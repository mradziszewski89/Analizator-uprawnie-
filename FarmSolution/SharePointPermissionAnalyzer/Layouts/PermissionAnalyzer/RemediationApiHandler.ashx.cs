using System;
using System.IO;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;

namespace SharePointPermissionAnalyzer.Layouts.PermissionAnalyzer
{
    /// <summary>
    /// HTTP Handler dla API remediacji uprawnień.
    /// Obsługuje żądania JSON z frontendu raportu hostowanego na SharePoint.
    ///
    /// Bezpieczeństwo:
    /// - Sprawdza autentykację Windows (Kerberos/NTLM)
    /// - Wymaga Farm Admin lub skonfigurowanej grupy
    /// - Waliduje wszystkie dane wejściowe
    /// - Zwraca tylko JSON (nie HTML)
    /// - Loguje wszystkie operacje
    ///
    /// Endpoint: /_layouts/15/PermissionAnalyzer/RemediationApiHandler.ashx
    /// Method: POST
    /// Content-Type: application/json
    /// Body: { "action": "...", "siteUrl": "...", "webUrl": "...", ... }
    /// </summary>
    public class RemediationApiHandler : IHttpHandler, System.Web.SessionState.IRequiresSessionState
    {
        private const string AdminGroupName = "SHAREPOINT\\Farm Administrators";

        public bool IsReusable { get { return false; } }

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "application/json";
            context.Response.Charset = "utf-8";

            // Ustaw nagłówki bezpieczeństwa
            context.Response.Headers["X-Content-Type-Options"] = "nosniff";
            context.Response.Headers["X-Frame-Options"] = "SAMEORIGIN";
            context.Response.Headers["Cache-Control"] = "no-store, no-cache, must-revalidate";
            context.Response.Headers["Pragma"] = "no-cache";

            // Tylko POST
            if (!context.Request.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase))
            {
                WriteError(context, 405, "Method Not Allowed. Use POST.");
                return;
            }

            // Sprawdź autentykację
            if (!context.Request.IsAuthenticated)
            {
                WriteError(context, 401, "Unauthorized. Windows authentication required.");
                return;
            }

            // Sprawdź uprawnienia
            if (!IsAuthorized(context))
            {
                WriteError(context, 403, "Forbidden. You do not have permission to perform remediation operations.");
                return;
            }

            // Parsuj żądanie
            ApiRequest request = null;
            try
            {
                var body = ReadRequestBody(context);
                if (string.IsNullOrEmpty(body))
                {
                    WriteError(context, 400, "Empty request body.");
                    return;
                }

                var serializer = new JavaScriptSerializer();
                request = serializer.Deserialize<ApiRequest>(body);

                if (request == null)
                {
                    WriteError(context, 400, "Invalid JSON request body.");
                    return;
                }
            }
            catch (Exception ex)
            {
                WriteError(context, 400, "JSON parse error: " + ex.Message);
                return;
            }

            // Walidacja żądania
            var validationError = ValidateRequest(request);
            if (!string.IsNullOrEmpty(validationError))
            {
                WriteError(context, 400, validationError);
                return;
            }

            // Wykonaj operację
            var result = new ApiResponse();

            try
            {
                var operatedBy = context.User.Identity.Name;

                if (request.DryRun)
                {
                    result.Success = true;
                    result.Message = "[DRY-RUN] Operacja " + request.Action + " na " + request.ObjectUrl + " przez " + request.PrincipalLoginName;
                    result.IsDryRun = true;
                }
                else
                {
                    ExecuteOperation(request, result);
                }

                // Zapisz do logu
                WriteAuditLog(context, request, result, operatedBy);
            }
            catch (UnauthorizedAccessException)
            {
                WriteError(context, 403, "Access denied during operation execution.");
                return;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = "Operation failed: " + ex.Message;
                WriteAuditLog(context, request, result, context.User.Identity.Name);
            }

            // Zwróć odpowiedź
            var serializer2 = new JavaScriptSerializer();
            context.Response.Write(serializer2.Serialize(result));
        }

        private bool IsAuthorized(HttpContext context)
        {
            try
            {
                // Sprawdź Farm Admin przez SPSecurity
                bool isFarmAdmin = false;
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    isFarmAdmin = SPFarm.Local.CurrentUserIsAdministrator();
                });
                if (isFarmAdmin) return true;

                // Sprawdź grupę Windows
                if (context.User.IsInRole(AdminGroupName)) return true;

                return false;
            }
            catch
            {
                return false;
            }
        }

        private string ValidateRequest(ApiRequest req)
        {
            if (string.IsNullOrEmpty(req.Action))
                return "Action is required.";

            var allowedActions = new[] {
                "RemoveDirectUserPermission",
                "RemoveSharePointGroupAssignment",
                "RemoveDomainGroupAssignment",
                "RestoreInheritance"
            };

            bool actionValid = false;
            foreach (var a in allowedActions)
                if (a == req.Action) { actionValid = true; break; }

            if (!actionValid)
                return "Unknown action: " + req.Action;

            if (string.IsNullOrEmpty(req.SiteCollectionUrl))
                return "SiteCollectionUrl is required.";

            if (!req.SiteCollectionUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                return "Invalid SiteCollectionUrl format.";

            if (string.IsNullOrEmpty(req.Reason))
                return "Reason is required for audit log.";

            if (req.Action != "RestoreInheritance" &&
                string.IsNullOrEmpty(req.PrincipalLoginName) &&
                string.IsNullOrEmpty(req.SharePointGroupName))
                return "PrincipalLoginName or SharePointGroupName is required for this action.";

            // Sprawdź ochronę konta
            if (!string.IsNullOrEmpty(req.PrincipalLoginName))
            {
                var protectedPatterns = new[] {
                    "SHAREPOINT\\system",
                    "NT AUTHORITY\\"
                };
                foreach (var p in protectedPatterns)
                {
                    if (req.PrincipalLoginName.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                        return "Protected account cannot be modified: " + req.PrincipalLoginName;
                }
            }

            return null;  // Brak błędów
        }

        private void ExecuteOperation(ApiRequest req, ApiResponse result)
        {
            SPSite site = null;
            SPWeb web = null;

            try
            {
                site = new SPSite(req.SiteCollectionUrl);
                var webRelUrl = string.IsNullOrEmpty(req.WebUrl)
                    ? ""
                    : req.WebUrl.Replace(req.SiteCollectionUrl, "").TrimStart('/');

                web = string.IsNullOrEmpty(webRelUrl) ? site.RootWeb : site.OpenWeb(webRelUrl);
                web.AllowUnsafeUpdates = true;

                switch (req.Action)
                {
                    case "RemoveDirectUserPermission":
                    {
                        var user = web.EnsureUser(req.PrincipalLoginName);
                        RemoveFromSecurable(web, req, user);
                        result.Success = true;
                        result.Message = "Removed user permission: " + req.PrincipalLoginName;
                        break;
                    }
                    case "RemoveSharePointGroupAssignment":
                    {
                        var group = web.SiteGroups[req.SharePointGroupName];
                        RemoveFromSecurable(web, req, group);
                        result.Success = true;
                        result.Message = "Removed SharePoint group: " + req.SharePointGroupName;
                        break;
                    }
                    case "RemoveDomainGroupAssignment":
                    {
                        var domainGroup = web.EnsureUser(req.PrincipalLoginName);
                        RemoveFromSecurable(web, req, domainGroup);
                        result.Success = true;
                        result.Message = "Removed domain group: " + req.PrincipalLoginName;
                        break;
                    }
                    case "RestoreInheritance":
                    {
                        RestoreInheritanceOnSecurable(web, req);
                        result.Success = true;
                        result.Message = "Inheritance restored on: " + (req.ObjectUrl ?? req.WebUrl);
                        break;
                    }
                }

                web.AllowUnsafeUpdates = false;
            }
            finally
            {
                if (web != null) { web.AllowUnsafeUpdates = false; web.Dispose(); }
                if (site != null) site.Dispose();
            }
        }

        private void RemoveFromSecurable(SPWeb web, ApiRequest req, SPPrincipal principal)
        {
            if (!string.IsNullOrEmpty(req.ListId) && req.ItemId > 0)
            {
                var list = web.Lists[new Guid(req.ListId)];
                var item = list.GetItemById(req.ItemId);
                item.RoleAssignments.Remove(principal);
                item.Update();
            }
            else if (!string.IsNullOrEmpty(req.ListId))
            {
                var list = web.Lists[new Guid(req.ListId)];
                list.RoleAssignments.Remove(principal);
                list.Update();
            }
            else
            {
                web.RoleAssignments.Remove(principal);
                web.Update();
            }
        }

        private void RestoreInheritanceOnSecurable(SPWeb web, ApiRequest req)
        {
            if (!string.IsNullOrEmpty(req.ListId) && req.ItemId > 0)
            {
                var list = web.Lists[new Guid(req.ListId)];
                var item = list.GetItemById(req.ItemId);
                item.ResetRoleInheritance();
                item.Update();
            }
            else if (!string.IsNullOrEmpty(req.ListId))
            {
                var list = web.Lists[new Guid(req.ListId)];
                list.ResetRoleInheritance();
                list.Update();
            }
            else
            {
                web.ResetRoleInheritance();
                web.Update();
            }
        }

        private string ReadRequestBody(HttpContext context)
        {
            using (var reader = new StreamReader(context.Request.InputStream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private void WriteError(HttpContext context, int statusCode, string message)
        {
            context.Response.StatusCode = statusCode;
            var serializer = new JavaScriptSerializer();
            context.Response.Write(serializer.Serialize(new { success = false, message = message, statusCode = statusCode }));
        }

        private void WriteAuditLog(HttpContext context, ApiRequest req, ApiResponse result, string operatedBy)
        {
            try
            {
                var logPath = context.Server.MapPath("~/_layouts/PermissionAnalyzer/ApiAuditLog.csv");
                var logDir = Path.GetDirectoryName(logPath);
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);

                bool exists = File.Exists(logPath);
                using (var writer = new StreamWriter(logPath, append: true, encoding: Encoding.UTF8))
                {
                    if (!exists)
                        writer.WriteLine("Timestamp;OperatedBy;Action;SiteUrl;WebUrl;ListId;ItemId;Principal;DryRun;Success;Message;Reason");

                    writer.WriteLine(string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}",
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        CsvE(operatedBy), CsvE(req.Action), CsvE(req.SiteCollectionUrl),
                        CsvE(req.WebUrl), CsvE(req.ListId), req.ItemId,
                        CsvE(req.PrincipalLoginName), req.DryRun,
                        result.Success, CsvE(result.Message), CsvE(req.Reason)
                    ));
                }
            }
            catch { /* Nie przerywaj głównej operacji */ }
        }

        private string CsvE(string s) { return "\"" + (s ?? "").Replace("\"", "\"\"") + "\""; }
    }

    public class ApiRequest
    {
        public string Action { get; set; }
        public string SiteCollectionUrl { get; set; }
        public string WebUrl { get; set; }
        public string ObjectUrl { get; set; }
        public string ListId { get; set; }
        public int ItemId { get; set; }
        public string PrincipalLoginName { get; set; }
        public string SharePointGroupName { get; set; }
        public string Reason { get; set; }
        public bool DryRun { get; set; }
    }

    public class ApiResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public bool IsDryRun { get; set; }
        public string Timestamp { get; set; } = DateTime.Now.ToString("o");
    }
}

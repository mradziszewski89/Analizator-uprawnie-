using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.ApplicationPages;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;

namespace SharePointPermissionAnalyzer.Layouts.PermissionAnalyzer
{
    /// <summary>
    /// Application Page dla bezpiecznej remediacji uprawnień SharePoint.
    /// Dostęp kontrolowany przez grupę AD/SharePoint zdefiniowaną w web.config lub stałej.
    /// Wszystkie operacje są audytowane do pliku logu.
    /// </summary>
    public partial class PermissionRemediationPage : LayoutsPageBase
    {
        // ============================================================
        // KONFIGURACJA - dostosuj do środowiska
        // ============================================================

        /// <summary>
        /// Nazwa grupy SharePoint lub grupy AD wymaganej do dostępu.
        /// Format: "DOMENA\Nazwa_Grupy" (AD) lub "Nazwa Grupy SharePoint" (SP).
        /// Można też skonfigurować przez web.config (klucz "SPAnalyzerAdminGroup").
        /// </summary>
        private const string AdminGroupName = "SHAREPOINT\\Farm Administrators";

        /// <summary>
        /// Ścieżka do pliku logu audytowego.
        /// Puste = automatycznie w _logs/ obok pliku.
        /// </summary>
        private string AuditLogPath
        {
            get
            {
                var webConfigValue = System.Web.Configuration.WebConfigurationManager.AppSettings["SPAnalyzerAuditLogPath"];
                if (!string.IsNullOrEmpty(webConfigValue)) return webConfigValue;
                return Path.Combine(Server.MapPath("~/_layouts/PermissionAnalyzer"), "AuditLog.csv");
            }
        }

        /// <summary>
        /// LoginName kont chronionych przed usunięciem uprawnień.
        /// </summary>
        private static readonly string[] ProtectedLoginPatterns = new[]
        {
            @"^SHAREPOINT\\system$",
            @"^NT AUTHORITY\\",
            @"^SHAREPOINT\\"
        };

        // ============================================================
        // ZDARZENIA STRONY
        // ============================================================

        protected void Page_Load(object sender, EventArgs e)
        {
            // Wymagaj HTTPS w produkcji
            if (!Request.IsSecureConnection && !Request.IsLocal)
            {
                Response.Redirect("https://" + Request.Url.Host + Request.Url.PathAndQuery, true);
                return;
            }

            // Sprawdź uprawnienia
            if (!IsCurrentUserAuthorized())
            {
                pnlAccessDenied.Visible = true;
                pnlMain.Visible = false;
                litRequiredGroup.Text = SPEncode.HtmlEncode(AdminGroupName);
                return;
            }

            pnlMain.Visible = true;
            litCurrentUser.Text = SPEncode.HtmlEncode(SPContext.Current.Web.CurrentUser.LoginName);

            if (!IsPostBack)
            {
                LoadAuditHistory();
            }
        }

        /// <summary>
        /// Podgląd operacji (preview bez wykonania).
        /// </summary>
        protected void btnPreview_Click(object sender, EventArgs e)
        {
            if (!IsCurrentUserAuthorized())
            {
                ShowMessage("Odmowa dostępu. Brak wymaganych uprawnień.", "danger");
                return;
            }

            var operation = BuildRemediationOperation();
            if (operation == null) return;

            operation.IsDryRun = true;  // Zawsze dry-run dla podglądu

            var previewHtml = BuildPreviewHtml(operation);
            litPreview.Text = previewHtml;
            pnlPreview.Visible = true;

            ShowMessage("Podgląd operacji wygenerowany (DRY-RUN - bez zmian).", "info");
        }

        /// <summary>
        /// Wykonanie operacji remediacji.
        /// </summary>
        protected void btnExecute_Click(object sender, EventArgs e)
        {
            if (!IsCurrentUserAuthorized())
            {
                ShowMessage("Odmowa dostępu. Brak wymaganych uprawnień.", "danger");
                return;
            }

            var operation = BuildRemediationOperation();
            if (operation == null) return;

            operation.IsDryRun = chkDryRun.Checked;

            // Zapisz stan przed operacją (do logu)
            string stateBefore = GetObjectPermissionState(operation);

            var log = new StringBuilder();
            bool success = false;
            string errorMessage = "";

            try
            {
                if (operation.IsDryRun)
                {
                    log.AppendLine("[DRY-RUN] Symulacja operacji: " + operation.Action);
                    log.AppendLine("[DRY-RUN] Obiekt: " + operation.ObjectUrl);
                    log.AppendLine("[DRY-RUN] Principal: " + operation.PrincipalLoginName);
                    log.AppendLine("[DRY-RUN] Powód: " + operation.Reason);
                    log.AppendLine("[DRY-RUN] Stan przed: " + stateBefore);
                    log.AppendLine("[DRY-RUN] Żadne zmiany nie zostały wprowadzone.");
                    success = true;
                }
                else
                {
                    ExecuteRemediationOperation(operation, log);
                    success = true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                log.AppendLine("[BŁĄD] " + ex.Message);
                if (ex.InnerException != null)
                    log.AppendLine("[BŁĄD Inner] " + ex.InnerException.Message);
            }

            // Stan po operacji
            string stateAfter = success && !operation.IsDryRun ? GetObjectPermissionState(operation) : "(nie dotyczy - dry-run lub błąd)";

            // Zapis do logu audytowego
            WriteAuditLog(new AuditLogEntry
            {
                Timestamp = DateTime.Now,
                OperatedBy = SPContext.Current.Web.CurrentUser.LoginName,
                Action = operation.Action,
                ObjectUrl = operation.ObjectUrl,
                SiteCollectionUrl = operation.SiteCollectionUrl,
                WebUrl = operation.WebUrl,
                ListId = operation.ListId,
                ItemId = operation.ItemId.ToString(),
                Principal = operation.PrincipalLoginName,
                IsDryRun = operation.IsDryRun,
                Result = success ? "Success" : "Error: " + errorMessage,
                Reason = operation.Reason,
                StateBefore = stateBefore,
                StateAfter = stateAfter
            });

            // Wyświetl log na stronie
            litLog.Text = SPEncode.HtmlEncode(log.ToString()).Replace("\n", "<br/>");
            pnlLog.Visible = true;

            if (success)
            {
                ShowMessage(operation.IsDryRun ? "DRY-RUN zakończony pomyślnie. Brak rzeczywistych zmian." : "Operacja zakończona pomyślnie.", "success");
            }
            else
            {
                ShowMessage("Błąd operacji: " + SPEncode.HtmlEncode(errorMessage), "danger");
            }

            // Odśwież historię
            LoadAuditHistory();
        }

        /// <summary>
        /// Eksport historii do CSV.
        /// </summary>
        protected void btnExportHistory_Click(object sender, EventArgs e)
        {
            if (!IsCurrentUserAuthorized()) return;

            Response.Clear();
            Response.ContentType = "text/csv";
            Response.Charset = "UTF-8";
            Response.AddHeader("Content-Disposition", "attachment; filename=AuditHistory_" + DateTime.Now.ToString("yyyy-MM-dd") + ".csv");
            Response.BinaryWrite(Encoding.UTF8.GetPreamble());

            if (File.Exists(AuditLogPath))
            {
                Response.Write(File.ReadAllText(AuditLogPath, Encoding.UTF8));
            }
            else
            {
                Response.Write("Brak danych historii.\r\n");
            }

            Response.End();
        }

        // ============================================================
        // LOGIKA BIZNESOWA
        // ============================================================

        /// <summary>
        /// Sprawdza czy bieżący użytkownik jest uprawniony do używania strony.
        /// Wymaga Farm Admin lub członkostwa w skonfigurowanej grupie.
        /// </summary>
        private bool IsCurrentUserAuthorized()
        {
            try
            {
                // Sprawdź Farm Admin
                if (SPContext.Current.Site.WebApplication.Farm.CurrentUserIsAdministrator())
                    return true;

                // Sprawdź SC Admin
                if (SPContext.Current.Site.RootWeb.CurrentUser.IsSiteAdmin)
                    return true;

                // Sprawdź grupę AD (przez Windows Identity)
                var windowsIdentity = HttpContext.Current.User as System.Security.Principal.WindowsPrincipal;
                if (windowsIdentity != null && windowsIdentity.IsInRole(AdminGroupName))
                    return true;

                // Sprawdź grupę SharePoint
                var web = SPContext.Current.Web;
                try
                {
                    var group = web.SiteGroups[AdminGroupName];
                    if (group != null)
                    {
                        var currentUser = web.CurrentUser;
                        foreach (SPUser member in group.Users)
                        {
                            if (member.LoginName.Equals(currentUser.LoginName, StringComparison.OrdinalIgnoreCase))
                                return true;
                        }
                    }
                }
                catch { /* Grupa SP nie istnieje - ignoruj */ }

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Sprawdza czy LoginName jest na liście chronionych kont.
        /// </summary>
        private bool IsProtectedAccount(string loginName)
        {
            if (string.IsNullOrEmpty(loginName)) return false;

            foreach (var pattern in ProtectedLoginPatterns)
            {
                if (Regex.IsMatch(loginName, pattern, RegexOptions.IgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Buduje obiekt operacji remediacji z danych formularza.
        /// Waliduje dane wejściowe.
        /// </summary>
        private RemediationOperation BuildRemediationOperation()
        {
            var action = ddlAction.SelectedValue;
            var siteUrl = txtSiteCollectionUrl.Text.Trim();
            var webUrl = txtWebUrl.Text.Trim();
            var listId = txtListId.Text.Trim();
            var itemIdStr = txtItemId.Text.Trim();
            var principalLogin = txtPrincipalLogin.Text.Trim();
            var spGroupName = txtSharePointGroupName.Text.Trim();
            var reason = txtReason.Text.Trim();

            // Walidacja
            if (string.IsNullOrEmpty(action))
            {
                ShowMessage("Wybierz typ operacji.", "danger");
                return null;
            }

            if (string.IsNullOrEmpty(siteUrl))
            {
                ShowMessage("Podaj URL Site Collection.", "danger");
                return null;
            }

            // Walidacja URL (zapobieganie SSRF)
            if (!IsValidSharePointUrl(siteUrl))
            {
                ShowMessage("Nieprawidłowy URL Site Collection.", "danger");
                return null;
            }

            if (string.IsNullOrEmpty(reason))
            {
                ShowMessage("Podaj powód operacji (wymagane do logu audytowego).", "danger");
                return null;
            }

            if (action != "RestoreInheritance" && string.IsNullOrEmpty(principalLogin) && string.IsNullOrEmpty(spGroupName))
            {
                ShowMessage("Podaj LoginName principal lub nazwę grupy SharePoint.", "danger");
                return null;
            }

            // Sprawdź ochronę konta
            if (!string.IsNullOrEmpty(principalLogin) && IsProtectedAccount(principalLogin))
            {
                ShowMessage("Odmowa: konto " + SPEncode.HtmlEncode(principalLogin) + " jest chronione i nie może być modyfikowane przez ten interfejs.", "danger");
                return null;
            }

            int itemId = 0;
            if (!string.IsNullOrEmpty(itemIdStr) && !int.TryParse(itemIdStr, out itemId))
            {
                ShowMessage("ID elementu musi być liczbą całkowitą.", "danger");
                return null;
            }

            Guid listGuid = Guid.Empty;
            if (!string.IsNullOrEmpty(listId) && !Guid.TryParse(listId.Trim('{', '}'), out listGuid))
            {
                ShowMessage("GUID listy ma nieprawidłowy format.", "danger");
                return null;
            }

            var effectiveWebUrl = string.IsNullOrEmpty(webUrl) ? siteUrl : webUrl;

            return new RemediationOperation
            {
                Action = action,
                SiteCollectionUrl = siteUrl,
                WebUrl = effectiveWebUrl,
                ObjectUrl = effectiveWebUrl,
                ListId = listGuid.ToString(),
                ItemId = itemId,
                PrincipalLoginName = principalLogin,
                SharePointGroupName = spGroupName,
                Reason = reason,
                IsDryRun = chkDryRun.Checked
            };
        }

        /// <summary>
        /// Walidacja czy URL należy do tej farmy SharePoint.
        /// Zapobiega SSRF przez podanie zewnętrznego URL.
        /// </summary>
        private bool IsValidSharePointUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return false;

            // Sprawdź czy URL zaczyna się od http:// lub https://
            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return false;

            // Sprawdź czy URL należy do tej farmy (jest zarejestrowany jako WebApplication)
            try
            {
                var farm = SPFarm.Local;
                var webService = SPWebService.ContentService;
                foreach (SPWebApplication webApp in webService.WebApplications)
                {
                    foreach (SPAlternateUrl altUrl in webApp.AlternateUrls)
                    {
                        if (url.StartsWith(altUrl.Uri.ToString(), StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }
            catch { }

            return false;
        }

        /// <summary>
        /// Pobiera aktualny stan uprawnień obiektu (do logu "przed").
        /// </summary>
        private string GetObjectPermissionState(RemediationOperation op)
        {
            var result = new StringBuilder();
            SPSite site = null;
            SPWeb web = null;

            try
            {
                site = new SPSite(op.SiteCollectionUrl);
                var webRelUrl = op.WebUrl.Replace(op.SiteCollectionUrl, "").TrimStart('/');
                web = string.IsNullOrEmpty(webRelUrl) ? site.RootWeb : site.OpenWeb(webRelUrl);

                object securable = null;

                if (!string.IsNullOrEmpty(op.ListId) && op.ItemId > 0)
                {
                    var list = web.Lists[new Guid(op.ListId)];
                    var item = list.GetItemById(op.ItemId);
                    securable = item;
                }
                else if (!string.IsNullOrEmpty(op.ListId))
                {
                    securable = web.Lists[new Guid(op.ListId)];
                }
                else
                {
                    securable = web;
                }

                var securableObj = securable as SPSecurableObject;
                if (securableObj != null)
                {
                    result.AppendLine("HasUniquePermissions: " + securableObj.HasUniqueRoleAssignments);
                    result.AppendLine("RoleAssignments:");
                    foreach (SPRoleAssignment ra in securableObj.RoleAssignments)
                    {
                        var levels = string.Join(", ", (from SPRoleDefinition rd in ra.RoleDefinitionBindings select rd.Name).ToArray());
                        result.AppendLine("  " + ra.Member.LoginName + " -> " + levels);
                    }
                }
            }
            catch (Exception ex)
            {
                result.AppendLine("Błąd pobierania stanu: " + ex.Message);
            }
            finally
            {
                if (web != null) web.Dispose();
                if (site != null) site.Dispose();
            }

            return result.ToString();
        }

        /// <summary>
        /// Wykonuje faktyczną operację remediacji.
        /// Wszystkie operacje opakowane w try/finally z Dispose.
        /// </summary>
        private void ExecuteRemediationOperation(RemediationOperation op, StringBuilder log)
        {
            SPSite site = null;
            SPWeb web = null;

            try
            {
                site = new SPSite(op.SiteCollectionUrl);
                var webRelUrl = op.WebUrl.Replace(op.SiteCollectionUrl, "").TrimStart('/');
                web = string.IsNullOrEmpty(webRelUrl) ? site.RootWeb : site.OpenWeb(webRelUrl);
                web.AllowUnsafeUpdates = true;

                log.AppendLine("[" + DateTime.Now.ToString("HH:mm:ss") + "] Otwarto Web: " + web.Url);

                switch (op.Action)
                {
                    case "RemoveDirectUserPermission":
                        ExecuteRemoveUserPermission(op, web, log);
                        break;

                    case "RemoveSharePointGroupAssignment":
                        ExecuteRemoveSharePointGroup(op, web, log);
                        break;

                    case "RemoveDomainGroupAssignment":
                        ExecuteRemoveDomainGroup(op, web, log);
                        break;

                    case "RestoreInheritance":
                        ExecuteRestoreInheritance(op, web, log);
                        break;

                    default:
                        throw new ArgumentException("Nieznana akcja: " + op.Action);
                }

                web.AllowUnsafeUpdates = false;
                log.AppendLine("[" + DateTime.Now.ToString("HH:mm:ss") + "] Operacja zakończona pomyślnie.");
            }
            finally
            {
                if (web != null) { web.AllowUnsafeUpdates = false; web.Dispose(); }
                if (site != null) site.Dispose();
            }
        }

        private void ExecuteRemoveUserPermission(RemediationOperation op, SPWeb web, StringBuilder log)
        {
            var user = web.EnsureUser(op.PrincipalLoginName);
            log.AppendLine("[INFO] Użytkownik: " + user.LoginName + " (" + user.Name + ")");

            if (!string.IsNullOrEmpty(op.ListId) && op.ItemId > 0)
            {
                var list = web.Lists[new Guid(op.ListId)];
                var item = list.GetItemById(op.ItemId);
                if (!item.HasUniqueRoleAssignments)
                    throw new InvalidOperationException("Element dziedziczy uprawnienia - nie można usunąć bezpośredniego przypisania.");
                item.RoleAssignments.Remove(user);
                item.Update();
                log.AppendLine("[OK] Usunięto uprawnienie użytkownika '" + user.LoginName + "' z elementu ID " + op.ItemId + " w liście '" + list.Title + "'");
            }
            else if (!string.IsNullOrEmpty(op.ListId))
            {
                var list = web.Lists[new Guid(op.ListId)];
                if (!list.HasUniqueRoleAssignments)
                    throw new InvalidOperationException("Lista dziedziczy uprawnienia - nie można usunąć bezpośredniego przypisania.");
                list.RoleAssignments.Remove(user);
                list.Update();
                log.AppendLine("[OK] Usunięto uprawnienie użytkownika '" + user.LoginName + "' z listy '" + list.Title + "'");
            }
            else
            {
                if (!web.HasUniqueRoleAssignments)
                    throw new InvalidOperationException("Witryna dziedziczy uprawnienia - nie można usunąć bezpośredniego przypisania.");

                // Ochrona: sprawdź czy nie usuwamy ostatniego Full Control
                ValidateNotRemovingLastFullControl(web, user, log);

                web.RoleAssignments.Remove(user);
                web.Update();
                log.AppendLine("[OK] Usunięto uprawnienie użytkownika '" + user.LoginName + "' z witryny '" + web.Url + "'");
            }
        }

        private void ExecuteRemoveSharePointGroup(RemediationOperation op, SPWeb web, StringBuilder log)
        {
            SPGroup group = null;
            try
            {
                group = web.SiteGroups[op.SharePointGroupName];
            }
            catch
            {
                // Próbuj przez EnsureUser jeśli podany jest LoginName
                if (!string.IsNullOrEmpty(op.PrincipalLoginName))
                {
                    var user = web.EnsureUser(op.PrincipalLoginName);
                    group = user as SPGroup;
                }
            }

            if (group == null)
                throw new ArgumentException("Nie znaleziono grupy SharePoint: " + op.SharePointGroupName);

            log.AppendLine("[INFO] Grupa SharePoint: " + group.Name);

            if (!string.IsNullOrEmpty(op.ListId) && op.ItemId > 0)
            {
                var list = web.Lists[new Guid(op.ListId)];
                var item = list.GetItemById(op.ItemId);
                item.RoleAssignments.Remove(group);
                item.Update();
                log.AppendLine("[OK] Usunięto grupę '" + group.Name + "' z elementu ID " + op.ItemId);
            }
            else if (!string.IsNullOrEmpty(op.ListId))
            {
                var list = web.Lists[new Guid(op.ListId)];
                list.RoleAssignments.Remove(group);
                list.Update();
                log.AppendLine("[OK] Usunięto grupę '" + group.Name + "' z listy '" + list.Title + "'");
            }
            else
            {
                ValidateNotRemovingLastFullControl(web, group, log);
                web.RoleAssignments.Remove(group);
                web.Update();
                log.AppendLine("[OK] Usunięto grupę '" + group.Name + "' z witryny '" + web.Url + "'");
            }
        }

        private void ExecuteRemoveDomainGroup(RemediationOperation op, SPWeb web, StringBuilder log)
        {
            var domainGroup = web.EnsureUser(op.PrincipalLoginName);
            log.AppendLine("[INFO] Grupa domenowa: " + domainGroup.LoginName);

            if (!string.IsNullOrEmpty(op.ListId) && op.ItemId > 0)
            {
                var list = web.Lists[new Guid(op.ListId)];
                var item = list.GetItemById(op.ItemId);
                item.RoleAssignments.Remove(domainGroup);
                item.Update();
                log.AppendLine("[OK] Usunięto grupę domenową '" + domainGroup.LoginName + "' z elementu ID " + op.ItemId);
            }
            else if (!string.IsNullOrEmpty(op.ListId))
            {
                var list = web.Lists[new Guid(op.ListId)];
                list.RoleAssignments.Remove(domainGroup);
                list.Update();
                log.AppendLine("[OK] Usunięto grupę domenową '" + domainGroup.LoginName + "' z listy '" + list.Title + "'");
            }
            else
            {
                ValidateNotRemovingLastFullControl(web, domainGroup, log);
                web.RoleAssignments.Remove(domainGroup);
                web.Update();
                log.AppendLine("[OK] Usunięto grupę domenową '" + domainGroup.LoginName + "' z witryny '" + web.Url + "'");
            }
        }

        private void ExecuteRestoreInheritance(RemediationOperation op, SPWeb web, StringBuilder log)
        {
            if (!string.IsNullOrEmpty(op.ListId) && op.ItemId > 0)
            {
                var list = web.Lists[new Guid(op.ListId)];
                var item = list.GetItemById(op.ItemId);
                if (!item.HasUniqueRoleAssignments)
                    throw new InvalidOperationException("Element już dziedziczy uprawnienia.");
                item.ResetRoleInheritance();
                item.Update();
                log.AppendLine("[OK] Przywrócono dziedziczenie dla elementu ID " + op.ItemId + " w liście '" + list.Title + "'");
            }
            else if (!string.IsNullOrEmpty(op.ListId))
            {
                var list = web.Lists[new Guid(op.ListId)];
                if (!list.HasUniqueRoleAssignments)
                    throw new InvalidOperationException("Lista już dziedziczy uprawnienia.");
                list.ResetRoleInheritance();
                list.Update();
                log.AppendLine("[OK] Przywrócono dziedziczenie dla listy '" + list.Title + "'");
            }
            else
            {
                if (!web.HasUniqueRoleAssignments)
                    throw new InvalidOperationException("Witryna już dziedziczy uprawnienia.");
                web.ResetRoleInheritance();
                web.Update();
                log.AppendLine("[OK] Przywrócono dziedziczenie dla witryny '" + web.Url + "'");
            }
        }

        /// <summary>
        /// Walidacja: czy po usunięciu nie zostaniemy bez żadnego Full Control na obiekcie.
        /// </summary>
        private void ValidateNotRemovingLastFullControl(SPWeb web, SPPrincipal principalToRemove, StringBuilder log)
        {
            int fullControlCount = 0;
            bool principalHasFullControl = false;

            foreach (SPRoleAssignment ra in web.RoleAssignments)
            {
                bool hasFC = false;
                foreach (SPRoleDefinition rd in ra.RoleDefinitionBindings)
                {
                    if (rd.Name == "Full Control") { hasFC = true; break; }
                }
                if (hasFC)
                {
                    fullControlCount++;
                    if (ra.Member.LoginName.Equals(principalToRemove.LoginName, StringComparison.OrdinalIgnoreCase))
                        principalHasFullControl = true;
                }
            }

            if (principalHasFullControl && fullControlCount <= 1)
            {
                throw new InvalidOperationException(
                    "Odmowa: usunięcie tego przypisania spowoduje brak kont z Full Control na witrynie. " +
                    "Najpierw nadaj Full Control innemu użytkownikowi lub grupie.");
            }

            log.AppendLine("[INFO] Walidacja Full Control: OK (" + fullControlCount + " kont z FC)");
        }

        // ============================================================
        // AUDIT LOG
        // ============================================================

        private void WriteAuditLog(AuditLogEntry entry)
        {
            try
            {
                var logDir = Path.GetDirectoryName(AuditLogPath);
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                bool fileExists = File.Exists(AuditLogPath);

                using (var writer = new StreamWriter(AuditLogPath, append: true, encoding: Encoding.UTF8))
                {
                    if (!fileExists)
                    {
                        // Nagłówek CSV
                        writer.WriteLine("Timestamp;OperatedBy;Action;ObjectUrl;SiteCollectionUrl;WebUrl;ListId;ItemId;Principal;IsDryRun;Result;Reason;StateBefore;StateAfter");
                    }

                    writer.WriteLine(string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13}",
                        entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"),
                        CsvEscape(entry.OperatedBy),
                        CsvEscape(entry.Action),
                        CsvEscape(entry.ObjectUrl),
                        CsvEscape(entry.SiteCollectionUrl),
                        CsvEscape(entry.WebUrl),
                        CsvEscape(entry.ListId),
                        entry.ItemId,
                        CsvEscape(entry.Principal),
                        entry.IsDryRun,
                        CsvEscape(entry.Result),
                        CsvEscape(entry.Reason),
                        CsvEscape(entry.StateBefore.Replace("\r\n", " | ")),
                        CsvEscape(entry.StateAfter.Replace("\r\n", " | "))
                    ));
                }
            }
            catch (Exception ex)
            {
                // Nie przerywaj operacji tylko dlatego że log się nie zapisał
                System.Diagnostics.EventLog.WriteEntry("SharePoint Permission Analyzer",
                    "Błąd zapisu logu audytowego: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
            }
        }

        private string CsvEscape(string s)
        {
            if (string.IsNullOrEmpty(s)) return "\"\"";
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        private void LoadAuditHistory()
        {
            var history = new List<AuditLogEntry>();

            if (!File.Exists(AuditLogPath))
            {
                gvHistory.DataSource = history;
                gvHistory.DataBind();
                return;
            }

            try
            {
                var lines = File.ReadAllLines(AuditLogPath, Encoding.UTF8)
                    .Skip(1) // Pomiń nagłówek
                    .Reverse()
                    .Take(50)
                    .ToList();

                foreach (var line in lines)
                {
                    var parts = ParseCsvLine(line);
                    if (parts.Length < 11) continue;

                    history.Add(new AuditLogEntry
                    {
                        Timestamp = DateTime.TryParse(parts[0], out DateTime dt) ? dt : DateTime.MinValue,
                        OperatedBy = parts[1],
                        Action = parts[2],
                        ObjectUrl = parts[3],
                        Principal = parts[8],
                        IsDryRun = parts[9] == "True",
                        Result = parts[10],
                        Reason = parts.Length > 11 ? parts[11] : ""
                    });
                }
            }
            catch { /* Ignoruj błędy parsowania logu */ }

            gvHistory.DataSource = history;
            gvHistory.DataBind();
        }

        private string[] ParseCsvLine(string line)
        {
            // Prosty parser CSV z obsługą cudzysłowów
            var result = new List<string>();
            bool inQuotes = false;
            var current = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ';' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
            result.Add(current.ToString());
            return result.ToArray();
        }

        // ============================================================
        // UI HELPERS
        // ============================================================

        private void ShowMessage(string message, string type)
        {
            litMessage.Text = SPEncode.HtmlEncode(message);
            divMessage.Attributes["class"] = "spa-alert spa-alert-" + type;
            pnlMessage.Visible = true;
        }

        private string BuildPreviewHtml(RemediationOperation op)
        {
            return "<table class='detail-table' style='width:100%;border-collapse:collapse'>"
                + "<tr><th style='padding:6px;background:#f3f2f1;border-bottom:1px solid #edebe9;width:200px'>Akcja</th><td style='padding:6px;border-bottom:1px solid #edebe9'>" + SPEncode.HtmlEncode(op.Action) + "</td></tr>"
                + "<tr><th style='padding:6px;background:#f3f2f1;border-bottom:1px solid #edebe9'>Site Collection</th><td style='padding:6px;border-bottom:1px solid #edebe9'>" + SPEncode.HtmlEncode(op.SiteCollectionUrl) + "</td></tr>"
                + "<tr><th style='padding:6px;background:#f3f2f1;border-bottom:1px solid #edebe9'>Web URL</th><td style='padding:6px;border-bottom:1px solid #edebe9'>" + SPEncode.HtmlEncode(op.WebUrl) + "</td></tr>"
                + "<tr><th style='padding:6px;background:#f3f2f1;border-bottom:1px solid #edebe9'>Lista GUID</th><td style='padding:6px;border-bottom:1px solid #edebe9'>" + SPEncode.HtmlEncode(op.ListId) + "</td></tr>"
                + "<tr><th style='padding:6px;background:#f3f2f1;border-bottom:1px solid #edebe9'>Element ID</th><td style='padding:6px;border-bottom:1px solid #edebe9'>" + op.ItemId + "</td></tr>"
                + "<tr><th style='padding:6px;background:#f3f2f1;border-bottom:1px solid #edebe9'>Principal</th><td style='padding:6px;border-bottom:1px solid #edebe9'>" + SPEncode.HtmlEncode(op.PrincipalLoginName) + "</td></tr>"
                + "<tr><th style='padding:6px;background:#f3f2f1;border-bottom:1px solid #edebe9'>Powód</th><td style='padding:6px;border-bottom:1px solid #edebe9'>" + SPEncode.HtmlEncode(op.Reason) + "</td></tr>"
                + "<tr><th style='padding:6px;background:#f3f2f1'>Tryb</th><td style='padding:6px'>" + (op.IsDryRun ? "<strong style='color:#c7a400'>DRY-RUN</strong>" : "<strong style='color:#d13438'>LIVE</strong>") + "</td></tr>"
                + "</table>";
        }
    }

    // ============================================================
    // MODELE DANYCH
    // ============================================================

    public class RemediationOperation
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
        public bool IsDryRun { get; set; }
    }

    public class AuditLogEntry
    {
        public DateTime Timestamp { get; set; }
        public string OperatedBy { get; set; }
        public string Action { get; set; }
        public string ObjectUrl { get; set; }
        public string SiteCollectionUrl { get; set; }
        public string WebUrl { get; set; }
        public string ListId { get; set; }
        public string ItemId { get; set; }
        public string Principal { get; set; }
        public bool IsDryRun { get; set; }
        public string Result { get; set; }
        public string Reason { get; set; }
        public string StateBefore { get; set; }
        public string StateAfter { get; set; }
    }
}

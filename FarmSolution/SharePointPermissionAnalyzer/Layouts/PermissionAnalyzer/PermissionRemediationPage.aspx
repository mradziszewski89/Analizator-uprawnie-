<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PermissionRemediationPage.aspx.cs" Inherits="SharePointPermissionAnalyzer.Layouts.PermissionAnalyzer.PermissionRemediationPage" DynamicMasterPageFile="~masterurl/default.master" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <title>SharePoint Permission Analyzer - Remediacja</title>
    <style>
        .spa-header { background: #0078d4; color: white; padding: 16px 24px; margin-bottom: 24px; }
        .spa-header h1 { font-size: 20px; margin: 0; }
        .spa-container { max-width: 1200px; margin: 0 auto; padding: 0 24px; }
        .spa-alert { padding: 12px 16px; border-radius: 4px; margin-bottom: 16px; }
        .spa-alert-danger { background: #fde7e9; border-left: 4px solid #d13438; color: #741e21; }
        .spa-alert-success { background: #dff6dd; border-left: 4px solid #107c10; color: #0a4a07; }
        .spa-alert-info { background: #deecf9; border-left: 4px solid #0078d4; color: #003e82; }
        .spa-form-group { margin-bottom: 16px; }
        .spa-form-group label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 13px; }
        .spa-form-group select, .spa-form-group input, .spa-form-group textarea { width: 100%; padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        .spa-button { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; }
        .spa-button-primary { background: #0078d4; color: white; }
        .spa-button-danger { background: #d13438; color: white; }
        .spa-button-secondary { background: #f3f2f1; color: #323130; border: 1px solid #ccc; }
        .spa-log { background: #1e1e1e; color: #d4d4d4; font-family: Consolas, monospace; font-size: 12px; padding: 16px; border-radius: 4px; max-height: 400px; overflow-y: auto; white-space: pre-wrap; }
        .spa-section { background: white; border: 1px solid #edebe9; border-radius: 8px; padding: 16px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
        .spa-section h2 { font-size: 16px; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 1px solid #edebe9; }
        .spa-access-denied { text-align: center; padding: 60px; color: #d13438; }
        .spa-access-denied h2 { font-size: 24px; }
    </style>
</asp:Content>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <div class="spa-container">

        <!-- Komunikat braku dostępu -->
        <asp:Panel ID="pnlAccessDenied" runat="server" Visible="false" CssClass="spa-access-denied">
            <h2>⛔ Odmowa dostępu</h2>
            <p>Nie masz uprawnień do wykonywania operacji remediacji uprawnień SharePoint.</p>
            <p>Wymagana rola: <strong><asp:Literal ID="litRequiredGroup" runat="server" /></strong></p>
        </asp:Panel>

        <!-- Główna treść strony (tylko dla uprawnionych) -->
        <asp:Panel ID="pnlMain" runat="server" Visible="false">

            <div class="spa-header">
                <h1>SharePoint Permission Analyzer - Remediacja uprawnień</h1>
                <small>Zalogowany jako: <asp:Literal ID="litCurrentUser" runat="server" /></small>
            </div>

            <!-- Komunikaty -->
            <asp:Panel ID="pnlMessage" runat="server" Visible="false">
                <div class="spa-alert" id="divMessage" runat="server">
                    <asp:Literal ID="litMessage" runat="server" />
                </div>
            </asp:Panel>

            <!-- Formularz remediacji -->
            <div class="spa-section">
                <h2>Nowa operacja remediacji</h2>

                <div class="spa-form-group">
                    <label>Typ operacji *</label>
                    <asp:DropDownList ID="ddlAction" runat="server" CssClass="">
                        <asp:ListItem Value="">-- wybierz --</asp:ListItem>
                        <asp:ListItem Value="RemoveDirectUserPermission">Usuń bezpośrednie uprawnienie użytkownika</asp:ListItem>
                        <asp:ListItem Value="RemoveSharePointGroupAssignment">Usuń przypisanie grupy SharePoint z obiektu</asp:ListItem>
                        <asp:ListItem Value="RemoveDomainGroupAssignment">Usuń przypisanie grupy domenowej z obiektu</asp:ListItem>
                        <asp:ListItem Value="RestoreInheritance">Przywróć dziedziczenie uprawnień</asp:ListItem>
                    </asp:DropDownList>
                </div>

                <div class="spa-form-group">
                    <label>URL Site Collection *</label>
                    <asp:TextBox ID="txtSiteCollectionUrl" runat="server" placeholder="np. http://portal.contoso.com/sites/moja-witryna" />
                </div>

                <div class="spa-form-group">
                    <label>URL Witryny (Web) - puste = SC root</label>
                    <asp:TextBox ID="txtWebUrl" runat="server" placeholder="np. http://portal.contoso.com/sites/moja-witryna/subsite" />
                </div>

                <div class="spa-form-group">
                    <label>GUID Listy (opcjonalnie - dla operacji na liście lub elemencie)</label>
                    <asp:TextBox ID="txtListId" runat="server" placeholder="np. {a1b2c3d4-e5f6-7890-abcd-ef1234567890}" />
                </div>

                <div class="spa-form-group">
                    <label>ID Elementu (opcjonalnie - dla operacji na elemencie)</label>
                    <asp:TextBox ID="txtItemId" runat="server" placeholder="np. 42" />
                </div>

                <div class="spa-form-group">
                    <label>LoginName Principal (użytkownik/grupa domenowa) - wymagane dla Remove*</label>
                    <asp:TextBox ID="txtPrincipalLogin" runat="server" placeholder="np. CONTOSO\jan.kowalski lub i:0#.w|contoso|jan.kowalski" />
                </div>

                <div class="spa-form-group">
                    <label>Nazwa grupy SharePoint (wymagane dla RemoveSharePointGroupAssignment)</label>
                    <asp:TextBox ID="txtSharePointGroupName" runat="server" placeholder="np. Odwiedzający witryny" />
                </div>

                <div class="spa-form-group">
                    <label>Powód operacji (do logu audytowego) *</label>
                    <asp:TextBox ID="txtReason" runat="server" TextMode="MultiLine" Rows="3" placeholder="Opisz powód odebrania uprawnień..." />
                </div>

                <div class="spa-form-group">
                    <label>
                        <asp:CheckBox ID="chkDryRun" runat="server" Checked="true" />
                        Tryb DRY-RUN (symulacja - bez rzeczywistych zmian)
                    </label>
                </div>

                <div>
                    <asp:Button ID="btnPreview" runat="server" Text="Podgląd (Preview)" CssClass="spa-button spa-button-secondary" OnClick="btnPreview_Click" />
                    <asp:Button ID="btnExecute" runat="server" Text="Wykonaj operację" CssClass="spa-button spa-button-danger" OnClick="btnExecute_Click" OnClientClick="return confirm('Czy na pewno chcesz wykonać tę operację? Ta akcja może być nieodwracalna.');" />
                </div>
            </div>

            <!-- Podgląd operacji -->
            <asp:Panel ID="pnlPreview" runat="server" Visible="false">
                <div class="spa-section">
                    <h2>Podgląd operacji</h2>
                    <asp:Literal ID="litPreview" runat="server" />
                </div>
            </asp:Panel>

            <!-- Log operacji -->
            <asp:Panel ID="pnlLog" runat="server" Visible="false">
                <div class="spa-section">
                    <h2>Log wykonania</h2>
                    <div class="spa-log">
                        <asp:Literal ID="litLog" runat="server" />
                    </div>
                </div>
            </asp:Panel>

            <!-- Historia operacji -->
            <div class="spa-section">
                <h2>Historia operacji remediacji</h2>
                <p style="color:#605e5c;font-size:13px">Ostatnie 50 operacji wykonanych przez tę stronę.</p>
                <asp:GridView ID="gvHistory" runat="server"
                    AutoGenerateColumns="false"
                    EmptyDataText="Brak historii operacji."
                    CssClass="spa-grid"
                    Width="100%">
                    <Columns>
                        <asp:BoundField DataField="Timestamp" HeaderText="Data/Czas" DataFormatString="{0:yyyy-MM-dd HH:mm:ss}" />
                        <asp:BoundField DataField="OperatedBy" HeaderText="Wykonał" />
                        <asp:BoundField DataField="Action" HeaderText="Akcja" />
                        <asp:BoundField DataField="ObjectUrl" HeaderText="Obiekt" />
                        <asp:BoundField DataField="Principal" HeaderText="Principal" />
                        <asp:BoundField DataField="IsDryRun" HeaderText="DryRun" />
                        <asp:BoundField DataField="Result" HeaderText="Wynik" />
                        <asp:BoundField DataField="Reason" HeaderText="Powód" />
                    </Columns>
                </asp:GridView>
                <br />
                <asp:Button ID="btnExportHistory" runat="server" Text="Eksportuj historię (CSV)" CssClass="spa-button spa-button-secondary" OnClick="btnExportHistory_Click" />
            </div>

        </asp:Panel>
    </div>
</asp:Content>

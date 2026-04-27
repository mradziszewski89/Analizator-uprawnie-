#Requires -Version 5.1
<#
.SYNOPSIS
    Modul generowania skryptow PowerShell do remediacji uprawnien SharePoint.

.DESCRIPTION
    Generuje gotowe, bezpieczne skrypty PS1 do wykonania przez administratora.
    Skrypt zawiera tryb dry-run, transcript, rollback i komentarze.
    Moze byc wywolany przez PowerShell zarowno z frontendu (generowanie) jak i backend SharePoint.
#>

Set-StrictMode -Version Latest

function New-RemediationScript {
    <#
    .SYNOPSIS
        Generuje skrypt PowerShell PS1 do wykonania remediacji uprawnien.

    .PARAMETER RemediationPlan
        Lista elementow do remediacji (z frontendu HTML lub backendu).
        Kazdy element: ObjectId, ObjectType, WebUrl, SiteCollectionUrl,
                       FullUrl, ServerRelativeUrl, ListId, ItemId,
                       Action, PrincipalLoginName, SharePointGroupName, Reason

    .PARAMETER DryRunDefault
        Czy skrypt domyslnie uruchamia sie w trybie dry-run.

    .PARAMETER ScriptTitle
        Tytul skryptu (w komentarzu naglowkowym).

    .PARAMETER GeneratedBy
        Informacja o generatorze (uzytkownik/system).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$RemediationPlan,

        [Parameter(Mandatory = $false)]
        [bool]$DryRunDefault = $true,

        [Parameter(Mandatory = $false)]
        [string]$ScriptTitle = "SharePoint Permission Remediation",

        [Parameter(Mandatory = $false)]
        [string]$GeneratedBy = "SharePoint Permission Analyzer",

        [Parameter(Mandatory = $false)]
        [string]$WhitelistPath = ""
    )

    $GenerationTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $DryRunStr = if ($DryRunDefault) { '$true' } else { '$false' }

    # Serialize remediation plan as JSON
    $PlanJson = $RemediationPlan | ConvertTo-Json -Depth 8 -Compress

    $ScriptContent = @"
#Requires -Version 5.1
<#
.SYNOPSIS
    $ScriptTitle

.DESCRIPTION
    Skrypt wygenerowany automatycznie przez SharePoint Permission Analyzer.
    PRZED URUCHOMIENIEM: Przeczytaj plan remediacji i zweryfikuj zmiany.
    ZAWSZE uruchom najpierw w trybie DryRun = `$true.

.NOTES
    Wygenerowany przez : $GeneratedBy
    Data wygenerowania : $GenerationTime
    Liczba operacji    : $($RemediationPlan.Count)
    Domyslny tryb      : $(if ($DryRunDefault) { 'DRY-RUN (symulacja)' } else { 'LIVE (rzeczywiste zmiany)' })

    INSTRUKCJA:
    1. Uruchom skrypt jako Farm Administrator na serwerze SharePoint.
    2. Sprawdz transcript log po zakonczeniu.
    3. W razie potrzeby skorzystaj z sekcji ROLLBACK.

    UWAGA: Skrypt NIE usuwa kont z whitelist (Farm Admin, System Account itp.)
#>

[CmdletBinding(SupportsShouldProcess = `$true)]
param(
    [Parameter(Mandatory = `$false)]
    [bool]`$DryRun = $DryRunStr,

    [Parameter(Mandatory = `$false)]
    [string]`$TranscriptPath = "",

    [Parameter(Mandatory = `$false)]
    [switch]`$SkipConfirmation
)

Set-StrictMode -Version Latest
`$ErrorActionPreference = "Stop"

# ============================================================
# NAGLOWEK
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SharePoint Permission Remediation Script" -ForegroundColor Cyan
Write-Host "  Wygenerowany: $GenerationTime" -ForegroundColor Cyan
Write-Host "  Operacji: $($RemediationPlan.Count)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

if (`$DryRun) {
    Write-Host ""
    Write-Host "  TRYB: DRY-RUN - zadne zmiany nie beda wprowadzone" -ForegroundColor Yellow
    Write-Host "  Aby wykonac rzeczywiste zmiany: -DryRun `$false" -ForegroundColor Yellow
}
else {
    Write-Host ""
    Write-Host "  TRYB: LIVE - zmiany beda wprowadzone w SharePoint!" -ForegroundColor Red
}
Write-Host ""

# Transcript log
`$ScriptDir = Split-Path -Parent `$MyInvocation.MyCommand.Definition
if (-not `$TranscriptPath) {
    `$TranscriptPath = Join-Path `$ScriptDir "Remediation_`$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
}
Start-Transcript -Path `$TranscriptPath -Force | Out-Null
Write-Host "Transcript: `$TranscriptPath"

# ============================================================
# ZALADUJ SHAREPOINT
# ============================================================

if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    try {
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
        Write-Host "[OK] Zaladowano SharePoint PowerShell" -ForegroundColor Green
    }
    catch {
        Write-Error "BLAD: Nie mozna zaladowac Microsoft.SharePoint.PowerShell. Uruchom na serwerze SharePoint."
        Stop-Transcript
        exit 1
    }
}

# ============================================================
# WHITELIST CHRONIONYCH KONT
# ============================================================

# Konta i wzorce LoginName, ktorych NIE WOLNO usunac
`$ProtectedLoginPatterns = @(
    "^SHAREPOINT\\\\system`$",
    "^NT AUTHORITY\\\\",
    "^SHAREPOINT\\\\",
    "i:0#\\.s\\|membership\\|",   # App-only tokens
    "c:0\\!\\."                    # Special built-in claims
)

function Test-IsProtectedAccount {
    param([string]`$LoginName)
    foreach (`$Pattern in `$ProtectedLoginPatterns) {
        if (`$LoginName -match `$Pattern) { return `$true }
    }
    return `$false
}

# ============================================================
# PLAN REMEDIACJI
# ============================================================

`$RemediationPlanJson = @'
$PlanJson
'@

try {
    `$Plan = `$RemediationPlanJson | ConvertFrom-Json
}
catch {
    Write-Error "BLAD: Nie mozna zaladowac planu remediacji z JSON: `$_"
    Stop-Transcript
    exit 1
}

Write-Host "Plan remediacji: `$(`$Plan.Count) operacji" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# PODSUMOWANIE PLANU - WYSWIETL PRZED POTWIERDZENIEM
# ============================================================

Write-Host "PLAN OPERACJI:" -ForegroundColor White
Write-Host "-" * 60 -ForegroundColor Gray
foreach (`$Op in `$Plan) {
    `$ActionDesc = switch (`$Op.Action) {
        "RemoveDirectUserPermission"       { "Usun uprawnienie uzytkownika" }
        "RemoveSharePointGroupAssignment"  { "Usun przypisanie grupy SharePoint z obiektu" }
        "RemoveDomainGroupAssignment"      { "Usun przypisanie grupy domenowej z obiektu" }
        "RestoreInheritance"               { "Przywroc dziedziczenie uprawnien" }
        default                            { `$Op.Action }
    }
    Write-Host "  [`$(`$Op.Action)]" -ForegroundColor Yellow -NoNewline
    Write-Host " `$(`$Op.FullUrl)" -ForegroundColor White
    Write-Host "    Principal: `$(`$Op.PrincipalLoginName)" -ForegroundColor Gray
    Write-Host "    Powod    : `$(`$Op.Reason)" -ForegroundColor Gray
    Write-Host ""
}
Write-Host "-" * 60 -ForegroundColor Gray

# ============================================================
# POTWIERDZENIE
# ============================================================

if (-not `$DryRun -and -not `$SkipConfirmation) {
    Write-Host ""
    Write-Host "UWAGA: Te operacje sa nieodwracalne (o ile nie masz backupu)!" -ForegroundColor Red
    `$Confirm = Read-Host "Wpisz 'POTWIERDZAM' aby kontynuowac (lub cokolwiek innego aby anulowac)"
    if (`$Confirm -ne "POTWIERDZAM") {
        Write-Host "Anulowano przez uzytkownika." -ForegroundColor Yellow
        Stop-Transcript
        exit 0
    }
}

# ============================================================
# WYKONANIE OPERACJI
# ============================================================

`$Results = @()
`$SuccessCount = 0
`$ErrorCount = 0
`$SkippedCount = 0

foreach (`$Op in `$Plan) {
    Write-Host ""
    Write-Host "Przetwarzam: [`$(`$Op.Action)] `$(`$Op.FullUrl)" -ForegroundColor Cyan

    # Sprawdz czy principal jest chroniony
    if (Test-IsProtectedAccount -LoginName `$Op.PrincipalLoginName) {
        Write-Host "  [POMINIETO] Chronione konto: `$(`$Op.PrincipalLoginName)" -ForegroundColor Yellow
        `$Results += [PSCustomObject]@{
            Status    = "Skipped_Protected"
            Operation = `$Op.Action
            Url       = `$Op.FullUrl
            Principal = `$Op.PrincipalLoginName
            Message   = "Konto chronione - nie usuwane"
            Timestamp = (Get-Date -Format "o")
        }
        `$SkippedCount++
        continue
    }

    try {
        `$Result = Invoke-RemediationOperation -Operation `$Op -DryRun `$DryRun
        `$Results += `$Result

        if (`$Result.Status -eq "Success" -or `$Result.Status -eq "DryRun") {
            Write-Host "  [`$(`$Result.Status)] `$(`$Result.Message)" -ForegroundColor $(if (`$DryRun) { 'Yellow' } else { 'Green' })
            `$SuccessCount++
        }
        else {
            Write-Host "  [BLAD] `$(`$Result.Message)" -ForegroundColor Red
            `$ErrorCount++
        }
    }
    catch {
        Write-Host "  [WYJĄTEK] `$_" -ForegroundColor Red
        `$Results += [PSCustomObject]@{
            Status    = "Exception"
            Operation = `$Op.Action
            Url       = `$Op.FullUrl
            Principal = `$Op.PrincipalLoginName
            Message   = `$_.ToString()
            Timestamp = (Get-Date -Format "o")
        }
        `$ErrorCount++
    }
}

# ============================================================
# FUNKCJA GLOWNA OPERACJI
# ============================================================

function Invoke-RemediationOperation {
    param(
        [object]`$Operation,
        [bool]`$DryRun
    )

    `$Status = if (`$DryRun) { "DryRun" } else { "Success" }
    `$Message = ""
    `$SiteUrl = `$Operation.SiteCollectionUrl
    `$WebUrl = `$Operation.WebUrl
    `$ListId = `$Operation.ListId
    `$ItemId = `$Operation.ItemId
    `$LoginName = `$Operation.PrincipalLoginName

    try {
        switch (`$Operation.Action) {

            "RemoveDirectUserPermission" {
                if (`$DryRun) {
                    `$Message = "[DRY-RUN] Usunelby uprawnienie uzytkownika '`$LoginName' z '`$(`$Operation.FullUrl)'"
                }
                else {
                    `$Disposed = `$false
                    `$SPSite = `$null
                    `$SPWeb = `$null
                    try {
                        `$SPSite = New-Object Microsoft.SharePoint.SPSite(`$SiteUrl)
                        `$SPWeb = `$SPSite.OpenWeb(`$WebUrl.Replace(`$SiteUrl, ""))
                        `$SPWeb.AllowUnsafeUpdates = `$true

                        if (`$ListId -and `$ItemId) {
                            # Poziom elementu
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$Item = `$List.GetItemById([int]`$ItemId)
                            `$User = `$SPWeb.EnsureUser(`$LoginName)
                            `$Item.RoleAssignments.Remove(`$User)
                            `$Item.Update()
                            `$Message = "Usunieto uprawnienie uzytkownika '`$LoginName' z elementu ID `$ItemId w liscie '`$(`$List.Title)'"
                        }
                        elseif (`$ListId) {
                            # Poziom listy
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$User = `$SPWeb.EnsureUser(`$LoginName)
                            `$List.RoleAssignments.Remove(`$User)
                            `$List.Update()
                            `$Message = "Usunieto uprawnienie uzytkownika '`$LoginName' z listy '`$(`$List.Title)'"
                        }
                        else {
                            # Poziom witryny
                            `$User = `$SPWeb.EnsureUser(`$LoginName)
                            `$SPWeb.RoleAssignments.Remove(`$User)
                            `$SPWeb.Update()
                            `$Message = "Usunieto uprawnienie uzytkownika '`$LoginName' z witryny '`$(`$SPWeb.Url)'"
                        }

                        `$SPWeb.AllowUnsafeUpdates = `$false
                    }
                    finally {
                        if (`$SPWeb) { `$SPWeb.Dispose() }
                        if (`$SPSite) { `$SPSite.Dispose() }
                    }
                }
            }

            "RemoveSharePointGroupAssignment" {
                if (`$DryRun) {
                    `$Message = "[DRY-RUN] Usunelby grupe SP '`$(`$Operation.SharePointGroupName)' z ACL '`$(`$Operation.FullUrl)'"
                }
                else {
                    `$SPSite = `$null
                    `$SPWeb = `$null
                    try {
                        `$SPSite = New-Object Microsoft.SharePoint.SPSite(`$SiteUrl)
                        `$SPWeb = `$SPSite.OpenWeb(`$WebUrl.Replace(`$SiteUrl, ""))
                        `$SPWeb.AllowUnsafeUpdates = `$true

                        # Znajdź grupę
                        `$Group = `$SPWeb.SiteGroups[`$Operation.SharePointGroupName]
                        if (`$null -eq `$Group) {
                            # Próbuj przez LoginName
                            `$Group = `$SPWeb.SiteUsers[`$LoginName]
                        }

                        if (`$ListId -and `$ItemId) {
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$Item = `$List.GetItemById([int]`$ItemId)
                            `$Item.RoleAssignments.Remove(`$Group)
                            `$Item.Update()
                            `$Message = "Usunieto grupe '`$(`$Operation.SharePointGroupName)' z elementu ID `$ItemId"
                        }
                        elseif (`$ListId) {
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$List.RoleAssignments.Remove(`$Group)
                            `$List.Update()
                            `$Message = "Usunieto grupe '`$(`$Operation.SharePointGroupName)' z listy '`$(`$List.Title)'"
                        }
                        else {
                            `$SPWeb.RoleAssignments.Remove(`$Group)
                            `$SPWeb.Update()
                            `$Message = "Usunieto grupe '`$(`$Operation.SharePointGroupName)' z witryny '`$(`$SPWeb.Url)'"
                        }

                        `$SPWeb.AllowUnsafeUpdates = `$false
                    }
                    finally {
                        if (`$SPWeb) { `$SPWeb.Dispose() }
                        if (`$SPSite) { `$SPSite.Dispose() }
                    }
                }
            }

            "RemoveDomainGroupAssignment" {
                if (`$DryRun) {
                    `$Message = "[DRY-RUN] Usunelby grupe domenowa '`$LoginName' z ACL '`$(`$Operation.FullUrl)'"
                }
                else {
                    `$SPSite = `$null
                    `$SPWeb = `$null
                    try {
                        `$SPSite = New-Object Microsoft.SharePoint.SPSite(`$SiteUrl)
                        `$SPWeb = `$SPSite.OpenWeb(`$WebUrl.Replace(`$SiteUrl, ""))
                        `$SPWeb.AllowUnsafeUpdates = `$true

                        `$DomainGroup = `$SPWeb.EnsureUser(`$LoginName)

                        if (`$ListId -and `$ItemId) {
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$Item = `$List.GetItemById([int]`$ItemId)
                            `$Item.RoleAssignments.Remove(`$DomainGroup)
                            `$Item.Update()
                            `$Message = "Usunieto grupe domenowa '`$LoginName' z elementu ID `$ItemId"
                        }
                        elseif (`$ListId) {
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$List.RoleAssignments.Remove(`$DomainGroup)
                            `$List.Update()
                            `$Message = "Usunieto grupe domenowa '`$LoginName' z listy '`$(`$List.Title)'"
                        }
                        else {
                            `$SPWeb.RoleAssignments.Remove(`$DomainGroup)
                            `$SPWeb.Update()
                            `$Message = "Usunieto grupe domenowa '`$LoginName' z witryny '`$(`$SPWeb.Url)'"
                        }

                        `$SPWeb.AllowUnsafeUpdates = `$false
                    }
                    finally {
                        if (`$SPWeb) { `$SPWeb.Dispose() }
                        if (`$SPSite) { `$SPSite.Dispose() }
                    }
                }
            }

            "RestoreInheritance" {
                if (`$DryRun) {
                    `$Message = "[DRY-RUN] Przywrocilby dziedziczenie uprawnien dla '`$(`$Operation.FullUrl)'"
                }
                else {
                    `$SPSite = `$null
                    `$SPWeb = `$null
                    try {
                        `$SPSite = New-Object Microsoft.SharePoint.SPSite(`$SiteUrl)
                        `$SPWeb = `$SPSite.OpenWeb(`$WebUrl.Replace(`$SiteUrl, ""))
                        `$SPWeb.AllowUnsafeUpdates = `$true

                        if (`$ListId -and `$ItemId) {
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$Item = `$List.GetItemById([int]`$ItemId)
                            `$Item.ResetRoleInheritance()
                            `$Item.Update()
                            `$Message = "Przywrocono dziedziczenie dla elementu ID `$ItemId w liscie '`$(`$List.Title)'"
                        }
                        elseif (`$ListId) {
                            `$List = `$SPWeb.Lists[[Guid]`$ListId]
                            `$List.ResetRoleInheritance()
                            `$List.Update()
                            `$Message = "Przywrocono dziedziczenie dla listy '`$(`$List.Title)'"
                        }
                        else {
                            `$SPWeb.ResetRoleInheritance()
                            `$SPWeb.Update()
                            `$Message = "Przywrocono dziedziczenie dla witryny '`$(`$SPWeb.Url)'"
                        }

                        `$SPWeb.AllowUnsafeUpdates = `$false
                    }
                    finally {
                        if (`$SPWeb) { `$SPWeb.Dispose() }
                        if (`$SPSite) { `$SPSite.Dispose() }
                    }
                }
            }

            default {
                `$Status = "Unknown"
                `$Message = "Nieznana akcja: `$(`$Operation.Action)"
            }
        }
    }
    catch {
        `$Status = "Error"
        `$Message = "BLAD: `$_"
        throw
    }

    return [PSCustomObject]@{
        Status    = `$Status
        Operation = `$Operation.Action
        Url       = `$Operation.FullUrl
        Principal = `$Operation.PrincipalLoginName
        Message   = `$Message
        Timestamp = (Get-Date -Format "o")
    }
}

# ============================================================
# PODSUMOWANIE
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PODSUMOWANIE REMEDIACJI" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Sukces  : `$SuccessCount" -ForegroundColor Green
Write-Host "  Bledy   : `$ErrorCount" -ForegroundColor $(if (`$ErrorCount -gt 0) { 'Red' } else { 'Green' })
Write-Host "  Pominieto: `$SkippedCount" -ForegroundColor Yellow
Write-Host "  Transcript: `$TranscriptPath" -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Eksport wynikow do JSON
`$ResultsPath = Join-Path (Split-Path `$TranscriptPath -Parent) "RemediationResults_`$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').json"
`$Results | ConvertTo-Json -Depth 5 | Out-File -FilePath `$ResultsPath -Encoding UTF8

Write-Host "Wyniki zapisane: `$ResultsPath" -ForegroundColor Gray

Stop-Transcript

# ============================================================
# ROLLBACK (INSTRUKCJA)
# ============================================================
# Aby cofnac zmiany:
# 1. Otwierz plik: `$ResultsPath
# 2. Dla kazdej operacji "RemoveDirectUserPermission":
#    -> Uruchom: `$web = Get-SPWeb -Identity [WebUrl]
#    ->           `$user = `$web.EnsureUser("[LoginName]")
#    ->           `$rol = [Microsoft.SharePoint.SPRoleDefinition]
#    ->           `$ra = New-Object Microsoft.SharePoint.SPRoleAssignment(`$user)
#    ->           `$ra.RoleDefinitionBindings.Add(`$web.RoleDefinitions["[PermissionLevel]"])
#    ->           [Obiekt].RoleAssignments.Add(`$ra)
# 3. Dla "RestoreInheritance":
#    -> Uzyj kopii zapasowej Site Collection lub Recycle Bin jesli dostepny
"@

    return $ScriptContent
}

Export-ModuleMember -Function @(
    "New-RemediationScript"
)

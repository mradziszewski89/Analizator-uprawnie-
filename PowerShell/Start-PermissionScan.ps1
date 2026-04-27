#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Permission Analyzer - Główny skrypt skanowania uprawnień farmy SharePoint SE on-premises.

.DESCRIPTION
    Skanuje całą farmę SharePoint SE i generuje raport uprawnień na poziomach:
    WebApplication > SiteCollection > Web > List/Library > Folder > File/Item
    Eksportuje dane do JSON, CSV oraz interaktywnego raportu HTML.

.PARAMETER ConfigPath
    Ścieżka do pliku konfiguracyjnego JSON. Domyślnie: .\Config\ScanConfig.json

.PARAMETER ExclusionsPath
    Ścieżka do pliku wykluczeń JSON. Domyślnie: .\Config\Exclusions.json

.PARAMETER WhitelistPath
    Ścieżka do pliku whitelist JSON. Domyślnie: .\Config\Whitelist.json

.PARAMETER OutputPath
    Nadpisuje OutputPath z konfiguracji.

.PARAMETER ResumeFromCheckpoint
    Wznawia skanowanie z ostatniego checkpointu.

.PARAMETER DryRun
    Tylko symuluje skanowanie bez zapisu wyników.

.PARAMETER Verbose
    Szczegółowe logowanie.

.EXAMPLE
    # Pełne skanowanie farmy
    .\Start-PermissionScan.ps1

    # Skanowanie z wznowieniem z checkpointu
    .\Start-PermissionScan.ps1 -ResumeFromCheckpoint

    # Skanowanie z automatyczna publikacja raportu do biblioteki SharePoint
    .\Start-PermissionScan.ps1 -SharePointLibraryUrl "http://portal.contoso.com/sites/Raporty/Documents"

    # Skanowanie konkretnej Web Application
    $env:SP_SCAN_WEBAPP = "http://portal.contoso.com"
    .\Start-PermissionScan.ps1

.PARAMETER SharePointLibraryUrl
    URL biblioteki lub folderu SharePoint, do ktorego zostanie opublikowany raport HTML.
    Przyklad: http://portal.contoso.com/sites/Raporty/Documents
    Gdy parametr jest podany, raport jest publikowany automatycznie bez interaktywnych pytan.
    Nadpisuje wartosc SharePointPublishUrl z pliku konfiguracyjnego.

.NOTES
    Wymagania: PowerShell 5.1, SharePoint SE on-premises, uprawnienia Farm Administrator.
    Uruchamiaj na serwerze SharePoint z załadowanym snap-in Microsoft.SharePoint.PowerShell.
    Autor: SharePoint Permission Analyzer v1.0
    Data: 2025
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "",

    [Parameter(Mandatory = $false)]
    [string]$ExclusionsPath = "",

    [Parameter(Mandatory = $false)]
    [string]$WhitelistPath = "",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",

    [Parameter(Mandatory = $false)]
    [switch]$ResumeFromCheckpoint,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$SkipItemLevelScan,

    [Parameter(Mandatory = $false)]
    [switch]$RawAssignmentsOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ExpandDomainGroups,

    [Parameter(Mandatory = $false)]
    [string]$SharePointLibraryUrl = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ============================================================
# SEKCJA 1: Inicjalizacja i walidacja środowiska
# ============================================================

$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) {
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SharePoint Permission Analyzer v1.0" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Sprawdź czy SharePoint PowerShell jest załadowany
if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    if (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -Registered -ErrorAction SilentlyContinue) {
        try {
            Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
            Write-Host "[OK] Zaladowano SharePoint PowerShell snap-in" -ForegroundColor Green
        }
        catch {
            # Moze byc juz zaladowany przez inny mechanizm (np. SharePoint Management Shell)
            Write-Host "[INFO] PSSnapin nie zostal zaladowany (moze byc aktywny przez inny mechanizm): $_" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "[INFO] PSSnapin Microsoft.SharePoint.PowerShell nie jest zarejestrowany - kontynuuję" -ForegroundColor Yellow
    }
}
else {
    Write-Host "[OK] SharePoint PowerShell snap-in juz zaladowany" -ForegroundColor Green
}

# Załaduj moduły pomocnicze
$ModulesPath = Join-Path $ScriptRoot "Modules"

$RequiredModules = @(
    "SPPermissionScanner.psm1",
    "ADGroupExpander.psm1",
    "ReportGenerator.psm1",
    "RemediationScriptGenerator.psm1"
)

foreach ($ModuleFile in $RequiredModules) {
    $ModulePath = Join-Path $ModulesPath $ModuleFile
    if (-not (Test-Path $ModulePath)) {
        Write-Error "BLAD KRYTYCZNY: Nie znaleziono modulu: $ModulePath"
        exit 1
    }
    try {
        Import-Module $ModulePath -Force -ErrorAction Stop
        Write-Host "[OK] Zaladowano modul: $ModuleFile" -ForegroundColor Green
    }
    catch {
        Write-Error "BLAD KRYTYCZNY: Nie mozna zaladowac modulu $ModuleFile`: $_"
        exit 1
    }
}

# Rozwiąż ścieżki konfiguracji
function Read-YesNoPrompt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $false)]
        [bool]$Default = $false
    )

    $suffix = if ($Default) { " [Y/n]" } else { " [y/N]" }

    while ($true) {
        $response = Read-Host "$Prompt$suffix"
        if ([string]::IsNullOrWhiteSpace($response)) {
            return $Default
        }

        switch -Regex ($response.Trim().ToLowerInvariant()) {
            '^(y|yes|t|tak)$' { return $true }
            '^(n|no|nie)$' { return $false }
            default {
                Write-Host "Wpisz 'y'/'tak' albo 'n'/'nie'." -ForegroundColor Yellow
            }
        }
    }
}

function Normalize-SPPublishUrlPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Uri]$Uri
    )

    $path = [System.Uri]::UnescapeDataString($Uri.AbsolutePath)
    $path = $path -replace '/Forms/[^/]+\.aspx$', ''
    $path = $path.TrimEnd('/')

    if (-not $path) {
        return '/'
    }

    if (-not $path.StartsWith('/')) {
        $path = "/$path"
    }

    return $path
}

function Resolve-SPPublishTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LibraryUrl
    )

    $uri = [System.Uri]$LibraryUrl
    $absolutePath = Normalize-SPPublishUrlPath -Uri $uri
    $authority = $uri.GetLeftPart([System.UriPartial]::Authority)
    $segments = @()

    if ($absolutePath -ne '/') {
        $segments = $absolutePath.Trim('/').Split('/')
    }

    $web = $null
    for ($i = $segments.Count; $i -ge 0; $i--) {
        $candidatePath = if ($i -gt 0) {
            '/' + ($segments[0..($i - 1)] -join '/')
        }
        else {
            '/'
        }

        try {
            $web = Get-SPWeb -Identity ($authority + $candidatePath) -ErrorAction Stop
            break
        }
        catch {
            $web = $null
        }
    }

    if (-not $web) {
        throw "Nie mozna odnalezc witryny SharePoint dla URL: $LibraryUrl"
    }

    try {
        $targetFolder = $web.GetFolder($absolutePath)
        if (-not $targetFolder.Exists) {
            throw "Docelowa biblioteka lub folder nie istnieje: $absolutePath"
        }

        return @{
            Web = $web
            TargetFolder = $targetFolder
            Authority = $authority
        }
    }
    catch {
        if ($web) {
            $web.Dispose()
        }
        throw
    }
}

function Ensure-SPFolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPWeb]$Web,

        [Parameter(Mandatory = $true)]
        [string]$ServerRelativePath
    )

    $normalizedPath = [System.Uri]::UnescapeDataString(($ServerRelativePath -replace '\\', '/')).TrimEnd('/')
    if (-not $normalizedPath) {
        $normalizedPath = '/'
    }
    if (-not $normalizedPath.StartsWith('/')) {
        $normalizedPath = "/$normalizedPath"
    }

    $rootPath = [System.Uri]::UnescapeDataString($Web.RootFolder.ServerRelativeUrl).TrimEnd('/')
    if (-not $rootPath) {
        $rootPath = '/'
    }

    if ($normalizedPath -eq $rootPath) {
        return $Web.RootFolder
    }

    if (-not $normalizedPath.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Sciezka '$normalizedPath' nie nalezy do witryny '$($Web.Url)'."
    }

    $relativePath = $normalizedPath.Substring($rootPath.Length).Trim('/')
    if (-not $relativePath) {
        return $Web.RootFolder
    }

    $currentFolder = $Web.RootFolder
    $currentPath = $rootPath

    foreach ($segment in $relativePath.Split('/')) {
        if ([string]::IsNullOrWhiteSpace($segment)) {
            continue
        }

        $currentPath = if ($currentPath -eq '/') {
            "/$segment"
        }
        else {
            "$currentPath/$segment"
        }

        $nextFolder = $Web.GetFolder($currentPath)
        if (-not $nextFolder.Exists) {
            $null = $currentFolder.SubFolders.Add($segment)
            $nextFolder = $Web.GetFolder($currentPath)
        }

        $currentFolder = $nextFolder
    }

    return $currentFolder
}

function Publish-ReportToSharePointLibrary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LibraryUrl,

        [Parameter(Mandatory = $true)]
        [string]$ReportFolderPath
    )

    if (-not (Test-Path -Path $ReportFolderPath -PathType Container)) {
        throw "Nie znaleziono folderu raportu: $ReportFolderPath"
    }

    $publishTarget = Resolve-SPPublishTarget -LibraryUrl $LibraryUrl
    $web = $publishTarget.Web

    try {
        $reportFolderName = Split-Path -Path $ReportFolderPath -Leaf
        if ([string]::IsNullOrWhiteSpace($reportFolderName)) {
            throw "Nie mozna ustalic nazwy folderu raportu dla: $ReportFolderPath"
        }

        $targetRootPath = ($publishTarget.TargetFolder.ServerRelativeUrl.TrimEnd('/') + '/' + $reportFolderName) -replace '//+', '/'
        $null = Ensure-SPFolderPath -Web $web -ServerRelativePath $targetRootPath

        # Tylko pliki będące zasobami webowymi - pomijaj logi, pliki tymczasowe itp.
        $webExtensions = @('.html', '.htm', '.js', '.css', '.png', '.gif', '.jpg', '.jpeg', '.ico', '.svg', '.woff', '.woff2', '.ttf', '.eot', '.map')
        $files = Get-ChildItem -Path $ReportFolderPath -File -Recurse |
            Where-Object { $_.Extension -in $webExtensions } |
            Sort-Object FullName
        if (-not $files) {
            throw "Folder raportu nie zawiera plikow webowych do publikacji: $ReportFolderPath"
        }

        $fileIndex = 0
        $skippedCount = 0
        foreach ($file in $files) {
            $fileIndex++
            $relativePath = $file.FullName.Substring($ReportFolderPath.Length).TrimStart([char[]]@('\', '/'))
            $relativeDirectory = Split-Path -Path $relativePath -Parent
            $targetFolderPath = $targetRootPath

            if ($relativeDirectory -and $relativeDirectory -ne '.') {
                $targetFolderPath = ($targetRootPath.TrimEnd('/') + '/' + ($relativeDirectory -replace '\\', '/')) -replace '//+', '/'
            }

            try {
                $targetFolder = Ensure-SPFolderPath -Web $web -ServerRelativePath $targetFolderPath
                $fileBytes = [System.IO.File]::ReadAllBytes($file.FullName)
                $spFile = $targetFolder.Files.Add($file.Name, $fileBytes, $true)

                # Jeśli biblioteka ma wersjonowanie minor (Draft), opublikuj jako wersję główną (Major).
                # Files.Add z overwrite=true na istniejącym pliku tworzy szkic zamiast nadpisać wersję opublikowaną.
                if ($spFile.Level -eq [Microsoft.SharePoint.SPFileLevel]::Draft) {
                    $spFile.Publish('')
                }
                # Jeśli biblioteka wymaga zatwierdzenia treści, zatwierdź plik.
                try {
                    $modInfo = $spFile.Item.ModerationInformation
                    if ($modInfo -ne $null -and $modInfo.Status -ne [Microsoft.SharePoint.SPModerationStatusType]::Approved) {
                        $spFile.Approve('')
                    }
                } catch { }

                Write-Host ("  [{0}/{1}] {2}" -f $fileIndex, $files.Count, ($relativePath -replace '\\', '/')) -ForegroundColor Gray
            }
            catch {
                $skippedCount++
                Write-Warning ("  [{0}/{1}] POMINIETY: {2} - {3}" -f $fileIndex, $files.Count, ($relativePath -replace '\\', '/'), $_.Exception.Message)
                Write-ScanLog -Level "Warning" -Message ("Pominiety plik podczas publikacji: {0} - {1}" -f ($relativePath -replace '\\', '/'), $_.Exception.Message)
            }
        }
        if ($skippedCount -gt 0) {
            Write-Host "  [WARN] Pominięto $skippedCount plik(ów) z powodu błędów odczytu." -ForegroundColor Yellow
        }

        # SP2016+ wymaga dwoch ustawien aby HTML otwierac w przegladarce (nie pobierac):
        #   1. BrowserFileHandling = Permissive  (int 1, uzywamy int bo enum moze byc niedostepny)
        #   2. AllowedInlineDownloadedMimeTypes musi zawierac "text/html"
        # Po kazdej zmianie wymagany jest recykl puli aplikacji IIS (lub iisreset).
        try {
            $targetWebApp = $web.Site.WebApplication
            $waChanged = $false

            if ([int]$targetWebApp.BrowserFileHandling -ne 1) {
                Write-Host "  [INFO] BrowserFileHandling = Strict - zmiana na Permissive" -ForegroundColor Yellow
                $targetWebApp.BrowserFileHandling = 1
                $waChanged = $true
            }

            if ($targetWebApp.AllowedInlineDownloadedMimeTypes -notcontains 'text/html') {
                Write-Host "  [INFO] Dodawanie text/html do AllowedInlineDownloadedMimeTypes (SP2016+)" -ForegroundColor Yellow
                $targetWebApp.AllowedInlineDownloadedMimeTypes.Add('text/html')
                $waChanged = $true
            }

            if ($waChanged) {
                $targetWebApp.Update()
                Write-Host "  [OK] Ustawienia Web Application zaktualizowane" -ForegroundColor Green
                Write-Host "  [!] Wymagany recykl puli IIS lub iisreset aby zmiany weszly w zycie." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Warning "Nie mozna zmienic ustawien Web Application: $_"
            Write-Host ""
            Write-Host "  [!] WYMAGANA AKCJA RECZNA (SharePoint Management Shell jako Farm Admin):" -ForegroundColor Red
            Write-Host "  `$wa = Get-SPWebApplication 'https://portalse.test.pl'" -ForegroundColor Yellow
            Write-Host "  `$wa.BrowserFileHandling = 1" -ForegroundColor Yellow
            Write-Host "  `$wa.AllowedInlineDownloadedMimeTypes.Add('text/html')" -ForegroundColor Yellow
            Write-Host "  `$wa.Update()" -ForegroundColor Yellow
            Write-Host "  iisreset" -ForegroundColor Yellow
            Write-Host ""
        }

        return @{
            FileCount = $files.Count
            ReportFolderUrl = $publishTarget.Authority + $targetRootPath
            IndexUrl = $publishTarget.Authority + $targetRootPath.TrimEnd('/') + '/index.html'
        }
    }
    finally {
        if ($web) {
            $web.Dispose()
        }
    }
}
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $ScriptRoot "Config\ScanConfig.json"
}
if (-not $ExclusionsPath) {
    $ExclusionsPath = Join-Path $ScriptRoot "Config\Exclusions.json"
}
if (-not $WhitelistPath) {
    $WhitelistPath = Join-Path $ScriptRoot "Config\Whitelist.json"
}

# Walidacja plików konfiguracji
foreach ($ConfigFile in @($ConfigPath, $ExclusionsPath, $WhitelistPath)) {
    if (-not (Test-Path $ConfigFile)) {
        Write-Error "BLAD: Nie znaleziono pliku konfiguracyjnego: $ConfigFile"
        exit 1
    }
}

# Wczytaj konfigurację
try {
    $Config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $Exclusions = Get-Content -Path $ExclusionsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $Whitelist = Get-Content -Path $WhitelistPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-Host "[OK] Wczytano pliki konfiguracyjne" -ForegroundColor Green
}
catch {
    Write-Error "BLAD KRYTYCZNY: Nie mozna wczytac plikow konfiguracyjnych: $_"
    exit 1
}

# Parametry wiersza polecenia nadpisują konfigurację
if ($OutputPath) {
    $Config.Output.OutputPath = $OutputPath
}
if ($SkipItemLevelScan) {
    $Config.ScanDepth.ScanFiles = $false
    $Config.ScanDepth.ScanListItems = $false
    Write-Host "[INFO] Tryb: Pomijanie skanowania na poziomie elementow (SkipItemLevelScan)" -ForegroundColor Yellow
}
if ($RawAssignmentsOnly) {
    $Config.PrincipalExpansion.RawAssignmentsOnly = $true
    $Config.PrincipalExpansion.ExpandSharePointGroups = $false
    $Config.PrincipalExpansion.ExpandDomainGroups = $false
    Write-Host "[INFO] Tryb: Tylko surowe przypisania SharePoint (RawAssignmentsOnly)" -ForegroundColor Yellow
}
if ($ExpandDomainGroups) {
    $Config.PrincipalExpansion.ExpandDomainGroups = $true
    Write-Host "[INFO] Tryb: Ekspansja grup domenowych wlaczona (ExpandDomainGroups)" -ForegroundColor Yellow
}

# Ustal ścieżki wyjściowe
if (-not $Config.Output.OutputPath) {
    $Config.Output.OutputPath = Join-Path $ScriptRoot "Output"
}
if (-not $Config.Logging.LogPath) {
    $Config.Logging.LogPath = Join-Path $ScriptRoot "Logs"
}
if (-not $Config.Performance.CheckpointPath) {
    $Config.Performance.CheckpointPath = Join-Path $ScriptRoot "Logs"
}

# Utwórz katalogi wyjściowe
foreach ($Dir in @($Config.Output.OutputPath, $Config.Logging.LogPath, $Config.Performance.CheckpointPath)) {
    if (-not (Test-Path $Dir)) {
        New-Item -ItemType Directory -Path $Dir -Force | Out-Null
    }
}

# Nazwa sesji skanowania
$ScanSessionId = [System.Guid]::NewGuid().ToString("N").Substring(0, 8).ToUpper()
$ScanStartTime = Get-Date
$ScanTimestamp = $ScanStartTime.ToString("yyyy-MM-dd_HH-mm-ss")
$ReportBaseName = "$($Config.Output.ReportName)_$ScanTimestamp"

# Inicjalizuj logger
$LogFilePath = Join-Path $Config.Logging.LogPath "Scan_$ScanTimestamp.log"
Initialize-ScanLogger -LogFilePath $LogFilePath -LogLevel $Config.Logging.LogLevel

Write-ScanLog -Level "Info" -Message "============================================================"
Write-ScanLog -Level "Info" -Message "SharePoint Permission Analyzer v1.0 - Sesja: $ScanSessionId"
Write-ScanLog -Level "Info" -Message "Poczatek skanowania: $($ScanStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-ScanLog -Level "Info" -Message "Konfiguracja: $ConfigPath"
Write-ScanLog -Level "Info" -Message "============================================================"

# ============================================================
# SEKCJA 2: Weryfikacja uprawnień wykonującego
# ============================================================

if ($Config.Security.RequireFarmAdminRole) {
    Write-Host ""
    Write-Host "Weryfikacja uprawnien administratora farmy..." -ForegroundColor Cyan

    try {
        $FarmAdminCheck = Test-SPFarmAdminRole
        if (-not $FarmAdminCheck) {
            Write-Error "BLAD BEZPIECZENSTWA: Biezacy uzytkownik nie jest administratorem farmy SharePoint.`nUruchom skrypt jako konto z uprawnieniami Farm Administrator."
            exit 1
        }
        Write-Host "[OK] Weryfikacja uprawnien Farm Administrator pomyslna" -ForegroundColor Green
        Write-ScanLog -Level "Info" -Message "Weryfikacja uprawnien Farm Admin: OK"
    }
    catch {
        Write-Warning "Nie mozna zweryfikowac uprawnien Farm Admin (kontynuowanie): $_"
        Write-ScanLog -Level "Warning" -Message "Weryfikacja uprawnien Farm Admin: $_"
    }
}

# ============================================================
# SEKCJA 3: Skanowanie
# ============================================================

Write-Host ""
Write-Host "Uruchamianie skanowania uprawnien farmy SharePoint..." -ForegroundColor Cyan
Write-Host "  Sesja: $ScanSessionId" -ForegroundColor Gray
Write-Host "  Timestamp: $ScanTimestamp" -ForegroundColor Gray
Write-Host "  Log: $LogFilePath" -ForegroundColor Gray
Write-Host ""

if ($DryRun) {
    Write-Host "[DRY-RUN] Tryb symulacji - zadnych zmian nie zostanie wprowadzonych" -ForegroundColor Yellow
    Write-ScanLog -Level "Info" -Message "TRYB DRY-RUN: Symulacja skanowania"
}

# Parametry skanowania
$ScanParams = @{
    Config             = $Config
    Exclusions         = $Exclusions
    Whitelist          = $Whitelist
    ScanSessionId      = $ScanSessionId
    ResumeFromCheckpoint = $ResumeFromCheckpoint.IsPresent
    CheckpointPath     = $Config.Performance.CheckpointPath
    DryRun             = $DryRun.IsPresent
}

try {
    # Uruchom główny scan
    $ScanResult = Invoke-SPFarmScan @ScanParams

    if ($null -eq $ScanResult) {
        Write-Error "BLAD: Skanowanie zwrocilo pusty wynik"
        exit 1
    }

    $ScanEndTime = Get-Date
    $ScanDuration = $ScanEndTime - $ScanStartTime

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  Skanowanie zakonczone pomyslnie!" -ForegroundColor Green
    Write-Host "  Czas trwania: $($ScanDuration.ToString('hh\:mm\:ss'))" -ForegroundColor Green
    Write-Host "  Obiekty przeskanowane: $($ScanResult.Statistics.TotalObjectsScanned)" -ForegroundColor Green
    Write-Host "  Przypisania zebrane: $($ScanResult.Statistics.TotalAssignments)" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green

    Write-ScanLog -Level "Info" -Message "Skanowanie zakonczone. Czas: $($ScanDuration.ToString('hh\:mm\:ss')). Obiektow: $($ScanResult.Statistics.TotalObjectsScanned). Przypizan: $($ScanResult.Statistics.TotalAssignments)"
}
catch {
    Write-Host ""
    Write-Host "[BLAD KRYTYCZNY] Skanowanie przerwane: $_" -ForegroundColor Red
    Write-ScanLog -Level "Error" -Message "BLAD KRYTYCZNY: $_`nStackTrace: $($_.ScriptStackTrace)"
    exit 1
}

# ============================================================
# SEKCJA 4: Eksport wyników
# ============================================================

Write-Host ""
Write-Host "Eksportowanie wynikow..." -ForegroundColor Cyan

# Uzupełnij metadane skanu
$ScanResult.ScanMetadata.ScanEndTime = $ScanEndTime.ToString("o")
$ScanResult.ScanMetadata.ScanDuration = [int]$ScanDuration.TotalSeconds
$ScanResult.ScanMetadata.ScanSessionId = $ScanSessionId
$ScanResult.ScanMetadata.LogFilePath = $LogFilePath

# Eksport JSON
if ($Config.Output.ExportJson) {
    $JsonPath = Join-Path $Config.Output.OutputPath "$ReportBaseName.json"
    try {
        $JsonDepth = 10
        if ($Config.Output.PrettifyJson) {
            $ScanResult | ConvertTo-Json -Depth $JsonDepth | Out-File -FilePath $JsonPath -Encoding UTF8
        }
        else {
            $ScanResult | ConvertTo-Json -Depth $JsonDepth -Compress | Out-File -FilePath $JsonPath -Encoding UTF8
        }
        Write-Host "[OK] Eksport JSON: $JsonPath" -ForegroundColor Green
        Write-ScanLog -Level "Info" -Message "Eksport JSON: $JsonPath"
    }
    catch {
        Write-Warning "Blad eksportu JSON: $_"
        Write-ScanLog -Level "Error" -Message "Blad eksportu JSON: $_"
    }
}

# Eksport CSV
if ($Config.Output.ExportCsv) {
    $CsvPath = Join-Path $Config.Output.OutputPath "$ReportBaseName.csv"
    try {
        Export-ScanResultToCsv -ScanResult $ScanResult -OutputPath $CsvPath
        Write-Host "[OK] Eksport CSV: $CsvPath" -ForegroundColor Green
        Write-ScanLog -Level "Info" -Message "Eksport CSV: $CsvPath"
    }
    catch {
        Write-Warning "Blad eksportu CSV: $_"
        Write-ScanLog -Level "Error" -Message "Blad eksportu CSV: $_"
    }
}

# Generowanie raportu HTML
if ($Config.Output.GenerateHtmlReport) {
    # Stala nazwa bez znacznika czasu - kazde uruchomienie nadpisuje poprzedni raport
    $HtmlOutputDir = Join-Path $Config.Output.OutputPath "Report"

    # Odczyt poprzedniego raportu do porownania (wersjonowanie)
    $DiffData = $null
    $PreviousGenerated = $null
    $PreviousDataJs = Join-Path $HtmlOutputDir "data.js"
    if (Test-Path $PreviousDataJs) {
        Write-Host "  [INFO] Znaleziono poprzedni raport - generowanie zestawienia zmian..." -ForegroundColor Cyan
        try {
            $PrevContent = Get-Content -Path $PreviousDataJs -Raw -Encoding UTF8
            $B64Match = [regex]::Match($PrevContent, 'var _b64 = "([A-Za-z0-9+/=]+)"')
            if ($B64Match.Success) {
                $PrevB64 = $B64Match.Groups[1].Value
                $PrevBytes = [System.Convert]::FromBase64String($PrevB64)
                $PrevJson = [System.Text.Encoding]::UTF8.GetString($PrevBytes)
                $PreviousScanData = $PrevJson | ConvertFrom-Json

                $PreviousGenerated = $null
                $GenMatch = [regex]::Match($PrevContent, 'window\.REPORT_GENERATED = "([^"]+)"')
                if ($GenMatch.Success) { $PreviousGenerated = $GenMatch.Groups[1].Value }

                $DiffData = Compare-PermissionScanResults `
                    -Previous $PreviousScanData `
                    -Current $ScanResult `
                    -PreviousGenerated $PreviousGenerated

                $addC = $DiffData.AddedObjects.Count
                $remC = $DiffData.RemovedObjects.Count
                $chgC = $DiffData.ChangedObjects.Count
                Write-Host "  [OK] Roznice: +$addC nowych, -$remC usunietych, ~$chgC zmienionych obiektow" -ForegroundColor Green
                Write-ScanLog -Level "Info" -Message "Wersjonowanie: +$addC / -$remC / ~$chgC"
            }
            else {
                Write-Host "  [WARN] Nie mozna odczytac danych z poprzedniego data.js" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Warning "Nie mozna porownac z poprzednim raportem: $_"
            Write-ScanLog -Level "Warning" -Message "Blad porownania raportow: $_"
        }
    }
    else {
        Write-Host "  [INFO] Brak poprzedniego raportu - to bedzie raport bazowy dla wersjonowania" -ForegroundColor Cyan
    }

    # Archiwizuj poprzedni raport (przed nadpisaniem przez nowy)
    if ((Test-Path $HtmlOutputDir) -and (Test-Path (Join-Path $HtmlOutputDir 'data.js'))) {
        try {
            $archiveGenerated = $PreviousGenerated
            if (-not $archiveGenerated) {
                $archiveDjs = Get-Content (Join-Path $HtmlOutputDir 'data.js') -Raw -Encoding UTF8
                $archiveGm = [regex]::Match($archiveDjs, 'window\.REPORT_GENERATED = "([^"]+)"')
                if ($archiveGm.Success) { $archiveGenerated = $archiveGm.Groups[1].Value }
            }
            if ($archiveGenerated) {
                $archiveTs = [datetime]::ParseExact($archiveGenerated, 'yyyy-MM-dd HH:mm:ss', $null).ToString('yyyy-MM-dd_HH-mm-ss')
                $archiveFolder = Join-Path $Config.Output.OutputPath "Report_$archiveTs"
                if (-not (Test-Path $archiveFolder)) {
                    Copy-Item -Path $HtmlOutputDir -Destination $archiveFolder -Recurse -Force
                    Write-Host "  [OK] Zarchiwizowano raport: Report_$archiveTs" -ForegroundColor Green
                    Write-ScanLog -Level "Info" -Message "Zarchiwizowano raport: Report_$archiveTs"
                }
            }
        }
        catch {
            Write-Warning "Nie mozna zarchiwizowac poprzedniego raportu: $_"
            Write-ScanLog -Level "Warning" -Message "Blad archiwizacji raportu: $_"
        }
    }

    # Utrzymuj maks. 30 archiwow raportow
    try {
        $allArchive = Get-ChildItem $Config.Output.OutputPath -Directory |
            Where-Object { $_.Name -match '^Report_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}$' } |
            Sort-Object Name
        if ($allArchive.Count -gt 30) {
            $toRemove = $allArchive | Select-Object -First ($allArchive.Count - 30)
            foreach ($old in $toRemove) {
                Remove-Item $old.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch { }

    # Zbierz historię ostatnich 7 archiwalnych raportów (posortowane od najstarszego)
    $ScanHistory = @()
    try {
        $histFolders = Get-ChildItem $Config.Output.OutputPath -Directory |
            Where-Object { $_.Name -match '^Report_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}$' } |
            Sort-Object Name -Descending |
            Select-Object -First 7 |
            Sort-Object Name
        foreach ($hf in $histFolders) {
            $hDataJs = Join-Path $hf.FullName 'data.js'
            if (Test-Path $hDataJs) {
                try {
                    $hContent = Get-Content $hDataJs -Raw -Encoding UTF8
                    $hb64m = [regex]::Match($hContent, 'var _b64 = "([A-Za-z0-9+/=]+)"')
                    $hgenm = [regex]::Match($hContent, 'window\.REPORT_GENERATED = "([^"]+)"')
                    if ($hb64m.Success) {
                        $hBytes = [Convert]::FromBase64String($hb64m.Groups[1].Value)
                        $hJson = [Text.Encoding]::UTF8.GetString($hBytes)
                        $hData = $hJson | ConvertFrom-Json
                        $hSizeRaw = (Get-ChildItem $hf.FullName -Recurse -File | Measure-Object -Property Length -Sum).Sum
                        $hSize = if ($null -eq $hSizeRaw) { 0L } else { [long]$hSizeRaw }
                        $ScanHistory += [ordered]@{
                            Generated              = if ($hgenm.Success) { $hgenm.Groups[1].Value } else { '' }
                            FolderName             = $hf.Name
                            FolderSizeBytes        = $hSize
                            TotalObjectsScanned    = [int]($hData.Statistics.TotalObjectsScanned)
                            TotalAssignments       = [int]($hData.Statistics.TotalAssignments)
                            UniquePermissionsCount = [int]($hData.Statistics.UniquePermissionsCount)
                            WebApplicationCount    = [int]($hData.Statistics.WebApplicationCount)
                            SiteCollectionCount    = [int]($hData.Statistics.SiteCollectionCount)
                            WebCount               = [int]($hData.Statistics.WebCount)
                            ListCount              = [int]($hData.Statistics.ListCount)
                            ItemCount              = [int]($hData.Statistics.ItemCount)
                        }
                    }
                }
                catch {
                    Write-Warning "Nie mozna odczytac historii z: $($hf.Name) - $_"
                }
            }
        }
        if ($ScanHistory.Count -gt 0) {
            Write-Host "  [OK] Historia raportow: $($ScanHistory.Count) wersji" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Blad zbierania historii raportow: $_"
        Write-ScanLog -Level "Warning" -Message "Blad historii raportow: $_"
    }

    # Ścieżka szablonu
    $TemplatePath = $Config.Output.HtmlTemplatePath
    if (-not $TemplatePath) {
        $TemplatePath = Join-Path (Split-Path $ScriptRoot -Parent) "Report"
    }

    try {
        $HtmlReportPath = New-HTMLPermissionReport `
            -ScanResult $ScanResult `
            -TemplatePath $TemplatePath `
            -OutputDirectory $HtmlOutputDir `
            -ReportTitle "Raport Uprawnien SharePoint - $($ScanResult.ScanMetadata.FarmName)" `
            -JsonDataPath $JsonPath `
            -DiffData $DiffData `
            -ScanHistory $ScanHistory

        Write-Host "[OK] Raport HTML: $HtmlOutputDir\index.html" -ForegroundColor Green
        Write-ScanLog -Level "Info" -Message "Raport HTML: $HtmlOutputDir\index.html"
    }
    catch {
        Write-Warning "Blad generowania raportu HTML: $_"
        Write-ScanLog -Level "Error" -Message "Blad generowania HTML: $_`n$($_.ScriptStackTrace)"
    }
}

# ============================================================
# SEKCJA 5: Podsumowanie końcowe
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PODSUMOWANIE SKANOWANIA" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Web Applications  : $($ScanResult.Statistics.WebApplicationCount)" -ForegroundColor White
Write-Host "  Site Collections  : $($ScanResult.Statistics.SiteCollectionCount)" -ForegroundColor White
Write-Host "  Witryny (Webs)    : $($ScanResult.Statistics.WebCount)" -ForegroundColor White
Write-Host "  Listy i biblioteki: $($ScanResult.Statistics.ListCount)" -ForegroundColor White
Write-Host "  Foldery           : $($ScanResult.Statistics.FolderCount)" -ForegroundColor White
Write-Host "  Pliki i elementy  : $($ScanResult.Statistics.ItemCount)" -ForegroundColor White
Write-Host "  Obiekty z unique  : $($ScanResult.Statistics.UniquePermissionsCount)" -ForegroundColor Yellow
Write-Host "  Wszystkie przyp.  : $($ScanResult.Statistics.TotalAssignments)" -ForegroundColor White
Write-Host "  Bledy skanowania  : $($ScanResult.Statistics.ErrorCount)" -ForegroundColor $(if ($ScanResult.Statistics.ErrorCount -gt 0) { "Red" } else { "Green" })
Write-Host ""
Write-Host "  Pliki wyjsciowe:" -ForegroundColor Cyan

if ($Config.Output.ExportJson -and (Test-Path $JsonPath)) {
    Write-Host "    JSON  : $JsonPath" -ForegroundColor Gray
}
if ($Config.Output.ExportCsv -and (Test-Path $CsvPath)) {
    Write-Host "    CSV   : $CsvPath" -ForegroundColor Gray
}
if ($Config.Output.GenerateHtmlReport -and (Test-Path $HtmlOutputDir)) {
    Write-Host "    HTML  : $HtmlOutputDir\index.html" -ForegroundColor Gray
}
Write-Host "    LOG   : $LogFilePath" -ForegroundColor Gray
Write-Host ""

if ($ScanResult.Statistics.ErrorCount -gt 0) {
    Write-Host "  [!] Wykryto bledy podczas skanowania. Sprawdz log: $LogFilePath" -ForegroundColor Red
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$PublishedReportInfo = $null
if ($Config.Output.GenerateHtmlReport -and (Test-Path $HtmlOutputDir)) {
    if ($DryRun) {
        Write-Host "[INFO] Tryb DRY-RUN - publikacja raportu do SharePoint zostala pominieta." -ForegroundColor Yellow
        Write-ScanLog -Level "Info" -Message "Pomijanie publikacji raportu do SharePoint (DryRun)"
    }
    else {
        # Ustal URL docelowy: parametr CLI > konfiguracja > brak (nie publikuj)
        if (-not [string]::IsNullOrWhiteSpace($SharePointLibraryUrl)) {
            $TargetLibraryUrl = $SharePointLibraryUrl
        } elseif (-not [string]::IsNullOrWhiteSpace($Config.Output.SharePointPublishUrl)) {
            $TargetLibraryUrl = $Config.Output.SharePointPublishUrl
        } else {
            $TargetLibraryUrl = ""
        }

        if (-not [string]::IsNullOrWhiteSpace($TargetLibraryUrl)) {
            Write-Host "[INFO] Publikacja raportu do SharePoint: $TargetLibraryUrl" -ForegroundColor Cyan
            Write-ScanLog -Level "Info" -Message "Rozpoczecie publikacji raportu HTML do SharePoint: $TargetLibraryUrl"
            try {
                $PublishedReportInfo = Publish-ReportToSharePointLibrary -LibraryUrl $TargetLibraryUrl -ReportFolderPath $HtmlOutputDir

                Write-Host "[OK] Raport opublikowany w SharePoint" -ForegroundColor Green
                Write-Host "    Folder: $($PublishedReportInfo.ReportFolderUrl)" -ForegroundColor Gray
                Write-Host "    HTML  : $($PublishedReportInfo.IndexUrl)" -ForegroundColor Gray
                Write-Host "    Pliki : $($PublishedReportInfo.FileCount)" -ForegroundColor Gray
                Write-Host "    Otworz index.html z biblioteki SharePoint, aby uzyc bezposredniej remediacji REST." -ForegroundColor Gray
                Write-Host ""

                Write-ScanLog -Level "Info" -Message "Raport opublikowany w SharePoint: $($PublishedReportInfo.IndexUrl)"
            }
            catch {
                Write-Warning "Blad publikacji raportu do SharePoint: $_"
                Write-ScanLog -Level "Error" -Message "Blad publikacji raportu do SharePoint: $_`n$($_.ScriptStackTrace)"
            }
        } else {
            Write-Host ""
            Write-Host "[INFO] Publikacja do SharePoint pominieta - brak skonfigurowanego URL." -ForegroundColor Yellow
            Write-Host "       Aby opublikowac raport podaj URL biblioteki na jeden z ponizszych sposobow:" -ForegroundColor Yellow
            Write-Host "       1. Parametr CLI:  .\Start-PermissionScan.ps1 -SharePointLibraryUrl 'http://portal/Dokumenty'" -ForegroundColor Yellow
            Write-Host "       2. Konfiguracja:  ustaw 'SharePointPublishUrl' w pliku Config\ScanConfig.json" -ForegroundColor Yellow
            Write-ScanLog -Level "Info" -Message "Publikacja do SharePoint pominieta - brak URL."
        }
    }
}
Write-ScanLog -Level "Info" -Message "Sesja $ScanSessionId zakonczona. Bledy: $($ScanResult.Statistics.ErrorCount)"
Write-Host "Gotowe." -ForegroundColor Green



#Requires -Version 5.1
<#
.SYNOPSIS
    Pobiera lokalne biblioteki front-end potrzebne do raportu HTML.
    Wymaga dostępu do Internetu tylko podczas tej jednorazowej instalacji.
    Po pobraniu raport działa w pełni offline.

.DESCRIPTION
    Pobiera:
    - jQuery 3.7.1
    - DataTables 1.13.8 (JS + CSS)
    - jsTree 3.3.16 (JS + CSS + ikony)
    - Chart.js 4.4.2
    Wszystkie pliki zapisuje w katalogu assets/ obok tego skryptu.

.NOTES
    Uruchom ten skrypt JEDEN RAZ podczas przygotowywania rozwiązania.
    Następnie skopiuj folder Report/ do serwera SharePoint (offline).
#>

$ReportDir = $PSScriptRoot
$AssetsDir = Join-Path $ReportDir "assets"
$JsDir     = Join-Path $AssetsDir "js"
$CssDir    = Join-Path $AssetsDir "css"

# Utwórz katalogi
foreach ($Dir in @($JsDir, $CssDir, (Join-Path $CssDir "jstree\default"))) {
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
}

Write-Host "Pobieranie bibliotek front-end..." -ForegroundColor Cyan

$Downloads = @(
    # jQuery
    @{ Url = "https://code.jquery.com/jquery-3.7.1.min.js"; Dest = "$JsDir\jquery.min.js"; Name = "jQuery 3.7.1" },

    # DataTables - JS + CSS (standalone, bez jQuery-ui)
    @{ Url = "https://cdn.datatables.net/1.13.8/js/jquery.dataTables.min.js"; Dest = "$JsDir\datatables.min.js"; Name = "DataTables 1.13.8 JS" },
    @{ Url = "https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css"; Dest = "$CssDir\datatables.min.css"; Name = "DataTables 1.13.8 CSS" },

    # jsTree - JS + CSS
    @{ Url = "https://cdnjs.cloudflare.com/ajax/libs/jstree/3.3.16/jstree.min.js"; Dest = "$JsDir\jstree.min.js"; Name = "jsTree 3.3.16 JS" },
    @{ Url = "https://cdnjs.cloudflare.com/ajax/libs/jstree/3.3.16/themes/default/style.min.css"; Dest = "$CssDir\jstree\default\style.min.css"; Name = "jsTree 3.3.16 CSS" },
    @{ Url = "https://cdnjs.cloudflare.com/ajax/libs/jstree/3.3.16/themes/default/throbber.gif"; Dest = "$CssDir\jstree\default\throbber.gif"; Name = "jsTree throbber" },
    @{ Url = "https://cdnjs.cloudflare.com/ajax/libs/jstree/3.3.16/themes/default/32px.png"; Dest = "$CssDir\jstree\default\32px.png"; Name = "jsTree icons" },
    @{ Url = "https://cdnjs.cloudflare.com/ajax/libs/jstree/3.3.16/themes/default/40px.png"; Dest = "$CssDir\jstree\default\40px.png"; Name = "jsTree icons 2" },

    # Chart.js
    @{ Url = "https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"; Dest = "$JsDir\chart.umd.min.js"; Name = "Chart.js 4.4.2" }
)

$TotalSuccess = 0
$TotalFailed = 0

foreach ($Item in $Downloads) {
    Write-Host "  Pobieranie: $($Item.Name)..." -NoNewline

    # Pomiń jeśli plik już istnieje
    if (Test-Path $Item.Dest) {
        Write-Host " [ISTNIEJE - pomijanie]" -ForegroundColor Gray
        $TotalSuccess++
        continue
    }

    try {
        $WebClient = New-Object System.Net.WebClient
        $WebClient.Headers.Add("User-Agent", "SharePoint-Permission-Analyzer-Installer/1.0")
        $WebClient.DownloadFile($Item.Url, $Item.Dest)
        $FileSize = [System.IO.FileInfo]::new($Item.Dest).Length
        Write-Host " [OK] ($([Math]::Round($FileSize / 1KB, 1)) KB)" -ForegroundColor Green
        $TotalSuccess++
        $WebClient.Dispose()
    }
    catch {
        Write-Host " [BLAD] $_" -ForegroundColor Red
        $TotalFailed++
    }
}

Write-Host ""
Write-Host "Wynik instalacji:" -ForegroundColor Cyan
Write-Host "  Pobrane: $TotalSuccess" -ForegroundColor Green
Write-Host "  Bledy  : $TotalFailed" -ForegroundColor $(if ($TotalFailed -gt 0) { "Red" } else { "Green" })

if ($TotalFailed -gt 0) {
    Write-Host ""
    Write-Host "UWAGA: Niektorych bibliotek nie udalo sie pobrac." -ForegroundColor Yellow
    Write-Host "Pobierz je recznie i umies w katalogu assets/js/ lub assets/css/:" -ForegroundColor Yellow
    Write-Host "  - jQuery: https://jquery.com/download/" -ForegroundColor Gray
    Write-Host "  - DataTables: https://datatables.net/download/" -ForegroundColor Gray
    Write-Host "  - jsTree: https://github.com/vakata/jstree/releases" -ForegroundColor Gray
    Write-Host "  - Chart.js: https://github.com/chartjs/Chart.js/releases" -ForegroundColor Gray
}
else {
    Write-Host ""
    Write-Host "Wszystkie biblioteki zostaly pobrane pomyslnie!" -ForegroundColor Green
    Write-Host "Folder Report/ jest gotowy do uzycia offline." -ForegroundColor Green
}

# Weryfikacja - sprawdź czy wszystkie pliki istnieją
Write-Host ""
Write-Host "Weryfikacja plikow:" -ForegroundColor Cyan
$RequiredFiles = @(
    "$JsDir\jquery.min.js",
    "$JsDir\datatables.min.js",
    "$JsDir\jstree.min.js",
    "$JsDir\chart.umd.min.js",
    "$CssDir\datatables.min.css",
    "$CssDir\jstree\default\style.min.css"
)

foreach ($File in $RequiredFiles) {
    $Exists = Test-Path $File
    $Status = if ($Exists) { "[OK]" } else { "[BRAK]" }
    $Color  = if ($Exists) { "Green" } else { "Red" }
    Write-Host "  $Status $File" -ForegroundColor $Color
}

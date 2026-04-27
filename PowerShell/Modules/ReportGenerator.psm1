#Requires -Version 5.1
<#
.SYNOPSIS
    Modul generowania raportu HTML z danych skanu uprawnien SharePoint.

.DESCRIPTION
    Kopiuje szablon HTML do katalogu wyjsciowego,
    generuje plik data.js z danymi JSON skanu,
    konfiguruje metadane raportu.
#>

Set-StrictMode -Version Latest

function New-HTMLPermissionReport {
    <#
    .SYNOPSIS
        Generuje kompletny raport HTML z danych skanu.

    .PARAMETER ScanResult
        Obiekt wynikowy z Invoke-SPFarmScan.

    .PARAMETER TemplatePath
        Sciezka do folderu z szablonem HTML (Report/).

    .PARAMETER OutputDirectory
        Katalog wyjsciowy dla raportu.

    .PARAMETER ReportTitle
        Tytul raportu widoczny w interfejsie.

    .PARAMETER JsonDataPath
        Sciezka do wygenerowanego pliku JSON (dla odnosnika w raporcie).

    .PARAMETER DiffData
        Dane roznic z poprzednim raportem (z Compare-PermissionScanResults). Null = pierwszy raport.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ScanResult,

        [Parameter(Mandatory = $true)]
        [string]$TemplatePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $false)]
        [string]$ReportTitle = "SharePoint Permission Report",

        [Parameter(Mandatory = $false)]
        [string]$JsonDataPath = "",

        [Parameter(Mandatory = $false)]
        [object]$DiffData = $null,

        [Parameter(Mandatory = $false)]
        [object]$ScanHistory = $null
    )

    Write-ScanLog -Level "Info" -Message "Generowanie raportu HTML..."
    Write-ScanLog -Level "Info" -Message "  Szablon: $TemplatePath"
    Write-ScanLog -Level "Info" -Message "  Wyjscie: $OutputDirectory"

    # Utwórz katalog wyjściowy
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    # Sprawdź szablon
    if (-not (Test-Path $TemplatePath)) {
        throw "Nie znaleziono folderu szablonu: $TemplatePath"
    }

    # Skopiuj pliki szablonu
    Copy-TemplateFiles -TemplatePath $TemplatePath -OutputDirectory $OutputDirectory

    # Generuj data.js
    $DataJsPath = Join-Path $OutputDirectory "data.js"
    $CacheBuster = Get-Date -Format 'yyyyMMddHHmmss'
    New-ReportDataJs -ScanResult $ScanResult -OutputPath $DataJsPath -ReportTitle $ReportTitle -DiffData $DiffData -ScanHistory $ScanHistory

    # Wstaw cache-busting do index.html dla data.js, report.js i report.css.
    # Wymusza pobranie swiezych plikow przez przegladarke po kazdym skanie.
    # Bez tego przegladarka moze serwowac stary report.js z cache, co powoduje
    # niezgodnosc z nowym index.html i brak danych na zakladce Dashboard.
    $IndexPath = Join-Path $OutputDirectory "index.html"
    if (Test-Path $IndexPath) {
        $IndexContent = Get-Content -Path $IndexPath -Raw -Encoding UTF8
        $IndexContent = $IndexContent -replace 'src="data\.js(?:\?v=[^"]*)?\"',         "src=`"data.js?v=$CacheBuster`""
        $IndexContent = $IndexContent -replace 'src="assets/js/report\.js(?:\?v=[^"]*)?\"', "src=`"assets/js/report.js?v=$CacheBuster`""
        $IndexContent = $IndexContent -replace 'href="assets/css/report\.css(?:\?v=[^"]*)?\"', "href=`"assets/css/report.css?v=$CacheBuster`""
        [System.IO.File]::WriteAllText($IndexPath, $IndexContent, [System.Text.Encoding]::UTF8)
    }
    else {
        Write-Warning "Brak index.html w katalogu wyjsciowym po kopiowaniu szablonu"
    }

    Write-ScanLog -Level "Info" -Message "Raport HTML wygenerowany: $OutputDirectory"
    return $OutputDirectory
}

function Copy-TemplateFiles {
    <#
    .SYNOPSIS
        Kopiuje pliki szablonu HTML do katalogu wyjsciowego.
    #>
    [CmdletBinding()]
    param(
        [string]$TemplatePath,
        [string]$OutputDirectory
    )

    try {
        # Kopiuj index.html
        $TemplateIndex = Join-Path $TemplatePath "index.html"
        if (Test-Path $TemplateIndex) {
            Copy-Item -Path $TemplateIndex -Destination $OutputDirectory -Force
        }
        else {
            Write-Warning "Brak index.html w szablonie. Raport moze nie dzialac poprawnie."
        }

        # Kopiuj assets/
        $TemplateAssets = Join-Path $TemplatePath "assets"
        if (Test-Path $TemplateAssets) {
            $AssetsOutput = Join-Path $OutputDirectory "assets"
            if (-not (Test-Path $AssetsOutput)) {
                New-Item -ItemType Directory -Path $AssetsOutput -Force | Out-Null
            }
            Copy-Item -Path "$TemplateAssets\*" -Destination $AssetsOutput -Recurse -Force
        }
        else {
            Write-Warning "Brak katalogu assets/ w szablonie."
        }
    }
    catch {
        throw "Blad kopiowania szablonu: $_"
    }
}

function New-ReportDataJs {
    <#
    .SYNOPSIS
        Generuje plik data.js z danymi skanu jako JavaScript.
        Plik jest ladowany przez index.html przez <script src="data.js">.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ScanResult,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [string]$ReportTitle = "SharePoint Permission Report",

        [Parameter(Mandatory = $false)]
        [object]$DiffData = $null,

        [Parameter(Mandatory = $false)]
        [object]$ScanHistory = $null
    )

    Write-ScanLog -Level "Info" -Message "Generowanie data.js..."

    try {
        # Konwertuj dane skanu do JSON
        $JsonData = $ScanResult | ConvertTo-Json -Depth 12 -Compress
        $JsonBytes = [System.Text.Encoding]::UTF8.GetBytes($JsonData)
        $JsonBase64 = [Convert]::ToBase64String($JsonBytes)

        # Przygotuj dane wersjonowania (roznice)
        $CurrentGenerated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        if ($DiffData -ne $null) {
            # Dodaj timestamp aktualnego raportu do danych diff
            $DiffData['CurrentGenerated'] = $CurrentGenerated
            $DiffJson = $DiffData | ConvertTo-Json -Depth 8 -Compress
        }
        else {
            # Pierwszy raport - brak poprzedniej wersji
            $DiffJson = "{`"IsFirstReport`":true,`"PreviousGenerated`":null,`"CurrentGenerated`":`"$CurrentGenerated`",`"AddedObjects`":[],`"RemovedObjects`":[],`"ChangedObjects`":[]}"
        }
        $DiffBytes = [System.Text.Encoding]::UTF8.GetBytes($DiffJson)
        $DiffBase64 = [Convert]::ToBase64String($DiffBytes)

        $DataJsContent = @"
// ============================================================
// SharePoint Permission Analyzer - Data File
// Wygenerowany: $CurrentGenerated
// NIE EDYTUJ TEGO PLIKU RECZNIE
// ============================================================

(function() {
    'use strict';

    // Dane zakodowane w Base64 dla bezpieczenstwa (unikanie problemow z escapowaniem)
    var _b64 = "$JsonBase64";

    try {
        var _json = decodeURIComponent(
            Array.prototype.map.call(
                atob(_b64),
                function(c) {
                    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
                }
            ).join('')
        );
        window.SCAN_DATA = JSON.parse(_json);
    } catch(e) {
        console.error('Blad ladowania danych raportu:', e);
        window.SCAN_DATA = null;
        window.SCAN_DATA_ERROR = e.message;
    }

    window.REPORT_TITLE = $(ConvertTo-Json $ReportTitle);
    window.REPORT_GENERATED = "$CurrentGenerated";
    window.REPORT_SERVER = "$env:COMPUTERNAME";

    // Dane wersjonowania (roznice z poprzednim raportem)
    var _diff_b64 = "$DiffBase64";
    try {
        var _diff_json = decodeURIComponent(
            Array.prototype.map.call(
                atob(_diff_b64),
                function(c) {
                    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
                }
            ).join('')
        );
        window.DIFF_DATA = JSON.parse(_diff_json);
    } catch(e) {
        console.error('Blad ladowania danych wersjonowania:', e);
        window.DIFF_DATA = null;
    }

})();
"@

        # Dane historii skanowania (opcjonalne)
        $HistorySectionJs = ''
        if ($null -ne $ScanHistory -and @($ScanHistory).Count -gt 0) {
            $HistJson = @($ScanHistory) | ConvertTo-Json -Depth 5 -Compress
            $HistBytes = [System.Text.Encoding]::UTF8.GetBytes($HistJson)
            $HistBase64 = [System.Convert]::ToBase64String($HistBytes)
            $HistCount = @($ScanHistory).Count
            $HistorySectionJs = (
                "`n// Dane historii skanowania (ostatnie $HistCount raportow)" +
                "`n(function() {" +
                "`n    var _hist_b64 = `"" + $HistBase64 + "`";" +
                "`n    try {" +
                "`n        var _hist_json = decodeURIComponent(Array.prototype.map.call(atob(_hist_b64), function(c) { return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2); }).join(''));" +
                "`n        window.SCAN_HISTORY = JSON.parse(_hist_json);" +
                "`n    } catch(e) {" +
                "`n        console.error('Blad ladowania historii:', e);" +
                "`n        window.SCAN_HISTORY = [];" +
                "`n    }" +
                "`n})();"
            )
        }

        ($DataJsContent + $HistorySectionJs) | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
        Write-ScanLog -Level "Info" -Message "data.js wygenerowany: $OutputPath ($([int](([System.IO.FileInfo]$OutputPath).Length / 1KB)) KB)"
    }
    catch {
        throw "Blad generowania data.js: $_"
    }
}

function Compare-PermissionScanResults {
    <#
    .SYNOPSIS
        Porownuje dwa wyniki skanu i zwraca roznice uprawnien.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Previous,

        [Parameter(Mandatory = $true)]
        [object]$Current,

        [Parameter(Mandatory = $false)]
        [string]$PreviousGenerated = ""
    )

    # Zbuduj mapy ObjectId -> obiekt
    $PrevMap = @{}
    if ($Previous.Objects) {
        foreach ($obj in $Previous.Objects) {
            if ($obj.ObjectId) { $PrevMap[$obj.ObjectId] = $obj }
        }
    }

    $CurrMap = @{}
    foreach ($obj in $Current.Objects) {
        if ($obj.ObjectId) { $CurrMap[$obj.ObjectId] = $obj }
    }

    $AddedObjects   = [System.Collections.Generic.List[hashtable]]::new()
    $RemovedObjects = [System.Collections.Generic.List[hashtable]]::new()
    $ChangedObjects = [System.Collections.Generic.List[hashtable]]::new()

    # Nowe obiekty (w Current ale nie w Previous)
    foreach ($objId in $CurrMap.Keys) {
        if (-not $PrevMap.ContainsKey($objId)) {
            $obj = $CurrMap[$objId]
            $AddedObjects.Add(@{
                ObjectId             = $obj.ObjectId
                ObjectType           = [string]($obj.ObjectType)
                Title                = [string]($obj.Title)
                ServerRelativeUrl    = [string]($obj.ServerRelativeUrl)
                HasUniquePermissions = [bool]($obj.HasUniquePermissions)
            })
        }
    }

    # Usuniete obiekty (w Previous ale nie w Current)
    foreach ($objId in $PrevMap.Keys) {
        if (-not $CurrMap.ContainsKey($objId)) {
            $obj = $PrevMap[$objId]
            $RemovedObjects.Add(@{
                ObjectId             = $obj.ObjectId
                ObjectType           = [string]($obj.ObjectType)
                Title                = [string]($obj.Title)
                ServerRelativeUrl    = [string]($obj.ServerRelativeUrl)
                HasUniquePermissions = [bool]($obj.HasUniquePermissions)
            })
        }
    }

    # Zmienione obiekty (w obu - porownaj assignments)
    foreach ($objId in $CurrMap.Keys) {
        if (-not $PrevMap.ContainsKey($objId)) { continue }

        $currObj = $CurrMap[$objId]
        $prevObj = $PrevMap[$objId]

        # Zbuduj mapy przypisań: klucz = LoginName + '|' + PermissionLevels (posortowane)
        # @() wymusza tablicę - PS5.1 ConvertFrom-Json może zwrócić scalar dla single-element array
        $prevA = @{}
        foreach ($a in @($prevObj.Assignments | Where-Object { $_ })) {
            if ($a.LoginName) {
                $key = [string]($a.LoginName) + '|' + (($a.PermissionLevels | Where-Object { $_ } | Sort-Object) -join ',')
                $prevA[$key] = $a
            }
        }

        $currA = @{}
        foreach ($a in @($currObj.Assignments | Where-Object { $_ })) {
            if ($a.LoginName) {
                $key = [string]($a.LoginName) + '|' + (($a.PermissionLevels | Where-Object { $_ } | Sort-Object) -join ',')
                $currA[$key] = $a
            }
        }

        $AddedAssignments   = [System.Collections.Generic.List[hashtable]]::new()
        $RemovedAssignments = [System.Collections.Generic.List[hashtable]]::new()

        foreach ($key in $currA.Keys) {
            if (-not $prevA.ContainsKey($key)) {
                $a = $currA[$key]
                $AddedAssignments.Add(@{
                    LoginName      = [string]($a.LoginName)
                    DisplayName    = [string]($a.DisplayName)
                    PrincipalType  = [string]($a.PrincipalType)
                    PermissionLevels = @($a.PermissionLevels | Where-Object { $_ })
                    SourceType     = [string]($a.SourceType)
                    SourceName     = [string]($a.SourceName)
                })
            }
        }

        foreach ($key in $prevA.Keys) {
            if (-not $currA.ContainsKey($key)) {
                $a = $prevA[$key]
                $RemovedAssignments.Add(@{
                    LoginName      = [string]($a.LoginName)
                    DisplayName    = [string]($a.DisplayName)
                    PrincipalType  = [string]($a.PrincipalType)
                    PermissionLevels = @($a.PermissionLevels | Where-Object { $_ })
                    SourceType     = [string]($a.SourceType)
                    SourceName     = [string]($a.SourceName)
                })
            }
        }

        $uniqueChanged = ([bool]($prevObj.HasUniquePermissions) -ne [bool]($currObj.HasUniquePermissions))

        if ($AddedAssignments.Count -gt 0 -or $RemovedAssignments.Count -gt 0 -or $uniqueChanged) {
            $ChangedObjects.Add(@{
                ObjectId                = $currObj.ObjectId
                ObjectType              = [string]($currObj.ObjectType)
                Title                   = [string]($currObj.Title)
                ServerRelativeUrl       = [string]($currObj.ServerRelativeUrl)
                WebApplicationUrl       = [string]($currObj.WebApplicationUrl)
                SiteCollectionUrl       = [string]($currObj.SiteCollectionUrl)
                UniquePermissionsChanged = $uniqueChanged
                OldHasUnique            = [bool]($prevObj.HasUniquePermissions)
                NewHasUnique            = [bool]($currObj.HasUniquePermissions)
                AddedAssignments        = $AddedAssignments.ToArray()
                RemovedAssignments      = $RemovedAssignments.ToArray()
            })
        }
    }

    return @{
        IsFirstReport      = $false
        PreviousGenerated  = $PreviousGenerated
        CurrentGenerated   = ""
        AddedObjects       = $AddedObjects.ToArray()
        RemovedObjects     = $RemovedObjects.ToArray()
        ChangedObjects     = $ChangedObjects.ToArray()
    }
}

Export-ModuleMember -Function @(
    "New-HTMLPermissionReport",
    "Copy-TemplateFiles",
    "New-ReportDataJs",
    "Compare-PermissionScanResults"
)

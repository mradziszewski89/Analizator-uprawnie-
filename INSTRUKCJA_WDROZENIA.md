# SharePoint Permission Analyzer v1.1 - Instrukcja wdrożenia

> **Wersja dokumentu:** 1.1 — 2026-04-21  
> **Dotyczy:** SharePoint Subscription Edition (SE) on-premises

## SPIS TREŚCI

1. [Architektura rozwiązania](#1-architektura-rozwiązania)
2. [Struktura katalogów](#2-struktura-katalogów)
3. [Wymagania systemowe](#3-wymagania-systemowe)
4. [Instalacja — krok po kroku](#4-instalacja--krok-po-kroku)
5. [Konfiguracja skanowania (ScanConfig.json)](#5-konfiguracja-skanowania-scanconfigjson)
6. [Uruchomienie skanu](#6-uruchomienie-skanu)
7. [Publikacja raportu na SharePoint](#7-publikacja-raportu-na-sharepoint)
8. [Konfiguracja SharePoint dla raportu HTML](#8-konfiguracja-sharepoint-dla-raportu-html)
9. [Opis raportu HTML — zakładki i funkcje](#9-opis-raportu-html--zakładki-i-funkcje)
10. [Remediacja — tryb lokalny (generowanie PS1)](#10-remediacja--tryb-lokalny-generowanie-ps1)
11. [Remediacja — backend SharePoint (Farm Solution)](#11-remediacja--backend-sharepoint-farm-solution)
12. [Archiwizacja i historia raportów](#12-archiwizacja-i-historia-raportów)
13. [Opis pól JSON](#13-opis-pól-json)
14. [Znane ograniczenia i obejścia](#14-znane-ograniczenia-i-obejścia)
15. [Recovery i rollback](#15-recovery-i-rollback)
16. [Checklist testów akceptacyjnych](#16-checklist-testów-akceptacyjnych)

---

## 1. ARCHITEKTURA ROZWIĄZANIA

```
┌─────────────────────────────────────────────────────────────────────┐
│                  WARSTWA SKANOWANIA (PowerShell 5.1)                │
│                                                                     │
│  Start-PermissionScan.ps1                                           │
│    ├── SPPermissionScanner.psm1   (SSOM — skanowanie ACL)           │
│    ├── ADGroupExpander.psm1       (LDAP — ekspansja grup AD)        │
│    ├── ReportGenerator.psm1       (generowanie HTML + data.js)      │
│    └── RemediationScriptGenerator.psm1 (generowanie PS1)            │
│                                                                     │
│  DANE WYJŚCIOWE:                                                    │
│    ├── PermissionReport_YYYY-MM-DD_HH-mm-ss.json   ← dane surowe   │
│    ├── PermissionReport_YYYY-MM-DD_HH-mm-ss.csv    ← eksport CSV   │
│    ├── Output/Report/                               ← bieżący HTML  │
│    │    ├── index.html                                              │
│    │    ├── data.js  (SCAN_DATA + DIFF_DATA + SCAN_HISTORY)         │
│    │    └── assets/  (lokalne biblioteki JS/CSS)                    │
│    └── Output/Report_YYYY-MM-DD_HH-mm-ss/          ← archiwa       │
│         └── (kopia każdego poprzedniego raportu)                    │
└─────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────┐
│                  WARSTWA RAPORTU (Przeglądarka)                     │
│                                                                     │
│  index.html — statyczny, interaktywny, działa offline               │
│    ├── report.js       (logika, filtry, tabela, drzewo, historia)   │
│    ├── report.css      (styl administracyjny, tryb ciemny)          │
│    ├── jquery.min.js                                                │
│    ├── datatables.min.js                                            │
│    ├── jstree.min.js                                                │
│    └── chart.umd.min.js   (wykresy dashboard + historia)            │
│                                                                     │
│  Źródła danych (data.js):                                           │
│    window.SCAN_DATA     — bieżące obiekty i uprawnienia (Base64)   │
│    window.DIFF_DATA     — różnice vs poprzedni skan (Base64)        │
│    window.SCAN_HISTORY  — historia ostatnich 7 skanów (Base64)      │
└─────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────┐
│                 WARSTWA REMEDIACJI (dual mode)                      │
│                                                                     │
│  TRYB LOKALNY (plik HTML / localhost):                              │
│    Przeglądarka → generuje PS1 → administrator uruchamia            │
│    Po wykonaniu PS1: przycisk "Odśwież dane raportu"                │
│    → pobiera nowy data.js przez <script src> (CSP-safe)             │
│                                                                     │
│  TRYB SHAREPOINT-HOSTED (Farm Solution):                            │
│    Przeglądarka → AJAX → RemediationApiHandler.ashx.cs              │
│    (kontrola uprawnień Farm Admin / grupa AD)                       │
│    → SPWeb.RoleAssignments.Remove() / ResetRoleInheritance()        │
│    → Zapis do AuditLog.csv                                          │
│    → Auto-reload data.js po sukcesie (CSP-safe)                     │
└─────────────────────────────────────────────────────────────────────┘
```

### Dane w pliku data.js

`data.js` zawiera trzy sekcje zakodowane jako Base64 (UTF-8 → atob → JSON.parse):

| Zmienna globalna | Zawartość | Generowana przez |
|---|---|---|
| `window.SCAN_DATA` | Pełny wynik bieżącego skanu (obiekty + przypisania) | `New-ReportDataJs` |
| `window.DIFF_DATA` | Zestawienie zmian vs poprzedni skan (+/-/~) | `New-ReportDataJs` |
| `window.SCAN_HISTORY` | Ostatnie 7 archiwów (statystyki + rozmiar dysku) | `New-ReportDataJs` |

---

## 2. STRUKTURA KATALOGÓW

```
SharePoint-Permission-Analyzer/
│
├── PowerShell/
│   ├── Start-PermissionScan.ps1          ← GŁÓWNY SKRYPT
│   ├── Config/
│   │   ├── ScanConfig.json               ← Konfiguracja skanowania
│   │   ├── Exclusions.json               ← Listy systemowe do wykluczenia
│   │   └── Whitelist.json                ← Konta chronione przed remediacja
│   ├── Modules/
│   │   ├── SPPermissionScanner.psm1      ← Skaner SSOM
│   │   ├── ADGroupExpander.psm1          ← Ekspansja grup AD (LDAP)
│   │   ├── ReportGenerator.psm1          ← Generator raportu HTML + data.js
│   │   └── RemediationScriptGenerator.psm1 ← Generator skryptów PS1
│   ├── Output/
│   │   ├── PermissionReport_*.json       ← Dane surowe skanów
│   │   ├── PermissionReport_*.csv        ← Eksport CSV skanów
│   │   ├── Report/                       ← BIEŻĄCY raport HTML (nadpisywany)
│   │   │   ├── index.html
│   │   │   ├── data.js
│   │   │   └── assets/
│   │   └── Report_YYYY-MM-DD_HH-mm-ss/  ← ARCHIWA poprzednich raportów
│   │       ├── index.html
│   │       ├── data.js
│   │       └── assets/
│   └── Logs/
│       └── Scan_YYYY-MM-DD_HH-mm-ss.log ← Log każdego skanu
│
├── Report/                               ← SZABLON (źródło kopiowania)
│   ├── index.html
│   ├── Install-LocalLibraries.ps1        ← Pobierz biblioteki JS/CSS
│   └── assets/
│       ├── js/
│       │   ├── report.js                 ← Logika raportu (dashboard/tabela/drzewo/historia/remediacja)
│       │   ├── jquery.min.js
│       │   ├── datatables.min.js
│       │   ├── jstree.min.js
│       │   └── chart.umd.min.js
│       └── css/
│           ├── report.css
│           ├── datatables.min.css
│           └── jstree/default/
│
├── FarmSolution/
│   └── SharePointPermissionAnalyzer/
│       ├── SharePointPermissionAnalyzer.csproj
│       ├── Properties/AssemblyInfo.cs
│       ├── Features/PermissionAnalyzerFeature/
│       │   └── PermissionAnalyzerFeature.feature
│       └── Layouts/PermissionAnalyzer/
│           ├── PermissionRemediationPage.aspx
│           ├── PermissionRemediationPage.aspx.cs
│           ├── RemediationApiHandler.ashx
│           └── RemediationApiHandler.ashx.cs
│
├── Samples/
│   ├── SampleScanResult.json
│   └── SampleScanLog.log
│
├── INSTRUKCJA_WDROZENIA.md               ← TEN PLIK
└── DOKUMENTACJA_TECHNICZNA.html          ← Dokumentacja techniczna modułów
```

---

## 3. WYMAGANIA SYSTEMOWE

### Serwer SharePoint

| Wymaganie | Minimalne | Zalecane |
|---|---|---|
| Platforma | SharePoint SE on-premises | SharePoint SE + SP2022 |
| PowerShell | 5.1 | 5.1 (wbudowany w WS 2019/2022) |
| Snap-in | Microsoft.SharePoint.PowerShell | Załadowany przez SP Management Shell |
| .NET Framework | 4.8 | 4.8 |
| Konto | Farm Administrator | Farm Administrator z dostępem do wszystkich SC |

> **Uwaga:** Skrypt jest w pełni kompatybilny z uruchamianiem zarówno z **SharePoint Management Shell** (snap-in załadowany automatycznie), jak i ze zwykłego **PowerShell 5.1** (snap-in ładowany dynamicznie przez skrypt). Jeśli snap-in jest niedostępny przez inny mechanizm, skrypt loguje informację i kontynuuje.

### Stacja administratora (opcjonalnie)

- Dostęp do Internetu (jednorazowo, dla `Install-LocalLibraries.ps1`)
- PowerShell 5.1

### Przeglądarka dla raportu HTML

- Chrome 90+, Firefox 88+, Edge 88+
- JavaScript włączony
- Brak wymogu połączenia internetowego po wygenerowaniu raportu (wszystkie biblioteki lokalne)

---

## 4. INSTALACJA — KROK PO KROKU

### KROK 1: Skopiuj pliki na serwer SharePoint

```powershell
$AnalyzerPath = "E:\AdminData\Analizator uprawnień"
New-Item -ItemType Directory -Path $AnalyzerPath -Force

Copy-Item -Path "\\INSTALL-SERVER\SP-Analyzer\PowerShell" -Destination $AnalyzerPath -Recurse
Copy-Item -Path "\\INSTALL-SERVER\SP-Analyzer\Report"      -Destination $AnalyzerPath -Recurse
```

### KROK 2: Pobierz lokalne biblioteki JS/CSS (wymaga Internetu — jednorazowo)

```powershell
Set-Location "$AnalyzerPath\Report"
.\Install-LocalLibraries.ps1
```

Jeśli serwer **nie ma dostępu do Internetu**:
1. Uruchom `Install-LocalLibraries.ps1` na komputerze z dostępem do Internetu.
2. Skopiuj wynikowy folder `Report/assets/` na serwer SharePoint.

### KROK 3: Dostosuj konfigurację

```powershell
notepad "$AnalyzerPath\PowerShell\Config\ScanConfig.json"
```

Kluczowe parametry do weryfikacji:

| Parametr | Domyślnie | Opis |
|---|---|---|
| `FarmSettings.WebApplicationUrls` | `[]` (wszystkie) | Ogranicz zakres do wybranych WebApp |
| `PrincipalExpansion.ExpandDomainGroups` | `true` | Ekspansja grup AD przez LDAP |
| `ScanDepth.ScanFiles` | `true` | Skanuj elementy list i pliki |
| `Performance.ThrottleDelayMs` | `10` | Opóźnienie ms między obiektami (10-50 dla dużych farm) |
| `Output.SharePointPublishUrl` | `""` | URL biblioteki do auto-publikacji (lub użyj parametru CLI) |

### KROK 4: Dostosuj Whitelist.json (WAŻNE!)

Dodaj do `Whitelist.json` własne konta serwisowe, które nie powinny być obiektem remediacji:

```json
{
  "ProtectedAccounts": {
    "Users": [
      {
        "LoginName": "TEST\\sp-farm-admin",
        "DisplayName": "Farm Administrator",
        "Reason": "Konto administracyjne farmy"
      },
      {
        "LoginName": "TEST\\sp-service-account",
        "Reason": "Konto serwisowe SharePoint"
      }
    ]
  }
}
```

### KROK 5: Weryfikacja snap-in

```powershell
Get-PSSnapin -Registered | Where-Object { $_.Name -like "*SharePoint*" }
# Powinno zwrócić: Microsoft.SharePoint.PowerShell
```

---

## 5. KONFIGURACJA SKANOWANIA (ScanConfig.json)

### Szybki skan — tylko metadane uprawnień bez plików/elementów

```json
{
  "ScanDepth": {
    "ScanFiles": false,
    "ScanListItems": false
  }
}
```

### Skan z ekspansją grup AD

```json
{
  "PrincipalExpansion": {
    "RawAssignmentsOnly": false,
    "ExpandSharePointGroups": true,
    "ExpandDomainGroups": true,
    "MaxGroupNestingDepth": 10
  }
}
```

### Skan tylko wybranych WebApp/SC

```json
{
  "FarmSettings": {
    "WebApplicationUrls": ["https://portalse.test.pl"],
    "SiteCollectionUrls": ["https://portalse.test.pl/sites/hr"]
  }
}
```

### Konfiguracja auto-publikacji na SharePoint

```json
{
  "Output": {
    "SharePointPublishUrl": "https://portalse.test.pl/RaportySP"
  }
}
```

Alternatywnie podaj URL przez parametr CLI (patrz sekcja 6).

---

## 6. URUCHOMIENIE SKANU

Otwórz **SharePoint Management Shell** (lub PowerShell 5.1) jako **Farm Administrator**.

### Pełny skan farmy

```powershell
cd "E:\AdminData\Analizator uprawnień\PowerShell"
.\Start-PermissionScan.ps1
```

### Skan z automatyczną publikacją raportu na SharePoint

```powershell
.\Start-PermissionScan.ps1 -SharePointLibraryUrl "https://portalse.test.pl/RaportySP"
```

Parametr `-SharePointLibraryUrl` nadpisuje wartość `SharePointPublishUrl` z `ScanConfig.json`.  
Gdy **żaden** z tych źródeł nie jest skonfigurowany, skrypt wyświetla komunikat informacyjny i **nie pyta interaktywnie**.

### Pozostałe opcje

```powershell
# Wznowienie z checkpointu (po przerwaniu)
.\Start-PermissionScan.ps1 -ResumeFromCheckpoint

# Skan z podaniem niestandardowego katalogu wyjściowego
.\Start-PermissionScan.ps1 -OutputPath "D:\Raporty\SP-Permissions"

# Symulacja (bez zapisu wyników)
.\Start-PermissionScan.ps1 -DryRun

# Szczegółowe logowanie
.\Start-PermissionScan.ps1 -Verbose
```

### Planowanie cyklicznego skanu (Task Scheduler)

```powershell
$Action = New-ScheduledTaskAction `
    -Execute "PowerShell.exe" `
    -Argument @"
-NonInteractive -File "E:\AdminData\Analizator uprawnień\PowerShell\Start-PermissionScan.ps1" -SharePointLibraryUrl "https://portalse.test.pl/RaportySP"
"@

$Trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At "02:00AM"
$Settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -WakeToRun

Register-ScheduledTask `
    -TaskName "SharePoint Permission Analyzer - Tygodniowy" `
    -Action   $Action  `
    -Trigger  $Trigger `
    -Settings $Settings `
    -RunLevel Highest `
    -User "TEST\sp-farm-admin"
```

---

## 7. PUBLIKACJA RAPORTU NA SHAREPOINT

### Metoda A: Automatyczna — przez parametr CLI (zalecana)

```powershell
.\Start-PermissionScan.ps1 -SharePointLibraryUrl "https://portalse.test.pl/RaportySP"
```

Skrypt:
1. Wgrywa cały folder `Output/Report/` do podanej biblioteki SharePoint.
2. Tworzy podfolder `Report/` w bibliotece (np. `https://portalse.test.pl/RaportySP/Report/`).
3. Automatycznie ustawia wymagane parametry Web Application (patrz sekcja 8).

### Metoda B: Ręczna — WebDAV

```powershell
net use Z: "\\portalse.test.pl@SSL\DavWWWRoot\RaportySP\" /user:TEST\sp-farm-admin
robocopy "E:\AdminData\Analizator uprawnień\PowerShell\Output\Report" "Z:\Report" /E /NP /MIR
net use Z: /delete
```

### Metoda C: Lokalny dostęp z dysku (bez SharePoint)

Otwórz `Output\Report\index.html` bezpośrednio w przeglądarce.

> **Uwaga Chrome/Edge:** Polityki CORS mogą blokować ładowanie `data.js` z dysku lokalnego (`file://`). Uruchom lokalny serwer HTTP:
> ```powershell
> python -m http.server 8080 -d "E:\AdminData\Analizator uprawnień\PowerShell\Output\Report"
> # Otwórz: http://localhost:8080
> ```

---

## 8. KONFIGURACJA SHAREPOINT DLA RAPORTU HTML

Aby plik `index.html` otwierał się **w przeglądarce** (nie był pobierany), w SharePoint 2016+ wymagane są dwa ustawienia na Web Application:

| Ustawienie | Wymagana wartość | Domyślna wartość |
|---|---|---|
| `BrowserFileHandling` | `Permissive` (int: 1) | `Strict` (int: 0) |
| `AllowedInlineDownloadedMimeTypes` | musi zawierać `text/html` | lista bez `text/html` |

> **Ważne dla SP 2016+:** Samo `BrowserFileHandling = Permissive` **nie wystarczy**. Typ `text/html` musi być jawnie dodany do `AllowedInlineDownloadedMimeTypes`. Jest to zmiana bezpieczeństwa wprowadzona w SharePoint 2016.

Skrypt `Start-PermissionScan.ps1` **ustawia oba parametry automatycznie** podczas publikacji. Jeśli z jakiegoś powodu automatyczne ustawienie nie powiedzie się (np. brak uprawnień Farm Admin w kontekście wykonania), użyj poniższego skryptu naprawczego.

### Ręczna naprawa (SharePoint Management Shell jako Farm Admin)

```powershell
$wa = Get-SPWebApplication "https://portalse.test.pl"

# 1. BrowserFileHandling — używamy int zamiast enum (enum może być niedostępny gdy snap-in załadowany przez inny mechanizm)
if ([int]$wa.BrowserFileHandling -ne 1) {
    $wa.BrowserFileHandling = 1   # 1 = Permissive
}

# 2. AllowedInlineDownloadedMimeTypes — wymagane w SP 2016+
if ($wa.AllowedInlineDownloadedMimeTypes -notcontains 'text/html') {
    $wa.AllowedInlineDownloadedMimeTypes.Add('text/html')
}

$wa.Update()

# 3. Wymagany recykl IIS po zmianie
iisreset
```

Lub użyj gotowego skryptu: `C:\Temp\fix_browserfilehandling.ps1 -WebApplicationUrl "https://portalse.test.pl"`

> **Po `iisreset`:** Odśwież stronę w bibliotece SharePoint i kliknij `index.html` — powinien otworzyć się w przeglądarce bez pobierania.

### Dlaczego nie używamy enum `[Microsoft.SharePoint.Administration.SPBrowserFileHandling]`

Gdy snap-in jest załadowany przez mechanizm inny niż bezpośrednie `Add-PSSnapin` (np. przez SharePoint Management Shell lub profil), typ enum może być niedostępny w bieżącej sesji PS. Użycie wartości `int` (`0`/`1`) jest zawsze bezpieczne i niezależne od dostępności typu.

---

## 9. OPIS RAPORTU HTML — ZAKŁADKI I FUNKCJE

### Dashboard

- Statystyki bieżącego skanu: liczba WebApp, SC, witryn, list, plików, przypisań.
- Wykresy kołowe: rozkład typów principalów, rozkład poziomów uprawnień.
- Zestawienie obiektów z unikatowymi ACL.
- Porównanie z poprzednim skanem: +/- nowych i usuniętych obiektów.

### Tabela uprawnień

- Każdy obiekt SharePoint wyświetlany **raz** (jeden wiersz = jeden obiekt).
- Wszystkie przypisania uprawnień dla danego obiektu połączone w jednej komórce (lista z separatorem).
- Filtry: Web Application, typ obiektu, typ principala, poziom uprawnień, flaga "tylko unikatowe ACL", wyszukiwanie po nazwie użytkownika.
- Sortowanie, paginacja, eksport CSV/JSON z aktualnych filtrów.
- Checkbox do zaznaczania obiektów do remediacji.

### Wyszukaj użytkownika/grupę

- Wyszukiwanie po LoginName lub DisplayName — zwraca wszystkie lokalizacje (obiekty SP) z dostępem.

### Drzewo lokalizacji

- Hierarchia: WebApp → SiteCollection → Web → Lista → Plik/Element.
- Kliknięcie węzła otwiera modal z pełnymi szczegółami obiektu i listą przypisań.

### Wersjonowanie (Historia)

- Porównanie z poprzednim skanem: tabela różnic (+nowe, -usunięte, ~zmienione obiekty).
- **Historia ostatnich 7 skanów** (z archiwów `Output/Report_*/`):
  - Tabela: data skanu, liczba obiektów, liczba przypisań, rozmiar na dysku, delta vs poprzedni skan.
  - Wykres mieszany (Chart.js): słupki = rozmiar folderu (KB, lewa oś), linie = obiekty i przypisania (prawa oś).
- Historia jest dostępna automatycznie gdy istnieją foldery `Output/Report_YYYY-MM-DD_HH-mm-ss/`.

### Remediacja

- Panel z zaznaczonymi obiektami z zakładki "Tabela uprawnień".
- Dostępne akcje:
  - Usuń bezpośrednie uprawnienie użytkownika
  - Usuń przypisanie grupy SharePoint z obiektu
  - Usuń przypisanie grupy domenowej
  - Przywróć dziedziczenie (usuwa wszystkie unikatowe ACL obiektu)
- Po wygenerowaniu skryptu PS1: pole powodu operacji (trafia do komentarza i logu audytowego).
- **Auto-reload po remediacji:** Przycisk "Odśwież dane raportu" pobiera nowy `data.js` przez tag `<script src>` (CSP-safe — nie używa `eval()` ani `new Function()`).

---

## 10. REMEDIACJA — TRYB LOKALNY (generowanie PS1)

### Przepływ pracy

1. Otwórz `index.html` w przeglądarce (lokalnie lub przez lokalny serwer HTTP).
2. Przejdź do **Tabela uprawnień** → zaznacz problematyczne obiekty checkboxem.
3. Przejdź do **Remediacja** → wybierz akcję → wpisz powód operacji.
4. Kliknij **"Generuj skrypt PS1"** — przeglądarka pobiera gotowy plik `.ps1`.
5. Przejrzyj skrypt (zawiera sekcję rollback i opis operacji).
6. Uruchom najpierw w trybie dry-run:
   ```powershell
   .\Remediation_Remove_2026-04-21.ps1 -DryRun $true
   ```
7. Po weryfikacji uruchom w trybie live:
   ```powershell
   .\Remediation_Remove_2026-04-21.ps1 -DryRun $false
   ```
8. Wróć do raportu → kliknij **"Odśwież dane raportu"** — raport automatycznie przeładuje `data.js` z nowym skanem.

### Zabezpieczenia wbudowane w generowany PS1

- Konta z `Whitelist.json` są **automatycznie pomijane** (SHAREPOINT\system, NT AUTHORITY\*, konta serwisowe).
- Skrypt tworzy **transcript log** (`.log`) w tym samym katalogu.
- Skrypt eksportuje wyniki operacji do JSON (plik `_results.json`).
- Tryb DRY-RUN nie wykonuje żadnych zmian — tylko symuluje i loguje.

---

## 11. REMEDIACJA — BACKEND SHAREPOINT (Farm Solution)

### Wdrożenie Farm Solution

```powershell
Add-PSSnapin Microsoft.SharePoint.PowerShell

# Skopiuj WSP na serwer SharePoint
$WspPath = "E:\AdminData\FarmSolution\SharePointPermissionAnalyzer.wsp"

Add-SPSolution -LiteralPath $WspPath
Install-SPSolution -Identity "SharePointPermissionAnalyzer.wsp" -GACDeployment -Force

# Sprawdź status
Get-SPSolution -Identity "SharePointPermissionAnalyzer.wsp" | Select-Object Name, Deployed, LastOperationResult

# Adres Application Page
# https://portalse.test.pl/_layouts/15/PermissionAnalyzer/PermissionRemediationPage.aspx
```

### Konfiguracja grupy administracyjnej

W `PermissionRemediationPage.aspx.cs` lub przez `web.config`:

```xml
<appSettings>
  <add key="SPAnalyzerAdminGroup"   value="TEST\SP-Permission-Admins" />
  <add key="SPAnalyzerAuditLogPath" value="E:\Logs\SPAnalyzer\AuditLog.csv" />
</appSettings>
```

### Aktualizacja i usunięcie

```powershell
# Aktualizacja
Update-SPSolution -Identity "SharePointPermissionAnalyzer.wsp" -LiteralPath $NewWspPath -GACDeployment

# Usunięcie
Uninstall-SPSolution -Identity "SharePointPermissionAnalyzer.wsp"
Remove-SPSolution    -Identity "SharePointPermissionAnalyzer.wsp"
```

---

## 12. ARCHIWIZACJA I HISTORIA RAPORTÓW

### Mechanizm archiwizacji

Każde uruchomienie skanu:
1. **Archiwizuje poprzedni raport** z `Output/Report/` do `Output/Report_YYYY-MM-DD_HH-mm-ss/` (gdzie timestamp pochodzi z `window.REPORT_GENERATED` poprzedniego `data.js`).
2. Generuje nowy raport w `Output/Report/`.
3. Odczytuje dane z ostatnich **7 folderów archiwalnych** i osadza je w nowym `data.js` jako `window.SCAN_HISTORY`.

### Dane historyczne w SCAN_HISTORY

Dla każdego archiwum zbierane są:

| Pole | Opis |
|---|---|
| `Generated` | Timestamp skanu (YYYY-MM-DD HH:mm:ss) |
| `FolderName` | Nazwa folderu archiwalnego |
| `FolderSizeBytes` | Rozmiar folderu na dysku |
| `TotalObjectsScanned` | Łączna liczba przeskanowanych obiektów |
| `TotalAssignments` | Łączna liczba przypisań |
| `UniquePermissionsCount` | Liczba obiektów z unikatowymi ACL |
| `WebApplicationCount` | Liczba Web Applications |
| `SiteCollectionCount` | Liczba Site Collections |
| `WebCount` | Liczba witryn (SPWeb) |
| `ListCount` | Liczba list i bibliotek |
| `ItemCount` | Liczba plików i elementów |

### Zarządzanie archiwami

Archiwa nie są automatycznie usuwane. W przypadku braku miejsca na dysku możesz bezpiecznie usunąć starsze foldery `Output/Report_*/` — historia w raporcie po prostu obejmie mniejszą liczbę wersji.

---

## 13. OPIS PÓL JSON

### Obiekt (ScanObject)

| Pole | Typ | Opis |
|---|---|---|
| `ObjectId` | string | Unikalny identyfikator obiektu (GUID lub wygenerowany) |
| `ParentObjectId` | string | ObjectId obiektu nadrzędnego w hierarchii |
| `ObjectType` | enum | `WebApplication`, `SiteCollection`, `Web`, `List`, `Library`, `Folder`, `File`, `ListItem` |
| `WebApplicationUrl` | string | URL Web Application |
| `SiteCollectionUrl` | string | URL Site Collection |
| `WebUrl` | string | URL witryny (SPWeb) |
| `FullUrl` | string | Pełny URL obiektu |
| `ServerRelativeUrl` | string | Względny URL od serwera |
| `Title` | string | Tytuł wyświetlany |
| `Name` | string | Techniczna nazwa (URL-safe) |
| `ListTitle` | string | Tytuł listy/biblioteki (dla elementów) |
| `ListId` | string | GUID listy/biblioteki |
| `ItemId` | int? | ID elementu listy (null dla list/web/SC) |
| `FileLeafRef` | string | Nazwa pliku z rozszerzeniem |
| `IsHidden` | bool | Czy obiekt jest ukryty |
| `IsSystem` | bool | Czy obiekt jest systemowy |
| `HasUniquePermissions` | bool | `true` = unikatowe ACL, `false` = dziedziczone |
| `InheritsFromUrl` | string | URL obiektu, z którego dziedziczy uprawnienia |
| `FirstUniqueAncestorUrl` | string | URL pierwszego przodka z unikatowymi ACL |
| `Assignments` | array | Lista przypisań uprawnień |
| `ScanTimestamp` | ISO 8601 | Czas skanowania tego obiektu |

### Przypisanie (AssignmentRecord)

| Pole | Typ | Opis |
|---|---|---|
| `PrincipalType` | enum | `User`, `SharePointGroup`, `DomainGroup`, `Claim`, `SpecialPrincipal` |
| `LoginName` | string | LoginName principala w SharePoint |
| `DisplayName` | string | Wyświetlana nazwa |
| `Email` | string | Adres email (jeśli dostępny) |
| `SID` | string | Security Identifier |
| `SourceType` | enum | `Direct`, `ViaSharePointGroup`, `ViaDomainGroup`, `Inherited` |
| `SourceName` | string | Nazwa źródła (np. nazwa grupy SP) |
| `PermissionLevels` | array | Lista poziomów uprawnień (np. `["Full Control"]`) |
| `IsLimitedAccessOnly` | bool | `true` = TYLKO Limited Access |
| `IsSiteAdmin` | bool | Czy principal jest administratorem SC |
| `IsActive` | bool | Czy konto jest aktywne (false = wyłączone w AD) |
| `IsOrphaned` | bool | Czy konto nie istnieje w AD |
| `InheritancePath` | array | Ścieżka przez grupy (np. `["SP Owners", "TEST\Managers"]`) |

---

## 14. ZNANE OGRANICZENIA I OBEJŚCIA

### Błąd dostępu do niektórych Site Collections

```
[Error] Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))
```

**Przyczyna:** Konto Farm Admin nie ma dostępu do SC użytkownika osobistego (`/personal/`), SC wewnętrznych (`sitemaster-*`).  
**Obejście:** Jest to oczekiwane zachowanie dla SC systemowych. Błąd jest logowany i skanowanie kontynuuje.

### Ostrzeżenie "NT AUTHORITY\authenticated users"

```
WARNING: Nie znaleziono grupy AD dla: authenticated users
```

**Przyczyna:** `NT AUTHORITY\authenticated users` to principal systemowy, nie grupa AD.  
**Obejście:** Oczekiwane. Jest on traktowany jako `SpecialPrincipal` i nie jest rozwijany przez LDAP.

### Błąd `The property 'Count' cannot be found`

**Przyczyna:** `Set-StrictMode -Version Latest` w PowerShell 5.1 rzuca wyjątek gdy próbujemy wywołać `.Count` na pojedynczym obiekcie (nie tablicy).  
**Rozwiązanie (już zastosowane w kodzie):** Wszystkie wywołania `.Count` opakowujemy w `@()` — np. `@($Assignments).Count`.

### Ostrzeżenie `Unable to find type [Microsoft.SharePoint.Administration.SPBrowserFileHandling]`

**Przyczyna:** Typ enum może być niedostępny gdy snap-in załadowany przez inny mechanizm niż `Add-PSSnapin`.  
**Rozwiązanie (już zastosowane):** Kod używa `[int]` zamiast enum (`1 = Permissive`, `0 = Strict`).

### Plik HTML pobierany zamiast otwierany w przeglądarce (SharePoint 2016+)

**Przyczyna:** Dwa oddzielne ustawienia muszą być skonfigurowane — `BrowserFileHandling = Permissive` **i** `text/html` w `AllowedInlineDownloadedMimeTypes`. Samo pierwsze ustawienie nie wystarczy w SP 2016+.  
**Rozwiązanie:** Patrz sekcja 8. Po zmianie wymagany `iisreset`.

### Raport nie wyświetla danych (Dashboard pusty)

**Przyczyna:** Przeglądarka blokuje ładowanie `data.js` z protokołu `file://` (CORS).  
**Rozwiązanie:** Uruchom lokalny serwer HTTP (`python -m http.server 8080`) lub otwórz raport przez SharePoint.

### CSP blokuje odświeżanie danych po remediacji

**Przyczyna:** Wcześniejsze implementacje używały `eval()` / `new Function()` do ładowania `data.js`, co jest blokowane przez Content Security Policy.  
**Rozwiązanie (już zastosowane):** Reload danych używa dynamicznego tagu `<script src="data.js?v=TIMESTAMP">` — jest CSP-safe.

---

## 15. RECOVERY I ROLLBACK

### Rollback operacji remediacji

Każdy wygenerowany skrypt PS1 zawiera sekcję rollback. Ogólna procedura:

```powershell
Add-PSSnapin Microsoft.SharePoint.PowerShell

# Przywróć uprawnienie użytkownika po usunięciu
$Site = New-Object Microsoft.SharePoint.SPSite("https://portalse.test.pl/sites/hr")
$Web  = $Site.RootWeb
$Web.AllowUnsafeUpdates = $true

$User           = $Web.EnsureUser("TEST\jan.kowalski")
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($User)
$RoleDef        = $Web.RoleDefinitions["Read"]
$RoleAssignment.RoleDefinitionBindings.Add($RoleDef)
$Web.RoleAssignments.Add($RoleAssignment)
$Web.Update()

$Web.AllowUnsafeUpdates = $false
$Web.Dispose()
$Site.Dispose()
```

> **Uwaga:** Operacja "Przywróć dziedziczenie" jest **nieodwracalna automatycznie** — usuwa wszystkie unikatowe ACL obiektu. Przed wykonaniem zawsze sprawdź log "Stan przed operacją" w `AuditLog.csv`.

### Wznowienie przerwanego skanowania

```powershell
# Wznów z ostatniego checkpointu
.\Start-PermissionScan.ps1 -ResumeFromCheckpoint

# Lub uruchom od nowa z konkretną konfiguracją
.\Start-PermissionScan.ps1 -ConfigPath ".\Config\ScanConfig_partial.json"
```

---

## 16. CHECKLIST TESTÓW AKCEPTACYJNYCH

### Testy skanowania

- [ ] **T-01** Skaner wykrywa wszystkie Web Applications (z wyłączeniem CA gdy `IncludeCentralAdministration=false`)
- [ ] **T-02** Skaner wchodzi we wszystkie Site Collections i witryny
- [ ] **T-03** Błąd dostępu do pojedynczej SC nie przerywa całego skanu
- [ ] **T-04** Skaner poprawnie oznacza `HasUniquePermissions=true/false`
- [ ] **T-05** `FirstUniqueAncestorUrl` wskazuje właściwy obiekt w hierarchii
- [ ] **T-06** `@($Assignments).Count` nie rzuca wyjątku pod `Set-StrictMode -Version Latest`
- [ ] **T-07** Checkpoint tworzy się co `BatchSize` obiektów; `-ResumeFromCheckpoint` wznawia poprawnie
- [ ] **T-08** Konta `NT AUTHORITY\*` logują ostrzeżenie, nie błąd krytyczny

### Testy ekspansji principalów

- [ ] **T-10** `ExpandSharePointGroups=true` rozszerza grupy SP do użytkowników (`SourceType=ViaSharePointGroup`)
- [ ] **T-11** `ExpandDomainGroups=true` rozszerza grupy AD rekurencyjnie do `MaxGroupNestingDepth`
- [ ] **T-12** Konta wyłączone w AD: `IsActive=false`; nieistniejące: `IsOrphaned=true`

### Testy generowania raportu

- [ ] **T-20** `data.js` zawiera trzy sekcje: `SCAN_DATA`, `DIFF_DATA`, `SCAN_HISTORY`
- [ ] **T-21** `SCAN_HISTORY` zawiera dane z ostatnich 7 folderów `Report_*/` (lub mniej jeśli archiwów jest mniej)
- [ ] **T-22** Bieżący raport jest archiwizowany przed generowaniem nowego
- [ ] **T-23** Cache-buster `?v=TIMESTAMP` jest dodawany do linków JS/CSS w `index.html`

### Testy raportu HTML

- [ ] **T-30** Dashboard wyświetla poprawne statystyki
- [ ] **T-31** Wykresy dashboard ładują się poprawnie (Chart.js lokalny)
- [ ] **T-32** Tabela uprawnień: każdy obiekt widoczny **raz** z połączonymi przypisaniami
- [ ] **T-33** Filtry (WebApp, typ obiektu, typ principala, poziom uprawnień) działają poprawnie
- [ ] **T-34** Eksport CSV/JSON z filtrów generuje poprawny plik
- [ ] **T-35** Zakładka Wersjonowanie: tabela różnic (poprzedni skan) wyświetlana poprawnie
- [ ] **T-36** Zakładka Wersjonowanie: tabela historii 7 skanów wyświetlana poprawnie
- [ ] **T-37** Wykres historii (Chart.js mixed) renderuje się bez błędów
- [ ] **T-38** Tryb ciemny działa i jest zapamiętywany (localStorage)
- [ ] **T-39** Nazwy z polskimi znakami wyświetlają się poprawnie (UTF-8)
- [ ] **T-40** Brak błędów XSS — nazwy specjalne są escapowane przed wstawieniem do HTML

### Testy publikacji

- [ ] **T-50** `-SharePointLibraryUrl` powoduje upload wszystkich 13 plików raportu
- [ ] **T-51** Brak skonfigurowanego URL wyświetla komunikat informacyjny (nie pyta interaktywnie)
- [ ] **T-52** `BrowserFileHandling = Permissive` ustawiane automatycznie podczas publikacji (używa `int`, nie enum)
- [ ] **T-53** `text/html` dodawane do `AllowedInlineDownloadedMimeTypes` automatycznie
- [ ] **T-54** Po `iisreset`: `index.html` z biblioteki SharePoint otwiera się w przeglądarce (nie pobierany)

### Testy remediacji — tryb lokalny

- [ ] **T-60** Checkbox w tabeli dodaje obiekty do panelu remediacji
- [ ] **T-61** Generowany PS1 zawiera poprawne polecenia SSOM
- [ ] **T-62** PS1 `-DryRun $true` nie wykonuje zmian, tylko loguje
- [ ] **T-63** PS1 pomija konta z Whitelist
- [ ] **T-64** Przycisk "Odśwież dane raportu" ładuje nowy `data.js` przez `<script src>` (bez CSP error)

### Testy bezpieczeństwa backendu (Farm Solution)

- [ ] **T-70** Nieuprawniony użytkownik → HTTP 403
- [ ] **T-71** Próba remediacji konta `SHAREPOINT\system` → blokada
- [ ] **T-72** Próba usunięcia ostatniego Full Control → blokada
- [ ] **T-73** Zewnętrzny URL w `SiteCollectionUrl` → odrzucany (SSRF protection)
- [ ] **T-74** Wszystkie operacje logowane do `AuditLog.csv` (kto, kiedy, co, na jakim obiekcie)
- [ ] **T-75** API wymaga POST (nie GET)

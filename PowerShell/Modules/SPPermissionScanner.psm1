#Requires -Version 5.1
<#
.SYNOPSIS
    Modul glownego skanera uprawnien SharePoint SE on-premises.

.DESCRIPTION
    Zawiera wszystkie funkcje do skanowania uprawnien na poziomach:
    WebApplication, SiteCollection, Web, List/Library, Folder, File, ListItem.
    Obsluguje uprawnienia dziedziczone i unikalne, rozpoznaje typy principalow,
    opcjonalnie rozszerza grupy SharePoint i domenowe.

.NOTES
    Wymaga: Microsoft.SharePoint.PowerShell, PowerShell 5.1
    Uzywa Server-Side Object Model SharePoint.
#>

Set-StrictMode -Version Latest

# ============================================================
# Zmienne globalne modulu
# ============================================================

$script:ScanLogger = $null
$script:ScanConfig = $null
$script:ScanExclusions = $null
$script:ScanWhitelist = $null
$script:ErrorLog = [System.Collections.Generic.List[hashtable]]::new()
$script:SkippedLog = [System.Collections.Generic.List[hashtable]]::new()
$script:CheckpointData = $null
$script:GroupMemberCache = [System.Collections.Generic.Dictionary[string, object[]]]::new()

# Liczniki statystyk
$script:Stats = @{
    WebApplicationCount    = 0
    SiteCollectionCount    = 0
    WebCount               = 0
    ListCount              = 0
    FolderCount            = 0
    ItemCount              = 0
    UniquePermissionsCount = 0
    TotalAssignments       = 0
    TotalObjectsScanned    = 0
    SkippedObjects         = 0
    ErrorCount             = 0
}

# ============================================================
# SEKCJA: Logowanie
# ============================================================

function Initialize-ScanLogger {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Verbose", "Info", "Warning", "Error")]
        [string]$LogLevel = "Info"
    )

    $script:LogFilePath = $LogFilePath
    $script:LogLevel = $LogLevel

    # Utwórz plik logu z nagłówkiem
    $Header = @"
================================================================================
SharePoint Permission Analyzer - Log skanowania
Data:    $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Serwer:  $env:COMPUTERNAME
Uzytkownik: $env:USERDOMAIN\$env:USERNAME
================================================================================
"@
    $Header | Out-File -FilePath $LogFilePath -Encoding UTF8 -Force
}

function Write-ScanLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Verbose", "Info", "Warning", "Error")]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$ObjectUrl = "",

        [Parameter(Mandatory = $false)]
        [object]$Exception = $null
    )

    $LevelPriority = @{ Verbose = 0; Info = 1; Warning = 2; Error = 3 }
    $ConfigPriority = $LevelPriority[$script:LogLevel]
    $MsgPriority = $LevelPriority[$Level]

    if ($MsgPriority -lt $ConfigPriority) { return }

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $ObjectPart = if ($ObjectUrl) { " [$ObjectUrl]" } else { "" }
    $LogLine = "[$Timestamp] [$Level]$ObjectPart $Message"

    if ($Exception) {
        $LogLine += "`n  Exception: $($Exception.Message)"
        if ($Exception.InnerException) {
            $LogLine += "`n  Inner: $($Exception.InnerException.Message)"
        }
    }

    # Zapis do pliku
    if ($script:LogFilePath) {
        try {
            $LogLine | Out-File -FilePath $script:LogFilePath -Encoding UTF8 -Append
        }
        catch {
            # Jeśli nie można pisać do logu, wypisz na konsolę
            Write-Warning "Nie mozna zapisac do logu: $_"
        }
    }

    # Wypisz na konsolę z kolorem
    $Color = switch ($Level) {
        "Verbose" { "Gray" }
        "Info"    { "White" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
    }

    if ($Level -ne "Verbose" -or $script:LogLevel -eq "Verbose") {
        Write-Host $LogLine -ForegroundColor $Color
    }
}

# ============================================================
# SEKCJA: Checkpointy
# ============================================================

function Save-ScanCheckpoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CheckpointPath,

        [Parameter(Mandatory = $true)]
        [string]$ScanSessionId,

        [Parameter(Mandatory = $true)]
        [hashtable]$State
    )

    $CheckpointFile = Join-Path $CheckpointPath "Checkpoint_$ScanSessionId.json"
    try {
        $State | ConvertTo-Json -Depth 5 | Out-File -FilePath $CheckpointFile -Encoding UTF8 -Force
        Write-ScanLog -Level "Verbose" -Message "Checkpoint zapisany: $CheckpointFile"
    }
    catch {
        Write-ScanLog -Level "Warning" -Message "Blad zapisu checkpointu: $_"
    }
}

function Get-LatestCheckpoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CheckpointPath
    )

    $Checkpoints = Get-ChildItem -Path $CheckpointPath -Filter "Checkpoint_*.json" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending

    if ($Checkpoints.Count -eq 0) {
        return $null
    }

    try {
        $Latest = $Checkpoints[0]
        $Data = Get-Content -Path $Latest.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        Write-ScanLog -Level "Info" -Message "Wczytano checkpoint: $($Latest.FullName)"
        return $Data
    }
    catch {
        Write-ScanLog -Level "Warning" -Message "Nie mozna wczytac checkpointu: $_"
        return $null
    }
}

# ============================================================
# SEKCJA: Pomocnicze funkcje weryfikacji
# ============================================================

function Test-SPFarmAdminRole {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try {
        $Farm = [Microsoft.SharePoint.Administration.SPFarm]::Local
        $CurrentUser = "$env:USERDOMAIN\$env:USERNAME"

        foreach ($Admin in $Farm.CurrentUserIsAdministrator()) {
            if ($Admin) { return $true }
        }

        # Alternatywna metoda
        $Farm.CurrentUserIsAdministrator("Farm")
        return $true
    }
    catch {
        # Jeśli nie ma dostępu do Farm.CurrentUserIsAdministrator, spróbuj inaczej
        try {
            $WebService = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
            if ($null -ne $WebService) { return $true }
        }
        catch {
            return $false
        }
    }
    return $false
}

function Test-IsSystemList {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPList]$List,

        [Parameter(Mandatory = $false)]
        [object]$Exclusions = $null
    )

    # UWAGA: $List.Hidden = true oznacza tylko ukrycie z nawigacji, NIE klasyfikuje listy jako systemowej.
    # Flaga Hidden jest obsługiwana przez ExcludeHiddenLists - nie używamy jej tutaj.

    # Sprawdź typ szablonu listy
    $SystemTemplates = @(
        [Microsoft.SharePoint.SPListTemplateType]::WebPageLibrary,
        [Microsoft.SharePoint.SPListTemplateType]::MasterPageCatalog,
        [Microsoft.SharePoint.SPListTemplateType]::ListTemplateCatalog,
        [Microsoft.SharePoint.SPListTemplateType]::WebTemplateCatalog,
        [Microsoft.SharePoint.SPListTemplateType]::SolutionCatalog,
        [Microsoft.SharePoint.SPListTemplateType]::ThemeCatalog
    )

    if ($SystemTemplates -contains $List.BaseTemplate) { return $true }

    # Sprawdź po tytule w exclusions
    if ($null -ne $Exclusions -and $Exclusions.SystemListTitles) {
        $SystemTitles = $Exclusions.SystemListTitles | ForEach-Object { $_.ToLower() }
        if ($SystemTitles -contains $List.Title.ToLower()) { return $true }
    }

    # Sprawdź po BaseTemplateId w exclusions
    if ($null -ne $Exclusions -and $Exclusions.SystemListBaseTemplateIds) {
        $BaseId = [int]$List.BaseTemplate
        if ($Exclusions.SystemListBaseTemplateIds -contains $BaseId) { return $true }
    }

    # Specjalne sprawdzenia (IsCatalog niedostepny na niektórych typach list w SP SE)
    $IsCatalogCheck = $false
    try { $IsCatalogCheck = $List.IsCatalog } catch { }
    if ($IsCatalogCheck) { return $true }

    return $false
}

function Test-IsExcludedWebApplication {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$WebAppUrl,
        [object]$Config
    )

    if ($Config.FarmSettings.WebApplicationUrls.Count -gt 0) {
        $NormalizedInput = $WebAppUrl.TrimEnd('/')
        foreach ($AllowedUrl in $Config.FarmSettings.WebApplicationUrls) {
            if ($AllowedUrl.TrimEnd('/') -eq $NormalizedInput) { return $false }
        }
        return $true  # Nie ma na liście dozwolonych = wyklucz
    }
    return $false  # Pusta lista = skanuj wszystkie
}

function Test-IsExcludedSiteCollection {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$SiteUrl,
        [object]$Config,
        [object]$Exclusions
    )

    # Sprawdź listę dozwolonych SC
    if ($Config.FarmSettings.SiteCollectionUrls.Count -gt 0) {
        $NormalizedInput = $SiteUrl.TrimEnd('/')
        foreach ($AllowedUrl in $Config.FarmSettings.SiteCollectionUrls) {
            if ($AllowedUrl.TrimEnd('/') -eq $NormalizedInput) { return $false }
        }
        return $true
    }

    # Sprawdź listę wykluczonych
    if ($Exclusions.ExcludedSiteCollections) {
        foreach ($ExcUrl in $Exclusions.ExcludedSiteCollections) {
            if ($ExcUrl -and $SiteUrl -like "*$($ExcUrl.TrimEnd('/'))*") { return $true }
        }
    }

    return $false
}

# ============================================================
# SEKCJA: Budowanie obiektów danych
# ============================================================

function New-ScanObject {
    <#
    .SYNOPSIS
        Tworzy ustandaryzowany obiekt reprezentujacy przeskanowany element.
    #>
    [CmdletBinding()]
    param(
        [string]$ObjectId = [System.Guid]::NewGuid().ToString(),
        [string]$ParentObjectId = "",
        [ValidateSet("WebApplication","SiteCollection","Web","List","Library","Folder","File","ListItem")]
        [string]$ObjectType = "Web",
        [string]$WebApplicationUrl = "",
        [string]$SiteCollectionUrl = "",
        [string]$WebUrl = "",
        [string]$FullUrl = "",
        [string]$ServerRelativeUrl = "",
        [string]$Title = "",
        [string]$Name = "",
        [string]$ListTitle = "",
        [string]$ListId = "",
        [object]$ItemId = $null,
        [string]$FileLeafRef = "",
        [bool]$IsHidden = $false,
        [bool]$IsSystem = $false,
        [bool]$IsCatalog = $false,
        [bool]$IsSiteAssets = $false,
        [bool]$HasUniquePermissions = $false,
        [string]$InheritsFromUrl = "",
        [string]$InheritsFromId = "",
        [string]$FirstUniqueAncestorId = "",
        [string]$FirstUniqueAncestorUrl = "",
        [object[]]$Assignments = @()
    )

    return [ordered]@{
        ObjectId               = $ObjectId
        ParentObjectId         = $ParentObjectId
        ObjectType             = $ObjectType
        WebApplicationUrl      = $WebApplicationUrl
        SiteCollectionUrl      = $SiteCollectionUrl
        WebUrl                 = $WebUrl
        FullUrl                = $FullUrl
        ServerRelativeUrl      = $ServerRelativeUrl
        Title                  = $Title
        Name                   = $Name
        ListTitle              = $ListTitle
        ListId                 = $ListId
        ItemId                 = $ItemId
        FileLeafRef            = $FileLeafRef
        IsHidden               = $IsHidden
        IsSystem               = $IsSystem
        IsCatalog              = $IsCatalog
        IsSiteAssets           = $IsSiteAssets
        HasUniquePermissions   = $HasUniquePermissions
        InheritsFromUrl        = $InheritsFromUrl
        InheritsFromId         = $InheritsFromId
        FirstUniqueAncestorId  = $FirstUniqueAncestorId
        FirstUniqueAncestorUrl = $FirstUniqueAncestorUrl
        Assignments            = $Assignments
        ScanTimestamp          = (Get-Date -Format "o")
    }
}

function New-AssignmentRecord {
    <#
    .SYNOPSIS
        Tworzy ustandaryzowany rekord przypisania uprawnienia.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet("User","SharePointGroup","DomainGroup","Claim","SpecialPrincipal")]
        [string]$PrincipalType = "User",
        [string]$LoginName = "",
        [string]$DisplayName = "",
        [string]$Email = "",
        [string]$SID = "",
        [ValidateSet("Direct","ViaSharePointGroup","ViaDomainGroup","Inherited")]
        [string]$SourceType = "Direct",
        [string]$SourceName = "",
        [string]$SourceId = "",
        [string[]]$PermissionLevels = @(),
        [bool]$IsLimitedAccessOnly = $false,
        [bool]$IsSiteAdmin = $false,
        [bool]$IsWebAppPolicy = $false,
        [bool]$IsActive = $true,
        [bool]$IsOrphaned = $false,
        [bool]$IsUnresolved = $false,
        [string[]]$InheritancePath = @()
    )

    return [ordered]@{
        PrincipalType     = $PrincipalType
        LoginName         = $LoginName
        DisplayName       = $DisplayName
        Email             = $Email
        SID               = $SID
        SourceType        = $SourceType
        SourceName        = $SourceName
        SourceId          = $SourceId
        PermissionLevels  = $PermissionLevels
        IsLimitedAccessOnly = $IsLimitedAccessOnly
        IsSiteAdmin       = $IsSiteAdmin
        IsWebAppPolicy    = $IsWebAppPolicy
        IsActive          = $IsActive
        IsOrphaned        = $IsOrphaned
        IsUnresolved      = $IsUnresolved
        InheritancePath   = $InheritancePath
    }
}

# ============================================================
# SEKCJA: Analiza principali i uprawnień
# ============================================================

function Get-PrincipalType {
    <#
    .SYNOPSIS
        Rozpoznaje typ principala na podstawie LoginName i obiektu SPPrincipal.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Principal
    )

    $LoginName = $Principal.LoginName
    $Type = $Principal.GetType().Name

    # SPGroup
    if ($Type -eq "SPGroup") { return "SharePointGroup" }

    # SPUser - analizuj LoginName
    if ($Type -eq "SPUser") {
        if ($LoginName -match "^c:0[.!]") {
            # Claims
            if ($LoginName -match "c:0\+\.f\|membership\|") { return "User" }
            if ($LoginName -match "c:0-\.f\|rolemanager\|") { return "DomainGroup" }
            if ($LoginName -match "c:0!\.s\|") { return "SpecialPrincipal" }
            return "Claim"
        }
        if ($LoginName -match "\\") {
            # Klasyczny NTLM: DOMAIN\user lub DOMAIN\group
            # Grupy AD mają DisplayName inne od LoginName i są isEveryone=false
            try {
                if ($Principal.IsDomainGroup) { return "DomainGroup" }
            }
            catch { }
            return "User"
        }
        if ($LoginName -eq "NT AUTHORITY\AUTHENTICATED USERS" -or
            $LoginName -eq "NT AUTHORITY\ALL USERS" -or
            $LoginName -eq "SHAREPOINT\system" -or
            $LoginName -match "^EVERYONE$" -or
            $LoginName -match "^NT AUTHORITY\\") {
            return "SpecialPrincipal"
        }
        return "User"
    }

    return "User"
}

function Get-PermissionLevelNames {
    <#
    .SYNOPSIS
        Pobiera nazwy poziomow uprawnien z role assignment.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPRoleAssignment]$RoleAssignment
    )

    $Levels = @()
    try {
        foreach ($RoleDef in $RoleAssignment.RoleDefinitionBindings) {
            if ($RoleDef.Name -ne "Limited Access") {
                $Levels += $RoleDef.Name
            }
            else {
                $Levels += "Limited Access"
            }
        }
    }
    catch {
        Write-ScanLog -Level "Warning" -Message "Blad pobierania poziomow uprawnien: $_"
    }
    return $Levels
}

function Test-IsLimitedAccessOnly {
    <#
    .SYNOPSIS
        Sprawdza czy principal ma TYLKO Limited Access (bez zadnych innych uprawnien).
    #>
    [CmdletBinding()]
    param(
        [Microsoft.SharePoint.SPRoleAssignment]$RoleAssignment
    )

    try {
        $Levels = @($RoleAssignment.RoleDefinitionBindings)
        if ($Levels.Count -eq 0) { return $false }
        $NonLimited = $Levels | Where-Object { $_.Name -ne "Limited Access" }
        return ($NonLimited.Count -eq 0)
    }
    catch {
        return $false
    }
}

function Get-RoleAssignmentDetails {
    <#
    .SYNOPSIS
        Przetwarza jedno role assignment i zwraca liste rekordow przypizan.
        Opcjonalnie rozwija grupy SharePoint i domenowe.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPRoleAssignment]$RoleAssignment,

        [Parameter(Mandatory = $false)]
        [string[]]$InheritancePath = @(),

        [Parameter(Mandatory = $false)]
        [bool]$ExpandSharePointGroups = $true,

        [Parameter(Mandatory = $false)]
        [bool]$ExpandDomainGroups = $false,

        [Parameter(Mandatory = $false)]
        [bool]$RawOnly = $false
    )

    $Results = [System.Collections.Generic.List[hashtable]]::new()

    try {
        $Principal = $RoleAssignment.Member
        $PrincipalType = Get-PrincipalType -Principal $Principal
        $PermLevels = Get-PermissionLevelNames -RoleAssignment $RoleAssignment
        $IsLimitedOnly = Test-IsLimitedAccessOnly -RoleAssignment $RoleAssignment
        $LoginName = $Principal.LoginName
        $DisplayName = $Principal.Name
        $Email = ""
        $SID = ""

        # Pobierz email i SID dla SPUser
        if ($Principal.GetType().Name -eq "SPUser") {
            try { $Email = $Principal.Email } catch { }
            try { $SID = $Principal.Sid.ToString() } catch { }
        }

        if ($RawOnly) {
            # Tylko surowe przypisanie bez ekspansji
            $Record = New-AssignmentRecord `
                -PrincipalType $PrincipalType `
                -LoginName $LoginName `
                -DisplayName $DisplayName `
                -Email $Email `
                -SID $SID `
                -SourceType "Direct" `
                -SourceName $DisplayName `
                -SourceId ($Principal.ID.ToString()) `
                -PermissionLevels $PermLevels `
                -IsLimitedAccessOnly $IsLimitedOnly `
                -InheritancePath $InheritancePath
            $Results.Add($Record)
            return $Results.ToArray()
        }

        # Grupy SharePoint - rozwiń do członków
        if ($PrincipalType -eq "SharePointGroup" -and $ExpandSharePointGroups) {
            $SPGroup = $Principal -as [Microsoft.SharePoint.SPGroup]

            # Dodaj sam rekord grupy
            $GroupRecord = New-AssignmentRecord `
                -PrincipalType "SharePointGroup" `
                -LoginName $LoginName `
                -DisplayName $DisplayName `
                -SourceType "Direct" `
                -SourceName $DisplayName `
                -SourceId ($SPGroup.ID.ToString()) `
                -PermissionLevels $PermLevels `
                -IsLimitedAccessOnly $IsLimitedOnly `
                -InheritancePath $InheritancePath
            $Results.Add($GroupRecord)

            # Rozwiń członków
            try {
                foreach ($Member in $SPGroup.Users) {
                    $MemberType = Get-PrincipalType -Principal $Member
                    $MemberEmail = ""
                    $MemberSID = ""
                    try { $MemberEmail = $Member.Email } catch { }
                    try { $MemberSID = $Member.Sid.ToString() } catch { }

                    if ($MemberType -eq "DomainGroup" -and $ExpandDomainGroups) {
                        # Rozwiń grupę domenową
                        $DomainMembers = Expand-DomainGroupMembers -LoginName $Member.LoginName -GroupName $Member.Name
                        foreach ($DomainMember in $DomainMembers) {
                            $DomainRecord = New-AssignmentRecord `
                                -PrincipalType "User" `
                                -LoginName $DomainMember.LoginName `
                                -DisplayName $DomainMember.DisplayName `
                                -Email $DomainMember.Email `
                                -SID $DomainMember.SID `
                                -SourceType "ViaSharePointGroup" `
                                -SourceName "$DisplayName > $($Member.Name)" `
                                -SourceId "$($SPGroup.ID)>$($Member.LoginName)" `
                                -PermissionLevels $PermLevels `
                                -IsLimitedAccessOnly $IsLimitedOnly `
                                -IsActive $DomainMember.IsActive `
                                -InheritancePath ($InheritancePath + @($DisplayName, $Member.Name))
                            $Results.Add($DomainRecord)
                        }
                    }
                    else {
                        $MemberRecord = New-AssignmentRecord `
                            -PrincipalType $MemberType `
                            -LoginName $Member.LoginName `
                            -DisplayName $Member.Name `
                            -Email $MemberEmail `
                            -SID $MemberSID `
                            -SourceType "ViaSharePointGroup" `
                            -SourceName $DisplayName `
                            -SourceId $SPGroup.ID.ToString() `
                            -PermissionLevels $PermLevels `
                            -IsLimitedAccessOnly $IsLimitedOnly `
                            -InheritancePath ($InheritancePath + @($DisplayName))
                        $Results.Add($MemberRecord)
                    }
                }
            }
            catch {
                Write-ScanLog -Level "Warning" -Message "Blad ekspansji grupy SP '$DisplayName': $_"
            }
        }
        # Grupy domenowe - bezpośrednio przypisane do ACL
        elseif ($PrincipalType -eq "DomainGroup") {
            $GroupRecord = New-AssignmentRecord `
                -PrincipalType "DomainGroup" `
                -LoginName $LoginName `
                -DisplayName $DisplayName `
                -Email $Email `
                -SourceType "Direct" `
                -SourceName $DisplayName `
                -SourceId ($Principal.ID.ToString()) `
                -PermissionLevels $PermLevels `
                -IsLimitedAccessOnly $IsLimitedOnly `
                -InheritancePath $InheritancePath
            $Results.Add($GroupRecord)

            if ($ExpandDomainGroups) {
                $DomainMembers = Expand-DomainGroupMembers -LoginName $LoginName -GroupName $DisplayName
                foreach ($DomainMember in $DomainMembers) {
                    $DomainRecord = New-AssignmentRecord `
                        -PrincipalType "User" `
                        -LoginName $DomainMember.LoginName `
                        -DisplayName $DomainMember.DisplayName `
                        -Email $DomainMember.Email `
                        -SID $DomainMember.SID `
                        -SourceType "ViaDomainGroup" `
                        -SourceName $DisplayName `
                        -SourceId ($Principal.ID.ToString()) `
                        -PermissionLevels $PermLevels `
                        -IsLimitedAccessOnly $IsLimitedOnly `
                        -IsActive $DomainMember.IsActive `
                        -InheritancePath ($InheritancePath + @($DisplayName))
                    $Results.Add($DomainRecord)
                }
            }
        }
        else {
            # Zwykły użytkownik lub specjalny principal
            $IsOrphaned = $false
            $IsUnresolved = $false

            if ($Principal.GetType().Name -eq "SPUser") {
                $SPUser = $Principal -as [Microsoft.SharePoint.SPUser]
                try {
                    # Sprawdź czy konto jest "orphaned" (nierozpoznane)
                    if ($SPUser.LoginName -match "^[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}#") {
                        $IsUnresolved = $true
                    }
                }
                catch { }
            }

            $UserRecord = New-AssignmentRecord `
                -PrincipalType $PrincipalType `
                -LoginName $LoginName `
                -DisplayName $DisplayName `
                -Email $Email `
                -SID $SID `
                -SourceType "Direct" `
                -SourceName $DisplayName `
                -SourceId ($Principal.ID.ToString()) `
                -PermissionLevels $PermLevels `
                -IsLimitedAccessOnly $IsLimitedOnly `
                -IsOrphaned $IsOrphaned `
                -IsUnresolved $IsUnresolved `
                -InheritancePath $InheritancePath
            $Results.Add($UserRecord)
        }
    }
    catch {
        Write-ScanLog -Level "Warning" -Message "Blad przetwarzania role assignment: $_"
    }

    return $Results.ToArray()
}

function Get-SecurableObjectAssignments {
    <#
    .SYNOPSIS
        Pobiera wszystkie przypisania uprawnien dla obiektu SPSecurableObject.
        Zwraca liste rekordow AssignmentRecord.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SecurableObject,

        [Parameter(Mandatory = $false)]
        [string[]]$InheritancePath = @(),

        [Parameter(Mandatory = $false)]
        [bool]$ExpandSharePointGroups = $true,

        [Parameter(Mandatory = $false)]
        [bool]$ExpandDomainGroups = $false,

        [Parameter(Mandatory = $false)]
        [bool]$RawOnly = $false
    )

    $AllAssignments = [System.Collections.Generic.List[hashtable]]::new()

    try {
        foreach ($RoleAssignment in $SecurableObject.RoleAssignments) {
            $Records = Get-RoleAssignmentDetails `
                -RoleAssignment $RoleAssignment `
                -InheritancePath $InheritancePath `
                -ExpandSharePointGroups $ExpandSharePointGroups `
                -ExpandDomainGroups $ExpandDomainGroups `
                -RawOnly $RawOnly

            foreach ($Record in $Records) {
                $AllAssignments.Add($Record)
            }
        }
    }
    catch {
        Write-ScanLog -Level "Warning" -Message "Blad pobierania RoleAssignments: $_"
    }

    return $AllAssignments.ToArray()
}

function Get-FirstUniqueAncestor {
    <#
    .SYNOPSIS
        Zwraca URL i ID pierwszego przodka z unikatowymi uprawnieniami.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SecurableObject,

        [Parameter(Mandatory = $false)]
        [string]$CurrentWebUrl = ""
    )

    $AncestorUrl = ""
    $AncestorId = ""

    try {
        $Ancestor = $SecurableObject.FirstUniqueAncestorSecurableObject
        if ($null -ne $Ancestor) {
            $AncestorType = $Ancestor.GetType().Name
            switch ($AncestorType) {
                "SPWeb" {
                    $AncestorUrl = $Ancestor.Url
                    $AncestorId = $Ancestor.ID.ToString()
                }
                "SPList" {
                    $AncestorUrl = "$CurrentWebUrl/$($Ancestor.RootFolder.Url)"
                    $AncestorId = $Ancestor.ID.ToString()
                }
                "SPListItem" {
                    $AncestorUrl = $Ancestor.Url
                    $AncestorId = "$($Ancestor.ParentList.ID)_$($Ancestor.ID)"
                }
                default {
                    $AncestorUrl = "Unknown"
                    $AncestorId = "Unknown"
                }
            }
        }
    }
    catch {
        Write-ScanLog -Level "Verbose" -Message "Blad Get-FirstUniqueAncestor: $_"
    }

    return @{ Url = $AncestorUrl; Id = $AncestorId }
}

# ============================================================
# SEKCJA: Skanowanie elementów
# ============================================================

function Get-SPItemPermissions {
    <#
    .SYNOPSIS
        Skanuje uprawnienia pojedynczego elementu listy lub pliku.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPListItem]$Item,

        [Parameter(Mandatory = $true)]
        [string]$ParentListObjectId,

        [Parameter(Mandatory = $true)]
        [string]$WebApplicationUrl,

        [Parameter(Mandatory = $true)]
        [string]$SiteCollectionUrl,

        [Parameter(Mandatory = $true)]
        [string]$WebUrl,

        [Parameter(Mandatory = $true)]
        [string]$ListTitle,

        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $false)]
        [bool]$ExpandSharePointGroups = $true,

        [Parameter(Mandatory = $false)]
        [bool]$ExpandDomainGroups = $false,

        [Parameter(Mandatory = $false)]
        [bool]$RawOnly = $false
    )

    $ItemId = $Item.ID.ToString()
    $ObjectId = "${ListId}_${ItemId}"

    try {
        # Określ typ obiektu
        $ObjectType = "ListItem"
        $FileLeafRef = ""
        $Title = ""
        $FullUrl = ""
        $ServerRelativeUrl = ""
        $Name = ""

        try {
            $FileLeafRef = [string]$Item["FileLeafRef"]
            $Name = $FileLeafRef
        }
        catch { }

        try {
            $Title = $Item.Title
            if (-not $Title) { $Title = $FileLeafRef }
        }
        catch { $Title = "Item $ItemId" }

        try {
            $FullUrl = $Item.Url
            $ServerRelativeUrl = $Item.ServerRelativeUrl
        }
        catch { }

        # Fallback ServerRelativeUrl dla folderów i plików
        if ([string]::IsNullOrEmpty($ServerRelativeUrl)) {
            try {
                if ($Item.FileSystemObjectType -eq [Microsoft.SharePoint.SPFileSystemObjectType]::Folder) {
                    $ServerRelativeUrl = $Item.Folder.ServerRelativeUrl
                }
            } catch { }
        }
        if ([string]::IsNullOrEmpty($ServerRelativeUrl)) {
            try {
                if ($Item.FileSystemObjectType -eq [Microsoft.SharePoint.SPFileSystemObjectType]::File) {
                    $ServerRelativeUrl = $Item.File.ServerRelativeUrl
                }
            } catch { }
        }
        if ([string]::IsNullOrEmpty($ServerRelativeUrl) -and -not [string]::IsNullOrEmpty($FullUrl)) {
            try {
                $WebSrvRel = ([Uri]$WebUrl).AbsolutePath.TrimEnd('/')
                $ServerRelativeUrl = $WebSrvRel + '/' + $FullUrl.TrimStart('/')
            } catch { }
        }
        if (-not [string]::IsNullOrEmpty($FullUrl) -and $FullUrl -notmatch '^[a-z]+://') {
            try {
                $FullUrl = $WebUrl.TrimEnd('/') + '/' + $FullUrl.TrimStart('/')
            } catch { }
        }
        elseif ([string]::IsNullOrEmpty($FullUrl) -and -not [string]::IsNullOrEmpty($ServerRelativeUrl)) {
            try {
                $WebAppRoot = ([Uri]$WebApplicationUrl).GetLeftPart([System.UriPartial]::Authority)
                $FullUrl = $WebAppRoot.TrimEnd('/') + $ServerRelativeUrl
            } catch { }
        }

        # Sprawdź czy to plik
        if ($Item.FileSystemObjectType -eq [Microsoft.SharePoint.SPFileSystemObjectType]::File) {
            $ObjectType = "File"
        }
        elseif ($Item.FileSystemObjectType -eq [Microsoft.SharePoint.SPFileSystemObjectType]::Folder) {
            $ObjectType = "Folder"
        }

        $HasUnique = $Item.HasUniqueRoleAssignments
        $Assignments = @()

        if ($HasUnique) {
            $AncestorInfo = Get-FirstUniqueAncestor -SecurableObject $Item -CurrentWebUrl $WebUrl
            $Assignments = Get-SecurableObjectAssignments `
                -SecurableObject $Item `
                -ExpandSharePointGroups $ExpandSharePointGroups `
                -ExpandDomainGroups $ExpandDomainGroups `
                -RawOnly $RawOnly

            $script:Stats.UniquePermissionsCount++
        }
        else {
            $AncestorInfo = @{ Url = ""; Id = $ParentListObjectId }
        }

        $script:Stats.ItemCount++
        $script:Stats.TotalObjectsScanned++
        $script:Stats.TotalAssignments += @($Assignments).Count

        return New-ScanObject `
            -ObjectId $ObjectId `
            -ParentObjectId $ParentListObjectId `
            -ObjectType $ObjectType `
            -WebApplicationUrl $WebApplicationUrl `
            -SiteCollectionUrl $SiteCollectionUrl `
            -WebUrl $WebUrl `
            -FullUrl $FullUrl `
            -ServerRelativeUrl $ServerRelativeUrl `
            -Title $Title `
            -Name $Name `
            -ListTitle $ListTitle `
            -ListId $ListId `
            -ItemId $Item.ID `
            -FileLeafRef $FileLeafRef `
            -HasUniquePermissions $HasUnique `
            -InheritsFromUrl $AncestorInfo.Url `
            -InheritsFromId $AncestorInfo.Id `
            -FirstUniqueAncestorId $AncestorInfo.Id `
            -FirstUniqueAncestorUrl $AncestorInfo.Url `
            -Assignments $Assignments
    }
    catch {
        $script:Stats.ErrorCount++
        Write-ScanLog -Level "Error" -Message "Blad skanowania elementu $ItemId w liscie $ListId`: $_"
        return $null
    }
}

function Get-SPListPermissions {
    <#
    .SYNOPSIS
        Skanuje uprawnienia listy lub biblioteki i jej elementow.
        Zwraca liste obiektow scanowanych.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPList]$List,

        [Parameter(Mandatory = $true)]
        [string]$ParentWebObjectId,

        [Parameter(Mandatory = $true)]
        [string]$WebApplicationUrl,

        [Parameter(Mandatory = $true)]
        [string]$SiteCollectionUrl,

        [Parameter(Mandatory = $true)]
        [string]$WebUrl,

        [Parameter(Mandatory = $false)]
        [object]$Config = $null,

        [Parameter(Mandatory = $false)]
        [object]$Exclusions = $null
    )

    $Results = [System.Collections.Generic.List[hashtable]]::new()
    $ListId = $List.ID.ToString()
    $ObjectId = $ListId

    try {
        $ListTitle = $List.Title
        $ListType = if ($List -is [Microsoft.SharePoint.SPDocumentLibrary]) { "Library" } else { "List" }
        $IsHidden = $List.Hidden
        $IsCatalog = $false
        try { $IsCatalog = $List.IsCatalog } catch { }
        $IsSiteAssets = ($ListTitle -eq "Site Assets" -or $ListTitle -eq "SiteAssets")

        # Sprawdź exclusions
        if ($Config) {
            if ($Config.FilterSettings.ExcludeHiddenLists -and $IsHidden) {
                # Biblioteki dokumentów skanujemy nawet gdy Hidden=true (ukryte w nawigacji, ale zawierają pliki i uprawnienia)
                if ($ListType -ne "Library") {
                    $script:Stats.SkippedObjects++
                    Write-ScanLog -Level "Verbose" -Message "Pomijanie ukrytej listy: $ListTitle ($ListId)"
                    return $Results.ToArray()
                }
                Write-ScanLog -Level "Verbose" -Message "Ukryta biblioteka dokumentów - skanowanie mimo Hidden=true: $ListTitle"
            }
            if ($Config.FilterSettings.ExcludeSystemLists -and (Test-IsSystemList -List $List -Exclusions $Exclusions)) {
                $script:Stats.SkippedObjects++
                Write-ScanLog -Level "Verbose" -Message "Pomijanie listy systemowej: $ListTitle ($ListId)"
                return $Results.ToArray()
            }
            if ($Config.FilterSettings.ExcludeCatalogLists -and $IsCatalog) {
                $script:Stats.SkippedObjects++
                Write-ScanLog -Level "Verbose" -Message "Pomijanie katalogu: $ListTitle ($ListId)"
                return $Results.ToArray()
            }
        }

        # Ścieżki listy
        $ListRootUrl = ""
        $ListServerRelativeUrl = ""
        try {
            $ListRootUrl = "$WebUrl/$($List.RootFolder.Url.TrimStart('/'))"
            $ListServerRelativeUrl = $List.RootFolder.ServerRelativeUrl
        }
        catch { }

        $HasUnique = $List.HasUniqueRoleAssignments
        $Assignments = @()
        $ExpandSPGroups = $true
        $ExpandDG = $false
        $RawOnly = $false

        if ($Config) {
            $ExpandSPGroups = [bool]$Config.PrincipalExpansion.ExpandSharePointGroups
            $ExpandDG = [bool]$Config.PrincipalExpansion.ExpandDomainGroups
            $RawOnly = [bool]$Config.PrincipalExpansion.RawAssignmentsOnly
        }

        if ($HasUnique) {
            $AncestorInfo = Get-FirstUniqueAncestor -SecurableObject $List -CurrentWebUrl $WebUrl
            $Assignments = Get-SecurableObjectAssignments `
                -SecurableObject $List `
                -ExpandSharePointGroups $ExpandSPGroups `
                -ExpandDomainGroups $ExpandDG `
                -RawOnly $RawOnly

            $script:Stats.UniquePermissionsCount++
        }
        else {
            $AncestorInfo = @{ Url = $WebUrl; Id = $ParentWebObjectId }
        }

        $script:Stats.ListCount++
        $script:Stats.TotalObjectsScanned++
        $script:Stats.TotalAssignments += @($Assignments).Count

        $ListObject = New-ScanObject `
            -ObjectId $ObjectId `
            -ParentObjectId $ParentWebObjectId `
            -ObjectType $ListType `
            -WebApplicationUrl $WebApplicationUrl `
            -SiteCollectionUrl $SiteCollectionUrl `
            -WebUrl $WebUrl `
            -FullUrl $ListRootUrl `
            -ServerRelativeUrl $ListServerRelativeUrl `
            -Title $ListTitle `
            -Name $ListTitle `
            -IsHidden $IsHidden `
            -IsCatalog $IsCatalog `
            -IsSiteAssets $IsSiteAssets `
            -HasUniquePermissions $HasUnique `
            -InheritsFromUrl $AncestorInfo.Url `
            -InheritsFromId $AncestorInfo.Id `
            -FirstUniqueAncestorId $AncestorInfo.Id `
            -FirstUniqueAncestorUrl $AncestorInfo.Url `
            -Assignments $Assignments

        $Results.Add($ListObject)

        # Skanowanie elementów listy/biblioteki (jeśli wymagane)
        $ScanItems = $true
        if ($Config) {
            if ($List -is [Microsoft.SharePoint.SPDocumentLibrary]) {
                $ScanItems = $Config.ScanDepth.ScanFiles
            }
            else {
                $ScanItems = $Config.ScanDepth.ScanListItems
            }

            if ($Config.FilterSettings.SkipListsWithInheritedPermissionsOnly -and -not $HasUnique) {
                # Sprawdź czy jakikolwiek element ma unikalne uprawnienia
                # Dla wydajności: sprawdź tylko jeśli flaga ustawiona
                $ScanItems = $false  # Pomiń jeśli lista dziedziczy i flaga ustawiona
            }
        }

        if ($ScanItems) {
            Write-ScanLog -Level "Verbose" -Message "  Skanowanie elementow listy: $ListTitle (ItemCount: $($List.ItemCount))"

            try {
                $MaxItems = 0
                if ($Config -and $Config.ScanDepth.MaxItemsPerList -gt 0) {
                    $MaxItems = [int]$Config.ScanDepth.MaxItemsPerList
                }

                # Query - pobierz tylko elementy z unikatowymi uprawnieniami lub wszystkie
                $Query = New-Object Microsoft.SharePoint.SPQuery
                $Query.ViewAttributes = "Scope='RecursiveAll'"
                $Query.Query = ""

                if ($MaxItems -gt 0) {
                    $Query.RowLimit = [uint32]$MaxItems
                }

                $Items = $List.GetItems($Query)
                $ItemCount = $Items.Count

                Write-ScanLog -Level "Verbose" -Message "    Elementow do przetworzenia: $ItemCount"

                $ProcessedItems = 0
                foreach ($Item in $Items) {
                    try {
                        # Pomiń elementy bez unikalnych uprawnień gdy konfiguracja na to zezwala
                        if ($Config -and $Config.FilterSettings.SkipListsWithInheritedPermissionsOnly) {
                            if (-not $Item.HasUniqueRoleAssignments) {
                                $script:Stats.SkippedObjects++
                                continue
                            }
                        }

                        $ItemObj = Get-SPItemPermissions `
                            -Item $Item `
                            -ParentListObjectId $ObjectId `
                            -WebApplicationUrl $WebApplicationUrl `
                            -SiteCollectionUrl $SiteCollectionUrl `
                            -WebUrl $WebUrl `
                            -ListTitle $ListTitle `
                            -ListId $ListId `
                            -ExpandSharePointGroups $ExpandSPGroups `
                            -ExpandDomainGroups $ExpandDG `
                            -RawOnly $RawOnly

                        if ($null -ne $ItemObj) {
                            $Results.Add($ItemObj)
                        }

                        $ProcessedItems++
                        if ($ProcessedItems % 100 -eq 0) {
                            Write-ScanLog -Level "Verbose" -Message "    Postep: $ProcessedItems / $ItemCount elementow"
                        }

                        # Throttling
                        if ($Config -and $Config.Performance.ThrottleDelayMs -gt 0) {
                            Start-Sleep -Milliseconds $Config.Performance.ThrottleDelayMs
                        }
                    }
                    catch {
                        $script:Stats.ErrorCount++
                        Write-ScanLog -Level "Error" -Message "Blad przetwarzania elementu w liscie $ListTitle`: $_ " -ObjectUrl $WebUrl
                        # Kontynuuj mimo błędu
                    }
                }
            }
            catch {
                $script:Stats.ErrorCount++
                Write-ScanLog -Level "Error" -Message "Blad pobierania elementow listy '$ListTitle': $_"
            }
        }

        return $Results.ToArray()
    }
    catch {
        $script:Stats.ErrorCount++
        Write-ScanLog -Level "Error" -Message "Blad skanowania listy '$($List.Title)' ($ListId): $_"
        return $Results.ToArray()
    }
}

function Get-SPWebPermissions {
    <#
    .SYNOPSIS
        Skanuje uprawnienia witryny (SPWeb) i wszystkich jej list.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPWeb]$Web,

        [Parameter(Mandatory = $true)]
        [string]$ParentSiteObjectId,

        [Parameter(Mandatory = $true)]
        [string]$WebApplicationUrl,

        [Parameter(Mandatory = $true)]
        [string]$SiteCollectionUrl,

        [Parameter(Mandatory = $false)]
        [object]$Config = $null,

        [Parameter(Mandatory = $false)]
        [object]$Exclusions = $null
    )

    $Results = [System.Collections.Generic.List[hashtable]]::new()
    $WebId = $Web.ID.ToString()
    $WebUrl = $Web.Url
    $WebTitle = $Web.Title

    Write-ScanLog -Level "Info" -Message "  Skanowanie Web: $WebUrl"

    try {
        $HasUnique = $Web.HasUniqueRoleAssignments
        $Assignments = @()
        $ExpandSPGroups = $true
        $ExpandDG = $false
        $RawOnly = $false

        if ($Config) {
            $ExpandSPGroups = [bool]$Config.PrincipalExpansion.ExpandSharePointGroups
            $ExpandDG = [bool]$Config.PrincipalExpansion.ExpandDomainGroups
            $RawOnly = [bool]$Config.PrincipalExpansion.RawAssignmentsOnly
        }

        if ($HasUnique) {
            $AncestorInfo = Get-FirstUniqueAncestor -SecurableObject $Web -CurrentWebUrl $WebUrl
            $Assignments = Get-SecurableObjectAssignments `
                -SecurableObject $Web `
                -ExpandSharePointGroups $ExpandSPGroups `
                -ExpandDomainGroups $ExpandDG `
                -RawOnly $RawOnly

            $script:Stats.UniquePermissionsCount++
        }
        else {
            $AncestorInfo = @{ Url = $SiteCollectionUrl; Id = $ParentSiteObjectId }
        }

        $script:Stats.WebCount++
        $script:Stats.TotalObjectsScanned++
        $script:Stats.TotalAssignments += @($Assignments).Count

        $IsRootWeb = ($Web.IsRootWeb)
        $WebObject = New-ScanObject `
            -ObjectId $WebId `
            -ParentObjectId $ParentSiteObjectId `
            -ObjectType "Web" `
            -WebApplicationUrl $WebApplicationUrl `
            -SiteCollectionUrl $SiteCollectionUrl `
            -WebUrl $WebUrl `
            -FullUrl $WebUrl `
            -ServerRelativeUrl $Web.ServerRelativeUrl `
            -Title $WebTitle `
            -Name $Web.Name `
            -HasUniquePermissions $HasUnique `
            -InheritsFromUrl $AncestorInfo.Url `
            -InheritsFromId $AncestorInfo.Id `
            -FirstUniqueAncestorId $AncestorInfo.Id `
            -FirstUniqueAncestorUrl $AncestorInfo.Url `
            -Assignments $Assignments

        $Results.Add($WebObject)

        # Skanowanie list i bibliotek
        if (-not $Config -or $Config.ScanDepth.ScanLists) {
            Write-ScanLog -Level "Verbose" -Message "  Skanowanie list w: $WebUrl (Lists: $($Web.Lists.Count))"

            foreach ($List in $Web.Lists) {
                try {
                    $ListResults = Get-SPListPermissions `
                        -List $List `
                        -ParentWebObjectId $WebId `
                        -WebApplicationUrl $WebApplicationUrl `
                        -SiteCollectionUrl $SiteCollectionUrl `
                        -WebUrl $WebUrl `
                        -Config $Config `
                        -Exclusions $Exclusions

                    foreach ($LR in $ListResults) {
                        $Results.Add($LR)
                    }
                }
                catch {
                    $script:Stats.ErrorCount++
                    Write-ScanLog -Level "Error" -Message "Blad skanowania listy '$($List.Title)' w $WebUrl`: $_"
                }
            }
        }

        return $Results.ToArray()
    }
    catch {
        $script:Stats.ErrorCount++
        Write-ScanLog -Level "Error" -Message "Blad skanowania Web $WebUrl`: $_"
        return $Results.ToArray()
    }
}

function Get-SPSiteCollectionPermissions {
    <#
    .SYNOPSIS
        Skanuje uprawnienia Site Collection i wszystkich jej witryn.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPSite]$Site,

        [Parameter(Mandatory = $true)]
        [string]$ParentWebAppObjectId,

        [Parameter(Mandatory = $true)]
        [string]$WebApplicationUrl,

        [Parameter(Mandatory = $false)]
        [object]$Config = $null,

        [Parameter(Mandatory = $false)]
        [object]$Exclusions = $null
    )

    $Results = [System.Collections.Generic.List[hashtable]]::new()
    $SiteId = $Site.ID.ToString()
    $SiteUrl = $Site.Url
    $SiteTitle = ""

    Write-ScanLog -Level "Info" -Message " Skanowanie Site Collection: $SiteUrl"

    try {
        # Pobierz tytuł ze strony głównej
        try {
            $SiteTitle = $Site.RootWeb.Title
        }
        catch {
            $SiteTitle = $SiteUrl
        }

        # Pobierz administratorów SC
        $SiteAdmins = @()
        try {
            foreach ($Admin in $Site.RootWeb.SiteAdministrators) {
                $AdminEmail = ""
                $AdminSID = ""
                try { $AdminEmail = $Admin.Email } catch { }
                try { $AdminSID = $Admin.Sid.ToString() } catch { }

                $AdminRecord = New-AssignmentRecord `
                    -PrincipalType (Get-PrincipalType -Principal $Admin) `
                    -LoginName $Admin.LoginName `
                    -DisplayName $Admin.Name `
                    -Email $AdminEmail `
                    -SID $AdminSID `
                    -SourceType "Direct" `
                    -SourceName "Site Collection Administrator" `
                    -PermissionLevels @("Site Collection Administrator") `
                    -IsSiteAdmin $true
                $SiteAdmins += $AdminRecord
            }
        }
        catch {
            Write-ScanLog -Level "Warning" -Message "Blad pobierania adminow SC $SiteUrl`: $_"
        }

        $SiteObject = New-ScanObject `
            -ObjectId $SiteId `
            -ParentObjectId $ParentWebAppObjectId `
            -ObjectType "SiteCollection" `
            -WebApplicationUrl $WebApplicationUrl `
            -SiteCollectionUrl $SiteUrl `
            -WebUrl $SiteUrl `
            -FullUrl $SiteUrl `
            -ServerRelativeUrl $Site.ServerRelativeUrl `
            -Title $SiteTitle `
            -Name $Site.RootWeb.Name `
            -HasUniquePermissions $true `
            -Assignments $SiteAdmins

        $Results.Add($SiteObject)
        $script:Stats.SiteCollectionCount++
        $script:Stats.TotalObjectsScanned++
        $script:Stats.UniquePermissionsCount++
        $script:Stats.TotalAssignments += $SiteAdmins.Count

        # Skanowanie wszystkich witryn w SC
        if (-not $Config -or $Config.ScanDepth.ScanWebs) {
            Write-ScanLog -Level "Verbose" -Message " Skanowanie Webs w SC: $SiteUrl"

            # Mapa URL→ObjectId do wyznaczania rzeczywistego rodzica sub-witryn
            $WebUrlMap = @{}
            $WebUrlMap[$SiteUrl.TrimEnd('/')] = $SiteId

            # Sortuj od najkrótszego URL - gwarantuje przetworzenie rodzica przed dziećmi
            $SortedWebs = $Site.AllWebs | Sort-Object { $_.Url.Length }

            foreach ($Web in $SortedWebs) {
                try {
                    # Sprawdź exclusions
                    if ($Config -and $Config.FarmSettings.WebUrls.Count -gt 0) {
                        if ($Config.FarmSettings.WebUrls -notcontains $Web.Url) {
                            [void]$Web.Dispose()
                            continue
                        }
                    }

                    # Wyznacz rzeczywistego rodzica przez dopasowanie najdłuższego prefiksu URL
                    $WebUrl = $Web.Url.TrimEnd('/')
                    $BestParentId = $SiteId
                    $BestParentLen = 0
                    foreach ($MapEntry in $WebUrlMap.GetEnumerator()) {
                        $MapUrl = $MapEntry.Key
                        if ($WebUrl.StartsWith($MapUrl + '/') -and $MapUrl.Length -gt $BestParentLen) {
                            $BestParentLen = $MapUrl.Length
                            $BestParentId = $MapEntry.Value
                        }
                    }

                    $WebResults = Get-SPWebPermissions `
                        -Web $Web `
                        -ParentSiteObjectId $BestParentId `
                        -WebApplicationUrl $WebApplicationUrl `
                        -SiteCollectionUrl $SiteUrl `
                        -Config $Config `
                        -Exclusions $Exclusions

                    foreach ($WR in $WebResults) {
                        $Results.Add($WR)
                    }

                    # Zapisz URL→ID przed Dispose (potrzebne dla głębiej zagnieżdżonych sub-witryn)
                    $WebUrlMap[$WebUrl] = $Web.ID.ToString()

                    # Dispose Web po użyciu
                    [void]$Web.Dispose()
                }
                catch {
                    $script:Stats.ErrorCount++
                    Write-ScanLog -Level "Error" -Message "Blad skanowania Web w SC $SiteUrl`: $_"
                    try { [void]$Web.Dispose() } catch { }
                }

                # Throttling
                if ($Config -and $Config.Performance.ThrottleDelayMs -gt 0) {
                    Start-Sleep -Milliseconds $Config.Performance.ThrottleDelayMs
                }
            }
        }

        return $Results.ToArray()
    }
    catch {
        $script:Stats.ErrorCount++
        Write-ScanLog -Level "Error" -Message "Blad skanowania SC $SiteUrl`: $_"
        return $Results.ToArray()
    }
}

function Get-SPWebApplicationPermissions {
    <#
    .SYNOPSIS
        Skanuje uprawnienia Web Application i wszystkich jej Site Collections.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Administration.SPWebApplication]$WebApplication,

        [Parameter(Mandatory = $false)]
        [object]$Config = $null,

        [Parameter(Mandatory = $false)]
        [object]$Exclusions = $null,

        # URL przekazany z zewnątrz (już rozwiązany przez Invoke-SPFarmScan)
        # – omija problematyczne $WebApplication.Url w środowiskach SP SE
        [Parameter(Mandatory = $false)]
        [string]$ResolvedUrl = ""
    )

    $Results = [System.Collections.Generic.List[hashtable]]::new()

    # Ustal URL – preferuj przekazany parametr, fallback na właściwość SSOM
    if (-not [string]::IsNullOrEmpty($ResolvedUrl)) {
        $WebAppUrl = $ResolvedUrl.TrimEnd('/')
    } else {
        $WebAppUrl = $null
        try { $WebAppUrl = $WebApplication.Url } catch { }
        if ([string]::IsNullOrEmpty($WebAppUrl)) {
            try {
                $Uri = $WebApplication.GetResponseUri([Microsoft.SharePoint.Administration.SPUrlZone]::Default)
                if ($null -ne $Uri) { $WebAppUrl = $Uri.AbsoluteUri }
            } catch { }
        }
        if ([string]::IsNullOrEmpty($WebAppUrl)) {
            try {
                foreach ($AltUrl in $WebApplication.AlternateUrls) {
                    if ($AltUrl.UrlZone -eq [Microsoft.SharePoint.Administration.SPUrlZone]::Default) {
                        $WebAppUrl = $AltUrl.PublicUrl; break
                    }
                }
                if ([string]::IsNullOrEmpty($WebAppUrl) -and $WebApplication.AlternateUrls.Count -gt 0) {
                    $WebAppUrl = $WebApplication.AlternateUrls[0].PublicUrl
                }
            } catch { }
        }
        if ([string]::IsNullOrEmpty($WebAppUrl)) {
            Write-ScanLog -Level "Warning" -Message "Get-SPWebApplicationPermissions: nie mozna ustalic URL dla '$($WebApplication.Name)'"
            return $Results
        }
        $WebAppUrl = $WebAppUrl.TrimEnd('/')
    }

    $WebAppId = $WebApplication.Id.ToString()

    Write-ScanLog -Level "Info" -Message "Skanowanie Web Application: $WebAppUrl"

    try {
        # Pobierz Web App Policies
        $PolicyAssignments = @()
        try {
            foreach ($Policy in $WebApplication.Policies) {
                $PolicyRecord = New-AssignmentRecord `
                    -PrincipalType "User" `
                    -LoginName $Policy.UserName `
                    -DisplayName $Policy.DisplayName `
                    -SourceType "Direct" `
                    -SourceName "Web Application Policy" `
                    -PermissionLevels ($Policy.PolicyRoleBindings | ForEach-Object { $_.Name }) `
                    -IsWebAppPolicy $true
                $PolicyAssignments += $PolicyRecord
            }
        }
        catch {
            Write-ScanLog -Level "Warning" -Message "Blad pobierania WebApp Policies dla $WebAppUrl`: $_"
        }

        $WebAppObject = New-ScanObject `
            -ObjectId $WebAppId `
            -ObjectType "WebApplication" `
            -WebApplicationUrl $WebAppUrl `
            -FullUrl $WebAppUrl `
            -Title $WebApplication.Name `
            -Name $WebApplication.Name `
            -HasUniquePermissions $true `
            -Assignments $PolicyAssignments

        $Results.Add($WebAppObject)
        $script:Stats.WebApplicationCount++
        $script:Stats.TotalObjectsScanned++
        $script:Stats.TotalAssignments += $PolicyAssignments.Count

        # Skanowanie Site Collections
        Write-ScanLog -Level "Info" -Message "Skanowanie Site Collections w $WebAppUrl (SC: $($WebApplication.Sites.Count))"

        foreach ($Site in $WebApplication.Sites) {
            try {
                # Sprawdź exclusions SC
                if ($null -ne $Exclusions -and $null -ne $Config) {
                    if (Test-IsExcludedSiteCollection -SiteUrl $Site.Url -Config $Config -Exclusions $Exclusions) {
                        Write-ScanLog -Level "Info" -Message " Pomijanie SC (exclusion): $($Site.Url)"
                        [void]$Site.Dispose()
                        continue
                    }
                }

                $SCResults = Get-SPSiteCollectionPermissions `
                    -Site $Site `
                    -ParentWebAppObjectId $WebAppId `
                    -WebApplicationUrl $WebAppUrl `
                    -Config $Config `
                    -Exclusions $Exclusions

                foreach ($SR in $SCResults) {
                    $Results.Add($SR)
                }

                [void]$Site.Dispose()
            }
            catch {
                $script:Stats.ErrorCount++
                Write-ScanLog -Level "Error" -Message "Blad skanowania SC $($Site.Url): $_"
                try { [void]$Site.Dispose() } catch { }
            }

            # Checkpoint co N obiektów
            if ($Config -and $Config.Performance.EnableCheckpointing) {
                if ($script:Stats.TotalObjectsScanned % $Config.Performance.BatchSize -eq 0) {
                    $CheckpointState = @{
                        LastProcessedSite = $Site.Url
                        ObjectsScanned    = $script:Stats.TotalObjectsScanned
                        Timestamp         = (Get-Date -Format "o")
                    }
                    Save-ScanCheckpoint `
                        -CheckpointPath $Config.Performance.CheckpointPath `
                        -ScanSessionId "CURRENT" `
                        -State $CheckpointState
                }
            }

            # Throttling
            if ($Config -and $Config.Performance.ThrottleDelayMs -gt 0) {
                Start-Sleep -Milliseconds $Config.Performance.ThrottleDelayMs
            }
        }

        return $Results.ToArray()
    }
    catch {
        $script:Stats.ErrorCount++
        Write-ScanLog -Level "Error" -Message "Blad skanowania WebApp $WebAppUrl`: $_"
        return $Results.ToArray()
    }
}

# ============================================================
# SEKCJA: Główna funkcja skanowania farmy
# ============================================================

function Invoke-SPFarmScan {
    <#
    .SYNOPSIS
        Glowna funkcja skanowania calej farmy SharePoint.
        Zwraca obiekt ScanResult z wszystkimi danymi.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config,

        [Parameter(Mandatory = $false)]
        [object]$Exclusions = $null,

        [Parameter(Mandatory = $false)]
        [object]$Whitelist = $null,

        [Parameter(Mandatory = $false)]
        [string]$ScanSessionId = "",

        [Parameter(Mandatory = $false)]
        [bool]$ResumeFromCheckpoint = $false,

        [Parameter(Mandatory = $false)]
        [string]$CheckpointPath = "",

        [Parameter(Mandatory = $false)]
        [bool]$DryRun = $false
    )

    # Inicjalizuj zmienne modułu
    $script:ScanConfig = $Config
    $script:ScanExclusions = $Exclusions
    $script:ScanWhitelist = $Whitelist
    $script:Stats = @{
        WebApplicationCount    = 0
        SiteCollectionCount    = 0
        WebCount               = 0
        ListCount              = 0
        FolderCount            = 0
        ItemCount              = 0
        UniquePermissionsCount = 0
        TotalAssignments       = 0
        TotalObjectsScanned    = 0
        SkippedObjects         = 0
        ErrorCount             = 0
    }

    $AllObjects = [System.Collections.Generic.List[hashtable]]::new()
    $ScanStartTime = Get-Date

    # Wznów z checkpointu jeśli wymagane
    if ($ResumeFromCheckpoint -and $CheckpointPath) {
        $CheckpointData = Get-LatestCheckpoint -CheckpointPath $CheckpointPath
        if ($null -ne $CheckpointData) {
            Write-ScanLog -Level "Info" -Message "Wznawiam skanowanie z checkpointu: $($CheckpointData.Timestamp)"
            # Checkpoint: tutaj można zaimplementować pomijanie już przeskanowanych SC
            # Dla uproszczenia - skanujemy od nowa ale logujemy
        }
    }

    try {
        # Pobierz informacje o farmie
        $Farm = [Microsoft.SharePoint.Administration.SPFarm]::Local
        $FarmName = $Farm.Name
        $FarmBuild = $Farm.BuildVersion.ToString()

        Write-ScanLog -Level "Info" -Message "Farma: $FarmName, Build: $FarmBuild"

        # Pobierz Web Applications
        $WebService = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
        $WebApplications = $WebService.WebApplications

        if ($Config.FarmSettings.IncludeCentralAdministration) {
            $AdminService = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local
            Write-ScanLog -Level "Info" -Message "Dodano Central Administration do skanowania"
        }

        $TotalWebApps = $WebApplications.Count
        Write-ScanLog -Level "Info" -Message "Znaleziono $TotalWebApps Web Application(s)"

        $WebAppIndex = 0
        foreach ($WebApp in $WebApplications) {
            $WebAppIndex++

            # --- Ustal URL aplikacji webowej (3 metody fallback) ---
            $WebAppUrl      = $null
            $UrlErrorMsg    = $null

            # Metoda 1: .Url (wywołuje GetResponseUri wewnętrznie)
            try { $WebAppUrl = $WebApp.Url } catch { $UrlErrorMsg = $_.Exception.Message }
            Write-ScanLog -Level "Info" -Message "  WebApp[$WebAppIndex] '$($WebApp.Name)' Url='$WebAppUrl' Type='$($WebApp.GetType().Name)'"

            # Metoda 2: GetResponseUri dla strefy Default
            if ([string]::IsNullOrEmpty($WebAppUrl)) {
                try {
                    $Uri = $WebApp.GetResponseUri([Microsoft.SharePoint.Administration.SPUrlZone]::Default)
                    if ($null -ne $Uri) { $WebAppUrl = $Uri.AbsoluteUri }
                } catch { }
            }

            # Metoda 3: AlternateUrls – strefa Default, a jeśli brak to pierwsza dostępna
            if ([string]::IsNullOrEmpty($WebAppUrl)) {
                try {
                    foreach ($AltUrl in $WebApp.AlternateUrls) {
                        if ($AltUrl.UrlZone -eq [Microsoft.SharePoint.Administration.SPUrlZone]::Default) {
                            $WebAppUrl = $AltUrl.PublicUrl
                            break
                        }
                    }
                    if ([string]::IsNullOrEmpty($WebAppUrl) -and $WebApp.AlternateUrls.Count -gt 0) {
                        $WebAppUrl = $WebApp.AlternateUrls[0].PublicUrl
                        Write-ScanLog -Level "Warning" -Message "  Brak AAM Default - uzywam pierwszej dostepnej strefy: $WebAppUrl"
                    }
                } catch { }
            }

            if ([string]::IsNullOrEmpty($WebAppUrl)) {
                Write-ScanLog -Level "Warning" -Message "Pomijanie WebApp '$($WebApp.Name)' - nie mozna ustalic URL$(if ($UrlErrorMsg) { ': ' + $UrlErrorMsg })"
                continue
            }

            $WebAppUrl = $WebAppUrl.TrimEnd('/')

            # Sprawdź exclusions WebApp
            if (Test-IsExcludedWebApplication -WebAppUrl $WebAppUrl -Config $Config) {
                Write-ScanLog -Level "Info" -Message "Pomijanie WebApp (exclusion): $WebAppUrl"
                continue
            }

            Write-Host "  [$WebAppIndex/$TotalWebApps] Web Application: $WebAppUrl" -ForegroundColor Cyan
            Write-ScanLog -Level "Info" -Message "[$WebAppIndex/$TotalWebApps] Przetwarzam WebApp: $WebAppUrl"

            $WebAppResults = Get-SPWebApplicationPermissions `
                -WebApplication $WebApp `
                -ResolvedUrl $WebAppUrl `
                -Config $Config `
                -Exclusions $Exclusions

            foreach ($Result in $WebAppResults) {
                $AllObjects.Add($Result)
            }

            Write-Host "    -> Zebrano $($AllObjects.Count) obiektow lacznie" -ForegroundColor Gray
        }

        # Opcjonalnie - skanuj Central Administration
        if ($Config.FarmSettings.IncludeCentralAdministration) {
            try {
                $CAWebApp = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local
                $CAResults = Get-SPWebApplicationPermissions `
                    -WebApplication $CAWebApp `
                    -Config $Config `
                    -Exclusions $Exclusions
                foreach ($Result in $CAResults) {
                    $AllObjects.Add($Result)
                }
            }
            catch {
                Write-ScanLog -Level "Warning" -Message "Blad skanowania Central Administration: $_"
            }
        }
    }
    catch {
        Write-ScanLog -Level "Error" -Message "BLAD KRYTYCZNY podczas skanowania farmy: $_`n$($_.ScriptStackTrace)"
        throw
    }

    $ScanEndTime = Get-Date
    $ScanDuration = $ScanEndTime - $ScanStartTime

    # Zbuduj obiekt wynikowy
    $ScanResult = [ordered]@{
        ScanMetadata = [ordered]@{
            ScanSessionId    = $ScanSessionId
            ScanStartTime    = $ScanStartTime.ToString("o")
            ScanEndTime      = $ScanEndTime.ToString("o")
            ScanDuration     = [int]$ScanDuration.TotalSeconds
            FarmName         = $(try { [Microsoft.SharePoint.Administration.SPFarm]::Local.Name } catch { "Unknown" })
            FarmBuild        = $(try { [Microsoft.SharePoint.Administration.SPFarm]::Local.BuildVersion.ToString() } catch { "Unknown" })
            ScannerVersion   = "1.0.0"
            ScanServer       = $env:COMPUTERNAME
            ScanUser         = "$env:USERDOMAIN\$env:USERNAME"
            Config           = @{
                RawAssignmentsOnly  = $Config.PrincipalExpansion.RawAssignmentsOnly
                ExpandDomainGroups  = $Config.PrincipalExpansion.ExpandDomainGroups
                ExpandSPGroups      = $Config.PrincipalExpansion.ExpandSharePointGroups
                ScanDepth           = $Config.ScanDepth
            }
            LogFilePath      = ""
        }
        Statistics = [ordered]@{
            WebApplicationCount    = $script:Stats.WebApplicationCount
            SiteCollectionCount    = $script:Stats.SiteCollectionCount
            WebCount               = $script:Stats.WebCount
            ListCount              = $script:Stats.ListCount
            FolderCount            = $script:Stats.FolderCount
            ItemCount              = $script:Stats.ItemCount
            UniquePermissionsCount = $script:Stats.UniquePermissionsCount
            TotalAssignments       = $script:Stats.TotalAssignments
            TotalObjectsScanned    = $script:Stats.TotalObjectsScanned
            SkippedObjects         = $script:Stats.SkippedObjects
            ErrorCount             = $script:Stats.ErrorCount
        }
        Errors  = $script:ErrorLog.ToArray()
        Objects = $AllObjects.ToArray()
    }

    Write-ScanLog -Level "Info" -Message "Laczna liczba obiektow: $($AllObjects.Count)"
    Write-ScanLog -Level "Info" -Message "Statystyki: WebApps=$($script:Stats.WebApplicationCount), SC=$($script:Stats.SiteCollectionCount), Webs=$($script:Stats.WebCount), Lists=$($script:Stats.ListCount), Items=$($script:Stats.ItemCount)"

    return $ScanResult
}

# ============================================================
# SEKCJA: Eksport CSV
# ============================================================

function Export-ScanResultToCsv {
    <#
    .SYNOPSIS
        Eksportuje wyniki skanu do pliku CSV.
        Kazde przypisanie to osobny wiersz.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ScanResult,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $CsvRows = [System.Collections.Generic.List[object]]::new()

    foreach ($Obj in $ScanResult.Objects) {
        if ($Obj.Assignments -and $Obj.Assignments.Count -gt 0) {
            foreach ($Assignment in $Obj.Assignments) {
                $Row = [ordered]@{
                    ObjectId               = $Obj.ObjectId
                    ObjectType             = $Obj.ObjectType
                    WebApplicationUrl      = $Obj.WebApplicationUrl
                    SiteCollectionUrl      = $Obj.SiteCollectionUrl
                    WebUrl                 = $Obj.WebUrl
                    FullUrl                = $Obj.FullUrl
                    ServerRelativeUrl      = $Obj.ServerRelativeUrl
                    Title                  = $Obj.Title
                    ListTitle              = $Obj.ListTitle
                    ItemId                 = $Obj.ItemId
                    FileLeafRef            = $Obj.FileLeafRef
                    IsHidden               = $Obj.IsHidden
                    IsSystem               = $Obj.IsSystem
                    HasUniquePermissions   = $Obj.HasUniquePermissions
                    InheritsFromUrl        = $Obj.InheritsFromUrl
                    PrincipalType          = $Assignment.PrincipalType
                    LoginName              = $Assignment.LoginName
                    DisplayName            = $Assignment.DisplayName
                    Email                  = $Assignment.Email
                    SID                    = $Assignment.SID
                    SourceType             = $Assignment.SourceType
                    SourceName             = $Assignment.SourceName
                    PermissionLevels       = ($Assignment.PermissionLevels -join "; ")
                    IsLimitedAccessOnly    = $Assignment.IsLimitedAccessOnly
                    IsSiteAdmin            = $Assignment.IsSiteAdmin
                    IsWebAppPolicy         = $Assignment.IsWebAppPolicy
                    IsOrphaned             = $Assignment.IsOrphaned
                    IsUnresolved           = $Assignment.IsUnresolved
                    InheritancePath        = ($Assignment.InheritancePath -join " > ")
                }
                $CsvRows.Add([PSCustomObject]$Row)
            }
        }
        else {
            # Obiekt bez przypisań (dziedziczy)
            $Row = [ordered]@{
                ObjectId               = $Obj.ObjectId
                ObjectType             = $Obj.ObjectType
                WebApplicationUrl      = $Obj.WebApplicationUrl
                SiteCollectionUrl      = $Obj.SiteCollectionUrl
                WebUrl                 = $Obj.WebUrl
                FullUrl                = $Obj.FullUrl
                ServerRelativeUrl      = $Obj.ServerRelativeUrl
                Title                  = $Obj.Title
                ListTitle              = $Obj.ListTitle
                ItemId                 = $Obj.ItemId
                FileLeafRef            = $Obj.FileLeafRef
                IsHidden               = $Obj.IsHidden
                IsSystem               = $Obj.IsSystem
                HasUniquePermissions   = $Obj.HasUniquePermissions
                InheritsFromUrl        = $Obj.InheritsFromUrl
                PrincipalType          = ""
                LoginName              = ""
                DisplayName            = ""
                Email                  = ""
                SID                    = ""
                SourceType             = "Inherited"
                SourceName             = ""
                PermissionLevels       = ""
                IsLimitedAccessOnly    = $false
                IsSiteAdmin            = $false
                IsWebAppPolicy         = $false
                IsOrphaned             = $false
                IsUnresolved           = $false
                InheritancePath        = ""
            }
            $CsvRows.Add([PSCustomObject]$Row)
        }
    }

    $CsvRows | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation -Delimiter ";"
}

Export-ModuleMember -Function @(
    "Initialize-ScanLogger",
    "Write-ScanLog",
    "Save-ScanCheckpoint",
    "Get-LatestCheckpoint",
    "Test-SPFarmAdminRole",
    "Test-IsSystemList",
    "Test-IsExcludedWebApplication",
    "Test-IsExcludedSiteCollection",
    "New-ScanObject",
    "New-AssignmentRecord",
    "Get-PrincipalType",
    "Get-PermissionLevelNames",
    "Test-IsLimitedAccessOnly",
    "Get-RoleAssignmentDetails",
    "Get-SecurableObjectAssignments",
    "Get-FirstUniqueAncestor",
    "Get-SPItemPermissions",
    "Get-SPListPermissions",
    "Get-SPWebPermissions",
    "Get-SPSiteCollectionPermissions",
    "Get-SPWebApplicationPermissions",
    "Invoke-SPFarmScan",
    "Export-ScanResultToCsv"
)


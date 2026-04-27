#Requires -Version 5.1
<#
.SYNOPSIS
    Modul rozwijania grup Active Directory do uzytkownikow indywidualnych.

.DESCRIPTION
    Wykorzystuje System.DirectoryServices.DirectorySearcher do ekspansji grup AD.
    Nie wymaga modulu RSAT ActiveDirectory.
    Obsluguje grupy zagniezdzone rekurencyjnie.
    Cachuje wyniki ekspansji.

.NOTES
    Wymaga dostepu do Active Directory z serwera SharePoint.
    Timeout LDAP konfigurowalny w ScanConfig.json.
#>

Set-StrictMode -Version Latest

# Cache grup i ich członków (LoginName -> tablica memberów)
$script:GroupExpansionCache = [System.Collections.Generic.Dictionary[string, object[]]]::new()

# Cache nieudanych ekspansji (LoginName -> komunikat błędu)
$script:FailedExpansions = [System.Collections.Generic.Dictionary[string, string]]::new()

# Konfiguracja
$script:LdapServer = ""
$script:LdapTimeout = 30
$script:MaxNestingDepth = 10

function Initialize-ADExpander {
    <#
    .SYNOPSIS
        Inicjalizuje modul z konfiguracja.
    #>
    [CmdletBinding()]
    param(
        [string]$LdapServer = "",
        [int]$LdapTimeout = 30,
        [int]$MaxNestingDepth = 10
    )

    $script:LdapServer = $LdapServer
    $script:LdapTimeout = $LdapTimeout
    $script:MaxNestingDepth = $MaxNestingDepth
}

function Get-LdapRoot {
    <#
    .SYNOPSIS
        Zwraca korzeń LDAP do przeszukiwania.
    #>
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param()

    if ($script:LdapServer) {
        return [System.DirectoryServices.DirectoryEntry]::new($script:LdapServer)
    }
    else {
        # Automatyczne wykrycie DC
        try {
            $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $DomainName = $Domain.Name
            return [System.DirectoryServices.DirectoryEntry]::new("LDAP://$DomainName")
        }
        catch {
            return [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
        }
    }
}

function Convert-SidToString {
    <#
    .SYNOPSIS
        Konwertuje bytes SID do stringa.
    #>
    param([byte[]]$SidBytes)

    try {
        $Sid = [System.Security.Principal.SecurityIdentifier]::new($SidBytes, 0)
        return $Sid.ToString()
    }
    catch {
        return ""
    }
}

function Get-ADObjectBySamAccountName {
    <#
    .SYNOPSIS
        Szuka obiektu AD po sAMAccountName.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,

        [Parameter(Mandatory = $false)]
        [string]$Domain = ""
    )

    try {
        $LdapRoot = Get-LdapRoot
        $Searcher = [System.DirectoryServices.DirectorySearcher]::new($LdapRoot)
        $Searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$SamAccountName))"
        $Searcher.PropertiesToLoad.AddRange(@("sAMAccountName", "displayName", "mail", "objectSid", "userAccountControl", "distinguishedName"))
        $Searcher.ClientTimeout = [TimeSpan]::FromSeconds($script:LdapTimeout)
        $Searcher.ServerTimeLimit = [TimeSpan]::FromSeconds($script:LdapTimeout)

        $Result = $Searcher.FindOne()
        if ($null -ne $Result) {
            return $Result
        }
    }
    catch {
        Write-Warning "Blad wyszukiwania AD dla '$SamAccountName': $_"
    }
    return $null
}

function Get-ADGroupByDistinguishedName {
    <#
    .SYNOPSIS
        Pobiera grupe AD po DistinguishedName.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DistinguishedName
    )

    try {
        $GroupEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$DistinguishedName")
        return $GroupEntry
    }
    catch {
        return $null
    }
}

function Get-GroupMembersFromAD {
    <#
    .SYNOPSIS
        Pobiera lista memberow grupy AD (1 poziom glebokosci).
        Obsluguje grupy wewnatrz domeny (member attribute).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupDN,

        [Parameter(Mandatory = $false)]
        [int]$PageSize = 1000
    )

    $Members = [System.Collections.Generic.List[object]]::new()

    try {
        $LdapRoot = Get-LdapRoot
        $Searcher = [System.DirectoryServices.DirectorySearcher]::new($LdapRoot)

        # Szukaj wszystkich memberow grupy przez memberOf lub przez member
        $Searcher.Filter = "(memberOf=$GroupDN)"
        $Searcher.PropertiesToLoad.AddRange(@(
            "sAMAccountName", "displayName", "mail", "objectSid",
            "objectCategory", "objectClass", "userAccountControl",
            "distinguishedName", "cn"
        ))
        $Searcher.PageSize = $PageSize
        $Searcher.ClientTimeout = [TimeSpan]::FromSeconds($script:LdapTimeout)
        $Searcher.ServerTimeLimit = [TimeSpan]::FromSeconds($script:LdapTimeout)

        $Results = $Searcher.FindAll()
        foreach ($Result in $Results) {
            $Members.Add($Result)
        }
        $Results.Dispose()

        # Alternatywnie - pobierz przez member attribute (dla nested/primary group)
        try {
            $GroupEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$GroupDN")
            $GroupEntry.RefreshCache(@("member"))

            if ($GroupEntry.Properties["member"].Count -gt 0) {
                foreach ($MemberDN in $GroupEntry.Properties["member"]) {
                    # Sprawdź czy już mamy tego członka
                    $AlreadyAdded = $Members | Where-Object {
                        $_.Properties["distinguishedName"].Count -gt 0 -and
                        $_.Properties["distinguishedName"][0] -eq $MemberDN
                    }
                    if (-not $AlreadyAdded) {
                        $MemberSearcher = [System.DirectoryServices.DirectorySearcher]::new($LdapRoot)
                        $MemberSearcher.Filter = "(distinguishedName=$MemberDN)"
                        $MemberSearcher.PropertiesToLoad.AddRange(@(
                            "sAMAccountName", "displayName", "mail", "objectSid",
                            "objectCategory", "objectClass", "userAccountControl",
                            "distinguishedName", "cn"
                        ))
                        $MemberResult = $MemberSearcher.FindOne()
                        if ($null -ne $MemberResult) {
                            $Members.Add($MemberResult)
                        }
                        $MemberSearcher.Dispose()
                    }
                }
            }
            $GroupEntry.Dispose()
        }
        catch {
            # Ignoruj błędy pobierania przez member attribute - mamy wyniki z memberOf
        }
    }
    catch {
        Write-Warning "Blad pobierania czlonkow grupy '$GroupDN': $_"
    }

    return $Members.ToArray()
}

function Test-IsADUserEnabled {
    <#
    .SYNOPSIS
        Sprawdza czy konto AD jest aktywne (userAccountControl).
    #>
    param([object]$SearchResult)

    try {
        if ($SearchResult.Properties["userAccountControl"].Count -gt 0) {
            $UAC = [int]$SearchResult.Properties["userAccountControl"][0]
            # Bit 2 (ACCOUNTDISABLE = 0x2) - konto wyłączone
            return (-not ($UAC -band 0x2))
        }
    }
    catch { }
    return $true  # Assume active if can't determine
}

function Test-IsADGroup {
    <#
    .SYNOPSIS
        Sprawdza czy obiekt AD jest grupa.
    #>
    param([object]$SearchResult)

    try {
        if ($SearchResult.Properties["objectClass"]) {
            return ($SearchResult.Properties["objectClass"] -contains "group")
        }
    }
    catch { }
    return $false
}

function Get-DomainFromLoginName {
    <#
    .SYNOPSIS
        Wyciaga nazwe domeny z LoginName (format: DOMAIN\user lub i:0#.w|domain\user).
    #>
    param([string]$LoginName)

    # Claims format: i:0#.w|domain\user lub i:0e.t|federation|user
    if ($LoginName -match "i:0[#e]\.w\|(.+)\\(.+)") {
        return $Matches[1]
    }
    # Classic format: DOMAIN\user
    if ($LoginName -match "^([^\\]+)\\(.+)$") {
        return $Matches[1]
    }
    return ""
}

function Get-SamAccountFromLoginName {
    <#
    .SYNOPSIS
        Wyciaga sAMAccountName z LoginName SharePoint.
    #>
    param([string]$LoginName)

    # Claims: i:0#.w|domain\user
    if ($LoginName -match "i:0[#e]\.w\|(.+)\\(.+)") {
        return $Matches[2]
    }
    # Claims grupy domenowej: c:0-.f|rolemanager|spo-grid-all-users|...
    if ($LoginName -match "c:0-\.f\|rolemanager\|(.+)$") {
        $RawName = $Matches[1]
        if ($RawName -match "\\(.+)$") {
            return $Matches[1]
        }
        return $RawName
    }
    # Classic: DOMAIN\user
    if ($LoginName -match "^[^\\]+\\(.+)$") {
        return $Matches[1]
    }
    return $LoginName
}

function Expand-DomainGroupMembers {
    <#
    .SYNOPSIS
        Rozwija grupe domenowa (AD) do listy uzytkownikow efektywnych.
        Rekurencyjnie przetwarza grupy zagniezdzone.
        Zwraca liste obiektow z polami LoginName, DisplayName, Email, SID, IsActive.

    .PARAMETER LoginName
        LoginName grupy w SharePoint (np. CONTOSO\GrupaSP lub claims).

    .PARAMETER GroupName
        Czytelna nazwa grupy (do logowania i sciezki dziedziczenia).

    .PARAMETER CurrentDepth
        Biezacy poziom zagniezdzenia (do ochrony przed petla).

    .PARAMETER ProcessedGroups
        Zbiór DN juz przetworzonych grup (anti-loop).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LoginName,

        [Parameter(Mandatory = $false)]
        [string]$GroupName = "",

        [Parameter(Mandatory = $false)]
        [int]$CurrentDepth = 0,

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.HashSet[string]]$ProcessedGroups = $null
    )

    if ($null -eq $ProcessedGroups) {
        $ProcessedGroups = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    # Sprawdź cache
    if ($script:GroupExpansionCache.ContainsKey($LoginName)) {
        return $script:GroupExpansionCache[$LoginName]
    }

    # Sprawdź czy nie mamy błędu dla tej grupy
    if ($script:FailedExpansions.ContainsKey($LoginName)) {
        return @()
    }

    # Limit głębokości
    if ($CurrentDepth -ge $script:MaxNestingDepth) {
        Write-Warning "Osiagnieto maksymalna glebokosc zagniezdzenia ($($script:MaxNestingDepth)) dla grupy: $LoginName"
        return @()
    }

    $Results = [System.Collections.Generic.List[hashtable]]::new()
    $SamAccount = Get-SamAccountFromLoginName -LoginName $LoginName
    $Domain = Get-DomainFromLoginName -LoginName $LoginName

    if (-not $SamAccount) {
        Write-Warning "Nie mozna wyciagnac sAMAccountName z LoginName: $LoginName"
        $script:FailedExpansions[$LoginName] = "Cannot parse sAMAccountName"
        return @()
    }

    try {
        # Znajdź grupę w AD
        $LdapRoot = Get-LdapRoot
        $Searcher = [System.DirectoryServices.DirectorySearcher]::new($LdapRoot)
        $Searcher.Filter = "(&(objectClass=group)(sAMAccountName=$SamAccount))"
        $Searcher.PropertiesToLoad.AddRange(@("distinguishedName", "cn", "objectSid"))
        $Searcher.ClientTimeout = [TimeSpan]::FromSeconds($script:LdapTimeout)

        $GroupResult = $Searcher.FindOne()
        $Searcher.Dispose()

        if ($null -eq $GroupResult) {
            Write-Warning "Nie znaleziono grupy AD dla: $SamAccount (LoginName: $LoginName)"
            $script:FailedExpansions[$LoginName] = "Group not found in AD"
            return @()
        }

        $GroupDN = $GroupResult.Properties["distinguishedName"][0]

        # Ochrona przed pętlą
        if ($ProcessedGroups.Contains($GroupDN)) {
            return @()
        }
        [void]$ProcessedGroups.Add($GroupDN)

        # Pobierz członków grupy
        $Members = Get-GroupMembersFromAD -GroupDN $GroupDN

        foreach ($Member in $Members) {
            try {
                $IsGroup = Test-IsADGroup -SearchResult $Member
                $MemberSam = ""
                $MemberDN = ""
                $MemberDisplay = ""
                $MemberEmail = ""
                $MemberSID = ""
                $IsActive = $true

                if ($Member.Properties["sAMAccountName"].Count -gt 0) {
                    $MemberSam = $Member.Properties["sAMAccountName"][0]
                }
                if ($Member.Properties["distinguishedName"].Count -gt 0) {
                    $MemberDN = $Member.Properties["distinguishedName"][0]
                }
                if ($Member.Properties["displayName"].Count -gt 0) {
                    $MemberDisplay = $Member.Properties["displayName"][0]
                }
                if ($Member.Properties["mail"].Count -gt 0) {
                    $MemberEmail = $Member.Properties["mail"][0]
                }
                if ($Member.Properties["objectSid"].Count -gt 0) {
                    $MemberSID = Convert-SidToString -SidBytes $Member.Properties["objectSid"][0]
                }

                $IsActive = Test-IsADUserEnabled -SearchResult $Member

                # Ustal domain prefix dla LoginName
                $MemberDomain = if ($Domain) { $Domain } else {
                    if ($MemberDN -match "DC=([^,]+)") { $Matches[1] } else { "" }
                }
                $MemberLoginName = if ($MemberDomain) { "$MemberDomain\$MemberSam" } else { $MemberSam }

                if ($IsGroup) {
                    # Rekurencyjnie rozwiń podgrupę
                    $SubMembers = Expand-DomainGroupMembers `
                        -LoginName $MemberLoginName `
                        -GroupName $MemberDisplay `
                        -CurrentDepth ($CurrentDepth + 1) `
                        -ProcessedGroups $ProcessedGroups

                    foreach ($SubMember in $SubMembers) {
                        $Results.Add($SubMember)
                    }
                }
                else {
                    # Użytkownik
                    if (-not $MemberDisplay) { $MemberDisplay = $MemberSam }

                    $UserRecord = @{
                        LoginName   = $MemberLoginName
                        DisplayName = $MemberDisplay
                        Email       = $MemberEmail
                        SID         = $MemberSID
                        IsActive    = $IsActive
                        SourceGroup = $LoginName
                    }
                    $Results.Add($UserRecord)
                }
            }
            catch {
                Write-Warning "Blad przetwarzania czlonka grupy '$LoginName': $_"
            }
        }

        # Zapisz do cache
        $ResultArray = $Results.ToArray()
        $script:GroupExpansionCache[$LoginName] = $ResultArray

        Write-ScanLog -Level "Verbose" -Message "Rozwiazano grupe '$LoginName' -> $($Results.Count) uzytkownikow (glebokosc: $CurrentDepth)"

        return $ResultArray
    }
    catch {
        Write-Warning "Blad ekspansji grupy domenowej '$LoginName': $_"
        $script:FailedExpansions[$LoginName] = $_.Message
        Write-ScanLog -Level "Warning" -Message "Blad ekspansji grupy '$LoginName': $_"
        return @()
    }
}

function Get-FailedGroupExpansions {
    <#
    .SYNOPSIS
        Zwraca liste grup, ktorych nie udalo sie rozwinac.
    #>
    [CmdletBinding()]
    param()

    return $script:FailedExpansions
}

function Clear-GroupExpansionCache {
    <#
    .SYNOPSIS
        Czysci cache rozwiniec grup.
    #>
    [CmdletBinding()]
    param()

    $script:GroupExpansionCache.Clear()
    $script:FailedExpansions.Clear()
}

Export-ModuleMember -Function @(
    "Initialize-ADExpander",
    "Expand-DomainGroupMembers",
    "Get-FailedGroupExpansions",
    "Clear-GroupExpansionCache",
    "Get-DomainFromLoginName",
    "Get-SamAccountFromLoginName"
)

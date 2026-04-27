if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell)) {
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
}

function Get-SPOSite {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$Identity,
        [Parameter(Mandatory=$false)]
        [switch]$Detailed
    )
    Test-M365Connected
    $params = @{}
    if ($PSBoundParameters.ContainsKey('Identity')) { $params.Identity = $Identity }
    if ($Detailed) { $params.Detailed = $Detailed }
    return Microsoft.Online.SharePoint.PowerShell\Get-SPOSite @params
}

function Get-SPOUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Site,
        [Parameter(Mandatory=$false)]
        [string]$LoginName
    )
    Test-M365Connected
    $params = @{}
    if ($PSBoundParameters.ContainsKey('Site')) { $params.Site = $Site }
    if ($PSBoundParameters.ContainsKey('LoginName')) { $params.LoginName = $LoginName }
    return Microsoft.Online.SharePoint.PowerShell\Get-SPOUser @params
}

function Get-SPOSiteGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Site,
        [Parameter(Mandatory=$false)]
        [string]$Group
    )
    Test-M365Connected
    $params = @{}
    if ($PSBoundParameters.ContainsKey('Site')) { $params.Site = $Site }
    if ($PSBoundParameters.ContainsKey('Group')) { $params.Group = $Group }
    return Microsoft.Online.SharePoint.PowerShell\Get-SPOSiteGroup @params
}

function Get-SPOUserOneDrive {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    Test-M365Connected
    return Microsoft.Online.SharePoint.PowerShell\Get-SPOSite -Filter "Owner -eq '$UserPrincipalName' -and Template -eq 'SPSPERS#10'" -ErrorAction SilentlyContinue
}

function Get-SPOSharingLink {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Site,
        [Parameter(Mandatory=$false)]
        [string]$FileUrl
    )
    Test-M365Connected
    $params = @{}
    if ($PSBoundParameters.ContainsKey('Site')) { $params.Site = $Site }
    if ($PSBoundParameters.ContainsKey('FileUrl')) { $params.FileUrl = $FileUrl }
    return Microsoft.Online.SharePoint.PowerShell\Get-SPOSharingLink @params
}

function Get-SPOSitePermission {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Site,
        [Parameter(Mandatory=$false)]
        [string]$Principal
    )
    Test-M365Connected
    $params = @{}
    if ($PSBoundParameters.ContainsKey('Site')) { $params.Site = $Site }
    if ($PSBoundParameters.ContainsKey('Principal')) { $params.Principal = $Principal }
    return Microsoft.Online.SharePoint.PowerShell\Get-SPOSitePermission @params
}

function Get-SPOStorageEntity {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Site,
        [Parameter(Mandatory=$false)]
        [string]$Name
    )
    Test-M365Connected
    $params = @{}
    if ($PSBoundParameters.ContainsKey('Site')) { $params.Site = $Site }
    if ($PSBoundParameters.ContainsKey('Name')) { $params.Name = $Name }
    return Microsoft.Online.SharePoint.PowerShell\Get-SPOStorageEntity @params
}

Export-ModuleMember -Function Get-SPOSite, Get-SPOUser, Get-SPOSiteGroup, Get-SPOUserOneDrive, Get-SPOSharingLink, Get-SPOSitePermission, Get-SPOStorageEntity

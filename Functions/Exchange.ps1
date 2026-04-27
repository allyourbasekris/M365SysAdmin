Import-Module ExchangeOnlineManagement -ErrorAction Stop

Function Get-EXOMailbox {
    [CmdletBinding()]
    param(
        [string]$Identity,
        [switch]$All
    )
    Test-M365Connected
    $params = @{}
    if ($Identity) { $params.Identity = $Identity }
    if ($All) { $params.All = $All }
    Get-EXOMailbox @params
}

Function Get-EXOMailboxPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )
    Test-M365Connected
    Get-EXOMailboxPermission -Identity $Identity
}

Function Get-EXORecipientPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )
    Test-M365Connected
    Get-EXORecipientPermission -Identity $Identity
}

Function Get-EXOCalendarFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )
    Test-M365Connected
    Get-EXOCalendarFolder -Identity $Identity
}

Function Get-EXOMailboxStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )
    Test-M365Connected
    Get-EXOMailboxStatistics -Identity $Identity
}

Function Test-EXOMailboxFolderPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity,
        [Parameter(Mandatory=$true)]
        [string]$Folder
    )
    Test-M365Connected
    Test-EXOMailboxFolderPermission -Identity $Identity -Folder $Folder
}

Function Get-EXOUser {
    [CmdletBinding()]
    param(
        [string]$Identity,
        [string]$Filter
    )
    Test-M365Connected
    $params = @{}
    if ($Identity) { $params.Identity = $Identity }
    if ($Filter) { $params.Filter = $Filter }
    Get-EXOUser @params
}

Export-ModuleMember -Function Get-EXOMailbox, Get-EXOMailboxPermission, Get-EXORecipientPermission, Get-EXOCalendarFolder, Get-EXOMailboxStatistics, Test-EXOMailboxFolderPermission, Get-EXOUser

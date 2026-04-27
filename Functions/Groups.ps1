# M365 Group Management Functions
# Requires Test-M365Connected (from module) to be available

function Get-M365Group {
    [CmdletBinding(DefaultParameterSetName='Filter')]
    param(
        [Parameter(ParameterSetName='ById')]
        [string]$Id,

        [Parameter(ParameterSetName='Filter')]
        [string]$DisplayName,

        [Parameter(ParameterSetName='Filter')]
        [switch]$Security,

        [Parameter(ParameterSetName='Filter')]
        [switch]$MailEnabled
    )

    Test-M365Connected

    if ($Id) {
        return Get-MgGroup -GroupId $Id
    }

    $filterClauses = @()
    if ($DisplayName) {
        $filterClauses += "DisplayName eq '$DisplayName'"
    }
    if ($Security.IsPresent) {
        $filterClauses += "SecurityEnabled eq true"
    }
    if ($MailEnabled.IsPresent) {
        $filterClauses += "MailEnabled eq true"
    }

    $filterString = $filterClauses -join ' and '
    if ($filterString) {
        Get-MgGroup -Filter $filterString
    } else {
        Get-MgGroup
    }
}

function New-M365Group {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DisplayName,

        [string]$Description,

        [switch]$MailEnabled,

        [switch]$Security,

        [string[]]$Members
    )

    Test-M365Connected

    $groupParams = @{
        DisplayName = $DisplayName
        Description = $Description
        MailEnabled = $MailEnabled.IsPresent
        SecurityEnabled = $Security.IsPresent
    }

    if ($PSCmdlet.ShouldProcess($DisplayName, "Create new M365 group")) {
        $newGroup = New-MgGroup @groupParams

        if ($Members) {
            foreach ($memberId in $Members) {
                New-MgGroupMember -GroupId $newGroup.Id -DirectoryObjectId $memberId
            }
        }

        return $newGroup
    }
}

function Add-M365GroupMember {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupId,

        [Parameter(Mandatory=$true)]
        [string]$MemberId
    )

    Test-M365Connected

    if ($PSCmdlet.ShouldProcess("Member $MemberId to Group $GroupId", "Add member to M365 group")) {
        New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $MemberId
    }
}

function Remove-M365GroupMember {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupId,

        [Parameter(Mandatory=$true)]
        [string]$MemberId
    )

    Test-M365Connected

    if ($PSCmdlet.ShouldProcess("Member $MemberId from Group $GroupId", "Remove member from M365 group")) {
        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $MemberId
    }
}

function Get-M365GroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupId
    )

    Test-M365Connected

    Get-MgGroupMember -GroupId $GroupId
}

function Get-M365GroupOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupId
    )

    Test-M365Connected

    Get-MgGroupOwner -GroupId $GroupId
}

Export-ModuleMember -Function Get-M365Group, New-M365Group, Add-M365GroupMember, Remove-M365GroupMember, Get-M365GroupMember, Get-M365GroupOwner

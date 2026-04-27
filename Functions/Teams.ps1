function Get-M365Team {
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName="ById")]
        [string]$GroupId,

        [Parameter(ParameterSetName="All")]
        [switch]$All
    )

    Test-M365Connected

    try {
        if ($GroupId) {
            $team = Get-MgTeam -TeamId $GroupId -ErrorAction Stop
            return $team
        }
        else {
            $teams = Get-MgTeam -All -ErrorAction Stop
            return $teams
        }
    }
    catch {
        Write-Error "Failed to get team: $_"
    }
}

function New-M365Team {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DisplayName,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [string[]]$Members,

        [Parameter()]
        [string[]]$Owners
    )

    Test-M365Connected

    try {
        $params = @{
            DisplayName = $DisplayName
            Description = $Description
        }

        $team = New-MgTeam -BodyParameter $params -ErrorAction Stop

        if ($Owners) {
            foreach ($ownerId in $Owners) {
                $ownerParams = @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    Roles = @("owner")
                    UserId = $ownerId
                }
                New-MgTeamMember -TeamId $team.Id -BodyParameter $ownerParams -ErrorAction Stop
            }
        }

        if ($Members) {
            foreach ($memberId in $Members) {
                $memberParams = @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    Roles = @()
                    UserId = $memberId
                }
                New-MgTeamMember -TeamId $team.Id -BodyParameter $memberParams -ErrorAction Stop
            }
        }

        Write-Host "Team created successfully: $DisplayName" -ForegroundColor Green
        return $team
    }
    catch {
        Write-Error "Failed to create team: $_"
    }
}

function Get-M365TeamChannel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamId,

        [Parameter()]
        [string]$ChannelId
    )

    Test-M365Connected

    try {
        if ($ChannelId) {
            $channel = Get-MgTeamChannel -TeamId $TeamId -ChannelId $ChannelId -ErrorAction Stop
            return $channel
        }
        else {
            $channels = Get-MgTeamChannel -TeamId $TeamId -All -ErrorAction Stop
            return $channels
        }
    }
    catch {
        Write-Error "Failed to get team channel: $_"
    }
}

function New-M365TeamChannel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamId,

        [Parameter(Mandatory)]
        [string]$DisplayName,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [ValidateSet("Standard", "Private", "Shared")]
        [string]$MembershipType = "Standard"
    )

    Test-M365Connected

    try {
        $params = @{
            DisplayName = $DisplayName
            Description = $Description
        }

        if ($MembershipType -ne "Standard") {
            $params.MembershipType = $MembershipType
        }

        $channel = New-MgTeamChannel -TeamId $TeamId -BodyParameter $params -ErrorAction Stop
        Write-Host "Channel created successfully: $DisplayName" -ForegroundColor Green
        return $channel
    }
    catch {
        Write-Error "Failed to create channel: $_"
    }
}

function Get-M365TeamMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamId,

        [Parameter()]
        [string]$MemberId
    )

    Test-M365Connected

    try {
        if ($MemberId) {
            $member = Get-MgTeamMember -TeamId $TeamId -ConversationMemberId $MemberId -ErrorAction Stop
            return $member
        }
        else {
            $members = Get-MgTeamMember -TeamId $TeamId -All -ErrorAction Stop
            return $members
        }
    }
    catch {
        Write-Error "Failed to get team member: $_"
    }
}

function Add-M365TeamMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamId,

        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter()]
        [ValidateSet("member", "owner")]
        [string]$Role = "member"
    )

    Test-M365Connected

    try {
        $roles = @()
        if ($Role -eq "owner") {
            $roles = @("owner")
        }

        $params = @{
            "@odata.type" = "#microsoft.graph.aadUserConversationMember"
            Roles = $roles
            UserId = $UserId
        }

        $member = New-MgTeamMember -TeamId $TeamId -BodyParameter $params -ErrorAction Stop
        Write-Host "Member added successfully to team: $UserId" -ForegroundColor Green
        return $member
    }
    catch {
        Write-Error "Failed to add team member: $_"
    }
}

function Remove-M365TeamMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamId,

        [Parameter(Mandatory)]
        [string]$MemberId
    )

    Test-M365Connected

    try {
        Remove-MgTeamMemberByRef -TeamId $TeamId -ConversationMemberId $MemberId -ErrorAction Stop
        Write-Host "Member removed successfully from team: $MemberId" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to remove team member: $_"
    }
}

function Get-M365TeamGuestSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamId
    )

    Test-M365Connected

    try {
        $team = Get-MgTeam -TeamId $TeamId -Property "GuestSettings" -ErrorAction Stop
        return $team.GuestSettings
    }
    catch {
        Write-Error "Failed to get team guest settings: $_"
    }
}

Export-ModuleMember -Function Get-M365Team, New-M365Team, Get-M365TeamChannel, New-M365TeamChannel, Get-M365TeamMember, Add-M365TeamMember, Remove-M365TeamMember, Get-M365TeamGuestSettings

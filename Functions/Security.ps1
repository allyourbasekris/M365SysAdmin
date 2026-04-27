function Get-M365LicenseUsage {
    [CmdletBinding()]
    param()
    
    Test-M365Connected
    
    $subs = Get-MgSubscribedSku -All
    $results = @()
    
    foreach ($sub in $subs) {
        $results += [PSCustomObject]@{
            SkuId = $sub.SkuId
            SkuPartNumber = $sub.SkuPartNumber
            ConsumedUnits = $sub.ConsumedUnits
            TotalUnits = $sub.TotalUnits
            Available = $sub.TotalUnits - $sub.ConsumedUnits
        }
    }
    
    return $results
}

function Get-M365InactiveUser {
    [CmdletBinding()]
    param(
        [Parameter()]
        [int]$Days = 90
    )
    
    Test-M365Connected
    
    $date = (Get-Date).AddDays(-$Days)
    $users = Get-MgUser -All -Property "id","displayName","userPrincipalName","lastSignInDateTime","accountEnabled"
    
    $inactive = $users | Where-Object { 
        $_.LastSignInDateTime -and ([DateTime]$_.LastSignInDateTime -lt $date) -and $_.AccountEnabled
    }
    
    return $inactive | Select-Object Id, DisplayName, UserPrincipalName, LastSignInDateTime
}

function Get-M365MFADisabledUser {
    [CmdletBinding()]
    param()
    
    Test-M365Connected
    
    $users = Get-MgUser -All -Property "id","displayName","userPrincipalName","authentication"
    
    $noMFA = @()
    foreach ($user in $users) {
        $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction SilentlyContinue
        if (-not $authMethods) {
            $noMFA += $user
        }
    }
    
    return $noMFA | Select-Object Id, DisplayName, UserPrincipalName
}

function Get-M365MailboxSizeReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [int]$Top = 50
    )
    
    Test-M365Connected
    
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
    
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Property "UserPrincipalName","ProhibitSendQuota","ProhibitSendReceiveQuota","TotalItemSize","MailboxProvisioningConstraint"
    
    $sizes = @()
    foreach ($mbx in $mailboxes) {
        $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue
        if ($stats) {
            $sizes += [PSCustomObject]@{
                User = $mbx.UserPrincipalName
                TotalItemSizeMB = [math]::Round($stats.TotalItemSize.Value.ToBytes() / 1MB, 2)
                ItemCount = $stats.ItemCount
            }
        }
    }
    
    return $sizes | Sort-Object TotalItemSizeMB -Descending | Select-Object -First $Top
}

function Get-M365GroupMemberReport {
    [CmdletBinding()]
    param()
    
    Test-M365Connected
    
    $groups = Get-MgGroup -All -Property "id","displayName","groupTypes"
    
    $report = @()
    foreach ($group in $groups) {
        $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
        $report += [PSCustomObject]@{
            GroupName = $group.DisplayName
            MemberCount = $members.Count
            GroupType = if ($group.GroupTypes -contains "Unified") { "Microsoft 365" } elseif ($group.SecurityEnabled) { "Security" } else { "Distribution" }
        }
    }
    
    return $report | Sort-Object MemberCount -Descending
}

function Get-M365GuestUserReport {
    [CmdletBinding()]
    param()
    
    Test-M365Connected
    
    $users = Get-MgUser -All -Filter "userType eq 'Guest'" -Property "id","displayName","userPrincipalName","mail","createdDateTime"
    
    return $users | Select-Object Id, DisplayName, UserPrincipalName, Mail, CreatedDateTime
}

function Get-M365PasswordExpiryReport {
    [CmdletBinding()]
    param()
    
    Test-M365Connected
    
    $users = Get-MgUser -All -Property "id","displayName","userPrincipalName","passwordProfile"
    
    $report = @()
    foreach ($user in $users) {
        if ($user.PasswordProfile -and $user.PasswordProfile.PasswordLastPasswordChangeTimestamp) {
            $lastChanged = [DateTime]$user.PasswordProfile.PasswordLastChangeTimestamp
            $daysSince = ((Get-Date) - $lastChanged).Days
            $report += [PSCustomObject]@{
                User = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                LastPasswordChange = $lastChanged
                DaysSinceChange = $daysSince
            }
        }
    }
    
    return $report | Sort-Object DaysSinceChange -Descending
}

Export-ModuleMember -Function Get-M365LicenseUsage, Get-M365InactiveUser, Get-M365MFADisabledUser, Get-M365MailboxSizeReport, Get-M365GroupMemberReport, Get-M365GuestUserReport, Get-M365PasswordExpiryReport
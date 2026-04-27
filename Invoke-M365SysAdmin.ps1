#!/usr/bin/env pwsh
<#
.SYNOPSIS
    M365 SysAdmin Tool - Entry point

.DESCRIPTION
    Unified management console for Microsoft 365 workloads
#>

$ErrorActionPreference = 'Stop'
$ModuleRoot = $PSScriptRoot

Get-ChildItem "$ModuleRoot/Functions/*.ps1" | ForEach-Object {
    Import-Module $_.FullName -Force
}

function Show-Menu {
    param([string]$Title, [array]$Items)
    
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) {
        Write-Host "  [$($i + 1)] $($Items[$i].Title)" -ForegroundColor White
    }
    Write-Host "  [Q] Quit" -ForegroundColor Yellow
    Write-Host ""
    
    $choice = Read-Host "Select option"
    
    if ($choice -eq "Q" -or $choice -eq "q") { return $null }
    if ($choice -match '^\d+$' -and $choice -ge 1 -and $choice -le $Items.Count) {
        return $Items[$choice - 1]
    }
    
    Write-Host "Invalid selection" -ForegroundColor Red
    return Show-Menu -Title $Title -Items $Items
}

function Show-MainMenu {
    $items = @(
        @{ Title = "Connect / Disconnect" }
        @{ Title = "User Management" }
        @{ Title = "Group Management" }
        @{ Title = "Exchange Online" }
        @{ Title = "Teams" }
        @{ Title = "SharePoint / OneDrive" }
        @{ Title = "Reports" }
    )
    
    while ($true) {
        Write-Host ""
        Write-Host "╔═══════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║     M365 SysAdmin Tool         ║" -ForegroundColor Cyan
        Write-Host "╚═══════════════════════════════╝" -ForegroundColor Cyan
        
        $status = Get-M365ConnectionStatus
        if ($status.Connected) {
            Write-Host "Connected: $($status.Account) ($($status.TenantId))" -ForegroundColor Green
        } else {
            Write-Host "Not connected" -ForegroundColor Yellow
        }
        
        $result = Show-Menu -Title "Main Menu" -Items $items
        if ($null -eq $result) { break }
        
        switch ($result.Title) {
            "Connect / Disconnect" { Show-ConnectionMenu }
            "User Management" { Show-UserMenu }
            "Group Management" { Show-GroupMenu }
            "Exchange Online" { Show-ExchangeMenu }
            "Teams" { Show-TeamsMenu }
            "SharePoint / OneDrive" { Show-SharePointMenu }
            "Reports" { Show-ReportsMenu }
        }
    }
}

function Show-ConnectionMenu {
    $status = Get-M365ConnectionStatus
    
    if ($status.Connected) {
        Write-Host "Tenant: $($status.TenantId)" -ForegroundColor White
        Write-Host "Account: $($status.Account)" -ForegroundColor White
        $confirm = Read-Host "Disconnect? (y/n)"
        if ($confirm -eq "y") {
            Disconnect-M365
        }
    } else {
        Connect-M365
    }
}

function Show-UserMenu {
    $items = @(
        @{ Title = "List Users" }
        @{ Title = "Create User" }
        @{ Title = "Reset Password" }
        @{ Title = "Disable User" }
        @{ Title = "Enable User" }
        @{ Title = "Assign License" }
    )
    
    while ($true) {
        $result = Show-Menu -Title "User Management" -Items $items
        if ($null -eq $result) { return }
        
        switch ($result.Title) {
            "List Users" { Get-M365User | Format-Table -AutoSize }
            "Create User" { 
                $name = Read-Host "Display Name"
                $upn = Read-Host "User Principal Name"
                $job = Read-Host "Job Title (optional)"
                if ($name -and $upn) {
                    New-M365User -DisplayName $name -UserPrincipalName $upn -JobTitle $job
                }
            }
            "Reset Password" {
                $upn = Read-Host "User Principal Name"
                if ($upn) { Set-M365UserPassword -UserId $upn }
            }
            "Disable User" {
                $upn = Read-Host "User Principal Name to disable"
                if ($upn) { Disable-M365User -UserId $upn }
            }
            "Enable User" {
                $upn = Read-Host "User Principal Name to enable"
                if ($upn) { Enable-M365User -UserId $upn }
            }
            "Assign License" {
                $upn = Read-Host "User Principal Name"
                $sku = Read-Host "License SKU ID"
                if ($upn -and $sku) { Set-M365UserLicense -UserId $upn -LicenseSkuId $sku }
            }
        }
    }
}

function Show-GroupMenu {
    $items = @(
        @{ Title = "List Groups" }
        @{ Title = "Create Group" }
        @{ Title = "View Group Members" }
        @{ Title = "Add Group Member" }
    )
    
    while ($true) {
        $result = Show-Menu -Title "Group Management" -Items $items
        if ($null -eq $result) { return }
        
        switch ($result.Title) {
            "List Groups" { Get-M365Group | Format-Table -AutoSize }
            "Create Group" {
                $name = Read-Host "Group Name"
                $desc = Read-Host "Description (optional)"
                if ($name) { New-M365Group -DisplayName $name -Description $desc }
            }
            "View Group Members" {
                $name = Read-Host "Group Display Name"
                $group = Get-M365Group -DisplayName $name
                if ($group) { Get-M365GroupMember -GroupId $group.Id | Format-Table -AutoSize }
            }
            "Add Group Member" {
                $gname = Read-Host "Group Name"
                $member = Read-Host "Member ID"
                $group = Get-M365Group -DisplayName $gname
                if ($group -and $member) { Add-M365GroupMember -GroupId $group.Id -MemberId $member }
            }
        }
    }
}

function Show-ExchangeMenu {
    $items = @(
        @{ Title = "List Mailboxes" }
        @{ Title = "Mailbox Permissions" }
        @{ Title = "Mailbox Statistics" }
    )
    
    while ($true) {
        $result = Show-Menu -Title "Exchange Online" -Items $items
        if ($null -eq $result) { return }
        
        switch ($result.Title) {
            "List Mailboxes" { Get-EXOMailbox | Format-Table -AutoSize }
            "Mailbox Permissions" {
                $upn = Read-Host "User Principal Name"
                if ($upn) { Get-EXOMailboxPermission -Identity $upn | Format-Table -AutoSize }
            }
            "Mailbox Statistics" {
                $upn = Read-Host "User Principal Name"
                if ($upn) { Get-EXOMailboxStatistics -Identity $upn }
            }
        }
    }
}

function Show-TeamsMenu {
    $items = @(
        @{ Title = "List Teams" }
        @{ Title = "Create Team" }
        @{ Title = "View Team Members" }
    )
    
    while ($true) {
        $result = Show-Menu -Title "Teams" -Items $items
        if ($null -eq $result) { return }
        
        switch ($result.Title) {
            "List Teams" { Get-M365Team | Format-Table -AutoSize }
            "Create Team" {
                $name = Read-Host "Team Name"
                $desc = Read-Host "Description (optional)"
                if ($name) { New-M365Team -DisplayName $name -Description $desc }
            }
            "View Team Members" {
                $id = Read-Host "Team ID"
                if ($id) { Get-M365TeamMember -TeamId $id | Format-Table -AutoSize }
            }
        }
    }
}

function Show-SharePointMenu {
    $items = @(
        @{ Title = "List Sites" }
        @{ Title = "Site Permissions" }
        @{ Title = "User OneDrive" }
    )
    
    while ($true) {
        $result = Show-Menu -Title "SharePoint / OneDrive" -Items $items
        if ($null -eq $result) { return }
        
        switch ($result.Title) {
            "List Sites" { Get-SPOSite | Format-Table -AutoSize }
            "Site Permissions" {
                $site = Read-Host "Site URL"
                if ($site) { Get-SPOSitePermission -Site $site | Format-Table -AutoSize }
            }
            "User OneDrive" {
                $uid = Read-Host "User ID"
                if ($uid) { Get-SPOUserOneDrive -UserId $uid }
            }
        }
    }
}

function Show-ReportsMenu {
    $items = @(
        @{ Title = "License Usage" }
        @{ Title = "Inactive Users" }
        @{ Title = "MFA Disabled Users" }
        @{ Title = "Mailbox Sizes" }
        @{ Title = "Group Members" }
        @{ Title = "Guest Users" }
    )
    
    while ($true) {
        $result = Show-Menu -Title "Reports" -Items $items
        if ($null -eq $result) { return }
        
        switch ($result.Title) {
            "License Usage" { Get-M365LicenseUsage | Format-Table -AutoSize }
            "Inactive Users" { Get-M365InactiveUser | Format-Table -AutoSize }
            "MFA Disabled Users" { Get-M365MFADisabledUser | Format-Table -AutoSize }
            "Mailbox Sizes" { Get-M365MailboxSizeReport | Format-Table -AutoSize }
            "Group Members" { Get-M365GroupMemberReport | Format-Table -AutoSize }
            "Guest Users" { Get-M365GuestUserReport | Format-Table -AutoSize }
        }
    }
}

function Invoke-M365TUI {
    Show-MainMenu
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-M365TUI
}

Export-ModuleMember -Function Invoke-M365TUI
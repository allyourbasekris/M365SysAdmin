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

if (-not (Get-Module -ListAvailable -Name Terminal.Gui)) {
    Write-Host "Installing Terminal.Gui..." -ForegroundColor Yellow
    Install-Module -Name Terminal.Gui -Scope CurrentUser -AllowClobber -Force
}

Import-Module Terminal.Gui -MinimumVersion 2.0

function Show-MainMenu {
    param([string]$Workload = "Main")
    
    $menu = @{
        Main = @(
            @{ Title = "Connect"; Action = { Show-ConnectionMenu } }
            @{ Title = "User Management"; Action = { Show-UserMenu } }
            @{ Title = "Group Management"; Action = { Show-GroupMenu } }
            @{ Title = "Exchange Online"; Action = { Show-ExchangeMenu } }
            @{ Title = "Teams"; Action = { Show-TeamsMenu } }
            @{ Title = "SharePoint/OneDrive"; Action = { Show-SharePointMenu } }
            @{ Title = "Reports"; Action = { Show-ReportsMenu } }
            @{ Title = "Quit"; Action = { [Terminal.Gui.Application]::RequestStop() }
        )
    }
    
    $result = [Terminal.Gui.MenuDialog]::Query("M365 SysAdmin Tool", "Select workload:", $menu[$Workload])
    return $result
}

function Show-ConnectionMenu {
    $status = Get-M365ConnectionStatus
    
    if ($status.Connected) {
        [Terminal.Gui.MessageBox]::Query(
            "Connected",
            "Tenant: $($status.TenantId)`nAccount: $($status.Account)`n`nDisconnect?",
            "Yes" , "No")
        Disconnect-M365
    }
    else {
        Connect-M365
    }
}

function Show-UserMenu {
    $items = @(
        @{ Title = "List Users"; Action = { Get-M365User | Out-Host } }
        @{ Title = "Create User"; Action = { Show-CreateUserDialog } }
        @{ Title = "Reset Password"; Action = { Show-ResetPasswordDialog } }
        @{ Title = "Disable User"; Action = { Show-DisableUserDialog } }
        @{ Title = "Assign License"; Action = { Show-AssignLicenseDialog } }
        @{ Title = "Back"; Action = { return } }
    )
    
    [Terminal.Gui.MenuDialog]::Query("User Management", "Select action:", $items)
}

function Show-CreateUserDialog {
    $name = [Terminal.Gui.Prompt]::Prompt("Enter display name:")
    if (-not $name) { return }
    
    $upn = [Terminal.Gui.Prompt]::Prompt("Enter user principal name (e.g., user@domain.onmicrosoft.com):")
    if (-not $upn) { return }
    
    $jobTitle = [Terminal.Gui.Prompt]::Prompt("Enter job title (optional):")
    
    try {
        New-M365User -DisplayName $name -UserPrincipalName $upn -JobTitle $jobTitle
        [Terminal.Gui.MessageBox]::Query("Success", "User created successfully!", "OK")
    }
    catch {
        [Terminal.Gui.MessageBox]::Query("Error", "Failed to create user: $_", "OK")
    }
}

function Show-ResetPasswordDialog {
    $upn = [Terminal.Gui.Prompt]::Prompt("Enter user principal name:")
    if (-not $upn) { return }
    
    try {
        $password = Set-M365UserPassword -UserId $upn
        [Terminal.Gui.MessageBox]::Query("Success", "New password: $password", "OK")
    }
    catch {
        [Terminal.Gui.MessageBox]::Query("Error", "Failed to reset password: $_", "OK")
    }
}

function Show-DisableUserDialog {
    $upn = [Terminal.Gui.Prompt]::Prompt("Enter user principal name to disable:")
    if (-not $upn) { return }
    
    $confirm = [Terminal.Gui.MessageBox]::Query(
        "Confirm",
        "Disable user $upn?",
        "Yes" , "No")
    
    if ($confirm -eq "Yes") {
        Disable-M365User -UserId $upn
        [Terminal.Gui.MessageBox]::Query("Success", "User disabled", "OK")
    }
}

function Show-AssignLicenseDialog {
    $upn = [Terminal.Gui.Prompt]::Prompt("Enter user principal name:")
    if (-not $upn) { return }
    
    $sku = [Terminal.Gui.Prompt]::Prompt("Enter license SKU ID (e.g., a1b2c3d4-...):")
    if (-not $sku) { return }
    
    Set-M365UserLicense -UserId $upn -LicenseSkuId $sku
    [Terminal.Gui.MessageBox]::Query("Success", "License assigned", "OK")
}

function Show-GroupMenu {
    $items = @(
        @{ Title = "List Groups"; Action = { Get-M365Group | Out-Host } }
        @{ Title = "Create Group"; Action = { Show-CreateGroupDialog } }
        @{ Title = "Manage Members"; Action = { Show-GroupMembersDialog } }
        @{ Title = "Back"; Action = { return } }
    )
    
    [Terminal.Gui.MenuDialog]::Query("Group Management", "Select action:", $items)
}

function Show-CreateGroupDialog {
    $name = [Terminal.Gui.Prompt]::Prompt("Enter group display name:")
    if (-not $name) { return }
    
    $desc = [Terminal.Gui.Prompt]::Prompt("Enter description (optional):")
    
    try {
        New-M365Group -DisplayName $name -Description $desc
        [Terminal.Gui.MessageBox]::Query("Success", "Group created successfully!", "OK")
    }
    catch {
        [Terminal.Gui.MessageBox]::Query("Error", "Failed to create group: $_", "OK")
    }
}

function Show-GroupMembersDialog {
    $groupName = [Terminal.Gui.Prompt]::Prompt("Enter group name:")
    if (-not $groupName) { return }
    
    $group = Get-M365Group -DisplayName $groupName
    if ($group) {
        Get-M365GroupMember -GroupId $group.Id | Out-Host
    }
}

function Show-ExchangeMenu {
    $items = @(
        @{ Title = "List Mailboxes"; Action = { Get-EXOMailbox -All | Out-Host } }
        @{ Title = "Mailbox Permissions"; Action = { Show-MailboxPermissionDialog } }
        @{ Title = "Mailbox Statistics"; Action = { Show-MailboxStatsDialog } }
        @{ Title = "Recipients"; Action = { Get-EXOUser | Out-Host } }
        @{ Title = "Back"; Action = { return } }
    )
    
    [Terminal.Gui.MenuDialog]::Query("Exchange Online", "Select action:", $items)
}

function Show-MailboxPermissionDialog {
    $upn = [Terminal.Gui.Prompt]::Prompt("Enter user principal name:")
    if (-not $upn) { return }
    
    Get-EXOMailboxPermission -Identity $upn | Out-Host
}

function Show-MailboxStatsDialog {
    $upn = [Terminal.Gui.Prompt]::Prompt("Enter user principal name:")
    if (-not $upn) { return }
    
    Get-EXOMailboxStatistics -Identity $upn | Out-Host
}

function Show-TeamsMenu {
    $items = @(
        @{ Title = "List Teams"; Action = { Get-M365Team | Out-Host } }
        @{ Title = "Create Team"; Action = { Show-CreateTeamDialog } }
        @{ Title = "Manage Channels"; Action = { Show-ChannelDialog } }
        @{ Title = "Team Members"; Action = { Show-TeamMembersDialog } }
        @{ Title = "Back"; Action = { return } }
    )
    
    [Terminal.Gui.MenuDialog]::Query("Teams", "Select action:", $items)
}

function Show-CreateTeamDialog {
    $name = [Terminal.Gui.Prompt]::Prompt("Enter team display name:")
    if (-not $name) { return }
    
    $desc = [Terminal.Gui.Prompt]::Prompt("Enter description (optional):")
    
    try {
        New-M365Team -DisplayName $name -Description $desc
        [Terminal.Gui.MessageBox]::Query("Success", "Team created!", "OK")
    }
    catch {
        [Terminal.Gui.MessageBox]::Query("Error", "Failed: $_", "OK")
    }
}

function Show-ChannelDialog {
    $teamId = [Terminal.Gui.Prompt]::Prompt("Enter team ID:")
    if (-not $teamId) { return }
    
    Get-M365TeamChannel -TeamId $teamId | Out-Host
}

function Show-TeamMembersDialog {
    $teamId = [Terminal.Gui.Prompt]::Prompt("Enter team ID:")
    if (-not $teamId) { return }
    
    Get-M365TeamMember -TeamId $teamId | Out-Host
}

function Show-SharePointMenu {
    $items = @(
        @{ Title = "List Sites"; Action = { Get-SPOSite | Out-Host } }
        @{ Title = "Site Permissions"; Action = { Show-SitePermissionDialog } }
        @{ Title = "User OneDrive"; Action = { Show-OneDriveDialog } }
        @{ Title = "Back"; Action = { return } }
    )
    
    [Terminal.Gui.MenuDialog]::Query("SharePoint/OneDrive", "Select action:", $items)
}

function Show-SitePermissionDialog {
    $site = [Terminal.Gui.Prompt]::Prompt("Enter site URL:")
    if (-not $site) { return }
    
    Get-SPOSitePermission -Site $site | Out-Host
}

function Show-OneDriveDialog {
    $user = [Terminal.Gui.Prompt]::Prompt("Enter user ID:")
    if (-not $user) { return }
    
    Get-SPOUserOneDrive -UserId $user | Out-Host
}

function Show-ReportsMenu {
    $items = @(
        @{ Title = "License Usage"; Action = { Get-M365LicenseUsage | Out-Host } }
        @{ Title = "Inactive Users"; Action = { Get-M365InactiveUser | Out-Host } }
        @{ Title = "MFA Disabled"; Action = { Get-M365MFADisabledUser | Out-Host } }
        @{ Title = "Mailbox Sizes"; Action = { Get-M365MailboxSizeReport | Out-Host } }
        @{ Title = "Group Members"; Action = { Get-M365GroupMemberReport | Out-Host } }
        @{ Title = "Guest Users"; Action = { Get-M365GuestUserReport | Out-Host } }
        @{ Title = "Password Expiry"; Action = { Get-M365PasswordExpiryReport | Out-Host } }
        @{ Title = "Back"; Action = { return } }
    )
    
    [Terminal.Gui.MenuDialog]::Query("Reports", "Select report:", $items)
}

function Invoke-M365TUI {
    [CmdletBinding()]
    param()
    
    [Terminal.Gui.Application]::Init()
    
    $window = [Terminal.Gui.Window]::new()
    $window.Title = "M365 SysAdmin Tool"
    $window.Height = 20
    $window.Width = 50
    
    $label = [Terminal.Gui.Label]::new()
    $label.Text = "Loading..."
    $label.Y = 1
    $label.X = [Terminal.Gui.Pos]::Center()
    
    $window.Add($label)
    
    [Terminal.Gui.Application]::Run($window)
    
    while ($true) {
        $result = Show-MainMenu
        if (-not $result) { break }
    }
    
    [Terminal.Gui.Application]::Shutdown()
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-M365TUI
}

Export-ModuleMember -Function Invoke-M365TUI
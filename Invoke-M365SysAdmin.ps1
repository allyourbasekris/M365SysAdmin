#!/usr/bin/env pwsh
<#
.SYNOPSIS
    M365 SysAdmin Tool - Entry point

.DESCRIPTION
    Unified management console for Microsoft 365 workloads
#>

$ErrorActionPreference = 'Stop'

$ModuleRoot = $PSScriptRoot

Import-Module "$ModuleRoot/Functions/Auth.ps1" -Force
Import-Module "$ModuleRoot/Functions/Utilities.ps1" -Force

if (-not (Get-Module -ListAvailable -Name Terminal.Gui)) {
    Write-Warning "Terminal.Gui module not found. Installing..."
    Install-Module -Name Terminal.Gui -Scope CurrentUser -AllowClobber
}

Import-Module Terminal.Gui -MinimumVersion 2.0

function Invoke-M365TUI {
    [CmdletBinding()]
    param()

    if (-not (Get-M365ConnectionStatus)) {
        $result = [Terminal.Gui.MessageBox]::Query(
            "Not Connected",
            "Connect to Microsoft 365 before continuing?",
            "Yes" , "No")
        if ($result -eq "Yes") {
            Connect-M365
        }
    }

    [Terminal.Gui.Application]::Init()
    $topLevel = [Terminal.Gui.Window]::new()
    $topLevel.Title = "M365 SysAdmin Tool"

    $menu = [Terminal.Gui.MenuBar]::new()
    $menu.Menus = @(
        [Terminal.Gui.MenuBarItem]::new("_File", @(
            [Terminal.Gui.MenuItem]::new("_Connect", "Connect-M365", { Connect-M365 }),
            [Terminal.Gui.MenuItem]::new("_Disconnect", "Disconnect-M365", { Disconnect-M365 }),
            [Terminal.Gui.MenuItem]::new("E_xit", "", { [Terminal.Gui.Application]::RequestStop() })
        )),
        [Terminal.Gui.MenuBarItem]::new("_Workloads", @(
            [Terminal.Gui.MenuItem]::new("_Users", "", { $script:CurrentWorkload = 'Users' }),
            [Terminal.Gui.MenuItem]::new("_Groups", "", { $script:CurrentWorkload = 'Groups' }),
            [Terminal.Gui.MenuItem]::new("_Exchange", "", { $script:CurrentWorkload = 'Exchange' }),
            [Terminal.Gui.MenuBarItem]::new("T_eams", "", { $script:CurrentWorkload = 'Teams' }),
            [Terminal.Gui.MenuItem]::new("S_harePoint", "", { $script:CurrentWorkload = 'SharePoint' }),
            [Terminal.Gui.MenuItem]::new("_Reports", "", { $script:CurrentWorkload = 'Reports' })
        ))
    )

    $topLevel.Add($menu)
    [Terminal.Gui.Application]::Run($topLevel)
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-M365TUI
}

Export-ModuleMember -Function Invoke-M365TUI

function Connect-M365 {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TenantId,
        
        [Parameter()]
        [string]$Username
    )
    
    Write-Host "Connecting to Microsoft 365..." -ForegroundColor Cyan
    
    if (-not $TenantId) {
        $TenantId = Read-Host "Enter Tenant ID (or press Enter for common)"
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            $TenantId = "common"
        }
    }
    
    $scopes = @(
        "User.Read.All"
        "Group.Read.All"
        "Directory.Read.All"
        "Mail.Read"
        "Mail.Send"
        "Calendars.Read"
        "Team.ReadBasic.All"
        "Sites.Read.All"
    )
    
    try {
        $connectParams = @{
            Scopes = $scopes
            ContextScope = "Process"
        }
        
        if ($TenantId -ne "common") {
            $connectParams.TenantId = $TenantId
        }
        
        Connect-MgGraph @connectParams -Device
        Write-Host "Connected to Microsoft 365 successfully!" -ForegroundColor Green
        
        $script:IsConnected = $true
    }
    catch {
        Write-Error "Failed to connect: $_"
        $script:IsConnected = $false
    }
}

function Disconnect-M365 {
    [CmdletBinding()]
    param()
    
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Microsoft 365" -ForegroundColor Yellow
        $script:IsConnected = $false
    }
    catch {
        Write-Warning "Error disconnecting: $_"
    }
}

function Get-M365ConnectionStatus {
    [CmdletBinding()]
    param()
    
    try {
        $context = Get-MgContext
        if ($context) {
            return @{
                Connected = $true
                TenantId = $context.TenantId
                Account = $context.Account
                Scopes = $context.Scopes
            }
        }
    }
    catch {}
    
    return @{
        Connected = $false
        TenantId = $null
        Account = $null
        Scopes = $null
    }
}

function Test-M365Connected {
    [CmdletBinding()]
    param()
    
    $status = Get-M365ConnectionStatus
    if (-not $status.Connected) {
        throw "Not connected to Microsoft 365. Run Connect-M365 first."
    }
}

Export-ModuleMember -Function Connect-M365, Disconnect-M365, Get-M365ConnectionStatus, Test-M365Connected
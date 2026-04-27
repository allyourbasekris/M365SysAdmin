# M365SysAdmin Utilities Module
# Helper functions for logging, Graph API requests, and output formatting

function Write-M365Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $logPath = Join-Path -Path $env:TEMP -ChildPath 'M365SysAdmin.log'
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "$timestamp [$Level] $Message"

    try {
        Add-Content -Path $logPath -Value $logEntry -ErrorAction Stop
    } catch {
        Write-Error "Failed to write to log file: $_"
    }
}

function Get-M365Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$Last = 10
    )

    $logPath = Join-Path -Path $env:TEMP -ChildPath 'M365SysAdmin.log'

    if (-not (Test-Path -Path $logPath)) {
        Write-Warning "Log file not found at $logPath"
        return @()
    }

    $logEntries = Get-Content -Path $logPath
    return $logEntries | Select-Object -Last $Last
}

function Get-M365GraphUri {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Resource,

        [Parameter(Mandatory = $false)]
        [hashtable]$QueryParameters = @{}
    )

    $baseUri = 'https://graph.microsoft.com/v1.0'
    $uri = "$baseUri/$Resource"

    if ($QueryParameters.Count -gt 0) {
        $queryParts = @()
        foreach ($key in $QueryParameters.Keys) {
            $queryParts += "$key=$($QueryParameters[$key])"
        }
        $uri += '?' + ($queryParts -join '&')
    }

    return $uri
}

function Invoke-M365GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [string]$Method = 'GET',

        [Parameter(Mandatory = $false)]
        [object]$Body,

        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{},

        [Parameter(Mandatory = $false)]
        [string]$ContentType = 'application/json'
    )

    if (-not $script:AuthToken) {
        Write-Error "Not connected to Microsoft 365. Use Connect-M365SysAdmin first."
        return
    }

    $Headers['Authorization'] = "Bearer $script:AuthToken"

    $invokeParams = @{
        Uri         = $Uri
        Method      = $Method
        Headers     = $Headers
        ContentType = $ContentType
    }

    if ($Body) {
        $invokeParams.Body = $Body | ConvertTo-Json -Depth 10 -Compress
    }

    try {
        Invoke-RestMethod @invokeParams
    } catch {
        Write-M365Log -Message "Graph API request failed: $_" -Level Error
        throw $_
    }
}

function Format-M365Table {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object[]]$InputObject,

        [Parameter(Mandatory = $false)]
        [string[]]$Property
    )

    begin {
        $collectedObjects = @()
    }

    process {
        $collectedObjects += $InputObject
    }

    end {
        if ($Property) {
            $collectedObjects | Format-Table -Property $Property
        } else {
            $collectedObjects | Format-Table
        }
    }
}

function Write-M365Output {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object[]]$InputObject,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Table', 'List', 'Json')]
        [string]$Format = 'Table'
    )

    begin {
        $collectedObjects = @()
    }

    process {
        $collectedObjects += $InputObject
    }

    end {
        switch ($Format) {
            'Table' { $collectedObjects | Format-Table }
            'List' { $collectedObjects | Format-List }
            'Json' { $collectedObjects | ConvertTo-Json -Depth 10 }
        }
    }
}

# Export module members
Export-ModuleMember -Function Write-M365Log, Get-M365Log, Get-M365GraphUri, Invoke-M365GraphRequest, Format-M365Table, Write-M365Output

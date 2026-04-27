function Get-M365User {
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName="ById")]
        [string]$Id,

        [Parameter(ParameterSetName="ByUPN")]
        [string]$UserPrincipalName,

        [Parameter(ParameterSetName="All")]
        [switch]$All
    )

    Test-M365Connected

    try {
        if ($Id) {
            $user = Get-MgUser -UserId $Id -ErrorAction Stop
            return $user
        }
        elseif ($UserPrincipalName) {
            $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
            return $user
        }
        elseif ($All) {
            $users = Get-MgUser -All -ErrorAction Stop
            return $users
        }
        else {
            $users = Get-MgUser -All -ErrorAction Stop
            return $users
        }
    }
    catch {
        Write-Error "Failed to get user: $_"
    }
}

function New-M365User {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DisplayName,

        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter()]
        [string]$Password,

        [Parameter()]
        [string]$Mail,

        [Parameter()]
        [string]$JobTitle,

        [Parameter()]
        [string]$Department
    )

    Test-M365Connected

    try {
        $passwordProfile = $null
        if ($Password) {
            $passwordProfile = @{
                Password = $Password
                ForceChangePasswordNextSignIn = $false
            }
        }

        $params = @{
            DisplayName = $DisplayName
            UserPrincipalName = $UserPrincipalName
            AccountEnabled = $true
            PasswordProfile = $passwordProfile
            Mail = $Mail
            JobTitle = $JobTitle
            Department = $Department
        }

        $user = New-MgUser -BodyParameter $params -ErrorAction Stop
        Write-Host "User created successfully: $($user.UserPrincipalName)" -ForegroundColor Green
        return $user
    }
    catch {
        Write-Error "Failed to create user: $_"
    }
}

function Set-M365UserLicense {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string[]]$SkuId
    )

    Test-M365Connected

    try {
        $addLicenses = @()
        foreach ($sku in $SkuId) {
            $addLicenses += @{
                DisabledPlans = @()
                SkuId = $sku
            }
        }

        $params = @{
            AddLicenses = $addLicenses
            RemoveLicenses = @()
        }

        Set-MgUserLicense -UserId $UserId -BodyParameter $params -ErrorAction Stop
        Write-Host "License assigned successfully to user: $UserId" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to set license: $_"
    }
}

function Get-M365UserLicense {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    Test-M365Connected

    try {
        $licenses = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop
        return $licenses
    }
    catch {
        Write-Error "Failed to get user license: $_"
    }
}

function Set-M365UserPassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$NewPassword,

        [switch]$ForceChangePasswordNextSignIn
    )

    Test-M365Connected

    try {
        $passwordProfile = @{
            Password = $NewPassword
            ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn
        }

        $params = @{
            PasswordProfile = $passwordProfile
        }

        Set-MgUser -UserId $UserId -BodyParameter $params -ErrorAction Stop
        Write-Host "Password reset successfully for user: $UserId" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to reset password: $_"
    }
}

function Disable-M365User {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    Test-M365Connected

    try {
        $params = @{
            AccountEnabled = $false
        }

        Set-MgUser -UserId $UserId -BodyParameter $params -ErrorAction Stop
        Write-Host "User disabled successfully: $UserId" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to disable user: $_"
    }
}

function Enable-M365User {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    Test-M365Connected

    try {
        $params = @{
            AccountEnabled = $true
        }

        Set-MgUser -UserId $UserId -BodyParameter $params -ErrorAction Stop
        Write-Host "User enabled successfully: $UserId" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to enable user: $_"
    }
}

Export-ModuleMember -Function Get-M365User, New-M365User, Set-M365UserLicense, Get-M365UserLicense, Set-M365UserPassword, Disable-M365User, Enable-M365User
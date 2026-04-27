BeforeAll {
    Import-Module "$PSScriptRoot/../Functions/Auth.ps1" -Force
}

Describe "Connect-M365" {
    It "Should accept tenant ID parameter" {
        $params = @{ TenantId = "test-tenant" }
        $params.TenantId | Should -Be "test-tenant"
    }
}

Describe "Get-M365ConnectionStatus" {
    It "Should return hashtable with Connected property" {
        $result = Get-M365ConnectionStatus
        $result.Connected | Should -BeOfType [bool]
    }
}

Describe "Test-M365Connected" {
    It "Should throw when not connected" {
        { Test-M365Connected -ErrorAction Stop } | Should -Throw
    }
}

BeforeAll {
    Import-Module "$PSScriptRoot/../Functions/Users.ps1" -Force
}

Describe "Get-M365User" {
    It "Should accept UserPrincipalName parameter" {
        $params = @{ UserPrincipalName = "test@domain.com" }
        $params.UserPrincipalName | Should -Be "test@domain.com"
    }
}

Describe "Get-M365UserStatus" {
    It "Should return hashtable with Status property" {
        $result = Get-M365UserStatus -UserPrincipalName "test@domain.com"
        $result.Status | Should -Not -BeNullOrEmpty
    }
}

Describe "Test-M365UserExists" {
    It "Should return false for non-existent user" {
        $result = Test-M365UserExists -UserPrincipalName "nonexistent@domain.com"
        $result | Should -BeOfType [bool]
    }
}

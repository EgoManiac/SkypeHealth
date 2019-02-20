import-Module .\SkypeHealth.psm1

Describe "Check Doughnut" {
    Context "Does the Function Create a doughnut png" {
        Mock -ModuleName SkypeHealth Get-CBUserCounts {  return @{
                'AllUsers'   = 1000;
                'PoolGroups' = @{Name = "FEServer"; Count = 1000};
                'EVusers'    = 450;
                'ExUMusers'  = 100;
            }
        }
    
        It "Returns a long BASE64 string" {
            (Get-PoolDoughnut).length -gt 100 | Should Be $true
        }
    }
}

Describe "Check Image Base 64 Conversion" {
    Context "Does the Function Create a Base64 String" {
           
        It "Returns a long BASE64 string" {
            (Get-CbImageBase64 -url "https://upload.wikimedia.org/wikipedia/commons/8/86/Microsoft_Skype_for_Business_logo.png").length -gt 100 | Should Be $true
        }
    }
}
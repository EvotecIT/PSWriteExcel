Import-Module PSWinDocumentation.AD -Force
Import-Module PSWriteExcel

if ($null -eq $ADForest) {
    $ADForest = Get-WinADForestInformation -Verbose -PasswordQuality
}

Excel -FilePath $PSScriptRoot\"Run-Demo01.xlsx" {
    $Number = 0
    [int] $Color = 0
    WorkbookProperties -Title 'PSWinDocumentation - Active Directory Demo - Automated'
    foreach ($Key in $ADForest.Keys | Where-Object { $_ -notin 'FoundDomains', 'Domains', 'ForestName', 'ForestNameDN' }) {
        Worksheet -DataTable $ADForest.$Key -Name $Key -TabColor BlueViolet #-AutoFilter -AutoFit

    }
    foreach ($FoundDomains in $ADForest.FoundDomains) {
        foreach ($D in $ADForest.FoundDomains.Keys) {
            $Color++
            foreach ($Section in $ADForest.FoundDomains.$D.Keys) {
                $Name = "$Section - $D" -replace 'DomainPassword', '' -replace 'Domain', '' -replace 'Password', 'Pass'
                $Number++
                Worksheet -DataTable $ADForest.FoundDomains.$D.$Section -Name $Number #-TabColor BlueViolet #-AutoFilter -AutoFit -Verbose

            }
        }
    }
} -Verbose -Open #-Parallel

return

Excel -FilePath $PSScriptRoot\"Run-Demo02.xlsx" {
    WorkbookProperties -Title 'PSWinDocumentation - Active Directory Demo - Automated'
    Worksheet -DataTable $ADForest.ForestInformation -Name 'Forest Information' -TabColor Green -AutoFilter -AutoFit

}
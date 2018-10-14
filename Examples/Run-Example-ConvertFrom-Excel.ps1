#Clear-Host
Import-Module PSWriteExcel -Force

$FilePath = "$Env:USERPROFILE\Desktop\Test.xlsx"

ConvertFrom-Excel -FilePath $FilePath -ExcelWorkSheetName 'Arkusz1' | Format-Table -AutoSize
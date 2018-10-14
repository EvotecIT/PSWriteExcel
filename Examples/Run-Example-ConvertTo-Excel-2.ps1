
Import-Module PSWriteExcel -Force
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-Test.xlsx"

$Excel = New-ExcelDocument -Verbose
Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook

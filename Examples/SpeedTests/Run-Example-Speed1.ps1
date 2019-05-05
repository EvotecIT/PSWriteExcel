if ($null -eq $Process) {
    $Process = Get-Process
}

Import-Module .\PSWriteExcel.psd1 -Force #-Verbose

$FilePath = "$PSScriptRoot\PSWriteExcel-Example-Test2.xlsx"
Measure-Collection -Name 'ConvertTo-Excel' -ScriptBlock {
    $Process | ConvertTo-Excel -FilePath $FilePath -ExcelWorkSheetName 'Test'
}


$FilePath = "$PSScriptRoot\PSWriteExcel-Example-Test.xlsx"
Measure-Collection -Name 'Add-ExcelWorkSheet' -ScriptBlock {
    $Excel = New-ExcelDocument -Verbose
    Add-ExcelWorksheetData -ExcelDocument $Excel -Supress $True -DataTable $Process -ExcelWorksheetName 'Test'
    Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath #-OpenWorkBook
}
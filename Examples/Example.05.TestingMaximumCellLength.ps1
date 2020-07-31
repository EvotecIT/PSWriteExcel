Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force

$FilePath = "$PSScriptRoot\Output\PSWriteExcel-Example-TestingMaximumCellLength.xlsx"
$Excel = New-ExcelDocument -Verbose
$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
$Value = ([string]3276999).PadRight(32767999, '0')
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 1 -CellValue $Value
Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook

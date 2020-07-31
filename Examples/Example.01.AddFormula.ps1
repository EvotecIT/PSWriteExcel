Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force #-Verbose

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-Test.xlsx"

$Excel = New-ExcelDocument -Verbose
$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 10' -Suppress $False -Option 'Replace'

Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 5 -CellValue 5
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 2 -CellColumn 5 -CellValue 5
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 3 -CellColumn 5 -CellValue 5

Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 7 -CellColumn 5 -CellFormula 'SUM(E1:E5)'
# this is also supported but we will remove = for user
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 8 -CellColumn 5 -CellFormula '=SUM(E1:E5)'

Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 2 -CellColumn 2 -CellFormula "HYPERLINK(B1)";
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 2 -CellValue "https://evotec.xyz";

# this forces Excel to calculate formulas in the worksheet
Request-ExcelWorkSheetCalculation -ExcelWorksheet $ExcelWorkSheet

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
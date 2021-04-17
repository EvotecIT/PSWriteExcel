Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force #-Verbose

$FilePath = "$PSScriptRoot\Output\PSWriteExcel.FormulaTest.xlsx"

$Excel = New-ExcelDocument -Verbose
$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 10' -Suppress $False -Option 'Replace'

Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 5 -CellValue 5
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 2 -CellColumn 5 -CellValue 5
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 3 -CellColumn 5 -CellValue 5

Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 7 -CellColumn 5 -CellFormula 'SUM(E1:E5)'
# this is also supported but we will remove = for user
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 8 -CellColumn 5 -CellFormula '=SUM(E1:E5)*10+22'

# this forces Excel to calculate formulas in the worksheet
Request-ExcelWorkSheetCalculation -ExcelWorksheet $ExcelWorkSheet

# This gets data from it, to confirm recalculation worked
Get-ExcelWorkSheetCell -CellRow 8 -CellColumn 5 -ExcelWorksheet $ExcelWorkSheet

$ExcelWorksheet.Cells[8, 5] | fl *

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
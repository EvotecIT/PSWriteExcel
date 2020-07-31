Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force -ErrorAction Stop

$FilePath = "$PSScriptRoot\Output\PSWriteExcel.FormulaTest7.xlsx"

$Excel = New-ExcelDocument -Verbose
$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 10' -Suppress $False -Option 'Replace'

#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 5 -CellValue 5
#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 2 -CellColumn 5 -CellValue 5
#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 3 -CellColumn 5 -CellValue 5

#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 7 -CellColumn 5 -CellFormula 'SUM(E1:E5)'
# this is also supported but we will remove = for user
#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 8 -CellColumn 5 -CellFormula '=SUM(E1:E5)'

#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 4 -CellColumn 2 -CellValue 10
Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 8 -CellColumn 6 -CellFormula '1+1'

# this forces Excel to calculate formulas in the worksheet
Request-ExcelWorkSheetCalculation -ExcelWorksheet $ExcelWorkSheet

# This gets data from it, to confirm recalculation worked
#Get-ExcelWorkSheetCell -CellRow 8 -CellColumn 5 -ExcelWorksheet $ExcelWorkSheet
Get-ExcelWorkSheetCell -CellRow 8 -CellColumn 6 -ExcelWorksheet $ExcelWorkSheet

$ExcelWorkSheet[0].Cells | Format-Table *

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
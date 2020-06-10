Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force

# this uses file created in Example02.
$FilePath = "$PSScriptRoot\Output\PSWriteExcel.FormulaTest.xlsx"

$Excel = Get-ExcelDocument -Path $FilePath
$ExcelWorkSheet = Get-ExcelWorkSheet -ExcelDocument $Excel -Name 'Test 10'

# This gets data from it, to confirm recalculation worked
Get-ExcelWorkSheetCell -CellRow 8 -CellColumn 5 -ExcelWorksheet $ExcelWorkSheet

Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 2 -CellColumn 5 -CellValue 4000

# this forces Excel to calculate formulas in the worksheet
Request-ExcelWorkSheetCalculation -ExcelWorksheet $ExcelWorkSheet

# This gets data from it, to confirm recalculation worked
Get-ExcelWorkSheetCell -CellRow 8 -CellColumn 5 -ExcelWorksheet $ExcelWorkSheet

Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 2 -CellColumn 5 -CellValue 5000

# You can also use Excel + Name of the worksheet to recalculate
Request-ExcelWorkSheetCalculation -Excel $Excel -Name 'Test 10'

# This gets data from it, to confirm recalculation worked
Get-ExcelWorkSheetCell -CellRow 8 -CellColumn 5 -ExcelWorksheet $ExcelWorkSheet

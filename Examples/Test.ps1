Import-Module -Name PSWriteExcel

$FilePath = "$PSScriptRoot\pswriteexcel_cell.xlsx"
$FilePath2 = "$PSScriptRoot\pswriteexcel_cell1.xlsx"
$Excel = Get-ExcelDocument -Path $FilePath
$Worksheet = Get-ExcelWorkSheet -ExcelDocument $Excel -Name 'Sheet1'

# read values from existing formulas
For ($i = 1; $i -le 40; $i++) {

    $RowFormula = $Worksheet.Cells[$i, 4].Formula.PadRight(30)
    $RowText = $Worksheet.Cells[$i, 4].Text.PadLeft(5)
    $Report = "Row " + $i + ": Formula: " + $RowFormula + "`t" + "Text: " + $RowText
    Write-Output $Report

}

$update = 5;

# update some data
For ($i = 1; $i -le 40; $i++) {

    Add-ExcelWorkSheetCell -ExcelWorksheet $WorkSheet -CellRow $i -CellColumn 3 -CellValue $update
    $Formula = $Worksheet.Cells[$i, 4].Formula
    #Add-ExcelWorkSheetCell -ExcelWorksheet $Worksheet -CellRow $i -CellColumn 4 -CellFormula "=($Formula)"
    #Add-ExcelWorkSheetCell -ExcelWorksheet $Worksheet -CellRow $i -CellColumn 4

}

# re-calculate sheet
Request-ExcelWorkSheetCalculation -ExcelWorksheet $WorkSheet

# re-read values
For ($i = 1; $i -le 40; $i++) {

    $RowFormula = $Worksheet.Cells[$i, 4].Formula.PadRight(30)
    $RowText = $Worksheet.Cells[$i, 4].Text.PadLeft(5)
    $Report = "Row " + $i + ": Formula: " + $RowFormula + "`t" + "Text: " + $RowText
    Write-Output $Report

}

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath2
function Get-ExcelWorkSheetCell {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [bool] $Supress
    )
    if ($ExcelWorksheet) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value
    }
}
function Get-ExcelWorkSheetCell {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [alias('Supress')][bool] $Suppress
    )
    if ($ExcelWorksheet) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value
    }
}
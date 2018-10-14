function Get-ExcelWorkSheetCell {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [bool] $Supress
    )
    if ($ExcelWorksheet) {
        $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value
    }
    return $Data
}
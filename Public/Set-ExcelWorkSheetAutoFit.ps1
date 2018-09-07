function Set-ExcelWorksheetAutoFit {
    [CmdletBinding()]
    param (
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet
    )
    if ($ExcelWorksheet) {
        Write-Verbose "Set-ExcelWorksheetAutoFit - Columns Count: $($ExcelWorksheet.Dimension.Columns)"
        if ($ExcelWorksheet.Dimension.Columns -gt 0) {
            $ExcelWorksheet.Cells.AutoFitColumns(0)
        }
    }
}
function Set-ExcelWorksheetAutoFit {
    [CmdletBinding()]
    param (
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet
    )
    if ($ExcelWorksheet) {
        Write-Verbose "Set-ExcelWorksheetAutoFit - Columns Count: $($ExcelWorksheet.Dimension.Columns)"
        if ($ExcelWorksheet.Dimension.Columns -gt 0) {
            try {
                $ExcelWorksheet.Cells.AutoFitColumns()
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                Write-Warning "Set-ExcelWorksheetAutoFit - Failed AutoFit with error message: $ErrorMessage"
            }
        }
    }
}
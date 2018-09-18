function Set-ExcelWorksheetAutoFilter {
    [CmdletBinding()]
    param (
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [string] $DataRange,
        [bool] $AutoFilter
    )
    if ($ExcelWorksheet) {
        if (-not $DataRange) {
            # if $DateRange was not provided try to get one from worksheet dimensions
            $DataRange = $ExcelWorksheet.Dimension
        }
        try {
            $ExcelWorksheet.Cells[$DataRange].AutoFilter = $AutoFilter
        } catch {
            $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
            Write-Warning "Set-ExcelWorksheetAutoFilter - Failed AutoFilter with error message: $ErrorMessage"
        }

    }
}
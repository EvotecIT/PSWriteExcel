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
        $ExcelWorksheet.Cells[$DataRange].AutoFilter = $AutoFilter

    }
}
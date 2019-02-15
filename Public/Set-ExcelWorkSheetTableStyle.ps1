function Set-ExcelWorkSheetTableStyle {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [string] $DataRange,
        [alias('TableStyles')][nullable[OfficeOpenXml.Table.TableStyles]] $TableStyle,
        [string] $TableName = $(Get-RandomStringName -LettersOnly -Size 5 -ToLower)
    )
    try {
        if ($null -ne $ExcelWorksheet) {
            if ($ExcelWorksheet.AutoFilterAddress) {
                # AutoFilter doesn't work with Styles
                return
            }
            if (-not $DataRange) {
                # if $DateRange was not provided try to get one from worksheet dimensions
                $DataRange = $ExcelWorksheet.Dimension
            }
            if ($null -ne $TableStyle) {
                Write-Verbose "Set-ExcelWorkSheetTableStyle - Setting style to $TableStyle"
                $ExcelWorkSheetTables = $ExcelWorksheet.Tables.Add($DataRange, $TableName) 
                $ExcelWorkSheetTables.TableStyle = $TableStyle
            }
        }
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Warning "Set-ExcelWorkSheetTableStyle - Worksheet: $($ExcelWorksheet.Name) error: $ErrorMessage"
    }
}
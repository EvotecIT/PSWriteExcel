function Add-ExcelWorkSheetCell {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [Object] $CellValue,
        [string] $CellFormula
    )
    if ($ExcelWorksheet) {
        if ($PSBoundParameters.Keys -contains 'CellValue') {
            Switch ($CellValue) {
                { $_ -is [PSCustomObject] } {
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                    break
                }
                { $_ -is [Array] } {
                    $Value = $CellValue -join [System.Environment]::NewLine
                    if ($Value.Length -gt 32767) {
                        Write-Warning "Add-ExcelWorkSheetCell - Triming cell lenght from $($CellValue.Length) to 32767 as maximum limit to prevent errors (Worksheet: $($ExcelWorksheet.Name) Row: $CellRow Column: $CellColumn)."
                        # We need to trim value as it's too long for a single cell
                        $Value = $Value.Substring(0, 32767)
                    }
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $Value
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.WrapText = $true
                    break
                }
                { $_ -is [DateTime] } {
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'm/d/yy h:mm'
                    break
                }
                { $_ -is [TimeSpan] } {
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'hh:mm:ss'
                    break
                }
                { $_ -is [Int64] } {
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = '#'
                    break
                }
                Default {
                    if ($CellValue.Length -gt 32767) {
                        # We need to trim value as it's too long for a single cell
                        Write-Warning "Add-ExcelWorkSheetCell - Triming cell lenght from $($CellValue.Length) to 32767 as maximum limit to prevent errors (Worksheet: $($ExcelWorksheet.Name) Row: $CellRow Column: $CellColumn)."
                        $CellValue = $CellValue.Substring(0, 32767)
                    }
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                }
            }
        } elseif ($PSBoundParameters.Keys -contains 'CellFormula') {
            # This makes sure = is removed as it's bad idea but Excel users may use it for whatever reason
            if ($CellFormula.StartsWith('=')) {
                $CellFormula = $CellFormula.Substring(1)
            }
            $ExcelWorksheet.Cells[$CellRow, $CellColumn].Formula = $CellFormula
        }
    }
}
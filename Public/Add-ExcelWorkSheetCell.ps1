function Add-ExcelWorkSheetCell {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet]  $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [Object] $CellValue,
        [string] $CellFormula
    )
    if ($ExcelWorksheet) {
        if (-not $CellFormula) {
            Switch ($CellValue) {
                { $_ -is [PSCustomObject] } {
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                    break
                }
                { $_ -is [Array] } {
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue -join [System.Environment]::NewLine
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
                    $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                }
            }
        } elseif ($CellFormula) {
            # This makes sure = is removed as it's bad idea but Excel users may use it for whatever reason
            if ($CellFormula.StartsWith('=')) {
                $CellFormula = $CellFormula.Substring(1)
            }
            $ExcelWorksheet.Cells[$CellRow, $CellColumn].Formula = $CellFormula
        }
    }
}
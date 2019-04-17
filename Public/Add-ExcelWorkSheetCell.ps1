function Add-ExcelWorkSheetCell {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet]  $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [Object] $CellValue
    )
    if ($ExcelWorksheet) {
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
            { $_ -is [DateTime]} {
                $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'm/d/yy h:mm'
                break
            }
            { $_ -is [TimeSpan]} {
                $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'hh:mm:ss'
                break
            }
            { $_ -is [Int64]} {
                $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = '#'
                break
            }
            Default {
                $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
            }
        }
    }
}
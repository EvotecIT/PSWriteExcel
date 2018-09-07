function Add-ExcelWorkSheetCell {
    param(
        [OfficeOpenXml.ExcelWorksheet]  $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [Object] $CellValue,
        [bool] $Supress
    )
    if ($ExcelWorksheet) {
        $Type = Get-ObjectType $CellValue
        Switch ($CellValue) {
            { $_ -and $Type.ObjectTypeName -eq 'PSCustomObject' } {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                break
            }
            { $_ -and $Type.ObjectTypeName -eq 'Object[]' } {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue -join [System.Environment]::NewLine
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.WrapText = $true
                break
            }
            { $_ -is [DateTime]} {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'm/d/yy h:mm'
                break
            }
            { $_ -is [TimeSpan]} {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'hh:mm:ss'
                break
            }
            { $_ -is [Int64]} {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = '#'
            }
            Default {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
            }
        }

    }
    if ($Supress) { return } else { $Data }
}
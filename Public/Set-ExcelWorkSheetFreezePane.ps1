function Set-ExcelWorkSheetFreezePane {
    param(
        [OfficeOpenXml.ExcelWorksheet]  $ExcelWorksheet,
        [Switch] $FreezeTopRow,
        [Switch] $FreezeFirstColumn,
        [Switch] $FreezeTopRowFirstColumn,
        [int[]]$FreezePane
    )
    try {
        if ($FreezeTopRowFirstColumn) {
            $ExcelWorksheet.View.FreezePanes(2, 2)
        } elseif ($FreezeTopRow -and $FreezeFirstColumn) {
            $ExcelWorksheet.View.FreezePanes(2, 2)
        } elseif ($FreezeTopRow) {
            $ExcelWorksheet.View.FreezePanes(2, 1)
        } elseif ($FreezeFirstColumn) {
            $ExcelWorksheet.View.FreezePanes(1, 2)
        }

        if ($FreezePane) {
            if ($FreezePane.Count -eq 2) {
                if ($FreezePane -notcontains 0) {
                    # check for row or column not being 0
                    if ($FreezePane[1] -gt 1) {
                        # check for column greater then 1
                        $ExcelWorksheet.View.FreezePanes($FreezePane[0], $FreezePane[1])
                    }
                }
            }
        }
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Warning "Set-ExcelWorkSheetFreezePane - $($ExcelWorksheet.Name) error: $ErrorMessage"
    }
}
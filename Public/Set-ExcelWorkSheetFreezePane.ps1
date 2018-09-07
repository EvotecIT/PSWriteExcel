function Set-ExcelWorkSheetFreezePane {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [Switch] $FreezeTopRow,
        [Switch] $FreezeFirstColumn,
        [Switch] $FreezeTopRowFirstColumn,
        [int[]]$FreezePane
    )
    try {

        if ($ExcelWorksheet -ne $null) {
            if ($FreezeTopRowFirstColumn) {
                Write-Verbose 'Set-ExcelWorkSheetFreezePane - Processing freezing panes FreezeTopRowFirstColumn'
                $ExcelWorksheet.View.FreezePanes(2, 2)
            } elseif ($FreezeTopRow -and $FreezeFirstColumn) {
                Write-Verbose 'Set-ExcelWorkSheetFreezePane - Processing freezing panes FreezeTopRow and FreezeFirstColumn'
                $ExcelWorksheet.View.FreezePanes(2, 2)
            } elseif ($FreezeTopRow) {
                Write-Verbose 'Set-ExcelWorkSheetFreezePane - Processing freezing panes FreezeTopRow'
                $ExcelWorksheet.View.FreezePanes(2, 1)
            } elseif ($FreezeFirstColumn) {
                Write-Verbose 'Set-ExcelWorkSheetFreezePane - Processing freezing panes FreezeFirstColumn'
                $ExcelWorksheet.View.FreezePanes(1, 2)
            }

            if ($FreezePane) {
                Write-Verbose 'Set-ExcelWorkSheetFreezePane - Processing freezing panes FreezePane'
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
        }
        #else {
        #    Write-Verbose 'Set-ExcelWorkSheetFreezePane - ExcelWorkSheet is null'
        #}
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Warning "Set-ExcelWorkSheetFreezePane - Worksheet: $($ExcelWorksheet.Name) error: $ErrorMessage"
    }
}
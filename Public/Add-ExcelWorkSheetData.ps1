function Add-ExcelWorksheetData {
    [CmdletBinding()]
    Param(
        [alias('ExcelWorkbook')][OfficeOpenXml.ExcelPackage] $ExcelDocument,
        $ExcelWorksheet, # [OfficeOpenXml.ExcelWorksheet]
        [Parameter(ValueFromPipeline = $true)][Object] $DataTable,
        [int]$StartRow = 1,
        [int]$StartColumn = 1,
        [alias("Autosize")][switch] $AutoFit,
        [switch] $AutoFilter,
        [Switch] $FreezeTopRow,
        [Switch] $FreezeFirstColumn,
        [Switch] $FreezeTopRowFirstColumn,
        [int[]]$FreezePane,
        [alias('Name', 'WorksheetName')][string] $ExcelWorksheetName,
        [alias('Rotate', 'RotateData', 'TransposeColumnsRows', 'TransposeData')][switch] $Transpose,
        [ValidateSet("ASC", "DESC", "NONE")][string] $TransposeSort = 'NONE',
        [switch] $PreScanHeaders, # this feature scans properties of an object for all objects it contains to make sure all headers are there
        [bool] $Supress
    )
    Begin {
        $FirstRun = $True
        $RowNr = if ($StartRow -ne $null -and $StartRow -ne 0) { $StartRow } else { 1 }
        $ColumnNr = if ($StartColumn -ne $null -and $StartColumn -ne 0 ) { $StartColumn } else { 1 }
        if ($ExcelWorksheet -ne $null) {
            Write-Verbose "Add-ExcelWorkSheetData - ExcelWorksheet given. Continuing..."
        } else {
            if ($ExcelDocument) {
                $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $ExcelWorksheetName
            } else {
                Write-Warning 'Add-ExcelWorksheetData - ExcelDocument and ExcelWorksheet not given. No data will be added...'
                # throw 'Add-ExcelWorksheetData - ExcelDocument and ExcelWorksheet not given. Terminating.'
            }
        }
        Write-Verbose "Add-ExcelWorksheetData - Excel Row: $RowNr Column: $ColumnNr"
    }
    Process {
        if ((Get-ObjectCount -Object $DataTable) -ne 0) {
            if ($FirstRun) {
                $FirstRun = $false
                #Write-Verbose "Add-ExcelWorksheetData - FirstRun - RowsToProcess: $($DataTable.Count) - Transpose: $Transpose AutoFit: $Autofit Autofilter: $Autofilter"
                if ($Transpose) { $DataTable = Format-TransposeTable -Object $DataTable -Sort $TransposeSort }
                $Data = Format-PSTable -Object $DataTable -ExcludeProperty $ExcludeProperty -NoAliasOrScriptProperties:$NoAliasOrScriptProperties -DisplayPropertySet:$DisplayPropertySet -PreScanHeaders:$PreScanHeaders # -SkipTitle:$NoHeader
                $WorksheetHeaders = $Data[0] # Saving Header information for later use
                #Write-Verbose "Add-ExcelWorksheetData - Headers: $($WorksheetHeaders -join ', ') - Data Count: $($Data.Count)"
                if ($NoHeader) {
                    $Data.RemoveAt(0);
                    #Write-Verbose "Removed header from ArrayList - Data Count: $($Data.Count)"
                }
                $ArrRowNr = 0
                foreach ($RowData in $Data) {
                    $ArrColumnNr = 0
                    $ColumnNr = $StartColumn
                    foreach ($Value in $RowData) {
                        #Write-Verbose "Row: $RowNr / $ArrRowNr Column: $ColumnNr / $ArrColumnNr Data: $Value Title: $($WorksheetHeaders[$ArrColumnNr])"
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Value -Supress $True
                        $ColumnNr++
                        $ArrColumnNr++
                    }
                    $ArrRowNr++
                    $RowNr++

                }
            } else {
                #Write-Verbose "Add-ExcelWorksheetData - NextRun - RowsToProcess: $($DataTable.Count) - Transpose: $Transpose AutoFit: $Autofit Autofilter: $Autofilter"
                if ($Transpose) { $DataTable = Format-TransposeTable -Object $DataTable -Sort $TransposeSort }
                $Data = Format-PSTable $DataTable -SkipTitle -ExcludeProperty $ExcludeProperty -NoAliasOrScriptProperties:$NoAliasOrScriptProperties -DisplayPropertySet:$DisplayPropertySet -OverwriteHeaders $WorksheetHeaders -PreScanHeaders:$PreScanHeaders
                $ArrRowNr = 0
                foreach ($RowData in $Data) {
                    $ArrColumnNr = 0
                    $ColumnNr = $StartColumn
                    foreach ($Value in $RowData) {
                        #Write-Verbose "Row: $RowNr / $ArrRowNr Column: $ColumnNr / $ArrColumnNr Data: $Value Title: $($WorksheetHeaders[$ArrColumnNr])"
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Value -Supress $True
                        $ColumnNr++; $ArrColumnNr++
                    }
                    $RowNr++; $ArrRowNr++
                }
            }
        }
    }
    End {
        if ($AutoFit) { Set-ExcelWorksheetAutoFit -ExcelWorksheet $ExcelWorksheet }
        if ($AutoFilter) { Set-ExcelWorksheetAutoFilter -ExcelWorksheet $ExcelWorksheet -DataRange $ExcelWorksheet.Dimension -AutoFilter $AutoFilter }
        if ($FreezeTopRow -or $FreezeFirstColumn -or $FreezeTopRowFirstColumn -or $FreezePane) {
            Set-ExcelWorkSheetFreezePane -ExcelWorksheet $ExcelWorksheet `
                -FreezeTopRow:$FreezeTopRow `
                -FreezeFirstColumn:$FreezeFirstColumn `
                -FreezeTopRowFirstColumn:$FreezeTopRowFirstColumn `
                -FreezePane $FreezePane
        }
        if ($Supress) { return } else { return $ExcelWorkSheet }
    }

}

function Add-ExcelWorksheetData {
    [CmdletBinding()]
    Param(
        [alias('ExcelWorkbook')][OfficeOpenXml.ExcelPackage] $ExcelDocument,
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet, # [OfficeOpenXml.ExcelWorksheet]
        [Parameter(ValueFromPipeline = $true)][Array] $DataTable,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
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
        [alias('PreScanHeaders')][switch] $AllProperties, # this feature scans properties of an object for all objects it contains to make sure all headers are there
        [alias('TableStyles')][nullable[OfficeOpenXml.Table.TableStyles]] $TableStyle,
        [string] $TableName,
        [string] $TabColor,
        [alias('Supress')][bool] $Suppress
    )
    Begin {
        $FirstRun = $True
        $RowNr = if ($null -ne $StartRow -and $StartRow -ne 0) { $StartRow } else { 1 }
        $ColumnNr = if ($null -ne $StartColumn -and $StartColumn -ne 0 ) { $StartColumn } else { 1 }
        if ($null -ne $ExcelWorksheet) {
            Write-Verbose "Add-ExcelWorkSheetData - ExcelWorksheet given. Continuing..."
        } else {
            if ($ExcelDocument) {
                $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $ExcelWorksheetName -Option $Option
                Write-Verbose "Add-ExcelWorkSheetData - ExcelWorksheet $($ExcelWorkSheet.Name)"
            } else {
                Write-Warning 'Add-ExcelWorksheetData - ExcelDocument and ExcelWorksheet not given. No data will be added...'
            }
        }
        if ($AutoFilter -and $TableStyle) {
            Write-Warning 'Add-ExcelWorksheetData - Using AutoFilter and TableStyle is not supported at same time. TableStyle will be skipped.'
        }
    }
    Process {
        if ($DataTable.Count -gt 0) {
            if ($FirstRun) {
                $FirstRun = $false
                if ($Transpose) {
                    $DataTable = Format-TransposeTable -Object $DataTable -Sort $TransposeSort
                }
                <#
                if ($DataTable[0] -is [System.Collections.IDictionary]) {
                    $Header = @('Name', 'Value')
                    foreach ($Head in $Header) {
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Head
                        $ColumnNr++
                    }
                    $RowNr++
                    foreach ($Data in $DataTable) {
                        foreach ($_ in $Data.GetEnumerator()) {
                            $ColumnNr = $StartColumn
                            Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $_.Name
                            Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn ($ColumnNr + 1) -CellValue $_.Value
                            $RowNr++
                        }
                    }
                } elseif ($DataTable[0].GetType().Name -match 'bool|byte|char|datetime|decimal|double|ExcelHyperLink|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort') {
                    foreach ($Data in $DataTable) {
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Data
                        $RowNr++
                    }
                } else {
                    # PSCustomobject
                    # Header processng
                    $Header = $DataTable[0].PSObject.Properties.Name
                    foreach ($Head in $Header) {
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Head
                        $ColumnNr++
                    }
                    $RowNr++
                    # Data processing
                    foreach ($Data in $DataTable) {
                        $ColumnNr = $StartColumn
                        foreach ($HeaderName in $Header) {
                            $Value = $Data.$HeaderName
                            Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Value
                             $ColumnNr++
                        }
                        $RowNr++
                    }
                }
#>

                $Data = Format-PSTable -Object $DataTable -ExcludeProperty $ExcludeProperty -PreScanHeaders:$AllProperties.IsPresent # -SkipTitle:$NoHeader
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
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Value
                        $ColumnNr++
                        $ArrColumnNr++
                    }
                    $ArrRowNr++
                    $RowNr++
                }
                <#
#>
            } else {
                #Write-Verbose "Add-ExcelWorksheetData - NextRun - RowsToProcess: $($DataTable.Count) - Transpose: $Transpose AutoFit: $Autofit Autofilter: $Autofilter"

                if ($Transpose) {
                    $DataTable = Format-TransposeTable -Object $DataTable -Sort $TransposeSort
                }

                <#
                if ($DataTable[0] -is [System.Collections.IDictionary]) {
                    foreach ($Data in $DataTable) {
                        foreach ($_ in $Data.GetEnumerator()) {
                            $ColumnNr = $StartColumn
                            Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $_.Name
                            Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn ($ColumnNr + 1) -CellValue $_.Value
                            $RowNr++
                        }
                    }
                } elseif ($DataTable[0].GetType().Name -match 'bool|byte|char|datetime|decimal|double|ExcelHyperLink|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort') {
                    foreach ($Data in $DataTable) {
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Data
                        $RowNr++
                    }
                } else {
                    # Data processing
                    foreach ($Data in $DataTable) {
                        $ColumnNr = $StartColumn
                        foreach ($HeaderName in $Header) {
                            $Value = $Data.$HeaderName
                            Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Value
                            $ColumnNr++
                        }
                        $RowNr++
                    }
                }
#>

                $Data = Format-PSTable -Object $DataTable -SkipTitle -ExcludeProperty $ExcludeProperty -OverwriteHeaders $WorksheetHeaders -PreScanHeaders:$PreScanHeaders
                $ArrRowNr = 0
                foreach ($RowData in $Data) {
                    $ArrColumnNr = 0
                    $ColumnNr = $StartColumn
                    foreach ($Value in $RowData) {
                        #Write-Verbose "Row: $RowNr / $ArrRowNr Column: $ColumnNr / $ArrColumnNr Data: $Value Title: $($WorksheetHeaders[$ArrColumnNr])"
                        Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorksheet -CellRow $RowNr -CellColumn $ColumnNr -CellValue $Value
                        $ColumnNr++; $ArrColumnNr++
                    }
                    $RowNr++; $ArrRowNr++
                }

            }

        }
    }
    End {
        if ($null -ne $ExcelWorksheet) {
            if ($AutoFit) { Set-ExcelWorksheetAutoFit -ExcelWorksheet $ExcelWorksheet }
            if ($AutoFilter) { Set-ExcelWorksheetAutoFilter -ExcelWorksheet $ExcelWorksheet -DataRange $ExcelWorksheet.Dimension -AutoFilter $AutoFilter }
            if ($FreezeTopRow -or $FreezeFirstColumn -or $FreezeTopRowFirstColumn -or $FreezePane) {
                Set-ExcelWorkSheetFreezePane -ExcelWorksheet $ExcelWorksheet `
                    -FreezeTopRow:$FreezeTopRow `
                    -FreezeFirstColumn:$FreezeFirstColumn `
                    -FreezeTopRowFirstColumn:$FreezeTopRowFirstColumn `
                    -FreezePane $FreezePane
            }
            if ($TableStyle) {
                Set-ExcelWorkSheetTableStyle -ExcelWorksheet $ExcelWorksheet -TableStyle $TableStyle -DataRange $ExcelWorksheet.Dimension -TableName $TableName
            }
            if ($TabColor) {
                $ExcelWorksheet.TabColor = ConvertFrom-Color -Color $TabColor -AsDrawingColor
            }
            #Write-Verbose 'Add-ExcelWorksheetData - Ending...'
            if ($Suppress) { return } else { return $ExcelWorkSheet }
        }
    }

}
$ScriptBlockColors = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    $Script:RGBColors.Keys | Where-Object { $_ -like "$wordToComplete*" }
}

Register-ArgumentCompleter -CommandName Add-ExcelWorksheetData -ParameterName TabColor -ScriptBlock $ScriptBlockColors
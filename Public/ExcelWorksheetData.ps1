function Add-ExcelWorksheetData {
    [CmdletBinding()]
    Param(
        [alias('ExcelWorkbook')][OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [Parameter(ValueFromPipeline = $true)]$DataTable,
        [int]$StartRow = 1,
        [int]$StartColumn = 1,
        [switch] $AutoFit,
        [switch] $AutoFilter,
        [alias('Name', 'WorksheetName')][string] $ExcelWorksheetName,
        [alias('Rotate', 'RotateData', 'TransposeColumnsRows', 'TransposeData')][switch] $Transpose,
        [ValidateSet("ASC", "DESC", "NONE")][string] $TransposeSort = 'NONE'
    )
    Begin {
        $FirstRun = $True
        $RowNr = $StartRow
        $ColumnNr = $StartColumn
        if ($ExcelWorksheet) {
            Write-Verbose "Add-ExcelWorkSheetData - ExcelWorksheet given. Continuing..."
        } else {
            if ($ExcelDocument) {
                $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $ExcelWorksheetName
            } else {
                Write-Warning 'Add-ExcelWorksheetData - ExcelDocument and ExcelWorksheet not given. No data will be added...'
                # throw 'Add-ExcelWorksheetData - ExcelDocument and ExcelWorksheet not given. Terminating.'
            }
        }
    }
    Process {
        if ((Get-ObjectCount -Object $DataTable) -ne 0) {
            if ($FirstRun) {
                $FirstRun = $false
                Write-Verbose "Add-ExcelWorksheetData - FirstRun - RowsToProcess: $($DataTable.Count) - Transpose: $Transpose AutoFit: $Autofit Autofilter: $Autofilter"
                if ($Transpose) { $DataTable = Format-TransposeTable -Object $DataTable -Sort $TransposeSort }
                $Data = Format-PSTable $DataTable -ExcludeProperty $ExcludeProperty -NoAliasOrScriptProperties:$NoAliasOrScriptProperties -DisplayPropertySet:$DisplayPropertySet # -SkipTitle:$NoHeader
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
                Write-Verbose "Add-ExcelWorksheetData - NextRun - RowsToProcess: $($DataTable.Count) - Transpose: $Transpose AutoFit: $Autofit Autofilter: $Autofilter"
                if ($Transpose) { $DataTable = Format-TransposeTable -Object $DataTable -Sort $TransposeSort }
                $Data = Format-PSTable $DataTable -SkipTitle -ExcludeProperty $ExcludeProperty -NoAliasOrScriptProperties:$NoAliasOrScriptProperties -DisplayPropertySet:$DisplayPropertySet -OverwriteHeaders $WorksheetHeaders
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
    }
}

function Set-ExcelWorksheetAutoFilter {
    [CmdletBinding()]
    param (
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [string] $DataRange,
        [bool] $AutoFilter
    )
    if ($ExcelWorksheet) {
        $ExcelWorksheet.Cells[$DataRange].AutoFilter = $AutoFilter
    }
}
function Set-ExcelWorksheetAutoFit {
    [CmdletBinding()]
    param (
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet
    )
    if ($ExcelWorksheet) {
        Write-Verbose "Set-ExcelWorksheetAutoFit - Columns Count: $($ExcelWorksheet.Dimension.Columns)"
        if ($ExcelWorksheet.Dimension.Columns -gt 0) {
            $ExcelWorksheet.Cells.AutoFitColumns(0)
        }
    }
}
function Set-ExcelTranslateFromR1C1 {
    [CmdletBinding()]
    param(
        [int]$Row,
        [int]$Column = 1
    )

    $Range = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[$Row]C[$Column]", 0, 0)
    return $Range
}
function Remove-ExcelWorksheet {
    [CmdletBinding()]
    param (
        [alias('ExcelWorkbook')][OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet
    )
    if ($ExcelDocument -and $ExcelWorksheet) {
        $ExcelDocument.Workbook.Worksheets.Delete($ExcelWorksheet)
    }
}

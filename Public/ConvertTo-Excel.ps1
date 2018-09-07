function ConvertTo-Excel {
    [CmdletBinding()]
    param(
        [alias("path")][string] $FilePath,
        [OfficeOpenXml.ExcelPackage] $Excel,
        [alias('Name', 'WorksheetName')][string] $ExcelWorkSheetName,
        [alias("TargetData")][Parameter(ValueFromPipeline = $true)][Object] $DataTable,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
        [switch] $AutoFilter,
        [alias("Autosize")][switch] $AutoFit,
        [Switch] $FreezeTopRow,
        [Switch] $FreezeFirstColumn,
        [Switch] $FreezeTopRowFirstColumn,
        [int[]]$FreezePane,
        [alias('Rotate', 'RotateData', 'TransposeColumnsRows', 'TransposeData')][switch] $Transpose,
        [ValidateSet("ASC", "DESC", "NONE")][string] $TransposeSort = 'NONE',
        [switch] $OpenWorkBook

    )
    Begin {
        $Data = @()
        $FirstRun = $true
        if (Test-Path $FilePath) {
            $Excel = Get-ExcelDocument -Path $FilePath
            Write-Verbose "ConvertTo-Excel - Excel exists, Excel is loaded from file"
        }
        if ($Excel -eq $null) {
            Write-Verbose "ConvertTo-Excel - Excel is null, creating new Excel"
            $Excel = New-ExcelDocument #-Verbose
        }
        #$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName $ExcelWorkSheetName -Supress $False -Option $Option #-Verbose
    }
    Process {
        $Data += $DataTable
    }
    End {
        $Return = Add-ExcelWorksheetData `
            -DataTable $Data `
            -ExcelDocument $Excel `
            -AutoFit:$AutoFit `
            -AutoFilter:$AutoFilter `
            -ExcelWorksheetName $ExcelWorkSheetName `
            -FreezeTopRow:$FreezeTopRow `
            -FreezeFirstColumn:$FreezeFirstColumn `
            -FreezeTopRowFirstColumn:$FreezeTopRowFirstColumn `
            -FreezePane $FreezePane `
            -Transpose:$Transpose `
            -TransposeSort $TransposeSort
        Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook:$OpenWorkBook
    }
}
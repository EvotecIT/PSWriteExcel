function ConvertTo-Excel {
    [CmdletBinding()]
    param(
        [alias("path")][string] $FilePath,
        [OfficeOpenXml.ExcelPackage] $Excel,
        [alias('Name', 'WorksheetName')][string] $ExcelWorkSheetName,
        [alias("TargetData")][Parameter(ValueFromPipeline = $true)][Array] $DataTable,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
        [switch] $AutoFilter,
        [alias("Autosize")][switch] $AutoFit,
        [Switch] $FreezeTopRow,
        [Switch] $FreezeFirstColumn,
        [Switch] $FreezeTopRowFirstColumn,
        [int[]]$FreezePane,
        [alias('Rotate', 'RotateData', 'TransposeColumnsRows', 'TransposeData')][switch] $Transpose,
        [ValidateSet("ASC", "DESC", "NONE")][string] $TransposeSort = 'NONE',
        [alias('TableStyles')][nullable[OfficeOpenXml.Table.TableStyles]] $TableStyle,
        [string] $TableName,
        [switch] $OpenWorkBook,
        [alias('PreScanHeaders')][switch] $AllProperties
    )
    Begin {
        $Fail = $false
        $Data = [System.Collections.Generic.List[Object]]::new()
        if ($FilePath -like '*.xlsx') {
            if (Test-Path $FilePath) {
                $Excel = Get-ExcelDocument -Path $FilePath
                Write-Verbose "ConvertTo-Excel - Excel exists, Excel is loaded from file"
            }
        } else {
            $Fail = $true
            Write-Warning "ConvertTo-Excel - Excel path not given or incorrect (no .xlsx file format)"
            return
        }
        if ($null -eq $Excel) {
            Write-Verbose "ConvertTo-Excel - Excel is null, creating new Excel"
            $Excel = New-ExcelDocument #-Verbose
        }
    }
    Process {
        if ($Fail) { return }
        foreach ($_ in $DataTable) {
            $Data.Add($_)
        }
    }
    End {
        if ($Fail) { return }
        Add-ExcelWorksheetData `
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
            -TransposeSort $TransposeSort `
            -Option $Option `
            -TableStyle $TableStyle `
            -TableName $TableName -AllProperties:$AllProperties.IsPresent -Suppress $true
        Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook:$OpenWorkBook
    }
}
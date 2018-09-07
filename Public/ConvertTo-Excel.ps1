function ConvertTo-Excel {
    [CmdletBinding()]
    param(
        [string] $FilePath,
        [OfficeOpenXml.ExcelPackage] $Excel,
        [string] $ExcelWorkSheetName,
        [Parameter(ValueFromPipeline = $true)][Object] $DataTable,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
        [switch] $AutoFilter,
        [switch] $AutoFit,
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
            $Excel = New-ExcelDocument -Verbose
        }
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName $ExcelWorkSheetName -Supress $False -Option $Option -Verbose
    }
    Process {
        $Data += $DataTable
    }
    End {
        $ExcelWorkSheet = Add-ExcelWorksheetData -DataTable $Data -ExcelDocument $Excel -ExcelWorksheet $ExcelWorkSheet -AutoFit:$AutoFit -AutoFilter:$AutoFilter -StartRow $StartRow -ExcelWorksheetName $ExcelWorkSheetName -Verbose
        Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook:$OpenWorkBook
    }
}
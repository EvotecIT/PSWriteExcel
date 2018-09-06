function ConvertTo-Excel {
    param(
        [string] $FilePath,
        [OfficeOpenXml.ExcelPackage] $Excel,
        [string] $ExcelWorkSheetName,
        [Object] $DataTable,
        [switch] $AutoFilter,
        [switch] $AutoFit,
        [switch] $OpenWorkBook

    )
    if ($Excel -eq $null) {
        $Excel = New-ExcelDocument -Verbose
    }
    Get-ObjectType -Object $Excel -Verbose -VerboseOnly
    $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName $ExcelWorkSheetName -Supress $False -Option 'Replace'
    Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $DataTable -AutoFit:$AutoFit -AutoFilter:$AutoFilter
    Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook:$OpenWorkBook

}
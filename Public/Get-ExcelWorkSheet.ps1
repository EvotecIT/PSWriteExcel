function Get-ExcelWorkSheet {
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [string] $Name,
        [switch] $All
    )
    if ($ExcelDocument) {
        if ($All) {
            $Data = $ExcelDocument.Workbook.Worksheets
        } else {
            $Data = $ExcelDocument.Workbook.Worksheets | Where { $_.Name -eq $Name }
        }
        return $Data
    } else {
        return
    }
}
function Get-ExcelWorkSheet {
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [string] $Name
    )
    $Data = $ExcelDocument.Workbook.Worksheets | Where { $_.Name -eq $Name }
    return $Data
}
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
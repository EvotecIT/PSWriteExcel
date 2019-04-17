function Get-ExcelProperties {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelPackage] $ExcelDocument
    )
    if ($ExcelDocument) {
        $Properties = [ordered] @{}
        foreach ($Key in $ExcelDocument.Workbook.Properties.PsObject.Properties.Name | Where { $_ -notlike '*Xml'} ) {
            $Properties.$Key = $ExcelDocument.Workbook.Properties.$Key
        }
        return $Properties
    }
}
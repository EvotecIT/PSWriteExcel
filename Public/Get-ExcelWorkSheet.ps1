function Get-ExcelWorkSheet {
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [string] $Name,
        [int] $Index,
        [switch] $All
    )
    if ($ExcelDocument) {
        if ($Name -and $Index) {
            Write-Warning 'Get-ExcelWorkSheet - Only $Name or $Index of Worksheet can be used.'
            return
        }
        if ($All) {
            $Data = $ExcelDocument.Workbook.Worksheets
        } elseif ($Name -or $Index) {
            if ($Name) {
                $Data = $ExcelDocument.Workbook.Worksheets | Where { $_.Name -eq $Name }
            }
            if ($Index) {
                if ($PSEdition -ne 'Core') {
                    $Index = $Index + 1
                }
                $Data = $ExcelDocument.Workbook.Worksheets[$Index]
            }
        }
    }
    return $Data
}
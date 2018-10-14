function Get-ExcelTranslateToR1C1 {
    [alias('Set-ExcelTranslateToR1C1')]
    [CmdletBinding()]
    param(
        [string] $Value
    )
    if ($Value -eq '') {
        return
    } else {
        $Range = [OfficeOpenXml.ExcelAddress]::TranslateToR1C1($Value, 0, 0)
        return $Range
    }
}

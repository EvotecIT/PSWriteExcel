function Get-ExcelTranslateFromR1C1 {
    [alias('Set-ExcelTranslateFromR1C1')]
    [CmdletBinding()]
    param(
        [int]$Row,
        [int]$Column = 1
    )
    $Range = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[$Row]C[$Column]", 0, 0)
    return $Range
}
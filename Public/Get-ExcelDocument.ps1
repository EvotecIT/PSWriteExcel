function Get-ExcelDocument {
    param(
        [alias("FilePath")][string] $Path
    )
    $Script:SaveCounter = 0
    $Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path
    return $Excel
}
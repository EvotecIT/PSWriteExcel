function Get-ExcelDocument {
    [CmdletBinding()]
    param(
        [alias("FilePath")][string] $Path
    )
    $Script:SaveCounter = 0
    if (Test-Path $Path) {
        $Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path
        return $Excel
    } else {
        return
    }
}
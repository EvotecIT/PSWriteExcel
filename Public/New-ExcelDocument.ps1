function New-ExcelDocument {
    param(

    )
    $Script:SaveCounter = 0
    $Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage
    return $Excel
}
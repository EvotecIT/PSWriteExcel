function New-ExcelDocument {
    param(

    )
    $Excel = New-Object -TypeName OfficeOpenXml.ExcelPackage
    return $Excel
}

function Save-ExcelDocument {
    param (
        [parameter(Mandatory = $true, ValueFromPipeline = $true)][Alias('Document', 'Excel', 'Package')] $ExcelDocument,
        [string] $FilePath,
        [alias('Show', 'Open')][switch] $OpenWorkBook
    )
    $ExcelDocument.SaveAs($FilePath)

    if ($OpenWorkBook) { Invoke-Item -Path $FilePath }
}
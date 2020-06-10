function New-ExcelDocument {
    [CmdletBinding()]
    param()
    $Script:SaveCounter = 0
    <#
    OfficeOpenXml.ExcelPackage new()
    OfficeOpenXml.ExcelPackage new(System.IO.FileInfo newFile)
    OfficeOpenXml.ExcelPackage new(System.IO.FileInfo newFile, string password)
    OfficeOpenXml.ExcelPackage new(System.IO.FileInfo newFile, System.IO.FileInfo template)
    OfficeOpenXml.ExcelPackage new(System.IO.FileInfo newFile, System.IO.FileInfo template, string password)
    OfficeOpenXml.ExcelPackage new(System.IO.FileInfo template, bool useStream)
    OfficeOpenXml.ExcelPackage new(System.IO.FileInfo template, bool useStream, string password)
    OfficeOpenXml.ExcelPackage new(System.IO.Stream newStream)
    OfficeOpenXml.ExcelPackage new(System.IO.Stream newStream, string Password)
    OfficeOpenXml.ExcelPackage new(System.IO.Stream newStream, System.IO.Stream templateStream)
    OfficeOpenXml.ExcelPackage new(System.IO.Stream newStream, System.IO.Stream templateStream, string Password)
    #>
    [OfficeOpenXml.ExcelPackage]::new()
}
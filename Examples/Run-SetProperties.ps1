Import-Module PSWriteExcel -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-SetProperties.xlsx"

$Excel = New-ExcelDocument -Verbose

$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'

$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True

Set-ExcelProperties -ExcelDocument $Excel -Author 'Przemyslaw Klys' -Title 'This is a test'
Set-ExcelProperties -ExcelDocument $Excel -Comments 'Testing PSWriteExcel' -Subject 'Subject'

Get-ExcelProperties -ExcelDocument $Excel

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook


Import-Module PSWriteExcel -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-Test.xlsx"
$FilePath1 = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-Test1.xlsx"

$Excel = New-ExcelDocument -Verbose

$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 10' -Supress $False -Option 'Replace'
$ExcelWorkSheet2 = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 2' -Supress $False -Option 'Replace'
$ExcelWorkSheet3 = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'This is very long title - Will be cut off at some point' -Supress $false -Option 'Replace'
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Option 'Replace' -Supress $True
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 2' -Option 'Skip' -Supress $True
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Option 'Replace' -Supress $True
#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 1 -CellValue 'Test'

$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)

Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet2 -DataTable $myitems0 -AutoFit -AutoFilter
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet3 -DataTable $myitems0 -AutoFit -AutoFilter
Add-ExcelWorksheetData -DataTable $myitems0 -Verbose -AutoFit -AutoFilter
Add-ExcelWorksheetData -DataTable $myitems0 -AutoFit -AutoFilter -ExcelDocument $Excel
Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
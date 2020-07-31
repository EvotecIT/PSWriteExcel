
#Import-Module .\PSWriteExcel.psd1 -Force #-Verbose

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-Test.xlsx"

$Excel = New-ExcelDocument -Verbose

$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 10' -Suppress $False -Option 'Replace'
$ExcelWorkSheet2 = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 2' -Suppress $False -Option 'Replace'
$ExcelWorkSheet3 = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'This is very long title - Will be cut off at some point' -Suppress $false -Option 'Replace'
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Option 'Replace' -Suppress $True
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 2' -Option 'Skip' -Suppress $True
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Option 'Replace' -Suppress $True
#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 1 -CellValue 'Test'

$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)

Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet2 -DataTable $myitems0 -AutoFit -AutoFilter  -Suppress $True -PreScanHeaders -TabColor ChartreuseYellow
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet3 -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True -TabColor DarkPurple
Add-ExcelWorksheetData -DataTable $myitems0 -Verbose -AutoFit -AutoFilter -Suppress $True -TabColor Bisque -ExcelDocument $Excel
Add-ExcelWorksheetData -DataTable $myitems0 -AutoFit -AutoFilter -ExcelDocument $Excel -Suppress $True -TabColor Grey

$myitems0 | Add-ExcelWorksheetData -AutoFit -AutoFilter -ExcelDocument $Excel -ExcelWorksheetName 'Hello Motto' -Suppress $True -FreezeTopRow -FreezeFirstColumn -Verbose

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
﻿Import-Module PSWriteExcel -Force #-Verbose

$FilePath = "$PSScriptRoot\PSWriteExcel-Example-Test.xlsx"

$myitems1 = [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}

$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)

$MyItems2 = @{
    name = "Joe"; age = 32; nfo = "Cat lover"
}

$MyItems3 = @(
    [ordered]@{name = "Joe"; age = 32; info = "Cat lover" }
    @{name = "Sue"; age = 29; info = "Dog lover" }
    @{name = "Jason another one"; age = 42; info = "Food lover" }
)

$Excel = New-ExcelDocument -Verbose
#Add-ExcelWorksheetData -ExcelDocument $Excel -DataTable $myitems0 -Verbose -AutoFit -AutoFilter -Supress $True
#Add-ExcelWorksheetData -ExcelDocument $Excel -DataTable $myitems1 -Verbose -AutoFit -AutoFilter -Supress $True
#$myitems0 | Add-ExcelWorksheetData -ExcelDocument $Excel -Verbose -AutoFit -AutoFilter -Supress $True
#$myitems1 | Add-ExcelWorksheetData -ExcelDocument $Excel -Verbose -AutoFit -AutoFilter -Supress $True
#Add-ExcelWorksheetData -ExcelDocument $Excel -DataTable $myitems2 -Verbose -AutoFit -AutoFilter -Supress $True
$myitems2 | Add-ExcelWorksheetData -ExcelDocument $Excel -Verbose -AutoFit -AutoFilter -Supress $True
$myitems3 | Add-ExcelWorksheetData -ExcelDocument $Excel -Verbose -AutoFit -AutoFilter -Supress $True
'testz','tests2' | Add-ExcelWorksheetData -ExcelDocument $Excel -Verbose -AutoFit -AutoFilter -Supress $True
#Add-ExcelWorksheetData -ExcelDocument $Excel -DataTable 'testz','tests2' -Verbose -AutoFit -AutoFilter -Supress $True
#Add-ExcelWorksheetData -ExcelDocument $Excel -DataTable $myitems2 -Verbose -AutoFit -AutoFilter -Supress $True
#Add-ExcelWorksheetData -ExcelDocument $Excel -DataTable $myitems3 -Verbose -AutoFit -AutoFilter -Supress $True
Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
<#
$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 10' -Supress $False -Option 'Replace'
$ExcelWorkSheet2 = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 2' -Supress $False -Option 'Replace'
$ExcelWorkSheet3 = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'This is very long title - Will be cut off at some point' -Supress $false -Option 'Replace'
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Option 'Replace' -Supress $True
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 2' -Option 'Skip' -Supress $True
Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Option 'Replace' -Supress $True
#Add-ExcelWorkSheetCell -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 1 -CellValue 'Test'


Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Supress $True
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet2 -DataTable $myitems0 -AutoFit -AutoFilter  -Supress $True
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet3 -DataTable $myitems0 -AutoFit -AutoFilter -Supress $True

#>



#Add-ExcelWorksheetData -DataTable $myitems0 -AutoFit -AutoFilter -ExcelDocument $Excel -Supress $True

#$myitems0 | Add-ExcelWorksheetData -AutoFit -AutoFilter -ExcelDocument $Excel -ExcelWorksheetName 'Hello Motto' -Supress $True -FreezeTopRow -FreezeFirstColumn -Verbose



#$myitems0[0].PSObject.Properties.Name


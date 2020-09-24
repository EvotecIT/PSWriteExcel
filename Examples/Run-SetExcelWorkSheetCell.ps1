Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-SetExcelWorkSheetCell.xlsx"

$Excel = New-ExcelDocument -Verbose

$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'

$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover" },
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover" },
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True

Set-ExcelWorkSheetCellStyle -ExcelWorksheet $ExcelWorkSheet -CellRow 1 -CellColumn 2 -Name 'Arial' -Size 20 -Color Red -Bold -Italic -UnderLine -UnderLineType Double

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
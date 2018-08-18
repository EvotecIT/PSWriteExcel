Import-Module PSWriteExcel -Force

$FilePath = "$Env:USERPROFILE\Desktop\new.xlsx"
$FilePathOutput = "$Env:USERPROFILE\Desktop\new1.xlsx"

$Excel = Get-ExcelDocument -Path $FilePath

$Worksheet = $Excel.Workbook.Worksheets[1]


$Worksheet.Cells[1, 2].Style.Font.Name = "B Zar"
$workSheet.Cells[1, 2].Style.Font.Size = 16
$workSheet.Cells[1, 2].Style.Font.Bold = $true

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePathOutput
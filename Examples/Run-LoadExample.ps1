Import-Module PSWriteExcel -Force

$FilePath = "$Env:USERPROFILE\Desktop\new.xlsx"
$FilePathOutput = "$Env:USERPROFILE\Desktop\new1.xlsx"

$Excel = Get-ExcelDocument -Path $FilePath

$Worksheet = $Excel.Workbook.Worksheets[1]


$Worksheet.Cells[1, 2].Style.Font.Name = "B Zar"
$workSheet.Cells[1, 2].Style.Font.Size = 16
$workSheet.Cells[1, 2].Style.Font.Bold = $true

# this shows you dimensions
$Worksheet.Dimension
$Worksheet.Dimension.Rows
$Worksheet.Dimension.Columns

for ($i = 1; $i -le $Worksheet.Dimension.Columns; $i++) {
    $Worksheet.Cells[3, $i].Style.Font.Name = 'B Zar'
    $workSheet.Cells[3, $i].Style.Font.Size = 16
}

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePathOutput
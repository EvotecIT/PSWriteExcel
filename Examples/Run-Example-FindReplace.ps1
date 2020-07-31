Import-Module PSWriteExcel -Force

<#
    Before you use this keep in mind that I've found some problem with URL as here: https://github.com/EvotecIT/PSWriteExcel/issues/1
#>


$FilePath = "$Env:USERPROFILE\Desktop\Book-Idea.xlsx"
$FilePathOutput1 = "$Env:USERPROFILE\Desktop\Book2.xlsx"
$FilePathOutput2 = "$Env:USERPROFILE\Desktop\Book3.xlsx"

# Display found cells
Find-ExcelDocumentText -FilePath $FilePath -Find 'Evotec'

# Display found cells, replace evotec with somethingelse but only if it matches case sensitivity. It won't replace Evotec. Finally save it to file
$Cells1 = Find-ExcelDocumentText -FilePath $FilePath -Find 'evotec' -Replace -ReplaceWith 'somethingelse' -FilePathTarget $FilePathOutput1

# Display found cells, replace evotec with somethingelse but don't check for case sensitivity. It will replace evotec, Evotec. Finally save it to file
$Cells2 = Find-ExcelDocumentText -FilePath $FilePath -Find 'evotec' -Replace -ReplaceWith 'somethingelse' -FilePathTarget $FilePathOutput2 -OpenWorkBook:$false -Regex

# Do not display found cells, replace evotec with somethingelse but don't check for case sensitivity. It will replace evotec, Evotec. Finally save it to file and open workbook.
Find-ExcelDocumentText -FilePath $FilePath -Find 'evotec' -Replace -ReplaceWith 'somethingelse' -FilePathTarget $FilePathOutput2 -OpenWorkBook -Regex -Suppress $true
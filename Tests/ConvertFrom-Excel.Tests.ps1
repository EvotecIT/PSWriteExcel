#Requires -Modules Pester
Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force #-Verbose

### Preparing Data Start
$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

if ($PSEdition -eq 'Core') {
    $WorkSheet = 0 # Core version has 0 based index for $Worksheets
} else {
    $WorkSheet = 1
}
$TemporaryFolder = [IO.Path]::GetTempPath()

Describe 'ConvertFrom-Excel - Should load Excel file into PSCustomObject)' {
    It 'Given (MyItems0) should convert it to Excel file and load Excel file to PSCustomObject ($Data) from specified Excel WorkSheet Name' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "17.xlsx") # same as $Env:TEMP\17.xlsx but platform independant
        $myitems0 | ConvertTo-Excel -Path $Path -AutoFilter -AutoSize -ExcelWorkSheetName 'MyRandomName'
        $Data = ConvertFrom-Excel -Path $Path -ExcelWorksheetName 'MyRandomName'
        $Data.Name[0] | Should -Be 'Joe'
        $Data.Name[1] | Should -Be 'Sue'
        $Data.Info[1] | Should -Be 'Dog lover'
    }
}
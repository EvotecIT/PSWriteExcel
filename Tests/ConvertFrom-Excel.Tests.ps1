Describe 'ConvertFrom-Excel - Should load Excel file into PSCustomObject)' {
    It 'Given (MyItems0) should convert it to Excel file and load Excel file to PSCustomObject ($Data) from specified Excel WorkSheet Name' {
        ### Preparing Data Start
        $myitems0 = @(
            [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover" },
            [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover" },
            [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover" }
        )

        $TemporaryFolder = [IO.Path]::GetTempPath()

        $Path = [IO.Path]::Combine($TemporaryFolder, "17.xlsx") # same as $Env:TEMP\17.xlsx but platform independant
        $myitems0 | ConvertTo-Excel -Path $Path -AutoFilter -AutoSize -ExcelWorkSheetName 'MyRandomName'
        $Data = ConvertFrom-Excel -Path $Path -ExcelWorksheetName 'MyRandomName'
        $Data.Name[0] | Should -Be 'Joe'
        $Data.Name[1] | Should -Be 'Sue'
        $Data.Info[1] | Should -Be 'Dog lover'
    }
    It 'Given (MyItems0) should convert it to Excel file and load Excel file to PSCustomObject ($Data) from specified Excel WorkSheet Name' {
        ### Preparing Data Start
        $myitems0 = @(
            [pscustomobject]@{bool = $true; age = 32; info = "Cat lover" },
            [pscustomobject]@{bool = $false; age = 29; info = 0 },
            [pscustomobject]@{bool = $null; age = 42; info = "Food lover" }
        )

        $TemporaryFolder = [IO.Path]::GetTempPath()

        $Path = [IO.Path]::Combine($TemporaryFolder, "17.xlsx") # same as $Env:TEMP\17.xlsx but platform independant
        $myitems0 | ConvertTo-Excel -Path $Path -AutoFilter -AutoSize -ExcelWorkSheetName 'MyRandomName'
        $Data = ConvertFrom-Excel -Path $Path -ExcelWorksheetName 'MyRandomName'
        $Data.Bool[0] | Should -Be $true
        $Data.Bool[1] | Should -Be $false
        $Data.Bool[2] | Should -Be ''
        $Data.Info[1] | Should -Be 0
    }
}
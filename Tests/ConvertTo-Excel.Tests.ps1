### Preparing Data Start
$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover" },
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover" },
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover" }
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover" }
)
$myitems2 = [PSCustomObject]@{
    name = "Joe"; age = 32; info = "Cat lover"
}

$InvoiceEntry1 = [ordered] @{ }
$InvoiceEntry1.Description = 'IT Services 1'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = [ordered] @{ }
$InvoiceEntry2.Description = 'IT Services 2'
$InvoiceEntry2.Amount = '$300'

$InvoiceEntry3 = [ordered] @{ }
$InvoiceEntry3.Description = 'IT Services 3'
$InvoiceEntry3.Amount = '$288'

$InvoiceEntry4 = [ordered]@{ }
$InvoiceEntry4.Description = 'IT Services 4'
$InvoiceEntry4.Amount = '$301'

$InvoiceEntry5 = [ordered] @{ }
$InvoiceEntry5.Description = 'IT Services 5'
$InvoiceEntry5.Amount = '$299'

$InvoiceData1 = @()
$InvoiceData1 += $InvoiceEntry1
$InvoiceData1 += $InvoiceEntry2
$InvoiceData1 += $InvoiceEntry3
$InvoiceData1 += $InvoiceEntry4
$InvoiceData1 += $InvoiceEntry5

$InvoiceData2 = $InvoiceData1.ForEach( { [PSCustomObject]$_ })

$InvoiceData3 = @()
$InvoiceData3 += $InvoiceEntry1

$InvoiceData4 = $InvoiceData3.ForEach( { [PSCustomObject]$_ })
### Preparing Data End

$Object1 = Get-Process | Select-Object ProcessName, Handle, StartTime -First 5
$Object2 = Get-PSDrive | Where-Object { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*' }
$Object3 = Get-PSDrive | Where-Object { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*' } | Select-Object * -First 2
$Object4 = Get-PSDrive | Where-Object { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*' } | Select-Object * -First 1

$obj = New-Object System.Object
$obj | Add-Member -type NoteProperty -Name Name -Value "Ryan_PC"
$obj | Add-Member -type NoteProperty -Name Manufacturer -Value "Dell"
$obj | Add-Member -type NoteProperty -Name ProcessorSpeed -Value "3 Ghz"
$obj | Add-Member -type NoteProperty -Name Memory -Value "6 GB"


$myObject2 = New-Object System.Object
$myObject2 | Add-Member -type NoteProperty -Name Name -Value "Doug_PC"
$myObject2 | Add-Member -type NoteProperty -Name Manufacturer -Value "HP"
$myObject2 | Add-Member -type NoteProperty -Name ProcessorSpeed -Value "2.6 Ghz"
$myObject2 | Add-Member -type NoteProperty -Name Memory -Value "4 GB"


$myObject3 = New-Object System.Object
$myObject3 | Add-Member -type NoteProperty -Name Name -Value "Julie_PC"
$myObject3 | Add-Member -type NoteProperty -Name Manufacturer -Value "Compaq"
$myObject3 | Add-Member -type NoteProperty -Name ProcessorSpeed -Value "2.0 Ghz"
$myObject3 | Add-Member -type NoteProperty -Name Memory -Value "2.5 GB"

$myArray1 = @($obj, $myobject2, $myObject3)
$myArray2 = @($obj)


$InvoiceEntry7 = [ordered]@{ }
$InvoiceEntry7.Description = 'IT Services 4'
$InvoiceEntry7.Amount = '$301'

$InvoiceEntry8 = [ordered]@{ }
$InvoiceEntry8.Description = 'IT Services 5'
$InvoiceEntry8.Amount = '$299'

$InvoiceDataOrdered1 = @()
$InvoiceDataOrdered1 += $InvoiceEntry7

$InvoiceDataOrdered2 = @()
$InvoiceDataOrdered2 += $InvoiceEntry7
$InvoiceDataOrdered2 += $InvoiceEntry8

if ($PSEdition -eq 'Core') {
    $WorkSheet = 0 # Core version has 0 based index for $Worksheets
} else {
    $WorkSheet = 1
}

$TemporaryFolder = [IO.Path]::GetTempPath()

$PSDefaultParameterValues = @{
    "It:TestCases" = @{
        InvoiceDataOrdered2 = $InvoiceDataOrdered2
        InvoiceDataOrdered1 = $InvoiceDataOrdered1
        InvoiceEntry8       = $InvoiceEntry8
        InvoiceEntry7       = $InvoiceEntry7
        myArray1            = $myArray1
        myArray2            = $myArray2
        myObject3           = $myObject3
        myObject2           = $myObject2
        Object1             = $Object1
        Object2             = $Object2
        Object3             = $Object3
        Object4             = $Object4
        InvoiceData4        = $InvoiceData4
        InvoiceData3        = $InvoiceData3
        InvoiceData2        = $InvoiceData2
        InvoiceData1        = $InvoiceData1
        InvoiceEntry1       = $InvoiceEntry1
        InvoiceEntry2       = $InvoiceEntry2
        InvoiceEntry3       = $InvoiceEntry3
        InvoiceEntry4       = $InvoiceEntry4
        InvoiceEntry5       = $InvoiceEntry5
        myitems0            = $myitems0
        myitems1            = $myitems1
        myitems2            = $myitems2
        obj                 = $obj
        TemporaryFolder     = $TemporaryFolder
        WorkSheet           = $WorkSheet
    }
}

Describe 'ConvertTo-Excel - Should deliver same results as Format-Table -Autosize (via pipeline)' {
    ## Cleanup of tests
    for ($i = 1; $i -le 30; $i++) {
        $Path = "$($i).xlsx"
        Remove-Item $Path -Confirm:$False -ErrorAction SilentlyContinue
    }

    It 'Given (MyItems0) should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "1.xlsx")
        $myitems0 | ConvertTo-Excel -Path $Path -AutoFilter -AutoSize -ExcelWorkSheetName 'MyRandomName'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$WorkSheet].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[$WorkSheet].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[$WorkSheet].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[$WorkSheet].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[$WorkSheet].Cells['A3'].Value | Should -Be 'Sue'
        $pkg.Workbook.Worksheets[$WorkSheet].Cells['C4'].Value | Should -Be 'Food lover'
        $pkg.Dispose()
    }
    It 'Given (MyItems1) should have 3 columns, 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "2.xlsx")
        $myitems1 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Joe'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'age'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()
    }
    It 'Given (MyItems2) should have 3 columns, 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "3.xlsx")
        $myitems1 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Joe'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'age'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()

    }
    It 'Given (InvoiceEntry1) should have 2 columns, 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "4.xlsx")
        $InvoiceEntry1 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData1) should have 2 columns, 10 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "5.xlsx")
        $InvoiceData1 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 11
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData2) should have 2 columns, 6 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "6.xlsx")
        $InvoiceData2 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 6
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'IT Services 1'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Amount'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData3) should have 2 columns, 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "7.xlsx")
        $InvoiceData3 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData4) should have 2 columns, 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "8.xlsx")
        $InvoiceData4 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'IT Services 1'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Amount'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }
    It 'Given (Object1) should have 3 columns, 6 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "9.xlsx")
        $Object1 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 6
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'ProcessName'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Handle'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()


    }

    It 'Given (Object2) should have 10 or more columns, Have 2 or more rows, data is in random order (unfortunately)' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "10.xlsx")
        $Object2 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Verbose
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -BeGreaterOrEqual 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -BeGreaterOrEqual 10
        # Not sure yet how to predict thje order. Seems order of FT -a is differnt then FL and script takes FL for now
        #$pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        #$pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        #$pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()
    }

    It 'Given (Object3) should have 10 or more columns, Have more then 1 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "11.xlsx")
        $Object3 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -BeGreaterThan 1
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -BeGreaterOrEqual 10
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Used'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Free'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (Object4) should have 10 or more columns, Have more then 1 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "12.xlsx")
        $Object4 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -BeGreaterThan 1
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -BeGreaterOrEqual 10
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Used'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Free'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (obj) should have 4 columns, Have 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "13.xlsx")
        $obj | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Manufacturer'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Dispose()

    }

    It 'Given (myArray1) should have 4 columns, Have 4 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "14.xlsx")
        $myArray1 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Manufacturer'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()

    }

    It 'Given (myArray2) should have 4 columns, Have 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "15.xlsx")
        $myArray2 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Manufacturer'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()
    }
    #>
    It 'Given (InvoiceEntry7) should have 2 columns, Have 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "16.xlsx")
        $InvoiceEntry7 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered1) should have 2 columns, Have 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "17.xlsx")
        $InvoiceDataOrdered1 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered2) should have 2 columns, Have 5 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "18.xlsx")
        $InvoiceDataOrdered2 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 5
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered2) should have 2 columns, Have 5 rows, data should be in proper columns and have proper style' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "19.xlsx")
        $InvoiceDataOrdered2 | ConvertTo-Excel -Path $Path -ExcelWorkSheetName 'MyRandomName' -TableStyle Dark10
        $Pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 5
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Tables[0].StyleName | Should -Be 'TableStyleDark10'
        $pkg.Dispose()
    }
    ## Cleanup of tests
    for ($i = 1; $i -le 30; $i++) {
        $Path = "$($i).xlsx"
        Remove-Item $Path -Confirm:$False -ErrorAction SilentlyContinue
    }
}

Describe 'ConvertTo-Excel - Should deliver same results as Format-Table -Autosize (without pipeline)' {
    ## Cleanup of tests
    for ($i = 1; $i -le 40; $i++) {
        $Path = "$($i).xlsx"
        Remove-Item $Path -Confirm:$False -ErrorAction SilentlyContinue
    }
    It 'Given (MyItems0) should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "1.xlsx")
        ConvertTo-Excel -Path $Path -AutoFilter -AutoSize -DataTable $myitems0 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Joe'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A3'].Value | Should -Be 'Sue'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C4'].Value | Should -Be 'Food lover'
        $pkg.Dispose()

    }
    It 'Given (MyItems1) should have 3 columns, 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "2.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $myitems1 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Joe'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'age'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()
    }
    It 'Given (MyItems2) should have 3 columns, 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "3.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $MyItems2 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Joe'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'age'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()

    }
    It 'Given (InvoiceEntry1) should have 2 columns, 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "4.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceEntry1 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData1) should have 2 columns, 10 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "5.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData1 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 11
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }

    It 'Given (InvoiceData2) should have 2 columns, 6 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "6.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData2 -ExcelWorkSheetName 'MyRandomName1' #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 6
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'IT Services 1'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Amount'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }

    It 'Given (InvoiceData3) should have 2 columns, 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "7.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData3 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }

    It 'Given (InvoiceData4) should have 2 columns, 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "8.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData4 -ExcelWorkSheetName 'MyRandomName1' #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'IT Services 1'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Amount'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }

    It 'Given (Object1) should have 3 columns, 6 rows, data should be in proper columns' {

        $Path = [IO.Path]::Combine($TemporaryFolder, "9.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $Object1 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 6
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'ProcessName'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Handle'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()


    }

    It 'Given (Object2) should have 10 or more columns, Have more 2 or more rows, data is in random order (unfortunately)' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "10.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $Object2 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -BeGreaterOrEqual 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -BeGreaterOrEqual 10
        # Not sure yet how to predict thje order. Seems order of FT -a is differnt then FL and script takes FL for now
        #$pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        #$pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        #$pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()
    }

    It 'Given (Object3) should have 10 or more columns, Have more then 1 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "11.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $Object3 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -BeGreaterThan 1
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -BeGreaterOrEqual 10
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Used'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Free'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (Object4) should have 10 or more columns, Have more then 1 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "12.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $Object4 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -BeGreaterThan 1
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -BeGreaterOrEqual 10
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Used'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Free'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (obj) should have 4 columns, Have 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "13.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $obj -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Manufacturer'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Dispose()

    }

    It 'Given (myArray1) should have 4 columns, Have 4 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "14.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $myArray1 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Manufacturer'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()

    }

    It 'Given (myArray2) should have 4 columns, Have 2 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "15.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $myArray2 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 4
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Manufacturer'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()
    }
    #>
    It 'Given (InvoiceEntry7) should have 2 columns, Have 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "16.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceEntry7 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered1) should have 2 columns, Have 3 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "17.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceDataOrdered1 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 3
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered2) should have 2 columns, Have 5 rows, data should be in proper columns' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "18.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceDataOrdered2 -ExcelWorkSheetName 'MyRandomName1'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 5
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered2) should have 2 columns, Have 5 rows, data should be in proper columns and have proper style' {
        $Path = [IO.Path]::Combine($TemporaryFolder, "19.xlsx")
        ConvertTo-Excel -Path $Path -DataTable $InvoiceDataOrdered2 -ExcelWorkSheetName 'MyRandomName2' -TableStyle Dark10
        $Pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Rows | Should -Be 5
        $Pkg.Workbook.Worksheets[$Worksheet].Dimension.Columns | Should -Be 2
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A1'].Value | Should -Be 'Name'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['B1'].Value | Should -Be 'Value'
        $Pkg.Workbook.Worksheets[$Worksheet].Cells['A2'].Value | Should -Be 'Description'
        $Pkg.Workbook.Worksheets[$Worksheet].Tables[0].StyleName | Should -Be 'TableStyleDark10'
        $pkg.Dispose()
    }

    ## Cleanup of tests
    for ($i = 1; $i -le 40; $i++) {
        $Path = [IO.Path]::Combine($TemporaryFolder, "$($i).xlsx")
        Remove-Item $Path -Confirm:$False -ErrorAction SilentlyContinue
    }
}
$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover" },
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover" },
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)

if ($PSEdition -eq 'Core') {
    $WorkSheet = 0 # Core version has 0 based index for $Worksheets
} else {
    $WorkSheet = 1
}

$TemporaryFolder = [IO.Path]::GetTempPath()

$PSDefaultParameterValues = @{
    "It:TestCases" = @{
        myitems0            = $myitems0
        TemporaryFolder     = $TemporaryFolder
        WorkSheet           = $WorkSheet
    }
}

Describe 'Set-ExcelProperties - Setting Excel Properties' {
    It 'Using Set-ExcelProperties - Setting Author, Title and Subject Properties should work' {
        $Excel = New-ExcelDocument -Verbose
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
        Set-ExcelProperties -ExcelDocument $Excel -Author 'Przemyslaw Klys' -Title 'PSWriteExcel Set-Properties' -Subject 'PSWriteExcel'

        $Excel.Workbook.Properties.Author | Should -Be 'Przemyslaw Klys'
        $Excel.Workbook.Properties.Title | Should -Be 'PSWriteExcel Set-Properties'
        $Excel.Workbook.Properties.Subject | Should -Be 'PSWriteExcel'
    }
    It 'Using Set-ExcelProperties - Setting Created, Modified and Category properties should work' {
        [DateTime] $Created = Get-Date -Year '2011' -Month '07' -Day '04' -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        [DateTime] $Modified = Get-Date -Year '2018' -Month '09' -Day '27' -Hour 0 -Minute 0 -Second 0 -Millisecond 0

        $Excel = New-ExcelDocument -Verbose
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
        Set-ExcelProperties -ExcelDocument $Excel -Created $Created -Modified $Modified -Category 'Excel'

        $Excel.Workbook.Properties.Created | Should -Be $Created
        $Excel.Workbook.Properties.Modified | Should -Be $Modified
        $Excel.Workbook.Properties.Category | Should -Be 'Excel'
    }
    It 'Using Set-ExcelProperties - Setting remaining properties should work' {
        [DateTime] $Created = Get-Date -Year '2011' -Month '07' -Day '04' -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        [DateTime] $Modified = Get-Date -Year '2018' -Month '09' -Day '27' -Hour 0 -Minute 0 -Second 0 -Millisecond 0

        $Excel = New-ExcelDocument -Verbose
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
        Set-ExcelProperties -ExcelDocument $Excel -Application 'Test 1' -AppVersion 'Test 2' -Keywords 'My key word' -LastModifiedBy 'Przemyslaw Klys' -LastPrinted 'Evotec' -LinksUpToDate $false -Manager 'Przemyslaw Klys' -ScaleCrop $true -SharedDoc $false -Status 'My status'

        $Excel.Workbook.Properties.Application | Should -Be 'Test 1'
        $Excel.Workbook.Properties.AppVersion | Should -Be 'Test 2'
        $Excel.Workbook.Properties.Keywords | Should -Be 'My key word'
        $Excel.Workbook.Properties.LastModifiedBy | Should -Be 'Przemyslaw Klys'
        $Excel.Workbook.Properties.LastPrinted | Should -Be 'Evotec'
        $Excel.Workbook.Properties.LinksUpToDate | Should -Be $false
        $Excel.Workbook.Properties.Manager | Should -Be 'Przemyslaw Klys'
        $Excel.Workbook.Properties.ScaleCrop | Should -Be $true
        $Excel.Workbook.Properties.SharedDoc | Should -Be $false
        $Excel.Workbook.Properties.Status | Should -Be 'My status'
    }
}
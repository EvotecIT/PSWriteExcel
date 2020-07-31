$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover" },
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover" },
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)

$TemporaryFolder = [IO.Path]::GetTempPath()

$PSDefaultParameterValues = @{
    "It:TestCases" = @{
        myitems0            = $myitems0
        TemporaryFolder     = $TemporaryFolder
        WorkSheet           = $WorkSheet
    }
}

Describe 'Get-ExcelProperties - Getting Excel Properties' {
    It 'Using Get-ExcelProperties - Getting Author, Title and Subject Properties should be readable' {
        $Excel = New-ExcelDocument
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
        Set-ExcelProperties -ExcelDocument $Excel -Author 'Przemyslaw Klys' -Title 'PSWriteExcel Set-Properties' -Subject 'PSWriteExcel'

        $Properties = Get-ExcelProperties -ExcelDocument $Excel
        $Properties.Author | Should -Be 'Przemyslaw Klys'
        $Properties.Title | Should -Be 'PSWriteExcel Set-Properties'
        $Properties.Subject | Should -Be 'PSWriteExcel'
    }
    It 'Using Get-ExcelProperties - Getting Created, Modified and Category properties should be readable' {
        [DateTime] $Created = Get-Date -Year '2011' -Month '07' -Day '04' -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        [DateTime] $Modified = Get-Date -Year '2018' -Month '09' -Day '27' -Hour 0 -Minute 0 -Second 0 -Millisecond 0

        $Excel = New-ExcelDocument
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
        Set-ExcelProperties -ExcelDocument $Excel -Created $Created -Modified $Modified -Category 'Excel'

        $Properties = Get-ExcelProperties -ExcelDocument $Excel
        $Properties.Created | Should -Be $Created
        $Properties.Modified | Should -Be $Modified
        $Properties.Category | Should -Be 'Excel'
    }
    It 'Using Get-ExcelProperties - Getting remaining properties should be readable' {
        [DateTime] $Created = Get-Date -Year '2011' -Month '07' -Day '04' -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        [DateTime] $Modified = Get-Date -Year '2018' -Month '09' -Day '27' -Hour 0 -Minute 0 -Second 0 -Millisecond 0

        $Excel = New-ExcelDocument
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
        Set-ExcelProperties -ExcelDocument $Excel -Application 'Test 1' -AppVersion 'Test 2' -Keywords 'My key word' -LastModifiedBy 'Przemyslaw Klys' -LastPrinted 'Evotec' -LinksUpToDate $false -Manager 'Przemyslaw Klys' -ScaleCrop $true -SharedDoc $false -Status 'My status'

        $Properties = Get-ExcelProperties -ExcelDocument $Excel
        $Properties.Application | Should -Be 'Test 1'
        $Properties.AppVersion | Should -Be 'Test 2'
        $Properties.Keywords | Should -Be 'My key word'
        $Properties.LastModifiedBy | Should -Be 'Przemyslaw Klys'
        $Properties.LastPrinted | Should -Be 'Evotec'
        $Properties.LinksUpToDate | Should -Be $false
        $Properties.Manager | Should -Be 'Przemyslaw Klys'
        $Properties.ScaleCrop | Should -Be $true
        $Properties.SharedDoc | Should -Be $false
        $Properties.Status | Should -Be 'My status'
    }
    It 'Using Get-ExcelProperties - Getting remaining properties should be readable after file has been saved and loaded again' {
        [DateTime] $Created = Get-Date -Year '2011' -Month '07' -Day '04' -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        [DateTime] $Modified = Get-Date -Year '2018' -Month '09' -Day '27' -Hour 0 -Minute 0 -Second 0 -Millisecond 0

        $FilePath = [IO.Path]::Combine($TemporaryFolder, "Get-ExpelProperties-Test.xlsx")

        $Excel = New-ExcelDocument
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True
        Set-ExcelProperties -ExcelDocument $Excel -Application 'Test 1' -AppVersion 'Test 2' -Keywords 'My key word' -LastModifiedBy 'Przemyslaw Klys' -LastPrinted 'Evotec' -LinksUpToDate $false -Manager 'Przemyslaw Klys' -ScaleCrop $true -SharedDoc $false -Status 'My status'
        Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath

        $ExcelOther = Get-ExcelDocument -Path $FilePath
        $Properties = Get-ExcelProperties -ExcelDocument $ExcelOther
        $Properties.Application | Should -Be 'Test 1'
        $Properties.AppVersion | Should -Be 'Test 2'
        $Properties.Keywords | Should -Be 'My key word'
        $Properties.LastModifiedBy | Should -Be 'Przemyslaw Klys'
        $Properties.LastPrinted | Should -Be 'Evotec'
        $Properties.LinksUpToDate | Should -Be $false
        $Properties.Manager | Should -Be 'Przemyslaw Klys'
        $Properties.ScaleCrop | Should -Be $true
        $Properties.SharedDoc | Should -Be $false
        $Properties.Status | Should -Be 'My status'

        Remove-Item -Path $FilePath -Confirm:$false -ErrorAction SilentlyContinue
    }
    It 'Using Get-ExcelProperties - Previous "Supress" parameter should still work on Add-ExcelWorkSheet and Add-ExcelWorkSheetData' {
        $Excel = New-ExcelDocument
        $ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Supress $False -Option 'Replace'
        Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Supress $True
        Set-ExcelProperties -ExcelDocument $Excel -Author 'Przemyslaw Klys' -Title 'PSWriteExcel Set-Properties' -Subject 'PSWriteExcel'

        $Properties = Get-ExcelProperties -ExcelDocument $Excel
        $Properties.Author | Should -Be 'Przemyslaw Klys'
        $Properties.Title | Should -Be 'PSWriteExcel Set-Properties'
        $Properties.Subject | Should -Be 'PSWriteExcel'
    }    
}
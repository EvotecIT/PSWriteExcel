#Import-Module .\PSWriteExcel.psd1 -Force
#Import-Module 'C:\Users\przemyslaw.klys\OneDrive - Evotec\Support\GitHub\PSSharedGoods\PSSharedGoods.psd1' -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-Test1.xlsx"

$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover - Very long line I need to test that works"
    }
)

# Standard way
ConvertTo-Excel -DataTable $myitems0 -FilePath $FilePath -ExcelWorkSheetName 'This is my test' -AutoFilter -AutoFit -FreezeTopRow #-OpenWorkBook #-Verbose
# pipeline
$myitems0 | ConvertTo-Excel -FilePath $FilePath -ExcelWorkSheetName 'This is my test2' -AutoFilter -AutoFit -Option Skip -FreezeTopRow -OpenWorkBook
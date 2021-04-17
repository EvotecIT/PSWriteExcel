@{
    AliasesToExport      = @('Set-ExcelTranslateFromR1C1', 'Set-ExcelTranslateToR1C1', 'Set-ExcelWorkSheetCellStyleFont')
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2020 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'Little project to create Excel files without Microsoft Excel being installed.'
    FunctionsToExport    = @('Add-ExcelWorkSheet', 'Add-ExcelWorkSheetCell', 'Add-ExcelWorksheetData', 'ConvertFrom-Excel', 'ConvertTo-Excel', 'Excel', 'WorkbookProperties', 'Find-ExcelDocumentText', 'Get-ExcelDocument', 'Get-ExcelProperties', 'Get-ExcelTranslateFromR1C1', 'Get-ExcelTranslateToR1C1', 'Get-ExcelWorkSheet', 'Get-ExcelWorkSheetCell', 'Get-ExcelWorkSheetData', 'New-ExcelDocument', 'Remove-ExcelWorksheet', 'Request-ExcelWorkSheetCalculation', 'Save-ExcelDocument', 'Set-ExcelProperties', 'Set-ExcelWorksheetAutoFilter', 'Set-ExcelWorksheetAutoFit', 'Set-ExcelWorkSheetCellStyle', 'Set-ExcelWorkSheetFreezePane', 'Set-ExcelWorkSheetTableStyle', 'Worksheet')
    GUID                 = '82232c6a-27f1-435d-a496-929f7221334b'
    ModuleVersion        = '0.1.13'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            Tags       = @('Excel', 'ConvertTo-Excel', 'ExportExcel', 'macOS', 'linux', 'windows')
            ProjectUri = 'https://github.com/EvotecIT/PSWriteExcel'
            IconUri    = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWriteExcel.png'
        }
    }
    RequiredModules      = @(@{
            ModuleVersion = '0.0.190'
            ModuleName    = 'PSSharedGoods'
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
        })
    RootModule           = 'PSWriteExcel.psm1'
}
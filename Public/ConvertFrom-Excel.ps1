function ConvertFrom-Excel {
    [CmdletBinding()]
    param(
        [alias('Excel', 'Path')][string] $FilePath,
        [alias('WorksheetName', 'Name')][string] $ExcelWorksheetName
    )
    if (Test-Path $FilePath) {
        $ExcelDocument = Get-ExcelDocument -Path $FilePath
        if ($ExcelWorksheetName) {
            $ExcelWorksheet = Get-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $ExcelWorksheetName
            if ($ExcelWorksheet) {
                $Data = Get-ExcelWorkSheetData -ExcelDocument $ExcelDocument -ExcelWorkSheet $ExcelWorksheet
                return $Data
            } else {
                Write-Warning "ConvertFrom-Excel - Worksheet with name $ExcelWorksheetName doesn't exists. Conversion terminated."
            }
        }
    } else {
        Write-Warning "ConvertFrom-Excel - File $FilePath doesn't exists. Conversion terminated."
    }
}
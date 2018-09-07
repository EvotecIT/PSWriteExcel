function Save-ExcelDocument {
    param (
        [parameter(Mandatory = $true, ValueFromPipeline = $true)][Alias('Document', 'Excel', 'Package')] $ExcelDocument,
        [string] $FilePath,
        [alias('Show', 'Open')][switch] $OpenWorkBook
    )
    if ($Script:SaveCounter -gt 5) {
        Write-Warning "Save-ExcelDocument - Couldnt save Excel. Terminating.."
        return
    }
    try {
        Write-Verbose "Save-ExcelDocument - Saving workbook $FilePath"
        $ExcelDocument.SaveAs($FilePath)
        $Script:SaveCounter = 0
    } catch {
        $Script:SaveCounter++
        $ErrorMessage = $_.Exception.Message
        if ($ErrorMessage -like "*The process cannot access the file*because it is being used by another process.*" -or
            $ErrorMessage -like "*Error saving file*") {
            $FilePath = "$($([System.IO.Path]::GetTempFileName()).Split('.')[0]).xlsx"
            Write-Warning "Save-ExcelDocument - Couldn't save file as it was in use or otherwise. Trying different name $FilePath"
            $ExcelDocument.File = $FilePath
            Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $FilePath -OpenWorkBook:$OpenWorkBook
        } else {
            Write-Warning $ErrorMessage
        }
    }

    if ($OpenWorkBook) {
        if (Test-Path $FilePath) {
            Invoke-Item -Path $FilePath
        } else {
            Write-Warning "Save-ExcelDocument - File $FilePath doesn't exists. Can't open Excel document."
        }
    }
}
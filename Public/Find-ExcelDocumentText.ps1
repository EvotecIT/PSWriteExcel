function Find-ExcelDocumentText {
    [CmdletBinding()]
    param(
        [string] $FilePath,
        [string] $FilePathTarget,
        [string] $Find,
        [switch] $Replace,
        [string] $ReplaceWith,
        [switch] $Regex,
        [switch] $OpenWorkBook,
        [alias('Supress')][bool] $Suppress
    )
    $Excel = Get-ExcelDocument -Path $FilePath
    if ($Excel) {
        $Addresses = @()
        $ExcelWorksheets = $Excel.Workbook.Worksheets
        #$i = 1
        foreach ($WorkSheet in $ExcelWorksheets) {
            #Write-Color 'Worksheet ', $i -Color White, Red
            $StartRow = $WorkSheet.Dimension.Start.Row
            $StartColumn = $WorkSheet.Dimension.Start.Column
            $EndRow = $WorkSheet.Dimension.End.Row + 1
            $EndColumn = $WorkSheet.Dimension.End.Column + 1

            for ($Row = $StartRow; $Row -le $EndRow; $Row++) {
                for ($Column = $StartColumn; $Column -le $EndColumn; $Column++) {
                    #Write-Color -Text 'Row: ', $Row, ' Column: ', $Column -Color White, Green, White, Green
                    $Value = $Worksheet.Cells[$Column, $Row].Value
                    if ($Value -like "*$Find*") {
                        if ($Replace) {
                            if ($Regex) {
                                $Worksheet.Cells[$Column, $Row].Value = $Value -Replace $Find, $ReplaceWith
                            } else {
                                $Worksheet.Cells[$Column, $Row].Value = $Value.Replace($Find, $ReplaceWith)
                            }
                        }
                        $Addresses += $WorkSheet.Cells[$Column, $Row].FullAddress
                    }
                }
            }
            #$i++
        }
        if ($Replace) {
            Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePathTarget -OpenWorkBook:$OpenWorkBook
        }
        if ($Suppress) {
            return
        } else {
            return $Addresses
        }
    }
}
function Add-ExcelWorkSheet {
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [alias('Name')][string] $WorksheetName,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Skip',
        [alias('Supress')][bool] $Suppress
    )
    $WorksheetName = $WorksheetName.Trim()
    if ($WorksheetName.Length -eq 0) {
        $WorksheetName = Get-RandomStringName -Size 31
        Write-Warning "Add-ExcelWorkSheet - Name is empty. Generated random name: '$WorksheetName'"
    } elseif ($WorksheetName.Length -gt 31) {
        $WorksheetName = $WorksheetName.Substring(0, 31)
    }

    $PreviousWorksheet = Get-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $WorksheetName
    if ($PreviousWorksheet) {
        #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName already exists"
        if ($Option -eq 'Skip') {
            #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName - skipping"
            Write-Warning "Add-ExcelWorkSheet - Worksheet '$WorksheetName' already exists. Skipping creation of new worksheet. Option: $Option"
            #Write-Warning "Add-ExcelWorkSheet - You can overwrite this setting with one of the Options: Replace, Skip, Rename"
            $Data = $PreviousWorksheet
        } elseif ($Option -eq 'Replace') {
            Write-Verbose "Add-ExcelWorkSheet - WorksheetName: '$WorksheetName' - exists. Replacing worksheet with empty worksheet."
            Remove-ExcelWorksheet -ExcelDocument $ExcelDocument -ExcelWorksheet $PreviousWorksheet
            $Data = Add-ExcelWorkSheet -ExcelDocument $ExcelDocument -WorksheetName $WorksheetName -Option $Option -Suppress $False
        } elseif ($Option -eq 'Rename') {
            Write-Verbose "Add-ExcelWorkSheet - Worksheet: '$WorksheetName' already exists. Renaming worksheet to random value."
            $WorksheetName = Get-RandomStringName -Size 31
            $Data = Add-ExcelWorkSheet -ExcelDocument $ExcelDocument -WorksheetName $WorksheetName -Option $Option -Suppress $False
            Write-Verbose "Add-ExcelWorkSheet - New worksheet name $WorksheetName"
        } else {
            #Write-Verbose "Future use..."
        }
    } else {
        Write-Verbose "Add-ExcelWorkSheet - WorksheetName: '$WorksheetName' doesn't exists in Workbook. Continuing..."
        $Data = $ExcelDocument.Workbook.Worksheets.Add($WorksheetName)

        # $data.Workbook.Worksheets | fv
        #$Data | fv

        #if ($Data.Name -ne $WorksheetName) {
        #   Write-Warning "Add-ExcelWorkSheet - WorksheetName was changed from:'$WorksheetName' to new name: '$($Data.Name)' (max chars 31)."
        #Write-Warning "Add-ExcelWorkSheet - Maximum amount of chars is 31 for worksheet name"
        #}
    }
    if ($Suppress) { return } else { return $data }
}

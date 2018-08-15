

function Add-ExcelWorkSheet {
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [string] $Name,
        [string] $Option = 'Skip',
        [bool] $Supress
    )
    if ($Name.Length -gt 31) {
        $WorksheetName = $Name.Substring(0, 31)
    } else {
        $WorksheetName = $Name
    }

    $PreviousWorksheet = Get-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $WorksheetName
    if ($PreviousWorksheet) {
        #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName already exists"
        if ($Option -eq 'Skip') {
            #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName - skipping"
            Write-Warning "Worksheet $WorksheetName already exists. Skipping."
            Write-Warning "You can overwrite this setting with one of the Options: Delete, Skip, Rename"
            return
        } elseif ($Option -eq 'Replace') {
            #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName - replace"
            $ExcelDocument.Workbook.Worksheets.Delete($PreviousWorksheet)
            Add-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $Name -Option $Option -Supress $Supress
        } elseif ($Option -eq 'Rename') {
            #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName - rename"
        } else {
            #Write-Verbose "Future use..."
        }

    } else {
        Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName doesn't exists"
        $Data = $ExcelDocument.Workbook.Worksheets.Add($WorksheetName)

        if ($Data.Name -ne $WorksheetName) {
            Write-Warning "Name was changed from:'$Name' to new name: '$($Data.Name)'."
            Write-Warning "Maximum amount of chars is 31 for worksheet name"
        }
    }
    if ($Supress) { return } else { return $Data }
}

function Get-ExcelWorkSheet {
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [string] $Name
    )
    $Data = $ExcelDocument.Workbook.Worksheets | Where { $_.Name -eq $Name }
    return $Data
}
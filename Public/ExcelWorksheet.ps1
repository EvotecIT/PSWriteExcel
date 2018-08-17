function Add-ExcelWorkSheet {
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [alias('Name')][string] $WorksheetName,
        [string] $Option = 'Skip',
        [bool] $Supress
    )
    $WorksheetName = $WorksheetName.Trim()
    if ($WorksheetName.Length -eq 0) {
        $WorksheetName = -join ((48..57) + (97..122) | Get-Random -Count 31 | % {[char]$_})
        Write-Warning "Add-ExcelWorkSheet - Name is empty. Generated random name: $WorksheetName"
    } elseif ($WorksheetName.Length -gt 31) {
        $WorksheetName = $WorksheetName.Substring(0, 31)
    }

    $PreviousWorksheet = Get-ExcelWorkSheet -ExcelDocument $ExcelDocument -Name $WorksheetName
    if ($PreviousWorksheet) {
        #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName already exists"
        if ($Option -eq 'Skip') {
            #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName - skipping"
            Write-Warning "Add-ExcelWorkSheet - Worksheet $WorksheetName already exists. Skipping."
            Write-Warning "Add-ExcelWorkSheet - You can overwrite this setting with one of the Options: Delete, Skip, Rename"
            return
        } elseif ($Option -eq 'Replace') {
            Write-Verbose "Add-ExcelWorkSheet - WorksheetName: $WorksheetName - exists. Replacing..."
            Remove-ExcelWorksheet -ExcelDocument $ExcelDocument -ExcelWorksheet $PreviousWorksheet
            Add-ExcelWorkSheet -ExcelDocument $ExcelDocument -WorksheetName $WorksheetName -Option $Option -Supress $Supress
        } elseif ($Option -eq 'Rename') {
            #Write-Verbose "Add-ExcelWorkSheet - Name: $WorksheetName - rename"
        } else {
            #Write-Verbose "Future use..."
        }

    } else {
        Write-Verbose "Add-ExcelWorkSheet - WorksheetName: $WorksheetName doesn't exists in Workbook. Continuing..."
        $Data = $ExcelDocument.Workbook.Worksheets.Add($WorksheetName)

        if ($Data.Name -ne $WorksheetName) {
            Write-Warning "Add-ExcelWorkSheet - WorksheetName was changed from:'$WorksheetName' to new name: '$($Data.Name)'."
            Write-Warning "Add-ExcelWorkSheet - Maximum amount of chars is 31 for worksheet name"
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

function Add-ExcelWorkSheetCell {
    param(
        [OfficeOpenXml.ExcelWorksheet]  $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [Object] $CellValue,
        [bool] $Supress
    )
    if ($ExcelWorksheet) {
        $Type = Get-ObjectType $CellValue
        Switch ($CellValue) {
            { $_ -and $Type.ObjectTypeName -eq 'PSCustomObject' } {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                break
            }
            { $_ -and $Type.ObjectTypeName -eq 'Object[]' } {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue -join [System.Environment]::NewLine
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.WrapText = $true
                break
            }
            { $_ -is [DateTime]} {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'm/d/yy h:mm'
                break
            }
            { $_ -is [TimeSpan]} {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Numberformat.Format = 'hh:mm:ss'
                break
            }
            Default {
                $Data = $ExcelWorksheet.Cells[$CellRow, $CellColumn].Value = $CellValue
            }
        }

    }
    if ($Supress) { return } else { $Data }
}
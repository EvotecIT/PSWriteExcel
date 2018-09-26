function Get-ExcelWorkSheetData {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelPackage] $ExcelDocument,
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorkSheet
    )

    $Dimensions = $ExcelWorkSheet.Dimension
    $CellRow = 1

    $ExcelDataArray = @()

    $Headers = @() # 1st row
    for ($CellColumn = 1; $CellColumn -lt $Dimensions.Columns + 1; $CellColumn++) {
        $Heading = $ExcelWorkSheet.Cells[$CellRow, $CellColumn].Value
        if ([string]::IsNullOrEmpty($Heading)) {
            $Heading = $ExcelWorkSheet.Cells[$CellRow, $CellColumn].Address
        }
        $Headers += $Heading
    }
    Write-Verbose "Get-ExcelWorkSheetData - Headers: $($Headers -join ',')"

    for ($CellRow = 2; $CellRow -lt $Dimensions.Rows + 1; $CellRow++) {

        $ExcelData = [PsCustomObject] @{  }
        for ($CellColumn = 1; $CellColumn -lt $Dimensions.Columns + 1; $CellColumn++) {
            $ValueContent = $ExcelWorkSheet.Cells[$CellRow, $CellColumn].Value
            #$ValueContent
            $ColumnName = $Headers[$CellColumn - 1]

            # Write-Verbose "CellRow: $CellRow  CellColumn: $CellColumn ColumnName: $ColumnName ValueContent: $ValueContent"
            Add-Member -InputObject $ExcelData -MemberType NoteProperty -Name $ColumnName -Value $ValueContent
            $ExcelData.$ColumnName = $ValueContent
        }
        $ExcelDataArray += $ExcelData
    }
    return $ExcelDataArray
}
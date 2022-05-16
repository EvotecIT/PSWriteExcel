function Get-ExcelWorkSheetData {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelPackage] $ExcelDocument,
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorkSheet
    )

    $Dimensions = $ExcelWorkSheet.Dimension
    $CellRow = 1

    $Headers = [System.Collections.Generic.List[string]]::new()
    for ($CellColumn = 1; $CellColumn -lt $Dimensions.Columns + 1; $CellColumn++) {
        $Heading = $ExcelWorkSheet.Cells[$CellRow, $CellColumn].Value
        if ([string]::IsNullOrEmpty($Heading)) {
            $Heading = $ExcelWorkSheet.Cells[$CellRow, $CellColumn].Address
        }
        if ($Headers.Contains($Heading)) {
            $Heading = $Heading + "_" + $ExcelWorkSheet.Cells[$CellRow, $CellColumn].Address
        }
        $Headers.Add($Heading)
    }
    [Array] $ExcelDataArray = for ($CellRow = 2; $CellRow -lt $Dimensions.Rows + 1; $CellRow++) {
        $ExcelData = [ordered] @{ }
        for ($CellColumn = 1; $CellColumn -lt $Dimensions.Columns + 1; $CellColumn++) {
            $ValueContent = $ExcelWorkSheet.Cells[$CellRow, $CellColumn].Value
            $ColumnName = $Headers[$CellColumn - 1]
            $ExcelData[$ColumnName] = $ValueContent
        }
        [PSCustomObject] $ExcelData
    }
    $ExcelDataArray
}
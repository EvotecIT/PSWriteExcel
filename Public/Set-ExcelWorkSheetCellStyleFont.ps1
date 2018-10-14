function Set-ExcelWorkSheetCellStyleFont {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [nullable[bool]] $Bold,
        [nullable]$Color,
        $Family,
        $Italic,
        [string] $Name,
        $Scheme,
        [nullable[int]] $Size,
        $Strike,
        $UnderLine,
        # [underlineType] $UnderLineType,
        $VerticalAlign
    )
    if (-not $ExcelWorksheet) { return }

    if ($Bold) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Bold = $Bold
    }
    if ($Color) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Color = $Color
    }
    if ($Family) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Family = $Family
    }
    if ($Italic) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Italic = $Italic
    }
    if ($Name) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Name = $Name
    }
    if ($Scheme) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Scheme = $Scheme
    }
    if ($Size) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Size = $Size
    }
    if ($Strike) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Strike = $Strike
    }
    # if ($UnderLine) {
    #     $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.UnderLine = $UnderLine
    # }
    if ($UnderLineType) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.UnderLineType = $UnderLineType
    }
    if ($VerticalAlign) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.VerticalAlign = $VerticalAlign
    }
}

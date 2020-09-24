function Set-ExcelWorkSheetCellStyle {
    [alias('Set-ExcelWorkSheetCellStyleFont')]
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,
        [int] $CellRow,
        [int] $CellColumn,
        [alias('FontName')][string] $Name,
        [switch] $Bold,
        [System.Drawing.KnownColor] $Color,
        [switch] $Italic,
        [int] $Size,
        [switch] $Strike,
        [switch] $UnderLine,
        [OfficeOpenXml.Style.ExcelUnderLineType] $UnderLineType,
        [OfficeOpenXml.Style.ExcelVerticalAlignment] $VerticalAlignment,
        [OfficeOpenXml.Style.ExcelFillStyle] $PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    )
    if (-not $ExcelWorksheet) { return }

    if ($PatternType) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Fill.PatternType = $PatternType
    }
    if ($Bold) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Bold = $Bold.IsPresent
    }
    if ($Color) {
        $DrawingColor = [System.Drawing.Color]::FromKnownColor($Color)
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Fill.BackgroundColor.SetColor($DrawingColor)
    }
    if ($Italic) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Italic = $Italic
    }
    if ($Name) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Name = $Name
    }
    if ($Size) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Size = $Size
    }
    if ($Strike) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.Strike = $Strike
    }
    if ($UnderLine) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.UnderLine = $UnderLine
    }
    if ($UnderLineType) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.UnderLineType = $UnderLineType
    }
    if ($VerticalAlign) {
        $ExcelWorksheet.Cells[$CellRow, $CellColumn].Style.Font.VerticalAlign = $VerticalAlignment
    }
}

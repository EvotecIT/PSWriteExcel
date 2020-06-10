function Request-ExcelWorkSheetCalculation {
    [cmdletBinding(DefaultParameterSetName = 'ExcelWorkSheetName')]
    param(
        [Parameter(ParameterSetName = 'ExcelWorkSheet')][OfficeOpenXml.ExcelWorksheet] $ExcelWorksheet,

        [Parameter(ParameterSetName = 'ExcelWorkSheetName', Mandatory)]
        [Parameter(ParameterSetName = 'ExcelWorkSheetIndex', Mandatory)]
        [OfficeOpenXml.ExcelPackage] $Excel,
        [Parameter(ParameterSetName = 'ExcelWorkSheetName')][string] $Name,
        [Parameter(ParameterSetName = 'ExcelWorkSheetIndex')][int] $Index
    )
    if ($ExcelWorksheet) {
        [OfficeOpenXml.CalculationExtension]::Calculate($ExcelWorkSheet)
    } elseif ($Name -and $Excel) {
        $ExcelWorksheet = Get-ExcelWorkSheet -Name $Name -ExcelDocument $Excel
        if ($ExcelWorksheet) {
            [OfficeOpenXml.CalculationExtension]::Calculate($ExcelWorkSheet)
        }
    } else {
        if ($PSBoundParameters.Contains('Index') -and $Excel) {
            $ExcelWorksheet = Get-ExcelWorkSheet -Index $Index -ExcelDocument $Excel
            if ($ExcelWorksheet) {
                [OfficeOpenXml.CalculationExtension]::Calculate($ExcelWorkSheet)
            }

        }
    }
}
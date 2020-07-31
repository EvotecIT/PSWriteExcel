function Worksheet {
    [CmdletBinding()]
    param(
        [Array] $DataTable,
        [string] $Name,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
        [string] $TabColor,
        [switch] $AutoFilter,
        [switch] $AutoFit
    )
    $ScriptBlock = {
        Param (
            $ExcelDocument,
            [Array] $DataTable,
            [string] $Name,
            [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Replace',
            [string] $TabColor,
            [alias('Supress')][bool] $Suppress,
            [switch] $AutoFilter,
            [switch] $AutoFit
        )
        $addExcelWorkSheetDataSplat = @{
            DataTable          = $DataTable
            TabColor           = $TabColor
            Suppress            = $Suppress
            Option             = $Option
            ExcelDocument      = $ExcelDocument
            ExcelWorksheetName = $Name
            AutoFit            = $AutoFit
            AutoFilter         = $AutoFilter
        }
        Add-ExcelWorksheetData @addExcelWorkSheetDataSplat -Verbose
    }
    $ExcelWorkSheetParameters = [ordered] @{
        DataTable     = $DataTable
        TabColor      = $TabColor
        Suppress       = $true
        Option        = $Option
        ExcelDocument = $Script:Excel.ExcelDocument
        Name          = $Name
        AutoFit       = $AutoFit
        AutoFilter    = $AutoFilter
    }

    if ($Script:Excel.Runspaces.Parallel) {
        $RunSpace = Start-Runspace -ScriptBlock $ScriptBlock -Parameters $ExcelWorkSheetParameters -RunspacePool $Script:Excel.Runspaces.RunspacesPool -Verbose:$Verbose
        $Script:Excel.Runspaces.Runspaces.Add($RunSpace)
    } else {
        & $ScriptBlock -Parameters @ExcelWorkSheetParameters
    }
}
$ScriptBlockColors = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    $Script:RGBColors.Keys | Where-Object { $_ -like "$wordToComplete*" }
}

Register-ArgumentCompleter -CommandName Worksheet -ParameterName TabColor -ScriptBlock $ScriptBlockColors
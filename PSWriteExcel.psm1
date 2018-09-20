#Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )
if ($PSEdition -eq 'Core') {
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Microsoft.Extensions.*.dll -ErrorAction SilentlyContinue )
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\*.NetCORE.dll -ErrorAction SilentlyContinue )
} else {
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\*.Net40.dll -ErrorAction SilentlyContinue )
}
#Dot source the files
Foreach ($Import in @($Public + $Private)) {
    Try {
        . $Import.Fullname
    } Catch {
        Write-Error -Message "Failed to import function $($import.Fullname): $_"
    }
}
Foreach ($Import in @($Assembly)) {
    Try {
        #Write-Verbose "Importing assembly name $($Import.Fullname)"
        Add-Type -Path $Import.Fullname
    } Catch {
        Write-Error -Message "Failed to import DLL $($Import.Fullname): $_"
    }
}

Export-ModuleMember -Function 'Add-ExcelWorkSheet' , 'Add-ExcelWorkSheetCell' , 'Add-ExcelWorksheetData' , 'ConvertTo-Excel' , 'Get-ExcelDocument' , 'Get-ExcelWorkSheet' , 'New-ExcelDocument' , 'Remove-ExcelWorksheet' , 'Save-ExcelDocument' , 'Set-ExcelTranslateFromR1C1' , 'Set-ExcelWorksheetAutoFilter' , 'Set-ExcelWorksheetAutoFit' , 'Set-ExcelWorkSheetFreezePane'

[string] $ManifestFile = '{0}.psd1' -f (Get-Item $PSCommandPath).BaseName;
$ManifestPathAndFile = Join-Path -Path $PSScriptRoot -ChildPath $ManifestFile;
if ( Test-Path -Path $ManifestPathAndFile) {
    $Manifest = (Get-Content -raw $ManifestPathAndFile) | iex;
    foreach ( $ScriptToProcess in $Manifest.ScriptsToProcess) {
        $ModuleToRemove = (Get-Item (Join-Path -Path $PSScriptRoot -ChildPath $ScriptToProcess)).BaseName;
        if (Get-Module $ModuleToRemove) {
            Remove-Module $ModuleToRemove;
        }
    }
}
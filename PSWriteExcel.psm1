#Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )
#if ($PSEdition -eq 'Core') {
#   $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\*.NetCORE.dll -ErrorAction SilentlyContinue )
#} else {
$Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\*.NET40.dll -ErrorAction SilentlyContinue )
#}

#Dot source the files
Foreach ($import in @($Public + $Private)) {
    Try {
        . $import.fullname
    } Catch {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}
Foreach ($import in @($Assembly)) {
    Try {
        #Write-Warning "Importing assembly name $($Import.Fullname)"
        Add-Type -Path $import.fullname
    } Catch {
        Write-Error -Message "Failed to import DLL $($import.fullname): $_"
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
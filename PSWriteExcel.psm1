#Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue -Recurse )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue -Recurse )
if ($PSEdition -eq 'Core') {
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Core\Microsoft.Extensions.*.dll -ErrorAction SilentlyContinue )
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Core\*.NetCORE.dll -ErrorAction SilentlyContinue )
} else {
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Default\*.Net40.dll -ErrorAction SilentlyContinue )
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

Export-ModuleMember -Function '*' -Alias '*'
#Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue -Recurse )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue -Recurse )
if ($PSEdition -eq 'Core') {
    # $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Core\Microsoft.Extensions.*.dll -ErrorAction SilentlyContinue )
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Core\*.dll -ErrorAction SilentlyContinue )
} else {
    #$Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Default\System.Text*.dll -ErrorAction SilentlyContinue )
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Default\*.dll -ErrorAction SilentlyContinue )
}
Foreach ($Import in @($Assembly)) {
    try {
        Add-Type -Path $Import.Fullname -ErrorAction Stop
    } catch [System.Reflection.ReflectionTypeLoadException] {
        Write-Error -Message "Message: $($_.Exception.Message)"
        #Write-Error -Message "StackTrace: $($_.Exception.StackTrace)"
        #Write-Error -Message "LoaderExceptions: $($_.Exception.LoaderExceptions)"
    } catch {
        Write-Error -Message "Message: $($_.Exception.Message)"
        #Write-Error -Message "StackTrace: $($_.Exception.StackTrace)"
        #Write-Error -Message "LoaderExceptions: $($_.Exception.LoaderExceptions)"
    }


}
#Dot source the files
Foreach ($Import in @($Public + $Private)) {
    Try {
        . $Import.Fullname
    } Catch {
        Write-Error -Message "Failed to import function $($import.Fullname): $_"
    }
}

Export-ModuleMember -Function '*' -Alias '*'
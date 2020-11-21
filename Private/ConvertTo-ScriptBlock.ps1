function ConvertTo-ScriptBlock {
    [cmdletBinding()]
    param(
        [Array] $Code,
        [string[]] $Include,
        [string[]] $Exclude
    )
    if ($Include) {
        $Output = foreach ($Line in $Code) {
            foreach ($I in $Include) {
                if ($Line.StartsWith($I)) {
                    $Line
                }
            }
        }
    }
    if ($Exclude) {
        $Output = foreach ($Line in $Code) {
            $Tests = foreach ($E in $Exclude) {
                if ($Line.StartsWith($E)) {
                    $true
                }
            }
            if ($Tests -notcontains $true) {
                $Line
            }
        }
    }
    if ($Output) {
        [ScriptBlock]::Create($Output)
    }
}
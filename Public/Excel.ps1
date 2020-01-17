function Excel {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Excel requires opening and closing brace."),
        [string] $FilePath,
        [switch] $Open,
        [switch] $Parallel
    )
    $Time = Start-TimeLog
    $ExcelDocument = New-ExcelDocument

    $Script:Excel = @{}
    $Script:Excel.ExcelDocument = $ExcelDocument

    $Script:Excel.Runspaces = @{}
    $Script:Excel.Runspaces.Parallel = $Parallel.IsPresent
    $Script:Excel.Runspaces.RunspacesPool = New-RunSpace
    $Script:Excel.Runspaces.Runspaces = [System.Collections.Generic.List[PSCustomObject]]::new()

    # We need to make sure some commands are executed in correct order, therefore we convert scriptblock into text, get the commands that we need to be executed last (Main)
    # and build ScriptBlock back
    [Array] $Output = ConvertFrom-ScriptBlock -ScriptBlock $Content
    #Write-Verbose -Message $(Stop-TimeLog -time $Time -Continue)
    $WorkbookProperties = ConvertTo-ScriptBlock -Code $Output -Include 'WorkbookProperties'
    #Write-Verbose -Message $(Stop-TimeLog -time $Time -Continue)
    $Everything = ConvertTo-ScriptBlock -Code $Output -Exclude 'WorkbookProperties'
    #Write-Verbose -Message $(Stop-TimeLog -time $Time -Continue)

    if ($Everything) {
        & $Everything
        $Script:Excel.Runspaces.End = Stop-Runspace -Runspaces  $Script:Excel.Runspaces.Runspaces -FunctionName "Excel" -RunspacePool $Script:RunspacesPool -Verbose:$Verbose -ErrorAction SilentlyContinue -ErrorVariable +AllErrors -ExtendedOutput:$ExtendedOutputF
    }
    if ($WorkbookProperties) {
        & $WorkbookProperties
        #Write-Verbose -Message $(Stop-TimeLog -time $Time -Continue)
        <#
        [Array] $Parameters = Invoke-Command -ScriptBlock $Main
        foreach ($Parameter in $Parameters) {
            switch ( $Parameter.Type ) {
                WorkbookProperties {
                    $Splat = $Parameter.ExcelProperties
                    Set-ExcelProperties @Splat -ExcelDocument $ExcelDocument
                    break
                }
            }

        }
        #>
    }
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $FilePath -OpenWorkBook:$Open
    $Script:Excel = $null
    Write-Verbose "Excel - Time to create - $EndTime"
}
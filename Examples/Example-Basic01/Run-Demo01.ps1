Import-Module $PSScriptRoot\..\..\PSWriteExcel.psd1 -Force #-Verbose

$Process = Get-Process | Select-Object -First 5

Excel -FilePath $PSScriptRoot\"Run-Demo01.xlsx" {
    WorkbookProperties -Title 'Test'
    Worksheet -DataTable $Process -Name 'Processes'

    Worksheet -DataTable $Process -Name 'Processes Test1' -TabColor Crimson
    Worksheet -DataTable $Process -Name 'Processes Test2' -TabColor BlueViolet
    Worksheet -DataTable $Process -Name 'Processes Test3' -TabColor Aquamarine
    Worksheet -DataTable $Process -Name 'Processes Test43' -TabColor Aquamarine
    Worksheet -DataTable $Process -Name 'Processes Test5' -TabColor Aquamarine
    Worksheet -DataTable $Process -Name 'Processes Test6' -TabColor Aquamarine
    Worksheet -DataTable $Process -Name 'Processes Test7' -TabColor Aquamarine
    Worksheet -DataTable $Process -Name 'Processes Test7' -TabColor Aquamarine
    Worksheet -DataTable $Process -Name 'Processes Test9' -TabColor Aquamarine
    Worksheet -DataTable $Process -Name 'Processes Test10' -TabColor Aquamarine

    for ($i = 0; $i -le 500; $i++) {
        #Worksheet -DataTable $Process -Name "Processes Test $i" -TabColor BlanchedAlmond
    }

} -Verbose -Open
# PSWriteExcel - PowerShell Module

[![Build status APPVEYOR](https://ci.appveyor.com/api/projects/status/n3ds0y45vc2dv6r2?svg=true)](https://ci.appveyor.com/project/PrzemyslawKlys/pswriteexcel)
[![Build Status Azure DevOps](https://dev.azure.com/evotecpl/PSWriteExcel/_apis/build/status/EvotecIT.PSWriteExcel)](https://dev.azure.com/evotecpl/PSWriteExcel/_build/latest?definitionId=1)
[![Build Status](https://travis-ci.org/EvotecIT/PSWriteExcel.svg?branch=master)](https://travis-ci.org/EvotecIT/PSWriteExcel)

[PSWriteExcel](https://evotec.xyz/hub/scripts/pswriteexcel-powershell-module/) is very basic (at the moment) PowerShell module to create Microsoft Excel workbooks without Microsoft Excel installed. It's main purpose is to support/create [Excel for PSWinDocumentation](https://evotec.xyz/hub/scripts/pswindocumentation-powershell-module/) project. Depending on requirements of this project (and maybe few others) it may evolve or cover more feature sets. After some basic testing it seems to work fine on **Windows** and **Linux** and **MacOS**.

More information can be found on dedicated page for [PSWriteExcel](https://evotec.xyz/hub/scripts/pswriteexcel-powershell-module/) module.

## There are 2 ways to use this module

-   Long way - `New-ExcelDocument`, `Add-ExcelWorkSheet`, `Add-ExcelWorksheetData` and finally `Save-ExcelDocument`
-   Short way - `$Object | ConvertTo-Excel -Path 'Export.xlsx' -WorkSheetName 'MyName'`

There are couple of more commands in play that may come useful. Feel free to explore.

## Example usage of Add-ExcelWorksheetData in action

![image](https://evotec.xyz/wp-content/uploads/2018/08/PSWriteExcel.gif.pagespeed.ce.WKvsf00WoC.gif)

## Changelog

-   0.1.2 - 23.06.2019
    -   Support for PSSharedGoods 0.0.79
    -   Some speed improvments
-   0.1.0 - 17.04.2019
    -   Big Performance improvements, removed some reduntant calls
    -   Updated .DLL to newest version (compiled from Source Code on day 15.04.2019 with all changes/fixes in EPPlus)
-   0.0.17 - 22.03.2019
    -   Added -PreScanHeaders to ConvertTo-Excel - objects are prescanned first so that property names are known before exporting
-   0.0.16 - 15.02.2019
    -   Added -TableStyle ConvertTo-Excel
    -   Added -TableStyle Add-ExcelWorksheetData

## Installing PowerShell Core on (Linux - Ubuntu)

```bash
# Download the Microsoft repository GPG keys
wget -q https://packages.microsoft.com/config/ubuntu/16.04/packages-microsoft-prod.deb
# Register the Microsoft repository GPG keys
sudo dpkg -i packages-microsoft-prod.deb
# Update the list of products
sudo apt-get update
# Install PowerShell
sudo apt-get install -y powershell
# Start PowerShell
pwsh
```

For anything else refer to great Microsoft Article - [Installing PowerShell Core on Linux](https://docs.microsoft.com/en-US/powershell/scripting/setup/installing-powershell-core-on-linux?view=powershell-6)

## Installing on Windows / Linux / MacOS

```powershell
Install-Module PSWriteExcel -Scope CurrentUser
#Update-Module PSWriteExcel
```

## Using on Linux

```
Get-Process | ConvertTo-Excel -Path 'ThisIsMyExcel.xlsx' -WorkSheetName 'AndWorksheet' -AutoFilter
```

![image](https://evotec.xyz/wp-content/uploads/2018/09/PSWriteExcel-ExportOnUbuntu.gif)

## Credits

This module is based on [EPPlus](https://github.com/JanKallman/EPPlus) and it's doing all the magic behind this project. PSWriteExcel is merely a wrapper around that with few PowerShell tricks around converting objects into tables.

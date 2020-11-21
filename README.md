<p align="center">
  <a href="https://dev.azure.com/evotecpl/PSWriteExcel/_build/results?buildId=latest"><img src="https://dev.azure.com/evotecpl/PSWriteExcel/_apis/build/status/EvotecIT.PSWriteExcel"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteExcel"><img src="https://img.shields.io/powershellgallery/v/PSWriteExcel.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteExcel"><img src="https://img.shields.io/powershellgallery/vpre/PSWriteExcel.svg?label=powershell%20gallery%20preview&colorB=yellow"></a>
  <a href="https://github.com/EvotecIT/PSWriteExcel"><img src="https://img.shields.io/github/license/EvotecIT/PSWriteExcel.svg"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/PSWriteExcel"><img src="https://img.shields.io/powershellgallery/p/PSWriteExcel.svg"></a>
  <a href="https://github.com/EvotecIT/PSWriteExcel"><img src="https://img.shields.io/github/languages/top/evotecit/PSWriteExcel.svg"></a>
  <a href="https://github.com/EvotecIT/PSWriteExcel"><img src="https://img.shields.io/github/languages/code-size/evotecit/PSWriteExcel.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteExcel"><img src="https://img.shields.io/powershellgallery/dt/PSWriteExcel.svg"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>

# PSWriteExcel - PowerShell Module

[PSWriteExcel](https://evotec.xyz/hub/scripts/pswriteexcel-powershell-module/) is very basic (at the moment) PowerShell module to create Microsoft Excel workbooks without Microsoft Excel installed. It's main purpose is to support/create [Excel for PSWinDocumentation](https://evotec.xyz/hub/scripts/pswindocumentation-powershell-module/) project. Depending on requirements of this project (and maybe few others) it may evolve or cover more feature sets. After some basic testing it seems to work fine on **Windows** and **Linux** and **MacOS**.

More information can be found on dedicated page for [PSWriteExcel](https://evotec.xyz/hub/scripts/pswriteexcel-powershell-module/) module.

## There are 2 ways to use this module

- Long way - `New-ExcelDocument`, `Add-ExcelWorkSheet`, `Add-ExcelWorksheetData` and finally `Save-ExcelDocument`
- Short way - `$Object | ConvertTo-Excel -Path 'Export.xlsx' -WorkSheetName 'MyName'`

There are couple of more commands in play that may come useful. Feel free to explore.

## Example usage of Add-ExcelWorksheetData in action

![image](https://evotec.xyz/wp-content/uploads/2018/08/PSWriteExcel.gif.pagespeed.ce.WKvsf00WoC.gif)

## Changelog

- 0.1.12 - 2020.11.21
  - Added back missing cmdlet which would prevent `Excel` cmdlet from working
- 0.1.11 - 2020.09.24
  - `Set-ExcelWorkSheetCellStyle` is now usable, it wasn't even half working before.
- 0.1.10 - 2020.07.31
  - Fix misspelling of "suppress" as "supress" (finally!) - tnx natescherer [#7](https://github.com/EvotecIT/PSWriteExcel/pull/7)
- 0.1.9 - 2020.07.30
  - Added verification for cell lenght to prevent errors. Cell will be trimmed to 32767 chars when lenght exceeds that.
- 0.1.8 - 2020.06.21
  - Small improvement
- 0.1.7 - 2020.06.21
  - Fix for not displaying $False and few other values
- 0.1.6 - 2020.06.10
  - Fix for colors, Colors should limit output while typing
  - Added `Request-ExcelWorkSheetCalculation`
  - Added ability to add CellFormula to `Add-ExcelWorkSheetCell`
  - Fix for `Transpose` in `ConvertTo-Excel`

- 0.1.5 - 2020.01.17
  - Merged `Excelimo` back into `PSWriteExcel`
  - Merged all dependencies into `PSWriteExcel` - requires additional modules only for development like all my other modules

- 0.1.2 - 23.06.2019
  - Support for PSSharedGoods 0.0.79
  - Some speed improvments
- 0.1.0 - 17.04.2019
  - Big Performance improvements, removed some reduntant calls
  - Updated .DLL to newest version (compiled from Source Code on day 15.04.2019 with all changes/fixes in EPPlus)
- 0.0.17 - 22.03.2019
  - Added -PreScanHeaders to ConvertTo-Excel - objects are prescanned first so that property names are known before exporting
- 0.0.16 - 15.02.2019
  - [x]   Added -TableStyle ConvertTo-Excel
  - [x] Added -TableStyle Add-ExcelWorksheetData

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

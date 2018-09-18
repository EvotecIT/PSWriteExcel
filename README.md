# PSWriteExcel - PowerShell Module

[![Build status](https://ci.appveyor.com/api/projects/status/n3ds0y45vc2dv6r2?svg=true)](https://ci.appveyor.com/project/PrzemyslawKlys/pswriteexcel)

*PSWriteExcel* is very basic (at the moment) PowerShell module to create Microsoft Excel workbooks without Microsoft Excel installed.

Overview of this module: https://evotec.xyz/hub/scripts/pswriteexcel-powershell-module/

It's main purpose is to support/create Excel for PSWinDocumentation (https://evotec.xyz/hub/scripts/pswindocumentation-powershell-module/) project. Depending on requirements of this project (and maybe few others) it may evolve or cover more feature sets.

## Example usage of Add-ExcelWorksheetData in action

![image](https://evotec.xyz/wp-content/uploads/2018/08/PSWriteExcel.gif.pagespeed.ce.WKvsf00WoC.gif)

## Using on PowerShell Core (Linux)

#Install PowerShell like: https://docs.microsoft.com/en-US/powershell/scripting/setup/installing-powershell-core-on-linux?view=powershell-6
#Install-Package Microsoft.Extensions.Configuration -Version 2.1.1

## Credits

This module is based on [EPPlus](<https://github.com/JanKallman/EPPlus>) and it's doing all the magic behind this project. PSWriteExcel is merely a wrapper around that with few PowerShell tricks around converting objects into tables.
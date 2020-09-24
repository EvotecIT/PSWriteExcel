---
external help file: PSWriteExcel-help.xml
Module Name: PSWriteExcel
online version:
schema: 2.0.0
---

# Request-ExcelWorkSheetCalculation

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

### ExcelWorkSheetName (Default)
```
Request-ExcelWorkSheetCalculation -Excel <ExcelPackage> [-Name <String>] [<CommonParameters>]
```

### ExcelWorkSheet
```
Request-ExcelWorkSheetCalculation [-ExcelWorksheet <ExcelWorksheet>] [<CommonParameters>]
```

### ExcelWorkSheetIndex
```
Request-ExcelWorkSheetCalculation -Excel <ExcelPackage> [-Index <Int32>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Excel
{{ Fill Excel Description }}

```yaml
Type: ExcelPackage
Parameter Sets: ExcelWorkSheetName, ExcelWorkSheetIndex
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelWorksheet
{{ Fill ExcelWorksheet Description }}

```yaml
Type: ExcelWorksheet
Parameter Sets: ExcelWorkSheet
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Index
{{ Fill Index Description }}

```yaml
Type: Int32
Parameter Sets: ExcelWorkSheetIndex
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
{{ Fill Name Description }}

```yaml
Type: String
Parameter Sets: ExcelWorkSheetName
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS

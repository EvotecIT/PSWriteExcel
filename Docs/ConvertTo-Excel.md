---
external help file: PSWriteExcel-help.xml
Module Name: PSWriteExcel
online version:
schema: 2.0.0
---

# ConvertTo-Excel

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
ConvertTo-Excel [[-FilePath] <String>] [[-Excel] <ExcelPackage>] [[-ExcelWorkSheetName] <String>]
 [[-DataTable] <Object>] [[-Option] <String>] [-AutoFilter] [-AutoFit] [-FreezeTopRow] [-FreezeFirstColumn]
 [-FreezeTopRowFirstColumn] [[-FreezePane] <Int32[]>] [-Transpose] [[-TransposeSort] <String>]
 [[-TableStyle] <TableStyles>] [[-TableName] <String>] [-OpenWorkBook] [<CommonParameters>]
```

## DESCRIPTION
{{Fill in the Description}}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -AutoFilter
{{Fill AutoFilter Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFit
{{Fill AutoFit Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: Autosize

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataTable
{{Fill DataTable Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases: TargetData

Required: False
Position: 3
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Excel
{{Fill Excel Description}}

```yaml
Type: ExcelPackage
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelWorkSheetName
{{Fill ExcelWorkSheetName Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases: Name, WorksheetName

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FilePath
{{Fill FilePath Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases: path

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FreezeFirstColumn
{{Fill FreezeFirstColumn Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FreezePane
{{Fill FreezePane Description}}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FreezeTopRow
{{Fill FreezeTopRow Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FreezeTopRowFirstColumn
{{Fill FreezeTopRowFirstColumn Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OpenWorkBook
{{Fill OpenWorkBook Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Option
{{Fill Option Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: Replace, Skip, Rename

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
{{Fill TableName Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle
{{Fill TableStyle Description}}

```yaml
Type: TableStyles
Parameter Sets: (All)
Aliases: TableStyles
Accepted values: None, Custom, Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10, Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20, Light21, Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10, Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19, Medium20, Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28, Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Transpose
{{Fill Transpose Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: Rotate, RotateData, TransposeColumnsRows, TransposeData

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TransposeSort
{{Fill TransposeSort Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: ASC, DESC, NONE

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.Object

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS

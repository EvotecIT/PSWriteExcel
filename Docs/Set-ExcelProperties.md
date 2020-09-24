---
external help file: PSWriteExcel-help.xml
Module Name: PSWriteExcel
online version:
schema: 2.0.0
---

# Set-ExcelProperties

## SYNOPSIS
Following function allows setting properties for Excel Workbook. It works with standard Excel Properties.

## SYNTAX

```
Set-ExcelProperties [[-ExcelDocument] <ExcelPackage>] [[-Title] <String>] [[-Subject] <String>]
 [[-Author] <String>] [[-Comments] <String>] [[-Keywords] <String>] [[-LastModifiedBy] <String>]
 [[-LastPrinted] <String>] [[-Created] <DateTime>] [[-Category] <String>] [[-Status] <String>]
 [[-Application] <String>] [[-HyperlinkBase] <String>] [[-AppVersion] <String>] [[-Company] <String>]
 [[-Manager] <String>] [[-Modified] <DateTime>] [[-LinksUpToDate] <Boolean>] [[-HyperlinksChanged] <Boolean>]
 [[-ScaleCrop] <Boolean>] [[-SharedDoc] <Boolean>] [<CommonParameters>]
```

## DESCRIPTION
Following function allows setting properties for Excel Workbook. It works with standard Excel Properties.

## EXAMPLES

### Example 1
```powershell
Import-Module PSWriteExcel -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteExcel-Example-SetProperties.xlsx"

$Excel = New-ExcelDocument -Verbose

$ExcelWorkSheet = Add-ExcelWorkSheet -ExcelDocument $Excel -WorksheetName 'Test 1' -Suppress $False -Option 'Replace'

$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason another one"; age = 42; info = "Food lover"
    }
)
Add-ExcelWorksheetData -ExcelWorksheet $ExcelWorkSheet -DataTable $myitems0 -AutoFit -AutoFilter -Suppress $True

Set-ExcelProperties -ExcelDocument $Excel -Author 'Przemyslaw Klys' -Title 'This is a test'
Set-ExcelProperties -ExcelDocument $Excel -Comments 'Testing PSWriteExcel' -Subject 'Subject'

Save-ExcelDocument -ExcelDocument $Excel -FilePath $FilePath -OpenWorkBook
```

Following example creates simple Excel file, adds some content to it and sets properties.

## PARAMETERS

### -AppVersion
{{Fill AppVersion Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 13
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Application
{{Fill Application Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Author
{{Fill Author Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Category
{{Fill Category Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Comments
{{Fill Comments Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Company
{{Fill Company Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 14
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Created
{{Fill Created Description}}

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelDocument
{{Fill ExcelDocument Description}}

```yaml
Type: ExcelPackage
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HyperlinkBase
{{Fill HyperlinkBase Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 12
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HyperlinksChanged
{{Fill HyperlinksChanged Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 18
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Keywords
{{Fill Keywords Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastModifiedBy
{{Fill LastModifiedBy Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastPrinted
{{Fill LastPrinted Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LinksUpToDate
{{Fill LinksUpToDate Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 17
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manager
{{Fill Manager Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 15
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Modified
{{Fill Modified Description}}

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

Required: False
Position: 16
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScaleCrop
{{Fill ScaleCrop Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 19
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SharedDoc
{{Fill SharedDoc Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 20
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Status
{{Fill Status Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
{{Fill Subject Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
{{Fill Title Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
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

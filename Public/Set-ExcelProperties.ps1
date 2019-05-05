function Set-ExcelProperties {
    [CmdletBinding()]
    param(
        [OfficeOpenXml.ExcelPackage] $ExcelDocument,
        [string] $Title,
        [string] $Subject,
        [string] $Author,
        [string] $Comments,
        [string] $Keywords,
        [string] $LastModifiedBy,
        [string] $LastPrinted,
        [nullable[DateTime]] $Created,
        [string] $Category,
        [string] $Status,
        [string] $Application,
        [string] $HyperlinkBase,
        [string] $AppVersion,
        [string] $Company,
        [string] $Manager,
        [nullable[DateTime]] $Modified,
        [nullable[bool]] $LinksUpToDate,
        [nullable[bool]] $HyperlinksChanged,
        [nullable[bool]] $ScaleCrop,
        [nullable[bool]] $SharedDoc
        #[hashtable] $CustomProperty,
        #[hashtable] $ExtendedProperty
    )
    if ($Title) {
        $ExcelDocument.Workbook.Properties.Title = $Title
    }
    if ($Subject) {
        $ExcelDocument.Workbook.Properties.Subject = $Subject
    }
    if ($Author) {
        $ExcelDocument.Workbook.Properties.Author = $Author
    }
    if ($Comments) {
        $ExcelDocument.Workbook.Properties.Comments = $Comments
    }
    if ($Keywords) {
        $ExcelDocument.Workbook.Properties.Keywords = $Keywords
    }
    if ($LastModifiedBy) {
        $ExcelDocument.Workbook.Properties.LastModifiedBy = $LastModifiedBy
    }
    if ($LastPrinted) {
        $ExcelDocument.Workbook.Properties.LastPrinted = $LastPrinted
    }
    if ($Created) {
        $ExcelDocument.Workbook.Properties.Created = $Created
    }
    if ($Category) {
        $ExcelDocument.Workbook.Properties.Category = $Category
    }
    if ($Status) {
        $ExcelDocument.Workbook.Properties.Status = $Status
    }
    if ($Application) {
        $ExcelDocument.Workbook.Properties.Application = $Application
    }
    if ($HyperlinkBase) {
        if ($HyperlinkBase -like '*://*') {
            $ExcelDocument.Workbook.Properties.HyperlinkBase = $HyperlinkBase
        } else {
            Write-Warning "Set-ExcelProperties - Hyperlinkbase is not an URL (doesn't contain ://)"
        }
    }
    if ($AppVersion) {
        $ExcelDocument.Workbook.Properties.AppVersion = $AppVersion
    }
    if ($Company) {
        $ExcelDocument.Workbook.Properties.Company = $Company
    }
    if ($Manager) {
        $ExcelDocument.Workbook.Properties.Manager = $Manager
    }
    if ($Modified) {
        $ExcelDocument.Workbook.Properties.Modified = $Modified
    }
    if ($LinksUpToDate -ne $null) {
        $ExcelDocument.Workbook.Properties.LinksUpToDate = $LinksUpToDate
    }
    if ($HyperlinksChanged -ne $null) {
        $ExcelDocument.Workbook.Properties.HyperlinksChanged = $HyperlinksChanged
    }
    if ($ScaleCrop -ne $null) {
        $ExcelDocument.Workbook.Properties.ScaleCrop = $ScaleCrop
    }
    if ($SharedDoc -ne $null) {
        $ExcelDocument.Workbook.Properties.SharedDoc = $SharedDoc
    }
    #foreach ($Key in $Custom.Keys) {
    #    $ExcelDocument.Workbook.Properties.SetCustomPropertyValue($Key, $Custom.$Key)
    #}
    #foreach ($Key in $ExtendedProperty.Keys) {
    #    $ExcelDocument.Workbook.Properties.SetExtendedPropertyValue($Key, $ExtendedProperty.$Key)
    #}
}
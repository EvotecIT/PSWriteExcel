function WorkbookProperties {
    [CmdletBinding()]
    param(
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
    )

    <#
    $Object = @{
        Type            = 'WorkbookProperties'
        ExcelProperties = @{
            HyperlinksChanged = $HyperlinksChanged
            ScaleCrop         = $ScaleCrop
            HyperlinkBase     = $HyperlinkBase
            Subject           = $Subject
            LastModifiedBy    = $LastModifiedBy
            Author            = $Author
            LinksUpToDate     = $LinksUpToDate
            Modified          = $Modified
            LastPrinted       = $LastPrinted
            Company           = $Company
            Comments          = $Comments
            Title             = $Title
            SharedDoc         = $SharedDoc
            Created           = $Created
            Category          = $Category
            #ExcelDocument     = $Script:ExcelDocument
            Status            = $Status
            AppVersion        = $AppVersion
            Keywords          = $Keywords
            Application       = $Application
            Manager           = $Manager
        }
    }
    #return $Object
    #>

    $ExcelProperties = @{
        HyperlinksChanged = $HyperlinksChanged
        ScaleCrop         = $ScaleCrop
        HyperlinkBase     = $HyperlinkBase
        Subject           = $Subject
        LastModifiedBy    = $LastModifiedBy
        Author            = $Author
        LinksUpToDate     = $LinksUpToDate
        Modified          = $Modified
        LastPrinted       = $LastPrinted
        Company           = $Company
        Comments          = $Comments
        Title             = $Title
        SharedDoc         = $SharedDoc
        Created           = $Created
        Category          = $Category
        ExcelDocument     = $Script:Excel.ExcelDocument
        Status            = $Status
        AppVersion        = $AppVersion
        Keywords          = $Keywords
        Application       = $Application
        Manager           = $Manager
    }
    Set-ExcelProperties @ExcelProperties

}
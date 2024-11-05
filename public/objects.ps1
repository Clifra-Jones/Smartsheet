function New-SmartSheetCell() {
    Param(
        [Parameter(Mandatory = $true)]
        [UInt64]$columnId,
        [string]$conditionalFormat,
        [string]$format,
        [string]$formula,
        [psobject]$hyperlink,
        [psobject]$image,
        [psobject]$linkInFromCell,
        [psobject[]]$linksOutFromCell,
        [Parameter(Mandatory)]
        [psobject]$value
    )

    $properties = @{
        columnId = $columnId
        value = $value
        #displayValue = $value
    }
    if ($conditionalFormat) { $properties.Add("conditionalFormat", $conditionalFormat) }
    if ($format) { $properties.Add("format", $format) }
    if ($formula) { $properties.Add("formula", $formula) }
    if ($hyperlink) { $properties.Add("hyperlink", $hyperlink) }
    if ($image) { $properties.Add("image", $image) }
    if ($linkInFromCell) { $properties.Add("linkInFromCell", $linkInFromCell) }
    if ($linksOutFromCell) { $properties.Add("linksOutFromCell", $linksOutFromCell) }
    $cell = [PSCustomObject]$properties
    return $cell
    <#
    .SYNOPSIS
    Creates a new Smartsheet Cell object
    .PARAMETER columnId
    Column ID of the cell
    .PARAMETER conditionalFormat
    A conditional format object.
    .PARAMETER format
    A format descriptor string.
    .PARAMETER formula
    A formula string
    .PARAMETER hyperlink
    A hyperlink object
    .PARAMETER image
    An image object.
    .PARAMETER linkInFromCell
    A cell link object
    .PARAMETER linksOutFromCell
    A cell link object
    .PARAMETER value
    The value of the cell
    .OUTPUTS
    A smartsheet cell object.
    #>
}

function New-Hyperlink() {
    Param(
        [Parameter(ParameterSetName = "reportId")]
        [Uint64]$reportId = 0,
        [Parameter(ParameterSetName = "sheetId")]
        [Uint64]$sheetId = 0,
        [Parameter(ParameterSetName = "sightId")]
        [Uint64]$sightId = 0,
        [Parameter(ParameterSetName = "url")]
        [string]$url=""
    )

    $hyperlink = [PSCustomObject]@{
        reportId = $reportId
        sheetId = $sheetId
        sightId = $sightId
        url = $url
    }

    return $hyperlink
    <#
    .SYNOPSIS
    Creates a new Smartsheet Hyperlink object.
    .PARAMETER reportId
    Target report Id.
    .PARAMETER sheetId
    Target sheet Id.
    .PARAMETER sightId
    Target sight od.
    .PARAMETER url
    Target URL.
    .OUTPUTS
    A hyperlink object.
    #>
}

function New-CellLink() {
    Param(
        [Parameter(Mandatory = $true)]
        [Uint64]$sheetId,
        [Parameter(Mandatory = $true)]
        [string]$sheetName,
        [Parameter(Mandatory = $true)]
        [Uint64]$columnId,
        [Parameter(Mandatory = $true)]
        [Uint64]$rowId
    )

    $cellLink = [PSCustomObject]@{
        columnId    = $columnId
        rowId       = $rowId
        sheetId     = $sheetid
        $sheetName  = $sheetName
    }
    return $cellLink
    <#
    .SYNOPSIS
    Creates a new cell link object.
    This method only creates the CellLink object that can be later inserted into a Cell. Set the LinkInFromCell property to this object.
    .PARAMETER sheetId
    Target Sheet Id.
    .PARAMETER sheetName
    Target Sheet name.
    .PARAMETER columnId
    Target column Id.
    .PARAMETER rowId
    Target row Id.
    .OUTPUTS
    A cell link object.
    #>
}

#Helper functions
function New-SmartSheetFormatString() {
    Param(
        [ValidateSet("Arial", "Tahoma", "Times New Roman", "Verdana")]
        [string]$fontFamily,
        [int]$fontSize,
        [switch]$bold,
        [switch]$italic,
        [switch]$underline,
        [switch]$stikethrough,
        [ValidateSet("left", "center", "right")]
        [string]$horizontalAlign,
        [ValidateSet("top", "middle", "bottom")]
        [string]$verticalAlign,
        [ValidateSet("Navajo White", "Black", "White", "White_2", "Lavender blush", "Sazerac", "Chilean Heath", "Panache", "Solitude", "French Lilac", "Merino", "Pastel Pink", "Navajo White_2", "Dolly", "Zanah", "French Pass", "French Lilac_2", "Almond", "Mercury", "Froly", "Chardonnay", "Yellow", "De York", "Malibu", "Light Wisteria", "Tan", "Silver", "Cinnabar", "Pizazz", "Turbo", "Chateau Green", "Denim", "Seance", "Brown", "Sonic Silver", "Tamarillo", "Trinidad", "Corn", "Forest Green", "Catalina Blue", "Purple", "Carnaby Tan")]
        [string]$textColor,
        [ValidateSet("Navajo White", "Black", "White", "White_2", "Lavender blush", "Sazerac", "Chilean Heath", "Panache", "Solitude", "French Lilac", "Merino", "Pastel Pink", "Navajo White_2", "Dolly", "Zanah", "French Pass", "French Lilac_2", "Almond", "Mercury", "Froly", "Chardonnay", "Yellow", "De York", "Malibu", "Light Wisteria", "Tan", "Silver", "Cinnabar", "Pizazz", "Turbo", "Chateau Green", "Denim", "Seance", "Brown", "Sonic Silver", "Tamarillo", "Trinidad", "Corn", "Forest Green", "Catalina Blue", "Purple", "Carnaby Tan")]
        [string]$backgroundColor,
        [ValidateSet("Navajo White", "Black", "White", "White_2", "Lavender blush", "Sazerac", "Chilean Heath", "Panache", "Solitude", "French Lilac", "Merino", "Pastel Pink", "Navajo White_2", "Dolly", "Zanah", "French Pass", "French Lilac_2", "Almond", "Mercury", "Froly", "Chardonnay", "Yellow", "De York", "Malibu", "Light Wisteria", "Tan", "Silver", "Cinnabar", "Pizazz", "Turbo", "Chateau Green", "Denim", "Seance", "Brown", "Sonic Silver", "Tamarillo", "Trinidad", "Corn", "Forest Green", "Catalina Blue", "Purple", "Carnaby Tan")]
        [string]$taskbarColor,
        [ValidateSet("RUB", "SEK", "AUD", "CAD", "KRW", "USD", "ARS", "NZD", "NOK", "INR", "EUR", "ILS", "SGD", "CNY", "none", "DKK", "BRL", "GBP", "HKD", "JPY", "CLP", "MXN", "CHF", "ZAR")]
        [string]$currency,
        [int]$decimalCount,
        [switch]$thousandsSeparator,
        [ValidateSet("none", "NUMBER", "CURRENCY", "PERCENT")]
        [string]$numberFormat,
        [switch]$textWrap,
        [ValidateSet("LOCALE_BASED", "MMMM_D_YYYY", "MMM_D_YYYY", "D_MMM_YYYY", "YYYY_MM_DD_HYPHEN", "YYYY_MM_DD_DOT", "DWWWW_MMMM_D_YYYY", "DWWW_DD_MMM_YYYY", "DWWW_MM_DD_YYYY", "MMMM_D", "D_MMMM")]
        [string]$dateFormat
    )
    $format = [ordered]@{
        fontFamily         = $null
        fontSize           = $null
        bold               = $null
        italic             = $null
        underlined         = $null
        strikethrough      = $null
        horizontalAlign    = $null
        verticalAlign      = $null
        textcolor          = $null
        backgroundColor    = $null
        taskbarColor       = $null
        currency           = 13
        decimalCount       = 2
        thousandsSeparator = $null
        numberFormat       = $null
        textWrap           = $null
        dateFormat         = $null
    }
    if ($fontFamily) { $format.fontFamily = $smformat.fontFamilies[$fontFamily] }
    if ($fontSize) { $format.fontSize = $fontSize }
    if ($bold) { $format.bold = 1 }
    if ($italic) { $format.italic = 1 }
    if ($underline) { $format.underlined = 1 }
    if ($stikethrough) { $format.strikethrough = 1 }
    if ($horizontalAlign) { $format.horizontalAlign = $smformat.horizontalAlign[$horizontalAlign] }
    if ($verticalAlign) { $format.verticalAlign = $smformat.verticalAlign[$verticalAlign] }
    if ($textColor) { $format.textcolor = $smformat.colors[$textColor] }
    if ($backgroundColor) { $format.backgroundColor = $smformat.colors[$backgroundColor] }
    if ($taskbarColor) { $format.backgroundColor = $smformat.colors[$backgroundColor] }
    if ($currency) { $format.currency = $smformat.currency[$currency] }
    if ($decimalCount) { $format.decimalCount = $decimalCount }
    if ($thousandsSeparator) { $format.thousandsSeparator = 1 }
    if ($numberFormat) { $format.numberFormat = $smformat.numberFormats[$numberFormat] }
    if ($textWrap) { $format.textWrap = 1 }
    if ($dateFormat) { $format.dateFormat = $smformat.dateFormats[$dateFormat] }
    
    $formatString = $format.values -join ","
    return $formatString

<#
    .SYNOPSIS 
    Creates a SMartsheet format string. Supports: autocomplete.
    .DESCRIPTION
    Creates a smartsheet format string to be used with column, row, and cell formatting.
    .PARAMETER fontFamily
    Sets the Font Family to use.
    .PARAMETER fontSize
    Sets the font size.
    .PARAMETER bold
    Sets to font to bold.
    .PARAMETER italic
    Sets the font to italic
    .PARAMETER underline
    Sets the font to underline
    .PARAMETER stikethrough
    Sets the font to strikethrough.
    .PARAMETER horizontalAlign
    Set the horizontal alignment
    .PARAMETER verticalAlign
    Set the vertical alignment
    .PARAMETER textColor
    Select the Text Color. Supports : autocomplete.
    .PARAMETER backgroundColor
    Select the Background color. Supports : autocomplete.
    .PARAMETER taskbarColor 
    Select the Taskbar color. Supports : autocomplete.
    .PARAMETER currency
    Select the Currency Symbol. Supports : autocomplete.
    .PARAMETER decimalCount
    Set the decimal count
    .PARAMETER thousandsSeparator
    Sets the thousands separator.
    .PARAMETER numberFormat
    Sets the Number format. Supports: autocomplete.
    .PARAMETER textWrap
    Sets textwrap.
    .PARAMETER dateFormat
    sets the date format. Supports : autocomplete.
    .OUTPUTS
    A string representing a Smartsheet formatting string.
#>
}
function New-SmartSheetCell() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$columnId,
        [string]$conditionalFormat,
        [string]$format,
        [string]$formula,
        [psobject]$hyperlink,
        [psobject]$image,
        [psobject]$linkInFromCell,
        [psobject[]]$linksOutFromCell,
        [psobject]$value
    )

    $properties = @{
        columnId = $columnId
    }
    if ($conditionalFormat) { $properties.Add("conditionalFormat", $conditionalFormat) }
    if ($format) { $properties.Add("format", $format) }
    if ($formula) { $properties.Add("formula", $formula) }
    if ($hyperlink) { $properties.Add("hyperlink", $hyperlink) }
    if ($image) { $properties.Add("image", $image) }
    if ($linkInFromCell) { $properties.Add("linkInFromCell", $linkInFromCell) }
    if ($linksOutFromCell) { $properties.Add("linksOutFromCell", $linksOutFromCell) }
    if ($value) { $properties.Add("value", $value) }
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
    A format descriptor sctring.
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
        [string]$reportId = 0,
        [Parameter(ParameterSetName = "sheetId")]
        [string]$sheetId = 0,
        [Parameter(ParameterSetName = "sightId")]
        [string]$sightId = 0,
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
    #>
}

function New-CellLink() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$sheetid,
        [Parameter(Mandatory = $true)]
        [string]$sheetName,
        [Parameter(Mandatory = $true)]
        [string]$columnId,
        [Parameter(Mandatory = $true)]
        [string]$rowId
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
    Creates a new cell link object,
    .PARAMETER sheetid
    Target Sheet Id.
    .PARAMETER sheetName
    Target Sheet name.
    .PARAMETER columnId
    Target column Id.
    .PARAMETER rowId
    Target row Id.
    #>
}
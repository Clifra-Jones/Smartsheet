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
}
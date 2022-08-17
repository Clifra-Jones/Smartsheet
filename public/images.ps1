
function Add-SmartSheetCellImage() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$rowId,
        [Parameter(Mandatory = $true)]
        [string]$columnId,
        [string]$altText,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $Uri = "{0}/sheets/{1}/rows/{2}/columns/{3}/cellimages" -f $BaseURI, $id, $rowId, $columnId

    if ($altText) {
        $Uri = "{0}?altText={1}" -f $Uri, $altText
    }

    $file = Get-Item -Path $Path
    $mimeType = $mimetypes[$file.Extension]

    $Headers = Get-Headers -ContentType $mimeType -ContentDisposition 'attachment' -filename $file.Name
    #$Headers.Add('Content-Length',$file.Length)

    [byte[]]$bytes = [System.IO.File]::ReadAllBytes($path)

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $bytes
        if ($response.message -eq 'SUCCESS') {
            return $response.result
        } else {
            throw $response.message
        }
    } catch {
        throw $_
    }
    <#
        .SYNOPSIS
        Adds an image to a cell
        .DESCRIPTION
        Added an image to the specified cell reference.
        .PARAMETER Id
        Smartsheet Id
        .PARAMETER rowId
        Smartsheet row Id.
        .PARAMETER columnId
        Smartsheet column Id.
        .PARAMETER altText
        ALternative text associated with the image.
        .PARAMETER Path
        Path tot he local image file.
        .OUTPUTS
        Smartsheet Row object that contains the image.
    #>
}

function Get-SmartsheetImageUrl() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('imageId')]
        [string]$Id,
        [string]$saveAs
    )
    $Headers = Get-Headers

    $Uri = "{0}/imageurls" -f $BaseURI

    $payload = @(
        @{
            imageId = $id
        }
    )
    $body = $payload | ConvertTo-Json -AsArray -Compress

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($saveAs) {
            [void](Invoke-WebRequest -Uri $response.Url -OutFile $saveAs)
        } else {
            return $response.imageUrls
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Returns Url to download the image.
    .PARAMETER Id
    Id of the image to get the Url for.
    .PARAMETER saveAs
    Path and filename to save the image to.
    .OUTPUTS
    Smartsheet Url image object containing imageId and Url.
    #>
}
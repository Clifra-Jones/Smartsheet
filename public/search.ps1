function Search-SmartsheetAccount() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$searchText,
        [switch]$exact
    )

    $Headers = Get-Headers -AuthOnly

    if ($exact) {
        $searchText = '"{0}"' -f $searchText
    }
    $Uri = '{0}/search?query={1}' -f $BaseURI, $searchText
    
    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.results
    } catch {
        throw $_
    }
}

function Search-Smartsheet() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$searchText,
        [switch]$exact
    )

    $Headers = Get-Headers -AuthOnly

    If ($exact) {
        $searchText = '"{0}"' -f $searchText
    }
    $Uri = "{0}/search/sheets/{1}?query={2}" -f $BaseURI, $Id, $searchText

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.results
    } catch {
        throw $_
    }
}
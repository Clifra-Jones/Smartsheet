function Search-SmartsheetAccount() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$searchText,
        [switch]$exact,
        [switch]$personalWorkspaces,
        [datetime]$modifiedSince,
        [switch]$favoriteFlag,
        [string[]]$scopes
    )

    $Headers = Get-Headers -AuthOnly

    if ($exact) {
        $searchText = '"{0}"' -f $searchText
    }
    $Uri = '{0}/search?query={1}' -f $BaseURI, $searchText
    if ($personalWorkspaces.IsPresent) {
        $Uri = "{0}&location={1}" -f $Uri, "personalWoekspaces"
    }
    if ($modified) {
        $Uri = "{0}&modifiedSince={1}" -f $Url, ($modifiedSince.tostring("s"))
    }
    if ($favoriteFlag) {
        $Url = "{0}&include{1}" -f $Url, "favoriteFlag"
    }
    if ($scopes) {
        $Url = "{0}&scopes={1}" -f $Uri, ($scopes -join ",")
    }
    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.results
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Searches all sheets that the user can access, for the specified text.

    #>
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
    <#
    .SYNOPSIS 
    Search a Smartsheet
    .DESCRIPTION
    Gets a list of the user's search results in a sheet based on query. 
    The list contains an abbreviated row object for each search result in a sheet. 
    Note Newly created or recently updated data may not be immediately discoverable via search.
    .PARAMETER Id
    Sheet ID of the sheet to search.
    .PARAMETER searchText
    Text to search for
    .PARAMETER exact
    Match text exactly.
    .OUTPUTS
    An array of search results.
    
    #>
}
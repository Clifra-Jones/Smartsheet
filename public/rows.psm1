function Add-SmartsheetRow() {
    [CmdletBinding(DefaultParameterSetName = "props")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = "props"
        )]
        [Parameter(ParameterSetName="row")]
        [string]$Id,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "row"
        )]
        [Row]$Row,
        [bool]$expanded,
        [string]$format,
        [psobject[]]$Cells,
        [bool]$locked,
        [switch]$top,
        [string]$aboveRow,
        [string]$belowRow
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $null
    if ($Row) {
        $body = $Row | ConvertTo-Json -Compress
    } else {
        $properties = @{}
        if ($top) { 
            $properties.Add("toTop", $true)
        }elseif ($aboveRow) {
            $properties.Add("siblingId". $aboveRow)
            $properties.Add("above", $true)
        } elseIf($belowRow) {
            $properties.Add("sibling", $belowRow)
        }
        if ($expanded) { $properties.Add("Expamded", $expanded)}
        if ($format) { $properties.Add("format", $format) }
        if ($locked) { Properties.Add("locked", $locked )}
        If ($Cells) {$properties.Add("cells", $Cells)}

        $row = [psCustpmObject]$properties
        
        $body = $Row | ConvertTo-Json -Compress
    }
    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $response.result
    } else {
        return $false
    }
}

function Add-SmartsheetRows() {
    Param(
        [Parameter(
            Mandatory = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [Row[]]$Rows
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $id

    $body = $Rows | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq 'SUCCESS') {
        return $response.result
    } else {
        return $false
    }
}
function Remove-SmartsheetRow() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$rowId,
        [bool]$ignoreRowsNotFound
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows?ids={2}" -f $BaseURI, $Id, $rowId

    If ($ignoreRowsNotFound) {
        $Uri = $Uri + "&ignoreRowsNotFound=true"
    }

    $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
    if ($response.message -eq "SUCCESS") {
        return r$true
    } else {
        return $false
    }
}

function Remove-SmartsheetRows() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string[]]$rowIds,
        [bool]$ignoreRowsNotFound
    )

    $Headers = Get-Headers
    $Ids = $rowIds -join ','

    $Uri = "{0}/sheets/{1}/rows?ids={2}" -f $BaseURI, $Id, $Ids

    If ($ignoreRowsNotFound) {
        $Uri = $Uri + "&ignoreRowsNotFound=true"
    }

    $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
    if ($response.message -eq "SUCCESS") {
        return r$true
    } else {
        return $false
    }
}

function Set-SmartsheetRow() {
    [CmdletBinding(DefaultParameterSetName = "props")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = "props"
        )]
        [Parameter(ParameterSetName = "row")]
        [string]$Id,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "row"
        )]
        [psobject]$Row,
        [bool]$expanded,
        [string]$format,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "props"
        )]
        [psobject[]]$Cells,
        [bool]$locked
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $null
    if ($Row) {
        $body = $Row | ConvertTo-Json -Compress
    } else {
        $row = [Row]::New()
        if ($expanded) { $Row.expanded = $expanded }
        if ($format) { $Row.format = $format }
        if ($locked) { $Row.locked = $locked}
        $Row.Cell = $Cells
    }
    $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $true
    } else {
        retuen $false
    }
}

function Set-SmartsheetRows() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Row[]]$Rows
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $Rows | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $true
    } else {
        return $false
    }
}

function Get-SmartsheetRow() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$rowId,
        [switch]$includeColumns
    )
    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows/{2}" -f $BaseURI, $Id, $rowId
    if ($includeColumns) {
        $Uri = $Uri + "?include=columns"
    }

    $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers ./.git 

    return $response
}
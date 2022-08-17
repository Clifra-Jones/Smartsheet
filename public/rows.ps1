using namespace System.Collections.Generic

function New-SmartsheetRow() {
    [CmdletBinding(DefaultParameterSetName = "props")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'row'
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'props'
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,            
            ParameterSetName = 'top'
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'top2'
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'above'
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'above2'
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'below'
        )]
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'below2'
        )]
        [string]$Id,
        [Parameter(Mandatory = $true, ParameterSetName = "row")]
        [Parameter(Mandatory = $true, ParameterSetName = 'top2')]
        [Parameter(Mandatory = $true, ParameterSetName = 'above2')]
        [Parameter(Mandatory = $true, ParameterSetName = 'below2')]
        [psObject]$Row,        
        [Parameter(ParameterSetName = 'props')]        
        [Parameter(ParameterSetName = 'top')]
        [Parameter(ParameterSetName = 'above')]
        [Parameter(ParameterSetName = 'below')]
        [bool]$expanded,
        [Parameter(ParameterSetName = 'props')]
        [Parameter(ParameterSetName = 'top')]
        [Parameter(ParameterSetName = 'above')]
        [Parameter(ParameterSetName = 'below')]
        [string]$format,
        [Parameter(ParameterSetName = 'top')]
        [Parameter(ParameterSetName = 'above')]
        [Parameter(ParameterSetName = 'below')]
        [psobject[]]$cells,
        [Parameter(ParameterSetName = 'props')]
        [Parameter(ParameterSetName = 'top')]
        [Parameter(ParameterSetName = 'above')]
        [Parameter(ParameterSetName = 'below')]
        [bool]$locked, 
        [Parameter(ParameterSetName = 'top')]
        [Parameter(ParameterSetName = "top2")]
        [switch]$top,
        [Parameter(ParameterSetName = 'above')]
        [Parameter(ParameterSetName = 'above2')]
        [string]$aboveRow,
        [Parameter(ParameterSetName = 'below')]
        [Parameter(ParameterSetName = 'below2')]
        [string]$belowRow #>
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $null
    if ($Row) {
        $body = $Row | ConvertTo-Json -Compress
    } else {
        $payload = [ordered]@{}
        if ($top) { 
            $payload.Add("toTop", $true)
        }elseif ($aboveRow) {
            $payload.Add("siblingId". $aboveRow)
            $payload.Add("above", $true)
        } elseIf($belowRow) {
            $payload.Add("sibling", $belowRow)
        }
        if ($expanded) { $payload.Add("Expamded", $expanded)}
        if ($format) { $payload.Add("format", $format) }
        if ($locked) { $payload.Add("locked", $locked )}
        If ($cells) {$payload.Add("cells", $cells)}
        
        $body = $payload | ConvertTo-Json -Compress
    }
    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $response.result
    } else {
        return $false
    }
    <#
    .SYNOPSIS
    Add a Smartsheet row,
    .DESCRIPTION
    Add a row to a smartsheet. The default location is the bottom of the sheet,
    .PARAMETER Id
    Id of the sheet to add row to.
    .PARAMETER Row
    A row object to add to the sheet. Cannot be used with individual row properties.
    .PARAMETER expanded
    Indicates whether the row is expanded or collapsed.
    .PARAMETER format
    Format descriptor. Use New-SmartsheetFormatString to create format descriptors.
    .PARAMETER Cells
    Cells belonging to the row.
    .PARAMETER locked
    Indicates whether the row is locked.
    .PARAMETER top
    place the new row at the top of the sheet (Cannot be used with belowRow or aboveRow)
    .PARAMETER aboveRow
    Place the new row above the row ID assigned to this parameter (cannot be used with top or belowRow).
    .PARAMETER belowRow
    Place the new row below the row ID assigned to this parameter (cannot be used with top or aboveRow).
    .OUTPUTS
    The newly added row object.
    #>
}

function New-SmartsheetRows() {
    Param(
        [Parameter(
            Mandatory = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [psobject[]]$Rows
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
    <#
    .SYNOPSIS
    Add rows to a smartsheet
    .DESCRIPTION
    Add an array of row objects to a smartsheet.
    .PARAMETER Id
    The Id of the smartsheet to add the rows to.
    .PARAMETER Rows
    An array of smartsheet row objects
    .OUTPUTS
    An array of the newly created rows.
    #>
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
    <#
    .SYNOPSIS
    Remove a Smartsheet Row
    .DESCRIPTION
    Remove a row from a smartsheet.
    .PARAMETER Id
    ID of Smartsheet to remove the row,
    .PARAMETER rowId
    The rowID of the row to remove.
    .PARAMETER ignoreRowsNotFound
    Supress errors if row not found.
    .OUTPUTS
    Boolean indicating success or failure
    #>
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
    <#
    .SYNOPSIS
    Remove a Smartsheet Rows
    .DESCRIPTION
    Remove rows from a smartsheet.
    .PARAMETER Id
    ID of Smartsheet to remove the rows,
    .PARAMETER rowIds
    An array of rowIDs to be remove.
    .PARAMETER ignoreRowsNotFound
    Supress errors if row not found.
    .OUTPUTS
    Boolean indicating success or failure    
    #>
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
        [Parameter(ParameterSetName = 'props')]
        [bool]$expanded,
        [Parameter(ParameterSetName = 'props')]
        [string]$format,
        [Parameter(ParameterSetName = "props")]
        [psobject[]]$Cells,
        [Parameter(ParameterSetName = 'props')]
        [bool]$locked
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $null
    if ($Row) {
        $body = $Row | ConvertTo-Json -Compress
    } else {
        $payload = [ordered]@{}
        if ($expanded) { $payload.Add("expanded", $expanded) }
        if ($format) { $payload.Add("format", $format) }
        if ($locked) { $payload.Add("locked", $locked) }
        If ($Cells) { $Payload.Add("Cells", $Cells)}        
        $body = $payload | ConvertTo-Json -Compress
    }
    
    try {
        $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            return $response.result
        } else {
            throw $response.message
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Updates a Smartsheet row.
    .DESCRIPTION
    Updates the properties of a smartsheet row.
    .PARAMETER Id
    Id os the Smartsheet to update.
    .PARAMETER Row
    A Smartsheet row object containing the updates (cannot be used with individual properties).
    .PARAMETER expanded
    True if the row is expanded, false if not.
    .PARAMETER format
    Format descriptor. Only returned if the include query string parameter contains format and this row has a non-default format applied.
    .PARAMETER Cells
    An array of Smartsheet cell objects.
    .PARAMETER locked
    Indicates if the row is locked or not.
    .OUTPUTS
    Boolean indicating suncess or failure.
    #>
}

function Set-SmartsheetRows() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [psobject[]]$Rows
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $Rows | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            return $response.result
        } else {
            throw $response.message
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Update multiple Smartsheet rows.
    .DESCRIPTION
    UPdate multiple rows in a Smartsheet.
    .PARAMETER Id
    ID of the Smartsheet to update.
    .PARAMETER Rows
    An array of smartsheet row objects. 
    .OUTPUTS
    Boolean indicating success or failure
    #>
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
    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers 

        return $response
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    retrieve a Smartsheet row.
    .DESCRIPTION
    Retrieve a row from a smartsheet.
    .PARAMETER Id
    Id of the Smartsheet to get the row from.
    .PARAMETER rowId
    Id of the row to get.
    .PARAMETER includeColumns
    Include column objects in the returned object.
    .OUTPUTS
    A row object (optionally with column objects).

    #>
}

function Export-SmartsheetRows() {
    Param(
        [Parameter(
            ValueFromPipeline = $true
        )]
        [psobject]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$sheetId,
        [switch]$blankRowAbove,
        [string]$title,
        [string]$titleFormat,
        [switch]$includeHeaders,
        [string]$headerFormat
    )

    # Get current sheet Columns
    Begin{
        $Columns = Get-SmartsheetColumns -SheetId $sheetId   

        if ($blankRowAbove) {
            [void](Add-SmartsheetRow -sheetId $sheetId)
        }

        if ($title) {
            $cell = New-SmartSheetCell -columnId $columns[0].Id -value $title -format $titleFormat
            $cells = @()
            $cells += $cell
            [void](Add-SmartsheetRow -Id $sheetId -cells $cells)
        }

        
    }
    Process{
        $PropCount = $inputObject.PSObject.Properties.Count
        # Get the number of properties in the input object
        # if needed add columns to the sheet.
        if ($PropCount -gt $Columns.Count) {
            $n = $PropCount - $Columns.Count
            1..$n | ForEach-Object {
                $index = (columns.count -1) + $_
                [void](Add-SmartsheetColumn -Id $sheetId -index $index -type:TEXT_NUMBER)
            }
        }

        if ($includeHeaders){
            # Add the headings
            $propNames = ($inputObject[0].psObject.Properties).Name
            $cells = @()
            foreach ($propName in $propNames) {
                $i = $PropNames.IndexOf($propName)
                $cell = new-SmartsheetCell -columnId $Columns[$i].Id -value $propName -format $headerFormat
                $cells += $cell
            }
            [void](Add-SmartsheetRow -Id $sheetId -cells $cells)        
        }

        $values = ($inputObject.PSObject.Properties).value
        $cells = @()
        foreach($value in $values) {
            $i = $values.IndexOf($value)
            $cell = New-SmartSheetCell -columnId $columns[$i].Id -value $value
            $cells += $cell            
        }
        [void](Add-SmartsheetRow -Id $sheetId -cells $cells)
    }

    <#
    .SYNOPSIS
    Export an array and appends to a smartsheet.
    .DESCRIPTION
    Exports a Powershell array and appends new rows to a smartsheet.
    If no columns ecist in the smartsheet they are created as generic Columns, i.e. Column1, Column2.    
    To generate a smartsheet with named columns from the objects of the array use Export-Smartsheet.
    .PARAMETER InputObject
    An array of Powershell objects.
    .PARAMETER sheetId
    The Smartsheet ID to put the data in.
    .PARAMETER blankRowAbove
    INsert a blank row above the data being exported.
    .PARAMETER title
    Insert a title row above the data.
    .PARAMETER titleFormat
    A Smartsheet format string for the title. To create a format string use New-SmartsheetFormatString.
    .PARAMETER includeHeaders
    Create a header row from the property names from the objects in the array.
    .PARAMETER headerFormat
    A Smartsheet format string for the headers. To create a format string use New-SmartsheetFormatString.
    .OUTPUTS
    Nothing
    #>
}
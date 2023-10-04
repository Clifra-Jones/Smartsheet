using namespace System.Collections.Generic

function Add-SmartsheetRow() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true, ParameterSetName = "row")]
        [psObject]$Row,        
        [Parameter(ParameterSetName = 'props')]        
        [bool]$expanded,
        [Parameter(ParameterSetName = 'props')]
        [string]$format,
        [Parameter(ParameterSetName = 'props')]
        [psobject[]]$cells,
        [Parameter(ParameterSetName = 'props')]
        [bool]$locked,
        [ValidateSet('top','bottom','above','below')]
        [string]$location = 'bottom',
        [string]$siblingRowId,
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $null
    if ($Row) {
        $body = $Row | ConvertTo-Json -Compress
    } else {
        $payload = [ordered]@{}
        if ($location -eq 'top') { 
            $payload.Add("toTop", $true)
        }elseif ($location -eq 'above') {
            if (-not $siblingRowId) {
                throw "siblingRowId must be provided when specifying location 'above'!"                
            }
            $payload.Add("siblingId". $siblingRowId)
            $payload.Add("above", $true)
        } elseIf($location -eq 'below') {
            $payload.Add("siblingId", $siblingRowId)
        }
        if ($expanded) { $payload.Add("expanded", $expanded)}
        if ($format) { $payload.Add("format", $format) }
        if ($locked) { $payload.Add("locked", $locked )}
        If ($cells) {$payload.Add("cells", $cells)}
        
        $body = $payload | ConvertTo-Json -Compress
    }

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            if ($PassThru) {
                return Get-Smartsheet -id $id
            } else {
                return $response.result
            }
        } else {
            return $response.message
        }
    } catch {
        throw $_
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
    .PARAMETER location
    The location to insert the row. Default is 'bottom'.
    .PARAMETER siblingRowId
    If location is above of below the row ID to insert the ro above or below. Required when specifying 'above' or 'below' for location.
    .PARAMETER PassThru
    REturn the updated sheet.
    .OUTPUTS
    The newly added row object.
    if PassThru is specified, return the updated sheet object.
    #>
}

function Add-SmartsheetRows() {
    Param(
        [Parameter(
            Mandatory = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [psobject[]]$Rows,
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $id

    $body = $Rows | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq 'SUCCESS') {
            if ($PassThru) {
                return Get-Smartsheet -id $id
            } else {
                return $response.result
            }
        } else {
            throw $response.message
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Add rows to a smartsheet
    .DESCRIPTION
    Add an array of row objects to a smartsheet.
    .PARAMETER Id
    The Id of the smartsheet to add the rows to.
    .PARAMETER Rows
    An array of smartsheet row objects.
    .PARAMETER PassThru
    Return then update sheet.
    .OUTPUTS
    An array of the newly created rows.
    If PassThru is specified returns the updated sheet object.
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
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows?ids={2}" -f $BaseURI, $Id, $rowId

    try{
        $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
        if ($response.message -eq "SUCCESS") {
            if ($PassThru) {
                return Get-SmartSheet -id $Id
            } else {
                return $true
            }
        } else {
            return $response.message
        }
    } catch {
        throw $_
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
    .PARAMETER PassThru
    Return the updated sheet.
    .OUTPUTS
    True is delete was successful.
    if PassThru is specified returns the updated sheet object.
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
        [bool]$ignoreRowsNotFound,
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Ids = $rowIds -join ','

    $Uri = "{0}/sheets/{1}/rows?ids={2}" -f $BaseURI, $Id, $Ids

    If ($ignoreRowsNotFound) {
        $Uri = $Uri + "&ignoreRowsNotFound=true"
    }

    try {
        $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
        if ($response.message -eq "SUCCESS") {
            if ($PassThru) {
                return Get-Smartsheet -id $Id
            } else {
                return $true
            }
        } else {
            throw $response.message
        }
    } catch {
        throw $_
    }
        <#
    .SYNOPSIS
    Remove Smartsheet Rows
    .DESCRIPTION
    Remove rows from a smartsheet.
    .PARAMETER Id
    ID of Smartsheet to remove the rows,
    .PARAMETER rowIds
    An array of rowIDs to be remove.
    .PARAMETER ignoreRowsNotFound
    Suppress errors if row not found.
    .PARAMETER PassThru
    Returns the updated sheet.
    .OUTPUTS
    True if successful.
    if PassThru is specified, return the updated smartsheet object.
    #>
}

function Set-SmartsheetRow() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,        
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "row"
        )]
        [psobject]$Row,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'props'
        )]
        [string]$rowId,
        [Parameter(ParameterSetName = 'props')]
        [bool]$expanded,
        [Parameter(ParameterSetName = 'props')]
        [string]$format,
        [Parameter(ParameterSetName = "props")]
        [psobject[]]$Cells,
        [Parameter(ParameterSetName = 'props')]
        [bool]$locked,
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $null
    if ($Row) {
        $body = $Row | ConvertTo-Json -Compress
    } else {
        $payload = [ordered]@{
            id = $rowId
        }
        if ($expanded) { $payload.Add("expanded", $expanded) }
        if ($format) { $payload.Add("format", $format) }
        if ($locked) { $payload.Add("locked", $locked) }
        If ($Cells) { $Payload.Add("cells", $Cells)}        
        $body = $payload | ConvertTo-Json -Compress
    }
    
    try {
        $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            if ($PassThru) {
                return Get-Smartsheet -id $id
            } else {
                return $response.result
            }
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
    .PARAMETER rowId
    Row ID of the row to be updated.
    .PARAMETER expanded
    True if the row is expanded, false if not.
    .PARAMETER format
    Format descriptor. Only returned if the include query string parameter contains format and this row has a non-default format applied.
    .PARAMETER Cells
    An array of Smartsheet cell objects.
    .PARAMETER locked
    Indicates if the row is locked or not.
    .PARAMETER PassThru
    Return the updated sheet
    .OUTPUTS
    Boolean indicating success or failure.
    if PassThru is specified, return the updated sheet object.
    #>
}

function Set-SmartsheetRows() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [psobject[]]$Rows,
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/rows" -f $BaseURI, $Id

    $body = $Rows | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            if ($PassThru) {
                return Get-Smartsheet -id $id
            } else {
                return $response.$result
            }
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
    .PARAMETER PassThru
    Return the updated sheet.
    .OUTPUTS
    An array of updated rows.
    If PassThru is specified, returns the updated sheet object,
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
        [string]$headerFormat,
        [switch]$PassThru
    )

    # Get current sheet Columns
    Begin {
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
        $HeadersNotSet = $true
    }   

    Process {
        $PropCount = $inputObject.PSObject.Properties.Name.Count
        # Get the number of properties in the input object
        # if needed add columns to the sheet.
        if ($PropCount -gt $Columns.Count) {
            $n = $PropCount - $Columns.Count
            1..$n | ForEach-Object {
                $index = ($Columns.count -1) + $_
                $title = "Column_$Index"
                [void](Add-SmartsheetColumn -Id $sheetId -index $index -type:TEXT_NUMBER -title $title)
            }
        }

        #Get the updated columns
        $Columns = Get-SmartsheetColumns -SheetId $sheetId

        if ($includeHeaders){
            # Add the headings
            if ($HeadersNotSet) {
                $propNames = ($inputObject[0].psObject.Properties).Name
                $cells = @()
                foreach ($propName in $propNames) {
                    $i = $PropNames.IndexOf($propName)
                    $cell = new-SmartsheetCell -columnId $Columns[$i].Id -value $propName -format $headerFormat
                    $cells += $cell
                }
                [void](Add-SmartsheetRow -Id $sheetId -cells $cells)     
                $HeadersNotSet = $False
            }
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

    End {
        if ($PassThru) {
            return Get-Smartsheet -id $sheetId
        }
    }

    <#
    .SYNOPSIS
    Export an array and appends to a smartsheet.
    .DESCRIPTION
    Exports a Powershell array and appends new rows to a smartsheet.
    If no columns exist in the smartsheet they are created as generic Columns, i.e. Column1, Column2.    
    To generate a smartsheet with named columns from the objects of the array use Export-Smartsheet.
    .PARAMETER InputObject
    An array of Powershell objects.
    .PARAMETER sheetId
    The Smartsheet ID to put the data in.
    .PARAMETER blankRowAbove
    Insert a blank row above the data being exported.
    .PARAMETER title
    Insert a title row above the data.
    .PARAMETER titleFormat
    A Smartsheet format string for the title. To create a format string use New-SmartsheetFormatString.
    .PARAMETER includeHeaders
    Create a header row from the property names of the objects in the array.
    .PARAMETER headerFormat
    A Smartsheet format string for the headers. To create a format string use New-SmartsheetFormatString.
    .PARAMETER PassThru
    Return the sheet object with the inserted rows,
    .OUTPUTS
    If -PassThru is omitted nothing is returned.
    If specifying -PassThru the sheet object is returned with the inserted rows,
    .NOTES
    This function is generally used to create the equivalent of an Excel table in a Smartsheet.
    This is sort of "out of functionality" for how Smartsheets work, but some may find it useful.
    You can use this function to append rows to an existing Smartsheet. See example 2 below.
    Data will always be added starting at the left most column. If you wanted to insert data into a Smartsheet
    starting at a certain column you would need to use the Add-SmartSheetRow function inserting blank cells into the beginning of
    the cell array.
    .EXAMPLE 
    The following example imports the array into a Smartsheet, creates a blank row above the data and adds a title and a header row.
    (To create the format variables use New-SmartsheetFormatString)
    $Array | Export-SmartsheetRows -id $Sheet.Id -blankRowAbove -title "My Title" -TitleFormat $titleFormat -includeHeaders -headerFormat $headerFormat
    .EXAMPLE
    The following example imports the array into a smartsheet appending the rows to the existing sheet without any title or headers. 
    This can be used to append rows to the Smartsheet. No attempt is made to prevent duplicate data.
    If the number of properties in the objects is more than the existing columns, then generic columns are created.
    (To update rows based in their primary column values use the Update-Smartsheet function.)
    $Array | Export-SmartsheetRows -id $Sheet.id
    #>
}

function Send-SmartsheetRowsViaEmail() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string[]]$rowIds,
        [string[]]$columnIds,
        [Parameter(Mandatory = $true)]
        [string[]]$To,
        [string]$subject,
        [string]$message,
        [switch]$includeAttachments,
        [switch]$includeDiscussions,
        [ValidateSet('HORIZONTAL','VERTICAL')]
        [string]$layout = 'HORIZONTAL',
        [switch]$ccMe
    )

    $Headers = Get-Headers

    $Uri = "{0}/sheets/{1}/rows/emails" -f $BaseURI, $Id

    $sendto = @()
    $To | ForEach-Object {
        $sendTo += [PSCustomObject]@{
            email = $_
        }
    }

    $payload = [ordered]@{
        rowIds = $rowIds
        columnIds = $columnIds
        includeAttachments = $includeAttachments.IsPresent
        includeDiscussions = $includeDiscussions.IsPresent        
        layout = $layout        
        ccMe = $ccMe.IsPresent
        message = $message
        sendTo = $sendTo
        subject = $subject
    }

    $body = $payload | ConvertTo-Json
    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq 'SUCCESS') {
            return $true
        } else {
            throw $response.message
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Send a select set of rows via email,
    .PARAMETER Id
    The Smartsheet Id.
    .PARAMETER rowIds
    An array row Ids to be included.
    .PARAMETER columnIds
    An array of column Ids to be included.
    If the columnIds attribute of the MultiRowEmail object is specified as an array of column Ids, those specific columns are included.
    If the columnIds attribute of the MultiRowEmail object is omitted, all columns except hidden columns shall be included.
    If the columnIds attribute of the MultiRowEmail object is specified as empty, no columns shall be included. 
    (NOTE: In this case, either includeAttachments=true or includeDiscussions=true must be specified.)
    .PARAMETER To
    An array of recipients.
    .PARAMETER subject
    The subject of the email.
    .PARAMETER message
    The message of the email.
    .PARAMETER ccMe
    Copy email to sender.
    .PARAMETER layout
    Layout of the rows. Either horizontal or Vertical. Default is horizontal for multiple rows, vertical for a single row.
    .PARAMETER includeAttachments
    Include attachment in email.
    .PARAMETER includeDiscussions
    Include Discussions in email.
    #>
}

function Copy-SmartSheetRows() {
    [CmdletBinding(DefaultParameterSetName = 'default')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sourceSheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$targetSheetId,
        [Parameter(Mandatory = $true)]
        [string[]]$rowIds,
        [Parameter(ParameterSetName='all')]
        [switch]$includeAll,
        [Parameter(ParameterSetName = 'each')]
        [switch]$includeAttachments,
        [Parameter(ParameterSetName = 'each')]
        [switch]$includeChildren,
        [Parameter(ParameterSetName = 'each')]
        [switch]$includeDiscussions,
        [switch]$ignoreRowsNotFound
    )
    
    $Headers = Get-Headers

    $Uri = "{0}/sheets/{1}/rows/copy" -f $BaseURI, $Id

    if ($includeAll) {
        $uri = "{0}?include=all" -f $Uri
    } else {
        $includes = @()
        if ($includeAttachments) {
            $Includes += "attachments"
        }
        if ($includeChildren) {
            $includes += "children"
        }
        if ($includeDiscussions) {
            $includes += "discussions"
        }
        if ($includes.Length -gt 0) {
            $str_Includes = $includes -split ","
            $Uri = "{0}?include=" -f $Uri, $str_includes
        }
    }

    If ($ignoreRowsNotFound) {
        if ($Uri.Contains("?") ) {
            $Uri = "{0}&ignoreRowsNotFound=true" -f $Uri
        } else {
            $Uri = "{0}?ignoreRowsNotFound=true" -f $Uri
        }
    }

    $payload = [ordered]@{
        rowIds = $rowIds
        to = @{
            sheetId = $id
        }
    }

    $body = $payload | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        return $response
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Copy rows from on Smartsheet to another.
    .DESCRIPTION
    Copies selected rows tro the bottom of the target sheet.
    .PARAMETER Id
    The source Sheet Id
    .PARAMETER targetSheetId
    The Target sheet Id.
    .PARAMETER rowIds
    An array of row Ids to copy to the target sheet. 
    .PARAMETER includeAll
    include all of 'attachments', 'children' and 'discussions'
    .PARAMETER includeAttachments
    Include row attachments.
    .PARAMETER includeChildren
    Include Child rows.
    If specified, any child rows of the rows specified in the request are also copied to the destination sheet, 
    and parent-child relationships amongst rows are preserved within the destination sheet; if not specified, 
    only the rows specified in the request are copied.
    .PARAMETER includeDiscussions
    Include row discussions.
    .PARAMETER ignoreRowsNotFound
    If specified, row Ids that do not exist within the source sheet does not cause an error response. If omitted, 
    specifying row Ids that do not exist within the source sheet causes an error response (and no rows are copied).
    .OUTPUTS
    An object containing the row mappings.
    #>
}

function Move-SmartSheetRows() {
    [CmdletBinding(DefaultParameterSetName = 'default')]
    Param(
        [Parameter(
            Mandatory =  $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sourceSheetId')]
        [string]$id,
        [Parameter(Mandatory = $true)]
        [string]$targetSheetId,
        [Parameter(Mandatory = $true)]
        [string[]]$rowIds,
        [Parameter(ParameterSetName='all')]
        [switch]$includeAll,
        [Parameter(ParameterSetName = 'each')]
        [switch]$includeAttachments,
         [Parameter(ParameterSetName = 'each')]
        [switch]$includeDiscussions,
        [switch]$ignoreRowsNotFound
    )

    $Headers = Get-Headers

    $Uri = "{0}/sheets/{1}/rows/move" -f $BaseURI, $Id

    if ($includeAll) {
        $uri = "{0}?include=attachments,discussions" -f $Uri
    } else {
        $includes = @()
        if ($includeAttachments) {
            $Includes += "attachments"
        }
        if ($includeDiscussions) {
            $includes += "discussions"
        }
        if ($includes.Length -gt 0) {
            $str_Includes = $includes -split ","
            $Uri = "{0}?include=" -f $Uri, $str_includes
        }
    }

    If ($ignoreRowsNotFound) {
        if ($Uri.Contains("?") ) {
            $Uri = "{0}&ignoreRowsNotFound=true" -f $Uri
        } else {
            $Uri = "{0}?ignoreRowsNotFound=true" -f $Uri
        }
    }

    $payload = [ordered]@{
        rowIds = $rowIds
        to = @{
            sheetId = $id
        }
    }

    $body = $payload | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        return $response
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Move rows from one Smartsheet to another.
    .DESCRIPTION
    Moves selected rows to the bottom of the target sheet.
    .PARAMETER Id
    The source Sheet Id
    .PARAMETER targetSheetId
    The Target sheet Id.
    .PARAMETER rowIds
    An array of row Ids to move to the target sheet. 
    .PARAMETER includeAll
    include both attachments and discussions.
    .PARAMETER includeAttachments
    Include row attachments.
    .PARAMETER includeDiscussions
    Include row discussions.
    .PARAMETER ignoreRowsNotFound
    If specified, row Ids that do not exist within the source sheet do not cause an error response. If omitted, 
    specifying row Ids that do not exist within the source sheet causes an error response (and no rows are copied).
    .OUTPUTS
    An object containing the row mappings.
    #>
}
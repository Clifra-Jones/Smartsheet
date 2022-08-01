function Get-SmartsheetColumn () {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$ColumnId
    )
    Begin {
        $Headers = Get-Headers
    }
    Process {
        $Uri = "{0}/sheets/{1}/columns/{2}" -f $BaseURI, $Id, $ColumnId
        $column = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $column
    }
    <#
    .SYNOPSIS
    Retrieve a Smartsheet column.
    .DESCRIPTION
    Retrieve a smartcheet column from a specified smartsheet.
    .PARAMETER Id
    The Id of the sheet to retrieve the column.
    .PARAMETER ColumnId
    The column ID to retrieve.
    .OUTPUTS
    A smartsheet column ofbject.

    #>
}

function Set-SmartsheetColumn {
    [CmdletBinding(DefaultParameterSetName = "props")]
    param (
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "props"
        )]
        [Parameter(ParameterSetName = "column")]
        [string]$Id,
        [Parameter(
            Mandatory = $true,    
            ParameterSetName="column",
            ValueFromPipeline = $true
        )]
        [Column]$column,
        [Parameter(
            Mandatory = $true
        )]
        [string]$ColumnId,
        [Parameter(Mandatory = $true)]
        [int]$Index,
        [string]$title,
        [string]$description,
        [ValidateSet("ABSTRACT_DATETIME", "CHECKBOX", "CONTACT_LIST", "DATE", 
            "DATETIME", "DURATION", "MULTI_CONTACT_LIST", "MULTI_PICKLIST", "PICKLIST", "PREDECESSOR", "TEXT_NUMBER")]
        [string]$type,
        [psobject]$formula,
        [bool]$hidden,
        [psobject]$autoNumberFormat,
        [psobject]$contactOptions,
        [string]$format,
        [bool]$locked,
        [string[]]$options,
        [string]$symbol,
        [ValidateSet("AUTO_NUMBER", "CREATED_BY", "CREATED_DATE", "MODIFIED_BY", "MODIFIED_DATE")]
        [string]$systemColumnType,
        [bool]$validation,
        [int]$version,
        [int]$width
    )


    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/columns/{2}" -f $BaseURI, $Id, $ColumnId
    $body = $null

    if ($column) {
        $body = $column | ConvertTo-Json -Compress
    } else {
        $properties = @{
            index = $index
        }
        if($validation -and ($type = "TEXT_NUMBER")) {
            throw "validation is invalid on this column Type 'TEXT_NUMBER'."
        }
        if ($title) { $properties.Add("title" , $title) }
        if ($type) { $properties.Add("type", $type) }
        if ($formula) { $properties.Add("formula", $formula) }
        if ($hidden) { $properties.Add("hidden", $hidden) }
        if ($autoNumberFormat) { $Properties.Add("autoNumberFormat", $autoNumberFormat) }
        if ($contactObject) { $properties.Add("contactObject", $contactObject) }
        if ($description) { $properties.Add("description", $description) }
        if ($format) { $properties.Add("format", $format) }
        if ($locked) { $properties.Add("locked", $locked) }
        if ($options) { $properties.Add("options", $options) }
        if ($symbol) { $properties.Add("symbol", $symbol) }
        if ($systemColumnType) { $properties.Add("systemColumnType", $systemColumnType) }
        if ($validation) { $properties.Add("validation", $validation) }
        if ($version) { $properties.Add("version", $version) }
        if ($width) { $properties.Add("width", $width) }       
       
        $body = $body | ConvertTo-Json -Compress
    }
    # remove the property 'lockedForUser as you cannot write that to the API.
   
    $response = Invoke-RestMethod -Method Put -Uri $Uri -Headers $Headers -Body $body
    return $response
    <#
    .SYNOPSIS
    Update a Smartsheet column
    .DESCRIPTION
    Update the properties of a Smartsheet column.
    .PARAMETER Id
    Id of the Smartsheet containing the column.
    .PARAMETER column
    A Smartsheet column object. Cannot be used with column property parameters.
    .PARAMETER ColumnId
    Id of teh column to update.
    .PARAMETER Index
    Index if the column to update.
    .PARAMETER title
    Column Title
    .PARAMETER description
    Column description
    .PARAMETER type
    Column type
    .PARAMETER formula
    The formula for a column, if set, for instance =data@row.
    .PARAMETER hidden
    Indicates visibility of the column.
    .PARAMETER autoNumberFormat
    Object that describes how the the System Column type of "AUTO_NUMBER" is auto-generated.
    .PARAMETER contactOptions
    Array of ContactOption objects to specify a pre-defined list of values for the column. Column type must be CONTACT_LIST.
    .PARAMETER format
    Format string.
    .PARAMETER locked
    Indicates whether the column is locked. A value of true indicates that the column has been locked by the sheet owner or the admin.
    .PARAMETER options
    Array of the options available for the column.
    .PARAMETER symbol
    When applicable for CHECKBOX or PICKLIST column types.
    .PARAMETER systemColumnType
    If this is a system column what type is it.
    .PARAMETER validation
    Indicates whether validation has been enabled for the column (value = true).
    .PARAMETER version
    0: CONTACT_LIST, PICKLIST, or TEXT_NUMBER.
    1: MULTI_CONTACT_LIST.
    2: MULTI_PICKLIST.
    .PARAMETER width
    Display width of the column in pixels.
    .OUTPUTS
    An updated column object.
    #>
} 

function Get-SmartsheetColumns () {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$SheetId
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/columns" -f $BaseURI, $SheetId

    $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
    return $response.data
    <#
    .SYNOPSIS
    Retrieve Smartsheet columns.
    .DESCRIPTION
    Returns an array of the columns in a smartsheet.
    .PARAMETER Id
    The Id of the SMartsheet to return columns from.
    .OUTPUTS
    An array of smartsheet column objects.
    #>
}

function Add-SmartsheetColumn() {
    [CmdletBinding(DefaultParameterSetName = "props")]
    param (
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "props"
        )]
        [Parameter(ParameterSetName = "column")]
        [string]$Id,
        [Parameter(
            Mandatory = $true,    
            ParameterSetName="column",
            ValueFromPipeline = $true
        )]
        [Column]$column,
        [Parameter(
            Mandatory = $true
        )]
        [string]$ColumnId,
        [Parameter(Mandatory = $true)]
        [int]$Index,
        [string]$title,
        [string]$description,
        [ValidateSet("ABSTRACT_DATETIME", "CHECKBOX", "CONTACT_LIST", "DATE", 
            "DATETIME", "DURATION", "MULTI_CONTACT_LIST", "MULTI_PICKLIST", "PICKLIST", "PREDECESSOR", "TEXT_NUMBER")]
        [string]$type,
        [psobject]$formula,
        [bool]$hidden,
        [psobject]$autoNumberFormat,
        [psobject]$contactObject,
        [string]$format,
        [bool]$locked,
        [string[]]$options,
        [string]$symbol,
        [ValidateSet("AUTO_NUMBER", "CREATED_BY", "CREATED_DATE", "MODIFIED_BY", "MODIFIED_DATE")]
        [string]$systemColumnType,
        [bool]$validation,
        [int]$version,
        [int]$width
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/columns" -f $BaseURI, $SheetId

    if ($column) {
        $body = $column | ConvertTo-Json -Compress
    } else {
        $properties = @{
            index = $index
        }
        
        if ($title) { $properties.Add("title" , $title) }
        if ($type) { $properties.Add("type", $type) }
        if ($formula) { $properties.Add("formula", $formula) }
        if ($hidden) { $properties.Add("hidden", $hidden) }
        if ($autoNumberFormat) { $Properties.Add("autoNumberFormat", $autoNumberFormat) }
        if ($contactObject) { $properties.Add("contactObject", $contactObject) }
        if ($description) { $properties.Add("description", $description) }
        if ($format) { $properties.Add("format", $format) }
        if ($locked) { $properties.Add("locked", $locked) }
        if ($options) { $properties.Add("options", $options) }
        if ($symbol) { $properties.Add("symbol", $symbol) }
        if ($systemColumnType) { $properties.Add("systemColumnType", $systemColumnType) }
        if ($validation) { $properties.Add("validation", $validation) }
        if ($version) { $properties.Add("version", $version) }
        if ($width) { $properties.Add("width", $width) }        
        $column = [psCustomObject]$properties
        $body = $column | ConvertTo-Json -Compress
    }
    
    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $response.result
    }
    <#
    .SYNOPSIS
    Add a column to a Smartsheet
    .DESCRIPTION
    Adds a new column to a smartsheet. Column can be specified as a column object or properties parameters.
    If is column index already exists the column will be inserted at that position. Columns after that index will have their index
    incremented by 1.
    .PARAMETER Id
    Id of the Smartsheet containing the column.
    .PARAMETER column
    A Smartsheet column object. Cannot be used with column property parameters.
    .PARAMETER ColumnId
    Id of teh column to update.
    .PARAMETER Index
    Index if the column to update.
    .PARAMETER title
    Column Title
    .PARAMETER description
    Column description
    .PARAMETER type
    Column type
    .PARAMETER formula
    The formula for a column, if set, for instance =data@row.
    .PARAMETER hidden
    Indicates visibility of the column.
    .PARAMETER autoNumberFormat
    Object that describes how the the System Column type of "AUTO_NUMBER" is auto-generated.
    .PARAMETER contactOptions
    Array of ContactOption objects to specify a pre-defined list of values for the column. Column type must be CONTACT_LIST.
    .PARAMETER format
    Format string.
    .PARAMETER locked
    Indicates whether the column is locked. A value of true indicates that the column has been locked by the sheet owner or the admin.
    .PARAMETER options
    Array of the options available for the column.
    .PARAMETER symbol
    When applicable for CHECKBOX or PICKLIST column types.
    .PARAMETER systemColumnType
    If this is a system column what type is it.
    .PARAMETER validation
    Indicates whether validation has been enabled for the column (value = true).
    .PARAMETER version
    0: CONTACT_LIST, PICKLIST, or TEXT_NUMBER.
    1: MULTI_CONTACT_LIST.
    2: MULTI_PICKLIST.
    .PARAMETER width
    Display width of the column in pixels.
    .OUTPUTS
    An updated column object.
    #>
}

function Add-SmartsheetColumns() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [List[Column]]$columns
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/columns" -f $BaseURI, $Id

    $body = $Columns | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $response.result
    } else{
        throw $response.message
    }
    <#
    .SYNOPSIS
    Add columns to a Smartsheet.
    .DESCRIPTION
    Adds an array of Smartsheet columns to a Smartsheet.
    .PARAMETER Id
    The Id fo the smartsheet to add columns to.
    .PARAMETER columns
    An array of smartsheet columns.
    .OUTPUTS
    An array of the newly added columns.
    #>
}

function Remove-SmartsheetColumn() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$columnId
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/columns/{2}" -f $BaseURI, $Id, $columnId

    $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
    If ($response.message = "SUCCESS") {
        return $true
    } else {
        return $false
    }
    <#
    .SYNOPSIS
    Remove a smartsheet column.
    .DESCRIPTION
    Remove a column from a smartsheet.
    .PARAMETER Id
    The Id of the Smartsheet to remove the column.
    .PARAMETER columnId
    The ID of the column to remove.
    .OUTPUTS
    Boolean indicating success or failue of the operation.
    #>    
}

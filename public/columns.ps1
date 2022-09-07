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
    The column Id to retrieve.
    .OUTPUTS
    A smartsheet column ofbject.

    #>
}

function Set-SmartsheetColumn {
    [CmdletBinding(DefaultParameterSetName = "default")]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(
            Mandatory = $true
        )]
        [string]$ColumnId,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'column'
        )]
        [psobject]$column,
        [Parameter(Mandatory = $true, ParameterSetName ='props')]
        [int]$Index,
        [Parameter(ParameterSetName ='props')]
        [string]$title,
        [Parameter(ParameterSetName ='props')]
        [string]$description,
        [ValidateSet("ABSTRACT_DATETIME", "CHECKBOX", "CONTACT_LIST", "DATE", 
            "DATETIME", "DURATION", "MULTI_CONTACT_LIST", "MULTI_PICKLIST", "PICKLIST", "PREDECESSOR", "TEXT_NUMBER")]
        [Parameter(ParameterSetName ='props')]
        [string]$type,
        [Parameter(ParameterSetName ='props')]
        [psobject]$formula,
        [Parameter(ParameterSetName ='props')]
        [bool]$hidden,
        [Parameter(ParameterSetName ='props')]
        [psobject]$autoNumberFormat,
        [Parameter(ParameterSetName ='props')]
        [psobject]$contactOptions,
        [Parameter(ParameterSetName ='props')]
        [string]$format,
        [Parameter(ParameterSetName ='props')]
        [bool]$locked,
        [Parameter(ParameterSetName ='props')]
        [string[]]$options,
        [Parameter(ParameterSetName ='props')]
        [string]$symbol,
        [ValidateSet("AUTO_NUMBER", "CREATED_BY", "CREATED_DATE", "MODIFIED_BY", "MODIFIED_DATE")]
        [Parameter(ParameterSetName ='props')]
        [string]$systemColumnType,
        [Parameter(ParameterSetName ='props')]
        [bool]$validation,
        [Parameter(ParameterSetName ='props')]
        [int]$width,
        [switch]$PassThru
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
    try {
        $response = Invoke-RestMethod -Method Put -Uri $Uri -Headers $Headers -Body $body
        if ($PassThru) {
            return Get-Smartsheet -id $id
        }
        return $response
    } catch {
        throw $_
    }
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
    Id of the column to update.
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
    .PARAMETER PassThru
    Return the updated sheet.
    .OUTPUTS
    An updated column object.
    If PassThru is provided returns the updated sheet object.
    #>
} 

function Get-SmartsheetColumns () {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id
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
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "column"
        )]
        [psObject]$column,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "props"
        )]
        [string]$ColumnId,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "props"
        )]
        [int]$Index,
        [Parameter(ParameterSetName = 'props')]
        [string]$title,
        [Parameter(ParameterSetName = 'props')]
        [string]$description,
        [Parameter(ParameterSetName = 'props')]
        [ValidateSet("ABSTRACT_DATETIME", "CHECKBOX", "CONTACT_LIST", "DATE", 
            "DATETIME", "DURATION", "MULTI_CONTACT_LIST", "MULTI_PICKLIST", "PICKLIST", "PREDECESSOR", "TEXT_NUMBER")]
        [string]$type,
        [Parameter(ParameterSetName = 'props')]
        [psobject]$formula,
        [Parameter(ParameterSetName = 'props')]
        [bool]$hidden,
        [Parameter(ParameterSetName = 'props')]
        [psobject]$autoNumberFormat,
        [Parameter(ParameterSetName = 'props')]
        [psobject]$contactOptions,
        [Parameter(ParameterSetName = 'props')]
        [string]$format,
        [Parameter(ParameterSetName = 'props')]
        [bool]$locked,
        [Parameter(ParameterSetName = 'props')]
        [string[]]$options,
        [Parameter(ParameterSetName = 'props')]
        [string]$symbol,
        [Parameter(ParameterSetName = 'props')]
        [ValidateSet("AUTO_NUMBER", "CREATED_BY", "CREATED_DATE", "MODIFIED_BY", "MODIFIED_DATE")]
        [string]$systemColumnType,
        [Parameter(ParameterSetName = 'props')]
        [bool]$validation,
        [Parameter(ParameterSetName = 'props')]
        [int]$width,
        [switch]$PassThru
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
        if ($contactOptions) { $properties.Add("contactObject", $contactOptions) }
        if ($description) { $properties.Add("description", $description) }
        if ($format) { $properties.Add("format", $format) }
        if ($locked) { $properties.Add("locked", $locked) }
        if ($options) { $properties.Add("options", $options) }
        if ($symbol) { $properties.Add("symbol", $symbol) }
        if ($systemColumnType) { $properties.Add("systemColumnType", $systemColumnType) }
        if ($validation) { $properties.Add("validation", $validation) }
        if ($width) { $properties.Add("width", $width) }        
        $column = [psCustomObject]$properties
        $body = $column | ConvertTo-Json -Compress
    }
    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            if ($PassThru) {
                return Get-Smartsheet -id $id
            } else {
                return $response.result
            }
        }
    } catch {
        throw $_
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
    The contact option object is in the form: 
    email = {email address}
    name = {contact name}
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
    .PARAMETER PassThru
    Return the updated sheet.
    .OUTPUTS
    An updated column object.
    if PassThru is provided returns the updated sheet object.
    .EXAMPLE
    To add a new colum to a Smartsheet.
    PS> $newColumn = $Sheet | Add-SmartsheetColumn -title "Title" -type:TEXT_NUMBER -description 'My new column'
    .EXAMPLE
    To insert a new column at position 4 (columns after position 4 are shifted to the right and thier index incremented).
    PS> $newColumn = $Sheet | Add-SmartsheetColumn -title "Asset" -type:TEXT_NUMBER -Description "Fixed asset" -index 4
    .Example 
    Add a new column with contact objects.
    PS> $contacts = @(
        @{
            email = "johndoe@example.com"
            name = "John Doe"
        },
        @{
            email = "janedoe@example.com"
            name = 'Jane Doe
        }
    )
    PS> $newColumn = $Sheet | Add-SmartsheetColumn -title "EmployeeName" -type:TEXT_NUMBER -contactOption $contacts
    #>    
}

function Add-SmartsheetColumns() {
    [CmdletBinding(DefaultParameterSetName = 'allq')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [psobject[]]$columns,
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/columns" -f $BaseURI, $Id

    $body = $Columns | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    try {
        if ($response.message -eq "SUCCESS") {
            return $response.result
        } else{
            throw $response.message
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Add columns to a Smartsheet.
    .DESCRIPTION
    Adds an array of Smartsheet columns to a Smartsheet.
    .PARAMETER Id
    The Id of the smartsheet to add columns to.
    .PARAMETER columns
    An array of smartsheet columns.
    .PARAMETER PassThru
    Return the updated Sheet.
    .OUTPUTS
    An array of the newly added columns.
    if PassThru is provided returns the updated sheet object.
    #>
}

function Remove-SmartsheetColumn() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$columnId,
        [switch]$PassThru
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/columns/{2}" -f $BaseURI, $Id, $columnId

    try {
        $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
        If ($response.message = "SUCCESS") {
            if ($PassThru) {
                return Get-SmartSheet -id $id
            }
            return $true
        } else {
            return $false
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Remove a smartsheet column.
    .DESCRIPTION
    Remove a column from a smartsheet.
    .PARAMETER Id
    The Id of the Smartsheet to remove the column.
    .PARAMETER columnId
    The Id of the column to remove.
    .PARAMETER PassThru
    Return the updated sheet.
    .OUTPUTS
    Boolean indicating success or failue of the operation.
    if PassThru is provided returns the updated sheet object.
    #>    
}

function New-SmartsheetColumn() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$title,
        [switch]$primary,
        [string]$description,        
        [ValidateSet("ABSTRACT_DATETIME", "CHECKBOX", "CONTACT_LIST", "DATE", 
            "DATETIME", "DURATION", "MULTI_CONTACT_LIST", "MULTI_PICKLIST", "PICKLIST", "PREDECESSOR", "TEXT_NUMBER")]
        [string]$type = "TEXT_NUMBER",
        [psobject]$formula,
        [switch]$hidden,
        [psobject]$autoNumberFormat,
        [psobject]$contactOptions,
        [string]$format,
        [switch]$locked,
        [string[]]$options,
        [string]$symbol,
        [ValidateSet("AUTO_NUMBER", "CREATED_BY", "CREATED_DATE", "MODIFIED_BY", "MODIFIED_DATE")]
        [string]$systemColumnType,
        [switch]$validation,
        [int]$width
    )

    $properties = [ordered]@{
        title = $title
        type = $type
        primary = $primary.IsPresent
    }
    if ($description) {$properties.Add("description", $description)}
    if ($formula) {$properties.Add("formula", $formula)}
    if ($hidden.IsPresent) {$properties.Add("hidden", $hidden.IsPresent)}
    if ($autoNumberFormat) {$properties.Add("autoNumberFormat", $autoNumberFormat)}
    if ($contactOptions) {$properties.Add("contactOptions", $contactOptions)}
    if ($format) {$properties.Add("format", $format)}
    if ($locked.IsPresent) {$properties.add("locked", $locaked.IsPresent)}
    if ($options) {
        if ($type -notin "PICKLIST","MULTI_PICKLIST") {
            throw "Option are only valid for column types PICKLIST or MULTI_PICKLIST!"
        }
        $properties.Add("options", $options)
    }
    if ($symbol) {$properties.Add("symbol", $symbol)}
    if ($systemColumnType) {$properties.Add("systemColumnType", $systemColumnType)}
    if ($validation.IsPresent) {$properties.Add("validation", $validation.IsPresent)}
    if ($width) {
        $properties.Add("width", $width)
    } else {
        $properties.Add("width", 100)
    }

    $column = [PSCustomObject]$Properties
    return $column
    <#
    .SYNOPSIS
    Create a Smartsheet Column object.
    .DESCRIPTION
    This function does not insert the column into a sheet. Use the Add-SmartsheetColumn function to add the column to a sheet.
    .PARAMETER title
    Column title
    .PARAMETER primary
    Set this column to the primary column.
    NOTE: if a primary column already exists this wil cause an error when added to a smartsheet.
    .PARAMETER description
    Colunm description.
    .PARAMETER type
    Column type.
    .PARAMETER formula
    Formula forthe column.
    .PARAMETER hidden
    Column is hidden
    .PARAMETER autoNumberFormat
    Object that describes how the the System Column type of "AUTO_NUMBER" is auto-generated.
    .PARAMETER contactOptions
    Array of ContactOption objects to specify a pre-defined list of values for the column. Column type must be CONTACT_LIST.
    .PARAMETER format
    The format descriptor.
    .PARAMETER locked
    Column is locked.
    .PARAMETER options
    Array of option for the column. Only valid for PICKLIST and MULTI_PICKLIST column types.
    .PARAMETER symbol
    Only applicable for CHECKBOX and PICKLIST column types.
    .PARAMETER systemColumnType
    If this is a SystemColumn type the type of system column.
    .PARAMETER validation
    Is validation enabled.
    .PARAMETER width
    Width of the column in pixels.
    .OUTPUTS
    A Smartsheet column object.
    #>
}
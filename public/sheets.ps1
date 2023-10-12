using namespace System.Collections.Generic

function New-Smartsheet() {
    [CmdletBinding()]
    Param(
        [Alias('folderId','workspaceId')]
        [Parameter(
            ValueFromPipelineByPropertyName = $true
        )]
        [UInt64]$Id,
        [ValidateSet('home','folder','workspace')]
        [string]$ContainerType = 'home',
        [Parameter(Mandatory = $true)]
        [string]$SheetName,
        [ValidateScript(
            {
                if (-not $_ -and ($templateId)) {
                    throw "Parameter Columns cannot be used with parameter TemplateId."
                } else {
                    $true
                }
            }
        )]        
        [psobject[]]$Columns,
        [ValidateScript(
            {
                if (-not $_ -and ($Columns)) {
                    throw "Parameter TemplateId cannot be used with parameter Columns."
                }
            }
        )]
        [UInt64]$TemplateId,
        [ValidateScript(
            {
                if ($_) {
                    if (
                        $null -ne $IncludeAttachments -or
                        $null -ne $includeCellLinks -or 
                        $null -ne $includeData -or 
                        $null -ne $includeDiscussions -or
                        $null -ne $IncludeFilters -or
                        $null -ne $includeForms -or
                        $null -ne $IncludeRuleReceipts -or
                        $Null -ne $IncludeRules
                    ) {
                        throw "Parameter Include all must not be used with other include parameters."
                    } else {
                        $true
                    }
                } 
            }
        )]
        [switch]$IncludeAll,
        [switch]$IncludeAttachments,
        [switch]$IncludeCellLinks,
        [switch]$IncludeData,
        [switch]$IncludeDiscussions,
        [switch]$IncludeFilters,
        [switch]$IncludeForms,
        [switch]$IncludeRuleReceipts,
        [switch]$IncludeRules 
    )

    switch ($ContainerType) {
        'folder' {
            $Uri = "{0}/folders/{1}/sheets" -f $BaseURI, $id
        }
        'workspace' {
            $Uri = "{0}/workspaces/{1}/sheets" -f $BaseURI, $Id
            if ($IncludeAll) {
                $includes = "attachments", "cellLinks", "data", "discussions", "filters", "forms", "ruleRecipients", "rules"
                $Uri = "{0}?include={1}" -f $Uri, ($includes -join ",")
            } else {
                $includes = [List[string]]::New()
                if ($includeAttachments) {$includes.Add("attachments")}
                if ($includeCellLinks) {$includes.Add("cellLinks")}
                if ($includeData) {$includes.Add("data")}
                if ($includeDiscussions) {$includes.Add("discussions")}
                if ($includeFilters) {$includes.Add("filters")}
                if ($includeForms) {$includes.Add("forms")}
                if ($includeRuleReceipts) {$includes.Add("ruleReciepts")}
                if ($includeRules) {$includes.Add("rules")}

                if ($includes.Count -gt 0) {
                    $Uri = "{0}?include={1}" -f $Uri, ($includes.ToArray() -join ",")
                }
            }
        }
        'home' {
            $Uri = "{0}/sheets" -f $BaseURI
        }
    }
    
    $Headers = Get-Headers

    if ($columns) {
        $payload = [ordered]@{
            name = $sheetName
            columns = $columns
        }
    } elseif ($templateId) {
        $payload = [ordered]@{
            fromId = $templateId
            name = $sheetName
        }
    } else {
        $columns = @()
        $columns += New-SmartsheetColumn -title "NewColumn" -primary
        $payload = [ordered]@{
            name = $sheetName
            columns = $columns
        }
    }
    $body = $payload | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
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
    Create a new Smartsheet
    .DESCRIPTION
    Creates a sheet from scratch or from the specified template.
    .PARAMETER Id
    Folder Id where you can create sheets, sights, reports, templates, and other folders.
    .PARAMETER ContainerType
    Type of container. Either 'home','folder', or'workspace'.
    .PARAMETER SheetName
    Name of the new Sheet
    .PARAMETER 
    .PARAMETER Columns
    An array of column objects.
    .PARAMETER TemplateId
    The Id of a sheet to use as a template.
    .PARAMETER IncludeAll
    Include all elements from the template sheet. cannot be used with other Include parameters.
    .PARAMETER IncludeAttachments
    Include attachments
    .PARAMETER IncludeCellLinks
    Include cell links
    .PARAMETER IncludeData
    Include data
    .PARAMETER IncludeDiscussions
    Include discussions
    .PARAMETER IncludeFilters
    Include filters
    .PARAMETER IncludeForms
    Include forms
    .PARAMETER IncludeRuleReceipts
    Include receipts
    .PARAMETER IncludeRules
    Include rules
    .OUTPUTS
    Result object containing a Sheet object for newly created sheet, corresponding to what was specified in the request.
    .NOTES
    Sheet names ARE NOT unique in folders. You can create a sheet with the same name as another sheet. 
    This function does not check for name conflicts in the target folder.
    You should check for the existence of a sheet of the same name using the Get-Smartsheets function.
    #>
    
}

function Get-Smartsheets () {
 
    $Uri = "{0}/sheets?includeAll=true" -f $BaseURI
    $Headers = Get-Headers

    $response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers

    return $response.data
    <#
    .SYNOPSIS
    Gets all smartsheet.
    .DESCRIPTION
    Gets an array of Smartsheet object associated the user has access to.
    .OUTPUTS
    AN array of Smartsheet objects.
    #>
} 


function Get-Smartsheet () {    
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = "id"
        )]
        [Alias('SheetId')]
        [UInt64]$id,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "name"
        )]
        [string]$Name,
        [ValidateSet(0,1,2)]
        [int]$level,
        [switch]$includeAll,
        [switch]$includeAttachments,
        [switch]$includeColumnTypes,
        [switch]$includeCrossSheetReferences,
        [switch]$includeDiscussions,
        [switch]$includeFilters,
        [switch]$includeFilterDefinitions,
        [switch]$includeFormat,
        [switch]$includeGantConfig,
        [switch]$includeObjectValue,
        [switch]$includeOwnerInfo,
        [switch]$includeRowPermalink,
        [switch]$includeSource,
        [switch]$includeWriterInfo,
        [switch]$excludeFilteredOutRows,
        [switch]$excludeLinkInFromCellDetails,
        [switch]$excludeLinkOutToCellDetails,
        [switch]$excludeNonexistentCells,
        [psobject[]]$columnIds,
        [psobject[]]$rowIds
    )
    
    Begin {
        $Header = Get-Headers
        $includes = [List[string]]::New()
        If ($includeAll) {
            "attachments","columnType","crossSheetReferences","discussions","filters","filterDefinitions","format","ganttConfig", `
            "objectValue","ownerInfo","rowPermalink","source","writerInfo" `
            | ForEach-Object { $includes.Add($_)}
        } else {
            if ($includeAttachments) {$includes.Add("attachments")}
            if ($includeColumnTypes) {$includes.Add("columnType")}
            if ($includeCrossSheetReferences) {$includes.Add("crossSheetReferences")}
            if ($includeDiscussions) {$includes.Add("discussions")}
            if ($includeFilters) {$includes.Add("filters")}
            if ($includeFilterDefinitions) {$includes.Add("filterDefinitions")}
            if ($includeFormat) {$includes.Add("format")}
            if ($includeGantConfig) {$includes.Add("ganttConfig")}
            if ($includeObjectValue) {$includes.Add("objectValue")}
            if ($includeOwnerInfo) {$includes.Add("ownerInfo")}
            if ($includeRowPermalink) {$includes.Add("rowPermalink")}
            if ($includeSource) {$includes.Add("source")}
            if ($includeWriterInfo) {$includes.Add("writerInfo")}
        }
        $excludes = [List[string]]::New()
        if ($excludeFilteredOutRows) {$excludes.Add("filteredOutRows")}
        if ($excludeLinkInFromCellDetails) {$excludes.Add("linkInFromCellDetails")}
        if ($excludeLinkOutToCellDetails) {$excludes.Add("linksOutToCellsDetails")}
        if ($excludeNonexistentCells) {$excludes.Add("nonexistentCells")}
    }

    Process {
        # Was a sheet name Provided
        if ($Name) {
            # Get Smartsheet(s) that match the name.
            $sheetInfo = Get-Smartsheets | Where-Object {$_.name -eq $Name}
            if (-not $sheetInfo) {
                $msg = "Sheet {0} not found!" -f $Name
                Throw $msg
            }
            # There may be more than one sheet that matches the name. Prompt the user to select the sheet.
            if ($sheetInfo -is [array]) {
                Write-Host "Select Which Sheet to load:"
                Do {
                    foreach ($Sheet in $sheetInfo) {
                        $Msg = "{0}:{1}:{2}" -f ($sheetInfo.IndexOf($Sheet) +1), $Sheet.Name, $Sheet.modifiedAt 
                        Write-Host $Msg
                    }
                    Write-Host "Q: Quit"
                    $R = Read-Host "Select SmartSheet:"
                    If ($R -eq "q") { exit}
                } while ( $R -notin 1..$SheetInfo.Count)
                $id = $sheetInfo[$R-1].id
            } else {
                $id = $sheetInfo.id
            }
        }

        $Uri = "{0}/sheets/{1}" -f $BaseURI, $id
        if ($level) {
            $Uri = "{0}?level={1}" -f $Uri, $level
        }
        if ($includes.Count -gt 0) {
            [string]$strIncludes = $includes.ToArray() -join ","
            if ($Uri.Contains("?")) {
                $Uri = "{0}&include={1}" -f $Uri, $strIncludes
            } else {
                $Uri = "{0}?include={1}" -f $Uri, $strIncludes
            }
        }
        If ($excludes.Count -gt 0) {
            [string]$strExcludes = $excludes.ToArray() -join ","
            if ($Uri.Contains("?")) {
                $Uri = "{0}&exclude={1}" -f $uri, $strExcludes
            } else {
                $Uri = "{0}?exclude={1}" -f $Uri, $strExcludes
            }
        }
        if ($columnIds) {
            $listOfColIds = [List[string]]::New()
            $columnIds | ForEach-Object {
                $listOfColIdss.Add($_)
            }
            $strColumnIds = $listOfColIds.ToArray() -join ","
            if ($Uri.Contains("?")) {
                $Uri = "{0}&columnIds={1}" -f $Uri, $strColumnIds
            } else {
                $Uri = "{0}?columnIds={1}" -f $Uri, $strColumnIds
            }
        }
        if ($RowIds) {
            $listOfRowIds = [List[string]]::New()
            $rowIds | ForEach-Object {
                $listOfRowIds.Add($_)
            }
            $strRowIds = $listOfRowIds.ToArray() -join ","
            if ($uri.Contains("?")) {
                $Uri = "{0}&rowIds={1}" -f $Uri, $strRowIds
            } else {
                $Uri = "{0}?rowIds={1}" -f $Uri, $strRowIds
            }
        }
        $Sheet = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Header
        $ToArray_Script = {
            param([boolean]$IncludeRowId)
            $psSheet = New-Object System.Collections.Generic.List[psobject]
            foreach ($Row in $this.rows) {                
                $Props = [ordered]@{}
                if ($IncludeRowId) {
                    $Props.Add("rowId", $row.id)
                }
                foreach ($Cell in $row.cells) {
                    $PropName = $this.columns.Where({$_.id -eq $Cell.columnId}).title
                    $Props.Add($PropName, $Cell.value)                    
                }
                $_row = New-Object -TypeName psobject -Property $props
                $psSheet.add($_row)
            }
            return $psSheet.ToArray()
        }
        $Sheet | Add-Member -MemberType ScriptMethod -Name ToArray -Value $ToArray_Script

        return $Sheet
    }
    <#
    .SYNOPSIS
    Retrieve an individual sheet.
    .DESCRIPTION
    Retrieves an individual sheet by either the sheet ID or the Name.
    Note: There can be multiple sheets with the same name. Using the Sheet ID is more accurate!
    The object returned has an additional method ToArray(), this method returns an array of PowerShell objects based on the sheet rows and columns.
    The ToArray method can accept a boolean parameter that instructs the method to include or exclude the RowId in the objects returned.
    True will return the row Id as a property of each object in the array. False will exclude the Row Id. False is the default.
    By using the Row id, you can update the values in a Smartsheet without the need for the primary column to be unique.
    The previous method of updating rows is included for backward compatibility and may eventually be removed.
    .PARAMETER id
    Folder ID, cannot be used with the Name parameter.
    .PARAMETER Name
    Sheet Name, cannot be used with the id parameter.
    .PARAMETER level
    Specifies whether new functionality, such as multi-contact data is returned in a backwards-compatible, 
    text format (level=0, default), multi-contact data (level=1), or multi-picklist data (level=2).
    .PARAMETER includeAll
    Include All Sheet objects
    .PARAMETER includeAttachments
    Includes the metadata for sheet-level and row-level attachments. 
    To include discussion attachments, both includeAttachments and includeDiscussions must be present.
    .PARAMETER includeColumnTypes
    Includes columnType attribute in the row's cells indicating the type of the column the cell resides in.
    .PARAMETER includeCrossSheetReferences
    Includes the cross-sheet references
    .PARAMETER includeDiscussions
    Includes sheet-level and row-level discussions. 
    To include discussion attachments, both includeAttachments and includeDiscussions must be present.
    .PARAMETER includeFilters
    Includes filteredOut attribute indicating if the row should be displayed or hidden according to the sheet's filters.
    .PARAMETER includeFilterDefinitions
    Includes type of filter, operators used, and criteria
    .PARAMETER includeFormat
    Includes column, row, cell, and summary fields formatting.
    .PARAMETER includeGantConfig
    Includes Gantt chart details.
    .PARAMETER includeObjectValue
    When used in combination with a level parameter, includes the email addresses for multi-contact data.
    .PARAMETER includeOwnerInfo
    Includes the workspace and the owner's email address and user Id.
    .PARAMETER includeRowPermalink
    Includes permalink attribute that represents a direct link to the row in the Smartsheet application.
    .PARAMETER includeSource
    Adds the Source object indicating which report, sheet Sight (aka dashboard), or template the sheet was created from, if any.
    .PARAMETER includeWriterInfo
    Includes createdBy and modifiedBy attributes on the row or summary fields, indicating the row or summary field's creator, and last modifier.
    .PARAMETER excludeFilteredOutRows
    Excludes filtered out rows from response payload if a sheet filter is applied; includes total number of filtered rows
    .PARAMETER excludeLinkInFromCellDetails
    Excludes the following attributes from the cell.linkInFromCell object: columnId, rowId, status
    .PARAMETER excludeLinkOutToCellDetails
    Excludes the following attributes from the cell.linksOutToCells array elements: columnId, rowId, status
    .PARAMETER excludeNonexistentCells
    Excludes cells that have never contained any data
    .PARAMETER columnIds
    An array of column ids. 
    The response contains only the specified columns in the "columns" array, and individual rows' "cells" array only 
    contains cells in the specified columns.
    .PARAMETER rowIds
    A array of row Ids on which to filter the rows included in the result.
    .NOTES
    When retrieving a smartsheet by name there is always the chance that there are multiple sheets with the same name in a folder.
    If more than one sheet have the same name, you will be prompted to select the sheet yu want from a list. 
    The list will show Sheet name and modified date.

    .OUTPUTS
    A Smartsheet sheet object.
    There is an added method named ToArray that returns the sheet as an array of PowerShell objects.
    #>
}

function Remove-Smartsheet() {   
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('SheetId')]
        [UInt64]$Id
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}" -f $BaseURI, $Id

    $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
    if ($response.message -eq "SUCCESS") {
        return $true
    } else {
        return $false
    }
     <#
    .SYNOPSIS
    Removes a smartsheet.
    .DESCRIPTION
    Removes a sheet by its SheetID.
    .PARAMETER Id
    Sheet Id, the sheet Id to remove.
    #>
}

function Copy-Smartsheet() {
    [CmdletBinding(DefaultParameterSetName = "props")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('SheetId')]
        [UInt64]$Id,
        [string]$newSheetName,
        [Parameter(ParameterSetName="container")]
        [string]$containerId,
        [Parameter(ParameterSetName="container")]
        [ValidateSet(
            "folder",
            "home",
            "workspace"
        )]
        [string]$containerType = "home",
        [switch]$includeAll,
        [switch]$includeAttachments,
        [switch]$includeCellLinks,
        [switch]$includeFormatting,
        [switch]$includeFilters,
        [switch]$includeForms,
        [switch]$includeRuleRecipients,
        [switch]$includeRules,
        [switch]$IncludeShares,
        [switch]$excludeSheetHyperlinks,
        [switch]$passThru

    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/copy" -f $BaseURI, $id
    if ($includeAll) {
        $uri = "{0}?include=all" -f $Uri
    } else {
        $includes = @()
        if ($includeAttachments) {
            $includes += "attachments"
        }
        if ($includeCellLinks) {
            $includes += "cellLinks"            
        }
        if ($includeFormatting) {
            $includes += "data"
        }
        if ($includeFilters) {
            $includes += "filters"
        }
        if ($includeForms) {
            $includes += "forms"
        }
        if ($includeRuleRecipients) {
            $includes += "ruleRecipients"
        }
        if ($includeRules) {
            $includes += "rules"
        }
        if ($IncludeShares) {
            $includes += "shares"
        }
        if ($includes.Length -gt 0) {
            $strIncludes = $includes -join ","
            if ($Uri.Contains("?")) {
                $Uri = "{0}&include=" -f $Uri, $strIncludes
            } else {
                $Uri = "{0}?include=" -f $Uri, $strIncludes
            }
        }
    }
    if ($excludeSheetHyperlinks) {
        if ($Uri.Contains("?")) {
            $Uri = "{0}&exclude=" -f $Uri, "sheetHyperlinks"
        } else {
            $Uri = "{0}?exclude=" -f $Uri, "sheetHyperlinks"
        }
    }
    $Payload = [ordered]@{}
    if ($containerId) {         
        $payload.Add("destinationId", $containerId) 
        $payload.Add("destinationType", $containerType)
    } else {
        $payload.Add("destinationType", $containerType)
    }
    $payload.Add("newName", $newSheetName)
    
    $body = $payload | ConvertTo-Json -Compress

    try{
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($response.message = "SUCCESS") {
            if ($passThru) {
                return Get-Smartsheet -sheetId $id
            }
        } else {
            throw $response.message
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS 
    Copy a smartsheet to a new name and/or into a folder.
    .DESCRIPTION
    Copies a smartsheet giving it a new name, or copying it to a folder or copying to a folder with a new name.
    .PARAMETER Id
    The sheet Id of the sheet to be copied.
    .PARAMETER newSheetName
    The name of the new sheet.
    .PARAMETER containerId
    The folder or workspace Id to copy the sheet to.
    .PARAMETER containerType
    One of 'folder', workspace' or 'home' if containerType - 'home' containerId must be omitted.
    'home' is the default value is omitted.
    .PARAMETER includeAll
    Include all elements of the sheet
    .PARAMETER includeAttachments
    Include attachments
    .PARAMETER includeCellLinks
    Include cell links.
    .PARAMETER includeFormatting
    Include formatting
    .PARAMETER includeFilters
    Include filters
    .PARAMETER includeForms
    Include forms
    .PARAMETER includeRuleRecipients
    Include rule recipients
    .PARAMETER includeRules
    Include rules.
    .PARAMETER IncludeShares
    Include Shares
    .PARAMETER excludeSheetHyperlinks
    Exclude sheet hyperlinks.
    .PARAMETER passThru
    Returns the copied Smartsheet object.
    #>
}
function Rename-SmartSheet() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('SheetId')]
        [Uint64]$Id,
        [Parameter(Mandatory = $true)]
        [String]$NewSheetName
    )
    Copy-Smartsheet -Id $Id -newSheetName $NewSheetName -includeAll
    Remove-Smartsheet -Id $Id
    <#
    .SYNOPSIS 
    Rename a Smartsheet
    .DESCRIPTION 
    Renames a smartsheet in the existing container.
    .PARAMETER Id
    Id of the sheet to rename.
    .PARAMETER newSheetname
    New name for the sheet
    .PARAMETER PassThru
    Return the copied smartsheet.
    #>
}

function Move-Smartsheet() {
    [CmdletBinding(DefaultParameterSetName = "props")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('SheetId')]
        [UInt64]$Id,
        [Parameter(ParameterSetName = "container")]
        [string]$containerId,
        [Parameter(ParameterSetName = "container")]
        [ValidateSet('folder','home','workspace')]
        [string]$containerType = "home"
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/move" -f $BaseURI, $Id

    $properties = @{}
    $properties.Add("destinationType", $containerType)        
    if ($containerType -ne 'home') {
        $properties.Add("destinationId", $containerId)
    } 
    $objBody = [PSCustomObject]$properties
    $body = $objBody | ConvertTo-Json -Compress

    try{
        $result = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($result.message -ne "SUCCESS") {
            throw $result.message
        }
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Move a Smartsheet
    .DESCRIPTION
    Move a Smartsheet into a different container.
    .PARAMETER Id
    ID of the the Smartsheet to move.
    .PARAMETER containerId
    Id of the container (folder/workspace) to move the Smartsheet to. 
    if omitted the container is 'home'
    .PARAMETER containerType
    Can be one of 'folder', 'workspace or 'home'. If 'home' then containerId must be omitted.
    The default for this property is 'home' if omitted.
    #>
}

function Get-SortedSmartsheet() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('SheetId')]
        [UInt64]$id,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "Multi")]
        [psobject[]]$sortCriteria,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "single"
        )]
        [Uint64]$columnId,
        [Parameter(ParameterSetName = "single")]
        [ValidateSet("ASCENDING","DESCENDING")]
        [string]$direction = "ASCENDING"
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/sort" -f $BaseURI, $id
    $body = $null
    if ($sortCriteria) {
        $body = $sortCriteria | ConvertTo-Json -Compress
    } else {
        $payload = @{
            sortCriteria = @{
                columnId = $columnId
                direction = $direction
            }
        }
        $body = $payload | ConvertTo-Json -Compress
    }
    try {
        return Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Sort rows in a Smartsheet.
    .DESCRIPTION
    Sort the rows in a smartsheet.
    .PARAMETER id
    Id of the sheet to sort rows in.
    .PARAMETER sortCriteria
    An array of sort criteria objects. The objects should have 2 properties, columnId and direction (ASCENDING or DESCENDING)
    .PARAMETER columnId
    Id of the column to sort on for a single column sort.
    .PARAMETER direction
    The direction of the sort. 
    .OUTPUTS
    Sheet object with the results of the sort operation.
    .NOTES
    If you are retrieving the Smartsheet to process the data within powershell it may be easier and more efficient to do the sorting within powershell.

    For Example:

    PS> $Sheet = Get-Smartsheet -sheetId 465987456
    PS> $Data = $Sheet.ToArray() | Sort-Object -Property Name, HireDate.

    .EXAMPLE
    How to create a multi-sort sortCriteria object.
    In this example we are going to sort a Smartsheet of employee salary information by Department and Salary in descending order.
    To create the criteria create an array of hash table object.
    PS> $sortCriteria - @(
            @{
                sortCriteria = @{
                    columnId = $Sheet.columns.Where({$_.title -eq "Department"}).ColumnId
                    direction = "ASCENDING"
                },
                @{
                {
                    columnId = $Sheet.Columns.Where({$_.title -eq "Salary"}).ColumnId
                    direction = "DESCENDING"
                }
            }
        )
    Now sort the sheet.
    PS >$SortedSheet = $sheet | Get-SortedSmartSheet -SortCriteria $sortCriteria
    #>
}

function Send-SmartsheetViaEmail() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('SheetId')]
        [UInt64]$Id,
        [ValidateSet('EXCEL','PDF','PDF_GANTT')]
        [string]$format,
        [ValidateSet(
            "LETTER","A4","LEGAL","TABLOID"
        )]
        [string]$paperSize,
        [Parameter(Mandatory = $true)]
        [string[]]$To,
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        [string]$Message,
        [switch]$ccMe
    )
    Begin {
        $Headers = Get-Headers 
        $sendTo = @()
        $To | ForEach-Object {
            $sendTo += @{email = $_}
        }       
        $properties = [ordered]@{
            format = $format
            sendTo = $sendTo
            subject = $Subject
        }
        if ($formatDetails) {
            $properties.Add("formatDetails", @{paperSize = $paperSize})
        }
        if ($Message) {
            $properties.Add("message", $Message)
        } else {
            $properties.Add("message", " ")
        }
        if ($ccMe) {
            $properties.Add("ccMe", "true")
        } else {
            $properties.Add("ccMe", "false")
        }

        $objBody = [PSCustomObject]$properties
        $body = $objBody | ConvertTo-Json
    }

    Process {
        $Uri = "{0}/sheets/{1}/emails" -f $BaseURI, $id
        try {
            $result = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
            if ($result.message -eq "SUCCESS") {
                return $true
            } else {
                return $false
            }
        } catch {
            throw $_
        }
    }
    <#
    .SYNOPSIS
    Send a Smartsheet via Email.
    .PARAMETER Id
    Sheet Id if the sheet to send.
    .PARAMETER format
    Attachment format.
    .PARAMETER paperSize
    Set the page size of the attached document.
    .PARAMETER To
    An array of email addresses to send to.
    .PARAMETER Subject
    .Subject of the email.
    .PARAMETER Message
    Body of the email.
    .PARAMETER ccMe
    Send a carbon copy to the sender.    
    #>
}


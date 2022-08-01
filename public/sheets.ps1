
function Get-Smartsheets () {
    <#
    .DESCRIPTION
    Retrieves all the sheets the user has access to.
    #>

    $Uri = "{0}/sheets" -f $BaseURI
    $Headers = Get-Headers

    $response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers

    return $response.data
}

function Get-Smartsheet () {    
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = "id"
        )]
        [string]$id,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "name"
        )]
        [string]$Name
    )

    Begin {
        $Header = Get-Headers
    }

    Process {
        if ($Name) {
            $sheetInfo = Get-Smartsheets | Where-Object {$_.name -eq $Name}
            if (-not $sheetInfo) {
                $msg = "Sheet {0} not found!" -f $Name
                Throw $msg
            }
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
        $Sheet = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Header
        $toPsObject_Script = {
            $psSheet = New-Object System.Collections.Generic.List[psobject]
            foreach ($Row in $sheet.rows) {                
               $Props = [ordered]@{}
               foreach ($Cell in $row.cells) {
                    $PropName = $sheet.columns.Where({$_.id -eq $Cell.columnId}).title
                    $Props.Add($PropName, $Cell.value)                    
                }
                $_row = New-Object -TypeName psobject -Property $props
                $psSheet.add($_row)
            }
            return $psSheet.ToArray()
        }
        $Sheet | Add-Member -MemberType ScriptMethod -Name ToPSObject -Value $toPsObject_Script

        return $Sheet
    }
    <#
    .SYNOPSIS
    Retrieve an indifivual sheet.
    .DESCRIPTION
    Retrieves an individual sheet by either the sheet ID or the Name.
    Note: There can be multiple sheets with the same name. Using the Sheet ID is more accurate!
    The object returned has an additional method ToPSObject, this method returns an array of objects based on the sheet rows and columns.
    .PARAMETER id
    Sheet ID, cannot be used with the Name parameter.
    .PARAMETER Name
    Sheet Name, cannot be used with the id parameter.
    #>
}

function Remove-Smartsheet() {   
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id
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
        [string]$Id,
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
        [switch]$includefilters,
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
        if ($includefilters) {
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
    $properties = @{}
    if ($containerId) {         
        $properties.Add("destinationId", $containerId) 
        $properties.Add("destinationType", $containerType)
    } else {
        $properties.Add("destinationType", $containerType)
    }
    $properties.Add("newName", $newSheetName)
    
    $psBody = [PSCustomObject]$properties

    $body = $psBody | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message = "SUCCESS") {
        if ($passThru) {
            return Get-Smartsheet -sheetId $id
        }
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
    'home' is the default value is ommitted.
    .PARAMETER includeAll
    Include all elements of the sheet
    .PARAMETER includeAttachments
    Include attachments
    .PARAMETER includeCellLinks
    Inlcude cell links.
    .PARAMETER includeFormatting
    Include formatting
    .PARAMETER includefilters
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
    #>
}
function Rename-SmartSheet() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [String]$newSheetname
    )
    Copy-Smartsheet -Id $Id -newSheetName $newSheetname
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
        [string]$Id,
        [Parameter(ParameterSetName = "container")]
        [string]$containerId,
        [Parameter(ParameterSetName = "container")]
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

    $result = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($result.message -ne "SUCCESS") {
        throw $result.message
    }
    <#
    .SYNOPSIS
    Move a Smartsheet
    .DESCRIPTION
    Move a Smartsheet into a different container.
    .PARAMETER Id
    ID of the the Smartsheet to move.
    .PARAMETER containerId
    Id of the container to move th4 smartsheet to. 
    .PARAMETER containerType
    Can be in of 'folder', 'workspace or 'home'. If 'home' then containerId must be omitted.
    The default for this property is 'home' if omitted.
    #>
}

function Get-SortedSmartsheet() {
    [CmdletBinding(DefaultParameterSetName="single")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$id,
        [Parameter(ParameterSetName = "Multi")]
        [psobject[]]$sortCriteria,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "single"
        )]
        [string]$columnId,
        [Parameter(  
            ParameterSetName = "Asc"
        )]
        [switch]$Ascending,
        [Parameter(
            Mandatory = $true,    
            ParameterSetName = "Desc"
        )]
        [switch]$Descending
    )

    $Headers = Get-Headers
    $Uri = "{0}/sheets/{1}/sort" -f $BaseURI, $id
    $body = $null
    if ($sortCriteria) {
        $body = $sortCriteria | ConvertTo-Json -Compress
    } else {
        $propreties = @{}
        $propreties.Add("columnId", $columnId)
        if ($Descending) {
            $propreties.Add("direction","DESCENDING")
        } else {
            $propreties.Add("direction","ASCENDING")
        }
        $objBody = [PSCustomObject]$propreties
        $body = $objBody | ConvertTo-Json -Compress
    }

    return Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body

    <#
    .SYNOPSIS
    Sort rows in a Smartsheet.
    .DESCRIPTION
    Sort the rows in a smartsheet.
    .PARAMETER id
    Id of the sheet to srot rows in.
    .PARAMETER sortCriteria
    An array of sort criteria objects. The objects shoud have 2 properties, columnId and direction (ASCENDING or DESCENDING)
    .PARAMETER columnId
    Id of the column to sort on for a single column sort.
    .PARAMETER Ascending
    Sort the column Ascending (default)
    .PARAMETER Descending
    Sort the column Descending
    .OUTPUTS
    Sheet object with the results of the sort operation.
    .EXAMPLE
    How to create a multi-sort sortCriteria object.
    I this example we are going to sort a Smartsheet of employee salary information by Department and Salary in descending order.
    To create the sort Criteria, 1st create an array.
    PS> $sortCriteria - @()
    The create a criteria objects from the columns collection of a sheet object named $Sheet.
    PS> $criteria - [PSCustomObject]@{
        columnId = $Sheet.columns.Where({$_.title -eq "Department"}).ColumnId
        direction = "ASCENDING"
    }
    Add this to the array.
    $sortCriteria += $criteria
    Create another criteria object
    PS >$criteria = [PSCustomObject]@{
        columnId = $Sheet.Columns.Where({$_,title -eq "Salary"}).ColumnId
        direction = "DESCENDING"
    }
    Add this to the array
    PS >$sortCriteria += $criteria
    Now sort the sheet.
    PS >$SortedSheet = $sheet | Get-SortedSmartSheet -SortCriteria $sortCriteria
    #>
}
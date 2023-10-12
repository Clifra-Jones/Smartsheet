function Get-SmartsheetWorkspaces() {

    $Uri = "{0}/workspaces" -f $BaseURI

    $Headers = Get-Headers

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.data
    } catch {
        Throw $response.message
    }
    <#
    .SYNOPSIS
    Retrieve a list of Smartsheet Workspaces.    
    #>
}

function Add-SmartsheetWorkspace() {
    [CmdletBinding(DefaultParameterSetName='default')]
    Param(
        [Parameter(Mandatory)]
        [string]$Name,
        [psobject[]]$Folders,
        [psobject[]]$Reports,
        [psobject[]]$Sheets,
        [psobject[]]$Sights,
        [psobject[]]$Templates,
        [Parameter(ParameterSetName='IncludeAll')]
        [switch]$IncludeAll,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeAttachments,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeBrands,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeCellLinks,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeData,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeDiscussions,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeFilters,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeForms,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeRuleRecipients,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeRules,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeShares,
        [switch]$ExcludeCellLinksRemap,
        [switch]$ExcludeReportsRemap,
        [switch]$ExcludeSheetHyperlinkRemap,
        [switch]$ExcludeSightsRemap
    )

    $Uri = "{0}/workspaces" -f $BaseURI

    $Headers = Get-Headers
    
    $Includes = $null
    $Excludes = $null

    if ($IncludeAll) {
        $Includes = "all"
    } else {
        if ($IncludeAttachments) {
            $Includes = "attachments"
        }
        if ($IncludeCellLinks) {
            if ($Includes) {
                $Includes = "{0},cellLinks" -f $Includes
            } else {
                $Includes = "cellLinks"
            }
        }
        if ($IncludeBrands) {
            if ($Includes) {
                $Includes = "{0},brands" -f $Includes
            } else {
                $Includes = "brands"
            }
        }
        if ($IncludeData) {
            if ($Includes) {
                $Includes = "{0},data" -f $Includes
            } else {
                $Includes = "data"
            }
        }
        if ($IncludeDiscussions) {
            if ($Includes) {
                $Includes = "{0},discussions" -f $Includes
            } else {
                $Includes = "discussions"
            }
        }
        if($IncludeFilters) {
            if ($Includes) {
                $Includes =     "{0},filters" -f $Includes
            } else {
                $Includes = "filters"
            }
        }
        if ($IncludeForms) {
            if ($Includes) {
                $Includes = "{0},forms" -f $Includes
            } else {
                $Includes = "forms"
            }
        }
        if ($IncludeRuleRecipients) {
            if ($includes) {
                $Includes = "{)},ruleRecipients" -f $Includes
            } else {
                $Includes = "ruleRecipients"
            }
        }
        if ($IncludeRules) {
            if ($Includes) {
                $Includes = "{0},rules" -f $Includes
            } else {
                $includes = "rules"
            }
        }
        if ($IncludeShares) {
            if ($Includes) {
                $Includes = "{0},shares" -f $Includes
            } else {
                $Includes = "shares"
            }
        }

        if ($ExcludeCellLinksRemap) {
            $Excludes = "cellLinks"
        }
        if ($ExcludeReportsRemap) {
            if ($Excludes) {
                $Excludes = "{0},reports" -f $Excludes
            } else {
                $Excludes = "reports"
            }
        }
        if ($ExcludeSheetHyperlinkRemap) {
            $Excludes = "{0},sheetHyperLinks" -f $Excludes
        } else {
            $Excludes = "sheetHyperLinks"
        }
        if ($ExcludeSightsRemap) {
            if ($Excludes) {
                $Excludes = "{0},sights" -f $Excludes
            } else {
                $Excludes = "sights"
            }
        }
    }

    if ($includes) {
        $Uri = "{0}?include={1}" -f $Url, $Includes
    }

    if ($Excludes) {
        if ($Includes) {
            $Url = "{0}&skipRemap={1}" -f $Excludes
        } else {
            $Uri = "{0}?skipRemap={1}" -f $Uri, $Excludes
        }
    }

    $payload = @{
        name = $Name
    }

    if ($Folders) {
        $payload.Add("folders", $Folders)
    }

    if ($Reports) {
        $Payload.Add("reports", $Reports)
    }

    if ($Sheets) {
        $payload.Add("sheets", $Sheets)
    }

    if ($Sights) {
        $payload.Add("sights", $Sights)
    }

    if ($Templates) {
        $payload.Add("templates", $Templates)
    }

    $Body = $payload | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $Body
        return $response
    } catch {
        Throw $response.message
    }
    <#
    .SYNOPSIS
    Add a new Smartsheet Workspace.
    .DESCRIPTION
    Add a new Smartsheet Workspace to the given account using the settings provided.
    .PARAMETER Name
    The name of the Workspace.
    .PARAMETER Folders
    An array of folder objects to add to the Workspace.
    .PARAMETER Reports
    An array of report objects to add to the Workspace.
    .PARAMETER Sheets
    An array of sheet objects to add to the Workspace.
    .PARAMETER Sights
    An array dashboards to add to the Workspace.
    .PARAMETER Templates
    An array templates to add to the Workspace.
    .PARAMETER IncludeAll
    Include all of the below elements in the Workspace.
    .PARAMETER IncludeAttachments
    Include attachments.
    .PARAMETER IncludeBrands
    Include brands.
    .PARAMETER IncludeCellLinks
    Include cell links.
    .PARAMETER IncludeData
    Include data.
    .PARAMETER IncludeDiscussions
    Include discussions.
    .PARAMETER IncludeFilters
    Include filters.
    .PARAMETER IncludeForms
    Include forms.
    .PARAMETER IncludeRuleRecipients
    Include recipients.
    .PARAMETER IncludeRules
    Include rules.
    .PARAMETER IncludeShares
    Include shares.
    .PARAMETER ExcludeCellLinksRemap
    Exclude cell link remaps.
    .PARAMETER ExcludeReportsRemap
    Exclude reports remaps.
    .PARAMETER ExcludeSheetHyperlinkRemap
    Exclude Sheet Hyperlink remaps.
    .PARAMETER ExcludeSightsRemap
    Exclude dashboard remaps.
    .OUTPUTS
    Object containing a Workspace object for the newly created workspace.
    #>
}

function Get-SmartsheetWorkspace() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias('WorkspaceId')]
        [UInt64]$Id,
        [switch]$IncludeSource,   
        [switch]$IncludeDistributionLink,
        [switch]$IncludeOwnerInfo,
        [switch]$IncludeSheetVersion,
        [switch]$IncludePermaLinks,
        [switch]$LoadNestedFolder
    )

    $Uri = "{0}/workspaces/{1}" -f $BaseURI, $Id

    $Headers = Get-Headers

    $includes = $null

    if ($IncludeSource) {
        $Includes = "source"
    }
    if ($IncludeDistributionLink) {
        if ($Includes) {
            $Includes = "{0},distributionLinks" -f $Includes
        } else {
            $Includes = "distributionLinks"
        }
    }
    if ($IncludeSheetVersion) {
        if ($Includes) {
            $Includes = "{0},sheetVersion" -f $Includes
        } else {
            $Includes = "sheetVersion"
        }
    }
    if ($IncludePermaLinks) {
        if ($Includes) {
            $Includes = "{0},permalinks" -f $Includes
        } else {
            $Includes = "permalinks"
        }
    }
    if ($IncludeOwnerInfo) {
        if ($Includes) {
            $Includes = "{0},ownerInfo" -f $Includes
        } else {
            $Includes = "ownerInfo"
        }
    }

    $LoadALL = $false

    if ($LoadNestedFolder) {
        $LoadALL = $true
    }

    if ($includes) {
        $Uri = "{0}?include={1}" -f $Uri, $Includes
    }

    if ($LoadALL) {
        if ($includes) {
            $Uri = "{0}?loadAll={1}" -f $Uri, $LoadAll
        } else {
            $Uri = "{0}?loadAll={1}" -f $Uri, $LoadALL
        }
    }

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response
    } catch {
        throw $response.message
    }
    <#
    .SYNOPSIS
    Retrieve a workspace.
    .DESCRIPTION
    Retrieve a workspace object.
    .PARAMETER WorkspaceId
    The ID of the workspace to retrieve.
    .PARAMETER IncludeSource
    Include the Source object indicating which object the folder was created from, if any.
    .PARAMETER IncludeDistributionLink
    INclude distribution links,
    .PARAMETER IncludeOwnerInfo
    Include owner information.
    .PARAMETER IncludeSheetVersion
    Include sheet version
    .PARAMETER IncludePermaLinks
    Include permalinks.
    .OUTPUTS 
    A workspace object.
    #>
}

function Remove-SmartSheetWorkspace() {
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [UInt64]$Id
    )

    Begin {
        $Headers = Get-Headers
    }

    Process {
        $WorkspaceName = (Get-Workspace -WorkspaceId $Id).Name
        $URI = "{0}/workspaces/{1}" -f $BaseURI, $Id

        if ($PSCmdlet.ShouldProcess("Remove", "Workspace:$WorkspaceName")) {
            try {
                Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
            } catch {
                throw $response.message
            }
        }
    }
    <#
    .SYNOPSIS 
    Delete a Smartsheet workspace.
    .DESCRIPTION
    Deletes the specified workspace.
    .PARAMETER Id
    The Id of thw workspace to delete.    
    #>
}

function Set-SmartSheetWorkspace() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [Uint64]$Id,
        [Parameter(Mandatory)]
        [string]$Name
    )

    $Uri = "{0}/workspaces/{1}" -f $BaseURI, $Id

    $Headers = Get-Headers

    $payload = @{
        name = $Name
    }

    $Body = $payload | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $Body
        return $response.result
    } catch {
        throw $response.message
    }
    <#
    .SYNOPSIS
    Rename a workspace
    .DESCRIPTION
    Rename a workspace with teh specified name.
    .PARAMETER Id
    The Id of thw workspace to rename.
    .PARAMETER Name
    The new name of the workspace.
    .OUTPUTS
    Object containing the renamed workspace.
    #>
}

function Copy-SmartsheetWorkspace() {
    [CmdletBinding(DefaultParameterSetName='default')]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [Uint64]$Id,
        [Parameter(Mandatory)]
        [string]$NewName,
        [Uint64]$DestinationId,
        [ValidateSet("Folder","Home","Workspace")]
        [ValidateScript(
            {
                ($_ -in "Folder","Workspace") -and $DestinationId
            }
        )]
        [string]$DestinationType,
        [Parameter(ParameterSetName='IncludeAll')]
        [switch]$IncludeAll,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeAttachments,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeBrands,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeCellLinks,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeData,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeDiscussions,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeFilters,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeForms,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeRuleRecipients,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeRules,
        [Parameter(ParameterSetName='includeSome')]
        [switch]$IncludeShares,
        [switch]$ExcludeCellLinksRemap,
        [switch]$ExcludeReportsRemap,
        [switch]$ExcludeSheetHyperlinkRemap,
        [switch]$ExcludeSightsRemap       
    )

    $Uri = "{0}/workspaces/{1}/copy" -f $BaseURI, $Id

    $Includes = $null
    $Excludes = $null

    if ($IncludeAll) {
        $Includes = "all"
    } else {
        if ($IncludeAttachments) {
            $Includes = "attachments"
        }
        if ($IncludeCellLinks) {
            if ($Includes) {
                $Includes = "{0},cellLinks" -f $Includes
            } else {
                $Includes = "cellLinks"
            }
        }
        if ($IncludeBrands) {
            if ($Includes) {
                $Includes = "{0},brands" -f $Includes
            } else {
                $Includes = "brands"
            }
        }
        if ($IncludeData) {
            if ($Includes) {
                $Includes = "{0},data" -f $Includes
            } else {
                $Includes = "data"
            }
        }
        if ($IncludeDiscussions) {
            if ($Includes) {
                $Includes = "{0},discussions" -f $Includes
            } else {
                $Includes = "discussions"
            }
        }
        if($IncludeFilters) {
            if ($Includes) {
                $Includes =     "{0},filters" -f $Includes
            } else {
                $Includes = "filters"
            }
        }
        if ($IncludeForms) {
            if ($Includes) {
                $Includes = "{0},forms" -f $Includes
            } else {
                $Includes = "forms"
            }
        }
        if ($IncludeRuleRecipients) {
            if ($includes) {
                $Includes = "{)},ruleRecipients" -f $Includes
            } else {
                $Includes = "ruleRecipients"
            }
        }
        if ($IncludeRules) {
            if ($Includes) {
                $Includes = "{0},rules" -f $Includes
            } else {
                $includes = "rules"
            }
        }
        if ($IncludeShares) {
            if ($Includes) {
                $Includes = "{0},shares" -f $Includes
            } else {
                $Includes = "shares"
            }
        }

        if ($ExcludeCellLinksRemap) {
            $Excludes = "cellLinks"
        }
        if ($ExcludeReportsRemap) {
            if ($Excludes) {
                $Excludes = "{0},reports" -f $Excludes
            } else {
                $Excludes = "reports"
            }
        }
        if ($ExcludeSheetHyperlinkRemap) {
            $Excludes = "{0},sheetHyperLinks" -f $Excludes
        } else {
            $Excludes = "sheetHyperLinks"
        }
        if ($ExcludeSightsRemap) {
            if ($Excludes) {
                $Excludes = "{0},sights" -f $Excludes
            } else {
                $Excludes = "sights"
            }
        }
    }

    if ($includes) {
        $Uri = "{0}?include={1}" -f $Url, $Includes
    }

    if ($Excludes) {
        if ($Includes) {
            $Url = "{0}&skipRemap={1}" -f $Excludes
        } else {
            $Uri = "{0}?skipRemap={1}" -f $Uri, $Excludes
        }
    }

    $payload = @{
        newName = $NewName
    }

    if ($DestinationId) {
        $payload.Add("destinationIf", $DestinationId)
    }

    if ($DestinationType) {
        $payload.Add("desctinationType", $DestinationType)
    }

    $Body = $payload | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method POST - Uri $Uri -Headers $Header -Body $Body
        return $response
    } catch {
        throw $response.message
    }
    <#
    .SYNOPSIS
    Copies a workspace.
    .DESCRIPTION
    Copies a workspace to the specified destination.
    .PARAMETER Id
    The Id of the workspace to copy.
    .PARAMETER NewName
    The new name of the workspace.
    .PARAMETER DestinationId
    The Id of the destination container (when copying or moving a sheet or a folder). Required if destinationType is "folder" or "workspace". 
    If destinationType is "home", this value must be null.
    .PARAMETER DestinationType
    Type of the destination container.
    .PARAMETER IncludeAll
    Include all of the below elements in the Workspace.
    .PARAMETER IncludeAttachments
    Include attachments.
    .PARAMETER IncludeBrands
    Include brands.
    .PARAMETER IncludeCellLinks
    Include cell links.
    .PARAMETER IncludeData
    Include data.
    .PARAMETER IncludeDiscussions
    Include discussions.
    .PARAMETER IncludeFilters
    Include filters.
    .PARAMETER IncludeForms
    Include forms.
    .PARAMETER IncludeRuleRecipients
    Include recipients.
    .PARAMETER IncludeRules
    Include rules.
    .PARAMETER IncludeShares
    Include shares.
    .PARAMETER ExcludeCellLinksRemap
    Exclude cell link remaps.
    .PARAMETER ExcludeReportsRemap
    Exclude reports remaps.
    .PARAMETER ExcludeSheetHyperlinkRemap
    Exclude Sheet Hyperlink remaps.
    .PARAMETER ExcludeSightsRemap
    Exclude dashboard remaps.
    .OUTPUTS
    Object containing a workspace object for the new workspace destination.
    #>
}

function Get-SmartsheetWorkspaceFolders() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [Uint64]$Id
    )

    Begin {
        $Headers = Get-Headers
    }

    Process{
        $Uri = "{0}/workspaces/{1}/folders?includeAll=true" -f $BaseURI, $Id

        try {
            $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
            return $response.data
        } catch {
            throw $response.message
        }
    }
    <#
    .SYNOPSIS
    Retrieve workspace folders.
    .DESCRIPTION
    Retrieve a collection of the top level folders in a workspace.
    .PARAMETER Id
    The Id of the workspace to retrieve folders from.
    .OUTPUTS
    An array of folder objects.
    .NOTES
    The returned collection consists of abbreviated folder objects. These object contain only the id, name, and permalink properties.
    You cannot return a recursive list with this function. To get a recursive list of subfolder use the Get-SmartsheetFolders function and provide the
    Id of one of the top level folders returned by this function.
    #>
}

function Add-SmartsheetWorkspaceFolder() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [Uint64]$Id,
        [Parameter(Mandatory)]
        [string]$Name
    )

    Begin {
        $Headers = Get-Headers
    }

    Process {
        $Uri = "{0}/workspaces/{1}/folders" -f $BaseURI, $Id

        $payload = @{
            name = $Name
        }

        $Body = $payload | ConvertTo-Json -Compress

        try {
            $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $Body
            return $response.result
        } catch {
            throw $response.message
        }
    }
    <#
    .SYNOPSIS
    Create a folder in a Smartsheet workspace.
    .DESCRIPTION
    Create a top level folder in a workspace.
    .PARAMETER Id
    The Id of the workspace to create the folder in.
    .PARAMETER Name
    The name of the folder.
    .OUTPUTS
    An object containing the newly created folder.
    .NOTES
    This function can only create top level folder. To create a subfolder us the Add-SmartsheetFolder function.
    #>
}


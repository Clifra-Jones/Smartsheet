function Get-SmartsheetWorkspaces() {

    $Uri = "{0}/workspaces" -f $BaseURI

    $Headers = Get-Headers

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.data
    } catch {
        Throw $response.message
    }
}

function Add-SmartsheetWorkspace() {
    [CmdletBinding(DefaultParameterSetName='default')]
    Param(
        [Parameter(Mandatory)]
        [string]$Name,
        [string]$PermaLink,
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

    if ($PermaLink) {
        $payload.Add("permalink", $PermaLink)
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
}

function Get-SmartsheetWorkspace() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$WorkspaceId,
        [switch]$IncludeSource,   
        [switch]$IncludeDistributionLink,
        [switch]$InlcudeOwnerInfo,
        [switch]$IncludeSheetVersion,
        [switch]$IncludePermaLinks,
        [switch]$LoadNestedFolder
    )

    $Uri = "{0}/workspaces/{1}" -f $BaseURI, $WorkspaceId

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
    if ($InlcudeOwnerInfo) {
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
}

function Remove-SmartSheetWorkspace() {
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [string]$Id
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
}

function Set-SmartSheetWorkspace() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [string]$Id,
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
}

function Copy-SmartsheetWorkspace() {
    [CmdletBinding(DefaultParameterSetName='default')]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [string]$Id,
        [Parameter(Mandatory)]
        [string]$NewName,
        [string]$DestinationId,
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
}

function Get-SmartsheetWorkspaceFolders() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [string]$Id
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
}

function Add-SmartsheetWorkspaceFolder() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias("WorkspaceId")]
        [string]$Id,
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
}


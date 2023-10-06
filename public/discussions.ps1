
function Get-SmartsheetDiscussions() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$Id,
        [switch]$includeAllComments,
        [switch]$includeAttachments
    )

    $Headers = Get-Headers -AuthOnly
    $Uri = "{0}/sheets/{1}/discussions" -f $BaseUri, $Id

    if ($includeAllComments) {
        $include = "comments"
    }

    if ($includeAttachments) {
        if (-not [string]::IsNullOrWhiteSpace($include)) {
            $include += ','
        }
        $include += "attachment"
    }

    if ($include) {
        $Uri = "{0}?include={1}" -f $Uri, $include 
    }

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.data   
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Get Smartsheet Discussions
    .DESCRIPTION 
    Gets all Discussions attached to the Smartsheet. 
    Returns both Sheet level and Row level discussions.
    .PARAMETER Id
    The Smartsheet Id
    .PARAMETER includeAllComments
    Include all comments. By default only the Discussion objects are returned.
    .PARAMETER includeAttachments
    Include all attachment. By default only the Discussion objects are returned.
    .OUTPUTS
    A Smartsheet discussion object.
    #>
}

function Get-SmartsheetDiscussion() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$discussionId
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/discussions/{2}" -f $BaseURI, $id, $discussionId

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
    <#
    .SYNOPSIS
    Get a smartsheet Discussion
    .PARAMETER Id
    The Smartsheet Id.
    .PARAMETER discussionId
    The discussion Id.
    .OUTPUTS
    A smartsheet Discussion object.
    #>
}

function Add-SmartsheetDiscussion() {
    [CmdletBinding(DefaultParameterSetName = 'default')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [string]$text
    )
    $Headers = Get-Headers -AuthOnly
    $uri = "{0}/sheets/{1}/discussions" -f $BaseURI, $id

    $Headers.Add("Content-Type", "application/json")
    $payload = [ordered]@{
        comment = @{
            text = $text
        }
    }
    $body = $payload | ConvertTo-Json
    try {
        $response = Invoke-RestMethod -Method Post -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            return $response.result
        } else {
            return $response.message
        }
    } catch {
        Write-Host $_.ErrorDetails.Message -ForegroundColor Red
        return
    }

    <#
    .SYNOPSIS
    Create a new Smartsheet discussion.
    .DESCRIPTION
    Creates a new discussion at the sheet level.
    To attach a file or URL to the comment use the New-SmartsheetCommentAttachment function.
    .PARAMETER Id
    The smartsheet Id.
    .PARAMETER text
    The text of the comment.
    .OUTPUTS    
    A Smartsheet Discussion object.
    #>
}

function Remove-SmartsheetDiscussion() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetid')]
        [UInt64]$id,
        [Parameter(Mandatory = $true)]
        [UInt64]$discussionId
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/discussions/{2}" -f $BaseURI, $id, $discussionId

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        if ($response.message -eq 'SECCESS') {
            return $true
        } else {
            return $false
        }
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $false
    }
    <#
    .SYNOPSIS
    Remove a Smartsheet discussion.
    .DESCRIPTION
    Removes a discussion from a smartsheet. This will remove all comments and attachments.
    .PARAMETER id
    The Smartsheet Id.
    .PARAMETER discussionId
    The discussion Id.
    .OUTPUTS
    Boolean indicating success of failure. True = success.
    #>
}

function Get-SmartsheetRowDiscussions() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$id,
        [Parameter(Mandatory = $true)]
        [UInt64]$rowId,
        [switch]$includeComments,
        [switch]$includeAttachments
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/rows/{2}/discussions" -f $BaseURI, $id, $rowId

    $includes = @()
    if ($includeComments) {
        $includes += "comments"
    }
    if ($includeAttachments) {
        $includes += "attachments"
    }
    if ($includes.Length -gt 0) {
        $strIncludes = $includes -join ","
        $uri = "{0}?include={1}" -f $Uri, $strIncludes
    }

    try{
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.data
    } catch {
        throw $_.Exception.message
    }
    <#
    .SYNOPSIS
    Get Smartsheet row discussions
    .DESCRIPTION
    Gets discussions attached to a row.
    .PARAMETER id
    The Smartsheet Id.
    .PARAMETER rowId
    The Row id.
    .PARAMETER includeComments
    Include comments. By default only the discussion objects are returned.
    .PARAMETER includeAttachments
    include attachments. By default only the discussion objects are returned.
    .OUTPUTS
    A smartsheet discussion object.
    #>
}

function Add-SmartsheetRowDiscussion() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$id,
        [Parameter(Mandatory = $true)]
        [UInt64]$rowId,
        [Parameter(Mandatory=$true)]
        [string]$text
    )

    $Uri = "{0}/sheets/{1}/rowes/{2}/discussions" -f $BaseURI, $id, $rowId

    $Headers = Get-Headers
    $payload = @{
        comment = @{
            text = $text
        }
    }

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -Body $payload
        return $response
    } catch {
        throw $_.Exception.Message
    }

    <#
    .SYNOPSIS
    Creates a new Smartsheet row discussion.
    .DESCRIPTION
    Creates a discussion on the specified row.
    To attach a file or URL to a comment use the New-SmartsheetCommentAttachment function.
    .PARAMETER id
    The smartsheet Id.
    .PARAMETER rowId
    The Row Id.
    .PARAMETER text
    The text of the comment.
    .OUTPUTS
    A smartsheet discussion object.
    #>
}

# Comment functions.
function Get-SmartSheetComment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [uint64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$commentId
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/comments/{2}" -f $BaseURI, $Id, $commentId

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response
    } catch {
        throw $_.Exception.Message
    }
    <#
    .SYNOPSIS 
    Gets a smartsheet discussion comment
    .PARAMETER Id
    The smartsheet Id.
    .PARAMETER commentId
    The comment Id.
    .OUTPUTS
    A smartsheet comment object
    #>
}

function Set-SmartSheetComment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [uint64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$commentId,
        [Parameter(Mandatory = $true)]
        [string]$text
    )

    $Headers = Get-Headers

    $Uri = "{0}/sheets/{1}/comments/{2}" -f $BaseURI, $Id, $commentId

    $payload = @{
        text = $text
    }
    $body = $payload | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            return $response.result
        } else {
            throw $response.message
        }
    } catch {
        throw $_.Exception.Message
    }
    <#
    .SYNOPSIS
    Updates a smartsheet comment.
    .DESCRIPTION
    Updates the text of a Smartsheet comment. Only the owner of the comment can update the text.
    .PARAMETER Id
    The Smartsheet Id.
    .PARAMETER commentId
    The Command d.
    .PARAMETER text
    The updated text for the comment.
    .OUTPUTS
    A smartsheet comment object
    #>
}

function Remove-SmartsheetComment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [string]$commentId
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/comments/{2}" -f $BaseURI, $Id, $commentId

    try {
        $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
        if ($response.message = 'SUCCESS') {
            return $true
        } else {
            return $false
        }
    } catch {
        throw $_.Exception.Message
    }
    <#
    .SYNOPSIS
    Remove a smartsheet comment.
    .PARAMETER Id
    The smartsheet Id.
    .PARAMETER commentId
    The comment Id.
    .OUTPUTS
    Boolean indicating success of failure. True = success.
    #>
}

function Add-SmartsheetComment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$discussionId,
        [Parameter(Mandatory = $true)]
        [string]$text
    )

    $Headers = Get-Headers

    $Uri = "{0}/sheets/{1}/discussions/{2}/comments" -f $BaseURI, $id, $discussionId

    $payload = @{
        text = $text
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
        throw $_.Exception.Message
    }
    <#
    .SYNOPSIS
    Adds a new comment to a smartsheet discussion.
    .PARAMETER Id
    The Smartsheet Id.
    .PARAMETER discussionId
    The discussion Id.
    .PARAMETER text
    The test of the new comment.
    .OUTPUTS
    A smartsheet comment object.
    #>
}

function Copy-SmartsheetDiscussions() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sourceSheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$targetSheetId
    )

    # Get the source Discussions
    $sourceDiscussions = Get-SmartsheetDiscussions -sheetId $sourceSheetId -includeAllComments

    # Process the Sheet level dicussions 1st.
    $sourceSheetDiscussions = $sourceDiscussions.Where({$_.parentType -eq 'SHEET'})
    foreach ($sourceSheetDiscussion in $sourceSheetDiscussions) {
        $newDiscussion = Add-SmartsheetDiscussion -sheetId $targetSheetId -text $sourceSheetDiscussion.title
        # Process the comments.
        foreach ($comment in $sourceSheetDiscussions.comments | Select-Object -Skip 1) {
            [void](Add-SmartsheetComment -sheetId $targetSheetId -discussionId $newDiscussion.id -text $comment.text)
        }        
    }
    <#
    .SYNOPSIS
    Copy discussions from one Smartsheet to another.
    .DESCRIPTION
    Copy all discussions from a source Smartsheet to another smartsheet.
    .PARAMETER Id
    The source Smartsheet Id.
    .PARAMETER targetSheetId
    The Target Smartsheet Id.
    #>
}
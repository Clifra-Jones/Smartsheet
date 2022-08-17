
function Get-SmartsheetDiscussions() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
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
    } catch {
        $ErrorDetails = $_.ErrorDetails | ConvertFrom-Json
        Write-Host $ErrorDetails.Message -ForegroundColor Red
        exit
    }
    return $response.data
}

function Get-SmartsheetDiscussion() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$discussionId
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/discussions/{2}" -f $BaseURI, $id, $discussionId

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

function New-SmartsheetDiscussion() {
    [CmdletBinding(DefaultParameterSetName = 'default')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$text,
        [PSObject]$Path
    )
    $Headers = Get-Headers -AuthOnly
    $uri = "{0}/sheets/{1}/discussions" -f $BaseURI, $id

    If ($Path) {
        $Headers.Add("Content-Type", "multipart/form-data")
        [byte[]]$file = [System.IO.File]::ReadAllBytes($Path)
        $payload = [ordered]@{
            discussion = [ordered]@{
                comment = @{
                    text = $text
                }
            }
            file = $file
        }
        $body = $payload | ConvertTo-Json
        try {
            $response = Invoke-RestMethod -Method Post -Uri $Uri -Headers $Headers -body $form
            if ($response.message -eq "SUCCESS") {
                return $response
            } else {
                return $response
            }
        } catch {
            throw $_.ErrorDetails.Message
        }
    } else {
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
    }
}

function Remove-SmartsheetDiscussion() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetid')]
        [string]$id,
        [Parameter(Mandatory = $true)]
        [string]$discussionId
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
}

function Get-SmartsheetRowDiscussions() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$id,
        [Parameter(Mandatory = $true)]
        [string]$rowId,
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
        $uri = "{0}?include={1}" -f $strIncludes
    }

    try{
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.data
    } catch {
        throw $_.Exception.message
    }
}

function New-SMartsheetRowDiscussion() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$id,
        [Parameter(Mandatory = $true)]
        [string]$rowId,
        [Parameter(Mandatory=$true)]
        [string]$text,
        [string]$Path
    )

    $Uri = "{0}/sheets/{1}/rowes/{2}/discussions" -f $BaseURI, $id, $rowId

    if ($Path) {
        $Headers = Get-Headers -ContentType "multipart/form-data"        

        $form = @{
            discussion = @{
                comment = @{
                    text = $text                    
                }
                type = "application/json"
            }
            file = Get-Item -Path $Path
            type = "multipart/form-data"
        }

        try {
            Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Form $form
            return $response
        } catch {
            throw $_.Exception.Message
        }
    } else {
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
    }
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
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$commentId
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/comments/{2}" -f $BaseURI, $Id, $commentId

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response
    } catch {
        throw $_.Exception.Message
    }
}

function Set-SmartSheetComment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$commentId,
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
}

function Remove-SmartsheetComment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$commentId
    )

    $Headers = Get-Headers -AuthOnly

    $Uri = "{0}/sheets/{1}/comments/{2}" -f $BaseId, $Id, $commentId

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
}

function New-SmartsheetComment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$discussionId,
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
}
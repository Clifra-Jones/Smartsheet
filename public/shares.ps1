function Add-SmartsheetShare() {
    [CmdletBinding(DefaultParameterSetName = 'none')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [ValidateSet(
            "ADMIN","COMMENTER","EDITOR","EDITOR_SHARE","OWNER","VIEWER"
        )]
        [string]$accessLevel,
        [Parameter(
            Mandatory = $true
        )]
        [switch]$sendEmail,
        [string]$email,
        [string]$subject,
        [string]$message,
        [switch]$ccMe
    )

    Begin {
        $Headers = Get-Headers

        $properties = [ordered]@{
            accessLevel = $accessLevel
            email = $email
        }               
        if ($subject) {$properties.Add("subject", $subject)}
        if ($message) {$properties.Add("message",$message)}
        if ($ccMe) {$properties.Add("ccMe", "true")}
    
        $objBody = [PSCustomObject]$properties
        $body = $objBody | ConvertTo-Json -Compress
    }

    Process {
        $Uri = "{0}/sheets/{1}/shares" -f $BaseURI, $Id
        if ($sendEmail) {
            $Uri = "{0}?sendEmail=true" -f $Uri
        }

        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            return $true
        } else {
            return $false
        }
    }
    <#
    .SYNOPSIS
    Share a smartsheet.
    .DESCRIPTION
    Adds a sharing object to the smartsheetoptionally sending an email to the person the sheet is shared with.
    .PARAMETER Id
    Sheet id of the sheet to share.
    .PARAMETER accessLevel
    Access level to grant to the user.
    .PARAMETER sendEmail
    Send an email to the user you are sharing the sheet with.
    .PARAMETER email
    Emai address of the person you are sharing the sheet with.
    .PARAMETER subject
    Subject of the email.
    .PARAMETER message
    Body of the email.
    .PARAMETER ccMe
    send a carbon copy tothe sender.
    #>
}

function Get-SmartsheetShares() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id
    )

    Begin{
        $Headers = Get-Headers
    }

    Process {
        $Uri = "{0}/sheets/{1}/shares" -f $BaseURI, $id
        try {
            $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
            return $response.data
        } catch {
            Throw $_
        }
    }
    <#
    .SYNOPSIS
    Get Smartsheet Shares
    .DESCRIPTION
    Get the Shares for this smartsheet.
    .PARAMETER Id
    Id of the Smartsheet
    .OUTPUTS
    An array of Smartsheet share Object.
    #>
}

function Get-SmartSheetShare() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$shareid
    )

    Begin {
        $Headers = Get-Headers
    }

    Process {
        $Uri = "{0}/sheets/{1}/shares/{2}" -f $BaseURI, $Id, $sheetId
        $response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers
        return $response
    }
    <#
    .SYNOPSIS
    Get a Smartsheet share.
    .DESCRIPTION
    Get an individual share from a Smartsheet.
    .PARAMETER Id
    ID of the Smartsheet.
    .PARAMETER shareid
    The Id of the share.
    .OUTPUTS
    A Smartsheet share object.
    #>
}

function Remove-SmartsheetShare() {
    [CmdletBinding(DefaultParameterSetName = 'none')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$shareId
    )

    Begin {
        $Headers = Get-Headers
    }

    Process {
        $Uri = "{0}/sheets/{1}/shares/{2}" -f $BaseURI, $Id, $shareId
        try {
            $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers -Body $body
        } catch {
            $ErrorDetails = $_.ErrorDetails | ConvertFrom-Json
            Write-Host $ErrorDetails.Message -ForegroundColor Red
            exit
        }
        if ($response.message -eq "SUCCESS") {
            return $true
        } else {
            return $false
        }
    }
    <#
    .SYNOPSIS
    remove a Smartsheet share.
    .PARAMETER Id
    Sheet id of the sheet to share.
    .PARAMETER shareId
    ID of teh share to remove.
    .OUTPUTS
    Boolean indicating success or failure.
    #>
}

function Set-SmartsheetShare() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true            
        )]
        [Alias("sheetId")]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$shareId,
        [Parameter(Mandatory = $true)]
        [ValidateSet(
            "ADMIN","COMMENTER","EDITOR","EDITOR_SHARE","OWNER","VIEWER"
        )]
        [string]$accessLevel
    )
    Begin {
        $Headers = Get-Headers

        $properties = [ordered]@{
            accessLevel = $accessLevel
        }                   
        $objBody = [PSCustomObject]$properties
        $body = $objBody | ConvertTo-Json -Compress

    }

    Process{
        $Uri = "{0}/sheets/{1}/shares/{2}" -f $BaseURI, $id, $shareId
        $response = Invoke-RestMethod -Method PUT -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            return $true
        } else {
            return $false
        }
    }
    <#
    .SYNOPSIS
    Updte the access level of a share.
    .PARAMETER Id
    Id of the Smartsheet.
    .PARAMETER shareId
    ID of the share.
    .PARAMETER accessLevel
    Access level to set.
    .OUTPUTS
    Boolean indicating success or failure.
    #>
}
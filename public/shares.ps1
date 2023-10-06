function Add-SmartsheetShare() {
    [CmdletBinding(DefaultParameterSetName = "default")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [ValidateSet(
            "ADMIN","COMMENTER","EDITOR","EDITOR_SHARE","OWNER","VIEWER"
        )]
        [string]$accessLevel,
        [Parameter(ParameterSetName="sendmail'")]
        [switch]$sendEmail,
        [Parameter(ParameterSetName="sendmail'")]
        [string]$email,
        [Parameter(ParameterSetName="sendmail'")]
        [string]$subject,
        [Parameter(ParameterSetName="sendmail'")]
        [string]$message,
        [Parameter(ParameterSetName="sendmail'")]
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
    Adds a sharing object to the smartsheet optionally sending an email to the person the sheet is shared with.
    .PARAMETER Id
    Sheet id of the sheet to share.
    .PARAMETER accessLevel
    Access level to grant to the user.
    .PARAMETER sendEmail
    Send an email to the user you are sharing the sheet with.
    .PARAMETER email
    Email address of the person you are sharing the sheet with.
    .PARAMETER subject
    Subject of the email.
    .PARAMETER message
    Body of the email.
    .PARAMETER ccMe
    send a carbon copy to the sender.
    #>
}

function Get-SmartsheetShares() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$Id
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
        [Alias('sheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$shareId
    )

    Begin {
        $Headers = Get-Headers
    }

    Process {
        $Uri = "{0}/sheets/{1}/shares" -f $BaseURI, $Id
        if ($shareId) {
            $Uri = "[0}/{1}" -f $shareId
        }
        $response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers
        return $response
    }
    <#
    .SYNOPSIS
    Get a Smartsheet share.
    .DESCRIPTION
    Get an individual share from a Smartsheet.
    .PARAMETER Id
    Id of the Smartsheet.
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
        [Alias('sheetId')]
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$shareId
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
    Remove a Smartsheet share.
    .PARAMETER Id
    Sheet id of the sheet to share.
    .PARAMETER shareId
    Id of the share to remove.
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
        [UInt64]$Id,
        [Parameter(Mandatory = $true)]
        [UInt64]$shareId,
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

function Copy-SmartsheetShares() {
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

    $Shares = (Get-SmartsheetShares -Id $sourceSheetId).Where({$_.accessLevel -ne 'OWNER'})
    foreach ($share in $shares) {
        $targetSheetId | Add-SmartsheetShare -email $share.email -accessLevel $share.accessLevel
    }
    <#
    .SYNOPSIS
    Copies shares from one sheet to another.
    .DESCRIPTION
    Copies the shares from one Smartsheet to another smartsheet.
    .PARAMETER sourceSheetId
    The source Smartsheet Id.
    .PARAMETER targetSheetId
    The Target Smartsheet Id
    #>
}
function Get-SmartsheetAttachments() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias("sheetId")]
        [UInt64]$id        
    )
    $Headers = Get-Headers 
    $Uri = "{0}/sheets/{1}/attachments" -f $BaseURI, $Id

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.data
    } catch {
        $ErrorDetails = $_.ErrorDetails | ConvertFrom-Json
        Write-Host $ErrorDetails.Message -ForegroundColor Red
        exit
    }
    <#
    .SYNOPSIS
    Get a Smartsheet Attachment.
    .PARAMETER id
    Smartsheet Id.
    .OUTPUTS
    An array of Smartsheet attachment objects.
    #>
} 

function Add-SmartsheetAttachment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true            
        )]
        [Alias("sheetId")]
        [UInt64]$Id,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'file'
        )]
        [string]$Path,
        [Parameter(ParameterSetName = 'url')]
        [string]$Url,
        [Parameter(ParameterSetName = 'url')]
        [ValidateSet("BOX_COM", "DROPBOX", "EGNYTE", "EVERNOTE", "FILE", "GOOGLE_DRIVE", "LINK", "ONEDRIVE")]
        [string]$Type = "LINK",
        [Parameter(ParameterSetName = 'url')]
        [ValidateSet("DOCUMENT", "DRAWING", "FOLDER", "PDF", "PRESENTATION", "SPREADSHEET")]
        [string]$subType = "DOCUMENT",
        [Parameter(ParameterSetName='url')]
        [Parameter(ParameterSetName='file')]
        [string]$description,
        [Parameter(ParameterSetName='url')]
        [Parameter(ParameterSetName='file')]
        [string]$name
    )
    
    $Uri = "{0}/sheets/{1}/attachments" -f $BaseURI, $id

    if ($Path) {
        If (Test-Path -Path $Path) {
            $file = Get-ChildItem -Path $Path
            $mimetype = $MimeTypes[$file.Extension]            
        } else {
            Write-Host "File not found!" -ForegroundColor Red
            exit
        }
        $Headers = Get-Headers -ContentType $mimetype -ContentDisposition 'attachment' -filename $file.name       
        $body = [System.IO.File]::ReadAllBytes($path)
        #$config = Read-Config
        #$token = ConvertTo-SecureString -string $config.APIKey -AsPlainText -Force
        try {
            $response = Invoke-RestMethod -Method 'POST' -Uri $Uri -Headers $Headers -Body $body

            if ($response.message -eq "SUCCESS") {
                return $response.result
            } else {
                throw $response.message
            }
        } catch {
            throw $_
        }
    } else {
        $Headers = Get-Headers
        $Properties = [ordered]@{
            attachmentType = $Type
            url = $url
        }
        If ($Type -in 'EGNYTE','GOOGLE_DRIVE') {
            $Properties.Add("attachmentSubType", $subType)
        }
        If ($description) { $Properties.Add('description', $description)}
        if ($name) { $Properties.Add('name', $name)}

        $objBody = [PScustomObject]$properties
        $body = $objBody | ConvertTo-Json -Compress

        Try {
            $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
            if ($response.message -eq "SUCCESS") {
                return $response.result
            } else {
                throw $response.message
            }
        } catch {
            throw $_
        }
    }
    <#
    .SYNOPSIS
    Adds a attachment to a Smartsheet.
    .DESCRIPTION
    Add either a file attachment or a URL attachment. URLs can point to links or cloud service files/folders.
    .PARAMETER Id
    Smartsheet Id.
    .PARAMETER Path
    Path to the file to attach.
    .PARAMETER Url
    URL to the cloud based resource.
    .PARAMETER Type
    The type of URL.
    .PARAMETER subType
    Subtype of URL. Only valid for EGNYTE and GOOGLE_DRIVE types.
    .PARAMETER description
    A description of the attachment.
    .PARAMETER name
    The name of the attachment.
    .OUTPUTS
    A Smartsheet attachment object.
    #>
} 

function Get-SmartsheetAttachment() {
    [CmdletBinding(DefaultParameterSetName = 'default')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [UInt64]$id,
        [Parameter(
            Mandatory = $true
        )]        
        [string]$attachmentId,
        [Parameter(ParameterSetName = 'saveas')]
        [string]$saveAs,
        [Parameter(ParameterSetName = 'bytes')]
        [byte[]]$asByteArray
    )
    $Headers = Get-Headers -AutoOnly
    $Uri = "{0}/sheets/{1}/attachments/{2}" -f $BaseURI, $Id, $attachmentId

    $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
    if ($response) {
        If ($asByteArray) {
            $webResponse = Invoke-WebRequest -Uri $response.url 
            return $webResponse.Content
        } elseif ($saveAs) {
            [void](Invoke-WebRequest -Uri $response.url -OutFile $saveAs)
        } else {
            return $response
        }
    }
    <#
    .SYNOPSIS
    Get a Smartsheet Attachment.
    .DESCRIPTION
    Gets s specific attachment to a Smartsheet.
    .PARAMETER id
    The Smartsheet Id
    .PARAMETER attachmentId
    The attachment Id.
    .PARAMETER saveAs
    Path and filename to save the attachment to.
    .PARAMETER asByteArray
    Returns the attachment as a byte array.
    .OUTPUTS
    If -saveAs and -asByteArray are not specified returns a smartsheet attachment object.

    if -saveAs is specified returns nothing.

    if -asByteArray is specified an array of bytes is returned.
    #>
}
function Remove-SmartSheetAttachment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$Id,
        [Parameter(
            Mandatory = $true
        )]
        [string]$attachmentId
    )

    $Headers = Get-Headers -AutoOnly
    $Uri = "{0}/sheets/{1}/attachments/{2}" -f $BaseURI, $Id, $attachmentId

    $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
    if ($response.message -eq 'SUCCESS') {
        return $true
    } else {
        return $false
    }
    <#
    .SYNOPSIS
    Removed a Smartsheet attachment
    .PARAMETER Id
    The Smartsheet Id.
    .PARAMETER attachmentId
    The attachment Id.
    .OUTPUTS
    Boolean indicating success of failure. True = success.
    #>
}

function Copy-SmartsheetAttachments() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sourceSheetId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$targetSheetId,
        [Parameter(Mandatory = $true)]
        [string]$tempDir        
    )

    $Attachments = Get-SmartsheetAttachments -sheetId $fromSheetId

    foreach ($Attachment in $Attachments) {
        if ($Attachment.attachmentType = "FILE") {
            $outfile = "{0}/{1}" -f $tempDir, $Attachment.name
            [void](Get-SmartsheetAttachment -sheetId $fromSheetId -attachmentId $Attachment.id -saveAs $outfile)
            [void](Add-SmartsheetAttachment -sheetId $toSheetIt -Path $outfile)
            Remove-Item $outfile
        }
    }
    <#
    .SYNOPSIS
    Copies Smartsheet attachments.
    .DESCRIPTION
    Copies all Smartsheet attachments from one sheet to another.
    .PARAMETER sourceSheetId
    The source Smartsheet Id.
    .PARAMETER targetSheetId
    The Target Smartsheet Id.
    .PARAMETER tempDir
    The temporary directory to save the files to. (Linux/Mac = /tmp, Windows = TEMP environment variable.)
    #>
}

function New-SmartSheetCommentAttachment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [Uint64]$id,
        [Parameter(Mandatory = $true)]
        [string]$commentId,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$Url,
        [ValidateSet("BOX_COM", "DROPBOX", "EGNYTE", "EVERNOTE", "FILE", "GOOGLE_DRIVE", "LINK", "ONEDRIVE")]
        [string]$Type = "LINK",
        [ValidateSet("DOCUMENT", "DRAWING", "FOLDER", "PDF", "PRESENTATION", "SPREADSHEET")]
        [string]$subType = "DOCUMENT",
        [string]$description,
        [string]$name
    )

    $Uri = "{0}/sheets/{1}/comments/{2}/attachments" -f $BaseURI, $id, $commentId

    if ($Path) {
        if (Test-Path -Path $Path) {
            $file = Get-ChildItem -Path $Path    
            $mimetype = $mimetypes[$file.Extension]        
        } else {
            Write-Host "File not found!" -ForegroundColor Red
            exit
        }
        $Headers = Get-Headers -ContentType $mimetype -ContentDisposition 'attachment' -filename $file.name
        $body = [System.IO.File]::ReadAllBytes($Path)

        try {
            $response = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers -Body $body
            if ($response.message -eq 'SUCCESS') {
                return $respose.result
            } else {
                throw $response.message
            }
        } catch {
            throw $_.Exception.Message
        }
    } else {
        $Headers = Get-Headers
        $payload = [ordered]@{
            attachmentType = $Type
            url = $url
        }
        if ($Type -in 'EGNYTE','GOOGLE_DRIVE') {
            $payload.Add('attachmentSubType', $subType)
        }
        if ($description) { $payload.Add('description', $description)}
        if ($name) { $payload.Add('name', $name)}

        $body = $payload | ConvertTo-Json -Compress

        try {
            $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
            if ($response.message -eq 'SUCCESS') {
                return $response.result
            } else {
                Throw $response.message
            }
        } catch {
            throw $_.Exception.Message
        }
    }
    <#
    .SYNOPSIS
    Create a smartsheet comment attachment.
    .DESCRIPTION
    Create an attachment tied to a Smartsheet comment.
    .PARAMETER id
    The Smartsheet Id.
    .PARAMETER commentId
    The Smartsheet Comment Id.
    .PARAMETER Path
    The path to the file to attach.
    .PARAMETER Url
    A url to a cloud resource.
    .PARAMETER Type
    The type of Url.
    .PARAMETER subType
    The URL subtype. Only valid for EGNYTE and GOOGLE_DRIVE types.
    .PARAMETER description
    The description of the URL.
    .PARAMETER name
    The name of the Url.
    .OUTPUTS
    A Smartsheet attachment object.
    #>
}


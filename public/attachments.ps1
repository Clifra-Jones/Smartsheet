function Get-SmartsheetAttachments() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias("sheetId")]
        [string]$id        
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
} 

function Add-SmartsheetAttachment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias("sheetId")]
        [string]$Id,
        [string]$Path,
        [string]$Url,
        [ValidateSet("BOX_COM", "DROPBOX", "EGNYTE", "EVERNOTE", "FILE", "GOOGLE_DRIVE", "LINK", "ONEDRIVE")]
        [string]$Type = "LINK",
        [ValidateSet("DOCUMENT", "DRAWING", "FOLDER", "PDF", "PRESENTATION", "SPREADSHEET")]
        [string]$subType = "DOCUMENT",
        [string]$description,
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
        $Headers = Get-Headers -ContentType $mimetype -ContentDisposition 'attachment' -filename $file.Fullname       
        $body = [System.IO.File]::ReadAllBytes($path)
        #$config = Read-Config
        #$token = ConvertTo-SecureString -string $config.APIKey -AsPlainText -Force
        $response = Invoke-RestMethod -Method 'POST' -Uri $Uri -Headers $Headers -Body $body
        

        if ($response.message -eq "SUCCESS") {
            return $response.result
        } else {
            return $false
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
                return false
            }
        } catch {
            Write-Host $response.message -ForegroundColor Red
            exit
        }
    }
} 

function Get-SmartsheetAttachment() {
    [CmdletBinding(DefaultParameterSetName = 'default')]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$id,
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
}

function Copy-SmartsheetAttachments() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$sourceSheetId,
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
}

function New-SmartSheetCommentAttachment() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('sheetId')]
        [string]$id,
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
}


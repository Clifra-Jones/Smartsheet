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

    if ($Path) {
        If (Test-Path -Path $Path) {
            $file = Get-ChildItem -Path $Path
            $mimetype = $MimeTypes[$file.Extension]            
        } else {
            Write-Host "File not found!" -ForegroundColor Red
            exit
        }
        $Headers = Get-Headers -ContentType $mimetype -ContentDisposition 'attachment' -filename $file.Fullname       
<#         $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Authorization", "Bearer VcjJj965oRo3E30PG1sXKkkvAFfFJhEhg5paj")
        $headers.Add("Content-Type", $mimetype)
        $headers.Add("Content-Disposition", "attachment; filename=`"CWilliams.docx`"")
 #>     $Uri = "{0}/sheets/{1}/attachments" -f $BaseURI, $id
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
            attachmentSubType = $subType
            Url = $url
        }
        If ($description) { $Properties.Add('description', $description)}
        if ($name) { $Properties.Add('name', $name)}

        $objBody = [PScustomObject]$properties
        $body = $objBody | ConvertTo-Json -Compress

        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if ($response.message -eq "SUCCESS") {
            return $response.result
        } else {
            return false
        }
    }
} 

function Get-SmartsheetAttachment() {
    [CmdletBinding(DefaultParameterSetName = "saveas")]
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
        [Parameter(Mandatory=$true, ParameterSetName = 'saveas')]
        [string]$saveAs,
    )
    $Headers = Get-Headers -AutoOnly
    $Uri = "{0}/sheets/{1}/attachments/{2}" -f $BaseURI, $Id, $attachmentId

    $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
    if ($response) {
        If ($asByteArray) {
            $webResponse = Invoke-WebRequest -Uri $response.url 
            return $webResponse.Content
        } else {
            [void](Invoke-WebRequest -Uri $response.url -OutFile $saveAs)
        }
    }
}
function Remove-SmartSheetAttachments() {
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
        [string]$fromSheetId,
        [Parameter(Mandatory = $true)]
        [string]$toSheetId,
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
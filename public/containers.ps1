
function Remove-SmartsheetFolder() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$folderId
    )

    $Headers = Get-Headers
    $Uri = "{0}/folders/{0}" -f $BaseURI, $folderId

    $response = Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $Headers
    if ($response.message -eq "SUCCESS") {
        return $true
    } else {
        return $false
    }
    <#
    .SYNOPSIS
    Remove a smartsheet folder
    .DESCRIPTION
    Removes a Smartsheet folder. WARNING: This function does not determine if the folder is empty before removing it. 
    Any contents of the folder will be lost.
    .PARAMETER folderId
    The Id of the folder to be removed.
    #>
}

function Get-SmartsheetFolders() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$folderId
    )

    $Headers = Get-Headers
    $Uri = "{0}/folders/{0}/folders" -f $BaseURI, $folderId

    $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
    return $response 
    <#
    .SYNOPSIS
    Retrieve a list of folders.
    .DESCRIPTION
    Returns an array of subfolder object from an existing folder.
    This will not return subfolders from the Home folder. Use Get-SmartsheetHomeFolders to get this list.
    .PARAMETER folderId
    The folder ID to retrieve subfolders from.
    .OUTPUTS
    An array of folder objects.
    #>   
}

function New-SmartsheetFolder() {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('folderId')]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$folderName
    )

    $Headers = Get-Headers
    $Uri = "{0}/folders/{1}/folders" -f $BaseURI, $Id

    $objBody = [PSCustomObject]@{
        Name = $folderName
    }
    $body = $objBody | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $response.result
    } else {
        return $fasle
    }
    <#
    .SYNOPSIS
    Add a Smartsheet folder.
    .DESCRIPTION
    Add a folder to an existing Smartsheet folder. 
    This function will not add folder to the home folder. Use Add-SmartsheetHomeFolder to add a folder to home.
    This function creates an empty folder. The functionality to create prepopulated folders my be included in the future.
    .PARAMETER Id
    Id of the folder to create the new folder in.
    .PARAMETER folderName
    Name of the new folder.
    .OUTPUTS
    The newly created folder object.
    #>
}

function Get-SmartsheetHome() {
    $Headers = Get-Headers 
    $Uri = "{0}/home" -f $BaseURI

    $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers

    return $response
    <#
    .SYNOPSIS
    Return all home objects.
    .DESCRIPTION
    Gets a nested list of all Home objects, including dashboards, folders, reports, sheets, templates, and workspaces, as shown on the "Home" tab.
    .OUTPUTS
    A nested array of home objects.
    #>
}

function Get-SmartsheetHomeFolders() {

    $Headers = Get-Headers
    $Uri = "{0}/home/folders" -f $BaseURI

    try {
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
        return $response.data
    } catch {
        throw $_
    }
    <#
    .SYNOPSIS
    Return folder in the home tab.
    .DESCRIPTION
    Gets a list of folders in your Home tab. The list contains an abbreviated Folder object for each folder.
    .OUTPUTS
    An array of abbreviated folder objects.
    #>
}

function New-SmartSheetHomeFolder() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$folderName
    )
    $Headers = Get-Headers
    $Uri = "{0}/home/folders" -f $BaseURI

    $objName = [PSCustomObject]@{
        Name = $folderName
    }
    $body = $objName | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -eq "SUCCESS") {
        return $response.result
    } else {
        return $false
    }
    <#
    .SYNOPSIS
    Create a folder int he hoome tab.
    .DESCRIPTION
    Create a new empty folder in the home tab.
    .PARAMETER folderName
    The name of the new folder.
    .OUTPUTS
    The newly created folder object.
    #>
}

function Get-SmartsheetFolder() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$folderId
    )

    $Headers = Get-Headers
    $Uri = "{0}/folders/{0}" -f $BaseURI, $folderId

    $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
    return $response    
    <#
    .SYNOPSIS
    Returns a folder object.
    .DESCRIPTION
    Returns the folder object specified by the folder Id.
    .PARAMETER folderId
    ID of the folder to retrieve.
    .OUTPUTS
    Folder object.
    #>
}

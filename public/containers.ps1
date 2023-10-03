using namespace System.Collections.Generic
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
    [CmdletBinding(DefaultParameterSetName='default')]    
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName
        )]
        [Alias('folderId')]
        [string]$Id,
        [Parameter(
            DontShow,
            ValueFromPipelineByPropertyName)]
        [string]$Name,        
        [switch]$recurse
    )

    Begin {
        $Headers = Get-Headers

        $folderList = [List[psobject]]::New()
        function Get-Subfolders() {
            Param (
                [string]$SubFolderId,
                [string]$name
            )

            $SubFolders = Get-SmartsheetFolders -folderId $SubfolderId
            foreach ($Subfolder in $Subfolders) {
                $Subfolder | Add-Member -MemberType NoteProperty -Name 'Fullname' -Value ("{0}/{1}" -f $Name, $Subfolder.Name)
                $folderList.Add($subfolder)
                Get-Subfolders -SubFolderId $Subfolder.Id -name $Subfolder.name
            }
        }
    }

    Process {
        $Uri = "{0}/folders/{1}/folders" -f $BaseURI, $Id

        try {
            $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers
            if ($recurse) {
                # Get the Name of the root folder
                $folders = $response.data
                foreach ($folder in $folders) {  
                    $folder | Add-Member -MemberType NoteProperty -Name FullName -Value $folder.name 
                    $folderList.Add($folder)                        
                    $Subfolders = (Get-Subfolders -SubFolderId $folder.Id -name $folder.name)
                }
            
                #return $FolderList.ToArray()
                return $folderList
            } else {
                return $response.data
            }
        } catch {
            throw $_
        }
    }
    <#
    .SYNOPSIS
    Retrieve a list of folders.
    .DESCRIPTION
    Returns an array of subfolder object from an existing folder.
    This will not return subfolders from the Home folder. Use Get-SmartsheetHomeFolders to get this list.
    .PARAMETER folderId
    The folder ID to retrieve subfolders from.
    .PARAMETER recurse
    This will return a list including all subfolders. This adds a new 'FullName'property which will be the full path of the folder from the folderID provided or the home folder.
    The returned array will look like this.
                    id name      permalink                                                                   FullName
                    -- ----      ---------                                                                   --------
        7306313035212676 folder2   https://app.smartsheet.com/folders/V5g7j44M52jf9GgHgJcM2XPC8VP7mrq33VXCg741 folder2
        5582828558673796 folder3   https://app.smartsheet.com/folders/fPpQw2qh24hFcVg9jRjCQqxH73Q85QVR243x77w1 folder2/folder3
        6462437860894596 folder4   https://app.smartsheet.com/folders/C73GCm6M4hxcQ3f38cr3x57hwGGjpqp4mWr8mGx1 folder3/folder4
        2079509634672516 folder3.1 https://app.smartsheet.com/folders/5P4JJfwF9Jj6rFvXJ94Gw9gG7rFm9cxM3QrxCxp1 folder2/folder3.1

    You can filter the results by comparing the full or partial path to the FullName property.
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
    You cannot get a recursive list from the home folder. To get a recursive list of subfolders you must use the Get-SMartsheetFolders
    function and specify a folder Id of one of the folder in this list.
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
    $Uri = "{0}/folders/{1}" -f $BaseURI, $folderId

    
    try {
        
        $response = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers       
        return $response
    } catch {
        throw $_
    }
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

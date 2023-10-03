using namespace System.Collections.Generic
<#
    Modulename: Smartsheet
    Description: Powershell module to interact with the SMartsheet API
    Object specific functions are in the ./public folder.
#>

#$BaseURI = "https://api.smartsheet.com/2.0"
$Mimes = import-csv -Path "$PSScriptRoot/private/mimetypes.csv"
$script:MimeTypes = [Dictionary[string,string]]::New()
foreach ($mime in $mimes) {
  $script:MimeTypes.Add($mime.Extension, $mime.MIMEType)
}

# dot source the following files.
. $PSScriptRoot/private/private.ps1
. $PSScriptRoot/public/images.ps1
. $PSScriptRoot/public/objects.ps1
. $PSScriptRoot/public/columns.ps1
. $PSScriptRoot/public/containers.ps1
. $PSScriptRoot/public/rows.ps1
. $PSScriptRoot/public/sheets.ps1
. $PSScriptRoot/public/shares.ps1
. $PSScriptRoot/public/attachments.ps1
. $PSScriptRoot/public/discussions.ps1
. $PSScriptRoot/public/search.ps1
. $PSScriptRoot/public/workspaces.ps1

# Setup Functions
function Set-SmartsheetAPIKey () {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$APIKey
    )
    $objConfig = @{
        APIKey = $APIKey        
    }
    $configPath = "{0}/.smartsheet" -f $HOME

    if (-not (Test-Path -Path $configPath)) {
        New-Item -Path $configPath -ItemType:Directory | Out-Null
    }

    $objConfig | ConvertTo-Json | Out-File -FilePath "$configPath/config.json" -Force

    <#
    .SYNOPSIS
    Set the API key.
    
    .DESCRIPTION
    Creates a file in the user profile folder in the .smartsheet folder named config.json.
    This file contains the users Smaretsheet API Token.

    .PARAMETER APIKey
    The Smartsheet API Access Token.
    #>    
} 

# End Setup Functions



#Export Functions
function Export-SmartSheet() {
    [CmdletBinding(DefaultParameterSetName = "default")]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true
        )]
        [psobject]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$SheetName,
        [Parameter(ParameterSetName='folder')] 
        [string]$FolderId,
        [Parameter(ParameterSetName='workspace')]
        [string]$WorkspaceId,
        [int]$headerRow,
        [int]$primaryColumn,
        [ValidateSet(
            "Replace",
            "Rename"
        )]
        [string]$overwriteAction,
        [string]$overwriteSheetId

    )
    
    $Headers = Get-Headers -ContentType 'text/csv' -ContentDisposition 'attachment'

    # convert input to csv        
    $inputCsv = $input | ConvertTo-Csv
    $inputString = $inputCsv | Out-String
    $encoder = New-Object System.Text.UTF8Encoding
    $body = $encoder.GetBytes($inputString)
    If ($FolderId) {
        $Uri = "{0}/folders/{1}/sheets/import?sheetName={2}" -f $BaseURI, $folderId, $SheetName
    } elseIf ($WorkspaceId) {
        $Uri = "{0}/workspaces/{1}/sheets/import?sheetName={2}" -f $BaseURI, $WorkspaceId, $SheetName
    }
    else {
        $Uri = "{0}/sheets/import?sheetName={1}" -f $BaseURI, $SheetName
    }
    if ($headerRow -ge 0) {
        $Uri = "{0}&headerRowIndex={1}" -f $Uri, $headerRow
    }
    if ($PrimaryColumn) {
        $Uri = "{0}&primaryColumnIndex={1}" -f $Uri, $PrimaryColumn
    }
<#      
    $Headers.Remove("Content-Type")
    $Headers.Add("Content-Type", "text/csv")
    $Headers.add("Content-Disposition", "attachment")
#>        
    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -ne "SUCCESS") {
        Throw "Import failed! $($_.Exception.Message)"
    }
    else {
        if ($overwriteAction) {       
            $newSheetid = $response.result.id
            Copy-SmartsheetShares -sourceSheetId $overwriteSheetId -targetSheetId $newSheetid
            Copy-SmartsheetAttachments -sourceSheetId $overwriteSheetId -targetSheetId $newSheetid
            Copy-SmartsheetDiscussions -sourceSheetId $overwriteSheetId -targetSheetId $newSheetid

            switch ($overwriteAction) {                
                "Replace" {
                    Remove-Smartsheet -Id $overwriteSheetId
                }
                "Rename" {
                    $sheetName = (Get-Smartsheet -id $overwriteSheetId).Name
                    $strDate = Get-Date -Format "yyyyMMdd-HHmm"
                    $newSheetName = "Copy Of_{0}_{1}" -f $SheetName, $strDate
                    Rename-SmartSheet -Id $overwriteSheetId -newSheetName $newSheetName
                }
            }
            $result = $response.result
        }
    }
    return $result

    <#
        .SYNOPSIS
        Exports a powershell array into a new Smartsheet.

        .DESCRIPTION 
        Exports an array of PSObjects into a new smartsheet. This function will always create a new sheet even if
        there is a sheet of the same name. The API will attempt to determine column types.
        To prevent Sheets of the same name being created, use the -overwriteAction and -overwriteSheetId parameters.

        .PARAMETER InputObject
        Array of object to create the Smartsheet from.

        .PARAMETER SheetName
        The name of the new Smartsheet. 
        
        .PARAMETER FolderId
        The folder ID of the folder to create the Smartsheet in. This can either be a folder fromthe home location or a folder in a Workspace.
        Use the Get-SmartsheetFolder or Get-SmartsheetWorkspaceFolders to get the Folder Id.

        .PARAMETER WorkspaceId
        The workspace to create the Smartsheet in. Use the Get-SmartsheetWorkspaces function to get the Workspace Id.
        This will create the sheet in the root of the workspace. To create a sheet in a folder in a workspace, 
        specify the folder ID of the folder inside the workspace. 
        At this time you cannot get a recursive list of all folders in a workspace, You can get a recursive list of all subfolders of a workspace folder.
        Use the Get-SmartsheetFolders function, specifying the top level folder ID and the Recursive property.

        
        .PARAMETER headerRow
        Row to use for column headers. 
        All rows above this row are ignored.        
        If ommitted the first row will be used. A value of -1 will create default headers in the form Column1, Column1, etc.
        
        .PARAMETER primaryColumn
        The column to use as the primary column. default is the 1st column.
        
        .PARAMETER overwriteAction
        What to do if the sheet name already exists. Choices are Replace or Rename.
        Multiple sheets can have the same name in a folder. If you omit this parameter a sheet with the same name will be created.
        Sheets are uniquely identified by the sheet ID.

        .PARAMETER overwriteSheetId
        Because you can have multiple sheets with the same name (with different sheet Ids) you must provide the sheet Id for the overwrite action.

        .EXAMPLE
        Create a new sheet in the home folder.
        PS> $ObjectArray | Export-Smartsheet -SheetName "MyNewSheet"
        .EXAMPLE
        Create a new sheet in the folder.
        $Folder = Get-SmartsheetHomeFolders -Recurse | Where-object {$_.FullName like "Inventory/Westcoast"}
        PS> $objectArray | Export-Smartsheet -SheetName "MyNewSheet" -folder 'myfolder1/myfolder2'
        .EXAMPLE
        Create a new sheet in a workspace folder.
        $Workspace = Get-SmartsheetWorkspaces | Where-Object {$_.Name -eq 'Accounting'}
        $APFolder = Get-SMartsheetWorkspaceFolders -WorkspaceId $Workspace.Id | Where-Object ($_.Name -eq 'Accounts Payable')
        $PaymentsFolder = Get-SmartsheetFolders -FolderId $APFolder.Id -Recurse | Where-Object {$_.FullName -eq "Microsoft/Payments"}
        $ObjectArray | Export-Smartsheet -Sheetname 'July Payments' -folder $PaymentsFolder.Id
        .EXAMPLE
        Overwrite an existing sheet of the same name.
        PS> $objectArray | Export-Smartsheet -SheetName "MySheet" -overwriteAction Replace -overwriteSheetId $oldsheet.Id
        .EXAMPLE
        Rename an exsiting sheet and create a new sheet from the Object.
        The renamed sheet name will be in the format originalname_yyyyMdd-HHmm.
        PS> $objectArray | Export-Smartsheet -SheetName "MySheet" -overwriteAction Rename -overwriteSheetId $oldsheet.Id

    #>
}

function Update-Smartsheet() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true
        )]
        [psObject]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$sheetId        
    )

    $sheet = Get-Smartsheet -id $sheetId

    # Verify columns, column names must match properties of the objects in the array.
    $columns = $sheet.columns

    $PrimaryColumn = ($Columns.Where({$_.Primary -eq $true}))[0]
    $pcIndex = $columns.IndexOf($PrimaryColumn)

    $properties = $input[0].PSObject.properties | Select-Object Name

    if ($columns.count -ne $properties.count) {
        throw "Column count does not match!"
    }

    foreach ($prop in $properties) {
        $col = $columns.Where({$_.title -eq $prop.Name})
        if (-not $col) {
            throw "Column names do not match"
        }
    }

    foreach ($object in $input) {
        $props = $object.PSObject.properties | Select-Object Name, Value
        $cells = @()
        foreach ($prop in $props) {
            $index = $props.indexOf($prop)
            $column = $sheet.columns[$index]
            switch ($column.type) {
                'PICKLIST' {
                    if ($prop.value) {
                        if ($column.options) {                        
                            if ($prop.value -notin $column.options) {
                                $options = $column.options
                                $options += $prop.value
                                $column.options = $options
                                $sheet = Set-SmartsheetColumn -Id $sheet.id -ColumnId $column.id -column $column -PassThru
                            }
                        }
                    }
                }
                'CHECKBOX' {
                    If ($prop.value) { 
                        $Prop.value = [System.Convert]::ToBoolean($prop.value) 
                     } else {
                        $prop.value = $false
                     }
                    #$prop.value = [System.Convert]::ToBoolean($prop.value)
                }
            }
            if ($null -eq $Prop.value) {
                $Prop.value = " "
            }
            $cell = New-SmartSheetCell -columnId $column.id -value $prop.value
            $cells += $cell
        }
        # Does the row exist based on the primary Column column
        $row = $sheet.rows.Where({$_.cells[$pcIndex].value -eq $props[$pcIndex].value})
        if ($row) {
            $sheet = Set-SmartsheetRow -id $sheetId -rowId $row.Id -Cells $cells -PassThru
        } else {       
            $index = $input.IndexOf($object)
            if ($index -lt ($sheet.rows.Count -1)) {
                $siblingRowId = $sheet.rows[$index].id
                $sheet = Add-SmartsheetRow -sheetId $sheet.id -siblingRowId $siblingRowId -cells $cells -location:below -PassThru
            } else {
                $sheet = Add-SmartsheetRow -sheetId $sheet.id -cells $cells -PassThru
            }
        }
    }
    <#
    .SYNOPSIS
    Update a smartsheet.
    .DESCRIPTION
    Update a Smartsheet from an array of powershell objects.
    1. The number and names of the columns is the same as the properties in the object in the array.
    2. The primary column is used to identify rows to be updated and must be unique.
    If condition 1 isn't met, and error will be thrown.
    .PARAMETER InputObject
    An array of powershell objects.
    .PARAMETER sheetId
    The Id of the sheet to update.
    
    #>
}

# End Export Functions
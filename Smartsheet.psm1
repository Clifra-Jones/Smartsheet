using namespace System.Collections.Generic
<#
    Modulename: Smartsheet
    Description: Powershell module to interact with the SMartsheet API
    Object specific functions are in the ./public folder.
#>


#$BaseURI = "https://api.smartsheet.com/2.0"
$Mimes = import-csv -Path "$PSScriptRoot/private/mimetypes.csv"
$script:MimeTypes = [Dictionary[string, string]]::New()
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
    [CmdletBinding(DefaultParameterSetName = 'default')]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$APIKey,
        [switch]$Secure
    )

    If ($Secure) {
        $Secretin = @{
            APIKey = $APIKey
        }
        $Secret = $SecretIn | ConvertTo-Json
        Set-Secret -Name SmartSheet -Secret $Secret 
        $objConfig = @{
            APIKey = 'secure'
        }
    }
    else {
        $objConfig = @{
            APIKey = $APIKey        
        }
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
    This file contains the users Smartsheet API Token.

    .PARAMETER APIKey
    The Smartsheet API Access Token.

    .PARAMETER Secure
    This switch instructs the function to save the APIKey into a secure vault.
    This vaule is created and managed by the modules:
    Microsoft.PowerShell.SecretManagement
    Microsoft.PowerShell.SecretStore
    See the note below on how to setup a secret store.

    .NOTES
    To use a secret store to securely store you API key, you must install the required modules
    and configure the store and create a vaule.

    To install the modules:
    Install-Module Microsoft.PowerShell.SecretManagement
    Install-Module Microsoft.PowerShell.SecretStore

    Configure the secret store:
    You must configure your secret store and set the authenticatio, there are several parameters:
    -Authentication: This can ne either 'Password" or 'None' 
    You should set a password for interactive sessions. For automation on a secure device you can set this to 'None'
    -Interaction: This can be set to 'Prompt' or 'None' If -Authentication is set to 'Password' This must be set to 'Prompt', id Authentication is set to 'None' this must ne set to 'None'.
    -PasswordTimeOut: The time in seconds befopre the password must be re-entered. The default is 600 seconds.
    -Password: The password to unlock the vault. If omitted you will be prompted for the password.

    Register the vault:
    You must register a vault to store your secrets. At this time the Secret Store does not support multiple vaules. 
    While you "can" register multiple vaults the store will just be duplicated. This is a known issue with the SecretStore and may be corrected in the future.
    Only create 1 vault and register it as the default vault.

    Register-SecretVault -Name "MyVault" -DefaultVault -ModuleName Microsoft.Powershell.SecretStore

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
        [Parameter(ParameterSetName = 'folder')] 
        [Uint64]$FolderId,
        [Parameter(ParameterSetName = 'workspace')]
        [UInt64]$WorkspaceId,
        [int]$headerRow,
        [int]$primaryColumn
    )
    
    $Headers = Get-Headers -ContentType 'text/csv' -ContentDisposition 'attachment'

    # convert input to csv        
    $inputCsv = $input | ConvertTo-Csv
    $inputString = $inputCsv | Out-String
    $encoder = New-Object System.Text.UTF8Encoding
    $body = $encoder.GetBytes($inputString)
    If ($FolderId) {
        $Uri = "{0}/folders/{1}/sheets/import?sheetName={2}" -f $BaseURI, $folderId, $SheetName
    }
    elseIf ($WorkspaceId) {
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
       
    $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
    if ($response.message -ne "SUCCESS") {
        Throw "Import failed! $($_.Exception.Message)"
    }
    return $result

    <#
        .SYNOPSIS
        Exports a powershell array into a new Smartsheet.

        .DESCRIPTION 
        Exports an array of PSObjects into a new smartsheet. This function will always create a new sheet even if
        there is a sheet of the same name. The API will attempt to determine column types.
        
        .PARAMETER InputObject
        Array of object to create the Smartsheet from.

        .PARAMETER SheetName
        The name of the new Smartsheet. 
        
        .PARAMETER FolderId
        The folder ID of the folder to create the Smartsheet in. This can either be a folder from the home location or a folder in a Workspace.
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
        If omitted the first row will be used. A value of -1 will create default headers in the form Column1, Column1, etc.
        
        .PARAMETER primaryColumn
        The column to use as the primary column. default is the 1st column.

        .NOTES
        As mentioned this function ALWAYS creates a new sheet even if the name already exists. Name IS NOT the unique identifyer in smartsheet, the sheet ID is.
        If you want to update a shee use the Update-SmartSheet funcction.
        
        .EXAMPLE
        Create a new sheet in the home folder.
        $ObjectArray | Export-Smartsheet -SheetName "MyNewSheet"
        .EXAMPLE
        Create a new sheet in the folder.
        $Folder = Get-SmartsheetHomeFolders -Recurse | Where-object {$_.FullName like "Inventory/Westcoast"}
        $objectArray | Export-Smartsheet -SheetName "MyNewSheet" -folder 'myfolder1/myfolder2'
        .EXAMPLE
        Create a new sheet in a workspace folder.
        $Workspace = Get-SmartsheetWorkspaces | Where-Object {$_.Name -eq 'Accounting'}
        $APFolder = Get-SMartsheetWorkspaceFolders -WorkspaceId $Workspace.Id | Where-Object ($_.Name -eq 'Accounts Payable')
        $PaymentsFolder = Get-SmartsheetFolders -FolderId $APFolder.Id -Recurse | Where-Object {$_.FullName -eq "Microsoft/Payments"}
        $ObjectArray | Export-Smartsheet -Sheetname 'July Payments' -folder $PaymentsFolder.Id        

    #>
}

function Update-Smartsheet() {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true
        )]
        [psObject[]]$InputObject,
        [Parameter(Mandatory = $true)]
        [UInt64]$sheetId,
        [switch]$UseRowId,
        [switch]$PassThru     
    )

    $sheet = Get-Smartsheet -id $sheetId

    $RowsToUpdate = @()
    $RowsToInsert = @()
    $RowsToAdd = @()

    # Verify columns, column names must match properties of the objects in the array.
    $columns = $sheet.columns

    $PrimaryColumn = ($Columns.Where({ $_.Primary -eq $true }))[0]
    $pcIndex = $columns.IndexOf($PrimaryColumn)

    $properties = $input[0].PSObject.properties | Select-Object Name
    If ($UseRowId) {
        $properties = $properties | Select-Object -Skip 1
    } 

    if ($columns.count -ne $properties.count) {
        throw "Column count does not match!"
    }

    foreach ($prop in $properties) {
        $col = $columns.Where({ $_.title -eq $prop.Name })
        if (-not $col) {
            throw "Column names do not match"
        }
    }

    [int]$RowsProcessed = 0
    [int]$RowsUpdated = 0

    foreach ($object in $input) {     
        $RowhasChanged = $False
        Write-Host ("Row:{0}" -f ($Input.IndexOf($Object) + 1))
        $RowsProcessed += 1
        $props = $object.PSObject.properties | Select-Object Name, Value
        if ($UseRowId) {
            $props = $props | Select-Object -Skip 1
        }
        $cells = @()
        
        try {
        foreach ($prop in $props) {
            Write-Host $prop.name
            #$Activity = "Row:{0}" -f ($Input.IndexOf($Object) + 1)
            #Write-Host $Activity
            #Write-Progress -Activity $Activity -Status $Prop.Name
            $index = $props.indexOf($prop)
            $column = $sheet.columns[$index]
            if ($column.formula) {
                Continue
            }
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
                    }
                    else {
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
        } catch {
            throw $_
        }


        if ($UseRowId) {
            # Get to row to update based on the RowId Property
            $row = $sheet.rows.where({ $_.id -eq $Object.RowId })
        }
        else {
            # Get the row based on a unique primary column.
            # Does the row exist based on the primary Column column
            $row = $sheet.rows.Where({ $_.cells[$pcIndex].value -eq $props[$pcIndex].value })
        }
        If ($row -is [array]) {
            throw "Primary column is not unique!"
        }
        if ($row) {
            # Check to see if the data has changed
            foreach ($rowCell in $row.Cells) {
                $i = $row.Cells.indexOf($rowCell)
                if ($cells[$i].value -ne $rowCell.value) {
                    $rowcell.value = $cells[$i].value
                    $RowhasChanged = $true
                }
            }
            if ($RowhasChanged) {
                $RowsToUpdate += $row
            }
            # if ($RowNeedsUpdating) {
            #     $RowsUpdated += 1
            #     $Time2 = Measure-Command {[void](Set-SmartsheetRow -id $sheetId -rowId $row.Id -Cells $cells)}
            #     $Time2.TotalMilliseconds
            # }
        }
        else { 
            $NewRow = [PSCustomObject]@{
                cells = $cells
            }
            # $index = $input.IndexOf($object)
            # if ($index -lt ($sheet.rows.Count - 1)) {
            #     $SiblingId = $sheet.rows[$index].id
            #     $NewRow | Add-Member -MemberType NoteProperty -Name "siblingId" -Value $SiblingId
            #     $RowsToInsert += $row
            # }
            # else {
            #$NewRow | Add-Member -MemberType NoteProperty -Name "ToBottom" -Value $true
            $RowsToAdd += $NewRow
            # }
            #$RowsUpdated += 1
        }
    }

    # Update sheet
    if ($RowsToUpdate.Count -gt 0) {
        #[array]$arrRows = $RowsToUpdate.ToArray()
        [void](Set-SmartsheetRows -Id $sheetId -Rows ($RowsToUpdate | Select-Object id,cells))
    }
    if ($RowsToAdd.count -gt 0) {
        [void](Add-SmartsheetRows -Id $sheetId -Rows $RowsToAdd)
    }
 
    $Stats = [PSCustomObject]@{
        RowsProcessed = $RowsProcessed
        RowsChanged = ($RowsUpdated.Count + $RowsToAdd.Count)
    }

    $Stats | Format-Table

    if ($PassThru) {
        $newSheet = Get-Smartsheet -SheetId $sheet.id
        return $newSheet
    }

    <#
    .SYNOPSIS
    Update a smartsheet.
    .DESCRIPTION
    Update a Smartsheet from an array of powershell objects.
    1. The number and names of the columns is the same as the properties in the object in the array.
    2. If the array objects do not contain a property RowId then primary column is used to identify rows to be updated and must be unique.
    3. If the Array objects contain the property RowId then this is used to identify the row to be updated. The primary column does not have to be unique.    
    .PARAMETER InputObject
    An array of powershell objects.
    .PARAMETER sheetId
    The Id of the sheet to update.
    .PARAMETER UseRowId
    This assumes that the objects in the array have a property called RowId which contains the Smartsheet row Id for the data,
    This will update the row associated with that Row Id.
    .PARAMETER PassThru
    Return the updated sheet object.
    .NOTES
    To return an array of objects from a smartsheet that contains the row id of the row the values are associated with use the ToArray method
    on the sheet object returned by Get-Smartsheet passing a value of $true. $sheet.ToArray($true).

    Updating sheet rows based on the primary column is maintained for backward compatibility. You should use the RowId method in any new projects
    as it is more accurate and does not depend on the primary column being unique. Smartsheet does not force uniqueness on the primary column.

    Updating sheets by primary column may be removed from future versions of this module.
    .EXAMPLE
    Update the rows in the smartsheet based on Primary Columns.
    
    $Array | Update-Smartsheet -SheetId $sheet.id 
    .EXAMPLE
    Update the rows in the smartsheet based on the RowId property.

    $Array | Update-Smartsheet -SheetId $Sheet.Id -UseRowId

    .EXAMPLE
    Update the rows in the smartsheet based on the RowId property and return the updated sheet.
    $Array | Update-Smartsheet -SheetId $Sheet.Id -UseRowId -PassThru
    #>
}

# End Export Functions
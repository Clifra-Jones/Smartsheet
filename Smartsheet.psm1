using namespace System.Collections.Generic
<#
    Modulename: Smartsheet
    Description: Powershell module to interact with the SMartsheet API
    Object specific functions are in the ./public folder.
#>
$script:BaseURI = "https://api.smartsheet.com/2.0"

#Private function
function Read-Config () {
    $ConfigPath = "$home/.smartsheet/config.json"
    $config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
    return $config
}

function ConvertTo-UTime () {
    Param(
        [datetime]$DateTime
    )

    $uTime = ([System.DateTimeOffset]$DateTime).ToUnixTimeMilliseconds() / 1000

    return $Utime
}

function ConvertFrom-UTime() {
    Param(
        [decimal]$Utime
    )

    [DateTime]$DateTime = [System.DateTimeOffset]::FromUnixTimeMilliseconds(1000 * $Utime).LocalDateTime

    return $DateTime
}

function Get-Headers() {
    $config = Read-Config
    $Authorization = "Bearer {0}" -f $Config.APIKey
    $Headers = @{
        "Authorization" = $Authorization
        "Content-Type" = 'application/json'
    }
    return $Headers
}

# end private functions

function Set-APIKey () {
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

    $objConfig | ConvertTo-Json | Out-File -FilePath "$configPath/config.json"

    <#
    .SYNOPSIS
    Set the API key.
    
    .DESCRIPTION
    Creates a file in the user profile folder in the .smartsheet folder named config.json.
    This file contains the users Meraki API Key.

    .PARAMETER APIKey
    The Smartsheet API Access Token.
    #>    
} 

function Export-SmartSheet() {
    Param(
        [Parameter(
            ValueFromPipeline = $true
        )]
        [psobject]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$SheetName,
        [string]$Folder,
        [int]$headerRow,
        [int]$primaryColumn,
        [ValidateSet(
            "Replace",
            "Rename"
        )]
        [string]$overwriteAction,
        [string]$overwriteSheetId
    )
    
    Begin {
        $Headers = Get-Headers
        $folderId = $null
        if ($folder) {
            if ($folder.Contains("/")) {                
                # Get an object that contains a nested list of objects & folders.
                $Uri = "{0}/home" -f $BaseURI
                $rootfolders = Get-SmartsheetHome
                # Split the folder path into its parts.
                $Folders = $folder.Split("/")
                $currentFolder = $rootfolders
                $folders | ForEach-Object {
                    $currentFolder = $currentFolder.Where($_.Name -eq $_)                        
                }
                #$FolderId = $currentFolder.Id
            } else {
                #get a folder off the root
                $rootfolders = Get-SmartsheetHomeFolders
                $currentFolder = $folders.Where({$_.name -eq $Folder})
            }
            $folderId = $currentFolder.Id
        }
        $ArList = [System.Collections.Generic.List[psobject]]::New()
    }
    Process {
        $ArList.Add($inputObject)
    }

    End {
        # convert input to csv
        $ArInput = $ArList.ToArray()
        $inputCsv = $ArInput | ConvertTo-Csv
        $inputString = $inputCsv | Out-String
        $encoder = New-Object System.Text.UTF8Encoding
        $body = $encoder.GetBytes($inputString)
        If ($FolderId) {
            $Uri = "{0}/folders/{1}/sheets/import?sheetName={2}" -f $BaseURI, $folderId, $SheetName
        } else {
            $Uri = "{0}/sheets/import?sheetName={1}" -f $BaseURI, $SheetName
        }
        if ($headerRow -ge 0) {
            $Uri = "{0}&headerRowIndex={1}" -f $Uri, $headerRow
        }
        if ($PrimaryColumn) {
            $Uri = "{0}&primaryColumnIndex={1}" -f $Uri, $PrimaryColumn
        }
        $Headers.Remove("Content-Type")
        $Headers.Add("Content-Type","text/csv")
        $Headers.add("Content-Disposition","attachment")
        $response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $body
        if (-not $response.message -eq "SUCCESS") {
            Throw "Import failed! $($_.Exception.Message)"
        } else {
            switch ($overwriteAction) {
                "Replace" {
                    Remove-Smartsheet -Id $overwriteSheetId
                }
                "Rename" {
                    $sheetName = (Get-Smartsheet -id $overwriteSheetId).Name
                    $strDate = Get-Date -Format "yyyyMMdd-HHmm"
                    $newSheetName = "copyOf_{0}_{1}" -f $SheetName, $strDate
                    Rename-SmartSheet -Id $overwriteSheetId -newSheetName $newSheetName
                }
            }
            $result = $response.result
        }
        return $result
    }
    <#
        .SYNOPSIS
        Exports a powershell array into a new Smartsheet.

        .DESCRIPTION 
        Exports an array of PSObjects into a smartsheet. This function will always create a new sheet even if
        there is a sheet of the same name. The API will attempt to determine column types.

        .PARAMETER InputObject
        Object to create the Smartsheet from

        .PARAMETER SheetName
        The name of the new Smartsheet. 
        
        .PARAMETER Folder
        The name and path to the folder to create the Smartsheet in in the for folder1/folder2/folder3
        folder(s) must exist.
        
        .PARAMETER headerRow
        Row to use for column headers. 
        All rows above this row are ignored.        
        If ommitted the first row will be used. A value of -1 will crete defaule headers in the form Column1, Column1, etc.
        
        .PARAMETER primaryColumn
        The column to use as the primary column. default is the 1st column.

        .EXAMPLE
        PS> $ObjectArray | Export-Smartsheet -SheetName "MyNewSheet"

        .EXAMPLE
        PS> $objectArray | Export-Smartsheet -SheetName "MyNewSheet" -folder 'myfolder1/myfolder2'
    #>
}
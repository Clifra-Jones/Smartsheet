using namespace System.Collections.Generic
<#
    Modulename: Smartsheet
    Description: Powershell module to interact with the SMartsheet API
    Object specific functions are in the ./public folder.
#>

$BaseURI = "https://api.smartsheet.com/2.0"

# dot source the following files.
. $PSScriptRoot/private/private.ps1
. $PSScriptRoot/public/objects.ps1
. $PSScriptRoot/public/columns.ps1
. $PSScriptRoot/public/containers.ps1
. $PSScriptRoot/public/rows.ps1
. $PSScriptRoot/public/sheets.ps1


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

# End Setup Functions

#Export Functions
function Export-SmartSheet() {
    [CmdletBinding(DefaultParameterSetName = "none")]
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
        [Parameter(ParameterSetName = 'exists', Mandatory = $true)]
        [string]$overwriteAction,
        [Parameter(ParameterSetName = 'exists', Mandatory = $true)]
        [string]$overwriteSheetId
    )
    
    Begin {
        $Headers = Get-Headers -ContentType:text/csv -ContentDisposition:attachment
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
            }
            else {
                #get a folder off the root
                $rootfolders = Get-SmartsheetHomeFolders
                $currentFolder = $folders.Where({ $_.name -eq $Folder })
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
        if (-not $response.message -eq "SUCCESS") {
            Throw "Import failed! $($_.Exception.Message)"
        }
        else {
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
        
        .PARAMETER overwriteAction
        What to do of the sheet name already exists. Choices are Replace or Rename.
        Multiple sheets can have the same name in a folder. If you omit this parameter a sheet with the same name will be created.

        .PARAMETER overwriteSheetId
        Because tou can have multiple sheets with the same name (with different sheet Ids) you must provide the sheet Id for the overwrite action.

        .EXAMPLE
        PS> $ObjectArray | Export-Smartsheet -SheetName "MyNewSheet"

        .EXAMPLE
        PS> $objectArray | Export-Smartsheet -SheetName "MyNewSheet" -folder 'myfolder1/myfolder2'
    #>
}

# End Export Functions

#Helper functions
function New-SmartSheetFormatString() {
    Param(
        [ValidateSet("Arial", "Tahoma", "Times New Roman", "Verdana")]
        [string]$fontFamily,
        [int]$fontSize,
        [switch]$bold,
        [switch]$italic,
        [switch]$underline,
        [switch]$stikethrew,
        [ValidateSet("left", "center", "right")]
        [string]$horizontalAlign,
        [ValidateSet("top", "middle", "bottom")]
        [string]$verticalAlign,
        [ValidateSet("Navajo White", "Black", "White", "White_2", "Lavender blush", "Sazerac", "Chilean Heath", "Panache", "Solitude", "French Lilac", "Merino", "Pastel Pink", "Navajo White_2", "Dolly", "Zanah", "French Pass", "French Lilac_2", "Almond", "Mercury", "Froly", "Chardonnay", "Yellow", "De York", "Malibu", "Light Wisteria", "Tan", "Silver", "Cinnabar", "Pizazz", "Turbo", "Chateau Green", "Denim", "Seance", "Brown", "Sonic Silver", "Tamarillo", "Trinidad", "Corn", "Forest Green", "Catalina Blue", "Purple", "Carnaby Tan")]
        [string]$textColor,
        [ValidateSet("Navajo White", "Black", "White", "White_2", "Lavender blush", "Sazerac", "Chilean Heath", "Panache", "Solitude", "French Lilac", "Merino", "Pastel Pink", "Navajo White_2", "Dolly", "Zanah", "French Pass", "French Lilac_2", "Almond", "Mercury", "Froly", "Chardonnay", "Yellow", "De York", "Malibu", "Light Wisteria", "Tan", "Silver", "Cinnabar", "Pizazz", "Turbo", "Chateau Green", "Denim", "Seance", "Brown", "Sonic Silver", "Tamarillo", "Trinidad", "Corn", "Forest Green", "Catalina Blue", "Purple", "Carnaby Tan")]
        [string]$backgroundColor,
        [ValidateSet("Navajo White", "Black", "White", "White_2", "Lavender blush", "Sazerac", "Chilean Heath", "Panache", "Solitude", "French Lilac", "Merino", "Pastel Pink", "Navajo White_2", "Dolly", "Zanah", "French Pass", "French Lilac_2", "Almond", "Mercury", "Froly", "Chardonnay", "Yellow", "De York", "Malibu", "Light Wisteria", "Tan", "Silver", "Cinnabar", "Pizazz", "Turbo", "Chateau Green", "Denim", "Seance", "Brown", "Sonic Silver", "Tamarillo", "Trinidad", "Corn", "Forest Green", "Catalina Blue", "Purple", "Carnaby Tan")]
        [string]$taskbarColor,
        [ValidateSet("RUB", "SEK", "AUD", "CAD", "KRW", "USD", "ARS", "NZD", "NOK", "INR", "EUR", "ILS", "SGD", "CNY", "none", "DKK", "BRL", "GBP", "HKD", "JPY", "CLP", "MXN", "CHF", "ZAR")]
        [string]$currency,
        [int]$decimalCount,
        [switch]$thousandsSeparator,
        [ValidateSet("none", "NUMBER", "CURRENCY", "PERCENT")]
        [string]$numberFormat,
        [switch]$textWrap,
        [ValidateSet("LOCALE_BASED", "MMMM_D_YYYY", "MMM_D_YYYY", "D_MMM_YYYY", "YYYY_MM_DD_HYPHEN", "YYYY_MM_DD_DOT", "DWWWW_MMMM_D_YYYY", "DWWW_DD_MMM_YYYY", "DWWW_MM_DD_YYYY", "MMMM_D", "D_MMMM")]
        [string]$dateFormat
    )
    $format = [ordered]@{
        fontFamily         = $null
        fontSize           = $null
        bold               = $null
        italic             = $null
        underlined         = $null
        strikethrough      = $null
        horizontalAlign    = $null
        verticalAlign      = $null
        textcolor          = $null
        backgroundColor    = $null
        taskbarColor       = $null
        currency           = 13
        decimalCount       = 2
        thousandsSeparator = $null
        numberFormat       = $null
        textWrap           = $null
        dateFormat         = $null
    }
    if ($fontFamily) { $format.fontFamily = $smformat.fontFamilies[$fontFamily] }
    if ($fontSize) { $format.fontSize = $fontSize }
    if ($bold) { $format.bold = 1 }
    if ($italic) { $format.italic = 1 }
    if ($underline) { $format.underlined = 1 }
    if ($stikethrew) { $format.strikethrough = 1 }
    if ($horizontalAlign) { $format.horizontalAlign = $smformat.horizontalAlign[$horizontalAlign] }
    if ($verticalAlign) { $format.verticalAlign = $smformat.verticalAlign[$verticalAlign] }
    if ($textColor) { $format.textcolor = $smformat.colors[$textColor] }
    if ($backgroundColor) { $format.backgroundColor = $smformat.colors[$backgroundColor] }
    if ($taskbarColor) { $format.backgroundColor = $smformat.colors[$backgroundColor] }
    if ($currency) { $format.currency = $smformat.currency[$currency] }
    if ($decimalCount) { $format.decimalCount = $decimalCount }
    if ($thousandsSeparator) { $format.thousandsSeparator = 1 }
    if ($numberFormat) { $format.numberFormat = $smformat.numberFormats[$numberFormat] }
    if ($textWrap) { $format.textWrap = 1 }
    if ($dateFormat) { $format.dateFormat = $smformat.dateFormats[$dateFormat] }
    
    $formatString = $format.values -join ","
    return $formatString
}

# End Helper functions
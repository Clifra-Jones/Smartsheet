#
# Module manifest for module 'Smartsheet'
#
# Generated by: Cliff Williams
#
# Generated on: 6/30/2022
#

@{

# Script module or binary module file associated with this manifest.
RootModule = '.\Smartsheet.psm1'

# Version number of this module.
ModuleVersion = '1.0.0'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = 'c770341c-98dc-4040-b936-d7b68bd87f92'

# Author of this module
Author = 'Cliff Williams'

# Company or vendor of this module
CompanyName = 'Balfour Beatty US'

# Copyright statement for this module
Copyright = '(c) Cliff Williams. All rights reserved.'

# Description of the functionality provided by this module
# Description = ''

# Minimum version of the PowerShell engine required by this module
# PowerShellVersion = ''

# Name of the PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# ClrVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
ScriptsToProcess = @('./public/public.ps1')

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
<# NestedModules = @(
    "./public/objects.psm1",
    "./public/sheets.psm1",
    "./public/rows.psm1",
    "./public/columns.psm1",
    "./public/containers.psm1"
) #>

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @(
    'Set-APIKey',
    'Get-ServerInfo',
    'Get-Smartsheets',
    'Get-Smartsheet',
    'Get-SortedSmartsheet',
    'Remove-Smartsheet',
    "Copy-Smartsheet",
    "Rename-SmartSheet",
    "Move-SmartSheet",
    'Get-SmartsheetColumn',
    'Set-SmartsheetColumn',
    'Get-SmartsheetColumns',
    'Add-SmartsheetColumn',
    'Add-SmartsheetColumns',
    'Remove-SmartsheetColumn',
    'Add-SmartsheetRow',
    'Add-SmartsheetRows',
    'Remove-SmartsheetRow',
    'Remove-SmartsheetRows',
    'Set-SmartsheetRow',
    'Set-SmartsheetRows',
    'Get-SmartsheetRow',
    'Get-SmartsheetFolder',
    'Remove-SmartsheetFolder'
    'Get-SmartsheetFolders',
    'Add-SmartsheetFolder',
    'Get-SmartsheetHome',
    'Get-SmartsheetHomeFolders',
    'Add-SmartsheetHomeFolder',
    'Export-SmartSheet',
    'New-SmartsheetCell',
    'New-HyperLink',
    'New-CellLink',
    'New-SmartSheetFormatString',
    'Export-SmartsheetRows',
    'Send-SmartsheetViaEmail',
    'Add-SmartsheetShare',
    'Get-SmartsheetShares',
    'Get-SmartSheetShare',
    'Remove-SmartsheetShare',
    'Set-SmartsheetShare',
    'Get-SmartsheetAttachments',
    'Add-SmartsheetAttachment',
    'Get-SmartSheetAttachment',
    'Remove-SmartSheetAttachment',
    'Copy-SmartsheetAttachments',
    'Copy-SmartsheetShares',
    'Get-SmartsheetDiscussions',
    'New-SmartsheetDiscussion',
    'Remove-SmartsheetDiscussion',
    'Get-SmartsheetRowDiscussions',
    'New-SMartsheetRowDiscussion'
)

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = '*'

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        # ProjectUri = ''

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

        # Prerelease string of this module
        # Prerelease = ''

        # Flag to indicate whether the module requires explicit user acceptance for install/update/save
        # RequireLicenseAcceptance = $false

        # External dependent modules of this module
        # ExternalModuleDependencies = @()

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}


Import-Module $PSScriptRoot/../Smartsheet.psd1

. $home/OneDrive/Scripts/psDoc/src/psDoc.ps1 -moduleName Smartsheet -Style HTML-Types -outputDir $PSScriptRoot -fileName "reference.html"

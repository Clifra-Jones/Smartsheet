#Requires -Modules @{ModuleName="Smartsheet"; ModuleVersion="1.0.0"}

# This script that generates the documentation is not part of this module.
# It was originally created by Chase Florell, the original can be downloaded from GitHub at https://github.com/ChaseFlorell/psDoc
# This is a modified version of the script that adds type definitions to the documentation. 
# You can download this modified version from my github page at https://github.com/Clifra-Jones/psDoc

# If you are using GitHub you can create githib pages from your readme.md or other markdown documents. Then you can link to the generated reference.html
# with /docs/reference.html

# To generate the documentation, execute this script from the source root folder.
# ./docs/makeDocs.ps1
#

# modify the path to the psDoc.ps1 file in the line below.
. $home/OneDrive/Scripts/psDoc/src/psDoc.ps1 -moduleName Smartsheet -Style HTML-Types -outputDir $PSScriptRoot -fileName "reference.html"

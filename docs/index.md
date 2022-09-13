# Powershell Module for the Smartsheet API

This module allows you to interact with the Smartsheet REST API using powershell.  

This module is 100% powershell and does not use the c# SDK available from smartsheet.
The reason is to not require including any additional .dll files to use this module. All calls to the API are done through the Invoke-RestMethod function.  

## Functionality

This module includes most of the functions to interact with Sheets. Functions for Dashboards, Reports, Workspaces, Users and Groups may be added at a later date.

## Authentication

OAuth authentication has not been implemented at this time. You must use an Access token generated from your Smartsheet account.
To generate an API Access Token click on the Account Icon at the bottom of the left sidebar. Select Personal Settings, select API Access and then Generate
an Access Token. Save the Access token in a safe place (a folder ONLY you have access to) you will not be able to retrieve the token after this. If you lose your token you will need to generate a new one.

## Developer Account

This module is still in an early stage of development. Even if the functions are working properly you can alter Smartsheets with unintended concequences.
Do not modify production Smartsheets unless you fully understand what you are doing.

It is a best practice to create a developers account and test your processes there before working on your production sheets.

To create a developer account go to [Register as a Developer](https://developers.smartsheet.com/register) and create an account using a different email address than you use with your production account. Here you will have access to all the Developer tools and can create and modify Smartsheets.
The account is limited to 2 users.

### [Module Referrence](./referrence.html)

The above link is a full module referrence that includes syntax, parameters and examples.

## Installation

To install the module, clone the repository into your module folder.

### User Scope

Change to your user module directory.

For **Windows**:

```dos
cd %USERPROFILE%\Documents\Powershell\Modules
```

For **Linux/Mac**:

```bash
>cd ~/.local/share/powershell/Modules
```

Clone the repository.

```bash
>git clone https://github.com/Clifra-Jones/Smartsheet.git
```

### System Scope

Change to the system module directory.

for **Windows**:

```bash
>cd %PROGRAMFILES%\PowerShell\Modules
```

For **Linux/Mac**:

```bash
>cd /usr/local/share/powershell/MOdules
```

Clone the repository.

```bash
>git clone https://github.com/Clifra-Jones/Smartsheet.git
```

## Usage

The primary usage for htis module is to create or consume Smartsheets within powershell.

**Warning**
: This module can be very dangerous and you can cause serious damage to a production Smartsheet if you are not careful and do not fully understand what you are are doing. See the section above about creating a Developers account to test your processes.

### Get a Smartsheet as an array of powershell object

to retrieve a Smartsheet and convert the data into an array of powershell objects use the [**Get-Smartsheet](https://clifra-jones.github.io/Smartsheet/referrence.html#Copy-Smartsheet) function.

Use the ToArray function to return an array of Powershell objects from the Sheet object.

```powershell
$Sheet = Get-Smartsheet -Name "MySmartsheet"
$array = $sheet.ToArray()
```

### Export Powershell array of objects to a Smartsheet

To export an array of powershell object into a Smartsheet you use the [**Export-Smartsheet**](https://clifra-jones.github.io/Smartsheet/referrence.html#Export-SmartSheet) function.
This funtion will **ALWAYS** create a new Smartsheet, even if a sheet of the same name exist in the target folder.
Smartsheets are uniquely identified by the Smartsheet's ID, not the name.

```powershell
$MyArray | Export-Smartsheet -sheetName "MyNewSheet"
```

If you want to overwrite an existing sheet you must retrieve its Id and supply that using the -overwriteSheetId parameter and also provider the -overwriteAction parameter with the value 'Replace'.

```powershell
>$oldSheet = Get-Smartsheet -Name "MySheet"
```

This assumes there is only 1 sheet named "MySheet" in the home folder.

Then create the new sheet overwriting the old sheet.

```powershell
$MyArray | Export-Smartsheet -sheetName "MySheet" -overwriteSheetId $oldsheet.id -overwriteAction Replace
```

**Note**
: This action creates a second sheet named "MySheet", copies the shares, discussions and comments from the old sheet and then deletes the old sheet. The old sheet can be recovered from the Deleted Items container on the Smartsheet web site.

You can also rename a smartsheet with this function by providing the -overwriteSheetID and the -overwriteAction with the 'Rename' value.

### Export a Powershell array as a set of Smartsheet rows into an existing Smartsheet

You can append/insert a powershell array into a Smartsheet using the [**Export-SmartsheetRows**](https://clifra-jones.github.io/Smartsheet/referrence.html#Export-SmartsheetRows) function.

This function is generally used to create the equivelent of an Excel table in a Smartsheet. This is sort of "out of functionality" for how Smartsheets works, but some may find it Useful. You can also use this finction to append rows to an existing Smartsheet.

The following example imports the array into a smnartsheet, creates a blank row above the data and adds a title and a header row.
(To create the format veriables use New-SmartsheetFormatString)

```powershell
> $Array | Export-SmartsheetRows -blankRowAbove -title "My Title" -TitleFormat $titleFormat -includeHeaders -headerFormat $headerFormat
```

The following example exports the array into a smartsheet appending the rows to the existing sheet without any title or headers.
This can be used to append rows to the Smartsheet. No attempt is made to prevent duplicate data.
If the number of properties in the objects is more than the existing columns, then generic columns are created.
(To update rows based in their primary column values use the [**Update-Smartsheet**](https://clifra-jones.github.io/Smartsheet/referrence.html#Update-Smartsheet) function.)

```powershell
$Array | Export-SmartsheetRows
```

### Update the rows in a smartsheet based on thier primary column value.

To update rows in a Smartsheet based on their primary column value use the [**Update-Smartsheet**](https://clifra-jones.github.io/Smartsheet/referrence.html#Update-Smartsheet) function.

This function makes the following assumptions:

1. The number and names of the columns are the same as the properties in the object in the array.
2. The primary column is the first column in the sheet and the column values are unique.

If condition 1 isn't met, an error will be thrown.
if Condition 2 isn't met, unpredictable results may occur.

```powershell
$MyArray | Update-Smartsheeet -SheetId MySheet.Id
```

### Add a new column to a SmartSheet

To add a new column to a Smartsheet use the [**Add-SmartsheetColumn**](https://clifra-jones.github.io/Smartsheet/referrence.html#Add-SmartsheetColumn) function.

The following example adds a new column to the end of the columns. Then updates the existing sheet object.

```powershell
>$Sheet = $Sheet | Add-SmartsheetColumn -Title "MyNewColumn" -Type TEXT_NUMBER -Passthru
```

To insert a column at a certain position use the -index parameter. The column will be inserted at that position shifting all columns after that to the right.

```powershell
>$Sheet = $Sheet | Add-SmartsheetColumn -title "MyNewColumn" -Type TEXT_NUMBER -index 3 -PassThru
```

### Add a new share to a Smartsheet

A share allows you to grant access to a Smartsheet to a user in your organization.
The following example grants the user with email johndoe@example.com EDITOR access to the smartsheet and emails him informing that the sheet has been shared with him. (Assumes we already have a sheet object in $Sheet)

```powershell
>$Sheet | Add-SmartsheetShare -AccessLevel EDITOR -SendEmail -Email johndoe@example.com -message "This is the employee data we discussed"
```

There are many more function that can add/remove/update sheets, rows, and columns. Manage Attachments, discussions, and commments. And add remove folders and more.

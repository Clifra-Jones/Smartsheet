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
>cd %USERPROFILE%\Documents\Powershell\Modules

For **Linux/Mac**:
>cd ~/.local/share/powershell/Modules

Clone the repository.
>git clone https://github.com/Clifra-Jones/Smartsheet.git

### System Scope

Change to the system module directory.

for **Windows**:
>cd %PROGRAMFILES%\PowerShell\Modules

For **Linux/Mac**:
>cd /usr/local/share/powershell/MOdules

Clone the repository.
>git clone https://github.com/Clifra-Jones/Smartsheet.git

<#
.SYNOPSIS
Manage-PFPermission.ps1

.DESCRIPTION 
This script is used to Manage Public Folder client level permission.
You could use this script with ADD/REMOVE/VIEWPERMISSION parameter or 
with PROMPT parameter which would show text input prompt messages to run this script.

.OUTPUTS
Display information about a given Public Folder and output to console.

.PARAMETER PublicFolderPath
The PublicFolderPath parameter is used to specify Public Folder name including full path under 'All Public Folders'.

.PARAMETER User
The User parameter is used to specify user email address or alias as user/security_group identity.

.PARAMETER AddPermission
The AddPermission parameter is used as switch to apply permission on specified Public Folder to specified User/security_group identity.

.PARAMETER RemovePermission
The RemovePermission parameter is used as switch to remove permission on specified Public Folder for specified User/security_group identity.

.PARAMETER ViewPermission
The ViewPermission parameter is used as switch to view permission on specified Public Folder.

.PARAMETER IncludeSubFolders
The IncludeSubFolders parameter is used if action need to be performed on all sub-folders.

.PARAMETER Prompt
The PROMPT parameter could be used to display multiple text input prompt messages to manage Public Folder permission related tasks.

.PARAMETER Log
Used to generate a log file for the script.

.EXAMPLE
Get-Help .\Manage-PFPermission.ps1 -Full
Returns full details about script usage method including parameters, syntax and examples.

.EXAMPLE
.\Manage-PFPermission.ps1 -Prompt
Returns multiple text input prompt messages which could be helpful for helpdesk support staff.

.EXAMPLE
.\Manage-PFPermission.ps1 -ViewPermission -PublicFolderPath "\Folder\SubFolder\SubFolder"
Returns information about Public Folder permission on specified Public Folder.

.EXAMPLE
.\Manage-PFPermission.ps1 -ViewPermission -PublicFolderPath "\Folder\SubFolder\SubFolder" -IncludeSubFolders
Returns information about Public Folder permission on specified Public Folder includin all sub-folders.

.EXAMPLE
.\Manage-PFPermission.ps1 -AddPermission -PublicFolderPath "\Folder\SubFolder\SubFolder" -User SAM999
Apply opted permission to specified User/security_group identity on specified Public Folder.

.EXAMPLE
.\Manage-PFPermission.ps1 -AddPermission -PublicFolderPath "\Folder\SubFolder\SubFolder" -User SAM999 -IncludeSubFolders
Apply opted permission to specified User/security_group identity on specified Public Folder includin all sub-folders.

.EXAMPLE
.\Manage-PFPermission.ps1 -RemovePermission -PublicFolderPath "\Folder\SubFolder\SubFolder" -User SAM999
Remove currently applied permission for specified User/security_group identity on specified Public Folder.

.EXAMPLE
.\Manage-PFPermission.ps1 -RemovePermission -PublicFolderPath "\Folder\SubFolder\SubFolder" -User SAM999 -IncludeSubFolders
Remove currently applied permission for specified User/security_group identity on specified Public Folder includin all sub-folders.

.EXAMPLE
.\Manage-PFPermission.ps1 -Log
Returns a report with a log file at the script file path when used with other parameters.

.LINK
https://

.NOTES
Written by: Sumit Dhuper

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Find me on:

* My Blog:	http://
* Twitter:	https://twitter.com/sumitdhuper
* LinkedIn:	https://www.linkedin.com/in/sumitdhuper
* Github:	https://github.com/sumit0009

Change Log:
V1.00, 21/07/2015 - Initial version
V1.01, 26/08/2015 - Added PublishingEditor permission option and did few minor changes in message prompt function
V1.02, 27/08/2015 - Added Parameters functionality in this script while keeping text input prompt method under PROMPT function
V1.03, 28/08/2015 - Updated script with help details
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Position=0,
        Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
	[string]$PublicFolderPath,

	[Parameter( Mandatory=$false)]
	[string]$User,

	[Parameter( Mandatory=$false)]
	[switch]$AddPermission,

	[Parameter( Mandatory=$false)]
	[switch]$RemovePermission,

	[Parameter( Mandatory=$false)]
	[switch]$ViewPermission,

	[Parameter( Mandatory=$false)]
	[switch]$IncludeSubFolders,

	[Parameter( Mandatory=$false)]
	[switch]$Prompt,

	[Parameter( Mandatory=$false)]
	[switch]$Log
	)


#...................................
# Variables
#...................................

$scriptname = $MyInvocation.MyCommand.Name
$now = Get-Date

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logfile = "$myDir\$scriptname.log"


#...................................
# Logfile Strings
#...................................

$logstring0 = "====================================="
$logstring1 = " $scriptname"


#...................................
# Initialization Strings
#...................................

$initstring0 = "Initializing..."
$initstring1 = "Loading the Exchange Server PowerShell snapin..."

#Try Exchange 2007 snapin first

$2007snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -Registered
if ($2007snapin)
{
    if (!(Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue))
    {
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
	}

#Set scope to include entire forest
$scope = "View Entire Forest"
if (!$AdminSessionADSettings.ViewEntireForest)
{
$scope = $AdminSessionADSettings.DefaultScope
}
}

else
{
    #Add Exchange 2010 snapin if not already loaded in the PowerShell session
    if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
    {
	    . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	    Connect-ExchangeServer -auto -AllowClobber
    }
    else
    {
        Write-Warning "Exchange Server management tools are not installed on this computer."
        EXIT
    }

    Set-ADServerSettings -ViewEntireForest $true
}

#...................................
# Functions
#...................................

#This function is used to write the log file if -Log is used
Function Write-Logfile()
{
	param( $logentry )
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logentry" | Out-File $logfile -Append
}


#...................................
# Initialize
#...................................

#Log file is overwritten each time the script is run to avoid
#very large log files from growing over time
if ($Log) {
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0
}

Write-Verbose $initstring0
if ($Log) {Write-Logfile $initstring0}

Write-Verbose $initstring1
if ($Log) {Write-Logfile $initstring1}


#...................................
# Script
#...................................

If($ViewPermission)
{
if ($IncludeSubfolders)
{
Get-PublicFolder -Identity "$PublicFolderPath" –Recurse | Get-PublicFolderClientPermission | ft @{n = "FolderPath"; e=

{$_.Identity}},user,access* -Wrap
}

Else
{
Get-PublicFolder -Identity "$PublicFolderPath" | Get-PublicFolderClientPermission | ft @{n = "FolderPath"; e={$_.Identity}},user,access* -Wrap
}
}

If($AddPermission)
{
if ($IncludeSubfolders)
{
Write-host "Below are the available options to apply permission on above mentioned Public Folder." -Fore DarkYellow
Write-host "1. Author" -Fore DarkYellow
Write-host "2. Reviewer" -Fore DarkYellow
Write-host "3. PublishingAuthor" -Fore DarkYellow
Write-host "4. PublishingEditor" -Fore DarkYellow
Write-host "5. Editor" -Fore DarkYellow
Write-host "6. Owner" -Fore DarkYellow
$AR = Read-host "Please enter an option 1 to 6..."

Switch( $AR ){
  1{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Add-PublicFolderClientPermission –User "$User" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Add-PublicFolderClientPermission –User "$User" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Add-PublicFolderClientPermission –User "$User" –AccessRights "PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Add-PublicFolderClientPermission –User "$User" –AccessRights "PublishingEditor"}
  5{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Add-PublicFolderClientPermission –User "$User" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Add-PublicFolderClientPermission –User "$User" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }
}
}

Else
{
Write-host "Below are the available options to apply permission on above mentioned Public Folder." -Fore DarkYellow
Write-host "1. Author" -Fore DarkYellow
Write-host "2. Reviewer" -Fore DarkYellow
Write-host "3. PublishingAuthor" -Fore DarkYellow
Write-host "4. PublishingEditor" -Fore DarkYellow
Write-host "5. Editor" -Fore DarkYellow
Write-host "6. Owner" -Fore DarkYellow
$AR = Read-host "Please enter an option 1 to 6..."

Switch( $AR ){
  1{Get-PublicFolder –Identity “$PublicFolderPath” | Add-PublicFolderClientPermission –User "$User" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PublicFolderPath” | Add-PublicFolderClientPermission –User "$User" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PublicFolderPath” | Add-PublicFolderClientPermission –User "$User" –AccessRights "PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PublicFolderPath” | Add-PublicFolderClientPermission –User "$User" –AccessRights "PublishingEditor"}
  5{Get-PublicFolder –Identity “$PublicFolderPath” | Add-PublicFolderClientPermission –User "$User" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PublicFolderPath” | Add-PublicFolderClientPermission –User "$User" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }
}
Write-host "Permission has been successfully applied!!!" -Fore Green
}
}

If($RemovePermission)
{
if ($IncludeSubfolders)
{
Write-host "Below are the available options to remove permission on above mentioned Public Folder." -Fore DarkYellow
Write-host "1. Author" -Fore DarkYellow
Write-host "2. Reviewer" -Fore DarkYellow
Write-host "3. PublishingAuthor" -Fore DarkYellow
Write-host "4. PublishingEditor" -Fore DarkYellow
Write-host "5. Editor" -Fore DarkYellow
Write-host "6. Owner" -Fore DarkYellow
$AR = Read-host "Please enter an option 1 to 6..."

Switch( $AR ){
  1{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Remove-PublicFolderClientPermission –User "$User" –AccessRights 

"PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Remove-PublicFolderClientPermission –User "$User" –AccessRights 

"PublishingEditor"}
  5{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PublicFolderPath” –Recurse | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }
}
}

Else
{
Write-host "Below are the available options to remove permission on above mentioned Public Folder." -Fore DarkYellow
Write-host "1. Author" -Fore DarkYellow
Write-host "2. Reviewer" -Fore DarkYellow
Write-host "3. PublishingAuthor" -Fore DarkYellow
Write-host "4. PublishingEditor" -Fore DarkYellow
Write-host "5. Editor" -Fore DarkYellow
Write-host "6. Owner" -Fore DarkYellow
$AR = Read-host "Please enter an option 1 to 6..."

Switch( $AR ){
  1{Get-PublicFolder –Identity “$PublicFolderPath” | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PublicFolderPath” | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PublicFolderPath” | Remove-PublicFolderClientPermission –User "$User" –AccessRights "PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PublicFolderPath” | Remove-PublicFolderClientPermission –User "$User" –AccessRights "PublishingEditor"}
  5{Get-PublicFolder –Identity “$PublicFolderPath” | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PublicFolderPath” | Remove-PublicFolderClientPermission –User "$User" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }
}
Write-host "Permission has been successfully removed!!!" -Fore Green
}
}

If($Prompt)
{

Function Add
{
Write-Host "         This Script is used to Add Public Folder Client Permission." -Fore Cyan
Write-Host "       ***************************************************************" -Fore Yellow
Write-Host ""

$Random = Get-Random -Maximum 1000 -Minimum 101
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; 

Write-Host " to continue on this Script."
$Code = Read-Host "Enter the verification Code as showing above"

if ($Random -eq "$Code")
{Write-Host "Code has been verified successfully!!!" -ForegroundColor Green}

Else
    {Write-Host "ERROR: Verification Failed!!! Kindly run the Script again." -ForegroundColor Yellow -BackgroundColor Black
    Return
    End}

$PF = Read-Host "Please enter the Public Folder path under All Public Folders including preceding '\' sign"
Write-host
$U = Read-Host "Please enter the user Email Address or Alias here"
Write-host
Write-Host "Kindly confirm if you would like to apply permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -

NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
$subfolder = Read-Host " "

if ($subfolder -eq "$null")
{Write-Host "ERROR: Permission scope has not been confirmed!!!" -BackgroundColor black -ForegroundColor Yellow
  Return
  End
  }

Write-host "Below are the available options to apply permission on above mentioned Public Folder." -Fore DarkYellow
Write-host "1. Author" -Fore DarkYellow
Write-host "2. Reviewer" -Fore DarkYellow
Write-host "3. PublishingAuthor" -Fore DarkYellow
Write-host "4. PublishingEditor" -Fore DarkYellow
Write-host "5. Editor" -Fore DarkYellow
Write-host "6. Owner" -Fore DarkYellow
$AR = Read-host "Please enter an option 1 to 6..."


if ($subfolder -eq "Y")
{
Switch( $AR ){
  1{Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User "$U" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User "$U" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User "$U" –AccessRights "PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User "$U" –AccessRights "PublishingEditor"}
  5{Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User "$U" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User "$U" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }

}
}

if ($subfolder -eq "N")
{
Switch( $AR ){
  1{Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User "$U" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User "$U" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User "$U" –AccessRights "PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User "$U" –AccessRights "PublishingEditor"}
  5{Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User "$U" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User "$U" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }

}
}

Write-host "Permission has been successfully applied!!!" -Fore Green

}


Function Remove
{
Write-Host "         This Script is used to Remove Public Folder Client Permission." -Fore Cyan
Write-Host "       ******************************************************************" -Fore Yellow
Write-Host ""

$Random = Get-Random -Maximum 1000 -Minimum 101
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; 

Write-Host " to continue on this Script."
$Code = Read-Host "Enter the verification Code as showing above"

if ($Random -eq "$Code")
{Write-Host "Code has been verified successfully!!!" -ForegroundColor Green}

Else
    {Write-Host "ERROR: Verification Failed!!! Kindly run the Script again." -ForegroundColor Yellow -BackgroundColor Black
    Return
    End}

$PF = Read-Host "Please enter the Public Folder path under All Public Folders including preceding '\' sign"
Write-host
$U = Read-Host "Please enter the user Email Address or Alias here"
Write-host
Write-Host "Kindly confirm if you would like to remove permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -

NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
$subfolder = Read-Host " "

if ($subfolder -eq "$null")
{Write-Host "ERROR: Permission scope has not been confirmed!!!" -BackgroundColor black -ForegroundColor Yellow
  Return
  End
  }

Write-host
Write-host "Below are the available options to remove permission on above mentioned Public Folder." -Fore DarkYellow
Write-host "1. Author" -Fore DarkYellow
Write-host "2. Reviewer" -Fore DarkYellow
Write-host "3. PublishingAuthor" -Fore DarkYellow
Write-host "4. PublishingEditor" -Fore DarkYellow
Write-host "5. Editor" -Fore DarkYellow
Write-host "6. Owner" -Fore DarkYellow
$AR = Read-host "Please enter an option 1 to 6..."

if ($subfolder -eq "Y")
{
Switch( $AR ){
  1{Get-PublicFolder –Identity “$PF” –Recurse | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PF” –Recurse | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PF” –Recurse | Remove-PublicFolderClientPermission –User "$U" –AccessRights "PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PF” –Recurse | Remove-PublicFolderClientPermission –User "$U" –AccessRights "PublishingEditor"}
  5{Get-PublicFolder –Identity “$PF” –Recurse | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PF” –Recurse | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }
}
}

if ($subfolder -eq "N")
{
Switch( $AR ){
  1{Get-PublicFolder –Identity “$PF” | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Author"}
  2{Get-PublicFolder –Identity “$PF” | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Reviewer"}
  3{Get-PublicFolder –Identity “$PF” | Remove-PublicFolderClientPermission –User "$U" –AccessRights "PublishingAuthor"}
  4{Get-PublicFolder –Identity “$PF” | Remove-PublicFolderClientPermission –User "$U" –AccessRights "PublishingEditor"}
  5{Get-PublicFolder –Identity “$PF” | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Editor"}
  6{Get-PublicFolder –Identity “$PF” | Remove-PublicFolderClientPermission –User "$U" –AccessRights "Owner"}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
  }
}
}

Write-host "Permission has been successfully removed!!!" -Fore Green

}

Function View
{
Write-Host "         This Script is used to View Public Folder Client Permission." -Fore Cyan
Write-Host "       *****************************************************************" -Fore Yellow
Write-Host ""

$Random = Get-Random -Maximum 1000 -Minimum 101
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; 

Write-Host " to continue on this Script."
$Code = Read-Host "Enter the verification Code as showing above"

if ($Random -eq "$Code")
{Write-Host "Code has been verified successfully!!!" -ForegroundColor Green}

Else
    {Write-Host "ERROR: Verification Failed!!! Kindly run the Script again." -ForegroundColor Yellow -BackgroundColor Black
    Return
    End}

$PF = Read-Host "Please enter the Public Folder path under All Public Folders including preceding '\' sign"
Write-host
Write-Host "Kindly confirm if you would like to view permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -

NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
$subfolder = Read-Host " "

if ($subfolder -eq "$null")
{Write-Host "ERROR: Permission scope has not been confirmed!!!" -BackgroundColor black -ForegroundColor Yellow
  Return
  End
  }

if ($subfolder -eq "Y")
{
Get-PublicFolder -Identity "$PF" –Recurse | Get-PublicFolderClientPermission | ft @{n = "FolderPath"; e={$_.Identity}},user,access* -Wrap
}

if ($subfolder -eq "N")
{
Get-PublicFolder -Identity "$PF" | Get-PublicFolderClientPermission | ft @{n = "FolderPath"; e={$_.Identity}},user,access* -Wrap
}
}


Function AddBulk
{
Write-Host "         This Script is used to Add Public Folder Client Permission - BULK REQUEST." -Fore Cyan
Write-Host "       ******************************************************************************" -Fore Yellow
Write-Host ""

$Random = Get-Random -Maximum 1000 -Minimum 101
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; 

Write-Host " to continue on this Script."
$Code = Read-Host "Enter the verification Code as showing above"

if ($Random -eq "$Code")
{Write-Host "Code has been verified successfully!!!" -ForegroundColor Green}

Else
    {Write-Host "ERROR: Verification Failed!!! Kindly run the Script again." -ForegroundColor Yellow -BackgroundColor Black
    Return
    End}

$CSVFile = Test-Path .\sam.csv
if ($CSVFile -eq "True")
{
$csv = Import-Csv ".\sam.csv"
$count = $csv.count
Write-Host "NOTE: CSV file sam.csv found under the Script folder and there are $count users SamAccountName/Alias found in the file!!!" -

BackgroundColor Cyan -ForegroundColor DarkBlue
}

Else
{
Write-Host "ERROR: Kindly make sure CSV file 'sam.csv' is already present in the Script folder" -BackgroundColor Cyan -ForegroundColor Red
Write-Host "and updated with all users SamAccountName/Alias (and column heading MUST be 'SAM')!!!" -BackgroundColor Cyan -ForegroundColor Red
Return
End
}

$PF = Read-Host "Please enter the Public Folder path under All Public Folders including preceding '\' sign"
Write-host
Write-Host "Kindly confirm if you would like to apply permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -

NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
$subfolder = Read-Host " "

if ($subfolder -eq "$null")
{Write-Host "ERROR: Permission scope has not been confirmed!!!" -BackgroundColor black -ForegroundColor Yellow
  Return
  End
  }


Write-host
Write-host "Below are the available options to apply permission on above mentioned Public Folder." -Fore DarkYellow
Write-host "1. Author" -Fore DarkYellow
Write-host "2. Reviewer" -Fore DarkYellow
Write-host "3. PublishingAuthor" -Fore DarkYellow
Write-host "4. PublishingEditor" -Fore DarkYellow
Write-host "5. Editor" -Fore DarkYellow
Write-host "6. Owner" -Fore DarkYellow
$AR = Read-host "Please enter an option 1 to 6..."


if ($subfolder -eq "Y")
{
Switch( $AR ){
  1{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"Author"}}
  2{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"Reviewer"}}
  3{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"PublishingAuthor"}}
  4{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"PublishingEditor"}}
  5{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"Editor"}}
  6{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"Owner"}}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
}
}}

if ($subfolder -eq "N")
{
Switch( $AR ){
  1{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Author"}}
  2{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Reviewer"}}
  3{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"PublishingAuthor"}}
  4{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights 

"PublishingEditor"}}
  5{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Editor"}}
  6{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Owner"}}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -

ForegroundColor Yellow
  Return
  End
}
}

Write-host "Permission has been successfully applied!!!" -Fore Green
}

}
Write-Host "         This Script is used to Manage Public Folder Client Permission." -Fore Cyan
Write-Host "       ******************************************************************" -Fore Yellow
Write-Host
Write-host "Below are the available options to manage Public Folder Client level permission." -Fore DarkYellow
Write-host "1. Add Permission on Public Folder" -Fore Yellow
Write-host "2. Remove Permission on Public Folder" -Fore Yellow
Write-host "3. View Permission on Public Folder" -Fore Yellow
Write-host "4. Add Permission on Public Folder - BULK REQUEST" -Fore Yellow
$call = Read-host "Please enter an option 1 to 4..."
Write-host

Switch( $call ){
  1{Add}
  2{Remove}
  3{View}
  4{AddBulk}
}
}

#...................................
# End
#...................................

if ($Log) {
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile -Append
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0
}

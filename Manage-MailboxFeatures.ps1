<#
.SYNOPSIS
Manage-MailboxFeatures.ps1

.DESCRIPTION 
This script displays user account details including user Mailbox Feature status on Exchange Server if script run with MAILBOX parameter.
If script run with PROMPT parameter then it will go through user input message prompt function to view and manage mailbox features.

.OUTPUTS
Displays user account details including Mailbox Feature status on Exchange Server.

.PARAMETER Mailbox
The Mailbox parameter could be used to specify mailbox name or email address of the mailbox to get feature status about the provided user name.

.PARAMETER EnableFeature
Used with MAILBOX parameter and would call user input message prompt function to Enable Mailbox Feature on specified mailbox.

.PARAMETER DisableFeature
Used with MAILBOX parameter and would call user input message prompt function to Disable Mailbox Feature on specified mailbox.

.PARAMETER Prompt
The PROMPT parameter could be used to run text input prompt messages to display information about specified mailbox.

.PARAMETER Log
Used to generate a log file for the script.

.EXAMPLE
Get-Help .\Manage-MailboxFeatures.ps1 -Full
Returns full details including Syntax and Examples used for this script.

.EXAMPLE
.\Manage-MailboxFeatures.ps1 -Prompt
Returns multiple text input prompt messages which could be helpful for helpdesk support staff.

.EXAMPLE
.\Manage-MailboxFeatures.ps1 -Mailbox SAM999
Returns Mailbox Feature details about a given mailbox when using Alias.

.EXAMPLE
.\Manage-MailboxFeatures.ps1 -Mailbox sumit.dhuper@example.net
Returns Mailbox Feature details about a given mailbox when using Email Address.

.EXAMPLE
.\Manage-MailboxFeatures.ps1 -Mailbox SAM999 -EnableFeature
Returns text input message prompts to Enable Mailbox Feature on specified mailbox.

.EXAMPLE
.\Manage-MailboxFeatures.ps1 -Mailbox SAM999 -DisableFeature
Returns text input message prompts to Disable Mailbox Feature on specified mailbox.

.EXAMPLE
$Mailboxes = Get-Mailbox -Server MBXSRV01
foreach ($Mailbox in $Mailboxes){.\Manage-MailboxFeatures.ps1 $Mailbox}
Returns information about multiple mailbox when using foreach loop method.

.EXAMPLE
.\Manage-MailboxFeatures.ps1 -Mailbox SAM999 -Log
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
V1.00, 18/08/2015 - Initial version
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( 
        Position=0,
        Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
	[string]$Mailbox,

	[Parameter( Mandatory=$false)]
	[switch]$EnableFeature,

	[Parameter( Mandatory=$false)]
	[switch]$DisableFeature,

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
$initstring2 = "Active Directory Scope has been set to Entire Forest..."
$initstring3 = "Function Enable has been called..."
$initstring4 = "Function Disable has been called..."
$initstring5 = "Mailbox attribute value has been updated on MAPIBlockOutlookRpcHttp to FALSE"
$initstring6 = "Mailbox attribute value has been updated on ActiveSyncEnabled to TRUE"
$initstring7 = "Mailbox attribute value has been updated on OWAEnabled to TRUE"
$initstring8 = "Mailbox attribute value has been updated on PopEnabled to TRUE"
$initstring9 = "Mailbox attribute value has been updated on ImapEnabled to TRUE"
$initstring10 = "Mailbox attribute value has been updated on MapiEnabled to TRUE"
$initstring11 = "Mailbox attribute value has been updated on MAPIBlockOutlookRpcHttp to TRUE"
$initstring12 = "Mailbox attribute value has been updated on ActiveSyncEnabled to FALSE"
$initstring13 = "Mailbox attribute value has been updated on OWAEnabled to FALSE"
$initstring14 = "Mailbox attribute value has been updated on PopEnabled to FALSE"
$initstring15 = "Mailbox attribute value has been updated on ImapEnabled to FALSE"
$initstring16 = "Mailbox attribute value has been updated on MapiEnabled to FALSE"

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

Function Enable
{
Write-Verbose $initstring3
if ($Log) {Write-Logfile $initstring3}
Write-host
Write-host "Below are the available options to Enable Mailbox Feature." -Fore Yellow
Write-host "1. Outlook Anywhere Access" -Fore DarkYellow
Write-host "2. Exchange ActiveSync" -Fore DarkYellow
Write-host "3. Outlook Web Access" -Fore DarkYellow
Write-host "4. POP3 Access for Outlook" -Fore DarkYellow
Write-host "5. IMAP Access for Outlook" -Fore DarkYellow
Write-host "6. MAPI Access for Outlook" -Fore DarkYellow
$Feature = Read-host "Please enter an option 1 to 6..."
Write-Host
Switch( $Feature ){
  1{Set-CasMailbox -Identity $Mailbox -MAPIBlockOutlookRpcHttp $false
    Write-host "Outlook Anywhere Access feature has been successfully Enabled for $Mailbox." -Fore Green
Write-Verbose $initstring5
if ($Log) {Write-Logfile $initstring5}
Return
End
  }
  2{Set-CASMailbox -Identity $Mailbox -ActiveSyncEnabled $true
    Write-host "Exchange ActiveSync feature has been successfully Enabled for $Mailbox." -Fore Green
Write-Verbose $initstring6
if ($Log) {Write-Logfile $initstring6}
Return
End
  }
  3{Set-CASMailbox -Identity $Mailbox -OWAEnabled $true
    Write-host "Outlook Web Access feature has been successfully Enabled for $Mailbox." -Fore Green
Write-Verbose $initstring7
if ($Log) {Write-Logfile $initstring7}
Return
End
  }
  4{Set-CASMailbox -Identity $Mailbox -PopEnabled $true
    Write-host "POP3 Access for Outlook feature has been successfully Enabled for $Mailbox." -Fore Green
Write-Verbose $initstring8
if ($Log) {Write-Logfile $initstring8}
Return
End
  }
  5{Set-CASMailbox -Identity $Mailbox -ImapEnabled $true
    Write-host "IMAP Access for Outlook feature has been successfully Enabled for $Mailbox." -Fore Green
Write-Verbose $initstring9
if ($Log) {Write-Logfile $initstring9}
Return
End
  }
  6{Set-CASMailbox -Identity $Mailbox -MapiEnabled $true
    Write-host "MAPI Access for Outlook feature has been successfully Enabled for $Mailbox." -Fore Green
Write-Verbose $initstring10
if ($Log) {Write-Logfile $initstring10}
Return
End
  }
  default{Write-Warning "Mailbox Feature to Enable has not been selected!!!"
Return
End
  }
}
}

Function Disable
{
Write-Verbose $initstring4
if ($Log) {Write-Logfile $initstring4}
Write-host
Write-host "Below are the available options to Disable Mailbox Feature." -Fore Yellow
Write-host "1. Outlook Anywhere Access" -Fore DarkYellow
Write-host "2. Exchange ActiveSync" -Fore DarkYellow
Write-host "3. Outlook Web Access" -Fore DarkYellow
Write-host "4. POP3 Access for Outlook" -Fore DarkYellow
Write-host "5. IMAP Access for Outlook" -Fore DarkYellow
Write-host "6. MAPI Access for Outlook" -Fore DarkYellow
$Feature = Read-host "Please enter an option 1 to 6..."
Write-Host
Switch( $Feature ){
  1{Set-CasMailbox -Identity $Mailbox -MAPIBlockOutlookRpcHttp $true
    Write-host "Outlook Anywhere Access feature has been successfully Disabled for $Mailbox." -Fore Green
Write-Verbose $initstring11
if ($Log) {Write-Logfile $initstring11}
Return
End
  }
  2{Set-CASMailbox -Identity $Mailbox -ActiveSyncEnabled $false
    Write-host "Exchange ActiveSync feature has been successfully Disabled for $Mailbox." -Fore Green
Write-Verbose $initstring12
if ($Log) {Write-Logfile $initstring12}
Return
End
  }
  3{Set-CASMailbox -Identity $Mailbox -OWAEnabled $false
    Write-host "Outlook Web Access feature has been successfully Disabled for $Mailbox." -Fore Green
Write-Verbose $initstring13
if ($Log) {Write-Logfile $initstring13}
Return
End
  }
  4{Set-CASMailbox -Identity $Mailbox -PopEnabled $false
    Write-host "POP3 Access for Outlook feature has been successfully Disabled for $Mailbox." -Fore Green
Write-Verbose $initstring14
if ($Log) {Write-Logfile $initstring14}
Return
End
  }
  5{Set-CASMailbox -Identity $Mailbox -ImapEnabled $false
    Write-host "IMAP Access for Outlook feature has been successfully Disabled for $Mailbox." -Fore Green
Write-Verbose $initstring15
if ($Log) {Write-Logfile $initstring15}
Return
End
  }
  6{Set-CASMailbox -Identity $Mailbox -MapiEnabled $false
    Write-host "MAPI Access for Outlook feature has been successfully Disabled for $Mailbox." -Fore Green
Write-Verbose $initstring16
if ($Log) {Write-Logfile $initstring16}
Return
End
  }
  default{Write-Warning "Mailbox Feature to Disable has not been selected!!!"
Return
End
  }
}
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

Write-Verbose $initstring2
if ($Log) {Write-Logfile $initstring2}


#...................................
# Script
#...................................

if(!$Mailbox)
{Write-Warning "No Parameter has been provided. Script has been terminated."}

If($Mailbox)
{
$name = Get-CasMailbox -Identity $Mailbox
Write-Host "User Display Name: " $name.DisplayName
Write-Host "Logon Account Name: " $name.SamAccountName
Write-Host "Primary Email Address: " $name.PrimarySmtpAddress
Write-Host "Organizational Unit: " $name.Identity
If($name.MAPIBlockOutlookRpcHttp -eq $false)
{
Write-Host "Outlook Anywhere Access: Enabled"
}
Else
{
Write-Host "Outlook Anywhere Access: Disabled"
}
If($name.ActiveSyncEnabled -eq $true)
{
Write-Host "Exchange ActiveSync: Enabled"
}
Else
{
Write-Host "Exchange ActiveSync: Disabled"
}
If($name.HasActiveSyncDevicePartnership -eq $true)
{
Write-Host "Using Mobile Device: Yes"
}
Else
{
Write-Host "Using Mobile Device: No"
}

If($name.OWAEnabled -eq $true)
{
Write-Host "Outlook Web Access: Enabled"
}
Else
{
Write-Host "Outlook Web Access: Disabled"
}
If($name.PopEnabled -eq $true)
{
Write-Host "POP3 Access: Enabled"
}
Else
{
Write-Host "POP3 Access: Disabled"
}
If($name.ImapEnabled -eq $true)
{
Write-Host "IMAP Access: Enabled"
}
Else
{
Write-Host "IMAP Access: Disabled"
}
If($name.MapiEnabled -eq $true)
{
Write-Host "MAPI Access: Enabled"
}
Else
{
Write-Host "MAPI Access: Disabled"
}
}

If($EnableFeature){
Switch($EnableFeature){
default{Enable
if($Log) {
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile -Append
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0

Return
End
}}}}

If($DisableFeature){
Switch($DisableFeature){
default{Disable
if($Log) {
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile -Append
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0
Return
End
}}}}

if ($Log) {
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile -Append
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0
Return
End
}

If($Prompt)
{
Write-Host ""
Write-Host "         This Script is used to View/Manage Mailbox Features." -Fore Cyan
Write-Host "       ********************************************************" -Fore Yellow
Write-Host ""
$Mailbox = Read-Host "Please Enter the User Email Address or Alias here"

If($Mailbox)
{
$name = Get-CasMailbox -Identity $Mailbox
Write-Host "User Display Name: " $name.DisplayName
Write-Host "Logon Account Name: " $name.SamAccountName
Write-Host "Primary Email Address: " $name.PrimarySmtpAddress
Write-Host "Organizational Unit: " $name.Identity
If($name.MAPIBlockOutlookRpcHttp -eq $false)
{
Write-Host "Outlook Anywhere Access: Enabled"
}
Else
{
Write-Host "Outlook Anywhere Access: Disabled"
}

If($name.ActiveSyncEnabled -eq $true)
{
Write-Host "Exchange ActiveSync: Enabled"
}
Else
{
Write-Host "Exchange ActiveSync: Disabled"
}
If($name.HasActiveSyncDevicePartnership -eq $true)
{
Write-Host "Using Mobile Device: Yes"
}
Else
{
Write-Host "Using Mobile Device: No"
}

If($name.OWAEnabled -eq $true)
{
Write-Host "Outlook Web Access: Enabled"
}
Else
{
Write-Host "Outlook Web Access: Disabled"
}
If($name.PopEnabled -eq $true)
{
Write-Host "POP3 Access: Enabled"
}
Else
{
Write-Host "POP3 Access: Disabled"
}
If($name.ImapEnabled -eq $true)
{
Write-Host "IMAP Access: Enabled"
}
Else
{
Write-Host "IMAP Access: Disabled"
}
If($name.MapiEnabled -eq $true)
{
Write-Host "MAPI Access: Enabled"
}
Else
{
Write-Host "MAPI Access: Disabled"
}

Write-Host "Kindly confirm if you would like to Enable or Disable Mailbox Feature [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No" 
$ActionRequired = Read-Host "Confirm"

If($ActionRequired -eq 'Y')
{
Write-host "Below are the available options to Manage Mailbox Features." -Fore Yellow
Write-host "1. Enable Mailbox Feature" -Fore Yellow
Write-host "2. Disable Mailbox Feature" -Fore Yellow
$call = Read-host "Please select any option and enter 1 or 2"
Write-host

Switch( $call ){
  1{Enable}
  2{Disable}
  default{Write-Warning "No Option has been selected. Script has been terminated."}
}
}

If($ActionRequired -eq 'N')
{Write-Host "This Script has been completed." -ForegroundColor Green
Return
End
}

If($ActionRequired -ne 'Y')
{Write-Warning "No Option has been selected. Script has been terminated."}

}
}

#...................................
# End
#...................................

if($Log) {
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile -Append
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0
}

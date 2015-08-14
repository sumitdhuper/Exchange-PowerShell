<#
.SYNOPSIS
Get-MBXInfo.ps1

.DESCRIPTION 
This script is used to get mailbox information comprising details from Get-Mailbox & Get-MailboxStatistics Cmdlet.
You could use this script with MAILBOX parameter or without any parameter which would show text input prompt message to use this script.

.OUTPUTS
Display information about a given mailbox.

.PARAMETER MAILBOX
The MAILBOX parameter could be used by administrator to bypass text input prompt message function of this script.

.PARAMETER Log
Used to generate a log file for the script.

.EXAMPLE
Get-Help .\Get-MBXInfo.ps1 -Full

.EXAMPLE
.\Get-MBXInfo.ps1
Returns multiple text input prompt messages which could be helpful for helpdesk support staff.

.EXAMPLE
.\Get-MBXInfo.ps1 -Mailbox SAM999
Returns information about a given mailbox when using Alias.

.EXAMPLE
.\Get-MBXInfo.ps1 -Mailbox sumit.dhuper@example.net
Returns information about a given mailbox when using Email Address.

.EXAMPLE
.\Get-MBXInfo.ps1 -Log
Returns a report with a log file at the script file path.

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
V1.00, 09/07/2015 - Initial version
V1.01, 14/08/2005 - Added Mailbox Parameter which is optional, Added Log file generation function and updated color codes
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[string]$Mailbox,

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

Write-Host $initstring0
if ($Log) {Write-Logfile $initstring0}

Write-Host $initstring1
if ($Log) {Write-Logfile $initstring1}


#...................................
# Script
#...................................

If($Mailbox)
{

$Email = (Get-Mailbox $Mailbox).PrimarySmtpAddress
$db = (Get-Mailbox $Mailbox).database
$quota = (Get-Mailbox $Mailbox).UseDatabaseQuotaDefaults
(Get-Mailbox $Mailbox | Get-MailboxStatistics | Add-Member -MemberType ScriptProperty -Name TotalItemSizeinMB -Value{$this.totalitemsize.value.ToMB()} -PassThru | Format-List DisplayName,TotalItemSizeinMB | Out-String).Trim()
Write-Host "PrimaryEmailAddress :" $Email
(Get-MailboxStatistics $Mailbox | Format-List Database,ItemCount,StorageLimitStatus,LastLogonTime | Out-String).Trim()

if ($quota -eq "true")
  {(Get-MailboxDatabase "$db" | Format-List ProhibitSendQuota,IssueWarningQuota | Out-String).Trim()

Write-Host "Note: Mailbox Quota Limit is getting inherited from Database." -fore DarkBlue -back Cyan
Write-Host "-----------------------------------------------------------" -fore Yellow
Return
End
}

Else
  {
(Get-Mailbox $Mailbox | Format-List ProhibitSendQuota, IssueWarningQuota | Out-String).Trim()

Write-Host "Note: Mailbox Quota Limit is applied on user's Mailbox." -fore DarkBlue -back Cyan
Write-Host "-------------------------------------------------------" -fore Yellow
Return
End
}

}

Write-Host "   This script will show the mailbox details for user's mailbox" -fore Yellow
Write-Host "******************************************************************" -fore Yellow
Write-Host
$Random = Get-Random -Maximum 1000 -Minimum 101
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host " to continue on this Script."
$Code = Read-Host "Enter the verification Code as showing above"

if ($Random -eq "$Code")
{Write-Host "Code has been verified successfully!!!" -ForegroundColor Green}

Else
    {Write-Host "ERROR: Verification Failed!!! Kindly run the Script again." -ForegroundColor Yellow -BackgroundColor Black
    Return
    End}
Write-Host
$user = Read-Host "Please enter user Alias or Email Address"
Write-Host
$Email = (Get-Mailbox $user).PrimarySmtpAddress
$db = (Get-Mailbox $user).database
$quota = (Get-Mailbox $user).UseDatabaseQuotaDefaults
(Get-Mailbox $user | Get-MailboxStatistics | Add-Member -MemberType ScriptProperty -Name TotalItemSizeinMB -Value{$this.totalitemsize.value.ToMB()} -PassThru | Format-List DisplayName,TotalItemSizeinMB | Out-String).Trim()
Write-Host "PrimaryEmailAddress :" $Email
(Get-MailboxStatistics $user | Format-List Database,ItemCount,StorageLimitStatus,LastLogonTime | Out-String).Trim()

if ($quota -eq "true")
  {(Get-MailboxDatabase "$db" | Format-List ProhibitSendQuota,IssueWarningQuota | Out-String).Trim()

Write-Host "Note: Mailbox Quota Limit is getting inherited from Database." -fore DarkBlue -back Cyan
Write-Host "-----------------------------------------------------------" -fore Yellow
}

Else
  {
(Get-Mailbox $user | Format-List ProhibitSendQuota, IssueWarningQuota | Out-String).Trim()

Write-Host "Note: Mailbox Quota Limit is applied on user's Mailbox." -fore DarkBlue -back Cyan
Write-Host "-------------------------------------------------------" -fore Yellow
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

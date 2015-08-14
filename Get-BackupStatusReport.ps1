<#
.SYNOPSIS
Get-BackupStatusReport.ps1

.DESCRIPTION 
This script is prepared to get Exchange Server Database status report and get the report in formatted HTML file via email if required.

.OUTPUTS
Exchange Server Database Backup status report to console and could send the report via mail in formatted HTML style.

.PARAMETER SERVER
Generates a report for all databases on the specified server.

.SWITCH PARAMETER SENDMAIL
The SendMail switch parameter could be used to schedule this script with service account.

.PARAMETER Log
Used to generate a log file for the script.

.EXAMPLE
Get-Help .\Get-BackupStatusReport.ps1 -Full
Returns information about this script.

.EXAMPLE
.\Get-BackupStatusReport.ps1
Returns a report with the backup statistics for all database on all mailbox servers.

.EXAMPLE
.\Get-BackupStatusReport.ps1 -Server MBX01
Returns a report with the backup statistics for all database on server MBX01.

.EXAMPLE
.\Get-BackupStatusReport.ps1 -SendEmail -MailFrom exchangereports@example.net -MailTo sumit.dhuper@example.net -MailServer smtp.example.net
.\Get-BackupStatusReport.ps1 -Server MBX01 -SendEmail -MailFrom exchangereports@example.net -MailTo sumit.dhuper@example.net -MailServer smtp.example.net

Returns a report with the backup statistics for all database on all mailbox servers and
sends an email report to the specified recipient. This switch could be used with -SERVER perameter if required.

.EXAMPLE
.\Get-BackupStatusReport.ps1 -Log
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
V1.00, 12/08/2015 - Initial version
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[string]$Server,

	[Parameter( Mandatory=$false)]
	[switch]$Log,

	[Parameter( Mandatory=$false)]
	[switch]$SendMail,

	[Parameter( Mandatory=$false)]
	[string]$MailFrom,

	[Parameter( Mandatory=$false)]
	[string]$MailTo,

	[Parameter( Mandatory=$false)]
	[string]$MailServer

	)


#...................................
# Variables
#...................................

$scriptname = $MyInvocation.MyCommand.Name
$now = Get-Date
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logfile = "$myDir\$scriptname.log"
$Attachment = ".\BackupStatus.html"
$messagesubject = "Exchange Server Database Backup Status - $now"

#...................................
# Email Settings
#...................................

$smtpsettings = @{
	To =  $MailTo
	From = $MailFrom
    Subject = $messagesubject
	SmtpServer = $MailServer
	}

#...................................
# SMTP Server Name
#...................................

$MailServer = "SIND042301.wsatkins.com"

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

If($Server)
{
#...................................
# Total Object Count
#...................................
$DATABASE_COUNT = (Get-MailboxServer -Identity $Server| Get-MailboxDatabase).count

#...................................
# Progress Status
#...................................

for ( $t =0; $t -lt $DATABASE_COUNT ; $t++) 
{$status = [int] (($t/$DATABASE_COUNT) * 100)
  Write-Progress -Activity "Working..." -PercentComplete $status -CurrentOperation "$status% complete" -Status "Please wait."
   Start-Sleep 1
}

$Databases = Get-MailboxServer -Identity $Server | Get-MailboxDatabase -Status | Select-Object Server,Name,LastFullBackup,LastIncrementalBackup
#Get report generation timestamp
$reportime = Get-Date
$a = "<style>"
$a = $a + "BODY{font-family: Arial; font-size: 8pt;}"
$a = $a + "H1{font-size: 16px;}"
$a = $a + "H2{font-size: 12px;}"
$a = $a + "TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}"
$a = $a + "TH{border: 1px solid black; background: #007fff; padding: 5px; color: #f0ffff;}"
$a = $a + "TD{border: 1px solid black; padding: 5px; }"
$a = $a + "td.pass{background: #7FFF00;}"
$a = $a + "td.warn{background: #FFE600;}"
$a = $a + "td.fail{background: #FF0000; color: #ffffff;}"
$a = $a + "</style>"
$a = $a + "<body>"
$a = $a + "<h1 align=""center"">Exchange Server Database Backup Report</h1>"
$a = $a + "<h2 align=""center"">Generated: $reportime</h2>"
$htmlfile = $Databases | Select-Object Server,Name,LastFullBackup,LastIncrementalBackup | ConvertTo-HTML -head $a | Out-File .\BackupStatus.html
$htmlreport = Get-Content -Path .\BackupStatus.html -Raw

#..........................................
# Presenting Database Status on PS Console
#..........................................

$Databases
}

#If no Server perameter provided then it will run against all servers.
Else
{
#...................................
# Total Object Count
#...................................
$DATABASE_COUNT = (Get-MailboxServer | Get-MailboxDatabase).count

#...................................
# Progress Status
#...................................

for ( $t =0; $t -lt $DATABASE_COUNT ; $t++) 
{$status = [int] (($t/$DATABASE_COUNT) * 100)
  Write-Progress -Activity "Working..." -PercentComplete $status -CurrentOperation "$status% complete" -Status "Please wait."
   Start-Sleep 1
}

$Databases = Get-MailboxServer | Get-MailboxDatabase -Status | Select-Object Server,Name,LastFullBackup,LastIncrementalBackup
#Get report generation timestamp
$reportime = Get-Date
$a = "<style>"
$a = $a + "BODY{font-family: Arial; font-size: 8pt;}"
$a = $a + "H1{font-size: 16px;}"
$a = $a + "H2{font-size: 12px;}"
$a = $a + "TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}"
$a = $a + "TH{border: 1px solid black; background: #007fff; padding: 5px; color: #f0ffff;}"
$a = $a + "TD{border: 1px solid black; padding: 5px; }"
$a = $a + "td.pass{background: #7FFF00;}"
$a = $a + "td.warn{background: #FFE600;}"
$a = $a + "td.fail{background: #FF0000; color: #ffffff;}"
$a = $a + "</style>"
$a = $a + "<body>"
$a = $a + "<h1 align=""center"">Exchange Server Database Backup Report</h1>"
$a = $a + "<h2 align=""center"">Generated: $reportime</h2>"
$htmlfile = $Databases | Select-Object Server,Name,LastFullBackup,LastIncrementalBackup | ConvertTo-HTML -head $a | Out-File .\BackupStatus.html
$htmlreport = Get-Content -Path .\BackupStatus.html -Raw

#..........................................
# Presenting Database Status on PS Console
#..........................................

$Databases

}

#..........................................
# SENDMAIL PARAMETER
#..........................................

If($SendMail)
{
Write-Host "Sending email report..."
Send-MailMessage @smtpsettings -Attachments $Attachment -Body $htmlreport -BodyAsHtml -ErrorAction STOP

Write-Host "INFO: Backup Report has been sent successfully on the above mentioned email address." -ForegroundColor Green
Return
End
}

#..........................................
# Email Report without -SendMail
#..........................................

$EmailRequired = Read-Host "Please Enter 'Y' if Backup Report is Required on Email or press any key to end this Script"

If ($EmailRequired -eq 'Y')
{
$MailTo = Read-Host "Mail Sending FROM"
$MailFrom = Read-Host "Mail Sending TO"

Write-Host "Sending email report..."
Send-MailMessage -Attachments $Attachment -To $MailTo -From $MailFrom -Subject "$messageSubject" -Body $htmlreport -BodyAsHtml -SmtpServer $MailServer

Write-Host "INFO: Backup Report has been sent successfully on the above mentioned email address." -ForegroundColor Green
}

Else
{
Write-Host "INFO: Script has been completed." -ForegroundColor Green
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

<#
.SYNOPSIS
Queue-Monitor.ps1

.DESCRIPTION 
This script could be used to monitor mail queues on all Hub Transport Servers for a given threshold count of number of emails in single queue.

.OUTPUTS
This script generate a CSV file report for mail queues above threshold count 
and attach file while sending alert email to provided recipient.

.SWITCH PARAMETER SENDMAIL
The SendMail switch parameter could be used to schedule this script with service account.

.PARAMETER Log
Used to generate a log file for the script.

.EXAMPLE
Get-Help .\Queue-Monitor.ps1 -Full
Returns information about this script.

.EXAMPLE
.\Queue-Monitor.ps1
Returns Mail Queue status if it is above/below threshold count 
and provide text input prompt to specify Sender and Recipient email address to get Alert email.

.EXAMPLE
.\Queue-Monitor.ps1 -SendEmail -MailFrom exchangereports@example.net -MailTo sumit.dhuper@example.net -MailServer smtp.example.net

Returns Mail Queue status if it is above/below threshold count 
and sends an email report to the specified recipient if any Queue found above threshold count.

.EXAMPLE
.\Queue-Monitor.ps1 -Log
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
V1.00, 06/08/2015 - Initial version
V1.00, 14/08/2015 - Added SendMail parameter into this script and now it could be used as scheduled task.
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[switch]$SendMail,

	[Parameter( Mandatory=$false)]
	[string]$MailFrom,

	[Parameter( Mandatory=$false)]
	[string]$MailTo,

	[Parameter( Mandatory=$false)]
	[string]$MailServer,

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
$Attachment = ".\QueueResult.csv"
$MessageSubject = "Exchange Server High Mail Queue Alert - $now"

#......................................................
# Change Alert Threshold Count as per your requirement
#......................................................

$AlertThresholdCount = 30

#...................................
# Email Settings
#...................................

$smtpsettings = @{
	To =  $MailTo
	From = $MailFrom
    Subject = $MessageSubject
	SmtpServer = $MailServer
	}

#.................................................
# Change SMTP Server Name as per your environment
#.................................................

$MailServer = "smtp.example.net"

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

If($SendMail)
{

#...................................
# Progress Status
#...................................

$TRANSPORTSERVER_COUNT = (Get-TransportServer).count
for ( $t =0; $t -lt $TRANSPORTSERVER_COUNT ; $t++) 
{$status = [int] (($t/$TRANSPORTSERVER_COUNT) * 100)
  Write-Progress -Activity "Working..." -PercentComplete $status -CurrentOperation "$status% complete" -Status "Please wait."
   Start-Sleep 1
}

#...................................
# Script Variable
#...................................

$Queue = Get-TransportServer | Get-Queue | where messagecount -gt "$AlertThresholdCount"
$reportfile = $Queue | Export-Csv .\QueueResult.csv
$a = "<style>"
$a = $a + "BODY{font-family: Arial; font-size: 8pt;}"
$a = $a + "H1{font-size: 16px;}"
$a = $a + "H2{font-size: 14px;}"
$a = $a + "H3{font-size: 12px;}"
$a = $a + "TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}"
$a = $a + "TH{border: 1px solid black; background: #007fff; padding: 5px; color: #f0ffff;}"
$a = $a + "TD{border: 1px solid black; padding: 5px; }"
$a = $a + "td.pass{background: #7FFF00;}"
$a = $a + "td.warn{background: #FFE600;}"
$a = $a + "td.fail{background: #FF0000; color: #ffffff;}"
$a = $a + "</style>"
$a = $a + "<body>"
$a = $a + "<h1 align=""center"">Exchange Server Mail Queue Report</h1>"
$a = $a + "<h3 align=""center"">Generated: $now</h3>"
$htmlfile = $Queue | Select-Object Identity,DeliveryType,Status,MessageCount,NextHopDomain | ConvertTo-HTML -head $a | Out-File .\QueueResult.html
$htmlreport = Get-Content -Path .\QueueResult.html -Raw
$Count = $Queue.MessageCount

If ($Count -gt $AlertThresholdCount)
{
Send-MailMessage @smtpsettings -Attachments $Attachment -Body $htmlreport -BodyAsHtml -ErrorAction Stop

Write-Host "INFO: High Queue Alert Mail has been sent to Exchange Administrator." -ForegroundColor Yellow
Return
End
}

Else
{

Write-Host "INFO: No Alert mail sent as all mail queues are below threshold limit of" -ForegroundColor Green -NoNewline; Write-Host " $AlertThresholdCount " -ForegroundColor Yellow -NoNewline; Write-Host "messages in All Local/Remote Delivery Queue." -ForegroundColor Green
Return
End
}
}

Else
{

Write-Host "   This script is used to track Mail Queue status and sending Alrets." -fore Yellow
Write-Host "************************************************************************" -fore Yellow
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
Write-Host "INFO: Checking Mail Queue Status on All Exchange Hub Transport Servers...." -ForegroundColor Yellow

#...................................
# Progress Status
#...................................
$TRANSPORTSERVER_COUNT = (Get-TransportServer).count
for ( $t =0; $t -lt $TRANSPORTSERVER_COUNT ; $t++) 
{$status = [int] (($t/$TRANSPORTSERVER_COUNT) * 100)
  Write-Progress -Activity "Working..." -PercentComplete $status -CurrentOperation "$status% complete" -Status "Please wait."
   Start-Sleep 1
}

#...................................
# Script Variable
#...................................

$Queue = Get-TransportServer | Get-Queue | where messagecount -gt "$AlertThresholdCount"
$reportfile = $Queue | Export-Csv .\QueueResult.csv
$a = "<style>"
$a = $a + "BODY{font-family: Arial; font-size: 8pt;}"
$a = $a + "H1{font-size: 16px;}"
$a = $a + "H2{font-size: 14px;}"
$a = $a + "H3{font-size: 12px;}"
$a = $a + "TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}"
$a = $a + "TH{border: 1px solid black; background: #007fff; padding: 5px; color: #f0ffff;}"
$a = $a + "TD{border: 1px solid black; padding: 5px; }"
$a = $a + "td.pass{background: #7FFF00;}"
$a = $a + "td.warn{background: #FFE600;}"
$a = $a + "td.fail{background: #FF0000; color: #ffffff;}"
$a = $a + "</style>"
$a = $a + "<body>"
$a = $a + "<h1 align=""center"">Exchange Server Mail Queue Report</h1>"
$a = $a + "<h3 align=""center"">Generated: $now</h3>"
$htmlfile = $Queue | Select-Object Identity,DeliveryType,Status,MessageCount,NextHopDomain | ConvertTo-HTML -head $a | Out-File .\QueueResult.html
$htmlreport = Get-Content -Path .\QueueResult.html -Raw
$Count = $Queue.MessageCount

If ($Count -gt $AlertThresholdCount)
{
$MailTo = Read-Host "Mail Sending FROM"
$MailFrom = Read-Host "Mail Sending TO"
Send-MailMessage -Attachments $Attachment -To $MailTo -From $MailFrom -Subject "$MessageSubject" -Body $htmlreport -BodyAsHtml -SmtpServer $MailServer -ErrorAction Stop

Write-Host "INFO: High Queue Alert Mail has been sent to Exchange Administrator." -ForegroundColor Yellow
}
Else
{

Write-Host "INFO: No Alert mail sent as all mail queues are below threshold limit of" -ForegroundColor Green -NoNewline; Write-Host " $AlertThresholdCount " -ForegroundColor Yellow -NoNewline; Write-Host "messages in All Local/Remote Delivery Queue." -ForegroundColor Green
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

Write-Host "Author: Sumit Dhuper @ 09/07/2015 v1.1" -fore DarkGray -back Cyan
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
$user = Read-Host "Please enter user 4x4 or email address"
Write-Host
$Email = (Get-Mailbox $user).PrimarySmtpAddress
$db = (Get-Mailbox $user).database
$quota = (Get-Mailbox $user).UseDatabaseQuotaDefaults
(Get-Mailbox $user | Get-MailboxStatistics | Add-Member -MemberType ScriptProperty -Name TotalItemSizeinMB -Value{$this.totalitemsize.value.ToMB()} -PassThru | Format-List DisplayName,TotalItemSizeinMB | Out-String).Trim()
Write-Host "PrimaryEmailAddress :" $Email
(Get-MailboxStatistics $user | Format-List Database,ItemCount,StorageLimitStatus,LastLogonTime | Out-String).Trim()

if ($quota -eq "true")
  {(Get-MailboxDatabase "$db" | Format-List ProhibitSendQuota,IssueWarningQuota | Out-String).Trim()

Write-Host "Note: Mailbox Quota Limit is getting applied from Database." -fore DarkGreen -back Cyan
Write-Host "-----------------------------------------------------------" -fore Yellow
}

Else
  {
(Get-Mailbox $user | Format-List ProhibitSendQuota, IssueWarningQuota | Out-String).Trim()

Write-Host "Note: Mailbox Quota Limit is applied on user's Mailbox." -fore DarkGreen -back Cyan
Write-Host "-------------------------------------------------------" -fore Yellow
}

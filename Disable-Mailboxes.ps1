$csv = Import-CSV .\Users.csv

$count = $csv.count

Write-Host "There are total " -NoNewline; Write-Host "$count" -ForegroundColor Yellow -NoNewline; Write-Host " users found in Users.csv file." 
Write-Host "Kindly confirm if would like to proceed to disable mailboxes? [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
$confirm = Read-Host " "

if($confirm -eq "Y")
{

$mailboxes = foreach($line in $csv){Get-Mailbox -Identity $line.users}

$mailboxes | Select ExchangeGuid,Database,Alias,PrimarySmtpAddress | Export-Csv .\DisabledMailboxes.csv -NoType

Foreach($line in $csv) {Disable-Mailbox -Identity $line.Users -Confirm:$False}

Write-Host "All $count users mailbox has been disabled." -fore "Green"

}

Else
{

Write-Host "Script has been terminated!!!" -fore "Yellow"

}

$csv = Import-CSV .\DisabledMailboxes.csv

$count = $csv.count

Write-Host "There are total " -NoNewline; Write-Host "$count" -ForegroundColor Yellow -NoNewline; Write-Host " users found in DisabledMailboxes.csv file." 
Write-Host "Kindly confirm if would like to proceed to connect mailboxes? [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
$confirm = Read-Host " "

if($confirm -eq "Y")
{

Foreach($line in $csv) {Connect-Mailbox -Identity $line.ExchangeGuid -Database $line.Database -Alias $line.Alias -User ("India\"+$line.Alias)}

Write-Host "All $count users mailbox has been successfully connected." -fore "Green"

}

Else
{

Write-Host "Script has been terminated!!!" -fore "Yellow"

}

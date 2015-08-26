Function Add
{
Write-Host "         This Script is used to Add Public Folder Client Permission." -Fore Cyan
Write-Host "       ***************************************************************" -Fore Yellow
Write-Host ""

$Random = Get-Random -Maximum 1000 -Minimum 101
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host " to continue on this Script."
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
Write-Host "Kindly confirm if you would like to apply permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
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
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -ForegroundColor Yellow
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
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -ForegroundColor Yellow
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
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host " to continue on this Script."
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
Write-Host "Kindly confirm if you would like to remove permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
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
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -ForegroundColor Yellow
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
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -ForegroundColor Yellow
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
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host " to continue on this Script."
$Code = Read-Host "Enter the verification Code as showing above"

if ($Random -eq "$Code")
{Write-Host "Code has been verified successfully!!!" -ForegroundColor Green}

Else
    {Write-Host "ERROR: Verification Failed!!! Kindly run the Script again." -ForegroundColor Yellow -BackgroundColor Black
    Return
    End}

$PF = Read-Host "Please enter the Public Folder path under All Public Folders including preceding '\' sign"
Write-host
Write-Host "Kindly confirm if you would like to apply permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
$subfolder = Read-Host " "

if ($subfolder -eq "$null")
{Write-Host "ERROR: Permission scope has not been confirmed!!!" -BackgroundColor black -ForegroundColor Yellow
  Return
  End
  }

if ($subfolder -eq "Y")
{
Get-PublicFolder -Identity "$PF" –Recurse | Get-PublicFolderClientPermission | ft user,access* -Wrap
}

if ($subfolder -eq "N")
{
Get-PublicFolder -Identity "$PF" | Get-PublicFolderClientPermission | ft user,access* -Wrap
}
}


Function AddBulk
{
Write-Host "         This Script is used to Add Public Folder Client Permission - BULK REQUEST." -Fore Cyan
Write-Host "       ******************************************************************************" -Fore Yellow
Write-Host ""

$Random = Get-Random -Maximum 1000 -Minimum 101
Write-Host "Please enter this verification code " -NoNewline; Write-Host " $Random " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host " to continue on this Script."
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
Write-Host "NOTE: CSV file sam.csv found under the Script folder and there are $count users SamAccountName/Alias found in the file!!!" -BackgroundColor Cyan -ForegroundColor DarkBlue
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
Write-Host "Kindly confirm if you would like to apply permission including Subfolders [" -NoNewline; Write-Host "Y" -ForegroundColor Yellow -NoNewline; Write-Host "]Yes or [" -NoNewline; Write-Host "N" -ForegroundColor Yellow -NoNewline; Write-Host "]No"  -Fore Cyan -NoNewline;
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
  1{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Author"}}
  2{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Reviewer"}}
  3{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights "PublishingAuthor"}}
  4{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights "PublishingEditor"}}
  5{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Editor"}}
  6{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” –Recurse | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Owner"}}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -ForegroundColor Yellow
  Return
  End
}
}}

if ($subfolder -eq "N")
{
Switch( $AR ){
  1{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Author"}}
  2{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Reviewer"}}
  3{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "PublishingAuthor"}}
  4{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "PublishingEditor"}}
  5{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Editor"}}
  6{foreach ($line in $csv) {Get-PublicFolder –Identity “$PF” | Add-PublicFolderClientPermission –User $line.sam –AccessRights "Owner"}}
  default{Write-Host "ERROR: Permission level has not been selected for the above mentioned public folder!!!" -BackgroundColor black -ForegroundColor Yellow
  Return
  End
}
}

Write-host "Permission has been successfully applied!!!" -Fore Green
}

}
Write-Host "Author: Sumit Dhuper @ 21/07/2015 v1.2" -fore DarkGray -back Cyan
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

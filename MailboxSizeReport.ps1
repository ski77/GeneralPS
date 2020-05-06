<#

.Requires -version 2 - Runs in Exchange Management Shell

.SYNOPSIS
.\MailboxSizeReport.ps1 - It Can Display all the Mailbox Size with Item Count,Database,Server Details

Or It can Export to a CSV file

Or You can Enter WildCard to Display or Export


Example 1

[PS] C:\>.\MailboxSizeReport.ps1


Mailbox Size Report
----------------------------

1.Display in Exchange Management Shell

2.Export to CSV File

3.Enter the Mailbox Name with Wild Card (Export)

4.Enter the Mailbox Name with Wild Card (Display)

Choose The Task: 1

Display Name                  Primary SMTP address          TotalItemSize                 ItemCount
------------                  --------------------          -------------                 ---------
Tes433                        Tes433@Welcome.com
Test                          Test@testcareexchange.biz     335.9 KB (343,933 bytes)      40
Test X500                     TestX500@Testexchange.biz     6.544 KB (6,701 bytes)        3
Test100                       test100@testcareexchange.biz  40.74 KB (41,719 bytes)       7
Test22                        Test22@Testexchange.biz       60.04 KB (61,483 bytes)       7
Test3                         Test3@testcareexchange.biz    364.7 KB (373,503 bytes)      31
Test33                        Test332@testcareexchange.biz  93.34 KB (95,585 bytes)       6
Test33                        Test33@FSD.com                5.335 KB (5,463 bytes)        3
Test3331                      Test3331@Testexchange.biz     24.14 KB (24,720 bytes)       2
Test46                        Test46@testcareexchange.biz   254 KB (260,071 bytes)        21

Example 2

[PS] C:\>.\MailboxSizeReport.ps1


Mailbox Size Report
----------------------------

1.Display in Exchange Management Shell

2.Export to CSV File

3.Enter the Mailbox Name with Wild Card (Export)

4.Enter the Mailbox Name with Wild Card (Display)

Choose The Task: 2
Enter the Path of CSV file (Eg. C:\Report.csv): C:\MailboxReport.csv

.Author
Written By: Satheshwaran Manoharan

Change Log
V1.0, 10/08/2014 - Initial version


#>

Write-host "

Mailbox Size Report
----------------------------

1.Display in Exchange Management Shell

2.Export to CSV File

3.Enter the Mailbox Name with Wild Card (Export)

4.Enter the Mailbox Name with Wild Card (Display)"-ForeGround "Cyan"

#----------------
# Script
#----------------

Write-Host "               "

$number = Read-Host "Choose The Task"
$output = @()
switch ($number) 
{

1 {

$AllMailbox = Get-mailbox -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount

Write-Output $Userobj

}

;Break}

2 {

$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\Report.csv)" 

$AllMailbox = Get-mailbox -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientType
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Database" -Value $Stats.Database
$userObj | Add-Member NoteProperty -Name "ServerName" -Value $Stats.ServerName
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize
$userObj | Add-Member NoteProperty -Name "DatabaseProhibitSendReceiveQuota" -Value $Stats.DatabaseProhibitSendReceiveQuota
$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime

$output += $UserObj  

}

$output | Export-csv -Path $CSVfile -NoTypeInformation

;Break}

3 {

$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)" 

$MailboxName = Read-Host "Enter the Mailbox name or Range (Eg. Mailboxname , Mi*,*Mik)"

$AllMailbox = Get-mailbox $MailboxName -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientType
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Database" -Value $Stats.Database
$userObj | Add-Member NoteProperty -Name "ServerName" -Value $Stats.ServerName
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize
$userObj | Add-Member NoteProperty -Name "DatabaseProhibitSendReceiveQuota" -Value $Stats.DatabaseProhibitSendReceiveQuota
$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime

$output += $UserObj  

}

$output | Export-csv -Path $CSVfile -NoTypeInformation

;Break}

4 {

$MailboxName = Read-Host "Enter the Mailbox name or Range (Eg. Mailboxname , Mi*,*Mik)"

$AllMailbox = Get-mailbox $MailboxName -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount

Write-Output $Userobj

}

;Break}

Default {Write-Host "No matches found , Enter Options 1 or 2" -ForeGround "red"}

}
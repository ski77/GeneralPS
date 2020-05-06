<#

.Synopsis
   ActiveSync Statistics is a small PowerShell Scipt which can be used in O365 to fetch ActiveSync Device Statistics of Mailboxes and manage rouge ActveSync Devices.

   Developed by: Noble K Varghese

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.1, 26 June 2015
		#Initial Release
	Version 1.2, 21 October 2015
		#Added CustomAttribute6 and CustomAttribute7 to the Script Output.

.DESCRIPTION
   ActiveSync Statistics.ps1 is a PowerShell Sciprt for Office365. It helps the Admin in collecting ActiveSync Device Statistics of Office 365 mailboxes, there by managing Rouge ActiveSync Devices.
   On completion, the Script creates one html report as the output in the current working directory. This scripts supports PowerShell 2.0 & 3.0. I am using 3.0 though. 

.ActiveSync Statistics.ps1
   To Run the Script go to PowerShell and Start It. Eg: PS E:\PowerShellWorkshop> .\ActiveSync Statistics.ps1

.Output Logs
   The Script creates one html report as the output in the present working directory in the format yyyymmddHHMMSS_ActiveSyncStatistics.html

.Function MobileDeviceConfig
    This Script works based on a Function. This function checks for the ActiveSync devices of the users. It will fetch multiple as well as single device reports.

#>

#Function
Function MobileDeviceConfig ([String]$Mbx,[String]$Upn,[String]$CustAtt6,[String]$CustAtt7)
{
	$i=-1
	$DevStat = Get-MobileDeviceStatistics -Mailbox $mbx
	foreach ($Dev in $DevStat)
	{
		if($i -eq "$Dev.Count")
		{
			$tab1+= "</tr><tr><td align=""center"">$($mbx)</td><td align=""center"">$($Upn)</td><td align=""center"">$($CustAtt6)</td><td align=""center"">$($CustAtt7)</td><td align=""center"">$($Dev.DeviceUserAgent)</td><td align=""center"">$($Dev.DeviceModel)</td><td align=""center"">$($Dev.DeviceOS)</td><td align=""center"">$($Dev.FirstSyncTime)</td><td align=""center"">$($Dev.LastSyncAttemptTime)</td><td align=""center"">$($Dev.LastSuccessSync)</td>"
		}
		else
		{
			$tab1+= "<td align=""center"">$($mbx)</td><td align=""center"">$($Upn)</td><td align=""center"">$($CustAtt6)</td><td align=""center"">$($CustAtt7)</td><td align=""center"">$($Dev.DeviceUserAgent)</td><td align=""center"">$($Dev.DeviceModel)</td><td align=""center"">$($Dev.DeviceOS)</td><td align=""center"">$($Dev.FirstSyncTime)</td><td align=""center"">$($Dev.LastSyncAttemptTime)</td><td align=""center"">$($Dev.LastSuccessSync)</td></tr>"
			$i++
		}	
	}
	$Global:MobileDevicesOut+=$tab1
}

#Main

$Watch = [System.Diagnostics.Stopwatch]::StartNew()
[String]$OutputFile1 = "$((Get-Date -uformat %Y%m%d%H%M%S).ToString())_ActiveSyncStatistics.html"
$j=1
$tabr1 = "<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""10""><font color=""#FFFFFF"">ActiveSync Statistics</font></th></tr><tr bgcolor=""#63635D"" align=""center""><th><font color=""#FFFFFF"">PrimarySmtpAddress</font></th><th><font color=""#FFFFFF"">UserPrincipaName</font></th><th><font color=""#FFFFFF"">CustomAttribute6</font></th><th><font color=""#FFFFFF"">CustomAttribute7</font></th><th><font color=""#FFFFFF"">DeviceUserAgent</font></th><th><font color=""#FFFFFF"">DeviceModel</font></th><th><font color=""#FFFFFF"">DeviceOS</font></th><th><font color=""#FFFFFF"">FirstSyncTime</font></th><th><font color=""#FFFFFF"">LastSyncAttemptTime</font></th><th><font color=""#FFFFFF"">LastSuccessSync</font></th></tr><tr>"
Write-Host -ForegroundColor Green -Object "please wait for me to fetch the total mailboxes in your organization..."
$mailall = Get-Mailbox -Resultsize Unlimited
$count = $mailall.Count
Write-Host -ForegroundColor Green -Object "found $count mailboxes..."
foreach ($mail in $mailall)
{
	$prsmtp = $mail.primarysmtpaddress
	$UserUpn = $mail.UserPrincipalName
	$CuAt6 = $mail.CustomAttribute6
	$CuAt7 = $mail.CustomAttribute7
	$DevDet = Get-MobileDeviceStatistics -Mailbox $prsmtp
	if($DevDet -ne $Null)
	{
		MobileDeviceConfig -Mbx $prsmtp -Upn $UserUpn -CustAtt6 $CuAt6 -CustAtt7 $CuAt7
	}
	Write-Progress -Activity "Processing User $mail" -status "$j Out Of $count Completed" -percentcomplete ($j / $count*100)
	$j++
}
$Header1="
	<html>
	<body>
	<font size=""1"" face=""Arial,sans-serif"">
	<h3 align=""center"">ActiveSync Statistics</h3>
	<a name=""top""><h4 align=""center"">Generated On $((Get-Date).ToString())</h4></a>
	"
$Footer1="</table></center><br><br>
	<font size=""1"" face=""Arial,sans-serif"">Scripted by <a href="""">Noble K Varghese</a> 
	Elapsed Time To Complete This Report(mm/ss): $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())
	<br><p>THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED AS IS WITHOUT
	<br>WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
	<br>LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
	<br>FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
	<br>RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	</font></body></html>"

$Watch.Stop()

$output = $Header1+$tabr1+$MobileDevicesOut+$Footer1
$output | Out-File $OutputFile1
Clear-Variable MobileDevicesOut -Scope global

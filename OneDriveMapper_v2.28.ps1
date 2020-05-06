######## 
#OneDriveMapper v2.28
#Copyright:     Free to use, please leave this header intact 
#Author:        Jos Lieben (OGD) 
#Company:       OGD (http://www.ogd.nl) 
#Script help:   http://www.liebensraum.nl 
#Purpose:       This script maps Onedrive for Business and maps a configurable number of Sharepoint Libraries
######## 
######## 
#Requirements: 
######## 
<# 
To use ADFS SSO, your federation URL (fs.domain.com) should be in Internet Explorer's Intranet Sites, windows authentication should be enabled in IE. 
Users should also have the same login UPN on their client, as in Office 365. The script will attempt to log into Office 365 using either the full  
user UPN from Active Directory (if lookupUPNbySAM is enabled) or by using the SamAccountName + the domain name. 
If you run any type of mapping scripts to your drive, make sure this script runs first or the drive will not be available. 
If you use a desktop management tool like RES PowerFuse, make sure your users are allowed to start a COM object and run powershell scripts. 
#> 
 
######## 
#Changelog: 
######## 
#V1.1: Added support for ADFS 
#V1.2: Added autoProvisioning, additional IE health checks 
#V1.3: Added checks for WebDav (WebClient) service 
#V1.4: Additional checks and automatic ProtectedMode fix 
#V1.5: Added DriveLabel parameter 
#V1.6: Form display fix (GPO bug) and Driveletter label persistence 
#V1.7: Added support for forcing a specific username and/or password, dealing with non domain joined machines using ADFS and non-standard library names 
#V1.8: Removed MaxAttempts setting, added automatic detection of changed usernames, added removal of existing failed drivemapping 
#V1.8a: Added conversion to lowercase of relevant user input to prevent problems when matching url's 
#V1.9: useADFS removed: this is now autodetected. Added sharepoint direct mapping.   
#V1.9a: added checks to verify Office is installed, Sharepoint is in Trusted Sites and WebDav file locking is disabled 
#V1.9b: added check for explorer.exe running, and option to restart it after mapping 
#V1.9c: added account splitter check (for people who use the same email for O365 and their normal MS account) 
#V2.0: enhanced the explorer.exe check to look only for own processes, added an IE runnning check and an option IE kill if found running 
#v2.1: added a check for the IE FirstRun wizard and a slight delay when restarting the IE Com Object to avoid issues with clean user profiles 
#V2.1: fixes a bug in Citrix, causing processess of other users to be returned 
#V2.1: revamped the ADFS redirection detection and login triggers to prevent slow responses to cause the script to fail 
#V2.1: improved zone map issue detection to include 3 alternate locations (machine, machine gpo, user gpo) where the registry can be saved 
#V2.1: Added detection of the 'HIDDEN' attribute of the redirection container for ADFS 
#V2.2: I got tired of the differences between attributes on the login page and the instability this causes, so several methods are implemented, don't forget to set adfsWaitTime
#V2.21: More ruthless cleanup of the COM objects used
#V2.22: Comments in Dutch -> English. Parameterised the ADFS control names for those who use a customized ADFS page. Cleanup. Additional zonecheck
#V2.23: Added a check to see IF the driveletter exists, it actually maps (approximately) to the right location, otherwise it will delete the mapping and remap it to the right location
#V2.23: Added an option to stop script execution if ADFS fails and two minor bugfixes reported by Martin Revard
#V2.24: added customization for stichting Sorg in the Netherlands to map a configurable number of Sharepoint Libraries in addition to O4B
#V2.25: solution for invisible drives when running as an admin. Make sure you set $restart_explorer to $True if you have users who are admin
#V2.26: Fixed multi-domain cookies not being registered (which causes sharepoint mappings to fail while O4B mappings work fine)
#V2.27: Fixed a bug in ProtectedMode storing values as String instead of DWord and better ADFS redirection and detection of invalid zonemaps configured through a GPO and added urlOpenAfter parameter
#V2.28: Support for Auto-Acceleration in Sharepoint Online (or O4B). https://support.office.com/en-us/article/Enable-auto-acceleration-for-your-SharePoint-Online-tenancy-74985ebf-39e1-4c59-a74a-dcdfd678ef83

######## 
#Configuration 
######## 
$domain             = "OGD.NL"                    #This should be your domain name in O365, and your UPN in Active Directory, for example: ogd.nl 
$driveLetter         = "X:"                         #This is the driveletter you'd like to use for OneDrive, for example: Z: 
$driveLabel         = "OGD"                     #If you enter a name here, the script will attempt to label the drive with this value 
$O365CustomerName    = "ogd"                        #This should be the name of your tenant (example, ogd as in ogd.onmicrosoft.com) 
$logfile            = ($env:APPDATA + "\OneDriveMapper.log")    #Logfile to log to 
$dontMapO4B         = $False                     #If you're only using Sharepoint Online mappings (see below), set this to True to keep the script from mapping the user's O4B (the user does still need to have a O4B license!)
$debugmode          = $False                      #Set to $True for debugging purposes. You'll be able to see the script navigate in Internet Explorer 
$lookupUPNbySAM     = $True                     #Look up the user's UPN by the SAMAccountName, use this if your UPN doesn't match your SamAccountName or if you have multiple domains 
$forceUserName      = ""                        #if anything is entered here, there will be no UPN lookup and the domain will be ignored. This is useful for machines that aren't domain joined. 
$forcePassword      = ""                        #if anything is entered here, the user won't be prompted for a password. This function is not recommended, as your password could be stolen from this file 
$autoProtectedMode  = $True                     #Automatically temporarily disable IE Protected Mode if it is enabled. ProtectedMode has to be disabled for the script to function 
$adfsWaitTime       = 10                         #Amount of seconds to allow for ADFS redirects, if set too low, the script may fail while just waiting for a slow ADFS redirect, this is because the IE object will report being ready even though it is not.  Set to 0 if not using ADFS. 
$libraryName        = "Documents"               #leave this default, unless you wish to map a non-default library you've created 
$restart_explorer   = $False                    #Set to true if drives are always invisible after the script runs, this will restart explorer.exe after mapping the drive 
$autoKillIE         = $True                     #Kill any running Internet Explorer processes prior to running the script to prevent security errors when mapping 
$abortIfNoAdfs      = $False                    #If set to True, will stop the script if no ADFS server has been detected during login
$displayErrors      = $False                    #show errors to user in visual popups
$buttonText         = "Login"                   #Text of the button on the password input popup box
$adfsLoginInput     = "userNameInput"           #Only modify this if you have a customized (skinned) ADFS implementation
$adfsPwdInput       = "passwordInput"
$adfsButton         = "submitButton"
$urlOpenAfter       = ""                        #This URL will be opened by the script after running if you configure it
$sharepointMappings = @()
$sharepointMappings += "https://ogd.sharepoint.com/site1/documentsLibrary,ExampleLabel,Y:"
#for each sharepoint site you wish to map 3 comma seperated values are required, the url to the library, the desired drive label, and the driveletter
#if you wish to add more, copy the example as you see above, if you don't wish to map any sharepoint sites, simply remove the line or clear everything between the quotes


######## 
#Required resources 
######## 
$mapresult = $False 
$protectedModeValues = @{} 
$privateSuffix = "-my" 
$script:errorsForUser = ""
$o365loginURL = "https://login.microsoftonline.com"

ac $logfile "-----$(Get-Date) OneDriveMapper V2.28 - $($env:USERNAME) on $($env:COMPUTERNAME) Session log-----" 
 
Write-Host "One moment please, your drive(s) are connecting..." 

$domain = $domain.ToLower() 
$O365CustomerName = $O365CustomerName.ToLower() 
$forceUserName = $forceUserName.ToLower() 
$finalURLs = @()
$finalURLs += "https://portal.office.com"
$finalURLs += "https://outlook.office365.com"
$finalURLs += "https://outlook.office.com"
$finalURLs += "https://$($O365CustomerName)-my.sharepoint.com"
$finalURLs += "https://$($O365CustomerName).sharepoint.com"

if($sharepointMappings[0] -eq "https://ogd.sharepoint.com/site1/documentsLibrary,ExampleLabel,Y:"){
    $sharepointMappings = @()
}

function checkIfAtO365URL{
    param(
        [String]$url,
        [Array]$finalURLs
    )
    foreach($item in $finalURLs){
        if($url.StartsWith($item)){
            return $True
        }
    }
    return $False
}

#region basicFunctions
function lookupUPN{ 
    try{ 
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
        $objSearcher.SearchRoot = $objDomain 
        $objSearcher.Filter = “(&(objectCategory=User)(SAMAccountName=$Env:USERNAME))” 
        $objSearcher.SearchScope = “Subtree” 
        $objSearcher.PropertiesToLoad.Add(“userprincipalname”) | Out-Null 
        $results = $objSearcher.FindAll() 
        return $results[0].Properties.userprincipalname 
    }catch{ 
        ac $logfile "Failed to lookup username, active directory connection failed, please disable lookupUPN" 
        $script:errorsForUser += "Could not find your username, are you connected to your network?`n"
        abort_OM 
    } 
} 
 
function CustomInputBox([string] $title, [string] $message)  
{ 
    if($forcePassword.Length -gt 2) { 
        return $forcePassword 
    } 
    $objBalloon = New-Object System.Windows.Forms.NotifyIcon  
    $objBalloon.BalloonTipIcon = "Info" 
    $objBalloon.BalloonTipTitle = "OneDriveMapper"  
    $objBalloon.BalloonTipText = "OneDriveMapper - www.liebensraum.nu" 
    $objBalloon.Visible = $True  
    $objBalloon.ShowBalloonTip(10000) 
 
    $userForm = New-Object 'System.Windows.Forms.Form' 
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState' 
    $Form_StateCorrection_Load= 
    { 
        $userForm.WindowState = $InitialFormWindowState 
    } 
 
    $userForm.Text = "$title" 
    $userForm.Size = New-Object System.Drawing.Size(350,200) 
    $userForm.StartPosition = "CenterScreen" 
    $userForm.AutoSize = $False 
    $userForm.MinimizeBox = $False 
    $userForm.MaximizeBox = $False 
    $userForm.SizeGripStyle= "Hide" 
    $userForm.WindowState = "Normal" 
    $userForm.FormBorderStyle="Fixed3D" 
    $userForm.KeyPreview = $True 
    $userForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$userForm.Close()}})   
    $OKButton = New-Object System.Windows.Forms.Button 
    $OKButton.Location = New-Object System.Drawing.Size(105,110) 
    $OKButton.Size = New-Object System.Drawing.Size(95,23) 
    $OKButton.Text = $buttonText 
    $OKButton.Add_Click({$userForm.Close()}) 
    $userForm.Controls.Add($OKButton) 
    $userLabel = New-Object System.Windows.Forms.Label 
    $userLabel.Location = New-Object System.Drawing.Size(10,20) 
    $userLabel.Size = New-Object System.Drawing.Size(300,50) 
    $userLabel.Text = "$message" 
    $userForm.Controls.Add($userLabel)  
    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.UseSystemPasswordChar = $True 
    $objTextBox.Location = New-Object System.Drawing.Size(70,75) 
    $objTextBox.Size = New-Object System.Drawing.Size(180,20) 
    $userForm.Controls.Add($objTextBox)  
    $userForm.Topmost = $True 
    $userForm.TopLevel = $True 
    $userForm.ShowIcon = $True 
    $userForm.Add_Shown({$userForm.Activate();$objTextBox.focus()}) 
    $InitialFormWindowState = $userForm.WindowState 
    $userForm.add_Load($Form_StateCorrection_Load) 
    [void] $userForm.ShowDialog() 
    return $objTextBox.Text 
} 
 
function labelDrive{ 
    Param( 
    [String]$lD_DriveLetter, 
    [String]$lD_MapURL, 
    [String]$lD_DriveLabel 
    ) 
 
    #try to set the drive label 
    if($lD_DriveLabel.Length -gt 0){ 
        ac $logfile "A drive label has been specified, attempting to set the label for $($lD_DriveLetter)" 
        try{ 
            $regURL = $lD_MapURL.Replace("\","#") 
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)" 
            $Null = New-Item -Path $path -Value "default value" –Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            $regURL = $regURL.Replace("DavWWWRoot#","") 
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)" 
            $Null = New-Item -Path $path -Value "default value" –Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            ac $logfile "Label has been set to $($lD_DriveLabel)" 
 
        }catch{ 
            ac $logfile "Failed to set the drive label registry keys" 
            ac $logfile $error[0] 
        } 
 
    } 
} 

function fixElevationVisibility{
    Param( 
    [String]$MD_DriveLetter, 
    [String]$MD_MapURL, 
    [String]$MD_DriveLabel 
    ) 
    if($MD_DriveLetter.Length -eq 2){
        $MD_DriveLetter = $MD_DriveLetter.SubString(0,1)
    }else{
        return $False
    }
    $path = "HKCU:\Network\$($MD_DriveLetter)"
    if(Test-Path $path){
        ac $logfile "$path key found, no further action required to make the driveletter visible for elevated users"
        return $True
    }else{
        $Null = New-Item -Path $path –Force -ErrorAction SilentlyContinue
        $Null = New-ItemProperty -Path $path -Name "ConnectionType" -PropertyType DWORD -Value 1 -ErrorAction SilentlyContinue
        $Null = New-ItemProperty -Path $path -Name "DeferFlags" -PropertyType DWORD -Value 4 -ErrorAction SilentlyContinue
        $Null = New-ItemProperty -Path $path -Name "ProviderName" -PropertyType String -Value "Web Client Network" -ErrorAction SilentlyContinue
        $Null = New-ItemProperty -Path $path -Name "ProviderType" -PropertyType DWORD -Value 3014656 -ErrorAction SilentlyContinue
        $Null = New-ItemProperty -Path $path -Name "RemotePath" -PropertyType String -Value $MD_MapURL -ErrorAction SilentlyContinue
        $Null = New-ItemProperty -Path $path -Name "UserName" -PropertyType String -Value $Env:USERNAME -ErrorAction SilentlyContinue
        if($restart_explorer){
            ac $logfile "$path key not found but created"
        }else{
            ac $logfile "$path key not found but created, BUT drive will not be visible until explorer.exe is restarted, set restart_explorer to True"
        }
        return $True
    }
        
}

function MapDrive{ 
    Param( 
    [String]$MD_DriveLetter, 
    [String]$MD_MapURL, 
    [String]$MD_DriveLabel 
    ) 
    ac $logfile "Mapping target: $($MD_MapURL)`n" 
    $del = NET USE $MD_DriveLetter /DELETE 2>&1 
    $out = NET USE $MD_DriveLetter $MD_MapURL /PERSISTENT:YES 2>&1 
    if($LASTEXITCODE -ne 0){ 
        if((Get-Service -Name WebClient).Status -ne "Running"){ 
            ac $logfile "CRITICAL ERROR: OneDriveMapper detected that the WebClient service was not started, please ensure this service is always running!`n" 
            $script:errorsForUser += "$MD_DriveLetter could not be mapped because the WebClient service is not running`n"
        } 
        ac $logfile "Failed to map $($MD_DriveLetter) to $($MD_MapURL), error: $($LASTEXITCODE) $($out)`n" 
        $script:errorsForUser += "$MD_DriveLetter could not be mapped because of error $($LASTEXITCODE) $($out)`n"
        return $False 
    } 
    if([System.IO.Directory]::Exists($MD_DriveLetter)){ 
        #set drive label 
        $Null = labelDrive $MD_DriveLetter $MD_MapURL $MD_DriveLabel
        #if admin, check registry keys for drive
        if($isElevated){
            $Null = fixElevationVisibility $MD_DriveLetter $MD_MapURL $MD_DriveLabel
        }
        ac $logfile "$($MD_DriveLetter) mapped successfully`n" 
        if($restart_explorer){ 
            ac $logfile "Restarting Explorer.exe to make the drive visible" 
            #kill all running explorer instances of this user 
            $explorerStatus = Get-ProcessWithOwner explorer 
            if($explorerStatus -eq 0){ 
                ac $logfile "WARNING: no instances of Explorer running yet, at least one should be running" 
            }elseif($explorerStatus -eq -1){ 
                ac $logfile "ERROR Checking status of Explorer.exe: unable to query WMI" 
            }else{ 
                ac $logfile "Detected running Explorer processes, attempting to shut them down..." 
                foreach($Process in $explorerStatus){ 
                    try{ 
                        Stop-Process $Process.handle | Out-Null 
                        ac $logfile "Stopped process with handle $($Process.handle)" 
                    }catch{ 
                        ac $logfile "Failed to kill process with handle $($Process.handle)" 
                    } 
                } 
            } 
        } 
        return $True 
    }else{ 
        if($LASTEXITCODE -eq 0){ 
            ac $logfile "failed to contact $($MD_DriveLetter) after mapping it to $($MD_MapURL), check if the URL is valid" 
            ac $logfile $error[0] 
        } 
        return $False 
    } 
} 
 
function revertProtectedMode(){ 
    ac $logfile "autoProtectedMode is set to True, reverting to old settings" 
    try{ 
        for($i=0; $i -lt 5; $i++){ 
            if($protectedModeValues[$i] -ne $Null){ 
                ac $logfile "Setting zone $i back to $($protectedModeValues[$i])" 
                Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value $protectedModeValues[$i] -Type Dword -ErrorAction SilentlyContinue 
            } 
        } 
    } 
    catch{ 
        ac $logfile "Failed to modify registry keys to change ProtectedMode back to the original settings" 
        ac $logfile $error[0] 
    } 
} 

function abort_OM{ 
    #find and kill all active COM objects for IE
    $ie.Quit() | Out-Null
    $shellapp = New-Object -ComObject "Shell.Application"
    $ShellWindows = $shellapp.Windows()
    for ($i = 0; $i -lt $ShellWindows.Count; $i++)
    {
      if ($ShellWindows.Item($i).FullName -like "*iexplore.exe")
      {
        $del = $ShellWindows.Item($i)
        $Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($del)  2>&1 
      }
    }
    $Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shellapp) 
    if($autoProtectedMode){ 
        revertProtectedMode 
    } 
    ac $logfile "OnedriveMapper has finished running"
    Write-Host "OnedriveMapper has finished running"
    if($urlOpenAfter){Start-Process iexplore.exe $urlOpenAfter}
    if($displayErrors){
        if($errorsForUser){ $OUTPUT= [System.Windows.Forms.MessageBox]::Show($errorsForUser, "Onedrivemapper Error" , 0) }
    }
    Exit 
} 
 
function askForPassword{ 
    do{ 
        $askAttempts++ 
        ac $logfile "asking user for password`n" 
        try{ 
            $password = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter the password for $($userUPN.ToLower()) to access $($driveLetter)" 
        }catch{ 
            ac $logfile "failed to display a password input box, exiting`n" 
            abort_OM              
        } 
    } 
    until($password.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 3) { 
        ac $logfile "user refused to enter a password, exiting`n" 
        $script:errorsForUser += "You did not enter a password, script cannot continue`n"
        abort_OM 
    }else{ 
        return $password 
    } 
} 
 
function Get-ProcessWithOwner { 
    param( 
        [parameter(mandatory=$true,position=0)]$ProcessName 
    ) 
    $ComputerName=$env:COMPUTERNAME 
    $UserName=$env:USERNAME 
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($(New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$('ProcessName','UserName','Domain','ComputerName','handle')))) 
    try { 
        $Processes = Get-wmiobject -Class Win32_Process -ComputerName $ComputerName -Filter "name LIKE '$ProcessName%'" 
    } catch { 
        return -1 
    } 
    if ($Processes -ne $null) { 
        $OwnedProcesses = @() 
        foreach ($Process in $Processes) { 
            if($Process.GetOwner().User -eq $UserName){ 
                $Process |  
                Add-Member -MemberType NoteProperty -Name 'Domain' -Value $($Process.getowner().domain) 
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $ComputerName  
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'UserName' -Value $($Process.GetOwner().User)  
                $Process |  
                Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers 
                $OwnedProcesses += $Process 
            } 
        } 
        return $OwnedProcesses 
    } else { 
        return 0 
    } 
 
} 
#endregion

function addMapping(){
    Param(
    [String]$driveLetter,
    [String]$url,
    [String]$label
    )
    $mapping = "" | Select-Object driveLetter, URL, Label, alreadyMapped
    $mapping.driveLetter = $driveLetter
    $mapping.url = $url
    $mapping.label = $label
    $mapping.alreadyMapped = $False
    ac $logfile "Adding to mapping list: $driveLetter ($url)"
    return $mapping
}

#this function checks if a given drivemapper is properly mapped to the given location, returns true if it is, otherwise false
function checkIfLetterIsMapped(){
    Param(
    [String]$driveLetter,
    [String]$url
    )
    if([System.IO.Directory]::Exists($driveLetter)){ 
        #check if mapped path is to at least the personal folder on Onedrive for Business, username detection would require a full login and slow things down
        #Ignore DavWWWRoot, as this does not consistently appear in the actual URL
        [String]$mapped_URL = (Get-PSDrive $driveLetter.Substring(0,1)).DisplayRoot.Replace("DavWWWRoot\","")
        [String]$url = $url.Replace("DavWWWRoot\","")
        if($mapped_URL.StartsWith($url)){
            ac $logfile "the mapped url for $driveLetter ($mapped_URL) matches the expected URL of $url, no need to remap"
            return $True
        }else{
            ac $logfile "the mapped url for $driveLetter ($mapped_URL) does not match the expected partial URL of $url"
            return $False
        } 
    }else{
        ac $logfile "$driveLetter is not yet mapped"
        return $False
    }
}

#region loginFunction
function login(){ 
    ac $logfile "Login attempt at Office 365 signin page" 
    #click to open up the login menu 
    do {sleep -m 100} until (-not ($ie.Busy))  
    if($ie.document.GetElementById("_link").tagName -ne $Null){ 
       $ie.document.GetElementById("_link").click()  
       ac $logfile "Found sign in elements type 1 on Office 365 login page, proceeding" 
    }elseif($ie.document.GetElementById("use_another_account").tagName -ne $Null){ 
       $ie.document.GetElementById("use_another_account").click() 
       ac $logfile "Found sign in elements type 2 on Office 365 login page, proceeding" 
    }elseif($ie.document.GetElementById("use_another_account_link").tagName -ne $Null){ 
       $ie.document.GetElementById("use_another_account_link").click() 
       ac $logfile "Found sign in elements type 3 on Office 365 login page, proceeding" 
    }elseif($ie.document.GetElementById("_use_another_account_link").tagName -ne $Null){ 
       $ie.document.GetElementById("_use_another_account_link").click() 
       ac $logfile "Found sign in elements type 4 on Office 365 login page, proceeding" 
    }elseif($ie.document.GetElementById("cred_keep_me_signed_in_checkbox").tagName -ne $Null){ 
       ac $logfile "Found sign in elements type 5 on Office 365 login page, proceeding" 
    }else{ 
       ac $logfile "Script was unable to find browser controls on the login page and cannot continue, please check your safe-sites or verify these elements are present" 
       $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
       abort_OM 
    } 
    do {sleep -m 100} until (-not ($ie.Busy))  
 
 
    #attempt to trigger redirect to detect if we're using ADFS automatically 
    try{ 
        ac $logfile "attempting to trigger a redirect to ADFS" 
        $ie.document.GetElementById("cred_keep_me_signed_in_checkbox").click() 
        $ie.document.GetElementById("cred_userid_inputtext").value = $userUPN 
        do {sleep -m 100} until (-not ($ie.Busy))  
        $ie.document.GetElementById("cred_password_inputtext").click() 
        do {sleep -m 100} until (-not ($ie.Busy))  
    }catch{ 
        ac $logfile "Failed to find the correct controls at $($ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script`n" 
        $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
        abort_OM  
    } 
 
    sleep -s 2 
    $redirWaited = 0 
    while($True){ 
        sleep -m 500 
        try{
            $found_Splitter = $ie.document.GetElementById("aad_account_tile_link").tagName
        }catch{
            $found_Splitter = $Null
        }
        #Select business account if the option is presented 
        if($found_Splitter -ne $Null){ 
            $ie.document.GetElementById("aad_account_tile_link").click() 
            ac $logfile "Login splitter detected, your account is both known as a personal and business account, selecting business account.." 
            sleep -s 2 
            $redirWaited += 2
        } 
        #check if the COM object is healthy, otherwise we're running into issues 
        if($ie.HWND -eq $null){ 
            ac $logfile "ERROR: the browser object was Nulled during login, this means IE ProtectedMode or other security settings are blocking the script." 
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            abort_OM 
        } 

        #If ADFS automatically signs us on, this will trigger
        if((checkIfAtO365URL -url $ie.LocationURL -finalURLs $finalURLs)){
            $useADFS = $True
            break            
        }

        #this is the ADFS login control ID, modify this if you have a custom IdP
        try{
            $found_ADFSControl = $ie.document.GetElementById($adfsLoginInput).tagName
        }catch{
            $found_ADFSControl = $Null
            ac $logfile "ADFS userNameInput element not found: $($Error[0]) with method 1"
        }
        #try alternative method for selecting the ID 
        if($found_ADFSControl.Length -lt 1){
            try{
                $found_ADFSControl = $ie.Document.IHTMLDocument3_getElementById($adfsLoginInput).tagName
            }catch{
                $found_ADFSControl = $Null
                ac $logfile "ADFS userNameInput element not found: $($Error[0]) with method 2"
            }
        }
        $redirWaited += 0.5 
        #found ADFS control
        if($found_ADFSControl){
            ac $logfile "ADFS Control found, we were redirected to: $($ie.LocationURL)" 
            $useADFS = $True
            break            
        } 

        if($redirWaited -ge $adfsWaitTime){ 
            ac $logfile "waited for more than $adfsWaitTime to get redirected to ADFS, checking if we were properly redirected or attempting normal signin" 
            $useADFS = $False    
            break 
        } 
    }     

    #if not using ADFS, sign in 
    if($useADFS -eq $False){ 
        if($abortIfNoAdfs){
            ac $logfile "abortIfNoAdfs was set to true, ADFS was not detected, script is exiting"
            $script:errorsForUser += "Onedrivemapper cannot login because ADFS is not available`n"
            abort_OM
        }
        if((checkIfAtO365URL -url $ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)" 
            return $True             
        }
        try{ 
            $ie.document.GetElementById("cred_password_inputtext").value = askForPassword 
            $ie.document.GetElementById("cred_sign_in_button").click() 
            do {sleep -m 100} until (-not ($ie.Busy))
        }catch{ 
            ac $logfile "Failed to find the correct controls at $($ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script`n" 
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            abort_OM  
        } 
    }else{ 
        #check if logged in now automatically after ADFS redirect 
        if((checkIfAtO365URL -url $ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)" 
            return $True 
        } 
    } 
 
    #Not logged in automatically, so ADFS requires us to sign in 
    do {sleep -m 100} until (-not ($ie.Busy))
 
    #Check if we arrived at a 404, or an actual page 
    if($ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
        ac $logfile "We received a 404 error after our signin attempt, this script cannot continue" 
        $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
        abort_OM          
    } 

    #check if logged in now 
    if((checkIfAtO365URL -url $ie.LocationURL -finalURLs $finalURLs)){
        #we've been logged in, we can abort the login function 
        ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)" 
        return $True 
    }else{ 
        if($useADFS){ 
            ac $logfile "ADFS did not automatically sign us on, attempting to enter credentials at $($ie.LocationURL)" 
            try{ 
                $ie.document.GetElementById($adfsLoginInput).value = $userUPN 
                $ie.document.GetElementById($adfsPwdInput).value = askForPassword 
                $ie.document.GetElementById($adfsButton).click() 
                do {sleep -m 100} until (-not ($ie.Busy))  
                sleep -s 1 
                do {sleep -m 100} until (-not ($ie.Busy))   
            }catch{ 
                ac $logfile "Failed to find the correct controls at $($ie.LocationURL) using method 1 to log in by script, will try method 2`n" 
                $tryMethod2 = $True
            } 
            if($tryMethod2 -eq $True){
                try{ 
                    $ie.document.IHTMLDocument3_getElementById($adfsLoginInput).value = $userUPN 
                    $ie.document.IHTMLDocument3_getElementById($adfsPwdInput).value = askForPassword 
                    $ie.document.IHTMLDocument3_getElementById($adfsButton).click() 
                    do {sleep -m 100} until (-not ($ie.Busy))  
                    sleep -s 1 
                    do {sleep -m 100} until (-not ($ie.Busy))   
                }catch{ 
                    ac $logfile "Failed to find the correct controls at $($ie.LocationURL) using method 2 to log in by script, check your browser and proxy settings or modify this script to match your ADFS form`n" 
                    $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
                    abort_OM 
                } 
            }
            do {sleep -m 100} until (-not ($ie.Busy))   
            #check if logged in now         
            if((checkIfAtO365URL -url $ie.LocationURL -finalURLs $finalURLs)){
                #we've been logged in, we can abort the login function 
                ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)" 
                return $True 
            }else{ 
                ac $logfile "We attempted to login with ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)" 
                $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
                abort_OM 
            } 
        }else{ 
            ac $logfile "We attempted to login without using ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)" 
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            abort_OM 
        } 
    } 
} 
#endregion


 
#get user login 
if($lookupUPNbySAM){ 
    ac $logfile "lookupUPNbySAM is set to True -> Using UPNlookup by SAMAccountName feature`n" 
    $userUPN = lookupUPN 
}else{ 
    $userUPN = ([Environment]::UserName)+"@"+$domain 
    ac $logfile "lookupUPNbySAM is set to False -> Using $userUPN from the currently logged in username + $domain`n" 
} 
if($forceUserName.Length -gt 2){ 
    ac $logfile "A username was already specified in the script configuration: $($forceUserName)`n" 
    $userUPN = $forceUserName 
} 

#region flightChecks

#check if the script is running elevated
If (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){   
    $isElevated = $True
    ac $logfile "Script elevation level: Administrator"
}else{
    $isElevated = $False
    ac $logfile "Script elevation level: User"
}
#Check if Office 365 libraries have been installed 
if([System.IO.File]::Exists("$(${env:ProgramFiles(x86)})\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll") -eq $False){ 
    ac $logfile "Possible critical error: Microsoft Office installation not detected, script may fail" 
} 
 
 
#Check if Zone Configuration is on a per machine or per user basis, then check the zones 
$privateZoneFound = $False
$publicZoneFound = $False
$BaseKeypath = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" 
try{ 
    $IEMO = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "Security HKLM only" -ErrorAction Stop | Select-Object 'Security HKLM only' 
}catch{ 
    ac $logfile "NOTICE: $($BaseKeypath)\Security HKLM only not found in registry, your zone configuration could be set on both levels" 
} 
if($IEMO.'Security HKLM only' -eq 1){ 
    ac $logfile "NOTICE: $($BaseKeypath)\Security HKLM only found in registry and set to 1, your zone configuration is set on a machine level"    
}else{ 
    #Check if sharepoint subtenant is in safe sites list of the user 
    $BaseKeypath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" 
    $BaseKeypath2 = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" 
    $zone = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https 
    if($zone -eq $Null){ 
        $zone = Get-ItemProperty -Path "$($BaseKeypath2)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https   
    }     
    if($zone.https -eq 2){ 
        ac $logfile "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level"  
        $privateZoneFound = $True
    }
    #Check if sharepoint tenant is in safe sites list of the user 
    $BaseKeypath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" 
    $BaseKeypath2 = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" 
    $zone = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https 
    if($zone -eq $Null){ 
        $zone = Get-ItemProperty -Path "$($BaseKeypath2)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https   
    }     
    if($zone.https -eq 2){ 
        ac $logfile "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level"  
        $publicZoneFound = $True
    }
} 
#Check if sharepoint subtenant is in safe sites list of the machine 
$BaseKeypath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" 
$BaseKeypath2 = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" 
$zone = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https 
if($zone -eq $Null){ 
    $zone = Get-ItemProperty -Path "$($BaseKeypath2)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https   
}     
if($zone.https -eq 2){ 
    ac $logfile "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level"  
    $privateZoneFound = $True
}
if($zoneFound -eq $False){
    ac $logfile "Possible critical error: $($O365CustomerName)$($privateSuffix).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail"
}
 
#Check if sharepoint tenant is in safe sites list of the machine 
$BaseKeypath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" 
$BaseKeypath2 = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" 
$zone = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https 
if($zone -eq $Null){ 
    $zone = Get-ItemProperty -Path "$($BaseKeypath2)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https   
}     
if($zone.https -eq 2){ 
    ac $logfile "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level"  
    $publicZoneFound = $True
}
if($publicZoneFound -eq $False){
    ac $logfile "Possible critical error: $($O365CustomerName).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail"
}
if($privateZoneFound -eq $False){
    ac $logfile "Possible critical error: $($O365CustomerName)$($privateSuffix).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail"
}
 
#Check if IE FirstRun is disabled 
$BaseKeypath = "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main" 
try{ 
    $IEFR = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "DisableFirstRunCustomize" -ErrorAction Stop | Select-Object DisableFirstRunCustomize 
}catch{ 
    ac $logfile "WARNING: $($BaseKeypath)\DisableFirstRunCustomize not found in registry, if script hangs this may be due to the First Run popup in IE" 
} 
if($IEFR.DisableFirstRunCustomize -ne 1){ 
    ac $logfile "Possible error: $($BaseKeypath)\DisableFirstRunCustomize not set"    
} 
 
 
#Check if WebDav file locking is enabled 
$BaseKeypath = "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" 
try{ 
    $wdlocking = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "SupportLocking" -ErrorAction Stop | Select-Object SupportLocking 
}catch{ 
    ac $logfile "WARNING: HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters registry location not accessible" 
} 
if($wdlocking.SupportLocking -ne 0){ 
    ac $logfile "WARNING: WebDav File Locking support is enabled, this could cause files to become locked in your OneDrive"    
} 

#check if any zones are configured with Protected Mode through group policy (which we can't modify) 
$BaseKeypath = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" 
for($i=0; $i -lt 5; $i++){ 
    $curr = Get-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue | select -exp 2500 
    if($curr -ne $Null -and $curr -ne 3){ 
        ac $logfile "WARNING: IE Zone $i protectedmode is enabled through group policy, autoprotectedmode cannot disable it. This will likely cause the script to fail." 
    }
} 

#endregion
 
#translate to URLs 
$userURL = ($userUPN.replace(".","_")).replace("@","_").ToLower() 
$mapURL = ("\\"+$O365CustomerName+$privateSuffix+".sharepoint.com@SSL\DavWWWRoot\personal\"+$userURL+"\"+$libraryName) 
$mapURLpersonal = ("\\"+$O365CustomerName+"-my.sharepoint.com@SSL\DavWWWRoot\personal\") 
$baseURL = ("https://"+$O365CustomerName+$privateSuffix+".sharepoint.com") 

$desiredMappings = @() #array with mappings to be made

#add the O4B mapping first, with an incorrect URL that will be updated later on because we haven't logged in yet and can't be sure of the URL
if($dontMapO4B){
    ac $logfile "Not mapping O4B because dontMapO4B is set to True"
}else{
    $desiredMappings += addMapping -driveLetter $driveLetter -url $mapURLpersonal -label $driveLabel
}

$WebAssemblyloaded = $True
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")
if(-NOT [appdomain]::currentdomain.getassemblies() -match "System.Web"){
    ac $logfile "Error loading System.Web library to decode sharepoint URL's, mapped sharepoint URL's may become read-only"   
    $WebAssemblyloaded = $False
}

#add any desired Sharepoint Mappings
$sharepointMappings | % {
    $data = $_.Split(",")
    if($data[0] -and $data[1] -and $data[2]){
        if($WebAssemblyloaded){
            $add = [System.Web.HttpUtility]::UrlDecode($data[0])
        }else{
            $add = $data[0]
        }
        $add = $add.Replace("https://","\\") 
        $add = $add.Replace("sharepoint.com/","sharepoint.com@SSL\") 
        $add = $add.Replace("/","\") 
        $desiredMappings += addMapping -driveLetter $data[2] -url $add -label $data[1]    
    }
}

$continue = $False
$countMapping = 0
#check if any of the mappings we should make is already mapped and update the corresponding property
$desiredMappings | % {
    if((checkIfLetterIsMapped -driveLetter $_.driveletter -url $_.url)){
        $desiredMappings[$countMapping].alreadyMapped = $True
        #reset the label for this drive (it might have changed)
        labelDrive $_.driveLetter ((Get-PSDrive $_.driveLetter.Substring(0,1)).DisplayRoot) $_.label
    }
    $countMapping++
}
 
if(@($desiredMappings | where-object{$_.alreadyMapped -eq $False}).Count -le 0){
    ac $logfile "no unmapped or incorrectly mapped drives detected"
    Write-Host "no unmapped or incorrectly mapped drives detected"
    abort_OM    
}

#load windows libraries to display things to the user 
try{ 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
}catch{ 
    ac $logfile "Error loading windows forms libraries, script will not be able to display a password input box" 
} 
 
ac $logfile "Base URL: $($baseURL) `n" 

#Start IE and stop it once to make sure IE sets default registry keys 
if($autoKillIE){ 
    #start invisible IE instance 
    $script:ie = new-object -com InternetExplorer.Application 
    $ie.visible = $debugmode 
    sleep 2 
 
    #kill all running IE instances of this user 
    $ieStatus = Get-ProcessWithOwner iexplore 
    if($ieStatus -eq 0){ 
        ac $logfile "WARNING: no instances of Internet Explorer running yet, at least one should be running" 
    }elseif($ieStatus -eq -1){ 
        ac $logfile "ERROR Checking status of iexplore.exe: unable to query WMI" 
    }else{ 
        ac $logfile "autoKillIE enabled, stopping IE processes" 
        foreach($Process in $ieStatus){ 
                Stop-Process $Process.handle -ErrorAction SilentlyContinue
                ac $logfile "Stopped process with handle $($Process.handle)"
        } 
    } 
}else{ 
    ac $logfile "ERROR: autoKillIE disabled, IE processes not stopped. This may cause the script to fail for users with a clean/new profile" 
} 

if($autoProtectedMode){ 
    ac $logfile "autoProtectedMode is set to True, disabling ProtectedMode temporarily" 
    $BaseKeypath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" 
     
    #store old values and change new ones 
    try{ 
        for($i=0; $i -lt 5; $i++){ 
            $curr = Get-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue| select -exp 2500 
            if($curr -ne $Null){ 
                $protectedModeValues[$i] = $curr 
                ac $logfile "Zone $i was set to $curr, setting it to 3" 
            }else{
                $protectedModeValues[$i] = 0 
                ac $logfile "Zone $i was not yet set, setting it to 3" 
            }
            Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value "3" -Type Dword -ErrorAction Stop
        } 
    } 
    catch{ 
        ac $logfile "Failed to modify registry keys to autodisable ProtectedMode $($error[0])" 
    } 
}else{
    ac $logfile "autoProtectedMode is set to False, IE ProtectedMode will not be disabled temporarily"
}
 
#start invisible IE instance 
try{ 
    $script:ie = new-object -com InternetExplorer.Application -ErrorAction Stop
    $ie.visible = $debugmode 
}catch{ 
    ac $logfile "failed to start Internet Explorer COM Object, check user permissions or already running instances`n$($error[0])"  
    $errorsForUser += "Mapping cannot continue because we could not start the browser`n"
    abort_OM 
} 

#navigate to the base URL of the tenant's Sharepoint to check if it exists 
try{ 
    $ie.navigate("https://login.microsoftonline.com/logout.srf")
    do {sleep -m 100} until (-not ($ie.Busy))
    $ie.navigate($o365loginURL) 
    do {sleep -m 100} until (-not ($ie.Busy))  
}catch{ 
    ac $logfile "Failed to browse to the Office 365 Sign in page, this is a fatal error $($error[0])`n" 
    $errorsForUser += "Mapping cannot continue because we could not contact Office 365`n"
    abort_OM 
} 
 
#check if we got a 404 not found 
if($ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
    ac $logfile "Failed to browse to the Office 365 Sign in page, exiting script" 
    $errorsForUser += "Mapping cannot continue because we could not start the browser`n"
    abort_OM 
} 
 
#check if the COM object is healthy, otherwise we're running into issues 
if($ie.HWND -eq $null){ 
    ac $logfile "ERROR: attempt to navigate caused the IE scripting object to be nulled. This means your security settings are too high (1)." 
    $errorsForUser += "Mapping cannot continue because we could not start the browser`n"
    abort_OM 
} 
ac $logfile "current URL: $($ie.LocationURL)" 
 
#log in 
if($ie.LocationURL.StartsWith($baseURL)){ 
    ac $logfile "ERROR: You were already logged in, skipping login attempt, please note this may fail if you did not log in with a persistent cookie" 
}else{ 
    #Check and log if Explorer is running 
    $explorerStatus = Get-ProcessWithOwner explorer 
    if($explorerStatus -eq 0){ 
        ac $logfile "WARNING: no instances of Explorer running yet, expected at least one running" 
    }elseif($explorerStatus -eq -1){ 
        ac $logfile "ERROR Checking status of explorer.exe: unable to query WMI" 
    }else{ 
        ac $logfile "Detected running explorer process" 
    } 
    login 
    $ie.navigate($baseURL) 
    do {sleep -m 100} until (-not ($ie.Busy))  
    do {sleep -m 100} until ($ie.ReadyState -eq 4 -or $ie.ReadyState -eq 0)  
    Sleep -s 2
} 

ac $logfile "Attempting to retrieve the username by browsing to $baseURL..." 
$url = $ie.LocationURL 
$timeSpent = 0
#find username
if($dontMapO4B -eq $False){
    while($url.IndexOf("/personal/") -eq -1){
        Sleep -s 3
        $timeSpent++
        $ie.navigate($baseURL)
        do {sleep -m 100} until (-not ($ie.Busy))  
        do {sleep -m 100} until ($ie.ReadyState -eq 4 -or $ie.ReadyState -eq 0)  
        $url = $ie.LocationURL
        if($timeSpent -gt 10){
            ac $logfile "Failed to get the username from the URL for 30 seconds while at $url, aborting" 
            $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
            abort_OM 
        }
    }
    try{
        $start = $url.IndexOf("/personal/")+10 
        $end = $url.IndexOf("/",$start) 
        $userURL = $url.Substring($start,$end-$start) 
        $mapURL = $mapURLpersonal + $userURL + "\" + $libraryName 
    }catch{
        ac $logfile "Failed to get the username while at $url, aborting" 
        $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
        abort_OM 
    }
    $desiredMappings[0].url = $mapURL 
    ac $logfile "Detected user: $($userURL)`n"
} 

ac $logfile "Current location: $($ie.LocationURL)" 
if($sharepointMappings.Count -gt 0){
    ac $logfile "browsing to Sharepoint to validate existence and set a cookie"
    $data = $sharepointMappings[0].Split(",")
    if($data[0] -and $data[1] -and $data[2]){
        $data = $data[0] #URL to browse to
        $ie.navigate($data) #check the URL
        do {sleep -m 100} until (-not ($ie.Busy))
        sleep 1 
        do {sleep -m 100} until (-not ($ie.Busy))
        if($ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*" -or $ie.HWND -eq $null) { 
            ac $logfile "Failed to browse to Sharepoint URL $data.`n" 
        } 
        ac $logfile "Current location: $($ie.LocationURL)" 
    }
}
ac $logfile "Cookies generated, attempting to map drive(s)`n" 
$desiredMappings | where-object {$_.alreadyMapped -eq $False} | % {
    $mapresult = MapDrive $_.driveLetter $_.url $_.label 
}
     
abort_OM
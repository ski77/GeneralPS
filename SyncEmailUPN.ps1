<#  
.SYNOPSIS 
    Script that copies the primary emailaddress from proxyAddresses to the userPrincipalName attribute.  
    It runs in test mode and just logs the changes that would have bee done without any parameters.  
    It identifies an exchange user with the legacyExchangeDN-attribute.  
.PARAMETER Production 
    Runs the script in production mode and makes the actual changes. 
.NOTES 
    Author: Scott Croucher  
    Email: sc@carnnell.co.uk 
    The script are provided “AS IS” with no guarantees, no warranties, and they confer no rights.     
#>
param(
[parameter(Mandatory=$false)]
[switch]
$Production = $false
)
#Define variables
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$DateStamp = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
$Logfile = $LogFile = ($PSScriptRoot + "\ProxyUPNSync-" + $DateStamp + ".log")
Function LogWrite
{
Param ([string]$logstring)
Add-content $Logfile -value $logstring
Write-Host $logstring
}
    try
    {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch
    {
        throw "Module ActiveDirectory not Installed"
    }
 
#For each AD-user with a legacyExchangeDN, look up primary SMTP: in proxyAddresses
#and use that as the UPN
$CollObjects=Get-ADObject -LDAPFilter "(&(legacyExchangeDN=*)(objectClass=user))" -Properties ProxyAddresses,distinguishedName,userPrincipalName
 
            foreach ($object in $CollObjects)
            {
                $Addresses = ""
                $DN=""
                $UserPrincipalName=""
                $Addresses = $object.proxyAddresses
                $ProxyArray=""
                $DN=$object.distinguishedName
                    foreach ($Address In $Addresses)
                    {
                        $ProxyArray=($ProxyArray + "," + $Address)
                        If ($Address -cmatch "SMTP:")
                            {
                                $PrimarySMTP = $Address
                                $UserPrincipalName=$Address -replace ("SMTP:","")
                                    #Found the object validating UserPrincipalName
                                    If ($object.userPrincipalName -notmatch $UserPrincipalName) {
                                        #Run in production mode if the production switch has been used
                                        If ($Production) {
                                            LogWrite ($DN + ";" + $object.userPrincipalName + ";NEW:" + $UserPrincipalName)
                                            Set-ADObject -Identity $DN -Replace @{userPrincipalName = $UserPrincipalName}
                                        }
                                        #Runs in test mode if the production switch has not been used
                                        else {
                                        LogWrite ($DN + ";" + $object.userPrincipalName + ";NEW:" + $UserPrincipalName)
                                        Set-ADObject -Identity $DN -WhatIf -Replace @{userPrincipalName = $UserPrincipalName}
                                        }
                            }
                            else
                            {
                            Write-Host "Info: All users primary email addresses are matching their userPrincipalName"
 
                            }
                        }
                    }
            }
 
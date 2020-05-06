############################################################################################################################################
#
# Script Name: Office 365 License Assignment Tool
# Version: 1
# Author: Brian Shipman
# Contact: bshipman@brianshipman.com OR bshipman@go-planet.com
#
# Credits
#  - Brent Challis | http://powershell.com/cs/media/p/10883.aspx | Select-TextItem Function
#  - Ed Wilson |  http://blogs.technet.com/b/heyscriptingguy/archive/2009/09/01/hey-scripting-guy-september-1.aspx | Get-Filename Function
#
############################################################################################################################################

#===========================================================================================================================================
#
# Functions
#
#===========================================================================================================================================

function Select-TextItem { 
    
    PARAM(
        [Parameter(Mandatory=$true)] 
        $options,
        $displayProperty 
    )
     
    [int]$optionPrefix = 1 

    # Create menu list 

    foreach ($option in $options) {

        if ($displayProperty -eq $null) { 

            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option) 

        } else { 

            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option.$displayProperty) 
        }

        $optionPrefix++ 

    }
     
    Write-Host ("{0,3}: {1}" -f 0,"To cancel")  

    [int]$response = Read-Host "Enter Selection" 

    $val = $null 

    if ($response -gt 0 -and $response -le $options.Count) {

        $val = $options[$response-1] 
    
    }

    return $val 

}

Function Get-FileName($initialDirectory,$FileTitle) {

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.ShowHelp = $True
    $OpenFileDialog.Title = $FileTitle
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV Files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename

} #end function Get-FileName

#===========================================================================================================================================
#
# Connect to Office 365
#
#===========================================================================================================================================

Clear

Do {

    Do {

        If($CloudUsername -eq $Null -or $ConvertCloudPassword -eq $Null) {

            Write-Host @"
Welcome to the Office 365 License Assignment Tool.

Before we get started, please enter your Office 365 Administrative username & password.

"@

            $CloudUsername = Read-Host "Enter your Cloud Username"
            $CloudPassword = Read-Host -AsSecureString "Enter your Cloud Password"
            $ConvertCloudPassword = ConvertFrom-SecureString($CloudPassword)

        } Else {

            $CloudPassword = ConvertTo-SecureString($ConvertCloudPassword)
        
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $CloudUsername, $CloudPassword

            $Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection

            Import-PSSession $Session

            Connect-MsolService –Credential $cred

       }

    } While ($Cred -eq $Null)

    $FindPSSession = Get-PSSession

} While ($FindPSSession -eq $Null)

#===========================================================================================================================================
#
# Choose Type of License Assignment
#
#===========================================================================================================================================

Clear

Do {

    If ( $Type -eq $Null ) {
    
        $MenuTitle = "Type of License Assignment"
        $MenuMessage = "Choose what type of license assignment you would like to perform."
        $Single = New-Object System.Management.Automation.Host.ChoiceDescription "&Single",""
        $Bulk = New-Object System.Management.Automation.Host.ChoiceDescription "&Bulk",""

        $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Single,$Bulk)
        $MenuChoice = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

        Switch($MenuChoice) {

            0 { $Type = "Single" }

            1 { $Type = "Bulk" }

        }

    }

} While($Type -eq $Null)

#===========================================================================================================================================
#
# License Assignment
#
#===========================================================================================================================================

$TenantName = (Get-MsolCompanyInformation).DisplayName

Do {

    Clear
    
    #$Type = $Null
    #$UPN = $Null
    #$UserList = $Null

    Switch($Type) {

        Single {

            Do {
            
                If($UPN -eq $Null -or $UPN -eq "") {

                    Clear

                    If($UPN -eq "") { Write-Host -ForegroundColor Yellow "A UserPrincipalName is required for this script to continue."; write-Host "" }

                    $UPN = Read-Host "Enter the UserPrincipalName to assign an Office 365 license(s) to"

                }

            } While ($UPN -eq "")
            
            Do {

                Clear
                
                Write-Host "You are about to assign Office 365 license(s) to $UPN"

                $MenuTitle = ""
                $MenuMessage = "Is the above UserPrincipalName correct?"
                $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
                $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""

                $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)
                $MenuChoice = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

                Switch($MenuChoice) {

                    0 {
                    
                        $UPNCheck = Get-MsolUser -UserPrincipalName $UPN

                        If($UPNCheck.UserPrincipalName -ne $UPN) { Clear; Write-Host -ForegroundColor Red "The UserPrincipalName " -NoNewLine; Write-Host -ForegroundColor Gray $UPN -NoNewLine; Write-Host -ForegroundColor Red " you entered could not be located in the " -NoNewline; Write-Host -ForeGroundColor Gray $TenantName -NoNewline; Write-Host -ForegroundColor Red " Office 365 tenant."; $UPN = Read-Host "Please enter the UserPrincipalName again"; $MenuChoice = 1 }
                    
                    }

                    1 {
                
                        Clear; $UPN = Read-Host "Enter the UserPrincipalName to assign an Office 365 license(s) to"
                
                    }

                }
            } While ($MenuChoice -ne 0)

        }

        Bulk {

            Do {
            
                If($UserList -eq $Null -or $UserList -eq "") {

                    $UserList = Get-FileName -initialDirectory "C:\" -FileTitle "Choose a UserList.csv to bulk assign licenses."

                }

            } While ($UserList -eq "")
            
            Do {

                Clear
                
                Write-Host "You are about to assign Office 365 license(s) to a list of users: $UserList"

                $MenuTitle = ""
                $MenuMessage = "Is the above User List file path correct?"
                $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
                $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""

                $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)
                $MenuChoice = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

                Switch($MenuChoice) {

                    0 {}

                    1 {
                
                        Do {

                            $UserList = Get-FileName -initialDirectory $DataDirectory -FileTitle "Choose a new UserList.csv to bulk assign licenses."

                        } While ($UserList -eq "")
                
                    }

                }

            } While ($MenuChoice -ne 0)
            
        }

    }

    Do {

    $Disabled = @()
    $AssignLicenseArray = @()

        Do {

            Clear

    
            $AccountSkuId = $null
            $AccountSkuIdWithoutDomain = $null
    

            Write-Host "Choose a primary license to assign to these users:"

            $values = Get-MsolAccountSku 
            $val = Select-TextItem $values "AccountSkuId" 
            $val.AccountSkuId
    
            $AccountSkuId = $val.AccountSkuId

            $Pos = $AccountSkuId.IndexOf(":")

            $AccountSkuIdWithoutDomain = $AccountSkuId.Substring($pos+1)

            $AssignLicenseArray += $AccountSkuId

            $AssignLicenseString = [String]$AssignLicenseArray

            Clear

            If($AssignLicenseArray.Length -eq 1) { $PrimaryLicense = $AccountSkuId }

            $MenuTitle = "Office 365 License Assignment"
            $MenuMessage = "Would you like to disable any service plans with in the $AccountSkuIdWithoutDomain subscription?"
            $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
            $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""

            $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)
            $MenuChoice = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

            Switch($MenuChoice) {

                0 {

                    Do {

                        Clear

                        Write-Host "Choose a service plan to disable."

                        $values = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq $AccountSkuIdWithoutDomain}
                        $s = $values.ServiceStatus
                        $o = $s.ServicePlan.ServiceName
                        $val = Select-TextItem $o

                        If($val -ne $Null) { $Disabled += $val }

                        $MenuTitle = "Office 365 License Assignment"
                        $MenuMessage = "Would you like to choose another service plan to disable?"
                        $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
                        $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""

                        $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)
                        $AnotherDisable = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

                    } While($AnotherDisable -ne 1)
        
                }

            }

            Clear

            $MenuTitle = "Office 365 License Assignment"
            $MenuMessage = "Would you like to assign another license to this user?"
            $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
            $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""

            $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)
            $LicenseToAssign = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

            Switch($LicenseToAssign) {

                0 { $LicenseToAssign2 = $TRUE }

                1 { $LicenseToAssign2 = $FALSE; $DisabledPlans = $Disabled -join ', '; $AssignLicense = $AssignLicenseArray -join ', ' }

            }

        } While($LicenseToAssign2 -eq $TRUE)

    Clear

    Write-Host -ForegroundColor Cyan "Tenant:"$TenantName
    Write-Host ""
    Switch($Type) { Single { Write-Host -ForegroundColor Yellow "User Account:" $UPN } Bulk { Write-Host -ForegroundColor Yellow "User List CSV:" $UserList } }
    Write-Host ""
    Write-Host -ForegroundColor Green "Primary Subscription Plan"
    Write-Host "   $AssignLicense"
    Write-Host ""
    Write-Host -ForeGroundColor Green "Service Plans to disable"
    Write-Host "   $DisabledPlans"
    Write-Host ""

    $MenuTitle = ""
    $MenuMessage = "Confirm the license assignment above."
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Correct",""
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&Incorrect",""

    $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)
    $ConfirmLicenseChoice = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

    Switch($ConfirmLicenseChoice) {

        0 {

            Write-Host ""
            $ConfirmLicense = $TRUE

            #==============================================

            $UsageLocation = Read-Host "Enter the two letter country code for Usage Location"
            Write-Host ""

            $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $PrimaryLicense -DisabledPlans $Disabled

            Switch($Type) {

                Bulk {

                    Import-CSV $UserList | % {

                        Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation $UsageLocation
                        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AssignLicenseArray -LicenseOptions $LicenseOptions

                        Write-Host "Licenses: $AssignLicenseArray have been assigned to" $_.UserPrincipalName

                    }

                }

                Single {
    
                    Set-MsolUser -UserPrincipalName $UPN -UsageLocation $UsageLocation
                    Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $AssignLicenseArray -LicenseOptions $LicenseOptions

                    Write-Host "Licenses: $AssignLicenseArray have been assigned to $UPN"

                }

            }

        }

        1 { $ConfirmLicense = $FALSE }

    }

    } While($ConfirmLicense -eq $FALSE)

    Write-Host ""

    $MenuTitle = ""
    $MenuMessage = "Would you like to assign Office 365 Licenses to another user or user list?"
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""

    $MenuOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No)
    $MainLicenseMenuChoice = $host.ui.PromptForChoice($MenuTitle, $MenuMessage, $MenuOptions, 0)

} While($MainLicenseMenuChoice -ne 1)

$Session = Get-PSSession | Where-Object {$_.ComputerName -like "*outlook.com*"}
Remove-PSSession -Id $Session.Id
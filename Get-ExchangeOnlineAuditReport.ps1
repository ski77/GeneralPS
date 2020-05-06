<#
.SYNOPSIS
    This script is used for creating Exchange Online Auditing Report in HTML format.


.PREREQUISITIES
    Make sure you create a Credential Reposityry Generic Credentials in order to include it under the $TeanantCredentialKey
    parameter (the name of the credentials).

.NOTES  
    Version                   : 0.1
    Rights Required           : Global administrator within O365
    Authors                   : Guy Bachar, Yoav Barzilay
    Last Update               : 20-May-2015
    Twitter/Blog              : @GuyBachar, http://guybachar.wordpress.com
    Twitter/Blog              : @y0avb, y0av.me

.REFRENCES
    http://mikepfeiffer.net/2010/08/administrator-audit-log-reports-in-html-format-exchange-2010-sp1/
    https://technet.microsoft.com/en-us/library/jj150497%28v=exchg.150%29.aspx?f=255&MSPPError=-2147217396



.VERSION
    0.1 - Initial Version for connecting Online resources

#>

param(
[Parameter(Position=0, Mandatory=$True) ][ValidateNotNullorEmpty()][string] $TenantCredentialKey,
[Parameter(Position=1, Mandatory=$false)][ValidateNotNullorEmpty()][string] $To,
[Parameter(Position=2, Mandatory=$false)][ValidateNotNullorEmpty()][string] $From,
[Parameter(Position=3, Mandatory=$false)][ValidateNotNullorEmpty()][string] $SmtpServer,
[Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $ExchangeOnlineWithProxy
)

######################################################################
# Common functions
######################################################################
#
# Connect Exchange Online
#
Function Connect-ExchangeOnline ([System.Management.Automation.PSCredential] $credential)
{
    if ($ExchangeOnlineWithProxy.IsPresent)
    {
    $ProxySettings = New-PSSessionOption -ProxyAccessType IEConfig
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection -SessionOption $ProxySettings
    }
    else
    {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection
    }
    Import-PSSession $Session -AllowClobber
}

#
# Auditing Log Report
#
function New-AuditLogReport {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        #[Microsoft.Exchange.Management.SystemConfigurationTasks.AdminAuditLogEvent]
        $AuditLogEntry	
        )
	begin {
$css = @'
	<style type="text/css">
	body { font-family: Tahoma, Geneva, Verdana, sans-serif;}
	table {border-collapse: separate; background-color: #EEF9EC; border: 3px solid #103E69; caption-side: bottom;}
	td { border:1px solid #103E69; margin: 3px; padding: 3px; vertical-align: top; background: #EEF9EC; color: #000;font-size: 12px;}
	thead th {background: #903; color:#fefdcf; text-align: left; font-weight: bold; padding: 3px;border: 1px solid #990033;}
	th {border:1px solid #CC9933; padding: 3px;}
	tbody th:hover {background-color: #fefdcf;}
	th a:link, th a:visited {color:#903; font-weight: normal; text-decoration: none; border-bottom:1px dotted #c93;}
	caption {background: #903; color:#fcee9e; padding: 4px 0; text-align: center; width: 40%; font-weight: bold;}
	tbody td a:link {color: #903;}
	tbody td a:visited {color:#633;}
	tbody td a:hover {color:#000; text-decoration: none;
	}
	</style>
'@	
		$sb = New-Object System.Text.StringBuilder
		[void]$sb.AppendLine($css)
		[void]$sb.AppendLine("<table cellspacing='0'>")
		[void]$sb.AppendLine("<tr><td colspan='6'><strong>Exchange Online (Office 365) Administrator Audit Log Report for $((get-date).ToShortDateString())</strong></td></tr>")
		[void]$sb.AppendLine("<tr>")
		[void]$sb.AppendLine("<td><strong>Caller</strong></td>")
		[void]$sb.AppendLine("<td><strong>Run Date</strong></td>")
		[void]$sb.AppendLine("<td><strong>Succeeded</strong></td>")
		[void]$sb.AppendLine("<td><strong>Cmdlet</strong></td>")
		[void]$sb.AppendLine("<td><strong>Parameters</strong></td>")
		[void]$sb.AppendLine("<td><strong>Object Modified</strong></td>")
		[void]$sb.AppendLine("</tr>")
	}
	
	process {
		[void]$sb.AppendLine("<tr>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.Caller.split("/")[-1])</td>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.RunDate.ToString())</td>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.Succeeded)</td>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.cmdletname)</td>")
		$cmdletparameters += $AuditLogEntry.cmdletparameters | %{
			"$($_.name) : $($_.value)<br>"
		}
		[void]$sb.AppendLine("<td>$cmdletparameters</td>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.ObjectModified)</td>")
		[void]$sb.AppendLine("</tr>")
		$cmdletparameters = $null
	}
	
	end {
		[void]$sb.AppendLine("</table>")
		Write-Output $sb.ToString()
	}
}

######################################################################
# API to load credential from generic credential store
######################################################################
$CredManager = @"
using System;
using System.Net;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace SyncSiteMailbox
{
    /// <summary>
    /// </summary>
    public class CredManager
    {
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CredReadW")]
        public static extern bool CredRead([MarshalAs(UnmanagedType.LPWStr)] string target, [MarshalAs(UnmanagedType.I4)] CRED_TYPE type, UInt32 flags, [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(CredentialMarshaler))] out Credential cred);
        
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto, EntryPoint = "CredFree")]
        public static extern void CredFree(IntPtr buffer);

        /// <summary>
        /// </summary>
        public enum CRED_TYPE : uint
        {
            /// <summary>
            /// </summary>
            CRED_TYPE_GENERIC = 1,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_PASSWORD = 2,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_CERTIFICATE = 3,

            /// <summary>
            /// </summary>
            CRED_TYPE_DOMAIN_VISIBLE_PASSWORD = 4,

            /// <summary>
            /// </summary>
            CRED_TYPE_MAXIMUM = 5, // Maximum supported cred type
        }
        
        /// <summary>
        /// </summary>
        public enum CRED_PERSIST : uint
        {
            /// <summary>
            /// </summary>
            CRED_PERSIST_SESSION = 1,

            /// <summary>
            /// </summary>
            CRED_PERSIST_LOCAL_MACHINE = 2,

            /// <summary>
            /// </summary>
            CRED_PERSIST_ENTERPRISE = 3
        }
        
        /// <summary>
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct CREDENTIAL
        {
            internal UInt32 flags;
            internal CRED_TYPE type;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string targetName;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string comment;
            internal System.Runtime.InteropServices.ComTypes.FILETIME lastWritten;
            internal UInt32 credentialBlobSize;
            internal IntPtr credentialBlob;
            internal CRED_PERSIST persist;
            internal UInt32 attributeCount;
            internal IntPtr credAttribute;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string targetAlias;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string userName;
        }
        
        /// <summary>
        /// Credential
        /// </summary>
        public class Credential
        {
            private SecureString secureString = null;

            /// <summary>
            /// </summary>
            internal Credential(CREDENTIAL cred)
            {
                this.credential = cred;
                unsafe
                {
                    this.secureString = new SecureString((char*)this.credential.credentialBlob.ToPointer(), (int)this.credential.credentialBlobSize / sizeof(char));
                }                
            }

            /// <summary>
            /// </summary>
            public string UserName
            {
                get { return this.credential.userName; }
            }

            /// <summary>
            /// </summary>
            public SecureString Password
            {
                get
                {
                    return this.secureString;
                }
            }

            /// <summary>
            /// </summary>
            internal CREDENTIAL Struct
            {
                get { return this.credential; }
            }

            private CREDENTIAL credential;
        }

        internal class CredentialMarshaler : ICustomMarshaler
        {
            public void CleanUpManagedData(object ManagedObj)
            {
                // Nothing to do since all data can be garbage collected.
            }

            public void CleanUpNativeData(IntPtr pNativeData)
            {
                if (pNativeData == IntPtr.Zero)
                {
                    return;
                }
                CredFree(pNativeData);
            }

            public int GetNativeDataSize()
            {
                return Marshal.SizeOf(typeof(CREDENTIAL));
            }

            public IntPtr MarshalManagedToNative(object obj)
            {
                Credential cred = (Credential)obj;
                if (cred == null)
                {
                    return IntPtr.Zero;
                }

                IntPtr nativeData = Marshal.AllocCoTaskMem(this.GetNativeDataSize());
                Marshal.StructureToPtr(cred.Struct, nativeData, false);

                return nativeData;
            }

            public object MarshalNativeToManaged(IntPtr pNativeData)
            {
                if (pNativeData == IntPtr.Zero)
                {
                    return null;
                }
                CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(pNativeData, typeof(CREDENTIAL));
                return new Credential(cred);
            }

            public static ICustomMarshaler GetInstance(string cookie)
            {
                return new CredentialMarshaler();
            }
        }    
        

        /// <summary>
        /// ReadCredential
        /// </summary>
        /// <param name="credentialKey"></param>
        /// <returns></returns>
        public static NetworkCredential ReadCredential(string credentialKey)
        {
            Credential credential;
            CredRead(credentialKey, CRED_TYPE.CRED_TYPE_GENERIC, 0, out credential);
            return credential != null ? new NetworkCredential(credential.UserName, credential.Password) : null;
        }
    }
}
"@

######################################################################
# Load credential APIs
######################################################################
$CredManagerType = $null
try
{
    $CredManagerType = [SyncSiteMailbox.CredManager]
}
catch [Exception]
{
}

if($null -eq $CredManagerType)
{
    $compilerParameters = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters
    $compilerParameters.CompilerOptions = "/unsafe"
    [void]$compilerParameters.ReferencedAssemblies.Add("System.dll")
    Add-Type $CredManager -CompilerParameters $compilerParameters
    $CredManagerType = [SyncSiteMailbox.CredManager]
}

######################################################################
# Load tenant credential from generic credential store
######################################################################
$TenantCredential = $null #Primary

#Write-Host "Load tenant credential is from generic credential store."
try
{
    $credential = $CredManagerType::ReadCredential($TenantCredentialKey)
    if ($null -ne $credential)
    {
        $TenantCredential = New-Object System.Management.Automation.PSCredential ($credential.UserName, $credential.SecurePassword);
    }
}
catch [Exception]
{
    $TenantCredential = $null
    $errorMessage = $_.Exception.Message
    Write-Host "Tenant credential cannot be loaded correctly: $errorMessage."
}

if ($null -eq $TenantCredential)
{
    Write-Host "Tenant credential cannot be loaded please ensure you have configured in credential manager correctly."
}

######################################################################
# Script Start
######################################################################
Connect-ExchangeOnline $TenantCredential

if (($From.Length -gt 0) -AND ($To.Length -gt 0) -AND ($SmtpServer.Length -gt 0))
{
    Send-MailMessage -To $To `
    -From $From `
    -Subject "Exchange Online (Office 365) Audit Log Report for $((get-date).ToShortDateString())" `
    -Body (Search-AdminAuditLog -StartDate ((Get-Date).AddHours(-24)) -EndDate (Get-Date) | New-AuditLogReport) `
    -SmtpServer $SmtpServer `
    -BodyAsHtml
}
else
{
    $FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
    $Report = Search-AdminAuditLog -StartDate ((Get-Date).AddHours(-24)) -EndDate (Get-Date) | New-AuditLogReport
    $Report | Out-File $env:TEMP"\ExchangeOnlineAuditReport-"$FileDate".html"
}
Function Get-TenantID
{

param 
(    
       [Parameter(Mandatory = $true,
       HelpMessage="Enter the Name of the tenant (e.g. contoso.onmicrosoft.com)")]
       [ValidateNotNullOrEmpty()]
       [string]$TenantName
)

$URL = "https://login.windows.net/$TenantName/federationmetadata/2007-06/federationmetadata.xml"

$tempfile = [System.IO.Path]::GetTempFileName()

Try
{
  invoke-webrequest -Uri $URL -outfile $tempfile -erroraction stop
}
Catch
{
  Write-host -Fore yellow "Tenant name provided does not appear to be valid, can't locate \ reach the URL : "
  Write-Host "https://login.windows.net/" -nonewline
  Write-Host -fore Red $TenantName -nonewline
  Write-Host "/federationmetadata/2007-06/federationmetadata.xml"
  exit
}

[xml]$FileInput = get-content $tempfile

$FileInput.EntityDescriptor.entityID.split("/")[3]

remove-item $TempFile

}
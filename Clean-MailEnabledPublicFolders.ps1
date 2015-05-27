<#
    .SYNOPSIS
    Remove proxy addressess for a selected protocol from mailo enabled public folders
   
   	Thomas Stensitzki
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.0, 2015-05-26

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    More information can be found at http://www.granikos.eu/en/scripts 
	
    .DESCRIPTION
	
    This script removes the proxy address(es) for a selected protocol from 
    mail enabled public folders.

    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - Exchange Server 2010/2013

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
	
	.PARAMETER ProtocolToRemove
    Proxy address protocol to remove, e.g. "MS:*", "CCMAIL:*"

    .PARAMETER UpdateAddresses
    Update proxy addresses by removing found protocol addresses

    .PARAMETER OutputFile
    File name for output file, default: RemovedAddresses.txt
 
	.EXAMPLE
    Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
    Do not remove and update addresses, but log found addresses to RemovedAddresses.txt
    .\Clean-EmailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" 

    .EXAMPLE
    Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
    Remove and update addresses and log found addresses to RemovedAddresses.txt
    .\Clean-EmailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" -UpdateAddresses

    #>
Param(
    [parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Proxy address protocol to remove')][string]$ProtocolToRemove,  
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Update proxy addresses by removing found protocol addresses')][switch]$UpdateAddresses,
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='File name for output file')][string]$OutputFile = "RemovedAddresses.txt"

)

Set-StrictMode -Version Latest

$PublicFolders = Get-MailPublicFolder -ResultSize Unlimited

$max = ($PublicFolders | Measure-Object).Count
$pf = 0
$updated = 0
$found = 0

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

Write-Host "Script started!"
Write-Host "Updating mail enabled public folders having $($ProtocolToRemove) addresses"

if ($UpdateAddresses) {
    Write-Host "Email addresses will be updated!"
}
else {
    Write-Host "Email addresses will NOT be updated. Dry run only!" 
}

foreach($Folder in $PublicFolders) {

    Write-Progress -Activity "Checking Public Folder $($Folder.Name)" -Status "Object ($pf/$max)" -PercentComplete((($pf+1)/$max)*100)
    
    $proxyUpdated = $false
    
    for ($i=0;$i -lt $Folder.EmailAddresses.Count; $i++)
    {
        $address = $Folder.EmailAddresses[$i]
        if ($address.IsPrimaryAddress -eq $true -and $address.ProxyAddressString -like $ProtocolToRemove )
        {
            $found++
            $proxyUpdated = $true
            
            # Remove found proxy address
            $Folder.EmailAddresses.RemoveAt($i)
            # Fix alias (mailNickname), if required
            if($Folder.Alias.Contains(" ")) {
                $Folder.Alias = $Folder.Alias.ToString().Replace(" ","-")
            }

            $line = "$($Folder.Name);$($address.ProxyAddressString);$($Folder.Alias)"
            $line | Out-File (Join-Path -Path $ScriptDir $OutputFile) -Append -Encoding utf8

            $i--
        }
    }

    if($UpdateAddresses -and $proxyUpdated) { 
        Set-MailPublicFolder -Identity $Folder.Identity -EmailAddresses $Folder.EmailAddresses -Alias $Folder.Alias
        $updated++
    }

    $pf++
}

Write-Host "Script finished!"
Write-Host "$($pf) folders parsed, $($found) folders found, $($updated) folders updated"
Write-Host "Check output file $($OutputFile) for found/updated public folders"
<#
    .SYNOPSIS
    Remove proxy addressess for a selected protocol from mailto enabled public folders
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.2, 2017-10-10

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    http://www.granikos.eu/en/scripts 
	
    .DESCRIPTION
    This script removes the proxy address(es) for a selected protocol from mail enabled public folders.

    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - Exchange Server 2010/2013/2016

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    1.1     FixAlias added, cleanup logic changed
    1.2     Some minor PowerShell updates
	
    .PARAMETER ProtocolToRemove
    Proxy address protocol to remove, e.g. "MS:*", "CCMAIL:*"

    .PARAMETER UpdateAddresses
    Update proxy addresses by removing found protocol addresses

    .PARAMETER OutputFile
    File name for output file, default: RemovedAddresses.txt
 
    .EXAMPLE
    Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
    Do not remove and update addresses, but log found addresses to RemovedAddresses.txt
    .\Clean-MailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" 

    .EXAMPLE
    Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
    Remove and update addresses and log found addresses to RemovedAddresses.txt
    .\Clean-MailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" -UpdateAddresses

#>
Param(
  [parameter(Mandatory,HelpMessage='Proxy address protocol to remove')][string]$ProtocolToRemove,  
  [switch]$UpdateAddresses,
  [string]$OutputFile = 'RemovedAddresses.txt'

)

# Set-StrictMode -Version Latest

# Fetch all public folders
$PublicFolders = Get-MailPublicFolder -ResultSize Unlimited

# Do some presets
$max = ($PublicFolders | Measure-Object).Count
$publicFoldersCount = 0
$updated = 0
$found = 0

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path

Write-Host 'Script started!'
Write-Host ('Updating mail enabled public folders having {0} addresses' -f $ProtocolToRemove)

if ($UpdateAddresses) {
  Write-Host 'Email addresses will be updated!'
}
else {
  Write-Host 'Email addresses will NOT be updated. Dry run only!' 
}

foreach($Folder in $PublicFolders) {

  # Some nice progrsss bar
  Write-Progress -Activity ('Checking Public Folder {0}' -f $Folder.Name) -Status ('Object ({0}/{1})' -f $publicFoldersCount, $max) -PercentComplete((($publicFoldersCount+1)/$max)*100)
    
  $proxyUpdated = $false
    
  for ($i=0;$i -lt $Folder.EmailAddresses.Count; $i++)
  {
    # Fetch proxy addresses
    $address = $Folder.EmailAddresses[$i]

    if ($address.IsPrimaryAddress -eq $true -and $address.ProxyAddressString -like $ProtocolToRemove )
    {
      $found++
      $proxyUpdated = $true
            
      # Remove found proxy address
      $Folder.EmailAddresses.RemoveAt($i)

      # Fix alias (mailNickname), if required
      if($Folder.Alias.Contains(' ')) {
        $Folder.Alias = $Folder.Alias.ToString().Replace(' ','-')
      }

      $line = ('{0};{1};{2}' -f $Folder.Name, $address.ProxyAddressString, $Folder.Alias)
      $line | Out-File -FilePath (Join-Path -Path $ScriptDir -ChildPath $OutputFile) -Append -Encoding utf8

      $i--
    }
  }

  if($UpdateAddresses -and $proxyUpdated) { 
    Set-MailPublicFolder -Identity $Folder.Identity -EmailAddresses $Folder.EmailAddresses -Alias $Folder.Alias
    $updated++
  }

  $publicFoldersCount++
}

Write-Host 'Script finished!'
Write-Host ('{0} folders parsed, {1} folders found, {2} folders updated' -f $publicFoldersCount, $found, $updated)
Write-Host ('Check output file {0} for found/updated public folders' -f $OutputFile)
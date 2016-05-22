<#
    .SYNOPSIS
    Remove proxy addressess for a selected protocol from mail enabled public folders and fix aliases
   
   	Thomas Stensitzki
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.1, 2016-01-21

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    More information can be found at http://www.granikos.eu/en/scripts 

    .DESCRIPTION
	This script removes the proxy address(es) for a selected protocol from mail enabled public folders.
    
    The script can fix the alias of mail enabled public folders as well. The code used is based opon a blog post by Shay Levy.
    http://blogs.microsoft.co.il/scriptfanatic/2011/08/15/exchange-removing-illegal-alias-characters-using-powershell/

    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - Exchange Server 2010/2013

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    1.1     FixAlias added, cleanup logic changed
	
	.PARAMETER ProtocolToRemove
    Proxy address protocol to remove, e.g. "MS:*", "CCMAIL:*"

    .PARAMETER UpdateAddresses
    Switch to update proxy addresses by removing found addresses matching protocol provided in parameter ProtocolToRemove

    .PARAMETER OutputFile
    File name for output file, default: RemovedAddresses.txt
    
    .PARAMETER FixAlias
    Switch to fix mail enabled public folder alias (mailNickname) and to remove illegal characters
 
	.EXAMPLE
    Check mail enabled public folders for proxy addresses having "MS:" as a protocol type.
    Do not remove and update addresses, but log found addresses to RemovedAddresses.txt
    .\Clean-EmailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" 

    .EXAMPLE
    Check mail enabled public folders for proxy addresses having "MS:" as a protocol type.
    Remove and update addresses and log found addresses to RemovedAddresses.txt
    .\Clean-EmailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" -UpdateAddresses

    #>
Param(
    [parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Proxy address protocol to remove')][string]$ProtocolToRemove,  
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Update proxy addresses by removing found protocol addresses')][switch]$UpdateAddresses,
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Fix Alias to remove illegalcharacters')][switch]$FixAlias,
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='File name for output file')][string]$OutputFile = "RemovedAddresses.txt"

)

Set-StrictMode -Version Latest

Write-Host "Please wait. Fetching mail enabled public folder..."
$PublicFolders = Get-MailPublicFolder -ResultSize Unlimited

$max = ($PublicFolders | Measure-Object).Count
$pf = 0
$updated = 0
$found = 0
$fixed = 0

# Thanks to Shay Levy
$IllegalCharacters = 0..34+40..41+44,47+58..60+62+64+91..93+127..160

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

Write-Host "$($max) mail enabled public folders fetched from Active Directory"
Write-Host "Updating mail enabled public folders having $($ProtocolToRemove) addresses"

if ($UpdateAddresses) {
    Write-Host "Email addresses will be updated!"
}
else {
    Write-Host "Email addresses will NOT be updated. Dry run only!" 
}

# Write file header
$line = "Name;ProxyAddress;OldAlias;NewAlias"
$line | Out-File (Join-Path -Path $ScriptDir $OutputFile) -Append -Encoding utf8

foreach($Folder in $PublicFolders) {

    # Write some nice progress bar
    Write-Progress -Activity "Checking Public Folder $($Folder.Name)" -Status "Object ($pf/$max)" -PercentComplete((($pf+1)/$max)*100)
    
    $proxyUpdated = $false
    $aliasUpdated = $false
    
    for ($i=0;$i -lt $Folder.EmailAddresses.Count; $i++)
    {
        $address = $Folder.EmailAddresses[$i]
        $newAlias = $Folder.Alias
        
        # Thanks to Shay Levy
        # Check alias for each illegal character
        foreach ($char in $IllegalCharacters) {

            $escapedChar = [regex]::Escape([char]$char)

            if($newAlias -match $escapedChar){
                $newAlias = $newAlias -replace $escapedChar
            }
        }
        
        $aliasUpdated = ($Folder.Alias -ne $newAlias)

        if ($address.IsPrimaryAddress -eq $true -and $address.ProxyAddressString -like $ProtocolToRemove )
        {
            # Yes, we've found an address to remove
            $found++
            $proxyUpdated = $true
         
            # Remove found proxy address
            $Folder.EmailAddresses.RemoveAt($i)
            
            # log folder name, address and alias to file
            $line = "$($Folder.Name);$($address.ProxyAddressString);$($Folder.Alias);$($newAlias)"
            $line | Out-File (Join-Path -Path $ScriptDir $OutputFile) -Append -Encoding utf8

            $i--
        }
    }

    if($UpdateAddresses -and $proxyUpdated) { 
        Set-MailPublicFolder -Identity $Folder.Identity -EmailAddresses $Folder.EmailAddresses
        $updated++
    }
    
    if($FixAlias -and $aliasUpdated) {
        Set-MailPublicFolder -Identity $Folder.Identity -Alias $newAlias
        $fixed++
    }

    $pf++
}

Write-Host "Script finished!"
Write-Host "$($pf) folders parsed, $($found) folders with $($ProtocolToRemove) address found, $($updated) folders updated, $($fixed) aliases fixed"
Write-Host "Check output file $($OutputFile) for found/updated public folders"
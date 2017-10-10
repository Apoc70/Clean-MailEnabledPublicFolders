# Clean-MailEnabledPublicFolders.ps1
Remove proxy addressess for a selected protocol from mailo enabled public folders

## Description
This script removes the proxy address(es) for a selected protocol from mail enabled public folders.

## Parameters
### ProtocolToRemove
Proxy address protocol to remove, e.g. "MS:*", "CCMAIL:*"

### UpdateAddresses
Update proxy addresses by removing found protocol addresses

### OutputFile
File name for output file, default: RemovedAddresses.txt


## Examples
```
.\Clean-MailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*"
```
Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
Do not remove and update addresses, but log found addresses to RemovedAddresses.txt

```
.\Clean-MailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" -UpdateAddresses
```
Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
Remove and update addresses and log found addresses to RemovedAddresses.txt

## TechNet Gallery
Find the script at TechNet Gallery
* https://gallery.technet.microsoft.com/Script-to-remove-unwanted-9d119c6b


## Credits
Written by: Thomas Stensitzki

Stay connected:

* My Blog: http://justcantgetenough.granikos.eu
* Twitter: https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github: https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog: http://blog.granikos.eu/
* Website: https://www.granikos.eu/en/
* Twitter: https://twitter.com/granikos_de

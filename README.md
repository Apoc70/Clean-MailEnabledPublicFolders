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

``` PowerShell
.\Clean-MailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*"
```

Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
Do not remove and update addresses, but log found addresses to RemovedAddresses.txt

``` PowerShell
.\Clean-MailEnabledPublicFolders.ps1 -ProtocolToRemove "MS:*" -UpdateAddresses
```

Check mal enabled public folders for proxy addresses having "MS:" as a protocol type.
Remove and update addresses and log found addresses to RemovedAddresses.txt

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)
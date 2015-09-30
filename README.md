# posh-o365
PowerShell scripts for management of Office 365

Author: Mark Wilson
Readme last updated 16 September 2015

posh-o365 is a collection of scripts I've used for managing Office 365. They may not be the "best" PowerShell in the world, but they worked for me, and my hope is that by open-sourcing them, others will improve and add to the collection.

Currently, there are four scripts:

* Check-ODBQuotas.ps1 is used to examine a list of users, calculate their OneDrive for Business folder URL and check to see if the quota is as expected (or if indeed the site exists). The script takes three parameters:

Check-ODBQuotas.ps1 filename quota tenant

Filename is a list of UPNs, in CSV format.
Quota is the intended storage quota size, in MB (i.e. the size to treat as correct and flag as green).
Tenant is the tenant name.

* Connect-O365.ps1 is used to issue a single command to log on to various Office 365 services in a single PowerShell window. The script takes one parameter:

Connect-O365.ps1 tenant

Tenant is the tenant name.

* Set-O365Licences.ps1 is used to examine a list of users, licences and locations before assigning licences according. The script takes two parameters:

Set-O365Licences.ps1 filename tenant

Filename should be in CSV format, with the following headings: UPN,Licence,Location. For example:

UPN,Licence,Location
user.1@domainname.tld,E3,GB
user.2@domainname.tld,None,GB

Tenant is the tenant name.

At this time, Set-O365Licences.ps1 has only been tested with AU/AT/CN/DE/FR/GB/HK/IT/SG/ZA usage locations and E1/E3 SKUs; however it would be simple to extend if required. Usage locations are based on ISO 3166 codes at https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2

* Set-ODBQuota.ps1 is used to set the OneDrive for Business storage quota for a user. Because of the way that OneDrive for Business sites are provisioned, it does need the user to have visited their OneDrive first, but the quota can be set anywhere between 1GB and 1TB (the current limits that Microsoft imposes on the service).  The command takes three parameters.

Set-ODBQuota.ps1 upn quota tenant

UPN is the UPN for the user whose quota is to be set.
Quota is the intended storage quota size, in MB.
Tenant is the tenant name (i.e. the tenantname part in tenantname.onmicrosoft.com).

These scripts are provided "as is" and with no warranty, express or implied, as to their suitability. No responsibility can be taken any actions as a result of the use of these scripts.

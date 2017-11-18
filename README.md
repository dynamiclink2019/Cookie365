# Cookie365 v.0.7
Utility to Mount OneDrive For Business As a Disk - Office365

Works also in ADSF integrated environments, without asking user/password

Usage: 

> 

> Cookie365 -s URL [-u user@domain.com | -d domain.com] [-p {password}] [-quiet] [-mount [disk] [-homedir]]

Mandatory



> -s <sharepoint URL(example https://yourdomain-my.sharepoint.com)

Optional:


>  -u <user>: if you are not using ADFS you can specify a username with the format user@domain.com


>  -p <password>: if you are not using ADFS you need to specify your password
 

> -d <domain>: if you are using ADFS but your internal domain is different from Office365 domain you 
              can specify a different Office365 domain with the format domain.com 


>  -quiet: be quiet...


>  -mount: if you specify a disk name (e.g. z:) it will be used, otherwise the OS will assign the first available disk. Optionally you can specify the -homedir option in order to mount the drive with the path for the specific OneDrive For Business user (\DavWWWRoot\personal\<user>)


Examples
In this example the utility will authenticate with user/password and will mount OneDrive repository as drive X:
Cookie365 -s https://yoursite.sharepoint.com -user user@domain.com -p password -mount x: -homedir

In this example the utility will only authenticate with user/password and create cookies needed to mount onedrive repository
Cookie365 -s https://yoursite.sharepoint.com -user user@domain.com -p password 

In this example the utility will authenticate using ADFS (and will not ask any user/password to the user) and will mount OneDrive repository as drive X: - In order to work, your environment should be ADFS enabled, and your users should be logged in in  Domain-Joined PC
Cookie365 -s https://yoursite.sharepoint.com -mount x: -homedir


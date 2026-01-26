
This is a Microsoft 365 Audit Script read-only tenant-wide audit using Microsoft Graph and Exchange Online (app-only authentication). It gathers detailed information on Entra ID groups, Microsoft 365 Groups/Teams, SharePoint usage, and Shared Mailboxes, 
including owners, members, email aliases, activity logs, and delegation settings. Results are exported to a structured Excel workbook to support governance, security, compliance, and client asset reporting.


##Prerequisites – Azure App Registration
i. Authentication method: Certificate-based authentication (PFX uploaded to the App Registration)
ii. Application permissions (not delegated) are required & Admin consent must be granted for all permissions

##Required Microsoft Graph API Permissions
##The App Registration must include the following Application permissions, depending on which report  you run:

API: Permission name:
Micrsooft Graph
Office 365 Exchange Online
SharePoint


| API / Permission Name                   | Permission Type | Description                                             | Admin Consent | Status  |
| --------------------------------------- | --------------- | ------------------------------------------------------- | ------------- | ------- |
| DeviceManagementApps.Read.All           | Application     | Read Microsoft Intune apps                              | Yes           | Granted |
| DeviceManagementConfiguration.Read.All  | Application     | Read Microsoft Intune device configuration and policies | Yes           | Granted |
| DeviceManagementManagedDevices.Read.All | Application     | Read Microsoft Intune devices                           | Yes           | Granted |
| Directory.Read.All                      | Delegated       | Read directory data                                     | Yes           | Granted |
| Directory.Read.All                      | Application     | Read directory data                                     | Yes           | Granted |
| Directory.ReadWrite.All                 | Delegated       | Read and write directory data                           | Yes           | Granted |
| Directory.ReadWrite.All                 | Application     | Read and write directory data                           | Yes           | Granted |
| Group.Read.All                          | Application     | Read all groups                                         | Yes           | Granted |
| Group.ReadWrite.All                     | Delegated       | Read and write all groups                               | Yes           | Granted |
| Group.ReadWrite.All                     | Application     | Read and write all groups                               | Yes           | Granted |
| MailboxSettings.Read                    | Application     | Read all user mailbox settings                          | Yes           | Granted |
| Policy.Read.All                         | Application     | Read organizational policies                            | Yes           | Granted |
| Reports.Read.All                        | Application     | Read usage and activity reports                         | Yes           | Granted |
| RoleManagement.Read.Directory           | Application     | Read directory RBAC role assignments and definitions    | Yes           | Granted |
| Sites.FullControl.All                   | Application     | Full control of all SharePoint site collections         | Yes           | Granted |
| Sites.ReadWrite.All                     | Application     | Read and write items in all site collections            | Yes           | Granted |
| User.Read                               | Delegated       | Sign in and read user profile                           | No            | Granted |
| User.Read.All                           | Application     | Read all users’ profiles                                | Yes           | Granted |


| API / Permission Name | Permission Type | Description                              | Admin Consent | Status  |
| --------------------- | --------------- | ---------------------------------------- | ------------- | ------- |
| Exchange.ManageAsApp  | Application     | Manage Exchange Online as an application | Yes           | Granted |



| API / Permission Name | Permission Type | Description                                     | Admin Consent | Status  |
| --------------------- | --------------- | ----------------------------------------------- | ------------- | ------- |
| AllSites.Write        | Delegated       | Read and write items in all site collections    | No            | Granted |
| Sites.FullControl.All | Application     | Full control of all SharePoint site collections | Yes           | Granted |


## iii. PowerShell Requirements: PowerShell 7.x , PnP PowerShell v3.1.0

## iv. Exchange Online PowerShell is required to pull the "Full Access" and "Send As" delegation lists, as Microsoft Graph does not expose these permissions in a simple way.
Run the script as Admin:
## Install the Exchange Online module
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser

# Verify the version (ensure it is 3.0.0 or higher for Certificate Auth support)
Get-Module -ListAvailable ExchangeOnlineManagement

## A tenant-level administrative roles is needed 
A user must hold one of the following Entra ID roles to grant consent:
Global Administrator 
Privileged Role Administrator 
Cloud Application Administrator 
Application Administrator  (can grant consent in most tenants, including Graph)

##Important Notes
Scripts will fail silently or return partial data if required permissions are missing
Always verify permissions after cloning or deploying to a new tenant
Certificate .cer must be uploaded to the App Registration before use

##Report  data: Privacy Settings (Required for Activity Data)
To see actual URLs and Activity, you must ensure your tenant is not hiding this data:
Go to Microsoft 365 Admin Center > Settings > Org Settings.
Go to the Services tab and select Reports.
Uncheck the box: "Display concealed user, group, and site names in all reports".
Save changes. (It may take 24 hours for the API to reflect clear-text URLs).

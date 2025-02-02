# M365PowerKit

Powershell module for Microsoft 365 / Entrata API and related Microsoft Cloud service automation tasks. Tasks primarily relate to maintenance, monitoring, and data retrieval for proactive maintenance of Information Security and Compliance.

## Installation

### Prerequisites

- Windows PowerShell 7.3 or later
- M365 User and Privileges required for the specific tasks

### Quick Start

```powershell
   git clone https://github.com/OrganisationServiceManagement/M365PowerKit.git
   cd .\M365PowerKit
  .\osm-m365-go.ps1
```

## Usage

```powershell
  # Use the moduel and sub-modules via :
  .\osm-m365-go.ps1
```

```powershell
[Parameter(Mandatory = $false)]
    [switch]$UpdateModule = $false
# Switch to indicate whether the M365PowerKit module should be forcefully removed and reimports from local source (if you make local code changes, but don't increment module version numbers, this is required).

[Parameter(Mandatory = $false)]
[string]$InputFunctionName = $null,
#The name of the function to run from the M365PowerKit module. Leave empty to show the M365PowerKit UI.

[Parameter(Mandatory = $false)]
[hashtable]$FunctionParameterHashTable = $null
# A hashtable of parameters to pass to the specified function. Not required, if the called function requires parameters, it will prompt you for them... but that is quite cumbersome, so in general  have a note pade with the FunctionParameterHashTable and InputFunctionNames as you use the tool.

.EXAMPLE
.\osm-m365-go.ps1
.\osm-m365-go.ps1 -UpdateModule
.\osm-m365-go.ps1 -UpdateModule -InputFunctionName "Get-AllSMTPAddresses"

###
$EMAIL_ATTACHMENT_EXPORT_JOB_HASH = @{
  UPN                 = 'security_report_reader-roleuser@mytenent.bleh'
  MailboxName         = 'mark.culhane@zoak.solutions'
  StartDate               = $(Get-Date).AddDays(-30).ToString('yyyy-MM-dd')
  Sender_Address          = 'siem-vendor.com'
  AttachmentExtension     = 'pdf'
  BASE_DIR                = 'C:\Temp\PSTExports'
}
.\osm-m365-go.ps1 -UpdateModule -InputFunctionName "Export-NewExchangeSearch" -FunctionParameterHashTable $EMAIL_ATTACHMENT_EXPORT_JOB_HASH
###
```

## Permissions

- When a user connects to Microsoft Graph using Connect-MgGraph with delegated permissions (i.e.,permissions tied to their user account), their permissions are ultimately limited by a combination of:

- Permissions Granted to the Application (App ID): The application (In this case, identified by App ID 14d82eec-204b-4c2f-b7e8-296a70dab67e) must have admin consent for specific permissions in Entra.
- This defines the maximum scope of permissions that any user of this app can request.

- Permissions of the User’s Role and Access: Even if the application has been granted permissions, the user’s access is restricted by their role and permissions within the tenant. For instance, a user with no admin privileges will not be able to perform admin-level actions, even if the app has consent for Directory.ReadWrite.All.

### Create a new search query and retrieve attachments

```powershell
  Export-NewExchangeSearch -MailboxName "user@example.com" -UPN "admin@example.com" -StartDate "2024-04-20" -Subject "Important Policy Docs" -Sender "importantsenderdomainoraddress.com" -AttachmentExtension "pdf"
```

### Retrieve attachments for an existing search query

```powershell
    Export-ExistingExchangeSearch -AttachmentExtension "pdf" -SkipModules -SkipConnIPS -SkipDownload -SearchName "20240429_015205-Export-Job"
```

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

### License

The documentation and written content in this repository are licensed under the Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International License.
For source code and software components, See [LICENSE](LICENSE.md) file.

## Disclaimer

This module is provided as-is without any warranty or support. Use it at your own risk.

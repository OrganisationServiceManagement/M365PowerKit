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

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

### License

The documentation and written content in this repository are licensed under the Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International License.
For source code and software components, See [LICENSE](LICENSE.md) file.

## Disclaimer

This module is provided as-is without any warranty or support. Use it at your own risk.

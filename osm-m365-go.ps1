<#
.SYNOPSIS
    Script to manage and run M365PowerKit functions.

.DESCRIPTION
    This script allows you to update the M365PowerKit module, import it, and run specified functions with provided parameters.

.PARAMETER UpdateModule
    Switch to indicate whether the M365PowerKit module should be forcefully removed and reimports from local source (if you make local code changes, but don't increment module version numbers, this is required).

.PARAMETER InputFunctionName
    The name of the function to run from the M365PowerKit module. Leave empty to show the M365PowerKit UI.

.PARAMETER FunctionParameterHashTable
    A hashtable of parameters to pass to the specified function. Not required, if the called function requires parameters, it will prompt you for them... but that is quite cumbersome, so in general  have a note pade with the FunctionParameterHashTable and InputFunctionNames as you use the tool.

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

.NOTES
    Author: Mark Culhane
    Date: February 2, 2025
    Version: 1.0
#>
param (
    [Parameter(Mandatory = $false)]
    [switch]$UpdateModule = $false,
    [Parameter(Mandatory = $false)]
    [string]$InputFunctionName = $null,
    [Parameter(Mandatory = $false)]
    [hashtable]$FunctionParameterHashTable = $null
)

#if flag for -UpdateModule is set
function Test-M365PowerKit {
    param (
        [string]$InputFunctionName = $null,
        [hashtable]$FunctionParameterHashTable = $null
    )
    if ($UpdateModule) {
        Write-Host 'Removing M365PowerKit modules...'
        Remove-Module M365PowerKi* -ErrorAction Continue
        # Unsetting variables
        Remove-Item env:M365PowerKit* -ErrorAction Continue
        Get-Module
        Write-Host 'Done removing M365PowerKit modules.'
    } 
    Write-Host 'Importing M365PowerKit main...'
    Import-Module .\M365PowerKit.psd1 -Force
    Write-Host 'Done importing M365PowerKit main.'

    if ($InputFunctionName) {
        Write-Host "Running M365PowerKit with function $InputFunctionName..."
        if ($FunctionParameterHashTable) {
            Write-Host '    Function parameters:'
            $FunctionParameterHashTable | ForEach-Object {
                Write-Host "        $_" -ForegroundColor Yellow
            }
            M365PowerKit -InputFunctionName $InputFunctionName -ProvidedParameters $FunctionParameterHashTable
        }
        else {
            Write-Host '      No function parameters provided.'
            M365PowerKit -InputFunctionName $InputFunctionName
        }
    }
    else {
        Write-Host 'No function name provided, showing UI...'
        M365PowerKit
    }
}
if ($InputFunctionName) {
    if ($FunctionParameterHashTable) {
        Test-M365PowerKit -InputFunctionName $InputFunctionName -FunctionParameterHashTable $FunctionParameterHashTable -UpdateModule $UpdateModule
    }
    else {
        Test-M365PowerKit -InputFunctionName $FunctionName -UpdateModule $UpdateModule
    }
}
else {
    Test-M365PowerKit -UpdateModule $UpdateModule
}
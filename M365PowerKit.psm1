<#
.SYNOPSIS
    Atlassian Cloud PowerKit module for interacting with Atlassian Cloud REST API.
.DESCRIPTION
    Atlassian Cloud PowerKit module for interacting with Atlassian Cloud REST API.
    - Dependencies: M365PowerKit-Shared
    - Functions:
      - M365PowerKit: Interactive function to run any function in the module.
    - Debug output is enabled by default. To disable, set $DisableDebug = $true before running functions.
.EXAMPLE
    M365PowerKit
    This example lists all functions in the M365PowerKit module.
.EXAMPLE
    M365PowerKit
    Simply run the function to see a list of all functions in the module and nested modules.
.EXAMPLE
    Get-DefinedPowerKitVariables
    This example lists all variables defined in the M365PowerKit module.
.LINK
    GitHub:

#>
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
$NESTED_MODULE_ARRAY = @(
    'M365PowerKit-Shared',
    'M365PowerKit-SharePoint',
    'M365PowerKit-ExchangeSearchExport'
)
function Import-NestedModules {
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $NESTED_MODULES
    )
    $NESTED_MODULES | ForEach-Object {
        $MODULE_NAME = $_
        Write-Debug "Importing nested module: $MODULE_NAME"
        #Find-Module psd1 file in the subdirectory and import it
        $PSD1_FILE = Get-ChildItem -Path ".\$MODULE_NAME" -Filter "$MODULE_NAME.psd1" -Recurse -ErrorAction SilentlyContinue
        if (-not $PSD1_FILE) {
            Write-Error "Module $MODULE_NAME not found. Exiting."
            throw "Nested module $MODULE_NAME not found. Exiting."
        }
        elseif ($PSD1_FILE.Count -gt 1) {
            Write-Error "Multiple module files found for $MODULE_NAME. Exiting."
            throw "Multiple module files found for $MODULE_NAME. Exiting."
        }
        Import-Module $PSD1_FILE.FullName -Force
        Write-Debug "Importing nested module: $PSD1_FILE,  -- $($PSD1_FILE.BaseName)"
        #Write-Debug "Importing nested module: .\$($_.BaseName)\$($_.Name)"
        # Validate the module is imported
        if (-not (Get-Module -Name $MODULE_NAME)) {
            Write-Error "Module $MODULE_NAME not found. Exiting."
            throw "Nested module $MODULE_NAME not found. Exiting."
        }
    }
    return $NESTED_MODULES
}
function Get-ClickOnceApplication {
    $Default_Path = "$($env:LOCALAPPDATA)\Apps\2.0\"
    $Default_Filename = 'microsoft.office.client.discovery.unifiedexporttool.exe'
    $Default_URL = 'https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application'
    function Write-ClickOnceInstructuions {
        Write-Host 'To install the Unified Export Tool manually, follow these steps:'
        Write-Host "   - Open a browser and navigate to: $Default_URL"
        Write-Host '   - Click on the "Install" button to download and install the application'
        Write-Host '   - Once the installation is complete, hit "C" to continue or any other key to exit'
        $Key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').VirtualKeyCode
        if ($Key -ne 67) {
            Write-Error 'Failed to get ClickOnceApplication'
            throw 'Failed to get ClickOnceApplication'
        }
    }
    while ((-not (Test-Path -Path $Default_Path -PathType Container)) -or (-not(Get-ChildItem -Path $Default_Path -Filter $Default_Filename -Recurse))) {
        Write-Debug 'Failed to get ClickOnceApplication, looking in '
        Write-ClickOnceInstructuions
    }
    $ClickOnceApp = (Get-ChildItem -Path $Default_Path -Filter $Default_Filename -Recurse).FullName | Where-Object { $_ -notmatch '_none_' } | Select-Object -First 1
    while (!$ClickOnceApp) {
        Write-Debug 'Failed to get ClickOnceApplication, try manual install see:'
        Write-ClickOnceInstructuions
    }
    Write-Debug "ClickOnce Application Installed - Path: $ClickOnceApp"
    $ClickOnceApp
}

# function to run provided functions with provided parameters (as hash table)
function Invoke-M365PowerKitFunction {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FunctionName,
        [Parameter(Mandatory = $false)]
        [hashtable]$ProvidedParameters
    )    
    $TEMP_DIR = "$env:OSM_HOME\M365PowerKit\.temp"
    if (-not (Test-Path $TEMP_DIR)) {
        New-Item -ItemType Directory -Path $TEMP_DIR -Force | Out-Null
    }
    $TIMESTAMP = Get-Date -Format 'yyyyMMdd-HHmmss'
    $LOG_FILE = "$TEMP_DIR\$FunctionName-$TIMESTAMP.log"
    # Start transcript logging
    Start-Transcript -Path $LOG_FILE -Append
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        # Invoke expression to run the function, splatting the parameters
        $stopwatch.Start()
        if ($ProvidedParameters) {
            $singleLineDefinition = $ProvidedParameters.Keys | ForEach-Object { "-   ->    $_ = $($ProvidedParameters.($_))" }
            Write-Debug "Running function: $FunctionName with parameters: $singleLineDefinition"
            & $FunctionName @ProvidedParameters
        }
        else {
            Invoke-Expression "$FunctionName"
        }
    }
    catch {
        Write-Error "Failed to run function: $FunctionName"
    }
    finally {
        $stopwatch.Stop()
        Write-Debug "Function: $FunctionName completed in $($stopwatch.Elapsed.TotalSeconds) seconds"
        # Stop transcript logging
        Stop-Transcript
        Write-Host "Log: $LOG_FILE"
         
    }
}

function Show-M365PowerKitFunctions {
    
    # List nested modules and their exported functions to the console in a readable format, grouped by module
    $colors = @('Green', 'Cyan', 'Red', 'Magenta', 'Yellow')
    $colorIndex = 0
    $functionReferences = @{}
    $NESTED_MODULE_ARRAY | ForEach-Object {
        $MODULE = Get-Module -Name $_
        # Select a color from the list
        $color = $colors[$colorIndex % $colors.Count]
        $spaces = ' ' * (52 - $MODULE.Name.Length)
        Write-Host '' -BackgroundColor Black
        Write-Host "Module: $($MODULE.Name)" -BackgroundColor $color -ForegroundColor White -NoNewline
        Write-Host $spaces -BackgroundColor $color -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $spaces = ' ' * 41
        Write-Host " Exported Commands:$spaces" -BackgroundColor "Dark$color" -ForegroundColor White -NoNewline
        Write-Host ' ' -BackgroundColor Black
        $MODULE.ExportedCommands.Keys | ForEach-Object {
            # Assign a letter reference to the function
            $functRefNum = $colorIndex
            $functionReferences[$functRefNum] = $_

            Write-Host ' ' -NoNewline -BackgroundColor "Dark$color"
            Write-Host '   ' -NoNewline -BackgroundColor Black
            Write-Host "$functRefNum -> " -NoNewline -BackgroundColor Black
            Write-Host "$_" -NoNewline -BackgroundColor Black -ForegroundColor $color
            # Calculate the number of spaces needed to fill the rest of the line
            $spaces = ' ' * (50 - $_.Length)
            Write-Host $spaces -NoNewline -BackgroundColor Black
            Write-Host ' ' -NoNewline -BackgroundColor "Dark$color"
            Write-Host ' ' -BackgroundColor Black
            # Increment the color index for the next function
            $colorIndex++
        }
        $spaces = ' ' * 60
        Write-Host $spaces -BackgroundColor "Dark$color" -NoNewline
        Write-Host ' ' -BackgroundColor Black
    }
    Write-Host 'Note: You can run functions without this interface by calling them directly.' 
    Write-Host "Example: Invoke-M365PowerKitFunction -FunctionName 'FunctionName' -ProvidedParameters @{ 'ParameterName' = 'ParameterValue' }" 
    # Write separator for readability
    Write-Host "`n" -BackgroundColor Black
    Write-Host '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++' -BackgroundColor Black -ForegroundColor DarkGray
    # Ask the user which function they want to run
    $selectedFunction = Read-Host -Prompt "`nSelect a function to run by ID, or FunctionName [parameters] (or hit enter to exit):"
    # if user enters a number, get the function name from the reference table and update the selectedFunction variable
    if ($selectedFunction -match '(\d+)') {
        $selectedFunction = [int]$selectedFunction
        $selectedFunction = $functionReferences[$selectedFunction]
    }
    # if the user enters a function name and parameters, run it with the provided parameters as a hash table
    if ($selectedFunction -match '(\w+)\s*\[(.*)\]') {
        $functionName = $matches[1]
        $ProvidedParameters = $matches[2] -split '\s*,\s*' | ForEach-Object {
            $key, $value = $_ -split '\s*=\s*'
            @{ $key = $value }
        }
        Write-Debug "Invoking: $functionName with parameters: $ProvidedParameters"
        Invoke-M365PowerKitFunction -FunctionName $functionName -ProvidedParameters $ProvidedParameters
    }
    elseif ($selectedFunction -match '(\w+)') {
        Write-Debug "Selected function: $selectedFunction withou parameters"
        Invoke-M365PowerKitFunction -FunctionName $selectedFunction
    }
    elseif ($selectedFunction -eq '') {
        return $null
    }
    else {
        Write-Host 'Invalid selection. Please try again.' -ForegroundColor Red
        Show-M365PowerKitFunctions
    }
    # Ask the user if they want to run another function
    $runAnother = Read-Host -Prompt 'Run another function? (Y / any key to exit)'
    if ($runAnother -eq 'Y') {
        Show-M365PowerKitFunctions
    }
    else {
        Write-Host 'Have a great day!'
        return $null
    }
}

function M365PowerKit {
    param (
        [Parameter(Mandatory = $false)]
        [string]$UPN,
        [Parameter(Mandatory = $false)]
        [string]$InputFunctionName,
        [Parameter(Mandatory = $false)]
        [hashtable]$ProvidedParameters,
        [Parameter(Mandatory = $false)]
        [switch]$ClearProfile = $false
    )
    # Import Shared Module
    $MODULE_LIST = Import-NestedModules -NESTED_MODULES $NESTED_MODULE_ARRAY
    Write-Debug "$($MyInvocation.MyCommand.Name) - Imported Nested Modules: $MODULE_LIST"

    if (! $env:M365PowerKit_DependenciesInstalled) {
        Install-M365Dependencies -DependencySet 'All'
    }
    if ($ClearProfile) {
        Write-Debug 'Clearing M365PowerKit Profile and Connections...'
        Clear-M365PowerKitProfile
        Write-Debug 'M365PowerKit Profile and Connections cleared - OK'
    }

    if (! $UPN) {
        if (!$env:M365PowerKitUPN) {
            $UPN = Read-Host 'Enter the User Principal Name (UPN) for the Exchange Online session'
        }
    }
    $env:M365PowerKitUPN = $UPN
    Write-Debug "M365PowerKit UPN set to: $UPN"
        
    #     Write-Debug 'Cleared M365PowerKit Profile and Connections'
    # }
    try {
        # New-IPPSSession
        # New-EXOSession
        # If function is called with a function name, run that function with the provided ProvidedParameters
        if ($InputFunctionName) {
            try {
                if ($ProvidedParameters) {
                    Write-Debug "$($MyInvocation.MyCommand.Name) - Running function: $InputFunctionName with ProvidedParameters provided"
                    Invoke-M365PowerKitFunction -FunctionName $InputFunctionName -ProvidedParameters $ProvidedParameters
                }
                else {
                    Write-Debug "$($MyInvocation.MyCommand.Name) - Running function: $InputFunctionName without parameters"
                    Invoke-M365PowerKitFunction -FunctionName $InputFunctionName
                }
                Invoke-M365PowerKitFunction -FunctionName $InputFunctionName -ProvidedParameters $ProvidedParameters
            }
            catch {
                Write-Debug "$($MyInvocation.MyCommand.Name) - FAILED: $_"
                Write-Error "$($MyInvocation.MyCommand.Name) - Failed to run function: $InputFunctionName with parameters: $ProvidedParameters"
            }
        }
        else {
            Write-Debug "$($MyInvocation.MyCommand.Name) - No function name provided, showing UI..."
            Show-M365PowerKitFunctions
        }
    }
    catch {
        Write-Debug "$($MyInvocation.MyCommand.Name) - Error: $_"
        Write-Error "M365PowerKit threw an error: $_"
    }
    finally {
        Write-Debug 'M365PowerKit completed successfully'
        #Write-Host 'Clearing profile...'
        #Clear-M365PowerKitProfile
    }
}

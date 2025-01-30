# Set exit on error
$ErrorActionPreference = 'Stop'; $DebugPreference = 'Continue'
# Function: Install-SharedDependencies
# Description: This function installs the shared dependencies for the M365PowerKit module.
$M365POWERKIT_DEPENDENCIES = @(
    # 'Microsoft.Graph.Authentication',
    # 'Microsoft.Graph.Sites',
    # 'Microsoft.Graph.Files'
    #  'WINGET_Microsoft.NuGet',
    #  'MSPACKAGE_Microsoft.Identity.Client',
    'ExchangeOnlineManagement'
)
# Main function to establish or reuse a connection to a SharePoint site
function Install-M365Dependencies {
    param (
        [Parameter(Mandatory = $false)]
        [string]$DependencySet = 'All'
    )
    $M365POWERKIT_DEPENDENCIES | ForEach-Object {
        # if the module is prefixed with WINGET_, use winget install -e $_
        if ($_ -like 'WINGET_*') {
            $packageName = $_ -replace 'WINGET_', ''
            Write-Debug "$($MyInvocation.MyCommand.Name) -- Checking if $packageName package is installed..."
            if (-not (Get-Command -Name $packageName -ErrorAction SilentlyContinue)) {
                try {
                    Write-Debug "Installing $packageName package..."
                    winget install -e -id $packageName -q -v -s  | Out-Null
                    Write-Debug "$packageName package installed successfully - OK"
                    # Get the installed package executable path and add it to the PATH
                    Get-ChildItem -Path $env:LOCALAPPDATA\Microsoft\WinGet\Packages\$packageName* -Recurse -Filter '*.exe' | ForEach-Object {
                        $env:PATH += ";$($_.DirectoryName)"
                    }
                }
                catch {
                    Write-Error "Failed to install $packageName package"
                }
            }
            else {
                Write-Debug "$packageName package already installed"
            }
        }
        # if the module is prefixed with MSPACKAGE_, use Install-Package $_ -Scope CurrentUser
        elseif ($_ -like 'MSPACKAGE_*') {
            $packageName = $_ -replace 'MSPACKAGE_', ''
            Write-Debug "$($MyInvocation.MyCommand.Name) -- Checking if $packageName package is installed..."
            if (-not (Get-Package -Name $packageName)) {
                try {
                    Write-Debug "Installing $packageName package... these can take a while to install, several minutes."
                    Install-Package -Name $packageName -Scope CurrentUser -Force -ErrorAction Continue
                    Write-Debug "$packageName package installed successfully - OK"
                }
                catch {
                    Write-Error "Failed to install $packageName package"
                }
            }
            else {
                Write-Debug "$packageName package already installed"
            }
        }
        else {
            Write-Debug "$($MyInvocation.MyCommand.Name) -- Checking if $_ module is installed..."
            Import-Module -Name $_ -Force
            
            if (-not (Get-Module -Name $_)) {
                try {
                    if (-not (Get-Module -ListAvailable -Name $_)) {
                        Write-Debug "$_ module not found locally, installing..."
                        Write-Debug "Installing $_ module..."
                        Install-Module -Name $_ -Force -ErrorAction Continue
                        Write-Debug "$_ module installed successfully - OK"
                    }
                }
                catch {
                    Write-Error "Failed to install $_ module"
                }
            }
            else {
                Write-Debug "$_ module already imported"
            }
            if (-not (Get-Module -Name $_)) {
                Import-Module -Name $_ -ErrorAction Continue | Out-Null
            }
        }
    }
    $env:M365PowerKit_DependenciesInstalled = $true
}
# We can get the access token using the Microsoft.Identity.Client assembly after doing an interactive login
function Get-M365PKAccessToken {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string[]]$Scopes = $(Get-MgContext).Scopes
    )
    $THIS_DIR = $MyInvocation.PSScriptRoot
    Write-Debug "Expected M365PowerKit-Shared Dir: $THIS_DIR"
    # Latest Microsoft.Identity.Client dir inside $THIS_DIR
    $MIC_DIR = Get-ChildItem -Path $THIS_DIR -Filter 'Microsoft.Identity.Client.*' -Directory | Sort-Object -Property Name -Descending | Select-Object -First 1
    #"Microsoft.Identity.Client.4.66.1\lib\net6.0\Microsoft.Identity.Client.dll"   
    $EXPECTED_DLL_PATH = "$($MIC_DIR.FullName)\lib\net6.0\Microsoft.Identity.Client.dll"
    # Test if file exists
    if (-not (Test-Path -Path $EXPECTED_DLL_PATH)) {
        Write-Error "Could not find the required DLL at $EXPECTED_DLL_PATH"
    }

    Add-Type -Path $EXPECTED_DLL_PATH
    # Create a PublicClientApplication instance
    $clientApp = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithTenantId($TenantId).Build()

    try {
        # Acquire token interactively
        $result = $clientApp.AcquireTokenInteractive($scopes).ExecuteAsync().GetAwaiter().GetResult()
        $accessToken = $result.AccessToken
        Write-Output "SUCCESSS: Access Token: $accessToken"
    }
    catch {
        Write-Error "Failed to acquire token: $_"
    }
    return $accessToken
}
# Function: New-IPPSSession
# Description: This function creates a new Exchange Online PowerShell session.
function New-IPPSSession {
    # Check if there is an existing session
    if (!$env:M365PowerKitUPN) {
        Write-Error 'No UPN found in the environment variable M365PowerKitUPN'
    }
    else {
        try {
            if (-not $env:M365PowerKit_IPPSSession) {
                Write-Debug 'Starting New-IPPSSession...'
                $env:M365PowerKit_IPPSSession = Connect-IPPSSession -UserPrincipalName $env:M365PowerKitUPN
            }
            else {
                Write-Debug 'Reusing existing IPS session.'
            }
        }
        catch {
            Write-Debug 'Failed to create Exchange Online PowerShell session, see:'
            Write-Debug '   - https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'
            Write-Error 'Failed establish IPS session'
        }
        Write-Debug 'IPS session created successfully'
    }
    return $env:M365PowerKitUPN
}

# New EXO Session
function New-EXOSession {
    # Check if there is an existing session
    if (!$env:M365PowerKitUPN) {
        Write-Error 'No UPN found in the environment variable M365PowerKitUPN'
    }
    else {
        try {
            if (-not $env:M365PowerKit_EXOSession) {
                Write-Debug 'Starting New-EXOSession...'
                $env:M365PowerKit_EXOSession = Connect-ExchangeOnline -UserPrincipalName $env:M365PowerKitUPN            
            }
            else {
                Write-Debug 'Reusing existing EXO session.'
            }
        }
        catch {
            Write-Debug 'Failed to create Exchange Online PowerShell session, see:'
            Write-Debug '   - https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'
            Write-Error 'Failed establish EXO session'
        }       
    }
}


# Function to authenticate using APP ID and Secret
function New-OAUTH2Session {
    param (
        [Parameter(Mandatory = $false)]
        [string]$AppID,
        [Parameter(Mandatory = $false)]
        [string]$TenantID,
        [Parameter(Mandatory = $false)]
        [securestring]$AppSecret
    )
    if (-not $AppID -or -not $TenantID -or -not $AppSecret) {
        Write-Output 'Required parameters: -AppID, -TenantID, and -AppSecret'
        Write-Error 'Parameters missing'
    }
    $OAUTH2Session = @{
        AppID        = $AppID
        TenantID     = $TenantID
        ClientSecret = $AppSecret
    }


    $body = @{
        Grant_Type    = 'client_credentials'
        Scope         = 'https://graph.microsoft.com/.default'
        Client_Id     = $appid
        Client_Secret = $secret
    }
 
    $connection = Invoke-RestMethod `
        -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token `
        -Method POST `
        -Body $body
 
    $token = $connection.access_token

    $secureToken = ConvertTo-SecureString $token -AsPlainText -Force
 
    Connect-MgGraph -AccessToken $secureToken -NoWelcome
    return $OAUTH2Session
}

function Clear-M365PowerKitProfile {
    Get-ChildItem env:M365PowerKit* | ForEach-Object {
        Write-Debug "Removing environment variable: $_"
        Remove-Item "env:$($_.Name)" -ErrorAction Continue
    }
    Write-Debug 'All M365PowerKit environment variables removed, disconnecting sessions...'
    if ($env:M365PowerKit_IPPSSession) {
        Disconnect-IPPSSession -Connection $env:M365PowerKit_IPPSSession
        $env:M365PowerKit_IPPSSession = $null
    }
    if ($env:M365PowerKit_EXOSession) {
        Disconnect-ExchangeOnline -Connection $env:EXOSession
        $env:M365PowerKit_EXOSession = $null
    }
    if ($env:M365PowerKit_PnPConnection) {
        Disconnect-PnPOnline -Connection $env:M365PowerKit_PnPConnection
        $env:M365PowerKit_PnPConnection = $null
    }
    Write-Debug 'All M365PowerKit sessions disconnected'
}
    
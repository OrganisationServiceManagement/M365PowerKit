# Set error and debug preferences for module initialization
$ErrorActionPreference = 'SilentlyContinue'; $DebugPreference = 'Continue'

$readOnlyScopes = @(
    'User.Read', # Read basic profile
    'User.Read.All', # Read all users' full profiles
    'Group.Read.All', # Read all groups
    'Directory.Read.All', # Read directory data
    'Files.Read.All', # Read all files the user has access to
    'Sites.Read.All', # Read items in all site collections
    'Calendars.Read', # Read user calendars
    'Mail.Read', # Read user mail
    'Contacts.Read', # Read user contacts
    'Tasks.Read', # Read user tasks
    'Notes.Read.All', # Read all user notes (OneNote)
    'Presence.Read', # Read user presence information
    'People.Read', # Read users' relevant people information
    'Reports.Read.All', # Read all usage reports
    'Device.Read.All', # Read all devices
    'Policy.Read.All', # Read all policies
    'IdentityRiskEvent.Read.All', # Read identity risk event information
    'Application.Read.All', # Read all applications
    'RoleManagement.Read.Directory', # Read role-based access control roles and assignments
    'SecurityEvents.Read.All'   # Read security events in the organization
    'ManagedTenants.Read.All' # Read organization data
    'MultiTenantOrganization.Read.All' # Read organization data
    'Organization.Read.All' # Read organization data
    'APIConnectors.Read.All' # Read API connectors
)


# Required Modules
# Function to list all files recursively within a given folder with depth control
function Get-RecursiveFileList {
    param (
        [string]$DriveId,
        [string]$FolderId,
        [int]$Depth = 0, # 0 means unlimited depth
        [int]$CurrentDepth = 1
    )

    # Base API URL for listing children
    $baseUrl = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$FolderId/children"
    
    # Fetch the children items in the current folder
    $items = Invoke-MgGraphRequest -Method GET -Uri $baseUrl

    foreach ($item in $items.value) {
        # If the item is a folder and we are within the depth limit, recurse into it
        if ($item.folder) {
            if ($Depth -eq 0 -or $CurrentDepth -lt $Depth) {
                # Recursive call with increased depth
                Get-RecursiveFileList -DriveId $DriveId -FolderId $item.id -Depth $Depth -CurrentDepth ($CurrentDepth + 1)
            }
        }
        else {
            # If it's a file, output the file name and path
            Write-Output "File: $($item.name) - Path: $($item.parentReference.path)/$($item.name)"
        }
    }
}

# Function to list all folders recursively within a given folder with depth control
function Get-RecursiveFolderList {
    param (
        [string]$DriveId,
        [string]$FolderId,
        [int]$Depth = 0, # 0 means unlimited depth
        [int]$CurrentDepth = 1
    )

    # Base API URL for listing children
    $baseUrl = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$FolderId/children"
    
    # Fetch the children items in the current folder
    $items = Invoke-MgGraphRequest -Method GET -Uri $baseUrl

    foreach ($item in $items.value) {
        # If the item is a folder, output the folder name and path
        if ($item.folder) {
            Write-Output "Folder: $($item.name) - Path: $($item.parentReference.path)/$($item.name)"

            # If within depth limit, recurse further
            if ($Depth -eq 0 -or $CurrentDepth -lt $Depth) {
                Get-RecursiveFolderList -DriveId $DriveId -FolderId $item.id -Depth $Depth -CurrentDepth ($CurrentDepth + 1)
            }
        }
    }
}

function Get-PdfFilesFromSharePointRelativeUrl {
    param (
        [Parameter(Mandatory = $false)]
        [string]$RawSharePointUrl = $null, # = 'https://zoaksolutions.sharepoint.com/sites/ISMS/Shared%20Documents/SSG_Export/4%20-%20Artefacts%20and%20Templates/Incident%20and%20Problem%20Artefacts/',
        [Parameter(Mandatory = $false)]
        [string]$SiteID, # = 'https://zoaksolutions.sharepoint.com:/sites/ISMS',
        [Parameter(Mandatory = $false)]
        [string]$FOLDER_ID, #= 'Shared%20Documents%2FSSG_Export%2F4%20-%20Artefacts%20and%20Templates%2FIncident%20and%20Problem%20Artefacts%2F'  # The relative URL of the library or folder within the site
        [Parameter(Mandatory = $false)]
        [string]$FolderName, # = 'Shared%20Documents/SSG_Export/4%20-%20Artefacts%20and%20Templates/Incident%20and%20Problem%20Artefacts/'  # The relative URL of the library or folder within the site
        [Parameter(Mandatory = $false)]
        [string]$SEARCH_TERM,
        [Parameter(Mandatory = $false)]
        [string]$URL_ENCODED_SEARCH_TERM,
        [Parameter(Mandatory = $false)]
        [switch]$ExternalTenant = $false, # = 'zoaksolutions.sharepoint.com'
        # = 'pdf'
        [Parameter(Mandatory = $false)]
        [switch]$ForceReconnect # = $false
    )
    $OUTPUT_DIR = "$env:OSM_HOME + '\M356_output'"
    if (-not (Test-Path $OUTPUT_DIR)) {
        New-Item -ItemType Directory -Path $OUTPUT_DIR | Out-Null
    }
    $TIME_STAMP = Get-Date -Format 'yyyyMMdd_HHmmss'
    if ($ExternalTenant) {
        Write-Debug 'Connecting to Microsoft Graph for external tenant...'
        if ($RawSharePointUrl -match 'https://([^.]+)\.sharepoint\.com') {
            $TenantID = "$($matches[1]).sharepoint.com" # Capture group 1 contains the tenant name
        }
        else {
            Write-Error "$($MyInvocation.InvocationName): Unable to parse TenantID from RawSharePointUrl: $RawSharePointUrl"
        }
        Write-Debug "Attempting to connect to External Tenant: $TenantID"
        Connect-MgGraph -Scopes 'Sites.Read.All' -TenantId $TenantID
    }
    else {
        Write-Debug 'Connecting to Microsoft Graph for internal tenant...'
        try {
            # Attempt to get the current Graph context
            $CUR_CONTEXT = Get-MgContext -ErrorAction Stop
        }
        catch {
            Write-Debug "Get-MgContext failed: $($_.Exception.Message)"
            $CUR_CONTEXT = $null
        }
        #Get-MgUser -Top 1 -ErrorAction Stop | Out-Null
        if ($null -ne $CUR_CONTEXT -and (! $ForceReconnect)) {
            $CUR_CONTEXT | ConvertTo-Json -Depth 30 | Write-Debug
            Write-Debug 'Session is still active, reusing...'
        }
        else {
            try {
                Write-Debug 'Session is not active, attempting to reconnect...'
                $CUR_CONTEXT = Connect-MgGraph -Scopes $readOnlyScopes # Include required scopes
                $CUR_CONTEXT | ConvertTo-Json -Depth 30 | Write-Debug
                Write-Debug 'Session reconnected successfully. OKOK'
            }
            catch {
                Write-Debug "$($MyInvocation.InvocationName): $_"
                Write-Error "$($MyInvocation.InvocationName): Failed to connect to Microsoft Graph with Connect-MgGraph ($readOnlyScopes)."
            }
        }
    }

    #Shared%20Documents%2FSSG_Export%2F4%20-%20Artefacts%20and%20Templates%2FIncident%20and%20Problem%20Artefacts%2F
    #Shared%20Documents%2FPolicies%2C%20Standards%20%26%20Processes%2F3%20-%20Policy%20and%20Process%2FPublished%20policies
    if ($null -ne $RawSharePointUrl) {
        Write-Debug "Parsing Raw SharePoint URL: $RawSharePointUrl"
        # Remove 'https://' prefix and split the URL
        $UrlWithoutProtocol = $RawSharePointUrl -replace '^https?://', ''
        $HostnameFU = $UrlWithoutProtocol.Split('/')[0]
        $RelativeURL = $UrlWithoutProtocol -replace '^[^/]+/', '' # removes the hostname, leaving the relative path
        $SiteRelativePath = ($RelativeURL -match '^(sites/[^/]+)') ? $matches[1] : ''
        # Get $PathToRoot by removing $SiteRelativePath from $RelativeURL and remove the trailing slash
        $PathToRoot = $RelativeURL -replace "^$SiteRelativePath/", '' -replace '/$', ''
        # If PathToRoot is starts with /Shared%20Documents

        Write-Debug "Hostname: $HostnameFU"
        Write-Debug "Relative Path: $RelativeURL"

        # Example API endpoint construction
        $SiteDetailsUrl = "https://graph.microsoft.com/v1.0/sites/${HostnameFU}:/${SiteRelativePath}"
        $REST_RESPONSE_SITE = Invoke-MgGraphRequest -Method GET -Uri $SiteDetailsUrl
        #$REST_RESPONSE_SITE | ConvertTo-Json -Depth 30 | Write-Debug
        $SiteID = $REST_RESPONSE_SITE.id
        #Write-Debug "Site ID: $SiteID"
    }
    elseif ($SiteID) {
        #Write-Debug 'Site ID and Relative URL provided.'
        Write-Debug "Site ID: $SiteID"
    }
    else {
        Write-Debug "$($MyInvocation.InvocationName): Either RawSharePointUrl or SiteID and RelativeUrl must be provided."
        Write-Debug "Parameters provided were: $($PSBoundParameters.GetEnumerator() | ForEach-Object { "$($_.Key) = $($_.Value)" } -join ', ')"
        Write-Error "$($MyInvocation.InvocationName): Invalid parameters provided."
    }
    # GET https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/{relative-path}:/children
    # GET https://graph.microsoft.com/v1.0/drives/your-site-id/root:/Shared Documents/YourFolder:/children
    $DRIVE_URL = "https://graph.microsoft.com/v1.0/sites/$SiteID/drive"
    $DRIVE_RESPONSE = Invoke-MgGraphRequest -Method GET -Uri $DRIVE_URL

    if ($URL_ENCODED_SEARCH_TERM) {
        $DRIVE_SEARCH_URL = "https://graph.microsoft.com/v1.0/drives/$($DRIVE_RESPONSE.id)/root/search(q='$($URL_ENCODED_SEARCH_TERM)')"
    }
    elseif ($SEARCH_TERM) {
        # Encode the search term
        $URL_ENCODED_SEARCH_TERM = [System.Web.HttpUtility]::UrlEncode($SEARCH_TERM)
        $DRIVE_SEARCH_URL = "https://graph.microsoft.com/v1.0/drives/$($DRIVE_RESPONSE.id)/root/search(q='$($URL_ENCODED_SEARCH_TERM)')"
    }
    if ($DRIVE_SEARH_URL) {
        $DRIVE_SEARCH_RESULTS = Invoke-MgGraphRequest -Method GET -Uri $DRIVE_SEARCH_URL
        $SEARCH_META = if ($($DRIVE_RESPONSE.webUrl) -match 'https://([^.]+)\.sharepoint\.com') { return $matches[1] } else { return 'Unn0wn' }
        $DRIVE_SEARCH_RESULTS | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OUTPUT_DIR\M365DriveSearch-$SEARCH_META-$TIME_STAMP.json"
    }
    else {
        Write-Debug 'Listing all directories from the root of the drive...'
        $DRIVE_ROOT_ITEMS = "https://graph.microsoft.com/v1.0/drives/$($DRIVE_RESPONSE.id)/root/children"
        $ROOT_OBJECTS = Invoke-MgGraphRequest -Method GET -Uri $DRIVE_ROOT_ITEMS
        # List the "name" and "id" of each item in the root of the drive where the item is a folder
        $ROOT_OBJECTS.value | Where-Object { $_.folder } | ForEach-Object {
            Write-Output "Folder: $($_.name) - ID: $($_.id)"
        }
        $NEXT_FOLDER = Read-Host 'Enter the ID of the folder to list its contents (null to exit):'
        if (!$NEXT_FOLDER) {
            Write-Debug 'Exiting...'
            return $true
        }
        else {
            $OPERATION = Read-Host 'Enter "f - files" / "d - dirs" / a - all to list folders [d]/f/a:'
            if ($OPERATION -eq 'f') {
                Get-RecursiveFileList -DriveId $DRIVE_RESPONSE.id -FolderId $NEXT_FOLDER
            }
            elseif ($OPERATION -eq 'd') {
                Get-RecursiveFolderList -DriveId $DRIVE_RESPONSE.id -FolderId $NEXT_FOLDER
            }
            elseif ($OPERATION -eq 'a') {
                Get-RecursiveFileList -DriveId $DRIVE_RESPONSE.id -FolderId $NEXT_FOLDER
                Get-RecursiveFolderList -DriveId $DRIVE_RESPONSE.id -FolderId $NEXT_FOLDER
            }
            else {
                Write-Debug 'Exiting...'
                return $true
            }
        }
    }
    Write-Host 'Congratulations! You win.'
}
<#
.SYNOPSIS
    Moves files in a SharePoint Online site using PnP.PowerShell. when a file is moved in a document
    library, if there is an associated file in the PDF subdirectory, then that file will also be
    moved into the corresponding PDF subdirectory in the destination.

    For example:

    source/folder/PDF/MyFile.pdf
    source/folder/MyFile.docx

    if you ask to move source/folder/MyFile.docx to target/newfolder/MyFile.docx, then the script
    will first check for PDF/MyFile.pdf. If it exists, then target/newfolder/PDF will be created
    (if it doesn't already exist), and then source/folder/PDF/MyFile.pdf will be moved to target/newfolder/PDF/MyFile.pdf.

    This is done in order to preserve the association between the original file and its PDF version, which is important
    sharing links and other functionality that relies on the PDF being in a predictable location relative to the original file

    If you ask to move a source pattern (like source/folder/*Training*.docx), then the script will get each file that matches
    the pattern and then do the move for each file, including checking for the associated PDF in the PDF subdirectory and moving it if it exists.

.DESCRIPTION
    This script helps move files within a SharePoint Online site. It includes checks to ensure
    PnP.PowerShell is installed and that you're connected to a SharePoint site.

.PARAMETER Check
    Verifies that PnP.PowerShell is installed and that you're connected to a SharePoint site.

.PARAMETER Connect
    Connects to a SharePoint site. Accepts multiple formats:
    - Tenant name only: 'contoso' → connects to https://contoso.sharepoint.com
    - Tenant/site: 'contoso/mysite' → connects to https://contoso.sharepoint.com/mysite
    - Full URL: 'https://contoso.sharepoint.com/mysite'
    This parameter is processed first, allowing you to combine it with other actions like -Check or -MoveFile.

.PARAMETER DeviceLogin
    Uses device code authentication instead of interactive browser authentication when connecting.
    This is useful when the default browser launches with the wrong profile because you can open
    the device login URL in the exact browser/profile you want.

.PARAMETER Tenant
    Optional tenant identifier for authentication. This is most useful with -DeviceLogin when
    PnP.PowerShell asks you to specify a tenant. Accepts tenant ID (GUID) or hostname
    (for example, 'contoso.onmicrosoft.com' or 'contoso.sharepoint.com').

.PARAMETER MoveFile
    The source file or pattern to move. Must include the folder path.
    Examples: "Shared Documents/file.docx" or "Shared Documents/*Training*.docx"
    Requires -Destination parameter.

.PARAMETER Destination
    The destination folder path where the file(s) will be moved. Required when -MoveFile is specified.

.PARAMETER List
    Lists files matching the specified pattern without moving them. Must include the folder path.
    Examples: "Shared Documents/Rules*.*" or "Documents/Subfolder/*.docx"

.PARAMETER Test
    Performs a dry run - shows what would be moved without actually moving files.
    When used with -Destination, shows where files would be moved to.
    When used without -Destination, simply lists the matching files (same as omitting -Destination).

.PARAMETER CheckAzure
    Path to a JSON configuration file for checking Azure Storage for mirrored PDF files.
    When provided, the script will also check if PDF files exist in Azure Blob Storage.
    When moving files, associated PDF blobs will be copied to the new location in Azure.
    The JSON file must have the following structure:
    {
        "StorageAccount": "<storage_account_name>",
        "StorageContainer": "<container_name>",
        "PathMappings": [
            [ "SPO_Path_Substring", "Azure_Path_Replacement" ]
        ]
    }
    Example: If a SharePoint URL is "https://tenant.sharepoint.com/umps/Reference/Helmets/PDF/file.pdf"
    and PathMappings has ["https://tenant.sharepoint.com/umps", "umps"], 
    the script will check for "umps/Reference/Helmets/PDF/file.pdf" in Azure Storage.
    Requires Az.Storage module and appropriate Azure authentication.

.PARAMETER DeleteOldAzureBlobs
    When specified with -CheckAzure, this switch will delete the old Azure blob after successfully
    copying it to the new location. Only applies when actually moving files (not in -Test or -List mode).
    Use with caution as this permanently removes the old blob.

.PARAMETER Verbose
    Shows detailed output including each PnP.PowerShell command before it's executed.
    This is useful for troubleshooting and understanding what operations the script is performing.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Check
    Checks if PnP.PowerShell is installed and if connected to a SharePoint site.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "contoso"
    Connects to https://contoso.sharepoint.com using interactive authentication.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "contoso/mysite"
    Connects to https://contoso.sharepoint.com/mysite using interactive authentication.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "https://contoso.sharepoint.com/ProjectSite"
    Connects to the specified SharePoint site URL using interactive authentication.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "contoso/mysite" -DeviceLogin
    Connects using device code authentication so you can complete sign-in in your chosen browser profile.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "contoso/mysite" -DeviceLogin -Tenant "contoso.onmicrosoft.com"
    Connects using device code authentication and an explicit tenant hostname.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "contoso/mysite" -Check
    Connects to SharePoint first, then verifies the environment setup.

.EXAMPLE
    .\Move-SPOFiles.ps1 -MoveFile "Shared Documents/MyFile.docx" -Destination "Shared Documents/Archive"
    Moves MyFile.docx to the Archive folder, including any associated PDF file.

.EXAMPLE
    .\Move-SPOFiles.ps1 -MoveFile "Shared Documents/*Training*.docx" -Destination "Shared Documents/Training"
    Moves all files matching the pattern to the Training folder, including associated PDFs.

.EXAMPLE
    .\Move-SPOFiles.ps1 -List "Shared Documents/*Training*.docx"
    Lists all files matching the pattern without moving them.

.EXAMPLE
    .\Move-SPOFiles.ps1 -MoveFile "Shared Documents/*Training*.docx" -Destination "Shared Documents/Training" -Test
    Shows what files would be moved without actually moving them (dry run).

.EXAMPLE
    .\Move-SPOFiles.ps1 -MoveFile "Shared Documents/file.docx" -Destination "Shared Documents/Archive" -Verbose
    Moves a file with verbose output showing each PnP.PowerShell command being executed.

.EXAMPLE
    .\Move-SPOFiles.ps1 -List "Shared Documents/*Training*.docx" -CheckAzure ".\azure-config.json"
    Lists all matching files and checks for associated PDFs in both SharePoint and Azure Storage.

.EXAMPLE
    .\Move-SPOFiles.ps1 -MoveFile "Shared Documents/file.docx" -Destination "Archive" -CheckAzure ".\azure-config.json" -Verbose
    Moves a file while checking both SharePoint and Azure Storage for associated PDFs, with verbose output.

.EXAMPLE
    .\Move-SPOFiles.ps1 -MoveFile "Shared Documents/*Training*.docx" -Destination "Archive" -CheckAzure ".\azure-config.json" -DeleteOldAzureBlobs
    Moves files and their associated PDFs. When PDFs are moved, copies the Azure blobs to new locations and deletes the old blobs.

.NOTES
    Author: Created for M365 Tools
    Date: February 16, 2026
    Last Updated: November 13, 2026
    Requires: PnP.PowerShell module
    Optional: Az.Storage module (for Azure blob checking and management)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [switch]$Check,
    
    [Parameter(Mandatory=$false)]
    [string]$Connect,

    [Parameter(Mandatory=$false)]
    [switch]$DeviceLogin,

    [Parameter(Mandatory=$false)]
    [string]$Tenant,
    
    [Parameter(Mandatory=$false)]
    [string]$MoveFile,
    
    [Parameter(Mandatory=$false)]
    [string]$Destination,
    
    [Parameter(Mandatory=$false)]
    [string]$List,
    
    [Parameter(Mandatory=$false)]
    [switch]$Test,
    
    [Parameter(Mandatory=$false)]
    [string]$CheckAzure,
    
    [Parameter(Mandatory=$false)]
    [switch]$DeleteOldAzureBlobs
)

# Script-level variable to store Azure configuration
$script:AzureConfig = $null
$script:AzureStorageContext = $null
$script:DeleteOldAzureBlobs = $false

function Load-AzureConfiguration {
    <#
    .SYNOPSIS
        Loads Azure storage configuration from JSON file.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            Write-Host "[X] Azure configuration file not found: $ConfigPath" -ForegroundColor Red
            return $null
        }
        
        $configContent = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        
        # Validate required properties
        if (-not $configContent.StorageAccount) {
            Write-Host "[X] Azure configuration missing 'StorageAccount' property" -ForegroundColor Red
            return $null
        }
        
        if (-not $configContent.StorageContainer) {
            Write-Host "[X] Azure configuration missing 'StorageContainer' property" -ForegroundColor Red
            return $null
        }
        
        if (-not $configContent.PathMappings) {
            Write-Host "[X] Azure configuration missing 'PathMappings' property" -ForegroundColor Red
            return $null
        }
        
        Write-Host "[✓] Azure configuration loaded successfully" -ForegroundColor Green
        Write-Host "    Storage Account: $($configContent.StorageAccount)" -ForegroundColor Gray
        Write-Host "    Container: $($configContent.StorageContainer)" -ForegroundColor Gray
        Write-Host "    Path Mappings: $($configContent.PathMappings.Count) mapping(s)" -ForegroundColor Gray
        
        return $configContent
    }
    catch {
        Write-Host "[X] Error loading Azure configuration: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Initialize-AzureStorageContext {
    <#
    .SYNOPSIS
        Creates and caches an Azure Storage context for reuse.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$AzureConfig
    )
    
    try {
        # Check if Az.Storage module is available
        $azStorageModule = Get-Module -ListAvailable -Name Az.Storage
        if (-not $azStorageModule) {
            Write-Host "[!] Az.Storage module not installed. Run: Install-Module Az.Storage" -ForegroundColor Yellow
            return $null
        }
        
        # Import the module if not already imported
        if (-not (Get-Module -Name Az.Storage)) {
            Write-Verbose "Importing Az.Storage module..."
            Import-Module Az.Storage -ErrorAction SilentlyContinue
        }
        
        $storageAccount = $AzureConfig.StorageAccount
        
        Write-Verbose "Creating storage context for account: $storageAccount"
        
        # Suppress progress bars to avoid flashing output
        $ProgressPreference = 'SilentlyContinue'
        
        # Try authenticated context first (required for write operations like copy/delete)
        $ctx = $null
        try {
            Write-Verbose "Attempting to create authenticated storage context"
            $ctx = New-AzStorageContext -StorageAccountName $storageAccount -UseConnectedAccount -ErrorAction Stop
            Write-Host "[✓] Azure Storage context created (authenticated)" -ForegroundColor Green
            Write-Verbose "Authenticated context created successfully"
        }
        catch {
            Write-Verbose "Authenticated context failed: $($_.Exception.Message)"
            # Fall back to anonymous if authenticated fails
            Write-Verbose "Trying anonymous context as fallback"
            $ctx = New-AzStorageContext -StorageAccountName $storageAccount -Anonymous -ErrorAction SilentlyContinue
            
            if ($null -ne $ctx) {
                Write-Host "[✓] Azure Storage context created (anonymous - read-only)" -ForegroundColor Green
                Write-Host "[!] Warning: Anonymous context may not support copy/delete operations" -ForegroundColor Yellow
            }
        }
        
        $ProgressPreference = 'Continue'
        
        return $ctx
    }
    catch {
        Write-Host "[X] Failed to create Azure Storage context: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Get-AzureBlobPath {
    <#
    .SYNOPSIS
        Converts a SharePoint URL to an Azure blob path using path mappings.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SharePointUrl,
        
        [Parameter(Mandatory=$true)]
        [object]$AzureConfig,
        
        [Parameter(Mandatory=$false)]
        [switch]$ShowDiagnostics
    )
    
    if ($ShowDiagnostics) {
        Write-Host "      [Azure] Mapping URL: $SharePointUrl" -ForegroundColor DarkGray
    }
    
    # Try each path mapping
    foreach ($mapping in $AzureConfig.PathMappings) {
        $spoPath = $mapping[0]
        $azurePath = $mapping[1]
        
        if ($ShowDiagnostics) {
            Write-Host "      [Azure] Testing mapping: '$spoPath' -> '$azurePath'" -ForegroundColor DarkGray
        }
        
        if ($SharePointUrl -like "*$spoPath*") {
            # Extract the part after the SPO path
            $urlParts = $SharePointUrl -split [regex]::Escape($spoPath), 2
            if ($urlParts.Count -eq 2) {
                $relativePath = $urlParts[1].TrimStart('/')
                $blobPath = "$azurePath/$relativePath"
                if ($ShowDiagnostics) {
                    Write-Host "      [Azure] ✓ Matched! Blob path: $blobPath" -ForegroundColor DarkGray
                }
                Write-Verbose "Mapped SharePoint URL to Azure blob: $blobPath"
                return $blobPath
            }
        }
    }
    
    if ($ShowDiagnostics) {
        Write-Host "      [Azure] ✗ No mapping matched" -ForegroundColor DarkGray
    }
    Write-Verbose "No path mapping found for: $SharePointUrl"
    return $null
}

function Copy-AzureBlob {
    <#
    .SYNOPSIS
        Copies a blob from one location to another in Azure Storage.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourceBlobPath,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationBlobPath,
        
        [Parameter(Mandatory=$true)]
        [object]$AzureConfig
    )
    
    try {
        if ($null -eq $script:AzureStorageContext) {
            Write-Verbose "Azure Storage context not initialized"
            return $false
        }
        
        $container = $AzureConfig.StorageContainer
        
        Write-Verbose "Copying Azure blob: $SourceBlobPath -> $DestinationBlobPath"
        
        # Suppress progress bars
        $ProgressPreference = 'SilentlyContinue'
        
        # Check if source blob exists
        $sourceBlob = Get-AzStorageBlob -Blob $SourceBlobPath -Container $container -Context $script:AzureStorageContext -ErrorAction SilentlyContinue
        
        if ($null -eq $sourceBlob) {
            Write-Verbose "Source blob not found: $SourceBlobPath"
            $ProgressPreference = 'Continue'
            return $false
        }
        
        # Start the copy operation
        $copyResult = Start-AzStorageBlobCopy -SrcBlob $SourceBlobPath -SrcContainer $container -DestBlob $DestinationBlobPath -DestContainer $container -Context $script:AzureStorageContext -ErrorAction Stop
        
        # Wait for copy to complete
        $copyStatus = $copyResult | Get-AzStorageBlobCopyState
        while ($copyStatus.Status -eq 'Pending') {
            Start-Sleep -Milliseconds 500
            $copyStatus = Get-AzStorageBlob -Blob $DestinationBlobPath -Container $container -Context $script:AzureStorageContext | Get-AzStorageBlobCopyState
        }
        
        $ProgressPreference = 'Continue'
        
        if ($copyStatus.Status -eq 'Success') {
            Write-Verbose "Azure blob copied successfully"
            return $true
        }
        else {
            Write-Verbose "Azure blob copy failed with status: $($copyStatus.Status)"
            return $false
        }
    }
    catch {
        $ProgressPreference = 'Continue'
        Write-Verbose "Error copying Azure blob: $($_.Exception.Message)"
        return $false
    }
}

function Remove-AzureBlob {
    <#
    .SYNOPSIS
        Deletes a blob from Azure Storage.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$BlobPath,
        
        [Parameter(Mandatory=$true)]
        [object]$AzureConfig
    )
    
    try {
        if ($null -eq $script:AzureStorageContext) {
            Write-Verbose "Azure Storage context not initialized"
            return $false
        }
        
        $container = $AzureConfig.StorageContainer
        
        Write-Verbose "Deleting Azure blob: $BlobPath"
        
        # Suppress progress bars
        $ProgressPreference = 'SilentlyContinue'
        
        $result = Remove-AzStorageBlob -Blob $BlobPath -Container $container -Context $script:AzureStorageContext -ErrorAction Stop
        
        $ProgressPreference = 'Continue'
        
        Write-Verbose "Azure blob deleted successfully"
        return $true
    }
    catch {
        $ProgressPreference = 'Continue'
        Write-Verbose "Error deleting Azure blob: $($_.Exception.Message)"
        return $false
    }
}

function Test-AzureBlobExists {
    <#
    .SYNOPSIS
        Checks if a blob exists in Azure Storage using cached context.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$BlobPath,
        
        [Parameter(Mandatory=$true)]
        [object]$AzureConfig,
        
        [Parameter(Mandatory=$false)]
        [switch]$ShowWarnings,
        
        [Parameter(Mandatory=$false)]
        [switch]$ShowDiagnostics
    )
    
    if ($ShowDiagnostics) {
        Write-Host "      [Azure] Checking blob: $BlobPath" -ForegroundColor DarkGray
    }
    
    try {
        # Use cached context if available
        if ($null -eq $script:AzureStorageContext) {
            if ($ShowWarnings) {
                Write-Host "    [!] Azure Storage context not initialized" -ForegroundColor Yellow
            }
            Write-Verbose "Azure Storage context is null"
            return $false
        }
        
        $container = $AzureConfig.StorageContainer
        
        if ($ShowDiagnostics) {
            Write-Host "      [Azure] Querying container '$container'..." -ForegroundColor DarkGray
        }
        
        Write-Verbose "Checking Azure blob: $BlobPath in container: $container"
        
        # Suppress progress bars to avoid flashing output
        $ProgressPreference = 'SilentlyContinue'
        $blob = Get-AzStorageBlob -Blob $BlobPath -Container $container -Context $script:AzureStorageContext -ErrorAction SilentlyContinue
        $ProgressPreference = 'Continue'
        
        if ($null -ne $blob) {
            if ($ShowDiagnostics) {
                Write-Host "      [Azure] ✓ Blob exists!" -ForegroundColor DarkGray
            }
            Write-Verbose "Blob exists in Azure: $BlobPath"
            return $true
        }
        else {
            if ($ShowDiagnostics) {
                Write-Host "      [Azure] ✗ Blob not found" -ForegroundColor DarkGray
            }
            Write-Verbose "Blob not found in Azure: $BlobPath"
            return $false
        }
    }
    catch {
        if ($ShowWarnings) {
            Write-Host "    [!] Azure error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        if ($ShowDiagnostics) {
            Write-Host "      [Azure] ✗ Error: $($_.Exception.Message)" -ForegroundColor DarkGray
        }
        Write-Verbose "Error checking Azure blob: $($_.Exception.Message)"
        return $false
    }
}

function Test-PnPModule {
    <#
    .SYNOPSIS
        Checks if PnP.PowerShell module is installed and imported.
    #>
    
    Write-Host "`n=== Checking PnP.PowerShell Module ===" -ForegroundColor Cyan
    
    # Check if module is installed
    $module = Get-Module -ListAvailable -Name PnP.PowerShell
    
    if ($null -eq $module) {
        Write-Host "[X] PnP.PowerShell module is NOT installed." -ForegroundColor Red
        Write-Host "`nTo install PnP.PowerShell, run the following command:" -ForegroundColor Yellow
        Write-Host "    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force" -ForegroundColor White
        Write-Host "`nNote: You may need to run PowerShell as Administrator or use -Scope CurrentUser" -ForegroundColor Gray
        return $false
    }
    
    Write-Host "[✓] PnP.PowerShell module is installed (Version: $($module[0].Version))" -ForegroundColor Green
    
    # Check if module is imported
    $importedModule = Get-Module -Name PnP.PowerShell
    
    if ($null -eq $importedModule) {
        Write-Host "[!] PnP.PowerShell module is not imported. Importing now..." -ForegroundColor Yellow
        try {
            Import-Module PnP.PowerShell -ErrorAction Stop
            Write-Host "[✓] PnP.PowerShell module imported successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "[X] Failed to import PnP.PowerShell module: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "[✓] PnP.PowerShell module is imported" -ForegroundColor Green
    }
    
    return $true
}

function Test-PnPConnection {
    <#
    .SYNOPSIS
        Checks if connected to a SharePoint site.
    #>
    
    Write-Host "`n=== Checking PnP Connection ===" -ForegroundColor Cyan
    
    try {
        Write-Verbose "Executing: Get-PnPConnection"
        $connection = Get-PnPConnection -ErrorAction Stop
        
        if ($null -eq $connection) {
            Write-Host "[X] Not connected to a SharePoint site." -ForegroundColor Red
            Write-Host "`nTo connect to a SharePoint site, run:" -ForegroundColor Yellow
            Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/yoursite' -Interactive" -ForegroundColor White
            Write-Host "`nOr use other authentication methods:" -ForegroundColor Yellow
            Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/yoursite' -DeviceLogin" -ForegroundColor White
            Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/yoursite' -UseWebLogin" -ForegroundColor White
            return $false
        }
        
        Write-Host "[✓] Connected to SharePoint site: $($connection.Url)" -ForegroundColor Green
        Write-Host "    Connection Type: $($connection.ConnectionType)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Write-Host "[X] Not connected to a SharePoint site." -ForegroundColor Red
        Write-Host "`nTo connect to a SharePoint site, run:" -ForegroundColor Yellow
        Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/yoursite' -Interactive" -ForegroundColor White
        Write-Host "`nOr use other authentication methods:" -ForegroundColor Yellow
        Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/yoursite' -DeviceLogin" -ForegroundColor White
        Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/yoursite' -UseWebLogin" -ForegroundColor White
        return $false
    }
}

function Show-ClientIdInstructions {
    <#
    .SYNOPSIS
        Displays instructions for registering or configuring PnP client ID.
    #>
    
    Write-Host "`n╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║  Client ID Registration Required                                  ║" -ForegroundColor Yellow
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    
    Write-Host "`nThe PnP.PowerShell app needs to be registered in your Entra ID tenant." -ForegroundColor Yellow
    Write-Host "`nFollow these steps:" -ForegroundColor Cyan
    
    Write-Host "`n1. Go to: " -NoNewline -ForegroundColor White
    Write-Host "https://entra.microsoft.com" -ForegroundColor Blue
    
    Write-Host "`n2. Navigate to: " -NoNewline -ForegroundColor White
    Write-Host "App Registrations" -ForegroundColor Blue
    
    Write-Host "`n3. Click on " -NoNewline -ForegroundColor White
    Write-Host "'All applications' " -NoNewline -ForegroundColor Blue
    Write-Host "and search for: " -NoNewline -ForegroundColor White
    Write-Host "'PnP.PowerShell'" -ForegroundColor Blue
    
    Write-Host "`n" -NoNewline
    Write-Host "═══ SCENARIO A: PnP.PowerShell App EXISTS ═══" -ForegroundColor Magenta
    Write-Host "`nIf you find the PnP.PowerShell app:" -ForegroundColor White
    Write-Host "   1. Click on it and copy the " -NoNewline -ForegroundColor Gray
    Write-Host "Application (client) ID" -ForegroundColor Cyan
    Write-Host "   2. Run this command in PowerShell:" -ForegroundColor Gray
    Write-Host "`n      Set-PnPManagedAppId -ClientId '<paste-your-client-id-here>'" -ForegroundColor Green
    Write-Host "`n      Example:" -ForegroundColor Gray
    Write-Host "      Set-PnPManagedAppId -ClientId '12345678-1234-1234-1234-123456789abc'" -ForegroundColor White
    
    Write-Host "`n" -NoNewline
    Write-Host "═══ SCENARIO B: PnP.PowerShell App DOES NOT EXIST ═══" -ForegroundColor Magenta
    Write-Host "`nIf PnP.PowerShell is not registered:" -ForegroundColor White
    Write-Host "   1. Run this command to register it:" -ForegroundColor Gray
    Write-Host "`n      Register-PnPEntraIDAppForInteractiveLogin ``" -ForegroundColor Green
    Write-Host "          -ApplicationName 'PnP.PowerShell' ``" -ForegroundColor Green
    Write-Host "          -Tenant '<yourtenant>.onmicrosoft.com' ``" -ForegroundColor Green
    Write-Host "          -Interactive" -ForegroundColor Green
    Write-Host "`n      Example:" -ForegroundColor Gray
    Write-Host "      Register-PnPEntraIDAppForInteractiveLogin ``" -ForegroundColor White
    Write-Host "          -ApplicationName 'PnP.PowerShell' ``" -ForegroundColor White
    Write-Host "          -Tenant 'contoso.onmicrosoft.com' ``" -ForegroundColor White
    Write-Host "          -Interactive" -ForegroundColor White
    Write-Host "`n   2. You will need Global Administrator or Application Administrator rights" -ForegroundColor Gray
    Write-Host "   3. After registration, the client ID will be set automatically" -ForegroundColor Gray
    
    Write-Host "`n" -NoNewline
    Write-Host "After completing either scenario, try connecting again." -ForegroundColor Cyan
    Write-Host ""
}

function Connect-ToPnP {
    <#
    .SYNOPSIS
        Connects to a SharePoint site with error handling.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,

        [Parameter(Mandatory=$false)]
        [switch]$UseDeviceLogin,

        [Parameter(Mandatory=$false)]
        [string]$Tenant
    )
    
    Write-Host "`n=== Connecting to SharePoint ===" -ForegroundColor Cyan
    Write-Host "Site URL: $SiteUrl" -ForegroundColor Gray
    
    # First ensure module is available
    $moduleOk = Test-PnPModule
    if (-not $moduleOk) {
        return $false
    }
    
    try {
        if ($UseDeviceLogin) {
            $tenantToUse = $Tenant

            if ([string]::IsNullOrWhiteSpace($tenantToUse)) {
                try {
                    $tenantToUse = ([Uri]$SiteUrl).Host
                }
                catch {
                    Write-Host "`n[X] Could not determine tenant from site URL '$SiteUrl'." -ForegroundColor Red
                    Write-Host "    Please try again with -Tenant '<tenant-id-or-hostname>'." -ForegroundColor Yellow
                    return $false
                }
            }

            Write-Host "`nAttempting to connect using DeviceLogin authentication..." -ForegroundColor Yellow
            Write-Host "Tenant: $tenantToUse" -ForegroundColor Gray
            Write-Verbose "Executing: Connect-PnPOnline -Url '$SiteUrl' -DeviceLogin -Tenant '$tenantToUse'"
            Connect-PnPOnline -Url $SiteUrl -DeviceLogin -Tenant $tenantToUse -ErrorAction Stop
        }
        else {
            Write-Host "`nAttempting to connect using Interactive authentication..." -ForegroundColor Yellow
            Write-Verbose "Executing: Connect-PnPOnline -Url '$SiteUrl' -Interactive"
            Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
        }
        Write-Host "[✓] Successfully connected to SharePoint site!" -ForegroundColor Green
        
        # Verify connection
        Write-Verbose "Executing: Get-PnPConnection"
        $connection = Get-PnPConnection
        Write-Host "    Connected to: $($connection.Url)" -ForegroundColor Gray
        Write-Host "    Connection Type: $($connection.ConnectionType)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        
        # Check for client ID related errors
        if ($errorMessage -match "specify a valid client id" -or 
            $errorMessage -match "client id" -or
            $errorMessage -match "AADSTS700016" -or
            $errorMessage -match "Application with identifier.*was not found") {
            
            Write-Host "`n[X] Connection failed: Client ID issue detected" -ForegroundColor Red
            Show-ClientIdInstructions
        }
        else {
            Write-Host "`n[X] Connection failed: $errorMessage" -ForegroundColor Red
            Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
            Write-Host "  • Verify the URL is correct and accessible" -ForegroundColor Gray
            Write-Host "  • Ensure you have permissions to access the site" -ForegroundColor Gray
            Write-Host "  • If browser profile selection is causing sign-in issues, try: Connect-PnPOnline -Url '$SiteUrl' -DeviceLogin" -ForegroundColor Gray
            Write-Host "  • If prompted for tenant, retry with: -Tenant '<tenant-id-or-hostname>'" -ForegroundColor Gray
            Write-Host "  • Try using -UseWebLogin instead: Connect-PnPOnline -Url '$SiteUrl' -UseWebLogin" -ForegroundColor Gray
        }
        
        return $false
    }
}

function Invoke-EnvironmentCheck {
    <#
    .SYNOPSIS
        Performs complete environment check.
    #>
    
    Write-Host "`n╔════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  PnP.PowerShell Environment Check             ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════╝" -ForegroundColor Cyan
    
    $moduleOk = Test-PnPModule
    $connectionOk = Test-PnPConnection
    
    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    
    if ($moduleOk -and $connectionOk) {
        Write-Host "[✓] All checks passed! Environment is ready." -ForegroundColor Green
        return $true
    }
    else {
        Write-Host "[X] Some checks failed. Please address the issues above." -ForegroundColor Red
        return $false
    }
}

function Get-ServerRelativePath {
    <#
    .SYNOPSIS
        Returns a path relative to the server/tenant (includes the site's base path).
        This format is used by Move-PnPFile, Get-PnPFile, etc.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        $connection = Get-PnPConnection -ErrorAction Stop
        $siteUrl = $connection.Url
        
        # Extract the server-relative part from the site URL
        # e.g., https://tenant.sharepoint.com/mysite -> /mysite
        $uri = [System.Uri]$siteUrl
        $serverRelativeSiteUrl = $uri.AbsolutePath.TrimEnd('/')
        
        # If path already has the full site prefix, return as-is
        if ($Path.StartsWith($serverRelativeSiteUrl + '/') -or $Path -eq $serverRelativeSiteUrl) {
            return $Path
        }
        
        # Remove leading slash to avoid double slashes
        $cleanPath = $Path.TrimStart('/')
        
        # Combine with the site's server-relative URL
        return "$serverRelativeSiteUrl/$cleanPath"
    }
    catch {
        return $Path
    }
}

function Get-SiteRelativePath {
    <#
    .SYNOPSIS
        Returns a path relative to the site (excludes site prefix).
        This format is used by Get-PnPFolderItem, Resolve-PnPFolder, etc.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        $connection = Get-PnPConnection -ErrorAction Stop
        $siteUrl = $connection.Url
        
        # Extract the server-relative part from the site URL
        $uri = [System.Uri]$siteUrl
        $serverRelativeSiteUrl = $uri.AbsolutePath.TrimEnd('/')
        
        # Remove leading slash for consistency
        $cleanPath = $Path.TrimStart('/')
        
        # If path starts with the site prefix, remove it
        if ($cleanPath.StartsWith($serverRelativeSiteUrl.TrimStart('/') + '/')) {
            $sitePrefix = $serverRelativeSiteUrl.TrimStart('/') + '/'
            return $cleanPath.Substring($sitePrefix.Length)
        }
        
        # If path already doesn't have the site prefix, return it
        return $cleanPath
    }
    catch {
        return $Path.TrimStart('/')
    }
}

function Get-NormalizedPath {
    <#
    .SYNOPSIS
        Deprecated: Use Get-ServerRelativePath or Get-SiteRelativePath instead.
        Defaults to Get-ServerRelativePath for backward compatibility.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    return Get-ServerRelativePath -Path $Path
}

function Move-SPOFileWithPDF {
    <#
    .SYNOPSIS
        Moves a file and its associated PDF if it exists.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Microsoft.SharePoint.Client.File]$File,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationFolder,
        
        [Parameter(Mandatory=$false)]
        [switch]$ListMode
    )
    
    try {
        # Validate File object has required properties
        if ($null -eq $File) {
            throw "File object is null"
        }
        
        # Ensure ServerRelativeUrl property is loaded (required for CSOM File objects)
        if ([string]::IsNullOrWhiteSpace($File.ServerRelativeUrl)) {
            Write-Verbose "ServerRelativeUrl not loaded, attempting to load property..."
            try {
                $null = Get-PnPProperty -ClientObject $File -Property ServerRelativeUrl -ErrorAction Stop
            }
            catch {
                Write-Verbose "Could not load ServerRelativeUrl property: $($_.Exception.Message)"
            }
        }
        
        if ([string]::IsNullOrWhiteSpace($File.Name)) {
            throw "File.Name is null or empty"
        }
        
        if ([string]::IsNullOrWhiteSpace($File.ServerRelativeUrl)) {
            throw "File.ServerRelativeUrl is null or empty (property may not be loaded)"
        }
        
        $fileName = $File.Name
        $fileServerRelativeUrl = $File.ServerRelativeUrl
        
        # Extract the folder path from the file's server relative URL
        $lastSlashIndex = $fileServerRelativeUrl.LastIndexOf('/')
        if ($lastSlashIndex -lt 0) {
            throw "Invalid file path format: $fileServerRelativeUrl (no folder separator found)"
        }
        $fileFolder = $fileServerRelativeUrl.Substring(0, $lastSlashIndex)
        
        # Normalize the destination folder with server-relative path (for Move-PnPFile)
        $normalizedDestination = Get-ServerRelativePath -Path $DestinationFolder
        
        # Get the base name without extension
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        
        Write-Host "  Processing: $fileName" -ForegroundColor Cyan
        
        # Check for associated PDF file
        $pdfPath = "$fileFolder/PDF/$baseName.pdf"
        $hasPDF = $false
        
        try {
            Write-Verbose "Executing: Get-PnPFile -Url '$pdfPath'"
            $pdfFile = Get-PnPFile -Url $pdfPath -ErrorAction SilentlyContinue
            if ($null -ne $pdfFile) {
                $hasPDF = $true
                Write-Host "    [✓] Found associated PDF in SharePoint: $pdfPath" -ForegroundColor Gray
            }
        }
        catch {
            # PDF doesn't exist, that's okay
        }
        
        # Check Azure storage if configured
        $hasAzurePDF = $false
        $azureBlobPath = $null
        if ($null -ne $script:AzureConfig) {
            # Get the connection to build the full SharePoint URL
            try {
                $connection = Get-PnPConnection -ErrorAction Stop
                # Get base URL (protocol://domain) without site path
                $uri = [System.Uri]$connection.Url
                $baseUrl = "$($uri.Scheme)://$($uri.Host)"
                $fullPdfUrl = "$baseUrl$pdfPath"
                
                Write-Verbose "Full PDF URL for Azure mapping: $fullPdfUrl"
                
                # Get the Azure blob path
                $azureBlobPath = Get-AzureBlobPath -SharePointUrl $fullPdfUrl -AzureConfig $script:AzureConfig 
                
                if ($null -ne $azureBlobPath) {
                    Write-Verbose "Checking Azure for PDF at: $azureBlobPath"
                    $hasAzurePDF = Test-AzureBlobExists -BlobPath $azureBlobPath -AzureConfig $script:AzureConfig -ShowWarnings
                    
                    if ($hasAzurePDF) {
                        Write-Host "    [✓] Found associated PDF in Azure: $azureBlobPath" -ForegroundColor Gray
                    }
                }
                else {
                    Write-Verbose "No Azure path mapping matched for: $fullPdfUrl"
                }
            }
            catch {
                Write-Host "    [!] Azure check error: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Verbose "Error checking Azure storage: $($_.Exception.Message)"
            }
        }
        
        # If in list mode, just show what would be moved
        if ($ListMode) {
            if ($hasPDF) {
                $pdfDestFolder = "$normalizedDestination/PDF"
                Write-Host "    [→] Would move PDF to: $pdfDestFolder/$baseName.pdf" -ForegroundColor Blue
            }
            $targetUrl = "$normalizedDestination/$fileName"
            Write-Host "    [→] Would move file to: $targetUrl" -ForegroundColor Blue
            return $true
        }
        
        # Move the PDF first if it exists
        if ($hasPDF) {
            $pdfDestFolder = "$normalizedDestination/PDF"
            
            # Ensure PDF destination folder exists
            try {
                Write-Verbose "Executing: Get-PnPFolder -Url '$pdfDestFolder'"
                $pdfFolder = Get-PnPFolder -Url $pdfDestFolder -ErrorAction SilentlyContinue
                if ($null -eq $pdfFolder) {
                    Write-Host "    [!] Creating PDF subfolder: $pdfDestFolder" -ForegroundColor Yellow
                    Write-Verbose "Executing: Add-PnPFolder -Name 'PDF' -Folder '$normalizedDestination'"
                    $null = Add-PnPFolder -Name "PDF" -Folder $normalizedDestination -ErrorAction Stop
                }
            }
            catch {
                Write-Host "    [!] Creating destination folder structure" -ForegroundColor Yellow
                # May need to create parent folders - Resolve-PnPFolder needs site-relative path
                $pdfDestFolderSiteRelative = Get-SiteRelativePath -Path $pdfDestFolder
                Write-Verbose "Executing: Resolve-PnPFolder -SiteRelativePath '$pdfDestFolderSiteRelative'"
                $null = Resolve-PnPFolder -SiteRelativePath $pdfDestFolderSiteRelative
            }
            
            # Move the PDF
            Write-Host "    [→] Moving PDF to: $pdfDestFolder/$baseName.pdf" -ForegroundColor Gray
            Write-Verbose "Executing: Move-PnPFile -SourceUrl '$pdfPath' -TargetUrl '$pdfDestFolder/$baseName.pdf' -Force"
            Move-PnPFile -SourceUrl $pdfPath -TargetUrl "$pdfDestFolder/$baseName.pdf" -Force -ErrorAction Stop
            Write-Host "    [✓] PDF moved successfully" -ForegroundColor Green
            
            # Copy Azure blob if it exists
            if ($hasAzurePDF -and $null -ne $azureBlobPath) {
                try {
                    # Build the destination blob path
                    $connection = Get-PnPConnection -ErrorAction Stop
                    $uri = [System.Uri]$connection.Url
                    $baseUrl = "$($uri.Scheme)://$($uri.Host)"
                    $destPdfPath = "$pdfDestFolder/$baseName.pdf"
                    $fullDestPdfUrl = "$baseUrl$destPdfPath"
                    
                    $destAzureBlobPath = Get-AzureBlobPath -SharePointUrl $fullDestPdfUrl -AzureConfig $script:AzureConfig
                    
                    if ($null -ne $destAzureBlobPath) {
                        Write-Host "    [→] Copying Azure blob to: $destAzureBlobPath" -ForegroundColor Gray
                        $copySuccess = Copy-AzureBlob -SourceBlobPath $azureBlobPath -DestinationBlobPath $destAzureBlobPath -AzureConfig $script:AzureConfig
                        
                        if ($copySuccess) {
                            Write-Host "    [✓] Azure blob copied successfully" -ForegroundColor Green
                            
                            # Delete old blob if requested
                            if ($script:DeleteOldAzureBlobs) {
                                Write-Verbose "DeleteOldAzureBlobs is TRUE - proceeding with deletion"
                                Write-Host "    [→] Deleting old Azure blob: $azureBlobPath" -ForegroundColor Gray
                                $deleteSuccess = Remove-AzureBlob -BlobPath $azureBlobPath -AzureConfig $script:AzureConfig
                                
                                if ($deleteSuccess) {
                                    Write-Host "    [✓] Old Azure blob deleted" -ForegroundColor Green
                                }
                                else {
                                    Write-Host "    [!] Failed to delete old Azure blob" -ForegroundColor Yellow
                                }
                            }
                        }
                        else {
                            Write-Host "    [!] Failed to copy Azure blob" -ForegroundColor Yellow
                        }
                    }
                }
                catch {
                    Write-Host "    [!] Error managing Azure blob: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Verbose "Azure blob management error: $($_.Exception.Message)"
                }
            }
        }
        
        # Move the main file
        $targetUrl = "$normalizedDestination/$fileName"
        Write-Host "    [→] Moving file to: $targetUrl" -ForegroundColor Gray
        Write-Verbose "Executing: Move-PnPFile -SourceUrl '$fileServerRelativeUrl' -TargetUrl '$targetUrl' -Force"
        Move-PnPFile -SourceUrl $fileServerRelativeUrl -TargetUrl $targetUrl -Force -ErrorAction Stop
        Write-Host "    [✓] File moved successfully" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Host "    [X] Error moving file: $($_.Exception.Message)" -ForegroundColor Red
        Write-Verbose "Error details: $($_.Exception.ToString())"
        Write-Verbose "File Name: $($File.Name)"
        Write-Verbose "File ServerRelativeUrl: $($File.ServerRelativeUrl)"
        return $false
    }
}

function Invoke-MoveFiles {
    <#
    .SYNOPSIS
        Processes file move operations based on source pattern.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePattern,
        
        [Parameter(Mandatory=$false)]
        [string]$DestinationFolder,
        
        [Parameter(Mandatory=$false)]
        [switch]$ListMode
    )
    
    if ($ListMode) {
        if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
            Write-Host "`n=== Listing Files ===" -ForegroundColor Cyan
        }
        else {
            Write-Host "`n=== Listing Files (Dry Run) ===" -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "`n=== Moving Files ===" -ForegroundColor Cyan
    }
    Write-Host "Source: $SourcePattern" -ForegroundColor Gray
    if (-not [string]::IsNullOrWhiteSpace($DestinationFolder)) {
        Write-Host "Destination: $DestinationFolder" -ForegroundColor Gray
    }
    
    # Ensure we're connected
    try {
        Write-Verbose "Executing: Get-PnPConnection"
        $connection = Get-PnPConnection -ErrorAction Stop
        if ($null -eq $connection) {
            Write-Host "`n[X] Not connected to SharePoint. Use -Connect first." -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "`n[X] Not connected to SharePoint. Use -Connect first." -ForegroundColor Red
        return $false
    }
    
    # Normalize the destination folder if provided (server-relative for Move-PnPFile operations)
    if (-not [string]::IsNullOrWhiteSpace($DestinationFolder)) {
        $DestinationFolder = Get-ServerRelativePath -Path $DestinationFolder
        Write-Verbose "Normalized destination path: $DestinationFolder"
    }
    
    # Ensure destination folder exists (skip in list mode or if no destination specified)
    if (-not $ListMode -and -not [string]::IsNullOrWhiteSpace($DestinationFolder)) {
        try {
            Write-Verbose "Executing: Get-PnPFolder -Url '$DestinationFolder'"
            $destFolder = Get-PnPFolder -Url $DestinationFolder -ErrorAction SilentlyContinue
            if ($null -eq $destFolder) {
                Write-Host "`n[!] Destination folder doesn't exist. Creating: $DestinationFolder" -ForegroundColor Yellow
                # Resolve-PnPFolder needs site-relative path
                $destFolderSiteRelative = Get-SiteRelativePath -Path $DestinationFolder
                Write-Verbose "Executing: Resolve-PnPFolder -SiteRelativePath '$destFolderSiteRelative'"
                $null = Resolve-PnPFolder -SiteRelativePath $destFolderSiteRelative
            }
        }
        catch {
            Write-Host "`n[X] Failed to create destination folder: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    
    # Check if source is a pattern or single file
    if ($SourcePattern -match '\*') {
        # Pattern - get matching files
        # Validate that pattern includes a folder path
        if ($SourcePattern -notmatch '/') {
            Write-Host "`n[X] Pattern must include a folder path." -ForegroundColor Red
            Write-Host "Example: 'Shared Documents/Rules*.*' or 'Documents/Subfolder/*Training*.docx'" -ForegroundColor Yellow
            Write-Host "You provided: '$SourcePattern'" -ForegroundColor Gray
            return $false
        }
        
        $folderPath = $SourcePattern.Substring(0, $SourcePattern.LastIndexOf('/'))
        $filePattern = $SourcePattern.Substring($SourcePattern.LastIndexOf('/') + 1)
        
        # Get-PnPFolderItem needs site-relative path
        $folderPath = Get-SiteRelativePath -Path $folderPath
        Write-Verbose "Normalized source folder path (site-relative): $folderPath"
        
        Write-Host "`nSearching for files matching pattern in: $folderPath" -ForegroundColor Yellow
        
        try {
            Write-Verbose "Executing: Get-PnPFolderItem -FolderSiteRelativeUrl '$folderPath' -ItemType File | Where-Object { Name -like '$filePattern' }"
            $files = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType File -ErrorAction Stop | 
                     Where-Object { $_.Name -like $filePattern }
            
            if ($files.Count -eq 0) {
                Write-Host "`n[!] No files found matching pattern: $filePattern" -ForegroundColor Yellow
                return $false
            }
            
            if ($ListMode) {
                if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
                    Write-Host "`nFound $($files.Count) file(s):" -ForegroundColor Green
                }
                else {
                    Write-Host "`nFound $($files.Count) file(s) that would be moved:" -ForegroundColor Green
                }
            }
            else {
                Write-Host "`nFound $($files.Count) file(s) to move" -ForegroundColor Green
            }
            
            $successCount = 0
            $failCount = 0
            
            foreach ($file in $files) {
                if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
                    # Just list the file
                    Write-Host "  $($file.Name)" -ForegroundColor White
                    Write-Verbose "  $($file.ServerRelativeUrl)" 
                    
                    # Check for associated PDF
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    $lastSlashIndex = $file.ServerRelativeUrl.LastIndexOf('/')
                    if ($lastSlashIndex -ge 0) {
                        $fileFolder = $file.ServerRelativeUrl.Substring(0, $lastSlashIndex)
                        $pdfPath = "$fileFolder/PDF/$baseName.pdf"
                        try {
                            Write-Verbose "Executing: Get-PnPFile -Url '$pdfPath'"
                            $pdfFile = Get-PnPFile -Url $pdfPath -ErrorAction SilentlyContinue
                            if ($null -ne $pdfFile) {
                                Write-Host "    [✓] Has associated PDF in SharePoint: $pdfPath" -ForegroundColor Gray
                            }
                        }
                        catch { }
                        
                        # Check Azure storage if configured
                        if ($null -ne $script:AzureConfig) {
                            try {
                                $connection = Get-PnPConnection -ErrorAction Stop
                                # Get base URL (protocol://domain) without site path
                                $uri = [System.Uri]$connection.Url
                                $baseUrl = "$($uri.Scheme)://$($uri.Host)"
                                $fullPdfUrl = "$baseUrl$pdfPath"
                                
                                Write-Verbose "Full PDF URL for Azure mapping: $fullPdfUrl"
                                
                                # Get the Azure blob path
                                $azureBlobPath = Get-AzureBlobPath -SharePointUrl $fullPdfUrl -AzureConfig $script:AzureConfig 
                                
                                if ($null -ne $azureBlobPath) {
                                    Write-Verbose "Checking Azure for PDF at: $azureBlobPath"
                                    $hasAzurePDF = Test-AzureBlobExists -BlobPath $azureBlobPath -AzureConfig $script:AzureConfig -ShowWarnings 
                                    
                                    if ($hasAzurePDF) {
                                        Write-Host "    [✓] Has associated PDF in Azure: $azureBlobPath" -ForegroundColor Gray
                                    }
                                }
                                else {
                                    Write-Verbose "No Azure path mapping matched for: $fullPdfUrl"
                                }
                            }
                            catch {
                                Write-Host "    [!] Azure check error: $($_.Exception.Message)" -ForegroundColor Yellow
                                Write-Verbose "Error checking Azure storage: $($_.Exception.Message)"
                            }
                        }
                    }
                    $successCount++
                }
                else {
                    if (Move-SPOFileWithPDF -File $file -DestinationFolder $DestinationFolder -ListMode:$ListMode) {
                        $successCount++
                    }
                    else {
                        $failCount++
                    }
                    Write-Host ""
                }
            }
            
            Write-Host "=== Summary ===" -ForegroundColor Cyan
            if ($ListMode -and [string]::IsNullOrWhiteSpace($DestinationFolder)) {
                Write-Host "[✓] Found: $successCount file(s)" -ForegroundColor Green
            }
            elseif ($ListMode) {
                Write-Host "[✓] Would move: $successCount file(s)" -ForegroundColor Green
                if ($failCount -gt 0) {
                    Write-Host "[X] Failed to process: $failCount file(s)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "[✓] Successfully moved: $successCount file(s)" -ForegroundColor Green
                if ($failCount -gt 0) {
                    Write-Host "[X] Failed to move: $failCount file(s)" -ForegroundColor Red
                }
            }
            
            return ($failCount -eq 0)
        }
        catch {
            Write-Host "`n[X] Error searching for files: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    else {
        # Single file - Get-PnPFile needs server-relative path
        $normalizedSource = Get-ServerRelativePath -Path $SourcePattern
        Write-Verbose "Normalized source file path (server-relative): $normalizedSource"
        
        try {
            Write-Verbose "Executing: Get-PnPFile -Url '$normalizedSource'"
            $fileObj = Get-PnPFile -Url $normalizedSource -ErrorAction Stop
            
            if ($null -eq $fileObj) {
                Write-Host "`n[X] File object is null after retrieval" -ForegroundColor Red
                return $false
            }
            
            # Explicitly load the ServerRelativeUrl property
            Write-Verbose "Loading ServerRelativeUrl property..."
            $null = Get-PnPProperty -ClientObject $fileObj -Property ServerRelativeUrl -ErrorAction Stop
            
            Write-Verbose "File retrieved successfully"
            Write-Verbose "  Type: $($fileObj.GetType().FullName)"
            Write-Verbose "  Name: $($fileObj.Name)"
            Write-Verbose "  ServerRelativeUrl: $($fileObj.ServerRelativeUrl)"
            
            Write-Host ""
            
            if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
                # Just list the file
                Write-Host "  $($fileObj.Name)" -ForegroundColor White
                Write-Host "    Location: $($fileObj.ServerRelativeUrl)" -ForegroundColor Gray
                
                # Check for associated PDF
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileObj.Name)
                $lastSlashIndex = $fileObj.ServerRelativeUrl.LastIndexOf('/')
                if ($lastSlashIndex -ge 0) {
                    $fileFolder = $fileObj.ServerRelativeUrl.Substring(0, $lastSlashIndex)
                    $pdfPath = "$fileFolder/PDF/$baseName.pdf"
                    try {
                        Write-Verbose "Executing: Get-PnPFile -Url '$pdfPath'"
                        $pdfFile = Get-PnPFile -Url $pdfPath -ErrorAction SilentlyContinue
                        if ($null -ne $pdfFile) {
                            Write-Host "    [✓] Has associated PDF in SharePoint: $pdfPath" -ForegroundColor Gray
                        }
                    }
                    catch { }
                    
                    # Check Azure storage if configured
                    if ($null -ne $script:AzureConfig) {
                        try {
                            $connection = Get-PnPConnection -ErrorAction Stop
                            # Get base URL (protocol://domain) without site path
                            $uri = [System.Uri]$connection.Url
                            $baseUrl = "$($uri.Scheme)://$($uri.Host)"
                            $fullPdfUrl = "$baseUrl$pdfPath"
                            
                            Write-Verbose "Full PDF URL for Azure mapping: $fullPdfUrl"
                            
                            # Get the Azure blob path
                            $azureBlobPath = Get-AzureBlobPath -SharePointUrl $fullPdfUrl -AzureConfig $script:AzureConfig
                            
                            if ($null -ne $azureBlobPath) {
                                Write-Verbose "Checking Azure for PDF at: $azureBlobPath"
                                $hasAzurePDF = Test-AzureBlobExists -BlobPath $azureBlobPath -AzureConfig $script:AzureConfig -ShowWarnings 
                                
                                if ($hasAzurePDF) {
                                    Write-Host "    [✓] Has associated PDF in Azure: $azureBlobPath" -ForegroundColor Gray
                                }
                            }
                            else {
                                Write-Verbose "No Azure path mapping matched for: $fullPdfUrl"
                            }
                        }
                        catch {
                            Write-Host "    [!] Azure check error: $($_.Exception.Message)" -ForegroundColor Yellow
                            Write-Verbose "Error checking Azure storage: $($_.Exception.Message)"
                        }
                    }
                }
                Write-Host ""
                Write-Host "[✓] File found" -ForegroundColor Green
                $result = $true
            }
            else {
                $result = Move-SPOFileWithPDF -File $fileObj -DestinationFolder $DestinationFolder -ListMode:$ListMode
                Write-Host ""
                
                if ($result) {
                    if ($ListMode) {
                        Write-Host "[✓] File list completed successfully" -ForegroundColor Green
                    }
                    else {
                        Write-Host "[✓] File move completed successfully" -ForegroundColor Green
                    }
                }
            }
            
            return $result
        }
        catch {
            Write-Host "`n[X] File not found or error accessing file: $($_.Exception.Message)" -ForegroundColor Red
            Write-Verbose "Full error details: $($_.Exception.ToString())"
            Write-Verbose "Error occurred at: $($_.InvocationInfo.PositionMessage)"
            return $false
        }
    }
}

# Main script logic

# Load Azure configuration if provided
if ($CheckAzure) {
    Write-Host "`n=== Loading Azure Configuration ===" -ForegroundColor Cyan
    $script:AzureConfig = Load-AzureConfiguration -ConfigPath $CheckAzure
    $script:DeleteOldAzureBlobs = $PSBoundParameters.ContainsKey('DeleteOldAzureBlobs')
    Write-Verbose "DeleteOldAzureBlobs switch value: $($script:DeleteOldAzureBlobs)"
    
    if ($null -eq $script:AzureConfig) {
        Write-Host "[X] Failed to load Azure configuration. Azure blob checking will be disabled." -ForegroundColor Yellow
    }
    else {
        # Check if Azure authentication is required (for MoveFile operations that need write access)
        # Skip authentication only for -List mode (read-only)
        if ($MoveFile -and -not $List) {
            Write-Verbose "MoveFile mode detected - checking Azure authentication for write operations"
            
            try {
                $azContext = Get-AzContext -ErrorAction SilentlyContinue
                
                if ($null -eq $azContext -or $null -eq $azContext.Account) {
                    Write-Host "[!] Azure authentication required for blob copy/delete operations" -ForegroundColor Yellow
                    Write-Host "[→] Running Connect-AzAccount..." -ForegroundColor Cyan
                    
                    # Suppress progress preference during connection
                    $savedProgress = $ProgressPreference
                    $ProgressPreference = 'SilentlyContinue'
                    
                    try {
                        # Suppress warnings during authentication
                        Connect-AzAccount -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                        $azContext = Get-AzContext -ErrorAction Stop
                        Write-Host "[✓] Azure authenticated as: $($azContext.Account)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "[X] Failed to authenticate to Azure: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "[!] Azure blob operations will be disabled" -ForegroundColor Yellow
                        $script:AzureConfig = $null
                    }
                    finally {
                        $ProgressPreference = $savedProgress
                    }
                }
                else {
                    Write-Host "[✓] Azure authenticated as: $($azContext.Account)" -ForegroundColor Green
                }
            }
            catch {
                Write-Verbose "Could not check Azure authentication status: $($_.Exception.Message)"
            }
        }
        else {
            Write-Verbose "List mode detected - skipping Azure authentication (anonymous context will be used for checking)"
        }
        
        # Initialize Azure Storage context once (only if config still valid)
        if ($null -ne $script:AzureConfig) {
            $script:AzureStorageContext = Initialize-AzureStorageContext -AzureConfig $script:AzureConfig
            
            if ($null -eq $script:AzureStorageContext) {
                Write-Host "[X] Failed to initialize Azure Storage context. Azure blob checking will be disabled." -ForegroundColor Yellow
                $script:AzureConfig = $null
            }
            else {
                # Display blob deletion mode
                Write-Verbose "Checking deletion mode: DeleteOldAzureBlobs = $($script:DeleteOldAzureBlobs)"
                if ($script:DeleteOldAzureBlobs) {
                    Write-Host "[!] Old Azure blobs will be deleted after copying" -ForegroundColor Yellow
                }
            }
        }
    }
}

# Handle -Connect first if provided
if ($Connect) {
    # Check if it's a full URL or just a tenant name
    $siteUrl = $Connect
    
    if ($Connect -notmatch '^https://') {
        # Not a full URL, parse the input
        # Remove .sharepoint.com if user included it
        $cleanInput = $Connect -replace '\.sharepoint\.com.*$', ''
        
        if ($cleanInput -match '^([^/]+)/(.+)$') {
            # Format: tenant/site
            $tenantName = $matches[1]
            $siteName = $matches[2]
            $siteUrl = "https://$tenantName.sharepoint.com/$siteName"
            Write-Host "Using tenant: $tenantName" -ForegroundColor Gray
            Write-Host "Using site: $siteName" -ForegroundColor Gray
            Write-Host "Connecting to: $siteUrl" -ForegroundColor Gray
        }
        else {
            # Just a tenant name
            $tenantName = $cleanInput
            $siteUrl = "https://$tenantName.sharepoint.com"
            Write-Host "Using tenant: $tenantName" -ForegroundColor Gray
            Write-Host "Connecting to: $siteUrl" -ForegroundColor Gray
        }
    }
    elseif ($Connect -notmatch '^https://.*\.sharepoint\.com') {
        Write-Host "[X] Invalid SharePoint URL format." -ForegroundColor Red
        Write-Host "Expected formats:" -ForegroundColor Yellow
        Write-Host "  - Tenant name: 'contoso'" -ForegroundColor Gray
        Write-Host "  - Tenant/site: 'contoso/mysite'" -ForegroundColor Gray
        Write-Host "  - Full URL: 'https://contoso.sharepoint.com/mysite'" -ForegroundColor Gray
        exit 1
    }
    
    $result = Connect-ToPnP -SiteUrl $siteUrl -UseDeviceLogin:$DeviceLogin -Tenant $Tenant
    
    if (-not $result) {
        exit 1
    }
    
    Write-Host "`n[✓] Connection successful! You can now use PnP cmdlets." -ForegroundColor Green
    
    # If no other actions specified, exit successfully
    if (-not $Check -and -not $MoveFile -and -not $List) {
        exit 0
    }
}

# Handle other actions
if ($Check) {
    $result = Invoke-EnvironmentCheck
    
    if ($result) {
        exit 0
    }
    else {
        exit 1
    }
}
elseif ($MoveFile) {
    # Validate that Destination is provided
    if ([string]::IsNullOrWhiteSpace($Destination)) {
        Write-Host "[X] -Destination parameter is required when using -MoveFile" -ForegroundColor Red
        Write-Host "`nUsage: .\Move-SPOFiles.ps1 -MoveFile '<source>' -Destination '<destination>'" -ForegroundColor Yellow
        Write-Host "To list files, use: .\Move-SPOFiles.ps1 -List '<pattern>'" -ForegroundColor Yellow
        exit 1
    }
    
    # Determine if we're in test mode (-Test flag) or move mode
    if ($Test) {
        # Dry run - show what would be moved
        $result = Invoke-MoveFiles -SourcePattern $MoveFile -DestinationFolder $Destination -ListMode
    }
    else {
        # Actually move files
        $result = Invoke-MoveFiles -SourcePattern $MoveFile -DestinationFolder $Destination
    }
    
    if ($result) {
        exit 0
    }
    else {
        exit 1
    }
}
elseif ($List) {
    # List files matching pattern
    $result = Invoke-MoveFiles -SourcePattern $List -ListMode
    
    if ($result) {
        exit 0
    }
    else {
        exit 1
    }
}
elseif (-not $Connect) {
    # Only show help if no parameters were provided at all
    Write-Host "No action specified. Available options:" -ForegroundColor Yellow
    Write-Host "  -Check                    Verify environment setup" -ForegroundColor Gray
    Write-Host "  -Connect <tenant|tenant/site|url>   Connect to a SharePoint site" -ForegroundColor Gray
    Write-Host "  -DeviceLogin              Use device code login (use with -Connect)" -ForegroundColor Gray
    Write-Host "  -Tenant <tenantId|hostname>   Optional auth tenant (especially with -DeviceLogin)" -ForegroundColor Gray
    Write-Host "  -List <pattern>           List files matching pattern" -ForegroundColor Gray
    Write-Host "  -MoveFile <src> -Destination <dest>   Move file(s) to destination" -ForegroundColor Gray
    Write-Host "  -MoveFile <src> -Destination <dest> -Test   Preview files that would be moved (dry run)" -ForegroundColor Gray
    Write-Host "  -Verbose                  Show detailed output including PnP commands" -ForegroundColor Gray
    Write-Host "`nExamples:" -ForegroundColor Cyan
    Write-Host "  .\Move-SPOFiles.ps1 -Check" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso' -DeviceLogin" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso' -DeviceLogin -Tenant 'contoso.onmicrosoft.com'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso/mysite'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'https://contoso.sharepoint.com/mysite'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso/mysite' -Check" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -List 'Shared Documents/*Training*.docx'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -MoveFile 'Shared Documents/file.docx' -Destination 'Shared Documents/Archive'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso/mysite' -MoveFile 'Shared Documents/*Training*.docx' -Destination 'Shared Documents/Training'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -MoveFile 'Shared Documents/*Training*.docx' -Destination 'Shared Documents/Training' -Test" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -MoveFile 'Shared Documents/file.docx' -Destination 'Archive' -Verbose" -ForegroundColor White
}

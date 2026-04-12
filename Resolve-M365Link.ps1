<#
.SYNOPSIS
    Resolves an anonymous Microsoft 365 sharing link to its real file path using Microsoft Graph API.

.DESCRIPTION
    This script takes a Microsoft 365 anonymous sharing link and uses the Microsoft Graph API
    to resolve it to the actual file path and metadata.

.PARAMETER Domain
    The SharePoint domain to filter links. Can be just the root name (e.g., 'contoso' for 'contoso.sharepoint.com').
    Only links matching this domain will be processed.

.PARAMETER ShareLink
    The anonymous M365 sharing link to resolve.

.PARAMETER InputList
    Path to a text file containing one sharing link per line. All links will be resolved.

.PARAMETER InputCsv
    Path to a CSV file with a 'Link URL' column containing sharing links. All links will be resolved.

.PARAMETER OutputFile
    Path to a CSV file where results will be saved (only used with -InputList).

.PARAMETER OutputCsv
    Path to a CSV file where results will be saved with original data preserved (only used with -InputCsv).

.PARAMETER TenantId
    The Azure AD Tenant ID for authentication.

.PARAMETER ClientId
    The Application (Client) ID for Microsoft Graph authentication.

.PARAMETER ClientSecret
    The Client Secret for app authentication (if using app-only authentication).

.PARAMETER UseInteractive
    Use interactive authentication instead of app-only authentication.

.PARAMETER UseAzCli
    Use Azure CLI credentials (requires 'az login' to be run first).

.PARAMETER VerifyAllLinks
    Verify that links not matching the domain are still valid by attempting to access them.
    This adds a "Link Status" column showing whether each non-matching link is accessible.

.PARAMETER TestMax
    Limit processing to the first N matching links and first N non-matching links (for testing).
    For example, -TestMax 5 will process at most 5 domain-matching links and verify at most 5 non-matching links.
    Valid range: 1-1000.

.PARAMETER RecheckExternalAfter
    Number of days before re-checking a previously verified external link (non-domain matching). Default is 7 days.
    When an external link is verified via HTTP, the result is cached. If the same link is encountered 
    again within RecheckExternalAfter days, the cached result is used instead of making a new HTTP request.
    Valid range: 1-365.

.PARAMETER RecheckDomainAfter
    Number of days before re-resolving a previously processed domain-matching link. Default is 5 days.
    When a domain link is resolved via Graph API, the result is cached. If the same link is encountered 
    again within RecheckDomainAfter days, the cached result is used instead of making a new API call.
    Valid range: 1-365.

.PARAMETER UseCacheFile
    Path to a JSON file for persistent cache storage. If the file exists, it will be loaded to
    pre-populate the cache. After processing, the cache will be saved back to this file.
    The cache stores both Graph API resolution results (for domain-matching links) and HTTP 
    verification results (for non-matching links when -VerifyAllLinks is used).
    This allows results to be reused across multiple script runs, significantly reducing API calls.

.PARAMETER All
    Display all file and site information. By default, only the file path is shown.

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -ShareLink "https://contoso.sharepoint.com/:w:/g/..." -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-secret"

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -ShareLink "https://contoso.sharepoint.com/:w:/g/..." -UseInteractive

.EXAMPLE
    az login
    .\Resolve-M365Link.ps1 -Domain "contoso" -ShareLink "https://contoso.sharepoint.com/:w:/g/..." -UseAzCli

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputList "links.txt" -UseInteractive

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputList "links.txt" -OutputFile "results.csv" -UseInteractive

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputCsv "data.csv" -OutputCsv "data-resolved.csv" -UseInteractive

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputCsv "data.csv" -OutputCsv "data-resolved.csv" -UseInteractive -VerifyAllLinks
    Resolves contoso links and verifies all other links for validity.

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputCsv "data.csv" -OutputCsv "test-output.csv" -UseInteractive -TestMax 5
    Test mode: Resolves only the first 5 contoso links and skips verification of other links.

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputCsv "data.csv" -OutputCsv "test-output.csv" -UseInteractive -VerifyAllLinks -TestMax 10
    Test mode: Resolves the first 10 contoso links and verifies the first 10 non-matching links.

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputCsv "data.csv" -OutputCsv "data-resolved.csv" -UseInteractive -VerifyAllLinks -UseCacheFile "link-cache.json"
    Resolves matching links and verifies non-matching links, saving all results to a cache file. On subsequent runs, 
    cached resolutions (within RecheckDomainAfter threshold) and verifications (within RecheckExternalAfter threshold) are reused.

.EXAMPLE
    .\Resolve-M365Link.ps1 -Domain "contoso" -InputCsv "data.csv" -OutputCsv "data-resolved.csv" -UseInteractive -VerifyAllLinks -UseCacheFile "link-cache.json" -RecheckDomainAfter 2 -RecheckExternalAfter 30
    Uses a 2-day cache for domain links and 30-day cache for external links. Domain links processed within 2 days use cached results; external links verified within 30 days use cached results.

.NOTES
    Requires Microsoft.Graph PowerShell module.
    For app-only authentication, the app needs Files.Read.All or Sites.Read.All permissions.
    
    Troubleshooting:
    - If you get "Microsoft.Identity.Client" errors with -UseInteractive, try:
      1. Update the module: Update-Module Microsoft.Graph -Force
      2. Restart PowerShell
      3. Or use -UseAzCli instead (requires: az login)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,
    
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$ShareLink,
    
    [Parameter(Mandatory = $false)]
    [string]$InputList,
    
    [Parameter(Mandatory = $false)]
    [string]$InputCsv,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputCsv,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [switch]$UseInteractive,
    
    [Parameter(Mandatory = $false)]
    [switch]$UseAzCli,
    
    [Parameter(Mandatory = $false)]
    [switch]$VerifyAllLinks,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$TestMax,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$RecheckExternalAfter = 7,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$RecheckDomainAfter = 5,
    
    [Parameter(Mandatory = $false)]
    [string]$UseCacheFile,
    
    [Parameter(Mandatory = $false)]
    [switch]$All
)

# Check if Microsoft.Graph.Authentication module is installed
$authModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1

if (-not $authModule) {
    Write-Host "Microsoft.Graph.Authentication module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
        Write-Host "Module installed successfully. Please restart your PowerShell session." -ForegroundColor Green
        exit 0
    }
    catch {
        Write-Error "Failed to install Microsoft.Graph.Authentication module: $_"
        exit 1
    }
}

# Import required module  
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}
catch {
    # Check if this is a version mismatch error
    if ($_.Exception.Message -like "*Could not load type*Microsoft.Graph*") {
        Write-Error "Module version conflict detected. Run these commands to fix:"
        Write-Host "`n  Get-Module Microsoft.Graph* | Remove-Module -Force" -ForegroundColor Cyan
        Write-Host "  Uninstall-Module Microsoft.Graph.Authentication -AllVersions -Force" -ForegroundColor Cyan
        Write-Host "  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber" -ForegroundColor Cyan
        Write-Host "`nThen restart PowerShell and try again.`n" -ForegroundColor Yellow
    }
    else {
        Write-Error "Failed to import Microsoft.Graph.Authentication module: $_"
    }
    exit 1
}

# Function to canonicalize URLs for consistent cache lookups
function Get-CanonicalUrl {
    param(
        [string]$Url
    )
    
    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $Url
    }
    
    # Normalize the URL
    $canonical = $Url.Trim()
    
    # Encode spaces as %20 (in case they're encoded differently or not at all)
    $canonical = $canonical.Replace(' ', '%20')
    
    # Remove trailing slash (but preserve if it's just the protocol://domain/)
    if ($canonical -match '^(https?://[^/]+)/$') {
        # Don't remove trailing slash if it's just protocol://domain/
        # Keep as is
    }
    elseif ($canonical.EndsWith('/')) {
        # Remove trailing slash from paths
        $canonical = $canonical.TrimEnd('/')
    }
    
    return $canonical
}

# Function to verify if a link is accessible
function Test-LinkValidity {
    param(
        [string]$Url,
        [hashtable]$Cache,
        [int]$RecheckAfterDays = 5
    )
    
    # Canonicalize URL for consistent cache lookups
    $canonicalUrl = Get-CanonicalUrl -Url $Url
    
    # Check cache first using canonical URL
    if ($Cache.ContainsKey($canonicalUrl)) {
        $cachedEntry = $Cache[$canonicalUrl]
        $daysSinceCheck = (Get-Date) - $cachedEntry.LastChecked
        
        if ($daysSinceCheck.TotalDays -lt $RecheckAfterDays) {
            # Cache is still valid, return cached result
            Write-Verbose "Using cached result for: $canonicalUrl (checked $([math]::Round($daysSinceCheck.TotalDays, 1)) days ago)"
            return @{
                IsValid = $cachedEntry.IsValid
                Status = $cachedEntry.Status + " (cached)"
                StatusCode = $cachedEntry.StatusCode
                FromCache = $true
            }
        }
        else {
            Write-Verbose "Cache expired for: $canonicalUrl (checked $([math]::Round($daysSinceCheck.TotalDays, 1)) days ago, threshold: $RecheckAfterDays days)"
        }
    }
    
    # Not in cache or cache expired, verify the link
    $result = $null
    
    try {
        # Try to get the headers with a HEAD request (faster than GET)
        $response = Invoke-WebRequest -Uri $canonicalUrl -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        $result = @{
            IsValid = $true
            Status = "Valid (HTTP $($response.StatusCode))"
            StatusCode = $response.StatusCode
            FromCache = $false
        }
    }
    catch {
        # If HEAD fails, try GET (some servers don't support HEAD)
        try {
            $response = Invoke-WebRequest -Uri $canonicalUrl -Method Get -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $result = @{
                IsValid = $true
                Status = "Valid (HTTP $($response.StatusCode))"
                StatusCode = $response.StatusCode
                FromCache = $false
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                $result = @{
                    IsValid = $false
                    Status = "Failed (HTTP $statusCode)"
                    StatusCode = $statusCode
                    FromCache = $false
                }
            }
            else {
                $result = @{
                    IsValid = $false
                    Status = "Failed: $($errorMsg.Substring(0, [Math]::Min(50, $errorMsg.Length)))"
                    StatusCode = $null
                    FromCache = $false
                }
            }
        }
    }
    
    # Update cache with the new result using canonical URL
    $Cache[$canonicalUrl] = @{
        Type = 'Verification'
        IsValid = $result.IsValid
        Status = $result.Status
        StatusCode = $result.StatusCode
        LastChecked = Get-Date
    }
    
    return $result
}

# Function to load verification cache from file
function Import-VerificationCache {
    param(
        [string]$Path
    )
    
    if (-not (Test-Path $Path)) {
        Write-Verbose "Cache file not found: $Path (will be created on save)"
        return @{}
    }
    
    try {
        Write-Host "Loading verification cache from: $Path" -ForegroundColor Cyan
        $jsonContent = Get-Content -Path $Path -Raw -ErrorAction Stop
        $cacheData = $jsonContent | ConvertFrom-Json -ErrorAction Stop
        
        # Convert JSON object back to hashtable with proper types
        $cache = @{}
        foreach ($property in $cacheData.PSObject.Properties) {
            $url = $property.Name
            $entry = $property.Value
            
            $cacheEntry = @{
                LastChecked = [DateTime]::Parse($entry.LastChecked)
            }
            
            # Determine cache entry type and restore appropriate fields
            if ($entry.PSObject.Properties.Name -contains 'WebUrl') {
                # This is a Graph API resolution result
                $cacheEntry.Type = 'Resolution'
                $cacheEntry.WebUrl = $entry.WebUrl
                $cacheEntry.Name = $entry.Name
                $cacheEntry.Status = $entry.Status
            }
            else {
                # This is an HTTP verification result
                $cacheEntry.Type = 'Verification'
                $cacheEntry.IsValid = $entry.IsValid
                $cacheEntry.Status = $entry.Status
                $cacheEntry.StatusCode = $entry.StatusCode
            }
            
            $cache[$url] = $cacheEntry
        }
        
        Write-Host "Loaded $($cache.Count) cached result(s)" -ForegroundColor Green
        return $cache
    }
    catch {
        Write-Warning "Failed to load cache file: $($_.Exception.Message)"
        Write-Warning "Starting with empty cache"
        return @{}
    }
}

# Function to save verification cache to file
function Export-VerificationCache {
    param(
        [hashtable]$Cache,
        [string]$Path
    )
    
    if ($Cache.Count -eq 0) {
        Write-Verbose "Cache is empty, skipping save"
        return
    }
    
    try {
        # Convert hashtable to JSON-friendly format
        $cacheObject = @{}
        foreach ($key in $Cache.Keys) {
            $entry = $Cache[$key]
            
            # Build cache entry based on type
            if ($entry.Type -eq 'Resolution') {
                $cacheObject[$key] = @{
                    WebUrl = $entry.WebUrl
                    Name = $entry.Name
                    Status = $entry.Status
                    LastChecked = $entry.LastChecked.ToString('o')
                }
            }
            else {
                # Verification type
                $cacheObject[$key] = @{
                    IsValid = $entry.IsValid
                    Status = $entry.Status
                    StatusCode = $entry.StatusCode
                    LastChecked = $entry.LastChecked.ToString('o')
                }
            }
        }
        
        $jsonContent = $cacheObject | ConvertTo-Json -Depth 10
        $jsonContent | Set-Content -Path $Path -Encoding UTF8 -ErrorAction Stop
        
        Write-Host "Saved verification cache to: $Path ($($Cache.Count) entries)" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to save cache file: $($_.Exception.Message)"
    }
}

# Function to export CSV with retry logic
function Export-CsvWithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [object[]]$Data,
        
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$false)]
        [string]$Description = "results"
    )
    
    $currentPath = $Path
    $success = $false
    
    while (-not $success) {
        try {
            $Data | Export-Csv -Path $currentPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            $success = $true
            Write-Host "`n$Description exported to: $currentPath" -ForegroundColor Green
            return $currentPath
        }
        catch {
            Write-Host "`nFailed to export $Description to: $currentPath" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "`nPossible issues:" -ForegroundColor Cyan
            Write-Host "  - File may be open in another program (close it and try again)" -ForegroundColor Gray
            Write-Host "  - Directory may not exist or you may not have write permissions" -ForegroundColor Gray
            Write-Host "  - Path may be invalid" -ForegroundColor Gray
            
            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "  [R] Retry with the same path" -ForegroundColor White
            Write-Host "  [N] Provide a new path" -ForegroundColor White
            Write-Host "  [S] Skip and continue without saving" -ForegroundColor White
            
            $choice = Read-Host "`nYour choice (R/N/S)"
            
            switch ($choice.ToUpper()) {
                "R" {
                    Write-Host "Retrying..." -ForegroundColor Cyan
                    continue
                }
                "N" {
                    $newPath = Read-Host "Enter new file path"
                    if ([string]::IsNullOrWhiteSpace($newPath)) {
                        Write-Host "Invalid path. Retrying with original path..." -ForegroundColor Yellow
                    }
                    else {
                        $currentPath = $newPath
                        Write-Host "Attempting to save to: $currentPath" -ForegroundColor Cyan
                    }
                    continue
                }
                "S" {
                    Write-Warning "Skipping export. Results will not be saved to file."
                    return $null
                }
                default {
                    Write-Host "Invalid choice. Please enter R, N, or S." -ForegroundColor Yellow
                    continue
                }
            }
        }
    }
}

# Validate parameters
if (-not $ShareLink -and -not $InputList -and -not $InputCsv) {
    Write-Error "Please provide either -ShareLink, -InputList, or -InputCsv parameter."
    exit 1
}

$paramCount = ($ShareLink -ne "") + ($InputList -ne "") + ($InputCsv -ne "")
if ($paramCount -gt 1) {
    Write-Error "Please provide only one of: -ShareLink, -InputList, or -InputCsv."
    exit 1
}

if ($OutputFile -and -not $InputList) {
    Write-Error "-OutputFile can only be used with -InputList."
    exit 1
}

if ($OutputCsv -and -not $InputCsv) {
    Write-Error "-OutputCsv can only be used with -InputCsv."
    exit 1
}

# Read links from file if InputList is specified
$linksToProcess = @()
$csvData = $null
$isCsvMode = $false

if ($InputList) {
    if (-not (Test-Path $InputList)) {
        Write-Error "Input file not found: $InputList"
        exit 1
    }
    
    $linksToProcess = Get-Content $InputList | Where-Object { $_.Trim() -ne "" -and -not $_.StartsWith("#") }
    
    if ($linksToProcess.Count -eq 0) {
        Write-Error "No valid links found in input file."
        exit 1
    }
    
    Write-Host "Found $($linksToProcess.Count) link(s) to process." -ForegroundColor Green
}
elseif ($InputCsv) {
    if (-not (Test-Path $InputCsv)) {
        Write-Error "Input CSV file not found: $InputCsv"
        exit 1
    }
    
    $csvData = Import-Csv $InputCsv
    
    if (-not $csvData) {
        Write-Error "Failed to read CSV file or file is empty."
        exit 1
    }
    
    # Check if 'Link URL' column exists
    $firstRow = $csvData | Select-Object -First 1
    if (-not ($firstRow.PSObject.Properties.Name -contains "Link URL")) {
        Write-Error "CSV file must contain a 'Link URL' column."
        exit 1
    }
    
    $urlColumnName = "Link URL"
    Write-Host "Using column: $urlColumnName" -ForegroundColor Cyan
    
    $linksToProcess = $csvData | ForEach-Object { $_.$urlColumnName } | Where-Object { $_.Trim() -ne "" }
    
    if ($linksToProcess.Count -eq 0) {
        Write-Error "No valid links found in 'Source URL' column."
        exit 1
    }
    
    $isCsvMode = $true
    Write-Host "Found $($linksToProcess.Count) link(s) to process from CSV." -ForegroundColor Green
}
else {
    $linksToProcess = @($ShareLink)
}

# Filter links by domain
$originalLinkCount = $linksToProcess.Count
$filteredLinks = @()
$skippedLinks = @()

foreach ($link in $linksToProcess) {
    # Check if the link is a SharePoint URL
    if ($link -notlike "*sharepoint.com*") {
        $skippedLinks += $link
        continue
    }
    
    # Check if the link matches the domain
    # Support both full domain (root.sharepoint.com) and short form (root)
    if ($link -match "https?://([^/]+)") {
        $linkDomain = $matches[1]
        
        # Extract the root part of the link domain (first part before first dot)
        $linkRoot = $linkDomain -replace "\..*$", ""
        
        # Extract the root part of the specified domain (first part before first dot)
        $specifiedRoot = $Domain -replace "\..*$", ""
        
        # Check if link domain matches the specified domain (exact match or root match)
        if ($linkDomain -eq $Domain -or $linkRoot -eq $specifiedRoot) {
            $filteredLinks += $link
        }
        else {
            $skippedLinks += $link
        }
    }
    else {
        Write-Warning "Invalid URL format: $link"
        $skippedLinks += $link
    }
}

$linksToProcess = $filteredLinks

if ($skippedLinks.Count -gt 0) {
    Write-Host "Filtered out $($skippedLinks.Count) link(s) not matching domain '$Domain'" -ForegroundColor Yellow
}

if ($linksToProcess.Count -eq 0) {
    Write-Error "No links matching domain '$Domain' found to process."
    exit 1
}

# Apply TestMax limiting if specified
if ($TestMax -gt 0) {
    $originalMatchingCount = $linksToProcess.Count
    $originalSkippedCount = $skippedLinks.Count
    
    Write-Host "`nTest mode enabled: Limiting to first $TestMax link(s) per category" -ForegroundColor Cyan
    
    if ($linksToProcess.Count -gt $TestMax) {
        $linksToProcess = $linksToProcess | Select-Object -First $TestMax
        Write-Host "  - Limited matching links from $originalMatchingCount to $TestMax" -ForegroundColor Yellow
    }
    else {
        Write-Host "  - Processing all $($linksToProcess.Count) matching link(s) (less than $TestMax)" -ForegroundColor Gray
    }
    
    if ($skippedLinks.Count -gt $TestMax) {
        $skippedLinks = $skippedLinks | Select-Object -First $TestMax
        Write-Host "  - Limited non-matching links from $originalSkippedCount to $TestMax" -ForegroundColor Yellow
    }
    elseif ($skippedLinks.Count -gt 0) {
        Write-Host "  - Processing all $($skippedLinks.Count) non-matching link(s) (less than $TestMax)" -ForegroundColor Gray
    }
    Write-Host ""
}

if ($originalLinkCount -ne $linksToProcess.Count) {
    Write-Host "Processing $($linksToProcess.Count) of $originalLinkCount link(s) matching domain '$Domain'" -ForegroundColor Green
}

# Load verification cache if UseCacheFile is specified
$verificationCache = @{}
if ($UseCacheFile) {
    $verificationCache = Import-VerificationCache -Path $UseCacheFile
}

try {
    # Authenticate to Microsoft Graph
    Write-Host "Authenticating to Microsoft Graph..." -ForegroundColor Cyan
    
    if ($UseAzCli) {
        # Check if Azure CLI is available
        $azCliAvailable = Get-Command az -ErrorAction SilentlyContinue
        if (-not $azCliAvailable) {
            Write-Error "Azure CLI (az) not found. Please install Azure CLI or use a different authentication method."
            exit 1
        }
        
        # Check if logged in to Azure CLI
        Write-Host "Checking Azure CLI login status..." -ForegroundColor Cyan
        $azAccount = az account show 2>$null | ConvertFrom-Json
        if (-not $azAccount) {
            Write-Error "Not logged in to Azure CLI. Please run 'az login' first."
            exit 1
        }
        
        Write-Host "Using Azure CLI account: $($azAccount.user.name)" -ForegroundColor Green
        
        # Get access token for Microsoft Graph from Azure CLI
        Write-Host "Getting access token from Azure CLI..." -ForegroundColor Cyan
        $tokenResponse = az account get-access-token --resource https://graph.microsoft.com | ConvertFrom-Json
        
        if (-not $tokenResponse.accessToken) {
            Write-Error "Failed to get access token from Azure CLI."
            exit 1
        }
        
        # Connect to Microsoft Graph using the access token
        $secureToken = ConvertTo-SecureString -String $tokenResponse.accessToken -AsPlainText -Force
        Connect-MgGraph -AccessToken $secureToken -NoWelcome
    }
    elseif ($UseInteractive) {
        # Interactive authentication
        try {
            Connect-MgGraph -Scopes "Files.Read.All", "Sites.Read.All" -NoWelcome -ErrorAction Stop
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Error "Interactive authentication failed: $errorMessage"
            Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
            
            # Check for specific Microsoft.Identity.Client error
            if ($errorMessage -like "*Microsoft.Identity.Client*" -or $errorMessage -like "*IMsalSFHttpClientFactory*") {
                Write-Host "1. This appears to be a module version compatibility issue." -ForegroundColor Cyan
                Write-Host "   Try updating the Microsoft.Graph module:" -ForegroundColor Cyan
                Write-Host "   Update-Module Microsoft.Graph -Force" -ForegroundColor White
                Write-Host "   Then restart PowerShell and try again`n" -ForegroundColor Cyan
                Write-Host "2. Alternatively, try using -UseAzCli instead (run 'az login' first)" -ForegroundColor Cyan
            }
            else {
                Write-Host "1. Make sure you have a web browser available" -ForegroundColor Cyan
                Write-Host "2. Try using -UseAzCli instead (run 'az login' first)" -ForegroundColor Cyan
                Write-Host "3. Check if you have the required permissions in your tenant" -ForegroundColor Cyan
            }
            Write-Host ""
            exit 1
        }
    }
    elseif ($TenantId -and $ClientId -and $ClientSecret) {
        # App-only authentication
        $SecureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecureClientSecret
        
        try {
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome -ErrorAction Stop
        }
        catch {
            Write-Error "App-only authentication failed: $_"
            exit 1
        }
    }
    else {
        Write-Error "Please specify an authentication method: -UseAzCli, -UseInteractive, or provide TenantId, ClientId, and ClientSecret."
        exit 1
    }
    
    # Verify authentication succeeded
    $context = Get-MgContext
    if (-not $context) {
        Write-Error "Authentication verification failed. Not connected to Microsoft Graph."
        exit 1
    }
    
    Write-Host "Authentication successful." -ForegroundColor Green
    
    # Process each link
    $results = @()
    $skippedResults = @()
    $linkIndex = 0
    $cachedResolutionCount = 0
    
    # Add skipped links to results for CSV mode
    if ($isCsvMode) {
        if ($VerifyAllLinks -and $skippedLinks.Count -gt 0) {
            Write-Host "`nVerifying $($skippedLinks.Count) non-matching link(s)..." -ForegroundColor Cyan
            $verifyIndex = 0
        }
        
        foreach ($skippedLink in $skippedLinks) {
            if ($VerifyAllLinks) {
                $verifyIndex++
                Write-Host "[$verifyIndex/$($skippedLinks.Count)] Verifying: $skippedLink" -ForegroundColor Gray
                $verifyResult = Test-LinkValidity -Url $skippedLink -Cache $verificationCache -RecheckAfterDays $RecheckExternalAfter
                
                $skippedResults += [PSCustomObject]@{
                    Link = $skippedLink
                    WebUrl = "NOT RESOLVED"
                    Name = "N/A"
                    Status = "Domain mismatch"
                    LinkStatus = $verifyResult.Status
                }
            }
            else {
                $skippedResults += [PSCustomObject]@{
                    Link = $skippedLink
                    WebUrl = "SKIPPED"
                    Name = "N/A"
                    Status = "Skipped: Domain mismatch"
                }
            }
        }
        
        if ($VerifyAllLinks -and $skippedLinks.Count -gt 0) {
            $validCount = ($skippedResults | Where-Object { $_.LinkStatus -like "Valid*" }).Count
            $failedCount = ($skippedResults | Where-Object { $_.LinkStatus -like "Failed*" }).Count
            $cachedCount = ($skippedResults | Where-Object { $_.LinkStatus -like "*(cached)" }).Count
            Write-Host "Verification complete: $validCount valid, $failedCount failed" -ForegroundColor Cyan
            if ($cachedCount -gt 0) {
                Write-Host "  ($cachedCount result(s) from cache, re-check threshold: $RecheckExternalAfter day(s))" -ForegroundColor Gray
            }
        }
    }
    
    foreach ($currentLink in $linksToProcess) {
        $linkIndex++
        
        if ($linksToProcess.Count -gt 1) {
            Write-Host "`n[$linkIndex/$($linksToProcess.Count)] Processing: $currentLink" -ForegroundColor Cyan
        }
        else {
            Write-Host "Resolving sharing link..." -ForegroundColor Cyan
        }
        
        # Check cache for this link first
        $canonicalLink = Get-CanonicalUrl -Url $currentLink
        $cachedResult = $null
        $usedCache = $false
        
        if ($verificationCache.ContainsKey($canonicalLink)) {
            $cachedEntry = $verificationCache[$canonicalLink]
            if ($cachedEntry.Type -eq 'Resolution') {
                $daysSinceCheck = (Get-Date) - $cachedEntry.LastChecked
                
                if ($daysSinceCheck.TotalDays -lt $RecheckDomainAfter) {
                    # Cache is still valid, use cached result
                    Write-Verbose "Using cached resolution for: $canonicalLink (resolved $([math]::Round($daysSinceCheck.TotalDays, 1)) days ago)"
                    $usedCache = $true
                    $cachedResolutionCount++
                    $cachedResult = @{
                        webUrl = $cachedEntry.WebUrl
                        name = $cachedEntry.Name
                    }
                }
                else {
                    Write-Verbose "Cache expired for: $canonicalLink (resolved $([math]::Round($daysSinceCheck.TotalDays, 1)) days ago, threshold: $RecheckDomainAfter days)"
                }
            }
        }
        
        try {
            if ($usedCache) {
                # Use cached response
                $response = $cachedResult
            }
            else {
                # Encode the sharing link for the Graph API
                # The sharing link needs to be base64-encoded and URL-safe
                $encodedUrl = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($currentLink))
                # Make it URL-safe by replacing characters
                $encodedUrl = $encodedUrl.TrimEnd('=').Replace('/', '_').Replace('+', '-')
                $shareId = "u!$encodedUrl"
                
                # Call Microsoft Graph API to get the shared item
                $uri = "https://graph.microsoft.com/v1.0/shares/$shareId/driveItem"
                
                $response = Invoke-MgGraphRequest -Uri $uri -Method GET
                
                # Cache the successful resolution
                $verificationCache[$canonicalLink] = @{
                    Type = 'Resolution'
                    WebUrl = $response.webUrl
                    Name = $response.name
                    Status = 'Success'
                    LastChecked = Get-Date
                }
            }
    
    # Construct the full path - only if not using cached data
    $fullPath = ""
    if (-not $usedCache -and $response.parentReference) {
        $siteId = $response.parentReference.siteId
        $driveId = $response.parentReference.driveId
        $itemId = $response.id
        
        # Try to get the full path from the parent reference
        if ($response.parentReference.path) {
            $path = $response.parentReference.path
            # Remove the /drive/root: prefix if present
            $path = $path -replace '^/drive/root:', ''
            $fullPath = "$path/$($response.name)"
        }
        
        # Get additional information if -All is specified
        if ($All) {
            # Get the site information for more context
            try {
                $siteUri = "https://graph.microsoft.com/v1.0/sites/$siteId"
                $siteInfo = Invoke-MgGraphRequest -Uri $siteUri -Method GET
            }
            catch {
                $siteInfo = $null
                Write-Warning "Could not retrieve site information: $_"
            }
            
            # Get the drive information
            try {
                $driveUri = "https://graph.microsoft.com/v1.0/drives/$driveId"
                $driveInfo = Invoke-MgGraphRequest -Uri $driveUri -Method GET
            }
            catch {
                $driveInfo = $null
                Write-Warning "Could not retrieve drive information: $_"
            }
        }
    }
    
            # Store result for batch processing
            if ($InputList) {
                $results += [PSCustomObject]@{
                    ViewLink = $currentLink
                    WebUrl = $response.webUrl
                    Name = $response.name
                    Status = "Success"
                }
            }
            elseif ($isCsvMode) {
                # For CSV mode, we'll update the original row data later
                $results += [PSCustomObject]@{
                    Link = $currentLink
                    WebUrl = $response.webUrl
                    Name = $response.name
                    Status = "Success"
                }
            }
            
            # Display output based on -All parameter
            if ($All) {
                # Show all information
                Write-Host "`n=== File Information ===" -ForegroundColor Green
                Write-Host "Name: $($response.name)" -ForegroundColor White
                Write-Host "ID: $($response.id)" -ForegroundColor White
                
                if ($response.webUrl) {
                    Write-Host "Web URL: $($response.webUrl)" -ForegroundColor White
                }
                
                Write-Host "`n=== Path Information ===" -ForegroundColor Green
                Write-Host "Site ID: $siteId" -ForegroundColor White
                Write-Host "Drive ID: $driveId" -ForegroundColor White
                Write-Host "Item ID: $itemId" -ForegroundColor White
                Write-Host "Relative Path: $fullPath" -ForegroundColor White
                
                if ($siteInfo) {
                    Write-Host "`n=== Site Information ===" -ForegroundColor Green
                    Write-Host "Site Name: $($siteInfo.displayName)" -ForegroundColor White
                    Write-Host "Site URL: $($siteInfo.webUrl)" -ForegroundColor White
                }
                
                if ($driveInfo) {
                    Write-Host "`n=== Drive Information ===" -ForegroundColor Green
                    Write-Host "Drive Name: $($driveInfo.name)" -ForegroundColor White
                    Write-Host "Drive Type: $($driveInfo.driveType)" -ForegroundColor White
                }
                
                # Additional metadata
                if ($response.size) {
                    $sizeKB = [math]::Round($response.size / 1KB, 2)
                    Write-Host "`n=== File Details ===" -ForegroundColor Green
                    Write-Host "Size: $sizeKB KB" -ForegroundColor White
                }
                
                if ($response.createdDateTime) {
                    Write-Host "Created: $($response.createdDateTime)" -ForegroundColor White
                }
                
                if ($response.lastModifiedDateTime) {
                    Write-Host "Modified: $($response.lastModifiedDateTime)" -ForegroundColor White
                }
                
                if ($linksToProcess.Count -eq 1) {
                    # Return the complete response object for single link
                    Write-Host "`n=== Complete Response ===" -ForegroundColor Green
                    return $response
                }
            }
            else {
                # Show only the web URL
                if ($usedCache) {
                    Write-Host "$($response.webUrl) (cached)" -ForegroundColor Gray
                }
                else {
                    Write-Host $response.webUrl
                }
            }
        }
        catch {
            Write-Warning "Failed to resolve link: $currentLink"
            Write-Warning "Error: $($_.Exception.Message)"
            
            if ($InputList) {
                $results += [PSCustomObject]@{
                    ViewLink = $currentLink
                    WebUrl = "ERROR"
                    Name = "N/A"
                    Status = "Failed: $($_.Exception.Message)"
                }
            }
            elseif ($isCsvMode) {
                $results += [PSCustomObject]@{
                    Link = $currentLink
                    WebUrl = "ERROR"
                    Name = "N/A"
                    Status = "Failed: $($_.Exception.Message)"
                }
            }
        }
    }
    
    # Display summary for batch processing
    if ($InputList -or $isCsvMode) {
        Write-Host "`n=== Processing Summary ===" -ForegroundColor Green
        Write-Host "Total links processed: $($results.Count)" -ForegroundColor White
        Write-Host "Successful: $(($results | Where-Object { $_.Status -eq 'Success' }).Count)" -ForegroundColor Green
        Write-Host "Failed: $(($results | Where-Object { $_.Status -ne 'Success' }).Count)" -ForegroundColor Red
        if ($cachedResolutionCount -gt 0) {
            Write-Host "  ($cachedResolutionCount resolution(s) from cache, re-check threshold: $RecheckDomainAfter day(s))" -ForegroundColor Gray
        }
        
        # Export to CSV if OutputFile is specified (InputList mode)
        if ($OutputFile) {
            Export-CsvWithRetry -Data $results -Path $OutputFile -Description "Results"
        }
        
        # Export to CSV with original data if OutputCsv is specified (InputCsv mode)
        if ($OutputCsv) {
            try {
                # Create enriched CSV with original data plus new columns
                $enrichedData = @()
                
                # Combine results and skipped results to match CSV order
                $allResults = @()
                $resultIndex = 0
                $skippedIndex = 0
                
                foreach ($csvRow in $csvData) {
                    $sourceUrl = $csvRow.$urlColumnName
                    if ($skippedLinks -contains $sourceUrl) {
                        $allResults += $skippedResults[$skippedIndex]
                        $skippedIndex++
                    }
                    else {
                        $allResults += $results[$resultIndex]
                        $resultIndex++
                    }
                }
                
                # Check if the columns already exist
                $firstRow = $csvData | Select-Object -First 1
                $hasWebUrl = $firstRow.PSObject.Properties.Name -contains "Web URL"
                $hasName = $firstRow.PSObject.Properties.Name -contains "Name"
                $hasStatus = $firstRow.PSObject.Properties.Name -contains "Status"
                $hasLinkStatus = $firstRow.PSObject.Properties.Name -contains "Link Status"
                
                for ($i = 0; $i -lt $csvData.Count; $i++) {
                    $originalRow = $csvData[$i]
                    $resultRow = $allResults[$i]
                    
                    # Create a new object with all original properties
                    $newRow = [PSCustomObject]@{}
                    
                    # Copy all original properties and update existing columns
                    foreach ($prop in $originalRow.PSObject.Properties) {
                        if ($prop.Name -eq "Web URL") {
                            # Update existing Web URL column
                            $newRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $resultRow.WebUrl
                        }
                        elseif ($prop.Name -eq "Name") {
                            # Update existing Name column
                            $newRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $resultRow.Name
                        }
                        elseif ($prop.Name -eq "Status") {
                            # Update existing Status column
                            $newRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $resultRow.Status
                        }
                        elseif ($prop.Name -eq "Link Status") {
                            # Update existing Link Status column (if VerifyAllLinks is used)
                            if ($VerifyAllLinks -and $resultRow.PSObject.Properties.Name -contains "LinkStatus") {
                                $newRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $resultRow.LinkStatus
                            }
                            else {
                                $newRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
                            }
                        }
                        else {
                            # Copy original value
                            $newRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
                        }
                    }
                    
                    # Add new columns if they don't exist
                    if (-not $hasWebUrl) {
                        $newRow | Add-Member -MemberType NoteProperty -Name "Web URL" -Value $resultRow.WebUrl
                    }
                    if (-not $hasName) {
                        $newRow | Add-Member -MemberType NoteProperty -Name "Name" -Value $resultRow.Name
                    }
                    if (-not $hasStatus) {
                        $newRow | Add-Member -MemberType NoteProperty -Name "Status" -Value $resultRow.Status
                    }
                    if (-not $hasLinkStatus -and $VerifyAllLinks) {
                        # Add Link Status column if VerifyAllLinks is enabled
                        if ($resultRow.PSObject.Properties.Name -contains "LinkStatus") {
                            $newRow | Add-Member -MemberType NoteProperty -Name "Link Status" -Value $resultRow.LinkStatus
                        }
                        else {
                            # For resolved links, leave Link Status empty
                            $newRow | Add-Member -MemberType NoteProperty -Name "Link Status" -Value ""
                        }
                    }
                    
                    $enrichedData += $newRow
                }
                
                # Determine description for export
                $exportDescription = if ($InputCsv -eq $OutputCsv) { "CSV file (in-place update)" } else { "Enriched results" }
                
                # Export with retry logic
                $finalPath = Export-CsvWithRetry -Data $enrichedData -Path $OutputCsv -Description $exportDescription
                
                # Update OutputCsv variable in case user provided a new path
                if ($finalPath -and $finalPath -ne $OutputCsv) {
                    Write-Host "Note: File was saved to a different location than originally specified." -ForegroundColor Yellow
                    $OutputCsv = $finalPath
                }
            }
            catch {
                Write-Warning "Failed to prepare enriched results: $_"
            }
        }
        
        if (-not $All -and $InputList) {
            Write-Host "`n=== Results ===" -ForegroundColor Green
            $results | Format-Table -AutoSize
        }
    }
}
catch {
    Write-Error "Failed to resolve sharing link: $_"
    Write-Error $_.Exception.Message
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd()
        Write-Error "Response: $responseBody"
    }
    exit 1
}
finally {
    # Save verification cache if UseCacheFile is specified
    if ($UseCacheFile -and $verificationCache -and $verificationCache.Count -gt 0) {
        Export-VerificationCache -Cache $verificationCache -Path $UseCacheFile
    }
    
    # Disconnect from Microsoft Graph if connected
    $context = Get-MgContext
    if ($context) {
        Disconnect-MgGraph | Out-Null
    }
}

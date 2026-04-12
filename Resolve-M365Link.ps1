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

if ($originalLinkCount -ne $linksToProcess.Count) {
    Write-Host "Processing $($linksToProcess.Count) of $originalLinkCount link(s) matching domain '$Domain'" -ForegroundColor Green
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
    
    # Add skipped links to results for CSV mode
    if ($isCsvMode) {
        foreach ($skippedLink in $skippedLinks) {
            $skippedResults += [PSCustomObject]@{
                Link = $skippedLink
                WebUrl = "SKIPPED"
                Name = "N/A"
                Status = "Skipped: Domain mismatch"
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
        
        try {
            # Encode the sharing link for the Graph API
            # The sharing link needs to be base64-encoded and URL-safe
            $encodedUrl = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($currentLink))
            # Make it URL-safe by replacing characters
            $encodedUrl = $encodedUrl.TrimEnd('=').Replace('/', '_').Replace('+', '-')
            $shareId = "u!$encodedUrl"
            
            # Call Microsoft Graph API to get the shared item
            $uri = "https://graph.microsoft.com/v1.0/shares/$shareId/driveItem"
            
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
    
    # Construct the full path
    $fullPath = ""
    if ($response.parentReference) {
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
                Write-Host $response.webUrl
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
        
        # Export to CSV if OutputFile is specified (InputList mode)
        if ($OutputFile) {
            try {
                $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                Write-Host "`nResults exported to: $OutputFile" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to export results to CSV: $_"
            }
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
                    
                    $enrichedData += $newRow
                }
                
                $enrichedData | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
                
                if ($InputCsv -eq $OutputCsv) {
                    Write-Host "`nCSV file updated in-place: $OutputCsv" -ForegroundColor Green
                }
                else {
                    Write-Host "`nEnriched results exported to: $OutputCsv" -ForegroundColor Green
                }
            }
            catch {
                Write-Warning "Failed to export enriched results to CSV: $_"
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
    # Disconnect from Microsoft Graph if connected
    $context = Get-MgContext
    if ($context) {
        Disconnect-MgGraph | Out-Null
    }
}

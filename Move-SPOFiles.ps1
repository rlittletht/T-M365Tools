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
    - Tenant/site: 'contoso/mysite' → connects to https://contoso.sharepoint.com/sites/mysite
    - Full URL: 'https://contoso.sharepoint.com/sites/mysite'
    This parameter is processed first, allowing you to combine it with other actions like -Check or -MoveFile.

.PARAMETER MoveFile
    The source file or pattern to move. Must include the folder path.
    Examples: "Shared Documents/file.docx" or "Shared Documents/*Training*.docx"

.PARAMETER Destination
    The destination folder path where the file(s) will be moved. Required when -MoveFile is specified.

.PARAMETER ListFiles
    Lists files matching the specified pattern without moving them. Must include the folder path.
    Takes the source file or pattern (same format as -MoveFile). 
    Examples: "Shared Documents/Rules*.*" or "Documents/Subfolder/*.docx"
    If -Destination is specified, shows where files would be moved to.
    If -Destination is omitted, simply lists the matching files.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Check
    Checks if PnP.PowerShell is installed and if connected to a SharePoint site.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "contoso"
    Connects to https://contoso.sharepoint.com using interactive authentication.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "contoso/mysite"
    Connects to https://contoso.sharepoint.com/sites/mysite using interactive authentication.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "https://contoso.sharepoint.com/sites/ProjectSite"
    Connects to the specified SharePoint site URL using interactive authentication.

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
    .\Move-SPOFiles.ps1 -ListFiles "Shared Documents/*Training*.docx"
    Lists all files matching the pattern.

.EXAMPLE
    .\Move-SPOFiles.ps1 -ListFiles "Shared Documents/*Training*.docx" -Destination "Shared Documents/Training"
    Lists files that would be moved without actually moving them (dry run).

.NOTES
    Author: Created for M365 Tools
    Date: February 16, 2026
    Requires: PnP.PowerShell module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [switch]$Check,
    
    [Parameter(Mandatory=$false)]
    [string]$Connect,
    
    [Parameter(Mandatory=$false)]
    [string]$MoveFile,
    
    [Parameter(Mandatory=$false)]
    [string]$Destination,
    
    [Parameter(Mandatory=$false)]
    [string]$ListFiles
)

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
        $connection = Get-PnPConnection -ErrorAction Stop
        
        if ($null -eq $connection) {
            Write-Host "[X] Not connected to a SharePoint site." -ForegroundColor Red
            Write-Host "`nTo connect to a SharePoint site, run:" -ForegroundColor Yellow
            Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/sites/yoursite' -Interactive" -ForegroundColor White
            Write-Host "`nOr use other authentication methods:" -ForegroundColor Yellow
            Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/sites/yoursite' -UseWebLogin" -ForegroundColor White
            return $false
        }
        
        Write-Host "[✓] Connected to SharePoint site: $($connection.Url)" -ForegroundColor Green
        Write-Host "    Connection Type: $($connection.ConnectionType)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Write-Host "[X] Not connected to a SharePoint site." -ForegroundColor Red
        Write-Host "`nTo connect to a SharePoint site, run:" -ForegroundColor Yellow
        Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/sites/yoursite' -Interactive" -ForegroundColor White
        Write-Host "`nOr use other authentication methods:" -ForegroundColor Yellow
        Write-Host "    Connect-PnPOnline -Url 'https://yourtenant.sharepoint.com/sites/yoursite' -UseWebLogin" -ForegroundColor White
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
        [string]$SiteUrl
    )
    
    Write-Host "`n=== Connecting to SharePoint ===" -ForegroundColor Cyan
    Write-Host "Site URL: $SiteUrl" -ForegroundColor Gray
    
    # First ensure module is available
    $moduleOk = Test-PnPModule
    if (-not $moduleOk) {
        return $false
    }
    
    try {
        Write-Host "`nAttempting to connect using Interactive authentication..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
        Write-Host "[✓] Successfully connected to SharePoint site!" -ForegroundColor Green
        
        # Verify connection
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
        $fileName = $File.Name
        $fileServerRelativeUrl = $File.ServerRelativeUrl
        $fileFolder = $File.ServerRelativeUrl.Substring(0, $File.ServerRelativeUrl.LastIndexOf('/'))
        
        # Get the base name without extension
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $extension = [System.IO.Path]::GetExtension($fileName)
        
        Write-Host "  Processing: $fileName" -ForegroundColor Cyan
        
        # Check for associated PDF file
        $pdfPath = "$fileFolder/PDF/$baseName.pdf"
        $hasPDF = $false
        
        try {
            $pdfFile = Get-PnPFile -Url $pdfPath -ErrorAction SilentlyContinue
            if ($null -ne $pdfFile) {
                $hasPDF = $true
                Write-Host "    [✓] Found associated PDF: $pdfPath" -ForegroundColor Gray
            }
        }
        catch {
            # PDF doesn't exist, that's okay
        }
        
        # If in list mode, just show what would be moved
        if ($ListMode) {
            if ($hasPDF) {
                $pdfDestFolder = "$DestinationFolder/PDF"
                Write-Host "    [→] Would move PDF to: $pdfDestFolder/$baseName.pdf" -ForegroundColor Blue
            }
            $targetUrl = "$DestinationFolder/$fileName"
            Write-Host "    [→] Would move file to: $targetUrl" -ForegroundColor Blue
            return $true
        }
        
        # Move the PDF first if it exists
        if ($hasPDF) {
            $pdfDestFolder = "$DestinationFolder/PDF"
            
            # Ensure PDF destination folder exists
            try {
                $pdfFolder = Get-PnPFolder -Url $pdfDestFolder -ErrorAction SilentlyContinue
                if ($null -eq $pdfFolder) {
                    Write-Host "    [!] Creating PDF subfolder: $pdfDestFolder" -ForegroundColor Yellow
                    $null = Add-PnPFolder -Name "PDF" -Folder $DestinationFolder -ErrorAction Stop
                }
            }
            catch {
                Write-Host "    [!] Creating destination folder structure" -ForegroundColor Yellow
                # May need to create parent folders
                $null = Resolve-PnPFolder -SiteRelativePath $pdfDestFolder
            }
            
            # Move the PDF
            Write-Host "    [→] Moving PDF to: $pdfDestFolder/$baseName.pdf" -ForegroundColor Gray
            Move-PnPFile -SourceUrl $pdfPath -TargetUrl "$pdfDestFolder/$baseName.pdf" -Force -ErrorAction Stop
            Write-Host "    [✓] PDF moved successfully" -ForegroundColor Green
        }
        
        # Move the main file
        $targetUrl = "$DestinationFolder/$fileName"
        Write-Host "    [→] Moving file to: $targetUrl" -ForegroundColor Gray
        Move-PnPFile -SourceUrl $fileServerRelativeUrl -TargetUrl $targetUrl -Force -ErrorAction Stop
        Write-Host "    [✓] File moved successfully" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Host "    [X] Error moving file: $($_.Exception.Message)" -ForegroundColor Red
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
            Write-Host "\n=== Listing Files ===" -ForegroundColor Cyan
        }
        else {
            Write-Host "\n=== Listing Files (Dry Run) ===" -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "\n=== Moving Files ===" -ForegroundColor Cyan
    }
    Write-Host "Source: $SourcePattern" -ForegroundColor Gray
    if (-not [string]::IsNullOrWhiteSpace($DestinationFolder)) {
        Write-Host "Destination: $DestinationFolder" -ForegroundColor Gray
    }
    
    # Ensure we're connected
    try {
        $connection = Get-PnPConnection -ErrorAction Stop
        if ($null -eq $connection) {
            Write-Host "\n[X] Not connected to SharePoint. Use -Connect first." -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "\n[X] Not connected to SharePoint. Use -Connect first." -ForegroundColor Red
        return $false
    }
    
    # Ensure destination folder exists (skip in list mode or if no destination specified)
    if (-not $ListMode -and -not [string]::IsNullOrWhiteSpace($DestinationFolder)) {
        try {
            $destFolder = Get-PnPFolder -Url $DestinationFolder -ErrorAction SilentlyContinue
            if ($null -eq $destFolder) {
                Write-Host "\n[!] Destination folder doesn't exist. Creating: $DestinationFolder" -ForegroundColor Yellow
                $null = Resolve-PnPFolder -SiteRelativePath $DestinationFolder
            }
        }
        catch {
            Write-Host "\n[X] Failed to create destination folder: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    
    # Check if source is a pattern or single file
    if ($SourcePattern -match '\*') {
        # Pattern - get matching files
        # Validate that pattern includes a folder path
        if ($SourcePattern -notmatch '/') {
            Write-Host "\n[X] Pattern must include a folder path." -ForegroundColor Red
            Write-Host "Example: 'Shared Documents/Rules*.*' or 'Documents/Subfolder/*Training*.docx'" -ForegroundColor Yellow
            Write-Host "You provided: '$SourcePattern'" -ForegroundColor Gray
            return $false
        }
        
        $folderPath = $SourcePattern.Substring(0, $SourcePattern.LastIndexOf('/'))
        $filePattern = $SourcePattern.Substring($SourcePattern.LastIndexOf('/') + 1)
        
        Write-Host "\nSearching for files matching pattern in: $folderPath" -ForegroundColor Yellow
        
        try {
            $files = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType File -ErrorAction Stop | 
                     Where-Object { $_.Name -like $filePattern }
            
            if ($files.Count -eq 0) {
                Write-Host "\n[!] No files found matching pattern: $filePattern" -ForegroundColor Yellow
                return $false
            }
            
            if ($ListMode) {
                if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
                    Write-Host "\nFound $($files.Count) file(s):\n" -ForegroundColor Green
                }
                else {
                    Write-Host "\nFound $($files.Count) file(s) that would be moved:\n" -ForegroundColor Green
                }
            }
            else {
                Write-Host "\nFound $($files.Count) file(s) to move\n" -ForegroundColor Green
            }
            
            $successCount = 0
            $failCount = 0
            
            foreach ($file in $files) {
                if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
                    # Just list the file
                    Write-Host "  $($file.Name)" -ForegroundColor White
                    Write-Host "    Location: $($file.ServerRelativeUrl)" -ForegroundColor Gray
                    
                    # Check for associated PDF
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    $fileFolder = $file.ServerRelativeUrl.Substring(0, $file.ServerRelativeUrl.LastIndexOf('/'))
                    $pdfPath = "$fileFolder/PDF/$baseName.pdf"
                    try {
                        $pdfFile = Get-PnPFile -Url $pdfPath -ErrorAction SilentlyContinue
                        if ($null -ne $pdfFile) {
                            Write-Host "    [✓] Has associated PDF: $pdfPath" -ForegroundColor Gray
                        }
                    }
                    catch { }
                    Write-Host ""
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
            Write-Host "\n[X] Error searching for files: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    else {
        # Single file
        try {
            $file = Get-PnPFile -Url $SourcePattern -AsListItem -ErrorAction Stop
            $fileObj = Get-PnPFile -Url $SourcePattern -ErrorAction Stop
            
            Write-Host ""
            
            if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
                # Just list the file
                Write-Host "  $($fileObj.Name)" -ForegroundColor White
                Write-Host "    Location: $($fileObj.ServerRelativeUrl)" -ForegroundColor Gray
                
                # Check for associated PDF
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileObj.Name)
                $fileFolder = $fileObj.ServerRelativeUrl.Substring(0, $fileObj.ServerRelativeUrl.LastIndexOf('/'))
                $pdfPath = "$fileFolder/PDF/$baseName.pdf"
                try {
                    $pdfFile = Get-PnPFile -Url $pdfPath -ErrorAction SilentlyContinue
                    if ($null -ne $pdfFile) {
                        Write-Host "    [✓] Has associated PDF: $pdfPath" -ForegroundColor Gray
                    }
                }
                catch { }
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
            Write-Host "\n[X] File not found or error accessing file: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
}

# Main script logic

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
            $siteUrl = "https://$tenantName.sharepoint.com/sites/$siteName"
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
        Write-Host "  - Full URL: 'https://contoso.sharepoint.com/sites/mysite'" -ForegroundColor Gray
        exit 1
    }
    
    $result = Connect-ToPnP -SiteUrl $siteUrl
    
    if (-not $result) {
        exit 1
    }
    
    Write-Host "`n[✓] Connection successful! You can now use PnP cmdlets." -ForegroundColor Green
    
    # If no other actions specified, exit successfully
    if (-not $Check -and -not $MoveFile -and -not $ListFiles) {
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
        Write-Host "\nUsage: .\Move-SPOFiles.ps1 -MoveFile '<source>' -Destination '<destination>'" -ForegroundColor Yellow
        exit 1
    }
    
    $result = Invoke-MoveFiles -SourcePattern $MoveFile -DestinationFolder $Destination
    
    if ($result) {
        exit 0
    }
    else {
        exit 1
    }
}
elseif ($ListFiles) {
    # Destination is optional for ListFiles
    if ([string]::IsNullOrWhiteSpace($Destination)) {
        $result = Invoke-MoveFiles -SourcePattern $ListFiles -ListMode
    }
    else {
        $result = Invoke-MoveFiles -SourcePattern $ListFiles -DestinationFolder $Destination -ListMode
    }
    
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
    Write-Host "  -MoveFile <src> -Destination <dest>   Move file(s) to destination" -ForegroundColor Gray
    Write-Host "  -ListFiles <pattern>      List files matching pattern" -ForegroundColor Gray
    Write-Host "  -ListFiles <src> -Destination <dest>  Preview files that would be moved (dry run)" -ForegroundColor Gray
    Write-Host "\nExamples:" -ForegroundColor Cyan
    Write-Host "  .\Move-SPOFiles.ps1 -Check" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso/mysite'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'https://contoso.sharepoint.com/sites/mysite'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso/mysite' -Check" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -MoveFile 'Shared Documents/file.docx' -Destination 'Shared Documents/Archive'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'contoso/mysite' -MoveFile 'Shared Documents/*Training*.docx' -Destination 'Shared Documents/Training'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -ListFiles 'Shared Documents/*Training*.docx'" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -ListFiles 'Shared Documents/*Training*.docx' -Destination 'Shared Documents/Training'" -ForegroundColor White
}

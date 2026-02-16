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
    Connects to a SharePoint site using the specified URL.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Check
    Checks if PnP.PowerShell is installed and if connected to a SharePoint site.

.EXAMPLE
    .\Move-SPOFiles.ps1 -Connect "https://contoso.sharepoint.com/sites/ProjectSite"
    Connects to the specified SharePoint site using interactive authentication.

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
    [string]$Connect
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

# Main script logic
if ($Check) {
    $result = Invoke-EnvironmentCheck
    
    if ($result) {
        exit 0
    }
    else {
        exit 1
    }
}
elseif ($Connect) {
    # Validate URL format
    if ($Connect -notmatch '^https://.*\.sharepoint\.com/') {
        Write-Host "[X] Invalid SharePoint URL format." -ForegroundColor Red
        Write-Host "Expected format: https://yourtenant.sharepoint.com/sites/yoursite" -ForegroundColor Yellow
        exit 1
    }
    
    $result = Connect-ToPnP -SiteUrl $Connect
    
    if ($result) {
        Write-Host "`n[✓] Connection successful! You can now use PnP cmdlets." -ForegroundColor Green
        exit 0
    }
    else {
        exit 1
    }
}
else {
    Write-Host "No action specified. Available options:" -ForegroundColor Yellow
    Write-Host "  -Check              Verify environment setup" -ForegroundColor Gray
    Write-Host "  -Connect <url>      Connect to a SharePoint site" -ForegroundColor Gray
    Write-Host "`nExamples:" -ForegroundColor Cyan
    Write-Host "  .\Move-SPOFiles.ps1 -Check" -ForegroundColor White
    Write-Host "  .\Move-SPOFiles.ps1 -Connect 'https://contoso.sharepoint.com/sites/ProjectSite'" -ForegroundColor White
}

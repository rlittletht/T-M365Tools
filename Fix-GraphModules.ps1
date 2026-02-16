# Fix-GraphModules.ps1
# This script fixes Microsoft.Graph module version conflicts

Write-Host "Fixing Microsoft.Graph module version conflicts..." -ForegroundColor Cyan
Write-Host ""

# Step 1: Remove loaded modules
Write-Host "Step 1: Unloading any loaded Microsoft.Graph modules..." -ForegroundColor Yellow
Get-Module Microsoft.Graph* | Remove-Module -Force
Write-Host "  Done." -ForegroundColor Green

# Step 2: Uninstall all versions
Write-Host ""
Write-Host "Step 2: Uninstalling all versions of Microsoft.Graph modules..." -ForegroundColor Yellow
try {
    Uninstall-Module Microsoft.Graph.Authentication -AllVersions -Force -ErrorAction SilentlyContinue
    Write-Host "  Removed Microsoft.Graph.Authentication" -ForegroundColor Gray
}
catch {
    Write-Host "  No Microsoft.Graph.Authentication to remove" -ForegroundColor Gray
}

try {
    Uninstall-Module Microsoft.Graph.Core -AllVersions -Force -ErrorAction SilentlyContinue
    Write-Host "  Removed Microsoft.Graph.Core" -ForegroundColor Gray
}
catch {
    Write-Host "  No Microsoft.Graph.Core to remove" -ForegroundColor Gray
}

# Step 3: Install latest version
Write-Host ""
Write-Host "Step 3: Installing latest Microsoft.Graph.Authentication module..." -ForegroundColor Yellow
try {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    Write-Host "  Installed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to install module: $_"
    exit 1
}

# Step 4: Verify installation
Write-Host ""
Write-Host "Step 4: Verifying installation..." -ForegroundColor Yellow
$authModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
$coreModule = Get-Module -ListAvailable -Name Microsoft.Graph.Core | Sort-Object Version -Descending | Select-Object -First 1

if ($authModule -and $coreModule) {
    Write-Host "  Microsoft.Graph.Authentication: $($authModule.Version)" -ForegroundColor Green
    Write-Host "  Microsoft.Graph.Core: $($coreModule.Version)" -ForegroundColor Green
    Write-Host ""
    Write-Host "Installation complete! You can now use Resolve-M365Link.ps1" -ForegroundColor Green
}
else {
    Write-Warning "Module verification failed. Please restart PowerShell and try again."
}

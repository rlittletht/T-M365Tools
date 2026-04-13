# Move-SPOFiles.ps1 - Development Prompts

This document chronicles the prompts and requests that shaped the development of the Move-SPOFiles.ps1 script.

---

## Initial State

The script existed with functionality to move files in SharePoint Online using PnP.PowerShell. It supported:
- Single file moves with `-MoveFile` and `-Destination` parameters
- Pattern-based moves (e.g., `*Training*.docx`) to move multiple files matching a pattern
- Automatic detection and moving of associated PDF files in PDF subdirectories
- Connection management with `-Connect` parameter (tenant name, tenant/site, or full URL)
- Environment checks with `-Check` parameter
- File listing with `-List` parameter
- Dry run mode with `-Test` parameter
- Verbose output with `-Verbose` parameter

The script's key feature was maintaining the association between files and their PDF versions by automatically detecting and moving PDFs stored in parallel PDF subdirectories.

---

## Enhancement Requests (Chronological)

### 1. Add Azure Blob Storage Checking
**Date:** April 12, 2026

**Prompt:** "Change this script to add functionality such that when checking for a supporting PDF file, do additional checks to see if there are mirrored Azure blobs in an azure storage account. Only do these azure checks if option CheckAzure is provided with a path to a json file describing the azure parameters."

**Requirements:**
- Add `-CheckAzure` parameter that accepts a path to a JSON configuration file
- JSON file format:
  ```json
  {
      "StorageAccount": "<azure_storage_account_name>",
      "StorageContainer": "<container_name>",
      "PathMappings": [
          ["SPO_Path_Substring", "Azure_Path_Replacement"]
      ]
  }
  ```
- PathMappings convert SharePoint URLs to Azure blob paths
- Example: SharePoint URL `https://washdist9.sharepoint.com/umps/Reference/Helmets%20and%20Flaps/PDF/Summary.pdf` with mapping `["https://washdist9.sharepoint.com/umps", "umps"]` becomes Azure blob path `umps/Reference/Helmets%20and%20Flaps/PDF/Summary.pdf`

**Implementation:**
- Added `-CheckAzure` parameter to accept JSON config file path
- Created `Load-AzureConfiguration` function:
  - Loads and parses JSON configuration file
  - Validates required properties (StorageAccount, StorageContainer, PathMappings)
  - Displays configuration summary upon successful load
  - Returns null on failure with appropriate error messages
- Created `Get-AzureBlobPath` function:
  - Converts SharePoint URLs to Azure blob paths using PathMappings
  - Iterates through path mappings to find matching patterns
  - Extracts relative path and combines with Azure path replacement
  - Returns null if no mapping matches
- Created `Test-AzureBlobExists` function:
  - Checks if Az.Storage module is available
  - Creates Azure Storage context (tries anonymous first, then authenticated)
  - Queries for blob existence using Get-AzStorageBlob
  - Returns true/false based on blob existence
  - Uses Write-Verbose for detailed logging
- Modified `Move-SPOFileWithPDF` function:
  - Added Azure checking logic after SharePoint PDF check
  - Constructs full SharePoint URL for PDF path
  - Calls Get-AzureBlobPath to translate to Azure blob path
  - Calls Test-AzureBlobExists to check Azure Storage
  - Displays results: "Found associated PDF in SharePoint" or "Found associated PDF in Azure"
  - Uses Write-Verbose for detailed Azure checking steps
- Added script-level variable `$script:AzureConfig` to store loaded configuration
- Added Azure configuration loading in main script logic:
  - Loads configuration at script start if `-CheckAzure` is provided
  - Displays configuration loading status
  - Sets `$script:AzureConfig` to null if loading fails (disables Azure checks)
- Updated help documentation:
  - Added `.PARAMETER CheckAzure` with detailed description
  - Added JSON structure example
  - Added path mapping example
  - Added examples showing CheckAzure usage with -List and -MoveFile
  - Updated .NOTES to include Az.Storage as optional module
  - Set Last Updated date to April 12, 2026

**Benefits:**
- Provides visibility into whether PDF files are backed up/mirrored in Azure Storage
- Non-intrusive: Azure checking is optional, script functions normally without it
- Flexible path mapping allows different SharePoint to Azure storage layouts
- Uses verbose logging so Azure checks can be monitored with -Verbose flag
- Graceful degradation: if Az.Storage module not installed, Azure checks are skipped

**Dependencies:**
- **Required:** PnP.PowerShell module (existing requirement)
- **Optional:** Az.Storage module (for Azure blob checking)
- **Optional:** Azure authentication (Az CLI or interactive Azure login)

**Example Usage:**
```powershell
# List files with Azure checking
.\Move-SPOFiles.ps1 -List "Shared Documents/*Training*.docx" -CheckAzure ".\azure-config.json"

# Move files with Azure checking and verbose output
.\Move-SPOFiles.ps1 -MoveFile "Shared Documents/file.docx" -Destination "Archive" -CheckAzure ".\azure-config.json" -Verbose

# Test mode with Azure checking
.\Move-SPOFiles.ps1 -MoveFile "Shared Documents/*.docx" -Destination "Archive" -Test -CheckAzure ".\azure-config.json"
```

**Bug Fix (Same Day):**
- **Issue:** Azure checking wasn't working with `-List` parameter without `-Destination`
- **Root Cause:** The listing code path bypassed `Move-SPOFileWithPDF` function where Azure checking was implemented
- **Solution:** Added Azure checking logic to both file listing code paths (pattern matching and single file listing)
- **Result:** Azure blob checking now works consistently across all script modes (List, Test, and Move)

**Bug Fix (Same Day):**
- **Issue:** Flashing/flickering output when running Azure checks
- **Root Cause:** Azure Storage cmdlets (Get-AzStorageBlob, New-AzStorageContext) display progress bars by default
- **Solution:** Set `$ProgressPreference = 'SilentlyContinue'` before Azure cmdlet calls, restore to 'Continue' after
- **Result:** Clean output without visual artifacts

**Performance Optimization (Same Day):**
- **Issue:** Creating new Azure Storage context for every blob check was slow and inefficient
- **Solution:** 
  - Added script-level `$script:AzureStorageContext` variable
  - Created `Initialize-AzureStorageContext` function to create context once at startup
  - Modified `Test-AzureBlobExists` to use cached context
- **Result:** Dramatically faster Azure checking, single authentication point

**JSON Configuration Example:**
```json
{
    "StorageAccount": "mystorageaccount",
    "StorageContainer": "sharepoint-backup",
    "PathMappings": [
        ["https://washdist9.sharepoint.com/umps", "umps"],
        ["https://washdist9.sharepoint.com/ops", "operations"]
    ]
}
```

---

### 2. Add Azure Blob Management on Move
**Date:** April 12, 2026

**Prompt:** "when performing a move for a matching PDF file, if azure is specified, copy the item from the old location to the new location in the container. if the optional 'DeleteOldAzureBlobs' is specified, delete the old azure blob."

**Requirements:**
- When moving PDF files, also copy corresponding Azure blobs to new locations
- Add `-DeleteOldAzureBlobs` switch parameter to optionally delete old blobs after copying
- Only applies during actual moves (not -List or -Test modes)

**Implementation:**
- Added `-DeleteOldAzureBlobs` switch parameter
- Added script-level variable `$script:DeleteOldAzureBlobs` to track deletion mode
- Created `Copy-AzureBlob` function:
  - Takes source and destination blob paths
  - Uses `Start-AzStorageBlobCopy` to initiate copy within same container
  - Waits for copy operation to complete
  - Returns success/failure status
  - Suppresses progress bars to avoid flashing output
- Created `Remove-AzureBlob` function:
  - Deletes a blob from Azure Storage
  - Returns success/failure status
  - Suppresses progress bars
- Modified `Move-SPOFileWithPDF` function:
  - Stores Azure blob path when checking for PDF existence
  - After successfully moving PDF in SharePoint, calculates destination blob path
  - Copies blob to new location if source blob exists
  - Optionally deletes old blob if `-DeleteOldAzureBlobs` is specified
  - Provides clear console output for each Azure operation
- Updated startup logic to display deletion mode warning
- Updated documentation with new parameter and examples

**Benefits:**
- Keeps Azure Storage synchronized with SharePoint when moving files
- Optional deletion prevents blob accumulation in old locations
- Non-intrusive: Azure operations only happen if CheckAzure is configured
- Clear feedback showing Azure copy/delete operations

**Example Usage:**
```powershell
# Move files and copy Azure blobs (keep old blobs)
.\Move-SPOFiles.ps1 -MoveFile "Reference/old/*.docx" -Destination "Reference/new" -CheckAzure ".\azure-config.json"

# Move files and copy Azure blobs, deleting old blobs
.\Move-SPOFiles.ps1 -MoveFile "Reference/old/*.docx" -Destination "Reference/new" -CheckAzure ".\azure-config.json" -DeleteOldAzureBlobs
```

**Safety Features:**
- Blob copy completes before deletion attempt
- Copy failure prevents deletion
- Detailed verbose logging for troubleshooting
- Warning message displayed when deletion mode is active

---

## Current Feature Set

The script now includes:

### Core Functionality
- ✅ Move single files or multiple files matching patterns
- ✅ Automatic detection and moving of associated PDF files
- ✅ Flexible connection options (tenant name, tenant/site, full URL)
- ✅ Environment validation and connection checking
- ✅ Dry run mode for previewing moves before execution

### PDF Management
- ✅ Checks for PDFs in parallel PDF subdirectories
- ✅ Automatically creates PDF destination folders if needed
- ✅ Moves PDFs alongside their source documents
- ✅ Checks both SharePoint and Azure Storage (when configured)

### Azure Storage Integration
- ✅ Optional Azure blob existence checking
- ✅ JSON-based configuration for storage account and path mappings
- ✅ Flexible path mapping for SharePoint-to-Azure translation
- ✅ Graceful degradation if Az.Storage module unavailable
- ✅ Detailed verbose logging for Azure operations
- ✅ Azure blob copying when moving PDF files
- ✅ Optional old blob deletion with `-DeleteOldAzureBlobs`
- ✅ Single Azure Storage context reused for all operations

### Usability
- ✅ List mode to preview files without moving
- ✅ Test mode to show what would be moved (dry run)
- ✅ Verbose mode showing all PnP.PowerShell commands
- ✅ Detailed error handling and user-friendly messages
- ✅ Progress tracking and summary statistics

---

## Technical Notes

### Path Management
The script handles two types of paths:
- **Server-relative paths**: Include site path (e.g., `/sites/mysite/Shared Documents/file.docx`)
  - Used by: Move-PnPFile, Get-PnPFile
- **Site-relative paths**: Exclude site path (e.g., `Shared Documents/file.docx`)
  - Used by: Get-PnPFolderItem, Resolve-PnPFolder

Functions `Get-ServerRelativePath` and `Get-SiteRelativePath` handle conversions.

### Azure Storage Context
The `Test-AzureBlobExists` function attempts two context creation methods:
1. **Anonymous context**: For public blob containers
2. **Authenticated context**: Using current Azure session (Az CLI or Connect-AzAccount)

### Error Handling
- JSON parsing errors are caught and reported
- Missing Az.Storage module is handled gracefully
- Azure authentication failures don't stop the script
- SharePoint connection issues block operations with clear error messages

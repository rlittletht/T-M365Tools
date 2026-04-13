<#
.SYNOPSIS
    Extracts external links from a WordPress Extended RSS XML export file.

.DESCRIPTION
    This script parses a WordPress Extended RSS (WXR) XML file and extracts all external links
    from items. It identifies links in:
    - Navigation menu items (_menu_item_url postmeta)
    - Other postmeta fields containing URLs
    - Content and excerpt fields with HTML links
    
    You can use Tools / Export on your WordPress site to generate the XML file.

.PARAMETER XMLFile
    Path to the WordPress Extended RSS XML file to parse.

.PARAMETER OutputFile
    Optional path to save the results as a CSV file.

.PARAMETER IncludeInternalLinks
    Switch to include internal links (same domain as base site). By default, only external links are returned.

.EXAMPLE
    .\Extract-WPExternalLinks.ps1 -XMLFile ".\wpdump.xml"
    
.EXAMPLE
    .\Extract-WPExternalLinks.ps1 -XMLFile ".\wpdump.xml" -OutputFile "external-links.csv"
    
.EXAMPLE
    .\Extract-WPExternalLinks.ps1 -XMLFile ".\wpdump.xml" -IncludeInternalLinks
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$XMLFile,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeInternalLinks
)

# Function to extract URLs from HTML content
function Get-URLsFromHTML {
    param(
        [string]$HTML
    )
    
    $urls = @()
    
    # Match href attributes in anchor tags
    $hrefPattern = 'href=["'']([^"'']+)["'']'
    $hrefMatches = [regex]::Matches($HTML, $hrefPattern)
    
    foreach ($match in $hrefMatches) {
        if ($match.Groups.Count -gt 1) {
            $urls += $match.Groups[1].Value
        }
    }
    
    return $urls
}

# Function to determine if a URL is external
function Test-IsExternalURL {
    param(
        [string]$URL,
        [string]$BaseDomain
    )
    
    if ([string]::IsNullOrWhiteSpace($URL)) {
        return $false
    }
    
    # Skip mailto:, tel:, and other non-http protocols
    if ($URL -notmatch '^https?://') {
        return $false
    }
    
    try {
        $uri = [System.Uri]$URL
        $baseUri = [System.Uri]$BaseDomain
        
        # Compare the host names
        return $uri.Host -ne $baseUri.Host
    }
    catch {
        # If we can't parse the URL, consider it invalid
        return $false
    }
}

Write-Host "Loading XML file: $XMLFile" -ForegroundColor Cyan

# Load the XML file
[xml]$xml = Get-Content -Path $XMLFile -Raw -Encoding UTF8

# Define the namespace manager for WordPress namespaces
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("content", "http://purl.org/rss/1.0/modules/content/")
$ns.AddNamespace("excerpt", "http://wordpress.org/export/1.2/excerpt/")
$ns.AddNamespace("wp", "http://wordpress.org/export/1.2/")
$ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")

# Get the base site URL
$baseSiteUrl = $xml.SelectSingleNode("//wp:base_blog_url", $ns).InnerText
Write-Host "Base site URL: $baseSiteUrl" -ForegroundColor Yellow

# Collection to store all found links
$allLinks = @()
$linkId = 1

# Iterate through all items
$items = $xml.SelectNodes("//item")
Write-Host "Processing $($items.Count) items..." -ForegroundColor Cyan

foreach ($item in $items) {
    $postId = $item.SelectSingleNode("wp:post_id", $ns).InnerText
    $postType = $item.SelectSingleNode("wp:post_type", $ns).InnerText
    $postTitle = $item.SelectSingleNode("title").InnerText
    
    # Check for nav_menu_item and extract _menu_item_url
    if ($postType -eq "nav_menu_item") {
        $menuItemUrl = $item.SelectSingleNode("wp:postmeta[wp:meta_key='_menu_item_url']/wp:meta_value", $ns)
        
        if ($menuItemUrl -and ![string]::IsNullOrWhiteSpace($menuItemUrl.InnerText)) {
            $url = $menuItemUrl.InnerText
            
            $isExternal = Test-IsExternalURL -URL $url -BaseDomain $baseSiteUrl
            
            if ($isExternal -or $IncludeInternalLinks) {
                $allLinks += [PSCustomObject]@{
                    ID = $linkId++
                    PostID = $postId
                    PostType = $postType
                    PostTitle = $postTitle
                    LinkType = "nav_menu_item_url"
                    "Link URL" = $url
                    IsExternal = $isExternal
                }
            }
        }
    }
    
    # Check all postmeta for URLs (looking for common URL fields)
    $urlMetaKeys = @('_menu_item_url', '_gallery_link_url', 'url', 'link_url', 'external_link')
    
    foreach ($metaKey in $urlMetaKeys) {
        $metaValues = $item.SelectNodes("wp:postmeta[wp:meta_key='$metaKey']/wp:meta_value", $ns)
        
        foreach ($metaValue in $metaValues) {
            if (![string]::IsNullOrWhiteSpace($metaValue.InnerText)) {
                $url = $metaValue.InnerText
                
                # Check if it's a URL
                if ($url -match '^https?://') {
                    $isExternal = Test-IsExternalURL -URL $url -BaseDomain $baseSiteUrl
                    
                    if ($isExternal -or $IncludeInternalLinks) {
                        # Check if we already have this exact URL for this post
                        $exists = $allLinks | Where-Object { $_.PostID -eq $postId -and $_.URL -eq $url }
                        
                        if (-not $exists) {
                            $allLinks += [PSCustomObject]@{
                                ID = $linkId++
                                PostID = $postId
                                PostType = $postType
                                PostTitle = $postTitle
                                LinkType = "postmeta:$metaKey"
                                "Link URL" = $url
                                IsExternal = $isExternal
                            }
                        }
                    }
                }
            }
        }
    }
    
    # Extract URLs from content:encoded
    $contentEncoded = $item.SelectSingleNode("content:encoded", $ns)
    if ($contentEncoded -and ![string]::IsNullOrWhiteSpace($contentEncoded.InnerText)) {
        $urls = Get-URLsFromHTML -HTML $contentEncoded.InnerText
        
        foreach ($url in $urls) {
            if ($url -match '^https?://') {
                $isExternal = Test-IsExternalURL -URL $url -BaseDomain $baseSiteUrl
                
                if ($isExternal -or $IncludeInternalLinks) {
                    # Check if we already have this exact URL for this post
                    $exists = $allLinks | Where-Object { $_.PostID -eq $postId -and $_.URL -eq $url }
                    
                    if (-not $exists) {
                        $allLinks += [PSCustomObject]@{
                            ID = $linkId++
                            PostID = $postId
                            PostType = $postType
                            PostTitle = $postTitle
                            LinkType = "content:encoded"
                            "Link URL" = $url
                            IsExternal = $isExternal
                        }
                    }
                }
            }
        }
    }
    
    # Extract URLs from excerpt:encoded
    $excerptEncoded = $item.SelectSingleNode("excerpt:encoded", $ns)
    if ($excerptEncoded -and ![string]::IsNullOrWhiteSpace($excerptEncoded.InnerText)) {
        $urls = Get-URLsFromHTML -HTML $excerptEncoded.InnerText
        
        foreach ($url in $urls) {
            if ($url -match '^https?://') {
                $isExternal = Test-IsExternalURL -URL $url -BaseDomain $baseSiteUrl
                
                if ($isExternal -or $IncludeInternalLinks) {
                    # Check if we already have this exact URL for this post
                    $exists = $allLinks | Where-Object { $_.PostID -eq $postId -and $_.URL -eq $url }
                    
                    if (-not $exists) {
                        $allLinks += [PSCustomObject]@{
                            ID = $linkId++
                            PostID = $postId
                            PostType = $postType
                            PostTitle = $postTitle
                            LinkType = "excerpt:encoded"
                            "Link URL" = $url
                            IsExternal = $isExternal
                        }
                    }
                }
            }
        }
    }
}

# Display results
Write-Host "`nFound $($allLinks.Count) links" -ForegroundColor Green

if ($allLinks.Count -gt 0) {
    # Group by external/internal
    $externalCount = ($allLinks | Where-Object { $_.IsExternal -eq $true }).Count
    $internalCount = ($allLinks | Where-Object { $_.IsExternal -eq $false }).Count
    
    Write-Host "  External: $externalCount" -ForegroundColor Yellow
    Write-Host "  Internal: $internalCount" -ForegroundColor Yellow
    
    # Display sample of links
    Write-Host "`nSample of found links:" -ForegroundColor Cyan
    $allLinks | Select-Object -First 10 | Format-Table -AutoSize
    
    # Export to CSV if requested
    if ($OutputFile) {
        $allLinks | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        Write-Host "`nResults exported to: $OutputFile" -ForegroundColor Green
    }
    
    # Return the collection
    return $allLinks
}
else {
    Write-Host "No links found in the XML file." -ForegroundColor Yellow
}

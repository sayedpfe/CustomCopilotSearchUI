# Deploy-SearchPage.ps1
# Creates the Bosch AI Search full-page experience on a Communication Site.
# Uses the single "Bosch AI Search" full-page web part (BoschSearchApp) which
# renders the complete landing page, search, AI answers, results, news, and chat.
#
# The web part manifest has "supportsFullBleed": true, so it can be placed
# in a OneColumnFullWidth section. If that fails (e.g., solution not yet
# redeployed), it falls back to a regular OneColumn section which still
# looks great on a Communication Site (no left nav).
#
# Prerequisites: PnP.PowerShell module, solution deployed to App Catalog, API permissions approved.
#
# Usage:
#   .\Deploy-SearchPage.ps1 -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/BoschAISearch"
#   .\Deploy-SearchPage.ps1 -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/BoschAISearch" -SetAsHomePage
#   .\Deploy-SearchPage.ps1 -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/BoschAISearch" -ClientId "your-app-client-id"

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$false)]
    [string]$ClientId,

    [Parameter(Mandatory=$false)]
    [string]$PageName = "BoschAISearch",

    [Parameter(Mandatory=$false)]
    [switch]$SetAsHomePage
)

# Connect to SharePoint
Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
if ($ClientId) {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

# ============================================================
# Bosch AI Search Full-Page Web Part ID (from manifest)
# ============================================================
$wpBoschSearchApp = "d1e2f3a4-b5c6-7890-abcd-ef1234567890"

$wpProperties = @{
    "groundingMode" = "work"
    "maxRetrievalResults" = 10
    "showCopilotLink" = $true
    "newsSourceSiteUrl" = ""
    "promotedResultsListName" = "SearchPromotedResults"
    "announcementsListName" = "SearchAnnouncements"
    "analyticsListName" = "SearchAnalyticsEvents"
}

# ============================================================
# 1. Create or recreate the page
# ============================================================
Write-Host "`nCreating page '$PageName'..." -ForegroundColor Yellow

$existingPage = Get-PnPPage -Identity $PageName -ErrorAction SilentlyContinue
if ($null -ne $existingPage) {
    Write-Host "  Page '$PageName' already exists. Removing to recreate..." -ForegroundColor Gray
    Remove-PnPPage -Identity $PageName -Force
}

# Use Home layout for a clean full-page look on Communication Sites
$page = Add-PnPPage -Name $PageName -LayoutType Home -Title "Bosch AI Search"
Write-Host "  Page created (Home layout)." -ForegroundColor Green

# ============================================================
# 2. Add section and web part
# ============================================================
Write-Host "`nAdding Bosch AI Search web part..." -ForegroundColor Yellow

# Try full-width section first (requires supportsFullBleed in manifest + redeployed .sppkg)
$fullWidthSuccess = $false
try {
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumnFullWidth -Order 1
    Add-PnPPageWebPart -Page $PageName -Component $wpBoschSearchApp -Section 1 -Column 1 -WebPartProperties $wpProperties
    $fullWidthSuccess = $true
    Write-Host "  Added in full-width section." -ForegroundColor Green
} catch {
    Write-Host "  Full-width section not supported for this web part. Falling back to single column..." -ForegroundColor Yellow
    Write-Host "  (Rebuild and redeploy the .sppkg to enable full-width support)" -ForegroundColor Gray

    # Remove the failed page and recreate
    Remove-PnPPage -Identity $PageName -Force
    $page = Add-PnPPage -Name $PageName -LayoutType Home -Title "Bosch AI Search"

    # Use regular OneColumn - still looks great on Communication Sites (no left nav)
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 1
    Add-PnPPageWebPart -Page $PageName -Component $wpBoschSearchApp -Section 1 -Column 1 -WebPartProperties $wpProperties
    Write-Host "  Added in single-column section." -ForegroundColor Green
}

# ============================================================
# 3. Publish the page
# ============================================================
Write-Host "`nPublishing page..." -ForegroundColor Yellow
Set-PnPPage -Identity $PageName -Publish
Write-Host "  Page published." -ForegroundColor Green

# ============================================================
# 4. Optionally set as home page
# ============================================================
if ($SetAsHomePage) {
    Write-Host "`nSetting page as site home page..." -ForegroundColor Yellow
    Set-PnPHomePage -RootFolderRelativeUrl "SitePages/$PageName.aspx"
    Write-Host "  Page set as home page." -ForegroundColor Green
}

# ============================================================
# Done
# ============================================================
$pageUrl = "$SiteUrl/SitePages/$PageName.aspx"
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Bosch AI Search page deployed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Page URL: $pageUrl"
Write-Host ""
if ($fullWidthSuccess) {
    Write-Host "Layout: Full-width section (edge-to-edge)" -ForegroundColor Green
} else {
    Write-Host "Layout: Single-column section (rebuild + redeploy .sppkg for full-width)" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "Web part: Bosch AI Search (full-page app)"
Write-Host "  - Landing page with BOSCH logo, Work/Web toggle, search box, news carousel"
Write-Host "  - Search results with AI answers (Copilot or Graph Search fallback)"
Write-Host "  - Promoted results from SearchPromotedResults list"
Write-Host "  - Announcements from SearchAnnouncements list"
Write-Host "  - Analytics tracking to SearchAnalyticsEvents list"
Write-Host "  - Chat panel for Copilot-licensed users"
Write-Host ""
if ($SetAsHomePage) {
    Write-Host "Home page: Yes (this page is the site landing page)"
} else {
    Write-Host "Tip: Re-run with -SetAsHomePage to make this the site landing page."
}
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Open $pageUrl and verify the full-page experience loads"
Write-Host "  2. Test search queries to verify AI answers and search results"
Write-Host "  3. Check promoted results by searching 'hr' or 'sales'"
Write-Host "  4. Verify news carousel shows SharePoint news pages"

Disconnect-PnPOnline

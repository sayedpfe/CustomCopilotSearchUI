# Deploy-Lists.ps1
# Creates the 3 SharePoint lists required by the Bosch Search Extensions web parts.
# Prerequisites: PnP.PowerShell module (Install-Module PnP.PowerShell)
#
# Usage:
#   .\Deploy-Lists.ps1 -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/BoschAISearch" -ClientId "your-app-client-id"
#
# The ClientId must be an Entra ID App Registration with:
#   - Redirect URI: http://localhost (Public client/native)
#   - Delegated permissions: Microsoft Graph > Sites.Manage.All, Sites.Read.All
#                            SharePoint > AllSites.Manage
#   - Admin consent granted
#   - Allow public client flows: Yes

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$ClientId
)

# Connect to SharePoint
Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ============================================================
# 1. SearchAnnouncements List
# ============================================================
Write-Host "`nCreating SearchAnnouncements list..." -ForegroundColor Yellow

$listExists = Get-PnPList -Identity "SearchAnnouncements" -ErrorAction SilentlyContinue
if ($null -eq $listExists) {
    New-PnPList -Title "SearchAnnouncements" -Template GenericList -Url "Lists/SearchAnnouncements"

    # Add columns
    Add-PnPField -List "SearchAnnouncements" -DisplayName "Message" -InternalName "Message" -Type Note -AddToDefaultView
    Add-PnPField -List "SearchAnnouncements" -DisplayName "Severity" -InternalName "Severity" -Type Choice -Choices "Info","Warning","Error","Success" -AddToDefaultView
    Add-PnPField -List "SearchAnnouncements" -DisplayName "StartDate" -InternalName "StartDate" -Type DateTime -AddToDefaultView
    Add-PnPField -List "SearchAnnouncements" -DisplayName "EndDate" -InternalName "EndDate" -Type DateTime -AddToDefaultView
    Add-PnPField -List "SearchAnnouncements" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
    Add-PnPField -List "SearchAnnouncements" -DisplayName "TargetAudience" -InternalName "TargetAudience" -Type Choice -Choices "All","HR","Sales","IT","Engineering" -AddToDefaultView
    Add-PnPField -List "SearchAnnouncements" -DisplayName "SortOrder" -InternalName "SortOrder" -Type Number -AddToDefaultView

    # Add sample data
    Add-PnPListItem -List "SearchAnnouncements" -Values @{
        "Title" = "System Maintenance Scheduled";
        "Message" = "Search indexing will be temporarily unavailable on March 15, 2026 from 2:00-4:00 AM CET.";
        "Severity" = "Warning";
        "StartDate" = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ");
        "EndDate" = (Get-Date).AddDays(30).ToString("yyyy-MM-ddTHH:mm:ssZ");
        "IsActive" = $true;
        "TargetAudience" = "All";
        "SortOrder" = 1
    }

    Add-PnPListItem -List "SearchAnnouncements" -Values @{
        "Title" = "New HR Portal Available";
        "Message" = "The new HR self-service portal is now live. Access it from the Quick Links section.";
        "Severity" = "Info";
        "StartDate" = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ");
        "EndDate" = (Get-Date).AddDays(60).ToString("yyyy-MM-ddTHH:mm:ssZ");
        "IsActive" = $true;
        "TargetAudience" = "All";
        "SortOrder" = 2
    }

    Write-Host "  SearchAnnouncements list created with sample data." -ForegroundColor Green
} else {
    Write-Host "  SearchAnnouncements list already exists. Skipping." -ForegroundColor Gray
}

# ============================================================
# 2. SearchPromotedResults List
# ============================================================
Write-Host "`nCreating SearchPromotedResults list..." -ForegroundColor Yellow

$listExists = Get-PnPList -Identity "SearchPromotedResults" -ErrorAction SilentlyContinue
if ($null -eq $listExists) {
    New-PnPList -Title "SearchPromotedResults" -Template GenericList -Url "Lists/SearchPromotedResults"

    # Add columns
    Add-PnPField -List "SearchPromotedResults" -DisplayName "Description" -InternalName "Description0" -Type Note -AddToDefaultView
    Add-PnPField -List "SearchPromotedResults" -DisplayName "Url" -InternalName "PromotedUrl" -Type URL -AddToDefaultView
    Add-PnPField -List "SearchPromotedResults" -DisplayName "Keywords" -InternalName "Keywords" -Type Note -AddToDefaultView
    Add-PnPField -List "SearchPromotedResults" -DisplayName "IconUrl" -InternalName "IconUrl" -Type URL
    Add-PnPField -List "SearchPromotedResults" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
    Add-PnPField -List "SearchPromotedResults" -DisplayName "StartDate" -InternalName "StartDate" -Type DateTime
    Add-PnPField -List "SearchPromotedResults" -DisplayName "EndDate" -InternalName "EndDate" -Type DateTime
    Add-PnPField -List "SearchPromotedResults" -DisplayName "SortOrder" -InternalName "SortOrder" -Type Number -AddToDefaultView

    # Add sample data
    Add-PnPListItem -List "SearchPromotedResults" -Values @{
        "Title" = "Bosch HR Self-Service Portal";
        "Description0" = "Access your payslips, request time off, and manage your benefits.";
        "PromotedUrl" = "https://hr.bosch.com, Bosch HR Portal";
        "Keywords" = "hr, human resources, payslip, time off, benefits, vacation";
        "IsActive" = $true;
        "SortOrder" = 1
    }

    Add-PnPListItem -List "SearchPromotedResults" -Values @{
        "Title" = "IT Service Desk";
        "Description0" = "Submit IT tickets, request new equipment, or get help with software.";
        "PromotedUrl" = "https://servicedesk.bosch.com, IT Service Desk";
        "Keywords" = "it help, service desk, ticket, computer, software, password, support";
        "IsActive" = $true;
        "SortOrder" = 2
    }

    Add-PnPListItem -List "SearchPromotedResults" -Values @{
        "Title" = "Q4 2025 Sales Dashboard";
        "Description0" = "View the latest sales performance metrics and regional breakdowns.";
        "PromotedUrl" = "https://powerbi.bosch.com/sales-dashboard, Sales Dashboard";
        "Keywords" = "sales, revenue, dashboard, q4, quarterly, performance";
        "IsActive" = $true;
        "SortOrder" = 3
    }

    Write-Host "  SearchPromotedResults list created with sample data." -ForegroundColor Green
} else {
    Write-Host "  SearchPromotedResults list already exists. Skipping." -ForegroundColor Gray
}

# ============================================================
# 3. SearchAnalyticsEvents List
# ============================================================
Write-Host "`nCreating SearchAnalyticsEvents list..." -ForegroundColor Yellow

$listExists = Get-PnPList -Identity "SearchAnalyticsEvents" -ErrorAction SilentlyContinue
if ($null -eq $listExists) {
    New-PnPList -Title "SearchAnalyticsEvents" -Template GenericList -Url "Lists/SearchAnalyticsEvents"

    # Add columns
    Add-PnPField -List "SearchAnalyticsEvents" -DisplayName "EventType" -InternalName "EventType" -Type Choice -Choices "Query","Click","ZeroResult" -AddToDefaultView
    Add-PnPField -List "SearchAnalyticsEvents" -DisplayName "ResultCount" -InternalName "ResultCount" -Type Number -AddToDefaultView
    Add-PnPField -List "SearchAnalyticsEvents" -DisplayName "ClickedUrl" -InternalName "ClickedUrl" -Type URL
    Add-PnPField -List "SearchAnalyticsEvents" -DisplayName "ClickPosition" -InternalName "ClickPosition" -Type Number
    Add-PnPField -List "SearchAnalyticsEvents" -DisplayName "Timestamp" -InternalName "Timestamp" -Type DateTime -AddToDefaultView
    Add-PnPField -List "SearchAnalyticsEvents" -DisplayName "SessionId" -InternalName "SessionId" -Type Text
    Add-PnPField -List "SearchAnalyticsEvents" -DisplayName "Vertical" -InternalName "Vertical" -Type Text

    # Create indexed columns for performance
    Set-PnPField -List "SearchAnalyticsEvents" -Identity "Timestamp" -Values @{Indexed=$true}
    Set-PnPField -List "SearchAnalyticsEvents" -Identity "EventType" -Values @{Indexed=$true}

    Write-Host "  SearchAnalyticsEvents list created with indexed columns." -ForegroundColor Green
} else {
    Write-Host "  SearchAnalyticsEvents list already exists. Skipping." -ForegroundColor Gray
}

# ============================================================
# Done
# ============================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "All lists provisioned successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Lists created:"
Write-Host "  1. SearchAnnouncements     - Banner messages for search page"
Write-Host "  2. SearchPromotedResults   - Keyword-triggered promoted results"
Write-Host "  3. SearchAnalyticsEvents   - Search query and click tracking"
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Deploy the .sppkg to the App Catalog"
Write-Host "  2. Approve Graph API permissions in SharePoint Admin"
Write-Host "  3. Add the web parts to your search page"

Disconnect-PnPOnline

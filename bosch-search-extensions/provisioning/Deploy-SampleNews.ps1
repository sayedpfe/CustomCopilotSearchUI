# Deploy-SampleNews.ps1
# Creates sample SharePoint news pages on the target site for the Bosch AI Search news carousel.
# News pages are site pages promoted as news (PromotedState = 2).
#
# Prerequisites: PnP.PowerShell module
#
# Usage:
#   .\Deploy-SampleNews.ps1 -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/BoschAISearch" -ClientId "14603fc7-543c-4dad-b17a-81de8aac9600"

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
# Sample news articles
# ============================================================
$newsArticles = @(
    @{
        PageName    = "Bosch-AI-Workplace-Launch"
        Title       = "Bosch Launches AI-Powered Workplace Search"
        Description = "Our new enterprise search platform brings AI answers, Copilot integration, and a unified search experience to every employee."
        ImageUrl    = "https://cdn.pixabay.com/photo/2024/02/01/06/55/ai-generated-8545283_1280.jpg"
        Content     = @"
<div>
<h2>Transforming How We Find Information</h2>
<p>We are excited to announce the launch of <strong>Bosch AI Search</strong> — a next-generation enterprise search experience powered by Microsoft 365 Copilot and Microsoft Graph.</p>
<p>The new platform provides:</p>
<ul>
<li><strong>AI-powered answers</strong> — Get instant, synthesized answers to your questions from enterprise content</li>
<li><strong>Copilot Chat integration</strong> — Have multi-turn conversations grounded in your company data</li>
<li><strong>Unified search</strong> — Search across SharePoint, OneDrive, Teams, emails, and external connectors from one place</li>
<li><strong>People search</strong> — Find experts with org chart, recent documents, and expertise tags</li>
</ul>
<p>Visit the <a href="/">Bosch AI Search</a> home page to try it now.</p>
<h3>What's Next</h3>
<p>We are rolling out the platform to all business units over the coming weeks. Training sessions are available — check the Learning Portal for schedules.</p>
</div>
"@
        Topic       = "Technology"
    },
    @{
        PageName    = "Q1-2026-Sales-Results"
        Title       = "Q1 2026 Sales Results: Strong Growth Across All Divisions"
        Description = "Bosch reports impressive growth in Q1 2026, with Mobility Solutions and Industrial Technology leading the way."
        ImageUrl    = "https://cdn.pixabay.com/photo/2020/07/08/04/12/work-5382501_1280.jpg"
        Content     = @"
<div>
<h2>Record-Breaking Quarter</h2>
<p>Bosch has reported strong financial results for Q1 2026, with total revenue up 8.3% year-over-year. Key highlights include:</p>
<ul>
<li><strong>Mobility Solutions</strong> — Revenue increased 12% driven by EV component demand</li>
<li><strong>Industrial Technology</strong> — Growth of 9.5% supported by automation and IoT solutions</li>
<li><strong>Consumer Goods</strong> — Steady 5% increase with strong performance in smart home products</li>
<li><strong>Energy and Building Technology</strong> — 7% growth from sustainable energy solutions</li>
</ul>
<h3>Regional Performance</h3>
<p>Asia-Pacific led regional growth at 14%, followed by Europe at 7% and North America at 6%. The strategic investments in electrification and AI-driven manufacturing continue to pay dividends.</p>
<p>For detailed financial reports, visit the Finance Portal or contact the Investor Relations team.</p>
</div>
"@
        Topic       = "Finance"
    },
    @{
        PageName    = "Cybersecurity-Awareness-2026"
        Title       = "Cybersecurity Awareness: New Policies and Training Resources"
        Description = "Updated security policies for 2026 and mandatory training requirements for all employees."
        ImageUrl    = "https://cdn.pixabay.com/photo/2018/05/14/16/54/cyber-security-3400657_1280.jpg"
        Content     = @"
<div>
<h2>Staying Secure in 2026</h2>
<p>As cyber threats continue to evolve, Bosch is strengthening its security posture with updated policies and new training programs.</p>
<h3>Key Policy Updates</h3>
<ul>
<li><strong>Multi-Factor Authentication (MFA)</strong> — Now required for all internal applications, not just external access</li>
<li><strong>Zero Trust Network</strong> — All access requests are verified regardless of location</li>
<li><strong>Data Classification</strong> — New 4-tier classification system (Public, Internal, Confidential, Restricted)</li>
<li><strong>AI Usage Policy</strong> — Guidelines for using generative AI tools with company data</li>
</ul>
<h3>Mandatory Training</h3>
<p>All employees must complete the updated Cybersecurity Fundamentals course by March 31, 2026. Access it through the Learning Management System.</p>
<p>Report suspicious activity to <strong>security@bosch.com</strong> or use the Security Incident button in ServiceNow.</p>
</div>
"@
        Topic       = "Security"
    },
    @{
        PageName    = "HR-Benefits-Open-Enrollment"
        Title       = "Open Enrollment 2026: New Benefits and Wellness Programs"
        Description = "Explore new health plans, expanded wellness benefits, and flexible work arrangements available this year."
        ImageUrl    = "https://cdn.pixabay.com/photo/2017/10/06/14/22/handshake-2824357_1280.jpg"
        Content     = @"
<div>
<h2>Your Benefits, Your Choice</h2>
<p>Open enrollment for 2026 benefits is now underway. This year, we are introducing several enhancements to support your well-being and work-life balance.</p>
<h3>New Offerings</h3>
<ul>
<li><strong>Enhanced Mental Health Support</strong> — Expanded counseling sessions from 6 to 12 per year, plus new digital therapy options</li>
<li><strong>Wellness Stipend</strong> — $500 annual stipend for fitness, nutrition, or mindfulness programs</li>
<li><strong>Family Care</strong> — Backup childcare and eldercare services now included</li>
<li><strong>Learning Budget</strong> — $2,000 annual professional development allowance</li>
</ul>
<h3>Flexible Work</h3>
<p>The hybrid work policy continues with a minimum of 2 days in office per week. New: Employees can now work from any Bosch office globally for up to 4 weeks per year.</p>
<p>Enroll by April 15, 2026 through the HR Portal. Questions? Contact your HR Business Partner.</p>
</div>
"@
        Topic       = "Human Resources"
    },
    @{
        PageName    = "Innovation-Day-2026"
        Title       = "Bosch Innovation Day 2026: Call for Projects"
        Description = "Submit your innovative ideas for the annual Innovation Day showcase. Prizes, mentorship, and funding for winning teams."
        ImageUrl    = "https://cdn.pixabay.com/photo/2018/09/04/10/27/never-stop-learning-3653430_1280.jpg"
        Content     = @"
<div>
<h2>Innovate. Create. Transform.</h2>
<p>The 2026 Bosch Innovation Day is coming in June, and we want to see your ideas! This annual event showcases breakthrough innovations from teams across all divisions.</p>
<h3>This Year's Themes</h3>
<ul>
<li><strong>AI for Everyone</strong> — Solutions that democratize AI across the organization</li>
<li><strong>Sustainable Operations</strong> — Ideas that reduce environmental impact</li>
<li><strong>Connected Experiences</strong> — IoT and digital twin innovations</li>
<li><strong>Future of Work</strong> — Tools and processes that transform how we collaborate</li>
</ul>
<h3>Prizes</h3>
<p><strong>Grand Prize:</strong> $50,000 project funding + executive mentorship program</p>
<p><strong>Runner-up:</strong> $25,000 project funding</p>
<p><strong>People's Choice:</strong> $10,000 team experience</p>
<h3>How to Submit</h3>
<p>Submit your project proposal through the Innovation Portal by May 1, 2026. Teams of 2-5 people from any division can participate.</p>
</div>
"@
        Topic       = "Innovation"
    },
    @{
        PageName    = "Sustainability-Report-2025"
        Title       = "2025 Sustainability Report: Carbon Neutral by 2030 Progress"
        Description = "Bosch has reduced Scope 1 and 2 emissions by 42% since 2018. See the full progress report and 2026 targets."
        ImageUrl    = "https://cdn.pixabay.com/photo/2022/01/07/11/45/solar-power-6921283_1280.jpg"
        Content     = @"
<div>
<h2>Our Path to Carbon Neutrality</h2>
<p>Bosch remains committed to achieving carbon neutrality across all operations by 2030. The 2025 Sustainability Report shows significant progress.</p>
<h3>Key Achievements in 2025</h3>
<ul>
<li><strong>42% reduction</strong> in Scope 1 and 2 emissions since 2018 baseline</li>
<li><strong>67% renewable energy</strong> usage across all global facilities</li>
<li><strong>35% recycled materials</strong> in manufacturing processes</li>
<li><strong>Zero waste to landfill</strong> achieved at 78% of production sites</li>
</ul>
<h3>2026 Targets</h3>
<p>We aim to reach 80% renewable energy, expand electric vehicle fleet to 50% of company vehicles, and launch the Supplier Sustainability Scorecard program.</p>
<p>Download the full report from the Sustainability Portal or contact the ESG team for details.</p>
</div>
"@
        Topic       = "Sustainability"
    }
)

# ============================================================
# Create news pages
# ============================================================
$created = 0
$skipped = 0

foreach ($article in $newsArticles) {
    Write-Host "`nCreating news page: $($article.Title)..." -ForegroundColor Yellow

    $existingPage = Get-PnPPage -Identity $article.PageName -ErrorAction SilentlyContinue
    if ($null -ne $existingPage) {
        Write-Host "  Page '$($article.PageName)' already exists, skipping." -ForegroundColor Gray
        $skipped++
        continue
    }

    # Create the page
    $page = Add-PnPPage -Name $article.PageName -LayoutType Article -Title $article.Title

    # Add a section so text can be inserted
    Add-PnPPageSection -Page $article.PageName -SectionTemplate OneColumn -Order 1

    # Add text content
    Add-PnPPageTextPart -Page $article.PageName -Text $article.Content -Section 1 -Column 1

    # Set page description via list item (Set-PnPPage doesn't support -Description)
    $pageItem = Get-PnPListItem -List "Site Pages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$($article.PageName).aspx</Value></Eq></Where></Query></View>"
    if ($null -ne $pageItem) {
        Set-PnPListItem -List "Site Pages" -Identity $pageItem.Id -Values @{ "Description" = $article.Description } | Out-Null
    }

    # Promote as news and publish
    Set-PnPPage -Identity $article.PageName -PromoteAs NewsArticle
    Set-PnPPage -Identity $article.PageName -Publish

    Write-Host "  Created and promoted as news." -ForegroundColor Green
    $created++
}

# ============================================================
# Done
# ============================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Sample news deployment complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Created: $created news pages"
Write-Host "Skipped: $skipped (already existed)"
Write-Host ""
Write-Host "News pages:"
foreach ($article in $newsArticles) {
    Write-Host "  - $($article.Title)"
    Write-Host "    $SiteUrl/SitePages/$($article.PageName).aspx"
}
Write-Host ""
Write-Host "These pages will appear in the Bosch AI Search news carousel."
Write-Host "You can edit them from Site Pages or create additional news pages manually."

Disconnect-PnPOnline

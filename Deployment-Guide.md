# Bosch Enterprise Search - Deployment Guide

This guide covers deploying the **bosch-search-extensions** SPFx solution to your SharePoint Online tenant. The solution includes the **Bosch AI Search** full-page experience (recommended) plus 7 standalone web parts.

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| **Node.js** | v16.x or v18.x (LTS) — **not v20+** (SPFx 1.18 incompatible) |
| **SPFx CLI tools** | `npm install -g gulp-cli yo @microsoft/generator-sharepoint` |
| **PnP PowerShell** | `Install-Module PnP.PowerShell` (for list provisioning) |
| **SharePoint Admin** | Access to the tenant App Catalog and API access page |
| **Target site** | `https://m365cpi90282478.sharepoint.com/sites/Bosch` |

---

## Step 1: Create a Communication Site

The **Bosch AI Search** full-page web part requires a **Communication Site** (not a Team Site) for full-width layout without the left navigation panel.

**Option A — Create a new Communication Site:**
1. Go to SharePoint Admin Center > **Sites** > **Active sites** > **Create**
2. Choose **Communication site**
3. Name: `Bosch AI Search` (or your preferred name)
4. URL: e.g., `https://m365cpi90282478.sharepoint.com/sites/BoschSearch`

**Option B — Use an existing Communication Site:**
- If you already have a Communication Site at `/sites/Bosch`, you can use it directly
- If it's currently a Team Site, you should create a new Communication Site instead (Team Sites cannot be converted)

> **Why Communication Site?** Team Sites have a permanent left navigation sidebar and don't support full-width page layouts. The Bosch AI Search landing page (large logo, centered search box, news carousel) needs a clean full-width canvas.

---

## Step 2: Install Dependencies

```bash
cd SearchCustomUI/bosch-search-extensions
npm install
```

---

## Step 3: Build and Validate

```bash
# Development build (catches TypeScript errors)
gulp build

# Production build
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` package will be generated at:
```
sharepoint/solution/bosch-search-extensions.sppkg
```

---

## Step 4: Provision SharePoint Lists

Three SharePoint lists must exist on the target site before the web parts that depend on them will work.

```powershell
cd provisioning
.\Deploy-Lists.ps1 -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/Bosch"
```

This creates:

| List | Purpose | Used By |
|------|---------|---------|
| **SearchAnnouncements** | Banner messages (title, message, severity, date range, audience) | Bosch AI Search app, Announcements Banner |
| **SearchPromotedResults** | Keyword-triggered promoted links (title, URL, keywords, date range) | Bosch AI Search app, Sponsored Results |
| **SearchAnalyticsEvents** | Query/click tracking (event type, result count, timestamp) | Search Analytics |

Sample data is included automatically. You can modify or remove it from the lists after deployment.

---

## Step 5: Deploy to App Catalog

1. Go to your **SharePoint Admin Center** > **More features** > **Apps** > **App Catalog**
2. If no App Catalog exists, create one first
3. Upload `sharepoint/solution/bosch-search-extensions.sppkg`
4. In the deployment dialog:
   - Check **"Make this solution available to all sites in the organization"** (recommended for tenant-wide deployment)
   - Click **Deploy**

> **Note:** The solution deploys **8 web parts** in a single package: the Bosch AI Search full-page app + 7 standalone web parts.

---

## Step 6: Approve Graph API Permissions

This is **critical**. The solution requests Microsoft Graph permissions that must be approved by a SharePoint or Global Admin.

1. Go to **SharePoint Admin Center** > **Advanced** > **API access**
2. You will see pending permission requests. Approve **all** of the following:

| Permission | Purpose |
|------------|---------|
| `User.Read` | Current user profile, user photo in header |
| `User.Read.All` | People search enrichment (org chart, profiles) |
| `People.Read` | People relevance API |
| `People.Read.All` | Extended people search |
| `Sites.Read.All` | Graph Search API (search results, news) |
| `Files.Read.All` | Copilot Retrieval API (file content) |
| `ExternalItem.Read.All` | Search external connectors (Graph Connectors) |
| `Directory.Read.All` | Copilot license detection |
| `Mail.Read` | Copilot Chat API grounding (email context) |
| `Chat.Read` | Copilot Chat API grounding (Teams chat context) |
| `ChannelMessage.Read.All` | Copilot Chat API grounding (Teams channels) |
| `OnlineMeetingTranscript.Read.All` | Copilot Meeting Insights |

> **Warning:** If permissions are not approved, the Copilot Chat API, Retrieval API, and Graph Search calls will fail with HTTP 403. The web parts will show error messages.

---

## Step 7: Set Up the Bosch AI Search Page (Recommended)

This is the **primary deployment option** — a full-page search experience matching the Bosch AI Search design.

### Create the Page

1. Go to your Communication Site
2. Click **New** > **Page** > **Blank** (choose the blank template)
3. Set the page layout to **Full-width** (Page details > Page layout)
4. Add the **"Bosch AI Search"** web part from the **Bosch Search** category in the toolbox
5. The web part takes over the full page — no other web parts are needed on this page
6. **Publish** the page
7. Optionally, set this page as the site **Home page** (Site Settings > Welcome page)

### What You Get

The single web part renders the entire experience:

```
+------------------------------------------------------------------+
|  Bosch AI Search   [Copilot] [Images] [Videos] [More]    [User]  |  <- Header
+------------------------------------------------------------------+
|                                                                    |
|                          B O S C H                                 |  <- Logo
|                                                                    |
|                    [Work] [Web]                                    |  <- Scope toggle
|                    [🔍 Search with Bosch AI...    🎤 🤖]          |  <- Search box
|                                                                    |
+------------------------------------------------------------------+
|  Bosch News                                                        |
|  [Card] [Card] [Card] [Card] [Card] [Card] [Card] ...    < >      |  <- News carousel
+------------------------------------------------------------------+
|                                                      [💬]          |  <- Chat FAB (Copilot users)
+------------------------------------------------------------------+
```

After searching, the view transitions to:

```
+------------------------------------------------------------------+
|  Bosch AI Search   [Copilot] [Images] [Videos] [More]    [User]  |
+------------------------------------------------------------------+
|  [🔍 query text...                              ]                 |  <- Compact search box
+------------------------------------------------------------------+
|  📌 Promoted Result (if keywords match)                           |  <- From SP list
|  ─────────────────────────────────────                            |
|  🤖 Copilot Summary  [Copilot badge]                             |  <- AI answer
|  Answer text with citations [1] [2]...                            |
|  [Open in M365 Copilot]                                          |
|  ─────────────────────────────────────                            |
|  About 1,234 results                                              |  <- Graph Search results
|  Result Title                                                     |
|  https://source.url                                               |
|  Result summary text...                                           |
|  ...                                                              |
+------------------------------------------------------------------+
```

### Configure the Web Part

Click **Edit** on the page, then click the web part to open its property pane:

| Setting | Default | Description |
|---------|---------|-------------|
| **AI Grounding Mode** | Work | `Work` = enterprise only, `Web` = web search, `Both` = work + web |
| **Max retrieval results** | 10 | Number of results from Copilot Retrieval API |
| **Show "Open in Copilot" link** | Yes | Deep link to M365 Copilot app |
| **News source site URL** | _(current site)_ | Override to pull news from a different site |
| **Promoted results list** | SearchPromotedResults | SP list name for keyword-matched promoted links |
| **Announcements list** | SearchAnnouncements | SP list name for banner messages |
| **Analytics list** | SearchAnalyticsEvents | SP list name for query/click tracking |

---

## Step 7b: Alternative — Standalone Web Parts on Existing Page

If you prefer to use individual web parts alongside PnP Modern Search (the original approach), you can still add them to any SharePoint page:

### Recommended Layout (with PnP Modern Search)

```
+----------------------------------------------------------+
|  [PnP Search Box]                                        |
+----------------------------------------------------------+
|  [Announcements Banner]                                  |
+----------------------------------------------------------+
|  [AI Answer Panel]          |  [Sponsored Results]       |
+----------------------------------------------------------+
|  [PnP Search Results]       |  [People Search Enrichment]|
|  [PnP Refiners]             |                            |
+----------------------------------------------------------+
|  [Chat Assistant - side panel trigger button]             |
+----------------------------------------------------------+
```

### Standalone Web Part Configuration

| Web Part | Key Settings |
|----------|-------------|
| **AI Answer Panel** | Grounding Mode: `Work` or `Both`. Max retrieval results: 10. Show Copilot link: Yes. |
| **Chat Assistant** | Grounding Mode: `Work`. Display mode: `Side Panel` or `Inline`. Max turns: 10. |
| **Sponsored Results** | Reads from `SearchPromotedResults` list automatically. |
| **Announcements Banner** | Set the site URL and list name (`SearchAnnouncements`) in properties. |
| **People Search Enrichment** | Activates on people-related queries. |
| **Search Analytics** | Set the site URL and list name (`SearchAnalyticsEvents`). Typically on a separate admin page. |
| **Quick Links** | Add links via the property pane (title, URL, icon, description). |

---

## Step 8: Verify End-to-End

### Bosch AI Search Full-Page App

| Test | Expected Result |
|------|----------------|
| Page loads | Large BOSCH logo, Work/Web toggle, search box, news carousel |
| News carousel | Shows SharePoint news pages (or placeholder cards if no news) |
| Type a query and press Enter | Transitions to results view with AI answer + search results |
| Copilot user sees AI answer | **"Copilot Summary"** with purple badge, synthesized answer with citations |
| Non-Copilot user sees AI answer | **"Search Summary"** with green **"Graph Search"** badge, structured results |
| Search "hr" or "sales" | Promoted results appear above AI answer (blue card with pin icon) |
| Click Copilot Chat FAB (bottom-right) | Side panel opens with multi-turn chat |
| Click "Bosch AI Search" in header | Returns to landing page |
| Header verticals (Images/Videos) | Tab switches (vertical filtering) |
| User avatar in header | Shows user's profile photo from Graph |

### Copilot Integration (requires Copilot license)

| Test | Expected Result |
|------|----------------|
| Search any query | AI answer powered by Copilot Chat API |
| Open Chat panel | Multi-turn Copilot conversation with citations |
| "Open in M365 Copilot" link | Links to `https://m365.cloud.microsoft/chat` |
| Second chat message | Reuses same conversation (no new conversation created) |

### Non-Copilot User Fallback

| Test | Expected Result |
|------|----------------|
| Search any query | AI answer falls back to Retrieval API or Graph Search |
| Chat panel | Not visible (FAB hidden for non-Copilot users) |
| Results | Standard Graph Search results displayed |

---

## Troubleshooting

### Web parts not appearing in toolbox
- Verify the `.sppkg` is deployed and the app is added to the site
- Check that **"Make this solution available to all sites"** was checked during deployment

### 403 Forbidden errors in browser console
- Graph API permissions not approved. Go to SharePoint Admin > API access and approve all pending requests.

### AI answer shows errors
- For Copilot users: Ensure the user has an active **Microsoft 365 Copilot** license
- For non-Copilot users: Ensure `Sites.Read.All` is approved (needed for Graph Search fallback)

### News carousel shows placeholder cards
- Ensure `Sites.Read.All` is approved
- Ensure the site (or the configured news source URL) has published SharePoint news pages
- The carousel searches for pages with `PromotedState:2` (promoted news)

### Promoted results not showing
- Verify the `SearchPromotedResults` list exists on the site
- Check that items have `IsActive = Yes` and keywords match the search query

### User photo not loading in header
- `User.Read` permission must be approved
- Some users may not have a profile photo set

### Chat FAB not visible
- The floating chat button only appears for Copilot-licensed users
- Verify Copilot license is assigned and `Directory.Read.All` is approved

### Page has left navigation (not full-width)
- You're on a **Team Site**, not a Communication Site
- Create a new Communication Site or use a full-width section layout

### Lists not provisioned
- Ensure PnP.PowerShell module is installed: `Install-Module PnP.PowerShell`
- Use `-Interactive` auth (the script handles this automatically)

---

## Architecture Overview

### Bosch AI Search Full-Page App (Primary)

```
User opens Bosch AI Search page
       |
       v
Landing View: BOSCH logo + search box + news carousel
       |
       v (user types query)
Results View:
       |
       +---> AI Answer
       |       |
       |       +--> Copilot user? --> CopilotChatService.askSingleQuestion()
       |       +--> No Copilot?  --> CopilotRetrievalService (pay-as-you-go)
       |                              --> GraphSearchService (final fallback)
       |
       +---> Graph Search Results --> GraphSearchService.search()
       |
       +---> Promoted Results --> SharePointListService (keyword match)
       |
       +---> Chat Panel (Copilot users only, multi-turn)
               |
               +--> CopilotChatService (conversation API)
```

### Standalone Web Parts (Alternative)

```
User types query in PnP Search Box
       |
       v
URL updated (?q=searchterm)
       |
       v
useSearchQuery hook detects change (polls every 200ms)
       |
       v
EventBus emits 'searchQueryChanged'
       |
       +---> AI Answer Panel
       +---> Sponsored Results --> SharePointListService (keyword match)
       +---> People Search --> PeopleGraphService (Graph People API)
       +---> Analytics --> AnalyticsTrackingService (batched writes)
       +---> Chat Assistant (user-initiated, multi-turn)
```

---

## Local Development

```bash
# Start the local workbench (opens browser)
gulp serve

# The workbench URL is:
# https://m365cpi90282478.sharepoint.com/sites/Bosch/_layouts/workbench.aspx
```

Add web parts from the toolbox in the workbench to test individually. Note that Copilot APIs require a real user context and will not work in the local workbench — deploy to the App Catalog for full testing.

---

## Solution Contents

| Web Part | Description |
|----------|-------------|
| **Bosch AI Search** | Full-page search experience: landing page, search, AI answers, results, news, chat |
| **Quick Links** | Configurable tile grid of shortcut links |
| **Announcements Banner** | Dismissible alert banners from SharePoint list |
| **Sponsored Results** | Keyword-matched promoted results from SharePoint list |
| **People Search Enrichment** | Enhanced people cards with org chart and recent docs |
| **Search Analytics** | Dashboard with top queries, zero-result tracking, CTR |
| **AI Answer Panel** | AI-powered answers via Copilot Chat API / Graph Search fallback |
| **Chat Assistant** | Multi-turn conversational search via Copilot Chat API / Graph Search fallback |

---

## Deployment Checklist

- [ ] Communication Site created (for full-page app)
- [ ] `npm install` completed
- [ ] `gulp build` passes without errors
- [ ] `gulp bundle --ship && gulp package-solution --ship` generates `.sppkg`
- [ ] SharePoint lists provisioned via `Deploy-Lists.ps1`
- [ ] `.sppkg` uploaded to App Catalog and deployed
- [ ] All 12 Graph API permissions approved in SharePoint Admin
- [ ] Bosch AI Search web part added to a full-width page
- [ ] Page published and set as home page (optional)
- [ ] AI answer works for Copilot user
- [ ] AI answer falls back for non-Copilot user
- [ ] Promoted results appear for matching keywords
- [ ] News carousel loads content
- [ ] Chat panel opens and works (Copilot users)

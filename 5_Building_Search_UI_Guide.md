# Building Custom Search UI for Microsoft 365 Enterprise Search

## Technology Stack Overview

### Option 1: SharePoint Search with Modern Web Parts (Easiest)

**Best for**: Quick deployment, no coding required

**Components**:
- Search Results Web Part
- Search Box Web Part
- Refinement Web Part
- Search Navigation Web Part

**Pros**:
- No code required
- Built-in authentication
- Native M365 integration
- Adaptive Cards support

**Cons**:
- Limited customization
- SharePoint dependency
- Fixed layout options

### Option 2: SPFx (SharePoint Framework) React Web Parts (Recommended)

**Best for**: Custom M365-integrated search experiences

**Tech Stack**:
- React/TypeScript
- SPFx Framework
- Microsoft Graph SDK
- Fluent UI React

**Pros**:
- Full customization
- Modern React development
- Deployable to Teams, SharePoint, Outlook
- Integrated authentication

**Cons**:
- Requires development skills
- Deployment complexity

### Option 3: Standalone React/Angular App (Most Flexible)

**Best for**: Fully custom enterprise portals

**Tech Stack**:
- React/Angular/Vue
- Microsoft Graph SDK
- MSAL (Microsoft Authentication Library)
- Azure Static Web Apps or App Service

**Pros**:
- Complete control
- Can be standalone portal
- Modern frameworks

**Cons**:
- Complex authentication setup
- Hosting required
- More maintenance

### Option 4: Power Apps (Low Code)

**Best for**: Citizen developers, rapid prototypes

**Components**:
- Power Apps Canvas/Model-driven
- Power Automate for workflows
- Dataverse/SharePoint lists

**Pros**:
- Low code/no code
- Fast development
- Built-in connectors

**Cons**:
- Limited Graph Connector support
- Performance limitations
- Licensing costs

## Detailed Implementation Guides

### Solution 1: SPFx React Search Web Part

#### 1.1 Set Up Development Environment

```powershell
# Install Node.js (v16 or v18 LTS)
# Download from: https://nodejs.org/

# Install Yeoman and Gulp
npm install -g yo gulp

# Install SPFx Generator
npm install -g @microsoft/generator-sharepoint

# Verify installation
yo --version
gulp --version
```

#### 1.2 Create SPFx Project

```powershell
# Create new SPFx project
yo @microsoft/sharepoint

# Configuration:
# - Solution name: enterprise-search-webpart
# - Baseline: SharePoint Online only
# - Location: Current folder
# - Component type: WebPart
# - WebPart name: EnterpriseSearch
# - Template: React
```

#### 1.3 Install Dependencies

```powershell
cd enterprise-search-webpart

# Install Microsoft Graph SDK
npm install @microsoft/microsoft-graph-client

# Install Fluent UI
npm install @fluentui/react

# Install MSAL for authentication
npm install @azure/msal-browser
```

#### 1.4 Create Search Component

```typescript
// src/webparts/enterpriseSearch/components/EnterpriseSearch.tsx

import * as React from 'react';
import { useState, useEffect } from 'react';
import { SearchBox, Stack, List, Text, Spinner } from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';

interface ISearchResult {
  id: string;
  title: string;
  summary: string;
  url: string;
  source: string;
}

export const EnterpriseSearch: React.FC<{ graphClient: MSGraphClientV3 }> = ({ graphClient }) => {
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');

  const executeSearch = async (query: string): Promise<void> => {
    if (!query || query.trim() === '') {
      setResults([]);
      return;
    }

    setLoading(true);
    setError('');

    try {
      const searchRequest = {
        requests: [
          {
            entityTypes: ['externalItem', 'driveItem', 'listItem'],
            query: {
              queryString: query
            },
            from: 0,
            size: 25,
            fields: ['title', 'summary', 'url', 'contentSource']
          }
        ]
      };

      const response = await graphClient
        .api('/search/query')
        .post(searchRequest);

      const hits = response.value[0]?.hitsContainers[0]?.hits || [];
      
      const searchResults: ISearchResult[] = hits.map((hit: any) => ({
        id: hit.resource.id,
        title: hit.resource.properties.title || 'Untitled',
        summary: hit.summary || hit.resource.properties.summary || '',
        url: hit.resource.webUrl || hit.resource.properties.url || '#',
        source: hit.resource.properties.contentSource || 'Unknown'
      }));

      setResults(searchResults);
    } catch (err) {
      console.error('Search error:', err);
      setError(`Search failed: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const onSearchQueryChange = (newValue?: string): void => {
    setSearchQuery(newValue || '');
  };

  const onSearch = (newValue: string): void => {
    executeSearch(newValue);
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <SearchBox
        placeholder="Search across all enterprise content..."
        onSearch={onSearch}
        onChange={(_, newValue) => onSearchQueryChange(newValue)}
        value={searchQuery}
      />

      {loading && <Spinner label="Searching..." />}

      {error && <Text style={{ color: 'red' }}>{error}</Text>}

      {!loading && results.length === 0 && searchQuery && (
        <Text>No results found for "{searchQuery}"</Text>
      )}

      {!loading && results.length > 0 && (
        <List
          items={results}
          onRenderCell={(item: ISearchResult) => (
            <Stack
              tokens={{ childrenGap: 5 }}
              style={{
                padding: '15px',
                borderBottom: '1px solid #edebe9',
                cursor: 'pointer'
              }}
              onClick={() => window.open(item.url, '_blank')}
            >
              <Text variant="large" style={{ color: '#0078d4', fontWeight: 600 }}>
                {item.title}
              </Text>
              <Text variant="small" style={{ color: '#666' }}>
                {item.summary}
              </Text>
              <Text variant="tiny" style={{ color: '#999' }}>
                Source: {item.source} • {item.url}
              </Text>
            </Stack>
          )}
        />
      )}
    </Stack>
  );
};
```

#### 1.5 Advanced Search with Refiners

```typescript
// EnterpriseSearchAdvanced.tsx - With filters and refiners

import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  SearchBox,
  Stack,
  List,
  Text,
  Spinner,
  Checkbox,
  Panel,
  IconButton,
  IStackTokens
} from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';

interface IRefiner {
  field: string;
  displayName: string;
  values: Array<{ value: string; count: number; selected: boolean }>;
}

export const EnterpriseSearchAdvanced: React.FC<{ graphClient: MSGraphClientV3 }> = ({ graphClient }) => {
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [results, setResults] = useState<any[]>([]);
  const [refiners, setRefiners] = useState<IRefiner[]>([]);
  const [selectedFilters, setSelectedFilters] = useState<Map<string, string[]>>(new Map());
  const [loading, setLoading] = useState<boolean>(false);
  const [showFilters, setShowFilters] = useState<boolean>(false);

  const executeSearchWithRefiners = async (query: string, filters?: Map<string, string[]>): Promise<void> => {
    if (!query || query.trim() === '') return;

    setLoading(true);

    try {
      // Build filter string
      const filterArray: string[] = [];
      if (filters) {
        filters.forEach((values, field) => {
          if (values.length > 0) {
            const filterValues = values.map(v => `"${v}"`).join(' OR ');
            filterArray.push(`${field}:(${filterValues})`);
          }
        });
      }

      const searchRequest = {
        requests: [
          {
            entityTypes: ['externalItem'],
            query: {
              queryString: query
            },
            from: 0,
            size: 25,
            aggregations: [
              {
                field: 'category',
                size: 10,
                bucketDefinition: {
                  sortBy: 'count',
                  isDescending: true
                }
              },
              {
                field: 'department',
                size: 10
              },
              {
                field: 'fileType',
                size: 10
              }
            ],
            aggregationFilters: filterArray
          }
        ]
      };

      const response = await graphClient
        .api('/search/query')
        .post(searchRequest);

      const hitsContainer = response.value[0]?.hitsContainers[0];
      const hits = hitsContainer?.hits || [];
      const aggregations = hitsContainer?.aggregations || [];

      setResults(hits.map((hit: any) => hit.resource));

      // Process refiners
      const processedRefiners: IRefiner[] = aggregations.map((agg: any) => ({
        field: agg.field,
        displayName: agg.field.charAt(0).toUpperCase() + agg.field.slice(1),
        values: agg.buckets.map((bucket: any) => ({
          value: bucket.key,
          count: bucket.count,
          selected: selectedFilters.get(agg.field)?.includes(bucket.key) || false
        }))
      }));

      setRefiners(processedRefiners);
    } catch (err) {
      console.error('Search error:', err);
    } finally {
      setLoading(false);
    }
  };

  const handleRefinerChange = (field: string, value: string, checked: boolean): void => {
    const newFilters = new Map(selectedFilters);
    const currentValues = newFilters.get(field) || [];

    if (checked) {
      newFilters.set(field, [...currentValues, value]);
    } else {
      newFilters.set(field, currentValues.filter(v => v !== value));
    }

    setSelectedFilters(newFilters);
    executeSearchWithRefiners(searchQuery, newFilters);
  };

  const stackTokens: IStackTokens = { childrenGap: 15 };

  return (
    <Stack horizontal tokens={stackTokens}>
      {/* Refiners Panel */}
      <Stack style={{ width: '250px', borderRight: '1px solid #edebe9', paddingRight: '15px' }}>
        <Text variant="large" style={{ marginBottom: '10px', fontWeight: 600 }}>
          Filters
        </Text>
        
        {refiners.map(refiner => (
          <Stack key={refiner.field} tokens={{ childrenGap: 5 }} style={{ marginBottom: '20px' }}>
            <Text variant="medium" style={{ fontWeight: 600 }}>
              {refiner.displayName}
            </Text>
            {refiner.values.map(val => (
              <Checkbox
                key={val.value}
                label={`${val.value} (${val.count})`}
                checked={val.selected}
                onChange={(_, checked) => handleRefinerChange(refiner.field, val.value, checked || false)}
              />
            ))}
          </Stack>
        ))}
      </Stack>

      {/* Results Area */}
      <Stack grow tokens={stackTokens} style={{ paddingLeft: '15px' }}>
        <SearchBox
          placeholder="Search enterprise content..."
          onSearch={(query) => executeSearchWithRefiners(query, selectedFilters)}
          value={searchQuery}
          onChange={(_, newValue) => setSearchQuery(newValue || '')}
        />

        {loading && <Spinner label="Searching..." />}

        {!loading && results.length > 0 && (
          <>
            <Text variant="medium">
              Found {results.length} results
            </Text>
            <List
              items={results}
              onRenderCell={(item: any) => (
                <Stack
                  tokens={{ childrenGap: 5 }}
                  style={{
                    padding: '15px',
                    borderBottom: '1px solid #edebe9',
                    cursor: 'pointer'
                  }}
                  onClick={() => window.open(item.properties.url, '_blank')}
                >
                  <Text variant="large" style={{ color: '#0078d4', fontWeight: 600 }}>
                    {item.properties.title}
                  </Text>
                  <Text variant="small">
                    {item.properties.summary || item.properties.description}
                  </Text>
                  <Stack horizontal tokens={{ childrenGap: 10 }}>
                    {item.properties.category && (
                      <Text variant="tiny" style={{ color: '#666' }}>
                        Category: {item.properties.category}
                      </Text>
                    )}
                    {item.properties.department && (
                      <Text variant="tiny" style={{ color: '#666' }}>
                        Department: {item.properties.department}
                      </Text>
                    )}
                  </Stack>
                </Stack>
              )}
            />
          </>
        )}
      </Stack>
    </Stack>
  );
};
```

#### 1.6 Build and Deploy

```powershell
# Build the solution
gulp build

# Bundle for production
gulp bundle --ship

# Package the solution
gulp package-solution --ship

# Deploy to App Catalog
# 1. Go to SharePoint Admin Center → More features → Apps → App Catalog
# 2. Upload the .sppkg file from sharepoint/solution folder
# 3. Deploy to all sites or specific sites
```

### Solution 2: Standalone React App with Graph API

#### 2.1 Create React App

```powershell
# Create new React app with TypeScript
npx create-react-app enterprise-search-portal --template typescript

cd enterprise-search-portal

# Install dependencies
npm install @azure/msal-browser @azure/msal-react
npm install @microsoft/microsoft-graph-client
npm install @fluentui/react
```

#### 2.2 Configure Azure AD App Registration

```powershell
# Register app in Azure Portal:
# 1. Azure AD → App registrations → New registration
#    Name: Enterprise Search Portal
#    Supported account types: Single tenant
#    Redirect URI: http://localhost:3000 (for dev)
#
# 2. API Permissions:
#    - Microsoft Graph → Delegated → User.Read
#    - Microsoft Graph → Delegated → Sites.Read.All
#    - Microsoft Graph → Delegated → ExternalItem.Read.All
#
# 3. Copy Application (client) ID and Directory (tenant) ID
```

#### 2.3 MSAL Configuration

```typescript
// src/authConfig.ts

import { Configuration, PopupRequest } from '@azure/msal-browser';

export const msalConfig: Configuration = {
  auth: {
    clientId: 'YOUR_CLIENT_ID',
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID',
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false
  }
};

export const loginRequest: PopupRequest = {
  scopes: ['User.Read', 'Sites.Read.All', 'ExternalItem.Read.All']
};

export const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',
  graphSearchEndpoint: 'https://graph.microsoft.com/v1.0/search/query'
};
```

#### 2.4 Main App Component

```typescript
// src/App.tsx

import React from 'react';
import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig, loginRequest } from './authConfig';
import { SearchComponent } from './components/SearchComponent';
import { PrimaryButton, Stack, Text } from '@fluentui/react';

const msalInstance = new PublicClientApplication(msalConfig);

const SignInButton: React.FC = () => {
  const { instance } = useMsal();

  const handleLogin = () => {
    instance.loginPopup(loginRequest).catch(e => {
      console.error(e);
    });
  };

  return (
    <Stack horizontalAlign="center" verticalAlign="center" style={{ height: '100vh' }}>
      <Text variant="xxLarge" style={{ marginBottom: '20px' }}>
        Enterprise Search Portal
      </Text>
      <PrimaryButton text="Sign In with Microsoft" onClick={handleLogin} />
    </Stack>
  );
};

const App: React.FC = () => {
  return (
    <MsalProvider instance={msalInstance}>
      <AuthenticatedTemplate>
        <SearchComponent />
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <SignInButton />
      </UnauthenticatedTemplate>
    </MsalProvider>
  );
};

export default App;
```

#### 2.5 Search Component with Graph Client

```typescript
// src/components/SearchComponent.tsx

import React, { useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { Client } from '@microsoft/microsoft-graph-client';
import { SearchBox, Stack, List, Spinner, Text } from '@fluentui/react';
import { loginRequest, graphConfig } from '../authConfig';

export const SearchComponent: React.FC = () => {
  const { instance, accounts } = useMsal();
  const [results, setResults] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);

  const getGraphClient = async (): Promise<Client> => {
    const account = accounts[0];
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account: account
    });

    return Client.init({
      authProvider: (done) => {
        done(null, response.accessToken);
      }
    });
  };

  const handleSearch = async (query: string): Promise<void> => {
    if (!query) return;

    setLoading(true);

    try {
      const client = await getGraphClient();

      const searchRequest = {
        requests: [
          {
            entityTypes: ['externalItem', 'driveItem', 'listItem', 'site'],
            query: {
              queryString: query
            },
            from: 0,
            size: 50
          }
        ]
      };

      const response = await client.api(graphConfig.graphSearchEndpoint).post(searchRequest);

      const hits = response.value[0]?.hitsContainers[0]?.hits || [];
      setResults(hits);
    } catch (error) {
      console.error('Search failed:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 20 }} style={{ padding: '40px', maxWidth: '1200px', margin: '0 auto' }}>
      <Text variant="xxLarge">Enterprise Search</Text>
      
      <SearchBox
        placeholder="Search across all content..."
        onSearch={handleSearch}
      />

      {loading && <Spinner label="Searching..." />}

      {!loading && results.length > 0 && (
        <List
          items={results}
          onRenderCell={(item: any) => (
            <Stack
              tokens={{ childrenGap: 8 }}
              style={{
                padding: '20px',
                borderBottom: '1px solid #edebe9',
                cursor: 'pointer'
              }}
              onClick={() => window.open(item.resource.webUrl, '_blank')}
            >
              <Text variant="large" style={{ color: '#0078d4', fontWeight: 600 }}>
                {item.resource.properties?.title || 'Untitled'}
              </Text>
              <Text variant="small">
                {item.summary || item.resource.properties?.summary || ''}
              </Text>
            </Stack>
          )}
        />
      )}
    </Stack>
  );
};
```

#### 2.6 Deploy to Azure Static Web Apps

```powershell
# Build for production
npm run build

# Install Azure CLI
# Download from: https://aka.ms/installazurecliwindows

# Login to Azure
az login

# Create static web app
az staticwebapp create \
  --name enterprise-search-portal \
  --resource-group rg-enterprise-search \
  --source ./build \
  --location "East US" \
  --branch main \
  --app-location "/" \
  --output-location "build"

# Get deployment URL
az staticwebapp show --name enterprise-search-portal --query "defaultHostname" -o tsv
```

## Copilot Search UI Differences

### Key Differences

1. **Adaptive Cards**: Not fully supported in Copilot search results
2. **Custom Refiners**: Limited in Copilot interface
3. **Result Templates**: Copilot uses AI-generated summaries instead of custom templates
4. **Citation Format**: Copilot shows citations differently than standard search

### Workaround for Copilot Limitations

```typescript
// Detect if running in Copilot context
function isCopilotContext(): boolean {
  return window.location.href.includes('microsoft365.com/chat') || 
         window.location.href.includes('teams.microsoft.com/v2');
}

// Adjust UI based on context
const renderSearchUI = () => {
  if (isCopilotContext()) {
    // Simplified UI for Copilot
    return <SimplifiedSearchResults />;
  } else {
    // Full-featured UI with Adaptive Cards
    return <AdvancedSearchResults />;
  }
};
```

## Next Steps

1. Choose technology stack based on requirements
2. Set up development environment
3. Implement authentication
4. Build search interface
5. Test with Graph Connectors data
6. Deploy to production
7. Gather user feedback and iterate

This guide provides complete code samples for each approach!

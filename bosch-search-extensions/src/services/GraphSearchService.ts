import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ISearchResult, ISearchResponse, ISpellingSuggestion } from '../models';
import { SimpleCache } from '../common/Utils';

const CACHE_TTL_MS = 60 * 1000; // 1 minute

export class GraphSearchService {
  private graphClient: MSGraphClientV3;
  private cache: SimpleCache<ISearchResponse> = new SimpleCache();

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  public async search(
    query: string,
    entityTypes: string[] = ['driveItem', 'listItem'],
    from: number = 0,
    size: number = 25,
    aggregations?: unknown[]
  ): Promise<ISearchResponse> {
    if (!query || query.trim() === '') {
      return { results: [], totalResults: 0, moreResultsAvailable: false };
    }

    const cacheKey = `${query}|${entityTypes.join(',')}|${from}|${size}`;
    const cached = this.cache.get(cacheKey);
    if (cached) return cached;

    // Filter out externalItem — it requires contentSources and fails with 400
    // if no Graph Connectors are configured on the tenant.
    // Use searchExternalItems() explicitly when connectors are available.
    const safeEntityTypes = entityTypes.filter((t) => t !== 'externalItem');
    if (safeEntityTypes.length === 0) {
      safeEntityTypes.push('driveItem', 'listItem');
    }

    const searchRequest: Record<string, unknown> = {
      requests: [
        {
          entityTypes: safeEntityTypes,
          query: { queryString: query },
          from,
          size,
          fields: ['title', 'summary', 'url', 'contentSource', 'fileType', 'lastModifiedDateTime', 'createdBy', 'department', 'category'],
          ...(aggregations ? { aggregations } : {}),
        },
      ],
    };

    console.log('[GraphSearchService] POST /search/query', JSON.stringify(searchRequest, null, 2));
    const response = await this.graphClient
      .api('/search/query')
      .post(searchRequest);

    const hitsContainer = response.value?.[0]?.hitsContainers?.[0];
    const hits = hitsContainer?.hits || [];
    const totalResults = hitsContainer?.total || 0;
    const moreResultsAvailable = hitsContainer?.moreResultsAvailable || false;

    console.log('[GraphSearchService] Raw response:', JSON.stringify(response, null, 2));

    const results: ISearchResult[] = hits.map((hit: Record<string, unknown>, idx: number) => {
      const resource = hit.resource as Record<string, unknown>;
      const properties = (resource.properties || {}) as Record<string, string>;
      // Graph Search returns the title in various locations depending on entity type
      const listItem = resource.listItem as Record<string, unknown> | undefined;
      const listItemFields = (listItem?.fields || {}) as Record<string, string>;
      const title =
        properties.title || properties.Title ||
        (resource.name as string) || properties.name || properties.Name ||
        (resource.displayName as string) || properties.displayName || properties.DisplayName ||
        properties.subject || properties.Subject ||
        listItemFields.title || listItemFields.Title ||
        (resource['@odata.type'] === '#microsoft.graph.driveItem'
          ? (resource.name as string)
          : '') ||
        'Untitled';

      // Build URL from multiple fallback sources
      const parentRef = resource.parentReference as Record<string, string> | undefined;
      const resolvedUrl =
        (resource.webUrl as string) ||
        properties.url || properties.Url || properties.URL ||
        properties.webUrl || properties.WebUrl ||
        properties.link || properties.Link ||
        properties.path || properties.Path ||
        (listItem?.webUrl as string) ||
        (parentRef?.siteUrl ? `${parentRef.siteUrl}/${resource.name || ''}` : '') ||
        '#';

      console.log(`[GraphSearchService] Hit[${idx}]: title="${title}", url="${resolvedUrl}", webUrl="${resource.webUrl}", props.url="${properties.url}", props.path="${properties.path}"`);

      return {
        id: resource.id as string,
        title,
        summary: (hit.summary as string) || properties.summary || '',
        url: resolvedUrl,
        source: properties.contentSource || 'SharePoint',
        fileType: properties.fileType,
        lastModified: properties.lastModifiedDateTime,
        author: properties.createdBy,
        department: properties.department,
        category: properties.category,
        hitHighlightedSummary: hit.summary as string,
      };
    });

    const searchResponse: ISearchResponse = { results, totalResults, moreResultsAvailable };
    this.cache.set(cacheKey, searchResponse, CACHE_TTL_MS);
    return searchResponse;
  }

  public async searchPeople(query: string): Promise<ISearchResult[]> {
    const response = await this.search(query, ['person'], 0, 10);
    return response.results;
  }

  public async searchExternalItems(query: string, size: number = 10): Promise<ISearchResult[]> {
    const response = await this.search(query, ['externalItem'], 0, size);
    return response.results;
  }

  /**
   * Search within a specific M365 source for source-specific paginated results.
   * Routes SharePoint / Outlook Mail / Teams to entity-type searches,
   * and external connector names to contentSources-filtered externalItem searches.
   *
   * @param source  Source name matching the sidebar label (e.g. "SharePoint", "CustomConnector1")
   */
  public async searchBySource(
    query: string,
    source: string,
    from: number = 0,
    size: number = 25
  ): Promise<ISearchResponse> {
    if (!query || query.trim() === '') {
      return { results: [], totalResults: 0, moreResultsAvailable: false };
    }

    let entityTypes: string[];
    let contentSources: string[] | undefined;

    switch (source) {
      case 'SharePoint':
        entityTypes = ['driveItem', 'listItem', 'site'];
        break;
      case 'Outlook Mail':
        entityTypes = ['message'];
        break;
      case 'Teams':
        entityTypes = ['chatMessage'];
        break;
      default:
        // External connector — source label is the connection ID
        entityTypes = ['externalItem'];
        contentSources = [`/external/connections/${source}`];
        break;
    }

    const requestBody: Record<string, unknown> = {
      entityTypes,
      query: { queryString: query },
      from,
      size,
      fields: ['title', 'name', 'summary', 'url', 'webUrl', 'contentSource', 'fileType', 'lastModifiedDateTime', 'createdBy'],
    };
    if (contentSources) requestBody.contentSources = contentSources;

    const request = { requests: [requestBody] };
    console.log(`[GraphSearchService] searchBySource [${source}] from=${from}`, JSON.stringify(request));

    const response = await this.graphClient.api('/search/query').post(request);
    const container = response.value?.[0]?.hitsContainers?.[0];
    if (!container?.hits) {
      return { results: [], totalResults: 0, moreResultsAvailable: false };
    }

    const hits = container.hits as Record<string, unknown>[];
    const results: ISearchResult[] = hits.map((hit) => {
      const resource = (hit.resource || {}) as Record<string, unknown>;
      const properties = (resource.properties || {}) as Record<string, string>;
      const odataType = (resource['@odata.type'] as string) || '';

      // Source label — default to requested source, normalize connector URIs
      let sourceLabel = source;
      if (odataType.includes('message')) sourceLabel = 'Outlook Mail';
      else if (odataType.includes('chatMessage')) sourceLabel = 'Teams';
      else if (odataType.includes('externalItem')) {
        const rawSource =
          properties.contentSource || properties.ContentSource ||
          (hit.contentSource as string) || source;
        sourceLabel = rawSource.includes('/external/connections/')
          ? rawSource.split('/').filter(Boolean).pop() || source
          : rawSource;
      }

      const listItem = resource.listItem as Record<string, unknown> | undefined;
      const listItemFields = (listItem?.fields || {}) as Record<string, string>;
      const title =
        (resource.subject as string) ||
        properties.title || properties.Title ||
        (resource.name as string) || properties.name || properties.Name ||
        (resource.displayName as string) || properties.displayName || properties.DisplayName ||
        listItemFields.title || listItemFields.Title ||
        'Untitled';

      const parentRef = resource.parentReference as Record<string, string> | undefined;
      const resolvedUrl =
        (resource.webUrl as string) ||
        (resource.webLink as string) ||
        properties.url || properties.Url || properties.URL ||
        properties.webUrl || properties.WebUrl ||
        properties.link || properties.Link ||
        (listItem?.webUrl as string) ||
        (parentRef?.siteUrl ? `${parentRef.siteUrl}/${resource.name || ''}` : '') ||
        '#';

      return {
        id: (resource.id as string) || '',
        title,
        summary: (hit.summary as string) || properties.summary || '',
        url: resolvedUrl,
        source: sourceLabel,
        fileType: properties.fileType || (odataType.includes('message') ? 'email' : undefined),
        lastModified: properties.lastModifiedDateTime || (resource.lastModifiedDateTime as string),
        author: properties.createdBy,
        hitHighlightedSummary: (hit.summary as string) || '',
      } as ISearchResult;
    });

    return {
      results,
      totalResults: container.total || 0,
      moreResultsAvailable: container.moreResultsAvailable || false,
    };
  }

  /**
   * Discovers all available Microsoft Graph external connections (Graph Connectors).
   * Returns array of { id, name } for each connection.
   */
  private async getExternalConnections(): Promise<{ id: string; name: string }[]> {
    try {
      const response = await this.graphClient
        .api('/external/connections')
        .select('id,name')
        .get();

      const connections = (response.value || []).map((c: Record<string, string>) => ({
        id: c.id,
        name: c.name || c.id,
      }));
      console.log(`[GraphSearchService] Found ${connections.length} external connections:`, connections.map((c: { id: string; name: string }) => c.name).join(', '));
      return connections;
    } catch (err) {
      console.warn('[GraphSearchService] Failed to get external connections:', (err as Error).message || err);
      return [];
    }
  }

  /**
   * Searches across ALL M365 entity types (files, email, chat, sites, external items)
   * using **independent parallel requests** so one failure doesn't lose all results.
   *
   * Graph Search API (`POST /search/query`) rules:
   *   - driveItem, listItem, site can share one request
   *   - message needs a separate request (Mail.Read)
   *   - chatMessage needs a separate request (Chat.Read / ChannelMessage.Read.All)
   *   - externalItem needs a separate request (ExternalItem.Read.All)
   *
   * Each is sent as a separate HTTP call via Promise.allSettled.
   * Failed calls are logged but don't block other sources.
   */
  public async searchAll(
    query: string,
    from: number = 0,
    size: number = 25
  ): Promise<ISearchResponse> {
    if (!query || query.trim() === '') {
      return { results: [], totalResults: 0, moreResultsAvailable: false };
    }

    const cacheKey = `all|${query}|${from}|${size}`;
    const cached = this.cache.get(cacheKey);
    if (cached) return cached;

    // Helper: send a single search request for one entity type group
    const searchEntityTypes = async (
      entityTypes: string[],
      requestFrom: number,
      requestSize: number,
      label: string,
      contentSources?: string[]
    ): Promise<{ results: ISearchResult[]; total: number; more: boolean }> => {
      const requestBody: Record<string, unknown> = {
        entityTypes,
        query: { queryString: query },
        from: requestFrom,
        size: requestSize,
      };
      if (contentSources && contentSources.length > 0) {
        requestBody.contentSources = contentSources;
      }

      const request: Record<string, unknown> = {
        requests: [requestBody],
      };

      console.log(`[GraphSearchService] searchAll: POST /search/query [${label}]`, JSON.stringify(request));
      const response = await this.graphClient
        .api('/search/query')
        .post(request);

      const container = response.value?.[0]?.hitsContainers?.[0];
      if (!container?.hits) {
        console.log(`[GraphSearchService] searchAll [${label}]: 0 hits`);
        return { results: [], total: 0, more: false };
      }

      console.log(`[GraphSearchService] searchAll [${label}]: ${container.total} total, ${container.hits.length} returned`);

      const results = container.hits.map((hit: Record<string, unknown>) => {
        const resource = hit.resource as Record<string, unknown>;
        const odataType = (resource['@odata.type'] as string) || '';
        const properties = (resource.properties || {}) as Record<string, string>;

        // Source label — sites are SharePoint content, labelled accordingly
        // For per-connector searches, the label carries the connector name as fallback
        let sourceLabel = 'SharePoint';
        const connectorFallback = label.startsWith('external:') ? label.replace('external:', '') : undefined;
        if (odataType.includes('message')) sourceLabel = 'Outlook Mail';
        else if (odataType.includes('chatMessage')) sourceLabel = 'Teams';
        else if (odataType.includes('externalItem')) sourceLabel = properties.contentSource || properties.ContentSource || (resource.contentSource as string) || connectorFallback || 'External';
        // 'site' odataType = SharePoint site — label as SharePoint (not 'Other Sites')
        else if (properties.contentSource) sourceLabel = properties.contentSource;

        // Title — try all common casing variants since connector schemas vary
        const listItem = resource.listItem as Record<string, unknown> | undefined;
        const listItemFields = (listItem?.fields || {}) as Record<string, string>;
        const title =
          (resource.subject as string) ||
          properties.title || properties.Title ||
          (resource.name as string) || properties.name || properties.Name ||
          (resource.displayName as string) || properties.displayName || properties.DisplayName ||
          properties.subject || properties.Subject ||
          listItemFields.title || listItemFields.Title ||
          'Untitled';

        // URL — emails use webLink, files use webUrl; connectors may use Url / URL
        const parentRef = resource.parentReference as Record<string, string> | undefined;
        const resolvedUrl =
          (resource.webUrl as string) ||
          (resource.webLink as string) ||
          properties.url || properties.Url || properties.URL ||
          properties.webUrl || properties.WebUrl ||
          properties.link || properties.Link ||
          properties.path || properties.Path ||
          (listItem?.webUrl as string) ||
          (parentRef?.siteUrl ? `${parentRef.siteUrl}/${resource.name || ''}` : '') ||
          '#';

        // Summary
        const summary =
          (hit.summary as string) ||
          (resource.bodyPreview as string) ||
          properties.summary || '';

        return {
          id: resource.id as string,
          title,
          summary,
          url: resolvedUrl,
          source: sourceLabel,
          contentSource: properties.contentSource,
          fileType: properties.fileType || (odataType.includes('message') ? 'email' : undefined),
          lastModified: properties.lastModifiedDateTime || (resource.lastModifiedDateTime as string),
          author: properties.createdBy ||
            ((resource.from as Record<string, unknown>)?.emailAddress as Record<string, string>)?.name,
          department: properties.department,
          category: properties.category,
          hitHighlightedSummary: (hit.summary as string) || '',
        };
      });

      return {
        results,
        total: container.total || 0,
        more: container.moreResultsAvailable || false,
      };
    };

    // Discover external connections for Graph Connector search
    const connections = await this.getExternalConnections();

    // Build parallel search requests: core M365 sources + one per connector
    const searchPromises: Promise<{ results: ISearchResult[]; total: number; more: boolean }>[] = [
      searchEntityTypes(['driveItem', 'listItem', 'site'], from, size, 'files+sites'),
      searchEntityTypes(['message'], 0, size, 'email'),
      searchEntityTypes(['chatMessage'], 0, size, 'teams'),
    ];
    // labels[0..2] = core sources; labels[3+] = per-connector
    const labels: string[] = ['files+sites', 'email', 'teams'];

    if (connections.length > 0) {
      // Search each connector separately for accurate per-connector totals
      connections.forEach((conn) => {
        labels.push(`connector:${conn.name}`);
        searchPromises.push(
          searchEntityTypes(
            ['externalItem'],
            0,
            size,
            `external:${conn.name}`,
            [`/external/connections/${conn.id}`]
          )
        );
      });
    } else {
      // No connections found — try a generic externalItem search as fallback
      labels.push('external');
      searchPromises.push(searchEntityTypes(['externalItem'], 0, size, 'external'));
    }

    // Fire ALL searches in parallel
    const searches = await Promise.allSettled(searchPromises);

    let allResults: ISearchResult[] = [];
    let grandTotal = 0;
    let anyMore = false;

    // Build per-source total counts from API totals (not page-limited)
    const sourceCounts: Record<string, number> = {};

    // Map from entity type group label → default source name
    const labelToDefaultSource: Record<string, string> = {
      'email': 'Outlook Mail',
      'teams': 'Teams',
    };

    searches.forEach((result, i) => {
      if (result.status === 'fulfilled') {
        const { results: groupResults, total, more } = result.value;
        allResults = allResults.concat(groupResults);
        grandTotal += total;
        if (more) anyMore = true;

        const label = labels[i];
        if (label.startsWith('connector:')) {
          // Per-connector search: use the connector name as source, total is exact
          const connectorName = label.replace('connector:', '');
          sourceCounts[connectorName] = total;
        } else if (labelToDefaultSource[label]) {
          // Single-source groups (email, teams): API total is accurate
          sourceCounts[labelToDefaultSource[label]] = total;
        } else {
          // Multi-source groups (files+sites): compute per-source
          // proportions from returned results, scale to API total.
          // 'Other Sites' (site odataType) is already relabelled as 'SharePoint' above.
          const pageCounts: Record<string, number> = {};
          groupResults.forEach((r) => {
            const src = r.source || 'Other';
            pageCounts[src] = (pageCounts[src] || 0) + 1;
          });
          const pageTotal = groupResults.length;
          if (pageTotal > 0 && total > 0) {
            Object.keys(pageCounts).forEach((src) => {
              const ratio = pageCounts[src] / pageTotal;
              sourceCounts[src] = Math.round(ratio * total);
            });
          }
        }
      } else {
        console.warn(`[GraphSearchService] searchAll [${labels[i]}] failed:`, result.reason?.message || result.reason);
      }
    });

    console.log(`[GraphSearchService] searchAll TOTAL: ${grandTotal} results from ${searches.filter(r => r.status === 'fulfilled').length}/${searches.length} sources`);
    console.log('[GraphSearchService] sourceCounts:', JSON.stringify(sourceCounts));

    const searchResponse: ISearchResponse = {
      results: allResults,
      totalResults: grandTotal,
      moreResultsAvailable: anyMore,
      sourceCounts,
    };
    this.cache.set(cacheKey, searchResponse, CACHE_TTL_MS);
    return searchResponse;
  }

  /**
   * Checks the given query for spelling errors using the Graph Search
   * `queryAlterationOptions` feature.
   *
   * - enableSuggestion: suggests a corrected term when typos are detected.
   * - enableModification: if no results for the original query, the API
   *   automatically searches the corrected term and returns those results.
   *
   * Supported types: listItem, driveItem, site, message, event, externalItem.
   * NOT supported: chatMessage, person.
   *
   * Returns null if the query needs no correction or if the call fails.
   */
  public async checkSpelling(query: string): Promise<ISpellingSuggestion | null> {
    if (!query || query.trim().length < 3) return null;
    try {
      const request = {
        requests: [{
          entityTypes: ['listItem'],
          query: { queryString: query },
          from: 0,
          size: 1,
          queryAlterationOptions: {
            enableSuggestion: true,
            enableModification: true,
          },
        }],
      };
      const response = await this.graphClient.api('/search/query').post(request);
      const alteration = response.value?.[0]?.queryAlterationResponse;
      if (!alteration?.queryAlteration?.alteredQueryString) return null;
      const alteredQuery = (alteration.queryAlteration.alteredQueryString as string).trim();
      if (alteredQuery.toLowerCase() === query.toLowerCase().trim()) return null;
      return {
        originalQuery: query.trim(),
        alteredQuery,
        type: alteration.queryAlterationType === 'Modification' ? 'Modification' : 'Suggestion',
      };
    } catch {
      return null;
    }
  }
}

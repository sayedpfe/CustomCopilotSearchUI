import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ISearchResult, ISearchResponse } from '../models';
import { COPILOT_SEARCH_ENDPOINT } from '../common/Constants';
import { SimpleCache } from '../common/Utils';

const CACHE_TTL_MS = 60 * 1000; // 1 minute

/**
 * Service for the Microsoft 365 Copilot Search API.
 * POST /beta/copilot/search
 *
 * This is the same API that powers the M365 Copilot Search experience
 * (m365.cloud.microsoft/search). It searches across ALL M365 sources in a
 * SINGLE request: SharePoint, OneDrive, Outlook Mail, Teams, Graph Connectors,
 * M365 Copilot Chats, etc. — no need to split by entity type.
 *
 * Key advantage over /search/query:
 *   - externalItem (Graph Connectors) works WITHOUT specifying contentSources
 *   - All entity types can be combined in one request body
 *   - Returns per-source aggregated counts automatically
 *
 * Request format is identical to /search/query (requests array).
 * Response format is identical to /search/query (value[0].hitsContainers[0]).
 *
 * Requires Copilot license.
 * Permissions: Sites.Read.All, Mail.Read, Chat.Read, Files.Read.All,
 *              ExternalItem.Read.All
 */
export class CopilotSearchService {
  private graphClient: MSGraphClientV3;
  private cache: SimpleCache<ISearchResponse> = new SimpleCache();

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Search across all M365 sources using the Copilot Search API.
   * Returns the same breadth of results as M365 Copilot Search.
   */
  public async search(
    query: string,
    from: number = 0,
    size: number = 25
  ): Promise<ISearchResponse> {
    if (!query || query.trim() === '') {
      return { results: [], totalResults: 0, moreResultsAvailable: false };
    }

    const cacheKey = `copilot-search|${query}|${from}|${size}`;
    const cached = this.cache.get(cacheKey);
    if (cached) return cached;

    // /beta/copilot/search uses a FLAT format (NOT the requests[] array of /search/query)
    // Sending entityTypes tells it which content to search; omitting defaults to all sources.
    const body: Record<string, unknown> = {
      queryString: query,
      from,
      size,
    };

    console.log(`[CopilotSearchService] POST ${COPILOT_SEARCH_ENDPOINT} (beta)`, JSON.stringify(body, null, 2));

    const response = await this.graphClient
      .api(COPILOT_SEARCH_ENDPOINT)
      .version('beta')
      .post(body);

    console.log('[CopilotSearchService] Raw response keys:', Object.keys(response || {}));
    console.log('[CopilotSearchService] Raw response:', JSON.stringify(response, null, 2));

    // Handle both the flat response format and the nested /search/query-style format
    const hitsContainer = response?.value?.[0]?.hitsContainers?.[0];
    const hits: Record<string, unknown>[] =
      hitsContainer?.hits ||
      (Array.isArray(response.hits) ? response.hits : null) ||
      [];
    let totalResults: number =
      hitsContainer?.total ??
      response.total ??
      response.totalCount ??
      (response['@odata.count'] as number | undefined) ??
      hits.length;
    const moreResultsAvailable: boolean =
      hitsContainer?.moreResultsAvailable ??
      response.moreResultsAvailable ??
      false;

    const results: ISearchResult[] = hits.map((hit: Record<string, unknown>, idx: number) => {
      const resource = (hit.resource || hit) as Record<string, unknown>;
      const properties = (resource.properties || {}) as Record<string, string>;
      const odataType = (resource['@odata.type'] as string) || '';

      // Determine source — sites are SharePoint content, connectors use their display name
      let sourceLabel = 'SharePoint';
      if (odataType.includes('message')) sourceLabel = 'Outlook Mail';
      else if (odataType.includes('chatMessage')) sourceLabel = 'Teams';
      else if (odataType.includes('copilotChat') || odataType.includes('AiChat')) sourceLabel = 'M365 Copilot Chats';
      else if (odataType.includes('externalItem')) {
        // Connector hit — contentSource may be a full URI like "/external/connections/ConnectorId"
        const rawSource =
          properties.contentSource ||
          properties.ContentSource ||
          (resource.contentSource as string) ||
          (resource['@search.contentSource'] as string) ||
          (hit.contentSource as string) ||
          (hit['@search.contentSource'] as string) ||
          '';
        // Normalize: extract connector ID from paths like "/external/connections/ConnectorId"
        sourceLabel = rawSource
          ? rawSource.split('/').filter(Boolean).pop() || rawSource
          : 'External';
      }
      // 'site' odataType = SharePoint site — label as SharePoint (not 'Other Sites')
      else if (resource.contentSource) sourceLabel = resource.contentSource as string;
      else if (properties.contentSource) sourceLabel = properties.contentSource;

      // Title resolution — try all common casing variants since connector schemas vary
      const title =
        (resource.subject as string) ||
        (resource.title as string) ||
        properties.title || properties.Title ||
        (resource.name as string) || properties.name || properties.Name ||
        (resource.displayName as string) || properties.displayName || properties.DisplayName ||
        properties.subject || properties.Subject ||
        'Untitled';

      // URL resolution — connectors may use Url / URL / link
      const resolvedUrl =
        (resource.webUrl as string) ||
        (resource.webLink as string) ||
        properties.url || properties.Url || properties.URL ||
        properties.webUrl || properties.WebUrl ||
        properties.link || properties.Link ||
        properties.path || properties.Path ||
        '#';

      // Summary
      const summary =
        (hit.summary as string) ||
        (resource.bodyPreview as string) ||
        (resource.summary as string) ||
        properties.summary || '';

      console.log(`[CopilotSearchService] Hit[${idx}]: source="${sourceLabel}", title="${title}", url="${resolvedUrl}"`);

      return {
        id: (resource.id as string) || String(from + idx),
        title,
        summary,
        url: resolvedUrl,
        source: sourceLabel,
        fileType: properties.fileType || (odataType.includes('message') ? 'email' : undefined),
        lastModified: properties.lastModifiedDateTime || (resource.lastModifiedDateTime as string),
        author: properties.createdBy || ((resource.from as Record<string, unknown>)?.emailAddress as Record<string, string>)?.name,
        hitHighlightedSummary: (hit.summary as string) || '',
      };
    });

    // Build per-source total counts.
    // The API may return aggregations; otherwise derive proportionally from returned results.
    // 'Other Sites' (SharePoint site type) is merged into 'SharePoint'.
    const sourceCounts: Record<string, number> = {};
    const aggBuckets = (
      hitsContainer?.aggregations?.[0]?.buckets ||
      response.aggregations?.[0]?.buckets ||
      []
    ) as Record<string, unknown>[];
    const aggTotal = aggBuckets.reduce((sum, b) => sum + (Number(b.count) || 0), 0);
    // Override totalResults if we only had hits.length as fallback and aggBuckets sum is larger
    if (totalResults <= hits.length && aggTotal > hits.length) totalResults = aggTotal;
    aggBuckets.forEach((bucket) => {
      if (!bucket.key) return;
      let key = bucket.key as string;
      // Normalize full URI paths like "/external/connections/ConnectorId" → "ConnectorId"
      if (key.includes('/')) key = key.split('/').filter(Boolean).pop() || key;
      // Merge 'Other Sites' into 'SharePoint'
      if (key === 'Other Sites') key = 'SharePoint';
      sourceCounts[key] = (sourceCounts[key] || 0) + (Number(bucket.count) || 0);
    });

    // Fall back to proportional counts when no aggregation data
    if (Object.keys(sourceCounts).length === 0 && hits.length > 0 && totalResults > 0) {
      const pageCounts: Record<string, number> = {};
      results.forEach((r) => { pageCounts[r.source || 'Other'] = (pageCounts[r.source || 'Other'] || 0) + 1; });
      const pageTotal = results.length;
      Object.keys(pageCounts).forEach((src) => {
        sourceCounts[src] = Math.round((pageCounts[src] / pageTotal) * totalResults);
      });
    }

    console.log('[CopilotSearchService] sourceCounts:', JSON.stringify(sourceCounts));

    const searchResponse: ISearchResponse = { results, totalResults, moreResultsAvailable, sourceCounts };
    this.cache.set(cacheKey, searchResponse, CACHE_TTL_MS);
    return searchResponse;
  }
}

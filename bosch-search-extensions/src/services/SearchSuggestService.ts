import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Autocomplete service using two strategies:
 *  1. /_api/search/suggest  — historical popular queries (requires tenant query analytics)
 *  2. /_api/search/query    — wildcard prefix search on content titles (always works)
 *
 * Uses native fetch() with credentials:'same-origin' to avoid SPHttpClient
 * adding a conflicting Accept header on top of our own.
 */
export class SearchSuggestService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getSuggestions(query: string, count: number = 5): Promise<string[]> {
    if (!query || query.trim().length < 2) return [];
    const trimmed = query.trim();
    try {
      const fromSuggest = await this.fetchFromSuggestApi(trimmed, count);
      if (fromSuggest.length > 0) return fromSuggest;
      return await this.fetchFromWildcardSearch(trimmed, count);
    } catch {
      return [];
    }
  }

  /**
   * Strategy 1: /_api/search/suggest
   * Returns popular query strings that match the prefix.
   * NOTE: Requires tenant-level query analytics data; may return empty on low-activity tenants.
   */
  private async fetchFromSuggestApi(query: string, count: number): Promise<string[]> {
    try {
      const encodedQuery = `'${encodeURIComponent(query)}'`;
      const url =
        `${this.context.pageContext.web.absoluteUrl}/_api/search/suggest` +
        `?querytext=${encodedQuery}` +
        `&numberOfQuerySuggestions=${count}` +
        `&numberOfPersonalResults=0`;

      // Use native fetch so we have sole control over the Accept header.
      // SPHttpClient.configurations.v1 injects its own Accept header which
      // conflicts with ours and causes SharePoint to return 'ACCEPT invalid'.
      const response = await fetch(url, {
        method: 'GET',
        credentials: 'same-origin',
        headers: { 'Accept': 'application/json;odata=verbose' }
      });
      if (!response.ok) return [];

      const data = await response.json() as Record<string, unknown>;

      // Verbose: { d: { suggest: { Queries: { results: [{ Query: string }] } } } }
      const d = data?.d as Record<string, unknown> | undefined;
      const suggest = d?.suggest as Record<string, unknown> | undefined;
      const queriesObj = suggest?.Queries as Record<string, unknown> | undefined;
      const results = queriesObj?.results;

      if (!Array.isArray(results)) return [];

      return (results as Array<Record<string, unknown>>)
        .map(item => ((item.Query as string) || '').trim())
        .filter(q => q && q.toLowerCase() !== query.toLowerCase())
        .slice(0, count);
    } catch {
      return [];
    }
  }

  /**
   * Strategy 2: /_api/search/query with wildcard on Title
   * Searches actual content matching the typed prefix — works on any tenant.
   */
  private async fetchFromWildcardSearch(query: string, count: number): Promise<string[]> {
    try {
      const encodedQuery = `'${encodeURIComponent(query)}*'`;
      const url =
        `${this.context.pageContext.web.absoluteUrl}/_api/search/query` +
        `?querytext=${encodedQuery}` +
        `&rowlimit=${count}` +
        `&selectproperties='Title'` +
        `&startrow=0`;

      const response = await fetch(url, {
        method: 'GET',
        credentials: 'same-origin',
        headers: { 'Accept': 'application/json;odata=verbose' }
      });
      if (!response.ok) return [];

      const data = await response.json() as Record<string, unknown>;

      // Verbose: { d: { query: { PrimaryQueryResult: { RelevantResults: { Table: { Rows: { results: [{Cells:{results:[...]}}] } } } } } } }
      const d = data?.d as Record<string, unknown> | undefined;
      const queryObj = d?.query as Record<string, unknown> | undefined;
      const primary = queryObj?.PrimaryQueryResult as Record<string, unknown> | undefined;
      const relevant = primary?.RelevantResults as Record<string, unknown> | undefined;
      const table = relevant?.Table as Record<string, unknown> | undefined;
      const rowsObj = table?.Rows as Record<string, unknown> | undefined;
      const rows = rowsObj?.results;

      if (!Array.isArray(rows)) return [];

      const titles: string[] = [];
      for (const row of rows as Array<Record<string, unknown>>) {
        const cellsObj = row?.Cells as Record<string, unknown> | undefined;
        const cells = cellsObj?.results;
        if (!Array.isArray(cells)) continue;
        const titleCell = (cells as Array<Record<string, unknown>>).find(c => c.Key === 'Title');
        const title = ((titleCell?.Value as string) || '').trim();
        if (title && title.toLowerCase() !== query.toLowerCase()) {
          titles.push(title);
        }
      }
      return titles.slice(0, count);
    } catch {
      return [];
    }
  }
}

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

/**
 * Autocomplete / query-suggestion service using the SharePoint Search Suggest API.
 * Endpoint: GET /_api/search/suggest
 *
 * This is the same suggestion engine used by the classic SharePoint search box.
 * It returns queries that other users in the tenant have run successfully,
 * so suggestions improve organically over time.
 *
 * No extra license required — built into SharePoint Online.
 * Uses the SPFx spHttpClient so auth is handled automatically.
 */
export class SearchSuggestService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Returns autocomplete suggestions for the given partial query string.
   * @param query  The text the user has typed so far (min 2 chars).
   * @param count  Max number of suggestions to return (default 5).
   */
  public async getSuggestions(query: string, count: number = 5): Promise<string[]> {
    if (!query || query.trim().length < 2) return [];
    try {
      // The querytext value must be wrapped in single quotes for the REST API
      const encodedQuery = encodeURIComponent(`'${query.trim()}'`);
      const url =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/search/suggest` +
        `?querytext=${encodedQuery}` +
        `&inpreflightmode=1` +
        `&numberOfQuerySuggestions=${count}` +
        `&numberOfPersonalResults=0`;

      const response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!response.ok) return [];

      const data = await response.json() as Record<string, unknown>;

      // With odata=nometadata, the response is: { "Queries": [{ "Query": "..." }, ...] }
      // Fallback verbose format: data.d.suggest.Queries.results (older SP REST style)
      let rawQueries: Array<{ Query: string }> = [];
      const direct = data?.Queries;
      if (Array.isArray(direct)) {
        rawQueries = direct as Array<{ Query: string }>;
      } else {
        // Safely traverse verbose format without complex generics
        const d = data?.d as Record<string, unknown> | undefined;
        const suggest = d?.suggest as Record<string, unknown> | undefined;
        const q = suggest?.Queries as Record<string, unknown> | undefined;
        const results = q?.results;
        if (Array.isArray(results)) rawQueries = results as Array<{ Query: string }>;
      }

      return rawQueries
        .map((q: { Query: string }) => (q.Query || '').trim())
        .filter((q: string) => q && q.toLowerCase() !== query.toLowerCase().trim())
        .slice(0, count);
    } catch {
      // Swallow — suggest failures should never block search
      return [];
    }
  }
}

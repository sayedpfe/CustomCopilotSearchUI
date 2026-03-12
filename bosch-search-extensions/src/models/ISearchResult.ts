export interface ISearchResult {
  id: string;
  title: string;
  summary: string;
  url: string;
  source: string;
  contentSource?: string;
  fileType?: string;
  lastModified?: string;
  author?: string;
  department?: string;
  category?: string;
  hitHighlightedSummary?: string;
}

export interface ISpellingSuggestion {
  originalQuery: string;
  alteredQuery: string;
  /**
   * Suggestion = original query was searched, corrected term is offered as "Did you mean?".
   * Modification = no results for original; the API auto-searched the corrected term.
   */
  type: 'Suggestion' | 'Modification';
}

export interface ISearchResponse {
  results: ISearchResult[];
  totalResults: number;
  moreResultsAvailable: boolean;
  /** Per-source total counts from the API (not page-limited). */
  sourceCounts?: Record<string, number>;
  /** Spelling suggestion / modification returned by the search API. */
  spellingSuggestion?: ISpellingSuggestion;
}

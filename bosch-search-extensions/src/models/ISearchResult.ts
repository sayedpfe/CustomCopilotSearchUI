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

export interface ISearchResponse {
  results: ISearchResult[];
  totalResults: number;
  moreResultsAvailable: boolean;
  /** Per-source total counts from the API (not page-limited). */
  sourceCounts?: Record<string, number>;
}

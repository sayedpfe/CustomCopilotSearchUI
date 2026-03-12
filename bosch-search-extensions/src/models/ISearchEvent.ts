export type SearchEventType = 'Query' | 'Click' | 'ZeroResult';

export interface ISearchEvent {
  query: string;
  eventType: SearchEventType;
  resultCount?: number;
  clickedUrl?: string;
  clickPosition?: number;
  userId?: string;
  timestamp: Date;
  sessionId: string;
  vertical?: string;
}

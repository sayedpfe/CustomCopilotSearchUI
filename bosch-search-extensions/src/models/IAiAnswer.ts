export interface ICitation {
  index: number;
  title: string;
  url: string;
  snippet: string;
}

export interface IAiAnswer {
  answer: string;
  citations: ICitation[];
  isStreaming: boolean;
  error?: string;
  source: 'copilot-chat' | 'copilot-retrieval' | 'graph-search';
}

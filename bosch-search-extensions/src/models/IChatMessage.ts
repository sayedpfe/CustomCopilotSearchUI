export interface IChatCitation {
  title: string;
  url: string;
}

export interface IChatMessage {
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  citations?: IChatCitation[];
  isStreaming?: boolean;
}

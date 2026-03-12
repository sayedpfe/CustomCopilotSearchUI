// SharePoint list names
export const LIST_ANNOUNCEMENTS = 'SearchAnnouncements';
export const LIST_PROMOTED_RESULTS = 'SearchPromotedResults';
export const LIST_ANALYTICS_EVENTS = 'SearchAnalyticsEvents';

// URL query parameter used by PnP Modern Search Box
export const SEARCH_QUERY_PARAM = 'q';

// EventBus event names
export const EVENT_SEARCH_QUERY_CHANGED = 'searchQueryChanged';
export const EVENT_RESULT_CLICKED = 'resultClicked';
export const EVENT_VERTICAL_CHANGED = 'verticalChanged';

// Graph API endpoints
export const GRAPH_SEARCH_ENDPOINT = '/search/query';
export const GRAPH_PEOPLE_ENDPOINT = '/me/people';
export const GRAPH_USERS_ENDPOINT = '/users';

// Copilot API endpoints (Microsoft Graph)
export const COPILOT_RETRIEVAL_ENDPOINT = '/copilot/retrieval';
export const COPILOT_CONVERSATIONS_ENDPOINT = '/copilot/conversations';
export const COPILOT_SEARCH_ENDPOINT = '/copilot/search';

// Copilot Chat API - default grounding
export const COPILOT_CHAT_DEFAULT_SYSTEM_PROMPT = 'You are a helpful Bosch enterprise search assistant. Answer questions based on the organization\'s documents, policies, and knowledge base. Be concise and professional. Cite your sources.';

// M365 Copilot app deep link
export const M365_COPILOT_SEARCH_URL = 'https://m365.cloud.microsoft/chat';

// Microsoft 365 Copilot SKU IDs
export const COPILOT_SKU_IDS = [
  'c815c93d-0759-4bb8-6b4e-5b5d5d93c68e',  // Microsoft 365 Copilot
  '639dec6b-bb19-468b-871c-c5c441c4b0cb',   // Microsoft Copilot Studio
];

// Cache durations (milliseconds)
export const CACHE_PROMOTED_RESULTS_MS = 5 * 60 * 1000;  // 5 minutes
export const CACHE_ANNOUNCEMENTS_MS = 2 * 60 * 1000;     // 2 minutes
export const CACHE_COPILOT_LICENSE_MS = 30 * 60 * 1000;  // 30 minutes

// Analytics batching
export const ANALYTICS_BATCH_SIZE = 10;
export const ANALYTICS_FLUSH_INTERVAL_MS = 5000;  // 5 seconds

// Search query polling interval (for PnP Search Box interop)
export const QUERY_POLL_INTERVAL_MS = 200;

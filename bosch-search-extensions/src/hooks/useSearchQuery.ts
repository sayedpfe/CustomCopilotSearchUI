import { useState, useEffect, useRef } from 'react';
import { EventBus } from '../common/EventBus';
import {
  SEARCH_QUERY_PARAM,
  EVENT_SEARCH_QUERY_CHANGED,
  QUERY_POLL_INTERVAL_MS,
} from '../common/Constants';
import { getQueryParam } from '../common/Utils';

/**
 * React hook that monitors the URL query parameter (?q=) for search query changes.
 * PnP Modern Search Box updates the URL via history.pushState, which does not fire
 * popstate. This hook polls the URL every 200ms to detect changes and emits
 * 'searchQueryChanged' on the EventBus.
 */
export function useSearchQuery(): string {
  const [query, setQuery] = useState<string>(() => getQueryParam(SEARCH_QUERY_PARAM));
  const lastQueryRef = useRef<string>(query);

  useEffect(() => {
    // Poll the URL for changes from PnP Search Box
    const intervalId = setInterval(() => {
      const currentQuery = getQueryParam(SEARCH_QUERY_PARAM);
      if (currentQuery !== lastQueryRef.current) {
        lastQueryRef.current = currentQuery;
        setQuery(currentQuery);
        EventBus.emit(EVENT_SEARCH_QUERY_CHANGED, { query: currentQuery });
      }
    }, QUERY_POLL_INTERVAL_MS);

    // Also listen for EventBus events from custom search boxes
    const handleQueryChange = (data: unknown): void => {
      const eventData = data as { query: string };
      if (eventData.query !== lastQueryRef.current) {
        lastQueryRef.current = eventData.query;
        setQuery(eventData.query);
      }
    };

    EventBus.on(EVENT_SEARCH_QUERY_CHANGED, handleQueryChange);

    return () => {
      clearInterval(intervalId);
      EventBus.off(EVENT_SEARCH_QUERY_CHANGED, handleQueryChange);
    };
  }, []);

  return query;
}

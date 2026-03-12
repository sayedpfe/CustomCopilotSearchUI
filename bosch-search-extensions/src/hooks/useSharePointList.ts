import { useState, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointListService } from '../services/SharePointListService';

/**
 * React hook for fetching items from a SharePoint list.
 */
export function useSharePointList<T>(
  context: WebPartContext,
  listTitle: string,
  filter?: string,
  select?: string[],
  orderBy?: string,
  top?: number
): {
  items: T[];
  loading: boolean;
  error: string | undefined;
  refresh: () => void;
} {
  const [items, setItems] = useState<T[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [refreshCounter, setRefreshCounter] = useState<number>(0);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(undefined);

    const service = new SharePointListService(context);
    service
      .getListItems<T>(listTitle, filter, select, orderBy, top)
      .then((data) => {
        if (!cancelled) {
          setItems(data);
          setLoading(false);
        }
      })
      .catch((err: Error) => {
        if (!cancelled) {
          console.error(`[useSharePointList] Error fetching "${listTitle}":`, err);
          setError(err.message);
          setLoading(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [context, listTitle, filter, orderBy, top, refreshCounter]);

  const refresh = (): void => {
    setRefreshCounter((c) => c + 1);
  };

  return { items, loading, error, refresh };
}

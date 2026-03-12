import { useState, useEffect } from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * React hook that acquires an MSGraphClientV3 instance from the SPFx context.
 */
export function useGraphClient(context: WebPartContext): {
  graphClient: MSGraphClientV3 | undefined;
  error: string | undefined;
} {
  const [graphClient, setGraphClient] = useState<MSGraphClientV3 | undefined>(undefined);
  const [error, setError] = useState<string | undefined>(undefined);

  useEffect(() => {
    let cancelled = false;

    context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        if (!cancelled) {
          setGraphClient(client);
        }
      })
      .catch((err: Error) => {
        if (!cancelled) {
          console.error('[useGraphClient] Failed to get Graph client:', err);
          setError(err.message);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [context]);

  return { graphClient, error };
}

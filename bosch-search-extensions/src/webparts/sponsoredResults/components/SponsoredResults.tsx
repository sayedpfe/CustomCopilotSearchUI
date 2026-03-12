import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointListService } from '../../../services/SharePointListService';
import { IPromotedResult } from '../../../models';
import { useSearchQuery } from '../../../hooks/useSearchQuery';
import { SimpleCache } from '../../../common/Utils';
import { CACHE_PROMOTED_RESULTS_MS } from '../../../common/Constants';
import styles from './SponsoredResults.module.scss';

export interface ISponsoredResultsProps {
  context: WebPartContext;
  listName: string;
}

const promotedCache = new SimpleCache<IPromotedResult[]>();
const CACHE_KEY = 'all_promoted';

function matchesQuery(keywords: string, query: string): boolean {
  if (!keywords || !query) return false;
  const queryLower = query.toLowerCase().trim();
  const keywordList = keywords.toLowerCase().split(',').map((k) => k.trim()).filter(Boolean);
  return keywordList.some((keyword) => {
    // Check if keyword is a word boundary match in the query
    const regex = new RegExp(`\\b${keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
    return regex.test(queryLower);
  });
}

export const SponsoredResults: React.FC<ISponsoredResultsProps> = ({ context, listName }) => {
  const searchQuery = useSearchQuery();
  const [allPromoted, setAllPromoted] = useState<IPromotedResult[]>([]);
  const serviceRef = useRef(new SharePointListService(context));

  useEffect(() => {
    const cached = promotedCache.get(CACHE_KEY);
    if (cached) {
      setAllPromoted(cached);
      return;
    }

    const now = new Date().toISOString();
    serviceRef.current
      .getListItems<IPromotedResult>(
        listName,
        `IsActive eq 1`,
        ['Id', 'Title', 'Description', 'Url', 'Keywords', 'IconUrl', 'IsActive', 'StartDate', 'EndDate', 'SortOrder'],
        'SortOrder asc',
        50
      )
      .then((items) => {
        // Client-side date filtering for flexibility
        const activeItems = items.filter((item) => {
          if (item.startDate && new Date(item.startDate) > new Date(now)) return false;
          if (item.endDate && new Date(item.endDate) < new Date(now)) return false;
          return true;
        });
        promotedCache.set(CACHE_KEY, activeItems, CACHE_PROMOTED_RESULTS_MS);
        setAllPromoted(activeItems);
      })
      .catch((err) => console.error('[SponsoredResults] Error:', err));
  }, [listName]);

  if (!searchQuery) return null;

  const matchingResults = allPromoted.filter((item) => matchesQuery(item.keywords, searchQuery));

  if (matchingResults.length === 0) return null;

  return (
    <div className={styles.sponsoredContainer}>
      {matchingResults.map((result) => (
        <div
          key={result.id}
          className={styles.sponsoredCard}
          onClick={() => window.open(result.url, '_blank', 'noopener,noreferrer')}
          role="link"
          tabIndex={0}
          onKeyDown={(e) => {
            if (e.key === 'Enter') window.open(result.url, '_blank', 'noopener,noreferrer');
          }}
        >
          <div className={styles.cardContent}>
            <span className={styles.badge}>Promoted</span>
            <h3 className={styles.cardTitle}>{result.title}</h3>
            <p className={styles.cardDescription}>{result.description}</p>
            <span className={styles.cardUrl}>{result.url}</span>
          </div>
        </div>
      ))}
    </div>
  );
};

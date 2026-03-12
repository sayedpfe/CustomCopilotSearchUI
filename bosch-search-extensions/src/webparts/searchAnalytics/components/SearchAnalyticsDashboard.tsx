import * as React from 'react';
import { useState, useEffect } from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointListService } from '../../../services/SharePointListService';
import styles from './SearchAnalyticsDashboard.module.scss';

export interface ISearchAnalyticsDashboardProps {
  context: WebPartContext;
  listName: string;
  defaultDays: number;
  maxRows: number;
}

interface IAnalyticsItem {
  Title: string;
  EventType: string;
  ResultCount: number;
  ClickedUrl?: { Url: string };
  ClickPosition: number;
  Timestamp: string;
  Vertical: string;
}

interface IQueryStats {
  query: string;
  count: number;
}

const DATE_OPTIONS: IDropdownOption[] = [
  { key: 7, text: 'Last 7 days' },
  { key: 30, text: 'Last 30 days' },
  { key: 90, text: 'Last 90 days' },
];

export const SearchAnalyticsDashboard: React.FC<ISearchAnalyticsDashboardProps> = ({
  context,
  listName,
  defaultDays,
  maxRows,
}) => {
  const [days, setDays] = useState<number>(defaultDays);
  const [loading, setLoading] = useState(true);
  const [totalQueries, setTotalQueries] = useState(0);
  const [totalClicks, setTotalClicks] = useState(0);
  const [zeroResultCount, setZeroResultCount] = useState(0);
  const [topQueries, setTopQueries] = useState<IQueryStats[]>([]);
  const [zeroResultQueries, setZeroResultQueries] = useState<IQueryStats[]>([]);

  useEffect(() => {
    setLoading(true);

    const service = new SharePointListService(context);
    const sinceDate = new Date();
    sinceDate.setDate(sinceDate.getDate() - days);
    const filter = `Timestamp ge datetime'${sinceDate.toISOString()}'`;

    service
      .getListItems<IAnalyticsItem>(
        listName,
        filter,
        ['Title', 'EventType', 'ResultCount', 'ClickPosition', 'Timestamp'],
        'Timestamp desc',
        5000
      )
      .then((items) => {
        // Aggregate stats
        const queries = items.filter((i) => i.EventType === 'Query' || i.EventType === 'ZeroResult');
        const clicks = items.filter((i) => i.EventType === 'Click');
        const zeroResults = items.filter((i) => i.EventType === 'ZeroResult');

        setTotalQueries(queries.length);
        setTotalClicks(clicks.length);
        setZeroResultCount(zeroResults.length);

        // Top queries by frequency
        const queryFreq = new Map<string, number>();
        queries.forEach((q) => {
          const key = q.Title.toLowerCase().trim();
          queryFreq.set(key, (queryFreq.get(key) || 0) + 1);
        });
        const sortedQueries = Array.from(queryFreq.entries())
          .map(([query, count]) => ({ query, count }))
          .sort((a, b) => b.count - a.count)
          .slice(0, maxRows);
        setTopQueries(sortedQueries);

        // Zero result queries
        const zeroFreq = new Map<string, number>();
        zeroResults.forEach((q) => {
          const key = q.Title.toLowerCase().trim();
          zeroFreq.set(key, (zeroFreq.get(key) || 0) + 1);
        });
        const sortedZero = Array.from(zeroFreq.entries())
          .map(([query, count]) => ({ query, count }))
          .sort((a, b) => b.count - a.count)
          .slice(0, maxRows);
        setZeroResultQueries(sortedZero);

        setLoading(false);
      })
      .catch((err) => {
        console.error('[SearchAnalytics] Error:', err);
        setLoading(false);
      });
  }, [context, listName, days, maxRows]);

  const clickThroughRate = totalQueries > 0 ? ((totalClicks / totalQueries) * 100).toFixed(1) : '0';
  const maxQueryCount = topQueries.length > 0 ? topQueries[0].count : 1;

  if (loading) return <Spinner label="Loading analytics..." />;

  return (
    <div className={styles.dashboard}>
      <div className={styles.header}>
        <span className={styles.title}>Search Analytics</span>
        <Dropdown
          options={DATE_OPTIONS}
          selectedKey={days}
          onChange={(_, option) => option && setDays(option.key as number)}
          style={{ width: 180 }}
        />
      </div>

      <div className={styles.statsGrid}>
        <div className={styles.statCard}>
          <span className={styles.statValue}>{totalQueries}</span>
          <span className={styles.statLabel}>Total Searches</span>
        </div>
        <div className={styles.statCard}>
          <span className={styles.statValue}>{clickThroughRate}%</span>
          <span className={styles.statLabel}>Click-Through Rate</span>
        </div>
        <div className={styles.statCard}>
          <span className={styles.statValue}>{zeroResultCount}</span>
          <span className={styles.statLabel}>Zero-Result Searches</span>
        </div>
      </div>

      <div className={styles.section}>
        <div className={styles.sectionTitle}>Top Search Queries</div>
        <table className={styles.table}>
          <thead>
            <tr>
              <th style={{ width: '50%' }}>Query</th>
              <th style={{ width: '15%' }}>Count</th>
              <th style={{ width: '35%' }}>Frequency</th>
            </tr>
          </thead>
          <tbody>
            {topQueries.map((q, i) => (
              <tr key={i}>
                <td>{q.query}</td>
                <td>{q.count}</td>
                <td>
                  <div className={styles.barContainer}>
                    <div
                      className={styles.bar}
                      style={{ width: `${(q.count / maxQueryCount) * 100}%` }}
                    />
                  </div>
                </td>
              </tr>
            ))}
            {topQueries.length === 0 && (
              <tr><td colSpan={3} style={{ textAlign: 'center', fontStyle: 'italic' }}>No search data yet</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {zeroResultQueries.length > 0 && (
        <div className={styles.section}>
          <div className={styles.sectionTitle}>Zero-Result Queries</div>
          <table className={styles.table}>
            <thead>
              <tr>
                <th style={{ width: '70%' }}>Query</th>
                <th style={{ width: '30%' }}>Count</th>
              </tr>
            </thead>
            <tbody>
              {zeroResultQueries.map((q, i) => (
                <tr key={i}>
                  <td className={styles.zeroResultHighlight}>{q.query}</td>
                  <td>{q.count}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

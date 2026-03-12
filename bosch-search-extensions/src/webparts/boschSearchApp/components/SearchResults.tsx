import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { Link } from '@fluentui/react/lib/Link';
import { CopilotChatService } from '../../../services/CopilotChatService';
import { CopilotRetrievalService } from '../../../services/CopilotRetrievalService';
import { CopilotSearchService } from '../../../services/CopilotSearchService';
import { GraphSearchService } from '../../../services/GraphSearchService';
import { SharePointListService } from '../../../services/SharePointListService';
import { ISearchResult, IPromotedResult, ICitation, ISpellingSuggestion } from '../../../models';
import { M365_COPILOT_SEARCH_URL } from '../../../common/Constants';
import { markdownToHtml, stripHtmlTags } from '../../../common/Utils';
import { SearchBox } from './SearchBox';
import { NewsCarousel } from './NewsCarousel';
import styles from './BoschSearchApp.module.scss';

type SearchScope = 'work' | 'web';
type SearchVertical = 'all' | 'copilot' | 'images' | 'videos' | 'more';

export interface ISearchResultsProps {
  query: string;
  searchScope: SearchScope;
  onSearchScopeChange: (scope: SearchScope) => void;
  onSearch: (query: string) => void;
  graphClient: MSGraphClientV3;
  hasCopilot: boolean | null;
  activeVertical: SearchVertical;
  groundingMode: 'work' | 'web' | 'both';
  maxRetrievalResults: number;
  showCopilotLink: boolean;
  promotedResultsListName: string;
  newsSourceSiteUrl: string;
  context: WebPartContext;
}

export const SearchResults: React.FC<ISearchResultsProps> = ({
  query,
  searchScope,
  onSearchScopeChange,
  onSearch,
  graphClient,
  hasCopilot,
  groundingMode,
  maxRetrievalResults,
  showCopilotLink,
  promotedResultsListName,
  newsSourceSiteUrl,
  context,
}) => {
  const [aiAnswer, setAiAnswer] = useState('');
  const [aiCitations, setAiCitations] = useState<ICitation[]>([]);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiSource, setAiSource] = useState('');
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [resultsLoading, setResultsLoading] = useState(false);
  const [totalResults, setTotalResults] = useState(0);
  const [promotedResults, setPromotedResults] = useState<IPromotedResult[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [aiExpanded, setAiExpanded] = useState(false);
  const [activeSourceFilter, setActiveSourceFilter] = useState<string>('All');
  const [sourceCountsFromApi, setSourceCountsFromApi] = useState<Record<string, number>>({});
  const [filteredApiResults, setFilteredApiResults] = useState<ISearchResult[]>([]);
  const [filteredApiTotal, setFilteredApiTotal] = useState(0);
  const [filteredPage, setFilteredPage] = useState(1);
  const [filteredLoading, setFilteredLoading] = useState(false);
  const [spellingSuggestion, setSpellingSuggestion] = useState<ISpellingSuggestion | null>(null);
  const abortRef = useRef(false);

  const PAGE_SIZE = 10;

  // Use API-level source counts (total across all pages) when available,
  // fall back to computing from current page results
  const sourceCounts = React.useMemo(() => {
    if (Object.keys(sourceCountsFromApi).length > 0) return sourceCountsFromApi;
    const counts: Record<string, number> = {};
    results.forEach((r) => {
      const src = r.source || 'Other';
      counts[src] = (counts[src] || 0) + 1;
    });
    return counts;
  }, [results, sourceCountsFromApi]);

  // Sort sources: well-known ones first, then alphabetical for connectors
  const sourceFilters = React.useMemo(() => {
    // Only show sources that have at least 1 result
    const sources = Object.keys(sourceCounts).filter((s) => sourceCounts[s] > 0);
    const knownOrder = ['SharePoint', 'Outlook Mail', 'Teams', 'M365 Copilot Chats'];
    const known = knownOrder.filter((s) => sources.includes(s));
    const custom = sources.filter((s) => !knownOrder.includes(s)).sort();
    return known.concat(custom);
  }, [sourceCounts]);

  const displayedResults = React.useMemo(() => {
    if (activeSourceFilter === 'All') return results;
    return filteredApiResults;
  }, [results, activeSourceFilter, filteredApiResults]);

  // Map source names to Fluent icons
  const getSourceIcon = (src: string): string => {
    if (src === 'SharePoint') return 'SharepointLogo';
    if (src === 'Outlook Mail') return 'Mail';
    if (src === 'Teams') return 'TeamsLogo';
    if (src === 'M365 Copilot Chats') return 'Chat';
    if (src === 'All') return 'Search';
    // Graph Connector sources
    return 'Database';
  };

  const totalPages = activeSourceFilter === 'All'
    ? Math.max(1, Math.ceil(totalResults / PAGE_SIZE))
    : Math.max(1, Math.ceil(filteredApiTotal / PAGE_SIZE));

  // Fetch results when query changes
  useEffect(() => {
    if (!query) return;
    abortRef.current = true;
    setCurrentPage(1);
    setActiveSourceFilter('All');
    setFilteredPage(1);
    setFilteredApiResults([]);
    setFilteredApiTotal(0);
    setSpellingSuggestion(null);

    const run = async (): Promise<void> => {
      abortRef.current = false;
      await Promise.all([
        fetchAiAnswer(),
        fetchSearchResults(0),
        fetchPromotedResults(),
      ]);
      // Spell check runs in parallel as a lightweight background call.
      // Uses the Graph Search queryAlterationOptions feature — no extra license needed.
      if (!abortRef.current && graphClient) {
        const graphSearch = new GraphSearchService(graphClient);
        graphSearch.checkSpelling(query)
          .then((suggestion) => { if (!abortRef.current) setSpellingSuggestion(suggestion); })
          .catch(() => { /* ignore — spell check is best-effort */ });
      }
    };
    run();

    return () => { abortRef.current = true; };
  }, [query, searchScope]);

  // Fetch when page changes (but not on initial load)
  useEffect(() => {
    if (!query || currentPage === 1) return;
    fetchSearchResults((currentPage - 1) * PAGE_SIZE);
  }, [currentPage]);

  // Fetch source-filtered results when filter or filter-page changes
  // (query is intentionally omitted: query change resets activeSourceFilter='All')
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!query || activeSourceFilter === 'All') return;
    void fetchSourceResults(activeSourceFilter, (filteredPage - 1) * PAGE_SIZE);
  }, [activeSourceFilter, filteredPage]);

  const fetchAiAnswer = async (): Promise<void> => {
    setAiLoading(true);
    setAiAnswer('');
    setAiCitations([]);
    setAiSource('');

    try {
      if (hasCopilot) {
        const chatService = new CopilotChatService(graphClient);
        const enableWeb = groundingMode === 'web' || groundingMode === 'both';

        // Try streaming first (text appears in ~2-3s), fall back to sync if it fails
        let streamingWorked = false;
        try {
          await chatService.askSingleQuestionStream(
            query,
            {
              onChunk: (textSoFar) => {
                if (abortRef.current) return;
                streamingWorked = true;
                setAiAnswer(textSoFar);
                setAiSource('copilot-chat');
                setAiLoading(false); // Hide spinner as soon as first chunk arrives
              },
              onDone: (fullText, attributions) => {
                if (abortRef.current) return;
                setAiAnswer(fullText);
                setAiCitations(
                  attributions
                    .filter((a) => a.url)
                    .map((attr, i) => ({
                      index: i + 1,
                      title: attr.title || 'Source',
                      url: attr.url || '',
                      snippet: '',
                    }))
                );
              },
              onError: (err) => {
                if (!abortRef.current) {
                  console.error('[SearchResults] Copilot stream error:', err);
                }
              },
            },
            enableWeb
          );
        } catch (streamErr) {
          console.warn('[SearchResults] Streaming threw:', streamErr);
        }

        // If streaming failed (no chunks arrived), fall back to sync chat API
        if (!streamingWorked && !abortRef.current) {
          console.log('[SearchResults] Streaming did not produce chunks, falling back to sync chat API');
          try {
            const result = await chatService.askSingleQuestion(query, enableWeb);
            if (!abortRef.current) {
              setAiAnswer(result.responseText);
              setAiSource('copilot-chat');
              setAiCitations(
                result.attributions
                  .filter((a) => a.url)
                  .map((attr, i) => ({
                    index: i + 1,
                    title: attr.title || 'Source',
                    url: attr.url || '',
                    snippet: '',
                  }))
              );
            }
          } catch (syncErr) {
            console.error('[SearchResults] Sync chat also failed:', syncErr);
            if (!abortRef.current) {
              setAiAnswer('Unable to get Copilot response. Please try again.');
              setAiSource('copilot-chat');
            }
          }
        }
        return;
      } else {
        // Try Retrieval API, fall back to Graph Search
        try {
          const retrievalService = new CopilotRetrievalService(graphClient);
          const hits = await retrievalService.retrieveFromAllSources(query, {
            maximumNumberOfResults: maxRetrievalResults,
            resourceMetadata: ['title', 'author'],
          });
          if (hits.length > 0 && !abortRef.current) {
            const { contextText, citations } = CopilotRetrievalService.formatAsContext(hits);
            setAiAnswer(contextText);
            setAiCitations(citations);
            setAiSource('copilot-retrieval');
            return;
          }
        } catch {
          // Retrieval not available
        }

        if (abortRef.current) return;
        const searchService = new GraphSearchService(graphClient);
        const response = await searchService.search(query, ['driveItem', 'listItem', 'externalItem'], 0, 5);
        if (abortRef.current) return;

        if (response.results.length > 0) {
          const text = response.results
            .slice(0, 3)
            .map((r, i) => `[${i + 1}] ${r.title}\n${r.summary || 'No summary available'}`)
            .join('\n\n');
          setAiAnswer(text);
          setAiSource('graph-search');
          setAiCitations(
            response.results.slice(0, 3).map((r, i) => ({
              index: i + 1,
              title: r.title,
              url: r.url,
              snippet: r.summary?.substring(0, 150) || '',
            }))
          );
        }
      }
    } catch (err) {
      if (!abortRef.current) {
        console.error('[SearchResults] AI answer error:', err);
      }
    } finally {
      if (!abortRef.current) setAiLoading(false);
    }
  };

  const fetchSourceResults = async (source: string, from: number): Promise<void> => {
    setFilteredLoading(true);
    try {
      const graphSearch = new GraphSearchService(graphClient);
      const response = await graphSearch.searchBySource(query, source, from, PAGE_SIZE);
      if (abortRef.current) return;
      setFilteredApiResults(response.results);
      setFilteredApiTotal(response.totalResults);
    } catch (err) {
      if (!abortRef.current) console.error('[SearchResults] Source filter search error:', err);
    } finally {
      if (!abortRef.current) setFilteredLoading(false);
    }
  };

  const fetchSearchResults = async (from: number = 0): Promise<void> => {
    setResultsLoading(true);
    try {
      if (hasCopilot) {
        // Copilot-licensed users: use the Copilot Search API
        // Same API that powers M365 Copilot Search (all sources, ranked results)
        console.log('[SearchResults] Using Copilot Search API for search results');
        try {
          const copilotSearch = new CopilotSearchService(graphClient);
          const response = await copilotSearch.search(query, from, PAGE_SIZE);
          if (abortRef.current) return;
          setResults(response.results);
          setTotalResults(response.totalResults);
          // Only update sidebar source counts on initial load (not on page navigation)
          if (from === 0 && response.sourceCounts) setSourceCountsFromApi(response.sourceCounts);
          return;
        } catch (copilotErr) {
          console.warn('[SearchResults] Copilot Search API failed, falling back to Graph Search:', copilotErr);
        }
      }

      // Non-Copilot users (or fallback): use Graph Search API with all entity types
      console.log('[SearchResults] Using Graph Search API for search results');
      const searchService = new GraphSearchService(graphClient);
      const response = await searchService.searchAll(query, from, PAGE_SIZE);
      if (abortRef.current) return;
      setResults(response.results);
      setTotalResults(response.totalResults);
      // Only update sidebar source counts on initial load (not on page navigation)
      if (from === 0 && response.sourceCounts) setSourceCountsFromApi(response.sourceCounts);
    } catch (err) {
      if (!abortRef.current) {
        console.error('[SearchResults] Search error:', err);
      }
    } finally {
      if (!abortRef.current) setResultsLoading(false);
    }
  };

  const fetchPromotedResults = async (): Promise<void> => {
    try {
      const listService = new SharePointListService(context);
      const items = await listService.getListItems<IPromotedResult>(
        promotedResultsListName,
        'IsActive eq 1',
        ['Id', 'Title', 'Description', 'Url', 'Keywords', 'IconUrl', 'IsActive', 'SortOrder'],
        'SortOrder asc',
        50
      );

      const queryLower = query.toLowerCase();
      const matched = items.filter((item) => {
        const keywords = (item.keywords || '').toLowerCase();
        return keywords.split(',').some(
          (kw: string) => queryLower.includes(kw.trim()) || kw.trim().includes(queryLower)
        );
      });

      if (!abortRef.current) setPromotedResults(matched);
    } catch {
      // List may not exist yet
    }
  };

  return (
    <div className={styles.resultsContainer}>
      <div className={styles.resultsHeader}>
        <SearchBox
          initialQuery={query}
          searchScope={searchScope}
          onSearchScopeChange={onSearchScopeChange}
          onSearch={onSearch}
          variant="results"
          hasCopilot={hasCopilot}
          context={context}
        />
      </div>

      <div className={styles.resultsBody}>
        <div className={styles.resultsMain}>
          {/* Promoted Results */}
          {promotedResults.length > 0 && (
            <div className={styles.promotedSection}>
              {promotedResults.map((pr) => (
                <a
                  key={pr.id}
                  href={pr.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className={styles.promotedCard}
                >
                  <div className={styles.promotedTitle}>
                    <Icon iconName="Pinned" className={styles.promotedIcon} />
                    {pr.title}
                  </div>
                  <div className={styles.promotedDescription}>{pr.description}</div>
                  <div className={styles.promotedUrl}>{pr.url}</div>
                </a>
              ))}
            </div>
          )}

          {/* Spell correction banner */}
          {spellingSuggestion && !resultsLoading && (
            <div className={styles.spellBanner}>
              {spellingSuggestion.type === 'Modification' ? (
                <>
                  Showing results for <strong>{spellingSuggestion.alteredQuery}</strong>.{' '}
                  Search instead for{' '}
                  <button
                    className={styles.spellBannerLink}
                    onClick={() => onSearch(spellingSuggestion.originalQuery)}
                  >
                    {spellingSuggestion.originalQuery}
                  </button>
                </>
              ) : (
                <>
                  Did you mean{' '}
                  <button
                    className={styles.spellBannerLink}
                    onClick={() => onSearch(spellingSuggestion.alteredQuery)}
                  >
                    {spellingSuggestion.alteredQuery}
                  </button>
                  ?
                </>
              )}
            </div>
          )}

          {/* AI Answer — collapsible */}
          <div className={styles.aiAnswerSection}>
            <div
              className={styles.aiAnswerHeader}
              onClick={() => !aiLoading && aiAnswer && setAiExpanded((v) => !v)}
              style={{ cursor: aiAnswer && !aiLoading ? 'pointer' : 'default' }}
            >
              <Icon iconName="Robot" className={styles.aiAnswerIcon} />
              <span className={styles.aiAnswerTitle}>
                {hasCopilot ? 'Copilot' : 'Search Summary'}
              </span>
              {aiSource && (
                <span
                  className={styles.aiSourceBadge}
                  style={{
                    background: hasCopilot ? '#6264a7' : '#107c10',
                  }}
                >
                  {aiSource === 'copilot-chat'
                    ? 'Copilot'
                    : aiSource === 'copilot-retrieval'
                    ? 'Retrieval API'
                    : 'Graph Search'}
                </span>
              )}
              {aiAnswer && !aiLoading && (
                <span className={styles.aiAnswerToggle}>
                  <Icon iconName={aiExpanded ? 'ChevronUp' : 'ChevronDown'} />
                  {aiExpanded ? 'Show less' : 'Continue reading'}
                </span>
              )}
            </div>

            {aiLoading && (
              <Spinner label={hasCopilot ? 'Asking Copilot...' : 'Searching...'} />
            )}

            {aiAnswer && !aiLoading && (
              <>
                <div className={styles.aiAnswerDisclaimer}>
                  AI-generated content may be incorrect
                </div>
                <div
                  className={`${styles.aiAnswerText} ${!aiExpanded ? styles.aiAnswerCollapsed : ''}`}
                  dangerouslySetInnerHTML={{ __html: markdownToHtml(aiAnswer) }}
                />
                {aiExpanded && aiCitations.length > 0 && (
                  <div className={styles.aiCitations}>
                    {aiCitations.map((c) => (
                      <a
                        key={c.index}
                        href={c.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        className={styles.aiCitationLink}
                      >
                        [{c.index}] {c.title}
                      </a>
                    ))}
                  </div>
                )}
                {aiExpanded && showCopilotLink && hasCopilot && (
                  <div className={styles.copilotDeepLink}>
                    <Link
                      href={`${M365_COPILOT_SEARCH_URL}?q=${encodeURIComponent(query)}`}
                      target="_blank"
                    >
                      <Icon iconName="OpenInNewWindow" style={{ marginRight: 4 }} />
                      Open in M365 Copilot
                    </Link>
                  </div>
                )}
              </>
            )}
          </div>

          {/* Search Results + Sidebar layout */}
          {(resultsLoading || filteredLoading || results.length > 0) && (
            <div className={styles.resultsWithSidebar}>
              {/* Results list — left side */}
              <div className={styles.resultsList}>
                {(resultsLoading || filteredLoading) && <Spinner label="Loading results..." />}

                {!(resultsLoading || filteredLoading) && displayedResults.length > 0 && (
                  <>
                    <div className={styles.resultsCount}>
                      {activeSourceFilter === 'All'
                        ? `About ${totalResults.toLocaleString()} results`
                        : `${(filteredApiTotal || sourceCounts[activeSourceFilter] || 0).toLocaleString()} ${activeSourceFilter} results`}
                    </div>
                    {displayedResults.map((result, i) => (
                      <div key={i} className={styles.resultItem}>
                        <div className={styles.resultMeta}>
                          <span className={styles.resultSourceBadge}>
                            <Icon iconName={getSourceIcon(result.source)} className={styles.resultSourceIcon} />
                            {result.source}
                          </span>
                          {result.lastModified && (
                            <span className={styles.resultDate}>
                              {new Date(result.lastModified).toLocaleDateString()}
                            </span>
                          )}
                        </div>
                        <a
                          href={result.url}
                          target="_blank"
                          rel="noopener noreferrer"
                          className={styles.resultTitle}
                        >
                          {result.title}
                        </a>
                        <div className={styles.resultUrl}>{result.url}</div>
                        {result.summary && (
                          <div className={styles.resultSummary}>{stripHtmlTags(result.summary)}</div>
                        )}
                      </div>
                    ))}
                  </>
                )}
              </div>

              {/* Source filter sidebar — right side, persists across page changes */}
              {Object.keys(sourceCounts).length > 0 && (
                <div className={styles.sourceFilterSidebar}>
                  <div className={styles.sourceFilterHeader}>
                    <Icon iconName="Filter" className={styles.sourceFilterHeaderIcon} />
                    <span>Results</span>
                  </div>
                  <button
                    className={`${styles.sourceFilterItem} ${activeSourceFilter === 'All' ? styles.sourceFilterItemActive : ''}`}
                    onClick={() => { setActiveSourceFilter('All'); setFilteredPage(1); setFilteredApiResults([]); setFilteredApiTotal(0); }}
                  >
                    <Icon iconName="Search" className={styles.sourceFilterItemIcon} />
                    <span className={styles.sourceFilterItemLabel}>All Results</span>
                    <span className={styles.sourceFilterItemCount}>{totalResults.toLocaleString()}</span>
                  </button>
                  {sourceFilters.map((src) => (
                    <button
                      key={src}
                      className={`${styles.sourceFilterItem} ${activeSourceFilter === src ? styles.sourceFilterItemActive : ''}`}
                      onClick={() => { setActiveSourceFilter(src); setFilteredPage(1); setFilteredApiResults([]); setFilteredApiTotal(0); }}
                    >
                      <Icon iconName={getSourceIcon(src)} className={styles.sourceFilterItemIcon} />
                      <span className={styles.sourceFilterItemLabel}>{src}</span>
                      <span className={styles.sourceFilterItemCount}>{(sourceCounts[src] || 0).toLocaleString()}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}

          {!(resultsLoading || filteredLoading) && displayedResults.length === 0 && !aiLoading && (
            <div className={styles.noResults}>
              No results found for &quot;{query}&quot;. Try different keywords.
            </div>
          )}

          {/* Pagination */}
          {!(resultsLoading || filteredLoading) && totalPages > 1 && (
            <div className={styles.pagination}>
              <button
                className={styles.paginationButton}
                disabled={(activeSourceFilter === 'All' ? currentPage : filteredPage) <= 1}
                onClick={() => activeSourceFilter === 'All'
                  ? setCurrentPage((p) => Math.max(1, p - 1))
                  : setFilteredPage((p) => Math.max(1, p - 1))}
              >
                <Icon iconName="ChevronLeft" /> Previous
              </button>
              <span className={styles.paginationInfo}>
                Page {activeSourceFilter === 'All' ? currentPage : filteredPage} of {totalPages}
              </span>
              <button
                className={styles.paginationButton}
                disabled={(activeSourceFilter === 'All' ? currentPage : filteredPage) >= totalPages}
                onClick={() => activeSourceFilter === 'All'
                  ? setCurrentPage((p) => Math.min(totalPages, p + 1))
                  : setFilteredPage((p) => Math.min(totalPages, p + 1))}
              >
                Next <Icon iconName="ChevronRight" />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* News carousel below results */}
      {graphClient && (
        <NewsCarousel graphClient={graphClient} siteUrl={newsSourceSiteUrl} />
      )}
    </div>
  );
};

import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useSearchQuery } from '../../../hooks/useSearchQuery';
import { useGraphClient } from '../../../hooks/useGraphClient';
import { CopilotChatService } from '../../../services/CopilotChatService';
import { CopilotRetrievalService } from '../../../services/CopilotRetrievalService';
import { CopilotDetectionService } from '../../../services/CopilotDetectionService';
import { GraphSearchService } from '../../../services/GraphSearchService';
import { ICitation } from '../../../models';
import { M365_COPILOT_SEARCH_URL } from '../../../common/Constants';
import styles from './AiAnswerPanel.module.scss';

export interface IAiAnswerPanelProps {
  context: WebPartContext;
  groundingMode: 'work' | 'web' | 'both';
  maxRetrievalResults: number;
  showCopilotLink: boolean;
}

export const AiAnswerPanel: React.FC<IAiAnswerPanelProps> = ({
  context,
  groundingMode,
  maxRetrievalResults,
  showCopilotLink,
}) => {
  const searchQuery = useSearchQuery();
  const { graphClient } = useGraphClient(context);
  const [answer, setAnswer] = useState<string>('');
  const [citations, setCitations] = useState<ICitation[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>('');
  const [hasCopilot, setHasCopilot] = useState<boolean | null>(null);
  const [source, setSource] = useState<string>('');
  const abortRef = useRef<boolean>(false);

  // Check Copilot license on mount
  useEffect(() => {
    if (!graphClient) return;
    const detector = new CopilotDetectionService(graphClient);
    detector.hasCopilotLicense().then(setHasCopilot);
  }, [graphClient]);

  // Generate answer when query changes
  useEffect(() => {
    if (!searchQuery || !graphClient || hasCopilot === null) {
      setAnswer('');
      setCitations([]);
      return;
    }

    abortRef.current = true;

    const generateAnswer = async (): Promise<void> => {
      abortRef.current = false;
      setLoading(true);
      setAnswer('');
      setError('');
      setCitations([]);

      try {
        if (hasCopilot) {
          // === COPILOT PATH: Chat API for AI-synthesized summary ===
          await generateCopilotAnswer(searchQuery);
        } else {
          // === NON-COPILOT PATH: Retrieval API (pay-as-you-go) or Graph Search ===
          await generateGraphSearchAnswer(searchQuery);
        }
      } catch (err) {
        if (!abortRef.current) {
          console.error('[AiAnswerPanel] Error:', err);
          setError(err instanceof Error ? err.message : 'Failed to generate answer');
        }
      } finally {
        if (!abortRef.current) {
          setLoading(false);
        }
      }
    };

    generateAnswer();

    return () => {
      abortRef.current = true;
    };
  }, [searchQuery, graphClient, hasCopilot]);

  /**
   * Copilot-licensed user path:
   * Uses Copilot Chat API — handles grounding internally.
   * Returns a fully synthesized answer with attributions.
   */
  const generateCopilotAnswer = async (query: string): Promise<void> => {
    const chatService = new CopilotChatService(graphClient!);
    const enableWebGrounding = groundingMode === 'web' || groundingMode === 'both';

    const result = await chatService.askSingleQuestion(query, enableWebGrounding);
    if (abortRef.current) return;

    setAnswer(result.responseText);
    setSource('copilot-chat');
    setCitations(
      result.attributions
        .filter((a) => a.url)
        .map((attr, i) => ({
          index: i + 1,
          title: attr.title || 'Source',
          url: attr.url || '',
          snippet: '',
        }))
    );
  };

  /**
   * Non-Copilot user path:
   * 1. Try Copilot Retrieval API (works with pay-as-you-go for SharePoint/Connectors)
   * 2. Fall back to Graph Search API if Retrieval is unavailable
   * 3. Displays relevant text chunks (no LLM synthesis without Copilot)
   */
  const generateGraphSearchAnswer = async (query: string): Promise<void> => {
    try {
      // Try Retrieval API first (may work with pay-as-you-go)
      const retrievalService = new CopilotRetrievalService(graphClient!);
      const hits = await retrievalService.retrieveFromAllSources(query, {
        maximumNumberOfResults: maxRetrievalResults,
        resourceMetadata: ['title', 'author'],
      });

      if (hits.length > 0 && !abortRef.current) {
        const { contextText, citations: citationList } = CopilotRetrievalService.formatAsContext(hits);
        setAnswer(contextText);
        setCitations(citationList);
        setSource('copilot-retrieval');
        return;
      }
    } catch {
      // Retrieval API not available, fall back to Graph Search
    }

    // Fallback: Graph Search API (no Copilot license needed)
    if (abortRef.current) return;
    const searchService = new GraphSearchService(graphClient!);
    const response = await searchService.search(query, ['driveItem', 'listItem', 'externalItem'], 0, 10);
    if (abortRef.current) return;

    if (response.results.length === 0) {
      setAnswer('No results found for your query.');
      setSource('graph-search');
      return;
    }

    const citationList: ICitation[] = response.results.map((r, i) => ({
      index: i + 1,
      title: r.title,
      url: r.url,
      snippet: r.summary?.substring(0, 200) || '',
    }));

    const summaryText = response.results
      .slice(0, 5)
      .map((r, i) => `[${i + 1}] ${r.title}\n${r.summary || 'No summary available'}`)
      .join('\n\n');

    setAnswer(summaryText);
    setCitations(citationList);
    setSource('graph-search');
  };

  if (!searchQuery) return null;

  return (
    <div className={styles.aiPanel}>
      <div className={styles.header}>
        <Icon iconName="Robot" className={styles.aiIcon} />
        <span className={styles.headerTitle}>
          {hasCopilot ? 'Copilot Summary' : 'Search Summary'}
        </span>
        {hasCopilot && <span className={styles.copilotBadge}>Copilot</span>}
        {!hasCopilot && source && (
          <span className={styles.copilotBadge} style={{ background: '#107c10' }}>
            {source === 'copilot-retrieval' ? 'Retrieval API' : 'Graph Search'}
          </span>
        )}
      </div>

      {error && <div className={styles.errorMessage}>{error}</div>}

      {loading && (
        <Spinner label={hasCopilot ? 'Asking Copilot...' : 'Searching...'} />
      )}

      {answer && !loading && (
        <div className={styles.answerText}>{answer}</div>
      )}

      {citations.length > 0 && !loading && (
        <div className={styles.citationsList}>
          <div className={styles.citationsTitle}>Sources</div>
          {citations.map((citation) => (
            <div key={citation.index}>
              <div className={styles.citation}>
                <span className={styles.citationIndex}>[{citation.index}]</span>
                <a
                  href={citation.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className={styles.citationLink}
                >
                  {citation.title}
                </a>
              </div>
              {citation.snippet && (
                <div className={styles.citationSnippet}>{citation.snippet}</div>
              )}
            </div>
          ))}
        </div>
      )}

      {showCopilotLink && hasCopilot && !loading && answer && (
        <div style={{ marginTop: 12, textAlign: 'right' }}>
          <Link href={`${M365_COPILOT_SEARCH_URL}?q=${encodeURIComponent(searchQuery)}`} target="_blank">
            <Icon iconName="OpenInNewWindow" style={{ marginRight: 4 }} />
            Open in M365 Copilot
          </Link>
        </div>
      )}
    </div>
  );
};

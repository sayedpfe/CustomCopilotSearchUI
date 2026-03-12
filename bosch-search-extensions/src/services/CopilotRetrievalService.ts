import { MSGraphClientV3 } from '@microsoft/sp-http';
import { COPILOT_RETRIEVAL_ENDPOINT } from '../common/Constants';

export interface IRetrievalExtract {
  text: string;
  relevanceScore?: number;
}

export interface IRetrievalHit {
  webUrl: string;
  extracts: IRetrievalExtract[];
  resourceType: string;
  resourceMetadata?: Record<string, string>;
  sensitivityLabel?: {
    sensitivityLabelId: string;
    displayName: string;
  };
}

export interface IRetrievalResponse {
  retrievalHits: IRetrievalHit[];
}

export type RetrievalDataSource = 'sharePoint' | 'oneDriveBusiness' | 'externalItem';

/**
 * Service for the Microsoft 365 Copilot Retrieval API.
 * POST /v1.0/copilot/retrieval (or /beta/copilot/retrieval)
 *
 * Returns relevant text chunks from SharePoint, OneDrive, and Copilot Connectors
 * using the same semantic index that powers M365 Copilot.
 *
 * Requires Copilot license (or pay-as-you-go for SharePoint/Connectors).
 * Permissions: Files.Read.All, Sites.Read.All, ExternalItem.Read.All
 */
export class CopilotRetrievalService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Retrieve grounding data from a single data source.
   * Call multiple times (or use batch) for multiple sources.
   */
  public async retrieve(
    queryString: string,
    dataSource: RetrievalDataSource = 'sharePoint',
    options?: {
      filterExpression?: string;
      maximumNumberOfResults?: number;
      resourceMetadata?: string[];
      connectionIds?: string[];
    }
  ): Promise<IRetrievalResponse> {
    if (!queryString || queryString.trim() === '') {
      return { retrievalHits: [] };
    }

    const body: Record<string, unknown> = {
      queryString,
      dataSource,
    };

    if (options?.filterExpression) {
      body.filterExpression = options.filterExpression;
    }
    if (options?.maximumNumberOfResults) {
      body.maximumNumberOfResults = String(options.maximumNumberOfResults);
    }
    if (options?.resourceMetadata) {
      body.resourceMetadata = options.resourceMetadata;
    }
    if (dataSource === 'externalItem' && options?.connectionIds?.length) {
      body.dataSourceConfiguration = {
        externalItem: {
          connections: options.connectionIds.map((id) => ({ connectionId: id })),
        },
      };
    }

    console.log(`[CopilotRetrievalService] POST ${COPILOT_RETRIEVAL_ENDPOINT} (v1.0) source=${dataSource}`, JSON.stringify(body, null, 2));
    const response = await this.graphClient
      .api(COPILOT_RETRIEVAL_ENDPOINT)
      .version('v1.0')
      .post(body);
    console.log(`[CopilotRetrievalService] Retrieved ${response.retrievalHits?.length || 0} hits from ${dataSource}`);

    return {
      retrievalHits: response.retrievalHits || [],
    };
  }

  /**
   * Retrieve from multiple data sources using Graph batch API.
   * Combines results from SharePoint, OneDrive, and external items.
   */
  public async retrieveFromAllSources(
    queryString: string,
    options?: {
      filterExpression?: string;
      maximumNumberOfResults?: number;
      resourceMetadata?: string[];
      connectionIds?: string[];
    }
  ): Promise<IRetrievalHit[]> {
    const sources: RetrievalDataSource[] = ['sharePoint', 'externalItem'];
    const results = await Promise.allSettled(
      sources.map((source) => this.retrieve(queryString, source, options))
    );

    const allHits: IRetrievalHit[] = [];
    for (const result of results) {
      if (result.status === 'fulfilled') {
        allHits.push(...result.value.retrievalHits);
      }
    }

    // Sort by relevance score (highest first)
    allHits.sort((a, b) => {
      const scoreA = a.extracts[0]?.relevanceScore ?? 0;
      const scoreB = b.extracts[0]?.relevanceScore ?? 0;
      return scoreB - scoreA;
    });

    return allHits;
  }

  /**
   * Formats retrieval hits into a context string with numbered citations
   * suitable for display or for passing to any LLM.
   */
  public static formatAsContext(hits: IRetrievalHit[]): {
    contextText: string;
    citations: Array<{ index: number; title: string; url: string; snippet: string }>;
  } {
    const citations = hits.map((hit, i) => ({
      index: i + 1,
      title: hit.resourceMetadata?.title || hit.webUrl.split('/').pop() || 'Document',
      url: hit.webUrl,
      snippet: hit.extracts[0]?.text?.substring(0, 200) || '',
    }));

    const contextText = hits
      .map((hit, i) => {
        const title = hit.resourceMetadata?.title || 'Document';
        const extracts = hit.extracts.map((e) => e.text).join('\n');
        return `[${i + 1}] Title: ${title}\nSource: ${hit.webUrl}\nContent: ${extracts}`;
      })
      .join('\n\n');

    return { contextText, citations };
  }
}

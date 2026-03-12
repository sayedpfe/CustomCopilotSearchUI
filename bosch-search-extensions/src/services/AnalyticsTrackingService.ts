import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointListService } from './SharePointListService';
import { ISearchEvent } from '../models';
import { LIST_ANALYTICS_EVENTS, ANALYTICS_BATCH_SIZE, ANALYTICS_FLUSH_INTERVAL_MS } from '../common/Constants';
import { getSessionId } from '../common/Utils';

export class AnalyticsTrackingService {
  private listService: SharePointListService;
  private eventBuffer: ISearchEvent[] = [];
  private flushIntervalId: ReturnType<typeof setInterval> | undefined;
  private sessionId: string;

  constructor(context: WebPartContext) {
    this.listService = new SharePointListService(context);
    this.sessionId = getSessionId();
    this.startAutoFlush();
  }

  public trackQuery(query: string, resultCount: number, vertical?: string): void {
    this.addEvent({
      query,
      eventType: resultCount === 0 ? 'ZeroResult' : 'Query',
      resultCount,
      timestamp: new Date(),
      sessionId: this.sessionId,
      vertical,
    });
  }

  public trackClick(query: string, clickedUrl: string, position: number): void {
    this.addEvent({
      query,
      eventType: 'Click',
      clickedUrl,
      clickPosition: position,
      timestamp: new Date(),
      sessionId: this.sessionId,
    });
  }

  private addEvent(event: ISearchEvent): void {
    this.eventBuffer.push(event);
    if (this.eventBuffer.length >= ANALYTICS_BATCH_SIZE) {
      this.flush();
    }
  }

  private startAutoFlush(): void {
    this.flushIntervalId = setInterval(() => {
      this.flush();
    }, ANALYTICS_FLUSH_INTERVAL_MS);

    // Flush on page unload
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', () => this.flush());
    }
  }

  public async flush(): Promise<void> {
    if (this.eventBuffer.length === 0) return;

    const eventsToFlush = [...this.eventBuffer];
    this.eventBuffer = [];

    try {
      const items = eventsToFlush.map((event) => ({
        Title: event.query,
        EventType: event.eventType,
        ResultCount: event.resultCount || 0,
        ClickedUrl: event.clickedUrl ? { Url: event.clickedUrl } : undefined,
        ClickPosition: event.clickPosition || 0,
        Timestamp: event.timestamp.toISOString(),
        SessionId: event.sessionId,
        Vertical: event.vertical || '',
      }));

      await this.listService.batchAddItems(LIST_ANALYTICS_EVENTS, items);
    } catch (err) {
      console.error('[AnalyticsTrackingService] Failed to flush events:', err);
      // Re-add failed events back to buffer (up to a limit to avoid memory leaks)
      if (this.eventBuffer.length < 100) {
        this.eventBuffer.unshift(...eventsToFlush);
      }
    }
  }

  public dispose(): void {
    if (this.flushIntervalId !== undefined) {
      clearInterval(this.flushIntervalId);
    }
    this.flush();
  }
}

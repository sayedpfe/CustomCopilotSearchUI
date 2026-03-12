import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphSearchService } from './GraphSearchService';
import { SharePointListService } from './SharePointListService';
import { PeopleGraphService } from './PeopleGraphService';
import { AnalyticsTrackingService } from './AnalyticsTrackingService';
import { CopilotChatService } from './CopilotChatService';
import { CopilotRetrievalService } from './CopilotRetrievalService';
import { CopilotDetectionService } from './CopilotDetectionService';

/**
 * Singleton factory that provides shared service instances across all web parts.
 * Call ServiceFactory.initialize(context) from the first web part that loads.
 */
export class ServiceFactory {
  private static context: WebPartContext;
  private static graphClient: MSGraphClientV3;
  private static graphSearchService: GraphSearchService;
  private static spListService: SharePointListService;
  private static peopleGraphService: PeopleGraphService;
  private static analyticsTrackingService: AnalyticsTrackingService;
  private static copilotChatService: CopilotChatService;
  private static copilotRetrievalService: CopilotRetrievalService;
  private static copilotDetectionService: CopilotDetectionService;
  private static initialized: boolean = false;

  public static async initialize(context: WebPartContext): Promise<void> {
    if (ServiceFactory.initialized) return;

    ServiceFactory.context = context;
    ServiceFactory.graphClient = await context.msGraphClientFactory.getClient('3') as MSGraphClientV3;
    ServiceFactory.initialized = true;
  }

  public static getContext(): WebPartContext {
    ServiceFactory.ensureInitialized();
    return ServiceFactory.context;
  }

  public static getGraphSearchService(): GraphSearchService {
    ServiceFactory.ensureInitialized();
    if (!ServiceFactory.graphSearchService) {
      ServiceFactory.graphSearchService = new GraphSearchService(ServiceFactory.graphClient);
    }
    return ServiceFactory.graphSearchService;
  }

  public static getSharePointListService(): SharePointListService {
    ServiceFactory.ensureInitialized();
    if (!ServiceFactory.spListService) {
      ServiceFactory.spListService = new SharePointListService(ServiceFactory.context);
    }
    return ServiceFactory.spListService;
  }

  public static getPeopleGraphService(): PeopleGraphService {
    ServiceFactory.ensureInitialized();
    if (!ServiceFactory.peopleGraphService) {
      ServiceFactory.peopleGraphService = new PeopleGraphService(ServiceFactory.graphClient);
    }
    return ServiceFactory.peopleGraphService;
  }

  public static getAnalyticsTrackingService(): AnalyticsTrackingService {
    ServiceFactory.ensureInitialized();
    if (!ServiceFactory.analyticsTrackingService) {
      ServiceFactory.analyticsTrackingService = new AnalyticsTrackingService(ServiceFactory.context);
    }
    return ServiceFactory.analyticsTrackingService;
  }

  public static getCopilotChatService(): CopilotChatService {
    ServiceFactory.ensureInitialized();
    if (!ServiceFactory.copilotChatService) {
      ServiceFactory.copilotChatService = new CopilotChatService(ServiceFactory.graphClient);
    }
    return ServiceFactory.copilotChatService;
  }

  public static getCopilotRetrievalService(): CopilotRetrievalService {
    ServiceFactory.ensureInitialized();
    if (!ServiceFactory.copilotRetrievalService) {
      ServiceFactory.copilotRetrievalService = new CopilotRetrievalService(ServiceFactory.graphClient);
    }
    return ServiceFactory.copilotRetrievalService;
  }

  public static getCopilotDetectionService(): CopilotDetectionService {
    ServiceFactory.ensureInitialized();
    if (!ServiceFactory.copilotDetectionService) {
      ServiceFactory.copilotDetectionService = new CopilotDetectionService(ServiceFactory.graphClient);
    }
    return ServiceFactory.copilotDetectionService;
  }

  public static getGraphClient(): MSGraphClientV3 {
    ServiceFactory.ensureInitialized();
    return ServiceFactory.graphClient;
  }

  private static ensureInitialized(): void {
    if (!ServiceFactory.initialized) {
      throw new Error('[ServiceFactory] Not initialized. Call ServiceFactory.initialize(context) first.');
    }
  }
}

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { COPILOT_SKU_IDS, CACHE_COPILOT_LICENSE_MS } from '../common/Constants';

export class CopilotDetectionService {
  private graphClient: MSGraphClientV3;
  private cachedResult: boolean | undefined;
  private cacheExpiry: number = 0;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  public async hasCopilotLicense(): Promise<boolean> {
    // Return cached result if still valid
    if (this.cachedResult !== undefined && Date.now() < this.cacheExpiry) {
      return this.cachedResult;
    }

    try {
      console.log('[CopilotDetectionService] GET /me/licenseDetails — Checking Copilot license');
      const response = await this.graphClient
        .api('/me/licenseDetails')
        .select('skuId')
        .get();

      const licenses: Array<{ skuId: string }> = response.value || [];
      const hasCopilot = licenses.some((license) =>
        COPILOT_SKU_IDS.includes(license.skuId)
      );
      console.log(`[CopilotDetectionService] Copilot license: ${hasCopilot}, total licenses: ${licenses.length}`);

      this.cachedResult = hasCopilot;
      this.cacheExpiry = Date.now() + CACHE_COPILOT_LICENSE_MS;
      return hasCopilot;
    } catch (err) {
      console.error('[CopilotDetectionService] Error checking license:', err);
      this.cachedResult = false;
      this.cacheExpiry = Date.now() + CACHE_COPILOT_LICENSE_MS;
      return false;
    }
  }
}

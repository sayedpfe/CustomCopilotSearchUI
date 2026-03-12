import { MSGraphClientV3 } from '@microsoft/sp-http';

const IMAGE_MIME_PREFIXES = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];
const CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

interface CacheEntry {
  url: string;
  expires: number;
}

interface ParsedLibraryUrl {
  hostname: string;
  sitePath: string;   // e.g. /sites/corp  or '' for root site
  libraryName: string; // e.g. SiteAssets
  subfolder: string;   // e.g. Backgrounds  or '' if root of library
}

/**
 * Loads images from a SharePoint picture/document library and returns a URL
 * for the image assigned to today's date (day-of-year modulo image count),
 * so the background changes once per day automatically.
 */
export class BackgroundService {
  private cache: Map<string, CacheEntry> = new Map();

  constructor(private graphClient: MSGraphClientV3) {}

  /**
   * Returns the direct download URL for today's background image, or undefined
   * if the library is empty, inaccessible, or the URL could not be parsed.
   *
   * @param libraryUrl  Full URL to the SharePoint library or library subfolder,
   *                    e.g. https://contoso.sharepoint.com/sites/corp/SiteAssets/Backgrounds
   */
  public async getDailyBackgroundImageUrl(libraryUrl: string): Promise<string | undefined> {
    if (!libraryUrl || !libraryUrl.startsWith('http')) return undefined;

    const cacheKey = `bg|${libraryUrl}`;
    const cached = this.cache.get(cacheKey);
    if (cached && cached.expires > Date.now()) return cached.url;

    try {
      const parsed = this.parseLibraryUrl(libraryUrl);
      if (!parsed) {
        console.warn('[BackgroundService] Could not parse library URL:', libraryUrl);
        return undefined;
      }

      const { hostname, sitePath, libraryName, subfolder } = parsed;
      console.log(`[BackgroundService] Parsed URL → hostname=${hostname} sitePath=${sitePath} library=${libraryName} subfolder=${subfolder}`);

      // 1. Resolve site ID via Graph
      const siteApiPath = sitePath
        ? `/sites/${hostname}:${sitePath}`
        : `/sites/${hostname}`;
      const site = await this.graphClient
        .api(siteApiPath)
        .select('id')
        .get();
      const siteId: string = site.id;
      console.log(`[BackgroundService] Site ID: ${siteId}`);

      // 2. Find the drive whose name matches the library.
      //    SharePoint libraries each appear as a separate Drive in Graph.
      const drivesResponse = await this.graphClient
        .api(`/sites/${siteId}/drives`)
        .select('id,name,webUrl')
        .get();
      const drives: Record<string, string>[] = drivesResponse.value || [];
      console.log(`[BackgroundService] Available drives: ${drives.map((d) => d.name).join(', ')}`);

      // Match by display name (case-insensitive) or by webUrl ending with the library name
      const libraryNameLower = libraryName.toLowerCase();
      let targetDrive = drives.find(
        (d) =>
          (d.name || '').toLowerCase() === libraryNameLower ||
          (d.webUrl || '').toLowerCase().endsWith(`/${libraryNameLower}`)
      );

      if (!targetDrive) {
        // Fall back: try the default drive (Documents library)
        console.warn(`[BackgroundService] Drive "${libraryName}" not found. Falling back to default drive.`);
        const defaultDrive = await this.graphClient
          .api(`/sites/${siteId}/drive`)
          .select('id,name')
          .get();
        targetDrive = defaultDrive;
      }

      const driveId: string = targetDrive.id;
      console.log(`[BackgroundService] Using drive: ${targetDrive.name} (${driveId})`);

      // 3. List items in the drive root or the specified subfolder
      const itemsApiPath = subfolder
        ? `/drives/${driveId}/root:/${subfolder}:/children`
        : `/drives/${driveId}/root/children`;

      const itemsResponse = await this.graphClient
        .api(itemsApiPath)
        .select('id,name,file')
        .top(200)
        .get();

      const allItems: Record<string, unknown>[] = itemsResponse.value || [];

      // 4. Keep only recognised image MIME types
      const imageItems = allItems.filter((item) => {
        const file = item.file as Record<string, string> | undefined;
        if (!file?.mimeType) return false;
        return IMAGE_MIME_PREFIXES.some((prefix) => file.mimeType.startsWith(prefix));
      });

      if (imageItems.length === 0) {
        console.warn('[BackgroundService] No image files found. Items in folder:', allItems.map((i) => i.name).join(', '));
        return undefined;
      }

      // 5. Pick by day-of-year — consistent all day, rotates daily
      const dayIndex = this.getDayOfYear() % imageItems.length;
      const pickedItem = imageItems[dayIndex];
      const pickedId = pickedItem.id as string;

      // 6. Fetch the pre-authenticated download URL for the chosen item
      const fileDetails = await this.graphClient
        .api(`/drives/${driveId}/items/${pickedId}`)
        .select('id,@microsoft.graph.downloadUrl,webUrl')
        .get();

      const url =
        (fileDetails['@microsoft.graph.downloadUrl'] as string) ||
        (fileDetails.webUrl as string);

      if (!url) return undefined;

      this.cache.set(cacheKey, { url, expires: Date.now() + CACHE_TTL_MS });
      console.log(`[BackgroundService] Daily background: ${pickedItem.name as string} (slot ${dayIndex + 1}/${imageItems.length})`);
      return url;

    } catch (err) {
      console.warn('[BackgroundService] Failed to fetch background image:', (err as Error).message || err);
      return undefined;
    }
  }

  /**
   * Parses a full SharePoint library URL into its components.
   *
   * Supported formats:
   *   https://tenant.sharepoint.com/sites/corp/SiteAssets
   *   https://tenant.sharepoint.com/sites/corp/SiteAssets/Backgrounds
   *   https://tenant.sharepoint.com/teams/team/LibraryName/SubFolder
   *   https://tenant.sharepoint.com/LibraryName          (root site)
   */
  private parseLibraryUrl(rawUrl: string): ParsedLibraryUrl | undefined {
    try {
      const parsed = new URL(rawUrl);
      const hostname = parsed.hostname;
      // Strip trailing slashes and split, filtering out empty segments
      const parts = parsed.pathname.replace(/\/+$/, '').split('/').filter(Boolean);

      // Determine how many path segments make up the site address
      // /sites/siteName  →  2 segments
      // /teams/teamName  →  2 segments
      // (root site)      →  0 segments
      let siteSegmentCount = 0;
      if (parts.length > 0 && (parts[0] === 'sites' || parts[0] === 'teams' || parts[0] === 'personal')) {
        siteSegmentCount = 2;
      }

      const siteSegments = parts.slice(0, siteSegmentCount);
      const sitePath = siteSegments.length > 0 ? '/' + siteSegments.join('/') : '';

      const afterSite = parts.slice(siteSegmentCount);
      if (afterSite.length === 0) {
        console.warn('[BackgroundService] URL contains no library name after site path:', rawUrl);
        return undefined;
      }

      // First segment after the site = library (Drive) name
      // Remaining segments = subfolder path inside that drive
      const libraryName = decodeURIComponent(afterSite[0]);
      const subfolder = afterSite.slice(1).map(decodeURIComponent).join('/');

      return { hostname, sitePath, libraryName, subfolder };
    } catch {
      return undefined;
    }
  }

  /** Returns a day-of-year integer (1–366). Consistent within a calendar day. */
  private getDayOfYear(): number {
    const now = new Date();
    const startOfYear = new Date(now.getFullYear(), 0, 0);
    const diff = now.getTime() - startOfYear.getTime();
    return Math.floor(diff / (1000 * 60 * 60 * 24));
  }
}

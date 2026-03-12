import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class SharePointListService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  private get siteUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  public async getListItems<T>(
    listTitle: string,
    filter?: string,
    select?: string[],
    orderBy?: string,
    top?: number
  ): Promise<T[]> {
    let url = `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/items?`;

    const params: string[] = [];
    if (select && select.length > 0) {
      params.push(`$select=${select.join(',')}`);
    }
    if (filter) {
      params.push(`$filter=${filter}`);
    }
    if (orderBy) {
      params.push(`$orderby=${orderBy}`);
    }
    if (top) {
      params.push(`$top=${top}`);
    }
    url += params.join('&');

    console.log(`[SharePointListService] GET ${url}`);
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to fetch list items from "${listTitle}": ${response.status} ${errorText}`);
    }

    const data = await response.json();
    return data.value as T[];
  }

  public async addListItem(listTitle: string, item: Record<string, unknown>): Promise<unknown> {
    const url = `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/items`;

    console.log(`[SharePointListService] POST ${url}`, JSON.stringify(item));
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
        body: JSON.stringify(item),
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to add item to "${listTitle}": ${response.status} ${errorText}`);
    }

    return response.json();
  }

  public async batchAddItems(listTitle: string, items: Record<string, unknown>[]): Promise<void> {
    // For simplicity, add items sequentially. For high-volume production use,
    // implement $batch API or switch to Application Insights.
    const promises = items.map((item) => this.addListItem(listTitle, item));
    await Promise.all(promises);
  }
}

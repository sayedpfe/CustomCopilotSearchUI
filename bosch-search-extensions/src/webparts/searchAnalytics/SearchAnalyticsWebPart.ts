import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider } from '@microsoft/sp-webpart-base';
import { SearchAnalyticsDashboard, ISearchAnalyticsDashboardProps } from './components/SearchAnalyticsDashboard';

export interface ISearchAnalyticsWebPartProps {
  listName: string;
  defaultDays: number;
  maxRows: number;
}

export default class SearchAnalyticsWebPart extends BaseClientSideWebPart<ISearchAnalyticsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISearchAnalyticsDashboardProps> = React.createElement(SearchAnalyticsDashboard, {
      context: this.context,
      listName: this.properties.listName || 'SearchAnalyticsEvents',
      defaultDays: this.properties.defaultDays || 30,
      maxRows: this.properties.maxRows || 20,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configure the search analytics dashboard.' },
          groups: [
            {
              groupName: 'Analytics Configuration',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'SharePoint list name',
                }),
                PropertyPaneSlider('defaultDays', {
                  label: 'Default date range (days)',
                  min: 7,
                  max: 90,
                  value: this.properties.defaultDays || 30,
                }),
                PropertyPaneSlider('maxRows', {
                  label: 'Maximum rows to display',
                  min: 5,
                  max: 50,
                  value: this.properties.maxRows || 20,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

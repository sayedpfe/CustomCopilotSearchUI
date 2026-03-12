import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
} from '@microsoft/sp-webpart-base';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { BoschSearchApp, IBoschSearchAppProps } from './components/BoschSearchApp';

// Register Fluent UI icon font once — required for all <Icon iconName="..." /> to render
initializeIcons();

export interface IBoschSearchAppWebPartProps {
  groundingMode: 'work' | 'web' | 'both';
  maxRetrievalResults: number;
  showCopilotLink: boolean;
  newsSourceSiteUrl: string;
  promotedResultsListName: string;
  announcementsListName: string;
  analyticsListName: string;
  // Dynamic background from SharePoint image library
  backgroundEnabled: boolean;
  backgroundLibraryUrl: string;
}

export default class BoschSearchAppWebPart extends BaseClientSideWebPart<IBoschSearchAppWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBoschSearchAppProps> = React.createElement(BoschSearchApp, {
      context: this.context,
      groundingMode: this.properties.groundingMode || 'work',
      maxRetrievalResults: this.properties.maxRetrievalResults || 10,
      showCopilotLink: this.properties.showCopilotLink !== false,
      newsSourceSiteUrl: this.properties.newsSourceSiteUrl || '',
      promotedResultsListName: this.properties.promotedResultsListName || 'SearchPromotedResults',
      announcementsListName: this.properties.announcementsListName || 'SearchAnnouncements',
      analyticsListName: this.properties.analyticsListName || 'SearchAnalyticsEvents',
      backgroundEnabled: this.properties.backgroundEnabled === true,
      backgroundLibraryUrl: this.properties.backgroundLibraryUrl || '',
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Refresh the property pane (and re-render) when backgroundEnabled toggle changes
  // so the conditional URL field appears / disappears immediately.
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'backgroundEnabled') {
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configure the Bosch AI Search experience.' },
          groups: [
            {
              groupName: 'AI Configuration',
              groupFields: [
                PropertyPaneDropdown('groundingMode', {
                  label: 'AI Grounding Mode',
                  options: [
                    { key: 'work', text: 'Work data only (enterprise search)' },
                    { key: 'web', text: 'Web search only' },
                    { key: 'both', text: 'Work + Web (default Copilot behavior)' },
                  ],
                  selectedKey: this.properties.groundingMode || 'work',
                }),
                PropertyPaneSlider('maxRetrievalResults', {
                  label: 'Max retrieval results',
                  min: 1,
                  max: 25,
                  value: this.properties.maxRetrievalResults || 10,
                }),
                PropertyPaneToggle('showCopilotLink', {
                  label: 'Show "Open in Copilot" link',
                  checked: this.properties.showCopilotLink !== false,
                }),
              ],
            },
            {
              groupName: 'Data Sources',
              groupFields: [
                PropertyPaneTextField('newsSourceSiteUrl', {
                  label: 'News source site URL (blank = current site)',
                }),
                PropertyPaneTextField('promotedResultsListName', {
                  label: 'Promoted results list name',
                }),
                PropertyPaneTextField('announcementsListName', {
                  label: 'Announcements list name',
                }),
                PropertyPaneTextField('analyticsListName', {
                  label: 'Analytics events list name',
                }),
              ],
            },
            {
              groupName: 'Background Image',
              groupFields: [
                PropertyPaneToggle('backgroundEnabled', {
                  label: 'Enable dynamic background',
                  onText: 'Enabled — rotates daily from image library',
                  offText: 'Disabled — default solid background',
                  checked: this.properties.backgroundEnabled === true,
                }),
                ...(this.properties.backgroundEnabled === true
                  ? [
                      PropertyPaneTextField('backgroundLibraryUrl', {
                        label: 'Image library URL',
                        description:
                          'Full URL to a SharePoint document or picture library (or a subfolder). ' +
                          'e.g. https://contoso.sharepoint.com/sites/corp/SiteAssets/Backgrounds',
                        placeholder: 'https://contoso.sharepoint.com/sites/…/LibraryName',
                        multiline: false,
                      }),
                    ]
                  : []),
              ],
            },
          ],
        },
      ],
    };
  }
}

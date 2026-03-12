import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider, PropertyPaneToggle } from '@microsoft/sp-webpart-base';
import { AnnouncementsBanner, IAnnouncementsBannerProps } from './components/AnnouncementsBanner';

export interface IAnnouncementsBannerWebPartProps {
  listName: string;
  maxItems: number;
  allowDismiss: boolean;
}

export default class AnnouncementsBannerWebPart extends BaseClientSideWebPart<IAnnouncementsBannerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAnnouncementsBannerProps> = React.createElement(AnnouncementsBanner, {
      context: this.context,
      listName: this.properties.listName || 'SearchAnnouncements',
      maxItems: this.properties.maxItems || 5,
      allowDismiss: this.properties.allowDismiss !== false,
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
          header: { description: 'Configure the announcements banner.' },
          groups: [
            {
              groupName: 'Announcements Configuration',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'SharePoint list name',
                  value: this.properties.listName || 'SearchAnnouncements',
                }),
                PropertyPaneSlider('maxItems', {
                  label: 'Maximum announcements to show',
                  min: 1,
                  max: 10,
                  value: this.properties.maxItems || 5,
                }),
                PropertyPaneToggle('allowDismiss', {
                  label: 'Allow users to dismiss announcements',
                  checked: this.properties.allowDismiss !== false,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

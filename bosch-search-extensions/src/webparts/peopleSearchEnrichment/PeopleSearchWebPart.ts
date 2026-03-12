import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneToggle, PropertyPaneSlider } from '@microsoft/sp-webpart-base';
import { PeopleSearch, IPeopleSearchProps } from './components/PeopleSearch';

export interface IPeopleSearchWebPartProps {
  showOrgChart: boolean;
  showRecentDocs: boolean;
  maxRecentDocs: number;
}

export default class PeopleSearchWebPart extends BaseClientSideWebPart<IPeopleSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPeopleSearchProps> = React.createElement(PeopleSearch, {
      context: this.context,
      showOrgChart: this.properties.showOrgChart !== false,
      showRecentDocs: this.properties.showRecentDocs !== false,
      maxRecentDocs: this.properties.maxRecentDocs || 5,
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
          header: { description: 'Configure enriched people search cards.' },
          groups: [
            {
              groupName: 'People Search Configuration',
              groupFields: [
                PropertyPaneToggle('showOrgChart', {
                  label: 'Show org chart (manager & reports)',
                  checked: this.properties.showOrgChart !== false,
                }),
                PropertyPaneToggle('showRecentDocs', {
                  label: 'Show recent documents',
                  checked: this.properties.showRecentDocs !== false,
                }),
                PropertyPaneSlider('maxRecentDocs', {
                  label: 'Max recent documents to show',
                  min: 1,
                  max: 10,
                  value: this.properties.maxRecentDocs || 5,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

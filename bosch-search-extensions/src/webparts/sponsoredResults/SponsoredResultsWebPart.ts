import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { SponsoredResults, ISponsoredResultsProps } from './components/SponsoredResults';

export interface ISponsoredResultsWebPartProps {
  listName: string;
}

export default class SponsoredResultsWebPart extends BaseClientSideWebPart<ISponsoredResultsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISponsoredResultsProps> = React.createElement(SponsoredResults, {
      context: this.context,
      listName: this.properties.listName || 'SearchPromotedResults',
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
          header: { description: 'Configure sponsored results.' },
          groups: [
            {
              groupName: 'Sponsored Results Configuration',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'SharePoint list name',
                  value: this.properties.listName || 'SearchPromotedResults',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

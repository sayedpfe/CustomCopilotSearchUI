import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneToggle } from '@microsoft/sp-webpart-base';
import { IQuickLink } from '../../models';
import { QuickLinks, IQuickLinksProps } from './components/QuickLinks';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IQuickLinksWebPartProps {
  links: IQuickLink[];
  columns: number;
  showDescriptions: boolean;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(QuickLinks, {
      links: this.properties.links || [],
      columns: this.properties.columns || 3,
      showDescriptions: this.properties.showDescriptions !== false,
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
          header: { description: 'Configure quick link tiles that appear above the search area.' },
          groups: [
            {
              groupName: 'Quick Links Configuration',
              groupFields: [
                PropertyFieldCollectionData('links', {
                  key: 'links',
                  label: 'Quick Links',
                  panelHeader: 'Manage Quick Links',
                  manageBtnLabel: 'Manage Links',
                  value: this.properties.links,
                  fields: [
                    {
                      id: 'title',
                      title: 'Title',
                      type: CustomCollectionFieldType.string,
                      required: true,
                    },
                    {
                      id: 'url',
                      title: 'URL',
                      type: CustomCollectionFieldType.url,
                      required: true,
                    },
                    {
                      id: 'iconName',
                      title: 'Icon Name (Fluent UI)',
                      type: CustomCollectionFieldType.string,
                      required: false,
                    },
                    {
                      id: 'description',
                      title: 'Description',
                      type: CustomCollectionFieldType.string,
                      required: false,
                    },
                    {
                      id: 'openInNewTab',
                      title: 'Open in new tab',
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                    },
                  ],
                }),
                PropertyPaneDropdown('columns', {
                  label: 'Number of columns',
                  options: [
                    { key: 2, text: '2 columns' },
                    { key: 3, text: '3 columns' },
                    { key: 4, text: '4 columns' },
                    { key: 6, text: '6 columns' },
                  ],
                  selectedKey: this.properties.columns || 3,
                }),
                PropertyPaneToggle('showDescriptions', {
                  label: 'Show descriptions',
                  checked: this.properties.showDescriptions !== false,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

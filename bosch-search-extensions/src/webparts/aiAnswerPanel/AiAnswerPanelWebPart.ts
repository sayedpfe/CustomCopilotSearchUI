import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneToggle, PropertyPaneDropdown, PropertyPaneSlider } from '@microsoft/sp-webpart-base';
import { AiAnswerPanel, IAiAnswerPanelProps } from './components/AiAnswerPanel';

export interface IAiAnswerPanelWebPartProps {
  groundingMode: 'work' | 'web' | 'both';
  maxRetrievalResults: number;
  showCopilotLink: boolean;
}

export default class AiAnswerPanelWebPart extends BaseClientSideWebPart<IAiAnswerPanelWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAiAnswerPanelProps> = React.createElement(AiAnswerPanel, {
      context: this.context,
      groundingMode: this.properties.groundingMode || 'work',
      maxRetrievalResults: this.properties.maxRetrievalResults || 10,
      showCopilotLink: this.properties.showCopilotLink !== false,
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
          header: { description: 'Configure the AI answer panel. Uses Copilot APIs for licensed users, falls back to Graph Search for others.' },
          groups: [
            {
              groupName: 'Copilot AI Configuration',
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
                  label: 'Max retrieval results (Copilot Retrieval API)',
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
          ],
        },
      ],
    };
  }
}

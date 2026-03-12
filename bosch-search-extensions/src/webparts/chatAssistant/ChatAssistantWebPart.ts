import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider, PropertyPaneDropdown, PropertyPaneToggle } from '@microsoft/sp-webpart-base';
import { ChatAssistant, IChatAssistantProps } from './components/ChatAssistant';

export interface IChatAssistantWebPartProps {
  groundingMode: 'work' | 'web' | 'both';
  maxConversationTurns: number;
  panelMode: 'sidePanel' | 'inline';
  welcomeMessage: string;
  suggestedQuestions: string;
  showCopilotLink: boolean;
}

export default class ChatAssistantWebPart extends BaseClientSideWebPart<IChatAssistantWebPartProps> {

  public render(): void {
    const suggestedQuestions = (this.properties.suggestedQuestions || '')
      .split(',')
      .map((q) => q.trim())
      .filter(Boolean);

    const element: React.ReactElement<IChatAssistantProps> = React.createElement(ChatAssistant, {
      context: this.context,
      groundingMode: this.properties.groundingMode || 'work',
      maxConversationTurns: this.properties.maxConversationTurns || 10,
      panelMode: this.properties.panelMode || 'sidePanel',
      welcomeMessage: this.properties.welcomeMessage,
      suggestedQuestions,
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
          header: { description: 'Configure the chat assistant. Uses Copilot Chat API for licensed users, falls back to Graph Search for others.' },
          groups: [
            {
              groupName: 'Copilot Chat Configuration',
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
                PropertyPaneSlider('maxConversationTurns', {
                  label: 'Max conversation turns',
                  min: 3,
                  max: 20,
                }),
                PropertyPaneDropdown('panelMode', {
                  label: 'Display mode',
                  options: [
                    { key: 'sidePanel', text: 'Side Panel' },
                    { key: 'inline', text: 'Inline' },
                  ],
                }),
                PropertyPaneToggle('showCopilotLink', {
                  label: 'Show "Open in Copilot" link',
                  checked: this.properties.showCopilotLink !== false,
                }),
                PropertyPaneTextField('welcomeMessage', {
                  label: 'Welcome message',
                }),
                PropertyPaneTextField('suggestedQuestions', {
                  label: 'Suggested questions (comma-separated)',
                  multiline: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

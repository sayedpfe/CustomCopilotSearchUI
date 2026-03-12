import * as React from 'react';
import * as ReactDom from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from '@microsoft/sp-webpart-base';
import { CopilotDiagnostics, ICopilotDiagnosticsProps } from './components/CopilotDiagnostics';

export interface ICopilotDiagnosticsWebPartProps {
  endpointMode: 'chat' | 'chatOverStream';
  groundingMode: 'work' | 'web' | 'both';
  defaultQuestion: string;
}

export default class CopilotDiagnosticsWebPart extends BaseClientSideWebPart<ICopilotDiagnosticsWebPartProps> {

  public onInit(): Promise<void> {
    initializeIcons();
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ICopilotDiagnosticsProps> = React.createElement(CopilotDiagnostics, {
      context: this.context,
      endpointMode: this.properties.endpointMode || 'chatOverStream',
      groundingMode: this.properties.groundingMode || 'work',
      defaultQuestion: this.properties.defaultQuestion || 'What are the key company priorities for this year?',
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
          header: {
            description: 'Diagnose and compare /chat vs /chatOverStream performance. Identifies if SharePoint HTTP proxy is buffering SSE responses.',
          },
          groups: [
            {
              groupName: 'Endpoint Configuration',
              groupFields: [
                PropertyPaneDropdown('endpointMode', {
                  label: 'Default endpoint mode',
                  options: [
                    { key: 'chat', text: '/chat — synchronous, single response' },
                    { key: 'chatOverStream', text: '/chatOverStream — SSE streaming, progressive' },
                  ],
                  selectedKey: this.properties.endpointMode || 'chatOverStream',
                }),
                PropertyPaneDropdown('groundingMode', {
                  label: 'Default grounding mode',
                  options: [
                    { key: 'work', text: 'Work data only (enterprise search)' },
                    { key: 'web', text: 'Web search only' },
                    { key: 'both', text: 'Work + Web' },
                  ],
                  selectedKey: this.properties.groundingMode || 'work',
                }),
                PropertyPaneTextField('defaultQuestion', {
                  label: 'Default test question',
                  placeholder: 'What are the key company priorities for this year?',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

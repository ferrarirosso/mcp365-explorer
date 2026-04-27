import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { AadTokenProvider } from '@microsoft/sp-http';

import * as strings from 'FoundryChatWebPartStrings';
import { FoundryChat } from './components/FoundryChat';
import type { IFoundryChatProps } from './components/IFoundryChatProps';

export interface IFoundryChatWebPartProps {
  backendUrl: string;
  backendApiResource: string;
}

export default class FoundryChatWebPart extends BaseClientSideWebPart<IFoundryChatWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _tokenProvider: AadTokenProvider | undefined;

  public render(): void {
    const element: React.ReactElement<IFoundryChatProps> = React.createElement(FoundryChat, {
      backendUrl: this.properties.backendUrl || '',
      backendApiResource: this.properties.backendApiResource || '',
      isDarkTheme: this._isDarkTheme,
      tokenProvider: this._tokenProvider
    });
    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [
          {
            groupName: strings.BackendGroupName,
            groupFields: [
              PropertyPaneTextField('backendUrl', {
                label: strings.BackendUrlFieldLabel,
                description: strings.BackendUrlFieldDescription
              }),
              PropertyPaneTextField('backendApiResource', {
                label: strings.BackendApiResourceFieldLabel,
                description: strings.BackendApiResourceFieldDescription
              })
            ]
          }
        ]
      }]
    };
  }
}

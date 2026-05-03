import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { AadTokenProvider } from '@microsoft/sp-http';

import * as strings from 'ListsChatWebPartStrings';
import { ListsChat } from './components/ListsChat';
import type { IListsChatProps } from './components/IListsChatProps';

export interface IListsChatWebPartProps {
  environmentId: string;
  backendUrl: string;
  backendApiResource: string;
  showTracePane: boolean;
}

export default class ListsChatWebPart extends BaseClientSideWebPart<IListsChatWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _tokenProvider: AadTokenProvider | undefined;

  public render(): void {
    // Compose the canonical Work IQ siteId for the current SharePoint web,
    // so the agent's system prompt can resolve "this site" without the LLM
    // guessing "root" and listing tenant-root system lists by mistake.
    const pc = this.context.pageContext;
    const hostname = new URL(pc.web.absoluteUrl).hostname;
    const currentSiteId = `${hostname},${pc.site.id.toString()},${pc.web.id.toString()}`;
    const currentSiteUrl = pc.web.absoluteUrl;

    const element: React.ReactElement<IListsChatProps> = React.createElement(ListsChat, {
      environmentId: this.properties.environmentId || '',
      backendUrl: this.properties.backendUrl || '',
      backendApiResource: this.properties.backendApiResource || '',
      showTracePane: this.properties.showTracePane !== false,
      currentSiteId,
      currentSiteUrl,
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
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('environmentId', {
                label: strings.EnvironmentIdFieldLabel,
                description: strings.EnvironmentIdFieldDescription
              })
            ]
          },
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
          },
          {
            groupName: strings.DisplayGroupName,
            groupFields: [
              PropertyPaneToggle('showTracePane', {
                label: strings.ShowTracePaneFieldLabel,
                onText: strings.ShowTracePaneOnText,
                offText: strings.ShowTracePaneOffText
              })
            ]
          }
        ]
      }]
    };
  }
}

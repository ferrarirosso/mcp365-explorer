import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { AadTokenProvider } from '@microsoft/sp-http';

import * as strings from 'OneDriveExplorerWebPartStrings';
import { OneDriveExplorer } from './components/OneDriveExplorer';
import type { IOneDriveExplorerProps } from './components/IOneDriveExplorerProps';

export interface IOneDriveExplorerWebPartProps {
  environmentId: string;
}

export default class OneDriveExplorerWebPart extends BaseClientSideWebPart<IOneDriveExplorerWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _tokenProvider: AadTokenProvider | undefined;

  public render(): void {
    const element: React.ReactElement<IOneDriveExplorerProps> = React.createElement(
      OneDriveExplorer,
      { environmentId: this.properties.environmentId || '', isDarkTheme: this._isDarkTheme, tokenProvider: this._tokenProvider, userEmail: this.context.pageContext.user.email, userId: String((this.context.pageContext as unknown as Record<string, unknown>).aadInfo ? ((this.context.pageContext as unknown as Record<string, Record<string, unknown>>).aadInfo.userId || '') : '') }
    );
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

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('environmentId', { label: strings.EnvironmentIdFieldLabel, description: strings.EnvironmentIdFieldDescription })
          ]
        }]
      }]
    };
  }
}

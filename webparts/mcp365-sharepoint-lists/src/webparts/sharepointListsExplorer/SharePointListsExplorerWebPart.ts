import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { AadTokenProvider } from '@microsoft/sp-http';

import * as strings from 'SharePointListsExplorerWebPartStrings';
import { SharePointListsExplorer } from './components/SharePointListsExplorer';
import type { ISharePointListsExplorerProps } from './components/ISharePointListsExplorerProps';

export interface ISharePointListsExplorerWebPartProps {
  environmentId: string;
}

export default class SharePointListsExplorerWebPart extends BaseClientSideWebPart<ISharePointListsExplorerWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _tokenProvider: AadTokenProvider | undefined;

  public render(): void {
    const element: React.ReactElement<ISharePointListsExplorerProps> = React.createElement(
      SharePointListsExplorer,
      { environmentId: this.properties.environmentId || '', isDarkTheme: this._isDarkTheme, tokenProvider: this._tokenProvider, userEmail: this.context.pageContext.user.email, spfxContext: this.context }
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

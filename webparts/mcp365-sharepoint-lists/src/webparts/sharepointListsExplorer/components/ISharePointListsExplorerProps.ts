import type { AadTokenProvider } from '@microsoft/sp-http';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

export interface ISharePointListsExplorerProps {
  environmentId: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
  userEmail: string;
  spfxContext: BaseComponentContext;
}

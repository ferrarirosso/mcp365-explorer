import type { AadTokenProvider } from '@microsoft/sp-http';

export interface IUserProfileExplorerProps {
  environmentId: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
}

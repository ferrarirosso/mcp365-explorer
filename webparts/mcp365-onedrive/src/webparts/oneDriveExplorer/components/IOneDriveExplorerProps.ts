import type { AadTokenProvider } from '@microsoft/sp-http';

export interface IOneDriveExplorerProps {
  environmentId: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
  userEmail: string;
  userId: string;
}

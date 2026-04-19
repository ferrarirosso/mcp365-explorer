import type { AadTokenProvider } from '@microsoft/sp-http';

export interface IWordExplorerProps {
  environmentId: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
  userEmail: string;
  userId: string;
}

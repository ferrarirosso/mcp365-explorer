import type { AadTokenProvider } from '@microsoft/sp-http';

export interface IMailExplorerProps {
  environmentId: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
}

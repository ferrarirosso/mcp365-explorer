import type { AadTokenProvider } from '@microsoft/sp-http';

export interface IFoundryChatProps {
  backendUrl: string;
  backendApiResource: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
}

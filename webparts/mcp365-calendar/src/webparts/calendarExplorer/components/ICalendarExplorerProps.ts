import type { AadTokenProvider } from '@microsoft/sp-http';

export interface ICalendarExplorerProps {
  environmentId: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
}

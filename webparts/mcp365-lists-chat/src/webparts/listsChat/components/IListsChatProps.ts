import type { AadTokenProvider } from '@microsoft/sp-http';

export interface IListsChatProps {
  environmentId: string;
  backendUrl: string;
  backendApiResource: string;
  showTracePane: boolean;
  /**
   * Pre-formatted Work IQ siteId for the page hosting the webpart:
   * `<hostname>,<siteCollectionGuid>,<webGuid>`. Injected into the agent's
   * system prompt so "this site" / "current site" resolves to the actual
   * SharePoint context instead of "root".
   */
  currentSiteId: string;
  /** Friendly URL for the same site, used in the system prompt for cite-ability. */
  currentSiteUrl: string;
  isDarkTheme: boolean;
  tokenProvider: AadTokenProvider | undefined;
}

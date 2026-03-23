declare interface ISharePointListsExplorerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  EnvironmentIdFieldLabel: string;
  EnvironmentIdFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'SharePointListsExplorerWebPartStrings' {
  const strings: ISharePointListsExplorerWebPartStrings;
  export = strings;
}

declare interface IMailExplorerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  EnvironmentIdFieldLabel: string;
  EnvironmentIdFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'MailExplorerWebPartStrings' {
  const strings: IMailExplorerWebPartStrings;
  export = strings;
}

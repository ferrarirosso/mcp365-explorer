declare interface IUserProfileExplorerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  EnvironmentIdFieldLabel: string;
  EnvironmentIdFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'UserProfileExplorerWebPartStrings' {
  const strings: IUserProfileExplorerWebPartStrings;
  export = strings;
}

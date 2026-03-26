declare interface ICalendarExplorerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  EnvironmentIdFieldLabel: string;
  EnvironmentIdFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'CalendarExplorerWebPartStrings' {
  const strings: ICalendarExplorerWebPartStrings;
  export = strings;
}

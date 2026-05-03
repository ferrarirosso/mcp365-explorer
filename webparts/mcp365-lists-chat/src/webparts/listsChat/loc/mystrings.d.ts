declare interface IListsChatWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  BackendGroupName: string;
  DisplayGroupName: string;
  EnvironmentIdFieldLabel: string;
  EnvironmentIdFieldDescription: string;
  BackendUrlFieldLabel: string;
  BackendUrlFieldDescription: string;
  BackendApiResourceFieldLabel: string;
  BackendApiResourceFieldDescription: string;
  ShowTracePaneFieldLabel: string;
  ShowTracePaneOnText: string;
  ShowTracePaneOffText: string;
}

declare module 'ListsChatWebPartStrings' {
  const strings: IListsChatWebPartStrings;
  export = strings;
}

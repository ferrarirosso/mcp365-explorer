declare interface IFoundryChatWebPartStrings {
  PropertyPaneDescription: string;
  BackendGroupName: string;
  BackendUrlFieldLabel: string;
  BackendUrlFieldDescription: string;
  BackendApiResourceFieldLabel: string;
  BackendApiResourceFieldDescription: string;
}

declare module 'FoundryChatWebPartStrings' {
  const strings: IFoundryChatWebPartStrings;
  export = strings;
}

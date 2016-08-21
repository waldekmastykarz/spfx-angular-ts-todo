declare interface IStrings {
  PropertyPaneDescription: string;
  DataGroupName: string;
  ViewGroupName: string;
  SharePointApiUrlFieldLabel: string;
  ToDoListNameFieldLabel: string;
  HideFinishedTasksFieldLabel: string;
}

declare module 'mystrings' {
  const strings: IStrings;
  export = strings;
}

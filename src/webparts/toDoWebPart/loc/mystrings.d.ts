declare interface IToDoStrings {
  PropertyPaneDescription: string;
  DataGroupName: string;
  ViewGroupName: string;
  SharePointApiUrlFieldLabel: string;
  ToDoListNameFieldLabel: string;
  HideFinishedTasksFieldLabel: string;
}

declare module 'toDoStrings' {
  const strings: IToDoStrings;
  export = strings;
}

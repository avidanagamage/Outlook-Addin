declare interface IOutlookAddinWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'OutlookAddinWebPartStrings' {
  const strings: IOutlookAddinWebPartStrings;
  export = strings;
}

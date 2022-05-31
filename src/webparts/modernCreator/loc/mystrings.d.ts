declare interface IModernCreatorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;

  analyticsWeb: string;
  analyticsList: string;
  analyticsListLog: string;

}

declare module 'ModernCreatorWebPartStrings' {
  const strings: IModernCreatorWebPartStrings;
  export = strings;
}

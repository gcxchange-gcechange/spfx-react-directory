declare interface ISpfxReactDirectoryWebpartWebPartStrings {
  userLang: string;
  TitleFieldLabel: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;

  SearchPlaceHolder: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  PagingLabel: string;
  DirectoryMessage: string;
  LoadingText: string;
  SearchBoxLabel: string;
  SendEmailLabel: string;
  StartChatLabel: string;
  NoUserFoundLabelText: string;
  NoUserFoundImageAltText: string;
  NoUserFoundEmailSubject: string;
  NoUserFoundEmailBody: string;
  NoUserFoundEmail: string;

  SearchButtonLabel: String;
}

declare module "SpfxReactDirectoryWebpartWebPartStrings" {
  const strings: ISpfxReactDirectoryWebpartWebPartStrings;
  export = strings;
}

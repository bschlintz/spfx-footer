declare interface IFooterApplicationCustomizerStrings {
  FooterOpenPanelText: string;
  FooterPanelHeaderText: string;
  EmptyFieldLabel: string;
  UnknownUserLabel: string;
  NoDataFieldLabel: string;
  SiteSponsorLabel: string;
  PeoplePickerPlaceholderText: string;
  PrimarySiteAdminLabel: string;
  SiteAdminsLabel: string;
  GroupOwnersLabel: string;
  StorageLabel: string;
  CreatedLabel: string;
  LastUpdatedLabel: string;
  SiteTemplateLabel: string;
  GroupIdLabel: string;
  SupportLabel: string;
  SupportLinkLabel: string;
}

declare module 'FooterApplicationCustomizerStrings' {
  const strings: IFooterApplicationCustomizerStrings;
  export = strings;
}

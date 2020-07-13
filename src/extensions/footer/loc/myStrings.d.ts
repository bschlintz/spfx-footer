declare interface IFooterApplicationCustomizerStrings {
  FooterOpenPanelText: string;
  FooterPanelHeaderText: string;
  EmptyFieldLabel: string;
  NoDataFieldLabel: string;
  SiteSponsorLabel: string;
  PeoplePickerPlaceholderText: string;
  PrimarySiteAdminLabel: string;
  SiteAdminsLabel: string;
  StorageLabel: string;
  CreatedLabel: string;
  LastUpdatedLabel: string;
  SiteTemplateLabel: string;
  SupportLabel: string;
  SupportLinkLabel: string;
}

declare module 'FooterApplicationCustomizerStrings' {
  const strings: IFooterApplicationCustomizerStrings;
  export = strings;
}

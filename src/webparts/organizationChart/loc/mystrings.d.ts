/*!
 * Modern SP Military Org Chart by TehGoodLivin
 * Copyright (c) 2026 Austin Livengood <https://github.com/TehGoodLivin/>
 */
declare interface IOrganizationChartWebPartStrings {
  // Property Pane
  WebPartTitle: string;
  WebPartSettingsGroup: string;
  DisplaySettingsGroup: string;
  LayoutSettingsGroup: string;
  HideWhiteSpaceLabel: string;
  HideMenusLabel: string;
  HideTopNavLabel: string;
  ToggleOnText: string;
  ToggleOffText: string;
  DataSourceGroup: string;
  CardGroup: string;
  CardLinesLabel: string;
  CardLinesDesc: string;
  DetailPanelGroup: string;
  SortGroup: string;
  WebPartTitleLabel: string;
  LineColorLabel: string;
  CardBorderColorLabel: string;
  SelectedGradientStartLabel: string;
  SelectedGradientEndLabel: string;
  SelectedBorderColorLabel: string;
  SelectedFontColorLabel: string;
  ButtonSettingsGroup: string;
  AddPositionBtnColorLabel: string;
  SaveImageBtnColorLabel: string;
  DeleteBtnColorLabel: string;
  CancelBtnColorLabel: string;
  SaveBtnColorLabel: string;
  AddFormBtnColorLabel: string;
  MoveBtnColorLabel: string;
  OkBtnColorLabel: string;
  SpecialtyCodeColorsLabel: string;
  SpecialtyCodeColorsDesc: string;
  SiteUrlLabel: string;
  SiteUrlDesc: string;
  ListNameLabel: string;
  ListNameDesc: string;
  UnitFilterLockLabel: string;
  UnitFilterLockDesc: string;
  TitleFieldLabel: string;
  TitleFieldDesc: string;
  UnitFieldLabel: string;
  UnitFieldDesc: string;
  OfficeFieldLabel: string;
  OfficeFieldDesc: string;
  RankFieldLabel: string;
  RankFieldDesc: string;
  NameFieldLabel: string;
  NameFieldDesc: string;
  BilletFieldLabel: string;
  BilletFieldDesc: string;
  OwnerFieldLabel: string;
  OwnerFieldDesc: string;
  SpecialtyCodeFieldLabel: string;
  SpecialtyCodeFieldDesc: string;
  PanelColumnsLabel: string;
  PanelColumnsDesc: string;
  OfficeSortPriorityLabel: string;
  OfficeSortPriorityDesc: string;
  UnitSortPriorityLabel: string;
  UnitSortPriorityDesc: string;
  LeadershipOfficesLabel: string;
  LeadershipOfficesDesc: string;

  // Page Edit Mode
  SavePageMessage: string;

  // Config & Loading
  ConfigureMessage: string;
  LoadingMessage: string;
  NoDataMessage: string;

  // Toolbar
  AddPositionButton: string;
  SaveAsImageButton: string;

  // Search & Filter
  SearchPlaceholder: string;
  AllUnitsOption: string;
  AllOfficesOption: string;

  // Warnings
  OrphanWarningPrefix: string;
  OrphanWarningSuffix_Singular: string;
  OrphanWarningSuffix_Plural: string;
  InvalidPanelColsPrefix: string;
  InvalidPanelColsSuffix_Singular: string;
  InvalidPanelColsSuffix_Plural: string;
  InvalidPanelColsTail: string;

  // Detail Panel
  UnitOfficeLabel: string;
  ReportsToLabel: string;
  DirectReportsLabel: string;
  SubordinatesLabel: string;
  BilletPanelPrefix: string;
  PanelConfigureHint: string;

  // Add Modal
  AddPositionTitle: string;
  BilletInputLabel: string;
  RankInputLabel: string;
  NameInputLabel: string;
  TitleInputLabel: string;
  OfficeInputLabel: string;
  UnitInputLabel: string;
  SpecialtyCodeInputLabel: string;
  OwnerInputLabel: string;
  OwnerInputPlaceholder: string;
  RequiredPrefix: string;
  AddSavingLabel: string;
  AddButtonLabel: string;
  AddFailedError: string;
  CancelButton: string;

  // Edit Modal
  DeleteButton: string;
  EditPositionTitle: string;
  BilletChangeWarning: string;
  BilletCollisionError: string;
  EditSavingLabel: string;
  SaveChangesButton: string;
  EditFailedError: string;

  // Remove Modal
  RemovePositionTitle: string;
  RemovePromptPrefix: string;
  RemoveReportsWarningPrefix: string;
  RemoveReportsWarningSuffix_Singular: string;
  RemoveReportsWarningSuffix_Plural: string;
  RemoveConfirmQuestion: string;
  RemovingLabel: string;
  YesRemoveButton: string;

  // Move (Reparent) Modal
  MovePositionTitle: string;
  MovePromptPrefix: string;
  MovePromptMiddle: string;
  MoveTopLevel: string;
  MoveSavingLabel: string;
  MoveButton: string;
  MoveBlockedCrossUnit: string;

  // Drag-to-scroll
  RootDropZoneText: string;

  // Card
  VacantLabel: string;
  EditPositionTooltip: string;
  RemovePositionTooltip: string;
}

declare module 'OrganizationChartWebPartStrings' {
  const strings: IOrganizationChartWebPartStrings;
  export = strings;
}

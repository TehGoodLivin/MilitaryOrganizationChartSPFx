/*!
 * Modern SP Military Org Chart by TehGoodLivin
 * Copyright (c) 2026 Austin Livengood <https://github.com/TehGoodLivin/>
 */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IOrganizationChartProps {
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;

  // List config
  siteUrl: string;
  list: string;
  unitFilterLock: string;

  // Field mappings (all plain string fields)
  titleField: string;
  unitField: string;
  officeField: string;
  rankField: string;
  nameField: string;
  billetField: string;
  ownerField: string;
  specialtyCodeField: string;

  // Detail panel extra columns (comma-separated)
  panelColumns: string;

  // Sort priorities (comma-separated)
  officeSortPriority: string;
  unitSortPriority: string;

  // Leadership offices (comma-separated) – always visible during search
  leadershipOffices: string;

  // Card extra lines (JSON array of arrays)
  cardLines: string;

  // Card colors (specialty-code driven)
  specialtyCodeColors: string;  // JSON array of {Code,GradientStart,GradientEnd,FontColor}
  lineColor: string;
  borderColor: string;

  // Selected card colors
  selectedGradientStart: string;
  selectedGradientEnd: string;
  selectedBorderColor: string;
  selectedFontColor: string;

  // Button colors
  addPositionBtnColor: string;
  saveImageBtnColor: string;
  deleteBtnColor: string;
  cancelBtnColor: string;
  saveBtnColor: string;
  addFormBtnColor: string;
  moveBtnColor: string;
  okBtnColor: string;

  // Display
  webpartTitle: string;

  // Layout toggles
  hideWhiteSpace: boolean;
  hideMenus: boolean;
  hideTopNav: boolean;

  // Page mode
  isPageEditMode: boolean;

  // SharePoint theme
  themeVariant?: IReadonlyTheme;
}

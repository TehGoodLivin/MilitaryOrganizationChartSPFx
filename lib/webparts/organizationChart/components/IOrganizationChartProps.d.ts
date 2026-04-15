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
    siteUrl: string;
    list: string;
    unitFilterLock: string;
    titleField: string;
    unitField: string;
    officeField: string;
    rankField: string;
    nameField: string;
    billetField: string;
    ownerField: string;
    specialtyCodeField: string;
    panelColumns: string;
    officeSortPriority: string;
    unitSortPriority: string;
    leadershipOffices: string;
    cardLines: string;
    specialtyCodeColors: string;
    lineColor: string;
    borderColor: string;
    selectedGradientStart: string;
    selectedGradientEnd: string;
    selectedBorderColor: string;
    selectedFontColor: string;
    addPositionBtnColor: string;
    saveImageBtnColor: string;
    deleteBtnColor: string;
    cancelBtnColor: string;
    saveBtnColor: string;
    addFormBtnColor: string;
    moveBtnColor: string;
    okBtnColor: string;
    webpartTitle: string;
    hideWhiteSpace: boolean;
    hideMenus: boolean;
    hideTopNav: boolean;
    isPageEditMode: boolean;
    themeVariant?: IReadonlyTheme;
}
//# sourceMappingURL=IOrganizationChartProps.d.ts.map
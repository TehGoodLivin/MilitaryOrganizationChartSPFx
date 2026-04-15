import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IOrganizationChartWebPartProps {
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
    specialtyCodeColors: string;
    cardLines: string;
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
}
export default class OrganizationChartWebPart extends BaseClientSideWebPart<IOrganizationChartWebPartProps> {
    private _themeProvider;
    private _themeVariant;
    private _handleThemeChangedEvent;
    render(): void;
    private _injectModernFullWidthStyles;
    protected onInit(): Promise<void>;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    private _rgbaToHex;
    private _renderColorField;
    private _colorField;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=OrganizationChartWebPart.d.ts.map
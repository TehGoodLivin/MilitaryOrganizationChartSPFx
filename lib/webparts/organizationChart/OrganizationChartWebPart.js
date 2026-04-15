var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
/*!
 * Modern SP Military Org Chart by TehGoodLivin
 * Copyright (c) 2026 Austin Livengood <https://github.com/TehGoodLivin/>
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { ThemeProvider } from '@microsoft/sp-component-base';
import { PropertyPaneFieldType, PropertyPaneTextField, PropertyPaneToggle, } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'OrganizationChartWebPartStrings';
import OrganizationChart from './components/OrganizationChart';
export default class OrganizationChartWebPart extends BaseClientSideWebPart {
    _handleThemeChangedEvent(args) {
        this._themeVariant = args.theme;
        this.render();
    }
    render() {
        const isEditMode = this.displayMode === DisplayMode.Edit;
        // Inject full-width styles for Modern pages
        this._injectModernFullWidthStyles();
        const element = React.createElement(OrganizationChart, {
            environmentMessage: '',
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName,
            context: this.context,
            siteUrl: this.properties.siteUrl,
            list: this.properties.list,
            unitFilterLock: this.properties.unitFilterLock,
            titleField: this.properties.titleField,
            unitField: this.properties.unitField,
            officeField: this.properties.officeField,
            rankField: this.properties.rankField,
            nameField: this.properties.nameField,
            billetField: this.properties.billetField,
            ownerField: this.properties.ownerField,
            specialtyCodeField: this.properties.specialtyCodeField,
            panelColumns: this.properties.panelColumns,
            officeSortPriority: this.properties.officeSortPriority,
            unitSortPriority: this.properties.unitSortPriority,
            leadershipOffices: this.properties.leadershipOffices,
            specialtyCodeColors: this.properties.specialtyCodeColors,
            cardLines: this.properties.cardLines,
            lineColor: this.properties.lineColor,
            borderColor: this.properties.borderColor,
            selectedGradientStart: this.properties.selectedGradientStart,
            selectedGradientEnd: this.properties.selectedGradientEnd,
            selectedBorderColor: this.properties.selectedBorderColor,
            selectedFontColor: this.properties.selectedFontColor,
            addPositionBtnColor: this.properties.addPositionBtnColor,
            saveImageBtnColor: this.properties.saveImageBtnColor,
            deleteBtnColor: this.properties.deleteBtnColor,
            cancelBtnColor: this.properties.cancelBtnColor,
            saveBtnColor: this.properties.saveBtnColor,
            addFormBtnColor: this.properties.addFormBtnColor,
            moveBtnColor: this.properties.moveBtnColor,
            okBtnColor: this.properties.okBtnColor,
            webpartTitle: this.properties.webpartTitle,
            hideWhiteSpace: this.properties.hideWhiteSpace !== false,
            hideMenus: this.properties.hideMenus !== false,
            hideTopNav: this.properties.hideTopNav !== false,
            isPageEditMode: isEditMode,
            themeVariant: this._themeVariant,
        });
        ReactDom.render(element, this.domElement);
    }
    // ── Modern page full-width & hide SP chrome ──────────────────────────────
    _injectModernFullWidthStyles() {
        // Remove previous injection so toggles take effect on re-render
        const existing = document.getElementById('org-chart-modern-styles');
        if (existing === null || existing === void 0 ? void 0 : existing.parentNode)
            existing.parentNode.removeChild(existing);
        const rules = [];
        if (this.properties.hideWhiteSpace !== false) {
            rules.push('.CanvasZone > :first-child, .CanvasZone, .CanvasComponent, .SPCanvas-canvas, .CanvasZoneContainer {', '  max-width: 100% !important; width: 100% !important; padding: 0px !important; margin: 0px !important;', '}', '.CanvasSection-xl4 > :first-child { margin: 0px !important; padding: 0px !important; }', '[data-automation-id="CanvasControl"] { margin: 0px !important; padding: 0px !important; }');
        }
        if (this.properties.hideMenus !== false) {
            rules.push('#spLeftNav, #sp-appBar { display: none !important; }');
        }
        if (this.properties.hideTopNav !== false) {
            rules.push('[data-automationid="SiteHeader"] { display: none !important; }');
        }
        if (rules.length > 0) {
            const style = document.createElement('style');
            style.id = 'org-chart-modern-styles';
            style.setAttribute('type', 'text/css');
            style.textContent = rules.join('\n');
            document.head.appendChild(style);
        }
        // Hide comments wrapper if present
        try {
            const commentsWrapper = document.getElementById('CommentsWrapper');
            if (commentsWrapper === null || commentsWrapper === void 0 ? void 0 : commentsWrapper.parentElement) {
                commentsWrapper.parentElement.style.display = 'none';
            }
        }
        catch ( /* ignore */_a) { /* ignore */ }
    }
    onInit() {
        return __awaiter(this, void 0, void 0, function* () {
            this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
            this._themeVariant = this._themeProvider.tryGetTheme();
            this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
            if (!this.properties.list)
                this.properties.list = '';
            if (!this.properties.titleField)
                this.properties.titleField = 'Title';
            if (!this.properties.unitField)
                this.properties.unitField = 'Unit';
            if (!this.properties.officeField)
                this.properties.officeField = 'Office';
            if (!this.properties.rankField)
                this.properties.rankField = 'Rank';
            if (!this.properties.nameField)
                this.properties.nameField = 'Name';
            if (!this.properties.billetField)
                this.properties.billetField = 'Billet';
            if (!this.properties.ownerField)
                this.properties.ownerField = 'Owner';
            if (!this.properties.specialtyCodeField)
                this.properties.specialtyCodeField = 'Code';
            if (this.properties.officeSortPriority === undefined)
                this.properties.officeSortPriority = 'CC,CD,DS,DO,SEL,CCF';
            if (this.properties.unitSortPriority === undefined)
                this.properties.unitSortPriority = '';
            if (this.properties.leadershipOffices === undefined)
                this.properties.leadershipOffices = 'CC,CD,DS,DO,SEL,CCF';
            if (!this.properties.panelColumns)
                this.properties.panelColumns = '';
            if (this.properties.specialtyCodeColors === undefined)
                this.properties.specialtyCodeColors = '';
            if (!this.properties.lineColor)
                this.properties.lineColor = 'rgba(170,170,170,1)';
            if (!this.properties.borderColor)
                this.properties.borderColor = 'rgba(68,68,68,1)';
            if (!this.properties.selectedGradientStart)
                this.properties.selectedGradientStart = 'rgba(247,107,28,1)';
            if (!this.properties.selectedGradientEnd)
                this.properties.selectedGradientEnd = 'rgba(250,217,97,1)';
            if (!this.properties.selectedBorderColor)
                this.properties.selectedBorderColor = 'rgba(230,81,0,1)';
            if (!this.properties.selectedFontColor)
                this.properties.selectedFontColor = 'rgba(255,255,255,1)';
            if (!this.properties.webpartTitle)
                this.properties.webpartTitle = strings.WebPartTitle;
            if (this.properties.hideWhiteSpace === undefined)
                this.properties.hideWhiteSpace = true;
            if (this.properties.hideMenus === undefined)
                this.properties.hideMenus = true;
            if (this.properties.hideTopNav === undefined)
                this.properties.hideTopNav = true;
        });
    }
    onDispose() {
        if (this._themeProvider) {
            this._themeProvider.themeChangedEvent.remove(this, this._handleThemeChangedEvent);
        }
        ReactDom.unmountComponentAtNode(this.domElement);
        // Clean up injected styles
        const modernStyle = document.getElementById('org-chart-modern-styles');
        if (modernStyle === null || modernStyle === void 0 ? void 0 : modernStyle.parentNode)
            modernStyle.parentNode.removeChild(modernStyle);
    }
    get dataVersion() {
        return Version.parse('1.0');
    }
    _rgbaToHex(rgba) {
        const match = rgba.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
        if (!match)
            return '#000000';
        const r = parseInt(match[1]).toString(16).padStart(2, '0');
        const g = parseInt(match[2]).toString(16).padStart(2, '0');
        const b = parseInt(match[3]).toString(16).padStart(2, '0');
        return `#${r}${g}${b}`;
    }
    _renderColorField(container, propKey, label) {
        const current = this.properties[propKey] || 'rgba(0,0,0,1)';
        let hexVal;
        if (/^#[0-9A-Fa-f]{6}$/.test(current)) {
            hexVal = current;
        }
        else {
            hexVal = this._rgbaToHex(current);
        }
        container.innerHTML = '';
        const wrapper = document.createElement('div');
        wrapper.style.cssText = 'margin-bottom: 12px;';
        const lbl = document.createElement('label');
        lbl.textContent = label;
        lbl.style.cssText = 'font-size: 14px; font-weight: 600; display: block; margin-bottom: 6px;';
        const row = document.createElement('div');
        row.style.cssText = 'display: flex; align-items: center; gap: 8px;';
        const colorInput = document.createElement('input');
        colorInput.type = 'color';
        colorInput.value = hexVal;
        colorInput.style.cssText = 'width: 48px; height: 32px; cursor: pointer; border: 1px solid #ccc; border-radius: 4px; padding: 2px; flex-shrink: 0;';
        const textInput = document.createElement('input');
        textInput.type = 'text';
        textInput.value = current;
        textInput.placeholder = 'rgba(0,0,0,1)';
        textInput.style.cssText = 'flex: 1; padding: 4px 8px; border: 1px solid #ccc; border-radius: 2px; font-size: 14px;';
        colorInput.addEventListener('input', () => {
            const hex = colorInput.value;
            const rv = parseInt(hex.slice(1, 3), 16);
            const gv = parseInt(hex.slice(3, 5), 16);
            const bv = parseInt(hex.slice(5, 7), 16);
            const rgba = `rgba(${rv},${gv},${bv},1)`;
            textInput.value = rgba;
            this.properties[propKey] = rgba;
            this.render();
        });
        textInput.addEventListener('blur', () => {
            const val = textInput.value.trim();
            if (/^#[0-9A-Fa-f]{6}$/.test(val) || /^rgba?\(\s*\d+\s*,\s*\d+\s*,\s*\d+/.test(val)) {
                if (/^#/.test(val)) {
                    colorInput.value = val;
                }
                else {
                    colorInput.value = this._rgbaToHex(val);
                }
                this.properties[propKey] = val;
                this.render();
            }
            else {
                textInput.value = this.properties[propKey] || 'rgba(0,0,0,1)';
            }
        });
        row.appendChild(colorInput);
        row.appendChild(textInput);
        wrapper.appendChild(lbl);
        wrapper.appendChild(row);
        container.appendChild(wrapper);
    }
    _colorField(propKey, label) {
        return {
            type: PropertyPaneFieldType.Custom,
            targetProperty: propKey,
            properties: {
                key: `${propKey}Picker`,
                onRender: (elem) => this._renderColorField(elem, propKey, label),
                onDispose: () => undefined,
            },
        };
    }
    getPropertyPaneConfiguration() {
        return {
            pages: [
                {
                    header: { description: strings.WebPartTitle },
                    groups: [
                        {
                            groupName: strings.WebPartSettingsGroup,
                            groupFields: [
                                PropertyPaneTextField('webpartTitle', { label: strings.WebPartTitleLabel }),
                                // Spacer to prevent toggle overlapping text field
                                {
                                    type: PropertyPaneFieldType.Custom,
                                    targetProperty: '_spacer',
                                    properties: {
                                        key: 'layoutSpacer',
                                        onRender: (elem) => { elem.style.height = '8px'; },
                                        onDispose: () => undefined,
                                    },
                                },
                                PropertyPaneToggle('hideWhiteSpace', {
                                    label: strings.HideWhiteSpaceLabel,
                                    onText: strings.ToggleOnText,
                                    offText: strings.ToggleOffText,
                                }),
                                PropertyPaneToggle('hideMenus', {
                                    label: strings.HideMenusLabel,
                                    onText: strings.ToggleOnText,
                                    offText: strings.ToggleOffText,
                                }),
                                PropertyPaneToggle('hideTopNav', {
                                    label: strings.HideTopNavLabel,
                                    onText: strings.ToggleOnText,
                                    offText: strings.ToggleOffText,
                                }),
                            ],
                        },
                        {
                            groupName: strings.DataSourceGroup,
                            groupFields: [
                                PropertyPaneTextField('siteUrl', {
                                    label: strings.SiteUrlLabel,
                                    description: strings.SiteUrlDesc,
                                }),
                                PropertyPaneTextField('list', {
                                    label: strings.ListNameLabel,
                                    description: strings.ListNameDesc,
                                }),
                                PropertyPaneTextField('unitFilterLock', {
                                    label: strings.UnitFilterLockLabel,
                                    description: strings.UnitFilterLockDesc,
                                }),
                                PropertyPaneTextField('titleField', {
                                    label: strings.TitleFieldLabel,
                                    description: strings.TitleFieldDesc,
                                }),
                                PropertyPaneTextField('unitField', {
                                    label: strings.UnitFieldLabel,
                                    description: strings.UnitFieldDesc,
                                }),
                                PropertyPaneTextField('officeField', {
                                    label: strings.OfficeFieldLabel,
                                    description: strings.OfficeFieldDesc,
                                }),
                                PropertyPaneTextField('rankField', {
                                    label: strings.RankFieldLabel,
                                    description: strings.RankFieldDesc,
                                }),
                                PropertyPaneTextField('nameField', {
                                    label: strings.NameFieldLabel,
                                    description: strings.NameFieldDesc,
                                }),
                                PropertyPaneTextField('billetField', {
                                    label: strings.BilletFieldLabel,
                                    description: strings.BilletFieldDesc,
                                }),
                                PropertyPaneTextField('ownerField', {
                                    label: strings.OwnerFieldLabel,
                                    description: strings.OwnerFieldDesc,
                                }),
                                PropertyPaneTextField('specialtyCodeField', {
                                    label: strings.SpecialtyCodeFieldLabel,
                                    description: strings.SpecialtyCodeFieldDesc,
                                }),
                            ],
                        },
                        {
                            groupName: strings.DisplaySettingsGroup,
                            groupFields: [
                                this._colorField('lineColor', strings.LineColorLabel),
                                this._colorField('borderColor', strings.CardBorderColorLabel),
                                this._colorField('selectedGradientStart', strings.SelectedGradientStartLabel),
                                this._colorField('selectedGradientEnd', strings.SelectedGradientEndLabel),
                                this._colorField('selectedBorderColor', strings.SelectedBorderColorLabel),
                                this._colorField('selectedFontColor', strings.SelectedFontColorLabel),
                                this._colorField('addPositionBtnColor', strings.AddPositionBtnColorLabel),
                                this._colorField('saveImageBtnColor', strings.SaveImageBtnColorLabel),
                                this._colorField('deleteBtnColor', strings.DeleteBtnColorLabel),
                                this._colorField('cancelBtnColor', strings.CancelBtnColorLabel),
                                this._colorField('saveBtnColor', strings.SaveBtnColorLabel),
                                this._colorField('addFormBtnColor', strings.AddFormBtnColorLabel),
                                this._colorField('moveBtnColor', strings.MoveBtnColorLabel),
                                this._colorField('okBtnColor', strings.OkBtnColorLabel),
                                PropertyPaneTextField('specialtyCodeColors', {
                                    label: strings.SpecialtyCodeColorsLabel,
                                    description: strings.SpecialtyCodeColorsDesc,
                                    multiline: true,
                                    rows: 6,
                                }),
                            ],
                        },
                        {
                            groupName: strings.CardGroup,
                            groupFields: [
                                PropertyPaneTextField('cardLines', {
                                    label: strings.CardLinesLabel,
                                    description: strings.CardLinesDesc,
                                    multiline: true,
                                }),
                            ],
                        },
                        {
                            groupName: strings.DetailPanelGroup,
                            groupFields: [
                                PropertyPaneTextField('panelColumns', {
                                    label: strings.PanelColumnsLabel,
                                    description: strings.PanelColumnsDesc,
                                    multiline: true,
                                }),
                            ],
                        },
                        {
                            groupName: strings.SortGroup,
                            groupFields: [
                                PropertyPaneTextField('unitSortPriority', {
                                    label: strings.UnitSortPriorityLabel,
                                    description: strings.UnitSortPriorityDesc,
                                }),
                                PropertyPaneTextField('officeSortPriority', {
                                    label: strings.OfficeSortPriorityLabel,
                                    description: strings.OfficeSortPriorityDesc,
                                }),
                                PropertyPaneTextField('leadershipOffices', {
                                    label: strings.LeadershipOfficesLabel,
                                    description: strings.LeadershipOfficesDesc,
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    }
}
//# sourceMappingURL=OrganizationChartWebPart.js.map
/*!
 * Modern SP Military Org Chart by TehGoodLivin
 * Copyright (c) 2026 Austin Livengood <https://github.com/TehGoodLivin/>
 */
import * as React from 'react';
import type { IOrganizationChartProps } from './IOrganizationChartProps';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/security';
export interface IOrgNode {
    listItemId: number;
    billet: string;
    owner: string;
    title: string;
    unit: string;
    office: string;
    rank: string;
    name: string;
    specialtyCode: string;
    details: Record<string, string>;
}
interface IAddForm {
    title: string;
    unit: string;
    office: string;
    rank: string;
    name: string;
    billet: string;
    owner: string;
    specialtyCode: string;
    extras: Record<string, string>;
}
interface IOrganizationChartState {
    nodes: IOrgNode[];
    loading: boolean;
    configError: boolean;
    selectedBillet: string | undefined;
    panelPosition: {
        top: number;
        left: number;
    } | undefined;
    searchQuery: string;
    unitFilter: string;
    officeFilter: string;
    canEdit: boolean;
    specialtyCodeExists: boolean;
    showAddModal: boolean;
    addForm: IAddForm;
    addSaving: boolean;
    addError: string;
    confirmRemoveBillet: string | undefined;
    removeSaving: boolean;
    dragBillet: string | undefined;
    dragOverBillet: string | undefined;
    confirmReparent: {
        childBillet: string;
        newParentBillet: string;
    } | undefined;
    reparentSaving: boolean;
    reparentError: string;
    showEditModal: boolean;
    editBillet: string | undefined;
    editForm: IAddForm;
    editSaving: boolean;
    editError: string;
    orphanBillets: string[];
    invalidPanelCols: string[];
}
export default class OrganizationChart extends React.Component<IOrganizationChartProps, IOrganizationChartState> {
    private _cachedSiteUrl;
    private _sp;
    /** Returns a PnP SP instance pointing to the configured site (or current site if blank). */
    private get sp();
    private _chartWrapperRef;
    private _chartInnerRef;
    /** Read a semantic color from the current SharePoint theme, falling back to a default. */
    private _tc;
    /** Read a palette color from the current SharePoint theme, falling back to a default. */
    private _tp;
    private _isDragScrolling;
    private _dragScrollStart;
    constructor(props: IOrganizationChartProps);
    componentDidMount(): Promise<void>;
    componentDidUpdate(prev: IOrganizationChartProps): Promise<void>;
    private _fetchData;
    private _isDescendant;
    private _nodeOrDescendantMatchesSearch;
    private _nodeOrDescendantMatchesUnit;
    private _nodeOrDescendantMatchesOffice;
    private _countDirectReports;
    private _countSubordinates;
    private _onDrop;
    private _reparentMember;
    private _addMember;
    private _addMemberAndAnother;
    private _removeMember;
    private _onSearchChange;
    private _onUnitChange;
    private _onOfficeChange;
    private _onNodeClick;
    private _cropCanvasBottom;
    private _doExport;
    private _exportToImage;
    private _onDragScrollStart;
    private _onDragScrollMove;
    private _onDragScrollEnd;
    private _parseSpecialtyColors;
    private _getColorForCode;
    private _buildTree;
    private _renderPanel;
    private _renderAddModal;
    private _renderConfirmRemove;
    private _startEdit;
    private _saveEdit;
    private _renderEditModal;
    private _renderConfirmReparent;
    render(): React.ReactElement;
}
export {};
//# sourceMappingURL=OrganizationChart.d.ts.map
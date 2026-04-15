/*!
 * Modern SP Military Org Chart by TehGoodLivin
 * Copyright (c) 2026 Austin Livengood <https://github.com/TehGoodLivin/>
 */
import * as React from 'react';
import type { IOrganizationChartProps } from './IOrganizationChartProps';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import html2canvas from 'html2canvas';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/security';
import * as strings from 'OrganizationChartWebPartStrings';

// ─── Data model ──────────────────────────────────────────────────────────────

export interface IOrgNode {
  listItemId: number;
  billet: string;     // node ID
  owner: string;      // parent billet (empty = root)
  title: string;
  unit: string;
  office: string;
  rank: string;
  name: string;
  specialtyCode: string;
  details: Record<string, string>;
}

// ─── Specialty code color entry ──────────────────────────────────────────────

interface ISpecialtyCodeColor {
  Code: string;
  GradientStart: string;
  GradientEnd: string;
  FontColor: string;
}

// ─── State ───────────────────────────────────────────────────────────────────

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
  panelPosition: { top: number; left: number } | undefined;
  searchQuery: string;
  unitFilter: string;
  officeFilter: string;
  canEdit: boolean;
  specialtyCodeExists: boolean;

  // Add member
  showAddModal: boolean;
  addForm: IAddForm;
  addSaving: boolean;
  addError: string;

  // Remove member
  confirmRemoveBillet: string | undefined;
  removeSaving: boolean;

  // Drag-and-drop reparenting
  dragBillet: string | undefined;
  dragOverBillet: string | undefined;   // '__root__' = hovering over root drop zone
  confirmReparent: { childBillet: string; newParentBillet: string } | undefined;
  reparentSaving: boolean;
  reparentError: string;

  // Edit member
  showEditModal: boolean;
  editBillet: string | undefined;   // billet of the node being edited
  editForm: IAddForm;
  editSaving: boolean;
  editError: string;

  // Data warnings
  orphanBillets: string[];      // billets whose owner doesn't exist in list
  invalidPanelCols: string[];   // panel column names not found in list

}

// ─── Constants / style helpers ────────────────────────────────────────────────

const CONNECTOR_SIZE = 20;

const BLANK_ADD_FORM: IAddForm = {
  title: '', unit: '', office: '', rank: '', name: '', billet: '', owner: '', specialtyCode: '', extras: {},
};

const DEFAULT_OTHER_COLOR: ISpecialtyCodeColor = {
  Code: 'Other',
  GradientStart: 'rgba(255,255,255,1)',
  GradientEnd: 'rgba(230,230,240,1)',
  FontColor: 'rgba(0,0,0,1)',
};

/** Resolve a SharePoint column name to its value on a node, checking core mapped fields first, then details. */
function resolveCardColumn(node: IOrgNode, colName: string, props: IOrganizationChartProps): string {
  const col = colName.trim();
  // Map SharePoint field names to core node properties
  if (col === props.billetField) return node.billet;
  if (col === props.rankField) return node.rank;
  if (col === props.nameField) return node.name;
  if (col === props.titleField) return node.title;
  if (col === props.unitField) return node.unit;
  if (col === props.officeField) return node.office;
  if (col === props.specialtyCodeField) return node.specialtyCode;
  // Fall back to details
  return node.details[col] || '';
}

/** Parse cardLines JSON setting. Returns array of arrays (each inner array = column names or literal separators for one line). */
function parseCardLines(json: string): string[][] {
  try {
    const parsed = JSON.parse(json || '[]');
    if (Array.isArray(parsed)) {
      return parsed
        .filter((line: unknown) => Array.isArray(line))
        .map((line: unknown[]) => (line as string[]).filter((c: unknown) => typeof c === 'string' && c !== ''));
    }
  } catch { /* ignore */ }
  return [];
}

/** Extract unique column names (non-separator tokens) from a JSON lines config string. */
function extractColumnNames(json: string): string[] {
  const lines = parseCardLines(json);
  const seen = new Set<string>();
  const cols: string[] = [];
  for (const tokens of lines) {
    for (const t of tokens) {
      if (!isLiteralSeparator(t) && !seen.has(t)) { seen.add(t); cols.push(t); }
    }
  }
  return cols;
}

/** Returns true if a token is a literal separator (no alphanumeric characters). */
function isLiteralSeparator(token: string): boolean {
  return !/[a-zA-Z0-9]/.test(token);
}

/** Build the display text for one card line from an array of column names and literal separators. */
function buildCardLineText(tokens: string[], node: IOrgNode, props: IOrganizationChartProps): string {
  // Resolve each token: columns become their value, separators stay literal
  const parts: { text: string; isSep: boolean }[] = tokens.map(t => {
    if (isLiteralSeparator(t)) return { text: t, isSep: true };
    const val = resolveCardColumn(node, t, props);
    return { text: val, isSep: false };
  });
  // Drop columns that resolved empty, then trim leading/trailing separators
  const filtered = parts.filter(p => p.isSep || p.text);
  // Trim leading separators
  while (filtered.length > 0 && filtered[0].isSep) filtered.shift();
  // Trim trailing separators
  while (filtered.length > 0 && filtered[filtered.length - 1].isSep) filtered.pop();
  return filtered.map(p => p.text).join('');
}

/** Build display text for a detail panel line from tokens, resolving values from node.details. */
function buildPanelLineText(tokens: string[], node: IOrgNode): string {
  const parts: { text: string; isSep: boolean }[] = tokens.map(t => {
    if (isLiteralSeparator(t)) return { text: t, isSep: true };
    const val = node.details[t] || '';
    return { text: val, isSep: false };
  });
  const filtered = parts.filter(p => p.isSep || p.text);
  while (filtered.length > 0 && filtered[0].isSep) filtered.shift();
  while (filtered.length > 0 && filtered[filtered.length - 1].isSep) filtered.pop();
  return filtered.map(p => p.text).join('');
}

function buildCardStyle(
  gradient: string,
  borderColor: string,
  fontColor: string,
  vacant: boolean,
  dragOver: boolean,
  canEdit: boolean
): React.CSSProperties {
  return {
    background: vacant ? 'none' : gradient,
    backgroundColor: vacant ? 'rgba(180,180,180,0.12)' : undefined,
    color: vacant ? '#888' : fontColor,
    padding: canEdit ? '10px 36px 10px 36px' : '10px 16px',
    border: `2px ${vacant ? 'dashed' : 'solid'} ${dragOver ? '#0078d4' : borderColor}`,
    borderRadius: '10px',
    boxShadow: dragOver
      ? '0 0 0 3px rgba(0,120,212,0.35), 0 4px 12px rgba(0,0,0,0.15)'
      : '0 4px 12px rgba(0,0,0,0.15)',
    cursor: 'pointer',
    minWidth: '160px',
    maxWidth: '260px',
    textAlign: 'center',
    fontFamily: "'Segoe UI', sans-serif",
    userSelect: 'none',
    flexShrink: 0,
    position: 'relative',
    opacity: vacant ? 0.72 : 1,
    transition: 'box-shadow 0.12s, border-color 0.12s',
  };
}


// ─── Format helpers ──────────────────────────────────────────────────────────

function formatPhone(val: string): string {
  const digits = val.replace(/\D/g, '');
  if (digits.length === 10) return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
  if (digits.length === 7) return `${digits.slice(0, 3)}-${digits.slice(3)}`;
  return val;
}

function formatPanelValue(colName: string, val: string): JSX.Element {
  const lower = colName.toLowerCase();
  if (lower === 'email') {
    return <a href={`mailto:${val}`} style={{ color: '#0078d4', textDecoration: 'none' }}>{val}</a>;
  }
  if (lower === 'commercial' || lower === 'dsn') {
    return <>{formatPhone(val)}</>;
  }
  return <>{val}</>;
}

// ─── Component ───────────────────────────────────────────────────────────────

export default class OrganizationChart extends React.Component<IOrganizationChartProps, IOrganizationChartState> {
  private _cachedSiteUrl: string | undefined;
  private _sp = spfi().using(SPFx(this.props.context));

  /** Returns a PnP SP instance pointing to the configured site (or current site if blank). */
  private get sp(): SPFI {
    const url = (this.props.siteUrl || '').trim() || undefined;
    if (url !== this._cachedSiteUrl) {
      this._cachedSiteUrl = url;
      this._sp = url ? spfi(url).using(SPFx(this.props.context)) : spfi().using(SPFx(this.props.context));
    }
    return this._sp;
  }
  private _chartWrapperRef = React.createRef<HTMLDivElement>();
  private _chartInnerRef = React.createRef<HTMLDivElement>();

  /** Read a semantic color from the current SharePoint theme, falling back to a default. */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _tc(key: string, fallback: string): string { return (this.props.themeVariant?.semanticColors as any)?.[key] ?? fallback; }
  /** Read a palette color from the current SharePoint theme, falling back to a default. */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _tp(key: string, fallback: string): string { return (this.props.themeVariant?.palette as any)?.[key] ?? fallback; }

  // Drag-to-scroll (instance vars to avoid re-renders)
  private _isDragScrolling = false;
  private _dragScrollStart: { x: number; y: number; scrollLeft: number; scrollTop: number } | null = null;

  constructor(props: IOrganizationChartProps) {
    super(props);
    this.state = {
      nodes: [],
      loading: false,
      configError: false,
      selectedBillet: undefined,
      panelPosition: undefined,
      searchQuery: '',
      unitFilter: '',
      officeFilter: '',
      canEdit: false,
      specialtyCodeExists: false,
      showAddModal: false,
      addForm: { ...BLANK_ADD_FORM },
      addSaving: false,
      addError: '',
      confirmRemoveBillet: undefined,
      removeSaving: false,
      dragBillet: undefined,
      dragOverBillet: undefined,
      confirmReparent: undefined,
      reparentSaving: false,
      reparentError: '',
      showEditModal: false,
      editBillet: undefined,
      editForm: { ...BLANK_ADD_FORM },
      editSaving: false,
      editError: '',
      orphanBillets: [],
      invalidPanelCols: [],
    };
  }

  // ── Lifecycle ─────────────────────────────────────────────────────────────

  public async componentDidMount(): Promise<void> {
    await this._fetchData();
  }

  public async componentDidUpdate(prev: IOrganizationChartProps): Promise<void> {
    const keys: (keyof IOrganizationChartProps)[] = [
      'siteUrl', 'list', 'unitFilterLock', 'titleField', 'unitField', 'officeField',
      'rankField', 'nameField', 'billetField', 'ownerField', 'panelColumns', 'specialtyCodeField', 'cardLines',
    ];
    if (keys.some(k => prev[k] !== this.props[k])) {
      await this._fetchData();
    }
  }

  // ── Data loading ──────────────────────────────────────────────────────────

  private async _fetchData(): Promise<void> {
    const {
      list, titleField, unitField, officeField, rankField,
      nameField, billetField, ownerField, panelColumns, specialtyCodeField, cardLines,
    } = this.props;

    // Config validation: missing list or field mappings
    const requiredFields = [titleField, unitField, officeField, rankField, nameField, billetField, ownerField, specialtyCodeField];
    if (!list || requiredFields.some(f => !f || !f.trim())) {
      this.setState({ configError: true, loading: false, nodes: [] });
      return;
    }

    this.setState({ loading: true, configError: false });

    try {
      // Validate list fields exist
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const listFields: { InternalName: string }[] = await this.sp.web.lists
        .getByTitle(list)
        .fields
        .select('InternalName')();

      const fieldNames = new Set(listFields.map(f => f.InternalName));

      for (const f of requiredFields) {
        if (!fieldNames.has(f.trim())) {
          this.setState({ configError: true, loading: false, nodes: [] });
          return;
        }
      }

      // Extract column names from panelColumns and cardLines JSON arrays
      const panelCols = extractColumnNames(panelColumns);
      const cardLineCols = extractColumnNames(cardLines);
      const reqSet = new Set(requiredFields.map(f => f.trim()));
      // Combine all extra columns, deduped, that exist in the list
      const allExtraCols = [...new Set([...panelCols, ...cardLineCols])];
      const validExtraCols = allExtraCols.filter(c => fieldNames.has(c) && !reqSet.has(c));
      const invalidPanelCols = panelCols.filter(c => !fieldNames.has(c));

      const selectFields = ['Id', ...requiredFields.map(f => f.trim()), ...validExtraCols];

      // ── Paginated fetch: handles lists >5 000 items ──────────────────────
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let allItems: Record<string, any>[] = [];
      const pageSize = 2000;
      let skip = 0;
      // eslint-disable-next-line no-constant-condition
      while (true) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const page: Record<string, any>[] = await this.sp.web.lists
          .getByTitle(list)
          .items
          .select(...selectFields)
          .top(pageSize)
          .skip(skip)();
        allItems = allItems.concat(page);
        if (page.length < pageSize) break;
        skip += pageSize;
      }

      // Build nodes
      const nodes: IOrgNode[] = allItems.map(item => ({
        listItemId: item.Id as number,
        billet: String(item[billetField] ?? '').trim(),
        owner: String(item[ownerField] ?? '').trim(),
        title: String(item[titleField] ?? '').trim(),
        unit: String(item[unitField] ?? '').trim(),
        office: String(item[officeField] ?? '').trim(),
        rank: String(item[rankField] ?? '').trim(),
        name: String(item[nameField] ?? '').trim(),
        specialtyCode: String(item[specialtyCodeField] ?? '').trim(),
        details: validExtraCols.reduce<Record<string, string>>((acc, col) => {
          const val = item[col];
          if (val !== undefined && val !== null && val !== '') acc[col] = String(val);
          return acc;
        }, {}),
      }));

      // ── Orphan detection ─────────────────────────────────────────────────
      const billetSet = new Set(nodes.map(n => n.billet));
      const orphanBillets = nodes
        .filter(n => n.owner !== '' && !billetSet.has(n.owner))
        .map(n => n.billet);

      // Check edit permissions
      let canEdit = false;
      try {
        const perms = await this.sp.web.lists.getByTitle(list).getCurrentUserEffectivePermissions();
        // EditListItems bit = 0x4 in Low
        canEdit = ((perms.Low >>> 0) & 0x4) !== 0;
      } catch {
        canEdit = false;
      }

      this.setState({
        nodes, configError: false, loading: false,
        canEdit, specialtyCodeExists: true, orphanBillets, invalidPanelCols,
      });
    } catch {
      this.setState({ configError: true, loading: false, nodes: [] });
    }
  }

  // ── Hierarchy helpers ─────────────────────────────────────────────────────

  private _isDescendant(ancestorBillet: string, potentialBillet: string, nodes: IOrgNode[]): boolean {
    if (potentialBillet === ancestorBillet) return true;
    const children = nodes.filter(n => n.owner === ancestorBillet);
    return children.some(c => this._isDescendant(c.billet, potentialBillet, nodes));
  }

  private _nodeOrDescendantMatchesSearch(billet: string, nodes: IOrgNode[], search: string): boolean {
    const node = nodes.find(n => n.billet === billet);
    if (!node) return false;
    if (node.name.toLowerCase().includes(search)) return true;
    return nodes
      .filter(n => n.owner === billet)
      .some(c => this._nodeOrDescendantMatchesSearch(c.billet, nodes, search));
  }

  private _nodeOrDescendantMatchesUnit(billet: string, nodes: IOrgNode[], unit: string): boolean {
    const node = nodes.find(n => n.billet === billet);
    if (!node) return false;
    if (node.unit.toLowerCase() === unit.toLowerCase()) return true;
    return nodes
      .filter(n => n.owner === billet)
      .some(c => this._nodeOrDescendantMatchesUnit(c.billet, nodes, unit));
  }

  private _nodeOrDescendantMatchesOffice(billet: string, nodes: IOrgNode[], office: string): boolean {
    const node = nodes.find(n => n.billet === billet);
    if (!node) return false;
    if (node.office.toLowerCase() === office.toLowerCase()) return true;
    return nodes
      .filter(n => n.owner === billet)
      .some(c => this._nodeOrDescendantMatchesOffice(c.billet, nodes, office));
  }

  private _countDirectReports(billet: string, nodes: IOrgNode[]): number {
    return nodes.filter(n => n.owner === billet).length;
  }

  private _countSubordinates(billet: string, nodes: IOrgNode[], visited: Set<string> = new Set()): number {
    if (visited.has(billet)) return 0;
    visited.add(billet);
    const children = nodes.filter(n => n.owner === billet);
    return children.reduce((sum, c) => sum + 1 + this._countSubordinates(c.billet, nodes, visited), 0);
  }

  // ── Drag-and-drop ────────────────────────────────────────────────────────

  private _onDrop = (targetBillet: string): void => {
    const { dragBillet, nodes } = this.state;
    if (!dragBillet) return;
    if (dragBillet === targetBillet) return;

    // Prevent dropping onto a descendant (would create a cycle)
    if (this._isDescendant(dragBillet, targetBillet, nodes)) return;

    // '__root__' is the sentinel for the root drop zone → newParentBillet = ''
    const newParentBillet = targetBillet === '__root__' ? '' : targetBillet;

    const dragNode = nodes.find(n => n.billet === dragBillet);

    // Prevent dropping into a different unit — show message
    if (dragNode && targetBillet !== '__root__') {
      const targetNode = nodes.find(n => n.billet === targetBillet);
      if (targetNode && dragNode.unit !== targetNode.unit) {
        this.setState({
          dragBillet: undefined,
          dragOverBillet: undefined,
          confirmReparent: { childBillet: dragBillet, newParentBillet },
          reparentError: strings.MoveBlockedCrossUnit,
        });
        return;
      }
    }

    // Don't re-confirm if nothing would change
    if (dragNode && dragNode.owner === newParentBillet) return;

    this.setState({
      dragBillet: undefined,
      dragOverBillet: undefined,
      confirmReparent: { childBillet: dragBillet, newParentBillet },
      reparentError: '',
    });
  };

  private _reparentMember = async (): Promise<void> => {
    const { confirmReparent, nodes } = this.state;
    const { list, ownerField } = this.props;
    if (!confirmReparent) return;

    const { childBillet, newParentBillet } = confirmReparent;
    const node = nodes.find(n => n.billet === childBillet);
    if (!node) return;

    this.setState({ reparentSaving: true });
    try {
      await this.sp.web.lists.getByTitle(list).items.getById(node.listItemId).update({
        [ownerField]: newParentBillet,
      });
      this.setState({ confirmReparent: undefined, reparentSaving: false });
      await this._fetchData();
    } catch {
      this.setState({ reparentSaving: false });
    }
  };

  // ── Add / Remove actions ──────────────────────────────────────────────────

  private _addMember = async (): Promise<void> => {
    const { list, titleField, unitField, officeField, rankField, nameField, billetField, ownerField, specialtyCodeField, panelColumns } = this.props;
    const { addForm } = this.state;

    // Required fields validation
    const missing: string[] = [];
    if (!addForm.billet.trim()) missing.push(strings.BilletInputLabel);
    if (!addForm.title.trim()) missing.push(strings.TitleInputLabel);
    if (!addForm.office.trim()) missing.push(strings.OfficeInputLabel);
    if (!addForm.unit.trim()) missing.push(strings.UnitInputLabel);
    if (missing.length > 0) {
      this.setState({ addError: `${strings.RequiredPrefix} ${missing.join(', ')}` });
      return;
    }

    this.setState({ addSaving: true, addError: '' });
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        [titleField]: addForm.title,
        [unitField]: addForm.unit,
        [officeField]: addForm.office,
        [rankField]: addForm.rank,
        [nameField]: addForm.name,
        [billetField]: addForm.billet.trim(),
        [ownerField]: addForm.owner.trim(),
        [specialtyCodeField]: addForm.specialtyCode,
      };
      // Include panel column extras
      const validPanelCols = extractColumnNames(panelColumns)
        .filter(c => !this.state.invalidPanelCols.includes(c));
      for (const col of validPanelCols) {
        if (addForm.extras[col] !== undefined) payload[col] = addForm.extras[col];
      }

      await this.sp.web.lists.getByTitle(list).items.add(payload);
      this.setState({ showAddModal: false, addForm: { ...BLANK_ADD_FORM }, addSaving: false, addError: '' });
      await this._fetchData();
    } catch {
      this.setState({ addSaving: false, addError: strings.AddFailedError });
    }
  };

  private _addMemberAndAnother = async (): Promise<void> => {
    const { list, titleField, unitField, officeField, rankField, nameField, billetField, ownerField, specialtyCodeField, panelColumns } = this.props;
    const { addForm } = this.state;

    const missing: string[] = [];
    if (!addForm.billet.trim()) missing.push(strings.BilletInputLabel);
    if (!addForm.title.trim()) missing.push(strings.TitleInputLabel);
    if (!addForm.office.trim()) missing.push(strings.OfficeInputLabel);
    if (!addForm.unit.trim()) missing.push(strings.UnitInputLabel);
    if (missing.length > 0) {
      this.setState({ addError: `${strings.RequiredPrefix} ${missing.join(', ')}` });
      return;
    }

    this.setState({ addSaving: true, addError: '' });
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        [titleField]: addForm.title,
        [unitField]: addForm.unit,
        [officeField]: addForm.office,
        [rankField]: addForm.rank,
        [nameField]: addForm.name,
        [billetField]: addForm.billet.trim(),
        [ownerField]: addForm.owner.trim(),
        [specialtyCodeField]: addForm.specialtyCode,
      };
      const validPanelCols = extractColumnNames(panelColumns)
        .filter(c => !this.state.invalidPanelCols.includes(c));
      for (const col of validPanelCols) {
        if (addForm.extras[col] !== undefined) payload[col] = addForm.extras[col];
      }

      await this.sp.web.lists.getByTitle(list).items.add(payload);
      // Keep modal open, reset form for next entry
      this.setState({ addForm: { ...BLANK_ADD_FORM }, addSaving: false, addError: '' });
      await this._fetchData();
    } catch {
      this.setState({ addSaving: false, addError: strings.AddFailedError });
    }
  };

  private _removeMember = async (billet: string): Promise<void> => {
    const { list, ownerField } = this.props;
    const { nodes } = this.state;
    const node = nodes.find(n => n.billet === billet);
    if (!node) return;

    this.setState({ removeSaving: true });
    try {
      const children = nodes.filter(n => n.owner === billet);
      await Promise.all(
        children.map(child =>
          this.sp.web.lists.getByTitle(list).items.getById(child.listItemId).update({
            [ownerField]: node.owner,
          })
        )
      );
      await this.sp.web.lists.getByTitle(list).items.getById(node.listItemId).delete();
      this.setState({
        confirmRemoveBillet: undefined,
        removeSaving: false,
        selectedBillet: this.state.selectedBillet === billet ? undefined : this.state.selectedBillet,
        panelPosition: this.state.selectedBillet === billet ? undefined : this.state.panelPosition,
      });
      await this._fetchData();
    } catch {
      this.setState({ removeSaving: false });
    }
  };

  // ── Interactions ──────────────────────────────────────────────────────────

  private _onSearchChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ searchQuery: e.target.value });
  };

  private _onUnitChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ unitFilter: e.target.value, officeFilter: '' });
  };

  private _onOfficeChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ officeFilter: e.target.value });
  };

  private _onNodeClick = (billet: string, event: React.MouseEvent<HTMLDivElement>): void => {
    if (this.state.selectedBillet === billet) {
      this.setState({ selectedBillet: undefined, panelPosition: undefined });
      return;
    }
    const nodeEl = event.currentTarget as HTMLDivElement;
    const wrapperEl = this._chartWrapperRef.current;
    const PANEL_WIDTH = 278;
    const MARGIN = 10;
    if (wrapperEl) {
      const nodeRect = nodeEl.getBoundingClientRect();
      const wrapperRect = wrapperEl.getBoundingClientRect();
      const nodeTop = nodeRect.top - wrapperRect.top;
      const nodeLeft = nodeRect.left - wrapperRect.left;
      const nodeRight = nodeRect.right - wrapperRect.left;
      const panelLeft = nodeLeft >= PANEL_WIDTH + MARGIN
        ? nodeLeft - PANEL_WIDTH - MARGIN
        : nodeRight + MARGIN;
      this.setState({
        selectedBillet: billet,
        panelPosition: { top: Math.max(0, nodeTop), left: panelLeft },
      });
    } else {
      this.setState({ selectedBillet: billet, panelPosition: { top: 0, left: 0 } });
    }
  };

  // ── Export ────────────────────────────────────────────────────────────────

  private _cropCanvasBottom(canvas: HTMLCanvasElement): HTMLCanvasElement {
    const ctx = canvas.getContext('2d');
    if (!ctx) return canvas;
    const { width, height } = canvas;
    let data: Uint8ClampedArray;
    try {
      data = ctx.getImageData(0, 0, width, height).data;
    } catch {
      return canvas; // cross-origin tainted canvas — return uncropped
    }
    let lastContentRow = 0;
    for (let y = height - 1; y >= 0; y--) {
      let hasContent = false;
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) { hasContent = true; break; }
      }
      if (hasContent) { lastContentRow = y; break; }
    }
    const croppedHeight = Math.min(lastContentRow + 40, height);
    const cropped = document.createElement('canvas');
    cropped.width = width;
    cropped.height = croppedHeight;
    cropped.getContext('2d')?.drawImage(canvas, 0, 0);
    return cropped;
  }

  private _doExport(onDone?: () => void): void {
    const el = this._chartInnerRef.current;
    if (!el) { onDone?.(); return; }

    // Reset scroll to (0,0) so html2canvas captures from the beginning of the content,
    // not from wherever the user has panned to. Restore afterward.
    const savedScrollLeft = el.scrollLeft;
    const savedScrollTop = el.scrollTop;
    el.scrollLeft = 0;
    el.scrollTop = 0;

    const shadowPad = 20; // extra padding so left/right box-shadows aren't clipped
    const fullWidth = el.scrollWidth + shadowPad * 2;
    const fullHeight = el.scrollHeight + 300;
    html2canvas(el, {
      backgroundColor: '#ffffff',
      scale: 2,
      useCORS: true,
      width: fullWidth,
      height: fullHeight,
      onclone: (_doc: Document, clonedEl: HTMLElement) => {
        clonedEl.style.overflow = 'visible';
        clonedEl.style.maxWidth = 'none';
        clonedEl.style.width = `${fullWidth}px`;
        clonedEl.style.height = `${fullHeight}px`;
        clonedEl.style.paddingTop = '20px';
        clonedEl.style.paddingBottom = '20px';
        clonedEl.style.paddingLeft = `${shadowPad}px`;
        clonedEl.style.paddingRight = `${shadowPad}px`;
        // Strip all box-shadows so no bleed artifacts appear in the exported image
        const noShadow = _doc.createElement('style');
        noShadow.textContent = '* { box-shadow: none !important; }';
        _doc.head.appendChild(noShadow);
        let parent = clonedEl.parentElement;
        while (parent) {
          parent.style.overflow = 'visible';
          parent.style.height = 'auto';
          parent.style.maxHeight = 'none';
          parent = parent.parentElement;
        }
        // Hide edit & delete buttons in exported image
        const editBtns = clonedEl.querySelectorAll('[data-export-hide]');
        editBtns.forEach(btn => {
          (btn as HTMLElement).style.display = 'none';
        });
      },
    })
      .then((canvas: HTMLCanvasElement) => {
        el.scrollLeft = savedScrollLeft;
        el.scrollTop = savedScrollTop;
        const cropped = this._cropCanvasBottom(canvas);
        const link = document.createElement('a');
        link.download = 'org-chart.png';
        link.href = cropped.toDataURL('image/png');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        onDone?.();
      })
      .catch(() => {
        el.scrollLeft = savedScrollLeft;
        el.scrollTop = savedScrollTop;
        onDone?.();
      });
  }

  private _exportToImage = (): void => {
    const { selectedBillet, panelPosition } = this.state;
    if (selectedBillet !== undefined) {
      // Close panel, export, then restore panel position
      this.setState({ selectedBillet: undefined, panelPosition: undefined }, () => {
        this._doExport(() => {
          this.setState({ selectedBillet, panelPosition });
        });
      });
    } else {
      this._doExport();
    }
  };

  // ── Drag-to-scroll ─────────────────────────────────────────────────────────

  private _onDragScrollStart = (e: React.MouseEvent<HTMLDivElement>): void => {
    const target = e.target as HTMLElement;
    if (target.closest('[data-card]') || target.closest('button') || target.closest('input') || target.closest('select')) return;
    if (e.button !== 0) return;
    const el = this._chartInnerRef.current;
    if (!el) return;
    this._isDragScrolling = true;
    this._dragScrollStart = {
      x: e.clientX,
      y: e.clientY,
      scrollLeft: el.scrollLeft,
      scrollTop: el.scrollTop,
    };
    el.style.cursor = 'grabbing';
    e.preventDefault();
  };

  private _onDragScrollMove = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (!this._isDragScrolling || !this._dragScrollStart) return;
    const el = this._chartInnerRef.current;
    if (!el) return;
    el.scrollLeft = this._dragScrollStart.scrollLeft - (e.clientX - this._dragScrollStart.x);
    el.scrollTop = this._dragScrollStart.scrollTop - (e.clientY - this._dragScrollStart.y);
  };

  private _onDragScrollEnd = (): void => {
    if (!this._isDragScrolling) return;
    this._isDragScrolling = false;
    this._dragScrollStart = null;
    const el = this._chartInnerRef.current;
    if (el) el.style.cursor = 'grab';
  };

  // ── Specialty code helpers ─────────────────────────────────────────────────

  private _parseSpecialtyColors(): ISpecialtyCodeColor[] {
    try {
      const arr: ISpecialtyCodeColor[] = JSON.parse(this.props.specialtyCodeColors || '[]');
      if (!arr.length || !arr.some(c => c.Code === 'Other')) {
        return [...arr, DEFAULT_OTHER_COLOR];
      }
      return arr;
    } catch {
      return [DEFAULT_OTHER_COLOR];
    }
  }

  private _getColorForCode(code: string, colors: ISpecialtyCodeColor[]): ISpecialtyCodeColor {
    if (code) {
      const match = colors.find(c => c.Code.toLowerCase() === code.toLowerCase() && c.Code !== 'Other');
      if (match) return match;
    }
    return colors.find(c => c.Code === 'Other') || DEFAULT_OTHER_COLOR;
  }

  // ── Tree rendering ────────────────────────────────────────────────────────

  private _buildTree(
    parentBillet: string,
    allNodes: IOrgNode[],
    visited: Set<string>,
    isTop: boolean
  ): JSX.Element[] {
    const { searchQuery, unitFilter: stateUnitFilter2, officeFilter, selectedBillet, dragBillet, dragOverBillet, canEdit: stateCanEdit } = this.state;
    const unitFilter = (this.props.unitFilterLock || '').trim() || stateUnitFilter2;
    const {
      borderColor,
      selectedGradientStart, selectedGradientEnd, selectedBorderColor, selectedFontColor,
      isPageEditMode,
    } = this.props;

    // Effective canEdit: requires list permissions AND page not in edit mode
    const canEdit = stateCanEdit && !isPageEditMode;

    const selGradient = `linear-gradient(135deg, ${selectedGradientStart} 0%, ${selectedGradientEnd} 100%)`;
    const search = (searchQuery || '').toLowerCase();

    // Parse specialty colors once
    const specialtyColors = this._parseSpecialtyColors();

    // Children of this parent — orphans treated as roots
    const billetSet = new Set(allNodes.map(n => n.billet));
    const children = allNodes.filter(n => {
      if (parentBillet === '') {
        return !n.owner || n.owner === '' || !billetSet.has(n.owner);
      }
      return n.owner === parentBillet;
    });

    // ── Sort siblings ─────────────────────────────────────────────────────
    const natSort = (x: string, y: string): number =>
      x.localeCompare(y, undefined, { numeric: true, sensitivity: 'base' });

    // Office priority (always applied)
    const officePriority = (this.props.officeSortPriority || '')
      .split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    const officeIdx = (v: string): number => {
      const i = officePriority.indexOf((v || '').toLowerCase());
      return i === -1 ? Infinity : i;
    };

    // Unit priority (only when no unit filter is active — all units shown)
    const unitPriority = (!unitFilter)
      ? (this.props.unitSortPriority || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean)
      : [];
    const unitIdx = (v: string): number => {
      if (!unitPriority.length) return Infinity;
      const i = unitPriority.indexOf((v || '').toLowerCase());
      return i === -1 ? Infinity : i;
    };

    children.sort((a, b) => {
      // 1. Unit priority (all-units view only)
      if (!unitFilter) {
        const ua = unitIdx(a.unit), ub = unitIdx(b.unit);
        if (ua !== ub) return ua - ub;
        const u = natSort(a.unit || '', b.unit || '');
        if (u !== 0) return u;
      }
      // 2. Office priority, then ABC
      const oa = officeIdx(a.office), ob = officeIdx(b.office);
      if (oa !== ob) return oa - ob;
      return natSort(a.office || '', b.office || '');
    });

    return children
      .filter(node => {
        if (visited.has(node.billet)) return false;
        if (search && !this._nodeOrDescendantMatchesSearch(node.billet, allNodes, search)) {
          const leadershipCodes = (this.props.leadershipOffices || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
          // Only keep leadership nodes if their parent is a pass-through ancestor
          // (parent has descendants matching but doesn't match search itself)
          if (!leadershipCodes.includes(node.office.toLowerCase()) || !parentBillet) return false;
          const parentNode = allNodes.find(n => n.billet === parentBillet);
          const parentMatchesDirectly = parentNode ? parentNode.name.toLowerCase().includes(search) : false;
          if (parentMatchesDirectly || !this._nodeOrDescendantMatchesSearch(parentBillet, allNodes, search)) return false;
        }
        if (unitFilter && !this._nodeOrDescendantMatchesUnit(node.billet, allNodes, unitFilter)) return false;
        if (officeFilter && !this._nodeOrDescendantMatchesOffice(node.billet, allNodes, officeFilter)) {
          const leadershipCodes = (this.props.leadershipOffices || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
          if (!leadershipCodes.includes(node.office.toLowerCase()) || !parentBillet) return false;
          const parentNode = allNodes.find(n => n.billet === parentBillet);
          if (!parentNode) return false;
          const parentMatchesOffice = parentNode.office.toLowerCase() === officeFilter.toLowerCase();
          if (parentMatchesOffice || !this._nodeOrDescendantMatchesOffice(parentBillet, allNodes, officeFilter)) return false;
        }
        return true;
      })
      .map(node => {
        const newVisited = new Set(visited);
        newVisited.add(node.billet);

        const isSelected = node.billet === selectedBillet;
        const isVacant = !node.name;
        const isDragOver = dragOverBillet === node.billet;
        const isDragging = dragBillet === node.billet;

        // Block dropping onto own descendant (cross-unit check handled in _onDrop)
        const isValidDropTarget = dragBillet
          && dragBillet !== node.billet
          && !this._isDescendant(dragBillet, node.billet, allNodes);

        const subChildren = this._buildTree(node.billet, allNodes, newVisited, false);

        const nameDisplay = isVacant
          ? strings.VacantLabel
          : [node.rank, node.name].filter(Boolean).join(' ');
        const unitOffice = [node.unit, node.office].filter(Boolean).join('/');

        // Per-node colors from specialty code
        const colorEntry = this._getColorForCode(node.specialtyCode, specialtyColors);
        const nodeGradient = `linear-gradient(135deg, ${colorEntry.GradientStart} 0%, ${colorEntry.GradientEnd} 100%)`;
        const nodeFontColor = colorEntry.FontColor;

        const activeGradient = isSelected ? selGradient : nodeGradient;
        const activeBorder = isSelected ? (selectedBorderColor || borderColor) : borderColor;
        const activeFontColor = isSelected ? selectedFontColor : nodeFontColor;

        return (
          <li
            key={node.billet}
            style={{
              position: 'relative',
              margin: isTop ? '0 3rem 3rem 3rem' : '0 0.75rem',
              paddingTop: isTop ? 0 : CONNECTOR_SIZE,
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              listStyle: 'none',
              opacity: isDragging ? 0.4 : 1,
              transition: 'opacity 0.15s',
            }}
          >
            {/* Vertical drop from parent bar */}
            {!isTop && (
              <div style={{
                position: 'absolute', top: 0, left: 'calc(50% - 1px)',
                width: 2, height: CONNECTOR_SIZE,
                background: this.props.lineColor,
              }} />
            )}

            {/* Card */}
            <div
              data-card="true"
              draggable={canEdit ? true : undefined}
              style={buildCardStyle(
                activeGradient, activeBorder, activeFontColor,
                isVacant, isDragOver && !!isValidDropTarget, canEdit
              )}
              onClick={(e) => this._onNodeClick(node.billet, e)}
              onDragStart={canEdit ? (e) => {
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', node.billet);
                this.setState({ dragBillet: node.billet, selectedBillet: undefined, panelPosition: undefined });
              } : undefined}
              onDragOver={canEdit ? (e) => {
                if (isValidDropTarget) {
                  e.preventDefault();
                  e.dataTransfer.dropEffect = 'move';
                  if (dragOverBillet !== node.billet) this.setState({ dragOverBillet: node.billet });
                }
              } : undefined}
              onDragLeave={canEdit ? () => {
                if (dragOverBillet === node.billet) this.setState({ dragOverBillet: undefined });
              } : undefined}
              onDrop={canEdit ? (e) => {
                e.preventDefault();
                this._onDrop(node.billet);
              } : undefined}
              onDragEnd={canEdit ? () => {
                this.setState({ dragBillet: undefined, dragOverBillet: undefined });
              } : undefined}
            >
              {/* Rank + Name — single line */}
              <div style={{
                fontWeight: 700, fontSize: '1rem', lineHeight: 1.3,
                color: isVacant ? '#777' : undefined,
                whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '100%',
              }}>
                {nameDisplay}
              </div>

              {/* Title */}
              {node.title && (
                <div style={{ fontSize: '0.82rem', opacity: 0.92, marginTop: 3, lineHeight: 1.3 }}>
                  {node.title}
                </div>
              )}

              {/* Unit/Office */}
              {unitOffice && (
                <div style={{ fontSize: '0.76rem', opacity: 0.85, marginTop: 2, lineHeight: 1.3 }}>
                  {unitOffice}
                </div>
              )}

              {/* Additional card lines from cardLines setting */}
              {(() => {
                const cardLines = parseCardLines(this.props.cardLines);
                return cardLines.map((tokens, li) => {
                  const text = buildCardLineText(tokens, node, this.props);
                  if (!text) return null;
                  return (
                    <div key={li} style={{ fontSize: '0.68rem', opacity: 0.7, marginTop: 2, lineHeight: 1.3, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '100%' }}>
                      {text}
                    </div>
                  );
                });
              })()}

              {/* Edit button (left) */}
              {canEdit && (
                <button
                  type="button"
                  data-export-hide="true"
                  title={strings.EditPositionTooltip}
                  onClick={(e) => {
                    e.stopPropagation();
                    this._startEdit(node.billet);
                  }}
                  style={{
                    position: 'absolute', top: 6, left: 6,
                    background: 'rgba(0,0,0,0.25)', border: 'none',
                    color: isSelected ? selectedFontColor : nodeFontColor,
                    cursor: 'pointer', borderRadius: '50%',
                    width: 22, height: 22, minWidth: 22, minHeight: 22, maxWidth: 22, maxHeight: 22,
                    fontSize: 11, lineHeight: '22px',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    padding: 0, margin: 0, boxSizing: 'border-box',
                    flexShrink: 0, flexGrow: 0, overflow: 'hidden',
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    WebkitAppearance: 'none', MozAppearance: 'none', appearance: 'none' as any,
                  }}
                >&#9998;</button>
              )}

              {/* Remove button (right) */}
              {canEdit && (
                <button
                  type="button"
                  data-export-hide="true"
                  title={strings.RemovePositionTooltip}
                  onClick={(e) => {
                    e.stopPropagation();
                    this.setState({ confirmRemoveBillet: node.billet });
                  }}
                  style={{
                    position: 'absolute', top: 6, right: 6,
                    background: 'rgba(0,0,0,0.25)', border: 'none',
                    color: isSelected ? selectedFontColor : nodeFontColor,
                    cursor: 'pointer', borderRadius: '50%',
                    width: 22, height: 22, minWidth: 22, minHeight: 22, maxWidth: 22, maxHeight: 22,
                    fontSize: 11, lineHeight: '22px',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    padding: 0, margin: 0, boxSizing: 'border-box',
                    flexShrink: 0, flexGrow: 0, overflow: 'hidden',
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    WebkitAppearance: 'none', MozAppearance: 'none', appearance: 'none' as any,
                  }}
                >&#10005;</button>
              )}
            </div>

            {/* Children subtree */}
            {subChildren.length > 0 && (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                <div style={{ width: 2, height: CONNECTOR_SIZE, background: this.props.lineColor }} />
                <div style={{ alignSelf: 'stretch', height: 2, background: this.props.lineColor }} />
                <ul style={{
                  listStyle: 'none', padding: 0, margin: 0,
                  display: 'inline-flex',
                }}>
                  {subChildren}
                </ul>
              </div>
            )}
          </li>
        );
      });
  }

  // ── Panel ─────────────────────────────────────────────────────────────────

  private _renderPanel(): JSX.Element | null {
    const { nodes, selectedBillet } = this.state;
    const {
      panelColumns, borderColor,
      selectedGradientStart, selectedGradientEnd, selectedBorderColor, selectedFontColor,
    } = this.props;

    const node = nodes.find(n => n.billet === selectedBillet);
    if (!node) return null;

    const panelLines = parseCardLines(panelColumns);
    const directReports = this._countDirectReports(node.billet, nodes);
    const subordinates = this._countSubordinates(node.billet, nodes, new Set());

    const panelLabelStyle: React.CSSProperties = {
      fontSize: '0.7rem', fontWeight: 700, color: this._tc('bodySubtext', '#888'),
      textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 2,
    };
    const panelValueStyle: React.CSSProperties = {
      fontSize: '0.9rem', color: this._tc('bodyText', '#1a1a1a'), marginBottom: 12, wordBreak: 'break-word',
    };

    const selGradient = `linear-gradient(135deg, ${selectedGradientStart} 0%, ${selectedGradientEnd} 100%)`;
    const headerBorder = selectedBorderColor || borderColor;
    const headerFont = selectedFontColor || 'rgba(255,255,255,1)';

    const ownerNode = nodes.find(n => n.billet === node.owner);
    const ownerDisplay = ownerNode
      ? [ownerNode.rank, ownerNode.name].filter(Boolean).join(' ') || ownerNode.billet
      : node.owner || undefined;

    const nameDisplay = node.name
      ? [node.rank, node.name].filter(Boolean).join(' ')
      : strings.VacantLabel;

    return (
      <div style={{
        width: '270px',
        border: `2px solid ${headerBorder}`,
        borderRadius: '10px',
        overflow: 'hidden',
        boxShadow: '0 6px 24px rgba(0,0,0,0.2)',
        fontFamily: "'Segoe UI', sans-serif",
        background: this._tc('bodyBackground', '#fff'),
      }}>
        {/* Header */}
        <div style={{ background: selGradient, color: headerFont, padding: '14px 16px', position: 'relative' }}>
          <button
            type="button"
            onClick={() => this.setState({ selectedBillet: undefined, panelPosition: undefined })}
            style={{
              position: 'absolute', top: 8, right: 8,
              background: 'rgba(255,255,255,0.25)', border: 'none',
              color: headerFont, cursor: 'pointer', borderRadius: '50%',
              width: 26, height: 26, minWidth: 26, minHeight: 26, maxWidth: 26, maxHeight: 26,
              fontSize: 14, lineHeight: '26px',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              padding: 0, margin: 0, boxSizing: 'border-box',
              flexShrink: 0, flexGrow: 0, overflow: 'hidden',
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              WebkitAppearance: 'none', MozAppearance: 'none', appearance: 'none' as any,
            }}
          >&#10005;</button>
          <div style={{ fontWeight: 700, fontSize: '1.05rem', paddingRight: 30 }}>{nameDisplay}</div>
          {node.title && (
            <div style={{ fontSize: '0.85rem', opacity: 0.9, marginTop: 3 }}>{node.title}</div>
          )}
          {node.billet && (
            <div style={{ fontSize: '0.78rem', opacity: 0.75, marginTop: 2 }}>{strings.BilletPanelPrefix} {node.billet}</div>
          )}
        </div>

        {/* Body */}
        <div style={{ padding: '14px 16px' }}>
          <div style={{ display: 'flex', gap: '12px', marginBottom: 14 }}>
            <div style={{ flex: 1, textAlign: 'center', background: this._tp('neutralLighter', '#f4f4f4'), borderRadius: '8px', padding: '8px 4px' }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: this._tc('bodyText', '#1a1a1a') }}>{directReports}</div>
              <div style={{ fontSize: '0.7rem', color: this._tc('bodySubtext', '#666'), textTransform: 'uppercase', letterSpacing: '0.05em' }}>{strings.DirectReportsLabel}</div>
            </div>
            <div style={{ flex: 1, textAlign: 'center', background: this._tp('neutralLighter', '#f4f4f4'), borderRadius: '8px', padding: '8px 4px' }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: this._tc('bodyText', '#1a1a1a') }}>{subordinates}</div>
              <div style={{ fontSize: '0.7rem', color: this._tc('bodySubtext', '#666'), textTransform: 'uppercase', letterSpacing: '0.05em' }}>{strings.SubordinatesLabel}</div>
            </div>
          </div>

          {(node.office || node.unit) && (
            <div>
              <div style={panelLabelStyle}>{strings.UnitOfficeLabel}</div>
              <div style={panelValueStyle}>{[node.unit, node.office].filter(Boolean).join('/')}</div>
            </div>
          )}

          {ownerDisplay && (
            <div>
              <div style={panelLabelStyle}>{strings.ReportsToLabel}</div>
              <div style={panelValueStyle}>{ownerDisplay}</div>
            </div>
          )}

          {panelLines.map((tokens, li) => {
            const colNames = tokens.filter(t => !isLiteralSeparator(t));
            const text = buildPanelLineText(tokens, node);
            if (!text) return null;
            const label = colNames.join(' | ');
            // Single column: use formatPanelValue for email/phone links
            const value = colNames.length === 1 && node.details[colNames[0]]
              ? formatPanelValue(colNames[0], node.details[colNames[0]])
              : <>{text}</>;
            return (
              <div key={li}>
                <div style={panelLabelStyle}>{label}</div>
                <div style={panelValueStyle}>{value}</div>
              </div>
            );
          })}

          {panelLines.length === 0 && !ownerDisplay && directReports === 0 && (
            <div style={{ color: '#999', fontSize: '0.85rem' }}>
              {strings.PanelConfigureHint}
            </div>
          )}
        </div>
      </div>
    );
  }

  // ── Add modal ─────────────────────────────────────────────────────────────

  private _renderAddModal(): JSX.Element | null {
    if (!this.state.showAddModal) return null;
    const { addForm, addSaving, addError, nodes, invalidPanelCols } = this.state;
    const billets = nodes.map(n => n.billet).filter(Boolean).sort();

    // Valid panel columns for extra form fields
    const validPanelCols = extractColumnNames(this.props.panelColumns)
      .filter(c => !invalidPanelCols.includes(c));

    const inputStyle: React.CSSProperties = { width: '100%', padding: '6px 8px', border: `1px solid ${this._tc('smallInputBorder', '#ccc')}`, borderRadius: '4px', fontSize: '0.9rem', boxSizing: 'border-box', background: this._tc('inputBackground', '#fff'), color: this._tc('bodyText', '#333') };
    const field = (label: string, key: keyof IAddForm, required?: boolean, placeholder?: string): JSX.Element => (
      <div style={{ marginBottom: 12 }} key={key}>
        <label style={{ fontWeight: 600, fontSize: '0.85rem', display: 'block', marginBottom: 4 }}>
          {label}{required && <span style={{ color: 'red' }}> *</span>}
        </label>
        <input
          type="text"
          value={addForm[key] as string}
          placeholder={placeholder}
          onChange={e => this.setState(prev => ({ addForm: { ...prev.addForm, [key]: e.target.value } }))}
          style={inputStyle}
          disabled={addSaving}
        />
      </div>
    );

    return (
      <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ background: this._tc('bodyBackground', '#fff'), color: this._tc('bodyText', '#1a1a1a'), borderRadius: '10px', padding: '24px', width: 480, maxWidth: '95vw', maxHeight: '85vh', overflowY: 'auto', boxShadow: '0 8px 32px rgba(0,0,0,0.25)', fontFamily: "'Segoe UI', sans-serif", position: 'relative' }}>
          <button
            type="button"
            onClick={() => this.setState({ showAddModal: false, addForm: { ...BLANK_ADD_FORM }, addError: '' })}
            disabled={addSaving}
            style={{ position: 'absolute', top: 8, right: 8, background: 'none', border: 'none', fontSize: '1.25rem', cursor: 'pointer', color: this._tc('bodySubtext', '#888'), lineHeight: 1, padding: '4px 8px' }}
            aria-label="Close"
          >&times;</button>
          <h3 style={{ margin: '0 0 16px', fontSize: '1.1rem' }}>{strings.AddPositionTitle}</h3>
          {field(strings.BilletInputLabel, 'billet', true)}

          <div style={{ marginBottom: 12 }}>
            <label style={{ fontWeight: 600, fontSize: '0.85rem', display: 'block', marginBottom: 4 }}>{strings.OwnerInputLabel}</label>
            <input
              type="text"
              value={addForm.owner}
              list="billet-options"
              placeholder={strings.OwnerInputPlaceholder}
              onChange={e => this.setState(prev => ({ addForm: { ...prev.addForm, owner: e.target.value } }))}
              style={inputStyle}
              disabled={addSaving}
            />
            <datalist id="billet-options">
              {billets.map(b => <option key={b} value={b} />)}
            </datalist>
          </div>

          {field(strings.RankInputLabel, 'rank')}
          {field(strings.NameInputLabel, 'name')}
          {field(strings.TitleInputLabel, 'title', true)}
          {field(strings.UnitInputLabel, 'unit', true)}
          {field(strings.OfficeInputLabel, 'office', true)}
          {field(strings.SpecialtyCodeInputLabel, 'specialtyCode')}

          {/* Panel column fields */}
          {validPanelCols.map(col => (
            <div style={{ marginBottom: 12 }} key={`extra-${col}`}>
              <label style={{ fontWeight: 600, fontSize: '0.85rem', display: 'block', marginBottom: 4 }}>{col}</label>
              <input
                type="text"
                value={addForm.extras[col] || ''}
                onChange={e => this.setState(prev => ({
                  addForm: { ...prev.addForm, extras: { ...prev.addForm.extras, [col]: e.target.value } }
                }))}
                style={inputStyle}
                disabled={addSaving}
              />
            </div>
          ))}

          {addError && <div style={{ color: 'red', fontSize: '0.85rem', marginBottom: 10 }}>{addError}</div>}
          <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', marginTop: 4 }}>
            <button type="button" onClick={this._addMemberAndAnother} disabled={addSaving} style={{ padding: '7px 18px', border: 'none', borderRadius: '4px', background: this.props.saveBtnColor || '#0078d4', color: '#fff', cursor: 'pointer', fontSize: '0.9rem', fontWeight: 600 }}>{addSaving ? strings.AddSavingLabel : 'Save & Add Another'}</button>
            <div style={{ display: 'flex', gap: 8 }}>
              <button type="button" onClick={() => this.setState({ showAddModal: false, addForm: { ...BLANK_ADD_FORM }, addError: '' })} disabled={addSaving} style={{ padding: '7px 18px', border: '1px solid #ccc', borderRadius: '4px', background: this.props.cancelBtnColor || '#f4f4f4', cursor: 'pointer', fontSize: '0.9rem' }}>{strings.CancelButton}</button>
              <button type="button" onClick={this._addMember} disabled={addSaving} style={{ padding: '7px 18px', border: 'none', borderRadius: '4px', background: this.props.addFormBtnColor || '#28a745', color: '#fff', cursor: 'pointer', fontSize: '0.9rem', fontWeight: 600 }}>{addSaving ? strings.AddSavingLabel : strings.AddButtonLabel}</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── Confirm remove ────────────────────────────────────────────────────────

  private _renderConfirmRemove(): JSX.Element | null {
    const { confirmRemoveBillet, removeSaving, nodes } = this.state;
    if (!confirmRemoveBillet) return null;
    const node = nodes.find(n => n.billet === confirmRemoveBillet);
    if (!node) return null;

    const childCount = this._countDirectReports(node.billet, nodes);
    const ownerNode = nodes.find(n => n.billet === node.owner);
    const ownerDisplay = ownerNode
      ? [ownerNode.rank, ownerNode.name].filter(Boolean).join(' ') || ownerNode.billet
      : node.owner ? node.owner : 'the top level';

    return (
      <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ background: this._tc('bodyBackground', '#fff'), color: this._tc('bodyText', '#1a1a1a'), borderRadius: '10px', padding: '24px', width: 360, maxWidth: '95vw', boxShadow: '0 8px 32px rgba(0,0,0,0.25)', fontFamily: "'Segoe UI', sans-serif" }}>
          <h3 style={{ margin: '0 0 12px', fontSize: '1.1rem', color: '#c00' }}>{strings.RemovePositionTitle}</h3>
          <p style={{ fontSize: '0.92rem', color: this._tc('bodyText', '#333'), lineHeight: 1.5, margin: '0 0 14px' }}>
            {strings.RemovePromptPrefix} <strong>{[node.rank, node.name].filter(Boolean).join(' ') || node.billet || 'this position'}</strong>.
          </p>
          {childCount > 0 && (
            <p style={{ fontSize: '0.88rem', background: '#fff3cd', border: '1px solid #ffc107', borderRadius: '6px', padding: '10px 12px', margin: '0 0 14px', color: '#7c5500' }}>
              {strings.RemoveReportsWarningPrefix} <strong>{childCount}</strong> {childCount !== 1 ? strings.RemoveReportsWarningSuffix_Plural : strings.RemoveReportsWarningSuffix_Singular} <strong>{ownerDisplay}</strong>.
            </p>
          )}
          <p style={{ fontSize: '0.88rem', color: this._tc('bodySubtext', '#555'), margin: '0 0 18px' }}>{strings.RemoveConfirmQuestion}</p>
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
            <button type="button" onClick={() => this.setState({ confirmRemoveBillet: undefined })} disabled={removeSaving} style={{ padding: '7px 18px', border: '1px solid #ccc', borderRadius: '4px', background: this.props.cancelBtnColor || '#f4f4f4', cursor: 'pointer', fontSize: '0.9rem' }}>{strings.CancelButton}</button>
            <button type="button" onClick={() => this._removeMember(confirmRemoveBillet)} disabled={removeSaving} style={{ padding: '7px 18px', border: 'none', borderRadius: '4px', background: this.props.deleteBtnColor || '#c00', color: '#fff', cursor: 'pointer', fontSize: '0.9rem', fontWeight: 600 }}>{removeSaving ? strings.RemovingLabel : strings.YesRemoveButton}</button>
          </div>
        </div>
      </div>
    );
  }

  // ── Edit member ───────────────────────────────────────────────────────────

  private _startEdit(billet: string): void {
    const node = this.state.nodes.find(n => n.billet === billet);
    if (!node) return;
    this.setState({
      showEditModal: true,
      editBillet: billet,
      editError: '',
      editForm: {
        title: node.title,
        unit: node.unit,
        office: node.office,
        rank: node.rank,
        name: node.name,
        billet: node.billet,
        owner: node.owner,
        specialtyCode: node.specialtyCode,
        extras: { ...node.details },
      },
    });
  }

  private _saveEdit = async (): Promise<void> => {
    const { list, titleField, unitField, officeField, rankField, nameField, billetField, ownerField, specialtyCodeField, panelColumns } = this.props;
    const { editBillet, editForm, nodes } = this.state;
    if (!editBillet) return;

    // Required fields validation
    const missing: string[] = [];
    if (!editForm.billet.trim()) missing.push(strings.BilletInputLabel);
    if (!editForm.title.trim()) missing.push(strings.TitleInputLabel);
    if (!editForm.office.trim()) missing.push(strings.OfficeInputLabel);
    if (!editForm.unit.trim()) missing.push(strings.UnitInputLabel);
    if (missing.length > 0) {
      this.setState({ editError: `${strings.RequiredPrefix} ${missing.join(', ')}` });
      return;
    }

    // If billet changed, make sure the new value doesn't collide with another existing node
    const billetChanged = editForm.billet.trim() !== editBillet;
    if (billetChanged && nodes.some(n => n.billet === editForm.billet.trim())) {
      this.setState({ editError: `"${editForm.billet.trim()}" ${strings.BilletCollisionError}` });
      return;
    }

    const node = nodes.find(n => n.billet === editBillet);
    if (!node) return;

    this.setState({ editSaving: true, editError: '' });
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        [titleField]: editForm.title,
        [unitField]: editForm.unit,
        [officeField]: editForm.office,
        [rankField]: editForm.rank,
        [nameField]: editForm.name,
        [billetField]: editForm.billet.trim(),
        [ownerField]: editForm.owner.trim(),
        [specialtyCodeField]: editForm.specialtyCode,
      };
      // Include panel column extras
      const validPanelCols = extractColumnNames(panelColumns)
        .filter(c => !this.state.invalidPanelCols.includes(c));
      for (const col of validPanelCols) {
        if (editForm.extras[col] !== undefined) payload[col] = editForm.extras[col];
      }

      await this.sp.web.lists.getByTitle(list).items.getById(node.listItemId).update(payload);

      // If the billet ID changed, cascade-update all direct children so they keep their parent
      if (billetChanged) {
        const children = nodes.filter(n => n.owner === editBillet);
        await Promise.all(
          children.map(child =>
            this.sp.web.lists.getByTitle(list).items.getById(child.listItemId).update({
              [ownerField]: editForm.billet.trim(),
            })
          )
        );
      }

      this.setState({ showEditModal: false, editBillet: undefined, editSaving: false, editError: '' });
      await this._fetchData();
    } catch {
      this.setState({ editSaving: false, editError: strings.EditFailedError });
    }
  };

  private _renderEditModal(): JSX.Element | null {
    if (!this.state.showEditModal) return null;
    const { editForm, editSaving, editError, editBillet, nodes, invalidPanelCols } = this.state;

    const billetChanged = editForm.billet.trim() !== editBillet;
    // Owner autocomplete: all billets except the one being edited (can't own itself)
    const billets = nodes.map(n => n.billet).filter(b => b !== editBillet).sort();

    // Valid panel columns for extra form fields
    const validPanelCols = extractColumnNames(this.props.panelColumns)
      .filter(c => !invalidPanelCols.includes(c));

    const inputStyle: React.CSSProperties = { width: '100%', padding: '6px 8px', border: `1px solid ${this._tc('smallInputBorder', '#ccc')}`, borderRadius: '4px', fontSize: '0.9rem', boxSizing: 'border-box', background: this._tc('inputBackground', '#fff'), color: this._tc('bodyText', '#333') };
    const field = (label: string, key: keyof IAddForm, required?: boolean, placeholder?: string): JSX.Element => (
      <div style={{ marginBottom: 12 }} key={key}>
        <label style={{ fontWeight: 600, fontSize: '0.85rem', display: 'block', marginBottom: 4 }}>
          {label}{required && <span style={{ color: 'red' }}> *</span>}
        </label>
        <input
          type="text"
          value={editForm[key] as string}
          placeholder={placeholder}
          onChange={e => this.setState(prev => ({ editForm: { ...prev.editForm, [key]: e.target.value } }))}
          style={inputStyle}
          disabled={editSaving}
        />
      </div>
    );

    return (
      <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ background: this._tc('bodyBackground', '#fff'), color: this._tc('bodyText', '#1a1a1a'), borderRadius: '10px', padding: '24px', width: 480, maxWidth: '95vw', maxHeight: '85vh', overflowY: 'auto', boxShadow: '0 8px 32px rgba(0,0,0,0.25)', fontFamily: "'Segoe UI', sans-serif", position: 'relative' }}>
          <button
            type="button"
            onClick={() => this.setState({ showEditModal: false, editBillet: undefined, editError: '' })}
            disabled={editSaving}
            style={{ position: 'absolute', top: 8, right: 8, background: 'none', border: 'none', fontSize: '1.25rem', cursor: 'pointer', color: this._tc('bodySubtext', '#888'), lineHeight: 1, padding: '4px 8px' }}
            aria-label="Close"
          >&times;</button>
          <h3 style={{ margin: '0 0 4px', fontSize: '1.1rem' }}>{strings.EditPositionTitle}</h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.8rem', color: this._tc('bodySubtext', '#888') }}>{strings.BilletPanelPrefix} {editBillet}</p>

          {/* Billet — editable with cascade warning */}
          <div style={{ marginBottom: 12 }}>
            <label style={{ fontWeight: 600, fontSize: '0.85rem', display: 'block', marginBottom: 4 }}>
              {strings.BilletInputLabel}<span style={{ color: 'red' }}> *</span>
            </label>
            <input
              type="text"
              value={editForm.billet}
              onChange={e => this.setState(prev => ({ editForm: { ...prev.editForm, billet: e.target.value } }))}
              style={{ ...inputStyle, border: `1px solid ${billetChanged ? '#f59e0b' : this._tc('smallInputBorder', '#ccc')}` }}
              disabled={editSaving}
            />
            {billetChanged && (
              <div style={{ fontSize: '0.78rem', color: '#92400e', background: '#fef3c7', border: '1px solid #f59e0b', borderRadius: '4px', padding: '5px 8px', marginTop: 4 }}>
                {strings.BilletChangeWarning}
              </div>
            )}
          </div>

          {/* Owner */}
          <div style={{ marginBottom: 12 }}>
            <label style={{ fontWeight: 600, fontSize: '0.85rem', display: 'block', marginBottom: 4 }}>{strings.OwnerInputLabel}</label>
            <input
              type="text"
              value={editForm.owner}
              list="edit-billet-options"
              placeholder={strings.OwnerInputPlaceholder}
              onChange={e => this.setState(prev => ({ editForm: { ...prev.editForm, owner: e.target.value } }))}
              style={inputStyle}
              disabled={editSaving}
            />
            <datalist id="edit-billet-options">
              {billets.map(b => <option key={b} value={b} />)}
            </datalist>
          </div>

          {field(strings.RankInputLabel, 'rank')}
          {field(strings.NameInputLabel, 'name')}
          {field(strings.TitleInputLabel, 'title', true)}
          {field(strings.UnitInputLabel, 'unit', true)}
          {field(strings.OfficeInputLabel, 'office', true)}
          {field(strings.SpecialtyCodeInputLabel, 'specialtyCode')}

          {/* Panel column fields */}
          {validPanelCols.map(col => (
            <div style={{ marginBottom: 12 }} key={`extra-${col}`}>
              <label style={{ fontWeight: 600, fontSize: '0.85rem', display: 'block', marginBottom: 4 }}>{col}</label>
              <input
                type="text"
                value={editForm.extras[col] || ''}
                onChange={e => this.setState(prev => ({
                  editForm: { ...prev.editForm, extras: { ...prev.editForm.extras, [col]: e.target.value } }
                }))}
                style={inputStyle}
                disabled={editSaving}
              />
            </div>
          ))}

          {editError && <div style={{ color: 'red', fontSize: '0.85rem', marginBottom: 10 }}>{editError}</div>}

          <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', marginTop: 4 }}>
            <button
              type="button"
              onClick={() => { this.setState({ showEditModal: false, editBillet: undefined, editError: '', confirmRemoveBillet: editBillet }); }}
              disabled={editSaving}
              style={{ padding: '7px 18px', border: 'none', borderRadius: '4px', background: this.props.deleteBtnColor || '#c00', color: '#fff', cursor: 'pointer', fontSize: '0.9rem', fontWeight: 600 }}
            >{strings.DeleteButton}</button>
            <div style={{ display: 'flex', gap: 8 }}>
              <button
                type="button"
                onClick={() => this.setState({ showEditModal: false, editBillet: undefined, editError: '' })}
                disabled={editSaving}
                style={{ padding: '7px 18px', border: '1px solid #ccc', borderRadius: '4px', background: this.props.cancelBtnColor || '#f4f4f4', cursor: 'pointer', fontSize: '0.9rem' }}
              >{strings.CancelButton}</button>
              <button
                type="button"
                onClick={this._saveEdit}
                disabled={editSaving}
                style={{ padding: '7px 18px', border: 'none', borderRadius: '4px', background: this.props.saveBtnColor || '#28a745', color: '#fff', cursor: 'pointer', fontSize: '0.9rem', fontWeight: 600 }}
              >{editSaving ? strings.EditSavingLabel : strings.SaveChangesButton}</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── Confirm reparent ──────────────────────────────────────────────────────

  private _renderConfirmReparent(): JSX.Element | null {
    const { confirmReparent, reparentSaving, reparentError, nodes } = this.state;
    if (!confirmReparent) return null;

    const child = nodes.find(n => n.billet === confirmReparent.childBillet);
    const newParent = confirmReparent.newParentBillet
      ? nodes.find(n => n.billet === confirmReparent.newParentBillet)
      : null;

    // If NOT vacant, use Rank + Name. If vacant, use the Billet.
    const childDisplay = child
      ? (child.name ? [child.rank, child.name].filter(Boolean).join(' ') : child.billet)
      : confirmReparent.childBillet;
    const parentDisplay = newParent
      ? (newParent.name ? [newParent.rank, newParent.name].filter(Boolean).join(' ') : newParent.billet)
      : strings.MoveTopLevel;

    return (
      <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ background: this._tc('bodyBackground', '#fff'), color: this._tc('bodyText', '#1a1a1a'), borderRadius: '10px', padding: '24px', width: 360, maxWidth: '95vw', boxShadow: '0 8px 32px rgba(0,0,0,0.25)', fontFamily: "'Segoe UI', sans-serif" }}>
          <h3 style={{ margin: '0 0 12px', fontSize: '1.1rem' }}>{strings.MovePositionTitle}</h3>
          {reparentError ? (
            <p style={{ fontSize: '0.92rem', color: this._tc('bodyText', '#000'), lineHeight: 1.6, margin: '0 0 18px' }}>
              {reparentError}
            </p>
          ) : (
            <p style={{ fontSize: '0.92rem', color: this._tc('bodyText', '#333'), lineHeight: 1.6, margin: '0 0 18px' }}>
              {strings.MovePromptPrefix} <strong>{childDisplay}</strong> {strings.MovePromptMiddle} <strong>{parentDisplay}</strong>?
            </p>
          )}
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
            <button type="button" onClick={() => this.setState({ confirmReparent: undefined, reparentError: '', dragBillet: undefined, dragOverBillet: undefined })} disabled={reparentSaving} style={{ padding: '7px 18px', border: '1px solid #ccc', borderRadius: '4px', background: reparentError ? (this.props.okBtnColor || '#f4f4f4') : (this.props.cancelBtnColor || '#f4f4f4'), cursor: 'pointer', fontSize: '0.9rem' }}>{reparentError ? 'OK' : strings.CancelButton}</button>
            {!reparentError && (
              <button type="button" onClick={this._reparentMember} disabled={reparentSaving} style={{ padding: '7px 18px', border: 'none', borderRadius: '4px', background: this.props.moveBtnColor || '#1a1a2e', color: '#fff', cursor: 'pointer', fontSize: '0.9rem', fontWeight: 600 }}>{reparentSaving ? strings.MoveSavingLabel : strings.MoveButton}</button>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ── Render ────────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const {
      nodes, loading, configError, searchQuery, unitFilter: stateUnitFilter, officeFilter,
      selectedBillet, canEdit, showAddModal, showEditModal,
      dragBillet, dragOverBillet, orphanBillets, invalidPanelCols,
    } = this.state;
    const { webpartTitle, isPageEditMode, unitFilterLock } = this.props;

    // When unitFilterLock is set, override the dropdown filter
    const unitFilter = (unitFilterLock || '').trim() || stateUnitFilter;

    // Effective canEdit: requires permissions AND page not in edit mode
    const effectiveCanEdit = canEdit && !isPageEditMode;

    const specialtyColors = this._parseSpecialtyColors();

    const wrapperStyle: React.CSSProperties = {
      fontFamily: "'Segoe UI', sans-serif",
      padding: 0,
      margin: 0,
      width: '100%',
      maxWidth: '100%',
      boxSizing: 'border-box',
    };

    if (configError) {
      return (
        <div style={wrapperStyle}>
          <div style={{ color: '#666', padding: '2rem', fontSize: '1rem', border: '2px dashed #ccc', borderRadius: '8px', margin: '1rem', textAlign: 'center' }}>
            {strings.ConfigureMessage}
          </div>
        </div>
      );
    }

    if (isPageEditMode) {
      return (
        <div style={wrapperStyle}>
          <div style={{ color: '#666', padding: '2rem', fontSize: '1rem', border: '2px dashed #ccc', borderRadius: '8px', margin: '1rem', textAlign: 'center' }}>
            {strings.SavePageMessage}
          </div>
        </div>
      );
    }

    // Unique unit values for dropdown — ordered by Unit Sort Priority, then alphabetical
    const uniqueUnits: string[] = [];
    const seenLower = new Set<string>();
    for (const node of nodes) {
      if (node.unit) {
        const lower = node.unit.toLowerCase();
        if (!seenLower.has(lower)) { seenLower.add(lower); uniqueUnits.push(node.unit); }
      }
    }

    // Sort units: prioritized units first (in priority order), then remaining alphabetically
    const unitPriorityList = (this.props.unitSortPriority || '').split(',').map(s => s.trim()).filter(Boolean);
    const unitPriorityLower = unitPriorityList.map(s => s.toLowerCase());
    uniqueUnits.sort((a, b) => {
      const aIdx = unitPriorityLower.indexOf(a.toLowerCase());
      const bIdx = unitPriorityLower.indexOf(b.toLowerCase());
      const aInPriority = aIdx !== -1;
      const bInPriority = bIdx !== -1;
      if (aInPriority && bInPriority) return aIdx - bIdx;
      if (aInPriority) return -1;
      if (bInPriority) return 1;
      return a.toLowerCase().localeCompare(b.toLowerCase());
    });

    const panel = selectedBillet !== undefined ? this._renderPanel() : null;

    // Determine if we should show the unit filter dropdown (hide if only 1 unit or locked to a unit)
    const showUnitFilter = uniqueUnits.length > 1 && !(unitFilterLock || '').trim();

    // Office filter: always enabled; collect offices scoped to selected unit or all units
    const effectiveUnit = unitFilter || (uniqueUnits.length === 1 ? uniqueUnits[0] : '');

    // Unique office values, scoped to effective unit if one is active, otherwise all nodes
    const uniqueOffices: string[] = [];
    {
      const seenOfficeLower = new Set<string>();
      for (const node of nodes) {
        if (!node.office) continue;
        if (effectiveUnit && node.unit.toLowerCase() !== effectiveUnit.toLowerCase()) continue;
        const lower = node.office.toLowerCase();
        if (!seenOfficeLower.has(lower)) { seenOfficeLower.add(lower); uniqueOffices.push(node.office); }
      }
      const officePriorityList = (this.props.officeSortPriority || '').split(',').map(s => s.trim()).filter(Boolean);
      const officePriorityLower = officePriorityList.map(s => s.toLowerCase());
      uniqueOffices.sort((a, b) => {
        const aIdx = officePriorityLower.indexOf(a.toLowerCase());
        const bIdx = officePriorityLower.indexOf(b.toLowerCase());
        const aInPriority = aIdx !== -1;
        const bInPriority = bIdx !== -1;
        if (aInPriority && bInPriority) return aIdx - bIdx;
        if (aInPriority) return -1;
        if (bInPriority) return 1;
        return a.toLowerCase().localeCompare(b.toLowerCase());
      });
    }

    // Chart inner style: horizontal overflow with drag scrolling for Modern pages
    const chartInnerStyle: React.CSSProperties = {
      width: '100%', maxWidth: '100%', overflowX: 'auto', overflowY: 'visible', cursor: 'grab', boxSizing: 'border-box',
    };

    return (
      <div style={wrapperStyle}>
        {/* Title */}
        <h2 style={{ fontSize: '2rem', fontWeight: 'bold', marginBottom: '1rem', textAlign: 'center' }}>{webpartTitle}</h2>

        {/* ── Orphan warning ──────────────────────────────────────────────── */}
        {orphanBillets.length > 0 && (
          <div style={{ background: '#fff3cd', border: '1px solid #ffc107', borderRadius: '6px', padding: '10px 14px', marginBottom: '0.75rem', fontSize: '0.85rem', color: '#7c5500' }}>
            <strong>{strings.OrphanWarningPrefix}</strong> {orphanBillets.length} {orphanBillets.length !== 1 ? strings.OrphanWarningSuffix_Plural : strings.OrphanWarningSuffix_Singular} <em>{orphanBillets.join(', ')}</em>
          </div>
        )}

        {/* ── Invalid panel columns notice ────────────────────────────────── */}
        {invalidPanelCols.length > 0 && (
          <div style={{ background: '#e8f4fd', border: '1px solid #90caf9', borderRadius: '6px', padding: '10px 14px', marginBottom: '0.75rem', fontSize: '0.85rem', color: '#1565c0' }}>
            <strong>{strings.InvalidPanelColsPrefix}</strong> {invalidPanelCols.length !== 1 ? strings.InvalidPanelColsSuffix_Plural : strings.InvalidPanelColsSuffix_Singular} <em>{invalidPanelCols.join(', ')}</em>. {strings.InvalidPanelColsTail}
          </div>
        )}

        {/* ── Toolbar + Search & filter (sticky) ────────────────────────── */}
        <div style={{ position: 'sticky', top: 0, zIndex: 10, background: this._tc('bodyBackground', '#fff'), padding: '0.5rem 0', marginBottom: '1rem' }}>
          <div style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center', gap: '0.5rem', marginBottom: '0.5rem' }}>
            {effectiveCanEdit && (
              <button type="button" onClick={() => this.setState({ showAddModal: true, addForm: { ...BLANK_ADD_FORM }, addError: '' })} style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', padding: '0.4rem 0.9rem', border: 'none', borderRadius: '4px', cursor: 'pointer', background: this.props.addPositionBtnColor || '#1a1a2e', color: '#fff', fontSize: '0.875rem', fontWeight: 600, whiteSpace: 'nowrap' }}>
                {strings.AddPositionButton}
              </button>
            )}
            <button type="button" onClick={this._exportToImage} title={strings.SaveAsImageButton} style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', padding: '0.4rem 0.9rem', border: 'none', borderRadius: '4px', cursor: 'pointer', background: this.props.saveImageBtnColor || '#1a1a2e', color: '#fff', fontSize: '0.875rem', fontWeight: 600, whiteSpace: 'nowrap' }}>
              {strings.SaveAsImageButton}
            </button>
          </div>
          <div style={{ display: 'flex', gap: '0.5rem' }}>
            <input type="text" placeholder={strings.SearchPlaceholder} value={searchQuery} onChange={this._onSearchChange} style={{ flex: 1, padding: '0.5rem', border: `1px solid ${this._tc('smallInputBorder', '#ccc')}`, borderRadius: '4px', fontSize: '1rem', boxSizing: 'border-box', background: this._tc('inputBackground', '#fff'), color: this._tc('bodyText', '#333') }} />
            {showUnitFilter && (
              <select value={unitFilter} onChange={this._onUnitChange} style={{ padding: '0.5rem 0.75rem', border: `1px solid ${this._tc('smallInputBorder', '#ccc')}`, borderRadius: '4px', fontSize: '1rem', background: this._tc('inputBackground', '#fff'), color: this._tc('bodyText', '#333'), minWidth: '160px', boxSizing: 'border-box' }}>
                <option value="">{strings.AllUnitsOption}</option>
                {uniqueUnits.map(u => <option key={u} value={u}>{u}</option>)}
              </select>
            )}
            <select value={officeFilter} onChange={this._onOfficeChange} style={{ padding: '0.5rem 0.75rem', border: `1px solid ${this._tc('smallInputBorder', '#ccc')}`, borderRadius: '4px', fontSize: '1rem', background: this._tc('inputBackground', '#fff'), color: this._tc('bodyText', '#333'), minWidth: '160px', boxSizing: 'border-box' }}>
              <option value="">{strings.AllOfficesOption}</option>
              {uniqueOffices.map(o => <option key={o} value={o}>{o}</option>)}
            </select>
          </div>
        </div>

        {/* ── Color key legend (inside export area) ────────────────────── */}
        {/* ── Chart + panel ─────────────────────────────────────────────────── */}
        <div style={{ position: 'relative', width: '100%', maxWidth: '100%', boxSizing: 'border-box' }} ref={this._chartWrapperRef}>
          <div
            ref={this._chartInnerRef}
            style={chartInnerStyle}
            onMouseDown={this._onDragScrollStart}
            onMouseMove={this._onDragScrollMove}
            onMouseUp={this._onDragScrollEnd}
            onMouseLeave={this._onDragScrollEnd}
          >

            {/* Color key legend — inside the export capture area */}
            {specialtyColors.some(c => c.Code !== 'Other') && (
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginBottom: '1rem', alignItems: 'center', paddingLeft: '20px' }}>
                {specialtyColors.map(sc => (
                  <div key={sc.Code} style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <div style={{
                      width: 18, height: 18, borderRadius: '3px',
                      background: `linear-gradient(135deg, ${sc.GradientStart} 0%, ${sc.GradientEnd} 100%)`,
                      border: '1px solid rgba(0,0,0,0.15)', flexShrink: 0,
                    }} />
                    <span style={{ fontSize: '0.75rem', color: '#555', whiteSpace: 'nowrap' }}>{sc.Code}</span>
                  </div>
                ))}
              </div>
            )}

            {/* Root drop zone — only visible while dragging */}
            {dragBillet && (
              <div
                onDragOver={(e) => { e.preventDefault(); if (dragOverBillet !== '__root__') this.setState({ dragOverBillet: '__root__' }); }}
                onDragLeave={() => { if (dragOverBillet === '__root__') this.setState({ dragOverBillet: undefined }); }}
                onDrop={(e) => { e.preventDefault(); this._onDrop('__root__'); }}
                style={{
                  width: '100%', padding: '12px', marginBottom: '12px',
                  border: `2px dashed ${dragOverBillet === '__root__' ? '#0078d4' : '#aaa'}`,
                  borderRadius: '8px', textAlign: 'center',
                  color: dragOverBillet === '__root__' ? '#0078d4' : '#999',
                  fontSize: '0.85rem',
                  background: dragOverBillet === '__root__' ? '#e3f2fd' : 'transparent',
                  transition: 'all 0.15s',
                  boxSizing: 'border-box',
                }}
              >
                {strings.RootDropZoneText}
              </div>
            )}

            {loading ? (
              <div style={{ padding: '2rem', color: '#666' }}>{strings.LoadingMessage}</div>
            ) : nodes.length === 0 ? (
              <div style={{ padding: '2rem', color: '#666' }}>{strings.NoDataMessage}</div>
            ) : (
              <div style={{ display: 'inline-block', minWidth: '100%', textAlign: 'center', padding: '0 1rem', boxSizing: 'border-box' }}>
                {(() => {
                  // Identify root nodes (no owner or owner not in list)
                  const billetSet = new Set(nodes.map(n => n.billet));
                  const roots = nodes.filter(n => !n.owner || n.owner === '' || !billetSet.has(n.owner));
                  // Sort roots by unit priority then office priority
                  const unitPriority = (this.props.unitSortPriority || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
                  const officePriority = (this.props.officeSortPriority || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
                  const unitIdx = (v: string): number => { const i = unitPriority.indexOf((v || '').toLowerCase()); return i === -1 ? Infinity : i; };
                  const officeIdx = (v: string): number => { const i = officePriority.indexOf((v || '').toLowerCase()); return i === -1 ? Infinity : i; };
                  roots.sort((a, b) => {
                    const uA = unitIdx(a.unit), uB = unitIdx(b.unit);
                    if (uA !== uB) return uA - uB;
                    // Same priority tier — sort alphabetically by unit name so different units don't intermix
                    if (uA === Infinity) {
                      const unitCmp = (a.unit || '').toLowerCase().localeCompare((b.unit || '').toLowerCase());
                      if (unitCmp !== 0) return unitCmp;
                    }
                    const oA = officeIdx(a.office), oB = officeIdx(b.office);
                    if (oA !== oB) return oA - oB;
                    return (a.office || '').toLowerCase().localeCompare((b.office || '').toLowerCase());
                  });
                  // Group consecutive roots by unit
                  const groups: IOrgNode[][] = [];
                  roots.forEach(r => {
                    const last = groups[groups.length - 1];
                    if (last && last[0].unit === r.unit) { last.push(r); }
                    else { groups.push([r]); }
                  });
                  // Render each unit group as its own row
                  return groups.map((group, gi) => {
                    // Create a virtual parent billet to scope _buildTree to this group's roots
                    const groupBillets = new Set(group.map(r => r.billet));
                    // Filter nodes so _buildTree only picks up roots in this group
                    const scopedNodes = nodes.map(n => {
                      if (groupBillets.has(n.billet)) {
                        // Ensure these roots share a common virtual parent
                        return { ...n, owner: `__group_${gi}__` };
                      }
                      return n;
                    });
                    return (
                      <ul key={`root-group-${gi}`} style={{ listStyle: 'none', padding: 0, margin: 0, display: 'inline-flex', flexWrap: 'wrap', justifyContent: 'center', width: '100%' }}>
                        {this._buildTree(`__group_${gi}__`, scopedNodes, new Set(), true)}
                      </ul>
                    );
                  });
                })()}
              </div>
            )}
          </div>

        </div>

        {/* Detail panel — fixed, viewport-centered so it never clips on any page type */}
        {panel && (
          <>
            {/* Backdrop: closes on outside click */}
            {!dragBillet && (
              <div
                style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 100, background: 'rgba(0,0,0,0.45)', cursor: 'default' }}
                onClick={() => this.setState({ selectedBillet: undefined, panelPosition: undefined })}
              />
            )}
            {/* Panel */}
            <div style={{
              position: 'fixed',
              top: '50%',
              left: '50%',
              transform: 'translate(-50%, -50%)',
              zIndex: 101,
              maxHeight: '85vh',
              overflowY: 'auto',
              borderRadius: '10px',
            }}>
              {panel}
            </div>
          </>
        )}

        {/* Modals */}
        {showAddModal && this._renderAddModal()}
        {showEditModal && this._renderEditModal()}
        {this._renderConfirmRemove()}
        {this._renderConfirmReparent()}
      </div>
    );
  }
}

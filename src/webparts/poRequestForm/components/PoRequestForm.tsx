// src/webparts/poRequestForm/components/PoRequestForm.tsx
import * as React from 'react';
import styles from './PoRequestForm.module.scss';
import { IPoRequestFormProps } from './IPoRequestFormProps';

// People Picker (PnP SPFx Controls)
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/controls/peoplepicker/IPeoplePickerContext';

// SPFx HTTP
import { SPHttpClient, HttpClient } from '@microsoft/sp-http';

// PnPjs v3
import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';

// PDF
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import autoTable from 'jspdf-autotable';

// LOGO
import flLogo from '../assets/fl_logo.jpg';

/* ==================== TIPOS ==================== */
type Priority = 'Normal' | 'Urgent';
type RequestType = 'Goods' | 'Services' | 'Software/License';
type RoleKey = 'requester' | 'supervisor' | 'staffManager' | 'staffManager2' | 'manager' | 'manager2' | 'director' | 'vp' | 'cfo' | 'ceo' | 'procurement' | 'finance';
type ApprovalStatus = 'Agree' | 'Disagree' | 'Pending';
type BudgetType = 'Budgeted' | 'Non-Budgeted';

interface PRHeader {
  requester?: IPeoplePickerUserItem | null;
  requesterName?: string;
  requesterEmail?: string;
  requestDate?: string;
  requiredByDate?: string;
  srded?: boolean;
  cmif?: boolean;
  companyId?: number | null;
  companyValue?: string;
  area?: string;
  areaProject?: string;
  glCode?: string;
  budgetType?: BudgetType;
  priority?: Priority;
  reqType?: RequestType;
  urgentJustification?: string;

  needObjective?: string;
  impactIfNot?: string;
  soleSource?: 'Yes' | 'No';
  soleSourceExplanation?: string;
  attachQuote?: boolean;
  attachSOW?: boolean;

  supervisor?: IPeoplePickerUserItem | null;
  staffManager?: IPeoplePickerUserItem | null;
  staffManager2?: IPeoplePickerUserItem | null;
  manager?: IPeoplePickerUserItem | null;
  manager2?: IPeoplePickerUserItem | null;
  director?: IPeoplePickerUserItem | null;
  vp?: IPeoplePickerUserItem | null;
  cfo?: IPeoplePickerUserItem | null;
  ceo?: IPeoplePickerUserItem | null;
  procurement?: IPeoplePickerUserItem | null;
  finance?: IPeoplePickerUserItem | null;

  supervisorDate?: string;
  staffManagerDate?: string;
  staffManager2Date?: string;
  managerDate?: string;
  manager2Date?: string;
  directorDate?: string;
  vpDate?: string;
  cfoDate?: string;
  ceoDate?: string;
  procurementDate?: string;
  financeDate?: string;

  supervisorStatus?: ApprovalStatus;
  staffManagerStatus?: ApprovalStatus;
  staffManager2Status?: ApprovalStatus;
  managerStatus?: ApprovalStatus;
  manager2Status?: ApprovalStatus;
  directorStatus?: ApprovalStatus;
  vpStatus?: ApprovalStatus;
  cfoStatus?: ApprovalStatus;
  ceoStatus?: ApprovalStatus;
  procurementStatus?: ApprovalStatus;
  financeStatus?: ApprovalStatus;

  poNumber?: string;
}

interface PRLine {
  id?: number;
  description?: string;
  sku?: string;
  qty?: number;
  uom?: string;
  unitPrice?: number;
  currency?: string;
  tax?: number;
  total?: number;
}

interface SupplierLine {
  id?: number;
  name?: string;
  contact?: string;
  email?: string;
}

interface FieldInfo {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden?: boolean;
  ReadOnlyField?: boolean;
  SchemaXml?: string;
}

interface GLCodeItem {
  Title: string;
  CostCenterName: string;
  CostCenterNumber: string;
  ActivityCodeName: string;
  ActivityCodeNumber: string;
  NaturalAccountName: string;
  NaturalAccountNumber: string;
  IsActive: boolean;
}

interface ProjectItem {
  Title: string;
  ProjectCode: string;
  ProjectDescription: string;
  IsActive: boolean;
}

interface CompanyItem {
  Id: number;
  Title: string;
  CompanyCodeforGLAccounts?: string;
  ProntoCompanyName?: string;
  CompanyName?: string;
  IsActive: boolean;
}

interface AuthorizationRule {
  Id: number;
  Title: string;
  Department: string;
  BudgetType: 'Budgeted' | 'Non-Budgeted';
  Position: string;
  Name: string;
  Email: string;
  MinLimit?: number;
  MaxLimit?: number;
  IsNil?: boolean;
  IsActive: boolean;
}

/* ==================== Constants ==================== */
const PERSON_FIELD_INTERNALS: Partial<Record<RoleKey, string>> = {
  requester: 'Requester',
  supervisor: 'Supervisor',
  staffManager: 'StaffManager',
  staffManager2: 'Staff2',
  manager: 'Manager',
  manager2: 'Manager2',
  director: 'Director',
  vp: 'VP',
  cfo: 'CFO',
  ceo: 'CEO',
  procurement: 'Procurement',
  finance: 'Finance'
};

const ROLE_PERSON_TITLE_CANDIDATES: Record<RoleKey, string[]> = {
  requester: ['Requester'],
  supervisor: ['Supervisor'],
  staffManager: ['Staff Manager', 'Staff'],
  staffManager2: ['Staff2'],
  manager: ['Manager'],
  manager2: ['Manager2'],
  director: ['Director'],
  vp: ['VP'],
  cfo: ['CFO'],
  ceo: ['CEO'],
  procurement: ['Procurement'],
  finance: ['Finance', 'Finance Final']
};

const ROLE_DATE_TITLES: Record<RoleKey, string[]> = {
  requester: [],
  supervisor: ['Supervisor Date'],
  staffManager: ['Staff Manager Date'],
  staffManager2: ['Staff2 Date'],
  manager: ['Manager Date'],
  manager2: ['Manager2 Date'],
  director: ['Director Date'],
  vp: ['VP Date'],
  cfo: ['CFO Date'],
  ceo: ['CEO Date'],
  procurement: ['Procurement Date'],
  finance: ['Finance Date', 'Finance (Final) Date']
};

const ROLE_STATUS_TITLES: Partial<Record<RoleKey, string[]>> = {
  supervisor: ['Supervisor status', 'Supervisor Status'],
  staffManager: ['Staff Manager status', 'Staff Manager Status'],
  staffManager2: ['Staff2 status', 'Staff2 Status'],
  manager: ['Manager status', 'Manager Status'],
  manager2: ['Manager2 status', 'Manager2 Status'],
  director: ['Director status', 'Director Status'],
  vp: ['VP status', 'VP Status'],
  cfo: ['CFO status', 'CFO Status'],
  ceo: ['CEO status', 'CEO Status'],
  procurement: ['Procurement status', 'Procurement Status'],
  finance: ['Finance status', 'Finance Status']
};

const CHILD_INTERNALS = {
  description: 'ItemDescription_x002f_Specificat',
  sku: 'SKU_x002f_Par',
  qty: 'Qty',
  uom: 'UoM',
  unitPrice: 'UnitPrice',
  currency: 'Currency',
  tax: 'Tax',
  total: 'Total'
};

const SUPPLIER_INTERNALS = {
  prid: 'PRId',
  name: 'SupplierLegalName',
  contact: 'SupplierContact',
  email: 'SupplierEmail'
};

const AREA_OPTIONS = [
  'Operations',
  'Geology',
  'Environment',
  'Finance',
  'Corporate Development',
  'Technology'
];

/* ==================== Utils ==================== */
const getTodayYmd = () => {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
};

const currencyFmt = (v?: number) => {
  const n = typeof v === 'number' ? v : 0;
  return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const safeNum = (v?: number | string) => {
  const n = typeof v === 'string' ? parseFloat(v) : (v ?? 0);
  return Number.isFinite(n) ? (n as number) : 0;
};

const sanitizeName = (name: string) => name.replace(/[^a-zA-Z0-9.\-_]/g, '_');

const isoToYmd = (iso?: string | null): string => {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
};

const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
const retryUntilOk = async <T,>(fn: () => Promise<T>, _label: string, max = 5, base = 800): Promise<T> => {
  let last: any;
  for (let i = 1; i <= max; i++) {
    try { return await fn(); } catch (e) { last = e; if (i < max) await sleep(base * i); }
  }
  throw last;
};

const norm = (s?: string) =>
  (s || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

const normLoose = (s?: string) =>
  (s || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '')
    .trim();

const toSharePointDateLocal = (ymd?: string) => {
  if (!ymd) return null;
  const [y, m, d] = ymd.split('-').map(Number);
  const dt = new Date(y, (m || 1) - 1, d || 1, 12, 0, 0);
  return dt.toISOString();
};

/* ==================== Component ==================== */
const PoRequestForm: React.FC<IPoRequestFormProps> = (props) => {
  const {
    context,
    parentListId,
    parentListTitle,
    childListId,
    childListTitle,
    suppliersListTitle,
    supplierChildListId,
    supplierChildListTitle,
    pageSize,
    glCodesListId,
    glCodesListTitle,
    projectsListId,
    projectsListTitle,
    companiesListId,
    companiesListTitle,
    authorizersListId,
    authorizersListTitle
  } = props;

  const sp = React.useMemo<SPFI>(() => spfi().using(SPFx(context)), [context]);

  const peopleCtx = React.useMemo<IPeoplePickerContext>(() => ({
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient
  }), [context]);

  // Tabs
  const [activeView, setActiveView] = React.useState<'new' | 'mysent' | 'tosign' | 'approved'>('new');

  // ===== State =====
  const initialHeader: PRHeader = {
    requestDate: getTodayYmd(),
    priority: 'Normal',
    reqType: 'Goods',
    budgetType: 'Budgeted',
    soleSource: 'No',
    attachQuote: true,
    attachSOW: false,
    supervisorStatus: 'Pending',
    staffManagerStatus: 'Pending',
    staffManager2Status: 'Pending',
    managerStatus: 'Pending',
    manager2Status: 'Pending',
    directorStatus: 'Pending',
    vpStatus: 'Pending',
    cfoStatus: 'Pending',
    ceoStatus: 'Pending',
    procurementStatus: 'Pending',
    financeStatus: 'Pending',
    srded: false,
    cmif: false
  };

  const [headerDraft, setHeaderDraft] = React.useState<PRHeader>(initialHeader);
  const [lines, setLines] = React.useState<PRLine[]>([{ id: 1, qty: 1, unitPrice: 0, tax: 0, currency: 'USD' }]);
  const [suppliers, setSuppliers] = React.useState<SupplierLine[]>([{ id: 1, name: '', contact: '', email: '' }]);
  const [files, setFiles] = React.useState<File[]>([]);
  const attachInputRef = React.useRef<HTMLInputElement | null>(null);

  // Lists
  const [, setMySent] = React.useState<any[]>([]);
  const [mySentWithStatus, setMySentWithStatus] = React.useState<Array<{
    item: any;
    pendingRoles: RoleKey[];
    approvedRoles: RoleKey[];
    disagreeRoles: RoleKey[];
  }>>([]);
  const [myApproved, setMyApproved] = React.useState<any[]>([]);
  const [myToSign, setMyToSign] = React.useState<Array<{ item: any; roles: RoleKey[] }>>([]);

  const [listLoading, setListLoading] = React.useState(false);
  const [editingItemId, setEditingItemId] = React.useState<number | null>(null);
  const [signFormRoles, setSignFormRoles] = React.useState<RoleKey[] | null>(null);

  const [childLookupInternal, setChildLookupInternal] = React.useState<string | null>(null);
  const [supLookupInternal, setSupLookupInternal] = React.useState<string | null>(null);
  const currentUserIdRef = React.useRef<number | null>(null);

  // Loading states
  const [submitLoading, setSubmitLoading] = React.useState(false);
  const [pdfLoading, setPdfLoading] = React.useState(false);

  // Estados para GL Codes, Projects, Companies y Authorization Rules
  const [glCodes, setGlCodes] = React.useState<GLCodeItem[]>([]);
  const [projects, setProjects] = React.useState<ProjectItem[]>([]);
  const [companies, setCompanies] = React.useState<CompanyItem[]>([]);
  const [authorizationRules, setAuthorizationRules] = React.useState<AuthorizationRule[]>([]);

  const [loadingGlCodes, setLoadingGlCodes] = React.useState<boolean>(false);
  const [loadingProjects, setLoadingProjects] = React.useState<boolean>(false);
  const [loadingCompanies, setLoadingCompanies] = React.useState<boolean>(false);
  const [, setLoadingAuthorizers] = React.useState<boolean>(false);

  // Snackbar
  const [snack, setSnack] = React.useState<{
    open: boolean;
    msg: string;
    variant: 'success' | 'error' | 'info';
  }>({ open: false, msg: '', variant: 'info' });

  const snackTimeoutRef = React.useRef<number | null>(null);

  const showSnack = React.useCallback(
    (msg: string, variant: 'success' | 'error' | 'info' = 'info', ms = 3200) => {
      setSnack({ open: true, msg, variant });

      if (snackTimeoutRef.current) {
        window.clearTimeout(snackTimeoutRef.current);
      }

      snackTimeoutRef.current = window.setTimeout(() => {
        setSnack(s => ({ ...s, open: false }));
      }, ms) as unknown as number;
    },
    [setSnack]
  );

  // Focus keeper
  const useFocusKeeper = () => {
    const last = React.useRef<{ id?: string; start?: number | null; end?: number | null } | null>(null);
    const keepFocus = <E extends HTMLInputElement | HTMLTextAreaElement>(id: string, updater: (ev: React.ChangeEvent<E>) => void) =>
      (ev: React.ChangeEvent<E>) => {
        const el = ev.currentTarget;
        last.current = { id, start: el.selectionStart, end: el.selectionEnd };
        updater(ev);
        requestAnimationFrame(() => {
          const target = document.getElementById(id) as E | null;
          if (target) {
            try {
              target.focus();
              if (last.current?.start != null && last.current?.end != null) {
                target.setSelectionRange(last.current.start!, last.current.end!);
              }
            } catch { /* ignore */ }
          }
        });
      };
    return { keepFocus };
  };
  const { keepFocus } = useFocusKeeper();

  // Lock por aprobaciones
  const [hasAnyApproval, setHasAnyApproval] = React.useState<boolean>(false);

  const siteUrl = React.useMemo(() => {
    const s = context.pageContext.web.absoluteUrl || '';
    return s.endsWith('/') ? s.slice(0, -1) : s;
  }, [context]);

  const parentRef = React.useMemo(() => {
    if (parentListId) return { by: 'id' as const, id: parentListId };
    if (parentListTitle) return { by: 'title' as const, title: parentListTitle };
    return null;
  }, [parentListId, parentListTitle]);

  const childRef = React.useMemo(() => {
    if (childListId) return { by: 'id' as const, id: childListId };
    if (childListTitle) return { by: 'title' as const, title: childListTitle };
    return null;
  }, [childListId, childListTitle]);

  /** Lista de proveedores sugeridos (B) */
  const suppliersRef = React.useMemo(() => {
    if (supplierChildListId) return { by: 'id' as const, id: supplierChildListId };
    if (supplierChildListTitle) return { by: 'title' as const, title: supplierChildListTitle };
    return null;
  }, [supplierChildListId, supplierChildListTitle]);

  const parentListEscName = React.useMemo(() => (parentListTitle || '').replace(/'/g, "''"), [parentListTitle]);
  const childListEscName = React.useMemo(() => (childListTitle || '').replace(/'/g, "''"), [childListTitle]);
  const suppliersListEscName = React.useMemo(
    () => (supplierChildListTitle || '').replace(/'/g, "''"),
    [supplierChildListTitle]
  );

  const listNameOrIdExpr = (ref: typeof parentRef, listEscapedName: string) => {
    if (!ref) return '';
    return ref.by === 'id' ? `lists(guid'${ref.id}')` : `lists/getByTitle('${listEscapedName}')`;
  };

  /* ====== REST helpers ====== */
  const spGet = React.useCallback(
    async (url: string, extraHeaders?: Record<string, string>): Promise<any> => {
      const accepts = [
        'application/json;odata=verbose',
        'application/json;odata=nometadata',
        'application/json;odata.metadata=none'
      ];
      let last = '';

      for (const accept of accepts) {
        try {
          const r = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1,
            {
              headers: { Accept: accept, ...(extraHeaders || {}) }
            }
          );

          const txt = await r.text();

          if (!r.ok) {
            last = txt || r.statusText || '';
            if (r.status === 406) {
              continue;
            }
            throw new Error(last);
          }

          if (!txt) return {};

          try {
            const j = JSON.parse(txt);
            if (j && (j as any).d !== undefined) {
              const d: any = (j as any).d;
              if (d && d.results !== undefined) {
                return { value: d.results };
              }
              return d;
            }
            return j;
          } catch {
            return {};
          }
        } catch (e: any) {
          last = e.message || String(e);
        }
      }

      throw new Error(last || 'GET failed');
    },
    [context.spHttpClient]
  );

  const getFormDigest = async (): Promise<string> => {
    const headers = [
      'application/json;odata=verbose',
      'application/json',
      '*/*'
    ];
    let last = '';
    for (const accept of headers) {
      try {
        const r = await context.spHttpClient.post(`${siteUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {
          headers: { 'Accept': accept, 'Content-Type': 'application/json;odata=verbose' }
        });
        if (!r.ok) { last = await r.text(); continue; }
        const txt = await r.text();
        let j: any; try { j = JSON.parse(txt); } catch { last = 'Invalid JSON'; continue; }
        const v = j?.d?.GetContextWebInformation?.FormDigestValue || j?.GetContextWebInformation?.FormDigestValue || j?.FormDigestValue;
        if (v) return v;
        last = 'No digest';
      } catch (e: any) { last = e.message; }
    }
    throw new Error(`Digest error: ${last}`);
  };

  const uploadAttachmentOnce = async (listNameEscaped: string, itemId: number, fileName: string, buffer: ArrayBuffer): Promise<void> => {
    const url = `${siteUrl}/_api/web/lists/getByTitle('${listNameEscaped}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`;
    let resp = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata, */*',
        'Content-Type': 'application/octet-stream'
      },
      body: buffer
    });
    const txt = await resp.text();
    if (resp.ok) return;

    const isSecurityValidation = resp.status === 403 && /security validation/i.test(txt);
    if (!isSecurityValidation) throw new Error(`Upload failed: ${txt}`);

    const digest = await getFormDigest();
    const r2 = await context.httpClient.post(url, HttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata, */*',
        'Content-Type': 'application/octet-stream',
        'X-RequestDigest': digest
      },
      body: buffer
    });
    if (!r2.ok) throw new Error(`Upload failed: ${await r2.text()}`);
  };

  const getChildFieldIndex = async (): Promise<{ byTitle: Map<string, FieldInfo>, byInternal: Map<string, FieldInfo> }> => {
    if (!childRef) return { byTitle: new Map(), byInternal: new Map() };
    const url = childRef.by === 'id'
      ? `${siteUrl}/_api/web/lists(guid'${childRef.id}')/fields?$select=InternalName,Title,TypeAsString`
      : `${siteUrl}/_api/web/lists/getByTitle('${childListEscName}')/fields?$select=InternalName,Title,TypeAsString`;

    const js = await spGet(url);
    const fields = (js.value || js) as FieldInfo[];
    const byTitle = new Map<string, FieldInfo>();
    const byInternal = new Map<string, FieldInfo>();
    for (const f of fields) {
      byTitle.set(norm(f.Title), f);
      byInternal.set(f.InternalName, f);
    }
    return { byTitle, byInternal };
  };

  const getSuppliersFieldIndex = async (): Promise<{ byTitle: Map<string, FieldInfo>, byInternal: Map<string, FieldInfo> }> => {
    if (!suppliersRef) return { byTitle: new Map(), byInternal: new Map() };
    const url = suppliersRef.by === 'id'
      ? `${siteUrl}/_api/web/lists(guid'${suppliersRef.id}')/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField,SchemaXml`
      : `${siteUrl}/_api/web/lists/getByTitle('${suppliersListEscName}')/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField,SchemaXml`;

    const js = await spGet(url);
    const fields = (js.value || js) as FieldInfo[];
    const byTitle = new Map<string, FieldInfo>();
    const byInternal = new Map<string, FieldInfo>();
    for (const f of fields) {
      byTitle.set(norm(f.Title), f);
      byInternal.set(f.InternalName, f);
    }
    return { byTitle, byInternal };
  };

  const childNameOfField = (
    dex: { byTitle: Map<string, FieldInfo>, byInternal: Map<string, FieldInfo> },
    titleOrInternal: string,
    overrideInternal?: string
  ): string | null => {
    if (overrideInternal && dex.byInternal.has(overrideInternal)) return overrideInternal;
    if (dex.byInternal.has(titleOrInternal)) return titleOrInternal;
    const f = dex.byTitle.get(norm(titleOrInternal));
    if (f) return f.InternalName;
    return null;
  };

  const getParentFieldIndex = async (): Promise<{ byTitle: Map<string, FieldInfo>, byInternal: Map<string, FieldInfo> }> => {
    const url = parentRef?.by === 'id'
      ? `${siteUrl}/_api/web/lists(guid'${parentRef.id}')/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField,SchemaXml`
      : `${siteUrl}/_api/web/lists/getByTitle('${parentListEscName}')/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField,SchemaXml`;

    const js = await spGet(url);
    const fields = (js.value || js) as FieldInfo[];
    const byTitle = new Map<string, FieldInfo>();
    const byInternal = new Map<string, FieldInfo>();
    for (const f of fields) {
      byTitle.set(norm(f.Title), f);
      byInternal.set(f.InternalName, f);
    }
    return { byTitle, byInternal };
  };

  const ensureUserIdSmart = React.useCallback(async (loginOrEmail?: string | null): Promise<number | null> => {
    if (!loginOrEmail) return null;
    const val = loginOrEmail.trim();
    const candidates: string[] = [val];
    if (!val.includes('|')) candidates.push(`i:0#.f|membership|${val}`);

    for (const logonName of candidates) {
      try {
        const r: any = await sp.web.ensureUser(logonName);
        const id = r?.data?.Id ?? r?.Id ?? null;
        if (id) return id;
      } catch { /* ignore */ }
    }

    try {
      const resp = await context.spHttpClient.post(`${siteUrl}/_api/web/ensureuser`, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: JSON.stringify({ logonName: candidates[candidates.length - 1] })
      });
      const txt = await resp.text();
      if (resp.ok) {
        const j = JSON.parse(txt);
        const id = j?.d?.Id ?? j?.Id ?? null;
        if (id) return id;
      }
    } catch { /* ignore */ }

    try {
      const esc = val.replace(/'/g, "''");
      const users = await sp.web.siteUsers.filter(`Email eq '${esc}'`)();
      if (users?.length) return users[0].Id;
    } catch { /* ignore */ }

    return null;
  }, [sp, context.spHttpClient, siteUrl]);

  const detectChildLookupInternal = React.useCallback(async (): Promise<string | null> => {
    if (!childRef || !parentRef) return null;

    const parent = parentRef.by === 'id'
      ? await spGet(`${siteUrl}/_api/web/lists(guid'${parentRef.id}')?$select=Id,Title`)
      : await spGet(`${siteUrl}/_api/web/lists/getByTitle('${parentListEscName}')?$select=Id,Title`);
    const parentGuid: string = (parent?.Id || '').toLowerCase();

    const childUrl = childRef.by === 'id'
      ? `${siteUrl}/_api/web/lists(guid'${childRef.id}')/fields?$select=InternalName,Title,TypeAsString,SchemaXml`
      : `${siteUrl}/_api/web/lists/getByTitle('${childListEscName}')/fields?$select=InternalName,Title,TypeAsString,SchemaXml`;

    const js = await spGet(childUrl);
    const fields = (js.value || js) as FieldInfo[];

    for (const f of fields) {
      if (f.TypeAsString === 'Lookup' && f.SchemaXml) {
        const xml = f.SchemaXml.toLowerCase();
        if (xml.includes(`list="{${parentGuid}}"`) || xml.includes(`list='${parentGuid}'`) || xml.includes(`list="{${parentGuid}}`)) {
          return f.InternalName;
        }
      }
    }
    const guess = fields.find(f => f.TypeAsString === 'Lookup' && /request|parent|po|header/i.test(f.InternalName));
    return guess?.InternalName || null;
  }, [childRef, parentRef, siteUrl, parentListEscName, childListEscName, spGet]);

  const detectSuppliersLookupInternal = React.useCallback(async (): Promise<string | null> => {
    console.log('\n✅ Suppliers usando PRId field (sin lookup detection)');
    return null;
  }, []);

  // Funciones para cargar las nuevas listas
  const loadAuthorizationRules = React.useCallback(async () => {
    if (!authorizersListId && !authorizersListTitle) {
      console.warn('⚠️ Authorizers list not configured');
      return;
    }

    try {
      setLoadingAuthorizers(true);

      let url: string;
      if (authorizersListId) {
        url = `${siteUrl}/_api/web/lists(guid'${authorizersListId}')/items?$select=Id,Title,Department,BudgetType,Position,Name,Email,MinLimit,MaxLimit,IsNil,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      } else {
        const listEscaped = authorizersListTitle!.replace(/'/g, "''");
        url = `${siteUrl}/_api/web/lists/getByTitle('${listEscaped}')/items?$select=Id,Title,Department,BudgetType,Position,Name,Email,MinLimit,MaxLimit,IsNil,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      }

      const js = await spGet(url);
      const items = ((js as any).value || js || []) as AuthorizationRule[];
      setAuthorizationRules(items);
      console.log(`✅ Loaded ${items.length} active Authorization Rules`);
    } catch (err: any) {
      console.error('❌ Error loading Authorization Rules:', err);
      setAuthorizationRules([]);
      showSnack(`Error loading Authorizers: ${err.message || err}`, 'error');
    } finally {
      setLoadingAuthorizers(false);
    }
  }, [authorizersListId, authorizersListTitle, siteUrl, spGet, showSnack]);

  const getRequiredAuthorizers = React.useCallback(
  (area: string, budgetType: BudgetType, totalAmount: number): RoleKey[] => {
    const required: RoleKey[] = ['supervisor'];

    if (!area || !budgetType) return required;

    const applicableRules = authorizationRules.filter(
      rule => norm(rule.Department) === norm(area) &&
              rule.BudgetType === budgetType &&
              rule.IsActive
    );

    for (const rule of applicableRules) {
      let matches = false;

      if (rule.IsNil) {
        matches = true;
      } else if (rule.MinLimit != null && rule.MaxLimit != null) {
        matches = totalAmount >= rule.MinLimit && totalAmount <= rule.MaxLimit;
      } else if (rule.MinLimit != null) {
        matches = totalAmount >= rule.MinLimit;
      } else if (rule.MaxLimit != null) {
        matches = totalAmount <= rule.MaxLimit;
      }

      if (matches) {
        const position = norm(rule.Position);
        // Staff/Staff2
        if (position.includes('staff2')) {
          required.push('staffManager2');
        } else if (position.includes('staff') && !position.includes('staff2')) {
          required.push('staffManager');
        }
        // Manager/Manager2
        if (position.includes('manager2')) {
          required.push('manager2');
        } else if (position.includes('manager') && !position.includes('manager2')) {
          required.push('manager');
        }
        if (position.includes('director')) required.push('director');
        if (position.includes('vp')) required.push('vp');
        if (position.includes('cfo')) required.push('cfo');
        if (position.includes('ceo')) required.push('ceo');
        if (position.includes('procurement')) required.push('procurement');
        if (position.includes('finance')) required.push('finance');
      }
    }

    return Array.from(new Set(required));
  },
  [authorizationRules]
);

 const autoPopulateApprovers = React.useCallback(() => {
  if (!headerDraft.area || !headerDraft.budgetType) {
    showSnack('Please select Area and Budget Type first', 'error');
    return;
  }

  const totalAmount = lines.reduce((s, l) => s + safeNum(l.qty) * safeNum(l.unitPrice), 0);
  const requiredRoles = getRequiredAuthorizers(headerDraft.area, headerDraft.budgetType, totalAmount);

  // Resetear todos los aprobadores excepto supervisor
  const resetApprovers: Partial<PRHeader> = {
    staffManager: null,
    manager: null,
    director: null,
    vp: null,
    cfo: null,
    ceo: null,
    procurement: null,
    finance: null
  };
  setHeaderDraft(prev => ({ ...prev, ...resetApprovers }));

  // Mostrar solo los aprobadores requeridos
  const rolesToShow: RoleKey[] = ['supervisor', ...requiredRoles.filter(r => r !== 'supervisor')];

  // Filtrar las reglas de autorización para el área y tipo de presupuesto
  const applicableRules = authorizationRules.filter(
    rule => norm(rule.Department) === norm(headerDraft.area) &&
            rule.BudgetType === headerDraft.budgetType &&
            rule.IsActive
  );

  // Mapeo de roles a palabras clave en Position
  const positionKeywords: Record<Exclude<RoleKey, 'requester'>, string[]> = {
    supervisor: ['supervisor'],
    staffManager: ['staff'],
    staffManager2: ['staff'],
    manager: ['manager'],
    manager2: ['manager'],
    director: ['director'],
    vp: ['vp'],
    cfo: ['cfo'],
    ceo: ['ceo'],
    procurement: ['procurement'],
    finance: ['finance']
  };

  // Para cada rol requerido, buscar en las reglas
  for (const role of rolesToShow) {
    if (role === 'requester' || role === 'supervisor') continue;

    const keywords = positionKeywords[role];
    let foundRule: AuthorizationRule | null = null;

    // Buscar la regla que coincida con las palabras clave del rol
    for (const rule of applicableRules) {
      const posNorm = norm(rule.Position);
      let matches = false;

      // Lógica especial para Staff/Staff2 y Manager/Manager2
      if (role === 'staffManager') {
        // Busca "staff" pero NO "staff2"
        matches = posNorm.includes('staff') && !posNorm.includes('staff2');
      } else if (role === 'staffManager2') {
        // Busca específicamente "staff2"
        matches = posNorm.includes('staff2');
      } else if (role === 'manager') {
        // Busca "manager" pero NO "manager2"
        matches = posNorm.includes('manager') && !posNorm.includes('manager2');
      } else if (role === 'manager2') {
        // Busca específicamente "manager2"
        matches = posNorm.includes('manager2');
      } else {
        matches = keywords.every(kw => posNorm.includes(kw));
      }

      if (matches) {
        foundRule = rule;
        break;
      }
    }

    if (foundRule) {
      const pickerItem: IPeoplePickerUserItem = {
        id: String(foundRule.Id),
        loginName: foundRule.Email,
        text: foundRule.Name,
        secondaryText: foundRule.Email,
        imageUrl: '',
        imageInitials: '',
        tertiaryText: '',
        optionalText: ''
      };

      setHeaderDraft(prev => ({
        ...prev,
        [role]: pickerItem
      }));
    }
  }
}, [headerDraft.area, headerDraft.budgetType, lines, getRequiredAuthorizers, authorizationRules]);
  
  const loadGlCodes = React.useCallback(async () => {
    if (!glCodesListId && !glCodesListTitle) return;
    try {
      setLoadingGlCodes(true);
      let url: string;
      if (glCodesListId) {
        url = `${siteUrl}/_api/web/lists(guid'${glCodesListId}')/items?$select=Title,CostCenterName,CostCenterNumber,ActivityCodeName,ActivityCodeNumber,NaturalAccountName,NaturalAccountNumber,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      } else {
        const listEscaped = glCodesListTitle!.replace(/'/g, "''");
        url = `${siteUrl}/_api/web/lists/getByTitle('${listEscaped}')/items?$select=Title,CostCenterName,CostCenterNumber,ActivityCodeName,ActivityCodeNumber,NaturalAccountName,NaturalAccountNumber,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      }
      const js = await spGet(url);
      const items = ((js as any).value || js || []) as GLCodeItem[];
      setGlCodes(items);
    } catch (err: any) {
      setGlCodes([]);
      showSnack(`Error loading GL Codes: ${err.message}`, 'error');
    } finally {
      setLoadingGlCodes(false);
    }
  }, [glCodesListId, glCodesListTitle, siteUrl, spGet, showSnack]);

  const loadProjects = React.useCallback(async () => {
    if (!projectsListId && !projectsListTitle) return;
    try {
      setLoadingProjects(true);
      let url: string;
      if (projectsListId) {
        url = `${siteUrl}/_api/web/lists(guid'${projectsListId}')/items?$select=Title,ProjectCode,ProjectDescription,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      } else {
        const listEscaped = projectsListTitle!.replace(/'/g, "''");
        url = `${siteUrl}/_api/web/lists/getByTitle('${listEscaped}')/items?$select=Title,ProjectCode,ProjectDescription,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      }
      const js = await spGet(url);
      const items = ((js as any).value || js || []) as ProjectItem[];
      setProjects(items);
    } catch (err: any) {
      setProjects([]);
      showSnack(`Error loading Projects: ${err.message}`, 'error');
    } finally {
      setLoadingProjects(false);
    }
  }, [projectsListId, projectsListTitle, siteUrl, spGet, showSnack]);

  const loadCompanies = React.useCallback(async () => {
    if (!companiesListId && !companiesListTitle) return;
    try {
      setLoadingCompanies(true);
      let url: string;
      if (companiesListId) {
        url = `${siteUrl}/_api/web/lists(guid'${companiesListId}')/items?$select=Id,Title,CompanyCodeforGLAccounts,ProntoCompanyName,CompanyName,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      } else {
        const listEscaped = companiesListTitle!.replace(/'/g, "''");
        url = `${siteUrl}/_api/web/lists/getByTitle('${listEscaped}')/items?$select=Id,Title,CompanyCodeforGLAccounts,ProntoCompanyName,CompanyName,IsActive&$filter=IsActive eq true&$orderby=Title asc`;
      }
      const js = await spGet(url);
      const items = ((js as any).value || js || []) as CompanyItem[];
      setCompanies(items);
    } catch (err: any) {
      setCompanies([]);
    } finally {
      setLoadingCompanies(false);
    }
  }, [companiesListId, companiesListTitle, siteUrl, spGet]);

  // ============================================
  // Dropdown "table" (search + pagination)
  // GL Code + Project + Company
  // ============================================
  const GL_DD_PAGE_SIZE = 3;
  const PROJECT_DD_PAGE_SIZE = 3;
  const COMPANY_DD_PAGE_SIZE = 3;

  // ===================== GL Code dropdown state =====================
  const [glOpen, setGlOpen] = React.useState<boolean>(false);
  const [glSearch, setGlSearch] = React.useState<string>("");
  const [glSearchDebounced, setGlSearchDebounced] = React.useState<string>("");
  const [glPage, setGlPage] = React.useState<number>(1);
  const [glOpenUp, setGlOpenUp] = React.useState<boolean>(false);
  const [glNoPagination, setGlNoPagination] = React.useState<boolean>(false);
  const glWrapRef = React.useRef<HTMLDivElement>(null);
  const glInputRef = React.useRef<HTMLInputElement>(null);

  // ===================== Project dropdown state =====================
  const [projOpen, setProjOpen] = React.useState<boolean>(false);
  const [projSearch, setProjSearch] = React.useState<string>("");
  const [projSearchDebounced, setProjSearchDebounced] = React.useState<string>("");
  const [projPage, setProjPage] = React.useState<number>(1);
  const [projOpenUp, setProjOpenUp] = React.useState<boolean>(false);
  const projWrapRef = React.useRef<HTMLDivElement>(null);
  const projInputRef = React.useRef<HTMLInputElement>(null);

  // ===================== Company dropdown state =====================
  const [companyOpen, setCompanyOpen] = React.useState<boolean>(false);
  const [companySearch, setCompanySearch] = React.useState<string>("");
  const [companySearchDebounced, setCompanySearchDebounced] = React.useState<string>("");
  const [companyPage, setCompanyPage] = React.useState<number>(1);
  const [companyOpenUp, setCompanyOpenUp] = React.useState<boolean>(false);
  const companyWrapRef = React.useRef<HTMLDivElement>(null);
  const companyInputRef = React.useRef<HTMLInputElement>(null);

  // Debounce search to avoid expensive filtering while typing
  React.useEffect(() => {
    const t = window.setTimeout(() => {
      setGlSearchDebounced(glSearch);
    }, 180);
    return () => window.clearTimeout(t);
  }, [glSearch]);

  React.useEffect(() => {
    const t = window.setTimeout(() => {
      setProjSearchDebounced(projSearch);
    }, 180);
    return () => window.clearTimeout(t);
  }, [projSearch]);

  React.useEffect(() => {
    const t = window.setTimeout(() => {
      setCompanySearchDebounced(companySearch);
    }, 180);
    return () => window.clearTimeout(t);
  }, [companySearch]);

  // Decide if dropdown should open upwards (when close to the bottom of the viewport)
  React.useLayoutEffect(() => {
    if (!glOpen) return;

    const compute = () => {
      const el = glInputRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      const desired = 360;
      const below = window.innerHeight - rect.bottom;
      const above = rect.top;
      setGlOpenUp(below < desired && above > below);
    };

    compute();
    window.addEventListener('resize', compute);
    window.addEventListener('scroll', compute, true);

    return () => {
      window.removeEventListener('resize', compute);
      window.removeEventListener('scroll', compute, true);
    };
  }, [glOpen]);

  React.useLayoutEffect(() => {
    if (!projOpen) return;

    const compute = () => {
      const el = projInputRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      const desired = 360;
      const below = window.innerHeight - rect.bottom;
      const above = rect.top;
      setProjOpenUp(below < desired && above > below);
    };

    compute();
    window.addEventListener('resize', compute);
    window.addEventListener('scroll', compute, true);

    return () => {
      window.removeEventListener('resize', compute);
      window.removeEventListener('scroll', compute, true);
    };
  }, [projOpen]);

  React.useLayoutEffect(() => {
    if (!companyOpen) return;

    const compute = () => {
      const el = companyInputRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      const desired = 360;
      const below = window.innerHeight - rect.bottom;
      const above = rect.top;
      setCompanyOpenUp(below < desired && above > below);
    };

    compute();
    window.addEventListener('resize', compute);
    window.addEventListener('scroll', compute, true);

    return () => {
      window.removeEventListener('resize', compute);
      window.removeEventListener('scroll', compute, true);
    };
  }, [companyOpen]);

  // ===================== GL Code dropdown logic =====================
  const glSelectedItem = React.useMemo<GLCodeItem | null>(() => {
    const code = (headerDraft.glCode || '').trim();
    if (!code) return null;
    return glCodes.find(g => (g.Title || '').trim() === code) || null;
  }, [glCodes, headerDraft.glCode]);

  const glDisplayValue = React.useMemo(() => {
    if (!glSelectedItem) return '';
    const parts = [
      (glSelectedItem.CostCenterName || '').trim(),
      (glSelectedItem.ActivityCodeName || '').trim(),
      (glSelectedItem.NaturalAccountName || '').trim(),
    ].filter(Boolean);

    return `${glSelectedItem.Title}${parts.length ? ` | ${parts.join(' | ')}` : ''}`;
  }, [glSelectedItem]);

  const glFiltered = React.useMemo(() => {
    const q = norm(glSearchDebounced);
    const src = glCodes.filter(g => !!g.IsActive);
    if (!q) return src;

    return src.filter(g => {
      const hay = [
        g.Title,
        g.CostCenterName,
        g.CostCenterNumber,
        g.ActivityCodeName,
        g.ActivityCodeNumber,
        g.NaturalAccountName,
        g.NaturalAccountNumber
      ].filter(Boolean).join(' ');
      return norm(hay).includes(q);
    });
  }, [glCodes, glSearchDebounced]);

  const glTotalPages = React.useMemo(
    () => Math.max(1, Math.ceil(glFiltered.length / GL_DD_PAGE_SIZE)),
    [glFiltered.length]
  );

  const glPaged = React.useMemo(() => {
    if (glNoPagination) {
      return glFiltered;
    } else {
      const safePage = Math.min(Math.max(1, glPage), glTotalPages);
      const start = (safePage - 1) * GL_DD_PAGE_SIZE;
      return glFiltered.slice(start, start + GL_DD_PAGE_SIZE);
    }
  }, [glFiltered, glPage, glTotalPages, glNoPagination]);

  React.useEffect(() => {
    if (glPage > glTotalPages) setGlPage(glTotalPages);
  }, [glPage, glTotalPages]);

  React.useEffect(() => {
    if (!glOpen) return;
    setGlPage(1);
    requestAnimationFrame(() => glInputRef.current?.focus({ preventScroll: true }));
  }, [glSearchDebounced, glOpen]);

  const pickGl = React.useCallback((item: GLCodeItem) => {
    setHeaderDraft(prev => ({ ...prev, glCode: item.Title }));
    setGlOpen(false);
    setGlSearch('');
    setGlNoPagination(false);
  }, []);

  // ===================== Project dropdown logic =====================
  const projSelectedItem = React.useMemo<ProjectItem | null>(() => {
    const title = (headerDraft.areaProject || '').trim();
    if (!title) return null;
    return projects.find(p => (p.Title || '').trim() === title) || null;
  }, [projects, headerDraft.areaProject]);

  const projDisplayValue = React.useMemo(() => {
    if (!projSelectedItem) return '';
    const code = (projSelectedItem.ProjectCode || '').trim();
    const desc = (projSelectedItem.ProjectDescription || '').trim();
    return `${projSelectedItem.Title}${code || desc ? ` | ${code} | ${desc}` : ''}`;
  }, [projSelectedItem]);

  const projFiltered = React.useMemo(() => {
    const q = norm(projSearchDebounced);
    const src = projects.filter(p => !!p.IsActive);
    if (!q) return src;

    return src.filter(p => {
      const hay = [p.Title, p.ProjectCode, p.ProjectDescription].filter(Boolean).join(' ');
      return norm(hay).includes(q);
    });
  }, [projects, projSearchDebounced]);

  const projTotalPages = React.useMemo(
    () => Math.max(1, Math.ceil(projFiltered.length / PROJECT_DD_PAGE_SIZE)),
    [projFiltered.length]
  );

  const projPaged = React.useMemo(() => {
    const safePage = Math.min(Math.max(1, projPage), projTotalPages);
    const start = (safePage - 1) * PROJECT_DD_PAGE_SIZE;
    return projFiltered.slice(start, start + PROJECT_DD_PAGE_SIZE);
  }, [projFiltered, projPage, projTotalPages]);

  React.useEffect(() => {
    if (projPage > projTotalPages) setProjPage(projTotalPages);
  }, [projPage, projTotalPages]);

  React.useEffect(() => {
    if (!projOpen) return;
    setProjPage(1);
    requestAnimationFrame(() => projInputRef.current?.focus({ preventScroll: true }));
  }, [projSearchDebounced, projOpen]);

  const pickProject = React.useCallback((item: ProjectItem) => {
    setHeaderDraft(prev => ({ ...prev, areaProject: item.Title }));
    setProjOpen(false);
    setProjSearch('');
  }, []);

  const clearProject = React.useCallback(() => {
    setHeaderDraft(prev => ({ 
      ...prev, 
      areaProject: ''
    }));
    setProjSearch('');
  }, []);

  // ===================== Company dropdown logic =====================
  const companySelectedItem = React.useMemo<CompanyItem | null>(() => {
    const company = (headerDraft.companyValue || '').trim();
    if (!company) return null;
    return companies.find(c => (c.CompanyCodeforGLAccounts || '').trim() === company) || null;
  }, [companies, headerDraft.companyValue]);

  const companyDisplayValue = React.useMemo(() => {
    if (!companySelectedItem) return '';
    const code = (companySelectedItem.CompanyCodeforGLAccounts || '').trim();
    const prontoName = (companySelectedItem.ProntoCompanyName || '').trim();
    const companyName = (companySelectedItem.CompanyName || '').trim();
    return `${companySelectedItem.CompanyCodeforGLAccounts}${code || prontoName || companyName ? ` | ${code} | ${prontoName} | ${companyName}` : ''}`;
  }, [companySelectedItem]);

  const companyFiltered = React.useMemo(() => {
    const q = norm(companySearchDebounced);
    const src = companies.filter(c => !!c.IsActive);
    if (!q) return src;

    return src.filter(c => {
      const hay = [c.Title, c.CompanyCodeforGLAccounts, c.ProntoCompanyName, c.CompanyName].filter(Boolean).join(' ');
      return norm(hay).includes(q);
    });
  }, [companies, companySearchDebounced]);

  const companyTotalPages = React.useMemo(
    () => Math.max(1, Math.ceil(companyFiltered.length / COMPANY_DD_PAGE_SIZE)),
    [companyFiltered.length]
  );

  const companyPaged = React.useMemo(() => {
    const safePage = Math.min(Math.max(1, companyPage), companyTotalPages);
    const start = (safePage - 1) * COMPANY_DD_PAGE_SIZE;
    return companyFiltered.slice(start, start + COMPANY_DD_PAGE_SIZE);
  }, [companyFiltered, companyPage, companyTotalPages]);

  React.useEffect(() => {
    if (companyPage > companyTotalPages) setCompanyPage(companyTotalPages);
  }, [companyPage, companyTotalPages]);

  React.useEffect(() => {
    if (!companyOpen) return;
    setCompanyPage(1);
    requestAnimationFrame(() => companyInputRef.current?.focus({ preventScroll: true }));
  }, [companySearchDebounced, companyOpen]);

  const pickCompany = React.useCallback((item: CompanyItem) => {
    setHeaderDraft(prev => ({ 
      ...prev, 
      companyValue: item.CompanyCodeforGLAccounts,
      companyId: item.Id 
    }));
    setCompanyOpen(false);
    setCompanySearch('');
  }, []);

  const clearCompany = React.useCallback(() => {
    setHeaderDraft(prev => ({ 
      ...prev, 
      companyValue: '',
      companyId: null
    }));
    setCompanySearch('');
  }, []);

  React.useEffect(() => {
    let isMounted = true;

    const onMouseDown = (e: MouseEvent) => {
      const t = e.target as Node;
      const glWrap = glWrapRef.current;
      const pjWrap = projWrapRef.current;
      const compWrap = companyWrapRef.current;

      const inGl = !!glWrap && glWrap.contains(t);
      const inPj = !!pjWrap && pjWrap.contains(t);
      const inComp = !!compWrap && compWrap.contains(t);

      if (!inGl) {
        setGlOpen(false);
        setGlSearch('');
        setGlNoPagination(false);
      }
      if (!inPj) {
        setProjOpen(false);
        setProjSearch('');
      }
      if (!inComp) {
        setCompanyOpen(false);
        setCompanySearch('');
      }
    };

    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      setGlOpen(false);
      setGlSearch('');
      setGlNoPagination(false);
      setProjOpen(false);
      setProjSearch('');
      setCompanyOpen(false);
      setCompanySearch('');
    };

    document.addEventListener('mousedown', onMouseDown);
    window.addEventListener('keydown', onKeyDown);

    (async () => {
      try {
        const me = await sp.web.currentUser();
        if (!isMounted) return;
        currentUserIdRef.current = me?.Id ?? null;
      } catch (err) {
        if (isMounted) currentUserIdRef.current = null;
      }

      try {
        const lk = await detectChildLookupInternal();
        if (!isMounted) return;
        setChildLookupInternal(lk);
      } catch (err: any) {
        if (isMounted) setChildLookupInternal(null);
      }

      try {
        const lk2 = await detectSuppliersLookupInternal();
        if (!isMounted) return;
        setSupLookupInternal(lk2);
      } catch (err: any) {
        if (isMounted) setSupLookupInternal(null);
      }

      try {
        await loadGlCodes();
        await loadProjects();
        await loadCompanies();
        await loadAuthorizationRules();
      } catch (err) {
        console.error('❌ Error loading initial data:', err);
      }
    })();

    return () => {
      isMounted = false;
      document.removeEventListener('mousedown', onMouseDown);
      window.removeEventListener('keydown', onKeyDown);
    };
  }, [sp, detectChildLookupInternal, detectSuppliersLookupInternal, loadGlCodes, loadProjects, loadCompanies, loadAuthorizationRules]);

  /* ====== name-of-field helpers ====== */
  const buildNameOfField = (dex: { byTitle: Map<string, FieldInfo>, byInternal: Map<string, FieldInfo> }) => {
    const allTitles = Array.from(dex.byTitle.keys());
    const allInternals = Array.from(dex.byInternal.keys());

    return (labelOrInternal: string | null | undefined) => {
      if (!labelOrInternal) return null;
      const want = String(labelOrInternal);

      const hit = dex.byTitle.get(norm(want));
      if (hit) return hit.InternalName;

      const i2 = allInternals.find(i => i.toLowerCase() === want.toLowerCase());
      if (i2) return i2;

      const wantLoose = normLoose(want);
      const byLoose = allTitles.find(t => normLoose(t) === wantLoose);
      if (byLoose) return dex.byTitle.get(byLoose)!.InternalName;

      return null;
    };
  };

  const nameOfFirstExisting = (
    dex: { byTitle: Map<string, FieldInfo>, byInternal: Map<string, FieldInfo> },
    candidates: string[]
  ): string | null => {
    const nf = buildNameOfField(dex);
    for (const c of candidates) {
      const i = nf(c);
      if (i) return i;
    }
    return null;
  };

  const getPersonInternalForRole = async (role: RoleKey): Promise<string | null> => {
    const dex = await getParentFieldIndex();

    const forced = PERSON_FIELD_INTERNALS[role];
    if (forced && dex.byInternal.has(forced)) return forced;

    const fromTitle = nameOfFirstExisting(dex, ROLE_PERSON_TITLE_CANDIDATES[role]);
    if (fromTitle) return fromTitle;

    const personFields = Array.from(dex.byInternal.values()).filter(f =>
      f.TypeAsString === 'User' || f.TypeAsString === 'UserMulti'
    );

    for (const field of personFields) {
      const titleNorm = norm(field.Title);
      const roleNorm = norm(ROLE_PERSON_TITLE_CANDIDATES[role][0]);
      if (titleNorm.includes(roleNorm) || roleNorm.includes(titleNorm)) {
        return field.InternalName;
      }
    }
    return null;
  };

  /* ============ NUEVA FUNCIÓN: Generar número de PR por compañía ============ */
  const getNextPRNumberForCompany = React.useCallback(async (companyCode: string): Promise<string> => {
    try {
      if (!companyCode || !parentRef) {
        return `R${new Date().getTime()}`;
      }
      
      // Definir rangos por compañía según las especificaciones
      const companyRanges: Record<string, { prefix: string; start: number; end: number }> = {
        'L01': { prefix: '1', start: 100000, end: 199999 },
        'L02': { prefix: '2', start: 200000, end: 299999 },
        'L03': { prefix: '3', start: 300000, end: 399999 },
        'L04': { prefix: '4', start: 400000, end: 499999 }
      };
      
      const range = companyRanges[companyCode];
      if (!range) {
        console.warn(`Company code not configured: ${companyCode}`);
        return `R${companyCode}-${new Date().getTime()}`;
      }
      
      // Buscar el último número de PR para esta compañía
      const filter = `startswith(Title, 'R${range.prefix}')`;
      const url = `${siteUrl}/_api/web/${listNameOrIdExpr(parentRef, parentListEscName)}/items?$filter=${filter}&$select=Title&$orderby=Title desc&$top=1`;
      
      const response = await spGet(url);
      const items = response.value || [];
      
      let nextNumber: number;
      
      if (items.length > 0) {
        const lastTitle = items[0].Title;
        const numberMatch = lastTitle.match(/R(\d+)/);
        
        if (numberMatch) {
          const lastNumber = parseInt(numberMatch[1], 10);
          
          if (lastNumber >= range.start && lastNumber <= range.end) {
            nextNumber = lastNumber + 1;
            if (nextNumber > range.end) {
              console.error(`Limit reached for ${companyCode}`);
              nextNumber = range.start;
            }
          } else {
            console.warn(`Number ${lastNumber} out of range for ${companyCode}`);
            nextNumber = range.start;
          }
        } else {
          nextNumber = range.start;
        }
      } else {
        nextNumber = range.start;
      }
      
      return `R${nextNumber}`;
      
    } catch (error) {
      console.error('Error in getNextPRNumberForCompany:', error);
      return `R${companyCode}-${new Date().getTime()}`;
    }
  }, [parentRef, siteUrl, parentListEscName, spGet, listNameOrIdExpr]);

  /* =================== Loads =================== */
const loadMySent = React.useCallback(async () => {
  if (!parentRef || !currentUserIdRef.current) return;
  setListLoading(true);
  try {
    const dex = await getParentFieldIndex();
    const nf = buildNameOfField(dex);
    const meId = currentUserIdRef.current!;
    const requesterInt = nf('Requester');

    const selectParts = [`Id`, `Title`, `Created`, `Author/Id`, `Author/Title`, `Author/EMail`];
    const expandParts = [`Author`];

    let filter = `Author/Id eq ${meId}`;
    if (requesterInt) {
      selectParts.push(`${requesterInt}/Id`, `${requesterInt}/Title`);
      expandParts.push(requesterInt);
      filter = `(${requesterInt}/Id eq ${meId} or Author/Id eq ${meId})`;
    }

    // Precompute internal field names for all roles
    const allRoles: RoleKey[] = ['supervisor', 'staffManager', 'manager', 'director', 'vp', 'cfo', 'ceo', 'procurement', 'finance'];

    const roleFieldMap: Record<string, {
      personInt: string | null;
      dateInt: string | null;
      statusInt: string | null;
    }> = {};

    for (const role of allRoles) {
      const personInt = nameOfFirstExisting(dex, ROLE_PERSON_TITLE_CANDIDATES[role]);
      const dateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES[role]);
      const statusInt = nameOfFirstExisting(dex, ROLE_STATUS_TITLES[role] || []);

      roleFieldMap[role] = { personInt, dateInt, statusInt };

      // Add person field to $select/$expand
      if (personInt) {
        selectParts.push(`${personInt}/Id`, `${personInt}/Title`);
        if (!expandParts.includes(personInt)) expandParts.push(personInt);
      }
      // Add date and status fields to $select
      if (dateInt) selectParts.push(dateInt);
      if (statusInt) selectParts.push(statusInt);
    }

    const url =
      `${siteUrl}/_api/web/${listNameOrIdExpr(parentRef, parentListEscName)}/items` +
      `?$select=${selectParts.join(',')}` +
      `&$expand=${expandParts.join(',')}` +
      `&$filter=${filter}` +
      `&$orderby=Created desc&$top=${pageSize ?? 25}`;

    const js = await spGet(url);
    const rawItems = (js.value || js || []) as any[];

    const itemsWithStatus: Array<{
      item: any;
      pendingRoles: RoleKey[];
      approvedRoles: RoleKey[];
      disagreeRoles: RoleKey[];
    }> = [];

    for (const item of rawItems) {
      const pendingRoles: RoleKey[] = [];
      const approvedRoles: RoleKey[] = [];
      const disagreeRoles: RoleKey[] = [];

      for (const role of allRoles) {
        const { personInt, dateInt, statusInt } = roleFieldMap[role];

        // ✅ Skip roles that have NO person assigned
        if (!personInt) continue;
        const personVal = item[personInt] || item[`${personInt}Id`];
        const hasPersonAssigned = personVal != null
          && personVal !== 0
          && personVal !== ''
          && !(typeof personVal === 'object' && personVal.Id == null);
        if (!hasPersonAssigned) continue;

        // Classify based on date and status
        const hasDate = dateInt && item[dateInt];
        const statusVal = statusInt ? item[statusInt] : undefined;

        if (hasDate) {
          if (statusVal === 'Agree') {
            approvedRoles.push(role);
          } else if (statusVal === 'Disagree') {
            disagreeRoles.push(role);
          } else {
            approvedRoles.push(role);
          }
        } else {
          if (statusVal === 'Agree') {
            approvedRoles.push(role);
          } else if (statusVal === 'Disagree') {
            disagreeRoles.push(role);
          } else {
            pendingRoles.push(role);
          }
        }
      }

      itemsWithStatus.push({
        item,
        pendingRoles,
        approvedRoles,
        disagreeRoles
      });
    }

    setMySentWithStatus(itemsWithStatus);
    setMySent(rawItems);
  } catch {
    setMySent([]);
    setMySentWithStatus([]);
  } finally {
    setListLoading(false);
  }
}, [parentRef, siteUrl, pageSize, spGet, parentListEscName]);

  // Approved
  const loadMyApproved = React.useCallback(async () => {
    if (!parentRef || !currentUserIdRef.current) return;
    setListLoading(true);
    try {
      const meId = currentUserIdRef.current!;

      const dex = await getParentFieldIndex();
      const supInt = await getPersonInternalForRole('supervisor');
      const staffManagerInt = await getPersonInternalForRole('staffManager');
      const managerInt = await getPersonInternalForRole('manager');
      const directorInt = await getPersonInternalForRole('director');
      const vpInt = await getPersonInternalForRole('vp');
      const cfoInt = await getPersonInternalForRole('cfo');
      const ceoInt = await getPersonInternalForRole('ceo');
      const procurementInt = await getPersonInternalForRole('procurement');
      const financeInt = await getPersonInternalForRole('finance');

      const supDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.supervisor);
      const staffManagerDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.staffManager);
      const managerDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.manager);
      const directorDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.director);
      const vpDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.vp);
      const cfoDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.cfo);
      const ceoDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.ceo);
      const procurementDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.procurement);
      const financeDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.finance);

      type Q = { role: RoleKey; url: string; dateInt?: string | null };

      const makeQ = (role: RoleKey, personInt?: string | null, dateInt?: string | null): Q | null => {
        if (!personInt) return null;
        const select: string[] = ['Id', 'Title', `${personInt}/Id`, `${personInt}/Title`];
        if (dateInt) select.push(dateInt);
        const url =
          `${siteUrl}/_api/web/${listNameOrIdExpr(parentRef, parentListEscName)}/items` +
          `?$select=${select.join(',')}` +
          `&$expand=${personInt}` +
          `&$filter=${personInt}/Id eq ${meId}` +
          `&$orderby=Created desc&$top=${pageSize ?? 50}`;
        return { role, url, dateInt };
      };

      const queries = [
        makeQ('supervisor', supInt, supDateInt),
        makeQ('staffManager', staffManagerInt, staffManagerDateInt),
        makeQ('manager', managerInt, managerDateInt),
        makeQ('director', directorInt, directorDateInt),
        makeQ('vp', vpInt, vpDateInt),
        makeQ('cfo', cfoInt, cfoDateInt),
        makeQ('ceo', ceoInt, ceoDateInt),
        makeQ('procurement', procurementInt, procurementDateInt),
        makeQ('finance', financeInt, financeDateInt),
      ].filter(Boolean) as Q[];

      const baskets: Record<number, any> = {};

      for (const { url, dateInt } of queries) {
        try {
          const js = await spGet(url);
          const items = (js.value || js || []) as any[];
          for (const it of items) {
            if (dateInt && !it[dateInt]) continue;
            const id = it.Id;
            if (!baskets[id]) baskets[id] = it;
          }
        } catch { /* ignore */ }
      }

      const arr = Object.values(baskets) as any[];
      arr.sort((a, b) => {
        const aC = a.Created || '';
        const bC = b.Created || '';
        if (aC > bC) return -1;
        if (aC < bC) return 1;
        return 0;
      });

      setMyApproved(arr);
    } catch {
      setMyApproved([]);
    } finally {
      setListLoading(false);
    }
  }, [parentRef, siteUrl, pageSize, spGet, parentListEscName]);

  // To Sign
  const loadMyToSign = React.useCallback(async () => {
    if (!parentRef || !currentUserIdRef.current) return;
    setListLoading(true);
    try {
      const meId = currentUserIdRef.current!;

      const dex = await getParentFieldIndex();
      const supInt = await getPersonInternalForRole('supervisor');
      const staffManagerInt = await getPersonInternalForRole('staffManager');
      const managerInt = await getPersonInternalForRole('manager');
      const directorInt = await getPersonInternalForRole('director');
      const vpInt = await getPersonInternalForRole('vp');
      const cfoInt = await getPersonInternalForRole('cfo');
      const ceoInt = await getPersonInternalForRole('ceo');
      const procurementInt = await getPersonInternalForRole('procurement');
      const financeInt = await getPersonInternalForRole('finance');

      const supDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.supervisor);
      const staffManagerDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.staffManager);
      const managerDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.manager);
      const directorDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.director);
      const vpDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.vp);
      const cfoDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.cfo);
      const ceoDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.ceo);
      const procurementDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.procurement);
      const financeDateInt = nameOfFirstExisting(dex, ROLE_DATE_TITLES.finance);

      type Q = { role: RoleKey; url: string; dateInt?: string | null };

      const makeQ = (role: RoleKey, personInt?: string | null, dateInt?: string | null): Q | null => {
        if (!personInt) return null;
        const select: string[] = ['Id', 'Title', `${personInt}/Id`, `${personInt}/Title`];
        if (dateInt) select.push(dateInt);
        const url =
          `${siteUrl}/_api/web/${listNameOrIdExpr(parentRef, parentListEscName)}/items` +
          `?$select=${select.join(',')}` +
          `&$expand=${personInt}` +
          `&$filter=${personInt}/Id eq ${meId}` +
          `&$orderby=Created desc&$top=${pageSize ?? 50}`;
        return { role, url, dateInt };
      };

      const queries = [
        makeQ('supervisor', supInt, supDateInt),
        makeQ('staffManager', staffManagerInt, staffManagerDateInt),
        makeQ('manager', managerInt, managerDateInt),
        makeQ('director', directorInt, directorDateInt),
        makeQ('vp', vpInt, vpDateInt),
        makeQ('cfo', cfoInt, cfoDateInt),
        makeQ('ceo', ceoInt, ceoDateInt),
        makeQ('procurement', procurementInt, procurementDateInt),
        makeQ('finance', financeInt, financeDateInt),
      ].filter(Boolean) as Q[];

      const baskets: Record<number, { item: any; roles: RoleKey[] }> = {};

      for (const { role, url, dateInt } of queries) {
        try {
          const js = await spGet(url);
          const items = (js.value || js || []) as any[];
          for (const it of items) {
            if (dateInt && it[dateInt]) continue;
            const id = it.Id;
            if (!baskets[id]) baskets[id] = { item: it, roles: [role] };
            else if (!baskets[id].roles.includes(role)) baskets[id].roles.push(role);
          }
        } catch { /* ignore */ }
      }
      setMyToSign(Object.values(baskets));
    } catch {
      setMyToSign([]);
    } finally {
      setListLoading(false);
    }
  }, [parentRef, siteUrl, spGet, parentListEscName, pageSize]);

  /* =================== write helpers =================== */
  const trySetUserByRole = async (
    dex: { byTitle: Map<string, FieldInfo>, byInternal: Map<string, FieldInfo> },
    role: RoleKey,
    userId: number | null,
    body: Record<string, any>
  ) => {
    if (userId == null) return;
    const iname = await getPersonInternalForRole(role);
    if (!iname) return;
    const info = dex.byInternal.get(iname);
    if (!info || info.Hidden || info.ReadOnlyField) return;
    const t = (info.TypeAsString || '').toLowerCase();

    if (t.includes('usermulti')) {
      body[`${iname}Id`] = { results: [userId] };
    } else {
      body[`${iname}Id`] = userId;
    }
  };

  /* =================== Create (PARENT + CHILD + SUPPLIERS) =================== */
  const createParentAndLines = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!parentRef || !childRef) {
      showSnack('Configure PARENT and CHILD lists in the Property Pane.', 'error');
      return;
    }

    try {
      setSubmitLoading(true);
      const dex = await getParentFieldIndex();
      const nameOfField = buildNameOfField(dex);

      const H = headerDraft;

      const requestIdUser =
        H.requester?.secondaryText ||
        H.requester?.loginName ||
        H.requester?.text ||
        null;
      const requesterId =
        (await ensureUserIdSmart(requestIdUser)) ?? currentUserIdRef.current;

      const supId = await ensureUserIdSmart(
        H.supervisor?.secondaryText || H.supervisor?.loginName || H.supervisor?.text || null
      );
      const staffManagerId = await ensureUserIdSmart(
        H.staffManager?.secondaryText || H.staffManager?.loginName || H.staffManager?.text || null
      );
      const managerId = await ensureUserIdSmart(
        H.manager?.secondaryText || H.manager?.loginName || H.manager?.text || null
      );
      const directorId = await ensureUserIdSmart(
        H.director?.secondaryText || H.director?.loginName || H.director?.text || null
      );
      const vpId = await ensureUserIdSmart(
        H.vp?.secondaryText || H.vp?.loginName || H.vp?.text || null
      );
      const cfoId = await ensureUserIdSmart(
        H.cfo?.secondaryText || H.cfo?.loginName || H.cfo?.text || null
      );
      const ceoId = await ensureUserIdSmart(
        H.ceo?.secondaryText || H.ceo?.loginName || H.ceo?.text || null
      );
      const procurementId = await ensureUserIdSmart(
        H.procurement?.secondaryText || H.procurement?.loginName || H.procurement?.text || null
      );
      const financeId = await ensureUserIdSmart(
        H.finance?.secondaryText || H.finance?.loginName || H.finance?.text || null
      );

      const canWrite = (iname: string | null): boolean => {
        if (!iname) return false;
        const info = dex.byInternal.get(iname);
        if (!info) return true;
        if (info.Hidden) return false;
        if (info.ReadOnlyField) return false;
        return true;
      };

      const body: Record<string, any> = {};
      body['Title'] = `PR - ${new Date().toISOString()}`;

      const trySetText = (titleLabel: string, v?: string) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname) && v !== undefined) body[iname] = v;
      };
      const trySetDate = (titleLabel: string, ymd?: string) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname)) {
          const iso = toSharePointDateLocal(ymd);
          if (iso) body[iname] = iso;
        }
      };
      const trySetBool = (titleLabel: string, v?: boolean) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname) && v !== undefined) body[iname] = v;
      };
      const trySetNumber = (titleLabel: string, n?: number) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname) && typeof n === 'number' && !Number.isNaN(n))
          body[iname] = n;
      };

      // A. Request Details
      await trySetUserByRole(dex, 'requester', requesterId ?? null, body);

      // Fallback: Requester Email
      if ((!H.requesterEmail || H.requesterEmail.trim() === '') && requesterId) {
        try {
          const u = await sp.web.getUserById(requesterId)();
          if (u?.Email) {
            const emailInternal = nameOfField('Requester Email');
            if (emailInternal) body[emailInternal] = u.Email;
            setHeaderDraft(prev => ({ ...prev, requesterEmail: u.Email }));
          }
        } catch {
          /* ignore */
        }
      }

      // Header fields
      trySetText('Requester Email', H.requesterEmail);
      trySetDate('Request Date', H.requestDate);
      trySetDate('Requiredbydate', H.requiredByDate);
      trySetText('Area', H.area);
      trySetText('Area / Project', H.areaProject);
      trySetText('Priority', H.priority);
      trySetText('Type', H.reqType);
      trySetText('GL Code', H.glCode);
      trySetText('Company', H.companyValue);
      trySetBool('SRDED', H.srded);
      trySetBool('CMIF', H.cmif);
      trySetText('Budget Type', H.budgetType);
      trySetText('Urgency justification', H.urgentJustification);

      // Para compatibilidad: tomar el primer proveedor sugerido como principal
      const firstSup = suppliers[0];
      const mainSupName = firstSup?.name;
      const mainSupContact = firstSup?.contact;
      const mainSupEmail = firstSup?.email;

      trySetText('SupplierLegalName', mainSupName);
      trySetText('SupplierContact', mainSupContact);
      trySetText('SupplierEmail', mainSupEmail);

      // D. Justification
      trySetText('Need / Objective', H.needObjective);
      trySetText('Impact if not purchased', H.impactIfNot);
      trySetBool('Sole-source', H.soleSource === 'Yes');
      trySetText('Sole-source explanation', H.soleSourceExplanation);
      trySetBool('Quote attached', !!H.attachQuote);
      trySetBool('SOW/Specification attached', !!H.attachSOW);

      // Aprobadores
      await trySetUserByRole(dex, 'supervisor', supId, body);
      await trySetUserByRole(dex, 'staffManager', staffManagerId, body);
      await trySetUserByRole(dex, 'manager', managerId, body);
      await trySetUserByRole(dex, 'director', directorId, body);
      await trySetUserByRole(dex, 'vp', vpId, body);
      await trySetUserByRole(dex, 'cfo', cfoId, body);
      await trySetUserByRole(dex, 'ceo', ceoId, body);
      await trySetUserByRole(dex, 'procurement', procurementId, body);
      await trySetUserByRole(dex, 'finance', financeId, body);

      // Totals
      const sub = lines.reduce((s, l) => s + safeNum(l.qty) * safeNum(l.unitPrice), 0);
      const tax = 0;
      const grand = sub;
      trySetNumber('Subtotal', Number(sub.toFixed(2)));
      trySetNumber('Tax Total', Number(tax.toFixed(2)));
      trySetNumber('Grand Total', Number(grand.toFixed(2)));

      // Generate PR Number usando la nueva función
      let prNumber = '';
      if (H.companyValue) {
        // H.companyValue es el CompanyCodeforGLAccounts (L01, L02, etc.)
        prNumber = await getNextPRNumberForCompany(H.companyValue);
      } else {
        prNumber = `R${new Date().getTime()}`;
      }

      body['Title'] = prNumber;

      // Create parent
      const parentList =
        parentRef.by === 'id'
          ? sp.web.lists.getById(parentRef.id)
          : sp.web.lists.getByTitle(parentListTitle!);
      const addResp: any = await retryUntilOk(
        async () => await parentList.items.add(body),
        'create parent'
      );
      const parentId =
        addResp?.data?.Id ?? addResp?.data?.ID ?? addResp?.Id ?? addResp?.ID;
      if (!parentId) throw new Error('Could not get parent item ID.');

      // Attachments
      for (const f of files) {
        const buf = await f.arrayBuffer();
        await uploadAttachmentOnce(
          parentListEscName,
          parentId,
          sanitizeName(f.name),
          buf
        );
      }

      // Create child lines
      const childList =
        childRef.by === 'id'
          ? sp.web.lists.getById(childRef.id)
          : sp.web.lists.getByTitle(childListTitle!);
      const lookupInternal =
        childLookupInternal || (await detectChildLookupInternal());

      if (lookupInternal) {
        const childDex = await getChildFieldIndex();

        for (const ln of lines) {
          const bodyLine: Record<string, any> = {};
          bodyLine[`${lookupInternal}Id`] = parentId;

          const putText = (title: string, v?: string, override?: string) => {
            const iname = childNameOfField(childDex, title, override);
            if (iname && v !== undefined) bodyLine[iname] = v;
          };
          const putNum = (title: string, v?: number, override?: string) => {
            const iname = childNameOfField(childDex, title, override);
            if (iname && v !== undefined) bodyLine[iname] = v;
          };

          putText('Item Description / Specification', ln.description, CHILD_INTERNALS.description);
          putText('SKU/Part', ln.sku, CHILD_INTERNALS.sku);
          putNum('Qty', safeNum(ln.qty), CHILD_INTERNALS.qty);
          putText('UoM', ln.uom, CHILD_INTERNALS.uom);
          putNum('Unit Price', safeNum(ln.unitPrice), CHILD_INTERNALS.unitPrice);
          putText('Currency', ln.currency || 'USD', CHILD_INTERNALS.currency);
          putNum('Tax', 0, CHILD_INTERNALS.tax);

          const lineTotal = Number(
            (safeNum(ln.qty) * safeNum(ln.unitPrice)).toFixed(2)
          );
          putNum('Total', lineTotal, CHILD_INTERNALS.total);

          await childList.items.add(bodyLine);
        }
      }

      // ✅ Create suggested suppliers
      if (suppliersRef && suppliers.length) {
        console.log('\n🚀 CREATING SUPPLIERS...');
        console.log('Number of suppliers to create:', suppliers.length);

        const supList =
          suppliersRef.by === 'id'
            ? sp.web.lists.getById(suppliersRef.id)
            : sp.web.lists.getByTitle(suppliersListTitle!);

        const supDex = await getSuppliersFieldIndex();

        const IN_PRID = nameOfFirstExisting(supDex, ['PRId']);
        const IN_NAME = nameOfFirstExisting(supDex, ['SupplierLegalName']);
        const IN_CONTACT = nameOfFirstExisting(supDex, ['SupplierContact']);
        const IN_EMAIL = nameOfFirstExisting(supDex, ['SupplierEmail']);

        console.log('Field mapping:', { IN_PRID, IN_NAME, IN_CONTACT, IN_EMAIL });

        if (!IN_NAME) {
          console.warn('⚠️ Supplier name field not found');
        } else {
          for (const supLine of suppliers) {
            if (!supLine.name?.trim() && !supLine.contact?.trim() && !supLine.email?.trim()) {
              console.log('Skipping empty supplier');
              continue;
            }

            const bodySup: Record<string, any> = {};
            
            // Guardar el ID del PR como texto
            if (IN_PRID) bodySup[IN_PRID] = String(parentId);
            
            if (supLine.name?.trim() && IN_NAME) bodySup[IN_NAME] = supLine.name;
            if (supLine.contact?.trim() && IN_CONTACT) bodySup[IN_CONTACT] = supLine.contact;
            if (supLine.email?.trim() && IN_EMAIL) bodySup[IN_EMAIL] = supLine.email;

            console.log('Creating supplier with body:', bodySup);

            if (Object.keys(bodySup).length > 0) {
              try {
                await supList.items.add(bodySup);
                console.log('✅ Supplier created successfully');
              } catch (err: any) {
                console.error('❌ Error creating supplier:', err);
                throw err;
              }
            }
          }
        }
      }

      showSnack(`PR saved successfully. ID: ${parentId}`, 'success');
      resetFormState();
    } catch (err: any) {
      console.error('Error creating PR:', err);
      if (/406/.test(String(err?.message ?? ''))) {
        showSnack('Request processed. (Note: a non-critical 406 was ignored)', 'info');
      } else {
        showSnack(`Error while saving: ${err.message}`, 'error');
      }
    } finally {
      setSubmitLoading(false);
    }
  };

  /* =================== Update (cuando NO hay aprobaciones) =================== */
  const replaceChildLines = async (parentId: number) => {
    if (!childRef) return;
    const lookupInternal = childLookupInternal || await detectChildLookupInternal();
    if (!lookupInternal) return;

    const childList = childRef.by === 'id' ? sp.web.lists.getById(childRef.id) : sp.web.lists.getByTitle(childListTitle!);

    const qDel =
      `${siteUrl}/_api/web/${listNameOrIdExpr(childRef, childListEscName)}/items` +
      `?$select=Id&$filter=${lookupInternal}/Id eq ${parentId}`;
    const jsDel = await spGet(qDel);
    const toDelete = (jsDel.value || jsDel || []) as Array<{ Id: number }>;
    for (const row of toDelete) {
      try { await childList.items.getById(row.Id).delete(); } catch { /* ignore */ }
    }

    const fieldsResp = await spGet(`${siteUrl}/_api/web/${listNameOrIdExpr(childRef, childListEscName)}/fields?$select=InternalName,Title,TypeAsString`);
    const cfs = (fieldsResp.value || fieldsResp) as FieldInfo[];
    const byTitle = (t: string) => cfs.find(f => norm(f.Title) === norm(t))?.InternalName;

    for (const ln of lines) {
      const bodyLine: Record<string, any> = {};
      bodyLine[`${lookupInternal}Id`] = parentId;
      const setText = (t: string, v?: string) => { const i = byTitle(t); if (i && v !== undefined) bodyLine[i] = v; };
      const setNum = (t: string, v?: number) => { const i = byTitle(t); if (i && v !== undefined) bodyLine[i] = v; };
      setText('Item Description / Specification', ln.description);
      setText('SKU/Part', ln.sku);
      setNum('Qty', safeNum(ln.qty));
      setText('UoM', ln.uom);
      setNum('Unit Price', safeNum(ln.unitPrice));
      setText('Currency', ln.currency || 'USD');
      setNum('Tax', 0);
      setNum('Total', Number((safeNum(ln.qty) * safeNum(ln.unitPrice)).toFixed(2)));
      await childList.items.add(bodyLine);
    }
  };

  const replaceSuppliersLines = async (parentId: number) => {
    if (!suppliersRef) return;

    console.log('\n🔄 REPLACING SUPPLIERS...');
    console.log('Parent ID:', parentId);

    const supList =
      suppliersRef.by === 'id'
        ? sp.web.lists.getById(suppliersRef.id)
        : sp.web.lists.getByTitle(suppliersListTitle!);

    const supDex = await getSuppliersFieldIndex();

    const IN_PRID = nameOfFirstExisting(supDex, ['PRId']);
    const IN_NAME = nameOfFirstExisting(supDex, ['SupplierLegalName']);
    const IN_CONTACT = nameOfFirstExisting(supDex, ['SupplierContact']);
    const IN_EMAIL = nameOfFirstExisting(supDex, ['SupplierEmail']);

    if (!IN_NAME || !IN_PRID) {
      throw new Error('Supplier name or PRId field not found');
    }

    console.log('Field mapping:', { IN_PRID, IN_NAME, IN_CONTACT, IN_EMAIL });

    // Eliminar proveedores antiguos (por PRId)
    const qDel =
      `${siteUrl}/_api/web/${listNameOrIdExpr(suppliersRef, suppliersListEscName)}/items` +
      `?$select=Id&$filter=${IN_PRID} eq '${parentId}'`;
    
    console.log('Delete query:', qDel);

    const jsDel = await spGet(qDel);
    const toDelete = (jsDel.value || jsDel || []) as Array<{ Id: number }>;
    
    console.log('Suppliers to delete:', toDelete.length);

    for (const row of toDelete) {
      try {
        await supList.items.getById(row.Id).delete();
        console.log(`✅ Deleted supplier ${row.Id}`);
      } catch (err) {
        console.error(`❌ Error deleting supplier ${row.Id}:`, err);
      }
    }

    // Crear nuevos proveedores
    console.log('Creating new suppliers:', suppliers.length);

    for (const supLine of suppliers) {
      if (!supLine.name?.trim() && !supLine.contact?.trim() && !supLine.email?.trim()) {
        continue;
      }

      const bodySup: Record<string, any> = {};
      bodySup[IN_PRID] = String(parentId);

      if (supLine.name?.trim() && IN_NAME) bodySup[IN_NAME] = supLine.name;
      if (supLine.contact?.trim() && IN_CONTACT) bodySup[IN_CONTACT] = supLine.contact;
      if (supLine.email?.trim() && IN_EMAIL) bodySup[IN_EMAIL] = supLine.email;

      if (Object.keys(bodySup).length > 1) {
        try {
          await supList.items.add(bodySup);
          console.log('✅ Supplier created:', bodySup);
        } catch (err: any) {
          console.error('❌ Error creating supplier:', err);
        }
      }
    }
  };

  const updateParentAndLines = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!parentRef || !childRef || !editingItemId) {
      showSnack('No PR loaded to update.', 'error');
      return;
    }
    if (hasAnyApproval) {
      showSnack('This PR is locked (approval present). Editing is disabled.', 'info');
      return;
    }
    try {
      setSubmitLoading(true);
      const dex = await getParentFieldIndex();
      const nameOfField = buildNameOfField(dex);

      const H = headerDraft;

      const requestIdUser =
        H.requester?.secondaryText || H.requester?.loginName || H.requester?.text || null;
      const requesterId = await ensureUserIdSmart(requestIdUser) ?? currentUserIdRef.current;

      const supId = await ensureUserIdSmart(H.supervisor?.secondaryText || H.supervisor?.loginName || H.supervisor?.text || null);
      const staffManagerId = await ensureUserIdSmart(H.staffManager?.secondaryText || H.staffManager?.loginName || H.staffManager?.text || null);
      const managerId = await ensureUserIdSmart(H.manager?.secondaryText || H.manager?.loginName || H.manager?.text || null);
      const directorId = await ensureUserIdSmart(H.director?.secondaryText || H.director?.loginName || H.director?.text || null);
      const vpId = await ensureUserIdSmart(H.vp?.secondaryText || H.vp?.loginName || H.vp?.text || null);
      const cfoId = await ensureUserIdSmart(H.cfo?.secondaryText || H.cfo?.loginName || H.cfo?.text || null);
      const ceoId = await ensureUserIdSmart(H.ceo?.secondaryText || H.ceo?.loginName || H.ceo?.text || null);
      const procurementId = await ensureUserIdSmart(H.procurement?.secondaryText || H.procurement?.loginName || H.procurement?.text || null);
      const financeId = await ensureUserIdSmart(H.finance?.secondaryText || H.finance?.loginName || H.finance?.text || null);

      const canWrite = (iname: string | null): boolean => {
        if (!iname) return false;
        const info = dex.byInternal.get(iname);
        if (!info) return true;
        if (info.Hidden) return false;
        if (info.ReadOnlyField) return false;
        return true;
      };

      const body: Record<string, any> = {};

      const trySetText = (titleLabel: string, v?: string) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname) && v !== undefined) body[iname] = v;
      };
      const trySetDate = (titleLabel: string, ymd?: string) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname)) {
          const iso = toSharePointDateLocal(ymd);
          if (iso) body[iname] = iso;
        }
      };
      const trySetBool = (titleLabel: string, v?: boolean) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname) && v !== undefined) body[iname] = v;
      };
      const trySetNumber = (titleLabel: string, n?: number) => {
        const iname = nameOfField(titleLabel);
        if (iname && canWrite(iname) && typeof n === 'number' && !Number.isNaN(n)) body[iname] = n;
      };
      const trySetRoleDate = (titleCandidates: string[], ymd?: string) => {
        const iname = nameOfFirstExisting(dex, titleCandidates);
        if (!iname || !canWrite(iname)) return;
        const iso = toSharePointDateLocal(ymd);
        if (iso) body[iname] = iso;
      };

      await trySetUserByRole(dex, 'requester', requesterId ?? null, body);
      await trySetUserByRole(dex, 'supervisor', supId, body);
      await trySetUserByRole(dex, 'staffManager', staffManagerId, body);
      await trySetUserByRole(dex, 'manager', managerId, body);
      await trySetUserByRole(dex, 'director', directorId, body);
      await trySetUserByRole(dex, 'vp', vpId, body);
      await trySetUserByRole(dex, 'cfo', cfoId, body);
      await trySetUserByRole(dex, 'ceo', ceoId, body);
      await trySetUserByRole(dex, 'procurement', procurementId, body);
      await trySetUserByRole(dex, 'finance', financeId, body);

      // Fallback Requester Email
      if ((!H.requesterEmail || H.requesterEmail.trim() === '') && requesterId) {
        try {
          const u = await sp.web.getUserById(requesterId)();
          if (u?.Email) {
            const emailInternal = nameOfField('Requester Email');
            if (emailInternal) body[emailInternal] = u.Email;
            setHeaderDraft(prev => ({ ...prev, requesterEmail: u.Email }));
          }
        } catch { /* ignore */ }
      }

      // Header fields
      trySetText('Requester Email', H.requesterEmail);
      trySetDate('Request Date', H.requestDate);
      trySetDate('Requiredbydate', H.requiredByDate);
      trySetText('Area', H.area);
      trySetText('Area / Project', H.areaProject);
      trySetText('Priority', H.priority);
      trySetText('Type', H.reqType);
      trySetText('GL Code', H.glCode);
      trySetText('Company', H.companyValue);
      trySetBool('SRDED', H.srded);
      trySetBool('CMIF', H.cmif);
      trySetText('Budget Type', H.budgetType);
      trySetText('Urgency justification', H.urgentJustification);

      const firstSup = suppliers[0];
      const mainSupName = firstSup?.name;
      const mainSupContact = firstSup?.contact;
      const mainSupEmail = firstSup?.email;

      trySetText('SupplierLegalName', mainSupName);
      trySetText('SupplierContact', mainSupContact);
      trySetText('SupplierEmail', mainSupEmail);

      trySetText('Need / Objective', H.needObjective);
      trySetText('Impact if not purchased', H.impactIfNot);
      trySetBool('Sole-source', H.soleSource === 'Yes');
      trySetText('Sole-source explanation', H.soleSourceExplanation);
      trySetBool('Quote attached', !!H.attachQuote);
      trySetBool('SOW/Specification attached', !!H.attachSOW);

      // Role dates
      trySetRoleDate(['Supervisor Date'], H.supervisorDate);
      trySetRoleDate(['Staff Manager Date'], H.staffManagerDate);
      trySetRoleDate(['Manager Date'], H.managerDate);
      trySetRoleDate(['Director Date'], H.directorDate);
      trySetRoleDate(['VP Date'], H.vpDate);
      trySetRoleDate(['CFO Date'], H.cfoDate);
      trySetRoleDate(['CEO Date'], H.ceoDate);
      trySetRoleDate(['Procurement Date'], H.procurementDate);
      trySetRoleDate(['Finance Date'], H.financeDate);
      trySetText('PO Number (assigned by Procurement)', H.poNumber);

      // Totals
      const sub = lines.reduce((s, l) => s + (safeNum(l.qty) * safeNum(l.unitPrice)), 0);
      const tax = 0;
      const grand = sub;
      trySetNumber('Subtotal', Number(sub.toFixed(2)));
      trySetNumber('Tax Total', Number(tax.toFixed(2)));
      trySetNumber('Grand Total', Number(grand.toFixed(2)));

      const parentList = parentRef.by === 'id' ? sp.web.lists.getById(parentRef.id) : sp.web.lists.getByTitle(parentListTitle!);
      await parentList.items.getById(editingItemId).update(body);

      await replaceChildLines(editingItemId);
      await replaceSuppliersLines(editingItemId);

      showSnack(`PR #${editingItemId} updated successfully.`, 'success');
      await loadMySent();

    } catch (err: any) {
      console.error('Error updating PR:', err);
      if (/406/.test(String(err?.message ?? ''))) {
        showSnack('Update processed. (A non-critical 406 was ignored)', 'info');
      } else {
        showSnack(`Error while updating: ${err.message}`, 'error');
      }
    } finally {
      setSubmitLoading(false);
    }
  };

  const resetFormState = () => {
    const fresh = { ...initialHeader, requestDate: getTodayYmd() };
    setHeaderDraft(fresh);
    setLines([{ id: 1, qty: 1, unitPrice: 0, tax: 0, currency: 'USD' }]);
    setSuppliers([{ id: 1, name: '', contact: '', email: '' }]);
    setFiles([]); if (attachInputRef.current) attachInputRef.current.value = '';
    setEditingItemId(null);
    setHasAnyApproval(false);
    setSignFormRoles(null);
    setGlNoPagination(false);
  };

  // SWITCH VIEW
  const switchView = async (v: typeof activeView) => {
    if (v === 'new') {
      resetFormState();
    } else {
      setEditingItemId(null);
      setHasAnyApproval(false);
      setSignFormRoles(null);
      setGlNoPagination(false);
    }
    setActiveView(v);
    if (v === 'mysent') await loadMySent();
    if (v === 'tosign') await loadMyToSign();
    if (v === 'approved') await loadMyApproved();
  };

  /* =================== Load item into form =================== */
  const loadItemIntoForm = async (itemId: number, opts?: { signRoles?: RoleKey[] }) => {
    if (!parentRef) return;
    try {
      const dex = await getParentFieldIndex();
      const nameOfField = buildNameOfField(dex);

      // Detect approvals para bloqueo
      const supInt = await getPersonInternalForRole('supervisor');
      const staffManagerInt = await getPersonInternalForRole('staffManager');
      const managerInt = await getPersonInternalForRole('manager');
      const directorInt = await getPersonInternalForRole('director');
      const vpInt = await getPersonInternalForRole('vp');
      const cfoInt = await getPersonInternalForRole('cfo');
      const ceoInt = await getPersonInternalForRole('ceo');
      const procurementInt = await getPersonInternalForRole('procurement');
      const financeInt = await getPersonInternalForRole('finance');

      let hasApproval = false;

      if (supInt || staffManagerInt || managerInt || directorInt || vpInt || cfoInt || ceoInt || procurementInt || financeInt) {
        try {
          const itApproval = await spGet(`${siteUrl}/_api/web/${listNameOrIdExpr(parentRef, parentListEscName)}/items(${itemId})?$select=${[supInt, staffManagerInt, managerInt, directorInt, vpInt, cfoInt, ceoInt, procurementInt, financeInt].filter(Boolean).join(',')}`);
          hasApproval =
            !!(supInt && (itApproval as any)[`${supInt}Id`]) ||
            !!(staffManagerInt && (itApproval as any)[`${staffManagerInt}Id`]) ||
            !!(managerInt && (itApproval as any)[`${managerInt}Id`]) ||
            !!(directorInt && (itApproval as any)[`${directorInt}Id`]) ||
            !!(vpInt && (itApproval as any)[`${vpInt}Id`]) ||
            !!(cfoInt && (itApproval as any)[`${cfoInt}Id`]) ||
            !!(ceoInt && (itApproval as any)[`${ceoInt}Id`]) ||
            !!(procurementInt && (itApproval as any)[`${procurementInt}Id`]) ||
            !!(financeInt && (itApproval as any)[`${financeInt}Id`]);
        } catch {
          hasApproval = false;
        }
      }
      setHasAnyApproval(!!hasApproval);

      // Fetch item
      const item = await spGet(`${siteUrl}/_api/web/${listNameOrIdExpr(parentRef, parentListEscName)}/items(${itemId})`);

      const getTxt = (t: string) => {
        const i = nameOfField(t); return i ? (item[i] ?? '') : '';
      };
      const getBool = (t: string) => {
        const i = nameOfField(t); return i ? !!item[i] : false;
      };
      const getDate = (t: string) => {
        const i = nameOfField(t); return isoToYmd(i ? item[i] : undefined);
      };

      const loadPickerByRole = async (role: RoleKey) => {
        const i = await getPersonInternalForRole(role);
        if (!i) return null;
        const idVal = item[`${i}Id`];
        const firstId = Array.isArray(idVal?.results) ? idVal.results[0] : idVal;
        const id = typeof firstId === 'number' ? firstId : null;
        if (!id) return null;
        try {
          const u = await sp.web.getUserById(id)();
          return { id: u.Id, loginName: u.LoginName, text: u.Title, secondaryText: u.Email } as any;
        } catch { return null; }
      };

      const loaded: PRHeader = {
        requester: await loadPickerByRole('requester'),
        requesterEmail: getTxt('Requester Email'),
        requestDate: getDate('Request Date'),
        requiredByDate: getDate('Requiredbydate'),
        area: getTxt('Area'),
        areaProject: getTxt('Area / Project'),
        glCode: getTxt('GL Code'),
        companyValue: getTxt('Company'),
        srded: getBool('SRDED'),
        cmif: getBool('CMIF'),
        budgetType: getTxt('Budget Type') as BudgetType,
        priority: getTxt('Priority') as Priority,
        reqType: getTxt('Type') as RequestType,
        urgentJustification: getTxt('Urgency justification'),

        needObjective: getTxt('Need / Objective'),
        impactIfNot: getTxt('Impact if not purchased'),
        soleSource: getBool('Sole-source') ? 'Yes' : 'No',
        soleSourceExplanation: getTxt('Sole-source explanation'),
        attachQuote: getBool('Quote attached'),
        attachSOW: getBool('SOW/Specification attached'),

        supervisor: await loadPickerByRole('supervisor'),
        staffManager: await loadPickerByRole('staffManager'),
        manager: await loadPickerByRole('manager'),
        director: await loadPickerByRole('director'),
        vp: await loadPickerByRole('vp'),
        cfo: await loadPickerByRole('cfo'),
        ceo: await loadPickerByRole('ceo'),
        procurement: await loadPickerByRole('procurement'),
        finance: await loadPickerByRole('finance'),

        supervisorDate: getDate('Supervisor Date'),
        staffManagerDate: getDate('Staff Manager Date'),
        managerDate: getDate('Manager Date'),
        directorDate: getDate('Director Date'),
        vpDate: getDate('VP Date'),
        cfoDate: getDate('CFO Date'),
        ceoDate: getDate('CEO Date'),
        procurementDate: getDate('Procurement Date'),
        financeDate: getDate('Finance Date'),

        supervisorStatus: (getTxt('Supervisor status') || 'Pending') as ApprovalStatus,
        staffManagerStatus: (getTxt('Staff Manager status') || 'Pending') as ApprovalStatus,
        managerStatus: (getTxt('Manager status') || 'Pending') as ApprovalStatus,
        directorStatus: (getTxt('Director status') || 'Pending') as ApprovalStatus,
        vpStatus: (getTxt('VP status') || 'Pending') as ApprovalStatus,
        cfoStatus: (getTxt('CFO status') || 'Pending') as ApprovalStatus,
        ceoStatus: (getTxt('CEO status') || 'Pending') as ApprovalStatus,
        procurementStatus: (getTxt('Procurement status') || 'Pending') as ApprovalStatus,
        financeStatus: (getTxt('Finance status') || 'Pending') as ApprovalStatus,

        poNumber: getTxt('PO Number (assigned by Procurement)')
      };

      setHeaderDraft(loaded);

      setEditingItemId(itemId);
      if (opts?.signRoles) setSignFormRoles(opts.signRoles!); else setSignFormRoles(null);

      // Load child lines
      if (childRef) {
        const lookupInt = childLookupInternal || await detectChildLookupInternal();
        if (lookupInt) {
          const cDex = await getChildFieldIndex();

          const IN_DESC = childNameOfField(
            cDex,
            'Item Description / Specification',
            'ItemDescription_x002f_Specificat'
          );
          const IN_SKU   = childNameOfField(cDex, 'SKU/Part', 'SKU_x002f_Par');
          const IN_QTY   = childNameOfField(cDex, 'Qty');
          const IN_UOM   = childNameOfField(cDex, 'UoM');
          const IN_UNIT  = childNameOfField(cDex, 'Unit Price', 'UnitPrice');
          const IN_CURR  = childNameOfField(cDex, 'Currency');
          const IN_TAX   = childNameOfField(cDex, 'Tax');
          const IN_TOTAL = childNameOfField(cDex, 'Total');

          const selectCols = [
            'Id', 'Title',
            IN_DESC, IN_SKU, IN_QTY, IN_UOM, IN_UNIT, IN_CURR, IN_TAX, IN_TOTAL
          ].filter(Boolean).join(',');

          const url =
            `${siteUrl}/_api/web/${listNameOrIdExpr(childRef, childListEscName)}/items` +
            `?$select=${selectCols}&$filter=${lookupInt}/Id eq ${itemId}&$orderby=Id asc`;

          const js = await spGet(url);
          const rows = (js.value || js || []) as any[];

          const num = (v: any) => {
            const n = Number(v);
            return Number.isFinite(n) ? n : 0;
          };

          const parsed: PRLine[] = rows.map((r, idx) => ({
            id: idx + 1,
            description: IN_DESC ? (r[IN_DESC] ?? '') : '',
            sku:         IN_SKU  ? (r[IN_SKU]  ?? '') : '',
            qty:         IN_QTY  ? num(r[IN_QTY])     : 0,
            uom:         IN_UOM  ? (r[IN_UOM]  ?? '') : '',
            unitPrice:   IN_UNIT ? num(r[IN_UNIT])    : 0,
            currency:    IN_CURR ? (r[IN_CURR] ?? 'USD') : 'USD',
            tax:         IN_TAX  ? num(r[IN_TAX])     : 0,
            total:       IN_TOTAL? num(r[IN_TOTAL])   : 0,
          }));

          setLines(parsed.length ? parsed : [{ id: 1, qty: 1, unitPrice: 0, tax: 0, currency: 'USD' }]);
        }
      }

      // Load suggested suppliers
      if (suppliersRef) {
        const supDex = await getSuppliersFieldIndex();

        const IN_PRID = nameOfFirstExisting(supDex, ['PRId']);
        const IN_NAME = nameOfFirstExisting(supDex, ['SupplierLegalName']);
        const IN_CONTACT = nameOfFirstExisting(supDex, ['SupplierContact']);
        const IN_EMAIL = nameOfFirstExisting(supDex, ['SupplierEmail']);

        console.log('Loading suppliers with fields:', { IN_PRID, IN_NAME, IN_CONTACT, IN_EMAIL });

        if (!IN_NAME || !IN_PRID) {
          console.warn('⚠️ Supplier name or PRId field not found');
          setSuppliers([{ id: 1, name: '', contact: '', email: '' }]);
        } else {
          const selectCols = ['Id', IN_PRID, IN_NAME];
          if (IN_CONTACT) selectCols.push(IN_CONTACT);
          if (IN_EMAIL) selectCols.push(IN_EMAIL);

          const url =
            `${siteUrl}/_api/web/${listNameOrIdExpr(suppliersRef, suppliersListEscName)}/items` +
            `?$select=${selectCols.join(',')}&$filter=${IN_PRID} eq '${itemId}'&$orderby=Id asc`;

          console.log('Loading suppliers URL:', url);

          const jsSup = await spGet(url);
          const rowsSup = (jsSup.value || jsSup || []) as any[];

          console.log('Suppliers loaded:', rowsSup.length);

          const supLines: SupplierLine[] = rowsSup.map((r, idx) => ({
            id: idx + 1,
            name: IN_NAME ? (r[IN_NAME] ?? '') : '',
            contact: IN_CONTACT ? (r[IN_CONTACT] ?? '') : '',
            email: IN_EMAIL ? (r[IN_EMAIL] ?? '') : ''
          }));

          setSuppliers(supLines.length ? supLines : [{ id: 1, name: '', contact: '', email: '' }]);
        }
      } else {
        setSuppliers([{ id: 1, name: '', contact: '', email: '' }]);
      }
    } catch (err: any) {
      showSnack(`Could not load item: ${err.message}`, 'error');
    }
  };

  // Helper: set role date by candidates
  const setRoleDateByCandidates = async (itemId: number, role: RoleKey, whenIso: string) => {
    const titles = ROLE_DATE_TITLES[role] || [];
    if (!titles.length) return;
    const dex = await getParentFieldIndex();
    const iname = nameOfFirstExisting(dex, titles);
    if (!iname) return;
    const parentList = parentRef!.by === 'id' ? sp.web.lists.getById(parentRef!.id) : sp.web.lists.getByTitle(parentListTitle!);
    await parentList.items.getById(itemId).update({ [iname]: whenIso });
  };

  const setRoleStatusByCandidates = async (itemId: number, role: RoleKey, status: ApprovalStatus) => {
    const titles = ROLE_STATUS_TITLES[role] || [];
    if (!titles.length) return;
    const dex = await getParentFieldIndex();
    const iname = nameOfFirstExisting(dex, titles);
    if (!iname) return;
    const parentList = parentRef!.by === 'id'
      ? sp.web.lists.getById(parentRef!.id)
      : sp.web.lists.getByTitle(parentListTitle!);
    await parentList.items.getById(itemId).update({ [iname]: status });
  };

  /* =================== Approve / Disagree =================== */
  const applyDecision = async (role: RoleKey, status: ApprovalStatus) => {
    if (!editingItemId || !parentRef) { showSnack('No item loaded.', 'error'); return; }
    if (role === 'requester') { showSnack('Requester does not use status fields.', 'error'); return; }

    try {
      const nowIso = new Date().toISOString();
      await setRoleDateByCandidates(editingItemId, role, nowIso);
      await setRoleStatusByCandidates(editingItemId, role, status);

      const ymd = isoToYmd(nowIso);
      setHeaderDraft(prev => {
        const patch: Partial<PRHeader> = {};
        if (role === 'supervisor') {
          patch.supervisorDate = ymd;
          patch.supervisorStatus = status;
        }
        if (role === 'staffManager') {
          patch.staffManagerDate = ymd;
          patch.staffManagerStatus = status;
        }
        if (role === 'manager') {
          patch.managerDate = ymd;
          patch.managerStatus = status;
        }
        if (role === 'director') {
          patch.directorDate = ymd;
          patch.directorStatus = status;
        }
        if (role === 'vp') {
          patch.vpDate = ymd;
          patch.vpStatus = status;
        }
        if (role === 'cfo') {
          patch.cfoDate = ymd;
          patch.cfoStatus = status;
        }
        if (role === 'ceo') {
          patch.ceoDate = ymd;
          patch.ceoStatus = status;
        }
        if (role === 'procurement') {
          patch.procurementDate = ymd;
          patch.procurementStatus = status;
        }
        if (role === 'finance') {
          patch.financeDate = ymd;
          patch.financeStatus = status;
        }
        return { ...prev, ...patch };
      });

      setHasAnyApproval(true);
      await loadMyToSign();
      await loadMyApproved();

      const label =
        role === 'supervisor' ? 'Supervisor' :
          role === 'staffManager' ? 'Staff Manager' :
            role === 'manager' ? 'Manager' :
              role === 'director' ? 'Director' :
                role === 'vp' ? 'VP' :
                  role === 'cfo' ? 'CFO' :
                    role === 'ceo' ? 'CEO' :
                      role === 'procurement' ? 'Procurement' :
                        role === 'finance' ? 'Finance' : role;

      const actionText = status === 'Agree' ? 'approved (Agree)' : 'marked as Disagree';
      showSnack(`${label} ${actionText} successfully.`, status === 'Agree' ? 'success' : 'info');
    } catch (err: any) {
      if (/406/.test(String(err?.message ?? ''))) {
        showSnack('Decision processed (non-critical 406 ignored).', 'info');
      } else {
        showSnack(`Error while applying decision: ${err.message}`, 'error');
      }
    }
  };

  const handleApprove = async (role: RoleKey) => {
    await applyDecision(role, 'Agree');
  };

  const handleDisagree = async (role: RoleKey) => {
    await applyDecision(role, 'Disagree');
  };

  /* =================== PDF =================== */
  const handleGeneratePdf = async (item?: any) => {
    try {
      setPdfLoading(true);
      const itemId = item?.Id ?? editingItemId;
      if (!itemId || !parentRef) { showSnack('No item selected.', 'error'); return; }

      // ==== Parent field index para obtener status ====
      const dex = await getParentFieldIndex();
      const nf = buildNameOfField(dex);
      const supStatusInt = nf('Supervisor status');
      const staffManagerStatusInt = nf('Staff Manager status');
      const managerStatusInt = nf('Manager status');
      const directorStatusInt = nf('Director status');
      const vpStatusInt = nf('VP status');
      const cfoStatusInt = nf('CFO status');
      const ceoStatusInt = nf('CEO status');
      const procurementStatusInt = nf('Procurement status');
      const financeStatusInt = nf('Finance status');

      const hasAllStatusColumns = !!(supStatusInt && staffManagerStatusInt && managerStatusInt && directorStatusInt && vpStatusInt && cfoStatusInt && ceoStatusInt && procurementStatusInt && financeStatusInt);
      if (!hasAllStatusColumns) {
        showSnack('Cannot generate PDF: status fields are not configured in the list.', 'error');
        return;
      }

      // ==== Personas (para nombres/emails en "Approvals") ====
      const supP = await getPersonInternalForRole('supervisor');
      const staffManagerP = await getPersonInternalForRole('staffManager');
      const staffManager2P = await getPersonInternalForRole('staffManager2');
      const managerP = await getPersonInternalForRole('manager');
      const manager2P = await getPersonInternalForRole('manager2');
      const directorP = await getPersonInternalForRole('director');
      const vpP = await getPersonInternalForRole('vp');
      const cfoP = await getPersonInternalForRole('cfo');
      const ceoP = await getPersonInternalForRole('ceo');
      const procurementP = await getPersonInternalForRole('procurement');
      const financeP = await getPersonInternalForRole('finance');

      const selectParts: string[] = ['Title'];
      if (supP) selectParts.push(`${supP}/Title`, `${supP}/EMail`);
      if (staffManagerP) selectParts.push(`${staffManagerP}/Title`, `${staffManagerP}/EMail`);
      if (staffManager2P) selectParts.push(`${staffManager2P}/Title`, `${staffManager2P}/EMail`);
      if (managerP) selectParts.push(`${managerP}/Title`, `${managerP}/EMail`);
      if (manager2P) selectParts.push(`${manager2P}/Title`, `${manager2P}/EMail`);
      if (directorP) selectParts.push(`${directorP}/Title`, `${directorP}/EMail`);
      if (vpP) selectParts.push(`${vpP}/Title`, `${vpP}/EMail`);
      if (cfoP) selectParts.push(`${cfoP}/Title`, `${cfoP}/EMail`);
      if (ceoP) selectParts.push(`${ceoP}/Title`, `${ceoP}/EMail`);
      if (procurementP) selectParts.push(`${procurementP}/Title`, `${procurementP}/EMail`);
      if (financeP) selectParts.push(`${financeP}/Title`, `${financeP}/EMail`);
      if (supStatusInt) selectParts.push(supStatusInt);
      if (staffManagerStatusInt) selectParts.push(staffManagerStatusInt);
      if (managerStatusInt) selectParts.push(managerStatusInt);
      if (directorStatusInt) selectParts.push(directorStatusInt);
      if (vpStatusInt) selectParts.push(vpStatusInt);
      if (cfoStatusInt) selectParts.push(cfoStatusInt);
      if (ceoStatusInt) selectParts.push(ceoStatusInt);
      if (procurementStatusInt) selectParts.push(procurementStatusInt);
      if (financeStatusInt) selectParts.push(financeStatusInt);

      const expandParts = [supP, staffManagerP, managerP, directorP, vpP, cfoP, ceoP, procurementP, financeP].filter(Boolean) as string[];

      const parentItemUrl =
        `${siteUrl}/_api/web/${listNameOrIdExpr(parentRef, parentListEscName)}/items(${itemId})` +
        `?$select=${selectParts.join(',')}` +
        (expandParts.length ? `&$expand=${expandParts.join(',')}` : ``);

      const spItem = await spGet(parentItemUrl);

      const supStatusVal: string | undefined = supStatusInt ? spItem?.[supStatusInt] : undefined;
      const staffManagerStatusVal: string | undefined = staffManagerStatusInt ? spItem?.[staffManagerStatusInt] : undefined;
      const managerStatusVal: string | undefined = managerStatusInt ? spItem?.[managerStatusInt] : undefined;
      const directorStatusVal: string | undefined = directorStatusInt ? spItem?.[directorStatusInt] : undefined;
      const vpStatusVal: string | undefined = vpStatusInt ? spItem?.[vpStatusInt] : undefined;
      const cfoStatusVal: string | undefined = cfoStatusInt ? spItem?.[cfoStatusInt] : undefined;
      const ceoStatusVal: string | undefined = ceoStatusInt ? spItem?.[ceoStatusInt] : undefined;
      const procurementStatusVal: string | undefined = procurementStatusInt ? spItem?.[procurementStatusInt] : undefined;
      const financeStatusVal: string | undefined = financeStatusInt ? spItem?.[financeStatusInt] : undefined;

      const normStatus = (v?: string) => (v || '').toLowerCase();
      const allAgree =
        normStatus(supStatusVal) === 'agree' &&
        normStatus(staffManagerStatusVal) === 'agree' &&
        normStatus(managerStatusVal) === 'agree' &&
        normStatus(directorStatusVal) === 'agree' &&
        normStatus(vpStatusVal) === 'agree' &&
        normStatus(cfoStatusVal) === 'agree' &&
        normStatus(ceoStatusVal) === 'agree' &&
        normStatus(procurementStatusVal) === 'agree' &&
        normStatus(financeStatusVal) === 'agree';

      if (!allAgree) {
        showSnack('Cannot generate PDF: All approvers must have status "Agree".', 'error');
        return;
      }

      const approverInfo: Record<RoleKey, { name?: string; email?: string }> = {
        requester: { name: headerDraft.requester?.text, email: headerDraft.requesterEmail || headerDraft.requester?.secondaryText },
        supervisor: { name: supP ? spItem?.[supP]?.Title : undefined, email: supP ? spItem?.[supP]?.EMail : undefined },
        staffManager: { name: staffManagerP ? spItem?.[staffManagerP]?.Title : undefined, email: staffManagerP ? spItem?.[staffManagerP]?.EMail : undefined },
        staffManager2: { name: staffManager2P ? spItem?.[staffManager2P]?.Title : undefined, email: staffManager2P ? spItem?.[staffManager2P]?.EMail : undefined },
        manager: { name: managerP ? spItem?.[managerP]?.Title : undefined, email: managerP ? spItem?.[managerP]?.EMail : undefined },
        manager2: { name: manager2P ? spItem?.[manager2P]?.Title : undefined, email: manager2P ? spItem?.[manager2P]?.EMail : undefined },
        director: { name: directorP ? spItem?.[directorP]?.Title : undefined, email: directorP ? spItem?.[directorP]?.EMail : undefined },
        vp: { name: vpP ? spItem?.[vpP]?.Title : undefined, email: vpP ? spItem?.[vpP]?.EMail : undefined },
        cfo: { name: cfoP ? spItem?.[cfoP]?.Title : undefined, email: cfoP ? spItem?.[cfoP]?.EMail : undefined },
        ceo: { name: ceoP ? spItem?.[ceoP]?.Title : undefined, email: ceoP ? spItem?.[ceoP]?.EMail : undefined },
        procurement: { name: procurementP ? spItem?.[procurementP]?.Title : undefined, email: procurementP ? spItem?.[procurementP]?.EMail : undefined },
        finance: { name: financeP ? spItem?.[financeP]?.Title : undefined, email: financeP ? spItem?.[financeP]?.EMail : undefined }
      };

      // ==== Asegurar líneas (por si estado vacío) ====
      let pdfLines = lines;
      if (!pdfLines || pdfLines.length === 0 || !pdfLines[0]?.description) {
        const lookupInt = childLookupInternal || await detectChildLookupInternal();
        if (childRef && lookupInt) {
          const url =
            `${siteUrl}/_api/web/${listNameOrIdExpr(childRef, childListEscName)}/items` +
            `?$select=Id,Title,*&$filter=${lookupInt}/Id eq ${itemId}&$orderby=Id asc`;
          const js = await spGet(url);
          const rows = (js.value || js || []) as any[];
          const pick = (o: any, keys: string[]) => {
            for (const k of Object.keys(o)) {
              const nk = norm(k);
              if (keys.some(s => nk === norm(s))) return o[k];
            }
            return '';
          };
          pdfLines = rows.map((r, idx) => ({
            id: idx + 1,
            description: pick(r, ['Item Description / Specification']),
            sku: pick(r, ['SKU/Part']),
            qty: Number(pick(r, ['Qty'])) || 0,
            uom: pick(r, ['UoM']),
            unitPrice: Number(pick(r, ['Unit Price'])) || 0,
            currency: pick(r, ['Currency']) || 'USD',
            tax: Number(pick(r, ['Tax'])) || 0,
            total: Number(pick(r, ['Total'])) || 0
          }));
        }
      }

      // ==== Asegurar proveedores sugeridos (por si estado vacío) ====
      let pdfSuppliers = suppliers;
      if (!pdfSuppliers || !pdfSuppliers.length || (!pdfSuppliers[0].name && !pdfSuppliers[0].contact && !pdfSuppliers[0].email)) {
        if (suppliersRef) {
          const lookupSup = supLookupInternal || await detectSuppliersLookupInternal();
          if (lookupSup) {
            const supDex = await getSuppliersFieldIndex();

            const IN_NAME = nameOfFirstExisting(supDex, SUPPLIER_INTERNALS.name ? [SUPPLIER_INTERNALS.name] : []);
            const IN_CONTACT = nameOfFirstExisting(supDex, SUPPLIER_INTERNALS.contact ? [SUPPLIER_INTERNALS.contact] : []);
            const IN_EMAIL = nameOfFirstExisting(supDex, SUPPLIER_INTERNALS.email ? [SUPPLIER_INTERNALS.email] : []);

            if (IN_NAME) {
              const selectCols = ['Id', IN_NAME];
              if (IN_CONTACT) selectCols.push(IN_CONTACT);
              if (IN_EMAIL) selectCols.push(IN_EMAIL);

              const url =
                `${siteUrl}/_api/web/${listNameOrIdExpr(suppliersRef, suppliersListEscName)}/items` +
                `?$select=${selectCols.join(',')}&$filter=${lookupSup}/Id eq ${itemId}&$orderby=Id asc`;

              const jsSup = await spGet(url);
              const rowsSup = (jsSup.value || jsSup || []) as any[];

              pdfSuppliers = rowsSup.map((r, idx) => ({
                id: idx + 1,
                name: IN_NAME ? (r[IN_NAME] ?? '') : '',
                contact: IN_CONTACT ? (r[IN_CONTACT] ?? '') : '',
                email: IN_EMAIL ? (r[IN_EMAIL] ?? '') : ''
              }));
            }
          }
        }
      }

      // ==== jsPDF setup ====
      const doc = new jsPDF('p', 'pt', 'letter');
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const marginLeft = 40;
      const marginTop = 40;
      const bottomMargin = 40;

      const RIGHT_MARGIN = 40;
      const LABEL_W = 160;
      let currentY = marginTop + 60;

      const ensureSpace = (needed = 40) => {
        if (currentY + needed > pageHeight - bottomMargin) {
          doc.addPage();
          currentY = marginTop;
        }
      };

      const sectionTitle = (txt: string) => {
        ensureSpace(28);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(14);
        doc.text(txt, marginLeft, currentY);
        currentY += 12;
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(12);
      };

      const writeWrapSmart = (label: string, value?: string) => {
        const safe = (value ?? '').toString().trim() || '-';
        const labelTxt = `${label}:`;
        const labelWidth = doc.getTextWidth(labelTxt);
        const MAX_LABEL_W = LABEL_W;
        const gap = 8;

        if (labelWidth <= MAX_LABEL_W) {
          ensureSpace(22);
          doc.setFont('helvetica', 'bold');
          doc.text(labelTxt, marginLeft, currentY);

          doc.setFont('helvetica', 'normal');
          const valX = marginLeft + MAX_LABEL_W + gap;
          const valWidth = pageWidth - valX - RIGHT_MARGIN;
          const linesTxt = doc.splitTextToSize(safe, valWidth);
          doc.text(linesTxt, valX, currentY);
          currentY += Math.max(18, linesTxt.length * 14);
        } else {
          ensureSpace(22);
          doc.setFont('helvetica', 'bold');
          const labelLines = doc.splitTextToSize(labelTxt, pageWidth - marginLeft - RIGHT_MARGIN);
          doc.text(labelLines, marginLeft, currentY);
          currentY += Math.max(16, labelLines.length * 14);

          doc.setFont('helvetica', 'normal');
          const valWidth = pageWidth - marginLeft - RIGHT_MARGIN;
          const linesTxt = doc.splitTextToSize(safe, valWidth);
          doc.text(linesTxt, marginLeft, currentY);
          currentY += Math.max(18, linesTxt.length * 14);
        }
      };

      // LOGO
      try { doc.addImage(flLogo as any, 'JPG', pageWidth - 140, marginTop, 100, 45); } catch { }

      // Título
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(16);
      doc.text('Purchase Requisition (PR)', marginLeft, marginTop + 18);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(10);
      

      // ===== 1. General Information =====
      sectionTitle('1. General Information');
      writeWrapSmart('Requester (Name/Email)',
        `${headerDraft.requester?.text || '(unknown)'}  (${headerDraft.requesterEmail || ''})`);
      writeWrapSmart('Request Date', headerDraft.requestDate);
      writeWrapSmart('Required by date', headerDraft.requiredByDate);
      writeWrapSmart('Area', headerDraft.area);
      writeWrapSmart('Project', headerDraft.areaProject);
      writeWrapSmart('GL Code', headerDraft.glCode);
      writeWrapSmart('Company', headerDraft.companyValue);
      writeWrapSmart('SRDED', headerDraft.srded ? 'Yes' : 'No');
      writeWrapSmart('CMIF', headerDraft.cmif ? 'Yes' : 'No');
      writeWrapSmart('Budget Type', headerDraft.budgetType);

      writeWrapSmart('Priority', headerDraft.priority);
      writeWrapSmart('Type', headerDraft.reqType);

      // ===== 2. Suggested Suppliers List (B) =====
      sectionTitle('2. Suggested Suppliers List');

      const suppliersToShow = (pdfSuppliers || []).filter(s => s.name || s.contact || s.email);
      if (suppliersToShow.length > 0) {
        const startY = currentY + 4;
        autoTable(doc, {
          startY,
          head: [['Supplier Name', 'Contact', 'Email']],
          body: suppliersToShow.map(s => [
            s.name || '',
            s.contact || '',
            s.email || ''
          ]),
          theme: 'grid',
          styles: { font: 'helvetica', fontSize: 10, cellPadding: 4, halign: 'center', valign: 'middle' },
          headStyles: { fillColor: [230, 230, 230], textColor: 20, fontStyle: 'bold' },
          columnStyles: {
            0: { cellWidth: 170 },
            1: { cellWidth: 170 },
            2: { cellWidth: 170 },
          },
          margin: { left: marginLeft, right: RIGHT_MARGIN }
        });
        // @ts-ignore
        currentY = (doc as any).lastAutoTable.finalY + 16;
      } else {
        writeWrapSmart('Suppliers', 'No suggested suppliers provided.');
      }

      // ===== 3. Items/Services =====
      sectionTitle('3. Items/Services');
      const itemsStartY = currentY;

      autoTable(doc, {
        startY: itemsStartY,
        head: [['Item Description / Specification', 'SKU/Part', 'Qty', 'UoM', 'Unit Price', 'Currency', 'Total']],
        body: (pdfLines || []).map(l => [
          l.description || '',
          l.sku || '',
          String(l.qty ?? ''),
          l.uom || '',
          currencyFmt(l.unitPrice),
          l.currency || 'USD',
          currencyFmt(l.total || 0)
        ]),
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 10, cellPadding: 4, halign: 'center', valign: 'middle' },
        headStyles: { fillColor: [230, 230, 230], textColor: 20, fontStyle: 'bold' },
        columnStyles: {
          0: { cellWidth: 170 },
          1: { cellWidth: 70 },
          2: { cellWidth: 40 },
          3: { cellWidth: 45 },
          4: { cellWidth: 70 },
          5: { cellWidth: 60 },
          6: { cellWidth: 70 }
        },
        margin: { left: marginLeft, right: RIGHT_MARGIN }
      });
      // @ts-ignore
      currentY = (doc as any).lastAutoTable.finalY + 16;

      const subTotalPdf = (pdfLines || []).reduce((s, l) => s + (safeNum(l.qty) * safeNum(l.unitPrice)), 0);
      const grandPdf = subTotalPdf;

      const RESERVED_FOR_TOTALS = 90;
      const RESERVED_FOR_NEXT_HEADER = 40;
      const RESERVED_FOR_ONE_LINE = 24;
      const NEED_BEFORE_SECTION4 = RESERVED_FOR_TOTALS + RESERVED_FOR_NEXT_HEADER + RESERVED_FOR_ONE_LINE;
      if (currentY + NEED_BEFORE_SECTION4 > pageHeight - bottomMargin) {
        doc.addPage();
        currentY = marginTop;
      }

      autoTable(doc, {
        startY: currentY,
        head: [['Subtotal', 'Grand Total']],
        body: [[currencyFmt(subTotalPdf), currencyFmt(grandPdf)]],
        theme: 'plain',
        styles: { font: 'helvetica', fontSize: 12, halign: 'right', cellPadding: 3 },
        headStyles: { fontStyle: 'bold' },
        margin: { left: pageWidth - RIGHT_MARGIN - 280, right: RIGHT_MARGIN }
      });
      // @ts-ignore
      currentY = (doc as any).lastAutoTable.finalY + 18;

      // ===== 4. Business Justification (D) =====
      if (currentY + 60 > pageHeight - bottomMargin) { doc.addPage(); currentY = marginTop; }
      sectionTitle('4. Business Justification');
      const yesNo = (v?: boolean | string) => (typeof v === 'string' ? v : v ? 'Yes' : 'No');

      writeWrapSmart('Need / Objective', headerDraft.needObjective);
      // Impact if not purchased no se imprime
      writeWrapSmart('Sole-source', yesNo(headerDraft.soleSource === 'Yes'));
      writeWrapSmart('Sole-source explanation', headerDraft.soleSourceExplanation || '-');
      writeWrapSmart('Attachments - Quote', yesNo(!!headerDraft.attachQuote));
      writeWrapSmart('Attachments - SOW/Specification', yesNo(!!headerDraft.attachSOW));

      // ===== 5. Approvals =====
      doc.addPage();
      currentY = marginTop;
      sectionTitle('5. Approvals');

      const approvalRows: Array<{ role: string; who?: string; date?: string }> = [
        { role: 'Requester',   who: `${approverInfo.requester.name || ''}${approverInfo.requester.email ? ` (${approverInfo.requester.email})` : ''}` },
        { role: 'Supervisor',  who: `${approverInfo.supervisor.name || ''}${approverInfo.supervisor.email ? ` (${approverInfo.supervisor.email})` : ''}`, date: headerDraft.supervisorDate },
        { role: 'Staff Manager',  who: `${approverInfo.staffManager.name || ''}${approverInfo.staffManager.email ? ` (${approverInfo.staffManager.email})` : ''}`, date: headerDraft.staffManagerDate },
        { role: 'Manager',  who: `${approverInfo.manager.name || ''}${approverInfo.manager.email ? ` (${approverInfo.manager.email})` : ''}`, date: headerDraft.managerDate },
        { role: 'Director',  who: `${approverInfo.director.name || ''}${approverInfo.director.email ? ` (${approverInfo.director.email})` : ''}`, date: headerDraft.directorDate },
        { role: 'VP',  who: `${approverInfo.vp.name || ''}${approverInfo.vp.email ? ` (${approverInfo.vp.email})` : ''}`, date: headerDraft.vpDate },
        { role: 'CFO',  who: `${approverInfo.cfo.name || ''}${approverInfo.cfo.email ? ` (${approverInfo.cfo.email})` : ''}`, date: headerDraft.cfoDate },
        { role: 'CEO',  who: `${approverInfo.ceo.name || ''}${approverInfo.ceo.email ? ` (${approverInfo.ceo.email})` : ''}`, date: headerDraft.ceoDate },
        { role: 'Procurement',  who: `${approverInfo.procurement.name || ''}${approverInfo.procurement.email ? ` (${approverInfo.procurement.email})` : ''}`, date: headerDraft.procurementDate },
        { role: 'Finance',  who: `${approverInfo.finance.name || ''}${approverInfo.finance.email ? ` (${approverInfo.finance.email})` : ''}`, date: headerDraft.financeDate }
      ];

      autoTable(doc, {
        startY: currentY + 4,
        head: [['Role', 'Name / Email', 'Date']],
        body: approvalRows.map(r => [r.role, r.who || '-', r.date || '']),
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 11, cellPadding: 6, valign: 'middle' },
        headStyles: { fillColor: [230, 230, 230], textColor: 20, fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 120 }, 1: { cellWidth: 260 }, 2: { cellWidth: 120 } },
        margin: { left: marginLeft, right: RIGHT_MARGIN }
      });

      doc.save(`R${itemId}.pdf`);
      showSnack('PDF generated.', 'success');
    } catch (err: any) {
      showSnack(`Error generating PDF: ${err.message}`, 'error');
    } finally {
      setPdfLoading(false);
    }
  };

  /* =================== UI helpers =================== */
  // Non-picker inputs
  const onHeaderText = (k: keyof PRHeader) =>
    (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
      const v = e.target.value;
      setHeaderDraft(prev => ({ ...prev, [k]: v as any }));
    };

  const onHeaderSelect = (k: keyof PRHeader) =>
    (e: React.ChangeEvent<HTMLSelectElement>) => {
      const v = e.target.value;
      setHeaderDraft(prev => ({ ...prev, [k]: v as any }));
    };

  const onHeaderCheck = (k: keyof PRHeader) =>
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const v = e.target.checked as any;
      setHeaderDraft(prev => ({ ...prev, [k]: v }));
    };

  const onHeaderRadio = <K extends keyof PRHeader>(k: K, val: PRHeader[K]) =>
    () => { setHeaderDraft(prev => ({ ...prev, [k]: val })); };

  // PeoplePicker
  const onPicker =
    (k: keyof PRHeader) =>
      async (items: any[]) => {
        const picked = items?.[0] ?? null;
        setHeaderDraft(prev => ({ ...prev, [k]: picked }));

        if (k === 'requester' && picked) {
          try {
            let email = picked.secondaryText || picked.text || '';
            let userId: number | null = null;
            const login = picked.secondaryText || picked.loginName || picked.text || '';
            if (login) userId = await ensureUserIdSmart(login);
            if (userId) {
              try {
                const u = await sp.web.getUserById(userId)();
                email = u?.Email || email;
              } catch { /* ignore */ }
            }
            if (email) {
              setHeaderDraft(prev => ({ ...prev, requesterEmail: email }));
            }
          } catch { /* ignore */ }
        }
      };

  const onAttach = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fs = Array.from(e.target.files || []);
    setFiles(prev => [...prev, ...fs]);
  };

  const addLine = () =>
    setLines(ls => {
      const maxId = ls.reduce((m, it) => Math.max(m, it.id ?? 0), 0);
      return [...ls, { id: maxId + 1, qty: 1, unitPrice: 0, tax: 0, currency: 'USD' }];
    });

  const calcLineTotal = (l: PRLine) => {
    const sub = safeNum(l.qty) * safeNum(l.unitPrice);
    return Number(sub.toFixed(2));
  };

  const updateLine = (id: number, patch: Partial<PRLine>) =>
    setLines(ls => ls.map(l => l.id === id ? { ...l, ...patch, total: calcLineTotal({ ...l, ...patch }) } : l));

  const removeLine = (id: number) => setLines(ls => ls.filter(l => l.id !== id));

  /** Acciones para proveedores sugeridos */
  const addSupplier = () =>
    setSuppliers(sups => {
      const maxId = sups.reduce((m, it) => Math.max(m, it.id ?? 0), 0);
      return [...sups, { id: maxId + 1, name: '', contact: '', email: '' }];
    });

  const updateSupplier = (id: number, patch: Partial<SupplierLine>) =>
    setSuppliers(sups => sups.map(s => s.id === id ? { ...s, ...patch } : s));

  const removeSupplier = (id: number) =>
    setSuppliers(sups => sups.filter(s => s.id !== id));

  React.useEffect(() => {
    setLines(ls => ls.map(l => ({ ...l, total: calcLineTotal(l) })));
  }, []);

  const Section = (p: { title: string; code?: string; children: React.ReactNode }) => (
    <section className={styles.sectionCard}>
      <div className={styles.sectionHead}>
        <div className={styles.secTitle}>{p.title}</div>
        {p.code ? <div className={styles.secCode}>{p.code}</div> : null}
      </div>
      {p.children}
    </section>
  );

  const isLocked = hasAnyApproval;

  const PersonView = ({ label, person }: { label: string; person?: IPeoplePickerUserItem | null }) => (
    <div className={styles.fieldGroup}>
      <label className={styles.fieldLabel}>{label}</label>
      <div className={styles.readonlyBox}>
        {person?.text || person?.secondaryText || '(not assigned)'}
      </div>
    </div>
  );

  const DateInput = ({ label, k }: { label: string; k: keyof PRHeader }) => {
    const isNew = !editingItemId;
    if (isNew) return null;
    return (
      <div className={styles.fieldGroup}>
        <label className={styles.fieldLabel}>{label}</label>
        <input
          className={styles.input}
          type="date"
          value={(headerDraft[k] as string) || ''}
          onChange={onHeaderText(k)}
          disabled={true}
        />
      </div>
    );
  };

  const renderForm = (readOnlyMode: null | 'mysent' | 'tosign') => {
    const roAll = readOnlyMode !== null;
    const isToSign = readOnlyMode === 'tosign';
    const enableApprover = (role: RoleKey) => isToSign && (signFormRoles || []).includes(role);

    // Determinar qué campos de aprobación mostrar
   const showApprovalFields = (role: RoleKey): boolean => {
  // Siempre mostrar supervisor
  if (role === 'supervisor') return true;

  // Helper: check if role has an assigned person
  const hasPersonAssigned = (r: RoleKey): boolean => {
    switch (r) {
      case 'staffManager': return !!headerDraft.staffManager;
      case 'staffManager2': return !!headerDraft.staffManager2;
      case 'manager': return !!headerDraft.manager;
      case 'manager2': return !!headerDraft.manager2;
      case 'director': return !!headerDraft.director;
      case 'vp': return !!headerDraft.vp;
      case 'cfo': return !!headerDraft.cfo;
      case 'ceo': return !!headerDraft.ceo;
      case 'procurement': return !!headerDraft.procurement;
      case 'finance': return !!headerDraft.finance;
      default: return false;
    }
  };

  // Si es modo lectura (mysent / tosign), solo mostrar roles con persona asignada
  if (roAll) return hasPersonAssigned(role);
  
  // Si estamos en modo nuevo, mostrar solo los requeridos
  if (!editingItemId) {
    const totalAmount = lines.reduce((s, l) => s + safeNum(l.qty) * safeNum(l.unitPrice), 0);
    const requiredRoles = getRequiredAuthorizers(headerDraft.area || '', headerDraft.budgetType || 'Budgeted', totalAmount);
    return requiredRoles.includes(role);
  }
  
  // Si estamos editando, mostrar el campo si ya tiene un valor asignado
  return hasPersonAssigned(role);
};
    return (
      <>
        {readOnlyMode === null && isLocked && (
          <div className={styles.lockBanner}>
            This PR already has approvals and is locked for editing.
          </div>
        )}
        <fieldset disabled={readOnlyMode === null ? isLocked : false} style={{ border: 'none', padding: 0, margin: 0 }}>
          {/* A. Request Details */}
          <Section title={`A. Request Details ${editingItemId ? `(PR #${editingItemId})` : ''}`} code="A">
            <div className={styles.grid2}>
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>Requester (Person)</label>
                {roAll ? (
                  <div className={styles.readonlyBox}>
                    {headerDraft.requester?.text || headerDraft.requester?.secondaryText || '(not assigned)'}
                  </div>
                ) : (
                  <div className={styles.peoplePicker}>
                    <PeoplePicker
                      context={peopleCtx}
                      personSelectionLimit={1}
                      ensureUser={true}
                      showtooltip={true}
                      principalTypes={[PrincipalType.User]}
                      onChange={onPicker('requester')}
                      defaultSelectedUsers={
                        headerDraft.requester ? [headerDraft.requester.secondaryText || headerDraft.requester.loginName || headerDraft.requester.text].filter(Boolean) as string[] : []
                      }
                    />
                  </div>
                )}
              </div>
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>Requester Email</label>
                <input
                  id="fld-requesterEmail"
                  className={styles.input}
                  value={headerDraft.requesterEmail || ''}
                  onChange={keepFocus('fld-requesterEmail', onHeaderText('requesterEmail'))}
                  disabled={roAll}
                />
              </div>
            </div>

            <div className={styles.grid3}>
              <div className={styles.fieldGroupSelect}>
                <label className={styles.fieldLabel}>Area</label>
                <div className={styles.selectContainer}>
                  <select
                    id="fld-area"
                    className={styles.select}
                    value={headerDraft.area || ''}
                    onChange={onHeaderSelect('area')}
                    disabled={roAll}
                  >
                    <option value="">Select Area</option>
                    {AREA_OPTIONS.map(area => (
                      <option key={area} value={area}>{area}</option>
                    ))}
                  </select>
                  <div className={styles.selectArrow}>
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                      <path fillRule="evenodd" d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" clipRule="evenodd" />
                    </svg>
                  </div>
                </div>
              </div>
              
              <div className={styles.fieldGroupSelect}>
                <label className={styles.fieldLabel}>Project</label>
                <div className={styles.selectContainer} ref={projWrapRef} style={{ position: 'relative', overflow: 'visible' }}>
                  <input
                    ref={projInputRef}
                    id="fld-areaProject"
                    className={styles.select}
                    value={projOpen ? projSearch : (projDisplayValue || '')}
                    placeholder="Search / Select Project"
                    onChange={keepFocus('fld-areaProject', (e) => {
                      if (roAll) return;
                      setProjSearch(e.target.value);
                      if (!projOpen) setProjOpen(true);
                      setGlOpen(false);
                      setCompanyOpen(false);
                    })}
                    onFocus={() => {
                      if (roAll) return;
                      setProjOpen(true);
                      setGlOpen(false);
                      setCompanyOpen(false);
                      setProjPage(1);
                    }}
                    disabled={roAll}
                    autoComplete="off"
                  />

                  <div
                    className={styles.selectArrow}
                    onMouseDown={(e) => {
                      e.preventDefault();
                      if (roAll) return;
                      setProjOpen((o) => {
                        const next = !o;
                        if (next) {
                          setGlOpen(false);
                          setCompanyOpen(false);
                          setProjPage(1);
                        } else {
                          setProjSearch('');
                        }
                        return next;
                      });
                    }}
                    style={{ cursor: roAll ? 'default' : 'pointer' }}
                    role="button"
                    aria-label="Toggle Project dropdown"
                  >
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                      <path
                        fillRule="evenodd"
                        d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z"
                        clipRule="evenodd"
                      />
                    </svg>
                  </div>

                  {/* Botón para limpiar proyecto */}
                  {!roAll && headerDraft.areaProject && (
                    <button
                      type="button"
                      className={styles.clearBtn}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => clearProject()}
                      aria-label="Clear Project"
                    >
                      ✕
                    </button>
                  )}

                  {loadingProjects && <div className={styles.selectHint}>Loading projects...</div>}

                  {projOpen && !roAll && (
                    <div
                      style={{
                        position: 'absolute',
                        zIndex: 50,
                        left: 0,
                        right: 0,
                        top: projOpenUp ? undefined : 'calc(100% + 6px)',
                        bottom: projOpenUp ? 'calc(100% + 6px)' : undefined,
                        background: 'white',
                        border: '1px solid rgba(0,0,0,0.12)',
                        borderRadius: 10,
                        boxShadow: '0 8px 20px rgba(0,0,0,0.12)',
                        maxHeight: 360,
                        display: 'flex',
                        flexDirection: 'column',
                        overflow: 'hidden'
                      }}
                    >
                      <div style={{ padding: 8, borderBottom: '1px solid rgba(0,0,0,0.08)' }}>
                        <div
                          style={{
                            display: 'grid',
                            gridTemplateColumns: '1fr 140px 1.4fr',
                            gap: 10,
                            fontSize: 12,
                            fontWeight: 700,
                            opacity: 0.8
                          }}
                        >
                          <div>Project Name</div>
                          <div>Project Code</div>
                          <div>Description</div>
                        </div>
                      </div>

                      <div style={{ flex: 1, overflow: 'auto' }}>
                        {projPaged.length === 0 ? (
                          <div style={{ padding: 10, fontSize: 13, opacity: 0.75 }}>No results</div>
                        ) : (
                          projPaged.map((p) => {
                            const isSel = (headerDraft.areaProject || '') === p.Title;
                            return (
                              <button
                                key={p.Title}
                                type="button"
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => pickProject(p)}
                                style={{
                                  width: '100%',
                                  textAlign: 'left',
                                  padding: '10px 8px',
                                  border: 'none',
                                  background: isSel ? 'rgba(0,0,0,0.05)' : 'transparent',
                                  cursor: 'pointer'
                                }}
                              >
                                <div
                                  style={{
                                    display: 'grid',
                                    gridTemplateColumns: '1fr 140px 1.4fr',
                                    gap: 10,
                                    fontSize: 13,
                                    lineHeight: 1.2
                                  }}
                                >
                                  <div>{p.Title}</div>
                                  <div
                                    style={{
                                      fontFamily:
                                        'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
                                    }}
                                  >
                                    {p.ProjectCode || ''}
                                  </div>
                                  <div>{p.ProjectDescription || ''}</div>
                                </div>
                              </button>
                            );
                          })
                        )}
                      </div>

                      <div
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'space-between',
                          gap: 8,
                          padding: 8,
                          borderTop: '1px solid rgba(0,0,0,0.08)',
                          background: 'white',
                          flexShrink: 0
                        }}
                      >
                        <button type="button" onMouseDown={(e) => e.preventDefault()}
                          disabled={projPage <= 1}
                          onClick={() => setProjPage((p) => Math.max(1, p - 1))}>
                          ◀ Prev
                        </button>

                        <div style={{ fontSize: 12, opacity: 0.8 }}>
                          Page {Math.min(projPage, projTotalPages)} / {projTotalPages} • {projFiltered.length} results
                        </div>

                        <button
                          type="button"
                          onMouseDown={(e) => e.preventDefault()}
                          disabled={projPage >= projTotalPages}
                          onClick={() => setProjPage((p) => Math.min(projTotalPages, p + 1))}
                        >
                          Next ▶
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>

              <div className={styles.fieldGroupSelect}>
                <label className={styles.fieldLabel}>GL Code</label>
                <div className={styles.selectContainer} ref={glWrapRef} style={{ position: 'relative', overflow: 'visible' }}>
                  <input
                    ref={glInputRef}
                    id="fld-glCode"
                    className={styles.select}
                    value={glOpen ? glSearch : (glDisplayValue || '')}
                    placeholder="Search / Select GL Code"
                    onChange={keepFocus('fld-glCode', (e) => {
                      if (roAll) return;
                      setGlSearch(e.target.value);
                      if (!glOpen) setGlOpen(true);
                      setProjOpen(false);
                      setCompanyOpen(false);
                    })}
                    onFocus={() => {
                      if (roAll) return;
                      setGlOpen(true);
                      setProjOpen(false);
                      setCompanyOpen(false);
                      setGlPage(1);
                      setGlNoPagination(false);
                    }}
                    disabled={roAll}
                    autoComplete="off"
                  />

                  <div
                    className={styles.selectArrow}
                    onMouseDown={(e) => {
                      e.preventDefault();
                      if (roAll) return;
                      setGlOpen((o) => {
                        const next = !o;
                        if (next) {
                          setProjOpen(false);
                          setCompanyOpen(false);
                          setGlPage(1);
                        } else {
                          setGlSearch('');
                          setGlNoPagination(false);
                        }
                        return next;
                      });
                    }}
                    style={{ cursor: roAll ? 'default' : 'pointer' }}
                    role="button"
                    aria-label="Toggle GL Code dropdown"
                  >
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                      <path
                        fillRule="evenodd"
                        d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z"
                        clipRule="evenodd"
                      />
                    </svg>
                  </div>

                  {loadingGlCodes && <div className={styles.selectHint}>Loading GL codes...</div>}

     {glOpen && !roAll && (
  <div
    style={{
      position: 'absolute',
      zIndex: 50,
      left: 0,
      right: 0,
      top: glOpenUp ? undefined : 'calc(100% + 6px)',
      bottom: glOpenUp ? 'calc(100% + 6px)' : undefined,
      width: 'clamp(120px, 90vw, 580px)',
      maxWidth: '98vw',
      background: 'white',
      border: '1px solid rgba(0,0,0,0.12)',
      borderRadius: 12,
      boxShadow: '0 18px 45px rgba(0,0,0,0.15)',
      maxHeight: glNoPagination ? '40vh' : 460,
      display: 'flex',
      flexDirection: 'column',
      overflow: 'hidden'
    }}
  >
    {(() => {
      const glGridCols = '140px minmax(150px, 1fr) minmax(150px, 1fr) minmax(180px, 1fr)';

      return (
        <>
          {/* ✅ HEADER SOLO CON TÍTULOS - SIN BOTÓN */}
          <div style={{ 
            padding: 10, 
            borderBottom: '1px solid rgba(0,0,0,0.08)',
            display: 'grid',
            gridTemplateColumns: glGridCols,
            gap: 6,
            fontSize: 10,
            fontWeight: 700,
            opacity: 0.85,
            whiteSpace: 'nowrap',
            flexShrink: 0,
            backgroundColor: '#fafafa'
          }}>
            <div>Account Code</div>
            <div>Cost Center Name</div>
            <div>Activity Code Name</div>
            <div>Natural Account Name</div>
          </div>

          {/* ✅ CONTENIDO SCROLLEABLE */}
          <div style={{ 
            flex: 1, 
            overflowY: 'auto', 
            overflowX: 'auto',
            maxHeight: glNoPagination ? 'calc(8 * 40px)' : undefined 
          }}>
            {glPaged.length === 0 ? (
              <div style={{ padding: 12, fontSize: 13, opacity: 0.75 }}>No results</div>
            ) : (
              glPaged.map((gl) => {
                const isSel = (headerDraft.glCode || '') === gl.Title;
                return (
                  <button
                    key={gl.Title}
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => pickGl(gl)}
                    style={{
                      width: '100%',
                      textAlign: 'left',
                      padding: '10px 10px',
                      border: 'none',
                      background: isSel ? 'rgba(0,0,0,0.05)' : 'transparent',
                      cursor: 'pointer',
                      borderBottom: '1px solid rgba(0,0,0,0.05)',
                      display: 'grid',
                      gridTemplateColumns: glGridCols,
                      gap: 8,
                      fontSize: 13,
                      lineHeight: 1.2,
                      whiteSpace: 'nowrap',
                      alignItems: 'center'
                    }}
                  >
                    <div style={{
                      fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
                    }}>
                      {gl.Title}
                    </div>
                    <div>{gl.CostCenterName || ''}</div>
                    <div>{gl.ActivityCodeName || ''}</div>
                    <div>{gl.NaturalAccountName || ''}</div>
                  </button>
                );
              })
            )}
          </div>

          {/* ✅ FOOTER CON BOTÓN + PAGINACIÓN - COMPACTO */}
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              gap: 6,
              padding: '8px 10px',
              borderTop: '1px solid rgba(0,0,0,0.08)',
              background: 'white',
              flexShrink: 0,
              flexWrap: 'wrap'
            }}
          >
            {/* BOTÓN DE PAGINACIÓN */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setGlNoPagination(!glNoPagination)}
              style={{
                background: glNoPagination ? '#4f46e5' : '#f3f4f6',
                color: glNoPagination ? 'white' : '#374151',
                border: glNoPagination ? 'none' : '1px solid #d1d5db',
                borderRadius: 4,
                padding: '5px 10px',
                fontSize: 10,
                fontWeight: 600,
                cursor: 'pointer',
                whiteSpace: 'nowrap',
                transition: 'all 0.2s',
                order: 1
              }}
            >
              {glNoPagination ? '▲ Paginated' : '▼ All'}
            </button>

            {/* CONTROLES DE PAGINACIÓN (solo si NOT sin paginación) */}
            {!glNoPagination && (
              <>
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={glPage <= 1}
                  onClick={() => setGlPage((p) => Math.max(1, p - 1))}
                  style={{ order: 2, padding: '4px 8px', fontSize: 10 }}
                >
                  ◀
                </button>

                <div style={{ fontSize: 10, opacity: 0.8, order: 3, whiteSpace: 'nowrap' }}>
                  {Math.min(glPage, glTotalPages)}/{glTotalPages}
                </div>

                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={glPage >= glTotalPages}
                  onClick={() => setGlPage((p) => Math.min(glTotalPages, p + 1))}
                  style={{ order: 4, padding: '4px 8px', fontSize: 10 }}
                >
                  ▶
                </button>
              </>
            )}

            {/* MENSAJE SIN PAGINACIÓN */}
            {glNoPagination && glFiltered.length > 0 && (
              <div
                style={{
                  fontSize: 10,
                  color: '#6b7280',
                  textAlign: 'center',
                  width: '100%',
                  order: 5
                }}
              >
                All {glFiltered.length} • Scroll ↓
              </div>
            )}
          </div>
        </>
      );
    })()}
  </div>
)}
                </div>
              </div>
            </div>

            <div className={styles.grid3}>
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>Request Date</label>
                <input
                  id="fld-requestDate"
                  className={styles.input}
                  type="date"
                  value={headerDraft.requestDate || ''}
                  onChange={keepFocus('fld-requestDate', onHeaderText('requestDate'))}
                  disabled={roAll}
                />
              </div>
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>Required by date</label>
                <input
                  id="fld-requiredByDate"
                  className={styles.input}
                  type="date"
                  value={headerDraft.requiredByDate || ''}
                  onChange={keepFocus('fld-requiredByDate', onHeaderText('requiredByDate'))}
                  disabled={roAll}
                />
              </div>
              <div className={styles.fieldGroupSelect}>
                <label className={styles.fieldLabel}>Company</label>
                <div className={styles.selectContainer} ref={companyWrapRef} style={{ position: 'relative', overflow: 'visible' }}>
                  <input
                    ref={companyInputRef}
                    id="fld-company"
                    className={styles.select}
                    value={companyOpen ? companySearch : (companyDisplayValue || '')}
                    placeholder="Search / Select Company"
                    onChange={keepFocus('fld-company', (e) => {
                      if (roAll) return;
                      setCompanySearch(e.target.value);
                      if (!companyOpen) setCompanyOpen(true);
                      setProjOpen(false);
                      setGlOpen(false);
                    })}
                    onFocus={() => {
                      if (roAll) return;
                      setCompanyOpen(true);
                      setProjOpen(false);
                      setGlOpen(false);
                      setCompanyPage(1);
                    }}
                    disabled={roAll}
                    autoComplete="off"
                  />

                  <div
                    className={styles.selectArrow}
                    onMouseDown={(e) => {
                      e.preventDefault();
                      if (roAll) return;
                      setCompanyOpen((o) => {
                        const next = !o;
                        if (next) {
                          setProjOpen(false);
                          setGlOpen(false);
                          setCompanyPage(1);
                        } else {
                          setCompanySearch('');
                        }
                        return next;
                      });
                    }}
                    style={{ cursor: roAll ? 'default' : 'pointer' }}
                    role="button"
                    aria-label="Toggle Company dropdown"
                  >
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                      <path
                        fillRule="evenodd"
                        d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z"
                        clipRule="evenodd"
                      />
                    </svg>
                  </div>

                  {/* Botón para limpiar company */}
                  {!roAll && headerDraft.companyValue && (
                    <button
                      type="button"
                      className={styles.clearBtn}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => clearCompany()}
                      aria-label="Clear Company"
                    >
                      ✕
                    </button>
                  )}

                  {loadingCompanies && <div className={styles.selectHint}>Loading companies...</div>}

                  {companyOpen && !roAll && (
                    <div
                      style={{
                        position: 'absolute',
                        zIndex: 50,
                        left: 0,
                        right: 0,
                        top: companyOpenUp ? undefined : 'calc(100% + 6px)',
                        bottom: companyOpenUp ? 'calc(100% + 6px)' : undefined,
                        background: 'white',
                        border: '1px solid rgba(0,0,0,0.12)',
                        borderRadius: 10,
                        boxShadow: '0 8px 20px rgba(0,0,0,0.12)',
                        maxHeight: 360,
                        display: 'flex',
                        flexDirection: 'column',
                        overflow: 'hidden'
                      }}
                    >
                      <div style={{ padding: 8, borderBottom: '1px solid rgba(0,0,0,0.08)' }}>
                        <div
                          style={{
                            display: 'grid',
                            gridTemplateColumns: '1fr 140px 1.4fr',
                            gap: 10,
                            fontSize: 12,
                            fontWeight: 700,
                            opacity: 0.8
                          }}
                        >
                          <div>Company Code</div>
                          <div>Code</div>
                          <div>Company Name</div>
                        </div>
                      </div>

                      <div style={{ flex: 1, overflow: 'auto' }}>
                        {companyPaged.length === 0 ? (
                          <div style={{ padding: 10, fontSize: 13, opacity: 0.75 }}>No results</div>
                        ) : (
                          companyPaged.map((c) => {
                            const isSel = (headerDraft.companyValue || '') === c.Title;
                            return (
                              <button
                                key={c.Title}
                                type="button"
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => pickCompany(c)}
                                style={{
                                  width: '100%',
                                  textAlign: 'left',
                                  padding: '10px 8px',
                                  border: 'none',
                                  background: isSel ? 'rgba(0,0,0,0.05)' : 'transparent',
                                  cursor: 'pointer'
                                }}
                              >
                                <div
                                  style={{
                                    display: 'grid',
                                    gridTemplateColumns: '1fr 140px 1.4fr',
                                    gap: 10,
                                    fontSize: 13,
                                    lineHeight: 1.2
                                  }}
                                >
                                  <div>{c.CompanyCodeforGLAccounts}</div>
                                  <div
                                    style={{
                                      fontFamily:
                                        'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace'
                                    }}
                                  >
                                    {c.ProntoCompanyName || ''}
                                  </div>
                                  <div>{c.CompanyName || ''}</div>
                                </div>
                              </button>
                            );
                          })
                        )}
                      </div>

                      <div
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'space-between',
                          gap: 8,
                          padding: 8,
                          borderTop: '1px solid rgba(0,0,0,0.08)',
                          background: 'white',
                          flexShrink: 0
                        }}
                      >
                        <button type="button" onMouseDown={(e) => e.preventDefault()}
                          disabled={companyPage <= 1}
                          onClick={() => setCompanyPage((p) => Math.max(1, p - 1))}>
                          ◀ Prev
                        </button>

                        <div style={{ fontSize: 12, opacity: 0.8 }}>
                          Page {Math.min(companyPage, companyTotalPages)} / {companyTotalPages} • {companyFiltered.length} results
                        </div>

                        <button
                          type="button"
                          onMouseDown={(e) => e.preventDefault()}
                          disabled={companyPage >= companyTotalPages}
                          onClick={() => setCompanyPage((p) => Math.min(companyTotalPages, p + 1))}
                        >
                          Next ▶
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Nuevos campos checkbox SRDED y CMIF */}
            <div className={styles.grid2}>
              <div className={styles.fieldGroup}>
                <div className={styles.fieldLabel}>Type set</div>
                <div className={styles.checkboxRow}>
                  <label className={styles.checkboxOption}>
                    <input 
                      type="checkbox" 
                      checked={headerDraft.srded || false} 
                      onChange={onHeaderCheck('srded')} 
                      disabled={roAll} 
                    />
                    <span>SR&ED?</span>
                  </label>
                  <label className={styles.checkboxOption}>
                    <input 
                      type="checkbox" 
                      checked={headerDraft.cmif || false} 
                      onChange={onHeaderCheck('cmif')} 
                      disabled={roAll} 
                    />
                    <span>CMIF</span>
                  </label>
                </div>
              </div>
            </div>

            <div className={styles.grid3}>
              <div className={styles.fieldGroup}>
                <div className={styles.fieldLabel}>Priority</div>
                <div className={styles.radioRow}>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.priority === 'Normal'} onChange={onHeaderRadio('priority', 'Normal')} disabled={roAll} />
                    <span>Normal</span>
                  </label>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.priority === 'Urgent'} onChange={onHeaderRadio('priority', 'Urgent')} disabled={roAll} />
                    <span>Urgent</span>
                  </label>
                </div>
              </div>
              {/* Type movido al lado de Priority */}
              <div className={styles.fieldGroup}>
                <div className={styles.fieldLabel}>Type</div>
                <div className={styles.radioRow}>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.reqType === 'Goods'} onChange={onHeaderRadio('reqType', 'Goods')} disabled={roAll} />
                    <span>Goods</span>
                  </label>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.reqType === 'Services'} onChange={onHeaderRadio('reqType', 'Services')} disabled={roAll} />
                    <span>Services</span>
                  </label>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.reqType === 'Software/License'} onChange={onHeaderRadio('reqType', 'Software/License')} disabled={roAll} />
                    <span>Software/License</span>
                  </label>
                </div>
              </div>
              <div className={styles.fieldGroup}>
                <div className={styles.fieldLabel}>Budget Type</div>
                <div className={styles.radioRow}>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.budgetType === 'Budgeted'} onChange={onHeaderRadio('budgetType', 'Budgeted')} disabled={roAll} />
                    <span>Budgeted</span>
                  </label>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.budgetType === 'Non-Budgeted'} onChange={onHeaderRadio('budgetType', 'Non-Budgeted')} disabled={roAll} />
                    <span>Non-Budgeted</span>
                  </label>
                </div>
              </div>
            </div>

            <div className={styles.grid2}>
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>Urgency justification</label>
                <input
                  id="fld-urgentJustification"
                  className={styles.input}
                  value={headerDraft.urgentJustification || ''}
                  onChange={keepFocus('fld-urgentJustification', onHeaderText('urgentJustification'))}
                  placeholder="Only if urgent"
                  disabled={roAll}
                />
              </div>
            </div>
          </Section>

          {/* B. Suggested Suppliers List */}
          <Section title="B. Suggested Suppliers List" code="B">
            <div className={styles.miniHint} style={{ marginBottom: 8 }}>
              Add one or more <b>suggested suppliers</b>. For each supplier, include Name, Contact (optional) and Email (optional).
            </div>

            <div className={styles.itemsTableWrap}>
              <div className={styles.itemsHeaderRow}>
                <div>Supplier Name</div>
                <div>Supplier Contact (optional)</div>
                <div>Supplier Email (optional)</div>
                <div />
              </div>

              {suppliers.map(s => (
                <div className={styles.itemsRow} key={s.id}>
                  <input
                    id={`sup-name-${s.id}`}
                    className={styles.input}
                    placeholder="Supplier Name"
                    value={s.name || ''}
                    onChange={keepFocus(`sup-name-${s.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateSupplier(s.id!, { name: e.target.value }))}
                    readOnly={roAll}
                  />
                  <input
                    id={`sup-contact-${s.id}`}
                    className={styles.input}
                    placeholder="Contact (optional)"
                    value={s.contact || ''}
                    onChange={keepFocus(`sup-contact-${s.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateSupplier(s.id!, { contact: e.target.value }))}
                    readOnly={roAll}
                  />
                  <input
                    id={`sup-email-${s.id}`}
                    className={styles.input}
                    type="email"
                    placeholder="Email (optional)"
                    value={s.email || ''}
                    onChange={keepFocus(`sup-email-${s.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateSupplier(s.id!, { email: e.target.value }))}
                    readOnly={roAll}
                  />
                  {!roAll && (
                    <button
                      type="button"
                      className={styles.removeBtn}
                      onClick={() => removeSupplier(s.id!)}
                    >
                      ✕
                    </button>
                  )}
                </div>
              ))}

              {!roAll && (
                <div className={styles.itemsActions}>
                  <button type="button" className={styles.addBtn} onClick={addSupplier}>+ Add supplier</button>
                </div>
              )}
            </div>
          </Section>

          {/* C. Items/Services */}
          <Section title="C. Items/Services" code="C">
            <div className={styles.itemsTableWrap}>
              <div className={styles.itemsHeaderRow}>
                <div>Description / Specification</div>
                <div>SKU/Part</div>
                <div>Qty</div>
                <div>UoM</div>
                <div>Unit Price</div>
                <div>Currency</div>
                <div>Total</div>
                <div />
              </div>

              {lines.map(l => (
                <div className={styles.itemsRow} key={l.id}>
                  <input
                    id={`ln-desc-${l.id}`}
                    className={styles.input}
                    placeholder="Describe the item/service"
                    value={l.description || ''}
                    onChange={keepFocus(`ln-desc-${l.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateLine(l.id!, { description: e.target.value }))}
                    readOnly={roAll}
                  />
                  <input
                    id={`ln-sku-${l.id}`}
                    className={styles.input}
                    placeholder="SKU/Part"
                    value={l.sku || ''}
                    onChange={keepFocus(`ln-sku-${l.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateLine(l.id!, { sku: e.target.value }))}
                    readOnly={roAll}
                  />
                  <input
                    id={`ln-qty-${l.id}`}
                    className={styles.input}
                    type="number"
                    min={0}
                    value={l.qty ?? 0}
                    onChange={keepFocus(`ln-qty-${l.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateLine(l.id!, { qty: parseFloat(e.target.value || '0') }))}
                    readOnly={roAll}
                  />
                  <input
                    id={`ln-uom-${l.id}`}
                    className={styles.input}
                    placeholder="UoM"
                    value={l.uom || ''}
                    onChange={keepFocus(`ln-uom-${l.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateLine(l.id!, { uom: e.target.value }))}
                    readOnly={roAll}
                  />
                  <input
                    id={`ln-unitPrice-${l.id}`}
                    className={styles.input}
                    type="number"
                    step="0.01"
                    min={0}
                    value={l.unitPrice ?? 0}
                    onChange={keepFocus(`ln-unitPrice-${l.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateLine(l.id!, { unitPrice: parseFloat(e.target.value || '0') }))}
                    readOnly={roAll}
                  />
                  <input
                    id={`ln-currency-${l.id}`}
                    className={styles.input}
                    placeholder="USD"
                    value={l.currency || 'USD'}
                    onChange={keepFocus(`ln-currency-${l.id}`, (e: React.ChangeEvent<HTMLInputElement>) => updateLine(l.id!, { currency: e.target.value }))}
                    readOnly={roAll}
                  />
                  <input className={styles.input} readOnly value={currencyFmt(l.total || 0)} />
                  {!roAll && <button type="button" className={styles.removeBtn} onClick={() => removeLine(l.id!)}>✕</button>}
                </div>
              ))}

              {!roAll && (
                <div className={styles.itemsActions}>
                  <button type="button" className={styles.addBtn} onClick={addLine}>+ Add line</button>
                </div>
              )}

              <div className={styles.totalsRow}>
                <div /><div /><div /><div /><div /><div />
                <div className={styles.totalLabel}>Subtotal</div>
                <div className={styles.totalValue}>
                  {currencyFmt(lines.reduce((s, l) => s + safeNum(l.qty) * safeNum(l.unitPrice), 0))}
                </div>
              </div>
              <div className={styles.totalsRow}>
                <div /><div /><div /><div /><div /><div />
                <div className={styles.totalLabel}>Grand Total</div>
                <div className={styles.totalValue}>
                  {currencyFmt(
                    lines.reduce((s, l) => s + safeNum(l.qty) * safeNum(l.unitPrice), 0)
                  )}
                </div>
              </div>
            </div>

            <div className={styles.miniHint} style={{ marginTop: 8 }}>
              Note: If this is a <b>Service</b>, please enter the total amount in <b>Qty</b> with a unit price of <b>$1</b>.
            </div>
          </Section>

          {/* D. Business Justification & Attachments */}
          <Section title="D. Business Justification & Attachments" code="D">
            <div className={styles.fieldGroup}>
              <label className={styles.fieldLabel}>Need / Objective</label>
              <textarea
                id="fld-needObjective"
                className={styles.textarea}
                style={{ height: '536px' }}
                value={headerDraft.needObjective || ''}
                onChange={keepFocus('fld-needObjective', onHeaderText('needObjective'))}
                readOnly={roAll}
              />
            </div>

            {/* "Impact if not purchased" section removed de la UI */}

            <div className={styles.grid3}>
              <div className={styles.fieldGroup}>
                <div className={styles.fieldLabel}>Sole-source</div>
                <div className={styles.radioRow}>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.soleSource === 'Yes'} onChange={onHeaderRadio('soleSource', 'Yes')} disabled={roAll} />
                    <span>Yes</span>
                  </label>
                  <label className={styles.radioOption}>
                    <input type="radio" checked={headerDraft.soleSource === 'No'} onChange={onHeaderRadio('soleSource', 'No')} disabled={roAll} />
                    <span>No</span>
                  </label>
                </div>
              </div>
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>Sole-source explanation</label>
                <input
                  id="fld-soleSourceExplanation"
                  className={styles.input}
                  value={headerDraft.soleSourceExplanation || ''}
                  onChange={keepFocus('fld-soleSourceExplanation', onHeaderText('soleSourceExplanation'))}
                  disabled={roAll}
                />
              </div>
            </div>

            <div className={styles.grid3}>
              <label className={styles.checkboxRow}>
                <input type="checkbox" checked={!!headerDraft.attachQuote} onChange={onHeaderCheck('attachQuote')} disabled={roAll} />
                <span>Quote attached</span>
              </label>
            </div>
            <div className={styles.grid3}>
              <label className={styles.checkboxRow}>
                <input type="checkbox" checked={!!headerDraft.attachSOW} onChange={onHeaderCheck('attachSOW')} disabled={roAll} />
                <span>SOW/Specification attached</span>
              </label>
            </div>

            {readOnlyMode === null && (
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>Attachments (PDF/Images/Docs)</label>
                <input ref={attachInputRef} type="file" multiple onChange={onAttach} />
                {files.length > 0 && (
                  <div className={styles.miniHint}>{files.length} file(s) ready to upload</div>
                )}
              </div>
            )}
          </Section>

          {/* F. Approvals - MODIFICADO PARA MOSTRAR SOLO APROBADORES REQUERIDOS */}
          <Section title="F. Approvals" code="F">
            {/* Botón para auto-poblar aprobadores CON VALIDACIÓN */}
            {!roAll && (
              <div className={styles.fieldGroup}>
                <button 
                  type="button" 
                  className={styles.saveBtn} 
                  onClick={autoPopulateApprovers} 
                  style={{ marginBottom: '16px' }}
                  disabled={!headerDraft.area || !headerDraft.budgetType || roAll || !!editingItemId}
                >
                  🔄 Auto-Populate Required Approvers
                </button>
                {headerDraft.area && headerDraft.budgetType && !editingItemId && (
                  <div style={{ marginBottom: '16px', padding: '8px', backgroundColor: '#f5f5f5', borderRadius: '4px' }}>
                    <strong>Total Amount:</strong> {currencyFmt(lines.reduce((s, l) => s + safeNum(l.qty) * safeNum(l.unitPrice), 0))}
                  </div>
                )}
              </div>
            )}

            {/* Supervisor - SIEMPRE VISIBLE */}
            {showApprovalFields('supervisor') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Supervisor" person={headerDraft.supervisor} />
                  <div className={styles.grid3}>
                    <DateInput label="Supervisor Date" k="supervisorDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Supervisor Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.supervisorStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Supervisor</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('supervisor')}
                        defaultSelectedUsers={
                          headerDraft.supervisor ? [headerDraft.supervisor.secondaryText || headerDraft.supervisor.loginName || headerDraft.supervisor.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Supervisor Date" k="supervisorDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Supervisor Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.supervisorStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('supervisor') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('supervisor')}
                            >
                              Agree as Supervisor
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('supervisor')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* Staff Manager - SOLO SI SE REQUIERE */}
            {showApprovalFields('staffManager') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Staff Manager" person={headerDraft.staffManager} />
                  <div className={styles.grid3}>
                    <DateInput label="Staff Manager Date" k="staffManagerDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Staff Manager Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.staffManagerStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Staff Manager</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('staffManager')}
                        defaultSelectedUsers={
                          headerDraft.staffManager ? [headerDraft.staffManager.secondaryText || headerDraft.staffManager.loginName || headerDraft.staffManager.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Staff Manager Date" k="staffManagerDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Staff Manager Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.staffManagerStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('staffManager') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('staffManager')}
                            >
                              Agree as Staff Manager
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('staffManager')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* Staff2 - SOLO SI SE REQUIERE */}
            {showApprovalFields('staffManager2') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Staff2" person={headerDraft.staffManager2} />
                  <div className={styles.grid3}>
                    <DateInput label="Staff2 Date" k="staffManager2Date" />
                    <div className={styles.fieldGroup}>gulp
                      <label className={styles.fieldLabel}>Staff2 Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.staffManager2Status || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Staff2</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('staffManager2')}
                        defaultSelectedUsers={
                          headerDraft.staffManager2 ? [headerDraft.staffManager2.secondaryText || headerDraft.staffManager2.loginName || headerDraft.staffManager2.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Staff2 Date" k="staffManager2Date" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Staff2 Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.staffManager2Status || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('staffManager2') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('staffManager2')}
                            >
                              Agree as Staff2
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('staffManager2')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* Manager - SOLO SI SE REQUIERE */}
            {showApprovalFields('manager') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Manager" person={headerDraft.manager} />
                  <div className={styles.grid3}>
                    <DateInput label="Manager Date" k="managerDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Manager Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.managerStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Manager</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('manager')}
                        defaultSelectedUsers={
                          headerDraft.manager ? [headerDraft.manager.secondaryText || headerDraft.manager.loginName || headerDraft.manager.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Manager Date" k="managerDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Manager Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.managerStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('manager') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('manager')}
                            >
                              Agree as Manager
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('manager')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* Manager2 - SOLO SI SE REQUIERE */}
            {showApprovalFields('manager2') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Manager2" person={headerDraft.manager2} />
                  <div className={styles.grid3}>
                    <DateInput label="Manager2 Date" k="manager2Date" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Manager2 Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.manager2Status || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Manager2</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('manager2')}
                        defaultSelectedUsers={
                          headerDraft.manager2 ? [headerDraft.manager2.secondaryText || headerDraft.manager2.loginName || headerDraft.manager2.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Manager2 Date" k="manager2Date" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Manager2 Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.manager2Status || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('manager2') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('manager2')}
                            >
                              Agree as Manager2
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('manager2')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* Director - SOLO SI SE REQUIERE */}
            {showApprovalFields('director') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Director" person={headerDraft.director} />
                  <div className={styles.grid3}>
                    <DateInput label="Director Date" k="directorDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Director Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.directorStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Director</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('director')}
                        defaultSelectedUsers={
                          headerDraft.director ? [headerDraft.director.secondaryText || headerDraft.director.loginName || headerDraft.director.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Director Date" k="directorDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Director Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.directorStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('director') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('director')}
                            >
                              Agree as Director
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('director')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* VP - SOLO SI SE REQUIERE */}
            {showApprovalFields('vp') && (
              readOnlyMode ? (
                <>
                  <PersonView label="VP" person={headerDraft.vp} />
                  <div className={styles.grid3}>
                    <DateInput label="VP Date" k="vpDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>VP Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.vpStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>VP</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('vp')}
                        defaultSelectedUsers={
                          headerDraft.vp ? [headerDraft.vp.secondaryText || headerDraft.vp.loginName || headerDraft.vp.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="VP Date" k="vpDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>VP Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.vpStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('vp') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('vp')}
                            >
                              Agree as VP
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('vp')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* CFO - SOLO SI SE REQUIERE */}
            {showApprovalFields('cfo') && (
              readOnlyMode ? (
                <>
                  <PersonView label="CFO" person={headerDraft.cfo} />
                  <div className={styles.grid3}>
                    <DateInput label="CFO Date" k="cfoDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>CFO Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.cfoStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>CFO</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('cfo')}
                        defaultSelectedUsers={
                          headerDraft.cfo ? [headerDraft.cfo.secondaryText || headerDraft.cfo.loginName || headerDraft.cfo.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="CFO Date" k="cfoDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>CFO Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.cfoStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('cfo') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('cfo')}
                            >
                              Agree as CFO
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('cfo')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* CEO - SOLO SI SE REQUIERE */}
            {showApprovalFields('ceo') && (
              readOnlyMode ? (
                <>
                  <PersonView label="CEO" person={headerDraft.ceo} />
                  <div className={styles.grid3}>
                    <DateInput label="CEO Date" k="ceoDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>CEO Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.ceoStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>CEO</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('ceo')}
                        defaultSelectedUsers={
                          headerDraft.ceo ? [headerDraft.ceo.secondaryText || headerDraft.ceo.loginName || headerDraft.ceo.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="CEO Date" k="ceoDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>CEO Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.ceoStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('ceo') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('ceo')}
                            >
                              Agree as CEO
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('ceo')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* Procurement - SOLO SI SE REQUIERE (NUEVO) */}
            {showApprovalFields('procurement') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Procurement" person={headerDraft.procurement} />
                  <div className={styles.grid3}>
                    <DateInput label="Procurement Date" k="procurementDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Procurement Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.procurementStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Procurement</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('procurement')}
                        defaultSelectedUsers={
                          headerDraft.procurement ? [headerDraft.procurement.secondaryText || headerDraft.procurement.loginName || headerDraft.procurement.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Procurement Date" k="procurementDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Procurement Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.procurementStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('procurement') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('procurement')}
                            >
                              Agree as Procurement
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('procurement')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            {/* Finance (Final) - SOLO SI SE REQUIERE (NUEVO) */}
            {showApprovalFields('finance') && (
              readOnlyMode ? (
                <>
                  <PersonView label="Finance (Final)" person={headerDraft.finance} />
                  <div className={styles.grid3}>
                    <DateInput label="Finance Date" k="financeDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Finance Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.financeStatus || 'Pending'}
                      </div>
                    </div>
                    <div className={styles.fieldGroup} />
                  </div>
                </>
              ) : (
                <>
                  <div className={styles.fieldGroup}>
                    <label className={styles.fieldLabel}>Finance (Final)</label>
                    <div className={styles.peoplePicker}>
                      <PeoplePicker
                        context={peopleCtx}
                        personSelectionLimit={1}
                        ensureUser={true}
                        showtooltip={true}
                        principalTypes={[PrincipalType.User]}
                        onChange={onPicker('finance')}
                        defaultSelectedUsers={
                          headerDraft.finance ? [headerDraft.finance.secondaryText || headerDraft.finance.loginName || headerDraft.finance.text].filter(Boolean) as string[] : []
                        }
                      />
                    </div>
                  </div>
                  <div className={styles.grid3}>
                    <DateInput label="Finance Date" k="financeDate" />
                    <div className={styles.fieldGroup}>
                      <label className={styles.fieldLabel}>Finance Status</label>
                      <div className={styles.readonlyBox}>
                        {headerDraft.financeStatus || 'Pending'}
                      </div>
                    </div>
                    {readOnlyMode === 'tosign' ? (
                      <div className={styles.fieldGroup}>
                        {enableApprover('finance') && (
                          <>
                            <button
                              type="button"
                              className={styles.saveBtn}
                              onClick={() => handleApprove('finance')}
                            >
                              Agree as Finance
                            </button>
                            <button
                              type="button"
                              className={styles.resetBtn}
                              style={{ marginLeft: 8 }}
                              onClick={() => handleDisagree('finance')}
                            >
                              Disagree
                            </button>
                          </>
                        )}
                      </div>
                    ) : (
                      <div className={styles.fieldGroup} />
                    )}
                  </div>
                </>
              )
            )}

            <div className={styles.grid3}>
              <div className={styles.fieldGroup}>
                <label className={styles.fieldLabel}>PO Number (assigned by Procurement)</label>
                <input
                  id="fld-poNumber"
                  className={styles.input}
                  value={headerDraft.poNumber || ''}
                  onChange={keepFocus('fld-poNumber', onHeaderText('poNumber'))}
                  placeholder="PO-####-####"
                  disabled={readOnlyMode !== null}
                />
              </div>
              <div className={styles.fieldGroup} />
              <div className={styles.fieldGroup} />
            </div>
          </Section>

          {/* Footer actions */}
          {readOnlyMode === null && (
            <div className={styles.footerRow}>
              <button
                type="reset"
                className={styles.resetBtn}
                onClick={() => { resetFormState(); showSnack('Form cleared', 'info'); }}
              >
                Clear
              </button>

              {!editingItemId && (
                <button
                  type="submit"
                  className={`${styles.saveBtn} ${submitLoading ? styles.btnLoading : ''}`}
                  disabled={submitLoading}
                >
                  {submitLoading ? 'Saving...' : 'Save / Submit'}
                </button>
              )}
              {editingItemId && !isLocked && (
                <button
                  type="submit"
                  className={`${styles.saveBtn} ${submitLoading ? styles.btnLoading : ''}`}
                  disabled={submitLoading}
                >
                  {submitLoading ? 'Updating...' : `Update PR #${editingItemId}`}
                </button>
              )}
            </div>
          )}
        </fieldset>
      </>
    );
  };

  // Estados para controlar si las listas están expandidas o colapsadas
  const [expandedLists, setExpandedLists] = React.useState<{
    mysent: boolean;
    tosign: boolean;
    approved: boolean;
  }>({
    mysent: false,
    tosign: false,
    approved: false
  });

  const toggleList = (list: keyof typeof expandedLists) => {
    setExpandedLists(prev => ({
      ...prev,
      [list]: !prev[list]
    }));
  };

  // Función para obtener el nombre legible del rol
  const getRoleDisplayName = (role: RoleKey): string => {
    const roleNames: Record<RoleKey, string> = {
      supervisor: 'Supervisor',
      staffManager: 'Staff Manager',
      staffManager2: 'Staff2 Manager',
      manager: 'Manager',
      manager2: 'Manager2',
      director: 'Director',
      vp: 'VP',
      cfo: 'CFO',
      ceo: 'CEO',
      procurement: 'Procurement',
      finance: 'Finance',
      requester: 'Requester'
    };
    return roleNames[role] || role;
  };

  /* =================== Render =================== */
  return (
    <div className={styles.poRoot}>
      <div className={styles.wrapper}>
        <header className={styles.formHeader}>
          <div className={styles.titleRow}>
            <h1 className={styles.h1}>Purchase Requisition</h1>
          </div>
          

          <div className={styles.viewTabs}>
            <button
              type="button"
              className={`${styles.tabBtn} ${activeView === 'new' ? styles.tabBtnActive : ''}`}
              onClick={() => switchView('new')}
            >
              1) New PR
            </button>
            <button
              type="button"
              className={`${styles.tabBtn} ${activeView === 'mysent' ? styles.tabBtnActive : ''}`}
              onClick={() => switchView('mysent')}
            >
              2) My PRs
            </button>
            <button
              type="button"
              className={`${styles.tabBtn} ${activeView === 'tosign' ? styles.tabBtnActive : ''}`}
              onClick={() => switchView('tosign')}
            >
              3) To Approve
            </button>
            <button
              type="button"
              className={`${styles.tabBtn} ${activeView === 'approved' ? styles.tabBtnActive : ''}`}
              onClick={() => switchView('approved')}
            >
              4) Approved / PDF
            </button>
          </div>
        </header>

        {/* NEW / EDIT */}
        {activeView === 'new' && (
          <form autoComplete="off" onSubmit={(e) => editingItemId ? updateParentAndLines(e) : createParentAndLines(e)}>
            {renderForm(null)}
          </form>
        )}

        {/* MY PRs */}
        {activeView === 'mysent' && (
          <form autoComplete="off" onSubmit={(e) => editingItemId ? updateParentAndLines(e) : createParentAndLines(e)}>
            <Section title="My PRs (sent)">
              <div className={styles.miniHint}>
                {listLoading ? 'Loading...' : `${mySentWithStatus.length} item(s) with pending approvals`}
              </div>
              
              {/* Botón para expandir/colapsar la lista */}
              <button
                type="button"
                className={styles.expandBtn}
                onClick={() => toggleList('mysent')}
              >
                {expandedLists.mysent ? '▲ Collapse' : '▼ Expand'} List ({mySentWithStatus.length} items)
              </button>
              
              {expandedLists.mysent && mySentWithStatus.map(({ item, pendingRoles, approvedRoles, disagreeRoles }) => (
                <div key={item.Id} className={styles.listRow}>
                  <div style={{ flex: 1 }}>
                    <div className={styles.bold}>
                      {item.Title}
                      <span style={{ marginLeft: '8px', fontSize: '12px', color: '#666' }}>
                        (Pending: {pendingRoles.length} roles)
                      </span>
                    </div>
                    <div className={styles.miniHint}>ID: {item.Id} • {item.Created}</div>
                    
                    {/* Mostrar estado de aprobación */}
                    <div style={{ marginTop: '8px' }}>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
                        {pendingRoles.map(role => (
                          <span
                            key={role}
                            style={{
                              backgroundColor: '#fff3cd',
                              color: '#856404',
                              padding: '2px 8px',
                              borderRadius: '12px',
                              fontSize: '12px',
                              border: '1px solid #ffeaa7'
                            }}
                          >
                            Pending: {getRoleDisplayName(role)}
                          </span>
                        ))}
                        {approvedRoles.map(role => (
                          <span
                            key={role}
                            style={{
                              backgroundColor: '#d4edda',
                              color: '#155724',
                              padding: '2px 8px',
                              borderRadius: '12px',
                              fontSize: '12px',
                              border: '1px solid #c3e6cb'
                            }}
                          >
                            Approved: {getRoleDisplayName(role)}
                          </span>
                        ))}
                      </div>
                    </div>
                  </div>
                  <div className={styles.rowBtns}>
                    <button
                      type="button"
                      className={styles.saveBtn}
                      onClick={async () => {
                        await loadItemIntoForm(item.Id);
                        showSnack(`Loaded PR #${item.Id} for view/edit`, 'info');
                      }}
                    >
                      Load & Edit
                    </button>
                  </div>
                </div>
              ))}
            </Section>

            {editingItemId && (
              <Section title={`Editing PR #${editingItemId}`}>
                {renderForm(null)}
              </Section>
            )}
          </form>
        )}

        {/* APPROVED / PDF */}
        {activeView === 'approved' && (
          <form autoComplete="off" onSubmit={(e) => e.preventDefault()}>
            <Section title="Approved PRs (with approval date)">
              <div className={styles.miniHint}>{listLoading ? 'Loading...' : `${myApproved.length} item(s)`}</div>
              
              {/* Botón para expandir/colapsar la lista */}
              <button
                type="button"
                className={styles.expandBtn}
                onClick={() => toggleList('approved')}
              >
                {expandedLists.approved ? '▲ Collapse' : '▼ Expand'} List ({myApproved.length} items)
              </button>
              
              {expandedLists.approved && myApproved.map(it => (
                <div key={it.Id} className={styles.listRow}>
                  <div>
                    <div className={styles.bold}>{it.Title}</div>
                    <div className={styles.miniHint}>ID: {it.Id} • {it.Created}</div>
                  </div>
                  <div className={styles.rowBtns}>
                    <button
                      type="button"
                      className={styles.saveBtn}
                      onClick={async () => {
                        await loadItemIntoForm(it.Id);
                        showSnack(`Loaded PR #${it.Id}`, 'info');
                      }}
                    >
                      Load & View
                    </button>
                    <button
                      type="button"
                      className={`${styles.resetBtn} ${pdfLoading ? styles.btnLoading : ''}`}
                      onClick={() => handleGeneratePdf(it)}
                      disabled={pdfLoading}
                    >
                      {pdfLoading ? 'Generating...' : 'Generate PDF'}
                    </button>
                  </div>
                </div>
              ))}
            </Section>

            {editingItemId && (
              <Section title={`PR #${editingItemId} (Read-Only)`}>
                {renderForm('mysent')}
              </Section>
            )}
          </form>
        )}

        {/* TO APPROVE */}
        {activeView === 'tosign' && (
          <form autoComplete="off" onSubmit={(e) => e.preventDefault()}>
            <Section title="Requests to Approve">
              <div className={styles.miniHint}>{listLoading ? 'Loading...' : `${myToSign.length} document(s)`}</div>
              
              {/* Botón para expandir/colapsar la lista */}
              <button
                type="button"
                className={styles.expandBtn}
                onClick={() => toggleList('tosign')}
              >
                {expandedLists.tosign ? '▲ Collapse' : '▼ Expand'} List ({myToSign.length} items)
              </button>
              
              {expandedLists.tosign && myToSign.map(({ item, roles }) => (
                <div key={item.Id} className={styles.listRow}>
                  <div>
                    <div className={styles.bold}>{item.Title} <span className={styles.miniHint}>ID: {item.Id}</span></div>
                    <div className={styles.miniHint}>You must approve as: {roles.map(r => getRoleDisplayName(r)).join(', ')}</div>
                  </div>
                  <div className={styles.rowBtns}>
                    <button
                      type="button"
                      className={styles.saveBtn}
                      onClick={async () => {
                        await loadItemIntoForm(item.Id, { signRoles: roles });
                        showSnack(`Loaded PR #${item.Id} to approve as ${roles.map(r => getRoleDisplayName(r)).join(', ')}`, 'info');
                      }}
                    >
                      Load to approve
                    </button>
                  </div>
                </div>
              ))}
            </Section>

            {editingItemId && signFormRoles && signFormRoles.length > 0 && (
              <Section title={`Approve PR #${editingItemId}`}>
                {(signFormRoles || []).map(r => {
                  const status =
                    r === 'supervisor' ? headerDraft.supervisorStatus :
                      r === 'staffManager' ? headerDraft.staffManagerStatus :
                        r === 'manager' ? headerDraft.managerStatus :
                          r === 'director' ? headerDraft.directorStatus :
                            r === 'vp' ? headerDraft.vpStatus :
                              r === 'cfo' ? headerDraft.cfoStatus :
                                r === 'ceo' ? headerDraft.ceoStatus :
                                  r === 'procurement' ? headerDraft.procurementStatus :
                                    r === 'finance' ? headerDraft.financeStatus : undefined;

                  const statusText = status || 'Pending';

                  return (
                    <div key={r} className={styles.grid3} style={{ alignItems: 'end' }}>
                      <div className={styles.fieldGroup}>
                        <div className={styles.fieldLabel}>Role</div>
                        <div className={styles.readonlyBox}>{getRoleDisplayName(r)}</div>
                      </div>
                      <div className={styles.fieldGroup}>
                        <button
                          type="button"
                          className={styles.saveBtn}
                          onClick={() => handleApprove(r)}
                        >
                          Agree as {getRoleDisplayName(r)}
                        </button>
                        <button
                          type="button"
                          className={styles.resetBtn}
                          onClick={() => handleDisagree(r)}
                          style={{ marginLeft: 8 }}
                        >
                          Disagree
                        </button>
                      </div>
                      <div className={styles.fieldGroup}>
                        <div className={styles.miniHint}>
                          Status: {statusText}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </Section>
            )}

            {editingItemId && (
              <Section title={`PR #${editingItemId} (Read-Only)`}>
                {renderForm('tosign')}
              </Section>
            )}
          </form>
        )}

        {/* Snackbar */}
        <div
          style={{
            position: 'fixed',
            top: 12,
            left: 0,
            right: 0,
            display: 'flex',
            justifyContent: 'center',
            zIndex: 99999,
            pointerEvents: 'none'
          }}
          aria-live="polite"
        >
          {snack.open && (
            <div
              role="status"
              style={{
                pointerEvents: 'auto',
                padding: '10px 14px',
                borderRadius: 8,
                boxShadow: '0 8px 24px rgba(0,0,0,.15)',
                fontSize: 13,
                fontWeight: 600,
                color: snack.variant === 'error' ? '#fff' : '#0a0a0a',
                background:
                  snack.variant === 'success' ? '#d1fadf' :
                    snack.variant === 'error' ? '#ef4444' :
                      '#e5e7eb'
              }}
            >
              {snack.msg}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
export default PoRequestForm;
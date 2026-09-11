import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import "../styles/common.css";

declare const _spPageContextInfo: any;

/* =====================================================================
   MASTER DATA PAGE (multi-table)
   ---------------------------------------------------------------------
   Merges one row per project across SharePoint tables:
     • Project List          – identity + Status / Contractual & Revised
                               Delivery Date / PR+ Quality / PR+ Submission
                               / PR+ Approval / NWC
     • Collection List       – NormalDue, NormalOverdue, RetentionDue,
                               RetentionOverdue, GstTds, WHT, ReceivableValue
     • Revenue Forecast List – OBL (Revenue)
     • InvoicingPlanning     – OBL (Balance to Invoice)
     • SEIBIES               – OBL (BIES/SEI)
     • Inventory             – OBL (Inventory)
   Every cell shows the SharePoint table the value came from.
   Import = project-first: missing Project Reference is CREATED on
   Project List; existing references UPDATE each table.
   Every child table links via Project lookup (Title) + ProjectReference
   lookup; writes set <Field>Id from the Project List item.
   ===================================================================== */

/* ------------------------------------------------------------------ */
/* Types                                                              */
/* ------------------------------------------------------------------ */

interface CollectionData {
  normalDue: number;
  normalOverdue: number;
  retentionDue: number;
  retentionOverdue: number;
  gstTds: number;
  wht: number;
  receivableValue: number;
}

interface MasterRow {
  id: number;
  projectReference: string;
  projectName: string;
  profitCentre: string;
  customerName: string;
  groupHead: string;
  projectManager: string;
  enggLead: string;
  engSubLead1: string;
  engSubLead2: string;
  status: string;
  contractualDeliveryDate: string;
  revisedContractualDeliveryDate: string;
  prQuality: string;
  prSubmissionStatus: string;
  prApprovalStatus: string;
  nwc: number;
  collection: CollectionData;
  revenueForecast: number;
  invoicingPlanning: number;
  seibies: number;
  inventory: number;
  collectionRowId?: number | null;
  revenueRowId?: number | null;
  invoicingRowId?: number | null;
  seibiesRowId?: number | null;
  inventoryRowId?: number | null;
}

interface SharePointField {
  internalName: string;
  title: string;
  type: string;
}

/** one importable/exportable logical column */
interface ColDef {
  /** logical key used across this module */
  key: string;
  /** header shown in table / template */
  label: string;
  /** every spelling seen in source Excel (normalized) */
  aliases: string[];
  /** SharePoint list that stores this column */
  source: string;
  kind: 'text' | 'number' | 'date';
}

/* ------------------------------------------------------------------ */
/* List names + source tags                                           */
/* ------------------------------------------------------------------ */

const PROJECT_LIST = "Project List";
const COLLECTION_LIST = "Collection List";
const REVENUE_LIST = "Revenue Forecast List";
const INVOICING_LIST = "InvoicingPlanning";
const SEIBIES_LIST = "SEIBIES";
const INVENTORY_LIST = "Inventory";

const CHILD_LISTS = [
  COLLECTION_LIST,
  REVENUE_LIST,
  INVOICING_LIST,
  SEIBIES_LIST,
  INVENTORY_LIST
];

/** auto-create missing rows on all child tables once after page load */
const AUTO_SYNC_ON_LOAD = true;

const siteUrl: string =
  typeof _spPageContextInfo !== "undefined"
    ? _spPageContextInfo.webAbsoluteUrl
    : window.location.origin + "/sites/DigitalTransformationTest";

/* ------------------------------------------------------------------ */
/* Small helpers                                                      */
/* ------------------------------------------------------------------ */

const norm = (s: string): string => s.trim().toLowerCase().replace(/[\s,％%()/]+/g, '');

const pad2 = (n: number): string => String(n).padStart(2, "0");

const inrNum = (n: number): string => Math.round(Number(n) || 0).toLocaleString("en-IN");

/** parse excel-ish numbers: " 18,647 ", " (46,305) ", " - ", "1,03,622 " */
const parseExcelNumber = (raw: any): number | null => {
  if (raw === undefined || raw === null) return null;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
  let s = String(raw).trim();
  if (s === "" || s === "-" || /^[-–—]$/.test(s)) return null;
  let negative = false;
  if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
  if (s.startsWith("-")) { negative = true; s = s.slice(1); }
  const cleaned = s.replace(/[^0-9.]/g, "");
  if (!cleaned) return null;
  const n = parseFloat(cleaned);
  if (!Number.isFinite(n)) return null;
  return Math.round((negative ? -n : n) * 10000) / 10000;
};

/** excel-ish value → number incl. percentages: "93%"→93, "(46,305)"→-46305, "1,03,622"→103622 */
const parseNumText = (v: any): number | null => {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  const s = String(v ?? "").trim().replace(/％/g, "%");
  if (!s) return null;
  return parseExcelNumber(s.endsWith("%") ? s.slice(0, -1) : s);
};

/** show a PR+ quality value; always ends with % when present */
const formatPrVal = (v: any): string => {
  if (v === undefined || v === null || String(v).trim() === "") return "";
  const s = String(v).trim();
  return /%$/.test(s) ? s : s + "%";
};

/** "10-07-2027" or "07/10/2027" → ISO "2027-10-07T00:00:00Z" (assume dd-mm-yyyy) */
const parseExcelDate = (raw: any): string => {
  if (raw === undefined || raw === null) return "";
  if (raw instanceof Date) return raw.toISOString();
  let s = String(raw).trim();
  if (!s) return "";
  const m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (m) {
    let dd = parseInt(m[1], 10), mo = parseInt(m[2], 10), yy = parseInt(m[3], 10);
    if (yy < 100) yy += 2000;
    return `${yy}-${pad2(mo)}-${pad2(dd)}T00:00:00Z`;
  }
  const iso = new Date(s);
  return Number.isNaN(iso.getTime()) ? "" : iso.toISOString();
};

const displayDate = (raw: string): string => {
  if (!raw) return "-";
  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return raw;
  return `${pad2(d.getDate())}-${pad2(d.getMonth() + 1)}-${d.getFullYear()}`;
};

/** display/ISO date → value for <input type="date"> (YYYY-MM-DD) */
const toDateInput = (raw: string): string => {
  if (!raw) return "";
  const m = String(raw).match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) return `${m[3]}-${pad2(parseInt(m[2], 10))}-${pad2(parseInt(m[1], 10))}`;
  const d = new Date(raw);
  return Number.isNaN(d.getTime()) ? "" : `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
};

const personName = (v: any): string => {
  if (!v) return "";
  if (typeof v === "object" && v !== null) return v.Title || "";
  return String(v);
};

const esc = (s: string): string =>
  String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/* ------------------------------------------------------------------ */
/* SharePoint REST helpers                                            */
/* ------------------------------------------------------------------ */

/* every REST call is bounded — a stuck request can never pin the loader forever */
/* netFetch = the REAL fetch, kept under a private name so the timeout wrapper can never wrap itself */
const netFetch = (url: string, init?: RequestInit): Promise<Response> => fetch(url, init);

const fetchWithTimeout = async (url: string, options: RequestInit = {}, ms = 45000): Promise<Response> => {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), ms);
  try {
    return await netFetch(url, { ...options, signal: ctrl.signal });
  } finally {
    clearTimeout(timer);
  }
};

/* session caches — avoid re-discovering the same schema on every call */
const listFieldsCache = new Map<string, SharePointField[]>();
const listExistsCache = new Set<string>();
const entityTypeCache = new Map<string, string>();

const getDigest = async (): Promise<string> => {
  try {
    if (typeof _spPageContextInfo !== "undefined" && _spPageContextInfo.formDigestValue) {
      return _spPageContextInfo.formDigestValue;
    }
    const res = await fetchWithTimeout(`${siteUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose" }
    });
    const data = await res.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  } catch (error) {
    console.error("Error getting digest:", error);
    throw error;
  }
};

const checkListExists = async (listName: string): Promise<boolean> => {
  if (listExistsCache.has(listName)) return true;
  try {
    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    if (res.ok) listExistsCache.add(listName);
    return res.ok;
  } catch (error) {
    return false;
  }
};

const getListFields = async (listName: string): Promise<SharePointField[]> => {
  const cached = listFieldsCache.get(listName);
  if (cached) return cached;
  try {
    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/fields` +
      `?$select=InternalName,Title,TypeAsString,Hidden&$filter=Hidden eq false`,
      { headers: { Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose" } }
    );
    if (!res.ok) return [];
    const data = await res.json();
    const fields = (data.d.results || []).filter((f: any) =>
      !f.InternalName.startsWith("_") &&
      f.TypeAsString !== "Attachments" &&
      f.TypeAsString !== "Computed" &&
      !String(f.InternalName).includes("ColorTag") &&
      f.InternalName !== "ID" &&
      f.InternalName !== "ContentType"
    ).map((f: any) => ({ internalName: f.InternalName, title: f.Title, type: f.TypeAsString }));
    if (fields.length > 0) listFieldsCache.set(listName, fields);
    return fields;
  } catch (error) {
    console.error(`Error getting fields for ${listName}:`, error);
    return [];
  }
};

const findField = (fields: SharePointField[], desired: string): SharePointField | null => {
  const dn = desired.trim().toLowerCase();
  const exact = fields.find(f =>
    f.internalName.toLowerCase() === dn || f.title.toLowerCase() === dn);
  if (exact) return exact;
  const clean = dn.replace(/[^a-z0-9]/g, "");
  const cleanMatch = fields.find(f =>
    f.internalName.replace(/[^a-z0-9]/g, "").toLowerCase() === clean ||
    f.title.replace(/[^a-z0-9]/g, "").toLowerCase() === clean);
  if (cleanMatch) return cleanMatch;
  const partial = fields.find(f =>
    f.internalName.toLowerCase().includes(dn) ||
    f.title.toLowerCase().includes(dn));
  return partial || null;
};

interface QuerySpec { select: string[]; expand: string[]; }

/** build $select/$expand for the fields we want (lookup & user → fetch Id + expand)
 *  Note: when $expand is used, $select MUST also name the expanded field's target
 *  sub-fields (e.g. `Project/Title`); selecting only `ProjectId` makes SPO answer
 *  400 "The query to field 'Project' is not valid…". */
const buildQuery = (fields: SharePointField[], desired: string[]): QuerySpec => {
  const select: string[] = ["Id", "Title"];
  const expand: string[] = [];
  desired.forEach(d => {
    const f = findField(fields, d);
    if (!f) return;
    if (f.type === "Lookup" || f.type === "User") {
      select.push(`${f.internalName}Id`, `${f.internalName}/Id`, `${f.internalName}/Title`);
      if (!expand.includes(f.internalName)) expand.push(f.internalName);
    } else if (!select.includes(f.internalName)) {
      select.push(f.internalName);
    }
  });
  return { select: Array.from(new Set(select)), expand: Array.from(new Set(expand)) };
};

const getItems = async (listName: string, desired: string[] = [], top = 5000): Promise<any[]> => {
  const fetchItems = async (query: string): Promise<any[] | null> => {
    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?${query}`,
      { headers: { Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose" } }
    );
    if (res.ok) {
      const data = await res.json();
      return data.d.results || [];
    }
    console.error(`Failed to fetch ${listName} (HTTP ${res.status}):`, await res.text());
    return null;
  };
  try {
    if (!(await checkListExists(listName))) {
      console.warn(`List does not exist: ${listName}`);
      return [];
    }
    const fields = await getListFields(listName);
    if (fields.length === 0) return [];

    if (desired.length === 0) {
      const flat = await fetchItems(`$select=${encodeURIComponent(fields.map(f => f.internalName).join(","))}&$top=${top}`);
      return flat || [];
    }
    const q = buildQuery(fields, desired);
    const primary =
      `$select=${encodeURIComponent(q.select.join(","))}` +
      (q.expand.length ? `&$expand=${encodeURIComponent(q.expand.join(","))}` : "") +
      `&$top=${top}`;
    const rows = await fetchItems(primary);
    if (rows) return rows;
    if (q.expand.length) {
      /* SPO may still reject $expand → retry scalar-only (lookups come back as <Name>Id) */
      console.warn(`Retrying ${listName} without $expand…`);
      const scalar = ["Id", "Title"];
      fields.forEach(f => {
        if (f.type === "Lookup" || f.type === "User") scalar.push(`${f.internalName}Id`);
        else if (!scalar.includes(f.internalName)) scalar.push(f.internalName);
      });
      const fallback = await fetchItems(`$select=${encodeURIComponent(Array.from(new Set(scalar)).join(","))}&$top=${top}`);
      return fallback || [];
    }
    return [];
  } catch (error) {
    console.error(`Error fetching ${listName}:`, error);
    return [];
  }
};

/* person-column id → display name (used when a list was fetched without $expand) */
let cachedUserMap: Record<number, string> = {};
const getUserNames = async (): Promise<Record<number, string>> => {
  if (Object.keys(cachedUserMap).length) return cachedUserMap;
  try {
    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/siteusers?$select=Id,Title&$top=5000`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    if (res.ok) {
      const d = await res.json();
      (((d.d && d.d.results) || []) as any[]).forEach((u: any) => {
        if (u && u.Id) cachedUserMap[Number(u.Id)] = String(u.Title ?? "");
      });
    }
  } catch (e) {
    console.error("Error loading site users:", e);
  }
  return cachedUserMap;
};

/* person display-name (lower-cased) → SharePoint user id, used to write
   GH / PM / Engg Lead / Engg sub leads when importing an Excel that stores names */
let userTitleCache: Record<string, number> = {};
const getUserIdsByTitle = async (): Promise<Record<string, number>> => {
  if (Object.keys(userTitleCache).length) return userTitleCache;
  try {
    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/siteusers?$select=Id,Title&$top=5000`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    if (res.ok) {
      const d = await res.json();
      ((((d.d && d.d.results) || []) as any[])).forEach((u: any) => {
        if (u && u.Id && u.Title) {
          const k = String(u.Title).trim().toLowerCase();
          if (k && userTitleCache[k] === undefined) userTitleCache[k] = Number(u.Id);
        }
      });
    }
  } catch (e) {
    console.error("Error loading site users (by title):", e);
  }
  return userTitleCache;
};

const getListEntityType = async (listName: string): Promise<string> => {
  const cached = entityTypeCache.get(listName);
  if (cached) return cached;
  try {
    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')?$select=ListItemEntityTypeFullName`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    const data = await res.json();
    const t = String(data.d.ListItemEntityTypeFullName);
    entityTypeCache.set(listName, t);
    return t;
  } catch (error) {
    console.error("Error getting entity type:", listName, error);
    return `SP.Data.${listName.replace(/[^a-zA-Z0-9]/g, "_x0020_")}ListItem`;
  }
};

/** generic row finder: by Project lookup id → ProjectReference lookup id → Title */
const findChildRow = async (
  listName: string,
  projectId: number,
  projectRef: string,
  projectTitle: string
): Promise<{ row: any | null; projectFieldId: string | null; refFieldId: string | null; refFieldText: string | null }> => {
  const fields = await getListFields(listName);
  const projField = findField(fields, "Project");
  const refField = findField(fields, "ProjectReference");
  const projectFieldId = projField && (projField.type === "Lookup" || projField.type === "User")
    ? `${projField.internalName}Id` : null;
  const refFieldId = refField && refField.type === "Lookup" ? `${refField.internalName}Id` : null;
  const refFieldText = refField && refField.type !== "Lookup" ? refField.internalName : null;

  const rows = await getItems(listName, ["Project", "ProjectReference"], 5000);
  const byProj = rows.find(r => projectFieldId && Number(r[projectFieldId]) === projectId);
  if (byProj) return { row: byProj, projectFieldId, refFieldId, refFieldText };
  const byRef = rows.find(r => refFieldId && Number(r[refFieldId]) === projectId);
  if (byRef) return { row: byRef, projectFieldId, refFieldId, refFieldText };
  const titleCombined = `${projectRef} - ${projectTitle}`;
  const byTitle = rows.find(r => r.Title === titleCombined || r.Title === projectRef || r.Title === projectTitle);
  return { row: byTitle || null, projectFieldId, refFieldId, refFieldText };
};

/* ------------------------------------------------------------------ */
/* Loading overlay + toasts (same as original page)                   */
/* ------------------------------------------------------------------ */

const createLoadingOverlay = (message = "Processing..."): HTMLElement => {
  removeLoadingOverlay(); // never stack two loaders
  const overlay = document.createElement("div");
  overlay.className = "loading-overlay";
  overlay.innerHTML = `
    <div class="loading-container">
      <div class="loading-spinner"></div>
      <div class="loading-text">${message}</div>
      <div class="loading-progress"></div>
    </div>`;
  if (!document.querySelector("#loading-styles")) {
    const style = document.createElement("style");
    style.id = "loading-styles";
    style.textContent = `
      .loading-overlay{position:absolute;top:0;left:0;width:100%;height:100%;background:rgba(255,255,255,0.94);display:flex;justify-content:center;align-items:center;z-index:99999;backdrop-filter:blur(3px);border-radius:12px;overflow:hidden;box-sizing:border-box}
      .loading-container{text-align:center;padding:40px;background:white;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,0.15);min-width:300px;border:1px solid #e8e8e8}
      .loading-spinner{width:50px;height:50px;border:4px solid #f3f3f3;border-top:4px solid #0078d4;border-radius:50%;animation:spin 1s linear infinite;margin:0 auto 20px}
      .loading-text{font-size:16px;color:#333;margin-bottom:15px;font-weight:500}
      .loading-progress{height:4px;background:#f0f0f0;border-radius:2px;overflow:hidden;position:relative}
      .loading-progress::after{content:'';position:absolute;top:0;left:-100%;width:100%;height:100%;background:linear-gradient(90deg,#0078d4 0%,#00a2ff 100%);animation:progress 1.5s ease-in-out infinite}
      @keyframes spin{0%{transform:rotate(0)}100%{transform:rotate(360deg)}}
      @keyframes progress{0%{left:-100%}100%{left:100%}}`;
    document.head.appendChild(style);
  }
  return overlay;
};

const updateLoadingMessage = (overlay: HTMLElement, message: string): void => {
  const el = overlay.querySelector(".loading-text");
  if (el) el.textContent = message;
};

const removeLoadingOverlay = (): void => {
  document.querySelectorAll(".loading-overlay").forEach(el => el.remove());
};

/** mount the loader inside the visible table/edit card — never the whole page */
const loaderHost = (): HTMLElement => {
  const cards = Array.from(document.querySelectorAll<HTMLElement>(".table-section, .edit-section"));
  const visible = cards.find(c => c.offsetParent !== null);
  return visible || (document.querySelector(".table-section") as HTMLElement | null) || document.body;
};

const showToast = (message: string, type: "success" | "error" | "info" | "warning" = "info", duration = 4000): void => {
  if (!document.querySelector("#toast-styles")) {
    const style = document.createElement("style");
    style.id = "toast-styles";
    style.textContent = `
      .custom-toast{position:fixed;top:20px;right:20px;padding:14px 18px;background:white;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,0.15);display:flex;align-items:center;gap:12px;z-index:10000;min-width:280px;max-width:420px;border-left:4px solid;transform:translateX(120%);opacity:0;transition:transform .3s,opacity .3s}
      .custom-toast.show{transform:translateX(0);opacity:1}
      .custom-toast.hide{transform:translateX(120%);opacity:0}
      .toast-success{border-left-color:#28a745;background:linear-gradient(135deg,#d4edda 0,#fff 100%)}
      .toast-error{border-left-color:#dc3545;background:linear-gradient(135deg,#f8d7da 0,#fff 100%)}
      .toast-info{border-left-color:#17a2b8;background:linear-gradient(135deg,#d1ecf1 0,#fff 100%)}
      .toast-warning{border-left-color:#ffc107;background:linear-gradient(135deg,#fff3cd 0,#fff 100%)}
      .toast-icon{font-size:20px}.toast-content{flex:1;font-size:14px;color:#333}
      .toast-close{background:none;border:none;font-size:18px;cursor:pointer;color:#999;width:24px;height:24px;border-radius:50%}
      .toast-close:hover{background:rgba(0,0,0,.1)}`;
    document.head.appendChild(style);
  }
  const toast = document.createElement("div");
  toast.className = `custom-toast toast-${type}`;
  toast.innerHTML = `<div class="toast-icon">${type === "success" ? "✅" : type === "error" ? "❌" : type === "warning" ? "⚠️" : "ℹ️"}</div><div class="toast-content">${message}</div><button class="toast-close">&times;</button>`;
  document.body.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add("show"));
  const dismiss = (): void => {
    toast.classList.add("hide");
    setTimeout(() => toast.remove(), 300);
  };
  toast.querySelector(".toast-close")?.addEventListener("click", dismiss);
  setTimeout(dismiss, duration);
};

/* ------------------------------------------------------------------ */
/* Excel column registry (logical key → aliases seen in real files)   */
/* ------------------------------------------------------------------ */

const PROJECT_TEXT_COLS: ColDef[] = [
  { key: "projectReference", label: "Project Reference", source: PROJECT_LIST, kind: "text",
    aliases: ["projectreference", "projectdefinition(reference)", "projectdefinitionreference", "project_x0020_definition_x0028_reference_x0029_"] },
  { key: "projectName", label: "Project Name", source: PROJECT_LIST, kind: "text",
    aliases: ["projectname", "description(projectname)", "descriptionprojectname", "description", "description(projectname)"] },
  { key: "profitCentre", label: "Profit Centre", source: PROJECT_LIST, kind: "text",
    aliases: ["profitcentre", "profitcenter", "pc", "profit_x0020_centre"] },
  { key: "customerName", label: "Customer Name", source: PROJECT_LIST, kind: "text",
    aliases: ["customername", "customer", "customer_x0020_name"] },
  { key: "groupHead", label: "GH", source: PROJECT_LIST, kind: "text", aliases: ["gh", "grouphead", "grouphead1"] },
  { key: "projectManager", label: "PM", source: PROJECT_LIST, kind: "text", aliases: ["pm", "projectmanager", "projectmanager1"] },
  { key: "enggLead", label: "Engg Lead", source: PROJECT_LIST, kind: "text", aliases: ["engglead", "englead", "engleads", "engg_lead"] },
  { key: "engSubLead1", label: "Engg sub lead 1", source: PROJECT_LIST, kind: "text", aliases: ["enggsublead1", "engsublead1", "enggsublead", "enggsublead_1"] },
  { key: "engSubLead2", label: "Engg sub lead 2", source: PROJECT_LIST, kind: "text", aliases: ["enggsublead2", "engsublead2", "enggsublead_2"] },
  { key: "status", label: "Status", source: PROJECT_LIST, kind: "text", aliases: ["status"] },
  { key: "contractualDeliveryDate", label: "Contractual Delivery Date", source: PROJECT_LIST, kind: "date",
    aliases: ["contractualdeliverydate", "contractualdeliverydate(approvedbycustomer)", "contractualdeliverydateapprovedbycustomer"] },
  { key: "revisedContractualDeliveryDate", label: "Revised Contractual Delivery Date", source: PROJECT_LIST, kind: "date",
    aliases: ["revisedcontractualdeliverydate", "revisedcontractualdeliverydate(approvedbycustomer)", "revisedcontractualdeliverydateapprovedbycustomer"] },
  { key: "prQuality", label: "PR+ Quality %", source: PROJECT_LIST, kind: "text",
    aliases: ["prquality", "prquality%", "pr+quality%", "prqualitypercent"] },
  { key: "prSubmissionStatus", label: "PR+ Submission Status", source: PROJECT_LIST, kind: "text",
    aliases: ["prsubmissionstatus", "pr+submissionstatus"] },
  { key: "prApprovalStatus", label: "PR+ Approval Status", source: PROJECT_LIST, kind: "text",
    aliases: ["prapprovalstatus", "pr+approvalstatus"] },
  { key: "nwc", label: "NWC", source: PROJECT_LIST, kind: "number", aliases: ["nwc", "nwc"] },
];

const COLLECTION_COLS: ColDef[] = [
  { key: "normalDue", label: "NormalDue", source: COLLECTION_LIST, kind: "number", aliases: ["normaldue", "normalreceivables", "normalreceivable"] },
  { key: "normalOverdue", label: "NormalOverdue", source: COLLECTION_LIST, kind: "number", aliases: ["normaloverdue", "normaloverduereceivables", "normaloverduereceivable"] },
  { key: "retentionDue", label: "RetentionDue", source: COLLECTION_LIST, kind: "number", aliases: ["retentiondue", "retentionreceivables", "retentionreceivable"] },
  { key: "retentionOverdue", label: "RetentionOverdue", source: COLLECTION_LIST, kind: "number", aliases: ["retentionoverdue", "retentionoverduereceivables", "retentionoverduereceivable"] },
  { key: "gstTds", label: "GstTds", source: COLLECTION_LIST, kind: "number", aliases: ["gsttds", "gsttds", "gst", "gsttdsx"] },
  { key: "wht", label: "IT TDS / Others (WHT)", source: COLLECTION_LIST, kind: "number", aliases: ["wht", "ittds", "ittdsothers", "ittds/others", "ittdsother"] },
  { key: "receivableValue", label: "Receivables (incl TDS)", source: COLLECTION_LIST, kind: "number",
    aliases: ["receivablevalue", "receivables(incltds)", "receivablesincltds", "receivables", "receivable"] },
  { key: "obl", label: "OBL", source: REVENUE_LIST, kind: "number",
    aliases: ["obl", "revenueforecastobl", "revenueobl", "oblrevenue"] },
  { key: "invoicingPlanning", label: "Balance to Invoice", source: INVOICING_LIST, kind: "number",
    aliases: ["balancetoinvoice", "balancebilling", "invoicingplanningobl", "invoicingplanning"] },
  { key: "seibies", label: "BIES/SEI", source: SEIBIES_LIST, kind: "number",
    aliases: ["bies/sei", "biessei", "seibies", "seibiesobl", "bies", "sei"] },
  { key: "inventory", label: "Inventory", source: INVENTORY_LIST, kind: "number",
    aliases: ["inventory", "inventoryobl", "inventoryactual"] }
];

const ALL_COLS: ColDef[] = [...PROJECT_TEXT_COLS, ...COLLECTION_COLS];

/** final template columns: EXACT header row of “master data excel.csv” (single header row, no band) */
const CSV_TEMPLATE: Array<{ key: string; h: string }> = [
  { key: "profitCentre", h: "Profit Centre" },
  { key: "projectReference", h: "Project definition(Reference)" },
  { key: "customerName", h: "Customer Name" },
  { key: "projectName", h: "Description (Project Name)" },
  { key: "groupHead", h: "GH" },
  { key: "projectManager", h: "PM" },
  { key: "enggLead", h: "Engg Lead" },
  { key: "engSubLead1", h: "Engg sub lead 1" },
  { key: "engSubLead2", h: "Engg sub lead 2" },
  { key: "status", h: "Status" },
  { key: "contractualDeliveryDate", h: "Contractual Delivery Date" },
  { key: "revisedContractualDeliveryDate", h: "Revised Contractual Delivery Date (approved by customer)" },
  { key: "normalDue", h: " Normal Receivables " },
  { key: "normalOverdue", h: " Normal Overdue Receivables " },
  { key: "retentionDue", h: " Retention Receivables " },
  { key: "retentionOverdue", h: " Retention Overdue Receivables " },
  { key: "gstTds", h: " GST TDS " },
  { key: "wht", h: " IT TDS/Others " },
  { key: "receivableValue", h: " Receivables (incl TDS) " },
  { key: "obl", h: " OBL " },
  { key: "invoicingPlanning", h: " Balance Billing " },
  { key: "seibies", h: "BIES/SEI" },
  { key: "nwc", h: "NWC" },
  { key: "inventory", h: "Inventory" },
  { key: "prQuality", h: "PR+ Quality %" },
  { key: "prSubmissionStatus", h: "PR+ Submission Status" },
  { key: "prApprovalStatus", h: "PR+ Approval Status" },
];

const findColByHeader = (header: string): ColDef | null => {
  const n = norm(header);
  if (!n) return null;
  /* aliases are compared raw and normalized, so headers like "PR+ Quality %" match */
  return ALL_COLS.find(c => c.aliases.some(a => a === n || norm(a) === n)) || null;
};

/* =====================================================================
   UI CONTENT + STYLE
   ===================================================================== */

const pageContent = `
<div class="project-page fade-in master-page">
  <section class="table-section card master-dashboard">
    <div class="page-header">
      <div class="header-actions">
        <button class="btn-primary import-master-btn">📤 Import Master Excel</button>
        <button class="btn-secondary sync-all-btn" title="Create any missing row on every source table for each project">🔄 Sync All Tables</button>
        <button class="btn-secondary download-template-btn">📥 Download Template</button>
      </div>
    </div>

    <div class="table-actions">
      <input type="text" class="search-input" placeholder="🔍 Search by Project Reference / Name / Customer / Status..." />
      <button class="btn-secondary reset-search-btn">Reset Filter</button>
    </div>

    <div style="overflow-x:auto;">
      <table class="modern-table master-table">
        <thead>
          <tr class="master-col-headers">
            <th rowspan="2">#</th>
            <th rowspan="2">Project Reference</th>
            <th rowspan="2">Project Name</th>
            <th rowspan="2">Customer / GH / PM</th>
            <th rowspan="2">Status &amp; Delivery</th>
            <th rowspan="2">Receivables</th>
            <th rowspan="2">OBL</th>
            <th rowspan="2">Balance to Invoice</th>
            <th rowspan="2">BIES/SEI</th>
            <th rowspan="2">NWC</th>
            <th rowspan="2">Inventory</th>
            <th rowspan="2">PR+ Quality / Sub / Appr</th>
            <th rowspan="2">Actions</th>
          </tr>
        </thead>
        <tbody class="project-table-body"></tbody>
      </table>
    </div>

    <div class="pagination-container">
      <div class="rows-per-page">
        <label for="rowsPerPageSelect">Rows per page:</label>
        <select id="rowsPerPageSelect" class="rows-per-page-select">
          <option value="10" >10</option>
          <option value="20"selected>25</option>
          <option value="50">50</option>
          <option value="100">100</option>
          <option value="0">All</option>
        </select>
      </div>
      <div class="pagination"></div>
    </div>

  </section>

  <section class="edit-section card" style="display:none;">
    <div class="form-header"><h2>✏️ Edit Master Data</h2></div>
    <form id="masterEditForm">
      <div class="form-section-header">Project Details</div>
      <div class="form-row">
        <div class="form-group"><label>Project Reference</label><input type="text" id="editProjectReference" readonly class="readonly-field" /></div>
        <div class="form-group"><label>Project Name</label><input type="text" id="editProjectName" /></div>
        <div class="form-group"><label>Profit Centre</label><input type="text" id="editProfitCentre" /></div>
        <div class="form-group"><label>Customer Name</label><input type="text" id="editCustomerName" /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Status</label><input type="text" id="editStatus" /></div>
        <div class="form-group"><label>Contractual Delivery Date</label><input type="date" id="editContractualDeliveryDate" /></div>
        <div class="form-group"><label>Revised Contractual Delivery Date</label><input type="date" id="editRevisedContractualDeliveryDate" /></div>
        <div class="form-group"><label>NWC</label><input type="number" id="editNWC" step="0.0001" /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>PR+ Quality %</label><input type="text" id="editPrQuality" /></div>
        <div class="form-group"><label>PR+ Submission Status</label><input type="text" id="editPrSubmissionStatus" /></div>
        <div class="form-group"><label>PR+ Approval Status</label><input type="text" id="editPrApprovalStatus" /></div>
      </div>

      <div class="form-section-header">Collection Data</div>
      <div class="form-row">
        <div class="form-group"><label>NormalDue</label><input type="number" id="editNormalDue" step="0.0001" /></div>
        <div class="form-group"><label>NormalOverdue</label><input type="number" id="editNormalOverdue" step="0.0001" /></div>
        <div class="form-group"><label>RetentionDue</label><input type="number" id="editRetentionDue" step="0.0001" /></div>
        <div class="form-group"><label>RetentionOverdue</label><input type="number" id="editRetentionOverdue" step="0.0001" /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>GstTds</label><input type="number" id="editGstTds" step="0.0001" /></div>
        <div class="form-group"><label>IT TDS / Others (WHT)</label><input type="number" id="editWHT" step="0.0001" /></div>
        <div class="form-group"><label>Receivables (incl TDS)</label><input type="number" id="editReceivableValue" step="0.0001" /></div>
      </div>

      <div class="form-section-header">Other Tables</div>
      <div class="form-row">
        <div class="form-group"><label>OBL</label><input type="number" id="editRevenueForecast" step="0.0001" /></div>
        <div class="form-group"><label>Balance to Invoice</label><input type="number" id="editInvoicingPlanning" step="0.0001" /></div>
        <div class="form-group"><label>BIES/SEI</label><input type="number" id="editSEIBIES" step="0.0001" /></div>
        <div class="form-group"><label>Inventory</label><input type="number" id="editInventory" step="0.0001" /></div>
      </div>

      <div class="form-actions">
        <button type="submit" class="btn-primary">💾 Update Data</button>
        <button type="button" class="btn-secondary back-to-table">← Back to List</button>
      </div>
    </form>
  </section>

  <!-- ================= Excel import/export popup ================= -->
  <div class="import-popup-overlay" id="importPopupOverlay" style="display:none;">
    <div class="import-popup">
      <div class="popup-head">
        <h2>📥 Master Data Excel — Import / Export</h2>
        <button type="button" class="import-popup-close" title="Close">✕</button>
      </div>
      <div class="popup-body">
        <div class="popup-actions">
          <button type="button" class="btn-primary import-choose-btn">📂 1. Choose Excel File…</button>
          <input type="file" class="import-file-input" accept=".xlsx,.xls" style="display:none;" />
          <button type="button" class="btn-secondary import-template-btn">📥 2. Download Template (contains ALL current data)</button>
        </div>
        <div class="import-hint">
          💡 <strong>Template contains every project with current values</strong> — edit existing rows or add new rows
          (new Project Reference) in Excel, save, then choose that file again.<br />
          Existing <strong>Project Reference → updated</strong> on every table · New <strong>Project Reference → project
          created</strong>, then rows ensured on every source table. Rows without a Project Reference are ignored.
        </div>
        <div class="import-progress" style="display:none;">
          <span class="popup-spin"></span><span class="import-progress-text">Reading file…</span>
        </div>
        <div class="import-results" style="display:none;">
          <div class="import-results-head">📊 Import status — one line per project row</div>
          <div class="import-summary-grid">
            <div class="is-item ok"><span class="lbl">Updated</span><span class="val" id="impUpdated">0</span></div>
            <div class="is-item new"><span class="lbl">New Projects</span><span class="val" id="impCreated">0</span></div>
            <div class="is-item err"><span class="lbl">Errors</span><span class="val" id="impErrors">0</span></div>
            <div class="is-item tot"><span class="lbl">Total</span><span class="val" id="impTotal">0</span></div>
          </div>
          <div class="import-status-list"></div>
        </div>
      </div>
      <div class="popup-foot">
        <button type="button" class="btn-primary popup-done-btn">✔ Done</button>
      </div>
    </div>
  </div>
</div>
`;

const pageStyle = `
.project-page { font-family:'Segoe UI', Arial, sans-serif; color:#333; animation:fadeIn 0.4s ease-in; }

.card { position:relative; background:#fff; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.08); padding:20px; margin-bottom:20px; transition:all 0.3s ease; border:1px solid #e8e8e8; }
.master-dashboard { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); }

.page-header { display:flex; justify-content:flex-end; align-items:center; flex-wrap:wrap; gap:15px; margin-bottom:20px; padding-bottom:15px; border-bottom:2px solid #0078d4; }
.page-header h1 { color:#0056b3; margin:0; font-size:22px; font-weight:600; }
.header-actions { display:flex; gap:10px; align-items:center; flex-wrap:wrap; }

.btn-primary { background: linear-gradient(135deg,#0078d4 0%,#005a9e 100%); color:#fff; border:none; padding:10px 18px; border-radius:8px; cursor:pointer; font-size:13px; font-weight:500; display:inline-flex; align-items:center; gap:8px; transition:all 0.3s ease; box-shadow:0 2px 4px rgba(0,120,212,0.2); }
.btn-primary:hover { background: linear-gradient(135deg,#005a9e 0%,#004578 100%); transform:translateY(-1px); }
.btn-secondary { background:#6c757d; color:#fff; border:none; padding:10px 18px; border-radius:8px; cursor:pointer; font-size:13px; font-weight:500; display:inline-flex; align-items:center; gap:8px; transition:all 0.3s ease; }
.btn-secondary:hover { background:#5a6268; transform:translateY(-1px); }
.btn-secondary:disabled { opacity:0.55; cursor:not-allowed; transform:none; }

.master-table { width:100%; border-collapse:collapse; font-size:13px; background:#fff; }
.master-table th { background: linear-gradient(135deg,#f3f2f1 0%,#e1e1e1 100%); text-align:left; padding:10px 8px; font-weight:600; border:1px solid #d4d9e0; color:#0056b3; position:sticky; top:0; z-index:10; vertical-align:bottom; }
.master-table td { padding:8px; border:1px solid #e8e8e8; vertical-align:top; }
.master-table tbody tr:hover { background:#f0f6fd; }
.master-table tbody tr:nth-child(even) { background:#fafafa; }
.master-table tbody tr:nth-child(even):hover { background:#f0f6fd; }


.cell-value { font-weight:500; color:#1b2430; white-space:nowrap; }
.cell-value.num { text-align:right; font-variant-numeric:tabular-nums; font-weight:600; color:#0f2c59; display:block; }
.cell-value.neg { color:#c00000; }
.cell-value.pos { color:#0a8a3a; }
.cell-value.muted { color:#6c757d; font-size:12px; }

.badge { display:inline-block; padding:2px 7px; border-radius:4px; font-size:11px; font-weight:600; margin:1px 2px 1px 0; white-space:nowrap; }
.badge.pl   { background:#e7ecf5; color:#26374e; }
.badge.due  { background:#cfe2ff; color:#084298; }
.badge.over { background:#f8d7da; color:#842029; }
.badge.gst  { background:#f8f1d7; color:#842029; }
.badge.wht  { background:#a9d2b8; color:#0f5132; }
.badge.recv { background:#d1e7dd; color:#0f5132; }
.badge.dark { background:#0f2c59; color:#fff; }

.pill-status { display:inline-block; font-weight:700; font-size:11px; padding:1px 8px; border-radius:10px; }
.pill-ok  { background:#d1e7dd; color:#0f5132; }
.pill-warn{ background:#fff3cd; color:#7a5b00; }
.pill-bad { background:#f8d7da; color:#842029; }
.pill-so  { background:#e7ecf5; color:#26374e; }

.dl-line { white-space:nowrap; font-size:12px; color:#333; }
.pr-badge { font-size:11px; font-weight:700; padding:1px 7px; border-radius:9px; background:#eef2f7; color:#26374e; display:inline-block; margin:1px 2px 1px 0; }

.table-actions { display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; flex-wrap:wrap; gap:10px; }
.search-input { padding:9px 14px; border:2px solid #e1e1e1; border-radius:8px; flex:1; min-width:250px; font-size:13px; transition:all 0.3s ease; }
.search-input:focus { border-color:#0078d4; box-shadow:0 0 0 3px rgba(0,120,212,0.1); outline:none; }

.import-summary { background:#f8f9fa; border-top:3px solid #28a745; }
.import-summary h3 { margin-top:0; color:#28a745; }
.summary-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:12px; }
.summary-item { padding:12px; border-radius:8px; text-align:center; }
.summary-item.success { background:#d4edda; border:1px solid #c3e6cb; }
.summary-item.warning { background:#fff3cd; border:1px solid #ffeaa7; }
.summary-item.error { background:#f8d7da; border:1px solid #f5c6cb; }
.summary-item.info { background:#d1ecf1; border:1px solid #bee5eb; }
.summary-label { display:block; font-size:13px; color:#666; margin-bottom:4px; }
.summary-value { display:block; font-size:22px; font-weight:600; }

.edit-section { max-width:1000px; margin:0 auto; }
.form-section-header { font-size:15px; font-weight:600; color:#0078d4; margin:18px 0 12px 0; padding-bottom:6px; border-bottom:2px solid #e8e8e8; }
.form-row { display:flex; flex-wrap:wrap; gap:14px; margin-bottom:12px; }
.form-group { flex:1 1 calc(25% - 14px); display:flex; flex-direction:column; min-width:200px; }
.form-group label { font-weight:500; color:#495057; margin-bottom:6px; font-size:13px; }
.form-group input, .form-group select { width:100%; padding:9px; border:2px solid #e1e1e1; border-radius:6px; font-size:13px; transition:all 0.3s ease; box-sizing:border-box; }
.form-group input:focus { border-color:#0078d4; box-shadow:0 0 0 3px rgba(0,120,212,0.1); outline:none; }
.readonly-field { background:#f8f9fa; color:#6c757d; cursor:not-allowed; }
.form-group input[type="number"]::-webkit-inner-spin-button { opacity:1; }
.form-actions { text-align:center; margin-top:24px; padding-top:16px; border-top:1px solid #e8e8e8; }
.form-actions button { margin:0 6px; }

.actions-cell { display:flex; gap:6px; }
.edit-btn { width:34px; height:34px; border:none; border-radius:6px; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; padding:0; font-size:16px; background:linear-gradient(135deg,#ffc107 0%,#e0a800 100%); color:#fff; transition:all 0.3s ease; }
.edit-btn:hover { transform:translateY(-2px); }

.pagination-container { margin-top:16px; display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:nowrap; overflow-x:auto; }
.rows-per-page { display:inline-flex; align-items:center; gap:6px; font-size:13px; color:#495057; white-space:nowrap; }
.rows-per-page label { margin:0; }
.rows-per-page select { padding:5px 8px; border:1.5px solid #c9d2db; border-radius:6px; font-size:13px; background:#fff; cursor:pointer; }
.rows-per-page select:focus { border-color:#0078d4; outline:none; }
.pagination { display:flex; gap:6px; flex-wrap:nowrap; }
.pagination button { padding:6px 12px; border:2px solid #dee2e6; background:#fff; border-radius:6px; cursor:pointer; font-size:13px; transition:all 0.3s ease; }
.pagination button:hover { border-color:#0078d4; background:#f8f9fa; }
.pagination button.active { background:#0078d4; color:#fff; border-color:#0078d4; }

@keyframes fadeIn { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:translateY(0); } }


/* ---------------- Excel import/export popup ---------------- */
.import-popup-overlay { position:fixed; inset:0; background:rgba(15,23,42,.55); display:flex; align-items:center; justify-content:center; z-index:100000; padding:20px; backdrop-filter:blur(2px); }
.import-popup { width:min(860px,96vw); max-height:92vh; background:#fff; border-radius:14px; box-shadow:0 24px 64px rgba(0,0,0,.3); display:flex; flex-direction:column; overflow:hidden; animation:popIn .22s ease; }
@keyframes popIn { from { opacity:0; transform:translateY(14px) scale(.98); } to { opacity:1; transform:none; } }
.popup-head { display:flex; justify-content:space-between; align-items:center; padding:14px 18px; border-bottom:2px solid #0078d4; background:linear-gradient(135deg,#f3f9ff,#e8f1fb); flex:0 0 auto; }
.popup-head h2 { margin:0; font-size:17px; color:#0056b3; }
.import-popup-close { background:#fff; border:1px solid #d4d9e0; border-radius:50%; width:32px; height:32px; font-size:15px; line-height:1; cursor:pointer; color:#666; }
.import-popup-close:hover { background:#f8d7da; color:#842029; border-color:#f5c6cb; }
.popup-body { padding:16px 18px; overflow:auto; display:flex; flex-direction:column; gap:12px; }
.popup-actions { display:flex; gap:10px; flex-wrap:wrap; align-items:center; }
.import-hint { font-size:12.5px; color:#495057; background:#fff8e6; border:1px solid #ffe3a3; border-radius:8px; padding:9px 12px; line-height:1.55; }
.import-progress { display:flex; align-items:center; gap:10px; background:#eef2f7; border:1px solid #dde4ee; border-radius:8px; padding:10px 14px; }
.popup-spin { width:18px; height:18px; border:3px solid #dde4ee; border-top-color:#0078d4; border-radius:50%; animation:popSpin .8s linear infinite; flex:0 0 auto; }
@keyframes popSpin { to { transform:rotate(360deg); } }
.import-progress-text { font-size:13px; color:#26374e; }
.import-results { border:1px solid #dde4ee; border-radius:10px; overflow:hidden; }
.import-results-head { padding:9px 13px; background:#f8f9fa; border-bottom:1px solid #e4e9f0; font-weight:600; color:#26374e; font-size:13.5px; }
.import-summary-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; padding:11px 13px; background:#fff; border-bottom:1px solid #eef1f5; }
.is-item { text-align:center; padding:9px 6px; border-radius:8px; }
.is-item.ok  { background:#d4edda; } .is-item.new { background:#cfe2ff; } .is-item.err { background:#f8d7da; } .is-item.tot { background:#eef2f7; }
.is-item .lbl { display:block; font-size:10.5px; color:#495057; text-transform:uppercase; letter-spacing:.4px; margin-bottom:3px; }
.is-item .val { display:block; font-size:20px; font-weight:700; }
.is-item.ok .val { color:#0f5132; } .is-item.new .val { color:#084298; } .is-item.err .val { color:#842029; } .is-item.tot .val { color:#26374e; }
.import-status-list { max-height:34vh; overflow:auto; padding:6px 13px; background:#fcfdff; }
.is-row { display:flex; gap:10px; align-items:flex-start; padding:6px 3px; border-bottom:1px dashed #e6eaf0; }
.is-row:last-child { border-bottom:none; }
.is-chip { flex:0 0 auto; font-size:10px; font-weight:800; letter-spacing:.5px; padding:2px 8px; border-radius:9px; margin-top:2px; }
.is-chip.is-ok  { background:#d4edda; color:#0f5132; }
.is-chip.is-new { background:#cfe2ff; color:#084298; }
.is-chip.is-err { background:#f8d7da; color:#842029; }
.is-txt { flex:1; min-width:0; }
.is-txt strong { font-size:12.5px; color:#1b2430; display:block; }
.is-msg { font-size:11.5px; color:#6c757d; line-height:1.45; margin-top:1px; word-break:break-word; }
.popup-foot { display:flex; justify-content:flex-end; gap:10px; padding:12px 18px; border-top:1px solid #e4e9f0; background:#fafbfc; flex:0 0 auto; }
.btn-primary:disabled, .btn-secondary:disabled { opacity:.5; cursor:not-allowed; transform:none; }

@media (max-width: 900px) {
  .page-header { flex-direction:column; align-items:stretch; }
  .header-actions { flex-direction:column; }
  .summary-grid { grid-template-columns:1fr; }
  .form-group { flex:1 1 100%; }
  .master-table { font-size:12px; }
}
`;

/* =====================================================================
   PAGE LOGIC
   ===================================================================== */

export default {
  content: pageContent,
  style: pageStyle,
  script: (): void => {
    setTimeout((): void => {
      const tableBody = document.querySelector(".project-table-body") as HTMLTableSectionElement | null;
      const importBtn = document.querySelector(".import-master-btn") as HTMLButtonElement | null;
      const syncAllBtn = document.querySelector(".sync-all-btn") as HTMLButtonElement | null;
      const searchInput = document.querySelector(".search-input") as HTMLInputElement | null;
      const resetSearchBtn = document.querySelector(".reset-search-btn") as HTMLButtonElement | null;
      const editSection = document.querySelector(".edit-section") as HTMLElement | null;
      const tableSection = document.querySelector(".table-section") as HTMLElement | null;
      const editForm = document.querySelector("#masterEditForm") as HTMLFormElement | null;
      const backBtn = document.querySelector(".back-to-table") as HTMLButtonElement | null;
      const downloadTemplateBtn = document.querySelector(".download-template-btn") as HTMLButtonElement | null;
      const paginationDiv = document.querySelector(".pagination") as HTMLElement | null;

      const popupOverlay = document.querySelector(".import-popup-overlay") as HTMLElement | null;
      const popupResults = document.querySelector(".import-results") as HTMLElement | null;
      const popupStatusList = document.querySelector(".import-status-list") as HTMLElement | null;
      const popupProgress = document.querySelector(".import-progress") as HTMLElement | null;
      const popupProgressText = document.querySelector(".import-progress-text") as HTMLElement | null;
      const importFileEl = document.querySelector(".import-file-input") as HTMLInputElement | null;
      const importChooseBtn = document.querySelector(".import-choose-btn") as HTMLButtonElement | null;
      const popupTemplateBtn = document.querySelector(".import-template-btn") as HTMLButtonElement | null;
      const popupCloseBtn = document.querySelector(".import-popup-close") as HTMLButtonElement | null;
      const popupDoneBtn = document.querySelector(".popup-done-btn") as HTMLButtonElement | null;
      const impUpdatedEl = document.getElementById("impUpdated") as HTMLElement | null;
      const impCreatedEl = document.getElementById("impCreated") as HTMLElement | null;
      const impErrorsEl = document.getElementById("impErrors") as HTMLElement | null;
      const impTotalEl = document.getElementById("impTotal") as HTMLElement | null;

      let masterData: MasterRow[] = [];
      let currentEditing: MasterRow | null = null;
      let currentPage = 1;
      let pageSize = 10;
      const rowsPerPageEl = document.getElementById("rowsPerPageSelect") as HTMLSelectElement | null;
      let renderedData: MasterRow[] = [];
      let autoSyncDone = false;

      /* ---------- loading / fetch master (loads + joins ALL tables) ---------- */

      let loadBusy = false;
      const loadMaster = async (): Promise<void> => {
        if (loadBusy) return;
        loadBusy = true;
        const loading = createLoadingOverlay("Loading master data...");
        loaderHost().appendChild(loading);
        try {
          updateLoadingMessage(loading, "Loading data from all 6 tables...");
          const TABLE_LOAD: Array<[string, string[]]> = [
            [PROJECT_LIST, ["Title", "ProjectReference", "ProfitCentre", "CustomerName", "GroupHead1", "ProjectManager", "EnggLead", "EnggSubLead1", "EnggSubLead2", "Status", "ContractualDeliveryDate", "RevisedContractualDeliveryDate", "PRQuality", "PRSubmissionStatus", "PRApprovalStatus", "NWC"]],
            [COLLECTION_LIST, ["Project", "ProjectReference", "NormalDue", "NormalOverdue", "RetentionDue", "RetentionOverdue", "GstTds", "WHT", "ReceivableValue", "OBL"]],
            [REVENUE_LIST, ["Project", "ProjectReference", "OBL"]],
            [INVOICING_LIST, ["Project", "ProjectReference", "OBL"]],
            [SEIBIES_LIST, ["Project", "ProjectReference", "OBL"]],
            [INVENTORY_LIST, ["Project", "ProjectReference", "OBL"]]
          ];
          const doneFlags = new Array<boolean>(TABLE_LOAD.length).fill(false);
          const loaders: Array<Promise<any[]>> = TABLE_LOAD.map(([listName, desired], i) =>
            getItems(listName, desired, 5000)
              .catch(() => [] as any[])
              .then(rows => {
                doneFlags[i] = true;
                updateLoadingMessage(loading, `Loading tables... ${doneFlags.filter(Boolean).length}/${TABLE_LOAD.length} done`);
                return rows;
              })
          );
          const [projects, coll, rev, inv, sei, invt] = await Promise.all(loaders);

          /* fallback mode: $expand was skipped → person columns only carry Ids,
             so resolve GH / PM / EnggLead display names from site users once */
          const projectsNeedUsers = projects.some((p: any) =>
            (Number(p.GroupHead1Id) && !personName(p.GroupHead1)) ||
            (Number(p.ProjectManagerId) && !personName(p.ProjectManager)) ||
            (Number(p.EnggLeadId) && !personName(p.EnggLead)));
          if (projectsNeedUsers) {
            const umap = await getUserNames();
            projects.forEach((p: any) => {
              const patchPerson = (idKey: string, objKey: string): void => {
                if (!personName(p[objKey])) {
                  const nm = umap[Number(p[idKey])];
                  if (nm) p[objKey] = { Title: nm };
                }
              };
              patchPerson("GroupHead1Id", "GroupHead1");
              patchPerson("ProjectManagerId", "ProjectManager");
              patchPerson("EnggLeadId", "EnggLead");
            });
          }

          /* ---- helper: match a child row to this project ---- */
          const lookupText = (v: any): string => {
            if (!v) return "";
            return typeof v === "object" ? String(v.Title ?? v.Value ?? v.Name ?? "") : String(v);
          };
          const findIn = (rows: any[], ref: string, title: string, id: number): any => {
            const byProjId = rows.find(r => {
              const pid = r.ProjectId ?? (r.Project && typeof r.Project === "object" ? r.Project.Id : null);
              return pid !== undefined && Number(pid) === id;
            });
            if (byProjId) return byProjId;
            const byRefId = rows.find(r => r.ProjectReferenceId !== undefined && Number(r.ProjectReferenceId) === id);
            if (byRefId) return byRefId;
            return rows.find(r =>
              (r.Title && (r.Title === ref || r.Title === `${ref} - ${title}` || r.Title === title)) ||
              (r.ProjectReference !== undefined && r.ProjectReference !== null && lookupText(r.ProjectReference) === ref)
            ) || null;
          };

          const numVal = (row: any, candidates: string[]): number => {
            if (!row) return 0;
            for (const c of candidates) {
              let v = row[c];
              if (v === undefined || v === null || v === "") continue;
              if (typeof v === "object" && v !== null) v = v.Title ?? v.Value ?? v.Name ?? v;
              const n = parseExcelNumber(v);
              if (n !== null) return n;
            }
            return 0;
          };

          masterData = projects.map((p: any): MasterRow => {
            const ref = String(p.ProjectReference ?? p.Title ?? "").trim();
            const title = String(p.Title ?? "").trim();
            // ✅ If the Project Name is blank or only a "#"/"#N/A" placeholder, show the
            //    Project Reference as the Project Name instead.
            const projectName = (!title || /^#\s*(?:N\/A)?\s*$/i.test(title)) ? "-" : title;
            const collRow = findIn(coll, ref, title, p.Id);
            const revRow = findIn(rev, ref, title, p.Id);
            const invRow = findIn(inv, ref, title, p.Id);
            const seiRow = findIn(sei, ref, title, p.Id);
            const invtRow = findIn(invt, ref, title, p.Id);

            return {
              id: p.Id,
              projectReference: ref,
              projectName,
              profitCentre: String(p.ProfitCentre ?? p.PC ?? ""),
              customerName: String(p.CustomerName ?? ""),
              groupHead: personName(p.GroupHead1),
              projectManager: personName(p.ProjectManager),
              enggLead: personName(p.EnggLead),
              engSubLead1: personName(p.EnggSubLead1),
              engSubLead2: personName(p.EnggSubLead2),
              status: String(p.Status ?? ""),
              contractualDeliveryDate: p.ContractualDeliveryDate ? displayDate(String(p.ContractualDeliveryDate)) : "",
              revisedContractualDeliveryDate: p.RevisedContractualDeliveryDate ? displayDate(String(p.RevisedContractualDeliveryDate)) : "",
              prQuality: String(p.PRQuality ?? ""),
              prSubmissionStatus: String(p.PRSubmissionStatus ?? ""),
              prApprovalStatus: String(p.PRApprovalStatus ?? ""),
              nwc: numVal(p, ["NWC"]),
              collection: {
                normalDue: numVal(collRow, ["NormalDue"]),
                normalOverdue: numVal(collRow, ["NormalOverdue"]),
                retentionDue: numVal(collRow, ["RetentionDue"]),
                retentionOverdue: numVal(collRow, ["RetentionOverdue"]),
                gstTds: numVal(collRow, ["GstTds", "GSTTDS"]),
                wht: numVal(collRow, ["WHT"]),
                receivableValue: numVal(collRow, ["ReceivableValue", "Receivables", "OBL"])
              },
              revenueForecast: numVal(revRow, ["OBL"]),
              invoicingPlanning: numVal(invRow, ["OBL"]),
              seibies: numVal(seiRow, ["OBL"]),
              inventory: numVal(invtRow, ["OBL"]),
              collectionRowId: collRow ? collRow.Id : null,
              revenueRowId: revRow ? revRow.Id : null,
              invoicingRowId: invRow ? invRow.Id : null,
              seibiesRowId: seiRow ? seiRow.Id : null,
              inventoryRowId: invtRow ? invtRow.Id : null
            };
          });

          currentPage = 1;
          renderedData = [...masterData];
          renderTable();

          const countWithData = masterData.filter(r =>
            Object.values(r.collection).some(v => v !== 0) ||
            r.revenueForecast !== 0 || r.invoicingPlanning !== 0 || r.seibies !== 0 ||
            r.inventory !== 0 || r.nwc !== 0
          ).length;
          showToast(`Loaded ${masterData.length} projects (${countWithData} with data across tables)`, "success");
        } catch (error) {
          console.error("Error loading master data:", error);
          masterData = [];
          renderedData = [];
          renderTable();
          showToast("Error loading master data. Check console.", "error");
        } finally {
          removeLoadingOverlay();
          loadBusy = false;
        }
      };

      /* ---------- table rendering ---------- */

      const valCell = (val: string, cls = ""): string =>
        `<div class="cell-value ${cls}">${esc(val || "-")}</div>`;

      const numCell = (n: number): string => {
        const cls = n < 0 ? "num neg" : n > 0 ? "num pos" : "num muted";
        return `<div class="cell-value ${cls}">${n === 0 ? "-" : inrNum(n)}</div>`;
      };

      const statusPill = (s: string): string => {
        if (!s) return `<span class="cell-value muted">-</span>`;
        const up = String(s).toUpperCase();
        const cls = /CLSD|TECO|REL/.test(up) ? "pill-warn" : /SO/.test(up) ? "pill-so" : "pill-ok";
        return `<span class="pill-status ${cls}">${esc(s)}</span>`;
      };

      const renderTable = (): void => {
        if (!tableBody) return;
        if (renderedData.length === 0) {
          tableBody.innerHTML = `<tr><td colspan="13" style="text-align:center;padding:40px;color:#666;"><h3>No projects found</h3><p>Import an Excel file or check the Project List.</p></td></tr>`;
          return;
        }
        const totalPages = Math.max(1, Math.ceil(renderedData.length / pageSize));
        if (currentPage > totalPages) currentPage = totalPages;
        const start = (currentPage - 1) * pageSize;
        const page = renderedData.slice(start, start + pageSize);

        const rowsHtml = page.map((r, idx) => {
          const c = r.collection;
          return `<tr>
            <td>${start + idx + 1}</td>
            <td><strong>${esc(r.projectReference || "N/A")}</strong></td>
            <td>${valCell(r.projectName)}</td>
            <td>
              <div class="cell-value">${esc(r.customerName || "-")}</div>
              <div class="cell-value muted">GH: ${esc(r.groupHead || "-")} · PM: ${esc(r.projectManager || "-")}</div>
            </td>
            <td>
              <div>${statusPill(r.status)}</div>
              <div class="dl-line">C.D.D: ${esc(r.contractualDeliveryDate || "-")}</div>
              <div class="dl-line">R.D.D: ${esc(r.revisedContractualDeliveryDate || "-")}</div>
            </td>
            <td>
              <div><span class="badge due">N: ${inrNum(c.normalDue)}</span> <span class="badge over">NO: ${inrNum(c.normalOverdue)}</span></div>
              <div><span class="badge due">R: ${inrNum(c.retentionDue)}</span> <span class="badge over">RO: ${inrNum(c.retentionOverdue)}</span></div>
              <div><span class="badge gst">GST: ${inrNum(c.gstTds)}</span> <span class="badge wht">WHT: ${inrNum(c.wht)}</span> <span class="badge recv">RECV: ${inrNum(c.receivableValue)}</span></div>
            </td>
            <td>${numCell(r.revenueForecast)}</td>
            <td>${numCell(r.invoicingPlanning)}</td>
            <td>${numCell(r.seibies)}</td>
            <td>${numCell(r.nwc)}</td>
            <td>${numCell(r.inventory)}</td>
            <td>
              <div><span class="pr-badge">${esc(formatPrVal(r.prQuality) || "-")}</span></div>
              <div class="dl-line" style="font-size:11px;">Sub: ${esc(r.prSubmissionStatus || "-")}</div>
              <div class="dl-line" style="font-size:11px;">Appr: ${esc(r.prApprovalStatus || "-")}</div>
            </td>
            <td class="actions-cell">
              <button class="edit-btn" data-ref="${esc(r.projectReference)}" title="Edit">✏️</button>
            </td>
          </tr>`;
        }).join("");

        tableBody.innerHTML = rowsHtml;
        attachEditHandlers();
        renderPagination();
      };

      const renderPagination = (): void => {
        if (!paginationDiv) return;
        const totalPages = Math.max(1, Math.ceil(renderedData.length / pageSize));
        paginationDiv.innerHTML = "";
        if (totalPages <= 1 && renderedData.length <= pageSize) return;
        const max = 7;
        let s = Math.max(1, currentPage - Math.floor(max / 2));
        let e = Math.min(totalPages, s + max - 1);
        if (e - s + 1 < max) s = Math.max(1, e - max + 1);
        if (s > 1) { const d = document.createElement("button"); d.textContent = "…"; d.disabled = true; paginationDiv.appendChild(d); }
        for (let i = s; i <= e; i++) {
          const b = document.createElement("button");
          b.textContent = String(i);
          b.className = i === currentPage ? "active" : "";
          b.addEventListener("click", () => { currentPage = i; renderTable(); });
          paginationDiv.appendChild(b);
        }
        if (e < totalPages) { const d = document.createElement("button"); d.textContent = "…"; d.disabled = true; paginationDiv.appendChild(d); }
      };

      const attachEditHandlers = (): void => {
        document.querySelectorAll(".edit-btn").forEach(btn => {
          btn.addEventListener("click", () => {
            const ref = (btn as HTMLElement).dataset.ref;
            const row = masterData.find(m => m.projectReference === ref);
            if (!row) return;
            currentEditing = row;
            const set = (id: string, v: string): void => {
              const el = document.getElementById(id) as HTMLInputElement | null;
              if (el) el.value = v ?? "";
            };
            set("editProjectReference", row.projectReference);
            set("editProjectName", row.projectName);
            set("editProfitCentre", row.profitCentre);
            set("editCustomerName", row.customerName);
            set("editStatus", row.status);
            set("editContractualDeliveryDate", toDateInput(row.contractualDeliveryDate));
            set("editRevisedContractualDeliveryDate", toDateInput(row.revisedContractualDeliveryDate));
            set("editPrQuality", row.prQuality);
            set("editPrSubmissionStatus", row.prSubmissionStatus);
            set("editPrApprovalStatus", row.prApprovalStatus);
            set("editNWC", String(row.nwc));
            set("editNormalDue", String(row.collection.normalDue));
            set("editNormalOverdue", String(row.collection.normalOverdue));
            set("editRetentionDue", String(row.collection.retentionDue));
            set("editRetentionOverdue", String(row.collection.retentionOverdue));
            set("editGstTds", String(row.collection.gstTds));
            set("editWHT", String(row.collection.wht));
            set("editReceivableValue", String(row.collection.receivableValue));
            set("editRevenueForecast", String(row.revenueForecast));
            set("editInvoicingPlanning", String(row.invoicingPlanning));
            set("editSEIBIES", String(row.seibies));
            set("editInventory", String(row.inventory));
            if (editSection) editSection.style.display = "block";
            if (tableSection) tableSection.style.display = "none";
          });
        });
      };

      /* ---------- generic upsert (updates by row; creates when missing) ---------- */

      interface UpsertOpts {
        listName: string;
        projectId: number;
        projectRef: string;
        projectTitle: string;
        /** logical values: { desiredFieldName: number|string } written via schema-aware mapping */
        values: Record<string, number | string>;
        rowId?: number | null;
        /** true when the row is already known to be missing — skip the lookup round-trip */
        assumeNew?: boolean;
      }

      const upsertRow = async (o: UpsertOpts): Promise<{ success: boolean; id: number | null; isNew: boolean; error?: any }> => {
        try {
          const fields = await getListFields(o.listName);
          const digest = await getDigest();
          const entityType = await getListEntityType(o.listName);

          let itemId: number | null = o.rowId ?? null;
          let isNew = false;
          if (!itemId && !o.assumeNew) {
            const found = await findChildRow(o.listName, o.projectId, o.projectRef, o.projectTitle);
            itemId = found.row ? found.row.Id : null;
          }
          if (!itemId) isNew = true;

          const payload: any = { __metadata: { type: entityType } };

          const write = (field: SharePointField | null, value: number | string): void => {
            if (!field) return;
            const t = field.type;
            if (t === "Lookup" || t === "User") {
              payload[`${field.internalName}Id`] = String(value);
            } else if (t === "Number" || t === "Currency") {
              const n = parseNumText(value);
              payload[field.internalName] = n !== null ? n : 0;
            } else if (t === "Boolean") {
              payload[field.internalName] = Boolean(value);
            } else {
              payload[field.internalName] = String(value);
            }
          };

          const title = `${o.projectRef} - ${o.projectTitle}`;
          write(findField(fields, "Title"), title);

          // Project lookup → Project List item Id; ProjectReference → lookup Id / text
          const projField = findField(fields, "Project");
          if (projField && (projField.type === "Lookup" || projField.type === "User")) {
            payload[`${projField.internalName}Id`] = String(o.projectId);
          } else if (projField) {
            payload[projField.internalName] = o.projectTitle;
          }
          const refField = findField(fields, "ProjectReference");
          if (refField && (refField.type === "Lookup" || refField.type === "User")) {
            payload[`${refField.internalName}Id`] = String(o.projectId);
          } else if (refField) {
            payload[refField.internalName] = o.projectRef;
          }

          for (const [desired, value] of Object.entries(o.values)) {
            if (desired === "Title" || desired === "Project" || desired === "ProjectReference") continue;
            write(findField(fields, desired), value);
          }

          let res: Response;
          if (!isNew && itemId) {
            res = await fetchWithTimeout(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(o.listName)}')/items(${itemId})`, {
              method: "POST",
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest,
                "X-HTTP-Method": "MERGE",
                "IF-MATCH": "*"
              },
              body: JSON.stringify(payload)
            });
          } else {
            res = await fetchWithTimeout(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(o.listName)}')/items`, {
              method: "POST",
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest
              },
              body: JSON.stringify(payload)
            });
          }

          if (res.ok) {
            let created: number | null = itemId;
            if (isNew) {
              try { const j = await res.json(); created = j.d.Id; } catch (e) { /* ignore */ }
            }
            return { success: true, id: created, isNew };
          }
          const body = await res.text();
          console.error(`❌ Failed upsert in ${o.listName}:`, body);
          return { success: false, id: null, isNew, error: body };
        } catch (error) {
          console.error(`Error upserting in ${o.listName}:`, error);
          return { success: false, id: null, isNew: false, error };
        }
      };

      /** create a bare row for a project on a child table (no figures yet) */
      const ensureChildRow = async (listName: string, row: MasterRow): Promise<boolean> => {
        const result = await upsertRow({
          listName,
          projectId: row.id,
          projectRef: row.projectReference,
          projectTitle: row.projectName,
          values: {},
          rowId: undefined,
          assumeNew: true // sync snapshot already proved the row is missing
        });
        return result.success;
      };

      /** 🔄 SYNC ALL TABLES: one row per project on every source table,
       *  idempotent (never duplicates, never overwrites existing rows). */
      const syncAllTables = async (silent: boolean): Promise<number> => {
        let created = 0;
        for (const listName of CHILD_LISTS) {
          if (!(await checkListExists(listName))) {
            console.warn(`⏭️ ${listName} does not exist - skipped`);
            continue;
          }
          // server-side snapshot to never duplicate
          let serverRows: any[] = [];
          try {
            const res = await fetchWithTimeout(
              `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$select=Id,ProjectId,Title&$top=5000`,
              { headers: { Accept: "application/json;odata=verbose" } }
            );
            if (res.ok) serverRows = (await res.json()).d.results || [];
          } catch (e) { console.error("sync snapshot failed", e); }
          const hasId = (row: MasterRow): boolean => serverRows.some(s => Number(s.ProjectId) === row.id);
          for (const row of masterData) {
            if (hasId(row)) continue;
            if (await ensureChildRow(listName, row)) {
              created++;
              console.log(`➕ ${listName}: created ${row.projectReference} - ${row.projectName}`);
            }
          }
        }
        if (!silent) showToast(`✅ Sync complete: ${created} missing row(s) created across tables`, created ? "success" : "info");
        console.log(`🔄 Sync all tables: ${created} row(s) created`);
        return created;
      };

      /* ---------- PROJECT-FIRST EXCEL IMPORT ---------- */

      const parseSheet = (ws: any): Array<Record<string, any>> => {
        const aoa: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
        // find the header row: the row containing a "project reference"-like header
        let headerIdx = -1;
        for (let ri = 0; ri < aoa.length; ri++) {
          const row = aoa[ri];
          for (let ci = 0; ci < row.length; ci++) {
            if (norm(String(row[ci])) && /projectref|projectdefinition|projectreference/.test(norm(String(row[ci])))) {
              headerIdx = ri;
              break;
            }
          }
          if (headerIdx >= 0) break;
        }
        if (headerIdx < 0) return [];

        const header = aoa[headerIdx];
        const colByIndex: Array<ColDef | null> = header.map(h => findColByHeader(String(h ?? "")));

        const out: Array<Record<string, any>> = [];
        for (let ri = headerIdx + 1; ri < aoa.length; ri++) {
          const row = aoa[ri];
          if (!row || !row.some((c: any) => String(c ?? "").trim() !== "")) continue;
          const rec: Record<string, any> = {};
          colByIndex.forEach((def, ci) => {
            if (!def) return;
            const raw = row[ci];
            if (raw === undefined || String(raw ?? "").trim() === "") return;
            if (def.kind === "number") rec[def.key] = parseExcelNumber(raw);
            else if (def.kind === "date") rec[def.key] = parseExcelDate(raw);
            else rec[def.key] = String(raw).trim();
          });
          const ref = String(rec.projectReference ?? "").trim();
          if (ref) out.push(rec);
        }
        return out;
      };

      /* ==================================================================
         Excel import / export popup (project-first create-or-update)
         ================================================================== */

      let importing = false;
      let impUpdated = 0, impCreated = 0, impErrors = 0;

      const describeError = (e: any): string => {
        if (!e) return "Unknown error";
        if (typeof e === "string") {
          try {
            const j = JSON.parse(e);
            const v = j && j.error && j.error.message &&
              (typeof j.error.message === "string" ? j.error.message : j.error.message.value);
            if (v) return String(v);
          } catch (err) { /* not JSON */ }
          return e.length > 220 ? e.slice(0, 220) + "…" : e;
        }
        if (e instanceof Error) return e.message || "Error";
        if (typeof e.message === "string") return e.message;
        return String(e);
      };

      const bumpChips = (): void => {
        if (impUpdatedEl) impUpdatedEl.textContent = String(impUpdated);
        if (impCreatedEl) impCreatedEl.textContent = String(impCreated);
        if (impErrorsEl) impErrorsEl.textContent = String(impErrors);
      };

      const addStatusRow = (kind: "updated" | "created" | "error", ref: string, name: string, msg: string): void => {
        if (kind === "created") impCreated++;
        else if (kind === "updated") impUpdated++;
        else impErrors++;
        bumpChips();
        if (!popupStatusList) return;
        const div = document.createElement("div");
        div.className = "is-row";
        const chip = document.createElement("span");
        chip.className = "is-chip " + (kind === "created" ? "is-new" : kind === "error" ? "is-err" : "is-ok");
        chip.textContent = kind === "created" ? "NEW" : kind === "error" ? "ERROR" : "UPDATED";
        const txt = document.createElement("div");
        txt.className = "is-txt";
        const strong = document.createElement("strong");
        strong.textContent = ref + (name && name !== ref ? "  —  " + name : "");
        const small = document.createElement("div");
        small.className = "is-msg";
        small.textContent = msg;
        txt.appendChild(strong);
        txt.appendChild(small);
        div.appendChild(chip);
        div.appendChild(txt);
        popupStatusList.appendChild(div);
        popupStatusList.scrollTop = popupStatusList.scrollHeight;
      };

      const openImportPopup = (): void => {
        if (!popupOverlay || importing) return;
        impUpdated = impCreated = impErrors = 0;
        bumpChips();
        if (impTotalEl) impTotalEl.textContent = "0";
        if (popupStatusList) popupStatusList.innerHTML = "";
        if (popupProgress) popupProgress.style.display = "none";
        if (popupResults) popupResults.style.display = "none";
        popupOverlay.style.display = "flex";
      };

      const closeImportPopup = (): void => {
        if (!popupOverlay || importing) return;
        popupOverlay.style.display = "none";
      };

      const doImport = async (file: File): Promise<void> => {
        if (importing) { showToast("Import already running — please wait for it to finish", "warning"); return; }
        importing = true;
        const setUiBusy = (busy: boolean): void => {
          if (importChooseBtn) importChooseBtn.disabled = busy;
          if (popupTemplateBtn) popupTemplateBtn.disabled = busy;
          if (popupCloseBtn) popupCloseBtn.disabled = busy;
          if (popupDoneBtn) popupDoneBtn.disabled = busy;
        };
        setUiBusy(true);
        impUpdated = impCreated = impErrors = 0;
        bumpChips();
        try {
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(buf);
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = parseSheet(ws);
          if (!rows.length) throw new Error("No rows with a Project Reference found. Check the header row (Project definition(Reference)).");

          const total = rows.length;
          if (impTotalEl) impTotalEl.textContent = String(total);
          if (popupResults) popupResults.style.display = "none";
          if (popupStatusList) popupStatusList.innerHTML = "";
          if (popupProgress) popupProgress.style.display = "flex";

          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const ref = String(row.projectReference ?? "").trim();
            // ✅ Excel Project Name blank or just "#"/"#N/A" → the Project Reference
            //    becomes the Project Name (same rule as the master table view).
            const excelName = String(row.projectName ?? "").trim();
            if (!excelName || /^#\s*(?:N\/A)?\s*$/i.test(excelName)) row.projectName = ref;
            if (popupProgressText) popupProgressText.textContent = `Processing ${i + 1} of ${total} — ${ref}`;
            try {
              /* 1 ── project-first: find by ProjectReference; create if missing ── */
              let projRes = await fetchWithTimeout(
                `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items` +
                `?$select=Id,Title,ProjectReference&$filter=ProjectReference eq '${encodeURIComponent(ref)}'&$top=1`,
                { headers: { Accept: "application/json;odata=verbose" } }
              );
              let projData: any = null;
              if (projRes.ok) {
                const j = await projRes.json();
                if (j.d.results.length > 0) projData = j.d.results[0];
              }

              let projectId: number;
              let projectTitle: string;
              let isNew = false;
              if (projData) {
                projectId = projData.Id;
                // ✅ Keep a real Project List Title; when the stored Title is blank
                //    or "#" the Project Reference is used as the Project Name.
                const storedTitle = String(projData.Title ?? "").trim();
                projectTitle = (!storedTitle || /^#\s*(?:N\/A)?\s*$/i.test(storedTitle)) ? ref : storedTitle;
              } else {
                /* NEW project → create on Project List first */
                isNew = true;
                projectTitle = String(row.projectName ?? row.customerName ?? ref).trim() || ref;
                const entityType = await getListEntityType(PROJECT_LIST);
                const plPayload: any = { __metadata: { type: entityType } };
                const plFields = await getListFields(PROJECT_LIST);
                const writePL = (desired: string, value: any): void => {
                  const f = findField(plFields, desired);
                  if (!f) return;
                  if (f.type === "Lookup") return;
                  if (f.type === "User") return; // people are handled below by display-name resolution
                  if (f.type === "Number" || f.type === "Currency") {
                    const n = parseNumText(value);
                    if (n !== null) plPayload[f.internalName] = n;
                  } else if (value !== undefined && value !== null && String(value) !== "") {
                    plPayload[f.internalName] = String(value);
                  }
                };
                plPayload.Title = projectTitle;
                writePL("ProjectReference", ref);
                writePL("ProfitCentre", row.profitCentre);
                writePL("CustomerName", row.customerName);
                writePL("Status", row.status);
                writePL("ContractualDeliveryDate", row.contractualDeliveryDate || "");
                writePL("RevisedContractualDeliveryDate", row.revisedContractualDeliveryDate || "");
                writePL("PRQuality", row.prQuality);
                writePL("PRSubmissionStatus", row.prSubmissionStatus);
                writePL("PRApprovalStatus", row.prApprovalStatus);
                writePL("NWC", row.nwc);

                /* people columns (GH / PM / Engg Lead / Engg sub leads) — match by display name */
                {
                  const byName = await getUserIdsByTitle();
                  const peopleDesired: Array<[string, string]> = [
                    ["GroupHead1", "groupHead"], ["ProjectManager", "projectManager"],
                    ["EnggLead", "enggLead"], ["EnggSubLead1", "engSubLead1"], ["EnggSubLead2", "engSubLead2"]
                  ];
                  peopleDesired.forEach(([desired, key]) => {
                    const nm = String(row[key] ?? "").trim();
                    const f = findField(plFields, desired);
                    if (!f || f.type !== "User" || !nm) return;
                    const uid = byName[nm.toLowerCase()];
                    if (uid) plPayload[f.internalName + "Id"] = String(uid);
                    else console.warn(`Skipping ${desired} (${nm}) — not found in site users`);
                  });
                }

                const createRes = await fetchWithTimeout(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items`, {
                  method: "POST",
                  headers: {
                    Accept: "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": await getDigest()
                  },
                  body: JSON.stringify(plPayload)
                });
                if (!createRes.ok) {
                  const bodyText = await createRes.text();
                  console.error("Project List create failed:", bodyText);
                  throw new Error(`Could not create project on Project List: ${describeError(bodyText)}`);
                }
                try {
                  const j = await createRes.json();
                  projectId = j.d.Id;
                } catch (e) {
                  /* re-query by reference */
                  projRes = await fetchWithTimeout(
                    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items` +
                    `?$select=Id&$filter=ProjectReference eq '${encodeURIComponent(ref)}'&$top=1`,
                    { headers: { Accept: "application/json;odata=verbose" } }
                  );
                  const j = await projRes.json();
                  projectId = j.d.results[0].Id;
                }
              }

              const notes: string[] = [];

              /* existing project → update Project List info too */
              if (!isNew) {
                const update: Record<string, number | string> = {};
                const plFields = await getListFields(PROJECT_LIST);
                const put = (desired: string, value: any): void => {
                  const f = findField(plFields, desired);
                  if (!f || f.type === "Lookup" || f.type === "User") return;
                  if (f.type === "Number" || f.type === "Currency") {
                    const n = typeof value === "number" ? value : parseNumText(value);
                    if (n !== null) update[f.internalName] = n;
                  } else if (value !== undefined && String(value).trim() !== "") {
                    update[f.internalName] = String(value);
                  }
                };
                put("Status", row.status);
                put("ContractualDeliveryDate", row.contractualDeliveryDate || "");
                put("RevisedContractualDeliveryDate", row.revisedContractualDeliveryDate || "");
                put("PRQuality", row.prQuality);
                put("PRSubmissionStatus", row.prSubmissionStatus);
                put("PRApprovalStatus", row.prApprovalStatus);
                put("NWC", row.nwc);
                put("CustomerName", row.customerName);
                put("ProfitCentre", row.profitCentre);
                // ✅ Fix a stored "#"/blank Title: write the Project Reference as the name
                put("Title", String(projData?.Title ?? "").trim() !== projectTitle ? projectTitle : undefined);

                /* people columns — match by display name */
                {
                  const byName = await getUserIdsByTitle();
                  const peopleDesired: Array<[string, string]> = [
                    ["GroupHead1", "groupHead"], ["ProjectManager", "projectManager"],
                    ["EnggLead", "enggLead"], ["EnggSubLead1", "engSubLead1"], ["EnggSubLead2", "engSubLead2"]
                  ];
                  peopleDesired.forEach(([desired, key]) => {
                    const nm = String(row[key] ?? "").trim();
                    const f = findField(plFields, desired);
                    if (!f || f.type !== "User" || !nm) return;
                    const uid = byName[nm.toLowerCase()];
                    if (uid) update[f.internalName + "Id"] = String(uid);
                    else console.warn(`Skipping ${desired} (${nm}) — not found in site users`);
                  });
                }
                if (Object.keys(update).length) {
                  const f = await fetchWithTimeout(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items(${projectId})`, {
                    method: "POST",
                    headers: {
                      Accept: "application/json;odata=verbose",
                      "Content-Type": "application/json;odata=verbose",
                      "X-RequestDigest": await getDigest(),
                      "X-HTTP-Method": "MERGE",
                      "IF-MATCH": "*"
                    },
                    body: JSON.stringify({ __metadata: { type: await getListEntityType(PROJECT_LIST) }, ...update })
                  });
                  if (!f.ok) notes.push("⚠ Project List update failed: " + describeError(await f.text()));
                }
              }

              /* 2 ── Collection List ── */
              const collValues: Record<string, number | string> = {};
              const cMap: Array<[string, string]> = [
                ["NormalDue", "normalDue"], ["NormalOverdue", "normalOverdue"],
                ["RetentionDue", "retentionDue"], ["RetentionOverdue", "retentionOverdue"],
                ["GstTds", "gstTds"], ["WHT", "wht"], ["ReceivableValue", "receivableValue"]
              ];
              cMap.forEach(([fieldKey, excelKey]) => {
                const v = row[excelKey];
                if (typeof v === "number" && v !== 0) collValues[fieldKey] = v;
              });
              const collRes = await upsertRow({
                listName: COLLECTION_LIST, projectId, projectRef: ref, projectTitle,
                values: collValues, rowId: undefined
              });
              if (!collRes.success) notes.push("⚠ Collection List: " + describeError(collRes.error));

              /* 3 ── single-value tables (OBL convention) ── */
              const single: Array<[string, string]> = [
                [REVENUE_LIST, "obl"],
                [INVOICING_LIST, "invoicingPlanning"],
                [SEIBIES_LIST, "seibies"],
                [INVENTORY_LIST, "inventory"]
              ];
              for (const [listName, excelKey] of single) {
                const v = row[excelKey];
                const values: Record<string, number | string> = (typeof v === "number" && v !== 0) ? { OBL: v } : {};
                const ur = await upsertRow({
                  listName, projectId, projectRef: ref, projectTitle,
                  values, rowId: undefined
                });
                if (!ur.success) notes.push("⚠ " + listName + ": " + describeError(ur.error));
              }

              const baseMsg = isNew
                ? "New project created on Project List, then rows ensured on all source tables."
                : "Existing project — values updated on Project List + source tables.";
              addStatusRow(isNew ? "created" : "updated", ref, projectTitle,
                notes.length ? baseMsg + " " + notes.join(" ") : baseMsg);
            } catch (err) {
              console.error("Error processing row:", row, err);
              addStatusRow("error", ref, String(row.projectName ?? ""),
                describeError(err instanceof Error ? err.message : err));
            }
          }

          if (popupProgress) popupProgress.style.display = "none";
          if (popupResults) popupResults.style.display = "block";
          showToast(`✅ Import complete — ${impCreated} new project(s), ${impUpdated} updated, ${impErrors} error(s) of ${total}`,
            impErrors ? "warning" : "success", 6000);
          await loadMaster();
        } catch (error) {
          console.error("Import error:", error);
          if (popupProgress) popupProgress.style.display = "none";
          if (popupResults) popupResults.style.display = "block";
          if (popupStatusList) popupStatusList.innerHTML = "";
          addStatusRow("error", "(file)", file.name, describeError(error)); // bumps the chip
          showToast("❌ Import failed: " + describeError(error), "error", 6000);
        } finally {
          importing = false;
          setUiBusy(false);
        }
      };


      /* ---------- edit form save ---------- */

      const onEditSubmit = async (e: Event): Promise<void> => {
        e.preventDefault();
        if (!editForm || !currentEditing) return;
        const g = (id: string): HTMLInputElement => document.getElementById(id) as HTMLInputElement;
        const num = (id: string): number => {
          const n = parseExcelNumber(g(id).value);
          return n ?? 0;
        };
        const txt = (id: string): string => g(id).value.trim();

        const overlay = createLoadingOverlay("Updating project data...");
        loaderHost().appendChild(overlay);
        try {
          const row = currentEditing;
          const ref = txt("editProjectReference");
          const name = txt("editProjectName");
          const projectId = row.id;

          /* 1 ── Project List ── */
          const plFields = await getListFields(PROJECT_LIST);
          const pl: Record<string, number | string> = {};
          const put = (desired: string, value: any, allowEmpty = false): void => {
            const f = findField(plFields, desired);
            if (!f || f.type === "Lookup" || f.type === "User") return;
            if (f.type === "Number" || f.type === "Currency") {
              const n = typeof value === "number" ? value : parseNumText(value);
              if (n !== null) pl[f.internalName] = n;
            } else if (value !== undefined && (allowEmpty || String(value).trim() !== "")) {
              pl[f.internalName] = String(value);
            }
          };
          put("Title", name, true);
          put("ProjectReference", ref, true);
          put("ProfitCentre", txt("editProfitCentre"), true);
          put("CustomerName", txt("editCustomerName"), true);
          put("Status", txt("editStatus"), true);
          put("ContractualDeliveryDate", g("editContractualDeliveryDate").value || "", true);
          put("RevisedContractualDeliveryDate", g("editRevisedContractualDeliveryDate").value || "", true);
          put("PRQuality", txt("editPrQuality"), true);
          put("PRSubmissionStatus", txt("editPrSubmissionStatus"), true);
          put("PRApprovalStatus", txt("editPrApprovalStatus"), true);
          put("NWC", num("editNWC"));

          if (Object.keys(pl).length) {
            const r = await fetchWithTimeout(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items(${projectId})`, {
              method: "POST",
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": await getDigest(),
                "X-HTTP-Method": "MERGE",
                "IF-MATCH": "*"
              },
              body: JSON.stringify({ __metadata: { type: await getListEntityType(PROJECT_LIST) }, ...pl })
            });
            if (!r.ok) console.error("Project List update failed:", await r.text());
          }

          /* 2 ── Collection List ── */
          const collVals: Record<string, number | string> = {
            NormalDue: num("editNormalDue"),
            NormalOverdue: num("editNormalOverdue"),
            RetentionDue: num("editRetentionDue"),
            RetentionOverdue: num("editRetentionOverdue"),
            GstTds: num("editGstTds"),
            WHT: num("editWHT"),
            ReceivableValue: num("editReceivableValue")
          };
          await upsertRow({
            listName: COLLECTION_LIST, projectId, projectRef: ref, projectTitle: name,
            values: collVals, rowId: row.collectionRowId
          });

          /* 3 ── other tables ── */
          updateLoadingMessage(overlay, "Updating other tables...");
          const others: Array<[string, string, number | null]> = [
            [REVENUE_LIST, "revenueRowId", num("editRevenueForecast")],
            [INVOICING_LIST, "invoicingRowId", num("editInvoicingPlanning")],
            [SEIBIES_LIST, "seibiesRowId", num("editSEIBIES")],
            [INVENTORY_LIST, "inventoryRowId", num("editInventory")]
          ];
          for (const [listName, rowIdKey, value] of others) {
            const vals: Record<string, number | string> = value !== 0 && value !== null ? { OBL: value } : {};
            await upsertRow({
              listName, projectId, projectRef: ref, projectTitle: name,
              values: vals,
              rowId: (row as any)[rowIdKey] ?? null
            });
          }

          removeLoadingOverlay();
          showToast("✅ Project data updated across all tables", "success");
          await loadMaster();
        } catch (error) {
          console.error("Edit save error:", error);
          removeLoadingOverlay();
          showToast("❌ Failed to update project data", "error");
        }

        if (editSection) editSection.style.display = "none";
        if (tableSection) tableSection.style.display = "block";
        currentEditing = null;
        editForm.reset();
      };

      /* ---------- download template ---------- */

      const downloadTemplate = (): void => {
        /* template mirrors the final "master data excel.csv" 1:1: a single header row
           (same 27 columns, same names, same sequence) followed by one row per project. */
        const headerRow: string[] = CSV_TEMPLATE.map(c => c.h);

        const rowValues = (row: MasterRow): Array<string | number> => {
          const out: Array<string | number> = [];
          const push = (v: any): void => { out.push(v === undefined || v === null ? "" : v); };
          CSV_TEMPLATE.forEach(c => {
            const r = row as any;
            const c2 = (row.collection || {}) as any;
            switch (c.key) {
              case "projectReference": push(r.projectReference); break;
              case "projectName": push(r.projectName); break;
              case "profitCentre": push(r.profitCentre); break;
              case "customerName": push(r.customerName); break;
              case "groupHead": push(r.groupHead); break;
              case "projectManager": push(r.projectManager); break;
              case "enggLead": push(r.enggLead); break;
              case "engSubLead1": push(r.engSubLead1); break;
              case "engSubLead2": push(r.engSubLead2); break;
              case "status": push(r.status); break;
              case "contractualDeliveryDate": push(r.contractualDeliveryDate); break;
              case "revisedContractualDeliveryDate": push(r.revisedContractualDeliveryDate); break;
              case "prQuality": push(formatPrVal(r.prQuality)); break;
              case "prSubmissionStatus": push(r.prSubmissionStatus); break;
              case "prApprovalStatus": push(r.prApprovalStatus); break;
              case "nwc": push(r.nwc === 0 ? "" : r.nwc); break;
              case "normalDue": push(c2.normalDue === 0 ? "" : c2.normalDue); break;
              case "normalOverdue": push(c2.normalOverdue === 0 ? "" : c2.normalOverdue); break;
              case "retentionDue": push(c2.retentionDue === 0 ? "" : c2.retentionDue); break;
              case "retentionOverdue": push(c2.retentionOverdue === 0 ? "" : c2.retentionOverdue); break;
              case "gstTds": push(c2.gstTds === 0 ? "" : c2.gstTds); break;
              case "wht": push(c2.wht === 0 ? "" : c2.wht); break;
              case "receivableValue": push(c2.receivableValue === 0 ? "" : c2.receivableValue); break;
              case "obl": push(r.revenueForecast === 0 ? "" : r.revenueForecast); break;
              case "invoicingPlanning": push(r.invoicingPlanning === 0 ? "" : r.invoicingPlanning); break;
              case "seibies": push(r.seibies === 0 ? "" : r.seibies); break;
              case "inventory": push(r.inventory === 0 ? "" : r.inventory); break;
              default: push("");
            }
          });
          return out;
        };

        const aoa: Array<Array<string | number>> = [headerRow];
        if (masterData.length) {
          masterData.forEach(m => aoa.push(rowValues(m)));
        } else {
          /* nothing loaded → one clearly-labelled sample row (delete it before importing) */
          const sample: Array<string | number> = [];
          const s: Record<string, any> = {
            profitCentre: "IN50334", projectReference: "REF-001", customerName: "Sample Customer",
            projectName: "Sample Customer Project — DELETE THIS ROW", groupHead: "KG",
            projectManager: "Kiran Tikate", enggLead: "Abhijeet Pokale",
            engSubLead1: "Yuvraj Bhade", engSubLead2: "Kiran Tikate", status: "REL",
            contractualDeliveryDate: "10-07-2027", revisedContractualDeliveryDate: "10-07-2027",
            prQuality: "93%", prSubmissionStatus: "Submitted", prApprovalStatus: "Approved",
            nwc: 45000, normalDue: 100000, normalOverdue: 50000, retentionDue: 20000,
            retentionOverdue: 10000, gstTds: 18000, wht: 10000, receivableValue: 150000,
            obl: 500000, invoicingPlanning: 300000, seibies: 200000, inventory: 100000
          };
          CSV_TEMPLATE.forEach(c => {
            const v = s[c.key];
            sample.push(v === undefined || v === "" ? "" : v);
          });
          aoa.push(sample);
        }

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Master Data");
        const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
        saveAs(new Blob([buf], { type: "application/octet-stream" }), "Master_Data_Template.xlsx");
        showToast(masterData.length
          ? `📥 Downloaded ${masterData.length} project(s) — same columns as "master data excel.csv". Edit / add rows, then import the same file again.`
          : "📥 Empty template downloaded (no data loaded yet).", "success");
      };

      /* ---------- event binding ---------- */

      importBtn?.addEventListener("click", openImportPopup);
      importFileEl?.addEventListener("change", (e: Event) => {
        const el = e.target as HTMLInputElement;
        const file = el.files?.[0];
        if (!file) return;
        void doImport(file);
        el.value = ""; // allow re-selecting the same file
      });
      importChooseBtn?.addEventListener("click", () => importFileEl?.click());
      popupTemplateBtn?.addEventListener("click", downloadTemplate);
      popupCloseBtn?.addEventListener("click", closeImportPopup);
      popupDoneBtn?.addEventListener("click", closeImportPopup);
      popupOverlay?.addEventListener("click", (e: MouseEvent) => {
        if (!importing && e.target === popupOverlay) closeImportPopup();
      });

      syncAllBtn?.addEventListener("click", async () => {
        if (!masterData.length) {
          showToast("Load the data first", "info");
          return;
        }
        syncAllBtn.disabled = true;
        try {
          await syncAllTables(false);
        } finally {
          syncAllBtn.disabled = false;
        }
        await loadMaster();
      });

      downloadTemplateBtn?.addEventListener("click", downloadTemplate);

      searchInput?.addEventListener("input", () => {
        currentPage = 1;
        const q = (searchInput.value || "").toLowerCase();
        renderedData = masterData.filter(r =>
          r.projectReference.toLowerCase().includes(q) ||
          r.projectName.toLowerCase().includes(q) ||
          r.customerName.toLowerCase().includes(q) ||
          r.status.toLowerCase().includes(q) ||
          r.projectManager.toLowerCase().includes(q) ||
          r.groupHead.toLowerCase().includes(q)
        );
        renderTable();
      });

      resetSearchBtn?.addEventListener("click", () => {
        if (searchInput) searchInput.value = "";
        renderedData = [...masterData];
        currentPage = 1;
        renderTable();
      });

      rowsPerPageEl?.addEventListener("change", () => {
        const v = parseInt(rowsPerPageEl.value, 10);
        pageSize = !v || v < 1 ? 0 : v; // "All"
        currentPage = 1;
        renderTable();
      });

      editForm?.addEventListener("submit", onEditSubmit);
      backBtn?.addEventListener("click", () => {
        if (editSection) editSection.style.display = "none";
        if (tableSection) tableSection.style.display = "block";
        currentEditing = null;
        editForm?.reset();
      });

      /* ---------- boot: load, then auto-sync all tables once ---------- */
      const boot = async (): Promise<void> => {
        await loadMaster();
        if (!AUTO_SYNC_ON_LOAD) return;
        if (autoSyncDone) return;
        autoSyncDone = true;
        if (!masterData.length) return;
        console.log("🔄 Master page load → syncing all tables...");
        const created = await syncAllTables(true);
        if (created > 0) await loadMaster();
      };
      void boot();
    }, 300);
  }
};

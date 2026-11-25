import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import "../styles/common.css";

declare const _spPageContextInfo: any;

/* =====================================================================
   MASTER DATA PAGE (multi-table)
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

interface ColDef {
  key: string;
  label: string;
  aliases: string[];
  source: string;
  kind: 'text' | 'number' | 'date';
}

interface ParsedRow extends Record<string, any> {
  __raw?: Array<string | number>;
  __rowNo?: number;
}

interface ImportErrorRow {
  rowNo: number;
  projectReference: string;
  projectName: string;
  missing: string[];
  message: string;
  raw: Array<string | number>;
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

const AUTO_SYNC_ON_LOAD = true;

const siteUrl: string =
  typeof _spPageContextInfo !== "undefined"
    ? _spPageContextInfo.webAbsoluteUrl
    : window.location.origin + "/sites/Kriya";

/* ------------------------------------------------------------------ */
/* Small helpers                                                      */
/* ------------------------------------------------------------------ */

const norm = (s: string): string => s.trim().toLowerCase().replace(/[\s,％%()/]+/g, '');

const pad2 = (n: number): string => String(n).padStart(2, "0");

const inrNum = (n: number): string =>  Math.round(Number(n) || 0).toLocaleString("en-US");

 const attachCommaFormatter = (el: HTMLInputElement | null): void => {
  if (!el) return;
  el.addEventListener("focus", () => {
    el.value = el.value.replace(/,/g, "");
  });
  el.addEventListener("blur", () => {
    const n = parseExcelNumber(el.value);
    el.value = n === null || n === 0 ? (n === 0 ? "0" : "") : n.toLocaleString("en-US");
  });
};

const setNumInput = (id: string, n: number): void => {
  const el = document.getElementById(id) as HTMLInputElement | null;
  if (el) el.value = n === 0 ? "0" : Number(n).toLocaleString("en-US");
};

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

const parseNumText = (v: any): number | null => {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  const s = String(v ?? "").trim().replace(/％/g, "%");
  if (!s) return null;
  return parseExcelNumber(s.endsWith("%") ? s.slice(0, -1) : s);
};

const formatPrVal = (v: any): string => {
  if (v === undefined || v === null || String(v).trim() === "") return "";
  const s = String(v).trim();
  return /%$/.test(s) ? s : s + "%";
};

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
/* Loading overlay + toasts                                           */
/* ------------------------------------------------------------------ */

const createLoadingOverlay = (message = "Processing..."): HTMLElement => {
  removeLoadingOverlay();
  console.log(`⏳ ${message}`);
  const overlay = document.createElement("div");
  overlay.className = "loading-overlay";
  overlay.setAttribute("role", "progressbar");
  overlay.setAttribute("aria-busy", "true");
  overlay.setAttribute("aria-label", message);
  overlay.innerHTML = `
    <div class="master-loader-card">
      <div class="master-page-loader">
        <div class="master-page-loader__track"><div class="master-page-loader__bar"></div></div>
      </div>
    </div>`;
  return overlay;
};

const updateLoadingMessage = (overlay: HTMLElement, message: string): void => {
  console.log(`⏳ ${message}`);
  if (overlay) overlay.setAttribute("aria-label", message);
};

const removeLoadingOverlay = (): void => {
  document.querySelectorAll(".loading-overlay").forEach(el => el.remove());
};

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
/* Excel column registry                                              */
/* ------------------------------------------------------------------ */

const PROJECT_TEXT_COLS: ColDef[] = [
  { key: "projectReference", label: "Project Reference", source: PROJECT_LIST, kind: "text",
    aliases: ["projectreference", "projectdefinition(reference)", "projectdefinitionreference", "project_x0020_definition_x0028_reference_x0029_"] },
  { key: "projectName", label: "Project Name", source: PROJECT_LIST, kind: "text",
    aliases: ["projectname", "description(projectname)", "descriptionprojectname", "description"] },
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
  { key: "nwc", label: "NWC", source: PROJECT_LIST, kind: "number", aliases: ["nwc"] },
];

const COLLECTION_COLS: ColDef[] = [
  { key: "normalDue", label: "NormalDue", source: COLLECTION_LIST, kind: "number", aliases: ["normaldue", "normalreceivables", "normalreceivable"] },
  { key: "normalOverdue", label: "NormalOverdue", source: COLLECTION_LIST, kind: "number", aliases: ["normaloverdue", "normaloverduereceivables", "normaloverduereceivable"] },
  { key: "retentionDue", label: "RetentionDue", source: COLLECTION_LIST, kind: "number", aliases: ["retentiondue", "retentionreceivables", "retentionreceivable"] },
  { key: "retentionOverdue", label: "RetentionOverdue", source: COLLECTION_LIST, kind: "number", aliases: ["retentionoverdue", "retentionoverduereceivables", "retentionoverduereceivable"] },
  { key: "gstTds", label: "GstTds", source: COLLECTION_LIST, kind: "number", aliases: ["gsttds", "gst", "gsttdsx"] },
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
  return ALL_COLS.find(c => c.aliases.some(a => a === n || norm(a) === n)) || null;
};

/* =====================================================================
   UI CONTENT + STYLE
   ===================================================================== */

const pageContent = `
<div class="project-page fade-in master-page">
  <section class="table-section card master-dashboard">
    <div class="page-header">
      <h1 style="color:#ffffff !important;-webkit-text-fill-color:#ffffff;">Master Data</h1>
    </div>

    <div class="header-actions">
      <button class="btn-primary import-master-btn">📤 Upload Master Data</button>
      <button class="btn-secondary sync-all-btn" title="Create any missing row on every source table for each project">🔄 Sync All Tables</button>
      <button class="btn-secondary download-template-btn">📥 Download Data</button>
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
          </tr>
        </thead>
        <tbody class="project-table-body"></tbody>
      </table>
    </div>

    <div class="pagination-container">
      <div class="pagination-left">
        <div class="rows-per-page">
          <label for="rowsPerPageSelect">Items per page:</label>
          <select id="rowsPerPageSelect" class="rows-per-page-select">
            <option value="10" >10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
            <option value="0">All</option>
          </select>
        </div>
        <div class="pagination-info">Showing <span class="record-start" id="masterRecordStart">0</span> - <span class="record-end" id="masterRecordEnd">0</span> of <span class="record-count" id="masterRecordCount">0</span> items</div>
      </div>
      <div class="pagination-right">
        <button type="button" class="master-pg-btn" id="masterFirstPage" disabled title="First page">&#171;</button>
        <button type="button" class="master-pg-btn" id="masterPrevPage" disabled title="Previous page">&#8249;</button>
        <div class="pagination" id="masterPageNumbers"></div>
        <span class="master-pg-of">of <span id="masterTotalPages">1</span></span>
        <button type="button" class="master-pg-btn" id="masterNextPage" disabled title="Next page">&#8250;</button>
        <button type="button" class="master-pg-btn" id="masterLastPage" disabled title="Last page">&#187;</button>
      </div>
    </div>

  </section>

  <section class="edit-section card" style="display:none;">
    <div class="form-header"><h2 style="color:#ffffff !important;-webkit-text-fill-color:#ffffff;">✏️ Edit Master Data</h2></div>
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
        <div class="form-group"><label>NWC</label><input type="text" inputmode="numeric" id="editNWC"  /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>PR+ Quality %</label><input type="text" id="editPrQuality" /></div>
        <div class="form-group"><label>PR+ Submission Status</label><input type="text" id="editPrSubmissionStatus" /></div>
        <div class="form-group"><label>PR+ Approval Status</label><input type="text" id="editPrApprovalStatus" /></div>
      </div>

      <div class="form-section-header">Collection Data</div>
      <div class="form-row">
        <div class="form-group"><label>NormalDue</label><input type="text" inputmode="numeric" id="editNormalDue"  /></div>
        <div class="form-group"><label>NormalOverdue</label><input type="text" inputmode="numeric" id="editNormalOverdue"  /></div>
        <div class="form-group"><label>RetentionDue</label><input type="text" inputmode="numeric" id="editRetentionDue"  /></div>
        <div class="form-group"><label>RetentionOverdue</label><input type="text" inputmode="numeric" id="editRetentionOverdue"  /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>GstTds</label><input type="text" inputmode="numeric" id="editGstTds"  /></div>
        <div class="form-group"><label>IT TDS / Others (WHT)</label><input type="text" inputmode="numeric" id="editWHT"  /></div>
        <div class="form-group"><label>Receivables (incl TDS)</label><input type="text" inputmode="numeric" id="editReceivableValue"  /></div>
      </div>

      <div class="form-section-header">Other Tables</div>
      <div class="form-row">
        <div class="form-group"><label>OBL</label><input type="text" inputmode="numeric" id="editRevenueForecast"  /></div>
        <div class="form-group"><label>Balance to Invoice</label><input type="text" inputmode="numeric" id="editInvoicingPlanning"  /></div>
        <div class="form-group"><label>BIES/SEI</label><input type="text" inputmode="numeric" id="editSEIBIES"  /></div>
        <div class="form-group"><label>Inventory</label><input type="text" inputmode="numeric" id="editInventory"  /></div>
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
          <button type="button" class="btn-secondary import-template-btn">📥 2. Download Data (contains ALL current data)</button>
        </div>
        <div class="import-hint">
          💡 <strong>Template contains every project with current values</strong> — edit existing rows or add new rows
          (new Project Reference) in Excel, save, then choose that file again.<br />
          Existing <strong>Project Reference → updated</strong> on every table · New <strong>Project Reference → project
          created</strong>, then rows ensured on every source table. Rows without a Project Reference are ignored.<br />
          <strong>Mandatory fields:</strong> Profit Centre, Project Ref, Customer name, Project Name, GH — any row
          missing one of these is rejected and listed in the downloadable error file.
        </div>
        <div class="import-progress" style="display:none;">
          <div class="master-page-loader"><div class="master-page-loader__track"><div class="master-page-loader__bar"></div></div></div>
        </div>
        <div class="import-results" style="display:none;">
          <div class="import-results-head">📊 Import status — one line per project row</div>
          <div class="import-summary-grid">
            <div class="is-item ok"><span class="lbl">Updated</span><span class="val" id="impUpdated">0</span></div>
            <div class="is-item new"><span class="lbl">New Projects</span><span class="val" id="impCreated">0</span></div>
            <div class="is-item err"><span class="lbl">Errors</span><span class="val" id="impErrors">0</span></div>
            <div class="is-item tot"><span class="lbl">Total</span><span class="val" id="impTotal">0</span></div>
          </div>

          <div class="import-error-actions">
            <button type="button" class="btn-primary download-errors-btn" style="display:none;">
              ⬇️ Download Error List (.xlsx)
            </button>
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
.page-header h1 { color:#0056b3; margin:0; }
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
.form-group input[type="text" inputmode="numeric"]::-webkit-inner-spin-button { opacity:1; }
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

/* NEW: download error list */
.import-error-actions { padding:10px 13px 2px; background:#fff; }
.import-error-actions .download-errors-btn { background:#dc3545 !important; border-color:#c82333 !important; }
.import-error-actions .download-errors-btn:hover:not(:disabled) { background:#c82333 !important; border-color:#b21f2d !important; }

@media (max-width: 900px) {
  .page-header { flex-direction:column; align-items:stretch; }
  .header-actions { flex-direction:column; }
  .summary-grid { grid-template-columns:1fr; }
  .form-group { flex:1 1 100%; }
  .master-table { font-size:12px; }
}

/* ---------------- 1. flat pagination bar ---------------- */
.pagination-container {
  margin-top:16px !important;
  display:flex !important;
  justify-content:space-between !important;
  align-items:center !important;
  gap:16px !important;
  flex-wrap:nowrap !important;
  overflow-x:auto !important;
}
.pagination-container .pagination-left { display:inline-flex !important; align-items:center !important; gap:14px !important; flex-wrap:nowrap !important; white-space:nowrap !important; }
.pagination-container .pagination-right { display:inline-flex !important; align-items:center !important; gap:8px !important; flex-wrap:nowrap !important; white-space:nowrap !important; }
.pagination-container .rows-per-page { display:inline-flex !important; align-items:center !important; gap:6px !important; font-size:13px !important; white-space:nowrap !important; }
.pagination-container .rows-per-page label { margin:0 !important; color:#667085 !important; font-weight:500 !important; }
.pagination-container .rows-per-page select {
  height:32px !important; padding:0 8px !important; border:1.5px solid #E4E7EC !important;
  border-radius:8px !important; font-size:13px !important; background:#fff !important; cursor:pointer !important;
}
.pagination-container .rows-per-page select:focus { border-color:#605EEE !important; outline:none !important; box-shadow:0 0 0 3px rgba(96,94,238,.12) !important; }

.pagination-container .pagination-info {
  display:inline-flex !important; align-items:center !important; gap:5px !important;
  padding:5px 12px !important; border-radius:999px !important;
  background:#F4F5FF !important; border:1px solid #E4E4FF !important;
  font-size:12.5px !important; font-weight:500 !important; color:#3F3F6B !important;
  white-space:nowrap !important;
}
.pagination-container .record-start,
.pagination-container .record-end,
.pagination-container .record-count { font-weight:700 !important; color:#4C49E0 !important; }

.pagination-container .pagination { display:inline-flex !important; align-items:center !important; gap:6px !important; flex-wrap:nowrap !important; }
.pagination-container .pagination button {
  min-width:32px !important; height:32px !important; padding:0 10px !important;
  display:inline-flex !important; align-items:center !important; justify-content:center !important;
  border:1.5px solid #E4E7EC !important; border-radius:8px !important; background:#fff !important;
  color:#1F2430 !important; font-size:13px !important; font-weight:500 !important;
  cursor:pointer !important; transition:all .18s ease !important;
}
.pagination-container .pagination button:hover:not(:disabled):not(.active) { border-color:#605EEE !important; color:#4C49E0 !important; background:#F7F7FF !important; }
.pagination-container .pagination button.active { background:#605EEE !important; border-color:#605EEE !important; color:#fff !important; font-weight:600 !important; }
.pagination-container .pagination button.dots { border-color:transparent !important; background:transparent !important; color:#98A2B3 !important; cursor:default !important; min-width:18px !important; padding:0 2px !important; }

.master-pg-btn {
  width:32px !important; height:32px !important; min-width:32px !important;
  display:inline-flex !important; align-items:center !important; justify-content:center !important;
  padding:0 !important; font-size:14px !important; line-height:1 !important;
  border:1.5px solid #E4E7EC !important; border-radius:8px !important;
  background:#fff !important; color:#475467 !important; cursor:pointer !important;
  transition:all .18s ease !important;
}
.master-pg-btn:hover:not(:disabled) { border-color:#605EEE !important; color:#4C49E0 !important; background:#EEEEFE !important; }
.master-pg-btn:disabled { opacity:.42 !important; cursor:not-allowed !important; }
.master-pg-of { font-size:13px !important; color:#667085 !important; white-space:nowrap !important; margin:0 2px !important; }
.master-pg-of > span { font-weight:700 !important; color:#1F2430 !important; }

@media (max-width:900px) {
  .pagination-container { gap:10px !important; }
  .pagination-container .pagination-left { gap:10px !important; }
  .pagination-container .pagination-info { padding:4px 10px !important; font-size:12px !important; }
  .master-pg-btn, .pagination-container .pagination button { min-width:30px !important; height:30px !important; }
}

/* ---------------- 2. page loader: grey track + indigo bar ---------------- */
.loading-overlay {
  position:absolute !important; top:0 !important; left:0 !important;
  width:100% !important; height:100% !important;
  background:rgba(255,255,255,.94) !important;
  display:flex !important; align-items:center !important; justify-content:center !important;
  z-index:99999 !important; backdrop-filter:blur(3px) !important;
  border-radius:12px !important; overflow:hidden !important; box-sizing:border-box !important;
}
.master-loader-card {
  background:#fff; border:1px solid #e8e8e8; border-radius:12px;
  box-shadow:0 8px 32px rgba(0,0,0,.15); padding:24px 28px; min-width:260px;
}
.master-page-loader { width:100%; }
.master-page-loader__track { position:relative; height:6px; width:100%; background:#DCDCDC; border-radius:999px; overflow:hidden; }
.master-page-loader__bar {
  position:absolute; top:0; left:-42%; width:42%; height:100%;
  background:#6664FF; border-radius:999px; animation:masterBarSlide 1.05s ease-in-out infinite;
}
@keyframes masterBarSlide { 0% { left:-42%; } 100% { left:100%; } }

.import-progress { align-items:center !important; gap:10px !important; background:#eef2f7 !important; border:1px solid #dde4ee !important; border-radius:8px !important; padding:12px 14px !important; }
.import-progress .master-page-loader { flex:1 1 auto !important; }
.import-progress .master-page-loader__track { height:5px !important; }

.loading-container, .loading-spinner, .loading-text, .loading-progress { display:none !important; }

/* ---------------- 3. pager alignment hardening ---------------- */
.pagination-container .pagination-right { align-items:center !important; }
.pagination-container .pagination,
.pagination-container .pagination button,
.pagination-container .pagination button.active,
.pagination-container .pagination button.dots,
.master-pg-btn {
  box-sizing:border-box !important;
  height:32px !important;
  min-height:32px !important;
  max-height:32px !important;
  padding-top:0 !important;
  padding-bottom:0 !important;
  line-height:1 !important;
  vertical-align:middle !important;
  align-self:center !important;
  margin:0 !important;
  transform:none !important;
  position:static !important;
}
.pagination-container .pagination button,
.pagination-container .pagination button.active,
.master-pg-btn {
  display:inline-flex !important;
  align-items:center !important;
  justify-content:center !important;
}
.master-pg-of { display:inline-flex !important; align-items:center !important; height:32px !important; line-height:1 !important; }

.pagination-container .master-pg-btn {
  width: auto !important;
  min-width: 26px !important;
  height: 32px !important;
  min-height: 32px !important;
  padding: 4px 5px !important;
  background: transparent !important;
  border: 0 !important;
  border-radius: 4px !important;
  color: #475467 !important;
  font-family: Georgia, "Times New Roman", serif !important;
  font-size: 16px !important;
  line-height: 1 !important;
  box-shadow: none !important;
  transform: none !important;
}
.pagination-container .master-pg-btn:hover:not(:disabled) {
  background: #f1f3f5 !important;
  border: 0 !important;
  color: #344054 !important;
  box-shadow: none !important;
  transform: none !important;
}
.pagination-container .master-pg-btn:disabled {
  background: transparent !important;
  border: 0 !important;
  color: #98a2b3 !important;
  box-shadow: none !important;
  transform: none !important;
}

/* ---------------- 4. title banner + header action row ---------------- */
.master-dashboard .page-header {
  margin: 0 0 14px 0;
}
.master-dashboard .page-header h1 { color: #fff !important; }
.master-dashboard .page-header, .master-dashboard .page-header h1, .master-dashboard .page-header h1 * { color: #fff !important; }
.master-dashboard .page-header, .master-dashboard .page-header h1, .master-dashboard .page-header h1 *,
.master-page .form-header, .master-page .form-header h2 { -webkit-text-fill-color: #fff !important; background-clip: border-box; }
.master-dashboard .header-actions { margin: 0 0 18px 0 !important; padding-left: 0 !important; }

/* ---------------- 5. Project Reference as the edit link ---------------- */
.master-table td .edit-link {
  display:inline-block;
  font-size:13px;
  font-weight:700;
  color:var(--lilac-primary, #605EEE);
  text-decoration:none;
  border-bottom:1px dashed rgba(96,94,238,0.45);
  cursor:pointer;
  white-space:nowrap;
  transition:color 0.18s ease, border-color 0.18s ease, background 0.18s ease;
}
.master-table td .edit-link:hover { color:var(--lilac-primary-dark, #4C49E0); border-bottom:1px solid var(--lilac-primary-dark, #4C49E0); }
.master-table td .edit-link:focus-visible { outline:2px solid #605EEE; outline-offset:2px; border-radius:3px; }
.master-table td .edit-link:active { transform:translateY(1px); }

/* ---------------- 6. lilac token palette ---------------- */
.project-page {
  --lilac-primary: #605EEE;
  --lilac-primary-dark: #4C49E0;
  --lilac-primary-soft: #EEEEFE;
  --lilac-ink: #1F2430;
  --lilac-muted: #667085;
  --lilac-line: #E4E7EC;
  --lilac-pill: 999px;
  --lilac-shadow-sm: 0 1px 2px rgba(16, 24, 40, .06);
}

/* ---------------- 7. Edit Master Data title wears the SAME banner ---------------- */
.master-page .form-header {
  margin: 0 0 18px 0;
}
.master-page .form-header h2 {
  color: #fff !important;
  margin: 0;
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
      const recordStartEl = document.getElementById("masterRecordStart") as HTMLElement | null;
      const recordEndEl = document.getElementById("masterRecordEnd") as HTMLElement | null;
      const recordCountEl = document.getElementById("masterRecordCount") as HTMLElement | null;
      const totalPagesEl = document.getElementById("masterTotalPages") as HTMLElement | null;
      const firstPageBtn = document.getElementById("masterFirstPage") as HTMLButtonElement | null;
      const prevPageBtn = document.getElementById("masterPrevPage") as HTMLButtonElement | null;
      const nextPageBtn = document.getElementById("masterNextPage") as HTMLButtonElement | null;
      const lastPageBtn = document.getElementById("masterLastPage") as HTMLButtonElement | null;

      const popupOverlay = document.querySelector(".import-popup-overlay") as HTMLElement | null;
      const popupResults = document.querySelector(".import-results") as HTMLElement | null;
      const popupStatusList = document.querySelector(".import-status-list") as HTMLElement | null;
      const popupProgress = document.querySelector(".import-progress") as HTMLElement | null;
      const importFileEl = document.querySelector(".import-file-input") as HTMLInputElement | null;
      const importChooseBtn = document.querySelector(".import-choose-btn") as HTMLButtonElement | null;
      const popupTemplateBtn = document.querySelector(".import-template-btn") as HTMLButtonElement | null;
      const popupCloseBtn = document.querySelector(".import-popup-close") as HTMLButtonElement | null;
      const popupDoneBtn = document.querySelector(".popup-done-btn") as HTMLButtonElement | null;
      const impUpdatedEl = document.getElementById("impUpdated") as HTMLElement | null;
      const impCreatedEl = document.getElementById("impCreated") as HTMLElement | null;
      const impErrorsEl = document.getElementById("impErrors") as HTMLElement | null;
      const impTotalEl = document.getElementById("impTotal") as HTMLElement | null;
      const downloadErrorsBtn = document.querySelector(".download-errors-btn") as HTMLButtonElement | null;

      let masterData: MasterRow[] = [];
      let currentEditing: MasterRow | null = null;
      let currentPage = 1;
      let pageSize = 25;
      const rowsPerPageEl = document.getElementById("rowsPerPageSelect") as HTMLSelectElement | null;
      let renderedData: MasterRow[] = [];
      let autoSyncDone = false;

      /* ---------- Profit Centre field resolution (fix) ---------- */
      let profitCentreInternalName: string | null = null;

      const resolveProfitCentreField = async (): Promise<string | null> => {
        if (profitCentreInternalName) return profitCentreInternalName;
        const fields = await getListFields(PROJECT_LIST);
        const candidates = [
          "ProfitCentre", "ProfitCenter", "Profit_x0020_Centre", "Profit_x0020_Center",
          "PC", "Profit Centre", "Profit Center"
        ];
        for (const c of candidates) {
          const f = findField(fields, c);
          if (f) {
            profitCentreInternalName = f.internalName;
            console.log(`✅ Profit Centre resolved → internal name "${f.internalName}" (title "${f.title}")`);
            return f.internalName;
          }
        }
        const fuzzy = fields.find(f => /profit/i.test(f.title) || /profit/i.test(f.internalName));
        if (fuzzy) {
          profitCentreInternalName = fuzzy.internalName;
          console.log(`✅ Profit Centre fuzzy-resolved → internal name "${fuzzy.internalName}" (title "${fuzzy.title}")`);
          return fuzzy.internalName;
        }
        console.warn("⚠ Profit Centre field not found. Available Project List fields:",
          fields.map(f => `${f.internalName} (${f.title})`).join(", "));
        return null;
      };

      /* ---------- loading / fetch master ---------- */

      let loadBusy = false;
      const loadMaster = async (): Promise<void> => {
        if (loadBusy) return;
        loadBusy = true;
        const loading = createLoadingOverlay("Loading master data...");
        loaderHost().appendChild(loading);
        try {
          updateLoadingMessage(loading, "Loading data from all 6 tables...");

          const pcInternal = await resolveProfitCentreField();
          const pcDesired = pcInternal || "ProfitCentre";

          const TABLE_LOAD: Array<[string, string[]]> = [
            [PROJECT_LIST, ["Title", "ProjectReference", pcDesired, "CustomerName", "GroupHead1", "ProjectManager", "EnggLead", "EnggSubLead1", "EnggSubLead2", "Status", "ContractualDeliveryDate", "RevisedContractualDeliveryDate", "PRQuality", "PRSubmissionStatus", "PRApprovalStatus", "NWC"]],
            [COLLECTION_LIST, ["Project", "ProjectReference", "NormalDue", "NormalOverdue", "RetentionDue", "RetentionOverdue", "GstTds", "WHT", "ReceivableValue", "OBL"]],
            [REVENUE_LIST, ["Project", "ProjectReference", "OBL"]],
            [INVOICING_LIST, ["Project", "ProjectReference", "OBL"]],
            [SEIBIES_LIST, ["Project", "ProjectReference", "OBL"]],
            [INVENTORY_LIST, ["Project", "ProjectReference", "OBL"]]
          ];
          console.log(`📋 Project List fields resolved for load: ${TABLE_LOAD[0][1].join(", ")}`);

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
           const projectName = title === "" ? "-" : title;
            const collRow = findIn(coll, ref, title, p.Id);
            const revRow = findIn(rev, ref, title, p.Id);
            const invRow = findIn(inv, ref, title, p.Id);
            const seiRow = findIn(sei, ref, title, p.Id);
            const invtRow = findIn(invt, ref, title, p.Id);

            return {
              id: p.Id,
              projectReference: ref,
              projectName,
              profitCentre: String(
                (pcInternal && p[pcInternal]) ?? p.ProfitCentre ?? p.ProfitCenter ?? p.PC ?? ""
              ),
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

          /* sanity check: catch a silently-dropped column before it exports blank */
          if (masterData.length > 0) {
            const checks: Array<[string, (r: MasterRow) => boolean]> = [
              ["Profit Centre", r => r.profitCentre !== ""],
              ["Project Reference", r => r.projectReference !== ""],
              ["Customer Name", r => r.customerName !== ""],
            ];
            const dead = checks.filter(([, test]) => masterData.every(r => !test(r))).map(([n]) => n);
            if (dead.length) {
              console.warn(`⚠ These columns are EMPTY on every loaded project: ${dead.join(", ")}. Check that the SharePoint internal names match.`);
              showToast(`⚠ Some columns loaded empty: ${dead.join(", ")}. See console.`, "warning", 8000);
            }
          }

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

      const editRefCell = (ref: string): string =>
        ref
          ? `<a href="#" class="edit-link" data-ref="${esc(ref)}" title="Click to edit this project">${esc(ref)}</a>`
          : `<strong>N/A</strong>`;

      const renderTable = (): void => {
        if (!tableBody) return;
        if (renderedData.length === 0) {
          tableBody.innerHTML = `<tr><td colspan="12" style="text-align:center;padding:40px;color:#666;"><h3>No projects found</h3><p>Import an Excel file or check the Project List.</p></td></tr>`;
          renderPagination();
          return;
        }
        const effectivePageSize = pageSize > 0 ? pageSize : Math.max(1, renderedData.length);
        const totalPages = Math.max(1, Math.ceil(renderedData.length / effectivePageSize));
        if (currentPage > totalPages) currentPage = totalPages;
        const start = (currentPage - 1) * effectivePageSize;
        const page = renderedData.slice(start, start + effectivePageSize);

        const rowsHtml = page.map((r, idx) => {
          const c = r.collection;
          return `<tr>
            <td>${start + idx + 1}</td>
            <td>${editRefCell(r.projectReference)}</td>
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
          </tr>`;
        }).join("");

        tableBody.innerHTML = rowsHtml;
        attachEditHandlers();
        renderPagination();
      };

      const goToPage = (page: number): void => {
        const effective = pageSize > 0 ? pageSize : Math.max(1, renderedData.length);
        const totalPages = Math.max(1, Math.ceil(renderedData.length / effective));
        if (page < 1 || page > totalPages || page === currentPage) return;
        currentPage = page;
        renderTable();
      };

      const updatePaginationInfo = (): void => {
        const effective = pageSize > 0 ? pageSize : Math.max(1, renderedData.length);
        const total = renderedData.length;
        const totalPages = Math.max(1, Math.ceil(total / effective));
        const startIndex = total === 0 ? 0 : (currentPage - 1) * effective + 1;
        const endIndex = total === 0 ? 0 : Math.min(currentPage * effective, total);
        if (recordStartEl) recordStartEl.textContent = String(startIndex);
        if (recordEndEl) recordEndEl.textContent = String(endIndex);
        if (recordCountEl) recordCountEl.textContent = String(total);
        if (totalPagesEl) totalPagesEl.textContent = String(totalPages);
      };

      const renderPagination = (): void => {
        if (!paginationDiv) return;
        updatePaginationInfo();

        const effective = pageSize > 0 ? pageSize : Math.max(1, renderedData.length);
        const totalPages = Math.max(1, Math.ceil(renderedData.length / effective));
        const noData = renderedData.length === 0;

        if (firstPageBtn) firstPageBtn.disabled = noData || currentPage === 1;
        if (prevPageBtn) prevPageBtn.disabled = noData || currentPage === 1;
        if (nextPageBtn) nextPageBtn.disabled = noData || currentPage >= totalPages;
        if (lastPageBtn) lastPageBtn.disabled = noData || currentPage >= totalPages;

        paginationDiv.innerHTML = "";
        if (noData) return;

        const max = 5;
        let s = Math.max(1, currentPage - Math.floor(max / 2));
        let e = Math.min(totalPages, s + max - 1);
        if (e - s + 1 < max) s = Math.max(1, e - max + 1);

        const addDots = (): void => {
          const d = document.createElement("button");
          d.type = "button";
          d.className = "dots";
          d.textContent = "…";
          d.disabled = true;
          paginationDiv.appendChild(d);
        };

        if (s > 1) addDots();
        for (let i = s; i <= e; i++) {
          const b = document.createElement("button");
          b.type = "button";
          b.textContent = String(i);
          b.className = i === currentPage ? "active" : "";
          b.title = `Page ${i}`;
          if (i === currentPage) b.setAttribute("aria-current", "page");
          b.addEventListener("click", () => goToPage(i));
          paginationDiv.appendChild(b);
        }
        if (e < totalPages) addDots();
      };

      const attachEditHandlers = (): void => {
        document.querySelectorAll(".edit-link, .edit-btn").forEach(btn => {
          btn.addEventListener("click", (evt: Event) => {
            if (evt && typeof evt.preventDefault === "function") { evt.preventDefault(); }
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
            setNumInput("editNWC", row.nwc);
            setNumInput("editNormalDue", row.collection.normalDue);
            setNumInput("editNormalOverdue", row.collection.normalOverdue);
            setNumInput("editRetentionDue", row.collection.retentionDue);
            setNumInput("editRetentionOverdue", row.collection.retentionOverdue);
            setNumInput("editGstTds", row.collection.gstTds);
            setNumInput("editWHT", row.collection.wht);
            setNumInput("editReceivableValue", row.collection.receivableValue);
            setNumInput("editRevenueForecast", row.revenueForecast);
            setNumInput("editInvoicingPlanning", row.invoicingPlanning);
            setNumInput("editSEIBIES", row.seibies);
            setNumInput("editInventory", row.inventory);
             [
              "editNWC","editNormalDue","editNormalOverdue","editRetentionDue",
              "editRetentionOverdue","editGstTds","editWHT","editReceivableValue",
              "editRevenueForecast","editInvoicingPlanning","editSEIBIES","editInventory"
            ].forEach(id => attachCommaFormatter(document.getElementById(id) as HTMLInputElement | null));
            if (editSection) editSection.style.display = "block";
            if (tableSection) tableSection.style.display = "none";
          });   
        });
      };

      /* ---------- generic upsert ---------- */

      interface UpsertOpts {
        listName: string;
        projectId: number;
        projectRef: string;
        projectTitle: string;
        values: Record<string, number | string>;
        rowId?: number | null;
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

      const ensureChildRow = async (listName: string, row: MasterRow): Promise<boolean> => {
        const result = await upsertRow({
          listName,
          projectId: row.id,
          projectRef: row.projectReference,
          projectTitle: row.projectName,
          values: {},
          rowId: undefined,
          assumeNew: true
        });
        return result.success;
      };

      const syncAllTables = async (silent: boolean): Promise<number> => {
        let created = 0;
        for (const listName of CHILD_LISTS) {
          if (!(await checkListExists(listName))) {
            console.warn(`⏭️ ${listName} does not exist - skipped`);
            continue;
          }
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

      /* ---------- EXCEL IMPORT ---------- */

      const parseSheet = (ws: any): ParsedRow[] => {
        const aoa: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
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

        const out: ParsedRow[] = [];
        for (let ri = headerIdx + 1; ri < aoa.length; ri++) {
          const row = aoa[ri];
          if (!row || !row.some((c: any) => String(c ?? "").trim() !== "")) continue;
          const rec: ParsedRow = {
            __raw: header.map((_: any, ci: number) => (row[ci] === undefined ? "" : row[ci])),
            __rowNo: ri + 1
          };
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
         Excel import / export popup
         ================================================================== */

      let importing = false;
      let impUpdated = 0, impCreated = 0, impErrors = 0;
      let importErrorRows: ImportErrorRow[] = [];

     const MANDATORY_FIELDS: Array<{ key: string; label: string }> = [
            { key: "profitCentre",     label: "Profit Centre" },
            { key: "projectReference", label: "Project definition(Reference)" },
            { key: "customerName",     label: "Customer Name" },
            { key: "projectName",      label: "Description (Project Name)" },
            { key: "groupHead",        label: "GH" },
            { key: "projectManager",   label: "PM" } // Added PM
          ];

     const isBlank = (v: any): boolean => {
      return String(v ?? "").trim() === "";
    };

      const missingMandatory = (row: ParsedRow): string[] => {
        const missing: string[] = [];
        MANDATORY_FIELDS.forEach(f => {
          if (isBlank(row[f.key])) missing.push(f.label);
        });
        return missing;
      };

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
        importErrorRows = [];
        bumpChips();
        if (impTotalEl) impTotalEl.textContent = "0";
        if (popupStatusList) popupStatusList.innerHTML = "";
        if (popupProgress) popupProgress.style.display = "none";
        if (popupResults) popupResults.style.display = "none";
        if (downloadErrorsBtn) downloadErrorsBtn.style.display = "none";
        popupOverlay.style.display = "flex";
      };

      const closeImportPopup = (): void => {
        if (!popupOverlay || importing) return;
        popupOverlay.style.display = "none";
      };

      const downloadImportErrors = (): void => {
        if (!importErrorRows.length) {
          showToast("No error rows to download", "info");
          return;
        }
        const headerRow: string[] = CSV_TEMPLATE.map(c => c.h);
        const diagHeaders = ["__Row #", "__Missing Fields", "__Error"];

        const aoa: Array<Array<string | number>> = [headerRow.concat(diagHeaders)];

        importErrorRows.forEach(err => {
          const tplRow: Array<string | number> = CSV_TEMPLATE.map((c, ci) => {
            const rawVal = err.raw[ci];
            return rawVal === undefined || rawVal === null ? "" : rawVal;
          });
          aoa.push(tplRow.concat([
            err.rowNo,
            err.missing.join(", "),
            err.message
          ]));
        });

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws["!cols"] = headerRow.map(() => ({ wch: 20 }))
          .concat([{ wch: 8 }, { wch: 40 }, { wch: 60 }]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Import Errors");
        const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
        const stamp = new Date().toISOString().slice(0, 10);
        saveAs(new Blob([buf], { type: "application/octet-stream" }),
          `Master_Import_Errors_${stamp}.xlsx`);
        showToast(`📥 Downloaded ${importErrorRows.length} error row(s)`, "success");
      };

     /**
 * Resolve a Project List item by Project Reference.
 * Exact OData filter first (works when text + indexed), then a
 * case-insensitive client-side scan (handles lookup-typed refs,
 * casing drift, unindexed columns, and >5000-item lists).
 */
 const resolveProjectByRef = async (ref: string): Promise<any | null> => {
  const wanted = String(ref ?? "").trim();
  if (!wanted) return null;
  const wantedKey = wanted.toLowerCase();

  // 1) Fast path: server-side exact filter
  try {
    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items` +
      `?$select=Id,Title,ProjectReference&$filter=ProjectReference eq '${encodeURIComponent(wanted)}'&$top=5`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    if (res.ok) {
      const rows: any[] = (await res.json()).d.results || [];
      const hit = rows.find(r =>
        String(r.ProjectReference ?? "").trim().toLowerCase() === wantedKey
      );
      if (hit) return hit;
    }
  } catch (e) {
    console.warn("resolveProjectByRef: exact filter failed, falling back", e);
  }

  // 2) Fallback: case-insensitive client-side scan (bounded page)
  try {
    const fields = await getListFields(PROJECT_LIST);
    const refField = findField(fields, "ProjectReference");
    const select = ["Id", "Title"];
    const expand: string[] = [];
    if (refField) {
      if (refField.type === "Lookup" || refField.type === "User") {
        select.push(`${refField.internalName}/Id`, `${refField.internalName}/Title`);
        expand.push(refField.internalName);
      } else {
        select.push(refField.internalName);
      }
    } else {
      select.push("ProjectReference");
    }

    const query =
      `$select=${encodeURIComponent(select.join(","))}` +
      (expand.length ? `&$expand=${encodeURIComponent(expand.join(","))}` : "") +
      `&$top=5000`;

    const res = await fetchWithTimeout(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items?${query}`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    if (!res.ok) return null;
    const rows: any[] = (await res.json()).d.results || [];

    const readRef = (r: any): string => {
      if (!refField) return String(r.ProjectReference ?? "");
      const v = r[refField.internalName];
      if (v && typeof v === "object") return String(v.Title ?? v.Value ?? "");
      return String(v ?? "");
    };

    return rows.find(r => readRef(r).trim().toLowerCase() === wantedKey) || null;
  } catch (e) {
    console.error("resolveProjectByRef: fallback scan failed", e);
    return null;
  }
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
  importErrorRows = [];
  bumpChips();

  const seenRefs = new Set<string>();
  const resolvedProjectCache = new Map<string, { id: number; title: string }>();

  const fixSPDate = (dateStr: any): string | null => {
    if (!dateStr || String(dateStr).trim() === "" || dateStr === "-") return null;
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? null : d.toISOString();
  };

  const getCurrentMonthKey = (): string => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  };

  const cleanOldMonths = (data: Record<string, any>): Record<string, any> => {
    const current = new Date();
    const cutoffDate = new Date(current.getFullYear(), current.getMonth() - 36, 1);
    const cleaned: Record<string, any> = {};
    const sortedKeys = Object.keys(data).sort();

    for (const key of sortedKeys) {
      const parts = key.split("-");
      if (parts.length === 2) {
        const year = Number(parts[0]);
        const month = Number(parts[1]);
        if (year && month) {
          const d = new Date(year, month - 1, 1);
          if (d >= cutoffDate) {
            cleaned[key] = data[key];
          }
        }
      }
    }
    return cleaned;
  };

  // Reusable helper to safely fetch existing JSON from a SharePoint item
  const fetchExistingMonthJson = async (listName: string, itemId: number, fieldName: string): Promise<Record<string, any>> => {
    try {
      const res = await fetchWithTimeout(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})?$select=${fieldName}`, {
        headers: { Accept: "application/json;odata=verbose" }
      });
      if (res.ok) {
        const j = await res.json();
        const rawJson = j.d[fieldName];
        if (rawJson) return JSON.parse(rawJson);
      }
    } catch (e) {
      console.warn(`Could not parse existing ${fieldName} JSON for item ${itemId} in ${listName}`, e);
    }
    return {};
  };

  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = parseSheet(ws);
    if (!rows.length) throw new Error("No rows with a Project Reference found. Check the header row (Project definition(Reference)).");

    const total = rows.length;
    const currentMonthKey = getCurrentMonthKey();

    if (impTotalEl) impTotalEl.textContent = String(total);
    if (popupResults) popupResults.style.display = "none";
    if (downloadErrorsBtn) downloadErrorsBtn.style.display = "none";
    if (popupStatusList) popupStatusList.innerHTML = "";
    if (popupProgress) popupProgress.style.display = "flex";

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const refRaw = String(row.projectReference ?? "").trim();
      if (!refRaw) continue;

      const ref = refRaw.replace(/\s+/g, " ");
      const refKey = ref.toLowerCase();

      // ✅ Excel Description is used verbatim — "-", "#", "#N/A", numbers, symbols all kept.
      //    No fallback to the reference. Only truly-empty stays empty.
      const excelName = String(row.projectName ?? "").trim();
      row.projectName = excelName;

      console.log(`⏳ Processing ${i + 1} of ${total} — ${ref}`);

      const missing = missingMandatory(row);
      if (missing.length) {
        const msg = "Missing mandatory field(s): " + missing.join(", ");
        importErrorRows.push({
          rowNo: row.__rowNo ?? (i + 2),
          projectReference: ref,
          projectName: String(row.projectName ?? "").trim(),
          missing,
          message: msg,
          raw: row.__raw ?? []
        });
        addStatusRow("error", ref || "(no ref)", String(row.projectName ?? ""), msg);
        continue;
      }

      try {
        let projectId: number | null = null;
        let projectTitle = "";
        let isNew = false;

        const cached = resolvedProjectCache.get(refKey);
        if (cached) {
          projectId = cached.id;
          projectTitle = cached.title;
          isNew = false;
        } else {
          const projData = await resolveProjectByRef(ref);
          if (projData) {
            projectId = projData.Id;

            const storedTitle = String(projData.Title ?? "").trim();
            const excelTitle  = String(row.projectName ?? "").trim();

            // ✅ Excel Description wins (even "-", "#", numbers, symbols).
            //    Fall back to stored SharePoint title only when Excel is truly empty.
            //    Never fall back to the reference.
            projectTitle = excelTitle !== "" ? excelTitle : storedTitle;

            isNew = false;
            resolvedProjectCache.set(refKey, { id: projectId as number, title: projectTitle });
            seenRefs.add(refKey);
          }
        }

        if (projectId === null) {
          if (seenRefs.has(refKey)) {
            const again = await resolveProjectByRef(ref);
            if (!again) throw new Error(`Project "${ref}" was created earlier in this import but could not be re-resolved.`);
            projectId = again.Id;
            projectTitle = String(again.Title ?? "").trim();
            isNew = false;
            resolvedProjectCache.set(refKey, { id: projectId as number, title: projectTitle });
          } else {
            isNew = true;

            // ✅ Title = Excel Description only. No fallback to the reference.
            projectTitle = String(row.projectName ?? "").trim();

            const entityType = await getListEntityType(PROJECT_LIST);
            const plPayload: any = { __metadata: { type: entityType } };
            const plFields = await getListFields(PROJECT_LIST);

            const writePL = (desired: string, value: any, isDate = false): void => {
              const f = findField(plFields, desired);
              if (!f) return;
              if (f.type === "Lookup" || f.type === "User") return;

              if (isDate || f.type === "DateTime") {
                plPayload[f.internalName] = fixSPDate(value);
                return;
              }

              if (f.type === "Number" || f.type === "Currency") {
                const n = parseNumText(value);
                if (n !== null) plPayload[f.internalName] = n;
              } else if (value !== undefined && value !== null && String(value) !== "") {
                plPayload[f.internalName] = String(value);
              }
            };

            // ✅ Only set Title when we actually have one (avoid blanking a new item’s required title)
            if (projectTitle !== "") plPayload.Title = projectTitle;

            writePL("ProjectReference", ref);
            writePL("PC", row.profitCentre);
            writePL("CustomerName", row.customerName);
            writePL("NWC", row.nwc);

            // Prepare current month JSON for Project List
            const currentProjectData = {
              status: row.status || "",
              contractualDeliveryDate: row.contractualDeliveryDate || "",
              revisedContractualDeliveryDate: row.revisedContractualDeliveryDate || "",
              prQuality: row.prQuality || "",
              prSubmissionStatus: row.prSubmissionStatus || "",
              prApprovalStatus: row.prApprovalStatus || "",
              nwc: row.nwc ?? 0
            };
            plPayload["ProjectMonth"] = JSON.stringify({ [currentMonthKey]: currentProjectData });

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
              const again = await resolveProjectByRef(ref);
              if (!again) throw new Error(`Project "${ref}" was created but could not be re-resolved.`);
              projectId = again.Id;
            }
            resolvedProjectCache.set(refKey, { id: projectId as number, title: projectTitle });
            seenRefs.add(refKey);
          }
        }

        if (projectId === null) {
          throw new Error(`Could not resolve or create project "${ref}".`);
        }
        const resolvedProjectId: number = projectId;

        const notes: string[] = [];

        // =====================================================================
        // 1. UPDATE PROJECT LIST (Existing Project -> Append ProjectMonth JSON)
        // =====================================================================
        if (!isNew) {
          const update: Record<string, number | string | null> = {};
          const plFields = await getListFields(PROJECT_LIST);

          const put = (desired: string, value: any, isDate = false): void => {
            const f = findField(plFields, desired);
            if (!f || f.type === "Lookup" || f.type === "User") return;

            if (isDate || f.type === "DateTime") {
              update[f.internalName] = fixSPDate(value);
              return;
            }

            if (f.type === "Number" || f.type === "Currency") {
              const n = typeof value === "number" ? value : parseNumText(value);
              if (n !== null) update[f.internalName] = n;
            } else if (value !== undefined && String(value).trim() !== "") {
              update[f.internalName] = String(value);
            }
          };

          put("NWC", row.nwc);
          put("CustomerName", row.customerName);
          put("PC", row.profitCentre);

          // ✅ Write the Excel Description verbatim when present.
          //    (projectTitle was already computed above from Excel or stored title.)
          if (projectTitle !== "") {
            put("Title", projectTitle);
          }

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

          // JSON Merge Logic for Project List
          const existingProjJson = await fetchExistingMonthJson(PROJECT_LIST, resolvedProjectId, "ProjectMonth");
          const currentProjectData = {
            status: row.status || "",
            contractualDeliveryDate: row.contractualDeliveryDate || "",
            revisedContractualDeliveryDate: row.revisedContractualDeliveryDate || "",
            prQuality: row.prQuality || "",
            prSubmissionStatus: row.prSubmissionStatus || "",
            prApprovalStatus: row.prApprovalStatus || "",
            nwc: row.nwc ?? 0
          };
          existingProjJson[currentMonthKey] = currentProjectData;
          update["ProjectMonth"] = JSON.stringify(cleanOldMonths(existingProjJson));

          if (Object.keys(update).length) {
            const f = await fetchWithTimeout(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PROJECT_LIST)}')/items(${resolvedProjectId})`, {
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

        // =====================================================================
        // 2. UPDATE COLLECTION LIST (Existing/New -> Append ReceivableValueMonth JSON)
        // =====================================================================
        const collValues: Record<string, number | string> = {};
        const cMap: Array<[string, string]> = [
          ["NormalDue", "normalDue"], ["NormalOverdue", "normalOverdue"],
          ["RetentionDue", "retentionDue"], ["RetentionOverdue", "retentionOverdue"],
          ["GstTds", "gstTds"], ["WHT", "wht"], ["ReceivableValue", "receivableValue"]
        ];

        const currentReceivableData = {
          normalDue: row.normalDue || 0,
          normalOverdue: row.normalOverdue || 0,
          retentionDue: row.retentionDue || 0,
          retentionOverdue: row.retentionOverdue || 0,
          gstTds: row.gstTds || 0,
          wht: row.wht || 0,
          receivableValue: row.receivableValue || 0
        };

        cMap.forEach(([fieldKey, excelKey]) => {
          const v = row[excelKey];
          if (typeof v === "number" && v !== 0) collValues[fieldKey] = v;
        });

        const existingCollRes = await findChildRow(COLLECTION_LIST, resolvedProjectId, ref, projectTitle);
        let existingCollJson = existingCollRes.row
          ? await fetchExistingMonthJson(COLLECTION_LIST, existingCollRes.row.Id, "ReceivableValueMonth")
          : {};

        existingCollJson[currentMonthKey] = currentReceivableData;
        collValues["ReceivableValueMonth"] = JSON.stringify(cleanOldMonths(existingCollJson));

        const collRes = await upsertRow({
          listName: COLLECTION_LIST, projectId: resolvedProjectId, projectRef: ref, projectTitle,
          values: collValues, rowId: existingCollRes.row ? existingCollRes.row.Id : undefined
        });
        if (!collRes.success) notes.push("⚠ Collection List: " + describeError(collRes.error));

        // =====================================================================
        // 3. UPDATE REVENUE FORECAST LIST (Existing/New -> Append OBLMonth JSON)
        // =====================================================================
        if (typeof row.obl === "number" || typeof row.obl === "string") {
          const existingRevRes = await findChildRow(REVENUE_LIST, resolvedProjectId, ref, projectTitle);
          let existingRevJson = existingRevRes.row
            ? await fetchExistingMonthJson(REVENUE_LIST, existingRevRes.row.Id, "OBLMonth")
            : {};

          existingRevJson[currentMonthKey] = row.obl || 0;
          const revValues: Record<string, number | string> = {
            OBLMonth: JSON.stringify(cleanOldMonths(existingRevJson))
          };
          if (typeof row.obl === "number" && row.obl !== 0) revValues["OBL"] = row.obl;

          const revRes = await upsertRow({
            listName: REVENUE_LIST, projectId: resolvedProjectId, projectRef: ref, projectTitle,
            values: revValues, rowId: existingRevRes.row ? existingRevRes.row.Id : undefined
          });
          if (!revRes.success) notes.push("⚠ Revenue Forecast List: " + describeError(revRes.error));
        }

        // =====================================================================
        // 4. UPDATE OTHER SINGLE TABLES (Invoicing, SEIBIES, Inventory)
        // =====================================================================
        const singleTables: Array<{ listName: string; excelKey: string; jsonColName: string }> = [
          { listName: INVOICING_LIST, excelKey: "invoicingPlanning", jsonColName: "OBLMonth" },
          { listName: SEIBIES_LIST, excelKey: "seibies", jsonColName: "OBLMonth" },
          { listName: INVENTORY_LIST, excelKey: "inventory", jsonColName: "OBLMonth" }
        ];

        for (const { listName, excelKey, jsonColName } of singleTables) {
          const v = row[excelKey];
          const existingRowRes = await findChildRow(listName, resolvedProjectId, ref, projectTitle);

          let existingJson = existingRowRes.row
            ? await fetchExistingMonthJson(listName, existingRowRes.row.Id, jsonColName)
            : {};

          existingJson[currentMonthKey] = v || 0;

          const values: Record<string, number | string> = {
            [jsonColName]: JSON.stringify(cleanOldMonths(existingJson))
          };
          if (typeof v === "number" && v !== 0) values["OBL"] = v;

          const ur = await upsertRow({
            listName, projectId: resolvedProjectId, projectRef: ref, projectTitle,
            values, rowId: existingRowRes.row ? existingRowRes.row.Id : undefined
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
        const msg = describeError(err instanceof Error ? err.message : err);

        importErrorRows.push({
          rowNo: row.__rowNo ?? (i + 2),
          projectReference: ref,
          projectName: String(row.projectName ?? "").trim(),
          missing: [],
          message: msg,
          raw: row.__raw ?? []
        });

        addStatusRow("error", ref, String(row.projectName ?? ""), msg);
      }
    }

    if (popupProgress) popupProgress.style.display = "none";
    if (popupResults) popupResults.style.display = "block";

    if (downloadErrorsBtn) {
      if (importErrorRows.length > 0) {
        downloadErrorsBtn.style.display = "inline-flex";
        downloadErrorsBtn.textContent =
          `⬇️ Download Error List (${importErrorRows.length} row${importErrorRows.length === 1 ? "" : "s"})`;
      } else {
        downloadErrorsBtn.style.display = "none";
      }
    }

    showToast(`✅ Import complete — ${impCreated} new project(s), ${impUpdated} updated, ${impErrors} error(s) of ${total}`,
      impErrors ? "warning" : "success", 6000);
    await loadMaster();
  } catch (error) {
    console.error("Import error:", error);
    if (popupProgress) popupProgress.style.display = "none";
    if (popupResults) popupResults.style.display = "block";
    if (downloadErrorsBtn) downloadErrorsBtn.style.display = "none";
    if (popupStatusList) popupStatusList.innerHTML = "";
    addStatusRow("error", "(file)", file.name, describeError(error));
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
          const sample: Array<string | number> = [];
          const s: Record<string, any> = {
            profitCentre: "IN50334", projectReference: "REF-001", customerName: "Sample Customer",
            projectName: "Sample Customer Project — DELETE THIS ROW", groupHead: "KG",
            projectManager: "Kiran", enggLead: "Abhijeet",
            engSubLead1: "Yuvraj", engSubLead2: "Kiran", status: "REL",
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

        /* never let a mandatory column silently export all-blank */
        const mandatoryKeys = ["profitCentre", "projectReference", "customerName", "projectName", "groupHead"];
        const blankMandatory = mandatoryKeys.filter(k =>
          aoa.length > 1 && aoa.slice(1).every(row => {
            const idx = CSV_TEMPLATE.findIndex(c => c.key === k);
            return idx >= 0 && String(row[idx] ?? "").trim() === "";
          })
        );
        if (blankMandatory.length) {
          console.warn(`⚠ About to download with these mandatory columns blank on every row: ${blankMandatory.join(", ")}`);
          showToast(`⚠ Column(s) ${blankMandatory.join(", ")} are blank on every row — check the SharePoint field names before re-importing.`,
            "warning", 9000);
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
        el.value = "";
      });
      importChooseBtn?.addEventListener("click", () => importFileEl?.click());
      popupTemplateBtn?.addEventListener("click", downloadTemplate);
      popupCloseBtn?.addEventListener("click", closeImportPopup);
      popupDoneBtn?.addEventListener("click", closeImportPopup);
      popupOverlay?.addEventListener("click", (e: MouseEvent) => {
        if (!importing && e.target === popupOverlay) closeImportPopup();
      });
      downloadErrorsBtn?.addEventListener("click", downloadImportErrors);

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
        const oldSize = pageSize > 0 ? pageSize : Math.max(1, renderedData.length);
        const firstVisibleIndex = (currentPage - 1) * oldSize;
        pageSize = !v || v < 1 ? 0 : v;
        const effective = pageSize > 0 ? pageSize : Math.max(1, renderedData.length);
        const totalPages = Math.max(1, Math.ceil(renderedData.length / effective));
        currentPage = Math.min(totalPages, Math.floor(firstVisibleIndex / effective) + 1);
        renderTable();
      });

      editForm?.addEventListener("submit", onEditSubmit);
      backBtn?.addEventListener("click", () => {
        if (editSection) editSection.style.display = "none";
        if (tableSection) tableSection.style.display = "block";
        currentEditing = null;
        editForm?.reset();
      });

      /* ---------- boot ---------- */

      const lockTitlesWhite = (): void => {
        const targets: HTMLElement[] = [];
        const banner = document.querySelector(".master-dashboard .page-header") as HTMLElement | null;
        const bannerTitle = document.querySelector(".master-dashboard .page-header h1") as HTMLElement | null;
        const formTitle = document.querySelector(".master-page .form-header h2") as HTMLElement | null;
        if (bannerTitle) targets.push(bannerTitle);
        if (formTitle) targets.push(formTitle);
        if (banner) targets.push(banner);
        for (let i = 0; i < targets.length; i++) {
          try {
            targets[i].style.setProperty("color", "#ffffff", "important");
            targets[i].style.setProperty("-webkit-text-fill-color", "#ffffff", "important");
          } catch (e) {
            targets[i].style.color = "#ffffff";
          }
        }
      };

      const boot = async (): Promise<void> => {
        lockTitlesWhite();
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
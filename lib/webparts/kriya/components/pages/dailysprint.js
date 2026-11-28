import "../styles/common.css";
import "../styles/dailySprint.css";
import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css";
import { authManager, getCurrentUserFullName, getPermissionLevel, getUserGroups, getAssignedProjects, // ← scopes non-admins to assigned projects
canCreateRecord, // ← Admin / GH Head / PM Manager only
canReadAllProjects, // ← Admin can read all
getDigest as getAuthDigest, getListEntityType } from "./auth";
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
/** how many months the sprint view projects forward (current + 2) */
const HORIZON_MONTHS = 3;
let currentViewMode = 'revenue';
let activeMonthKey = '';
/** [currentMonthKey, nextMonthKey, monthAfterKey] */
let horizonKeys = [];
let pageMounted = false;
let dataLoaded = false;
let projects = [];
let revenueForecasts = [];
let collectionItems = [];
// ─── pagination state ───────────────────────────────
let currentPage = 1;
let pageSize = 25;
let totalPages = 1;
let filteredData = [];
// ─── sort state ─────────────────────────────────────
let sortColumn = '';
let sortDirection = 'asc';
/* ============================================================
   PURE HELPERS
   ============================================================ */
const toast = (icon, title) => {
    void Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
};
const getSafeString = (value, defaultValue = "-") => {
    if (value === null || value === undefined)
        return defaultValue;
    if (typeof value === 'string')
        return value.trim() || defaultValue;
    if (typeof value === 'object' && 'Title' in value) {
        return String(value.Title || '').trim() || defaultValue;
    }
    return String(value) || defaultValue;
};
const getPersonDisplayName = (value, fallback) => {
    if (!value)
        return fallback;
    if (typeof value === "object" && value !== null)
        return value.Title || fallback;
    return value;
};
const getGroupHeadDisplayName = (v) => getPersonDisplayName(v, "");
const getProjectManagerDisplayName = (v) => getPersonDisplayName(v, "");
const getEnggLeadDisplayName = (v) => getPersonDisplayName(v, "-");
const pad2 = (n) => String(n).padStart(2, "0");
const safeJsonParse = (jsonString, fallback = {}) => {
    if (!jsonString || !jsonString.trim() || jsonString.trim() === '""' || jsonString.trim() === "null") {
        return fallback;
    }
    try {
        const parsed = JSON.parse(jsonString.trim());
        return parsed && typeof parsed === 'object' ? parsed : fallback;
    }
    catch (error) {
        console.error("JSON parse error:", error, "Input:", jsonString);
        return fallback;
    }
};
const safeNumberMap = (jsonString) => {
    if (!jsonString || !jsonString.trim() || jsonString.trim() === '""' || jsonString.trim() === "null")
        return {};
    try {
        const parsed = JSON.parse(jsonString.trim());
        return parsed && typeof parsed === 'object' ? parsed : {};
    }
    catch (_a) {
        return {};
    }
};
const escapeAttr = (value) => value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
/** escape text going into element CONTENT (not into an attribute) */
const escapeHtml = (value) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const dataField = (mode) => mode === 'revenue' ? 'ForecastData' : 'CollectionData';
const SPRINT_CELL_FIELD = 'REMARKS_x0028_SPRINTCELL_x0029_';
const INTERNAL_REMARKS_FIELD = 'REMARKS_x0028_INTERNAL_x0029_';
const splitRemarks = (remarks) => {
    const raw = remarks || "";
    if (raw.includes('|||')) {
        const parts = raw.split('|||');
        return { internal: parts[0] || "", sprintCell: parts[1] || "" };
    }
    return { internal: raw, sprintCell: "" };
};
const inr = (n) => Math.round(Number(n) || 0).toLocaleString('en-US');
/**
 * "Concerns" text. Prefers the dedicated `REMARKS (SPRINT CELL)` column, and
 * falls back to the legacy `Remarks` field split on '|||' for older rows.
 */
const getConcerns = (item) => {
    const dedicated = item[SPRINT_CELL_FIELD];
    if (typeof dedicated === 'string' && dedicated.trim())
        return dedicated.trim();
    return splitRemarks(item.Remarks).sprintCell.trim();
};
/** Columns are "Single line of text" in SharePoint, so strip commas/spaces before Number(). */
const toNum = (v) => {
    if (typeof v === 'number')
        return Number.isFinite(v) ? v : 0;
    if (typeof v !== 'string')
        return 0;
    const n = Number(v.replace(/[,\s₹]/g, ''));
    return Number.isFinite(n) ? n : 0;
};
/** Receivables — display name "ReceivableValue", internal name `OBL`. */
const getReceivables = (item) => toNum(item.OBL);
/** "2026-07" → "2026-08" (rolls the year over correctly) */
const addMonths = (key, delta) => {
    const [ys, ms] = key.split('-');
    const y = parseInt(ys, 10);
    const m = parseInt(ms, 10) - 1 + delta;
    const ny = y + Math.floor(m / 12);
    const nm = ((m % 12) + 12) % 12;
    return `${ny}-${pad2(nm + 1)}`;
};
/** "2026-07" → "Jul'26" */
const shortMonthLabel = (key) => {
    const [y, m] = key.split('-');
    return `${MONTHS[parseInt(m, 10) - 1]}'${y.slice(2)}`;
};
const buildHorizon = (startKey) => Array.from({ length: HORIZON_MONTHS }, (_, i) => addMonths(startKey, i));
const siteUrl = typeof _spPageContextInfo !== "undefined"
    ? _spPageContextInfo.webAbsoluteUrl
    : window.location.origin + "/sites/Kriya";
const revenueListName = "Revenue Forecast List";
const collectionListName = "Collection List";
const projectListName = "Project List";
const LEFT_COLUMNS = {
    revenue: [
        { key: 'projectRef', label: 'Project reference', width: 120, sort: 'projectRef' },
        { key: 'customer', label: 'Customer Name', width: 150, sort: 'customer' },
        { key: 'project', label: 'Project Name', width: 180, sort: 'project' },
        { key: 'gh', label: 'Group Head', width: 90, align: 'center', sort: 'groupHead' },
        { key: 'pm', label: 'Project Manager', width: 90, align: 'center', sort: 'projectManager' },
        { key: 'enggLead', label: 'Engineering Lead', width: 110, align: 'center', sort: 'enggLead' },
        { key: 'concerns', label: 'Add Concerns', width: 180 },
        { key: 'obl', label: 'OBL', width: 110, align: 'right', sort: 'obl' }
    ],
    collection: [
        { key: 'projectRef', label: 'Project reference', width: 120, sort: 'projectRef' },
        { key: 'customer', label: 'Customer Name', width: 150, sort: 'customer' },
        { key: 'project', label: 'Project', width: 180, sort: 'project' },
        { key: 'gh', label: 'Group Head', width: 90, align: 'center', sort: 'groupHead' },
        { key: 'pm', label: 'Project Manager', width: 90, align: 'center', sort: 'projectManager' },
        { key: 'enggLead', label: 'Engineering Lead', width: 110, align: 'center', sort: 'enggLead' },
        { key: 'concerns', label: 'Concerns', width: 180 },
        { key: 'receivables', label: 'Receivables', width: 130, align: 'right', sort: 'receivables' }
    ]
};
const leftCols = () => LEFT_COLUMNS[currentViewMode];
/** cumulative left offset for sticky positioning */
const leftOffset = (index) => leftCols().slice(0, index).reduce((sum, c) => sum + c.width, 0);
/** total columns = left block + 5 (current month) + 2 per future month */
const totalCols = () => leftCols().length + 5 + 2 * (HORIZON_MONTHS - 1);
/* ============================================================
   CONTENT
   ============================================================ */
const content = `
<div class="daily-sprint-page">

  <div class="daily-sprint-header">
    <div class="daily-sprint-title">
      <h1>Daily Sprint</h1>
      <span id="activeViewIndicator" class="daily-sprint-view-indicator">📊 Revenue View</span>
    </div>

    <div class="daily-sprint-header-actions">
      <div class="daily-sprint-view-toggle">
        <button id="toggleRevenue" type="button" class="daily-sprint-toggle-btn toggle-btn-active">Revenue</button>
        <button id="toggleCollection" type="button" class="daily-sprint-toggle-btn toggle-btn-inactive">Collection</button>
      </div>
    </div>
  </div>

  <div class="daily-sprint-table-section">

    <!-- ✅ THE FILTER BUTTONS NOW SIT INSIDE THE "DAILY SPRINT TOTALS" BAND
         (renderTotalsHeader + mountTotalsToolbar in the script below).
         The funnel trigger, its active-count badge and the compact "Apply"
         pill (shown only while the popup holds unapplied changes) are mounted
         at the LEFT of that band - i.e. directly above the "All Projects"
         cell - so nothing floats above the table any more.
         The eight dropdowns live INSIDE the Filters popup below, keeping their
         original ids so filterTable() and the exports keep working unchanged.
         There is no Reset button in the strip - the popup owns Reset / Clear,
         exactly like the other pages.

         This holder is NOT displayed: it only keeps the two buttons in the DOM
         so the script can find them on the first run. mountTotalsToolbar()
         MOVES them into the band (appendChild moves, it never clones), which
         is what keeps their click listeners, the badge count and the Apply
         pill's hidden / disabled state alive across every re-render. -->
    <div class="sprint-toolbar-holder" hidden>
      <div class="sprint-table-toolbar">
        <div class="sprint-toolbar-left">
          <button type="button" class="sprint-filter-trigger" id="sprintFilterTrigger" aria-haspopup="dialog" aria-expanded="false" title="Filter the daily sprint grid">
            <svg class="sprint-filter-funnel" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8v5l-4 2v-7L3 5z"/></svg><span>Filters</span><span class="sprint-filter-count" id="sprintFilterCount">0</span><svg class="sprint-filter-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>
          </button>
          <button type="button" class="sprint-filter-applypill" id="sprintFilterApplyPill" title="Apply the selected filters">Apply</button>
        </div>
      </div>
    </div>

    <!-- applied filters are listed here as removable chips -->
   <!-- <div class="sprint-filter-summary" id="sprintFilterChips" hidden></div> -->

    <div class="sprint-filter-overlay" id="sprintFilterOverlay" role="dialog" aria-modal="true" aria-hidden="true" aria-labelledby="sprintFilterTitle">
      <div class="sprint-filter-modal">
        <div class="sprint-filter-modal-head">
          <h2 id="sprintFilterTitle">Filters</h2>
          <button type="button" id="sprintFilterClose" aria-label="Close">&times;</button>
        </div>

        <div class="sprint-filter-grid">
          <label class="daily-sprint-filter">Year
            <select id="sprintFilterYear" class="sprint-filter-select"><option value="">All Years</option></select>
          </label>
          <label class="daily-sprint-filter">Month
            <select id="sprintFilterMonth" class="sprint-filter-select"><option value="">All Months</option></select>
          </label>
          <label class="daily-sprint-filter">Project
            <select id="sprintFilterProject" class="sprint-filter-select"><option value="">All Projects</option></select>
          </label>
          <label class="daily-sprint-filter">Customer
            <select id="sprintFilterCustomer" class="sprint-filter-select"><option value="">All Customers</option></select>
          </label>
          <label class="daily-sprint-filter">Project Manager
            <select id="sprintFilterProjectManager" class="sprint-filter-select"><option value="">All Managers</option></select>
          </label>
          <label class="daily-sprint-filter">Group Head
            <select id="sprintFilterGroupHead" class="sprint-filter-select"><option value="">All Group Heads</option></select>
          </label>
          <label class="daily-sprint-filter">Project Reference
            <select id="sprintFilterProjectReference" class="sprint-filter-select"><option value="">All References</option></select>
          </label>
          <label class="daily-sprint-filter">Engg Lead
            <select id="sprintFilterEnggLead" class="sprint-filter-select"><option value="">All Engg Leads</option></select>
          </label>
          <div class="sprint-filter-reset-cell">
            <button type="button" class="daily-sprint-clear-filters-btn" id="sprintResetFilterBtn" title="Clear all filters" disabled>Reset Filter</button>
          </div>
        </div>

        <div class="sprint-filter-modal-foot">
          <button type="button" class="sprint-btn-secondary" id="sprintFilterCancel">Cancel</button>
          <button type="button" class="sprint-btn-secondary sprint-filter-btn-clear" id="sprintFilterClear" title="Clear all filters" disabled>Clear</button>
          <button type="button" class="sprint-btn-primary sprint-filter-btn-apply" id="sprintFilterApply" disabled>Apply</button>
        </div>
      </div>
    </div>

    <!-- 🔐 AUTO-SYNC BAR (only visible to roles with create rights: Admin / GH Head / PM Manager).
         Mirrors the Revenue and Collection page sync bars - console-only reporting. -->
    <div id="sprintSyncBar"
         style="display:none; margin:6px 0 10px 0; align-items:center; gap:12px; flex-wrap:wrap;">
      <button type="button" id="sprintSyncProjectsBtn"
              style="padding:6px 14px; border:1px solid #0078d4; background:#ffffff; color:#0078d4;
                     border-radius:4px; cursor:pointer; font-size:13px; font-weight:600;"
              title="Compare Project List with the active forecast list and add any missing project row">➕ Sync Projects</button>
      <span id="sprintSyncHint" style="font-size:12px; color:#6c757d;"></span>
    </div>

    <!-- ✏️ Concerns cells are editable (shown only to users who may edit).
         Text is MERGEd to the underlying list the moment you leave the cell.
         The hint strip itself is switched off in the markup below - the cells
         still save exactly as before (the script finds no hint node and skips
         the line that would have shown it). -->
    <!--<div id="sprintConcernsHint" class="sprint-concerns-hint" style="display:none;">
      💬 <b>Concerns</b> cells are editable — click a cell, type, then press
      <kbd>Enter</kbd> or click away to save straight to the list.
      <kbd>Shift</kbd>+<kbd>Enter</kbd> adds a line, <kbd>Esc</kbd> cancels.
    </div> -->

    <div class="sprint-data-container">
        <table class="daily-sprint-table">
          <thead id="sprintTotalsHeader"></thead>
          <thead id="sprintMonthColumnsHeader"></thead>
          <tbody class="sprint-table-body">
            <tr><td colspan="20" class="daily-sprint-empty">Loading data…</td></tr>
          </tbody>
        </table>

      <!-- ✅ PAGINATION - one flat bar, exact mock text + layout:
           LEFT  : Items per page: [n]   Showing a - b of c items
           RIGHT : « ‹ [tiles] of N › »                                            -->
      <div class="sprint-pagination-container">
        <div class="sprint-pagination-left">
          <div class="sprint-pagination-size-selector">
            <label for="sprintPageSize">Items per page:</label>
            <select id="sprintPageSize" class="sprint-page-size-select">
              <option value="10">10</option>
              <option value="25" selected>25</option>
              <option value="50">50</option>
              <option value="100">100</option>
            </select>
          </div>

          <div class="sprint-pagination-info sprint-pg-total">
            Showing <span class="sprint-pg-range"><span id="sprintStartItem">0</span> - <span id="sprintEndItem">0</span></span> of <span class="sprint-pg-count" id="sprintTotalItems">0</span> items
          </div>
        </div>

        <div class="sprint-pagination-right">
          <button class="sprint-pagination-btn sprint-pg-icon" id="sprintFirstPage" disabled title="First page">«</button>
          <button class="sprint-pagination-btn sprint-pg-icon" id="sprintPrevPage" disabled title="Previous page">‹</button>
          <div class="sprint-pagination-pages" id="sprintPageNumbers"></div>
          <span class="sprint-pg-of">of <span id="sprintTotalPages">1</span></span>
          <button class="sprint-pagination-btn sprint-pg-icon" id="sprintNextPage" disabled title="Next page">›</button>
          <button class="sprint-pagination-btn sprint-pg-icon" id="sprintLastPage" disabled title="Last page">»</button>
        </div>
      </div>
    </div>
  </div>

</div>
`;
/* ============================================================
   STYLE — sticky thead rows + sticky left columns
   ============================================================ */
const style = `
  /* ---- Concerns cell: inline editor ------------------------------------
     (kept here so the module is self-contained; safe to move verbatim into
      ../styles/dailySprint.css if you prefer all styling in one place) ---- */

  .sprint-concerns-hint{
    margin:6px 0 10px 0; padding:6px 10px; font-size:12px; color:#444;
    background:#f6f8fa; border:1px solid #dde3e8; border-radius:4px;
  }
  .sprint-concerns-hint kbd{
    font:600 11px/1.4 inherit; background:#fff; border:1px solid #c8d0d8;
    border-bottom-width:2px; border-radius:3px; padding:0 4px; color:#333;
  }

  /* wrapper (NOT the td — the td needs position:sticky for the left block) */
  .sprint-concerns-wrap{position:relative;display:flex;align-items:flex-start;gap:2px;}

  .sprint-concerns-editor{
    flex:1 1 auto;
    min-height:20px; max-height:64px; overflow-y:auto;
    padding:2px 4px; margin:-2px -4px;           /* keep text aligned with siblings */
    white-space:pre-wrap; word-break:break-word;
    border-radius:3px; outline:none; cursor:text;
    transition:background .12s ease, box-shadow .12s ease;
  }
  .sprint-concerns-cell.is-editable .sprint-concerns-editor:hover{
    background:#f3f6fb; box-shadow:inset 0 0 0 1px #cfe0f5;
  }
  .sprint-concerns-cell.is-editing .sprint-concerns-editor{
    background:#fff; box-shadow:inset 0 0 0 2px #0078d4;
  }
  /* placeholder via CSS so it never becomes real content (and never saves) */
  .sprint-concerns-editor.is-empty::before{
    content:attr(data-placeholder); color:#9aa5b1; font-style:italic;
  }

  .sprint-concerns-status{
    flex:0 0 auto; width:14px; font-size:11px; line-height:20px; text-align:center;
  }
  .sprint-concerns-cell.is-error .sprint-concerns-editor{
    box-shadow:inset 0 0 0 2px #d13438;
  }
  .sprint-concerns-cell.is-saving .sprint-concerns-editor{opacity:.65;}

  /* ---- page loader: bar only (grey track #DCDCDC + indigo bar #6664FF),
          no spinner, no text - messages go to the console ---- */
  .sprint-loading-overlay{
    position:absolute; inset:0; display:flex; align-items:center; justify-content:center;
    background:rgba(248,249,250,.92); z-index:999;
  }
  .sprint-loader-card{
    width:min(520px, 84%); padding:22px 26px; background:#fff;
    border:1px solid #e6e9ec; border-radius:12px;
    box-shadow:0 10px 26px rgba(15,45,85,.12);
  }
  .sprint-page-loader{ display:block; width:100%; }
  .sprint-page-loader__track{
    position:relative; width:100%; height:6px; overflow:hidden;
    background:#DCDCDC; border-radius:999px;
  }
  .sprint-page-loader__bar{
    position:absolute; top:0; left:0; width:50%; height:100%;
    background:#6664FF; border-radius:999px;
    animation:sprintPageLoaderSlide 1.25s cubic-bezier(.65,.05,.36,1) infinite;
  }
  @keyframes sprintPageLoaderSlide{
    0%{transform:translateX(-100%)}
    100%{transform:translateX(200%)}
  }
`;
/* ============================================================
   DATA LOADING
   ============================================================ */
const loadProjects = async () => {
    var _a;
    try {
        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items` +
            `?$select=Id,Title,CustomerName,PC,ProjectReference,` +
            `GroupHead1/Title,GroupHead1Id,` +
            `ProjectManager/Title,ProjectManagerId,` +
            `EnggLead/Title,EnggLeadId,` +
            `EnggSubLead1/Title,EnggSubLead1Id,` +
            `EnggSubLead2/Title,EnggSubLead2Id,` +
            `EngineeringManagers/Title,EngineeringManagersId` +
            `&$expand=GroupHead1,ProjectManager,EnggLead,EnggSubLead1,EnggSubLead2,EngineeringManagers` +
            `&$top=5000`, { headers: { Accept: "application/json;odata=verbose" } });
        const data = await res.json();
        const raw = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
        const byTitle = new Map();
        raw.forEach(p => {
            if (p && p.Title && !byTitle.has(p.Title))
                byTitle.set(p.Title, p);
        });
        projects = Array.from(byTitle.values());
        console.log('📋 Projects loaded:', projects.length);
    }
    catch (error) {
        console.error("Error loading projects:", error);
        toast("error", "Failed to load projects");
    }
};
const normalizeRecords = (rows, jsonFields) => rows.map(row => {
    const out = Object.assign({}, row);
    jsonFields.forEach(field => {
        out[field] = typeof row[field] === 'string' ? row[field] : JSON.stringify(row[field] || {});
    });
    out.Freeze = row.Freeze || false;
    return out;
});
/**
 * `PreviousPlanData` is an OPTIONAL column. If a list doesn't have it, SharePoint
 * rejects the whole $select with 400 and we'd load zero rows. So we try with it
 * once, and transparently fall back to a $select without it.
 */
/**
 * Fields that may not exist on every list. Rather than guessing, we let
 * SharePoint tell us: its 400 body names the offending column
 * ("The field or property 'X' does not exist."), so we drop exactly that one
 * and retry. This self-heals for ANY missing column, not just a known list.
 */
const OPTIONAL_FIELDS = ['PreviousPlanData', SPRINT_CELL_FIELD, INTERNAL_REMARKS_FIELD];
/** per-list memo of columns SharePoint has rejected, so we only pay the retry once */
const unsupportedFields = {};
/** pull the bad field name out of an SP error payload */
const parseMissingField = (body) => {
    const m = body.match(/field or property '([^']+)' does not exist/i);
    return m ? m[1] : null;
};
/** remember that SharePoint rejected a column so we stop trying to use it */
const markUnsupportedField = (listName, field) => {
    if (!unsupportedFields[listName])
        unsupportedFields[listName] = new Set();
    unsupportedFields[listName].add(field);
};
/** has SharePoint accepted this column for this list? (optimistic until told no) */
const fieldSupported = (listName, field) => { var _a; return !((_a = unsupportedFields[listName]) === null || _a === void 0 ? void 0 : _a.has(field)); };
const fetchListItems = async (listName, baseSelect, extra = '') => {
    var _a;
    const bad = unsupportedFields[listName] || (unsupportedFields[listName] = new Set());
    const build = (fields) => `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items` +
        `?$select=${fields.join(',')}&$expand=Project&$top=5000${extra}` +
        `&_=${Date.now()}`; // 🔧 cache-buster: always read live data from the database
    const attempt = async (fields) => fetch(build(fields), { headers: { Accept: "application/json;odata=verbose" } });
    let fields = [...baseSelect, ...OPTIONAL_FIELDS].filter(f => !bad.has(f));
    let res = await attempt(fields);
    // Drop rejected columns one at a time until the query is accepted.
    for (let i = 0; i < OPTIONAL_FIELDS.length + 2 && !res.ok; i++) {
        const body = await res.text();
        const missing = parseMissingField(body);
        if (!missing) {
            throw new Error(`HTTP ${res.status} loading "${listName}": ${body}`);
        }
        const match = fields.find(f => f === missing || f.split('/')[0] === missing);
        if (!match) {
            throw new Error(`HTTP ${res.status} loading "${listName}": column '${missing}' is missing ` +
                `but is required by this page.`);
        }
        console.warn(`⚠️ "${listName}": column '${match}' does not exist — dropping it and retrying. ` +
            `Check List Settings for its INTERNAL name (the Field= value in the column's edit URL).`);
        bad.add(match);
        fields = fields.filter(f => f !== match);
        res = await attempt(fields);
    }
    if (!res.ok) {
        throw new Error(`HTTP ${res.status} loading "${listName}": ${await res.text()}`);
    }
    const data = await res.json();
    return ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
};
const loadRevenueForecasts = async () => {
    try {
        const rows = await fetchListItems(revenueListName, [
            'Id', 'Title', 'Project/Id', 'Project/Title', 'ProfitCentre', 'OBL',
            'ForecastData', 'RevenueActual', 'Remarks', 'Created', 'Modified', 'Freeze'
        ]);
        revenueForecasts = normalizeRecords(rows, ['ForecastData', 'RevenueActual', 'PreviousPlanData']);
        console.log("✅ Revenue forecasts loaded:", revenueForecasts.length);
    }
    catch (error) {
        console.error("Error loading revenue forecasts:", error);
        toast("error", "Failed to load revenue data");
    }
};
const loadCollectionItems = async () => {
    try {
        const rows = await fetchListItems(collectionListName, [
            'Id', 'Title', 'Project/Id', 'Project/Title', 'ProfitCentre', 'OBL',
            'CollectionData', 'RevenueActual', 'Remarks', 'Created', 'Modified', 'Freeze',
            // 'OBL' is the internal name of the column displayed as "ReceivableValue"
            'NormalDue', 'NormalOverdue', 'RetentionDue', 'RetentionOverdue', 'GSTTDS', 'WHT'
        ]);
        collectionItems = normalizeRecords(rows, ['CollectionData', 'RevenueActual', 'PreviousPlanData']);
        console.log("✅ Collection items loaded:", collectionItems.length);
    }
    catch (error) {
        console.error("Error loading collection items:", error);
        toast("error", "Failed to load collection data");
    }
};
const currentRecords = () => currentViewMode === 'revenue' ? revenueForecasts : collectionItems;
/* ============================================================
   CONCERNS CELL — INLINE EDIT, SAVED STRAIGHT TO SHAREPOINT
   ------------------------------------------------------------
   The Concerns cell is typed straight into the grid and MERGEd to the list
   behind the active view (Revenue Forecast List / Collection List) the moment
   the editor loses focus or Enter is pressed. No "Save" button, no modal.

   Where it writes:
     1. the dedicated `REMARKS (SPRINT CELL)` column
        (internal name `REMARKS_x0028_SPRINTCELL_x0029_`) when the list has it;
     2. otherwise the legacy `Remarks` column, re-packed as
        "<internal remarks>|||<sprint cell text>" so the internal half of the
        old combined field is never lost.
   Either way `getConcerns()` keeps reading the new value back, because it
   already prefers the dedicated column and falls back to the split.
   ============================================================ */
/** which list the active view is editing */
const currentListName = () => currentViewMode === 'revenue' ? revenueListName : collectionListName;
/**
 * 🔐 Who is allowed to type into a Concerns cell.
 * Default is the same create-rights gate as the sync bar
 * (canCreateRecord → Admin FULL_CONTROL / GH Heads / PM Managers).
 * Change to `() => true` if every viewer should be able to log a concern.
 */
const canEditConcerns = () => canCreateRecord();
/** cache list entity types — they never change at runtime */
const entityTypeCache = {};
const entityTypeFor = async (listName) => {
    if (!entityTypeCache[listName]) {
        entityTypeCache[listName] = await getListEntityType(siteUrl, listName);
    }
    return entityTypeCache[listName];
};
/** single-item MERGE (partial update — untouched columns are left alone) */
const mergeListItem = async (listName, itemId, payload) => {
    const digest = await getAuthDigest(siteUrl);
    return fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`, {
        method: 'POST',
        headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': digest,
            'IF-Match': '*', // last write wins; the cell is a one-liner
            'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
    });
};
/** unwrap a SharePoint/REST error into something a human can read */
const readErrorText = async (res) => {
    var _a, _b;
    const body = await res.text();
    try {
        const parsed = JSON.parse(body);
        return ((_b = (_a = parsed === null || parsed === void 0 ? void 0 : parsed.error) === null || _a === void 0 ? void 0 : _a.message) === null || _b === void 0 ? void 0 : _b.value) || body || `HTTP ${res.status}`;
    }
    catch (_c) {
        return body || `HTTP ${res.status}`;
    }
};
/**
 * Write one Concerns value to SharePoint.
 * Falls back from the dedicated column to `Remarks` if SharePoint rejects the
 * column (list without it, renamed internal name, read-only, …).
 */
const saveConcernsToSharePoint = async (listName, record, text) => {
    try {
        const type = await entityTypeFor(listName);
        const legacyRemarks = () => `${splitRemarks(record.Remarks).internal}|||${text}`;
        if (fieldSupported(listName, SPRINT_CELL_FIELD)) {
            const res = await mergeListItem(listName, record.Id, {
                __metadata: { type },
                [SPRINT_CELL_FIELD]: text
            });
            if (res.ok)
                return { ok: true, field: 'sprintCell', message: '' };
            const message = await readErrorText(res);
            const missing = parseMissingField(message);
            // only swallow the "column does not exist" case — a real failure (403,
            // locked row, …) must surface to the user instead of silently retrying
            if (!missing || missing.split('/')[0] !== SPRINT_CELL_FIELD) {
                return { ok: false, field: null, message };
            }
            markUnsupportedField(listName, SPRINT_CELL_FIELD);
            console.warn(`⚠️ "${listName}" has no '${SPRINT_CELL_FIELD}' column — Concerns will be ` +
                `written to the legacy Remarks column instead ` +
                `("internal|||sprint cell"). Check List Settings for the real INTERNAL name.`);
        }
        const fallback = await mergeListItem(listName, record.Id, {
            __metadata: { type },
            Remarks: legacyRemarks()
        });
        if (fallback.ok)
            return { ok: true, field: 'remarks', message: '' };
        return { ok: false, field: null, message: await readErrorText(fallback) };
    }
    catch (error) {
        return { ok: false, field: null, message: error instanceof Error ? error.message : String(error) };
    }
};
/** mirror a saved value onto the in-memory row so sorts/pages keep it */
const applyConcernsLocally = (record, text, field) => {
    if (field === 'sprintCell') {
        record[SPRINT_CELL_FIELD] = text;
    }
    else {
        record.Remarks = `${splitRemarks(record.Remarks).internal}|||${text}`;
    }
};
const syncMonthKey = (date) => `${date.getFullYear()}-${pad2(date.getMonth() + 1)}`;
/** Revenue skeleton - same month grid as the Revenue page:
 *  36 historical + 3 detailed (current + next 2) + 15 future months. */
const syncIdealForecastMonths = () => {
    const months = [];
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth();
    for (let i = -36; i < 0; i++)
        months.push(syncMonthKey(new Date(year, month + i, 1)));
    for (let i = 0; i < 3; i++)
        months.push(syncMonthKey(new Date(year, month + i, 1)));
    for (let i = 3; i < 18; i++)
        months.push(syncMonthKey(new Date(year, month + i, 1)));
    return Array.from(new Set(months)).sort();
};
const syncEmptyForecastItems = () => {
    const items = {};
    for (let i = 1; i <= 10; i++) {
        items[i.toString()] = { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
    }
    return items;
};
const syncFreshYearlyForecast = (lastMonthKey) => {
    const yearlyData = {};
    let startYear;
    if (lastMonthKey) {
        startYear = parseInt(lastMonthKey.split('-')[0], 10);
    }
    else {
        const futureDate = new Date();
        futureDate.setMonth(futureDate.getMonth() + 15);
        startYear = futureDate.getFullYear();
    }
    const endYear = startYear + 4;
    for (let year = startYear; year <= endYear; year++) {
        yearlyData[year.toString()] = { plan: 0, actual: 0 };
    }
    return yearlyData;
};
/** Empty ForecastData JSON - byte-identical shape to the Revenue page's
 *  createFreshForecastStructure(). */
const syncFreshForecastStructure = () => {
    const months = syncIdealForecastMonths();
    const freshData = {};
    const currentMonthKey = syncMonthKey(new Date());
    const currentMonthIndex = months.indexOf(currentMonthKey);
    const detailedStartIndex = currentMonthIndex !== -1 ? currentMonthIndex : 36;
    months.forEach((monthKey, index) => {
        const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
        freshData[monthKey] = {
            plan: 0,
            actual: 0,
            forecast: 0,
            monthActual: 0,
            items: isDetailedMonth ? syncEmptyForecastItems() : {}
        };
    });
    freshData.yearly = syncFreshYearlyForecast(months[months.length - 1]);
    return freshData;
};
/** Collection skeleton - same month grid as the Collection page:
 *  3 + 33 historical, 3 detailed (current + next 2), 16 future months. */
const syncIdealCollectionMonths = () => {
    const months = [];
    const now = new Date();
    for (let i = -3; i < 0; i++) {
        const d = new Date(now);
        d.setMonth(d.getMonth() + i);
        months.push(syncMonthKey(d));
    }
    for (let i = -36; i < -3; i++) {
        const d = new Date(now);
        d.setMonth(d.getMonth() + i);
        months.push(syncMonthKey(d));
    }
    for (let i = 0; i < 3; i++) {
        const d = new Date(now);
        d.setMonth(d.getMonth() + i);
        months.push(syncMonthKey(d));
    }
    for (let i = 3; i < 19; i++) {
        const d = new Date(now);
        d.setMonth(d.getMonth() + i);
        months.push(syncMonthKey(d));
    }
    return months.sort();
};
const syncCollectionCategory = (name) => ({ name, plan: 0, actual: 0, forecast: 0, date: '' });
const syncCollectionCategories = (simple) => {
    if (simple) {
        return { Normal: syncCollectionCategory('Normal'), Retention: syncCollectionCategory('Retention') };
    }
    return {
        'Normal': syncCollectionCategory('Normal'),
        'Normal Overdue': syncCollectionCategory('Normal Overdue'),
        'Retention': syncCollectionCategory('Retention'),
        'Retention Overdue': syncCollectionCategory('Retention Overdue')
    };
};
const syncFreshYearlyCollection = (lastMonthKey) => {
    const yearlyData = {};
    let startYear;
    if (lastMonthKey) {
        startYear = parseInt(lastMonthKey.split('-')[0], 10);
    }
    else {
        startYear = new Date().getFullYear() + 1;
    }
    const endYear = startYear + 3;
    for (let year = startYear; year <= endYear; year++) {
        yearlyData[year.toString()] = {
            plan: 0,
            actual: 0,
            items: {
                Normal: { name: 'Normal', plan: 0, actual: 0 },
                Retention: { name: 'Retention', plan: 0, actual: 0 }
            }
        };
    }
    return yearlyData;
};
/** Empty CollectionData JSON - byte-identical shape to the Collection page's
 *  createFreshCollectionStructure(). */
const syncFreshCollectionStructure = () => {
    const months = syncIdealCollectionMonths();
    const freshData = {};
    const currentMonthKey = syncMonthKey(new Date());
    const currentMonthIndex = months.indexOf(currentMonthKey);
    const detailedStartIndex = currentMonthIndex !== -1 ? currentMonthIndex : 36;
    months.forEach((monthKey, index) => {
        const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
        let items = {};
        if (isDetailedMonth) {
            items = syncCollectionCategories(false);
        }
        else if (index >= detailedStartIndex + 3) {
            items = syncCollectionCategories(true);
        }
        freshData[monthKey] = {
            plan: 0,
            actual: 0,
            monthActual: 0,
            forecast: 0,
            date: '',
            items
        };
    });
    freshData.yearly = syncFreshYearlyCollection(months[months.length - 1]);
    return freshData;
};
/** Title format used by the Revenue page Add-form:
 *  "<Project> - <Year> Forecast - <sum of monthly actuals>" */
const syncForecastTitle = (project, data) => {
    const currentYear = new Date().getFullYear();
    let totalValue = 0;
    Object.keys(data).forEach(key => {
        if (key.match(/^\d{4}-\d{2}$/)) {
            const monthData = data[key];
            if (monthData)
                totalValue += Number(monthData.actual) || 0;
        }
    });
    return `${project.Title} - ${currentYear} Forecast - ${totalValue.toLocaleString()}`;
};
/** Title format used by the Collection page Add-form:
 *  "<Project> - <Year> Collection - <sum of monthly actuals>" */
const syncCollectionTitle = (project, data) => {
    const currentYear = new Date().getFullYear();
    let totalValue = 0;
    Object.keys(data).forEach(key => {
        if (key.match(/^\d{4}-\d{2}$/)) {
            const monthData = data[key];
            if (monthData)
                totalValue += Number(monthData.actual) || 0;
        }
    });
    return `${project.Title} - ${currentYear} Collection - ${totalValue.toLocaleString()}`;
};
/** Cached check: does Revenue Forecast List have a 'ProjectReference'
 *  column, and is it plain text or a Lookup to Project List? */
let sprintRevenueRefField = null;
const syncRevenueRefField = async () => {
    var _a;
    if (sprintRevenueRefField)
        return sprintRevenueRefField;
    try {
        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${revenueListName}')/fields?$filter=InternalName eq 'ProjectReference'&$select=InternalName,TypeAsString,ReadOnlyField`, { headers: { Accept: "application/json;odata=verbose" } });
        const data = await res.json();
        const results = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
        if (results.length > 0) {
            sprintRevenueRefField = { exists: true, type: String(results[0].TypeAsString || "Text") };
            console.log(`📋 Revenue Forecast List ProjectReference column type = ${sprintRevenueRefField.type}`);
        }
        else {
            sprintRevenueRefField = { exists: false, type: "" };
            console.log("📋 Revenue Forecast List has no ProjectReference column");
        }
    }
    catch (error) {
        console.error("⚠️ Could not check Revenue ProjectReference column:", error);
        sprintRevenueRefField = { exists: false, type: "" };
    }
    return sprintRevenueRefField;
};
/** Cached check: does Collection List have a 'ProjectReference' column,
 *  and is it plain text or a Lookup to Project List? */
let sprintCollectionRefField = null;
const syncCollectionRefField = async () => {
    var _a;
    if (sprintCollectionRefField)
        return sprintCollectionRefField;
    try {
        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${collectionListName}')/fields?$filter=InternalName eq 'ProjectReference'&$select=InternalName,TypeAsString,ReadOnlyField`, { headers: { Accept: "application/json;odata=verbose" } });
        const data = await res.json();
        const results = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
        if (results.length > 0) {
            sprintCollectionRefField = { exists: true, type: String(results[0].TypeAsString || "Text") };
            console.log(`📋 Collection List ProjectReference column type = ${sprintCollectionRefField.type}`);
        }
        else {
            sprintCollectionRefField = { exists: false, type: "" };
            console.log("📋 Collection List has no ProjectReference column");
        }
    }
    catch (error) {
        console.error("⚠️ Could not check Collection ProjectReference column:", error);
        sprintCollectionRefField = { exists: false, type: "" };
    }
    return sprintCollectionRefField;
};
const syncMissingRows = async (cfg) => {
    var _a, _b, _c, _d;
    if (!canCreateRecord()) {
        console.log('👁️ No create rights - sync skipped (Admin/GH Head/PM Manager only)');
        return;
    }
    if (!projects || projects.length === 0) {
        console.log('⚠️ Project List is empty - nothing to sync');
        return;
    }
    try {
        const visibleProjects = canReadAllProjects()
            ? projects
            : getAssignedProjects(projects);
        // Safety re-check against SharePoint itself: the in-memory list may be
        // empty/stale, so never create rows for projects that already exist.
        let serverRows = [];
        try {
            const chkRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${cfg.listName}')/items?$select=Id,ProjectId,Title&$top=5000`, { headers: { Accept: "application/json;odata=verbose" } });
            if (chkRes.ok)
                serverRows = ((_b = (_a = (await chkRes.json())) === null || _a === void 0 ? void 0 : _a.d) === null || _b === void 0 ? void 0 : _b.results) || [];
        }
        catch (chkErr) {
            console.error(`⚠️ ${cfg.summaryLabel} safety-check fetch failed:`, chkErr);
        }
        // If the list really has rows on SharePoint but the page loaded 0, the
        // main load is broken - creating rows now would duplicate everything.
        if (serverRows.length > 0 && cfg.records.length === 0) {
            console.warn(`⚠️ ${cfg.listName} already has ${serverRows.length} row(s) on SharePoint but the page loaded 0 - ` +
                `sync stopped to avoid duplicates. Check the main list load.`);
            return;
        }
        const existsOnServer = (p) => serverRows.some(c => Number(c.ProjectId) === p.Id);
        const existsInMemory = (p) => cfg.records.some(r => { var _a; return ((_a = r.Project) === null || _a === void 0 ? void 0 : _a.Id) === p.Id || (r.Project && r.Project.Title === p.Title); });
        const missingProjects = visibleProjects.filter(p => !existsOnServer(p) && !existsInMemory(p));
        if (missingProjects.length === 0) {
            console.log(`✅ In sync: all ${visibleProjects.length} visible project(s) have a ${cfg.listName} row`);
            return;
        }
        console.log(`🆕 ${missingProjects.length} missing: ${missingProjects.map(p => p.Title).join(', ')}`);
        const digest = await getAuthDigest(siteUrl);
        const entityType = await getListEntityType(siteUrl, cfg.listName);
        let added = 0;
        let failed = 0;
        const errors = [];
        for (const project of missingProjects) {
            let ok = false;
            // Up to 3 payload variants (same as the Revenue/Collection pages):
            //   1 = full create - ProjectReference lookup id as STRING + OBL as STRING
            //   2 = without ProjectReference (column may be read-only/misconfigured)
            //   3 = without OBL too (column may be read-only/absent)
            for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
                try {
                    const freshData = cfg.skeleton();
                    // odata=verbose values must be STRINGs (Edm.String): lookup ids and
                    // numeric fields sent as JSON numbers get a 400 "Cannot convert a
                    // primitive value to the expected type 'Edm.String'".
                    const payload = {
                        __metadata: { type: entityType },
                        Title: cfg.titleFor(project, freshData),
                        ProjectId: project.Id.toString(),
                        ProfitCentre: project.PC || "",
                        Remarks: "",
                        OBL: String(project.PC ? 100000 : 0),
                        CreatedDate: new Date().toISOString(),
                        UpdatedDate: new Date().toISOString()
                    };
                    payload[cfg.jsonField] = JSON.stringify(freshData);
                    if (cfg.jsonField === 'ForecastData') {
                        payload.RevenueActual = "{}";
                    }
                    // Store Project Reference when the column exists - same rule as the
                    // Revenue/Collection page syncs: a LOOKUP column receives the target
                    // item ID as a string (ProjectReferenceId), a text column receives
                    // the text.
                    const refField = await cfg.probeRef();
                    if (refField.exists && attempt === 1) {
                        if (refField.type === "Lookup") {
                            payload.ProjectReferenceId = String(project.Id);
                        }
                        else {
                            payload.ProjectReference = project.ProjectReference || "";
                        }
                    }
                    if (attempt >= 2) {
                        delete payload.ProjectReference;
                        delete payload.ProjectReferenceId;
                    }
                    if (attempt >= 3) {
                        delete payload.OBL;
                    }
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${cfg.listName}')/items`, {
                        method: "POST",
                        headers: {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest
                        },
                        body: JSON.stringify(payload)
                    });
                    if (res.ok) {
                        ok = true;
                        added++;
                        console.log(`✅ Added: ${project.Title}${project.ProjectReference ? ' (' + project.ProjectReference + ')' : ''}`);
                    }
                    else {
                        const bodyText = await res.text();
                        let message = `HTTP ${res.status}`;
                        try {
                            const errJson = JSON.parse(bodyText);
                            message = ((_d = (_c = errJson === null || errJson === void 0 ? void 0 : errJson.error) === null || _c === void 0 ? void 0 : _c.message) === null || _d === void 0 ? void 0 : _d.value) || message;
                        }
                        catch (e) { /* keep http message */ }
                        console.error(`❌ Attempt ${attempt} failed for ${project.Title}: ${bodyText}`);
                        if (attempt === 2) {
                            failed++;
                            errors.push(`${project.Title} -> ${message}`);
                        }
                    }
                }
                catch (projectError) {
                    console.error(`❌ Attempt ${attempt} error for ${project.Title}:`, projectError);
                    if (attempt === 2) {
                        failed++;
                        errors.push(`${project.Title} -> ${projectError instanceof Error ? projectError.message : String(projectError)}`);
                    }
                }
                if (!ok)
                    await new Promise(resolve => setTimeout(resolve, 300));
            }
            await new Promise(resolve => setTimeout(resolve, 100));
        }
        if (added > 0) {
            await cfg.reload();
            console.log(`➕ ${cfg.summaryLabel} added ${added} project(s)` + (failed > 0 ? ` (${failed} failed)` : ""));
        }
        else if (failed > 0) {
            console.error(`❌ ${cfg.summaryLabel} failed for ${failed} project(s): ${errors.join(' | ')}`);
        }
    }
    catch (error) {
        console.error(`❌ ${cfg.summaryLabel} failed:`, error);
    }
};
/** Auto-sync: Project List → Revenue Forecast List (one row per missing
 *  visible project, empty ForecastData skeleton). Same logic as the Revenue
 *  page. Console-only reporting. */
const syncMissingProjectsToRevenueList = async () => {
    await syncMissingRows({
        listName: revenueListName,
        summaryLabel: 'Revenue auto-sync',
        records: revenueForecasts,
        skeleton: syncFreshForecastStructure,
        titleFor: syncForecastTitle,
        probeRef: syncRevenueRefField,
        jsonField: 'ForecastData',
        reload: loadRevenueForecasts
    });
};
/** Auto-sync: Project List → Collection List (one row per missing visible
 *  project, empty CollectionData skeleton). Same logic as the Collection
 *  page. Console-only reporting. */
const syncMissingProjectsToCollectionList = async () => {
    await syncMissingRows({
        listName: collectionListName,
        summaryLabel: 'Collection auto-sync',
        records: collectionItems,
        skeleton: syncFreshCollectionStructure,
        titleFor: syncCollectionTitle,
        probeRef: syncCollectionRefField,
        jsonField: 'CollectionData',
        reload: loadCollectionItems
    });
};
/* ============================================================
   PAGE MODULE
   ============================================================ */
const sprintPage = {
    content,
    style,
    script: () => {
        const remount = pageMounted;
        pageMounted = true;
        /* ---------- DOM LOOKUPS ---------- */
        const root = (document.querySelector('.daily-sprint-page') || document);
        const q = (sel) => root.querySelector(sel);
        /* ---------- loading overlay (same pattern as master.ts) ---------- */
        /**
         * ✅ PAGE LOADER - bar only, same as every other page: grey track
         * (#DCDCDC) + indigo sliding bar (#6664FF), no spinner and no text.
         * The status messages go to the console instead of the screen.
         * Inline styles on the shell keep it correct even before the CSS loads;
         * the bar itself is styled by .sprint-page-loader* (module style + css).
         */
        const createLoadingOverlay = (message = "Processing...") => {
            removeLoadingOverlay(); // never stack two loaders
            console.log(`⏳ ${message}`);
            const overlay = document.createElement("div");
            overlay.className = "sprint-loading-overlay";
            overlay.setAttribute("role", "progressbar");
            overlay.setAttribute("aria-busy", "true");
            overlay.setAttribute("aria-label", message);
            overlay.style.cssText =
                "position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;" +
                    "justify-content:center;background:rgba(248,249,250,0.92);z-index:999;";
            overlay.innerHTML = `
        <div class="sprint-loader-card">
          <div class="sprint-page-loader">
            <div class="sprint-page-loader__track">
              <div class="sprint-page-loader__bar"></div>
            </div>
          </div>
        </div>`;
            return overlay;
        };
        /** progress messages are console-only now (the loader shows no text) */
        const updateLoadingMessage = (overlay, message) => {
            void overlay;
            console.log(`⏳ ${message}`);
        };
        const removeLoadingOverlay = () => {
            document.querySelectorAll(".sprint-loading-overlay, .loading-overlay").forEach(el => el.remove());
        };
        /** mount the loader over the table container — never the whole page */
        const loaderHost = () => q(".sprint-data-container") || root;
        const toggleRevenueBtn = q("#toggleRevenue");
        const toggleCollectionBtn = q("#toggleCollection");
        const activeViewIndicator = q("#activeViewIndicator");
        const sprintTableBody = q(".sprint-table-body");
        const sprintTotalsHeader = q("#sprintTotalsHeader");
        const sprintMonthColumnsHeader = q("#sprintMonthColumnsHeader");
        const pageSizeSelect = q("#sprintPageSize");
        const firstPageBtn = q("#sprintFirstPage");
        const prevPageBtn = q("#sprintPrevPage");
        const nextPageBtn = q("#sprintNextPage");
        const lastPageBtn = q("#sprintLastPage");
        const currentPageSpan = q("#sprintCurrentPage");
        const totalPagesSpan = q("#sprintTotalPages");
        const pageNumbersDiv = q("#sprintPageNumbers");
        const startItemSpan = q("#sprintStartItem");
        const endItemSpan = q("#sprintEndItem");
        const totalItemsSpan = q("#sprintTotalItems");
        const filterYear = q("#sprintFilterYear");
        const filterMonth = q("#sprintFilterMonth");
        const filterProject = q("#sprintFilterProject");
        const filterCustomer = q("#sprintFilterCustomer");
        const filterProjectManager = q("#sprintFilterProjectManager");
        const filterGroupHead = q("#sprintFilterGroupHead");
        const filterProjectReference = q("#sprintFilterProjectReference");
        const filterEnggLead = q("#sprintFilterEnggLead");
        const clearFiltersBtn = q(".daily-sprint-clear-filters-btn");
        const concernsHint = q("#sprintConcernsHint");
        const required = [
            [".sprint-table-body", sprintTableBody],
            ["#toggleRevenue", toggleRevenueBtn],
            ["#toggleCollection", toggleCollectionBtn],
            ["#sprintFilterYear", filterYear],
            ["#sprintPageSize", pageSizeSelect],
            ["#sprintTotalsHeader", sprintTotalsHeader],
            ["#sprintMonthColumnsHeader", sprintMonthColumnsHeader]
        ];
        const missing = required.filter(([, el]) => !el).map(([s]) => s);
        if (missing.length)
            console.error(`❌ sprint page: missing element(s): ${missing.join(", ")}`);
        /* ---------- PER-MONTH VALUE READERS ---------- */
        /** Actual revenue for a month, read from the separate RevenueActual column. */
        const getActual = (item, monthKey) => {
            const map = safeNumberMap(item.RevenueActual);
            const n = Number(map[monthKey]);
            return Number.isFinite(n) ? n : 0;
        };
        const getMonthEntry = (item, monthKey) => safeJsonParse(item[dataField(currentViewMode)])[monthKey] || { plan: 0, actual: 0, forecast: 0 };
        /** Previous (last cycle's) plan for a month.
         *  Priority: PreviousPlanData["YYYY-MM"] → monthEntry.prevPlan / previousPlan → 0 */
        const getPrevPlan = (item, monthKey) => {
            var _a;
            const snap = safeNumberMap(item.PreviousPlanData);
            if (Number.isFinite(Number(snap[monthKey])))
                return Number(snap[monthKey]) || 0;
            const md = getMonthEntry(item, monthKey);
            const v = (_a = md.prevPlan) !== null && _a !== void 0 ? _a : md.previousPlan;
            return Number.isFinite(Number(v)) ? Number(v) || 0 : 0;
        };
        const getCurrPlan = (item, monthKey) => getMonthEntry(item, monthKey).plan || 0;
        const getFurther = (item, monthKey) => getMonthEntry(item, monthKey).forecast || 0;
        const getForecast = (item, monthKey) => getActual(item, monthKey) + getFurther(item, monthKey);
        const computeTotals = (data) => horizonKeys.map(key => {
            const t = { prevPlan: 0, plan: 0, actual: 0, further: 0, forecast: 0 };
            data.forEach(item => {
                t.prevPlan += getPrevPlan(item, key);
                t.plan += getCurrPlan(item, key);
                t.actual += getActual(item, key);
                t.further += getFurther(item, key);
            });
            t.forecast = t.actual + t.further;
            return t;
        });
        /* ---------- HEADER RENDERING ---------- */
        const sortBtns = (col) => {
            const asc = sortColumn === col && sortDirection === 'asc';
            const desc = sortColumn === col && sortDirection === 'desc';
            return `
        <div class="sprint-sort-buttons-container">
          <button class="sprint-sort-btn sprint-sort-asc${asc ? ' active' : ''}" data-sort="${col}" data-direction="asc" title="Sort Ascending">▲</button>
          <button class="sprint-sort-btn sprint-sort-desc${desc ? ' active' : ''}" data-sort="${col}" data-direction="desc" title="Sort Descending">▼</button>
        </div>`;
        };
        const subTh = (col, label, extraClass = '') => `
      <th class="sprint-subheader ${extraClass}" data-col="${col}">
        <div class="sprint-sort-header-container">
          <span class="sprint-header-label">${label}</span>
          ${sortBtns(col)}
        </div>
      </th>`;
        /** inline sticky-left geometry for a left column at position `i` */
        const leftCellStyle = (i) => {
            const c = leftCols()[i];
            const off = leftOffset(i);
            return `left:${off}px; width:${c.width}px; min-width:${c.width}px; max-width:${c.width}px;`;
        };
        const leftCellClass = (i, prefix) => {
            const c = leftCols()[i];
            const last = i === leftCols().length - 1;
            return [
                'sprint-left-cell',
                `sprint-${prefix}-${c.key}`,
                `sprint-col-${c.key}`,
                c.align ? `align-${c.align}` : '',
                last ? 'sprint-left-last' : ''
            ].filter(Boolean).join(' ');
        };
        const renderMonthColumnsHeader = () => {
            if (!sprintMonthColumnsHeader)
                return;
            const [m1, m2, m3] = horizonKeys;
            const m1Label = `${shortMonthLabel(m1)} (Current Month)`;
            /* Left block — sequence comes from LEFT_COLUMNS[currentViewMode] */
            const leftHeaders = leftCols().map((c, i) => `
          <th class="${leftCellClass(i, 'th')}" style="${leftCellStyle(i)}" rowspan="2">
            <div class="sprint-sort-header-container">
              <span class="sprint-header-label">${c.label}</span>
              ${c.sort ? sortBtns(c.sort) : ''}
            </div>
          </th>`).join('');
            /* Row 2 — month group bar: current month spans 5 cols, each future month spans 2 */
            const row2 = `
        <tr class="sprint-month-header">
          ${leftHeaders}

          <th class="month-current month-sep" colspan="5">
            <div class="sprint-sort-header-container">
              <span class="sprint-header-label">${m1Label}</span>
              ${sortBtns('forecast')}
            </div>
          </th>

          <th class="month-future month-sep" colspan="2">
            <div class="sprint-sort-header-container">
              <span class="sprint-header-label">${shortMonthLabel(m2)}</span>
              ${sortBtns('m2Plan')}
            </div>
          </th>

          <th class="month-future month-sep" colspan="2">
            <div class="sprint-sort-header-container">
              <span class="sprint-header-label">${shortMonthLabel(m3)}</span>
              ${sortBtns('m3Plan')}
            </div>
          </th>
        </tr>`;
            /* Row 3 — sub-columns */
            const row3 = `
        <tr class="sprint-subheader-row">
          ${subTh('prevPlan', 'Previous Plan', 'month-sep')}
          ${subTh('plan', 'Current Plan')}
          ${subTh('actual', 'Actuals')}
          ${subTh('further', 'Further')}
          ${subTh('forecast', 'Forecast')}

          ${subTh('m2PrevPlan', 'Previous Plan', 'month-sep')}
          ${subTh('m2Plan', 'Current Plan')}

          ${subTh('m3PrevPlan', 'Previous Plan', 'month-sep')}
          ${subTh('m3Plan', 'Current Plan')}
        </tr>`;
            sprintMonthColumnsHeader.innerHTML = row2 + row3;
            // table must be at least as wide as left block + month columns
            const table = root.querySelector('.daily-sprint-table');
            if (table) {
                const leftWidth = leftCols().reduce((sum, c) => sum + c.width, 0);
                table.style.minWidth = `${leftWidth + 5 * 105 + 4 * 105}px`;
            }
        };
        /**
         * ✅ THE FILTER BUTTONS LIVE IN THE TOTALS BAND.
         * The trigger + Apply pill used to sit in a toolbar card above the table.
         * They are now mounted at the LEFT of the "📊 Daily Sprint Totals" band -
         * directly above the "All Projects" cell.
         *
         * ⚠️ The band is re-painted on every filter / sort / page change, so the
         * two buttons must NOT be rebuilt with innerHTML: writing the band drops
         * the old nodes out of the DOM. The SAME nodes are therefore moved into
         * the fresh slot every time (appendChild moves, it never clones), which is
         * what keeps their click listeners, the badge count and the Apply pill's
         * hidden / disabled state alive across renders.
         *
         * totalsToolbar() prefers a node that is currently in the DOM (first run,
         * and again after an SPA re-mount that re-creates the markup); after the
         * band has been re-painted it falls back to the captured node, which is
         * detached but still alive and still holds its listeners.
         */
        let toolbarNode = null;
        const totalsToolbar = () => {
            const live = q('.sprint-table-toolbar');
            if (live)
                toolbarNode = live;
            return toolbarNode;
        };
        /** move the persistent toolbar into the freshly painted band */
        const mountTotalsToolbar = () => {
            const slot = sprintTotalsHeader === null || sprintTotalsHeader === void 0 ? void 0 : sprintTotalsHeader.querySelector('.sprint-totals-slot');
            const toolbar = totalsToolbar();
            if (!slot || !toolbar)
                return;
            if (toolbar.parentElement !== slot)
                slot.appendChild(toolbar);
        };
        const renderTotalsHeader = () => {
            if (!sprintTotalsHeader)
                return;
            const allData = currentRecords();
            const filtered = filteredData;
            const all = computeTotals(allData);
            const filt = computeTotals(filtered);
            const showFiltered = filtered.length > 0 && filtered.length !== allData.length;
            const sumOf = (rows, read) => rows.reduce((acc, r) => acc + (Number(read(r)) || 0), 0);
            const buildRow = (label, t, labelClass, rows) => {
                const cols = leftCols();
                const oblTotal = sumOf(rows, r => toNum(r.OBL));
                const recvTotal = sumOf(rows, getReceivables);
                // The trailing money column (OBL on Revenue, Receivables on Collection)
                // keeps its own cell so its total lines up under the header. Everything
                // before it collapses into ONE merged label cell:
                //   Revenue    → 7 merged (Project reference … Concerns) + OBL
                //   Collection → 6 merged (Project reference … Concerns) + Receivables
                const moneyIdx = cols.findIndex(c => c.key === 'obl' || c.key === 'receivables');
                const mergeCount = moneyIdx === -1 ? cols.length : moneyIdx;
                // merged cell spans the combined width of the columns it replaces so the
                // sticky-left geometry still matches the body rows
                const mergedWidth = cols.slice(0, mergeCount).reduce((sum, c) => sum + c.width, 0);
                const mergedStyle = `left:0px; width:${mergedWidth}px; min-width:${mergedWidth}px; max-width:${mergedWidth}px;`;
                const mergedCell = `
          <th class="sprint-left-cell sprint-th-merged ${labelClass}"
              style="${mergedStyle}" colspan="${mergeCount}">${label}</th>`;
                const moneyCell = moneyIdx === -1 ? '' : `
          <th class="${leftCellClass(moneyIdx, 'th')} sprint-totals-num"
              style="${leftCellStyle(moneyIdx)}">${cols[moneyIdx].key === 'obl' ? inr(oblTotal) : inr(recvTotal)}</th>`;
                return `
        <tr class="sprint-totals-detail">
          ${mergedCell}${moneyCell}

          <th class="sprint-totals-num month-sep">${inr(t[0].prevPlan)}</th>
          <th class="sprint-totals-num">${inr(t[0].plan)}</th>
          <th class="sprint-totals-num">${inr(t[0].actual)}</th>
          <th class="sprint-totals-num">${inr(t[0].further)}</th>
          <th class="sprint-totals-forecast">${inr(t[0].forecast)}</th>

          <th class="sprint-totals-num month-sep">${inr(t[1].prevPlan)}</th>
          <th class="sprint-totals-num">${inr(t[1].plan)}</th>

          <th class="sprint-totals-num month-sep">${inr(t[2].prevPlan)}</th>
          <th class="sprint-totals-num">${inr(t[2].plan)}</th>
        </tr>`;
            };
            const band = `
        <tr class="sprint-totals-bar">
          <th colspan="${totalCols()}" class="sprint-totals-band">
            <div class="sprint-totals-band-inner">
              <!-- the persistent filter toolbar is MOVED into this slot -->
              <span class="sprint-totals-slot"></span>
              <span class="sprint-totals-band-title">📊 Daily Sprint Totals — ${horizonKeys.map(shortMonthLabel).join(' · ')}</span>
            </div>
          </th>
        </tr>`;
            let html = band;
            if (allData.length > 0)
                html += buildRow('All Projects', all, 'sprint-totals-label', allData);
            if (showFiltered)
                html += buildRow('Filtered', filt, 'sprint-totals-label-filtered', filtered);
            sprintTotalsHeader.innerHTML = html;
            mountTotalsToolbar();
            applyStickyHeaderOffsets();
        };
        /**
         * Assign a cumulative `top` to every header row so that all thead rows
         * (totals band → totals detail → month bar → sub-columns) stack and stay
         * pinned while the table body scrolls underneath.
         */
        /**
         * Find the app's fixed/sticky main header (SharePoint suite bar, SPA
         * chrome, etc.) so our table header parks directly beneath it instead of
         * sliding underneath. Returns its bottom edge in px.
         */
        /**
         * position:sticky silently stops working if ANY ancestor has an overflow
         * other than visible (or a transform/filter/contain, which create a
         * containing block). SPFx wraps web parts in several such divs. Walk up
         * from the table and neutralise the offenders, remembering the nearest
         * real scrolling ancestor so we can listen to the right scroll events.
         */
        let scrollParent = window;
        /**
         * READ-ONLY. Find the element that actually scrolls our table.
         *
         * ⚠️ This deliberately does NOT modify any ancestor. An earlier version
         * force-set overflow/transform on ancestors to "unlock" sticky, but those
         * are the host app's own layout nodes — clobbering them broke page
         * scrolling and hid the sidebar (which is positioned with a transform).
         * We only ever style elements inside our own root.
         */
        /** ancestors we relaxed, so we can put them back exactly as they were */
        const patchedAncestors = [];
        /**
         * Find the element that actually scrolls our table, and clear the one
         * thing that silently defeats sticky: a NON-scrolling ancestor with
         * overflow hidden/clip between us and that scroller.
         *
         * ⚠️ Deliberately conservative. An earlier version rewrote overflow AND
         * transform on every ancestor, which broke the host page's scrolling and
         * hid the sidebar (positioned via transform). Now we:
         *   - stop at the scroller and never look past it
         *   - only touch elements that clip WITHOUT scrolling (a no-op for layout)
         *   - never touch transform/filter/contain
         *   - record every change so restoreAncestors() can undo it
         */
        const detectScrollParent = () => {
            const table = root.querySelector('.daily-sprint-table');
            if (!table) {
                scrollParent = window;
                return;
            }
            const clipping = [];
            let el = table.parentElement;
            let found = window;
            while (el && el !== document.body && el !== document.documentElement) {
                const cs = getComputedStyle(el);
                const scrolls = /(auto|scroll|overlay)/.test(cs.overflowY)
                    && el.scrollHeight > el.clientHeight + 1;
                if (scrolls) {
                    found = el;
                    break;
                } // reached the scrollport - stop
                // clips but never scrolls => harmless to relax, fatal to sticky if left
                if (/(hidden|clip)/.test(cs.overflowY) && el.scrollHeight <= el.clientHeight + 1) {
                    clipping.push(el);
                }
                el = el.parentElement;
            }
            clipping.forEach(node => {
                if (patchedAncestors.some(p => p.el === node))
                    return;
                patchedAncestors.push({ el: node, prev: node.style.overflow });
                node.style.overflow = 'visible';
            });
            scrollParent = found;
        };
        /** put every ancestor we relaxed back the way we found it */
        const restoreAncestors = () => {
            patchedAncestors.forEach(({ el, prev }) => { el.style.overflow = prev; });
            patchedAncestors.length = 0;
        };
        /**
         * Find the app's fixed/sticky main header (SharePoint suite bar, SPA
         * chrome, etc.) so our table header parks directly beneath it instead of
         * sliding underneath. Returns its bottom edge in px.
         */
        const measureAppHeaderOffset = () => {
            const candidates = [
                '#SuiteNavPlaceHolder', '#spSiteHeader', '#sp-appBar', '#sp-appBar-mobile',
                '.od-TopBar', '.ms-CommandBar', '.sp-appBar',
                '[class*="commandBarWrapper"]', '[data-automation-id="pageCommandBar"]',
                '.app-header', '.main-header', 'header.sticky', '[data-sprint-header]'
            ];
            // collect every docked bar first
            const bars = [];
            const seen = new Set();
            candidates.forEach(sel => {
                let nodes;
                try {
                    nodes = document.querySelectorAll(sel);
                }
                catch (_a) {
                    return;
                }
                nodes.forEach(el => {
                    if (seen.has(el))
                        return;
                    seen.add(el);
                    const cs = getComputedStyle(el);
                    if (cs.position !== 'fixed' && cs.position !== 'sticky')
                        return;
                    if (cs.display === 'none' || cs.visibility === 'hidden')
                        return;
                    const r = el.getBoundingClientRect();
                    if (r.height > 0 && r.top < 200)
                        bars.push({ top: r.top, bottom: r.bottom });
                });
            });
            // Chrome bars STACK (suite bar at 0, command bar beneath it). Walk the
            // chain: start at the viewport top and keep absorbing any bar that
            // begins at/above the current bottom edge.
            bars.sort((a, b) => a.top - b.top);
            let bottom = 0;
            let grew = true;
            while (grew) {
                grew = false;
                bars.forEach(bar => {
                    if (bar.top <= bottom + 4 && bar.bottom > bottom) {
                        bottom = bar.bottom;
                        grew = true;
                    }
                });
            }
            return Math.round(bottom);
        };
        /**
         * Stack every thead row with a cumulative `top`, starting below the app
         * header, so all four rows stay pinned to the MAIN header while the page
         * scrolls.
         */
        const applyStickyHeaderOffsets = () => {
            const rows = [];
            if (sprintTotalsHeader)
                rows.push(...Array.from(sprintTotalsHeader.querySelectorAll('tr')));
            if (sprintMonthColumnsHeader)
                rows.push(...Array.from(sprintMonthColumnsHeader.querySelectorAll('tr')));
            detectScrollParent();
            // `top` on a sticky cell is measured from the top edge of its scrollport.
            // When the window scrolls, that edge IS the viewport top, so the app
            // header's bottom is the right offset. But when an inner region scrolls
            // (SharePoint modern), the region already starts below the app header —
            // using the viewport figure would push our header that much too far down.
            const headerBottom = measureAppHeaderOffset();
            let base = headerBottom;
            if (scrollParent !== window) {
                const portTop = scrollParent.getBoundingClientRect().top;
                base = Math.max(0, headerBottom - portTop);
            }
            root.style.setProperty('--sprint-header-offset', `${base}px`);
            let offset = base;
            rows.forEach((row, idx) => {
                const height = row.getBoundingClientRect().height;
                row.querySelectorAll('th, td').forEach(cell => {
                    cell.style.position = 'sticky';
                    cell.style.top = `${offset}px`;
                    // later rows must paint over earlier ones when they collapse together
                    const isLeftSticky = /sprint-(th|td)-/.test(cell.className);
                    cell.style.zIndex = String((isLeftSticky ? 30 : 20) + idx);
                });
                offset += height;
            });
        };
        let activeEditor = null;
        /** contenteditable gives us innerText, so
    /<div> come back as newlines */
        const readEditorText = (el) => (el.innerText || el.textContent || '')
            .replace(/\u00a0/g, ' ') // contenteditable inserts nbsp
            .replace(/\r\n/g, '\n')
            .replace(/\n+$/, '');
        const setConcernsStatus = (itemId, state, glyph = '') => {
            const cell = root.querySelector(`td[data-concerns-cell="${itemId}"]`);
            const status = root.querySelector(`[data-concerns-status="${itemId}"]`);
            if (!cell && !status)
                return; // row re-rendered / off-page
            cell === null || cell === void 0 ? void 0 : cell.classList.remove('is-saving', 'is-saved', 'is-error');
            if (state !== 'idle')
                cell === null || cell === void 0 ? void 0 : cell.classList.add(`is-${state}`);
            if (status)
                status.textContent = glyph;
            if (state === 'saved' || state === 'error') {
                window.setTimeout(() => {
                    cell === null || cell === void 0 ? void 0 : cell.classList.remove('is-saved', 'is-error');
                    const s = root.querySelector(`[data-concerns-status="${itemId}"]`);
                    if (s)
                        s.textContent = '';
                }, state === 'saved' ? 1600 : 4000);
            }
        };
        /** fire the MERGE and keep the UI (and the in-memory row) in step */
        const persistConcerns = async (itemId, text) => {
            const record = currentRecords().find(r => Number(r.Id) === Number(itemId));
            if (!record) {
                toast('error', 'Row not found — please refresh the page');
                return;
            }
            const listName = currentListName();
            setConcernsStatus(itemId, 'saving', '💾');
            const result = await saveConcernsToSharePoint(listName, record, text);
            if (!result.ok) {
                console.error(`❌ Concerns save failed — ${listName} #${itemId}:`, result.message);
                setConcernsStatus(itemId, 'error', '⚠️');
                toast('error', /403|denied|access|permission/i.test(result.message)
                    ? 'You do not have permission to edit this row'
                    : 'Could not save the concern');
                return;
            }
            applyConcernsLocally(record, text, result.field);
            setConcernsStatus(itemId, 'saved', '✅');
            const editor = root.querySelector(`.sprint-concerns-editor[data-concerns-id="${itemId}"]`);
            if (editor) {
                editor.title = text;
                editor.classList.toggle('is-empty', !text.trim());
            }
            console.log(`💾 Concerns saved → ${listName} #${itemId} (${result.field}): "${text}"`);
        };
        /**
         * Close the open editor and save it if the text actually changed.
         * Called on blur, on Enter, and defensively before every re-render
         * (paging/sorting/filtering would otherwise drop the typed text).
         */
        const commitActiveEditor = () => {
            var _a;
            if (!activeEditor)
                return;
            const { el, itemId, original } = activeEditor;
            activeEditor = null;
            (_a = el.closest('td')) === null || _a === void 0 ? void 0 : _a.classList.remove('is-editing');
            const cancelled = el.dataset.cancelled === '1';
            delete el.dataset.cancelled;
            if (cancelled) {
                el.innerText = original;
                el.classList.toggle('is-empty', !original.trim());
                return;
            }
            const next = readEditorText(el);
            if (next === original)
                return;
            void persistConcerns(itemId, next);
        };
        /* ---------- BODY RENDERING ---------- */
        const renderTableBody = (rows) => {
            if (!sprintTableBody)
                return;
            // never silently throw away text that is being typed
            commitActiveEditor();
            if (rows.length === 0) {
                sprintTableBody.innerHTML = `<tr><td colspan="${totalCols()}" class="daily-sprint-empty">No data found</td></tr>`;
                return;
            }
            const [m1, m2, m3] = horizonKeys;
            sprintTableBody.innerHTML = rows.map(item => {
                var _a;
                const projectTitle = ((_a = item.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                const proj = projects.find(p => p.Title === projectTitle);
                const projectRef = getSafeString(proj === null || proj === void 0 ? void 0 : proj.ProjectReference);
                const customerName = getSafeString(proj === null || proj === void 0 ? void 0 : proj.CustomerName);
                const groupHeadName = getGroupHeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.GroupHead1) || "-";
                const pmName = getProjectManagerDisplayName(proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-";
                const enggLeadName = getEnggLeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.EnggLead);
                const sprintCell = getConcerns(item);
                const receivables = getReceivables(item);
                return `
          <tr data-item-id="${item.Id}">
            ${leftCols().map((c, i) => {
                    const style = leftCellStyle(i);
                    const cls = leftCellClass(i, 'td');
                    switch (c.key) {
                        case 'projectRef':
                            return `<td class="${cls}" style="${style}">${escapeAttr(projectRef)}</td>`;
                        case 'customer':
                            return `<td class="${cls}" style="${style}">${escapeAttr(customerName)}</td>`;
                        case 'project':
                            return `<td class="${cls}" style="${style}font-weight:600;color:#333;">${escapeAttr(projectTitle)}</td>`;
                        case 'gh':
                            return `<td class="${cls}" style="${style}">${escapeAttr(groupHeadName)}</td>`;
                        case 'pm':
                            return `<td class="${cls}" style="${style}">${escapeAttr(pmName)}</td>`;
                        case 'enggLead':
                            return `<td class="${cls}" style="${style}"><span class="daily-sprint-pill">${escapeAttr(enggLeadName)}</span></td>`;
                        case 'concerns': {
                            // ✏️ editable → MERGEd to the list on blur / Enter
                            if (!canEditConcerns()) {
                                return `<td class="${cls}" style="${style}" title="${escapeAttr(sprintCell)}">${escapeAttr(sprintCell || '-')}</td>`;
                            }
                            return `
                    <td class="${cls} sprint-concerns-cell" style="${style}"
                        data-concerns-cell="${item.Id}" title="${escapeAttr(sprintCell)}">
                      <div class="sprint-concerns-wrap">
                        <div class="sprint-concerns-editor${sprintCell ? '' : ' is-empty'}"
                             contenteditable="true" role="textbox" aria-multiline="true"
                             aria-label="Concerns for ${escapeAttr(projectTitle)}"
                             data-concerns-id="${item.Id}"
                             data-placeholder="Add concern…"
                             title="${escapeAttr(sprintCell)}">${escapeHtml(sprintCell)}</div>
                        <span class="sprint-concerns-status" data-concerns-status="${item.Id}"></span>
                      </div>
                    </td>`;
                        }
                        case 'obl':
                            return `<td class="${cls}" style="${style}">${inr(toNum(item.OBL))}</td>`;
                        case 'receivables':
                            return `<td class="${cls}" style="${style}">${inr(receivables)}</td>`;
                        default:
                            return `<td class="${cls}" style="${style}"></td>`;
                    }
                }).join('')}

            <td class="num prevplan month-sep">${inr(getPrevPlan(item, m1))}</td>
            <td class="num plan">${inr(getCurrPlan(item, m1))}</td>
            <td class="num actual">${inr(getActual(item, m1))}</td>
            <td class="num">${inr(getFurther(item, m1))}</td>
            <td class="num forecast">${inr(getForecast(item, m1))}</td>

            <td class="num prevplan future-cell month-sep">${inr(getPrevPlan(item, m2))}</td>
            <td class="num plan future-cell">${inr(getCurrPlan(item, m2))}</td>

            <td class="num prevplan future-cell month-sep">${inr(getPrevPlan(item, m3))}</td>
            <td class="num plan future-cell">${inr(getCurrPlan(item, m3))}</td>
          </tr>`;
            }).join('');
        };
        /* ---------- FILTERS ---------- */
        /**
         * The three-month horizon follows the APPLIED filters (not the dropdowns),
         * so opening the popup and picking a year never moves the table until
         * "Apply" is pressed. With no year/month applied it falls back to the
         * current month exactly as before.
         */
        const syncMonthSelection = () => {
            const year = appliedFilters.year || String(new Date().getFullYear());
            const month = pad2(appliedFilters.month || String(new Date().getMonth() + 1));
            activeMonthKey = `${year}-${month}`;
            horizonKeys = buildHorizon(activeMonthKey);
        };
        const initializeFilters = () => {
            if (!filterYear || !filterMonth)
                return;
            const currentYear = new Date().getFullYear();
            // ✅ CLEAN START - no year/month is pre-selected (badge 0, no chips,
            // Apply gated), same as the other six pages.
            filterYear.innerHTML = '<option value="">All Years</option>' +
                Array.from({ length: 4 }, (_, i) => currentYear - 1 + i)
                    .map(y => `<option value="${y}">${y}</option>`)
                    .join('');
            filterMonth.innerHTML = '<option value="">All Months</option>' +
                MONTHS.map((m, i) => `<option value="${i + 1}">${m}</option>`)
                    .join('');
            const assignedProjects = getAssignedProjects(projects);
            const collect = (read) => {
                const values = new Set();
                assignedProjects.forEach(p => {
                    const value = read(p);
                    if (value && value.trim())
                        values.add(value.trim());
                });
                return Array.from(values).sort()
                    .map(v => `<option value="${escapeAttr(v)}">${escapeAttr(v)}</option>`)
                    .join('');
            };
            const fill = (sel, allLabel, read) => {
                if (sel)
                    sel.innerHTML = `<option value="">${allLabel}</option>` + collect(read);
            };
            fill(filterProject, 'All Projects', p => p.Title);
            fill(filterCustomer, 'All Customers', p => p.CustomerName);
            fill(filterProjectManager, 'All Managers', p => getProjectManagerDisplayName(p.ProjectManager));
            fill(filterGroupHead, 'All Group Heads', p => getGroupHeadDisplayName(p.GroupHead1));
            fill(filterProjectReference, 'All References', p => p.ProjectReference);
            fill(filterEnggLead, 'All Engg Leads', p => getEnggLeadDisplayName(p.EnggLead));
            // popup + badge start in sync with the (empty) applied filter set
            restoreAppliedFilters();
            syncFilterUI();
        };
        /* ---------- SORT ---------- */
        const getSortValue = (item, col) => {
            var _a;
            const projectTitle = ((_a = item.Project) === null || _a === void 0 ? void 0 : _a.Title) || '';
            const proj = projects.find(p => p.Title === projectTitle);
            const [m1, m2, m3] = horizonKeys;
            switch (col) {
                case 'projectRef': return (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || '';
                case 'project': return projectTitle;
                case 'customer': return (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || '';
                case 'groupHead': return getGroupHeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.GroupHead1);
                case 'projectManager': return getProjectManagerDisplayName(proj === null || proj === void 0 ? void 0 : proj.ProjectManager);
                case 'enggLead': return getEnggLeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.EnggLead);
                case 'obl': return toNum(item.OBL);
                case 'receivables': return getReceivables(item);
                case 'prevPlan': return getPrevPlan(item, m1);
                case 'plan': return getCurrPlan(item, m1);
                case 'actual': return getActual(item, m1);
                case 'further': return getFurther(item, m1);
                case 'forecast': return getForecast(item, m1);
                case 'm2PrevPlan': return getPrevPlan(item, m2);
                case 'm2Plan': return getCurrPlan(item, m2);
                case 'm3PrevPlan': return getPrevPlan(item, m3);
                case 'm3Plan': return getCurrPlan(item, m3);
                default: return '';
            }
        };
        const sortData = (data) => {
            if (!sortColumn)
                return [...data];
            const col = sortColumn;
            return [...data].sort((a, b) => {
                const va = getSortValue(a, col);
                const vb = getSortValue(b, col);
                if (typeof va === 'number' && typeof vb === 'number') {
                    return sortDirection === 'asc' ? va - vb : vb - va;
                }
                const sa = String(va).toLowerCase();
                const sb = String(vb).toLowerCase();
                return sortDirection === 'asc' ? sa.localeCompare(sb) : sb.localeCompare(sb);
            });
        };
        const toggleSort = (col, direction) => {
            if (direction) {
                sortColumn = col;
                sortDirection = direction;
            }
            else if (sortColumn === col) {
                sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
            }
            else {
                sortColumn = col;
                sortDirection = 'asc';
            }
            renderMonthColumnsHeader();
            renderCurrentPage();
            applyStickyHeaderOffsets();
        };
        /* ---------- PAGINATION ---------- */
        const calculateTotalPages = () => {
            totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
        };
        const updatePageNumbers = () => {
            if (!pageNumbersDiv)
                return;
            pageNumbersDiv.innerHTML = "";
            if (filteredData.length === 0)
                return;
            const max = 5;
            let start = Math.max(1, currentPage - Math.floor(max / 2));
            const end = Math.min(totalPages, start + max - 1);
            if (end - start + 1 < max)
                start = Math.max(1, end - max + 1);
            if (start > 1) {
                const dots = document.createElement("button");
                dots.className = "sprint-page-number";
                dots.textContent = "…";
                dots.disabled = true;
                pageNumbersDiv.appendChild(dots);
            }
            for (let i = start; i <= end; i++) {
                const b = document.createElement("button");
                b.className = `sprint-page-number${i === currentPage ? " active" : ""}`;
                b.textContent = String(i);
                b.addEventListener("click", () => goToPage(i));
                pageNumbersDiv.appendChild(b);
            }
            if (end < totalPages) {
                const dots = document.createElement("button");
                dots.className = "sprint-page-number";
                dots.textContent = "…";
                dots.disabled = true;
                pageNumbersDiv.appendChild(dots);
            }
        };
        const updatePaginationControls = () => {
            if (currentPageSpan)
                currentPageSpan.textContent = String(currentPage);
            if (totalPagesSpan)
                totalPagesSpan.textContent = String(totalPages);
            if (filteredData.length === 0) {
                if (startItemSpan)
                    startItemSpan.textContent = "0";
                if (endItemSpan)
                    endItemSpan.textContent = "0";
                if (totalItemsSpan)
                    totalItemsSpan.textContent = "0";
            }
            else {
                const start = (currentPage - 1) * pageSize + 1;
                const end = Math.min(currentPage * pageSize, filteredData.length);
                if (startItemSpan)
                    startItemSpan.textContent = String(start);
                if (endItemSpan)
                    endItemSpan.textContent = String(end);
                if (totalItemsSpan)
                    totalItemsSpan.textContent = String(filteredData.length);
            }
            if (pageSizeSelect)
                pageSizeSelect.value = String(pageSize);
            const noData = filteredData.length === 0;
            if (firstPageBtn)
                firstPageBtn.disabled = noData || currentPage === 1;
            if (prevPageBtn)
                prevPageBtn.disabled = noData || currentPage === 1;
            if (nextPageBtn)
                nextPageBtn.disabled = noData || currentPage >= totalPages;
            if (lastPageBtn)
                lastPageBtn.disabled = noData || currentPage >= totalPages;
            updatePageNumbers();
        };
        const goToPage = (page) => {
            if (page < 1 || page > totalPages || page === currentPage)
                return;
            if (filteredData.length === 0)
                return;
            currentPage = page;
            updatePaginationControls();
            renderCurrentPage();
            // bring the table top just under the pinned header, on whichever
            // element actually scrolls (window, or an SPFx inner region)
            const table = root.querySelector('.daily-sprint-table');
            if (table) {
                const gap = measureAppHeaderOffset() + 8;
                if (scrollParent === window) {
                    const top = window.scrollY + table.getBoundingClientRect().top - gap;
                    window.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
                }
                else {
                    const host = scrollParent;
                    const delta = table.getBoundingClientRect().top - host.getBoundingClientRect().top;
                    host.scrollTo({ top: Math.max(0, host.scrollTop + delta - 8), behavior: 'smooth' });
                }
            }
        };
        const changePageSize = (newSize) => {
            // ✅ keep the first visible record in view when the page size changes
            // (same behaviour as Revenue / Collection / Procurement)
            const firstVisibleIndex = Math.max(0, (currentPage - 1) * pageSize);
            pageSize = (!newSize || newSize <= 0) ? 10 : newSize;
            calculateTotalPages();
            currentPage = Math.min(totalPages, Math.floor(firstVisibleIndex / pageSize) + 1);
            updatePaginationControls();
            renderCurrentPage();
        };
        const renderCurrentPage = () => {
            const sorted = sortData(filteredData);
            if (sorted.length === 0) {
                renderTableBody([]);
                renderTotalsHeader();
                return;
            }
            if (currentPage > totalPages)
                currentPage = totalPages;
            const startIndex = (currentPage - 1) * pageSize;
            const endIndex = Math.min(startIndex + pageSize, sorted.length);
            renderTableBody(sorted.slice(startIndex, endIndex));
            renderTotalsHeader();
        };
        const EMPTY_FILTERS = {
            year: '', month: '', project: '', customer: '', manager: '', head: '', reference: '', lead: ''
        };
        let appliedFilters = Object.assign({}, EMPTY_FILTERS);
        /** key → <select> map used by read / restore / clear / chips */
        const FILTER_FIELDS = [
            ['year', filterYear], ['month', filterMonth], ['project', filterProject],
            ['customer', filterCustomer], ['manager', filterProjectManager],
            ['head', filterGroupHead], ['reference', filterProjectReference], ['lead', filterEnggLead]
        ];
        const FILTER_LABELS = {
            year: 'Year', month: 'Month', project: 'Project', customer: 'Customer',
            manager: 'Project Manager', head: 'Group Head', reference: 'Project Reference', lead: 'Engg Lead'
        };
        const filterOverlay = q('#sprintFilterOverlay');
        const filterTrigger = q('#sprintFilterTrigger');
        const filterCountBadge = q('#sprintFilterCount');
        const filterApplyPill = q('#sprintFilterApplyPill');
        const filterChipsHost = q('#sprintFilterChips');
        const filterCancelBtn = q('#sprintFilterCancel');
        const filterClearBtn = q('#sprintFilterClear');
        const filterApplyBtn = q('#sprintFilterApply');
        const filterResetBtn = q('#sprintResetFilterBtn');
        const filterCloseBtn = q('#sprintFilterClose');
        /** read the eight dropdowns (= the draft) */
        const readDraftFilters = () => {
            const out = Object.assign({}, EMPTY_FILTERS);
            FILTER_FIELDS.forEach(([key, sel]) => { out[key] = sel ? sel.value : ''; });
            return out;
        };
        /** push the applied filters back into the dropdowns */
        const restoreAppliedFilters = () => {
            FILTER_FIELDS.forEach(([key, sel]) => { if (sel)
                sel.value = appliedFilters[key]; });
        };
        const sameFilters = (a, b) => Object.keys(EMPTY_FILTERS).every(k => a[k] === b[k]);
        const activeFilterCount = (f) => Object.keys(EMPTY_FILTERS).filter(k => !!f[k]).length;
        const filtersAreDirty = () => !sameFilters(readDraftFilters(), appliedFilters);
        /** chips for the APPLIED filters (chip × removes one, "Clear all" removes every one) */
        const renderFilterChips = () => {
            if (!filterChipsHost)
                return;
            if (activeFilterCount(appliedFilters) === 0) {
                filterChipsHost.hidden = true;
                filterChipsHost.innerHTML = '';
                return;
            }
            const chips = Object.keys(EMPTY_FILTERS)
                .filter(k => !!appliedFilters[k])
                .map(k => `
          <span class="sprint-filter-chip">
            <span class="sprint-filter-chip__label">${FILTER_LABELS[k]}:</span>
            <span class="sprint-filter-chip__value">${escapeHtml(appliedFilters[k])}</span>
            <button type="button" data-clear-filter="${k}" title="Remove this filter" aria-label="Remove ${FILTER_LABELS[k]} filter">&times;</button>
          </span>`)
                .join('');
            filterChipsHost.innerHTML =
                chips + `<button type="button" class="sprint-filter-chip__all" data-clear-filter="__all">Clear all</button>`;
            filterChipsHost.hidden = false;
        };
        /** badge count, Apply pill, disabled states - always from the applied state */
        const syncFilterUI = () => {
            const active = activeFilterCount(appliedFilters);
            const draftCount = activeFilterCount(readDraftFilters());
            const dirty = filtersAreDirty();
            if (filterCountBadge) {
                filterCountBadge.textContent = String(active);
                filterCountBadge.classList.toggle('is-zero', active === 0);
            }
            if (filterTrigger)
                filterTrigger.classList.toggle('is-active', active > 0);
            if (filterApplyPill)
                filterApplyPill.hidden = !dirty;
            if (filterApplyBtn)
                filterApplyBtn.disabled = !dirty;
            if (filterClearBtn)
                filterClearBtn.disabled = draftCount === 0;
            if (filterResetBtn)
                filterResetBtn.disabled = draftCount === 0;
            renderFilterChips();
        };
        const openFilterModal = () => {
            if (!filterOverlay)
                return;
            restoreAppliedFilters(); // always start from what is applied
            filterOverlay.classList.add('is-open');
            filterOverlay.setAttribute('aria-hidden', 'false');
            filterTrigger === null || filterTrigger === void 0 ? void 0 : filterTrigger.setAttribute('aria-expanded', 'true');
            document.body.classList.add('sprint-filter-open');
            syncFilterUI();
            window.setTimeout(() => { var _a; return (_a = FILTER_FIELDS[0][1]) === null || _a === void 0 ? void 0 : _a.focus(); }, 30);
        };
        const closeFilterModal = (commit = false) => {
            if (!filterOverlay)
                return;
            if (!commit)
                restoreAppliedFilters(); // cancel = drop the draft
            filterOverlay.classList.remove('is-open');
            filterOverlay.setAttribute('aria-hidden', 'true');
            filterTrigger === null || filterTrigger === void 0 ? void 0 : filterTrigger.setAttribute('aria-expanded', 'false');
            document.body.classList.remove('sprint-filter-open');
            syncFilterUI();
        };
        /** Apply: the draft becomes the applied filter set and the grid refreshes */
        const commitFilters = () => {
            appliedFilters = readDraftFilters();
            closeFilterModal(true);
            filterTable();
            syncFilterUI();
            const n = activeFilterCount(appliedFilters);
            toast('info', n === 0 ? 'Filters cleared' : `${n} filter${n > 1 ? 's' : ''} applied`);
        };
        /** chip × - remove one applied filter ("__all" = remove every one) */
        const clearAppliedFilter = (key) => {
            if (key === '__all') {
                appliedFilters = Object.assign({}, EMPTY_FILTERS);
            }
            else {
                appliedFilters = Object.assign(Object.assign({}, appliedFilters), { [key]: '' });
            }
            restoreAppliedFilters();
            filterTable();
            syncFilterUI();
        };
        const setupFilterEventListeners = () => {
            filterTrigger === null || filterTrigger === void 0 ? void 0 : filterTrigger.addEventListener('click', openFilterModal);
            filterApplyPill === null || filterApplyPill === void 0 ? void 0 : filterApplyPill.addEventListener('click', commitFilters);
            filterCloseBtn === null || filterCloseBtn === void 0 ? void 0 : filterCloseBtn.addEventListener('click', () => closeFilterModal(false));
            filterCancelBtn === null || filterCancelBtn === void 0 ? void 0 : filterCancelBtn.addEventListener('click', () => closeFilterModal(false));
            filterApplyBtn === null || filterApplyBtn === void 0 ? void 0 : filterApplyBtn.addEventListener('click', commitFilters);
            filterClearBtn === null || filterClearBtn === void 0 ? void 0 : filterClearBtn.addEventListener('click', clearFilters);
            filterResetBtn === null || filterResetBtn === void 0 ? void 0 : filterResetBtn.addEventListener('click', clearFilters);
            // click on the scrim closes without saving
            filterOverlay === null || filterOverlay === void 0 ? void 0 : filterOverlay.addEventListener('click', (e) => {
                if (e.target === filterOverlay)
                    closeFilterModal(false);
            });
            // Esc closes without saving
            filterOverlay === null || filterOverlay === void 0 ? void 0 : filterOverlay.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    e.preventDefault();
                    closeFilterModal(false);
                }
            });
            // changing a dropdown only updates the UI state (badge, Apply pill, buttons)
            FILTER_FIELDS.forEach(([, sel]) => sel === null || sel === void 0 ? void 0 : sel.addEventListener('change', syncFilterUI));
            // chip × buttons (delegated - chips are re-rendered on every apply)
            filterChipsHost === null || filterChipsHost === void 0 ? void 0 : filterChipsHost.addEventListener('click', (e) => {
                const btn = e.target.closest('[data-clear-filter]');
                if (!btn)
                    return;
                clearAppliedFilter(btn.dataset.clearFilter || '');
            });
        };
        /* ============================================================
           🔐 ACCESS-LEVEL PROJECT VISIBILITY
           ------------------------------------------------------------
           Admin (canReadAllProjects)            → all projects in Project List.
           Everyone else (GH Head / PM Manager) → only getAssignedProjects(projects).
           Same rule used by:
             • the filter dropdowns (initializeFilters → assignedProjects)
             • the auto-sync        (syncMissingRows  → visibleProjects)
             • the Revenue / Collection pages
           Titles are compared because forecast rows only expose Project.Title.
           ============================================================ */
        const getVisibleProjectTitles = () => {
            const visible = canReadAllProjects() ? projects : getAssignedProjects(projects);
            return new Set(visible
                .map(p => ((p === null || p === void 0 ? void 0 : p.Title) || '').trim())
                .filter(Boolean));
        };
        /* ---------- FILTER ---------- */
        const filterTable = () => {
            // Keep the 3-month horizon in sync with the APPLIED Year/Month filters
            syncMonthSelection();
            // ── read the APPLIED filters (dropdowns are only a draft while popup is open)
            const year = appliedFilters.year;
            const month = appliedFilters.month;
            const project = appliedFilters.project;
            const customer = appliedFilters.customer;
            const manager = appliedFilters.manager;
            const head = appliedFilters.head;
            const reference = appliedFilters.reference;
            const lead = appliedFilters.lead;
            // 🔐 Build the access-level allow-list ONCE per filter pass.
            //    Admin → every project title; everyone else → assigned titles only.
            const visibleTitles = getVisibleProjectTitles();
            const matches = (item) => {
                var _a;
                const projectTitle = (_a = item.Project) === null || _a === void 0 ? void 0 : _a.Title;
                if (!projectTitle)
                    return false;
                // 🔐 ACCESS GATE — non-admins only see rows for their assigned projects.
                //    This single check scopes the GRID, TOTALS band, SORT, PAGINATION
                //    and any downstream EXPORT identically to the Revenue/Collection pages.
                if (!visibleTitles.has(projectTitle.trim()))
                    return false;
                // Keep the row only if ANY month in the 3-month horizon has data
                // (only when BOTH Year + Month are applied — same as before)
                if (year && month) {
                    const data = safeJsonParse(item[dataField(currentViewMode)]);
                    const actuals = safeNumberMap(item.RevenueActual);
                    const hasAny = horizonKeys.some(k => (k in data) || (k in actuals));
                    if (!hasAny)
                        return false;
                }
                // Exact Project filter (dropdown already scoped by assignedProjects,
                // but the row check stays for safety / URL-driven states)
                if (project && projectTitle !== project)
                    return false;
                // Project-attribute filters — need the Project List row
                if (customer || manager || head || reference || lead) {
                    const proj = projects.find(p => p.Title === projectTitle);
                    if (!proj)
                        return false;
                    if (customer && proj.CustomerName !== customer)
                        return false;
                    if (manager && getProjectManagerDisplayName(proj.ProjectManager) !== manager)
                        return false;
                    if (head && getGroupHeadDisplayName(proj.GroupHead1) !== head)
                        return false;
                    if (reference && proj.ProjectReference !== reference)
                        return false;
                    if (lead && getEnggLeadDisplayName(proj.EnggLead) !== lead)
                        return false;
                }
                return true;
            };
            // Apply the gate + user filters to the active list's rows
            filteredData = currentRecords().filter(matches);
            // Reset paging, recompute pages, repaint header + body + sticky offsets
            currentPage = 1;
            calculateTotalPages();
            updatePaginationControls();
            renderMonthColumnsHeader(); // month labels follow the Year/Month filter
            renderCurrentPage();
            applyStickyHeaderOffsets();
        };
        /**
         * Reset / Clear (both live INSIDE the popup) empty the DRAFT only - the
         * grid keeps showing the currently applied filters until "Apply" is
         * pressed, exactly like the other pages.
         */
        const clearFilters = () => {
            FILTER_FIELDS.forEach(([, sel]) => { if (sel)
                sel.value = ""; });
            syncFilterUI();
            toast("info", "Filters cleared - press Apply to refresh the grid");
        };
        /* ---------- VIEW MODE ---------- */
        const paintViewMode = () => {
            const mode = currentViewMode;
            toggleRevenueBtn === null || toggleRevenueBtn === void 0 ? void 0 : toggleRevenueBtn.classList.toggle('toggle-btn-active', mode === 'revenue');
            toggleRevenueBtn === null || toggleRevenueBtn === void 0 ? void 0 : toggleRevenueBtn.classList.toggle('toggle-btn-inactive', mode !== 'revenue');
            toggleCollectionBtn === null || toggleCollectionBtn === void 0 ? void 0 : toggleCollectionBtn.classList.toggle('toggle-btn-active', mode === 'collection');
            toggleCollectionBtn === null || toggleCollectionBtn === void 0 ? void 0 : toggleCollectionBtn.classList.toggle('toggle-btn-inactive', mode !== 'collection');
            if (activeViewIndicator) {
                activeViewIndicator.textContent = mode === 'revenue' ? '📊 Revenue View' : '📋 Collection View';
            }
            // Concerns cells are only editable for editors — tell them so
            if (concernsHint)
                concernsHint.style.display = canEditConcerns() ? 'block' : 'none';
            refreshSyncBar();
        };
        /**
         * 🔐 RBAC SYNC BAR. Visible only to roles with create rights
         * (canCreateRecord → Admin / GH Head / PM Manager) and always follows the
         * active view: Revenue view syncs Project List → Revenue Forecast List,
         * Collection view syncs Project List → Collection List.
         */
        const refreshSyncBar = () => {
            const bar = q("#sprintSyncBar");
            if (!bar)
                return;
            const allowed = canCreateRecord();
            bar.style.display = allowed ? 'flex' : 'none';
            const hint = q("#sprintSyncHint");
            if (hint) {
                // hint.textContent = allowed
                //   ? (currentViewMode === 'revenue'
                //       ? 'Project List → Revenue Forecast List: any visible project without a forecast row is added automatically.'
                //       : 'Project List → Collection List: any visible project without a collection row is added automatically.')
                //   : '';
            }
        };
        const runSyncForCurrentView = async () => {
            const btn = q("#sprintSyncProjectsBtn");
            if (btn)
                btn.disabled = true;
            const syncLoading = createLoadingOverlay("Syncing projects...");
            loaderHost().appendChild(syncLoading);
            console.log("⏳ Syncing projects — the grid stays visible behind the loader bar");
            try {
                if (currentViewMode === 'revenue') {
                    await syncMissingProjectsToRevenueList();
                }
                else {
                    await syncMissingProjectsToCollectionList();
                }
                syncMonthSelection();
                filterTable();
            }
            catch (error) {
                console.error("❌ Sync failed:", error);
            }
            finally {
                if (btn)
                    btn.disabled = false;
                refreshSyncBar();
                removeLoadingOverlay();
            }
        };
        const refreshData = async () => {
            const loading = createLoadingOverlay(currentViewMode === 'revenue' ? 'Loading revenue data...' : 'Loading collection data...');
            loaderHost().appendChild(loading);
            console.log(currentViewMode === 'revenue'
                ? "⏳ Loading revenue data…"
                : "⏳ Loading collection data…");
            try {
                await (currentViewMode === 'revenue' ? loadRevenueForecasts() : loadCollectionItems());
                syncMonthSelection();
                filterTable();
            }
            finally {
                removeLoadingOverlay();
            }
        };
        const setViewMode = (mode) => {
            if (currentViewMode === mode)
                return;
            currentViewMode = mode;
            sortColumn = ''; // sort keys differ between the two layouts
            paintViewMode();
            renderMonthColumnsHeader(); // left column sequence differs per view
            void refreshData().then(() => toast('info', mode === 'revenue' ? 'Switched to Revenue View' : 'Switched to Collection View'));
        };
        /* ---------- EVENT BINDING ---------- */
        const bind = () => {
            toggleRevenueBtn === null || toggleRevenueBtn === void 0 ? void 0 : toggleRevenueBtn.addEventListener("click", () => setViewMode('revenue'));
            toggleCollectionBtn === null || toggleCollectionBtn === void 0 ? void 0 : toggleCollectionBtn.addEventListener("click", () => setViewMode('collection'));
            // filter popup: trigger / Apply pill / overlay / selects / footer buttons
            // (Apply-gated: the grid only re-filters when Apply is pressed)
            setupFilterEventListeners();
            void clearFiltersBtn; // queried by the button row above; wired in setupFilterEventListeners()
            // 🔐 Manual re-run of the auto-sync for the active view (button only
            // visible to create-role users). Console-only reporting.
            const sprintSyncProjectsBtn = q("#sprintSyncProjectsBtn");
            sprintSyncProjectsBtn === null || sprintSyncProjectsBtn === void 0 ? void 0 : sprintSyncProjectsBtn.addEventListener("click", () => {
                void runSyncForCurrentView();
            });
            pageSizeSelect === null || pageSizeSelect === void 0 ? void 0 : pageSizeSelect.addEventListener("change", (e) => {
                changePageSize(parseInt(e.target.value, 10));
            });
            firstPageBtn === null || firstPageBtn === void 0 ? void 0 : firstPageBtn.addEventListener("click", () => goToPage(1));
            prevPageBtn === null || prevPageBtn === void 0 ? void 0 : prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
            nextPageBtn === null || nextPageBtn === void 0 ? void 0 : nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
            lastPageBtn === null || lastPageBtn === void 0 ? void 0 : lastPageBtn.addEventListener("click", () => goToPage(totalPages));
            sprintMonthColumnsHeader === null || sprintMonthColumnsHeader === void 0 ? void 0 : sprintMonthColumnsHeader.addEventListener('click', (e) => {
                const btn = e.target.closest('.sprint-sort-btn');
                if (!btn)
                    return;
                const col = btn.dataset.sort;
                if (!col)
                    return;
                toggleSort(col, btn.dataset.direction || undefined);
            });
            /* ---- Concerns cell: edit + auto-save (delegated, survives re-renders) ---- */
            const editorFrom = (target) => { var _a; return (_a = target === null || target === void 0 ? void 0 : target.closest) === null || _a === void 0 ? void 0 : _a.call(target, '.sprint-concerns-editor'); };
            // entering a cell: close any previous editor first, then remember state
            sprintTableBody === null || sprintTableBody === void 0 ? void 0 : sprintTableBody.addEventListener('focusin', (e) => {
                var _a;
                const el = editorFrom(e.target);
                if (!el || (activeEditor === null || activeEditor === void 0 ? void 0 : activeEditor.el) === el)
                    return;
                commitActiveEditor(); // jumping cell → cell
                const itemId = Number(el.dataset.concernsId);
                if (!Number.isFinite(itemId))
                    return;
                activeEditor = { itemId, el, original: readEditorText(el) };
                (_a = el.closest('td')) === null || _a === void 0 ? void 0 : _a.classList.add('is-editing');
            });
            // leaving the cell = save
            sprintTableBody === null || sprintTableBody === void 0 ? void 0 : sprintTableBody.addEventListener('focusout', (e) => {
                var _a;
                const el = editorFrom(e.target);
                if (!el)
                    return;
                // ignore focus moving around inside the same cell (e.g. its scrollbar)
                const next = e.relatedTarget;
                if (next && ((_a = el.parentElement) === null || _a === void 0 ? void 0 : _a.contains(next)))
                    return;
                commitActiveEditor();
            });
            sprintTableBody === null || sprintTableBody === void 0 ? void 0 : sprintTableBody.addEventListener('keydown', (e) => {
                const el = editorFrom(e.target);
                if (!el)
                    return;
                if (e.key === 'Enter' && !e.shiftKey) { // Enter saves, Shift+Enter = newline
                    e.preventDefault();
                    el.blur();
                }
                else if (e.key === 'Escape') { // Esc discards
                    e.preventDefault();
                    el.dataset.cancelled = '1';
                    el.blur();
                }
            });
            sprintTableBody === null || sprintTableBody === void 0 ? void 0 : sprintTableBody.addEventListener('input', (e) => {
                const el = editorFrom(e.target);
                if (el)
                    el.classList.toggle('is-empty', !readEditorText(el).trim());
            });
            // paste as plain text only — never let HTML into a contenteditable
            sprintTableBody === null || sprintTableBody === void 0 ? void 0 : sprintTableBody.addEventListener('paste', (e) => {
                var _a;
                const el = editorFrom(e.target);
                if (!el)
                    return;
                e.preventDefault();
                const text = (((_a = e.clipboardData) === null || _a === void 0 ? void 0 : _a.getData('text/plain')) || '').replace(/\r\n/g, '\n');
                const selection = window.getSelection();
                if (selection && selection.rangeCount) {
                    const range = selection.getRangeAt(0);
                    range.deleteContents();
                    range.insertNode(document.createTextNode(text));
                    range.collapse(false);
                    selection.removeAllRanges();
                    selection.addRange(range);
                }
                else {
                    el.textContent = (el.textContent || '') + text;
                }
                el.classList.toggle('is-empty', !readEditorText(el).trim());
            });
        };
        /* ---------- INIT ---------- */
        const initialize = async () => {
            const loading = createLoadingOverlay("Loading sprint data...");
            loaderHost().appendChild(loading);
            try {
                // ✅ the loader is now a pure bar: no spinner, no text, no overlay
                // message - every status line below is logged to the console instead
                syncMonthSelection();
                paintViewMode();
                renderMonthColumnsHeader();
                bind();
                // paint the totals band - and with it the filter trigger - before the
                // first data render, so the strip is in place even if a list load fails
                renderTotalsHeader();
                // 🔧 Always re-fetch from the database on every mount - the old
                // cached-data shortcut re-rendered the in-memory copy and could show
                // stale actuals until a hard refresh. No early return: fall through
                // to the full SharePoint load below (same as a fresh page load).
                if (remount && dataLoaded) {
                    console.log("🔄 Daily Sprint page re-mounted - forcing a fresh load from SharePoint");
                }
                updateLoadingMessage(loading, "Initializing permissions...");
                console.log("⏳ Initializing permissions…");
                await authManager.initialize({ siteUrl });
                console.log(`👤 ${getCurrentUserFullName()} | ${getPermissionLevel().toUpperCase()} | [${getUserGroups().join(", ")}]`);
                updateLoadingMessage(loading, "Loading projects...");
                console.log("⏳ Loading projects…");
                await loadProjects();
                updateLoadingMessage(loading, "Loading forecast data...");
                console.log("⏳ Loading forecast data…");
                await Promise.all([loadRevenueForecasts(), loadCollectionItems()]);
                // 🔐 AUTO-SYNC (same as the Revenue & Collection pages): add a row for
                // every visible project missing from each list - empty skeleton JSON,
                // idempotent, never duplicates/overwrites. Only roles with create
                // rights run it; reporting is console-only.
                updateLoadingMessage(loading, "Syncing missing projects...");
                console.log("⏳ Syncing missing projects…");
                await syncMissingProjectsToRevenueList();
                await syncMissingProjectsToCollectionList();
                // show the manual Sync button only to roles that can create rows
                refreshSyncBar();
                dataLoaded = true;
                initializeFilters();
                syncMonthSelection();
                filterTable();
                console.log("✅ Daily Sprint page initialized — horizon:", horizonKeys.join(', '));
            }
            catch (error) {
                console.error("❌ Failed to initialize sprint page:", error);
                toast("error", "Failed to initialize page");
            }
            finally {
                removeLoadingOverlay();
            }
        };
        // keep sticky offsets accurate after layout / resize / content changes
        window.requestAnimationFrame(() => {
            applyStickyHeaderOffsets();
            teardownWhenDetached();
        });
        // If the SPA swaps pages, undo our ancestor tweaks so other pages are
        // unaffected. (SPFx removes the web part node from the DOM.)
        const teardownWhenDetached = () => {
            const host = root;
            if (!host || !host.isConnected)
                return;
            const obs = new MutationObserver(() => {
                if (!host.isConnected) {
                    restoreAncestors();
                    obs.disconnect();
                }
            });
            obs.observe(document.body, { childList: true, subtree: true });
        };
        const scheduleStickyRecalc = () => {
            clearTimeout(window._sprintStickyResizeTimer);
            window._sprintStickyResizeTimer = setTimeout(applyStickyHeaderOffsets, 120);
        };
        window.addEventListener('resize', scheduleStickyRecalc);
        // app chrome may collapse/expand as the page scrolls — re-measure lazily
        window.addEventListener('scroll', scheduleStickyRecalc, { passive: true });
        // SPFx often scrolls an inner region rather than the window
        document.addEventListener('scroll', scheduleStickyRecalc, { passive: true, capture: true });
        void initialize();
    }
};
export default sprintPage;
export { sprintPage };
//# sourceMappingURL=dailysprint.js.map
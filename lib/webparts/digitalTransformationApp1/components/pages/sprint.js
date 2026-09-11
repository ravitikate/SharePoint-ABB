import "../styles/common.css";
import "../styles/dailySprint.css";
import Swal from "sweetalert2";
import { authManager, hasFullControl, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canCreateRecord, getAssignedProjects, updateUIBasedOnPermissions, getDigest as getAuthDigest, getListEntityType } from "./auth";
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const NUMERIC_FIELDS = ["Plan", "Actual", "Further"];
let currentViewMode = 'revenue';
let activeMonthKey = '';
let editingId = null;
let editingType = null;
// All three flags are module-level on purpose: `script()` runs again every time the user
// navigates back to this page, and the web part only swaps child markup - listeners
// attached to the same long-lived nodes would stack up (duplicate toasts, double saves).
let pageMounted = false;
let dataLoaded = false;
// NOTE: no "already bound" guard. loadPage() rewrites .dt-body.innerHTML, so every mount
// hands us brand-new nodes and bind() must attach again. What must NOT be repeated is the
// network/auth work - that is guarded by `dataLoaded`.
let projects = [];
let revenueForecasts = [];
let collectionItems = [];
const siteUrl = typeof _spPageContextInfo !== "undefined"
    ? _spPageContextInfo.webAbsoluteUrl
    : window.location.origin + "/sites/DigitalTransformationTest";
const revenueListName = "Revenue Forecast List";
const collectionListName = "Collection List";
const projectListName = "Project List";
/* ============================================================
   PURE HELPERS (no DOM, safe anywhere)
   ============================================================ */
const toast = (icon, title) => {
    void Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
};
const getDigest = async () => getAuthDigest(siteUrl);
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
const escapeAttr = (value) => value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const dataField = (mode) => mode === 'revenue' ? 'ForecastData' : 'CollectionData';
const getNumericValue = (input) => {
    if (!input)
        return 0;
    const value = input.value.trim();
    if (value === '' || value === '0')
        return 0;
    return parseInt(value, 10) || 0;
};
const setNumericValue = (input, value) => {
    if (!input)
        return;
    if (value === 0) {
        input.value = '';
        input.classList.remove('sprint-input-with-value');
        input.classList.add('sprint-input-empty');
    }
    else {
        input.value = String(value);
        input.classList.remove('sprint-input-empty');
        input.classList.add('sprint-input-with-value');
    }
};
const splitRemarks = (remarks) => {
    const raw = remarks || "";
    if (raw.includes('|||')) {
        const parts = raw.split('|||');
        return { internal: parts[0] || "", sprintCell: parts[1] || "" };
    }
    return { internal: raw, sprintCell: "" };
};
const detailRows = (p) => {
    if (!p) {
        return `<div class="sprint-detail-item"><span class="sprint-detail-label">Project:</span><span class="sprint-detail-value">-</span></div>`;
    }
    const rows = [
        ["Project", p.Title || '-'],
        ["Project Ref", p.ProjectReference || '-'],
        ["Profit Center", p.PC || '-'],
        ["Project Manager", getProjectManagerDisplayName(p.ProjectManager) || '-'],
        ["Group Head", getGroupHeadDisplayName(p.GroupHead1) || '-'],
        ["Customer", p.CustomerName || '-'],
        ["Engg Lead", getEnggLeadDisplayName(p.EnggLead) || '-']
    ];
    return rows
        .map(([label, value]) => `<div class="sprint-detail-item"><span class="sprint-detail-label">${label}:</span><span class="sprint-detail-value">${value}</span></div>`)
        .join('');
};
/* ============================================================
   CONTENT (static markup only — no ${} that needs runtime data)
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
      <button type="button" class="daily-sprint-show-form-btn daily-sprint-btn-primary">+ Add Entry</button>
    </div>
  </div>

  <div class="daily-sprint-table-section">

    <div class="daily-sprint-filters">
      <label class="daily-sprint-filter">Year
        <select id="sprintFilterYear"><option value="">All Years</option></select>
      </label>
      <label class="daily-sprint-filter">Month
        <select id="sprintFilterMonth"><option value="">All Months</option></select>
      </label>
      <label class="daily-sprint-filter">Project
        <select id="sprintFilterProject"><option value="">All Projects</option></select>
      </label>
      <label class="daily-sprint-filter">Customer
        <select id="sprintFilterCustomer"><option value="">All Customers</option></select>
      </label>
      <label class="daily-sprint-filter">Project Manager
        <select id="sprintFilterProjectManager"><option value="">All Managers</option></select>
      </label>
      <label class="daily-sprint-filter">Group Head
        <select id="sprintFilterGroupHead"><option value="">All Group Heads</option></select>
      </label>
      <label class="daily-sprint-filter">Project Reference
        <select id="sprintFilterProjectReference"><option value="">All References</option></select>
      </label>
      <label class="daily-sprint-filter">Engg Lead
        <select id="sprintFilterEnggLead"><option value="">All Engg Leads</option></select>
      </label>
      <button type="button" class="daily-sprint-clear-filters-btn">Clear filters</button>
    </div>

    <div class="sprint-data-container">
      <table class="daily-sprint-table">
        <thead>
          <tr>
            <th>View</th>
            <th>Project Ref</th>
            <th>Project</th>
            <th>Customer</th>
            <th>Group Head</th>
            <th>Project Manager</th>
            <th>Engg Lead</th>
            <th>Plan</th>
            <th>Actual</th>
            <th>Further</th>
            <th>Forecast</th>
            <th>Sprint remarks</th>
          </tr>
        </thead>
        <tbody class="sprint-table-body">
          <tr><td colspan="12" class="daily-sprint-empty">Loading data…</td></tr>
        </tbody>
      </table>
      <button id="loadMoreBtn" type="button" class="daily-sprint-load-more">Load more</button>
    </div>
  </div>

  <div class="daily-sprint-form-section" style="display:none;">
    <h2 id="sprintFormTitle">Add New Revenue Forecast</h2>

    <form id="dailySprintForm">
      <div class="daily-sprint-form-grid">
        <label class="daily-sprint-field">Project
          <select id="sprintProjectSelect" required>
            <option value="">Select Project</option>
          </select>
        </label>
        <div id="sprintProjectDetails" class="sprint-project-details"></div>
      </div>

      <div class="daily-sprint-month-grid-wrap">
        <h3>Month values (<span class="sprint-active-month-label">this month</span>) — in ₹</h3>
        <table class="sprint-month-grid">
          <thead>
            <tr>
              <th>Month</th><th>Plan</th><th>Actual</th><th>Further Forecast</th>
            </tr>
          </thead>
          <tbody>
            ${MONTHS.map((month, i) => `
            <tr>
              <td>${month}</td>
              ${NUMERIC_FIELDS.map(field => `
                <td><input type="number" min="0" step="1" class="sprint-num-input sprint-input-empty"
                     id="sprintRevenue${month}${field}" data-month="${pad2(i + 1)}" data-field="${field.toLowerCase()}" /></td>
              `).join('')}
            </tr>`).join('')}
          </tbody>
        </table>
        <p class="daily-sprint-hint">Only the highlighted month is written to the list item; the other months are preserved.</p>
      </div>

      <div class="daily-sprint-form-grid">
        <label class="daily-sprint-field">Internal remarks
          <textarea id="sprintRemarksInternal" rows="3" placeholder="Visible to the team only"></textarea>
        </label>
        <label class="daily-sprint-field">Sprint cell remarks
          <textarea id="sprintRemarksSprintCell" rows="3" placeholder="Shown in the table's remarks column"></textarea>
        </label>
      </div>

      <div class="daily-sprint-form-actions">
        <button type="submit" class="daily-sprint-btn-primary">Save</button>
        <button type="button" class="daily-sprint-btn-secondary daily-sprint-back-to-table">Back to table</button>
      </div>
    </form>
  </div>
</div>
`;
/* ============================================================
   STYLE (inlined so the page still looks right if the global
   CSS import is not wired into the web part)
   ============================================================ */
const style = `
.daily-sprint-page { font-family:'Segoe UI',system-ui,sans-serif; color:#1f2328; }
.daily-sprint-header { display:flex; align-items:center; justify-content:space-between; gap:16px; margin-bottom:14px; flex-wrap:wrap; }
.daily-sprint-title { display:flex; align-items:baseline; gap:12px; }
.daily-sprint-title h1 { margin:0; font-size:22px; font-weight:600; }
.daily-sprint-view-indicator { font-size:12px; color:#0b6bcb; background:#eaf3fd; border:1px solid #cfe4fb; border-radius:12px; padding:2px 10px; }
.daily-sprint-header-actions { display:flex; align-items:center; gap:10px; }
.daily-sprint-view-toggle { display:inline-flex; border:1px solid #d0d7de; border-radius:6px; overflow:hidden; }
.daily-sprint-toggle-btn { border:0; background:#fff; padding:7px 14px; font-size:13px; cursor:pointer; }
.toggle-btn-active { background:#0078d4; color:#fff; }
.toggle-btn-inactive { background:#fff; color:#57606a; }
.daily-sprint-btn-primary { background:#0078d4; color:#fff; border:0; border-radius:4px; padding:8px 16px; font-size:13px; cursor:pointer; }
.daily-sprint-btn-primary:hover { background:#106ebe; }
.daily-sprint-btn-secondary { background:#fff; color:#333; border:1px solid #d0d7de; border-radius:4px; padding:8px 16px; font-size:13px; cursor:pointer; }
.daily-sprint-filters { display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end; background:#fff; border:1px solid #eaeef2; border-radius:8px; padding:12px; margin-bottom:12px; }
.daily-sprint-filter { display:flex; flex-direction:column; gap:4px; font-size:11px; color:#57606a; text-transform:uppercase; letter-spacing:.03em; }
.daily-sprint-filter select { min-width:130px; padding:6px 8px; font-size:13px; border:1px solid #d0d7de; border-radius:4px; background:#fff; color:#1f2328; }
.daily-sprint-clear-filters-btn { padding:7px 12px; font-size:12px; border:1px solid #d0d7de; background:#f6f8fa; border-radius:4px; cursor:pointer; }
.sprint-data-container { background:#fff; border:1px solid #eaeef2; border-radius:8px; padding:12px; overflow-x:auto; }
.daily-sprint-table { width:100%; border-collapse:collapse; font-size:12px; }
.daily-sprint-table th { text-align:left; padding:8px 10px; border-bottom:2px solid #eaeef2; color:#57606a; font-weight:600; white-space:nowrap; }
.daily-sprint-table td { padding:10px 8px; border-bottom:1px solid #f0f2f4; vertical-align:middle; }
.daily-sprint-empty { text-align:center; padding:20px; color:#6c757d; }
.daily-sprint-load-more { display:block; width:100%; margin-top:10px; padding:8px; border:1px dashed #d0d7de; background:#fbfcfd; color:#57606a; border-radius:4px; font-size:12px; cursor:pointer; }
.daily-sprint-form-section { background:#fff; border:1px solid #eaeef2; border-radius:8px; padding:16px; margin-top:4px; }
.daily-sprint-form-section h2 { margin:0 0 14px; font-size:17px; font-weight:600; }
.daily-sprint-form-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(280px,1fr)); gap:14px; margin-bottom:14px; }
.daily-sprint-field { display:flex; flex-direction:column; gap:6px; font-size:12px; font-weight:600; color:#57606a; }
.daily-sprint-field select, .daily-sprint-field textarea { padding:8px; font-size:13px; font-weight:400; color:#1f2328; border:1px solid #d0d7de; border-radius:4px; font-family:inherit; }
.sprint-project-details { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:8px; background:#f6f8fa; border:1px solid #eaeef2; border-radius:6px; padding:10px; align-self:start; }
.sprint-detail-item { display:flex; gap:6px; font-size:12px; }
.sprint-detail-label { color:#6b7280; min-width:92px; }
.sprint-detail-value { color:#1f2328; font-weight:600; }
.daily-sprint-month-grid-wrap { margin-bottom:14px; }
.daily-sprint-month-grid-wrap h3 { margin:0 0 8px; font-size:13px; color:#57606a; font-weight:600; }
.sprint-month-grid { width:100%; border-collapse:collapse; font-size:12px; }
.sprint-month-grid th { text-align:left; padding:6px 8px; background:#f6f8fa; border-bottom:1px solid #eaeef2; color:#57606a; font-weight:600; }
.sprint-month-grid td { padding:5px 8px; border-bottom:1px solid #f0f2f4; }
.sprint-month-grid tr.is-active { background:#f0f8ff; }
.sprint-num-input { width:110px; padding:5px 7px; font-size:12px; border:1px solid #d0d7de; border-radius:4px; }
.sprint-input-empty { background:#fff; }
.sprint-input-with-value { background:#fffbe6; border-color:#d4a72c; }
.daily-sprint-form-actions { display:flex; gap:10px; }
.daily-sprint-hint { margin:6px 0 0; font-size:11px; color:#8c959f; }
.sprint-edit-trigger { color:#0078d4; cursor:pointer; text-decoration:underline; }
.daily-sprint-badge { display:inline-block; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:600; background:#0078d4; color:#fff; }
.daily-sprint-badge.collection { background:#6f42c1; }
.daily-sprint-pill { display:inline-block; padding:2px 8px; border-radius:4px; font-size:11px; background:#e3f2fd; color:#1565c0; }
.daily-sprint-table td.num { text-align:right; font-variant-numeric:tabular-nums; }
.daily-sprint-table td.plan { font-weight:600; color:#0078d4; }
.daily-sprint-table td.actual { font-weight:700; color:#2e7d32; background:#e8f5e9; }
.daily-sprint-table td.forecast { font-weight:700; color:#0c5460; background:#f0f8ff; }
.daily-sprint-table td.remarks { max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:#6c757d; font-style:italic; }
.sprint-month-grid tr.is-dirty td { background:#fffdf5; }
.daily-sprint-table tbody tr:hover { background:#f9fbfd; }
`;
/* ============================================================
   DATA LOADING
   ============================================================ */
const loadProjects = async () => {
    var _a;
    try {
        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items` +
            `?$select=Id,Title,CustomerName,GroupHead1/Title,GroupHead1Id,PC,ProjectReference,ProjectManager/Title,ProjectManagerId,EnggLead/Title,EnggLeadId` +
            `&$expand=GroupHead1,ProjectManager,EnggLead`, { headers: { Accept: "application/json;odata=verbose" } });
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
const loadRevenueForecasts = async () => {
    var _a;
    try {
        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${revenueListName}')/items` +
            `?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,RevenueActual,Remarks,Created,Modified,Freeze` +
            `&$expand=Project&$top=5000`, { headers: { Accept: "application/json;odata=verbose" } });
        const data = await res.json();
        revenueForecasts = normalizeRecords(((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [], ['ForecastData', 'RevenueActual']);
        console.log("✅ Revenue forecasts loaded:", revenueForecasts.length);
    }
    catch (error) {
        console.error("Error loading revenue forecasts:", error);
        toast("error", "Failed to load revenue data");
    }
};
const loadCollectionItems = async () => {
    var _a;
    try {
        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${collectionListName}')/items` +
            `?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,CollectionData,RevenueActual,Remarks,Created,Modified,Freeze,NormalDue,NormalOverdue,RetentionDue,RetentionOverdue,GSTTDS,WHT` +
            `&$expand=Project&$top=5000`, { headers: { Accept: "application/json;odata=verbose" } });
        const data = await res.json();
        collectionItems = normalizeRecords(((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [], ['CollectionData', 'RevenueActual']);
        console.log("✅ Collection items loaded:", collectionItems.length);
    }
    catch (error) {
        console.error("Error loading collection items:", error);
        toast("error", "Failed to load collection data");
    }
};
const currentRecords = () => currentViewMode === 'revenue' ? revenueForecasts : collectionItems;
/* ============================================================
   PAGE MODULE
   ============================================================ */
const sprintPage = {
    content,
    style,
    // Called by the web part AFTER `content` has been written into the DOM.
    script: () => {
        // Fresh mount: drop any half-finished edit, keep the data cache.
        editingId = null;
        editingType = null;
        const remount = pageMounted; // true when we've been here before in this tab
        pageMounted = true;
        /* ---------- DOM LOOKUPS: now, not at import time ---------- */
        const root = (document.querySelector('.daily-sprint-page') || document);
        const q = (sel) => root.querySelector(sel);
        const qa = (sel) => Array.from(root.querySelectorAll(sel));
        const pageHeader = q(".daily-sprint-header");
        const toggleRevenueBtn = q("#toggleRevenue");
        const toggleCollectionBtn = q("#toggleCollection");
        const activeViewIndicator = q("#activeViewIndicator");
        const loadMoreBtn = q("#loadMoreBtn");
        const sprintTableBody = q(".sprint-table-body");
        const formSection = q(".daily-sprint-form-section");
        const tableSection = q(".daily-sprint-table-section");
        const showFormBtn = q(".daily-sprint-show-form-btn");
        const projectSelect = q("#sprintProjectSelect");
        const projectDetails = q("#sprintProjectDetails");
        const backBtn = q(".daily-sprint-back-to-table");
        const form = q("#dailySprintForm");
        const remarksInternal = q("#sprintRemarksInternal");
        const remarksSprintCell = q("#sprintRemarksSprintCell");
        const formTitle = q("#sprintFormTitle");
        const activeMonthLabel = q(".sprint-active-month-label");
        const monthGridRows = qa(".sprint-month-grid tbody tr");
        const filterYear = q("#sprintFilterYear");
        const filterMonth = q("#sprintFilterMonth");
        const filterProject = q("#sprintFilterProject");
        const filterCustomer = q("#sprintFilterCustomer");
        const filterProjectManager = q("#sprintFilterProjectManager");
        const filterGroupHead = q("#sprintFilterGroupHead");
        const filterProjectReference = q("#sprintFilterProjectReference");
        const filterEnggLead = q("#sprintFilterEnggLead");
        const clearFiltersBtn = q(".daily-sprint-clear-filters-btn");
        const missing = [];
        const requireEls = [
            ["#dailySprintForm", form], ["#sprintProjectSelect", projectSelect],
            [".sprint-table-body", sprintTableBody], ["#toggleRevenue", toggleRevenueBtn],
            ["#toggleCollection", toggleCollectionBtn], ["#sprintFilterYear", filterYear]
        ];
        requireEls.forEach(([sel, el]) => { if (!el)
            missing.push(sel); });
        if (missing.length) {
            console.error(`❌ sprint page: markup is missing required element(s): ${missing.join(", ")}`);
        }
        const activeMonthNum = () => activeMonthKey.slice(5);
        const inputFor = (field) => q(`.sprint-num-input[data-month="${activeMonthNum()}"][data-field="${field}"]`);
        /* ---------- MONTH GRID ---------- */
        const highlightActiveMonth = () => {
            const active = MONTHS[parseInt(activeMonthNum(), 10) - 1];
            monthGridRows.forEach(row => {
                var _a;
                const cell = row.querySelector('td');
                row.classList.toggle('is-active', !!cell && ((_a = cell.textContent) === null || _a === void 0 ? void 0 : _a.trim()) === active);
            });
            if (activeMonthLabel) {
                activeMonthLabel.textContent = activeMonthKey
                    ? `${MONTHS[parseInt(activeMonthKey.slice(5), 10) - 1]} ${activeMonthKey.slice(0, 4)}`
                    : "this month";
            }
        };
        const clearMonthGrid = () => {
            qa('.sprint-num-input').forEach(el => setNumericValue(el, 0));
        };
        const readMonthValues = () => ({
            plan: getNumericValue(inputFor('plan')),
            actual: getNumericValue(inputFor('actual')),
            forecast: getNumericValue(inputFor('further'))
        });
        /** Pre-fills the active month's row from the record so an edit never wipes other months. */
        const ensureMonthGrid = (record) => {
            clearMonthGrid();
            if (!record || !activeMonthKey)
                return;
            const data = safeJsonParse(record[dataField(currentViewMode)]);
            const entry = data[activeMonthKey] || { plan: 0, actual: 0, forecast: 0 };
            setNumericValue(inputFor('plan'), entry.plan || 0);
            setNumericValue(inputFor('actual'), entry.actual || 0);
            setNumericValue(inputFor('further'), entry.forecast || 0);
        };
        /** Merges the active month into the existing map (all other months preserved). */
        const mergeMonthData = (existing) => {
            const data = safeJsonParse(existing);
            data[activeMonthKey] = readMonthValues();
            return JSON.stringify(data);
        };
        /* ---------- TABLE ---------- */
        const renderSprintTable = (rows) => {
            if (!sprintTableBody)
                return;
            const dataToRender = rows || currentRecords();
            if (dataToRender.length === 0) {
                sprintTableBody.innerHTML = `<tr><td colspan="12" class="daily-sprint-empty">No data found</td></tr>`;
                return;
            }
            sprintTableBody.innerHTML = dataToRender.map(item => {
                var _a;
                const projectTitle = ((_a = item.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                const proj = projects.find(p => p.Title === projectTitle);
                const projectRef = getSafeString(proj === null || proj === void 0 ? void 0 : proj.ProjectReference);
                const customerName = getSafeString(proj === null || proj === void 0 ? void 0 : proj.CustomerName);
                const groupHeadName = getGroupHeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.GroupHead1) || "-";
                const pmName = getProjectManagerDisplayName(proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-";
                const enggLeadName = getEnggLeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.EnggLead);
                const { sprintCell } = splitRemarks(item.Remarks);
                const monthData = safeJsonParse(item[dataField(currentViewMode)])[activeMonthKey] || { plan: 0, actual: 0, forecast: 0 };
                const plan = monthData.plan || 0;
                const actual = monthData.actual || 0;
                const further = monthData.forecast || 0;
                const forecast = actual + further;
                const inr = (n) => n.toLocaleString('en-IN');
                return `
          <tr data-item-id="${item.Id}" style="border-bottom:1px solid #e9ecef;">
            <td><span class="daily-sprint-badge ${currentViewMode}">${currentViewMode === 'revenue' ? 'R' : 'C'}</span></td>
            <td><span class="sprint-edit-trigger" data-id="${item.Id}" data-type="${currentViewMode}">${projectRef}</span></td>
            <td style="font-weight:600;color:#333;">${escapeAttr(projectTitle)}</td>
            <td>${escapeAttr(customerName)}</td>
            <td>${escapeAttr(groupHeadName)}</td>
            <td>${escapeAttr(pmName)}</td>
            <td><span class="daily-sprint-pill">${escapeAttr(enggLeadName)}</span></td>
            <td class="num plan">${inr(plan)}</td>
            <td class="num actual">${inr(actual)}</td>
            <td class="num">${inr(further)}</td>
            <td class="num forecast">${inr(forecast)}</td>
            <td class="remarks" title="${escapeAttr(sprintCell)}">💬 ${escapeAttr(sprintCell || '-')}</td>
          </tr>`;
            }).join('');
            sprintTableBody.querySelectorAll('.sprint-edit-trigger').forEach(link => {
                link.addEventListener('click', (e) => {
                    e.preventDefault();
                    const el = e.currentTarget;
                    const id = parseInt(el.dataset.id || '0', 10);
                    const type = (el.dataset.type || currentViewMode);
                    void editItem(id, type);
                });
            });
        };
        /* ---------- FILTERS ---------- */
        const selectedOr = (sel, fallback) => (sel && sel.value) ? sel.value : fallback;
        const syncMonthSelection = () => {
            const year = selectedOr(filterYear, String(new Date().getFullYear()));
            const month = pad2(selectedOr(filterMonth, String(new Date().getMonth() + 1)));
            activeMonthKey = `${year}-${month}`;
            highlightActiveMonth();
        };
        const initializeFilters = () => {
            if (!filterYear || !filterMonth)
                return;
            const currentYear = new Date().getFullYear();
            filterYear.innerHTML = '<option value="">All Years</option>' +
                Array.from({ length: 4 }, (_, i) => currentYear - 1 + i)
                    .map(y => `<option value="${y}"${y === currentYear ? ' selected' : ''}>${y}</option>`)
                    .join('');
            filterMonth.innerHTML = '<option value="">All Months</option>' +
                MONTHS.map((m, i) => `<option value="${i + 1}"${i + 1 === new Date().getMonth() + 1 ? ' selected' : ''}>${m}</option>`)
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
        };
        const filterTable = () => {
            syncMonthSelection();
            const year = (filterYear === null || filterYear === void 0 ? void 0 : filterYear.value) || "";
            const month = (filterMonth === null || filterMonth === void 0 ? void 0 : filterMonth.value) || "";
            const project = (filterProject === null || filterProject === void 0 ? void 0 : filterProject.value) || "";
            const customer = (filterCustomer === null || filterCustomer === void 0 ? void 0 : filterCustomer.value) || "";
            const manager = (filterProjectManager === null || filterProjectManager === void 0 ? void 0 : filterProjectManager.value) || "";
            const head = (filterGroupHead === null || filterGroupHead === void 0 ? void 0 : filterGroupHead.value) || "";
            const reference = (filterProjectReference === null || filterProjectReference === void 0 ? void 0 : filterProjectReference.value) || "";
            const lead = (filterEnggLead === null || filterEnggLead === void 0 ? void 0 : filterEnggLead.value) || "";
            const matches = (item) => {
                var _a;
                const projectTitle = (_a = item.Project) === null || _a === void 0 ? void 0 : _a.Title;
                if (!projectTitle)
                    return false;
                if (year && month) {
                    const data = safeJsonParse(item[dataField(currentViewMode)]);
                    if (!(activeMonthKey in data))
                        return false;
                }
                if (project && projectTitle !== project)
                    return false;
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
            renderSprintTable(currentRecords().filter(matches));
        };
        const clearFilters = () => {
            [filterYear, filterMonth, filterProject, filterCustomer, filterProjectManager, filterGroupHead, filterProjectReference, filterEnggLead]
                .forEach(sel => { if (sel)
                sel.value = ""; });
            filterTable();
            toast("info", "All filters cleared");
        };
        /* ---------- FORM ---------- */
        const showForm = (isEditMode, itemType) => {
            if (formSection)
                formSection.style.display = "block";
            if (tableSection)
                tableSection.style.display = "none";
            if (showFormBtn)
                showFormBtn.style.display = "none";
            editingType = itemType;
            if (formTitle) {
                formTitle.textContent = `${isEditMode ? 'Edit' : 'Add New'} ${itemType === 'revenue' ? 'Revenue' : 'Collection'} Forecast`;
            }
            if (projectSelect) {
                const assignedProjects = getAssignedProjects(projects);
                projectSelect.innerHTML = `<option value="">Select Project</option>` +
                    assignedProjects.map(p => `<option value="${p.Id}">${escapeAttr(p.Title)}</option>`).join('');
            }
            if (isEditMode)
                highlightActiveMonth();
            window.setTimeout(() => formSection === null || formSection === void 0 ? void 0 : formSection.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
        };
        const hideForm = () => {
            if (formSection)
                formSection.style.display = "none";
            if (tableSection)
                tableSection.style.display = "block";
            if (showFormBtn)
                showFormBtn.style.display = "inline-block";
            editingType = null;
        };
        const resetForm = () => {
            editingId = null;
            editingType = null;
            form === null || form === void 0 ? void 0 : form.reset();
            if (projectSelect)
                projectSelect.value = "";
            if (projectDetails)
                projectDetails.innerHTML = detailRows(undefined);
            if (remarksInternal)
                remarksInternal.value = "";
            if (remarksSprintCell)
                remarksSprintCell.value = "";
            clearMonthGrid();
        };
        const editItem = async (id, itemType) => {
            editingId = id;
            editingType = itemType;
            const source = itemType === 'revenue' ? revenueForecasts : collectionItems;
            const item = source.find(r => r.Id === id);
            if (!item) {
                toast("error", "Item not found");
                return;
            }
            if (!canRead()) {
                toast("error", "❌ You don't have permission to view this item");
                return;
            }
            const assignedProjects = getAssignedProjects(projects);
            const isAssigned = assignedProjects.some(p => { var _a; return p.Id === ((_a = item.Project) === null || _a === void 0 ? void 0 : _a.Id); });
            if (!hasFullControl() && !isGHHead() && !isAssigned) {
                toast("error", "❌ You don't have permission to view this project");
                return;
            }
            const project = projects.find(p => { var _a; return p.Id === ((_a = item.Project) === null || _a === void 0 ? void 0 : _a.Id); });
            showForm(true, itemType);
            if (projectSelect && project)
                projectSelect.value = String(project.Id);
            if (projectDetails)
                projectDetails.innerHTML = detailRows(project);
            if (remarksInternal)
                remarksInternal.value = splitRemarks(item.Remarks).internal;
            if (remarksSprintCell)
                remarksSprintCell.value = splitRemarks(item.Remarks).sprintCell;
            ensureMonthGrid(item);
            toast("info", `Editing ${itemType === 'revenue' ? 'Revenue' : 'Collection'} forecast`);
        };
        /* ---------- SAVE ---------- */
        const saveRecord = async (e) => {
            e.preventDefault();
            const projectId = parseInt((projectSelect === null || projectSelect === void 0 ? void 0 : projectSelect.value) || '0', 10);
            const project = projects.find(p => p.Id === projectId);
            if (!project) {
                toast("error", "Please select a project");
                return;
            }
            const isEditMode = editingId !== null;
            const itemType = editingType || currentViewMode;
            const source = itemType === 'revenue' ? revenueForecasts : collectionItems;
            let existing;
            if (isEditMode) {
                existing = source.find(r => r.Id === editingId);
                if (!existing) {
                    toast("error", "Item not found");
                    return;
                }
                if (existing.Freeze && !canFreeze()) {
                    toast("error", "❌ This record is frozen; you cannot edit it");
                    return;
                }
            }
            else if (!canCreateRecord()) {
                toast("error", "❌ You don't have permission to create records");
                return;
            }
            try {
                const listName = itemType === 'revenue' ? revenueListName : collectionListName;
                const listItemType = await getListEntityType(siteUrl, listName);
                const digest = await getDigest();
                const field = dataField(itemType);
                const combinedRemarks = `${((remarksInternal === null || remarksInternal === void 0 ? void 0 : remarksInternal.value) || '').trim()}|||${((remarksSprintCell === null || remarksSprintCell === void 0 ? void 0 : remarksSprintCell.value) || '').trim()}`;
                const payload = {
                    __metadata: { type: listItemType },
                    Title: `${project.Title} - ${new Date().getFullYear()} ${itemType === 'revenue' ? 'Revenue' : 'Collection'}`,
                    ProjectId: String(project.Id),
                    ProfitCentre: project.PC || "",
                    Remarks: combinedRemarks,
                    [field]: mergeMonthData(existing === null || existing === void 0 ? void 0 : existing[field]),
                    Modified: new Date().toISOString()
                };
                if (!isEditMode)
                    payload.Created = new Date().toISOString();
                const url = isEditMode
                    ? `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${editingId})`
                    : `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
                const headers = {
                    Accept: "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": digest
                };
                if (isEditMode) {
                    headers["IF-MATCH"] = "*";
                    headers["X-HTTP-Method"] = "MERGE";
                }
                const res = await fetch(url, {
                    method: "POST",
                    headers,
                    body: JSON.stringify(payload)
                });
                if (!res.ok) {
                    const detail = await res.text().catch(() => "");
                    console.error("Save failed:", res.status, detail);
                    toast("error", `Error saving: ${res.status}`);
                    return;
                }
                toast("success", `${itemType === 'revenue' ? 'Revenue' : 'Collection'} ${isEditMode ? 'Updated' : 'Added'}!`);
                await refreshData();
                hideForm();
                resetForm();
            }
            catch (error) {
                console.error("Error saving:", error);
                toast("error", "Unexpected error while saving!");
            }
        };
        /* ---------- VIEW MODE ---------- */
        /**
         * Paints the toggle/indicator from `currentViewMode`. Must run on EVERY mount:
         * the module keeps `currentViewMode` in memory, but a re-mount hands us markup whose
         * buttons default to "revenue" - without this, state and UI disagree and the first
         * click on the already-active button is swallowed by the equality guard below.
         */
        const paintViewMode = () => {
            const mode = currentViewMode;
            toggleRevenueBtn === null || toggleRevenueBtn === void 0 ? void 0 : toggleRevenueBtn.classList.toggle('toggle-btn-active', mode === 'revenue');
            toggleRevenueBtn === null || toggleRevenueBtn === void 0 ? void 0 : toggleRevenueBtn.classList.toggle('toggle-btn-inactive', mode !== 'revenue');
            toggleCollectionBtn === null || toggleCollectionBtn === void 0 ? void 0 : toggleCollectionBtn.classList.toggle('toggle-btn-active', mode === 'collection');
            toggleCollectionBtn === null || toggleCollectionBtn === void 0 ? void 0 : toggleCollectionBtn.classList.toggle('toggle-btn-inactive', mode !== 'collection');
            if (activeViewIndicator) {
                activeViewIndicator.textContent = mode === 'revenue' ? '📊 Revenue View' : '📋 Collection View';
            }
        };
        const refreshData = async () => {
            await (currentViewMode === 'revenue' ? loadRevenueForecasts() : loadCollectionItems());
            syncMonthSelection();
            renderSprintTable();
        };
        const setViewMode = (mode) => {
            if (currentViewMode === mode)
                return;
            currentViewMode = mode;
            paintViewMode();
            if (editingId !== null)
                hideForm();
            void refreshData().then(() => toast('info', mode === 'revenue' ? 'Switched to Revenue View' : 'Switched to Collection View'));
        };
        /* ---------- EVENT BINDING (once, after DOM exists) ---------- */
        const bind = () => {
            toggleRevenueBtn === null || toggleRevenueBtn === void 0 ? void 0 : toggleRevenueBtn.addEventListener("click", () => setViewMode('revenue'));
            toggleCollectionBtn === null || toggleCollectionBtn === void 0 ? void 0 : toggleCollectionBtn.addEventListener("click", () => setViewMode('collection'));
            showFormBtn === null || showFormBtn === void 0 ? void 0 : showFormBtn.addEventListener("click", () => {
                resetForm();
                ensureMonthGrid();
                showForm(false, currentViewMode);
            });
            backBtn === null || backBtn === void 0 ? void 0 : backBtn.addEventListener("click", () => {
                hideForm();
                renderSprintTable();
            });
            form === null || form === void 0 ? void 0 : form.addEventListener("submit", (e) => void saveRecord(e));
            projectSelect === null || projectSelect === void 0 ? void 0 : projectSelect.addEventListener("change", (e) => {
                const value = parseInt(e.target.value || '0', 10);
                if (!value) {
                    resetForm();
                    return;
                }
                const project = projects.find(p => p.Id === value);
                if (!project || !projectDetails)
                    return;
                const existingItem = currentRecords().find(r => { var _a; return ((_a = r.Project) === null || _a === void 0 ? void 0 : _a.Id) === value; });
                if (!existingItem) {
                    editingId = null;
                    editingType = currentViewMode;
                    showForm(false, currentViewMode);
                    projectDetails.innerHTML = detailRows(project);
                    ensureMonthGrid();
                    return;
                }
                if (editingId === existingItem.Id)
                    return;
                void Swal.fire({
                    icon: "question",
                    title: "Existing Record Found",
                    text: `A ${currentViewMode} record already exists for ${project.Title}. Would you like to edit it?`,
                    showCancelButton: true,
                    confirmButtonText: "Yes, Edit",
                    cancelButtonText: "No, Create New"
                }).then((result) => {
                    if (result.isConfirmed) {
                        void editItem(existingItem.Id, currentViewMode);
                    }
                    else {
                        editingId = null;
                        editingType = currentViewMode;
                        showForm(false, currentViewMode);
                        projectDetails.innerHTML = detailRows(project);
                        ensureMonthGrid();
                    }
                });
            });
            [filterYear, filterMonth, filterProject, filterCustomer, filterProjectManager,
                filterGroupHead, filterProjectReference, filterEnggLead]
                .forEach(sel => sel === null || sel === void 0 ? void 0 : sel.addEventListener("change", filterTable));
            clearFiltersBtn === null || clearFiltersBtn === void 0 ? void 0 : clearFiltersBtn.addEventListener("click", clearFilters);
            loadMoreBtn === null || loadMoreBtn === void 0 ? void 0 : loadMoreBtn.addEventListener("click", () => {
                toast("info", "Loading more data…");
                void refreshData();
            });
            // Keep the month grid and the table in step with the filters.
            monthGridRows.forEach(row => row.addEventListener("input", () => {
                const hasValue = Array.from(row.querySelectorAll('.sprint-num-input'))
                    .some(el => getNumericValue(el) !== 0);
                row.classList.toggle('is-dirty', hasValue);
            }));
        };
        /* ---------- INIT ---------- */
        const initialize = async () => {
            try {
                syncMonthSelection();
                paintViewMode();
                bind();
                if (remount && dataLoaded) {
                    // Re-entry: markup is fresh but the data cache is warm - just re-render.
                    syncMonthSelection();
                    renderSprintTable();
                    console.log("✅ Daily Sprint page re-mounted (cached data)");
                    return;
                }
                await authManager.initialize({ siteUrl });
                console.log(`👤 ${getCurrentUserFullName()} | ${getPermissionLevel().toUpperCase()} | [${getUserGroups().join(", ")}]`);
                await loadProjects();
                // BOTH lists, so switching views is instant and Collection isn't blank on first click.
                await Promise.all([loadRevenueForecasts(), loadCollectionItems()]);
                dataLoaded = true;
                initializeFilters();
                syncMonthSelection();
                renderSprintTable();
                updateUIBasedOnPermissions({
                    addBtn: '.daily-sprint-show-form-btn',
                    saveBtn: '.daily-sprint-btn-primary[type="submit"]'
                });
                if (pageHeader)
                    pageHeader.setAttribute('data-ready', 'true');
                console.log("✅ Daily Sprint page initialized");
            }
            catch (error) {
                console.error("❌ Failed to initialize sprint page:", error);
                toast("error", "Failed to initialize page");
            }
        };
        void initialize();
    }
};
// THIS is the line your web part was missing → TS1192 "has no default export".
export default sprintPage;
export { sprintPage };
//# sourceMappingURL=sprint.js.map
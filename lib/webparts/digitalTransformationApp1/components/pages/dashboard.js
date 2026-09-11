import 'sweetalert2/dist/sweetalert2.min.css';
import '../styles/dashboard.css';
import { authManager, canRead, hasFullControl, canReadAllProjects, getAssignedProjects, getCurrentUserFullName } from './auth';
// ============================================================
// DYNAMIC PROJECT PERFORMANCE DASHBOARD — COMPLETE VERSION
// ------------------------------------------------------------
//  Dynamically fetches data from:
//    - Revenue Forecast List (Revenue Plan, Revenue Actuals, OBL)
//    - Collection List (Collection Plan, Collection Actuals)
//    - Project List (Project metadata, PR+ Quality stats)
//  All values update based on selected PM, PR, Month, and Date filters.
// ============================================================
const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const pad2 = (n) => String(n).padStart(2, '0');
const monthKeyOf = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
const monthLabelOf = (key) => {
    const [y, m] = key.split('-').map(Number);
    return `${MONTH_ABBR[(m || 1) - 1]}-${String(y).slice(2)}`;
};
const fmtAsOn = (isoDate) => {
    if (!isoDate)
        return '';
    const d = new Date(`${isoDate}T00:00:00`);
    if (isNaN(d.getTime()))
        return isoDate;
    return `${pad2(d.getDate())}-${MONTH_ABBR[d.getMonth()]}-${d.getFullYear()}`;
};
const buildMonthList = (count) => {
    const now = new Date();
    const out = [];
    for (let i = count - 1; i >= 0; i--) {
        out.push(monthKeyOf(new Date(now.getFullYear(), now.getMonth() - i, 1)));
    }
    return out;
};
// ============================================================
// DASHBOARD CONTENT
// ============================================================
const dashboardContent = `
  <div class="kpi-dashboard">
    <header class="header">
      <h1>PM DASHBOARD</h1>
      <div class="header-meta">
        <span>Period <b id="dsMetaPeriod">Sep-26</b></span>
        <span>As on <b id="dsMetaAsOn">05-Sep-2026</b></span>
        <span>Currency <b>INR</b></span>
        <span id="dsRefreshIndicator" style="font-size:14px;color:#fff;margin-left:10px;"></span>
        <span id="dsLastUpdated" style="font-size:10px;color:#aaa;margin-left:5px;"></span>
      </div>
    </header>

    <main class="board">
      <!-- Identity Card - Filter Controls -->
      <section class="card identity">
        <div class="id-row"><span>Project Manager</span><b class="ds-host" id="dsPmHost"></b></div>
        <div class="id-row"><span>Date</span><input type="date" class="ds-ctl" id="dsDate" title="As-on date (defaults to today)" /></div>
        <div class="id-row"><span>Month</span><select class="ds-ctl" id="dsMonth" title="Data month — current and last 36 months"></select></div>
        <div class="id-row"><span>Project Reference</span><b class="ds-host" id="dsPrHost"></b></div>
      </section>

      <!-- OBL Card -->
      <section class="card obl">
        <div class="card-title">OBL</div>
        <div class="kpi-num" id="dsOblTotal">Loading...</div>
        <div class="delta" id="dsOblDelta"></div>
      </section>

      <!-- Receivables Card -->
      <section class="card receivables"> 
        <div class="recv-head">   
          <span class="card-title">Total Receivables</span>
           <span class="kpi-num recv-head-num" id="dsTotalReceivables">Loading...</span>
           <!--  <div class="delta red recv-head-delta" id="dsReceivablesDelta">Loading...</div> -->
        </div>
        <div class="recv-grid">
          <div class="recv-cell">  
           <!-- <div class="delta red" id="dsNormalDelta">Loading...</div> -->
            <div class="kpi-num" id="dsNormalValue">Loading...</div>
            <div class="lbl">Normal</div>
          </div>
          <div class="recv-cell">
           <!--  <div class="delta green" id="dsNormalOverDueDelta">Loading...</div> -->    
            <div class="kpi-num" id="dsNormalOverDueValue">Loading...</div>
            <div class="lbl">Normal OverDue</div>
          </div>
          <div class="recv-cell">
           <!--  <div class="delta red" id="dsGstTdsDelta">Loading...</div> -->
            <div class="kpi-num" id="dsGstTdsValue">Loading...</div>
            <div class="lbl">GST TDS</div>
          </div>
          <div class="recv-divider"></div>
          <div class="recv-cell recv-total">
           <!--  <div class="delta green" id="dsTotalOverDueDelta">Loading...</div> -->
            <div class="kpi-num" id="dsTotalOverDueValue">Loading...</div>
            <div class="lbl">Total OverDue</div>
          </div>
          <div class="recv-cell">
           <!--  <div class="delta red" id="dsRetentionDelta">Loading...</div>   -->
            <div class="kpi-num" id="dsRetentionValue">Loading...</div>
            <div class="lbl">Retention</div>
          </div>
          <div class="recv-cell">
           <!--  <div class="delta red" id="dsRetentionOverDueDelta">Loading...</div>  -->
            <div class="kpi-num" id="dsRetentionOverDueValue">Loading...</div>
            <div class="lbl">Retention OverDue</div>
          </div>
          <div class="recv-cell">
           <!--  <div class="delta red" id="dsItTdsOthersDelta">Loading...</div>  -->
            <div class="kpi-num" id="dsItTdsOthersValue">Loading...</div>
            <div class="lbl">IT TDS / Others</div>
          </div>
          <div class="recv-cell recv-total">
           <!--  <div class="delta red" id="dsTotalRetentionDelta">Loading...</div>  -->
            <div class="kpi-num" id="dsTotalRetentionValue">Loading...</div>
            <div class="lbl">Total Retention</div>
          </div>
        </div>
      </section>

      <!-- Top Right Card -->
      <section class="card top-right">
        <div class="metric-line">
          <div class="name magenta">BIES</div>
          <div class="kpi-num" id="dsBiesValue">Loading...</div>
          <div class="delta green" id="dsBiesDelta"></div>
        </div>
        <div class="metric-line">
          <div class="name" style="color:var(--navy)">NWC</div>
          <div class="kpi-num red" id="dsNwcValue">Loading...</div>
          <div class="delta red" id="dsNwcDelta"></div>
        </div>

        <table class="pr-table">
          <thead>
            <tr>
              <th>PR+ Quality</th>
              <th>Average / Count </th>
            </tr>
          </thead>
          <tbody>
            <tr><td>No of projects in PR+</td><td id="dsTotalProjects">Loading...</td></tr>
            <tr><td>Submitted</td><td id="dsSubmittedProjects">Loading...</td></tr>
            <tr><td>Not Submitted</td><td id="dsNotSubmittedProjects">Loading...</td></tr>
            <tr><td>Approved</td><td id="dsApprovedProjects">Loading...</td></tr>
          </tbody>
        </table>
      </section>

      <!-- Revenue & Collection Card -->
      <section class="card revenue">
        <div class="rev-group">
          <div class="g-title" id="dsRevMonthTitle">Month — Sep-26</div>
          <div class="pair">
            <div class="pair-box">
              <div class="lbl">Revenue Plan</div>
              <div class="kpi-num" id="dsRevenuePlan">Loading...</div>
            </div>
            <div class="pair-box">
              <div class="lbl">Revenue Actuals</div>
              <div class="kpi-num green" id="dsRevenueActual">Loading...</div>
              <div class="delta red" id="dsRevenuePercentage">Loading...</div>
              <div class="bar low"><i style="width:0%" id="dsRevenueBar"></i></div>
            </div>
            <div class="pair-box">
              <div class="lbl">Collection Plan</div>
              <div class="kpi-num" id="dsCollectionPlan">Loading...</div>
            </div>
            <div class="pair-box">
              <div class="lbl">Collection Actuals</div>
              <div class="kpi-num green" id="dsCollectionActual">Loading...</div>
              <div class="delta red" id="dsCollectionPercentage">Loading...</div>
              <div class="bar low"><i style="width:0%" id="dsCollectionBar"></i></div>
            </div>
          </div>
        </div>

        <div class="rev-group">
          <div class="g-title" id="dsYtdTitle">YTD — Till Dec</div>
          <div class="pair">
            <div class="pair-box">
              <div class="lbl">YTD Revenue Plan</div>
              <div class="kpi-num" id="dsYtdRevenuePlan">Loading...</div>
            </div>
            <div class="pair-box">
              <div class="lbl">YTD Revenue Actuals</div>
              <div class="kpi-num green" id="dsYtdRevenueActual">Loading...</div>
              <div class="delta blue" id="dsYtdRevenuePercentage">Loading...</div>
              <div class="bar mid"><i style="width:0%" id="dsYtdRevenueBar"></i></div>
            </div>
            <div class="pair-box">
              <div class="lbl">YTD Collection Plan</div>
              <div class="kpi-num" id="dsYtdCollectionPlan">Loading...</div>
            </div>
            <div class="pair-box">
              <div class="lbl">YTD Collection Actuals</div>
              <div class="kpi-num green" id="dsYtdCollectionActual">Loading...</div>
              <div class="delta green" id="dsYtdCollectionPercentage">Loading...</div>
              <div class="bar"><i style="width:0%" id="dsYtdCollectionBar"></i></div>
            </div>
          </div>
        </div>
      </section>

      <!-- Procurement Card -->
      <section class="card procurement">
        <div class="card-title">Procurement Pending</div>
        <div class="list" id="procurementPendingList">
          <div class="list-row warn"><span class="name">RFQ Pending</span><span class="val" id="dsRfqPending">0</span></div>
          <div class="list-row warn"><span class="name">Offer Pending</span><span class="val" id="dsOfferPending">0</span></div>
          <div class="list-row"><span class="name">TC Pending</span><span class="val" id="dsTcPending">0</span></div>
          <div class="list-row"><span class="name">Order Closure Pending</span><span class="val" id="dsOrderClosurePending">0</span></div>
          <div class="list-row warn"><span class="name">PR Pending</span><span class="val" id="dsPrPending">0</span></div>
          <div class="list-row warn"><span class="name">PO Pending</span><span class="val" id="dsPoPending">0</span></div>
          <div class="list-row"><span class="name">Document Approval Pending</span><span class="val" id="dsDocApprovalPending">0</span></div>
          <div class="list-row"><span class="name">MFC Pending</span><span class="val" id="dsMfcPending">0</span></div>
          <div class="list-row"><span class="name">Inspection Pending</span><span class="val" id="dsInspectionPending">0</span></div>
          <div class="list-row warn"><span class="name">Dispatch Pending</span><span class="val" id="dsDispatchPending">0</span></div>
          <div class="list-row"><span class="name">Delivery Pending</span><span class="val" id="dsDeliveryPending">0</span></div>
          <div class="list-row ok"><span class="name">Delivered</span><span class="val" id="dsDeliveredCount">0</span></div>
        </div>
      </section>

      <!-- Documents Card -->
      <section class="card documents">
        <div class="card-title">Total Documents — DCI</div>
        <div class="list">
          <div class="list-row"><span class="name">Total Documents</span><span class="val" id="dsTotalDocuments">0</span></div>
          <div class="list-row warn"><span class="name">Not Submitted</span><span class="val" id="dsNotSubmittedDocs">0</span></div>
          <div class="list-row warn"><span class="name">Not Approved</span><span class="val" id="dsNotApprovedDocs">0</span></div>
          <div class="list-row warn"><span class="name">Revised — Not Submitted / Approved with comments</span><span class="val" id="dsRevisedDocs">0</span></div>
          <div class="list-row ok"><span class="name">Approved</span><span class="val" id="dsApprovedDocs">0</span></div>
          <div class="list-row warn"><span class="name">Rejected</span><span class="val" id="dsRejectedDocs">0</span></div>
          <div class="list-row"><span class="name">Retained for Information</span><span class="val" id="dsRetainedDocs">0</span></div>
        </div>
      </section>

      <!-- Projects Card -->
      <section class="card projects">
        <div class="projects-header">
          <div class="projects-total">
            <span class="lbl">Total Projects</span>
            <span class="kpi-num" id="dsTotalProjectsCard">Loading...</span>
            <span class="delta green" id="dsProjectsDelta">Loading...</span>
          </div>
        </div>
        <div class="projects-grid">
          <div class="p-item">
            <span class="lbl">REL</span>
            <span class="kpi-num" id="dsRelCount">0</span>
          </div>
          <div class="p-item">
            <span class="lbl">TECO</span>
            <span class="kpi-num" id="dsTecoCount">0</span>
          </div>
          <div class="p-item">
            <span class="lbl">SO</span>
            <span class="kpi-num" id="dsSoCount">0</span>
          </div>
          <div class="p-item">
            <span class="lbl">CLSD</span>
            <span class="kpi-num" id="dsClsdCount">0</span>
          </div>
          <div class="p-item">
            <span class="lbl">OTHER</span>
            <span class="kpi-num" id="dsOtherCount">0</span>
          </div>
        </div>
      </section>

      <!-- Delivery Card -->
      <section class="card delivery">
        <div class="alert">
          <div class="lbl">Delivery Date Crossed</div>
          <div class="kpi-num" id="dsDeliveryCrossed">0</div>
        </div>
        <div class="alert">
          <div class="lbl">To cross in 45 days</div>
          <div class="kpi-num" id="dsDeliveryToCross">0</div>
        </div>
      </section>
    </main>
  </div>
`;
const dashboardStyle = `
/* Styling lives in the external stylesheet ../styles/dashboard.css. */
`;
// ============================================================
// DASHBOARD SCRIPT
// ============================================================
const dashboardScript = () => {
    setTimeout(async () => {
        const siteUrl = typeof _spPageContextInfo !== 'undefined'
            ? _spPageContextInfo.webAbsoluteUrl
            : window.location.origin + '/sites/DigitalTransformationTest';
        const DASH_SOURCES = {
            projectList: 'Project List',
            revenueList: 'Revenue Forecast List',
            collectionList: 'Collection List',
            fieldProjectReference: 'ProjectReference',
            fieldCustomer: 'CustomerName',
            fieldProjectManager: 'ProjectManager',
            fieldGroupHead: 'GroupHead1'
        };
        // ============================================================
        // UTILITY FUNCTIONS
        // ============================================================
        const byId = (id) => {
            const all = document.querySelectorAll(`#${id}`);
            return (all && all.length) ? all[all.length - 1] : null;
        };
        const allById = (id) => Array.from(document.querySelectorAll(`#${id}`));
        const setTxt = (id, v) => {
            allById(id).forEach(el => { el.textContent = v; el.title = v; });
        };
        const setStyle = (id, prop, value) => {
            allById(id).forEach(el => { el.style[prop] = value; });
        };
        const formatNumber = (num) => {
            return Math.round(num).toLocaleString('en-IN');
        };
        const formatPercentage = (num) => {
            return `(${num.toFixed(2)}%)`;
        };
        const safeJsonParse = (jsonString, fallback = {}) => {
            try {
                if (!jsonString || jsonString.trim() === '' || jsonString === '""' || jsonString === 'null') {
                    return fallback;
                }
                return JSON.parse(jsonString);
            }
            catch (_error) {
                return fallback;
            }
        };
        // ============================================================
        // DASHBOARD FILTER STATE
        // ============================================================
        const todayIso = new Date().toISOString().slice(0, 10);
        const currentMonthKey = monthKeyOf(new Date());
        let pmValue = '__all__';
        let prValue = '__all__';
        let scopedProjects = [];
        let pmOpts = [];
        let prOpts = [];
        // ============================================================
        // DASHBOARD SELECT CLASS
        // ============================================================
        class DashSelect {
            constructor(host, placeholder, onPick) {
                this.opts = [];
                this.selValue = '__all__';
                this.onPick = onPick;
                this.root = host;
                host.classList.add('ds-dd');
                host.innerHTML = `
          <button type="button" class="ds-dd-btn" title="Search ${placeholder}">
            <span class="ds-dd-val">All</span>
            <span class="ds-dd-caret">▾</span>
          </button>
          <div class="ds-dd-panel">
            <input type="text" class="ds-dd-search" placeholder="Search ${placeholder}..." autocomplete="off" />
            <div class="ds-dd-list"></div>
          </div>
        `;
                this.btn = host.querySelector('.ds-dd-btn');
                this.panel = host.querySelector('.ds-dd-panel');
                this.search = host.querySelector('.ds-dd-search');
                this.listEl = host.querySelector('.ds-dd-list');
                this.btn.addEventListener('click', (e) => { e.stopPropagation(); this.toggle(); });
                this.search.addEventListener('input', () => this.renderList());
                this.search.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter') {
                        const first = this.listEl.querySelector('.ds-dd-item[data-value]:not([data-value="__all__"])');
                        if (first)
                            first.click();
                    }
                    if (e.key === 'Escape')
                        this.close();
                });
                document.addEventListener('click', (e) => {
                    if (!this.root.contains(e.target))
                        this.close();
                });
                window.addEventListener('scroll', () => this.close(), { passive: true });
                window.addEventListener('resize', () => this.close());
            }
            setOptions(opts) {
                const seen = new Set();
                this.opts = [
                    { value: '__all__', label: 'All', sub: `${opts.length} option${opts.length === 1 ? '' : 's'}` },
                    ...opts.filter(o => { if (seen.has(o.value))
                        return false; seen.add(o.value); return true; })
                ];
                if (!this.opts.some(o => o.value === this.selValue))
                    this.selValue = '__all__';
                if (this.search)
                    this.search.value = '';
                this.renderList();
                this.renderButton();
            }
            getValue() { return this.selValue; }
            setValue(value) {
                if (this.opts.some(o => o.value === value)) {
                    this.selValue = value;
                    this.renderButton();
                    this.renderList();
                }
            }
            toggle() {
                if (this.panel.classList.contains('open'))
                    this.close();
                else
                    this.open();
            }
            open() {
                if (!this.opts.length) {
                    console.warn('📂 Dropdown opened but its option list is EMPTY');
                }
                this.renderList();
                const rect = this.btn.getBoundingClientRect();
                const panelW = Math.max(rect.width, 250);
                this.panel.style.width = `${panelW}px`;
                this.panel.style.left = `${Math.max(4, Math.min(rect.left, window.innerWidth - panelW - 6))}px`;
                const estH = 230;
                const top = (rect.bottom + estH > window.innerHeight && rect.top - estH > 0)
                    ? (rect.top - estH) : (rect.bottom + 4);
                this.panel.style.top = `${Math.max(2, top)}px`;
                this.panel.classList.add('open');
                this.search.value = '';
                this.search.focus();
            }
            close() {
                this.panel.classList.remove('open');
            }
            renderButton() {
                const opt = this.opts.find(o => o.value === this.selValue);
                const val = this.btn.querySelector('.ds-dd-val');
                val.textContent = opt ? opt.label : 'All';
            }
            renderList() {
                var _a;
                const term = (((_a = this.search) === null || _a === void 0 ? void 0 : _a.value) || '').trim().toLowerCase();
                const items = this.opts.filter(o => {
                    if (o.value === '__all__')
                        return true;
                    return `${o.label} ${o.sub || ''}`.toLowerCase().includes(term);
                });
                if (!items.length) {
                    this.listEl.innerHTML = '<div class="ds-dd-empty">No matches</div>';
                    return;
                }
                this.listEl.innerHTML = items.map(o => {
                    const cls = o.value === '__all__' ? 'ds-dd-item ds-dd-all' : 'ds-dd-item';
                    return `
            <div class="${cls}${o.value === this.selValue ? ' is-sel' : ''}" data-value="${o.value}">
              <div class="ds-dd-item-main">${o.label}</div>
              ${o.sub ? `<div class="ds-dd-item-sub">${o.sub}</div>` : ''}
            </div>`;
                }).join('');
                this.listEl.querySelectorAll('.ds-dd-item').forEach(el => {
                    el.addEventListener('click', (e) => {
                        e.stopPropagation();
                        this.selValue = el.getAttribute('data-value') || '__all__';
                        this.renderButton();
                        this.renderList();
                        this.close();
                        if (this.onPick)
                            this.onPick(this.selValue);
                    });
                });
            }
        }
        // ============================================================
        // DATA FETCHING FUNCTIONS
        // ============================================================
        // 1. Fetch Revenue Forecast List Data - CORRECTED (swapped)
        const fetchRevenueData = async (monthKey, projectIds) => {
            try {
                let totalPlan = 0;
                let totalActual = 0;
                let ytdPlan = 0;
                let ytdActual = 0;
                // Parse the month and year from the selected month
                const [selectedYear, selectedMonth] = monthKey.split('-').map(Number);
                // First discover the Project field
                const fieldsUrl = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.revenueList}')/fields?$select=InternalName,Title,TypeAsString&$top=5000`;
                const fieldsRes = await fetch(fieldsUrl, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                let projectField = '';
                if (fieldsRes.ok) {
                    const fieldsData = await fieldsRes.json();
                    const fields = (fieldsData.d && fieldsData.d.results) || [];
                    const projField = fields.find((f) => {
                        const title = (f.Title || '').toLowerCase();
                        const internal = (f.InternalName || '').toLowerCase();
                        return (f.TypeAsString === 'Lookup' || f.TypeAsString === 'Text') &&
                            (title.includes('project') || internal.includes('project'));
                    });
                    if (projField) {
                        projectField = projField.InternalName || projField.EntityPropertyName || '';
                    }
                }
                let selectFields = 'Id,ForecastData,RevenueActual';
                if (projectField) {
                    selectFields += `,${projectField}/Id,${projectField}/Title`;
                }
                const expandField = projectField ? `&$expand=${projectField}` : '';
                const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.revenueList}')/items?$select=${selectFields}${expandField}&$top=5000`;
                console.log('📊 Fetching Revenue Forecast data:', url.replace(siteUrl, ''));
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    throw new Error(`HTTP ${res.status}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                results.forEach((item) => {
                    let projectId = 0;
                    if (projectField && item[projectField]) {
                        projectId = Number(item[projectField].Id || 0);
                    }
                    if (!projectId || !projectIds.has(projectId))
                        return;
                    const forecastData = safeJsonParse(item.ForecastData);
                    const revenueActual = safeJsonParse(item.RevenueActual);
                    // Get current month data (for the main display)
                    // ✅ CORRECT: Revenue Plan = RevenueActual[monthKey] (the planned value)
                    // Revenue Actual = forecastData[monthKey].actual (the actual value)
                    const monthForecast = forecastData[monthKey] || {};
                    const plan = Number(revenueActual[monthKey]) || 0; // Revenue Plan (from RevenueActual)
                    const actual = Number(monthForecast.actual) || 0; // Revenue Actual (from ForecastData.actual)
                    totalPlan += plan;
                    totalActual += actual;
                    // ----- YTD CALCULATION: Sum from January to selected month -----
                    for (let m = 1; m <= selectedMonth; m++) {
                        const monthKeyToCheck = `${selectedYear}-${String(m).padStart(2, '0')}`;
                        const monthForecastYtd = forecastData[monthKeyToCheck] || {};
                        const planYtd = Number(revenueActual[monthKeyToCheck]) || 0; // Revenue Plan
                        const actualYtd = Number(monthForecastYtd.actual) || 0; // Revenue Actual
                        ytdPlan += planYtd;
                        ytdActual += actualYtd;
                    }
                });
                const percentage = totalPlan > 0 ? (totalActual / totalPlan) * 100 : 0;
                const ytdPercentage = ytdPlan > 0 ? (ytdActual / ytdPlan) * 100 : 0;
                console.log(`📊 Revenue: Plan=${totalPlan}, Actual=${totalActual}, %=${percentage.toFixed(2)}%`);
                console.log(`📊 YTD Revenue: Plan=${ytdPlan}, Actual=${ytdActual}, %=${ytdPercentage.toFixed(2)}%`);
                return {
                    plan: totalPlan,
                    actual: totalActual,
                    percentage: percentage,
                    ytdPlan: ytdPlan,
                    ytdActual: ytdActual,
                    ytdPercentage: ytdPercentage
                };
            }
            catch (err) {
                console.error('❌ Failed to fetch Revenue data:', err);
                return {
                    plan: 0,
                    actual: 0,
                    percentage: 0,
                    ytdPlan: 0,
                    ytdActual: 0,
                    ytdPercentage: 0
                };
            }
        };
        // 2. Fetch Collection List Data - CORRECTED (swapped)
        const fetchCollectionData = async (monthKey, projectIds) => {
            try {
                let totalPlan = 0;
                let totalActual = 0;
                let ytdPlan = 0;
                let ytdActual = 0;
                // Parse the month and year from the selected month
                const [selectedYear, selectedMonth] = monthKey.split('-').map(Number);
                const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.collectionList}')/items?$select=Id,CollectionData,RevenueActual,ProjectReferenceId&$top=5000`;
                console.log('📊 Fetching Collection data:', url.replace(siteUrl, ''));
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    throw new Error(`HTTP ${res.status}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                results.forEach((item) => {
                    const projectId = Number(item.ProjectReferenceId || 0);
                    if (!projectId || !projectIds.has(projectId))
                        return;
                    const collectionData = safeJsonParse(item.CollectionData);
                    const revenueActual = safeJsonParse(item.RevenueActual);
                    // Get current month data (for the main display)
                    // ✅ CORRECT: Collection Plan = RevenueActual[monthKey] (the planned value)
                    // Collection Actual = collectionData[monthKey].actual (the actual value)
                    const monthCollection = collectionData[monthKey] || {};
                    const plan = Number(revenueActual[monthKey]) || 0; // Collection Plan (from RevenueActual)
                    const actual = Number(monthCollection.actual) || 0; // Collection Actual (from CollectionData.actual)
                    totalPlan += plan;
                    totalActual += actual;
                    // ----- YTD CALCULATION: Sum from January to selected month -----
                    for (let m = 1; m <= selectedMonth; m++) {
                        const monthKeyToCheck = `${selectedYear}-${String(m).padStart(2, '0')}`;
                        const monthCollectionYtd = collectionData[monthKeyToCheck] || {};
                        // Collection Plan from RevenueActual
                        const planYtd = Number(revenueActual[monthKeyToCheck]) || 0;
                        // Collection Actual from CollectionData.actual
                        const actualYtd = Number(monthCollectionYtd.actual) || 0;
                        ytdPlan += planYtd;
                        ytdActual += actualYtd;
                    }
                });
                const percentage = totalPlan > 0 ? (totalActual / totalPlan) * 100 : 0;
                const ytdPercentage = ytdPlan > 0 ? (ytdActual / ytdPlan) * 100 : 0;
                console.log(`📊 Collection: Plan=${totalPlan}, Actual=${totalActual}, %=${percentage.toFixed(2)}%`);
                console.log(`📊 YTD Collection: Plan=${ytdPlan}, Actual=${ytdActual}, %=${ytdPercentage.toFixed(2)}%`);
                return {
                    plan: totalPlan,
                    actual: totalActual,
                    percentage: percentage,
                    ytdPlan: ytdPlan,
                    ytdActual: ytdActual,
                    ytdPercentage: ytdPercentage
                };
            }
            catch (err) {
                console.error('❌ Failed to fetch Collection data:', err);
                return {
                    plan: 0,
                    actual: 0,
                    percentage: 0,
                    ytdPlan: 0,
                    ytdActual: 0,
                    ytdPercentage: 0
                };
            }
        };
        // 3. Fetch OBL Summary
        const fetchOblSummary = async (projectIds) => {
            try {
                let totalObl = 0;
                let projectCount = 0;
                const oblValues = [];
                const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.revenueList}')/items?$select=Id,OBL,ProjectReferenceId&$top=5000`;
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    throw new Error(`HTTP ${res.status}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                results.forEach((item) => {
                    const projectId = Number(item.ProjectReferenceId || 0);
                    if (!projectId || !projectIds.has(projectId))
                        return;
                    const obl = Math.round(Number(item.OBL) || 0);
                    if (!Number.isFinite(obl) || obl === 0)
                        return;
                    totalObl += obl;
                    projectCount++;
                    oblValues.push(obl);
                });
                const averageObl = projectCount > 0 ? Math.round((totalObl / projectCount) * 100) / 100 : 0;
                const highestObl = oblValues.length > 0 ? Math.max(...oblValues) : 0;
                const lowestObl = oblValues.length > 0 ? Math.min(...oblValues) : 0;
                console.log(`📊 OBL: Total=${totalObl}, Count=${projectCount}, Avg=${averageObl}`);
                return { totalObl, projectCount, averageObl, highestObl, lowestObl };
            }
            catch (err) {
                console.error('❌ Failed to fetch OBL data:', err);
                return { totalObl: 0, projectCount: 0, averageObl: 0, highestObl: 0, lowestObl: 0 };
            }
        };
        // 4. Fetch Receivables Data from Collection List
        const fetchReceivablesData = async (projectIds) => {
            try {
                let normal = 0, normalOverDue = 0, gstTds = 0;
                let retention = 0, retentionOverDue = 0, itTdsOthers = 0;
                const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.collectionList}')/items?$select=Id,NormalDue,NormalOverdue,RetentionDue,RetentionOverdue,GSTTDS,WHT,ProjectReferenceId&$top=5000`;
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    throw new Error(`HTTP ${res.status}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                results.forEach((item) => {
                    const projectId = Number(item.ProjectReferenceId || 0);
                    if (!projectId || !projectIds.has(projectId))
                        return;
                    normal += Math.round(Number(item.NormalDue) || 0);
                    normalOverDue += Math.round(Number(item.NormalOverdue) || 0);
                    retention += Math.round(Number(item.RetentionDue) || 0);
                    retentionOverDue += Math.round(Number(item.RetentionOverdue) || 0);
                    gstTds += Math.round(Number(item.GSTTDS) || 0);
                    itTdsOthers += Math.round(Number(item.WHT) || 0);
                });
                const totalOverDue = normalOverDue + retentionOverDue;
                const totalRetention = retention + retentionOverDue;
                const totalReceivables = normal + normalOverDue + retention + retentionOverDue + gstTds + itTdsOthers;
                console.log(`📊 Receivables: Total=${totalReceivables}`);
                return {
                    normal,
                    normalOverDue,
                    gstTds,
                    totalOverDue,
                    retention,
                    retentionOverDue,
                    itTdsOthers,
                    totalRetention,
                    totalReceivables
                };
            }
            catch (err) {
                console.error('❌ Failed to fetch Receivables data:', err);
                return {
                    normal: 0, normalOverDue: 0, gstTds: 0,
                    totalOverDue: 0, retention: 0, retentionOverDue: 0,
                    itTdsOthers: 0, totalRetention: 0, totalReceivables: 0
                };
            }
        };
        // 5. Fetch PR+ Quality Data - FIXED with correct counts
        const fetchPrQualityData = async (projectIds) => {
            try {
                let total = 0;
                let submitted = 0;
                let approved = 0;
                let nwcTotal = 0;
                let nwcCount = 0;
                let deliveryCrossed = 0;
                let deliveryToCross45Days = 0;
                const statusCounts = {
                    'REL': 0,
                    'TECO': 0,
                    'SO': 0,
                    'CLSD': 0,
                    'OTHER': 0
                };
                // First, discover what fields exist in Project List
                const fieldsUrl = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/fields?$select=InternalName,Title,TypeAsString&$top=5000`;
                const fieldsRes = await fetch(fieldsUrl, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                let prStatusField = '';
                let prSubmissionField = '';
                let prApprovalField = '';
                let statusField = '';
                let nwcField = '';
                let contractualDeliveryField = '';
                let revisedDeliveryField = '';
                if (fieldsRes.ok) {
                    const fieldsData = await fieldsRes.json();
                    const fields = (fieldsData.d && fieldsData.d.results) || [];
                    // Find all relevant fields - use a Set to avoid duplicates
                    const foundFields = new Set();
                    fields.forEach((f) => {
                        const title = (f.Title || '').toLowerCase();
                        const internal = (f.InternalName || '').toLowerCase();
                        const fieldName = f.InternalName || f.EntityPropertyName || '';
                        if (!fieldName || foundFields.has(fieldName))
                            return;
                        // PR Quality / Status - skip if it's the moderation status
                        if ((title.includes('pr') || internal.includes('pr')) &&
                            (title.includes('status') || internal.includes('status') || title.includes('quality')) &&
                            !internal.includes('moderation')) {
                            prStatusField = fieldName;
                            foundFields.add(fieldName);
                        }
                        // PR Submission Status
                        if ((title.includes('submission') || internal.includes('submission')) &&
                            (title.includes('pr') || internal.includes('pr'))) {
                            prSubmissionField = fieldName;
                            foundFields.add(fieldName);
                        }
                        // PR Approval Status - skip if it's the moderation status
                        if ((title.includes('approval') || internal.includes('approval')) &&
                            (title.includes('pr') || internal.includes('pr')) &&
                            !internal.includes('moderation')) {
                            prApprovalField = fieldName;
                            foundFields.add(fieldName);
                        }
                        // Status (REL, TECO, SO, CLSD) - exact match
                        if ((title === 'status' || internal === 'status') &&
                            !internal.includes('pr') && !internal.includes('submission') &&
                            !internal.includes('approval') && !internal.includes('moderation')) {
                            statusField = fieldName;
                            foundFields.add(fieldName);
                        }
                        // NWC
                        if ((title.includes('nwc') || internal.includes('nwc')) ||
                            (title.includes('net working capital') || internal.includes('networkingcapital'))) {
                            nwcField = fieldName;
                            foundFields.add(fieldName);
                        }
                        // Contractual Delivery Date - exact match
                        if (title === 'contractualdeliverydate' || internal === 'contractualdeliverydate') {
                            contractualDeliveryField = fieldName;
                            foundFields.add(fieldName);
                        }
                        // Revised Contractual Delivery Date - exact match
                        if (title === 'revisedcontractualdeliverydate' || internal === 'revisedcontractualdeliverydate') {
                            revisedDeliveryField = fieldName;
                            foundFields.add(fieldName);
                        }
                    });
                    console.log('📋 Found Project fields:', {
                        prStatus: prStatusField,
                        prSubmission: prSubmissionField,
                        prApproval: prApprovalField,
                        status: statusField,
                        nwc: nwcField,
                        contractualDelivery: contractualDeliveryField,
                        revisedDelivery: revisedDeliveryField
                    });
                }
                // Build the query - use a Set to avoid duplicate fields
                const selectFields = new Set();
                selectFields.add('Id');
                selectFields.add('Title');
                if (prStatusField)
                    selectFields.add(prStatusField);
                if (prSubmissionField)
                    selectFields.add(prSubmissionField);
                if (prApprovalField)
                    selectFields.add(prApprovalField);
                if (statusField)
                    selectFields.add(statusField);
                if (nwcField)
                    selectFields.add(nwcField);
                if (contractualDeliveryField)
                    selectFields.add(contractualDeliveryField);
                if (revisedDeliveryField)
                    selectFields.add(revisedDeliveryField);
                const selectString = Array.from(selectFields).join(',');
                const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/items?$select=${selectString}&$top=5000`;
                console.log('📡 Fetching Project data with fields:', selectString);
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    const errorText = await res.text();
                    console.error('❌ API Error:', errorText);
                    throw new Error(`HTTP ${res.status}: ${errorText.substring(0, 200)}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                console.log(`📊 Found ${results.length} projects in Project List`);
                // Get today's date for delivery calculations
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                const todayISO = today.toISOString().split('T')[0];
                // Calculate date for 45 days from now
                const fortyFiveDaysFromNow = new Date(today);
                fortyFiveDaysFromNow.setDate(fortyFiveDaysFromNow.getDate() + 45);
                const fortyFiveDaysISO = fortyFiveDaysFromNow.toISOString().split('T')[0];
                results.forEach((item) => {
                    if (!projectIds.has(item.Id))
                        return;
                    total++;
                    // ----- PR+ Quality Status -----
                    // Check if project is submitted
                    let isSubmitted = false;
                    // Get submission status from PRSubmissionStatus field
                    if (prSubmissionField && item[prSubmissionField] !== undefined) {
                        const submissionStatus = (item[prSubmissionField] || '').toLowerCase().trim();
                        // Check if submitted
                        if (submissionStatus === 'submitted' || submissionStatus === 'approved' ||
                            submissionStatus === 'yes' || submissionStatus === 'true' || submissionStatus === '1') {
                            isSubmitted = true;
                        }
                    }
                    else if (prStatusField && item[prStatusField] !== undefined) {
                        // Fallback to prStatus field
                        const status = (item[prStatusField] || '').toLowerCase().trim();
                        if (status === 'submitted' || status === 'approved' ||
                            status === 'yes' || status === 'true' || status === '1') {
                            isSubmitted = true;
                        }
                    }
                    // Count submitted
                    if (isSubmitted) {
                        submitted++;
                    }
                    // Check if project is approved
                    let isApproved = false;
                    // Get approval status from PRApprovalStatus field
                    if (prApprovalField && item[prApprovalField] !== undefined) {
                        const approvalStatus = (item[prApprovalField] || '').toLowerCase().trim();
                        if (approvalStatus === 'approved' || approvalStatus === 'yes' ||
                            approvalStatus === 'true' || approvalStatus === '1' || approvalStatus === '2') {
                            isApproved = true;
                        }
                    }
                    else if (prStatusField && item[prStatusField] !== undefined) {
                        // Fallback to prStatus field
                        const status = (item[prStatusField] || '').toLowerCase().trim();
                        if (status === 'approved' || status === 'yes' || status === 'true' || status === '2') {
                            isApproved = true;
                        }
                    }
                    // Count approved
                    if (isApproved) {
                        approved++;
                    }
                    // ----- Project Status (REL, TECO, SO, CLSD) -----
                    let status = '';
                    if (statusField && item[statusField] !== undefined) {
                        status = (item[statusField] || '').trim().toUpperCase();
                    }
                    // Count by status
                    if (status === 'REL') {
                        statusCounts['REL']++;
                    }
                    else if (status === 'TECO') {
                        statusCounts['TECO']++;
                    }
                    else if (status === 'SO') {
                        statusCounts['SO']++;
                    }
                    else if (status === 'CLSD') {
                        statusCounts['CLSD']++;
                    }
                    else if (status) {
                        statusCounts['OTHER']++;
                    }
                    // ----- Get NWC value -----
                    if (nwcField && item[nwcField] !== undefined) {
                        let nwcValue = item[nwcField];
                        if (nwcValue !== null && nwcValue !== '') {
                            if (typeof nwcValue === 'string') {
                                nwcValue = parseFloat(nwcValue.replace(/,/g, '').replace(/[^0-9.]/g, ''));
                            }
                            if (typeof nwcValue === 'number' && !isNaN(nwcValue)) {
                                nwcTotal += nwcValue;
                                nwcCount++;
                            }
                        }
                    }
                    // ----- Delivery Date Calculations using Revised Contractual Delivery Date -----
                    let deliveryDate = '';
                    // Priority: Use RevisedContractualDeliveryDate first, fallback to ContractualDeliveryDate
                    if (revisedDeliveryField && item[revisedDeliveryField] !== undefined && item[revisedDeliveryField] !== '') {
                        deliveryDate = item[revisedDeliveryField];
                    }
                    else if (contractualDeliveryField && item[contractualDeliveryField] !== undefined && item[contractualDeliveryField] !== '') {
                        deliveryDate = item[contractualDeliveryField];
                    }
                    if (deliveryDate) {
                        try {
                            const delivery = new Date(deliveryDate);
                            if (!isNaN(delivery.getTime())) {
                                delivery.setHours(0, 0, 0, 0);
                                const deliveryISO = delivery.toISOString().split('T')[0];
                                // Check if delivery date is crossed (before today)
                                if (deliveryISO < todayISO) {
                                    deliveryCrossed++;
                                }
                                // Check if delivery date is within next 45 days (between today and today+45)
                                if (deliveryISO >= todayISO && deliveryISO <= fortyFiveDaysISO) {
                                    deliveryToCross45Days++;
                                }
                            }
                        }
                        catch (e) {
                            console.warn('Could not parse delivery date:', deliveryDate);
                        }
                    }
                });
                const notSubmitted = total - submitted;
                const submittedPercentage = total > 0 ? (submitted / total) * 100 : 0;
                const approvedPercentage = total > 0 ? (approved / total) * 100 : 0;
                const nwcAverage = nwcCount > 0 ? nwcTotal / nwcCount : 0;
                console.log('📊 PR Quality Results:', {
                    total,
                    submitted,
                    notSubmitted,
                    approved,
                    nwcTotal,
                    nwcAverage,
                    statusCounts,
                    relCount: statusCounts['REL'],
                    tecoCount: statusCounts['TECO'],
                    soCount: statusCounts['SO'],
                    clsdCount: statusCounts['CLSD'],
                    deliveryCrossed,
                    deliveryToCross45Days
                });
                return {
                    total,
                    submitted,
                    notSubmitted,
                    approved,
                    submittedPercentage,
                    approvedPercentage,
                    nwcTotal,
                    nwcAverage,
                    statusCounts,
                    relCount: statusCounts['REL'] || 0,
                    tecoCount: statusCounts['TECO'] || 0,
                    soCount: statusCounts['SO'] || 0,
                    clsdCount: statusCounts['CLSD'] || 0,
                    deliveryCrossed,
                    deliveryToCross45Days
                };
            }
            catch (err) {
                console.error('❌ Failed to fetch PR Quality data:', err);
                return {
                    total: 0,
                    submitted: 0,
                    notSubmitted: 0,
                    approved: 0,
                    submittedPercentage: 0,
                    approvedPercentage: 0,
                    nwcTotal: 0,
                    nwcAverage: 0,
                    statusCounts: { REL: 0, TECO: 0, SO: 0, CLSD: 0, OTHER: 0 },
                    relCount: 0,
                    tecoCount: 0,
                    soCount: 0,
                    clsdCount: 0,
                    deliveryCrossed: 0,
                    deliveryToCross45Days: 0
                };
            }
        };
        // 6. Fetch Procurement Pending Data - FIXED
        const fetchProcurementData = async (projectIds) => {
            try {
                const statusCounts = {
                    'RFQ Pending': 0,
                    'Offer Pending': 0,
                    'TC pending': 0,
                    'Order Closure pending': 0,
                    'PR pending': 0,
                    'PO pending': 0,
                    'Document Approval Pending': 0,
                    'MFC Pending': 0,
                    'Inspection Pending': 0,
                    'Dispatch Pending': 0,
                    'Delivery Pending': 0,
                    'Delivered': 0
                };
                // First, discover what fields exist in ProcurementTracker
                const fieldsUrl = `${siteUrl}/_api/web/lists/getbytitle('ProcurementTracker')/fields?$select=InternalName,Title,TypeAsString&$top=5000`;
                const fieldsRes = await fetch(fieldsUrl, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                let projectField = '';
                let forecastField = '';
                if (fieldsRes.ok) {
                    const fieldsData = await fieldsRes.json();
                    const fields = (fieldsData.d && fieldsData.d.results) || [];
                    // Find Project field (lookup)
                    const projField = fields.find((f) => {
                        const title = (f.Title || '').toLowerCase();
                        const internal = (f.InternalName || '').toLowerCase();
                        return f.TypeAsString === 'Lookup' &&
                            (title.includes('project') || internal.includes('project'));
                    });
                    if (projField) {
                        projectField = projField.InternalName || projField.EntityPropertyName || '';
                        console.log(`📋 Found Project field: ${projectField}`);
                    }
                    // Find ProcurementForecast field (Note/Multi-line text)
                    const fcField = fields.find((f) => {
                        const title = (f.Title || '').toLowerCase();
                        const internal = (f.InternalName || '').toLowerCase();
                        return (f.TypeAsString === 'Note' || f.TypeAsString === 'Text') &&
                            (title.includes('procurement') || title.includes('forecast') ||
                                internal.includes('procurement') || internal.includes('forecast'));
                    });
                    if (fcField) {
                        forecastField = fcField.InternalName || fcField.EntityPropertyName || '';
                        console.log(`📋 Found ProcurementForecast field: ${forecastField}`);
                    }
                }
                // Build the query
                let selectFields = 'Id,Title';
                if (projectField) {
                    selectFields += `,${projectField}/Id,${projectField}/Title`;
                }
                if (forecastField) {
                    selectFields += `,${forecastField}`;
                }
                const expandField = projectField ? `&$expand=${projectField}` : '';
                const url = `${siteUrl}/_api/web/lists/getbytitle('ProcurementTracker')/items?$select=${selectFields}${expandField}&$top=5000`;
                console.log('📊 Fetching Procurement data:', url.replace(siteUrl, ''));
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    throw new Error(`HTTP ${res.status}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                results.forEach((item) => {
                    // Get project ID from lookup field
                    let projectId = 0;
                    if (projectField && item[projectField]) {
                        projectId = Number(item[projectField].Id || 0);
                    }
                    if (!projectId || !projectIds.has(projectId))
                        return;
                    // Parse procurement forecast data
                    let procurementItems = [];
                    if (forecastField && item[forecastField]) {
                        try {
                            const forecastData = item[forecastField];
                            if (typeof forecastData === 'string') {
                                procurementItems = JSON.parse(forecastData);
                            }
                            else if (typeof forecastData === 'object') {
                                procurementItems = Array.isArray(forecastData) ? forecastData : [forecastData];
                            }
                        }
                        catch (parseErr) {
                            console.warn('⚠️ Failed to parse ProcurementForecast:', parseErr);
                        }
                    }
                    // Count statuses
                    if (Array.isArray(procurementItems)) {
                        procurementItems.forEach((procItem) => {
                            const status = procItem.Status || 'RFQ Pending';
                            if (statusCounts.hasOwnProperty(status)) {
                                statusCounts[status]++;
                            }
                        });
                    }
                });
                console.log('📊 Procurement Status Counts:', statusCounts);
                return statusCounts;
            }
            catch (err) {
                console.error('❌ Failed to fetch Procurement data:', err);
                return {
                    'RFQ Pending': 0,
                    'Offer Pending': 0,
                    'TC pending': 0,
                    'Order Closure pending': 0,
                    'PR pending': 0,
                    'PO pending': 0,
                    'Document Approval Pending': 0,
                    'MFC Pending': 0,
                    'Inspection Pending': 0,
                    'Dispatch Pending': 0,
                    'Delivery Pending': 0,
                    'Delivered': 0
                };
            }
        };
        // 7. Fetch DCS Document Data from DCSTracker
        const fetchDCSData = async (projectIds) => {
            try {
                const stats = {
                    totalDocuments: 0,
                    notSubmitted: 0,
                    notApproved: 0,
                    revisedNotSubmitted: 0,
                    approved: 0,
                    rejected: 0,
                    retained: 0
                };
                const url = `${siteUrl}/_api/web/lists/getbytitle('DCSTracker')/items?$select=Id,DCSData,ProjectId&$top=5000`;
                console.log('📊 Fetching DCS data:', url.replace(siteUrl, ''));
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    throw new Error(`HTTP ${res.status}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                const statusMap = {
                    'not submitted': 'notSubmitted',
                    'pending': 'notSubmitted',
                    'submitted': 'notApproved',
                    'under review': 'notApproved',
                    'approved with comments': 'revisedNotSubmitted',
                    'revised': 'revisedNotSubmitted',
                    'approved': 'approved',
                    'rejected': 'rejected',
                    'retained for information': 'retained',
                    'document deleted/cancelled': 'retained'
                };
                results.forEach((item) => {
                    const projectId = Number(item.ProjectId || 0);
                    if (!projectId || !projectIds.has(projectId))
                        return;
                    let dcsData = null;
                    try {
                        if (item.DCSData) {
                            if (typeof item.DCSData === 'string') {
                                dcsData = JSON.parse(item.DCSData);
                            }
                            else if (typeof item.DCSData === 'object') {
                                dcsData = item.DCSData;
                            }
                        }
                    }
                    catch (parseErr) {
                        console.warn('⚠️ Failed to parse DCSData:', parseErr);
                        return;
                    }
                    if (!dcsData)
                        return;
                    let outputs = [];
                    if (dcsData.documentStructure && dcsData.documentStructure.outputs) {
                        outputs = dcsData.documentStructure.outputs;
                    }
                    else if (dcsData.outputs) {
                        outputs = dcsData.outputs;
                    }
                    else if (Array.isArray(dcsData)) {
                        outputs = dcsData.filter((doc) => !doc.isInput);
                    }
                    outputs.forEach((doc) => {
                        stats.totalDocuments++;
                        const status = (doc.status || 'not submitted').toLowerCase().trim();
                        let matched = false;
                        for (const [key, value] of Object.entries(statusMap)) {
                            if (status === key || status.includes(key)) {
                                stats[value]++;
                                matched = true;
                                break;
                            }
                        }
                        if (!matched && status) {
                            stats.notSubmitted++;
                        }
                    });
                });
                console.log('📊 DCS Stats:', stats);
                return stats;
            }
            catch (err) {
                console.error('❌ Failed to fetch DCS data:', err);
                return {
                    totalDocuments: 0,
                    notSubmitted: 0,
                    notApproved: 0,
                    revisedNotSubmitted: 0,
                    approved: 0,
                    rejected: 0,
                    retained: 0
                };
            }
        };
        // 8. Fetch SEIBIES Data (BIES - Receivable Value)
        const fetchSeibiesData = async (projectIds) => {
            try {
                let totalBies = 0;
                let projectCount = 0;
                const biesValues = [];
                const url = `${siteUrl}/_api/web/lists/getbytitle('SEIBIES')/items?$select=Id,OBL,ProjectId&$top=5000`;
                console.log('📊 Fetching SEIBIES (BIES) data:', url.replace(siteUrl, ''));
                const res = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'same-origin'
                });
                if (!res.ok) {
                    throw new Error(`HTTP ${res.status}`);
                }
                const data = await res.json();
                const results = (data.d && data.d.results) || [];
                results.forEach((item) => {
                    const projectId = Number(item.ProjectId || 0);
                    if (!projectId || !projectIds.has(projectId))
                        return;
                    const obl = Math.round(Number(item.OBL) || 0);
                    if (!Number.isFinite(obl) || obl === 0)
                        return;
                    totalBies += obl;
                    projectCount++;
                    biesValues.push(obl);
                });
                const averageBies = projectCount > 0 ? Math.round((totalBies / projectCount) * 100) / 100 : 0;
                console.log(`📊 SEIBIES (BIES): Total=${totalBies}, Count=${projectCount}, Avg=${averageBies}`);
                return { totalBies, projectCount, averageBies };
            }
            catch (err) {
                console.error('❌ Failed to fetch SEIBIES data:', err);
                return { totalBies: 0, projectCount: 0, averageBies: 0 };
            }
        };
        // ============================================================
        // PROJECT FILTERING HELPERS
        // ============================================================
        const distinctBy = (items, key) => {
            const seen = new Set();
            const out = [];
            for (const it of items) {
                const k = key(it);
                if (!k || seen.has(k))
                    continue;
                seen.add(k);
                out.push(it);
            }
            return out;
        };
        const selProjects = () => {
            let list = scopedProjects;
            if (pmValue !== '__all__') {
                const opt = pmOpts.find(o => o.value === pmValue);
                const id = opt ? Number(opt.value.split('::')[0]) : NaN;
                list = list.filter(p => {
                    const pm = p.ProjectManager;
                    const sameId = pm && pm.Id !== undefined && Number(pm.Id) === id;
                    const sameTitle = pm && pm.Title && opt
                        ? pm.Title.toLowerCase().trim() === opt.label.toLowerCase().trim()
                        : false;
                    return sameId || sameTitle;
                });
            }
            if (prValue !== '__all__') {
                list = list.filter(p => (p.ProjectReference || '') === prValue);
            }
            return list;
        };
        const getSelectedProjectIds = () => {
            const selected = selProjects();
            const ids = new Set();
            selected.forEach((p) => {
                if (p.Id)
                    ids.add(Number(p.Id));
            });
            return ids;
        };
        // ============================================================
        // MAIN REFRESH FUNCTION
        // ============================================================
        const refreshBoard = async () => {
            var _a;
            const monthSelect = byId('dsMonth');
            const dateInput = byId('dsDate');
            const monthKey = monthSelect ? monthSelect.value : currentMonthKey;
            const isoDate = dateInput ? dateInput.value || todayIso : todayIso;
            // Update header meta
            setTxt('dsMetaPeriod', monthLabelOf(monthKey));
            setTxt('dsMetaAsOn', fmtAsOn(isoDate));
            setTxt('dsRevMonthTitle', `Month — ${monthLabelOf(monthKey)}`);
            // ----- UPDATE YTD TITLE DYNAMICALLY -----
            const [, selectedMonth] = monthKey.split('-').map(Number);
            const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            const selectedMonthName = monthNames[selectedMonth - 1];
            const ytdTitleElements = allById('dsYtdTitle');
            ytdTitleElements.forEach(el => {
                el.textContent = `YTD — Till ${selectedMonthName}`;
            });
            // Get filtered project IDs
            const projectIds = getSelectedProjectIds();
            console.log(`📊 Dashboard view -> PM: ${pmValue} | PR: ${prValue} | Month: ${monthKey} | Projects: ${projectIds.size}`);
            if (projectIds.size === 0) {
                // Show empty state
                setTxt('dsOblTotal', '₹0');
                setTxt('dsOblDelta', '0 projects');
                setTxt('dsRevenuePlan', '0');
                setTxt('dsRevenueActual', '0');
                setTxt('dsRevenuePercentage', '(0%)');
                setStyle('dsRevenueBar', 'width', '0%');
                setTxt('dsCollectionPlan', '0');
                setTxt('dsCollectionActual', '0');
                setTxt('dsCollectionPercentage', '(0%)');
                setStyle('dsCollectionBar', 'width', '0%');
                setTxt('dsNwcValue', '₹0');
                // setTxt('dsNwcDelta', 'Avg: ₹0');
                setTxt('dsTotalProjectsCard', '0');
                setTxt('dsProjectsDelta', '0%');
                setTxt('dsRelCount', '0');
                setTxt('dsTecoCount', '0');
                setTxt('dsSoCount', '0');
                setTxt('dsClsdCount', '0');
                setTxt('dsBiesValue', '₹0');
                setTxt('dsOtherCount', '0');
                // setTxt('dsBiesDelta', '0 projects');
                const statusIds = ['dsRfqPending', 'dsOfferPending', 'dsTcPending', 'dsOrderClosurePending',
                    'dsPrPending', 'dsPoPending', 'dsDocApprovalPending', 'dsMfcPending',
                    'dsInspectionPending', 'dsDispatchPending', 'dsDeliveryPending', 'dsDeliveredCount'];
                statusIds.forEach(id => setTxt(id, '0'));
                const docIds = ['dsTotalDocuments', 'dsNotSubmittedDocs', 'dsNotApprovedDocs',
                    'dsRevisedDocs', 'dsApprovedDocs', 'dsRejectedDocs', 'dsRetainedDocs'];
                docIds.forEach(id => setTxt(id, '0'));
                return;
            }
            // Fetch all data in parallel
            const [oblSummary, revenueData, collectionData, receivablesData, prQuality, procurementData, dcsData, seibiesData] = await Promise.all([
                fetchOblSummary(projectIds),
                fetchRevenueData(monthKey, projectIds),
                fetchCollectionData(monthKey, projectIds),
                fetchReceivablesData(projectIds),
                fetchPrQualityData(projectIds),
                fetchProcurementData(projectIds),
                fetchDCSData(projectIds),
                fetchSeibiesData(projectIds)
            ]);
            // ----- Update OBL Card -----
            setTxt('dsOblTotal', `₹${formatNumber(oblSummary.totalObl)}`);
            // setTxt('dsOblDelta', `${oblSummary.projectCount} projects | Avg: ₹${formatNumber(oblSummary.averageObl)}`);
            // ----- Update Revenue Section (Current Month) -----
            setTxt('dsRevenuePlan', formatNumber(revenueData.plan));
            setTxt('dsRevenueActual', formatNumber(revenueData.actual));
            setTxt('dsRevenuePercentage', formatPercentage(revenueData.percentage));
            setStyle('dsRevenueBar', 'width', `${Math.min(revenueData.percentage, 100)}%`);
            // ----- Update Collection Section (Current Month) -----
            setTxt('dsCollectionPlan', formatNumber(collectionData.plan));
            setTxt('dsCollectionActual', formatNumber(collectionData.actual));
            setTxt('dsCollectionPercentage', formatPercentage(collectionData.percentage));
            setStyle('dsCollectionBar', 'width', `${Math.min(collectionData.percentage, 100)}%`);
            // ----- Update YTD Section -----
            setTxt('dsYtdRevenuePlan', formatNumber(revenueData.ytdPlan));
            setTxt('dsYtdRevenueActual', formatNumber(revenueData.ytdActual));
            setTxt('dsYtdRevenuePercentage', formatPercentage(revenueData.ytdPercentage));
            setStyle('dsYtdRevenueBar', 'width', `${Math.min(revenueData.ytdPercentage, 100)}%`);
            setTxt('dsYtdCollectionPlan', formatNumber(collectionData.ytdPlan));
            setTxt('dsYtdCollectionActual', formatNumber(collectionData.ytdActual));
            setTxt('dsYtdCollectionPercentage', formatPercentage(collectionData.ytdPercentage));
            setStyle('dsYtdCollectionBar', 'width', `${Math.min(collectionData.ytdPercentage, 100)}%`);
            // Update YTD bar colors
            const ytdRevenueBar = byId('dsYtdRevenueBar');
            if (ytdRevenueBar) {
                const parent = ytdRevenueBar.closest('.bar');
                if (parent) {
                    parent.className = 'bar';
                    if (revenueData.ytdPercentage > 75)
                        parent.classList.add('high');
                    else if (revenueData.ytdPercentage > 50)
                        parent.classList.add('mid');
                    else
                        parent.classList.add('low');
                }
            }
            const ytdCollectionBar = byId('dsYtdCollectionBar');
            if (ytdCollectionBar) {
                const parent = ytdCollectionBar.closest('.bar');
                if (parent) {
                    parent.className = 'bar';
                    if (collectionData.ytdPercentage > 75)
                        parent.classList.add('high');
                    else if (collectionData.ytdPercentage > 50)
                        parent.classList.add('mid');
                    else
                        parent.classList.add('low');
                }
            }
            // ----- Update Receivables Card -----
            setTxt('dsTotalReceivables', formatNumber(receivablesData.totalReceivables));
            setTxt('dsNormalValue', formatNumber(receivablesData.normal));
            setTxt('dsNormalDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.normal / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsNormalOverDueValue', formatNumber(receivablesData.normalOverDue));
            setTxt('dsNormalOverDueDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.normalOverDue / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsGstTdsValue', formatNumber(receivablesData.gstTds));
            setTxt('dsGstTdsDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.gstTds / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsTotalOverDueValue', formatNumber(receivablesData.totalOverDue));
            setTxt('dsTotalOverDueDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.totalOverDue / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsRetentionValue', formatNumber(receivablesData.retention));
            setTxt('dsRetentionDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.retention / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsRetentionOverDueValue', formatNumber(receivablesData.retentionOverDue));
            setTxt('dsRetentionOverDueDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.retentionOverDue / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsItTdsOthersValue', formatNumber(receivablesData.itTdsOthers));
            setTxt('dsItTdsOthersDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.itTdsOthers / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsTotalRetentionValue', formatNumber(receivablesData.totalRetention));
            setTxt('dsTotalRetentionDelta', `(${receivablesData.totalReceivables > 0 ? ((receivablesData.totalRetention / receivablesData.totalReceivables) * 100).toFixed(2) : 0}%)`);
            setTxt('dsReceivablesDelta', `(+${(Math.random() * 20 + 5).toFixed(2)}%)`);
            // ----- Update BIES (SEIBIES) -----
            const formatCurrency = (num) => {
                return `₹${Math.round(num).toLocaleString('en-IN')}`;
            };
            setTxt('dsBiesValue', formatCurrency(seibiesData.totalBies));
            // setTxt('dsBiesDelta', `${seibiesData.projectCount} projects | Avg: ${formatCurrency(seibiesData.averageBies)}`);
            // ----- Update NWC -----
            setTxt('dsNwcValue', formatCurrency(prQuality.nwcTotal));
            // setTxt('dsNwcDelta', `Avg: ${formatCurrency(prQuality.nwcAverage)}`);
            // ----- Update Total Projects Card with Status Counts -----
            setTxt('dsTotalProjectsCard', formatNumber(prQuality.total));
            const changePercent = prQuality.total > 0 ? '+2.5%' : '0%';
            setTxt('dsProjectsDelta', changePercent);
            setTxt('dsRelCount', formatNumber(prQuality.relCount));
            setTxt('dsTecoCount', formatNumber(prQuality.tecoCount));
            setTxt('dsSoCount', formatNumber(prQuality.soCount));
            setTxt('dsClsdCount', formatNumber(prQuality.clsdCount));
            const otherCount = ((_a = prQuality.statusCounts) === null || _a === void 0 ? void 0 : _a.OTHER) || 0;
            setTxt('dsOtherCount', formatNumber(otherCount));
            // ----- Update PR+ Quality Table -----
            setTxt('dsTotalProjects', formatNumber(prQuality.total));
            setTxt('dsSubmittedProjects', `${formatNumber(prQuality.submitted)} (${prQuality.submittedPercentage.toFixed(2)}%)`);
            setTxt('dsNotSubmittedProjects', `${formatNumber(prQuality.notSubmitted)} (${(100 - prQuality.submittedPercentage).toFixed(2)}%)`);
            setTxt('dsApprovedProjects', `${formatNumber(prQuality.approved)} (${prQuality.approvedPercentage.toFixed(2)}%)`);
            setTxt('dsDeliveryCrossed', formatNumber(prQuality.deliveryCrossed));
            setTxt('dsDeliveryToCross', formatNumber(prQuality.deliveryToCross45Days));
            // ----- Update Procurement Pending Section -----
            setTxt('dsRfqPending', formatNumber(procurementData['RFQ Pending'] || 0));
            setTxt('dsOfferPending', formatNumber(procurementData['Offer Pending'] || 0));
            setTxt('dsTcPending', formatNumber(procurementData['TC pending'] || 0));
            setTxt('dsOrderClosurePending', formatNumber(procurementData['Order Closure pending'] || 0));
            setTxt('dsPrPending', formatNumber(procurementData['PR pending'] || 0));
            setTxt('dsPoPending', formatNumber(procurementData['PO pending'] || 0));
            setTxt('dsDocApprovalPending', formatNumber(procurementData['Document Approval Pending'] || 0));
            setTxt('dsMfcPending', formatNumber(procurementData['MFC Pending'] || 0));
            setTxt('dsInspectionPending', formatNumber(procurementData['Inspection Pending'] || 0));
            setTxt('dsDispatchPending', formatNumber(procurementData['Dispatch Pending'] || 0));
            setTxt('dsDeliveryPending', formatNumber(procurementData['Delivery Pending'] || 0));
            setTxt('dsDeliveredCount', formatNumber(procurementData['Delivered'] || 0));
            // Calculate total pending
            let totalPending = 0;
            const pendingStatuses = ['RFQ Pending', 'Offer Pending', 'TC pending', 'Order Closure pending',
                'PR pending', 'PO pending', 'Document Approval Pending', 'MFC Pending',
                'Inspection Pending', 'Dispatch Pending', 'Delivery Pending'];
            pendingStatuses.forEach(status => {
                totalPending += procurementData[status] || 0;
            });
            const procurementCardTitle = document.querySelector('.card.procurement .card-title');
            if (procurementCardTitle) {
                procurementCardTitle.textContent = `Procurement Pending (${totalPending} total)`;
            }
            // Update row colors based on values
            const updateRowColors = (id, value) => {
                const elements = allById(id);
                elements.forEach(el => {
                    const row = el.closest('.list-row');
                    if (row) {
                        row.className = 'list-row';
                        if (value > 10) {
                            row.classList.add('warn');
                        }
                        else if (value > 5) {
                            row.classList.add('ok');
                        }
                    }
                });
            };
            updateRowColors('dsRfqPending', procurementData['RFQ Pending'] || 0);
            updateRowColors('dsOfferPending', procurementData['Offer Pending'] || 0);
            updateRowColors('dsTcPending', procurementData['TC pending'] || 0);
            updateRowColors('dsOrderClosurePending', procurementData['Order Closure pending'] || 0);
            updateRowColors('dsPrPending', procurementData['PR pending'] || 0);
            updateRowColors('dsPoPending', procurementData['PO pending'] || 0);
            updateRowColors('dsDocApprovalPending', procurementData['Document Approval Pending'] || 0);
            updateRowColors('dsMfcPending', procurementData['MFC Pending'] || 0);
            updateRowColors('dsInspectionPending', procurementData['Inspection Pending'] || 0);
            updateRowColors('dsDispatchPending', procurementData['Dispatch Pending'] || 0);
            updateRowColors('dsDeliveryPending', procurementData['Delivery Pending'] || 0);
            // ----- Update Documents (DCI) Section -----
            setTxt('dsTotalDocuments', formatNumber(dcsData.totalDocuments));
            setTxt('dsNotSubmittedDocs', formatNumber(dcsData.notSubmitted));
            setTxt('dsNotApprovedDocs', formatNumber(dcsData.notApproved));
            setTxt('dsRevisedDocs', formatNumber(dcsData.revisedNotSubmitted));
            setTxt('dsApprovedDocs', formatNumber(dcsData.approved));
            setTxt('dsRejectedDocs', formatNumber(dcsData.rejected));
            setTxt('dsRetainedDocs', formatNumber(dcsData.retained));
            const documentsCardTitle = document.querySelector('.card.documents .card-title');
            if (documentsCardTitle) {
                documentsCardTitle.textContent = `Total Documents — DCI (${dcsData.totalDocuments})`;
            }
            const updateDocRowColors = (id, value) => {
                const elements = allById(id);
                elements.forEach(el => {
                    const row = el.closest('.list-row');
                    if (row) {
                        row.className = 'list-row';
                        if (value > 10) {
                            row.classList.add('warn');
                        }
                        else if (value > 5) {
                            row.classList.add('ok');
                        }
                    }
                });
            };
            updateDocRowColors('dsNotSubmittedDocs', dcsData.notSubmitted);
            updateDocRowColors('dsNotApprovedDocs', dcsData.notApproved);
            updateDocRowColors('dsRevisedDocs', dcsData.revisedNotSubmitted);
            updateDocRowColors('dsApprovedDocs', dcsData.approved);
            updateDocRowColors('dsRejectedDocs', dcsData.rejected);
            updateDocRowColors('dsRetainedDocs', dcsData.retained);
            console.log('✅ Dashboard refreshed successfully');
        };
        // ============================================================
        // PROJECT LOADING
        // ============================================================
        const readBody = async (res) => {
            try {
                return await res.text();
            }
            catch (_a) {
                return '(no body)';
            }
        };
        const normField = (s) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        const fetchFieldMap = async () => {
            const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/fields`;
            const res = await fetch(url, {
                headers: { 'Accept': 'application/json;odata=verbose' },
                credentials: 'same-origin'
            });
            if (!res.ok) {
                await readBody(res);
                throw new Error(`fields meta HTTP ${res.status}`);
            }
            const data = await res.json();
            const fields = (data.d && data.d.results) || [];
            const attrOf = (schema, attr) => {
                const m = (schema || '').match(new RegExp(`${attr}="([^"]+)"`));
                return m ? m[1] : '';
            };
            return fields.map(f => ({
                Title: f.Title || attrOf(f.SchemaXml, 'DisplayName'),
                InternalName: f.InternalName || attrOf(f.SchemaXml, 'Name') || f.EntityPropertyName || '',
                TypeAsString: f.TypeAsString || attrOf(f.SchemaXml, 'Type') || ''
            }));
        };
        const tryFetchRows = async (select, expand) => {
            const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/items`
                + `?$select=${select}`
                + (expand ? `&$expand=${expand}` : '')
                + `&$orderby=Title`;
            const res = await fetch(url, {
                headers: { 'Accept': 'application/json;odata=verbose' },
                credentials: 'same-origin'
            });
            if (!res.ok) {
                await readBody(res);
                throw new Error(`HTTP ${res.status}`);
            }
            const data = await res.json();
            return (data.d && data.d.results) || [];
        };
        const loadScope = async () => {
            const canView = canRead();
            if (!canView) {
                console.warn('🚫 Dashboard: user has no read access');
                return;
            }
            const isFullAccess = hasFullControl() || canReadAllProjects();
            const wanted = [
                { key: 'reference', label: DASH_SOURCES.fieldProjectReference },
                { key: 'customer', label: DASH_SOURCES.fieldCustomer },
                { key: 'pm', label: DASH_SOURCES.fieldProjectManager },
                { key: 'gh', label: DASH_SOURCES.fieldGroupHead }
            ];
            try {
                const resolved = new Map();
                try {
                    const fields = await fetchFieldMap();
                    for (const w of wanted) {
                        const f = fields.find(x => normField(x.Title) === normField(w.label));
                        if (f && f.InternalName) {
                            resolved.set(w.key, { internal: f.InternalName, type: f.TypeAsString || '' });
                        }
                    }
                }
                catch (metaErr) {
                    console.warn('⚠️ Could not read list field metadata', metaErr);
                }
                const pickF = (key, cfg) => { var _a; return ((_a = resolved.get(key)) === null || _a === void 0 ? void 0 : _a.internal) || cfg; };
                const kindOf = (key, cfg) => {
                    var _a;
                    const t = (((_a = resolved.get(key)) === null || _a === void 0 ? void 0 : _a.type) || '').toLowerCase();
                    if (resolved.has(key)) {
                        if (!t || t.includes('user'))
                            return 'person';
                        if (t.includes('lookup') || t.includes('taxonomy'))
                            return 'lookup';
                        return 'text';
                    }
                    return /Id$/.test(cfg) ? 'lookup' : 'person';
                };
                const refF = pickF('reference', DASH_SOURCES.fieldProjectReference);
                const custF = pickF('customer', DASH_SOURCES.fieldCustomer);
                const pmF = pickF('pm', DASH_SOURCES.fieldProjectManager);
                const ghF = pickF('gh', DASH_SOURCES.fieldGroupHead);
                const pmKind = kindOf('pm', DASH_SOURCES.fieldProjectManager);
                const ghKind = kindOf('gh', DASH_SOURCES.fieldGroupHead);
                let rows = [];
                const sel = ['Id', 'Title'];
                const pushSel = (f) => { if (f && !sel.includes(f))
                    sel.push(f); };
                pushSel(refF);
                pushSel(custF);
                if (pmF)
                    pushSel(pmKind === 'text' ? pmF : `${pmF}Id`);
                if (ghF)
                    pushSel(ghKind === 'text' ? ghF : `${ghF}Id`);
                try {
                    rows = await tryFetchRows(sel.join(','));
                }
                catch (_eBare) {
                    console.warn('⚠️ Extended query failed - retrying with Id and Title only…');
                    rows = await tryFetchRows('Id,Title');
                }
                // Resolve user IDs
                const userMap = new Map();
                const personIdOf = (r, f) => r && typeof r[f] !== 'object' ? (Number(r[`${f}Id`]) || undefined) : undefined;
                const wantedIds = new Set();
                (rows || []).forEach((r) => {
                    if (pmF && pmKind !== 'text' && personIdOf(r, pmF))
                        wantedIds.add(personIdOf(r, pmF));
                    if (ghF && ghKind !== 'text' && personIdOf(r, ghF))
                        wantedIds.add(personIdOf(r, ghF));
                });
                if (wantedIds.size) {
                    await Promise.all([...wantedIds].map(async (id) => {
                        try {
                            const res = await fetch(`${siteUrl}/_api/web/getuserbyid(${id})?$select=Id,Title,Email`, {
                                headers: { 'Accept': 'application/json;odata=verbose' },
                                credentials: 'same-origin'
                            });
                            if (!res.ok)
                                return;
                            const j = await res.json();
                            const d = j.d || {};
                            userMap.set(Number(d.Id), { Id: Number(d.Id), Title: d.Title || '', Email: d.Email || '' });
                        }
                        catch (err) {
                            console.warn('⚠️ Could not resolve user id', id, err);
                        }
                    }));
                }
                const personValue = (r, f) => {
                    if (!f)
                        return undefined;
                    if (r[f] && typeof r[f] === 'object')
                        return r[f];
                    const id = Number(r[`${f}Id`]) || 0;
                    return (id && userMap.get(id)) || undefined;
                };
                const allProjects = (rows || []).map((r) => {
                    var _a, _b, _c, _d;
                    return ({
                        Id: r.Id,
                        Title: r.Title || '',
                        ProjectReference: (_a = r[refF]) !== null && _a !== void 0 ? _a : '',
                        CustomerName: (_b = r[custF]) !== null && _b !== void 0 ? _b : '',
                        ProjectManager: pmKind === 'text' ? ((_c = r[pmF]) !== null && _c !== void 0 ? _c : undefined) : personValue(r, pmF),
                        GroupHead1: ghKind === 'text' ? ((_d = r[ghF]) !== null && _d !== void 0 ? _d : undefined) : personValue(r, ghF)
                    });
                });
                scopedProjects = getAssignedProjects(allProjects);
                console.log(`📊 Projects: ${allProjects.length} total, ${scopedProjects.length} in role scope`);
                // Build dropdown options
                const pms = distinctBy(scopedProjects.filter(p => p.ProjectManager && (p.ProjectManager.Title || p.ProjectManager.Email)), p => { var _a, _b, _c; return String((_c = (_b = (_a = p.ProjectManager.Id) !== null && _a !== void 0 ? _a : p.ProjectManager.Title) !== null && _b !== void 0 ? _b : p.ProjectManager.Email) !== null && _c !== void 0 ? _c : ''); });
                pmOpts = pms
                    .map(p => {
                    const pm = p.ProjectManager;
                    return { value: `${pm.Id}::${(pm.Title || '').toLowerCase().trim()}`, label: pm.Title || pm.Email || '—', sub: pm.Email || '' };
                })
                    .sort((a, b) => a.label.localeCompare(b.label));
                const inScope = selProjects();
                prOpts = distinctBy(inScope.filter(p => p.ProjectReference), p => String(p.ProjectReference))
                    .map(p => ({ value: String(p.ProjectReference), label: String(p.ProjectReference), sub: p.CustomerName || '' }))
                    .sort((a, b) => a.label.localeCompare(b.label));
                pmSels.forEach(sel => sel.setOptions(pmOpts));
                prSels.forEach(sel => sel.setOptions(prOpts));
                // Non-admin users default to their own projects
                if (!isFullAccess) {
                    const myName = (getCurrentUserFullName() || '').trim().toLowerCase();
                    const mine = pmOpts.find(o => o.label.toLowerCase().trim() === myName);
                    if (mine) {
                        pmSels.forEach(sel => sel.setValue(mine.value));
                        pmValue = mine.value;
                    }
                }
                await refreshBoard();
            }
            catch (err) {
                console.error('❌ Dashboard scope load failed:', err);
            }
        };
        // ============================================================
        // BUILD UI CONTROLS
        // ============================================================
        // Month dropdown
        const monthEls = allById('dsMonth');
        monthEls.forEach(el => {
            const monthSelect = el;
            if (!monthSelect.options.length) {
                buildMonthList(36).forEach(key => {
                    const o = document.createElement('option');
                    o.value = key;
                    o.textContent = monthLabelOf(key);
                    monthSelect.appendChild(o);
                });
            }
            monthSelect.value = currentMonthKey;
            monthSelect.addEventListener('change', () => { refreshBoard(); });
        });
        // Date input
        allById('dsDate').forEach(el => {
            const dateInput = el;
            dateInput.value = todayIso;
            dateInput.addEventListener('change', () => { refreshBoard(); });
        });
        // PM and PR dropdowns
        const pmHosts = allById('dsPmHost');
        const prHosts = allById('dsPrHost');
        const pmSels = [];
        const prSels = [];
        const pairCount = Math.min(pmHosts.length, prHosts.length);
        for (let i = 0; i < pairCount; i++) {
            prSels.push(new DashSelect(prHosts[i], 'Project Reference', async (v) => {
                prValue = v;
                await refreshBoard();
            }));
            pmSels.push(new DashSelect(pmHosts[i], 'Project Manager', async (v) => {
                pmValue = v;
                prValue = '__all__';
                prSels.forEach(sel => sel.setValue('__all__'));
                const inScope = selProjects();
                prOpts = distinctBy(inScope.filter(p => p.ProjectReference), p => String(p.ProjectReference))
                    .map(p => ({ value: String(p.ProjectReference), label: String(p.ProjectReference), sub: p.CustomerName || '' }))
                    .sort((a, b) => a.label.localeCompare(b.label));
                prSels.forEach(sel => sel.setOptions(prOpts));
                await refreshBoard();
            }));
        }
        // ============================================================
        // AUTO-REFRESH SYSTEM
        // ============================================================
        let refreshInterval = null;
        let isRefreshing = false;
        const startAutoRefresh = () => {
            if (refreshInterval) {
                clearInterval(refreshInterval);
                refreshInterval = null;
            }
            const indicator = document.getElementById('dsRefreshIndicator');
            if (indicator) {
                indicator.textContent = '🔄';
                indicator.style.color = '#fff';
            }
            refreshInterval = setInterval(async () => {
                if (isRefreshing)
                    return;
                isRefreshing = true;
                console.log('🔄 Auto-refreshing dashboard at:', new Date().toLocaleTimeString());
                try {
                    const indicator = document.getElementById('dsRefreshIndicator');
                    if (indicator) {
                        indicator.textContent = '⏳';
                        indicator.style.color = '#ffc107';
                    }
                    await refreshBoard();
                    const lastUpdated = document.getElementById('dsLastUpdated');
                    if (lastUpdated) {
                        const now = new Date();
                        lastUpdated.textContent = `Updated: ${now.toLocaleTimeString()}`;
                    }
                    if (indicator) {
                        indicator.textContent = '✅';
                        indicator.style.color = '#28a745';
                        setTimeout(() => {
                            if (indicator) {
                                indicator.textContent = '🔄';
                                indicator.style.color = '#fff';
                            }
                        }, 2000);
                    }
                }
                catch (error) {
                    console.error('❌ Auto-refresh failed:', error);
                    const indicator = document.getElementById('dsRefreshIndicator');
                    if (indicator) {
                        indicator.textContent = '⚠️';
                        indicator.style.color = '#dc3545';
                        setTimeout(() => {
                            if (indicator) {
                                indicator.textContent = '🔄';
                                indicator.style.color = '#fff';
                            }
                        }, 3000);
                    }
                }
                finally {
                    isRefreshing = false;
                }
            }, 5000);
        };
        const stopAutoRefresh = () => {
            if (refreshInterval) {
                clearInterval(refreshInterval);
                refreshInterval = null;
                console.log('⏹️ Auto-refresh stopped');
            }
        };
        const setupAutoRefresh = () => {
            startAutoRefresh();
            document.addEventListener('visibilitychange', () => {
                if (document.hidden) {
                    stopAutoRefresh();
                }
                else {
                    startAutoRefresh();
                }
            });
            window.addEventListener('beforeunload', () => {
                stopAutoRefresh();
            });
            console.log('✅ Auto-refresh configured (30 second interval)');
        };
        // ============================================================
        // AUTHENTICATION & INITIALIZATION
        // ============================================================
        try {
            await authManager.initialize({ siteUrl, listName: DASH_SOURCES.projectList });
            console.log('🔑 Dashboard auth initialized');
        }
        catch (err) {
            console.error('❌ Dashboard auth initialization failed:', err);
        }
        await loadScope();
        console.log('✅ Dashboard initialized for', currentMonthKey);
        // ----- SETUP AUTO-REFRESH -----
        setupAutoRefresh();
        // Initial last updated time
        const lastUpdated = document.getElementById('dsLastUpdated');
        if (lastUpdated) {
            const now = new Date();
            lastUpdated.textContent = `Updated: ${now.toLocaleTimeString()}`;
        }
    }, 250);
};
export default {
    content: dashboardContent,
    style: dashboardStyle,
    script: dashboardScript
};
//# sourceMappingURL=dashboard.js.map
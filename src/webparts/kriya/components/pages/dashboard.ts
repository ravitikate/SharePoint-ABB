import 'sweetalert2/dist/sweetalert2.min.css';
import '../styles/dashboard.css';
import {
  authManager,
  canRead,
  hasFullControl, 
  canReadAllProjects, 
  getAssignedProjects,
  getCurrentUserFullName,
  canCreateRecord,
  getPermissionLevel,
  getDigest as getAuthDigest
} from './auth';
declare const _spPageContextInfo: any;
 
// ============================================================
// DYNAMIC PROJECT PERFORMANCE DASHBOARD — COMPLETE VERSION
// ------------------------------------------------------------
//  Dynamically fetches data from:    
//    - Revenue Forecast List (Revenue Plan, Revenue Actuals, OBL)
//    - Collection List (Collection Plan, Collection Actuals)
//    - Project List (Project metadata, PR+ Quality stats)
//    - Inventory (InventoryActual)
//  All values update based on selected PM, PR, Month, and Date filters.
//
//  NOTE: All styling lives in ../styles/dashboard.css
// ============================================================

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const pad2 = (n: number): string => String(n).padStart(2, '0');
const monthKeyOf = (d: Date): string => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
const monthLabelOf = (key: string): string => {
  const [y, m] = key.split('-').map(Number);
  return `${MONTH_ABBR[(m || 1) - 1]}-${String(y).slice(2)}`;
};
const fmtAsOn = (isoDate: string): string => {
  if (!isoDate) return '';
  const d = new Date(`${isoDate}T00:00:00`);
  if (isNaN(d.getTime())) return isoDate;
  return `${pad2(d.getDate())}-${MONTH_ABBR[d.getMonth()]}-${d.getFullYear()}`;
};
const buildMonthList = (count: number): string[] => {
  const now = new Date();
  const out: string[] = [];
  for (let i = count - 1; i >= 0; i--) {
    out.push(monthKeyOf(new Date(now.getFullYear(), now.getMonth() - i, 1)));
  }
  return out;
};

// ============================================================
// TYPE DEFINITIONS
// ============================================================
interface DashOption { 
  value: string;
  label: string;
  sub?: string;
}  

interface OblSummary {
  totalObl: number;
  projectCount: number;
  averageObl: number;
  highestObl: number;
  lowestObl: number;
  previousObl: number;          // ← NEW
  oblChangePercentage: number;  // ← NEW
}

interface RevenueData {
  plan: number;
  actual: number;
  percentage: number;
   monthActual: number;
  ytdPlan: number;
  ytdActual: number;
  ytdPercentage: number;
  
}

interface CollectionData {
  plan: number;
  actual: number;
   monthActual: number;
  percentage: number;
  ytdPlan: number;
  ytdActual: number;
  ytdPercentage: number;
}

interface ReceivablesData {
  normal: number;
  normalOverDue: number;
  gstTds: number;
  totalOverDue: number;
  retention: number;
  retentionOverDue: number;
  itTdsOthers: number;
  totalRetention: number;
  totalReceivables: number;
  previousReceivables: number;         // ← NEW
  receivablesChangePercentage: number; // ← NEW
}

interface InventoryData {
  totalInventory: number;
  projectCount: number;
  averageInventory: number;
}

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

    <div class="board">
      <main class="board-main">

        <section class="row-top">
          <section class="card obl">
            <div class="card-title">Order Backlog (OBL)</div>
            <div class="kpi-num obl-num" id="dsOblTotal">0</div>
            <div class="delta muted" id="dsOblDelta">0</div>
          </section>

          <section class="card receivables">
            <div class="recv-main">
              <span class="card-title">Total Receivables</span>
              <span class="kpi-num recv-total-num" id="dsTotalReceivables">0</span>
              <div class="delta red" id="dsReceivablesDelta">0</div>
            </div>
            <div class="recv-grid">
              <div class="recv-cell">
                <div class="kpi-num val" id="dsNormalValue">0</div>
                <div class="lbl">Normal</div>
                <div class="delta" id="dsNormalDelta">0</div>
              </div>
              <div class="recv-cell">
                <div class="kpi-num val" id="dsNormalOverDueValue">0</div>
                <div class="lbl">Normal OverDue</div>
                <div class="delta" id="dsNormalOverDueDelta">0</div>
              </div>
              <div class="recv-cell">
                <div class="kpi-num val" id="dsGstTdsValue">0</div>
                <div class="lbl">GST TDS</div>
                <div class="delta" id="dsGstTdsDelta">0</div>
              </div>
              <div class="recv-cell danger col-total">
                <div class="kpi-num val" id="dsTotalOverDueValue">0</div>
                <div class="lbl">Total OverDue</div>
                <div class="delta" id="dsTotalOverDueDelta">0</div>
              </div>
              <div class="recv-cell">
                <div class="kpi-num val" id="dsRetentionValue">0</div>
                <div class="lbl">Retention</div>
                <div class="delta" id="dsRetentionDelta">0</div>
              </div>
              <div class="recv-cell">
                <div class="kpi-num val" id="dsRetentionOverDueValue">0</div>
                <div class="lbl">Retention OverDue</div>
                <div class="delta" id="dsRetentionOverDueDelta">0</div>
              </div>
              <div class="recv-cell">
                <div class="kpi-num val" id="dsItTdsOthersValue">0</div>
                <div class="lbl">IT TDS / Others</div>
                <div class="delta" id="dsItTdsOthersDelta">0</div>
              </div>
              <div class="recv-cell danger col-total">
                <div class="kpi-num val" id="dsTotalRetentionValue">0</div>
                <div class="lbl">Total Retention</div>
                <div class="delta" id="dsTotalRetentionDelta">0</div>
              </div>
            </div>
          </section>

          <section class="card projects">
            <div class="proj-head">
              <div class="proj-left">
                <span class="card-title">Total Projects</span>
                <span class="kpi-num proj-num" id="dsTotalProjectsCard">0</span> 
              </div>
              <div class="proj-chips">
                <div class="chip"><span class="n" id="dsRelCount">0</span><span class="t">REL</span></div>
                <div class="chip"><span class="n" id="dsTecoCount">0</span><span class="t">TECO</span></div>
                <div class="chip"><span class="n" id="dsSoCount">0</span><span class="t">SO</span></div>
                <div class="chip"><span class="n" id="dsOtherCount">0</span><span class="t">Other</span></div>
              </div>
            </div>
            <div class="delivery-row">
              <div class="alert alert-red">
                <div class="lbl">Delivery Date Crossed</div>
                <div class="kpi-num" id="dsDeliveryCrossed">0</div>
              </div>
              <div class="alert alert-steel">
                <div class="lbl">To Cross In 45 Days</div>
                <div class="kpi-num" id="dsDeliveryToCross">0</div>
              </div>
            </div>
          </section>
        </section>

        <!-- ============ SECTION 2 : MONTH · YTD · WORKING CAPITAL ============ -->
        <section class="row-mid">
          <section class="card revenue perf-card">
            <div class="perf-head">
              <span class="g-title" id="dsRevMonthTitle">Month — Sep-26</span>
              <span class="perf-tag">Plan vs Actual</span>
            </div>
            <div class="pair">
              <div class="pair-box">
                <div class="lbl">Revenue Plan</div>
                <div class="kpi-num plan" id="dsRevenuePlan">0</div>
              </div>
              <div class="pair-box">
                <div class="lbl">Collection Plan</div>
                <div class="kpi-num plan" id="dsCollectionPlan">0</div>
              </div>
            </div>
            <div class="perf-div"></div>
            <div class="pair">
              <div class="pair-box">
                <div class="lbl">Revenue Actual</div>
                <div class="kpi-num green actual" id="dsRevenueActual">0</div>
                <div class="delta" id="dsRevenuePercentage">0</div>
                <div class="bar low"><i style="width:0%" id="dsRevenueBar"></i></div>
              </div>
              <div class="pair-box">
                <div class="lbl">Collection Actual</div>
                <div class="kpi-num green actual" id="dsCollectionActual">0</div>
                <div class="delta" id="dsCollectionPercentage">0</div>
                <div class="bar low"><i style="width:0%" id="dsCollectionBar"></i></div>
              </div>
            </div>
          </section>

          <section class="card revenue perf-card">
            <div class="perf-head">
              <span class="g-title" id="dsYtdTitle">YTD — Till Dec</span>
              <span class="perf-tag">Plan vs Actual</span>
            </div>
            <div class="pair">
              <div class="pair-box">
                <div class="lbl">YTD Revenue Plan</div>
                <div class="kpi-num plan" id="dsYtdRevenuePlan">0</div>
              </div>
              <div class="pair-box">
                <div class="lbl">YTD Collection Plan</div>
                <div class="kpi-num plan" id="dsYtdCollectionPlan">0</div>
              </div>
            </div>
            <div class="perf-div"></div>
            <div class="pair">
              <div class="pair-box">
                <div class="lbl">YTD Revenue Actual</div>
                <div class="kpi-num green actual" id="dsYtdRevenueActual">0</div>
                <div class="delta" id="dsYtdRevenuePercentage">0</div>
                <div class="bar mid"><i style="width:0%" id="dsYtdRevenueBar"></i></div>
              </div>
              <div class="pair-box">
                <div class="lbl">YTD Collection Actual</div>
                <div class="kpi-num green actual" id="dsYtdCollectionActual">0</div>
                <div class="delta" id="dsYtdCollectionPercentage">0</div>
                <div class="bar"><i style="width:0%" id="dsYtdCollectionBar"></i></div>
              </div>
            </div>
          </section>

          <section class="card top-right working-capital">
            <div class="wc-head">
              <span class="card-title">Working Capital</span>
              <span class="cur">INR</span>
            </div>
            <div class="wc-grid">
              <div class="metric-line wc-cell">
                <div class="name">BIES</div>
                <div class="kpi-num" id="dsBiesValue">0</div>
              </div>
              <div class="metric-line wc-cell">
                <div class="name">NWC</div>
                <div class="kpi-num red" id="dsNwcValue">0</div>
              </div>
              <div class="metric-line wc-cell">
                <div class="name">SEI</div>
                <div class="kpi-num" id="dsSeiValue">0</div>
              </div>
              <div class="metric-line wc-cell">
                <div class="name">Inventory</div>
                <div class="kpi-num" id="dsInventoryValue">0</div>
              </div>
            </div>
          </section>
        </section>

        <section class="row-bot">
          <section class="card procurement">
            <div class="card-title">Procurement Pending</div>
            <div class="list proc-list" style="grid-template-columns:repeat(4,minmax(0,1fr));grid-auto-rows:minmax(52px,auto);">
              <div class="list-row"><span class="name">RFQ Pending</span><span class="val" id="dsRfqPending">0</span></div>
              <div class="list-row"><span class="name">Offer Pending</span><span class="val" id="dsOfferPending">0</span></div>
              <div class="list-row"><span class="name">TC Pending</span><span class="val" id="dsTcPending">0</span></div>
              <div class="list-row"><span class="name">Order Closure Pending</span><span class="val" id="dsOrderClosurePending">0</span></div>
              <div class="list-row"><span class="name">PR Pending</span><span class="val" id="dsPrPending">0</span></div>
              <div class="list-row"><span class="name">PO Pending</span><span class="val" id="dsPoPending">0</span></div>
              <div class="list-row"><span class="name">Document Approval Pending</span><span class="val" id="dsDocApprovalPending">0</span></div>
              <div class="list-row"><span class="name">MFC Pending</span><span class="val" id="dsMfcPending">0</span></div>
              <div class="list-row"><span class="name">Inspection Pending</span><span class="val" id="dsInspectionPending">0</span></div>
              <div class="list-row"><span class="name">Dispatch Pending</span><span class="val" id="dsDispatchPending">0</span></div>
              <div class="list-row"><span class="name">Delivery Pending</span><span class="val" id="dsDeliveryPending">0</span></div>
              <div class="list-row ok"><span class="name">Delivered</span><span class="val" id="dsDeliveredCount">0</span></div>
            </div>
          </section>

          <section class="card documents">
            <div class="card-title">Total Documents — DCI</div>
            <div class="list doc-list">
              <div class="list-row"><span class="name">Total Documents</span><span class="val" id="dsTotalDocuments">0</span></div>
              <div class="list-row"><span class="name">Not Submitted</span><span class="val" id="dsNotSubmittedDocs">0</span></div>
              <div class="list-row"><span class="name">Not Approved</span><span class="val" id="dsNotApprovedDocs">0</span></div>
              <div class="list-row doc-revised"><span class="name">Revised — Not Submitted / Approved with comments</span><span class="val" id="dsRevisedDocs">0</span></div>
              <div class="list-row ok"><span class="name">Approved</span><span class="val" id="dsApprovedDocs">0</span></div>
              <div class="list-row"><span class="name">Rejected</span><span class="val" id="dsRejectedDocs">0</span></div>
              <div class="list-row"><span class="name">Retained for Information</span><span class="val" id="dsRetainedDocs">0</span></div>
            </div>
          </section>

          <section class="card prq" style="font-size:9px;">
            <div class="card-title" style="font-size:9.5px;margin-bottom:3px;">PR+ Status</div>
            <div class="prq-sections" style="gap:6px;">
              <div class="prq-section">
                <div class="prq-section-header">
                  <span>Submission & Approval Status</span>
                  <span class="prq-total" id="dsTotalProjects">0</span>
                </div>
                <div class="prq-bar-row">
                  <div class="prq-bar-label">
                    <span>Submitted</span>
                    <span class="prq-bar-val submitted-val" id="dsSubmittedProjects">0</span>
                  </div>
                  <div class="prq-bar-track"><i class="prq-bar-fill submitted" id="dsSubmittedBar" style="width:0%"></i></div>
                  <div class="prq-bar-pct" id="dsSubmittedPct">0</div>
                </div>
                <div class="prq-bar-row">
                  <div class="prq-bar-label">  
                    <span>Not Submitted</span>
                    <span class="prq-bar-val not-submitted-val" id="dsNotSubmittedProjects">0</span>
                  </div>
                  <div class="prq-bar-track"><i class="prq-bar-fill not-submitted" id="dsNotSubmittedBar" style="width:0%"></i></div>
                  <div class="prq-bar-pct" id="dsNotSubmittedPct">0</div>
                </div>
                <div class="prq-bar-row">
                  <div class="prq-bar-label">
                    <span>Approved</span>
                    <span class="prq-bar-val approved-val" id="dsApprovedProjects">0</span>
                  </div>
                  <div class="prq-bar-track"><i class="prq-bar-fill approved" id="dsApprovedBar" style="width:0%"></i></div>
                  <div class="prq-bar-pct" id="dsApprovedPct">0</div>
                </div>
              </div>   

              <div class="prq-section">
                <div class="prq-section-header">
                  <span>Data Quality</span>
                </div> 
                <div class="prq-sub-row">
                  <span class="prq-sub-label" id="dsPrQualityMonth">0</span>
                  <div class="prq-bar-track flex-1"><i class="prq-bar-fill submitted" id="dsPrQualityBar" style="width:0%"></i></div>
                  <span class="prq-sub-pct" id="dsPrQualityAverage">0</span>
                </div>                 
              </div>
            </div>
          </section>
        </section> 
      </main>

      <aside class="card filters">
        <div class="f-cap">Filters</div>
        <div class="f-row">
          <label>Project Manager</label>
          <b class="ds-host" id="dsPmHost"></b>
        </div>
        <div class="f-row">
          <label>Date</label>
          <input type="date" class="ds-ctl" id="dsDate" readonly tabindex="-1" title="As-on date (Defaults to today)" />
        </div>
        <div class="f-row">
          <label>Month</label>
          <select class="ds-ctl" id="dsMonth" title="Data month — current and last 36 months"></select>
        </div>
        <div class="f-row">
          <label>Project Reference</label>
          <b class="ds-host" id="dsPrHost"></b>
        </div>
      </aside>
    </div>
  </div>
`;

// ============================================================
// DASHBOARD SCRIPT
// ============================================================
let dashboardRunId = 0;
let dashboardStartupTimer: number | null = null;
let dashboardCleanup: (() => void) | null = null;

const dashboardScript = (): void => {
  dashboardCleanup?.();
  dashboardCleanup = null;
  if (dashboardStartupTimer !== null) {
    window.clearTimeout(dashboardStartupTimer);
  }
  const runId = ++dashboardRunId;
  const isCurrentRun = (): boolean =>
    runId === dashboardRunId && (window as any).__dtActivePageKey === 'dashboard';

  dashboardStartupTimer = window.setTimeout(async () => {
    dashboardStartupTimer = null;
    if (!isCurrentRun()) return;

    const siteUrl = typeof _spPageContextInfo !== 'undefined'
      ? _spPageContextInfo.webAbsoluteUrl
      : window.location.origin + '/sites/Kriya';

    const DASH_SOURCES = {
      projectList: 'Project List',
      revenueList: 'Revenue Forecast List',
      collectionList: 'Collection List',
      inventoryList: 'Inventory',
      fieldProjectReference: 'ProjectReference',
      fieldCustomer: 'CustomerName',
      fieldProjectManager: 'ProjectManager',
      fieldGroupHead: 'GroupHead1',
      fieldInventoryActual: 'OBLMonth'
    };

    // ============================================================
    // UTILITY FUNCTIONS
    // ============================================================
    const byId = <T extends HTMLElement = HTMLElement>(id: string): T | null => {
      const all = document.querySelectorAll(`#${id}`);
      return (all && all.length) ? (all[all.length - 1] as T) : null;
    };

    const allById = (id: string): HTMLElement[] =>
      Array.from(document.querySelectorAll<HTMLElement>(`#${id}`));

    const setTxt = (id: string, v: string): void => {
      allById(id).forEach(el => { el.textContent = v; el.title = v; });
    };

    const setStyle = (id: string, prop: string, value: string): void => {
      allById(id).forEach(el => { (el as HTMLElement).style[prop as any] = value; });
    };

    const formatNumber = (num: number): string => {
      return Math.round(num).toLocaleString('en-US');
    };

    const formatPercentage = (num: number): string => {
      return `${num.toFixed(2)}%`;
    };

      const safeJsonParse = (jsonString: string | undefined, fallback: any = {}): any => {
      try {
        if (!jsonString || jsonString.trim() === '' || jsonString === '""' || jsonString === 'null') {
          return fallback;
        }
        return JSON.parse(jsonString);
      } catch (_error) {
        return fallback;
      }
    };

    // ⚡ SAFE FETCH: Prevents 429 Too Many Requests from crashing the dashboard
      // ⚡ SAFE FETCH: Prevents 429 Too Many Requests from crashing the dashboard
    const safeFetch = async (url: string, options: any = {}): Promise<Response> => {
      // ✅ FIX: Use native `fetch` here, NOT safeFetch!
      let res = await fetch(url, {
        credentials: 'same-origin',
        ...options,
        headers: {
          'Accept': 'application/json;odata=verbose',
          ...(options.headers || {})
        }
      });
      
      let retries = 0;
      while (res.status === 429 && retries < 5) {
        const waitSec = parseInt(res.headers.get("Retry-After") || "5", 10);
        console.warn(`⚠️ Dashboard throttled (429). Waiting ${waitSec}s...`);
        await new Promise(r => setTimeout(r, waitSec * 1000));
        
        // ✅ FIX: Use native `fetch` here too
        res = await fetch(url, {
          credentials: 'same-origin',
          ...options,
          headers: {
            'Accept': 'application/json;odata=verbose',
            ...(options.headers || {})
          }
        });
        retries++;
      }
      return res;
    };

    // ============================================================
    // DASHBOARD FILTER STATE
    // ============================================================
    const todayIso = new Date().toISOString().slice(0, 10);
    const currentMonthKey = monthKeyOf(new Date());
    let pmValue = '__all__';
    let prValue = '__all__';
    let scopedProjects: any[] = [];
    let pmOpts: DashOption[] = [];
    let prOpts: DashOption[] = [];

    // ============================================================
    // DASHBOARD SELECT CLASS
    // ============================================================
    class DashSelect {
      private root: HTMLElement;
      private btn: HTMLButtonElement;
      private panel: HTMLElement;
      private search: HTMLInputElement;
      private listEl: HTMLElement;
      private opts: DashOption[] = [];
      private selValue: string = '__all__';
      private onPick?: (value: string) => void;

      constructor(host: HTMLElement, placeholder: string, onPick?: (value: string) => void) {
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
        this.btn = host.querySelector('.ds-dd-btn') as HTMLButtonElement;
        this.panel = host.querySelector('.ds-dd-panel') as HTMLElement;
        this.search = host.querySelector('.ds-dd-search') as HTMLInputElement;
        this.listEl = host.querySelector('.ds-dd-list') as HTMLElement;

        this.btn.addEventListener('click', (e) => { e.stopPropagation(); this.toggle(); });
        this.search.addEventListener('input', () => this.renderList());
        this.search.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') {
            const first = this.listEl.querySelector<HTMLElement>('.ds-dd-item[data-value]:not([data-value="__all__"])');
            if (first) first.click();
          }
          if (e.key === 'Escape') this.close();
        });
        document.addEventListener('click', (e) => {
          if (!this.root.contains(e.target as Node)) this.close();
        });
        window.addEventListener('scroll', () => this.close(), { passive: true });
        window.addEventListener('resize', () => this.close());
      }

      setOptions(opts: DashOption[]): void {
        const seen = new Set<string>();
        this.opts = [
          { value: '__all__', label: 'All', sub: `${opts.length} option${opts.length === 1 ? '' : 's'}` },
          ...opts.filter(o => { if (seen.has(o.value)) return false; seen.add(o.value); return true; })
        ];
        if (!this.opts.some(o => o.value === this.selValue)) this.selValue = '__all__';
        if (this.search) this.search.value = '';
        this.renderList();
        this.renderButton();
      }

      getValue(): string { return this.selValue; }

      setValue(value: string): void {
        if (this.opts.some(o => o.value === value)) {
          this.selValue = value;
          this.renderButton();
          this.renderList();
        }
      }

      private toggle(): void {
        if (this.panel.classList.contains('open')) this.close();
        else this.open();
      }

      private open(): void {
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

      private close(): void {
        this.panel.classList.remove('open');
      }

      private renderButton(): void {
        const opt = this.opts.find(o => o.value === this.selValue);
        const val = this.btn.querySelector('.ds-dd-val') as HTMLElement;
        val.textContent = opt ? opt.label : 'All';
      }

      private renderList(): void {
        const term = (this.search?.value || '').trim().toLowerCase();
        const items = this.opts.filter(o => {
          if (o.value === '__all__') return true;
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
        this.listEl.querySelectorAll<HTMLElement>('.ds-dd-item').forEach(el => {
          el.addEventListener('click', (e) => {
            e.stopPropagation();
            this.selValue = el.getAttribute('data-value') || '__all__';
            this.renderButton();
            this.renderList();
            this.close();
            if (this.onPick) this.onPick(this.selValue);
          });
        });
      }
    }


    const previousMonthKeyOf = (key: string): string => {
  const [y, m] = key.split('-').map(Number);
  const d = new Date(y, (m || 1) - 1, 1);
  d.setMonth(d.getMonth() - 1);
  return monthKeyOf(d);
};

const changePct = (current: number, previous: number): number => {
  if (previous === 0) return current === 0 ? 0 : 100;
  return ((current - previous) / Math.abs(previous)) * 100;
};
    // ============================================================
    // DATA FETCHING FUNCTIONS
    // ============================================================
    const dashProjectRefFields = new Map<string, { exists: boolean; type: string }>();

    const dashListProjectRefField = async (listName: string): Promise<{ exists: boolean; type: string }> => {
      const cached = dashProjectRefFields.get(listName);
      if (cached) return cached;

      let result: { exists: boolean; type: string };
      try {
        const res = await safeFetch(
          `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq 'ProjectReference'&$select=InternalName,TypeAsString,ReadOnlyField`,
          { headers: { 'Accept': 'application/json;odata=verbose' }, credentials: 'same-origin' }
        );
        const data = await res.json();
        const results = data?.d?.results || [];
        if (results.length > 0) {
          result = { exists: true, type: String(results[0].TypeAsString || 'Text') };
          console.log(`📋 ${listName} ProjectReference column type = ${result.type}`);
        } else {
          result = { exists: false, type: '' };
          console.log(`📋 ${listName} has no ProjectReference column`);
        }
      } catch (error) {
        console.error(`⚠️ Could not check ${listName} ProjectReference column:`, error);
        result = { exists: false, type: '' };
      }

      dashProjectRefFields.set(listName, result);
      return result;
    };

    const dashFetchListItems = async (listName: string, columns: string): Promise<any[]> => {
      const refField = await dashListProjectRefField(listName);
      const linkColumn = refField.exists
        ? (refField.type === 'Lookup' ? ',ProjectReferenceId' : ',ProjectReference')
        : '';

      const attempts = [
        `Id,${columns},ProjectId${linkColumn}`,
        `Id,${columns},ProjectId`,
        `Id,${columns}`
      ];

      for (let i = 0; i < attempts.length; i++) {
        const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${attempts[i]}&$top=5000`;
        console.log(`📊 Fetching ${listName} data:`, url.replace(siteUrl, ''));
        const res = await safeFetch(url, {
          headers: { 'Accept': 'application/json;odata=verbose' },
          credentials: 'same-origin'
        });

        if (res.ok) {
          const data = await res.json();
          return (data.d && data.d.results) || [];
        }

        if (i === attempts.length - 1) {
          throw new Error(`HTTP ${res.status}`);
        }
        console.warn(`⚠️ ${listName} read with $select=${attempts[i]} → HTTP ${res.status} - retrying without the project link column`);
      }

      return [];
    };

    const dashProjectReferenceIndex = (): Map<string, number> => {
      const map = new Map<string, number>();
      (scopedProjects || []).forEach((project: any) => {
        const reference = String(project && project.ProjectReference || '').trim().toLowerCase();
        if (reference) map.set(reference, Number(project.Id));
      });
      return map;
    };

 const dashProjectIdOf = (item: any, referenceToId: Map<string, number>): number => {
  const byProjectId = Number(item.ProjectId || 0);
  if (byProjectId) return byProjectId;

  const byReferenceId = Number(item.ProjectReferenceId || 0);
  if (byReferenceId) return byReferenceId;

  const rawRef = String(item.ProjectReference?.Title || item.ProjectReference || '').trim().toLowerCase();
  if (rawRef && referenceToId.has(rawRef)) {
    return referenceToId.get(rawRef) as number;
  }

  return 0;
};

const fetchRevenueData = async (monthKey: string, projectIds: Set<number>): Promise<RevenueData> => {
  try {
    let totalPlan = 0;
    let totalActual = 0;
    let ytdPlan = 0;
    let ytdActual = 0;
    const [selectedYear, selectedMonth] = monthKey.split('-').map(Number);
    const referenceToId = dashProjectReferenceIndex();
    
    const results = await dashFetchListItems(DASH_SOURCES.revenueList, 'ForecastData');

    results.forEach((item: any) => {
      const projectId = dashProjectIdOf(item, referenceToId);
      
      // ✅ FIX: If Admin is viewing 'All', bypass the broken project link
      const isAdminAll = (pmValue === '__all__' && prValue === '__all__');
      if (!isAdminAll && (!projectId || !projectIds.has(projectId))) return;

      const forecastData = safeJsonParse(item.ForecastData);

      // --- CURRENT MONTH ---
      const monthForecast = forecastData[monthKey] || {};
      
      // Plan = 'actual' field in JSON (from your SharePoint columns)
      const plan = Number(monthForecast.actual) || 0; 
      
      // Actual = 'monthActual' field in JSON
      const actual = Number(monthForecast.monthActual) || 0; 

      totalPlan += plan;
      totalActual += actual;

      // --- YTD ---
      for (let m = 1; m <= selectedMonth; m++) {
        const monthKeyToCheck = `${selectedYear}-${String(m).padStart(2, '0')}`;
        const monthForecastYtd = forecastData[monthKeyToCheck] || {};
        
        const planYtd = Number(monthForecastYtd.actual) || 0;
        const actualYtd = Number(monthForecastYtd.monthActual) || 0;

        ytdPlan += planYtd;
        ytdActual += actualYtd;
      }
    });

    // Mathematically safe percentage
    // ✅ CHANGE 3 — 0 plan ⇒ 100 %
const percentage    = totalPlan    > 0 ? (totalActual    / totalPlan)    * 100 : 100;
const ytdPercentage = ytdPlan      > 0 ? (ytdActual      / ytdPlan)      * 100 : 100;

    return { 
      plan: totalPlan, 
      actual: totalActual, 
      monthActual: totalActual, 
      percentage, 
      ytdPlan, 
      ytdActual, 
      ytdPercentage 
    };

  } catch (err) {
    console.error('❌ Failed to fetch Revenue data:', err);
    return { plan: 0, actual: 0, monthActual: 0, percentage: 0, ytdPlan: 0, ytdActual: 0, ytdPercentage: 0 };
  }
};

const fetchCollectionData = async (monthKey: string, projectIds: Set<number>): Promise<CollectionData> => {
  try {
    let totalPlan = 0;
    let totalActual = 0;
    let ytdPlan = 0;
    let ytdActual = 0;
    const [selectedYear, selectedMonth] = monthKey.split('-').map(Number);
    const referenceToId = dashProjectReferenceIndex();
    
    const results = await dashFetchListItems(DASH_SOURCES.collectionList, 'CollectionData');

    results.forEach((item: any) => {
      const projectId = dashProjectIdOf(item, referenceToId);
      
      // ✅ FIX: If Admin is viewing 'All', bypass the broken project link
      const isAdminAll = (pmValue === '__all__' && prValue === '__all__');
      if (!isAdminAll && (!projectId || !projectIds.has(projectId))) return;

      const collectionData = safeJsonParse(item.CollectionData);
      
      // --- CURRENT MONTH ---
      const monthCollection = collectionData[monthKey] || {};
      
      const plan = Number(monthCollection.actual) || 0;
      const actual = Number(monthCollection.monthActual) || 0;

      totalPlan += plan;
      totalActual += actual;

      // --- YTD ---
      for (let m = 1; m <= selectedMonth; m++) {
        const monthKeyToCheck = `${selectedYear}-${String(m).padStart(2, '0')}`;
        const monthCollectionYtd = collectionData[monthKeyToCheck] || {};
        
        const planYtd = Number(monthCollectionYtd.actual) || 0;
        const actualYtd = Number(monthCollectionYtd.monthActual) || 0;

        ytdPlan += planYtd;
        ytdActual += actualYtd;
      }
    });

    // ✅ FIX: Force 0% if plan is 0 to prevent the 100% logic error
    const percentage = totalPlan > 0 ? (totalActual / totalPlan) * 100 : 0;
    const ytdPercentage = ytdPlan > 0 ? (ytdActual / ytdPlan) * 100 : 0;

    return { 
      plan: totalPlan, 
      actual: totalActual, 
      monthActual: totalActual, 
      percentage, 
      ytdPlan, 
      ytdActual, 
      ytdPercentage 
    };

  } catch (err) {
    console.error('❌ Failed to fetch Collection data:', err);
    return { plan: 0, actual: 0, monthActual: 0, percentage: 0, ytdPlan: 0, ytdActual: 0, ytdPercentage: 0 };
  }
};

const fetchOblSummary = async (monthKey: string, projectIds: Set<number>): Promise<OblSummary> => {
  try {
    let totalObl = 0;
    let previousObl = 0;
    let projectCount = 0;
    const oblValues: number[] = [];

    const prevMonthKey = previousMonthKeyOf(monthKey);
    const referenceToId = dashProjectReferenceIndex();
    const results = await dashFetchListItems(DASH_SOURCES.revenueList, 'OBLMonth');

    results.forEach((item: any) => {
      const projectId = dashProjectIdOf(item, referenceToId);
      if (!projectId || !projectIds.has(projectId)) return;

      const oblJson = safeJsonParse(item.OBLMonth);

      const cur = Math.round(Number(oblJson[monthKey]) || 0);
      if (Number.isFinite(cur) && cur !== 0) {
        totalObl += cur;
        projectCount++;
        oblValues.push(cur);
      }

      const prev = Math.round(Number(oblJson[prevMonthKey]) || 0);
      if (Number.isFinite(prev)) previousObl += prev;
    });

    const averageObl = projectCount > 0 ? Math.round((totalObl / projectCount) * 100) / 100 : 0;
    const highestObl = oblValues.length > 0 ? Math.max(...oblValues) : 0;
    const lowestObl  = oblValues.length > 0 ? Math.min(...oblValues) : 0;

    return {
      totalObl,
      projectCount,
      averageObl,
      highestObl,
      lowestObl,
      previousObl,
      oblChangePercentage: changePct(totalObl, previousObl)
    };

  } catch (err) {
    console.error('❌ Failed to fetch OBL data:', err);
    return {
      totalObl: 0, projectCount: 0, averageObl: 0, highestObl: 0, lowestObl: 0,
      previousObl: 0, oblChangePercentage: 0
    };
  }
};

const fetchReceivablesData = async (monthKey: string, projectIds: Set<number>): Promise<ReceivablesData> => {
  try {
    let normal = 0, normalOverDue = 0, gstTds = 0;
    let retention = 0, retentionOverDue = 0, itTdsOthers = 0;
    let previousReceivables = 0;

    const prevMonthKey = previousMonthKeyOf(monthKey);
    const referenceToId = dashProjectReferenceIndex();
    const results = await dashFetchListItems(DASH_SOURCES.collectionList, 'ReceivableValueMonth');

    results.forEach((item: any) => {
      const projectId = dashProjectIdOf(item, referenceToId);
      if (!projectId || !projectIds.has(projectId)) return;

      const recvJson = safeJsonParse(item.ReceivableValueMonth);

      // ---- current month ----
      const monthData = recvJson[monthKey] || {};
      normal          += Math.round(Number(monthData.normalDue) || 0);
      normalOverDue   += Math.round(Number(monthData.normalOverdue) || 0);
      retention       += Math.round(Number(monthData.retentionDue) || 0);
      retentionOverDue+= Math.round(Number(monthData.retentionOverdue) || 0);
      gstTds          += Math.round(Number(monthData.gstTds) || 0);
      itTdsOthers     += Math.round(Number(monthData.wht) || 0);

      // ---- previous month (for delta) ----
      const prevData = recvJson[prevMonthKey] || {};
      previousReceivables +=
        Math.round(Number(prevData.normalDue) || 0) +
        Math.round(Number(prevData.normalOverdue) || 0) +
        Math.round(Number(prevData.retentionDue) || 0) +
        Math.round(Number(prevData.retentionOverdue) || 0) +
        Math.round(Number(prevData.gstTds) || 0) +
        Math.round(Number(prevData.wht) || 0);
    });

    const totalOverDue    = normalOverDue + retentionOverDue;
    const totalRetention  = retention + retentionOverDue;
    const totalReceivables =
      normal + normalOverDue + retention + retentionOverDue + gstTds + itTdsOthers;

    return {
      normal, normalOverDue, gstTds, totalOverDue,
      retention, retentionOverDue, itTdsOthers,
      totalRetention, totalReceivables,
      previousReceivables,
      receivablesChangePercentage: changePct(totalReceivables, previousReceivables)
    };

  } catch (err) {
    console.error('❌ Failed to fetch Receivables data:', err);
    return {
      normal: 0, normalOverDue: 0, gstTds: 0,
      totalOverDue: 0, retention: 0, retentionOverDue: 0,
      itTdsOthers: 0, totalRetention: 0, totalReceivables: 0,
      previousReceivables: 0, receivablesChangePercentage: 0
    };
  }
};

    // ============================================================
    // INVENTORY — read InventoryActual from the Inventory list
    // ============================================================
   const fetchInventoryData = async (monthKey: string, projectIds: Set<number>): Promise<InventoryData> => {
  try {
    let totalInventory = 0;
    let projectCount = 0;

    // Fetch OBLMonth instead of InventoryActual
    const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.inventoryList}')/items?$select=Id,ProjectId,OBLMonth&$top=5000`;

    const res = await safeFetch(url, {
      headers: { 'Accept': 'application/json;odata=verbose' },
      credentials: 'same-origin'
    });

    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }

    const data = await res.json();
    const results = (data.d && data.d.results) || [];

    results.forEach((item: any) => {
      const projectId = Number(item.ProjectId || 0);
      if (!projectId || !projectIds.has(projectId)) return;

      // Parse the JSON string in OBLMonth
      const oblJson = safeJsonParse(item.OBLMonth);
      
      // Read the value for the selected month (e.g., "2026-11")
      const inventoryActual = Math.round(Number(oblJson[monthKey]) || 0);
      
      if (!Number.isFinite(inventoryActual) || inventoryActual === 0) return;

      totalInventory += inventoryActual;
      projectCount++;
    });

    const averageInventory = projectCount > 0
      ? Math.round((totalInventory / projectCount) * 100) / 100
      : 0;

    console.log(`📊 Inventory: Total=${totalInventory}, Count=${projectCount}, Avg=${averageInventory}`);

    return { totalInventory, projectCount, averageInventory };

  } catch (err) {
    console.error('❌ Failed to fetch Inventory data:', err);
    return { totalInventory: 0, projectCount: 0, averageInventory: 0 };
  }
};
        // ============================================================
    // OBL — sum the OBL column from the Inventory list
    // ============================================================
 

const fetchPrQualityData = async (monthKey: string, projectIds: Set<number>): Promise<{
  projectCount: number;
  total: number;
  submitted: number;
  notSubmitted: number;
  approved: number;
  submittedPercentage: number;
  approvedPercentage: number;
  dataQualityAverage: number;
  nwcTotal: number;
  nwcAverage: number;
  statusCounts: { [key: string]: number };
  relCount: number;
  tecoCount: number;
  soCount: number;
  deliveryCrossed: number;
  deliveryToCross45Days: number;
}> => {
  try {
    let projectCount = 0;
    let prTotal = 0;
    let prSubmitted = 0;
    let prApproved = 0;
    let prQualityTotal = 0;
    let prQualityCount = 0;
    let nwcTotal = 0;
    let nwcCount = 0;
    let deliveryCrossed = 0;
    let deliveryToCross45Days = 0;

    const statusCounts: { [key: string]: number } = {
      'REL': 0, 'TECO': 0, 'SO': 0, 'CLSD': 0, 'OTHER': 0, 'NA': 0
    };

    // ------------------------------------------------------------
    // Single fetch — ProjectMonth JSON + NWC column
    // ------------------------------------------------------------
    const selectString = 'Id,Title,NWC,ProjectMonth';
    const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/items?$select=${selectString}&$top=5000`;

    const res = await safeFetch(url, {
      headers: { 'Accept': 'application/json;odata=verbose' },
      credentials: 'same-origin'
    });

    if (!res.ok) {
      const errorText = await res.text();
      throw new Error(`HTTP ${res.status}: ${errorText.substring(0, 200)}`);
    }

    const data = await res.json();
    const results = (data.d && data.d.results) || [];

    // ------------------------------------------------------------
    // Date helpers (local ISO, no timezone drift)
    // ------------------------------------------------------------
    const localISO = (d: Date): string =>
      `${d.getFullYear()}-` +
      `${String(d.getMonth() + 1).padStart(2, '0')}-` +
      `${String(d.getDate()).padStart(2, '0')}`;

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayISO = localISO(today);

    const fortyFiveDaysFromNow = new Date(today);
    fortyFiveDaysFromNow.setDate(fortyFiveDaysFromNow.getDate() + 45);
    const fortyFiveDaysISO = localISO(fortyFiveDaysFromNow);

    // ------------------------------------------------------------
    // Iterate rows in scope
    // ------------------------------------------------------------
    results.forEach((item: any) => {
      if (!projectIds.has(item.Id)) return;
      projectCount++;

      // 1. Parse ProjectMonth JSON for the selected month
      const projectJson = safeJsonParse(item.ProjectMonth);
      const monthData = projectJson[monthKey] || {};

      // 2. Status (REL / TECO / SO / CLSD / OTHER / NA)
      const status = (monthData.status || '').trim().toUpperCase();
      if (status === 'REL') statusCounts['REL']++;
      else if (status === 'TECO') statusCounts['TECO']++;
      else if (status === 'SO') statusCounts['SO']++;
      else if (status === 'CLSD') statusCounts['CLSD']++;
      else if (status) statusCounts['OTHER']++;
      else statusCounts['NA']++;

      // 3. PR Quality (number or numeric string, strips % and commas)
      const rawPrQuality = monthData.prQuality;
      if (rawPrQuality !== null && rawPrQuality !== undefined && String(rawPrQuality).trim() !== '') {
        const qualityValue = typeof rawPrQuality === 'number'
          ? rawPrQuality
          : Number(String(rawPrQuality).replace(/,/g, '').replace(/%/g, '').trim());
        if (Number.isFinite(qualityValue)) {
          prQualityTotal += qualityValue;
          prQualityCount++;
        }
      }

      // 4. PR Submission / Approval (only for REL / TECO)
      const prApplicable = (status === 'REL' || status === 'TECO');
      if (prApplicable) {
        prTotal++;

        const subStr = (monthData.prSubmissionStatus || '').toLowerCase().trim();
        if (['submitted', 'approved', 'yes', 'true', '1'].includes(subStr)) {
          prSubmitted++;
        }

        const appStr = (monthData.prApprovalStatus || '').toLowerCase().trim();
        if (['approved', 'yes', 'true', '1', '2'].includes(appStr)) {
          prApproved++;
        }
      }

      // 5. NWC column (top-level list field, not JSON)
     // ✅ Prefer per-month NWC from ProjectMonth JSON; fall back to the flat NWC column
 

const rawNwc = (monthData as any).nwc;
let monthNwc: number = 0;                       // ✅ fresh, number-only
if (rawNwc !== undefined && rawNwc !== null && String(rawNwc).trim() !== '') {
  const parsed = typeof rawNwc === 'number'
    ? rawNwc
    : parseFloat(String(rawNwc).replace(/,/g, ''));
  if (!isNaN(parsed)) monthNwc = parsed;
}
nwcTotal += monthNwc;
nwcCount++;

      // 6. Delivery dates from JSON (revised overrides contractual)
      const revisedStr = monthData.revisedContractualDeliveryDate;
      const contractualStr = monthData.contractualDeliveryDate;
      const deliveryRaw = revisedStr || contractualStr;

      if (deliveryRaw) {
        const delivery = new Date(deliveryRaw);
        if (!isNaN(delivery.getTime())) {
          delivery.setHours(0, 0, 0, 0);
          const deliveryISO = localISO(delivery);

          if (deliveryISO < todayISO) deliveryCrossed++;
          if (deliveryISO >= todayISO && deliveryISO <= fortyFiveDaysISO) {
            deliveryToCross45Days++;
          }
        }
      }
    });

    // ------------------------------------------------------------
    // Derived metrics
    // ------------------------------------------------------------
    const prNotSubmitted = prTotal - prSubmitted;
    const submittedPercentage = prTotal > 0 ? (prSubmitted / prTotal) * 100 : 0;
    const approvedPercentage = prTotal > 0 ? (prApproved / prTotal) * 100 : 0;
    const dataQualityAverage = prQualityCount > 0 ? prQualityTotal / prQualityCount : 0;
    const nwcAverage = nwcCount > 0 ? nwcTotal / nwcCount : 0;

    return {
      projectCount,
      total: prTotal,
      submitted: prSubmitted,
      notSubmitted: prNotSubmitted,
      approved: prApproved,
      submittedPercentage,
      approvedPercentage,
      dataQualityAverage,
      nwcTotal,
      nwcAverage,
      statusCounts,
      relCount: statusCounts['REL'] || 0,
      tecoCount: statusCounts['TECO'] || 0,
      soCount: statusCounts['SO'] || 0,
      deliveryCrossed,
      deliveryToCross45Days
    };

  } catch (err) {
    console.error('❌ Failed to fetch PR Quality data:', err);
    return {
      projectCount: 0,
      total: 0,
      submitted: 0,
      notSubmitted: 0,
      approved: 0,
      submittedPercentage: 0,
      approvedPercentage: 0,
      dataQualityAverage: 0,
      nwcTotal: 0,
      nwcAverage: 0,
      statusCounts: { REL: 0, TECO: 0, SO: 0, CLSD: 0, OTHER: 0, NA: 0 },
      relCount: 0,
      tecoCount: 0,
      soCount: 0,
      deliveryCrossed: 0,
      deliveryToCross45Days: 0
    };
  }
}; 
    const fetchProcurementData = async (projectIds: Set<number>): Promise<{ [key: string]: number }> => {
      try {
        const statusCounts: { [key: string]: number } = {
          'RFQ Pending': 0, 'Offer Pending': 0, 'TC pending': 0,
          'Order Closure pending': 0, 'PR pending': 0, 'PO pending': 0,
          'Document Approval Pending': 0, 'MFC Pending': 0,
          'Inspection Pending': 0, 'Dispatch Pending': 0,
          'Delivery Pending': 0, 'Delivered': 0
        };

        const fieldsUrl = `${siteUrl}/_api/web/lists/getbytitle('ProcurementTracker')/fields?$select=InternalName,Title,TypeAsString&$top=5000`;
        const fieldsRes = await safeFetch(fieldsUrl, {
          headers: { 'Accept': 'application/json;odata=verbose' },
          credentials: 'same-origin'
        });

        let projectField = '';
        let forecastField = '';

        if (fieldsRes.ok) {
          const fieldsData = await fieldsRes.json();
          const fields = (fieldsData.d && fieldsData.d.results) || [];

          const projField = fields.find((f: any) => {
            const title = (f.Title || '').toLowerCase();
            const internal = (f.InternalName || '').toLowerCase();
            return f.TypeAsString === 'Lookup' && (title.includes('project') || internal.includes('project'));
          });
          if (projField) projectField = projField.InternalName || projField.EntityPropertyName || '';

          const fcField = fields.find((f: any) => {
            const title = (f.Title || '').toLowerCase();
            const internal = (f.InternalName || '').toLowerCase();
            return (f.TypeAsString === 'Note' || f.TypeAsString === 'Text') &&
              (title.includes('procurement') || title.includes('forecast') ||
                internal.includes('procurement') || internal.includes('forecast'));
          });
          if (fcField) forecastField = fcField.InternalName || fcField.EntityPropertyName || '';
        }

        let selectFields = 'Id,Title';
        if (projectField) selectFields += `,${projectField}/Id,${projectField}/Title`;
        if (forecastField) selectFields += `,${forecastField}`;

        const expandField = projectField ? `&$expand=${projectField}` : '';
        const url = `${siteUrl}/_api/web/lists/getbytitle('ProcurementTracker')/items?$select=${selectFields}${expandField}&$top=5000`;

        const res = await safeFetch(url, {
          headers: { 'Accept': 'application/json;odata=verbose' },
          credentials: 'same-origin'
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);

        const data = await res.json();
        const results = (data.d && data.d.results) || [];

        results.forEach((item: any) => {
          let projectId = 0;
          if (projectField && item[projectField]) projectId = Number(item[projectField].Id || 0);
          if (!projectId || !projectIds.has(projectId)) return;

          let procurementItems = [];
          if (forecastField && item[forecastField]) {
            try {
              const forecastData = item[forecastField];
              if (typeof forecastData === 'string') procurementItems = JSON.parse(forecastData);
              else if (typeof forecastData === 'object') procurementItems = Array.isArray(forecastData) ? forecastData : [forecastData];
            } catch (parseErr) {
              console.warn('⚠️ Failed to parse ProcurementForecast:', parseErr);
            }
          }

          if (Array.isArray(procurementItems)) {
            procurementItems.forEach((procItem: any) => {
              const status = procItem.Status || 'RFQ Pending';
              if (statusCounts.hasOwnProperty(status)) statusCounts[status]++;
            });
          }
        });

        return statusCounts;

      } catch (err) {
        console.error('❌ Failed to fetch Procurement data:', err);
        return {
          'RFQ Pending': 0, 'Offer Pending': 0, 'TC pending': 0,
          'Order Closure pending': 0, 'PR pending': 0, 'PO pending': 0,
          'Document Approval Pending': 0, 'MFC Pending': 0,
          'Inspection Pending': 0, 'Dispatch Pending': 0,
          'Delivery Pending': 0, 'Delivered': 0
        };
      }
    };

    const fetchDCSData = async (projectIds: Set<number>): Promise<{
      totalDocuments: number;
      notSubmitted: number;
      notApproved: number;
      revisedNotSubmitted: number;
      approved: number;
      rejected: number;
      retained: number;
    }> => {
      try {
        const stats = {
          totalDocuments: 0, notSubmitted: 0, notApproved: 0,
          revisedNotSubmitted: 0, approved: 0, rejected: 0, retained: 0
        };

        const url = `${siteUrl}/_api/web/lists/getbytitle('DCSTracker')/items?$select=Id,DCSData,ProjectId&$top=5000`;
        const res = await safeFetch(url, {
          headers: { 'Accept': 'application/json;odata=verbose' },
          credentials: 'same-origin'
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);

        const data = await res.json();
        const results = (data.d && data.d.results) || [];

        const statusMap: { [key: string]: keyof typeof stats } = {
          'not submitted': 'notSubmitted', 'pending': 'notSubmitted',
          'submitted': 'notApproved', 'under review': 'notApproved',
          'approved with comments': 'revisedNotSubmitted', 'revised': 'revisedNotSubmitted',
          'approved': 'approved', 'rejected': 'rejected',
          'retained for information': 'retained', 'document deleted/cancelled': 'retained'
        };

        results.forEach((item: any) => {
          const projectId = Number(item.ProjectId || 0);
          if (!projectId || !projectIds.has(projectId)) return;

          let dcsData = null;
          try {
            if (item.DCSData) {
              if (typeof item.DCSData === 'string') dcsData = JSON.parse(item.DCSData);
              else if (typeof item.DCSData === 'object') dcsData = item.DCSData;
            }
          } catch (parseErr) {
            console.warn('⚠️ Failed to parse DCSData:', parseErr);
            return;
          }

          if (!dcsData) return;

          let outputs: any[] = [];
          if (dcsData.documentStructure && dcsData.documentStructure.outputs) outputs = dcsData.documentStructure.outputs;
          else if (dcsData.outputs) outputs = dcsData.outputs;
          else if (Array.isArray(dcsData)) outputs = dcsData.filter((doc: any) => !doc.isInput);

          outputs.forEach((doc: any) => {
            stats.totalDocuments++;
            const status = (doc.status || 'not submitted').toLowerCase().trim();

            let matched = false;
            for (const [key, value] of Object.entries(statusMap)) {
              if (status === key || status.includes(key)) {
                stats[value as keyof typeof stats]++;
                matched = true;
                break;
              }
            }
            if (!matched && status) stats.notSubmitted++;
          });
        });

        return stats;

      } catch (err) {
        console.error('❌ Failed to fetch DCS data:', err);
        return {
          totalDocuments: 0, notSubmitted: 0, notApproved: 0,
          revisedNotSubmitted: 0, approved: 0, rejected: 0, retained: 0
        };
      }
    };

  // ✅ CHANGE 1 — accept monthKey, read the monthly JSON
const fetchSeibiesData = async (
  monthKey: string,               // ← NEW
  projectIds: Set<number>
): Promise<{ totalBies: number; totalSei: number }> => {
  try {
    let totalBies = 0;
    let totalSei = 0;

    // Read the monthly JSON column (adjust InternalName if yours differs)
    const url = `${siteUrl}/_api/web/lists/getbytitle('SEIBIES')/items?$select=Id,ProjectId,OBLMonth&$top=5000`;
    const res = await safeFetch(url, {
      headers: { 'Accept': 'application/json;odata=verbose' },
      credentials: 'same-origin'
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const data = await res.json();
    const results = (data.d && data.d.results) || [];

    results.forEach((item: any) => {
      const projectId = Number(item.ProjectId || 0);
      if (!projectId || !projectIds.has(projectId)) return;

      // Parse the JSON and pull the selected month
      const oblJson = safeJsonParse(item.OBLMonth);
      const obl = Math.round(Number(oblJson[monthKey]) || 0);
      if (!Number.isFinite(obl) || obl === 0) return;

      if (obl < 0) totalBies += obl;
      else totalSei += obl;
    });

    return { totalBies, totalSei };
  } catch (err) {
    console.error('❌ Failed to fetch SEIBIES data:', err);
    return { totalBies: 0, totalSei: 0 };
  }
};

    // ============================================================
    // PROJECT FILTERING HELPERS
    // ============================================================
    const distinctBy = <T,>(items: T[], key: (t: T) => string): T[] => {
      const seen = new Set<string>();
      const out: T[] = [];
      for (const it of items) {
        const k = key(it);
        if (!k || seen.has(k)) continue;
        seen.add(k);
        out.push(it);
      }
      return out;
    };

 const selProjects = (): any[] => {
  let list = scopedProjects;
  
  // If a specific PM is selected in the dropdown
  if (pmValue !== '__all__') {
    const opt = pmOpts.find(o => o.value === pmValue);
    const id = opt ? Number(opt.value.split('::')[0]) : NaN;
    list = list.filter(p => {
      const pm: any = p.ProjectManager;
      const sameId = pm && pm.Id !== undefined && Number(pm.Id) === id;
      const sameTitle = pm && pm.Title && opt
        ? pm.Title.toLowerCase().trim() === opt.label.toLowerCase().trim()
        : false;
      return sameId || sameTitle;
    });
  }
  
  // If a specific Project Reference is selected
  if (prValue !== '__all__') {
    list = list.filter(p => (p.ProjectReference || '') === prValue);
  }
  
  return list;
};

const getSelectedProjectIds = (): Set<number> => {
  const selected = selProjects();
  const ids = new Set<number>();
  selected.forEach((p: any) => { if (p.Id) ids.add(Number(p.Id)); });
  return ids;
};

    const prApplicableInScope = (
      ids: Set<number>,
      statusCounts: { [key: string]: number }
    ): boolean => {
      if (!ids || ids.size === 0) return false;
      const rel = statusCounts?.['REL'] || 0;
      const teco = statusCounts?.['TECO'] || 0;
      return (rel + teco) > 0;
    };

    // ============================================================
    // MAIN REFRESH FUNCTION
    // ============================================================
   const refreshBoard = async (): Promise<void> => {
  const monthSelect = byId<HTMLSelectElement>('dsMonth');
  
  const monthKey = monthSelect ? monthSelect.value : currentMonthKey;
  // const prevMonthLabel = monthLabelOf(previousMonthKeyOf(monthKey));
  const isoDate = todayIso;

  setTxt('dsMetaPeriod', monthLabelOf(monthKey));
  setTxt('dsMetaAsOn', fmtAsOn(isoDate));
  setTxt('dsRevMonthTitle', `Month — ${monthLabelOf(monthKey)}`);
  setTxt('dsPrQualityMonth', monthLabelOf(currentMonthKey));

  const [, selectedMonth] = monthKey.split('-').map(Number);
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const selectedMonthName = monthNames[selectedMonth - 1];

  allById('dsYtdTitle').forEach(el => {
    el.textContent = `YTD — Till ${selectedMonthName}`;
  });

  const projectIds = getSelectedProjectIds();

  // ------------------------------------------------------------
  // Empty-scope branch — no delta data available, show placeholders
  // ------------------------------------------------------------
  if (projectIds.size === 0) {
    setTxt('dsOblTotal', '0');
    setTxt('dsOblDelta', '—');              // ✅ no reference to oblSummary
    setTxt('dsReceivablesDelta', '—');      // ✅ no reference to receivablesData
    setTxt('dsTotalReceivables', '0');
    // ... all your other '0' placeholders unchanged ...
    setTxt('dsRevenuePlan', '0');
    setTxt('dsRevenueActual', '0');
    setTxt('dsRevenuePercentage', '0');
    setStyle('dsRevenueBar', 'width', '0%');
    setTxt('dsCollectionPlan', '0');
    setTxt('dsCollectionActual', '0');
    setTxt('dsCollectionPercentage', '0');
    setStyle('dsCollectionBar', 'width', '0%');
    setTxt('dsNwcValue', '0');
    setTxt('dsInventoryValue', '0');
    setTxt('dsTotalProjectsCard', '0');
    setTxt('dsRelCount', '0');
    setTxt('dsTecoCount', '0');
    setTxt('dsSoCount', '0');
    setTxt('dsBiesValue', '0');
    setTxt('dsSeiValue', '0');
    setTxt('dsOtherCount', '0');
    setTxt('dsTotalProjects', '0');
    setTxt('dsSubmittedProjects', '0');
    setTxt('dsNotSubmittedProjects', '0');
    setTxt('dsApprovedProjects', '0');
    setTxt('dsDeliveryCrossed', '0');
    setTxt('dsDeliveryToCross', '0');
    setTxt('dsSubmittedPct', '0');
    setTxt('dsNotSubmittedPct', '0');
    setTxt('dsApprovedPct', '0');
    setTxt('dsPrQualityAverage', '0');
    setStyle('dsSubmittedBar', 'width', '0%');
    setStyle('dsNotSubmittedBar', 'width', '0%');
    setStyle('dsApprovedBar', 'width', '0%');
    setStyle('dsPrQualityBar', 'width', '0%');

    const statusIds = ['dsRfqPending', 'dsOfferPending', 'dsTcPending', 'dsOrderClosurePending',
      'dsPrPending', 'dsPoPending', 'dsDocApprovalPending', 'dsMfcPending',
      'dsInspectionPending', 'dsDispatchPending', 'dsDeliveryPending', 'dsDeliveredCount'];
    statusIds.forEach(id => setTxt(id, '0'));

    const docIds = ['dsTotalDocuments', 'dsNotSubmittedDocs', 'dsNotApprovedDocs',
      'dsRevisedDocs', 'dsApprovedDocs', 'dsRejectedDocs', 'dsRetainedDocs'];
    docIds.forEach(id => setTxt(id, '0'));
    return;
  }

  // ------------------------------------------------------------
  // Fetch all sources in parallel
  // ------------------------------------------------------------
  const [
    oblSummary,
    revenueData,
    collectionData,
    receivablesData,
    prQuality,
    procurementData,
    dcsData,
    seibiesData,
    inventoryData
  ] = await Promise.all([
    fetchOblSummary(monthKey, projectIds),
    fetchRevenueData(monthKey, projectIds),
    fetchCollectionData(monthKey, projectIds),
    fetchReceivablesData(monthKey, projectIds),
    fetchPrQualityData(monthKey, projectIds),
    fetchProcurementData(projectIds),
    fetchDCSData(projectIds),
    fetchSeibiesData(monthKey, projectIds),
    fetchInventoryData(monthKey, projectIds)
  ]);

  // ------------------------------------------------------------
  // OBL + delta (declared AFTER oblSummary exists)
  // ------------------------------------------------------------
  setTxt('dsOblTotal', formatNumber(oblSummary.totalObl));
  const oblDeltaText =
    `${oblSummary.oblChangePercentage >= 0 ? '▲' : '▼'} ` +
    `${Math.abs(oblSummary.oblChangePercentage).toFixed(2)}% vs ${monthLabelOf(previousMonthKeyOf(monthKey))}`;
  setTxt('dsOblDelta', oblDeltaText);

  const oblEl = byId('dsOblDelta');
  if (oblEl) {
    oblEl.classList.remove('red', 'green', 'muted');
    oblEl.classList.add(oblSummary.oblChangePercentage >= 0 ? 'green' : 'red');
  }

  // ------------------------------------------------------------
  // Receivables + delta (declared AFTER receivablesData exists)
  // ------------------------------------------------------------
  setTxt('dsTotalReceivables', formatNumber(receivablesData.totalReceivables));
  const recvDeltaText =
    `${receivablesData.receivablesChangePercentage >= 0 ? '▲' : '▼'} ` +
    `${Math.abs(receivablesData.receivablesChangePercentage).toFixed(2)}% vs ${monthLabelOf(previousMonthKeyOf(monthKey))}`;
  setTxt('dsReceivablesDelta', recvDeltaText);

  const recvEl = byId('dsReceivablesDelta');
  if (recvEl) {
    recvEl.classList.remove('red', 'green', 'muted');
    recvEl.classList.add(receivablesData.receivablesChangePercentage >= 0 ? 'green' : 'red');
  }
 
     // ---- Revenue Month ----
  const revPlan = revenueData.plan;
  const revActual = revenueData.actual;

  setTxt('dsRevenuePlan', formatNumber(revPlan));
  setTxt('dsRevenueActual', formatNumber(revActual));

  // ✅ Plan = 0 ⇒ 100 %
  const revPct = revPlan > 0 ? revenueData.percentage : 100;
  setStyle('dsRevenueBar', 'width', `${Math.min(revPct, 100)}%`);

  const revPctEl = byId('dsRevenuePercentage');
  if (revPctEl) {
    const isAchieved = revActual >= revPlan;
    revPctEl.textContent = `${isAchieved ? '▲' : '▼'} ${formatPercentage(Math.abs(revPct))} (Plan vs Actual)`;
    revPctEl.style.color = isAchieved ? '#28a745' : '#dc3545';
  }

  // ---- Collection Month ----
  const colPlan = collectionData.plan;
  const colActual = collectionData.actual;

  setTxt('dsCollectionPlan', formatNumber(colPlan));
  setTxt('dsCollectionActual', formatNumber(colActual));

  // ✅ Plan = 0 ⇒ 100 %
  const colPct = colPlan > 0 ? collectionData.percentage : 100;
  setStyle('dsCollectionBar', 'width', `${Math.min(colPct, 100)}%`);

  const colPctEl = byId('dsCollectionPercentage');
  if (colPctEl) {
    const isAchieved = colActual >= colPlan;
    colPctEl.textContent = `${isAchieved ? '▲' : '▼'} ${formatPercentage(Math.abs(colPct))} (Plan vs Actual)`;
    colPctEl.style.color = isAchieved ? '#28a745' : '#dc3545';
  }

  // ---- Revenue YTD ----
  setTxt('dsYtdRevenuePlan', formatNumber(revenueData.ytdPlan));
  setTxt('dsYtdRevenueActual', formatNumber(revenueData.ytdActual));

  // ✅ Plan = 0 ⇒ 100 %
  const revYtdPct = revenueData.ytdPlan > 0 ? revenueData.ytdPercentage : 100;
  setStyle('dsYtdRevenueBar', 'width', `${Math.min(revYtdPct, 100)}%`);

  const revYtdPctEl = byId('dsYtdRevenuePercentage');
  if (revYtdPctEl) {
    const isAchieved = revenueData.ytdActual >= revenueData.ytdPlan;
    revYtdPctEl.textContent = `${isAchieved ? '▲' : '▼'} ${formatPercentage(Math.abs(revYtdPct))} (Plan vs Actual)`;
    revYtdPctEl.style.color = isAchieved ? '#28a745' : '#dc3545';
  }

  // ---- Collection YTD ----
  setTxt('dsYtdCollectionPlan', formatNumber(collectionData.ytdPlan));
  setTxt('dsYtdCollectionActual', formatNumber(collectionData.ytdActual));

  // ✅ Plan = 0 ⇒ 100 %
  const colYtdPct = collectionData.ytdPlan > 0 ? collectionData.ytdPercentage : 100;
  setStyle('dsYtdCollectionBar', 'width', `${Math.min(colYtdPct, 100)}%`);

  const colYtdPctEl = byId('dsYtdCollectionPercentage');
  if (colYtdPctEl) {
    const isAchieved = collectionData.ytdActual >= collectionData.ytdPlan;
    colYtdPctEl.textContent = `${isAchieved ? '▲' : '▼'} ${formatPercentage(Math.abs(colYtdPct))} (Plan vs Actual)`;
    colYtdPctEl.style.color = isAchieved ? '#28a745' : '#dc3545';
  }
    
       
          
  const ytdRevenueBar = byId<HTMLElement>('dsYtdRevenueBar');
  if (ytdRevenueBar) {
    const parent = ytdRevenueBar.closest('.bar');      
    if (parent) {
      parent.className = 'bar';
      if (revenueData.ytdPercentage > 75) parent.classList.add('high');
      else if (revenueData.ytdPercentage > 50) parent.classList.add('mid');
      else parent.classList.add('low');
    }
  }

  const ytdCollectionBar = byId<HTMLElement>('dsYtdCollectionBar');
  if (ytdCollectionBar) {
    const parent = ytdCollectionBar.closest('.bar');
    if (parent) {
      parent.className = 'bar';
      if (collectionData.ytdPercentage > 75) parent.classList.add('high');
      else if (collectionData.ytdPercentage > 50) parent.classList.add('mid');
      else parent.classList.add('low');
    }
  }

  // ---- receivables breakdown cells (share-of-total percentages) ----
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
  // ✅ NOTE: removed the trailing `setTxt('dsReceivablesDelta', '—')` that was overwriting the real delta.

  // ---- working capital ----
  const formatCurrency = (num: number): string => `₹${Math.round(num).toLocaleString('en-US')}`;
// const biesValue = seibiesData.totalBies < 0
//   ? `-₹${Math.abs(Math.round(seibiesData.totalBies)).toLocaleString('en-US')}`
//   : formatCurrency(0);
  setTxt('dsBiesValue', formatCurrency(Math.abs(seibiesData.totalBies)));
  setTxt('dsSeiValue', formatCurrency(seibiesData.totalSei));
  setTxt('dsNwcValue', formatCurrency(prQuality.nwcTotal));
  setTxt('dsInventoryValue', formatCurrency(inventoryData.totalInventory));

  // ---- project counts / PR+ / delivery ----
  setTxt('dsTotalProjectsCard', formatNumber(prQuality.projectCount));
  setTxt('dsRelCount', formatNumber(prQuality.relCount));
  setTxt('dsTecoCount', formatNumber(prQuality.tecoCount));
  setTxt('dsSoCount', formatNumber(prQuality.soCount));
  setTxt('dsOtherCount', formatNumber(prQuality.statusCounts?.OTHER || 0));
  setTxt('dsPrQualityAverage', `${prQuality.dataQualityAverage.toFixed(2)}%`);
  setStyle('dsPrQualityBar', 'width', `${Math.max(0, Math.min(prQuality.dataQualityAverage, 100))}%`);

  if (prApplicableInScope(projectIds, prQuality.statusCounts)) {
    const total = prQuality.total || 0;
    const submitted = prQuality.submitted || 0;
    const notSubmitted = prQuality.notSubmitted || 0;
    const approved = prQuality.approved || 0;

    setTxt('dsTotalProjects', formatNumber(total));
    setTxt('dsSubmittedProjects', formatNumber(submitted));
    setTxt('dsNotSubmittedProjects', formatNumber(notSubmitted));
    setTxt('dsApprovedProjects', formatNumber(approved));

    const submittedPct = total > 0 ? (submitted / total) * 100 : 0;
    const notSubmittedPct = total > 0 ? (notSubmitted / total) * 100 : 0;
    const approvedPct = total > 0 ? (approved / total) * 100 : 0;

    setTxt('dsSubmittedPct', `${submittedPct.toFixed(2)}%`);
    setTxt('dsNotSubmittedPct', `${notSubmittedPct.toFixed(2)}%`);
    setTxt('dsApprovedPct', `${approvedPct.toFixed(2)}%`);

    setStyle('dsSubmittedBar', 'width', `${Math.min(submittedPct, 100)}%`);
    setStyle('dsNotSubmittedBar', 'width', `${Math.min(notSubmittedPct, 100)}%`);
    setStyle('dsApprovedBar', 'width', `${Math.min(approvedPct, 100)}%`);
  } else {
    setTxt('dsTotalProjects', '0');
    setTxt('dsSubmittedProjects', '0');
    setTxt('dsNotSubmittedProjects', '0');
    setTxt('dsApprovedProjects', '0');
    setTxt('dsSubmittedPct', '0.00%');
    setTxt('dsNotSubmittedPct', '0.00%');
    setTxt('dsApprovedPct', '0.00%');
    setStyle('dsSubmittedBar', 'width', '0%');
    setStyle('dsNotSubmittedBar', 'width', '0%');
    setStyle('dsApprovedBar', 'width', '0%');
  }

  setTxt('dsDeliveryCrossed', formatNumber(prQuality.deliveryCrossed));
  setTxt('dsDeliveryToCross', formatNumber(prQuality.deliveryToCross45Days));

  // ---- procurement ----
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

  // ---- documents ----
  setTxt('dsTotalDocuments', formatNumber(dcsData.totalDocuments));
  setTxt('dsNotSubmittedDocs', formatNumber(dcsData.notSubmitted));
  setTxt('dsNotApprovedDocs', formatNumber(dcsData.notApproved));
  setTxt('dsRevisedDocs', formatNumber(dcsData.revisedNotSubmitted));
  setTxt('dsApprovedDocs', formatNumber(dcsData.approved));
  setTxt('dsRejectedDocs', formatNumber(dcsData.rejected));
  setTxt('dsRetainedDocs', formatNumber(dcsData.retained));
};

    // ============================================================
    // PROJECT LOADING
    // ============================================================
    const readBody = async (res: Response): Promise<string> => {
      try { return await res.text(); } catch { return '(no body)'; }
    };

    const normField = (s: string): string => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

    const fetchFieldMap = async (): Promise<Array<{ Title: string; InternalName: string; TypeAsString: string }>> => {
      const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/fields`;
      const res = await safeFetch(url, {
        headers: { 'Accept': 'application/json;odata=verbose' },
        credentials: 'same-origin'
      });
      if (!res.ok) {
        await readBody(res);
        throw new Error(`fields meta HTTP ${res.status}`);
      }
      const data = await res.json();
      const fields: any[] = (data.d && data.d.results) || [];
      const attrOf = (schema: string, attr: string): string => {
        const m = (schema || '').match(new RegExp(`${attr}="([^"]+)"`));
        return m ? m[1] : '';
      };
      return fields.map(f => ({
        Title: f.Title || attrOf(f.SchemaXml, 'DisplayName'),
        InternalName: f.InternalName || attrOf(f.SchemaXml, 'Name') || f.EntityPropertyName || '',
        TypeAsString: f.TypeAsString || attrOf(f.SchemaXml, 'Type') || ''
      }));
    };

    const tryFetchRows = async (select: string, expand?: string): Promise<any[]> => {
      const url = `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/items`
        + `?$select=${select}`
        + (expand ? `&$expand=${expand}` : '')
        + `&$orderby=Title&$top=5000`;
      const res = await safeFetch(url, {
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

      const loadScope = async (): Promise<void> => {
      const canView = canRead();
      if (!canView) return;

      const isFullAccess = hasFullControl() || canReadAllProjects();
      const wanted = [
        { key: 'reference',         label: DASH_SOURCES.fieldProjectReference },
        { key: 'customer',          label: DASH_SOURCES.fieldCustomer },
        { key: 'pm',                label: DASH_SOURCES.fieldProjectManager },
        { key: 'gh',                label: DASH_SOURCES.fieldGroupHead },
        // ✅ Engineering role fields (used by ROLE_OWN_FIELDS in auth.ts)
        { key: 'enggLead',          label: 'EnggLead' },
        { key: 'enggSubLead1',      label: 'EnggSubLead1' },
        { key: 'enggSubLead2',      label: 'EnggSubLead2' },
        { key: 'enggManagers',      label: 'EngineeringManagers' }
      ];

      try {
        const resolved = new Map<string, { internal: string; type: string }>();
        try {
          const fields = await fetchFieldMap();
          for (const w of wanted) {
            const f = fields.find(x => normField(x.Title) === normField(w.label));
            if (f && f.InternalName) {
              resolved.set(w.key, { internal: f.InternalName, type: f.TypeAsString || '' });
            }
          }
        } catch (metaErr) {
          console.warn('⚠️ Could not read list field metadata', metaErr);
        }

        const pickF = (key: string, cfg: string): string => (resolved.get(key)?.internal) || cfg;
        const kindOf = (key: string, cfg: string): string => {
          const t = (resolved.get(key)?.type || '').toLowerCase();
          if (resolved.has(key)) {
            if (!t || t.includes('user')) return 'person';
            if (t.includes('lookup') || t.includes('taxonomy')) return 'lookup';
            return 'text';
          }
          return /Id$/.test(cfg) ? 'lookup' : 'person';
        };

        const refF  = pickF('reference', DASH_SOURCES.fieldProjectReference);
        const custF = pickF('customer',  DASH_SOURCES.fieldCustomer);
        const pmF   = pickF('pm',        DASH_SOURCES.fieldProjectManager);
        const ghF   = pickF('gh',        DASH_SOURCES.fieldGroupHead);
        // ✅ Engineering role fields
        const elF   = pickF('enggLead',      'EnggLead');
        const es1F  = pickF('enggSubLead1',  'EnggSubLead1');
        const es2F  = pickF('enggSubLead2',  'EnggSubLead2');
        const emF   = pickF('enggManagers',  'EngineeringManagers');

        const pmKind  = kindOf('pm',        DASH_SOURCES.fieldProjectManager);
        const ghKind  = kindOf('gh',        DASH_SOURCES.fieldGroupHead);
        const elKind  = kindOf('enggLead',  'EnggLead');
        const es1Kind = kindOf('enggSubLead1', 'EnggSubLead1');
        const es2Kind = kindOf('enggSubLead2', 'EnggSubLead2');
        const emKind  = kindOf('enggManagers', 'EngineeringManagers');

        let rows: any[] = [];
        const sel: string[] = ['Id', 'Title'];
        const pushSel = (f: string): void => { if (f && !sel.includes(f)) sel.push(f); };
        pushSel(refF);
        pushSel(custF);
        if (pmF)  pushSel(pmKind  === 'text' ? pmF  : `${pmF}Id`);
        if (ghF)  pushSel(ghKind  === 'text' ? ghF  : `${ghF}Id`);
        if (elF)  pushSel(elKind  === 'text' ? elF  : `${elF}Id`);
        if (es1F) pushSel(es1Kind === 'text' ? es1F : `${es1F}Id`);
        if (es2F) pushSel(es2Kind === 'text' ? es2F : `${es2F}Id`);
        if (emF)  pushSel(emKind  === 'text' ? emF  : `${emF}Id`);

        try {
          rows = await tryFetchRows(sel.join(','));
        } catch (_eBare) {
          console.warn('⚠️ Extended query failed - retrying with Id and Title only…');
          rows = await tryFetchRows('Id,Title');
        }

        const userMap = new Map<number, { Id: number; Title: string; Email: string }>();
        const personIdOf = (r: any, f: string): number | undefined =>
          r && typeof r[f] !== 'object' ? (Number(r[`${f}Id`]) || undefined) : undefined;
        const wantedIds = new Set<number>();
        
        (rows || []).forEach((r: any) => {
          if (pmF  && pmKind  !== 'text' && personIdOf(r, pmF))  wantedIds.add(personIdOf(r, pmF)  as number);
          if (ghF  && ghKind  !== 'text' && personIdOf(r, ghF))  wantedIds.add(personIdOf(r, ghF)  as number);
          if (elF  && elKind  !== 'text' && personIdOf(r, elF))  wantedIds.add(personIdOf(r, elF)  as number);
          if (es1F && es1Kind !== 'text' && personIdOf(r, es1F)) wantedIds.add(personIdOf(r, es1F) as number);
          if (es2F && es2Kind !== 'text' && personIdOf(r, es2F)) wantedIds.add(personIdOf(r, es2F) as number);
          if (emF  && emKind  !== 'text' && personIdOf(r, emF))  wantedIds.add(personIdOf(r, emF)  as number);
        });

        if (wantedIds.size) {
          await Promise.all([...wantedIds].map(async (id) => {
            try {
              // ⚡ Cleaned up: safeFetch already handles headers and credentials!
              const res = await safeFetch(`${siteUrl}/_api/web/getuserbyid(${id})?$select=Id,Title,Email`);
              if (!res.ok) return;
              const j = await res.json();
              const d = j.d || {};
              userMap.set(Number(d.Id), { Id: Number(d.Id), Title: d.Title || '', Email: d.Email || '' });
            } catch (err) {
              console.warn('⚠️ Could not resolve user id', id, err);
            }
          }));
        }

        const personValue = (r: any, f: string): any => {
          if (!f) return undefined;
          if (r[f] && typeof r[f] === 'object') return r[f];
          const id = Number(r[`${f}Id`]) || 0;
          return (id && userMap.get(id)) || undefined;
        };

        const allProjects = (rows || []).map((r: any) => ({
          Id: r.Id,
          Title: r.Title || '',
          ProjectReference: r[refF]  ?? '',
          CustomerName:     r[custF] ?? '',
          ProjectManager:   pmKind  === 'text' ? (r[pmF]  ?? undefined) : personValue(r, pmF),
          GroupHead1:       ghKind  === 'text' ? (r[ghF]  ?? undefined) : personValue(r, ghF),
          // ✅ Engineering role fields
          EnggLead:           elKind  === 'text' ? (r[elF]  ?? undefined) : personValue(r, elF),
          EnggSubLead1:       es1Kind === 'text' ? (r[es1F] ?? undefined) : personValue(r, es1F),
          EnggSubLead2:       es2Kind === 'text' ? (r[es2F] ?? undefined) : personValue(r, es2F),
          EngineeringManagers: emKind === 'text' ? (r[emF]  ?? undefined) : personValue(r, emF)
        }));

        scopedProjects = getAssignedProjects(allProjects);

        const pms = distinctBy(
          scopedProjects.filter(p => p.ProjectManager && (p.ProjectManager.Title || p.ProjectManager.Email)),
          p => String((p.ProjectManager as any).Id ?? (p.ProjectManager as any).Title ?? (p.ProjectManager as any).Email ?? '')
        );
        pmOpts = pms
          .map(p => {
            const pm: any = p.ProjectManager;
            return { value: `${pm.Id}::${(pm.Title || '').toLowerCase().trim()}`, label: pm.Title || pm.Email || '—', sub: pm.Email || '' };
          })
          .sort((a, b) => a.label.localeCompare(b.label));

        const inScope = selProjects();
        prOpts = distinctBy(inScope.filter(p => p.ProjectReference), p => String(p.ProjectReference))
          .map(p => ({ value: String(p.ProjectReference), label: String(p.ProjectReference), sub: p.CustomerName || '' }))
          .sort((a, b) => a.label.localeCompare(b.label));

        pmSels.forEach(sel => sel.setOptions(pmOpts));
        prSels.forEach(sel => sel.setOptions(prOpts));

        if (!isFullAccess) {
          const myName = (getCurrentUserFullName() || '').trim().toLowerCase();
          const mine = pmOpts.find(o => o.label.toLowerCase().trim() === myName);
          if (mine) {
            pmSels.forEach(sel => sel.setValue(mine.value));
            pmValue = mine.value;
          }
        }

        await refreshBoard();
      } catch (err) {
        console.error('❌ Dashboard scope load failed:', err);
      }
    };

    // ============================================================
    // BUILD UI CONTROLS
    // ============================================================
    const monthEls = allById('dsMonth');
    monthEls.forEach(el => {
      const monthSelect = el as HTMLSelectElement;
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

    allById('dsDate').forEach(el => {
      const dateInput = el as HTMLInputElement;
      dateInput.readOnly = true;
      dateInput.tabIndex = -1;
      dateInput.setAttribute('readonly', 'readonly');
      dateInput.setAttribute('aria-readonly', 'true');
      dateInput.value = todayIso;

      ['mousedown', 'click'].forEach(evt => {
        dateInput.addEventListener(evt, (e) => { e.preventDefault(); });
      });

      ['keydown', 'keypress', 'paste', 'cut', 'drop', 'input', 'change'].forEach(evt => {
        dateInput.addEventListener(evt, (e) => {
          e.preventDefault();
          dateInput.value = todayIso;
        });
      });

      dateInput.addEventListener('blur', () => {
        if (dateInput.value !== todayIso) {
          dateInput.value = todayIso;
          refreshBoard();
        }
      });
    });

    const pmHosts = allById('dsPmHost');
    const prHosts = allById('dsPrHost');
    const pmSels: DashSelect[] = [];
    const prSels: DashSelect[] = [];
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
    let refreshInterval: ReturnType<typeof setInterval> | null = null;
    let isRefreshing = false;
    let backgroundSyncTimer: number | null = null;
    const handleVisibilityChange = (): void => {
      if (document.hidden || !isCurrentRun()) stopAutoRefresh();
      else startAutoRefresh();
    };
    const handlePageDeactivated = (event: Event): void => {
      const detail = (event as CustomEvent<{ pageKey: string }>).detail;
      if (detail?.pageKey !== 'dashboard') return;
      stopAutoRefresh();
      if (backgroundSyncTimer !== null) {
        window.clearTimeout(backgroundSyncTimer);
        backgroundSyncTimer = null;
      }
    };
    const handleBeforeUnload = (): void => stopAutoRefresh();

    const startAutoRefresh = (): void => {
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
        if (!isCurrentRun()) {
          stopAutoRefresh();
          return;
        }
        if (isRefreshing) return;
        isRefreshing = true;

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
            }, 100000);
          }
        } catch (error) {
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
            }, 150000);
          }
        } finally {
          isRefreshing = false;
        }
      }, 900000);
    };

    const stopAutoRefresh = (): void => {
      if (refreshInterval) {
        clearInterval(refreshInterval);
        refreshInterval = null;
      }
    };

    const setupAutoRefresh = (): void => {
      startAutoRefresh();
      document.addEventListener('visibilitychange', handleVisibilityChange);
      document.addEventListener('dt-page-deactivated', handlePageDeactivated);
      window.addEventListener('beforeunload', handleBeforeUnload);
    };

    // ============================================================
    // AUTHENTICATION & INITIALIZATION
    // ============================================================
    try {
      await authManager.initialize({ siteUrl, listName: DASH_SOURCES.projectList });
      console.log('🔑 Dashboard auth initialized');
    } catch (err) {
      console.error('❌ Dashboard auth initialization failed:', err);
    }

    if (!isCurrentRun()) return;
    await loadScope();
    if (!isCurrentRun()) return;
    console.log('✅ Dashboard initialized for', currentMonthKey);
    setupAutoRefresh();

    const lastUpdated = document.getElementById('dsLastUpdated');
    if (lastUpdated) {
      const now = new Date();
      lastUpdated.textContent = `Updated: ${now.toLocaleTimeString()}`;
    }

    // ============================================================
    // BACKGROUND AUTO-SYNC + AUTO-PROGRESSION
    // ============================================================
    interface DashSyncTarget {
      key: string;
      label: string;
      listName: string;
      entityFallback: string;
      jsonField: string;
      monthlyKind: 'standard' | 'revenue' | 'collection';
      fresh: () => any;
      buildPayload: (
        project: any,
        fresh: any,
        entityType: string,
        attempt: number,
        refField: { exists: boolean; type: string }
      ) => any;
    }

    const dashMonthKeyOf = (d: Date): string =>
      `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;

    const dashIdealMonths = (kind: 'standard' | 'revenue' | 'collection'): string[] => {
      const now = new Date();
      const y = now.getFullYear();
      const m = now.getMonth();
      const keys: string[] = [];
      const push = (offset: number): void => { keys.push(dashMonthKeyOf(new Date(y, m + offset, 1))); };

      if (kind === 'collection') {
        for (let i = -3; i < 0; i++) push(i);
        for (let i = -36; i < -3; i++) push(i);
        for (let i = 0; i < 3; i++) push(i);
        for (let i = 3; i < 19; i++) push(i);
        return keys.sort();
      }

      if (kind === 'revenue') {
        for (let i = -36; i < 0; i++) push(i);
        for (let i = 0; i < 3; i++) push(i);
        for (let i = 3; i < 18; i++) push(i);
        return Array.from(new Set(keys)).sort();
      }

      for (let i = -3; i < 0; i++) push(i);
      for (let i = 0; i < 3; i++) push(i);
      for (let i = 3; i < 18; i++) push(i);
      return keys;
    };

    const dashEmptyItems = (): any => {
      const items: any = {};
      for (let i = 1; i <= 10; i++) {
        items[i.toString()] = { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
      }
      return items;
    };

    const dashEmptyCategories = (): any => ({
      'Normal': { name: 'Normal', plan: 0, actual: 0, forecast: 0, date: '' },
      'Normal Overdue': { name: 'Normal Overdue', plan: 0, actual: 0, forecast: 0, date: '' },
      'Retention': { name: 'Retention', plan: 0, actual: 0, forecast: 0, date: '' },
      'Retention Overdue': { name: 'Retention Overdue', plan: 0, actual: 0, forecast: 0, date: '' }
    });

    const dashSimpleCategories = (): any => ({
      'Normal': { name: 'Normal', plan: 0, actual: 0, forecast: 0, date: '' },
      'Retention': { name: 'Retention', plan: 0, actual: 0, forecast: 0, date: '' }
    });

    const dashFreshYearly = (lastMonthKey?: string): any => {
      const yearly: any = {};
      let startYear: number;
      if (lastMonthKey) {
        startYear = parseInt(lastMonthKey.split('-')[0]);
      } else {
        const futureDate = new Date();
        futureDate.setMonth(futureDate.getMonth() + 15);
        startYear = futureDate.getFullYear();
      }
      for (let year = startYear; year <= startYear + 4; year++) {
        yearly[year.toString()] = { plan: 0, actual: 0 };
      }
      return yearly;
    };

    const dashFreshData = (kind: 'standard' | 'revenue' | 'collection', withItems: boolean): any => {
      const months = dashIdealMonths(kind);
      const currentKey = dashMonthKeyOf(new Date());
      const currentIndex = months.indexOf(currentKey);
      const detailedStart = currentIndex !== -1 ? currentIndex : (kind === 'revenue' ? 36 : 0);
      const fresh: any = {};

      months.forEach((monthKey, index) => {
        const isDetailed = index >= detailedStart && index < detailedStart + 3;
        const monthData: any = { plan: 0, actual: 0, forecast: 0 };

        if (kind === 'revenue') {
          monthData.monthActual = 0;
          monthData.items = isDetailed ? dashEmptyItems() : {};
        } else if (kind === 'collection') {
          monthData.monthActual = 0;
          monthData.date = '';
          monthData.items = isDetailed
            ? dashEmptyCategories()
            : (index >= detailedStart + 3 ? dashSimpleCategories() : {});
        } else {
          monthData.date = monthKey === currentKey ? '' : undefined;
          if (withItems) {
            monthData.items = (index >= 3 && index <= 5) ? dashEmptyItems() : {};
          }
        }

        fresh[monthKey] = monthData;
      });

      fresh.yearly = dashFreshYearly(months[months.length - 1]);
      return fresh;
    };

    const dashNeedsProgression = (data: any, kind: 'standard' | 'revenue' | 'collection'): boolean => {
      if (!data || typeof data !== 'object') return true;
      const months = Object.keys(data).filter(key => /^\d{4}-\d{2}$/.test(key)).sort();
      if (months.length === 0) return true;

      const idealMonths = dashIdealMonths(kind);
      const monthlyProgressionNeeded = months[0] !== idealMonths[0];

      const currentYear = new Date().getFullYear();
      let yearlyProgressionNeeded = false;

      if (data.yearly) {
        const years = Object.keys(data.yearly).map(y => parseInt(y, 10)).sort();
        yearlyProgressionNeeded = years.length > 0 && years[0] < currentYear;
      } else {
        yearlyProgressionNeeded = true;
      }

      return monthlyProgressionNeeded || yearlyProgressionNeeded;
    };

    const dashRollForward = (data: any, kind: 'standard' | 'revenue' | 'collection'): any => {
      const idealMonths = dashIdealMonths(kind);
      const currentKey = dashMonthKeyOf(new Date());
      const originalMonths = Object.keys(data || {}).filter(key => /^\d{4}-\d{2}$/.test(key)).sort();

      if (originalMonths.length > 0 && originalMonths[0] === idealMonths[0]) return data;

      const updated: any = {};
      idealMonths.forEach(monthKey => {
        const existing = data[monthKey];
        const isMonthData = existing && typeof existing === 'object' &&
          typeof existing.plan === 'number' && typeof existing.actual === 'number';

        updated[monthKey] = isMonthData
          ? JSON.parse(JSON.stringify(existing))
          : { plan: 0, actual: 0 };
      });

      const currentIndex = idealMonths.indexOf(currentKey);
      idealMonths.forEach((monthKey, index) => {
        const monthData = updated[monthKey];
        if (monthData && typeof monthData.plan === 'number' && typeof monthData.actual === 'number') {
          if (index >= currentIndex) {
            monthData.plan = monthData.actual;
            monthData.actual = 0;
          }
        }
      });

      if (data.yearly) {
        updated.yearly = JSON.parse(JSON.stringify(data.yearly));
      } else {
        updated.yearly = dashFreshYearly(idealMonths[idealMonths.length - 1]);
      }

      if (originalMonths[0] !== idealMonths[0]) {
        const rolled: any = {};
        Object.keys(updated.yearly).forEach(year => {
          const yearData = updated.yearly[year];
          if (yearData) rolled[year] = { plan: yearData.actual, actual: 0 };
        });
        updated.yearly = rolled;
      }

      return updated;
    };

    const dashLog = (...parts: any[]): void => { console.log('[DASH-SYNC]', ...parts); };

    const dashStatusChip = (): HTMLElement | null => {
      let chip = document.getElementById('dsBgSyncStatus');
      if (chip) return chip;

      const meta = document.querySelector('.header-meta') || document.querySelector('.header');
      if (!meta) return null;

      chip = document.createElement('span');
      chip.id = 'dsBgSyncStatus';
      chip.style.cssText = 'font-size:10px;color:#9ecbff;margin-left:8px;white-space:nowrap;';
      chip.textContent = '';
      meta.appendChild(chip);
      return chip;
    };

    const setDashSyncStatus = (text: string, color: string = '#9ecbff'): void => {
      const chip = dashStatusChip();
      if (chip) { chip.textContent = text; chip.style.color = color; }
    };

    const dashEnsureSiteLevelAuth = async (): Promise<void> => {
      try {
        dashLog('re-initialising shared auth without a list name…');
        await authManager.initialize({ siteUrl });
        dashLog('✅ site-level auth ready');
      } catch (error) {
        console.error('[DASH-SYNC] ❌ could not re-initialise auth:', error);
      }
    };

    const dashPermissionReport = (): string => {
      const safe = (name: string, fn: any): string => {
        if (typeof fn !== 'function') return `${name}=<not exported by ./auth>`;
        try { return `${name}=${fn()}`; } catch (error) { return `${name}=<error:${(error as Error).message}>`; }
      };
      return [
        safe('level', getPermissionLevel),
        safe('read', canRead),
        safe('readAll', canReadAllProjects),
        safe('create', canCreateRecord)
      ].join(' ');
    };

    const dashCanWrite = (): boolean => {
      try {
        if (typeof canCreateRecord !== 'function') {
          console.error('[DASH-SYNC] ❌ canCreateRecord is not exported by ./auth');
          return false;
        }
        return canCreateRecord();
      } catch (error) {
        console.error('[DASH-SYNC] ❌ create-right check threw:', error);
        return false;
      }
    };

    const dashScopeProjects = (): any[] =>
      (scopedProjects || []).filter((p: any) => p && p.Id);

    const dashVisibleProjects = (): any[] => {
      const scoped = dashScopeProjects();
      if (canReadAllProjects()) return scoped;
      const assigned = getAssignedProjects(scoped);
      return (assigned && assigned.length ? assigned : scoped);
    };

    const dashListEntityType = async (listName: string, fallback: string): Promise<string> => {
      try {
        const res = await safeFetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
          headers: { 'Accept': 'application/json;odata=verbose' },
          credentials: 'same-origin'
        });
        const data = await res.json();
        const type = data?.d?.ListItemEntityTypeFullName;
        if (!type) {
          dashLog(`⚠️ '${listName}': no ListItemEntityTypeFullName - using fallback ${fallback}`);
          return fallback;
        }
        return type;
      } catch (error) {
        console.error(`[DASH-SYNC] ⚠️ could not read entity type of '${listName}' - using fallback ${fallback}:`, error);
        return fallback;
      }
    };

    let dashPcMap: Map<number, string> | null = null;

    const dashProjectProfitCentres = async (): Promise<Map<number, string>> => {
      if (dashPcMap) return dashPcMap;

      const map = new Map<number, string>();
      try {
        const res = await safeFetch(
          `${siteUrl}/_api/web/lists/getbytitle('${DASH_SOURCES.projectList}')/items?$select=Id,PC&$top=5000`,
          { headers: { 'Accept': 'application/json;odata=verbose' }, credentials: 'same-origin' }
        );
        if (res.ok) {
          const rows = (await res.json())?.d?.results || [];
          rows.forEach((row: any) => {
            if (row && row.Id) map.set(Number(row.Id), String(row.PC || ''));
          });
        } else {
          console.warn(`[DASH-SYNC] ⚠️ Project List PC read HTTP ${res.status}`);
        }
      } catch (error) {
        console.warn('[DASH-SYNC] ⚠️ could not read the Project List PC column', error);
      }

      dashPcMap = map;
      return map;
    };

    const dashExistingProjectIds = async (target: DashSyncTarget, projects: any[]): Promise<Set<number> | null> => {
      try {
        const res = await safeFetch(
          `${siteUrl}/_api/web/lists/getbytitle('${target.listName}')/items?$select=Id,ProjectId,Title&$top=5000`,
          { headers: { 'Accept': 'application/json;odata=verbose' }, credentials: 'same-origin' }
        );
        if (res.ok) {
          const rows = (await res.json())?.d?.results || [];
          dashLog(`'${target.label}': ${rows.length} row(s) already on SharePoint`);
          return new Set(rows.map((row: any) => Number(row.ProjectId)).filter((id: number) => !!id));
        }
        dashLog(`⚠️ '${target.label}': bulk row read HTTP ${res.status} - falling back to per-project probes`);
      } catch (error) {
        dashLog(`⚠️ '${target.label}': bulk row read failed - falling back to per-project probes`, error);
      }

      const present = new Set<number>();
      let probeFailed = 0;

      for (const project of projects) {
        try {
          const res = await safeFetch(
            `${siteUrl}/_api/web/lists/getbytitle('${target.listName}')/items?$filter=ProjectId eq ${Number(project.Id)}&$select=Id&$top=1`,
            { headers: { 'Accept': 'application/json;odata=verbose' }, credentials: 'same-origin' }
          );
          if (res.ok) {
            const rows = (await res.json())?.d?.results || [];
            if (rows.length > 0) present.add(Number(project.Id));
          } else {
            probeFailed++;
          }
        } catch (probeError) {
          probeFailed++;
        }
      }

      dashLog(`'${target.label}': per-project probes found ${present.size} existing row(s)` +
        (probeFailed > 0 ? ` - ${probeFailed} probe(s) failed` : ''));

      return probeFailed > 0 && present.size === 0 ? null : present;
    };

    const dashForecastTitle = (project: any, fresh: any, suffix: string): string => {
      const currentYear = new Date().getFullYear();
      let totalValue = 0;

      Object.keys(fresh || {}).forEach(key => {
        if (/^\d{4}-\d{2}$/.test(key)) {
          const monthData = fresh[key];
          if (monthData && typeof monthData.actual === 'number') totalValue += monthData.actual;
        }
      });

      return `${project.Title} - ${currentYear} ${suffix} - ${totalValue.toLocaleString('en-US')}`;
    };

    const dashFieldNamesByList = new Map<string, Set<string> | null>();
    const dashOblFieldByList = new Map<string, string | null>();

    const dashPrimeFieldNames = async (listName: string): Promise<void> => {
      if (dashFieldNamesByList.has(listName)) return;

      try {
        const res = await safeFetch(
          `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$top=5000`,
          { headers: { 'Accept': 'application/json;odata=verbose' }, credentials: 'same-origin' }
        );
        if (!res.ok) {
          dashFieldNamesByList.set(listName, null);
          dashLog(`⚠️ '${listName}': field list HTTP ${res.status}`);
          return;
        }
        const rows = (await res.json())?.d?.results || [];
        const names = new Set<string>(rows.map((f: any) => String(f.InternalName || '')).filter(Boolean));
        dashFieldNamesByList.set(listName, names);
        dashOblFieldByList.set(
          listName,
          names.has('OBL') ? 'OBL' : (names.has('OBLData') ? 'OBLData' : null)
        );
        dashLog(`'${listName}': OBL column = ${dashOblFieldByList.get(listName) || '<none>'}`);
      } catch (error) {
        dashFieldNamesByList.set(listName, null);
        dashLog(`⚠️ '${listName}': field list could not be read`, error);
      }
    };

    const dashNormalizeOblField = (payload: any, listName: string): void => {
      if (!('OBL' in payload)) return;
      if (!dashFieldNamesByList.has(listName) || dashFieldNamesByList.get(listName) === null) return;
      if (dashOblFieldByList.get(listName) === 'OBLData') {
        payload.OBLData = payload.OBL;
        delete payload.OBL;
      }
    };

    const dashApplyRefField = (
      payload: any,
      attempt: number,
      refField: { exists: boolean; type: string },
      project: any
    ): void => {
      if (refField.exists && attempt === 1) {
        if (refField.type === 'Lookup') {
          payload.ProjectReferenceId = String(project.Id);
        } else {
          payload.ProjectReference = project.ProjectReference || '';
        }
      }
      if (attempt >= 2) {
        delete payload.ProjectReference;
        delete payload.ProjectReferenceId;
      }
    };

    const dashSyncMissingProjects = async (target: DashSyncTarget): Promise<number> => {
      const visibleProjects = dashVisibleProjects();
      if (visibleProjects.length === 0) {
        dashLog(`⚠️ '${target.label}': no visible project in scope - skipped`);
        return 0;
      }

      const existingIds = await dashExistingProjectIds(target, visibleProjects);
      if (!isCurrentRun()) return 0;
      if (existingIds === null) {
        dashLog(`⚠️ '${target.label}': existing rows could not be verified - skipped`);
        return 0;
      }

      const missing = visibleProjects.filter((p: any) => !existingIds.has(Number(p.Id)));
      if (missing.length === 0) {
        dashLog(`✅ '${target.label}' in sync (${visibleProjects.length} project(s))`);
        return 0;
      }

      let digest: string;
      try {
        digest = await getAuthDigest(siteUrl);
      } catch (digestError) {
        console.error(`[DASH-SYNC] ❌ '${target.label}': could not get digest:`, digestError);
        return 0;
      }
      if (!isCurrentRun()) return 0;

      const entityType = await dashListEntityType(target.listName, target.entityFallback);
      if (!isCurrentRun()) return 0;
      const pcMap = await dashProjectProfitCentres();
      if (!isCurrentRun()) return 0;
      const refField = await dashListProjectRefField(target.listName);
      if (!isCurrentRun()) return 0;
      await dashPrimeFieldNames(target.listName);
      if (!isCurrentRun()) return 0;

      dashLog(`➕ '${target.label}': adding ${missing.length} row(s): ${missing.map((p: any) => p.Title).join(', ')}`);

      let added = 0;
      let failed = 0;
      let permissionDenied = 0;

      for (const project of missing) {
        if (!isCurrentRun()) break;
        const projectWithPc = { ...project, PC: project.PC || pcMap.get(Number(project.Id)) || '' };
        let ok = false;

        for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
          try {
            const fresh = target.fresh();
            const payload = target.buildPayload(projectWithPc, fresh, entityType, attempt, refField);
            dashNormalizeOblField(payload, target.listName);

            const res = await safeFetch(`${siteUrl}/_api/web/lists/getbytitle('${target.listName}')/items`, {
              method: 'POST',
              headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'X-RequestDigest': digest
              },
              body: JSON.stringify(payload)
            });
            if (!isCurrentRun()) return added;

            if (res.ok) {
              ok = true;
              added++;
              dashLog(`✅ '${target.listName}': added ${projectWithPc.Title}` + (attempt > 1 ? ` (variant ${attempt})` : ''));
            } else {
              const bodyText = await res.text();
              if (res.status === 401 || res.status === 403) permissionDenied++;
              console.error(`[DASH-SYNC] ❌ '${target.listName}' create failed for ${projectWithPc.Title} ` +
                `(variant ${attempt}, HTTP ${res.status}): ${bodyText.substring(0, 300)}`);
            }
          } catch (projectError) {
            console.error(`[DASH-SYNC] ❌ '${target.listName}' create error for ${projectWithPc.Title} ` +
              `(variant ${attempt}):`, projectError);
          }

          if (!isCurrentRun()) break;
          if (!ok) await new Promise(resolve => setTimeout(resolve, 300));
        }

        if (!isCurrentRun()) break;
        if (!ok) failed++;
        await new Promise(resolve => setTimeout(resolve, 100));
      }

      if (added > 0) {
        dashLog(`➕ '${target.label}' → added ${added} missing row(s)` + (failed > 0 ? ` - ${failed} failed` : ''));
      } else if (failed > 0) {
        console.error(`[DASH-SYNC] ❌ '${target.label}' → 0 added, ${failed} failed` +
          (permissionDenied > 0 ? ` (${permissionDenied} permission-denied)` : ''));
      }

      return added;
    };

    const dashApplyAutoProgression = async (target: DashSyncTarget): Promise<number> => {
      const scopeIds = new Set<number>();
      dashVisibleProjects().forEach((p: any) => scopeIds.add(Number(p.Id)));
      if (scopeIds.size === 0) return 0;

      let rows: any[] = [];
      try {
        const res = await safeFetch(
          `${siteUrl}/_api/web/lists/getbytitle('${target.listName}')/items?$select=Id,ProjectId,${target.jsonField}&$top=5000`,
          { headers: { 'Accept': 'application/json;odata=verbose' } as any, credentials: 'same-origin' }
        );
        if (!isCurrentRun()) return 0;
        if (!res.ok) {
          dashLog(`⚠️ '${target.label}': progression read HTTP ${res.status} - skipped`);
          return 0;
        }
        rows = (await res.json())?.d?.results || [];
      } catch (error) {
        console.error(`[DASH-SYNC] ⚠️ '${target.label}' progression fetch failed:`, error);
        return 0;
      }
      if (!isCurrentRun()) return 0;

      const inScope = rows.filter((row: any) =>
        row.ProjectId !== undefined && row.ProjectId !== null && scopeIds.has(Number(row.ProjectId))
      );

      const blankRows: any[] = [];
      const staleRows: any[] = [];

      inScope.forEach((row: any) => {
        const raw = row[target.jsonField];
        if (!raw || raw === '""' || raw === 'null') {
          blankRows.push(row);
          return;
        }
        const data = safeJsonParse(raw, {});
        if (!data || typeof data !== 'object' || Object.keys(data).length === 0) {
          blankRows.push(row);
          return;
        }
        if (dashNeedsProgression(data, target.monthlyKind)) staleRows.push(row);
      });

      const upToDate = inScope.length - blankRows.length - staleRows.length;
      dashLog(`'${target.label}': ${inScope.length} row(s) in scope | blank: ${blankRows.length} | needs progression: ${staleRows.length} | up to date: ${upToDate}`);

      if (blankRows.length === 0 && staleRows.length === 0) {
        dashLog(`✅ '${target.label}': nothing to write`);
        return 0;
      }

      const work = [
        ...blankRows.map(row => ({ row, mode: 'fill' as const })),
        ...staleRows.map(row => ({ row, mode: 'roll' as const }))
      ];

      let digest: string;
      try {
        digest = await getAuthDigest(siteUrl);
      } catch (digestError) {
        console.error(`[DASH-SYNC] ❌ '${target.label}': could not get digest:`, digestError);
        return 0;
      }
      if (!isCurrentRun()) return 0;

      const entityType = await dashListEntityType(target.listName, target.entityFallback);
      if (!isCurrentRun()) return 0;
      dashLog(`🔄 '${target.label}': writing ${work.length} row(s)`);

      let progressed = 0;

      for (const item of work) {
        if (!isCurrentRun()) break;
        const row = item.row;
        try {
          const progressedData = item.mode === 'fill'
            ? target.fresh()
            : dashRollForward(safeJsonParse(row[target.jsonField], {}), target.monthlyKind);
          const payload: any = {
            __metadata: { type: entityType },
            UpdatedDate: new Date().toISOString()
          };
          payload[target.jsonField] = JSON.stringify(progressedData);

          const response = await safeFetch(`${siteUrl}/_api/web/lists/getbytitle('${target.listName}')/items(${row.Id})`, {
            method: 'POST',
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'X-RequestDigest': digest,
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE'
            },
            body: JSON.stringify(payload)
          });
          if (!isCurrentRun()) return progressed;

          if (response.ok) {
            progressed++;
          } else {
            const bodyText = await response.text();
            console.error(`[DASH-SYNC] ❌ '${target.listName}' progression failed for item ${row.Id} ` +
              `(HTTP ${response.status}): ${bodyText.substring(0, 300)}`);
          }

          await new Promise(resolve => setTimeout(resolve, 100));
        } catch (rowError) {
          console.error(`[DASH-SYNC] ⚠️ '${target.listName}' progression error for item ${row.Id}:`, rowError);
        }
      }

      if (progressed > 0) {
        dashLog(`🔄 '${target.label}' → wrote ${progressed} row(s)` +
          (blankRows.length > 0 ? ` (${blankRows.length} had blank ${target.jsonField})` : ''));
      }
      return progressed;
    };

    const DASH_AUTO_SYNC_TARGETS: DashSyncTarget[] = [
      {
        key: 'revenue',
        label: 'Revenue Forecast List',
        listName: DASH_SOURCES.revenueList,
        entityFallback: 'SP.Data.RevenueForecastListItem',
        jsonField: 'ForecastData',
        monthlyKind: 'revenue',
        fresh: () => dashFreshData('revenue', false),
        buildPayload: (project, fresh, entityType, attempt, refField) => {
          const payload: any = {
            __metadata: { type: entityType },
            Title: dashForecastTitle(project, fresh, 'Forecast'),
            ProjectId: project.Id.toString(),
            ProfitCentre: project.PC || '',
            ForecastData: JSON.stringify(fresh),
            RevenueActual: '{}',
            Remarks: '',
            OBL: String(project.PC ? 0 : 0),
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
          };
          dashApplyRefField(payload, attempt, refField, project);
          if (attempt >= 3) delete payload.OBL;
          return payload;
        }
      },
      {
        key: 'collection',
        label: 'Collection List',
        listName: DASH_SOURCES.collectionList,
        entityFallback: 'SP.Data.CollectionListItem',
        jsonField: 'CollectionData',
        monthlyKind: 'collection',
        fresh: () => dashFreshData('collection', false),
        buildPayload: (project, fresh, entityType, attempt, refField) => {
          const payload: any = {
            __metadata: { type: entityType },
            Title: dashForecastTitle(project, fresh, 'Collection'),
            ProjectId: project.Id.toString(),
            ProfitCentre: project.PC || '',
            CollectionData: JSON.stringify(fresh),
            Remarks: '',
            OBL: project.PC ? '0' : '0',
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
          };
          dashApplyRefField(payload, attempt, refField, project);
          if (attempt >= 3) delete payload.OBL;
          return payload;
        }
      },
      {
        key: 'seibies',
        label: 'SEIBIES',
        listName: 'SEIBIES',
        entityFallback: 'SP.Data.SEIBIESListItem',
        jsonField: 'ForecastData',
        monthlyKind: 'standard',
        fresh: () => dashFreshData('standard', false),
        buildPayload: (project, fresh, entityType, attempt, refField) => {
          const payload: any = {
            __metadata: { type: entityType },
            Title: dashForecastTitle(project, fresh, 'SEIBIES Forecast'),
            ProjectId: project.Id,
            ProfitCentre: project.PC || '',
            ForecastData: JSON.stringify(fresh),
            OBL: project.PC ? 0 : 0,
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
          };
          dashApplyRefField(payload, attempt, refField, project);
          if (attempt >= 3) delete payload.OBL;
          return payload;
        }
      },
            {
        key: 'inventory',
        label: 'Inventory',
        listName: DASH_SOURCES.inventoryList,
        entityFallback: 'SP.Data.InventoryListItem',
        jsonField: 'ForecastData',                          // ← note column (JSON store)
        monthlyKind: 'standard',
        fresh: () => dashFreshData('standard', false),
        buildPayload: (project, fresh, entityType, attempt, refField) => {
          const payload: any = {
            __metadata: { type: entityType },
            Title: dashForecastTitle(project, fresh, 'Inventory'),
            ProjectId: project.Id,
            ProfitCentre: project.PC || '',
            ForecastData: JSON.stringify(fresh),            // ← JSON goes here
            OBL: project.PC ? 0 : 0,
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
          };
          dashApplyRefField(payload, attempt, refField, project);
          if (attempt >= 3) delete payload.OBL;
          return payload;
        }
      },
      {
        key: 'invoicing',
        label: 'InvoicingPlanning',
        listName: 'InvoicingPlanning',
        entityFallback: 'SP.Data.InvoicingPlanningListItem',
        jsonField: 'ForecastData',
        monthlyKind: 'standard',
        fresh: () => dashFreshData('standard', true),
        buildPayload: (project, fresh, entityType, attempt, refField) => {
          const payload: any = {
            __metadata: { type: entityType },
            Title: dashForecastTitle(project, fresh, 'Forecast'),
            ProjectId: project.Id,
            ProfitCentre: project.PC || '',
            ForecastData: JSON.stringify(fresh),
            OBL: project.PC ? 0 : 0,
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
          };
          dashApplyRefField(payload, attempt, refField, project);
          if (attempt >= 3) delete payload.OBL;
          return payload;
        }
      }
    ];

    let dashAutoSyncRunning = false;
    const DASH_SYNC_AUTO_LOCK_MS = 100000;

    const runDashboardBackgroundSync = async (trigger: string = 'auto'): Promise<void> => {
      const w = window as any;

      if (!isCurrentRun()) return;
      if (dashAutoSyncRunning || w.__dashSyncRunning) {
        dashLog('already running - second call ignored');
        return;
      }
      if (trigger === 'auto' && w.__dashSyncLockUntil && Date.now() < w.__dashSyncLockUntil) {
        dashLog('skipped - another dashboard instance synced just now');
        return;
      }

      dashAutoSyncRunning = true;
      w.__dashSyncRunning = true;
      w.__dashSyncLockUntil = Date.now() + DASH_SYNC_AUTO_LOCK_MS;

      const summary: string[] = [];
      let changed = 0;

      dashLog(`started (${trigger}) →`, DASH_AUTO_SYNC_TARGETS.map(t => t.label).join(' · '));
      dashLog('permissions:', dashPermissionReport());
      setDashSyncStatus('⏳ Syncing…', '#ffc107');

      try {
        await dashEnsureSiteLevelAuth();
        if (!isCurrentRun()) return;
        dashLog('permissions after re-init:', dashPermissionReport());

        if (!canRead()) {
          dashLog('skipped - user has no read access');
          setDashSyncStatus('Sync skipped (no read access)', '#ffc107');
          return;
        }

        if (!dashCanWrite()) {
          dashLog('skipped - no create rights (Admin / GH Head / PM Manager only)');
          setDashSyncStatus('Sync skipped (read-only account)', '#ffc107');
          return;
        }

        dashLog(`project scope: ${dashVisibleProjects().length} project(s)`);

        for (const target of DASH_AUTO_SYNC_TARGETS) {
          if (!isCurrentRun()) break;
          try {
            const added = await dashSyncMissingProjects(target);
            if (!isCurrentRun()) break;
            const progressed = await dashApplyAutoProgression(target);
            if (!isCurrentRun()) break;
            changed += added + progressed;
            summary.push(`${target.label}: +${added} new, ${progressed} written`);
            setDashSyncStatus(`⏳ Syncing… ${summary.length}/${DASH_AUTO_SYNC_TARGETS.length}`, '#ffc107');
          } catch (targetError) {
            console.error(`[DASH-SYNC] ⚠️ '${target.label}' failed:`, targetError);
            summary.push(`${target.label}: error`);
          }
        }
      } finally {
        dashAutoSyncRunning = false;
        (window as any).__dashSyncRunning = false;
        if (isCurrentRun()) {
          dashLog(`finished (${trigger}) →`, summary.join(' | '));

          if (changed > 0) {
            setDashSyncStatus(`✅ Synced ${changed} change(s)`, '#28a745');
            try {
              await refreshBoard();
            } catch (refreshError) {
              console.error('[DASH-SYNC] ⚠️ refresh after background sync failed:', refreshError);
            }
          } else {
            setDashSyncStatus('✅ Already in sync', '#28a745');
          }

          window.setTimeout(() => {
            if (isCurrentRun()) setDashSyncStatus('');
          }, 15000);
        }
      }
    };

    const startDashboardBackgroundSync = (): void => {
      (window as any).__dashSyncLockUntil = 0;
      if (!canRead()) {
        dashLog('skipped - user has no read access');
        setDashSyncStatus('Sync skipped (no read access)', '#ffc107');
        return;
      }
      if (!dashCanWrite()) {
        dashLog('skipped - no create rights (Admin / GH Head / PM Manager only)');
        setDashSyncStatus('Sync skipped (read-only account)', '#ffc107');
        return;
      }

      backgroundSyncTimer = window.setTimeout(() => {
        backgroundSyncTimer = null;
        if (isCurrentRun()) void runDashboardBackgroundSync('auto');
      }, 1500);
    };

    const mountDashboardSyncButton = (): void => {
      const meta = document.querySelector('.header-meta') || document.querySelector('.header');
      if (!meta || document.getElementById('dsBgSyncNow')) return;

      const btn = document.createElement('button');
      btn.id = 'dsBgSyncNow';
      btn.type = 'button';
      btn.textContent = '🔄 Sync now';
      btn.title = 'Run the same background sync + monthly/yearly progression the Inventory / SEIBIES / Revenue / Collection / Invoicing pages run on open';
      btn.style.cssText =
        'margin-left:10px;padding:3px 10px;font-size:11px;border-radius:4px;border:1px solid #6ea8fe;' +
        'background:rgba(255,255,255,0.12);color:#fff;cursor:pointer;';
      btn.addEventListener('click', () => { void runDashboardBackgroundSync('button'); });
      meta.appendChild(btn);
    };

    (window as any).dashboardSyncNow = runDashboardBackgroundSync;

    mountDashboardSyncButton();
    startDashboardBackgroundSync();

    dashboardCleanup = (): void => {
      stopAutoRefresh();
      if (backgroundSyncTimer !== null) {
        window.clearTimeout(backgroundSyncTimer);
        backgroundSyncTimer = null;
      }
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      document.removeEventListener('dt-page-deactivated', handlePageDeactivated);
      window.removeEventListener('beforeunload', handleBeforeUnload);
      document.getElementById('dsBgSyncNow')?.remove();
      if ((window as any).dashboardSyncNow === runDashboardBackgroundSync) {
        delete (window as any).dashboardSyncNow;
      }
    };
  }, 250);
};

export default {
  content: dashboardContent,
  style: '',                 // CSS moved to ../styles/dashboard.css
  script: dashboardScript
};
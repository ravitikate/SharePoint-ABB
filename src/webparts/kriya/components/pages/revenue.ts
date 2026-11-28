declare const _spPageContextInfo: any;
// import Swal from "sweetalert2";

import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/revenue.css";
import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css";
import { createSharePointItemsInBatches, setBackgroundAutoUpdateIndicator } from "./sharePointBatch";

import { 
  authManager, 
  hasFullControl,
  canContribute, 
  canRead,
  canEditOwn,
  canReadAllProjects,
  canFreeze,
  isGHHead,
  getUserGroups,
  getCurrentUserFullName,
  getPermissionLevel,
  canEditProject as authCanEditProject,
  canCreateRecord,
  getAssignedProjects,
  updateUIBasedOnPermissions as authUpdateUI
} from "./auth";
     
interface ProjectItem {
  Id: number;
  Title: string;
  CustomerName?: string;
  ProjectManagerId?: number;
  ProjectManager?: string | { Title: string; Email?: string };
  PC?: string;
  ProjectReference?: string;
  GroupHead1?: string | { Title: string; Email?: string };
  GroupHead1Id?: number;

  // ✅ Engineering role fields (used by ROLE_OWN_FIELDS in auth.ts)
  EnggLead?: string | { Title: string; Email?: string };
  EnggLeadId?: number;
  EnggSubLead1?: string | { Title: string; Email?: string };
  EnggSubLead1Id?: number;
  EnggSubLead2?: string | { Title: string; Email?: string };
  EnggSubLead2Id?: number;
  EngineeringManagers?: string | { Title: string; Email?: string };
  EngineeringManagersId?: number;
}       

interface SortConfig {
  column: string;
  direction: 'asc' | 'desc';
  subColumn?: string;
  monthIndex?: number;
}

let currentSortConfig: SortConfig = {
  column: '',
  direction: 'asc'
};

interface RevenueForecastItem {
  Id: number;
  Title: string;
  __metadata?: { etag?: string };
  Project?: { 
    Id: number;
    Title: string;
  };
  ProfitCentre?: string;
  OBL?: number;
  ForecastData?: string;
  RevenueActual?: string;
  Remarks?: string | null;
  CreatedDate?: string;
  UpdatedDate?: string;
  Freeze?: boolean;
  [key: string]: any;
}

interface SharePointResponse<T> {
  d: {
    results: T[];
  };
}

interface ForecastItem {
    name: string;
    plan: number;
    actual: number;
    forecast?: number;
    date?: string;
    previous?: number;
}

interface ForecastMonthData {
    plan: number;
    actual: number;
    forecast?: number;
    date?: string;

    // Revenue actual belongs to this specific month
    monthActual?: number;

    items?: {
        [itemNumber: string]: ForecastItem;
    };
}

interface YearlyForecastData {
    [year: string]: {
        plan: number;
        actual: number;
        forecast?: number;
    };
}

interface ForecastData {
    [monthKey: string]: ForecastMonthData | YearlyForecastData | undefined;
    yearly?: YearlyForecastData;
}

export default {
  content: `
<div class="revenue-page fade-in">

  <!-- ============================================ -->
  <!-- LOADER - covers ONLY this section (the page   -->
  <!-- wrapper), so the sidebar / top nav stay       -->
  <!-- visible and usable. Inline styles so it shows -->
  <!-- before CSS loads; hidden by hideRevenueLoader -->
  <!-- () in the initialize() finally block.         -->
  <!-- ============================================ -->
  <div id="revenuePageLoader" class="revenue-loader" style="position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;background:rgba(248,249,250,0.92);z-index:999;">
    <div class="revenue-loader-box">
      <div class="revenue-loader-spinner"></div>
      <div class="revenue-loader-text">Loading Revenue Data</div>
      <div class="revenue-loader-subtext">Please wait... fetching data</div>
      <div class="revenue-loader-bar"><span></span></div>
    </div>
  </div>

  <!-- ============================================ -->
  <!-- TABLE SECTION - toggled by showForm()/       -->
  <!-- hideForm(): while the Add/Edit form is open  -->
  <!-- the whole list (header, filters, table,      -->
  <!-- pagination) is hidden, so ONLY ONE of        -->
  <!-- form / project list is visible at a time.    -->
  <!-- (The closing </section> after the pagination -->
  <!-- belongs to this opening tag.)                -->
  <!-- ============================================ -->
  <section class="table-section">

    <div class="page-header">
      <h1>Revenue plan 5QRF</h1>
       
    </div>
    
    <!-- ✅ TOOLBAR ABOVE THE TABLE (inventory / procurement style)
         LEFT  = Filters (funnel icon + active count badge) + the compact
                 "Apply" pill that appears when the popup has unapplied changes
         RIGHT = Upload Actual / Download data (summary) / Download data (with
                 item details) / Download Data Historical
         The page header (title) is untouched - the toolbar never sits next to it. -->
    <div class="revenue-table-toolbar">
      <div class="revenue-toolbar-left">
        <button type="button" class="revenue-btn-secondary revenue-filter-trigger" id="revenueFilterTrigger" aria-haspopup="dialog" aria-expanded="false" title="Filter Revenue forecast data">
          <svg class="revenue-filter-funnel" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8v5l-4 2v-7L3 5z"/></svg><span>Filters</span><span class="revenue-filter-count" id="revenueFilterCount">0</span><svg class="revenue-filter-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>
        </button>
        <!-- appears only while the popup has unapplied changes -->
        <button type="button" class="revenue-btn-primary revenue-filter-applypill" id="revenueFilterApplyPill" title="Apply the selected filters">Apply</button>
      </div>
      <div class="revenue-table-toolbar-actions">
        <button class="revenue-btn-secondary revenue-import-actual-btn">📥 Upload Actual</button>
        <button class="revenue-btn-secondary revenue-export-btn">📊 Download data (summary)</button>
        <button class="revenue-btn-secondary revenue-export-details-btn">📋 Download data (with item details)</button>
        <button class="revenue-btn-secondary revenue-export-historical-btn">📈 Download Data Historical</button>
      </div>
    </div>
    <!-- applied filters show up here as removable chips -->
     <!-- <div class="revenue-filter-summary" id="revenueFilterChips" hidden></div> -->

    <div class="revenue-table-actions revenue-filter-overlay" id="revenueFilterOverlay" role="dialog" aria-modal="true" aria-labelledby="revenueFilterTitle">
      <div class="revenue-filter-modal">
        <div class="revenue-filter-modal-head"><h2 id="revenueFilterTitle">Filters</h2><button type="button" id="revenueFilterClose" aria-label="Close">&times;</button></div>
        <div class="revenue-table-actions-row">
        <div class="revenue-table-col-p4">
          <label for="filterProject">Project</label>
          <select id="filterProject" class="revenue-filter-select"><option value="">All Projects</option></select>
        </div>
        <div class="revenue-table-col-p4">
          <label for="filterCustomer">Customer</label>
          <select id="filterCustomer" class="revenue-filter-select"><option value="">All Customers</option></select>
        </div>
        <div class="revenue-table-col-p4">
          <label for="filterProjectManager">Project Manager</label>
          <select id="filterProjectManager" class="revenue-filter-select"><option value="">All Managers</option></select>
        </div>
        <div class="revenue-table-col-p4">
          <label for="filterGroupHead">Group Head</label>
          <select id="filterGroupHead" class="revenue-filter-select"><option value="">All Group Heads</option></select>
        </div>
        <div class="revenue-table-col-p4">
          <label for="filterProjectReference">Project Reference</label>
          <select id="filterProjectReference" class="revenue-filter-select"><option value="">All References</option></select>
        </div>
        <div class="revenue-table-col-p4">
          <label> </label>
          <button type="button" class="revenue-btn-secondary revenue-clear-filters-btn" id="revenueResetFilterBtn" title="Clear All Filters" disabled>Reset Filter</button>
        </div>
      </div>

      <div class="revenue-table-actions-row">
        <div class="revenue-table-col-p4">
          <label for="filterYear">Year</label>
          <select id="filterYear" class="revenue-filter-select"><option value="">All Years</option></select>
        </div>
        <div class="revenue-table-col-p4">
          <label for="filterMonth">Month</label>
          <select id="filterMonth" class="revenue-filter-select"><option value="">All Months</option></select>
        </div>
        <div class="revenue-table-col-p4 revenue-filter-dupe-actions">
          <label> </label>
          <button class="revenue-btn-secondary revenue-import-actual-btn">📥 Upload Actual</button>
        </div>
        <div class="revenue-table-col-p4 revenue-filter-dupe-actions">
          <label> </label>
          <button class="revenue-btn-secondary revenue-export-btn">📊 Download data (summary)</button>
        </div>
        <div class="revenue-table-col-p4 revenue-filter-dupe-actions">
          <label> </label>
          <button class="revenue-btn-secondary revenue-export-details-btn">📋 Download data (with item details)</button>
        </div>
        <div class="revenue-table-col-p4 revenue-filter-dupe-actions">
          <label> </label>
          <button class="revenue-btn-secondary revenue-export-historical-btn">📈 Download Data Historical</button>
        </div>
      </div>
        <div class="revenue-filter-modal-foot">
          <button type="button" class="revenue-btn-secondary" id="revenueFilterCancel">Cancel</button>
          <button type="button" class="revenue-btn-secondary revenue-filter-btn-clear" id="revenueFilterClear" title="Clear all filters" disabled>Clear</button>
          <button type="button" class="revenue-btn-primary revenue-filter-btn-apply" id="revenueFilterApply" disabled>Apply</button>
        </div>
      </div>
    </div>
    <div class="revenue-table-container-rev">
      <div class="revenue-table-scroll-wrapper">
        <table class="revenue-month-columns-table-rev"> 
           <thead id="totalsHeader"></thead> 
           <thead id="monthColumnsHeader"></thead> 
        <tbody class="revenue-table-body-rev"></tbody>
      </table>
      </div> 
      
      
    </div>
    <!-- ✅ PAGINATION - one flat bar, same text + layout as the mock-up
         LEFT  : Items per page: [n]   Showing 1 - 25 of 60 items
         RIGHT : « ‹ [tiles] of 6 › » -->
    <div class="revenue-pagination-container">
      <div class="revenue-pagination-left">
        <div class="revenue-page-size-selector revenue-pg-size">
          <label for="pageSizeSelect">Items per page:</label>
          <select id="pageSizeSelect" class="revenue-page-size-select">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>
        <div class="revenue-pagination-info revenue-pg-total">
          Showing <span class="revenue-pg-range"><span id="startItem">0</span> - <span id="endItem">0</span></span> of <span class="revenue-pg-count" id="totalItems">0</span> items
        </div>
      </div>
      <div class="revenue-pagination-right">
        <button class="revenue-pagination-btn revenue-pg-icon" id="firstPage" disabled title="First page">«</button>
        <button class="revenue-pagination-btn revenue-pg-icon" id="prevPage" disabled title="Previous page">‹</button>
        <div class="revenue-pagination-pages" id="pageNumbers"></div>
        <span class="revenue-pg-of">of <span id="totalPages">1</span></span>
        <button class="revenue-pagination-btn revenue-pg-icon" id="nextPage" disabled title="Next page">›</button>
        <button class="revenue-pagination-btn revenue-pg-icon" id="lastPage" disabled title="Last page">»</button>
      </div>
    </div>
  </section>
 <section class="form-section revenue-card" style="display:none;">
  <form id="revenueForm" novalidate>

    <div class="page-header">
      <h1 id="formTitle"></h1>
      <div class="revenue-form-header-actions">
        <button type="submit"
                class="revenue-btn-primary revenue-save-btn revenue-save-top-btn"
                title="Save Revenue Plan">💾 Save Revenue Plan</button>
        <button type="button" class="revenue-btn-secondary revenue-back-top-btn"
                title="Go back without saving">⬅ Back</button>
      </div>
    </div>
    
    <!-- Project Selection -->
    <div class="revenue-form-row revenue-hidden-section">
      <div class="revenue-form-group revenue-project-select-container">
        <label>Project</label>
        <select id="projectSelect" required></select>
      </div>
    </div> 

    <!-- MAIN WRAPPER - 2 Column Layout -->
    <div class="revenue-project-main-wrapper">

      <!-- LEFT SIDE - Project Details (6 Rows) -->
<div class="revenue-project-left-side" id="projectDetails">
  <div class="revenue-detail-item">
    <span class="revenue-detail-label">Project:</span>
    <span class="revenue-detail-value" id="projectNameDisplay">-</span>
    <!-- ✅ Extra SAVE button at the TOP of the edit section, right side of
         the Project name row. Same type="submit" + .revenue-save-btn class
         as the bottom button, so both trigger the same submit handler and
         stay in sync (visibility, frozen text/style, saving state). -->
    
  </div>
  <div class="revenue-detail-item">
    <span class="revenue-detail-label">Project Ref:</span>
    <span class="revenue-detail-value" id="projectRef">-</span>
  </div>
  <div class="revenue-detail-item">
    <span class="revenue-detail-label">Profit Center:</span>
    <span class="revenue-detail-value" id="profitCenter">-</span>
  </div>
  <div class="revenue-detail-item">
    <span class="revenue-detail-label">Project Manager:</span>
    <span class="revenue-detail-value" id="projectManager">-</span>
  </div>
  <div class="revenue-detail-item">
    <span class="revenue-detail-label">Group Head:</span>
    <span class="revenue-detail-value" id="groupHead">-</span>
  </div>
  <div class="revenue-detail-item">
    <span class="revenue-detail-label">Customer:</span>
    <span class="revenue-detail-value" id="customer">-</span>
  </div>
</div>

      <!-- RIGHT SIDE - Remarks & OBL -->
      <div class="revenue-project-right-side">
        <!-- Remarks Internal -->
        <div class="revenue-remarks-group">
          <label for="remarksInternal">REMARKS (INTERNAL)</label>
          <textarea id="remarksInternal" class="revenue-remarks-textarea" rows="3" placeholder="Enter internal remarks..."></textarea>
        </div>
        
        <!-- Remarks Sprint Cell  -->
        <div class="revenue-remarks-group">
          <label for="remarksSprintCell">REMARKS (SPRINT CELL)</label>
          <textarea id="remarksSprintCell" class="revenue-remarks-textarea" rows="3" placeholder="Enter sprint cell remarks..."></textarea>
        </div>
        
        <!-- OBL Section -->
        <div class="revenue-obl-section-container revenue-hidden-section" id="oblContainer">
          <div class="revenue-obl-content">
            <div class="revenue-obl-left">
              <div class="revenue-obl-label-group">
                <label>OBL</label> 
              </div>
              <div class="revenue-obl-value-display">
                <input type="text" id="oblLastMonth" class="revenue-obl-input revenue-number-input" placeholder="Enter OBL" readonly />                
              </div>
            </div>
            <div class="revenue-obl-right">
              <div class="revenue-validation-summary">
                <div class="revenue-validation-item">
                  <span class="revenue-validation-label">Total Current Value:</span>
                  <span class="revenue-validation-value" id="totalCurrentValue">0</span>
                </div>
                <div class="revenue-validation-item">
                  <span class="revenue-validation-label">Remaining Balance:</span>
                  <span class="revenue-validation-value" id="remainingValue">0</span>
                </div>
                <div class="revenue-validation-line">
                  <div class="revenue-line-fill" id="oblLineFill"></div>
                </div>
              </div>
            </div>
          </div>
          <div class="revenue-validation-message" id="oblValidationMessage">
            ✅ All values are within OBL limit
          </div> 
        </div>
      </div> 

    </div>

    <!-- Forecast Sections -->
<div class="revenue-forecast-section revenue-hidden-section" id="historicalForecastSection">
  <div class="revenue-historical-toggle-header">
  <h3>📊 Historical Data (Last 3 Months - Read only)</h3>

  <button
    type="button"
    id="revenue-toggle-historical-btn"
    aria-controls="historicalContainer"
    aria-expanded="false"
    aria-label="Show historical data"
    title="Show historical data"
  >
    <span class="revenue-toggle-arrow" aria-hidden="true">▼</span>
    <span class="revenue-toggle-label">Show</span>
  </button>
</div>
  <div id="historicalContainer" class="revenue-historical-container"></div>
</div>
    
    
    <div class="revenue-forecast-section revenue-hidden-section" id="detailedForecastSection">
      <h3>Detailed Forecast (Current + Next 2 Months)</h3>
      <div id="firstRowContainer" class="revenue-first-row-container"></div>
    </div>
    
    <div class="revenue-forecast-section revenue-hidden-section" id="longtermForecastSection">
      <h3>Long-term Forecast (Next 15 Months)</h3>
      <div id="remainingContainer" class="revenue-remaining-container"></div>
    </div>

    <div class="revenue-forecast-section revenue-hidden-section" id="yearlyForecastSection">
      <h3>5-Year Forecast</h3>
      <div id="yearlyContainer" class="revenue-yearly-container-rev"></div>
    </div>
    
    <div class="form-actions">
      <button type="submit" class="revenue-btn-primary revenue-save-btn">💾 Save Revenue Plan</button>
      <button type="button" class="revenue-btn-secondary revenue-back-to-table">⬅ Back</button>
    </div>
  </form>
</section>
    
    
</div>
  `,

script: () => {
  if ((window as any).__dtActivePageKey !== 'revenue') return;
  
  // ============================================
  // LOAD SWEETALERT2 WITH GUARANTEED FALLBACK
  // ============================================
  const loadSwal = (): Promise<any> => new Promise((resolve) => {
    // Already loaded?
    if ((window as any).Swal && typeof (window as any).Swal.fire === 'function') {
      return resolve((window as any).Swal);
    }

    const siteUrl = typeof _spPageContextInfo !== 'undefined'
      ? _spPageContextInfo.webAbsoluteUrl
      : window.location.origin + '/sites/Kriya';

    const tryLoad = (url: string): Promise<any> => new Promise((res) => {
      const s = document.createElement('script');
      s.src = url;
      s.onload = () => {
        console.log(`✅ Swal loaded from ${url}`);
        res((window as any).Swal);
      };
      s.onerror = () => {
        console.warn(`❌ Failed to load Swal from ${url}`);
        res(undefined);
      };
      document.head.appendChild(s);
    });

    (async () => {
      // 1. Try local SiteAssets
      let swal = await tryLoad(`${siteUrl}/SiteAssets/sweetalert2.all.min.js`);

      // 2. Try CDN
      if (!swal || typeof swal.fire !== 'function') {
        swal = await tryLoad('https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js');
      }

      // 3. Native fallback shim — GUARANTEES Swal.fire exists
      if (!swal || typeof swal.fire !== 'function') {
        console.error('❌ SweetAlert2 unavailable everywhere. Using native fallback shim.');
        swal = {
          fire: async (opts: any) => {
            const raw = opts?.html || opts?.text || '';
            const tmp = document.createElement('DIV');
            tmp.innerHTML = raw;
            const txt = tmp.textContent || tmp.innerText || '';

            if (opts?.showDenyButton) {
              // 3-way approx: OK = Yes/Save, Cancel = No/Don't Save
              const yes = window.confirm(
                `${opts.title || ''}\n\n${txt}\n\nClick OK to save, Cancel to discard.`
              );
              return { isConfirmed: yes, isDenied: !yes, isDismissed: false };
            }

            if (opts?.showCancelButton) {
              const ok = window.confirm(`${opts.title || ''}\n\n${txt}`);
              return { isConfirmed: ok, isDenied: false, isDismissed: !ok };
            }

            window.alert(opts?.title || txt);
            return { isConfirmed: true, isDenied: false, isDismissed: false };
          },
          close: () => {},
          isVisible: () => false
        };
      }

      // Expose globally
      (window as any).Swal = swal;
      resolve(swal);
    })();
  });

  // Kick off the load, then start the page
  loadSwal().then(() => {
    setTimeout(async () => {
      if ((window as any).__dtActivePageKey !== 'revenue') return;
      // ============================================
      // ⚠️ CRITICAL: rebind local Swal from window
      // ============================================
      //  const Swal = (window as any).Swal;

      // ... everything else that was already inside your setTimeout stays the same ...
      // ============================================
      // SECTION LOADER (not fullscreen)
      // 'Loading Revenue Data' covers only .revenue-page, so the sidebar
      // stays visible. Hidden by hideRevenueLoader() in the initialize()
      // finally block; this watchdog is only a safety net for a hung
      // request and never hides the loader while work is still in flight.
      // ============================================



      (window as any).__revenueLoaderDeadline = Date.now() + 30000;

      window.setTimeout(() => {
        const loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;
        if (loader && Date.now() > Number((window as any).__revenueLoaderDeadline || 0)) {
          console.warn('Loader watchdog: hiding loader after timeout');
          // hideRevenueLoader();
        }
      }, 31000);

      const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/Kriya";
      const listName = "Revenue Forecast List";
      const projectListName = "Project List";

      
 

// ============================================
// LOCAL PERMISSION WRAPPERS
// ============================================

// Re-export auth functions with consistent naming for local use
const canEditProject = (project?: ProjectItem, forecast?: RevenueForecastItem): boolean => authCanEditProject(project);
const canCreateForecast = (): boolean => canCreateRecord();
// const getAssignedProjectsLocal = (allProjects: ProjectItem[]): ProjectItem[] => getAssignedProjects(allProjects);
const canImportData = (): boolean => canContribute();
const canExportData = (): boolean => canRead();

// ============================================
// UI PERMISSION UPDATE
// ============================================

const updateUIBasedOnPermissions = (): void => {
  // 1. Run the standard auth UI updates
  authUpdateUI({
    importBtn: '.revenue-import-actual-btn',
    exportBtn: '.revenue-export-btn',
    addBtn: '.show-form-btn',
    saveBtn: '.revenue-btn-primary[type="submit"]',
    freezeCheckboxes: '.revenue-freeze-checkbox'
  });
  
  // 2. 🔐 FORCE ADMIN-ONLY for "Upload Actual"
  // We override the standard canContribute() check and strictly enforce Admin status.
  const isAdmin = hasFullControl(); 
  const uploadActualBtns = document.querySelectorAll<HTMLElement>('.revenue-import-actual-btn');
  
  uploadActualBtns.forEach(btn => {
    if (isAdmin) {
      btn.style.display = '';        // Show to Admin
      (btn as HTMLButtonElement).disabled = false;
    } else {
      btn.style.display = 'none';    // Hide from everyone else
      (btn as HTMLButtonElement).disabled = true;
    }
  });

  // 3. Export buttons (visible to every reader)
  const exportDetailsBtn = document.querySelector('.revenue-export-details-btn') as HTMLElement;
  if (exportDetailsBtn) {
    exportDetailsBtn.style.display = canExportData() ? '' : 'none';
  }

  const exportHistoricalBtn = document.querySelector('.revenue-export-historical-btn') as HTMLElement;
  if (exportHistoricalBtn) {
    exportHistoricalBtn.style.display = canExportData() ? '' : 'none';
  }
};

// ============================================
// PERMISSION CHECK FUNCTIONS
// ============================================

const canViewForm = (isEdit: boolean = false, project?: ProjectItem, forecast?: RevenueForecastItem): boolean => {
  if (!canRead()) return false;
  
  if (isEdit) {
    // 🔐 RBAC: opening an EXISTING forecast is allowed for any role that may
    // VIEW that row - READ_ALL roles (Admin, SC Team, Controlling) see all
    // projects; every other role only rows of projects assigned to them.
    // Actual EDIT rights are enforced separately (read-only form + hidden
    // Save button via canEditProject).
    if (canEditProject(project, forecast)) return true;
    if (canReadAllProjects()) return true;
    if (project && getAssignedProjects(projects).some(p => p.Id === project.Id)) return true;
    return false;
  } else {
    // 🔐 RBAC: creating a NEW forecast requires create rights
    // (FULL_CONTROL / CONTRIBUTE only - Admin, GH Heads, PM Managers).
    return canCreateForecast();
  }
};

      // DOM Elements
      const tableBody = document.querySelector(".revenue-table-body-rev") as HTMLElement;
      const monthColumnsHeader = document.querySelector("#monthColumnsHeader") as HTMLElement;
      const formSection = document.querySelector(".form-section") as HTMLElement;
      const tableSection = document.querySelector(".table-section") as HTMLElement; 
      const projectSelect = document.querySelector("#projectSelect") as HTMLSelectElement;
      const historicalContainer = document.querySelector("#historicalContainer") as HTMLElement;
      const firstRowContainer = document.querySelector("#firstRowContainer") as HTMLElement;
      const remainingContainer = document.querySelector("#remainingContainer") as HTMLElement;
      const yearlyContainer = document.querySelector("#yearlyContainer") as HTMLElement;
      // const projectDetails = document.querySelector("#projectDetails") as HTMLElement;
      const backBtn = document.querySelector(".revenue-back-to-table") as HTMLButtonElement;
      const form = document.querySelector("#revenueForm") as HTMLFormElement;
      const exportBtn = document.querySelector(".revenue-export-btn") as HTMLButtonElement;
      const oblLastMonthInput = document.querySelector("#oblLastMonth") as HTMLInputElement;
      const totalCurrentValueSpan = document.querySelector("#totalCurrentValue") as HTMLSpanElement;
      const remainingValueSpan = document.querySelector("#remainingValue") as HTMLSpanElement;
      const oblLineFill = document.querySelector("#oblLineFill") as HTMLDivElement;
      const oblValidationMessage = document.querySelector("#oblValidationMessage") as HTMLDivElement;
      const formTitle = document.querySelector("#formTitle") as HTMLElement;

      // Filter Elements
      const filterYear = document.querySelector("#filterYear") as HTMLSelectElement;
      const filterMonth = document.querySelector("#filterMonth") as HTMLSelectElement;
      const filterProject = document.querySelector("#filterProject") as HTMLSelectElement;
      const filterCustomer = document.querySelector("#filterCustomer") as HTMLSelectElement;
      const filterProjectManager = document.querySelector("#filterProjectManager") as HTMLSelectElement;
      const filterGroupHead = document.querySelector("#filterGroupHead") as HTMLSelectElement;
      const filterProjectReference = document.querySelector("#filterProjectReference") as HTMLSelectElement;

      // Action Buttons
      const importActualBtns = document.querySelector(".revenue-import-actual-btn") as HTMLButtonElement;
      const clearFiltersBtn = document.querySelector(".revenue-clear-filters-btn") as HTMLButtonElement;
      
      const exportDetailsBtn = document.querySelector(".revenue-export-details-btn") as HTMLButtonElement;

      // Pagination Elements
      const firstPageBtn = document.querySelector("#firstPage") as HTMLButtonElement;
      const prevPageBtn = document.querySelector("#prevPage") as HTMLButtonElement;
      const nextPageBtn = document.querySelector("#nextPage") as HTMLButtonElement;
      const lastPageBtn = document.querySelector("#lastPage") as HTMLButtonElement;
      const totalPagesSpan = document.querySelector("#totalPages") as HTMLSpanElement;
      const pageNumbersDiv = document.querySelector("#pageNumbers") as HTMLDivElement;
      const startItemSpan = document.querySelector("#startItem") as HTMLSpanElement;
      const endItemSpan = document.querySelector("#endItem") as HTMLSpanElement;
      const totalItemsSpan = document.querySelector("#totalItems") as HTMLSpanElement; 
      const remarksInternal = document.querySelector("#remarksInternal") as HTMLTextAreaElement;
      const remarksSprintCell = document.querySelector("#remarksSprintCell") as HTMLTextAreaElement;

      let projects: ProjectItem[] = [];
      let forecasts: RevenueForecastItem[] = [];
      let editingId: number | null = null;
      let allMonths: string[] = [];

      // Pagination variables
      let currentPage = 1;
      let pageSize = 25;
      let totalPages = 1;
      let filteredData: RevenueForecastItem[] = [];

      // ---- Filter state (inventory / procurement style) ---------------------
      type RevenueFilters = Record<'year' | 'month' | 'project' | 'customer' | 'manager' | 'groupHead' | 'reference', string>;
      const EMPTY_FILTERS: RevenueFilters = { year: '', month: '', project: '', customer: '', manager: '', groupHead: '', reference: '' };
      const FILTER_LABELS: Record<keyof RevenueFilters, string> = {
        year: 'Year', month: 'Month', project: 'Project', customer: 'Customer',
        manager: 'Project Manager', groupHead: 'Group Head', reference: 'Project Reference'
      };
      // ✅ Filters start CLEAN on every load - nothing is pre-selected.
      let appliedFilters: RevenueFilters = { ...EMPTY_FILTERS };

      // toolbar badge / chips / floating Apply pill / popup Reset Filter
      const filterTriggerBtn = document.querySelector('#revenueFilterTrigger') as HTMLButtonElement | null;
      const filterOverlay = document.querySelector('#revenueFilterOverlay') as HTMLElement | null;
      const filterCloseBtn = document.querySelector('#revenueFilterClose') as HTMLButtonElement | null;
      const filterCancelBtn = document.querySelector('#revenueFilterCancel') as HTMLButtonElement | null;
      const filterClearBtn = document.querySelector('#revenueFilterClear') as HTMLButtonElement | null;
      const filterApplyBtn = document.querySelector('#revenueFilterApply') as HTMLButtonElement | null;
      const filterCount = document.querySelector('#revenueFilterCount') as HTMLElement | null;
      const filterChipsHost = document.querySelector('#revenueFilterChips') as HTMLElement | null;
      const filterApplyPill = document.querySelector('#revenueFilterApplyPill') as HTMLButtonElement | null;
      const resetFilterBtn = document.querySelector('#revenueResetFilterBtn') as HTMLButtonElement | null;

      const FILTER_FIELD_IDS: Record<keyof RevenueFilters, string> = {
        year: 'filterYear', month: 'filterMonth', project: 'filterProject', customer: 'filterCustomer',
        manager: 'filterProjectManager', groupHead: 'filterGroupHead', reference: 'filterProjectReference'
      };

      const readDraftFilters = (): RevenueFilters => ({
        year: filterYear.value, month: filterMonth.value, project: filterProject.value,
        customer: filterCustomer.value, manager: filterProjectManager.value,
        groupHead: filterGroupHead.value, reference: filterProjectReference.value
      });

      const restoreAppliedFilters = (): void => {
        filterYear.value = appliedFilters.year; filterMonth.value = appliedFilters.month;
        filterProject.value = appliedFilters.project; filterCustomer.value = appliedFilters.customer;
        filterProjectManager.value = appliedFilters.manager; filterGroupHead.value = appliedFilters.groupHead;
        filterProjectReference.value = appliedFilters.reference;
      };

      const sameFilters = (a: RevenueFilters, b: RevenueFilters): boolean =>
        (Object.keys(EMPTY_FILTERS) as (keyof RevenueFilters)[]).every(k => (a[k] || '') === (b[k] || ''));

      const activeFilterCount = (): number => Object.values(appliedFilters).filter(Boolean).length;
      const hasDraftFilters = (): boolean => Object.values(readDraftFilters()).some(Boolean);
      const filtersAreDirty = (): boolean => !sameFilters(readDraftFilters(), appliedFilters);

      // chips for the applied filters (each removable, plus "Clear all")
      const renderFilterChips = (): void => {
        if (!filterChipsHost) return;
        const entries = (Object.keys(appliedFilters) as (keyof RevenueFilters)[])
          .filter(k => !!appliedFilters[k])
          .map(k => ({ key: k, value: appliedFilters[k] }));

        filterChipsHost.innerHTML = '';
        if (entries.length === 0) { filterChipsHost.hidden = true; return; }
        filterChipsHost.hidden = false;

        const label = document.createElement('span');
        label.className = 'revenue-filter-summary__label';
        label.textContent = 'Active filters:';
        filterChipsHost.appendChild(label);

        entries.forEach(({ key, value }) => {
          const chip = document.createElement('span');
          chip.className = 'revenue-filter-chip';
          const text = document.createElement('span');
          text.innerHTML = `${FILTER_LABELS[key]}: <strong>${value}</strong>`;
          const remove = document.createElement('button');
          remove.type = 'button';
          remove.title = 'Remove this filter';
          remove.setAttribute('aria-label', `Remove ${FILTER_LABELS[key]} filter`);
          remove.textContent = '×';
          remove.addEventListener('click', () => {
            appliedFilters = { ...appliedFilters, [key]: '' };
            restoreAppliedFilters();
            syncFilterUI();
            filterTable();
          });
          chip.appendChild(text);
          chip.appendChild(remove);
          filterChipsHost.appendChild(chip);
        });

        const clearAll = document.createElement('button');
        clearAll.type = 'button';
        clearAll.className = 'revenue-filter-chip__all';
        clearAll.textContent = 'Clear all';
        clearAll.addEventListener('click', () => clearFilters());
        filterChipsHost.appendChild(clearAll);
      };

      // single source of truth for badge / trigger state / popup buttons / pill / chips
      const syncFilterUI = (): void => {
        const count = activeFilterCount();
        if (filterCount) filterCount.textContent = String(count);
        filterTriggerBtn?.classList.toggle('is-active', count > 0);

        const canClear = count > 0 || hasDraftFilters();
        if (filterClearBtn) {
          filterClearBtn.disabled = !canClear;
          filterClearBtn.classList.toggle('is-enabled', canClear);
        }
        if (resetFilterBtn) {
          resetFilterBtn.disabled = !canClear;
          resetFilterBtn.classList.toggle('is-enabled', canClear);
        }

        const canApply = filtersAreDirty();
        if (filterApplyBtn) {
          filterApplyBtn.disabled = !canApply;
          filterApplyBtn.classList.toggle('is-enabled', canApply);
        }
        filterApplyPill?.classList.toggle('is-visible', canApply);

        // highlight the fields that carry a value (draft = what you see in the popup)
        const draft = readDraftFilters();
        Object.keys(draft).forEach(key => {
          const el = document.querySelector(`#${FILTER_FIELD_IDS[key as keyof RevenueFilters]}`) as HTMLSelectElement | null;
          el?.classList.toggle('is-set', !!draft[key as keyof RevenueFilters]);
        });

        renderFilterChips();
      };

      const openFilterModal = (): void => {
        restoreAppliedFilters();          // the draft always starts from what is applied
        filterOverlay?.classList.add('is-open');
        filterTriggerBtn?.setAttribute('aria-expanded', 'true');
        document.body.style.overflow = 'hidden';
        syncFilterUI();
      };

      const closeFilterModal = (restore = true): void => {
        if (restore) restoreAppliedFilters();
        filterOverlay?.classList.remove('is-open');
        filterTriggerBtn?.setAttribute('aria-expanded', 'false');
        document.body.style.overflow = '';
        syncFilterUI();
      };

      const commitFilters = (): void => {
        appliedFilters = readDraftFilters();
        syncFilterUI();
        filterTable();
      };

      const setupFilterEventListeners = (): void => {
        filterTriggerBtn?.addEventListener('click', openFilterModal);
        filterCloseBtn?.addEventListener('click', () => closeFilterModal());
        filterCancelBtn?.addEventListener('click', () => closeFilterModal());
        filterOverlay?.addEventListener('click', (event) => { if (event.target === filterOverlay) closeFilterModal(); });

        // choosing a value in the popup only enables the buttons + shows the pill -
        // the table changes only when Apply (or the pill) is pressed
        [filterYear, filterMonth, filterProject, filterCustomer, filterProjectManager, filterGroupHead, filterProjectReference].forEach(sel =>
          sel?.addEventListener('change', syncFilterUI));

        filterClearBtn?.addEventListener('click', () => clearFilters());
        resetFilterBtn?.addEventListener('click', () => clearFilters());
        filterApplyBtn?.addEventListener('click', () => { commitFilters(); closeFilterModal(false); });
        filterApplyPill?.addEventListener('click', commitFilters);

        syncFilterUI();
      };

      const pageSizeSelect = document.querySelector("#pageSizeSelect") as HTMLSelectElement;

      // const getSafeString = (value: any, defaultValue: string = "-"): string => {
      //   if (value === null || value === undefined) {
      //     return defaultValue;
      //   }
        
      //   if (typeof value === 'string') {
      //     return value.trim() || defaultValue;
      //   }
        
      //   if (typeof value === 'object' && 'Title' in value) {
      //     return value.Title?.trim() || defaultValue;
      //   }
        
      //   return String(value) || defaultValue;
      // };

      const toast = (icon: any, title: string) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });

      const getDigest = async (): Promise<string> => {
        const res = await fetch(`${siteUrl}/_api/contextinfo`, { 
          method: "POST", 
          headers: { Accept: "application/json;odata=verbose" } 
        });
        const data = await res.json();
        return data.d.GetContextWebInformation.FormDigestValue;
      };

      const safeJsonParse = (jsonString: string | undefined, fallback: any = {}): any => {
        try {
          if (!jsonString || jsonString.trim() === '') {
            return fallback;
          }
          
          const cleanedString = jsonString.trim();
          
          if (cleanedString === '""' || cleanedString === "null") {
            return fallback;
          }
          
          return JSON.parse(cleanedString);
        } catch (error) {
          console.error("JSON parse error:", error, "Input:", jsonString);
          return fallback;
        }
      };

      const formatMonthKey = (date: Date): string => {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
      };

      const formatMonthLabel = (date: Date): string => {
        return date.toLocaleString("default", { month: "short", year: "2-digit" });
      };

      const generateForecastTitle = (project: ProjectItem, forecastData: ForecastData): string => {
        const currentYear = new Date().getFullYear();
        let totalValue = 0;
        
        Object.keys(forecastData).forEach(key => {
          if (key.match(/^\d{4}-\d{2}$/)) {
            const monthData = forecastData[key] as ForecastMonthData;
            if (monthData) {
              totalValue += monthData.actual;
            }
          }
        });
        
        return `${project.Title} - ${currentYear} Forecast - ${totalValue.toLocaleString('en-US')}`;
      };

      const getDisplayMonths = (): string[] => {
        const now = new Date();
        const months: string[] = [];
        
        for (let i = 0; i < 3; i++) {
          const month = new Date(now.getFullYear(), now.getMonth() + i, 1);
          months.push(formatMonthKey(month));
        }
        
        return months;
      };

const getAllMonthsForStorage = (): string[] => {
  return generateIdealMonthStructure(); // Returns all months including 36 historical
};

      const allowOnlyNumbers = (input: HTMLInputElement): void => {
        input.addEventListener('input', function() {
          let value = this.value.replace(/[^\d]/g, '');
          
          if (value.length > 1 && value.startsWith('0')) {
            value = value.replace(/^0+/, '');
            if (value === '') value = '0';
          }
          
          this.value = value;
          
          if (value === '' || value === '0') {
            this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
          } else {
            this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
          }
        });
        
        input.addEventListener('keydown', function(e) {
          if ([8, 9, 13, 27, 46].includes(e.keyCode) || 
              (e.keyCode === 65 && e.ctrlKey === true) || 
              (e.keyCode === 67 && e.ctrlKey === true) ||
              (e.keyCode === 86 && e.ctrlKey === true) ||
              (e.keyCode === 88 && e.ctrlKey === true) ||
              (e.keyCode >= 35 && e.keyCode <= 39)) {
            return;
          }
          
          if ((e.shiftKey || (e.keyCode < 48 || e.keyCode > 57)) && (e.keyCode < 96 || e.keyCode > 105)) {
            e.preventDefault();
          }
        });
        
        input.addEventListener('paste', function(e) {
          const clipboardData = e.clipboardData || (window as any).clipboardData;
          const pastedData = clipboardData?.getData('Text') || '';
          
          if (/[^\d]/.test(pastedData)) {
            e.preventDefault();
          }
        });
      };
// ============================================
// OBL TOAST FUNCTIONS - CORRECTED
// ============================================

// Helper function to get numeric value from input
const getNumericValue = (input: HTMLInputElement): number => {
  const value = input.value.trim();
  if (value === '' || value === '0') return 0;
  return parseInt(value) || 0;
};

// Helper function to set numeric value in input
const setNumericValue = (input: HTMLInputElement, value: number): void => {
  if (value === 0) {
    input.value = '';
    input.className = input.className.replace('revenue-input-with-value', 'revenue-input-empty');
  } else {
    input.value = value.toString();
    input.className = input.className.replace('revenue-input-empty', 'revenue-input-with-value');
  }
};

// SHOW OBL TOAST - CORRECTED 
const showOBLToast = (type: 'success' | 'info' | 'warning' | 'danger', message: string): void => {
  // Remove existing toast
  const existingToast = document.querySelector('.revenue-obl-error-toast');
  if (existingToast) existingToast.remove();
  
  const toast = document.createElement('div');
  toast.className = `revenue-obl-error-toast ${type}`;
  
  document.body.appendChild(toast);
  
  // Auto dismiss after 5 seconds
  const timeoutId = setTimeout(() => {
    if (toast.parentNode) {
      toast.style.animation = 'slideOutRight 0.3s ease-in';
      setTimeout(() => {
        if (toast.parentNode) toast.remove();
      }, 300);
    }
  }, 5000);
  
  // Close button handler
  toast.querySelector('.revenue-toast-close')?.addEventListener('click', () => {
    clearTimeout(timeoutId);
    toast.style.animation = 'slideOutRight 0.3s ease-in';
    setTimeout(() => {
      if (toast.parentNode) toast.remove();
    }, 300);
  });
};

// CALCULATE TOTAL CURRENT VALUE - CORRECTED 
const calculateTotalCurrentValue = (): number => {
  let total = 0;
  
  // Helper to get numeric value
  const getVal = (input: HTMLInputElement): number => getNumericValue(input);
  
  // ONLY count revenue-item-current-input values (these represent individual line items)
  // The main "Curr. Plan" is the sum of all items, so counting both would double count
  document.querySelectorAll<HTMLInputElement>('.revenue-detailed-month-box .revenue-item-current-input').forEach(input => {
    total += getVal(input);
  });
  
  // Long-term Forecast (Next 15 Months) - Curr. Plan values (these are standalone, not itemized)
  document.querySelectorAll<HTMLInputElement>('.revenue-simple-month-rev .revenue-main-current-input:not([readonly])').forEach(input => {
    total += getVal(input);
  });
  
  // 5-Year Forecast - Curr. Plan values (these are standalone, not itemized)
  document.querySelectorAll<HTMLInputElement>('.revenue-yearly-current-input-rev:not([readonly])').forEach(input => {
    total += getVal(input);
  });
  
  return total;
};

// VALIDATE OBL LIMIT - FIXED
const validateOBLLimit = (
  changedInput?: HTMLInputElement
): boolean => {
  const oblValue = getNumericValue(oblLastMonthInput);
  const totalCurrent = calculateTotalCurrentValue();

  totalCurrentValueSpan.textContent =
    totalCurrent.toLocaleString('en-US');

  const remaining = oblValue - totalCurrent;

  remainingValueSpan.textContent =
    remaining.toLocaleString('en-US');

  // Remove existing OBL toast
  const existingToast =
    document.querySelector(".revenue-obl-error-toast");

  if (existingToast) {
    existingToast.remove();
  }

  // =====================================================
  // OBL = 0
  // =====================================================
  if (oblValue === 0) {
    oblLineFill.style.width = "0%";
    oblLineFill.className = "revenue-line-fill";

    if (totalCurrent > 0) {
      oblValidationMessage.textContent =
        `❌ Liquidation is more than OBL (More by ${totalCurrent.toLocaleString('en-US')})`;

      oblValidationMessage.className =
        "revenue-validation-message danger";

      showOBLToast(
        "danger",
        `❌ Liquidation is more than OBL (More by ${totalCurrent.toLocaleString('en-US')})`
      );

      // 🔧 Liquidation > OBL: warn only - do NOT lock the inputs
      // (no not-allowed pointer); the form must still be submittable.
      unblockAllCurrentInputs();

      return false;
    }

    oblValidationMessage.textContent =
      "ℹ️ OBL is set to zero.";

    oblValidationMessage.className =
      "revenue-validation-message";

    showOBLToast(
      "info",
      "ℹ️ OBL is set to zero."
    );

    unblockAllCurrentInputs();

    return true;
  }

  // =====================================================
  // Calculate percentage
  // =====================================================
  const percentage =
    Math.min(
      (totalCurrent / oblValue) * 100,
      100
    );

  oblLineFill.style.width =
    `${percentage}%`;

  oblLineFill.className =
    "revenue-line-fill";

  if (percentage > 90) {
    oblLineFill.classList.add("danger");
  } else if (percentage > 75) {
    oblLineFill.classList.add("warning");
  }

  // =====================================================
  // 1. LIQUIDATION > OBL
  // =====================================================
  if (totalCurrent > oblValue) {
    const exceedAmount =
      totalCurrent - oblValue;

    oblValidationMessage.textContent =
      `❌ Liquidation is more than OBL (More by ${exceedAmount.toLocaleString('en-US')})`;

    oblValidationMessage.className =
      "revenue-validation-message danger";

    showOBLToast(
      "danger",
      `❌ Liquidation is more than OBL (More by ${exceedAmount.toLocaleString('en-US')})`
    );

    // 🔧 Liquidation > OBL: warn only - do NOT lock the inputs
    // (no not-allowed pointer); the form must still be submittable.
    unblockAllCurrentInputs();

    return false;
  }

  // =====================================================
  // 2. LIQUIDATION = OBL
  // =====================================================
  if (totalCurrent === oblValue) {
    oblValidationMessage.textContent =
      `✅ OBL is fully liquidated (${totalCurrent.toLocaleString('en-US')})`;

    oblValidationMessage.className =
      "revenue-validation-message success";

    showOBLToast(
      "success",
      `✅ OBL is fully liquidated (${totalCurrent.toLocaleString('en-US')})`
    );

    // Don't allow increasing values beyond OBL
    blockIncreasingCurrentInputs(
      changedInput
    );

    return true;
  }

  // =====================================================
  // 3. LIQUIDATION < OBL
  // =====================================================
  if (totalCurrent < oblValue) {
    const remainingAmount =
      oblValue - totalCurrent;

    oblValidationMessage.textContent =
      `⚠️ OBL is not fully liquidated (Remaining ${remainingAmount.toLocaleString('en-US')})`;

    oblValidationMessage.className =
      "revenue-validation-message warning";

    showOBLToast(
      "warning",
      `⚠️ OBL is not fully liquidated (Remaining ${remainingAmount.toLocaleString('en-US')})`
    );

    unblockAllCurrentInputs();

    return true;
  }

  // =====================================================
  // Default
  // =====================================================
  unblockAllCurrentInputs();

  return true;
};

// UPDATE OBL VALIDATION DISPLAY - CORRECTED
const updateOBLValidationDisplay = (changedInput?: HTMLInputElement): void => {
  validateOBLLimit(changedInput);
  validateMonthItemTotals(changedInput);
};

      const createEmptyItems = (): { [key: string]: ForecastItem } => {
          const items: { [key: string]: ForecastItem } = {};
          for (let i = 1; i <= 10; i++) {
              items[i.toString()] = { 
                  name: '', 
                  plan: 0, 
                  actual: 0, 
                  forecast: 0, 
                  date: '' 
              };
          }
          return items;
      };

const generateIdealMonthStructure = (): string[] => {
    const months: string[] = [];

    // IMPORTANT:
    // Always use the FIRST day of the current month.
    // Otherwise dates like 29/30/31 can cause month skipping
    // when setMonth() is used.
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();

    // 36 historical months
    // Includes the previous 36 months.
    for (let i = -36; i < 0; i++) {
        const date = new Date(currentYear, currentMonth + i, 1);
        months.push(formatMonthKey(date));
    }

    // 3 detailed months:
    // Current month + next 2 months
    for (let i = 0; i < 3; i++) {
        const date = new Date(currentYear, currentMonth + i, 1);
        months.push(formatMonthKey(date));
    }

    // 15 future months
    for (let i = 3; i < 18; i++) {
        const date = new Date(currentYear, currentMonth + i, 1);
        months.push(formatMonthKey(date));
    }

    // Remove duplicates and make sure months are chronological
    return Array.from(new Set(months)).sort();
};

      const getLastMonthKey = (existingData: ForecastData): string | undefined => {
        const monthKeys = Object.keys(existingData).filter(key => key.match(/^\d{4}-\d{2}$/));
        
        if (monthKeys.length === 0) {
          return undefined;
        }
        
        const sortedMonths = monthKeys.sort();
        const lastMonth = sortedMonths[sortedMonths.length - 1];
         
        return lastMonth;
      };

      const createFreshYearlyStructure = (lastMonthKey?: string): YearlyForecastData => {
        const yearlyData: YearlyForecastData = {};
        
        let startYear: number;
        
        if (lastMonthKey) {
          startYear = parseInt(lastMonthKey.split('-')[0]); 
        } else {
          const futureDate = new Date();
          futureDate.setMonth(futureDate.getMonth() + 15);
          startYear = futureDate.getFullYear(); 
        }
        
        const endYear = startYear + 4;
        
        for (let year = startYear; year <= endYear; year++) {
          yearlyData[year.toString()] = {
            plan: 0,
            actual: 0
          };
        }
        
        return yearlyData;
      };

const createFreshForecastStructure = (): ForecastData => {
    const allMonths = generateIdealMonthStructure();
    const freshData: ForecastData = {};

    const now = new Date();
    const currentMonthKey = formatMonthKey(now);

    const currentMonthIndex = allMonths.indexOf(currentMonthKey);

    let detailedStartIndex = 0;

    if (currentMonthIndex !== -1) {
        detailedStartIndex = currentMonthIndex;
    } else {
        detailedStartIndex = 36;
    }

    allMonths.forEach((monthKey, index) => {

        const isDetailedMonth =
            index >= detailedStartIndex &&
            index < detailedStartIndex + 3;

        const monthData: ForecastMonthData = {
            plan: 0,
            actual: 0,
            forecast: 0,
            monthActual: 0,  // ✅ Initialize monthActual
            items: isDetailedMonth
                ? createEmptyItems()
                : {}
        };

        freshData[monthKey] = monthData;
    });

    const lastMonthKey = allMonths[allMonths.length - 1];

    freshData.yearly =
        createFreshYearlyStructure(lastMonthKey);

    return freshData;
};

      const isForecastMonthData = (data: any): data is ForecastMonthData => {
          return data && 
                typeof data.plan === 'number' && 
                typeof data.actual === 'number';
      };

const propagateMonthValues = (forecastData: ForecastData): ForecastData => {
    const now = new Date();
    const currentMonthKey = formatMonthKey(now);
    
    const originalMonths = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
    const idealMonths = generateIdealMonthStructure();
    
    const hasCurrentMonth = originalMonths.includes(currentMonthKey);
    const lastDataMonth = originalMonths[originalMonths.length - 1];
    const lastIdealMonth = idealMonths[idealMonths.length - 1];
    
    if (hasCurrentMonth && lastDataMonth === lastIdealMonth) {
        return forecastData;
    }
    
    const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
    const detailedStartIndex = currentMonthIndex !== -1 ? currentMonthIndex : 36;
    
    const updatedData: ForecastData = {};
    
    idealMonths.forEach((monthKey, index) => {
        const existingData = forecastData[monthKey];
        const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
        
        if (existingData && isForecastMonthData(existingData)) {
            updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
        } else {
            updatedData[monthKey] = {
                plan: 0,
                actual: 0,
                monthActual: 0,
                items: isDetailedMonth ? createEmptyItems() : {}
            };
        }
    });
    
    console.log("=== APPLYING VALUE PROPAGATION ===");
    
    for (let i = 1; i < idealMonths.length; i++) {
        const prevMonthKey = idealMonths[i - 1];
        const currMonthKey = idealMonths[i];
        
        const prevMonthData = updatedData[prevMonthKey];
        const currMonthData = updatedData[currMonthKey];
        
        if (prevMonthData && isForecastMonthData(prevMonthData) && 
            currMonthData && isForecastMonthData(currMonthData)) {
            
            const isPrevMonthOlder = prevMonthKey < currentMonthKey;
            const isCurrMonthNew = !originalMonths.includes(currMonthKey);
            
            if (isPrevMonthOlder && isCurrMonthNew) {
                // ✅ Use monthActual from previous month if available
                const prevMonthActual = prevMonthData.monthActual || 0;
                
                if (prevMonthActual > 0) {
                    currMonthData.plan = prevMonthActual;
                } else {
                    currMonthData.plan = prevMonthData.actual;
                }
                
                currMonthData.actual = 0;
                currMonthData.monthActual = 0;
                
                const monthIndex = idealMonths.indexOf(currMonthKey);
                const isDetailedMonth = monthIndex >= detailedStartIndex && monthIndex < detailedStartIndex + 3;
                
                if (isDetailedMonth && currMonthData.items) {
                    const prevItems = prevMonthData.items || {};
                    Object.keys(currMonthData.items).forEach(itemKey => {
                        const prevItem = prevItems[itemKey];
                        const currItem = currMonthData.items![itemKey];
                        if (prevItem && currItem) {
                            if (prevMonthActual > 0) {
                                currItem.plan = prevMonthActual;
                            } else {
                                currItem.plan = prevItem.actual || 0;
                            }
                            currItem.actual = 0;
                        }
                    });
                }
            }
        }
    }
    
    if (forecastData.yearly) {
        updatedData.yearly = JSON.parse(JSON.stringify(forecastData.yearly));
    } else {
        const lastMonthKey = idealMonths[idealMonths.length - 1];
        updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
    }
    
    return updatedData;
};
const applyAutoProgression = async (): Promise<void> => { 
    if (forecasts.length === 0) { 
        return;
    }

    const digest = await getDigest();
    let progressedCount = 0;
    let errorCount = 0;
    const entityType = await getListEntityType();
    
    for (const forecast of forecasts) {
        try {
            if (!forecast.ForecastData) { 
                continue;
            }
            
            const originalData = safeJsonParse(forecast.ForecastData, {});
            
            if (checkIfNeedsProgression(originalData)) {  
                const etag = forecast.__metadata?.etag;
                if (!etag) {
                    errorCount++;
                    console.error(`Cannot safely progress Revenue forecast ${forecast.Id}: SharePoint ETag is unavailable`);
                    continue;
                }
                // ✅ FIX: Call with only 1 argument
                const progressedData = propagateCollectionValues(originalData);
                
                const payload = {
                    __metadata: { type: entityType },
                    ForecastData: JSON.stringify(progressedData),
                    UpdatedDate: new Date().toISOString()
                };
                
                const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${forecast.Id})`, {
                    method: "POST",
                    headers: {
                        Accept: "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": digest,
                        "IF-MATCH": etag,
                        "X-HTTP-Method": "MERGE"
                    },
                    body: JSON.stringify(payload)
                });
                
                if (response.ok) {
                    progressedCount++; 
                } else {
                    errorCount++; 
                    console.error(`Failed to progress Revenue forecast ${forecast.Id} (HTTP ${response.status})`);
                }
                
                await new Promise(resolve => setTimeout(resolve, 100));
            }
        } catch (error) {
            errorCount++; 
        }
    }
    
    if (progressedCount > 0) { 
        await loadForecasts();
        toast("success", `Monthly data progressed for ${progressedCount} revenue forecasts`);
    }
};

const loadProjects = async (): Promise<void> => {
  try {
    // ✅ Expand GroupHead1, ProjectManager AND the engineering person/lookup
    // fields, because getAssignedProjects() (auth.ts) matches the logged-in
    // user against these fields for KRIYA_Engg_Lead / Engg_Sub1 / Engg_Sub2 /
    // Engg_Mgr. Without them in $select+$expand those roles get 0 projects.
    const res = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items` +
      `?$select=Id,Title,CustomerName,PC,ProjectReference,` +
      `GroupHead1/Title,GroupHead1Id,` +
      `ProjectManager/Title,ProjectManagerId,` +
      `EnggLead/Title,EnggLeadId,` +
      `EnggSubLead1/Title,EnggSubLead1Id,` +
      `EnggSubLead2/Title,EnggSubLead2Id,` +
      `EngineeringManagers/Title,EngineeringManagersId` +
      `&$expand=GroupHead1,ProjectManager,EnggLead,EnggSubLead1,EnggSubLead2,EngineeringManagers` +
      `&$top=5000`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    const data: SharePointResponse<ProjectItem> = await res.json();
    const rawProjects: ProjectItem[] = data.d?.results || [];

    const cleanedProjects = rawProjects.map((p: ProjectItem) => ({
      ...p,
      CustomerName: p.CustomerName || "",
      GroupHead1: p.GroupHead1 || "",
      GroupHead1Id: p.GroupHead1Id,
      PC: p.PC || "",
      ProjectReference: p.ProjectReference || "",
      ProjectManager: p.ProjectManager || "",
      ProjectManagerId: p.ProjectManagerId,
      // ✅ Keep the engineering fields so getAssignedProjects() can match
      EnggLead: p.EnggLead || "",
      EnggLeadId: p.EnggLeadId,
      EnggSubLead1: p.EnggSubLead1 || "",
      EnggSubLead1Id: p.EnggSubLead1Id,
      EnggSubLead2: p.EnggSubLead2 || "",
      EnggSubLead2Id: p.EnggSubLead2Id,
      EngineeringManagers: p.EngineeringManagers || "",
      EngineeringManagersId: p.EngineeringManagersId
    }));

    const projectMap = new Map<string, ProjectItem>();
    cleanedProjects.forEach(p => {
      projectMap.set(p.Title, p);
    });

    projects = Array.from(projectMap.values());

    // Log all projects with every role owner for debugging
    console.log('📋 All projects with role assignments:');
    projects.forEach(p => {
      const pmName = typeof p.ProjectManager === 'object'
        ? p.ProjectManager?.Title
        : p.ProjectManager;
      const ghName = typeof p.GroupHead1 === 'object'
        ? p.GroupHead1?.Title
        : p.GroupHead1;
      const elName = typeof p.EnggLead === 'object'
        ? p.EnggLead?.Title
        : p.EnggLead;
      const es1Name = typeof p.EnggSubLead1 === 'object'
        ? p.EnggSubLead1?.Title
        : p.EnggSubLead1;
      const es2Name = typeof p.EnggSubLead2 === 'object'
        ? p.EnggSubLead2?.Title
        : p.EnggSubLead2;
      const emName = typeof p.EngineeringManagers === 'object'
        ? p.EngineeringManagers?.Title
        : p.EngineeringManagers;
      console.log(
        `  - ${p.Title}: PM=${pmName || 'None'}, GH=${ghName || 'None'}, ` +
        `EL=${elName || 'None'}, ES1=${es1Name || 'None'}, ` +
        `ES2=${es2Name || 'None'}, EM=${emName || 'None'}`
      );
    });

    const assignedProjects = getAssignedProjects(projects);

    console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);
    console.log('📋 Assigned project titles:', assignedProjects.map(p => p.Title));

    // ✅ Populate project select dropdown with assigned projects ONLY
    projectSelect.innerHTML = `<option value="">Select Project</option>` +
      assignedProjects.map(p => `<option value="${p.Id}">${p.Title}</option>`).join("");

    // ✅ Also update the filter dropdown with assigned projects
    filterProject.innerHTML = '<option value="">All Projects</option>' +
      [...new Set(assignedProjects.map(p => p.Title))]
        .sort()
        .map(project => `<option value="${project}">${project}</option>`)
        .join("");

  } catch (error) {
    console.error("Error loading projects:", error);
    toast("error", "Failed to load projects");
  }
};

// Helper functions for Person/Group display
const getGroupHeadDisplayName = (groupHead: string | { Title: string; Email?: string } | undefined): string => {
  if (!groupHead) return "";
  if (typeof groupHead === "object" && groupHead !== null) {
    return groupHead.Title || "";
  }
  return groupHead;
};

const getProjectManagerDisplayName = (manager: string | { Title: string; Email?: string } | undefined): string => {
  if (!manager) return "";
  if (typeof manager === "object" && manager !== null) {
    return manager.Title || "";
  }
  return manager;
};

const getAssignedProjectValues = (key: keyof ProjectItem): string[] => {
  const assignedProjects = getAssignedProjects(projects);
  const values = new Set<string>();
  
  assignedProjects.forEach(project => {
    let value = project[key];
    if (value) {
      // Handle ProjectManager which can be string or object
      if (key === 'ProjectManager' && typeof value === 'object') {
        value = (value as { Title: string }).Title;
      }
      // ✅ Handle GroupHead1 which can be string or object
      if (key === 'GroupHead1' && typeof value === 'object') {
        value = (value as { Title: string }).Title;
      }
      if (typeof value === 'string' && value.trim()) {
        values.add(value.trim());
      }
    }
  });
  
  return Array.from(values).sort();
};

  const loadForecasts = async (projectsReady?: Promise<void>): Promise<void> => {
  try {
    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,RevenueActual,Remarks,CreatedDate,UpdatedDate,Freeze&$expand=Project&$top=5000`, {
      headers: { Accept: "application/json;odata=verbose" }
    });
    const data: SharePointResponse<RevenueForecastItem> = await res.json();
    if (projectsReady) await projectsReady;
    
    forecasts = (data.d?.results || []).map(forecast => ({
      ...forecast,
      ForecastData: typeof forecast.ForecastData === 'string' ? 
        forecast.ForecastData : 
        JSON.stringify(forecast.ForecastData || {}),
      RevenueActual: typeof forecast.RevenueActual === 'string' ? 
        forecast.RevenueActual : 
        JSON.stringify(forecast.RevenueActual || {}),
      Freeze: forecast.Freeze || false
    }));
    
    console.log("✅ Forecasts loaded:", forecasts.length);
  } catch (error) { 
    console.error("Error loading forecasts:", error);
    toast("error", "Failed to load forecasts");
  }
};

// ============================================
// AUTO-SYNC MISSING PROJECTS → REVENUE FORECAST LIST
// Every project that exists in Project List but has NO row in Revenue
// Forecast List yet is added automatically, with an empty ForecastData
// skeleton (createFreshForecastStructure: historical months + current/next
// 2 + long-term months + 4-year yearly block).
// 🔐 RBAC: runs only for roles with create rights (Admin, GH Heads, PM
// Managers) and only adds projects inside the role's visible scope.
// Idempotent: existing rows are never duplicated or overwritten.
// ============================================

// Cached check: does Revenue Forecast List have a 'ProjectReference' column,
// and is it plain text or a Lookup to Project List?
let revenueProjectRefField: { exists: boolean; type: string } | null = null;
const revenueListProjectRefField = async (): Promise<{ exists: boolean; type: string }> => {
  if (revenueProjectRefField) return revenueProjectRefField;
  try {
    const res = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq 'ProjectReference'&$select=InternalName,TypeAsString,ReadOnlyField`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    const data = await res.json();
    const results = data?.d?.results || [];
    if (results.length > 0) {
      revenueProjectRefField = { exists: true, type: String(results[0].TypeAsString || "Text") };
      console.log(`📋 Revenue Forecast List ProjectReference column type = ${revenueProjectRefField.type}`);
    } else {
      revenueProjectRefField = { exists: false, type: "" };
      console.log("📋 Revenue Forecast List has no ProjectReference column");
    }
  } catch (error) {
    console.error("⚠️ Could not check Revenue ProjectReference column:", error);
    revenueProjectRefField = { exists: false, type: "" };
  }
  return revenueProjectRefField;
};
  
const syncMissingProjectsToRevenueList = async (): Promise<void> => {
  const syncStatus = document.querySelector('#revenueSyncStatus') as HTMLElement | null;
  const syncBtn = document.querySelector('#revenueSyncProjectsBtn') as HTMLButtonElement | null;
  const syncLockKey = `${siteUrl}::${listName}`;
  const windowState = window as any;
  const syncLocks: Set<string> = windowState.__revenueSyncLocks || new Set<string>();
  windowState.__revenueSyncLocks = syncLocks;

  const setStatus = (text: string): void => {
    console.log(text);
    if (syncStatus) syncStatus.textContent = text;
  };

  if (!canCreateForecast()) {
    setStatus('👁️ No create rights - sync skipped (Admin/GH Head/PM Manager only)');
    return;
  }

  if (!projects || projects.length === 0) {
    setStatus('⚠️ Project List is empty - nothing to sync');
    return;
  }

  if (syncLocks.has(syncLockKey)) {
    setStatus('⏳ Revenue project sync is already running');
    return;
  }
  syncLocks.add(syncLockKey);

  if (syncBtn) syncBtn.disabled = true;
  try {
    const visibleProjects = canReadAllProjects()
      ? projects
      : getAssignedProjects(projects);
    const refField = await revenueListProjectRefField();
    const hasTextProjectReference = refField.exists && refField.type !== "Lookup";
    const normalizeReference = (value: unknown): string =>
      String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
    const getServerProjectIds = (rows: any[]): Set<number> => {
      const ids = new Set<number>();
      rows.forEach(row => {
        [row.ProjectId, row.ProjectReferenceId, row.Project?.Id].forEach(value => {
          const id = Number(value);
          if (value !== null && value !== undefined && Number.isFinite(id) && id > 0) ids.add(id);
        });
      });
      return ids;
    };

    // Safety re-check against SharePoint itself: the in-memory list may be
    // empty/stale, so never create rows for projects that already exist.
    const selectProjectReference = refField.exists
      ? (hasTextProjectReference ? ",ProjectReference" : ",ProjectReferenceId")
      : "";
    const existingRowsUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,ProjectId,Title,Project/Id,Project/Title${selectProjectReference}&$expand=Project&$top=5000`;
    const chkRes = await fetch(existingRowsUrl, {
      headers: { Accept: "application/json;odata=verbose" }
    });
    if (!chkRes.ok) {
      throw new Error(`Could not verify existing Revenue rows (HTTP ${chkRes.status}); sync stopped to prevent duplicates.`);
    }
    const serverRows: any[] = (await chkRes.json())?.d?.results || [];

    // If the list really has rows on SharePoint but the page loaded 0, the
    // main load is broken - creating rows now would duplicate everything.
    if (serverRows.length > 0 && forecasts.length === 0) {
      setStatus(`⚠️ Revenue Forecast List already has ${serverRows.length} row(s) on SharePoint but the page loaded 0 - sync stopped to avoid duplicates. Check the main forecast load.`);
      toast("error", `Sync stopped: list has ${serverRows.length} rows but page loaded 0`);
      return;
    }

    const serverProjectIds = getServerProjectIds(serverRows);
    const serverProjectReferences = new Set(
      hasTextProjectReference
        ? serverRows.map(row => normalizeReference(row.ProjectReference)).filter(Boolean)
        : []
    );
    const forecastProjectIds = new Set<number>();
    const forecastProjectTitles = new Set<string>();
    const forecastProjectReferences = new Set<string>();
    const projectsById = new Map<number, ProjectItem>();
    const projectsByTitle = new Map<string, ProjectItem>();
    projects.forEach(project => {
      projectsById.set(project.Id, project);
      projectsByTitle.set(project.Title, project);
    });
    forecasts.forEach(forecast => {
      if (forecast.Project?.Id !== undefined) forecastProjectIds.add(forecast.Project.Id);
      if (forecast.Project?.Title) forecastProjectTitles.add(forecast.Project.Title);
      if (hasTextProjectReference) {
        const linkedProjectId = forecast.Project?.Id;
        const linkedProjectTitle = forecast.Project?.Title;
        const project = (linkedProjectId !== undefined ? projectsById.get(linkedProjectId) : undefined) ||
          (linkedProjectTitle ? projectsByTitle.get(linkedProjectTitle) : undefined);
        const reference = normalizeReference(project?.ProjectReference);
        if (reference) forecastProjectReferences.add(reference);
      }
    });

    const existsOnServer = (p: ProjectItem): boolean =>
      serverProjectIds.has(p.Id) ||
      (hasTextProjectReference && serverProjectReferences.has(normalizeReference(p.ProjectReference)));
    const existsInMemory = (p: ProjectItem): boolean =>
      forecastProjectIds.has(p.Id) ||
      forecastProjectTitles.has(p.Title) ||
      (hasTextProjectReference && forecastProjectReferences.has(normalizeReference(p.ProjectReference)));

    const missingProjects = visibleProjects.filter(p => !existsOnServer(p) && !existsInMemory(p));

    if (missingProjects.length === 0) {
      // setStatus(`✅ In sync: all ${visibleProjects.length} visible project(s) have a Revenue Forecast List row`);
      return;
    }

    // setStatus(`🆕 ${missingProjects.length} missing: ${missingProjects.map(p => p.Title).join(', ')}`);
    // toast("info", `Checking ${visibleProjects.length} projects... adding ${missingProjects.length} missing`);

    const digest = await getDigest();
    const entityType = await getListEntityType();
    const freshData = createFreshForecastStructure();
    const freshDataJson = JSON.stringify(freshData);
    let added = 0;
    let failed = 0;
    const errors: string[] = [];

    const createForecastForProject = async (project: ProjectItem): Promise<void> => {
      let ok = false;
      let stopAfterThrottle = false;
      let lastError = "Unknown SharePoint error";

      // Keep the existing payload fallbacks for optional/read-only columns.
      for (let payloadAttempt = 1; payloadAttempt <= 3 && !ok && !stopAfterThrottle; payloadAttempt++) {
        const payload: any = {
          __metadata: { type: entityType },
          Title: generateForecastTitle(project, freshData),
          ProjectId: project.Id.toString(),
          ProfitCentre: project.PC || "",
          ForecastData: freshDataJson,
          RevenueActual: "{}",
          Remarks: "",
          OBL: "0",
          CreatedDate: new Date().toISOString(),
          UpdatedDate: new Date().toISOString()
        };

        if (refField.exists && payloadAttempt === 1) {
          if (refField.type === "Lookup") {
            payload.ProjectReferenceId = String(project.Id);
          } else {
            payload.ProjectReference = project.ProjectReference || "";
          }
        }
        if (payloadAttempt >= 2) {
          delete payload.ProjectReference;
          delete payload.ProjectReferenceId;
        }
        if (payloadAttempt >= 3) delete payload.OBL;

        let throttleAttempt = 0;
        while (!ok && !stopAfterThrottle) {
          try {
            const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
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
              break;
            }

            const bodyText = await res.text();
            lastError = `HTTP ${res.status}`;
            try {     
              const errJson = JSON.parse(bodyText);
              lastError = errJson?.error?.message?.value || lastError;
            } catch {
              if (bodyText) lastError = bodyText.slice(0, 250);
            }

            if (res.status === 408 || res.status >= 500) {
              stopAfterThrottle = true;
              console.error(`Revenue create returned HTTP ${res.status} for ${project.Title}; it will not be reposted automatically.`);
              break;
            }

            if (res.status === 429) {
              const retryAfter = res.headers.get("Retry-After");
              const retryAfterSeconds = retryAfter === null ? NaN : Number(retryAfter);
              const retryAfterDate = retryAfter && Number.isNaN(retryAfterSeconds)
                ? Date.parse(retryAfter)
                : NaN;
              if (throttleAttempt >= 5) {
                stopAfterThrottle = true;
                console.error(`❌ Throttled too many times while creating ${project.Title}: ${bodyText}`);
                break;
              }

              const waitMs = Number.isFinite(retryAfterSeconds)
                ? Math.max(0, retryAfterSeconds * 1000)
                : Number.isFinite(retryAfterDate)
                  ? Math.max(0, retryAfterDate - Date.now())
                  : Math.min(60000, 1000 * (2 ** throttleAttempt));
              console.warn(`SharePoint throttled Revenue sync; retrying in ${Math.ceil(waitMs / 1000)}s.`);
              throttleAttempt++;
              await new Promise(resolve => setTimeout(resolve, waitMs));
              continue;
            }

            console.error(`❌ Payload attempt ${payloadAttempt} failed for ${project.Title}: ${bodyText}`);
            break;
          } catch (projectError) {
            lastError = projectError instanceof Error ? projectError.message : String(projectError);
            console.error(`❌ Payload attempt ${payloadAttempt} error for ${project.Title}:`, projectError);
            stopAfterThrottle = true;
            break;
          }
        }

        if (!ok && !stopAfterThrottle && payloadAttempt < 3) {
          await new Promise(resolve => setTimeout(resolve, 300));
        }
      }

      if (!ok) {
        failed++;
        errors.push(`${project.Title} -> ${lastError}`);
      }
      await new Promise(resolve => setTimeout(resolve, 100));
    };

    const batchPayloads = missingProjects.map((project): Record<string, unknown> => {
      const payload: Record<string, unknown> = {
        __metadata: { type: entityType },
        Title: generateForecastTitle(project, freshData),
        ProjectId: project.Id.toString(),
        ProfitCentre: project.PC || "",
        ForecastData: freshDataJson,
        RevenueActual: "{}",
        Remarks: "",
        OBL: "0",
        CreatedDate: new Date().toISOString(),
        UpdatedDate: new Date().toISOString()
      };
      if (refField.exists) {
        if (refField.type === "Lookup") payload.ProjectReferenceId = String(project.Id);
        else payload.ProjectReference = project.ProjectReference || "";
      }
      return payload;
    });
    const batchResults = await createSharePointItemsInBatches(siteUrl, listName, digest, batchPayloads);
    const failedBatchProjects: ProjectItem[] = [];
    batchResults.forEach((result, index) => {
      if (result.success) {
        added++;
      } else {
        failedBatchProjects.push(missingProjects[index]);
        console.error(`Revenue batch create failed for ${missingProjects[index].Title}: ${result.error}`);
      }
    });

    if (failedBatchProjects.length > 0) {
      const verifyResponse = await fetch(existingRowsUrl, {
        headers: { Accept: "application/json;odata=verbose" }
      });
      if (!verifyResponse.ok) {
        throw new Error(`Could not verify Revenue batch results (HTTP ${verifyResponse.status}); individual retries were stopped to avoid duplicate rows.`);
      }
      const verifiedRows = (await verifyResponse.json())?.d?.results || [];
      const verifiedProjectIds = getServerProjectIds(verifiedRows);
      const verifiedProjectReferences = new Set(
        hasTextProjectReference
          ? verifiedRows.map((row: any) => normalizeReference(row.ProjectReference)).filter(Boolean)
          : []
      );
      const fallbackProjects = failedBatchProjects.filter(project =>
        !verifiedProjectIds.has(project.Id) &&
        !(hasTextProjectReference && verifiedProjectReferences.has(normalizeReference(project.ProjectReference)))
      );

      if (fallbackProjects.length < failedBatchProjects.length) {
        console.warn("Revenue batch verification found rows already created; skipping duplicate retries.");
      }
      const fallbackConcurrency = 4;
      for (let i = 0; i < fallbackProjects.length; i += fallbackConcurrency) {
        await Promise.all(fallbackProjects.slice(i, i + fallbackConcurrency).map(createForecastForProject));
      }
    }

    if (added > 0) {
      await loadForecasts();
      console.log(`➕ Revenue auto-sync added ${added} project(s)` + (failed > 0 ? ` (${failed} failed)` : ""));
    } else if (failed > 0) {
      console.error(`❌ Revenue auto-sync failed for ${failed} project(s): ${errors.join(' | ')}`);
    }
  } catch (error) {
    console.error("❌ Revenue auto-sync failed:", error);
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`❌ Revenue sync stopped: ${message}`);
    toast("error", `Revenue sync stopped: ${message}`);
  } finally {
    if (syncBtn) syncBtn.disabled = false;
    syncLocks.delete(syncLockKey);
  }
};

const getMonthName = (monthNumber: string): string => {
  const months = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];
  return months[parseInt(monthNumber) - 1] || "Month";
};
const exportToExcelWithFullStyle = (): void => { 

  // 🔐 RBAC: exporting data requires at least READ access
  if (!canRead()) {
    toast("error", "❌ You don't have permission to export data");
    return;
  }
  const selectedYear = filterYear.value;
  const selectedMonth = filterMonth.value;
  
  if (!selectedYear || !selectedMonth) {
    toast("error", "Please select year and month to export");
    return;
  }

  const exportData = filteredData.length > 0 ? filteredData : forecasts;
  
  if (exportData.length === 0) {
    toast("error", "No data available to export");
    return;
  }

  try {
    const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);
     
    const excelData = exportData.map(forecast => {
      const projectTitle = forecast.Project?.Title || "-";
      const proj = projects.find(p => p.Title === projectTitle);
      
      const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
      const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
      
      const oblValue = forecast.OBL || 0;
      
      const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
      const monthName = getMonthName(selectedMonth);
      
      const selectedMonthForecastData = forecastData[selectedMonthKey] as ForecastMonthData || { plan: 0, actual: 0, forecast: 0 };
      const selectedMonthPlan = selectedMonthForecastData.plan || 0;
      const selectedMonthCurrent = selectedMonthForecastData.actual || 0;
      const selectedMonthActual = revenueActualData[selectedMonthKey] || 0;
      // Further = forecast value from backend
      const selectedMonthFurther = selectedMonthForecastData.forecast || 0;
      // Forecast = Curr. Plan + Further
      const selectedMonthForecast = selectedMonthCurrent + selectedMonthFurther;

      const rowData: any = {
        "Profit Centre": forecast.ProfitCentre || "-",
        "SO/Project Ref": proj?.ProjectReference || "-",
        "Customer Name": proj?.CustomerName || "-",
        "Project Name": projectTitle,
       "Group Head": getGroupHeadDisplayName(proj?.GroupHead1),
        "Project Manager": getProjectManagerDisplayName(proj?.ProjectManager),
        "OBL": oblValue,
        // Order: Prev. Plan → Curr. Plan → Actual → Further → Forecast
        [`${monthName} ${selectedYear} (Prev. Plan)`]: selectedMonthPlan,
        [`${monthName} ${selectedYear} (Curr. Plan)`]: selectedMonthCurrent,
        [`${monthName} ${selectedYear} (Actual)`]: selectedMonthActual,
        [`${monthName} ${selectedYear} (Further)`]: selectedMonthFurther,  // backend forecast field
        [`${monthName} ${selectedYear} (Forecast)`]: selectedMonthForecast, // Curr. Plan + Further
      };
      
      allMonths.forEach(monthKey => {
        const [year, month] = monthKey.split('-');
        const monthLabel = getMonthName(month);
        const displayLabel = `${monthLabel} ${year}`;
        
        const forecastMonthData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
        const planValue = forecastMonthData.plan || 0;
        const currentValue = forecastMonthData.actual || 0;
        
        rowData[`${displayLabel} (Prev. Plan)`] = planValue;
        rowData[`${displayLabel} (Curr. Plan)`] = currentValue;
      });
      
      lastFiveYears.forEach(year => {
        const yearData = forecastData.yearly?.[year] || { plan: 0, actual: 0 };
        rowData[`${year} (Prev. Plan)`] = yearData.plan;
        rowData[`${year} (Curr. Plan)`] = yearData.actual;
      });     
      
      const totalMonthlyCurrent = Object.keys(forecastData)
        .filter(key => key.match(/^\d{4}-\d{2}$/))
        .reduce((sum, monthKey) => {
          const monthData = forecastData[monthKey] as ForecastMonthData;
          return sum + (monthData?.actual || 0);
        }, 0);

      const totalYearlyCurrent = forecastData.yearly ? 
        Object.values(forecastData.yearly).reduce((sum, yearData) => {
          return sum + (yearData.actual || 0);
        }, 0) : 0;

      rowData["Balance"] = oblValue - (totalMonthlyCurrent + totalYearlyCurrent);
      
      return rowData;
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData, { header: headers });

    const headerStyle = {
      fill: {
        fgColor: { rgb: "FF0078D4" }
      },
      font: {
        name: "Arial",
        sz: 10,
        bold: true,
        color: { rgb: "FFFFFFFF" }
      },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
      border: {
        top: { style: "thin", color: { rgb: "FF000000" } },
        left: { style: "thin", color: { rgb: "FF000000" } },
        bottom: { style: "thin", color: { rgb: "FF000000" } },
        right: { style: "thin", color: { rgb: "FF000000" } }
      }
    };

    const dataStyle = {
      font: {
        name: "Arial",
        sz: 10
      },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
      border: {
        top: { style: "thin", color: { rgb: "FFD0D0D0" } },
        left: { style: "thin", color: { rgb: "FFD0D0D0" } },
        bottom: { style: "thin", color: { rgb: "FFD0D0D0" } },
        right: { style: "thin", color: { rgb: "FFD0D0D0" } }
      }
    };

    const numberStyle = {
      ...dataStyle,
      alignment: {
        horizontal: "right",
        vertical: "center"
      }
    };

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      
      if (!ws[cellAddress].s) {
        ws[cellAddress].s = {};
      }
      ws[cellAddress].s = headerStyle;
    }
    
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;
        
        const isNumberColumn = col >= 6;
        const styleToApply = isNumberColumn ? numberStyle : dataStyle;
        
        if (!ws[cellAddress].s) {
          ws[cellAddress].s = {};
        }
        ws[cellAddress].s = styleToApply;
      }
    }

    const colWidths = headers.map((header: string) => {
      if (header.includes('Project Name')) return { wch: 25 };
      if (header.includes('Customer Name')) return { wch: 20 };
      if (header.includes('Profit Centre') || header.includes('SO/Project Ref')) return { wch: 12 };
      if (header.includes('Group Head') || header.includes('Project Manager')) return { wch: 15 };
      if (header.includes('OBL') || header.includes('Balance')) return { wch: 12 };
      if (header.includes('(Prev. Plan)') || header.includes('(Curr. Plan)') || header.includes('(Further)') || header.includes('(Actual)') || header.includes('(Forecast)')) return { wch: 10 };
      return { wch: 12 };
    });
    ws['!cols'] = colWidths;

    ws['!freeze'] = { x: 0, y: 1 };

    XLSX.utils.book_append_sheet(wb, ws, "Revenue Forecast");

    const monthName = getMonthName(selectedMonth);
    const fileName = `Revenue_Forecast_${monthName}_${selectedYear}_Filtered.xlsx`;
    
    XLSX.writeFile(wb, fileName);
     
    toast("success", `Exported ${exportData.length} filtered records for ${monthName} ${selectedYear}`);
    
  } catch (error) { 
    toast("error", "Failed to export data to Excel");
  }
};

const generateHeadersStructure = (selectedYear: string, selectedMonth: string, exportData: RevenueForecastItem[]) => {
  const baseHeaders = [
    "Profit Centre", 
    "SO/Project Ref", 
    "Customer Name", 
    "Project Name", 
    "Group Head", 
    "Project Manager", 
    "OBL",
  ];

  const selectedMonthName = getMonthName(selectedMonth);
  // Order: Prev. Plan → Curr. Plan → Actual → Further → Forecast
  const selectedMonthHeaders = [
    `${selectedMonthName} ${selectedYear} (Prev. Plan)`,
    `${selectedMonthName} ${selectedYear} (Curr. Plan)`,
    `${selectedMonthName} ${selectedYear} (Actual)`,
    `${selectedMonthName} ${selectedYear} (Further)`,   // backend forecast field
    `${selectedMonthName} ${selectedYear} (Forecast)`   // Curr. Plan + Further
  ];

  const allMonthKeys = new Set<string>();
  
  exportData.forEach(forecast => {
    const forecastData = safeJsonParse(forecast.ForecastData, {});
    Object.keys(forecastData).forEach(key => {
      if (key && key.match(/^\d{4}-\d{2}$/)) {
        allMonthKeys.add(key);
      }
    });
  });

  const sortedMonthKeys = Array.from(allMonthKeys).sort();
  const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
  const selectedIndex = sortedMonthKeys.indexOf(selectedMonthKey);
  
  const startIndex = selectedIndex >= 0 ? selectedIndex + 1 : 0;
  const next17Months = sortedMonthKeys.slice(startIndex, startIndex + 17);

  const monthlyHeaders: string[] = [];
  next17Months.forEach(monthKey => {
    const [year, month] = monthKey.split('-');
    const monthLabel = getMonthName(month);
    const displayLabel = `${monthLabel} ${year}`;
    monthlyHeaders.push(`${displayLabel} (Prev. Plan)`, `${displayLabel} (Curr. Plan)`);
  });

  const lastFiveYears = getLastFiveYears(exportData);
  const yearlyHeaders: string[] = [];
  lastFiveYears.forEach(year => {
    yearlyHeaders.push(`${year} (Prev. Plan)`, `${year} (Curr. Plan)`);
  });

  const finalHeaders = [
    ...baseHeaders,
    ...selectedMonthHeaders,
    ...monthlyHeaders,
    ...yearlyHeaders,
    "Balance"
  ];
  
  return { 
    headers: finalHeaders, 
    allMonths: next17Months,
    lastFiveYears
  };
};

const getLastFiveYears = (data: RevenueForecastItem[]): string[] => {
  const years = new Set<string>();
  
  data.forEach(forecast => {
    const forecastData = safeJsonParse(forecast.ForecastData, {});
    if (forecastData.yearly) {
      Object.keys(forecastData.yearly).forEach(year => {
        years.add(year);
      });
    }
  });
  
  const sortedYears = Array.from(years).sort();
  return sortedYears.slice(-5);
};
const exportDetailedItemsToExcel = (): void => { 

  // 🔐 RBAC: exporting data requires at least READ access
  if (!canRead()) {
    toast("error", "❌ You don't have permission to export data");
    return;
  }
  const selectedYear = filterYear.value;
  const selectedMonth = filterMonth.value;
  
  if (!selectedYear || !selectedMonth) {
    toast("error", "Please select year and month to export");
    return;
  }

  const exportData = filteredData.length > 0 ? filteredData : forecasts;
  
  if (exportData.length === 0) {
    toast("error", "No data available to export");
    return;
  }

  try {
    const currentDate = new Date();
    const detailedMonths: string[] = [];
    for (let i = 0; i < 3; i++) {
      const monthDate = new Date(currentDate);
      monthDate.setMonth(monthDate.getMonth() + i);
      const monthKey = formatMonthKey(monthDate);
      if (monthKey) {
        detailedMonths.push(monthKey);
      }
    }

    const nextFifteenMonths: string[] = [];
    for (let i = 3; i < 18; i++) {
      const monthDate = new Date(currentDate);
      monthDate.setMonth(monthDate.getMonth() + i);
      const monthKey = formatMonthKey(monthDate);
      if (monthKey) {
        nextFifteenMonths.push(monthKey);
      }
    }

    const lastFiveYears = getLastFiveYearsFromForecastData(exportData);

    const selectedMonthName = getMonthName(selectedMonth) || selectedMonth;
    const currentMonthName = getMonthName(detailedMonths[0]?.split('-')[1] || '') || detailedMonths[0]?.split('-')[1] || '';
    const nextMonth1Name = getMonthName(detailedMonths[1]?.split('-')[1] || '') || detailedMonths[1]?.split('-')[1] || '';
    const nextMonth2Name = getMonthName(detailedMonths[2]?.split('-')[1] || '') || detailedMonths[2]?.split('-')[1] || '';
    const currentMonthYear = detailedMonths[0]?.split('-')[0] || '';
    const nextMonth1Year = detailedMonths[1]?.split('-')[0] || '';
    const nextMonth2Year = detailedMonths[2]?.split('-')[0] || '';

    const excelData: any[][] = [];

    // First row headers - Order: Prev. Plan → Curr. Plan → Actual → Further → Forecast
    const firstRowHeaders = [
      "Profit Centre",
      "SO/Project Ref", 
      "Project Name", 
      "Customer Name",       
      "Group Head", 
      "Project Manager", 
      "OBL",
      
      `${selectedMonthName} ${selectedYear} (Prev. Plan)`,
      `${selectedMonthName} ${selectedYear} (Curr. Plan)`,
      `${selectedMonthName} ${selectedYear} (Actual)`,
      `${selectedMonthName} ${selectedYear} (Further)`,   // backend forecast field
      `${selectedMonthName} ${selectedYear} (Forecast)`,  // Curr. Plan + Further
      
      `${currentMonthName} ${currentMonthYear}`.toUpperCase(),
      "", "", "", "", "", "",
      
      `${nextMonth1Name} ${nextMonth1Year} (Prev. Plan)`,
      `${nextMonth1Name} ${nextMonth1Year} (Curr. Plan)`,
      
      `${nextMonth1Name} ${nextMonth1Year}`.toUpperCase(),
      "", "", "", "",
      
      `${nextMonth2Name} ${nextMonth2Year} (Prev. Plan)`,
      `${nextMonth2Name} ${nextMonth2Year} (Curr. Plan)`,
      
      `${nextMonth2Name} ${nextMonth2Year}`.toUpperCase(),
      "", "", "", "",
    ];

    nextFifteenMonths.forEach(monthKey => {
      const [year, month] = monthKey.split('-');
      const monthLabel = getMonthName(month) || month;
      firstRowHeaders.push(`${monthLabel} ${year} (Prev. Plan)`, `${monthLabel} ${year} (Curr. Plan)`);
    });

    lastFiveYears.forEach(year => {
      firstRowHeaders.push(`${year} (Prev. Plan)`, `${year} (Curr. Plan)`);
    });

    firstRowHeaders.push("Balance");

    // Second row headers - Order: Item Description → Prev. Plan → Curr. Plan → Actual → Further → Forecast
    const secondRowHeaders = [
      "", "", "", "", "", "", "",
      "", "", "", "", "",
      "Item", "Prev. Plan", "Curr. Plan", "Actual", "Further", "Forecast",
      "", "",
      "Item", "Prev. Plan", "Curr. Plan", "Actual", "Further", "Forecast",
      "", "",
      "Item", "Prev. Plan", "Curr. Plan", "Actual", "Further", "Forecast",
    ];

    const remainingColumns = firstRowHeaders.length - secondRowHeaders.length;
    for (let i = 0; i < remainingColumns; i++) {
      secondRowHeaders.push("");
    }

    excelData.push(firstRowHeaders);
    excelData.push(secondRowHeaders);

    exportData.forEach(forecast => {
      const projectTitle = forecast.Project?.Title || "-";
      const proj = projects.find(p => p.Title === projectTitle);
      
      const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
      const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
      
      const oblValue = forecast.OBL || 0;

      const detailedMonthsData: any[] = [];
      for (let i = 0; i < 3; i++) {
        const monthKey = detailedMonths[i];
        if (monthKey) {
          const monthForecastData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0, forecast: 0, items: createEmptyItems() };
          const monthActualData = revenueActualData[monthKey] || 0;
          // Further = forecast value from backend
          const further = monthForecastData.forecast || 0;
          // Forecast = Curr. Plan + Further
          const totalForecast = (monthForecastData.actual || 0) + further;
          detailedMonthsData.push({
            plan: monthForecastData.plan || 0,
            current: monthForecastData.actual || 0,
            actual: monthActualData,
            further: further,
            total: totalForecast,
            items: monthForecastData.items || createEmptyItems()
          });
        } else {
          detailedMonthsData.push({
            plan: 0,
            current: 0,
            actual: 0,
            further: 0,
            total: 0,
            items: createEmptyItems()
          });
        }
      }

      const totalMonthlyCurrent = Object.keys(forecastData)
        .filter(key => key.match(/^\d{4}-\d{2}$/))
        .reduce((sum, monthKey) => {
          const monthData = forecastData[monthKey] as ForecastMonthData;
          return sum + (monthData?.actual || 0);
        }, 0);

      const totalYearlyCurrent = forecastData.yearly ? 
        Object.values(forecastData.yearly).reduce((sum: number, yearData: any) => {
          return sum + (yearData.actual || 0);
        }, 0) : 0;

      const balance = oblValue - (totalMonthlyCurrent + totalYearlyCurrent);

      for (let itemNum = 1; itemNum <= 10; itemNum++) {
        const itemKey = itemNum.toString();
        
        const rowData: any[] = [
        forecast.ProfitCentre || "-",
        proj?.ProjectReference || "-",
        projectTitle,
        proj?.CustomerName || "-",
        
        // ✅ GroupHead1 - Person/Group type
        getGroupHeadDisplayName(proj?.GroupHead1),
        // ✅ ProjectManager - Person/Group type
        getProjectManagerDisplayName(proj?.ProjectManager),
        oblValue,
  
  detailedMonthsData[0]?.plan || 0,
  detailedMonthsData[0]?.current || 0,
  detailedMonthsData[0]?.actual || 0,
  detailedMonthsData[0]?.further || 0,   // backend forecast field
  detailedMonthsData[0]?.total || 0,      // Curr. Plan + Further
];

        const currentMonthItems = detailedMonthsData[0]?.items || createEmptyItems();
        const currentMonthItem = currentMonthItems[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
        // Further = forecast value from backend
        const furtherValue = currentMonthItem.forecast || 0;
        // Forecast = Curr. Plan + Further
        const itemTotal = (currentMonthItem.actual || 0) + furtherValue;
        rowData.push(
          currentMonthItem.name || `Item ${itemNum}`,
          currentMonthItem.plan,
          currentMonthItem.actual,
          currentMonthItem.forecast || 0,  // Further (backend forecast field)
          furtherValue,                    // Further (same value)
          itemTotal                       // Forecast = Curr. Plan + Further
        );

        rowData.push(
          detailedMonthsData[1]?.plan || 0,
          detailedMonthsData[1]?.current || 0
        );

        const nextMonth1Items = detailedMonthsData[1]?.items || createEmptyItems();
        const nextMonth1Item = nextMonth1Items[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0 };
        const nextMonth1Further = nextMonth1Item.forecast || 0;
        const nextMonth1Total = (nextMonth1Item.actual || 0) + nextMonth1Further;
        rowData.push(
          nextMonth1Item.name || `Item ${itemNum}`,
          nextMonth1Item.plan,
          nextMonth1Item.actual,
          nextMonth1Item.forecast || 0,  // Further (backend forecast field)
          nextMonth1Further,              // Further (same value)
          nextMonth1Total                 // Forecast = Curr. Plan + Further
        );

        rowData.push(
          detailedMonthsData[2]?.plan || 0,
          detailedMonthsData[2]?.current || 0
        );

        const nextMonth2Items = detailedMonthsData[2]?.items || createEmptyItems();
        const nextMonth2Item = nextMonth2Items[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0 };
        const nextMonth2Further = nextMonth2Item.forecast || 0;
        const nextMonth2Total = (nextMonth2Item.actual || 0) + nextMonth2Further;
        rowData.push(
          nextMonth2Item.name || `Item ${itemNum}`,
          nextMonth2Item.plan,
          nextMonth2Item.actual,
          nextMonth2Item.forecast || 0,  // Further (backend forecast field)
          nextMonth2Further,              // Further (same value)
          nextMonth2Total                 // Forecast = Curr. Plan + Further
        );

        nextFifteenMonths.forEach(monthKey => {
          const monthForecastData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
          rowData.push(
            monthForecastData.plan || 0,
            monthForecastData.actual || 0
          );
        });

        lastFiveYears.forEach(year => {
          const yearData = forecastData.yearly?.[year] || { plan: 0, actual: 0 };
          rowData.push(
            yearData.plan,
            yearData.actual
          );
        });

        if (itemNum === 1) {
          rowData.push(balance);
        } else {
          rowData.push("");
        }

        excelData.push(rowData);
      }
    });

    if (excelData.length <= 2) {
      toast("info", "No detailed items data found for export");
      return;
    }
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(excelData);

    const mainHeaderStyle = {
      fill: { fgColor: { rgb: "FF0078D4" } },
      font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top: { style: "thin", color: { rgb: "FF000000" } },
        left: { style: "thin", color: { rgb: "FF000000" } },
        bottom: { style: "thin", color: { rgb: "FF000000" } },
        right: { style: "thin", color: { rgb: "FF000000" } }
      }
    };

    const subHeaderStyle = {
      fill: { fgColor: { rgb: "FF0056B3" } },
      font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top: { style: "thin", color: { rgb: "FF000000" } },
        left: { style: "thin", color: { rgb: "FF000000" } },
        bottom: { style: "thin", color: { rgb: "FF000000" } },
        right: { style: "thin", color: { rgb: "FF000000" } }
      }
    };

    const dataStyle = {
      font: { name: "Arial", sz: 10 },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top: { style: "thin", color: { rgb: "FFD0D0D0" } },
        left: { style: "thin", color: { rgb: "FFD0D0D0" } },
        bottom: { style: "thin", color: { rgb: "FFD0D0D0" } },
        right: { style: "thin", color: { rgb: "FFD0D0D0" } }
      }
    };

    const numberStyle = {
      ...dataStyle,
      alignment: { horizontal: "right", vertical: "center" }
    };

    const textStyle = {
      ...dataStyle,
      alignment: { horizontal: "left", vertical: "center" }
    };

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s = mainHeaderStyle;
    }
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 1, c: col });
      if (!ws[cellAddress]) continue;
      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s = subHeaderStyle;
    }
    
    for (let row = range.s.r + 2; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;
        
        let styleToApply = dataStyle;
        
        if (col === 0 || col === 1 || col === 2 ||
            col === 12 || col === 21 || col === 30) {
          styleToApply = textStyle;
        } 
        else if (col === 6 ||
                (col >= 7 && col <= 11) ||
                col === 13 || col === 14 || col === 15 || col === 16 || col === 17 || col === 18 || col === 19 || col === 20 ||
                col === 22 || col === 23 ||
                col === 31 || col === 32 ||
                (col >= 33)
        ) {
          styleToApply = numberStyle;
        }
        
        if (!ws[cellAddress].s) {
          ws[cellAddress].s = {};
        }
        ws[cellAddress].s = styleToApply;
      }
    }

    // Updated colWidths - Order: Prev. Plan → Curr. Plan → Actual → Further → Forecast
    const colWidths = [
      { wch: 12 },  // SO/Project Ref
      { wch: 25 },  // Project Name
      { wch: 20 },  // Customer Name
      { wch: 12 },  // Profit Centre
      { wch: 15 },  // Group Head
      { wch: 15 },  // Project Manager
      { wch: 10 },  // OBL
      { wch: 12 },  // Selected Month Prev. Plan
      { wch: 12 },  // Selected Month Curr. Plan
      { wch: 12 },  // Selected Month Actual
      { wch: 12 },  // Selected Month Further
      { wch: 12 },  // Selected Month Forecast
      { wch: 20 },  // Current Month Header
      { wch: 15 },  // Item Description
      { wch: 10 },  // Prev. Plan
      { wch: 10 },  // Curr. Plan
      { wch: 10 },  // Actual
      { wch: 10 },  // Further (backend forecast)
      { wch: 10 },  // Forecast (Curr. Plan + Further)
      { wch: 12 },  // Next Month 1 Prev. Plan
      { wch: 12 },  // Next Month 1 Curr. Plan
      { wch: 20 },  // Next Month 1 Header
      { wch: 15 },  // Item Description
      { wch: 10 },  // Prev. Plan
      { wch: 10 },  // Curr. Plan
      { wch: 10 },  // Actual
      { wch: 10 },  // Further (backend forecast)
      { wch: 10 },  // Forecast (Curr. Plan + Further)
      { wch: 12 },  // Next Month 2 Prev. Plan
      { wch: 12 },  // Next Month 2 Curr. Plan
      { wch: 20 },  // Next Month 2 Header
      { wch: 15 },  // Item Description
      { wch: 10 },  // Prev. Plan
      { wch: 10 },  // Curr. Plan
      { wch: 10 },  // Actual
      { wch: 10 },  // Further (backend forecast)
      { wch: 10 },  // Forecast (Curr. Plan + Further)
    ];

    for (let i = 0; i < nextFifteenMonths.length * 2; i++) {
      colWidths.push({ wch: 12 });
    }

    for (let i = 0; i < lastFiveYears.length * 2; i++) {
      colWidths.push({ wch: 12 });
    }

    colWidths.push({ wch: 12 }); // Balance

    ws['!cols'] = colWidths;

    if (!ws['!merges']) {
      ws['!merges'] = [];
    }

    // Merge headers for Current Month (columns 12-20)
    ws['!merges'].push({ s: { r: 0, c: 12 }, e: { r: 0, c: 20 } });
    // Merge headers for Next Month 1 (columns 22-29)
    ws['!merges'].push({ s: { r: 0, c: 22 }, e: { r: 0, c: 29 } });
    // Merge headers for Next Month 2 (columns 31-38)
    ws['!merges'].push({ s: { r: 0, c: 31 }, e: { r: 0, c: 38 } });

    // Merge rows for each project (10 items per project)
    let currentRow = 2;
    while (currentRow < excelData.length) {
      const endRow = Math.min(currentRow + 9, excelData.length - 1);
      
      // Merge project info columns (0-6) for all 10 items
      for (let col = 0; col < 7; col++) {
        if (currentRow < endRow && ws['!merges']) {
          ws['!merges'].push({ s: { r: currentRow, c: col }, e: { r: endRow, c: col } });
        }
      }
      
      // Merge numeric columns for the 10 items
      const totalColumnsToMerge = [
        7, 8, 9, 10, 11,  // Selected Month values (Prev Plan, Curr Plan, Actual, Further, Forecast)
        21, 22,           // Next Month 1 Prev Plan, Curr Plan
        30, 31,           // Next Month 2 Prev Plan, Curr Plan
      ];
      
      const monthlyStartCol = 39;
      for (let col = monthlyStartCol; col < firstRowHeaders.length; col++) {
        totalColumnsToMerge.push(col);
      }
      
      totalColumnsToMerge.forEach(col => {
        if (currentRow < endRow && ws['!merges']) {
          ws['!merges'].push({ s: { r: currentRow, c: col }, e: { r: endRow, c: col } });
        }
      });
      
      currentRow = endRow + 1;
    }

    // Remove empty cells from merged areas
    for (let col = 13; col <= 20; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      delete ws[cellAddress];
    }
    for (let col = 23; col <= 29; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      delete ws[cellAddress];
    }
    for (let col = 32; col <= 38; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      delete ws[cellAddress];
    }

    ws['!freeze'] = { x: 7, y: 2 };

    XLSX.utils.book_append_sheet(wb, ws, "Detailed Items");

    const monthName = getMonthName(selectedMonth) || selectedMonth;
    const fileName = `Revenue_Forecast_Detailed_Items_${monthName}_${selectedYear}.xlsx`;
    
    XLSX.writeFile(wb, fileName); 
    toast("success", `Exported ${excelData.length - 2} detailed item records`);
    
  } catch (error) { 
    toast("error", "Failed to export detailed items data to Excel");
  }
};

const getLastFiveYearsFromForecastData = (data: RevenueForecastItem[]): string[] => {
  const years = new Set<string>();
  
  data.forEach(forecast => {
    const forecastData = safeJsonParse(forecast.ForecastData, {});
    if (forecastData.yearly) {
      Object.keys(forecastData.yearly).forEach(year => {
        years.add(year);
      });
    }
  });
  
  const sortedYears = Array.from(years).sort();
  return sortedYears.slice(-5);
};
 const exportHistoricalToExcel = (): void => {

  // 🔐 RBAC: exporting data requires at least READ access
  if (!canRead()) {
    toast("error", "❌ You don't have permission to export data");
    return;
  }

  // ✅ Same validation as the other two export buttons — a year and month
  // must be selected in the Filters popup before exporting.
  const selectedYear = filterYear.value;
  const selectedMonth = filterMonth.value;

  if (!selectedYear || !selectedMonth) {
    toast("error", "Please select year and month to export");
    return;
  }

  const exportData = filteredData.length > 0 ? filteredData : forecasts;

  if (exportData.length === 0) {
    toast("error", "No data available to export");
    return;
  }

  try {
    // Get all 36 historical months + current month
    const allMonths = generateIdealMonthStructure();
    const historicalMonths = allMonths.slice(0, 36); // All 36 historical months
    const currentMonthKey = formatMonthKey(new Date());

    // Build headers
    const baseHeaders = [
      "Profit Centre",
      "SO/Project Ref",
      "Customer Name",
      "Project Name",
      "Group Head",
      "Project Manager",
      "OBL"
    ];

    // Historical month headers (36 months × 2 = 72 columns)
    const historicalHeaders: string[] = [];
    historicalMonths.forEach(monthKey => {
      const [year, month] = monthKey.split('-');
      const monthLabel = getMonthName(month) || month;
      const displayLabel = `${monthLabel} ${year}`;
      historicalHeaders.push(`${displayLabel} (Curr. Plan)`, `${displayLabel} (Revenue Actual)`);
    });

    // Current month headers — 5 columns: Prev. Plan, Curr. Plan, Actual, Further, Forecast
    const currentMonthHeaders = [
      `${getMonthName(currentMonthKey.split('-')[1])} ${currentMonthKey.split('-')[0]} (Prev. Plan)`,
      `${getMonthName(currentMonthKey.split('-')[1])} ${currentMonthKey.split('-')[0]} (Curr. Plan)`,
      `${getMonthName(currentMonthKey.split('-')[1])} ${currentMonthKey.split('-')[0]} (Actual)`,
      `${getMonthName(currentMonthKey.split('-')[1])} ${currentMonthKey.split('-')[0]} (Further)`,
      `${getMonthName(currentMonthKey.split('-')[1])} ${currentMonthKey.split('-')[0]} (Forecast)`
    ];

    const headers = [...baseHeaders, ...historicalHeaders, ...currentMonthHeaders];

    // Build data rows
    const excelData = exportData.map(forecast => {
      const projectTitle = forecast.Project?.Title || "-";
      const proj = projects.find(p => p.Title === projectTitle);

      const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
      const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
      const oblValue = forecast.OBL || 0;

      const rowData: any = {
        "Profit Centre": forecast.ProfitCentre || "-",
        "SO/Project Ref": proj?.ProjectReference || "-",
        "Customer Name": proj?.CustomerName || "-",
        "Project Name": projectTitle,
        "Group Head": getGroupHeadDisplayName(proj?.GroupHead1),
        "Project Manager": getProjectManagerDisplayName(proj?.ProjectManager),
        "OBL": oblValue
      };

      // Add historical month data (Curr. Plan and Revenue Actual)
      historicalMonths.forEach(monthKey => {
        const monthData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
        const currentValue = monthData.actual || 0;
        const actualValue = revenueActualData[monthKey] || 0;

        const [year, month] = monthKey.split('-');
        const monthLabel = getMonthName(month) || month;
        const displayLabel = `${monthLabel} ${year}`;

        rowData[`${displayLabel} (Curr. Plan)`] = currentValue;
        rowData[`${displayLabel} (Revenue Actual)`] = actualValue;
      });

      // Add current month data
      const currentMonthData = forecastData[currentMonthKey] as ForecastMonthData || { plan: 0, actual: 0, forecast: 0 };
      const currentMonthActual = revenueActualData[currentMonthKey] || 0;
      const currentMonthForecast = currentMonthData.forecast || 0;
      const totalForecast = currentMonthActual + currentMonthForecast;

      const monthName = getMonthName(currentMonthKey.split('-')[1]) || currentMonthKey.split('-')[1];
      const yearName = currentMonthKey.split('-')[0];
      const label = `${monthName} ${yearName}`;

      rowData[`${label} (Prev. Plan)`] = currentMonthData.plan || 0;
      rowData[`${label} (Curr. Plan)`] = currentMonthData.actual || 0;
      rowData[`${label} (Actual)`] = currentMonthActual;
      rowData[`${label} (Further)`] = currentMonthForecast;
      rowData[`${label} (Forecast)`] = totalForecast;

      return rowData;
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData, { header: headers });

    // Styling
    const headerStyle = {
      fill: { fgColor: { rgb: "FF0078D4" } },
      font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top: { style: "thin", color: { rgb: "FF000000" } },
        left: { style: "thin", color: { rgb: "FF000000" } },
        bottom: { style: "thin", color: { rgb: "FF000000" } },
        right: { style: "thin", color: { rgb: "FF000000" } }
      }
    };

    const dataStyle = {
      font: { name: "Arial", sz: 10 },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top: { style: "thin", color: { rgb: "FFD0D0D0" } },
        left: { style: "thin", color: { rgb: "FFD0D0D0" } },
        bottom: { style: "thin", color: { rgb: "FFD0D0D0" } },
        right: { style: "thin", color: { rgb: "FFD0D0D0" } }
      }
    };

    const numberStyle = {
      ...dataStyle,
      alignment: { horizontal: "right", vertical: "center" }
    };

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

    // Apply header style
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s = headerStyle;
    }

    // Apply data styles
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;

        // Columns 0-6 are text (Profit Centre, SO/Project Ref, Customer, Project, GH, PM, OBL),
        // everything from column 7 onwards is numeric (monthly plan/actual data).
        const isNumberColumn = col >= 7;
        const styleToApply = isNumberColumn ? numberStyle : dataStyle;

        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        ws[cellAddress].s = styleToApply;
      }
    }

    // Column widths
    const colWidths = [
      { wch: 12 },  // Profit Centre
      { wch: 15 },  // SO/Project Ref
      { wch: 20 },  // Customer Name
      { wch: 25 },  // Project Name
      { wch: 15 },  // Group Head
      { wch: 15 },  // Project Manager
      { wch: 12 },  // OBL
    ];

    // 36 historical months × 2 columns each (Curr. Plan + Revenue Actual)
    for (let i = 0; i < historicalMonths.length * 2; i++) {
      colWidths.push({ wch: 12 });
    }

    // 5 current-month columns
    for (let i = 0; i < 5; i++) {
      colWidths.push({ wch: 12 });
    }

    ws['!cols'] = colWidths;

    // Freeze the header row + the first 7 columns (project metadata)
    ws['!freeze'] = { x: 7, y: 1 };

    XLSX.utils.book_append_sheet(wb, ws, "Historical Data");

    const monthNameOut = getMonthName(selectedMonth) || selectedMonth;
    const fileName = `Revenue_Forecast_Historical_${monthNameOut}_${selectedYear}.xlsx`;

    XLSX.writeFile(wb, fileName);

    toast("success", `Exported ${exportData.length} historical records for ${monthNameOut} ${selectedYear}`);

  } catch (error) {
    console.error("❌ Historical export failed:", error);
    toast("error", "Failed to export historical data to Excel");
  }
};
 
// Pager bindings - wired ONCE (the old version cloned the buttons, which left
// the module holding detached nodes, so disabled states never showed).
const setupPaginationEventListeners = (): void => {
  if (pageSizeSelect) {
    pageSize = parseInt(pageSizeSelect.value) || 25;
  }

  firstPageBtn?.addEventListener("click", () => goToPage(1));
  prevPageBtn?.addEventListener("click", () => goToPage(currentPage - 1));
  nextPageBtn?.addEventListener("click", () => goToPage(currentPage + 1));
  lastPageBtn?.addEventListener("click", () => goToPage(totalPages));
  pageSizeSelect?.addEventListener("change", (e) => {
    changePageSize(parseInt((e.target as HTMLSelectElement).value, 10));
  });
};

// COMPLETE REPLACEMENT FOR changePageSize
const changePageSize = (newSize: number): void => {
  if (!newSize || newSize <= 0) {
    return;
  }

  // ✅ keep the FIRST visible record in view: work out its index with the
  //    OLD page size BEFORE switching, then open the page that contains it.
  const firstVisibleIndex = (currentPage - 1) * pageSize;

  pageSize = newSize;
  currentPage = Math.floor(firstVisibleIndex / pageSize) + 1;

  calculateTotalPages();

  if (currentPage > totalPages) {
    currentPage = totalPages;
  }

  updatePaginationControls();
  renderCurrentPage();

  const displayMonths = getDisplayMonths();
  const currentMonthKey = formatMonthKey(new Date());
  setTimeout(() => {
    renderTotalsHeader(displayMonths, currentMonthKey);
  }, 100);

  toast("info", `Showing ${pageSize} records per page`);
};

// COMPLETE REPLACEMENT FOR renderCurrentPage
const renderCurrentPage = (): void => {
  console.log(`📊 renderCurrentPage called - filteredData length: ${filteredData.length}, pageSize: ${pageSize}, currentPage: ${currentPage}`);
  
  if (!filteredData || filteredData.length === 0) {
    console.log("📊 No data to render");
    renderTableWithMonthColumns([]);
    return;
  }

  // Ensure currentPage is within bounds
  if (currentPage > totalPages) {
    currentPage = totalPages;
  }

  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, filteredData.length);
  const pageData = filteredData.slice(startIndex, endIndex);
  
  console.log(`📊 Rendering page ${currentPage}: items ${startIndex + 1} to ${endIndex} of ${filteredData.length}`);
  console.log(`📊 Page data length: ${pageData.length}`);
  
  // Force re-render the table
  renderTableWithMonthColumns(pageData);
};

// COMPLETE REPLACEMENT FOR updatePaginationControls
const updatePaginationControls = (): void => {
  console.log(`📊 updatePaginationControls - currentPage: ${currentPage}, totalPages: ${totalPages}, filteredData length: ${filteredData.length}`);
  
  totalPagesSpan.textContent = totalPages.toString();

  if (filteredData.length === 0) {
    startItemSpan.textContent = '0';
    endItemSpan.textContent = '0';
    totalItemsSpan.textContent = '0';
  } else {
    const startItem = ((currentPage - 1) * pageSize) + 1;
    const endItem = Math.min(currentPage * pageSize, filteredData.length);

    startItemSpan.textContent = startItem.toString();
    endItemSpan.textContent = endItem.toString();
    totalItemsSpan.textContent = filteredData.length.toString();
    
    console.log(`📊 Showing items ${startItem} to ${endItem} of ${filteredData.length}`);
  }

  // Keep dropdown synchronized with actual pageSize - IMPORTANT!
  if (pageSizeSelect) {
    pageSizeSelect.value = pageSize.toString();
    console.log(`📊 Set pageSizeSelect value to: ${pageSize}`);
  }

  firstPageBtn.disabled = currentPage === 1 || filteredData.length === 0;
  prevPageBtn.disabled = currentPage === 1 || filteredData.length === 0;
  nextPageBtn.disabled = currentPage >= totalPages || filteredData.length === 0;
  lastPageBtn.disabled = currentPage >= totalPages || filteredData.length === 0;

  updatePageNumbers();
};

// COMPLETE REPLACEMENT FOR goToPage
const goToPage = (page: number): void => {
  if (page < 1 || page > totalPages || page === currentPage) return;
  if (filteredData.length === 0) return;

  currentPage = page;
  updatePaginationControls();
  renderCurrentPage();

  const tableSectionEl = document.querySelector('.table-section') as HTMLElement;
  tableSectionEl?.scrollIntoView({ behavior: 'smooth', block: 'start' });
};

// COMPLETE REPLACEMENT FOR initializePagination
const initializePagination = (data: RevenueForecastItem[]): void => { 
  console.log("🔍 initializePagination called with data length:", data.length);
  filteredData = data; 
  console.log("📊 filteredData set to:", filteredData.length);
  
  // Ensure pageSize is set from dropdown
  if (pageSizeSelect) {
    pageSize = parseInt(pageSizeSelect.value) || 10;
  }
  console.log(`📊 Using page size: ${pageSize}`);
  
  // Reset to page 1
  currentPage = 1;
  
  // Calculate total pages
  calculateTotalPages();
  
  // Update pagination controls
  updatePaginationControls();
  
  // Render the first page
  renderCurrentPage();
  
  const displayMonths = getDisplayMonths();
  const currentMonthKey = formatMonthKey(new Date());
  setTimeout(() => {
    renderTotalsHeader(displayMonths, currentMonthKey);
  }, 100);
};

// COMPLETE REPLACEMENT FOR calculateTotalPages
const calculateTotalPages = (): void => {
  totalPages = Math.ceil(filteredData.length / pageSize);
  if (totalPages === 0) totalPages = 1;
  console.log(`📄 Total pages: ${totalPages} (${filteredData.length} items / ${pageSize} per page)`);
};

// COMPLETE REPLACEMENT FOR updatePageNumbers
const updatePageNumbers = (): void => {
  pageNumbersDiv.innerHTML = '';
  
  if (totalPages <= 1) {
    // Only show page 1 if there's data
    if (filteredData.length > 0) {
      const pageBtn = document.createElement('button');
      pageBtn.className = `revenue-page-number revenue-active`;
      pageBtn.textContent = '1';
      pageBtn.addEventListener('click', () => goToPage(1));
      pageNumbersDiv.appendChild(pageBtn);
    }
    return;
  }
  
  const maxVisiblePages = 5;
  let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
  let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
  
  if (endPage - startPage + 1 < maxVisiblePages) {
    startPage = Math.max(1, endPage - maxVisiblePages + 1);
  }
  
  if (startPage > 1) {
    const dotsBtn = document.createElement('button');
    dotsBtn.className = 'revenue-page-number revenue-dots';
    dotsBtn.textContent = '...';
    dotsBtn.disabled = true;
    pageNumbersDiv.appendChild(dotsBtn);
  }
  
  for (let i = startPage; i <= endPage; i++) {
    const pageBtn = document.createElement('button');
    pageBtn.className = `revenue-page-number ${i === currentPage ? 'revenue-active' : ''}`;
    pageBtn.textContent = i.toString();
    pageBtn.addEventListener('click', () => goToPage(i));
    pageNumbersDiv.appendChild(pageBtn);
  }
  
  if (endPage < totalPages) {
    const dotsBtn = document.createElement('button');
    dotsBtn.className = 'revenue-page-number revenue-dots';
    dotsBtn.textContent = '...';
    dotsBtn.disabled = true;
    pageNumbersDiv.appendChild(dotsBtn);
  }
}; 
 

const calculateTotals = (
  data: RevenueForecastItem[],
  displayMonths: string[],
  currentMonthKey: string
): any => {
  const totals = {
    obl: 0,
    monthly: {} as {
      [key: string]: { plan: number; current: number; actual: number; forecast: number; total: number }
    }
  };

  // Initialize monthly totals
  displayMonths.forEach(monthKey => {
    totals.monthly[monthKey] = { plan: 0, current: 0, actual: 0, forecast: 0, total: 0 };
  });

  data.forEach(forecast => {
    totals.obl += Math.round(Number(forecast.OBL) || 0);

    const forecastData = safeJsonParse(forecast.ForecastData, {});
    const revenueActualData = safeJsonParse(forecast.RevenueActual, {});

    displayMonths.forEach(monthKey => {
      const m = forecastData[monthKey] || {};
      
      // Ensure all values are numbers
      const plan = Math.round(Number(m.plan) || 0);
      const current = Math.round(Number(m.actual) || 0);
      const actual = Math.round(Number(revenueActualData[monthKey]) || 0);
      const forecastVal = Math.round(Number(m.forecast) || 0);

      totals.monthly[monthKey].plan += plan;
      totals.monthly[monthKey].current += current;
      totals.monthly[monthKey].actual += actual;
      totals.monthly[monthKey].forecast += forecastVal;
      
      // ✅ FIX: Forecast = Actual + Further (RevenueActual + forecast field)
      totals.monthly[monthKey].total = totals.monthly[monthKey].actual + totals.monthly[monthKey].forecast;
    });
  });

  return totals;
};
 
const setFormReadOnly = (isFrozen: boolean): void => {
    // Define which selectors should be forced readonly when frozen
    const forcedReadonlySelectors = [
        '.revenue-simple-month-rev .revenue-main-current-input',
        '.revenue-yearly-current-input-rev',
        '.revenue-detailed-month-box .revenue-item-current-input',
        // '.revenue-detailed-month-box .revenue-item-description-input',
        '#remarksInternal',
        '#remarksSprintCell'
    ];
    
    const inputs = document.querySelectorAll<HTMLInputElement | HTMLTextAreaElement>(forcedReadonlySelectors.join(','));
    
    inputs.forEach(input => {
        if (isFrozen) {
            input.setAttribute('readonly', 'true');
            input.style.backgroundColor = '#f5f5f5';
            input.style.color = '#6c757d';
            input.style.cursor = 'not-allowed';
            input.style.opacity = '0.7';
            input.style.pointerEvents = 'none';
        } else {
            input.removeAttribute('readonly');
            input.style.backgroundColor = '';
            input.style.color = '';
            input.style.cursor = '';
            input.style.opacity = '1';
            input.style.pointerEvents = '';
        }
    });
    
    // ============================================================
    // MAIN-FORECAST-INPUT - ALWAYS READONLY (calculated from items)
    // ============================================================
   const mainForecastInputs =
    document.querySelectorAll<HTMLInputElement>('.revenue-main-forecast-input');

mainForecastInputs.forEach(input => {

    if (isFrozen) {
        // Frozen:
        // Further is still editable
        input.removeAttribute('readonly');
        input.disabled = false;

        input.style.backgroundColor = '#ffffff';
        input.style.color = '#000000';
        input.style.cursor = 'text';
        input.style.opacity = '1';
        input.style.pointerEvents = 'auto';
        input.style.borderColor = '#28a745';
        input.style.borderWidth = '2px';
        input.style.borderStyle = 'solid';

    } else {
        // Unfrozen:
        // Normal editable mode
        input.removeAttribute('readonly');
        input.disabled = false;

        input.style.backgroundColor = '';
        input.style.color = '';
        input.style.cursor = 'text';
        input.style.opacity = '1';
        input.style.pointerEvents = 'auto';
        input.style.borderColor = '';
        input.style.borderWidth = '';
        input.style.borderStyle = '';
    }
});
    
    // ============================================================
    // ITEM-FORECAST-INPUT - ALWAYS EDITABLE (even when frozen)
    // ============================================================
    const itemForecastInputs = document.querySelectorAll<HTMLInputElement>('.revenue-item-forecast-input');
    
    itemForecastInputs.forEach(input => {
        // Always editable
        input.removeAttribute('readonly');
        input.disabled = false;
        input.style.backgroundColor = '#ffffff';
        input.style.color = '#000000';
        input.style.cursor = 'text';
        input.style.opacity = '1';
        input.style.pointerEvents = 'auto';
        input.style.borderColor = '#28a745';
        input.style.borderWidth = '2px';
        input.style.borderStyle = 'solid';
    });
    
    // ============================================================
    // DATE INPUTS - ALWAYS EDITABLE (even when frozen)
    // ============================================================
       // ============================================================
    // MAIN (SECTION-HEADER) DATE INPUT - READONLY (auto-filled from max item date)
    // ============================================================
    const mainDateInputs = document.querySelectorAll<HTMLInputElement>('.revenue-forecast-date-input');

    mainDateInputs.forEach(input => {
        input.setAttribute('readonly', 'true');
        input.disabled = false;
        input.style.backgroundColor = '#f5f5f5';
        input.style.color = '#000000';
        input.style.cursor = 'default';
        input.style.opacity = '1';
        input.style.pointerEvents = 'none';
        input.style.borderColor = '#fff';
        input.style.borderWidth = '2px';
        input.style.borderStyle = 'solid';
    });

    // ============================================================
    // ITEM-LEVEL DATE INPUTS - ALWAYS EDITABLE (even when frozen) 
    // ============================================================
    const itemDateInputs = document.querySelectorAll<HTMLInputElement>('.revenue-item-date-input');

    itemDateInputs.forEach(input => {
        input.removeAttribute('readonly');
        input.disabled = false;
        input.style.backgroundColor = '#ffffff';
        input.style.color = '#000000';
        input.style.cursor = 'text';
        input.style.opacity = '1';
        input.style.pointerEvents = 'auto';
        input.style.borderColor = '#fff';
        input.style.borderWidth = '2px';
        input.style.borderStyle = 'solid';
    });
    
    // ============================================================
    // ⚠️ DO NOT HIDE OR DISABLE THE SAVE BUTTON HERE
    // The save button visibility is controlled in editForecast
    // ============================================================
    
    // Handle OBL input separately
    const oblInput = document.querySelector('#oblLastMonth') as HTMLInputElement;
    if (oblInput) {
        const isExistingProject = editingId !== null;
        if (isExistingProject || isFrozen) {
            oblInput.setAttribute('readonly', 'true');
            oblInput.style.backgroundColor = '#f5f5f5';
            oblInput.style.cursor = 'not-allowed';
        } else {
            oblInput.removeAttribute('readonly');
            oblInput.style.backgroundColor = '';
            oblInput.style.cursor = '';
        }
    }
    
    // Update the freeze checkbox state with toast
    // Toast is now shown in editForecast to avoid duplication
};
 

const renderTotalsHeader = (displayMonths: string[], currentMonthKey: string): void => {
  const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;
  if (!totalsHeader) return;
  
  totalsHeader.innerHTML = '';
  
  const allData = forecasts;
  const filtered = filteredData.length ? filteredData : allData;
  
  // Calculate totals for ALL data
  const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
  // Calculate totals for FILTERED data
  const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);
  
  // Check if filters are applied
  const hasFilters = filteredData.length > 0 && filteredData.length !== allData.length;
  
  // Calculate total columns
  const baseColumns = 7;
  const monthColumns = displayMonths.length * 2;
  const currentMonthExtraColumns = displayMonths.includes(currentMonthKey) ? 3 : 0;
  const totalColspan = baseColumns + monthColumns + currentMonthExtraColumns;
  
  const buildTotalsRow = (label: string, totals: any, css: string, showData: boolean = true) => {
    if (!showData) return '';
    
    // Build month cells
    const monthCells = displayMonths.map((m, index) => {
      const isCurrent = m === currentMonthKey;
      const t = totals.monthly[m] || { plan: 0, current: 0, actual: 0, forecast: 0, total: 0 };
      
      if (isCurrent) {
        // ✅ Current month: Prev. Plan, Curr. Plan, Actual, Further, Forecast
        // Forecast = Actual + Further (calculated in calculateTotals)
        return `
          <td class="revenue-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle; background-color: #e8f5e9;">${t.actual.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.forecast.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col revenue-total-value" style="min-width:60px; text-align:center; vertical-align:middle; font-weight:bold; background-color: #f0f8ff;">${t.total.toLocaleString('en-US')}</td>
        `;
      } else {
        return `
          <td class="revenue-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-US')}</td>
        `;
      }
    }).join('');
    
    return `
      <tr class="${css}">
        <td class="revenue-freeze-col" style="text-align:center; vertical-align:middle; min-width:40px; max-width:40px; width:40px;">${css === 'revenue-all-projects-total' ? '★' : '●'}</td>
        <td class="revenue-project-ref-col" style="text-align:center; vertical-align:middle; min-width:120px; max-width:120px; width:120px;"></td>
        <td class="revenue-customer-col" style="text-align:center; vertical-align:middle; min-width:130px; max-width:130px; width:130px;"></td>
        <td class="revenue-project-col" style="text-align:left; vertical-align:middle; min-width:180px; max-width:180px; width:180px; font-weight:bold; background-color:${css === 'revenue-all-projects-total' ? '#d1ecf1' : '#d4edda'};">
          ${label}
        </td>
        <td class="revenue-gh-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
        <td class="revenue-pm-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
        <td class="revenue-obl-col" style="text-align:right; vertical-align:middle; min-width:90px; max-width:90px; width:90px; font-weight:bold;">${totals.obl.toLocaleString('en-US')}</td>
        ${monthCells}
      </tr>
    `;
  };
  
  // Build the totals header
  let html = `
    <tr>
      <th colspan="${totalColspan}" style="background:#f8f9fa; padding:8px; text-align:center; font-size:13px; font-weight:bold; color:#495057; border:1px solid #dee2e6;">
        📊 Revenue Totals
      </th>
    </tr>
  `;
  
  // Always show ALL totals
  if (allData.length > 0) {
    html += buildTotalsRow(`All Projects (${allData.length})`, allTotals, 'revenue-all-projects-total', true);
  }
  
  // Only show FILTERED totals if filters are applied AND data differs
  if (hasFilters && filtered.length > 0) {
    html += buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'revenue-filtered-total', true);
  }
  
  totalsHeader.innerHTML = html;
   applyStickyHeaderOffsets();
};
  

const applyStickyHeaderOffsets = (): void => {
  const table = document.querySelector('.revenue-month-columns-table-rev') as HTMLTableElement;
  if (!table) return;

  const headerRows = table.querySelectorAll<HTMLTableRowElement>('thead tr');
  let offset = 0;

  headerRows.forEach(row => {
    row.querySelectorAll<HTMLElement>('th, td').forEach(cell => {
      cell.style.position = 'sticky';
      cell.style.top = `${offset}px`;
    });
    offset += row.getBoundingClientRect().height;
  });
};


// Move attachEditTriggerListeners BEFORE renderTableWithMonthColumns
const attachEditTriggerListeners = (): void => {
  document.querySelectorAll('.revenue-edit-trigger').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const id = parseInt((e.currentTarget as HTMLElement).dataset.id!);
      editForecast(id);
    });
  });
};
// COMPLETE REPLACEMENT FOR renderTableWithMonthColumns
const renderTableWithMonthColumns = (data: RevenueForecastItem[]): void => {
  console.log("📊 renderTableWithMonthColumns called with data length:", data.length);
  
  if (!data || data.length === 0) {
    const displayMonths = getDisplayMonths();
    const currentMonthKey = formatMonthKey(new Date());
    const totalCols = 7 + (displayMonths.length * 2) + (displayMonths.includes(currentMonthKey) ? 3 : 0);
    tableBody.innerHTML = `<tr><td colspan="${totalCols}" style="text-align:center;padding:20px;">No records found</td></tr>`;
    const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;
    if (totalsHeader) totalsHeader.innerHTML = '';
    return;
  }

  // FILTER OUT EMPTY/INVALID DATA - Skip records with no project title
  const validData = data.filter(forecast => {
    const projectTitle = forecast.Project?.Title;
    // Skip if project title is null, undefined, empty, or just whitespace
    if (!projectTitle || projectTitle.trim() === '') {
      console.log(`⚠️ Skipping record with empty project title:`, forecast);
      return false;
    }
    return true;
  });

  console.log(`📊 Valid data length after filtering: ${validData.length} (removed ${data.length - validData.length} empty records)`);

  // If after filtering there's no valid data, show empty state
  if (validData.length === 0) {
    const displayMonths = getDisplayMonths();
    const currentMonthKey = formatMonthKey(new Date());
    const totalCols = 7 + (displayMonths.length * 2) + (displayMonths.includes(currentMonthKey) ? 3 : 0);
    tableBody.innerHTML = `<tr><td colspan="${totalCols}" style="text-align:center;padding:20px;">No valid records found</td></tr>`;
    const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;
    if (totalsHeader) totalsHeader.innerHTML = '';
    return;
  }

  const displayMonths = getDisplayMonths();
  const currentMonthKey = formatMonthKey(new Date());
  
  renderTotalsHeader(displayMonths, currentMonthKey);
    
  // Main header - order: PR → Customer → Project → GH → PM → OBL
  const mainHeaderHTML = `
    <tr>  
      <th rowspan="2" class="revenue-freeze-col" style="text-align:center; min-width:40px; max-width:40px; width:40px;">Lock</th>
      <th rowspan="2" class="revenue-project-ref-col" style="min-width:120px; max-width:120px; width:120px;">
        <div class="revenue-sort-header-container">
          <span class="revenue-header-label">Project Reference</span>
          <div class="revenue-sort-buttons-container">
            <button class="revenue-sort-btn ${currentSortConfig.column === 'projectRef' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                    data-sort="projectRef" data-direction="asc" title="Sort Ascending">▲</button>
            <button class="revenue-sort-btn ${currentSortConfig.column === 'projectRef' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                    data-sort="projectRef" data-direction="desc" title="Sort Descending">▼</button>
          </div>
        </div>
      </th>
      <th rowspan="2" class="revenue-customer-col" style="min-width:130px; max-width:130px; width:130px;">
        <div class="revenue-sort-header-container">
          <span class="revenue-header-label">Customer</span>
          <div class="revenue-sort-buttons-container">
            <button class="revenue-sort-btn ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                    data-sort="customer" data-direction="asc" title="Sort Ascending">▲</button>
            <button class="revenue-sort-btn ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                    data-sort="customer" data-direction="desc" title="Sort Descending">▼</button>
          </div>
        </div>
      </th>
      <th rowspan="2" class="revenue-project-col" style="min-width:180px; max-width:180px; width:180px;">
        <div class="revenue-sort-header-container">
          <span class="revenue-header-label">Project</span>
          <div class="revenue-sort-buttons-container">
            <button class="revenue-sort-btn ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                    data-sort="project" data-direction="asc" title="Sort Ascending">▲</button>
            <button class="revenue-sort-btn ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                    data-sort="project" data-direction="desc" title="Sort Descending">▼</button>
          </div>
        </div>
      </th>
      <th rowspan="2" class="revenue-gh-col" style="min-width:100px; max-width:100px; width:100px;">
        <div class="revenue-sort-header-container">
          <span class="revenue-header-label">Group Head</span>
          <div class="revenue-sort-buttons-container">
            <button class="revenue-sort-btn ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                    data-sort="groupHead" data-direction="asc" title="Sort Ascending">▲</button>
            <button class="revenue-sort-btn ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                    data-sort="groupHead" data-direction="desc" title="Sort Descending">▼</button>
          </div>
        </div>
      </th>
      <th rowspan="2" class="revenue-pm-col" style="min-width:100px; max-width:100px; width:100px;">
        <div class="revenue-sort-header-container">
          <span class="revenue-header-label">Project Manager</span>
          <div class="revenue-sort-buttons-container">
            <button class="revenue-sort-btn ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                    data-sort="projectManager" data-direction="asc" title="Sort Ascending">▲</button>
            <button class="revenue-sort-btn ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                    data-sort="projectManager" data-direction="desc" title="Sort Descending">▼</button>
          </div>
        </div>
      </th>
      <th rowspan="2" class="revenue-obl-col" style="min-width:90px; max-width:90px; width:90px;">
        <div class="revenue-sort-header-container">
          <span class="revenue-header-label">OBL</span>
          <div class="revenue-sort-buttons-container">
            <button class="revenue-sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                    data-sort="obl" data-direction="asc" title="Sort Ascending">▲</button>
            <button class="revenue-sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                    data-sort="obl" data-direction="desc" title="Sort Descending">▼</button>
          </div>
        </div>
      </th>
      
      ${displayMonths.map((month, monthIndex) => {
        const isCurrentMonth = month === currentMonthKey;
        const monthDate = new Date(month + "-01");
        const monthLabel = formatMonthLabel(monthDate);
        const colspan = isCurrentMonth ? 5 : 2;
        const isCurrentDisplay = isCurrentMonth ? ' (Current Month)' : '';
        
        return `
          <th colspan="${colspan}" class="revenue-month-scroll-col revenue-month-column-header ${isCurrentMonth ? 'revenue-current-month-header' : ''}" style="min-width:${isCurrentMonth ? '110px' : '70px'};">
            <div class="revenue-sort-header-container">
              <span class="revenue-header-label">${monthLabel}${isCurrentDisplay}</span>
              <div class="revenue-sort-buttons-container">
                <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                        data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort Ascending">▲</button>
                <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                        data-sort="month" data-direction="desc" data-month-index="${monthIndex}" title="Sort Descending">▼</button>
              </div>
            </div>
          </th>
        `;
      }).join('')}
    </tr>
    
    <tr>
      ${displayMonths.map((month, monthIndex) => {
        const isCurrentMonth = month === currentMonthKey;
        
        if (isCurrentMonth) {
          return `
            <th class="revenue-month-scroll-col revenue-month-subheader revenue-current-month-header-table" style="min-width:60px;">
              <div class="revenue-sort-header-container">
                <span>Prev. Plan</span>
                <div class="revenue-sort-buttons-container">
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}">▲</button>
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="revenue-month-scroll-col revenue-month-subheader revenue-current-month-header-table" style="min-width:60px;">
              <div class="revenue-sort-header-container">
                <span>Curr. Plan</span>
                <div class="revenue-sort-buttons-container">
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}">▲</button>
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="current" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="revenue-month-scroll-col revenue-month-subheader revenue-current-month-header-table" style="min-width:60px;">
              <div class="revenue-sort-header-container">
                <span>Actual</span>
                <div class="revenue-sort-buttons-container">
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'actual' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="actual" data-month-index="${monthIndex}">▲</button>
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'actual' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="actual" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="revenue-month-scroll-col revenue-month-subheader revenue-current-month-header-table" style="min-width:60px;">
              <div class="revenue-sort-header-container">
                <span>Further</span>
                <div class="revenue-sort-buttons-container">
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'forecast' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="forecast" data-month-index="${monthIndex}">▲</button>
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'forecast' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="forecast" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="revenue-month-scroll-col revenue-month-subheader revenue-current-month-header-table" style="min-width:60px;">
              <div class="revenue-sort-header-container">
                <span>Forecast</span>
                <div class="revenue-sort-buttons-container">
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'total' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="total" data-month-index="${monthIndex}">▲</button>
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'total' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="total" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
          `;
        } else {
          return `
            <th class="revenue-month-scroll-col revenue-month-subheader" style="min-width:60px;">
              <div class="revenue-sort-header-container">
                <span>Prev. Plan</span>
                <div class="revenue-sort-buttons-container">
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}">▲</button>
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="revenue-month-scroll-col revenue-month-subheader" style="min-width:60px;">
              <div class="revenue-sort-header-container">
                <span>Curr. Plan</span>
                <div class="revenue-sort-buttons-container">
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'revenue-sort-asc revenue-active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}">▲</button>
                  <button class="revenue-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'revenue-sort-desc revenue-active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="current" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
          `;
        }
      }).join('')}
    </tr>
  `;

  monthColumnsHeader.innerHTML = mainHeaderHTML;
  let tableBodyHTML = '';

  validData.forEach((forecast, index) => {
    const projectTitle = forecast.Project?.Title || "Unknown Project";
    const proj = projects.find(p => p.Title === projectTitle);
    
    const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
    const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
    const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-US');
    
    const isFrozen = forecast.Freeze || false;
    const rowStyle = isFrozen ? 'background-color: #fff8e1;' : '';
    
    const canUserEdit = canEditProject(proj, forecast);
    const canUserFreeze = canFreeze();
    
    // Determine if checkbox should be disabled
    const isCheckboxDisabled = !canUserFreeze;

    // Build project reference link - only clickable if user can edit
    const projectRef = getSafeString(proj?.ProjectReference);
    const projectRefLink = projectRef !== "-" 
      ? (canUserEdit 
          ? `<a href="#" class="revenue-edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectRef}</a>`
          : `<span style="color:#6c757d; cursor:default;">${projectRef}</span>`)
      : '<span style="color:#999;">-</span>';

    // ✅ Project name: NO hyperlink - ONE single lock icon when the row is
    // frozen OR the user cannot edit it (previously TWO locks were rendered:
    // one inside this span and a second freezeBadge after it).
    const lockIcon = (!canUserEdit || isFrozen) ? ' 🔒' : '';
    const projectNameLink = `<span style="color:${projectTitle === 'Unknown Project' ? '#999' : '#333'}; cursor:default;">${projectTitle}${lockIcon}</span>`;

    const ghName = proj ? getGroupHeadDisplayName(proj.GroupHead1) : "-";
    const pmName = proj ? getProjectManagerDisplayName(proj.ProjectManager) : "-";
    const customerName = getSafeString(proj?.CustomerName);

    // Row cells order: # → PR → Customer → Project → GH → PM → OBL → months
    let rowHTML = `
      <tr style="${rowStyle}" data-forecast-id="${forecast.Id}"> 
        <td class="revenue-freeze-col" style="text-align:center; vertical-align:middle; min-width:40px; max-width:40px; width:40px;">
          <input type="checkbox" class="revenue-freeze-checkbox" 
       data-id="${forecast.Id}" 
       data-freeze="${isFrozen}"
       ${isFrozen ? 'checked' : ''}
       ${isCheckboxDisabled ? 'disabled' : ''}
       style="pointer-events: ${isCheckboxDisabled ? 'none' : 'auto'} !important; 
              cursor: ${isCheckboxDisabled ? 'not-allowed' : 'pointer'} !important; 
              position: relative; z-index: 100; 
              -webkit-appearance: checkbox !important; 
              appearance: checkbox !important;"
       title="${canUserFreeze ? 'Toggle freeze status' : 'Freeze permission restricted'}" />
        </td>
        <td class="revenue-project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
        <td class="revenue-customer-col" style="min-width:130px; max-width:130px; width:130px;">${customerName}</td>
        <td class="revenue-project-col" style="min-width:180px; max-width:180px; width:180px;">${projectNameLink}</td>
        <td class="revenue-gh-col" style="min-width:100px; max-width:100px; width:100px;">${ghName}</td>
        <td class="revenue-pm-col" style="min-width:100px; max-width:100px; width:100px;">${pmName}</td>
        <td class="revenue-obl-col" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
    `;

    displayMonths.forEach((monthKey, monthIndex) => {
      const monthData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0, forecast: 0 };
      const isCurrentMonth = monthKey === currentMonthKey;
      
      const planValue = monthData.plan || 0;
      const currentValue = monthData.actual || 0;
      
      if (isCurrentMonth) {
        // ✅ Forecast = RevenueActual + Further
        const revenueActualValue = revenueActualData[monthKey] || 0;  // From RevenueActual field
        const furtherValue = monthData.forecast || 0;  // Further field
        const totalForecast = revenueActualValue + furtherValue;  // Actual + Further
        
        rowHTML += `
          <td class="revenue-month-scroll-col" style="min-width:60px;">${planValue.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px;">${currentValue.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px; background-color: #e8f5e9;">${revenueActualValue.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px;">${furtherValue.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col revenue-total-cell" style="min-width:60px; font-weight: bold; background-color: #f0f8ff;">${totalForecast.toLocaleString('en-US')}</td>
        `;
      } else {
        rowHTML += `
          <td class="revenue-month-scroll-col" style="min-width:60px;">${planValue.toLocaleString('en-US')}</td>
          <td class="revenue-month-scroll-col" style="min-width:60px;">${currentValue.toLocaleString('en-US')}</td>
        `;
      }
    });

    tableBodyHTML += rowHTML + `</tr>`;
  });

  tableBody.innerHTML = tableBodyHTML;
  
  attachSortButtonListeners();
  attachFreezeEventListeners();
  attachEditTriggerListeners();
   applyStickyHeaderOffsets(); 
};

const attachSortButtonListeners = (): void => {
  document.querySelectorAll('.revenue-sort-btn[data-sort]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const button = e.currentTarget as HTMLButtonElement;
      const sortColumn = button.dataset.sort;
      const direction = button.dataset.direction as 'asc' | 'desc';
      const subColumn = button.dataset.subcolumn;
      const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;
      
      if (sortColumn && direction) {
        handleRevenueSort(sortColumn, direction, subColumn, monthIndex);
      }
    });
  });
};


const attachFreezeEventListeners = (): void => {
  console.log("🔧 Attaching freeze event listeners...");
  
  const checkboxes = document.querySelectorAll('.revenue-freeze-checkbox');
  const canUserFreeze = canFreeze();
  
  checkboxes.forEach(checkbox => {
    const cb = checkbox as HTMLInputElement;
    
    // Remove existing listeners
    cb.onchange = null;
    
    // Force proper checkbox appearance
    cb.style.pointerEvents = canUserFreeze ? 'auto' : 'none';
    cb.style.cursor = canUserFreeze ? 'pointer' : 'not-allowed';
    cb.style.opacity = canUserFreeze ? '1' : '0.6';
    cb.style.display = 'inline-block';
    cb.style.width = '18px';
    cb.style.height = '18px';
    cb.style.accentColor = '#0078d4';
    cb.style.verticalAlign = 'middle';
    cb.style.position = 'relative';
    cb.style.zIndex = '100';
    cb.style.margin = '0 auto';
    
    // Set checked state
    const isFrozen = cb.dataset.freeze === 'true';
    cb.checked = isFrozen;
    
    // Set disabled based on permission
    cb.disabled = !canUserFreeze;
    
    console.log(`✅ Checkbox ${cb.dataset.id} rendered, checked: ${cb.checked}, disabled: ${cb.disabled}`);
    
    // Only attach change listener if NOT disabled
    if (!cb.disabled) {
      cb.addEventListener('change', async (e) => {
        e.preventDefault();
        e.stopPropagation();
        
        const target = e.target as HTMLInputElement;
        const id = parseInt(target.dataset.id!);
        const isChecked = target.checked;
        
        console.log(`🔄 Checkbox changed: ID=${id}, checked=${isChecked}`);
        
        // Verify user has permission (double-check)
        if (!canFreeze()) {
          toast("error", "❌ You don't have permission to freeze/unfreeze projects");
          target.checked = !isChecked;
          return;
        }
        
        // Get the forecast
        const forecast = forecasts.find(f => f.Id === id);
        if (!forecast) {
          toast("error", "Forecast not found");
          target.checked = !isChecked;
          return;
        }
        
        try {
          const digest = await getDigest();
          const entityType = await getListEntityType();
          
          // Create payload
          const payload = {
            __metadata: { type: entityType },
            Freeze: isChecked
          };
          
          const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`, {
            method: "POST",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": digest,
              "IF-MATCH": "*",
              "X-HTTP-Method": "MERGE"
            },
            body: JSON.stringify(payload)
          });
          
          if (res.ok) {
            // Update the freeze status
            forecast.Freeze = isChecked;
            
            // Update filteredData
            const filteredForecast = filteredData.find(f => f.Id === id);
            if (filteredForecast) {
              filteredForecast.Freeze = isChecked;
            }
            
            // If this is the currently edited project, update the form
            const currentEditingId = editingId;
            if (currentEditingId === id) {
              setFormReadOnly(isChecked);
            }
            
            toast("success", `Project ${isChecked ? 'frozen' : 'unfrozen'} successfully`);
            
            // Re-render table
            renderCurrentPage();
          } else {
            const errorText = await res.text();
            console.error("❌ Error response:", errorText);
            toast("error", "Failed to update freeze status");
            target.checked = !isChecked;
          }
        } catch (error) {
          console.error("Error updating freeze status:", error);
          toast("error", "Failed to update freeze status");
          target.checked = !isChecked;
        }
      });
    }
  });
};
 
const getListEntityType = async (): Promise<string> => {
  try {
    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
      headers: { Accept: "application/json;odata=verbose" },
    });
    const data = await res.json();
    return data.d.ListItemEntityTypeFullName;
  } catch (error) {
    console.error("Error getting entity type:", error);
    // Fallback to a common type name
    return "SP.Data.Revenue_x0020_Forecast_x0020_ListListItem";
  }
};   

      const handleRevenueSort = (column: string, direction: 'asc' | 'desc', subColumn?: string, monthIndex?: number): void => {
        currentSortConfig = {
          column,
          direction,
          subColumn,
          monthIndex
        };

        const dataToSort = filteredData.length > 0 ? filteredData : forecasts;
        
        if (dataToSort.length === 0) return;

        const sortedData = [...dataToSort].sort((a, b) => {
          let valueA: any = 0;
          let valueB: any = 0;

          switch (column) {
            case 'obl':
              valueA = a.OBL || 0;
              valueB = b.OBL || 0;
              break;
            
            case 'month':
              const displayMonths = getDisplayMonths();
              
              const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
              const monthKeyToUse = displayMonths[monthIndexToUse];
              
              const dataA = safeJsonParse(a.ForecastData, {});
              const dataB = safeJsonParse(b.ForecastData, {});
              const monthDataA = dataA[monthKeyToUse] as ForecastMonthData || { plan: 0, actual: 0, forecast: 0 };
              const monthDataB = dataB[monthKeyToUse] as ForecastMonthData || { plan: 0, actual: 0, forecast: 0 };
              
              if (subColumn) {
                switch (subColumn) {
                  case 'plan':
                    valueA = monthDataA.plan || 0;
                    valueB = monthDataB.plan || 0;
                    break;
                  case 'current':
                    valueA = monthDataA.actual || 0;
                    valueB = monthDataB.actual || 0;
                    break;
                  case 'forecast':
                    valueA = monthDataA.forecast || 0;
                    valueB = monthDataB.forecast || 0;
                    break;
                  default:
                    valueA = monthDataA.actual || 0;
                    valueB = monthDataB.actual || 0;
                }
              } else {
                valueA = monthDataA.actual || 0;
                valueB = monthDataB.actual || 0;
              }
              break;
            
            case 'projectRef':
              valueA = a.ProjectReference || "";
              valueB = b.ProjectReference || "";
              break;

            case 'project':
              valueA = a.Project?.Title || "";
              valueB = b.Project?.Title || "";
              break;
            
            case 'customer':
              const projA = projects.find(p => p.Title === a.Project?.Title);
              const projB = projects.find(p => p.Title === b.Project?.Title);
              valueA = projA?.CustomerName || "";
              valueB = projB?.CustomerName || "";
              break;
            
            case 'groupHead': {
              const ghA = projects.find(p => p.Title === a.Project?.Title);
              const ghB = projects.find(p => p.Title === b.Project?.Title);
              valueA = getGroupHeadDisplayName(ghA?.GroupHead1) || "";
              valueB = getGroupHeadDisplayName(ghB?.GroupHead1) || "";
              break;
            }
            
            case 'projectManager': {
              const pmA = projects.find(p => p.Title === a.Project?.Title);
              const pmB = projects.find(p => p.Title === b.Project?.Title);
              valueA = getProjectManagerDisplayName(pmA?.ProjectManager) || "";
              valueB = getProjectManagerDisplayName(pmB?.ProjectManager) || "";
              break;
            }
            
            default:
              return 0;
          }

          if (typeof valueA === 'number' && typeof valueB === 'number') {
            return direction === 'asc' ? valueA - valueB : valueB - valueA;
          } else {
            const strA = String(valueA).toLowerCase();
            const strB = String(valueB).toLowerCase();
            return direction === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
          }
        });

        if (filteredData.length > 0) {
          filteredData = sortedData;
        } else {
          forecasts = sortedData;
        }

        currentPage = 1;
        updatePaginationControls();
        renderCurrentPage();
        
        const sortDirectionText = direction === 'asc' ? 'Ascending' : 'Descending';
        let columnName = '';
        
        if (subColumn) {
          const subColNames: { [key: string]: string } = {
            'plan': 'Prev. Plan',
            'current': 'Curr. Plan',
            'forecast': 'Forecast',
            'total': 'Further'
          };
          columnName = subColNames[subColumn] || subColumn;
        } else if (column === 'obl') {
          columnName = 'OBL';
        } else if (column === 'project') {
          columnName = 'Project';
        } else if (column === 'projectRef') {
          columnName = 'Project Reference';
        } else if (column === 'customer') {
          columnName = 'Customer';
        } else if (column === 'groupHead') {
          columnName = 'Group Head';
        } else if (column === 'projectManager') {
          columnName = 'Project Manager';
        } else {
          const displayMonths = getDisplayMonths();
          const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
          const monthKey = displayMonths[monthIndexToUse];
          const monthDate = new Date(monthKey + "-01");
          columnName = formatMonthLabel(monthDate);
        }
        
        toast("info", `Sorted by ${columnName} (${sortDirectionText})`);
      };

/* Clear = wipe the popup fields AND the applied filters (used by the popup
   Clear button, the popup Reset Filter button and the "Clear all" chip). */
const clearFilters = (): void => {
  appliedFilters = { ...EMPTY_FILTERS };
  restoreAppliedFilters();
  syncFilterUI();
  
  currentSortConfig = {
    column: '',
    direction: 'asc'
  };
  
  // ✅ After clearing, re-apply the project filter
  filterTable();
  
  const displayMonths = getDisplayMonths();
  const currentMonthKey = formatMonthKey(new Date());
  setTimeout(() => {
    renderTotalsHeader(displayMonths, currentMonthKey);
  }, 100);
  
  toast("info", "All filters cleared");
};

const showForm = (isEditMode: boolean = false, project?: ProjectItem, forecast?: RevenueForecastItem): void => {
  // Check if user can view this form
  if (!canViewForm(isEditMode, project, forecast)) {
    if (!canRead()) {
      toast("error", "❌ You don't have permission to access this form");
      return;
    }
    if (isEditMode && !canEditProject(project, forecast)) {
      toast("error", "❌ You don't have permission to edit this project");
      return;
    }
    toast("error", "❌ You don't have permission to create forecasts");
    return;
  }

  // ✅ FIX: Show form section, hide table section
  if (formSection) {
    formSection.style.display = "block";
    formSection.style.visibility = "visible";
    formSection.style.opacity = "1";
  }
  
  if (tableSection) {
    tableSection.style.display = "none";
  }
  
  // Show/hide save button(s) based on permissions
  // (top button beside the project name + bottom button stay in sync)
  const saveBtns = document.querySelectorAll<HTMLButtonElement>('button.revenue-save-btn');
  if (saveBtns.length > 0) {
    const canSave = isEditMode ? canEditProject(project, forecast) : canCreateForecast();
    saveBtns.forEach(btn => { btn.style.display = canSave ? '' : 'none'; });
  }
  
  if (isEditMode) {
    formSection?.classList.add('revenue-edit-mode');
    if (projectSelect) {
      projectSelect.disabled = true;
    }
    if (formTitle) {
      formTitle.textContent = "";
    }
  } else {
    formSection?.classList.remove('revenue-edit-mode');
    if (projectSelect) {
      projectSelect.disabled = false;
    }
    if (formTitle) {
      formTitle.textContent = "Add Forecast";
    }
  }
  
  // ✅ Force scroll to form
  setTimeout(() => {
    if (formSection) {
      formSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, 100);
};

      const hideForm = (): void => {
        // 🔧 Re-query the live DOM (fall back to the captured refs) and guard
        // every access - after a dashboard re-render the captured nodes can
        // be stale/missing, which used to crash the post-save cleanup here.
        const fs = (document.querySelector(".form-section") as HTMLElement | null) || formSection;
        const ts = (document.querySelector(".table-section") as HTMLElement | null) || tableSection;
        const ps = (document.querySelector("#projectSelect") as HTMLSelectElement | null) || projectSelect;

        if (fs) {
          fs.style.display = "none";
          fs.classList.remove('revenue-edit-mode');
        }
        if (ts) {
          ts.style.display = "block";
        }
        if (ps) {
          ps.disabled = false;
        }
      };

      const resetForm = (): void => {
        editingId = null;
        // 🔧 Guard every DOM access (same reason as hideForm) - a missing
        // element must never abort the post-save cleanup.
        if (form) {
          form.reset();
        }
        if (projectSelect) {
          projectSelect.value = "";
        }
        // 🔧 Do NOT wipe #projectDetails with innerHTML = "" - that destroys
        // the Project / Ref / PC / PM / GH / Customer rows AND the top Save
        // button that lives in the Project row. Reset the shown values
        // instead so the panel (and its button) survive every save cycle.
        ['#projectNameDisplay', '#projectRef', '#profitCenter', '#projectManager', '#groupHead', '#customer'].forEach(selector => {
          const detailEl = document.querySelector(selector) as HTMLElement | null;
          if (detailEl) detailEl.textContent = "-";
        });
        if (historicalContainer) historicalContainer.innerHTML = "";
        if (firstRowContainer) firstRowContainer.innerHTML = "";
        if (remainingContainer) remainingContainer.innerHTML = "";
        if (yearlyContainer) yearlyContainer.innerHTML = "";

        if (remarksInternal) {
          remarksInternal.value = "";
        }
        if (remarksSprintCell) {
          remarksSprintCell.value = "";
        }
        
        toggleForecastSections(false);
  
        const oblContainer = document.querySelector('.revenue-obl-section-container') as HTMLElement | null;
        if (oblContainer) {
          oblContainer.className = 'revenue-obl-section-container revenue-hidden-section';

          const fieldNote = oblContainer.querySelector('.revenue-field-note') as HTMLElement;
          if (fieldNote) fieldNote.textContent = "For new projects: Enter initial OBL value";
        }
        if (oblLastMonthInput) {
          oblLastMonthInput.removeAttribute('readonly');
          oblLastMonthInput.placeholder = "Enter OBL";
          setNumericValue(oblLastMonthInput, 0);
        }
        
        if (oblValidationMessage) {
          oblValidationMessage.textContent = "✅ All values are within OBL limit";
          oblValidationMessage.className = 'revenue-validation-message';
        }
        
        // Only re-run OBL validation when its live elements exist
        if (oblLastMonthInput && oblValidationMessage && oblLineFill && totalCurrentValueSpan && remainingValueSpan) {
          updateOBLValidationDisplay();
        }
      };

      // ============================================
      // LOADER HELPERS (page load / refresh / import)
      // Same names/classes as the collection page loader; the
      // overlay is anchored to .revenue-page, not the viewport.
      // ============================================

      const getRevenueLoaderMarkup = (): string => `
  <!-- ============================================ -->
  <!-- LOADER - covers ONLY this section (the page   -->
  <!-- wrapper), so the sidebar / top nav stay       -->
  <!-- visible and usable. Inline styles so it shows -->
  <!-- before CSS loads; hidden by hideRevenueLoader -->
  <!-- () in the initialize() finally block.         -->
  <!-- ============================================ -->
  <div id="revenuePageLoader" class="revenue-loader" style="position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;background:rgba(248,249,250,0.92);z-index:999;">
    <div class="revenue-loader-box">
      <div class="revenue-loader-spinner"></div>
      <div class="revenue-loader-text">Loading Revenue Data</div>
      <div class="revenue-loader-subtext">Please wait... fetching data</div>
      <div class="revenue-loader-bar"><span></span></div>
    </div>
  </div>`;

      const showRevenueLoader = (text?: string, subtext?: string): void => {
        let loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;

        // re-create it if hideRevenueLoader() already removed it from the DOM
        if (!loader) {
          const host = document.querySelector('.revenue-page') || document.body;
          host.insertAdjacentHTML('afterbegin', getRevenueLoaderMarkup());
          loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;
        }
        if (!loader) return;

        // ✅ No text on screen - the loader is the bar only (grey track +
        //    indigo bar). Progress messages go to the console instead.
        if (text || subtext) console.log(`⏳ ${text || 'Loading'}${subtext ? ' - ' + subtext : ''}`);

        loader.classList.remove('revenue-loader-hidden');
        loader.style.display = 'flex';
        loader.style.opacity = '1';
        loader.style.visibility = 'visible';

        // lock only the section, never <body> - the sidebar must stay usable
        loader.parentElement?.classList.add('revenue-loading');
        positionLoaderOverlay();
        window.setTimeout(positionLoaderOverlay, 0);

        const textEl = loader.querySelector('.revenue-loader-text');
        if (textEl && text) textEl.textContent = text;

        const subEl = loader.querySelector('.revenue-loader-subtext');
        if (subEl && subtext) subEl.textContent = subtext;
      };

// keeps the section-anchored overlay centred in the VISIBLE part of the
// section: its top edge follows the viewport (so a sticky header never covers
// it) and it is at least one viewport tall.
const positionLoaderOverlay = (): void => {
  const loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;
  if (!loader || !loader.parentElement) return;

  const rect = loader.parentElement.getBoundingClientRect();
  const top = Math.max(0, rect.top);
  const height = Math.max(window.innerHeight - top, rect.height - top, 320);

  loader.style.setProperty('--loader-top', `${top}px`);
  loader.style.height = `${height}px`;
};

window.addEventListener('resize', positionLoaderOverlay);
window.addEventListener('scroll', positionLoaderOverlay, { passive: true });

      const hideRevenueLoader = (): void => {
        const loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;
        document.querySelector('.revenue-page')?.classList.remove('revenue-loading');
        if (!loader) return;

        loader.classList.add('revenue-loader-hidden');

        // fade out, then remove from the DOM
        window.setTimeout(() => {
          if (loader.parentNode) loader.parentNode.removeChild(loader);
        }, 350);
      };

      const reloadRevenuePage = async (): Promise<void> => {
        showRevenueLoader('Reloading Data', 'Refreshing projects and forecasts...');
        try {
          const projectsLoad = loadProjects();
          const forecastsLoad = loadForecasts(projectsLoad);
          await Promise.all([projectsLoad, forecastsLoad]);
          initializeFilters();
         // renderTableHeader();
          filteredData = [...forecasts];
          currentPage = 1;
          totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
          updatePaginationControls();
          renderCurrentPage();
        } finally {
          hideRevenueLoader();
        }
      };

      // Manual refresh / debug hooks (also callable from the browser console)
      (window as any).reloadRevenueForecastData = reloadRevenuePage;
      (window as any).hideRevenueLoader = hideRevenueLoader;
 
 
 

      const blockIncreasingCurrentInputs = (allowedInput?: HTMLInputElement): void => {
        const currentInputs = document.querySelectorAll<HTMLInputElement>(
  'input.revenue-item-current-input:not([readonly])'
);
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>('input.revenue-yearly-current-input-rev:not([readonly])');
        
        currentInputs.forEach(input => {
          if (input === allowedInput) {
            input.style.backgroundColor = '';
            input.style.cursor = '';
            input.removeAttribute('readonly');
          } else {
            input.setAttribute('data-block-increase', 'true');
            input.style.backgroundColor = '#fff3cd';
            input.style.cursor = 'not-allowed';
          }
        });
        
        yearlyCurrentInputs.forEach(input => {
          if (input === allowedInput) {
            input.style.backgroundColor = '';
            input.style.cursor = '';
            input.removeAttribute('readonly');
          } else {
            input.setAttribute('data-block-increase', 'true');
            input.style.backgroundColor = '#fff3cd';
            input.style.cursor = 'not-allowed';
          }
        });
      };

const unblockAllCurrentInputs = (): void => {

  // Editable Current Plan inputs
  // Long-term Forecast + 5-Year Forecast
  const currentInputs = document.querySelectorAll<HTMLInputElement>(
    'input.revenue-item-current-input'
  );

  const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>(
    'input.revenue-yearly-current-input-rev'
  );

  currentInputs.forEach(input => {
    input.style.backgroundColor = '';
    input.style.cursor = '';
    input.removeAttribute('data-block-increase');

    // Editable
    input.removeAttribute('readonly');
  });

  yearlyCurrentInputs.forEach(input => {
    input.style.backgroundColor = '';
    input.style.cursor = '';
    input.removeAttribute('data-block-increase');

    // Editable
    input.removeAttribute('readonly');
  });

  // IMPORTANT:
  // Detailed Forecast main Curr. Plan must stay readonly.
  document.querySelectorAll<HTMLInputElement>(
    '.revenue-detailed-month-box .revenue-main-current-input'
  ).forEach(input => {
    input.setAttribute('readonly', 'true');
    input.style.cursor = 'default';
  });
};

      const allowInputIncrease = (input: HTMLInputElement, newValue: number): boolean => {
        const oldValue = parseInt(input.getAttribute('data-old-value') || '0');
        const isIncreasing = newValue > oldValue;
        
        if (!isIncreasing) {
          return true;
        }
        
        const oblValue = getNumericValue(oblLastMonthInput);
        const currentTotal = calculateTotalCurrentValue();
        const increaseAmount = newValue - oldValue;
        
        return (currentTotal + increaseAmount) <= oblValue;
      };

      const allowItemInputIncrease = (input: HTMLInputElement, newValue: number): boolean => {
        const oldValue = parseInt(input.getAttribute('data-old-value') || '0');
        const isIncreasing = newValue > oldValue;
        
        if (!isIncreasing) {
            return true;
        }
        
        const monthKey = input.dataset.month!;
        const mainCurrentInput = document.querySelector<HTMLInputElement>(`input.revenue-main-current-input[data-month="${monthKey}"]`);
        
        if (!mainCurrentInput) return true;
        
        const mainCurrentValue = getNumericValue(mainCurrentInput);
        const itemCurrentInputs = document.querySelectorAll<HTMLInputElement>(`input.revenue-item-current-input[data-month="${monthKey}"]`);
        
        let currentItemsTotal = 0;
        itemCurrentInputs.forEach(itemInput => {
            if (itemInput === input) {
                currentItemsTotal += newValue;
            } else {
                currentItemsTotal += getNumericValue(itemInput);
            }
        });
        
        return currentItemsTotal <= mainCurrentValue;
      };

      const validateMonthItemTotals = (changedInput?: HTMLInputElement): boolean => {
        let allValid = true;
        
        for (let i = 3; i <= 5; i++) {
          if (i < allMonths.length) {
            const monthKey = allMonths[i];
            const monthValid = validateSingleMonthItems(monthKey, changedInput);
            if (!monthValid) {
              allValid = false;
            }
          }
        }
        
        return allValid;
      };

      const validateSingleMonthItems = (monthKey: string, changedInput?: HTMLInputElement): boolean => {
        
       
        
        return true;
      
      };

      const onProjectChange = (projectId: number, existingForecast?: RevenueForecastItem): void => {
        const project = projects.find(p => p.Id === projectId);
        if (!project) return;

        toggleForecastSections(true);

        renderProjectDetails(project);
        
        let forecastData: ForecastData = {};
        if (existingForecast) {
          forecastData = safeJsonParse(existingForecast.ForecastData, {});
        }

        const isNewProject = !existingForecast;
        
        const oblContainer = document.querySelector('.revenue-obl-section-container') as HTMLElement;
        const fieldNote = oblContainer.querySelector('.revenue-field-note') as HTMLElement;
        
        if (isNewProject) {
          oblContainer.className = 'revenue-obl-section-container revenue-visible-section revenue-new-project';
          oblLastMonthInput.removeAttribute('readonly');
          oblLastMonthInput.placeholder = "Enter initial OBL for new project";
          if (fieldNote) fieldNote.textContent = "For new projects: Enter initial OBL value";
          oblValidationMessage.textContent = "Enter initial OBL value for new project";
          oblValidationMessage.className = 'revenue-validation-message';
          
          const defaultOBL = project.PC ? 0 : 0;
          setNumericValue(oblLastMonthInput, defaultOBL);
        } else {
          oblContainer.className = 'revenue-obl-section-container revenue-visible-section revenue-existing-project';
          oblLastMonthInput.setAttribute('readonly', 'true');
          
          const oblValue = existingForecast?.OBL || 0;
          setNumericValue(oblLastMonthInput, oblValue);
          
          oblLastMonthInput.placeholder = "OBL from database";
          if (fieldNote) fieldNote.textContent = "OBL value from database record";
          
          if (oblValue > 0) {
            oblLastMonthInput.value = oblValue.toString();
            oblLastMonthInput.className = oblLastMonthInput.className.replace('revenue-input-empty', 'revenue-input-with-value');
          }
        }

        renderForecastInterface(forecastData, getNumericValue(oblLastMonthInput), existingForecast);
      };

      const toggleForecastSections = (show: boolean): void => {
        const sections = [
          '.revenue-obl-section-container',
          '#historicalForecastSection',
          '#detailedForecastSection', 
          '#longtermForecastSection',
          '#yearlyForecastSection'
        ];
        
        sections.forEach(selector => {
          const element = document.querySelector(selector) as HTMLElement;
          if (element) {
            if (show) {
              element.classList.remove('revenue-hidden-section');
              element.classList.add('revenue-visible-section');
            } else {
              element.classList.remove('revenue-visible-section');
              element.classList.add('revenue-hidden-section');
            }
          }
        });
      };

const renderProjectDetails = (project: ProjectItem): void => {

   const canEdit = canEditProject(project);
  
  // Update Project Name
  const projectNameDisplay = document.querySelector('#projectNameDisplay');
  if (projectNameDisplay) {
    projectNameDisplay.textContent = project.Title || "-";
    // Add lock icon if read-only
    if (!canEdit) {
      projectNameDisplay.innerHTML = `${project.Title || "-"} 🔒`;
    }
  }
  
  // Update Project Name (this should be the FIRST row) 
  if (projectNameDisplay) projectNameDisplay.textContent = project.Title || "-";
  
  // Update Project Ref (this should be the SECOND row)  
  const projectRef = document.querySelector('#projectRef');
  if (projectRef) projectRef.textContent = project.ProjectReference || "-";
  
  // Update Profit Center (third row)
  const profitCenter = document.querySelector('#profitCenter');
  if (profitCenter) profitCenter.textContent = project.PC || "-";
  
  // Update Project Manager (fourth row)
   const projectManager = document.querySelector('#projectManager');
  if (projectManager) {
    projectManager.textContent = getProjectManagerDisplayName(project.ProjectManager) || "-";
  }
  
  // ✅ Group Head - Person/Group type
  const groupHead = document.querySelector('#groupHead');
  if (groupHead) {
    groupHead.textContent = getGroupHeadDisplayName(project.GroupHead1) || "-";
  }
  // Update Customer (sixth row)
  const customer = document.querySelector('#customer');
  if (customer) customer.textContent = project.CustomerName || "-";
};
const getForecastMonths = (): string[] => {
  return generateIdealMonthStructure(); // This should generate 36+ months
};

// COMPLETE REPLACEMENT FOR renderForecastInterface
const renderForecastInterface = (existingData: ForecastData, displayOBL: number, existingForecast?: RevenueForecastItem): void => {
    allMonths = getForecastMonths();
    
    setNumericValue(oblLastMonthInput, displayOBL);
    
    // IMPORTANT: Get all existing months from the data to preserve them
    const existingMonths = Object.keys(existingData).filter(key => key.match(/^\d{4}-\d{2}$/));
    
    // Get all 54 months
    const allStorageMonths = generateIdealMonthStructure();
    
    // Store ALL 36 historical months
    const storeHistoricalMonths = allStorageMonths.slice(0, 36);
    const displayHistoricalMonths = allStorageMonths.slice(33, 36);
    
    // Ensure ALL historical months exist in the data (preserve existing values)
    storeHistoricalMonths.forEach((monthKey) => {
        if (!existingData[monthKey]) {
            existingData[monthKey] = {
                plan: 0,
                actual: 0,
                items: {}
            };
        } else {
            // Keep existing data, don't overwrite
            // Just ensure items exists
            const monthData = existingData[monthKey] as ForecastMonthData;
            if (!monthData.items) {
                monthData.items = {};
            }
        }
    });
    
    // Also preserve any months that might be outside the ideal structure
    existingMonths.forEach(monthKey => {
        if (!storeHistoricalMonths.includes(monthKey) && 
            !allStorageMonths.slice(36).includes(monthKey)) {
            if (!existingData[monthKey]) {
                existingData[monthKey] = {
                    plan: 0,
                    actual: 0,
                    items: {}
                };
            }
        }
    });
    
    // Render only the LAST 3 months in the UI
    renderHistoricalMonths(displayHistoricalMonths, existingData);
    
    setTimeout(() => {
        setupHistoricalToggle();
    }, 50);
    
    // Pass existingForecast to get RevenueActual for current month
    renderFirstRow(allStorageMonths.slice(36, 39), existingData, existingForecast);
    renderRemainingMonths(allStorageMonths.slice(39), existingData);
    renderYearlyForecast(existingData.yearly || createFreshYearlyStructure(getLastMonthKey(existingData)));
    
    protectHistoricalElements();
    
    updateOBLValidationDisplay();
    attachForecastEventListeners();
    
    // Apply freeze state
    if (existingForecast) {
    const isFrozen = existingForecast.Freeze === true;

    setTimeout(() => {
        setFormReadOnly(isFrozen);

        const formTitleEl =
            document.querySelector("#formTitle") as HTMLElement;

        if (formTitleEl) {
            if (isFrozen) {
                formTitleEl.innerHTML =
                    '🔒 Edit Forecast (Frozen - Further fields editable)';
                formTitleEl.style.color = '#856404';
            } else {
                formTitleEl.innerHTML = 'Edit Forecast';
                formTitleEl.style.color = '';
            }
        }
    }, 250);
}
};

       
    
const setupHistoricalToggle = (): void => {
  const toggleBtn = document.querySelector("#revenue-toggle-historical-btn") as HTMLButtonElement;
  const historicalContainer = document.querySelector("#historicalContainer") as HTMLElement;
  const header = document.querySelector(".revenue-historical-toggle-header") as HTMLElement;
  
  if (!toggleBtn || !historicalContainer || !header) return;
  
  let isVisible = false;
  
  // Build the compact arrow + label button (image 1 style)
  const newBtn = document.createElement('button');
  newBtn.id = 'revenue-toggle-historical-btn';
  newBtn.type = 'button';
  newBtn.setAttribute('aria-controls', 'historicalContainer');
  newBtn.setAttribute('aria-expanded', 'false');
  newBtn.setAttribute('aria-label', 'Show historical data');
  newBtn.title = 'Show historical data';
  newBtn.innerHTML = '<span class="revenue-toggle-arrow" aria-hidden="true">▼</span><span class="revenue-toggle-label">Show</span>';
  
  toggleBtn.parentNode?.replaceChild(newBtn, toggleBtn);
  
  // Start hidden
  historicalContainer.style.display = 'none';
  
  const applyState = (visible: boolean) => {
    const btn = document.querySelector("#revenue-toggle-historical-btn") as HTMLButtonElement;
    const container = document.querySelector("#historicalContainer") as HTMLElement;
    if (!btn || !container) return;
    const label = btn.querySelector('.revenue-toggle-label');
    if (visible) {
      container.style.display = "";
      if (label) label.textContent = "Hide";
      btn.setAttribute('aria-expanded', 'true');
      btn.setAttribute('aria-label', 'Hide historical data');
      btn.title = 'Hide historical data';
    } else {
      container.style.display = "none";
      if (label) label.textContent = "Show";
      btn.setAttribute('aria-expanded', 'false');
      btn.setAttribute('aria-label', 'Show historical data');
      btn.title = 'Show historical data';
    }
    isVisible = visible;
  };
  
  // Click anywhere on the header (including button) toggles
  const headerEl = document.querySelector(".revenue-historical-toggle-header") as HTMLElement;
  if (headerEl) {
    headerEl.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      applyState(!isVisible);
    });
  }
};

       const protectHistoricalElements = (): void => {
        document.querySelectorAll('.revenue-historical-month input, .revenue-historical-month select, .revenue-historical-month button').forEach(element => {
          const newElement = element.cloneNode(true) as HTMLElement;
          element.parentNode?.replaceChild(newElement, element);
        });
        
        document.querySelectorAll('.revenue-historical-month').forEach(section => {
          section.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            toast("warning", "Historical data cannot be modified");
            return false;
          });
          
          section.addEventListener('keydown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            return false;
          });
          
          section.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            return false;
          });
        });
      };

const renderHistoricalMonths = (months: string[], existingData: ForecastData): void => {
  let html = '';

  months.forEach((monthKey, index) => {
    const monthDate = new Date(monthKey + "-01");
    const monthLabel = formatMonthLabel(monthDate);
    const data = existingData[monthKey] as ForecastMonthData || { plan: 0, actual: 0, items: {} };
    
    html += renderHistoricalMonthBox(monthKey, monthLabel, data, index);
  });

  historicalContainer.innerHTML = html;
};

      const renderHistoricalMonthBox = (monthKey: string, monthLabel: string, data: ForecastMonthData, index: number): string => {
        const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
        const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
        const planInputClass = data.plan > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
        const actualInputClass = data.actual > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
        
        const items = data.items || createEmptyItems();
        
        return `
          <div class="revenue-month-box-rev revenue-historical-month">
            <div class="revenue-month-header-rev">${monthLabel} (Historical - Readonly)</div>
            
            <div class="revenue-month-block">
              <div class="revenue-main-inputs-container-rev">
                <div class="revenue-main-input-group-rev">
                  <label>Prev. Plan</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="revenue-main-plan-input ${planInputClass} revenue-number-input" />
                </div>
                <div class="revenue-main-input-group-rev">
                  <label>Curr. Plan</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        readonly class="revenue-main-current-input ${actualInputClass} revenue-number-input" />
                </div>
              </div>
            </div>

            <div class="revenue-items-section revenue-historical-items-section">
              <div class="revenue-items-header revenue-historical-items-header">
                <div>Item</div>
                <div>Prev. Plan</div>
                <div>Curr. Plan</div>
              </div>
              
              ${Array.from({length: 10}, (_, i) => {
                const itemNum = i + 1;
                const itemKey = itemNum.toString();
                const item = items[itemKey] || { name: '', plan: 0, actual: 0 };
                const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                const itemPlanClass = item.plan > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
                const itemActualClass = item.actual > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
                const itemNameClass = item.name ? 'revenue-input-with-value' : 'revenue-input-empty';
                
                return `
                  <div class="revenue-item-row revenue-historical-item-row">
                   <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
      data-item="${itemNum}" data-month="${monthKey}" 
      class="${itemNameClass}" style="pointer-events: auto !important; cursor: text !important;" />
                    <input type="text" value="${itemPlanValue}" data-item-type="plan" data-item="${itemKey}" data-month="${monthKey}" 
                          readonly class="revenue-item-plan-input ${itemPlanClass} revenue-number-input" />
                    <input type="text" value="${itemActualValue}" data-item-type="actual" data-item="${itemKey}" data-month="${monthKey}" 
                          readonly class="revenue-item-current-input ${itemActualClass} revenue-number-input" />
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      };

const renderFirstRow = (months: string[], existingData: ForecastData, existingForecast?: RevenueForecastItem): void => {
    let html = '';

    html += `<div class="revenue-detailed-months-row">`;
    
    // Get current month key
    const now = new Date();
    const currentMonthKey = formatMonthKey(now);
    
    // Find current month in the months array
    const currentMonthIndex = months.indexOf(currentMonthKey);
    
    // Detailed months: current month and next 2 months
    let detailedMonths: string[] = [];
    if (currentMonthIndex !== -1 && currentMonthIndex + 3 <= months.length) {
        detailedMonths = months.slice(currentMonthIndex, currentMonthIndex + 3);
    } else if (currentMonthIndex !== -1) {
        detailedMonths = months.slice(currentMonthIndex);
    } else {
        detailedMonths = months.slice(0, 3);
    }
    
    console.log(`📊 Detailed months for rendering: ${detailedMonths.join(', ')}`);
    
    // ✅ Load RevenueActual data from separate column
    let revenueActualData: { [key: string]: number } = {};
    if (existingForecast) {
        revenueActualData = safeJsonParse(existingForecast.RevenueActual, {});
        console.log("📊 RevenueActual data:", revenueActualData);
    }
    
    // Current month is the first of the detailed months
    const currentMonth = detailedMonths[0];
    if (currentMonth) {
        const monthDate = new Date(currentMonth + "-01");
        const monthLabel = formatMonthLabel(monthDate);
        const data = existingData[currentMonth] as ForecastMonthData || { 
            plan: 0, 
            actual: 0, 
            monthActual: 0,
            items: {} 
        };
        
        // ✅ Set monthActual from RevenueActual column for current month
        if (revenueActualData[currentMonth] !== undefined) {
            data.monthActual = revenueActualData[currentMonth];
            console.log(`📊 Set monthActual for ${currentMonth}:`, data.monthActual);
        }
        
        html += `<div class="revenue-current-month-container">`;
        html += renderDetailedMonthBox(currentMonth, monthLabel, data, true);
        html += `</div>`;
    }
    
    // Next 2 months - also get monthActual from RevenueActual if available
    html += `<div class="revenue-next-months-container">`;
    for (let i = 1; i < 3 && i < detailedMonths.length; i++) {
        const monthKey = detailedMonths[i];
        const monthDate = new Date(monthKey + "-01");
        const monthLabel = formatMonthLabel(monthDate);
        const data = existingData[monthKey] as ForecastMonthData || { 
            plan: 0, 
            actual: 0, 
            monthActual: 0,
            items: {} 
        };
        
        // ✅ Set monthActual from RevenueActual for next months
        if (revenueActualData[monthKey] !== undefined) {
            data.monthActual = revenueActualData[monthKey];
        }
        
        html += `<div class="revenue-next-month-box">`;
        html += renderDetailedMonthBox(monthKey, monthLabel, data, false);
        html += `</div>`;
    }
    html += `</div>`;
    
    html += `</div>`;

    firstRowContainer.innerHTML = html;
};
      const renderRemainingMonths = (months: string[], existingData: ForecastData): void => {
        let html = '';

        for (let row = 0; row < 3; row++) {
          html += `<div class="revenue-remaining-row-rev">`;
          for (let col = 0; col < 5; col++) {
            const index = (row * 5) + col;
            if (index >= months.length) break;
            
            const monthKey = months[index];
            const monthDate = new Date(monthKey + "-01");
            const monthLabel = formatMonthLabel(monthDate);
            const data = existingData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
            
            html += renderSimpleMonthBox(monthKey, monthLabel, data);
          }
          html += `</div>`;
        }

        remainingContainer.innerHTML = html;
      };

      const renderYearlyForecast = (yearlyData: YearlyForecastData): void => {
        let html = '';

        const years = Object.keys(yearlyData).sort();
        
        years.forEach(year => {
          const data = yearlyData[year] || { plan: 0, actual: 0 };
          html += renderYearlyBox(year, data);
        });

        yearlyContainer.innerHTML = html;
      };

      const renderYearlyBox = (year: string, data: { plan: number; actual: number }): string => {
        const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
        const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
        const planInputClass = data.plan > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
        const actualInputClass = data.actual > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
        
        return `
          <div class="revenue-yearly-box-rev">
            <div class="revenue-yearly-header-rev">${year}</div>
            
            <div class="revenue-yearly-inputs-container-rev">
              <div class="revenue-yearly-input-group-rev">
                <label>Prev. Plan</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                      readonly class="revenue-yearly-plan-input-rev ${planInputClass} revenue-number-input" />
              </div>
              <div class="revenue-yearly-input-group-rev">
                <label>Curr. Plan</label>
                <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                      class="revenue-yearly-current-input-rev ${actualInputClass} revenue-number-input" />
              </div>
            </div>
          </div>
        `;
      };

const renderDetailedMonthBox = (monthKey: string, monthLabel: string, data: ForecastMonthData, isCurrentMonth: boolean = false): string => {
    console.log(`📊 renderDetailedMonthBox for ${monthLabel} (${monthKey}):`, {
        plan: data.plan,
        actual: data.actual,
        monthActual: data.monthActual,
        itemsCount: data.items ? Object.keys(data.items).length : 0
    });
    
    const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
    const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
    const displayForecastValue = data.forecast || 0;
    const displayDateValue = data.date || '';
    const displayMonthActual = data.monthActual || 0;
    
    const planInputClass = data.plan > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
    const actualInputClass = data.actual > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
    const forecastInputClass = displayForecastValue > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
    const dateInputClass = displayDateValue ? 'revenue-input-with-value' : 'revenue-input-empty';
    
    const items = data.items || createEmptyItems();

    // Compute first/last day of this month for date-picker bounds
    const [yStr, mStr] = monthKey.split('-');
    const yNum = parseInt(yStr, 10);
    const mNum = parseInt(mStr, 10);
    const monthFirstDay = `${yStr}-${mStr}-01`;
    const lastDayDate = new Date(yNum, mNum, 0);
    const monthLastDay = `${yStr}-${mStr}-${String(lastDayDate.getDate()).padStart(2, '0')}`;

    // 🔧 Month Actual badge - ONLY for the CURRENT month (hidden on next 2)
    const hasMonthActual = displayMonthActual > 0;
    const monthActualBadge = isCurrentMonth ? `
        <span class="revenue-month-actual-badge ${hasMonthActual ? 'revenue-has-value' : ''}" style="
            background-color: ${hasMonthActual ? '#c8e6c9' : '#f5f5f5'};
            border: 1px solid ${hasMonthActual ? '#4caf50' : '#ddd'};
            border-radius: 4px;
            padding: 2px 12px;
            margin-left: auto;
            font-size: 13px;
            font-weight: 600;
            color: ${hasMonthActual ? '#1b5e20' : '#999'};
            display: inline-flex;
            align-items: center;
            gap: 6px;
            white-space: nowrap;
        ">
            <span>📊 Actual Revenue Till date:</span>
            <span style="font-size: 16px; font-weight: 700;">${displayMonthActual.toLocaleString('en-US')}</span>
        </span>
    ` : '';
    
    return `
        <div class="revenue-detailed-month-box ${isCurrentMonth ? 'revenue-current-month-box' : 'revenue-next-month-box'}">
            <div class="revenue-month-header-rev" style="
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 8px 12px;
                background: #f8f9fa;
                border-bottom: 2px solid #007bff;
                border-radius: 4px 4px 0 0;
            ">
                <span style="font-size: 15px; font-weight: 700; color: #333;">
                    ${monthLabel} ${isCurrentMonth ? '<span class="revenue-current-month-indicator" style="margin-left: 8px; font-size: 12px; color: #007bff; background: #e3f2fd; padding: 2px 10px; border-radius: 12px;">Current</span>' : ''}
                </span>
                ${monthActualBadge}
            </div>
            
            <div class="revenue-month-block" style="padding: 8px 12px;">
                <div class="revenue-enhanced-inputs-container" style="
                    display: grid;
                    grid-template-columns: ${isCurrentMonth ? '1fr 1fr 1fr 1fr' : '1fr 1fr'};
                    gap: 8px;
                ">
                    <div class="revenue-enhanced-input-group">
                        <label style="font-size: 11px; font-weight: 600; color: #666;">Prev. Plan</label>
                        <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                              readonly class="revenue-main-plan-input ${planInputClass} revenue-number-input" style="width: 100%; padding: 4px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center;" />
                    </div>
                    
                    <div class="revenue-enhanced-input-group">
                        <label style="font-size: 11px; font-weight: 600; color: #666;">Curr. Plan</label>
                        <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                              readonly class="revenue-main-current-input ${actualInputClass} revenue-number-input" style="width: 100%; padding: 4px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center;" data-calc-total="true" />
                    </div>
                    
                    ${isCurrentMonth ? `
                        <div class="revenue-enhanced-input-group">
                            <label style="font-size: 11px; font-weight: 600; color: #666;">Further</label>
                            <input type="text" value="${displayForecastValue}" data-type="forecast" data-month="${monthKey}"
                                  class="revenue-main-forecast-input ${forecastInputClass} revenue-number-input" style="width: 100%; padding: 4px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center; background-color: #ffffff;" data-calc-total="true" />
                        </div>
                        
                        <div class="revenue-enhanced-input-group">
                            <label style="font-size: 11px; font-weight: 600; color: #666;">Date</label>
                            <input type="date" value="${displayDateValue}" data-type="date" data-month="${monthKey}"
                                  readonly class="revenue-forecast-date-input ${dateInputClass}" style="width: 100%; padding: 4px 6px; border: 1px solid #28a745; border-radius: 4px; background-color: #f5f5f5; cursor: default;" />
                        </div>
                    ` : ''}
                </div>
            </div>

            <div class="revenue-items-section" style="padding: 0 12px 12px 12px;">
                <div class="revenue-items-header" style="
                    display: grid;
                    grid-template-columns: ${isCurrentMonth ? '1fr 0.8fr 0.8fr 0.8fr 0.8fr' : '1fr 0.8fr 0.8fr'};
                    gap: 4px;
                    padding: 6px 4px;
                    background: #e9ecef;
                    border-radius: 4px 4px 0 0;
                    font-weight: 600;
                    font-size: 11px;
                    color: #495057;
                ">
                    <div>Item</div>
                    <div>Prev. Plan</div>
                    <div>Curr. Plan</div>
                    ${isCurrentMonth ? `
                        <div>Further</div>
                        <div>Date</div>
                    ` : ''}
                </div>
                
                ${Array.from({length: 10}, (_, i) => i + 1).map(itemNum => {
                    const item = items[itemNum] || { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
                    const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                    const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                    const itemForecastValue = (item.forecast || 0) > 0 ? (item.forecast || 0).toString() : '';
                    const itemDateValue = item.date || '';
                    
                    const itemPlanClass = item.plan > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
                    const itemActualClass = item.actual > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
                    const itemForecastClass = (item.forecast || 0) > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
                    const itemDateClass = itemDateValue ? 'revenue-input-with-value' : 'revenue-input-empty';
                    const itemNameClass = item.name ? 'revenue-input-with-value' : 'revenue-input-empty';
                    
                    return `
                        <div class="revenue-item-row ${isCurrentMonth ? 'revenue-row' : ''}" style="
                            display: grid;
                            grid-template-columns: ${isCurrentMonth ? '1fr 0.8fr 0.8fr 0.8fr 0.8fr' : '1fr 0.8fr 0.8fr'};
                            gap: 4px;
                            padding: 3px 4px;
                            border-bottom: 1px solid #f0f0f0;
                            align-items: center;
                        ">
                            <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
                                  data-item="${itemNum}" data-month="${monthKey}" 
                                  class="revenue-item-description-input ${itemNameClass}" 
                                  style="width: 100%; padding: 3px 6px; border: 1px solid #eee27e; border-radius: 4px; font-size: 11px; pointer-events: auto !important; cursor: text !important; background-color: #fffde7 !important;" />
                            <input type="text" value="${itemPlanValue}" 
                                  data-item-type="plan" data-item="${itemNum}" data-month="${monthKey}" 
                                  readonly class="revenue-item-plan-input ${itemPlanClass} revenue-number-input" style="width: 100%; padding: 3px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center; font-size: 11px;" />
                            
                            <input type="text" value="${itemActualValue}" 
                                  data-item-type="actual" data-item="${itemNum}" data-month="${monthKey}" 
                                  class="revenue-item-current-input ${itemActualClass} revenue-number-input" style="width: 100%; padding: 3px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center; font-size: 11px;" data-calc-total="true" />
                            
                            ${isCurrentMonth ? `
                                <input type="text" value="${itemForecastValue}" 
                                      data-item-type="forecast" data-item="${itemNum}" data-month="${monthKey}" 
                                      class="revenue-item-forecast-input ${itemForecastClass} revenue-number-input" style="width: 100%; padding: 3px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center; font-size: 11px; background-color: #ffffff;" data-calc-total="true" />
                                
                                <input type="date" value="${itemDateValue}" 
                                      data-item-type="date" data-item="${itemNum}" data-month="${monthKey}"
                                      min="${monthFirstDay}" max="${monthLastDay}"
                                      class="revenue-item-date-input ${itemDateClass}" style="width: 100%; padding: 3px 6px; border: 1px solid #ddd; border-radius: 4px; font-size: 11px; background-color: #ffffff;" />
                            ` : ''}
                        </div>
                    `;
                }).join('')}
            </div>
        </div>
    `;
};
      const renderSimpleMonthBox = (monthKey: string, monthLabel: string, data: ForecastMonthData): string => {
        const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
        const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
        const planInputClass = data.plan > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
        const actualInputClass = data.actual > 0 ? 'revenue-input-with-value' : 'revenue-input-empty';
        
        return `
          <div class="revenue-month-box-rev revenue-simple-month-rev">
            <div class="revenue-month-header-rev">${monthLabel}</div>
            
            <div class="revenue-main-inputs-container-rev">
              <div class="revenue-main-input-group-rev">
                <label>Prev. Plan</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}" 
                      readonly class="revenue-main-plan-input ${planInputClass} revenue-number-input" />
              </div>
              <div class="revenue-main-input-group-rev">
                <label>Curr. Plan</label>
                <input type="text" value="${displayActualValue}"  id="actualCurrent" data-type="actual" data-month="${monthKey}" 
                      class="revenue-main-current-input ${actualInputClass} revenue-number-input" />
              </div>
            </div>
          </div>
        `;
      };
      const attachForecastEventListeners = (): void => {
        document.querySelectorAll<HTMLInputElement>('.revenue-number-input').forEach(input => {
            allowOnlyNumbers(input);
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-number-input:not([readonly])').forEach(input => {
            input.addEventListener('focus', function() {
                this.setAttribute('data-old-value', getNumericValue(this).toString());
            });
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-main-current-input:not([readonly])').forEach(input => {
            input.addEventListener('input', function() {
                const currentValue = getNumericValue(this);
                const oldValue = parseInt(this.getAttribute('data-old-value') || '0');
                
                this.setAttribute('data-old-value', currentValue.toString());
                
                const isIncreasing = currentValue > oldValue;
                
                if (isIncreasing) {
                    const allowByOBL = allowInputIncrease(this, currentValue);
                    
                    if (!allowByOBL) {
                        setNumericValue(this, oldValue);
                        updateOBLValidationDisplay(this);
                        return;
                    }
                }
                
                if (currentValue === 0) {
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else {
                    this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                }

                updateOBLValidationDisplay(this);
            });
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-main-forecast-input:not([readonly])').forEach(input => {
            input.addEventListener('input', function() {
                const currentValue = getNumericValue(this);
                const oldValue = parseInt(this.getAttribute('data-old-value') || '0');
                
                this.setAttribute('data-old-value', currentValue.toString());
                
                const isIncreasing = currentValue > oldValue;
                
                if (isIncreasing) {
                    const allowByOBL = allowInputIncrease(this, currentValue);
                    
                    if (!allowByOBL) {
                        setNumericValue(this, oldValue);
                        updateOBLValidationDisplay(this);
                        return;
                    }
                }
                
                if (currentValue === 0) {
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else {
                    this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                }

                updateOBLValidationDisplay(this);
            });
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-item-current-input:not([readonly])').forEach(input => {
            input.addEventListener('input', function() {
                const currentValue = getNumericValue(this);
                const oldValue = parseInt(this.getAttribute('data-old-value') || '0');
                
                this.setAttribute('data-old-value', currentValue.toString());
                
                const isIncreasing = currentValue > oldValue;
                
                if (isIncreasing) {
                    const allowByOBL = allowInputIncrease(this, currentValue);
                    const allowByMonth = allowItemInputIncrease(this, currentValue);
                    
                    if (!allowByOBL || !allowByMonth) {
                        setNumericValue(this, oldValue);
                        updateOBLValidationDisplay(this);
                        return;
                    }
                }
                
                if (currentValue === 0) {
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else {
                    this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                }

                updateOBLValidationDisplay(this);
            });
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-item-forecast-input:not([readonly])').forEach(input => {
            input.addEventListener('input', function() {
                const currentValue = getNumericValue(this);
                const oldValue = parseInt(this.getAttribute('data-old-value') || '0');
                
                this.setAttribute('data-old-value', currentValue.toString());
                
                const isIncreasing = currentValue > oldValue;
                
                if (isIncreasing) {
                    const allowByOBL = allowInputIncrease(this, currentValue);
                    const allowByMonth = allowItemInputIncrease(this, currentValue);
                    
                    if (!allowByOBL || !allowByMonth) {
                        setNumericValue(this, oldValue);
                        updateOBLValidationDisplay(this);
                        return;
                    }
                }
                
                if (currentValue === 0) {
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else {
                    this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                }

                updateOBLValidationDisplay(this);
            });
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-yearly-current-input-rev:not([readonly])').forEach(input => {
            input.addEventListener('input', function() {
                const currentValue = getNumericValue(this);
                const oldValue = parseInt(this.getAttribute('data-old-value') || '0');
                
                this.setAttribute('data-old-value', currentValue.toString());
                
                const isIncreasing = currentValue > oldValue;
                
                if (isIncreasing) {
                    const allowByOBL = allowInputIncrease(this, currentValue);
                    
                    if (!allowByOBL) {
                        setNumericValue(this, oldValue);
                        updateOBLValidationDisplay(this);
                        return;
                    }
                }
                
                if (currentValue === 0) {
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else {
                    this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                }

                updateOBLValidationDisplay(this);
            });
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-forecast-date-input:not([readonly])').forEach(input => {
            input.addEventListener('change', function() {
                if (this.value === '') {
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else {
                    this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                }
            });
        });

        document.querySelectorAll<HTMLInputElement>('.revenue-item-date-input:not([readonly])').forEach(input => {
            input.addEventListener('change', function() {
                const monthKey = this.dataset.month || '';
                const minVal = this.min;
                const maxVal = this.max;

                // Enforce same-month: reject dates outside this month's min/max
                if (this.value && ((minVal && this.value < minVal) || (maxVal && this.value > maxVal))) {
                    toast("warning", `Date must be within the same month (${minVal} to ${maxVal})`);
                    this.value = '';
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else if (this.value === '') {
                    this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                } else {
                    this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                }

                // Auto-sync the main (summary) Date input with the GREATEST item date in this month
                if (monthKey) {
                    syncMainDateFromItems(monthKey);
                }
            });
             const currentMonthKey = formatMonthKey(new Date());
        setTimeout(() => syncMainDateFromItems(currentMonthKey), 0)
        });

        document.querySelectorAll<HTMLInputElement>('input[data-item]:not([readonly])').forEach(input => {
            if (!input.classList.contains('revenue-number-input')) {
                input.addEventListener('input', function() {
                    if (this.value === '') {
                        this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
                    } else {
                        this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
                    }
                });
            }
        });

 

        document.querySelectorAll<HTMLInputElement>('.revenue-main-forecast-input').forEach(input => {
        // Ensure it's always editable
        input.removeAttribute('readonly');
        input.style.backgroundColor = '#ffffff';
        input.style.pointerEvents = 'auto';
        input.disabled = false;
        
        input.addEventListener('input', function() {
            const currentValue = getNumericValue(this);
            const oldValue = parseInt(this.getAttribute('data-old-value') || '0');
            
            this.setAttribute('data-old-value', currentValue.toString());
            
            const isIncreasing = currentValue > oldValue;
            
            if (isIncreasing) {
                const allowByOBL = allowInputIncrease(this, currentValue);
                
                if (!allowByOBL) {
                    setNumericValue(this, oldValue);
                    updateOBLValidationDisplay(this);
                    return;
                }
            }
            
            if (currentValue === 0) {
                this.className = this.className.replace('revenue-input-with-value', 'revenue-input-empty');
            } else {
                this.className = this.className.replace('revenue-input-empty', 'revenue-input-with-value');
            }

            updateOBLValidationDisplay(this);
        });
    });
    
     // OBL input listener
  oblLastMonthInput.addEventListener('input', function() {
    // Remove existing toast
    const existingToast = document.querySelector('.revenue-obl-error-toast');
    if (existingToast) existingToast.remove();
    
    // Update validation display
    updateOBLValidationDisplay(this);
    
    // Show appropriate toast based on current state
    const oblValue = getNumericValue(oblLastMonthInput);
    const totalCurrent = calculateTotalCurrentValue();
    
    if (oblValue === 0 && totalCurrent === 0) {
      showOBLToast('info', `ℹ️ OBL is zero. No budget limit applied.`);
    } else if (oblValue > 0 && totalCurrent <= oblValue) {
      const remaining = oblValue - totalCurrent;
      const percentage = Math.round((totalCurrent / oblValue) * 100);
      
      if (totalCurrent === oblValue) {
        showOBLToast('success', `✅ OBL FULLY UTILIZED: ${totalCurrent.toLocaleString('en-US')} = ${oblValue.toLocaleString('en-US')}`);
      } else if (percentage > 90) {
        showOBLToast('warning', `⚠️ OBL nearly reached: ${percentage}% used. ${remaining.toLocaleString('en-US')} remaining.`);
      } else {
        showOBLToast('success', `✅ OBL within limit: ${percentage}% used. ${remaining.toLocaleString('en-US')} remaining.`);
      }
    } else if (oblValue > 0 && totalCurrent > oblValue) {
      const exceedAmount = totalCurrent - oblValue;
      showOBLToast('danger', `❌ OBL EXCEEDED by ${exceedAmount.toLocaleString('en-US')}! Total: ${totalCurrent.toLocaleString('en-US')} > OBL: ${oblValue.toLocaleString('en-US')}`);
    }
  });

  // ===== GLOBAL INPUT LISTENER FOR ALL CURRENT INPUTS - FIXED =====
  // Use event delegation for better performance
  document.addEventListener('input', function(e) {
    const target = e.target as HTMLElement;
    
    // Check if the input is one we care about
    if (target.matches('.revenue-item-current-input, .revenue-main-current-input, .revenue-yearly-current-input-rev, #oblLastMonth')) {
      // Small delay to let the value update
      setTimeout(() => {
        // Remove existing toast
        const existingToast = document.querySelector('.revenue-obl-error-toast');
        if (existingToast) existingToast.remove();
        
        // Calculate values
        const oblValue = getNumericValue(oblLastMonthInput);
        const totalCurrent = calculateTotalCurrentValue();
         
        // Update display
        totalCurrentValueSpan.textContent = totalCurrent.toLocaleString('en-US');
        const remaining = oblValue - totalCurrent;
        remainingValueSpan.textContent = remaining.toLocaleString('en-US');
        
        // Update line fill 
        if (oblValue > 0) {
          const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
          oblLineFill.style.width = `${percentage}%`;
          oblLineFill.className = 'revenue-line-fill';
          if (percentage > 90) oblLineFill.classList.add('danger');
          else if (percentage > 75) oblLineFill.classList.add('warning');
        } else {
          oblLineFill.style.width = '0%';
          oblLineFill.className = 'revenue-line-fill';
        }
        
        // Show toast based on state
        if (oblValue === 0) {
          if (totalCurrent > 0) {
            showOBLToast('warning', `⚠️ OBL is zero but values exist (${totalCurrent.toLocaleString('en-US')})`);
          } else {
            showOBLToast('info', `ℹ️ OBL is zero. No budget limit.`);
          }
        } else if (totalCurrent > oblValue) {
          showOBLToast('danger', `❌ OBL EXCEEDED by ${(totalCurrent - oblValue).toLocaleString('en-US')}! Total: ${totalCurrent.toLocaleString('en-US')} > OBL: ${oblValue.toLocaleString('en-US')}`);
        } else if (totalCurrent === oblValue) {
          showOBLToast('success', `✅ OBL FULLY UTILIZED: ${totalCurrent.toLocaleString('en-US')} = ${oblValue.toLocaleString('en-US')}`);
        } else {
          const pct = Math.round((totalCurrent / oblValue) * 100);
          const remainingAmt = oblValue - totalCurrent;
          if (pct > 90) {
            showOBLToast('warning', `⚠️ OBL nearly reached: ${pct}% used. ${remainingAmt.toLocaleString('en-US')} remaining.`);
          } else if (pct > 75) {
            showOBLToast('info', `📊 OBL usage: ${pct}%. ${remainingAmt.toLocaleString('en-US')} remaining.`);
          } else {
            showOBLToast('success', `✅ OBL within limit: ${pct}% used. ${remainingAmt.toLocaleString('en-US')} remaining.`);
          }
        }
      }, 50);
    }
  }, true);

// Also trigger validation when any input changes
document.querySelectorAll<HTMLInputElement>('.revenue-item-current-input:not([readonly]), .revenue-main-current-input:not([readonly]), .revenue-yearly-current-input-rev:not([readonly])').forEach(input => {
  input.addEventListener('input', function() {
    // Small delay to allow value to be set
    setTimeout(() => {
      updateOBLValidationDisplay(this);
    }, 50);
  });
});

// ===== OBL REAL-TIME VALIDATION =====
// Force OBL validation on any input change using event delegation

// 1. Remove any existing OBL toast and show fresh on any input change
document.addEventListener('input', function(e) {
  const target = e.target as HTMLElement;
  
  // Check if the input is one we care about
  if (target.matches('.revenue-item-current-input, .revenue-main-current-input, .revenue-yearly-current-input-rev, #oblLastMonth')) {
    // Small delay to let the value update
    setTimeout(() => {
      // Remove existing toast
      const existingToast = document.querySelector('.revenue-obl-error-toast');
      if (existingToast) existingToast.remove();
      
      // Calculate values
      const oblValue = getNumericValue(oblLastMonthInput);
      const totalCurrent = calculateTotalCurrentValue();
      
      // Update display
      totalCurrentValueSpan.textContent = totalCurrent.toLocaleString('en-US');
      const remaining = oblValue - totalCurrent;
      remainingValueSpan.textContent = remaining.toLocaleString('en-US');
      
      // Update line fill
      if (oblValue > 0) {
        const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
        oblLineFill.style.width = `${percentage}%`;
        oblLineFill.className = 'revenue-line-fill';
        if (percentage > 90) oblLineFill.classList.add('danger');
        else if (percentage > 75) oblLineFill.classList.add('warning');
      } else {
        oblLineFill.style.width = '0%';
        oblLineFill.className = 'revenue-line-fill';
      }
      
      // Show toast based on state
      if (oblValue === 0) {
        if (totalCurrent > 0) {
          showOBLToast('warning', `⚠️ OBL is zero but values exist (${totalCurrent.toLocaleString('en-US')})`);
        } else {
          showOBLToast('info', `ℹ️ OBL is zero. No budget limit.`);
        }
      } else if (totalCurrent > oblValue) {
        showOBLToast('danger', `❌ OBL EXCEEDED by ${(totalCurrent - oblValue).toLocaleString('en-US')}! Total: ${totalCurrent.toLocaleString('en-US')} > OBL: ${oblValue.toLocaleString('en-US')}`);
      } else if (totalCurrent === oblValue) {
        showOBLToast('success', `✅ OBL FULLY UTILIZED: ${totalCurrent.toLocaleString('en-US')} = ${oblValue.toLocaleString('en-US')}`);
      } else {
        const pct = Math.round((totalCurrent / oblValue) * 100);
        const remainingAmt = oblValue - totalCurrent;
        if (pct > 90) {
          showOBLToast('warning', `⚠️ OBL nearly reached: ${pct}% used. ${remainingAmt.toLocaleString('en-US')} remaining.`);
        } else if (pct > 75) {
          showOBLToast('info', `📊 OBL usage: ${pct}%. ${remainingAmt.toLocaleString('en-US')} remaining.`);
        } else {
          showOBLToast('success', `✅ OBL within limit: ${pct}% used. ${remainingAmt.toLocaleString('en-US')} remaining.`);
        }
      }
    }, 50);
  }
}, true);


        document.querySelectorAll<HTMLInputElement>('.revenue-item-current-input:not([readonly])').forEach(input => {
            input.addEventListener('input', function() {
                const monthKey = this.dataset.month;
                if (monthKey) {
                    validateSingleMonthItems(monthKey, this);
                }
            });
        });

const calculateMonthTotals = (monthKey: string): void => {
    const itemCurrentInputs = document.querySelectorAll<HTMLInputElement>(
        `input.revenue-item-current-input[data-month="${monthKey}"]`
    );
    const itemForecastInputs = document.querySelectorAll<HTMLInputElement>(
        `input.revenue-item-forecast-input[data-month="${monthKey}"]`
    );
    
    let totalCurrent = 0;
    let totalForecast = 0;
    
    itemCurrentInputs.forEach(input => {
        totalCurrent += getNumericValue(input);
    });
    
    itemForecastInputs.forEach(input => {
        totalForecast += getNumericValue(input);
    });

    // Update Curr. Plan main input
    const mainCurrentInput = document.querySelector<HTMLInputElement>(
        `input.revenue-main-current-input[data-month="${monthKey}"]`
    );
    if (mainCurrentInput) {
        setNumericValue(mainCurrentInput, totalCurrent);
        mainCurrentInput.dispatchEvent(new Event('input'));
    }
    
    // ✅ Update main forecast input (READONLY - calculated from items)
    const mainForecastInput = document.querySelector<HTMLInputElement>(
        `input.revenue-main-forecast-input[data-month="${monthKey}"]`
    );
    if (mainForecastInput) {
        setNumericValue(mainForecastInput, totalForecast);
        mainForecastInput.dispatchEvent(new Event('input'));
    }
};
   document.querySelectorAll<HTMLInputElement>('.revenue-item-current-input:not([readonly]), .revenue-item-forecast-input:not([readonly])').forEach(input => {
    input.addEventListener('input', function() {
      const monthKey = this.dataset.month;
      if (monthKey) {
        calculateMonthTotals(monthKey);
        validateSingleMonthItems(monthKey, this);
        updateOBLValidationDisplay(this);
      }
    });
  });
        const validationInputs = document.querySelectorAll<HTMLInputElement>(`
            .revenue-main-current-input:not([readonly]),
            .revenue-item-current-input:not([readonly]),
            .revenue-yearly-current-input-rev:not([readonly]),
            .revenue-main-forecast-input:not([readonly]),
            .revenue-item-forecast-input:not([readonly])
        `);

        validationInputs.forEach(input => {
            input.addEventListener('input', function() {
                setTimeout(() => {
                    updateOBLValidationDisplay(this);
                }, 100);
            });
        });
      };
 
const syncMainDateFromItems = (monthKey: string): void => {
    const itemDateInputs = document.querySelectorAll<HTMLInputElement>(
        `.revenue-item-date-input[data-month="${monthKey}"]`
    );
    const mainDateInput = document.querySelector<HTMLInputElement>(
        `input.revenue-forecast-date-input[data-month="${monthKey}"]`
    );
    if (!mainDateInput) return;

    let maxDate = '';
    itemDateInputs.forEach(inp => {
        if (inp.value && (!maxDate || inp.value > maxDate)) {
            maxDate = inp.value;
        }
    });

    // If no item date is set, default to today's system date (YYYY-MM-DD)
    if (!maxDate) {
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        maxDate = `${yyyy}-${mm}-${dd}`;
    }

    mainDateInput.value = maxDate;
    if (maxDate) {
        mainDateInput.classList.remove('revenue-input-empty');
        mainDateInput.classList.add('revenue-input-with-value');
    } else {
        mainDateInput.classList.remove('revenue-input-with-value');
        mainDateInput.classList.add('revenue-input-empty');
    }
};
// ============================================================
// Updated editForecast with RevenueActual loading
// ============================================================
const editForecast = async (id: number): Promise<void> => {
    editingId = id;
    const forecast = forecasts.find(f => f.Id === id);
    if (!forecast) {
        toast("error", "Forecast not found");
        return;
    }
    
    // Get project for permission check
    const project = projects.find(p => p.Id === forecast.Project?.Id);
    
    // Check if user can view this forecast
    if (!canRead()) {
        toast("error", "❌ You don't have permission to view this forecast");
        return;
    }
    
    // 🔐 RBAC: READ_ALL roles (Admin, SC Team, Controlling) may open ANY row;
    // every other role (GH Heads, PM Managers, Engineering) may open only rows
    // of projects assigned to them.
    const assignedProjects = getAssignedProjects(projects);
    const isAssigned = assignedProjects.some(p => p.Id === forecast.Project?.Id);
    
    if (!canReadAllProjects() && !isAssigned) {
        toast("error", "❌ You don't have permission to view this project");
        return;
    }
    
    // Check if user can edit this project
    const canEdit = canEditProject(project, forecast);
    const isFrozen = forecast.Freeze || false;
    
    // ✅ Show form FIRST
    showForm(true, project, forecast);
    
    // Load remarks
    if (forecast.Remarks) {
        const remarksData = forecast.Remarks;
        if (remarksData.includes('|||')) {
          const parts = remarksData.split('|||');
            if (remarksInternal) {
                remarksInternal.value = parts[0] || '';
            }
            if (remarksSprintCell) {
                remarksSprintCell.value = parts[1] || '';
            }
        } else {
            if (remarksInternal) {
                remarksInternal.value = remarksData;
            }
        }
    }
    
    let forecastData: ForecastData = {};
    let revenueActualData: { [key: string]: number } = {};
    
    try {
        // Load ForecastData
        let forecastDataString = forecast.ForecastData || "{}";
        forecastDataString = forecastDataString.trim();
        
        if (!forecastDataString || forecastDataString === '""' || forecastDataString === "null") {
            forecastDataString = "{}";
        }
        
        forecastData = JSON.parse(forecastDataString);
        
        // Load RevenueActual separately
        let revenueActualString = forecast.RevenueActual || "{}";
        revenueActualString = revenueActualString.trim();
        
        if (!revenueActualString || revenueActualString === '""' || revenueActualString === "null") {
            revenueActualString = "{}";
        }
        
        revenueActualData = JSON.parse(revenueActualString);
        console.log("📊 Loaded RevenueActual in editForecast:", revenueActualData);
        
    } catch (e) {
        forecastData = createFreshForecastStructure();
        revenueActualData = {};
        toast("warning", "Corrupted forecast data detected. Loaded empty structure.");
    }
    
    // Apply progression
    forecastData = propagateMonthValues(forecastData);
    
    if (forecast.Project) {
        projectSelect.value = String(forecast.Project.Id);
        onProjectChange(forecast.Project.Id, forecast);
    }
    
    // ============================================================
    // ✅ Apply freeze state to form after rendering
    // ============================================================
    setTimeout(() => {
        // If frozen OR user can't edit, set read-only for most fields
        const isReadOnly = isFrozen || !canEdit;
        setFormReadOnly(isReadOnly);
        
        // ✅ SHOW SAVE BUTTONS (top + bottom) - shown if user can edit the project
        const saveBtns = document.querySelectorAll<HTMLButtonElement>('button.revenue-save-btn');
        saveBtns.forEach(saveBtn => {
            // Only hide if user doesn't have edit permission at all
            saveBtn.style.display = canEdit ? '' : 'none';
            
            // Change button text and style based on freeze state
            if (isFrozen) {
                saveBtn.textContent = '💾 Save Further Changes';
                saveBtn.title = 'Save only Further field changes (Curr. Plan is revenue-locked)';
                saveBtn.style.backgroundColor = '#28a745';
                saveBtn.style.borderColor = '#28a745';
                saveBtn.style.color = '#ffffff';
                saveBtn.style.fontWeight = 'bold';
                saveBtn.style.cursor = 'pointer';
                saveBtn.setAttribute('type', 'submit');
            } else {
                saveBtn.textContent = '💾 Save Revenue Plan';
                saveBtn.title = '';
                saveBtn.style.backgroundColor = '';
                saveBtn.style.borderColor = '';
                saveBtn.style.color = '';
                saveBtn.style.fontWeight = '';
                saveBtn.style.cursor = '';
                saveBtn.setAttribute('type', 'submit');
            }
        });
        
           const formTitleEl = document.querySelector("#formTitle") as HTMLElement;
if (formTitleEl) {
    // const projectName = project?.Title || forecast?.Title || forecast?.Project?.Title || "";
    const projectNameShort = "Revenue Plan 5QRF";

    if (isFrozen) { 
        formTitleEl.innerHTML = `🔒 ${projectNameShort} (Locked By Admin)`;
        formTitleEl.style.color = '#856404';
    } else if (!canEdit) {
        formTitleEl.innerHTML = `🔒 ${projectNameShort} (Read Only)`;
        formTitleEl.style.color = '#856404';
    } else {
        formTitleEl.innerHTML = projectNameShort;
        formTitleEl.style.color = '';
    }
}
        // ✅ Show/hide OBL section based on freeze state
        const oblContainer = document.querySelector('.revenue-obl-section-container') as HTMLElement;
        if (oblContainer) {
            if (isFrozen) {
                oblContainer.style.opacity = '0.6';
                oblContainer.style.pointerEvents = 'none';
            } else {
                oblContainer.style.opacity = '1';
                oblContainer.style.pointerEvents = '';
            }
        }
        
        // ✅ Disable project selection when frozen
        if (projectSelect) {
            projectSelect.disabled = isFrozen || true; // Always disabled in edit mode
        }
        
        // Show toast notification about frozen state
        if (isFrozen) {
            showOBLToast('info', '🔒 Project is frozen. Only Further and Date fields are editable. Click "Save Further Changes" to save.');
        } else if (!canEdit) {
            showOBLToast('info', '🔒 You have view-only access to this project.');
        } else {
            showOBLToast('success', '🔓 All fields are editable. Click "Save Revenue Plan" to save.');
        }
        
    }, 300);
     
}; 

 
 
 // ============================================================
// ✅ SAVE — SweetAlert confirmation FIRST (same look as procurement.ts / dci.ts)
//    • 2 buttons only: Yes, Save / Cancel  (no "Don't Save" option)
//    • Title + body adapt to the mode:  New / Edit / Frozen
//    • z-index bumped so the loader (.revenue-loader / .loading-overlay)
//      can never cover the dialog
//    • isSaving guard prevents double-submit
//    • saveButtons hoisted BEFORE the outer try, so the outer finally
//      ALWAYS restores them — even if Swal.fire itself throws
// ============================================================

const escapeHtml = (unsafe: string | number | null | undefined): string => {
    if (unsafe === null || unsafe === undefined) return "";
    return String(unsafe)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
};

let isSaving = false;

form.addEventListener("submit", async (e: SubmitEvent): Promise<void> => {
    e.preventDefault();
    e.stopPropagation();

    // ✅ Only run the flow when a Save button triggered the submit
    const submitter = (e as any).submitter as HTMLButtonElement | null;
    const isSaveClick = submitter && submitter.classList.contains("revenue-save-btn");
    if (!isSaveClick) return;

    // ---- Guard: only block while a save is actually in flight ----
    if (isSaving) {
        console.log("⛔ submit blocked: save already in progress");
        return;
    }

    // ---- Pre-checks (nothing shown yet) ----
    const projectId = parseInt(projectSelect.value, 10);
    const project = projects.find(p => p.Id === projectId);

    if (!project) {
        toast("error", "Please select a project");
        return;
    }

    const existingForecast = editingId
        ? forecasts.find(f => f.Id === editingId)
        : undefined;

    if (editingId) {
        if (!canEditProject(project, existingForecast)) {
            toast("error", "❌ You don't have permission to edit this forecast");
            return;
        }
    } else {
        if (!canCreateForecast()) {
            toast("error", "❌ You don't have permission to create forecasts");
            return;
        }
    }

    const isFrozen = existingForecast?.Freeze === true;
    const isNewProject =
        !editingId &&
        !forecasts.some(f => f.Project?.Id === projectId);

    if (
        !isNewProject &&
        forecasts.some(f => f.Project?.Id === projectId && f.Id !== editingId)
    ) {
        toast("error", "Forecast already exists for this project!");
        return;
    }

    // ✅ HOIST: query the buttons ONCE, before the outer try,
    //    so the outer finally can always restore them.
    const saveButtons = form.querySelectorAll<HTMLButtonElement>(
        'button[type="submit"]'
    );

    isSaving = true;
    try {
        // ========== 1) CONFIRMATION POPUP ==========
        const confirmResult = await Swal.fire({
            title: isFrozen
                ? "Save Further Changes?"
                : editingId
                    ? "Save Changes?"
                    : "Save Revenue Plan?",
            html: isFrozen
                ? `Do you want to save the <strong>Further</strong> field changes for <strong>${escapeHtml(project.Title)}</strong>?`
                : editingId
                    ? `Do you want to save the changes made to <strong>${escapeHtml(project.Title)}</strong>?`
                    : `Do you want to save the revenue plan for <strong>${escapeHtml(project.Title)}</strong>?`,
            icon: "question",
            showCancelButton: true,
            confirmButtonText: isFrozen
                ? "Yes, Save Further"
                : editingId
                    ? "Yes, Save Changes"
                    : "Yes, Save",
            cancelButtonText: "Cancel",
            confirmButtonColor: "#605eee",
            cancelButtonColor: "#6c757d",
            reverseButtons: true,
            focusCancel: true,
            allowOutsideClick: false,
            customClass: { popup: "proc-confirm-popup" },
            didOpen: () => {
                const c = document.querySelector(".swal2-container") as HTMLElement | null;
                if (c) c.style.zIndex = "2147483647";
            },
        });

        if (!confirmResult.isConfirmed) {
            console.log("🔵 save cancelled by user");
            return; // finally will reset isSaving + buttons
        }

        // ========== 2) OBL VALIDATION (skip when frozen) ==========
        if (!isFrozen) {
            const oblValue = getNumericValue(oblLastMonthInput);

            if (isNewProject && oblValue === 0) {
                const confirmZeroOBL = await Swal.fire({
                    icon: "question",
                    title: "Zero OBL Value",
                    html: "You're setting OBL to zero for a new project.<br>This means there's no budget limit. Continue?",
                    showCancelButton: true,
                    confirmButtonText: "Yes, Continue",
                    cancelButtonText: "No, Set OBL",
                    confirmButtonColor: "#605eee",
                    cancelButtonColor: "#6c757d",
                    reverseButtons: true,
                    focusCancel: true,
                    customClass: { popup: "proc-confirm-popup" },
                    didOpen: () => {
                        const c = document.querySelector(".swal2-container") as HTMLElement | null;
                        if (c) c.style.zIndex = "2147483647";
                    },
                });

                if (!confirmZeroOBL.isConfirmed) {
                    oblLastMonthInput.focus();
                    return;
                }
            }

            const totalCurrent = calculateTotalCurrentValue();
            if (oblValue > 0 && totalCurrent > oblValue) {
                toast(
                    "warning",
                    `⚠️ Total current value exceeds OBL by ${(totalCurrent - oblValue).toLocaleString('en-US')} - saving anyway`
                );
            }
        }

        // ========== 3) LOCK SAVE BUTTONS ==========
        saveButtons.forEach(saveButton => {
            saveButton.disabled = true;
            saveButton.dataset.saving = "true";
            saveButton.textContent = isFrozen ? "💾 Saving Further..." : "💾 Saving...";
        });

        // ========== 4) ACTUAL SAVE (logic unchanged) ==========
        let forecastData: ForecastData = {};
        let revenueActualData: { [key: string]: number } = {};

        if (existingForecast) {
            const existingData = safeJsonParse(existingForecast.ForecastData, {});
            forecastData = JSON.parse(JSON.stringify(existingData));
            revenueActualData = safeJsonParse(existingForecast.RevenueActual, {});
        }

        if (Object.keys(forecastData).length === 0) {
            forecastData = createFreshForecastStructure();
        }

        const allStorageMonths = getAllMonthsForStorage();
        const currentMonthKey = formatMonthKey(new Date());
        const currentMonthIndex = allStorageMonths.indexOf(currentMonthKey);

        let detailedMonths: string[] = [];
        if (currentMonthIndex !== -1) {
            detailedMonths = allStorageMonths.slice(
                currentMonthIndex,
                Math.min(currentMonthIndex + 3, allStorageMonths.length)
            );
        } else {
            detailedMonths = allStorageMonths.slice(36, 39);
        }

        // ---- PROCESS EACH MONTH ----
        allStorageMonths.forEach((monthKey: string) => {
            const existingMonthData = forecastData[monthKey] as ForecastMonthData || {
                plan: 0, actual: 0, monthActual: 0, forecast: 0, date: "", items: {}
            };

            const monthData: ForecastMonthData = {
                plan: existingMonthData.plan || 0,
                actual: existingMonthData.actual || 0,
                monthActual: existingMonthData.monthActual || 0,
                forecast: existingMonthData.forecast || 0,
                date: existingMonthData.date || "",
                items: existingMonthData.items
                    ? JSON.parse(JSON.stringify(existingMonthData.items))
                    : {}
            };

            if (isFrozen) {
                // FROZEN MODE: only Further + Date are editable
                const mainForecastInput = document.querySelector<HTMLInputElement>(
                    `input.revenue-main-forecast-input[data-month="${monthKey}"]`
                );
                const mainDateInput = document.querySelector<HTMLInputElement>(
                    `input.revenue-forecast-date-input[data-month="${monthKey}"]`
                );

                if (!monthData.items) monthData.items = {};

                let totalItemForecast = 0;
                for (let itemNum = 1; itemNum <= 10; itemNum++) {
                    const itemKey = itemNum.toString();
                    const itemForecastInput = document.querySelector<HTMLInputElement>(
                        `.revenue-item-forecast-input[data-item="${itemKey}"][data-month="${monthKey}"]`
                    );
                    const itemDateInput = document.querySelector<HTMLInputElement>(
                        `.revenue-item-date-input[data-item="${itemKey}"][data-month="${monthKey}"]`
                    );

                    if (monthData.items[itemKey]) {
                        if (itemForecastInput) {
                            const v = getNumericValue(itemForecastInput);
                            monthData.items[itemKey].forecast = v;
                            totalItemForecast += v;
                        }
                        if (itemDateInput) {
                            monthData.items[itemKey].date =
                                itemDateInput.value || monthData.items[itemKey].date || "";
                        }
                    }
                }

                monthData.forecast = totalItemForecast;
                if (mainForecastInput) setNumericValue(mainForecastInput, totalItemForecast);
                if (mainDateInput) monthData.date = mainDateInput.value || monthData.date;

            } else {
                // UNFROZEN MODE: update everything
                const mainPlanInput = document.querySelector<HTMLInputElement>(
                    `input.revenue-main-plan-input[data-month="${monthKey}"]`
                );
                const mainCurrentInput = document.querySelector<HTMLInputElement>(
                    `input.revenue-main-current-input[data-month="${monthKey}"]`
                );
                const mainForecastInput = document.querySelector<HTMLInputElement>(
                    `input.revenue-main-forecast-input[data-month="${monthKey}"]`
                );
                const mainDateInput = document.querySelector<HTMLInputElement>(
                    `input.revenue-forecast-date-input[data-month="${monthKey}"]`
                );

                if (mainPlanInput) monthData.plan = getNumericValue(mainPlanInput);
                if (mainCurrentInput) monthData.actual = getNumericValue(mainCurrentInput);

                if (detailedMonths.includes(monthKey)) {
                    if (!monthData.items) monthData.items = {};

                    let totalItemCurrent = 0;
                    let totalItemForecast = 0;

                    for (let itemNum = 1; itemNum <= 10; itemNum++) {
                        const itemKey = itemNum.toString();
                        const itemNameInput = document.querySelector<HTMLInputElement>(
                            `.revenue-item-description-input[data-item="${itemKey}"][data-month="${monthKey}"]`
                        );
                        const itemPlanInput = document.querySelector<HTMLInputElement>(
                            `.revenue-item-plan-input[data-item="${itemKey}"][data-month="${monthKey}"]`
                        );
                        const itemCurrentInput = document.querySelector<HTMLInputElement>(
                            `.revenue-item-current-input[data-item="${itemKey}"][data-month="${monthKey}"]`
                        );
                        const itemForecastInput = document.querySelector<HTMLInputElement>(
                            `.revenue-item-forecast-input[data-item="${itemKey}"][data-month="${monthKey}"]`
                        );
                        const itemDateInput = document.querySelector<HTMLInputElement>(
                            `.revenue-item-date-input[data-item="${itemKey}"][data-month="${monthKey}"]`
                        );

                        if (itemNameInput || itemPlanInput || itemCurrentInput || itemForecastInput || itemDateInput) {
                            const existingItem = monthData.items[itemKey] || {
                                name: "", plan: 0, actual: 0, forecast: 0, date: ""
                            };

                            const itemCurrentValue = itemCurrentInput
                                ? getNumericValue(itemCurrentInput)
                                : existingItem.actual || 0;
                            const itemForecastValue = itemForecastInput
                                ? getNumericValue(itemForecastInput)
                                : existingItem.forecast || 0;

                            totalItemCurrent += itemCurrentValue;
                            totalItemForecast += itemForecastValue;

                            monthData.items[itemKey] = {
                                name: itemNameInput ? itemNameInput.value || "" : existingItem.name || "",
                                plan: itemPlanInput ? getNumericValue(itemPlanInput) : existingItem.plan || 0,
                                actual: itemCurrentValue,
                                forecast: itemForecastValue,
                                date: itemDateInput ? itemDateInput.value || "" : existingItem.date || ""
                            };
                        }
                    }

                    monthData.actual = totalItemCurrent;
                    if (mainCurrentInput) setNumericValue(mainCurrentInput, totalItemCurrent);

                    monthData.forecast = totalItemForecast;
                    if (mainForecastInput) setNumericValue(mainForecastInput, totalItemForecast);

                } else {
                    if (mainCurrentInput) monthData.actual = getNumericValue(mainCurrentInput);
                    if (mainForecastInput) monthData.forecast = getNumericValue(mainForecastInput);
                }

                if (mainDateInput) monthData.date = mainDateInput.value || monthData.date;
            }

            forecastData[monthKey] = monthData;
        });

        // ---- PROCESS YEARLY DATA (only when not frozen) ----
        if (!isFrozen) {
            const yearlyPlanInputs = document.querySelectorAll<HTMLInputElement>(".revenue-yearly-plan-input-rev");
            const existingYearly = forecastData.yearly || {};

            if (yearlyPlanInputs.length > 0) {
                const updatedYearlyData: YearlyForecastData = { ...existingYearly };
                const years = new Set<string>();

                yearlyPlanInputs.forEach((input: HTMLInputElement) => {
                    const year = input.dataset.year;
                    if (year) years.add(year);
                });

                Array.from(years).sort().forEach((year: string) => {
                    const planInput = document.querySelector<HTMLInputElement>(
                        `.revenue-yearly-plan-input-rev[data-year="${year}"]`
                    );
                    const currentInput = document.querySelector<HTMLInputElement>(
                        `.revenue-yearly-current-input-rev[data-year="${year}"]`
                    );

                    if (planInput && currentInput) {
                        updatedYearlyData[year] = {
                            plan: getNumericValue(planInput),
                            actual: getNumericValue(currentInput)
                        };
                    }
                });

                forecastData.yearly = updatedYearlyData;
            } else if (existingYearly && Object.keys(existingYearly).length > 0) {
                forecastData.yearly = existingYearly;
            }
        }

        // ---- APPLY PROPAGATION (only when not frozen) ----
        let finalData = forecastData;
        if (!isFrozen) {
            finalData = propagateCollectionValues(forecastData);
        }

        // ---- BUILD PAYLOAD ----
        const digest = await getDigest();

        const internalRemark = remarksInternal?.value || "";
        const sprintCellRemark = remarksSprintCell?.value || "";
        const combinedRemarks = internalRemark + "|||" + sprintCellRemark;

        const forecastDataString = JSON.stringify(finalData);

        // ✅ Preserve RevenueActual — DO NOT overwrite with plan values
        let revenueActualString = JSON.stringify(revenueActualData);

        if (revenueActualString.length > 255) {
            console.warn("⚠️ RevenueActual too long:", revenueActualString.length);
            const cleanedRevenueActual: { [key: string]: number } = {};
            Object.keys(revenueActualData).forEach((key: string) => {
                const value = revenueActualData[key];
                if (value > 0 && /^\d{4}-\d{2}$/.test(key)) {
                    cleanedRevenueActual[key] = value;
                }
            });
            revenueActualString = JSON.stringify(cleanedRevenueActual);
        }

        const payload: any = {
            __metadata: { type: "SP.Data.Revenue_x0020_Forecast_x0020_ListListItem" },
            Title: generateForecastTitle(project, finalData),
            ProjectId: project.Id,
            ProfitCentre: project.PC || "",
            ForecastData: forecastDataString,
            RevenueActual: revenueActualString,
            Remarks: combinedRemarks,
            UpdatedDate: new Date().toISOString()
        };

        if (!editingId && isNewProject) {
            payload.OBL = getNumericValue(oblLastMonthInput);
        }
        if (!editingId) {
            payload.CreatedDate = new Date().toISOString();
        }

        const url = editingId
            ? `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${editingId})`
            : `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

        const headers: Record<string, string> = {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest
        };

        if (editingId) {
            headers["IF-MATCH"] = "*";
            headers["X-HTTP-Method"] = "MERGE";
        }

        const response = await fetch(url, {
            method: "POST",
            headers,
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("❌ SharePoint save failed:", response.status, errorText);
            toast("error", `Error saving forecast: ${response.status}`);
            return;
        }

        toast(
            "success",
            isFrozen
                ? "✅ Further changes saved successfully!"
                : editingId
                    ? "✅ Revenue Plan updated successfully!"
                    : "✅ Revenue Plan saved successfully!"
        );

        await loadForecasts();
        filterTable();
        initializePagination(filteredData.length ? filteredData : forecasts);

        hideForm();
        resetForm();

    } catch (outerError) {
        console.error("Unexpected error in revenue submit:", outerError);
        const message = outerError instanceof Error ? outerError.message : String(outerError);
        toast("error", `Unexpected error: ${message}`);
    } finally {
        // ✅ Always restore the save buttons + clear the guard —
        //    no path (success, cancel, error, or a thrown Swal) can leave
        //    them stuck in the "Saving..." state.
        saveButtons.forEach(saveButton => {
            saveButton.disabled = false;
            saveButton.dataset.saving = "false";
            saveButton.textContent = isFrozen
                ? "💾 Save Further Changes"
                : "💾 Save Revenue Plan";
        });
        isSaving = false;
    }
});
 
const getSafeString = (value: any, defaultValue: string = "-"): string => {
    if (value === null || value === undefined) {
        return defaultValue;
    }
    if (typeof value === 'string') {
        return value.trim() || defaultValue;
    }
    if (typeof value === 'object' && 'Title' in value) {
        return value.Title?.trim() || defaultValue;
    }
    return String(value) || defaultValue;
};
 
// Add this near the other event listeners
// const backTopBtn = document.querySelector('.revenue-back-top-btn') as HTMLButtonElement;
// if (backTopBtn) {
//     backTopBtn.addEventListener('click', async () => {
//         const confirmBack = await Swal.fire({
//             icon: "question",
//             title: "Unsaved Changes",
//             text: "Do you want to save the changes?",
//             showCancelButton: true,
//             showDenyButton: true,
//             confirmButtonText: "Yes, Save",
//             denyButtonText: "No, Don't Save",
//             cancelButtonText: "Cancel"
//         });

//         if (confirmBack.isDismissed) {
//             // Cancel - stay on form
//             return;
//         }

//         if (confirmBack.isDenied) {
//             // No, Don't Save - go back without saving
//             hideForm();
//             resetForm();
//             return;
//         }

//         // Yes, Save - trigger form submit
//         form.requestSubmit();
//     });
// }

// ============================================================
// BACK BUTTONS (Top and Bottom) - Confirmation Popup ONLY
// ============================================================
 

 // ============================================================
// ✅ BACK — SweetAlert confirmation, same look as procurement.ts
//    • Only one prompt: Leave (Yes/Stay) — no Save/Don't Save/Cancel
//    • Only shown when a form is actually open (add or edit)
//    • z-index bumped so the revenue-loader / loading-overlay cannot
//      cover it.
// ============================================================
let isLeavingForm = false;

const handleBackAction = async (e: Event): Promise<void> => {
  e.preventDefault();
  e.stopPropagation();

  if (isLeavingForm) {
    console.log("⛔ back blocked: already leaving");
    return;
  }

  // Nothing to confirm when the form is hidden
  const formVisible = formSection && formSection.style.display !== "none";
  if (!formVisible) {
    hideForm();
    resetForm();
    return;
  }

  isLeavingForm = true;
  try {
    const confirmResult = await Swal.fire({
      title: "Leave this page?",
      html:
        "Any <strong>unsaved changes</strong> will be lost.<br>" +
        "Do you want to go back to the revenue list?",
      icon: "warning",
      showCancelButton: true,
      confirmButtonText: "Yes, Go Back",
      cancelButtonText: "Cancel",
      confirmButtonColor: "#605eee",
      cancelButtonColor: "#6c757d",
      reverseButtons: true,
      focusCancel: true,
      allowOutsideClick: false,
      customClass: { popup: "proc-confirm-popup" },
      didOpen: () => {
        const c = document.querySelector(".swal2-container") as HTMLElement | null;
        if (c) c.style.zIndex = "2147483647";
      },
    });

    if (!confirmResult.isConfirmed) {
      console.log("🔵 back cancelled by user");
      return;
    }

    hideForm();
    resetForm();
  } catch (err) {
    console.error("Error while leaving the form:", err);
    const message = err instanceof Error ? err.message : String(err);
    toast("error", `Unexpected error: ${message}`);
  } finally {
    isLeavingForm = false;
  }
};

// Attach to the Bottom Back Button
backBtn?.addEventListener("click", handleBackAction);

// Attach to the Top Back Button
const backTopBtn = document.querySelector(".revenue-back-top-btn") as HTMLButtonElement | null;
if (backTopBtn) {
  backTopBtn.addEventListener("click", handleBackAction);
}

const initializeFilters = (): void => {
  const currentYear = new Date().getFullYear();

  // ✅ "check filter": Filters start CLEAN - the placeholder option is selected
  //    and nothing is pre-picked (no default Year/Month, badge stays 0).
  filterYear.innerHTML = '<option value="">All Years</option>';
  for (let year = currentYear - 1; year <= currentYear + 2; year++) {
    filterYear.innerHTML += `<option value="${year}">${year}</option>`;
  }

  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  filterMonth.innerHTML = '<option value="">All Months</option>';
  for (let month = 1; month <= 12; month++) {
    filterMonth.innerHTML += `<option value="${month}">${monthNames[month - 1]}</option>`;
  }

  const assignedProjectTitles = getAssignedProjectValues('Title');
  const assignedCustomers = getAssignedProjectValues('CustomerName');
   const assignedManagers = getAssignedProjectValues('ProjectManager');
  const assignedGroupHeads = getAssignedProjectValues('GroupHead1');
  const assignedReferences = getAssignedProjectValues('ProjectReference');

  filterProject.innerHTML = '<option value="">All Projects</option>' + 
    assignedProjectTitles.map(project => `<option value="${project}">${project}</option>`).join("");

  filterCustomer.innerHTML = '<option value="">All Customers</option>' + 
    assignedCustomers.map(customer => `<option value="${customer}">${customer}</option>`).join("");

  filterProjectManager.innerHTML = '<option value="">All Managers</option>' + 
    assignedManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");

  filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' + 
    assignedGroupHeads.map(head => `<option value="${head}">${head}</option>`).join("");

  filterProjectReference.innerHTML = '<option value="">All References</option>' + 
    assignedReferences.map(ref => `<option value="${ref}">${ref}</option>`).join("");

  // ✅ nothing is pre-selected - mirror the applied (empty) filter state
  restoreAppliedFilters();
  syncFilterUI();
};

const filterTable = (): void => {
  // ✅ the table is filtered by the APPLIED filters (Apply / chips), never by
  //    whatever happens to sit in the popup right now
  const year = appliedFilters.year;
  const month = appliedFilters.month;
  const project = appliedFilters.project;
  const customer = appliedFilters.customer;
  const projectManager = appliedFilters.manager;
  const groupHead = appliedFilters.groupHead;
  const projectReference = appliedFilters.reference; 
  
  let filtered = [...forecasts];

  const assignedProjects = getAssignedProjects(projects);
  const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
  
  // 🔐 RBAC: READ_ALL roles (Admin, SC Team, Controlling) see ALL rows;
  // every other role (GH Heads, PM Managers, Engineering) sees only rows of
  // projects assigned to them.
  if (!canReadAllProjects()) {
    filtered = filtered.filter(forecast => {
      const projectTitle = forecast.Project?.Title;
      return projectTitle && assignedProjectTitles.has(projectTitle);
    });
  }

  filtered = filtered.filter(forecast => {
    const projectTitle = forecast.Project?.Title;
    return projectTitle && projectTitle.trim() !== '';
  });

  filtered = filtered.filter(forecast => { 
    if (year && month) {
      const monthKey = `${year}-${month.padStart(2, "0")}`;
      const data: ForecastData = safeJsonParse(forecast.ForecastData, {});
      if (!(monthKey in data)) return false;
    }
    
    if (project && forecast.Project?.Title !== project) return false;
    
    if (customer) {
      const proj = projects.find(p => p.Title === forecast.Project?.Title);
      if (proj?.CustomerName !== customer) return false;
    }
    
    if (projectManager) {
      const proj = projects.find(p => p.Title === forecast.Project?.Title);
      const managerName = getProjectManagerDisplayName(proj?.ProjectManager);
      if (managerName !== projectManager) return false;
    }
    
    if (groupHead) {
      const proj = projects.find(p => p.Title === forecast.Project?.Title);
      const groupHeadName = getGroupHeadDisplayName(proj?.GroupHead1);
      if (groupHeadName !== groupHead) return false;
    }
    
    if (projectReference) {
      const proj = projects.find(p => p.Title === forecast.Project?.Title);
      if (proj?.ProjectReference !== projectReference) return false;
    }
    
    return true;
  }); 
  
  if (pageSizeSelect) {
    pageSize = parseInt(pageSizeSelect.value) || 10;
  }
  
  if (filtered.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:20px;">No records found for selected filters</td></tr>`;
    
    const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;
    if (totalsHeader) totalsHeader.innerHTML = '';
    
    filteredData = [];
    currentPage = 1;
    calculateTotalPages();
    updatePaginationControls();
    renderCurrentPage();
  } else {
    filteredData = filtered;
    currentPage = 1;
    calculateTotalPages();
    updatePaginationControls();
    renderCurrentPage();
    
    const displayMonths = getDisplayMonths();
    const currentMonthKey = formatMonthKey(new Date());
    setTimeout(() => {
      renderTotalsHeader(displayMonths, currentMonthKey);
    }, 100);
  }
};
const checkIfNeedsProgression = (forecastData: ForecastData): boolean => {
    const now = new Date();
    const currentMonthKey = formatMonthKey(now);
    
    const months = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
    
    if (months.length === 0) return true;
    
    const lastDataMonth = months[months.length - 1];
    
    const idealMonths = generateIdealMonthStructure();
    const lastIdealMonth = idealMonths[idealMonths.length - 1];
    
    const hasCurrentMonth = months.includes(currentMonthKey);
    
    if (!hasCurrentMonth) {
        return true;
    }
    
    const monthlyProgressionNeeded = lastDataMonth < lastIdealMonth;
    
    const currentYear = now.getFullYear();
    let yearlyProgressionNeeded = false;
    
    if (forecastData.yearly) {
        const years = Object.keys(forecastData.yearly).map(y => parseInt(y)).sort();
        yearlyProgressionNeeded = years.length > 0 && years[0] < currentYear;
    } else {
        yearlyProgressionNeeded = true;
    }
    
    return monthlyProgressionNeeded || yearlyProgressionNeeded;
};
 
const propagateCollectionValues = (forecastData: ForecastData): ForecastData => {
    const updatedData: ForecastData = { ...forecastData };
    
    if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) {
        const longTermMonths = generateIdealMonthStructure().slice(6);
        const lastMonthKey = longTermMonths[longTermMonths.length - 1];
        updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
    }

    const originalMonths = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
    const idealMonths = generateIdealMonthStructure();
    const isNewMonth = originalMonths[0] !== idealMonths[0];
     
    const monthlyProgressed = propagateMonthValues(updatedData);
    const yearlyProgressed = propagateYearlyValues(monthlyProgressed, isNewMonth); 
    return yearlyProgressed;
};

const propagateYearlyValues = (forecastData: ForecastData, isNewMonth: boolean = false): ForecastData => {
    if (!isNewMonth) { 
        return forecastData;
    }

    const updatedData: ForecastData = { ...forecastData };
    
    if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) { 
        return updatedData;
    }

    const yearlyData = updatedData.yearly; 
    const updatedYearlyData: YearlyForecastData = {};

    Object.keys(yearlyData).forEach(year => {
        const yearData = yearlyData[year];
        if (yearData) {
            updatedYearlyData[year] = {
                plan: yearData.actual,
                actual: 0
            };
        }
    });

    updatedData.yearly = updatedYearlyData; 
    return updatedData;
};

     const importActualDatas = (): void => { 
  const modalHTMLRev = `
    <div class="revenue-import-modal" style="display: none;">
      <div class="revenue-import-modal-content">
        <div class="revenue-import-modal-header">
          <h3>Import Monthly Actual Revenue Data</h3>
          <button class="revenue-close-modal">&times;</button>
        </div>

        <div class="revenue-import-section">
          <h4>Step 1: Download Template</h4>
          <button class="revenue-btn-primary" id="downloadTemplate">⬇️ Download Template</button>
        </div>

        <div class="revenue-import-section">
          <h4>Step 2: Upload Excel File</h4>
          <div class="revenue-file-upload-area" id="fileUploadArea">
            <div class="revenue-upload-icon">📁</div>
            <p>Click to upload or drag & drop Excel file</p>
            <p><small>Supported format: Excel with Project Reference, Month, and Actual Revenue columns</small></p>
            <input type="file" id="excelFile" class="revenue-file-input" accept=".xlsx,.xls">
            <div class="revenue-file-info" id="fileInfo"></div>
          </div>
        </div>

        <div class="revenue-import-section" id="previewSection" style="display:none;">
          <h4>Step 3: Preview & Import</h4>
          <div class="revenue-import-preview" id="importPreview"></div>
          <div class="revenue-import-actions">
            <button class="revenue-btn-secondary" id="cancelImport">Cancel</button>
            <button class="revenue-btn-primary" id="confirmImport">Import Data</button>
          </div>
          <div class="revenue-import-status" id="importStatus"></div>
        </div>
      </div>
    </div>
  `;

  const existingModal = document.querySelector('.revenue-import-modal');
  if (existingModal) {
    existingModal.remove();
  }
  
  document.body.insertAdjacentHTML("beforeend", modalHTMLRev);
  const modal = document.querySelector(".revenue-import-modal") as HTMLElement;
  
  setTimeout(() => {
    modal.style.display = "flex";
  }, 10);

  const fileUploadArea = document.querySelector("#fileUploadArea") as HTMLElement;
  const excelFileInput = document.querySelector("#excelFile") as HTMLInputElement;
  const fileInfo = document.querySelector("#fileInfo") as HTMLElement;
  const previewSection = document.querySelector("#previewSection") as HTMLElement;
  const importPreview = document.querySelector("#importPreview") as HTMLElement;
  const confirmImport = document.querySelector("#confirmImport") as HTMLButtonElement;
  const cancelImport = document.querySelector("#cancelImport") as HTMLButtonElement;
  const closeModal = document.querySelector(".revenue-close-modal") as HTMLButtonElement;
  const importStatus = document.querySelector("#importStatus") as HTMLElement;
  const downloadTemplateBtn = document.querySelector("#downloadTemplate") as HTMLButtonElement;

  let importData: any[] = [];
  let listEntityType: string = "";

  const monthMap: { [key: string]: number } = {
    jan: 1, january: 1, "jan-": 1,
    feb: 2, february: 2, "feb-": 2,
    mar: 3, march: 3, "mar-": 3,
    apr: 4, april: 4, "apr-": 4,
    may: 5, "may-": 5,
    jun: 6, june: 6, "jun-": 6,
    jul: 7, july: 7, "jul-": 7,
    aug: 8, august: 8, "aug-": 8,
    sep: 9, september: 9, "sep-": 9,
    oct: 10, october: 10, "oct-": 10,
    nov: 11, november: 11, "nov-": 11,
    dec: 12, december: 12, "dec-": 12,
  };

  const parseMonth = (monthStr: string): string => {
    if (!monthStr) return "";
    if (/^\d{4}-\d{2}$/.test(monthStr)) return monthStr;

    const match = monthStr.toLowerCase().match(/([a-z]+)[-\s]*'?(\d{2,4})?/);
    if (match) {
      const monthName = match[1];
      let year = match[2];

      const monthNum = monthMap[monthName];
      if (monthNum) {
        if (!year) {
          year = new Date().getFullYear().toString();
        } else if (year.length === 2) {
          const currentYear = new Date().getFullYear();
          const century = Math.floor(currentYear / 100) * 100;
          const yearNum = parseInt(year);
          year = (century + yearNum).toString();
        }
        return `${year}-${monthNum.toString().padStart(2, "0")}`;
      }
    }
    return "";
  };

  const getListEntityTypeForImport = async (): Promise<string> => {
    if (listEntityType) return listEntityType;
    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
      headers: { Accept: "application/json;odata=verbose" },
    });
    const data = await res.json();
    listEntityType = data.d.ListItemEntityTypeFullName; 
    return listEntityType;
  };

  fileUploadArea.addEventListener("click", () => excelFileInput.click());
  fileUploadArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    fileUploadArea.classList.add("revenue-dragover");
  });
  fileUploadArea.addEventListener("dragleave", () => fileUploadArea.classList.remove("revenue-dragover"));
  fileUploadArea.addEventListener("drop", (e) => {
    e.preventDefault();
    fileUploadArea.classList.remove("revenue-dragover");
    const files = e.dataTransfer?.files;
    if (files && files[0]) {
      excelFileInput.files = files;
      handleFileSelection(files[0]);
    }
  });
  excelFileInput.addEventListener("change", (e) => {
    const files = (e.target as HTMLInputElement).files;
    if (files && files[0]) handleFileSelection(files[0]);
  });

  const handleFileSelection = async (file: File) => { 
    if (!file.name.toLowerCase().match(/\.(xlsx|xls)$/)) {
      toast("error", "Please select an Excel file (.xlsx or .xls)");
      return;
    }

    fileInfo.textContent = `Selected: ${file.name} (${(file.size / 1024).toFixed(2)} KB)`;

    try {
      const parsedData = await parseExcelFile(file); 
      const validation = validateImportData(parsedData);
      if (validation.valid) { 
        importData = parsedData;
        showImportPreview(parsedData);
        previewSection.style.display = "block";
        previewSection.style.opacity = "1"; 
      } else { 
        toast("error", `Invalid Excel format. Missing required column(s): ${validation.missing.join(", ")}. Expected columns: Project Reference, Month, Actual Revenue.`);
      }
    } catch (error) { 
      toast("error", "Error reading Excel file");
    }
  };

  const parseExcelFile = (file: File): Promise<any[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];

          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); 
          if (jsonData.length < 2) return resolve([]);

          const headers = jsonData[0] as string[];  
          const rows: any[] = [];

          for (let i = 1; i < jsonData.length; i++) {
            const rowData = jsonData[i] as any[];
            if (!rowData || rowData.length === 0) continue;

            const row: any = {};
            headers.forEach((header, idx) => {
              if (header && rowData[idx] !== undefined && rowData[idx] !== null) {
                row[header.trim()] = rowData[idx];
              }
            });

            if (row["Month"] || row["month"]) {
              let monthValue = row["Month"] || row["month"];
              if (typeof monthValue === "number") {
                const excelDate = new Date(Math.round((monthValue - 25569) * 86400 * 1000));
                const yyyy = excelDate.getFullYear();
                const mm = String(excelDate.getMonth() + 1).padStart(2, '0');
                monthValue = `${yyyy}-${mm}`;
              }
              row.parsedMonth = parseMonth(monthValue.toString()); 
            }

            const hasProjectRef = row["Project Reference"] || row["project_reference"];
            const hasMonth = row.parsedMonth;

            if (hasProjectRef && hasMonth) {
              rows.push(row); 
            }
          } 
          resolve(rows);
        } catch (err) { 
          reject(err);
        }
      };
      reader.onerror = (error) => reject(new Error("Failed to read file"));
      reader.readAsArrayBuffer(file);
    });  
  };

  const validateImportData = (data: any[]): { valid: boolean; missing: string[] } => { 
    if (data.length === 0) return { valid: false, missing: ["No data rows found"] };

    const first = data[0];
    const keys = Object.keys(first).map((k) => k.toLowerCase()); 
    const hasProjectRef = keys.some((k) => k.includes("project") && k.includes("reference"));
    const hasMonth = keys.some((k) => k.includes("month"));
    const hasRevenue = keys.some((k) => k.includes("actual") && (k.includes("revenue") || k.includes("rev")));

    const missing: string[] = [];
    if (!hasProjectRef) missing.push("Project Reference");
    if (!hasMonth) missing.push("Month");
    if (!hasRevenue) missing.push("Actual Revenue");

    return { valid: missing.length === 0, missing };
  }; 

  const showImportPreview = (data: any[]) => { 
    const monthSummary: { [key: string]: number } = {};
    data.forEach((row) => {
      const month = row.parsedMonth;
      if (month) {
        monthSummary[month] = (monthSummary[month] || 0) + 1;
      }
    });

    const monthSummaryText = Object.keys(monthSummary)
      .map((month) => `${month}: ${monthSummary[month]} records`)
      .join(", ");

    const previewHTML = `
      <div style="padding: 20px; background: #f8f9fa; border-radius: 8px;">
        <h4 style="color: #28a745; margin-bottom: 15px;">✅ File Ready for Import</h4>
        <p><strong>Found ${data.length} records to import:</strong></p>
        <p><strong>Month Summary:</strong> ${monthSummaryText || "No months found"}</p>
        <div class="revenue-preview-table-container" style="max-height: 400px; overflow-y: auto;">
          <table class="revenue-preview-table" style="width: 100%; border-collapse: collapse;">
            <thead style="background: #007bff; color: white;">
              <tr>
                <th style="padding: 10px; border: 1px solid #ddd;">Project Reference</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Month</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Actual Revenue</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Project Match</th>
              </tr>
            </thead>
            <tbody>
              ${data.slice(0, 10).map((row, index) => {
                const projectRef = row["Project Reference"] || row["project_reference"] || "N/A";
                const month = row["Month"] || row["month"] || "N/A";
                const actualRevenue = row["Actual Revenue"] || row["actual_revenue"] || row["Actual"] || row["actual"] || "0";
                const project = projects.find((p) => p.ProjectReference === projectRef);

                return `
                  <tr style="${index % 2 === 0 ? "background: #f8f9fa;" : "background: white;"}">
                    <td style="padding: 8px; border: 1px solid #ddd;">${projectRef}</td>
                    <td style="padding: 8px; border: 1px solid #ddd;">${month} (${row.parsedMonth || "Invalid"})</td>
                    <td style="padding: 8px; border: 1px solid #ddd;">${actualRevenue}</td>
                    <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${project ? "✅" : "❌"}</td>
                  </tr>
                `;
              }).join("")}
            </tbody>
          </table>
        </div>
        ${data.length > 10 ? `<p style="margin-top: 10px; font-style: italic;">... and ${data.length - 10} more records</p>` : ""}
        <div style="margin-top: 20px; padding: 15px; background: #e7f3ff; border-radius: 5px;">
          <p style="margin: 0; color: #0056b3;"><strong>Ready to import?</strong> Click the "Import Data" button below.</p>
          <p style="margin: 5px 0 0 0; font-size: 12px;"><strong>Note:</strong> Both the RevenueActual column and ForecastData JSON will be updated.</p>
        </div>
      </div>`;

    importPreview.innerHTML = previewHTML; 
  };

  const performImport = async () => {
    confirmImport.disabled = true;
    confirmImport.textContent = "Importing...";
    importStatus.style.display = "block";
    importStatus.textContent = "Updating Revenue Data...";

    showRevenueLoader('Importing Data', `Please wait... updating ${importData.length} record(s)`);

    let success = 0, error = 0, notFound = 0, invalidMonth = 0, skipped = 0;
    const entityType = await getListEntityTypeForImport(); 

    for (const row of importData) {
      const projectRef = row["Project Reference"] || row["project_reference"];
      const month = row.parsedMonth;
      const actualRevenue = parseFloat(row["Actual Revenue"] || row["actual_revenue"] || row["Actual"] || row["actual"] || "0");

      if (!month) { 
        invalidMonth++;
        continue;
      }

      const project = projects.find((p) => p.ProjectReference === projectRef);
      if (!project) { 
        notFound++;
        continue;
      }

      const forecast = forecasts.find((f) => f.Project?.Title === project.Title);
      if (!forecast?.Id) { 
        notFound++;
        continue;
      }

      try {
        const digest = await getDigest(); 
        
        // ✅ FETCH BOTH COLUMNS TO PRESERVE EXISTING DATA
        const currentForecastResponse = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${forecast.Id})?$select=Id,RevenueActual,ForecastData`, {
          method: "GET",
          headers: { Accept: "application/json;odata=verbose" },
        });

        if (!currentForecastResponse.ok) {
          throw new Error(`Failed to fetch current data: ${currentForecastResponse.status}`);
        }

        const currentData = await currentForecastResponse.json();
        
        // 1. Update RevenueActual column
        const currentRevenueData = currentData.d.RevenueActual || "{}"; 
        const updatedRevenueData = updateRevenueData(currentRevenueData, month, actualRevenue); 

        // 2. Update ForecastData JSON (monthActual)
        const currentForecastDataStr = currentData.d.ForecastData || "{}";
        let currentForecastData: any = {};
        try {
          currentForecastData = JSON.parse(currentForecastDataStr);
        } catch(e) { currentForecastData = {}; }

        if (!currentForecastData[month]) {
          currentForecastData[month] = { plan: 0, actual: 0, forecast: 0, monthActual: 0, items: {} };
        }
        
        // ✅ Sync the monthActual inside the JSON
        currentForecastData[month].monthActual = actualRevenue;

        // ✅ PREPARE PAYLOAD
        const payload = {
          __metadata: { type: entityType },
          RevenueActual: updatedRevenueData,
          ForecastData: JSON.stringify(currentForecastData),
          UpdatedDate: new Date().toISOString(),
        }; 

        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${forecast.Id})`, {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify(payload),
        });

        if (res.ok) {
          success++; 
        } else {
          error++;
          const text = await res.text();
          console.warn(`⚠️ Failed for ${projectRef} (${month}):`, text);
        }

        await new Promise((r) => setTimeout(r, 100));
      } catch (err) {
        error++;
        console.error(`❌ Error updating ${projectRef} (${month}):`, err);
      }
    }

    skipped = importData.length - (success + error + notFound + invalidMonth);

    importStatus.className = "revenue-import-status success";
    importStatus.innerHTML = `
      <strong>Import Complete!</strong>
      ✅ Success: ${success} records imported
      ❌ Errors: ${error} records
      🔍 Project Not Found: ${notFound} records
      📅 Invalid Month: ${invalidMonth} records
      ⏭️ Skipped: ${skipped} records
      <small>Only the imported month was updated, all other data preserved</small>
    `;

    await loadForecasts();
    initializePagination(forecasts);

    hideRevenueLoader();

    confirmImport.disabled = false;
    confirmImport.textContent = "Import Data";
  };

  downloadTemplateBtn.addEventListener("click", () => {
    if (!projects || projects.length === 0) {
      toast("error", "No projects available to create template");
      return;
    }

    const wsData = [["Project Reference", "Month", "Actual Revenue"]];
    projects.forEach((p) => {
      wsData.push([p.ProjectReference || "", "", ""]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Revenue Template");

    const fileName = `Revenue_Import_Template.xlsx`;
    XLSX.writeFile(wb, fileName);
  });

  const closeModalHandler = () => {
    hideRevenueLoader();
    modal.remove();
  };
  closeModal.addEventListener("click", closeModalHandler);
  cancelImport.addEventListener("click", closeModalHandler);
  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeModalHandler();
  });

  confirmImport.addEventListener("click", performImport);
};
        const updateRevenueData = (existingData: string, month: string, actualRevenue: number): string => {
            let finalData: { [key: string]: number } = {};

            if (existingData && existingData.trim() !== "" && existingData !== "{}" && existingData !== "null" && existingData !== "undefined") {
                try {
                    const parsed = JSON.parse(existingData);
                    finalData = { ...parsed };
                } catch (parseError) {
                    finalData = {};
                }
            } else {
                finalData = {};
            }

            if (actualRevenue > 0) {
                finalData[month] = actualRevenue;
            } else {
                delete finalData[month];
            }

            const result: { [key: string]: number } = {};
            Object.keys(finalData).forEach(key => {
                if (finalData[key] > 0) {
                    result[key] = finalData[key];
                }
            });

            return JSON.stringify(result);
        };
const initialize = async (): Promise<void> => {
    try {
        showRevenueLoader('Loading Revenue Data', 'Initializing permissions...');
        await authManager.initialize({ siteUrl });

        console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
        console.log(`👤 User: ${getCurrentUserFullName()}`);
        console.log(`📋 User Groups:`, getUserGroups());
        console.log(`✅ Has Full Control: ${hasFullControl()}`);
        console.log(`✅ Can Contribute: ${canContribute()}`);
        console.log(`✅ Can Freeze: ${canFreeze()}`);
        console.log(`✅ Is GH Head: ${isGHHead()}`);

        showRevenueLoader('Loading Revenue Data', 'Loading projects and forecasts...');
        const projectsLoad = loadProjects();
        const forecastsLoad = loadForecasts(projectsLoad);
        await Promise.all([projectsLoad, forecastsLoad]);
        
        console.log("🚀 Total forecasts loaded:", forecasts.length);

        // Show the manual Sync button only to roles that can create rows
        const revenueSyncBar = document.querySelector('#revenueSyncBar') as HTMLElement | null;
        if (revenueSyncBar) {
          revenueSyncBar.style.display = canCreateForecast() ? 'flex' : 'none';
        }
        
        if (pageSizeSelect) {
            pageSize = parseInt(pageSizeSelect.value) || 10;
        }
        
        initializeFilters();
        
        setupPaginationEventListeners();
        
        filterTable();
        
        updateUIBasedOnPermissions();
        
        if (canImportData()) {
            importActualBtns.addEventListener("click", importActualDatas);
        }

        const runBackgroundAutoUpdate = async (): Promise<void> => {
            try {
                setBackgroundAutoUpdateIndicator(".revenue-page", true);
                await syncMissingProjectsToRevenueList();

                if (forecasts.length > 0) {
                    if (canEditOwn()) {
                        await applyAutoProgression();
                    } else {
                        forecasts = forecasts.map((forecast: RevenueForecastItem) => {
                            const data = safeJsonParse(forecast.ForecastData, {});
                            if (!checkIfNeedsProgression(data)) return forecast;
                            return { ...forecast, ForecastData: JSON.stringify(propagateCollectionValues(data)) };
                        });
                        console.log('👁️ Read-only viewer: progression applied locally for display only (no SharePoint writes)');
                    }
                }
            } catch (error) {
                console.error("Background Revenue update failed:", error);
                toast("error", "Revenue background update failed");
            } finally {
                try {
                    filterTable();
                } catch (error) {
                    console.error("Could not refresh the Revenue table after background update:", error);
                    toast("error", "Could not refresh the Revenue table");
                } finally {
                    setBackgroundAutoUpdateIndicator(".revenue-page", false);
                }
            }
        };
        void runBackgroundAutoUpdate();
         
        console.log("✅ Application initialized successfully");
        
    } catch (error) {
        console.error("❌ Failed to initialize application:", error);
        toast("error", "Failed to initialize application");
      } finally {
        // hide the section loader (success AND failure paths)
        hideRevenueLoader();
    }
};
      // ========== EVENT LISTENERS ==========

       
           
      
      projectSelect.addEventListener("change", (e) => {
        if (editingId) {
          toast("warning", "Cannot change project in edit mode");
          const forecast = forecasts.find(f => f.Id === editingId);
          if (forecast?.Project) {
            projectSelect.value = String(forecast.Project.Id);
          }
          return;
        }

                const projectId = parseInt((e.target as HTMLSelectElement).value);
        if (projectId) {
          const project = projects.find(p => p.Id === projectId);
          if (project) {
            const existingForecast = forecasts.find(f => f.Project?.Id === projectId);
            
            // ✅ Just open the form directly without the "Existing Forecast Found" popup
            if (existingForecast && !editingId) {
              editingId = existingForecast.Id;
              showForm(true);
              onProjectChange(projectId, existingForecast);
            } else {
              onProjectChange(projectId, existingForecast);
            }
          }
        }
      });

      // Filter popup + chips + Apply pill + Reset Filter: ONE set of bindings
      // (the table only re-filters when Apply is pressed). The pager is wired
      // by setupPaginationEventListeners() inside initialize().
      setupFilterEventListeners();

      // Action button event listeners
      exportBtn.addEventListener("click", exportToExcelWithFullStyle);
      exportDetailsBtn.addEventListener("click", exportDetailedItemsToExcel);
      const exportHistoricalBtn = document.querySelector(".revenue-export-historical-btn") as HTMLButtonElement;
      if (exportHistoricalBtn) {
        exportHistoricalBtn.addEventListener("click", exportHistoricalToExcel);
      }
      const revenueSyncProjectsBtn = document.querySelector("#revenueSyncProjectsBtn") as HTMLButtonElement;
      if (revenueSyncProjectsBtn) {
        revenueSyncProjectsBtn.addEventListener("click", () => {
          syncMissingProjectsToRevenueList();
        });
      }
      clearFiltersBtn.addEventListener("click", clearFilters);
 window.addEventListener('resize', () => {
  clearTimeout((window as any)._revStickyResizeTimer);
  (window as any)._revStickyResizeTimer = setTimeout(applyStickyHeaderOffsets, 150);
});
      // Initialize the application      
      initialize();
    }, 100);
  }   
     )}
}; 

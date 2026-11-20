declare const _spPageContextInfo: any;
import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css"; 
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/collection.css";

import {
  authManager,
  canRead,
  canReadAllProjects,
  canFreeze,
  canEditOwn,
  getUserGroups,
  getCurrentUserFullName,
  getPermissionLevel,
  canEditProject as authCanEditProject,
  canCreateRecord,
  getAssignedProjects,
  updateUIBasedOnPermissions,
  getDigest as getAuthDigest,
  getListEntityType,
  hasFullControl  
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

declare global {
  interface Window {
    _headerRendered?: boolean;
  }
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

interface CollectionItem {
  Id: number;
  Title: string;
  Project?: {
    Id: number;
    Title: string;
  };
  ProfitCentre?: string;
  OBL?: number;
  CollectionData?: string;
  RevenueActual?: string;
  Remarks?: string | null;
  CreatedDate?: string;
  UpdatedDate?: string;
  Freeze?: boolean;
  NormalDue?: number;
  NormalOverdue?: number;
  RetentionDue?: number;
  RetentionOverdue?: number;
  GSTTDS?: number;
  WHT?: number;
  _revenueActual?: { [key: string]: number };
  [key: string]: any;
}

interface OBLBreakdown {
  normalDue: number;
  normalOverdue: number;
  retentionDue: number;
  retentionOverdue: number;
  gstTds: number;
  wht: number;
}

interface SharePointResponse<T> {
  d: {
    results: T[];
  };
}

interface CollectionCategory {
  name: string;
  plan: number;
  actual: number;
  forecast?: number;
  date?: string;
}

interface CollectionMonthData {
    plan: number;
    actual: number;
    monthActual?: number;
    forecast?: number;
    date?: string;
    items?: {
        [categoryName: string]: CollectionCategory;
    };
}

interface YearlyCollectionData {
  [year: string]: {
    plan: number;
    actual: number;
    items?: {
      [categoryName: string]: CollectionCategory;
    };
  };
}

interface CollectionData {
  [monthKey: string]: CollectionMonthData | YearlyCollectionData | undefined;
  yearly?: YearlyCollectionData;
}

export default { 
  content: `
<div class="collection-page fade-in">
  <section class="collection-section collection-card">
    <div class="page-header"> 
      <h1>Collection Plan 5QRF</h1>
        <button type="button" class="collection-show-form-btn" title="Add a new collection">+ Add Collection</button>
    </div> 

    <!-- ✅ TOOLBAR ABOVE THE TABLE (inventory style)
         LEFT  = Filters (funnel icon + active count badge) + the compact
                 "Apply" pill that appears when the popup has unapplied changes
         RIGHT = Upload Actual / Download Data (Summary) / Download Historical Data
         The page header (title) is untouched - the toolbar never sits next to it. -->
    <div class="collection-table-toolbar">
      <div class="collection-toolbar-left">
        <button type="button" class="collection-btn-secondary collection-filter-trigger" id="collectionFilterTrigger" aria-haspopup="dialog" aria-expanded="false" title="Filter Collection data">
          <svg class="collection-filter-funnel" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8v5l-4 2v-7L3 5z"/></svg><span>Filters</span><span class="collection-filter-count" id="collectionFilterCount">0</span><svg class="collection-filter-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>
        </button>
        <!-- appears only while the popup has unapplied changes -->
        <button type="button" class="collection-btn-primary collection-filter-applypill" id="collectionFilterApplyPill" title="Apply the selected filters">Apply</button>
      </div>
      <div class="collection-table-toolbar-actions">
        <button class="collection-btn-secondary collection-import-actual-btn">📥 Upload Actual</button>
        <button class="collection-btn-secondary collection-export-btn">📊 Download Data (Summary)</button>
        <button class="collection-btn-secondary collection-export-historical-btn">📊 Download Historical Data</button>
      </div>
    </div>
    <!-- applied filters show up here as removable chips -->
    <div class="collection-filter-summary" id="collectionFilterChips" hidden></div>

    <div class="collection-table-actions collection-filter-overlay" id="collectionFilterOverlay" role="dialog" aria-modal="true" aria-labelledby="collectionFilterTitle">
      <div class="collection-filter-modal">
        <div class="collection-filter-modal-head"><h2 id="collectionFilterTitle">Filters</h2><button type="button" id="collectionFilterClose" aria-label="Close">&times;</button></div>
      <div class="collection-table-actions-row">
        <div class="collection-tableColp4">
          <label for="filterProject">Project</label>
          <select id="filterProject" class="collection-filter-select"><option value="">All Projects</option></select>
        </div>
        <div class="collection-tableColp4">
          <label for="filterCustomer">Customer</label>
          <select id="filterCustomer" class="collection-filter-select"><option value="">All Customers</option></select>
        </div>
        <div class="collection-tableColp4">
          <label for="filterProjectManager">Project Manager</label>
          <select id="filterProjectManager" class="collection-filter-select"><option value="">All Managers</option></select>
        </div>
        <div class="collection-tableColp4">
          <label for="filterGroupHead">Group Head</label>
          <select id="filterGroupHead" class="collection-filter-select"><option value="">All Group Heads</option></select>
        </div>
        <div class="collection-tableColp4">
          <label for="filterProjectReference">Project Reference</label>
          <select id="filterProjectReference" class="collection-filter-select"><option value="">All References</option></select>
        </div>
        <div class="collection-tableColp4">
          <label> </label>
          <button class="collection-btn-secondary collection-clear-filters-btn" id="collectionResetFilterBtn" title="Clear All Filters" disabled>Reset Filter</button>
        </div>
      </div>

      <div class="collection-table-actions-row">
        <div class="collection-tableColp4">
          <label for="filterYear">Year</label>
          <select id="filterYear" class="collection-filter-select"><option value="">All Years</option></select>
        </div>
        <div class="collection-tableColp4">
          <label for="filterMonth">Month</label>
          <select id="filterMonth" class="collection-filter-select"><option value="">All Months</option></select>
        </div>
        <div class="collection-tableColp4 collection-filter-dupe-actions">
          <label> </label>
          <button class="collection-btn-secondary collection-import-actual-btn">📥 Upload Actual</button>
        </div>
        <div class="collection-tableColp4 collection-filter-dupe-actions">
          <label> </label>
          <button class="collection-btn-secondary collection-export-btn">📊 Download Data (Summary)</button>
        </div>
        <div class="collection-tableColp4 collection-filter-dupe-actions">
    <label> </label> 
    <button class="collection-btn-secondary collection-export-historical-btn">📊 Download Historical Data</button>
      </div>
    </div>
        <div class="collection-filter-modal-foot">
          <button type="button" class="collection-btn-secondary" id="collectionFilterCancel">Cancel</button>
          <button type="button" class="collection-btn-secondary collection-filter-btn-clear" id="collectionFilterClear" title="Clear all filters" disabled>Clear</button>
          <button type="button" class="collection-btn-primary collection-filter-btn-apply" id="collectionFilterApply" disabled>Apply</button>
        </div>
      </div>
    </div>

    <div id="collectionSyncBar" style="display:none; margin:6px 0 10px 0; align-items:center; gap:10px; flex-wrap:wrap;">
      <button type="button" class="collection-btn-secondary collection-sync-projects-btn" id="collectionSyncProjectsBtn" title="Compare Project List with Collection List and add any missing project row">➕ Sync Projects</button>
      <span id="collectionSyncStatus" style="font-size:12px; color:#6c757d;"></span>
    </div>

    <div class="collection-table-container">
      <div class="collection-month-columns-table">
        <table class="collection-month-columns-table">
          <thead id="totalsHeader"></thead>
          <thead id="monthColumnsHeader"></thead>
          <tbody class="collection-table-body" aria-busy="true">
            <!-- ✅ Bar loader (grey track + indigo bar, no text) while the
                 SharePoint rows are fetched. -->
            <tr class="collection-table-loading-row">
              <td colspan="28">
                <div class="collection-table-loading-content" role="status" aria-live="polite">
                  <div class="collection-page-loader">
                    <div class="collection-page-loader-track"><div class="collection-page-loader-bar"></div></div>
                  </div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      
    </div>
    <!-- ✅ PAGINATION - one flat bar, same text + layout as the mock-up
         LEFT  : Items per page: [n]   Showing 1 - 25 of 60 items
         RIGHT : « ‹ [tiles] of 6 › » -->
    <div class="collection-pagination-container">
      <div class="collection-pagination-left">
        <div class="collection-pagination-size-selector collection-pg-size">
          <label for="pageSize">Items per page:</label>
          <select id="pageSize" class="collection-page-size-select">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>
        <div class="collection-pagination-info collection-pg-total">
          Showing <span class="collection-pg-range"><span id="startItem">0</span> - <span id="endItem">0</span></span> of <span class="collection-pg-count" id="totalItems">0</span> items
        </div>
      </div>
      <div class="collection-pagination-right">
        <button class="collection-pagination-btn collection-pg-icon" id="firstPage" disabled title="First page">«</button>
        <button class="collection-pagination-btn collection-pg-icon" id="prevPage" disabled title="Previous page">‹</button>
        <div class="collection-pagination-pages" id="pageNumbers"></div>
        <span class="collection-pg-of">of <span id="totalPages">1</span></span>
        <button class="collection-pagination-btn collection-pg-icon" id="nextPage" disabled title="Next page">›</button>
        <button class="collection-pagination-btn collection-pg-icon" id="lastPage" disabled title="Last page">»</button>
      </div>

  </section>

 <section class="collection-form-section collection-card" style="display:none;">
   <form id="collectionForm">

  <div class="collection-page-header">
    <h1 id="formTitle">Collection Plan 5QRF</h1> 
    <button type="submit"
            form="collectionForm"
            class="collection-btn-primary collection-btn-save-top"
            title="Save Collection Plan">💾 Save Collection Plan</button>
             <button type="button" class="collection-btn-secondary collection-back-to-table">⬅ Back</button>
  </div>

      <div class="collection-form-row collection-hidden-section">
        <div class="collection-form-group collection-project-select-container">
          <label>Project</label>
          <select id="projectSelect" required></select>
        </div>
      </div>

      <div class="collection-project-main-wrapper">
        <div class="collection-project-left-side" id="projectDetails">
          <div class="collection-detail-item">
            <span class="collection-detail-label">Project:</span>
            <span class="collection-detail-value" id="projectNameDisplay">-</span>
          </div>
          <div class="collection-detail-item">
            <span class="collection-detail-label">Project Ref:</span>
            <span class="collection-detail-value" id="projectRef">-</span>
          </div>
          <div class="collection-detail-item">
            <span class="collection-detail-label">Profit Center:</span>
            <span class="collection-detail-value" id="profitCenter">-</span>
          </div>
          <div class="collection-detail-item">
            <span class="collection-detail-label">Project Manager:</span>
            <span class="collection-detail-value" id="projectManager">-</span>
          </div>
          <div class="collection-detail-item">
            <span class="collection-detail-label">Group Head:</span>
            <span class="collection-detail-value" id="groupHead">-</span>
          </div>
          <div class="collection-detail-item">
            <span class="collection-detail-label">Customer:</span>
            <span class="collection-detail-value" id="customer">-</span>
          </div>
        </div>

        <div class="collection-project-right-side">
          <div class="collection-remarks-group">
            <label for="remarksInternal">REMARKS (INTERNAL)</label>
            <textarea id="remarksInternal" class="collection-remarks-textarea" rows="3" placeholder="Enter internal remarks..."></textarea>
          </div>

          <div class="collection-remarks-group">
            <label for="remarksSprintCell">REMARKS (SPRINT CALL)</label>
            <textarea id="remarksSprintCell" class="collection-remarks-textarea" rows="3" placeholder="Enter sprint cell remarks..."></textarea>
          </div>

          <div class="collection-obl-section-container collection-hidden-section" id="oblContainer">
            <div class="collection-obl-content">
              <div class="collection-obl-left">
                <div class="collection-obl-label-group">
                  <label>Receivable Value</label>
                </div>
                <div class="collection-obl-value-display">
                  <input type="text" id="oblLastMonth" class="collection-obl-input collection-number-input" placeholder="Enter OBL" readonly />
                </div>
              </div>
              <div class="collection-obl-right">
                <div class="collection-validation-summary">
                  <div class="collection-validation-item">
                    <span class="collection-validation-label">Total Current Value:</span>
                    <span class="collection-validation-value" id="totalCurrentValue">0</span>
                  </div>
                  <div class="collection-validation-item">
                    <span class="collection-validation-label">Remaining Balance:</span>
                    <span class="collection-validation-value" id="remainingValue">0</span>
                  </div>
                  <div class="collection-validation-line">
                    <div class="collection-line-fill" id="oblLineFill"></div>
                  </div>
                </div>
              </div>
            </div>
            <div class="collection-validation-message" id="oblValidationMessage">
              ✅ All values are within Receivables limit
            </div>
          </div>
        </div>
      </div>

      <div class="collection-receivable-breakdown-section">
        <div class="collection-breakdown-header-title"></div>
        <div class="collection-breakdown-grid">
          <div class="collection-breakdown-field">
            <label>NORMAL DUE</label>
            <input type="text" id="oblNormalDue" class="collection-breakdown-input collection-number-input" readonly placeholder="0" />
          </div>
          <div class="collection-breakdown-field">
            <label>NORMAL OVERDUE</label>
            <input type="text" id="oblNormalOverdue" class="collection-breakdown-input collection-number-input" readonly placeholder="0" />
          </div>
          <div class="collection-breakdown-field">
            <label>RETENTION DUE</label>
            <input type="text" id="oblRetentionDue" class="collection-breakdown-input collection-number-input" readonly placeholder="0" />
          </div>
          <div class="collection-breakdown-field">
            <label>RETENTION OVERDUE</label>
            <input type="text" id="oblRetentionOverdue" class="collection-breakdown-input collection-number-input" readonly placeholder="0" />
          </div>
          <div class="collection-breakdown-field">
            <label>GST TDS</label>
            <input type="text" id="oblGstTds" class="collection-breakdown-input collection-number-input" readonly placeholder="0" />
          </div>
          <div class="collection-breakdown-field">
            <label>WHT</label>
            <input type="text" id="oblWht" class="collection-breakdown-input collection-number-input" readonly placeholder="0" />
          </div>
        </div>
      </div>

      <div class="collection-collection-section collection-hidden-section" id="historicalCollectionSection">
        <div class="collection-section-header" id="historicalSectionHeader" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; cursor: pointer; padding: 8px 12px; border-radius: 4px; transition: background-color 0.2s;">
          <h3 style="margin: 0; color: #495057; border-bottom: 2px solid #007bff; padding-bottom: 8px; flex: 1; pointer-events: none;">
            📊 Historical Data (Last 3 Months - Readonly)
          </h3>
          <button type="button" class="collection-toggle-historical-btn" id="historicalToggleBtn" style="font-size: 12px; padding: 4px 12px; border-radius: 4px; border: 1px solid #6c757d; background-color: #6c757d; color: white; cursor: pointer; min-width: 70px; transition: all 0.3s ease;">
            ▼ Show
        </button>
        </div>
        <div id="historicalContainer" class="collection-historical-container" style="display: none; opacity: 0; max-height: 0; overflow: hidden; transition: opacity 0.4s ease, max-height 0.4s ease, padding 0.3s ease; padding: 0;">
        </div>
      </div> 

      <div class="collection-collection-section collection-hidden-section" id="detailedCollectionSection">
        <h3>Detailed Collection (Current + Next 2 Months)</h3>
        <div id="firstRowContainer" class="collection-first-row-container"></div>
      </div>

      <div class="collection-collection-section collection-hidden-section" id="longtermCollectionSection">
        <h3>Long-term Collection (Next 16 Months)</h3>
        <div id="remainingContainer" class="collection-remaining-container"></div>
      </div>

      <div class="collection-collection-section collection-hidden-section" id="yearlyCollectionSection">
        <h3>4-Year Collection</h3>
        <div id="yearlyContainer" class="collection-yearly-container"></div>
      </div>

      <div class="collection-form-actions">
        <button type="submit" class="collection-btn-primary">💾 Save Collection Plan</button>
        <button type="button" class="collection-btn-secondary collection-back-to-table">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,


  script: () => {
    setTimeout(async () => {
      const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/Kriya";
      const listName = "Collection List";
      const projectListName = "Project List";

      const formSection = document.querySelector(".collection-form-section") as HTMLElement;
      const tableSection = document.querySelector(".collection-section") as HTMLElement;
      const showFormBtn = document.querySelector(".collection-show-form-btn") as HTMLButtonElement;
      const projectSelect = document.querySelector("#projectSelect") as HTMLSelectElement;
      const historicalContainer = document.querySelector("#historicalContainer") as HTMLElement;
      const firstRowContainer = document.querySelector("#firstRowContainer") as HTMLElement;
      const remainingContainer = document.querySelector("#remainingContainer") as HTMLElement;
      const yearlyContainer = document.querySelector("#yearlyContainer") as HTMLElement;
      const backBtn = document.querySelector(".collection-back-to-table") as HTMLButtonElement;
      const form = document.querySelector("#collectionForm") as HTMLFormElement;
      const exportBtn = document.querySelector(".collection-export-btn") as HTMLButtonElement;
      const oblLastMonthInput = document.querySelector("#oblLastMonth") as HTMLInputElement;
      const totalCurrentValueSpan = document.querySelector("#totalCurrentValue") as HTMLSpanElement;
      const remainingValueSpan = document.querySelector("#remainingValue") as HTMLSpanElement;
      const oblLineFill = document.querySelector("#oblLineFill") as HTMLDivElement;
      const oblValidationMessage = document.querySelector("#oblValidationMessage") as HTMLSpanElement;
      const formTitle = document.querySelector("#formTitle") as HTMLElement;

      const filterYear = document.querySelector("#filterYear") as HTMLSelectElement;
      const filterMonth = document.querySelector("#filterMonth") as HTMLSelectElement;
      const filterProject = document.querySelector("#filterProject") as HTMLSelectElement;
      const filterCustomer = document.querySelector("#filterCustomer") as HTMLSelectElement;
      const filterProjectManager = document.querySelector("#filterProjectManager") as HTMLSelectElement;
      const filterGroupHead = document.querySelector("#filterGroupHead") as HTMLSelectElement;
      const filterProjectReference = document.querySelector("#filterProjectReference") as HTMLSelectElement;

      const filterTriggerBtn = document.querySelector("#collectionFilterTrigger") as HTMLButtonElement;
      const filterOverlay = document.querySelector("#collectionFilterOverlay") as HTMLElement;
      const filterCloseBtn = document.querySelector("#collectionFilterClose") as HTMLButtonElement;
      const filterCancelBtn = document.querySelector("#collectionFilterCancel") as HTMLButtonElement;
      const filterClearBtn = document.querySelector("#collectionFilterClear") as HTMLButtonElement;
      const filterApplyBtn = document.querySelector("#collectionFilterApply") as HTMLButtonElement;
      const filterCount = document.querySelector("#collectionFilterCount") as HTMLElement;

      const importActualBtn = document.querySelector(".collection-import-actual-btn") as HTMLButtonElement;

      const pageSizeSelect = document.querySelector("#pageSize") as HTMLSelectElement;
      const firstPageBtn = document.querySelector("#firstPage") as HTMLButtonElement;
      const prevPageBtn = document.querySelector("#prevPage") as HTMLButtonElement;
      const nextPageBtn = document.querySelector("#nextPage") as HTMLButtonElement;
      const lastPageBtn = document.querySelector("#lastPage") as HTMLButtonElement;
      const totalPagesSpan = document.querySelector("#totalPages") as HTMLSpanElement;
      const pageNumbersDiv = document.querySelector("#pageNumbers") as HTMLDivElement;
      const startItemSpan = document.querySelector("#startItem") as HTMLSpanElement;
      const endItemSpan = document.querySelector("#endItem") as HTMLSpanElement;
      const totalItemsSpan = document.querySelector("#totalItems") as HTMLSpanElement;

      // ============================================================
// ✅ SHARED HELPERS for the Save / Update confirmation dialogs
//    (same recipe as revenue.ts / procurement.ts / dci.ts)
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

// Guards so neither flow can be re-entered by double-clicking
let isSavingCollection = false;
let isLeavingCollectionForm = false;

// Dialogs must render ABOVE every overlay of this module
const bumpSwalZIndex = (): void => {
  const c = document.querySelector(".swal2-container") as HTMLElement | null;
  if (c) c.style.zIndex = "2147483647";
};

      let projects: ProjectItem[] = [];
      let collections: CollectionItem[] = [];
      let editingId: number | null = null;
      let allMonths: string[] = [];

      let currentPage = 1;
      let pageSize = 25;
      let totalPages = 1;
      let filteredData: CollectionItem[] = [];
      // ---- Filter state (inventory style) ----------------------------------
      type CollectionFilters = Record<'year' | 'month' | 'project' | 'customer' | 'manager' | 'groupHead' | 'reference', string>;
      const EMPTY_FILTERS: CollectionFilters = { year: '', month: '', project: '', customer: '', manager: '', groupHead: '', reference: '' };
      const FILTER_LABELS: Record<keyof CollectionFilters, string> = {
        year: 'Year',
        month: 'Month',
        project: 'Project',
        customer: 'Customer',
        manager: 'Project Manager',
        groupHead: 'Group Head',
        reference: 'Project Reference'
      };
      // ✅ Filters start CLEAN on every load - nothing is pre-selected.
      let appliedFilters: CollectionFilters = { ...EMPTY_FILTERS };

      // toolbar badge / chips / floating Apply pill / popup Reset Filter
      const filterChipsHost = document.querySelector('#collectionFilterChips') as HTMLElement | null;
      const filterApplyPill = document.querySelector('#collectionFilterApplyPill') as HTMLButtonElement | null;
      const resetFilterBtn = document.querySelector('#collectionResetFilterBtn') as HTMLButtonElement | null;

      const readDraftFilters = (): CollectionFilters => ({
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

      const sameFilters = (a: CollectionFilters, b: CollectionFilters): boolean =>
        (Object.keys(EMPTY_FILTERS) as (keyof CollectionFilters)[]).every(k => (a[k] || '') === (b[k] || ''));

      const activeFilterCount = (): number => Object.values(appliedFilters).filter(Boolean).length;
      const hasDraftFilters = (): boolean => Object.values(readDraftFilters()).some(Boolean);
      const filtersAreDirty = (): boolean => !sameFilters(readDraftFilters(), appliedFilters);

      // chips for the applied filters (each removable, plus "Clear all")
      const renderFilterChips = (): void => {
        if (!filterChipsHost) return;
        const entries = (Object.keys(appliedFilters) as (keyof CollectionFilters)[])
          .filter(k => !!appliedFilters[k])
          .map(k => ({ key: k, value: appliedFilters[k] }));

        filterChipsHost.innerHTML = '';
        if (entries.length === 0) { filterChipsHost.hidden = true; return; }
        filterChipsHost.hidden = false;

        const label = document.createElement('span');
        label.className = 'collection-filter-summary__label';
        label.textContent = 'Active filters:';
        filterChipsHost.appendChild(label);

        entries.forEach(({ key, value }) => {
          const chip = document.createElement('span');
          chip.className = 'collection-filter-chip';
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
        clearAll.className = 'collection-filter-chip__all';
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
          const el = document.querySelector(`#${FILTER_FIELD_IDS[key as keyof CollectionFilters]}`) as HTMLSelectElement | null;
          el?.classList.toggle('is-set', !!draft[key as keyof CollectionFilters]);
        });

        renderFilterChips();
      };

      const FILTER_FIELD_IDS: Record<keyof CollectionFilters, string> = {
        year: 'filterYear',
        month: 'filterMonth',
        project: 'filterProject',
        customer: 'filterCustomer',
        manager: 'filterProjectManager',
        groupHead: 'filterGroupHead',
        reference: 'filterProjectReference'
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
      // ========== UTILITY FUNCTIONS ==========

      const toast = (icon: any, title: string) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });

      /* ---------- loading overlay (bar loader - image 4.png) ---------- */
      /*  Grey track + indigo bar, NO text and NO spinner: every in-code loader
          (first-paint table row + save/freeze overlay) is this one component.
          Progress messages go to the console instead of the screen.          */

      const createLoadingOverlay = (message = "Processing..."): HTMLElement => {
        removeLoadingOverlay(); // never stack two loaders
        const overlay = document.createElement("div");
        overlay.className = "loading-overlay";
        overlay.setAttribute("aria-label", message);
        overlay.innerHTML = `
          <div class="loading-container">
            <div class="collection-page-loader">
              <div class="collection-page-loader-track"><div class="collection-page-loader-bar"></div></div>
            </div>
          </div>`;
        if (!document.querySelector("#loading-styles")) {
          const style = document.createElement("style");
          style.id = "loading-styles";
          style.textContent = `
            .loading-overlay{position:absolute;top:0;left:0;width:100%;height:100%;background:rgba(255,255,255,0.94);display:flex;justify-content:center;align-items:center;z-index:99999;backdrop-filter:blur(3px);border-radius:12px;overflow:hidden;box-sizing:border-box}
            .loading-container{padding:26px 30px;background:#fff;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,0.15);width:min(520px,85vw);border:1px solid #e8e8e8;box-sizing:border-box}
            .collection-page-loader{display:block;width:100%;height:auto;background:none;border-radius:0;overflow:visible}
            .collection-page-loader-track{position:relative;width:100%;height:6px;overflow:hidden;background:#DCDCDC;border-radius:999px}
            .collection-page-loader-bar{position:absolute;top:0;left:0;width:50%;height:100%;background:#6664FF;border-radius:999px;animation:collectionPageLoaderSlide 1.25s cubic-bezier(.65,.05,.36,1) infinite}
            @keyframes collectionPageLoaderSlide{0%{transform:translateX(-100%)}100%{transform:translateX(200%)}}`;
          document.head.appendChild(style);
        }
        return overlay;
      };

      const updateLoadingMessage = (overlay: HTMLElement, message: string): void => {
        // ✅ No text on screen - the loader is the bar only. Progress goes to the console.
        console.log(`⏳ ${message}`);
        if (overlay) overlay.setAttribute("aria-label", message);
      };

      const removeLoadingOverlay = (): void => {
        document.querySelectorAll(".loading-overlay").forEach(el => el.remove());
      };

      /** mount the loader inside the visible collection card — never the whole page */
      const loaderHost = (): HTMLElement => {
        const cards = Array.from(document.querySelectorAll<HTMLElement>(".collection-section, .collection-form-section"));
        const visible = cards.find(c => c.offsetParent !== null);
        return visible || (document.querySelector(".collection-section") as HTMLElement | null) || document.body;
      };


    const getDigest = async (): Promise<string> => {
  return getAuthDigest(siteUrl);
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



      const generateCollectionTitle = (project: ProjectItem, collectionData: CollectionData): string => {
        const currentYear = new Date().getFullYear();
        let totalValue = 0;

        Object.keys(collectionData).forEach(key => {
          if (key.match(/^\d{4}-\d{2}$/)) {
            const monthData = collectionData[key] as CollectionMonthData;
            if (monthData) {
              totalValue += monthData.actual;
            }
          }
        });

        return `${project.Title} - ${currentYear} Collection - ${totalValue.toLocaleString('en-US')}`;
      };

      const getNumericValue = (input: HTMLInputElement): number => {
        const value = input.value.trim();
        if (value === '' || value === '0') return 0;
        return parseInt(value) || 0;
      };

      const setNumericValue = (input: HTMLInputElement, value: number): void => {
        if (value === 0) {
          input.value = '';
          input.className = input.className.replace('collection-input-with-value', 'collection-input-empty');
        } else {
          input.value = value.toString();
          input.className = input.className.replace('collection-input-empty', 'collection-input-with-value');
        }
      };



      // Utility function to allow only numbers 0-9 in text inputs
      const allowOnlyNumbers = (input: HTMLInputElement): void => {
        input.addEventListener('input', function() {
          let value = this.value.replace(/[^\d]/g, '');

          if (value.length > 1 && value.startsWith('0')) {
            value = value.replace(/^0+/, '');
            if (value === '') value = '0';
          }

          this.value = value;

          if (value === '' || value === '0') {
            this.className = this.className.replace('collection-input-with-value', 'collection-input-empty');
          } else {
            this.className = this.className.replace('collection-input-empty', 'collection-input-with-value');
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


      // Enhanced OBL validation with toast notifications
      const showOBLErrorToast = (message: string, isWarning: boolean = false): void => {
        const existingToast = document.querySelector('.collection-obl-error-toast');
        if (existingToast) existingToast.remove();

        const toastEl = document.createElement('div');
        toastEl.className = `collection-obl-error-toast ${isWarning ? 'collection-warning' : ''}`;
        toastEl.innerHTML = `
          ${message}
          <button class="collection-toast-close">&times;</button>
        `;

        document.body.appendChild(toastEl);

        setTimeout(() => {
          if (toastEl.parentNode) {
            toastEl.remove();
          }
        }, 5000);

        toastEl.querySelector('.collection-toast-close')?.addEventListener('click', () => {
          toastEl.remove();
        });
      };


      // Type guard to check if data is CollectionMonthData
      const isCollectionMonthData = (data: any): data is CollectionMonthData => {
        return data && typeof data.plan === 'number' && typeof data.actual === 'number';
      };


// 1. First, update the checkIfNeedsProgression function in collection:
const checkIfNeedsProgression = (collectionData: CollectionData): boolean => {
    const now = new Date();
    const currentMonthKey = formatMonthKey(now);

    const months = Object.keys(collectionData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();

    if (months.length === 0) return true;

    // Check if data has current month
    const hasCurrentMonth = months.includes(currentMonthKey);
    if (!hasCurrentMonth) {
        console.log(`🔄 Month progression needed: Missing current month ${currentMonthKey}`);
        return true;
    }

    // Get the last month in the data
    const lastDataMonth = months[months.length - 1];
    const idealMonths = generateIdealMonthStructure();
    const lastIdealMonth = idealMonths[idealMonths.length - 1];

    // Check monthly progression
    const monthlyProgressionNeeded = lastDataMonth < lastIdealMonth;

    // Check yearly progression
    const currentYear = now.getFullYear();
    let yearlyProgressionNeeded = false;

    if (collectionData.yearly) {
        const years = Object.keys(collectionData.yearly).map(y => parseInt(y)).sort();
        yearlyProgressionNeeded = years.length > 0 && years[0] < currentYear;
    } else {
        yearlyProgressionNeeded = true;
    }

    console.log(`🔍 Progression check: monthly=${monthlyProgressionNeeded}, yearly=${yearlyProgressionNeeded}`);
    return monthlyProgressionNeeded || yearlyProgressionNeeded;
};

 const propagateMonthValues = (collectionData: CollectionData): CollectionData => {
    const now = new Date();
    const currentMonthKey = formatMonthKey(now);

    const originalMonths = Object.keys(collectionData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
    const idealMonths = generateIdealMonthStructure();

    const hasCurrentMonth = originalMonths.includes(currentMonthKey);
    const lastDataMonth = originalMonths[originalMonths.length - 1];
    const lastIdealMonth = idealMonths[idealMonths.length - 1];

    if (hasCurrentMonth && lastDataMonth === lastIdealMonth) {
        return collectionData;
    }

    const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
    const detailedStartIndex = currentMonthIndex !== -1 ? currentMonthIndex : 36;

    const updatedData: CollectionData = {};

    // First pass: Create or copy all months
    idealMonths.forEach((monthKey, index) => {
        const existingData = collectionData[monthKey];
        const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
        const isLongTermMonth = index >= detailedStartIndex + 3;

        if (existingData && isCollectionMonthData(existingData)) {
              const copied: CollectionMonthData = JSON.parse(JSON.stringify(existingData));
              if (copied.monthActual === undefined) {
                  copied.monthActual = 0;
              }
              updatedData[monthKey] = copied;
          } else {
            let items = {};
            if (isDetailedMonth) {
                items = createEmptyCategories();
            } else if (isLongTermMonth) {
                items = createSimpleCategories();
            }
            updatedData[monthKey] = {
                plan: 0,
                actual: 0,
                monthActual: 0,
                forecast: 0,
                date: '',
                items: items
            };
        }
    });

    // Second pass: Carry values forward
    for (let i = 1; i < idealMonths.length; i++) {
        const prevMonthKey = idealMonths[i - 1];
        const currMonthKey = idealMonths[i];

        const prevMonthData = updatedData[prevMonthKey];
        const currMonthData = updatedData[currMonthKey];

        if (!prevMonthData || !isCollectionMonthData(prevMonthData) ||
            !currMonthData || !isCollectionMonthData(currMonthData)) {
            continue;
        }

        const isCurrMonthNew = !originalMonths.includes(currMonthKey);

        if (isCurrMonthNew) {
            const prevMonthValue = (prevMonthData.monthActual !== undefined && prevMonthData.monthActual > 0)
                ? prevMonthData.monthActual
                : prevMonthData.actual;

            currMonthData.plan = prevMonthValue;
            currMonthData.actual = 0;
            currMonthData.monthActual = 0;

            const monthIndex = idealMonths.indexOf(currMonthKey);
            const isDetailedMonth = monthIndex >= detailedStartIndex && monthIndex < detailedStartIndex + 3;

            if (isDetailedMonth) {
                if (!currMonthData.items) {
                    currMonthData.items = {};
                }
                if (!prevMonthData.items) {
                    prevMonthData.items = {};
                }

                const currItems = currMonthData.items;
                const prevItems = prevMonthData.items;

                Object.keys(currItems).forEach(categoryKey => {
                    const prevItem = prevItems[categoryKey];
                    const currItem = currItems[categoryKey];

                    if (prevItem && currItem) {
                        currItem.plan = prevItem.actual || 0;
                        currItem.actual = 0;
                        currItem.forecast = 0;
                        currItem.date = '';
                    }
                });
            }
        }
    }

    // Preserve yearly data
    if (collectionData.yearly) {
        updatedData.yearly = JSON.parse(JSON.stringify(collectionData.yearly));
    } else {
        const lastMonthKey = idealMonths[idealMonths.length - 1];
        updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
    }

    return updatedData;
};
     // SET FORM READONLY FUNCTION
      // ============================================

const setFormReadOnly = (isFrozen: boolean): void => {
  const editableSelectors = [
    '.collection-detailed-month-box .collection-category-current-input',
    '.collection-simple-month',
    '.collection-yearly-box',
    '.collection-category-current-input',
    '.collection-category-forecast-input',  
    '#remarksInternal',
    '#remarksSprintCell'
  ];

  const inputs = document.querySelectorAll<HTMLInputElement | HTMLTextAreaElement>(editableSelectors.join(','));

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

  // month-actual-input should ALWAYS be readonly
  document.querySelectorAll('.collection-month-actual-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.setAttribute('readonly', 'true');
    el.style.backgroundColor = '#e8f5e9';
    el.style.cursor = 'not-allowed';
    el.style.pointerEvents = 'none';
  });

  // forecast-total-input should ALWAYS be readonly
  document.querySelectorAll('.collection-forecast-total-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.setAttribute('readonly', 'true');
    el.style.backgroundColor = '#f0f8ff';
    el.style.cursor = 'not-allowed';
    el.style.pointerEvents = 'none';
  });
document.querySelectorAll('.collection-forecast-date-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.setAttribute('readonly', 'true');
    el.style.backgroundColor = '#f0f8ff';
    el.style.cursor = 'not-allowed';
    el.style.pointerEvents = 'none';
  });

   

  // main-forecast-input (Further) should be editable when frozen
  document.querySelectorAll('.collection-main-forecast-input').forEach((input) => {
    const el = input as HTMLInputElement; 
    el.disabled = false;
    el.style.backgroundColor = '#ffffff';
    el.style.color = '#000000';
    el.style.cursor = 'text';
    el.style.opacity = '1';
    el.style.pointerEvents = 'auto';    
    el.style.borderColor = '#28a745';
    el.style.borderWidth = '2px';
    el.style.borderStyle = 'solid';
  });

  // category-forecast-input (Further) should be editable when frozen
  document.querySelectorAll('.collection-category-forecast-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.removeAttribute('readonly');
    el.disabled = false;
    el.style.backgroundColor = '#ffffff';
    el.style.color = '#000000';
    el.style.cursor = 'text';
    el.style.opacity = '1';
    el.style.pointerEvents = 'auto';
    el.style.borderColor = '#28a745';
    el.style.borderWidth = '2px';
    el.style.borderStyle = 'solid';
  });

  const oblInput = document.querySelector('#oblLastMonth') as HTMLInputElement;
  if (oblInput) {
    if (isFrozen) {
      oblInput.setAttribute('readonly', 'true');
      oblInput.style.backgroundColor = '#f5f5f5';
      oblInput.style.cursor = 'not-allowed';
    } else {
      oblInput.setAttribute('readonly', 'true');
      oblInput.style.backgroundColor = '';
      oblInput.style.cursor = '';
    }
  }

  if (isFrozen) {
    // toast("info", '🔒 This project is frozen. Further and Date fields are editable.');
  } else {
    // toast("info", '🔓 Project unfrozen. You can now edit all fields.');
  }
};

// ═══════════════════════════════════════════════════════════
// FIX 1: SINGLE SOURCE OF TRUTH FOR THE EDIT FORM TITLE
// Always appends the project name, so the title can never be
// overwritten by a later timeout without the project name.
// ═══════════════════════════════════════════════════════════
// const getCurrentProjectName = (): string => {
//   const projectNameEl = document.querySelector("#projectNameDisplay") as HTMLElement | null;
//   const projectName = projectNameEl?.textContent?.trim() || '';
//   return (projectName && projectName !== '-') ? projectName : '';
// };

const updateEditFormHeader = (isFrozen: boolean, canEdit: boolean): void => {
  const formTitleEl = document.querySelector("#formTitle") as HTMLElement | null;
  if (!formTitleEl) return;

  // const projectName = getCurrentProjectName();
  // ✅ "Edit Collection — " prefix removed: the project name IS the title
  // (with the 🔒 / (Locked By Admin) / (Read Only) suffixes preserved for
  //  frozen or read-only rows, and a safe fallback when no project is set).
  const baseTitle = "Collection Plan 5QRF";

  if (isFrozen) {
    formTitleEl.textContent = `🔒 ${baseTitle} (Locked By Admin)`;
    formTitleEl.style.color = '#856404';
  } else if (!canEdit) {
    formTitleEl.textContent = `🔒 ${baseTitle} (Read Only)`;
    formTitleEl.style.color = '#856404';
  } else {
    formTitleEl.textContent = baseTitle;
    formTitleEl.style.color = '';
  }
};

      // ============================================
      // MONTH STRUCTURE FUNCTIONS
      // ============================================

      const createEmptyCategories = (): { [key: string]: CollectionCategory } => {
        return {
          "Normal": {
            name: 'Normal',
            plan: 0,
            actual: 0,
            forecast: 0,
            date: ''
          },
          "Normal Overdue": {
            name: 'Normal Overdue',
            plan: 0,
            actual: 0,
            forecast: 0,
            date: ''
          },
          "Retention": {
            name: 'Retention',
            plan: 0,
            actual: 0,
            forecast: 0,
            date: ''
          },
          "Retention Overdue": {
            name: 'Retention Overdue',
            plan: 0,
            actual: 0,
            forecast: 0,
            date: ''
          }
        };
      };

      const createSimpleCategories = (): { [key: string]: CollectionCategory } => {
        return {
          "Normal": {
            name: 'Normal',
            plan: 0,
            actual: 0,
            forecast: 0,
            date: ''
          },
          "Retention": {
            name: 'Retention',
            plan: 0,
            actual: 0,
            forecast: 0,
            date: ''
          }
        };
      };

      const generateIdealMonthStructure = (): string[] => {
        const months: string[] = [];
        const currentDate = new Date();

        for (let i = -3; i < 0; i++) {
          const historicalDate = new Date(currentDate);
          historicalDate.setMonth(historicalDate.getMonth() + i);
          months.push(formatMonthKey(historicalDate));
        }

        for (let i = -36; i < -3; i++) {
          const historicalDate = new Date(currentDate);
          historicalDate.setMonth(historicalDate.getMonth() + i);
          months.push(formatMonthKey(historicalDate));
        }

        for (let i = 0; i < 3; i++) {
          const detailedDate = new Date(currentDate);
          detailedDate.setMonth(detailedDate.getMonth() + i);
          months.push(formatMonthKey(detailedDate));
        }

        for (let i = 3; i < 19; i++) {
          const futureDate = new Date(currentDate);
          futureDate.setMonth(futureDate.getMonth() + i);
          months.push(formatMonthKey(futureDate));
        }

        return months.sort();
      };

      const createFreshYearlyStructure = (lastMonthKey?: string): YearlyCollectionData => {
        const yearlyData: YearlyCollectionData = {};

        let startYear: number;

        if (lastMonthKey) {
          startYear = parseInt(lastMonthKey.split('-')[0]);
        } else {
          startYear = new Date().getFullYear() + 1;
        }

        const endYear = startYear + 3;

        for (let year = startYear; year <= endYear; year++) {
          yearlyData[year.toString()] = {
            plan: 0,
            actual: 0,
            items: {
              "Normal": { name: 'Normal', plan: 0, actual: 0 },
              "Retention": { name: 'Retention', plan: 0, actual: 0 }
            }
          };
        }

        return yearlyData;
      };

const createFreshCollectionStructure = (): CollectionData => {
    const idealMonths = generateIdealMonthStructure();
    const freshData: CollectionData = {};

    const now = new Date();
    const currentMonthKey = formatMonthKey(now);

    const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
    const detailedStartIndex = currentMonthIndex !== -1 ? currentMonthIndex : 36;

    idealMonths.forEach((monthKey, index) => {
        const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;

        let items = {};
        if (isDetailedMonth) {
            items = createEmptyCategories();
        } else if (index >= detailedStartIndex + 3) {
            items = createSimpleCategories();
        }

        freshData[monthKey] = {
            plan: 0,
            actual: 0,
            monthActual: 0,
            forecast: 0,
            date: '',
            items: items
        };
    });

    const lastMonthKey = idealMonths[idealMonths.length - 1];
    freshData.yearly = createFreshYearlyStructure(lastMonthKey);

    return freshData;
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
        return generateIdealMonthStructure();
      };

      // ============================================
      // DATA LOADING FUNCTIONS
      // ============================================
const loadProjects = async (): Promise<void> => {
  try {
    // ✅ $select must include every field that ROLE_OWN_FIELDS in auth.ts
    // matches against, otherwise getAssignedProjects() returns [] for the
    // KRIYA_Engg_* roles.
    const res = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items` +
      `?$select=Id,Title,CustomerName,PC,ProjectReference,` +
      `GroupHead1/Title,GroupHead1Id,` +
      `ProjectManager/Title,ProjectManagerId,` +
      `EnggLead/Title,EnggLeadId,` +
      `EnggSubLead1/Title,EnggSubLead1Id,` +
      `EnggSubLead2/Title,EnggSubLead2Id,` +
      `EngineeringManagers/Title,EngineeringManagersId` +
      `&$expand=GroupHead1,ProjectManager,EnggLead,EnggSubLead1,EnggSubLead2,EngineeringManagers`,
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
      // ✅ keep engineering fields so getAssignedProjects() can match
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

    console.log('📋 All projects with role assignments:');
    projects.forEach(p => {
      const nameOf = (v: any) => typeof v === 'object' ? v?.Title : v;
      console.log(
        `  - ${p.Title}: ` +
        `PM=${nameOf(p.ProjectManager) || 'None'}, ` +
        `GH=${nameOf(p.GroupHead1) || 'None'}, ` +
        `EL=${nameOf(p.EnggLead) || 'None'}, ` +
        `ES1=${nameOf(p.EnggSubLead1) || 'None'}, ` +
        `ES2=${nameOf(p.EnggSubLead2) || 'None'}, ` +
        `EM=${nameOf(p.EngineeringManagers) || 'None'}`
      );
    });

    // 🔐 RBAC: single source of truth for project scope (auth.ts)
    //   - canReadAllProjects()  → Admin / SC Team / Controlling: ALL projects
    //   - everyone else         → getAssignedProjects() filters by the role's
    //                             own field(s): GH Heads → GroupHead1,
    //                             PM Managers → ProjectManager,
    //                             Engg Lead / Sub1 / Sub2 / Mgr → the
    //                             matching Engg* field.
    const assignedProjects = canReadAllProjects()
      ? projects
      : getAssignedProjects(projects);

    console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);
    console.log('📋 Assigned project titles:', assignedProjects.map(p => p.Title));

    // Populate project select dropdown (Add / Edit form)
    projectSelect.innerHTML = `<option value="">Select Project</option>` +
      assignedProjects.map(p => `<option value="${p.Id}">${p.Title}</option>`).join("");

    // ✅ Also populate the filter dropdown (was missing before)
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
  return typeof groupHead === "string" ? groupHead : "";
};

const getProjectManagerDisplayName = (manager: string | { Title: string; Email?: string } | undefined): string => {
  if (!manager) return "";
  if (typeof manager === "object" && manager !== null) {
    return manager.Title || "";
  }
  return typeof manager === "string" ? manager : "";
};

// ============================================
// UPDATED LOAD COLLECTIONS
// ============================================

/* ✅ Bar loader row (grey track + indigo bar, no text) - replaces the old
   spinner + "Loading collection data..." text row. */
const COLLECTION_LOADING_ROW_HTML = `
    <tr class="collection-table-loading-row">
      <td colspan="28">
        <div class="collection-table-loading-content" role="status" aria-live="polite">
          <div class="collection-page-loader">
            <div class="collection-page-loader-track"><div class="collection-page-loader-bar"></div></div>
          </div>
        </div>
      </td>
    </tr>
  `;

const showCollectionTableLoading = (): void => {
  const tableBody = document.querySelector(".collection-table-body") as HTMLElement | null;
  if (!tableBody) return;

  tableBody.setAttribute("aria-busy", "true");
  tableBody.innerHTML = COLLECTION_LOADING_ROW_HTML;
};

const showCollectionTableLoadError = (): void => {
  const tableBody = document.querySelector(".collection-table-body") as HTMLElement | null;
  if (!tableBody) return;

  tableBody.setAttribute("aria-busy", "false");
  tableBody.innerHTML = `
    <tr class="collection-table-loading-row collection-table-loading-error-row">
      <td colspan="28">
        <div class="collection-table-loading-content" role="alert">
          <span>Unable to load collection data. Please refresh the page and try again.</span>
        </div>
      </td>
    </tr>
  `;
};

const loadCollections = async (): Promise<void> => {
  // Show the table loading state immediately while SharePoint data is being fetched.
  // Keep existing rows visible during background refreshes after save/update.
  if (collections.length === 0) {
    showCollectionTableLoading();
  }

  try {
    const res = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?` +
      `$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,CollectionData,RevenueActual,Remarks,CreatedDate,UpdatedDate,NormalDue,NormalOverdue,RetentionDue,RetentionOverdue,GstTds,WHT,Freeze` +
      `&$expand=Project&$top=5000`,
      {
        headers: { Accept: "application/json;odata=verbose" }
      }
    );
   
    if (!res.ok) {
      const errorText = await res.text(); 
      console.error("❌ API Error:", errorText);
      throw new Error(`HTTP ${res.status}: ${errorText}`);
    }

    const data: SharePointResponse<CollectionItem> = await res.json();

    // 🔐 RBAC: role-based project scope (must match the dropdown scope above).
    //   - Admin / SC Team / Controlling (READ_ALL) → ALL projects
    //   - GH Heads / PM Managers / Engg roles     → OWN assigned projects only
    let allowedProjectTitles: Set<string> = new Set();

    if (canReadAllProjects()) {
      allowedProjectTitles = new Set(projects.map(p => p.Title));
      console.log(`👁️ READ_ALL role - showing ALL projects (${allowedProjectTitles.size})`);
    } else {
      const assignedProjects = getAssignedProjects(projects);
      allowedProjectTitles = new Set(assignedProjects.map(p => p.Title));
      console.log(`👁️ Scoped role - showing OWN projects only (${allowedProjectTitles.size})`);
    }

    collections = (data.d?.results || []).map(collection => {
      let revenueActualObj = {};
      try {
        if (typeof collection.RevenueActual === 'string') {
          revenueActualObj = JSON.parse(collection.RevenueActual) || {};
        } else if (typeof collection.RevenueActual === 'object') {
          revenueActualObj = collection.RevenueActual || {};
        }
      } catch (e) {
        revenueActualObj = {};
      }

      return {
        ...collection,
        CollectionData: typeof collection.CollectionData === 'string' ?
          collection.CollectionData :
          JSON.stringify(collection.CollectionData || {}),
        _revenueActual: revenueActualObj,
        ReceivableValue: Number(collection.OBL) || 0,
        GSTTDS: Number(collection.GSTTDS) || 0,
        NormalDue: Number(collection.NormalDue) || 0,
        NormalOverdue: Number(collection.NormalOverdue) || 0,
        RetentionDue: Number(collection.RetentionDue) || 0,
        RetentionOverdue: Number(collection.RetentionOverdue) || 0,
        WHT: Number(collection.WHT) || 0
      };
    });

    // FILTER: Only keep collections for allowed projects
    const beforeCount = collections.length;
    collections = collections.filter(collection => {
      const projectTitle = collection.Project?.Title;
      const isAllowed = projectTitle && allowedProjectTitles.has(projectTitle);
      if (!isAllowed && projectTitle) {
        console.log(`🔍 Filtering out: ${projectTitle} (not allowed for this user)`);
      }
      return isAllowed;
    });

    console.log(`✅ Loaded ${collections.length} collections (filtered from ${beforeCount})`);

  } catch (error) {
    console.error("Error loading collections:", error);
    showCollectionTableLoadError();
    toast("error", "Failed to load collections");
  }
};


// 🔐 FORCE ADMIN-ONLY for "Upload Actual"
// We override the standard canContribute() check and strictly enforce Admin status.
const isAdmin = hasFullControl(); 
const uploadActualBtns = document.querySelectorAll<HTMLElement>('.collection-import-actual-btn');

uploadActualBtns.forEach(btn => {
  if (isAdmin) {
    btn.style.display = '';        // Show to Admin
    (btn as HTMLButtonElement).disabled = false;
  } else {
    btn.style.display = 'none';    // Hide from everyone else
    (btn as HTMLButtonElement).disabled = true;
  }
});

// ============================================
// AUTO-SYNC MISSING PROJECTS → COLLECTION LIST
// "Collection List" is the internal list name. Every project that exists in
// Project List but has NO row in Collection List yet is added automatically,
// with an empty CollectionData skeleton generated by
// createFreshCollectionStructure():
//   36 historical months + current & next 2 (detailed) + next 16 months
//   (long-term) + 4-year yearly block.
// 🔐 RBAC: only roles with create rights (canCreateRecord → Admin, GH Heads,
// PM Managers) may run it, and it only ever adds projects inside the role's
// own visible scope (READ_ALL → all projects; others → assigned only).
// Idempotent: rows that already exist are never duplicated or overwritten.
// ============================================
const syncMissingProjectsToCollectionList = async (): Promise<void> => {
  const syncStatus = document.querySelector('#collectionSyncStatus') as HTMLElement | null;
  const syncBtn = document.querySelector('#collectionSyncProjectsBtn') as HTMLButtonElement | null;

  const setStatus = (text: string): void => {
    console.log(text);
    if (syncStatus) syncStatus.textContent = text;
  };

  if (!canCreateRecord()) {
    setStatus('👁️ No create rights - sync skipped (Admin/GH Head/PM Manager only)');
    return;
  }

  if (!projects || projects.length === 0) {
    setStatus('⚠️ Project List is empty - nothing to sync');
    return;
  } 

  if (syncBtn) syncBtn.disabled = true;
  try {
    const visibleProjects = canReadAllProjects()
      ? projects
      : getAssignedProjects(projects);

    // Safety re-check against SharePoint itself: the in-memory list may be
    // empty/stale, so never create rows for projects that already exist.        
    let serverRows: any[] = [];
    try {
      const chkRes = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,ProjectId,Title&$top=5000`,
        { headers: { Accept: "application/json;odata=verbose" } }
      );
      if (chkRes.ok) serverRows = (await chkRes.json())?.d?.results || [];
    } catch (chkErr) {
      console.error("⚠️ Sync safety-check fetch failed:", chkErr);
    }

    // If the list really has rows on SharePoint but the page loaded 0, the
    // main load is broken - creating rows now would duplicate everything.
    if (serverRows.length > 0 && collections.length === 0) {
      setStatus(`⚠️ Collection List already has ${serverRows.length} row(s) on SharePoint but the page loaded 0 - sync stopped to avoid duplicates. Check the main collection load.`);
      toast("error", `Sync stopped: list has ${serverRows.length} rows but page loaded 0`);
      return;
    }

    const existsOnServer = (p: ProjectItem): boolean =>
      serverRows.some(c => Number(c.ProjectId) === p.Id);

    const existsInMemory = (p: ProjectItem): boolean =>
      collections.some(c =>
        c.Project?.Id === p.Id || (c.Project && c.Project.Title === p.Title)
      );

    const missingProjects = visibleProjects.filter(p => !existsOnServer(p) && !existsInMemory(p));

    if (missingProjects.length === 0) {
      // setStatus(`✅ In sync: all ${visibleProjects.length} visible project(s) have a Collection List row`);
      return;
    }

    // setStatus(`🆕 ${missingProjects.length} missing: ${missingProjects.map(p => p.Title).join(', ')}`);
    // toast("info", `Checking ${visibleProjects.length} projects... adding ${missingProjects.length} missing`);

    const digest = await getDigest();
    const entityType = await getListItemType();
    let added = 0;
    let failed = 0;
    const errors: string[] = [];

    for (const project of missingProjects) {
      let ok = false;
      // Try twice before giving up on a project
      // Up to 3 payload variants:
      //   1 = full create - ProjectReference lookup id as STRING + OBL as STRING
      //   2 = without ProjectReference (column may be read-only/misconfigured)
      //   3 = without OBL too (column may be read-only/absent)
      for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
        try {
          const freshCollectionData = createFreshCollectionStructure();

          // odata=verbose values must be STRINGs (Edm.String); lookup ids and
          // numeric fields sent as JSON numbers get a 400 "Cannot convert a
          // primitive value to the expected type 'Edm.String'".
          const payload: any = {
            __metadata: { type: entityType },
            Title: generateCollectionTitle(project, freshCollectionData),
            ProjectId: project.Id.toString(),
            ProfitCentre: project.PC || "",
            CollectionData: JSON.stringify(freshCollectionData),
            Remarks: "",
            OBL: project.PC ? "0" : "0",
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
          };

          // Store Project Reference when the column exists: a LOOKUP column
          // receives the target item ID as a STRING (ProjectReferenceId) - the
          // same rule used by the Revenue sync so both behave identically.
          const refField = await collectionListProjectRefField();
          if (refField.exists && attempt === 1) {
            if (refField.type === "Lookup") {
              payload.ProjectReferenceId = String(project.Id);
            } else {
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
            // setStatus(`✅ Added: ${project.Title}${project.ProjectReference ? ' (' + project.ProjectReference + ')' : ''}`);
          } else {
            const bodyText = await res.text();
            let message = `HTTP ${res.status}`;
            try {
              const errJson = JSON.parse(bodyText);
              message = errJson?.error?.message?.value || message;
            } catch (e) { /* keep http message */ }
            console.error(`❌ Attempt ${attempt} failed for ${project.Title}: ${bodyText}`);
            if (attempt === 2) {
              failed++;
              errors.push(`${project.Title} -> ${message}`);
            }
          }
        } catch (projectError) {
          console.error(`❌ Attempt ${attempt} error for ${project.Title}:`, projectError);
          if (attempt === 2) {
            failed++;
            errors.push(`${project.Title} -> ${projectError instanceof Error ? projectError.message : String(projectError)}`);
          }
        }
        if (!ok) await new Promise(resolve => setTimeout(resolve, 300));
      }
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    if (added > 0) {
      await loadCollections();
      console.log(`➕ Auto-sync added ${added} project(s)` + (failed > 0 ? ` (${failed} failed)` : ""));
    } else if (failed > 0) {
      console.error(`❌ Auto-sync failed for ${failed} project(s): ${errors.join(' | ')}`);
    }
  } catch (error) {
    console.error("❌ Auto-sync failed:", error);
  } finally {
    if (syncBtn) syncBtn.disabled = false;
  }
};

// Cached check: does Collection List physically have a 'ProjectReference' column,
// and is it a plain Text field or a Lookup to Project List?
let collectionProjectRefField: { exists: boolean; type: string } | null = null;
const collectionListProjectRefField = async (): Promise<{ exists: boolean; type: string }> => {
  if (collectionProjectRefField) return collectionProjectRefField;    
  try {
    const res = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq 'ProjectReference'&$select=InternalName,TypeAsString,ReadOnlyField`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    const data = await res.json();
    const results = data?.d?.results || [];
    if (results.length > 0) {  
      collectionProjectRefField = {
        exists: true,
        type: String(results[0].TypeAsString || "Text")
      };
      console.log(`📋 Collection List ProjectReference column type = ${collectionProjectRefField.type}`);
    } else {
      collectionProjectRefField = { exists: false, type: "" };
      console.log("📋 Collection List has no ProjectReference column");
    }
  } catch (error) {
    console.error("⚠️ Could not check ProjectReference column:", error);
    collectionProjectRefField = { exists: false, type: "" };
  }
  return collectionProjectRefField;
};

const getListItemType = async (): Promise<string> => {
  return getListEntityType(siteUrl, listName);
};

             // ========== TABLE RENDERER WITH VIRTUAL SCROLLING ==========
             const attachEditTriggerListeners = (): void => {
  document.querySelectorAll('.collection-edit-trigger').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const id = parseInt((e.currentTarget as HTMLElement).dataset.id!);
      const collection = collections.find(f => f.Id === id);
      editCollection(id, collection);
    });
  });
};
const attachFreezeEventListeners = (): void => {
  console.log("🔧 Attaching freeze event listeners...");

  const checkboxes = document.querySelectorAll('.collection-freeze-checkbox');
  const canUserFreeze = canFreeze();

  checkboxes.forEach(checkbox => {
    const cb = checkbox as HTMLInputElement;

    // Remove existing listeners by cloning
    const newCb = cb.cloneNode(true) as HTMLInputElement;
    cb.parentNode?.replaceChild(newCb, cb);

    newCb.style.pointerEvents = canUserFreeze ? 'auto' : 'none';
    newCb.style.cursor = canUserFreeze ? 'pointer' : 'not-allowed';
    newCb.style.opacity = canUserFreeze ? '1' : '0.6';
    newCb.style.display = 'inline-block';
    newCb.style.width = '18px';
    newCb.style.height = '18px';
    newCb.style.accentColor = '#0078d4';
    newCb.style.verticalAlign = 'middle';
    newCb.style.position = 'relative';
    newCb.style.zIndex = '100';
    newCb.style.margin = '0 auto';

    const isFrozen = newCb.dataset.freeze === 'true';
    newCb.checked = isFrozen;
    newCb.disabled = !canUserFreeze;

    if (!newCb.disabled) {
      // Use a single event handler
      newCb.onchange = async function(e) {
        e.preventDefault();
        e.stopPropagation();

        const target = e.target as HTMLInputElement;
        const id = parseInt(target.dataset.id!);
        const isChecked = target.checked;

        console.log(`🔄 Freeze toggle: ID=${id}, checked=${isChecked}`);

        const collection = collections.find(f => f.Id === id);
        if (!collection) {
          toast("error", "Collection not found");
          target.checked = !isChecked;
          return;
        }

        const freezeLoading = createLoadingOverlay("Updating freeze status...");
        loaderHost().appendChild(freezeLoading);
        try {
          const digest = await getDigest();
          const entityType = await getListItemType();

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
            collection.Freeze = isChecked;

            // Update filtered data
            const filteredCollection = filteredData.find(f => f.Id === id);
            if (filteredCollection) {
              filteredCollection.Freeze = isChecked;
            }

            if (editingId === id) {
              applyFreezeToAllFields(isChecked);
            }

            toast("success", `Project ${isChecked ? 'frozen' : 'unfrozen'}`);

            // Refresh the table
            await loadCollections();
            renderCurrentPage();

          } else {
            toast("error", "Failed to update freeze status");
            target.checked = !isChecked;
          }
        } catch (error) {
          console.error("Error updating freeze status:", error);
          toast("error", "Failed to update freeze status");
          target.checked = !isChecked;
        } finally {
          removeLoadingOverlay();
        }
      };
    }
  });
};

/**
 * Apply freeze status to ALL form fields - comprehensive version
 */
const applyFreezeToAllFields = (isFrozen: boolean): void => {
  console.log(`🔒 ${isFrozen ? 'FREEZING' : 'UNFREEZING'} all form fields`);

  // 1. OBL Input
  const oblInput = document.querySelector('#oblLastMonth') as HTMLInputElement;
  if (oblInput) {
    if (isFrozen) {
      oblInput.setAttribute('readonly', 'true');
      oblInput.style.backgroundColor = '#f5f5f5';
      oblInput.style.cursor = 'not-allowed';
      oblInput.style.opacity = '0.7';
    } else {
      oblInput.setAttribute('readonly', 'true');
      oblInput.style.backgroundColor = '';
      oblInput.style.cursor = '';
      oblInput.style.opacity = '1';
    }
  }

  // 2. OBL Breakdown fields (readonly always, but style changes)
  const breakdownFields = ['oblNormalDue', 'oblNormalOverdue', 'oblRetentionDue',
                           'oblRetentionOverdue', 'oblGstTds', 'oblWht'];
  breakdownFields.forEach(id => {
    const el = document.getElementById(id) as HTMLInputElement;
    if (el) {
      if (isFrozen) {
        el.style.backgroundColor = '#f5f5f5';
        el.style.opacity = '0.7';
      } else {
        el.style.backgroundColor = '';
        el.style.opacity = '1';
      }
    }
  });

  // 3. ALL input fields in the form (comprehensive)
  const allInputs = document.querySelectorAll('input, textarea');
  allInputs.forEach(input => {
    const el = input as HTMLInputElement | HTMLTextAreaElement;

    // Skip checkbox inputs
    if (el.type === 'checkbox') return;

    // Skip the freeze checkbox itself
    if (el.classList.contains('collection-freeze-checkbox')) return;

    // IMPORTANT: ALWAYS skip historical months - they should never be editable
    if (el.closest('.collection-historical-month')) {
      // Ensure historical inputs are always readonly
      el.setAttribute('readonly', 'true');
      el.style.backgroundColor = '#e9ecef';
      el.style.cursor = 'not-allowed';
      el.style.pointerEvents = 'none';
      return;
    }

    // Apply freeze
    if (isFrozen) {
      el.setAttribute('readonly', 'true');
      el.style.backgroundColor = '#f5f5f5';
      el.style.color = '#6c757d';
      el.style.cursor = 'not-allowed';
      el.style.opacity = '0.7';
      el.style.pointerEvents = 'none';
    } else {
      el.removeAttribute('readonly');
      el.style.backgroundColor = '';
      el.style.color = '';
      el.style.cursor = '';
      el.style.opacity = '1';
      el.style.pointerEvents = '';

      // Restore specific styles for editable fields
      if (el.classList.contains('collection-category-current-input') && !el.closest('.collection-historical-month')) {
        el.style.backgroundColor = '#ffffff';
        el.style.cursor = 'text';
        el.style.pointerEvents = 'auto';
      }

      // Restore main current inputs - they should be readonly but styled properly
      if (el.classList.contains('collection-main-current-input') && !el.closest('.collection-historical-month')) {
        el.setAttribute('readonly', 'true');
        el.style.backgroundColor = '#e9ecef';
        el.style.cursor = 'not-allowed';
        el.style.pointerEvents = 'none';
      }

      // Restore main plan inputs - they should be readonly
      if (el.classList.contains('collection-main-plan-input') || el.classList.contains('collection-category-plan-input')) {
        el.setAttribute('readonly', 'true');
        el.style.backgroundColor = '#e9ecef';
        el.style.cursor = 'not-allowed';
        el.style.pointerEvents = 'none';
      }

      // Restore yearly current inputs - they should be readonly
      if (el.classList.contains('collection-yearly-current-input') || el.classList.contains('collection-yearly-plan-input')) {
        el.setAttribute('readonly', 'true');
        el.style.backgroundColor = '#e9ecef';
        el.style.cursor = 'not-allowed';
        el.style.pointerEvents = 'none';
      }
    }
  });

  // 4. Special handling for category-forecast-input
  if (isFrozen) {
    // When frozen, make forecast inputs readonly
    document.querySelectorAll('.collection-category-forecast-input').forEach((input) => {
      const el = input as HTMLInputElement;
      el.setAttribute('readonly', 'true');
      el.style.backgroundColor = '#f5f5f5';
      el.style.cursor = 'not-allowed';
      el.style.pointerEvents = 'none';
      el.style.opacity = '0.7';
    });
  } else {
    // When unfrozen, make forecast inputs editable
    document.querySelectorAll('.collection-category-forecast-input').forEach((input) => {
      const el = input as HTMLInputElement;
      el.removeAttribute('readonly');
      el.style.backgroundColor = '#ffffff';
      el.style.cursor = 'text';
      el.style.pointerEvents = 'auto';
      el.style.opacity = '1';
    });
  }

  // 5. Remarks fields
  const remarksInternal = document.querySelector('#remarksInternal') as HTMLTextAreaElement;
  const remarksSprintCell = document.querySelector('#remarksSprintCell') as HTMLTextAreaElement;

  [remarksInternal, remarksSprintCell].forEach(el => {
    if (el) {
      if (isFrozen) {
        el.setAttribute('readonly', 'true');
        el.style.backgroundColor = '#f5f5f5';
        el.style.cursor = 'not-allowed';
        el.style.opacity = '0.7';
      } else {
        el.removeAttribute('readonly');
        el.style.backgroundColor = '';
        el.style.cursor = '';
        el.style.opacity = '1';
      }
    }
  });

  // 6. Date inputs - ALWAYS editable (even when frozen)
document.querySelectorAll('.collection-forecast-date-input, .collection-category-date-input').forEach((input) => {
  const el = input as HTMLInputElement;
  el.removeAttribute('readonly');
  el.disabled = false;
  el.style.backgroundColor = '#ffffff';
  el.style.cursor = 'pointer';
  el.style.pointerEvents = 'auto';
  el.style.opacity = '1';
});

  // 7. Update validation message
  const validationMessage = document.querySelector('.collection-validation-message') as HTMLElement;
  if (validationMessage) {
    if (isFrozen) {
      validationMessage.textContent = '🔒 Project is frozen - all fields are read-only';
      validationMessage.className = 'collection-validation-message info';
    } else {
      validationMessage.textContent = '✅ All values are within Receivables limit';
      validationMessage.className = 'collection-validation-message success';
      // Re-validate
      validateOBLLimit();
    }
  }

  // 8. Update the form title with freeze indicator
  const formTitle = document.querySelector("#formTitle") as HTMLElement;
  if (formTitle) {
    if (isFrozen) {
      if (!formTitle.textContent?.includes('🔒')) {
        formTitle.textContent = `🔒 ${formTitle.textContent?.replace('🔒 ', '') || 'Edit Collection'}`;
      }
    } else {
      formTitle.textContent = formTitle.textContent?.replace('🔒 ', '') || 'Edit Collection';
    } 
  }

  // toast("info", isFrozen ? '🔒 Project frozen. All fields are read-only.' : '🔓 Project unfrozen. You can now edit fields.');
};


      const getSafeString = (value: any, defaultValue: string = "-"): string => {
        if (value === null || value === undefined) return defaultValue;
        if (typeof value === 'string') return value.trim() || defaultValue;
        if (typeof value === 'object' && 'Title' in value) return value.Title?.trim() || defaultValue;
        return String(value) || defaultValue;
      };

      // ========== EXPORT FUNCTIONS ==========

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

  const currentExportData = filteredData.length > 0 ? filteredData : collections;

  if (currentExportData.length === 0) {
    toast("error", "No data available to export");
    return;
  }

  try {
    // Get ALL months from the ideal structure (54 months total)
    const allIdealMonths = generateIdealMonthStructure();

    const currentMonthKey = formatMonthKey(new Date());

    // Detailed months (3 months: current + next 2) - shown in table as detailed view
    const detailedMonths = allIdealMonths.slice(36, 39);

    // Long-term months (15 months) - shown in table as simple view
    const longTermMonths = allIdealMonths.slice(39, 54);

    // Get the last month from long-term to determine yearly start
    const lastMonthKey = longTermMonths[longTermMonths.length - 1];
    const startYear = parseInt(lastMonthKey.split('-')[0]);

    // Build headers - with correct order: PR → Customer → Project → GH → PM
    const baseHeaders = [
      "Freeze",
      "Project Reference",
      "Customer",
      "Project Name",
      "Group Head",
      "Project Manager",
      "Receivables(TDS)"
    ];

    // 1. Detailed Months (3 months: current + next 2)
    const detailedHeaders: string[] = [];
    detailedMonths.forEach((monthKey) => {
      const monthDate = new Date(monthKey + "-01");
      const monthLabel = formatMonthLabel(monthDate);
      const isCurrentMonth = monthKey === currentMonthKey;

      // Previous columns
      detailedHeaders.push(`${monthLabel} - Prev Normal`);
      detailedHeaders.push(`${monthLabel} - Prev Retention`);
      detailedHeaders.push(`${monthLabel} - Prev Total`);

      // Current columns
      detailedHeaders.push(`${monthLabel} - Curr Normal`);
      detailedHeaders.push(`${monthLabel} - Curr Retention`);
      detailedHeaders.push(`${monthLabel} - Curr Total`);

      // Current month only: Actual, Further, Forecast
      if (isCurrentMonth) {
        detailedHeaders.push(`${monthLabel} - Actual`);
        detailedHeaders.push(`${monthLabel} - Further`);
        detailedHeaders.push(`${monthLabel} - Forecast`);
      }
    });

    // 2. Long-term Months (15 months) - Previous and Current only
    const longTermHeaders: string[] = [];
    longTermMonths.forEach((monthKey) => {
      const monthDate = new Date(monthKey + "-01");
      const monthLabel = formatMonthLabel(monthDate);
      // Long-term months show Previous and Current (Normal and Retention only)
      longTermHeaders.push(`${monthLabel} - Prev Normal`);
      longTermHeaders.push(`${monthLabel} - Prev Retention`);
      longTermHeaders.push(`${monthLabel} - Curr Normal`);
      longTermHeaders.push(`${monthLabel} - Curr Retention`);
    });

    // 3. Yearly Data (4 years)
    const yearlyHeaders: string[] = [];
    for (let i = 0; i < 3; i++) {
      const year = (startYear + i).toString();
      yearlyHeaders.push(`${year} - Prev`);
      yearlyHeaders.push(`${year} - Curr`);
      yearlyHeaders.push(`${year} - Prev Normal`);
      yearlyHeaders.push(`${year} - Prev Retention`);
      yearlyHeaders.push(`${year} - Curr Normal`);
      yearlyHeaders.push(`${year} - Curr Retention`);
    }

    // Combine all headers (NO historical headers)
    const finalHeaders = [
      ...baseHeaders,
      ...detailedHeaders,
      ...longTermHeaders,
      ...yearlyHeaders
    ];

    // Prepare data for Excel
    const excelData = currentExportData.map(collection => {
      const projectTitle = collection.Project?.Title || "-";
      const proj = projects.find(p => p.Title === projectTitle);

      // Parse data from CollectionData
      const collectionData: CollectionData = safeJsonParse(collection.CollectionData, {});

      // Get OBL breakdown data
      const oblBreakdown = getOBLBreakdownForExport(collection);

      // Calculate total receivables
      const safeOblTotal = Number(oblBreakdown.normalDue || 0) +
                          Number(oblBreakdown.normalOverdue || 0) +
                          Number(oblBreakdown.retentionDue || 0) +
                          Number(oblBreakdown.retentionOverdue || 0) +
                          Number(oblBreakdown.gstTds || 0) +
                          Number(oblBreakdown.wht || 0);

      const isFrozen = collection.Freeze || false;

      const groupHeadName = getGroupHeadDisplayName(proj?.GroupHead1);
      const managerName = getProjectManagerDisplayName(proj?.ProjectManager);

      const rowData: any = {
        "Freeze": isFrozen ? "🔒" : "",
        "Project Reference": proj?.ProjectReference || "-",
        "Customer": proj?.CustomerName || "-",
        "Project Name": projectTitle,
        "Group Head": groupHeadName,
        "Project Manager": managerName,
        "Receivables(TDS)": safeOblTotal
      };

      // Helper function to add detailed month data
      const addDetailedMonthData = (monthKey: string) => {
        const monthData = collectionData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, forecast: 0, items: {} };
        const items = monthData.items || {};
        const monthLabel = formatMonthLabel(new Date(monthKey + "-01"));
        const isCurrentMonth = monthKey === currentMonthKey;

        const prevNormal = (items.Normal?.plan || 0) + (items['Normal Overdue']?.plan || 0);
        const prevRetention = (items.Retention?.plan || 0) + (items['Retention Overdue']?.plan || 0);
        const prevTotal = prevNormal + prevRetention;

        const currNormal = (items.Normal?.actual || 0) + (items['Normal Overdue']?.actual || 0);
        const currRetention = (items.Retention?.actual || 0) + (items['Retention Overdue']?.actual || 0);
        const currTotal = currNormal + currRetention;

        rowData[`${monthLabel} - Prev Normal`] = prevNormal;
        rowData[`${monthLabel} - Prev Retention`] = prevRetention;
        rowData[`${monthLabel} - Prev Total`] = prevTotal;
        rowData[`${monthLabel} - Curr Normal`] = currNormal;
        rowData[`${monthLabel} - Curr Retention`] = currRetention;
        rowData[`${monthLabel} - Curr Total`] = currTotal;

        if (isCurrentMonth) {
          // Get actual revenue value (RevenueActual)
          let actualValue = 0;
          try {
            if (collection._revenueActual && typeof collection._revenueActual === 'object') {
              actualValue = Number(collection._revenueActual[monthKey]) || 0;
            } else if (typeof collection.RevenueActual === 'string') {
              const parsed = JSON.parse(collection.RevenueActual);
              if (typeof parsed === 'object' && parsed !== null) {
                actualValue = Number(parsed[monthKey]) || 0;
              }
            } else if (typeof collection.RevenueActual === 'object') {
              actualValue = Number((collection.RevenueActual as any)[monthKey]) || 0;
            }
          } catch (e) {
            actualValue = 0;
          }

          const furtherValue = monthData.forecast || 0;
          // Forecast = Actual + Further
          const forecastValue = actualValue + furtherValue;

          rowData[`${monthLabel} - Actual`] = actualValue;
          rowData[`${monthLabel} - Further`] = furtherValue;
          rowData[`${monthLabel} - Forecast`] = forecastValue;
        }
      };

      // Helper function to add long-term month data
      const addLongTermMonthData = (monthKey: string) => {
        const monthData = collectionData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, items: {} };
        const items = monthData.items || {};
        const monthLabel = formatMonthLabel(new Date(monthKey + "-01"));

        const prevNormal = (items.Normal?.plan || 0) + (items['Normal Overdue']?.plan || 0);
        const prevRetention = (items.Retention?.plan || 0) + (items['Retention Overdue']?.plan || 0);

        const currNormal = (items.Normal?.actual || 0) + (items['Normal Overdue']?.actual || 0);
        const currRetention = (items.Retention?.actual || 0) + (items['Retention Overdue']?.actual || 0);

        rowData[`${monthLabel} - Prev Normal`] = prevNormal;
        rowData[`${monthLabel} - Prev Retention`] = prevRetention;
        rowData[`${monthLabel} - Curr Normal`] = currNormal;
        rowData[`${monthLabel} - Curr Retention`] = currRetention;
      };

      // Helper function to add yearly data
      const addYearlyData = (year: string) => {
        const yearData = collectionData.yearly?.[year] || { plan: 0, actual: 0, items: {} };
        const items = yearData.items || {};

        rowData[`${year} - Prev`] = yearData.plan || 0;
        rowData[`${year} - Curr`] = yearData.actual || 0;
        rowData[`${year} - Prev Normal`] = items.Normal?.plan || 0;
        rowData[`${year} - Prev Retention`] = items.Retention?.plan || 0;
        rowData[`${year} - Curr Normal`] = items.Normal?.actual || 0;
        rowData[`${year} - Curr Retention`] = items.Retention?.actual || 0;
      };

      // 1. Add Detailed Months (3 months: current + next 2)
      detailedMonths.forEach((monthKey) => {
        addDetailedMonthData(monthKey);
      });

      // 2. Add Long-term Months (15 months)
      longTermMonths.forEach((monthKey) => {
        addLongTermMonthData(monthKey);
      });

      // 3. Add Yearly Data (4 years)
      for (let i = 0; i < 3; i++) {
        const year = (startYear + i).toString();
        addYearlyData(year);
      }

      return rowData;
    });

    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData, { header: finalHeaders });

    // Apply styles
    const headerStyle = {
      fill: { fgColor: { rgb: "FF0078D4" } },
      font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
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

    // Apply styles to all cells
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

    // Style header row
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;

      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s = headerStyle;
    }

    // Style data rows
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;

        const isNumberColumn = col >= 6;
        const styleToApply = isNumberColumn ? numberStyle : dataStyle;

        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        ws[cellAddress].s = styleToApply;
      }
    }

    // Set column widths
    const colWidths = finalHeaders.map((header: string) => {
      if (header === 'Freeze') return { wch: 8 };
      if (header.includes('Project Name')) return { wch: 30 };
      if (header.includes('Project Reference')) return { wch: 18 };
      if (header.includes('Project Manager') || header.includes('Group Head')) return { wch: 18 };
      if (header.includes('Customer')) return { wch: 20 };
      if (header.includes('Receivables')) return { wch: 18 };
      if (header.includes('Prev') || header.includes('Curr') ||
          header.includes('Actual') || header.includes('Further') || header.includes('Forecast')) {
        return { wch: 14 };
      }
      return { wch: 12 };
    });
    ws['!cols'] = colWidths;

    // Freeze header row
    ws['!freeze'] = { x: 0, y: 1 };

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, "Collection");

    // Generate filename and export
    const monthName = getMonthName(selectedMonth);
    const fileName = `Collection_${monthName}_${selectedYear}_Current_to_Future.xlsx`;

    XLSX.writeFile(wb, fileName);

    toast("success", `Exported ${currentExportData.length} records from current month to future for ${monthName} ${selectedYear}`);

  } catch (error) {
    console.error("❌ Error during styled Excel export:", error);
    toast("error", "Failed to export data to Excel");
  }
};
 
// ========== EXPORT HISTORICAL DATA ONLY - Simplified ==========
const exportHistoricalToExcel = (): void => {
  // 🔐 RBAC: exporting data requires at least READ access
  if (!canRead()) {
    toast("error", "❌ You don't have permission to export data");
    return;
  }

  // ✅ Same validation as the Summary export — a year and month must be
  // selected in the Filters popup before exporting.
  const selectedYear = filterYear.value;
  const selectedMonth = filterMonth.value;

  if (!selectedYear || !selectedMonth) {
    toast("error", "Please select year and month to export");
    return;
  }

  const currentExportData = filteredData.length > 0 ? filteredData : collections;

  if (currentExportData.length === 0) {
    toast("error", "No data available to export");
    return;
  }

  try {
    // Get ALL months from the ideal structure (54 months total)
    const allIdealMonths = generateIdealMonthStructure();

    // Historical months (36 months) - ONLY these
    const historicalMonths = allIdealMonths.slice(0, 36);

    // Build headers with correct order: PR → Customer → Project → GH → PM
    const baseHeaders = [
      "Freeze",
      "Project Reference",
      "Customer",
      "Project Name",
      "Group Head",
      "Project Manager",
      "Receivables (incl TDS)"
    ];

    // Historical Months (36 months) - ONLY Actual, Current Normal, Current Retention
    const historicalHeaders: string[] = [];
    historicalMonths.forEach((monthKey) => {
      const monthDate = new Date(monthKey + "-01");
      const monthLabel = formatMonthLabel(monthDate);
      historicalHeaders.push(`${monthLabel} - Current Normal`);
      historicalHeaders.push(`${monthLabel} - Current Retention`);
      historicalHeaders.push(`${monthLabel} - Actual`);
    });

    // Combine all headers
    const finalHeaders = [
      ...baseHeaders,
      ...historicalHeaders
    ];

    // Prepare data for Excel
    const excelData = currentExportData.map(collection => {
      const projectTitle = collection.Project?.Title || "-";
      const proj = projects.find(p => p.Title === projectTitle);

      const collectionData: CollectionData = safeJsonParse(collection.CollectionData, {});

      const oblBreakdown = getOBLBreakdownForExport(collection);

      const safeOblTotal = Number(oblBreakdown.normalDue || 0) +
                          Number(oblBreakdown.normalOverdue || 0) +
                          Number(oblBreakdown.retentionDue || 0) +
                          Number(oblBreakdown.retentionOverdue || 0) +
                          Number(oblBreakdown.gstTds || 0) +
                          Number(oblBreakdown.wht || 0);

      const isFrozen = collection.Freeze || false;

      // Parse RevenueActual data for ALL months
      let revenueActualData: { [key: string]: number } = {};
      try {
        if (collection._revenueActual && typeof collection._revenueActual === 'object') {
          revenueActualData = collection._revenueActual;
        } else if (typeof collection.RevenueActual === 'string') {
          const parsed = JSON.parse(collection.RevenueActual);
          if (typeof parsed === 'object' && parsed !== null) {
            revenueActualData = parsed;
          }
        } else if (typeof collection.RevenueActual === 'object') {
          revenueActualData = collection.RevenueActual || {};
        }
      } catch (e) {
        revenueActualData = {};
      }

      const groupHeadName = getGroupHeadDisplayName(proj?.GroupHead1);
      const managerName = getProjectManagerDisplayName(proj?.ProjectManager);

      const rowData: any = {
        "Freeze": isFrozen ? "🔒" : "",
        "Project Reference": proj?.ProjectReference || "-",
        "Customer": proj?.CustomerName || "-",
        "Project Name": projectTitle,
        "Group Head": groupHeadName,
        "Project Manager": managerName,
        "Receivables (incl TDS)": safeOblTotal
      };

      // Add historical month data - ONLY Current Normal, Current Retention, Actual
      historicalMonths.forEach((monthKey) => {
        const monthData = collectionData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, items: {} };
        const items = monthData.items || {};
        const monthLabel = formatMonthLabel(new Date(monthKey + "-01"));

        // Current Normal = Normal actual + Normal Overdue actual
        const currentNormal = (items.Normal?.actual || 0) + (items['Normal Overdue']?.actual || 0);

        // Current Retention = Retention actual + Retention Overdue actual
        const currentRetention = (items.Retention?.actual || 0) + (items['Retention Overdue']?.actual || 0);

        // Actual from RevenueActual
        const actualValue = Number(revenueActualData[monthKey]) || 0;

        rowData[`${monthLabel} - Current Normal`] = currentNormal;
        rowData[`${monthLabel} - Current Retention`] = currentRetention;
        rowData[`${monthLabel} - Actual`] = actualValue;
      });

      return rowData;
    });

    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData, { header: finalHeaders });

    // Apply styles
    const headerStyle = {
      fill: { fgColor: { rgb: "FF0078D4" } },
      font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
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

    // Apply styles to all cells
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

    // Style header row
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;

      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s = headerStyle;
    }

    // Style data rows
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;

        const isNumberColumn = col >= 6;
        const styleToApply = isNumberColumn ? numberStyle : dataStyle;

        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        ws[cellAddress].s = styleToApply;
      }
    }

    // Set column widths
    const colWidths = finalHeaders.map((header: string) => {
      if (header === 'Freeze') return { wch: 8 };
      if (header.includes('Project Name')) return { wch: 30 };
      if (header.includes('Project Reference')) return { wch: 18 };
      if (header.includes('Project Manager') || header.includes('Group Head')) return { wch: 18 };
      if (header.includes('Customer')) return { wch: 20 };
      if (header.includes('Receivables')) return { wch: 18 };
      if (header.includes('Current Normal') || header.includes('Current Retention') || header.includes('Actual')) {
        return { wch: 16 };
      }
      return { wch: 12 };
    });
    ws['!cols'] = colWidths;

    // Freeze header row
    ws['!freeze'] = { x: 0, y: 1 };

    XLSX.utils.book_append_sheet(wb, ws, "Historical Data");

    // ✅ Filename now includes the selected month + year
    const monthNameOut = getMonthName(selectedMonth) || selectedMonth;
    const fileName = `Historical_Collection_Data_${monthNameOut}_${selectedYear}.xlsx`;

    XLSX.writeFile(wb, fileName);

    // ✅ Toast mentions the selected month + year
    toast("success", `Exported ${currentExportData.length} records with 36 months of historical data for ${monthNameOut} ${selectedYear} (Current Normal, Current Retention, Actual)`);

  } catch (error) {
    console.error("❌ Error during historical Excel export:", error);
    toast("error", "Failed to export historical data");
  }
};

const getOBLBreakdownForExport = (collection: CollectionItem): OBLBreakdown => {
  const defaultBreakdown: OBLBreakdown = {
    normalDue: 0,
    normalOverdue: 0,
    retentionDue: 0,
    retentionOverdue: 0,
    gstTds: 0,
    wht: 0
  };

  try {
    const breakdown: OBLBreakdown = {
      normalDue: getNumericFieldValue(collection, 'NormalDue'),
      normalOverdue: getNumericFieldValue(collection, 'NormalOverdue'),
      retentionDue: getNumericFieldValue(collection, 'RetentionDue'),
      retentionOverdue: getNumericFieldValue(collection, 'RetentionOverdue'),
      gstTds: getNumericFieldValue(collection, 'GSTTDS'),
      wht: getNumericFieldValue(collection, 'WHT')
    };

    return breakdown;

  } catch (error) {
    console.error("Error in getOBLBreakdownForExport:", error);
    return defaultBreakdown;
  }
};

const getNumericFieldValue = (collection: CollectionItem, fieldName: string): number => {
  const value = collection[fieldName as keyof CollectionItem];

  if (value === undefined || value === null || value === '') {
    return 0;
  }

  if (typeof value === 'number') {
    return value;
  }

  if (typeof value === 'string') {
    const num = parseFloat(value);
    return isNaN(num) ? 0 : num;
  }

  return 0;
};

// COMPLETE REPLACEMENT FOR setupPaginationEventListeners
const setupPaginationEventListeners = (): void => {

  if (!pageSizeSelect) {
    return;
  }

  pageSizeSelect.addEventListener(
    "change",
    (e) => {

      const newSize = parseInt(
        (e.target as HTMLSelectElement).value,
        10
      );

      changePageSize(newSize);
    }
  );
};

// COMPLETE REPLACEMENT FOR changePageSize
const changePageSize = (newSize: number): void => {

  const validPageSizes = [10, 25, 50, 100];

  if (!validPageSizes.includes(newSize)) {
    newSize = 25;
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

  toast(
    "info",
    `Showing ${pageSize} records per page`
  );
};
// COMPLETE REPLACEMENT FOR renderCurrentPage
const renderCurrentPage = (): void => {
  const tableBody = document.querySelector(".collection-table-body") as HTMLElement;

  if (!tableBody) {
    console.error("Table body element not found");
    return;
  }

  const displayMonths = getDisplayMonths();
  const currentMonthKey = formatMonthKey(new Date());

  const selectedYear = filterYear.value;
  const selectedMonth = filterMonth.value;

  const selectedMonthKey = selectedYear && selectedMonth
    ? `${selectedYear}-${selectedMonth.padStart(2, "0")}`
    : "";

  tableBody.setAttribute("aria-busy", "false");

  // Handle empty state
  if (filteredData.length === 0) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="28" style="text-align:center;padding:20px;border:1px solid #ddd;">
          No records found
        </td>
      </tr>
    `;
    renderTotalsHeader(displayMonths, selectedMonthKey);
    return;
  }

  // Apply current sort to the FILTERED data
  const sortedData = sortCollectionData(filteredData);

  // Ensure currentPage is within bounds
  if (currentPage > totalPages) {
    currentPage = totalPages;
  }

  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, sortedData.length);

  const pageData = sortedData.slice(startIndex, endIndex);

  console.log(`📊 Rendering page ${currentPage}: items ${startIndex + 1} to ${endIndex} of ${sortedData.length}`);
  console.log(`📊 Page data length: ${pageData.length}`);
  console.log(`📊 Using pageSize: ${pageSize}`);

  // Use the existing renderTableBodyOnly function
  renderTableBodyOnly(pageData, displayMonths, currentMonthKey);
  renderTotalsHeader(displayMonths, selectedMonthKey);
};

const updatePaginationControls = (): void => {

  totalPagesSpan.textContent = totalPages.toString();

  if (filteredData.length === 0) {

    startItemSpan.textContent = "0";
    endItemSpan.textContent = "0";
    totalItemsSpan.textContent = "0";

  } else {

    const startItem =
      (currentPage - 1) * pageSize + 1;

    const endItem =
      Math.min(
        currentPage * pageSize,
        filteredData.length
      );

    startItemSpan.textContent =
      startItem.toString();

    endItemSpan.textContent =
      endItem.toString();

    totalItemsSpan.textContent =
      filteredData.length.toString();
  }

  // IMPORTANT: dropdown follows actual pageSize
  if (pageSizeSelect) {
    pageSizeSelect.value = pageSize.toString();
  }

  firstPageBtn.disabled =
    currentPage === 1 ||
    filteredData.length === 0;

  prevPageBtn.disabled =
    currentPage === 1 ||
    filteredData.length === 0;

  nextPageBtn.disabled =
    currentPage >= totalPages ||
    filteredData.length === 0;

  lastPageBtn.disabled =
    currentPage >= totalPages ||
    filteredData.length === 0;

  updatePageNumbers();
};

const goToPage = (page: number): void => {

  if (
    page < 1 ||
    page > totalPages ||
    page === currentPage
  ) {
    return;
  }

  currentPage = page;

  updatePaginationControls();
  renderCurrentPage();
};


const initializePagination = (
  data: CollectionItem[]
): void => {

  filteredData = [...data];

  // Default page size must always be 25
  pageSize = 25;
  currentPage = 1;

  // Keep dropdown synchronized
  if (pageSizeSelect) {
    pageSizeSelect.value = "25";
  }

  calculateTotalPages();

  updatePaginationControls();

  renderCurrentPage();
};
const calculateTotalPages = (): void => {
  totalPages = Math.ceil(filteredData.length / pageSize);

  if (totalPages === 0) {
    totalPages = 1;
  }
};

// COMPLETE REPLACEMENT FOR updatePageNumbers
const updatePageNumbers = (): void => {

  pageNumbersDiv.innerHTML = "";

  if (filteredData.length === 0) {
    return;
  }

  const maxVisiblePages = 5;

  let startPage = Math.max(
    1,
    currentPage - Math.floor(maxVisiblePages / 2)
  );

  let endPage = Math.min(
    totalPages,
    startPage + maxVisiblePages - 1
  );

  if (endPage - startPage + 1 < maxVisiblePages) {
    startPage = Math.max(
      1,
      endPage - maxVisiblePages + 1
    );
  }

  if (startPage > 1) {

    const dotsBtn =
      document.createElement("button");

    dotsBtn.className =
      "collection-page-number collection-dots";

    dotsBtn.textContent = "...";
    dotsBtn.disabled = true;

    pageNumbersDiv.appendChild(dotsBtn);
  }

  for (
    let i = startPage;
    i <= endPage;
    i++
  ) {

    const pageBtn =
      document.createElement("button");

    pageBtn.className =
      `collection-page-number ${
        i === currentPage
          ? "active"
          : ""
      }`;

    pageBtn.textContent = i.toString();

    pageBtn.addEventListener(
      "click",
      () => goToPage(i)
    );

    pageNumbersDiv.appendChild(pageBtn);
  }

  if (endPage < totalPages) {

    const dotsBtn =
      document.createElement("button");

    dotsBtn.className =
      "collection-page-number collection-dots";

    dotsBtn.textContent = "...";
    dotsBtn.disabled = true;

    pageNumbersDiv.appendChild(dotsBtn);
  }
};

const updatePagination = (): void => {
  totalPages = Math.max(
    1,
    Math.ceil(
      filteredData.length / pageSize
    )
  );

  if (
    currentPage > totalPages
  ) {
    currentPage = totalPages;
  }

  if (currentPage < 1) {
    currentPage = 1;
  }

  updatePaginationControls();

  renderCurrentPage();
};


// Updated sorting function with month index support
const handleEnhancedSort = (column: string, direction: 'asc' | 'desc', subColumn?: string, monthIndex?: number): void => {
  // Update current sort state
  currentSortConfig = {
    column,
    direction,
    subColumn,
    monthIndex
  };

  const dataToSort = filteredData.length > 0 ? filteredData : collections;

  if (dataToSort.length === 0) return;

  const sortedData = [...dataToSort].sort((a, b) => {
    let valueA: any = 0;
    let valueB: any = 0;

    switch (column) {
      case 'receivable':
        valueA = a.ReceivableValue || 0;
        valueB = b.ReceivableValue || 0;
        break;

      case 'collection':
        const selectedYear = filterYear.value;
        const selectedMonth = filterMonth.value;
        const selectedMonthKey = selectedYear && selectedMonth ? `${selectedYear}-${selectedMonth.padStart(2, '0')}` : '';

        if (selectedMonthKey) {
          const revenueA = safeJsonParse(a.RevenueActual, {});
          const revenueB = safeJsonParse(b.RevenueActual, {});
          valueA = revenueA[selectedMonthKey] || 0;
          valueB = revenueB[selectedMonthKey] || 0;
        } else {
          valueA = 0;
          valueB = 0;
        }
        break;

      case 'month':
        const displayMonths = getDisplayMonths();

        // Use the specific month index if provided, otherwise default to first month
        const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
        const monthKeyToUse = displayMonths[monthIndexToUse];

        const dataA = safeJsonParse(a.CollectionData, {});
        const dataB = safeJsonParse(b.CollectionData, {});
        const monthDataA = dataA[monthKeyToUse] as CollectionMonthData || { plan: 0, actual: 0, items: {} };
        const monthDataB = dataB[monthKeyToUse] as CollectionMonthData || { plan: 0, actual: 0, items: {} };

        const itemsA = monthDataA.items || {};
        const itemsB = monthDataB.items || {};

        if (subColumn) {
          switch (subColumn) {
            case 'normal-previous':
              valueA = (itemsA.Normal?.plan || 0) + (itemsA['Normal Overdue']?.plan || 0);
              valueB = (itemsB.Normal?.plan || 0) + (itemsB['Normal Overdue']?.plan || 0);
              break;
            case 'normal-current':
              valueA = (itemsA.Normal?.actual || 0) + (itemsA['Normal Overdue']?.actual || 0);
              valueB = (itemsB.Normal?.actual || 0) + (itemsB['Normal Overdue']?.actual || 0);
              break;
            case 'retention-previous':
              valueA = (itemsA.Retention?.plan || 0) + (itemsA['Retention Overdue']?.plan || 0);
              valueB = (itemsB.Retention?.plan || 0) + (itemsB['Retention Overdue']?.plan || 0);
              break;
            case 'retention-current':
              valueA = (itemsA.Retention?.actual || 0) + (itemsA['Retention Overdue']?.actual || 0);
              valueB = (itemsB.Retention?.actual || 0) + (itemsB['Retention Overdue']?.actual || 0);
              break;
            case 'total-previous':
              valueA = monthDataA.plan || 0;
              valueB = monthDataB.plan || 0;
              break;
            case 'total-current':
              valueA = monthDataA.actual || 0;
              valueB = monthDataB.actual || 0;
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

      default:
        return 0;
    }

    if (direction === 'asc') {
      return valueA - valueB;
    } else {
      return valueB - valueA;
    }
  });

  if (filteredData.length > 0) {
    filteredData = sortedData;
  } else {
    collections = sortedData;
  }

  currentPage = 1;
  updatePagination();

  // const sortDirectionText = direction === 'asc' ? 'Ascending' : 'Descending';
  // let columnName = '';

  // if (subColumn) {
  //   columnName = subColumn.replace('-', ' ').replace(/\b\w/g, l => l.toUpperCase());
  // } else if (column === 'receivable') {
  //   columnName = 'Receivable';
  // } else if (column === 'collection') {
  //   columnName = 'Collection';
  // } else {
  //   const displayMonths = getDisplayMonths();
  //   const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
  //   const monthKey = displayMonths[monthIndexToUse];
  //   const monthDate = new Date(monthKey + "-01");
  //   // columnName = formatMonthLabel(monthDate);
  // }
 
};

// Updated sort button listener
/* ✅ Header sorting - ONE ⇅ arrow inside a circle per sortable header.
   A single delegated listener sits on the (persistent) <table>; the header rows
   themselves are rebuilt by renderTableHeader(). Clicking a header sorts by that
   column, a second click flips the direction, and the active column gets the
   tinted circle + solid arrow. */
const updateSortIndicators = (): void => {
  document.querySelectorAll<HTMLElement>('.collection-month-columns-table th.collection-sortable-th').forEach((th) => {
    th.classList.remove('collection-sorted-asc', 'collection-sorted-desc');

    const attrSub = th.getAttribute('data-subcolumn') || undefined;
    const attrMonth = th.getAttribute('data-month-index');
    const monthIndex = attrMonth !== null && attrMonth !== '' ? parseInt(attrMonth, 10) : undefined;

    // ✅ a month column is only "active" when the month INDEX matches too -
    //    otherwise every month's identical sub-header lights up at once.
    const isActive =
      !!currentSortConfig.column &&
      th.getAttribute('data-sort') === currentSortConfig.column &&
      attrSub === (currentSortConfig.subColumn || undefined) &&
      (currentSortConfig.column !== 'month' || monthIndex === (currentSortConfig.monthIndex as number | undefined));

    if (isActive) {
      th.classList.add(currentSortConfig.direction === 'asc' ? 'collection-sorted-asc' : 'collection-sorted-desc');
    }
  });
};

// Updated sort listener - delegated on the table so it survives header re-renders
const attachSortButtonListeners = (): void => {
  const table = document.querySelector('.collection-month-columns-table') as (HTMLElement & { _collectionSortBound?: boolean }) | null;
  if (!table) return;

  if (!table._collectionSortBound) {
    table._collectionSortBound = true;
    table.addEventListener('click', (e) => {
      const th = (e.target as HTMLElement).closest('th.collection-sortable-th') as HTMLElement | null;
      if (!th) return;

      const column = th.getAttribute('data-sort') || '';
      if (!column) return;

      const subColumn = th.getAttribute('data-subcolumn') || undefined;
      const attrMonth = th.getAttribute('data-month-index');
      const monthIndex = attrMonth !== null && attrMonth !== '' ? parseInt(attrMonth, 10) : undefined;

      const isSame =
        currentSortConfig.column === column &&
        (currentSortConfig.subColumn || undefined) === subColumn &&
        (currentSortConfig.monthIndex as number | undefined) === monthIndex;

      const direction: 'asc' | 'desc' = isSame && currentSortConfig.direction === 'asc' ? 'desc' : 'asc';

      handleEnhancedSort(column, direction, subColumn, monthIndex);
      updateSortIndicators();
    });
  }

  updateSortIndicators();
};

// ========== FIXED RENDER TABLE WITH MONTH COLUMNS - SHOW REVENUE ACTUAL ==========
const renderTableBodyOnly = (data: CollectionItem[], displayMonths: string[], currentMonthKey: string): void => {
  const tableBody = document.querySelector(".collection-table-body") as HTMLElement;

  if (!tableBody) return;

  if (!data || data.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="21" style="text-align:center;padding:20px;border:1px solid #ddd;">No records found</td></tr>`;
    return;
  }

  let tableBodyHTML = '';

  data.forEach((collection) => {
    const projectTitle = collection.Project?.Title || "-";
    const proj = projects.find(p => p.Title === projectTitle);

    const collectionData = safeJsonParse(collection.CollectionData, {});
    const totalReceivable = collection.ReceivableValue || 0;
    const formattedOBL = Math.round(totalReceivable).toLocaleString('en-US');

    const isFrozen = collection.Freeze || false;
    const rowStyle = isFrozen ? 'background-color: #fff8e1;' : '';

    const canUserEdit = authCanEditProject(proj);
    const canUserFreeze = canFreeze();
    const isCheckboxDisabled = !canUserFreeze;

    const projectRef = getSafeString(proj?.ProjectReference);
    const projectRefLink = projectRef !== "-"
      ? (canUserEdit
          ? `<a href="#" class="collection-edit-trigger" data-id="${collection.Id}" data-project="${collection.Project?.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectRef}</a>`
          : `<span style="color:#6c757d; cursor:default;">${projectRef}</span>`)
      : '<span style="color:#999;">-</span>';

    const customerName = getSafeString(proj?.CustomerName);
    const groupHead = getGroupHeadDisplayName(proj?.GroupHead1);
    const pmName = getProjectManagerDisplayName(proj?.ProjectManager);

    const projectNameLink = `<span style="color:#333;">${projectTitle}</span>`;

    const freezeBadge = isFrozen ? ' 🔒' : '';

    // Table column order: Freeze | Project Ref | Customer | Project | Group Head | Project Manager | OBL
    let rowHTML = `
      <tr style="${rowStyle}" data-collection-id="${collection.Id}">
        <td class="collection-freeze-col" style="text-align:center; vertical-align:middle; min-width:40px; max-width:40px; width:40px;">
          <input type="checkbox" class="collection-freeze-checkbox"
                data-id="${collection.Id}"
                data-freeze="${isFrozen}"
                ${isFrozen ? 'checked' : ''}
                ${isCheckboxDisabled ? 'disabled' : ''}
                style="pointer-events: ${isCheckboxDisabled ? 'none' : 'auto'} !important;
                       cursor: ${isCheckboxDisabled ? 'not-allowed' : 'pointer'} !important;
                       position: relative; z-index: 100;
                       -webkit-appearance: checkbox !important;
                       appearance: checkbox !important;
                       width: 18px; height: 18px;"
                title="${canUserFreeze ? 'Toggle freeze status' : 'Freeze permission restricted'}" />
        </td>
        <td class="collection-project-ref-col" style="min-width:120px; max-width:120px; width:120px; word-wrap:break-word; white-space:normal;">${projectRefLink}</td>
        <td class="collection-customer-col" style="min-width:130px; max-width:130px; width:130px; word-wrap:break-word; white-space:normal;">${customerName}</td>
        <td class="collection-project-col" style="min-width:180px; max-width:180px; width:180px; word-wrap:break-word; white-space:normal;">${projectNameLink}${freezeBadge}</td>
        <td class="collection-gh-col" style="min-width:100px; max-width:100px; width:100px; word-wrap:break-word; white-space:normal;">${groupHead}</td>
<td class="collection-pm-col" style="min-width:100px; max-width:100px; width:100px; word-wrap:break-word; white-space:normal;">${pmName}</td>
        <td class="collection-obl-col" style="min-width:100px; max-width:100px; width:100px; text-align:center; word-wrap:break-word; white-space:normal;">${formattedOBL}</td>
    `;

    // Monthly data columns - display ONLY numeric data, no category names
    displayMonths.forEach((monthKey) => {
      const monthData = collectionData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, forecast: 0, items: {} };
      const isCurrentMonth = monthKey === currentMonthKey;

      const items = monthData.items || {};

      // Previous columns
      const prevNormal = (items.Normal?.plan || 0) + (items['Normal Overdue']?.plan || 0);
      const prevRetention = (items.Retention?.plan || 0) + (items['Retention Overdue']?.plan || 0);
      const prevTotal = prevNormal + prevRetention;

      // Current columns
      const currentNormal = (items.Normal?.actual || 0) + (items['Normal Overdue']?.actual || 0);
      const currentRetention = (items.Retention?.actual || 0) + (items['Retention Overdue']?.actual || 0);
      const currentTotal = currentNormal + currentRetention;

      // Add Previous columns
      rowHTML += `
        <td class="collection-month-scroll-col collection-prev-cell" style="min-width:55px; text-align:center;">${prevNormal.toLocaleString('en-US')}</td>
        <td class="collection-month-scroll-col collection-prev-cell" style="min-width:55px; text-align:center;">${prevRetention.toLocaleString('en-US')}</td>
        <td class="collection-month-scroll-col collection-prev-cell" style="min-width:55px; text-align:center; font-weight:bold;">${prevTotal.toLocaleString('en-US')}</td>
      `;

      // Add Current columns
      rowHTML += `
        <td class="collection-month-scroll-col" style="min-width:55px; text-align:center;">${currentNormal.toLocaleString('en-US')}</td>
        <td class="collection-month-scroll-col" style="min-width:55px; text-align:center;">${currentRetention.toLocaleString('en-US')}</td>
        <td class="collection-month-scroll-col" style="min-width:55px; text-align:center; font-weight:bold;">${currentTotal.toLocaleString('en-US')}</td>
      `;

      // Current month only: Actual (RevenueActual), Further, Forecast
      if (isCurrentMonth) {
        let revenueActualData: { [key: string]: number } = {};
        try {
          if (collection._revenueActual && typeof collection._revenueActual === 'object') {
            revenueActualData = collection._revenueActual;
          } else if (typeof collection.RevenueActual === 'string') {
            const parsed = JSON.parse(collection.RevenueActual);
            if (typeof parsed === 'object' && parsed !== null) {
              revenueActualData = parsed;
            }
          } else if (typeof collection.RevenueActual === 'object') {
            revenueActualData = collection.RevenueActual || {};
          }
        } catch (e) {
          revenueActualData = {};
        }

        // Actual = RevenueActual
        const actualValue = Number(revenueActualData[monthKey]) || 0;
        const furtherValue = monthData.forecast || 0;
        // Forecast = Actual + Further
        const forecastValue = actualValue + furtherValue;

        rowHTML += `
          <td class="collection-month-scroll-col" style="min-width:55px; text-align:center; background-color: #e8f5e9; font-weight: bold; color: #1a8a34;">${actualValue.toLocaleString('en-US')}</td>
          <td class="collection-month-scroll-col" style="min-width:55px; text-align:center;">${furtherValue.toLocaleString('en-US')}</td>
          <td class="collection-month-scroll-col collection-total-cell" style="min-width:55px; text-align:right; font-weight: bold; background-color: #f0f8ff; color: #0c5460;">${forecastValue.toLocaleString('en-US')}</td>
        `;
      }
    });

    rowHTML += `</tr>`;
    tableBodyHTML += rowHTML;
  });

  tableBody.innerHTML = tableBodyHTML;

  // Re-attach event listeners
  attachFreezeEventListeners();
  attachEditTriggerListeners();
};

// Replace the renderTableHeader function
const renderTableHeader = (): void => {
  const monthColumnsHeader = document.querySelector("#monthColumnsHeader") as HTMLElement;
  if (!monthColumnsHeader) return;

  const displayMonths = getDisplayMonths();
  const currentMonthKey = formatMonthKey(new Date());
  const headerHTML = buildHeaderHTML(displayMonths, currentMonthKey);
  monthColumnsHeader.innerHTML = headerHTML;
  attachSortButtonListeners();
  applyStickyHeaderOffsets();
};    
  
/* ✅ Sortable headers - ONE ⇅ arrow inside a circle (mock-up style).
   The <th> itself carries data-sort / data-subcolumn / data-month-index and the
   delegated listener on the table turns a click into asc → desc → asc. */
const buildHeaderHTML = (displayMonths: string[], currentMonthKey: string): string => {
  const sortArrows = `<span class="collection-sort-arrows" aria-hidden="true"><span class="collection-arrow-up">▲</span><span class="collection-arrow-down">▼</span></span>`;

  // ── ROW 1: sticky columns (rowspan=3) + month names ──
  let html = `
    <tr>
      <th rowspan="3" class="collection-freeze-col" title="Freeze / Unfreeze">
        <span class="collection-header-label">Lock</span>
      </th>
      <th rowspan="3" class="collection-project-ref-col collection-sortable-th" data-sort="projectReference" title="Sort by Project Reference">
        <div class="collection-sort-header-container">
          <span class="collection-header-label">Project Reference</span>
          ${sortArrows}
        </div>
      </th>
      <th rowspan="3" class="collection-customer-col collection-sortable-th" data-sort="customer" title="Sort by Customer">
        <div class="collection-sort-header-container">
          <span class="collection-header-label">Customer</span>
          ${sortArrows}
        </div>
      </th>
      <th rowspan="3" class="collection-project-col collection-sortable-th" data-sort="project" title="Sort by Project">
        <div class="collection-sort-header-container">
          <span class="collection-header-label">Project</span>
          ${sortArrows}
        </div>
      </th>
      <th rowspan="3" class="collection-gh-col collection-sortable-th" data-sort="groupHead" title="Sort by Group Head">
        <div class="collection-sort-header-container">
          <span class="collection-header-label">Group Head</span>
          ${sortArrows}
        </div>
      </th>
      <th rowspan="3" class="collection-pm-col collection-sortable-th" data-sort="projectManager" title="Sort by Project Manager">
        <div class="collection-sort-header-container">
          <span class="collection-header-label">Project Manager</span>
          ${sortArrows}
        </div>
      </th>
      <th rowspan="3" class="collection-obl-col collection-sortable-th" data-sort="receivable" title="Sort by Receivables (TDS)">
        <div class="collection-sort-header-container">
          <span class="collection-header-label">Receivables (TDS)</span>
          ${sortArrows}
        </div>
      </th>
  `;

  // Month name cells (row 1)
  displayMonths.forEach((month, monthIndex) => {
    const isCurrentMonth = month === currentMonthKey;
    const monthDate = new Date(month + "-01");
    const monthLabel = formatMonthLabel(monthDate);
    const colspan = isCurrentMonth ? 9 : 6;   // 3 prev + 3 curr (+3 actual/further/forecast)

    html += `
      <th colspan="${colspan}" class="collection-month-column-header ${isCurrentMonth ? 'collection-current-month-header' : ''} collection-sortable-th" data-sort="month" data-month-index="${monthIndex}" title="Sort by ${monthLabel}">
        <div class="collection-sort-header-container" style="display:inline-flex; align-items:center; gap:4px;">
          <span class="collection-header-label">${monthLabel}${isCurrentMonth ? ' (Current)' : ''}</span>
          ${sortArrows}
        </div>
      </th>
    `;
  });
  html += `</tr>`;

  // ── ROW 2: Previous / Current group headers ──
  html += `<tr>`;
  displayMonths.forEach((month) => {
    const isCurrentMonth = month === currentMonthKey;
    html += `
      <th colspan="3" class="collection-month-group-header collection-prev-group">Previous</th>
      <th colspan="${isCurrentMonth ? 6 : 3}" class="collection-month-group-header collection-curr-group">Current</th>
    `;
  });
  html += `</tr>`;

  // ── ROW 3: Normal / Retention / Total sub-columns (each sortable) ──
  const subCol = (label: string, sortKey: string, monthIndex: number) => `
    <th class="collection-month-subheader collection-sortable-th" data-sort="month" data-subcolumn="${sortKey}" data-month-index="${monthIndex}" title="Sort by ${label}">
      <div class="collection-sort-header-container" style="display:inline-flex; align-items:center; gap:2px; width:63px;">
        <span class="collection-header-label">${label}</span>
        ${sortArrows}
      </div>
    </th>
  `;

  html += `<tr>`;
  displayMonths.forEach((month, monthIndex) => {
    const isCurrentMonth = month === currentMonthKey;

    // Previous group
    html += subCol('Normal', 'normal-previous', monthIndex);
    html += subCol('Ret.n', 'retention-previous', monthIndex);
    html += subCol('Total', 'total-previous', monthIndex);

    // Current group
    html += subCol('Normal', 'normal-current', monthIndex);
    html += subCol('Ret.n', 'retention-current', monthIndex);
    html += subCol('Total', 'total-current', monthIndex);

    // Current month extras
    if (isCurrentMonth) {
      html += subCol('Actual', 'actual', monthIndex);
      html += subCol('Further', 'forecast', monthIndex);
      html += subCol('Forecast', 'forecast-total', monthIndex);
    }
  }); 
  html += `</tr>`;

  return html;
};

// ========== FIXED RENDER TOTALS HEADER ==========        
const renderTotalsHeader = (  displayMonths: string[],selectedMonthKey: string ): void => {   
  const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;

  if (!totalsHeader) { 
    console.error("Totals header element not found");
    return;
  }

  const allData = collections;

  const hasActiveFilters =
    filterYear.value !== "" ||
    filterMonth.value !== "" ||
    filterProject.value !== "" ||
    filterCustomer.value !== "" ||
    filterProjectManager.value !== "" ||
    filterGroupHead.value !== "" ||
    filterProjectReference.value !== "";

  const filtered = hasActiveFilters
    ? filteredData
    : collections;

  const allTotals = calculateTotals(
    allData,
    displayMonths,
    selectedMonthKey
  );

  const filteredTotals = calculateTotals(
    filtered,
    displayMonths,
    selectedMonthKey
  );

  const buildTotalsRow = (
    label: string,
    totals: any,
    cssClass: string,
    count: number
  ): string => {
    let rowHTML = `
      <tr class="${cssClass}">
        <td
          class="collection-freeze-col"
          style="
            min-width:40px;
            max-width:40px;
            width:40px;
            text-align:center;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="collection-project-ref-col"
          style="
            min-width:120px;
            max-width:120px;
            width:120px;
            box-sizing:border-box;
          "
        >
          <strong>${label}</strong>
        </td>

        <td
          class="collection-customer-col"
          style="
            min-width:130px;
            max-width:130px;
            width:130px;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="collection-project-col"
          style="
            min-width:180px;
            max-width:180px;
            width:180px;
            box-sizing:border-box;
          "
        >
          <span style="font-size:11px; color:#6c757d;">
            ${
              label === "All Projects"
                ? `All Projects (${count})`
                : `Filtered (${count})`
            }
          </span>
        </td>

        <td
          class="collection-gh-col"
          style="
            min-width:100px;
            max-width:100px;
            width:100px;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="collection-pm-col"
          style="
            min-width:100px;
            max-width:100px;
            width:100px;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="collection-obl-col"
          style="
            min-width:100px;
            max-width:100px;
            width:100px;
            text-align:center;
            font-weight:bold;
            color:#0078d4;
            box-sizing:border-box;
          "
        >
          ${Math.round(totals.receivable || 0).toLocaleString('en-US')}
        </td>
    `;

    const currentMonthKey = formatMonthKey(new Date());

    displayMonths.forEach((monthKey) => {
      const t = totals.monthly[monthKey] || {
        normalPrevious: 0,
        retentionPrevious: 0,
        totalPrevious: 0,
        normalCurrent: 0,
        retentionCurrent: 0,
        totalCurrent: 0,
        actual: 0,
        further: 0,
        forecast: 0
      };

      const isCurrentMonth = monthKey === currentMonthKey;

      // Previous
      rowHTML += `
        <td
          class="collection-month-scroll-col collection-prev-cell"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:center;
            font-weight:bold;
            color:#0056b3;
          "
        >
          ${Math.round(t.normalPrevious || 0).toLocaleString('en-US')}
        </td>

        <td
          class="collection-month-scroll-col collection-prev-cell"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:center;
            font-weight:bold;
            color:#0056b3;
          "
        >
          ${Math.round(t.retentionPrevious || 0).toLocaleString('en-US')}
        </td>

        <td
          class="collection-month-scroll-col collection-prev-cell"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:center;
            font-weight:bold;
            color:#0078d4;
          "
        >
          ${Math.round(t.totalPrevious || 0).toLocaleString('en-US')}
        </td>
      `;

      // Current
      rowHTML += `
        <td
          class="collection-month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:center;
            font-weight:bold;
            color:#28a745;
          "
        >
          ${Math.round(t.normalCurrent || 0).toLocaleString('en-US')}
        </td>

        <td
          class="collection-month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:center;
            font-weight:bold;
            color:#28a745;
          "
        >
          ${Math.round(t.retentionCurrent || 0).toLocaleString('en-US')}
        </td>

        <td
          class="collection-month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:center;
            font-weight:bold;
            color:#1a8a34;
          "
        >
          ${Math.round(t.totalCurrent || 0).toLocaleString('en-US')}
        </td>
      `;

      // Current month: Actual / Further / Forecast
      if (isCurrentMonth) {
        rowHTML += `
          <td
            class="collection-month-scroll-col"
            style="
              min-width:55px;
              box-sizing:border-box;
              text-align:center;
              font-weight:bold;
              color:#856404;
              background-color:#fff3cd;
            "
          >
            ${Math.round(t.actual || 0).toLocaleString('en-US')}
          </td>

          <td
            class="collection-month-scroll-col"
            style="
              min-width:55px;
              box-sizing:border-box;
              text-align:center;
              font-weight:bold;
              color:#856404;
            "
          >
            ${Math.round(t.further || 0).toLocaleString('en-US')}
          </td>

          <td
            class="collection-month-scroll-col"
            style="
              min-width:55px;
              box-sizing:border-box;
              text-align:center;
              font-weight:bold;
              color:#0c5460;
              background-color:#d1ecf1;
            "
          >
            ${Math.round(t.forecast || 0).toLocaleString('en-US')}
          </td>
        `;
      }
    });

    rowHTML += `</tr>`;

    return rowHTML;
  };

  // 7 fixed columns
  const totalColumns =
    7 +
    displayMonths.length * 6 +
    (displayMonths.includes(formatMonthKey(new Date())) ? 3 : 0);

  totalsHeader.innerHTML = `
    <tr>
      <th
        colspan="${totalColumns}"
        style="
          background:#f8f9fa;
          padding:8px;
          text-align:center;
          font-size:13px;
          font-weight:bold;
          color:#495057;
          border:1px solid #dee2e6;
          box-sizing:border-box;
        "
      >
        📊 Collection Totals
      </th>
    </tr>

    ${buildTotalsRow(
      "All Projects",
      allTotals,
      "all-projects-total",
      allData.length
    )}

    ${buildTotalsRow(
      "Filtered", 
      filteredTotals,
      "filtered-total",
      filtered.length
    )}
  `; 

   applyStickyHeaderOffsets();
};

 const calculateTotals = (
  data: CollectionItem[],
  displayMonths: string[],
  selectedMonthKey: string
): any => {
  const dataToCalculate = data;

  const totals = {
    receivable: 0,
    actualCollection: 0,
    monthly: {} as {
      [key: string]: {
        normalPrevious: number;
        retentionPrevious: number;
        totalPrevious: number;
        normalCurrent: number;
        retentionCurrent: number;
        totalCurrent: number;
        actual: number;
        further: number;
        forecast: number;
      }
    }
  };

  // Init months
  displayMonths.forEach(monthKey => {
    totals.monthly[monthKey] = {
      normalPrevious: 0,
      retentionPrevious: 0,
      totalPrevious: 0,
      normalCurrent: 0,
      retentionCurrent: 0,
      totalCurrent: 0,
      actual: 0,
      further: 0,
      forecast: 0
    };
  });

  dataToCalculate.forEach(collection => {
    totals.receivable += Math.round(Number(collection.ReceivableValue) || 0);

    // Parse RevenueActual ONCE per collection (was parsed inside the month loop)
    let revenueActualData: { [key: string]: number } = {};
    try {
      if (typeof collection.RevenueActual === 'string') {
        revenueActualData = JSON.parse(collection.RevenueActual) || {};
      } else if (typeof collection.RevenueActual === 'object') {
        revenueActualData = collection.RevenueActual || {};
      }
    } catch (e) {
      revenueActualData = {};
    }

    if (selectedMonthKey) {
      totals.actualCollection += Math.round(Number(revenueActualData[selectedMonthKey]) || 0);
    }

    const collectionData = safeJsonParse(collection.CollectionData, {});

    displayMonths.forEach(monthKey => {
      const m = collectionData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, items: {} };
      const items = m.items || {};

      const normalPrevious = Math.round((items.Normal?.plan || 0) + (items['Normal Overdue']?.plan || 0));
      const retentionPrevious = Math.round((items.Retention?.plan || 0) + (items['Retention Overdue']?.plan || 0));

      // ✅ FIX: Total = Normal + Retention (same formula as the body rows)
      const totalPrevious = normalPrevious + retentionPrevious;

      const normalCurrent = Math.round((items.Normal?.actual || 0) + (items['Normal Overdue']?.actual || 0));
      const retentionCurrent = Math.round((items.Retention?.actual || 0) + (items['Retention Overdue']?.actual || 0));

      // ✅ FIX: Total = Normal + Retention (was m.actual — caused the mismatch)
      const totalCurrent = normalCurrent + retentionCurrent;

      const actual = Math.round(Number(revenueActualData[monthKey]) || 0);
      const further = Math.round(m.forecast || 0);
      const forecast = actual + further;

      totals.monthly[monthKey].normalPrevious += normalPrevious;
      totals.monthly[monthKey].retentionPrevious += retentionPrevious;
      totals.monthly[monthKey].totalPrevious += totalPrevious;
      totals.monthly[monthKey].normalCurrent += normalCurrent;
      totals.monthly[monthKey].retentionCurrent += retentionCurrent;
      totals.monthly[monthKey].totalCurrent += totalCurrent;
      totals.monthly[monthKey].actual += actual;
      totals.monthly[monthKey].further += further;
      totals.monthly[monthKey].forecast += forecast;
    });
  });

  return totals;
}; 


      // ========== FORM MANAGEMENT FUNCTIONS ==========

const showForm = (isEditMode: boolean = false): void => {
    formSection.style.display = "block";
    tableSection.style.display = "none";
    showFormBtn.style.display = "none";

    if (isEditMode) {
        formSection.classList.add('collection-edit-mode');
        projectSelect.disabled = true;
        formTitle.textContent = "Collection Plan 5QRF";
    } else {
        formSection.classList.remove('collection-edit-mode');
        projectSelect.disabled = false;
        formTitle.textContent = "Collection Plan 5QRF";
    }
};

      const hideForm = (): void => {
        formSection.style.display = "none";
        tableSection.style.display = "block";
        showFormBtn.style.display = "inline-block";
        formSection.classList.remove('collection-edit-mode');
        projectSelect.disabled = false;
      };

const resetForm = (): void => {
  editingId = null;

  form.reset();

  projectSelect.value = "";

  const projectNameDisplay =
    document.querySelector("#projectNameDisplay") as HTMLElement | null;

  const projectRef =
    document.querySelector("#projectRef") as HTMLElement | null;

  const profitCenter =
    document.querySelector("#profitCenter") as HTMLElement | null;

  const projectManager =
    document.querySelector("#projectManager") as HTMLElement | null;

  const groupHead =
    document.querySelector("#groupHead") as HTMLElement | null;

  const customer =
    document.querySelector("#customer") as HTMLElement | null;

  if (projectNameDisplay) {
    projectNameDisplay.textContent = "-";
  }

  if (projectRef) {
    projectRef.textContent = "-";
  }

  if (profitCenter) {
    profitCenter.textContent = "-";
  }

  if (projectManager) {
    projectManager.textContent = "-";
  }

  if (groupHead) {
    groupHead.textContent = "-";
  }

  if (customer) {
    customer.textContent = "-";
  }

  // Clear collection data containers
  historicalContainer.innerHTML = "";
  firstRowContainer.innerHTML = "";
  remainingContainer.innerHTML = "";
  yearlyContainer.innerHTML = "";

  // Reset remarks
  const remarksInternal =
    document.querySelector(
      "#remarksInternal"
    ) as HTMLTextAreaElement | null;

  const remarksSprintCell =
    document.querySelector(
      "#remarksSprintCell"
    ) as HTMLTextAreaElement | null;

  if (remarksInternal) {
    remarksInternal.value = "";
  }

  if (remarksSprintCell) {
    remarksSprintCell.value = "";
  }

  // Reset breakdown fields
  const breakdownFields = [
    "oblNormalDue",
    "oblNormalOverdue",
    "oblRetentionDue",
    "oblRetentionOverdue",
    "oblGstTds",
    "oblWht"
  ];

  breakdownFields.forEach((id) => {
    const element =
      document.getElementById(
        id
      ) as HTMLInputElement | null;

    if (element) {
      element.value = "";

      element.className =
        element.className.replace(
          "collection-input-with-value",
          "collection-input-empty"
        );

      if (!element.className.includes("collection-input-empty")) {
        element.classList.add("collection-input-empty");
      }

      element.placeholder = "0";
    }
  });

  // Hide collection sections
  toggleCollectionSections(false);

  const oblContainer =
    document.querySelector(
      ".collection-obl-section-container"
    ) as HTMLElement | null;

  if (oblContainer) {
    oblContainer.className =
      "collection-obl-section-container collection-hidden-section";
  }

  oblLastMonthInput.removeAttribute(
    "readonly"
  );

  oblLastMonthInput.placeholder =
    "Enter OBL";

  setNumericValue(
    oblLastMonthInput,
    0
  );

  const fieldNote =
    oblContainer?.querySelector(
      ".field-note"
    ) as HTMLElement | null;

  if (fieldNote) {
    fieldNote.textContent =
      "For new projects: Enter initial Receivables value";
  }

  oblValidationMessage.textContent =
    "✅ All values are within Receivables limit";

  oblValidationMessage.className =
    "collection-validation-message";

  updateOBLValidationDisplay();
};

      // ========== VALIDATION FUNCTIONS ==========
const calculateTotalCurrentValue = (): number => {
  let total = 0;

  const allEditableMonths = allMonths.slice(36);

  allEditableMonths.forEach(monthKey => {
    const currentInput = document.querySelector<HTMLInputElement>(
      `input.collection-main-current-input[data-month="${monthKey}"]:not([readonly])`
    );
    if (currentInput) {
      total += getNumericValue(currentInput);
    }
  });

  const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>(
    `input.collection-yearly-current-input:not([readonly])`
  );
  yearlyCurrentInputs.forEach(input => {
    total += getNumericValue(input);
  });

  return total;
};


const validateOBLLimit = (changedInput?: HTMLInputElement): boolean => {
  const oblValue = getNumericValue(oblLastMonthInput);
  const totalCurrent = calculateTotalCurrentValue();

  // Update displays
  totalCurrentValueSpan.textContent = totalCurrent.toLocaleString('en-US');
  const remaining = oblValue - totalCurrent;
  remainingValueSpan.textContent = remaining.toLocaleString('en-US');

  // Update progress bar
  if (oblValue > 0) {
    const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
    oblLineFill.style.width = `${percentage}%`;

    oblLineFill.className = 'collection-line-fill';
    if (percentage > 90) {
      oblLineFill.classList.add('collection-danger');
    } else if (percentage > 75) {
      oblLineFill.classList.add('collection-warning');
    }
  } else {
    oblLineFill.style.width = `0%`;
    oblLineFill.className = 'collection-line-fill';
  }

  // Check if OBL is zero but values exist
  if (oblValue === 0 && totalCurrent > 0) {
    showOBLErrorToast('⚠️ Receivables is zero but liquidation values exist. Please set Receivables value.', true);
    blockAllCurrentInputs();

    oblValidationMessage.textContent = `❌ Receivables is zero but liquidation values exist (${totalCurrent.toLocaleString('en-US')})`;
    oblValidationMessage.className = 'collection-validation-message danger';
    return false;
  }

  // Check if total exceeds OBL (Liquidation > Receivables)
  if (totalCurrent > oblValue) {
    const exceedAmount = totalCurrent - oblValue;
    const lastToast = (window as any)._lastOBLToast;
    const now = Date.now();
    if (!lastToast || now - lastToast > 3000) {
      toast("error", `❌ Liquidation (${totalCurrent.toLocaleString('en-US')}) exceeds Receivables (${oblValue.toLocaleString('en-US')}) by ${exceedAmount.toLocaleString('en-US')}! Reduce liquidation values to continue.`);
      (window as any)._lastOBLToast = now;
    }
    oblValidationMessage.textContent = `❌ LIQUIDATION EXCEEDS RECEIVABLES: ${totalCurrent.toLocaleString('en-US')} > ${oblValue.toLocaleString('en-US')} (Excess: ${exceedAmount.toLocaleString('en-US')})`;
    oblValidationMessage.className = 'collection-validation-message danger';

    // Block ALL inputs when OBL is exceeded
    blockAllCurrentInputs();
    return false;
  }

  // If within OBL limit, unblock all inputs
  unblockAllCurrentInputs();

  // Update status message based on relationship
  if (oblValue > 0 && totalCurrent === 0) {
    // No liquidation yet
    oblValidationMessage.textContent = `ℹ️ No liquidation recorded yet. Receivables: ${oblValue.toLocaleString('en-US')} - Awaiting liquidation`;
    oblValidationMessage.className = 'collection-validation-message info';

  } else if (totalCurrent === oblValue && oblValue > 0) {
    // Fully liquidated
    const lastToast = (window as any)._lastOBLToast;
    const now = Date.now();
    if (!lastToast || now - lastToast > 3000) {
      toast("success", `✅ Receivables fully liquidated: ${totalCurrent.toLocaleString('en-US')} = ${oblValue.toLocaleString('en-US')}`);
      (window as any)._lastOBLToast = now;
    }
    oblValidationMessage.textContent = `✅ RECEIVABLES FULLY LIQUIDATED: ${totalCurrent.toLocaleString('en-US')} = ${oblValue.toLocaleString('en-US')}`;
    oblValidationMessage.className = 'collection-validation-message success';

  } else if (totalCurrent < oblValue && totalCurrent > 0) {
    // Partial liquidation
    const percentage = (totalCurrent / oblValue) * 100;
    const remainingBalance = oblValue - totalCurrent;

    if (percentage > 90) {
      const lastToast = (window as any)._lastOBLToast;
      const now = Date.now();
      if (!lastToast || now - lastToast > 3000) {
        toast("warning", `⚠️ Receivables nearly fully liquidated! ${percentage.toFixed(1)}% used (${remainingBalance.toLocaleString('en-US')} remaining)`);
        (window as any)._lastOBLToast = now;
      }
      oblValidationMessage.textContent = `⚠️ RECEIVABLES NEARLY FULLY LIQUIDATED: ${percentage.toFixed(1)}% | ${totalCurrent.toLocaleString('en-US')} / ${oblValue.toLocaleString('en-US')} (${remainingBalance.toLocaleString('en-US')} remaining)`;
      oblValidationMessage.className = 'collection-validation-message warning';

    } else if (percentage > 75) {
      const lastToast = (window as any)._lastOBLToast;
      const now = Date.now();
      if (!lastToast || now - lastToast > 5000) {
        // toast("info", `ℹ️ Receivables ${percentage.toFixed(1)}% liquidated (${remainingBalance.toLocaleString('en-US')} remaining)`);
        (window as any)._lastOBLToast = now;
      }
      oblValidationMessage.textContent = `ℹ️ RECEIVABLES PARTIALLY LIQUIDATED: ${percentage.toFixed(1)}% | ${totalCurrent.toLocaleString('en-US')} / ${oblValue.toLocaleString('en-US')} (${remainingBalance.toLocaleString('en-US')} remaining)`;
      oblValidationMessage.className = 'collection-validation-message info';

    } else {
      // Less than 75% liquidated
      const lastToast = (window as any)._lastOBLToast;
      const now = Date.now();
      if (!lastToast || now - lastToast > 5000) {
        // toast("info", `ℹ️ Receivables ${percentage.toFixed(1)}% liquidated (${remainingBalance.toLocaleString('en-US')} remaining)`);
        (window as any)._lastOBLToast = now;
      }
      oblValidationMessage.textContent = `ℹ️ RECEIVABLES PARTIALLY LIQUIDATED: ${percentage.toFixed(1)}% | ${totalCurrent.toLocaleString('en-US')} / ${oblValue.toLocaleString('en-US')} (${remainingBalance.toLocaleString('en-US')} remaining)`;
      oblValidationMessage.className = 'collection-validation-message info';
    }

  } else if (oblValue === 0 && totalCurrent === 0) {
    // No Receivables set
    // oblValidationMessage.textContent = `ℹ️ No Receivables value set. Please set Receivables to track liquidation.`;
    oblValidationMessage.className = 'collection-validation-message info';
  }

  return true;
};


const blockAllCurrentInputs = (): void => {
  const currentInputs = document.querySelectorAll<HTMLInputElement>(
    'input.collection-main-current-input:([readonly]), input.collection-yearly-current-input:([readonly])'
  );  

  currentInputs.forEach(input => {
    input.setAttribute('readonly', 'true');
    input.style.backgroundColor = '#f8f9fa';
    input.style.cursor = 'not-allowed';
    input.style.opacity = '0.7';
  });
};

const unblockAllCurrentInputs = (): void => {
  const currentInputs = document.querySelectorAll<HTMLInputElement>(
    'input.collection-main-current-input[readonly], input.collection-yearly-current-input[readonly]'
  );

  currentInputs.forEach(input => {
    input.removeAttribute('readonly');
    input.style.backgroundColor = '';
    input.style.cursor = '';
    input.style.opacity = '1';
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

  // Check if adding this increase would exceed OBL
  return (currentTotal + increaseAmount) <= oblValue;
};

const validateSingleMonthCategories = (monthKey: string, changedInput?: HTMLInputElement): boolean => {
  const mainCurrentInput = document.querySelector<HTMLInputElement>(`input.collection-main-current-input[data-month="${monthKey}"]`);
  const categoryCurrentInputs = document.querySelectorAll<HTMLInputElement>(`input.collection-category-current-input[data-month="${monthKey}"]`);

  if (!mainCurrentInput) return true;

  const mainCurrentValue = getNumericValue(mainCurrentInput);
  let categoriesTotal = 0;

  categoryCurrentInputs.forEach(input => {
    categoriesTotal += getNumericValue(input);
  });

  // Find or create validation display - SINGLE one per month
  let validationDisplay = document.querySelector(`.collection-month-total-validation[data-month="${monthKey}"]`) as HTMLDivElement;
  if (!validationDisplay) {
    validationDisplay = document.createElement('div');
    validationDisplay.className = 'collection-month-total-validation';
    validationDisplay.setAttribute('data-month', monthKey);
    validationDisplay.style.marginTop = '8px';
    validationDisplay.style.fontSize = '11px';
    validationDisplay.style.textAlign = 'center';

    const monthBlock = mainCurrentInput.closest('.collection-month-box, .collection-detailed-month-box, .collection-simple-month');
    if (monthBlock) {
      const existingValidations = monthBlock.querySelectorAll('.collection-month-total-validation');
      existingValidations.forEach(el => el.remove());

      const oldMessages = monthBlock.querySelectorAll('.validation-message-text');
      oldMessages.forEach(el => el.remove());

      monthBlock.appendChild(validationDisplay);
    }
  }

  const monthBlock = mainCurrentInput.closest('.collection-month-box, .collection-detailed-month-box, .collection-simple-month');
  if (monthBlock) {
    const oldMessages = monthBlock.querySelectorAll('.validation-message-text');
    oldMessages.forEach(el => el.remove());
  }

  if (categoriesTotal > mainCurrentValue) {
    blockIncreasingMonthCategoryInputs(monthKey, changedInput);
    return false;
  } else if (categoriesTotal < mainCurrentValue) {
    unblockAllMonthCategoryInputs(monthKey);
    return true;
  } else {
    unblockAllMonthCategoryInputs(monthKey);
    return true;
  }
};

      const blockIncreasingMonthCategoryInputs = (monthKey: string, allowedInput?: HTMLInputElement): void => {
        const categoryInputs = document.querySelectorAll<HTMLInputElement>(`input.collection-category-current-input[data-month="${monthKey}"]`);

        categoryInputs.forEach(input => {
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

      const unblockAllMonthCategoryInputs = (monthKey: string): void => {
        const categoryInputs = document.querySelectorAll<HTMLInputElement>(`input.collection-category-current-input[data-month="${monthKey}"]`);

        categoryInputs.forEach(input => {
          input.style.backgroundColor = '';
          input.style.cursor = '';
          input.removeAttribute('data-block-increase');
          input.removeAttribute('readonly');
        });
      };

// Yearly category current inputs
document.querySelectorAll<HTMLInputElement>('.collection-category-current-input[data-year]:not([readonly])').forEach(input => {
  input.addEventListener('input', function() {
    const currentValue = getNumericValue(this);
    const oldValue = parseInt(this.getAttribute('data-old-value') || '0');

    this.setAttribute('data-old-value', currentValue.toString());

    const isIncreasing = currentValue > oldValue;

    if (isIncreasing) {
      const allowByOBL = allowInputIncrease(this, currentValue);
      const allowByYear = allowYearCategoryInputIncrease(this, currentValue);

      if (!allowByOBL || !allowByYear) {
        setNumericValue(this, oldValue);
        updateOBLValidationDisplay(this);
        return;
      }
    }

    if (currentValue === 0) {
      this.className = this.className.replace('collection-input-with-value', 'collection-input-empty');
    } else {
      this.className = this.className.replace('collection-input-empty', 'collection-input-with-value');
    }

    updateOBLValidationDisplay(this);
    validateYearlyCategoryTotals(this);
  });
});

// Helper function for yearly category validation
const allowYearCategoryInputIncrease = (input: HTMLInputElement, newValue: number): boolean => {
  const oldValue = parseInt(input.getAttribute('data-old-value') || '0');
  const isIncreasing = newValue > oldValue;

  if (!isIncreasing) {
    return true;
  }

  const year = input.dataset.year!;
  const yearlyCurrentInput = document.querySelector<HTMLInputElement>(`input.collection-yearly-current-input[data-year="${year}"]`);

  if (!yearlyCurrentInput) return true;

  const yearlyCurrentValue = getNumericValue(yearlyCurrentInput);
  const categoryInputs = document.querySelectorAll<HTMLInputElement>(`input.collection-category-current-input[data-year="${year}"]`);

  let currentCategoriesTotal = 0;
  categoryInputs.forEach(categoryInput => {
    if (categoryInput === input) {
      currentCategoriesTotal += newValue;
    } else {
      currentCategoriesTotal += getNumericValue(categoryInput);
    }
  });

  return currentCategoriesTotal <= yearlyCurrentValue;
};

const updateOBLValidationDisplay = (changedInput?: HTMLInputElement): void => {
  validateOBLLimit(changedInput);
};

const validateFormSubmission = (): boolean => {
  const oblValue = getNumericValue(oblLastMonthInput);
  const totalCurrent = calculateTotalCurrentValue();

  // OBL Validation - total cannot exceed OBL
  if (oblValue > 0 && totalCurrent > oblValue) {
    toast("error", `Total current value (${totalCurrent.toLocaleString('en-US')}) exceeds OBL (${oblValue.toLocaleString('en-US')})! Please reduce values.`);
    return false;
  }

  // Check if OBL is zero but there are values
  if (oblValue === 0 && totalCurrent > 0) {
    toast("error", `OBL is set to 0 but there are values (${totalCurrent.toLocaleString('en-US')}). Please set OBL value.`);
    return false;
  }

  // Month Category Validation - categories shouldn't exceed month total
  const invalidMonths: string[] = [];

  // Check all detailed months (indices 36-38)
  for (let i = 36; i <= 38 && i < allMonths.length; i++) {
    const monthKey = allMonths[i];
    const mainCurrentInput = document.querySelector<HTMLInputElement>(
      `input.collection-main-current-input[data-month="${monthKey}"]`
    );

    if (mainCurrentInput) {
      const mainCurrentValue = getNumericValue(mainCurrentInput);
      const categoryInputs = document.querySelectorAll<HTMLInputElement>(
        `input.collection-category-current-input[data-month="${monthKey}"]`
      );
      const categoriesTotal = Array.from(categoryInputs).reduce((sum, input) => sum + getNumericValue(input), 0);

      if (categoriesTotal > mainCurrentValue) {
        const monthDate = new Date(monthKey + "-01");
        invalidMonths.push(formatMonthLabel(monthDate));
      }
    }
  }

  if (invalidMonths.length > 0) {
    toast("error", `Categories exceed month totals in: ${invalidMonths.join(', ')}`);
    return false;
  }

  // Yearly Category Validation - categories shouldn't exceed year total
  const invalidYears: string[] = [];
  const yearlyInputs = document.querySelectorAll<HTMLInputElement>('.collection-yearly-current-input');

  yearlyInputs.forEach(input => {
    const year = input.dataset.year;
    if (year) {
      const yearlyCurrentValue = getNumericValue(input);
      const categoryInputs = document.querySelectorAll<HTMLInputElement>(
        `input.collection-category-current-input[data-year="${year}"]`
      );
      const categoriesTotal = Array.from(categoryInputs).reduce((sum, categoryInput) => sum + getNumericValue(categoryInput), 0);

      if (categoriesTotal > yearlyCurrentValue) {
        invalidYears.push(year);
      }
    }
  });

  if (invalidYears.length > 0) {
    toast("error", `Categories exceed year totals in: ${invalidYears.join(', ')}`);
    return false;
  }

  // Forecast Validation - categories forecast shouldn't exceed main forecast
  const currentMonthKey = allMonths[36];
  if (currentMonthKey) {
    const mainForecastInput = document.querySelector<HTMLInputElement>(
      `input.collection-main-forecast-input[data-month="${currentMonthKey}"]`
    );

    if (mainForecastInput) {
      const mainForecastValue = getNumericValue(mainForecastInput);
      let categoriesForecastTotal = 0;

      const categoryForecastInputs = document.querySelectorAll<HTMLInputElement>(
        `input.collection-category-forecast-input[data-month="${currentMonthKey}"]`
      );
      categoryForecastInputs.forEach(input => {
        categoriesForecastTotal += getNumericValue(input);
      });

      if (categoriesForecastTotal > mainForecastValue) {
        const monthDate = new Date(currentMonthKey + "-01");
        toast("error", `Category forecasts exceed main forecast for ${formatMonthLabel(monthDate)}`);
        return false;
      }
    }
  }

  return true;
};

      // Add this function to show/hide collection sections
      const toggleCollectionSections = (show: boolean): void => {
        const sections = [
          '.collection-obl-section-container',
          '#historicalCollectionSection',
          '#detailedCollectionSection',
          '#longtermCollectionSection',
          '#yearlyCollectionSection'
        ];

        sections.forEach(selector => {
          const element = document.querySelector(selector) as HTMLElement;
          if (element) {
            if (show) {
              element.classList.remove('collection-hidden-section');
              element.classList.add('collection-visible-section');
            } else {
              element.classList.remove('collection-visible-section');
              element.classList.add('collection-hidden-section');
            }
          }
        });
      };

const calculateOBLBreakdown = async (projectId: number): Promise<OBLBreakdown> => {
  const breakdown: OBLBreakdown = {
    normalDue: 0,
    normalOverdue: 0,
    retentionDue: 0,
    retentionOverdue: 0,
    gstTds: 0,
    wht: 0
  };

  try {
    const collection = collections.find(f => f.Project?.Id === projectId);
    if (!collection) {
      console.log(`⚠️ No collection found for project ${projectId}`);
      return breakdown;
    }

    console.log(`📊 Found collection for project:`, {
      OBL: collection.OBL,
      NormalDue: collection.NormalDue,
      NormalOverdue: collection.NormalOverdue,
      RetentionDue: collection.RetentionDue,
      RetentionOverdue: collection.RetentionOverdue,
      GSTTDS: collection.GSTTDS,
      WHT: collection.WHT
    });

    const result = {
      normalDue: Number(collection.NormalDue) || 0,
      normalOverdue: Number(collection.NormalOverdue) || 0,
      retentionDue: Number(collection.RetentionDue) || 0,
      retentionOverdue: Number(collection.RetentionOverdue) || 0,
      gstTds: Number(collection.GSTTDS) || 0,
      wht: Number(collection.WHT) || 0
    };

    console.log(`✅ Breakdown result:`, result);
    return result;

  } catch (error) {
    console.error("❌ Error in calculateOBLBreakdown:", error);
    return breakdown;
  }
};

 const renderOBLBreakdown = (breakdown: OBLBreakdown): void => {

  const setBreakdownValue = (elementId: string, value: number) => {
    const element = document.getElementById(elementId) as HTMLInputElement;
    if (element) {
      // Always show the actual value (0 if no data)
      setNumericValue(element, value);
    } else {
      console.error(`Element not found: ${elementId}`);
    }
  };

  setBreakdownValue('oblNormalDue', breakdown.normalDue);
  setBreakdownValue('oblNormalOverdue', breakdown.normalOverdue);
  setBreakdownValue('oblRetentionDue', breakdown.retentionDue);
  setBreakdownValue('oblRetentionOverdue', breakdown.retentionOverdue);
  setBreakdownValue('oblGstTds', breakdown.gstTds);
  setBreakdownValue('oblWht', breakdown.wht);
};

const onProjectChange = async (projectId: number, existingCollection?: CollectionItem): Promise<void> => {
  const project = projects.find(p => p.Id === projectId);
  if (!project) return;

  console.log(`🔄 Loading project: ${project.Title} (ID: ${projectId})`);

  // Show all collection sections
  toggleCollectionSections(true);
  renderProjectDetails(project);

  let collectionData: CollectionData = {};
  if (existingCollection) {
    collectionData = safeJsonParse(existingCollection.CollectionData, {});
  }

  const isNewProject = !existingCollection;

  const oblContainer = document.querySelector('.collection-obl-section-container') as HTMLElement;
  const fieldNote = oblContainer.querySelector('.field-note') as HTMLElement;

  // Calculate OBL breakdown FIRST
  const breakdown = await calculateOBLBreakdown(projectId);
  console.log(`📊 Breakdown for ${project.Title}:`, breakdown);

  // Render breakdown
  renderOBLBreakdown(breakdown);

  if (isNewProject) {
    oblContainer.className = 'collection-obl-section-container collection-visible-section new-project';
    oblLastMonthInput.setAttribute('readonly', 'true');
    oblLastMonthInput.placeholder = "Enter initial OBL for new project";
    if (fieldNote) fieldNote.textContent = "For new projects: Enter initial OBL value";
    oblValidationMessage.textContent = "Enter initial OBL value for new project";
    oblValidationMessage.className = 'collection-validation-message';

    const defaultOBL = project.PC ? 0 : 0;
    setNumericValue(oblLastMonthInput, defaultOBL);
  } else {
    oblContainer.className = 'collection-obl-section-container collection-visible-section existing-project';
    oblLastMonthInput.setAttribute('readonly', 'true');

    const oblValue = Number(existingCollection?.OBL) || 0;
    console.log(`💰 OBL Value for ${project.Title}: ${oblValue}`);

    setNumericValue(oblLastMonthInput, oblValue);

    oblLastMonthInput.placeholder = "Receivable value";
    if (fieldNote) fieldNote.textContent = "Receivable value record";

    if (oblValue > 0) {
      oblLastMonthInput.value = oblValue.toString();
      oblLastMonthInput.className = oblLastMonthInput.className.replace('collection-input-empty', 'collection-input-with-value');
    }
  }

  // Render collection interface
  renderCollectionInterface(collectionData, getNumericValue(oblLastMonthInput), existingCollection);

  // Initialize validation after rendering
  initializeFormValidation();

  // Force breakdown to render again after collection interface
  setTimeout(() => {
    renderOBLBreakdown(breakdown);
  }, 100);
};

// ========== HISTORICAL SECTION TOGGLE ==========
let historicalVisible: boolean = false;

const setupHistoricalToggle = (): void => {
  // Use setTimeout with 0 to ensure DOM is ready
  setTimeout(() => {
    const toggleBtn = document.getElementById('historicalToggleBtn') as HTMLButtonElement;
    const header = document.getElementById('historicalSectionHeader') as HTMLElement;
    const historicalContainer = document.getElementById('historicalContainer') as HTMLElement;

    if (!toggleBtn || !header || !historicalContainer) {
      console.warn('Historical toggle elements not found');
      return;
    }

    // Reset initial state - HIDDEN
    historicalVisible = false;
    historicalContainer.style.display = 'block';
    historicalContainer.style.opacity = '0';
    historicalContainer.style.maxHeight = '0';
    historicalContainer.style.padding = '0';
    historicalContainer.style.overflow = 'hidden';
    toggleBtn.textContent = '▼ Show';

    // Remove any existing event listeners by cloning
    const newToggleBtn = toggleBtn.cloneNode(true) as HTMLButtonElement;
    toggleBtn.parentNode?.replaceChild(newToggleBtn, toggleBtn);

    const performToggle = () => {
      historicalVisible = !historicalVisible;

      if (historicalVisible) {
        // SHOW - with animation
        historicalContainer.style.display = 'block';
        historicalContainer.style.padding = '5px 0';
        // Force reflow
        void historicalContainer.offsetHeight;
        historicalContainer.style.opacity = '1';
        historicalContainer.style.maxHeight = '2000px';
        historicalContainer.style.overflow = 'visible';
        newToggleBtn.textContent = '▲ Hide';
        newToggleBtn.style.backgroundColor = '#007bff';
        newToggleBtn.style.borderColor = '#007bff';
      } else {
        // HIDE - with animation
        historicalContainer.style.opacity = '0';
        historicalContainer.style.maxHeight = '0';
        historicalContainer.style.overflow = 'hidden';
        historicalContainer.style.padding = '0';
        setTimeout(() => {
          if (!historicalVisible) {
            historicalContainer.style.display = 'none';
          }
        }, 400);
        newToggleBtn.textContent = '▼ Show';
        newToggleBtn.style.backgroundColor = '#6c757d';
        newToggleBtn.style.borderColor = '#6c757d';
      }
    };

    // Click handler for toggle button
    // FIX 2: preventDefault() + the button now has type="button" in the template,
    // so clicking "▼ Show" can no longer submit the form / navigate to the list page.
    newToggleBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      performToggle();
    };

    // Click handler for header (excluding button)
    header.onclick = (e) => {
      e.preventDefault();
      if (newToggleBtn.contains(e.target as Node)) return;
      performToggle();
    };

    // Hover effects
    header.onmouseenter = () => { header.style.backgroundColor = '#f8f9fa'; };
    header.onmouseleave = () => { header.style.backgroundColor = 'transparent'; };

  }, 50);
};


const renderProjectDetails = (
  project: ProjectItem
): void => {
  if (!project) {
    console.warn(
      "renderProjectDetails: project is missing"
    );
    return;
  }

  const projectNameDisplay =
    document.querySelector(
      "#projectNameDisplay"
    ) as HTMLElement | null;

  const projectRef =
    document.querySelector(
      "#projectRef"
    ) as HTMLElement | null;

  const profitCenter =
    document.querySelector(
      "#profitCenter"
    ) as HTMLElement | null;

  const projectManager =
    document.querySelector(
      "#projectManager"
    ) as HTMLElement | null;

  const groupHead =
    document.querySelector(
      "#groupHead"
    ) as HTMLElement | null;

  const customer =
    document.querySelector(
      "#customer"
    ) as HTMLElement | null;

   // ✅ Use helper functions for Person/Group types
  const managerName = getProjectManagerDisplayName(project.ProjectManager);
  const groupHeadName = getGroupHeadDisplayName(project.GroupHead1);

  if (projectNameDisplay) projectNameDisplay.textContent = project.Title || "-";
  if (projectRef) projectRef.textContent = project.ProjectReference || "-";
  if (profitCenter) profitCenter.textContent = project.PC || "-";
  if (projectManager) projectManager.textContent = managerName || "-";
  if (groupHead) groupHead.textContent = groupHeadName || "-";
  if (customer) customer.textContent = project.CustomerName || "-";

  console.log(
    "✅ Project details rendered:",
    {
      project: project.Title,
    projectRef: project.ProjectReference,
    profitCenter: project.PC,
    projectManager: managerName,
    groupHead: groupHeadName,
    customer: project.CustomerName
    }
  );
};

const getCollectionMonths = (): string[] => {
  return generateIdealMonthStructure();
};


// ========== FIXED RENDER COLLECTION INTERFACE ==========

const renderCollectionInterface = (existingData: CollectionData, displayOBL: number, existingCollection?: CollectionItem): void => {
    allMonths = getCollectionMonths();

    setNumericValue(oblLastMonthInput, displayOBL);

    // CRITICAL FIX: If existingData is empty, create fresh structure
    if (!existingData || typeof existingData !== 'object' || Object.keys(existingData).length === 0) {
        console.log("🔄 No existing data found, creating fresh collection structure");
        existingData = createFreshCollectionStructure();
    }

    // Get the ideal month structure
    const idealMonths = generateIdealMonthStructure();

    const displayHistoricalMonths = idealMonths.slice(33, 36);

    // FIX: Ensure ALL months (historical, detailed, long-term) have monthActual
    idealMonths.forEach((monthKey) => {
        if (!existingData[monthKey]) {
            // Create new month data with monthActual
            existingData[monthKey] = {
                plan: 0,
                actual: 0,
                monthActual: 0,
                forecast: 0,
                date: '',
                items: {}
            };
        } else {
            const monthData = existingData[monthKey] as CollectionMonthData;
            // Ensure monthActual exists for ALL months
            if (monthData.monthActual === undefined) {
                monthData.monthActual = 0;
            }
            if (monthData.forecast === undefined) {
                monthData.forecast = 0;
            }
            if (monthData.date === undefined) {
                monthData.date = '';
            }
            if (!monthData.items) {
                monthData.items = {};
            }
        }
    });

    // Ensure detailed months (36-38) have proper structure
    for (let i = 36; i <= 38 && i < idealMonths.length; i++) {
        const monthKey = idealMonths[i];
        if (existingData[monthKey]) {
            const monthData = existingData[monthKey] as CollectionMonthData;
            if (!monthData.items || Object.keys(monthData.items).length === 0) {
                monthData.items = createEmptyCategories();
            } else {
                const categories = ['Normal', 'Normal Overdue', 'Retention', 'Retention Overdue'];
                categories.forEach(cat => {
                    if (!monthData.items![cat]) {
                        monthData.items![cat] = {
                            name: cat,
                            plan: 0,
                            actual: 0,
                            forecast: 0,
                            date: ''
                        };
                    }
                    if (monthData.items![cat].forecast === undefined) {
                        monthData.items![cat].forecast = 0;
                    }
                    if (monthData.items![cat].date === undefined) {
                        monthData.items![cat].date = '';
                    }
                });
            }
        }
    }

    // Ensure long-term months (39+) have proper structure
    for (let i = 39; i < idealMonths.length; i++) {
        const monthKey = idealMonths[i];
        if (existingData[monthKey]) {
            const monthData = existingData[monthKey] as CollectionMonthData;
            if (!monthData.items || Object.keys(monthData.items).length === 0) {
                monthData.items = createSimpleCategories();
            } else {
                const categories = ['Normal', 'Retention'];
                categories.forEach(cat => {
                    if (!monthData.items![cat]) {
                        monthData.items![cat] = {
                            name: cat,
                            plan: 0,
                            actual: 0
                        };
                    }
                });
            }
        }
    }

    // Ensure yearly data exists
    if (!existingData.yearly || Object.keys(existingData.yearly).length === 0) {
        existingData.yearly = createFreshYearlyStructure(idealMonths[idealMonths.length - 1]);
    }

    // ALWAYS sync monthActual from RevenueActual so "Actual" = RevenueActual
    if (existingCollection) {
        const revActual = (existingCollection as any)._revenueActual || {};
        idealMonths.forEach(monthKey => {
            const monthData = existingData[monthKey] as CollectionMonthData;
            if (monthData && revActual[monthKey] !== undefined) {
                monthData.monthActual = Number(revActual[monthKey]) || 0;
            }
        });
    }

    // Render only the LAST 3 months in the UI
    renderHistoricalMonths(displayHistoricalMonths, existingData);

    // setTimeout(() => {
    //     setupHistoricalToggle();
    // }, 50);

    // Pass existingCollection to get RevenueActual for current month
    renderFirstRow(idealMonths.slice(36, 39), existingData);

    renderRemainingMonths(idealMonths.slice(39), existingData);

    renderYearlyCollection(existingData.yearly || createFreshYearlyStructure(idealMonths[idealMonths.length - 1]));

    protectHistoricalElements();

    // Calculate totals for all months first
    calculateMonthTotals();

    // Then set all sections readonly/editable
    setAllSectionsReadOnly();

    updateOBLValidationDisplay();
    attachCollectionEventListeners();

    // Initialize validation after rendering
    initializeFormValidation();

    // Initialize forecast validation
    setupForecastValidation();

  // Apply freeze state
if (existingCollection) {
    const isFrozen = existingCollection.Freeze === true;
    setTimeout(() => {
        setFormReadOnly(isFrozen);
        // FIX 1: use the shared helper so the project name is ALWAYS included in the
        // form title (previously this got overwritten 50ms later by editCollection
        // with plain "Edit Collection", making the project name flash and disappear).
        const titleProject = projects.find(p => p.Id === parseInt(projectSelect.value));
        updateEditFormHeader(isFrozen, titleProject ? authCanEditProject(titleProject) : false);
    }, 250);
}
};
const protectHistoricalElements = (): void => {
  // Remove any existing historical elements and replace with readonly versions
  document.querySelectorAll('.collection-historical-month').forEach(monthElement => {
    // Clone the element to remove event listeners
    const newElement = monthElement.cloneNode(true) as HTMLElement;
    monthElement.parentNode?.replaceChild(newElement, monthElement);

    // Add readonly attributes to all inputs and ensure they are properly styled
    newElement.querySelectorAll('input').forEach(input => {
      input.setAttribute('readonly', 'true');
      input.style.backgroundColor = '#f8f9fa';
      input.style.cursor = 'not-allowed';
      input.style.pointerEvents = 'none';

      // Ensure proper classes are maintained
      if (input.classList.contains('collection-main-current-input')) {
        input.className = input.className.replace('collection-input-empty', 'collection-input-with-value');
      }
      if (input.classList.contains('collection-category-current-input')) {
        // Keep the class as is - it will show input-empty or input-with-value based on data
      }
    });
  });

  // Add global event prevention
  document.querySelectorAll('.collection-historical-month').forEach(section => {
    section.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      return false;
    });

    section.addEventListener('keydown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      return false;
    });
  });
};

// Force validation on form load
const initializeFormValidation = (): void => {
  setTimeout(() => {
    // Validate all detailed months
    const detailedMonths = allMonths.slice(3, 6);
    detailedMonths.forEach(monthKey => {
      validateSingleMonthCategories(monthKey);
    });

    // Update OBL display
    updateOBLValidationDisplay();

  }, 500);
};


const renderHistoricalMonths = (months: string[], existingData: CollectionData): void => {
  let html = '';

  // Sort months chronologically (oldest to newest)
  const sortedMonths = [...months].sort();

  if (sortedMonths.length === 0) {
    html = `<div style="padding: 20px; text-align: center; color: #6c757d;">No historical data available</div>`;
    historicalContainer.innerHTML = html;
    // Keep container hidden by default
    historicalContainer.style.display = 'none';
    historicalContainer.style.opacity = '0';
    historicalContainer.style.maxHeight = '0';
    historicalContainer.style.padding = '0';
    historicalContainer.style.overflow = 'hidden';
    return;
  }

  // Create a grid layout for historical months
  html = `<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 15px;">`;

  sortedMonths.forEach((monthKey, index) => {
    const monthDate = new Date(monthKey + "-01");
    const monthLabel = formatMonthLabel(monthDate);
    const data = existingData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, items: {} };

    html += renderHistoricalMonthBox(monthKey, monthLabel, data, index);
  });

  html += `</div>`;

  historicalContainer.innerHTML = html;

  // ✅ FIX: Make container visible but collapsed
  historicalContainer.style.display = 'block';
  historicalContainer.style.opacity = '0';
  historicalContainer.style.maxHeight = '0';
  historicalContainer.style.padding = '0';
  historicalContainer.style.overflow = 'hidden';

  // Setup toggle functionality after rendering
  setupHistoricalToggle();

  // Protect historical elements after rendering
  protectHistoricalElements();
};

const renderHistoricalMonthBox = (monthKey: string, monthLabel: string, data: CollectionMonthData, index: number): string => {
    // Get values
    const planValue = data.plan || 0;
    const actualValue = data.actual || 0;
    const monthActualValue = data.monthActual || 0;

    const displayPlanValue = planValue > 0 ? planValue.toString() : '';
    const displayActualValue = actualValue > 0 ? actualValue.toString() : '';
    const displayMonthActual = monthActualValue > 0 ? monthActualValue.toString() : '';

    const planInputClass = planValue > 0 ? 'collection-input-with-value' : 'collection-input-empty';
    const actualInputClass = actualValue > 0 ? 'collection-input-with-value' : 'collection-input-empty';
    const monthActualClass = monthActualValue > 0 ? 'collection-input-with-value' : 'collection-input-empty';

    const categories = data.items || createEmptyCategories();

    // Calculate category values - ensure they match the main total
    const normalPlan = (categories.Normal?.plan || 0);
    const normalActual = (categories.Normal?.actual || 0);
    const normalOverduePlan = (categories['Normal Overdue']?.plan || 0);
    const normalOverdueActual = (categories['Normal Overdue']?.actual || 0);
    const retentionPlan = (categories.Retention?.plan || 0);
    const retentionActual = (categories.Retention?.actual || 0);
    const retentionOverduePlan = (categories['Retention Overdue']?.plan || 0);
    const retentionOverdueActual = (categories['Retention Overdue']?.actual || 0);

    // Combined totals (should equal main values)
    const totalNormalPlan = normalPlan + normalOverduePlan;
    const totalNormalActual = normalActual + normalOverdueActual;
    const totalRetentionPlan = retentionPlan + retentionOverduePlan;
    const totalRetentionActual = retentionActual + retentionOverdueActual;
    const categoriesTotalPlan = totalNormalPlan + totalRetentionPlan;
    const categoriesTotalActual = totalNormalActual + totalRetentionActual;

    return `
        <div class="collection-month-box collection-historical-month" style="border: 1px solid #dee2e6; border-radius: 8px; padding: 12px; margin-bottom: 10px; background: #f8f9fa;">
            <div class="collection-month-header" style="font-weight: bold; font-size: 14px; color: #495057; margin-bottom: 10px; border-bottom: 2px solid #007bff; padding-bottom: 5px;">
                ${monthLabel}
                <span style="font-size: 11px; font-weight: normal; color: #6c757d; margin-left: 8px;">(Historical - Readonly)</span>
            </div>

            <!-- Main Inputs -->
            <div class="collection-month-block" style="margin-bottom: 10px;">
                <div class="collection-main-inputs-container" style="display: flex; gap: 15px; flex-wrap: wrap;">
                    <div class="collection-main-input-group" style="flex: 1; min-width: 100px;">
                        <label style="font-size: 11px; color: #6c757d; display: block; margin-bottom: 2px;">Previous</label>
                        <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                              readonly class="collection-main-plan-input ${planInputClass} collection-number-input"
                              style="width: 100%; padding: 6px 8px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 14px; text-align: right; cursor: not-allowed;" />
                    </div>
                    <div class="collection-main-input-group" style="flex: 1; min-width: 100px;">
                        <label style="font-size: 11px; color: #6c757d; display: block; margin-bottom: 2px;">Current</label>
                        <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                              readonly class="collection-main-current-input ${actualInputClass} collection-number-input"
                              style="width: 100%; padding: 6px 8px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 14px; text-align: right; cursor: not-allowed;" />
                    </div>
                    <div class="collection-main-input-group" style="flex: 1; min-width: 100px;">
                        <label style="font-size: 11px; color: #6c757d; display: block; margin-bottom: 2px;">Actual</label>
                        <input type="text" value="${displayMonthActual}" data-type="month-actual" data-month="${monthKey}"
                              readonly class="collection-month-actual-input ${monthActualClass} collection-number-input"
                              style="width: 100%; padding: 6px 8px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 14px; text-align: right; font-weight: bold; color: #1a8a34; cursor: not-allowed;" />
                    </div>
                </div>
            </div>

            <!-- Categories Section -->
            <div class="collection-categories-section" style="border-top: 1px solid #e9ecef; padding-top: 10px;">
                <div class="collection-categories-header" style="display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 5px; font-size: 11px; font-weight: bold; color: #495057; padding: 5px 0; border-bottom: 1px solid #dee2e6; margin-bottom: 5px;">
                    <div>Category</div>
                    <div style="text-align: right;">Previous</div>
                    <div style="text-align: right;">Current</div>
                </div>

                <!-- Normal Category -->
                <div class="collection-category-row" style="display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Normal" data-category="Normal" data-month="${monthKey}"
                          class="collection-input-with-value" readonly
                          style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${normalPlan > 0 ? normalPlan.toString() : ''}"
                          data-category-type="plan" data-category="Normal" data-month="${monthKey}"
                          readonly class="collection-category-plan-input ${normalPlan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${normalActual > 0 ? normalActual.toString() : ''}"
                          data-category-type="actual" data-category="Normal" data-month="${monthKey}"
                          readonly class="collection-category-current-input ${normalActual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>

                <!-- Normal Overdue Category -->
                <div class="collection-category-row" style="display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Normal Overdue" data-category="Normal Overdue" data-month="${monthKey}"
                          class="collection-input-with-value" readonly
                          style="border: none; background: transparent; font-weight: 500; color: #856404; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${normalOverduePlan > 0 ? normalOverduePlan.toString() : ''}"
                          data-category-type="plan" data-category="Normal Overdue" data-month="${monthKey}"
                          readonly class="collection-category-plan-input ${normalOverduePlan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${normalOverdueActual > 0 ? normalOverdueActual.toString() : ''}"
                          data-category-type="actual" data-category="Normal Overdue" data-month="${monthKey}"
                          readonly class="collection-category-current-input ${normalOverdueActual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>

                <!-- Retention Category -->
                <div class="collection-category-row" style="display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Retention" data-category="Retention" data-month="${monthKey}"
                          class="collection-input-with-value" readonly
                          style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${retentionPlan > 0 ? retentionPlan.toString() : ''}"
                          data-category-type="plan" data-category="Retention" data-month="${monthKey}"
                          readonly class="collection-category-plan-input ${retentionPlan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${retentionActual > 0 ? retentionActual.toString() : ''}"
                          data-category-type="actual" data-category="Retention" data-month="${monthKey}"
                          readonly class="collection-category-current-input ${retentionActual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>

                <!-- Retention Overdue Category -->
                <div class="collection-category-row" style="display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Retention Overdue" data-category="Retention Overdue" data-month="${monthKey}"
                          class="collection-input-with-value" readonly
                          style="border: none; background: transparent; font-weight: 500; color: #856404; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${retentionOverduePlan > 0 ? retentionOverduePlan.toString() : ''}"
                          data-category-type="plan" data-category="Retention Overdue" data-month="${monthKey}"
                          readonly class="collection-category-plan-input ${retentionOverduePlan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${retentionOverdueActual > 0 ? retentionOverdueActual.toString() : ''}"
                          data-category-type="actual" data-category="Retention Overdue" data-month="${monthKey}"
                          readonly class="collection-category-current-input ${retentionOverdueActual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
            </div>

            <!-- Summary Display - Shows totals match -->
            <div style="margin-top: 10px; padding: 8px 12px; background: ${categoriesTotalActual === actualValue ? '#d4edda' : '#f8d7da'}; border-radius: 4px; font-size: 12px; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 5px; border: 1px solid ${categoriesTotalActual === actualValue ? '#c3e6cb' : '#f5c6cb'};">
                <span style="font-weight: 600; color: #495057;">📊 Summary:</span>
                <span style="color: #0056b3;">Normal: ${totalNormalPlan} → ${totalNormalActual}</span>
                <span style="color: #28a745;">Retention: ${totalRetentionPlan} → ${totalRetentionActual}</span>
                <span style="font-weight: 600; color: ${categoriesTotalActual === actualValue ? '#155724' : '#721c24'};">
                    Total: ${categoriesTotalPlan} → ${categoriesTotalActual}
                    ${categoriesTotalActual === actualValue ? '✅' : '⚠️ ' + (categoriesTotalActual - actualValue)}
                </span>
                <span style="color: #1a8a34; font-weight: 600;">
                    📊 Actual: ${monthActualValue}
                </span>
            </div>
        </div>
    `;
};

const renderFirstRow = (months: string[], existingData: CollectionData): void => {
  let html = '';

  // Create single row container for all three months
  html += `<div class="collection-detailed-months-row">`;

  // Current Month (first of the detailed months)
  const currentMonth = months[0];
  if (currentMonth) {
    const monthDate = new Date(currentMonth + "-01");
    const monthLabel = formatMonthLabel(monthDate);
    const data = existingData[currentMonth] as CollectionMonthData || { plan: 0, actual: 0, items: createEmptyCategories() };

    html += `<div class="collection-current-month-container">`;
    html += renderDetailedMonthBox(currentMonth, monthLabel, data, true);
    html += `</div>`;
  }

  // Next 2 Months
  html += `<div class="collection-next-months-container">`;
  for (let i = 1; i <= 2; i++) {
    if (i < months.length) {
      const monthKey = months[i];
      const monthDate = new Date(monthKey + "-01");
      const monthLabel = formatMonthLabel(monthDate);
      const data = existingData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, items: createEmptyCategories() };

      html += `<div class="collection-next-month-box">`;
      html += renderDetailedMonthBox(monthKey, monthLabel, data, false);
      html += `</div>`;
    }
  }
  html += `</div>`;

  html += `</div>`;

  firstRowContainer.innerHTML = html;
};

const renderRemainingMonths = (months: string[], existingData: CollectionData): void => {
        let html = '';

        for (let row = 0; row < 4; row++) {
    html += `<div class="collection-remaining-row">`;
    for (let col = 0; col < 4; col++) {
      const index = (row * 4) + col;
      if (index >= months.length) break;

      const monthKey = months[index];
      const monthDate = new Date(monthKey + "-01");
      const monthLabel = formatMonthLabel(monthDate);
      const data = existingData[monthKey] as CollectionMonthData || { plan: 0, actual: 0, items: createSimpleCategories() };

      html += renderSimpleMonthBox(monthKey, monthLabel, data);
    }
    html += `</div>`;
  }

        remainingContainer.innerHTML = html;
      };

      const renderYearlyCollection = (yearlyData: YearlyCollectionData): void => {
        let html = '';

        const years = Object.keys(yearlyData).sort();

        years.forEach(year => {
          const data = yearlyData[year] || { plan: 0, actual: 0 };
          html += renderYearlyBox(year, data);
        });

        yearlyContainer.innerHTML = html;
      };

const renderYearlyBox = (year: string, data: { plan: number; actual: number, items?: { [key: string]: CollectionCategory } }): string => {
  const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
  const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
  const planInputClass = data.plan > 0 ? 'collection-input-with-value' : 'collection-input-empty';
  const actualInputClass = data.actual > 0 ? 'collection-input-with-value' : 'collection-input-empty';

  const categories = data.items || {
    "Normal": { name: 'Normal', plan: 0, actual: 0 },
    "Retention": { name: 'Retention', plan: 0, actual: 0 }
  };

  return `
    <div class="collection-yearly-box">
      <div class="collection-yearly-header">${year}</div>

      <div class="collection-yearly-inputs-container">
        <div class="collection-yearly-input-group">
          <label>Previous</label>
          <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                readonly class="collection-yearly-plan-input ${planInputClass} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
        <div class="collection-yearly-input-group">
          <label>Current</label>
          <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                readonly class="collection-yearly-current-input ${actualInputClass} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
      </div>

      <div class="collection-categories-section">
        <div class="collection-categories-header">
          <div>Category</div>
          <div>Previous</div>
          <div>Current</div>
        </div>

        <!-- Normal Category - READONLY -->
        <div class="collection-category-row">
          <input type="text" value="Normal" data-category="Normal" data-year="${year}"
                class="collection-input-with-value" readonly
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${categories.Normal?.plan > 0 ? categories.Normal.plan.toString() : ''}"
                data-category-type="plan" data-category="Normal" data-year="${year}"
                readonly class="collection-category-plan-input ${categories.Normal?.plan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${categories.Normal?.actual > 0 ? categories.Normal.actual.toString() : ''}"
                data-category-type="actual" data-category="Normal" data-year="${year}"
                readonly class="collection-category-current-input ${categories.Normal?.actual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>

        <!-- Retention Category - READONLY -->
        <div class="collection-category-row">
          <input type="text" value="Retention" data-category="Retention" data-year="${year}"
                class="collection-input-with-value" readonly
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${categories.Retention?.plan > 0 ? categories.Retention.plan.toString() : ''}"
                data-category-type="plan" data-category="Retention" data-year="${year}"
                readonly class="collection-category-plan-input ${categories.Retention?.plan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${categories.Retention?.actual > 0 ? categories.Retention.actual.toString() : ''}"
                data-category-type="actual" data-category="Retention" data-year="${year}"
                readonly class="collection-category-current-input ${categories.Retention?.actual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>

        <!-- Validation Display - Simplified -->
        <div class="collection-yearly-total-validation" data-year="${year}" style="margin-top: 10px; font-size: 11px; text-align: center;"></div>
      </div>
    </div>
  `;
};

// ========== FIXED RENDER DETAILED MONTH BOX ==========

const renderDetailedMonthBox = (monthKey: string, monthLabel: string, data: CollectionMonthData, isCurrentMonth: boolean = false): string => {
    // Get values from data
    const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
    const displayCurrentValue = data.actual > 0 ? data.actual.toString() : '';

    // Get monthActual from data (stored in CollectionData)
    let monthActualValue = data.monthActual || 0;

    // Get forecast/further value
    const forecastValue = data.forecast || 0;
    const displayForecastValue = forecastValue > 0 ? forecastValue.toString() : '';

    // If monthActual is not in data and this is current month, try to get from existingCollection
    if (data.monthActual === undefined && isCurrentMonth) {
        try {
            const projectId = parseInt(projectSelect.value);
            const collection = collections.find(f => f.Project?.Id === projectId);
            if (collection && collection._revenueActual) {
                monthActualValue = Number(collection._revenueActual[monthKey]) || 0;
            }
        } catch (e) {
            monthActualValue = 0;
        }
    }

    // Forecast = RevenueActual + Further
    const forecastTotal = monthActualValue + forecastValue;
    const displayForecastTotal = forecastTotal > 0 ? forecastTotal.toString() : '';

    // CSS classes
    const planInputClass = data.plan > 0 ? 'collection-input-with-value' : 'collection-input-empty';
    const currentInputClass = data.actual > 0 ? 'collection-input-with-value' : 'collection-input-empty';
    const forecastInputClass = forecastValue > 0 ? 'collection-input-with-value' : 'collection-input-empty';
    const forecastTotalClass = forecastTotal > 0 ? 'collection-input-with-value' : 'collection-input-empty';

    // Get categories with proper forecast values
    const categories = data.items || {};

    // For detailed months (Current + Next 2 Months) - Show ALL 4 categories
    const categoryKeys = ['Normal', 'Normal Overdue', 'Retention', 'Retention Overdue'];

    const categoryRows = categoryKeys.map(categoryKey => {
        const category = categories[categoryKey] || {
            name: categoryKey,
            plan: 0,
            actual: 0,
            forecast: 0,
            date: ''
        };
        const isOverdue = categoryKey.includes('Overdue');
        const color = isOverdue ? '#856404' : '#495057';

        const categoryForecast = category.forecast || 0;
        const displayCategoryForecast = categoryForecast > 0 ? categoryForecast.toString() : '';
        const categoryForecastClass = categoryForecast > 0 ? 'collection-input-with-value' : 'collection-input-empty';
        const categoryDateClass = category.date ? 'collection-input-with-value' : 'collection-input-empty';

        return `
            <tr style="border-bottom: 1px solid #dee2e6;">
                <td style="padding: 6px 8px; font-weight: 500; color: ${color};">${categoryKey}</td>
                <td style="padding: 6px 8px; text-align: right;">
                    <input type="text" value="${(category.plan || 0) > 0 ? category.plan.toString() : ''}"
                          data-category-type="plan" data-category="${categoryKey}" data-month="${monthKey}"
                          readonly class="collection-category-plan-input ${(category.plan || 0) > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 80px; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #e9ecef; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </td>
                <td style="padding: 6px 8px; text-align: right;">
                    <input type="text" value="${(category.actual || 0) > 0 ? category.actual.toString() : ''}"
                          data-category-type="actual" data-category="${categoryKey}" data-month="${monthKey}"
                          class="collection-category-current-input ${(category.actual || 0) > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                          style="width: 80px; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #ffffff; font-size: 13px; text-align: right; cursor: text;" />
                </td>
                ${isCurrentMonth ? `
                    <td style="padding: 6px 8px; text-align: right;">
                        <input type="text" value="${displayCategoryForecast}"
                              data-category-type="forecast" data-category="${categoryKey}" data-month="${monthKey}"
                              class="collection-category-forecast-input ${categoryForecastClass} collection-number-input"
                              style="width: 80px; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #ffffff; font-size: 13px; text-align: right; cursor: text;"
                              placeholder="0" />
                    </td>
                    <td style="padding: 6px 8px; text-align: center;">
                        <input type="date" value="${category.date || ''}"
                              data-category-type="date" data-category="${categoryKey}" data-month="${monthKey}"
                              class="collection-category-date-input ${categoryDateClass}"
                              style="width: 100px; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #ffffff; font-size: 13px;color:black !important; text-align: center; cursor: text;"
                              placeholder="Date" />
                    </td>
                ` : ''}
            </tr>
        `; 
    }).join('');

    // ✅ Month Actual badge (only for the current month)
    const hasMonthActual = monthActualValue > 0;
        const monthActualBadge = isCurrentMonth ? `
        <span class="collection-month-actual-badge ${hasMonthActual ? 'collection-has-value' : ''}" style="
            background-color: ${hasMonthActual ? '#c8e6c9' : '#f5f5f5'};
            border: 1px solid ${hasMonthActual ? '#4caf50' : '#ddd'};
            border-radius: 4px;
            padding: 2px 12px;
            font-size: 13px;
            font-weight: 600;
            color: ${hasMonthActual ? '#1b5e20' : '#999'};
            display: inline-flex;
            align-items: center;
            gap: 6px;
            white-space: nowrap;
            height: 19px;
            line-height: 1;
            box-sizing: border-box;
        ">
            <span>📊 Month Actual:</span>
            <span style="font-size: 16px; font-weight: 700; line-height: 1;">${monthActualValue.toLocaleString('en-US')}</span>
        </span>
    ` : '';

    // ✅ NEW: Date badge (only for current month) - mirrors revenue.ts
    // Picks the HIGHEST (max) date among this month's category item dates.
    // Falls back to today's system date when no item date is set.
    let headerDateValue = '';
    if (isCurrentMonth) {
        const catDates: string[] = [];
        categoryKeys.forEach(categoryKey => {
            const c = categories[categoryKey];
            if (c && c.date) catDates.push(c.date);
        });

        if (catDates.length > 0) {
            headerDateValue = catDates.sort().slice(-1)[0];  // max date
        } else {
            const today = new Date();
            const yyyy = today.getFullYear();
            const mm = String(today.getMonth() + 1).padStart(2, '0');
            const dd = String(today.getDate()).padStart(2, '0');
            headerDateValue = `${yyyy}-${mm}-${dd}`;
        }
    }

    const hasHeaderDate = !!headerDateValue;
        const dateBadge = isCurrentMonth ? `
        <span class="collection-month-date-badge ${hasHeaderDate ? 'collection-has-value' : ''}" style="
            background-color: ${hasHeaderDate ? '#e3f2fd' : '#f5f5f5'};
            border: 1px solid ${hasHeaderDate ? '#007bff' : '#ddd'};
            border-radius: 4px;
            padding: 2px 12px;
            font-size: 13px;
            font-weight: 600;
            color: ${hasHeaderDate ? '#0c5460' : '#999'};
            display: inline-flex;
            align-items: center;
            gap: 6px;
            white-space: nowrap;
            height: 28px;
            line-height: 1;
            box-sizing: border-box;
        ">
            <span>📅</span>
            <span>${hasHeaderDate ? headerDateValue : ''}</span>
        </span>
    ` : '';

    return `
         <div class="collection-detailed-month-box ${isCurrentMonth ? 'collection-current-month-box' : 'collection-next-month-box'}" style="border: 1px solid #dee2e6; border-radius: 8px; padding: 12px; margin-bottom: 15px;">
            <div class="collection-month-header" style="display: flex; justify-content: space-between; align-items: center; font-weight: bold; font-size: 14px; color: #495057; margin-bottom: 10px; border-bottom: 2px solid #007bff; padding-bottom: 5px; gap: 10px; flex-wrap: nowrap; white-space: nowrap;">
                <span style="flex-shrink: 0;">${monthLabel} ${isCurrentMonth ? '<span style="background: #007bff; color: white; padding: 2px 8px; border-radius: 4px; font-size: 10px; margin-left: 8px; vertical-align: middle;">Current</span>' : ''}</span>
                <span style="display: inline-flex; align-items: center; gap: 10px; margin-left: auto; flex-shrink: 0;">
                    ${monthActualBadge}
                    ${dateBadge}
                </span>
            </div>

            <!-- Top Row: Previous, Current, Further, Forecast, Date (Actual removed) -->
            <div style="display: grid; grid-template-columns: ${isCurrentMonth ? '1fr 1fr 1fr 1fr 1fr' : '1fr 1fr'}; gap: 8px; margin-bottom: 12px; background: #f8f9fa; padding: 8px; border-radius: 4px;">
                <div>
                    <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Previous</label>
                    <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                          readonly class="collection-main-plan-input ${planInputClass} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
                <div>
                    <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Current</label>
                    <input type="text" value="${displayCurrentValue}" data-type="actual" data-month="${monthKey}"
                          readonly class="collection-main-current-input ${currentInputClass} collection-number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
                ${isCurrentMonth ? `
                    <div>
                        <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Further</label>
                        <input type="text" value="${displayForecastValue}" data-type="forecast" data-month="${monthKey}"
                              class="collection-main-forecast-input ${forecastInputClass} collection-number-input"
                              style="width: 100%; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #ffffff; font-size: 13px; text-align: right; cursor: text;" />
                    </div>
                    <div>
                        <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Forecast</label>
                        <input type="text" value="${displayForecastTotal}" data-type="forecast-total" data-month="${monthKey}"
                              readonly class="collection-forecast-total-input ${forecastTotalClass} collection-number-input"
                              style="width: 100%; padding: 4px 6px; border: 1px solid #007bff; border-radius: 4px; background: #f0f8ff; font-size: 13px; text-align: right; font-weight: bold; color: #0c5460;" />
                    </div>
                    <div>
                        <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Date</label>
                        <input type="date" value="${data.date || ''}" data-type="date" data-month="${monthKey}"
                              readonly class="collection-forecast-date-input collection-number-input"
                              style="width: 100%; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #f5f5f5; font-size: 13px; color: #000; text-align: center; cursor: default;" />
                    </div>
                ` : ''}
            </div>

            <!-- Category Table -->
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                    <thead>
                        <tr style="background: #f8f9fa; border-bottom: 2px solid #dee2e6;">
                            <th style="padding: 6px 8px; text-align: left; font-weight: bold; color: #495057;">Category</th>
                            <th style="padding: 6px 8px; text-align: right; font-weight: bold; color: #495057;">Previous</th>
                            <th style="padding: 6px 8px; text-align: right; font-weight: bold; color: #495057;">Current</th>
                            ${isCurrentMonth ? `
                                <th style="padding: 6px 8px; text-align: right; font-weight: bold; color: #495057;">Further</th>
                                <th style="padding: 6px 8px; text-align: center; font-weight: bold; color: #495057;">Date</th>
                            ` : ''}
                        </tr>
                    </thead>
                    <tbody>
                        ${categoryRows}
                    </tbody>
                </table>
            </div>

            <!-- Forecast validation display -->
            ${isCurrentMonth ? `
                <div class="collection-forecast-validation" data-month="${monthKey}" style="margin-top: 10px; font-size: 11px; text-align: center; color: #6c757d;">
                    <span style="display: block; padding: 4px 8px; background: #f8f9fa; border-radius: 4px; border: 1px solid #dee2e6;">
                        💡 Forecast = Actual + Further
                    </span>
                </div>
            ` : ''}
        </div>
    `;
};

// Sync the main (summary) Date input from the HIGHEST category item date
const syncMainDateFromItems = (monthKey: string): void => {
    const categoryDateInputs = document.querySelectorAll<HTMLInputElement>(
        `input.collection-category-date-input[data-month="${monthKey}"]`
    );
    const mainDateInput = document.querySelector<HTMLInputElement>(
        `input.collection-forecast-date-input[data-month="${monthKey}"]`
    );
    if (!mainDateInput) return;

    let maxDate = '';
    categoryDateInputs.forEach(inp => {
        if (inp.value && (!maxDate || inp.value > maxDate)) {
            maxDate = inp.value;
        }
    });

    // If no item date is set, default to today's system date
    if (!maxDate) {
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        maxDate = `${yyyy}-${mm}-${dd}`;
    }

    mainDateInput.value = maxDate;
    if (maxDate) {
        mainDateInput.classList.remove('collection-input-empty');
        mainDateInput.classList.add('collection-input-with-value');
    } else {
        mainDateInput.classList.remove('collection-input-with-value');
        mainDateInput.classList.add('collection-input-empty');
    }

    // Also refresh the header Date badge in real time
    const box = mainDateInput.closest('.collection-detailed-month-box') as HTMLElement | null;
    const badge = box?.querySelector('.collection-month-date-badge') as HTMLElement | null;
    if (badge) {
        const dateSpan = badge.querySelectorAll('span')[1] as HTMLElement | null;
        if (dateSpan) dateSpan.textContent = maxDate || '-';
        if (maxDate) {
            badge.classList.add('collection-has-value');
            badge.style.backgroundColor = '#e3f2fd';
            badge.style.borderColor = '#007bff';
            badge.style.color = '#0c5460';
        }
    }
};
      

const renderSimpleMonthBox = (monthKey: string, monthLabel: string, data: CollectionMonthData): string => {
  const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
  const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
  const planInputClass = data.plan > 0 ? 'collection-input-with-value' : 'collection-input-empty';
  const actualInputClass = data.actual > 0 ? 'collection-input-with-value' : 'collection-input-empty';

  const categories = data.items || createSimpleCategories();

  return `
    <div class="collection-simple-month">
      <div class="collection-month-header">${monthLabel}</div>

      <div class="collection-main-inputs-container">
        <div class="collection-main-input-group">
          <label>Prev. Plan</label>
          <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                readonly class="collection-main-plan-input ${planInputClass} collection-number-input"
                style="background: #e9ecef; border: 1px solid #e1e5e8; border-radius: 4px; cursor: not-allowed;" />
        </div>
        <div class="collection-main-input-group">
          <label>Curr. Plan</label>
          <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                readonly class="collection-main-current-input ${actualInputClass} collection-number-input"
                style="background: #e9ecef; border: 1px solid #e1e5e8; border-radius: 4px; cursor: not-allowed;" />
        </div>
      </div>

      <div class="collection-categories-section">
        <div class="collection-categories-header">
          <div>Category</div>
          <div>Prev. Plan</div>
          <div>Curr. Plan</div>
        </div>

        <!-- Normal Category -->
        <div class="collection-category-row">
          <input type="text" value="Normal" data-category="Normal" data-month="${monthKey}"
                class="collection-input-with-value" readonly
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${categories.Normal?.plan > 0 ? categories.Normal.plan.toString() : ''}"
                data-category-type="plan" data-category="Normal" data-month="${monthKey}"
                readonly class="collection-category-plan-input ${categories.Normal?.plan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${categories.Normal?.actual > 0 ? categories.Normal.actual.toString() : ''}"
                data-category-type="actual" data-category="Normal" data-month="${monthKey}"
                readonly class="collection-category-current-input ${categories.Normal?.actual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>

        <!-- Retention Category -->
        <div class="collection-category-row">
          <input type="text" value="Retention" data-category="Retention" data-month="${monthKey}"
                class="collection-input-with-value" readonly
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${categories.Retention?.plan > 0 ? categories.Retention.plan.toString() : ''}"
                data-category-type="plan" data-category="Retention" data-month="${monthKey}"
                readonly class="collection-category-plan-input ${categories.Retention?.plan > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${categories.Retention?.actual > 0 ? categories.Retention.actual.toString() : ''}"
                data-category-type="actual" data-category="Retention" data-month="${monthKey}"
                readonly class="collection-category-current-input ${categories.Retention?.actual > 0 ? 'collection-input-with-value' : 'collection-input-empty'} collection-number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
      </div>

      <!-- Validation display area -->
    </div>
  `;
};

// ========== FIXED SET ALL SECTIONS READONLY ==========

const setAllSectionsReadOnly = (): void => {
  // MAIN-CURRENT-INPUT - READONLY (shows total)
  document.querySelectorAll('.collection-main-current-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.setAttribute('readonly', 'true');
    el.style.backgroundColor = '#e9ecef';
    el.style.cursor = 'not-allowed';
    el.style.pointerEvents = 'none';

    const value = el.value.trim();
    if (value !== '' && value !== '0') {
      el.classList.remove('collection-input-empty');
      el.classList.add('collection-input-with-value');
    } else {
      el.classList.remove('collection-input-with-value');
      el.classList.add('collection-input-empty');
    }
  });

  // CATEGORY-CURRENT-INPUT - EDITABLE (user input)
  document.querySelectorAll('.collection-category-current-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.removeAttribute('readonly');
    el.style.backgroundColor = '#ffffff';
    el.style.cursor = 'text';
    el.style.pointerEvents = 'auto';

    const value = el.value.trim();
    if (value !== '' && value !== '0') {
      el.classList.remove('collection-input-empty');
      el.classList.add('collection-input-with-value');
    } else {
      el.classList.remove('collection-input-with-value');
      el.classList.add('collection-input-empty');
    }
  });

  // YEARLY-CURRENT-INPUT - READONLY (shows total)
  document.querySelectorAll('.collection-yearly-current-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.setAttribute('readonly', 'true');
    el.style.backgroundColor = '#e9ecef';
    el.style.cursor = 'not-allowed';
    el.style.pointerEvents = 'none';

    const value = el.value.trim();
    if (value !== '' && value !== '0') {
      el.classList.remove('collection-input-empty');
      el.classList.add('collection-input-with-value');
    } else {
      el.classList.remove('collection-input-with-value');
      el.classList.add('collection-input-empty');
    }
  });

  // MAIN-FORECAST-INPUT - READONLY (auto-calculated from categories)
  document.querySelectorAll('.collection-main-forecast-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.setAttribute('readonly', 'true');
    el.style.backgroundColor = '#e9ecef';
    el.style.cursor = 'not-allowed';
    el.style.pointerEvents = 'none';

    const value = el.value.trim();
    if (value !== '' && value !== '0') {
      el.classList.remove('collection-input-empty');
      el.classList.add('collection-input-with-value');
    } else {
      el.classList.remove('collection-input-with-value');
      el.classList.add('collection-input-empty');
    }
  });

  // CATEGORY-FORECAST-INPUT - EDITABLE
  document.querySelectorAll('.collection-category-forecast-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.removeAttribute('readonly');
    el.style.backgroundColor = '#ffffff';
    el.style.cursor = 'text';
    el.style.pointerEvents = 'auto';
    el.style.opacity = '1';

    const value = el.value.trim();
    if (value !== '' && value !== '0') {
      el.classList.remove('collection-input-empty');
      el.classList.add('collection-input-with-value');
    } else {
      el.classList.remove('collection-input-with-value');
      el.classList.add('collection-input-empty');
    }
  });

  // DATE INPUTS - EDITABLE
  document.querySelectorAll('.collection-forecast-date-input, .collection-category-date-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.removeAttribute('readonly');
    el.style.backgroundColor = '#ffffff';
    el.style.cursor = 'text';
    el.style.pointerEvents = 'auto';
    el.style.opacity = '1';
  });

  // PLAN INPUTS (Previous) - READONLY
  document.querySelectorAll('.collection-main-plan-input, .collection-category-plan-input, .collection-yearly-plan-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.setAttribute('readonly', 'true');
    el.style.backgroundColor = '#e9ecef';
    el.style.cursor = 'not-allowed';
    el.style.pointerEvents = 'none';
  });
};

const calculateMonthTotals = (): void => {
  // Find all unique months
  const monthKeys = new Set<string>();
  document.querySelectorAll('.collection-category-current-input[data-month]').forEach((input) => {
    const el = input as HTMLInputElement;
    const month = el.dataset.month;
    if (month) monthKeys.add(month);
  });

  // For each month, calculate total and update main-current-input
  monthKeys.forEach(monthKey => {
    const categoryInputs = document.querySelectorAll<HTMLInputElement>(`.collection-category-current-input[data-month="${monthKey}"]`);
    let total = 0;

    categoryInputs.forEach(input => {
      const value = getNumericValue(input);
      total += value;
    });

    const mainInput = document.querySelector<HTMLInputElement>(`.collection-main-current-input[data-month="${monthKey}"]`);
    if (mainInput) {
      // Set the total value
      if (total > 0) {
        mainInput.value = total.toString();
        mainInput.className = mainInput.className.replace('collection-input-empty', 'collection-input-with-value');
      } else {
        mainInput.value = '';
        mainInput.className = mainInput.className.replace('collection-input-with-value', 'collection-input-empty');
      }

      // Ensure it's readonly
      mainInput.setAttribute('readonly', 'true');
      mainInput.style.backgroundColor = '#e9ecef';
      mainInput.style.cursor = 'not-allowed';
      mainInput.style.pointerEvents = 'none';
    }
  });

  // Also calculate yearly totals
  const yearKeys = new Set<string>();
  document.querySelectorAll('.collection-category-current-input[data-year]').forEach((input) => {
    const el = input as HTMLInputElement;
    const year = el.dataset.year;
    if (year) yearKeys.add(year);
  });

  yearKeys.forEach(year => {
    const categoryInputs = document.querySelectorAll<HTMLInputElement>(`.collection-category-current-input[data-year="${year}"]`);
    let total = 0;

    categoryInputs.forEach(input => {
      const value = getNumericValue(input);
      total += value;
    });

    const yearlyInput = document.querySelector<HTMLInputElement>(`.collection-yearly-current-input[data-year="${year}"]`);
    if (yearlyInput) {
      if (total > 0) {
        yearlyInput.value = total.toString();
        yearlyInput.className = yearlyInput.className.replace('collection-input-empty', 'collection-input-with-value');
      } else {
        yearlyInput.value = '';
        yearlyInput.className = yearlyInput.className.replace('collection-input-with-value', 'collection-input-empty');
      }
    }
  });
};

// Add this to the OBL input change event
oblLastMonthInput.addEventListener('input', function() {
  const value = getNumericValue(this);
  if (value > 0) {
    this.className = this.className.replace('collection-input-empty', 'collection-input-with-value');
  } else {
    this.className = this.className.replace('collection-input-with-value', 'collection-input-empty');
  }
  validateOBLLimit(this);
});

const applyStickyHeaderOffsets = (): void => {
  const table = document.querySelector('.collection-month-columns-table') as HTMLTableElement;
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

// ========== FIXED ATTACH COLLECTION EVENT LISTENERS ==========

const attachCollectionEventListeners = (): void => {
  // Set up number inputs with allowOnlyNumbers
  document.querySelectorAll<HTMLInputElement>('.collection-number-input').forEach(input => {
    allowOnlyNumbers(input);
  });

  // Set up focus handlers for all number inputs to store old values
  document.querySelectorAll<HTMLInputElement>('.collection-number-input:not([readonly])').forEach(input => {
    input.addEventListener('focus', function() {
      this.setAttribute('data-old-value', getNumericValue(this).toString());
    });
  });

  // Category current inputs - handle changes
  document.querySelectorAll('.collection-category-current-input:not([readonly])').forEach((input) => {
    const el = input as HTMLInputElement;
    el.removeEventListener('input', handleCategoryChange);
    el.addEventListener('input', handleCategoryChange);
  });

  // Yearly category current inputs
  document.querySelectorAll('.collection-category-current-input[data-year]:not([readonly])').forEach((input) => {
    const el = input as HTMLInputElement;
    el.removeEventListener('input', handleCategoryChange);
    el.addEventListener('input', handleCategoryChange);
  });

 // Category Date inputs: when a category date changes, sync the main date
document.querySelectorAll<HTMLInputElement>('.collection-category-date-input').forEach(input => {
    input.addEventListener('change', function () {
        const monthKey = this.dataset.month;
        if (monthKey) syncMainDateFromItems(monthKey);
    });
});

// On form load, sync the main date for the current month
const currentMonthKey = formatMonthKey(new Date());
setTimeout(() => syncMainDateFromItems(currentMonthKey), 0);

  // Category forecast inputs - handle changes
  document.querySelectorAll('.collection-category-forecast-input').forEach((input) => {
    const el = input as HTMLInputElement;
    el.removeAttribute('readonly');
    el.style.backgroundColor = '#ffffff';
    el.style.cursor = 'text';
    el.style.pointerEvents = 'auto';

    el.removeEventListener('input', handleForecastCategoryChange);
    el.addEventListener('input', handleForecastCategoryChange);

    el.removeEventListener('focus', storeOldValue);
    el.addEventListener('focus', storeOldValue);
  });

  // Setup forecast validation
  setupForecastValidation();
};

// ===== HANDLE CATEGORY CHANGE - Updates main-current-input total =====
const handleCategoryChange = function(this: HTMLInputElement) {
  const month = this.dataset.month;
  const year = this.dataset.year;

  // Get old value for increase check
  const oldValue = parseInt(this.getAttribute('data-old-value') || '0');
  const currentValue = getNumericValue(this);
  const isIncreasing = currentValue > oldValue;

  // Check if this increase would exceed OBL
  if (isIncreasing) {
    const oblValue = getNumericValue(oblLastMonthInput);
    const currentTotal = calculateTotalCurrentValue();
    const increaseAmount = currentValue - oldValue;

    if (currentTotal + increaseAmount > oblValue) {
      // Revert to old value
      setNumericValue(this, oldValue);
      toast("warning", `Cannot increase - would exceed OBL limit`);
      return;
    }
  }

  // Update month total if this is a monthly category
  if (month) {
    const categoryInputs = document.querySelectorAll<HTMLInputElement>(
      `.collection-category-current-input[data-month="${month}"]`
    );
    let total = 0;
    categoryInputs.forEach(input => {
      total += getNumericValue(input);
    });

    const mainInput = document.querySelector<HTMLInputElement>(
      `.collection-main-current-input[data-month="${month}"]`
    );
    if (mainInput) {
      if (total > 0) {
        mainInput.value = total.toString();
        mainInput.className = mainInput.className.replace('collection-input-empty', 'collection-input-with-value');
      } else {
        mainInput.value = '';
        mainInput.className = mainInput.className.replace('collection-input-with-value', 'collection-input-empty');
      }
    }
  }

  // Update yearly total if this is a yearly category
  if (year) {
    const categoryInputs = document.querySelectorAll<HTMLInputElement>(
      `.collection-category-current-input[data-year="${year}"]`
    );
    let total = 0;
    categoryInputs.forEach(input => {
      total += getNumericValue(input);
    });

    const yearlyInput = document.querySelector<HTMLInputElement>(
      `.collection-yearly-current-input[data-year="${year}"]`
    );
    if (yearlyInput) {
      if (total > 0) {
        yearlyInput.value = total.toString();
        yearlyInput.className = yearlyInput.className.replace('collection-input-empty', 'collection-input-with-value');
      } else {
        yearlyInput.value = '';
        yearlyInput.className = yearlyInput.className.replace('collection-input-with-value', 'collection-input-empty');
      }
    }
  }

  // Update OBL validation after a short delay
  setTimeout(() => {
    validateOBLLimit(this);
  }, 50);
};

const getSaveButtons = (): HTMLButtonElement[] =>
  Array.from(document.querySelectorAll<HTMLButtonElement>(
    'button.collection-btn-primary[type="submit"][form="collectionForm"], .collection-form-actions button.collection-btn-primary[type="submit"]'
  ));

 const editCollection = async (id: number, existingCollection?: CollectionItem): Promise<void> => {
    const collection = existingCollection || collections.find(f => f.Id === id);

    if (!collection) {
        toast("error", "Collection not found");
        return;
    }

    editingId = collection.Id;

    const project = projects.find(p => p.Id === collection.Project?.Id);

    if (!canRead()) {
        toast("error", "❌ You don't have permission to view this collection");
        return;
    }

    // 🔐 RBAC view gate (mirrors auth.ts):
    //   - READ_ALL / FULL_CONTROL (Admin, SC Team, Controlling) → may open ANY project
    //   - every other role (GH Head, PM, Engg Lead/Sub1/Sub2/Mgr) → OWN projects only
    const assignedProjects = getAssignedProjects(projects);
    const isAssigned = assignedProjects.some(p => p.Id === collection.Project?.Id);

    if (!canReadAllProjects() && !isAssigned) {
        toast("error", "❌ You don't have permission to view this project");
        return;
    }

    const canEdit = authCanEditProject(project);
    const isFrozen = collection.Freeze || false;

    showForm(true);

    const remarksInternal = document.querySelector("#remarksInternal") as HTMLTextAreaElement;
    const remarksSprintCell = document.querySelector("#remarksSprintCell") as HTMLTextAreaElement;

    if (collection.Remarks) {
        const remarksData = collection.Remarks;
        if (remarksData.includes("|||")) {
            const parts = remarksData.split("|||");
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
            if (remarksSprintCell) {
                remarksSprintCell.value = '';
            }
        }
    }

    let collectionData: CollectionData = {};
    let revenueActualData: { [key: string]: number } = {};

    try {
        let collectionDataString = collection.CollectionData || "{}";
        collectionDataString = collectionDataString.trim();

        if (!collectionDataString || collectionDataString === '""' || collectionDataString === "null" || collectionDataString === "{}") {
            collectionDataString = "{}";
        }

        collectionData = JSON.parse(collectionDataString);

        if (!collectionData || typeof collectionData !== 'object' || Object.keys(collectionData).length === 0) {
            console.log("🔄 Collection data is empty, creating fresh structure");
            collectionData = createFreshCollectionStructure();
        }

        let revenueActualString = collection.RevenueActual || "{}";
        revenueActualString = revenueActualString.trim();

        if (!revenueActualString || revenueActualString === '""' || revenueActualString === "null") {
            revenueActualString = "{}";
        }

        revenueActualData = JSON.parse(revenueActualString);
        console.log("📊 Loaded RevenueActual in editCollection:", revenueActualData);

    } catch (error) {
        console.error("Error parsing CollectionData:", error);
        collectionData = createFreshCollectionStructure();
        revenueActualData = {};
        toast("warning", "Corrupted collection data detected. Loaded empty structure.");
    }

    // Apply progression
    collectionData = propagateMonthValues(collectionData);

    const formLoading = createLoadingOverlay("Loading collection form...");
    loaderHost().appendChild(formLoading);

    try {
        if (collection.Project?.Id) {
            const projectId = Number(collection.Project.Id);
            projectSelect.value = String(projectId);
            await onProjectChange(projectId, collection);
        }
    } finally {
        // ✅ Only the form loader is released here.
        // Save-button re-enable logic lives in the SUBMIT handler — editCollection
        // does not save anything, so it must not touch saveBtns.
        removeLoadingOverlay();
    }

    setTimeout(() => {
        const isReadOnly = isFrozen || !canEdit;
        setFormReadOnly(isReadOnly);

        // ✅ Sync BOTH save buttons (top + bottom) - label + visibility + frozen style
        const saveBtns: HTMLButtonElement[] = getSaveButtons();
        saveBtns.forEach((saveBtn: HTMLButtonElement) => {
            saveBtn.style.display = canEdit ? '' : 'none';

            if (isFrozen) {
                saveBtn.textContent = '💾 Save Collection Plan';
                saveBtn.title = 'Save only Further and Date field changes (Curr. Plan is locked)';
                saveBtn.style.backgroundColor = '#28a745';
                saveBtn.style.borderColor = '#28a745';
                saveBtn.style.color = '#ffffff';
                saveBtn.style.fontWeight = 'bold';
                saveBtn.style.cursor = 'pointer';
                saveBtn.setAttribute('type', 'submit');
            } else {
                saveBtn.textContent = '💾 Save Collection Plan';
                saveBtn.title = '';
                saveBtn.style.backgroundColor = '';
                saveBtn.style.borderColor = '';
                saveBtn.style.color = '';
                saveBtn.style.fontWeight = '';
                saveBtn.style.cursor = '';
                saveBtn.setAttribute('type', 'submit');
            }
        });

        // FIX 1: Use the shared helper so the project name is ALWAYS appended to the
        // title. Previously this block overwrote the "Edit Collection — {Project}"
        // title set 50ms earlier by renderCollectionInterface, causing the project
        // name to flash and disappear.
        updateEditFormHeader(isFrozen, canEdit);

        if (isFrozen) {
            // toast('info', '🔒 Project is frozen. Only Further and Date fields are editable. Click "Save Further Changes" to save.');
        } else if (!canEdit) {
            // toast('info', '🔒 You have view-only access to this project.');
        } else {
            // toast('success', '🔓 All fields are editable. Click "Save Collection Plan" to save.');
        }

    }, 300);

    // toast("info", `Viewing collection for ${collection.Title}`);
};
const applyAutoProgression = async (): Promise<void> => {
    if (collections.length === 0) {
        console.log("📊 No collections to progress");
        return;
    }

    const digest = await getDigest();
    const entityType = await getListItemType();
    let progressedCount = 0;
    let errorCount = 0;

    for (const collection of collections) {
        try {
            if (!collection.CollectionData) {
                continue;
            }

            const originalData = safeJsonParse(collection.CollectionData, {});

            // Check if progression is needed
            if (checkIfNeedsProgression(originalData)) {
                console.log(`🔄 Progressing collection: ${collection.Title}`);
                const progressedData = propagateCollectionValues(originalData);

                const payload = {
                    __metadata: { type: entityType },
                    CollectionData: JSON.stringify(progressedData),
                    UpdatedDate: new Date().toISOString()
                };

                const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${collection.Id})`, {
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

                if (response.ok) {
                    progressedCount++;
                } else {
                    errorCount++;
                    console.error(`❌ Failed to progress collection: ${collection.Id}`);
                }

                await new Promise(resolve => setTimeout(resolve, 100));
            }
        } catch (error) {
            errorCount++;
            console.error(`❌ Error progressing collection: ${collection.Id}`, error);
        }
    }

    if (progressedCount > 0) {
        await loadCollections();
        toast("success", `Monthly data progressed for ${progressedCount} collections`);
    } else if (errorCount > 0) {
        toast("warning", `Some errors occurred during monthly progression`);
    }
};

const propagateCollectionValues = (collectionData: CollectionData): CollectionData => {

  const updatedData: CollectionData = { ...collectionData };

  // Ensure yearly data exists before propagation
  if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) {
     const longTermMonths = generateIdealMonthStructure().slice(6);
    const lastMonthKey = longTermMonths[longTermMonths.length - 1];
    updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
  }

  // Detect if it's a new month by checking current month structure
  const originalMonths = Object.keys(collectionData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
  const idealMonths = generateIdealMonthStructure();
  const isNewMonth = originalMonths[0] !== idealMonths[0];

  // Apply monthly progression (long-term collection)
  const monthlyProgressed = propagateMonthValues(updatedData);

  // Apply yearly progression only when month changes
  const yearlyProgressed = propagateYearlyValues(monthlyProgressed, isNewMonth);
  return yearlyProgressed;
};

// ========== FIXED YEARLY PROPAGATION ==========

const propagateYearlyValues = (collectionData: CollectionData, isNewMonth: boolean = false): CollectionData => {
  if (!isNewMonth) {
    return collectionData;
  }

  const updatedData: CollectionData = { ...collectionData };

  if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) {
    return updatedData;
  }

  const yearlyData = updatedData.yearly;
  const updatedYearlyData: YearlyCollectionData = {};

  // CORRECT LOGIC:
  // For each year, current ACTUAL values become the new PLAN values for the SAME YEAR
  // ACTUAL values reset to 0

  Object.keys(yearlyData).forEach(year => {
    const yearData = yearlyData[year];
    if (yearData) {
      updatedYearlyData[year] = {
        plan: yearData.actual, // Current actual becomes new plan for same year
        actual: 0, // Reset actual to 0
        items: {}
      };

      // Apply same logic to categories
      if (yearData.items) {
        Object.keys(yearData.items).forEach(categoryKey => {
          const category = yearData.items![categoryKey];
          updatedYearlyData[year].items![categoryKey] = {
            name: categoryKey,
            plan: category.actual, // Category actual becomes category plan
            actual: 0 // Reset category actual to 0
          };
        });
      }
    }
  });

  updatedData.yearly = updatedYearlyData;
  return updatedData;
};

const importActualData = (): void => {
  // 🔐 RBAC: Upload Actual = a WRITE. Only FULL_CONTROL / CONTRIBUTE
  if (!canCreateRecord()) {
    toast("error", "❌ You don't have permission to import actual data");
    return;
  }

  console.log("📤 Import Actual button clicked - Opening modal");

  // Remove any existing modals first
  const existingModal = document.querySelector(".collection-import-modal");
  if (existingModal) {
    existingModal.remove();
  }

  // Create modal with INLINE STYLES (not relying on CSS classes)
  const modal = document.createElement('div');
  modal.className = 'collection-import-modal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.6);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 999999;
    overflow-y: auto;
    padding: 20px;
  `;

  modal.innerHTML = `
    <div class="collection-import-modal-content" style="
      background: white;
      border-radius: 12px;
      padding: 30px;
      max-width: 850px;
      width: 100%;
      max-height: 90vh;
      overflow-y: auto;
      box-shadow: 0 10px 40px rgba(0,0,0,0.3);
      position: relative;
    ">
      <div class="collection-import-modal-header" style="
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 2px solid #007bff;
        padding-bottom: 15px;
        margin-bottom: 25px;
      ">
        <h3 style="margin: 0; color: #495057; font-size: 20px;">
          📊 Import Monthly Actual Collection Data
        </h3>
        <button class="collection-close-modal" style="
          background: none;
          border: none;
          font-size: 28px;
          cursor: pointer;
          color: #6c757d;
          padding: 0 10px;
          transition: color 0.3s;
        ">&times;</button>
      </div>

      <div class="collection-import-section" style="margin-bottom: 25px;">
        <h4 style="color: #495057; margin-bottom: 12px; font-size: 15px;">
          Step 1: Download Template
        </h4>
        <button id="downloadTemplate" style="
          background: #007bff;
          color: white;
          border: none;
          padding: 10px 20px;
          border-radius: 6px;
          cursor: pointer;
          font-size: 14px;
          transition: background 0.3s;
        ">⬇️ Download Template</button>
      </div>

      <div class="collection-import-section" style="margin-bottom: 25px;">
        <h4 style="color: #495057; margin-bottom: 12px; font-size: 15px;">
          Step 2: Upload Excel File
        </h4>
        <div id="fileUploadArea" style="
          border: 2px dashed #ced4da;
          border-radius: 10px;
          padding: 50px 20px;
          text-align: center;
          cursor: pointer;
          transition: all 0.3s ease;
          background: #f8f9fa;
        ">
          <div style="font-size: 50px; margin-bottom: 10px;">📁</div>
          <p style="margin: 0; color: #495057; font-size: 16px;">
            Click to upload or drag & drop Excel file
          </p>
          <p style="margin: 5px 0 0 0; font-size: 12px; color: #6c757d;">
            Supported format: Excel with Project Reference, Month, and Actual Collection columns
          </p>
          <input type="file" id="excelFile" accept=".xlsx,.xls" style="display: none;">
          <div id="fileInfo" style="margin-top: 10px; font-size: 14px; color: #28a745;"></div>
        </div>
      </div>

      <div id="previewSection" style="display: none; margin-bottom: 25px;">
        <h4 style="color: #495057; margin-bottom: 12px; font-size: 15px;">
          Step 3: Preview & Import
        </h4>
        <div id="importPreview"></div>
        <div style="display: flex; gap: 10px; margin-top: 15px;">
          <button id="cancelImport" style="
            background: #6c757d;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
          ">Cancel</button>
          <button id="confirmImport" style="
            background: #28a745;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
          ">Import Data</button>
        </div>
        <div id="importStatus" style="
          margin-top: 10px;
          padding: 12px;
          border-radius: 6px;
          display: none;
        "></div>
      </div>
    </div>
  `;

  // Append modal to body
  document.body.appendChild(modal);
  console.log("✅ Modal created and appended to DOM");

  // Close modal when clicking outside
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });

  // Get all elements
  const fileUploadArea = document.getElementById('fileUploadArea') as HTMLElement;
  const excelFileInput = document.getElementById('excelFile') as HTMLInputElement;
  const fileInfo = document.getElementById('fileInfo') as HTMLElement;
  const previewSection = document.getElementById('previewSection') as HTMLElement;
  const importPreview = document.getElementById('importPreview') as HTMLElement;
  const confirmImport = document.getElementById('confirmImport') as HTMLButtonElement;
  const cancelImport = document.getElementById('cancelImport') as HTMLButtonElement;
  const closeModal = document.querySelector('.collection-close-modal') as HTMLButtonElement;
  const importStatus = document.getElementById('importStatus') as HTMLElement;
  const downloadTemplateBtn = document.getElementById('downloadTemplate') as HTMLButtonElement;

  let importData: any[] = [];
  let listEntityType: string = "";

  // Month name to number mapping
  const monthMap: { [key: string]: number } = {
    'jan': 1, 'january': 1, 'jan-': 1,
    'feb': 2, 'february': 2, 'feb-': 2,
    'mar': 3, 'march': 3, 'mar-': 3,
    'apr': 4, 'april': 4, 'apr-': 4,
    'may': 5, 'may-': 5,
    'jun': 6, 'june': 6, 'jun-': 6,
    'jul': 7, 'july': 7, 'jul-': 7,
    'aug': 8, 'august': 8, 'aug-': 8,
    'sep': 9, 'september': 9, 'sep-': 9,
    'oct': 10, 'october': 10, 'oct-': 10,
    'nov': 11, 'november': 11, 'nov-': 11,
    'dec': 12, 'december': 12, 'dec-': 12
  };

  // Parse month string to YYYY-MM format
  const parseMonth = (monthStr: string): string => {
    if (!monthStr) return '';

    if (/^\d{4}-\d{2}$/.test(monthStr)) {
      return monthStr;
    }

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

        return `${year}-${monthNum.toString().padStart(2, '0')}`;
      }
    }

    return '';
  };

  // Get correct List Entity Type
  const getListEntityType = async (): Promise<string> => {
    if (listEntityType) return listEntityType;
    try {
      const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
        headers: { Accept: "application/json;odata=verbose" },
      });
      const data = await res.json();
      listEntityType = data.d.ListItemEntityTypeFullName;
      return listEntityType;
    } catch (error) {
      console.error("Error getting list entity type:", error);
      throw error;
    }
  };

  // File Upload Handling
  fileUploadArea.addEventListener("click", () => excelFileInput.click());
  fileUploadArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    fileUploadArea.style.borderColor = '#007bff';
    fileUploadArea.style.background = '#e7f3ff';
  });
  fileUploadArea.addEventListener("dragleave", () => {
    fileUploadArea.style.borderColor = '#ced4da';
    fileUploadArea.style.background = '#f8f9fa';
  });
  fileUploadArea.addEventListener("drop", (e) => {
    e.preventDefault();
    fileUploadArea.style.borderColor = '#ced4da';
    fileUploadArea.style.background = '#f8f9fa';
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

  // Handle Excel File Selection
  const handleFileSelection = async (file: File) => {
    if (!file.name.toLowerCase().match(/\.(xlsx|xls)$/)) {
      toast("error", "Please select an Excel file (.xlsx or .xls)");
      return;
    }

    fileInfo.textContent = `Selected: ${file.name} (${(file.size / 1024).toFixed(2)} KB)`;

    try {
      const parsedData = await parseExcelFile(file);
      if (validateImportData(parsedData)) {
        importData = parsedData;
        showImportPreview(parsedData);
        previewSection.style.display = "block";
        previewSection.style.opacity = "1";
      } else {
        toast("error", "Invalid Excel format. Please check the required columns.");
      }
    } catch (error) {
      console.error("Error reading Excel file:", error);
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

          if (jsonData.length < 2) {
            return resolve([]);
          }

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

            const monthValue = row["Month"] || row["month"];

            if (monthValue) {
              if (typeof monthValue === "number") {
                const excelDate = new Date(Math.round((monthValue - 25569) * 86400 * 1000));
                const yyyy = excelDate.getFullYear();
                const mm = String(excelDate.getMonth() + 1).padStart(2, "0");
                row.parsedMonth = `${yyyy}-${mm}`;
              } else {
                const monthStr = monthValue.toString();
                row.parsedMonth = parseMonth(monthStr);
              }
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
      reader.onerror = (error) => {
        reject(new Error("Failed to read file"));
      };
      reader.readAsArrayBuffer(file);
    });
  };

  // Validate Columns
  const validateImportData = (data: any[]): boolean => {
    if (data.length === 0) {
      return false;
    }
    const first = data[0];
    const keys = Object.keys(first).map(k => k.toLowerCase());

    const hasProjectRef = keys.some(k => k.includes('project') && k.includes('reference'));
    const hasMonth = keys.some(k => k.includes('month'));

    return hasProjectRef && hasMonth;
  };

  // Show Preview
  const showImportPreview = (data: any[]) => {
    const monthSummary: { [key: string]: number } = {};

    data.forEach(row => {
      const month = row.parsedMonth;
      if (month) {
        monthSummary[month] = (monthSummary[month] || 0) + 1;
      }
    });

    const monthSummaryText = Object.keys(monthSummary)
      .map(month => {
        return `${month}: ${monthSummary[month]} records`;
      })
      .join(', ');

    const previewHTML = `
      <div style="padding: 20px; background: #f8f9fa; border-radius: 8px;">
        <h4 style="color: #28a745; margin-bottom: 15px;">✅ File Ready for Import</h4>
        <p><strong>Found ${data.length} records to import:</strong></p>
        <p><strong>Month Summary:</strong> ${monthSummaryText || 'No months found'}</p>
        <div style="max-height: 400px; overflow-y: auto;">
          <table style="width: 100%; border-collapse: collapse;">
            <thead style="background: #007bff; color: white;">
              <tr>
                <th style="padding: 10px; border: 1px solid #ddd;">Project Reference</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Month</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Actual Collection</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Project Match</th>
              </tr>
            </thead>
            <tbody>
              ${data
                .slice(0, 10)
                .map((row, index) => {
                  const projectRef = row["Project Reference"] || row["project_reference"] || 'N/A';
                  const month = row["Month"] || row["month"] || 'N/A';
                  const actualRevenue = row["Actual Collection"] || row["actual_collection"] || row["Actual"] || row["actual"] || '0';
                  const project = projects.find((p) => p.ProjectReference === projectRef);

                  return `
                    <tr style="${index % 2 === 0 ? 'background: #f8f9fa;' : 'background: white;'}">
                      <td style="padding: 8px; border: 1px solid #ddd;">${projectRef}</td>
                      <td style="padding: 8px; border: 1px solid #ddd;">${month} (${row.parsedMonth || 'Invalid'})</td>
                      <td style="padding: 8px; border: 1px solid #ddd;">${actualRevenue}</td>
                      <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${project ? "✅" : "❌"}</td>
                    </tr>
                  `;
                })
                .join("")}
            </tbody>
          </table>
        </div>
        ${data.length > 10 ? `<p style="margin-top: 10px; font-style: italic;">... and ${data.length - 10} more records</p>` : ''}
        <div style="margin-top: 20px; padding: 15px; background: #e7f3ff; border-radius: 5px;">
          <p style="margin: 0; color: #0056b3;"><strong>Ready to import?</strong> Click the "Import Data" button below.</p>
          <p style="margin: 5px 0 0 0; font-size: 12px;"><strong>Note:</strong> Only the imported month will be updated, all other data preserved</p>
        </div>
      </div>`;

    importPreview.innerHTML = previewHTML;
  };

  // Perform Import
  const performImport = async (): Promise<void> => {
    confirmImport.disabled = true;
    confirmImport.textContent = "Importing...";
    importStatus.style.display = "block";
    importStatus.className = "collection-import-status";
    importStatus.textContent = "Updating Collection Data...";

    let success = 0;
    let error = 0;
    let notFound = 0;
    let invalidMonth = 0;
    let invalidAmount = 0;

    try {
      // Helper functions
      const normalizeProjectRef = (value: any): string => {
        return String(value ?? "")
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase();
      };

      const getExcelValue = (row: any, keys: string[]): any => {
        for (const key of keys) {
          if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") {
            return row[key];
          }
        }
        return null;
      };

      const parseActualAmount = (value: any): number => {
        if (value === null || value === undefined || value === "") return 0;
        const cleaned = String(value).replace(/,/g, "").replace(/\s/g, "").trim();
        const numberValue = Number(cleaned);
        return Number.isFinite(numberValue) ? numberValue : NaN;
      };

      const parseJsonObject = (value: any): { [key: string]: any } => {
        if (value === null || value === undefined || value === "" || value === "{}" || value === "null" || value === "undefined") {
          return {};
        }
        if (typeof value === "object") {
          return Array.isArray(value) ? {} : { ...value };
        }
        try {
          const parsed = JSON.parse(String(value));
          if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
            return { ...parsed };
          }
        } catch (jsonError) {
          console.warn("⚠️ Invalid JSON found:", value, jsonError);
        }
        return {};
      };

      const entityType = await getListEntityType();

      if (!entityType) {
        throw new Error("Unable to determine SharePoint List Entity Type.");
      }

      console.log("========== EXCEL IMPORT START ==========");

      for (const row of importData) {
        try {
          // Get Project Reference
          const rawProjectRef = getExcelValue(row, [
            "Project Reference",
            "project_reference",
            "ProjectReference",
            "project reference"
          ]);

          const projectRef = normalizeProjectRef(rawProjectRef);

          if (!projectRef) {
            notFound++;
            console.warn("⚠️ Project Reference missing:", row);
            continue;
          }

          // Get Month
          let month = row.parsedMonth;

          if (!month) {
            const rawMonth = getExcelValue(row, ["Month", "month"]);

            if (rawMonth) {
              if (typeof rawMonth === "number") {
                const excelDate = new Date(Math.round((rawMonth - 25569) * 86400 * 1000));
                if (!isNaN(excelDate.getTime())) {
                  const yyyy = excelDate.getFullYear();
                  const mm = String(excelDate.getMonth() + 1).padStart(2, "0");
                  month = `${yyyy}-${mm}`;
                }
              } else {
                month = parseMonth(String(rawMonth).trim());
              }
            }
          }

          if (!month || !/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
            invalidMonth++;
            console.warn("⚠️ Invalid month:", rawProjectRef, row);
            continue;
          }

          // Get Actual Collection
          const rawActual = getExcelValue(row, [
            "Actual Collection",
            "actual_collection",
            "Actual",
            "actual",
            "ActualCollection"
          ]);

          const actualRevenue = parseActualAmount(rawActual);

          if (!Number.isFinite(actualRevenue)) {
            invalidAmount++;
            console.warn("⚠️ Invalid Actual Collection:", { projectRef: rawProjectRef, month, value: rawActual });
            continue;
          }

          // Find Project
          const project = projects.find(
            (p) => normalizeProjectRef(p.ProjectReference) === projectRef
          );

          if (!project) {
            notFound++;
            console.warn(`⚠️ Project not found: ${rawProjectRef}`);
            continue;
          }

          // Find Collection Record
          const collection = collections.find((f) => {
            const collectionProjectTitle = typeof f.Project === "object" ? f.Project?.Title : f.Project;
            return String(collectionProjectTitle ?? "").trim().toLowerCase() ===
                   String(project.Title ?? "").trim().toLowerCase();
          });

          if (!collection?.Id) {
            notFound++;
            console.warn(`⚠️ Collection record not found for project: ${rawProjectRef}`, {
              projectId: project.Id,
              projectTitle: project.Title
            });
            continue;
          }

          console.log("----------------------------------------");
          console.log("Importing:", {
            projectReference: rawProjectRef,
            normalizedProjectReference: projectRef,
            projectId: project.Id,
            projectTitle: project.Title,
            collectionId: collection.Id,
            month,
            actualRevenue
          });

          // Get current collection record
          const currentCollectionResponse = await fetch(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${collection.Id})?$select=Id,RevenueActual,CollectionData`,
            {
              method: "GET",
              headers: { Accept: "application/json;odata=verbose" }
            }
          );

          if (!currentCollectionResponse.ok) {
            const errorText = await currentCollectionResponse.text();
            throw new Error(`Failed to read Collection record. HTTP ${currentCollectionResponse.status}: ${errorText}`);
          }

          const currentData = await currentCollectionResponse.json();
          const existingRevenueActual = currentData?.d?.RevenueActual ?? "{}";
          const existingCollectionData = currentData?.d?.CollectionData ?? "{}";

          // Update RevenueActual
          const revenueActualObject = parseJsonObject(existingRevenueActual) as { [key: string]: number };
          revenueActualObject[month] = actualRevenue;
          const updatedRevenueActual = JSON.stringify(revenueActualObject);

          // Update CollectionData
          const collectionDataObject = parseJsonObject(existingCollectionData) as { [key: string]: any };

          if (!collectionDataObject[month] || typeof collectionDataObject[month] !== "object") {
            collectionDataObject[month] = {
              plan: 0,
              actual: 0,
              monthActual: 0,
              forecast: 0,
              date: "",
              items: {}
            };
          }

          const monthData = collectionDataObject[month];
          if (typeof monthData.plan !== "number") monthData.plan = Number(monthData.plan) || 0;
          monthData.actual = Number(monthData.actual) || 0;
          monthData.monthActual = actualRevenue;
          if (typeof monthData.forecast !== "number") monthData.forecast = Number(monthData.forecast) || 0;
          if (monthData.date === undefined || monthData.date === null) monthData.date = "";
          if (!monthData.items || typeof monthData.items !== "object") monthData.items = {};

          const updatedCollectionData = JSON.stringify(collectionDataObject);

          // Create SharePoint payload
          const payload = {
            __metadata: { type: entityType },
            RevenueActual: updatedRevenueActual,
            CollectionData: updatedCollectionData,
            UpdatedDate: new Date().toISOString()
          };

          // Get Request Digest
          const digest = await getDigest();

          if (!digest) {
            throw new Error("SharePoint Request Digest is empty.");
          }

          // Save to SharePoint
          const saveResponse = await fetch(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${collection.Id})`,
            {
              method: "POST",
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest,
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE"
              },
              body: JSON.stringify(payload)
            }
          );

          if (saveResponse.ok) {
            success++;
            console.log(`✅ Successfully saved: Project Reference = ${rawProjectRef}, Month = ${month}, Actual = ${actualRevenue}`);
          } else {
            error++;
            const responseText = await saveResponse.text();
            console.error("❌ SharePoint SAVE FAILED", {
              projectReference: rawProjectRef,
              month,
              actualRevenue,
              collectionId: collection.Id,
              status: saveResponse.status,
              response: responseText
            });
          }

          await new Promise((resolve) => setTimeout(resolve, 100));

        } catch (rowError) {
          error++;
          console.error("❌ Excel import row failed:", { row, error: rowError });
        }
      }

      const totalProcessed = success + error + notFound + invalidMonth + invalidAmount;
      const skipped = Math.max(0, importData.length - totalProcessed);

      console.log("========== EXCEL IMPORT COMPLETE ==========");
      console.log({ total: importData.length, success, error, notFound, invalidMonth, invalidAmount, skipped });

      if (error === 0 && notFound === 0 && invalidMonth === 0 && invalidAmount === 0) {
        importStatus.className = "collection-import-status success";
        importStatus.style.background = "#d4edda";
        importStatus.style.color = "#155724";
        importStatus.style.border = "1px solid #c3e6cb";
      } else {
        importStatus.className = "collection-import-status warning";
        importStatus.style.background = "#fff3cd";
        importStatus.style.color = "#856404";
        importStatus.style.border = "1px solid #ffeeba";
      }

      importStatus.innerHTML = `
        <strong>Excel Import Complete</strong>

        ✅ Success: ${success} records

        ❌ Errors: ${error} records

        🔍 Project Not Found: ${notFound} records

        📅 Invalid Month: ${invalidMonth} records

        💰 Invalid Amount: ${invalidAmount} records

        ⏭️ Skipped: ${skipped} records

        <small>RevenueActual and CollectionData were updated. Existing months were preserved.</small>
      `;

      await loadCollections();
      initializePagination(collections);

    } catch (importError) {
      console.error("❌ EXCEL IMPORT FAILED:", importError);
      importStatus.className = "collection-import-status error";
      importStatus.style.background = "#f8d7da";
      importStatus.style.color = "#721c24";
      importStatus.style.border = "1px solid #f5c6cb";
      importStatus.innerHTML = `
        <strong>Import Failed</strong>

        ❌ ${importError instanceof Error ? importError.message : String(importError)}
      `;
    } finally {
      confirmImport.disabled = false;
      confirmImport.textContent = "Import Data";
    }
  };

  // Download Template
  downloadTemplateBtn.addEventListener("click", () => {
    if (!projects || projects.length === 0) {
      toast("error", "No projects available to create template");
      return;
    }

    const wsData = [["Project Reference", "Month", "Actual Collection"]];

    projects.forEach((p) => {
      wsData.push([p.ProjectReference || "", "", ""]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Collection Template");

    const fileName = `Collection_Import_Template.xlsx`;

    XLSX.writeFile(wb, fileName);
  });

  // Modal Controls
  const closeModalHandler = () => {
    if (modal.parentNode) {
      modal.remove();
    }
  };

  closeModal.addEventListener("click", closeModalHandler);
  cancelImport.addEventListener("click", closeModalHandler);
  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeModalHandler();
  });

  confirmImport.addEventListener("click", performImport);
};

const initializeFilters = (): void => {
  const currentYear = new Date().getFullYear();

  // ✅ "check filter": Filters start CLEAN - the placeholder option is selected
  //    and nothing is pre-picked (no default Year/Month, badge stays 0).
  filterYear.innerHTML = '<option value="">All Years</option>';
  for (let year = currentYear - 1; year <= currentYear + 2; year++) {
    filterYear.innerHTML += `<option value="${year}">${year}</option>`;
  }

  // Month filter
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  filterMonth.innerHTML = '<option value="">All Months</option>';
  for (let month = 1; month <= 12; month++) {
    filterMonth.innerHTML += `<option value="${month}">${monthNames[month - 1]}</option>`;
  }

  // Helper function to get unique values from assigned projects
  // 🔐 RBAC: filter options always mirror the projects the role may see
  // (READ_ALL roles → all projects; everyone else → OWN projects only).
  const getFilterValues = (key: keyof ProjectItem): string[] => {
    const values = new Set<string>();

    const filterProjects = canReadAllProjects()
      ? projects
      : getAssignedProjects(projects);

    filterProjects.forEach(project => {
      let value = project[key];
      if (value) {
        if (key === 'ProjectManager' && typeof value === 'object') {
          value = (value as { Title: string }).Title;
        }
        // ✅ Handle GroupHead1 as Person/Group
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

  // Project filter
  const filterProjectTitles = getFilterValues('Title');
  filterProject.innerHTML = '<option value="">All Projects</option>' +
    filterProjectTitles.map(project => `<option value="${project}">${project}</option>`).join("");

  // Customer filter
  const filterCustomers = getFilterValues('CustomerName');
  filterCustomer.innerHTML = '<option value="">All Customers</option>' +
    filterCustomers.map(customer => `<option value="${customer}">${customer}</option>`).join("");

  // Project Manager filter
  const filterManagers = getFilterValues('ProjectManager');
  filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
    filterManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");

  // ✅ Group Head filter - using GroupHead1
  const filterGroupHeads = getFilterValues('GroupHead1');
  filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
    filterGroupHeads.map(head => `<option value="${head}">${head}</option>`).join("");

  // Project Reference filter
  const filterReferences = getFilterValues('ProjectReference');
  filterProjectReference.innerHTML = '<option value="">All References</option>' +
    filterReferences.map(ref => `<option value="${ref}">${ref}</option>`).join("");
};

const sortCollectionData = (
  data: CollectionItem[]
): CollectionItem[] => {
  if (!currentSortConfig.column) {
    return [...data];
  }

  const sorted = [...data];

  const getProject = (item: CollectionItem): any => {
    return projects.find(
      p => p.Title === item.Project?.Title
    );
  };

  const getMonthValue = (
    item: CollectionItem,
    monthKey: string,
    subColumn?: string
  ): number => {
    const collectionData: CollectionData =
      safeJsonParse(
        item.CollectionData,
        {}
      );

    const monthData =
      collectionData[monthKey] as
        | CollectionMonthData
        | undefined;

    if (!monthData) {
      return 0;
    }

    const items = monthData.items || {};

    switch (subColumn) {
      case "normal-previous":
        return (
          (items.Normal?.plan || 0) +
          (items["Normal Overdue"]?.plan || 0)
        );

      case "normal-current":
        return (
          (items.Normal?.actual || 0) +
          (items["Normal Overdue"]?.actual || 0)
        );

      case "retention-previous":
        return (
          (items.Retention?.plan || 0) +
          (items["Retention Overdue"]?.plan || 0)
        );

      case "retention-current":
        return (
          (items.Retention?.actual || 0) +
          (items["Retention Overdue"]?.actual || 0)
        );

      case "total-previous":
        return (
          (items.Normal?.plan || 0) +
          (items["Normal Overdue"]?.plan || 0) +
          (items.Retention?.plan || 0) +
          (items["Retention Overdue"]?.plan || 0)
        );

      case "total-current":
        return (
          (items.Normal?.actual || 0) +
          (items["Normal Overdue"]?.actual || 0) +
          (items.Retention?.actual || 0) +
          (items["Retention Overdue"]?.actual || 0)
        );

      case "actual":
        return monthData.monthActual || 0;

      case "forecast":
        return monthData.forecast || 0;

      case "forecast-total":
        return (monthData.monthActual || 0) + (monthData.forecast || 0);

      default:
        return 0;
    }
  };

  const getSortValue = (
    item: CollectionItem
  ): string | number => {
    const project = getProject(item);

    switch (currentSortConfig.column) {

      case "freeze":
        return item.Freeze ? 1 : 0;

      case "projectReference":
        return project?.ProjectReference || "";

      case "project":
        return item.Project?.Title || "";

      case "projectManager": {
        const manager =
          typeof project?.ProjectManager === "object"
            ? project?.ProjectManager?.Title
            : project?.ProjectManager;

        return manager || "";
      }

      case "groupHead": {
  const groupHead1 = getGroupHeadDisplayName(project?.GroupHead1);
  return groupHead1 || "";
}

      case "customer":
        return project?.CustomerName || "";

      case "receivable":
      case "obl":
        return Number(item.OBL || 0);

      case "month": {
        const displayMonths =
          getDisplayMonths();

        const monthIndex =
          currentSortConfig.monthIndex ?? 0;

        const monthKey =
          displayMonths[monthIndex];

        if (!monthKey) {
          return 0;
        }

        return getMonthValue(
          item,
          monthKey,
          currentSortConfig.subColumn
        );
      }

      default:
        return "";
    }
  };

  sorted.sort((a, b) => {
    const valueA = getSortValue(a);
    const valueB = getSortValue(b);

    let comparison = 0;

    if (
      typeof valueA === "number" &&
      typeof valueB === "number"
    ) {
      comparison = valueA - valueB;
    } else {
      comparison =
        String(valueA)
          .toLowerCase()
          .localeCompare(
            String(valueB).toLowerCase()
          );
    }

    return currentSortConfig.direction === "asc"
      ? comparison
      : -comparison;
  });

  return sorted;
};
// ========== UPDATED FILTER TABLE ==========
const filterTable = (): void => {
  const f = appliedFilters;

  filteredData = collections.filter(collection => {
    const title = collection.Project?.Title;
    const proj = projects.find(p => p.Title === title);

    if (f.project && title !== f.project) return false;
    if (f.customer && getSafeString(proj?.CustomerName, '') !== f.customer) return false;
    if (f.manager && getProjectManagerDisplayName(proj?.ProjectManager) !== f.manager) return false;
    if (f.groupHead && getGroupHeadDisplayName(proj?.GroupHead1) !== f.groupHead) return false;
    if (f.reference && getSafeString(proj?.ProjectReference, '') !== f.reference) return false;

    // Year / Month are matched against the stored month keys of the row
    if (f.year || f.month) {
      const data = safeJsonParse(collection.CollectionData, {});
      const monthKeys = Object.keys(data).filter(key => /^\d{4}-\d{2}$/.test(key));

      if (f.year && f.month) {
        if (!monthKeys.includes(`${f.year}-${f.month}`)) return false;
      } else if (f.year) {
        if (!monthKeys.some(key => key.startsWith(`${f.year}-`))) return false;
      } else if (f.month) {
        if (!monthKeys.some(key => key.endsWith(`-${f.month}`))) return false;
      }
    }

    return true;
  });

  currentPage = 1;
  updatePagination();

  console.log(`🔎 Filtered ${filteredData.length} of ${collections.length} collection rows`);
};

/* Clear = wipe the popup fields AND the applied filters (used by the popup
   Clear button, the popup Reset Filter button and the "Clear all" chip). */
const clearFilters = (): void => {
  appliedFilters = { ...EMPTY_FILTERS };
  restoreAppliedFilters();
  syncFilterUI();
  filterTable();

  // totals header reflects the (now unfiltered) data again
  renderTotalsHeader(getDisplayMonths(), appliedFilters.year && appliedFilters.month ? `${appliedFilters.year}-${appliedFilters.month}` : '');
};

form.addEventListener("submit", async (e: Event) => {
  e.preventDefault();

  // ✅ Only run the flow when a Save button triggered the submit
  const submitter = (e as any).submitter as HTMLButtonElement | null;
  const isSaveClick =
    submitter && submitter.classList.contains("collection-btn-primary");
  if (!isSaveClick) return;

  if (isSavingCollection) {
    console.log("⛔ submit blocked: save already in progress");
    return;
  }

  // ---- Pre-checks (nothing shown yet) ----
  const projectId = parseInt(projectSelect.value);
  const project = projects.find(p => p.Id === projectId);

  let currentCollection: CollectionItem | undefined;
  let isFrozenSave = false;

  if (editingId) {
    currentCollection = collections.find(f => f.Id === editingId);

    if (!authCanEditProject(project)) {
      toast("error", "❌ You don't have permission to edit this collection");
      return;
    }

    if (currentCollection && currentCollection.Freeze) {
      isFrozenSave = true;
    }
  } else {
    if (!canCreateRecord()) {
      toast("error", "❌ You don't have permission to create collections");
      return;
    }
  }

  if (!project) {
    toast("error", "Please select a project");
    return;
  }

  // Frozen → its own further-only flow (its own dialog is inside saveFurtherOnly)
  if (isFrozenSave && currentCollection) {
    await saveFurtherOnly(currentCollection, project);
    return;
  }

  const isNewProject =
    !editingId && !collections.some(f => f.Project?.Id === projectId);

  if (
    !isNewProject &&
    collections.some(f => f.Project?.Id === projectId && f.Id !== editingId)
  ) {
    toast("error", "Collection already exists for this project!");
    return;
  }

  // ============================================================
  // ✅ 1) CONFIRMATION POPUP (same look as revenue.ts)
  //    Save the guard here so the outer finally can always clear it,
  //    and hoist the buttons so the outer finally can always restore them.
  // ============================================================
  isSavingCollection = true;
  const saveBtns: HTMLButtonElement[] = getSaveButtons();

  try {
    const confirmResult = await Swal.fire({
      title: editingId ? "Save Changes?" : "Save Collection Plan?",
      html: editingId
        ? `Do you want to save the changes made to <strong>${escapeHtml(project.Title)}</strong>?`
        : `Do you want to save the collection plan for <strong>${escapeHtml(project.Title)}</strong>?`,
      icon: "question",
      showCancelButton: true,
      confirmButtonText: editingId ? "Yes, Save Changes" : "Yes, Save",
      cancelButtonText: "Cancel",
      confirmButtonColor: "#605eee",
      cancelButtonColor: "#6c757d",
      reverseButtons: true,
      focusCancel: true,
      allowOutsideClick: false,
      customClass: { popup: "proc-confirm-popup" },
      didOpen: bumpSwalZIndex,
    });

    if (!confirmResult.isConfirmed) {
      console.log("🔵 save cancelled by user");
      return; // finally resets the guard + buttons
    }

    // ============================================================
    // ✅ 2) ZERO-OBL confirm (unchanged logic, now runs AFTER the save popup)
    // ============================================================
    const oblValue = getNumericValue(oblLastMonthInput);
    if (isNewProject && oblValue === 0) {
      const confirmZeroOBL = await Swal.fire({
        icon: "question",
        title: "Zero OBL Value",
        text: "You're setting OBL to zero for a new project. This means there's no budget limit. Continue?",
        showCancelButton: true,
        confirmButtonText: "Yes, Continue",
        cancelButtonText: "No, Set OBL",
        confirmButtonColor: "#605eee",
        cancelButtonColor: "#6c757d",
        reverseButtons: true,
        focusCancel: true,
        customClass: { popup: "proc-confirm-popup" },
        didOpen: bumpSwalZIndex,
      });

      if (!confirmZeroOBL.isConfirmed) {
        oblLastMonthInput.focus();
        return; // finally resets the guard + buttons
      }
    }

    // ============================================================
    // ✅ 3) VALIDATION — bail out cleanly if anything fails
    // ============================================================
    if (!validateForecastData()) return;
    if (!validateFormSubmission()) return;

    // ============================================================
    // ✅ 4) LOCK SAVE BUTTONS + show the loader
    // ============================================================
    saveBtns.forEach((btn: HTMLButtonElement) => {
      btn.disabled = true;
      btn.dataset.saving = "true";
      btn.textContent = "💾 Saving...";
    });

    const saveLoading = createLoadingOverlay("Saving collection...");
    loaderHost().appendChild(saveLoading);

    // ============================================================
    // ✅ 5) ACTUAL SAVE (unchanged logic)
    // ============================================================
    const listItemType = await getListItemType();
    const digest = await getDigest();

    // ===== GET REMARKS VALUES =====
    const remarksInternalInput = document.querySelector(
      "#remarksInternal"
    ) as HTMLTextAreaElement;
    const remarksSprintCellInput = document.querySelector(
      "#remarksSprintCell"
    ) as HTMLTextAreaElement;

    const internalRemark = remarksInternalInput
      ? remarksInternalInput.value.trim()
      : "";
    const sprintCellRemark = remarksSprintCellInput
      ? remarksSprintCellInput.value.trim()
      : "";
    const combinedRemarks = internalRemark + "|||" + sprintCellRemark;

    // ===== BUILD COLLECTION DATA =====
    const collectionData: CollectionData = {};

    const allStorageMonths = getAllMonthsForStorage();

    // IMPORTANT: If editing, load existing data first to preserve all months
    let existingCollectionData: CollectionData = {};
    if (editingId) {
      const existingCollection = collections.find(f => f.Id === editingId);
      if (existingCollection && existingCollection.CollectionData) {
        existingCollectionData = safeJsonParse(
          existingCollection.CollectionData,
          {}
        );
      }
    }

    // Type guard to check if data is CollectionMonthData
    const isCollectionMonthDataLocal = (
      data: any
    ): data is CollectionMonthData => {
      return (
        data && typeof data === "object" && "plan" in data && "actual" in data
      );
    };

    // FIRST: Preserve ALL existing historical months from the database
    const idealMonths = generateIdealMonthStructure();
    const historicalMonths = idealMonths.slice(0, 36);

    historicalMonths.forEach(monthKey => {
      const existingData = existingCollectionData[monthKey];
      if (existingData && isCollectionMonthDataLocal(existingData)) {
        collectionData[monthKey] = JSON.parse(JSON.stringify(existingData));
      } else {
        collectionData[monthKey] = {
          plan: 0,
          actual: 0,
          monthActual: 0,
          forecast: 0,
          date: "",
          items: {},
        };
      }
    });

    // SECOND: Process detailed months (36-38) and long-term months (39+)
    allStorageMonths.forEach((monthKey, index) => {
      if (index < 36) return;

      const mainCurrentInput = document.querySelector<HTMLInputElement>(
        `input.collection-main-current-input[data-month="${monthKey}"]`
      );
      const mainPlanInput = document.querySelector<HTMLInputElement>(
        `input.collection-main-plan-input[data-month="${monthKey}"]`
      );
      const mainForecastInput = document.querySelector<HTMLInputElement>(
        `input.collection-main-forecast-input[data-month="${monthKey}"]`
      );
      const mainDateInput = document.querySelector<HTMLInputElement>(
        `input.collection-forecast-date-input[data-month="${monthKey}"]`
      );

      if (mainCurrentInput && mainPlanInput) {
        // ✅ FIX: Start with existing data to preserve monthActual
        let monthData: CollectionMonthData;

        const existingData = existingCollectionData[monthKey];
        if (existingData && isCollectionMonthDataLocal(existingData)) {
          // Preserve existing data including monthActual
          monthData = JSON.parse(JSON.stringify(existingData));
          // Update plan and actual from inputs
          monthData.plan = getNumericValue(mainPlanInput);
          monthData.actual = getNumericValue(mainCurrentInput);
        } else {
          // Create new month data
          monthData = {
            plan: getNumericValue(mainPlanInput),
            actual: getNumericValue(mainCurrentInput),
            monthActual: 0,
            forecast: 0,
            date: "",
            items: {},
          };
        }

        // ✅ FIX: For current month (index 36), preserve monthActual from existing data
        if (index === 36) {
          const existingDataForMonth = existingCollectionData[monthKey];
          if (
            existingDataForMonth &&
            isCollectionMonthDataLocal(existingDataForMonth) &&
            existingDataForMonth.monthActual !== undefined
          ) {
            monthData.monthActual = existingDataForMonth.monthActual;
          } else {
            // Try to get from RevenueActual
            let monthActualValue = 0;
            try {
              if (editingId) {
                const collection = collections.find(f => f.Id === editingId);
                if (collection && collection._revenueActual) {
                  monthActualValue =
                    Number(collection._revenueActual[monthKey]) || 0;
                }
              } else {
                const collection = collections.find(
                  f => f.Project?.Id === projectId
                );
                if (collection && collection._revenueActual) {
                  monthActualValue =
                    Number(collection._revenueActual[monthKey]) || 0;
                }
              }
            } catch (e) {
              monthActualValue = 0;
            }
            monthData.monthActual = monthActualValue;
          }

          // Update forecast and date from inputs
          if (mainForecastInput) {
            monthData.forecast = getNumericValue(mainForecastInput);
          }
          if (mainDateInput) {
            monthData.date = mainDateInput.value;
          }
        } else {
          if (monthData.monthActual === undefined) {
            monthData.monthActual = 0;
          }
        }

        // Detailed months (36-38): full categories
        if (index >= 36 && index <= 38) {
          const categories = [
            "Normal",
            "Normal Overdue",
            "Retention",
            "Retention Overdue",
          ];
          let totalActual = 0;

          categories.forEach(categoryKey => {
            const categoryPlanInput = document.querySelector<HTMLInputElement>(
              `input.collection-category-plan-input[data-category="${categoryKey}"][data-month="${monthKey}"]`
            );
            const categoryCurrentInput =
              document.querySelector<HTMLInputElement>(
                `input.collection-category-current-input[data-category="${categoryKey}"][data-month="${monthKey}"]`
              );
            const categoryForecastInput =
              document.querySelector<HTMLInputElement>(
                `input.collection-category-forecast-input[data-category="${categoryKey}"][data-month="${monthKey}"]`
              );
            const categoryDateInput = document.querySelector<HTMLInputElement>(
              `input.collection-category-date-input[data-category="${categoryKey}"][data-month="${monthKey}"]`
            );

            if (categoryPlanInput && categoryCurrentInput) {
              let categoryData: CollectionCategory;
              const existingCat = monthData.items?.[categoryKey];

              if (existingCat) {
                categoryData = JSON.parse(JSON.stringify(existingCat));
                categoryData.plan = getNumericValue(categoryPlanInput);
                categoryData.actual = getNumericValue(categoryCurrentInput);
              } else {
                categoryData = {
                  name: categoryKey,
                  plan: getNumericValue(categoryPlanInput),
                  actual: getNumericValue(categoryCurrentInput),
                  forecast: 0,
                  date: "",
                };
              }

              if (index === 36) {
                if (categoryForecastInput) {
                  categoryData.forecast = getNumericValue(categoryForecastInput);
                }
                if (categoryDateInput) {
                  categoryData.date = categoryDateInput.value;
                }
              }

              totalActual += categoryData.actual;
              monthData.items![categoryKey] = categoryData;
            }
          });

          monthData.actual = totalActual;
        }
        // Long-term months (39+): simple categories
        else if (index >= 39) {
          const categories = ["Normal", "Retention"];
          let totalActual = 0;

          categories.forEach(categoryKey => {
            const categoryPlanInput = document.querySelector<HTMLInputElement>(
              `input.collection-category-plan-input[data-category="${categoryKey}"][data-month="${monthKey}"]`
            );
            const categoryCurrentInput =
              document.querySelector<HTMLInputElement>(
                `input.collection-category-current-input[data-category="${categoryKey}"][data-month="${monthKey}"]`
              );

            if (categoryPlanInput && categoryCurrentInput) {
              let categoryData;
              const existingCat = monthData.items?.[categoryKey];

              if (existingCat) {
                categoryData = JSON.parse(JSON.stringify(existingCat));
                categoryData.plan = getNumericValue(categoryPlanInput);
                categoryData.actual = getNumericValue(categoryCurrentInput);
              } else {
                categoryData = {
                  name: categoryKey,
                  plan: getNumericValue(categoryPlanInput),
                  actual: getNumericValue(categoryCurrentInput),
                };
              }

              totalActual += categoryData.actual;
              monthData.items![categoryKey] = categoryData;
            }
          });

          monthData.actual = totalActual;
        }

        collectionData[monthKey] = monthData;
      }
    });

    // ===== PROCESS YEARLY DATA =====
    const yearlyData: YearlyCollectionData = {};

    const yearlyInputs = document.querySelectorAll<HTMLInputElement>(
      ".collection-yearly-current-input, .collection-yearly-plan-input"
    );
    const yearlyYears = new Set<string>();

    yearlyInputs.forEach(input => {
      const year = input.dataset.year;
      if (year) yearlyYears.add(year);
    });

    Array.from(yearlyYears).forEach(year => {
      const yearlyPlanInput = document.querySelector<HTMLInputElement>(
        `input.collection-yearly-plan-input[data-year="${year}"]`
      );
      const yearlyCurrentInput = document.querySelector<HTMLInputElement>(
        `input.collection-yearly-current-input[data-year="${year}"]`
      );

      if (yearlyPlanInput && yearlyCurrentInput) {
        yearlyData[year] = {
          plan: getNumericValue(yearlyPlanInput),
          actual: getNumericValue(yearlyCurrentInput),
          items: {},
        };

        const categoryPlanInputs = document.querySelectorAll<HTMLInputElement>(
          `input.collection-category-plan-input[data-year="${year}"]`
        );

        categoryPlanInputs.forEach(planInput => {
          const category = planInput.dataset.category;
          if (category) {
            const currentInput = document.querySelector<HTMLInputElement>(
              `input.collection-category-current-input[data-category="${category}"][data-year="${year}"]`
            );

            if (currentInput) {
              yearlyData[year].items![category] = {
                name: category,
                plan: getNumericValue(planInput),
                actual: getNumericValue(currentInput),
              };
            }
          }
        });
      }
    });

    collectionData.yearly = yearlyData;

    // ===== PROPAGATE VALUES =====
    const propagatedData = propagateCollectionValues(collectionData);

    // ===== CREATE PAYLOAD =====
    const payload: any = {
      __metadata: {
        type: listItemType,
      },
      Title: generateCollectionTitle(project, propagatedData),
      ProjectId: project.Id.toString(),
      ProfitCentre: project.PC || "",
      CollectionData: JSON.stringify(propagatedData),
      Remarks: combinedRemarks,
      UpdatedDate: new Date().toISOString(),
    };

    if (!editingId && isNewProject) {
      const oblValueToSave = getNumericValue(oblLastMonthInput);
      payload.OBL = oblValueToSave.toString();
    }

    if (!editingId) {
      payload.CreatedDate = new Date().toISOString();
    }

    const url = editingId
      ? `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${editingId})`
      : `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const headers: any = {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
    };

    if (editingId) {
      headers["IF-MATCH"] = "*";
      headers["X-HTTP-Method"] = "MERGE";
    }

    const res = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });

    if (res.ok) {
      toast(
        "success",
        editingId ? "Collection Updated!" : "Collection Added!"
      );
      await loadCollections();
      initializePagination(collections);
      hideForm();
      resetForm();
    } else {
      const errorText = await res.text();
      console.error("❌ SharePoint Error Response:", errorText);

      try {
        const errorData = JSON.parse(errorText);
        if (errorData.error && errorData.error.message) {
          toast(
            "error",
            `SharePoint Error: ${errorData.error.message.value}`
          );
        } else {
          toast("error", `Error saving collection: ${res.status}`);
        }
      } catch (parseError) {
        toast("error", `Error saving collection: ${res.status}`);
      }
    }
  } catch (error) {
    console.error("💥 Error saving collection:", error);
    toast("error", "Unexpected error saving collection!");
  } finally {
    // ✅ Always restore the buttons + clear the guard — no path
    //    (success, cancel, validation fail, or thrown error) can leave
    //    them stuck in the "Saving..." state.
    removeLoadingOverlay();
    saveBtns.forEach((btn: HTMLButtonElement) => {
      btn.disabled = false;
      btn.dataset.saving = "false";
      btn.textContent = "💾 Save Collection Plan";
    });
    isSavingCollection = false;
  }
});

const saveFurtherOnly = async (
  collection: CollectionItem,
  project: ProjectItem
): Promise<void> => {
  const saveLoading = createLoadingOverlay("Saving further changes...");
  loaderHost().appendChild(saveLoading);
  try {
    const listItemType = await getListItemType();
    const digest = await getDigest();

    // Start from the SAVED data so nothing else is touched
    const existing: CollectionData = safeJsonParse(collection.CollectionData, {});

    const currentMonthKey = formatMonthKey(new Date());

    // --- main Further + Date for current month ---
    const mainForecastInput = document.querySelector<HTMLInputElement>(
      `input.collection-main-forecast-input[data-month="${currentMonthKey}"]`
    );
    const mainDateInput = document.querySelector<HTMLInputElement>(
      `input.collection-forecast-date-input[data-month="${currentMonthKey}"]`
    );

    let monthData = existing[currentMonthKey] as CollectionMonthData | undefined;
    if (!monthData || !isCollectionMonthData(monthData)) {
      monthData = { plan: 0, actual: 0, monthActual: 0, forecast: 0, date: '', items: {} };
      existing[currentMonthKey] = monthData;
    }
    if (!monthData.items) monthData.items = {};

    if (mainForecastInput) monthData.forecast = getNumericValue(mainForecastInput);
    if (mainDateInput) monthData.date = mainDateInput.value;

    // --- category Further + Date for current month ---
    const categories = ["Normal", "Normal Overdue", "Retention", "Retention Overdue"];
    categories.forEach(categoryKey => {
      const fInput = document.querySelector<HTMLInputElement>(
        `input.collection-category-forecast-input[data-category="${categoryKey}"][data-month="${currentMonthKey}"]`
      );
      const dInput = document.querySelector<HTMLInputElement>(
        `input.collection-category-date-input[data-category="${categoryKey}"][data-month="${currentMonthKey}"]`
      );

      if (!fInput && !dInput) return;

      const existingCat = monthData!.items![categoryKey];
      const cat: CollectionCategory = existingCat
        ? { ...existingCat }
        : { name: categoryKey, plan: 0, actual: 0, forecast: 0, date: '' };

      if (fInput) cat.forecast = getNumericValue(fInput);
      if (dInput) cat.date = dInput.value;

      monthData!.items![categoryKey] = cat;
    });

    // Remarks are also editable while frozen
    const remarksInternalInput = document.querySelector('#remarksInternal') as HTMLTextAreaElement;
    const remarksSprintCellInput = document.querySelector('#remarksSprintCell') as HTMLTextAreaElement;
    const combinedRemarks =
      (remarksInternalInput?.value.trim() || "") + "|||" + (remarksSprintCellInput?.value.trim() || "");

    const payload: any = {
      __metadata: { type: listItemType },
      CollectionData: JSON.stringify(existing),
      Remarks: combinedRemarks,
      UpdatedDate: new Date().toISOString()
    };

    const res = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${collection.Id})`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE"
        },
        body: JSON.stringify(payload)
      }
    );

    if (res.ok) {
      toast("success", "✅ Further & Date changes saved (project stays frozen)");
      await loadCollections();
      initializePagination(collections);
      hideForm();
      resetForm();
    } else {
      const errorText = await res.text();
      console.error("❌ Frozen save failed:", errorText);
      toast("error", `Error saving further changes: ${res.status}`);
    }
  } catch (err) {
    console.error("💥 saveFurtherOnly error:", err);
    toast("error", "Unexpected error saving further changes!");
  } finally {
    removeLoadingOverlay();
  }
};

// Main forecast validation function for form submission
const validateForecastData = (): boolean => {
  let allValid = true;
  const invalidMonths: string[] = [];

  // Current month is at index 36 in the 54-month structure
  const currentMonthKey = allMonths[36];

  if (currentMonthKey) {
    const mainForecastInput = document.querySelector<HTMLInputElement>(
      `input.collection-main-forecast-input[data-month="${currentMonthKey}"]`
    );

    if (mainForecastInput) {
      const mainForecastValue = getNumericValue(mainForecastInput);
      let categoriesForecastTotal = 0;

      // Sum all category forecasts
      const categoryForecastInputs = document.querySelectorAll<HTMLInputElement>(
        `input.collection-category-forecast-input[data-month="${currentMonthKey}"]`
      );
      categoryForecastInputs.forEach(input => {
        categoriesForecastTotal += getNumericValue(input);
      });

      // Validate that categories forecast doesn't exceed main forecast
      if (categoriesForecastTotal > mainForecastValue) {
        const monthDate = new Date(currentMonthKey + "-01");
        invalidMonths.push(`${formatMonthLabel(monthDate)} (Categories forecast: ${categoriesForecastTotal} > Main forecast: ${mainForecastValue})`);
        allValid = false;
      }
    }
  }

  if (!allValid) {
    toast("error", `Category forecasts exceed main forecast in: ${invalidMonths.join(', ')}`);
    return false;
  }

  return true;
};
const updateMainForecastTotal = (monthKey: string): void => {
  const mainForecastInput = document.querySelector<HTMLInputElement>(
    `input.collection-main-forecast-input[data-month="${monthKey}"]`
  );
  const monthActualInput = document.querySelector<HTMLInputElement>(
    `input.collection-month-actual-input[data-month="${monthKey}"]`
  );
  const forecastTotalInput = document.querySelector<HTMLInputElement>(
    `input.collection-forecast-total-input[data-month="${monthKey}"]`
  );

  // ---- 1. Further (main) = sum of category Further ----
  const categoryForecastInputs = document.querySelectorAll<HTMLInputElement>(
    `input.collection-category-forecast-input[data-month="${monthKey}"]`
  );

  let furtherTotal = 0;
  categoryForecastInputs.forEach(input => {
    furtherTotal += getNumericValue(input);
  });

  if (mainForecastInput) {
    if (furtherTotal > 0) {
      mainForecastInput.value = furtherTotal.toString();
      mainForecastInput.classList.remove('collection-input-empty');
      mainForecastInput.classList.add('collection-input-with-value');
    } else {
      mainForecastInput.value = '';
      mainForecastInput.classList.remove('collection-input-with-value');
      mainForecastInput.classList.add('collection-input-empty');
    }
  }

  // ---- 2. Forecast = Actual + Further ----
  const actualValue = monthActualInput ? getNumericValue(monthActualInput) : 0;
  const forecastValue = actualValue + furtherTotal;

  if (forecastTotalInput) {
    if (forecastValue > 0) {
      forecastTotalInput.value = forecastValue.toString();
      forecastTotalInput.classList.remove('collection-input-empty');
      forecastTotalInput.classList.add('collection-input-with-value');
    } else {
      forecastTotalInput.value = '';
      forecastTotalInput.classList.remove('collection-input-with-value');
      forecastTotalInput.classList.add('collection-input-empty');
    }
  }
};
// ========== FIXED SETUP FORECAST VALIDATION ==========

const setupForecastValidation = (): void => {

  const forecastInputs = document.querySelectorAll<HTMLInputElement>(
    'input.collection-category-forecast-input'
  );

  forecastInputs.forEach(input => {

    // Category Forecast = editable
    input.removeAttribute('readonly');

    input.style.backgroundColor = '#ffffff';
    input.style.cursor = 'text';
    input.style.pointerEvents = 'auto';

    // Remove old listeners
    input.removeEventListener('input', handleForecastCategoryChange);

    // Add new listener
    input.addEventListener('input', handleForecastCategoryChange);

    // Allow only numbers
    allowOnlyNumbers(input);

    // Store old value
    input.setAttribute(
      'data-old-value',
      getNumericValue(input).toString()
    );
  });

  // Update Forecast total for every month
  const months = new Set<string>();

  forecastInputs.forEach(input => {
    if (input.dataset.month) {
      months.add(input.dataset.month);
    }
  });

  months.forEach(month => {
    updateMainForecastTotal(month);
  });
};


const handleForecastCategoryChange = function(this: HTMLInputElement) {
  const month = this.dataset.month;

  if (!month) return;

  // Update input visual state
  const currentValue = getNumericValue(this);

  if (currentValue === 0) {
    this.classList.remove('collection-input-with-value');
    this.classList.add('collection-input-empty');
  } else {
    this.classList.remove('collection-input-empty');
    this.classList.add('collection-input-with-value');
  }

  // Store current value
  this.setAttribute('data-old-value', currentValue.toString());

  // Calculate total of all category forecast inputs
  updateMainForecastTotal(month);
};

// Store old value helper

const storeOldValue = function(this: HTMLInputElement) {
  this.setAttribute('data-old-value', getNumericValue(this).toString());
};

const debugDataTypes = (data: any, path: string = ''): void => {
  if (typeof data === 'object' && data !== null) {
    Object.keys(data).forEach(key => {
      const value = data[key];
      const currentPath = path ? `${path}.${key}` : key;
      if (typeof value === 'object' && value !== null) {
        debugDataTypes(value, currentPath);
      }
    });
  }
};

 const initialize = async (): Promise<void> => {
  const loading = createLoadingOverlay("Loading collection data...");
  loaderHost().appendChild(loading);
  try {
    // Initialize shared authentication
    updateLoadingMessage(loading, "Initializing permissions...");
    await authManager.initialize({ siteUrl });

    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
    console.log(`👤 User: ${getCurrentUserFullName()}`);
    console.log(`📋 User Groups:`, getUserGroups());

    // Load projects first
    updateLoadingMessage(loading, "Loading projects...");
    await loadProjects();

    // Then load collections
    updateLoadingMessage(loading, "Loading collections...");
    await loadCollections();

    console.log(`📊 Loaded ${collections.length} collections`);

    // 🔐 AUTO-SYNC: compare Project List with Collection List and add any
    // missing project (skeleton JSON: 36 historical + current/next 2 +
    // next 16 months + 4-year yearly). Only runs for roles with create
    // rights and only for projects inside that role's visible scope.
    updateLoadingMessage(loading, "Syncing missing projects...");
    await syncMissingProjectsToCollectionList();

    // Show the manual Sync button only to roles that can create rows
    const syncBar = document.querySelector('#collectionSyncBar') as HTMLElement | null;
    if (syncBar) {
      syncBar.style.display = canCreateRecord() ? 'flex' : 'none';
    }

    // 🔐 RBAC AUTO-PROGRESSION (monthly roll-over):
    //   - Roles with edit rights (canEditOwn: Admin, GH Heads, PM Managers,
    //     Engg Lead, Engg Sub1) → perform the real SharePoint bulk update.
    //   - Read-only roles (SC Team, Controlling, Engg Sub2, Engg Mgr) → the
    //     same roll-over is applied IN MEMORY only, so their screen shows the
    //     current month structure WITHOUT writing to SharePoint.
    if (collections.length > 0) {
        if (canEditOwn()) {
            updateLoadingMessage(loading, "Applying monthly progression...");
            await applyAutoProgression();
            // Reload after progression
            await loadCollections();
        } else {
            const viewerProgressed = collections.map((c: CollectionItem) => {
                const data = safeJsonParse(c.CollectionData, {});
                if (!checkIfNeedsProgression(data)) return c;
                return { ...c, CollectionData: JSON.stringify(propagateCollectionValues(data)) };
            });
            collections = viewerProgressed;
            console.log('👁️ Read-only viewer: progression applied locally for display only (no SharePoint writes)');
        }
    }

    // IMPORTANT: Force page size to 25 on initial load
    pageSize = 25;
    if (pageSizeSelect) {
        pageSizeSelect.value = "25";
        pageSize = parseInt(pageSizeSelect.value) || 25;
        console.log(`📊 Initial page size set to: ${pageSize}`);
    }

    // Initialize filters
    initializeFilters();

    // Initialize pagination
    filteredData = [...collections];
    currentPage = 1;
    totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));

    // Render the table header
    renderTableHeader();

    updatePaginationControls();
    renderCurrentPage();
    setupPaginationEventListeners();

    // CRITICAL FIX: Set dropdown value again
    if (pageSizeSelect) {
        pageSizeSelect.value = "25";
        pageSize = 25;
        console.log(`📊 Final page size dropdown value: ${pageSizeSelect.value}`);
    }

    // Update UI based on permissions
    updateUIBasedOnPermissions({
      importBtn: '.collection-import-actual-btn',
      exportBtn: '.collection-export-btn',
      addBtn: '.collection-show-form-btn',
      saveBtn: '.collection-btn-primary[type="submit"]',
      freezeCheckboxes: '.collection-freeze-checkbox'
    });

    // 🔐 RBAC: hide the historical export button for non-readers too
    // (updateUIBasedOnPermissions only covers the main export button).
    if (!canRead()) {
      const historicalExportBtn = document.querySelector('.collection-export-historical-btn') as HTMLElement | null;
      if (historicalExportBtn) historicalExportBtn.style.display = 'none';
    }

    console.log("✅ Application initialized successfully");

  } catch (error) {
    console.error("Initialization error:", error);
    toast("error", "Failed to initialize application");
  } finally {
    removeLoadingOverlay();
  }
};

      showFormBtn.addEventListener("click", () => {
        resetForm();
        showForm(false);
      });

      // ============================================================
// ✅ BACK — SweetAlert confirmation, same look as revenue.ts
//    • Only one prompt: Leave (Yes, Go Back / Stay)
//    • Only shown when the form is actually open
//    • z-index bumped so the loader cannot cover the dialog
// ============================================================
const handleCollectionBackAction = async (e: Event): Promise<void> => {
  e.preventDefault();
  e.stopPropagation();

  if (isLeavingCollectionForm) {
    console.log("⛔ back blocked: already leaving");
    return;
  }

  // Nothing to confirm when the form is already hidden
  const formVisible = formSection && formSection.style.display !== "none";
  if (!formVisible) {
    hideForm();
    resetForm();
    return;
  }

  isLeavingCollectionForm = true;
  try {
    const confirmResult = await Swal.fire({
      title: "Leave this page?",
      html:
        "Any <strong>unsaved changes</strong> will be lost.<br>" +
        "Do you want to go back to the collection list?",
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
      didOpen: bumpSwalZIndex,
    });

    if (!confirmResult.isConfirmed) {
      console.log("🔵 back cancelled by user");
      return;
    }

    hideForm();
    resetForm();
  } catch (err) {
    console.error("Error while leaving the Collection form:", err);
    const message = err instanceof Error ? err.message : String(err);
    toast("error", `Unexpected error: ${message}`);
  } finally {
    isLeavingCollectionForm = false;
  }
};
      backBtn.addEventListener("click", handleCollectionBackAction);

      projectSelect.addEventListener("change", (e) => {
  const projectId = parseInt((e.target as HTMLSelectElement).value);

  if (!projectId) {
    resetForm();
    return;
  }

  if (editingId) {
    toast("warning", "Cannot change project in edit mode");
    const collection = collections.find(f => f.Id === editingId);
    if (collection?.Project) {
      projectSelect.value = String(collection.Project.Id);
    }
    return;
  }

  const project = projects.find(p => p.Id === projectId);
  if (!project) return;

  const existingCollection = collections.find(f => f.Project?.Id === projectId);

  const formLoading = createLoadingOverlay("Loading collection form...");
  loaderHost().appendChild(formLoading);

  if (existingCollection) {
    // 🔐 RBAC: opening an existing collection is allowed for every role that
    // may VIEW it (READ_ALL roles see all, others only their OWN) — editing
    // rights are enforced inside by canEditProject().
    editingId = existingCollection.Id;
    showForm(true);
    onProjectChange(projectId, existingCollection).finally(() => removeLoadingOverlay());
    // toast("info", `Editing existing collection for ${project.Title}`);
  } else {
    // 🔐 RBAC: creating a NEW collection record requires create rights
    // (FULL_CONTROL / CONTRIBUTE only — Admin, GH Heads, PM Managers).
    if (!canCreateRecord()) {
      removeLoadingOverlay();
      toast("error", "❌ You don't have permission to create collection records");
      projectSelect.value = "";
      resetForm();
      return;
    }
    onProjectChange(projectId).finally(() => removeLoadingOverlay());
  }
});


const validateYearlyCategoryTotals = (changedInput?: HTMLInputElement): boolean => {
  let allValid = true;

  const yearlyInputs = document.querySelectorAll<HTMLInputElement>('.collection-yearly-current-input');
  const yearlyYears = new Set<string>();

  yearlyInputs.forEach(input => {
    const year = input.dataset.year;
    if (year) yearlyYears.add(year);
  });

  Array.from(yearlyYears).forEach(year => {
    const yearValid = validateSingleYearCategories(year, changedInput);
    if (!yearValid) {
      allValid = false;
    }
  });

  return allValid;
};

const validateSingleYearCategories = (year: string, changedInput?: HTMLInputElement): boolean => {
  const yearlyCurrentInput = document.querySelector<HTMLInputElement>(`input.collection-yearly-current-input[data-year="${year}"]`);
  const categoryCurrentInputs = document.querySelectorAll<HTMLInputElement>(`input.collection-category-current-input[data-year="${year}"]`);

  if (!yearlyCurrentInput) return true;

  const yearlyCurrentValue = getNumericValue(yearlyCurrentInput);
  let categoriesTotal = 0;

  categoryCurrentInputs.forEach(input => {
    categoriesTotal += getNumericValue(input);
  });

  let validationDisplay = document.querySelector(`.collection-yearly-total-validation[data-year="${year}"]`) as HTMLDivElement;
  if (!validationDisplay) {
    validationDisplay = document.createElement('div');
    validationDisplay.className = 'collection-yearly-total-validation';
    validationDisplay.setAttribute('data-year', year);

    const yearBlock = yearlyCurrentInput.closest('.collection-yearly-box');
    if (yearBlock) {
      const existingValidation = yearBlock.querySelector('.collection-yearly-total-validation');
      if (existingValidation) {
        existingValidation.remove();
      }
      yearBlock.appendChild(validationDisplay);
    }
  }

  // Simplified validation - just show match or mismatch
  if (categoriesTotal > yearlyCurrentValue) {
    blockIncreasingYearCategoryInputs(year, changedInput);
    return false;
  } else if (categoriesTotal < yearlyCurrentValue) {
    unblockAllYearCategoryInputs(year);
    return true;
  } else {
    unblockAllYearCategoryInputs(year);
    return true;
  }
};

const blockIncreasingYearCategoryInputs = (year: string, allowedInput?: HTMLInputElement): void => {
  const categoryInputs = document.querySelectorAll<HTMLInputElement>(`input.collection-category-current-input[data-year="${year}"]`);

  categoryInputs.forEach(input => {
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

const unblockAllYearCategoryInputs = (year: string): void => {
  const categoryInputs = document.querySelectorAll<HTMLInputElement>(`input.collection-category-current-input[data-year="${year}"]`);

  categoryInputs.forEach(input => {
    input.style.backgroundColor = '';
    input.style.cursor = '';
    input.removeAttribute('data-block-increase');
    input.removeAttribute('readonly');
  });
};

      // Pagination event listeners
      firstPageBtn.addEventListener("click", () => goToPage(1));
      prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
      nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
      lastPageBtn.addEventListener("click", () => goToPage(totalPages));

      // Filter popup + chips + Apply pill + Reset Filter: ONE set of bindings
      // (inventory style - the table only re-filters when Apply is pressed).
      setupFilterEventListeners();

      // Action button event listeners
      exportBtn.addEventListener("click", exportToExcelWithFullStyle);
      importActualBtn.addEventListener("click", importActualData);
      const syncProjectsBtn = document.querySelector("#collectionSyncProjectsBtn") as HTMLButtonElement;
      if (syncProjectsBtn) {
        syncProjectsBtn.addEventListener("click", async () => {
          const syncLoading = createLoadingOverlay("Syncing projects...");
          loaderHost().appendChild(syncLoading);
          try {
            await syncMissingProjectsToCollectionList();
          } finally {
            removeLoadingOverlay();
          }
        });
      }
      const exportHistoricalBtn = document.querySelector(".collection-export-historical-btn") as HTMLButtonElement;

// Add event listener for historical export
if (exportHistoricalBtn) {
  exportHistoricalBtn.addEventListener("click", exportHistoricalToExcel);
}
      window.addEventListener('resize', () => {
        clearTimeout((window as any)._stickyResizeTimer);
        (window as any)._stickyResizeTimer = setTimeout(applyStickyHeaderOffsets, 150);
      });
      // Initialize application
      initialize();
    }, 100); 
  } 
}; 

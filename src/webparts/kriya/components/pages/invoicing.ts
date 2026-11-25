declare const _spPageContextInfo: any;
import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css"; 
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/invoicing.css";

import { 
  authManager,
  hasFullControl,
  canContribute,
  canRead,
  canFreeze,
  isGHHead,
  getUserGroups,
  getCurrentUserFullName,
  getPermissionLevel,
  canEditProject,
  canCreateRecord,
  canReadAllProjects,
  isPMManager,
  isEngineeringUser,
  // canViewForm,
  getAssignedProjects,
  // getAssignedProjectValues,
  updateUIBasedOnPermissions,
  getDigest as getAuthDigest,
  // getListEntityType
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

interface InvoicingForecastItem {
  Id: number;
  Title: string;
  Project?: { 
    Id: number;
    Title: string;
  };
  ProfitCentre?: string;
  OBL?: number;
  ForecastData?: string;
  InvoiceActual?: string;
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
}

interface ForecastMonthData {
    plan: number;
    actual: number;
    forecast?: number;
    date?: string;
    items?: {
        [itemNumber: string]: ForecastItem;
    };
}

interface YearlyForecastData {
    [year: string]: {
        plan: number;
        actual: number;
    };
}

interface ForecastData {
    [monthKey: string]: ForecastMonthData | YearlyForecastData | undefined;
    yearly?: YearlyForecastData;
}

export default {
  content: `
<div class="invoicing-page fade-in">
  <section class="table-section invoicing-card">
    <div class="invoicing-page-header">
      <h1>Invoicing Plan 5QRF</h1>
    </div>

    <!-- ✅ TOOLBAR ABOVE THE TABLE - the Filters popup trigger + its Apply pill (left)
         and the two Excel exports (right) live here (NOT next to the page title any more).
         There is no Reset button in the toolbar: the popup's Clear button resets every
         filter. The seven filter dropdowns moved inside the Filters popup below, keeping
         their original ids (#filterProject, #filterYear, ...) so initializeFilters(),
         filterTable() and the exports keep working unchanged. -->
    <div class="invoicing-toolbar">
      <div class="invoicing-toolbar__left">
        <button type="button" class="invoicing-btn-secondary invoicing-filter-trigger" id="filterTriggerInvoicing"
                aria-haspopup="dialog" aria-expanded="false" title="Filter Invoicing">
          <svg class="invoicing-filter-funnel" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8v5l-4 2v-7L3 5z"/></svg>
          <span class="invoicing-filter-label">Filters</span>
          <span class="invoicing-filter-count" id="filterCountInvoicing">0</span>
          <svg class="invoicing-filter-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>
        </button>

        <!-- compact Apply pill - shown only while the popup holds changes that
             have not been applied yet -->
        <button type="button" class="invoicing-btn-primary invoicing-filter-applypill" id="floatingFilterApplyInvoicing">Apply</button>
      </div>

      <div class="invoicing-toolbar__right">
        <button type="button" class="invoicing-btn-secondary invoicing-export-btn">📊 Download Excel</button>
        <button type="button" class="invoicing-btn-secondary invoicing-export-details-btn">📋 Details Download Excel</button>
      </div>
    </div>

    <!-- chips summarising the filters that are currently applied -->
   <!-- <div class="invoicing-filter-summary" id="filterSummaryInvoicing" style="display:none;"></div> -->

    <!-- ✅ CHANGE 6: same scroll structure as inventory (vertical + horizontal, sticky totals/header) -->
    <div class="table-container-inv">
      <div class="invoicing-table-scroll-wrapper">
        <table class="month-columns-table-inv">
          <thead id="totalsHeader" class="sticky-totals-header"></thead>
          <thead id="monthColumnsHeader" class="sticky-main-header"></thead>
          <tbody class="invoicing-table-body-inv">
            <!-- ✅ CHANGE 1: Static loading row - visible from first paint until the JS render
                 replaces it, so the table never looks blank while data loads.
                 ✅ Now shows the indeterminate page-loader bar (grey track + indigo bar, no text). -->
            <tr class="invoicing-table-loading-row">
              <td colspan="25">
                <div class="invoicing-page-loader">
                  <div class="invoicing-page-loader__track">
                    <div class="invoicing-page-loader__bar"></div>
                  </div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- ✅ PAGINATION - one flat bar under the table (mock-up layout)
           LEFT  = Items per page: [n]   Showing 1 - 25 of 60 items
           RIGHT = «  ‹  [page tiles] of N  ›  »
           Every id below is unchanged (#pageSizeSelect, #startItem, #endItem,
           #totalItems, #pageNumbers, #totalPages, #first/#prev/#next/#lastPage),
           so updatePaginationControls / updatePageNumbers / changePageSize /
           goToPage keep working. The "Page X of Y" pill is gone - it is now the
           "of N" after the tiles. -->
      <div class="invoicing-pagination-container">
        <!-- LEFT: items per page + showing info -->
        <div class="invoicing-pagination-left">
          <div class="invoicing-page-size-selector invoicing-pg-size">
            <label for="pageSizeSelect">Items per page:</label>
            <select id="pageSizeSelect" class="invoicing-page-size-select">
              <option value="5">5</option>
              <option value="10">10</option>
              <option value="25" selected>25</option>
              <option value="50">50</option>
              <option value="100">100</option>
            </select>
          </div>

          <div class="invoicing-pagination-info invoicing-pg-showing">
            Showing <span id="startItem">0</span> - <span id="endItem">0</span> of <span id="totalItems">0</span> items
          </div>
        </div>

        <!-- RIGHT: «  ‹  [page tiles] of N  ›  » -->
        <div class="invoicing-pagination-right">
          <button class="invoicing-pagination-btn invoicing-pg-icon" id="firstPage" title="First page">«</button>
          <button class="invoicing-pagination-btn invoicing-pg-icon" id="prevPage" title="Previous page">‹</button>

          <div class="invoicing-pagination-pages" id="pageNumbers"></div>

          <span class="invoicing-pg-of">of <span id="totalPages">1</span></span>

          <button class="invoicing-pagination-btn invoicing-pg-icon" id="nextPage" title="Next page">›</button>
          <button class="invoicing-pagination-btn invoicing-pg-icon" id="lastPage" title="Last page">»</button>
        </div>
      </div>
    </div>
  </section>

  <!-- =====================================================================
       ✅ FILTER POPUP (header · fields · Cancel / Clear / Apply)
       The seven selects keep their ORIGINAL ids (#filterProject,
       #filterCustomer, #filterProjectManager, #filterGroupHead,
       #filterProjectReference, #filterYear, #filterMonth) so
       initializeFilters(), filterTable(), clearFilters() and both Excel
       exports keep working.
       Footer: Cancel (revert the draft) · Clear (RESET every filter, kept
       inside the popup) · Apply.
       ===================================================================== -->
  <div class="invoicing-filter-overlay" id="invoicingFilterOverlay" role="dialog" aria-modal="true" aria-labelledby="invoicingFilterTitle">
    <div class="invoicing-filter-modal">
      <div class="invoicing-filter-modal__head">
        <h2 id="invoicingFilterTitle">Filters</h2>
        <button type="button" class="invoicing-filter-modal__close" id="filterModalCloseInvoicing" title="Close" aria-label="Close">✕</button>
      </div>

      <div class="invoicing-filter-modal__body">
        <div class="invoicing-filter-grid">
          <div class="invoicing-filter-field invoicing-filter-field--wide">
            <label for="filterProject">Project</label>
            <select id="filterProject" data-filter="project"><option value="">All Projects</option></select>
          </div>
          <div class="invoicing-filter-field">
            <label for="filterCustomer">Customer</label>
            <select id="filterCustomer" data-filter="customer"><option value="">All Customers</option></select>
          </div>
          <div class="invoicing-filter-field">
            <label for="filterProjectManager">Project Manager</label>
            <select id="filterProjectManager" data-filter="manager"><option value="">All Managers</option></select>
          </div>
          <div class="invoicing-filter-field">
            <label for="filterGroupHead">Group Head</label>
            <select id="filterGroupHead" data-filter="groupHead"><option value="">All Group Heads</option></select>
          </div>
          <div class="invoicing-filter-field">
            <label for="filterProjectReference">Project Reference</label>
            <select id="filterProjectReference" data-filter="reference"><option value="">All References</option></select>
          </div>
          <div class="invoicing-filter-field">
            <label for="filterYear">Year</label>
            <select id="filterYear" data-filter="year"><option value="">All Years</option></select>
          </div>
          <div class="invoicing-filter-field">
            <label for="filterMonth">Month</label>
            <select id="filterMonth" data-filter="month"><option value="">All Months</option></select>
          </div>
        </div>
      </div>

      <div class="invoicing-filter-modal__foot">
        <button type="button" class="invoicing-filter-btn-text" id="filterCancelInvoicing">Cancel</button>
        <button type="button" class="invoicing-filter-btn-clear" id="filterClearModalInvoicing" title="Reset all filters" disabled>Clear</button>
        <button type="button" class="invoicing-filter-btn-apply" id="filterApplyModalInvoicing" disabled>Apply</button>
      </div>
    </div>
  </div>

 <section class="invoicing-form-section invoicing-card" style="display:none;">
   <form id="invoicingForm"> 
  <div class="invoicing-page-header">
  <h1 id="formTitle">Invoicing Plan 5QRF</h1>
  
      <!-- ✅ Added wrapper div here -->
      <div class="invoicing-header-actions">
        <button type="submit" form="invoicingForm" class="invoicing-btn-primary invoicing-btn-save-top" title="Save Invoicing Forecast">💾 Save</button>
        <button type="button" class="invoicing-btn-secondary back-to-table">⬅ Back</button>
      </div>
    </div>

      <!-- Project Selection (hidden in edit mode) -->
      <div class="invoicing-form-row invoicing-hidden-section">
        <div class="invoicing-form-group invoicing-project-select-container">
          <label>Project</label>
          <select id="projectSelect" required></select>
        </div>
      </div>

      <!-- MAIN WRAPPER - 2 Column Layout (Matching Revenue/SEIBIES) -->
      <div class="project-main-wrapper-invoicing">

        <!-- LEFT SIDE - Project Details (6 Rows) -->
        <div class="project-left-side-invoicing" id="projectDetailsInvoicing">
          <div class="detail-item-invoicing">
            <span class="detail-label-invoicing">Project:</span>
            <span class="detail-value-invoicing" id="projectNameDisplayInvoicing">-</span>
          </div>
          <div class="detail-item-invoicing">
            <span class="detail-label-invoicing">Project Ref:</span>
            <span class="detail-value-invoicing" id="projectRefInvoicing">-</span>
          </div>
          <div class="detail-item-invoicing">
            <span class="detail-label-invoicing">Profit Center:</span>
            <span class="detail-value-invoicing" id="profitCenterInvoicing">-</span>
          </div>
          <div class="detail-item-invoicing">
            <span class="detail-label-invoicing">Project Manager:</span>
            <span class="detail-value-invoicing" id="projectManagerInvoicing">-</span>
          </div>
          <div class="detail-item-invoicing">
            <span class="detail-label-invoicing">Group Head:</span>
            <span class="detail-value-invoicing" id="groupHeadInvoicing">-</span>
          </div>
          <div class="detail-item-invoicing">
            <span class="detail-label-invoicing">Customer:</span>
            <span class="detail-value-invoicing" id="customerInvoicing">-</span>
          </div>
        </div>

        <!-- RIGHT SIDE - OBL Section (No Remarks) -->
        <div class="project-right-side-invoicing">
          <!-- OBL Section -->
          <div class="obl-section-container-invoicing invoicing-hidden-section" id="oblContainerInvoicing">
            <div class="obl-content-invoicing">
              <div class="obl-left-invoicing">
                <div class="obl-label-group-invoicing">
                  <label>Balance to Invoice </label>
                </div>
                <div class="obl-value-display-invoicing">
                  <input type="text" id="oblLastMonthInvoicing" class="obl-input-invoicing number-input" placeholder="Enter Invoicing" readonly />
                </div>
              </div>
              <div class="obl-right-invoicing">
                <div class="validation-summary-invoicing">
                  <div class="validation-item-invoicing">
                    <span class="validation-label-invoicing">Total Current Value:</span>
                    <span class="validation-value-invoicing" id="totalCurrentValueInvoicing">0</span>
                  </div>
                  <div class="validation-item-invoicing">
                    <span class="validation-label-invoicing">Remaining Balance:</span>
                    <span class="validation-value-invoicing" id="remainingValueInvoicing">0</span>
                  </div>
                  <div class="validation-line-invoicing">
                    <div class="line-fill-invoicing" id="oblLineFillInvoicing"></div>
                  </div>
                </div>
              </div>
            </div>
            <div class="validation-message-invoicing" id="oblValidationMessageInvoicing">
              ✅ All values are within INVOICING limit
            </div>
          </div>
        </div>

      </div>
      
      <!-- Historical Months (Last 3 Months - Readonly) -->
      <div class="invoicing-forecast-section invoicing-hidden-section" id="historicalForecastSection">
        <div class="invoicing-historical-toggle-header">
          <h3>📊 Historical Data (Last 3 Months - Readonly)</h3>
          <!-- ✅ CHANGE 2: Arrow on top + "Show / Hide" label below (same as inventory) -->
          <button type="button" id="historicalToggleBtn" class="invoicing-toggle-arrow-btn" aria-expanded="false" title="Show historical data">
            <span class="invoicing-toggle-arrow">▼</span>
            <span class="invoicing-toggle-label">Show</span>
          </button>
        </div>
        <div id="historicalContainer" class="invoicing-historical-container" style="display:none;"></div>
      </div>
      
      <!-- Current + Next 2 Months - Hidden by default -->
      <div class="invoicing-forecast-section invoicing-hidden-section" id="detailedForecastSection">
      <h3>Detailed Forecast (Current + Next 2 Months)</h3>
      <div id="firstRowContainer" class="invoicing-first-row-container">
         
      </div>
    </div>
      
      <!-- Next 15 Months - Hidden by default -->
      <div class="invoicing-forecast-section invoicing-hidden-section" id="longtermForecastSection">
        <h3>Long-term Forecast (Next 15 Months)</h3>
        <div id="remainingContainer" class="invoicing-remaining-container"></div>
      </div>

      <!-- 5-Year Forecast Section - Hidden by default -->
      <div class="invoicing-forecast-section invoicing-hidden-section" id="yearlyForecastSection">
        <h3>5-Year Forecast</h3>
        <div id="yearlyContainer" class="yearly-container-inv"></div>
      </div>
      
      <div class="invoicing-form-actions">
        <button type="submit" class="invoicing-btn-primary invoicing-btn-save-top">💾 Save</button>
         <button type="button" class="invoicing-btn-secondary back-to-table">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,



  script: () => {
    setTimeout(async () => {
      const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/Kriya";
      const listName = "InvoicingPlanning";
      const projectListName = "Project List";

      // DOM Elements
      const tableBody = document.querySelector(".invoicing-table-body-inv") as HTMLElement;
      const monthColumnsHeader = document.querySelector("#monthColumnsHeader") as HTMLElement;
      const formSection = document.querySelector(".invoicing-form-section") as HTMLElement;
      const tableSection = document.querySelector(".table-section") as HTMLElement;
      const projectSelect = document.querySelector("#projectSelect") as HTMLSelectElement;
      const historicalContainer = document.querySelector("#historicalContainer") as HTMLElement;
      const historicalToggleBtn = document.querySelector("#historicalToggleBtn") as HTMLButtonElement;
      const firstRowContainer = document.querySelector("#firstRowContainer") as HTMLElement;
      const remainingContainer = document.querySelector("#remainingContainer") as HTMLElement;
      const yearlyContainer = document.querySelector("#yearlyContainer") as HTMLElement;
      const backBtn = document.querySelector(".back-to-table") as HTMLButtonElement;
      const form = document.querySelector("#invoicingForm") as HTMLFormElement;
      const exportBtn = document.querySelector(".invoicing-export-btn") as HTMLButtonElement;
      const oblLastMonthInput = document.querySelector("#oblLastMonthInvoicing") as HTMLInputElement;
      const totalCurrentValueSpan = document.querySelector("#totalCurrentValueInvoicing") as HTMLSpanElement;
      const remainingValueSpan = document.querySelector("#remainingValueInvoicing") as HTMLSpanElement;
      const oblLineFill = document.querySelector("#oblLineFillInvoicing") as HTMLDivElement;
      const oblValidationMessage = document.querySelector("#oblValidationMessageInvoicing") as HTMLDivElement;
      const formTitle = document.querySelector("#formTitle") as HTMLElement;

      // Project Details Display Elements
      const projectNameDisplay = document.querySelector("#projectNameDisplayInvoicing") as HTMLElement;
      const projectRefDisplay = document.querySelector("#projectRefInvoicing") as HTMLElement;
      const profitCenterDisplay = document.querySelector("#profitCenterInvoicing") as HTMLElement;
      const projectManagerDisplay = document.querySelector("#projectManagerInvoicing") as HTMLElement;
      const groupHeadDisplay = document.querySelector("#groupHeadInvoicing") as HTMLElement;
      const customerDisplay = document.querySelector("#customerInvoicing") as HTMLElement;

      // Filter Elements
      const filterYear = document.querySelector("#filterYear") as HTMLSelectElement;
      const filterMonth = document.querySelector("#filterMonth") as HTMLSelectElement;
      const filterProject = document.querySelector("#filterProject") as HTMLSelectElement;
      const filterCustomer = document.querySelector("#filterCustomer") as HTMLSelectElement;
      const filterProjectManager = document.querySelector("#filterProjectManager") as HTMLSelectElement;
      const filterGroupHead = document.querySelector("#filterGroupHead") as HTMLSelectElement;
      const filterProjectReference = document.querySelector("#filterProjectReference") as HTMLSelectElement;

      // The Reset/clear action now lives INSIDE the Filters popup (Clear button),
      // so the old toolbar button - and its reference - are gone.

      const exportDetailsBtn = document.querySelector(".invoicing-export-details-btn") as HTMLButtonElement;

      // Pagination Elements
      const firstPageBtn = document.querySelector("#firstPage") as HTMLButtonElement;
      const prevPageBtn = document.querySelector("#prevPage") as HTMLButtonElement;
      const nextPageBtn = document.querySelector("#nextPage") as HTMLButtonElement;
      const lastPageBtn = document.querySelector("#lastPage") as HTMLButtonElement;
      // kept for backwards compatibility - the "Page X of Y" pill was replaced by
      // the "« ‹ [page tiles] of N › »" layout, so this node no longer exists
      const currentPageSpan = document.querySelector("#currentPage") as HTMLSpanElement | null;
      const totalPagesSpan = document.querySelector("#totalPages") as HTMLSpanElement;
      const pageNumbersDiv = document.querySelector("#pageNumbers") as HTMLDivElement;
      const startItemSpan = document.querySelector("#startItem") as HTMLSpanElement;
      const endItemSpan = document.querySelector("#endItem") as HTMLSpanElement;
      const totalItemsSpan = document.querySelector("#totalItems") as HTMLSpanElement;

      let projects: ProjectItem[] = [];
      let forecasts: InvoicingForecastItem[] = [];
      let editingId: number | null = null;
      let allMonths: string[] = [];
      
      let currentPage = 1;
      let pageSize = 25; // ✅ default 25 per page (same as DCS / Procurement / SEIBIES / Inventory)
      let totalPages = 1;
      let filteredData: InvoicingForecastItem[] = [];
   // ============================================================
// ✅ SHARED HELPERS for the Save / Back confirmation dialogs
//    (same recipe as procurement.ts / dci.ts / revenue.ts / project.ts / inventory.ts)
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
let isSavingInvoicing = false;
let isLeavingInvoicingForm = false;

// Dialogs must render ABOVE every overlay of this module
const bumpSwalZIndex = (): void => {
  const c = document.querySelector(".swal2-container") as HTMLElement | null;
  if (c) c.style.zIndex = "2147483647";
};
      // Add page size selector element
      const pageSizeSelect = document.querySelector("#pageSizeSelect") as HTMLSelectElement;

      // Add changePageSize function
      const changePageSize = (newSize: number): void => {
        if (newSize === pageSize) return;

        // ✅ Keep the SAME first row visible after the switch: read the offset while
        //    pageSize still has its OLD value. (The previous version read pageSize
        //    after assigning it, so the maths always cancelled out and the table
        //    jumped back to an earlier page.)
        const firstVisibleIndex = (currentPage - 1) * pageSize;

        pageSize = newSize;

        calculateTotalPages();

        currentPage = Math.floor(firstVisibleIndex / pageSize) + 1;
        if (currentPage > totalPages) currentPage = totalPages || 1;
        if (currentPage < 1) currentPage = 1;

        updatePaginationControls();
        renderCurrentPage();

        toast("info", `Showing ${pageSize} records per page`);
      };

      // Add event listener for page size selector
      if (pageSizeSelect) {
        pageSizeSelect.addEventListener("change", (e) => {
          const newSize = parseInt((e.target as HTMLSelectElement).value, 10) || 25;
          changePageSize(newSize);
        });
      }

      // ========== UTILITY FUNCTIONS ==========

      const toast = (icon: any, title: string) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });

      /**
       * ✅ PAGE LOADER - bar only (same as the Inventory / SEIBIES modules).
       *    The overlay keeps this exact API - createLoadingOverlay(message) returns the
       *    element to append, removeLoadingOverlay() cleans up - and the .loading-overlay
       *    lifecycle class, so every existing call site keeps working unchanged.
       *    The circular spinner + "Loading ..." text are gone: the grey track + indigo bar
       *    (.invoicing-page-loader, styled in styles/invoicing.css) are the only thing on
       *    screen, and the step message is logged to the console instead.
       */
      const createLoadingOverlay = (message = "Processing..."): HTMLElement => {
        removeLoadingOverlay(); // never stack two loaders
        if (message) console.log(`⏳ ${message}`);

        const overlay = document.createElement("div");
        overlay.className = "loading-overlay invoicing-loading-overlay";
        overlay.innerHTML = `
          <div class="invoicing-loader-panel">
            <div class="invoicing-page-loader">
              <div class="invoicing-page-loader__track">
                <div class="invoicing-page-loader__bar"></div>
              </div>
            </div>
          </div>`;
        return overlay;
      };

      const removeLoadingOverlay = (): void => {
        document.querySelectorAll(".loading-overlay, .invoicing-loading-overlay").forEach(el => el.remove());
      };


      // ✅ EVENT DELEGATION - Attach to document body
document.body.addEventListener("click", (e: Event) => {
  const target = e.target as HTMLElement;
  
  // Check if the clicked element (or its parent) has the class 'back-to-table'
  const backBtn = target.closest(".back-to-table");
  
  if (backBtn) {
    e.preventDefault(); // 🛑 Stop default behavior
    console.log("🔘 Back button clicked via delegation!");

    try {
      // ✅ 1. Force UI changes directly
      const formSec = document.querySelector(".invoicing-form-section") as HTMLElement;
      const tableSec = document.querySelector(".table-section") as HTMLElement;
      
      if (formSec) formSec.style.display = "none";
      if (tableSec) tableSec.style.display = "block";

      // ✅ 2. Try your custom functions safely
      if (typeof hideForm === 'function') hideForm();
      if (typeof reloadInvoicingData === 'function') reloadInvoicingData();
      
    } catch (err) {
      console.error("⚠️ Error while navigating back:", err);
    }
  }
});

      /** mount the loader on the VISIBLE card (table section or form section) —
          the table card while data loads, the form card while saving */
      const loaderHost = (): HTMLElement => {
        const sections = Array.from(document.querySelectorAll<HTMLElement>('.table-section, .invoicing-form-section'));
        const visible = sections.find(s => s.offsetParent !== null);
        return visible || (document.querySelector('.table-section') as HTMLElement | null) || document.body;
      };


      const getDigest = async (): Promise<string> => {
        return getAuthDigest(siteUrl);
      };

      // Safe JSON parsing function
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
        
        // Calculate total from monthly data only (exclude yearly)
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

      // Get current and next 2 months for display
      const getDisplayMonths = (): string[] => {
        const now = new Date();
        const months: string[] = [];
        
        // Current month + next 2 months
        for (let i = 0; i < 3; i++) {
          const month = new Date(now.getFullYear(), now.getMonth() + i, 1);
          months.push(formatMonthKey(month));
        }
        
        return months;
      };

      const getAllMonthsForStorage = (): string[] => {
        return generateIdealMonthStructure();
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
            this.className = this.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
          } else {
            this.className = this.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
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

      // Function to get numeric value from text input - 0 is placeholder
      const getNumericValue = (input: HTMLInputElement): number => {
        const value = input.value.trim();
        if (value === '' || value === '0') return 0;
        return parseInt(value) || 0;
      };

      // Function to set numeric value to text input - 0 shows as empty
      const setNumericValue = (input: HTMLInputElement, value: number): void => {
        if (value === 0) {
          input.value = '';
          input.className = input.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
        } else {
          input.value = value.toString();
          input.className = input.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
        }
      };

      // Enhanced OBL validation with toast notifications
      const showOBLErrorToast = (message: string, isWarning: boolean = false): void => {
        const existingToast = document.querySelector('.invoicing-obl-error-toast');
        if (existingToast) existingToast.remove();
        
        const toastEl = document.createElement('div');
        toastEl.className = `invoicing-obl-error-toast ${isWarning ? 'invoicing-warning' : ''}`;
        toastEl.innerHTML = `
          ${message}
          <button class="invoicing-toast-close">&times;</button>
        `;
        
        document.body.appendChild(toastEl);
        
        setTimeout(() => {
          if (toastEl.parentNode) {
            toastEl.remove();
          }
        }, 5000);
        
        toastEl.querySelector('.invoicing-toast-close')?.addEventListener('click', () => {
          toastEl.remove();
        });
      };

      // ========== MONTH STRUCTURE FUNCTIONS ==========
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
        const currentDate = new Date();
         
        // Historical: 3 months before current (readonly)
        for (let i = -3; i < 0; i++) {
          const historicalDate = new Date(currentDate);
          historicalDate.setMonth(historicalDate.getMonth() + i);
          months.push(formatMonthKey(historicalDate));
        }
        
        // Detailed: Current + Next 2 months (editable with items)
        for (let i = 0; i < 3; i++) {
          const detailedDate = new Date(currentDate);
          detailedDate.setMonth(detailedDate.getMonth() + i);
          months.push(formatMonthKey(detailedDate));
        }
        
        // Future: Next 15 months (simple inputs)
        for (let i = 3; i < 18; i++) {
          const futureDate = new Date(currentDate);
          futureDate.setMonth(futureDate.getMonth() + i);
          months.push(formatMonthKey(futureDate));
        } 
        return months;
      };

      // ========== YEARLY FORECAST FUNCTIONS ==========

      const getLastMonthKey = (existingData: ForecastData): string | undefined => {
        const monthKeys = Object.keys(existingData).filter(key => key.match(/^\d{4}-\d{2}$/));
        
        if (monthKeys.length === 0) {
          return undefined;
        }
        
        // Sort months and get the last one
        const sortedMonths = monthKeys.sort();
        const lastMonth = sortedMonths[sortedMonths.length - 1];
         
        return lastMonth;
      };

      const createFreshYearlyStructure = (lastMonthKey?: string): YearlyForecastData => {
        const yearlyData: YearlyForecastData = {};
        
        // Determine start year dynamically
        let startYear: number;
        
        if (lastMonthKey) {
          // Extract year from last month key (format: "YYYY-MM")
          startYear = parseInt(lastMonthKey.split('-')[0]); 
        } else {
          // If no last month key, calculate from current date + 15 months
          const futureDate = new Date();
          futureDate.setMonth(futureDate.getMonth() + 15);
          startYear = futureDate.getFullYear(); 
        }
        
        const endYear = startYear + 4; // 5 years total including start year 
        
        for (let year = startYear; year <= endYear; year++) {
          yearlyData[year.toString()] = {
            plan: 0,
            actual: 0
          };
        }
        
        return yearlyData;
      };

      const createFreshForecastStructure = (): ForecastData => {
        const idealMonths = generateIdealMonthStructure();
        const freshData: ForecastData = {};
        
        const currentMonthKey = formatMonthKey(new Date());
        
        idealMonths.forEach((monthKey, index) => {
          const isCurrentMonth = monthKey === currentMonthKey;
          freshData[monthKey] = {
            plan: 0,
            actual: 0,
            forecast: 0,
            date: isCurrentMonth ? '' : undefined,
            items: (index >= 3 && index <= 5) ? createEmptyItems() : {}
          };
        });
        
        // Initialize yearly data dynamically
        const lastMonthKey = idealMonths[idealMonths.length - 1];
        freshData.yearly = createFreshYearlyStructure(lastMonthKey);
        
        return freshData;
      };

      // Type guard to check if data is ForecastMonthData
      const isForecastMonthData = (data: any): data is ForecastMonthData => {
        return data && 
               typeof data.plan === 'number' && 
               typeof data.actual === 'number';
      };

      const handleMonthProgression = (forecastData: ForecastData): ForecastData => {
        const now = new Date();
        const currentMonthKey = formatMonthKey(now);
         
        const originalMonths = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
        const idealMonths = generateIdealMonthStructure();
        
        // Check if progression is needed
        if (originalMonths[0] === idealMonths[0]) { 
          return forecastData;
        }
 
        
        const updatedData: ForecastData = {};

        // STEP 1: Create new structure with ideal months
        idealMonths.forEach((monthKey, index) => {
          // Check if this month exists in original data
          const existingData = forecastData[monthKey];
          if (existingData && isForecastMonthData(existingData)) {
            // Month exists - copy as-is
            updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
          } else {
            // New month - create empty
            updatedData[monthKey] = {
              plan: 0,
              actual: 0,
              items: (index >= 3 && index <= 5) ? createEmptyItems() : {}
            };
          }
        });
 
        
        // Find the current month index in ideal structure
        const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
        
        idealMonths.forEach((monthKey, index) => {
          const monthData = updatedData[monthKey];
          
          if (monthData && isForecastMonthData(monthData)) {
            // Check if this is CURRENT or FUTURE month (not historical)
            const isCurrentOrFuture = index >= currentMonthIndex;
            
            if (isCurrentOrFuture) {
              // Store the current actual value before resetting
              const oldActual = monthData.actual;
              
              // PROPAGATION: Same month's actual becomes plan, actual resets to 0
              monthData.plan = oldActual;  // Carry forward actual as plan
              monthData.actual = 0;        // Reset actual to 0
              
              // Apply same logic to items for detailed months
              if (monthData.items && Object.keys(monthData.items).length > 0) {
                Object.keys(monthData.items).forEach(itemKey => {
                  const item = monthData.items![itemKey];
                  const oldItemActual = item.actual;
                  
                  item.plan = oldItemActual;  // Carry forward item actual as plan
                  item.actual = 0;            // Reset item actual to 0
                  
                 });
              }
            } else {
              // HISTORICAL month - keep original values
             }
          }
        });

        // Preserve yearly data if it exists
        if (forecastData.yearly) {
          updatedData.yearly = JSON.parse(JSON.stringify(forecastData.yearly));
        } else {
          // Create fresh yearly data based on new last month
          const lastMonthKey = idealMonths[idealMonths.length - 1];
          updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
        }

         return updatedData;
      };

      const propagateMonthValues = (forecastData: ForecastData): ForecastData => {
        if (Object.keys(forecastData).length === 0) {
          return forecastData;
        }
        
        return handleMonthProgression(forecastData);
      };

      // ========== DATA LOADING FUNCTIONS ==========
const loadProjects = async (): Promise<void> => {
  try {
    // ✅ $select + $expand must include every field ROLE_OWN_FIELDS in auth.ts
    // matches against, otherwise getAssignedProjects() returns [] for the
    // KRIYA_Engg_* roles (and their Invoicing rows are filtered out entirely).
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

    if (!res.ok) {
      // ---------------------------------------------------------------
      // Fallback path - still carries the engineering IDs so
      // getAssignedProjects() can match via a separate user lookup.
      // ---------------------------------------------------------------
      console.warn('⚠️ Expand query failed, trying fallback...');
      const fallbackRes = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items` +
          `?$select=Id,Title,CustomerName,PC,ProjectReference,` +
          `GroupHead1Id,ProjectManagerId,` +
          `EnggLeadId,EnggSubLead1Id,EnggSubLead2Id,EngineeringManagersId`,
        { headers: { Accept: "application/json;odata=verbose" } }
      );

      if (!fallbackRes.ok) {
        throw new Error(`Failed to load projects: ${fallbackRes.status} ${fallbackRes.statusText}`);
      }

      const fallbackData = await fallbackRes.json();
      const rawProjects: any[] = fallbackData.d?.results || [];

      // Memoised user-id → name resolver (avoids duplicate calls)
      const userCache = new Map<number, string>();
      const getUserName = async (id: number | undefined): Promise<string> => {
        if (!id) return "";
        if (userCache.has(id)) return userCache.get(id)!;
        try {
          const uRes = await fetch(
            `${siteUrl}/_api/web/getuserbyid(${id})?$select=Id,Title,Email`,
            { headers: { Accept: "application/json;odata=verbose" } }
          );
          if (uRes.ok) {
            const uData = await uRes.json();
            const name = uData.d?.Title || "";
            userCache.set(id, name);
            return name;
          }
        } catch (e) {
          console.warn(`Could not fetch user ${id}:`, e);
        }
        userCache.set(id, "");
        return "";
      };

      const cleanedProjects: ProjectItem[] = await Promise.all(
        rawProjects.map(async (p: any) => ({
          ...p,
          CustomerName: p.CustomerName || "",
          PC: p.PC || "",
          ProjectReference: p.ProjectReference || "",
          GroupHead1: await getUserName(p.GroupHead1Id),
          GroupHead1Id: p.GroupHead1Id,
          ProjectManager: await getUserName(p.ProjectManagerId),
          ProjectManagerId: p.ProjectManagerId,
          // ✅ Engineering fields
          EnggLead: await getUserName(p.EnggLeadId),
          EnggLeadId: p.EnggLeadId,
          EnggSubLead1: await getUserName(p.EnggSubLead1Id),
          EnggSubLead1Id: p.EnggSubLead1Id,
          EnggSubLead2: await getUserName(p.EnggSubLead2Id),
          EnggSubLead2Id: p.EnggSubLead2Id,
          EngineeringManagers: await getUserName(p.EngineeringManagersId),
          EngineeringManagersId: p.EngineeringManagersId
        }))
      );

      const projectMap = new Map<string, ProjectItem>();
      cleanedProjects.forEach(p => projectMap.set(p.Title, p));
      projects = Array.from(projectMap.values());

      console.log(`📋 Total projects from SharePoint (fallback): ${projects.length}`);

      const assignedProjects = getAssignedProjects(projects);
      console.log(`📋 Assigned projects: ${assignedProjects.length}`);
      console.log('📋 Assigned project titles:', assignedProjects.map(p => p.Title));

      projectSelect.innerHTML =
        `<option value="">Select Project</option>` +
        assignedProjects.map(p => `<option value="${p.Id}">${p.Title}</option>`).join("");

      filterProject.innerHTML =
        '<option value="">All Projects</option>' +
        [...new Set(assignedProjects.map(p => p.Title))]
          .sort()
          .map(project => `<option value="${project}">${project}</option>`)
          .join("");
      return;
    }

    // ---------------------------------------------------------------
    // Normal (expanded) path
    // ---------------------------------------------------------------
    const data = await res.json();
    const rawProjects: any[] = data.d?.results || [];

    const cleanedProjects: ProjectItem[] = rawProjects.map((p: any) => ({
      ...p,
      CustomerName: p.CustomerName || "",
      PC: p.PC || "",
      ProjectReference: p.ProjectReference || "",
      GroupHead1: p.GroupHead1 || "",
      GroupHead1Id: p.GroupHead1Id,
      ProjectManager: p.ProjectManager || "",
      ProjectManagerId: p.ProjectManagerId,
      // ✅ Engineering fields
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
    cleanedProjects.forEach(p => projectMap.set(p.Title, p));
    projects = Array.from(projectMap.values());

    if (projects.length > 0) {
      console.log('📋 Sample project data:', {
        Title: projects[0].Title,
        ProjectManager: projects[0].ProjectManager,
        GroupHead1: projects[0].GroupHead1,
        EnggLead: projects[0].EnggLead,
        EnggSubLead1: projects[0].EnggSubLead1,
        EnggSubLead2: projects[0].EnggSubLead2,
        EngineeringManagers: projects[0].EngineeringManagers
      });
    }

    const assignedProjects = getAssignedProjects(projects);
    console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);
    console.log('📋 Assigned project titles:', assignedProjects.map(p => p.Title));

    projectSelect.innerHTML =
      `<option value="">Select Project</option>` +
      assignedProjects.map(p => `<option value="${p.Id}">${p.Title}</option>`).join("");

    filterProject.innerHTML =
      '<option value="">All Projects</option>' +
      [...new Set(assignedProjects.map(p => p.Title))]
        .sort()
        .map(project => `<option value="${project}">${project}</option>`)
        .join("");

  } catch (error) {
    console.error("Error loading projects:", error);
    toast("error", "Failed to load projects");
  }
};

      const loadForecasts = async (): Promise<void> => {
        try {
          const loading = createLoadingOverlay("Loading Invoicing data...");
          loaderHost().appendChild(loading);
          const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,InvoiceActual,CreatedDate,UpdatedDate&$expand=Project&$top=5000`, {
            headers: { Accept: "application/json;odata=verbose" }
          });
          const data: SharePointResponse<InvoicingForecastItem> = await res.json();

          // ✅ Get assigned projects
          const assignedProjects = getAssignedProjects(projects);
          const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));

          const allForecasts = (data.d?.results || []).map(forecast => ({
            ...forecast,
            ForecastData: typeof forecast.ForecastData === 'string' ? 
              forecast.ForecastData : 
              JSON.stringify(forecast.ForecastData || {}),
            InvoiceActual: typeof forecast.InvoiceActual === 'string' ? 
              forecast.InvoiceActual : 
              JSON.stringify(forecast.InvoiceActual || {}),
            Freeze: forecast.Freeze || false
          }));

          // ✅ Only keep forecasts for assigned projects
          forecasts = allForecasts.filter(forecast => {
            const projectTitle = forecast.Project?.Title;
            return projectTitle && assignedProjectTitles.has(projectTitle);
          });

          console.log(`✅ Loaded ${forecasts.length} invoicing forecasts`);

        } catch (error) {
          console.error("Error loading invoicing forecasts:", error);
          toast("error", "Failed to load invoicing forecasts");
        } finally {
          removeLoadingOverlay();
        }
      };

const getGroupHeadDisplayName = (groupHead: string | { Title: string; Email?: string } | undefined): string => {
  if (!groupHead) return "-";
  if (typeof groupHead === "object" && groupHead !== null) {
    return groupHead.Title || "-";
  }
  // ✅ cast: the object branch was handled above, so this is the plain string
  return groupHead as string;
};

      // ========== EXPORT FUNCTIONS ==========

      // Add this utility function at the top with other utility functions
      const getMonthName = (monthNumber: string): string => {
        const months = [
          "Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ];
        return months[parseInt(monthNumber) - 1] || "Month";
      };

      // UPDATED: Main Export Excel - Only Previous and Current columns
      const exportToExcelWithFullStyle = (): void => { 
        // Get selected year and month from filters
        const selectedYear = filterYear.value;
        const selectedMonth = filterMonth.value;
        
        if (!selectedYear || !selectedMonth) {
          toast("error", "Please select year and month to export");
          return;
        }

        // Use filtered data from pagination (current view)
        const exportData = filteredData.length > 0 ? filteredData : forecasts;
        
        if (exportData.length === 0) {
          toast("error", "No data available to export");
          return;
        }

        const exportLoading = createLoadingOverlay("Preparing the Excel workbook...");
        loaderHost().appendChild(exportLoading);

        try {
          // Generate headers according to your specified structure
          const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);
           
          // Prepare data for Excel
          const excelData = exportData.map(forecast => {
            const projectTitle = forecast.Project?.Title || "-";
            const proj = projects.find(p => p.Title === projectTitle);
            
            // Parse data from both sources
            const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
            
            // Get OBL value from database
            const oblValue = forecast.OBL || 0;
            
            // Selected month key
            const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
            const monthName = getMonthName(selectedMonth);
            
            // Get data for selected month
            const selectedMonthForecastData = forecastData[selectedMonthKey] as ForecastMonthData || { plan: 0, actual: 0 };
            const selectedMonthPlan = selectedMonthForecastData.plan || 0;
            const selectedMonthCurrent = selectedMonthForecastData.actual || 0;

            // Create base row data - ONLY Previous and Current columns
            const rowData: any = {
              "SO/Project Ref": proj?.ProjectReference || "-",
              "Project Name": projectTitle,
              "Project Manager": typeof proj?.ProjectManager === "object" ? proj.ProjectManager?.Title : proj?.ProjectManager || "-",
              "Group Head": getGroupHeadDisplayName(proj?.GroupHead1),
              "Customer Name": proj?.CustomerName || "-",
              "Profit Centre": forecast.ProfitCentre || "-",
              "Invoicing": oblValue,
              // Selected month columns - only Previous and Current
              [`${monthName} ${selectedYear} (P)`]: selectedMonthPlan,
              [`${monthName} ${selectedYear} (C)`]: selectedMonthCurrent,
            };
            
            // Add next 17 months data (previous + current only)
            allMonths.forEach(monthKey => {
              const [year, month] = monthKey.split('-');
              const monthLabel = getMonthName(month);
              const displayLabel = `${monthLabel} ${year}`;
              
              // Get forecast data for plan and current
              const forecastMonthData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
              const planValue = forecastMonthData.plan || 0;
              const currentValue = forecastMonthData.actual || 0;
              
              // Add both previous (plan) and current columns for each month
              rowData[`${displayLabel} (P)`] = planValue;
              rowData[`${displayLabel} (C)`] = currentValue;
            });
            
            // Add 5-year data (previous + current)
            lastFiveYears.forEach(year => {
              const yearData = forecastData.yearly?.[year] || { plan: 0, actual: 0 };
              rowData[`${year} (P)`] = yearData.plan;
              rowData[`${year} (C)`] = yearData.actual;
            });     
            
            // Calculate balance
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

          // Create workbook with styling 
          const wb = XLSX.utils.book_new();
          
          // Convert data to worksheet with headers
          const ws = XLSX.utils.json_to_sheet(excelData, { header: headers });

          // Define styles
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

          // Apply styles to all cells
          const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
          
          // Style header row (row 0)
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            if (!ws[cellAddress]) continue;
            
            if (!ws[cellAddress].s) {
              ws[cellAddress].s = {};
            }
            ws[cellAddress].s = headerStyle;
          }
          
          // Style data rows
          for (let row = range.s.r + 1; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
              const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
              if (!ws[cellAddress]) continue;
              
              // Determine if this is a number column
              const isNumberColumn = col >= 6; // Starting from OBL column
              const styleToApply = isNumberColumn ? numberStyle : dataStyle;
              
              if (!ws[cellAddress].s) {
                ws[cellAddress].s = {};
              }
              ws[cellAddress].s = styleToApply;
            }
          }

          // Set column widths
          const colWidths = headers.map((header: string) => {
            if (header.includes('Project Name')) return { wch: 25 };
            if (header.includes('Customer Name')) return { wch: 20 };
            if (header.includes('Profit Centre') || header.includes('SO/Project Ref')) return { wch: 12 };
            if (header.includes('Group Head') || header.includes('Project Manager')) return { wch: 15 };
            if (header.includes('OBL') || header.includes('Balance')) return { wch: 12 };
            if (header.includes('(P)') || header.includes('(C)')) return { wch: 10 };
            return { wch: 12 };
          });
          ws['!cols'] = colWidths;

          // Freeze header row
          ws['!freeze'] = { x: 0, y: 1 };

          // Add worksheet to workbook
          XLSX.utils.book_append_sheet(wb, ws, "Invoicing Forecast");

          // Generate filename and export
          const monthName = getMonthName(selectedMonth);
          const fileName = `Invoicing_Forecast_${monthName}_${selectedYear}.xlsx`;
          
          XLSX.writeFile(wb, fileName);
           
          toast("success", `Exported ${exportData.length} records for ${monthName} ${selectedYear}`);
          
        } catch (error) { 
          toast("error", "Failed to export data to Excel");
        } finally {
          removeLoadingOverlay();
        }
      };

      // Updated helper function to generate headers - ONLY Previous and Current
      const generateHeadersStructure = (selectedYear: string, selectedMonth: string, exportData: InvoicingForecastItem[]) => {
        const baseHeaders = [
          "SO/Project Ref",
          "Project Name", 
          "Project Manager", 
          "Group Head", 
          "Customer Name", 
          "Profit Centre",
          "Invoicing",
        ];

        // Add selected month columns - only Previous and Current
        const selectedMonthName = getMonthName(selectedMonth);
        const selectedMonthHeaders = [
          `${selectedMonthName} ${selectedYear} (P)`,
          `${selectedMonthName} ${selectedYear} (C)`
        ];

        // Get all unique months from the data
        const allMonthKeys = new Set<string>();
        
        exportData.forEach(forecast => {
          const forecastData = safeJsonParse(forecast.ForecastData, {});
          Object.keys(forecastData).forEach(key => {
            if (key && key.match(/^\d{4}-\d{2}$/)) {
              allMonthKeys.add(key);
            }
          });
        });

        // Sort months and take NEXT 17 months (excluding current month)
        const sortedMonthKeys = Array.from(allMonthKeys).sort();
        const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
        const selectedIndex = sortedMonthKeys.indexOf(selectedMonthKey);
        
        // Start from the month AFTER the selected month
        const startIndex = selectedIndex >= 0 ? selectedIndex + 1 : 0;
        const next17Months = sortedMonthKeys.slice(startIndex, startIndex + 17);
       
        // Add monthly columns (previous + current for each of the next 17 months)
        const monthlyHeaders: string[] = [];
        next17Months.forEach(monthKey => {
          const [year, month] = monthKey.split('-');
          const monthLabel = getMonthName(month);
          const displayLabel = `${monthLabel} ${year}`;
          monthlyHeaders.push(`${displayLabel} (P)`, `${displayLabel} (C)`);
        });

        // Get last 5 years from yearly forecast data
        const lastFiveYears = getLastFiveYears(exportData);
        const yearlyHeaders: string[] = [];
        lastFiveYears.forEach(year => {
          yearlyHeaders.push(`${year} (P)`, `${year} (C)`);
        });

        // Add balance column
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

      // Helper function to get last 5 years from data
      const getLastFiveYears = (data: InvoicingForecastItem[]): string[] => {
        const years = new Set<string>();
        
        data.forEach(forecast => {
          const forecastData = safeJsonParse(forecast.ForecastData, {});
          if (forecastData.yearly) {
            Object.keys(forecastData.yearly).forEach(year => {
              years.add(year);
            });
          }
        });
        
        // Sort years and get last 5
        const sortedYears = Array.from(years).sort();
        return sortedYears.slice(-5);
      };

      // DETAILED EXPORT - KEPT AS IS (unchanged)
      const exportDetailedItemsToExcel = (): void => { 
        // Get selected year and month from filters
        const selectedYear = filterYear.value;
        const selectedMonth = filterMonth.value;
        
        if (!selectedYear || !selectedMonth) {
          toast("error", "Please select year and month to export");
          return;
        }

        // Use filtered data from pagination (current view)
        const exportData = filteredData.length > 0 ? filteredData : forecasts;
        
        if (exportData.length === 0) {
          toast("error", "No data available to export");
          return;
        }

        const detailsLoading = createLoadingOverlay("Preparing the detailed workbook...");
        loaderHost().appendChild(detailsLoading);

        try {
          // Get current and next 2 months for detailed items
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

          // Get next 15 months for monthly totals
          const nextFifteenMonths: string[] = [];
          for (let i = 3; i < 18; i++) {
            const monthDate = new Date(currentDate);
            monthDate.setMonth(monthDate.getMonth() + i);
            const monthKey = formatMonthKey(monthDate);
            if (monthKey) {
              nextFifteenMonths.push(monthKey);
            }
          }

          // Get last 5 years from actual forecast data
          const lastFiveYears = getLastFiveYearsFromForecastData(exportData);
       
          // Get month names for headers with safety checks
          const selectedMonthName = getMonthName(selectedMonth) || selectedMonth;
          const currentMonthName = getMonthName(detailedMonths[0]?.split('-')[1] || '') || detailedMonths[0]?.split('-')[1] || '';
          const nextMonth1Name = getMonthName(detailedMonths[1]?.split('-')[1] || '') || detailedMonths[1]?.split('-')[1] || '';
          const nextMonth2Name = getMonthName(detailedMonths[2]?.split('-')[1] || '') || detailedMonths[2]?.split('-')[1] || '';
          const currentMonthYear = detailedMonths[0]?.split('-')[0] || '';
          const nextMonth1Year = detailedMonths[1]?.split('-')[0] || '';
          const nextMonth2Year = detailedMonths[2]?.split('-')[0] || '';

          // Prepare data for Excel with two-row header structure
          const excelData: any[][] = [];

          // First row headers - Current month has 5 columns, Next months have 3 columns
          const firstRowHeaders = [
            "SO/Project Ref", 
            "Project Name", 
            "Customer Name",
            "Profit Centre", 
            "Group Head", 
            "Project Manager", 
            "Invoicing",
            
            // Selected month totals (4 columns) - H-K
            `${selectedMonthName} ${selectedYear} (P)`,
            `${selectedMonthName} ${selectedYear} (C)`,
            `${selectedMonthName} ${selectedYear} (R)`,
            `${selectedMonthName} ${selectedYear} (F)`,
            
            // Current month items section header - will be merged across 5 columns - L-P
            `${currentMonthName} ${currentMonthYear}`.toUpperCase(),
            "", "", "", "", // Placeholders for merged columns M-P
            
            // Next month 1 totals (2 columns) - Q-R
            `${nextMonth1Name} ${nextMonth1Year} (P)`,
            `${nextMonth1Name} ${nextMonth1Year} (C)`,
            
            // Next month 1 items section header - will be merged across 3 columns - S-U
            `${nextMonth1Name} ${nextMonth1Year}`.toUpperCase(),
            "", "", // Placeholders for merged columns T-U
            
            // Next month 2 totals (2 columns) - V-W
            `${nextMonth2Name} ${nextMonth2Year} (P)`,
            `${nextMonth2Name} ${nextMonth2Year} (C)`,
            
            // Next month 2 items section header - will be merged across 3 columns - X-Z
            `${nextMonth2Name} ${nextMonth2Year}`.toUpperCase(),
            "", "", // Placeholders for merged columns Y-Z
          ];

          // Add next 15 months headers (P and C for each month) - starting from column AA
          nextFifteenMonths.forEach(monthKey => {
            const [year, month] = monthKey.split('-');
            const monthLabel = getMonthName(month) || month;
            firstRowHeaders.push(`${monthLabel} ${year} (P)`, `${monthLabel} ${year} (C)`);
          });

          // Add 5-year headers (P and C for each year)
          lastFiveYears.forEach(year => {
            firstRowHeaders.push(`${year} (P)`, `${year} (C)`);
          });

          // Add Balance header
          firstRowHeaders.push("Balance");

          // Second row headers
          const secondRowHeaders = [
            "", "", "", "", "", "", "", // Empty for first 7 columns (A-G)
            
            // Selected month totals (4 columns) - H-K
            "", "", "", "",
            
            // Current month items headers (5 columns) - L-P
            "Item Description", "Previous", "Current", "Forecast", "Date",
            
            // Next month 1 totals (2 columns) - Q-R
            "", "",
            
            // Next month 1 items headers (3 columns) - S-U
            "Item Description", "Previous", "Current",
            
            // Next month 2 totals (2 columns) - V-W
            "", "",
            
            // Next month 2 items headers (3 columns) - X-Z
            "Item Description", "Previous", "Current",
          ];

          // Add empty headers for next 15 months and 5 years
          const remainingColumns = firstRowHeaders.length - secondRowHeaders.length;
          for (let i = 0; i < remainingColumns; i++) {
            secondRowHeaders.push("");
          }

          excelData.push(firstRowHeaders);
          excelData.push(secondRowHeaders);

          // Prepare detailed data with merged project information
          exportData.forEach(forecast => {
            const projectTitle = forecast.Project?.Title || "-";
            const proj = projects.find(p => p.Title === projectTitle);
            
            // Parse data from both sources
            const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
            
            // Get OBL value from database
            const oblValue = forecast.OBL || 0;

            // Get data for detailed months with safety checks
            const detailedMonthsData: any[] = [];
            for (let i = 0; i < 3; i++) {
              const monthKey = detailedMonths[i];
              if (monthKey) {
                const monthForecastData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0, forecast: 0, items: createEmptyItems() };
                detailedMonthsData.push({
                  plan: monthForecastData.plan || 0,
                  current: monthForecastData.actual || 0,
                  forecast: monthForecastData.forecast || 0,
                  items: monthForecastData.items || createEmptyItems()
                });
              } else {
                detailedMonthsData.push({
                  plan: 0,
                  current: 0,
                  forecast: 0,
                  items: createEmptyItems()
                });
              }
            }

            // Calculate total current value for balance
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

            // Create one row for each item (1-10)
            for (let itemNum = 1; itemNum <= 10; itemNum++) {
              const itemKey = itemNum.toString();
              
              const rowData: any[] = [
                // Base information (7 columns) - A-G
                proj?.ProjectReference || "-",
                projectTitle,
                proj?.CustomerName || "-",
                forecast.ProfitCentre || "-",
                getGroupHeadDisplayName(proj?.GroupHead1),
                typeof proj?.ProjectManager === "object" ? proj.ProjectManager?.Title : proj?.ProjectManager || "-",
                oblValue,
                
                // Selected month totals (4 columns) - H-K
                detailedMonthsData[0]?.plan || 0,
                detailedMonthsData[0]?.current || 0,
                0, // Removed actual invoicing from detailed export
                detailedMonthsData[0]?.forecast || 0,
              ];

              // Current month items (5 columns) - L-P
              const currentMonthItems = detailedMonthsData[0]?.items || createEmptyItems();
              const currentMonthItem = currentMonthItems[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
              rowData.push(
                currentMonthItem.name || `Item ${itemNum}`,
                currentMonthItem.plan,
                currentMonthItem.actual,
                currentMonthItem.forecast || 0,
                currentMonthItem.date || ""
              );

              // Next month 1 totals (2 columns) - Q-R
              rowData.push(
                detailedMonthsData[1]?.plan || 0,
                detailedMonthsData[1]?.current || 0
              );

              // Next month 1 items (3 columns) - S-U
              const nextMonth1Items = detailedMonthsData[1]?.items || createEmptyItems();
              const nextMonth1Item = nextMonth1Items[itemKey] || { name: '', plan: 0, actual: 0 };
              rowData.push(
                nextMonth1Item.name || `Item ${itemNum}`,
                nextMonth1Item.plan,
                nextMonth1Item.actual
              );

              // Next month 2 totals (2 columns) - V-W
              rowData.push(
                detailedMonthsData[2]?.plan || 0,
                detailedMonthsData[2]?.current || 0
              );

              // Next month 2 items (3 columns) - X-Z
              const nextMonth2Items = detailedMonthsData[2]?.items || createEmptyItems();
              const nextMonth2Item = nextMonth2Items[itemKey] || { name: '', plan: 0, actual: 0 };
              rowData.push(
                nextMonth2Item.name || `Item ${itemNum}`,
                nextMonth2Item.plan,
                nextMonth2Item.actual
              );

              // Add next 15 months data (P and C for each month) - starting from column AA
              nextFifteenMonths.forEach(monthKey => {
                const monthForecastData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
                rowData.push(
                  monthForecastData.plan || 0,
                  monthForecastData.actual || 0
                );
              });

              // Add 5-year data (P and C for each year)
              lastFiveYears.forEach(year => {
                const yearData = forecastData.yearly?.[year] || { plan: 0, actual: 0 };
                rowData.push(
                  yearData.plan,
                  yearData.actual
                );
              });

              // Add Balance - only for first item of each project, empty for others
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
       
          // Create workbook
          const wb = XLSX.utils.book_new();
          const ws = XLSX.utils.aoa_to_sheet(excelData);

          // Define styles
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

          // Apply styles to all cells
          const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
          
          // Style first header row (row 0)
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            if (!ws[cellAddress]) continue;
            if (!ws[cellAddress].s) ws[cellAddress].s = {};
            ws[cellAddress].s = mainHeaderStyle;
          }
          
          // Style second header row (row 1)
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 1, c: col });
            if (!ws[cellAddress]) continue;
            if (!ws[cellAddress].s) ws[cellAddress].s = {};
            ws[cellAddress].s = subHeaderStyle;
          }
          
          // Style data rows (starting from row 2)
          for (let row = range.s.r + 2; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
              const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
              if (!ws[cellAddress]) continue;
              
              // Determine column type for styling
              let styleToApply = dataStyle;
              
              // Text columns: Item Description columns and text fields
              if (col === 0 || col === 1 || col === 2 || // SO/Project Ref, Project Name, Customer Name
                  col === 11 || col === 17 || col === 23) { // Item Description columns
                styleToApply = textStyle;
              } 
              // Number columns: Previous, Current, Forecast, OBL, Month totals
              else if (col === 6 || // OBL
                       (col >= 7 && col <= 10) || // Selected month totals (H-K)
                       col === 12 || col === 13 || col === 14 || // Current Month Previous, Current, Forecast
                       col === 18 || col === 19 || // Next month 1 Previous, Current
                       col === 24 || col === 25 || // Next month 2 Previous, Current
                       (col >= 26) // All remaining columns (next 15 months, 5 years, balance)
              ) {
                styleToApply = numberStyle;
              }
              
              if (!ws[cellAddress].s) {
                ws[cellAddress].s = {};
              }
              ws[cellAddress].s = styleToApply;
            }
          }

          // Set column widths
          const colWidths = [
            { wch: 12 }, // A: SO/Project Ref
            { wch: 25 }, // B: Project Name
            { wch: 20 }, // C: Customer Name
            { wch: 12 }, // D: Profit Centre
            { wch: 15 }, // E: Group Head
            { wch: 15 }, // F: Project Manager
            { wch: 10 }, // G: OBL
            { wch: 12 }, // H: Selected Month (P)
            { wch: 12 }, // I: Selected Month (C)
            { wch: 12 }, // J: Selected Month (R)
            { wch: 12 }, // K: Selected Month (F)
            { wch: 20 }, // L: Current Month Item Description
            { wch: 10 }, // M: Current Month Previous
            { wch: 10 }, // N: Current Month Current
            { wch: 10 }, // O: Current Month Forecast
            { wch: 12 }, // P: Current Month Date
            { wch: 12 }, // Q: Next Month 1 (P)
            { wch: 12 }, // R: Next Month 1 (C)
            { wch: 20 }, // S: Next Month 1 Item Description
            { wch: 10 }, // T: Next Month 1 Previous
            { wch: 10 }, // U: Next Month 1 Current
            { wch: 12 }, // V: Next Month 2 (P)
            { wch: 12 }, // W: Next Month 2 (C)
            { wch: 20 }, // X: Next Month 2 Item Description
            { wch: 10 }, // Y: Next Month 2 Previous
            { wch: 10 }, // Z: Next Month 2 Current
          ];

          // Add widths for next 15 months (each month has 2 columns)
          for (let i = 0; i < nextFifteenMonths.length * 2; i++) {
            colWidths.push({ wch: 12 });
          }

          // Add widths for 5 years (each year has 2 columns)
          for (let i = 0; i < lastFiveYears.length * 2; i++) {
            colWidths.push({ wch: 12 });
          }

          // Add width for Balance
          colWidths.push({ wch: 12 });

          ws['!cols'] = colWidths;

          // Ensure merges array exists
          if (!ws['!merges']) {
            ws['!merges'] = [];
          }

          // Merge header cells for month sections
          ws['!merges'].push({ s: { r: 0, c: 11 }, e: { r: 0, c: 15 } }); // Current month header - spans 5 columns (L-P)
          ws['!merges'].push({ s: { r: 0, c: 17 }, e: { r: 0, c: 19 } }); // Next month 1 header - spans 3 columns (S-U)
          ws['!merges'].push({ s: { r: 0, c: 23 }, e: { r: 0, c: 25 } }); // Next month 2 header - spans 3 columns (X-Z)

          // Merge project information and totals (excluding item details)
          let currentRow = 2; // Start from first data row
          while (currentRow < excelData.length) {
            const endRow = Math.min(currentRow + 9, excelData.length - 1); // 10 rows per project
            
            // Merge project information (columns A-G)
            for (let col = 0; col < 7; col++) {
              if (currentRow < endRow && ws['!merges']) {
                ws['!merges'].push({ s: { r: currentRow, c: col }, e: { r: endRow, c: col } });
              }
            }
            
            // Merge ALL totals columns (selected month totals, detailed month totals, next 15 months, 5 years, balance)
            const totalColumnsToMerge = [
              7, 8, 9, 10, // H-K: Selected month totals
              16, 17, // Q-R: Next month 1 totals
              21, 22, // V-W: Next month 2 totals
            ];
            
            // Add next 15 months and yearly columns (starting from column AA)
            const monthlyStartCol = 26; // Column AA (26 = 0-based index for AA)
            for (let col = monthlyStartCol; col < firstRowHeaders.length; col++) {
              totalColumnsToMerge.push(col);
            }
            
            // Apply all merges - WITH NULL CHECK
            totalColumnsToMerge.forEach(col => {
              if (currentRow < endRow && ws['!merges']) {
                ws['!merges'].push({ s: { r: currentRow, c: col }, e: { r: endRow, c: col } });
              }
            });
            
            currentRow = endRow + 1; // Move to next project group
          }

          // Clear the placeholder cells in first row after merging
          for (let col = 12; col <= 15; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            delete ws[cellAddress];
          }
          for (let col = 18; col <= 19; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            delete ws[cellAddress];
          }
          for (let col = 24; col <= 25; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            delete ws[cellAddress];
          }

          // Freeze header rows and first few columns
          ws['!freeze'] = { x: 7, y: 2 };

          // Add worksheet to workbook
          XLSX.utils.book_append_sheet(wb, ws, "Detailed Items");

          // Generate filename and export
          const monthName = getMonthName(selectedMonth) || selectedMonth;
          const fileName = `Invoicing_Forecast_Detailed_Items_${monthName}_${selectedYear}.xlsx`;
          
          XLSX.writeFile(wb, fileName);
           
          toast("success", `Exported ${excelData.length - 2} detailed item records`);
          
        } catch (error) { 
          toast("error", "Failed to export detailed items data to Excel");
        } finally {
          removeLoadingOverlay();
        }
      };
      
      // Helper function to get last 5 years from forecast data
      const getLastFiveYearsFromForecastData = (data: InvoicingForecastItem[]): string[] => {
        const years = new Set<string>();
        
        data.forEach(forecast => {
          const forecastData = safeJsonParse(forecast.ForecastData, {});
          if (forecastData.yearly) {
            Object.keys(forecastData.yearly).forEach(year => {
              years.add(year);
            });
          }
        });
        
        // Sort years and get last 5
        const sortedYears = Array.from(years).sort();
        return sortedYears.slice(-5);
      };

      // ======  PAGINATION FUNCTIONS ======

      const initializePagination = (data: InvoicingForecastItem[]): void => { 
        // Store whatever data is passed
        filteredData = data || []; 
        calculateTotalPages();
        currentPage = 1;
        updatePaginationControls();
        renderCurrentPage();
        
        // Refresh totals header
        const displayMonths = getDisplayMonths();
        const currentMonthKey = formatMonthKey(new Date());
        setTimeout(() => {
          renderTotalsHeader(displayMonths, currentMonthKey);
        }, 100);
      };

      const calculateTotalPages = (): void => {
        if (!filteredData || filteredData.length === 0) {
          totalPages = 1;
        } else {
          totalPages = Math.ceil(filteredData.length / pageSize);
        }
      };

      const updatePaginationControls = (): void => {
        // ✅ Every element is OPTIONAL: a missing node must never abort the whole
        //    update (the old "Page X of Y" pill is gone - it is now the "of N"
        //    after the page tiles).
        const setText = (el: HTMLElement | null, value: string): void => {
          if (el) el.textContent = value;
        };

        const totalItems = filteredData ? filteredData.length : 0;
        const startItem = totalItems > 0 ? (currentPage - 1) * pageSize + 1 : 0;
        const endItem = totalItems > 0 ? Math.min(currentPage * pageSize, totalItems) : 0;

        if (currentPageSpan) currentPageSpan.textContent = currentPage.toString();

        setText(totalPagesSpan, totalPages.toString());
        setText(startItemSpan, startItem.toString());
        setText(endItemSpan, endItem.toString());
        setText(totalItemsSpan, totalItems.toString());

        if (firstPageBtn) firstPageBtn.disabled = currentPage === 1 || totalItems === 0;
        if (prevPageBtn) prevPageBtn.disabled = currentPage === 1 || totalItems === 0;
        if (nextPageBtn) nextPageBtn.disabled = currentPage === totalPages || totalItems === 0;
        if (lastPageBtn) lastPageBtn.disabled = currentPage === totalPages || totalItems === 0;

        updatePageNumbers();
      };

      const updatePageNumbers = (): void => {
        pageNumbersDiv.innerHTML = '';
        
        if (!filteredData || filteredData.length === 0) return;
        
        const maxVisiblePages = 5;
        let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
        let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
        
        if (endPage - startPage + 1 < maxVisiblePages) {
          startPage = Math.max(1, endPage - maxVisiblePages + 1);
        }
        
        for (let i = startPage; i <= endPage; i++) {
          const pageBtn = document.createElement('button');
          pageBtn.className = `invoicing-page-number ${i === currentPage ? 'invoicing-active' : ''}`;
          pageBtn.textContent = i.toString();
          pageBtn.addEventListener('click', () => goToPage(i));
          pageNumbersDiv.appendChild(pageBtn);
        }
      };

      const goToPage = (page: number): void => {
        if (page < 1 || page > totalPages || page === currentPage) return;
        
        currentPage = page;
        updatePaginationControls();
        renderCurrentPage();
      };
      
      const renderCurrentPage = (): void => {
        // Check if filteredData has data
        if (!filteredData || filteredData.length === 0) {
          tableBody.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:20px;">No records found</td></tr>`;
          const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;
          if (totalsHeader) totalsHeader.innerHTML = '';
          return;
        }
        
        const startIndex = (currentPage - 1) * pageSize;
        const endIndex = Math.min(startIndex + pageSize, filteredData.length);
        const pageData = filteredData.slice(startIndex, endIndex);
        
        renderTableWithMonthColumns(pageData);
      };

      // ========== TABLE RENDERING FUNCTIONS ==========

      // ✅ CHANGE 6: keep totals thead + main thead stacked and frozen while scrolling vertically (same as inventory)
      // Measures each header <tr> (from the totals thead AND the main thead, in DOM order)
      // and stacks them with cumulative top offsets so they freeze contiguously on scroll.
      const measureStickyHeaderOffsets = (): void => {
        const table = document.querySelector('.month-columns-table-inv') as HTMLTableElement | null;
        // Only meaningful when the table is actually laid out & visible.
        if (!table) return;
        const wrapper = table.closest('.invoicing-table-scroll-wrapper') as HTMLElement | null;
        if (wrapper && wrapper.offsetParent === null) return; // hidden by display:none ancestors

        // 1) Clear any previous inline tops so we measure from natural (unstuck) positions.
        const allCells = table.querySelectorAll<HTMLElement>(
          '#totalsHeader th, #totalsHeader td, #monthColumnsHeader th, #monthColumnsHeader td'
        );
        allCells.forEach(cell => { cell.style.top = ''; });

        // 2) Force layout so getBoundingClientRect returns real heights.
        void table.offsetHeight;

        // 3) Stack all header rows across BOTH theads in order.
        const headerRows = table.querySelectorAll<HTMLTableRowElement>(
          '#totalsHeader tr, #monthColumnsHeader tr'
        );
        let offset = 0;
        headerRows.forEach(row => {
          const h = row.getBoundingClientRect().height || 0;
          row.querySelectorAll<HTMLElement>('th, td').forEach(cell => {
            cell.style.position = 'sticky';
            cell.style.top = `${offset}px`;
          });
          offset += h;
        });
      };

      // Public wrapper: measures immediately and again on the next frame so heights are
      // correct once fonts/scrollbars have settled.
      const applyStickyHeaderOffsets = (): void => {
        measureStickyHeaderOffsets();
        if (typeof window !== 'undefined') {
          requestAnimationFrame(measureStickyHeaderOffsets);
        }
      };

      // ✅ CHANGE 1 + 3: Totals header - same structure, colours and column order as inventory
      // Rows are rendered directly inside <thead id="totalsHeader"> (no nested table) so they align with the main columns.
      const renderTotalsHeader = (displayMonths: string[], currentMonthKey: string): void => {
        const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;
        if (!totalsHeader) return;

        totalsHeader.innerHTML = '';

        const allData = forecasts;
        const filtered = filteredData.length > 0 ? filteredData : forecasts;

        const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
        const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);

        // const baseColumns = 6;
        // const monthColumns = displayMonths.length * 2;
        // const totalColspan = baseColumns + monthColumns;

        const buildTotalsRow = (label: string, totals: any, css: string) => {
          const monthCells = displayMonths.map(m => {
            const t = totals.monthly[m] || { plan: 0, current: 0 };
            return `
              <td class="invoicing-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-US')}</td>
              <td class="invoicing-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-US')}</td>
            `;
          }).join('');

          return `
            <tr class="${css}">
              <td class="invoicing-project-ref-col" style="text-align:center; vertical-align:middle; min-width:120px; max-width:120px; width:120px;">${css === 'all-projects-total' ? '★' : '●'}</td>
              <td class="invoicing-customer-col" style="text-align:center; vertical-align:middle; min-width:130px; max-width:130px; width:130px;"></td>
              <td class="invoicing-project-col" style="text-align:left; vertical-align:middle; min-width:180px; max-width:180px; width:180px; font-weight:bold; background-color:${css === 'all-projects-total' ? '#d1ecf1' : '#d4edda'};">
                ${label}
              </td>
              <td class="invoicing-gh-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
              <td class="invoicing-pm-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
              <td class="inv-col" style="text-align:right; vertical-align:middle; min-width:90px; max-width:90px; width:90px; font-weight:bold;">${totals.obl.toLocaleString('en-US')}</td>
              ${monthCells}
            </tr>
          `;
        };

        let html = `
          <tr class="invoicing-totals-title-row">
             
          </tr>
        `;

        if (allData.length > 0) {
          html += buildTotalsRow(`All Projects (${allData.length})`, allTotals, 'all-projects-total');
        }

        // Filtered row is always shown (same as before)
        if (filtered.length > 0) {
          html += buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total');
        }

        totalsHeader.innerHTML = html;
        applyStickyHeaderOffsets();
      };

      const calculateTotals = (
        data: InvoicingForecastItem[],
        displayMonths: string[],
        currentMonthKey: string
      ): any => {

        const totals = {
          obl: 0,
          monthly: {} as {
            [key: string]: { plan: number; current: number }
          }
        };

        // Init months
        displayMonths.forEach(monthKey => {
          totals.monthly[monthKey] = { plan: 0, current: 0 };
        });

        data.forEach(forecast => {
          totals.obl += Math.round(Number(forecast.OBL) || 0);

          const forecastData = safeJsonParse(forecast.ForecastData, {});

          displayMonths.forEach(monthKey => {
            const m = forecastData[monthKey] || {};

            const plan = Math.round(Number(m.plan) || 0);
            const current = Math.round(Number(m.actual) || 0);

            totals.monthly[monthKey].plan += plan;
            totals.monthly[monthKey].current += current;
          });
        });

        return totals;
      };

      const renderTableWithMonthColumns = (data: InvoicingForecastItem[]): void => {
        if (!data.length) {
          tableBody.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:20px;">No records found</td></tr>`;
          const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement;
          if (totalsHeader) totalsHeader.innerHTML = ''; 
          return;
        }
        
        const displayMonths = getDisplayMonths();
        const currentMonthKey = formatMonthKey(new Date());
        
        // First, render the totals rows
        renderTotalsHeader(displayMonths, currentMonthKey);
        
        // Determine current month index
        const currentMonthIndex = displayMonths.indexOf(currentMonthKey);
        
        // ✅ CHANGE 3: Column order same as inventory:
        // Project Reference → Customer → Project → Group Head → Project Manager → Invoicing → Month columns
        let mainHeaderHTML = `
          <tr>
            <!-- Project Reference - FIRST (clickable edit link) -->
            <th rowspan="2" class="invoicing-project-ref-col" style="min-width:120px; max-width:120px; width:120px;">
              <div class="invoicing-sort-header-container">
                <span class="invoicing-header-label">Project Reference</span>
                <div class="invoicing-sort-buttons-container">
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'projectReference' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                          data-sort="projectReference" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'projectReference' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                          data-sort="projectReference" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Customer -->
            <th rowspan="2" class="invoicing-customer-col" style="min-width:130px; max-width:130px; width:130px;">
              <div class="invoicing-sort-header-container">
                <span class="invoicing-header-label">Customer</span>
                <div class="invoicing-sort-buttons-container">
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                          data-sort="customer" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                          data-sort="customer" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Project Column with Sort -->
            <th rowspan="2" class="invoicing-project-col" style="min-width:180px; max-width:180px; width:180px;">
              <div class="invoicing-sort-header-container">
                <span class="invoicing-header-label">Project</span>
                <div class="invoicing-sort-buttons-container">
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                          data-sort="project" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                          data-sort="project" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Group Head -->
            <th rowspan="2" class="invoicing-gh-col" style="min-width:100px; max-width:100px; width:100px;">
              <div class="invoicing-sort-header-container">
                <span class="invoicing-header-label">Group Head</span>
                <div class="invoicing-sort-buttons-container">
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                          data-sort="groupHead" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                          data-sort="groupHead" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Project Manager -->
            <th rowspan="2" class="invoicing-pm-col" style="min-width:100px; max-width:100px; width:100px;">
              <div class="invoicing-sort-header-container">
                <span class="invoicing-header-label">Project Manager</span>
                <div class="invoicing-sort-buttons-container">
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                          data-sort="projectManager" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                          data-sort="projectManager" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Invoicing Column -->
            <th rowspan="2" class="inv-col" style="min-width:90px; max-width:90px; width:90px;">
              <div class="invoicing-sort-header-container">
                <span class="invoicing-header-label">Invoicing</span>
                <div class="invoicing-sort-buttons-container">
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                          data-sort="obl" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="invoicing-sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                          data-sort="obl" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            ${displayMonths.map((month, monthIndex) => {
              const isCurrentMonth = monthIndex === currentMonthIndex;
              const monthDate = new Date(month + "-01");
              const monthLabel = formatMonthLabel(monthDate);
              const monthHeaderClass = isCurrentMonth ? 'invoicing-current-month-header' : 'other-month-header';
              const monthLabelDisplay = isCurrentMonth ? `${monthLabel} (Current Month)` : monthLabel;
              
              return `
                <th colspan="2" class="invoicing-month-scroll-col invoicing-month-column-header ${monthHeaderClass}" style="min-width:60px;">
                  <div class="invoicing-sort-header-container">
                    <span class="invoicing-header-label">${monthLabelDisplay}</span>
                    <div class="invoicing-sort-buttons-container">
                      <button class="invoicing-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                              data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort by Month Total Ascending">▲</button>
                      <button class="invoicing-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                              data-sort="month" data-direction="desc" data-month-index="${monthIndex}" title="Sort by Month Total Descending">▼</button>
                    </div>
                  </div>
                </th>
              `;
            }).join('')}
          </tr>
          
          <tr>
            ${displayMonths.map((month, monthIndex) => {
              const isCurrentMonth = monthIndex === currentMonthIndex;
              const subHeaderClass = isCurrentMonth ? 'current-month-subheader current-month-header-table' : 'other-month-subheader';
              
              return `
                <th class="invoicing-month-scroll-col invoicing-month-subheader ${subHeaderClass}" style="min-width:60px;">
                  <div class="invoicing-sort-header-container">
                    <span>Previous</span>
                    <div class="invoicing-sort-buttons-container">
                      <button class="invoicing-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Ascending">▲</button>
                      <button class="invoicing-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Descending">▼</button>
                    </div>
                  </div>
                </th>
                <th class="invoicing-month-scroll-col invoicing-month-subheader ${subHeaderClass}" style="min-width:60px;">
                  <div class="invoicing-sort-header-container">
                    <span>Current</span>
                    <div class="invoicing-sort-buttons-container">
                      <button class="invoicing-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'invoicing-sort-asc invoicing-active' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Ascending">▲</button>
                      <button class="invoicing-sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'invoicing-sort-desc invoicing-active' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Descending">▼</button>
                    </div>
                  </div>
                </th>
              `;
            }).join('')}
          </tr>
        `;

        // Set main header HTML
        monthColumnsHeader.innerHTML = mainHeaderHTML;

        // Build table body with actual data
        let tableBodyHTML = '';

  data.forEach((forecast) => {
  const projectTitle = forecast.Project?.Title || "-";
  const proj = projects.find(p => p.Title === projectTitle);
  
  const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
  const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-US');
  
  // ✅ Project Manager - Person/Group type
  let projectManagerDisplay = "-";
  if (proj?.ProjectManager) {
    projectManagerDisplay = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
      ? proj.ProjectManager.Title || ""
      : (proj.ProjectManager as string) || "";
  }
  
  // ✅ GroupHead1 - Person/Group type
  let groupHeadDisplayName = "-";
  if (proj?.GroupHead1) {
    groupHeadDisplayName = typeof proj.GroupHead1 === "object" && proj.GroupHead1 !== null
      ? proj.GroupHead1.Title || ""
      : (proj.GroupHead1 as string) || "";
  }
  
  // ✅ CHANGE 5: only Project Reference is the clickable edit link (same as inventory).
  // Project name is plain text. If a project has no reference, fall back to "#Id" so every row is still editable.
  const projectRefText = proj?.ProjectReference || `#${forecast.Id}`;
  const projectRefLink = `<a href="#" class="invoicing-edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;" title="Edit ${projectTitle}">${projectRefText}</a>`;
  
  let rowHTML = `
    <tr data-forecast-id="${forecast.Id}">
      <td class="invoicing-project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
      <td class="invoicing-customer-col" style="min-width:130px; max-width:130px; width:130px;">${proj?.CustomerName || "-"}</td>
      <td class="invoicing-project-col" style="min-width:180px; max-width:180px; width:180px;">${projectTitle}</td>
      <td class="invoicing-gh-col" style="min-width:100px; max-width:100px; width:100px;">${groupHeadDisplayName}</td>
      <td class="invoicing-pm-col" style="min-width:100px; max-width:100px; width:100px;">${projectManagerDisplay}</td>
      <td class="inv-col" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
  `;

  // Add month columns data (only Previous and Current)
  displayMonths.forEach((monthKey) => {
    const monthData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
    
    const planValue = monthData.plan || 0;
    const currentValue = monthData.actual || 0;
    
    rowHTML += `
      <td class="invoicing-month-scroll-col" style="min-width:60px;">${planValue}</td>
      <td class="invoicing-month-scroll-col" style="min-width:60px;">${currentValue}</td>
    `;
  });

  rowHTML += `</tr>`;
  tableBodyHTML += rowHTML;
});

        tableBody.innerHTML = tableBodyHTML;

        // ✅ CHANGE 6: recompute sticky offsets after header + body render
        applyStickyHeaderOffsets();
        
        // Attach event listeners
        attachSortButtonListeners();
        attachEditTriggerListeners();
      };

      const attachSortButtonListeners = (): void => {
        document.querySelectorAll('.invoicing-sort-btn[data-sort]').forEach(btn => {
          btn.addEventListener('click', (e) => {
            const button = e.currentTarget as HTMLButtonElement;
            const sortColumn = button.dataset.sort;
            const direction = button.dataset.direction as 'asc' | 'desc';
            const subColumn = button.dataset.subcolumn;
            const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;
            
            if (sortColumn && direction) {
              handleInvoicingSort(sortColumn, direction, subColumn, monthIndex);
            }
          });
        });
      };

      const attachEditTriggerListeners = (): void => {
        document.querySelectorAll('.invoicing-edit-trigger').forEach(link => {
          link.addEventListener('click', (e) => {
            e.preventDefault();
            const id = parseInt((e.currentTarget as HTMLElement).dataset.id!);
            editForecast(id);
          });
        });
      };

      const handleInvoicingSort = (column: string, direction: 'asc' | 'desc', subColumn?: string, monthIndex?: number): void => {
        // Update current sort state
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
            
            case 'projectReference':
              const projRefA = projects.find(p => p.Title === a.Project?.Title);
              const projRefB = projects.find(p => p.Title === b.Project?.Title);
              valueA = projRefA?.ProjectReference || "";
              valueB = projRefB?.ProjectReference || "";
              break;
            
            case 'month':
              const displayMonths = getDisplayMonths();
              
              const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
              const monthKeyToUse = displayMonths[monthIndexToUse];
              
              const dataA = safeJsonParse(a.ForecastData, {});
              const dataB = safeJsonParse(b.ForecastData, {});
              const monthDataA = dataA[monthKeyToUse] as ForecastMonthData || { plan: 0, actual: 0 };
              const monthDataB = dataB[monthKeyToUse] as ForecastMonthData || { plan: 0, actual: 0 };
              
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
                  default:
                    valueA = monthDataA.actual || 0;
                    valueB = monthDataB.actual || 0;
                }
              } else {
                valueA = monthDataA.actual || 0;
                valueB = monthDataB.actual || 0;
              }
              break;
            
            case 'project':
              valueA = a.Project?.Title || "";
              valueB = b.Project?.Title || "";
              break;
            
            case 'projectManager':
              const projMA = projects.find(p => p.Title === a.Project?.Title);
              const projMB = projects.find(p => p.Title === b.Project?.Title);
              valueA = typeof projMA?.ProjectManager === "object" ? projMA.ProjectManager?.Title : projMA?.ProjectManager || "";
              valueB = typeof projMB?.ProjectManager === "object" ? projMB.ProjectManager?.Title : projMB?.ProjectManager || "";
              break;
            
case 'groupHead':
  const projGA = projects.find(p => p.Title === a.Project?.Title);
  const projGB = projects.find(p => p.Title === b.Project?.Title);
  // ✅ Handle GroupHead1 as Person/Group
  if (projGA?.GroupHead1) {
    valueA = typeof projGA.GroupHead1 === "object" && projGA.GroupHead1 !== null
      ? projGA.GroupHead1.Title || ""
      : projGA.GroupHead1 || "";
  } else {
    valueA = "";
  }
  if (projGB?.GroupHead1) {
    valueB = typeof projGB.GroupHead1 === "object" && projGB.GroupHead1 !== null
      ? projGB.GroupHead1.Title || ""
      : projGB.GroupHead1 || "";
  } else {
    valueB = "";
  }
  break;
            
            case 'customer':
              const projA = projects.find(p => p.Title === a.Project?.Title);
              const projB = projects.find(p => p.Title === b.Project?.Title);
              valueA = projA?.CustomerName || "";
              valueB = projB?.CustomerName || "";
              break;
            
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
          columnName = subColumn.replace('-', ' ').replace(/\b\w/g, l => l.toUpperCase());
        } else if (column === 'obl') {
          columnName = 'Invoicing';
        } else if (column === 'project') {
          columnName = 'Project';
        } else if (column === 'projectReference') {
          columnName = 'Project Reference';
        } else if (column === 'projectManager') {
          columnName = 'Project Manager';
        } else if (column === 'groupHead') {
          columnName = 'Group Head';
        } else if (column === 'customer') {
          columnName = 'Customer';
        } else {
          const displayMonths = getDisplayMonths();
          const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
          const monthKey = displayMonths[monthIndexToUse];
          const monthDate = new Date(monthKey + "-01");
          columnName = formatMonthLabel(monthDate);
        }
        
        toast("info", `Sorted by ${columnName} (${sortDirectionText})`);
      };

      const clearFilters = (): void => {
        filterYear.value = "";
        filterMonth.value = "";
        filterProject.value = "";
        filterCustomer.value = "";
        filterProjectManager.value = "";
        filterGroupHead.value = "";
        filterProjectReference.value = "";
        
        // Reset sorting
        currentSortConfig = {
          column: '',
          direction: 'asc'
        };
        
        // Re-initialize filters
        initializeFilters();
        
        // Reset to all data
        filteredData = [...forecasts];
        calculateTotalPages();
        currentPage = 1;
        updatePaginationControls();
        renderCurrentPage();
        
        // Refresh totals header
        const displayMonths = getDisplayMonths();
        const currentMonthKey = formatMonthKey(new Date());
        setTimeout(() => {
          renderTotalsHeader(displayMonths, currentMonthKey);
        }, 100);
        
        toast("info", "All filters and sorting cleared");
      };

      // ========== HISTORICAL TOGGLE (CHANGE 2 - same as inventory) ==========

      const setupHistoricalToggle = (): void => {
        const toggleBtn = document.querySelector("#historicalToggleBtn") as HTMLButtonElement | null;
        const container = document.querySelector("#historicalContainer") as HTMLElement | null;
        const header = document.querySelector(".invoicing-historical-toggle-header") as HTMLElement | null;

        if (!toggleBtn || !container || !header) return;

        let isVisible = false;

        const updateHistoricalToggle = (): void => {
          const arrow = toggleBtn.querySelector(".invoicing-toggle-arrow") as HTMLElement | null;
          const label = toggleBtn.querySelector(".invoicing-toggle-label") as HTMLElement | null;

          if (isVisible) {
            container.style.display = "grid";
            toggleBtn.setAttribute("aria-expanded", "true");
            toggleBtn.title = "Hide historical data";
            if (arrow) arrow.textContent = "▲";
            if (label) label.textContent = "Hide";
          } else {
            container.style.display = "none";
            toggleBtn.setAttribute("aria-expanded", "false");
            toggleBtn.title = "Show historical data";
            if (arrow) arrow.textContent = "▼";
            if (label) label.textContent = "Show";
          }
        };

        const toggleHistorical = (event?: Event): void => {
          if (event) {
            event.preventDefault();
            event.stopPropagation();
          }

          isVisible = !isVisible;
          updateHistoricalToggle();
        };

        // Prevent duplicate handlers without replacing the button/header.
        toggleBtn.onclick = (event: MouseEvent): void => {
          toggleHistorical(event);
        };

        header.onclick = (event: MouseEvent): void => {
          const target = event.target as HTMLElement;

          // Button click is already handled by the button itself.
          if (target.closest("#historicalToggleBtn")) {
            return;
          }

          toggleHistorical(event);
        };

        // Historical data is collapsed by default.
        updateHistoricalToggle();
      };

      // Reset the toggle button to its collapsed state (used by resetForm)
      const resetHistoricalToggle = (): void => {
        historicalContainer.style.display = "none";
        if (historicalToggleBtn) {
          historicalToggleBtn.setAttribute("aria-expanded", "false");
          historicalToggleBtn.title = "Show historical data";
          const arrow = historicalToggleBtn.querySelector(".invoicing-toggle-arrow") as HTMLElement | null;
          const label = historicalToggleBtn.querySelector(".invoicing-toggle-label") as HTMLElement | null;
          if (arrow) arrow.textContent = "▼";
          if (label) label.textContent = "Show";
        }
      };

      // ========== FORM MANAGEMENT FUNCTIONS ==========

      // const showForm = (isEditMode: boolean = false): void => {
      //   formSection.style.display = "block";
      //   tableSection.style.display = "none";
        
      //   if (isEditMode) {
      //     formSection.classList.add('edit-mode');
      //     projectSelect.disabled = true;
      //     formTitle.textContent = "Edit Forecast";
      //   } else {
      //     formSection.classList.remove('edit-mode');
      //     projectSelect.disabled = true;
      //     formTitle.textContent = "Edit Forecast";
      //   }
      // };

      const hideForm = (): void => {
            formSection.style.display = "none";
            tableSection.style.display = "block";
            formSection.classList.remove('invoicing-edit-mode');
            projectSelect.disabled = false;
            
            // Reset any form state
            resetForm();

            // ✅ CHANGE 6: table is visible again → recompute sticky offsets
            setTimeout(applyStickyHeaderOffsets, 50);
        };

      const onSaveSuccess = (): void => {
        // Show success message
        toast("success", editingId ? "Forecast updated successfully!" : "Forecast created successfully!");
        
        // Reset form state
        resetForm();
        
        // Show table section
        tableSection.style.display = "block";
        formSection.style.display = "none";
        
        // Reload data to reflect changes - keeps the filters the user has applied
        reloadInvoicingData();
      };
      
      const resetForm = (): void => {
        editingId = null;
        form.reset();
        projectSelect.value = "";
        
        // Clear project details display
        if (projectNameDisplay) projectNameDisplay.textContent = "-";
        if (projectRefDisplay) projectRefDisplay.textContent = "-";
        if (profitCenterDisplay) profitCenterDisplay.textContent = "-";
        if (projectManagerDisplay) projectManagerDisplay.textContent = "-";
        if (groupHeadDisplay) groupHeadDisplay.textContent = "-";
        if (customerDisplay) customerDisplay.textContent = "-";
        
        historicalContainer.innerHTML = "";
        resetHistoricalToggle();
        firstRowContainer.innerHTML = "";
        remainingContainer.innerHTML = "";
        yearlyContainer.innerHTML = "";
        
        // Hide all forecast sections
        toggleForecastSections(false);
        
        const oblContainer = document.querySelector('.obl-section-container-invoicing') as HTMLElement;
        if (oblContainer) {
          oblContainer.className = 'obl-section-container-invoicing invoicing-hidden-section';
        }
        oblLastMonthInput.removeAttribute('readonly');
        oblLastMonthInput.placeholder = "Enter OBL";
        setNumericValue(oblLastMonthInput, 0);
        
        const fieldNote = document.querySelector('.field-note') as HTMLElement;
        if (fieldNote) fieldNote.textContent = "For new projects: Enter initial OBL value";
        
        oblValidationMessage.textContent = "✅ All values are within INVOICING limit";
        oblValidationMessage.className = 'validation-message-invoicing';
        
        updateOBLValidationDisplay();
      };

      // ========== VALIDATION FUNCTIONS ==========

      const calculateTotalCurrentValue = (): number => {
        let total = 0;
        
        // Include all future current values EXCEPT derived-readonly ones.
        // Detailed months (indices 3-5): top "Current" is READONLY (auto = items sum), so the
        // real editable values live in the item rows. Long-term months (6+): editable single Current.
        const futureMonths = allMonths.slice(3);
        futureMonths.forEach(monthKey => {
          const idx = allMonths.indexOf(monthKey);
          if (idx >= 3 && idx <= 5) {
            document.querySelectorAll<HTMLInputElement>(`input.item-current-input[data-month="${monthKey}"]:not([readonly])`).forEach(input => {
              total += getNumericValue(input);
            });
          } else {
            document.querySelectorAll<HTMLInputElement>(`input.main-current-input[data-month="${monthKey}"]:not([readonly])`).forEach(input => {
              total += getNumericValue(input);
            });
          }
        });
        
        // Include yearly actual values
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>(`input.yearly-current-input-inv:not([readonly])`);
        yearlyCurrentInputs.forEach(input => {
          total += getNumericValue(input);
        });
        
        return total;
      };

      const validateOBLLimit = (changedInput?: HTMLInputElement): boolean => {
        const oblValue = getNumericValue(oblLastMonthInput);
        const totalCurrent = calculateTotalCurrentValue();
        
        totalCurrentValueSpan.textContent = totalCurrent.toLocaleString('en-US');
        
        const remaining = oblValue - totalCurrent;
        remainingValueSpan.textContent = remaining.toLocaleString('en-US');
        
        if (oblValue === 0) {
          oblLineFill.style.width = `0%`;
          oblLineFill.className = 'line-fill-invoicing';
          
          if (totalCurrent > 0) {
            showOBLErrorToast('⚠️ OBL is zero but current values exist. Please set OBL value.', true);
            oblValidationMessage.textContent = '⚠️ OBL is zero but current values exist. Please set OBL value.';
            oblValidationMessage.className = 'validation-message-invoicing danger-invoicing';
            blockAllCurrentInputs();
            return false;
          } else {
            const existingToast = document.querySelector('.invoicing-obl-error-toast');
            if (existingToast) existingToast.remove();
            oblValidationMessage.textContent = '✅ All values are within INVOICING limit';
            oblValidationMessage.className = 'validation-message-invoicing success-invoicing';
            unblockAllCurrentInputs();
            return true;
          }
        }
        
        const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
        oblLineFill.style.width = `${percentage}%`;
        
        oblLineFill.className = 'line-fill-invoicing';
        if (percentage > 90) {
          oblLineFill.classList.add('danger-invoicing');
        } else if (percentage > 75) {
          oblLineFill.classList.add('warning-invoicing');
        }
        
        oblValidationMessage.className = 'validation-message-invoicing';
        
        if (totalCurrent > oblValue) {
          const exceedAmount = totalCurrent - oblValue;
          showOBLErrorToast(`❌ Total current value exceeds INVOICING by ${exceedAmount.toLocaleString('en-US')}! Reduce values to continue.`);
          oblValidationMessage.textContent = `❌ Total current value exceeds INVOICING by ${exceedAmount.toLocaleString('en-US')}! Reduce values to continue.`;
          oblValidationMessage.classList.add('danger-invoicing');
          blockIncreasingCurrentInputs(changedInput);
          return false;
        } else if (percentage > 90) {
          showOBLErrorToast('⚠️ INVOICING limit nearly reached! Consider reducing values.', true);
          oblValidationMessage.textContent = '⚠️ INVOICING limit nearly reached! Consider reducing values.';
          oblValidationMessage.classList.add('warning-invoicing');
          unblockAllCurrentInputs();
          return true;
        } else {
          const existingToast = document.querySelector('.invoicing-obl-error-toast');
          if (existingToast) existingToast.remove();
          
          oblValidationMessage.textContent = '✅ All values are within INVOICING limit';
          oblValidationMessage.classList.add('success-invoicing');
          unblockAllCurrentInputs();
          return true;
        }
      };

      const blockAllCurrentInputs = (): void => {
        const currentInputs = document.querySelectorAll<HTMLInputElement>('input[data-type="actual"]:not([readonly])');
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>('input.yearly-current-input-inv:not([readonly])');
        
        currentInputs.forEach(input => {
          input.setAttribute('readonly', 'true');
          input.style.backgroundColor = '#f8f9fa';
          input.style.cursor = 'not-allowed';
        });
        
        yearlyCurrentInputs.forEach(input => {
          input.setAttribute('readonly', 'true');
          input.style.backgroundColor = '#f8f9fa';
          input.style.cursor = 'not-allowed';
        });
      };

      const blockIncreasingCurrentInputs = (allowedInput?: HTMLInputElement): void => {
        const currentInputs = document.querySelectorAll<HTMLInputElement>('input[data-type="actual"]:not([readonly])');
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>('input.yearly-current-input-inv:not([readonly])');
        
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
        const currentInputs = document.querySelectorAll<HTMLInputElement>('input[data-type="actual"]:not([data-month="0"])');
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>('input.yearly-current-input-inv');
        
        currentInputs.forEach(input => {
          // Derived Current (auto sum) must NEVER lose readonly.
          const isDerived = input.classList.contains('derived-month-current') || input.getAttribute('data-derived') === 'current';
          input.style.backgroundColor = '';
          input.style.cursor = '';
          input.removeAttribute('data-block-increase');
          if (!isDerived) {
            input.removeAttribute('readonly');
          } else {
            input.setAttribute('readonly', 'true');
          }
        });
        
        yearlyCurrentInputs.forEach(input => {
          input.style.backgroundColor = '';
          input.style.cursor = '';
          input.removeAttribute('data-block-increase');
          input.removeAttribute('readonly');
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
        const mainCurrentInput = document.querySelector<HTMLInputElement>(`input.main-current-input[data-month="${monthKey}"]`);
        const itemCurrentInputs = document.querySelectorAll<HTMLInputElement>(`input.item-current-input[data-month="${monthKey}"]`);
        
        // Top "Current" is now DERIVED = sum of the item Currents (always equal, no conflict).
        let itemsTotal = 0;
        itemCurrentInputs.forEach(input => {
            itemsTotal += getNumericValue(input);
        });
        
        if (mainCurrentInput) {
            if (itemsTotal > 0) {
                mainCurrentInput.value = itemsTotal.toString();
                mainCurrentInput.classList.add('invoicing-input-with-value');
                mainCurrentInput.classList.remove('invoicing-input-empty');
            } else {
                mainCurrentInput.value = '';
                mainCurrentInput.classList.add('invoicing-input-empty');
                mainCurrentInput.classList.remove('invoicing-input-with-value');
            }
            // Keep the derived field read-only even if an OBL block/unblock touched it.
            mainCurrentInput.setAttribute('readonly', 'true');
            mainCurrentInput.classList.add('invoicing-readonly-field');
        }
        
        unblockAllMonthItemInputs(monthKey);
        return true;
      };
      
      const unblockAllMonthItemInputs = (monthKey: string): void => {
        const itemInputs = document.querySelectorAll<HTMLInputElement>(`input.item-current-input[data-month="${monthKey}"]`);
        
        itemInputs.forEach(input => {
            input.style.backgroundColor = '';
            input.style.cursor = '';
            input.removeAttribute('data-block-increase');
            input.removeAttribute('readonly');
        });
      };

      const updateOBLValidationDisplay = (changedInput?: HTMLInputElement): void => {
        validateOBLLimit(changedInput);
        validateMonthItemTotals(changedInput);
      };

      // ========== FORECAST FORM FUNCTIONS ==========

      const onProjectChange = (projectId: number, existingForecast?: InvoicingForecastItem, preloadedData?: ForecastData): void => {
        const project = projects.find(p => p.Id === projectId);
        if (!project) return;

        // Show all forecast sections when project is selected
        toggleForecastSections(true);

        renderProjectDetails(project);
        
        let forecastData: ForecastData = {};
        if (existingForecast) {
          // preloadedData = the same record already parsed + month-progressed by editForecast()
          forecastData = preloadedData || safeJsonParse(existingForecast.ForecastData, {});
        }

        const isNewProject = !existingForecast;
        
        const oblContainer = document.querySelector('.obl-section-container-invoicing') as HTMLElement;
        const fieldNote = document.querySelector('.field-note') as HTMLElement;
        
        if (isNewProject) {
          if (oblContainer) {
            oblContainer.className = 'obl-section-container-invoicing invoicing-visible-section new-project';
          }
          oblLastMonthInput.removeAttribute('readonly');
          oblLastMonthInput.placeholder = "Enter initial OBL for new project";
          if (fieldNote) fieldNote.textContent = "For new projects: Enter initial OBL value";
          oblValidationMessage.textContent = "Enter initial OBL value for new project";
          oblValidationMessage.className = 'validation-message-invoicing';
          
          const defaultOBL = project.PC ? 100000 : 0;
          setNumericValue(oblLastMonthInput, defaultOBL);
        } else {
          if (oblContainer) {
            oblContainer.className = 'obl-section-container-invoicing invoicing-visible-section existing-project';
          }
          oblLastMonthInput.setAttribute('readonly', 'true');
          
          const oblValue = existingForecast?.OBL || 0;
          setNumericValue(oblLastMonthInput, oblValue);
          
          oblLastMonthInput.placeholder = "OBL from database";
          if (fieldNote) fieldNote.textContent = "OBL value from database record";
          
          if (oblValue > 0) {
            oblLastMonthInput.value = oblValue.toString();
            oblLastMonthInput.className = oblLastMonthInput.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
          }
        }

        renderForecastInterface(forecastData, getNumericValue(oblLastMonthInput), existingForecast);
      };

      // Add this function to show/hide forecast sections
      const toggleForecastSections = (show: boolean): void => {
        const sections = [
          '.obl-section-container-invoicing',
          '#historicalForecastSection',
          '#detailedForecastSection', 
          '#longtermForecastSection',
          '#yearlyForecastSection'
        ];
        
        sections.forEach(selector => {
          const element = document.querySelector(selector) as HTMLElement;
          if (element) {
            if (show) {
              element.classList.remove('invoicing-hidden-section');
              element.classList.add('invoicing-visible-section');
            } else {
              element.classList.remove('invoicing-visible-section');
              element.classList.add('invoicing-hidden-section');
            }
          }
        });
      };

      // Updated renderProjectDetails to use the 6-row layout
const renderProjectDetails = (project: ProjectItem): void => {
  if (projectNameDisplay) projectNameDisplay.textContent = project.Title || "-";
  if (projectRefDisplay) projectRefDisplay.textContent = project.ProjectReference || "-";
  if (profitCenterDisplay) profitCenterDisplay.textContent = project.PC || "-";
  if (projectManagerDisplay) {
    projectManagerDisplay.textContent = typeof project.ProjectManager === "object" ? 
      project.ProjectManager?.Title : project.ProjectManager || "-";
  }
  // ✅ GroupHead1 - Person/Group type
  if (groupHeadDisplay) {
    groupHeadDisplay.textContent = getGroupHeadDisplayName(project.GroupHead1);
  }
  if (customerDisplay) customerDisplay.textContent = project.CustomerName || "-";
};

      const getForecastMonths = (): string[] => {
        return generateIdealMonthStructure();
      };

      const renderForecastInterface = (existingData: ForecastData, displayOBL: number, existingForecast?: InvoicingForecastItem): void => {
        allMonths = getForecastMonths();
        
        setNumericValue(oblLastMonthInput, displayOBL);

        renderHistoricalMonths(allMonths.slice(0, 3), existingData);
        renderFirstRow(allMonths.slice(3, 6), existingData);
        renderRemainingMonths(allMonths.slice(6), existingData);
        renderYearlyForecast(existingData.yearly || createFreshYearlyStructure(getLastMonthKey(existingData)));

        // Protect historical elements from any interactions
        protectHistoricalElements();

        // ✅ CHANGE 2: (re)bind the historical toggle after the section is rendered - collapsed by default
        setupHistoricalToggle();
        
        updateOBLValidationDisplay();
        attachForecastEventListeners();
      };

      const protectHistoricalElements = (): void => {
        // Remove any existing event listeners from historical elements by cloning and replacing
        document.querySelectorAll('.invoicing-historical-month input, .invoicing-historical-month select, .invoicing-historical-month button').forEach(element => {
          const newElement = element.cloneNode(true) as HTMLElement;
          element.parentNode?.replaceChild(newElement, element);
        });
        
        // Add global event prevention for historical sections
        document.querySelectorAll('.invoicing-historical-month').forEach(section => {
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
        const planInputClass = data.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        const actualInputClass = data.actual > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        
        const items = data.items || createEmptyItems();
        
        return `
          <div class="month-box-inv invoicing-historical-month">
            <div class="month-header-inv">${monthLabel} (Historical - Readonly)</div>
            
            <div class="month-block">
              <div class="main-inputs-container-inv">
                <div class="main-input-group-inv">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input ${planInputClass} number-input" />
                </div>
                <div class="main-input-group-inv">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        readonly class="main-current-input ${actualInputClass} number-input" />
                </div>
              </div>
            </div>

            <div class="invoicing-items-section invoicing-historical-items-section">
              <div class="items-header invoicing-historical-items-header">
                <div>Item Description</div>
                <div>Previous</div>
                <div>Current</div>
              </div>
              
              ${Array.from({length: 10}, (_, i) => {
                const itemNum = i + 1;
                const itemKey = itemNum.toString();
                const item = items[itemKey] || { name: '', plan: 0, actual: 0 };
                const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                const itemPlanClass = item.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const itemActualClass = item.actual > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const itemNameClass = item.name ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                
                return `
                  <div class="invoicing-item-row invoicing-historical-item-row">
                    <input type="text" placeholder="Item ${itemNum}" value="${item.name}" data-item="${itemKey}" data-month="${monthKey}" readonly 
                          class="${itemNameClass}" />
                    <input type="text" value="${itemPlanValue}" data-item-type="plan" data-item="${itemKey}" data-month="${monthKey}" 
                          readonly class="invoicing-item-plan-input ${itemPlanClass} number-input" />
                    <input type="text" value="${itemActualValue}" data-item-type="actual" data-item="${itemKey}" data-month="${monthKey}" 
                          readonly class="item-current-input ${itemActualClass} number-input" />
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      };

      // UPDATED: renderFirstRow - Current month (5 columns), Next 2 months (3 columns)
      const renderFirstRow = (months: string[], existingData: ForecastData): void => {
        let html = '';

        // Create single row container for all three months
        html += `<div class="invoicing-detailed-months-row">`;
        
        // Current Month - 44% - 5 columns (Item Description, Previous, Current, Forecast, Date)
        const currentMonth = months[0];
        if (currentMonth) {
          const monthDate = new Date(currentMonth + "-01");
          const monthLabel = formatMonthLabel(monthDate);
          const data = existingData[currentMonth] as ForecastMonthData || { plan: 0, actual: 0, items: {} };
          
          html += `<div class="invoicing-current-month-container">`;
          html += renderDetailedMonthBox5Col(currentMonth, monthLabel, data, true);
          html += `</div>`;
        }
        
        // Next 2 Months - 56% (split equally) - 3 columns (Item Description, Previous, Current)
        html += `<div class="invoicing-next-months-container">`;
        for (let i = 1; i <= 2; i++) {
          if (i < months.length) {
            const monthKey = months[i];
            const monthDate = new Date(monthKey + "-01");
            const monthLabel = formatMonthLabel(monthDate);
            const data = existingData[monthKey] as ForecastMonthData || { plan: 0, actual: 0, items: {} };
            
            html += `<div class="invoicing-next-month-box">`;
            html += renderDetailedMonthBox3Col(monthKey, monthLabel, data, false);
            html += `</div>`;
          }
        }
        html += `</div>`;
        
        html += `</div>`;

        firstRowContainer.innerHTML = html;
      };

      // ✅ DATE LOCK (same as inventory/SEIBIES): restrict dates to the CURRENT calendar month only.
      // min = 1st of the current month, max = last day of the current month. This ensures a current-month
      // date can never fall into the previous or the next month.
      const getCurrentMonthMinDate = (): string => {
        const now = new Date();
        return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`;
      };
      const getCurrentMonthMaxDate = (): string => {
        const now = new Date();
        const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        return `${lastDay.getFullYear()}-${String(lastDay.getMonth() + 1).padStart(2, "0")}-${String(lastDay.getDate()).padStart(2, "0")}`;
      };

      // RENDER: Current Month - 5 columns (Item Description, Previous, Current, Forecast, Date)
      const renderDetailedMonthBox5Col = (monthKey: string, monthLabel: string, data: ForecastMonthData, isCurrentMonth: boolean = false): string => {
        const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
        
        const planInputClass = data.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        
        const items = data.items || createEmptyItems();

        // ✅ Main Date is READONLY and auto-shows the HIGHEST (latest) date entered on the items.
        const computeMaxItemDate = (): string => {
          let latest = '';
          for (let i = 1; i <= 10; i++) {
            const d = items[i.toString()]?.date || '';
            if (d && (!latest || d > latest)) latest = d;
          }
          // fall back to the previously stored top-level date if no item has a date yet
          return latest || data.date || '';
        };
        const autoMainDate = computeMaxItemDate();

        // ✅ Month Current is READONLY and derived = SUM of that month's item Currents below.
        let itemsTotalSum = 0;
        Array.from({length: 10}, (_, i) => i + 1).forEach(itemNum => {
          const item = items[itemNum] || { name: '', plan: 0, actual: 0 };
          itemsTotalSum += item.actual || 0;
        });
        const derivedMonthCurrent = itemsTotalSum;
        const derivedCurrentDisplay = derivedMonthCurrent > 0 ? derivedMonthCurrent.toString() : '';
        const derivedCurrentClass = derivedMonthCurrent > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';

        // ✅ DATE LOCK: only dates INSIDE the current calendar month can be entered
        // (min = 1st of this month, max = last day of this month) - no previous or next month.
        const minDate = getCurrentMonthMinDate();
        const maxDate = getCurrentMonthMaxDate();
        
        return `
          <div class="invoicing-detailed-month-box invoicing-current-month-box">
            <div class="month-header-inv">
              ${monthLabel} 
             
            </div>
            
            <div class="month-block">
              <div class="invoicing-enhanced-inputs-container-5col">
                <!-- Previous Column -->
                <div class="invoicing-enhanced-input-group">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input ${planInputClass} number-input" />
                </div>
                
                <!-- Current Column - READONLY, auto = sum of item Currents -->
                <div class="invoicing-enhanced-input-group">
                  <label>Current</label>
                  <input type="text" value="${derivedCurrentDisplay}" data-type="actual" data-month="${monthKey}"
                        readonly data-derived="current" class="main-current-input derived-month-current invoicing-readonly-field ${derivedCurrentClass} number-input" title="Auto total of the item Currents below" />
                </div>
                
                <!-- Date Column - READONLY, auto = highest date set on the items -->
                <div class="invoicing-enhanced-input-group">
                  <label>Date</label>
                  <input type="date" value="${autoMainDate}" data-type="date" data-month="${monthKey}"
                        readonly min="${minDate}" max="${maxDate}" title="Auto date = highest date entered on the items"
                        class="main-date-input invoicing-readonly-field" />
                </div>
              </div>
            </div>

            <div class="invoicing-items-section">
              <div class="invoicing-items-header-5col">
                <div>Item Description</div>
                <div>Previous</div>
                <div>Current</div>
                <div>Date</div>
              </div>
              
              ${Array.from({length: 10}, (_, i) => i + 1).map(itemNum => {
                const item = items[itemNum] || { name: '', plan: 0, actual: 0, date: '' };
                const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                const itemDateValue = item.date || '';
                
                const itemPlanClass = item.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const itemActualClass = item.actual > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const itemNameClass = item.name ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                
                return `
                  <div class="invoicing-item-row-5col">
                    <!-- Item Description -->
                    <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
                          data-item="${itemNum}" data-month="${monthKey}" 
                          class="${itemNameClass}" />
                    
                    <!-- Previous (Plan) -->
                    <input type="text" value="${itemPlanValue}" 
                          data-item-type="plan" data-item="${itemNum}" data-month="${monthKey}" 
                          readonly class="invoicing-item-plan-input ${itemPlanClass} number-input" />
                    
                    <!-- Current (Actual) -->
                    <input type="text" value="${itemActualValue}" 
                          data-item-type="actual" data-item="${itemNum}" data-month="${monthKey}" 
                          class="item-current-input ${itemActualClass} number-input" />
                    
                    <!-- Date (locked to current month) -->
                    <input type="date" value="${itemDateValue}" 
                          data-item-type="date" data-item="${itemNum}" data-month="${monthKey}" 
                          min="${minDate}" max="${maxDate}" title="Select a date within the current month only"
                          class="item-date-input" />
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      };

      // RENDER: Next 2 Months - 3 columns (Item Description, Previous, Current)
      const renderDetailedMonthBox3Col = (monthKey: string, monthLabel: string, data: ForecastMonthData, isCurrentMonth: boolean = false): string => {
        const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
        
        const planInputClass = data.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        
        const items = data.items || createEmptyItems();

        // ✅ Month Current is READONLY and derived = SUM of that month's item Currents below.
        let itemsTotalSum = 0;
        Array.from({length: 10}, (_, i) => i + 1).forEach(itemNum => {
          const item = items[itemNum] || { name: '', plan: 0, actual: 0 };
          itemsTotalSum += item.actual || 0;
        });
        const derivedMonthCurrent = itemsTotalSum;
        const derivedCurrentDisplay = derivedMonthCurrent > 0 ? derivedMonthCurrent.toString() : '';
        const derivedCurrentClass = derivedMonthCurrent > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        
        return `
          <div class="invoicing-detailed-month-box invoicing-next-month-box">
            <div class="month-header-inv">
              ${monthLabel}
            </div>
            
            <div class="month-block">
              <div class="invoicing-enhanced-inputs-container-3col">
                <!-- Previous Column -->
                <div class="invoicing-enhanced-input-group">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input ${planInputClass} number-input" />
                </div>
                
                <!-- Current Column - READONLY, auto = sum of item Currents -->
                <div class="invoicing-enhanced-input-group">
                  <label>Current</label>
                  <input type="text" value="${derivedCurrentDisplay}" data-type="actual" data-month="${monthKey}"
                        readonly data-derived="current" class="main-current-input derived-month-current invoicing-readonly-field ${derivedCurrentClass} number-input" title="Auto total of the item Currents below" />
                </div>
              </div>
            </div>

            <div class="invoicing-items-section">
              <div class="invoicing-items-header-3col">
                <div>Item Description</div>
                <div>Previous</div>
                <div>Current</div>
              </div>
              
              ${Array.from({length: 10}, (_, i) => i + 1).map(itemNum => {
                const item = items[itemNum] || { name: '', plan: 0, actual: 0 };
                const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                
                const itemPlanClass = item.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const itemActualClass = item.actual > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const itemNameClass = item.name ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                
                return `
                  <div class="invoicing-item-row-3col">
                    <!-- Item Description -->
                    <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
                          data-item="${itemNum}" data-month="${monthKey}" 
                          class="${itemNameClass}" />
                    
                    <!-- Previous (Plan) -->
                    <input type="text" value="${itemPlanValue}" 
                          data-item-type="plan" data-item="${itemNum}" data-month="${monthKey}" 
                          readonly class="invoicing-item-plan-input ${itemPlanClass} number-input" />
                    
                    <!-- Current (Actual) -->
                    <input type="text" value="${itemActualValue}" 
                          data-item-type="actual" data-item="${itemNum}" data-month="${monthKey}" 
                          class="item-current-input ${itemActualClass} number-input" />
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      };

      const renderRemainingMonths = (months: string[], existingData: ForecastData): void => {
        let html = '';

        for (let row = 0; row < 3; row++) {
          html += `<div class="remaining-row-inv">`;
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
        const planInputClass = data.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        const actualInputClass = data.actual > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        
        return `
          <div class="yearly-box-inv">
            <div class="yearly-header-inv">${year}</div>
            
            <div class="yearly-inputs-container-inv">
              <div class="yearly-input-group-inv">
                <label>Previous</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                      readonly class="yearly-plan-input-inv ${planInputClass} number-input" />
              </div>
              <div class="yearly-input-group-inv">
                <label>Current</label>
                <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                      class="yearly-current-input-inv ${actualInputClass} number-input" />
              </div>
            </div>
          </div>
        `;
      };

      const renderSimpleMonthBox = (monthKey: string, monthLabel: string, data: ForecastMonthData): string => {
        const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
        const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
        const planInputClass = data.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        const actualInputClass = data.actual > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
        
        return `
          <div class="month-box-inv simple-month-inv">
            <div class="month-header-inv">${monthLabel}</div>
            
            <div class="main-inputs-container-inv">
              <div class="main-input-group-inv">
                <label>Previous</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}" 
                      readonly class="main-plan-input ${planInputClass} number-input" />
              </div>
              <div class="main-input-group-inv">
                <label>Current</label>
                <input type="text" value="${displayActualValue}"  id="actualCurrent" data-type="actual" data-month="${monthKey}" 
                      class="main-current-input ${actualInputClass} number-input" />
              </div>
            </div>
          </div>
        `;
      };

      // ========== FORM EVENT LISTENERS ==========

      const attachForecastEventListeners = (): void => {
        // ✅ Numeric-only inputs (all of them; the historical ones stay readonly and are skipped anyway)
        document.querySelectorAll<HTMLInputElement>('.number-input').forEach(input => {
          if (input.dataset.numericBound === 'true') return;
          input.dataset.numericBound = 'true';
          // Historical inputs are readonly - never attach the numeric filter to them.
          if (input.hasAttribute('readonly')) return;
          allowOnlyNumbers(input);
        });

        // ✅ Value inputs (historical inputs are readonly, so they are skipped)
        document.querySelectorAll<HTMLInputElement>('input[data-type], input[data-item-type]').forEach(input => {
          if (input.hasAttribute('readonly')) return;
          if (input.dataset.valueBound === 'true') return;
          input.dataset.valueBound = 'true';

          // store the previous value so we can decide "increase" vs "decrease"
          input.addEventListener('focus', () => {
            input.setAttribute('data-old-value', getNumericValue(input).toString());
          });

          input.addEventListener('input', () => {
            updateOBLValidationDisplay(input);
            if (input.getAttribute('data-block-increase') === 'true') {
              const newValue = getNumericValue(input);
              if (!allowInputIncrease(input, newValue)) {
                setNumericValue(input, parseInt(input.getAttribute('data-old-value') || '0'));
                showOBLErrorToast('Cannot increase - INVOICING limit already reached. Reduce other values first.', true);
                return;
              }
            }
            if (input.getAttribute('data-old-value') !== null) {
              input.setAttribute('data-old-value', getNumericValue(input).toString());
            }
          });

          input.addEventListener('change', () => {
            updateOBLValidationDisplay(input);
          });
        });

        // ✅ Item Description inputs - keep the class in sync with the content
        document.querySelectorAll<HTMLInputElement>('input[data-item]:not([data-item-type])').forEach(input => {
          if (input.hasAttribute('readonly')) return;
          if (input.dataset.nameBound === 'true') return;
          input.dataset.nameBound = 'true';
          input.addEventListener('input', () => {
            const value = input.value.trim();
            input.className = value ? 'invoicing-input-with-value' : 'invoicing-input-empty';
          });
        });

        // ✅ Item date changes - recompute the derived month total AND the auto (highest) main Date
        document.querySelectorAll<HTMLInputElement>('input.item-date-input[data-month]').forEach(input => {
          if (input.dataset.dateBound === 'true') return;
          input.dataset.dateBound = 'true';
          input.addEventListener('change', () => {
            const monthKey = input.getAttribute('data-month') || '';
            validateSingleMonthItems(monthKey, input);
            syncAutoMonthDate(monthKey);
            updateOBLValidationDisplay(input);
          });
        });
      };

      // ✅ Auto date: the readonly top-level Date mirrors the HIGHEST item date of that month.
      const syncAutoMonthDate = (monthKey: string): void => {
        if (!monthKey) return;

        const mainDateInput = document.querySelector<HTMLInputElement>(`input.main-date-input[data-month="${monthKey}"]`);
        if (!mainDateInput) return;

        let latest = '';
        document.querySelectorAll<HTMLInputElement>(`input.item-date-input[data-month="${monthKey}"]`).forEach(itemDateInput => {
          const value = itemDateInput.value || '';
          if (value && (!latest || value > latest)) latest = value;
        });

        mainDateInput.value = latest;
        mainDateInput.setAttribute('readonly', 'true');
      };

      const validateForm = (): boolean => {
        const oblValue = getNumericValue(oblLastMonthInput);
        
        // For new projects, OBL must be > 0
        if (!editingId && oblValue <= 0) {
          toast("error", "Please enter a valid OBL value for the new project");
          oblLastMonthInput.focus();
          return false;
        }
        
        // Check OBL limit validation
        if (!validateOBLLimit()) {
          toast("error", "Cannot save - INVOICING limit exceeded. Please reduce current values.");
          return false;
        }
        
        // Check month item totals
        if (!validateMonthItemTotals()) {
          toast("error", "Cannot save - Item totals exceed month limits.");
          return false;
        }
        
        return true;
      };

      const saveForecast = async (e: Event): Promise<void> => {
        e.preventDefault();

        // ✅ Role gate (same as the sibling modules): creating needs create rights,
        //    editing needs edit rights on that project.
        const targetProject = editingId
          ? projects.find(p => p.Title === forecasts.find(f => f.Id === editingId)?.Project?.Title)
          : projects.find(p => p.Id === parseInt(projectSelect.value));

        if (editingId) {
          if (!canEditProject(targetProject)) {
            toast("error", "❌ You don't have permission to edit this invoicing forecast");
            return;
          }
        } else if (!canCreateRecord()) {
          toast("error", "❌ You don't have permission to create invoicing forecasts");
          return;
        }

        if (!validateForm()) return;
        
        // ✅ Bar-only loader (no text on screen) - the message goes to the console
        const savingLoader = createLoadingOverlay(editingId ? "Updating the invoicing forecast..." : "Creating the invoicing forecast...");
        loaderHost().appendChild(savingLoader);
        
        try {
          const digest = await getDigest();
          const oblValue = getNumericValue(oblLastMonthInput);

          // ✅ Storage months come from the module's own generator (same as the sibling
          //    modules): 3 historical + 3 detailed + 15 long-term = the 21 rendered months.
          const allStorageMonths = getAllMonthsForStorage();
          let projectId: number;
          let projectTitle: string;
          
          if (editingId) {
            const existing = forecasts.find(f => f.Id === editingId);
            projectId = existing?.Project?.Id || 0;
            projectTitle = existing?.Project?.Title || "";
          } else {
            projectId = parseInt(projectSelect.value);
            const project = projects.find(p => p.Id === projectId);
            projectTitle = project?.Title || "";
          }

          // Build forecast data object
          const forecastData: ForecastData = {};
          
          // Historical months
          allStorageMonths.slice(0, 3).forEach(monthKey => {
            const planInput = document.querySelector<HTMLInputElement>(`input.main-plan-input[data-month="${monthKey}"]`);
            const actualInput = document.querySelector<HTMLInputElement>(`input.main-current-input[data-month="${monthKey}"]`);
            const items: { [key: string]: ForecastItem } = {};
            
            Array.from({length: 10}, (_, i) => i + 1).forEach(itemNum => {
              const nameInput = document.querySelector<HTMLInputElement>(`input[data-item="${itemNum}"][data-month="${monthKey}"]:not([data-item-type])`);
              const planInputItem = document.querySelector<HTMLInputElement>(`input[data-item-type="plan"][data-item="${itemNum}"][data-month="${monthKey}"]`);
              const actualInputItem = document.querySelector<HTMLInputElement>(`input[data-item-type="actual"][data-item="${itemNum}"][data-month="${monthKey}"]`);
              
              items[itemNum.toString()] = {
                name: nameInput?.value || '',
                plan: getNumericValue(planInputItem as HTMLInputElement),
                actual: getNumericValue(actualInputItem as HTMLInputElement),
                forecast: 0,
                date: ''
              };
            });
            
            forecastData[monthKey] = {
              plan: getNumericValue(planInput as HTMLInputElement),
              actual: getNumericValue(actualInput as HTMLInputElement),
              forecast: 0,
              date: '',
              items
            };
          });
          
          // Detailed months (3-5)
          allStorageMonths.slice(3, 6).forEach(monthKey => {
            const planInput = document.querySelector<HTMLInputElement>(`input.main-plan-input[data-month="${monthKey}"]`);
            const actualInput = document.querySelector<HTMLInputElement>(`input.main-current-input[data-month="${monthKey}"]`);
            const dateInput = document.querySelector<HTMLInputElement>(`input.main-date-input[data-month="${monthKey}"]`);
            const items: { [key: string]: ForecastItem } = {};
            
            Array.from({length: 10}, (_, i) => i + 1).forEach(itemNum => {
              const nameInput = document.querySelector<HTMLInputElement>(`input[data-item="${itemNum}"][data-month="${monthKey}"]:not([data-item-type])`);
              const planInputItem = document.querySelector<HTMLInputElement>(`input[data-item-type="plan"][data-item="${itemNum}"][data-month="${monthKey}"]`);
              const actualInputItem = document.querySelector<HTMLInputElement>(`input[data-item-type="actual"][data-item="${itemNum}"][data-month="${monthKey}"]`);
              const dateInputItem = document.querySelector<HTMLInputElement>(`input[data-item-type="date"][data-item="${itemNum}"][data-month="${monthKey}"]`);
              
              items[itemNum.toString()] = {
                name: nameInput?.value || '',
                plan: getNumericValue(planInputItem as HTMLInputElement),
                actual: getNumericValue(actualInputItem as HTMLInputElement),
                forecast: 0,
                date: dateInputItem?.value || ''
              };
            });
            
            forecastData[monthKey] = {
              plan: getNumericValue(planInput as HTMLInputElement),
              actual: getNumericValue(actualInput as HTMLInputElement),
              forecast: 0,
              date: dateInput?.value || '',
              items
            };
          });
          
          // Long-term months (6+)
          allStorageMonths.slice(6).forEach(monthKey => {
            const planInput = document.querySelector<HTMLInputElement>(`input.main-plan-input[data-month="${monthKey}"]`);
            const actualInput = document.querySelector<HTMLInputElement>(`input.main-current-input[data-month="${monthKey}"]`);
            
            forecastData[monthKey] = {
              plan: getNumericValue(planInput as HTMLInputElement),
              actual: getNumericValue(actualInput as HTMLInputElement),
              forecast: 0
            };
          });
          
          // Yearly data
          const yearlyData: YearlyForecastData = {};
          document.querySelectorAll<HTMLInputElement>('input.yearly-plan-input-inv[data-year], input.yearly-current-input-inv[data-year]').forEach(input => {
            const year = input.getAttribute('data-year')!;
            if (!yearlyData[year]) yearlyData[year] = { plan: 0, actual: 0 };
            if (input.classList.contains('yearly-plan-input-inv')) {
              yearlyData[year].plan = getNumericValue(input);
            } else {
              yearlyData[year].actual = getNumericValue(input);
            }
          });
          forecastData.yearly = yearlyData;

          // Save to SharePoint
          const itemData: any = {
            __metadata: { type: "SP.Data.InvoicingPlanningListItem" },
            Title: generateForecastTitle(projects.find(p => p.Title === projectTitle)!, forecastData),
            ProjectId: projectId,
            ProfitCentre: projectTitle ? projects.find(p => p.Title === projectTitle)?.PC || "" : "",
            OBL: oblValue,
            ForecastData: JSON.stringify(forecastData),
            InvoiceActual: JSON.stringify({})
          };

          const url = editingId 
            ? `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${editingId})`
            : `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
          
          const method = editingId ? "POST" : "POST";
          const headers: any = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest
          };
          
          if (editingId) {
            headers["IF-MATCH"] = "*";
            headers["X-HTTP-Method"] = "MERGE";
          }

          const response = await fetch(url, {
            method: method,
            headers: headers,
            body: JSON.stringify(itemData)
          });

          if (!response.ok) {
            const errorText = await response.text();
            console.error("❌ Save failed:", errorText);
            throw new Error(`Save failed: ${response.status}`);
          }

          const result = await response.json();
          console.log("✅ Invoicing forecast saved:", result);
          
          onSaveSuccess();
          
        } catch (error: any) {
          console.error("❌ Error saving forecast:", error);
          toast("error", error.message || "Failed to save forecast");
        } finally {
          removeLoadingOverlay();
        }
      };

      // ========== EDIT/DELETE FUNCTIONS ==========

      const editForecast = (id: number): void => {
        const forecast = forecasts.find(f => f.Id === id);
        if (!forecast) {
          toast("error", "Forecast not found");
          return;
        }

        const project = projects.find(p => p.Title === forecast.Project?.Title);

        // ✅ Same gate as the sibling modules
        if (!canEditProject(project)) {
          toast("error", "❌ You don't have permission to edit this invoicing forecast");
          return;
        }

        // ✅ Month progression (same as Inventory / SEIBIES): a record saved in an
        //    earlier month is rolled forward - each month's Current becomes the next
        //    month's Previous - and corrupted data falls back to a fresh structure.
        let progressedData: ForecastData;
        try {
          progressedData = safeJsonParse(forecast.ForecastData, {});
          if (!progressedData || Object.keys(progressedData).length === 0) {
            progressedData = createFreshForecastStructure();
          }
        } catch (error) {
          console.error("Corrupted invoicing forecast data:", error);
          progressedData = createFreshForecastStructure();
          toast("warning", "Corrupted invoicing forecast data detected. Loaded empty structure.");
        }
        progressedData = propagateMonthValues(progressedData);

        editingId = id;
        
        // Show form, hide table
        formSection.style.display = "block";
        tableSection.style.display = "none";
        formSection.classList.add('invoicing-edit-mode');
        projectSelect.disabled = true;
        formTitle.textContent = "Invoicing Plan 5QRF";
        
        // Set the project
        if (project) {
          projectSelect.value = project.Id.toString();
          
          // Progressively render the form (keep the loader visible until done)
          const editLoader = createLoadingOverlay("Preparing the invoicing forecast form...");
          loaderHost().appendChild(editLoader);
          
          setTimeout(() => {
            try {
              onProjectChange(project.Id, forecast, progressedData);
            } finally {
              removeLoadingOverlay();
            }
          }, 100);
        } else {
          toast("error", "Project not found in the project list");
        }
      };

      // ========== FILTER FUNCTIONS ==========

      // =====================================================================
      // ✅ FILTER ENGINE + FILTER POPUP  (same behaviour as the Inventory / SEIBIES modules)
      //
      //    appliedFilters               = what the TABLE is filtered by
      //    the seven <select> values    = the DRAFT inside the popup
      //      • Apply              → snapshot the draft into appliedFilters and re-filter
      //      • Clear              → clearFilters() (this module's own logic) + resets the
      //                             currently seeded Year / Month defaults
      //      • Cancel             → a real cancel: the applied values go back
      //      • ✕ / Esc / backdrop → close but KEEP the draft, so the compact
      //                             "Apply" pill appears (mock-up image 8)
      //    A missing element never breaks the flow (every node is optional).
      // =====================================================================

      const MONTH_NAMES_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

      // ---------------------------------------------------------------------
      // Option population. The old loadProjects() only filled #filterProject,
      // so the other six selects stayed empty (and the Excel export, which
      // requires a year + month, could never be used). This fills them all.
      // ---------------------------------------------------------------------
      const populateFilterOptions = (): void => {
        const yearSet = new Set<string>();
        yearSet.add(new Date().getFullYear().toString());

        forecasts.forEach(f => {
          const data = safeJsonParse(f.ForecastData, {});
          Object.keys(data).forEach(key => {
            if (key && /^\d{4}-\d{2}$/.test(key)) yearSet.add(key.substring(0, 4));
          });
          if (data.yearly && typeof data.yearly === "object") {
            Object.keys(data.yearly).forEach(year => yearSet.add(year));
          }
        });

        const sortedYears = Array.from(yearSet).sort().reverse();

        const currentYear = filterYear.value;
        filterYear.innerHTML = '<option value="">All Years</option>' +
          sortedYears.map(y => `<option value="${y}">${y}</option>`).join("");
        if (currentYear && yearSet.has(currentYear)) filterYear.value = currentYear;

        const currentMonth = filterMonth.value;
        filterMonth.innerHTML = '<option value="">All Months</option>' +
          MONTH_NAMES_SHORT.map((m, idx) => `<option value="${idx + 1}">${m}</option>`).join("");
        if (currentMonth) filterMonth.value = currentMonth;

        const currentCustomer = filterCustomer.value;
        filterCustomer.innerHTML = '<option value="">All Customers</option>' +
          [...new Set(projects.map(p => p.CustomerName).filter(Boolean))]
            .sort().map(c => `<option value="${c}">${c}</option>`).join("");
        if (currentCustomer) filterCustomer.value = currentCustomer;

        const currentManager = filterProjectManager.value;
        filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
          [...new Set(projects.map(p => displayNameOf(p.ProjectManager)).filter(Boolean))]
            .sort().map(m => `<option value="${m}">${m}</option>`).join("");
        if (currentManager) filterProjectManager.value = currentManager;

        const currentGroupHead = filterGroupHead.value;
        filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
          [...new Set(projects.map(p => displayNameOf(p.GroupHead1)).filter(Boolean))]
            .sort().map(g => `<option value="${g}">${g}</option>`).join("");
        if (currentGroupHead) filterGroupHead.value = currentGroupHead;

        const currentReference = filterProjectReference.value;
        filterProjectReference.innerHTML = '<option value="">All References</option>' +
          [...new Set(projects.map(p => p.ProjectReference).filter(Boolean))]
            .sort().map(r => `<option value="${r}">${r}</option>`).join("");
        if (currentReference) filterProjectReference.value = currentReference;

        // Known side effect: #filterProject is rebuilt by loadProjects(), but re-running it
        // here keeps every reload path (Back button, after save) idempotent.
        const currentProject = filterProject.value;
        if (currentProject) filterProject.value = currentProject;
      };

      // ---------------------------------------------------------------------
      // Filtering (replaces "the selects did nothing" with a real filter)
      // ---------------------------------------------------------------------
      const getProjectByForecast = (forecast: InvoicingForecastItem): ProjectItem | undefined => {
        const projectTitle = forecast.Project?.Title;
        if (!projectTitle) return undefined;
        return projects.find(p => p.Title === projectTitle);
      };

      const displayNameOf = (value: string | { Title: string } | undefined): string => {
        if (!value) return "";
        if (typeof value === "object" && value !== null) return value.Title || "";
        // ✅ cast: the object branch was handled above, so this is the plain string
        return value as string;
      };

      const forecastHasMonth = (data: any, year: string, month: string): boolean => {
        const monthNum = parseInt(month, 10);

        return Object.keys(data || {}).some(key => {
          if (!key || !/^\d{4}-\d{2}$/.test(key)) return false;

          const keyYear = key.substring(0, 4);
          if (year && keyYear !== year) return false;

          if (month) {
            const keyMonth = parseInt(key.substring(5, 7), 10);
            if (keyMonth !== monthNum) return false;
          }

          return true;
        });
      };

      const forecastHasYear = (data: any, year: string): boolean => {
        if (!year) return false;

        const hasMonthly = Object.keys(data || {}).some(key =>
          key && /^\d{4}-\d{2}$/.test(key) && key.substring(0, 4) === year
        );

        const yearlyData = data?.yearly || {};
        const hasYearly = typeof yearlyData === "object" && Object.prototype.hasOwnProperty.call(yearlyData, year);

        return hasMonthly || hasYearly;
      };

      // ✅ Role gate kept from this module: a restricted user only ever sees the
      //    forecasts of his assigned projects (canReadAllProjects() opens it up).
      const forecastIsVisibleToUser = (forecast: InvoicingForecastItem): boolean => {
        if (canReadAllProjects()) return true;
        const projectTitle = forecast.Project?.Title || "";
        if (!projectTitle) return false;
        return getAssignedProjects(projects).some(p => p.Title === projectTitle);
      };

      const filterTable = (): void => {
        const selectedYear = filterYear.value;
        const selectedMonth = filterMonth.value;
        const selectedProject = filterProject.value;
        const selectedCustomer = filterCustomer.value;
        const selectedManager = filterProjectManager.value;
        const selectedGroupHead = filterGroupHead.value;
        const selectedReference = filterProjectReference.value;

        const visibleForecasts = forecasts.filter(forecastIsVisibleToUser);

        const noFilters = !selectedYear && !selectedMonth && !selectedProject &&
          !selectedCustomer && !selectedManager && !selectedGroupHead && !selectedReference;

        if (noFilters) {
          initializePagination(visibleForecasts);
          return;
        }

        const filtered = visibleForecasts.filter(f => {
          if (selectedProject && f.Project?.Title !== selectedProject) return false;

          const proj = getProjectByForecast(f);
          if (proj) {
            if (selectedCustomer && (proj.CustomerName || "") !== selectedCustomer) return false;
            if (selectedManager && displayNameOf(proj.ProjectManager) !== selectedManager) return false;
            if (selectedGroupHead && displayNameOf(proj.GroupHead1) !== selectedGroupHead) return false;
            if (selectedReference && (proj.ProjectReference || "") !== selectedReference) return false;
          } else if (selectedCustomer || selectedManager || selectedGroupHead || selectedReference) {
            return false;
          }

          // Year + Month together = "this record holds that exact month" (the module's
          // original rule); one on its own still narrows the list sensibly.
          if (selectedYear || selectedMonth) {
            const data = safeJsonParse(f.ForecastData, {});
            if (selectedYear && selectedMonth) {
              if (!forecastHasMonth(data, selectedYear, selectedMonth)) return false;
            } else if (selectedYear) {
              if (!forecastHasYear(data, selectedYear)) return false;
            } else if (selectedMonth) {
              if (!forecastHasMonth(data, "", selectedMonth)) return false;
            }
          }

          return true;
        });

        if (filtered.length === 0) {
          // No results: clear the totals band and show the empty state
          const totalsHeader = document.querySelector("#totalsHeader") as HTMLElement | null;
          if (totalsHeader) totalsHeader.innerHTML = '';
          initializePagination([]);
          return;
        }

        initializePagination(filtered);
      };

      // ---------------------------------------------------------------------
      // Popup / toolbar state
      // ---------------------------------------------------------------------
      const filterTriggerBtn = document.querySelector("#filterTriggerInvoicing") as HTMLButtonElement | null;
      const filterOverlay = document.querySelector("#invoicingFilterOverlay") as HTMLElement | null;
      const filterModalCloseBtn = document.querySelector("#filterModalCloseInvoicing") as HTMLButtonElement | null;
      const filterCancelBtn = document.querySelector("#filterCancelInvoicing") as HTMLButtonElement | null;
      const filterClearModalBtn = document.querySelector("#filterClearModalInvoicing") as HTMLButtonElement | null;
      const filterApplyModalBtn = document.querySelector("#filterApplyModalInvoicing") as HTMLButtonElement | null;
      const floatingApplyBtn = document.querySelector("#floatingFilterApplyInvoicing") as HTMLButtonElement | null;
      const filterCountBadge = document.querySelector("#filterCountInvoicing") as HTMLElement | null;
      const filterSummaryBox = document.querySelector("#filterSummaryInvoicing") as HTMLElement | null;

      let appliedFilters: { [key: string]: string } = {};

      const FILTER_FIELDS = ['project', 'customer', 'manager', 'groupHead', 'reference', 'year', 'month'];

      const FILTER_SELECTS: { [key: string]: HTMLSelectElement } = {
        project: filterProject,
        customer: filterCustomer,
        manager: filterProjectManager,
        groupHead: filterGroupHead,
        reference: filterProjectReference,
        year: filterYear,
        month: filterMonth
      };

      const FILTER_LABELS: { [key: string]: string } = {
        project: 'Project',
        customer: 'Customer',
        manager: 'Manager',
        groupHead: 'Group Head',
        reference: 'Reference',
        year: 'Year',
        month: 'Month'
      };

      const isFilterModalOpen = (): boolean =>
        !!filterOverlay && filterOverlay.classList.contains('is-open');

      const readDraftFilters = (): { [key: string]: string } => {
        const out: { [key: string]: string } = {};
        FILTER_FIELDS.forEach(field => {
          const el = FILTER_SELECTS[field];
          if (el && el.value) out[field] = el.value;
        });
        return out;
      };

      const sameFilters = (a: { [key: string]: string }, b: { [key: string]: string }): boolean =>
        FILTER_FIELDS.every(field => (a[field] || '') === (b[field] || ''));

      const countFilters = (source: { [key: string]: string }): number =>
        FILTER_FIELDS.filter(field => !!source[field]).length;

      const syncFilterFieldStyles = (): void => {
        FILTER_FIELDS.forEach(field => {
          const el = FILTER_SELECTS[field];
          if (el) el.classList.toggle('is-set', !!el.value);
        });
      };

      // Apply/Clear enable only when there is something to do; the floating Apply
      // pill shows only while there are unapplied changes and the popup is closed
      const syncFilterButtons = (): void => {
        const draft = readDraftFilters();
        const dirty = !sameFilters(draft, appliedFilters);
        const anyDraft = countFilters(draft) > 0;
        const anyApplied = countFilters(appliedFilters) > 0;

        if (filterApplyModalBtn) {
          filterApplyModalBtn.disabled = !dirty;
          filterApplyModalBtn.classList.toggle('is-enabled', dirty);
        }
        if (filterClearModalBtn) {
          const canClear = anyDraft || anyApplied;
          filterClearModalBtn.disabled = !canClear;
          filterClearModalBtn.classList.toggle('is-enabled', canClear);
        }
        if (floatingApplyBtn) {
          floatingApplyBtn.classList.toggle('is-visible', dirty && !isFilterModalOpen());
        }
      };

      const syncFilterTrigger = (): void => {
        const n = countFilters(appliedFilters);
        if (filterCountBadge) filterCountBadge.textContent = String(n);
        if (filterTriggerBtn) filterTriggerBtn.classList.toggle('is-active', n > 0);
      };

      // chips above the table: one per applied filter, each removable
      const renderFilterChips = (): void => {
        if (!filterSummaryBox) return;

        const activeFields = FILTER_FIELDS.filter(field => !!appliedFilters[field]);

        if (activeFields.length === 0) {
          filterSummaryBox.style.display = 'none';
          filterSummaryBox.innerHTML = '';
          return;
        }

        filterSummaryBox.style.display = 'flex';
        filterSummaryBox.innerHTML =
          '<span class="invoicing-filter-summary__label">Active filters:</span>' +
          activeFields.map(field => {
            const text = field === 'month'
              ? `Month: ${MONTH_NAMES_SHORT[parseInt(appliedFilters[field], 10) - 1] || appliedFilters[field]}`
              : `${FILTER_LABELS[field]}: ${appliedFilters[field]}`;
            return `<span class="invoicing-filter-chip">${text}<button type="button" data-clear-filter="${field}" title="Remove this filter">&times;</button></span>`;
          }).join('') +
          '<button type="button" class="invoicing-filter-chip__all" id="filterChipsClearAllInvoicing">Clear all</button>';

        filterSummaryBox.querySelectorAll<HTMLButtonElement>('[data-clear-filter]').forEach(btn => {
          btn.addEventListener('click', () => {
            const field = btn.dataset.clearFilter as string;
            const el = FILTER_SELECTS[field];
            if (el) el.value = '';
            delete appliedFilters[field];
            syncFilterFieldStyles();
            syncFilterTrigger();
            renderFilterChips();
            syncFilterButtons();
            filterTable();
          });
        });

        const clearAllBtn = filterSummaryBox.querySelector('#filterChipsClearAllInvoicing') as HTMLButtonElement | null;
        if (clearAllBtn) {
          clearAllBtn.addEventListener('click', () => {
            clearFilters();                  // this module's own logic (selects + sort + table)
            syncFilterUIFromSelects();
          });
        }
      };

      // after a Clear / Reset the selects are the source of truth again
      const syncFilterUIFromSelects = (): void => {
        appliedFilters = readDraftFilters();
        syncFilterFieldStyles();
        syncFilterTrigger();
        renderFilterChips();
        syncFilterButtons();
      };

      const openFilterModal = (): void => {
        if (!filterOverlay) return;

        // the popup always opens showing what is currently applied
        FILTER_FIELDS.forEach(field => {
          const el = FILTER_SELECTS[field];
          if (el) el.value = appliedFilters[field] || '';
        });

        syncFilterFieldStyles();
        syncFilterButtons();
        filterOverlay.classList.add('is-open');
        if (filterTriggerBtn) {
          filterTriggerBtn.classList.add('is-open');
          filterTriggerBtn.setAttribute('aria-expanded', 'true');
        }
        document.body.style.overflow = 'hidden';
      };

      const closeFilterModal = (restoreApplied: boolean = true): void => {
        if (!filterOverlay) return;

        if (restoreApplied) {
          // real cancel: discard the draft, show the applied values again
          FILTER_FIELDS.forEach(field => {
            const el = FILTER_SELECTS[field];
            if (el) el.value = appliedFilters[field] || '';
          });
        }

        filterOverlay.classList.remove('is-open');
        if (filterTriggerBtn) {
          filterTriggerBtn.classList.remove('is-open');
          filterTriggerBtn.setAttribute('aria-expanded', 'false');
        }
        document.body.style.overflow = '';
        syncFilterFieldStyles();
        syncFilterButtons();
      };

      const applyFilterModal = (): void => {
        appliedFilters = readDraftFilters();
        syncFilterTrigger();
        renderFilterChips();
        syncFilterButtons();
        filterTable();                 // real filtering + totals + pagination
        closeFilterModal(false);       // keep the values that were just applied
        const n = countFilters(appliedFilters);
        toast("info", n > 0 ? `${n} invoicing filter(s) applied` : "Invoicing filters cleared");
      };

      const clearFilterDraft = (): void => {
        clearFilters();                // this module's own clear logic, popup stays open
        syncFilterUIFromSelects();
      };

      // ---------------------------------------------------------------------
      // ✅ Kept for compatibility: the rest of this module (clearFilters(),
      //    Back button, after-save reloads) calls initializeFilters().
      //
      //    It only (re)fills the seven dropdowns. NOTHING is filtered on load -
      //    the Filters button starts with no badge, no chips and the table shows
      //    every row until something is applied in the popup.
      // ---------------------------------------------------------------------
      const initializeFilters = (): void => {
        populateFilterOptions();
      };

      // ========== PAGINATION EVENT LISTENERS ==========

      const setupPaginationEventListeners = (): void => {
        if (firstPageBtn) firstPageBtn.addEventListener('click', () => goToPage(1));
        if (prevPageBtn) prevPageBtn.addEventListener('click', () => goToPage(currentPage - 1));
        if (nextPageBtn) nextPageBtn.addEventListener('click', () => goToPage(currentPage + 1));
        if (lastPageBtn) lastPageBtn.addEventListener('click', () => goToPage(totalPages));
      };

      // ✅ One reload path for every "the data may have changed" case
      //    (Back button, after a save, manual refresh): projects + forecasts,
      //    then re-fill the dropdowns and re-run the filter the user has applied.
      const reloadInvoicingData = async (): Promise<void> => {
        const loading = createLoadingOverlay("Reloading invoicing data...");
        loaderHost().appendChild(loading);

        try {
          await loadProjects();
          await loadForecasts();
          populateFilterOptions();
          syncFilterUIFromSelects();
          filterTable();
        } catch (error) {
          console.error('❌ Reload failed:', error);
          toast("error", "Reload failed - showing current data");
        } finally {
          removeLoadingOverlay();
        }
      };

      (window as any).reloadInvoicingData = reloadInvoicingData;

      // ========== EVENT LISTENERS ==========

      exportBtn?.addEventListener("click", exportToExcelWithFullStyle);
      exportDetailsBtn?.addEventListener("click", exportDetailedItemsToExcel);

      // ============================================================
// ✅ SAVE — SweetAlert confirmation FIRST (same look as procurement.ts)
//    • 2 buttons only: Yes, Save / Cancel  (no "Don't Save" option)
//    • z-index bumped so the loader cannot cover the dialog
//    • isSavingInvoicing guard prevents double-submit
//    • saveBtns re-enabled in the outer finally, so no path can leave them stuck
// ============================================================
form.addEventListener("submit", async (e: Event) => {
  e.preventDefault();

  // ✅ Only run the flow when a Save button triggered the submit
  const submitter = (e as any).submitter as HTMLButtonElement | null;
  const isSaveClick = submitter && submitter.classList.contains("invoicing-btn-primary");
  if (!isSaveClick) return;

  if (isSavingInvoicing) {
    console.log("⛔ submit blocked: save already in progress");
    return;
  }

  // ---- Pre-checks (nothing shown yet) ----
  const targetProject = editingId
    ? projects.find(p => p.Title === forecasts.find(f => f.Id === editingId)?.Project?.Title)
    : projects.find(p => p.Id === parseInt(projectSelect.value));

  if (!targetProject) {
    toast("error", "Please select a project");
    return;
  }

  if (editingId) {
    if (!canEditProject(targetProject)) {
      toast("error", "❌ You don't have permission to edit this invoicing forecast");
      return;
    }
  } else if (!canCreateRecord()) {
    toast("error", "❌ You don't have permission to create invoicing forecasts");
    return;
  }

  if (!validateForm()) return;

  // ✅ HOIST the save buttons so the outer finally can always restore them
  const saveBtns = Array.from(
    document.querySelectorAll<HTMLButtonElement>('button.invoicing-btn-primary[type="submit"]')
  );

  isSavingInvoicing = true;
  try {
    // ========== CONFIRMATION POPUP ==========
    const confirmResult = await Swal.fire({
      title: editingId ? "Save Changes?" : "Save Invoicing Forecast?",
      html: editingId
        ? `Do you want to save the changes made to <strong>${escapeHtml(targetProject.Title)}</strong>?`
        : `Do you want to save the Invoicing forecast for <strong>${escapeHtml(targetProject.Title)}</strong>?`,
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
      return; // finally will reset the guard + buttons
    }

    // ========== LOCK SAVE BUTTONS ==========
    saveBtns.forEach(btn => {
      btn.disabled = true;
      btn.dataset.saving = "true";
    });

    // ========== ACTUAL SAVE (unchanged logic) ==========
    await saveForecast(e);

  } catch (outerError) {
    console.error("Unexpected error in Invoicing submit:", outerError);
    const message = outerError instanceof Error ? outerError.message : String(outerError);
    toast("error", `Unexpected error: ${message}`);
  } finally {
    // ✅ Always restore the buttons + clear the guard
    saveBtns.forEach(btn => {
      btn.disabled = false;
      btn.dataset.saving = "false";
    });
    isSavingInvoicing = false;
  }
});

  // ============================================================
// ✅ BACK — SweetAlert confirmation, same look as procurement.ts
//    • Only one prompt: Leave (Yes, Go Back / Stay)
//    • Only shown when the form is actually open
//    • z-index bumped so the loader cannot cover the dialog 
// ============================================================
if (backBtn) {
  backBtn.addEventListener("click", async (e) => {
    e.preventDefault();
    e.stopPropagation();
 
    if (isLeavingInvoicingForm) {
      console.log("⛔ back blocked: already leaving");
      return;
    }

    // Nothing to confirm when the form is already hidden
    const formVisible = formSection && formSection.style.display !== "none";
    if (!formVisible) {
      editingId = null;
      hideForm();
      reloadInvoicingData();
      return;
    }

    isLeavingInvoicingForm = true;
    try {
      const confirmResult = await Swal.fire({
        title: "Leave this page?",
        html:
          "Any <strong>unsaved changes</strong> will be lost.<br>" +
          "Do you want to go back to the Invoicing list?",
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

      editingId = null;
      hideForm();
      reloadInvoicingData();
    } catch (err) {
      console.error("Error while leaving the Invoicing form:", err);
      const message = err instanceof Error ? err.message : String(err);
      toast("error", `Unexpected error: ${message}`);
    } finally {
      isLeavingInvoicingForm = false;
    }
  });
}
// ✅ Select ALL back buttons
const backBtns = document.querySelectorAll(".back-to-table");

if (backBtns && backBtns.length > 0) {
  backBtns.forEach((btn) => {
    btn.addEventListener("click", async (e) => {
      e.preventDefault();
      e.stopPropagation();

      if (isLeavingInvoicingForm) {
        console.log("⛔ back blocked: already leaving");
        return;
      }

      // Nothing to confirm when the form is already hidden
      const formVisible = formSection && formSection.style.display !== "none";
      if (!formVisible) {
        editingId = null;
        hideForm();
        reloadInvoicingData();
        return;
      }

      isLeavingInvoicingForm = true;
      try {
        const confirmResult = await Swal.fire({
          title: "Leave this page?",
          html:
            "Any <strong>unsaved changes</strong> will be lost.<br>" +
            "Do you want to go back to the Invoicing list?",
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

        editingId = null;
        hideForm();
        reloadInvoicingData();
      } catch (err) {
        console.error("Error while leaving the Invoicing form:", err);
        const message = err instanceof Error ? err.message : String(err);
        toast("error", `Unexpected error: ${message}`);
      } finally {
        isLeavingInvoicingForm = false;
      }
    });
  });
}
      projectSelect.addEventListener("change", (e: Event) => {
        const projectId = parseInt((e.target as HTMLSelectElement).value);

        if (projectId) {
          onProjectChange(projectId);
        } else {
          toggleForecastSections(false);
        }
      });

      // ✅ The filter selects no longer do anything on their own - they are inside
      //    the popup and only Apply (or the floating Apply pill) filters the table.
      //    Picking a value just enables Apply / marks the field as set.
      FILTER_FIELDS.forEach(field => {
        const el = FILTER_SELECTS[field];
        if (el) {
          el.addEventListener("change", () => {
            syncFilterFieldStyles();
            syncFilterButtons();
          });
        }
      });

      if (filterTriggerBtn) {
        filterTriggerBtn.addEventListener("click", () => {
          isFilterModalOpen() ? closeFilterModal() : openFilterModal();
        });
      }
      // ✕ keeps the draft (the floating Apply pill then offers to apply it)
      if (filterModalCloseBtn) filterModalCloseBtn.addEventListener("click", () => closeFilterModal(false));
      if (filterCancelBtn) filterCancelBtn.addEventListener("click", () => closeFilterModal());
      if (filterApplyModalBtn) filterApplyModalBtn.addEventListener("click", () => { if (!filterApplyModalBtn.disabled) applyFilterModal(); });
      if (filterClearModalBtn) filterClearModalBtn.addEventListener("click", () => { if (!filterClearModalBtn.disabled) clearFilterDraft(); });
      if (floatingApplyBtn) floatingApplyBtn.addEventListener("click", applyFilterModal);
      if (filterOverlay) {
        filterOverlay.addEventListener("mousedown", (e: MouseEvent) => {
          if (e.target === filterOverlay) closeFilterModal(false);   // keep the draft
        });
      }
      document.addEventListener("keydown", (e: KeyboardEvent) => {
        // Esc = light dismiss: keep the draft, the floating Apply pill stays available
        if (e.key === "Escape" && isFilterModalOpen()) closeFilterModal(false);
      });

      setupPaginationEventListeners();

      // ========== INITIAL LOADS ==========

      try {
        await authManager.initialize({ siteUrl });

        console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
        console.log(`👤 User: ${getCurrentUserFullName()}`);
        console.log(`📋 User Groups:`, getUserGroups());
        console.log(`✅ Full Control: ${hasFullControl()} | Can Contribute: ${canContribute()} | Can Read: ${canRead()}`);
        console.log(`✅ Can Read All Projects: ${canReadAllProjects()} | Can Freeze: ${canFreeze()}`);
        console.log(`✅ Is GH Head: ${isGHHead()} | Is PM Manager: ${isPMManager()} | Is Engg User: ${isEngineeringUser()}`);

        await loadProjects();
        await loadForecasts();

        setupPaginationEventListeners();

        // ✅ fill all seven filter dropdowns and paint the toolbar state
        //    (badge 0, no chips, Apply disabled) - no filter is applied on load
        initializeFilters();
        syncFilterUIFromSelects();
        filterTable();

        updateUIBasedOnPermissions({
          exportBtn: ".invoicing-export-btn",
          saveBtn: ".invoicing-btn-primary[type=\"submit\"]"
        });

      } catch (error) {
        console.error('❌ Error initializing Invoicing module:', error);
        toast("error", "Failed to initialize application");
      }

      console.log("✅ Invoicing Forecast Module Initialized - Default page size: 25 per page");

    }, 100);
  }
};

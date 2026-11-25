import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/seibies.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canReadAllProjects, getAssignedProjects, isPMManager, isEngineeringUser, updateUIBasedOnPermissions } from "./auth";
let currentSortConfig = {
    column: '',
    direction: 'asc'
};
export default {
    content: `
<div class="seibies-page fade-in">

  <!-- ============================================ -->
  <!-- LOADER - covers ONLY this section              -->
  <!-- ✅ bar-only page loader (grey track + indigo bar, no text) -->
  <!-- ============================================ -->
  <div id="revenuePageLoader" class="revenue-loader" style="position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;background:rgba(248,249,250,0.92);z-index:999;">
    <div class="seibies-loader-panel">
      <div class="seibies-page-loader">
        <div class="seibies-page-loader__track">
          <div class="seibies-page-loader__bar"></div>
        </div>
      </div>
    </div>
  </div>
  <section class="table-section-seibies card-seibies">
    <div class="page-header">
      <h1>SEI BIES Liquidation 5QRF</h1>
    </div>

    <!-- ✅ TOOLBAR ABOVE THE TABLE - the Filters popup trigger + its Apply pill (left)
         and Download Excel (right) live here (NOT next to the page title any more).
         There is no Reset button in the toolbar: the popup's Clear button resets
         every filter. The seven filter dropdowns themselves moved inside the Filters
         popup below, keeping their original ids so clearFilters() / the export logic
         and every existing listener keep working. -->
    <div class="table-toolbar">
  <!-- LEFT: Filters trigger -->
  <div class="table-toolbar__left"> 
    <button type="button" class="btn-secondary-seibies seibies-filter-trigger" id="filterTriggerSeibies"
            aria-haspopup="dialog" aria-expanded="false" title="Filter SEI-BIES">
      <svg class="seibies-filter-funnel" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8v5l-4 2v-7L3 5z"/></svg>
      <span class="seibies-filter-label">Filters</span>
      <span class="seibies-filter-count" id="filterCountSeibies">0</span>
      <svg class="seibies-filter-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>
    </button>
  </div>
 <!-- RIGHT: Note -->
  <div class="table-toolbar__center">
    <span class="table-toolbar__note">Note: "Negative=BIES, Positive=SEI"</span>
  </div>
  <!-- MIDDLE: Apply pill + Download Excel -->
  <div class="table-toolbar__right">
    <button type="button" class="btn-primary-seibies seibies-filter-applypill" id="floatingFilterApplySeibies">Apply</button>
    <button type="button" class="btn-secondary-seibies export-btn-seibies">📊 Download Excel</button>
  </div>

 
</div>

   
    <div class="table-container-seibies">
      <div class="table-scroll-wrapper-seibies">
        <table class="month-columns-table-seibies">
          <thead id="totalsHeaderSeibies" class="sticky-totals-header-seibies"></thead>
          <thead id="monthColumnsHeaderSeibies" class="sticky-main-header-seibies"></thead>
          <tbody class="seibies-table-body">
            <!-- Static loading row - shows the indeterminate page-loader bar
                 (grey track + indigo bar, no text) until the JS render replaces it -->
            <tr class="seibies-table-loading-row">
              <td colspan="25">
                <div class="seibies-page-loader">
                  <div class="seibies-page-loader__track">
                    <div class="seibies-page-loader__bar"></div>
                  </div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- ✅ PAGINATION - matches the Inventory mock-up (one flat bar under the table)
           LEFT  = Items per page: [n]   Showing 1 - 25 of 60 items
           RIGHT = «  ‹  [page tiles] of N  ›  »
           Every id below is unchanged (#pageSizeSelectSeibies, #startItemSeibies,
           #endItemSeibies, #totalItemsSeibies, #pageNumbersSeibies,
           #totalPagesSeibies, #first/#prev/#next/#lastPageSeibies), so
           updatePaginationControls / updatePageNumbers / changePageSize / goToPage
           + setupPaginationEventListeners keep working as-is.
           The "Page X of Y" pill is gone (it is now the "of N" after the tiles) -
           see edit 9 for the two-line guard. -->
      <div class="pagination-container-seibies">
        <!-- LEFT: items per page + showing info -->
        <div class="pagination-left-seibies">
          <div class="page-size-selector-seibies seibies-pg-size">
            <label for="pageSizeSelectSeibies">Items per page:</label>
            <select id="pageSizeSelectSeibies" class="page-size-select-seibies">
              <option value="5">5</option>
              <option value="10">10</option>
              <option value="25" selected>25</option>
              <option value="50">50</option>
              <option value="100">100</option>
            </select>
          </div>

          <div class="pagination-info-seibies seibies-pg-showing">
            Showing <span id="startItemSeibies">0</span> - <span id="endItemSeibies">0</span> of <span id="totalItemsSeibies">0</span> items
          </div>
        </div>

        <!-- RIGHT: «  ‹  [page tiles] of N  ›  » -->
        <div class="pagination-right-seibies">
          <button class="pagination-btn-seibies seibies-pg-icon" id="firstPageSeibies" title="First page">«</button>
          <button class="pagination-btn-seibies seibies-pg-icon" id="prevPageSeibies" title="Previous page">‹</button>

          <div class="pagination-pages-seibies" id="pageNumbersSeibies"></div>

          <span class="seibies-pg-of">of <span id="totalPagesSeibies">1</span></span>

          <button class="pagination-btn-seibies seibies-pg-icon" id="nextPageSeibies" title="Next page">›</button>
          <button class="pagination-btn-seibies seibies-pg-icon" id="lastPageSeibies" title="Last page">»</button>
        </div>
      </div>
    </div>
  </section>

  <!-- =====================================================================
       ✅ FILTER POPUP (mock-up image 7: header · fields · Cancel / Clear / Apply)
       The seven selects keep their ORIGINAL ids (#filterProjectSeibies,
       #filterCustomerSeibies, #filterProjectManagerSeibies,
       #filterGroupHeadSeibies, #filterProjectReferenceSeibies,
       #filterYearSeibies, #filterMonthSeibies) so clearFilters(), the export
       (filterYear/filterMonth) and every existing listener keep working.
       Footer: Cancel (revert the draft) · Clear (RESET every filter, kept inside
       the popup per request) · Apply.
       ===================================================================== -->
  <div class="seibies-filter-overlay" id="seibiesFilterOverlay" role="dialog" aria-modal="true" aria-labelledby="seibiesFilterTitle">
    <div class="seibies-filter-modal">
      <div class="seibies-filter-modal__head">
        <h2 id="seibiesFilterTitle">Filters</h2>
        <button type="button" class="seibies-filter-modal__close" id="filterModalCloseSeibies" title="Close" aria-label="Close">✕</button>
      </div>

      <div class="seibies-filter-modal__body">
        <div class="seibies-filter-grid">
          <div class="seibies-filter-field seibies-filter-field--wide">
            <label for="filterProjectSeibies">Project</label>
            <select id="filterProjectSeibies" data-filter="project"><option value="">All Projects</option></select>
          </div>
          <div class="seibies-filter-field">
            <label for="filterCustomerSeibies">Customer</label>
            <select id="filterCustomerSeibies" data-filter="customer"><option value="">All Customers</option></select>
          </div>
          <div class="seibies-filter-field">
            <label for="filterProjectManagerSeibies">Project Manager</label>
            <select id="filterProjectManagerSeibies" data-filter="manager"><option value="">All Managers</option></select>
          </div>
          <div class="seibies-filter-field">
            <label for="filterGroupHeadSeibies">Group Head</label>
            <select id="filterGroupHeadSeibies" data-filter="groupHead"><option value="">All Group Heads</option></select>
          </div>
          <div class="seibies-filter-field">
            <label for="filterProjectReferenceSeibies">Project Reference</label>
            <select id="filterProjectReferenceSeibies" data-filter="reference"><option value="">All References</option></select>
          </div>
          <div class="seibies-filter-field">
            <label for="filterYearSeibies">Year</label>
            <select id="filterYearSeibies" data-filter="year"><option value="">All Years</option></select>
          </div>
          <div class="seibies-filter-field">
            <label for="filterMonthSeibies">Month</label>
            <select id="filterMonthSeibies" data-filter="month"><option value="">All Months</option></select>
          </div>
        </div>
      </div>

      <div class="seibies-filter-modal__foot">
        <button type="button" class="seibies-filter-btn-text" id="filterCancelSeibies">Cancel</button>
        <button type="button" class="seibies-filter-btn-clear" id="filterClearModalSeibies" title="Reset all filters" disabled>Clear</button>
        <button type="button" class="seibies-filter-btn-apply" id="filterApplyModalSeibies" disabled>Apply</button>
      </div>
    </div>
  </div>

  <section class="form-section-seibies card-seibies" style="display:none;">
  <form id="seibiesForm">
  <div class="page-header"> 
  <h1 id="formTitleSeibies">SEI BIES Liquidation 5QRF</h1>
  
  <!-- Wrap the buttons here -->
  <div class="header-actions">
    <button type="submit"     
            form="seibiesForm"
            class="btn-primary-seibies btn-save-top-seibies"
            title="Save SEIBIES Forecast">💾 Save</button>
    <button type="button" class="btn-secondary-seibies back-to-table-seibies">⬅ Back</button>
  </div>
</div>
  
      
      <!-- Project Selection (hidden in edit mode) -->
      <div class="form-row-seibies hidden-section">
        <div class="form-group-seibies project-select-container-seibies">
          <label>Project</label>
          <select id="projectSelectSeibies" required></select>
        </div>
      </div>

      <!-- MAIN WRAPPER - 2 Column Layout -->
      <div class="project-main-wrapper-seibies">

        <!-- LEFT SIDE - Project Details (6 Rows) -->
        <div class="project-left-side-seibies" id="projectDetailsSeibies">
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Project:</span>
            <span class="detail-value-seibies" id="projectNameDisplaySeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Project Ref:</span>
            <span class="detail-value-seibies" id="projectRefSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Profit Center:</span>
            <span class="detail-value-seibies" id="profitCenterSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Project Manager:</span>
            <span class="detail-value-seibies" id="projectManagerSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Group Head:</span>
            <span class="detail-value-seibies" id="groupHeadSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Customer:</span>
            <span class="detail-value-seibies" id="customerSeibies">-</span>
          </div>
        </div>

        <!-- RIGHT SIDE - OBL Section -->
        <div class="project-right-side-seibies">
          <div class="obl-section-container-seibies hidden-section" id="oblContainerSeibies">
            <div class="obl-content-seibies">
              <div class="obl-left-seibies">
                <div class="obl-label-group-seibies">
                  <label>SEI BIES </label>
                </div>
                <div class="obl-value-display-seibies">
                  <input type="text" id="oblLastMonthSeibies" class="obl-input-seibies number-input-seibies" placeholder="Enter SEIBIES" readonly />
                </div>
              </div>
              <div class="obl-right-seibies">
                <div class="validation-summary-seibies">
                  <div class="validation-item-seibies">
                    <span class="validation-label-seibies">Total Current Value:</span>
                    <span class="validation-value-seibies" id="totalCurrentValueSeibies">0</span>
                  </div>
                  <div class="validation-item-seibies">
                    <span class="validation-label-seibies">Remaining Balance:</span>
                    <span class="validation-value-seibies" id="remainingValueSeibies">0</span>
                  </div>
                  <div class="validation-line-seibies">
                    <div class="line-fill-seibies" id="oblLineFillSeibies"></div>
                  </div>
                </div>
              </div>
            </div>
            <div class="validation-message-seibies" id="oblValidationMessageSeibies">
              ✅ All values are within SEI-BIES limit
            </div>
            <div class="table-toolbar__note" style="margin-top: 15px;">
              Note: "Negative=BIES, Positive=SEI"
            </div>
          </div>
        </div>

      </div>

      <!-- Forecast Sections -->
      <div class="forecast-section-seibies hidden-section" id="historicalForecastSectionSeibies">
        <div class="historical-toggle-header-seibies">
          <h3>📊 Historical Data (Last 3 Months - Readonly)</h3>
          <button type="button" id="toggleHistoricalBtnSeibies" class="toggle-arrow-btn-seibies" aria-expanded="false" title="Show historical data">
            <span class="toggle-arrow-seibies">▼</span>
            <span class="toggle-label-seibies">Show</span>
          </button>
        </div>
        <div id="historicalContainerSeibies" class="historical-container-seibies" style="display:none;"></div>
      </div>
      
      <div class="forecast-section-seibies hidden-section" id="detailedForecastSectionSeibies">
        <h3>Detailed Forecast (Current + Next 2 Months)</h3>
        <div id="firstRowContainerSeibies" class="first-row-container-seibies"></div>
      </div>
      
      <div class="forecast-section-seibies hidden-section" id="longtermForecastSectionSeibies">
        <h3>Long-term Forecast (Next 15 Months)</h3>
        <div id="remainingContainerSeibies" class="remaining-container-seibies"></div>
      </div>

      <div class="forecast-section-seibies hidden-section" id="yearlyForecastSectionSeibies">
        <h3>5-Year Forecast</h3>
        <div id="yearlyContainerSeibies" class="yearly-container-seibies"></div>
      </div>
      
      <div class="form-actions-seibies">
        <button type="submit" class="btn-primary-seibies">💾 Save</button>
        <button type="button" class="btn-secondary-seibies back-to-table-seibies">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,
    script: () => {
        setTimeout(async () => {
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/Kriya";
            const listName = "SEIBIES";
            const projectListName = "Project List";
            // DOM Elements
            const tableBody = document.querySelector(".seibies-table-body");
            const monthColumnsHeader = document.querySelector("#monthColumnsHeaderSeibies");
            const totalsHeader = document.querySelector("#totalsHeaderSeibies");
            const formSection = document.querySelector(".form-section-seibies");
            const tableSection = document.querySelector(".table-section-seibies");
            const showFormBtn = document.querySelector(".show-form-btn-seibies");
            const projectSelect = document.querySelector("#projectSelectSeibies");
            const historicalContainer = document.querySelector("#historicalContainerSeibies");
            const firstRowContainer = document.querySelector("#firstRowContainerSeibies");
            const remainingContainer = document.querySelector("#remainingContainerSeibies");
            const yearlyContainer = document.querySelector("#yearlyContainerSeibies");
            const backBtn = document.querySelector(".back-to-table-seibies");
            const form = document.querySelector("#seibiesForm");
            const exportBtn = document.querySelector(".export-btn-seibies");
            const oblLastMonthInput = document.querySelector("#oblLastMonthSeibies");
            const totalCurrentValueSpan = document.querySelector("#totalCurrentValueSeibies");
            const remainingValueSpan = document.querySelector("#remainingValueSeibies");
            const oblLineFill = document.querySelector("#oblLineFillSeibies");
            const oblValidationMessage = document.querySelector("#oblValidationMessageSeibies");
            const formTitle = document.querySelector("#formTitleSeibies");
            // Filter Elements
            const filterYear = document.querySelector("#filterYearSeibies");
            const filterMonth = document.querySelector("#filterMonthSeibies");
            const filterProject = document.querySelector("#filterProjectSeibies");
            const filterCustomer = document.querySelector("#filterCustomerSeibies");
            const filterProjectManager = document.querySelector("#filterProjectManagerSeibies");
            const filterGroupHead = document.querySelector("#filterGroupHeadSeibies");
            const filterProjectReference = document.querySelector("#filterProjectReferenceSeibies");
            // The Reset/clear action now lives INSIDE the Filters popup (Clear button),
            // so the old toolbar button - and its reference - are gone.
            // Pagination Elements
            const firstPageBtn = document.querySelector("#firstPageSeibies");
            const prevPageBtn = document.querySelector("#prevPageSeibies");
            const nextPageBtn = document.querySelector("#nextPageSeibies");
            const lastPageBtn = document.querySelector("#lastPageSeibies");
            // kept for backwards compatibility - the "Page X of Y" pill was replaced by
            // the "+ page tiles of N" layout, so this node may not exist any more
            const currentPageSpan = document.querySelector("#currentPageSeibies");
            const totalPagesSpan = document.querySelector("#totalPagesSeibies");
            const pageNumbersDiv = document.querySelector("#pageNumbersSeibies");
            const startItemSpan = document.querySelector("#startItemSeibies");
            const endItemSpan = document.querySelector("#endItemSeibies");
            const totalItemsSpan = document.querySelector("#totalItemsSeibies");
            const pageSizeSelect = document.querySelector("#pageSizeSelectSeibies");
            // ============================================================
            // ✅ SHARED HELPERS for the Save / Back confirmation dialogs
            //    (same recipe as procurement.ts / dci.ts / revenue.ts / project.ts / inventory.ts / invoicing.ts)
            // ============================================================
            const escapeHtml = (unsafe) => {
                if (unsafe === null || unsafe === undefined)
                    return "";
                return String(unsafe)
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;")
                    .replace(/"/g, "&quot;")
                    .replace(/'/g, "&#039;");
            };
            // Guards so neither flow can be re-entered by double-clicking
            let isSavingSeibies = false;
            let isLeavingSeibiesForm = false;
            // Dialogs must render ABOVE every overlay of this module
            const bumpSwalZIndex = () => {
                const c = document.querySelector(".swal2-container");
                if (c)
                    c.style.zIndex = "2147483647";
            };
            let projects = [];
            let forecasts = [];
            let editingId = null;
            let allMonths = [];
            // Pagination variables
            let pageSize = 25;
            let currentPage = 1;
            let totalPages = 1;
            let filteredData = [];
            // ========== UTILITY FUNCTIONS ==========
            const toast = (icon, title) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
            /**
             * ✅ SINGLE SOURCE OF TRUTH for every save button in the SEIBIES form.
             *    Both the top button (in the sticky project details panel) and the
             *    bottom button (in .form-actions-seibies) carry the .btn-primary-seibies
             *    class + type="submit", so this helper always returns BOTH.
             *    Any code that shows / hides / disables / re-labels a save button MUST
             *    use this helper so the two buttons never drift out of sync.
             */
            const getSaveButtons = () => Array.from(document.querySelectorAll('button.btn-primary-seibies[type="submit"]'));
            // ============================================
            // SECTION LOADER
            // ============================================
            // ✅ PAGE LOADER instead of the old spinner + text - bar only
            const getRevenueLoaderMarkup = () => `
  <div id="revenuePageLoader" class="revenue-loader" style="position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;background:rgba(248,249,250,0.92);z-index:999;">
    <div class="seibies-loader-panel">
      <div class="seibies-page-loader">
        <div class="seibies-page-loader__track">
          <div class="seibies-page-loader__bar"></div>
        </div>
      </div>
    </div>
  </div>`;
            const showRevenueLoader = (text, subtext) => {
                var _a;
                let loader = document.querySelector('#revenuePageLoader');
                if (!loader) {
                    const host = document.querySelector('.seibies-page') || document.body;
                    host.insertAdjacentHTML('afterbegin', getRevenueLoaderMarkup());
                    loader = document.querySelector('#revenuePageLoader');
                }
                if (!loader)
                    return;
                loader.classList.remove('revenue-loader-hidden');
                loader.style.display = 'flex';
                loader.style.opacity = '1';
                loader.style.visibility = 'visible';
                (_a = loader.parentElement) === null || _a === void 0 ? void 0 : _a.classList.add('revenue-loading');
                positionLoaderOverlay();
                window.setTimeout(positionLoaderOverlay, 0);
                // ✅ The loader is bar-only now (nothing is shown on screen), so the step
                //    messages are kept in the console to stay traceable. Every existing
                //    showRevenueLoader('Saving', '…') call site keeps working unchanged.
                if (text)
                    console.log(`⏳ ${text}${subtext ? ` - ${subtext}` : ''}`);
            };
            const positionLoaderOverlay = () => {
                const loader = document.querySelector('#revenuePageLoader');
                if (!loader || !loader.parentElement)
                    return;
                const rect = loader.parentElement.getBoundingClientRect();
                const top = Math.max(0, rect.top);
                const height = Math.max(window.innerHeight - top, rect.height - top, 320);
                loader.style.setProperty('--loader-top', `${top}px`);
                loader.style.height = `${height}px`;
            };
            window.addEventListener('resize', positionLoaderOverlay);
            window.addEventListener('scroll', positionLoaderOverlay, { passive: true });
            const hideRevenueLoader = () => {
                var _a;
                const loader = document.querySelector('#revenuePageLoader');
                (_a = document.querySelector('.seibies-page')) === null || _a === void 0 ? void 0 : _a.classList.remove('revenue-loading');
                if (!loader)
                    return;
                loader.classList.add('revenue-loader-hidden');
                window.setTimeout(() => {
                    if (loader.parentNode)
                        loader.parentNode.removeChild(loader);
                }, 350);
            };
            const reloadSeibiesData = async () => {
                showRevenueLoader('Reloading Data', 'Refreshing projects and forecasts...');
                try {
                    await loadProjects();
                    await loadForecasts();
                    populateFilterOptions();
                    syncFilterUIFromSelects();
                    initializePagination(forecasts);
                }
                catch (error) {
                    console.error('❌ Reload failed:', error);
                    toast("error", "Reload failed - showing current data");
                }
                finally {
                    hideRevenueLoader();
                }
            };
            window.reloadSeibiesData = reloadSeibiesData;
            window.hideRevenueLoader = hideRevenueLoader;
            const getDigest = async () => {
                const res = await fetch(`${siteUrl}/_api/contextinfo`, {
                    method: "POST",
                    headers: { Accept: "application/json;odata=verbose" }
                });
                const data = await res.json();
                return data.d.GetContextWebInformation.FormDigestValue;
            };
            const safeJsonParse = (jsonString, fallback = {}) => {
                try {
                    if (!jsonString || jsonString.trim() === '') {
                        return fallback;
                    }
                    const cleanedString = jsonString.trim();
                    if (cleanedString === '""' || cleanedString === "null") {
                        return fallback;
                    }
                    return JSON.parse(cleanedString);
                }
                catch (error) {
                    console.error("JSON parse error:", error, "Input:", jsonString);
                    return fallback;
                }
            };
            const formatMonthKey = (date) => {
                return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
            };
            const formatMonthLabel = (date) => {
                return date.toLocaleString("default", { month: "short", year: "2-digit" });
            };
            const generateForecastTitle = (project, forecastData) => {
                const currentYear = new Date().getFullYear();
                let totalValue = 0;
                Object.keys(forecastData).forEach(key => {
                    if (key.match(/^\d{4}-\d{2}$/)) {
                        const monthData = forecastData[key];
                        if (monthData) {
                            totalValue += monthData.actual;
                        }
                    }
                });
                return `${project.Title} - ${currentYear} SEIBIES Forecast - ${totalValue.toLocaleString('en-US')}`;
            };
            const getDisplayMonths = () => {
                const now = new Date();
                const months = [];
                for (let i = 0; i < 3; i++) {
                    const month = new Date(now.getFullYear(), now.getMonth() + i, 1);
                    months.push(formatMonthKey(month));
                }
                return months;
            };
            const getAllMonthsForStorage = () => {
                return generateIdealMonthStructure();
            };
            const allowOnlyNumbers = (input) => {
                input.addEventListener('input', function () {
                    let value = this.value.replace(/[^\d]/g, '');
                    if (value.length > 1 && value.startsWith('0')) {
                        value = value.replace(/^0+/, '');
                        if (value === '')
                            value = '0';
                    }
                    this.value = value;
                    if (value === '' || value === '0') {
                        this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                    }
                    else {
                        this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                    }
                });
                input.addEventListener('keydown', function (e) {
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
                input.addEventListener('paste', function (e) {
                    const clipboardData = e.clipboardData || window.clipboardData;
                    const pastedData = (clipboardData === null || clipboardData === void 0 ? void 0 : clipboardData.getData('Text')) || '';
                    if (/[^\d]/.test(pastedData)) {
                        e.preventDefault();
                    }
                });
            };
            const getNumericValue = (input) => {
                const value = input.value.trim();
                if (value === '' || value === '0')
                    return 0;
                return parseInt(value) || 0;
            };
            const setNumericValue = (input, value) => {
                if (value === 0) {
                    input.value = '';
                    input.className = input.className.replace('input-with-value-seibies', 'input-empty-seibies');
                }
                else {
                    input.value = value.toString();
                    input.className = input.className.replace('input-empty-seibies', 'input-with-value-seibies');
                }
            };
            const showOBLErrorToast = (message, isWarning = false) => {
                var _a;
                const existingToast = document.querySelector('.obl-error-toast-seibies');
                if (existingToast)
                    existingToast.remove();
                const toast = document.createElement('div');
                toast.className = `obl-error-toast-seibies ${isWarning ? 'warning-seibies' : ''}`;
                toast.innerHTML = `
          ${message}
          <button class="toast-close-seibies">&times;</button>
        `;
                document.body.appendChild(toast);
                setTimeout(() => {
                    if (toast.parentNode) {
                        toast.remove();
                    }
                }, 5000);
                (_a = toast.querySelector('.toast-close-seibies')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
                    toast.remove();
                });
            };
            // ========== MONTH STRUCTURE FUNCTIONS ==========
            const generateIdealMonthStructure = () => {
                const months = [];
                const currentDate = new Date();
                for (let i = -3; i < 0; i++) {
                    const historicalDate = new Date(currentDate);
                    historicalDate.setMonth(historicalDate.getMonth() + i);
                    months.push(formatMonthKey(historicalDate));
                }
                for (let i = 0; i < 3; i++) {
                    const detailedDate = new Date(currentDate);
                    detailedDate.setMonth(detailedDate.getMonth() + i);
                    months.push(formatMonthKey(detailedDate));
                }
                for (let i = 3; i < 18; i++) {
                    const futureDate = new Date(currentDate);
                    futureDate.setMonth(futureDate.getMonth() + i);
                    months.push(formatMonthKey(futureDate));
                }
                return months;
            };
            // ========== YEARLY FORECAST FUNCTIONS ==========
            const getLastMonthKey = (existingData) => {
                const monthKeys = Object.keys(existingData).filter(key => key.match(/^\d{4}-\d{2}$/));
                if (monthKeys.length === 0) {
                    return undefined;
                }
                const sortedMonths = monthKeys.sort();
                const lastMonth = sortedMonths[sortedMonths.length - 1];
                return lastMonth;
            };
            const createFreshYearlyStructure = (lastMonthKey) => {
                const yearlyData = {};
                let startYear;
                if (lastMonthKey) {
                    startYear = parseInt(lastMonthKey.split('-')[0]);
                }
                else {
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
            const createFreshForecastStructure = () => {
                const idealMonths = generateIdealMonthStructure();
                const freshData = {};
                const currentMonthKey = formatMonthKey(new Date());
                idealMonths.forEach((monthKey) => {
                    const isCurrentMonth = monthKey === currentMonthKey;
                    freshData[monthKey] = {
                        plan: 0,
                        actual: 0,
                        forecast: 0,
                        date: isCurrentMonth ? '' : undefined
                    };
                });
                const lastMonthKey = idealMonths[idealMonths.length - 1];
                freshData.yearly = createFreshYearlyStructure(lastMonthKey);
                return freshData;
            };
            const isForecastMonthData = (data) => {
                return data &&
                    typeof data.plan === 'number' &&
                    typeof data.actual === 'number';
            };
            const handleMonthProgression = (forecastData) => {
                const now = new Date();
                const currentMonthKey = formatMonthKey(now);
                const originalMonths = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
                const idealMonths = generateIdealMonthStructure();
                if (originalMonths[0] === idealMonths[0]) {
                    return forecastData;
                }
                const updatedData = {};
                idealMonths.forEach((monthKey) => {
                    const existingData = forecastData[monthKey];
                    if (existingData && isForecastMonthData(existingData)) {
                        updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
                    }
                    else {
                        updatedData[monthKey] = {
                            plan: 0,
                            actual: 0
                        };
                    }
                });
                const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
                idealMonths.forEach((monthKey, index) => {
                    const monthData = updatedData[monthKey];
                    if (monthData && isForecastMonthData(monthData)) {
                        const isCurrentOrFuture = index >= currentMonthIndex;
                        if (isCurrentOrFuture) {
                            const oldActual = monthData.actual;
                            monthData.plan = oldActual;
                            monthData.actual = 0;
                        }
                    }
                });
                if (forecastData.yearly) {
                    updatedData.yearly = JSON.parse(JSON.stringify(forecastData.yearly));
                }
                else {
                    const lastMonthKey = idealMonths[idealMonths.length - 1];
                    updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
                }
                return updatedData;
            };
            const propagateMonthValues = (forecastData) => {
                if (Object.keys(forecastData).length === 0) {
                    return forecastData;
                }
                return handleMonthProgression(forecastData);
            };
            // ========== DATA LOADING FUNCTIONS ==========
            const loadProjects = async () => {
                var _a;
                try {
                    const listPath = `${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items`;
                    const headers = { Accept: "application/json;odata=verbose" };
                    // ✅ $select + $expand must include every field ROLE_OWN_FIELDS in auth.ts
                    // matches against, otherwise getAssignedProjects() returns [] for the
                    // KRIYA_Engg_* roles and every SEIBIES forecast gets filtered out.
                    const fullQuery = `${listPath}?$select=Id,Title,CustomerName,PC,ProjectReference,` +
                        `GroupHead1/Title,GroupHead1Id,` +
                        `ProjectManager/Title,ProjectManagerId,` +
                        `EnggLead/Title,EnggLeadId,` +
                        `EnggSubLead1/Title,EnggSubLead1Id,` +
                        `EnggSubLead2/Title,EnggSubLead2Id,` +
                        `EngineeringManagers/Title,EngineeringManagersId` +
                        `&$expand=GroupHead1,ProjectManager,EnggLead,EnggSubLead1,EnggSubLead2,EngineeringManagers` +
                        `&$top=5000`;
                    let res = await fetch(fullQuery, { headers });
                    // ✅ No "| null" — start as [] so .map() below type-checks cleanly
                    let rawProjects = [];
                    // Fallback 1 — engineering fields as IDs only (still carries Engg*Id
                    // columns so names can be resolved via getuserbyid)
                    if (!res.ok) {
                        console.warn(`Project List expanded query failed (${res.status}); retrying with ID-only query`);
                        const idQuery = `${listPath}?$select=Id,Title,CustomerName,PC,ProjectReference,` +
                            `GroupHead1Id,ProjectManagerId,` +
                            `EnggLeadId,EnggSubLead1Id,EnggSubLead2Id,EngineeringManagersId` +
                            `&$top=5000`;
                        res = await fetch(idQuery, { headers });
                    }
                    // Fallback 2 — bare fetch, prune to what we need
                    if (!res.ok) {
                        console.warn(`Project List ID query failed (${res.status}); retrying with full fetch`);
                        res = await fetch(`${listPath}?$top=5000`, { headers });
                    }
                    if (!res.ok)
                        throw new Error(`HTTP error! status: ${res.status}`);
                    const data = await res.json();
                    rawProjects = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    // Memoised user-id → name resolver (used by the fallback paths)
                    const userCache = new Map();
                    const getUserName = async (id) => {
                        var _a;
                        if (!id)
                            return "";
                        if (userCache.has(id))
                            return userCache.get(id);
                        try {
                            const uRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${id})?$select=Id,Title,Email`, { headers });
                            if (uRes.ok) {
                                const uData = await uRes.json();
                                const name = ((_a = uData.d) === null || _a === void 0 ? void 0 : _a.Title) || "";
                                userCache.set(id, name);
                                return name;
                            }
                        }
                        catch (e) {
                            console.warn(`Could not fetch user ${id}:`, e);
                        }
                        userCache.set(id, "");
                        return "";
                    };
                    const nameOf = (val) => {
                        if (!val)
                            return "";
                        if (typeof val === "object")
                            return val.Title || "";
                        return String(val);
                    };
                    const cleanedProjects = await Promise.all(rawProjects.map(async (p) => {
                        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m;
                        return ({
                            Id: p.Id,
                            Title: p.Title || "",
                            CustomerName: p.CustomerName || "",
                            PC: p.PC || "",
                            ProjectReference: p.ProjectReference || "",
                            GroupHead1: p.GroupHead1 ? nameOf(p.GroupHead1) : await getUserName(p.GroupHead1Id),
                            GroupHead1Id: (_b = (_a = p.GroupHead1) === null || _a === void 0 ? void 0 : _a.Id) !== null && _b !== void 0 ? _b : p.GroupHead1Id,
                            ProjectManager: p.ProjectManager ? nameOf(p.ProjectManager) : await getUserName(p.ProjectManagerId),
                            ProjectManagerId: (_d = (_c = p.ProjectManager) === null || _c === void 0 ? void 0 : _c.Id) !== null && _d !== void 0 ? _d : p.ProjectManagerId,
                            // ✅ Engineering fields
                            EnggLead: p.EnggLead ? nameOf(p.EnggLead) : await getUserName(p.EnggLeadId),
                            EnggLeadId: (_f = (_e = p.EnggLead) === null || _e === void 0 ? void 0 : _e.Id) !== null && _f !== void 0 ? _f : p.EnggLeadId,
                            EnggSubLead1: p.EnggSubLead1 ? nameOf(p.EnggSubLead1) : await getUserName(p.EnggSubLead1Id),
                            EnggSubLead1Id: (_h = (_g = p.EnggSubLead1) === null || _g === void 0 ? void 0 : _g.Id) !== null && _h !== void 0 ? _h : p.EnggSubLead1Id,
                            EnggSubLead2: p.EnggSubLead2 ? nameOf(p.EnggSubLead2) : await getUserName(p.EnggSubLead2Id),
                            EnggSubLead2Id: (_k = (_j = p.EnggSubLead2) === null || _j === void 0 ? void 0 : _j.Id) !== null && _k !== void 0 ? _k : p.EnggSubLead2Id,
                            EngineeringManagers: p.EngineeringManagers ? nameOf(p.EngineeringManagers) : await getUserName(p.EngineeringManagersId),
                            EngineeringManagersId: (_m = (_l = p.EngineeringManagers) === null || _l === void 0 ? void 0 : _l.Id) !== null && _m !== void 0 ? _m : p.EngineeringManagersId
                        });
                    }));
                    const projectMap = new Map();
                    cleanedProjects.forEach((p) => projectMap.set(p.Title, p));
                    projects = Array.from(projectMap.values());
                    console.log(`Loaded ${projects.length} projects from Project List`);
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
                    console.log('📋 Assigned project titles:', assignedProjects.map((p) => p.Title));
                    projectSelect.innerHTML =
                        `<option value="">Select Project</option>` +
                            assignedProjects.map((p) => `<option value="${p.Id}">${p.Title}</option>`).join("");
                    filterProject.innerHTML =
                        '<option value="">All Projects</option>' +
                            [...new Set(assignedProjects.map((p) => p.Title))]
                                .sort()
                                .map(project => `<option value="${project}">${project}</option>`)
                                .join("");
                }
                catch (error) {
                    console.error("Error loading projects:", error);
                    toast("error", "Failed to load projects");
                }
            };
            const loadForecasts = async () => {
                var _a;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,SeibiesActual,CreatedDate,UpdatedDate&$expand=Project&$top=5000`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const data = await res.json();
                    const assignedProjects = getAssignedProjects(projects);
                    const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                    const allForecasts = (((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || []).map(forecast => (Object.assign(Object.assign({}, forecast), { ForecastData: typeof forecast.ForecastData === 'string' ?
                            forecast.ForecastData :
                            JSON.stringify(forecast.ForecastData || {}), SeibiesActual: typeof forecast.SeibiesActual === 'string' ?
                            forecast.SeibiesActual :
                            JSON.stringify(forecast.SeibiesActual || {}) })));
                    forecasts = allForecasts.filter(forecast => {
                        var _a;
                        const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        return projectTitle && assignedProjectTitles.has(projectTitle);
                    });
                    console.log(`✅ Loaded ${forecasts.length} SEIBIES forecasts`);
                    console.log(`👑 Is GH Head: ${isGHHead()}`);
                }
                catch (error) {
                    console.error("Error loading SEIBIES forecasts:", error);
                    toast("error", "Failed to load SEIBIES forecasts");
                }
            };
            const getGroupHeadDisplayName = (groupHead) => {
                if (!groupHead)
                    return "-";
                if (typeof groupHead === "object" && groupHead !== null) {
                    return groupHead.Title || "-";
                }
                // ✅ cast: the object branch was handled above, so this is the plain string
                return groupHead;
            };
            // ========== EXPORT FUNCTIONS ==========
            const getMonthName = (monthNumber) => {
                const months = [
                    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
                ];
                return months[parseInt(monthNumber) - 1] || "Month";
            };
            const generateHeadersStructure = (selectedYear, selectedMonth, exportData) => {
                const baseHeaders = [
                    "Profit Centre",
                    "SO/Project Ref",
                    "Customer Name",
                    "Project Name",
                    "Group Head",
                    "Project Manager",
                    "SEI-BIES",
                ];
                const selectedMonthName = getMonthName(selectedMonth);
                const selectedMonthHeaders = [
                    `${selectedMonthName} ${selectedYear} (P)`,
                    `${selectedMonthName} ${selectedYear} (C)`,
                ];
                const allMonthKeys = new Set();
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
                const monthlyHeaders = [];
                next17Months.forEach(monthKey => {
                    const [year, month] = monthKey.split('-');
                    const monthLabel = getMonthName(month);
                    const displayLabel = `${monthLabel} ${year}`;
                    monthlyHeaders.push(`${displayLabel} (P)`, `${displayLabel} (C)`);
                });
                const lastFiveYears = getLastFiveYears(exportData);
                const yearlyHeaders = [];
                lastFiveYears.forEach(year => {
                    yearlyHeaders.push(`${year} (P)`, `${year} (C)`);
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
            const getLastFiveYears = (data) => {
                const years = new Set();
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
            const exportToExcelWithFullStyle = () => {
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
                    toast("error", "No SEIBIES data available to export");
                    return;
                }
                showRevenueLoader('Exporting', 'Preparing the Excel workbook...');
                try {
                    const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);
                    const excelData = exportData.map(forecast => {
                        var _a;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        const oblValue = forecast.OBL || 0;
                        const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
                        const monthName = getMonthName(selectedMonth);
                        const selectedMonthForecastData = forecastData[selectedMonthKey] || { plan: 0, actual: 0 };
                        const selectedMonthPlan = selectedMonthForecastData.plan || 0;
                        const selectedMonthCurrent = selectedMonthForecastData.actual || 0;
                        let groupHeadDisplay = "-";
                        if (proj === null || proj === void 0 ? void 0 : proj.GroupHead1) {
                            groupHeadDisplay = typeof proj.GroupHead1 === "object" && proj.GroupHead1 !== null
                                ? proj.GroupHead1.Title || ""
                                : proj.GroupHead1 || "";
                        }
                        if (!groupHeadDisplay)
                            groupHeadDisplay = "-";
                        let projectManagerDisplay = "-";
                        if (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) {
                            projectManagerDisplay = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
                                ? proj.ProjectManager.Title || ""
                                : proj.ProjectManager || "";
                        }
                        if (!projectManagerDisplay)
                            projectManagerDisplay = "-";
                        const rowData = {
                            "Profit Centre": forecast.ProfitCentre || "-",
                            "SO/Project Ref": (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                            "Customer Name": (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                            "Project Name": projectTitle,
                            "Group Head": groupHeadDisplay,
                            "Project Manager": projectManagerDisplay,
                            "SEI-BIES": oblValue,
                            [`${monthName} ${selectedYear} (P)`]: selectedMonthPlan,
                            [`${monthName} ${selectedYear} (C)`]: selectedMonthCurrent,
                        };
                        allMonths.forEach(monthKey => {
                            const [year, month] = monthKey.split('-');
                            const monthLabel = getMonthName(month);
                            const displayLabel = `${monthLabel} ${year}`;
                            const forecastMonthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                            const planValue = forecastMonthData.plan || 0;
                            const currentValue = forecastMonthData.actual || 0;
                            rowData[`${displayLabel} (P)`] = planValue;
                            rowData[`${displayLabel} (C)`] = currentValue;
                        });
                        lastFiveYears.forEach(year => {
                            var _a;
                            const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0 };
                            rowData[`${year} (P)`] = yearData.plan;
                            rowData[`${year} (C)`] = yearData.actual;
                        });
                        const totalMonthlyCurrent = Object.keys(forecastData)
                            .filter(key => key.match(/^\d{4}-\d{2}$/))
                            .reduce((sum, monthKey) => {
                            const monthData = forecastData[monthKey];
                            return sum + ((monthData === null || monthData === void 0 ? void 0 : monthData.actual) || 0);
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
                    const numberStyle = Object.assign(Object.assign({}, dataStyle), { alignment: { horizontal: "right", vertical: "center" } });
                    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s) {
                            ws[cellAddress].s = {};
                        }
                        ws[cellAddress].s = headerStyle;
                    }
                    for (let row = range.s.r + 1; row <= range.e.r; row++) {
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                            if (!ws[cellAddress])
                                continue;
                            const isNumberColumn = col >= 6;
                            const styleToApply = isNumberColumn ? numberStyle : dataStyle;
                            if (!ws[cellAddress].s) {
                                ws[cellAddress].s = {};
                            }
                            ws[cellAddress].s = styleToApply;
                        }
                    }
                    const colWidths = headers.map((header) => {
                        if (header.includes('Project Name'))
                            return { wch: 25 };
                        if (header.includes('Customer Name'))
                            return { wch: 20 };
                        if (header.includes('Profit Centre') || header.includes('SO/Project Ref'))
                            return { wch: 12 };
                        if (header.includes('Group Head') || header.includes('Project Manager'))
                            return { wch: 15 };
                        if (header.includes('SEI-BIES') || header.includes('Balance'))
                            return { wch: 12 };
                        if (header.includes('(P)') || header.includes('(C)'))
                            return { wch: 10 };
                        return { wch: 12 };
                    });
                    ws['!cols'] = colWidths;
                    ws['!freeze'] = { x: 0, y: 1 };
                    XLSX.utils.book_append_sheet(wb, ws, "SEIBIES Forecast");
                    const monthName = getMonthName(selectedMonth);
                    const fileName = `SEIBIES_Forecast_${monthName}_${selectedYear}_Filtered.xlsx`;
                    XLSX.writeFile(wb, fileName);
                    hideRevenueLoader();
                    toast("success", `Exported ${exportData.length} filtered SEIBIES records for ${monthName} ${selectedYear}`);
                }
                catch (error) {
                    hideRevenueLoader();
                    console.error("Export error:", error);
                    toast("error", "Failed to export SEIBIES data to Excel");
                }
            };
            // ========== PAGINATION FUNCTIONS ==========
            const initializePagination = (data) => {
                filteredData = [...data];
                calculateTotalPages();
                currentPage = 1;
                updatePaginationControls();
                renderCurrentPage();
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                setTimeout(() => {
                    renderTotalsHeader(displayMonths, currentMonthKey);
                }, 100);
            };
            const calculateTotalPages = () => {
                totalPages = Math.ceil(filteredData.length / pageSize);
                if (totalPages === 0)
                    totalPages = 1;
            };
            const changePageSize = (newSize) => {
                if (newSize === pageSize)
                    return;
                pageSize = newSize;
                const oldStartItem = (currentPage - 1) * pageSize;
                currentPage = Math.floor(oldStartItem / pageSize) + 1;
                calculateTotalPages();
                if (currentPage > totalPages) {
                    currentPage = totalPages || 1;
                }
                updatePaginationControls();
                renderCurrentPage();
                toast("info", `Showing ${pageSize} SEIBIES records per page`);
            };
            const updatePaginationControls = () => {
                // ✅ Every element is OPTIONAL: a missing node must never abort the whole
                //    update (that used to freeze every pager button).
                const setText = (el, value) => {
                    if (el)
                        el.textContent = value;
                };
                const totalItems = filteredData.length;
                const startItem = totalItems === 0 ? 0 : (currentPage - 1) * pageSize + 1;
                const endItem = Math.min(currentPage * pageSize, totalItems);
                if (currentPageSpan)
                    currentPageSpan.textContent = currentPage.toString();
                setText(totalPagesSpan, totalPages.toString());
                setText(startItemSpan, startItem.toString());
                setText(endItemSpan, endItem.toString());
                setText(totalItemsSpan, totalItems.toString());
                if (pageSizeSelect) {
                    pageSizeSelect.value = pageSize.toString();
                }
                if (firstPageBtn)
                    firstPageBtn.disabled = currentPage === 1 || totalItems === 0;
                if (prevPageBtn)
                    prevPageBtn.disabled = currentPage === 1 || totalItems === 0;
                if (nextPageBtn)
                    nextPageBtn.disabled = currentPage === totalPages || totalItems === 0;
                if (lastPageBtn)
                    lastPageBtn.disabled = currentPage === totalPages || totalItems === 0;
                updatePageNumbers();
            };
            const updatePageNumbers = () => {
                pageNumbersDiv.innerHTML = '';
                if (filteredData.length === 0)
                    return;
                const maxVisiblePages = 5;
                let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
                let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
                if (endPage - startPage + 1 < maxVisiblePages) {
                    startPage = Math.max(1, endPage - maxVisiblePages + 1);
                }
                if (startPage > 1) {
                    const dotsBtn = document.createElement('button');
                    dotsBtn.className = 'page-number-seibies dots-seibies';
                    dotsBtn.textContent = '...';
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
                for (let i = startPage; i <= endPage; i++) {
                    const pageBtn = document.createElement('button');
                    pageBtn.className = `page-number-seibies ${i === currentPage ? 'active-seibies' : ''}`;
                    pageBtn.textContent = i.toString();
                    pageBtn.addEventListener('click', () => goToPage(i));
                    pageNumbersDiv.appendChild(pageBtn);
                }
                if (endPage < totalPages) {
                    const dotsBtn = document.createElement('button');
                    dotsBtn.className = 'page-number-seibies dots-seibies';
                    dotsBtn.textContent = '...';
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
            };
            const goToPage = (page) => {
                if (page < 1 || page > totalPages || page === currentPage)
                    return;
                currentPage = page;
                updatePaginationControls();
                renderCurrentPage();
            };
            const setupPaginationEventListeners = () => {
                if (firstPageBtn)
                    firstPageBtn.addEventListener("click", () => goToPage(1));
                if (prevPageBtn)
                    prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
                if (nextPageBtn)
                    nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
                if (lastPageBtn)
                    lastPageBtn.addEventListener("click", () => goToPage(totalPages));
                if (pageSizeSelect) {
                    pageSizeSelect.addEventListener("change", (e) => {
                        const newSize = parseInt(e.target.value);
                        changePageSize(newSize);
                    });
                }
            };
            // ========== TABLE RENDERING FUNCTIONS ==========
            const calculateTotals = (data, displayMonths, currentMonthKey) => {
                const totals = {
                    obl: 0,
                    monthly: {}
                };
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
            const applyStickyHeaderOffsets = () => {
                const table = document.querySelector('.month-columns-table-seibies');
                if (!table)
                    return;
                const headerRows = table.querySelectorAll('thead tr');
                let offset = 0;
                headerRows.forEach(row => {
                    row.querySelectorAll('th, td').forEach(cell => {
                        cell.style.position = 'sticky';
                        cell.style.top = `${offset}px`;
                    });
                    offset += row.getBoundingClientRect().height;
                });
            };
            const renderTotalsHeader = (displayMonths, currentMonthKey) => {
                if (!totalsHeader)
                    return;
                totalsHeader.innerHTML = '';
                const allData = forecasts;
                const filtered = filteredData.length ? filteredData : allData;
                const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
                const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);
                // const baseColumns = 6;
                // const monthColumns = displayMonths.length * 2;
                // const totalColspan = baseColumns + monthColumns;
                const buildTotalsRow = (label, totals, css) => {
                    const monthCells = displayMonths.map(m => {
                        const t = totals.monthly[m] || { plan: 0, current: 0 };
                        return `
              <td class="month-scroll-col-seibies" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-US')}</td>
              <td class="month-scroll-col-seibies" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-US')}</td>
            `;
                    }).join('');
                    return `
            <tr class="${css}">
              <td class="project-ref-col-seibies" style="text-align:center; vertical-align:middle; min-width:120px; max-width:120px; width:120px;">${css === 'all-projects-total-seibies' ? '★' : '●'}</td>
              <td class="customer-col-seibies" style="text-align:center; vertical-align:middle; min-width:130px; max-width:130px; width:130px;"></td>
              <td class="project-col-seibies" style="text-align:left; vertical-align:middle; min-width:180px; max-width:180px; width:180px; font-weight:bold; background-color:${css === 'all-projects-total-seibies' ? '#d1ecf1' : '#d4edda'};">
                ${label}
              </td>
              <td class="gh-col-seibies" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
              <td class="pm-col-seibies" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
              <td class="obl-col-seibies" style="text-align:right; vertical-align:middle; min-width:90px; max-width:90px; width:90px; font-weight:bold;">${totals.obl.toLocaleString('en-US')}</td>
              ${monthCells}
            </tr>
          `;
                };
                let html = `
          <tr class="totals-title-row-seibies">
             
          </tr>
        `;
                if (allData.length > 0) {
                    html += buildTotalsRow(`All Projects (${allData.length})`, allTotals, 'all-projects-total-seibies');
                }
                if (filtered.length > 0) {
                    html += buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total-seibies');
                }
                totalsHeader.innerHTML = html;
                applyStickyHeaderOffsets();
            };
            const renderCurrentPage = () => {
                if (filteredData.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No SEIBIES records found</td></tr>`;
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    return;
                }
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, filteredData.length);
                const pageData = filteredData.slice(startIndex, endIndex);
                renderTableWithMonthColumns(pageData);
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                setTimeout(() => {
                    renderTotalsHeader(displayMonths, currentMonthKey);
                }, 100);
            };
            const attachEditTriggerListeners = () => {
                document.querySelectorAll('.edit-trigger-seibies').forEach(link => {
                    link.addEventListener('click', (e) => {
                        e.preventDefault();
                        const id = parseInt(e.currentTarget.dataset.id);
                        editForecast(id);
                    });
                });
            };
            const renderTableWithMonthColumns = (data) => {
                if (!data.length) {
                    tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No SEIBIES records found</td></tr>`;
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    return;
                }
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                renderTotalsHeader(displayMonths, currentMonthKey);
                const mainHeaderHTML = `
          <tr>
            <th rowspan="2" class="project-ref-col-seibies" style="min-width:120px; max-width:120px; width:120px;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Project Reference</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectRef' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="projectRef" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectRef' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="projectRef" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <th rowspan="2" class="customer-col-seibies" style="min-width:130px; max-width:130px; width:130px;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Customer</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="customer" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="customer" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <th rowspan="2" class="project-col-seibies" style="min-width:180px; max-width:180px; width:180px;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Project</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="project" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="project" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <th rowspan="2" class="gh-col-seibies" style="min-width:100px; max-width:100px; width:100px;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Group Head</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="groupHead" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="groupHead" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <th rowspan="2" class="pm-col-seibies" style="min-width:100px; max-width:100px; width:100px;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Project Manager</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="projectManager" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="projectManager" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <th rowspan="2" class="obl-col-seibies" style="min-width:90px; max-width:90px; width:90px;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">SEI-BIES</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="obl" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="obl" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            ${displayMonths.map((month, monthIndex) => {
                    const isCurrentMonth = month === currentMonthKey;
                    const monthLabel = formatMonthLabel(new Date(month + "-01"));
                    const displayLabel = isCurrentMonth ? `${monthLabel} (Current Month)` : monthLabel;
                    const headerClass = isCurrentMonth ? 'current-month-header-seibies' : '';
                    return `
                <th colspan="2" class="month-scroll-col-seibies month-column-header-seibies ${headerClass}" style="min-width:60px;">
                  <div class="sort-header-container-seibies">
                    <span class="header-label-seibies">${displayLabel}</span>
                    <div class="sort-buttons-container-seibies">
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort by Month Total Ascending">▲</button>
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="desc" data-month-index="${monthIndex}" title="Sort by Month Total Descending">▼</button>
                    </div>
                  </div>
                </th>
              `;
                }).join('')}
          </tr>
          
          <tr>
            ${displayMonths.map((month, monthIndex) => {
                    const isCurrentMonth = month === currentMonthKey;
                    const subHeaderClass = isCurrentMonth ? 'current-month-header-table-seibies' : '';
                    return `
                <th class="month-scroll-col-seibies month-subheader-seibies ${subHeaderClass}" style="min-width:60px;">
                  <div class="sort-header-container-seibies">
                    <span>Previous</span>
                    <div class="sort-buttons-container-seibies">
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Ascending">▲</button>
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Descending">▼</button>
                    </div>
                  </div>
                </th>
                <th class="month-scroll-col-seibies month-subheader-seibies ${subHeaderClass}" style="min-width:60px;">
                  <div class="sort-header-container-seibies">
                    <span>Current</span>
                    <div class="sort-buttons-container-seibies">
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Ascending">▲</button>
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Descending">▼</button>
                    </div>
                  </div>
                </th>
              `;
                }).join('')}
          </tr>
        `;
                monthColumnsHeader.innerHTML = mainHeaderHTML;
                let tableBodyHTML = '';
                data.forEach((forecast) => {
                    var _a;
                    const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                    const proj = projects.find(p => p.Title === projectTitle);
                    const forecastData = safeJsonParse(forecast.ForecastData, {});
                    const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-US');
                    const projectRefText = (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || `#${forecast.Id}`;
                    const projectRefLink = `<a href="#" class="edit-trigger-seibies" data-id="${forecast.Id}" title="Open / edit this forecast">${projectRefText}</a>`;
                    const projectName = projectTitle;
                    let projectManagerDisplay = "-";
                    if (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) {
                        projectManagerDisplay = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
                            ? proj.ProjectManager.Title || "-"
                            : proj.ProjectManager || "-";
                    }
                    let rowHTML = `
            <tr data-forecast-id="${forecast.Id}">
              <td class="project-ref-col-seibies" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
              
              <td class="customer-col-seibies" style="min-width:130px; max-width:130px; width:130px;">${(proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-"}</td>
              
              <td class="project-col-seibies" style="min-width:180px; max-width:180px; width:180px;">${projectName}</td>
              
              <td class="gh-col-seibies" style="min-width:100px; max-width:100px; width:100px;">${getGroupHeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.GroupHead1)}</td>
              
              <td class="pm-col-seibies" style="min-width:100px; max-width:100px; width:100px;">${projectManagerDisplay}</td>
              
              <td class="obl-col-seibies" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
          `;
                    displayMonths.forEach((monthKey) => {
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                        const planValue = monthData.plan || 0;
                        const currentValue = monthData.actual || 0;
                        rowHTML += `
              <td class="month-scroll-col-seibies" style="min-width:60px;">${planValue.toLocaleString('en-US')}</td>
              <td class="month-scroll-col-seibies" style="min-width:60px;">${currentValue.toLocaleString('en-US')}</td>
            `;
                    });
                    rowHTML += `</tr>`;
                    tableBodyHTML += rowHTML;
                });
                tableBody.innerHTML = tableBodyHTML;
                attachSortButtonListeners();
                attachEditTriggerListeners();
                applyStickyHeaderOffsets();
            };
            const attachSortButtonListeners = () => {
                document.querySelectorAll('.sort-btn-seibies[data-sort]').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const button = e.currentTarget;
                        const sortColumn = button.dataset.sort;
                        const direction = button.dataset.direction;
                        const subColumn = button.dataset.subcolumn;
                        const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;
                        if (sortColumn && direction) {
                            handleSeibiesSort(sortColumn, direction, subColumn, monthIndex);
                        }
                    });
                });
            };
            const handleSeibiesSort = (column, direction, subColumn, monthIndex) => {
                currentSortConfig = {
                    column,
                    direction,
                    subColumn,
                    monthIndex
                };
                const dataToSort = filteredData.length > 0 ? filteredData : forecasts;
                if (dataToSort.length === 0)
                    return;
                const sortedData = [...dataToSort].sort((a, b) => {
                    var _a, _b;
                    let valueA = 0;
                    let valueB = 0;
                    switch (column) {
                        case 'projectRef':
                            const projA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            valueA = (projA === null || projA === void 0 ? void 0 : projA.ProjectReference) || "";
                            valueB = (projB === null || projB === void 0 ? void 0 : projB.ProjectReference) || "";
                            break;
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
                            const monthDataA = dataA[monthKeyToUse] || { plan: 0, actual: 0 };
                            const monthDataB = dataB[monthKeyToUse] || { plan: 0, actual: 0 };
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
                            }
                            else {
                                valueA = monthDataA.actual || 0;
                                valueB = monthDataB.actual || 0;
                            }
                            break;
                        case 'project':
                            valueA = ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title) || "";
                            valueB = ((_b = b.Project) === null || _b === void 0 ? void 0 : _b.Title) || "";
                            break;
                        case 'customer':
                            const projACust = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projBCust = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            valueA = (projACust === null || projACust === void 0 ? void 0 : projACust.CustomerName) || "";
                            valueB = (projBCust === null || projBCust === void 0 ? void 0 : projBCust.CustomerName) || "";
                            break;
                        case 'projectManager':
                            const projMA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projMB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            if (projMA === null || projMA === void 0 ? void 0 : projMA.ProjectManager) {
                                valueA = typeof projMA.ProjectManager === "object" && projMA.ProjectManager !== null
                                    ? projMA.ProjectManager.Title || ""
                                    : projMA.ProjectManager || "";
                            }
                            else {
                                valueA = "";
                            }
                            if (projMB === null || projMB === void 0 ? void 0 : projMB.ProjectManager) {
                                valueB = typeof projMB.ProjectManager === "object" && projMB.ProjectManager !== null
                                    ? projMB.ProjectManager.Title || ""
                                    : projMB.ProjectManager || "";
                            }
                            else {
                                valueB = "";
                            }
                            break;
                        case 'groupHead':
                            const projGA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projGB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            if (projGA === null || projGA === void 0 ? void 0 : projGA.GroupHead1) {
                                valueA = typeof projGA.GroupHead1 === "object" && projGA.GroupHead1 !== null
                                    ? projGA.GroupHead1.Title || ""
                                    : projGA.GroupHead1 || "";
                            }
                            else {
                                valueA = "";
                            }
                            if (projGB === null || projGB === void 0 ? void 0 : projGB.GroupHead1) {
                                valueB = typeof projGB.GroupHead1 === "object" && projGB.GroupHead1 !== null
                                    ? projGB.GroupHead1.Title || ""
                                    : projGB.GroupHead1 || "";
                            }
                            else {
                                valueB = "";
                            }
                            break;
                        default:
                            return 0;
                    }
                    if (typeof valueA === 'number' && typeof valueB === 'number') {
                        return direction === 'asc' ? valueA - valueB : valueB - valueA;
                    }
                    else {
                        const strA = String(valueA).toLowerCase();
                        const strB = String(valueB).toLowerCase();
                        return direction === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
                    }
                });
                if (filteredData.length > 0) {
                    filteredData = sortedData;
                }
                else {
                    forecasts = sortedData;
                }
                currentPage = 1;
                updatePaginationControls();
                renderCurrentPage();
                const sortDirectionText = direction === 'asc' ? 'Ascending' : 'Descending';
                let columnName = '';
                if (subColumn) {
                    columnName = subColumn.replace('-', ' ').replace(/\b\w/g, l => l.toUpperCase());
                }
                else if (column === 'projectRef') {
                    columnName = 'Project Reference';
                }
                else if (column === 'obl') {
                    columnName = 'SEI-BIES';
                }
                else if (column === 'project') {
                    columnName = 'Project';
                }
                else if (column === 'customer') {
                    columnName = 'Customer';
                }
                else if (column === 'projectManager') {
                    columnName = 'Project Manager';
                }
                else if (column === 'groupHead') {
                    columnName = 'Group Head';
                }
                else {
                    const displayMonths = getDisplayMonths();
                    const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
                    const monthKey = displayMonths[monthIndexToUse];
                    const monthDate = new Date(monthKey + "-01");
                    columnName = formatMonthLabel(monthDate);
                }
                toast("info", `Sorted by ${columnName} (${sortDirectionText})`);
            };
            const clearFilters = () => {
                filterYear.value = "";
                filterMonth.value = "";
                filterProject.value = "";
                filterCustomer.value = "";
                filterProjectManager.value = "";
                filterGroupHead.value = "";
                filterProjectReference.value = "";
                currentSortConfig = {
                    column: '',
                    direction: 'asc'
                };
                initializePagination(forecasts);
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                setTimeout(() => {
                    renderTotalsHeader(displayMonths, currentMonthKey);
                }, 100);
                toast("info", "All SEIBIES filters and sorting cleared");
            };
            // ========== RENDER PROJECT DETAILS ==========
            const renderProjectDetails = (project) => {
                var _a;
                const projectNameDisplay = document.querySelector('#projectNameDisplaySeibies');
                if (projectNameDisplay)
                    projectNameDisplay.textContent = project.Title || "-";
                const projectRef = document.querySelector('#projectRefSeibies');
                if (projectRef)
                    projectRef.textContent = project.ProjectReference || "-";
                const profitCenter = document.querySelector('#profitCenterSeibies');
                if (profitCenter)
                    profitCenter.textContent = project.PC || "-";
                const projectManager = document.querySelector('#projectManagerSeibies');
                if (projectManager) {
                    projectManager.textContent = typeof project.ProjectManager === "object" ?
                        (_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title : project.ProjectManager || "-";
                }
                const groupHead = document.querySelector('#groupHeadSeibies');
                if (groupHead) {
                    groupHead.textContent = getGroupHeadDisplayName(project.GroupHead1);
                }
                const customer = document.querySelector('#customerSeibies');
                if (customer)
                    customer.textContent = project.CustomerName || "-";
            };
            // ========== HISTORICAL TOGGLE ==========
            const setupHistoricalToggle = () => {
                const toggleBtn = document.querySelector("#toggleHistoricalBtnSeibies");
                const historicalContainer = document.querySelector("#historicalContainerSeibies");
                const header = document.querySelector(".historical-toggle-header-seibies");
                if (!toggleBtn || !historicalContainer || !header)
                    return;
                let isVisible = false;
                const updateHistoricalToggle = () => {
                    const arrow = toggleBtn.querySelector(".toggle-arrow-seibies");
                    const label = toggleBtn.querySelector(".toggle-label-seibies");
                    if (isVisible) {
                        historicalContainer.style.display = "grid";
                        toggleBtn.setAttribute("aria-expanded", "true");
                        toggleBtn.title = "Hide historical data";
                        if (arrow)
                            arrow.textContent = "▲";
                        if (label)
                            label.textContent = "Hide";
                    }
                    else {
                        historicalContainer.style.display = "none";
                        toggleBtn.setAttribute("aria-expanded", "false");
                        toggleBtn.title = "Show historical data";
                        if (arrow)
                            arrow.textContent = "▼";
                        if (label)
                            label.textContent = "Show";
                    }
                };
                const toggleHistorical = (event) => {
                    if (event) {
                        event.preventDefault();
                        event.stopPropagation();
                    }
                    isVisible = !isVisible;
                    updateHistoricalToggle();
                };
                toggleBtn.onclick = (event) => {
                    toggleHistorical(event);
                };
                header.onclick = (event) => {
                    const target = event.target;
                    if (target.closest("#toggleHistoricalBtnSeibies")) {
                        return;
                    }
                    toggleHistorical(event);
                };
                updateHistoricalToggle();
            };
            // ========== FORM MANAGEMENT FUNCTIONS ==========
            const canViewForecastForm = (isEdit = false, project) => {
                if (!canRead())
                    return false;
                if (isEdit) {
                    if (canEditProject(project))
                        return true;
                    if (canReadAllProjects())
                        return true;
                    if (project && getAssignedProjects(projects).some(p => p.Id === project.Id))
                        return true;
                    return false;
                }
                return canCreateRecord();
            };
            const showForm = (isEditMode = false, project) => {
                console.log(`📝 showForm called: isEditMode=${isEditMode}, project=${project === null || project === void 0 ? void 0 : project.Title}`);
                if (!canViewForecastForm(isEditMode, project)) {
                    if (!canRead()) {
                        toast("error", "❌ You don't have permission to access this form");
                        return;
                    }
                    if (isEditMode) {
                        toast("error", "❌ You don't have permission to view this forecast");
                        return;
                    }
                    toast("error", "❌ You don't have permission to create SEIBIES forecasts");
                    return;
                }
                formSection.style.display = "block";
                tableSection.style.display = "none";
                if (showFormBtn)
                    showFormBtn.style.display = "none";
                // ✅ Show/hide BOTH save buttons together (top + bottom)
                const canSave = isEditMode ? canEditProject(project) : canCreateRecord();
                const saveBtns = getSaveButtons();
                saveBtns.forEach(btn => {
                    btn.style.display = canSave ? '' : 'none';
                });
                console.log(`💾 Save buttons ${canSave ? 'visible' : 'hidden'} (${saveBtns.length} button(s))`);
                if (isEditMode) {
                    formSection.classList.add('edit-mode-seibies');
                    projectSelect.disabled = true;
                    if (formTitle) {
                        const projectName = "SEI BIES Liquidation 5QRF";
                        // const projectReference = project?.ProjectReference || "";
                        formTitle.textContent = projectName
                            ? `${projectName}`
                            : "SEI BIES Liquidation 5QRF";
                        formTitle.style.color = '';
                    }
                }
                else {
                    formSection.classList.remove('edit-mode-seibies');
                    projectSelect.disabled = false;
                    if (formTitle) {
                        formTitle.textContent = "Add SEIBIES Forecast";
                        formTitle.style.color = '';
                    }
                }
            };
            const hideForm = () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                if (showFormBtn)
                    showFormBtn.style.display = "inline-block";
                formSection.classList.remove('edit-mode-seibies');
                projectSelect.disabled = false;
            };
            const resetForm = () => {
                editingId = null;
                form.reset();
                projectSelect.value = "";
                const projectNameDisplay = document.querySelector('#projectNameDisplaySeibies');
                if (projectNameDisplay)
                    projectNameDisplay.textContent = "-";
                const projectRef = document.querySelector('#projectRefSeibies');
                if (projectRef)
                    projectRef.textContent = "-";
                const profitCenter = document.querySelector('#profitCenterSeibies');
                if (profitCenter)
                    profitCenter.textContent = "-";
                const projectManager = document.querySelector('#projectManagerSeibies');
                if (projectManager)
                    projectManager.textContent = "-";
                const groupHead = document.querySelector('#groupHeadSeibies');
                if (groupHead)
                    groupHead.textContent = "-";
                const customer = document.querySelector('#customerSeibies');
                if (customer)
                    customer.textContent = "-";
                historicalContainer.innerHTML = "";
                firstRowContainer.innerHTML = "";
                remainingContainer.innerHTML = "";
                yearlyContainer.innerHTML = "";
                toggleForecastSections(false);
                const oblContainer = document.querySelector('.obl-section-container-seibies');
                if (oblContainer) {
                    oblContainer.className = 'obl-section-container-seibies hidden-section';
                }
                oblLastMonthInput.removeAttribute('readonly');
                oblLastMonthInput.placeholder = "Enter OBL";
                setNumericValue(oblLastMonthInput, 0);
                oblValidationMessage.textContent = "✅ All values are within OBL limit";
                oblValidationMessage.className = 'validation-message-seibies';
                updateOBLValidationDisplay();
            };
            // ========== VALIDATION FUNCTIONS ==========
            const calculateTotalCurrentValue = () => {
                let total = 0;
                const editableMonths = allMonths.slice(3);
                editableMonths.forEach(monthKey => {
                    const currentInputs = document.querySelectorAll(`input.main-current-input-seibies[data-month="${monthKey}"]:not([readonly])`);
                    currentInputs.forEach(input => {
                        total += getNumericValue(input);
                    });
                });
                const yearlyCurrentInputs = document.querySelectorAll(`input.yearly-current-input-seibies:not([readonly])`);
                yearlyCurrentInputs.forEach(input => {
                    total += getNumericValue(input);
                });
                return total;
            };
            const validateOBLLimit = (changedInput) => {
                const oblValue = getNumericValue(oblLastMonthInput);
                const totalCurrent = calculateTotalCurrentValue();
                totalCurrentValueSpan.textContent = totalCurrent.toLocaleString('en-US');
                const remaining = Math.abs(oblValue) - totalCurrent;
                remainingValueSpan.textContent = remaining.toLocaleString('en-US');
                if (oblValue === 0) {
                    oblLineFill.style.width = `0%`;
                    oblLineFill.className = 'line-fill-seibies';
                    if (totalCurrent > 0) {
                        showOBLErrorToast('⚠️ OBL is zero but current values exist. Please set OBL value.', true);
                        oblValidationMessage.textContent = '⚠️ OBL is zero but current values exist. Please set OBL value.';
                        oblValidationMessage.className = 'validation-message-seibies danger-seibies';
                        blockAllCurrentInputs();
                        return false;
                    }
                    else {
                        const existingToast = document.querySelector('.obl-error-toast-seibies');
                        if (existingToast)
                            existingToast.remove();
                        oblValidationMessage.textContent = '✅ All values are within SEI-BIES limit';
                        oblValidationMessage.className = 'validation-message-seibies success-seibies';
                        unblockAllCurrentInputs();
                        return true;
                    }
                }
                const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
                oblLineFill.style.width = `${percentage}%`;
                oblLineFill.className = 'line-fill-seibies';
                if (percentage > 90) {
                    oblLineFill.classList.add('danger-seibies');
                }
                else if (percentage > 75) {
                    oblLineFill.classList.add('warning-seibies');
                }
                oblValidationMessage.className = 'validation-message-seibies';
                if (totalCurrent > oblValue) {
                    const exceedAmount = totalCurrent - oblValue;
                    showOBLErrorToast(`❌ Total current value exceeds SEI-BIES by ${exceedAmount.toLocaleString('en-US')}! Reduce values to continue.`);
                    oblValidationMessage.textContent = `❌ Total current value exceeds SEI-BIES by ${exceedAmount.toLocaleString('en-US')}! Reduce values to continue.`;
                    oblValidationMessage.classList.add('danger-seibies');
                    blockIncreasingCurrentInputs(changedInput);
                    return false;
                }
                else if (percentage > 90) {
                    showOBLErrorToast('⚠️ SEI-BIES limit nearly reached! Consider reducing values.', true);
                    oblValidationMessage.textContent = '⚠️ SEI-BIES limit nearly reached! Consider reducing values.';
                    oblValidationMessage.classList.add('warning-seibies');
                    unblockAllCurrentInputs();
                    return true;
                }
                else {
                    const existingToast = document.querySelector('.obl-error-toast-seibies');
                    if (existingToast)
                        existingToast.remove();
                    oblValidationMessage.textContent = '✅ All values are within SEI-BIES limit';
                    oblValidationMessage.classList.add('success-seibies');
                    unblockAllCurrentInputs();
                    return true;
                }
            };
            const blockAllCurrentInputs = () => {
                const currentInputs = document.querySelectorAll('input[data-type="actual"]:not([readonly])');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-seibies:not([readonly])');
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
            const blockIncreasingCurrentInputs = (allowedInput) => {
                const currentInputs = document.querySelectorAll('input[data-type="actual"]:not([readonly])');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-seibies:not([readonly])');
                currentInputs.forEach(input => {
                    if (input === allowedInput) {
                        input.style.backgroundColor = '';
                        input.style.cursor = '';
                        input.removeAttribute('readonly');
                    }
                    else {
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
                    }
                    else {
                        input.setAttribute('data-block-increase', 'true');
                        input.style.backgroundColor = '#fff3cd';
                        input.style.cursor = 'not-allowed';
                    }
                });
            };
            const unblockAllCurrentInputs = () => {
                const currentInputs = document.querySelectorAll('input[data-type="actual"]:not([data-month="0"])');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-seibies');
                currentInputs.forEach(input => {
                    input.style.backgroundColor = '';
                    input.style.cursor = '';
                    input.removeAttribute('data-block-increase');
                    input.removeAttribute('readonly');
                });
                yearlyCurrentInputs.forEach(input => {
                    input.style.backgroundColor = '';
                    input.style.cursor = '';
                    input.removeAttribute('data-block-increase');
                    input.removeAttribute('readonly');
                });
            };
            const allowInputIncrease = (input, newValue) => {
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
            const updateOBLValidationDisplay = (changedInput) => {
                validateOBLLimit(changedInput);
            };
            // ========== FORECAST FORM FUNCTIONS ==========
            const toggleForecastSections = (show) => {
                const sections = [
                    '.obl-section-container-seibies',
                    '#historicalForecastSectionSeibies',
                    '#detailedForecastSectionSeibies',
                    '#longtermForecastSectionSeibies',
                    '#yearlyForecastSectionSeibies'
                ];
                sections.forEach(selector => {
                    const element = document.querySelector(selector);
                    if (element) {
                        if (show) {
                            element.classList.remove('hidden-section');
                            element.classList.add('visible-section-seibies');
                        }
                        else {
                            element.classList.remove('visible-section-seibies');
                            element.classList.add('hidden-section');
                        }
                    }
                });
            };
            // ========== REMAINING FUNCTIONS ==========
            const onProjectChange = (projectId, existingForecast) => {
                const project = projects.find(p => p.Id === projectId);
                if (!project)
                    return;
                toggleForecastSections(true);
                renderProjectDetails(project);
                let forecastData = {};
                if (existingForecast) {
                    forecastData = safeJsonParse(existingForecast.ForecastData, {});
                }
                const isNewProject = !existingForecast;
                const oblContainer = document.querySelector('.obl-section-container-seibies');
                const fieldNote = oblContainer === null || oblContainer === void 0 ? void 0 : oblContainer.querySelector('.field-note-seibies');
                if (isNewProject) {
                    if (oblContainer) {
                        oblContainer.className = 'obl-section-container-seibies visible-section-seibies new-project-seibies';
                    }
                    oblLastMonthInput.removeAttribute('readonly');
                    oblLastMonthInput.placeholder = "Enter initial OBL for new project";
                    if (fieldNote)
                        fieldNote.textContent = "For new projects: Enter initial OBL value";
                    oblValidationMessage.textContent = "Enter initial OBL value for new project";
                    oblValidationMessage.className = 'validation-message-seibies';
                    const defaultOBL = project.PC ? 0 : 0;
                    setNumericValue(oblLastMonthInput, defaultOBL);
                }
                else {
                    if (oblContainer) {
                        oblContainer.className = 'obl-section-container-seibies visible-section-seibies existing-project-seibies';
                    }
                    oblLastMonthInput.setAttribute('readonly', 'true');
                    const oblValue = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                    setNumericValue(oblLastMonthInput, oblValue);
                    oblLastMonthInput.placeholder = "SEI BIES";
                    if (fieldNote)
                        fieldNote.textContent = "OBL value from database record";
                    if (oblValue > 0) {
                        oblLastMonthInput.value = oblValue.toString();
                        oblLastMonthInput.className = oblLastMonthInput.className.replace('input-empty-seibies', 'input-with-value-seibies');
                    }
                }
                renderForecastInterface(forecastData, existingForecast);
            };
            const getForecastMonths = () => {
                return generateIdealMonthStructure();
            };
            const renderForecastInterface = (existingData, existingForecast) => {
                allMonths = getForecastMonths();
                const displayOBL = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                setNumericValue(oblLastMonthInput, displayOBL);
                renderHistoricalMonths(allMonths.slice(0, 3), existingData);
                renderFirstRow(allMonths.slice(3, 6), existingData);
                renderRemainingMonths(allMonths.slice(6), existingData);
                renderYearlyForecast(existingData.yearly || createFreshYearlyStructure(getLastMonthKey(existingData)));
                protectHistoricalElements();
                updateOBLValidationDisplay();
                attachForecastEventListeners();
                setTimeout(() => {
                    setupHistoricalToggle();
                }, 50);
            };
            // ========== RENDER FUNCTIONS FOR FORECAST SECTIONS ==========
            const renderHistoricalMonths = (months, existingData) => {
                let html = '';
                months.forEach((monthKey, index) => {
                    const monthDate = new Date(monthKey + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[monthKey] || { plan: 0, actual: 0 };
                    html += renderHistoricalMonthBox(monthKey, monthLabel, data, index);
                });
                historicalContainer.innerHTML = html;
            };
            const renderHistoricalMonthBox = (monthKey, monthLabel, data, index) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="month-box-seibies historical-month-seibies">
            <div class="month-header-seibies">${monthLabel} (Historical - Readonly)</div>

            <div class="month-block-seibies">
              <div class="main-inputs-container-seibies">
                <div class="main-input-group-seibies">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input-seibies ${planInputClass} number-input-seibies" />
                </div>
                <div class="main-input-group-seibies">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        readonly class="main-current-input-seibies ${actualInputClass} number-input-seibies" />
                </div>
              </div>
            </div>
          </div>
        `;
            };
            const renderFirstRow = (months, existingData) => {
                let html = '';
                html += `<div class="detailed-months-row-seibies">`;
                const currentMonth = months[0];
                if (currentMonth) {
                    const monthDate = new Date(currentMonth + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[currentMonth] || { plan: 0, actual: 0 };
                    html += `<div class="current-month-container-seibies">`;
                    html += renderDetailedMonthBox(currentMonth, monthLabel, data, true);
                    html += `</div>`;
                }
                html += `<div class="next-months-container-seibies">`;
                for (let i = 1; i <= 2; i++) {
                    if (i < months.length) {
                        const monthKey = months[i];
                        const monthDate = new Date(monthKey + "-01");
                        const monthLabel = formatMonthLabel(monthDate);
                        const data = existingData[monthKey] || { plan: 0, actual: 0 };
                        html += `<div class="next-month-box-seibies">`;
                        html += renderDetailedMonthBox(monthKey, monthLabel, data, false);
                        html += `</div>`;
                    }
                }
                html += `</div>`;
                html += `</div>`;
                firstRowContainer.innerHTML = html;
            };
            const renderRemainingMonths = (months, existingData) => {
                let html = '';
                for (let row = 0; row < 3; row++) {
                    html += `<div class="remaining-row-seibies">`;
                    for (let col = 0; col < 5; col++) {
                        const index = (row * 5) + col;
                        if (index >= months.length)
                            break;
                        const monthKey = months[index];
                        const monthDate = new Date(monthKey + "-01");
                        const monthLabel = formatMonthLabel(monthDate);
                        const data = existingData[monthKey] || { plan: 0, actual: 0 };
                        html += renderSimpleMonthBox(monthKey, monthLabel, data);
                    }
                    html += `</div>`;
                }
                remainingContainer.innerHTML = html;
            };
            const renderYearlyForecast = (yearlyData) => {
                let html = '';
                const years = Object.keys(yearlyData).sort();
                years.forEach(year => {
                    const data = yearlyData[year] || { plan: 0, actual: 0 };
                    html += renderYearlyBox(year, data);
                });
                yearlyContainer.innerHTML = html;
            };
            const renderYearlyBox = (year, data) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="yearly-box-seibies">
            <div class="yearly-header-seibies">${year}</div>

            <div class="yearly-inputs-container-seibies">
              <div class="yearly-input-group-seibies">
                <label>Previous</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                      readonly class="yearly-plan-input-seibies ${planInputClass} number-input-seibies" />
              </div>
              <div class="yearly-input-group-seibies">
                <label>Current</label>
                <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                      class="yearly-current-input-seibies ${actualInputClass} number-input-seibies" />
              </div>
            </div>
          </div>
        `;
            };
            const getCurrentMonthMaxDate = () => {
                const now = new Date();
                const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
                return `${lastDay.getFullYear()}-${String(lastDay.getMonth() + 1).padStart(2, "0")}-${String(lastDay.getDate()).padStart(2, "0")}`;
            };
            /**
             * ✅ renderDetailedMonthBox
             *    - Current month  : Previous / Current / Date  (Further input REMOVED)
             *    - Next 2 months  : Previous / Current
             *    The data model still carries monthData.forecast so previously-saved
             *    "Further" values are preserved on save (just not shown/editable here).
             */
            const renderDetailedMonthBox = (monthKey, monthLabel, data, isCurrentMonth = false) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const displayDateValue = data.date || '';
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const dateInputClass = displayDateValue ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="detailed-month-box-seibies ${isCurrentMonth ? 'current-month-box-seibies' : 'next-month-box-seibies'}">
            <div class="month-header-seibies">
              ${monthLabel}
              ${isCurrentMonth ? '<span class="current-month-indicator-seibies">Current</span>' : ''}
            </div>

            <div class="month-block-seibies">
              <div class="enhanced-inputs-container-seibies ${isCurrentMonth ? 'current-month-inputs-seibies' : ''}">
                <div class="enhanced-input-group-seibies">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input-seibies ${planInputClass} number-input-seibies" />
                </div>

                <div class="enhanced-input-group-seibies">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        class="main-current-input-seibies ${actualInputClass} number-input-seibies" />
                </div>

                ${isCurrentMonth ? `
                  <div class="enhanced-input-group-seibies">
                    <label>Date</label>
                    <input type="date" value="${displayDateValue}" data-type="date" data-month="${monthKey}" max="${getCurrentMonthMaxDate()}" title="Select a date up to the end of the current month"
                          class="forecast-date-input-seibies ${dateInputClass}" />
                  </div>
                ` : ''}
              </div>
            </div>
          </div>
        `;
            };
            const renderSimpleMonthBox = (monthKey, monthLabel, data) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="month-box-seibies simple-month-seibies">
            <div class="month-header-seibies">${monthLabel}</div>

            <div class="main-inputs-container-seibies">
              <div class="main-input-group-seibies">
                <label>Previous</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                      readonly class="main-plan-input-seibies ${planInputClass} number-input-seibies" />
              </div>
              <div class="main-input-group-seibies">
                <label>Current</label>
                <input type="text" value="${displayActualValue}" id="actualCurrent" data-type="actual" data-month="${monthKey}"
                      class="main-current-input-seibies ${actualInputClass} number-input-seibies" />
              </div>
            </div>
          </div>
        `;
            };
            const protectHistoricalElements = () => {
                document.querySelectorAll('.historical-month-seibies input, .historical-month-seibies select, .historical-month-seibies button').forEach(element => {
                    var _a;
                    const newElement = element.cloneNode(true);
                    (_a = element.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newElement, element);
                });
                document.querySelectorAll('.historical-month-seibies').forEach(section => {
                    section.addEventListener('click', (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        toast("warning", "Historical SEIBIES data cannot be modified");
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
            const attachForecastEventListeners = () => {
                document.querySelectorAll('.number-input-seibies').forEach(input => {
                    allowOnlyNumbers(input);
                });
                document.querySelectorAll('.number-input-seibies:not([readonly])').forEach(input => {
                    input.addEventListener('focus', function () {
                        this.setAttribute('data-old-value', getNumericValue(this).toString());
                    });
                });
                document.querySelectorAll('.main-current-input-seibies:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
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
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                        }
                        else {
                            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.yearly-current-input-seibies:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
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
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                        }
                        else {
                            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.forecast-date-input-seibies:not([readonly])').forEach(input => {
                    input.addEventListener('change', function () {
                        const maxAllowed = this.getAttribute('max') || getCurrentMonthMaxDate();
                        if (this.value && this.value > maxAllowed) {
                            this.value = '';
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                            toast("warning", `Date cannot be beyond the current month (max ${maxAllowed})`);
                            return;
                        }
                        if (this.value === '') {
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                        }
                        else {
                            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                        }
                    });
                });
                const validationInputs = document.querySelectorAll(`
          .main-current-input-seibies:not([readonly]),
          .yearly-current-input-seibies:not([readonly])
        `);
                validationInputs.forEach(input => {
                    input.addEventListener('input', function () {
                        setTimeout(() => {
                            updateOBLValidationDisplay(this);
                        }, 100);
                    });
                });
            };
            const editForecast = async (id) => {
                editingId = id;
                const forecast = forecasts.find(f => f.Id === id);
                if (!forecast) {
                    toast("error", "SEIBIES forecast not found");
                    return;
                }
                if (!canRead()) {
                    toast("error", "❌ You don't have permission to view this forecast");
                    return;
                }
                const project = projects.find(p => { var _a; return p.Id === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Id); });
                const canEdit = canEditProject(project);
                const isFrozen = forecast.Freeze || false;
                showForm(true, project);
                let forecastData = {};
                try {
                    let forecastDataString = forecast.ForecastData || "{}";
                    forecastDataString = forecastDataString.trim();
                    if (!forecastDataString || forecastDataString === '""' || forecastDataString === "null") {
                        forecastDataString = "{}";
                    }
                    forecastData = JSON.parse(forecastDataString);
                }
                catch (e) {
                    forecastData = createFreshForecastStructure();
                    toast("warning", "Corrupted SEIBIES forecast data detected. Loaded empty structure.");
                }
                forecastData = propagateMonthValues(forecastData);
                if (forecast.Project) {
                    projectSelect.value = String(forecast.Project.Id);
                    onProjectChange(forecast.Project.Id, forecast);
                }
                setTimeout(() => {
                    const isReadOnly = isFrozen || !canEdit;
                    setFormReadOnly(isReadOnly);
                    // ✅ Show/hide BOTH save buttons together (top + bottom)
                    const saveBtns = getSaveButtons();
                    saveBtns.forEach(btn => {
                        btn.style.display = isReadOnly ? 'none' : '';
                    });
                    const formTitleEl = document.querySelector("#formTitleSeibies");
                    if (formTitleEl) {
                        const projectName = "SEI BIES Liquidation 5QRF";
                        // const projectReference = project?.ProjectReference || "";
                        const projectTitle = projectName
                            ? `${projectName}`
                            : "";
                        if (isFrozen) {
                            formTitleEl.innerHTML = `🔒 ${projectTitle || "SEIBIES Forecast"} (Frozen - Read Only)`;
                            formTitleEl.style.color = '#856404';
                        }
                        else if (!canEdit) {
                            formTitleEl.innerHTML = `🔒 ${projectTitle || "SEIBIES Forecast"} (Read Only)`;
                            formTitleEl.style.color = '#856404';
                        }
                        else {
                            formTitleEl.innerHTML = projectTitle || "";
                            formTitleEl.style.color = '';
                        }
                    }
                }, 300);
                // toast("info", `Editing SEIBIES forecast for ${forecast.Title}`);
            };
            // ========== FORM SUBMISSION ==========
            // ============================================================
            // ✅ SAVE — SweetAlert confirmation FIRST (same look as procurement.ts)
            //    • 2 buttons only: Yes, Save / Cancel  (no "Don't Save" option)
            //    • z-index bumped so the loader cannot cover the dialog
            //    • isSavingSeibies guard prevents double-submit
            //    • saveBtns re-enabled in the outer finally, so no path can leave them stuck
            // ============================================================
            form.addEventListener("submit", async (e) => {
                e.preventDefault();
                // ✅ Only run the flow when a Save button triggered the submit
                const submitter = e.submitter;
                const isSaveClick = submitter && submitter.classList.contains("btn-primary-seibies");
                if (!isSaveClick)
                    return;
                if (isSavingSeibies) {
                    console.log("⛔ submit blocked: save already in progress");
                    return;
                }
                // ---- Pre-checks (nothing shown yet) ----
                const projectId = parseInt(projectSelect.value);
                const project = projects.find(p => p.Id === projectId);
                if (!project) {
                    toast("error", "Please select a project");
                    return;
                }
                if (editingId) {
                    if (!canEditProject(project)) {
                        toast("error", "❌ You don't have permission to edit this SEIBIES forecast");
                        return;
                    }
                }
                else {
                    if (!canCreateRecord()) {
                        toast("error", "❌ You don't have permission to create SEIBIES forecasts");
                        return;
                    }
                }
                const isNewProject = !editingId && !forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                if (!isNewProject && forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId && f.Id !== editingId; })) {
                    toast("error", "SEIBIES forecast already exists for this project!");
                    return;
                }
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
                        didOpen: bumpSwalZIndex,
                    });
                    if (!confirmZeroOBL.isConfirmed) {
                        oblLastMonthInput.focus();
                        return;
                    }
                }
                const totalCurrent = calculateTotalCurrentValue();
                if (oblValue > 0 && totalCurrent > oblValue) {
                    toast("error", "Total current value cannot exceed OBL!");
                    return;
                }
                // ✅ HOIST the save buttons so the outer finally can always restore them
                const saveBtns = getSaveButtons();
                isSavingSeibies = true;
                try {
                    // ========== 1) CONFIRMATION POPUP ==========
                    const confirmResult = await Swal.fire({
                        title: editingId ? "Save Changes?" : "Save SEIBIES Plan?",
                        html: editingId
                            ? `Do you want to save the changes made to <strong>${escapeHtml(project.Title)}</strong>?`
                            : `Do you want to save the SEIBIES forecast for <strong>${escapeHtml(project.Title)}</strong>?`,
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
                    // ========== 2) LOCK SAVE BUTTONS ==========
                    saveBtns.forEach(btn => {
                        btn.disabled = true;
                        btn.dataset.saving = "true";
                    });
                    // ========== 3) ACTUAL SAVE (unchanged logic) ==========
                    showRevenueLoader(editingId ? 'Updating' : 'Saving', editingId ? 'Updating the SEIBIES forecast on SharePoint...' : 'Saving the SEIBIES forecast to SharePoint...');
                    try {
                        const forecastData = {};
                        const allStorageMonths = getAllMonthsForStorage();
                        allStorageMonths.forEach(monthKey => {
                            var _a, _b;
                            const mainCurrentInput = document.querySelector(`input.main-current-input-seibies[data-month="${monthKey}"]`);
                            const mainPlanInput = document.querySelector(`input.main-plan-input-seibies[data-month="${monthKey}"]`);
                            const mainDateInput = document.querySelector(`input.forecast-date-input-seibies[data-month="${monthKey}"]`);
                            if (mainCurrentInput && mainPlanInput) {
                                const monthData = {
                                    plan: getNumericValue(mainPlanInput),
                                    actual: getNumericValue(mainCurrentInput)
                                };
                                // ✅ Preserve any previously-saved "Further" value
                                const existingData = safeJsonParse((_a = forecasts.find(f => f.Id === editingId)) === null || _a === void 0 ? void 0 : _a.ForecastData, {});
                                const prevForecast = (_b = existingData[monthKey]) === null || _b === void 0 ? void 0 : _b.forecast;
                                if (typeof prevForecast === 'number' && prevForecast > 0) {
                                    monthData.forecast = prevForecast;
                                }
                                if (mainDateInput && mainDateInput.value) {
                                    monthData.date = mainDateInput.value;
                                }
                                forecastData[monthKey] = monthData;
                            }
                        });
                        const yearlyData = {};
                        const yearlyPlanInputs = document.querySelectorAll('.yearly-plan-input-seibies');
                        if (yearlyPlanInputs.length > 0) {
                            const years = new Set();
                            yearlyPlanInputs.forEach(input => {
                                const year = input.dataset.year;
                                if (year)
                                    years.add(year);
                            });
                            Array.from(years).sort().forEach(year => {
                                const planInput = document.querySelector(`input.yearly-plan-input-seibies[data-year="${year}"]`);
                                const currentInput = document.querySelector(`input.yearly-current-input-seibies[data-year="${year}"]`);
                                if (planInput && currentInput) {
                                    yearlyData[year] = {
                                        plan: getNumericValue(planInput),
                                        actual: getNumericValue(currentInput)
                                    };
                                }
                            });
                            forecastData.yearly = yearlyData;
                        }
                        const propagatedData = propagateCollectionValues(forecastData);
                        const digest = await getDigest();
                        const payload = {
                            __metadata: { type: "SP.Data.SEIBIESListItem" },
                            Title: generateForecastTitle(project, propagatedData),
                            ProjectId: project.Id,
                            ProfitCentre: project.PC || "",
                            ForecastData: JSON.stringify(propagatedData),
                            UpdatedDate: new Date().toISOString()
                        };
                        if (!editingId && isNewProject) {
                            payload.OBL = oblValue;
                        }
                        if (!editingId) {
                            payload.CreatedDate = new Date().toISOString();
                        }
                        const url = editingId
                            ? `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${editingId})`
                            : `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
                        const headers = {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest
                        };
                        if (editingId) {
                            headers["IF-MATCH"] = "*";
                            headers["X-HTTP-Method"] = "MERGE";
                        }
                        const res = await fetch(url, {
                            method: "POST",
                            headers,
                            body: JSON.stringify(payload)
                        });
                        if (res.ok) {
                            // toast("success", editingId ? "SEIBIES Forecast Updated!" : "SEIBIES Forecast Added!");
                            await loadForecasts();
                            initializePagination(forecasts);
                            hideForm();
                            resetForm();
                        }
                        else {
                            const errorText = await res.text();
                            console.error("SharePoint Error:", errorText);
                            toast("error", `Error saving SEIBIES forecast: ${res.status}`);
                        }
                    }
                    catch (error) {
                        console.error("Error saving SEIBIES forecast:", error);
                        toast("error", "Error saving SEIBIES forecast!");
                    }
                    finally {
                        hideRevenueLoader();
                    }
                }
                catch (outerError) {
                    console.error("Unexpected error in SEIBIES submit:", outerError);
                    const message = outerError instanceof Error ? outerError.message : String(outerError);
                    toast("error", `Unexpected error: ${message}`);
                }
                finally {
                    // ✅ Always restore the buttons + clear the guard
                    saveBtns.forEach(btn => {
                        btn.disabled = false;
                        btn.dataset.saving = "false";
                    });
                    isSavingSeibies = false;
                }
            });
            // ========== SET FORM READONLY ==========
            const setFormReadOnly = (isReadOnly) => {
                const editableSelectors = [
                    '.main-current-input-seibies',
                    '.yearly-current-input-seibies',
                    '.forecast-date-input-seibies'
                ];
                const alwaysReadonlySelectors = [
                    '.main-plan-input-seibies',
                    '.yearly-plan-input-seibies'
                ];
                const inputs = document.querySelectorAll(editableSelectors.join(','));
                inputs.forEach(input => {
                    if (isReadOnly) {
                        input.setAttribute('readonly', 'true');
                        input.style.backgroundColor = '#f5f5f5';
                        input.style.color = '#6c757d';
                        input.style.cursor = 'not-allowed';
                        input.style.opacity = '0.7';
                        input.style.pointerEvents = 'none';
                    }
                    else {
                        input.removeAttribute('readonly');
                        input.style.backgroundColor = '';
                        input.style.color = '';
                        input.style.cursor = '';
                        input.style.opacity = '1';
                        input.style.pointerEvents = '';
                    }
                });
                const readonlyInputs = document.querySelectorAll(alwaysReadonlySelectors.join(','));
                readonlyInputs.forEach(input => {
                    input.setAttribute('readonly', 'true');
                    if (isReadOnly) {
                        input.style.backgroundColor = '#f5f5f5';
                        input.style.color = '#6c757d';
                        input.style.cursor = 'not-allowed';
                        input.style.opacity = '0.7';
                        input.style.pointerEvents = 'none';
                    }
                    else {
                        input.style.backgroundColor = '';
                        input.style.color = '';
                        input.style.cursor = '';
                        input.style.opacity = '';
                        input.style.pointerEvents = '';
                    }
                });
                const oblInput = document.querySelector('#oblLastMonthSeibies');
                if (oblInput) {
                    oblInput.setAttribute('readonly', 'true');
                    if (isReadOnly) {
                        oblInput.style.backgroundColor = '#f5f5f5';
                        oblInput.style.cursor = 'not-allowed';
                    }
                    else {
                        oblInput.style.backgroundColor = '';
                        oblInput.style.cursor = '';
                    }
                }
                if (isReadOnly) {
                    toast("info", '🔒 This project is read-only.');
                }
                else {
                    // toast("info", '🔓 Project is now editable.');
                }
            };
            const propagateCollectionValues = (forecastData) => {
                const updatedData = Object.assign({}, forecastData);
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
            const propagateYearlyValues = (forecastData, isNewMonth = false) => {
                if (!isNewMonth) {
                    return forecastData;
                }
                const updatedData = Object.assign({}, forecastData);
                if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) {
                    return updatedData;
                }
                const yearlyData = updatedData.yearly;
                const updatedYearlyData = {};
                Object.keys(yearlyData).forEach(year => {
                    const yearData = yearlyData[year];
                    if (yearData) {
                        updatedYearlyData[year] = {
                            plan: yearData.actual,
                            actual: 0,
                        };
                    }
                });
                updatedData.yearly = updatedYearlyData;
                return updatedData;
            };
            // =====================================================================
            // ✅ FILTER ENGINE + FILTER POPUP  (same behaviour as the Inventory module)
            //
            //    appliedFilters               = what the TABLE is filtered by
            //    the seven <select> values    = the DRAFT inside the popup
            //      • Apply              → snapshot the draft into appliedFilters and re-filter
            //      • Clear              → clearFilters() (this module's own logic)
            //      • Cancel             → a real cancel: the applied values go back
            //      • ✕ / Esc / backdrop → close but KEEP the draft, so the compact
            //                             "Apply" pill appears (mock-up image 8)
            //    A missing element never breaks the flow (every node is optional).
            // =====================================================================
            const MONTH_NAMES_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            // ---------------------------------------------------------------------
            // Option population. The old loadProjects() only filled #filterProjectSeibies,
            // so the other six selects stayed empty (and the Excel export, which requires
            // a year + month, could never be used). This fills them all.
            // ---------------------------------------------------------------------
            const populateFilterOptions = () => {
                const yearSet = new Set();
                yearSet.add(new Date().getFullYear().toString());
                forecasts.forEach(f => {
                    const data = safeJsonParse(f.ForecastData, {});
                    Object.keys(data).forEach(key => {
                        if (key && /^\d{4}-\d{2}$/.test(key))
                            yearSet.add(key.substring(0, 4));
                    });
                    if (data.yearly && typeof data.yearly === "object") {
                        Object.keys(data.yearly).forEach(year => yearSet.add(year));
                    }
                });
                const sortedYears = Array.from(yearSet).sort().reverse();
                const currentYear = filterYear.value;
                filterYear.innerHTML = '<option value="">All Years</option>' +
                    sortedYears.map(y => `<option value="${y}">${y}</option>`).join("");
                if (currentYear && yearSet.has(currentYear))
                    filterYear.value = currentYear;
                const currentMonth = filterMonth.value;
                filterMonth.innerHTML = '<option value="">All Months</option>' +
                    MONTH_NAMES_SHORT.map((m, idx) => `<option value="${idx + 1}">${m}</option>`).join("");
                if (currentMonth)
                    filterMonth.value = currentMonth;
                const currentCustomer = filterCustomer.value;
                filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                    [...new Set(projects.map(p => p.CustomerName).filter(Boolean))]
                        .sort().map(c => `<option value="${c}">${c}</option>`).join("");
                if (currentCustomer)
                    filterCustomer.value = currentCustomer;
                const currentManager = filterProjectManager.value;
                filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                    [...new Set(projects.map(p => displayNameOf(p.ProjectManager)).filter(Boolean))]
                        .sort().map(m => `<option value="${m}">${m}</option>`).join("");
                if (currentManager)
                    filterProjectManager.value = currentManager;
                const currentGroupHead = filterGroupHead.value;
                filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
                    [...new Set(projects.map(p => displayNameOf(p.GroupHead1)).filter(Boolean))]
                        .sort().map(g => `<option value="${g}">${g}</option>`).join("");
                if (currentGroupHead)
                    filterGroupHead.value = currentGroupHead;
                const currentReference = filterProjectReference.value;
                filterProjectReference.innerHTML = '<option value="">All References</option>' +
                    [...new Set(projects.map(p => p.ProjectReference).filter(Boolean))]
                        .sort().map(r => `<option value="${r}">${r}</option>`).join("");
                if (currentReference)
                    filterProjectReference.value = currentReference;
            };
            // ---------------------------------------------------------------------
            // Filtering (replaces "the selects did nothing" with a real filter)
            // ---------------------------------------------------------------------
            const getProjectByForecast = (forecast) => {
                var _a;
                const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                if (!projectTitle)
                    return undefined;
                return projects.find(p => p.Title === projectTitle);
            };
            const displayNameOf = (value) => {
                if (!value)
                    return "";
                if (typeof value === "object" && value !== null)
                    return value.Title || "";
                return String(value);
            };
            const forecastHasMonth = (data, year, month) => {
                const monthNum = parseInt(month, 10);
                return Object.keys(data || {}).some(key => {
                    if (!key || !/^\d{4}-\d{2}$/.test(key))
                        return false;
                    const keyYear = key.substring(0, 4);
                    if (year && keyYear !== year)
                        return false;
                    if (month) {
                        const keyMonth = parseInt(key.substring(5, 7), 10);
                        if (keyMonth !== monthNum)
                            return false;
                    }
                    return true;
                });
            };
            const forecastHasYear = (data, year) => {
                if (!year)
                    return false;
                const hasMonthly = Object.keys(data || {}).some(key => key && /^\d{4}-\d{2}$/.test(key) && key.substring(0, 4) === year);
                const yearlyData = (data === null || data === void 0 ? void 0 : data.yearly) || {};
                const hasYearly = typeof yearlyData === "object" && Object.prototype.hasOwnProperty.call(yearlyData, year);
                return hasMonthly || hasYearly;
            };
            const filterTable = () => {
                const selectedYear = filterYear.value;
                const selectedMonth = filterMonth.value;
                const selectedProject = filterProject.value;
                const selectedCustomer = filterCustomer.value;
                const selectedManager = filterProjectManager.value;
                const selectedGroupHead = filterGroupHead.value;
                const selectedReference = filterProjectReference.value;
                const noFilters = !selectedYear && !selectedMonth && !selectedProject &&
                    !selectedCustomer && !selectedManager && !selectedGroupHead && !selectedReference;
                if (noFilters) {
                    initializePagination(forecasts);
                    return;
                }
                const filtered = forecasts.filter(f => {
                    var _a;
                    if (selectedProject && ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Title) !== selectedProject)
                        return false;
                    const proj = getProjectByForecast(f);
                    if (proj) {
                        if (selectedCustomer && (proj.CustomerName || "") !== selectedCustomer)
                            return false;
                        if (selectedManager && displayNameOf(proj.ProjectManager) !== selectedManager)
                            return false;
                        if (selectedGroupHead && displayNameOf(proj.GroupHead1) !== selectedGroupHead)
                            return false;
                        if (selectedReference && (proj.ProjectReference || "") !== selectedReference)
                            return false;
                    }
                    else if (selectedCustomer || selectedManager || selectedGroupHead || selectedReference) {
                        return false;
                    }
                    if (selectedYear || selectedMonth) {
                        const data = safeJsonParse(f.ForecastData, {});
                        if (selectedYear && selectedMonth) {
                            if (!forecastHasMonth(data, selectedYear, selectedMonth))
                                return false;
                        }
                        else if (selectedYear) {
                            if (!forecastHasYear(data, selectedYear))
                                return false;
                        }
                        else if (selectedMonth) {
                            if (!forecastHasMonth(data, "", selectedMonth))
                                return false;
                        }
                    }
                    return true;
                });
                if (filtered.length === 0) {
                    // No results: clear the totals band and show the empty state
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    initializePagination([]);
                    return;
                }
                initializePagination(filtered);
            };
            // ---------------------------------------------------------------------
            // Popup / toolbar state
            // ---------------------------------------------------------------------
            const filterTriggerBtn = document.querySelector("#filterTriggerSeibies");
            const filterOverlay = document.querySelector("#seibiesFilterOverlay");
            const filterModalCloseBtn = document.querySelector("#filterModalCloseSeibies");
            const filterCancelBtn = document.querySelector("#filterCancelSeibies");
            const filterClearModalBtn = document.querySelector("#filterClearModalSeibies");
            const filterApplyModalBtn = document.querySelector("#filterApplyModalSeibies");
            const floatingApplyBtn = document.querySelector("#floatingFilterApplySeibies");
            const filterCountBadge = document.querySelector("#filterCountSeibies");
            const filterSummaryBox = document.querySelector("#filterSummarySeibies");
            let appliedFilters = {};
            const FILTER_FIELDS = ['project', 'customer', 'manager', 'groupHead', 'reference', 'year', 'month'];
            const FILTER_SELECTS = {
                project: filterProject,
                customer: filterCustomer,
                manager: filterProjectManager,
                groupHead: filterGroupHead,
                reference: filterProjectReference,
                year: filterYear,
                month: filterMonth
            };
            const FILTER_LABELS = {
                project: 'Project',
                customer: 'Customer',
                manager: 'Manager',
                groupHead: 'Group Head',
                reference: 'Reference',
                year: 'Year',
                month: 'Month'
            };
            const isFilterModalOpen = () => !!filterOverlay && filterOverlay.classList.contains('is-open');
            const readDraftFilters = () => {
                const out = {};
                FILTER_FIELDS.forEach(field => {
                    const el = FILTER_SELECTS[field];
                    if (el && el.value)
                        out[field] = el.value;
                });
                return out;
            };
            const sameFilters = (a, b) => FILTER_FIELDS.every(field => (a[field] || '') === (b[field] || ''));
            const countFilters = (source) => FILTER_FIELDS.filter(field => !!source[field]).length;
            const syncFilterFieldStyles = () => {
                FILTER_FIELDS.forEach(field => {
                    const el = FILTER_SELECTS[field];
                    if (el)
                        el.classList.toggle('is-set', !!el.value);
                });
            };
            // Apply/Clear enable only when there is something to do; the floating Apply
            // pill shows only while there are unapplied changes and the popup is closed
            const syncFilterButtons = () => {
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
            const syncFilterTrigger = () => {
                const n = countFilters(appliedFilters);
                if (filterCountBadge)
                    filterCountBadge.textContent = String(n);
                if (filterTriggerBtn)
                    filterTriggerBtn.classList.toggle('is-active', n > 0);
            };
            // chips above the table: one per applied filter, each removable
            const renderFilterChips = () => {
                if (!filterSummaryBox)
                    return;
                const activeFields = FILTER_FIELDS.filter(field => !!appliedFilters[field]);
                if (activeFields.length === 0) {
                    filterSummaryBox.style.display = 'none';
                    filterSummaryBox.innerHTML = '';
                    return;
                }
                filterSummaryBox.style.display = 'flex';
                filterSummaryBox.innerHTML =
                    '<span class="seibies-filter-summary__label">Active filters:</span>' +
                        activeFields.map(field => {
                            const text = field === 'month'
                                ? `Month: ${MONTH_NAMES_SHORT[parseInt(appliedFilters[field], 10) - 1] || appliedFilters[field]}`
                                : `${FILTER_LABELS[field]}: ${appliedFilters[field]}`;
                            return `<span class="seibies-filter-chip">${text}<button type="button" data-clear-filter="${field}" title="Remove this filter">&times;</button></span>`;
                        }).join('') +
                        '<button type="button" class="seibies-filter-chip__all" id="filterChipsClearAllSeibies">Clear all</button>';
                filterSummaryBox.querySelectorAll('[data-clear-filter]').forEach(btn => {
                    btn.addEventListener('click', () => {
                        const field = btn.dataset.clearFilter;
                        const el = FILTER_SELECTS[field];
                        if (el)
                            el.value = '';
                        delete appliedFilters[field];
                        syncFilterFieldStyles();
                        syncFilterTrigger();
                        renderFilterChips();
                        syncFilterButtons();
                        filterTable();
                    });
                });
                const clearAllBtn = filterSummaryBox.querySelector('#filterChipsClearAllSeibies');
                if (clearAllBtn) {
                    clearAllBtn.addEventListener('click', () => {
                        clearFilters(); // this module's own logic (selects + sort + table)
                        syncFilterUIFromSelects();
                    });
                }
            };
            // after a Clear / Reset the selects are the source of truth again
            const syncFilterUIFromSelects = () => {
                appliedFilters = readDraftFilters();
                syncFilterFieldStyles();
                syncFilterTrigger();
                renderFilterChips();
                syncFilterButtons();
            };
            const openFilterModal = () => {
                if (!filterOverlay)
                    return;
                // the popup always opens showing what is currently applied
                FILTER_FIELDS.forEach(field => {
                    const el = FILTER_SELECTS[field];
                    if (el)
                        el.value = appliedFilters[field] || '';
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
            const closeFilterModal = (restoreApplied = true) => {
                if (!filterOverlay)
                    return;
                if (restoreApplied) {
                    // real cancel: discard the draft, show the applied values again
                    FILTER_FIELDS.forEach(field => {
                        const el = FILTER_SELECTS[field];
                        if (el)
                            el.value = appliedFilters[field] || '';
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
            const applyFilterModal = () => {
                appliedFilters = readDraftFilters();
                syncFilterTrigger();
                renderFilterChips();
                syncFilterButtons();
                filterTable(); // real filtering + totals + pagination
                closeFilterModal(false); // keep the values that were just applied
                const n = countFilters(appliedFilters);
                toast("info", n > 0 ? `${n} SEI-BIES filter(s) applied` : "SEI-BIES filters cleared");
            };
            const clearFilterDraft = () => {
                clearFilters(); // this module's own clear logic, popup stays open
                syncFilterUIFromSelects();
            };
            // ========== EVENT LISTENERS ==========
            showFormBtn === null || showFormBtn === void 0 ? void 0 : showFormBtn.addEventListener("click", () => {
                editingId = null;
                resetForm();
                showForm(false);
            });
            // ============================================================
            // ✅ BACK — SweetAlert confirmation, same look as procurement.ts
            //    • Only one prompt: Leave (Yes, Go Back / Stay)
            //    • Only shown when the form is actually open
            //    • z-index bumped so the loader cannot cover the dialog
            // ============================================================
            const backBtns = document.querySelectorAll(".back-to-table-seibies");
            if (backBtns && backBtns.length > 0) {
                backBtns.forEach(btn => {
                    btn.addEventListener("click", async (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        if (isLeavingSeibiesForm) {
                            console.log("⛔ back blocked: already leaving");
                            return;
                        }
                        // Nothing to confirm when the form is already hidden
                        const formVisible = formSection && formSection.style.display !== "none";
                        if (!formVisible) {
                            editingId = null;
                            hideForm();
                            // reloadInventoryData();
                            return;
                        }
                        isLeavingSeibiesForm = true;
                        try {
                            const confirmResult = await Swal.fire({
                                title: "Leave this page?",
                                html: "Any <strong>unsaved changes</strong> will be lost.<br>" +
                                    "Do you want to go back to the SEI-BIES list?",
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
                            // reloadInventoryData();
                        }
                        catch (err) {
                            console.error("Error while leaving the SEI-BIES form:", err);
                            const message = err instanceof Error ? err.message : String(err);
                            toast("error", `Unexpected error: ${message}`);
                        }
                        finally {
                            isLeavingSeibiesForm = false;
                        }
                    });
                });
            }
            backBtn.addEventListener("click", async (e) => {
                e.preventDefault();
                e.stopPropagation();
                if (isLeavingSeibiesForm) {
                    console.log("⛔ back blocked: already leaving");
                    return;
                }
                // Nothing to confirm when the form is already hidden
                const formVisible = formSection && formSection.style.display !== "none";
                if (!formVisible) {
                    hideForm();
                    resetForm();
                    editingId = null;
                    loadForecasts();
                    return;
                }
                isLeavingSeibiesForm = true;
                try {
                    const confirmResult = await Swal.fire({
                        title: "Leave this page?",
                        html: "Any <strong>unsaved changes</strong> will be lost.<br>" +
                            "Do you want to go back to the SEI-BIES list?",
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
                    editingId = null;
                    loadForecasts();
                }
                catch (err) {
                    console.error("Error while leaving the SEIBIES form:", err);
                    const message = err instanceof Error ? err.message : String(err);
                    toast("error", `Unexpected error: ${message}`);
                }
                finally {
                    isLeavingSeibiesForm = false;
                }
            });
            projectSelect.addEventListener("change", (e) => {
                const projectId = parseInt(e.target.value);
                if (projectId) {
                    onProjectChange(projectId);
                }
                else {
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
            if (filterModalCloseBtn)
                filterModalCloseBtn.addEventListener("click", () => closeFilterModal(false));
            if (filterCancelBtn)
                filterCancelBtn.addEventListener("click", () => closeFilterModal());
            if (filterApplyModalBtn)
                filterApplyModalBtn.addEventListener("click", () => { if (!filterApplyModalBtn.disabled)
                    applyFilterModal(); });
            if (filterClearModalBtn)
                filterClearModalBtn.addEventListener("click", () => { if (!filterClearModalBtn.disabled)
                    clearFilterDraft(); });
            if (floatingApplyBtn)
                floatingApplyBtn.addEventListener("click", applyFilterModal);
            if (filterOverlay) {
                filterOverlay.addEventListener("mousedown", (e) => {
                    if (e.target === filterOverlay)
                        closeFilterModal(false); // keep the draft
                });
            }
            document.addEventListener("keydown", (e) => {
                // Esc = light dismiss: keep the draft, the floating Apply pill stays available
                if (e.key === "Escape" && isFilterModalOpen())
                    closeFilterModal(false);
            });
            exportBtn === null || exportBtn === void 0 ? void 0 : exportBtn.addEventListener("click", exportToExcelWithFullStyle);
            // ========== INITIAL LOADS ==========
            showRevenueLoader();
            window.__revenueLoaderDeadline = Date.now() + 30000;
            window.setTimeout(() => {
                const loader = document.querySelector('#revenuePageLoader');
                if (loader && Date.now() > Number(window.__revenueLoaderDeadline || 0)) {
                    console.warn('⏱ Loader deadline exceeded - forcing hide');
                    hideRevenueLoader();
                }
            }, 31000);
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
                // ✅ fill all seven filter dropdowns + paint the toolbar state (badge "0",
                //    no chips, Apply disabled)
                populateFilterOptions();
                syncFilterUIFromSelects();
                initializePagination(forecasts);
                updateUIBasedOnPermissions({
                    exportBtn: ".export-btn-seibies",
                    addBtn: ".show-form-btn-seibies",
                    saveBtn: '.btn-primary-seibies[type="submit"]'
                });
            }
            catch (error) {
                console.error('❌ Error initializing SEIBIES module:', error);
                toast("error", "Failed to initialize application");
            }
            finally {
                hideRevenueLoader();
            }
            // console.log("✅ SEIBIES Forecast Module Initialized - Default page size: 25 per page");
        }, 100);
    }
};
//# sourceMappingURL=seibies.js.map
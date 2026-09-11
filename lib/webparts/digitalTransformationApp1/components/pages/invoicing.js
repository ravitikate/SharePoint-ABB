import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/invoicing.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canReadAllProjects, isPMManager, isEngineeringUser, canViewForm, getAssignedProjects, getAssignedProjectValues, updateUIBasedOnPermissions, getDigest as getAuthDigest,
// getListEntityType
 } from "./auth";
let currentSortConfig = {
    column: '',
    direction: 'asc'
};
export default {
    content: `
<div class="invoicing-page fade-in">
  <section class="table-section invoicing-card">
    <div class="invoicing-page-header">
      <h1>Invoicing Plan+</h1>
    </div>
    
    <div class="invoicing-table-actions">
          <div class="invoicing-table-actions-row">
        <div class="invoicing-tableColp4">
          <label for="filterProject">Project</label>
          <select id="filterProject"><option value="">All Projects</option></select>
        </div>
        <div class="invoicing-tableColp4">
          <label for="filterCustomer">Customer</label>
          <select id="filterCustomer"><option value="">All Customers</option></select>
        </div>
        <div class="invoicing-tableColp4">
          <label for="filterProjectManager">Project Manager</label>
          <select id="filterProjectManager"><option value="">All Managers</option></select>
        </div>
        <div class="invoicing-tableColp4">
          <label for="filterGroupHead">Group Head</label>
          <select id="filterGroupHead"><option value="">All Group Heads</option></select>
        </div>
        <div class="invoicing-tableColp4">
          <label for="filterProjectReference">Project Reference</label>
          <select id="filterProjectReference"><option value="">All References</option></select>
        </div>
        <div class="invoicing-tableColp4"> 
          <label> </label>
          <button class="invoicing-btn-secondary invoicing-clear-filters-btn" title="Clear All Filters">Reset Filter</button> 
        </div>
      </div>
      
      <!-- Second Row: Year, Month, Export Excel -->
      <div class="invoicing-table-actions-row">
        <div class="invoicing-tableColp4">
          <label for="filterYear">Year</label>
          <select id="filterYear"><option value="">All Years</option></select>
        </div>
        <div class="invoicing-tableColp4">
          <label for="filterMonth">Month</label>
          <select id="filterMonth"><option value="">All Months</option></select>
        </div>
        <div class="invoicing-tableColp4"> 
          <label> </label>
          <button class="invoicing-btn-secondary invoicing-export-btn">📊 Download Excel</button>
        </div>
        <div class="invoicing-tableColp4"> 
    <label> </label>
    <button class="invoicing-btn-secondary invoicing-export-details-btn">📋 Details Download Excel</button>
  </div>
      </div>
    </div>
       
    <!-- ✅ CHANGE 6: same scroll structure as inventory (vertical + horizontal, sticky totals/header) -->
    <div class="table-container-inv">
      <div class="invoicing-table-scroll-wrapper">
        <table class="month-columns-table-inv">
          <thead id="totalsHeader" class="sticky-totals-header"></thead>
          <thead id="monthColumnsHeader" class="sticky-main-header"></thead>
          <tbody class="invoicing-table-body-inv">
            <!-- ✅ CHANGE 1: Static loading row - visible from first paint until the JS render
                 replaces it, so the table never looks blank while data loads -->
            <tr class="invoicing-table-loading-row">
              <td colspan="25">
                <span class="invoicing-table-loading-spinner"></span>Loading Invoicing data…
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      
      <!-- ✅ CHANGE 2: DCS / Procurement-style pagination bar (same look as the other modules):
           LEFT  = « ‹ [Page X of Y pill] [page tiles] › »
           RIGHT = Show [size] per page + "Showing X-Y of N records" badge.
           All ids unchanged - the existing pagination logic works as-is. -->
      <div class="invoicing-pagination-container">
        <div class="invoicing-pagination-row">
          <div class="invoicing-pagination-left">
            <button class="invoicing-pagination-btn invoicing-pg-icon" id="firstPage" title="First page">«</button>
            <button class="invoicing-pagination-btn invoicing-pg-icon" id="prevPage" title="Previous page">‹</button>

            <div class="invoicing-pagination-info invoicing-pg-info">
              Page <span id="currentPage">1</span> of <span id="totalPages">1</span>
            </div>

            <div class="invoicing-pagination-pages" id="pageNumbers"></div>

            <button class="invoicing-pagination-btn invoicing-pg-icon" id="nextPage" title="Next page">›</button>
            <button class="invoicing-pagination-btn invoicing-pg-icon" id="lastPage" title="Last page">»</button>
          </div>
        </div>

        <div class="invoicing-pagination-row invoicing-pagination-bottom-row">
          <div class="invoicing-pagination-right">
            <div class="invoicing-page-size-selector invoicing-pg-size">
              <label for="pageSizeSelect">Show:</label>
              <select id="pageSizeSelect" class="invoicing-page-size-select">
                <option value="10">10</option>
                <option value="25" selected>25</option>
                <option value="50">50</option>
                <option value="100">100</option>
              </select>
              <span>per page</span>
            </div>

            <div class="invoicing-pagination-info invoicing-pg-total">
              Showing <span class="invoicing-pg-range"><span id="startItem">0</span>-<span id="endItem">0</span></span>
              of <span class="invoicing-pg-count" id="totalItems">0</span> records
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>

 <section class="invoicing-form-section invoicing-card" style="display:none;">
   <form id="invoicingForm">
  <div class="invoicing-page-header">
    <h1 id="formTitle">Edit Forecast</h1>
    <!-- ✅ TOP SAVE BUTTON - same type="submit" + .invoicing-btn-primary
         as the bottom button, so both trigger the same submit handler
         and stay in sync (visibility, saving state). The header is
         position:sticky, so this button stays visible while scrolling
         through the long forecast form. -->
    <button type="submit"
            form="invoicingForm"
            class="invoicing-btn-primary invoicing-btn-save-top"
            title="Save Invoicing Forecast">💾 Save Invoicing Forecast</button>
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
                  <label>Invoicing </label>
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
        <button type="submit" class="invoicing-btn-primary">💾 Save Forecast</button>
        <button type="button" class="invoicing-btn-secondary back-to-table">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,
    script: () => {
        setTimeout(async () => {
            const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/DigitalTransformationTest";
            const listName = "InvoicingPlanning";
            const projectListName = "Project List";
            // DOM Elements
            const tableBody = document.querySelector(".invoicing-table-body-inv");
            const monthColumnsHeader = document.querySelector("#monthColumnsHeader");
            const formSection = document.querySelector(".invoicing-form-section");
            const tableSection = document.querySelector(".table-section");
            const projectSelect = document.querySelector("#projectSelect");
            const historicalContainer = document.querySelector("#historicalContainer");
            const historicalToggleBtn = document.querySelector("#historicalToggleBtn");
            const firstRowContainer = document.querySelector("#firstRowContainer");
            const remainingContainer = document.querySelector("#remainingContainer");
            const yearlyContainer = document.querySelector("#yearlyContainer");
            const backBtn = document.querySelector(".back-to-table");
            const form = document.querySelector("#invoicingForm");
            const exportBtn = document.querySelector(".invoicing-export-btn");
            const oblLastMonthInput = document.querySelector("#oblLastMonthInvoicing");
            const totalCurrentValueSpan = document.querySelector("#totalCurrentValueInvoicing");
            const remainingValueSpan = document.querySelector("#remainingValueInvoicing");
            const oblLineFill = document.querySelector("#oblLineFillInvoicing");
            const oblValidationMessage = document.querySelector("#oblValidationMessageInvoicing");
            const formTitle = document.querySelector("#formTitle");
            // Project Details Display Elements
            const projectNameDisplay = document.querySelector("#projectNameDisplayInvoicing");
            const projectRefDisplay = document.querySelector("#projectRefInvoicing");
            const profitCenterDisplay = document.querySelector("#profitCenterInvoicing");
            const projectManagerDisplay = document.querySelector("#projectManagerInvoicing");
            const groupHeadDisplay = document.querySelector("#groupHeadInvoicing");
            const customerDisplay = document.querySelector("#customerInvoicing");
            // Filter Elements
            const filterYear = document.querySelector("#filterYear");
            const filterMonth = document.querySelector("#filterMonth");
            const filterProject = document.querySelector("#filterProject");
            const filterCustomer = document.querySelector("#filterCustomer");
            const filterProjectManager = document.querySelector("#filterProjectManager");
            const filterGroupHead = document.querySelector("#filterGroupHead");
            const filterProjectReference = document.querySelector("#filterProjectReference");
            // Action Buttons
            const clearFiltersBtn = document.querySelector(".invoicing-clear-filters-btn");
            const exportDetailsBtn = document.querySelector(".invoicing-export-details-btn");
            // Pagination Elements
            const firstPageBtn = document.querySelector("#firstPage");
            const prevPageBtn = document.querySelector("#prevPage");
            const nextPageBtn = document.querySelector("#nextPage");
            const lastPageBtn = document.querySelector("#lastPage");
            const currentPageSpan = document.querySelector("#currentPage");
            const totalPagesSpan = document.querySelector("#totalPages");
            const pageNumbersDiv = document.querySelector("#pageNumbers");
            const startItemSpan = document.querySelector("#startItem");
            const endItemSpan = document.querySelector("#endItem");
            const totalItemsSpan = document.querySelector("#totalItems");
            let projects = [];
            let forecasts = [];
            let editingId = null;
            let allMonths = [];
            let currentPage = 1;
            let pageSize = 25; // ✅ default 25 per page (same as DCS / Procurement / SEIBIES / Inventory)
            let totalPages = 1;
            let filteredData = [];
            // Add page size selector element
            const pageSizeSelect = document.querySelector("#pageSizeSelect");
            // Add changePageSize function
            const changePageSize = (newSize) => {
                if (newSize === pageSize)
                    return;
                pageSize = newSize;
                // Calculate new page number to keep the same first item visible
                const oldStartItem = (currentPage - 1) * pageSize;
                currentPage = Math.floor(oldStartItem / pageSize) + 1;
                calculateTotalPages();
                // Adjust current page if it exceeds new total pages
                if (currentPage > totalPages) {
                    currentPage = totalPages || 1;
                }
                updatePaginationControls();
                renderCurrentPage();
                toast("info", `Showing ${pageSize} records per page`);
            };
            // Add event listener for page size selector
            if (pageSizeSelect) {
                pageSizeSelect.addEventListener("change", (e) => {
                    const newSize = parseInt(e.target.value, 10) || 25;
                    changePageSize(newSize);
                });
            }
            // ========== UTILITY FUNCTIONS ==========
            const toast = (icon, title) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
            const createLoadingOverlay = (message = "Processing...") => {
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
            .loading-overlay{position:absolute;top:0;left:0;width:100%;height:100%;background:rgba(255,255,255,0.94);display:flex;justify-content:center;align-items:center;z-index:99999;backdrop-filter:blur(3px);border-radius:10px;overflow:hidden;box-sizing:border-box}
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
            const removeLoadingOverlay = () => {
                document.querySelectorAll(".loading-overlay").forEach(el => el.remove());
            };
            /** mount the loader on the VISIBLE card (table section or form section) —
                the table card while data loads, the form card while saving */
            const loaderHost = () => {
                const sections = Array.from(document.querySelectorAll('.table-section, .invoicing-form-section'));
                const visible = sections.find(s => s.offsetParent !== null);
                return visible || document.querySelector('.table-section') || document.body;
            };
            const getDigest = async () => {
                return getAuthDigest(siteUrl);
            };
            // Safe JSON parsing function
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
                // Calculate total from monthly data only (exclude yearly)
                Object.keys(forecastData).forEach(key => {
                    if (key.match(/^\d{4}-\d{2}$/)) {
                        const monthData = forecastData[key];
                        if (monthData) {
                            totalValue += monthData.actual;
                        }
                    }
                });
                return `${project.Title} - ${currentYear} Forecast - ${totalValue.toLocaleString()}`;
            };
            // Get current and next 2 months for display
            const getDisplayMonths = () => {
                const now = new Date();
                const months = [];
                // Current month + next 2 months
                for (let i = 0; i < 3; i++) {
                    const month = new Date(now.getFullYear(), now.getMonth() + i, 1);
                    months.push(formatMonthKey(month));
                }
                return months;
            };
            const getAllMonthsForStorage = () => {
                return generateIdealMonthStructure();
            };
            // Utility function to allow only numbers 0-9 in text inputs
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
                        this.className = this.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
                    }
                    else {
                        this.className = this.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
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
            // Function to get numeric value from text input - 0 is placeholder
            const getNumericValue = (input) => {
                const value = input.value.trim();
                if (value === '' || value === '0')
                    return 0;
                return parseInt(value) || 0;
            };
            // Function to set numeric value to text input - 0 shows as empty
            const setNumericValue = (input, value) => {
                if (value === 0) {
                    input.value = '';
                    input.className = input.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
                }
                else {
                    input.value = value.toString();
                    input.className = input.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
                }
            };
            // Enhanced OBL validation with toast notifications
            const showOBLErrorToast = (message, isWarning = false) => {
                var _a;
                const existingToast = document.querySelector('.invoicing-obl-error-toast');
                if (existingToast)
                    existingToast.remove();
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
                (_a = toastEl.querySelector('.invoicing-toast-close')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
                    toastEl.remove();
                });
            };
            // ========== MONTH STRUCTURE FUNCTIONS ==========
            const createEmptyItems = () => {
                const items = {};
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
            const generateIdealMonthStructure = () => {
                const months = [];
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
            const getLastMonthKey = (existingData) => {
                const monthKeys = Object.keys(existingData).filter(key => key.match(/^\d{4}-\d{2}$/));
                if (monthKeys.length === 0) {
                    return undefined;
                }
                // Sort months and get the last one
                const sortedMonths = monthKeys.sort();
                const lastMonth = sortedMonths[sortedMonths.length - 1];
                return lastMonth;
            };
            const createFreshYearlyStructure = (lastMonthKey) => {
                const yearlyData = {};
                // Determine start year dynamically
                let startYear;
                if (lastMonthKey) {
                    // Extract year from last month key (format: "YYYY-MM")
                    startYear = parseInt(lastMonthKey.split('-')[0]);
                }
                else {
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
            const createFreshForecastStructure = () => {
                const idealMonths = generateIdealMonthStructure();
                const freshData = {};
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
                // Check if progression is needed
                if (originalMonths[0] === idealMonths[0]) {
                    return forecastData;
                }
                const updatedData = {};
                // STEP 1: Create new structure with ideal months
                idealMonths.forEach((monthKey, index) => {
                    // Check if this month exists in original data
                    const existingData = forecastData[monthKey];
                    if (existingData && isForecastMonthData(existingData)) {
                        // Month exists - copy as-is
                        updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
                    }
                    else {
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
                            monthData.plan = oldActual; // Carry forward actual as plan
                            monthData.actual = 0; // Reset actual to 0
                            // Apply same logic to items for detailed months
                            if (monthData.items && Object.keys(monthData.items).length > 0) {
                                Object.keys(monthData.items).forEach(itemKey => {
                                    const item = monthData.items[itemKey];
                                    const oldItemActual = item.actual;
                                    item.plan = oldItemActual; // Carry forward item actual as plan
                                    item.actual = 0; // Reset item actual to 0
                                });
                            }
                        }
                        else {
                            // HISTORICAL month - keep original values
                        }
                    }
                });
                // Preserve yearly data if it exists
                if (forecastData.yearly) {
                    updatedData.yearly = JSON.parse(JSON.stringify(forecastData.yearly));
                }
                else {
                    // Create fresh yearly data based on new last month
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
                var _a, _b;
                try {
                    // ✅ Include GroupHead1 with expand
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items?$select=Id,Title,CustomerName,GroupHead1/Title,GroupHead1Id,PC,ProjectReference,ProjectManager/Title,ProjectManagerId&$expand=GroupHead1,ProjectManager`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    if (!res.ok) {
                        // Fallback without expand
                        console.warn('⚠️ Expand query failed, trying fallback...');
                        const fallbackRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items?$select=Id,Title,CustomerName,GroupHead1Id,PC,ProjectReference,ProjectManagerId`, {
                            headers: { Accept: "application/json;odata=verbose" }
                        });
                        if (!fallbackRes.ok) {
                            throw new Error(`Failed to load projects: ${fallbackRes.status} ${fallbackRes.statusText}`);
                        }
                        const fallbackData = await fallbackRes.json();
                        const rawProjects = ((_a = fallbackData.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                        // Fetch manager details separately
                        const cleanedProjects = await Promise.all(rawProjects.map(async (p) => {
                            let projectManager = "";
                            let groupHead1 = "";
                            // Fetch ProjectManager details
                            if (p.ProjectManagerId) {
                                try {
                                    const pmRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${p.ProjectManagerId})?$select=Id,Title,Email`, {
                                        headers: { Accept: "application/json;odata=verbose" }
                                    });
                                    if (pmRes.ok) {
                                        const pmData = await pmRes.json();
                                        projectManager = pmData.d.Title || "";
                                    }
                                }
                                catch (e) {
                                    console.warn('Could not fetch ProjectManager:', e);
                                }
                            }
                            // Fetch GroupHead1 details
                            if (p.GroupHead1Id) {
                                try {
                                    const ghRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${p.GroupHead1Id})?$select=Id,Title,Email`, {
                                        headers: { Accept: "application/json;odata=verbose" }
                                    });
                                    if (ghRes.ok) {
                                        const ghData = await ghRes.json();
                                        groupHead1 = ghData.d.Title || "";
                                    }
                                }
                                catch (e) {
                                    console.warn('Could not fetch GroupHead1:', e);
                                }
                            }
                            return Object.assign(Object.assign({}, p), { CustomerName: p.CustomerName || "", GroupHead1: groupHead1, GroupHead1Id: p.GroupHead1Id, PC: p.PC || "", ProjectReference: p.ProjectReference || "", ProjectManager: projectManager, ProjectManagerId: p.ProjectManagerId });
                        }));
                        const projectMap = new Map();
                        cleanedProjects.forEach((p) => {
                            projectMap.set(p.Title, p);
                        });
                        projects = Array.from(projectMap.values());
                        console.log(`📋 Total projects from SharePoint (fallback): ${projects.length}`);
                        const assignedProjects = getAssignedProjects(projects);
                        projectSelect.innerHTML = `<option value="">Select Project</option>` +
                            assignedProjects.map((p) => `<option value="${p.Id}">${p.Title}</option>`).join("");
                        filterProject.innerHTML = '<option value="">All Projects</option>' +
                            [...new Set(assignedProjects.map((p) => p.Title))]
                                .sort()
                                .map(project => `<option value="${project}">${project}</option>`)
                                .join("");
                        return;
                    }
                    const data = await res.json();
                    const rawProjects = ((_b = data.d) === null || _b === void 0 ? void 0 : _b.results) || [];
                    const cleanedProjects = rawProjects.map((p) => (Object.assign(Object.assign({}, p), { CustomerName: p.CustomerName || "", GroupHead1: p.GroupHead1 || "", GroupHead1Id: p.GroupHead1Id, PC: p.PC || "", ProjectReference: p.ProjectReference || "", ProjectManager: p.ProjectManager || "", ProjectManagerId: p.ProjectManagerId })));
                    const projectMap = new Map();
                    cleanedProjects.forEach((p) => {
                        projectMap.set(p.Title, p);
                    });
                    projects = Array.from(projectMap.values());
                    // Debug: Log sample project
                    if (projects.length > 0) {
                        console.log('📋 Sample project data:', {
                            Title: projects[0].Title,
                            ProjectManager: projects[0].ProjectManager,
                            ProjectManagerId: projects[0].ProjectManagerId,
                            GroupHead1: projects[0].GroupHead1,
                            GroupHead1Id: projects[0].GroupHead1Id
                        });
                    }
                    const assignedProjects = getAssignedProjects(projects);
                    console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);
                    projectSelect.innerHTML = `<option value="">Select Project</option>` +
                        assignedProjects.map((p) => `<option value="${p.Id}">${p.Title}</option>`).join("");
                    filterProject.innerHTML = '<option value="">All Projects</option>' +
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
                    const loading = createLoadingOverlay("Loading Invoicing data...");
                    loaderHost().appendChild(loading);
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,InvoiceActual,CreatedDate,UpdatedDate&$expand=Project&$top=5000`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const data = await res.json();
                    // ✅ Get assigned projects
                    const assignedProjects = getAssignedProjects(projects);
                    const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                    const allForecasts = (((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || []).map(forecast => (Object.assign(Object.assign({}, forecast), { ForecastData: typeof forecast.ForecastData === 'string' ?
                            forecast.ForecastData :
                            JSON.stringify(forecast.ForecastData || {}), InvoiceActual: typeof forecast.InvoiceActual === 'string' ?
                            forecast.InvoiceActual :
                            JSON.stringify(forecast.InvoiceActual || {}), Freeze: forecast.Freeze || false })));
                    // ✅ Only keep forecasts for assigned projects
                    forecasts = allForecasts.filter(forecast => {
                        var _a;
                        const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        return projectTitle && assignedProjectTitles.has(projectTitle);
                    });
                    console.log(`✅ Loaded ${forecasts.length} invoicing forecasts`);
                }
                catch (error) {
                    console.error("Error loading invoicing forecasts:", error);
                    toast("error", "Failed to load invoicing forecasts");
                }
                finally {
                    removeLoadingOverlay();
                }
            };
            const getGroupHeadDisplayName = (groupHead) => {
                if (!groupHead)
                    return "-";
                if (typeof groupHead === "object" && groupHead !== null) {
                    return groupHead.Title || "-";
                }
                return groupHead;
            };
            // ========== EXPORT FUNCTIONS ==========
            // Add this utility function at the top with other utility functions
            const getMonthName = (monthNumber) => {
                const months = [
                    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
                ];
                return months[parseInt(monthNumber) - 1] || "Month";
            };
            // UPDATED: Main Export Excel - Only Previous and Current columns
            const exportToExcelWithFullStyle = () => {
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
                try {
                    // Generate headers according to your specified structure
                    const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);
                    // Prepare data for Excel
                    const excelData = exportData.map(forecast => {
                        var _a, _b;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        // Parse data from both sources
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        // Get OBL value from database
                        const oblValue = forecast.OBL || 0;
                        // Selected month key
                        const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
                        const monthName = getMonthName(selectedMonth);
                        // Get data for selected month
                        const selectedMonthForecastData = forecastData[selectedMonthKey] || { plan: 0, actual: 0 };
                        const selectedMonthPlan = selectedMonthForecastData.plan || 0;
                        const selectedMonthCurrent = selectedMonthForecastData.actual || 0;
                        // Create base row data - ONLY Previous and Current columns
                        const rowData = {
                            "SO/Project Ref": (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                            "Project Name": projectTitle,
                            "Project Manager": typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                            "Group Head": getGroupHeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.GroupHead1),
                            "Customer Name": (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
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
                            const forecastMonthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                            const planValue = forecastMonthData.plan || 0;
                            const currentValue = forecastMonthData.actual || 0;
                            // Add both previous (plan) and current columns for each month
                            rowData[`${displayLabel} (P)`] = planValue;
                            rowData[`${displayLabel} (C)`] = currentValue;
                        });
                        // Add 5-year data (previous + current)
                        lastFiveYears.forEach(year => {
                            var _a;
                            const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0 };
                            rowData[`${year} (P)`] = yearData.plan;
                            rowData[`${year} (C)`] = yearData.actual;
                        });
                        // Calculate balance
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
                    const numberStyle = Object.assign(Object.assign({}, dataStyle), { alignment: {
                            horizontal: "right",
                            vertical: "center"
                        } });
                    // Apply styles to all cells
                    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                    // Style header row (row 0)
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s) {
                            ws[cellAddress].s = {};
                        }
                        ws[cellAddress].s = headerStyle;
                    }
                    // Style data rows
                    for (let row = range.s.r + 1; row <= range.e.r; row++) {
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                            if (!ws[cellAddress])
                                continue;
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
                    const colWidths = headers.map((header) => {
                        if (header.includes('Project Name'))
                            return { wch: 25 };
                        if (header.includes('Customer Name'))
                            return { wch: 20 };
                        if (header.includes('Profit Centre') || header.includes('SO/Project Ref'))
                            return { wch: 12 };
                        if (header.includes('Group Head') || header.includes('Project Manager'))
                            return { wch: 15 };
                        if (header.includes('OBL') || header.includes('Balance'))
                            return { wch: 12 };
                        if (header.includes('(P)') || header.includes('(C)'))
                            return { wch: 10 };
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
                }
                catch (error) {
                    toast("error", "Failed to export data to Excel");
                }
            };
            // Updated helper function to generate headers - ONLY Previous and Current
            const generateHeadersStructure = (selectedYear, selectedMonth, exportData) => {
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
                const allMonthKeys = new Set();
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
                const monthlyHeaders = [];
                next17Months.forEach(monthKey => {
                    const [year, month] = monthKey.split('-');
                    const monthLabel = getMonthName(month);
                    const displayLabel = `${monthLabel} ${year}`;
                    monthlyHeaders.push(`${displayLabel} (P)`, `${displayLabel} (C)`);
                });
                // Get last 5 years from yearly forecast data
                const lastFiveYears = getLastFiveYears(exportData);
                const yearlyHeaders = [];
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
                // Sort years and get last 5
                const sortedYears = Array.from(years).sort();
                return sortedYears.slice(-5);
            };
            // DETAILED EXPORT - KEPT AS IS (unchanged)
            const exportDetailedItemsToExcel = () => {
                var _a, _b, _c, _d, _e, _f, _g, _h, _j;
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
                try {
                    // Get current and next 2 months for detailed items
                    const currentDate = new Date();
                    const detailedMonths = [];
                    for (let i = 0; i < 3; i++) {
                        const monthDate = new Date(currentDate);
                        monthDate.setMonth(monthDate.getMonth() + i);
                        const monthKey = formatMonthKey(monthDate);
                        if (monthKey) {
                            detailedMonths.push(monthKey);
                        }
                    }
                    // Get next 15 months for monthly totals
                    const nextFifteenMonths = [];
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
                    const currentMonthName = getMonthName(((_a = detailedMonths[0]) === null || _a === void 0 ? void 0 : _a.split('-')[1]) || '') || ((_b = detailedMonths[0]) === null || _b === void 0 ? void 0 : _b.split('-')[1]) || '';
                    const nextMonth1Name = getMonthName(((_c = detailedMonths[1]) === null || _c === void 0 ? void 0 : _c.split('-')[1]) || '') || ((_d = detailedMonths[1]) === null || _d === void 0 ? void 0 : _d.split('-')[1]) || '';
                    const nextMonth2Name = getMonthName(((_e = detailedMonths[2]) === null || _e === void 0 ? void 0 : _e.split('-')[1]) || '') || ((_f = detailedMonths[2]) === null || _f === void 0 ? void 0 : _f.split('-')[1]) || '';
                    const currentMonthYear = ((_g = detailedMonths[0]) === null || _g === void 0 ? void 0 : _g.split('-')[0]) || '';
                    const nextMonth1Year = ((_h = detailedMonths[1]) === null || _h === void 0 ? void 0 : _h.split('-')[0]) || '';
                    const nextMonth2Year = ((_j = detailedMonths[2]) === null || _j === void 0 ? void 0 : _j.split('-')[0]) || '';
                    // Prepare data for Excel with two-row header structure
                    const excelData = [];
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
                        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        // Parse data from both sources
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        // Get OBL value from database
                        const oblValue = forecast.OBL || 0;
                        // Get data for detailed months with safety checks
                        const detailedMonthsData = [];
                        for (let i = 0; i < 3; i++) {
                            const monthKey = detailedMonths[i];
                            if (monthKey) {
                                const monthForecastData = forecastData[monthKey] || { plan: 0, actual: 0, forecast: 0, items: createEmptyItems() };
                                detailedMonthsData.push({
                                    plan: monthForecastData.plan || 0,
                                    current: monthForecastData.actual || 0,
                                    forecast: monthForecastData.forecast || 0,
                                    items: monthForecastData.items || createEmptyItems()
                                });
                            }
                            else {
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
                            const monthData = forecastData[monthKey];
                            return sum + ((monthData === null || monthData === void 0 ? void 0 : monthData.actual) || 0);
                        }, 0);
                        const totalYearlyCurrent = forecastData.yearly ?
                            Object.values(forecastData.yearly).reduce((sum, yearData) => {
                                return sum + (yearData.actual || 0);
                            }, 0) : 0;
                        const balance = oblValue - (totalMonthlyCurrent + totalYearlyCurrent);
                        // Create one row for each item (1-10)
                        for (let itemNum = 1; itemNum <= 10; itemNum++) {
                            const itemKey = itemNum.toString();
                            const rowData = [
                                // Base information (7 columns) - A-G
                                (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                                projectTitle,
                                (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                                forecast.ProfitCentre || "-",
                                getGroupHeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.GroupHead1),
                                typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                                oblValue,
                                // Selected month totals (4 columns) - H-K
                                ((_c = detailedMonthsData[0]) === null || _c === void 0 ? void 0 : _c.plan) || 0,
                                ((_d = detailedMonthsData[0]) === null || _d === void 0 ? void 0 : _d.current) || 0,
                                0, // Removed actual invoicing from detailed export
                                ((_e = detailedMonthsData[0]) === null || _e === void 0 ? void 0 : _e.forecast) || 0,
                            ];
                            // Current month items (5 columns) - L-P
                            const currentMonthItems = ((_f = detailedMonthsData[0]) === null || _f === void 0 ? void 0 : _f.items) || createEmptyItems();
                            const currentMonthItem = currentMonthItems[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
                            rowData.push(currentMonthItem.name || `Item ${itemNum}`, currentMonthItem.plan, currentMonthItem.actual, currentMonthItem.forecast || 0, currentMonthItem.date || "");
                            // Next month 1 totals (2 columns) - Q-R
                            rowData.push(((_g = detailedMonthsData[1]) === null || _g === void 0 ? void 0 : _g.plan) || 0, ((_h = detailedMonthsData[1]) === null || _h === void 0 ? void 0 : _h.current) || 0);
                            // Next month 1 items (3 columns) - S-U
                            const nextMonth1Items = ((_j = detailedMonthsData[1]) === null || _j === void 0 ? void 0 : _j.items) || createEmptyItems();
                            const nextMonth1Item = nextMonth1Items[itemKey] || { name: '', plan: 0, actual: 0 };
                            rowData.push(nextMonth1Item.name || `Item ${itemNum}`, nextMonth1Item.plan, nextMonth1Item.actual);
                            // Next month 2 totals (2 columns) - V-W
                            rowData.push(((_k = detailedMonthsData[2]) === null || _k === void 0 ? void 0 : _k.plan) || 0, ((_l = detailedMonthsData[2]) === null || _l === void 0 ? void 0 : _l.current) || 0);
                            // Next month 2 items (3 columns) - X-Z
                            const nextMonth2Items = ((_m = detailedMonthsData[2]) === null || _m === void 0 ? void 0 : _m.items) || createEmptyItems();
                            const nextMonth2Item = nextMonth2Items[itemKey] || { name: '', plan: 0, actual: 0 };
                            rowData.push(nextMonth2Item.name || `Item ${itemNum}`, nextMonth2Item.plan, nextMonth2Item.actual);
                            // Add next 15 months data (P and C for each month) - starting from column AA
                            nextFifteenMonths.forEach(monthKey => {
                                const monthForecastData = forecastData[monthKey] || { plan: 0, actual: 0 };
                                rowData.push(monthForecastData.plan || 0, monthForecastData.actual || 0);
                            });
                            // Add 5-year data (P and C for each year)
                            lastFiveYears.forEach(year => {
                                var _a;
                                const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0 };
                                rowData.push(yearData.plan, yearData.actual);
                            });
                            // Add Balance - only for first item of each project, empty for others
                            if (itemNum === 1) {
                                rowData.push(balance);
                            }
                            else {
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
                    const numberStyle = Object.assign(Object.assign({}, dataStyle), { alignment: { horizontal: "right", vertical: "center" } });
                    const textStyle = Object.assign(Object.assign({}, dataStyle), { alignment: { horizontal: "left", vertical: "center" } });
                    // Apply styles to all cells
                    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                    // Style first header row (row 0)
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s)
                            ws[cellAddress].s = {};
                        ws[cellAddress].s = mainHeaderStyle;
                    }
                    // Style second header row (row 1)
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 1, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s)
                            ws[cellAddress].s = {};
                        ws[cellAddress].s = subHeaderStyle;
                    }
                    // Style data rows (starting from row 2)
                    for (let row = range.s.r + 2; row <= range.e.r; row++) {
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                            if (!ws[cellAddress])
                                continue;
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
                }
                catch (error) {
                    toast("error", "Failed to export detailed items data to Excel");
                }
            };
            // Helper function to get last 5 years from forecast data
            const getLastFiveYearsFromForecastData = (data) => {
                const years = new Set();
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
            const initializePagination = (data) => {
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
            const calculateTotalPages = () => {
                if (!filteredData || filteredData.length === 0) {
                    totalPages = 1;
                }
                else {
                    totalPages = Math.ceil(filteredData.length / pageSize);
                }
            };
            const updatePaginationControls = () => {
                currentPageSpan.textContent = currentPage.toString();
                totalPagesSpan.textContent = totalPages.toString();
                const totalItems = filteredData ? filteredData.length : 0;
                const startItem = totalItems > 0 ? (currentPage - 1) * pageSize + 1 : 0;
                const endItem = totalItems > 0 ? Math.min(currentPage * pageSize, totalItems) : 0;
                startItemSpan.textContent = startItem.toString();
                endItemSpan.textContent = endItem.toString();
                totalItemsSpan.textContent = totalItems.toString();
                firstPageBtn.disabled = currentPage === 1 || totalItems === 0;
                prevPageBtn.disabled = currentPage === 1 || totalItems === 0;
                nextPageBtn.disabled = currentPage === totalPages || totalItems === 0;
                lastPageBtn.disabled = currentPage === totalPages || totalItems === 0;
                updatePageNumbers();
            };
            const updatePageNumbers = () => {
                pageNumbersDiv.innerHTML = '';
                if (!filteredData || filteredData.length === 0)
                    return;
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
            const goToPage = (page) => {
                if (page < 1 || page > totalPages || page === currentPage)
                    return;
                currentPage = page;
                updatePaginationControls();
                renderCurrentPage();
            };
            const renderCurrentPage = () => {
                // Check if filteredData has data
                if (!filteredData || filteredData.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:20px;">No records found</td></tr>`;
                    const totalsHeader = document.querySelector("#totalsHeader");
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
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
            const measureStickyHeaderOffsets = () => {
                const table = document.querySelector('.month-columns-table-inv');
                // Only meaningful when the table is actually laid out & visible.
                if (!table)
                    return;
                const wrapper = table.closest('.invoicing-table-scroll-wrapper');
                if (wrapper && wrapper.offsetParent === null)
                    return; // hidden by display:none ancestors
                // 1) Clear any previous inline tops so we measure from natural (unstuck) positions.
                const allCells = table.querySelectorAll('#totalsHeader th, #totalsHeader td, #monthColumnsHeader th, #monthColumnsHeader td');
                allCells.forEach(cell => { cell.style.top = ''; });
                // 2) Force layout so getBoundingClientRect returns real heights.
                void table.offsetHeight;
                // 3) Stack all header rows across BOTH theads in order.
                const headerRows = table.querySelectorAll('#totalsHeader tr, #monthColumnsHeader tr');
                let offset = 0;
                headerRows.forEach(row => {
                    const h = row.getBoundingClientRect().height || 0;
                    row.querySelectorAll('th, td').forEach(cell => {
                        cell.style.position = 'sticky';
                        cell.style.top = `${offset}px`;
                    });
                    offset += h;
                });
            };
            // Public wrapper: measures immediately and again on the next frame so heights are
            // correct once fonts/scrollbars have settled.
            const applyStickyHeaderOffsets = () => {
                measureStickyHeaderOffsets();
                if (typeof window !== 'undefined') {
                    requestAnimationFrame(measureStickyHeaderOffsets);
                }
            };
            // ✅ CHANGE 1 + 3: Totals header - same structure, colours and column order as inventory
            // Rows are rendered directly inside <thead id="totalsHeader"> (no nested table) so they align with the main columns.
            const renderTotalsHeader = (displayMonths, currentMonthKey) => {
                const totalsHeader = document.querySelector("#totalsHeader");
                if (!totalsHeader)
                    return;
                totalsHeader.innerHTML = '';
                const allData = forecasts;
                const filtered = filteredData.length > 0 ? filteredData : forecasts;
                const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
                const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);
                const baseColumns = 6;
                const monthColumns = displayMonths.length * 2;
                const totalColspan = baseColumns + monthColumns;
                const buildTotalsRow = (label, totals, css) => {
                    const monthCells = displayMonths.map(m => {
                        const t = totals.monthly[m] || { plan: 0, current: 0 };
                        return `
              <td class="invoicing-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-IN')}</td>
              <td class="invoicing-month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-IN')}</td>
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
              <td class="inv-col" style="text-align:right; vertical-align:middle; min-width:90px; max-width:90px; width:90px; font-weight:bold;">${totals.obl.toLocaleString('en-IN')}</td>
              ${monthCells}
            </tr>
          `;
                };
                let html = `
          <tr class="invoicing-totals-title-row">
            <th colspan="${totalColspan}" style="background:#f8f9fa; padding:8px; text-align:center; font-size:13px; font-weight:bold; color:#495057; border:1px solid #dee2e6;">
              📊 Invoicing Totals
            </th>
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
            const calculateTotals = (data, displayMonths, currentMonthKey) => {
                const totals = {
                    obl: 0,
                    monthly: {}
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
            const renderTableWithMonthColumns = (data) => {
                if (!data.length) {
                    tableBody.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:20px;">No records found</td></tr>`;
                    const totalsHeader = document.querySelector("#totalsHeader");
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
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
                    var _a;
                    const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                    const proj = projects.find(p => p.Title === projectTitle);
                    const forecastData = safeJsonParse(forecast.ForecastData, {});
                    const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-IN');
                    // ✅ Project Manager - Person/Group type
                    let projectManagerDisplay = "-";
                    if (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) {
                        projectManagerDisplay = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
                            ? proj.ProjectManager.Title || ""
                            : proj.ProjectManager || "";
                    }
                    // ✅ GroupHead1 - Person/Group type
                    let groupHeadDisplayName = "-";
                    if (proj === null || proj === void 0 ? void 0 : proj.GroupHead1) {
                        groupHeadDisplayName = typeof proj.GroupHead1 === "object" && proj.GroupHead1 !== null
                            ? proj.GroupHead1.Title || ""
                            : proj.GroupHead1 || "";
                    }
                    // ✅ CHANGE 5: only Project Reference is the clickable edit link (same as inventory).
                    // Project name is plain text. If a project has no reference, fall back to "#Id" so every row is still editable.
                    const projectRefText = (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || `#${forecast.Id}`;
                    const projectRefLink = `<a href="#" class="invoicing-edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;" title="Edit ${projectTitle}">${projectRefText}</a>`;
                    let rowHTML = `
    <tr data-forecast-id="${forecast.Id}">
      <td class="invoicing-project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
      <td class="invoicing-customer-col" style="min-width:130px; max-width:130px; width:130px;">${(proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-"}</td>
      <td class="invoicing-project-col" style="min-width:180px; max-width:180px; width:180px;">${projectTitle}</td>
      <td class="invoicing-gh-col" style="min-width:100px; max-width:100px; width:100px;">${groupHeadDisplayName}</td>
      <td class="invoicing-pm-col" style="min-width:100px; max-width:100px; width:100px;">${projectManagerDisplay}</td>
      <td class="inv-col" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
  `;
                    // Add month columns data (only Previous and Current)
                    displayMonths.forEach((monthKey) => {
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
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
            const attachSortButtonListeners = () => {
                document.querySelectorAll('.invoicing-sort-btn[data-sort]').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const button = e.currentTarget;
                        const sortColumn = button.dataset.sort;
                        const direction = button.dataset.direction;
                        const subColumn = button.dataset.subcolumn;
                        const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;
                        if (sortColumn && direction) {
                            handleInvoicingSort(sortColumn, direction, subColumn, monthIndex);
                        }
                    });
                });
            };
            const attachEditTriggerListeners = () => {
                document.querySelectorAll('.invoicing-edit-trigger').forEach(link => {
                    link.addEventListener('click', (e) => {
                        e.preventDefault();
                        const id = parseInt(e.currentTarget.dataset.id);
                        editForecast(id);
                    });
                });
            };
            const handleInvoicingSort = (column, direction, subColumn, monthIndex) => {
                // Update current sort state
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
                    var _a, _b, _c, _d;
                    let valueA = 0;
                    let valueB = 0;
                    switch (column) {
                        case 'obl':
                            valueA = a.OBL || 0;
                            valueB = b.OBL || 0;
                            break;
                        case 'projectReference':
                            const projRefA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projRefB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            valueA = (projRefA === null || projRefA === void 0 ? void 0 : projRefA.ProjectReference) || "";
                            valueB = (projRefB === null || projRefB === void 0 ? void 0 : projRefB.ProjectReference) || "";
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
                        case 'projectManager':
                            const projMA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projMB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            valueA = typeof (projMA === null || projMA === void 0 ? void 0 : projMA.ProjectManager) === "object" ? (_c = projMA.ProjectManager) === null || _c === void 0 ? void 0 : _c.Title : (projMA === null || projMA === void 0 ? void 0 : projMA.ProjectManager) || "";
                            valueB = typeof (projMB === null || projMB === void 0 ? void 0 : projMB.ProjectManager) === "object" ? (_d = projMB.ProjectManager) === null || _d === void 0 ? void 0 : _d.Title : (projMB === null || projMB === void 0 ? void 0 : projMB.ProjectManager) || "";
                            break;
                        case 'groupHead':
                            const projGA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projGB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            // ✅ Handle GroupHead1 as Person/Group
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
                        case 'customer':
                            const projA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            valueA = (projA === null || projA === void 0 ? void 0 : projA.CustomerName) || "";
                            valueB = (projB === null || projB === void 0 ? void 0 : projB.CustomerName) || "";
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
                else if (column === 'obl') {
                    columnName = 'Invoicing';
                }
                else if (column === 'project') {
                    columnName = 'Project';
                }
                else if (column === 'projectReference') {
                    columnName = 'Project Reference';
                }
                else if (column === 'projectManager') {
                    columnName = 'Project Manager';
                }
                else if (column === 'groupHead') {
                    columnName = 'Group Head';
                }
                else if (column === 'customer') {
                    columnName = 'Customer';
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
            const setupHistoricalToggle = () => {
                const toggleBtn = document.querySelector("#historicalToggleBtn");
                const container = document.querySelector("#historicalContainer");
                const header = document.querySelector(".invoicing-historical-toggle-header");
                if (!toggleBtn || !container || !header)
                    return;
                let isVisible = false;
                const updateHistoricalToggle = () => {
                    const arrow = toggleBtn.querySelector(".invoicing-toggle-arrow");
                    const label = toggleBtn.querySelector(".invoicing-toggle-label");
                    if (isVisible) {
                        container.style.display = "grid";
                        toggleBtn.setAttribute("aria-expanded", "true");
                        toggleBtn.title = "Hide historical data";
                        if (arrow)
                            arrow.textContent = "▲";
                        if (label)
                            label.textContent = "Hide";
                    }
                    else {
                        container.style.display = "none";
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
                // Prevent duplicate handlers without replacing the button/header.
                toggleBtn.onclick = (event) => {
                    toggleHistorical(event);
                };
                header.onclick = (event) => {
                    const target = event.target;
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
            const resetHistoricalToggle = () => {
                historicalContainer.style.display = "none";
                if (historicalToggleBtn) {
                    historicalToggleBtn.setAttribute("aria-expanded", "false");
                    historicalToggleBtn.title = "Show historical data";
                    const arrow = historicalToggleBtn.querySelector(".invoicing-toggle-arrow");
                    const label = historicalToggleBtn.querySelector(".invoicing-toggle-label");
                    if (arrow)
                        arrow.textContent = "▼";
                    if (label)
                        label.textContent = "Show";
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
            const hideForm = () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                formSection.classList.remove('invoicing-edit-mode');
                projectSelect.disabled = false;
                // Reset any form state
                resetForm();
                // ✅ CHANGE 6: table is visible again → recompute sticky offsets
                setTimeout(applyStickyHeaderOffsets, 50);
            };
            const onSaveSuccess = () => {
                // Show success message
                toast("success", editingId ? "Forecast updated successfully!" : "Forecast created successfully!");
                // Reset form state
                resetForm();
                // Show table section
                tableSection.style.display = "block";
                formSection.style.display = "none";
                // Reload data to reflect changes
                loadForecasts().then(() => {
                    initializePagination(forecasts);
                });
            };
            const resetForm = () => {
                editingId = null;
                form.reset();
                projectSelect.value = "";
                // Clear project details display
                if (projectNameDisplay)
                    projectNameDisplay.textContent = "-";
                if (projectRefDisplay)
                    projectRefDisplay.textContent = "-";
                if (profitCenterDisplay)
                    profitCenterDisplay.textContent = "-";
                if (projectManagerDisplay)
                    projectManagerDisplay.textContent = "-";
                if (groupHeadDisplay)
                    groupHeadDisplay.textContent = "-";
                if (customerDisplay)
                    customerDisplay.textContent = "-";
                historicalContainer.innerHTML = "";
                resetHistoricalToggle();
                firstRowContainer.innerHTML = "";
                remainingContainer.innerHTML = "";
                yearlyContainer.innerHTML = "";
                // Hide all forecast sections
                toggleForecastSections(false);
                const oblContainer = document.querySelector('.obl-section-container-invoicing');
                if (oblContainer) {
                    oblContainer.className = 'obl-section-container-invoicing invoicing-hidden-section';
                }
                oblLastMonthInput.removeAttribute('readonly');
                oblLastMonthInput.placeholder = "Enter OBL";
                setNumericValue(oblLastMonthInput, 0);
                const fieldNote = document.querySelector('.field-note');
                if (fieldNote)
                    fieldNote.textContent = "For new projects: Enter initial OBL value";
                oblValidationMessage.textContent = "✅ All values are within INVOICING limit";
                oblValidationMessage.className = 'validation-message-invoicing';
                updateOBLValidationDisplay();
            };
            // ========== VALIDATION FUNCTIONS ==========
            const calculateTotalCurrentValue = () => {
                let total = 0;
                // Include all future current values EXCEPT derived-readonly ones.
                // Detailed months (indices 3-5): top "Current" is READONLY (auto = items sum), so the
                // real editable values live in the item rows. Long-term months (6+): editable single Current.
                const futureMonths = allMonths.slice(3);
                futureMonths.forEach(monthKey => {
                    const idx = allMonths.indexOf(monthKey);
                    if (idx >= 3 && idx <= 5) {
                        document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]:not([readonly])`).forEach(input => {
                            total += getNumericValue(input);
                        });
                    }
                    else {
                        document.querySelectorAll(`input.main-current-input[data-month="${monthKey}"]:not([readonly])`).forEach(input => {
                            total += getNumericValue(input);
                        });
                    }
                });
                // Include yearly actual values
                const yearlyCurrentInputs = document.querySelectorAll(`input.yearly-current-input-inv:not([readonly])`);
                yearlyCurrentInputs.forEach(input => {
                    total += getNumericValue(input);
                });
                return total;
            };
            const validateOBLLimit = (changedInput) => {
                const oblValue = getNumericValue(oblLastMonthInput);
                const totalCurrent = calculateTotalCurrentValue();
                totalCurrentValueSpan.textContent = totalCurrent.toLocaleString();
                const remaining = oblValue - totalCurrent;
                remainingValueSpan.textContent = remaining.toLocaleString();
                if (oblValue === 0) {
                    oblLineFill.style.width = `0%`;
                    oblLineFill.className = 'line-fill-invoicing';
                    if (totalCurrent > 0) {
                        showOBLErrorToast('⚠️ OBL is zero but current values exist. Please set OBL value.', true);
                        oblValidationMessage.textContent = '⚠️ OBL is zero but current values exist. Please set OBL value.';
                        oblValidationMessage.className = 'validation-message-invoicing danger-invoicing';
                        blockAllCurrentInputs();
                        return false;
                    }
                    else {
                        const existingToast = document.querySelector('.invoicing-obl-error-toast');
                        if (existingToast)
                            existingToast.remove();
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
                }
                else if (percentage > 75) {
                    oblLineFill.classList.add('warning-invoicing');
                }
                oblValidationMessage.className = 'validation-message-invoicing';
                if (totalCurrent > oblValue) {
                    const exceedAmount = totalCurrent - oblValue;
                    showOBLErrorToast(`❌ Total current value exceeds INVOICING by ${exceedAmount.toLocaleString()}! Reduce values to continue.`);
                    oblValidationMessage.textContent = `❌ Total current value exceeds INVOICING by ${exceedAmount.toLocaleString()}! Reduce values to continue.`;
                    oblValidationMessage.classList.add('danger-invoicing');
                    blockIncreasingCurrentInputs(changedInput);
                    return false;
                }
                else if (percentage > 90) {
                    showOBLErrorToast('⚠️ INVOICING limit nearly reached! Consider reducing values.', true);
                    oblValidationMessage.textContent = '⚠️ INVOICING limit nearly reached! Consider reducing values.';
                    oblValidationMessage.classList.add('warning-invoicing');
                    unblockAllCurrentInputs();
                    return true;
                }
                else {
                    const existingToast = document.querySelector('.invoicing-obl-error-toast');
                    if (existingToast)
                        existingToast.remove();
                    oblValidationMessage.textContent = '✅ All values are within INVOICING limit';
                    oblValidationMessage.classList.add('success-invoicing');
                    unblockAllCurrentInputs();
                    return true;
                }
            };
            const blockAllCurrentInputs = () => {
                const currentInputs = document.querySelectorAll('input[data-type="actual"]:not([readonly])');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-inv:not([readonly])');
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
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-inv:not([readonly])');
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
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-inv');
                currentInputs.forEach(input => {
                    // Derived Current (auto sum) must NEVER lose readonly.
                    const isDerived = input.classList.contains('derived-month-current') || input.getAttribute('data-derived') === 'current';
                    input.style.backgroundColor = '';
                    input.style.cursor = '';
                    input.removeAttribute('data-block-increase');
                    if (!isDerived) {
                        input.removeAttribute('readonly');
                    }
                    else {
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
            const validateMonthItemTotals = (changedInput) => {
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
            const validateSingleMonthItems = (monthKey, changedInput) => {
                const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                const itemCurrentInputs = document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]`);
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
                    }
                    else {
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
            const unblockAllMonthItemInputs = (monthKey) => {
                const itemInputs = document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]`);
                itemInputs.forEach(input => {
                    input.style.backgroundColor = '';
                    input.style.cursor = '';
                    input.removeAttribute('data-block-increase');
                    input.removeAttribute('readonly');
                });
            };
            const updateOBLValidationDisplay = (changedInput) => {
                validateOBLLimit(changedInput);
                validateMonthItemTotals(changedInput);
            };
            // ========== FORECAST FORM FUNCTIONS ==========
            const onProjectChange = (projectId, existingForecast) => {
                const project = projects.find(p => p.Id === projectId);
                if (!project)
                    return;
                // Show all forecast sections when project is selected
                toggleForecastSections(true);
                renderProjectDetails(project);
                let forecastData = {};
                if (existingForecast) {
                    forecastData = safeJsonParse(existingForecast.ForecastData, {});
                }
                const isNewProject = !existingForecast;
                const oblContainer = document.querySelector('.obl-section-container-invoicing');
                const fieldNote = document.querySelector('.field-note');
                if (isNewProject) {
                    if (oblContainer) {
                        oblContainer.className = 'obl-section-container-invoicing invoicing-visible-section new-project';
                    }
                    oblLastMonthInput.removeAttribute('readonly');
                    oblLastMonthInput.placeholder = "Enter initial OBL for new project";
                    if (fieldNote)
                        fieldNote.textContent = "For new projects: Enter initial OBL value";
                    oblValidationMessage.textContent = "Enter initial OBL value for new project";
                    oblValidationMessage.className = 'validation-message-invoicing';
                    const defaultOBL = project.PC ? 100000 : 0;
                    setNumericValue(oblLastMonthInput, defaultOBL);
                }
                else {
                    if (oblContainer) {
                        oblContainer.className = 'obl-section-container-invoicing invoicing-visible-section existing-project';
                    }
                    oblLastMonthInput.setAttribute('readonly', 'true');
                    const oblValue = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                    setNumericValue(oblLastMonthInput, oblValue);
                    oblLastMonthInput.placeholder = "OBL from database";
                    if (fieldNote)
                        fieldNote.textContent = "OBL value from database record";
                    if (oblValue > 0) {
                        oblLastMonthInput.value = oblValue.toString();
                        oblLastMonthInput.className = oblLastMonthInput.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
                    }
                }
                renderForecastInterface(forecastData, getNumericValue(oblLastMonthInput), existingForecast);
            };
            // Add this function to show/hide forecast sections
            const toggleForecastSections = (show) => {
                const sections = [
                    '.obl-section-container-invoicing',
                    '#historicalForecastSection',
                    '#detailedForecastSection',
                    '#longtermForecastSection',
                    '#yearlyForecastSection'
                ];
                sections.forEach(selector => {
                    const element = document.querySelector(selector);
                    if (element) {
                        if (show) {
                            element.classList.remove('invoicing-hidden-section');
                            element.classList.add('invoicing-visible-section');
                        }
                        else {
                            element.classList.remove('invoicing-visible-section');
                            element.classList.add('invoicing-hidden-section');
                        }
                    }
                });
            };
            // Updated renderProjectDetails to use the 6-row layout
            const renderProjectDetails = (project) => {
                var _a;
                if (projectNameDisplay)
                    projectNameDisplay.textContent = project.Title || "-";
                if (projectRefDisplay)
                    projectRefDisplay.textContent = project.ProjectReference || "-";
                if (profitCenterDisplay)
                    profitCenterDisplay.textContent = project.PC || "-";
                if (projectManagerDisplay) {
                    projectManagerDisplay.textContent = typeof project.ProjectManager === "object" ?
                        (_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title : project.ProjectManager || "-";
                }
                // ✅ GroupHead1 - Person/Group type
                if (groupHeadDisplay) {
                    groupHeadDisplay.textContent = getGroupHeadDisplayName(project.GroupHead1);
                }
                if (customerDisplay)
                    customerDisplay.textContent = project.CustomerName || "-";
            };
            const getForecastMonths = () => {
                return generateIdealMonthStructure();
            };
            const renderForecastInterface = (existingData, displayOBL, existingForecast) => {
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
            const protectHistoricalElements = () => {
                // Remove any existing event listeners from historical elements by cloning and replacing
                document.querySelectorAll('.invoicing-historical-month input, .invoicing-historical-month select, .invoicing-historical-month button').forEach(element => {
                    var _a;
                    const newElement = element.cloneNode(true);
                    (_a = element.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newElement, element);
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
            const renderHistoricalMonths = (months, existingData) => {
                let html = '';
                months.forEach((monthKey, index) => {
                    const monthDate = new Date(monthKey + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[monthKey] || { plan: 0, actual: 0, items: {} };
                    html += renderHistoricalMonthBox(monthKey, monthLabel, data, index);
                });
                historicalContainer.innerHTML = html;
            };
            const renderHistoricalMonthBox = (monthKey, monthLabel, data, index) => {
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
              
              ${Array.from({ length: 10 }, (_, i) => {
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
            const renderFirstRow = (months, existingData) => {
                let html = '';
                // Create single row container for all three months
                html += `<div class="invoicing-detailed-months-row">`;
                // Current Month - 44% - 5 columns (Item Description, Previous, Current, Forecast, Date)
                const currentMonth = months[0];
                if (currentMonth) {
                    const monthDate = new Date(currentMonth + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[currentMonth] || { plan: 0, actual: 0, items: {} };
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
                        const data = existingData[monthKey] || { plan: 0, actual: 0, items: {} };
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
            const getCurrentMonthMinDate = () => {
                const now = new Date();
                return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`;
            };
            const getCurrentMonthMaxDate = () => {
                const now = new Date();
                const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
                return `${lastDay.getFullYear()}-${String(lastDay.getMonth() + 1).padStart(2, "0")}-${String(lastDay.getDate()).padStart(2, "0")}`;
            };
            // RENDER: Current Month - 5 columns (Item Description, Previous, Current, Forecast, Date)
            const renderDetailedMonthBox5Col = (monthKey, monthLabel, data, isCurrentMonth = false) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const planInputClass = data.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const items = data.items || createEmptyItems();
                // ✅ Main Date is READONLY and auto-shows the HIGHEST (latest) date entered on the items.
                const computeMaxItemDate = () => {
                    var _a;
                    let latest = '';
                    for (let i = 1; i <= 10; i++) {
                        const d = ((_a = items[i.toString()]) === null || _a === void 0 ? void 0 : _a.date) || '';
                        if (d && (!latest || d > latest))
                            latest = d;
                    }
                    // fall back to the previously stored top-level date if no item has a date yet
                    return latest || data.date || '';
                };
                const autoMainDate = computeMaxItemDate();
                // ✅ Month Current is READONLY and derived = SUM of that month's item Currents below.
                let itemsTotalSum = 0;
                Array.from({ length: 10 }, (_, i) => i + 1).forEach(itemNum => {
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
              
              ${Array.from({ length: 10 }, (_, i) => i + 1).map(itemNum => {
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
            const renderDetailedMonthBox3Col = (monthKey, monthLabel, data, isCurrentMonth = false) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const planInputClass = data.plan > 0 ? 'invoicing-input-with-value' : 'invoicing-input-empty';
                const items = data.items || createEmptyItems();
                // ✅ Month Current is READONLY and derived = SUM of that month's item Currents below.
                let itemsTotalSum = 0;
                Array.from({ length: 10 }, (_, i) => i + 1).forEach(itemNum => {
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
              
              ${Array.from({ length: 10 }, (_, i) => i + 1).map(itemNum => {
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
            const renderRemainingMonths = (months, existingData) => {
                let html = '';
                for (let row = 0; row < 3; row++) {
                    html += `<div class="remaining-row-inv">`;
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
            const renderSimpleMonthBox = (monthKey, monthLabel, data) => {
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
            const attachForecastEventListeners = () => {
                // Apply number validation to all number inputs
                document.querySelectorAll('.number-input').forEach(input => {
                    allowOnlyNumbers(input);
                });
                // Set up focus handlers for all number inputs to store old values
                document.querySelectorAll('.number-input:not([readonly])').forEach(input => {
                    input.addEventListener('focus', function () {
                        this.setAttribute('data-old-value', getNumericValue(this).toString());
                    });
                });
                // ===== MAIN CURRENT INPUTS =====
                document.querySelectorAll('.main-current-input:not([readonly])').forEach(input => {
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
                        // Update visual state
                        if (currentValue === 0) {
                            this.className = this.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
                        }
                        else {
                            this.className = this.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                // ===== ITEM CURRENT INPUTS =====
                document.querySelectorAll('.item-current-input:not([readonly])').forEach(input => {
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
                        // Update visual state
                        if (currentValue === 0) {
                            this.className = this.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
                        }
                        else {
                            this.className = this.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                // ===== YEARLY CURRENT INPUTS =====
                document.querySelectorAll('.yearly-current-input-inv:not([readonly])').forEach(input => {
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
                        // Update visual state
                        if (currentValue === 0) {
                            this.className = this.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
                        }
                        else {
                            this.className = this.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                // ===== ITEM NAME INPUTS =====
                document.querySelectorAll('input[data-item]:not([readonly])').forEach(input => {
                    if (!input.classList.contains('number-input')) {
                        input.addEventListener('input', function () {
                            if (this.value === '') {
                                this.className = this.className.replace('invoicing-input-with-value', 'invoicing-input-empty');
                            }
                            else {
                                this.className = this.className.replace('invoicing-input-empty', 'invoicing-input-with-value');
                            }
                        });
                    }
                });
                // ===== AUTO MAIN DATE (readonly) =====
                // The current-month main Date is READONLY and always reflects the HIGHEST
                // (latest) date entered across that month's items.
                const updateAutoMainDate = (monthKey) => {
                    const mainDate = document.querySelector(`input.main-date-input[data-month="${monthKey}"]`);
                    const itemDates = document.querySelectorAll(`input.item-date-input[data-month="${monthKey}"]`);
                    if (!mainDate)
                        return;
                    let latest = '';
                    itemDates.forEach(input => {
                        const d = input.value || '';
                        if (d && (!latest || d > latest))
                            latest = d;
                    });
                    mainDate.value = latest;
                };
                // ===== DATE INPUTS (DATE LOCK - same as inventory/SEIBIES) =====
                // Blocks any date outside the CURRENT calendar month - whether picked, typed or pasted.
                // Dates before the 1st (previous month) or after the last day (next month) are rejected.
                // Only ITEM dates are user-editable; the main date is derived (readonly).
                document.querySelectorAll('.item-date-input:not([readonly])').forEach(input => {
                    input.addEventListener('change', function () {
                        const minAllowed = this.getAttribute('min') || getCurrentMonthMinDate();
                        const maxAllowed = this.getAttribute('max') || getCurrentMonthMaxDate();
                        if (this.value) {
                            if (this.value < minAllowed || this.value > maxAllowed) {
                                this.value = '';
                                toast("warning", `Date must be within the current month (${minAllowed} to ${maxAllowed})`);
                                return;
                            }
                        }
                        const monthKey = this.dataset.month;
                        if (monthKey)
                            updateAutoMainDate(monthKey);
                    });
                });
                // ===== VALIDATION FOR ITEM TOTALS =====
                document.querySelectorAll('.main-current-input:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
                        const monthKey = this.dataset.month;
                        if (monthKey) {
                            // 🔧 Only the detailed months (indices 3-5: current + next 2) have item
                            // rows, so their top Current is the DERIVED sum of the items.
                            // Long-term months (6+) have NO item rows - running this on them wiped
                            // the just-typed Current to '' and locked the input readonly.
                            const monthIndex = allMonths.indexOf(monthKey);
                            if (monthIndex >= 3 && monthIndex <= 5) {
                                validateSingleMonthItems(monthKey, this);
                            }
                        }
                    });
                });
                document.querySelectorAll('.item-current-input:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
                        const monthKey = this.dataset.month;
                        if (monthKey) {
                            validateSingleMonthItems(monthKey, this);
                        }
                    });
                });
                // ===== REAL-TIME OBL VALIDATION =====
                const validationInputs = document.querySelectorAll(`
            .main-current-input:not([readonly]),
            .item-current-input:not([readonly]),
            .yearly-current-input-inv:not([readonly])
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
                    toast("error", "Forecast not found");
                    return;
                }
                // ✅ Check if user can read
                if (!canRead()) {
                    toast("error", "❌ You don't have permission to view this forecast");
                    return;
                }
                // ✅ Get project and check if user can edit
                const project = projects.find(p => { var _a; return p.Id === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Id); });
                const canEdit = canEditProject(project);
                const isFrozen = forecast.Freeze || false;
                // ✅ Show form FIRST
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
                    toast("warning", "Corrupted forecast data detected. Loaded empty structure.");
                }
                forecastData = propagateMonthValues(forecastData);
                if (forecast.Project) {
                    projectSelect.value = String(forecast.Project.Id);
                    onProjectChange(forecast.Project.Id, forecast);
                }
                // ✅ Apply permission-based readonly state AFTER rendering
                setTimeout(() => {
                    var _a;
                    // If frozen OR user can't edit, set read-only
                    const isReadOnly = isFrozen || !canEdit;
                    setFormReadOnly(isReadOnly);
                    // ✅ Show/hide BOTH save buttons together (top + bottom)
                    getSaveButtons().forEach(btn => {
                        btn.style.display = isReadOnly ? 'none' : '';
                    });
                    // ✅ CHANGE 4: Update form title - same format as inventory
                    // "Edit Invoicing Forecast - {Project Name} ({Project Ref})" + Frozen / Read Only suffix
                    const formTitleEl = document.querySelector("#formTitle");
                    if (formTitleEl) {
                        const projectName = (project === null || project === void 0 ? void 0 : project.Title) || ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        const projectRef = (project === null || project === void 0 ? void 0 : project.ProjectReference) || "";
                        // ✅ "Edit Invoicing Forecast - " prefix removed; show only the project name (Ref)
                        const projectTitle = projectName
                            ? `${projectName}${projectRef ? ` (${projectRef})` : ""}`
                            : "";
                        if (isFrozen) {
                            formTitleEl.innerHTML = `🔒 ${projectTitle || "Invoicing Forecast"} (Frozen - Read Only)`;
                            formTitleEl.style.color = '#856404';
                        }
                        else if (!canEdit) {
                            formTitleEl.innerHTML = `🔒 ${projectTitle || "Invoicing Forecast"} (Read Only)`;
                            formTitleEl.style.color = '#856404';
                        }
                        else {
                            formTitleEl.innerHTML = projectTitle || "Edit Invoicing Forecast";
                            formTitleEl.style.color = '';
                        }
                    }
                }, 300);
                toast("info", `Editing forecast for ${forecast.Title}`);
            };
            /**
       * ✅ SINGLE SOURCE OF TRUTH for every save button in the Invoicing form.
       *    Both the top button (in the sticky .invoicing-page-header) and the
       *    bottom button (in .invoicing-form-actions) carry
       *    .invoicing-btn-primary + type="submit", so this helper always
       *    returns BOTH. Any code that shows / hides / disables / re-labels a
       *    save button MUST use this helper so the two never drift out of sync.
       */
            const getSaveButtons = () => Array.from(document.querySelectorAll('button.invoicing-btn-primary[type="submit"]'));
            // ========== FORM SUBMISSION ==========
            form.addEventListener("submit", async (e) => {
                e.preventDefault();
                const projectId = parseInt(projectSelect.value);
                const project = projects.find(p => p.Id === projectId);
                if (!project) {
                    toast("error", "Please select a project");
                    return;
                }
                // ✅ Check if user can edit this forecast
                if (editingId) {
                    const existingForecast = forecasts.find(f => f.Id === editingId);
                    if (!canEditProject(project)) {
                        toast("error", "❌ You don't have permission to edit this invoicing forecast");
                        return;
                    }
                    if (existingForecast && existingForecast.Freeze) {
                        toast("error", "❌ Cannot edit: This project is frozen!");
                        return;
                    }
                }
                else {
                    if (!canCreateRecord()) {
                        toast("error", "❌ You don't have permission to create invoicing forecasts");
                        return;
                    }
                }
                const isNewProject = !editingId && !forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                if (!isNewProject && forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId && f.Id !== editingId; })) {
                    toast("error", "Forecast already exists for this project!");
                    return;
                }
                const oblValue = getNumericValue(oblLastMonthInput);
                if (isNewProject && oblValue === 0) {
                    const confirmZeroOBL = await Swal.fire({
                        icon: "question",
                        title: "Zero OBL Value",
                        text: "You're setting OBL to zero for a new project. This means there's no budget limit. Continue?",
                        showCancelButton: true,
                        confirmButtonText: "Yes, Continue",
                        cancelButtonText: "No, Set OBL"
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
                // ✅ Disable BOTH save buttons while saving
                const saveBtns = getSaveButtons();
                saveBtns.forEach(btn => {
                    btn.disabled = true;
                    btn.dataset.saving = "true";
                    btn.textContent = "💾 Saving...";
                });
                const saveLoading = createLoadingOverlay("Saving Invoicing forecast...");
                loaderHost().appendChild(saveLoading);
                try {
                    const forecastData = {};
                    const allStorageMonths = getAllMonthsForStorage();
                    // Process monthly data including current month items with dates
                    allStorageMonths.forEach(monthKey => {
                        const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                        const mainPlanInput = document.querySelector(`input.main-plan-input[data-month="${monthKey}"]`);
                        if (mainCurrentInput && mainPlanInput) {
                            const monthData = {
                                plan: getNumericValue(mainPlanInput),
                                actual: getNumericValue(mainCurrentInput),
                                forecast: 0,
                                date: '',
                                items: {}
                            };
                            // Process items for detailed months (index 3-5 in allMonths array = current + next 2 months)
                            const monthIndex = allStorageMonths.indexOf(monthKey);
                            if (monthIndex >= 3 && monthIndex <= 5) {
                                const isCurrentMonth = monthIndex === 3;
                                // For the current month: item rows have Date (Forecast removed).
                                // For the next 2 months: no Date columns.
                                for (let itemNum = 1; itemNum <= 10; itemNum++) {
                                    const itemKey = itemNum.toString();
                                    const itemNameInput = document.querySelector(`input[data-item="${itemKey}"][data-month="${monthKey}"]:not([data-item-type])`);
                                    const itemPlanInput = document.querySelector(`input.invoicing-item-plan-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    const itemCurrentInput = document.querySelector(`input.item-current-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    if (itemNameInput && itemPlanInput && itemCurrentInput) {
                                        const itemData = {
                                            name: itemNameInput.value || '',
                                            plan: getNumericValue(itemPlanInput),
                                            actual: getNumericValue(itemCurrentInput),
                                            forecast: 0,
                                            date: ''
                                        };
                                        // Only the current month captures a Date per item
                                        if (isCurrentMonth) {
                                            const itemDateInput = document.querySelector(`input.item-date-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                            itemData.date = itemDateInput ? itemDateInput.value : '';
                                        }
                                        monthData.items[itemKey] = itemData;
                                    }
                                }
                            }
                            // Detailed months: the top "Current" is derived/readonly, so persist
                            // the month's actual = SUM of the item Currents.
                            if (monthIndex >= 3 && monthIndex <= 5 && monthData.items) {
                                let monthCurrent = 0;
                                Object.keys(monthData.items).forEach(k => {
                                    monthCurrent += monthData.items[k].actual || 0;
                                });
                                monthData.actual = monthCurrent;
                            }
                            // Main Date is READONLY = highest date among current-month items.
                            // (Forecast column was removed from the current month.)
                            if (monthIndex === 3 && monthData.items) {
                                let latest = '';
                                Object.keys(monthData.items).forEach(k => {
                                    const d = monthData.items[k].date || '';
                                    if (d && (!latest || d > latest))
                                        latest = d;
                                });
                                monthData.date = latest;
                            }
                            forecastData[monthKey] = monthData;
                        }
                    });
                    // Process yearly data
                    const yearlyData = {};
                    const yearlyPlanInputs = document.querySelectorAll('.yearly-plan-input-inv');
                    if (yearlyPlanInputs.length > 0) {
                        const years = new Set();
                        yearlyPlanInputs.forEach(input => {
                            const year = input.dataset.year;
                            if (year)
                                years.add(year);
                        });
                        Array.from(years).sort().forEach(year => {
                            const planInput = document.querySelector(`input.yearly-plan-input-inv[data-year="${year}"]`);
                            const currentInput = document.querySelector(`input.yearly-current-input-inv[data-year="${year}"]`);
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
                        __metadata: { type: "SP.Data.InvoicingPlanningListItem" },
                        Title: generateForecastTitle(project, propagatedData),
                        ProjectId: project.Id,
                        ProfitCentre: project.PC || "",
                        ForecastData: JSON.stringify(propagatedData),
                        UpdatedDate: new Date().toISOString()
                    };
                    // Add OBL for new projects
                    if (!editingId && isNewProject) {
                        const oblValueToSave = getNumericValue(oblLastMonthInput);
                        payload.OBL = oblValueToSave;
                    }
                    // Add CreatedDate for new items
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
                        toast("success", editingId ? "Forecast Updated!" : "Forecast Added!");
                        // IMPORTANT: Reload data and redirect to table
                        await loadForecasts();
                        initializePagination(forecasts);
                        // Force redirect to table view
                        hideForm();
                        resetForm();
                        // Ensure table section is visible
                        tableSection.style.display = "block";
                        onSaveSuccess();
                    }
                    else {
                        const errorText = await res.text();
                        console.error("SharePoint Error:", errorText);
                        toast("error", `Error saving forecast: ${res.status}`);
                    }
                }
                catch (error) {
                    console.error("Error saving forecast:", error);
                    toast("error", "Error saving forecast!");
                }
                finally {
                    removeLoadingOverlay();
                    // ✅ Re-enable BOTH save buttons
                    saveBtns.forEach(btn => {
                        btn.disabled = false;
                        btn.dataset.saving = "false";
                        btn.textContent = "💾 Save Invoicing Forecast";
                    });
                }
            });
            const initializeFilters = () => {
                const currentYear = new Date().getFullYear();
                const currentMonth = new Date().getMonth() + 1;
                filterYear.innerHTML = '<option value="">All Years</option>';
                for (let year = currentYear - 1; year <= currentYear + 2; year++) {
                    const selected = year === currentYear ? 'selected' : '';
                    filterYear.innerHTML += `<option value="${year}" ${selected}>${year}</option>`;
                }
                const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                filterMonth.innerHTML = '<option value="">All Months</option>';
                for (let month = 1; month <= 12; month++) {
                    const selected = month === currentMonth ? 'selected' : '';
                    filterMonth.innerHTML += `<option value="${month}" ${selected}>${monthNames[month - 1]}</option>`;
                }
                // ✅ Use getAssignedProjectValues to get filtered values
                const assignedProjectTitles = getAssignedProjectValues(projects, 'Title');
                const assignedCustomers = getAssignedProjectValues(projects, 'CustomerName');
                const assignedManagers = getAssignedProjectValues(projects, 'ProjectManager');
                const assignedGroupHeads = getAssignedProjectValues(projects, 'GroupHead1');
                const assignedReferences = getAssignedProjectValues(projects, 'ProjectReference');
                filterProject.innerHTML = '<option value="">All Projects</option>' +
                    assignedProjectTitles.map(project => `<option value="${project}">${project}</option>`).join("");
                filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                    assignedCustomers.map(customer => `<option value="${customer}">${customer}</option>`).join("");
                filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                    assignedManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");
                // ✅ GroupHead1 - Person/Group type
                filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
                    assignedGroupHeads.map(head => `<option value="${head}">${head}</option>`).join("");
                filterProjectReference.innerHTML = '<option value="">All References</option>' +
                    assignedReferences.map(ref => `<option value="${ref}">${ref}</option>`).join("");
            };
            const filterTable = () => {
                const year = filterYear.value;
                const month = filterMonth.value;
                const project = filterProject.value;
                const customer = filterCustomer.value;
                const projectManager = filterProjectManager.value;
                const groupHead = filterGroupHead.value;
                const projectReference = filterProjectReference.value;
                // ✅ Use getAssignedProjects for permission-based filtering
                const assignedProjects = getAssignedProjects(projects);
                const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                let filtered = [...forecasts];
                // ✅ Role access (auth.ts - same as SEIBIES / collection / revenue / sprint):
                //    FULL_CONTROL + READ_ALL roles see all projects; every other role only sees
                //    the data of projects assigned to them (own projects).
                if (!canReadAllProjects()) {
                    filtered = filtered.filter(forecast => {
                        var _a;
                        const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        return projectTitle && assignedProjectTitles.has(projectTitle);
                    });
                }
                // ✅ Apply other filters
                filtered = filtered.filter(forecast => {
                    var _a;
                    if (year && month) {
                        const monthKey = `${year}-${month.padStart(2, "0")}`;
                        const data = safeJsonParse(forecast.ForecastData, {});
                        if (!(monthKey in data))
                            return false;
                    }
                    if (project && ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) !== project)
                        return false;
                    if (customer) {
                        const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                        if ((proj === null || proj === void 0 ? void 0 : proj.CustomerName) !== customer)
                            return false;
                    }
                    // ✅ ProjectManager - Person/Group type
                    if (projectManager) {
                        const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                        let managerName = "";
                        if (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) {
                            managerName = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
                                ? proj.ProjectManager.Title || ""
                                : proj.ProjectManager || "";
                        }
                        if (managerName !== projectManager)
                            return false;
                    }
                    // ✅ GroupHead1 - Person/Group type
                    if (groupHead) {
                        const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                        let groupHeadName = "";
                        if (proj === null || proj === void 0 ? void 0 : proj.GroupHead1) {
                            groupHeadName = typeof proj.GroupHead1 === "object" && proj.GroupHead1 !== null
                                ? proj.GroupHead1.Title || ""
                                : proj.GroupHead1 || "";
                        }
                        if (groupHeadName !== groupHead)
                            return false;
                    }
                    if (projectReference) {
                        const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                        if ((proj === null || proj === void 0 ? void 0 : proj.ProjectReference) !== projectReference)
                            return false;
                    }
                    return true;
                });
                if (filtered.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:20px;">No records found for selected filters</td></tr>`;
                    const totalsHeader = document.querySelector("#totalsHeader");
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    calculateTotalPages();
                    currentPage = 1;
                    updatePaginationControls();
                }
                else {
                    filteredData = filtered;
                    calculateTotalPages();
                    currentPage = 1;
                    updatePaginationControls();
                    renderCurrentPage();
                    const displayMonths = getDisplayMonths();
                    const currentMonthKey = formatMonthKey(new Date());
                    setTimeout(() => {
                        renderTotalsHeader(displayMonths, currentMonthKey);
                    }, 100);
                }
            };
            // ============================================
            // UPDATED: SET FORM READONLY
            // ============================================
            const setFormReadOnly = (isReadOnly) => {
                const editableSelectors = [
                    '.main-current-input',
                    '.item-current-input',
                    '.yearly-current-input-inv',
                    '.item-date-input'
                ];
                const inputs = document.querySelectorAll(editableSelectors.join(','));
                inputs.forEach(input => {
                    // Derived Current (auto sum of items) must STAY locked even when the project is editable.
                    const isDerived = input.classList.contains('derived-month-current') || input.getAttribute('data-derived') === 'current';
                    if (isReadOnly || isDerived) {
                        input.setAttribute('readonly', 'true');
                        input.style.backgroundColor = '#e9ecef';
                        input.style.color = '#343a40';
                        input.style.cursor = 'not-allowed';
                        if (!isDerived)
                            input.style.opacity = '0.7';
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
                // ✅ Invoicing (OBL) value is ALWAYS readonly in the edit section - it is a
                //    saved database value. Creating a NEW forecast never calls
                //    setFormReadOnly, so entering an initial value there still works.
                const oblInput = document.querySelector('#oblLastMonthInvoicing');
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
                    toast("info", '🔓 Project is now editable.');
                }
            };
            // ============================================
            // UPDATED: FORM MANAGEMENT FUNCTIONS
            // ============================================
            const showForm = (isEditMode = false, project) => {
                // ✅ Check if user can view this form
                if (!canViewForm(isEditMode, project)) {
                    if (!canRead()) {
                        toast("error", "❌ You don't have permission to access this form");
                        return;
                    }
                    if (isEditMode && !canEditProject(project)) {
                        toast("error", "❌ You don't have permission to edit this project");
                        return;
                    }
                    toast("error", "❌ You don't have permission to create invoicing forecasts");
                    return;
                }
                formSection.style.display = "block";
                tableSection.style.display = "none";
                // ✅ Show/hide save button based on permissions
                // ✅ Show/hide BOTH save buttons together (top + bottom)
                const canSave = isEditMode ? canEditProject(project) : canCreateRecord();
                getSaveButtons().forEach(btn => {
                    btn.style.display = canSave ? '' : 'none';
                });
                // ✅ CHANGE 4: Edit title same as inventory - "Edit Invoicing Forecast - {Project} ({Ref})"
                const projectName = (project === null || project === void 0 ? void 0 : project.Title) || "";
                const projectRef = (project === null || project === void 0 ? void 0 : project.ProjectReference) || "";
                // ✅ "Edit Invoicing Forecast - " prefix removed; show only the project name (Ref)
                const editTitle = projectName
                    ? `${projectName}${projectRef ? ` (${projectRef})` : ""}`
                    : "Edit Invoicing Forecast";
                if (isEditMode) {
                    formSection.classList.add('invoicing-edit-mode');
                    projectSelect.disabled = true;
                    formTitle.textContent = editTitle;
                    formTitle.style.color = '';
                }
                else {
                    formSection.classList.remove('invoicing-edit-mode');
                    projectSelect.disabled = true;
                    formTitle.textContent = editTitle;
                    formTitle.style.color = '';
                }
            };
            // ========== SYNC MISSING PROJECTS TO INVOICING PLANNING LIST ==========
            // (same feature as Revenue / SEIBIES) - auto-adds a row for any Project List
            // project that has no InvoicingPlanning row yet. Runs silently in the
            // background (console + toast only - no on-page button/status bar).
            let invoicingProjectRefField = null;
            const invoicingListProjectRefField = async () => {
                var _a;
                if (invoicingProjectRefField)
                    return invoicingProjectRefField;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq 'ProjectReference'&$select=InternalName,TypeAsString,ReadOnlyField`, { headers: { Accept: "application/json;odata=verbose" } });
                    const data = await res.json();
                    const results = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    if (results.length > 0) {
                        invoicingProjectRefField = { exists: true, type: String(results[0].TypeAsString || "Text") };
                        console.log(`📋 InvoicingPlanning list ProjectReference column type = ${invoicingProjectRefField.type}`);
                    }
                    else {
                        invoicingProjectRefField = { exists: false, type: "" };
                        console.log("📋 InvoicingPlanning list has no ProjectReference column");
                    }
                }
                catch (error) {
                    console.error("⚠️ Could not check InvoicingPlanning ProjectReference column:", error);
                    invoicingProjectRefField = { exists: false, type: "" };
                }
                return invoicingProjectRefField;
            };
            const getInvoicingListEntityType = async () => {
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const data = await res.json();
                    return data.d.ListItemEntityTypeFullName;
                }
                catch (error) {
                    console.error("Error getting InvoicingPlanning entity type:", error);
                    // Fallback - same type constant already used by the normal Invoicing save
                    return "SP.Data.InvoicingPlanningListItem";
                }
            };
            const syncMissingProjectsToInvoicingList = async () => {
                var _a, _b, _c, _d;
                // ✅ Role guard (same as Revenue / SEIBIES): export/create needs create rights.
                if (!canCreateRecord()) {
                    console.log('👁️ No create rights - Invoicing sync skipped (Admin/GH Head/PM Manager only)');
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
                        const chkRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,ProjectId,Title&$top=5000`, { headers: { Accept: "application/json;odata=verbose" } });
                        if (chkRes.ok)
                            serverRows = ((_b = (_a = (await chkRes.json())) === null || _a === void 0 ? void 0 : _a.d) === null || _b === void 0 ? void 0 : _b.results) || [];
                    }
                    catch (chkErr) {
                        console.error("⚠️ Invoicing sync safety-check fetch failed:", chkErr);
                    }
                    // If the list really has rows on SharePoint but the page loaded 0, the
                    // main load is broken - creating rows now would duplicate everything.
                    if (serverRows.length > 0 && forecasts.length === 0) {
                        console.log(`⚠️ InvoicingPlanning list already has ${serverRows.length} row(s) on SharePoint but the page loaded 0 - sync stopped to avoid duplicates. Check the main Invoicing load.`);
                        toast("error", `Sync stopped: list has ${serverRows.length} rows but page loaded 0`);
                        return;
                    }
                    const existsOnServer = (p) => serverRows.some(c => Number(c.ProjectId) === p.Id);
                    const existsInMemory = (p) => forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === p.Id || (f.Project && f.Project.Title === p.Title); });
                    const missingProjects = visibleProjects.filter(p => !existsOnServer(p) && !existsInMemory(p));
                    if (missingProjects.length === 0) {
                        console.log(`✅ In sync: all ${visibleProjects.length} visible project(s) have an InvoicingPlanning row`);
                        return;
                    }
                    console.log(`🆕 ${missingProjects.length} missing: ${missingProjects.map(p => p.Title).join(', ')}`);
                    toast("info", `Checking ${visibleProjects.length} projects... adding ${missingProjects.length} missing`);
                    const digest = await getDigest();
                    const entityType = await getInvoicingListEntityType();
                    let added = 0;
                    let failed = 0;
                    const errors = [];
                    for (const project of missingProjects) {
                        let ok = false;
                        // Up to 3 payload variants:
                        //   1 = full create - Project lookup id as number + ProjectReference (when the
                        //       column exists: Lookup -> ProjectReferenceId, Text -> text) + OBL
                        //   2 = without ProjectReference (column may be read-only/misconfigured)
                        //   3 = without OBL too (column may be read-only/absent)
                        for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
                            try {
                                const freshData = createFreshForecastStructure();
                                // Same shape as the normal Invoicing Add-form save (ProjectId is sent as
                                // a number there and this list accepts it that way).
                                const payload = {
                                    __metadata: { type: entityType },
                                    Title: generateForecastTitle(project, freshData),
                                    ProjectId: project.Id,
                                    ProfitCentre: project.PC || "",
                                    ForecastData: JSON.stringify(freshData),
                                    // Same OBL default as the Invoicing Add form (OBL becomes readonly after
                                    // creation, so a 0 here would leave the row unusable)
                                    OBL: project.PC ? 0 : 0,
                                    CreatedDate: new Date().toISOString(),
                                    UpdatedDate: new Date().toISOString()
                                };
                                // Store Project Reference when the column exists - same rule as the
                                // Revenue sync: a LOOKUP column receives the target item ID as a string
                                // (ProjectReferenceId), a text column receives the text.
                                const refField = await invoicingListProjectRefField();
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
                        // Reload + rebuild the table so the freshly created rows show up.
                        await loadForecasts();
                        initializePagination(forecasts);
                        console.log(`➕ Invoicing sync added ${added} project(s)` + (failed > 0 ? ` (${failed} failed)` : ""));
                        toast("info", `Added ${added} missing Invoicing Planning row(s)` + (failed > 0 ? `, ${failed} failed` : ""));
                    }
                    else if (failed > 0) {
                        console.error(`❌ Invoicing sync failed for ${failed} project(s): ${errors.join(' | ')}`);
                    }
                }
                catch (error) {
                    console.error("❌ Invoicing sync failed:", error);
                }
            };
            // ========== AUTO PROGRESSION FUNCTIONS ==========
            const applyAutoProgression = async () => {
                if (forecasts.length === 0) {
                    return;
                }
                const digest = await getDigest();
                let progressedCount = 0;
                let errorCount = 0;
                for (const forecast of forecasts) {
                    try {
                        if (!forecast.ForecastData) {
                            continue;
                        }
                        const originalData = safeJsonParse(forecast.ForecastData, {});
                        // Use the enhanced check that includes yearly progression
                        if (checkIfNeedsProgression(originalData)) {
                            // Use the enhanced progression that includes yearly progression
                            const progressedData = propagateCollectionValues(originalData);
                            const payload = {
                                __metadata: { type: "SP.Data.InvoicingPlanningListItem" },
                                ForecastData: JSON.stringify(progressedData),
                                UpdatedDate: new Date().toISOString()
                            };
                            const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${forecast.Id})`, {
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
                            }
                            else {
                                errorCount++;
                            }
                            await new Promise(resolve => setTimeout(resolve, 100));
                        }
                        else {
                        }
                    }
                    catch (error) {
                        errorCount++;
                    }
                }
                if (progressedCount > 0) {
                    await loadForecasts();
                    toast("success", `Monthly and yearly data progressed for ${progressedCount} invoicing forecasts`);
                }
                else if (errorCount > 0) {
                    toast("warning", `Some errors occurred during monthly and yearly progression`);
                }
                else {
                }
            };
            const checkIfNeedsProgression = (forecastData) => {
                const months = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
                if (months.length === 0)
                    return true;
                const idealMonths = generateIdealMonthStructure();
                const firstIdealMonth = idealMonths[0];
                const firstDataMonth = months[0];
                // Check if monthly progression is needed
                const monthlyProgressionNeeded = firstDataMonth !== firstIdealMonth;
                // ALWAYS check for yearly progression, even if monthly is not needed
                const currentYear = new Date().getFullYear();
                let yearlyProgressionNeeded = false;
                if (forecastData.yearly) {
                    const years = Object.keys(forecastData.yearly).map(y => parseInt(y)).sort();
                    yearlyProgressionNeeded = years.length > 0 && years[0] < currentYear;
                }
                else {
                    yearlyProgressionNeeded = true;
                }
                return monthlyProgressionNeeded || yearlyProgressionNeeded;
            };
            // ========== ENHANCED YEARLY PROPAGATION FOR INVOICING ==========
            const propagateCollectionValues = (forecastData) => {
                const updatedData = Object.assign({}, forecastData);
                // Ensure yearly data exists before progression
                if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) {
                    const longTermMonths = generateIdealMonthStructure().slice(6);
                    const lastMonthKey = longTermMonths[longTermMonths.length - 1];
                    updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
                }
                // Detect if it's a new month by checking current month structure
                const originalMonths = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
                const idealMonths = generateIdealMonthStructure();
                const isNewMonth = originalMonths[0] !== idealMonths[0];
                // Apply monthly progression (long-term collection)
                const monthlyProgressed = propagateMonthValues(updatedData);
                // Apply yearly progression only when month changes
                const yearlyProgressed = propagateYearlyValues(monthlyProgressed, isNewMonth);
                return yearlyProgressed;
            };
            // ========== FIXED YEARLY PROPAGATION FOR INVOICING ==========
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
            const initialize = async () => {
                try {
                    const initialLoading = createLoadingOverlay("Loading Invoicing data...");
                    loaderHost().appendChild(initialLoading);
                    // ✅ Initialize shared authentication FIRST
                    await authManager.initialize({ siteUrl });
                    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                    console.log(`👤 User: ${getCurrentUserFullName()}`);
                    console.log(`📋 User Groups:`, getUserGroups());
                    console.log(`✅ Has Full Control: ${hasFullControl()}`);
                    console.log(`✅ Can Contribute: ${canContribute()}`);
                    console.log(`✅ Can Freeze: ${canFreeze()}`);
                    console.log(`✅ Is GH Head: ${isGHHead()}`);
                    console.log(`✅ Is PM Manager: ${isPMManager()}`); // ✅ NEW
                    console.log(`✅ Is Engineering User: ${isEngineeringUser()}`); // ✅ NEW
                    await loadProjects();
                    await loadForecasts();
                    // 🔐 AUTO-SYNC (same as Revenue / SEIBIES): compare the Project List with the
                    // InvoicingPlanning List and add any missing project row (empty ForecastData
                    // skeleton). Role-gated inside the function - runs silently for read-only roles
                    // and only creates rows within this role's visible project scope. Reloads
                    // internally only when it actually added rows.
                    await syncMissingProjectsToInvoicingList();
                    if (forecasts.length > 0) {
                        await applyAutoProgression();
                    }
                    initializeFilters();
                    // ✅ Call filterTable() to apply project filtering
                    filterTable();
                    // ✅ Update UI based on permissions
                    updateUIBasedOnPermissions({
                        exportBtn: '.invoicing-export-btn',
                        addBtn: '.show-form-btn',
                        saveBtn: '.invoicing-btn-primary[type="submit"]',
                        freezeCheckboxes: '.freeze-checkbox'
                    });
                    console.log("✅ Invoicing application initialized successfully");
                }
                catch (error) {
                    console.error("Failed to initialize Invoicing application:", error);
                    toast("error", "Failed to initialize Invoicing application");
                }
                finally {
                    removeLoadingOverlay();
                }
            };
            // ========== EVENT LISTENERS ==========
            backBtn.addEventListener("click", hideForm);
            projectSelect.addEventListener("change", (e) => {
                if (editingId) {
                    toast("warning", "Cannot change project in edit mode");
                    const forecast = forecasts.find(f => f.Id === editingId);
                    if (forecast === null || forecast === void 0 ? void 0 : forecast.Project) {
                        projectSelect.value = String(forecast.Project.Id);
                    }
                    return;
                }
                const projectId = parseInt(e.target.value);
                if (projectId) {
                    const project = projects.find(p => p.Id === projectId);
                    if (project) {
                        const existingForecast = forecasts.find(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                        if (existingForecast && !editingId) {
                            Swal.fire({
                                icon: "question",
                                title: "Existing Forecast Found",
                                text: `A forecast already exists for ${project.Title}. Would you like to edit it?`,
                                showCancelButton: true,
                                confirmButtonText: "Yes, Edit",
                                cancelButtonText: "No, Create New"
                            }).then((result) => {
                                if (result.isConfirmed) {
                                    editingId = existingForecast.Id;
                                    showForm(true, project);
                                    onProjectChange(projectId, existingForecast);
                                }
                                else {
                                    onProjectChange(projectId);
                                }
                            });
                        }
                        else {
                            onProjectChange(projectId, existingForecast);
                        }
                    }
                }
            });
            // Pagination event listeners
            firstPageBtn.addEventListener("click", () => goToPage(1));
            prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
            nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
            lastPageBtn.addEventListener("click", () => goToPage(totalPages));
            // Filter event listeners
            filterYear.addEventListener("change", filterTable);
            filterMonth.addEventListener("change", filterTable);
            filterProject.addEventListener("change", filterTable);
            filterCustomer.addEventListener("change", filterTable);
            filterProjectManager.addEventListener("change", filterTable);
            filterGroupHead.addEventListener("change", filterTable);
            filterProjectReference.addEventListener("change", filterTable);
            // Action button event listeners
            exportBtn.addEventListener("click", exportToExcelWithFullStyle);
            exportDetailsBtn.addEventListener("click", exportDetailedItemsToExcel);
            clearFiltersBtn.addEventListener("click", clearFilters);
            // ✅ CHANGE 2: Historical Data show/hide toggle (arrow + label, collapsed by default)
            // Bound once at startup and re-bound after each form render via setupHistoricalToggle().
            setupHistoricalToggle();
            // ✅ DERIVED CURRENT PERMANENT-LOCK: delegated capture-phase guard.
            // The top Current (auto sum of the item Currents) must NEVER accept typed/pasted/dropped values,
            // even if some OBL/readonly helper momentarily strips readonly. Veto and recompute from items.
            const protectDerivedCurrent = (e) => {
                const target = e.target;
                if (!(target instanceof HTMLInputElement))
                    return;
                const isDerived = target.classList.contains('derived-month-current') ||
                    target.getAttribute('data-derived') === 'current';
                if (!isDerived)
                    return;
                e.preventDefault();
                e.stopPropagation();
                const monthKey = target.getAttribute('data-month');
                if (monthKey) {
                    let total = 0;
                    document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]`).forEach(inp => { total += getNumericValue(inp); });
                    target.value = total > 0 ? total.toString() : '';
                }
                target.setAttribute('readonly', 'true');
                target.classList.add('invoicing-readonly-field');
                try {
                    target.blur();
                }
                catch (_) { /* noop */ }
            };
            ['beforeinput', 'input', 'change', 'paste', 'drop'].forEach(evtType => {
                document.addEventListener(evtType, protectDerivedCurrent, true);
            });
            // ✅ CHANGE 6: keep sticky header offsets correct on resize (same as inventory)
            window.addEventListener('resize', () => {
                clearTimeout(window._invoicingStickyResizeTimer);
                window._invoicingStickyResizeTimer = setTimeout(applyStickyHeaderOffsets, 150);
            });
            // Initialize the application
            initialize();
        }, 100);
    }
};
//# sourceMappingURL=invoicing.js.map
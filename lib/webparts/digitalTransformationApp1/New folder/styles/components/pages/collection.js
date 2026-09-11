import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/collection.css";
import { authManager, hasFullControl, 
// canContribute,
canRead, canReadAllProjects, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject as authCanEditProject, canCreateRecord, 
// canViewForm,
getAssignedProjects, 
// getAssignedProjectValues,
updateUIBasedOnPermissions, getDigest as getAuthDigest, getListEntityType } from "./auth";
let currentSortConfig = {
    column: '',
    direction: 'asc'
};
export default {
    content: `
<div class="collection-page fade-in">
  <section class="table-section card">
    <div class="page-header">
      <h1>Collection Management</h1>
      <button class="btn-primary show-form-btn" style="display:none">+ Add Collection</button>
    </div>
    
    <div class="table-actions">
      <div class="table-actions-row">
        <div class="tableColp4">
          <label for="filterProject">Project</label>
          <select id="filterProject"><option value="">All Projects</option></select>
        </div>
        <div class="tableColp4">
          <label for="filterCustomer">Customer</label>
          <select id="filterCustomer"><option value="">All Customers</option></select>
        </div>
        <div class="tableColp4">
          <label for="filterProjectManager">Project Manager</label>
          <select id="filterProjectManager"><option value="">All Managers</option></select>
        </div>
        <div class="tableColp4">
          <label for="filterGroupHead">Group Head</label>
          <select id="filterGroupHead"><option value="">All Group Heads</option></select>
        </div>
        <div class="tableColp4">
          <label for="filterProjectReference">Project Reference</label>
          <select id="filterProjectReference"><option value="">All References</option></select>
        </div>
        <div class="tableColp4"> 
          <label>&nbsp;</label>
          <button class="btn-secondary clear-filters-btn" title="Clear All Filters">Reset Filter</button> 
        </div>
      </div>
      
      <div class="table-actions-row">
        <div class="tableColp4">
          <label for="filterYear">Year</label>
          <select id="filterYear"><option value="">All Years</option></select>
        </div>
        <div class="tableColp4">
          <label for="filterMonth">Month</label>
          <select id="filterMonth"><option value="">All Months</option></select>
        </div>
        <div class="tableColp4"> 
          <label>&nbsp;</label>
          <button class="btn-secondary import-actual-btn">📥 Upload Actual</button>
        </div>
        <div class="tableColp4"> 
          <label>&nbsp;</label>
          <button class="btn-secondary export-btn">📊 Export Excel</button>
        </div>
      </div>
    </div>
       
    <div class="table-container">
      <div style="overflow-x:auto;">
        <table class="month-columns-table">
          <thead id="totalsHeader"></thead>
          <thead id="monthColumnsHeader"></thead>
          <tbody class="collection-table-body"></tbody>
        </table>
      </div>
       
      <div class="pagination-container">
        <div class="pagination-controls">
          <div class="pagination-size-selector">
            <label for="pageSize">Show:</label>
              <select id="pageSize" class="page-size-select">
                  <option value="10">10</option>
                  <option value="25">25</option>
                  <option value="50">50</option>
                  <option value="100">100</option>
              </select>
          </div>
          
          <div class="pagination-navigation">
            <button class="pagination-btn" id="firstPage" disabled>First</button>
            <button class="pagination-btn" id="prevPage" disabled>Previous</button>
            
            <div class="pagination-info">
              Page <span id="currentPage">1</span> of <span id="totalPages">1</span>
            </div>
            
            <div class="pagination-pages" id="pageNumbers"></div>
            
            <button class="pagination-btn" id="nextPage" disabled>Next</button>
            <button class="pagination-btn" id="lastPage" disabled>Last</button>
            
            <div class="pagination-info">
              Showing <span id="startItem">0</span>-<span id="endItem">0</span> of <span id="totalItems">0</span> records
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>

  <section class="form-section card" style="display:none;">
    <div class="page-header">
      <h1 id="formTitle">Add / Edit Collection</h1>
    </div>
    <form id="collectionForm">
      
      <div class="form-row hidden-section">
        <div class="form-group project-select-container">
          <label>Project</label>
          <select id="projectSelect" required></select>
        </div>
      </div>

      <div class="project-main-wrapper">
        <div class="project-left-side" id="projectDetails">
          <div class="detail-item">
            <span class="detail-label">Project:</span>
            <span class="detail-value" id="projectNameDisplay">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project Ref:</span>
            <span class="detail-value" id="projectRef">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Profit Center:</span>
            <span class="detail-value" id="profitCenter">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project Manager:</span>
            <span class="detail-value" id="projectManager">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Group Head:</span>
            <span class="detail-value" id="groupHead">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Customer:</span>
            <span class="detail-value" id="customer">-</span>
          </div>
        </div>

        <div class="project-right-side">
          <div class="remarks-group">
            <label for="remarksInternal">REMARKS (INTERNAL)</label>
            <textarea id="remarksInternal" class="remarks-textarea" rows="3" placeholder="Enter internal remarks..."></textarea>
          </div>
          
          <div class="remarks-group">
            <label for="remarksSprintCell">REMARKS (SPRINT CALL)</label>
            <textarea id="remarksSprintCell" class="remarks-textarea" rows="3" placeholder="Enter sprint cell remarks..."></textarea>
          </div>
          
          <div class="obl-section-container hidden-section" id="oblContainer">
            <div class="obl-content">
              <div class="obl-left">
                <div class="obl-label-group">
                  <label>Receivable Value</label>
                </div>
                <div class="obl-value-display">
                  <input type="text" id="oblLastMonth" class="obl-input number-input" placeholder="Enter OBL" readonly />
                </div>
              </div>
              <div class="obl-right">
                <div class="validation-summary">
                  <div class="validation-item">
                    <span class="validation-label">Total Current Value:</span>
                    <span class="validation-value" id="totalCurrentValue">0</span>
                  </div>
                  <div class="validation-item">
                    <span class="validation-label">Remaining Balance:</span>
                    <span class="validation-value" id="remainingValue">0</span>
                  </div>
                  <div class="validation-line">
                    <div class="line-fill" id="oblLineFill"></div>
                  </div>
                </div>
              </div>
            </div>
            <div class="validation-message" id="oblValidationMessage">
              ✅ All values are within Receivables limit
            </div>
          </div>
        </div>
      </div>

      <div class="receivable-breakdown-section">
        <div class="breakdown-header-title"></div>
        <div class="breakdown-grid">
          <div class="breakdown-field">
            <label>NORMAL DUE</label>
            <input type="text" id="oblNormalDue" class="breakdown-input number-input" readonly placeholder="0" />
          </div>
          <div class="breakdown-field">
            <label>NORMAL OVERDUE</label>
            <input type="text" id="oblNormalOverdue" class="breakdown-input number-input" readonly placeholder="0" />
          </div>
          <div class="breakdown-field">
            <label>RETENTION DUE</label>
            <input type="text" id="oblRetentionDue" class="breakdown-input number-input" readonly placeholder="0" />
          </div>
          <div class="breakdown-field">
            <label>RETENTION OVERDUE</label>
            <input type="text" id="oblRetentionOverdue" class="breakdown-input number-input" readonly placeholder="0" />
          </div>
          <div class="breakdown-field">
            <label>GST TDS</label>
            <input type="text" id="oblGstTds" class="breakdown-input number-input" readonly placeholder="0" />
          </div>
          <div class="breakdown-field">
            <label>WHT</label>
            <input type="text" id="oblWht" class="breakdown-input number-input" readonly placeholder="0" />
          </div>
        </div>
      </div>

      <div class="collection-section hidden-section" id="historicalCollectionSection">
        <div class="section-header" id="historicalSectionHeader" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; cursor: pointer; padding: 8px 12px; border-radius: 4px; transition: background-color 0.2s;">
          <h3 style="margin: 0; color: #495057; border-bottom: 2px solid #007bff; padding-bottom: 8px; flex: 1; pointer-events: none;">
            📊 Historical Data (Last 3 Months - Readonly)
          </h3>
          <button class="toggle-historical-btn" id="historicalToggleBtn" style="font-size: 12px; padding: 4px 12px; border-radius: 4px; border: 1px solid #6c757d; background-color: #6c757d; color: white; cursor: pointer; min-width: 70px; transition: all 0.3s ease;">
            ▼ Show
          </button>
        </div>
        <div id="historicalContainer" class="historical-container" style="display: none; opacity: 0; max-height: 0; overflow: hidden; transition: opacity 0.4s ease, max-height 0.4s ease, padding 0.3s ease; padding: 0;">
        </div>
      </div>
      
      <div class="collection-section hidden-section" id="detailedCollectionSection">
        <h3>Detailed Collection (Current + Next 2 Months)</h3>
        <div id="firstRowContainer" class="first-row-container"></div>
      </div>
      
      <div class="collection-section hidden-section" id="longtermCollectionSection">
        <h3>Long-term Collection (Next 16 Months)</h3>
        <div id="remainingContainer" class="remaining-container"></div>
      </div>

      <div class="collection-section hidden-section" id="yearlyCollectionSection">
        <h3>4-Year Collection</h3>
        <div id="yearlyContainer" class="yearly-container"></div>
      </div>
      
      <div class="form-actions">
        <button type="submit" class="btn-primary">💾 Save Collection Plan</button>
        <button type="button" class="btn-secondary back-to-table">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,
    script: () => {
        setTimeout(async () => {
            const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/DigitalTransformationTest";
            const listName = "Collection List";
            const projectListName = "Project List";
            // ============================================
            // UI ELEMENT VISIBILITY CONTROLS
            // ============================================
            //  const tableBody = document.querySelector(".collection-table-body") as HTMLElement;
            const formSection = document.querySelector(".form-section");
            const tableSection = document.querySelector(".table-section");
            const showFormBtn = document.querySelector(".show-form-btn");
            const projectSelect = document.querySelector("#projectSelect");
            const historicalContainer = document.querySelector("#historicalContainer");
            const firstRowContainer = document.querySelector("#firstRowContainer");
            const remainingContainer = document.querySelector("#remainingContainer");
            const yearlyContainer = document.querySelector("#yearlyContainer");
            const backBtn = document.querySelector(".back-to-table");
            const form = document.querySelector("#collectionForm");
            const exportBtn = document.querySelector(".export-btn");
            const oblLastMonthInput = document.querySelector("#oblLastMonth");
            const totalCurrentValueSpan = document.querySelector("#totalCurrentValue");
            const remainingValueSpan = document.querySelector("#remainingValue");
            const oblLineFill = document.querySelector("#oblLineFill");
            const oblValidationMessage = document.querySelector("#oblValidationMessage");
            const formTitle = document.querySelector("#formTitle");
            const filterYear = document.querySelector("#filterYear");
            const filterMonth = document.querySelector("#filterMonth");
            const filterProject = document.querySelector("#filterProject");
            const filterCustomer = document.querySelector("#filterCustomer");
            const filterProjectManager = document.querySelector("#filterProjectManager");
            const filterGroupHead = document.querySelector("#filterGroupHead");
            const filterProjectReference = document.querySelector("#filterProjectReference");
            const importActualBtn = document.querySelector(".import-actual-btn");
            const clearFiltersBtn = document.querySelector(".clear-filters-btn");
            const pageSizeSelect = document.querySelector("#pageSize");
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
            let collections = [];
            let editingId = null;
            let allMonths = [];
            let currentPage = 1;
            let pageSize = 10;
            let totalPages = 1;
            let filteredData = [];
            // let _headerRendered = false;
            // ========== UTILITY FUNCTIONS ==========    
            const toast = (icon, title) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
            const getDigest = async () => {
                return getAuthDigest(siteUrl);
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
            const generateCollectionTitle = (project, collectionData) => {
                const currentYear = new Date().getFullYear();
                let totalValue = 0;
                // Calculate total from monthly data only (exclude yearly)
                Object.keys(collectionData).forEach(key => {
                    if (key.match(/^\d{4}-\d{2}$/)) {
                        const monthData = collectionData[key];
                        if (monthData) {
                            totalValue += monthData.actual;
                        }
                    }
                });
                return `${project.Title} - ${currentYear} Collection - ${totalValue.toLocaleString()}`;
            };
            // ========== FIXED GET DISPLAY MONTHS ==========
            const getNumericValue = (input) => {
                const value = input.value.trim();
                if (value === '' || value === '0')
                    return 0;
                return parseInt(value) || 0;
            };
            const setNumericValue = (input, value) => {
                if (value === 0) {
                    input.value = '';
                    input.className = input.className.replace('input-with-value', 'input-empty');
                }
                else {
                    input.value = value.toString();
                    input.className = input.className.replace('input-empty', 'input-with-value');
                }
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
                        this.className = this.className.replace('input-with-value', 'input-empty');
                    }
                    else {
                        this.className = this.className.replace('input-empty', 'input-with-value');
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
            // Enhanced OBL validation with toast notifications
            const showOBLErrorToast = (message, isWarning = false) => {
                var _a;
                const existingToast = document.querySelector('.obl-error-toast');
                if (existingToast)
                    existingToast.remove();
                const toast = document.createElement('div');
                toast.className = `obl-error-toast ${isWarning ? 'warning' : ''}`;
                toast.innerHTML = `
          ${message}
          <button class="toast-close">&times;</button>
        `;
                document.body.appendChild(toast);
                setTimeout(() => {
                    if (toast.parentNode) {
                        toast.remove();
                    }
                }, 5000);
                (_a = toast.querySelector('.toast-close')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
                    toast.remove();
                });
            };
            // ========== MONTH STRUCTURE FUNCTIONS ==========
            // ========== YEARLY COLLECTION FUNCTIONS ==========    
            // Type guard to check if data is CollectionMonthData
            const isCollectionMonthData = (data) => {
                return data && typeof data.plan === 'number' && typeof data.actual === 'number';
            };
            // 1. First, update the checkIfNeedsProgression function in collection:
            const checkIfNeedsProgression = (collectionData) => {
                const now = new Date();
                const currentMonthKey = formatMonthKey(now);
                const months = Object.keys(collectionData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
                if (months.length === 0)
                    return true;
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
                }
                else {
                    yearlyProgressionNeeded = true;
                }
                console.log(`🔍 Progression check: monthly=${monthlyProgressionNeeded}, yearly=${yearlyProgressionNeeded}`);
                return monthlyProgressionNeeded || yearlyProgressionNeeded;
            };
            const propagateMonthValues = (collectionData) => {
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
                const updatedData = {};
                // First pass: Create or copy all months
                idealMonths.forEach((monthKey, index) => {
                    const existingData = collectionData[monthKey];
                    const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
                    const isLongTermMonth = index >= detailedStartIndex + 3;
                    if (existingData && isCollectionMonthData(existingData)) {
                        const copied = JSON.parse(JSON.stringify(existingData));
                        if (copied.monthActual === undefined) {
                            copied.monthActual = 0;
                        }
                        updatedData[monthKey] = copied;
                    }
                    else {
                        let items = {};
                        if (isDetailedMonth) {
                            items = createEmptyCategories();
                        }
                        else if (isLongTermMonth) {
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
                    // Skip if either month data is missing or invalid
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
                        // ✅ CRITICAL FIX: Only process if both items exist
                        if (isDetailedMonth) {
                            // Ensure items objects exist
                            if (!currMonthData.items) {
                                currMonthData.items = {};
                            }
                            if (!prevMonthData.items) {
                                prevMonthData.items = {};
                            }
                            const currItems = currMonthData.items;
                            const prevItems = prevMonthData.items;
                            // Iterate over all categories in current month
                            Object.keys(currItems).forEach(categoryKey => {
                                // ✅ Safe access - check both exist
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
                }
                else {
                    const lastMonthKey = idealMonths[idealMonths.length - 1];
                    updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
                }
                return updatedData;
            };
            // ========== DATA LOADING FUNCTIONS ==========
            // SET FORM READONLY FUNCTION
            // ============================================
            const setFormReadOnly = (isFrozen) => {
                const editableSelectors = [
                    '.detailed-month-box .category-current-input',
                    '.simple-month .main-current-input',
                    '.yearly-box .yearly-current-input',
                    '.category-current-input',
                    '.category-forecast-input', // ✅ Further fields editable when frozen
                    // '.forecast-date-input',       // ✅ Date fields editable when frozen
                    // '.category-date-input',       // ✅ Date fields editable when frozen
                    '#remarksInternal',
                    '#remarksSprintCell'
                ];
                const inputs = document.querySelectorAll(editableSelectors.join(','));
                inputs.forEach(input => {
                    if (isFrozen) {
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
                // ✅ month-actual-input should ALWAYS be readonly (like revenue.ts)
                document.querySelectorAll('.month-actual-input').forEach((input) => {
                    const el = input;
                    el.setAttribute('readonly', 'true');
                    el.style.backgroundColor = '#e8f5e9';
                    el.style.cursor = 'not-allowed';
                    el.style.pointerEvents = 'none';
                });
                // ✅ forecast-total-input should ALWAYS be readonly (like revenue.ts)
                document.querySelectorAll('.forecast-total-input').forEach((input) => {
                    const el = input;
                    el.setAttribute('readonly', 'true');
                    el.style.backgroundColor = '#f0f8ff';
                    el.style.cursor = 'not-allowed';
                    el.style.pointerEvents = 'none';
                });
                // ✅ main-forecast-input (Further) should be editable when frozen
                document.querySelectorAll('.main-forecast-input').forEach((input) => {
                    const el = input;
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
                // ✅ category-forecast-input (Further) should be editable when frozen
                document.querySelectorAll('.category-forecast-input').forEach((input) => {
                    const el = input;
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
                const oblInput = document.querySelector('#oblLastMonth');
                if (oblInput) {
                    if (isFrozen) {
                        oblInput.setAttribute('readonly', 'true');
                        oblInput.style.backgroundColor = '#f5f5f5';
                        oblInput.style.cursor = 'not-allowed';
                    }
                    else {
                        oblInput.removeAttribute('readonly');
                        oblInput.style.backgroundColor = '';
                        oblInput.style.cursor = '';
                    }
                }
                if (isFrozen) {
                    toast("info", '🔒 This project is frozen. Further and Date fields are editable.');
                }
                else {
                    toast("info", '🔓 Project unfrozen. You can now edit all fields.');
                }
            };
            // ============================================
            // MONTH STRUCTURE FUNCTIONS
            // ============================================
            const createEmptyCategories = () => {
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
            const createSimpleCategories = () => {
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
            const generateIdealMonthStructure = () => {
                const months = [];
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
            const createFreshYearlyStructure = (lastMonthKey) => {
                const yearlyData = {};
                let startYear;
                if (lastMonthKey) {
                    startYear = parseInt(lastMonthKey.split('-')[0]);
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
                            "Normal": { name: 'Normal', plan: 0, actual: 0 },
                            "Retention": { name: 'Retention', plan: 0, actual: 0 }
                        }
                    };
                }
                return yearlyData;
            };
            const createFreshCollectionStructure = () => {
                const idealMonths = generateIdealMonthStructure();
                const freshData = {};
                const now = new Date();
                const currentMonthKey = formatMonthKey(now);
                const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
                const detailedStartIndex = currentMonthIndex !== -1 ? currentMonthIndex : 36;
                idealMonths.forEach((monthKey, index) => {
                    const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
                    let items = {};
                    if (isDetailedMonth) {
                        items = createEmptyCategories();
                    }
                    else if (index >= detailedStartIndex + 3) {
                        items = createSimpleCategories();
                    }
                    // ✅ ALWAYS include monthActual for ALL months
                    freshData[monthKey] = {
                        plan: 0,
                        actual: 0,
                        monthActual: 0, // ✅ Always initialize monthActual
                        forecast: 0,
                        date: '',
                        items: items
                    };
                });
                const lastMonthKey = idealMonths[idealMonths.length - 1];
                freshData.yearly = createFreshYearlyStructure(lastMonthKey);
                return freshData;
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
            // ============================================
            // DATA LOADING FUNCTIONS
            // ============================================
            const loadProjects = async () => {
                var _a;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items?$select=Id,Title,CustomerName,GroupHead,PC,ProjectReference,ProjectManager/Title&$expand=ProjectManager`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const data = await res.json();
                    const rawProjects = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    const cleanedProjects = rawProjects.map((p) => (Object.assign(Object.assign({}, p), { CustomerName: p.CustomerName || "", GroupHead: p.GroupHead || "", PC: p.PC || "", ProjectReference: p.ProjectReference || "", ProjectManager: p.ProjectManager || "" })));
                    const projectMap = new Map();
                    cleanedProjects.forEach(p => {
                        projectMap.set(p.Title, p);
                    });
                    projects = Array.from(projectMap.values());
                    console.log('📋 All projects with PMs:');
                    projects.forEach(p => {
                        var _a;
                        const pmName = typeof p.ProjectManager === 'object'
                            ? (_a = p.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title
                            : p.ProjectManager;
                        console.log(`  - ${p.Title}: PM = ${pmName || 'None'}`);
                    });
                    // ✅ Get projects for dropdown based on role using auth functions
                    const isGHHeadUser = isGHHead();
                    const isFullControlUser = hasFullControl();
                    const isReadAll = canReadAllProjects();
                    let dropdownProjects = projects;
                    // Only filter if NOT GH Head, NOT Full Control, and NOT Read All
                    if (!isGHHeadUser && !isFullControlUser && !isReadAll) {
                        dropdownProjects = getAssignedProjects(projects);
                    }
                    console.log(`📋 Total projects: ${projects.length}, Dropdown projects: ${dropdownProjects.length}`);
                    console.log('📋 Dropdown project titles:', dropdownProjects.map(p => p.Title));
                    // ✅ Populate project select dropdown
                    projectSelect.innerHTML = `<option value="">Select Project</option>` +
                        dropdownProjects.map(p => `<option value="${p.Id}">${p.Title}</option>`).join("");
                }
                catch (error) {
                    console.error("Error loading projects:", error);
                    toast("error", "Failed to load projects");
                }
            };
            // ============================================
            // UPDATED LOAD COLLECTIONS
            // ============================================
            const loadCollections = async () => {
                var _a;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?` +
                        `$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,CollectionData,RevenueActual,Remarks,CreatedDate,UpdatedDate,NormalDue,NormalOverdue,RetentionDue,RetentionOverdue,GSTTDS,WHT,Freeze` +
                        `&$expand=Project&$top=5000`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    if (!res.ok) {
                        const errorText = await res.text();
                        console.error("❌ API Error:", errorText);
                        throw new Error(`HTTP ${res.status}: ${errorText}`);
                    }
                    const data = await res.json();
                    // ✅ Determine which projects to show based on role using auth functions
                    const isGHHeadUser = isGHHead();
                    const isFullControlUser = hasFullControl();
                    const isReadAll = canReadAllProjects();
                    let allowedProjectTitles = new Set();
                    if (isGHHeadUser || isFullControlUser || isReadAll) {
                        allowedProjectTitles = new Set(projects.map(p => p.Title));
                        console.log(`👁️ Showing ALL projects (${allowedProjectTitles.size})`);
                    }
                    else {
                        const assignedProjects = getAssignedProjects(projects);
                        allowedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                        console.log(`👁️ Showing assigned projects only (${allowedProjectTitles.size})`);
                    }
                    collections = (((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || []).map(collection => {
                        let revenueActualObj = {};
                        try {
                            if (typeof collection.RevenueActual === 'string') {
                                revenueActualObj = JSON.parse(collection.RevenueActual) || {};
                            }
                            else if (typeof collection.RevenueActual === 'object') {
                                revenueActualObj = collection.RevenueActual || {};
                            }
                        }
                        catch (e) {
                            revenueActualObj = {};
                        }
                        return Object.assign(Object.assign({}, collection), { CollectionData: typeof collection.CollectionData === 'string' ?
                                collection.CollectionData :
                                JSON.stringify(collection.CollectionData || {}), _revenueActual: revenueActualObj, ReceivableValue: Number(collection.OBL) || 0, GSTTDS: Number(collection.GSTTDS) || 0, NormalDue: Number(collection.NormalDue) || 0, NormalOverdue: Number(collection.NormalOverdue) || 0, RetentionDue: Number(collection.RetentionDue) || 0, RetentionOverdue: Number(collection.RetentionOverdue) || 0, WHT: Number(collection.WHT) || 0 });
                    });
                    // ✅ FILTER: Only keep collections for allowed projects
                    const beforeCount = collections.length;
                    collections = collections.filter(collection => {
                        var _a;
                        const projectTitle = (_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        const isAllowed = projectTitle && allowedProjectTitles.has(projectTitle);
                        if (!isAllowed && projectTitle) {
                            console.log(`🔍 Filtering out: ${projectTitle} (not allowed for this user)`);
                        }
                        return isAllowed;
                    });
                    console.log(`✅ Loaded ${collections.length} collections (filtered from ${beforeCount})`);
                }
                catch (error) {
                    console.error("Error loading collections:", error);
                    toast("error", "Failed to load collections");
                }
            };
            const getListItemType = async () => {
                return getListEntityType(siteUrl, listName);
            };
            // ========== TABLE RENDERER WITH VIRTUAL SCROLLING ==========
            const attachEditTriggerListeners = () => {
                document.querySelectorAll('.edit-trigger').forEach(link => {
                    link.addEventListener('click', (e) => {
                        e.preventDefault();
                        const id = parseInt(e.currentTarget.dataset.id);
                        // ✅ Find the collection and pass it
                        const collection = collections.find(f => f.Id === id);
                        editCollection(id, collection);
                    });
                });
            };
            const attachFreezeEventListeners = () => {
                console.log("🔧 Attaching freeze event listeners...");
                const checkboxes = document.querySelectorAll('.freeze-checkbox');
                const canUserFreeze = canFreeze();
                checkboxes.forEach(checkbox => {
                    var _a;
                    const cb = checkbox;
                    // Remove existing listeners by cloning
                    const newCb = cb.cloneNode(true);
                    (_a = cb.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newCb, cb);
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
                        newCb.onchange = async function (e) {
                            e.preventDefault();
                            e.stopPropagation();
                            const target = e.target;
                            const id = parseInt(target.dataset.id);
                            const isChecked = target.checked;
                            console.log(`🔄 Freeze toggle: ID=${id}, checked=${isChecked}`);
                            const collection = collections.find(f => f.Id === id);
                            if (!collection) {
                                toast("error", "Collection not found");
                                target.checked = !isChecked;
                                return;
                            }
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
                                }
                                else {
                                    toast("error", "Failed to update freeze status");
                                    target.checked = !isChecked;
                                }
                            }
                            catch (error) {
                                console.error("Error updating freeze status:", error);
                                toast("error", "Failed to update freeze status");
                                target.checked = !isChecked;
                            }
                        };
                    }
                });
            };
            /**
             * Apply freeze status to ALL form fields - comprehensive version
             */
            const applyFreezeToAllFields = (isFrozen) => {
                var _a, _b, _c;
                console.log(`🔒 ${isFrozen ? 'FREEZING' : 'UNFREEZING'} all form fields`);
                // 1. OBL Input
                const oblInput = document.querySelector('#oblLastMonth');
                if (oblInput) {
                    if (isFrozen) {
                        oblInput.setAttribute('readonly', 'true');
                        oblInput.style.backgroundColor = '#f5f5f5';
                        oblInput.style.cursor = 'not-allowed';
                        oblInput.style.opacity = '0.7';
                    }
                    else {
                        oblInput.removeAttribute('readonly');
                        oblInput.style.backgroundColor = '';
                        oblInput.style.cursor = '';
                        oblInput.style.opacity = '1';
                    }
                }
                // 2. OBL Breakdown fields (readonly always, but style changes)
                const breakdownFields = ['oblNormalDue', 'oblNormalOverdue', 'oblRetentionDue',
                    'oblRetentionOverdue', 'oblGstTds', 'oblWht'];
                breakdownFields.forEach(id => {
                    const el = document.getElementById(id);
                    if (el) {
                        if (isFrozen) {
                            el.style.backgroundColor = '#f5f5f5';
                            el.style.opacity = '0.7';
                        }
                        else {
                            el.style.backgroundColor = '';
                            el.style.opacity = '1';
                        }
                    }
                });
                // 3. ALL input fields in the form (comprehensive)
                const allInputs = document.querySelectorAll('input, textarea');
                allInputs.forEach(input => {
                    const el = input;
                    // Skip checkbox inputs
                    if (el.type === 'checkbox')
                        return;
                    // Skip the freeze checkbox itself
                    if (el.classList.contains('freeze-checkbox'))
                        return;
                    // ✅ IMPORTANT: ALWAYS skip historical months - they should never be editable
                    if (el.closest('.historical-month')) {
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
                    }
                    else {
                        el.removeAttribute('readonly');
                        el.style.backgroundColor = '';
                        el.style.color = '';
                        el.style.cursor = '';
                        el.style.opacity = '1';
                        el.style.pointerEvents = '';
                        // Restore specific styles for editable fields
                        if (el.classList.contains('category-current-input') && !el.closest('.historical-month')) {
                            el.style.backgroundColor = '#ffffff';
                            el.style.cursor = 'text';
                            el.style.pointerEvents = 'auto';
                        }
                        // Restore main current inputs - they should be readonly but styled properly
                        if (el.classList.contains('main-current-input') && !el.closest('.historical-month')) {
                            el.setAttribute('readonly', 'true');
                            el.style.backgroundColor = '#e9ecef';
                            el.style.cursor = 'not-allowed';
                            el.style.pointerEvents = 'none';
                        }
                        // Restore main plan inputs - they should be readonly
                        if (el.classList.contains('main-plan-input') || el.classList.contains('category-plan-input')) {
                            el.setAttribute('readonly', 'true');
                            el.style.backgroundColor = '#e9ecef';
                            el.style.cursor = 'not-allowed';
                            el.style.pointerEvents = 'none';
                        }
                        // Restore yearly current inputs - they should be readonly
                        if (el.classList.contains('yearly-current-input') || el.classList.contains('yearly-plan-input')) {
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
                    document.querySelectorAll('.category-forecast-input').forEach((input) => {
                        const el = input;
                        el.setAttribute('readonly', 'true');
                        el.style.backgroundColor = '#f5f5f5';
                        el.style.cursor = 'not-allowed';
                        el.style.pointerEvents = 'none';
                        el.style.opacity = '0.7';
                    });
                }
                else {
                    // When unfrozen, make forecast inputs editable
                    document.querySelectorAll('.category-forecast-input').forEach((input) => {
                        const el = input;
                        el.removeAttribute('readonly');
                        el.style.backgroundColor = '#ffffff';
                        el.style.cursor = 'text';
                        el.style.pointerEvents = 'auto';
                        el.style.opacity = '1';
                    });
                }
                // 5. Remarks fields
                const remarksInternal = document.querySelector('#remarksInternal');
                const remarksSprintCell = document.querySelector('#remarksSprintCell');
                [remarksInternal, remarksSprintCell].forEach(el => {
                    if (el) {
                        if (isFrozen) {
                            el.setAttribute('readonly', 'true');
                            el.style.backgroundColor = '#f5f5f5';
                            el.style.cursor = 'not-allowed';
                            el.style.opacity = '0.7';
                        }
                        else {
                            el.removeAttribute('readonly');
                            el.style.backgroundColor = '';
                            el.style.cursor = '';
                            el.style.opacity = '1';
                        }
                    }
                });
                // 6. Date inputs (always readonly)
                // document.querySelectorAll('.forecast-date-input, .category-date-input').forEach((input) => {
                //   const el = input as HTMLInputElement;
                //   el.setAttribute('readonly', 'true');
                //   if (isFrozen) {
                //     el.style.backgroundColor = '#f5f5f5';
                //     el.style.opacity = '0.7';
                //   } else {
                //     el.style.backgroundColor = '#e9ecef';
                //     el.style.opacity = '1';
                //   }
                //   el.style.cursor = 'not-allowed';
                //   el.style.pointerEvents = 'none';
                // });
                // DATE INPUTS - ALWAYS EDITABLE
                // 6. Date inputs - ALWAYS editable (even when frozen)
                document.querySelectorAll('.forecast-date-input, .category-date-input').forEach((input) => {
                    const el = input;
                    el.removeAttribute('readonly');
                    el.disabled = false;
                    el.style.backgroundColor = '#ffffff';
                    el.style.cursor = 'pointer';
                    el.style.pointerEvents = 'auto';
                    el.style.opacity = '1';
                });
                // 7. Update validation message
                const validationMessage = document.querySelector('.validation-message');
                if (validationMessage) {
                    if (isFrozen) {
                        validationMessage.textContent = '🔒 Project is frozen - all fields are read-only';
                        validationMessage.className = 'validation-message info';
                    }
                    else {
                        validationMessage.textContent = '✅ All values are within Receivables limit';
                        validationMessage.className = 'validation-message success';
                        // Re-validate
                        validateOBLLimit();
                    }
                }
                // 8. Update the form title with freeze indicator
                const formTitle = document.querySelector('#formTitle');
                if (formTitle) {
                    if (isFrozen) {
                        if (!((_a = formTitle.textContent) === null || _a === void 0 ? void 0 : _a.includes('🔒'))) {
                            formTitle.textContent = `🔒 ${((_b = formTitle.textContent) === null || _b === void 0 ? void 0 : _b.replace('🔒 ', '')) || 'Edit Collection'}`;
                        }
                    }
                    else {
                        formTitle.textContent = ((_c = formTitle.textContent) === null || _c === void 0 ? void 0 : _c.replace('🔒 ', '')) || 'Edit Collection';
                    }
                }
                toast("info", isFrozen ? '🔒 Project frozen. All fields are read-only.' : '🔓 Project unfrozen. You can now edit fields.');
            };
            const getSafeString = (value, defaultValue = "-") => {
                var _a;
                if (value === null || value === undefined)
                    return defaultValue;
                if (typeof value === 'string')
                    return value.trim() || defaultValue;
                if (typeof value === 'object' && 'Title' in value)
                    return ((_a = value.Title) === null || _a === void 0 ? void 0 : _a.trim()) || defaultValue;
                return String(value) || defaultValue;
            };
            // ========== EXPORT FUNCTIONS ==========
            const getMonthName = (monthNumber) => {
                const months = [
                    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
                ];
                return months[parseInt(monthNumber) - 1] || "Month";
            };
            const exportToExcelWithFullStyle = () => {
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
                    // REMOVED: Historical months (not needed)
                    // Detailed months (3 months: current + next 2) - shown in table as detailed view
                    const detailedMonths = allIdealMonths.slice(36, 39);
                    // Long-term months (15 months) - shown in table as simple view
                    const longTermMonths = allIdealMonths.slice(39, 54);
                    // Get the last month from long-term to determine yearly start
                    const lastMonthKey = longTermMonths[longTermMonths.length - 1];
                    const startYear = parseInt(lastMonthKey.split('-')[0]);
                    // Build headers - ONLY detailed, long-term, and yearly
                    const baseHeaders = [
                        "Freeze",
                        "Project Reference",
                        "Project Name",
                        "Project Manager",
                        "Group Head",
                        "Customer",
                        "Receivables (incl TDS)"
                    ];
                    // 1. Detailed Months (3 months: current + next 2)
                    const detailedHeaders = [];
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
                    const longTermHeaders = [];
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
                    const yearlyHeaders = [];
                    for (let i = 0; i < 4; i++) {
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
                        var _a, _b;
                        const projectTitle = ((_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        // Parse data from CollectionData
                        const collectionData = safeJsonParse(collection.CollectionData, {});
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
                        // Create row data
                        const rowData = {
                            "Freeze": isFrozen ? "🔒" : "",
                            "Project Reference": (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                            "Project Name": projectTitle,
                            "Project Manager": typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                            "Group Head": (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
                            "Customer": (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                            "Receivables (incl TDS)": safeOblTotal
                        };
                        // Helper function to add detailed month data
                        const addDetailedMonthData = (monthKey) => {
                            var _a, _b, _c, _d, _e, _f, _g, _h;
                            const monthData = collectionData[monthKey] || { plan: 0, actual: 0, forecast: 0, items: {} };
                            const items = monthData.items || {};
                            const monthLabel = formatMonthLabel(new Date(monthKey + "-01"));
                            const isCurrentMonth = monthKey === currentMonthKey;
                            const prevNormal = (((_a = items.Normal) === null || _a === void 0 ? void 0 : _a.plan) || 0) + (((_b = items['Normal Overdue']) === null || _b === void 0 ? void 0 : _b.plan) || 0);
                            const prevRetention = (((_c = items.Retention) === null || _c === void 0 ? void 0 : _c.plan) || 0) + (((_d = items['Retention Overdue']) === null || _d === void 0 ? void 0 : _d.plan) || 0);
                            const prevTotal = prevNormal + prevRetention;
                            const currNormal = (((_e = items.Normal) === null || _e === void 0 ? void 0 : _e.actual) || 0) + (((_f = items['Normal Overdue']) === null || _f === void 0 ? void 0 : _f.actual) || 0);
                            const currRetention = (((_g = items.Retention) === null || _g === void 0 ? void 0 : _g.actual) || 0) + (((_h = items['Retention Overdue']) === null || _h === void 0 ? void 0 : _h.actual) || 0);
                            const currTotal = currNormal + currRetention;
                            rowData[`${monthLabel} - Prev Normal`] = prevNormal;
                            rowData[`${monthLabel} - Prev Retention`] = prevRetention;
                            rowData[`${monthLabel} - Prev Total`] = prevTotal;
                            rowData[`${monthLabel} - Curr Normal`] = currNormal;
                            rowData[`${monthLabel} - Curr Retention`] = currRetention;
                            rowData[`${monthLabel} - Curr Total`] = currTotal;
                            if (isCurrentMonth) {
                                // Get actual revenue value
                                let actualValue = 0;
                                try {
                                    if (collection._revenueActual && typeof collection._revenueActual === 'object') {
                                        actualValue = Number(collection._revenueActual[monthKey]) || 0;
                                    }
                                    else if (typeof collection.RevenueActual === 'string') {
                                        const parsed = JSON.parse(collection.RevenueActual);
                                        if (typeof parsed === 'object' && parsed !== null) {
                                            actualValue = Number(parsed[monthKey]) || 0;
                                        }
                                    }
                                    else if (typeof collection.RevenueActual === 'object') {
                                        actualValue = Number(collection.RevenueActual[monthKey]) || 0;
                                    }
                                }
                                catch (e) {
                                    actualValue = 0;
                                }
                                const furtherValue = monthData.forecast || 0;
                                const forecastValue = currTotal + furtherValue;
                                rowData[`${monthLabel} - Actual`] = actualValue;
                                rowData[`${monthLabel} - Further`] = furtherValue;
                                rowData[`${monthLabel} - Forecast`] = forecastValue;
                            }
                        };
                        // Helper function to add long-term month data
                        const addLongTermMonthData = (monthKey) => {
                            var _a, _b, _c, _d, _e, _f, _g, _h;
                            const monthData = collectionData[monthKey] || { plan: 0, actual: 0, items: {} };
                            const items = monthData.items || {};
                            const monthLabel = formatMonthLabel(new Date(monthKey + "-01"));
                            const prevNormal = (((_a = items.Normal) === null || _a === void 0 ? void 0 : _a.plan) || 0) + (((_b = items['Normal Overdue']) === null || _b === void 0 ? void 0 : _b.plan) || 0);
                            const prevRetention = (((_c = items.Retention) === null || _c === void 0 ? void 0 : _c.plan) || 0) + (((_d = items['Retention Overdue']) === null || _d === void 0 ? void 0 : _d.plan) || 0);
                            const currNormal = (((_e = items.Normal) === null || _e === void 0 ? void 0 : _e.actual) || 0) + (((_f = items['Normal Overdue']) === null || _f === void 0 ? void 0 : _f.actual) || 0);
                            const currRetention = (((_g = items.Retention) === null || _g === void 0 ? void 0 : _g.actual) || 0) + (((_h = items['Retention Overdue']) === null || _h === void 0 ? void 0 : _h.actual) || 0);
                            rowData[`${monthLabel} - Prev Normal`] = prevNormal;
                            rowData[`${monthLabel} - Prev Retention`] = prevRetention;
                            rowData[`${monthLabel} - Curr Normal`] = currNormal;
                            rowData[`${monthLabel} - Curr Retention`] = currRetention;
                        };
                        // Helper function to add yearly data
                        const addYearlyData = (year) => {
                            var _a, _b, _c, _d, _e;
                            const yearData = ((_a = collectionData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0, items: {} };
                            const items = yearData.items || {};
                            rowData[`${year} - Prev`] = yearData.plan || 0;
                            rowData[`${year} - Curr`] = yearData.actual || 0;
                            rowData[`${year} - Prev Normal`] = ((_b = items.Normal) === null || _b === void 0 ? void 0 : _b.plan) || 0;
                            rowData[`${year} - Prev Retention`] = ((_c = items.Retention) === null || _c === void 0 ? void 0 : _c.plan) || 0;
                            rowData[`${year} - Curr Normal`] = ((_d = items.Normal) === null || _d === void 0 ? void 0 : _d.actual) || 0;
                            rowData[`${year} - Curr Retention`] = ((_e = items.Retention) === null || _e === void 0 ? void 0 : _e.actual) || 0;
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
                        for (let i = 0; i < 4; i++) {
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
                    const numberStyle = Object.assign(Object.assign({}, dataStyle), { alignment: { horizontal: "right", vertical: "center" } });
                    // Apply styles to all cells
                    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                    // Style header row
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s)
                            ws[cellAddress].s = {};
                        ws[cellAddress].s = headerStyle;
                    }
                    // Style data rows
                    for (let row = range.s.r + 1; row <= range.e.r; row++) {
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                            if (!ws[cellAddress])
                                continue;
                            const isNumberColumn = col >= 6;
                            const styleToApply = isNumberColumn ? numberStyle : dataStyle;
                            if (!ws[cellAddress].s)
                                ws[cellAddress].s = {};
                            ws[cellAddress].s = styleToApply;
                        }
                    }
                    // Set column widths
                    const colWidths = finalHeaders.map((header) => {
                        if (header === 'Freeze')
                            return { wch: 8 };
                        if (header.includes('Project Name'))
                            return { wch: 30 };
                        if (header.includes('Project Reference'))
                            return { wch: 18 };
                        if (header.includes('Project Manager') || header.includes('Group Head'))
                            return { wch: 18 };
                        if (header.includes('Customer'))
                            return { wch: 20 };
                        if (header.includes('Receivables'))
                            return { wch: 18 };
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
                }
                catch (error) {
                    console.error("❌ Error during styled Excel export:", error);
                    toast("error", "Failed to export data to Excel");
                }
            };
            // const generateSelectedMonthHeaders = (selectedYear: string, selectedMonth: string, exportData: CollectionItem[]) => {
            //   // CORRECTED: New column sequence
            //   const baseHeaders = [
            //     "Project Reference",  // Moved to first
            //     "Project Name",       // Moved to second  
            //     "PC",                 // Moved to third
            //     "Customer Name",      // Moved to fourth
            //     "Group Head", 
            //     "Project Manager", 
            //     "Normal Due", 
            //     "Normal Overdue", 
            //     "Retention Due", 
            //     "Retention Overdue", 
            //     "GST-TDS", 
            //     "WHT", 
            //     "Total (normal to wht)"
            //   ];
            //   const allMonthKeys = new Set<string>();
            //   exportData.forEach(collection => {
            //     const collectionData = safeJsonParse(collection.CollectionData, {});
            //     Object.keys(collectionData).forEach(key => {
            //       if (key && key.match(/^\d{4}-\d{2}$/)) {
            //         allMonthKeys.add(key);
            //       }
            //     });
            //   });
            //   const sortedMonthKeys = Array.from(allMonthKeys).sort();
            //   const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
            //   // Get months after selected month (EXCLUDING selected month)
            //   const allMonthsAfterSelected = sortedMonthKeys.filter(monthKey => monthKey > selectedMonthKey);
            //   const individualMonthHeaders: string[] = [];
            //   allMonthsAfterSelected.forEach(monthKey => {
            //     const [year, month] = monthKey.split('-');
            //     const monthLabel = getMonthName(month);
            //     const displayLabel = `${monthLabel} ${year}`;
            //     individualMonthHeaders.push(`${displayLabel} (P)`, `${displayLabel} (C)`);
            //   });
            //   // Get the year from the LAST month in the headers
            //   let lastMonthYear = selectedYear;
            //   if (allMonthsAfterSelected.length > 0) {
            //     const lastMonthKey = allMonthsAfterSelected[allMonthsAfterSelected.length - 1];
            //     lastMonthYear = lastMonthKey.split('-')[0];
            //   }
            //   // Use the last month year to generate next 4 years
            //   const nextFourYears = getNextFourYears(lastMonthYear);
            //   const yearlyHeaders: string[] = [];
            //   nextFourYears.forEach(year => {
            //     yearlyHeaders.push(`${year} (P)`, `${year} (C)`);
            //   });
            //   // CORRECTED: Headers WITHOUT selected month columns
            //   const finalHeaders = [
            //     ...baseHeaders,
            //     ...individualMonthHeaders, // Only months after selected month (NO Nov 2025)
            //     ...yearlyHeaders,
            //     "Balance"
            //   ];
            //   return { 
            //     headers: finalHeaders, 
            //     allMonthsAfterSelected,
            //     nextFourYears
            //   };
            // };
            // const getNextFourYears = (lastMonthYear: string): string[] => {
            //   const startYear = parseInt(lastMonthYear);
            //   const years: string[] = [];
            //   // Generate 4 years: last month year + next 3 years
            //   for (let i = 0; i < 4; i++) {
            //     years.push((startYear + i).toString());
            //   }
            //   return years;
            // };
            // Enhanced OBL breakdown function with multiple fallbacks
            const getOBLBreakdownForExport = (collection) => {
                const defaultBreakdown = {
                    normalDue: 0,
                    normalOverdue: 0,
                    retentionDue: 0,
                    retentionOverdue: 0,
                    gstTds: 0,
                    wht: 0
                };
                try {
                    // Use the exact field names from your database
                    const breakdown = {
                        normalDue: getNumericFieldValue(collection, 'NormalDue'),
                        normalOverdue: getNumericFieldValue(collection, 'NormalOverdue'),
                        retentionDue: getNumericFieldValue(collection, 'RetentionDue'),
                        retentionOverdue: getNumericFieldValue(collection, 'RetentionOverdue'),
                        gstTds: getNumericFieldValue(collection, 'GstTds'), // This is one field in your DB
                        wht: getNumericFieldValue(collection, 'WHT')
                    };
                    return breakdown;
                }
                catch (error) {
                    console.error("Error in getOBLBreakdownForExport:", error);
                    return defaultBreakdown;
                }
            };
            const getNumericFieldValue = (collection, fieldName) => {
                const value = collection[fieldName];
                if (value === undefined || value === null || value === '') {
                    return 0;
                }
                // Handle different value types
                if (typeof value === 'number') {
                    return value;
                }
                if (typeof value === 'string') {
                    const num = parseFloat(value);
                    return isNaN(num) ? 0 : num;
                }
                return 0;
            };
            // At the top of your script, initialize pageSize from dropdown
            // Replace the existing initialization:
            // let pageSize = 10;
            // With:
            // COMPLETE REPLACEMENT FOR setupPaginationEventListeners
            const setupPaginationEventListeners = () => {
                if (!pageSizeSelect) {
                    return;
                }
                pageSizeSelect.addEventListener("change", (e) => {
                    const newSize = parseInt(e.target.value, 10);
                    changePageSize(newSize);
                });
            };
            // Separate handler for page size change
            // const handlePageSizeChange = function(e: Event) {
            //   const target = e.target as HTMLSelectElement;
            //   const newSize = parseInt(target.value);
            //   console.log(`📊 Page size changed to: ${newSize}`);
            //   changePageSize(newSize);
            // };
            // ✅ Separate handler for document-level event listener
            // const handleDocumentPageSizeChange = function(e: Event) {
            //   const target = e.target as HTMLElement;
            //   if (target && target.id === 'pageSize') {
            //     const newSize = parseInt((target as HTMLSelectElement).value);
            //     console.log(`📊 Page size changed (document): ${newSize}`);
            //     changePageSize(newSize);
            //   }
            // };
            // COMPLETE REPLACEMENT FOR changePageSize
            const changePageSize = (newSize) => {
                const validPageSizes = [10, 25, 50, 100];
                if (!validPageSizes.includes(newSize)) {
                    newSize = 10;
                }
                pageSize = newSize;
                currentPage = 1;
                calculateTotalPages();
                updatePaginationControls();
                renderCurrentPage();
                toast("info", `Showing ${pageSize} SEIBIES records per page`);
            };
            // COMPLETE REPLACEMENT FOR renderCurrentPage
            const renderCurrentPage = () => {
                const tableBody = document.querySelector(".collection-table-body");
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
                // ✅ Handle empty state
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
                // ✅ Ensure currentPage is within bounds
                if (currentPage > totalPages) {
                    currentPage = totalPages;
                }
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, sortedData.length);
                const pageData = sortedData.slice(startIndex, endIndex);
                console.log(`📊 Rendering page ${currentPage}: items ${startIndex + 1} to ${endIndex} of ${sortedData.length}`);
                console.log(`📊 Page data length: ${pageData.length}`);
                console.log(`📊 Using pageSize: ${pageSize}`);
                // ✅ Use the existing renderTableBodyOnly function
                renderTableBodyOnly(pageData, displayMonths, currentMonthKey);
                renderTotalsHeader(displayMonths, selectedMonthKey);
            };
            const updatePaginationControls = () => {
                currentPageSpan.textContent = currentPage.toString();
                totalPagesSpan.textContent = totalPages.toString();
                if (filteredData.length === 0) {
                    startItemSpan.textContent = "0";
                    endItemSpan.textContent = "0";
                    totalItemsSpan.textContent = "0";
                }
                else {
                    const startItem = (currentPage - 1) * pageSize + 1;
                    const endItem = Math.min(currentPage * pageSize, filteredData.length);
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
            // COMPLETE REPLACEMENT FOR goToPage
            // Replace the goToPage function
            const goToPage = (page) => {
                if (page < 1 ||
                    page > totalPages ||
                    page === currentPage) {
                    return;
                }
                currentPage = page;
                updatePaginationControls();
                renderCurrentPage();
            };
            // COMPLETE REPLACEMENT FOR initializePagination
            const initializePagination = (data) => {
                filteredData = [...data];
                // Default page size must always be 10
                pageSize = 10;
                currentPage = 1;
                // Keep dropdown synchronized
                if (pageSizeSelect) {
                    pageSizeSelect.value = "10";
                }
                calculateTotalPages();
                updatePaginationControls();
                renderCurrentPage();
            };
            // COMPLETE REPLACEMENT FOR calculateTotalPages
            const calculateTotalPages = () => {
                totalPages = Math.ceil(filteredData.length / pageSize);
                if (totalPages === 0) {
                    totalPages = 1;
                }
            };
            // COMPLETE REPLACEMENT FOR updatePageNumbers
            const updatePageNumbers = () => {
                pageNumbersDiv.innerHTML = "";
                if (filteredData.length === 0) {
                    return;
                }
                const maxVisiblePages = 5;
                let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
                let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
                if (endPage - startPage + 1 < maxVisiblePages) {
                    startPage = Math.max(1, endPage - maxVisiblePages + 1);
                }
                if (startPage > 1) {
                    const dotsBtn = document.createElement("button");
                    dotsBtn.className =
                        "page-number-seibies dots-seibies";
                    dotsBtn.textContent = "...";
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
                for (let i = startPage; i <= endPage; i++) {
                    const pageBtn = document.createElement("button");
                    pageBtn.className =
                        `page-number-seibies ${i === currentPage
                            ? "active-seibies"
                            : ""}`;
                    pageBtn.textContent = i.toString();
                    pageBtn.addEventListener("click", () => goToPage(i));
                    pageNumbersDiv.appendChild(pageBtn);
                }
                if (endPage < totalPages) {
                    const dotsBtn = document.createElement("button");
                    dotsBtn.className =
                        "page-number-seibies dots-seibies";
                    dotsBtn.textContent = "...";
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
            };
            // Helper function to add page button
            // const addPageButton = (page: number): void => {
            //   const pageBtn = document.createElement('button');
            //   pageBtn.className = `page-number ${page === currentPage ? 'active' : ''}`;
            //   pageBtn.textContent = page.toString();
            //   pageBtn.addEventListener('click', () => goToPage(page));
            //   pageNumbersDiv.appendChild(pageBtn);
            // };
            // Helper function to add ellipsis
            // const addEllipsis = (): void => {
            //   const ellipsis = document.createElement('span');
            //   ellipsis.className = 'pagination-ellipsis';
            //   ellipsis.textContent = '...';
            //   ellipsis.style.padding = '8px 6px';
            //   ellipsis.style.color = '#666';
            //   pageNumbersDiv.appendChild(ellipsis);
            // };
            const updatePagination = () => {
                totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
                if (currentPage > totalPages) {
                    currentPage = totalPages;
                }
                if (currentPage < 1) {
                    currentPage = 1;
                }
                updatePaginationControls();
                renderCurrentPage();
            };
            // const handlePageSizeChange = (): void => {
            //   const selectedSize = parseInt(
            //     pageSizeSelect.value,
            //     10
            //   );
            //   if (!Number.isFinite(selectedSize) || selectedSize <= 0) {
            //     return;
            //   }
            //   // Update page size
            //   pageSize = selectedSize;
            //   // Always go back to page 1
            //   currentPage = 1;
            //   // Recalculate total pages
            //   totalPages = Math.max(
            //     1,
            //     Math.ceil(
            //       filteredData.length / pageSize
            //     )
            //   );
            //   // Update pagination text/buttons
            //   updatePaginationControls();
            //   // Re-render current page using the new page size
            //   renderCurrentPage();
            //   toast(
            //     "info",
            //     `Now showing ${pageSize} records per page`
            //   );
            // };
            // Updated sorting function with month index support
            const handleEnhancedSort = (column, direction, subColumn, monthIndex) => {
                // Update current sort state
                currentSortConfig = {
                    column,
                    direction,
                    subColumn,
                    monthIndex
                };
                const dataToSort = filteredData.length > 0 ? filteredData : collections;
                if (dataToSort.length === 0)
                    return;
                const sortedData = [...dataToSort].sort((a, b) => {
                    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r;
                    let valueA = 0;
                    let valueB = 0;
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
                            }
                            else {
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
                            const monthDataA = dataA[monthKeyToUse] || { plan: 0, actual: 0, items: {} };
                            const monthDataB = dataB[monthKeyToUse] || { plan: 0, actual: 0, items: {} };
                            const itemsA = monthDataA.items || {};
                            const itemsB = monthDataB.items || {};
                            if (subColumn) {
                                switch (subColumn) {
                                    case 'normal-previous':
                                        valueA = (((_a = itemsA.Normal) === null || _a === void 0 ? void 0 : _a.plan) || 0) + (((_b = itemsA['Normal Overdue']) === null || _b === void 0 ? void 0 : _b.plan) || 0);
                                        valueB = (((_c = itemsB.Normal) === null || _c === void 0 ? void 0 : _c.plan) || 0) + (((_d = itemsB['Normal Overdue']) === null || _d === void 0 ? void 0 : _d.plan) || 0);
                                        break;
                                    case 'normal-current':
                                        valueA = (((_e = itemsA.Normal) === null || _e === void 0 ? void 0 : _e.actual) || 0) + (((_f = itemsA['Normal Overdue']) === null || _f === void 0 ? void 0 : _f.actual) || 0);
                                        valueB = (((_g = itemsB.Normal) === null || _g === void 0 ? void 0 : _g.actual) || 0) + (((_h = itemsB['Normal Overdue']) === null || _h === void 0 ? void 0 : _h.actual) || 0);
                                        break;
                                    case 'retention-previous':
                                        valueA = (((_j = itemsA.Retention) === null || _j === void 0 ? void 0 : _j.plan) || 0) + (((_k = itemsA['Retention Overdue']) === null || _k === void 0 ? void 0 : _k.plan) || 0);
                                        valueB = (((_l = itemsB.Retention) === null || _l === void 0 ? void 0 : _l.plan) || 0) + (((_m = itemsB['Retention Overdue']) === null || _m === void 0 ? void 0 : _m.plan) || 0);
                                        break;
                                    case 'retention-current':
                                        valueA = (((_o = itemsA.Retention) === null || _o === void 0 ? void 0 : _o.actual) || 0) + (((_p = itemsA['Retention Overdue']) === null || _p === void 0 ? void 0 : _p.actual) || 0);
                                        valueB = (((_q = itemsB.Retention) === null || _q === void 0 ? void 0 : _q.actual) || 0) + (((_r = itemsB['Retention Overdue']) === null || _r === void 0 ? void 0 : _r.actual) || 0);
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
                            }
                            else {
                                valueA = monthDataA.actual || 0;
                                valueB = monthDataB.actual || 0;
                            }
                            break;
                        default:
                            return 0;
                    }
                    if (direction === 'asc') {
                        return valueA - valueB;
                    }
                    else {
                        return valueB - valueA;
                    }
                });
                if (filteredData.length > 0) {
                    filteredData = sortedData;
                }
                else {
                    collections = sortedData;
                }
                currentPage = 1;
                updatePagination();
                const sortDirectionText = direction === 'asc' ? 'Ascending' : 'Descending';
                let columnName = '';
                if (subColumn) {
                    columnName = subColumn.replace('-', ' ').replace(/\b\w/g, l => l.toUpperCase());
                }
                else if (column === 'receivable') {
                    columnName = 'Receivable';
                }
                else if (column === 'collection') {
                    columnName = 'Collection';
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
            // Updated sort button listener
            const attachSortButtonListeners = () => {
                document.querySelectorAll('.sort-btn[data-sort]').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const button = e.currentTarget;
                        const sortColumn = button.dataset.sort;
                        const direction = button.dataset.direction;
                        const subColumn = button.dataset.subcolumn;
                        const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;
                        if (sortColumn && direction) {
                            handleEnhancedSort(sortColumn, direction, subColumn, monthIndex);
                        }
                    });
                });
            };
            // ========== FIXED RENDER TABLE WITH MONTH COLUMNS - SHOW REVENUE ACTUAL ==========
            const renderTableBodyOnly = (data, displayMonths, currentMonthKey) => {
                const tableBody = document.querySelector(".collection-table-body");
                if (!tableBody)
                    return;
                if (!data || data.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="21" style="text-align:center;padding:20px;border:1px solid #ddd;">No records found</td></tr>`;
                    return;
                }
                let tableBodyHTML = '';
                data.forEach((collection) => {
                    var _a, _b;
                    const projectTitle = ((_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                    const proj = projects.find(p => p.Title === projectTitle);
                    const collectionData = safeJsonParse(collection.CollectionData, {});
                    const totalReceivable = collection.ReceivableValue || 0;
                    const formattedOBL = Math.round(totalReceivable).toLocaleString('en-IN');
                    const isFrozen = collection.Freeze || false;
                    const rowStyle = isFrozen ? 'background-color: #fff8e1;' : '';
                    const canUserEdit = authCanEditProject(proj);
                    const canUserFreeze = canFreeze();
                    const isCheckboxDisabled = !canUserFreeze;
                    const projectRef = getSafeString(proj === null || proj === void 0 ? void 0 : proj.ProjectReference);
                    const projectRefLink = projectRef !== "-"
                        ? (canUserEdit
                            ? `<a href="#" class="edit-trigger" data-id="${collection.Id}" data-project="${(_b = collection.Project) === null || _b === void 0 ? void 0 : _b.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectRef}</a>`
                            : `<span style="color:#6c757d; cursor:default;">${projectRef}</span>`)
                        : '<span style="color:#999;">-</span>';
                    const projectNameLink = canUserEdit
                        ? `<a href="#" class="edit-trigger" data-id="${collection.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectTitle}</a>`
                        : `<span style="color:#6c757d; cursor:default;">${projectTitle}</span>`;
                    const freezeBadge = isFrozen ? ' 🔒' : '';
                    let rowHTML = `
      <tr style="${rowStyle}" data-collection-id="${collection.Id}"> 
        <td class="freeze-col" style="text-align:center; vertical-align:middle; min-width:40px; max-width:40px; width:40px;">
          <input type="checkbox" class="freeze-checkbox" 
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
        <td class="project-ref-col" style="min-width:150px; max-width:150px; width:150px; word-wrap:break-word; white-space:normal;">${projectRefLink}</td>
        <td class="project-col" style="min-width:200px; max-width:200px; width:200px; word-wrap:break-word; white-space:normal;">${projectNameLink}${freezeBadge}</td>
        <td class="pm-col" style="min-width:120px; max-width:120px; width:120px; word-wrap:break-word; white-space:normal;">${proj ? getSafeString(proj.ProjectManager) : "-"}</td>
        <td class="gh-col" style="min-width:120px; max-width:120px; width:120px; word-wrap:break-word; white-space:normal;">${proj ? getSafeString(proj.GroupHead) : "-"}</td>
        <td class="customer-col" style="min-width:150px; max-width:150px; width:150px; word-wrap:break-word; white-space:normal;">${getSafeString(proj === null || proj === void 0 ? void 0 : proj.CustomerName)}</td>
        <td class="obl-col" style="min-width:140px; max-width:140px; width:140px; text-align:right; word-wrap:break-word; white-space:normal;">${formattedOBL}</td>
    `;
                    // Monthly data columns - display ONLY numeric data, no category names
                    displayMonths.forEach((monthKey) => {
                        var _a, _b, _c, _d, _e, _f, _g, _h;
                        const monthData = collectionData[monthKey] || { plan: 0, actual: 0, forecast: 0, items: {} };
                        const isCurrentMonth = monthKey === currentMonthKey;
                        const items = monthData.items || {};
                        // Previous columns
                        const prevNormal = (((_a = items.Normal) === null || _a === void 0 ? void 0 : _a.plan) || 0) + (((_b = items['Normal Overdue']) === null || _b === void 0 ? void 0 : _b.plan) || 0);
                        const prevRetention = (((_c = items.Retention) === null || _c === void 0 ? void 0 : _c.plan) || 0) + (((_d = items['Retention Overdue']) === null || _d === void 0 ? void 0 : _d.plan) || 0);
                        const prevTotal = prevNormal + prevRetention;
                        // Current columns  
                        const currentNormal = (((_e = items.Normal) === null || _e === void 0 ? void 0 : _e.actual) || 0) + (((_f = items['Normal Overdue']) === null || _f === void 0 ? void 0 : _f.actual) || 0);
                        const currentRetention = (((_g = items.Retention) === null || _g === void 0 ? void 0 : _g.actual) || 0) + (((_h = items['Retention Overdue']) === null || _h === void 0 ? void 0 : _h.actual) || 0);
                        const currentTotal = currentNormal + currentRetention;
                        // Add Previous columns
                        rowHTML += `
        <td class="month-scroll-col" style="min-width:55px; text-align:right;">${prevNormal.toLocaleString('en-IN')}</td>
        <td class="month-scroll-col" style="min-width:55px; text-align:right;">${prevRetention.toLocaleString('en-IN')}</td>
        <td class="month-scroll-col" style="min-width:55px; text-align:right; font-weight:bold;">${prevTotal.toLocaleString('en-IN')}</td>
      `;
                        // Add Current columns
                        rowHTML += `
        <td class="month-scroll-col" style="min-width:55px; text-align:right;">${currentNormal.toLocaleString('en-IN')}</td>
        <td class="month-scroll-col" style="min-width:55px; text-align:right;">${currentRetention.toLocaleString('en-IN')}</td>
        <td class="month-scroll-col" style="min-width:55px; text-align:right; font-weight:bold;">${currentTotal.toLocaleString('en-IN')}</td>
      `;
                        // Current month only: Actual, Further, Forecast
                        if (isCurrentMonth) {
                            let revenueActualData = {};
                            try {
                                if (collection._revenueActual && typeof collection._revenueActual === 'object') {
                                    revenueActualData = collection._revenueActual;
                                }
                                else if (typeof collection.RevenueActual === 'string') {
                                    const parsed = JSON.parse(collection.RevenueActual);
                                    if (typeof parsed === 'object' && parsed !== null) {
                                        revenueActualData = parsed;
                                    }
                                }
                                else if (typeof collection.RevenueActual === 'object') {
                                    revenueActualData = collection.RevenueActual || {};
                                }
                            }
                            catch (e) {
                                revenueActualData = {};
                            }
                            const actualValue = Number(revenueActualData[monthKey]) || 0;
                            const furtherValue = monthData.forecast || 0;
                            const forecastValue = currentTotal + furtherValue;
                            rowHTML += `
          <td class="month-scroll-col" style="min-width:55px; text-align:right; background-color: #e8f5e9; font-weight: bold; color: #1a8a34;">${actualValue.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:55px; text-align:right;">${furtherValue.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col total-cell" style="min-width:55px; text-align:right; font-weight: bold; background-color: #f0f8ff; color: #0c5460;">${forecastValue.toLocaleString('en-IN')}</td>
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
            // Replace the renderTableHeader function with this:
            const renderTableHeader = () => {
                const monthColumnsHeader = document.querySelector("#monthColumnsHeader");
                if (!monthColumnsHeader)
                    return;
                // ALWAYS render header - don't use _headerRendered flag
                // The flag was preventing header re-render after filters
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                const headerHTML = buildHeaderHTML(displayMonths, currentMonthKey);
                monthColumnsHeader.innerHTML = headerHTML;
                attachSortButtonListeners();
                // Just set the flag to true after rendering
                // _headerRendered = true;
            };
            const buildHeaderHTML = (displayMonths, currentMonthKey) => {
                // Build the main header HTML
                let html = `
    <tr>  
      <th rowspan="3" class="freeze-col" style="text-align:center; min-width:40px; max-width:40px; width:40px; white-space:nowrap;">#</th>
      <th rowspan="3" class="project-ref-col" style="min-width:150px; max-width:150px; width:150px; word-wrap:break-word; white-space:nowrap;">Project Reference</th>
      <th rowspan="3" class="project-col" style="min-width:200px; max-width:200px; width:200px; word-wrap:break-word; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:4px;">
          <span class="header-label">Project</span>
          <div class="sort-buttons-container" style="display:flex; gap:2px;">
            <button class="sort-btn" data-sort="project" data-direction="asc" title="Sort Ascending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="project" data-direction="desc" title="Sort Descending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
      <th rowspan="3" class="pm-col" style="min-width:120px; max-width:120px; width:120px; word-wrap:break-word; white-space:nowrap;">Project Manager</th>
      <th rowspan="3" class="gh-col" style="min-width:120px; max-width:120px; width:120px; word-wrap:break-word; white-space:nowrap;">Group Head</th>
      <th rowspan="3" class="customer-col" style="min-width:150px; max-width:150px; width:150px; word-wrap:break-word; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:4px;">
          <span class="header-label">Customer</span>
          <div class="sort-buttons-container" style="display:flex; gap:2px;">
            <button class="sort-btn" data-sort="customer" data-direction="asc" title="Sort Ascending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="customer" data-direction="desc" title="Sort Descending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
      <th rowspan="3" class="obl-col" style="min-width:140px; max-width:140px; width:140px; word-wrap:break-word; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:4px;">
          <span class="header-label">Receivables (incl TDS)</span>
          <div class="sort-buttons-container" style="display:flex; gap:2px;">
            <button class="sort-btn" data-sort="receivable" data-direction="asc" title="Sort Ascending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="receivable" data-direction="desc" title="Sort Descending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
  `;
                // Add month columns
                displayMonths.forEach((month, monthIndex) => {
                    const isCurrentMonth = month === currentMonthKey;
                    const monthDate = new Date(month + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const colspan = isCurrentMonth ? 9 : 6;
                    const isCurrentDisplay = isCurrentMonth ? ' (Current)' : '';
                    html += `
      <th colspan="${colspan}" class="month-scroll-col month-column-header ${isCurrentMonth ? 'current-month-header' : ''}" style="min-width:${isCurrentMonth ? '300px' : '200px'}; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:4px;">
          <span class="header-label">${monthLabel}${isCurrentDisplay}</span>
          <div class="sort-buttons-container" style="display:flex; gap:2px;">
            <button class="sort-btn" data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort Ascending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="month" data-direction="desc" data-month-index="${monthIndex}" title="Sort Descending" style="font-size:10px; padding:0 3px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
    `;
                });
                html += `</tr>
    
    <!-- Row 2: Previous (colspan 3) | Current -->
    <tr>
  `;
                displayMonths.forEach((month) => {
                    const isCurrentMonth = month === currentMonthKey;
                    html += `
      <th colspan="3" class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; text-align:center; background-color: ${isCurrentMonth ? '#e3f2fd' : '#f8f9fa'}; white-space:nowrap;">
        Prev
      </th>
    `;
                    const currentColspan = isCurrentMonth ? 6 : 3;
                    html += `
      <th colspan="${currentColspan}" class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; text-align:center; background-color: ${isCurrentMonth ? '#e3f2fd' : '#f8f9fa'}; white-space:nowrap;">
        Curr
      </th>
    `;
                });
                html += `</tr>
    
    <!-- Row 3: Actual column names -->
    <tr>
  `;
                displayMonths.forEach((month, monthIndex) => {
                    const isCurrentMonth = month === currentMonthKey;
                    // Previous columns
                    html += `
      <th class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
          <span style="font-size:11px;">Normal</span>
          <div class="sort-buttons-container" style="display:flex; gap:1px;">
            <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="prev-normal" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="prev-normal" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
      <th class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
          <span style="font-size:11px;">Retention</span>
          <div class="sort-buttons-container" style="display:flex; gap:1px;">
            <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="prev-retention" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="prev-retention" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
      <th class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
          <span style="font-size:11px;">Total</span>
          <div class="sort-buttons-container" style="display:flex; gap:1px;">
            <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="prev-total" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="prev-total" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
    `;
                    // Current columns
                    html += `
      <th class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
          <span style="font-size:11px;">Normal</span>
          <div class="sort-buttons-container" style="display:flex; gap:1px;">
            <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="normal" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="normal" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
      <th class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
          <span style="font-size:11px;">Retention</span>
          <div class="sort-buttons-container" style="display:flex; gap:1px;">
            <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="retention" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="retention" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
      <th class="month-scroll-col month-subheader ${isCurrentMonth ? 'current-month-header-table' : ''}" style="min-width:55px; white-space:nowrap;">
        <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
          <span style="font-size:11px;">Total</span>
          <div class="sort-buttons-container" style="display:flex; gap:1px;">
            <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="total" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
            <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="total" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
          </div>
        </div>
      </th>
    `;
                    // Current month only: Actual, Further, Forecast
                    if (isCurrentMonth) {
                        html += `
        <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:55px; white-space:nowrap;">
          <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
            <span style="font-size:11px;">Actual</span>
            <div class="sort-buttons-container" style="display:flex; gap:1px;">
              <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="actual" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
              <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="actual" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
            </div>
          </div>
        </th>
        <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:55px; white-space:nowrap;">
          <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
            <span style="font-size:11px;">Further</span>
            <div class="sort-buttons-container" style="display:flex; gap:1px;">
              <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="forecast" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
              <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="forecast" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
            </div>
          </div>
        </th>
        <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:55px; white-space:nowrap;">
          <div class="sort-header-container" style="display:flex; align-items:center; gap:2px; justify-content:center;">
            <span style="font-size:11px;">Forecast</span>
            <div class="sort-buttons-container" style="display:flex; gap:1px;">
              <button class="sort-btn" data-sort="month" data-direction="asc" data-subcolumn="forecast-total" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▲</button>
              <button class="sort-btn" data-sort="month" data-direction="desc" data-subcolumn="forecast-total" data-month-index="${monthIndex}" style="font-size:8px; padding:0 2px; border:none; background:transparent; cursor:pointer;">▼</button>
            </div>
          </div>
        </th>
      `;
                    }
                });
                html += `</tr>`;
                return html;
            };
            // ========== FIXED RENDER TOTALS HEADER ==========
            const renderTotalsHeader = (displayMonths, selectedMonthKey) => {
                const totalsHeader = document.querySelector("#totalsHeader");
                if (!totalsHeader) {
                    console.error("Totals header element not found");
                    return;
                }
                const allData = collections;
                // IMPORTANT:
                // filteredData can legitimately be empty when a filter returns no records.
                // Do NOT fall back to allData in that case.
                const hasActiveFilters = filterYear.value !== "" ||
                    filterMonth.value !== "" ||
                    filterProject.value !== "" ||
                    filterCustomer.value !== "" ||
                    filterProjectManager.value !== "" ||
                    filterGroupHead.value !== "" ||
                    filterProjectReference.value !== "";
                const filtered = hasActiveFilters
                    ? filteredData
                    : collections;
                const allTotals = calculateTotals(allData, displayMonths, selectedMonthKey);
                const filteredTotals = calculateTotals(filtered, displayMonths, selectedMonthKey);
                const buildTotalsRow = (label, totals, cssClass, count) => {
                    let rowHTML = `
      <tr class="${cssClass}">
        <td
          class="freeze-col"
          style="
            min-width:40px;
            max-width:40px;
            width:40px;
            text-align:center;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="project-ref-col"
          style="
            min-width:150px;
            max-width:150px;
            width:150px;
            box-sizing:border-box;
          "
        >
          <strong>${label}</strong>
        </td>

        <td
          class="project-col"
          style="
            min-width:200px;
            max-width:200px;
            width:200px;
            box-sizing:border-box;
          "
        >
          <span style="font-size:11px; color:#6c757d;">
            ${label === "All Projects"
                        ? `All Projects (${count})`
                        : `Filtered (${count})`}
          </span>
        </td>

        <td
          class="pm-col"
          style="
            min-width:120px;
            max-width:120px;
            width:120px;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="gh-col"
          style="
            min-width:120px;
            max-width:120px;
            width:120px;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="customer-col"
          style="
            min-width:150px;
            max-width:150px;
            width:150px;
            box-sizing:border-box;
          "
        ></td>

        <td
          class="obl-col"
          style="
            min-width:140px;
            max-width:140px;
            width:140px;
            text-align:right;
            font-weight:bold;
            color:#0078d4;
            box-sizing:border-box;
          "
        >
          ${Math.round(totals.receivable || 0).toLocaleString("en-IN")}
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
          class="month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:right;
            font-weight:bold;
            color:#0056b3;
          "
        >
          ${Math.round(t.normalPrevious || 0).toLocaleString("en-IN")}
        </td>

        <td
          class="month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:right;
            font-weight:bold;
            color:#0056b3;
          "
        >
          ${Math.round(t.retentionPrevious || 0).toLocaleString("en-IN")}
        </td>

        <td
          class="month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:right;
            font-weight:bold;
            color:#0078d4;
          "
        >
          ${Math.round(t.totalPrevious || 0).toLocaleString("en-IN")}
        </td>
      `;
                        // Current
                        rowHTML += `
        <td
          class="month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:right;
            font-weight:bold;
            color:#28a745;
          "
        >
          ${Math.round(t.normalCurrent || 0).toLocaleString("en-IN")}
        </td>

        <td
          class="month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:right;
            font-weight:bold;
            color:#28a745;
          "
        >
          ${Math.round(t.retentionCurrent || 0).toLocaleString("en-IN")}
        </td>

        <td
          class="month-scroll-col"
          style="
            min-width:55px;
            box-sizing:border-box;
            text-align:right;
            font-weight:bold;
            color:#1a8a34;
          "
        >
          ${Math.round(t.totalCurrent || 0).toLocaleString("en-IN")}
        </td>
      `;
                        // Current month: Actual / Further / Forecast
                        if (isCurrentMonth) {
                            rowHTML += `
          <td
            class="month-scroll-col"
            style="
              min-width:55px;
              box-sizing:border-box;
              text-align:right;
              font-weight:bold;
              color:#856404;
              background-color:#fff3cd;
            "
          >
            ${Math.round(t.actual || 0).toLocaleString("en-IN")}
          </td>

          <td
            class="month-scroll-col"
            style="
              min-width:55px;
              box-sizing:border-box;
              text-align:right;
              font-weight:bold;
              color:#856404;
            "
          >
            ${Math.round(t.further || 0).toLocaleString("en-IN")}
          </td>

          <td
            class="month-scroll-col"
            style="
              min-width:55px;
              box-sizing:border-box;
              text-align:right;
              font-weight:bold;
              color:#0c5460;
              background-color:#d1ecf1;
            "
          >
            ${Math.round(t.forecast || 0).toLocaleString("en-IN")}
          </td>
        `;
                        }
                    });
                    rowHTML += `</tr>`;
                    return rowHTML;
                };
                // IMPORTANT:
                // 7 fixed columns
                // + 6 columns for each displayed month
                // + 3 additional current-month columns
                const totalColumns = 7 +
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

    ${buildTotalsRow("All Projects", allTotals, "all-projects-total", allData.length)}

    ${buildTotalsRow("Filtered", filteredTotals, "filtered-total", filtered.length)}
  `;
            };
            const calculateTotals = (data, displayMonths, selectedMonthKey) => {
                const dataToCalculate = data;
                const totals = {
                    receivable: 0,
                    actualCollection: 0,
                    monthly: {}
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
                    if (selectedMonthKey) {
                        let revenueData = {};
                        try {
                            if (typeof collection.RevenueActual === 'string') {
                                revenueData = JSON.parse(collection.RevenueActual) || {};
                            }
                            else if (typeof collection.RevenueActual === 'object') {
                                revenueData = collection.RevenueActual || {};
                            }
                        }
                        catch (e) {
                            revenueData = {};
                        }
                        totals.actualCollection += Math.round(Number(revenueData[selectedMonthKey]) || 0);
                    }
                    const collectionData = safeJsonParse(collection.CollectionData, {});
                    displayMonths.forEach(monthKey => {
                        var _a, _b, _c, _d, _e, _f, _g, _h;
                        // Parse RevenueActual data INSIDE the loop where monthKey is available
                        let revenueActualData = {};
                        try {
                            if (typeof collection.RevenueActual === 'string') {
                                revenueActualData = JSON.parse(collection.RevenueActual) || {};
                            }
                            else if (typeof collection.RevenueActual === 'object') {
                                revenueActualData = collection.RevenueActual || {};
                            }
                        }
                        catch (e) {
                            revenueActualData = {};
                        }
                        const m = collectionData[monthKey] || { plan: 0, actual: 0, items: {} };
                        const items = m.items || {};
                        const normalPrevious = Math.round((((_a = items.Normal) === null || _a === void 0 ? void 0 : _a.plan) || 0) + (((_b = items['Normal Overdue']) === null || _b === void 0 ? void 0 : _b.plan) || 0));
                        const retentionPrevious = Math.round((((_c = items.Retention) === null || _c === void 0 ? void 0 : _c.plan) || 0) + (((_d = items['Retention Overdue']) === null || _d === void 0 ? void 0 : _d.plan) || 0));
                        const totalPrevious = Math.round(m.plan || 0);
                        const normalCurrent = Math.round((((_e = items.Normal) === null || _e === void 0 ? void 0 : _e.actual) || 0) + (((_f = items['Normal Overdue']) === null || _f === void 0 ? void 0 : _f.actual) || 0));
                        const retentionCurrent = Math.round((((_g = items.Retention) === null || _g === void 0 ? void 0 : _g.actual) || 0) + (((_h = items['Retention Overdue']) === null || _h === void 0 ? void 0 : _h.actual) || 0));
                        const totalCurrent = Math.round(m.actual || 0);
                        const actual = Math.round(Number(revenueActualData[monthKey]) || 0);
                        const further = Math.round(m.forecast || 0);
                        const forecast = Math.round(totalCurrent + further);
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
            const showForm = (isEditMode = false) => {
                formSection.style.display = "block";
                tableSection.style.display = "none";
                showFormBtn.style.display = "none";
                if (isEditMode) {
                    formSection.classList.add('edit-mode');
                    projectSelect.disabled = true;
                    formTitle.textContent = "Edit Collection";
                }
                else {
                    formSection.classList.remove('edit-mode');
                    projectSelect.disabled = false;
                    formTitle.textContent = "Add Collection";
                }
            };
            const hideForm = () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                showFormBtn.style.display = "inline-block";
                formSection.classList.remove('edit-mode');
                projectSelect.disabled = false;
            };
            const resetForm = () => {
                editingId = null;
                form.reset();
                projectSelect.value = "";
                // IMPORTANT:
                // Do NOT clear projectDetails.innerHTML.
                // The project detail HTML is static and must remain in the DOM.
                // Only reset its displayed values.
                const projectNameDisplay = document.querySelector("#projectNameDisplay");
                const projectRef = document.querySelector("#projectRef");
                const profitCenter = document.querySelector("#profitCenter");
                const projectManager = document.querySelector("#projectManager");
                const groupHead = document.querySelector("#groupHead");
                const customer = document.querySelector("#customer");
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
                const remarksInternal = document.querySelector("#remarksInternal");
                const remarksSprintCell = document.querySelector("#remarksSprintCell");
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
                    const element = document.getElementById(id);
                    if (element) {
                        element.value = "";
                        element.className =
                            element.className.replace("input-with-value", "input-empty");
                        if (!element.className.includes("input-empty")) {
                            element.classList.add("input-empty");
                        }
                        element.placeholder = "0";
                    }
                });
                // Hide collection sections
                toggleCollectionSections(false);
                const oblContainer = document.querySelector(".obl-section-container");
                if (oblContainer) {
                    oblContainer.className =
                        "obl-section-container hidden-section";
                }
                oblLastMonthInput.removeAttribute("readonly");
                oblLastMonthInput.placeholder =
                    "Enter OBL";
                setNumericValue(oblLastMonthInput, 0);
                const fieldNote = oblContainer === null || oblContainer === void 0 ? void 0 : oblContainer.querySelector(".field-note");
                if (fieldNote) {
                    fieldNote.textContent =
                        "For new projects: Enter initial Receivables value";
                }
                oblValidationMessage.textContent =
                    "✅ All values are within Receivables limit";
                oblValidationMessage.className =
                    "validation-message";
                updateOBLValidationDisplay();
            };
            // ========== VALIDATION FUNCTIONS ==========
            const calculateTotalCurrentValue = () => {
                let total = 0;
                // Include monthly current values from ALL editable months (detailed + long-term)
                // These are months with index 36-53 (detailed and long-term)
                const allEditableMonths = allMonths.slice(36); // From index 36 onwards
                allEditableMonths.forEach(monthKey => {
                    const currentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]:not([readonly])`);
                    if (currentInput) {
                        total += getNumericValue(currentInput);
                    }
                });
                // Include yearly current values
                const yearlyCurrentInputs = document.querySelectorAll(`input.yearly-current-input:not([readonly])`);
                yearlyCurrentInputs.forEach(input => {
                    total += getNumericValue(input);
                });
                return total;
            };
            const validateOBLLimit = (changedInput) => {
                const oblValue = getNumericValue(oblLastMonthInput);
                const totalCurrent = calculateTotalCurrentValue();
                // Update displays
                totalCurrentValueSpan.textContent = totalCurrent.toLocaleString();
                const remaining = oblValue - totalCurrent;
                remainingValueSpan.textContent = remaining.toLocaleString();
                // Update progress bar
                if (oblValue > 0) {
                    const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
                    oblLineFill.style.width = `${percentage}%`;
                    oblLineFill.className = 'line-fill';
                    if (percentage > 90) {
                        oblLineFill.classList.add('danger');
                    }
                    else if (percentage > 75) {
                        oblLineFill.classList.add('warning');
                    }
                }
                else {
                    oblLineFill.style.width = `0%`;
                    oblLineFill.className = 'line-fill';
                }
                // Check if OBL is zero but values exist
                if (oblValue === 0 && totalCurrent > 0) {
                    showOBLErrorToast('⚠️ Receivables is zero but liquidation values exist. Please set Receivables value.', true);
                    blockAllCurrentInputs();
                    oblValidationMessage.textContent = `❌ Receivables is zero but liquidation values exist (${totalCurrent.toLocaleString()})`;
                    oblValidationMessage.className = 'validation-message danger';
                    return false;
                }
                // Check if total exceeds OBL (Liquidation > Receivables)
                if (totalCurrent > oblValue) {
                    const exceedAmount = totalCurrent - oblValue;
                    const lastToast = window._lastOBLToast;
                    const now = Date.now();
                    if (!lastToast || now - lastToast > 3000) {
                        toast("error", `❌ Liquidation (${totalCurrent.toLocaleString()}) exceeds Receivables (${oblValue.toLocaleString()}) by ${exceedAmount.toLocaleString()}! Reduce liquidation values to continue.`);
                        window._lastOBLToast = now;
                    }
                    oblValidationMessage.textContent = `❌ LIQUIDATION EXCEEDS RECEIVABLES: ${totalCurrent.toLocaleString()} > ${oblValue.toLocaleString()} (Excess: ${exceedAmount.toLocaleString()})`;
                    oblValidationMessage.className = 'validation-message danger';
                    // Block ALL inputs when OBL is exceeded
                    blockAllCurrentInputs();
                    return false;
                }
                // If within OBL limit, unblock all inputs
                unblockAllCurrentInputs();
                // Update status message based on relationship
                if (oblValue > 0 && totalCurrent === 0) {
                    // No liquidation yet
                    oblValidationMessage.textContent = `ℹ️ No liquidation recorded yet. Receivables: ${oblValue.toLocaleString()} - Awaiting liquidation`;
                    oblValidationMessage.className = 'validation-message info';
                }
                else if (totalCurrent === oblValue && oblValue > 0) {
                    // Fully liquidated
                    const lastToast = window._lastOBLToast;
                    const now = Date.now();
                    if (!lastToast || now - lastToast > 3000) {
                        toast("success", `✅ Receivables fully liquidated: ${totalCurrent.toLocaleString()} = ${oblValue.toLocaleString()}`);
                        window._lastOBLToast = now;
                    }
                    oblValidationMessage.textContent = `✅ RECEIVABLES FULLY LIQUIDATED: ${totalCurrent.toLocaleString()} = ${oblValue.toLocaleString()}`;
                    oblValidationMessage.className = 'validation-message success';
                }
                else if (totalCurrent < oblValue && totalCurrent > 0) {
                    // Partial liquidation - Receivables not fully liquidated
                    const percentage = (totalCurrent / oblValue) * 100;
                    const remainingBalance = oblValue - totalCurrent;
                    if (percentage > 90) {
                        const lastToast = window._lastOBLToast;
                        const now = Date.now();
                        if (!lastToast || now - lastToast > 3000) {
                            toast("warning", `⚠️ Receivables nearly fully liquidated! ${percentage.toFixed(1)}% used (${remainingBalance.toLocaleString()} remaining)`);
                            window._lastOBLToast = now;
                        }
                        oblValidationMessage.textContent = `⚠️ RECEIVABLES NEARLY FULLY LIQUIDATED: ${percentage.toFixed(1)}% | ${totalCurrent.toLocaleString()} / ${oblValue.toLocaleString()} (${remainingBalance.toLocaleString()} remaining)`;
                        oblValidationMessage.className = 'validation-message warning';
                    }
                    else if (percentage > 75) {
                        const lastToast = window._lastOBLToast;
                        const now = Date.now();
                        if (!lastToast || now - lastToast > 5000) {
                            toast("info", `ℹ️ Receivables ${percentage.toFixed(1)}% liquidated (${remainingBalance.toLocaleString()} remaining)`);
                            window._lastOBLToast = now;
                        }
                        oblValidationMessage.textContent = `ℹ️ RECEIVABLES PARTIALLY LIQUIDATED: ${percentage.toFixed(1)}% | ${totalCurrent.toLocaleString()} / ${oblValue.toLocaleString()} (${remainingBalance.toLocaleString()} remaining)`;
                        oblValidationMessage.className = 'validation-message info';
                    }
                    else {
                        // Less than 75% liquidated
                        const lastToast = window._lastOBLToast;
                        const now = Date.now();
                        if (!lastToast || now - lastToast > 5000) {
                            toast("info", `ℹ️ Receivables ${percentage.toFixed(1)}% liquidated (${remainingBalance.toLocaleString()} remaining)`);
                            window._lastOBLToast = now;
                        }
                        oblValidationMessage.textContent = `ℹ️ RECEIVABLES PARTIALLY LIQUIDATED: ${percentage.toFixed(1)}% | ${totalCurrent.toLocaleString()} / ${oblValue.toLocaleString()} (${remainingBalance.toLocaleString()} remaining)`;
                        oblValidationMessage.className = 'validation-message info';
                    }
                }
                else if (oblValue === 0 && totalCurrent === 0) {
                    // No Receivables set
                    oblValidationMessage.textContent = `ℹ️ No Receivables value set. Please set Receivables to track liquidation.`;
                    oblValidationMessage.className = 'validation-message info';
                }
                return true;
            };
            const blockAllCurrentInputs = () => {
                // Block all editable current inputs (both monthly and yearly)
                const currentInputs = document.querySelectorAll('input.main-current-input:not([readonly]), input.yearly-current-input:not([readonly])');
                currentInputs.forEach(input => {
                    input.setAttribute('readonly', 'true');
                    input.style.backgroundColor = '#f8f9fa';
                    input.style.cursor = 'not-allowed';
                    input.style.opacity = '0.7';
                });
            };
            const unblockAllCurrentInputs = () => {
                // Unblock all editable current inputs
                const currentInputs = document.querySelectorAll('input.main-current-input[readonly], input.yearly-current-input[readonly]');
                currentInputs.forEach(input => {
                    input.removeAttribute('readonly');
                    input.style.backgroundColor = '';
                    input.style.cursor = '';
                    input.style.opacity = '1';
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
                // Check if adding this increase would exceed OBL
                return (currentTotal + increaseAmount) <= oblValue;
            };
            const validateSingleMonthCategories = (monthKey, changedInput) => {
                const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                const categoryCurrentInputs = document.querySelectorAll(`input.category-current-input[data-month="${monthKey}"]`);
                if (!mainCurrentInput)
                    return true;
                const mainCurrentValue = getNumericValue(mainCurrentInput);
                let categoriesTotal = 0;
                categoryCurrentInputs.forEach(input => {
                    categoriesTotal += getNumericValue(input);
                });
                // Find or create validation display - SINGLE one per month
                let validationDisplay = document.querySelector(`.month-total-validation[data-month="${monthKey}"]`);
                if (!validationDisplay) {
                    validationDisplay = document.createElement('div');
                    validationDisplay.className = 'month-total-validation';
                    validationDisplay.setAttribute('data-month', monthKey);
                    validationDisplay.style.marginTop = '8px';
                    validationDisplay.style.fontSize = '11px';
                    validationDisplay.style.textAlign = 'center';
                    const monthBlock = mainCurrentInput.closest('.month-box, .detailed-month-box, .simple-month');
                    if (monthBlock) {
                        // Remove any existing validation elements
                        const existingValidations = monthBlock.querySelectorAll('.month-total-validation');
                        existingValidations.forEach(el => el.remove());
                        // Remove any old validation messages
                        const oldMessages = monthBlock.querySelectorAll('.validation-message-text');
                        oldMessages.forEach(el => el.remove());
                        monthBlock.appendChild(validationDisplay);
                    }
                }
                // Clear any previous duplicate validation messages
                const monthBlock = mainCurrentInput.closest('.month-box, .detailed-month-box, .simple-month');
                if (monthBlock) {
                    const oldMessages = monthBlock.querySelectorAll('.validation-message-text');
                    oldMessages.forEach(el => el.remove());
                }
                // Single validation message - Simplified, just show totals match or warning
                if (categoriesTotal > mainCurrentValue) {
                    blockIncreasingMonthCategoryInputs(monthKey, changedInput);
                    return false;
                }
                else if (categoriesTotal < mainCurrentValue) {
                    // Keep it simple - just show a warning 
                    unblockAllMonthCategoryInputs(monthKey);
                    return true;
                }
                else {
                    // ✅ Simple confirmation only 
                    unblockAllMonthCategoryInputs(monthKey);
                    return true;
                }
            };
            const blockIncreasingMonthCategoryInputs = (monthKey, allowedInput) => {
                const categoryInputs = document.querySelectorAll(`input.category-current-input[data-month="${monthKey}"]`);
                categoryInputs.forEach(input => {
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
            const unblockAllMonthCategoryInputs = (monthKey) => {
                const categoryInputs = document.querySelectorAll(`input.category-current-input[data-month="${monthKey}"]`);
                categoryInputs.forEach(input => {
                    input.style.backgroundColor = '';
                    input.style.cursor = '';
                    input.removeAttribute('data-block-increase');
                    input.removeAttribute('readonly');
                });
            };
            // Yearly category current inputs
            document.querySelectorAll('.category-current-input[data-year]:not([readonly])').forEach(input => {
                input.addEventListener('input', function () {
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
                        this.className = this.className.replace('input-with-value', 'input-empty');
                    }
                    else {
                        this.className = this.className.replace('input-empty', 'input-with-value');
                    }
                    updateOBLValidationDisplay(this);
                    validateYearlyCategoryTotals(this);
                });
            });
            // Helper function for yearly category validation
            const allowYearCategoryInputIncrease = (input, newValue) => {
                const oldValue = parseInt(input.getAttribute('data-old-value') || '0');
                const isIncreasing = newValue > oldValue;
                if (!isIncreasing) {
                    return true;
                }
                const year = input.dataset.year;
                const yearlyCurrentInput = document.querySelector(`input.yearly-current-input[data-year="${year}"]`);
                if (!yearlyCurrentInput)
                    return true;
                const yearlyCurrentValue = getNumericValue(yearlyCurrentInput);
                const categoryInputs = document.querySelectorAll(`input.category-current-input[data-year="${year}"]`);
                let currentCategoriesTotal = 0;
                categoryInputs.forEach(categoryInput => {
                    if (categoryInput === input) {
                        currentCategoriesTotal += newValue;
                    }
                    else {
                        currentCategoriesTotal += getNumericValue(categoryInput);
                    }
                });
                return currentCategoriesTotal <= yearlyCurrentValue;
            };
            const updateOBLValidationDisplay = (changedInput) => {
                validateOBLLimit(changedInput);
            };
            const validateFormSubmission = () => {
                const oblValue = getNumericValue(oblLastMonthInput);
                const totalCurrent = calculateTotalCurrentValue();
                // OBL Validation - total cannot exceed OBL
                if (oblValue > 0 && totalCurrent > oblValue) {
                    toast("error", `Total current value (${totalCurrent.toLocaleString()}) exceeds OBL (${oblValue.toLocaleString()})! Please reduce values.`);
                    return false;
                }
                // Check if OBL is zero but there are values
                if (oblValue === 0 && totalCurrent > 0) {
                    toast("error", `OBL is set to 0 but there are values (${totalCurrent.toLocaleString()}). Please set OBL value.`);
                    return false;
                }
                // Month Category Validation - categories shouldn't exceed month total
                const invalidMonths = [];
                // Check all detailed months (indices 36-38)
                for (let i = 36; i <= 38 && i < allMonths.length; i++) {
                    const monthKey = allMonths[i];
                    const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                    if (mainCurrentInput) {
                        const mainCurrentValue = getNumericValue(mainCurrentInput);
                        const categoryInputs = document.querySelectorAll(`input.category-current-input[data-month="${monthKey}"]`);
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
                const invalidYears = [];
                const yearlyInputs = document.querySelectorAll('.yearly-current-input');
                yearlyInputs.forEach(input => {
                    const year = input.dataset.year;
                    if (year) {
                        const yearlyCurrentValue = getNumericValue(input);
                        const categoryInputs = document.querySelectorAll(`input.category-current-input[data-year="${year}"]`);
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
                const currentMonthKey = allMonths[36]; // Current month is at index 36
                if (currentMonthKey) {
                    const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${currentMonthKey}"]`);
                    if (mainForecastInput) {
                        const mainForecastValue = getNumericValue(mainForecastInput);
                        let categoriesForecastTotal = 0;
                        const categoryForecastInputs = document.querySelectorAll(`input.category-forecast-input[data-month="${currentMonthKey}"]`);
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
            if (!validateFormSubmission()) {
                return;
            }
            // Add this function to show/hide collection sections
            const toggleCollectionSections = (show) => {
                const sections = [
                    '.obl-section-container',
                    '#historicalCollectionSection',
                    '#detailedCollectionSection',
                    '#longtermCollectionSection',
                    '#yearlyCollectionSection'
                ];
                sections.forEach(selector => {
                    const element = document.querySelector(selector);
                    if (element) {
                        if (show) {
                            element.classList.remove('hidden-section');
                            element.classList.add('visible-section');
                        }
                        else {
                            element.classList.remove('visible-section');
                            element.classList.add('hidden-section');
                        }
                    }
                });
            };
            const calculateOBLBreakdown = async (projectId) => {
                const breakdown = {
                    normalDue: 0,
                    normalOverdue: 0,
                    retentionDue: 0,
                    retentionOverdue: 0,
                    gstTds: 0,
                    wht: 0
                };
                try {
                    const collection = collections.find(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
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
                        gstTds: Number(collection.GSTTDS) || 0, // ✅ Use GSTTDS
                        wht: Number(collection.WHT) || 0
                    };
                    console.log(`✅ Breakdown result:`, result);
                    return result;
                }
                catch (error) {
                    console.error("❌ Error in calculateOBLBreakdown:", error);
                    return breakdown;
                }
            };
            const renderOBLBreakdown = (breakdown) => {
                const setBreakdownValue = (elementId, value) => {
                    const element = document.getElementById(elementId);
                    if (element) {
                        // Always show the actual value (0 if no data)
                        setNumericValue(element, value);
                    }
                    else {
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
            const onProjectChange = async (projectId, existingCollection) => {
                const project = projects.find(p => p.Id === projectId);
                if (!project)
                    return;
                console.log(`🔄 Loading project: ${project.Title} (ID: ${projectId})`);
                // Show all collection sections
                toggleCollectionSections(true);
                renderProjectDetails(project);
                let collectionData = {};
                if (existingCollection) {
                    collectionData = safeJsonParse(existingCollection.CollectionData, {});
                }
                const isNewProject = !existingCollection;
                const oblContainer = document.querySelector('.obl-section-container');
                const fieldNote = oblContainer.querySelector('.field-note');
                // Calculate OBL breakdown FIRST
                const breakdown = await calculateOBLBreakdown(projectId);
                console.log(`📊 Breakdown for ${project.Title}:`, breakdown);
                // Render breakdown
                renderOBLBreakdown(breakdown);
                if (isNewProject) {
                    oblContainer.className = 'obl-section-container visible-section new-project';
                    oblLastMonthInput.removeAttribute('readonly');
                    oblLastMonthInput.placeholder = "Enter initial OBL for new project";
                    if (fieldNote)
                        fieldNote.textContent = "For new projects: Enter initial OBL value";
                    oblValidationMessage.textContent = "Enter initial OBL value for new project";
                    oblValidationMessage.className = 'validation-message';
                    const defaultOBL = project.PC ? 100000 : 0;
                    setNumericValue(oblLastMonthInput, defaultOBL);
                }
                else {
                    oblContainer.className = 'obl-section-container visible-section existing-project';
                    oblLastMonthInput.setAttribute('readonly', 'true');
                    // ✅ Use OBL field from the collection
                    const oblValue = Number(existingCollection === null || existingCollection === void 0 ? void 0 : existingCollection.OBL) || 0;
                    console.log(`💰 OBL Value for ${project.Title}: ${oblValue}`);
                    setNumericValue(oblLastMonthInput, oblValue);
                    oblLastMonthInput.placeholder = "Receivable value";
                    if (fieldNote)
                        fieldNote.textContent = "Receivable value record";
                    if (oblValue > 0) {
                        oblLastMonthInput.value = oblValue.toString();
                        oblLastMonthInput.className = oblLastMonthInput.className.replace('input-empty', 'input-with-value');
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
            // Function to create the OBL breakdown list in SharePoint (run once)
            // const createOBLBreakdownList = async (): Promise<void> => {
            //   try {
            //     const digest = await getDigest();
            //     // Check if list already exists
            //     const checkResponse = await fetch(`${siteUrl}/_api/web/lists/getbytitle('OBLBreakdown')`, {
            //       method: 'GET',
            //       headers: {
            //         'Accept': 'application/json;odata=verbose',
            //         'Content-Type': 'application/json;odata=verbose'
            //       }
            //     });
            //     if (checkResponse.ok) { 
            //       return;
            //     }
            //     // Create list only if it doesn't exist
            //     const listDefinition = {
            //       __metadata: { type: 'SP.List' },
            //       Title: 'OBLBreakdown',
            //       Description: 'Stores OBL breakdown details for projects',
            //       BaseTemplate: 100
            //     };
            //     const createListResponse = await fetch(`${siteUrl}/_api/web/lists`, {
            //       method: 'POST',
            //       headers: {
            //         'Accept': 'application/json;odata=verbose',
            //         'Content-Type': 'application/json;odata=verbose',
            //         'X-RequestDigest': digest
            //       },
            //       body: JSON.stringify(listDefinition)
            //     });
            //     if (createListResponse.ok) { 
            //       // Add fields to the list (your existing field creation code)
            //     } else { 
            //     }
            //   } catch (error) { 
            //   }
            // };
            // ========== HISTORICAL SECTION TOGGLE ========== 
            let historicalVisible = false;
            const setupHistoricalToggle = () => {
                // Wait for the toggle elements to exist
                setTimeout(() => {
                    const toggleBtn = document.getElementById('historicalToggleBtn');
                    const header = document.getElementById('historicalSectionHeader');
                    const historicalContainer = document.getElementById('historicalContainer');
                    if (!toggleBtn || !header || !historicalContainer) {
                        console.warn('Historical toggle elements not found');
                        return;
                    }
                    // Set initial state - HIDDEN by default
                    historicalVisible = false;
                    historicalContainer.style.display = 'none';
                    historicalContainer.style.opacity = '0';
                    historicalContainer.style.maxHeight = '0';
                    historicalContainer.style.padding = '0';
                    historicalContainer.style.overflow = 'hidden';
                    toggleBtn.textContent = '▼ Show';
                    toggleBtn.style.backgroundColor = '#6c757d';
                    toggleBtn.style.borderColor = '#6c757d';
                    // Toggle function
                    const toggleHistorical = (e) => {
                        // Don't toggle if clicking on the button itself (it has its own handler)
                        if (e.target === toggleBtn) {
                            e.stopPropagation();
                            return;
                        }
                        // If clicking on the header (not the button), toggle
                        if (header.contains(e.target) && e.target !== toggleBtn) {
                            performToggle();
                        }
                    };
                    // Button click handler
                    toggleBtn.addEventListener('click', function (e) {
                        e.stopPropagation();
                        performToggle();
                    });
                    // Header click handler (entire header)
                    header.addEventListener('click', toggleHistorical);
                    // Hover effect for header
                    header.addEventListener('mouseenter', function () {
                        this.style.backgroundColor = '#f8f9fa';
                    });
                    header.addEventListener('mouseleave', function () {
                        this.style.backgroundColor = 'transparent';
                    });
                    // Perform toggle function
                    const performToggle = () => {
                        historicalVisible = !historicalVisible;
                        if (historicalVisible) {
                            // Show historical data with animation
                            historicalContainer.style.display = 'block';
                            historicalContainer.style.padding = '5px 0';
                            // Trigger animation
                            requestAnimationFrame(() => {
                                historicalContainer.style.opacity = '1';
                                historicalContainer.style.maxHeight = '2000px';
                                historicalContainer.style.overflow = 'visible';
                            });
                            toggleBtn.textContent = '▲ Hide';
                            toggleBtn.style.backgroundColor = '#007bff';
                            toggleBtn.style.borderColor = '#007bff';
                            toggleBtn.style.color = 'white';
                        }
                        else {
                            // Hide historical data with animation
                            historicalContainer.style.opacity = '0';
                            historicalContainer.style.maxHeight = '0';
                            historicalContainer.style.overflow = 'hidden';
                            historicalContainer.style.padding = '0';
                            // After animation complete, hide completely
                            setTimeout(() => {
                                if (!historicalVisible) {
                                    historicalContainer.style.display = 'none';
                                }
                            }, 400);
                            toggleBtn.textContent = '▼ Show';
                            toggleBtn.style.backgroundColor = '#6c757d';
                            toggleBtn.style.borderColor = '#6c757d';
                            toggleBtn.style.color = 'white';
                        }
                    };
                    console.log('✅ Historical toggle initialized - default hidden');
                }, 300);
            };
            const renderProjectDetails = (project) => {
                var _a, _b;
                if (!project) {
                    console.warn("renderProjectDetails: project is missing");
                    return;
                }
                const projectNameDisplay = document.querySelector("#projectNameDisplay");
                const projectRef = document.querySelector("#projectRef");
                const profitCenter = document.querySelector("#profitCenter");
                const projectManager = document.querySelector("#projectManager");
                const groupHead = document.querySelector("#groupHead");
                const customer = document.querySelector("#customer");
                if (projectNameDisplay) {
                    projectNameDisplay.textContent =
                        project.Title || "-";
                }
                if (projectRef) {
                    projectRef.textContent =
                        project.ProjectReference || "-";
                }
                if (profitCenter) {
                    profitCenter.textContent =
                        project.PC || "-";
                }
                if (projectManager) {
                    const manager = typeof project.ProjectManager === "object"
                        ? (_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title
                        : project.ProjectManager;
                    projectManager.textContent =
                        manager || "-";
                }
                if (groupHead) {
                    groupHead.textContent =
                        project.GroupHead || "-";
                }
                if (customer) {
                    customer.textContent =
                        project.CustomerName || "-";
                }
                console.log("✅ Project details rendered:", {
                    project: project.Title,
                    projectRef: project.ProjectReference,
                    profitCenter: project.PC,
                    projectManager: typeof project.ProjectManager === "object"
                        ? (_b = project.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                        : project.ProjectManager,
                    groupHead: project.GroupHead,
                    customer: project.CustomerName
                });
            };
            const getCollectionMonths = () => {
                return generateIdealMonthStructure();
            };
            // ========== FIXED RENDER COLLECTION INTERFACE ==========
            const renderCollectionInterface = (existingData, displayOBL, existingCollection) => {
                allMonths = getCollectionMonths();
                setNumericValue(oblLastMonthInput, displayOBL);
                // ✅ CRITICAL FIX: If existingData is empty, create fresh structure
                if (!existingData || typeof existingData !== 'object' || Object.keys(existingData).length === 0) {
                    console.log("🔄 No existing data found, creating fresh collection structure");
                    existingData = createFreshCollectionStructure();
                }
                // ✅ Get the ideal month structure
                const idealMonths = generateIdealMonthStructure();
                // const storeHistoricalMonths = idealMonths.slice(0, 36);
                const displayHistoricalMonths = idealMonths.slice(33, 36);
                // ✅ FIX: Ensure ALL months (historical, detailed, long-term) have monthActual
                // Loop through ALL months in idealMonths
                idealMonths.forEach((monthKey) => {
                    if (!existingData[monthKey]) {
                        // Create new month data with monthActual
                        existingData[monthKey] = {
                            plan: 0,
                            actual: 0,
                            monthActual: 0, // ✅ Always initialize monthActual
                            forecast: 0,
                            date: '',
                            items: {}
                        };
                    }
                    else {
                        const monthData = existingData[monthKey];
                        // ✅ Ensure monthActual exists for ALL months
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
                // ✅ Ensure detailed months (36-38) have proper structure
                for (let i = 36; i <= 38 && i < idealMonths.length; i++) {
                    const monthKey = idealMonths[i];
                    if (existingData[monthKey]) {
                        const monthData = existingData[monthKey];
                        // ✅ Ensure all 4 categories exist
                        if (!monthData.items || Object.keys(monthData.items).length === 0) {
                            monthData.items = createEmptyCategories();
                        }
                        else {
                            const categories = ['Normal', 'Normal Overdue', 'Retention', 'Retention Overdue'];
                            categories.forEach(cat => {
                                if (!monthData.items[cat]) {
                                    monthData.items[cat] = {
                                        name: cat,
                                        plan: 0,
                                        actual: 0,
                                        forecast: 0,
                                        date: ''
                                    };
                                }
                                // ✅ Ensure forecast and date exist for existing categories
                                if (monthData.items[cat].forecast === undefined) {
                                    monthData.items[cat].forecast = 0;
                                }
                                if (monthData.items[cat].date === undefined) {
                                    monthData.items[cat].date = '';
                                }
                            });
                        }
                    }
                }
                // ✅ Ensure long-term months (39+) have proper structure
                for (let i = 39; i < idealMonths.length; i++) {
                    const monthKey = idealMonths[i];
                    if (existingData[monthKey]) {
                        const monthData = existingData[monthKey];
                        if (!monthData.items || Object.keys(monthData.items).length === 0) {
                            monthData.items = createSimpleCategories();
                        }
                        else {
                            const categories = ['Normal', 'Retention'];
                            categories.forEach(cat => {
                                if (!monthData.items[cat]) {
                                    monthData.items[cat] = {
                                        name: cat,
                                        plan: 0,
                                        actual: 0
                                    };
                                }
                            });
                        }
                    }
                }
                // ✅ Ensure yearly data exists
                if (!existingData.yearly || Object.keys(existingData.yearly).length === 0) {
                    existingData.yearly = createFreshYearlyStructure(idealMonths[idealMonths.length - 1]);
                }
                // Render only the LAST 3 months in the UI
                renderHistoricalMonths(displayHistoricalMonths, existingData);
                setTimeout(() => {
                    setupHistoricalToggle();
                }, 50);
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
                        const formTitleEl = document.querySelector("#formTitle");
                        if (formTitleEl) {
                            if (isFrozen) {
                                formTitleEl.innerHTML = '🔒 Edit Collection (Frozen - Further fields editable)';
                                formTitleEl.style.color = '#856404';
                            }
                            else {
                                formTitleEl.innerHTML = 'Edit Collection';
                                formTitleEl.style.color = '';
                            }
                        }
                    }, 250);
                }
            };
            const protectHistoricalElements = () => {
                // Remove any existing historical elements and replace with readonly versions
                document.querySelectorAll('.historical-month').forEach(monthElement => {
                    var _a;
                    // Clone the element to remove event listeners
                    const newElement = monthElement.cloneNode(true);
                    (_a = monthElement.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newElement, monthElement);
                    // Add readonly attributes to all inputs and ensure they are properly styled
                    newElement.querySelectorAll('input').forEach(input => {
                        input.setAttribute('readonly', 'true');
                        input.style.backgroundColor = '#f8f9fa';
                        input.style.cursor = 'not-allowed';
                        input.style.pointerEvents = 'none';
                        // Ensure proper classes are maintained
                        if (input.classList.contains('main-current-input')) {
                            input.className = input.className.replace('input-empty', 'input-with-value');
                        }
                        if (input.classList.contains('category-current-input')) {
                            // Keep the class as is - it will show input-empty or input-with-value based on data
                        }
                    });
                });
                // Add global event prevention
                document.querySelectorAll('.historical-month').forEach(section => {
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
            const initializeFormValidation = () => {
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
            const renderHistoricalMonths = (months, existingData) => {
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
                    const data = existingData[monthKey] || { plan: 0, actual: 0, items: {} };
                    html += renderHistoricalMonthBox(monthKey, monthLabel, data, index);
                });
                html += `</div>`;
                historicalContainer.innerHTML = html;
                // Keep container hidden by default
                historicalContainer.style.display = 'none';
                historicalContainer.style.opacity = '0';
                historicalContainer.style.maxHeight = '0';
                historicalContainer.style.padding = '0';
                historicalContainer.style.overflow = 'hidden';
                // Setup toggle functionality after rendering
                setupHistoricalToggle();
                // Protect historical elements after rendering
                protectHistoricalElements();
            };
            const renderHistoricalMonthBox = (monthKey, monthLabel, data, index) => {
                var _a, _b, _c, _d, _e, _f, _g, _h;
                // Get values
                const planValue = data.plan || 0;
                const actualValue = data.actual || 0;
                const monthActualValue = data.monthActual || 0; // ✅ Get monthActual (RevenueActual)
                const displayPlanValue = planValue > 0 ? planValue.toString() : '';
                const displayActualValue = actualValue > 0 ? actualValue.toString() : '';
                const displayMonthActual = monthActualValue > 0 ? monthActualValue.toString() : '';
                const planInputClass = planValue > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = actualValue > 0 ? 'input-with-value' : 'input-empty';
                const monthActualClass = monthActualValue > 0 ? 'input-with-value' : 'input-empty';
                const categories = data.items || createEmptyCategories();
                // Calculate category values - ensure they match the main total
                const normalPlan = (((_a = categories.Normal) === null || _a === void 0 ? void 0 : _a.plan) || 0);
                const normalActual = (((_b = categories.Normal) === null || _b === void 0 ? void 0 : _b.actual) || 0);
                const normalOverduePlan = (((_c = categories['Normal Overdue']) === null || _c === void 0 ? void 0 : _c.plan) || 0);
                const normalOverdueActual = (((_d = categories['Normal Overdue']) === null || _d === void 0 ? void 0 : _d.actual) || 0);
                const retentionPlan = (((_e = categories.Retention) === null || _e === void 0 ? void 0 : _e.plan) || 0);
                const retentionActual = (((_f = categories.Retention) === null || _f === void 0 ? void 0 : _f.actual) || 0);
                const retentionOverduePlan = (((_g = categories['Retention Overdue']) === null || _g === void 0 ? void 0 : _g.plan) || 0);
                const retentionOverdueActual = (((_h = categories['Retention Overdue']) === null || _h === void 0 ? void 0 : _h.actual) || 0);
                // Combined totals (should equal main values)
                const totalNormalPlan = normalPlan + normalOverduePlan;
                const totalNormalActual = normalActual + normalOverdueActual;
                const totalRetentionPlan = retentionPlan + retentionOverduePlan;
                const totalRetentionActual = retentionActual + retentionOverdueActual;
                const categoriesTotalPlan = totalNormalPlan + totalRetentionPlan;
                const categoriesTotalActual = totalNormalActual + totalRetentionActual;
                return `
        <div class="month-box historical-month" style="border: 1px solid #dee2e6; border-radius: 8px; padding: 12px; margin-bottom: 10px; background: #f8f9fa;">
            <div class="month-header" style="font-weight: bold; font-size: 14px; color: #495057; margin-bottom: 10px; border-bottom: 2px solid #007bff; padding-bottom: 5px;">
                ${monthLabel} 
                <span style="font-size: 11px; font-weight: normal; color: #6c757d; margin-left: 8px;">(Historical - Readonly)</span>
            </div>
            
            <!-- Main Inputs -->
            <div class="month-block" style="margin-bottom: 10px;">
                <div class="main-inputs-container" style="display: flex; gap: 15px; flex-wrap: wrap;">
                    <div class="main-input-group" style="flex: 1; min-width: 100px;">
                        <label style="font-size: 11px; color: #6c757d; display: block; margin-bottom: 2px;">Previous</label>
                        <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                              readonly class="main-plan-input ${planInputClass} number-input" 
                              style="width: 100%; padding: 6px 8px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 14px; text-align: right; cursor: not-allowed;" />
                    </div>
                    <div class="main-input-group" style="flex: 1; min-width: 100px;">
                        <label style="font-size: 11px; color: #6c757d; display: block; margin-bottom: 2px;">Current</label>
                        <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                              readonly class="main-current-input ${actualInputClass} number-input"
                              style="width: 100%; padding: 6px 8px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 14px; text-align: right; cursor: not-allowed;" />
                    </div>
                    <!-- ✅ Display monthActual (RevenueActual) for historical months -->
                    <div class="main-input-group" style="flex: 1; min-width: 100px;">
                        <label style="font-size: 11px; color: #6c757d; display: block; margin-bottom: 2px;">Actual</label>
                        <input type="text" value="${displayMonthActual}" data-type="month-actual" data-month="${monthKey}"
                              readonly class="month-actual-input ${monthActualClass} number-input"
                              style="width: 100%; padding: 6px 8px; border: 1px solid #28a745; border-radius: 4px; background: #e8f5e9; font-size: 14px; text-align: right; font-weight: bold; color: #1a8a34; cursor: not-allowed;" />
                    </div>
                </div>
            </div>

            <!-- Categories Section -->
            <div class="categories-section" style="border-top: 1px solid #e9ecef; padding-top: 10px;">
                <div class="categories-header" style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 5px; font-size: 11px; font-weight: bold; color: #495057; padding: 5px 0; border-bottom: 1px solid #dee2e6; margin-bottom: 5px;">
                    <div>Category</div>
                    <div style="text-align: right;">Previous</div>
                    <div style="text-align: right;">Current</div>
                </div>
                
                <!-- Normal Category -->
                <div class="category-row" style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Normal" data-category="Normal" data-month="${monthKey}" 
                          class="input-with-value" readonly 
                          style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${normalPlan > 0 ? normalPlan.toString() : ''}" 
                          data-category-type="plan" data-category="Normal" data-month="${monthKey}" 
                          readonly class="category-plan-input ${normalPlan > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${normalActual > 0 ? normalActual.toString() : ''}" 
                          data-category-type="actual" data-category="Normal" data-month="${monthKey}" 
                          readonly class="category-current-input ${normalActual > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
                
                <!-- Normal Overdue Category -->
                <div class="category-row" style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Normal Overdue" data-category="Normal Overdue" data-month="${monthKey}" 
                          class="input-with-value" readonly 
                          style="border: none; background: transparent; font-weight: 500; color: #856404; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${normalOverduePlan > 0 ? normalOverduePlan.toString() : ''}" 
                          data-category-type="plan" data-category="Normal Overdue" data-month="${monthKey}" 
                          readonly class="category-plan-input ${normalOverduePlan > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${normalOverdueActual > 0 ? normalOverdueActual.toString() : ''}" 
                          data-category-type="actual" data-category="Normal Overdue" data-month="${monthKey}" 
                          readonly class="category-current-input ${normalOverdueActual > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
                
                <!-- Retention Category -->
                <div class="category-row" style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Retention" data-category="Retention" data-month="${monthKey}" 
                          class="input-with-value" readonly 
                          style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${retentionPlan > 0 ? retentionPlan.toString() : ''}" 
                          data-category-type="plan" data-category="Retention" data-month="${monthKey}" 
                          readonly class="category-plan-input ${retentionPlan > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${retentionActual > 0 ? retentionActual.toString() : ''}" 
                          data-category-type="actual" data-category="Retention" data-month="${monthKey}" 
                          readonly class="category-current-input ${retentionActual > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
                
                <!-- Retention Overdue Category -->
                <div class="category-row" style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 5px; align-items: center; padding: 3px 0;">
                    <input type="text" value="Retention Overdue" data-category="Retention Overdue" data-month="${monthKey}" 
                          class="input-with-value" readonly 
                          style="border: none; background: transparent; font-weight: 500; color: #856404; padding: 4px 0; cursor: default;" />
                    <input type="text" value="${retentionOverduePlan > 0 ? retentionOverduePlan.toString() : ''}" 
                          data-category-type="plan" data-category="Retention Overdue" data-month="${monthKey}" 
                          readonly class="category-plan-input ${retentionOverduePlan > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #f8f9fa; font-size: 13px; text-align: right; cursor: not-allowed;" />
                    <input type="text" value="${retentionOverdueActual > 0 ? retentionOverdueActual.toString() : ''}" 
                          data-category-type="actual" data-category="Retention Overdue" data-month="${monthKey}" 
                          readonly class="category-current-input ${retentionOverdueActual > 0 ? 'input-with-value' : 'input-empty'} number-input"
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
            const renderFirstRow = (months, existingData) => {
                let html = '';
                // Create single row container for all three months
                html += `<div class="detailed-months-row">`;
                // Current Month (first of the detailed months)
                const currentMonth = months[0];
                if (currentMonth) {
                    const monthDate = new Date(currentMonth + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[currentMonth] || { plan: 0, actual: 0, items: createEmptyCategories() };
                    html += `<div class="current-month-container">`;
                    html += renderDetailedMonthBox(currentMonth, monthLabel, data, true);
                    html += `</div>`;
                }
                // Next 2 Months
                html += `<div class="next-months-container">`;
                for (let i = 1; i <= 2; i++) {
                    if (i < months.length) {
                        const monthKey = months[i];
                        const monthDate = new Date(monthKey + "-01");
                        const monthLabel = formatMonthLabel(monthDate);
                        const data = existingData[monthKey] || { plan: 0, actual: 0, items: createEmptyCategories() };
                        html += `<div class="next-month-box">`;
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
                for (let row = 0; row < 4; row++) {
                    html += `<div class="remaining-row">`;
                    for (let col = 0; col < 4; col++) {
                        const index = (row * 4) + col;
                        if (index >= months.length)
                            break;
                        const monthKey = months[index];
                        const monthDate = new Date(monthKey + "-01");
                        const monthLabel = formatMonthLabel(monthDate);
                        const data = existingData[monthKey] || { plan: 0, actual: 0, items: createSimpleCategories() };
                        html += renderSimpleMonthBox(monthKey, monthLabel, data);
                    }
                    html += `</div>`;
                }
                remainingContainer.innerHTML = html;
            };
            const renderYearlyCollection = (yearlyData) => {
                let html = '';
                const years = Object.keys(yearlyData).sort();
                years.forEach(year => {
                    const data = yearlyData[year] || { plan: 0, actual: 0 };
                    html += renderYearlyBox(year, data);
                });
                yearlyContainer.innerHTML = html;
            };
            const renderYearlyBox = (year, data) => {
                var _a, _b, _c, _d, _e, _f, _g, _h;
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
                const categories = data.items || {
                    "Normal": { name: 'Normal', plan: 0, actual: 0 },
                    "Retention": { name: 'Retention', plan: 0, actual: 0 }
                };
                return `
    <div class="yearly-box">
      <div class="yearly-header">${year}</div>
      
      <div class="yearly-inputs-container">
        <div class="yearly-input-group">
          <label>Previous</label>
          <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                readonly class="yearly-plan-input ${planInputClass} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
        <div class="yearly-input-group">
          <label>Current</label>
          <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                readonly class="yearly-current-input ${actualInputClass} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
      </div>

      <div class="categories-section">
        <div class="categories-header">
          <div>Category</div>
          <div>Previous</div>
          <div>Current</div>
        </div>
        
        <!-- Normal Category - READONLY -->
        <div class="category-row">
          <input type="text" value="Normal" data-category="Normal" data-year="${year}" 
                class="input-with-value" readonly 
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${((_a = categories.Normal) === null || _a === void 0 ? void 0 : _a.plan) > 0 ? categories.Normal.plan.toString() : ''}" 
                data-category-type="plan" data-category="Normal" data-year="${year}" 
                readonly class="category-plan-input ${((_b = categories.Normal) === null || _b === void 0 ? void 0 : _b.plan) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${((_c = categories.Normal) === null || _c === void 0 ? void 0 : _c.actual) > 0 ? categories.Normal.actual.toString() : ''}" 
                data-category-type="actual" data-category="Normal" data-year="${year}" 
                readonly class="category-current-input ${((_d = categories.Normal) === null || _d === void 0 ? void 0 : _d.actual) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
        
        <!-- Retention Category - READONLY -->
        <div class="category-row">
          <input type="text" value="Retention" data-category="Retention" data-year="${year}" 
                class="input-with-value" readonly 
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${((_e = categories.Retention) === null || _e === void 0 ? void 0 : _e.plan) > 0 ? categories.Retention.plan.toString() : ''}" 
                data-category-type="plan" data-category="Retention" data-year="${year}" 
                readonly class="category-plan-input ${((_f = categories.Retention) === null || _f === void 0 ? void 0 : _f.plan) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${((_g = categories.Retention) === null || _g === void 0 ? void 0 : _g.actual) > 0 ? categories.Retention.actual.toString() : ''}" 
                data-category-type="actual" data-category="Retention" data-year="${year}" 
                readonly class="category-current-input ${((_h = categories.Retention) === null || _h === void 0 ? void 0 : _h.actual) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
        
        <!-- Validation Display - Simplified -->
        <div class="yearly-total-validation" data-year="${year}" style="margin-top: 10px; font-size: 11px; text-align: center;"></div>
      </div>
    </div>
  `;
            };
            // ========== FIXED RENDER DETAILED MONTH BOX - FORECAST EDITABLE ==========
            const renderDetailedMonthBox = (monthKey, monthLabel, data, isCurrentMonth = false) => {
                // Get values from data
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayCurrentValue = data.actual > 0 ? data.actual.toString() : '';
                // ✅ Get monthActual from data (stored in CollectionData) - like revenue.ts
                let monthActualValue = data.monthActual || 0;
                let displayMonthActual = monthActualValue > 0 ? monthActualValue.toString() : '';
                // ✅ Get forecast/further value
                const forecastValue = data.forecast || 0;
                const displayForecastValue = forecastValue > 0 ? forecastValue.toString() : '';
                // const displayDateValue = data.date || '';
                // If monthActual is not in data and this is current month, try to get from existingCollection
                if (data.monthActual === undefined && isCurrentMonth) {
                    try {
                        const projectId = parseInt(projectSelect.value);
                        const collection = collections.find(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                        if (collection && collection._revenueActual) {
                            monthActualValue = Number(collection._revenueActual[monthKey]) || 0;
                            displayMonthActual = monthActualValue > 0 ? monthActualValue.toString() : '';
                        }
                    }
                    catch (e) {
                        monthActualValue = 0;
                    }
                }
                // ✅ Forecast = RevenueActual + Further (like revenue.ts)
                const forecastTotal = monthActualValue + forecastValue;
                const displayForecastTotal = forecastTotal > 0 ? forecastTotal.toString() : '';
                // CSS classes
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const currentInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
                const forecastInputClass = forecastValue > 0 ? 'input-with-value' : 'input-empty';
                // const dateInputClass = displayDateValue ? 'input-with-value' : 'input-empty';
                const monthActualClass = monthActualValue > 0 ? 'input-with-value' : 'input-empty';
                const forecastTotalClass = forecastTotal > 0 ? 'input-with-value' : 'input-empty';
                // Get categories with proper forecast values
                const categories = data.items || {};
                // ✅ For detailed months (Current + Next 2 Months) - Show ALL 4 categories
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
                    const categoryForecastClass = categoryForecast > 0 ? 'input-with-value' : 'input-empty';
                    const categoryDateClass = category.date ? 'input-with-value' : 'input-empty';
                    return `
            <tr style="border-bottom: 1px solid #dee2e6;">
                <td style="padding: 6px 8px; font-weight: 500; color: ${color};">${categoryKey}</td>
                <td style="padding: 6px 8px; text-align: right;">
                    <input type="text" value="${(category.plan || 0) > 0 ? category.plan.toString() : ''}" 
                          data-category-type="plan" data-category="${categoryKey}" data-month="${monthKey}" 
                          readonly class="category-plan-input ${(category.plan || 0) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 80px; padding: 4px 6px; border: 1px solid #e9ecef; border-radius: 4px; background: #e9ecef; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </td>
                <td style="padding: 6px 8px; text-align: right;">
                    <input type="text" value="${(category.actual || 0) > 0 ? category.actual.toString() : ''}" 
                          data-category-type="actual" data-category="${categoryKey}" data-month="${monthKey}" 
                          class="category-current-input ${(category.actual || 0) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                          style="width: 80px; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #ffffff; font-size: 13px; text-align: right; cursor: text;" />
                </td>
                ${isCurrentMonth ? `
                    <td style="padding: 6px 8px; text-align: right;">
                        <input type="text" value="${displayCategoryForecast}" 
                              data-category-type="forecast" data-category="${categoryKey}" data-month="${monthKey}" 
                              class="category-forecast-input ${categoryForecastClass} number-input"
                              style="width: 80px; padding: 4px 6px; border: 1px solid #28a745; border-radius: 4px; background: #ffffff; font-size: 13px; text-align: right; cursor: text;"
                              placeholder="0" />
                    </td>
                    <td style="padding: 6px 8px; text-align: center;">
                        <input type="date" value="${category.date || ''}" 
                              data-category-type="date" data-category="${categoryKey}" data-month="${monthKey}" 
                              class="category-date-input ${categoryDateClass}"
                              style="width: 100px; padding: 4px 6px; border: 1px solid #28a745; border-radius: 4px; background: #ffffff; font-size: 13px; text-align: center; cursor: text;"
                              placeholder="Date" />
                    </td>
                ` : ''}
            </tr>
        `;
                }).join('');
                // Month Actual badge for current month (like revenue.ts)
                const monthActualBadge = isCurrentMonth ? `
        <span class="month-actual-badge" style="
            background-color: ${monthActualValue > 0 ? '#c8e6c9' : '#f5f5f5'};
            border: 1px solid ${monthActualValue > 0 ? '#4caf50' : '#ddd'};
            border-radius: 4px;
            padding: 2px 12px;
            font-size: 13px;
            font-weight: 600;
            color: ${monthActualValue > 0 ? '#1b5e20' : '#999'};
            display: inline-flex;
            align-items: center;
            gap: 6px;
            white-space: nowrap;
        ">
            <span>📊 Month Actual:</span>
            <span style="font-size: 16px; font-weight: 700;">${monthActualValue.toLocaleString()}</span>
        </span>
    ` : '';
                return `
        <div class="detailed-month-box ${isCurrentMonth ? 'current-month-box' : 'next-month-box'}" style="border: 1px solid #dee2e6; border-radius: 8px; padding: 12px; margin-bottom: 15px;">
            <div class="month-header" style="display: flex; justify-content: space-between; align-items: center; font-weight: bold; font-size: 14px; color: #495057; margin-bottom: 10px; border-bottom: 2px solid #007bff; padding-bottom: 5px;">
                <span>${monthLabel} ${isCurrentMonth ? '<span style="background: #007bff; color: white; padding: 2px 8px; border-radius: 4px; font-size: 10px; margin-left: 8px;">Current</span>' : ''}</span>
                ${monthActualBadge}
            </div>
            
            <!-- Top Row: Previous, Current, Actual, Further, Forecast -->
            <div style="display: grid; grid-template-columns: ${isCurrentMonth ? '1fr 1fr 1fr 1fr 1fr' : '1fr 1fr'}; gap: 8px; margin-bottom: 12px; background: #f8f9fa; padding: 8px; border-radius: 4px;">
                <div>
                    <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Previous</label>
                    <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                          readonly class="main-plan-input ${planInputClass} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
                <div>
                    <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Current</label>
                    <input type="text" value="${displayCurrentValue}" data-type="actual" data-month="${monthKey}"
                          readonly class="main-current-input ${currentInputClass} number-input"
                          style="width: 100%; padding: 4px 6px; border: 1px solid #ced4da; border-radius: 4px; background: #e9ecef; font-size: 13px; text-align: right; cursor: not-allowed;" />
                </div>
                ${isCurrentMonth ? `
                    <div>
                        <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Actual</label>
                        <input type="text" value="${displayMonthActual}" data-type="month-actual" data-month="${monthKey}"
                              readonly class="month-actual-input ${monthActualClass} number-input"
                              style="width: 100%; padding: 4px 6px; border: 1px solid #28a745; border-radius: 4px; background: #e8f5e9; font-size: 13px; text-align: right; font-weight: bold; color: #1a8a34;" />
                    </div>
                    <div>
                        <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Further</label>
                        <input type="text" value="${displayForecastValue}" data-type="forecast" data-month="${monthKey}"
                              class="main-forecast-input ${forecastInputClass} number-input"
                              style="width: 100%; padding: 4px 6px; border: 1px solid #28a745; border-radius: 4px; background: #ffffff; font-size: 13px; text-align: right; cursor: text;" />
                    </div>
                    <div>
                        <label style="font-size: 10px; color: #6c757d; display: block; text-align: center; font-weight: bold;">Forecast</label>
                        <input type="text" value="${displayForecastTotal}" data-type="forecast-total" data-month="${monthKey}"
                              readonly class="forecast-total-input ${forecastTotalClass} number-input"
                              style="width: 100%; padding: 4px 6px; border: 1px solid #007bff; border-radius: 4px; background: #f0f8ff; font-size: 13px; text-align: right; font-weight: bold; color: #0c5460;" />
                    </div>
                ` : ''}
            </div>
            
            <!-- Category Table - ALL 4 categories for Current + Next 2 Months -->
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
                <div class="forecast-validation" data-month="${monthKey}" style="margin-top: 10px; font-size: 11px; text-align: center; color: #6c757d;">
                    <span style="display: block; padding: 4px 8px; background: #f8f9fa; border-radius: 4px; border: 1px solid #dee2e6;">
                        💡 Forecast = Actual + Further
                    </span>
                </div>
            ` : ''}
        </div>
    `;
            };
            const renderSimpleMonthBox = (monthKey, monthLabel, data) => {
                var _a, _b, _c, _d, _e, _f, _g, _h;
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
                const categories = data.items || createSimpleCategories();
                return `
    <div class="month-box simple-month">
      <div class="month-header">${monthLabel}</div>
      
      <div class="main-inputs-container">
        <div class="main-input-group">
          <label>Previous</label>
          <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}" 
                readonly class="main-plan-input ${planInputClass} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
        <div class="main-input-group">
          <label>Current</label>
          <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}" 
                readonly class="main-current-input ${actualInputClass} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
      </div>

      <div class="categories-section">
        <div class="categories-header">
          <div>Category</div>
          <div>Previous</div>
          <div>Current</div>
        </div>
        
        <!-- Normal Category - READONLY -->
        <div class="category-row">
          <input type="text" value="Normal" data-category="Normal" data-month="${monthKey}" 
                class="input-with-value" readonly 
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${((_a = categories.Normal) === null || _a === void 0 ? void 0 : _a.plan) > 0 ? categories.Normal.plan.toString() : ''}" 
                data-category-type="plan" data-category="Normal" data-month="${monthKey}" 
                readonly class="category-plan-input ${((_b = categories.Normal) === null || _b === void 0 ? void 0 : _b.plan) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${((_c = categories.Normal) === null || _c === void 0 ? void 0 : _c.actual) > 0 ? categories.Normal.actual.toString() : ''}" 
                data-category-type="actual" data-category="Normal" data-month="${monthKey}" 
                readonly class="category-current-input ${((_d = categories.Normal) === null || _d === void 0 ? void 0 : _d.actual) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
        
        <!-- Retention Category - READONLY -->
        <div class="category-row">
          <input type="text" value="Retention" data-category="Retention" data-month="${monthKey}" 
                class="input-with-value" readonly 
                style="border: none; background: transparent; font-weight: 500; color: #495057; padding: 4px 0; cursor: default;" />
          <input type="text" value="${((_e = categories.Retention) === null || _e === void 0 ? void 0 : _e.plan) > 0 ? categories.Retention.plan.toString() : ''}" 
                data-category-type="plan" data-category="Retention" data-month="${monthKey}" 
                readonly class="category-plan-input ${((_f = categories.Retention) === null || _f === void 0 ? void 0 : _f.plan) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
          <input type="text" value="${((_g = categories.Retention) === null || _g === void 0 ? void 0 : _g.actual) > 0 ? categories.Retention.actual.toString() : ''}" 
                data-category-type="actual" data-category="Retention" data-month="${monthKey}" 
                readonly class="category-current-input ${((_h = categories.Retention) === null || _h === void 0 ? void 0 : _h.actual) > 0 ? 'input-with-value' : 'input-empty'} number-input"
                style="background: #e9ecef; cursor: not-allowed;" />
        </div>
      </div>
      
      <!-- Validation display area --> 
    </div>
  `;
            };
            // ========== FIXED SET ALL SECTIONS READONLY ==========
            const setAllSectionsReadOnly = () => {
                // MAIN-CURRENT-INPUT - READONLY (shows total)
                document.querySelectorAll('.main-current-input').forEach((input) => {
                    const el = input;
                    el.setAttribute('readonly', 'true');
                    el.style.backgroundColor = '#e9ecef';
                    el.style.cursor = 'not-allowed';
                    el.style.pointerEvents = 'none';
                    const value = el.value.trim();
                    if (value !== '' && value !== '0') {
                        el.classList.remove('input-empty');
                        el.classList.add('input-with-value');
                    }
                    else {
                        el.classList.remove('input-with-value');
                        el.classList.add('input-empty');
                    }
                });
                // CATEGORY-CURRENT-INPUT - EDITABLE (user input)
                document.querySelectorAll('.category-current-input').forEach((input) => {
                    const el = input;
                    el.removeAttribute('readonly');
                    el.style.backgroundColor = '#ffffff';
                    el.style.cursor = 'text';
                    el.style.pointerEvents = 'auto';
                    const value = el.value.trim();
                    if (value !== '' && value !== '0') {
                        el.classList.remove('input-empty');
                        el.classList.add('input-with-value');
                    }
                    else {
                        el.classList.remove('input-with-value');
                        el.classList.add('input-empty');
                    }
                });
                // YEARLY-CURRENT-INPUT - READONLY (shows total)
                document.querySelectorAll('.yearly-current-input').forEach((input) => {
                    const el = input;
                    el.setAttribute('readonly', 'true');
                    el.style.backgroundColor = '#e9ecef';
                    el.style.cursor = 'not-allowed';
                    el.style.pointerEvents = 'none';
                    const value = el.value.trim();
                    if (value !== '' && value !== '0') {
                        el.classList.remove('input-empty');
                        el.classList.add('input-with-value');
                    }
                    else {
                        el.classList.remove('input-with-value');
                        el.classList.add('input-empty');
                    }
                });
                // MAIN-FORECAST-INPUT - READONLY (auto-calculated from categories)
                document.querySelectorAll('.main-forecast-input').forEach((input) => {
                    const el = input;
                    el.setAttribute('readonly', 'true');
                    el.style.backgroundColor = '#e9ecef';
                    el.style.cursor = 'not-allowed';
                    el.style.pointerEvents = 'none';
                    const value = el.value.trim();
                    if (value !== '' && value !== '0') {
                        el.classList.remove('input-empty');
                        el.classList.add('input-with-value');
                    }
                    else {
                        el.classList.remove('input-with-value');
                        el.classList.add('input-empty');
                    }
                });
                // ===== IMPORTANT FIX: CATEGORY-FORECAST-INPUT - EDITABLE =====
                document.querySelectorAll('.category-forecast-input').forEach((input) => {
                    const el = input;
                    // REMOVE readonly attribute - make it editable
                    el.removeAttribute('readonly');
                    el.style.backgroundColor = '#ffffff';
                    el.style.cursor = 'text';
                    el.style.pointerEvents = 'auto';
                    el.style.opacity = '1';
                    const value = el.value.trim();
                    if (value !== '' && value !== '0') {
                        el.classList.remove('input-empty');
                        el.classList.add('input-with-value');
                    }
                    else {
                        el.classList.remove('input-with-value');
                        el.classList.add('input-empty');
                    }
                });
                // PLAN INPUTS (Previous) - READONLY
                document.querySelectorAll('.main-plan-input, .category-plan-input, .yearly-plan-input').forEach((input) => {
                    const el = input;
                    el.setAttribute('readonly', 'true');
                    el.style.backgroundColor = '#e9ecef';
                    el.style.cursor = 'not-allowed';
                    el.style.pointerEvents = 'none';
                });
            };
            const calculateMonthTotals = () => {
                // Find all unique months
                const monthKeys = new Set();
                document.querySelectorAll('.category-current-input[data-month]').forEach((input) => {
                    const el = input;
                    const month = el.dataset.month;
                    if (month)
                        monthKeys.add(month);
                });
                // For each month, calculate total and update main-current-input
                monthKeys.forEach(monthKey => {
                    const categoryInputs = document.querySelectorAll(`.category-current-input[data-month="${monthKey}"]`);
                    let total = 0;
                    categoryInputs.forEach(input => {
                        const value = getNumericValue(input);
                        total += value;
                    });
                    const mainInput = document.querySelector(`.main-current-input[data-month="${monthKey}"]`);
                    if (mainInput) {
                        // Set the total value
                        if (total > 0) {
                            mainInput.value = total.toString();
                            mainInput.className = mainInput.className.replace('input-empty', 'input-with-value');
                        }
                        else {
                            mainInput.value = '';
                            mainInput.className = mainInput.className.replace('input-with-value', 'input-empty');
                        }
                        // Ensure it's readonly
                        mainInput.setAttribute('readonly', 'true');
                        mainInput.style.backgroundColor = '#e9ecef';
                        mainInput.style.cursor = 'not-allowed';
                        mainInput.style.pointerEvents = 'none';
                    }
                });
                // Also calculate yearly totals
                const yearKeys = new Set();
                document.querySelectorAll('.category-current-input[data-year]').forEach((input) => {
                    const el = input;
                    const year = el.dataset.year;
                    if (year)
                        yearKeys.add(year);
                });
                yearKeys.forEach(year => {
                    const categoryInputs = document.querySelectorAll(`.category-current-input[data-year="${year}"]`);
                    let total = 0;
                    categoryInputs.forEach(input => {
                        const value = getNumericValue(input);
                        total += value;
                    });
                    const yearlyInput = document.querySelector(`.yearly-current-input[data-year="${year}"]`);
                    if (yearlyInput) {
                        if (total > 0) {
                            yearlyInput.value = total.toString();
                            yearlyInput.className = yearlyInput.className.replace('input-empty', 'input-with-value');
                        }
                        else {
                            yearlyInput.value = '';
                            yearlyInput.className = yearlyInput.className.replace('input-with-value', 'input-empty');
                        }
                    }
                });
            };
            // Add this to the OBL input change event
            oblLastMonthInput.addEventListener('input', function () {
                const value = getNumericValue(this);
                if (value > 0) {
                    this.className = this.className.replace('input-empty', 'input-with-value');
                }
                else {
                    this.className = this.className.replace('input-with-value', 'input-empty');
                }
                validateOBLLimit(this);
            });
            // ========== FIXED ATTACH COLLECTION EVENT LISTENERS ==========
            const attachCollectionEventListeners = () => {
                // Set up number inputs with allowOnlyNumbers
                document.querySelectorAll('.number-input').forEach(input => {
                    allowOnlyNumbers(input);
                });
                // Set up focus handlers for all number inputs to store old values
                document.querySelectorAll('.number-input:not([readonly])').forEach(input => {
                    input.addEventListener('focus', function () {
                        this.setAttribute('data-old-value', getNumericValue(this).toString());
                    });
                });
                // Category current inputs - handle changes
                document.querySelectorAll('.category-current-input:not([readonly])').forEach((input) => {
                    const el = input;
                    el.removeEventListener('input', handleCategoryChange);
                    el.addEventListener('input', handleCategoryChange);
                });
                // Yearly category current inputs
                document.querySelectorAll('.category-current-input[data-year]:not([readonly])').forEach((input) => {
                    const el = input;
                    el.removeEventListener('input', handleCategoryChange);
                    el.addEventListener('input', handleCategoryChange);
                });
                document.querySelectorAll('.category-date-input, .forecast-date-input').forEach(el => {
                    el.addEventListener('click', () => { var _a, _b; return (_b = (_a = el).showPicker) === null || _b === void 0 ? void 0 : _b.call(_a); });
                });
                // Category forecast inputs - handle changes (REMOVED readonly)
                document.querySelectorAll('.category-forecast-input').forEach((input) => {
                    const el = input;
                    // Remove readonly attribute - make editable
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
            const handleCategoryChange = function () {
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
                    const categoryInputs = document.querySelectorAll(`.category-current-input[data-month="${month}"]`);
                    let total = 0;
                    categoryInputs.forEach(input => {
                        total += getNumericValue(input);
                    });
                    const mainInput = document.querySelector(`.main-current-input[data-month="${month}"]`);
                    if (mainInput) {
                        if (total > 0) {
                            mainInput.value = total.toString();
                            mainInput.className = mainInput.className.replace('input-empty', 'input-with-value');
                        }
                        else {
                            mainInput.value = '';
                            mainInput.className = mainInput.className.replace('input-with-value', 'input-empty');
                        }
                    }
                }
                // Update yearly total if this is a yearly category
                if (year) {
                    const categoryInputs = document.querySelectorAll(`.category-current-input[data-year="${year}"]`);
                    let total = 0;
                    categoryInputs.forEach(input => {
                        total += getNumericValue(input);
                    });
                    const yearlyInput = document.querySelector(`.yearly-current-input[data-year="${year}"]`);
                    if (yearlyInput) {
                        if (total > 0) {
                            yearlyInput.value = total.toString();
                            yearlyInput.className = yearlyInput.className.replace('input-empty', 'input-with-value');
                        }
                        else {
                            yearlyInput.value = '';
                            yearlyInput.className = yearlyInput.className.replace('input-with-value', 'input-empty');
                        }
                    }
                }
                // Update OBL validation after a short delay
                setTimeout(() => {
                    validateOBLLimit(this);
                }, 50);
            };
            // ========== EDIT COLLECTION FUNCTION ==========
            const editCollection = async (id, existingCollection) => {
                var _a;
                const collection = existingCollection || collections.find(f => f.Id === id);
                if (!collection) {
                    toast("error", "Collection not found");
                    return;
                }
                editingId = collection.Id;
                const project = projects.find(p => { var _a; return p.Id === ((_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Id); });
                if (!canRead()) {
                    toast("error", "❌ You don't have permission to view this collection");
                    return;
                }
                const assignedProjects = getAssignedProjects(projects);
                const isAssigned = assignedProjects.some(p => { var _a; return p.Id === ((_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Id); });
                const isGHHead = getUserGroups().some((g) => g === 'KRIYA_GH_Heads');
                const isFullControl = hasFullControl();
                if (!isFullControl && !isGHHead && !isAssigned) {
                    toast("error", "❌ You don't have permission to view this project");
                    return;
                }
                const canEdit = authCanEditProject(project);
                const isFrozen = collection.Freeze || false;
                showForm(true);
                const remarksInternal = document.querySelector("#remarksInternal");
                const remarksSprintCell = document.querySelector("#remarksSprintCell");
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
                    }
                    else {
                        if (remarksInternal) {
                            remarksInternal.value = remarksData;
                        }
                        if (remarksSprintCell) {
                            remarksSprintCell.value = '';
                        }
                    }
                }
                let collectionData = {};
                let revenueActualData = {};
                try {
                    let collectionDataString = collection.CollectionData || "{}";
                    collectionDataString = collectionDataString.trim();
                    if (!collectionDataString || collectionDataString === '""' || collectionDataString === "null" || collectionDataString === "{}") {
                        collectionDataString = "{}";
                    }
                    collectionData = JSON.parse(collectionDataString);
                    // ✅ Check if data is empty or invalid
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
                }
                catch (error) {
                    console.error("Error parsing CollectionData:", error);
                    // ✅ Auto-create fresh structure on error
                    collectionData = createFreshCollectionStructure();
                    revenueActualData = {};
                    toast("warning", "Corrupted collection data detected. Loaded empty structure.");
                }
                // Apply progression
                collectionData = propagateMonthValues(collectionData);
                if ((_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Id) {
                    const projectId = Number(collection.Project.Id);
                    projectSelect.value = String(projectId);
                    await onProjectChange(projectId, collection);
                }
                setTimeout(() => {
                    const isReadOnly = isFrozen || !canEdit;
                    setFormReadOnly(isReadOnly);
                    const saveBtn = document.querySelector('.btn-primary[type="submit"]');
                    if (saveBtn) {
                        saveBtn.style.display = canEdit ? '' : 'none';
                        if (isFrozen) {
                            saveBtn.textContent = '💾 Save Further Changes';
                            saveBtn.title = 'Save only Further and Date field changes (Curr. Plan is locked)';
                            saveBtn.style.backgroundColor = '#28a745';
                            saveBtn.style.borderColor = '#28a745';
                            saveBtn.style.color = '#ffffff';
                            saveBtn.style.fontWeight = 'bold';
                            saveBtn.style.cursor = 'pointer';
                            saveBtn.setAttribute('type', 'submit');
                        }
                        else {
                            saveBtn.textContent = '💾 Save Collection Plan';
                            saveBtn.title = '';
                            saveBtn.style.backgroundColor = '';
                            saveBtn.style.borderColor = '';
                            saveBtn.style.color = '';
                            saveBtn.style.fontWeight = '';
                            saveBtn.style.cursor = '';
                            saveBtn.setAttribute('type', 'submit');
                        }
                    }
                    const formTitleEl = document.querySelector("#formTitle");
                    if (formTitleEl) {
                        if (isFrozen) {
                            formTitleEl.innerHTML = '🔒 Edit Collection (Frozen - Further fields editable)';
                            formTitleEl.style.color = '#856404';
                        }
                        else if (!canEdit) {
                            formTitleEl.innerHTML = '🔒 Edit Collection (Read Only)';
                            formTitleEl.style.color = '#856404';
                        }
                        else {
                            formTitleEl.innerHTML = 'Edit Collection';
                            formTitleEl.style.color = '';
                        }
                    }
                    if (isFrozen) {
                        toast('info', '🔒 Project is frozen. Only Further and Date fields are editable. Click "Save Further Changes" to save.');
                    }
                    else if (!canEdit) {
                        toast('info', '🔒 You have view-only access to this project.');
                    }
                    else {
                        toast('success', '🔓 All fields are editable. Click "Save Collection Plan" to save.');
                    }
                }, 300);
                toast("info", `Viewing collection for ${collection.Title}`);
            };
            // ========== FREEZE FUNCTIONS ==========
            // 3. Update the applyAutoProgression function for collection:
            const applyAutoProgression = async () => {
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
                            }
                            else {
                                errorCount++;
                                console.error(`❌ Failed to progress collection: ${collection.Id}`);
                            }
                            await new Promise(resolve => setTimeout(resolve, 100));
                        }
                    }
                    catch (error) {
                        errorCount++;
                        console.error(`❌ Error progressing collection: ${collection.Id}`, error);
                    }
                }
                if (progressedCount > 0) {
                    await loadCollections();
                    toast("success", `Monthly data progressed for ${progressedCount} collections`);
                }
                else if (errorCount > 0) {
                    toast("warning", `Some errors occurred during monthly progression`);
                }
            };
            const propagateCollectionValues = (collectionData) => {
                const updatedData = Object.assign({}, collectionData);
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
            const propagateYearlyValues = (collectionData, isNewMonth = false) => {
                if (!isNewMonth) {
                    return collectionData;
                }
                const updatedData = Object.assign({}, collectionData);
                if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) {
                    return updatedData;
                }
                const yearlyData = updatedData.yearly;
                const updatedYearlyData = {};
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
                                const category = yearData.items[categoryKey];
                                updatedYearlyData[year].items[categoryKey] = {
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
            // ========== IMPORT FUNCTIONALITY ==========
            const importActualData = () => {
                const modalHTML = `
          <div class="import-modal">
            <div class="import-modal-content">
              <div class="import-modal-header">
                <h3>Import Monthly Actual Collection Data</h3>
                <button class="close-modal">&times;</button>
              </div> 
              <div class="import-section">
                <h4>Step 1: Download Template</h4>
                <button class="btn-primary" id="downloadTemplate">⬇️ Download Template</button>
              </div>

              <div class="import-section">
                <h4>Step 2: Upload Excel File</h4>
                <div class="file-upload-area" id="fileUploadArea">
                  <div class="upload-icon">📁</div>
                  <p>Click to upload or drag & drop Excel file</p>
                  <p><small>Supported format: Excel with Project Reference, Month, and Actual Collection columns</small></p>
                  <input type="file" id="excelFile" class="file-input" accept=".xlsx,.xls">
                  <div class="file-info" id="fileInfo"></div>
                </div>
              </div>

              <div class="import-section" id="previewSection" style="display:none;">
                <h4>Step 3: Preview & Import</h4>
                <div class="import-preview" id="importPreview"></div>
                <div class="import-actions">
                  <button class="btn-secondary" id="cancelImport">Cancel</button>
                  <button class="btn-primary" id="confirmImport">Import Data</button>
                </div>
                <div class="import-status" id="importStatus"></div>
              </div>
            </div>
          </div>
        `;
                document.body.insertAdjacentHTML("beforeend", modalHTML);
                const modal = document.querySelector(".import-modal");
                modal.style.display = "flex";
                // Elements
                const fileUploadArea = document.querySelector("#fileUploadArea");
                const excelFileInput = document.querySelector("#excelFile");
                const fileInfo = document.querySelector("#fileInfo");
                const previewSection = document.querySelector("#previewSection");
                const importPreview = document.querySelector("#importPreview");
                const confirmImport = document.querySelector("#confirmImport");
                const cancelImport = document.querySelector("#cancelImport");
                const closeModal = document.querySelector(".close-modal");
                const importStatus = document.querySelector("#importStatus");
                const downloadTemplateBtn = document.querySelector("#downloadTemplate");
                let importData = [];
                let listEntityType = "";
                // Month name to number mapping
                const monthMap = {
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
                const parseMonth = (monthStr) => {
                    if (!monthStr)
                        return '';
                    // If already in YYYY-MM format
                    if (/^\d{4}-\d{2}$/.test(monthStr)) {
                        return monthStr;
                    }
                    // Handle formats like "Nov-25", "November 2025", "Nov 2025"
                    const match = monthStr.toLowerCase().match(/([a-z]+)[-\s]*'?(\d{2,4})?/);
                    if (match) {
                        const monthName = match[1];
                        let year = match[2];
                        const monthNum = monthMap[monthName];
                        if (monthNum) {
                            // Handle year
                            if (!year) {
                                // If no year provided, use current year
                                year = new Date().getFullYear().toString();
                            }
                            else if (year.length === 2) {
                                // Convert 2-digit year to 4-digit
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
                const getListEntityType = async () => {
                    if (listEntityType)
                        return listEntityType;
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                        headers: { Accept: "application/json;odata=verbose" },
                    });
                    const data = await res.json();
                    listEntityType = data.d.ListItemEntityTypeFullName;
                    return listEntityType;
                };
                // const updateRevenueData = (existingData: string, month: string, actualRevenue: number): string => {
                //     let finalData: { [key: string]: number } = {};
                //     if (existingData && existingData.trim() !== '' && existingData !== '{}' && existingData !== 'null' && existingData !== 'undefined') {
                //         try {
                //             const parsed = JSON.parse(existingData);
                //             // Only assign if parsed is an object
                //             if (typeof parsed === 'object' && parsed !== null) {
                //                 finalData = { ...parsed };
                //             }
                //         } catch (parseError) {
                //             finalData = {};
                //         }
                //     }
                //     finalData[month] = actualRevenue;
                //     const getFourMonthWindow = (): string[] => {
                //         const currentDate = new Date();
                //         const validMonths: string[] = [];
                //         for (let i = 0; i < 4; i++) {
                //             const date = new Date(currentDate.getFullYear(), currentDate.getMonth() - i, 1);
                //             const yearStr = date.getFullYear();
                //             const monthStr = (date.getMonth() + 1).toString().padStart(2, '0');
                //             validMonths.push(`${yearStr}-${monthStr}`);
                //         }
                //         return validMonths.sort().reverse();
                //     };
                //     const fourMonthWindow = getFourMonthWindow();
                //     const resultData: { [key: string]: number } = {};
                //     fourMonthWindow.forEach(monthKey => {
                //         resultData[monthKey] = finalData[monthKey] !== undefined ? finalData[monthKey] : 0;
                //     });
                //     return JSON.stringify(resultData);
                // };
                // File Upload Handling
                fileUploadArea.addEventListener("click", () => excelFileInput.click());
                fileUploadArea.addEventListener("dragover", (e) => {
                    e.preventDefault();
                    fileUploadArea.classList.add("dragover");
                });
                fileUploadArea.addEventListener("dragleave", () => fileUploadArea.classList.remove("dragover"));
                fileUploadArea.addEventListener("drop", (e) => {
                    var _a;
                    e.preventDefault();
                    fileUploadArea.classList.remove("dragover");
                    const files = (_a = e.dataTransfer) === null || _a === void 0 ? void 0 : _a.files;
                    if (files && files[0]) {
                        excelFileInput.files = files;
                        handleFileSelection(files[0]);
                    }
                });
                excelFileInput.addEventListener("change", (e) => {
                    const files = e.target.files;
                    if (files && files[0])
                        handleFileSelection(files[0]);
                });
                // Handle Excel File Selection
                const handleFileSelection = async (file) => {
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
                        }
                        else {
                            toast("error", "Invalid Excel format. Please check the required columns.");
                        }
                    }
                    catch (error) {
                        toast("error", "Error reading Excel file");
                    }
                };
                // Find the parseExcelFile function and update the row parsing:
                const parseExcelFile = (file) => {
                    return new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onload = (e) => {
                            var _a;
                            try {
                                const data = new Uint8Array((_a = e.target) === null || _a === void 0 ? void 0 : _a.result);
                                const workbook = XLSX.read(data, { type: "array" });
                                const sheetName = workbook.SheetNames[0];
                                const worksheet = workbook.Sheets[sheetName];
                                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                                if (jsonData.length < 2) {
                                    return resolve([]);
                                }
                                const headers = jsonData[0];
                                const rows = [];
                                for (let i = 1; i < jsonData.length; i++) {
                                    const rowData = jsonData[i];
                                    if (!rowData || rowData.length === 0)
                                        continue;
                                    const row = {};
                                    headers.forEach((header, idx) => {
                                        if (header && rowData[idx] !== undefined && rowData[idx] !== null) {
                                            row[header.trim()] = rowData[idx];
                                        }
                                    });
                                    // ✅ Use type-safe access
                                    const monthValue = row["Month"] || row["month"];
                                    if (monthValue) {
                                        if (typeof monthValue === "number") {
                                            const excelDate = new Date(Math.round((monthValue - 25569) * 86400 * 1000));
                                            const yyyy = excelDate.getFullYear();
                                            const mm = String(excelDate.getMonth() + 1).padStart(2, "0");
                                            row.parsedMonth = `${yyyy}-${mm}`;
                                        }
                                        else {
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
                            }
                            catch (err) {
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
                const validateImportData = (data) => {
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
                const showImportPreview = (data) => {
                    const monthSummary = {};
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
                    <div class="preview-table-container" style="max-height: 400px; overflow-y: auto;">
                        <table class="preview-table" style="width: 100%; border-collapse: collapse;">
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
                const performImport = async () => {
                    var _a, _b, _c, _d;
                    confirmImport.disabled = true;
                    confirmImport.textContent = "Importing...";
                    importStatus.style.display = "block";
                    importStatus.className = "import-status";
                    importStatus.textContent = "Updating Collection Data...";
                    let success = 0;
                    let error = 0;
                    let notFound = 0;
                    let invalidMonth = 0;
                    let invalidAmount = 0;
                    try {
                        // --------------------------------------------------------
                        // Helper: Normalize Project Reference
                        // --------------------------------------------------------
                        const normalizeProjectRef = (value) => {
                            return String(value !== null && value !== void 0 ? value : "")
                                .trim()
                                .replace(/\s+/g, " ")
                                .toLowerCase();
                        };
                        // --------------------------------------------------------
                        // Helper: Get Excel value safely
                        // --------------------------------------------------------
                        const getExcelValue = (row, keys) => {
                            for (const key of keys) {
                                if (row[key] !== undefined &&
                                    row[key] !== null &&
                                    String(row[key]).trim() !== "") {
                                    return row[key];
                                }
                            }
                            return null;
                        };
                        // --------------------------------------------------------
                        // Helper: Parse Actual Collection amount
                        // Handles values like:
                        // 1000
                        // "1000"
                        // "1,000"
                        // "1000.50"
                        // --------------------------------------------------------
                        const parseActualAmount = (value) => {
                            if (value === null || value === undefined || value === "") {
                                return 0;
                            }
                            const cleaned = String(value)
                                .replace(/,/g, "")
                                .replace(/\s/g, "")
                                .trim();
                            const numberValue = Number(cleaned);
                            return Number.isFinite(numberValue)
                                ? numberValue
                                : NaN;
                        };
                        // --------------------------------------------------------
                        // Helper: Safely parse JSON
                        // --------------------------------------------------------
                        const parseJsonObject = (value) => {
                            if (value === null ||
                                value === undefined ||
                                value === "" ||
                                value === "{}" ||
                                value === "null" ||
                                value === "undefined") {
                                return {};
                            }
                            if (typeof value === "object") {
                                return Array.isArray(value)
                                    ? {}
                                    : Object.assign({}, value);
                            }
                            try {
                                const parsed = JSON.parse(String(value));
                                if (parsed &&
                                    typeof parsed === "object" &&
                                    !Array.isArray(parsed)) {
                                    return Object.assign({}, parsed);
                                }
                            }
                            catch (jsonError) {
                                console.warn("⚠️ Invalid JSON found:", value, jsonError);
                            }
                            return {};
                        };
                        // --------------------------------------------------------
                        // Get SharePoint List Entity Type
                        // --------------------------------------------------------
                        const entityType = await getListEntityType();
                        if (!entityType) {
                            throw new Error("Unable to determine SharePoint List Entity Type.");
                        }
                        console.log("========== EXCEL IMPORT START ==========");
                        console.log("Total Excel records:", importData.length);
                        // --------------------------------------------------------
                        // Process every Excel row
                        // --------------------------------------------------------
                        for (const row of importData) {
                            try {
                                // ====================================================
                                // 1. GET PROJECT REFERENCE
                                // ====================================================
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
                                // ====================================================
                                // 2. GET MONTH
                                // ====================================================
                                let month = row.parsedMonth;
                                if (!month) {
                                    const rawMonth = getExcelValue(row, [
                                        "Month",
                                        "month"
                                    ]);
                                    if (rawMonth) {
                                        if (typeof rawMonth === "number") {
                                            // Excel serial date
                                            const excelDate = new Date(Math.round((rawMonth - 25569) *
                                                86400 *
                                                1000));
                                            if (!isNaN(excelDate.getTime())) {
                                                const yyyy = excelDate.getFullYear();
                                                const mm = String(excelDate.getMonth() + 1).padStart(2, "0");
                                                month = `${yyyy}-${mm}`;
                                            }
                                        }
                                        else {
                                            month = parseMonth(String(rawMonth).trim());
                                        }
                                    }
                                }
                                // Validate YYYY-MM
                                if (!month ||
                                    !/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
                                    invalidMonth++;
                                    console.warn("⚠️ Invalid month:", rawProjectRef, row);
                                    continue;
                                }
                                // ====================================================
                                // 3. GET ACTUAL COLLECTION
                                // ====================================================
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
                                    console.warn("⚠️ Invalid Actual Collection:", {
                                        projectRef: rawProjectRef,
                                        month,
                                        value: rawActual
                                    });
                                    continue;
                                }
                                // ====================================================
                                // 4. FIND PROJECT
                                // ====================================================
                                const project = projects.find((p) => normalizeProjectRef(p.ProjectReference) === projectRef);
                                if (!project) {
                                    notFound++;
                                    console.warn(`⚠️ Project not found: ${rawProjectRef}`);
                                    continue;
                                }
                                // ====================================================
                                // 5. FIND COLLECTION RECORD
                                //
                                // Existing code matches using Project.Title.
                                // Keep that behavior, but normalize it.
                                // ====================================================
                                const collection = collections.find((f) => {
                                    var _a, _b;
                                    const collectionProjectTitle = typeof f.Project === "object"
                                        ? (_a = f.Project) === null || _a === void 0 ? void 0 : _a.Title
                                        : f.Project;
                                    return (String(collectionProjectTitle !== null && collectionProjectTitle !== void 0 ? collectionProjectTitle : "")
                                        .trim()
                                        .toLowerCase() ===
                                        String((_b = project.Title) !== null && _b !== void 0 ? _b : "")
                                            .trim()
                                            .toLowerCase());
                                });
                                if (!(collection === null || collection === void 0 ? void 0 : collection.Id)) {
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
                                // ====================================================
                                // 6. GET CURRENT COLLECTION RECORD
                                //
                                // IMPORTANT:
                                // We now read BOTH RevenueActual and CollectionData.
                                // ====================================================
                                const currentCollectionResponse = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${collection.Id})?$select=Id,RevenueActual,CollectionData`, {
                                    method: "GET",
                                    headers: {
                                        Accept: "application/json;odata=verbose"
                                    }
                                });
                                if (!currentCollectionResponse.ok) {
                                    const errorText = await currentCollectionResponse.text();
                                    throw new Error(`Failed to read Collection record. HTTP ${currentCollectionResponse.status}: ${errorText}`);
                                }
                                const currentData = await currentCollectionResponse.json();
                                const existingRevenueActual = (_b = (_a = currentData === null || currentData === void 0 ? void 0 : currentData.d) === null || _a === void 0 ? void 0 : _a.RevenueActual) !== null && _b !== void 0 ? _b : "{}";
                                const existingCollectionData = (_d = (_c = currentData === null || currentData === void 0 ? void 0 : currentData.d) === null || _c === void 0 ? void 0 : _c.CollectionData) !== null && _d !== void 0 ? _d : "{}";
                                // ====================================================
                                // 7. UPDATE RevenueActual
                                //
                                // IMPORTANT:
                                // Do NOT limit to only 4 months.
                                // Preserve ALL existing months.
                                // ====================================================
                                const revenueActualObject = parseJsonObject(existingRevenueActual);
                                revenueActualObject[month] =
                                    actualRevenue;
                                const updatedRevenueActual = JSON.stringify(revenueActualObject);
                                // ====================================================
                                // 8. UPDATE CollectionData
                                // ====================================================
                                const collectionDataObject = parseJsonObject(existingCollectionData);
                                // ----------------------------------------------------
                                // Ensure month object exists
                                // ----------------------------------------------------
                                if (!collectionDataObject[month] ||
                                    typeof collectionDataObject[month] !==
                                        "object") {
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
                                // ----------------------------------------------------
                                // Preserve existing plan
                                // ----------------------------------------------------
                                if (typeof monthData.plan !== "number") {
                                    monthData.plan =
                                        Number(monthData.plan) || 0;
                                }
                                // ----------------------------------------------------
                                // Update Actual
                                // ----------------------------------------------------
                                monthData.actual =
                                    actualRevenue;
                                // ----------------------------------------------------
                                // Update Month Actual
                                // ----------------------------------------------------
                                monthData.monthActual =
                                    actualRevenue;
                                // ----------------------------------------------------
                                // Preserve forecast
                                // ----------------------------------------------------
                                if (typeof monthData.forecast !==
                                    "number") {
                                    monthData.forecast =
                                        Number(monthData.forecast) || 0;
                                }
                                // ----------------------------------------------------
                                // Preserve date
                                // ----------------------------------------------------
                                if (monthData.date === undefined ||
                                    monthData.date === null) {
                                    monthData.date = "";
                                }
                                // ----------------------------------------------------
                                // Ensure items exists
                                // ----------------------------------------------------
                                if (!monthData.items ||
                                    typeof monthData.items !== "object") {
                                    monthData.items = {};
                                }
                                const updatedCollectionData = JSON.stringify(collectionDataObject);
                                // ====================================================
                                // 9. CREATE SHAREPOINT PAYLOAD
                                // ====================================================
                                const payload = {
                                    __metadata: {
                                        type: entityType
                                    },
                                    RevenueActual: updatedRevenueActual,
                                    CollectionData: updatedCollectionData,
                                    UpdatedDate: new Date().toISOString()
                                };
                                console.log("📦 SharePoint Payload:", payload);
                                // ====================================================
                                // 10. GET REQUEST DIGEST
                                // ====================================================
                                const digest = await getDigest();
                                if (!digest) {
                                    throw new Error("SharePoint Request Digest is empty.");
                                }
                                // ====================================================
                                // 11. SAVE TO SHAREPOINT
                                // ====================================================
                                const saveResponse = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${collection.Id})`, {
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
                                // ====================================================
                                // 12. HANDLE SAVE RESPONSE
                                // ====================================================
                                if (saveResponse.ok) {
                                    success++;
                                    console.log(`✅ Successfully saved:
Project Reference = ${rawProjectRef}
Month = ${month}
Actual = ${actualRevenue}
Collection ID = ${collection.Id}`);
                                }
                                else {
                                    error++;
                                    const responseText = await saveResponse.text();
                                    console.error("❌ SharePoint SAVE FAILED", {
                                        projectReference: rawProjectRef,
                                        month,
                                        actualRevenue,
                                        collectionId: collection.Id,
                                        status: saveResponse.status,
                                        statusText: saveResponse.statusText,
                                        response: responseText,
                                        payload
                                    });
                                }
                                // Small delay to avoid SharePoint throttling
                                await new Promise((resolve) => setTimeout(resolve, 100));
                            }
                            catch (rowError) {
                                error++;
                                console.error("❌ Excel import row failed:", {
                                    row,
                                    error: rowError
                                });
                            }
                        }
                        // ============================================================
                        // 13. FINAL RESULT
                        // ============================================================
                        const totalProcessed = success +
                            error +
                            notFound +
                            invalidMonth +
                            invalidAmount;
                        const skipped = Math.max(0, importData.length -
                            totalProcessed);
                        console.log("========== EXCEL IMPORT COMPLETE ==========");
                        console.log({
                            total: importData.length,
                            success,
                            error,
                            notFound,
                            invalidMonth,
                            invalidAmount,
                            skipped
                        });
                        // ============================================================
                        // 14. SHOW RESULT
                        // ============================================================
                        if (error === 0 && notFound === 0 && invalidMonth === 0 && invalidAmount === 0) {
                            importStatus.className =
                                "import-status success";
                        }
                        else {
                            importStatus.className =
                                "import-status warning";
                        }
                        importStatus.innerHTML = `
            <strong>Excel Import Complete</strong><br>

            ✅ Success: ${success} records<br>

            ❌ Errors: ${error} records<br>

            🔍 Project Not Found: ${notFound} records<br>

            📅 Invalid Month: ${invalidMonth} records<br>

            💰 Invalid Amount: ${invalidAmount} records<br>

            ⏭️ Skipped: ${skipped} records<br>

            <small>
                RevenueActual and CollectionData were updated.
                Existing months were preserved.
            </small>
        `;
                        // ============================================================
                        // 15. RELOAD COLLECTION DATA
                        // ============================================================
                        await loadCollections();
                        initializePagination(collections);
                    }
                    catch (importError) {
                        console.error("❌ EXCEL IMPORT FAILED:", importError);
                        importStatus.className =
                            "import-status error";
                        importStatus.innerHTML = `
            <strong>Import Failed</strong><br>
            ❌ ${importError instanceof Error
                            ? importError.message
                            : String(importError)}
        `;
                    }
                    finally {
                        // ============================================================
                        // 16. RESTORE BUTTON
                        // ============================================================
                        confirmImport.disabled = false;
                        confirmImport.textContent =
                            "Import Data";
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
                const closeModalHandler = () => modal.remove();
                closeModal.addEventListener("click", closeModalHandler);
                cancelImport.addEventListener("click", closeModalHandler);
                modal.addEventListener("click", (e) => {
                    if (e.target === modal)
                        closeModalHandler();
                });
                confirmImport.addEventListener("click", performImport);
            };
            // ========== UPDATED FILTERS INITIALIZATION ==========
            const initializeFilters = () => {
                const currentYear = new Date().getFullYear();
                const currentMonth = new Date().getMonth() + 1;
                // Year filter
                filterYear.innerHTML = '<option value="">All Years</option>';
                for (let year = currentYear - 1; year <= currentYear + 2; year++) {
                    const selected = year === currentYear ? 'selected' : '';
                    filterYear.innerHTML += `<option value="${year}" ${selected}>${year}</option>`;
                }
                // Month filter
                const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                filterMonth.innerHTML = '<option value="">All Months</option>';
                for (let month = 1; month <= 12; month++) {
                    const selected = month === currentMonth ? 'selected' : '';
                    filterMonth.innerHTML += `<option value="${month}" ${selected}>${monthNames[month - 1]}</option>`;
                }
                // ✅ Get assigned projects based on user permissions
                // const assignedProjects = getAssignedProjects(projects);
                // ✅ Helper function to get unique values from assigned projects
                const getFilterValues = (key) => {
                    const values = new Set();
                    // Determine which projects to use for filters
                    const isGHHead = getUserGroups().some((g) => g === 'KRIYA_GH_Heads');
                    const isFullControl = hasFullControl();
                    const isReadAll = canReadAllProjects();
                    let filterProjects = projects;
                    // Only filter if NOT GH Head, NOT Full Control, and NOT Read All
                    if (!isGHHead && !isFullControl && !isReadAll) {
                        filterProjects = getAssignedProjects(projects);
                    }
                    filterProjects.forEach(project => {
                        let value = project[key];
                        if (value) {
                            // Handle ProjectManager which can be string or object
                            if (key === 'ProjectManager' && typeof value === 'object') {
                                value = value.Title;
                            }
                            if (typeof value === 'string' && value.trim()) {
                                values.add(value.trim());
                            }
                        }
                    });
                    return Array.from(values).sort();
                };
                // ✅ Project filter
                const filterProjectTitles = getFilterValues('Title');
                filterProject.innerHTML = '<option value="">All Projects</option>' +
                    filterProjectTitles.map(project => `<option value="${project}">${project}</option>`).join("");
                // ✅ Customer filter
                const filterCustomers = getFilterValues('CustomerName');
                filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                    filterCustomers.map(customer => `<option value="${customer}">${customer}</option>`).join("");
                // ✅ Project Manager filter
                const filterManagers = getFilterValues('ProjectManager');
                filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                    filterManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");
                // ✅ Group Head filter
                const filterGroupHeads = getFilterValues('GroupHead');
                filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
                    filterGroupHeads.map(head => `<option value="${head}">${head}</option>`).join("");
                // ✅ Project Reference filter
                const filterReferences = getFilterValues('ProjectReference');
                filterProjectReference.innerHTML = '<option value="">All References</option>' +
                    filterReferences.map(ref => `<option value="${ref}">${ref}</option>`).join("");
            };
            // Pagination button event listeners
            firstPageBtn.addEventListener("click", () => goToPage(1));
            prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
            nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
            lastPageBtn.addEventListener("click", () => goToPage(totalPages));
            const sortCollectionData = (data) => {
                if (!currentSortConfig.column) {
                    return [...data];
                }
                const sorted = [...data];
                const getProject = (item) => {
                    return projects.find(p => { var _a; return p.Title === ((_a = item.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                };
                const getMonthValue = (item, monthKey, subColumn) => {
                    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t;
                    const collectionData = safeJsonParse(item.CollectionData, {});
                    const monthData = collectionData[monthKey];
                    if (!monthData) {
                        return 0;
                    }
                    const items = monthData.items || {};
                    switch (subColumn) {
                        case "normal-previous":
                            return ((((_a = items.Normal) === null || _a === void 0 ? void 0 : _a.plan) || 0) +
                                (((_b = items["Normal Overdue"]) === null || _b === void 0 ? void 0 : _b.plan) || 0));
                        case "normal-current":
                            return ((((_c = items.Normal) === null || _c === void 0 ? void 0 : _c.actual) || 0) +
                                (((_d = items["Normal Overdue"]) === null || _d === void 0 ? void 0 : _d.actual) || 0));
                        case "retention-previous":
                            return ((((_e = items.Retention) === null || _e === void 0 ? void 0 : _e.plan) || 0) +
                                (((_f = items["Retention Overdue"]) === null || _f === void 0 ? void 0 : _f.plan) || 0));
                        case "retention-current":
                            return ((((_g = items.Retention) === null || _g === void 0 ? void 0 : _g.actual) || 0) +
                                (((_h = items["Retention Overdue"]) === null || _h === void 0 ? void 0 : _h.actual) || 0));
                        case "total-previous":
                            return ((((_j = items.Normal) === null || _j === void 0 ? void 0 : _j.plan) || 0) +
                                (((_k = items["Normal Overdue"]) === null || _k === void 0 ? void 0 : _k.plan) || 0) +
                                (((_l = items.Retention) === null || _l === void 0 ? void 0 : _l.plan) || 0) +
                                (((_m = items["Retention Overdue"]) === null || _m === void 0 ? void 0 : _m.plan) || 0));
                        case "total-current":
                            return ((((_o = items.Normal) === null || _o === void 0 ? void 0 : _o.actual) || 0) +
                                (((_p = items["Normal Overdue"]) === null || _p === void 0 ? void 0 : _p.actual) || 0) +
                                (((_q = items.Retention) === null || _q === void 0 ? void 0 : _q.actual) || 0) +
                                (((_r = items["Retention Overdue"]) === null || _r === void 0 ? void 0 : _r.actual) || 0));
                        case "actual":
                            return monthData.actual || 0;
                        case "further":
                            return ((((_s = items.Normal) === null || _s === void 0 ? void 0 : _s.plan) || 0) +
                                (((_t = items["Normal Overdue"]) === null || _t === void 0 ? void 0 : _t.plan) || 0));
                        case "forecast":
                            return monthData.plan || 0;
                        default:
                            return 0;
                    }
                };
                const getSortValue = (item) => {
                    var _a, _b, _c;
                    const project = getProject(item);
                    switch (currentSortConfig.column) {
                        case "freeze":
                            return item.Freeze ? 1 : 0;
                        case "projectReference":
                            return (project === null || project === void 0 ? void 0 : project.ProjectReference) || "";
                        case "project":
                            return ((_a = item.Project) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        case "projectManager": {
                            const manager = typeof (project === null || project === void 0 ? void 0 : project.ProjectManager) === "object"
                                ? (_b = project === null || project === void 0 ? void 0 : project.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                                : project === null || project === void 0 ? void 0 : project.ProjectManager;
                            return manager || "";
                        }
                        case "groupHead":
                            return (project === null || project === void 0 ? void 0 : project.GroupHead) || "";
                        case "customer":
                            return (project === null || project === void 0 ? void 0 : project.CustomerName) || "";
                        case "receivable":
                        case "obl":
                            return Number(item.OBL || 0);
                        case "month": {
                            const displayMonths = getDisplayMonths();
                            const monthIndex = (_c = currentSortConfig.monthIndex) !== null && _c !== void 0 ? _c : 0;
                            const monthKey = displayMonths[monthIndex];
                            if (!monthKey) {
                                return 0;
                            }
                            return getMonthValue(item, monthKey, currentSortConfig.subColumn);
                        }
                        default:
                            return "";
                    }
                };
                sorted.sort((a, b) => {
                    const valueA = getSortValue(a);
                    const valueB = getSortValue(b);
                    let comparison = 0;
                    if (typeof valueA === "number" &&
                        typeof valueB === "number") {
                        comparison = valueA - valueB;
                    }
                    else {
                        comparison =
                            String(valueA)
                                .toLowerCase()
                                .localeCompare(String(valueB).toLowerCase());
                    }
                    return currentSortConfig.direction === "asc"
                        ? comparison
                        : -comparison;
                });
                return sorted;
            };
            // ========== UPDATED FILTER TABLE ==========
            const filterTable = () => {
                const year = filterYear.value;
                const month = filterMonth.value;
                const project = filterProject.value;
                const customer = filterCustomer.value;
                const projectManager = filterProjectManager.value;
                const groupHead = filterGroupHead.value;
                const projectReference = filterProjectReference.value;
                // Get assigned project titles
                const assignedProjects = getAssignedProjects(projects);
                const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                let filtered = [...collections];
                // ✅ First filter by assigned projects (unless Full Control or GH Head)
                const isGHHead = getUserGroups().some((g) => g === 'KRIYA_GH_Heads');
                const isFullControl = hasFullControl();
                if (!isGHHead && !isFullControl) {
                    filtered = filtered.filter(collection => {
                        var _a;
                        const projectTitle = (_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        return projectTitle && assignedProjectTitles.has(projectTitle);
                    });
                }
                // Then apply other filters
                filtered = filtered.filter((collection) => {
                    var _a, _b;
                    // Year + Month
                    if (year && month) {
                        const selectedMonthKey = `${year}-${month.padStart(2, "0")}`;
                        const data = safeJsonParse(collection.CollectionData, {});
                        const hasMonthData = Object.keys(data).some((key) => /^\d{4}-\d{2}$/.test(key) && key === selectedMonthKey);
                        if (!hasMonthData) {
                            return false;
                        }
                    }
                    // Project
                    if (project && ((_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Title) !== project) {
                        return false;
                    }
                    const proj = projects.find((p) => { var _a; return p.Title === ((_a = collection.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                    // Customer
                    if (customer && (!proj || proj.CustomerName !== customer)) {
                        return false;
                    }
                    // Project Manager
                    if (projectManager) {
                        const managerName = typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object"
                            ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                            : proj === null || proj === void 0 ? void 0 : proj.ProjectManager;
                        if (!managerName || managerName !== projectManager) {
                            return false;
                        }
                    }
                    // Group Head
                    if (groupHead && (!proj || proj.GroupHead !== groupHead)) {
                        return false;
                    }
                    // Project Reference
                    if (projectReference && (!proj || proj.ProjectReference !== projectReference)) {
                        return false;
                    }
                    return true;
                });
                // Always start filtered results from page 1
                currentPage = 1;
                filteredData = filtered;
                totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
                // Update pagination controls and re-render
                updatePaginationControls();
                renderCurrentPage();
            };
            // const renderTableBodyOnlyWithTotals = (): void => {
            //   const tableBody = document.querySelector(
            //     ".collection-table-body"
            //   ) as HTMLElement;
            //   if (!tableBody) {
            //     console.error("Table body element not found");
            //     return;
            //   }
            //   const displayMonths = getDisplayMonths();
            //   const currentMonthKey = formatMonthKey(new Date());
            //   const selectedYear = filterYear.value;
            //   const selectedMonth = filterMonth.value;
            //   const selectedMonthKey =
            //     selectedYear && selectedMonth
            //       ? `${selectedYear}-${selectedMonth.padStart(2, "0")}`
            //       : "";
            //   // Empty filtered result
            //   if (filteredData.length === 0) {
            //     tableBody.innerHTML = `
            //       <tr>
            //         <td
            //           colspan="28"
            //           style="
            //             text-align:center;
            //             padding:20px;
            //             border:1px solid #ddd;
            //           "
            //         >
            //           No records found
            //         </td>
            //       </tr>
            //     `;
            //     renderTotalsHeader(
            //       displayMonths,
            //       selectedMonthKey
            //     );
            //     return;
            //   }
            //   const startIndex =
            //     (currentPage - 1) * pageSize;
            //   const endIndex =
            //     Math.min(
            //       startIndex + pageSize,
            //       filteredData.length
            //     );
            //   const pageData =
            //     filteredData.slice(
            //       startIndex,
            //       endIndex
            //     );
            //   renderTableBodyOnly(
            //     pageData,
            //     displayMonths,
            //     currentMonthKey
            //   );
            //   renderTotalsHeader(
            //     displayMonths,
            //     selectedMonthKey
            //   );
            // };
            // Enhanced clear filters function
            const clearFilters = () => {
                filterYear.value = "";
                filterMonth.value = "";
                filterProject.value = "";
                filterCustomer.value = "";
                filterProjectManager.value = "";
                filterGroupHead.value = "";
                filterProjectReference.value = "";
                currentPage = 1;
                pageSize = 10;
                pageSizeSelect.value = "10";
                currentSortConfig = {
                    column: "",
                    direction: "asc"
                };
                filteredData = [...collections];
                totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
                updatePaginationControls();
                // Do not initialize filters again.
                // Do not attach listeners again.
                renderCurrentPage();
                toast("info", "All filters and sorting cleared");
            };
            // ========== FORM SUBMISSION ==========
            form.addEventListener("submit", async (e) => {
                e.preventDefault();
                const projectId = parseInt(projectSelect.value);
                const project = projects.find(p => p.Id === projectId);
                let currentCollection;
                let isFrozenSave = false;
                if (editingId) {
                    currentCollection = collections.find(f => f.Id === editingId);
                    if (!authCanEditProject(project)) {
                        toast("error", "❌ You don't have permission to edit this collection");
                        return;
                    }
                    // ✅ Frozen => allowed, but only Further + Date are saved
                    if (currentCollection && currentCollection.Freeze) {
                        isFrozenSave = true;
                    }
                }
                else {
                    if (!canCreateRecord()) {
                        toast("error", "❌ You don't have permission to create collections");
                        return;
                    }
                }
                if (!project) {
                    toast("error", "Please select a project");
                    return;
                }
                // ✅ FROZEN PATH: save only forecast (Further) + date, keep everything else
                if (isFrozenSave && currentCollection) {
                    await saveFurtherOnly(currentCollection, project);
                    return;
                }
                const isNewProject = !editingId && !collections.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                if (!isNewProject && collections.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId && f.Id !== editingId; })) {
                    toast("error", "Collection already exists for this project!");
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
                // Add forecast validation
                if (!validateForecastData()) {
                    return;
                }
                if (!validateFormSubmission()) {
                    return;
                }
                try {
                    const listItemType = await getListItemType();
                    const digest = await getDigest();
                    // ===== GET REMARKS VALUES =====
                    const remarksInternalInput = document.querySelector('#remarksInternal');
                    const remarksSprintCellInput = document.querySelector('#remarksSprintCell');
                    const internalRemark = remarksInternalInput ? remarksInternalInput.value.trim() : "";
                    const sprintCellRemark = remarksSprintCellInput ? remarksSprintCellInput.value.trim() : "";
                    const combinedRemarks = internalRemark + "|||" + sprintCellRemark;
                    // ===== BUILD COLLECTION DATA =====
                    const collectionData = {};
                    // Get ALL months from the ideal structure (54 months total)
                    const allStorageMonths = getAllMonthsForStorage(); // This returns 54 months
                    // IMPORTANT: If editing, load existing data first to preserve all months
                    let existingCollectionData = {};
                    if (editingId) {
                        const existingCollection = collections.find(f => f.Id === editingId);
                        if (existingCollection && existingCollection.CollectionData) {
                            existingCollectionData = safeJsonParse(existingCollection.CollectionData, {});
                        }
                    }
                    // FIRST: Preserve ALL existing historical months from the database
                    // These are months 0-35 (36 historical months)
                    const idealMonths = generateIdealMonthStructure();
                    const historicalMonths = idealMonths.slice(0, 36); // First 36 months are historical
                    historicalMonths.forEach((monthKey) => {
                        // If the month exists in existing data, preserve it
                        if (existingCollectionData[monthKey]) {
                            collectionData[monthKey] = JSON.parse(JSON.stringify(existingCollectionData[monthKey]));
                        }
                        else {
                            // If it doesn't exist, create a new empty entry
                            collectionData[monthKey] = {
                                plan: 0,
                                actual: 0,
                                items: {}
                            };
                        }
                    });
                    // SECOND: Process detailed months (36-38) and long-term months (39+)
                    // These are the editable months
                    allStorageMonths.forEach((monthKey, index) => {
                        // Skip historical months (0-35) since we already handled them above
                        if (index < 36)
                            return;
                        const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                        const mainPlanInput = document.querySelector(`input.main-plan-input[data-month="${monthKey}"]`);
                        const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${monthKey}"]`);
                        const mainDateInput = document.querySelector(`input.forecast-date-input[data-month="${monthKey}"]`);
                        if (mainCurrentInput && mainPlanInput) {
                            const monthData = {
                                plan: getNumericValue(mainPlanInput),
                                actual: getNumericValue(mainCurrentInput),
                                items: {}
                            };
                            // Check if this is the current month (index 36)
                            if (index === 36) {
                                if (mainForecastInput) {
                                    monthData.forecast = getNumericValue(mainForecastInput);
                                }
                                if (mainDateInput) {
                                    monthData.date = mainDateInput.value;
                                }
                            }
                            // Detailed months (36-38): full categories
                            if (index >= 36 && index <= 38) {
                                const categories = ["Normal", "Normal Overdue", "Retention", "Retention Overdue"];
                                categories.forEach(categoryKey => {
                                    const categoryPlanInput = document.querySelector(`input.category-plan-input[data-category="${categoryKey}"][data-month="${monthKey}"]`);
                                    const categoryCurrentInput = document.querySelector(`input.category-current-input[data-category="${categoryKey}"][data-month="${monthKey}"]`);
                                    const categoryForecastInput = document.querySelector(`input.category-forecast-input[data-category="${categoryKey}"][data-month="${monthKey}"]`);
                                    const categoryDateInput = document.querySelector(`input.category-date-input[data-category="${categoryKey}"][data-month="${monthKey}"]`);
                                    if (categoryPlanInput && categoryCurrentInput) {
                                        const categoryData = {
                                            name: categoryKey,
                                            plan: getNumericValue(categoryPlanInput),
                                            actual: getNumericValue(categoryCurrentInput)
                                        };
                                        if (index === 36) {
                                            if (categoryForecastInput) {
                                                categoryData.forecast = getNumericValue(categoryForecastInput);
                                            }
                                            if (categoryDateInput) {
                                                categoryData.date = categoryDateInput.value;
                                            }
                                        }
                                        monthData.items[categoryKey] = categoryData;
                                    }
                                });
                            }
                            // Long-term months (39+): simple categories
                            else if (index >= 39) {
                                const categories = ["Normal", "Retention"];
                                categories.forEach(categoryKey => {
                                    const categoryPlanInput = document.querySelector(`input.category-plan-input[data-category="${categoryKey}"][data-month="${monthKey}"]`);
                                    const categoryCurrentInput = document.querySelector(`input.category-current-input[data-category="${categoryKey}"][data-month="${monthKey}"]`);
                                    if (categoryPlanInput && categoryCurrentInput) {
                                        monthData.items[categoryKey] = {
                                            name: categoryKey,
                                            plan: getNumericValue(categoryPlanInput),
                                            actual: getNumericValue(categoryCurrentInput)
                                        };
                                    }
                                });
                            }
                            collectionData[monthKey] = monthData;
                        }
                    });
                    // ===== PROCESS YEARLY DATA =====
                    const yearlyData = {};
                    const yearlyInputs = document.querySelectorAll('.yearly-current-input, .yearly-plan-input');
                    const yearlyYears = new Set();
                    yearlyInputs.forEach(input => {
                        const year = input.dataset.year;
                        if (year)
                            yearlyYears.add(year);
                    });
                    Array.from(yearlyYears).forEach(year => {
                        const yearlyPlanInput = document.querySelector(`input.yearly-plan-input[data-year="${year}"]`);
                        const yearlyCurrentInput = document.querySelector(`input.yearly-current-input[data-year="${year}"]`);
                        if (yearlyPlanInput && yearlyCurrentInput) {
                            yearlyData[year] = {
                                plan: getNumericValue(yearlyPlanInput),
                                actual: getNumericValue(yearlyCurrentInput),
                                items: {}
                            };
                            const categoryPlanInputs = document.querySelectorAll(`input.category-plan-input[data-year="${year}"]`);
                            categoryPlanInputs.forEach(planInput => {
                                const category = planInput.dataset.category;
                                if (category) {
                                    const currentInput = document.querySelector(`input.category-current-input[data-category="${category}"][data-year="${year}"]`);
                                    if (currentInput) {
                                        yearlyData[year].items[category] = {
                                            name: category,
                                            plan: getNumericValue(planInput),
                                            actual: getNumericValue(currentInput)
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
                    const payload = {
                        __metadata: {
                            type: listItemType
                        },
                        Title: generateCollectionTitle(project, propagatedData),
                        ProjectId: project.Id.toString(),
                        ProfitCentre: project.PC || "",
                        CollectionData: JSON.stringify(propagatedData),
                        Remarks: combinedRemarks,
                        UpdatedDate: new Date().toISOString()
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
                        toast("success", editingId ? "Collection Updated!" : "Collection Added!");
                        await loadCollections();
                        initializePagination(collections);
                        hideForm();
                        resetForm();
                    }
                    else {
                        const errorText = await res.text();
                        console.error("❌ SharePoint Error Response:", errorText);
                        try {
                            const errorData = JSON.parse(errorText);
                            if (errorData.error && errorData.error.message) {
                                toast("error", `SharePoint Error: ${errorData.error.message.value}`);
                            }
                            else {
                                toast("error", `Error saving collection: ${res.status}`);
                            }
                        }
                        catch (parseError) {
                            toast("error", `Error saving collection: ${res.status}`);
                        }
                    }
                }
                catch (error) {
                    console.error("💥 Error saving collection:", error);
                    toast("error", "Unexpected error saving collection!");
                }
            });
            const saveFurtherOnly = async (collection, project) => {
                try {
                    const listItemType = await getListItemType();
                    const digest = await getDigest();
                    // Start from the SAVED data so nothing else is touched
                    const existing = safeJsonParse(collection.CollectionData, {});
                    const currentMonthKey = formatMonthKey(new Date());
                    // --- main Further + Date for current month ---
                    const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${currentMonthKey}"]`);
                    const mainDateInput = document.querySelector(`input.forecast-date-input[data-month="${currentMonthKey}"]`);
                    let monthData = existing[currentMonthKey];
                    if (!monthData || !isCollectionMonthData(monthData)) {
                        monthData = { plan: 0, actual: 0, monthActual: 0, forecast: 0, date: '', items: {} };
                        existing[currentMonthKey] = monthData;
                    }
                    if (!monthData.items)
                        monthData.items = {};
                    if (mainForecastInput)
                        monthData.forecast = getNumericValue(mainForecastInput);
                    if (mainDateInput)
                        monthData.date = mainDateInput.value;
                    // --- category Further + Date for current month ---
                    const categories = ["Normal", "Normal Overdue", "Retention", "Retention Overdue"];
                    categories.forEach(categoryKey => {
                        const fInput = document.querySelector(`input.category-forecast-input[data-category="${categoryKey}"][data-month="${currentMonthKey}"]`);
                        const dInput = document.querySelector(`input.category-date-input[data-category="${categoryKey}"][data-month="${currentMonthKey}"]`);
                        if (!fInput && !dInput)
                            return;
                        const existingCat = monthData.items[categoryKey];
                        const cat = existingCat
                            ? Object.assign({}, existingCat) : { name: categoryKey, plan: 0, actual: 0, forecast: 0, date: '' };
                        if (fInput)
                            cat.forecast = getNumericValue(fInput);
                        if (dInput)
                            cat.date = dInput.value;
                        monthData.items[categoryKey] = cat;
                    });
                    // Remarks are also editable while frozen
                    const remarksInternalInput = document.querySelector('#remarksInternal');
                    const remarksSprintCellInput = document.querySelector('#remarksSprintCell');
                    const combinedRemarks = ((remarksInternalInput === null || remarksInternalInput === void 0 ? void 0 : remarksInternalInput.value.trim()) || "") + "|||" + ((remarksSprintCellInput === null || remarksSprintCellInput === void 0 ? void 0 : remarksSprintCellInput.value.trim()) || "");
                    const payload = {
                        __metadata: { type: listItemType },
                        CollectionData: JSON.stringify(existing), // ⚠️ NOT re-propagated
                        Remarks: combinedRemarks,
                        UpdatedDate: new Date().toISOString()
                    };
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${collection.Id})`, {
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
                        toast("success", "✅ Further & Date changes saved (project stays frozen)");
                        await loadCollections();
                        initializePagination(collections);
                        hideForm();
                        resetForm();
                    }
                    else {
                        const errorText = await res.text();
                        console.error("❌ Frozen save failed:", errorText);
                        toast("error", `Error saving further changes: ${res.status}`);
                    }
                }
                catch (err) {
                    console.error("💥 saveFurtherOnly error:", err);
                    toast("error", "Unexpected error saving further changes!");
                }
            };
            // Main forecast validation function for form submission
            const validateForecastData = () => {
                let allValid = true;
                const invalidMonths = [];
                // Current month is at index 36 in the 54-month structure
                const currentMonthKey = allMonths[36];
                if (currentMonthKey) {
                    const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${currentMonthKey}"]`);
                    if (mainForecastInput) {
                        const mainForecastValue = getNumericValue(mainForecastInput);
                        let categoriesForecastTotal = 0;
                        // Sum all category forecasts
                        const categoryForecastInputs = document.querySelectorAll(`input.category-forecast-input[data-month="${currentMonthKey}"]`);
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
            const updateMainForecastTotal = (monthKey) => {
                const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${monthKey}"]`);
                const monthActualInput = document.querySelector(`input.month-actual-input[data-month="${monthKey}"]`);
                const forecastTotalInput = document.querySelector(`input.forecast-total-input[data-month="${monthKey}"]`);
                // ---- 1. Further (main) = sum of category Further ----
                const categoryForecastInputs = document.querySelectorAll(`input.category-forecast-input[data-month="${monthKey}"]`);
                let furtherTotal = 0;
                categoryForecastInputs.forEach(input => {
                    furtherTotal += getNumericValue(input);
                });
                if (mainForecastInput) {
                    if (furtherTotal > 0) {
                        mainForecastInput.value = furtherTotal.toString();
                        mainForecastInput.classList.remove('input-empty');
                        mainForecastInput.classList.add('input-with-value');
                    }
                    else {
                        mainForecastInput.value = '';
                        mainForecastInput.classList.remove('input-with-value');
                        mainForecastInput.classList.add('input-empty');
                    }
                }
                // ---- 2. Forecast = Actual + Further ----
                const actualValue = monthActualInput ? getNumericValue(monthActualInput) : 0;
                const forecastValue = actualValue + furtherTotal;
                if (forecastTotalInput) {
                    if (forecastValue > 0) {
                        forecastTotalInput.value = forecastValue.toString();
                        forecastTotalInput.classList.remove('input-empty');
                        forecastTotalInput.classList.add('input-with-value');
                    }
                    else {
                        forecastTotalInput.value = '';
                        forecastTotalInput.classList.remove('input-with-value');
                        forecastTotalInput.classList.add('input-empty');
                    }
                }
            };
            // ========== FIXED SETUP FORECAST VALIDATION ==========
            const setupForecastValidation = () => {
                const forecastInputs = document.querySelectorAll('input.category-forecast-input');
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
                    input.setAttribute('data-old-value', getNumericValue(input).toString());
                });
                // Update Forecast total for every month
                const months = new Set();
                forecastInputs.forEach(input => {
                    if (input.dataset.month) {
                        months.add(input.dataset.month);
                    }
                });
                months.forEach(month => {
                    updateMainForecastTotal(month);
                });
            };
            // ========== FIXED HANDLE FORECAST CATEGORY CHANGE ==========
            const handleForecastCategoryChange = function () {
                const month = this.dataset.month;
                if (!month)
                    return;
                // Update input visual state
                const currentValue = getNumericValue(this);
                if (currentValue === 0) {
                    this.classList.remove('input-with-value');
                    this.classList.add('input-empty');
                }
                else {
                    this.classList.remove('input-empty');
                    this.classList.add('input-with-value');
                }
                // Store current value
                this.setAttribute('data-old-value', currentValue.toString());
                // Calculate total of all category forecast inputs
                updateMainForecastTotal(month);
            };
            // Store old value helper
            const storeOldValue = function () {
                this.setAttribute('data-old-value', getNumericValue(this).toString());
            };
            // Debug helper to check data types
            const debugDataTypes = (data, path = '') => {
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
            // In the initialize function, make sure this is correct:
            const initialize = async () => {
                try {
                    // Initialize shared authentication
                    await authManager.initialize({ siteUrl });
                    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                    console.log(`👤 User: ${getCurrentUserFullName()}`);
                    console.log(`📋 User Groups:`, getUserGroups());
                    // Load projects first
                    await loadProjects();
                    // Then load collections
                    await loadCollections();
                    console.log(`📊 Loaded ${collections.length} collections`);
                    // ✅ RUN AUTO-PROGRESSION ON PAGE LOAD (like revenue)
                    if (collections.length > 0) {
                        await applyAutoProgression();
                        // Reload after progression
                        await loadCollections();
                    }
                    // ✅ IMPORTANT: Force page size to 10 on initial load
                    pageSize = 10;
                    if (pageSizeSelect) {
                        pageSizeSelect.value = "10";
                        pageSize = parseInt(pageSizeSelect.value) || 10;
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
                    // ✅ CRITICAL FIX: Set dropdown value again
                    if (pageSizeSelect) {
                        pageSizeSelect.value = "10";
                        pageSize = 10;
                        console.log(`📊 Final page size dropdown value: ${pageSizeSelect.value}`);
                    }
                    // Update UI based on permissions
                    updateUIBasedOnPermissions({
                        importBtn: '.import-actual-btn',
                        exportBtn: '.export-btn',
                        addBtn: '.show-form-btn',
                        saveBtn: '.btn-primary[type="submit"]',
                        freezeCheckboxes: '.freeze-checkbox'
                    });
                    console.log("✅ Application initialized successfully");
                }
                catch (error) {
                    console.error("Initialization error:", error);
                    toast("error", "Failed to initialize application");
                }
            };
            // ========== EVENT LISTENERS ==========
            showFormBtn.addEventListener("click", () => {
                resetForm();
                showForm(false);
            });
            backBtn.addEventListener("click", hideForm);
            projectSelect.addEventListener("change", (e) => {
                const projectId = parseInt(e.target.value);
                if (!projectId) {
                    // No project selected, reset form
                    resetForm();
                    return;
                }
                if (editingId) {
                    toast("warning", "Cannot change project in edit mode");
                    // Revert to the original project
                    const collection = collections.find(f => f.Id === editingId);
                    if (collection === null || collection === void 0 ? void 0 : collection.Project) {
                        projectSelect.value = String(collection.Project.Id);
                    }
                    return;
                }
                const project = projects.find(p => p.Id === projectId);
                if (!project)
                    return;
                const existingCollection = collections.find(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                if (existingCollection) {
                    // Project exists - automatically switch to edit mode
                    editingId = existingCollection.Id;
                    showForm(true);
                    onProjectChange(projectId, existingCollection);
                    // Show notification instead of confirmation dialog
                    toast("info", `Editing existing collection for ${project.Title}`);
                }
                else {
                    // New project - stay in add mode
                    onProjectChange(projectId);
                }
            });
            const validateYearlyCategoryTotals = (changedInput) => {
                let allValid = true;
                const yearlyInputs = document.querySelectorAll('.yearly-current-input');
                const yearlyYears = new Set();
                yearlyInputs.forEach(input => {
                    const year = input.dataset.year;
                    if (year)
                        yearlyYears.add(year);
                });
                Array.from(yearlyYears).forEach(year => {
                    const yearValid = validateSingleYearCategories(year, changedInput);
                    if (!yearValid) {
                        allValid = false;
                    }
                });
                return allValid;
            };
            const validateSingleYearCategories = (year, changedInput) => {
                const yearlyCurrentInput = document.querySelector(`input.yearly-current-input[data-year="${year}"]`);
                const categoryCurrentInputs = document.querySelectorAll(`input.category-current-input[data-year="${year}"]`);
                if (!yearlyCurrentInput)
                    return true;
                const yearlyCurrentValue = getNumericValue(yearlyCurrentInput);
                let categoriesTotal = 0;
                // ✅ Now properly using categoryCurrentInputs
                categoryCurrentInputs.forEach(input => {
                    categoriesTotal += getNumericValue(input);
                });
                let validationDisplay = document.querySelector(`.yearly-total-validation[data-year="${year}"]`);
                if (!validationDisplay) {
                    validationDisplay = document.createElement('div');
                    validationDisplay.className = 'yearly-total-validation';
                    validationDisplay.setAttribute('data-year', year);
                    const yearBlock = yearlyCurrentInput.closest('.yearly-box');
                    if (yearBlock) {
                        const existingValidation = yearBlock.querySelector('.yearly-total-validation');
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
                }
                else if (categoriesTotal < yearlyCurrentValue) {
                    unblockAllYearCategoryInputs(year);
                    return true;
                }
                else {
                    unblockAllYearCategoryInputs(year);
                    return true;
                }
            };
            const blockIncreasingYearCategoryInputs = (year, allowedInput) => {
                const categoryInputs = document.querySelectorAll(`input.category-current-input[data-year="${year}"]`);
                categoryInputs.forEach(input => {
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
            const unblockAllYearCategoryInputs = (year) => {
                const categoryInputs = document.querySelectorAll(`input.category-current-input[data-year="${year}"]`);
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
            clearFiltersBtn.addEventListener("click", clearFilters);
            importActualBtn.addEventListener("click", importActualData);
            // Initialize application
            initialize();
        }, 100);
    }
};
//# sourceMappingURL=collection.js.map
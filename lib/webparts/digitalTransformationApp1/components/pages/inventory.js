import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/inventory.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canReadAllProjects, canCreateRecord, 
// canViewForm,
getAssignedProjects, 
// getAssignedProjectValues,
updateUIBasedOnPermissions, getDigest as getAuthDigest,
// getListEntityType
 } from "./auth";
let currentSortConfig = {
    column: '',
    direction: 'asc'
};
export default {
    content: `
<div class="inventory-page fade-in">
  <section class="table-section-inventory card">
    <div class="page-header">
      <h1>Inventory Liquidation</h1>
    </div>
    
    <div class="table-actions-inventory">
      <div class="table-actions-row-inventory">
        <div class="tableColp4-inventory">
          <label for="filterProjectInventory">Project</label>
          <select id="filterProjectInventory"><option value="">All Projects</option></select>
        </div>
        <div class="tableColp4-inventory">
          <label for="filterCustomerInventory">Customer</label>
          <select id="filterCustomerInventory"><option value="">All Customers</option></select>
        </div>
        <div class="tableColp4-inventory">
          <label for="filterProjectManagerInventory">Project Manager</label>
          <select id="filterProjectManagerInventory"><option value="">All Managers</option></select>
        </div>
        <div class="tableColp4-inventory">
          <label for="filterGroupHeadInventory">Group Head</label>
          <select id="filterGroupHeadInventory"><option value="">All Group Heads</option></select>
        </div>
        <div class="tableColp4-inventory">
          <label for="filterProjectReferenceInventory">Project Reference</label>
          <select id="filterProjectReferenceInventory"><option value="">All References</option></select>
        </div>
        <div class="tableColp4-inventory"> 
          <label> </label>
          <button class="btn-secondary-inventory clear-filters-btn-inventory" title="Clear All Filters">Reset Filter</button> 
        </div>
      </div>
      
      <div class="table-actions-row-inventory">
        <div class="tableColp4-inventory">
          <label for="filterYearInventory">Year</label>
          <select id="filterYearInventory"><option value="">All Years</option></select>
        </div>
        <div class="tableColp4-inventory">
          <label for="filterMonthInventory">Month</label>
          <select id="filterMonthInventory"><option value="">All Months</option></select>
        </div>
        <div class="tableColp4-inventory"> 
          <label> </label>
          <button class="btn-secondary-inventory export-btn-inventory">📊 Download Excel</button>
        </div> 
      </div>
    </div>
       
    <!-- ✅ SYNC MISSING PROJECTS (same as SEIBIES / revenue) - visible to Admin / GH Head / PM Manager only -->
    

    <div class="table-container-inventory">
      <div class="table-scroll-wrapper">
        <table class="month-columns-table-inventory">
          <thead id="totalsHeaderInventory" class="sticky-totals-header"></thead>
          <thead id="monthColumnsHeaderInventory" class="sticky-main-header"></thead>
          <tbody class="inventory-table-body">
            <!-- Static loading row - visible from first paint until the JS render
                 replaces it, so the table never looks blank while data loads -->
            <tr class="inventory-table-loading-row">
              <td colspan="25">
                <span class="inventory-table-loading-spinner"></span>Loading Inventory data…
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      
      <!-- ✅ DCS / Procurement-style pagination bar (same look as the other modules):
           LEFT  = « ‹ [Page X of Y pill] [page tiles] › »
           RIGHT = Show [size] per page + "Showing X-Y of N records" badge.
           All ids unchanged - the existing inventory.ts logic works as-is. -->
      <div class="pagination-container-inventory">
        <div class="pagination-row-inventory">
          <div class="pagination-left-inventory">
            <button class="pagination-btn-inventory inventory-pg-icon" id="firstPageInventory" title="First page">«</button>
            <button class="pagination-btn-inventory inventory-pg-icon" id="prevPageInventory" title="Previous page">‹</button>

            <div class="pagination-info-inventory inventory-pg-info">
              Page <span id="currentPageInventory">1</span> of <span id="totalPagesInventory">1</span>
            </div>

            <div class="pagination-pages-inventory" id="pageNumbersInventory"></div>

            <button class="pagination-btn-inventory inventory-pg-icon" id="nextPageInventory" title="Next page">›</button>
            <button class="pagination-btn-inventory inventory-pg-icon" id="lastPageInventory" title="Last page">»</button>
          </div>
        </div>

        <div class="pagination-row-inventory pagination-bottom-row-inventory">
          <div class="pagination-right-inventory">
            <div class="page-size-selector-inventory inventory-pg-size">
              <label for="pageSizeSelectInventory">Show:</label>
              <select id="pageSizeSelectInventory" class="page-size-select-inventory">
                <option value="10">10</option>
                <option value="25" selected>25</option>
                <option value="50">50</option>
                <option value="100">100</option>
              </select>
              <span>per page</span>
            </div>

            <div class="pagination-info-inventory inventory-pg-total">
              Showing <span class="inventory-pg-range"><span id="startItemInventory">0</span>-<span id="endItemInventory">0</span></span>
              of <span class="inventory-pg-count" id="totalItemsInventory">0</span> records
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>

<section class="form-section-inventory card" style="display:none;">
 <form id="inventoryForm">
  <div class="page-header">
    <h1 id="formTitleInventory">Add / Edit Inventory Forecast</h1>
    <!-- ✅ TOP SAVE BUTTON - same type="submit" + .btn-primary-inventory
         as the bottom button, so both trigger the same submit handler
         and stay in sync (visibility, saving state). The header is
         position:sticky, so this button stays visible while scrolling
         through the long forecast form. -->
    <button type="submit"
            form="inventoryForm"
            class="btn-primary-inventory btn-save-top-inventory"
            title="Save Inventory Forecast">💾 Save Inventory Forecast</button>
  </div>
 
      
      <!-- Project Selection (hidden in edit mode) -->
      <div class="form-row-inventory hidden-section">
        <div class="form-group-inventory project-select-container-inventory">
          <label>Project</label>
          <select id="projectSelectInventory" required></select>
        </div>
      </div>

      <!-- MAIN WRAPPER - 2 Column Layout -->
      <div class="project-main-wrapper-inventory">

        <!-- LEFT SIDE - Project Details (6 Rows) -->
        <div class="project-left-side-inventory" id="projectDetailsInventory">
          <div class="detail-item-inventory">
            <span class="detail-label-inventory">Project:</span>
            <span class="detail-value-inventory" id="projectNameDisplayInventory">-</span>
          </div>
          <div class="detail-item-inventory">
            <span class="detail-label-inventory">Project Ref:</span>
            <span class="detail-value-inventory" id="projectRefInventory">-</span>
          </div>
          <div class="detail-item-inventory">
            <span class="detail-label-inventory">Profit Center:</span>
            <span class="detail-value-inventory" id="profitCenterInventory">-</span>
          </div>
          <div class="detail-item-inventory">
            <span class="detail-label-inventory">Project Manager:</span>
            <span class="detail-value-inventory" id="projectManagerInventory">-</span>
          </div>
          <div class="detail-item-inventory">
            <span class="detail-label-inventory">Group Head:</span>
            <span class="detail-value-inventory" id="groupHeadInventory">-</span>
          </div>
          <div class="detail-item-inventory">
            <span class="detail-label-inventory">Customer:</span>
            <span class="detail-value-inventory" id="customerInventory">-</span>
          </div>
        </div>

        <!-- RIGHT SIDE - OBL Section -->
        <div class="project-right-side-inventory">
          <div class="obl-section-container-inventory hidden-section" id="oblContainerInventory">
            <div class="obl-content-inventory">
              <div class="obl-left-inventory">
                <div class="obl-label-group-inventory">
                  <label>Inventory</label>
                </div>
                <div class="obl-value-display-inventory">
                  <input type="text" id="oblLastMonthInventory" class="obl-input-inventory number-input-inventory" placeholder="Enter Inventory" readonly />
                </div>
              </div>
              <div class="obl-right-inventory">
                <div class="validation-summary-inventory">
                  <div class="validation-item-inventory">
                    <span class="validation-label-inventory">Total Current Value:</span>
                    <span class="validation-value-inventory" id="totalCurrentValueInventory">0</span>
                  </div>
                  <div class="validation-item-inventory">
                    <span class="validation-label-inventory">Remaining Balance:</span>
                    <span class="validation-value-inventory" id="remainingValueInventory">0</span>
                  </div>
                  <div class="validation-line-inventory">
                    <div class="line-fill-inventory" id="oblLineFillInventory"></div>
                  </div>
                </div>
              </div>
            </div>
            <div class="validation-message-inventory" id="oblValidationMessageInventory">
              ✅ All values are within Inventory limit
            </div>
          </div>
        </div>

      </div>

      <!-- Forecast Sections -->
      <div class="forecast-section-inventory hidden-section" id="historicalForecastSectionInventory">
        <div class="historical-toggle-header-inventory">
          <h3>📊 Historical Data (Last 3 Months - Readonly)</h3>
          <!-- ✅ CHANGE 1: Arrow on top + "Show / Hide" label below (matches screenshot 1) -->
          <button type="button" id="toggleHistoricalBtnInventory" class="toggle-arrow-btn-inventory" aria-expanded="false" title="Show historical data">
            <span class="toggle-arrow-inventory">▼</span>
            <span class="toggle-label-inventory">Show</span>
          </button>
        </div>
        <div id="historicalContainerInventory" class="historical-container-inventory" style="display:none;"></div>
      </div>
      
      <div class="forecast-section-inventory hidden-section" id="detailedForecastSectionInventory">
        <h3>Detailed Forecast (Current + Next 2 Months)</h3>
        <div id="firstRowContainerInventory" class="first-row-container-inventory"></div>
      </div>
      
      <div class="forecast-section-inventory hidden-section" id="longtermForecastSectionInventory">
        <h3>Long-term Forecast (Next 15 Months)</h3>
        <div id="remainingContainerInventory" class="remaining-container-inventory"></div>
      </div>

      <div class="forecast-section-inventory hidden-section" id="yearlyForecastSectionInventory">
        <h3>5-Year Forecast</h3>
        <div id="yearlyContainerInventory" class="yearly-container-inventory"></div>
      </div>
      
      <div class="form-actions-inventory">
        <button type="submit" class="btn-primary-inventory">💾 Save Inventory Forecast</button>
        <button type="button" class="btn-secondary-inventory back-to-table-inventory">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,
    /*
     * NOTE: The full stylesheet lives in ../styles/inventory.css (imported above).
     * Previously this block was a 100% duplicate of that file, which made the two
     * sources fight each other. Only the targeted UI overrides for the two requested
     * changes are kept here so they always win regardless of load order.
     */
    style: `
  /* ============================================ */
  /* CHANGE 1 - HISTORICAL TOGGLE (arrow + label) */
  /* ============================================ */
  .historical-toggle-header-inventory {
    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    cursor: pointer !important;
    padding: 10px 16px !important;
    background: #f8f9fa !important;
    border-radius: 6px !important;
    border: 1px solid #dee2e6 !important;
    margin-bottom: 12px !important;
    user-select: none !important;
  }

  .toggle-arrow-btn-inventory {
    display: inline-flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 2px !important;
    min-width: 52px !important;
    padding: 5px 10px !important;
    background: linear-gradient(180deg, #ffffff 0%, #f1f3f5 100%) !important;
    color: #495057 !important;
    border: 1px solid #ced4da !important;
    border-radius: 4px !important;
    cursor: pointer !important;
    line-height: 1 !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.06) !important;
    transition: background 0.2s ease, border-color 0.2s ease !important;
  }

  .toggle-arrow-btn-inventory:hover {
    background: #e9ecef !important;
    border-color: #adb5bd !important;
  }

  .toggle-arrow-btn-inventory:focus {
    outline: none !important;
    box-shadow: 0 0 0 2px rgba(0, 120, 212, 0.25) !important;
  }

  .toggle-arrow-inventory {
    display: block !important;
    font-size: 12px !important;
    line-height: 1 !important;
    transform: none !important;
  }

  .toggle-label-inventory {
    display: block !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    line-height: 1 !important;
    white-space: nowrap !important;
  }

  /* ============================================================ */
  /* CHANGE 2 - DETAILED FORECAST (Current + Next 2 Months) cards */
  /* ============================================================ */
  .detailed-months-row-inventory {
    display: grid !important;
    /* CHANGE - current month (2 units) is wider than each next month (1 unit):
       because current has 4 fields and each next has 2, this keeps every input
       the same width. */
    grid-template-columns: minmax(0, 2fr) minmax(0, 1fr) minmax(0, 1fr) !important;
    gap: 15px !important;
    align-items: stretch !important;
    width: 100% !important;
  }

  .current-month-container-inventory {
    min-width: 0 !important;
    display: flex !important;
  }

  .next-months-container-inventory {
    min-width: 0 !important;
    display: grid !important;
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
    gap: 15px !important;
  }

  .next-month-wrapper-inventory {
    min-width: 0 !important;
    display: flex !important;
  }

  /* Card look for all three boxes */
  .detailed-month-box-inventory {
    position: relative !important;
    width: 100% !important;
    min-width: 0 !important;
    height: 100% !important;
    box-sizing: border-box !important;
    padding: 15px !important;
    background: #fff !important;
    border: 1px solid #e0e0e0 !important;
    border-radius: 8px !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1) !important;
  }

  /* Current month = highlighted card (blue border + light tint + badge) */
  .detailed-month-box-inventory.current-month-box-inventory {
    border: 2px solid #0078d4 !important;
    background: #f5f9ff !important;
  }

  .detailed-month-box-inventory .month-header-inventory {
    text-align: center !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    color: #0078d4 !important;
    margin: 0 0 15px 0 !important;
    padding: 4px 0 8px 0 !important;
    border-bottom: 1px solid #e5e5e5 !important;
  }

  .detailed-month-box-inventory .current-month-indicator-inventory {
    position: absolute !important;
    top: 10px !important;
    right: 10px !important;
    background: #0078d4 !important;
    color: #fff !important;
    padding: 4px 8px !important;
    border-radius: 4px !important;
    font-size: 10px !important;
    font-weight: 700 !important;
  }

  .detailed-month-box-inventory .enhanced-inputs-container-inventory {
    display: grid !important;
    grid-template-columns: 1fr 1fr !important;
    gap: 14px !important;
    margin-bottom: 8px !important;
  }

  .detailed-month-box-inventory .enhanced-inputs-container-inventory.four-col-inventory {
    grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
  }

  .detailed-month-box-inventory .enhanced-input-group-inventory label {
    font-size: 11px !important;
    font-weight: 600 !important;
    color: #666 !important;
    margin-bottom: 6px !important;
  }

  .detailed-month-box-inventory .enhanced-input-group-inventory input {
    width: 100% !important;
    height: 38px !important;
    min-height: 38px !important;
    box-sizing: border-box !important;
    padding: 6px 8px !important;
    font-size: 12px !important;
    text-align: center !important;
    border: 1px solid #ddd !important;
    border-radius: 4px !important;
  }

  /* Forecast shows "0" in grey like the mock-up (not invisible) */
  .detailed-month-box-inventory .main-forecast-input-inventory.input-empty-inventory {
    color: #6c757d !important;
  }

  @media (max-width: 1200px) {
    .detailed-months-row-inventory {
      grid-template-columns: 1fr !important;
    }
  }

  @media (max-width: 768px) {
    .next-months-container-inventory {
      grid-template-columns: 1fr !important;
    }
    .detailed-month-box-inventory .enhanced-inputs-container-inventory.four-col-inventory {
      grid-template-columns: 1fr 1fr !important;
    }
  }

  /* ============================================================ */
  /* CHANGE 3 - FORM LAYOUT                                       */
  /*   historical  = 3 month boxes in a row                       */
  /*   detailed    = current + next 2 = 3 cards in a row          */
  /*   long-term   = next 15 months = rows of 5 cards             */
  /* ============================================================ */

  .historical-container-inventory {
    /* ✅ display must NOT be !important - the JS toggle sets
       container.style.display = "none" / "grid" inline (same as SEIBIES).
       An !important here would override the inline style and the historical
       boxes could never be hidden (Show/Hide + default-collapsed broken). */
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
    gap: 15px !important;
    align-items: stretch !important;
    width: 100% !important;
    overflow: visible !important;
    padding-bottom: 0 !important;
  }

  .historical-container-inventory .historical-month-inventory {
    min-width: 0 !important;
    width: auto !important;
  }

  /* Detailed: current month (2 units) + next 2 months (1 unit each) in a row.
     Wider current card matches the mock-up (2 : 1 : 1) and, because the current
     month has 4 fields while each next month has 2, keeps every input equal width. */
  .detailed-months-row-inventory {
    grid-template-columns: minmax(0, 2fr) minmax(0, 1fr) minmax(0, 1fr) !important;
  }

  /* Long-term: 15 months = rows of 5 month cards */
  .remaining-container-inventory {
    display: flex !important;
    flex-direction: column !important;
    gap: 15px !important;
    width: 100% !important;
  }

  .remaining-row-inventory {
    display: grid !important;
    grid-template-columns: repeat(5, minmax(0, 1fr)) !important;
    gap: 15px !important;
    align-items: stretch !important;
  }

  .remaining-row-inventory > .simple-month-inventory {
    min-width: 0 !important;
    height: 100% !important;
  }

  .remaining-month-title-inventory {
    display: block;
    text-align: center;
    font-weight: 600;
    font-size: 13px;
    color: #495057;
    margin-bottom: 10px;
    padding-bottom: 8px;
    border-bottom: 1px solid #e5e5e5;
  }

  .simple-month-inventory {
    background: #fff !important;
    border: 1px solid #e0e0e0 !important;
    border-radius: 8px !important;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1) !important;
    padding: 15px !important;
    box-sizing: border-box !important;
  }

  .simple-month-inventory .enhanced-inputs-container-inventory {
    display: grid !important;
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
    gap: 10px !important;
    margin-bottom: 0 !important;
  }

  .simple-month-inventory .enhanced-input-group-inventory input {
    width: 100% !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
    text-align: center !important;
    height: 36px !important;
    font-size: 12px !important;
  }

  /* ============================================================ */
  /* CHANGE 4 - SYNC BAR (same as SEIBIES / revenue)              */
  /* ============================================================ */
  .sync-bar-inventory {
    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 12px !important;
    background: #eef4fb !important;
    border: 1px solid #cfe2f3 !important;
    border-radius: 6px !important;
    padding: 8px 14px !important;
    margin: 0 0 14px 0 !important;
    font-size: 13px !important;
  }

  .sync-status-inventory {
    color: #495057 !important;
    flex: 1 !important;
    font-size: 13px !important;
  }

  .inventory-table-loading-spinner {
    display: inline-block;
    width: 14px;
    height: 14px;
    margin-right: 8px;
    border: 2px solid #cbd5e1;
    border-top-color: #0d6efd;
    border-radius: 50%;
    animation: inventoryTableSpin 0.8s linear infinite;
    vertical-align: middle;
  }
  @keyframes inventoryTableSpin { to { transform: rotate(360deg); } }

  /* Graceful wrap-down on smaller screens */
  @media (max-width: 1200px) {
    .detailed-months-row-inventory {
      grid-template-columns: minmax(0, 2fr) minmax(0, 1fr) minmax(0, 1fr) !important;
    }
    .remaining-row-inventory {
      grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
    }
  }

  @media (max-width: 900px) {
    .historical-container-inventory {
      grid-template-columns: 1fr !important;
    }
    .detailed-months-row-inventory {
      grid-template-columns: 1fr !important;
    }
    .remaining-row-inventory {
      grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
    }
  }

  @media (max-width: 560px) {
    .remaining-row-inventory {
      grid-template-columns: 1fr !important;
    }
  }
  `,
    script: () => {
        setTimeout(async () => {
            try {
                const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/DigitalTransformationTest";
                const listName = "Inventory";
                const projectListName = "Project List";
                console.log("Site URL:", siteUrl);
                // DOM Elements
                const tableBody = document.querySelector(".inventory-table-body");
                const monthColumnsHeader = document.querySelector("#monthColumnsHeaderInventory");
                const totalsHeader = document.querySelector("#totalsHeaderInventory");
                const formSection = document.querySelector(".form-section-inventory");
                const tableSection = document.querySelector(".table-section-inventory");
                const showFormBtn = document.querySelector(".show-form-btn-inventory");
                const projectSelect = document.querySelector("#projectSelectInventory");
                const historicalContainer = document.querySelector("#historicalContainerInventory");
                const firstRowContainer = document.querySelector("#firstRowContainerInventory");
                const remainingContainer = document.querySelector("#remainingContainerInventory");
                const yearlyContainer = document.querySelector("#yearlyContainerInventory");
                const backBtn = document.querySelector(".back-to-table-inventory");
                const form = document.querySelector("#inventoryForm");
                const exportBtn = document.querySelector(".export-btn-inventory");
                const oblLastMonthInput = document.querySelector("#oblLastMonthInventory");
                const totalCurrentValueSpan = document.querySelector("#totalCurrentValueInventory");
                const remainingValueSpan = document.querySelector("#remainingValueInventory");
                const oblLineFill = document.querySelector("#oblLineFillInventory");
                const oblValidationMessage = document.querySelector("#oblValidationMessageInventory");
                const formTitle = document.querySelector("#formTitleInventory");
                // Filter Elements
                const filterYear = document.querySelector("#filterYearInventory");
                const filterMonth = document.querySelector("#filterMonthInventory");
                const filterProject = document.querySelector("#filterProjectInventory");
                const filterCustomer = document.querySelector("#filterCustomerInventory");
                const filterProjectManager = document.querySelector("#filterProjectManagerInventory");
                const filterGroupHead = document.querySelector("#filterGroupHeadInventory");
                const filterProjectReference = document.querySelector("#filterProjectReferenceInventory");
                // Action Buttons
                const clearFiltersBtn = document.querySelector(".clear-filters-btn-inventory");
                // Sync Bar Elements
                const inventorySyncBar = document.querySelector("#inventorySyncBar");
                const inventorySyncStatus = document.querySelector("#inventorySyncStatus");
                const inventorySyncProjectsBtn = document.querySelector("#inventorySyncProjectsBtn");
                // Pagination Elements
                const firstPageBtn = document.querySelector("#firstPageInventory");
                const prevPageBtn = document.querySelector("#prevPageInventory");
                const nextPageBtn = document.querySelector("#nextPageInventory");
                const lastPageBtn = document.querySelector("#lastPageInventory");
                const currentPageSpan = document.querySelector("#currentPageInventory");
                const totalPagesSpan = document.querySelector("#totalPagesInventory");
                const pageNumbersDiv = document.querySelector("#pageNumbersInventory");
                const startItemSpan = document.querySelector("#startItemInventory");
                const endItemSpan = document.querySelector("#endItemInventory");
                const totalItemsSpan = document.querySelector("#totalItemsInventory");
                const pageSizeSelect = document.querySelector("#pageSizeSelectInventory");
                let projects = [];
                let forecasts = [];
                let editingId = null;
                // Pagination variables
                let currentPage = 1;
                let pageSize = 25; // ✅ default 25 per page (same as DCS / Procurement / SEIBIES)
                let totalPages = 1;
                let filteredData = [];
                // ========== PROCESSING / LOADING OVERLAY (same as SEIBIES) ==========
                let inventoryProcessingOverlay = null;
                const getInventoryProcessingOverlay = () => {
                    if (inventoryProcessingOverlay && document.body.contains(inventoryProcessingOverlay)) {
                        return inventoryProcessingOverlay;
                    }
                    const overlay = document.createElement('div');
                    overlay.id = 'inventoryProcessingOverlay';
                    overlay.style.cssText =
                        'position:fixed; top:0; left:0; right:0; bottom:0; z-index:1000;' +
                            'background:rgba(255,255,255,0.72); display:flex; align-items:center; justify-content:center;';
                    overlay.innerHTML = `
            <div style="background:#ffffff; border:1px solid #dee2e6; border-radius:10px; padding:26px 34px;
                        display:flex; align-items:center; gap:18px;
                        box-shadow:0 10px 32px rgba(0,0,0,0.16); max-width:85vw;">
              <div style="width:36px; height:36px; flex:0 0 36px; border-radius:50%;
                          border:4px solid #e2e6ea; border-top-color:#0d6efd;
                          animation:inventoryProcessingSpin 0.8s linear infinite;"></div>
              <div>
                <div style="font-size:15px; font-weight:600; color:#212529; margin-bottom:5px;">Processing…</div>
                <div id="inventoryProcessingText" style="font-size:12px; color:#6c757d; max-width:520px;">Please wait</div>
              </div>
            </div>`;
                    if (!document.getElementById('inventoryProcessingStyle')) {
                        const style = document.createElement('style');
                        style.id = 'inventoryProcessingStyle';
                        style.textContent = '@keyframes inventoryProcessingSpin { to { transform: rotate(360deg); } }';
                        document.head.appendChild(style);
                    }
                    document.body.appendChild(overlay);
                    inventoryProcessingOverlay = overlay;
                    return overlay;
                };
                // Show (or update the message of) the full-screen processing screen
                const showInventoryProcessing = (message = "Please wait") => {
                    const overlay = getInventoryProcessingOverlay();
                    const text = overlay.querySelector('#inventoryProcessingText');
                    if (text)
                        text.textContent = message;
                };
                const hideInventoryProcessing = () => {
                    if (inventoryProcessingOverlay && document.body.contains(inventoryProcessingOverlay)) {
                        inventoryProcessingOverlay.remove();
                    }
                    inventoryProcessingOverlay = null;
                };
                // ========== UTILITY FUNCTIONS ==========
                const toast = (icon, title) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
                // ============================================
                // LOADING OVERLAY (same pattern as master.ts)
                // ============================================
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
                    const sections = Array.from(document.querySelectorAll('.table-section-inventory, .form-section-inventory'));
                    const visible = sections.find(s => s.offsetParent !== null);
                    return visible || document.querySelector('.table-section-inventory') || document.body;
                };
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
                const localDateKey = (date) => {
                    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
                };
                const getMonthBounds = (date) => {
                    const first = new Date(date.getFullYear(), date.getMonth(), 1);
                    const last = new Date(date.getFullYear(), date.getMonth() + 1, 0);
                    return { min: localDateKey(first), max: localDateKey(last) };
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
                    return `${project.Title} - ${currentYear} Inventory Forecast - ${totalValue.toLocaleString()}`;
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
                const getNumericValue = (input) => {
                    if (!input)
                        return 0;
                    const value = input.value.trim();
                    if (value === '' || value === '0')
                        return 0;
                    return parseInt(value) || 0;
                };
                const setNumericValue = (input, value) => {
                    if (!input)
                        return;
                    if (value === 0) {
                        input.value = '';
                        input.className = input.className.replace('input-with-value-inventory', 'input-empty-inventory');
                    }
                    else {
                        input.value = value.toString();
                        input.className = input.className.replace('input-empty-inventory', 'input-with-value-inventory');
                    }
                };
                const allowOnlyNumbers = (input) => {
                    if (!input)
                        return;
                    input.addEventListener('input', function () {
                        let value = this.value.replace(/[^\d]/g, '');
                        if (value.length > 1 && value.startsWith('0')) {
                            value = value.replace(/^0+/, '');
                            if (value === '')
                                value = '0';
                        }
                        this.value = value;
                        if (value === '' || value === '0') {
                            this.className = this.className.replace('input-with-value-inventory', 'input-empty-inventory');
                        }
                        else {
                            this.className = this.className.replace('input-empty-inventory', 'input-with-value-inventory');
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
                // ========== AUTO-PROGRESSION FUNCTIONS (same as SEIBIES) ==========
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
                const checkIfNeedsProgression = (forecastData) => {
                    const months = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
                    if (months.length === 0)
                        return true;
                    const idealMonths = generateIdealMonthStructure();
                    const firstIdealMonth = idealMonths[0];
                    const firstDataMonth = months[0];
                    const monthlyProgressionNeeded = firstDataMonth !== firstIdealMonth;
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
                            if (checkIfNeedsProgression(originalData)) {
                                const progressedData = propagateCollectionValues(originalData);
                                const payload = {
                                    __metadata: { type: "SP.Data.InventoryListItem" },
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
                        }
                        catch (error) {
                            errorCount++;
                        }
                    }
                    if (progressedCount > 0) {
                        await loadForecasts();
                        toast("success", `Monthly and yearly data progressed for ${progressedCount} Inventory forecasts`);
                    }
                    else if (errorCount > 0) {
                        toast("warning", `Some errors occurred during Inventory monthly and yearly progression`);
                    }
                };
                // ========== DATA LOADING FUNCTIONS ==========
                // Helper function to get GroupHead display name
                const getGroupHeadDisplayName = (groupHead) => {
                    if (!groupHead)
                        return "-";
                    if (typeof groupHead === "object" && groupHead !== null) {
                        return groupHead.Title || "-";
                    }
                    return groupHead;
                };
                const loadProjects = async () => {
                    var _a, _b;
                    try {
                        console.log("Loading projects from list:", projectListName);
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
                            if (projectSelect) {
                                projectSelect.innerHTML = `<option value="">Select Project</option>` +
                                    assignedProjects.map((p) => `<option value="${p.Id}">${p.Title}</option>`).join("");
                            }
                            if (filterProject) {
                                filterProject.innerHTML = '<option value="">All Projects</option>' +
                                    [...new Set(assignedProjects.map((p) => p.Title))]
                                        .sort()
                                        .map(project => `<option value="${project}">${project}</option>`)
                                        .join("");
                            }
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
                        if (projectSelect) {
                            projectSelect.innerHTML = `<option value="">Select Project</option>` +
                                assignedProjects.map((p) => `<option value="${p.Id}">${p.Title}</option>`).join("");
                        }
                        if (filterProject) {
                            filterProject.innerHTML = '<option value="">All Projects</option>' +
                                [...new Set(assignedProjects.map((p) => p.Title))]
                                    .sort()
                                    .map(project => `<option value="${project}">${project}</option>`)
                                    .join("");
                        }
                    }
                    catch (error) {
                        console.error("Error loading projects:", error);
                        toast("error", "Failed to load projects: " + error.message);
                    }
                };
                const loadForecasts = async () => {
                    var _a;
                    try {
                        // Only mount the CARD overlay when the full-screen "Processing…"
                        // overlay is NOT already showing - otherwise two loaders appear
                        // at once during initialize() / auto-sync / auto-progression.
                        const processingVisible = !!(inventoryProcessingOverlay && document.body.contains(inventoryProcessingOverlay));
                        if (!processingVisible) {
                            const loading = createLoadingOverlay("Loading Inventory data...");
                            loaderHost().appendChild(loading);
                        }
                        console.log("Loading forecasts from list:", listName);
                        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,InventoryActual,CreatedDate,UpdatedDate&$expand=Project&$top=5000`, {
                            headers: { Accept: "application/json;odata=verbose" }
                        });
                        if (!res.ok) {
                            throw new Error(`Failed to load forecasts: ${res.status} ${res.statusText}`);
                        }
                        const responseData = await res.json();
                        // ✅ Get assigned projects
                        const assignedProjects = getAssignedProjects(projects);
                        const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                        const allForecasts = (((_a = responseData.d) === null || _a === void 0 ? void 0 : _a.results) || []).map((forecast) => (Object.assign(Object.assign({}, forecast), { ForecastData: typeof forecast.ForecastData === 'string' ?
                                forecast.ForecastData :
                                JSON.stringify(forecast.ForecastData || {}), InventoryActual: typeof forecast.InventoryActual === 'string' ?
                                forecast.InventoryActual :
                                JSON.stringify(forecast.InventoryActual || {}), Freeze: forecast.Freeze || false })));
                        // ✅ Only keep forecasts for assigned projects
                        forecasts = allForecasts.filter((forecast) => {
                            var _a;
                            const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                            return projectTitle && assignedProjectTitles.has(projectTitle);
                        });
                        console.log(`✅ Loaded ${forecasts.length} Inventory forecasts`);
                    }
                    catch (error) {
                        console.error("Error loading Inventory forecasts:", error);
                        toast("error", "Failed to load Inventory forecasts: " + error.message);
                    }
                    finally {
                        removeLoadingOverlay();
                    }
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
                        "Inventory",
                    ];
                    const selectedMonthName = getMonthName(selectedMonth);
                    const selectedMonthHeaders = [
                        `${selectedMonthName} ${selectedYear}`,
                        `${selectedMonthName} ${selectedYear}`,
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
                        monthlyHeaders.push(`${displayLabel}`, `${displayLabel}`);
                    });
                    const lastFiveYears = getLastFiveYears(exportData);
                    const yearlyHeaders = [];
                    lastFiveYears.forEach(year => {
                        yearlyHeaders.push(`${year}`, `${year}`);
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
                    if (!exportBtn)
                        return;
                    const selectedYear = filterYear ? filterYear.value : "";
                    const selectedMonth = filterMonth ? filterMonth.value : "";
                    if (!selectedYear || !selectedMonth) {
                        toast("error", "Please select year and month to export");
                        return;
                    }
                    const exportData = filteredData.length > 0 ? filteredData : forecasts;
                    if (exportData.length === 0) {
                        toast("error", "No Inventory data available to export");
                        return;
                    }
                    try {
                        const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);
                        const excelData = exportData.map(forecast => {
                            var _a, _b;
                            const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                            const proj = projects.find(p => p.Title === projectTitle);
                            const forecastData = safeJsonParse(forecast.ForecastData, {});
                            const oblValue = forecast.OBL || 0;
                            const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
                            const monthName = getMonthName(selectedMonth);
                            const selectedMonthForecastData = forecastData[selectedMonthKey] || { plan: 0, actual: 0 };
                            const selectedMonthPlan = selectedMonthForecastData.plan || 0;
                            const selectedMonthCurrent = selectedMonthForecastData.actual || 0;
                            const rowData = {
                                "Profit Centre": forecast.ProfitCentre || "-",
                                "SO/Project Ref": (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                                "Customer Name": (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                                "Project Name": projectTitle,
                                "Group Head": getGroupHeadDisplayName(proj === null || proj === void 0 ? void 0 : proj.GroupHead1),
                                "Project Manager": typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                                "Inventory": oblValue,
                                [`${monthName} ${selectedYear}`]: selectedMonthPlan,
                                [`${monthName} ${selectedYear}`]: selectedMonthCurrent,
                            };
                            allMonths.forEach(monthKey => {
                                const [year, month] = monthKey.split('-');
                                const monthLabel = getMonthName(month);
                                const displayLabel = `${monthLabel} ${year}`;
                                const forecastMonthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                                const planValue = forecastMonthData.plan || 0;
                                const currentValue = forecastMonthData.actual || 0;
                                rowData[`${displayLabel}`] = planValue;
                                rowData[`${displayLabel}`] = currentValue;
                            });
                            lastFiveYears.forEach(year => {
                                var _a;
                                const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0 };
                                rowData[`${year}`] = yearData.plan;
                                rowData[`${year}`] = yearData.actual;
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
                            if (!ws[cellAddress].s)
                                ws[cellAddress].s = {};
                            ws[cellAddress].s = headerStyle;
                        }
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
                        const colWidths = headers.map((header) => {
                            if (header.includes('Project Name'))
                                return { wch: 25 };
                            if (header.includes('Customer Name'))
                                return { wch: 20 };
                            if (header.includes('Profit Centre') || header.includes('SO/Project Ref'))
                                return { wch: 12 };
                            if (header.includes('Group Head') || header.includes('Project Manager'))
                                return { wch: 15 };
                            if (header.includes('Inventory') || header.includes('Balance'))
                                return { wch: 12 };
                            if (header.includes('(P)') || header.includes('(C)'))
                                return { wch: 10 };
                            return { wch: 12 };
                        });
                        ws['!cols'] = colWidths;
                        ws['!freeze'] = { x: 0, y: 1 };
                        XLSX.utils.book_append_sheet(wb, ws, "Inventory Forecast");
                        const monthName = getMonthName(selectedMonth);
                        const fileName = `Inventory_Forecast_${monthName}_${selectedYear}_Filtered.xlsx`;
                        XLSX.writeFile(wb, fileName);
                        toast("success", `Exported ${exportData.length} filtered Inventory records for ${monthName} ${selectedYear}`);
                    }
                    catch (error) {
                        console.error("❌ Error during styled Inventory Excel export:", error);
                        toast("error", "Failed to export Inventory data to Excel");
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
                    toast("info", `Showing ${pageSize} Inventory records per page`);
                };
                const updatePaginationControls = () => {
                    if (!currentPageSpan || !totalPagesSpan || !startItemSpan || !endItemSpan || !totalItemsSpan)
                        return;
                    currentPageSpan.textContent = currentPage.toString();
                    totalPagesSpan.textContent = totalPages.toString();
                    const startItem = (currentPage - 1) * pageSize + 1;
                    const endItem = Math.min(currentPage * pageSize, filteredData.length);
                    startItemSpan.textContent = startItem.toString();
                    endItemSpan.textContent = endItem.toString();
                    totalItemsSpan.textContent = filteredData.length.toString();
                    if (pageSizeSelect) {
                        pageSizeSelect.value = pageSize.toString();
                    }
                    if (firstPageBtn)
                        firstPageBtn.disabled = currentPage === 1 || filteredData.length === 0;
                    if (prevPageBtn)
                        prevPageBtn.disabled = currentPage === 1 || filteredData.length === 0;
                    if (nextPageBtn)
                        nextPageBtn.disabled = currentPage === totalPages || filteredData.length === 0;
                    if (lastPageBtn)
                        lastPageBtn.disabled = currentPage === totalPages || filteredData.length === 0;
                    updatePageNumbers();
                };
                const updatePageNumbers = () => {
                    if (!pageNumbersDiv)
                        return;
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
                        dotsBtn.className = 'page-number-inventory dots-inventory';
                        dotsBtn.textContent = '...';
                        dotsBtn.disabled = true;
                        pageNumbersDiv.appendChild(dotsBtn);
                    }
                    for (let i = startPage; i <= endPage; i++) {
                        const pageBtn = document.createElement('button');
                        pageBtn.className = `page-number-inventory ${i === currentPage ? 'active-inventory' : ''}`;
                        pageBtn.textContent = i.toString();
                        pageBtn.addEventListener('click', () => goToPage(i));
                        pageNumbersDiv.appendChild(pageBtn);
                    }
                    if (endPage < totalPages) {
                        const dotsBtn = document.createElement('button');
                        dotsBtn.className = 'page-number-inventory dots-inventory';
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
                    const tableContainer = document.querySelector('.table-container-inventory');
                    if (tableContainer) {
                        tableContainer.scrollTop = 0;
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
                // ✅ Sticky header offsets - totals rows + main header stack correctly
                const applyStickyHeaderOffsets = () => {
                    const table = document.querySelector('.month-columns-table-inventory');
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
                // ✅ Totals header - same structure & colours as SEIBIES / revenue
                const renderTotalsHeader = (displayMonths, currentMonthKey) => {
                    if (!totalsHeader)
                        return;
                    totalsHeader.innerHTML = '';
                    const allData = forecasts;
                    const filtered = filteredData.length ? filteredData : allData;
                    const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
                    const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);
                    const baseColumns = 6;
                    const monthColumns = displayMonths.length * 2;
                    const totalColspan = baseColumns + monthColumns;
                    const buildTotalsRow = (label, totals, css) => {
                        const monthCells = displayMonths.map(m => {
                            const t = totals.monthly[m] || { plan: 0, current: 0 };
                            return `
                <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-IN')}</td>
                <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-IN')}</td>
              `;
                        }).join('');
                        return `
              <tr class="${css}">
                <td class="project-ref-col" style="text-align:center; vertical-align:middle; min-width:120px; max-width:120px; width:120px;">${css === 'all-projects-total-inventory' ? '★' : '●'}</td>
                <td class="customer-col" style="text-align:center; vertical-align:middle; min-width:130px; max-width:130px; width:130px;"></td>
                <td class="project-col" style="text-align:left; vertical-align:middle; min-width:180px; max-width:180px; width:180px; font-weight:bold; background-color:${css === 'all-projects-total-inventory' ? '#d1ecf1' : '#d4edda'};">
                  ${label}
                </td>
                <td class="gh-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
                <td class="pm-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
                <td class="inv-col" style="text-align:right; vertical-align:middle; min-width:90px; max-width:90px; width:90px; font-weight:bold;">${totals.obl.toLocaleString('en-IN')}</td>
                ${monthCells}
              </tr>
            `;
                    };
                    let html = `
            <tr class="totals-title-row-inventory">
              <th colspan="${totalColspan}" style="background:#f8f9fa; padding:8px; text-align:center; font-size:13px; font-weight:bold; color:#495057; border:1px solid #dee2e6;">
                📊 Inventory Totals
              </th>
            </tr>
          `;
                    if (allData.length > 0) {
                        html += buildTotalsRow(`All Projects (${allData.length})`, allTotals, 'all-projects-total-inventory');
                    }
                    // Always show the Filtered row (same behaviour as before the colour update)
                    if (filtered.length > 0) {
                        html += buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total-inventory');
                    }
                    totalsHeader.innerHTML = html;
                    applyStickyHeaderOffsets();
                };
                const renderCurrentPage = () => {
                    if (filteredData.length === 0) {
                        tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No Inventory records found</td></tr>`;
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
                    // Only attach to .edit-trigger (which is now only on Project Reference)
                    document.querySelectorAll('.edit-trigger').forEach(link => {
                        link.addEventListener('click', (e) => {
                            e.preventDefault();
                            const id = parseInt(e.currentTarget.dataset.id);
                            editForecast(id);
                        });
                    });
                };
                const renderTableWithMonthColumns = (data) => {
                    if (!data.length) {
                        tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No Inventory records found</td></tr>`;
                        if (totalsHeader)
                            totalsHeader.innerHTML = '';
                        return;
                    }
                    const displayMonths = getDisplayMonths();
                    const currentMonthKey = formatMonthKey(new Date());
                    renderTotalsHeader(displayMonths, currentMonthKey);
                    // Build main table header - NEW ORDER: Project Reference, Customer, Project, Group Head, Project Manager, Inventory
                    const mainHeaderHTML = `
            <tr>
              <!-- Project Reference - FIRST - clickable -->
              <th rowspan="2" class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">
                <div class="sort-header-container-inventory">
                  <span class="header-label-inventory">Project Reference</span>
                  <div class="sort-buttons-container-inventory">
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'projectReference' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                            data-sort="projectReference" data-direction="asc" title="Sort Ascending">▲</button>
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'projectReference' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                            data-sort="projectReference" data-direction="desc" title="Sort Descending">▼</button>
                  </div>
                </div>
              </th>
              
              <!-- Customer - SECOND -->
              <th rowspan="2" class="customer-col" style="min-width:130px; max-width:130px; width:130px;">
                <div class="sort-header-container-inventory">
                  <span class="header-label-inventory">Customer</span>
                  <div class="sort-buttons-container-inventory">
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                            data-sort="customer" data-direction="asc" title="Sort Ascending">▲</button>
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                            data-sort="customer" data-direction="desc" title="Sort Descending">▼</button>
                  </div>
                </div>
              </th>
              
              <!-- Project - THIRD - NOT clickable -->
              <th rowspan="2" class="project-col" style="min-width:180px; max-width:180px; width:180px;">
                <div class="sort-header-container-inventory">
                  <span class="header-label-inventory">Project</span>
                  <div class="sort-buttons-container-inventory">
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                            data-sort="project" data-direction="asc" title="Sort Ascending">▲</button>
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                            data-sort="project" data-direction="desc" title="Sort Descending">▼</button>
                  </div>
                </div>
              </th>
              
              <!-- Group Head - FOURTH -->
              <th rowspan="2" class="gh-col" style="min-width:100px; max-width:100px; width:100px;">
                <div class="sort-header-container-inventory">
                  <span class="header-label-inventory">Group Head</span>
                  <div class="sort-buttons-container-inventory">
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                            data-sort="groupHead" data-direction="asc" title="Sort Ascending">▲</button>
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                            data-sort="groupHead" data-direction="desc" title="Sort Descending">▼</button>
                  </div>
                </div>
              </th>
              
              <!-- Project Manager - FIFTH -->
              <th rowspan="2" class="pm-col" style="min-width:100px; max-width:100px; width:100px;">
                <div class="sort-header-container-inventory">
                  <span class="header-label-inventory">Project Manager</span>
                  <div class="sort-buttons-container-inventory">
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                            data-sort="projectManager" data-direction="asc" title="Sort Ascending">▲</button>
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                            data-sort="projectManager" data-direction="desc" title="Sort Descending">▼</button>
                  </div>
                </div>
              </th>
              
              <!-- Inventory Column - SIXTH -->
              <th rowspan="2" class="inv-col" style="min-width:90px; max-width:90px; width:90px;">
                <div class="sort-header-container-inventory">
                  <span class="header-label-inventory">Inventory</span>
                  <div class="sort-buttons-container-inventory">
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                            data-sort="obl" data-direction="asc" title="Sort Ascending">▲</button>
                    <button class="sort-btn-inventory ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                            data-sort="obl" data-direction="desc" title="Sort Descending">▼</button>
                  </div>
                </div>
              </th>
              
              ${displayMonths.map((month, monthIndex) => {
                        const isCurrentMonth = month === currentMonthKey;
                        const monthLabel = formatMonthLabel(new Date(month + "-01"));
                        const displayLabel = isCurrentMonth ? `${monthLabel} (Current Month)` : monthLabel;
                        const headerClass = isCurrentMonth ? 'current-month-header-inventory' : '';
                        return `
                  <th colspan="2" class="month-scroll-col month-column-header-inventory ${headerClass}" style="min-width:60px;">
                    <div class="sort-header-container-inventory">
                      <span class="header-label-inventory">${displayLabel}</span>
                      <div class="sort-buttons-container-inventory">
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort by Month Total Ascending">▲</button>
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
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
                        const subHeaderClass = isCurrentMonth ? 'current-month-header-table-inventory' : '';
                        return `
                  <th class="month-scroll-col month-subheader-inventory ${subHeaderClass}" style="min-width:60px;">
                    <div class="sort-header-container-inventory">
                      <span>Previous</span>
                      <div class="sort-buttons-container-inventory">
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Ascending">▲</button>
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Descending">▼</button>
                      </div>
                    </div>
                  </th>
                  <th class="month-scroll-col month-subheader-inventory ${subHeaderClass}" style="min-width:60px;">
                    <div class="sort-header-container-inventory">
                      <span>Current</span>
                      <div class="sort-buttons-container-inventory">
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Ascending">▲</button>
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="desc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Descending">▼</button>
                      </div>
                    </div>
                  </th>
                `;
                    }).join('')}
            </tr>
          `;
                    if (monthColumnsHeader) {
                        monthColumnsHeader.innerHTML = mainHeaderHTML;
                    }
                    let tableBodyHTML = '';
                    data.forEach((forecast) => {
                        var _a;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-IN');
                        // Project Reference - clickable link (ONLY edit trigger here)
                        const projectRefLink = (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) ?
                            `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${proj.ProjectReference}</a>` :
                            "-";
                        // Project name - plain text (NOT clickable)
                        const projectName = projectTitle;
                        // ✅ Project Manager - Person/Group type
                        let projectManagerDisplay = "-";
                        if (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) {
                            projectManagerDisplay = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
                                ? proj.ProjectManager.Title || ""
                                : proj.ProjectManager || "";
                        }
                        // ✅ GroupHead1 - Person/Group type
                        let groupHeadDisplay = "-";
                        if (proj === null || proj === void 0 ? void 0 : proj.GroupHead1) {
                            groupHeadDisplay = typeof proj.GroupHead1 === "object" && proj.GroupHead1 !== null
                                ? proj.GroupHead1.Title || ""
                                : proj.GroupHead1 || "";
                        }
                        let rowHTML = `
              <tr data-forecast-id="${forecast.Id}">
                <!-- Project Reference - clickable link -->
                <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
                
                <!-- Customer -->
                <td class="customer-col" style="min-width:130px; max-width:130px; width:130px;">${(proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-"}</td>
                
                <!-- Project - plain text (NOT clickable) -->
                <td class="project-col" style="min-width:180px; max-width:180px; width:180px;">${projectName}</td>
                
                <!-- Group Head -->
                <td class="gh-col" style="min-width:100px; max-width:100px; width:100px;">${groupHeadDisplay}</td>
                
                <!-- Project Manager -->
                <td class="pm-col" style="min-width:100px; max-width:100px; width:100px;">${projectManagerDisplay}</td>
                
                <!-- Inventory -->
                <td class="inv-col" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
            `;
                        displayMonths.forEach((monthKey) => {
                            const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                            const planValue = monthData.plan || 0;
                            const currentValue = monthData.actual || 0;
                            rowHTML += `
                <td class="month-scroll-col" style="min-width:60px;">${planValue.toLocaleString('en-IN')}</td>
                <td class="month-scroll-col" style="min-width:60px;">${currentValue.toLocaleString('en-IN')}</td>
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
                    document.querySelectorAll('.sort-btn-inventory[data-sort]').forEach(btn => {
                        btn.addEventListener('click', (e) => {
                            const button = e.currentTarget;
                            const sortColumn = button.dataset.sort;
                            const direction = button.dataset.direction;
                            const subColumn = button.dataset.subcolumn;
                            const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;
                            if (sortColumn && direction) {
                                handleInventorySort(sortColumn, direction, subColumn, monthIndex);
                            }
                        });
                    });
                };
                const handleInventorySort = (column, direction, subColumn, monthIndex) => {
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
                            case 'projectReference':
                                const projRefA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                                const projRefB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                                valueA = (projRefA === null || projRefA === void 0 ? void 0 : projRefA.ProjectReference) || "";
                                valueB = (projRefB === null || projRefB === void 0 ? void 0 : projRefB.ProjectReference) || "";
                                break;
                            case 'customer':
                                const projA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                                const projB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                                valueA = (projA === null || projA === void 0 ? void 0 : projA.CustomerName) || "";
                                valueB = (projB === null || projB === void 0 ? void 0 : projB.CustomerName) || "";
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
                    else if (column === 'obl') {
                        columnName = 'Inventory';
                    }
                    else if (column === 'project') {
                        columnName = 'Project';
                    }
                    else if (column === 'projectReference') {
                        columnName = 'Project Reference';
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
                    if (filterYear)
                        filterYear.value = "";
                    if (filterMonth)
                        filterMonth.value = "";
                    if (filterProject)
                        filterProject.value = "";
                    if (filterCustomer)
                        filterCustomer.value = "";
                    if (filterProjectManager)
                        filterProjectManager.value = "";
                    if (filterGroupHead)
                        filterGroupHead.value = "";
                    if (filterProjectReference)
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
                    toast("info", "All Inventory filters and sorting cleared");
                };
                // ========== FILTER OPTION POPULATION ==========
                const MONTH_NAMES_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                const populateFilterOptions = () => {
                    // Year options: current year + every year present in monthly keys / yearly data
                    const yearSet = new Set();
                    yearSet.add(new Date().getFullYear().toString());
                    forecasts.forEach(f => {
                        const data = safeJsonParse(f.ForecastData, {});
                        Object.keys(data).forEach(key => {
                            if (key && /^\d{4}-\d{2}$/.test(key)) {
                                yearSet.add(key.substring(0, 4));
                            }
                        });
                        if (data.yearly && typeof data.yearly === "object") {
                            Object.keys(data.yearly).forEach(year => yearSet.add(year));
                        }
                    });
                    const sortedYears = Array.from(yearSet).sort().reverse();
                    if (filterYear) {
                        const currentSelection = filterYear.value;
                        filterYear.innerHTML = '<option value="">All Years</option>' +
                            sortedYears.map(y => `<option value="${y}">${y}</option>`).join("");
                        if (currentSelection && yearSet.has(currentSelection)) {
                            filterYear.value = currentSelection;
                        }
                    }
                    if (filterMonth) {
                        filterMonth.innerHTML = '<option value="">All Months</option>' +
                            MONTH_NAMES_SHORT.map((m, idx) => `<option value="${idx + 1}">${m}</option>`).join("");
                    }
                    if (filterCustomer) {
                        const currentSelection = filterCustomer.value;
                        filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                            [...new Set(projects.map(p => p.CustomerName).filter(Boolean))]
                                .sort()
                                .map(c => `<option value="${c}">${c}</option>`)
                                .join("");
                        if (currentSelection)
                            filterCustomer.value = currentSelection;
                    }
                    if (filterProjectManager) {
                        const currentSelection = filterProjectManager.value;
                        filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                            [...new Set(projects.map(p => {
                                    if (!p.ProjectManager)
                                        return "";
                                    return typeof p.ProjectManager === "object" && p.ProjectManager !== null
                                        ? p.ProjectManager.Title || ""
                                        : String(p.ProjectManager);
                                }).filter(Boolean))]
                                .sort()
                                .map(m => `<option value="${m}">${m}</option>`)
                                .join("");
                        if (currentSelection)
                            filterProjectManager.value = currentSelection;
                    }
                    if (filterGroupHead) {
                        const currentSelection = filterGroupHead.value;
                        filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
                            [...new Set(projects.map(p => {
                                    if (!p.GroupHead1)
                                        return "";
                                    return typeof p.GroupHead1 === "object" && p.GroupHead1 !== null
                                        ? p.GroupHead1.Title || ""
                                        : String(p.GroupHead1);
                                }).filter(Boolean))]
                                .sort()
                                .map(g => `<option value="${g}">${g}</option>`)
                                .join("");
                        if (currentSelection)
                            filterGroupHead.value = currentSelection;
                    }
                    if (filterProjectReference) {
                        const currentSelection = filterProjectReference.value;
                        filterProjectReference.innerHTML = '<option value="">All References</option>' +
                            [...new Set(projects.map(p => p.ProjectReference).filter(Boolean))]
                                .sort()
                                .map(r => `<option value="${r}">${r}</option>`)
                                .join("");
                        if (currentSelection)
                            filterProjectReference.value = currentSelection;
                    }
                };
                // ========== FILTER TABLE (same role-gated behaviour as SEIBIES) ==========
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
                    const selectedYear = filterYear ? filterYear.value : "";
                    const selectedMonth = filterMonth ? filterMonth.value : "";
                    const selectedProject = filterProject ? filterProject.value : "";
                    const selectedCustomer = filterCustomer ? filterCustomer.value : "";
                    const selectedManager = filterProjectManager ? filterProjectManager.value : "";
                    const selectedGroupHead = filterGroupHead ? filterGroupHead.value : "";
                    const selectedReference = filterProjectReference ? filterProjectReference.value : "";
                    const noFilters = !selectedYear && !selectedMonth && !selectedProject &&
                        !selectedCustomer && !selectedManager && !selectedGroupHead && !selectedReference;
                    // 🔐 Role gate (same as SEIBIES): users who cannot read ALL projects only see
                    // rows for the projects assigned to them (by Group Head / PM role).
                    let visibleSource = forecasts;
                    if (!canReadAllProjects() && projects.length > 0) {
                        const assignedProjects = getAssignedProjects(projects);
                        const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                        visibleSource = forecasts.filter(f => {
                            var _a;
                            const title = (_a = f.Project) === null || _a === void 0 ? void 0 : _a.Title;
                            return !!title && assignedProjectTitles.has(title);
                        });
                    }
                    if (noFilters) {
                        initializePagination(visibleSource);
                        return;
                    }
                    const filtered = visibleSource.filter(f => {
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
                        // No results: clear the totals band and show the empty state (same as SEIBIES)
                        if (totalsHeader)
                            totalsHeader.innerHTML = '';
                        initializePagination([]);
                        return;
                    }
                    initializePagination(filtered);
                };
                // ========== HISTORICAL TOGGLE (same as SEIBIES) ==========
                const setupHistoricalToggle = () => {
                    const toggleBtn = document.querySelector("#toggleHistoricalBtnInventory");
                    const historicalContainer = document.querySelector("#historicalContainerInventory");
                    const header = document.querySelector(".historical-toggle-header-inventory");
                    if (!toggleBtn || !historicalContainer || !header)
                        return;
                    let isVisible = false;
                    const updateHistoricalToggle = () => {
                        const arrow = toggleBtn.querySelector(".toggle-arrow-inventory");
                        const label = toggleBtn.querySelector(".toggle-label-inventory");
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
                    // Prevent duplicate handlers without replacing the button/header.
                    toggleBtn.onclick = (event) => {
                        toggleHistorical(event);
                    };
                    header.onclick = (event) => {
                        const target = event.target;
                        // Button click is already handled by the button itself.
                        if (target.closest("#toggleHistoricalBtnInventory")) {
                            return;
                        }
                        toggleHistorical(event);
                    };
                    // Historical data is collapsed by default.
                    updateHistoricalToggle();
                };
                // ========== FORM MANAGEMENT (same view rules as SEIBIES) ==========
                // 🔐 ROLE-BASED ACCESS - same rule as collection / revenue / sprint / SEIBIES:
                //   • a role may open an EXISTING forecast when it can edit it, when it is a
                //     READ_ALL role (Admin / SC Team / Controlling - opens read-only), or when
                //     the project is assigned to it (own project - read-only if not editable)
                //   • creating a NEW forecast requires create rights (FULL_CONTROL / CONTRIBUTE)
                const canViewInventoryForm = (isEdit = false, project) => {
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
                    console.log(`📋 Permission Level: ${getPermissionLevel()}`);
                    console.log(`🔑 Can Edit: ${canEditProject(project)}`);
                    console.log(`📝 Can Create: ${canCreateRecord()}`);
                    console.log(`👁️ Can View: ${canViewInventoryForm(isEditMode, project)}`);
                    if (!canViewInventoryForm(isEditMode, project)) {
                        if (!canRead()) {
                            toast("error", "❌ You don't have permission to access this form");
                            return;
                        }
                        if (isEditMode) {
                            toast("error", "❌ You don't have permission to view this forecast");
                            return;
                        }
                        toast("error", "❌ You don't have permission to create Inventory forecasts");
                        return;
                    }
                    if (!formSection || !tableSection)
                        return;
                    formSection.style.display = "block";
                    tableSection.style.display = "none";
                    if (isEditMode) {
                        formSection.classList.add('edit-mode-inventory');
                    }
                    else {
                        formSection.classList.remove('edit-mode-inventory');
                    }
                    // Project selector: hidden while editing an existing row
                    const projectSelectContainer = document.querySelector('.project-select-container-inventory');
                    if (projectSelectContainer) {
                        if (isEditMode) {
                            projectSelectContainer.classList.add('hidden-section');
                        }
                        else {
                            projectSelectContainer.classList.remove('hidden-section');
                        }
                    }
                    if (formTitle) {
                        if (isEditMode) {
                            const projectName = (project === null || project === void 0 ? void 0 : project.Title) || "";
                            const projectReference = (project === null || project === void 0 ? void 0 : project.ProjectReference) || "";
                            // ✅ "Edit Inventory Forecast - " prefix removed; show only the project name (Ref)
                            formTitle.textContent = projectName
                                ? `${projectName}${projectReference ? ` (${projectReference})` : ""}`
                                : "Edit Inventory Forecast";
                        }
                        else {
                            formTitle.textContent = "Add Inventory Forecast";
                        }
                        formTitle.style.color = '';
                    }
                    const canSave = isEditMode ? canEditProject(project) : canCreateRecord();
                    getSaveButtons().forEach(btn => {
                        btn.style.display = canSave ? '' : 'none';
                    });
                };
                const getSaveButtons = () => Array.from(document.querySelectorAll('button.btn-primary-inventory[type="submit"]'));
                const hideForm = () => {
                    if (!formSection || !tableSection)
                        return;
                    formSection.style.display = "none";
                    tableSection.style.display = "block";
                    formSection.classList.remove('edit-mode-inventory');
                    resetForm();
                };
                const resetForm = () => {
                    editingId = null;
                    form.reset();
                    projectSelect.value = "";
                    // Hide the value-only sections again until a project is chosen
                    document.querySelectorAll('.forecast-section-inventory, #oblContainerInventory').forEach(el => {
                        el.classList.add('hidden-section');
                    });
                    const projectSelectContainer = document.querySelector('.project-select-container-inventory');
                    if (projectSelectContainer)
                        projectSelectContainer.classList.remove('hidden-section');
                    // ✅ Historical cleared + collapsed (same as SEIBIES) -
                    //    setupHistoricalToggle() re-runs on the next render and resets
                    //    the arrow / label / aria state itself.
                    if (historicalContainer) {
                        historicalContainer.innerHTML = '';
                        historicalContainer.style.display = 'none';
                    }
                    if (oblValidationMessage) {
                        oblValidationMessage.textContent = '✅ All values are within Inventory limit';
                        oblValidationMessage.className = 'validation-message-inventory success-inventory';
                    }
                };
                // ========== FORM FORECAST RENDERING ==========
                const monthEditorInputHTML = (monthKey, monthData, kind, opts = {}) => {
                    const value = kind === 'plan' ? ((monthData === null || monthData === void 0 ? void 0 : monthData.plan) || 0)
                        : kind === 'actual' ? ((monthData === null || monthData === void 0 ? void 0 : monthData.actual) || 0)
                            : kind === 'forecast' ? ((monthData === null || monthData === void 0 ? void 0 : monthData.forecast) || 0) : 0;
                    const isDate = kind === 'date';
                    const dateValue = (monthData === null || monthData === void 0 ? void 0 : monthData.date) || '';
                    const baseClass = kind === 'plan' ? 'main-plan-input-inventory'
                        : kind === 'actual' ? 'main-current-input-inventory'
                            : kind === 'forecast' ? 'main-forecast-input-inventory'
                                : 'forecast-date-input-inventory';
                    const cssClass = isDate
                        ? baseClass
                        : value > 0
                            ? `${baseClass} input-with-value-inventory`
                            : `${baseClass} input-empty-inventory`;
                    if (isDate) {
                        const monthIsCurrent = monthKey === formatMonthKey(new Date());
                        const todayKey = localDateKey(new Date());
                        // Default the current month's date to today so it always displays.
                        const resolvedDate = monthIsCurrent && !dateValue ? todayKey : dateValue;
                        let minAttr = '';
                        let maxAttr = '';
                        // Only allow dates within the current month (no last/next month).
                        if (monthIsCurrent) {
                            const bounds = getMonthBounds(new Date());
                            minAttr = `min="${bounds.min}"`;
                            maxAttr = `max="${bounds.max}"`;
                        }
                        return `
              <div class="enhanced-input-group-inventory">
                <label>Date</label>
                <input type="date" class="${cssClass}" data-month="${monthKey}" data-kind="date" value="${resolvedDate}" ${minAttr} ${maxAttr} ${opts.readonly ? 'readonly' : ''} />
              </div>
            `;
                    }
                    return `
            <div class="enhanced-input-group-inventory">
              <label>${opts.label || (kind === 'plan' ? 'Previous' : kind === 'actual' ? 'Current' : 'Further')}</label>
              <input type="text" inputmode="numeric" class="${cssClass} number-input-inventory" data-month="${monthKey}" data-kind="${kind}" value="${value > 0 ? value : ''}" placeholder="${opts.placeholder || '0'}" ${opts.readonly ? 'readonly' : ''} />
            </div>
          `;
                };
                const buildMonthKeyLabel = (monthKey) => {
                    const d = new Date(monthKey + "-01");
                    return d.toLocaleString("default", { month: "short", year: "numeric" });
                };
                // ✅ Historical boxes - same structure as SEIBIES:
                //    month box + grey header band "MMM YY (Historical - Readonly)"
                //    + Previous / Current read-only inputs (+ READ ONLY watermark).
                const renderHistoricalSection = (forecastData) => {
                    if (!historicalContainer)
                        return;
                    const historicalMonths = generateIdealMonthStructure().filter((m, i) => i < 3); // last 3 months, read-only
                    let html = '';
                    historicalMonths.forEach(monthKey => {
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                        const monthLabel = formatMonthLabel(new Date(monthKey + "-01"));
                        const planValue = (monthData === null || monthData === void 0 ? void 0 : monthData.plan) || 0;
                        const actualValue = (monthData === null || monthData === void 0 ? void 0 : monthData.actual) || 0;
                        const planInputClass = planValue > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                        const actualInputClass = actualValue > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                        html += `
              <div class="month-box-inventory historical-month-inventory">
                <div class="month-header-inventory">${monthLabel} (Historical - Readonly)</div>

                <div class="month-block-inventory">
                  <div class="main-inputs-container-inventory">
                    <div class="main-input-group-inventory">
                      <label>Previous</label>
                      <input type="text" value="${planValue > 0 ? planValue : ''}" data-type="plan" data-kind="plan" data-month="${monthKey}"
                            readonly class="main-plan-input-inventory ${planInputClass} number-input-inventory" />
                    </div>
                    <div class="main-input-group-inventory">
                      <label>Current</label>
                      <input type="text" value="${actualValue > 0 ? actualValue : ''}" data-type="actual" data-kind="actual" data-month="${monthKey}"
                            readonly class="main-current-input-inventory ${actualInputClass} number-input-inventory" />
                    </div>
                  </div>
                </div>
              </div>
            `;
                    });
                    historicalContainer.innerHTML = html;
                };
                const renderDetailedMonths = (forecastData) => {
                    if (!firstRowContainer)
                        return;
                    const now = new Date();
                    const currentMonthKey = formatMonthKey(now);
                    const detailedHtml = `
            <div class="detailed-months-row-inventory">
              ${getDisplayMonths().map((monthKey, idx) => {
                        const isCurrentMonth = monthKey === currentMonthKey;
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0, forecast: 0, date: '' };
                        const subInputs = isCurrentMonth ? 'four-col-inventory' : '';
                        const extraCells = isCurrentMonth
                            ? `${monthEditorInputHTML(monthKey, monthData, 'date', {})}`
                            : '';
                        const boxHtml = `
                  <div class="detailed-month-box-inventory ${isCurrentMonth ? 'current-month-box-inventory' : ''}">
                    ${isCurrentMonth ? '<span class="current-month-indicator-inventory">Current</span>' : ''}
                    <div class="month-header-inventory">${buildMonthKeyLabel(monthKey)}</div>
                    <div class="enhanced-inputs-container-inventory ${subInputs}">
                      ${monthEditorInputHTML(monthKey, monthData, 'plan', { label: 'Previous' })}
                      ${monthEditorInputHTML(monthKey, monthData, 'actual', { label: 'Current' })}
                      ${extraCells}
                    </div>
                  </div>
                `;
                        return `
                  <div class="${idx === 0 ? 'current-month-container-inventory' : 'next-month-wrapper-inventory'}">
                    ${boxHtml}
                  </div>
                `;
                    }).join('')}
            </div>
          `;
                    firstRowContainer.innerHTML = detailedHtml;
                };
                const renderRemainingMonths = (forecastData) => {
                    if (!remainingContainer)
                        return;
                    const idealMonths = generateIdealMonthStructure();
                    // Months AFTER the detailed window (current month + next 2):
                    // exactly the next 15 months, rendered as rows of 5 cards each.
                    const remainingMonths = idealMonths.slice(6);
                    const buildMonthCard = (monthKey) => {
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                        return `
              <div class="simple-month-inventory"> 
                <div class="remaining-month-title-inventory">${buildMonthKeyLabel(monthKey)}</div>
                <div class="enhanced-inputs-container-inventory">
                  ${monthEditorInputHTML(monthKey, monthData, 'plan', { label: 'Previous' })}
                  ${monthEditorInputHTML(monthKey, monthData, 'actual', { label: 'Current' })}
                </div>
              </div>
            `;
                    };
                    let rowsHtml = '';
                    for (let i = 0; i < remainingMonths.length; i += 5) {
                        const chunk = remainingMonths.slice(i, i + 5);
                        rowsHtml += `
              <div class="remaining-row-inventory">
                ${chunk.map(monthKey => buildMonthCard(monthKey)).join('')}
              </div>
            `;
                    }
                    remainingContainer.innerHTML = rowsHtml;
                };
                const renderYearlyBoxes = (forecastData) => {
                    var _a;
                    if (!yearlyContainer)
                        return;
                    const idealMonths = generateIdealMonthStructure();
                    const lastMonthKey = idealMonths[idealMonths.length - 1];
                    const startYear = parseInt(lastMonthKey.split('-')[0], 10);
                    const endYear = startYear + 4;
                    let boxes = '';
                    for (let year = startYear; year <= endYear; year++) {
                        const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year.toString()]) || { plan: 0, actual: 0 };
                        const planValue = yearData.plan || 0;
                        const actualValue = yearData.actual || 0;
                        boxes += `
              <div class="yearly-box-inventory">
                <div class="yearly-header-inventory">${year}</div>
                <div class="yearly-inputs-container-inventory">
                  <div class="yearly-input-group-inventory">
                    <label>Previous</label>
                    <input type="text" inputmode="numeric" class="yearly-plan-input-inventory number-input-inventory ${planValue > 0 ? 'input-with-value-inventory' : 'input-empty-inventory'}" data-year="${year}" value="${planValue > 0 ? planValue : ''}" placeholder="0" />
                  </div>
                  <div class="yearly-input-group-inventory">
                    <label>Current</label>
                    <input type="text" inputmode="numeric" class="yearly-current-input-inventory number-input-inventory ${actualValue > 0 ? 'input-with-value-inventory' : 'input-empty-inventory'}" data-year="${year}" value="${actualValue > 0 ? actualValue : ''}" placeholder="0" />
                  </div>
                </div>
              </div>
            `;
                    }
                    yearlyContainer.innerHTML = boxes;
                };
                // ========== OBL / INVENTORY VALIDATION ==========
                const calculateTotalCurrentValue = () => {
                    // Current + future monthly "Current" values + yearly current values
                    let total = 0;
                    document.querySelectorAll('input.main-current-input-inventory[data-kind="actual"]').forEach(input => {
                        if (input.closest('.historical-month-inventory'))
                            return; // historical months are read-only past data
                        total += getNumericValue(input);
                    });
                    document.querySelectorAll('input.yearly-current-input-inventory').forEach(input => {
                        total += getNumericValue(input);
                    });
                    return total;
                };
                const updateOBLValidation = () => {
                    if (!oblLastMonthInput || !totalCurrentValueSpan || !remainingValueSpan || !oblLineFill || !oblValidationMessage)
                        return;
                    const oblValue = getNumericValue(oblLastMonthInput);
                    const totalCurrent = calculateTotalCurrentValue();
                    totalCurrentValueSpan.textContent = totalCurrent.toLocaleString('en-IN');
                    // ✅ Show the REAL remaining balance (goes negative when the limit is exceeded)
                    remainingValueSpan.textContent = (oblValue - totalCurrent).toLocaleString('en-IN');
                    let fillPercent = 0;
                    if (oblValue > 0) {
                        fillPercent = Math.min(100, (totalCurrent / oblValue) * 100);
                    }
                    oblLineFill.style.width = `${fillPercent}%`;
                    if (oblValue > 0 && totalCurrent > oblValue) {
                        oblValidationMessage.textContent = '⚠️ Total Current value exceeds the Inventory limit!';
                        oblValidationMessage.className = 'validation-message-inventory danger-inventory';
                        oblLineFill.style.background = '#dc3545';
                    }
                    else if (oblValue > 0 && totalCurrent > oblValue * 0.9) {
                        // ✅ Same as invoicing: near the limit -> YELLOW warning box
                        oblValidationMessage.textContent = '⚠️ Inventory limit nearly reached! Consider reducing values.';
                        oblValidationMessage.className = 'validation-message-inventory warning-inventory';
                        oblLineFill.style.background = '#ffc107';
                    }
                    else {
                        oblValidationMessage.textContent = '✅ All values are within Inventory limit';
                        oblValidationMessage.className = 'validation-message-inventory success-inventory';
                        oblLineFill.style.background = '#28a745';
                    }
                };
                const bindMonthInputListeners = () => {
                    const inputs = document.querySelectorAll('input.main-plan-input-inventory, input.main-current-input-inventory, input.main-forecast-input-inventory, input.forecast-date-input-inventory, input.yearly-plan-input-inventory, input.yearly-current-input-inventory');
                    inputs.forEach(input => {
                        allowOnlyNumbers(input);
                        input.addEventListener('input', updateOBLValidation);
                    });
                };
                const setFormReadOnly = (isReadOnly) => {
                    // ✅ EDIT SECTION RULE: only the CURRENT / FORECAST / DATE inputs follow the edit
                    //    permission. Everything "previous" (saved/database values) ALWAYS stays
                    //    readonly in the edit section - even for fully editable users:
                    //      • Previous (plan) values - monthly + yearly  → always readonly
                    //      • Inventory (OBL) value from the database    → always readonly
                    const editableSelectors = [
                        '.main-current-input-inventory',
                        '.yearly-current-input-inventory',
                        '.main-forecast-input-inventory',
                        '.forecast-date-input-inventory'
                    ];
                    // ✅ ALWAYS-READONLY fields - never unlocked by the permission check
                    const alwaysReadonlySelectors = [
                        '.main-plan-input-inventory',
                        '.yearly-plan-input-inventory'
                    ];
                    const inputs = document.querySelectorAll(editableSelectors.join(','));
                    inputs.forEach(input => {
                        if (input.closest('.historical-month-inventory'))
                            return; // historical stays readonly either way
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
                    // ✅ Previous (plan) inputs: forced readonly in BOTH branches - greyed out only
                    //    when the whole form is read-only, otherwise normal readonly styling
                    //    (the stylesheet keeps .main-plan-input-inventory / .yearly-plan-input-inventory grey).
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
                    // ✅ Inventory (OBL) value: ALWAYS readonly in the edit section - it is a saved
                    //    database value. Creating a NEW forecast never calls setFormReadOnly.
                    const oblInput = document.querySelector('#oblLastMonthInventory');
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
                };
                // ========== PROJECT CHANGE + DETAILS FILL ==========
                const fillProjectDetails = (project) => {
                    const setText = (id, value) => {
                        const el = document.getElementById(id);
                        if (el)
                            el.textContent = value || "-";
                    };
                    if (!project) {
                        setText('projectNameDisplayInventory', '-');
                        setText('projectRefInventory', '-');
                        setText('profitCenterInventory', '-');
                        setText('projectManagerInventory', '-');
                        setText('groupHeadInventory', '-');
                        setText('customerInventory', '-');
                        return;
                    }
                    setText('projectNameDisplayInventory', project.Title);
                    setText('projectRefInventory', project.ProjectReference || '-');
                    setText('profitCenterInventory', project.PC || '-');
                    setText('projectManagerInventory', displayNameOf(project.ProjectManager) || '-');
                    setText('groupHeadInventory', displayNameOf(project.GroupHead1) || '-');
                    setText('customerInventory', project.CustomerName || '-');
                };
                const onProjectChange = (projectId, existingForecast) => {
                    const project = projects.find(p => p.Id === projectId);
                    if (!project)
                        return;
                    fillProjectDetails(project);
                    const oblInput = oblLastMonthInput;
                    const oblValue = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                    if (oblInput) {
                        setNumericValue(oblInput, Math.round(oblValue));
                        oblInput.style.backgroundColor = '#f5f5f5';
                    }
                    let forecastData = createFreshForecastStructure();
                    if (existingForecast) {
                        try {
                            const parsed = safeJsonParse(existingForecast.ForecastData, {});
                            if (parsed && typeof parsed === "object" && Object.keys(parsed).length > 0) {
                                forecastData = parsed;
                            }
                        }
                        catch (e) {
                            console.warn("Could not parse existing ForecastData", e);
                        }
                    }
                    forecastData = propagateMonthValues(forecastData);
                    renderHistoricalSection(forecastData);
                    renderDetailedMonths(forecastData);
                    renderRemainingMonths(forecastData);
                    renderYearlyBoxes(forecastData);
                    bindMonthInputListeners();
                    // Reveal forecast sections + inventory (OBL) box
                    document.querySelectorAll('.forecast-section-inventory, #oblContainerInventory').forEach(el => {
                        el.classList.remove('hidden-section');
                    });
                    updateOBLValidation();
                    // Setup historical toggle after rendering - HIDE BY DEFAULT (same as SEIBIES)
                    setTimeout(() => {
                        setupHistoricalToggle();
                    }, 50);
                };
                // ========== EDIT FORECAST ==========
                const editForecast = async (id) => {
                    editingId = id;
                    const forecast = forecasts.find(f => f.Id === id);
                    if (!forecast) {
                        toast("error", "Inventory forecast not found");
                        return;
                    }
                    if (!canRead()) {
                        toast("error", "❌ You don't have permission to view this forecast");
                        return;
                    }
                    const project = projects.find(p => { var _a; return p.Id === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Id); });
                    const canEdit = canEditProject(project);
                    const isFrozen = forecast.Freeze || false;
                    // Show form FIRST (role check happens inside showForm)
                    showForm(true, project);
                    let forecastData = createFreshForecastStructure();
                    try {
                        const parsed = safeJsonParse(forecast.ForecastData, {});
                        if (parsed && typeof parsed === "object" && Object.keys(parsed).length > 0) {
                            forecastData = parsed;
                        }
                    }
                    catch (e) {
                        toast("warning", "Corrupted Inventory forecast data detected. Loaded empty structure.");
                    }
                    forecastData = propagateMonthValues(forecastData);
                    if (forecast.Project) {
                        projectSelect.value = String(forecast.Project.Id);
                        onProjectChange(forecast.Project.Id, forecast);
                    }
                    // Apply permission-based readonly state AFTER rendering
                    setTimeout(() => {
                        const isReadOnly = isFrozen || !canEdit;
                        setFormReadOnly(isReadOnly);
                        // ✅ Show/hide BOTH save buttons together (top + bottom)
                        getSaveButtons().forEach(btn => {
                            btn.style.display = isReadOnly ? 'none' : '';
                        });
                        if (formTitle) {
                            const projectName = (project === null || project === void 0 ? void 0 : project.Title) || "";
                            const projectReference = (project === null || project === void 0 ? void 0 : project.ProjectReference) || "";
                            const projectTitle = projectName
                                ? `${projectName}${projectReference ? ` (${projectReference})` : ""}`
                                : "";
                            if (isFrozen) {
                                // ✅ "Edit Inventory Forecast - " prefix removed; project name (Ref) is the title
                                formTitle.innerHTML = `🔒 ${projectTitle || "Inventory Forecast"} (Frozen - Read Only)`;
                                formTitle.style.color = '#856404';
                            }
                            else if (!canEdit) {
                                formTitle.innerHTML = `🔒 ${projectTitle || "Inventory Forecast"} (Read Only)`;
                                formTitle.style.color = '#856404';
                            }
                            else {
                                formTitle.innerHTML = projectTitle || "Edit Inventory Forecast";
                                formTitle.style.color = '';
                            }
                        }
                    }, 300);
                    toast("info", `Editing Inventory forecast for ${forecast.Title}`);
                };
                // ========== FORM SUBMISSION ==========
                form.addEventListener("submit", async (e) => {
                    e.preventDefault();
                    const projectId = parseInt(projectSelect.value);
                    const project = projects.find(p => p.Id === projectId);
                    if (!project) {
                        toast("error", "Please select a project");
                        return;
                    }
                    // 🔐 Permission checks (same as SEIBIES)
                    if (editingId) {
                        if (!canEditProject(project)) {
                            toast("error", "❌ You don't have permission to edit this Inventory forecast");
                            return;
                        }
                    }
                    else {
                        if (!canCreateRecord()) {
                            toast("error", "❌ You don't have permission to create Inventory forecasts");
                            return;
                        }
                    }
                    const isNewProject = !editingId && !forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                    if (!isNewProject && forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId && f.Id !== editingId; })) {
                        toast("error", "Inventory forecast already exists for this project!");
                        return;
                    }
                    const totalCurrent = calculateTotalCurrentValue();
                    const oblValue = getNumericValue(oblLastMonthInput);
                    if (oblValue > 0 && totalCurrent > oblValue) {
                        toast("error", "Total current value cannot exceed the Inventory limit!");
                        return;
                    }
                    // ✅ Disable BOTH save buttons while saving
                    const saveBtns = getSaveButtons();
                    saveBtns.forEach(btn => {
                        btn.disabled = true;
                        btn.dataset.saving = "true";
                    });
                    const saveLoading = createLoadingOverlay("Saving Inventory forecast...");
                    loaderHost().appendChild(saveLoading);
                    try {
                        const forecastData = {};
                        const allStorageMonths = getAllMonthsForStorage();
                        allStorageMonths.forEach(monthKey => {
                            const mainCurrentInput = document.querySelector(`input.main-current-input-inventory[data-month="${monthKey}"]`);
                            const mainPlanInput = document.querySelector(`input.main-plan-input-inventory[data-month="${monthKey}"]`);
                            const mainForecastInput = document.querySelector(`input.main-forecast-input-inventory[data-month="${monthKey}"]`);
                            const mainDateInput = document.querySelector(`input.forecast-date-input-inventory[data-month="${monthKey}"]`);
                            if (mainCurrentInput && mainPlanInput) {
                                const monthData = {
                                    plan: getNumericValue(mainPlanInput),
                                    actual: getNumericValue(mainCurrentInput)
                                };
                                if (mainForecastInput) {
                                    monthData.forecast = getNumericValue(mainForecastInput);
                                }
                                if (mainDateInput && mainDateInput.value) {
                                    const chosenDate = new Date(mainDateInput.value + "T00:00:00");
                                    const isCurrentMonth = monthKey === formatMonthKey(new Date());
                                    let dateToSave = mainDateInput.value;
                                    // Enforce current-month-only dates for the current month card.
                                    if (isCurrentMonth) {
                                        const bounds = getMonthBounds(new Date());
                                        if (chosenDate < new Date(bounds.min + "T00:00:00") || chosenDate > new Date(bounds.max + "T00:00:00")) {
                                            dateToSave = localDateKey(new Date());
                                        }
                                    }
                                    monthData.date = dateToSave;
                                }
                                forecastData[monthKey] = monthData;
                            }
                        });
                        const yearlyData = {};
                        const yearlyPlanInputs = document.querySelectorAll('.yearly-plan-input-inventory');
                        if (yearlyPlanInputs.length > 0) {
                            const years = new Set();
                            yearlyPlanInputs.forEach(input => {
                                const year = input.dataset.year;
                                if (year)
                                    years.add(year);
                            });
                            Array.from(years).sort().forEach(year => {
                                const planInput = document.querySelector(`input.yearly-plan-input-inventory[data-year="${year}"]`);
                                const currentInput = document.querySelector(`input.yearly-current-input-inventory[data-year="${year}"]`);
                                if (planInput && currentInput) {
                                    yearlyData[year] = {
                                        plan: getNumericValue(planInput),
                                        actual: getNumericValue(currentInput)
                                    };
                                }
                            });
                            forecastData.yearly = yearlyData;
                        }
                        const propagatedData = propagateMonthValues(forecastData);
                        const digest = await getDigest();
                        const payload = {
                            __metadata: { type: "SP.Data.InventoryListItem" },
                            Title: generateForecastTitle(project, propagatedData),
                            ProjectId: project.Id,
                            ProfitCentre: project.PC || "",
                            ForecastData: JSON.stringify(propagatedData),
                            UpdatedDate: new Date().toISOString()
                        };
                        if (!editingId && isNewProject) {
                            const oblValueToSave = getNumericValue(oblLastMonthInput);
                            payload.OBL = oblValueToSave;
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
                            toast("success", editingId ? "Inventory Forecast Updated!" : "Inventory Forecast Added!");
                            await loadForecasts();
                            populateFilterOptions();
                            filterTable();
                            hideForm();
                        }
                        else {
                            const errorText = await res.text();
                            console.error("SharePoint Error:", errorText);
                            toast("error", `Error saving Inventory forecast: ${res.status}`);
                        }
                    }
                    catch (error) {
                        console.error("Error saving Inventory forecast:", error);
                        toast("error", "Error saving Inventory forecast!");
                    }
                    finally {
                        removeLoadingOverlay();
                        // ✅ Re-enable BOTH save buttons
                        saveBtns.forEach(btn => {
                            btn.disabled = false;
                            btn.dataset.saving = "false";
                        });
                    }
                });
                // ========== AUTO-SYNC (same as SEIBIES) ==========
                let inventoryProjectRefField = null;
                const inventoryListProjectRefField = async () => {
                    var _a;
                    if (inventoryProjectRefField)
                        return inventoryProjectRefField;
                    try {
                        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq 'ProjectReference'&$select=InternalName,TypeAsString,ReadOnlyField`, { headers: { Accept: "application/json;odata=verbose" } });
                        const data = await res.json();
                        const results = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                        if (results.length > 0) {
                            inventoryProjectRefField = { exists: true, type: String(results[0].TypeAsString || "Text") };
                            console.log(`📋 Inventory list ProjectReference column type = ${inventoryProjectRefField.type}`);
                        }
                        else {
                            inventoryProjectRefField = { exists: false, type: "" };
                            console.log("📋 Inventory list has no ProjectReference column");
                        }
                    }
                    catch (error) {
                        console.error("⚠️ Could not check Inventory ProjectReference column:", error);
                        inventoryProjectRefField = { exists: false, type: "" };
                    }
                    return inventoryProjectRefField;
                };
                const getInventoryListEntityType = async () => {
                    try {
                        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                            headers: { Accept: "application/json;odata=verbose" }
                        });
                        const data = await res.json();
                        return data.d.ListItemEntityTypeFullName;
                    }
                    catch (error) {
                        console.error("Error getting entity type:", error);
                        // Fallback - same type constant already used by the normal Inventory save
                        return "SP.Data.InventoryListItem";
                    }
                };
                const syncMissingProjectsToInventoryList = async () => {
                    var _a, _b, _c, _d;
                    const setStatus = (text) => {
                        console.log(text);
                        if (inventorySyncStatus)
                            inventorySyncStatus.textContent = text;
                    };
                    // 🔐 Role gate (same as SEIBIES): only FULL_CONTROL / CONTRIBUTE roles can create rows
                    if (!canCreateRecord()) {
                        setStatus('👁️ No create rights - sync skipped (Admin/GH Head/PM Manager only)');
                        return;
                    }
                    if (inventorySyncProjectsBtn)
                        inventorySyncProjectsBtn.disabled = true;
                    try {
                        if (!projects || projects.length === 0) {
                            setStatus('⚠️ Project List is empty - nothing to sync');
                            return;
                        }
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
                            console.error("⚠️ Inventory sync safety-check fetch failed:", chkErr);
                        }
                        // If the list really has rows on SharePoint but the page loaded 0, the
                        // main load is broken - creating rows now would duplicate everything.
                        if (serverRows.length > 0 && forecasts.length === 0) {
                            setStatus(`⚠️ Inventory list already has ${serverRows.length} row(s) on SharePoint but the page loaded 0 - sync stopped to avoid duplicates. Check the main Inventory load.`);
                            toast("error", `Sync stopped: list has ${serverRows.length} rows but page loaded 0`);
                            return;
                        }
                        const existsOnServer = (p) => serverRows.some(c => Number(c.ProjectId) === p.Id);
                        const existsInMemory = (p) => forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === p.Id || (f.Project && f.Project.Title === p.Title); });
                        const missingProjects = visibleProjects.filter(p => !existsOnServer(p) && !existsInMemory(p));
                        if (missingProjects.length === 0) {
                            // setStatus(`✅ In sync: all ${visibleProjects.length} visible project(s) have an Inventory row`);
                            return;
                        }
                        setStatus(`🆕 ${missingProjects.length} missing: ${missingProjects.map(p => p.Title).join(', ')}`);
                        toast("info", `Checking ${visibleProjects.length} projects... adding ${missingProjects.length} missing`);
                        const digest = await getDigest();
                        const entityType = await getInventoryListEntityType();
                        let added = 0;
                        let failed = 0;
                        const errors = [];
                        for (const project of missingProjects) {
                            let ok = false;
                            // Up to 3 payload variants (same fallback chain as SEIBIES):
                            //   1 = full create - Project lookup id as STRING + ProjectReference (when the
                            //       column exists: Lookup -> ProjectReferenceId, Text -> text) + OBL as STRING
                            //   2 = without ProjectReference (column may be read-only/misconfigured)
                            //   3 = without OBL too (column may be read-only/absent)
                            for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
                                try {
                                    const freshData = createFreshForecastStructure();
                                    const payload = {
                                        __metadata: { type: entityType },
                                        Title: generateForecastTitle(project, freshData),
                                        ProjectId: project.Id,
                                        ProfitCentre: project.PC || "",
                                        ForecastData: JSON.stringify(freshData),
                                        // Same OBL default rule as the SEIBIES sync (OBL becomes readonly after
                                        // creation, so a 0 here would leave the row unusable)
                                        OBL: project.PC ? 0 : 0,
                                        CreatedDate: new Date().toISOString(),
                                        UpdatedDate: new Date().toISOString()
                                    };
                                    // Store Project Reference when the column exists - same rule as the
                                    // Revenue / SEIBIES sync: a LOOKUP column receives the target item ID
                                    // as a string (ProjectReferenceId), a text column receives the text.
                                    const refField = await inventoryListProjectRefField();
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
                            await loadForecasts();
                            populateFilterOptions();
                            filterTable();
                            setStatus(`✅ Added ${added} missing Inventory row(s)` + (failed > 0 ? ` - ${failed} failed` : ""));
                            console.log(`➕ Inventory sync added ${added} project(s)` + (failed > 0 ? ` (${failed} failed)` : ""));
                        }
                        else if (failed > 0) {
                            setStatus(`❌ Sync failed for ${failed} project(s) - see console`);
                            console.error(`❌ Inventory sync failed for ${failed} project(s): ${errors.join(' | ')}`);
                        }
                    }
                    catch (error) {
                        console.error("❌ Inventory sync failed:", error);
                        setStatus("❌ Sync failed - see console");
                    }
                    finally {
                        if (inventorySyncProjectsBtn)
                            inventorySyncProjectsBtn.disabled = false;
                    }
                };
                // ========== INITIALIZE ==========
                const initialize = async () => {
                    try {
                        console.log("Initializing Inventory application…");
                        // ✅ Processing screen from the very start - the table body is empty at this
                        // point, so the user must never stare at a blank table while data loads.
                        showInventoryProcessing("Loading Inventory application…");
                        // ✅ Initialize shared authentication
                        await authManager.initialize({ siteUrl });
                        console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                        console.log(`👤 User: ${getCurrentUserFullName()}`);
                        console.log(`📋 User Groups:`, getUserGroups());
                        console.log(`✅ Has Full Control: ${hasFullControl()}`);
                        console.log(`✅ Can Contribute: ${canContribute()}`);
                        console.log(`✅ Can Freeze: ${canFreeze()}`);
                        console.log(`✅ Is GH Head: ${isGHHead()}`);
                        showInventoryProcessing("Loading Project List…");
                        await loadProjects();
                        showInventoryProcessing("Loading Inventory forecasts…");
                        await loadForecasts();
                        // 🔐 AUTO-SYNC (same as SEIBIES): compare Project List with the Inventory List
                        // and add any missing project row (empty ForecastData skeleton). Role-gated
                        // inside the function - runs silently for read-only roles, only creates rows
                        // within this role's visible project scope.
                        showInventoryProcessing("Syncing Inventory with Project List – adding missing rows…");
                        await syncMissingProjectsToInventoryList();
                        // ✅ AUTO-PROGRESSION (same as SEIBIES): roll monthly + yearly data forward,
                        // then reload so the table shows the progressed figures on page open.
                        if (forecasts.length > 0) {
                            showInventoryProcessing("Checking monthly / yearly progression…");
                            await applyAutoProgression();
                            showInventoryProcessing("Reloading Inventory forecasts…");
                            await loadForecasts();
                        }
                        // 🔐 Apply visible scope + filters + totals + pagination
                        showInventoryProcessing("Building Inventory table…");
                        populateFilterOptions();
                        filterTable();
                        // ✅ Update UI based on permissions (export hidden when read-only etc.)
                        updateUIBasedOnPermissions({
                            exportBtn: '.export-btn-inventory',
                            saveBtn: '.btn-primary-inventory[type="submit"]'
                        });
                        // ✅ Sync bar (same as revenue / SEIBIES): only Admin / GH Head / PM Manager can sync
                        if (inventorySyncBar) {
                            inventorySyncBar.style.display = canCreateRecord() ? 'flex' : 'none';
                        }
                        console.log("✅ Inventory application initialized successfully");
                        hideInventoryProcessing();
                    }
                    catch (error) {
                        hideInventoryProcessing();
                        console.error("Failed to initialize Inventory application:", error);
                        toast("error", "Failed to initialize Inventory application");
                    }
                    finally {
                        removeLoadingOverlay();
                    }
                };
                // ========== EVENT LISTENERS ==========
                if (showFormBtn) {
                    showFormBtn.addEventListener("click", () => {
                        resetForm();
                        showForm(false);
                    });
                }
                if (backBtn)
                    backBtn.addEventListener("click", hideForm);
                if (projectSelect) {
                    projectSelect.addEventListener("change", (e) => {
                        if (editingId) {
                            toast("warning", "Cannot change project in Inventory edit mode");
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
                                        title: "Existing Inventory Forecast Found",
                                        text: `An Inventory forecast already exists for ${project.Title}. Would you like to edit it?`,
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
                }
                if (firstPageBtn)
                    firstPageBtn.addEventListener("click", () => goToPage(1));
                if (prevPageBtn)
                    prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
                if (nextPageBtn)
                    nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
                if (lastPageBtn)
                    lastPageBtn.addEventListener("click", () => goToPage(totalPages));
                if (pageSizeSelect) {
                    pageSizeSelect.addEventListener("change", () => {
                        changePageSize(parseInt(pageSizeSelect.value, 10) || 25);
                    });
                }
                if (filterYear)
                    filterYear.addEventListener("change", filterTable);
                if (filterMonth)
                    filterMonth.addEventListener("change", filterTable);
                if (filterProject)
                    filterProject.addEventListener("change", filterTable);
                if (filterCustomer)
                    filterCustomer.addEventListener("change", filterTable);
                if (filterProjectManager)
                    filterProjectManager.addEventListener("change", filterTable);
                if (filterGroupHead)
                    filterGroupHead.addEventListener("change", filterTable);
                if (filterProjectReference)
                    filterProjectReference.addEventListener("change", filterTable);
                if (exportBtn)
                    exportBtn.addEventListener("click", exportToExcelWithFullStyle);
                if (clearFiltersBtn)
                    clearFiltersBtn.addEventListener("click", clearFilters);
                // ✅ Sync Projects button (same as revenue / SEIBIES) - manual, never runs automatically
                if (inventorySyncProjectsBtn) {
                    inventorySyncProjectsBtn.addEventListener("click", async () => {
                        showInventoryProcessing("Syncing Inventory with Project List – adding missing rows…");
                        try {
                            await syncMissingProjectsToInventoryList();
                        }
                        finally {
                            hideInventoryProcessing();
                        }
                    });
                }
                // ✅ Historical toggle is wired inside setupHistoricalToggle()
                //    (onclick assignment instead of addEventListener to avoid duplicate handlers,
                //    exactly like SEIBIES - collapsed by default, re-initialised after each render)
                // ✅ Keep sticky header offsets correct on resize (same as SEIBIES)
                window.addEventListener('resize', () => {
                    clearTimeout(window._inventoryStickyResizeTimer);
                    window._inventoryStickyResizeTimer = setTimeout(applyStickyHeaderOffsets, 150);
                });
                // ✅ Start the app (runs auto-sync + auto-progression inside)
                initialize();
            }
            catch (error) {
                console.error("Failed to start Inventory application:", error);
            }
        }, 100);
    }
};
//# sourceMappingURL=inventory.js.map
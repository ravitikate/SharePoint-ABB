import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/inventory.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canViewForm, getAssignedProjects, getAssignedProjectValues, updateUIBasedOnPermissions, getDigest as getAuthDigest,
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
    <div class="page-header-inventory">
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
          <label>&nbsp;</label>
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
          <label>&nbsp;</label>
          <button class="btn-secondary-inventory export-btn-inventory">📊 Export Excel</button>
        </div> 
      </div>
    </div>
       
    <div class="table-container-inventory">
      <div style="overflow-x:auto;">
        <table class="month-columns-table-inventory">
          <thead id="totalsHeaderInventory"></thead>
          <thead id="monthColumnsHeaderInventory"></thead>
          <tbody class="inventory-table-body"></tbody>
        </table>
      </div>
      
      <div class="pagination-container-inventory">
        <div class="pagination-left-inventory">
          <button class="pagination-btn-inventory" id="firstPageInventory" disabled>First</button>
          <button class="pagination-btn-inventory" id="prevPageInventory" disabled>Previous</button>
          
          <div class="pagination-info-inventory">
            Page <span id="currentPageInventory">1</span> of <span id="totalPagesInventory">1</span>
          </div>
          
          <div class="pagination-pages-inventory" id="pageNumbersInventory"></div>
          
          <button class="pagination-btn-inventory" id="nextPageInventory" disabled>Next</button>
          <button class="pagination-btn-inventory" id="lastPageInventory" disabled>Last</button>
        </div>
        
        <div class="pagination-right-inventory">
          <div class="page-size-selector-inventory">
            <label for="pageSizeSelectInventory">Show:</label>
            <select id="pageSizeSelectInventory" class="page-size-select-inventory">
              <option value="10">10</option>
              <option value="25">25</option>
              <option value="50">50</option>
              <option value="100">100</option>
            </select>
            <span>per page</span>
          </div>
          
          <div class="pagination-info-inventory">
            Showing <span id="startItemInventory">0</span>-<span id="endItemInventory">0</span> of <span id="totalItemsInventory">0</span> records
          </div>
        </div>
      </div>
    </div>
  </section>

  <section class="form-section-inventory card" style="display:none;">
    <div class="page-header-inventory">
      <h1 id="formTitleInventory">Add / Edit Inventory Forecast</h1>
    </div>
    <form id="inventoryForm">
      
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
          <button type="button" id="toggleHistoricalBtnInventory" class="btn-secondary-inventory">Show</button>
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
    style: `
  /* ----- Core Layout & Components ----- */
  .inventory-page {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  }

  .hidden-section {
    display: none !important;
  }

  .visible-section-inventory {
    display: block !important;     
  }

  .forecast-section-inventory.visible-section-inventory {
    display: block;
  }

  .obl-section-container-inventory.visible-section-inventory {
    display: block;
  }

  .card { 
    background:#fff; 
    border-radius:10px; 
    padding:20px; 
    margin-bottom:20px; 
    box-shadow:0 2px 8px rgba(0,0,0,0.08);
  }

  .detailed-month-inventory {
    position: relative;
  }

  .current-month-indicator-inventory {
    position: absolute;
    top: 8px;
    right: 8px;
    background: #0078d4;
    color: white;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 10px;
    font-weight: bold;
  }

  .enhanced-inputs-container-inventory {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
    margin-bottom: 15px;
  }

  .enhanced-inputs-container-inventory.four-col-inventory {
    grid-template-columns: repeat(4, 1fr);
  }

  .enhanced-input-group-inventory {
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .enhanced-input-group-inventory.full-width-inventory {
    grid-column: 1 / -1;
  }

  .enhanced-input-group-inventory label {
    font-size: 11px;
    color: #666;
    margin-bottom: 4px;
    font-weight: 600;
    text-align: center;
  }

  .enhanced-input-group-inventory input,
  .enhanced-input-group-inventory select {
    width: 100%;
    padding: 6px 8px;
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    font-size: 12px;
    background: #fff;
    transition: all 0.2s ease;
    color: #000 !important;
  }

  .enhanced-input-group-inventory input:focus,
  .enhanced-input-group-inventory select:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
    outline: none;
  }

  .forecast-date-input-inventory {
    background: #fff !important;
    color: #333 !important;
    border: 1px solid #ddd !important;
  }

  /* ----- Table Actions & Filters ----- */
  .table-actions-inventory {
    display: flex;
    flex-wrap: wrap;
    gap: 15px;
    margin-bottom: 20px;
    padding: 15px;
    background: #f8f9fa;
    border-radius: 8px;
    border: 1px solid #e9ecef;
  }

  .table-actions-row-inventory {
    display: grid;
    grid-template-columns: repeat(6, 1fr);
    gap: 15px;
    align-items: end;
  }

  .tableColp4-inventory {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .tableColp4-inventory label {
    font-size: 12px;
    font-weight: 600;
    color: #495057;
    margin-bottom: 4px;
  }

  .tableColp4-inventory select {
    padding: 8px 12px;
    border: 1px solid #ddd;
    border-radius: 5px;
    font-size: 14px;
    background: #fff;
    width: 100%;
  }

  /* ----- Buttons & Actions ----- */
  .btn-primary-inventory {
    background: #0078d4;
    color: white;
    padding: 8px 12px;
    border: none;
    border-radius: 5px;
    font-size: 14px;
    cursor: pointer;
    transition: all 0.2s ease;
  }

  .btn-primary-inventory:hover:not(:disabled) {
    background: #106ebe;
    transform: translateY(-1px);
  }

  .btn-primary-inventory:disabled {
    background: #a19f9d;
    cursor: not-allowed;
  }

  .btn-secondary-inventory {
    background: #6c757d;
    color: white;
    padding: 8px 12px;
    border: none;
    border-radius: 5px;
    font-size: 14px;
    cursor: pointer;
    transition: all 0.2s ease;
  }

  .btn-secondary-inventory:hover:not(:disabled) {
    background: #5a6268;
    transform: translateY(-1px);
  }

  .clear-filters-btn-inventory {
    background: #dc3545 !important;
  }

  .clear-filters-btn-inventory:hover {
    background: #c82333 !important;
  }

  .export-btn-inventory {
    background: #17a2b8 !important;
  }

  .export-btn-inventory:hover {
    background: #138496 !important;
  }

  /* ----- Forms & Inputs ----- */
  .form-row-inventory { 
    display:flex; 
    flex-wrap:wrap; 
    gap:20px; 
    margin-bottom:20px;
  }   

  .form-group-inventory { 
    flex:1 1 250px; 
    display:flex; 
    flex-direction:column;     
  }

  .form-group-inventory label {
    font-weight:600;
    margin-bottom:5px;
    color:#333;
  }

  .form-group-inventory input,
  .form-group-inventory select {
    padding:8px 12px;
    border:1px solid #ddd;
    border-radius:5px;
    font-size:14px;
  }

  .edit-mode-inventory .project-select-container-inventory select {
    background-color: #f8f9fa !important;
    color: #666 !important;
    border-color: #d0d0d0 !important;
    cursor: not-allowed !important;
  }

  /* ----- OBL Section ----- */
  .obl-section-container-inventory {
    background: #f8f9fa;
    padding: 20px;
    border-radius: 10px;
    border: 1px solid #e9ecef;
    margin-bottom: 25px;
  }

  .obl-main-section-inventory {
    display: grid;
    grid-template-columns: 1fr 2fr;
    gap: 20px;
    align-items: start;
  }

  .obl-input-inventory {
    font-size: 18px !important;
    font-weight: 700 !important;
    text-align: center !important;
    background: #fff !important;
    border: 1px solid #ddd !important;
    padding: 12px !important;
  }

  .validation-summary-inventory {
    background: #fff;
    padding: 15px;
    border-radius: 8px;
    border: 1px solid #e9ecef;
  }

  .validation-item-inventory {
    display: flex;
    justify-content: space-between;
    margin-bottom: 8px;
    font-size: 14px;
  }

  .validation-label-inventory {
    font-weight: 600;
    color: #666;
  }

  .validation-value-inventory {
    font-weight: 700;
  }

  .validation-line-inventory {
    height: 8px;
    background: #e9ecef;
    border-radius: 4px;
    margin-top: 10px;
    overflow: hidden;
  }

  .line-fill-inventory {
    height: 100%;
    background: #28a745;
    border-radius: 4px;
    transition: all 0.3s ease;
    width: 0%;
  }

  .line-fill-inventory.warning-inventory { background: #ffc107; }
  .line-fill-inventory.danger-inventory { background: #dc3545; }

  .validation-message-inventory {
    margin-top: 10px;
    padding: 10px;
    border-radius: 5px;
    font-size: 14px;
    font-weight: 600;
    text-align: center;
  }

  .validation-message-inventory.success-inventory {
    background: #d4edda;
    color: #155724;
    border: 1px solid #c3e6cb;
  }

  .validation-message-inventory.warning-inventory {
    background: #fff3cd;
    color: #856404;
    border: 1px solid #ffeaa7;
  }

  .validation-message-inventory.danger-inventory {
    background: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
  }

  /* ----- Forecast Sections ----- */
  .forecast-section-inventory {
    margin: 25px 0;
    padding: 20px;
    background: #f8f9fa;
    border-radius: 10px;
    border: 1px solid #e9ecef;
  }

  .forecast-section-inventory h3 {
    margin: 0 0 15px 0;
    color: #0078d4;
    font-size: 16px;
    font-weight: 600;
  }

  .first-row-container-inventory {
    display: grid;
    grid-template-columns: repeat(1, 1fr);
    gap: 15px;
  }

  .historical-container-inventory  {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 15px;
  }

  .remaining-container-inventory {
    display: flex;
    flex-direction: column;
    gap: 15px;
  }

  .remaining-row-inventory {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 15px;
  }

  /* ----- Detailed Forecast Row Layout (Current + Next 2 Months) ----- */
  .detailed-months-row-inventory {
    display: flex;
    gap: 15px;
    align-items: stretch;
  }

  .current-month-container-inventory {
    flex: 1 1 0;
    display: flex;
  }

  .next-months-container-inventory {
    flex: 2 1 0;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 15px;
  }

  .detailed-month-box-inventory {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 15px;
    background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    position: relative;
    width: 100%;
  }

  .current-month-box-inventory {
    border: 2px solid #0078d4;
    box-shadow: 0 2px 10px rgba(0,120,212,0.2);
  }

  .next-month-box-inventory {
    border: 1px solid #e0e0e0;
  }

  /* ----- 5-Year Forecast Row Layout ----- */
  .yearly-container-inventory {
    display: flex;
    gap: 15px;
  }

  .yearly-box-inventory {
    flex: 1 1 0;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 15px;
    background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  }

  .yearly-header-inventory {
    text-align: center;
    font-weight: 600;
    color: #0078d4;
    margin-bottom: 15px;
    padding-bottom: 8px;
    border-bottom: 1px solid #eee;
    font-size: 15px;
  }

  .yearly-inputs-container-inventory {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 10px;
  }

  .yearly-input-group-inventory {
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .yearly-input-group-inventory label {
    font-size: 11px;
    color: #666;
    margin-bottom: 4px;
    font-weight: 600;
  }

  .yearly-input-group-inventory input {
    width: 100%;
    padding: 6px 8px;
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    font-size: 12px;
    background: #f8f9fa;
    color: #000 !important;
  }

  .yearly-input-group-inventory input:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
    outline: none;
  }

  /* ----- Month Box Styles ----- */
  .month-box-inventory {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 15px;
    background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    position: relative;
  }

  .month-header-inventory {
    text-align: center;
    font-weight: 600;
    color: #0078d4;
    margin-bottom: 15px;
    padding-bottom: 8px;
    border-bottom: 1px solid #eee;
  }

  /* ----- Historical Month Specific Styles ----- */
  .historical-month-inventory {
    background: linear-gradient(135deg, #f5f5f5 0%, #e8e8e8 100%) !important;
    border: 1px solid #d0d0d0 !important;
    pointer-events: none !important;
    user-select: none !important;
    opacity: 0.8;
  }

  .historical-month-inventory * {
    pointer-events: none !important;
    user-select: none !important;
  }

  .historical-month-inventory::before {
    content: "READ ONLY";
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%) rotate(-45deg);
    font-size: 24px;
    font-weight: bold;
    color: rgba(108, 117, 125, 0.1);
    z-index: 1;
    pointer-events: none;
  }

  .historical-month-inventory .month-header-inventory {
    background: #6c757d !important;
    color: white !important;
    padding: 8px;
    margin: -15px -15px 15px -15px;
    border-radius: 8px 8px 0 0;
    border-bottom: 1px solid #ccc;
  }

  .historical-month-inventory input {
    background-color: #f8f9fa !important;
    color: #6c757d !important;
    border-color: #ced4da !important;
    cursor: not-allowed !important;
  }

  /* ----- Input Groups ----- */
  .main-inputs-container-inventory {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 15px;
    margin-bottom: 15px;
  }

  .main-input-group-inventory {
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .main-input-group-inventory label {
    font-size: 12px;
    color: #666;
    margin-bottom: 6px;
    font-weight: 600;
  }

  .main-input-group-inventory input {
    width: 100%;
    padding: 8px 10px;
    color:#000 !important;
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    font-size: 14px;
    background: #fff;
    transition: all 0.2s ease;
  }

  .main-input-group-inventory input:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
    outline: none;
  }

  /* ----- Items Section ----- */
  .items-section-inventory {
    margin-top: 15px;
    padding-top: 15px;
    border-top: 2px solid #e0e0e0;
  }

  .items-header-inventory {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 8px;
    font-size: 11px;
    color: #666;
    text-align: center;
    font-weight: 600;
    padding-bottom: 4px;
    border-bottom: 1px solid #eee;
  }

  .item-row-inventory {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 6px;
    align-items: center;
  }

  .item-row-inventory input {
    padding: 5px 8px;
    border: 1px solid #ddd;
    border-radius: 3px;
    font-size: 12px;
    background: #fff;
    transition: all 0.2s ease;
  }

  .item-row-inventory input[type="text"] { text-align: left; }
  .item-row-inventory input[type="number"] { text-align: right; }

  .item-row-inventory input:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 1px rgba(0,120,212,0.2);
    outline: none;
  }

  /* ----- Pagination ----- */
  .pagination-container-inventory {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 20px;
    padding: 15px;
    background-color: #f8f9fa;
    border-radius: 8px;
    border: 1px solid #dee2e6;
    flex-wrap: wrap;
    gap: 15px;
  }

  .pagination-left-inventory {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
  }

  .pagination-right-inventory {
    display: flex;
    align-items: center;
    gap: 20px;
    flex-wrap: wrap;
  }

  .page-size-selector-inventory {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .page-size-select-inventory {
    padding: 5px 10px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    background-color: white;
    font-size: 14px;
    cursor: pointer;
  }

  .pagination-info-inventory {
    font-size: 14px;
    color: #495057;
    white-space: nowrap;
  }

  .pagination-btn-inventory {
    padding: 6px 12px;
    background-color: #007bff;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
    transition: background-color 0.2s;
  }

  .pagination-btn-inventory:hover:not(:disabled) {
    background-color: #0056b3;
  }

  .pagination-btn-inventory:disabled {
    background-color: #6c757d;
    cursor: not-allowed;
    opacity: 0.6;
  }

  .pagination-pages-inventory {
    display: flex;
    gap: 5px;
  }

  .page-number-inventory {
    padding: 6px 12px;
    background-color: white;
    border: 1px solid #dee2e6;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
    transition: all 0.2s;
  }

  .page-number-inventory:hover:not(.active-inventory) {
    background-color: #f8f9fa;
    border-color: #007bff;
  }

  .page-number-inventory.active-inventory {
    background-color: #007bff;
    color: white;
    border-color: #007bff;
  }

  .page-number-inventory.dots-inventory {
    background-color: transparent;
    border: none;
    cursor: default;
    color: #6c757d;
  }

  /* ----- Utility Classes ----- */
  .input-with-value-inventory { background: #fff !important; color: #333 !important; }
  .input-empty-inventory { background: #f8f9fa !important; color: transparent !important; }

  /* ----- OBL Error Toast ----- */
  .obl-error-toast-inventory {
    position: fixed;
    bottom: 20px;
    right: 20px;
    background: #dc3545;
    color: white;
    padding: 15px 20px;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    z-index: 10000;
    max-width: 300px;
    animation: slideInRight 0.3s ease;
  }

  .obl-error-toast-inventory.warning-inventory {
    background: #ffc107;
    color: #856404;
  }

  @keyframes slideInRight {
    from {
      transform: translateX(100%);
      opacity: 0;
    }
    to {
      transform: translateX(0);
      opacity: 1;
    }
  }

  .toast-close-inventory {
    background: none;
    border: none;
    color: inherit;
    font-size: 18px;
    cursor: pointer;
    margin-left: 10px;
    float: right;
  }

  /* ============================================ */
  /* MAIN TABLE - MONTH COLUMNS                   */
  /* ============================================ */

  .month-columns-table-inventory {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 12px;
    min-width: 1200px;
  }

  .month-columns-table-inventory th,
  .month-columns-table-inventory td {
    padding: 6px 4px;
    border: 1px solid #dee2e6;
    text-align: center;
    white-space: normal !important;
    word-wrap: break-word !important;
    word-break: break-word !important;
    overflow-wrap: break-word !important;
  }

  /* ----- STICKY COLUMNS ----- */

  .project-ref-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 0 !important;
    min-width: 120px !important;
    max-width: 120px !important;
    width: 120px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-inventory thead .project-ref-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .project-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 120px !important;
    min-width: 180px !important;
    max-width: 180px !important;
    width: 180px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-inventory thead .project-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .pm-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 300px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-inventory thead .pm-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .gh-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 400px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-inventory thead .gh-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .customer-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 500px !important;
    min-width: 130px !important;
    max-width: 130px !important;
    width: 130px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-inventory thead .customer-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .inv-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 630px !important;
    min-width: 90px !important;
    max-width: 90px !important;
    width: 90px !important;
    text-align: right !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-inventory thead .inv-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .month-scroll-col {
    min-width: 60px !important;
    width: 60px !important;
  }

  /* ----- TOTALS HEADER TABLE ----- */
  .totals-header-table-inventory {
    margin-bottom: 10px;
    background-color: #f8f9fa;
    border-collapse: separate;
    border-spacing: 0;
    table-layout: fixed !important;
    width: 100% !important;
  }

  .totals-header-table-inventory th,
  .totals-header-table-inventory td {
    width: auto !important;
    min-width: 0 !important;
    padding: 6px 4px !important;
    font-size: 12px;
    font-weight: bold;
    text-align: center;
    vertical-align: middle;
    border: 1px solid #dee2e6;
    white-space: normal !important;
    word-wrap: break-word !important;
    word-break: break-word !important;
    overflow-wrap: break-word !important;
  }

  .totals-header-table-inventory .totals-label-inventory {
    font-weight: 600;
    color: #495057;
    background-color: #f1f3f4 !important;
  }

  .totals-header-table-inventory .all-projects-total-inventory td {
    background-color: #d1ecf1 !important;
    color: #0c5460;
  }

  .totals-header-table-inventory .filtered-total-inventory td {
    background-color: #d4edda !important;
    color: #155724;
  }

  .totals-header-table-inventory .total-value-inventory {
    font-weight: bold;
    font-size: 13px;
  }

  /* Totals header sticky columns */
  .totals-header-table-inventory .project-ref-col {
    left: 0 !important;
    min-width: 120px !important;
    max-width: 120px !important;
    width: 120px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-inventory .project-col {
    left: 120px !important;
    min-width: 180px !important;
    max-width: 180px !important;
    width: 180px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-inventory .pm-col {
    left: 300px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-inventory .gh-col {
    left: 400px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-inventory .customer-col {
    left: 500px !important;
    min-width: 130px !important;
    max-width: 130px !important;
    width: 130px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-inventory .inv-col {
    left: 630px !important;
    min-width: 90px !important;
    max-width: 90px !important;
    width: 90px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-inventory .month-scroll-col {
    min-width: 60px !important;
    width: 60px !important;
  }

  /* ----- MONTH COLUMN HEADERS ----- */
  .month-column-header-inventory {
    background: #0078d4;
    color: white;
    font-weight: 600;
    font-size: 11px;
    white-space: nowrap !important;
  }

  .month-subheader-inventory {
    background: #e3f2fd;
    color: #0078d4;
    font-weight: 600;
    font-size: 10px;
    white-space: nowrap !important;
  }

  .current-month-header-inventory {
    background: #2196f3 !important;
    color: white !important;
  }

  .current-month-header-table-inventory {
    background: #ff9800 !important;
    color: white !important;
  }

  /* ----- SORT HEADERS ----- */
  .sort-header-container-inventory {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
  }

  .sort-buttons-container-inventory {
    display: flex;
    flex-direction: column;
    gap: 1px;
  }

  .sort-btn-inventory {
    background: none;
    border: none;
    cursor: pointer;
    font-size: 8px;
    padding: 1px 2px;
    border-radius: 2px;
    color: #ccc;
    transition: all 0.2s ease;
  }

  .sort-btn-inventory:hover {
    background-color: rgba(255, 255, 255, 0.2);
    color: white;
  }

  .sort-btn-inventory.active-inventory {
    color: #ffeb3b;
  }

  .sort-btn-inventory.sort-asc-inventory.active-inventory {
    color: #4caf50;
  }

  .sort-btn-inventory.sort-desc-inventory.active-inventory {
    color: #f44336;
  }

  .header-label-inventory {
    margin-right: 4px;
  }

  /* ----- Clickable Edit Link ----- */
  .edit-trigger {
    color: #0078d4 !important;
    text-decoration: underline !important;
    cursor: pointer !important;
  }

  .edit-trigger:hover {
    color: #0056b3 !important;
  }

  /* ----- RESPONSIVE ----- */
  @media (max-width: 1400px) {
    .project-ref-col {
      left: 0 !important;
      min-width: 100px !important;
      max-width: 100px !important;
      width: 100px !important;
    }
    .project-col {
      left: 100px !important;
      min-width: 150px !important;
      max-width: 150px !important;
      width: 150px !important;
    }
    .pm-col {
      left: 250px !important;
      min-width: 80px !important;
      max-width: 80px !important;
      width: 80px !important;
    }
    .gh-col {
      left: 330px !important;
      min-width: 80px !important;
      max-width: 80px !important;
      width: 80px !important;
    }
    .customer-col {
      left: 410px !important;
      min-width: 110px !important;
      max-width: 110px !important;
      width: 110px !important;
    }
    .inv-col {
      left: 520px !important;
      min-width: 70px !important;
      max-width: 70px !important;
      width: 70px !important;
    }
    .table-actions-row-inventory {
      grid-template-columns: repeat(3, 1fr);
    }
    .historical-container-inventory,
    .first-row-container-inventory { grid-template-columns: repeat(2, 1fr); }
    .remaining-row-inventory { grid-template-columns: repeat(3, 1fr); }
    .obl-main-section-inventory { grid-template-columns: 1fr; }
  }

  @media (max-width: 1024px) {
    .table-actions-row-inventory {
      grid-template-columns: repeat(2, 1fr);
    }
    .historical-container-inventory,
    .first-row-container-inventory { grid-template-columns: 1fr; }
    .remaining-row-inventory { grid-template-columns: repeat(2, 1fr); }
    .detailed-months-row-inventory { flex-direction: column; }
    .next-months-container-inventory { grid-template-columns: 1fr; }
    .yearly-container-inventory { flex-wrap: wrap; }
    .yearly-box-inventory { flex: 1 1 45%; }
  }

  @media (max-width: 768px) {
    .project-ref-col {
      left: 0 !important;
      min-width: 70px !important;
      max-width: 70px !important;
      width: 70px !important;
    }
    .project-col {
      left: 70px !important;
      min-width: 100px !important;
      max-width: 100px !important;
      width: 100px !important;
    }
    .pm-col {
      left: 170px !important;
      min-width: 60px !important;
      max-width: 60px !important;
      width: 60px !important;
    }
    .gh-col {
      left: 230px !important;
      min-width: 60px !important;
      max-width: 60px !important;
      width: 60px !important;
    }
    .customer-col {
      left: 290px !important;
      min-width: 80px !important;
      max-width: 80px !important;
      width: 80px !important;
    }
    .inv-col {
      left: 370px !important;
      min-width: 55px !important;
      max-width: 55px !important;
      width: 55px !important;
    }
    .table-actions-row-inventory {
      grid-template-columns: 1fr;
    }
    .table-actions-inventory {
      padding: 10px;
    }
    .tableColp4-inventory {
      min-width: 100%;
    }
    .historical-container-inventory,
    .first-row-container-inventory,
    .remaining-row-inventory { grid-template-columns: 1fr; }
    .forecast-section-inventory { padding: 15px; }
    .pagination-container-inventory { flex-direction: column; gap: 10px; }
    .pagination-left-inventory,
    .pagination-right-inventory { justify-content: center; width: 100%; }
    .next-months-container-inventory {
      grid-template-columns: 1fr;
    }
    .enhanced-inputs-container-inventory {
      grid-template-columns: 1fr 1fr;
    }
    .enhanced-inputs-container-inventory.four-col-inventory {
      grid-template-columns: 1fr 1fr;
    }
    .items-header-inventory,
    .item-row-inventory {
      grid-template-columns: 1fr;
      gap: 4px;
    }
    .items-header-inventory div:nth-child(n+4),
    .item-row-inventory input:nth-child(n+4) {
      display: none;
    }
    .yearly-box-inventory { flex: 1 1 100%; }
  }

  @media (max-width: 1200px) {
    .detailed-months-row-inventory {
      flex-direction: column;
    }
    .current-month-container-inventory,
    .next-months-container-inventory {
      flex: 1;
      max-width: 100%;
    }
    .next-months-container-inventory {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 15px;
    }
  }

  /* ----- ENHANCED VALIDATION STYLES ----- */
  .validation-message-inventory.success-inventory {
    background: #d4edda;
    color: #155724;
    border: 1px solid #c3e6cb;
  }

  .validation-message-inventory.warning-inventory {
    background: #fff3cd;
    color: #856404;
    border: 1px solid #ffeaa7;
  }

  .validation-message-inventory.error-inventory {
    background: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
  }

  .input-invalid-inventory {
    border-color: #dc3545 !important;
    background-color: #f8f9fa !important;
  }

  .input-warning-inventory {
    border-color: #ffc107 !important;
    background-color: #fffbf0 !important;
  }

  .btn-loading-inventory {
    opacity: 0.7;
    cursor: not-allowed !important;
  }

  .btn-loading-inventory:hover {
    transform: none !important;
  }

  .obl-section-container-inventory.critical-inventory {
    border-left: 4px solid #dc3545;
  }

  .obl-section-container-inventory.warning-inventory {
    border-left: 4px solid #ffc107;
  }

  .obl-section-container-inventory.normal-inventory {
    border-left: 4px solid #28a745;
  }

  .role-indicator-inventory {
    display: none;
  }

  .table-container-inventory > div > table:last-child {
    margin-top: 0;
  }

  tr.filtered-total-inventory {
    background-color: #00ffff;
  }
  tr.all-projects-total-inventory {
    background-color: #fff8dc;
  }

  /* ----- HISTORICAL TOGGLE HEADER ----- */
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
    transition: background 0.2s !important;
    user-select: none !important;
  }

  .historical-toggle-header-inventory:hover {
    background: #e9ecef !important;
  }

  .historical-toggle-header-inventory h3 {
    margin: 0 !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    color: #495057 !important;
  }

  #toggleHistoricalBtnInventory {
    padding: 4px 16px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    background: #28a745 !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    cursor: pointer !important;
    transition: background 0.2s !important;
    flex-shrink: 0 !important;
  }

  #toggleHistoricalBtnInventory:hover {
    opacity: 0.85 !important;
  }

  .historical-container-inventory {
    transition: all 0.3s ease !important;
  }

  /* ----- PROJECT MAIN WRAPPER - 2 COLUMN LAYOUT ----- */
  .project-main-wrapper-inventory {
    display: grid !important;
    grid-template-columns: 1fr 1fr !important;
    gap: 0 !important;
    background: #f8f9fa !important;
    border-radius: 6px !important;
    border: 1px solid #e9ecef !important;
    overflow: hidden !important;
    margin: 4px 0 8px 0 !important;
    min-height: 234px !important;
  }

  .project-left-side-inventory {
    padding: 8px 16px !important;
    background: #ffffff !important;
    border-right: 1px solid #e9ecef !important;
    display: flex !important;
    flex-direction: column !important;
    gap: 0 !important;
    justify-content: flex-start !important;
    align-items: flex-start !important;
  }

  .project-left-side-inventory .detail-item-inventory {
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    gap: 10px !important;
    padding: 8px 0 !important;
    font-size: 12px !important;
    border-bottom: 1px solid #f0f0f0 !important;
    text-align: left !important;
    justify-content: flex-start !important;
    width: 100% !important;
  }

  .project-left-side-inventory .detail-label-inventory {
    font-weight: 700 !important;
    white-space: nowrap !important;
    min-width: 100px !important;
  }

  .project-left-side-inventory .detail-value-inventory {
    white-space: nowrap !important;
  }

  .project-right-side-inventory {
    padding: 8px 16px !important;
    background: #ffffff !important;
    display: flex !important;
    flex-direction: column !important;
    gap: 6px !important;
    justify-content: flex-start !important;
  }

  /* ----- OBL SECTION (matching SEIBIES) ----- */
  .obl-section-container-inventory {
    margin: 2px 0 0 0 !important;
    padding: 6px 12px !important;
    background: #f8f9fa !important;
    border-radius: 4px !important;
    border: 1px solid #e9ecef !important;
  }

  .obl-content-inventory {
    display: grid !important;
    grid-template-columns: 1fr 1.5fr !important;
    gap: 12px !important;
    align-items: start !important;
  }

  .obl-left-inventory {
    display: flex !important;
    flex-direction: row !important;
  }

  .obl-label-group-inventory {
    display: flex !important;
    align-items: center !important;
    gap: 4px !important;
    margin-bottom: 2px !important;
    text-align: left !important;
  }

  .obl-label-group-inventory label {
    font-size: 12px !important;
    font-weight: 700 !important;
    color: #495057 !important;
    margin: 0 !important;
  }

  .info-tooltip-inventory {
    font-size: 12px !important;
    color: #007bff !important;
    cursor: help !important;
  }

  .obl-value-display-inventory {
    display: flex !important;
    flex-direction: column !important;
  }

  .obl-input-inventory {
    font-size: 18px !important;
    font-weight: 700 !important;
    text-align: center !important;
    background: #fff !important;
    border: 1px solid #ddd !important;
    padding: 2px 8px !important;
    height: 36px !important;
    width: 100% !important;
    max-width: 140px !important;
    border-radius: 4px !important;
    color: #0078d4 !important;
  }

  .field-note-inventory {
    font-size: 9px !important;
    color: #6c757d !important;
    margin-top: 2px !important;
    font-style: italic !important;
    text-align: left !important;
  }

  .obl-right-inventory {
    display: flex !important;
    align-items: center !important;
  }

  .validation-summary-inventory {
    background: #fff !important;
    padding: 4px 12px !important;
    border-radius: 4px !important;
    border: 1px solid #e9ecef !important;
    width: 100% !important;
  }

  .validation-item-inventory {
    display: flex !important;
    justify-content: space-between !important;
    margin-bottom: 1px !important;
    font-size: 11px !important;
    text-align: left !important;
  }

  .validation-label-inventory {
    font-weight: 600 !important;
    color: #666 !important;
    font-size: 10px !important;
    text-align: left !important;
  }

  .validation-value-inventory {
    font-weight: 700 !important;
    font-size: 11px !important;
    color: #333 !important;
    text-align: right !important;
  }

  .validation-line-inventory {
    height: 4px !important;
    background: #e9ecef !important;
    border-radius: 2px !important;
    margin-top: 3px !important;
    overflow: hidden !important;
  }

  .line-fill-inventory {
    height: 100% !important;
    background: #28a745 !important;
    border-radius: 2px !important;
    transition: all 0.3s ease !important;
    width: 0% !important;
  }

  .line-fill-inventory.warning-inventory {
    background: #ffc107 !important;
  }

  .line-fill-inventory.danger-inventory {
    background: #dc3545 !important;
  }

  .validation-message-inventory {
    display: block !important;
    padding: 8px 12px;
    border-radius: 4px;
    margin-top: 5px;
    font-size: 12px;
    font-weight: 500;
  }

  .validation-message-inventory.danger-inventory {
    background-color: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
  }

  .validation-message-inventory.warning-inventory {
    background-color: #fff3cd;
    color: #856404;
    border: 1px solid #ffc107;
  }

  /* ----- RESPONSIVE for project wrapper ----- */
  @media (max-width: 992px) {
    .project-main-wrapper-inventory {
      grid-template-columns: 1fr !important;
    }
    
    .project-left-side-inventory {
      border-right: none !important;
      border-bottom: 1px solid #e9ecef !important;
    }
    
    .project-right-side-inventory {
      padding: 8px 14px !important;
    }
  }

  @media (max-width: 768px) {
    .obl-content-inventory {
      grid-template-columns: 1fr !important;
      gap: 8px !important;
    }
    
    .obl-input-inventory {
      max-width: 100% !important;
    }
    
    .project-left-side-inventory .detail-label-inventory {
      min-width: 80px !important;
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
                let allMonths = [];
                // Pagination variables
                let currentPage = 1;
                let pageSize = 10;
                let totalPages = 1;
                let filteredData = [];
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
                            //const inventoryActualData = safeJsonParse(forecast.InventoryActual, {});
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
                                "Group Head": (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
                                "Project Manager": typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                                "Inventory": oblValue,
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
                const renderTotalsHeader = (displayMonths, currentMonthKey) => {
                    if (!totalsHeader)
                        return;
                    totalsHeader.innerHTML = '';
                    const allData = forecasts;
                    const filtered = filteredData.length ? filteredData : allData;
                    const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
                    const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);
                    const buildTotalsRow = (label, totals, css) => {
                        return `
              <tr class="${css}">
                <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;"></td>
                <td class="project-col" style="min-width:180px; max-width:180px; width:180px; text-align:left; font-weight:bold; background-color:${css === 'all-projects-total-inventory' ? '#d1ecf1' : '#d4edda'};">
                  ${label}
                </td>
                <td class="pm-col" style="min-width:100px; max-width:100px; width:100px;"></td>
                <td class="gh-col" style="min-width:100px; max-width:100px; width:100px;"></td>
                <td class="customer-col" style="min-width:130px; max-width:130px; width:130px;"></td>
                <td class="inv-col" style="min-width:90px; max-width:90px; width:90px; text-align:right; font-weight:bold;">${totals.obl.toLocaleString('en-IN')}</td>
                ${displayMonths.map(m => {
                            const t = totals.monthly[m] || { plan: 0, current: 0 };
                            return `
                    <td class="month-scroll-col" style="min-width:60px;">${t.plan.toLocaleString('en-IN')}</td>
                    <td class="month-scroll-col" style="min-width:60px;">${t.current.toLocaleString('en-IN')}</td>
                  `;
                        }).join('')}
              </tr>
            `;
                    };
                    totalsHeader.innerHTML = `
            <table class="totals-header-table-inventory" style="width: 100%;">
              ${buildTotalsRow(`All (${allData.length})`, allTotals, 'all-projects-total-inventory')}
              ${buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total-inventory')}
            </table>
          `;
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
                    // Build main table header - ONLY Plan and Current columns
                    const mainHeaderHTML = `
            <tr>
              <!-- Project Reference - NOW CLICKABLE LIKE PROJECT -->
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
              
              <!-- Project Column with Sort - clickable to edit -->
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
              
              <!-- Project Manager -->
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
              
              <!-- Group Head -->
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
              
              <!-- Customer -->
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
              
              <!-- Inventory Column -->
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
                        // All months have 2 columns: Plan and Current
                        return `
                  <th colspan="2" class="month-scroll-col month-column-header-inventory ${month === currentMonthKey ? 'current-month-header-inventory' : ''}" style="min-width:60px;">
                    <div class="sort-header-container-inventory">
                      <span class="header-label-inventory">${formatMonthLabel(new Date(month + "-01"))}</span>
                      <div class="sort-buttons-container-inventory">
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort by Month Total Ascending">▲</button>
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="desc" data-month-index="${monthIndex}" title="Sort by Month Total Descending">▼</button>
                      </div>
                    </div>
                    ${month === currentMonthKey ? ' <span style="font-size:10px;color:#ffeb3b;">(Current)</span>' : ''}
                  </th>
                `;
                    }).join('')}
            </tr>
            
            <tr>
              ${displayMonths.map((month, monthIndex) => {
                        // All months have Plan and Current sub-headers only
                        return `
                  <th class="month-scroll-col month-subheader-inventory ${month === currentMonthKey ? 'current-month-header-table-inventory' : ''}" style="min-width:60px;">
                    <div class="sort-header-container-inventory">
                      <span>Plan</span>
                      <div class="sort-buttons-container-inventory">
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Ascending">▲</button>
                        <button class="sort-btn-inventory ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc-inventory active-inventory' : ''}" 
                                data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Descending">▼</button>
                      </div>
                    </div>
                  </th>
                  <th class="month-scroll-col month-subheader-inventory ${month === currentMonthKey ? 'current-month-header-table-inventory' : ''}" style="min-width:60px;">
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
                        var _a, _b;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-IN');
                        // Build project name link - click to edit
                        const projectNameLink = `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectTitle}</a>`;
                        // Build project reference link - NOW CLICKABLE LIKE PROJECT
                        const projectRefLink = `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${(proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-"}</a>`;
                        let rowHTML = `
              <tr data-forecast-id="${forecast.Id}">
                <!-- Project Reference - NOW CLICKABLE -->
                <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
                
                <!-- Project - clickable link -->
                <td class="project-col" style="min-width:180px; max-width:180px; width:180px;">${projectNameLink}</td>
                
                <!-- Project Manager -->
                <td class="pm-col" style="min-width:100px; max-width:100px; width:100px;">${typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-"}</td>
                
                <!-- Group Head -->
                <td class="gh-col" style="min-width:100px; max-width:100px; width:100px;">${(proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-"}</td>
                
                <!-- Customer -->
                <td class="customer-col" style="min-width:130px; max-width:130px; width:130px;">${(proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-"}</td>
                
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
                        var _a, _b, _c, _d;
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
                                valueA = typeof (projMA === null || projMA === void 0 ? void 0 : projMA.ProjectManager) === "object" ? (_c = projMA.ProjectManager) === null || _c === void 0 ? void 0 : _c.Title : (projMA === null || projMA === void 0 ? void 0 : projMA.ProjectManager) || "";
                                valueB = typeof (projMB === null || projMB === void 0 ? void 0 : projMB.ProjectManager) === "object" ? (_d = projMB.ProjectManager) === null || _d === void 0 ? void 0 : _d.Title : (projMB === null || projMB === void 0 ? void 0 : projMB.ProjectManager) || "";
                                break;
                            case 'groupHead':
                                const projGA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                                const projGB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                                valueA = (projGA === null || projGA === void 0 ? void 0 : projGA.GroupHead) || "";
                                valueB = (projGB === null || projGB === void 0 ? void 0 : projGB.GroupHead) || "";
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
                // ============================================
                // UPDATED: FILTER FUNCTIONS
                // ============================================
                const initializeFilters = () => {
                    const currentYear = new Date().getFullYear();
                    const currentMonth = new Date().getMonth() + 1;
                    if (filterYear) {
                        filterYear.innerHTML = '<option value="">All Years</option>';
                        for (let year = currentYear - 1; year <= currentYear + 2; year++) {
                            const selected = year === currentYear ? 'selected' : '';
                            filterYear.innerHTML += `<option value="${year}" ${selected}>${year}</option>`;
                        }
                    }
                    if (filterMonth) {
                        const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                        filterMonth.innerHTML = '<option value="">All Months</option>';
                        for (let month = 1; month <= 12; month++) {
                            const selected = month === currentMonth ? 'selected' : '';
                            filterMonth.innerHTML += `<option value="${month}" ${selected}>${monthNames[month - 1]}</option>`;
                        }
                    }
                    // ✅ Use getAssignedProjectValues to get filtered values
                    try {
                        const assignedProjectTitles = getAssignedProjectValues(projects, 'Title');
                        const assignedCustomers = getAssignedProjectValues(projects, 'CustomerName');
                        const assignedManagers = getAssignedProjectValues(projects, 'ProjectManager');
                        const assignedGroupHeads = getAssignedProjectValues(projects, 'GroupHead');
                        const assignedReferences = getAssignedProjectValues(projects, 'ProjectReference');
                        if (filterProject) {
                            filterProject.innerHTML = '<option value="">All Projects</option>' +
                                assignedProjectTitles.map(project => `<option value="${project}">${project}</option>`).join("");
                        }
                        if (filterCustomer) {
                            filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                                assignedCustomers.map(customer => `<option value="${customer}">${customer}</option>`).join("");
                        }
                        if (filterProjectManager) {
                            filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                                assignedManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");
                        }
                        if (filterGroupHead) {
                            filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
                                assignedGroupHeads.map(head => `<option value="${head}">${head}</option>`).join("");
                        }
                        if (filterProjectReference) {
                            filterProjectReference.innerHTML = '<option value="">All References</option>' +
                                assignedReferences.map(ref => `<option value="${ref}">${ref}</option>`).join("");
                        }
                    }
                    catch (error) {
                        console.error("Error initializing filters:", error);
                        // Fallback to empty filters if something goes wrong
                        if (filterProject)
                            filterProject.innerHTML = '<option value="">All Projects</option>';
                        if (filterCustomer)
                            filterCustomer.innerHTML = '<option value="">All Customers</option>';
                        if (filterProjectManager)
                            filterProjectManager.innerHTML = '<option value="">All Managers</option>';
                        if (filterGroupHead)
                            filterGroupHead.innerHTML = '<option value="">All Group Heads</option>';
                        if (filterProjectReference)
                            filterProjectReference.innerHTML = '<option value="">All References</option>';
                    }
                };
                const filterTable = () => {
                    const year = filterYear ? filterYear.value : "";
                    const month = filterMonth ? filterMonth.value : "";
                    const project = filterProject ? filterProject.value : "";
                    const customer = filterCustomer ? filterCustomer.value : "";
                    const projectManager = filterProjectManager ? filterProjectManager.value : "";
                    const groupHead = filterGroupHead ? filterGroupHead.value : "";
                    const projectReference = filterProjectReference ? filterProjectReference.value : "";
                    // ✅ Get assigned projects
                    const assignedProjects = getAssignedProjects(projects);
                    const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                    let filtered = [...forecasts];
                    // ✅ Filter by assigned projects (unless Full Control or GH Head)
                    const isGHHeadUser = isGHHead();
                    const isFullControlUser = hasFullControl();
                    if (!isGHHeadUser && !isFullControlUser) {
                        filtered = filtered.filter(forecast => {
                            var _a;
                            const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                            return projectTitle && assignedProjectTitles.has(projectTitle);
                        });
                    }
                    // ✅ Apply other filters
                    filtered = filtered.filter((forecast) => {
                        var _a, _b;
                        // Year and Month filter
                        if (year && month) {
                            const monthKey = `${year}-${month.padStart(2, "0")}`;
                            const forecastData = safeJsonParse(forecast.ForecastData, {});
                            if (!(monthKey in forecastData))
                                return false;
                        }
                        // Project filter
                        if (project && ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) !== project)
                            return false;
                        // Customer filter
                        if (customer) {
                            const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            if (!proj || proj.CustomerName !== customer)
                                return false;
                        }
                        // Project Manager filter
                        if (projectManager) {
                            const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            if (!proj)
                                return false;
                            const managerName = typeof proj.ProjectManager === 'object'
                                ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                                : proj.ProjectManager;
                            if (managerName !== projectManager)
                                return false;
                        }
                        // Group Head filter
                        if (groupHead) {
                            const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            if (!proj || proj.GroupHead !== groupHead)
                                return false;
                        }
                        // Project Reference filter
                        if (projectReference) {
                            const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            if (!proj || proj.ProjectReference !== projectReference)
                                return false;
                        }
                        return true;
                    });
                    // Handle results
                    if (filtered.length === 0) {
                        if (tableBody) {
                            tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No Inventory records found for selected filters</td></tr>`;
                        }
                        if (totalsHeader) {
                            totalsHeader.innerHTML = '';
                        }
                        initializePagination([]);
                    }
                    else {
                        initializePagination(filtered);
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
                        '.main-current-input-inventory',
                        '.yearly-current-input-inventory',
                        '.main-forecast-input-inventory',
                        '.forecast-date-input-inventory'
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
                    // Handle OBL input
                    const oblInput = document.querySelector('#oblLastMonthInventory');
                    if (oblInput) {
                        if (isReadOnly) {
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
                        toast("error", "❌ You don't have permission to create Inventory forecasts");
                        return;
                    }
                    if (!formSection || !tableSection)
                        return;
                    formSection.style.display = "block";
                    tableSection.style.display = "none";
                    // ✅ Show/hide save button based on permissions
                    const saveBtn = document.querySelector('.btn-primary-inventory[type="submit"]');
                    if (saveBtn) {
                        const canSave = isEditMode ? canEditProject(project) : canCreateRecord();
                        saveBtn.style.display = canSave ? '' : 'none';
                    }
                    if (isEditMode) {
                        formSection.classList.add('edit-mode-inventory');
                        if (projectSelect)
                            projectSelect.disabled = true;
                        if (formTitle)
                            formTitle.textContent = "Edit Inventory Forecast";
                    }
                    else {
                        formSection.classList.remove('edit-mode-inventory');
                        if (projectSelect)
                            projectSelect.disabled = false;
                        if (formTitle)
                            formTitle.textContent = "Add Inventory Forecast";
                    }
                };
                const hideForm = () => {
                    if (!formSection || !tableSection)
                        return;
                    formSection.style.display = "none";
                    tableSection.style.display = "block";
                    formSection.classList.remove('edit-mode-inventory');
                    if (projectSelect)
                        projectSelect.disabled = false;
                };
                const resetForm = () => {
                    editingId = null;
                    if (form)
                        form.reset();
                    if (projectSelect)
                        projectSelect.value = "";
                    // Clear project details display
                    const projectNameDisplay = document.querySelector('#projectNameDisplayInventory');
                    if (projectNameDisplay)
                        projectNameDisplay.textContent = "-";
                    const projectRef = document.querySelector('#projectRefInventory');
                    if (projectRef)
                        projectRef.textContent = "-";
                    const profitCenter = document.querySelector('#profitCenterInventory');
                    if (profitCenter)
                        profitCenter.textContent = "-";
                    const projectManager = document.querySelector('#projectManagerInventory');
                    if (projectManager)
                        projectManager.textContent = "-";
                    const groupHead = document.querySelector('#groupHeadInventory');
                    if (groupHead)
                        groupHead.textContent = "-";
                    const customer = document.querySelector('#customerInventory');
                    if (customer)
                        customer.textContent = "-";
                    if (historicalContainer)
                        historicalContainer.innerHTML = "";
                    if (firstRowContainer)
                        firstRowContainer.innerHTML = "";
                    if (remainingContainer)
                        remainingContainer.innerHTML = "";
                    if (yearlyContainer)
                        yearlyContainer.innerHTML = "";
                    toggleForecastSections(false);
                    const oblContainer = document.querySelector('.obl-section-container-inventory');
                    if (oblContainer) {
                        oblContainer.className = 'obl-section-container-inventory hidden-section';
                    }
                    if (oblLastMonthInput) {
                        oblLastMonthInput.removeAttribute('readonly');
                        oblLastMonthInput.placeholder = "Enter OBL";
                        setNumericValue(oblLastMonthInput, 0);
                    }
                    if (oblValidationMessage) {
                        oblValidationMessage.textContent = "✅ All values are within OBL limit";
                        oblValidationMessage.className = 'validation-message-inventory';
                    }
                    updateOBLValidationDisplay();
                };
                // ========== VALIDATION FUNCTIONS ==========
                const calculateTotalCurrentValue = () => {
                    let total = 0;
                    const editableMonths = allMonths.slice(3);
                    editableMonths.forEach(monthKey => {
                        const currentInputs = document.querySelectorAll(`input.main-current-input-inventory[data-month="${monthKey}"]:not([readonly])`);
                        currentInputs.forEach(input => {
                            total += getNumericValue(input);
                        });
                    });
                    const yearlyCurrentInputs = document.querySelectorAll(`input.yearly-current-input-inventory:not([readonly])`);
                    yearlyCurrentInputs.forEach(input => {
                        total += getNumericValue(input);
                    });
                    return total;
                };
                const validateOBLLimit = (changedInput) => {
                    if (!oblLastMonthInput || !totalCurrentValueSpan || !remainingValueSpan || !oblLineFill || !oblValidationMessage)
                        return true;
                    const oblValue = getNumericValue(oblLastMonthInput);
                    const totalCurrent = calculateTotalCurrentValue();
                    totalCurrentValueSpan.textContent = totalCurrent.toLocaleString();
                    const remaining = oblValue - totalCurrent;
                    remainingValueSpan.textContent = remaining.toLocaleString();
                    if (oblValue === 0) {
                        oblLineFill.style.width = `0%`;
                        oblLineFill.className = 'line-fill-inventory';
                        if (totalCurrent > 0) {
                            showOBLErrorToast('⚠️ OBL is zero but current values exist. Please set OBL value.', true);
                            blockAllCurrentInputs();
                            return false;
                        }
                        else {
                            const existingToast = document.querySelector('.obl-error-toast-inventory');
                            if (existingToast)
                                existingToast.remove();
                            unblockAllCurrentInputs();
                            return true;
                        }
                    }
                    const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
                    oblLineFill.style.width = `${percentage}%`;
                    oblLineFill.className = 'line-fill-inventory';
                    if (percentage > 90) {
                        oblLineFill.classList.add('danger-inventory');
                    }
                    else if (percentage > 75) {
                        oblLineFill.classList.add('warning-inventory');
                    }
                    oblValidationMessage.className = 'validation-message-inventory';
                    if (totalCurrent > oblValue) {
                        const exceedAmount = totalCurrent - oblValue;
                        showOBLErrorToast(`❌ Total current value exceeds OBL by ${exceedAmount.toLocaleString()}! Reduce values to continue.`);
                        oblValidationMessage.classList.add('danger-inventory');
                        blockIncreasingCurrentInputs(changedInput);
                        return false;
                    }
                    else if (percentage > 90) {
                        showOBLErrorToast('⚠️ OBL limit nearly reached! Consider reducing values.', true);
                        oblValidationMessage.classList.add('warning-inventory');
                        unblockAllCurrentInputs();
                        return true;
                    }
                    else {
                        const existingToast = document.querySelector('.obl-error-toast-inventory');
                        if (existingToast)
                            existingToast.remove();
                        oblValidationMessage.textContent = '✅ All values are within OBL limit';
                        unblockAllCurrentInputs();
                        return true;
                    }
                };
                const updateOBLValidationDisplay = (changedInput) => {
                    validateOBLLimit(changedInput);
                };
                const blockAllCurrentInputs = () => {
                    const currentInputs = document.querySelectorAll('input[data-type="actual"]:not([readonly])');
                    const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-inventory:not([readonly])');
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
                    const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-inventory:not([readonly])');
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
                    const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-inventory');
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
                    if (!input)
                        return true;
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
                const showOBLErrorToast = (message, isWarning = false) => {
                    var _a;
                    const existingToast = document.querySelector('.obl-error-toast-inventory');
                    if (existingToast)
                        existingToast.remove();
                    const toast = document.createElement('div');
                    toast.className = `obl-error-toast-inventory ${isWarning ? 'warning-inventory' : ''}`;
                    toast.innerHTML = `
            ${message}
            <button class="toast-close-inventory">&times;</button>
          `;
                    document.body.appendChild(toast);
                    setTimeout(() => {
                        if (toast.parentNode) {
                            toast.remove();
                        }
                    }, 5000);
                    (_a = toast.querySelector('.toast-close-inventory')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
                        toast.remove();
                    });
                };
                // ========== FORECAST FORM FUNCTIONS ==========
                const toggleForecastSections = (show) => {
                    const sections = [
                        '.obl-section-container-inventory',
                        '#historicalForecastSectionInventory',
                        '#detailedForecastSectionInventory',
                        '#longtermForecastSectionInventory',
                        '#yearlyForecastSectionInventory'
                    ];
                    sections.forEach(selector => {
                        const element = document.querySelector(selector);
                        if (element) {
                            if (show) {
                                element.classList.remove('hidden-section');
                                element.classList.add('visible-section-inventory');
                            }
                            else {
                                element.classList.remove('visible-section-inventory');
                                element.classList.add('hidden-section');
                            }
                        }
                    });
                };
                // ========== RENDER PROJECT DETAILS (2-COLUMN LAYOUT) ==========
                const renderProjectDetails = (project) => {
                    var _a;
                    const projectNameDisplay = document.querySelector('#projectNameDisplayInventory');
                    if (projectNameDisplay)
                        projectNameDisplay.textContent = project.Title || "-";
                    const projectRef = document.querySelector('#projectRefInventory');
                    if (projectRef)
                        projectRef.textContent = project.ProjectReference || "-";
                    const profitCenter = document.querySelector('#profitCenterInventory');
                    if (profitCenter)
                        profitCenter.textContent = project.PC || "-";
                    const projectManager = document.querySelector('#projectManagerInventory');
                    if (projectManager) {
                        projectManager.textContent = typeof project.ProjectManager === "object" ?
                            (_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title : project.ProjectManager || "-";
                    }
                    const groupHead = document.querySelector('#groupHeadInventory');
                    if (groupHead)
                        groupHead.textContent = project.GroupHead || "-";
                    const customer = document.querySelector('#customerInventory');
                    if (customer)
                        customer.textContent = project.CustomerName || "-";
                };
                // ========== HISTORICAL TOGGLE FUNCTION ==========
                const setupHistoricalToggle = () => {
                    var _a, _b;
                    const toggleBtn = document.querySelector("#toggleHistoricalBtnInventory");
                    const historicalContainer = document.querySelector("#historicalContainerInventory");
                    const header = document.querySelector(".historical-toggle-header-inventory");
                    if (!toggleBtn || !historicalContainer || !header)
                        return;
                    // Set initial state - container should be HIDDEN by default
                    let isVisible = false;
                    historicalContainer.style.display = 'none';
                    toggleBtn.textContent = 'Show';
                    toggleBtn.style.background = '#28a745';
                    // Replace button to remove old listeners
                    const newBtn = document.createElement('button');
                    newBtn.id = 'toggleHistoricalBtnInventory';
                    newBtn.className = 'btn-secondary-inventory';
                    newBtn.textContent = 'Show';
                    newBtn.type = 'button';
                    newBtn.style.cssText = 'padding: 4px 16px; font-size: 13px; background: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; flex-shrink: 0; z-index: 10; position: relative;';
                    (_a = toggleBtn.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newBtn, toggleBtn);
                    // Replace header to remove all old listeners
                    const newHeader = header.cloneNode(true);
                    (_b = header.parentNode) === null || _b === void 0 ? void 0 : _b.replaceChild(newHeader, header);
                    // Button click handler
                    const btn = document.querySelector("#toggleHistoricalBtnInventory");
                    const container = document.querySelector("#historicalContainerInventory");
                    if (btn && container) {
                        btn.addEventListener("click", function (e) {
                            e.stopPropagation();
                            e.preventDefault();
                            if (isVisible) {
                                container.style.display = "none";
                                this.textContent = "Show";
                                this.style.background = "#28a745";
                            }
                            else {
                                container.style.display = "";
                                this.textContent = "Hide";
                                this.style.background = "#007bff";
                            }
                            isVisible = !isVisible;
                        });
                    }
                    // Header click handler (not the button)
                    const headerEl = document.querySelector(".historical-toggle-header-inventory");
                    if (headerEl) {
                        headerEl.addEventListener("click", function (e) {
                            // If click is on the button or its children, ignore
                            if (e.target.closest("#toggleHistoricalBtnInventory")) {
                                return;
                            }
                            e.preventDefault();
                            e.stopPropagation();
                            const btn = document.querySelector("#toggleHistoricalBtnInventory");
                            const container = document.querySelector("#historicalContainerInventory");
                            if (!container || !btn)
                                return;
                            const isCurrentlyHidden = container.style.display === "none";
                            if (isCurrentlyHidden) {
                                container.style.display = "";
                                btn.textContent = "Hide";
                                btn.style.background = "#007bff";
                                isVisible = true;
                            }
                            else {
                                container.style.display = "none";
                                btn.textContent = "Show";
                                btn.style.background = "#28a745";
                                isVisible = false;
                            }
                        });
                    }
                };
                // ========== ON PROJECT CHANGE ==========
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
                    const oblContainer = document.querySelector('.obl-section-container-inventory');
                    const fieldNote = oblContainer === null || oblContainer === void 0 ? void 0 : oblContainer.querySelector('.field-note-inventory');
                    if (isNewProject) {
                        if (oblContainer) {
                            oblContainer.className = 'obl-section-container-inventory visible-section-inventory new-project-inventory';
                        }
                        if (oblLastMonthInput) {
                            oblLastMonthInput.removeAttribute('readonly');
                            oblLastMonthInput.placeholder = "Enter initial OBL for new project";
                        }
                        if (fieldNote)
                            fieldNote.textContent = "For new projects: Enter initial OBL value";
                        if (oblValidationMessage) {
                            oblValidationMessage.textContent = "Enter initial OBL value for new project";
                            oblValidationMessage.className = 'validation-message-inventory';
                        }
                        const defaultOBL = project.PC ? 100000 : 0;
                        if (oblLastMonthInput)
                            setNumericValue(oblLastMonthInput, defaultOBL);
                    }
                    else {
                        if (oblContainer) {
                            oblContainer.className = 'obl-section-container-inventory visible-section-inventory existing-project-inventory';
                        }
                        if (oblLastMonthInput) {
                            oblLastMonthInput.setAttribute('readonly', 'true');
                            const oblValue = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                            setNumericValue(oblLastMonthInput, oblValue);
                            oblLastMonthInput.placeholder = "OBL from database";
                        }
                        if (fieldNote)
                            fieldNote.textContent = "OBL value from database record";
                        const oblValue = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                        if (oblLastMonthInput && oblValue > 0) {
                            oblLastMonthInput.value = oblValue.toString();
                            oblLastMonthInput.className = oblLastMonthInput.className.replace('input-empty-inventory', 'input-with-value-inventory');
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
                    if (oblLastMonthInput)
                        setNumericValue(oblLastMonthInput, displayOBL);
                    renderHistoricalMonths(allMonths.slice(0, 3), existingData);
                    renderFirstRow(allMonths.slice(3, 6), existingData);
                    renderRemainingMonths(allMonths.slice(6), existingData);
                    renderYearlyForecast(existingData.yearly || createFreshYearlyStructure(getLastMonthKey(existingData)));
                    protectHistoricalElements();
                    updateOBLValidationDisplay();
                    attachForecastEventListeners();
                    // Setup historical toggle after rendering - DEFAULT HIDDEN
                    setTimeout(() => {
                        setupHistoricalToggle();
                    }, 50);
                };
                // ========== RENDER FUNCTIONS FOR FORECAST SECTIONS ==========
                const renderHistoricalMonths = (months, existingData) => {
                    if (!historicalContainer)
                        return;
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
                    const planInputClass = data.plan > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    const actualInputClass = data.actual > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    return `
            <div class="month-box-inventory historical-month-inventory">
              <div class="month-header-inventory">${monthLabel} (Historical - Readonly)</div>

              <div class="month-block-inventory">
                <div class="main-inputs-container-inventory">
                  <div class="main-input-group-inventory">
                    <label>Previous</label>
                    <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                          readonly class="main-plan-input-inventory ${planInputClass} number-input-inventory" />
                  </div>
                  <div class="main-input-group-inventory">
                    <label>Current</label>
                    <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                          readonly class="main-current-input-inventory ${actualInputClass} number-input-inventory" />
                  </div>
                </div>
              </div>
            </div>
          `;
                };
                const renderFirstRow = (months, existingData) => {
                    if (!firstRowContainer)
                        return;
                    let html = '';
                    html += `<div class="detailed-months-row-inventory">`;
                    const currentMonth = months[0];
                    if (currentMonth) {
                        const monthDate = new Date(currentMonth + "-01");
                        const monthLabel = formatMonthLabel(monthDate);
                        const data = existingData[currentMonth] || { plan: 0, actual: 0 };
                        html += `<div class="current-month-container-inventory">`;
                        html += renderDetailedMonthBox(currentMonth, monthLabel, data, true);
                        html += `</div>`;
                    }
                    html += `<div class="next-months-container-inventory">`;
                    for (let i = 1; i <= 2; i++) {
                        if (i < months.length) {
                            const monthKey = months[i];
                            const monthDate = new Date(monthKey + "-01");
                            const monthLabel = formatMonthLabel(monthDate);
                            const data = existingData[monthKey] || { plan: 0, actual: 0 };
                            html += `<div class="next-month-box-inventory">`;
                            html += renderDetailedMonthBox(monthKey, monthLabel, data, false);
                            html += `</div>`;
                        }
                    }
                    html += `</div>`;
                    html += `</div>`;
                    firstRowContainer.innerHTML = html;
                };
                const renderRemainingMonths = (months, existingData) => {
                    if (!remainingContainer)
                        return;
                    let html = '';
                    for (let row = 0; row < 3; row++) {
                        html += `<div class="remaining-row-inventory">`;
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
                    if (!yearlyContainer)
                        return;
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
                    const planInputClass = data.plan > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    const actualInputClass = data.actual > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    return `
            <div class="yearly-box-inventory">
              <div class="yearly-header-inventory">${year}</div>

              <div class="yearly-inputs-container-inventory">
                <div class="yearly-input-group-inventory">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                        readonly class="yearly-plan-input-inventory ${planInputClass} number-input-inventory" />
                </div>
                <div class="yearly-input-group-inventory">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                        class="yearly-current-input-inventory ${actualInputClass} number-input-inventory" />
                </div>
              </div>
            </div>
          `;
                };
                const renderDetailedMonthBox = (monthKey, monthLabel, data, isCurrentMonth = false) => {
                    const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                    const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                    const displayForecastValue = data.forecast || 0;
                    const displayDateValue = data.date || '';
                    const planInputClass = data.plan > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    const actualInputClass = data.actual > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    const forecastInputClass = displayForecastValue > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    const dateInputClass = displayDateValue ? 'input-with-value-inventory' : 'input-empty-inventory';
                    return `
            <div class="detailed-month-box-inventory ${isCurrentMonth ? 'current-month-box-inventory' : 'next-month-box-inventory'}">
              <div class="month-header-inventory">
                ${monthLabel}
                ${isCurrentMonth ? '<span class="current-month-indicator-inventory">Current</span>' : ''}
              </div>

              <div class="month-block-inventory">
                <div class="enhanced-inputs-container-inventory ${isCurrentMonth ? 'four-col-inventory' : ''}">
                  <div class="enhanced-input-group-inventory">
                    <label>Previous</label>
                    <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                          readonly class="main-plan-input-inventory ${planInputClass} number-input-inventory" />
                  </div>

                  <div class="enhanced-input-group-inventory">
                    <label>Current</label>
                    <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                          class="main-current-input-inventory ${actualInputClass} number-input-inventory" />
                  </div>

                  ${isCurrentMonth ? `
                    <div class="enhanced-input-group-inventory">
                      <label>Forecast</label>
                      <input type="text" value="${displayForecastValue}" data-type="forecast" data-month="${monthKey}"
                            class="main-forecast-input-inventory ${forecastInputClass} number-input-inventory" />
                    </div>

                    <div class="enhanced-input-group-inventory">
                      <label>Date</label>
                      <input type="date" value="${displayDateValue}" data-type="date" data-month="${monthKey}"
                            class="forecast-date-input-inventory ${dateInputClass}" />
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
                    const planInputClass = data.plan > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    const actualInputClass = data.actual > 0 ? 'input-with-value-inventory' : 'input-empty-inventory';
                    return `
            <div class="month-box-inventory simple-month-inventory">
              <div class="month-header-inventory">${monthLabel}</div>

              <div class="main-inputs-container-inventory">
                <div class="main-input-group-inventory">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input-inventory ${planInputClass} number-input-inventory" />
                </div>
                <div class="main-input-group-inventory">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" id="actualCurrent" data-type="actual" data-month="${monthKey}"
                        class="main-current-input-inventory ${actualInputClass} number-input-inventory" />
                </div>
              </div>
            </div>
          `;
                };
                const protectHistoricalElements = () => {
                    document.querySelectorAll('.historical-month-inventory input, .historical-month-inventory select, .historical-month-inventory button').forEach(element => {
                        var _a;
                        const newElement = element.cloneNode(true);
                        (_a = element.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newElement, element);
                    });
                    document.querySelectorAll('.historical-month-inventory').forEach(section => {
                        section.addEventListener('click', (e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            toast("warning", "Historical Inventory data cannot be modified");
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
                    document.querySelectorAll('.number-input-inventory').forEach(input => {
                        allowOnlyNumbers(input);
                    });
                    document.querySelectorAll('.number-input-inventory:not([readonly])').forEach(input => {
                        input.addEventListener('focus', function () {
                            this.setAttribute('data-old-value', getNumericValue(this).toString());
                        });
                    });
                    document.querySelectorAll('.main-current-input-inventory:not([readonly])').forEach(input => {
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
                                this.className = this.className.replace('input-with-value-inventory', 'input-empty-inventory');
                            }
                            else {
                                this.className = this.className.replace('input-empty-inventory', 'input-with-value-inventory');
                            }
                            updateOBLValidationDisplay(this);
                        });
                    });
                    document.querySelectorAll('.main-forecast-input-inventory:not([readonly])').forEach(input => {
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
                                this.className = this.className.replace('input-with-value-inventory', 'input-empty-inventory');
                            }
                            else {
                                this.className = this.className.replace('input-empty-inventory', 'input-with-value-inventory');
                            }
                            updateOBLValidationDisplay(this);
                        });
                    });
                    document.querySelectorAll('.yearly-current-input-inventory:not([readonly])').forEach(input => {
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
                                this.className = this.className.replace('input-with-value-inventory', 'input-empty-inventory');
                            }
                            else {
                                this.className = this.className.replace('input-empty-inventory', 'input-with-value-inventory');
                            }
                            updateOBLValidationDisplay(this);
                        });
                    });
                    document.querySelectorAll('.forecast-date-input-inventory:not([readonly])').forEach(input => {
                        input.addEventListener('change', function () {
                            if (this.value === '') {
                                this.className = this.className.replace('input-with-value-inventory', 'input-empty-inventory');
                            }
                            else {
                                this.className = this.className.replace('input-empty-inventory', 'input-with-value-inventory');
                            }
                        });
                    });
                    const validationInputs = document.querySelectorAll(`
            .main-current-input-inventory:not([readonly]),
            .yearly-current-input-inventory:not([readonly]),
            .main-forecast-input-inventory:not([readonly])
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
                        toast("error", "Inventory forecast not found");
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
                        toast("warning", "Corrupted Inventory forecast data detected. Loaded empty structure.");
                    }
                    forecastData = propagateMonthValues(forecastData);
                    if (forecast.Project && projectSelect) {
                        projectSelect.value = String(forecast.Project.Id);
                        onProjectChange(forecast.Project.Id, forecast);
                    }
                    // ✅ Apply permission-based readonly state AFTER rendering
                    setTimeout(() => {
                        // If frozen OR user can't edit, set read-only
                        const isReadOnly = isFrozen || !canEdit;
                        setFormReadOnly(isReadOnly);
                        // Show/hide save button
                        const saveBtn = document.querySelector('.btn-primary-inventory[type="submit"]');
                        if (saveBtn) {
                            saveBtn.style.display = isReadOnly ? 'none' : '';
                        }
                        // Update form title
                        const formTitleEl = document.querySelector("#formTitleInventory");
                        if (formTitleEl) {
                            if (isFrozen) {
                                formTitleEl.innerHTML = '🔒 Edit Inventory Forecast (Frozen - Read Only)';
                                formTitleEl.style.color = '#856404';
                            }
                            else if (!canEdit) {
                                formTitleEl.innerHTML = '🔒 Edit Inventory Forecast (Read Only)';
                                formTitleEl.style.color = '#856404';
                            }
                            else {
                                formTitleEl.innerHTML = 'Edit Inventory Forecast';
                                formTitleEl.style.color = '';
                            }
                        }
                    }, 300);
                    toast("info", `Editing Inventory forecast for ${forecast.Title}`);
                };
                const loadProjects = async () => {
                    var _a;
                    try {
                        console.log("Loading projects from list:", projectListName);
                        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items?$select=Id,Title,CustomerName,GroupHead,PC,ProjectReference,ProjectManager/Title&$expand=ProjectManager`, {
                            headers: { Accept: "application/json;odata=verbose" }
                        });
                        if (!res.ok) {
                            throw new Error(`Failed to load projects: ${res.status} ${res.statusText}`);
                        }
                        const data = await res.json();
                        const rawProjects = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                        const cleanedProjects = rawProjects.map((p) => (Object.assign(Object.assign({}, p), { CustomerName: p.CustomerName || "", GroupHead: p.GroupHead || "", PC: p.PC || "", ProjectReference: p.ProjectReference || "", ProjectManager: p.ProjectManager || "" })));
                        const projectMap = new Map();
                        cleanedProjects.forEach(p => {
                            projectMap.set(p.Title, p);
                        });
                        projects = Array.from(projectMap.values());
                        // ✅ Use getAssignedProjects to filter by permissions
                        const assignedProjects = getAssignedProjects(projects);
                        console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);
                        // ✅ Populate project select dropdown with assigned projects ONLY
                        if (projectSelect) {
                            projectSelect.innerHTML = `<option value="">Select Project</option>` +
                                assignedProjects.map(p => `<option value="${p.Id}">${p.Title}</option>`).join("");
                        }
                        // ✅ Also update the filter dropdown with assigned projects
                        if (filterProject) {
                            filterProject.innerHTML = '<option value="">All Projects</option>' +
                                [...new Set(assignedProjects.map(p => p.Title))]
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
                        console.log("Loading forecasts from list:", listName);
                        const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,InventoryActual,CreatedDate,UpdatedDate&$expand=Project&$top=5000`, {
                            headers: { Accept: "application/json;odata=verbose" }
                        });
                        if (!res.ok) {
                            throw new Error(`Failed to load forecasts: ${res.status} ${res.statusText}`);
                        }
                        // ✅ Parse the response into a variable
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
                // ========== FORM SUBMISSION ==========
                if (form) {
                    form.addEventListener("submit", async (e) => {
                        e.preventDefault();
                        const projectId = parseInt(projectSelect ? projectSelect.value : "0");
                        const project = projects.find(p => p.Id === projectId);
                        if (!project) {
                            toast("error", "Please select a project");
                            return;
                        }
                        if (editingId) {
                            const existingForecast = forecasts.find(f => f.Id === editingId);
                            if (!canEditProject(project)) {
                                toast("error", "❌ You don't have permission to edit this Inventory forecast");
                                return;
                            }
                            if (existingForecast && existingForecast.Freeze) {
                                toast("error", "❌ Cannot edit: This project is frozen!");
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
                                if (oblLastMonthInput)
                                    oblLastMonthInput.focus();
                                return;
                            }
                        }
                        const totalCurrent = calculateTotalCurrentValue();
                        if (oblValue > 0 && totalCurrent > oblValue) {
                            toast("error", "Total current value cannot exceed OBL!");
                            return;
                        }
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
                                        monthData.date = mainDateInput.value;
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
                            const propagatedData = propagateCollectionValues(forecastData);
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
                                initializePagination(forecasts);
                                hideForm();
                                resetForm();
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
                    });
                }
                // ========== INITIALIZE ==========
                const initialize = async () => {
                    try {
                        // ✅ Initialize shared authentication FIRST
                        await authManager.initialize({ siteUrl });
                        console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                        console.log(`👤 User: ${getCurrentUserFullName()}`);
                        console.log(`📋 User Groups:`, getUserGroups());
                        console.log(`✅ Has Full Control: ${hasFullControl()}`);
                        console.log(`✅ Can Contribute: ${canContribute()}`);
                        console.log(`✅ Can Freeze: ${canFreeze()}`);
                        console.log(`✅ Is GH Head: ${isGHHead()}`);
                        // Show loading state
                        if (tableBody) {
                            tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">Loading Inventory data...</td></tr>`;
                        }
                        // Load data
                        await loadProjects();
                        await loadForecasts();
                        // ✅ Apply auto progression if needed
                        if (forecasts.length > 0) {
                            await applyAutoProgression();
                            // Reload forecasts after progression
                            await loadForecasts();
                        }
                        // Initialize filters
                        initializeFilters();
                        // ✅ Call filterTable() to apply project filtering
                        filterTable();
                        // ✅ Update UI based on permissions
                        updateUIBasedOnPermissions({
                            exportBtn: '.export-btn-inventory',
                            addBtn: '.show-form-btn-inventory',
                            saveBtn: '.btn-primary-inventory[type="submit"]',
                            freezeCheckboxes: '.freeze-checkbox'
                        });
                        console.log("✅ Inventory application initialized successfully");
                    }
                    catch (error) {
                        console.error("Error initializing Inventory management:", error);
                        toast("error", "Failed to initialize Inventory management: " + error.message);
                    }
                };
                // ========== EVENT LISTENERS ==========
                if (backBtn) {
                    backBtn.addEventListener("click", hideForm);
                }
                if (projectSelect) {
                    projectSelect.addEventListener("change", (e) => {
                        if (editingId) {
                            toast("warning", "Cannot change project in Inventory edit mode");
                            const forecast = forecasts.find(f => f.Id === editingId);
                            if ((forecast === null || forecast === void 0 ? void 0 : forecast.Project) && projectSelect) {
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
                // Pagination event listeners
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
                // Filter event listeners
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
                // Action button event listeners
                if (exportBtn)
                    exportBtn.addEventListener("click", exportToExcelWithFullStyle);
                if (clearFiltersBtn)
                    clearFiltersBtn.addEventListener("click", clearFilters);
                // Initialize the application
                await initialize();
            }
            catch (error) {
                console.error("Fatal error in inventory script:", error);
                const tableBody = document.querySelector(".inventory-table-body");
                if (tableBody) {
                    tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;color:red;">Error loading Inventory data. Please check console for details.</td></tr>`;
                }
            }
        }, 100);
    }
};
//# sourceMappingURL=inventory.js.map
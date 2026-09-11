import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/invoicing.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canViewForm, getAssignedProjects, getAssignedProjectValues, updateUIBasedOnPermissions, getDigest as getAuthDigest,
// getListEntityType
 } from "./auth";
let currentSortConfig = {
    column: '',
    direction: 'asc'
};
export default {
    content: `
<div class="invoicing-page fade-in">
  <section class="table-section card">
    <div class="page-header">
      <h1>Invoicing Plan++</h1>
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
      
      <!-- Second Row: Year, Month, Export Excel -->
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
          <button class="btn-secondary export-btn">📊 Export Excel</button>
        </div>
        <div class="tableColp4"> 
    <label>&nbsp;</label>
    <button class="btn-secondary export-details-btn">📋 Details Export Excel</button>
  </div>
      </div>
    </div>
       
    <div class="table-container-inv">
      <div style="overflow-x:auto;">
        <table class="month-columns-table-inv">
          <thead id="totalsHeader"></thead>
          <thead id="monthColumnsHeader"></thead>
          <tbody class="invoicing-table-body-inv"></tbody>
        </table>
      </div>
      
      <!-- Pagination -->
      <div class="pagination-container">
        <div class="pagination-left">
          <button class="pagination-btn" id="firstPage" disabled>First</button>
          <button class="pagination-btn" id="prevPage" disabled>Previous</button>
          
          <div class="pagination-info">
            Page <span id="currentPage">1</span> of <span id="totalPages">1</span>
          </div>
          
          <div class="pagination-pages" id="pageNumbers"></div>
          
          <button class="pagination-btn" id="nextPage" disabled>Next</button>
          <button class="pagination-btn" id="lastPage" disabled>Last</button>
        </div>
  
        <div class="pagination-right">
          <div class="page-size-selector">
            <label for="pageSizeSelect">Show:</label>
            <select id="pageSizeSelect" class="page-size-select">
              <option value="10">10</option>
              <option value="25">25</option>
              <option value="50">50</option>
              <option value="100">100</option>
            </select>
            <span>per page</span>
          </div>
          
          <div class="pagination-info">
            Showing <span id="startItem">0</span>-<span id="endItem">0</span> of <span id="totalItems">0</span> records
          </div>
        </div>
      </div>
    </div>
  </section>

  <section class="form-section card" style="display:none;">
    <div class="page-header">
      <h1 id="formTitle">Edit Forecast</h1>
    </div>
    <form id="invoicingForm">
      <!-- Project Selection (hidden in edit mode) -->
      <div class="form-row hidden-section">
        <div class="form-group project-select-container">
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
          <div class="obl-section-container-invoicing hidden-section" id="oblContainerInvoicing">
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
              ✅ All values are within OBL limit
            </div>
          </div>
        </div>

      </div>
      
      <!-- Historical Months (Last 3 Months - Readonly) -->
       <div class="forecast-section hidden-section" id="historicalForecastSection">
        <h3 class="historical-toggle-header">
          Historical Data (Last 3 Months - Readonly)
          <button type="button" class="btn-secondary1 historical-toggle-btn" id="historicalToggleBtn">Show</button>
        </h3>
        <div id="historicalContainer" class="historical-container" style="display:none;"></div>
      </div>
      
      <!-- Current + Next 2 Months - Hidden by default -->
      <div class="forecast-section hidden-section" id="detailedForecastSection">
      <h3>Detailed Forecast (Current + Next 2 Months)</h3>
      <div id="firstRowContainer" class="first-row-container">
         
      </div>
    </div>
      
      <!-- Next 15 Months - Hidden by default -->
      <div class="forecast-section hidden-section" id="longtermForecastSection">
        <h3>Long-term Forecast (Next 15 Months)</h3>
        <div id="remainingContainer" class="remaining-container"></div>
      </div>

      <!-- 5-Year Forecast Section - Hidden by default -->
      <div class="forecast-section hidden-section" id="yearlyForecastSection">
        <h3>5-Year Forecast</h3>
        <div id="yearlyContainer" class="yearly-container-inv"></div>
      </div>
      
      <div class="form-actions">
        <button type="submit" class="btn-primary">💾 Save Forecast</button>
        <button type="button" class="btn-secondary back-to-table">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,
    style: `
/* ----- Core Layout & Components ----- */
.hidden-section {
  display: none !important;
}

/* Add to your existing CSS */
.export-details-btn {
  background: #6f42c1 !important;
}

.export-details-btn:hover {
  background: #5e35b1 !important;
}

.edit-action-btn {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 3px;
  padding: 4px 6px;
  background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
  color: white;
  border: none;
  border-radius: 4px;
  font-size: 9px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s ease;
  min-width: 50px;
  box-shadow: 0 1px 2px rgba(0,0,0,0.1);
  line-height: 1.1;
}

.edit-action-btn:hover {
  background: linear-gradient(135deg, #106ebe 0%, #005a9e 100%);
  transform: translateY(-1px);
  box-shadow: 0 2px 4px rgba(0,0,0,0.15);
}

.visible-section {
  display: block !important;
}

/* Ensure forecast sections have proper display when visible */
.forecast-section.visible-section {
  display: block;
}

.obl-section-container.visible-section {
  display: block;
}

.card {
  background: #fff;
  border-radius: 10px;
  padding: 20px;
  margin-bottom: 20px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}

.detailed-month {
  position: relative;
}

/* ============================================ */
/* HISTORICAL DATA TOGGLE - FIXED               */
/* ============================================ */

.historical-toggle-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 15px;
}

.historical-toggle-btn {
  padding: 6px 16px;
  font-size: 13px;
  background: #28a745 !important;
  color: white !important;
  border: none;
  border-radius: 5px;
  cursor: pointer;
  transition: all 0.2s ease;
  font-weight: 600;
}

.historical-toggle-btn:hover {
  background: #218838 !important;
  transform: translateY(-1px);
}

.historical-container {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 15px;
  margin-top: 10px;
}

/* ============================================ */
/* PROJECT DETAILS - FIXED                      */
/* ============================================ */

.project-main-wrapper-invoicing {
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

/* LEFT SIDE */
.project-left-side-invoicing {
  padding: 8px 16px !important;
  background: #ffffff !important;
  border-right: 1px solid #e9ecef !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 0 !important;
  justify-content: flex-start !important;
  align-items: flex-start !important;
}

.project-left-side-invoicing .detail-item-invoicing {
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

.project-left-side-invoicing .detail-label-invoicing {
  font-weight: 700 !important;
  white-space: nowrap !important;
  min-width: 100px !important;
}

.project-left-side-invoicing .detail-value-invoicing {
  white-space: nowrap !important;
}

/* RIGHT SIDE */
.project-right-side-invoicing {
  padding: 8px 16px !important;
  background: #ffffff !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 6px !important;
  justify-content: flex-start !important;
}

/* ============================================ */
/* OBL SECTION                                 */
/* ============================================ */

.obl-section-container-invoicing {
  margin: 2px 0 0 0 !important;
  padding: 6px 12px !important;
  background: #f8f9fa !important;
  border-radius: 4px !important;
  border: 1px solid #e9ecef !important;
}

.obl-content-invoicing {
  display: grid !important;
  grid-template-columns: 1fr 1.5fr !important;
  gap: 12px !important;
  align-items: start !important;
}

.obl-left-invoicing {
  display: flex !important;
  flex-direction: row !important;
}

.obl-label-group-invoicing {
  display: flex !important;
  align-items: center !important;
  gap: 4px !important;
  margin-bottom: 2px !important;
  text-align: left !important;
}

.obl-label-group-invoicing label {
  font-size: 12px !important;
  font-weight: 700 !important;
  color: #495057 !important;
  margin: 0 !important;
}

.obl-value-display-invoicing {
  display: flex !important;
  flex-direction: column !important;
}

.obl-input-invoicing {
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

.obl-right-invoicing {
  display: flex !important;
  align-items: center !important;
}

.validation-summary-invoicing {
  background: #fff !important;
  padding: 4px 12px !important;
  border-radius: 4px !important;
  border: 1px solid #e9ecef !important;
  width: 100% !important;
}

.validation-item-invoicing {
  display: flex !important;
  justify-content: space-between !important;
  margin-bottom: 1px !important;
  font-size: 11px !important;
  text-align: left !important;
}

.validation-label-invoicing {
  font-weight: 600 !important;
  color: #666 !important;
  font-size: 10px !important;
  text-align: left !important;
}

.validation-value-invoicing {
  font-weight: 700 !important;
  font-size: 11px !important;
  color: #333 !important;
  text-align: right !important;
}

.validation-line-invoicing {
  height: 4px !important;
  background: #e9ecef !important;
  border-radius: 2px !important;
  margin-top: 3px !important;
  overflow: hidden !important;
}

.line-fill-invoicing {
  height: 100% !important;
  background: #28a745 !important;
  border-radius: 2px !important;
  transition: all 0.3s ease !important;
  width: 0% !important;
}

.line-fill-invoicing.warning-invoicing {
  background: #ffc107 !important;
}

.line-fill-invoicing.danger-invoicing {
  background: #dc3545 !important;
}

.validation-message-invoicing {
  display: block !important;
  padding: 8px 12px;
  border-radius: 4px;
  margin-top: 5px;
  font-size: 12px;
  font-weight: 500;
}

.validation-message-invoicing.danger-invoicing {
  background-color: #f8d7da;
  color: #721c24;
  border: 1px solid #f5c6cb;
}

.validation-message-invoicing.warning-invoicing {
  background-color: #fff3cd;
  color: #856404;
  border: 1px solid #ffc107;
}

/* ============================================ */
/* RESPONSIVE - PROJECT MAIN WRAPPER            */
/* ============================================ */

@media (max-width: 992px) {
  .project-main-wrapper-invoicing {
    grid-template-columns: 1fr !important;
  }
  
  .project-left-side-invoicing {
    border-right: none !important;
    border-bottom: 1px solid #e9ecef !important;
  }
  
  .project-right-side-invoicing {
    padding: 8px 14px !important;
  }
}

@media (max-width: 768px) {
  .obl-content-invoicing {
    grid-template-columns: 1fr !important;
    gap: 8px !important;
  }
  
  .obl-input-invoicing {
    max-width: 100% !important;
  }
  
  .project-left-side-invoicing .detail-label-invoicing {
    min-width: 80px !important;
  }
}

/* ============================================ */
/* TABLE ACTIONS & FILTERS                     */
/* ============================================ */

.table-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 15px;
  margin-bottom: 20px;
  padding: 15px;
  background: #f8f9fa;
  border-radius: 8px;
  border: 1px solid #e9ecef;
}

.table-actions-row {
  display: grid;
  grid-template-columns: repeat(6, 1fr);
  gap: 15px;
  align-items: end;
}

.tableColp4 {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.tableColp4 label {
  font-size: 12px;
  font-weight: 600;
  color: #495057;
  margin-bottom: 4px;
}

.tableColp4 select {
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 5px;
  font-size: 14px;
  background: #fff;
  width: 100%;
}

/* ============================================ */
/* BUTTONS                                     */
/* ============================================ */

.btn-primary {
  background: #0078d4;
  color: white;
  padding: 8px 12px;
  border: none;
  border-radius: 5px;
  font-size: 14px;
  cursor: pointer;
  transition: all 0.2s ease;
}

.btn-primary:hover:not(:disabled) {
  background: #106ebe;
  transform: translateY(-1px);
}

.btn-primary:disabled {
  background: #a19f9d;
  cursor: not-allowed;
}

.clear-filters-btn {
  background: #dc3545 !important;
}

.clear-filters-btn:hover {
  background: #c82333 !important;
}

.import-actual-btn {
  background: #28a745 !important;
}

.import-actual-btn:hover {
  background: #218838 !important;
}

.export-btn {
  background: #17a2b8 !important;
}

.export-btn:hover {
  background: #138496 !important;
}

.export-details-btn {
  background: #6f42c1 !important;
}

.export-details-btn:hover {
  background: #5e35b1 !important;
}

.btn-secondary {
  background: #6c757d;
  color: white;
  padding: 8px 12px;
  border: none;
  border-radius: 5px;
  font-size: 14px;
  cursor: pointer;
  transition: all 0.2s ease;
}

.btn-secondary:hover:not(:disabled) {
  background: #5a6268;
  transform: translateY(-1px);
}

/* ============================================ */
/* FORECAST SECTIONS                           */
/* ============================================ */

.forecast-section {
  margin: 25px 0;
  padding: 20px;
  background: #f8f9fa;
  border-radius: 10px;
  border: 1px solid #e9ecef;
}

.forecast-section h3 {
  margin: 0 0 15px 0;
  color: #0078d4;
  font-size: 16px;
  font-weight: 600;
}

.first-row-container {
  display: grid;
  grid-template-columns: repeat(1, 1fr);
  gap: 15px;
}

/* ============================================ */
/* LONG-TERM FORECAST - NEXT 15 MONTHS         */
/* ============================================ */

.remaining-container {
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.remaining-row-inv {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 15px;
}

/* ============================================ */
/* MONTH BOX STYLES                            */
/* ============================================ */

.month-box-inv {
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  padding: 15px;
  background: #fff;
  box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  position: relative;
}

.month-header-inv {
  text-align: center;
  font-weight: 600;
  color: #0078d4;
  margin-bottom: 15px;
  padding-bottom: 8px;
  border-bottom: 1px solid #eee;
}

/* ============================================ */
/* HISTORICAL MONTH - READONLY                 */
/* ============================================ */

.historical-month {
  background: linear-gradient(135deg, #f5f5f5 0%, #e8e8e8 100%) !important;
  border: 1px solid #d0d0d0 !important;
  pointer-events: none !important;
  user-select: none !important;
  opacity: 0.8;
}

.historical-month * {
  pointer-events: none !important;
  user-select: none !important;
}

.historical-month::before {
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

.historical-month .month-header-inv {
  background: #6c757d !important;
  color: white !important;
  padding: 8px;
  margin: -15px -15px 15px -15px;
  border-radius: 8px 8px 0 0;
  border-bottom: 1px solid #ccc;
}

.historical-month input {
  background-color: #f8f9fa !important;
  color: #6c757d !important;
  border-color: #ced4da !important;
  cursor: not-allowed !important;
}

/* Historical Items Section - 3 columns only */
.historical-items-section .historical-items-header {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr !important;
  gap: 8px;
  margin-bottom: 8px;
  font-size: 11px;
  color: #666;
  text-align: center;
  font-weight: 600;
  padding-bottom: 4px;
  border-bottom: 1px solid #eee;
}

.historical-items-section .historical-item-row {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr !important;
  gap: 8px;
  margin-bottom: 6px;
  align-items: center;
}

/* Ensure historical items don't show forecast/date columns */
.historical-items-section .historical-items-header div:nth-child(4),
.historical-items-section .historical-items-header div:nth-child(5),
.historical-items-section .historical-item-row input:nth-child(4),
.historical-items-section .historical-item-row input:nth-child(5) {
  display: none !important;
}

/* Historical main inputs - 2 columns only */
.historical-month .main-inputs-container-inv {
  grid-template-columns: 1fr 1fr !important;
}

.historical-month .main-input-group-inv:nth-child(3),
.historical-month .main-input-group-inv:nth-child(4) {
  display: none !important;
}

/* ============================================ */
/* INPUT GROUPS                                */
/* ============================================ */

.main-inputs-container-inv {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 15px;
  margin-bottom: 15px;
}

.main-input-group-inv {
  display: flex;
  flex-direction: column;
  align-items: center;
}

.main-input-group-inv label {
  font-size: 12px;
  color: #666;
  margin-bottom: 6px;
  font-weight: 600;
}

.main-input-group-inv input {
  width: 100%;
  padding: 8px 10px;
  border: 1px solid #ddd;
  border-radius: 4px;
  text-align: center;
  font-size: 14px;
  background: #fff;
  transition: all 0.2s ease;
}

.main-input-group-inv input:focus {
  border-color: #0078d4;
  box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
  outline: none;
}

/* ============================================ */
/* ITEMS SECTION - DEFAULT (3 columns)          */
/* ============================================ */

.items-section {
  margin-top: 15px;
  padding-top: 15px;
  border-top: 2px solid #e0e0e0;
}

.items-header {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr !important;
  gap: 8px;
  margin-bottom: 8px;
  font-size: 11px;
  color: #666;
  text-align: center;
  font-weight: 600;
  padding-bottom: 4px;
  border-bottom: 1px solid #eee;
}

.item-row {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr !important;
  gap: 8px;
  margin-bottom: 6px;
  align-items: center;
}

.item-row input {
  padding: 5px 8px;
  border: 1px solid #ddd;
  border-radius: 3px;
  font-size: 12px;
  background: #fff;
  transition: all 0.2s ease;
  width: 100%;
  box-sizing: border-box;
}

.item-row input[type="text"] {
  text-align: left;
}
.item-row input[type="number"] {
  text-align: right;
}

.item-row input:focus {
  border-color: #0078d4;
  box-shadow: 0 0 0 1px rgba(0,120,212,0.2);
  outline: none;
}

/* ============================================ */
/* DETAILED FORECAST - FIXED LAYOUT             */
/* ============================================ */

/* Main row: Current month + Next 2 months */
.detailed-months-row {
  display: flex;
  gap: 24px;
  width: 100%;
  margin-bottom: 20px;
  box-sizing: border-box;
}

/* ============================================ */
/* CURRENT MONTH - 44%                          */
/* ============================================ */

.current-month-container {
  flex: 0 0 44%;
  width: 44%;
  max-width: 44%;
  min-width: 0;
  box-sizing: border-box;
}

.current-month-container .detailed-month-box {
  width: 100%;
  box-sizing: border-box;
}

/* Current month main inputs - 4 columns */
.current-month-container .enhanced-inputs-container-5col {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr)) !important;
  gap: 14px;
  width: 100%;
  margin-bottom: 15px;
  box-sizing: border-box;
}

.current-month-container .enhanced-inputs-container-5col .enhanced-input-group {
  min-width: 0;
}

.current-month-container .enhanced-inputs-container-5col label {
  text-align: center;
  white-space: nowrap;
}

.current-month-container .enhanced-inputs-container-5col input {
  width: 100%;
  min-width: 0;
  box-sizing: border-box;
}

/* ============================================ */
/* CURRENT MONTH ITEMS - 5 COLUMNS              */
/* ============================================ */

.current-month-container .items-header-5col {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr 1fr 1.4fr !important;
  gap: 10px;
  width: 100%;
  padding: 8px 4px;
  margin-bottom: 5px;
  box-sizing: border-box;
  font-size: 12px;
  font-weight: 600;
  color: #495057;
  border-bottom: 1px solid #dee2e6;
  text-align: center;
}

.current-month-container .item-row-5col {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr 1fr 1.4fr !important;
  gap: 10px;
  width: 100%;
  margin-bottom: 7px;
  align-items: center;
  box-sizing: border-box;
}

.current-month-container .item-row-5col input {
  width: 100%;
  min-width: 0;
  box-sizing: border-box;
  padding: 7px 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 12px;
  text-align: center;
}

.current-month-container .item-row-5col input:first-child {
  text-align: left;
}

/* ============================================ */
/* NEXT TWO MONTHS - 56%                        */
/* ============================================ */

.next-months-container {
  flex: 1 1 56%;
  width: 56%;
  max-width: 56%;
  min-width: 0;
  display: flex;
  gap: 20px;
  box-sizing: border-box;
}

/* Each next month gets equal width */
.next-month-box {
  flex: 1 1 0;
  min-width: 0;
  width: 50%;
  box-sizing: border-box;
}

.next-month-box .detailed-month-box {
  width: 100%;
  box-sizing: border-box;
}

/* ============================================ */
/* NEXT MONTH MAIN INPUTS - 2 COLUMNS           */
/* ============================================ */

.next-month-box .enhanced-inputs-container-3col {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
  gap: 14px;
  width: 100%;
  margin-bottom: 15px;
  box-sizing: border-box;
}

.next-month-box .enhanced-inputs-container-3col .enhanced-input-group {
  min-width: 0;
}

.next-month-box .enhanced-inputs-container-3col input {
  width: 100%;
  min-width: 0;
  box-sizing: border-box;
}

/* ============================================ */
/* NEXT MONTH ITEMS - 3 COLUMNS                 */
/* ============================================ */

.next-month-box .items-header-3col {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr !important;
  gap: 10px;
  width: 100%;
  padding: 8px 4px;
  margin-bottom: 5px;
  box-sizing: border-box;
  font-size: 12px;
  font-weight: 600;
  color: #495057;
  border-bottom: 1px solid #dee2e6;
  text-align: center;
}

.next-month-box .item-row-3col {
  display: grid;
  grid-template-columns: 2fr 1fr 1fr !important;
  gap: 10px;
  width: 100%;
  margin-bottom: 7px;
  align-items: center;
  box-sizing: border-box;
}

.next-month-box .item-row-3col input {
  width: 100%;
  min-width: 0;
  box-sizing: border-box;
  padding: 7px 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 12px;
  text-align: center;
}

.next-month-box .item-row-3col input:first-child {
  text-align: left;
}

/* ============================================ */
/* MONTH BOX                                    */
/* ============================================ */

.detailed-month-box {
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  padding: 20px;
  background: #fff;
  box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  width: 100%;
  box-sizing: border-box;
  height: fit-content;
  overflow: hidden;
}

/* Current month */
.current-month-box {
  background: linear-gradient(135deg, #ffffff 0%, #f0f8ff 100%);
  border: 2px solid #0078d4;
}

/* Next months */
.next-month-box .detailed-month-box {
  background: linear-gradient(135deg, #ffffff 0%, #f9f9f9 100%);
  border: 1px solid #ddd;
}

/* ============================================ */
/* MONTH HEADER                                 */
/* ============================================ */

.month-header-inv {
  font-weight: 700;
  font-size: 16px;
  color: #333;
  margin-bottom: 14px;
  padding-bottom: 8px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-bottom: 2px solid #e9ecef;
  gap: 8px !important;
}

.current-month-box .month-header-inv {
  color: #0078d4;
}

.current-month-indicator {
  display: inline-block;
  font-size: 10px;
  background-color: #ff9800;
  color: white;
  padding: 1px 6px;
  border-radius: 10px;
  font-weight: bold;
  
}

/* ============================================ */
/* VALIDATION BAR                               */
/* ============================================ */

.month-total-validation {
  width: 100%;
  box-sizing: border-box;
  padding: 10px 12px;
  border-radius: 6px;
  margin: 10px 0;
  font-size: 12px;
  font-weight: 600;
  text-align: center;
}

.month-total-valid {
  background-color: #d4edda;
  color: #155724;
  border: 1px solid #c3e6cb;
}

.month-total-invalid {
  background-color: #f8d7da;
  color: #721c24;
  border: 1px solid #f5c6cb;
}

/* ============================================ */
/* ITEMS SECTION                                */
/* ============================================ */

.items-section {
  width: 100%;
  margin-top: 12px;
  border-top: 1px solid #dee2e6;
  padding-top: 10px;
  box-sizing: border-box;
}

/* ============================================ */
/* ENHANCED INPUTS - FIXED FOR DETAILED        */
/* ============================================ */

.enhanced-inputs-container-invoice {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px;
  margin-bottom: 15px;
}

.enhanced-input-group {
  display: flex;
  flex-direction: column;
  align-items: center;
}

.enhanced-input-group.full-width {
  grid-column: 1 / -1;
}

.enhanced-input-group label {
  font-size: 11px;
  color: #666;
  margin-bottom: 4px;
  font-weight: 600;
  text-align: center;
}

.enhanced-input-group input,
.enhanced-input-group select {
  width: 100%;
  padding: 6px 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  text-align: center;
  font-size: 12px;
  background: #fff;
  transition: all 0.2s ease;
}

.enhanced-input-group input:focus,
.enhanced-input-group select:focus {
  border-color: #0078d4;
  box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
  outline: none;
}

.forecast-date-input {
  background: #fff !important;
  color: #333 !important;
  border: 1px solid #ddd !important;
}

/* ============================================ */
/* FORMS                                       */
/* ============================================ */

.form-row {
  display: flex;
  flex-wrap: wrap;
  gap: 20px;
  margin-bottom: 20px;
}

.form-group {
  flex: 1 1 250px;
  display: flex;
  flex-direction: column;
}

.form-group label {
  font-weight: 600;
  margin-bottom: 5px;
  color: #333;
}

.form-group input,
.form-group select {
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 5px;
  font-size: 14px;
}

.form-actions {
  display: flex;
  gap: 15px;
  margin-top: 20px;
  padding-top: 20px;
  border-top: 1px solid #e9ecef;
}

/* ============================================ */
/* YEARLY FORECAST                             */
/* ============================================ */

.yearly-container-inv {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 15px;
  margin-top: 15px;
}

.yearly-box-inv {
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  padding: 15px;
  background: #fff;
  box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  position: relative;
}

.yearly-header-inv {
  text-align: center;
  font-weight: 600;
  color: #0078d4;
  margin-bottom: 15px;
  padding-bottom: 8px;
  border-bottom: 1px solid #eee;
}

.yearly-inputs-container-inv {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 15px;
}

.yearly-input-group-inv {
  display: flex;
  flex-direction: column;
  align-items: center;
}

.yearly-input-group-inv label {
  font-size: 12px;
  color: #666;
  margin-bottom: 6px;
  font-weight: 600;
}

.yearly-input-group-inv input {
  width: 100%;
  padding: 8px 10px;
  border: 1px solid #ddd;
  border-radius: 4px;
  text-align: center;
  font-size: 14px;
  background: #fff;
  transition: all 0.2s ease;
}

.yearly-input-group-inv input:focus {
  border-color: #0078d4;
  box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
  outline: none;
}

.yearly-plan-input-inv {
  background-color: #f8f9fa !important;
  color: #6c757d !important;
  border-color: #ced4da !important;
}

input.yearly-current-input-inv {
  background: #fff !important;
}

/* ============================================ */
/* READONLY FIELDS                             */
/* ============================================ */

.readonly-field {
  background-color: #f5f5f5 !important;
  color: #666 !important;
  border-color: #ddd !important;
}

.historical-month input,
input[readonly] {
  background-color: #f8f9fa !important;
  color: #6c757d !important;
  border-color: #ced4da !important;
  cursor: not-allowed !important;
}

.main-plan-input,
.item-plan-input {
  background-color: #f8f9fa !important;
  color: #6c757d !important;
  border-color: #ced4da !important;
}

input:read-only,
select:disabled {
  background-color: #f8f9fa !important;
  color: #6c757d !important;
  border-color: #ced4da !important;
  cursor: not-allowed !important;
}

/* ============================================ */
/* UTILITY CLASSES                             */
/* ============================================ */

.input-with-value {
  background: #fff !important;
  color: #333 !important;
}
.input-empty {
  background: #f8f9fa !important;
  color: transparent !important;
}
.edit-mode .project-select-container select {
  background-color: #f8f9fa !important;
  color: #666 !important;
  border-color: #d0d0d0 !important;
  cursor: not-allowed !important;
}

/* ============================================ */
/* TOASTS                                      */
/* ============================================ */

.obl-error-toast {
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

.obl-error-toast.warning {
  background: #ffc107;
  color: #856404;
}

.obl-section-container-invoicing.critical-invoicing {
  border-left: 4px solid #dc3545;
}

.obl-section-container-invoicing.warning-invoicing {
  border-left: 4px solid #ffc107;
}

.obl-section-container-invoicing.normal-invoicing {
  border-left: 4px solid #28a745;
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

.toast-close {
  background: none;
  border: none;
  color: inherit;
  font-size: 18px;
  cursor: pointer;
  margin-left: 10px;
  float: right;
}

/* ============================================ */
/* PAGINATION                                  */
/* ============================================ */

.pagination-container {
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

.pagination-left {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}

.pagination-right {
  display: flex;
  align-items: center;
  gap: 20px;
  flex-wrap: wrap;
}

.page-size-selector {
  display: flex;
  align-items: center;
  gap: 8px;
}

.page-size-select {
  padding: 5px 10px;
  border: 1px solid #ced4da;
  border-radius: 4px;
  background-color: white;
  font-size: 14px;
  cursor: pointer;
}

.page-size-select:focus {
  outline: none;
  border-color: #007bff;
  box-shadow: 0 0 0 2px rgba(0, 123, 255, 0.25);
}

.pagination-btn {
  padding: 6px 12px;
  background-color: #007bff;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  transition: background-color 0.2s;
}

.pagination-btn:hover:not(:disabled) {
  background-color: #0056b3;
}

.pagination-btn:disabled {
  background-color: #6c757d;
  cursor: not-allowed;
  opacity: 0.6;
}

.pagination-info {
  font-size: 14px;
  color: #495057;
  white-space: nowrap;
}

.pagination-pages {
  display: flex;
  gap: 5px;
}

.page-number {
  padding: 6px 12px;
  background-color: white;
  border: 1px solid #dee2e6;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  transition: all 0.2s;
}

.page-number:hover:not(.active) {
  background-color: #f8f9fa;
  border-color: #007bff;
}

.page-number.active {
  background-color: #007bff;
  color: white;
  border-color: #007bff;
}

/* ============================================ */
/* SORT HEADERS                                */
/* ============================================ */

.sort-header-container {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
}

.sort-buttons-container {
  display: flex;
  flex-direction: column;
  gap: 1px;
}

.sort-btn {
  background: none;
  border: none;
  cursor: pointer;
  font-size: 8px;
  padding: 1px 2px;
  border-radius: 2px;
  color: #ccc;
  transition: all 0.2s ease;
}

.sort-btn:hover {
  background-color: rgba(255, 255, 255, 0.2);
  color: white;
}

.sort-btn.active {
  color: #ffeb3b;
}

.sort-btn.sort-asc.active {
  color: #4caf50;
}

.sort-btn.sort-desc.active {
  color: #f44336;
}

.header-label {
  margin-right: 4px;
}

.month-column-header,
.project-info-cell {
  padding: 8px 4px !important;
}

/* ============================================ */
/* SIMPLE MONTH BOX - LONG-TERM FORECAST        */
/* ============================================ */

.simple-month-inv {
  background: #fff;
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, .1);
  padding: 15px;
  position: relative;
}

/* ============================================ */
/* MAIN TABLE - MONTH COLUMNS                   */
/* ============================================ */

.month-columns-table-inv {
  border-collapse: separate;
  border-spacing: 0;
  width: 100%;
  min-width: 100%;
}

.month-columns-table-inv th,
.month-columns-table-inv td {
  border: 1px solid #68e0f0 !important;
  padding: 8px 5px !important;
  text-align: center;
  white-space: normal !important;
  word-wrap: break-word !important;
  word-break: break-word !important;
  overflow-wrap: break-word !important;
}

.month-columns-table-inv thead th {
  background: #68e0f0 !important;
  font-weight: bold;
  border-bottom: 2px solid #007bff !important;
}

.month-column-header {
  background: #0078d4 !important;
  color: white !important;
  font-weight: 600;
  font-size: 11px;
  white-space: nowrap !important;
}

.month-subheader {
  background: #e3f2fd !important;
  color: #0078d4 !important;
  font-weight: 600;
  font-size: 10px;
  white-space: nowrap !important;
}

/* ============================================ */
/* ZEBRA STRIPING & ROW HOVER                   */
/* ============================================ */

.month-columns-table-inv tbody tr:nth-child(even) {
  background-color: #f8f9fa !important;
}

.month-columns-table-inv tbody tr:hover {
  background-color: #e3f2fd !important;
}

/* ============================================ */
/* CURRENT MONTH HIGHLIGHTING                   */
/* ============================================ */

.current-month-header {
  background: #2196f3 !important;
  color: white !important;
}

.current-month-header-table {
  background: #ff9800 !important;
  color: white !important;
}

.month-subheader.current-month-header-table {
  background: #ff9800 !important;
  color: white !important;
}

.month-columns-table-inv thead tr:first-child th.month-scroll-col.current-month-header {
  background: #2196f3 !important;
  color: white !important;
}

/* ============================================ */
/* TOTALS HEADER                                */
/* ============================================ */

.totals-header-table {
  margin-bottom: 10px;
  background-color: #f8f9fa;
  border-collapse: separate;
  border-spacing: 0;
  table-layout: fixed !important;
  width: 100% !important;
}

.totals-header-table th,
.totals-header-table td {
  padding: 6px 4px !important;
  font-size: 12px;
  font-weight: bold;
  text-align: center;
  vertical-align: middle;
  border: 1px solid #dee2e6;
}

.totals-header-table .all-projects-total td {
  background-color: #d1ecf1 !important;
  color: #0c5460;
}

.totals-header-table .filtered-total td {
  background-color: #d4edda !important;
  color: #155724;
}

.totals-header-table .total-value {
  font-weight: bold;
  font-size: 13px;
}

tr.all-projects-total td {
  background-color: #d1ecf1 !important;
}

tr.filtered-total td {
  background-color: #d4edda !important;
}

/* ============================================ */
/* STICKY COLUMNS                               */
/* ============================================ */

.month-columns-table-inv .project-ref-col,
.month-columns-table-inv .project-col,
.month-columns-table-inv .pm-col,
.month-columns-table-inv .gh-col,
.month-columns-table-inv .customer-col,
.month-columns-table-inv .inv-col {
  position: sticky !important;
  z-index: 2 !important;
  background: white !important;
}

.month-columns-table-inv thead .project-ref-col,
.month-columns-table-inv thead .project-col,
.month-columns-table-inv thead .pm-col,
.month-columns-table-inv thead .gh-col,
.month-columns-table-inv thead .customer-col,
.month-columns-table-inv thead .inv-col {
  z-index: 5 !important;
  background: #e9ecef !important;
}

.project-ref-col {
  left: 0 !important;
  min-width: 120px !important;
  max-width: 120px !important;
  width: 120px !important;
  text-align: left !important;
  padding: 4px 6px !important;
  font-size: 11px !important;
}

.project-col {
  left: 120px !important;
  min-width: 180px !important;
  max-width: 180px !important;
  width: 180px !important;
  text-align: left !important;
  padding: 4px 6px !important;
  font-size: 11px !important;
}

.pm-col {
  left: 300px !important;
  min-width: 100px !important;
  max-width: 100px !important;
  width: 100px !important;
  text-align: left !important;
  padding: 4px 6px !important;
  font-size: 11px !important;
}

.gh-col {
  left: 400px !important;
  min-width: 100px !important;
  max-width: 100px !important;
  width: 100px !important;
  text-align: left !important;
  padding: 4px 6px !important;
  font-size: 11px !important;
}

.customer-col {
  left: 500px !important;
  min-width: 130px !important;
  max-width: 130px !important;
  width: 130px !important;
  text-align: left !important;
  padding: 4px 6px !important;
  font-size: 11px !important;
}

.inv-col {
  left: 630px !important;
  min-width: 90px !important;
  max-width: 90px !important;
  width: 90px !important;
  text-align: right !important;
  padding: 4px 6px !important;
  font-size: 11px !important;
}

.month-columns-table-inv .month-scroll-col {
  min-width: 60px !important;
  width: 60px !important;
}

/* ============================================ */
/* TOTALS HEADER - STICKY COLUMNS               */
/* ============================================ */

.totals-header-table .project-ref-col {
  left: 0 !important;
  min-width: 120px !important;
  max-width: 120px !important;
  width: 120px !important;
  position: sticky !important;
  z-index: 5 !important;
  background: #e9ecef !important;
}

.totals-header-table .project-col {
  left: 120px !important;
  min-width: 180px !important;
  max-width: 180px !important;
  width: 180px !important;
  position: sticky !important;
  z-index: 5 !important;
  background: #e9ecef !important;
}

.totals-header-table .pm-col {
  left: 300px !important;
  min-width: 100px !important;
  max-width: 100px !important;
  width: 100px !important;
  position: sticky !important;
  z-index: 5 !important;
  background: #e9ecef !important;
}

.totals-header-table .gh-col {
  left: 400px !important;
  min-width: 100px !important;
  max-width: 100px !important;
  width: 100px !important;
  position: sticky !important;
  z-index: 5 !important;
  background: #e9ecef !important;
}

.totals-header-table .customer-col {
  left: 500px !important;
  min-width: 130px !important;
  max-width: 130px !important;
  width: 130px !important;
  position: sticky !important;
  z-index: 5 !important;
  background: #e9ecef !important;
}

.totals-header-table .inv-col {
  left: 630px !important;
  min-width: 90px !important;
  max-width: 90px !important;
  width: 90px !important;
  position: sticky !important;
  z-index: 5 !important;
  background: #e9ecef !important;
}

.totals-header-table .month-scroll-col {
  min-width: 60px !important;
  width: 60px !important;
}

/* ============================================ */
/* SCROLLBAR                                    */
/* ============================================ */

.table-container-inv::-webkit-scrollbar {
  width: 8px;
  height: 8px;
}

.table-container-inv::-webkit-scrollbar-track {
  background: #f1f1f1;
}

.table-container-inv::-webkit-scrollbar-thumb {
  background: #888;
  border-radius: 4px;
}

.table-container-inv::-webkit-scrollbar-thumb:hover {
  background: #555;
}

/* ============================================ */
/* IMPORT MODAL                                 */
/* ============================================ */

.import-modal {
  display: none;
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: rgba(0,0,0,0.5);
  z-index: 10000;
  align-items: center;
  justify-content: center;
}

.import-modal-content {
  background: white;
  padding: 30px;
  border-radius: 10px;
  width: 90%;
  max-width: 600px;
  max-height: 80vh;
  overflow-y: auto;
  box-shadow: 0 4px 20px rgba(0,0,0,0.15);
}

.import-modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 25px;
  border-bottom: 2px solid #e9ecef;
  padding-bottom: 15px;
}

.import-modal-header h3 {
  margin: 0;
  color: #0078d4;
  font-size: 20px;
  font-weight: 600;
}

.close-modal {
  background: none;
  border: none;
  font-size: 28px;
  cursor: pointer;
  color: #666;
  padding: 0;
  width: 35px;
  height: 35px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
  transition: all 0.2s ease;
}

.close-modal:hover {
  background-color: #f3f2f1;
  color: #333;
}

.import-section {
  margin-bottom: 25px;
  padding: 20px;
  background: #f8f9fa;
  border-radius: 8px;
  border: 1px solid #e9ecef;
}

.import-section h4 {
  margin: 0 0 15px 0;
  color: #495057;
  font-size: 16px;
  font-weight: 600;
}

.file-upload-area {
  border: 2px dashed #0078d4;
  border-radius: 8px;
  padding: 40px 20px;
  text-align: center;
  background: #f8f9fa;
  cursor: pointer;
  transition: all 0.3s ease;
  margin-bottom: 15px;
}

.file-upload-area:hover {
  background: #e3f2fd;
  border-color: #0056b3;
}

.file-upload-area.dragover {
  background: #e3f2fd;
  border-color: #0056b3;
}

.upload-icon {
  font-size: 48px;
  color: #0078d4;
  margin-bottom: 15px;
}

.file-input {
  display: none;
}

.file-info {
  margin-top: 15px;
  padding: 10px;
  background: #fff;
  border-radius: 4px;
  font-size: 14px;
  color: #495057;
  border: 1px solid #dee2e6;
}

.import-preview {
  margin: 20px 0;
}

.preview-table-container {
  max-height: 200px;
  overflow-y: auto;
  border: 1px solid #dee2e6;
  border-radius: 4px;
}

.preview-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}

.preview-table th,
.preview-table td {
  border: 1px solid #dee2e6;
  padding: 8px 10px;
  text-align: left;
}

.preview-table th {
  background: #e9ecef;
  font-weight: 600;
  color: #495057;
  position: sticky;
  top: 0;
}

.preview-table tbody tr:nth-child(even) {
  background: #f8f9fa;
}
.preview-table tbody tr:hover {
  background: #e3f2fd;
}

.import-actions {
  display: flex;
  gap: 12px;
  justify-content: flex-end;
  margin-top: 25px;
  padding-top: 20px;
  border-top: 1px solid #e9ecef;
}

.import-status {
  margin-top: 20px;
  padding: 15px;
  border-radius: 6px;
  display: none;
  font-size: 14px;
  line-height: 1.5;
}

.import-status.success {
  background: #d4edda;
  color: #155724;
  border: 1px solid #c3e6cb;
  display: block;
}

.import-status.error {
  background: #f8d7da;
  color: #721c24;
  border: 1px solid #f5c6cb;
  display: block;
}

/* ============================================ */
/* RESPONSIVE                                   */
/* ============================================ */

@media(max-width: 1400px) {
  .table-actions-row {
    grid-template-columns: repeat(3, 1fr);
  }
  .historical-container,
  .first-row-container {
    grid-template-columns: repeat(2, 1fr);
  }
  .remaining-row-inv {
    grid-template-columns: repeat(3, 1fr);
  }
  .obl-main-section {
    grid-template-columns: 1fr;
  }

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

  .totals-header-table .project-ref-col {
    left: 0 !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
  }
  .totals-header-table .project-col {
    left: 100px !important;
    min-width: 150px !important;
    max-width: 150px !important;
    width: 150px !important;
  }
  .totals-header-table .pm-col {
    left: 250px !important;
    min-width: 80px !important;
    max-width: 80px !important;
    width: 80px !important;
  }
  .totals-header-table .gh-col {
    left: 330px !important;
    min-width: 80px !important;
    max-width: 80px !important;
    width: 80px !important;
  }
  .totals-header-table .customer-col {
    left: 410px !important;
    min-width: 110px !important;
    max-width: 110px !important;
    width: 110px !important;
  }
  .totals-header-table .inv-col {
    left: 520px !important;
    min-width: 70px !important;
    max-width: 70px !important;
    width: 70px !important;
  }
}

@media(max-width: 1200px) {
  .detailed-months-row {
    gap: 16px;
  }

  .current-month-container {
    flex-basis: 43%;
    width: 43%;
    max-width: 43%;
  }

  .next-months-container {
    width: 57%;
    max-width: 57%;
    gap: 12px;
  }

  .detailed-month-box {
    padding: 14px;
  }

  .current-month-container .item-row-5col,
  .current-month-container .items-header-5col {
    gap: 6px;
  }

  .next-month-box .item-row-3col,
  .next-month-box .items-header-3col {
    gap: 6px;
  }
}

@media(max-width: 1024px) {
  .table-actions-row {
    grid-template-columns: repeat(2, 1fr);
  }
  .historical-container,
  .first-row-container {
    grid-template-columns: 1fr;
  }
  .remaining-row-inv {
    grid-template-columns: repeat(2, 1fr);
  }
}

@media(max-width: 992px) {
  .detailed-months-row {
    flex-direction: column;
  }

  .current-month-container,
  .next-months-container {
    width: 100%;
    max-width: 100%;
    flex-basis: 100%;
  }

  .next-months-container {
    display: flex;
  }
}

@media(max-width: 768px) {
  .table-actions-row {
    grid-template-columns: 1fr;
  }

  .table-actions {
    padding: 10px;
  }

  .tableColp4 {
    min-width: 100%;
  }

  .historical-container,
  .first-row-container,
  .remaining-row-inv {
    grid-template-columns: 1fr;
  }

  .forecast-section {
    padding: 15px;
  }

  .project-details-grid {
    grid-template-columns: 1fr;
  }

  .pagination-container {
    flex-direction: column;
    gap: 10px;
  }

  .import-modal-content {
    padding: 20px;
    margin: 20px;
  }

  .import-actions {
    flex-direction: column;
  }

  .import-actions button {
    width: 100%;
  }

  .next-months-container {
    flex-direction: column;
  }

  .next-month-box {
    width: 100%;
  }

  .current-month-container .enhanced-inputs-container-5col {
    grid-template-columns: repeat(2, 1fr) !important;
  }

  .current-month-container .items-header-5col,
  .current-month-container .item-row-5col {
    grid-template-columns: 1.5fr 0.8fr 0.8fr 0.8fr 1fr !important;
    gap: 4px;
    font-size: 10px;
  }

  .next-month-box .items-header-3col,
  .next-month-box .item-row-3col {
    grid-template-columns: 1.5fr 0.8fr 0.8fr !important;
    gap: 4px;
    font-size: 10px;
  }

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

  .month-columns-table-inv .month-scroll-col {
    min-width: 50px;
    width: 50px;
  }

  .table-container-inv {
    overflow-x: auto;
  }
}

@media(max-width: 650px) {
  .next-months-container {
    flex-direction: column;
  }

  .next-month-box {
    width: 100%;
  }

  .current-month-container .enhanced-inputs-container-5col {
    grid-template-columns: repeat(2, 1fr) !important;
  }
}
  `,
    script: () => {
        setTimeout(async () => {
            const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/DigitalTransformationTest";
            const listName = "InvoicingPlanning";
            const projectListName = "Project List";
            // DOM Elements
            const tableBody = document.querySelector(".invoicing-table-body-inv");
            const monthColumnsHeader = document.querySelector("#monthColumnsHeader");
            const formSection = document.querySelector(".form-section");
            const tableSection = document.querySelector(".table-section");
            const projectSelect = document.querySelector("#projectSelect");
            const historicalContainer = document.querySelector("#historicalContainer");
            const historicalToggleBtn = document.querySelector("#historicalToggleBtn");
            const firstRowContainer = document.querySelector("#firstRowContainer");
            const remainingContainer = document.querySelector("#remainingContainer");
            const yearlyContainer = document.querySelector("#yearlyContainer");
            const backBtn = document.querySelector(".back-to-table");
            const form = document.querySelector("#invoicingForm");
            const exportBtn = document.querySelector(".export-btn");
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
            const clearFiltersBtn = document.querySelector(".clear-filters-btn");
            const exportDetailsBtn = document.querySelector(".export-details-btn");
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
            let pageSize = 10;
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
                    const newSize = parseInt(e.target.value);
                    changePageSize(newSize);
                });
            }
            // ========== UTILITY FUNCTIONS ==========
            const toast = (icon, title) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
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
                    input.className = input.className.replace('input-with-value', 'input-empty');
                }
                else {
                    input.value = value.toString();
                    input.className = input.className.replace('input-empty', 'input-with-value');
                }
            };
            // Enhanced OBL validation with toast notifications
            const showOBLErrorToast = (message, isWarning = false) => {
                var _a;
                const existingToast = document.querySelector('.obl-error-toast');
                if (existingToast)
                    existingToast.remove();
                const toastEl = document.createElement('div');
                toastEl.className = `obl-error-toast ${isWarning ? 'warning' : ''}`;
                toastEl.innerHTML = `
          ${message}
          <button class="toast-close">&times;</button>
        `;
                document.body.appendChild(toastEl);
                setTimeout(() => {
                    if (toastEl.parentNode) {
                        toastEl.remove();
                    }
                }, 5000);
                (_a = toastEl.querySelector('.toast-close')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
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
                    // ✅ Use getAssignedProjects to filter by permissions
                    const assignedProjects = getAssignedProjects(projects);
                    console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);
                    // ✅ Populate project select dropdown with assigned projects ONLY
                    projectSelect.innerHTML = `<option value="">Select Project</option>` +
                        assignedProjects.map(p => `<option value="${p.Id}">${p.Title}</option>`).join("");
                    // ✅ Also update the filter dropdown with assigned projects
                    filterProject.innerHTML = '<option value="">All Projects</option>' +
                        [...new Set(assignedProjects.map(p => p.Title))]
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
                            "Group Head": (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
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
                                (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
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
                                col === 12 || col === 13 || col === 14 || // Current month Previous, Current, Forecast
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
                    pageBtn.className = `page-number ${i === currentPage ? 'active' : ''}`;
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
            const renderTotalsHeader = (displayMonths, currentMonthKey) => {
                const totalsHeader = document.querySelector("#totalsHeader");
                if (!totalsHeader)
                    return;
                totalsHeader.innerHTML = '';
                const allData = forecasts;
                const filtered = filteredData.length > 0 ? filteredData : forecasts;
                const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
                const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);
                const buildTotalsRow = (label, totals, css) => `
          <tr class="${css}">
            <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;"></td>
            <td class="project-col" style="min-width:180px; max-width:180px; width:180px; text-align:left; font-weight:bold; background-color:${css === 'all-projects-total' ? '#d1ecf1' : '#d4edda'};">
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
                totalsHeader.innerHTML = `
          <table class="totals-header-table" style="width: 100%;">
            ${buildTotalsRow(`All (${allData.length})`, allTotals, 'all-projects-total')}
            ${buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total')}
          </table>
        `;
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
                // Build main table header - Column order: Project Reference → Project → Project Manager → Group Head → Customer → Invoicing → Month columns
                let mainHeaderHTML = `
          <tr>
            <!-- Project Reference - FIRST - Now clickable like project -->
            <th rowspan="2" class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">
              <div class="sort-header-container">
                <span class="header-label">Project Reference</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'projectReference' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="projectReference" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'projectReference' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="projectReference" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Project Column with Sort -->
            <th rowspan="2" class="project-col" style="min-width:180px; max-width:180px; width:180px;">
              <div class="sort-header-container">
                <span class="header-label">Project</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="project" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="project" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Project Manager -->
            <th rowspan="2" class="pm-col" style="min-width:100px; max-width:100px; width:100px;">
              <div class="sort-header-container">
                <span class="header-label">Project Manager</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="projectManager" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="projectManager" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Group Head -->
            <th rowspan="2" class="gh-col" style="min-width:100px; max-width:100px; width:100px;">
              <div class="sort-header-container">
                <span class="header-label">Group Head</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="groupHead" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="groupHead" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Customer -->
            <th rowspan="2" class="customer-col" style="min-width:130px; max-width:130px; width:130px;">
              <div class="sort-header-container">
                <span class="header-label">Customer</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="customer" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="customer" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Invoicing Column -->
            <th rowspan="2" class="inv-col" style="min-width:90px; max-width:90px; width:90px;">
              <div class="sort-header-container">
                <span class="header-label">Invoicing</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="obl" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="obl" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            ${displayMonths.map((month, monthIndex) => {
                    const isCurrentMonth = monthIndex === currentMonthIndex;
                    const monthDate = new Date(month + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const monthHeaderClass = isCurrentMonth ? 'current-month-header' : 'other-month-header';
                    const monthLabelDisplay = isCurrentMonth ? `${monthLabel} (Current Month)` : monthLabel;
                    return `
                <th colspan="2" class="month-scroll-col month-column-header ${monthHeaderClass}" style="min-width:60px;">
                  <div class="sort-header-container">
                    <span class="header-label">${monthLabelDisplay}</span>
                    <div class="sort-buttons-container">
                      <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                              data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort by Month Total Ascending">▲</button>
                      <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
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
                    const subHeaderClass = isCurrentMonth ? 'current-month-subheader' : 'other-month-subheader';
                    return `
                <th class="month-scroll-col month-subheader ${subHeaderClass}" style="min-width:60px;">
                  <div class="sort-header-container">
                    <span>Previous</span>
                    <div class="sort-buttons-container">
                      <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Ascending">▲</button>
                      <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Descending">▼</button>
                    </div>
                  </div>
                </th>
                <th class="month-scroll-col month-subheader ${subHeaderClass}" style="min-width:60px;">
                  <div class="sort-header-container">
                    <span>Current</span>
                    <div class="sort-buttons-container">
                      <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Ascending">▲</button>
                      <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
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
                    var _a, _b;
                    const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                    const proj = projects.find(p => p.Title === projectTitle);
                    // Parse forecast data
                    const forecastData = safeJsonParse(forecast.ForecastData, {});
                    const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-IN');
                    // Build project name link - clickable to edit
                    const projectNameLink = `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectTitle}</a>`;
                    const projectRefLink = `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${(proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-"}</a>`;
                    // Start building the row - Column order: Project Reference → Project → PM → GH → Customer → Invoicing
                    let rowHTML = `
            <tr data-forecast-id="${forecast.Id}">
              <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
              <td class="project-col" style="min-width:180px; max-width:180px; width:180px;">${projectNameLink}</td>
              <td class="pm-col" style="min-width:100px; max-width:100px; width:100px;">${proj ? (typeof proj.ProjectManager === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : proj.ProjectManager) || "-" : "-"}</td>
              <td class="gh-col" style="min-width:100px; max-width:100px; width:100px;">${(proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-"}</td>
              <td class="customer-col" style="min-width:130px; max-width:130px; width:130px;">${(proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-"}</td>
              <td class="inv-col" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
          `;
                    // Add month columns data (only Previous and Current)
                    displayMonths.forEach((monthKey) => {
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                        const planValue = monthData.plan || 0;
                        const currentValue = monthData.actual || 0;
                        rowHTML += `
              <td class="month-scroll-col" style="min-width:60px;">${planValue}</td>
              <td class="month-scroll-col" style="min-width:60px;">${currentValue}</td>
            `;
                    });
                    rowHTML += `</tr>`;
                    tableBodyHTML += rowHTML;
                });
                tableBody.innerHTML = tableBodyHTML;
                // Attach event listeners
                attachSortButtonListeners();
                attachEditTriggerListeners();
            };
            const attachSortButtonListeners = () => {
                document.querySelectorAll('.sort-btn[data-sort]').forEach(btn => {
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
                document.querySelectorAll('.edit-trigger').forEach(link => {
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
                            valueA = (projGA === null || projGA === void 0 ? void 0 : projGA.GroupHead) || "";
                            valueB = (projGB === null || projGB === void 0 ? void 0 : projGB.GroupHead) || "";
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
                formSection.classList.remove('edit-mode');
                projectSelect.disabled = false;
                // Reset any form state
                resetForm();
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
                historicalContainer.style.display = "none";
                if (historicalToggleBtn)
                    historicalToggleBtn.textContent = "Show";
                firstRowContainer.innerHTML = "";
                remainingContainer.innerHTML = "";
                yearlyContainer.innerHTML = "";
                // Hide all forecast sections
                toggleForecastSections(false);
                const oblContainer = document.querySelector('.obl-section-container-invoicing');
                if (oblContainer) {
                    oblContainer.className = 'obl-section-container-invoicing hidden-section';
                }
                oblLastMonthInput.removeAttribute('readonly');
                oblLastMonthInput.placeholder = "Enter OBL";
                setNumericValue(oblLastMonthInput, 0);
                const fieldNote = document.querySelector('.field-note');
                if (fieldNote)
                    fieldNote.textContent = "For new projects: Enter initial OBL value";
                oblValidationMessage.textContent = "✅ All values are within OBL limit";
                oblValidationMessage.className = 'validation-message-invoicing';
                updateOBLValidationDisplay();
            };
            // ========== VALIDATION FUNCTIONS ==========
            const calculateTotalCurrentValue = () => {
                let total = 0;
                // Include monthly actual values
                const editableMonths = allMonths.slice(3);
                editableMonths.forEach(monthKey => {
                    const currentInputs = document.querySelectorAll(`input.main-current-input[data-month="${monthKey}"]:not([readonly])`);
                    currentInputs.forEach(input => {
                        total += getNumericValue(input);
                    });
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
                        blockAllCurrentInputs();
                        return false;
                    }
                    else {
                        const existingToast = document.querySelector('.obl-error-toast');
                        if (existingToast)
                            existingToast.remove();
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
                    showOBLErrorToast(`❌ Total current value exceeds OBL by ${exceedAmount.toLocaleString()}! Reduce values to continue.`);
                    oblValidationMessage.classList.add('danger-invoicing');
                    blockIncreasingCurrentInputs(changedInput);
                    return false;
                }
                else if (percentage > 90) {
                    showOBLErrorToast('⚠️ OBL limit nearly reached! Consider reducing values.', true);
                    oblValidationMessage.classList.add('warning-invoicing');
                    unblockAllCurrentInputs();
                    return true;
                }
                else {
                    const existingToast = document.querySelector('.obl-error-toast');
                    if (existingToast)
                        existingToast.remove();
                    oblValidationMessage.textContent = '✅ All values are within OBL limit';
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
            const allowItemInputIncrease = (input, newValue) => {
                const oldValue = parseInt(input.getAttribute('data-old-value') || '0');
                const isIncreasing = newValue > oldValue;
                if (!isIncreasing) {
                    return true;
                }
                const monthKey = input.dataset.month;
                const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                if (!mainCurrentInput)
                    return true;
                const mainCurrentValue = getNumericValue(mainCurrentInput);
                const itemCurrentInputs = document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]`);
                let currentItemsTotal = 0;
                itemCurrentInputs.forEach(itemInput => {
                    if (itemInput === input) {
                        currentItemsTotal += newValue;
                    }
                    else {
                        currentItemsTotal += getNumericValue(itemInput);
                    }
                });
                return currentItemsTotal <= mainCurrentValue;
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
                if (!mainCurrentInput)
                    return true;
                const mainCurrentValue = getNumericValue(mainCurrentInput);
                let itemsTotal = 0;
                itemCurrentInputs.forEach(input => {
                    itemsTotal += getNumericValue(input);
                });
                let validationDisplay = document.querySelector(`.month-total-validation[data-month="${monthKey}"]`);
                if (!validationDisplay) {
                    validationDisplay = document.createElement('div');
                    validationDisplay.className = 'month-total-validation';
                    validationDisplay.setAttribute('data-month', monthKey);
                    const monthBlock = mainCurrentInput.closest('.month-block');
                    if (monthBlock) {
                        const existingValidation = monthBlock.querySelector('.month-total-validation');
                        if (existingValidation) {
                            existingValidation.remove();
                        }
                        monthBlock.appendChild(validationDisplay);
                    }
                }
                if (itemsTotal > mainCurrentValue) {
                    validationDisplay.textContent = `❌ Items total (${itemsTotal}) exceeds Month current (${mainCurrentValue})! Reduce item values to continue.`;
                    validationDisplay.className = 'month-total-validation month-total-invalid';
                    blockIncreasingMonthItemInputs(monthKey, changedInput);
                    return false;
                }
                else if (itemsTotal === mainCurrentValue && itemsTotal > 0) {
                    validationDisplay.textContent = `✅ Items total matches Month current: ${itemsTotal}`;
                    validationDisplay.className = 'month-total-validation month-total-valid';
                    unblockAllMonthItemInputs(monthKey);
                    return true;
                }
                else {
                    validationDisplay.textContent = `📊 Items total: ${itemsTotal} / ${mainCurrentValue}`;
                    validationDisplay.className = 'month-total-validation month-total-valid';
                    unblockAllMonthItemInputs(monthKey);
                    return true;
                }
            };
            const blockIncreasingMonthItemInputs = (monthKey, allowedInput) => {
                const itemInputs = document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]`);
                itemInputs.forEach(input => {
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
                        oblContainer.className = 'obl-section-container-invoicing visible-section new-project';
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
                        oblContainer.className = 'obl-section-container-invoicing visible-section existing-project';
                    }
                    oblLastMonthInput.setAttribute('readonly', 'true');
                    const oblValue = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                    setNumericValue(oblLastMonthInput, oblValue);
                    oblLastMonthInput.placeholder = "OBL from database";
                    if (fieldNote)
                        fieldNote.textContent = "OBL value from database record";
                    if (oblValue > 0) {
                        oblLastMonthInput.value = oblValue.toString();
                        oblLastMonthInput.className = oblLastMonthInput.className.replace('input-empty', 'input-with-value');
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
                if (groupHeadDisplay)
                    groupHeadDisplay.textContent = project.GroupHead || "-";
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
                updateOBLValidationDisplay();
                attachForecastEventListeners();
            };
            const protectHistoricalElements = () => {
                // Remove any existing event listeners from historical elements by cloning and replacing
                document.querySelectorAll('.historical-month input, .historical-month select, .historical-month button').forEach(element => {
                    var _a;
                    const newElement = element.cloneNode(true);
                    (_a = element.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newElement, element);
                });
                // Add global event prevention for historical sections
                document.querySelectorAll('.historical-month').forEach(section => {
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
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
                const items = data.items || createEmptyItems();
                return `
          <div class="month-box-inv historical-month">
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

            <div class="items-section historical-items-section">
              <div class="items-header historical-items-header">
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
                    const itemPlanClass = item.plan > 0 ? 'input-with-value' : 'input-empty';
                    const itemActualClass = item.actual > 0 ? 'input-with-value' : 'input-empty';
                    const itemNameClass = item.name ? 'input-with-value' : 'input-empty';
                    return `
                  <div class="item-row historical-item-row">
                    <input type="text" placeholder="Item ${itemNum}" value="${item.name}" data-item="${itemKey}" data-month="${monthKey}" readonly 
                          class="${itemNameClass}" />
                    <input type="text" value="${itemPlanValue}" data-item-type="plan" data-item="${itemKey}" data-month="${monthKey}" 
                          readonly class="item-plan-input ${itemPlanClass} number-input" />
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
                html += `<div class="detailed-months-row">`;
                // Current Month - 44% - 5 columns (Item Description, Previous, Current, Forecast, Date)
                const currentMonth = months[0];
                if (currentMonth) {
                    const monthDate = new Date(currentMonth + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[currentMonth] || { plan: 0, actual: 0, items: {} };
                    html += `<div class="current-month-container">`;
                    html += renderDetailedMonthBox5Col(currentMonth, monthLabel, data, true);
                    html += `</div>`;
                }
                // Next 2 Months - 56% (split equally) - 3 columns (Item Description, Previous, Current)
                html += `<div class="next-months-container">`;
                for (let i = 1; i <= 2; i++) {
                    if (i < months.length) {
                        const monthKey = months[i];
                        const monthDate = new Date(monthKey + "-01");
                        const monthLabel = formatMonthLabel(monthDate);
                        const data = existingData[monthKey] || { plan: 0, actual: 0, items: {} };
                        html += `<div class="next-month-box">`;
                        html += renderDetailedMonthBox3Col(monthKey, monthLabel, data, false);
                        html += `</div>`;
                    }
                }
                html += `</div>`;
                html += `</div>`;
                firstRowContainer.innerHTML = html;
            };
            // RENDER: Current Month - 5 columns (Item Description, Previous, Current, Forecast, Date)
            const renderDetailedMonthBox5Col = (monthKey, monthLabel, data, isCurrentMonth = false) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const displayForecastValue = data.forecast && data.forecast > 0 ? data.forecast.toString() : '';
                const displayDateValue = data.date || '';
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
                const forecastInputClass = data.forecast && data.forecast > 0 ? 'input-with-value' : 'input-empty';
                const items = data.items || createEmptyItems();
                // Calculate items total for the summary bar
                let itemsTotalSum = 0;
                Array.from({ length: 10 }, (_, i) => i + 1).forEach(itemNum => {
                    const item = items[itemNum] || { name: '', plan: 0, actual: 0 };
                    itemsTotalSum += item.actual || 0;
                });
                const mainCurrentTotal = data.actual || 0;
                const itemsTotalClass = itemsTotalSum > mainCurrentTotal ? 'month-total-invalid' : 'month-total-valid';
                const itemsTotalText = itemsTotalSum > mainCurrentTotal
                    ? `❌ Items total (${itemsTotalSum}) exceeds Month current (${mainCurrentTotal})! Reduce item values to continue.`
                    : (itemsTotalSum === mainCurrentTotal && itemsTotalSum > 0
                        ? `✅ Items total matches Month current: ${itemsTotalSum}`
                        : `📊 Items total: ${itemsTotalSum} / ${mainCurrentTotal}`);
                return `
          <div class="detailed-month-box current-month-box">
            <div class="month-header-inv">
              ${monthLabel} 
             
            </div>
            
            <div class="month-block">
              <div class="enhanced-inputs-container-5col">
                <!-- Previous Column -->
                <div class="enhanced-input-group">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input ${planInputClass} number-input" />
                </div>
                
                <!-- Current Column -->
                <div class="enhanced-input-group">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        class="main-current-input ${actualInputClass} number-input" />
                </div>
                
                <!-- Forecast Column -->
                <div class="enhanced-input-group">
                  <label>Forecast</label>
                  <input type="text" value="${displayForecastValue}" data-type="forecast" data-month="${monthKey}"
                        readonly class="main-forecast-input ${forecastInputClass} number-input" />
                </div>
                
                <!-- Date Column -->
                <div class="enhanced-input-group">
                  <label>Date</label>
                  <input type="date" value="${displayDateValue}" data-type="date" data-month="${monthKey}"
                        class="main-date-input" />
                </div>
              </div>
              <div id="validation-${monthKey}" class="month-total-validation ${itemsTotalClass}">${itemsTotalText}</div>
            </div>

            <div class="items-section">
              <div class="items-header-5col">
                <div>Item Description</div>
                <div>Previous</div>
                <div>Current</div>
                <div>Forecast</div>
                <div>Date</div>
              </div>
              
              ${Array.from({ length: 10 }, (_, i) => i + 1).map(itemNum => {
                    const item = items[itemNum] || { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
                    const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                    const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                    const itemForecastValue = item.forecast && item.forecast > 0 ? item.forecast.toString() : '';
                    const itemDateValue = item.date || '';
                    const itemPlanClass = item.plan > 0 ? 'input-with-value' : 'input-empty';
                    const itemActualClass = item.actual > 0 ? 'input-with-value' : 'input-empty';
                    const itemForecastClass = item.forecast && item.forecast > 0 ? 'input-with-value' : 'input-empty';
                    const itemNameClass = item.name ? 'input-with-value' : 'input-empty';
                    return `
                  <div class="item-row-5col">
                    <!-- Item Description -->
                    <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
                          data-item="${itemNum}" data-month="${monthKey}" 
                          class="${itemNameClass}" />
                    
                    <!-- Previous (Plan) -->
                    <input type="text" value="${itemPlanValue}" 
                          data-item-type="plan" data-item="${itemNum}" data-month="${monthKey}" 
                          readonly class="item-plan-input ${itemPlanClass} number-input" />
                    
                    <!-- Current (Actual) -->
                    <input type="text" value="${itemActualValue}" 
                          data-item-type="actual" data-item="${itemNum}" data-month="${monthKey}" 
                          class="item-current-input ${itemActualClass} number-input" />
                    
                    <!-- Forecast -->
                    <input type="text" value="${itemForecastValue}" 
                          data-item-type="forecast" data-item="${itemNum}" data-month="${monthKey}" 
                          readonly class="item-forecast-input ${itemForecastClass} number-input" />
                    
                    <!-- Date -->
                    <input type="date" value="${itemDateValue}" 
                          data-item-type="date" data-item="${itemNum}" data-month="${monthKey}" 
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
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
                const items = data.items || createEmptyItems();
                // Calculate items total for the summary bar
                let itemsTotalSum = 0;
                Array.from({ length: 10 }, (_, i) => i + 1).forEach(itemNum => {
                    const item = items[itemNum] || { name: '', plan: 0, actual: 0 };
                    itemsTotalSum += item.actual || 0;
                });
                const mainCurrentTotal = data.actual || 0;
                const itemsTotalClass = itemsTotalSum > mainCurrentTotal ? 'month-total-invalid' : 'month-total-valid';
                const itemsTotalText = itemsTotalSum > mainCurrentTotal
                    ? `❌ Items total (${itemsTotalSum}) exceeds Month current (${mainCurrentTotal})! Reduce item values to continue.`
                    : (itemsTotalSum === mainCurrentTotal && itemsTotalSum > 0
                        ? `✅ Items total matches Month current: ${itemsTotalSum}`
                        : `📊 Items total: ${itemsTotalSum} / ${mainCurrentTotal}`);
                return `
          <div class="detailed-month-box next-month-box">
            <div class="month-header-inv">
              ${monthLabel}
            </div>
            
            <div class="month-block">
              <div class="enhanced-inputs-container-3col">
                <!-- Previous Column -->
                <div class="enhanced-input-group">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input ${planInputClass} number-input" />
                </div>
                
                <!-- Current Column -->
                <div class="enhanced-input-group">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        class="main-current-input ${actualInputClass} number-input" />
                </div>
              </div>
              <div id="validation-${monthKey}" class="month-total-validation ${itemsTotalClass}">${itemsTotalText}</div>
            </div>

            <div class="items-section">
              <div class="items-header-3col">
                <div>Item Description</div>
                <div>Previous</div>
                <div>Current</div>
              </div>
              
              ${Array.from({ length: 10 }, (_, i) => i + 1).map(itemNum => {
                    const item = items[itemNum] || { name: '', plan: 0, actual: 0 };
                    const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                    const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                    const itemPlanClass = item.plan > 0 ? 'input-with-value' : 'input-empty';
                    const itemActualClass = item.actual > 0 ? 'input-with-value' : 'input-empty';
                    const itemNameClass = item.name ? 'input-with-value' : 'input-empty';
                    return `
                  <div class="item-row-3col">
                    <!-- Item Description -->
                    <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
                          data-item="${itemNum}" data-month="${monthKey}" 
                          class="${itemNameClass}" />
                    
                    <!-- Previous (Plan) -->
                    <input type="text" value="${itemPlanValue}" 
                          data-item-type="plan" data-item="${itemNum}" data-month="${monthKey}" 
                          readonly class="item-plan-input ${itemPlanClass} number-input" />
                    
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
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
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
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
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
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
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
                            const allowByMonth = allowItemInputIncrease(this, currentValue);
                            if (!allowByOBL || !allowByMonth) {
                                setNumericValue(this, oldValue);
                                updateOBLValidationDisplay(this);
                                return;
                            }
                        }
                        // Update visual state
                        if (currentValue === 0) {
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
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
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                // ===== ITEM NAME INPUTS =====
                document.querySelectorAll('input[data-item]:not([readonly])').forEach(input => {
                    if (!input.classList.contains('number-input')) {
                        input.addEventListener('input', function () {
                            if (this.value === '') {
                                this.className = this.className.replace('input-with-value', 'input-empty');
                            }
                            else {
                                this.className = this.className.replace('input-empty', 'input-with-value');
                            }
                        });
                    }
                });
                // ===== VALIDATION FOR ITEM TOTALS =====
                document.querySelectorAll('.main-current-input:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
                        const monthKey = this.dataset.month;
                        if (monthKey) {
                            validateSingleMonthItems(monthKey, this);
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
                    // If frozen OR user can't edit, set read-only
                    const isReadOnly = isFrozen || !canEdit;
                    setFormReadOnly(isReadOnly);
                    // Show/hide save button
                    const saveBtn = document.querySelector('.btn-primary[type="submit"]');
                    if (saveBtn) {
                        saveBtn.style.display = isReadOnly ? 'none' : '';
                    }
                    // Update form title
                    const formTitleEl = document.querySelector("#formTitle");
                    if (formTitleEl) {
                        if (isFrozen) {
                            formTitleEl.innerHTML = '🔒 Edit Forecast (Frozen - Read Only)';
                            formTitleEl.style.color = '#856404';
                        }
                        else if (!canEdit) {
                            formTitleEl.innerHTML = '🔒 Edit Forecast (Read Only)';
                            formTitleEl.style.color = '#856404';
                        }
                        else {
                            formTitleEl.innerHTML = 'Edit Forecast';
                            formTitleEl.style.color = '';
                        }
                    }
                }, 300);
                toast("info", `Editing forecast for ${forecast.Title}`);
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
                let hasMonthValidationError = false;
                for (let i = 3; i <= 5; i++) {
                    if (i < allMonths.length) {
                        const monthKey = allMonths[i];
                        const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                        const itemCurrentInputs = document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]`);
                        if (mainCurrentInput) {
                            const mainCurrentValue = getNumericValue(mainCurrentInput);
                            let itemsTotal = 0;
                            itemCurrentInputs.forEach(input => {
                                itemsTotal += getNumericValue(input);
                            });
                            if (itemsTotal > mainCurrentValue) {
                                hasMonthValidationError = true;
                                toast("error", `Items total exceeds month current value for ${formatMonthLabel(new Date(monthKey + "-01"))}!`);
                                break;
                            }
                        }
                    }
                }
                if (hasMonthValidationError) {
                    return;
                }
                try {
                    const forecastData = {};
                    const allStorageMonths = getAllMonthsForStorage();
                    // Process monthly data including current month items with dates
                    allStorageMonths.forEach(monthKey => {
                        const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                        const mainPlanInput = document.querySelector(`input.main-plan-input[data-month="${monthKey}"]`);
                        const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${monthKey}"]`);
                        const mainDateInput = document.querySelector(`input.main-date-input[data-month="${monthKey}"]`);
                        if (mainCurrentInput && mainPlanInput) {
                            const monthData = {
                                plan: getNumericValue(mainPlanInput),
                                actual: getNumericValue(mainCurrentInput),
                                forecast: mainForecastInput ? getNumericValue(mainForecastInput) : 0,
                                date: mainDateInput ? mainDateInput.value : '',
                                items: {}
                            };
                            // Process items for detailed months (index 3-5 in allMonths array = current + next 2 months)
                            const monthIndex = allStorageMonths.indexOf(monthKey);
                            if (monthIndex >= 3 && monthIndex <= 5) {
                                // For current month (index 3) - 5 columns with Forecast and Date
                                // For next 2 months (index 4-5) - 3 columns without Forecast and Date
                                const isCurrentMonth = monthIndex === 3;
                                for (let itemNum = 1; itemNum <= 10; itemNum++) {
                                    const itemKey = itemNum.toString();
                                    const itemNameInput = document.querySelector(`input[data-item="${itemKey}"][data-month="${monthKey}"]:not([data-item-type])`);
                                    const itemPlanInput = document.querySelector(`input.item-plan-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    const itemCurrentInput = document.querySelector(`input.item-current-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    if (itemNameInput && itemPlanInput && itemCurrentInput) {
                                        const itemData = {
                                            name: itemNameInput.value || '',
                                            plan: getNumericValue(itemPlanInput),
                                            actual: getNumericValue(itemCurrentInput),
                                            forecast: 0,
                                            date: ''
                                        };
                                        // Only get forecast and date for current month
                                        if (isCurrentMonth) {
                                            const itemForecastInput = document.querySelector(`input.item-forecast-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                            const itemDateInput = document.querySelector(`input.item-date-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                            itemData.forecast = itemForecastInput ? getNumericValue(itemForecastInput) : 0;
                                            itemData.date = itemDateInput ? itemDateInput.value : '';
                                        }
                                        monthData.items[itemKey] = itemData;
                                    }
                                }
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
                const assignedGroupHeads = getAssignedProjectValues(projects, 'GroupHead');
                const assignedReferences = getAssignedProjectValues(projects, 'ProjectReference');
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
            };
            const filterTable = () => {
                const year = filterYear.value;
                const month = filterMonth.value;
                const project = filterProject.value;
                const customer = filterCustomer.value;
                const projectManager = filterProjectManager.value;
                const groupHead = filterGroupHead.value;
                const projectReference = filterProjectReference.value;
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
                filtered = filtered.filter(forecast => {
                    var _a, _b;
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
                    if (projectManager) {
                        const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                        const managerName = typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === 'object' ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : proj === null || proj === void 0 ? void 0 : proj.ProjectManager;
                        if (managerName !== projectManager)
                            return false;
                    }
                    if (groupHead) {
                        const proj = projects.find(p => { var _a; return p.Title === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                        if ((proj === null || proj === void 0 ? void 0 : proj.GroupHead) !== groupHead)
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
                    '.main-date-input',
                    '.item-date-input'
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
                const oblInput = document.querySelector('#oblLastMonthInvoicing');
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
                    toast("error", "❌ You don't have permission to create invoicing forecasts");
                    return;
                }
                formSection.style.display = "block";
                tableSection.style.display = "none";
                // ✅ Show/hide save button based on permissions
                const saveBtn = document.querySelector('.btn-primary[type="submit"]');
                if (saveBtn) {
                    const canSave = isEditMode ? canEditProject(project) : canCreateRecord();
                    saveBtn.style.display = canSave ? '' : 'none';
                }
                if (isEditMode) {
                    formSection.classList.add('edit-mode');
                    projectSelect.disabled = true;
                    formTitle.textContent = "Edit Forecast";
                }
                else {
                    formSection.classList.remove('edit-mode');
                    projectSelect.disabled = true;
                    formTitle.textContent = "Edit Forecast";
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
                    // ✅ Initialize shared authentication FIRST
                    await authManager.initialize({ siteUrl });
                    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                    console.log(`👤 User: ${getCurrentUserFullName()}`);
                    console.log(`📋 User Groups:`, getUserGroups());
                    console.log(`✅ Has Full Control: ${hasFullControl()}`);
                    console.log(`✅ Can Contribute: ${canContribute()}`);
                    console.log(`✅ Can Freeze: ${canFreeze()}`);
                    console.log(`✅ Is GH Head: ${isGHHead()}`);
                    await loadProjects();
                    await loadForecasts();
                    if (forecasts.length > 0) {
                        await applyAutoProgression();
                    }
                    initializeFilters();
                    // ✅ Call filterTable() to apply project filtering
                    filterTable();
                    // ✅ Update UI based on permissions
                    updateUIBasedOnPermissions({
                        exportBtn: '.export-btn',
                        addBtn: '.show-form-btn',
                        saveBtn: '.btn-primary[type="submit"]',
                        freezeCheckboxes: '.freeze-checkbox'
                    });
                    console.log("✅ Invoicing application initialized successfully");
                }
                catch (error) {
                    console.error("Failed to initialize Invoicing application:", error);
                    toast("error", "Failed to initialize Invoicing application");
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
            // Historical Data show/hide toggle (default: hidden)
            if (historicalToggleBtn) {
                historicalToggleBtn.addEventListener("click", () => {
                    const isCurrentlyHidden = historicalContainer.style.display === "none";
                    if (isCurrentlyHidden) {
                        historicalContainer.style.display = "";
                        historicalToggleBtn.textContent = "Hide";
                    }
                    else {
                        historicalContainer.style.display = "none";
                        historicalToggleBtn.textContent = "Show";
                    }
                });
            }
            // Initialize the application
            initialize();
        }, 100);
    }
};
//# sourceMappingURL=invoicing.js.map
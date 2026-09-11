import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/revenue.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject as authCanEditProject, canCreateRecord, getAssignedProjects, updateUIBasedOnPermissions as authUpdateUI } from "./auth";
let currentSortConfig = {
    column: '',
    direction: 'asc'
};
export default {
    content: `
<div class="revenue-page fade-in">
  <section class="table-section card">
    <div class="page-header">
      <h1>Revenue Forecast Management</h1>
       
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
        <div class="tableColp4"> 
          <label>&nbsp;</label>
          <button class="btn-secondary export-details-btn">📋 Details Export Excel</button>
        </div>
        <div class="tableColp4"> 
          <label>&nbsp;</label>
          <button class="btn-secondary export-historical-btn">📈 Historical Export</button>
        </div>
      </div>
    </div>
       
    <div class="table-container-rev">
      <div class="table-scroll-wrapper">
        <table class="month-columns-table-rev"> 
           <thead id="totalsHeader"></thead> 
           <thead id="monthColumnsHeader"></thead> 
        <tbody class="revenue-table-body-rev"></tbody>
      </table>
      </div>
      
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
    <h1 id="formTitle">Add / Edit Forecast</h1>
  </div>
  <form id="revenueForm">
    
    <!-- Project Selection -->
    <div class="form-row hidden-section">
      <div class="form-group project-select-container">
        <label>Project</label>
        <select id="projectSelect" required></select>
      </div>
    </div>

    <!-- MAIN WRAPPER - 2 Column Layout -->
    <div class="project-main-wrapper">

      <!-- LEFT SIDE - Project Details (6 Rows) -->
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

      <!-- RIGHT SIDE - Remarks & OBL -->
      <div class="project-right-side">
        <!-- Remarks Internal -->
        <div class="remarks-group">
          <label for="remarksInternal">REMARKS (INTERNAL)</label>
          <textarea id="remarksInternal" class="remarks-textarea" rows="3" placeholder="Enter internal remarks..."></textarea>
        </div>
        
        <!-- Remarks Sprint Cell  -->
        <div class="remarks-group">
          <label for="remarksSprintCell">REMARKS (SPRINT CALL)</label>
          <textarea id="remarksSprintCell" class="remarks-textarea" rows="3" placeholder="Enter sprint cell remarks..."></textarea>
        </div>
        
        <!-- OBL Section -->
        <div class="obl-section-container hidden-section" id="oblContainer">
          <div class="obl-content">
            <div class="obl-left">
              <div class="obl-label-group">
                <label>OBL</label>
                
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
            ✅ All values are within OBL limit
          </div> 
        </div>
      </div>

    </div>

    <!-- Forecast Sections -->
<div class="forecast-section hidden-section" id="historicalForecastSection">
  <div class="historical-toggle-header">
    <h3>📊 Historical Data (Last 3 Months - Readonly)</h3>
    <button type="button" id="toggleHistoricalBtn" class="btn-secondary">Hide</button>
  </div>
  <div id="historicalContainer" class="historical-container"></div>
</div>
    
    
    <div class="forecast-section hidden-section" id="detailedForecastSection">
      <h3>Detailed Forecast (Current + Next 2 Months)</h3>
      <div id="firstRowContainer" class="first-row-container"></div>
    </div>
    
    <div class="forecast-section hidden-section" id="longtermForecastSection">
      <h3>Long-term Forecast (Next 15 Months)</h3>
      <div id="remainingContainer" class="remaining-container"></div>
    </div>

    <div class="forecast-section hidden-section" id="yearlyForecastSection">
      <h3>5-Year Forecast</h3>
      <div id="yearlyContainer" class="yearly-container-rev"></div>
    </div>
    
    <div class="form-actions">
      <button type="submit" class="btn-primary">💾 Save Revenue Plan</button>
      <button type="button" class="btn-secondary back-to-table">⬅ Back</button>
    </div>
  </form>
</section>
    
   
</div>
  `,
    style: `
.page-number.dots {
  background-color: transparent;
  border: none;
  cursor: default;
  color: #6c757d;
}

.pagination-btn:focus,
.page-number:focus,
.page-size-select:focus {
  outline: 2px solid #007bff;
  outline-offset: 2px;
}

.table-loading {
  opacity: 0.7;
  pointer-events: none;
}

@keyframes fadeIn {
  from { opacity: 0.5; }
  to { opacity: 1; }
}

.revenue-table-body-rev tr {
  animation: fadeIn 0.3s ease-in;
}

.table-container-rev {
  overflow: auto;
  position: relative;
}

.table-scroll-wrapper {
  overflow: auto;
  position: relative;
}

.month-columns-table-rev {
  border-collapse: separate;
  border-spacing: 0;
  width: 100%;
  min-width: 100%;
}

/* Fixed columns - left side */
.month-columns-table-rev .freeze-col, 
.month-columns-table-rev .project-ref-col,
.month-columns-table-rev .project-col,
.month-columns-table-rev .pm-col,
.month-columns-table-rev .gh-col,
.month-columns-table-rev .customer-col,
.month-columns-table-rev .obl-col {
  position: sticky;
  z-index: 2;
  background: white;
}

.month-columns-table-rev thead .freeze-col, 
.month-columns-table-rev thead .project-ref-col,
.month-columns-table-rev thead .project-col,
.month-columns-table-rev thead .pm-col,
.month-columns-table-rev thead .gh-col,
.month-columns-table-rev thead .customer-col,
.month-columns-table-rev thead .obl-col {
  z-index: 5;
  background: #e9ecef;
}

/* Updated column sizes */
.freeze-col { left: 0; min-width: 80px; max-width: 80px; } 
.project-ref-col { left: 160px; min-width: 100px; max-width: 100px; }
.project-col { left: 260px; min-width: 140px; max-width: 140px; }
.pm-col { left: 400px; min-width: 90px; max-width: 90px; }
.gh-col { left: 490px; min-width: 130px; max-width: 130px; }
.customer-col { left: 620px; min-width: 70px; max-width: 70px; }
.obl-col { left: 690px; min-width: 100px; max-width: 100px; }

/* Scrollable month columns */
.month-columns-table-rev .month-scroll-col {
  min-width: 120px;
}

.month-columns-table-rev tbody td,
.month-columns-table-rev thead th {
  border: 1px solid #68e0f0;
  padding: 8px 5px;
  text-align: center;
  white-space: nowrap;
}

.month-columns-table-rev thead th {
  background: #68e0f0;
  font-weight: bold;
  border-bottom: 2px solid #007bff;
}

.month-columns-table-rev tbody tr:nth-child(even) {
  background-color: #f8f9fa;
}

.month-columns-table-rev tbody tr:hover {
  background-color: #e3f2fd;
}

/* Fix freeze checkbox - ensure pointer events work */
.freeze-checkbox {
  width: 18px;
  height: 18px;
  cursor: pointer;
  pointer-events: auto !important;
  z-index: 10;
  position: relative;
}

/* Action button icon style */
.edit-btn {
  background: none;
  border: none;
  cursor: pointer;
  font-size: 18px;
  padding: 4px 6px;
  color: #007bff;
  transition: color 0.2s, transform 0.2s;
}

.edit-btn:hover {
  color: #0056b3;
  transform: scale(1.2);
}

.edit-btn:focus {
  outline: 2px solid #007bff;
  outline-offset: 2px;
}

/* Action column - center align */
 

.freeze-checkbox:disabled {
  cursor: not-allowed;
  opacity: 0.6;
  pointer-events: auto !important;
}

.table-container-rev::-webkit-scrollbar {
  width: 8px;
  height: 8px;
}

.table-container-rev::-webkit-scrollbar-track {
  background: #f1f1f1;
}

.table-container-rev::-webkit-scrollbar-thumb {
  background: #888;
  border-radius: 4px;
}

.table-container-rev::-webkit-scrollbar-thumb:hover {
  background: #555;
}

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

@media (max-width: 768px) {
  .pagination-container {
    flex-direction: column;
    align-items: stretch;
    gap: 15px;
  }
  
  .pagination-left, .pagination-right {
    justify-content: center;
    width: 100%;
  }
  
  .freeze-col { left: 0; min-width: 50px; max-width: 50px; } 
  .project-ref-col { left: 110px; min-width: 80px; max-width: 80px; }
  .project-col { left: 190px; min-width: 100px; max-width: 100px; }
  .pm-col { left: 290px; min-width: 70px; max-width: 70px; }
  .gh-col { left: 360px; min-width: 80px; max-width: 80px; }
  .customer-col { left: 440px; min-width: 60px; max-width: 60px; }
  .obl-col { left: 500px; min-width: 80px; max-width: 80px; }
}

/* Totals header - merged cells from Freeze to OBL */
.totals-header-table {
  margin-bottom: 10px;
  background-color: #f8f9fa;
  border-collapse: separate;
  border-spacing: 0;
  width: 100%;
  table-layout: fixed;
}

.totals-header-table th {
  padding: 8px 5px;
  font-size: 12px;
  font-weight: bold;
  text-align: center;
  vertical-align: middle;
  border: 1px solid #dee2e6;
  border-bottom: 2px solid #007bff;
  background-color: #e9ecef;
}

.totals-header-table .totals-label {
  font-weight: 600;
  color: #495057;
  background-color: #f1f3f4 !important;
}

.totals-header-table .all-projects-total {
  background-color: #d1ecf1 !important;
  color: #0c5460;
}

.totals-header-table .filtered-total {
  background-color: #d4edda !important;
  color: #155724;
}

.totals-header-table .total-value {
  font-weight: bold;
  font-size: 13px; 
}

tr.filtered-total {
    background-color: #00ffff;
}
tr.all-projects-total {
    background-color: #fff8dc;
}
    .role-indicator {
    display: none;
}
 /* ============================================ */
/* PROJECT MAIN WRAPPER - 2 COLUMN LAYOUT       */
/* ============================================ */

.project-main-wrapper {
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

/* LEFT SIDE - Start from left, not middle */
.project-left-side {
  padding: 8px 16px !important;
  background: #ffffff !important;
  border-right: 1px solid #e9ecef !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 0 !important;
  justify-content: flex-start !important;  /* Changed from center to flex-start */
  align-items: flex-start !important;      /* Added: align items to left */
}

 
.project-left-side .detail-item {
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

.project-left-side .detail-label {
    font-weight: 700 !important;
    white-space: nowrap !important;
}

.project-left-side .detail-value {
    white-space: nowrap !important;
}

/* RIGHT SIDE */
.project-right-side {
  padding: 8px 16px !important;
  background: #ffffff !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 6px !important;
  justify-content: flex-start !important;  /* Changed from center to flex-start */
}

/* Remarks Group - Remove the extra row for Project Name */
.remarks-group {
  display: flex !important;
  flex-direction: column !important;
  width: 100% !important;
}

.remarks-group label {
  font-size: 12x !important;
  font-weight: 600 !important;
  color: #495057 !important;
  margin-bottom: 2px !important;
  letter-spacing: 0.3px !important;
  text-align: left !important;
}

.remarks-textarea {
  width: 100% !important;
  padding: 4px 8px !important;
  border: 1px solid #ced4da !important;
  border-radius: 4px !important;
  font-size: 12px !important;
  font-family: inherit !important;
  resize: vertical !important;
  min-height: 32px !important;
  max-height: 50px !important;
  background-color: #fff !important;
  transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out !important;
  line-height: 1.4 !important;
  text-align: left !important;
}

.remarks-textarea:focus {
  border-color: #80bdff !important;
  outline: 0 !important;
  box-shadow: 0 0 0 0.15rem rgba(0, 123, 255, 0.15) !important;
}

.remarks-textarea::placeholder {
  color: #6c757d !important;
  font-style: italic !important;
  font-size: 10px !important;
}

/* OBL Section */
.obl-section-container {
  margin: 2px 0 0 0 !important;
  padding: 6px 12px !important;
  background: #f8f9fa !important;
  border-radius: 4px !important;
  border: 1px solid #e9ecef !important;
}

.obl-content {
  display: grid !important;
  grid-template-columns: 1fr 1.5fr !important;
  gap: 12px !important;
  align-items: start !important;
}

.obl-left {
  display: flex !important;
  flex-direction: row !important;
}

.obl-label-group {
  display: flex !important;
  align-items: center !important;
  gap: 4px !important;
  margin-bottom: 2px !important;
  text-align: left !important;
}

.obl-label-group label {
  font-size: 12px !important;
  font-weight: 700 !important;
  color: #495057 !important;
  margin: 0 !important;
}

.info-tooltip {
  font-size: 12px !important;
  color: #007bff !important;
  cursor: help !important;
}

.obl-value-display {
  display: flex !important;
  flex-direction: column !important;
}

.obl-input {
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

.field-note {
  font-size: 9px !important;
  color: #6c757d !important;
  margin-top: 2px !important;
  font-style: italic !important;
  text-align: left !important;
}

.obl-right {
  display: flex !important;
  align-items: center !important;
}

.validation-summary {
  background: #fff !important;
  padding: 4px 12px !important;
  border-radius: 4px !important;
  border: 1px solid #e9ecef !important;
  width: 100% !important;
}

.validation-item {
  display: flex !important;
  justify-content: space-between !important;
  margin-bottom: 1px !important;
  font-size: 11px !important;
  text-align: left !important;
}

.validation-label {
  font-weight: 600 !important;
  color: #666 !important;
  font-size: 10px !important;
  text-align: left !important;
}

.validation-value {
  font-weight: 700 !important;
  font-size: 11px !important;
  color: #333 !important;
  text-align: right !important;
}

.validation-line {
  height: 4px !important;
  background: #e9ecef !important;
  border-radius: 2px !important;
  margin-top: 3px !important;
  overflow: hidden !important;
}

.line-fill {
  height: 100% !important;
  background: #28a745 !important;
  border-radius: 2px !important;
  transition: all 0.3s ease !important;
  width: 0% !important;
}

.line-fill.warning {
  background: #ffc107 !important;
}

.line-fill.danger {
  background: #dc3545 !important;
}

.validation-message {
  display: block !important; /* Changed from none to block */
  padding: 8px 12px;
  border-radius: 4px;
  margin-top: 5px;
  font-size: 12px;
  font-weight: 500;
}

.validation-message.danger {
  background-color: #f8d7da;
  color: #721c24;
  border: 1px solid #f5c6cb;
}

.validation-message.warning {
  background-color: #fff3cd;
  color: #856404;
  border: 1px solid #ffc107;
}

/* ============================================ */
/* RESPONSIVE                                   */
/* ============================================ */

@media (max-width: 992px) {
  .project-main-wrapper {
    grid-template-columns: 1fr !important;
  }
  
  .project-left-side {
    border-right: none !important;
    border-bottom: 1px solid #e9ecef !important;
  }
  
  .project-right-side {
    padding: 8px 14px !important;
  }
}

@media (max-width: 768px) {
  .obl-content {
    grid-template-columns: 1fr !important;
    gap: 8px !important;
  }
  
  .obl-input {
    max-width: 100% !important;
  }
  
  .project-left-side .detail-label {
    min-width: 80px !important;
  }
}
  .historical-toggle-header {
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

.historical-toggle-header:hover {
  background: #e9ecef !important;
}

.historical-toggle-header h3 {
  margin: 0 !important;
  font-size: 16px !important;
  font-weight: 600 !important;
  color: #495057 !important;
}

#toggleHistoricalBtn {
  padding: 4px 16px !important;
  font-size: 13px !important;
  font-weight: 500 !important;
  background: #007bff !important;
  color: white !important;
  border: none !important;
  border-radius: 4px !important;
  cursor: pointer !important;
  transition: background 0.2s !important;
  flex-shrink: 0 !important;
}

#toggleHistoricalBtn:hover {
  opacity: 0.85 !important;
}

.historical-container {
  transition: all 0.3s ease !important;
}
span.actual-revenue-display {
    float: right;
    border: 1px solid red;
    padding: 3px 8px 3px 7px;
}

/* Frozen state styles */
.frozen-project {
  background-color: #fff8e1 !important;
  border-left: 4px solid #ffc107 !important;
}

.frozen-input {
  background-color: #f5f5f5 !important;
  color: #6c757d !important;
  cursor: not-allowed !important;
  opacity: 0.7 !important;
  pointer-events: none !important;
}

.frozen-textarea {
  background-color: #f5f5f5 !important;
  color: #6c757d !important;
  cursor: not-allowed !important;
  opacity: 0.7 !important;
  pointer-events: none !important;
}

.freeze-badge {
  display: inline-block;
  padding: 2px 8px;
  background-color: #ffc107;
  color: #856404;
  border-radius: 4px;
  font-size: 11px;
  font-weight: bold;
  margin-left: 8px;
}

/* Frozen row in table */
tr.frozen-row {
  background-color: #fff8e1 !important;
  border-left: 4px solid #ffc107 !important;
}

tr.frozen-row:hover {
  background-color: #fff3cd !important;
}
   .revenue-row {
  grid-template-columns: 1fr 1fr 1fr 1fr 1fr !important;
}
  .month-actual-input {
    width: 100% !important;
    padding: 4px 8px !important;
    border: 1px solid #ced4da !important;
    border-radius: 4px !important;
    font-size: 12px !important;
    text-align: center !important;
    background-color: #e8f5e9 !important;
    transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out !important;
}

.month-actual-input:focus {
    border-color: #28a745 !important;
    outline: 0 !important;
    box-shadow: 0 0 0 0.15rem rgba(40, 167, 69, 0.25) !important;
}

.month-actual-input.input-with-value {
    background-color: #c8e6c9 !important;
    font-weight: bold !important;
    color: #1b5e20 !important;
}

.month-actual-input.input-empty {
    background-color: #f5f5f5 !important;
    color: #999 !important;
}


`,
    script: () => {
        setTimeout(async () => {
            const siteUrl = typeof _spPageContextInfo !== "undefined" ? _spPageContextInfo.webAbsoluteUrl : window.location.origin + "/sites/DigitalTransformationTest";
            const listName = "Revenue Forecast List";
            const projectListName = "Project List";
            // ============================================
            // LOCAL PERMISSION WRAPPERS
            // ============================================
            // Re-export auth functions with consistent naming for local use
            const canEditProject = (project, forecast) => authCanEditProject(project);
            const canCreateForecast = () => canCreateRecord();
            // const getAssignedProjectsLocal = (allProjects: ProjectItem[]): ProjectItem[] => getAssignedProjects(allProjects);
            const canImportData = () => canContribute();
            const canExportData = () => canRead();
            // ============================================
            // UI PERMISSION UPDATE
            // ============================================
            const updateUIBasedOnPermissions = () => {
                // Use authUpdateUI with correct selectors
                authUpdateUI({
                    importBtn: '.import-actual-btn',
                    exportBtn: '.export-btn',
                    addBtn: '.show-form-btn',
                    saveBtn: '.btn-primary[type="submit"]',
                    freezeCheckboxes: '.freeze-checkbox'
                });
                // Additional UI updates specific to this page
                const exportDetailsBtn = document.querySelector('.export-details-btn');
                if (exportDetailsBtn) {
                    exportDetailsBtn.style.display = canExportData() ? '' : 'none';
                }
                // In the event listeners section, after exportDetailsBtn
                const exportHistoricalBtn = document.querySelector('.export-historical-btn');
                if (exportHistoricalBtn) {
                    exportHistoricalBtn.style.display = canExportData() ? '' : 'none';
                }
            };
            // ============================================
            // PERMISSION CHECK FUNCTIONS
            // ============================================
            const canViewForm = (isEdit = false, project, forecast) => {
                if (!canRead())
                    return false;
                if (isEdit) {
                    return canEditProject(project, forecast);
                }
                else {
                    return canCreateForecast();
                }
            };
            // DOM Elements
            const tableBody = document.querySelector(".revenue-table-body-rev");
            const monthColumnsHeader = document.querySelector("#monthColumnsHeader");
            const formSection = document.querySelector(".form-section");
            const tableSection = document.querySelector(".table-section");
            const projectSelect = document.querySelector("#projectSelect");
            const historicalContainer = document.querySelector("#historicalContainer");
            const firstRowContainer = document.querySelector("#firstRowContainer");
            const remainingContainer = document.querySelector("#remainingContainer");
            const yearlyContainer = document.querySelector("#yearlyContainer");
            const projectDetails = document.querySelector("#projectDetails");
            const backBtn = document.querySelector(".back-to-table");
            const form = document.querySelector("#revenueForm");
            const exportBtn = document.querySelector(".export-btn");
            const oblLastMonthInput = document.querySelector("#oblLastMonth");
            const totalCurrentValueSpan = document.querySelector("#totalCurrentValue");
            const remainingValueSpan = document.querySelector("#remainingValue");
            const oblLineFill = document.querySelector("#oblLineFill");
            const oblValidationMessage = document.querySelector("#oblValidationMessage");
            const formTitle = document.querySelector("#formTitle");
            // Filter Elements
            const filterYear = document.querySelector("#filterYear");
            const filterMonth = document.querySelector("#filterMonth");
            const filterProject = document.querySelector("#filterProject");
            const filterCustomer = document.querySelector("#filterCustomer");
            const filterProjectManager = document.querySelector("#filterProjectManager");
            const filterGroupHead = document.querySelector("#filterGroupHead");
            const filterProjectReference = document.querySelector("#filterProjectReference");
            // Action Buttons
            const importActualBtns = document.querySelector(".import-actual-btn");
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
            const remarksInternal = document.querySelector("#remarksInternal");
            const remarksSprintCell = document.querySelector("#remarksSprintCell");
            let projects = [];
            let forecasts = [];
            let editingId = null;
            let allMonths = [];
            // Pagination variables
            let currentPage = 1;
            let pageSize = 10;
            let totalPages = 1;
            let filteredData = [];
            const pageSizeSelect = document.querySelector("#pageSizeSelect");
            const getSafeString = (value, defaultValue = "-") => {
                var _a;
                if (value === null || value === undefined) {
                    return defaultValue;
                }
                if (typeof value === 'string') {
                    return value.trim() || defaultValue;
                }
                if (typeof value === 'object' && 'Title' in value) {
                    return ((_a = value.Title) === null || _a === void 0 ? void 0 : _a.trim()) || defaultValue;
                }
                return String(value) || defaultValue;
            };
            const toast = (icon, title) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });
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
                return `${project.Title} - ${currentYear} Forecast - ${totalValue.toLocaleString()}`;
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
                return generateIdealMonthStructure(); // Returns all months including 36 historical
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
            // ============================================
            // OBL TOAST FUNCTIONS - CORRECTED
            // ============================================
            // Helper function to get numeric value from input
            const getNumericValue = (input) => {
                const value = input.value.trim();
                if (value === '' || value === '0')
                    return 0;
                return parseInt(value) || 0;
            };
            // Helper function to set numeric value in input
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
            // SHOW OBL TOAST - CORRECTED 
            const showOBLToast = (type, message) => {
                var _a;
                // Remove existing toast
                const existingToast = document.querySelector('.obl-error-toast');
                if (existingToast)
                    existingToast.remove();
                const toast = document.createElement('div');
                toast.className = `obl-error-toast ${type}`;
                const icons = {
                    success: '✅',
                    info: 'ℹ️',
                    warning: '⚠️',
                    danger: '❌'
                };
                const colors = {
                    success: '#d4edda',
                    info: '#d1ecf1',
                    warning: '#fff3cd',
                    danger: '#f8d7da'
                };
                const textColors = {
                    success: '#155724',
                    info: '#0c5460',
                    warning: '#856404',
                    danger: '#721c24'
                };
                const borderColors = {
                    success: '#28a745',
                    info: '#17a2b8',
                    warning: '#ffc107',
                    danger: '#dc3545'
                };
                toast.innerHTML = `
    <span class="toast-icon">${icons[type]}</span>
    <span class="toast-message">${message}</span>
    <button class="toast-close">&times;</button>
  `;
                toast.style.cssText = `
    position: fixed !important;
    top: 20px !important;
    right: 20px !important;
    background-color: ${colors[type]} !important;
    color: ${textColors[type]} !important;
    padding: 15px 20px !important;
    border-radius: 8px !important;
    border-left: 4px solid ${borderColors[type]} !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
    z-index: 10000 !important;
    display: flex !important;
    align-items: center !important;
    gap: 12px !important;
    min-width: 350px !important;
    max-width: 500px !important;
    animation: slideInRight 0.3s ease-out !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    opacity: 1 !important;
    visibility: visible !important;
  `;
                // Add animation styles if not already present
                if (!document.querySelector('#toast-animations')) {
                    const style = document.createElement('style');
                    style.id = 'toast-animations';
                    style.textContent = `
      @keyframes slideInRight {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
      }
      @keyframes slideOutRight {
        from { transform: translateX(0); opacity: 1; }
        to { transform: translateX(100%); opacity: 0; }
      }
      .obl-error-toast .toast-icon {
        font-size: 20px;
        flex-shrink: 0;
      }
      .obl-error-toast .toast-message {
        flex: 1;
      }
      .obl-error-toast .toast-close {
        background: none;
        border: none;
        font-size: 20px;
        cursor: pointer;
        color: inherit;
        opacity: 0.6;
        padding: 0 4px;
        flex-shrink: 0;
      }
      .obl-error-toast .toast-close:hover {
        opacity: 1;
      }
    `;
                    document.head.appendChild(style);
                }
                document.body.appendChild(toast);
                // Auto dismiss after 5 seconds
                const timeoutId = setTimeout(() => {
                    if (toast.parentNode) {
                        toast.style.animation = 'slideOutRight 0.3s ease-in';
                        setTimeout(() => {
                            if (toast.parentNode)
                                toast.remove();
                        }, 300);
                    }
                }, 5000);
                // Close button handler
                (_a = toast.querySelector('.toast-close')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
                    clearTimeout(timeoutId);
                    toast.style.animation = 'slideOutRight 0.3s ease-in';
                    setTimeout(() => {
                        if (toast.parentNode)
                            toast.remove();
                    }, 300);
                });
            };
            // CALCULATE TOTAL CURRENT VALUE - CORRECTED 
            const calculateTotalCurrentValue = () => {
                let total = 0;
                // Helper to get numeric value
                const getVal = (input) => getNumericValue(input);
                // ONLY count item-current-input values (these represent individual line items)
                // The main "Curr. Plan" is the sum of all items, so counting both would double count
                document.querySelectorAll('.detailed-month-box .item-current-input').forEach(input => {
                    total += getVal(input);
                });
                // Long-term Forecast (Next 15 Months) - Curr. Plan values (these are standalone, not itemized)
                document.querySelectorAll('.simple-month-rev .main-current-input:not([readonly])').forEach(input => {
                    total += getVal(input);
                });
                // 5-Year Forecast - Curr. Plan values (these are standalone, not itemized)
                document.querySelectorAll('.yearly-current-input-rev:not([readonly])').forEach(input => {
                    total += getVal(input);
                });
                return total;
            };
            // VALIDATE OBL LIMIT - FIXED
            const validateOBLLimit = (changedInput) => {
                const oblValue = getNumericValue(oblLastMonthInput);
                const totalCurrent = calculateTotalCurrentValue();
                totalCurrentValueSpan.textContent =
                    totalCurrent.toLocaleString();
                const remaining = oblValue - totalCurrent;
                remainingValueSpan.textContent =
                    remaining.toLocaleString();
                // Remove existing OBL toast
                const existingToast = document.querySelector(".obl-error-toast");
                if (existingToast) {
                    existingToast.remove();
                }
                // =====================================================
                // OBL = 0
                // =====================================================
                if (oblValue === 0) {
                    oblLineFill.style.width = "0%";
                    oblLineFill.className = "line-fill";
                    if (totalCurrent > 0) {
                        oblValidationMessage.textContent =
                            `❌ Liquidation is more than OBL (More by ${totalCurrent.toLocaleString()})`;
                        oblValidationMessage.className =
                            "validation-message danger";
                        showOBLToast("danger", `❌ Liquidation is more than OBL (More by ${totalCurrent.toLocaleString()})`);
                        blockAllCurrentInputs();
                        return false;
                    }
                    oblValidationMessage.textContent =
                        "ℹ️ OBL is set to zero.";
                    oblValidationMessage.className =
                        "validation-message";
                    showOBLToast("info", "ℹ️ OBL is set to zero.");
                    unblockAllCurrentInputs();
                    return true;
                }
                // =====================================================
                // Calculate percentage
                // =====================================================
                const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
                oblLineFill.style.width =
                    `${percentage}%`;
                oblLineFill.className =
                    "line-fill";
                if (percentage > 90) {
                    oblLineFill.classList.add("danger");
                }
                else if (percentage > 75) {
                    oblLineFill.classList.add("warning");
                }
                // =====================================================
                // 1. LIQUIDATION > OBL
                // =====================================================
                if (totalCurrent > oblValue) {
                    const exceedAmount = totalCurrent - oblValue;
                    oblValidationMessage.textContent =
                        `❌ Liquidation is more than OBL (More by ${exceedAmount.toLocaleString()})`;
                    oblValidationMessage.className =
                        "validation-message danger";
                    showOBLToast("danger", `❌ Liquidation is more than OBL (More by ${exceedAmount.toLocaleString()})`);
                    blockIncreasingCurrentInputs(changedInput);
                    return false;
                }
                // =====================================================
                // 2. LIQUIDATION = OBL
                // =====================================================
                if (totalCurrent === oblValue) {
                    oblValidationMessage.textContent =
                        `✅ OBL is fully liquidated (${totalCurrent.toLocaleString()})`;
                    oblValidationMessage.className =
                        "validation-message success";
                    showOBLToast("success", `✅ OBL is fully liquidated (${totalCurrent.toLocaleString()})`);
                    // Don't allow increasing values beyond OBL
                    blockIncreasingCurrentInputs(changedInput);
                    return true;
                }
                // =====================================================
                // 3. LIQUIDATION < OBL
                // =====================================================
                if (totalCurrent < oblValue) {
                    const remainingAmount = oblValue - totalCurrent;
                    oblValidationMessage.textContent =
                        `⚠️ OBL is not fully liquidated (Remaining ${remainingAmount.toLocaleString()})`;
                    oblValidationMessage.className =
                        "validation-message warning";
                    showOBLToast("warning", `⚠️ OBL is not fully liquidated (Remaining ${remainingAmount.toLocaleString()})`);
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
            const updateOBLValidationDisplay = (changedInput) => {
                validateOBLLimit(changedInput);
                validateMonthItemTotals(changedInput);
            };
            // ============================================
            // ATTACH FORECAST EVENT LISTENERS - ADD OBL REAL-TIME VALIDATION
            // ============================================
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
                const allMonths = generateIdealMonthStructure();
                const freshData = {};
                const now = new Date();
                const currentMonthKey = formatMonthKey(now);
                const currentMonthIndex = allMonths.indexOf(currentMonthKey);
                let detailedStartIndex = 0;
                if (currentMonthIndex !== -1) {
                    detailedStartIndex = currentMonthIndex;
                }
                else {
                    detailedStartIndex = 36;
                }
                allMonths.forEach((monthKey, index) => {
                    const isDetailedMonth = index >= detailedStartIndex &&
                        index < detailedStartIndex + 3;
                    const monthData = {
                        plan: 0,
                        actual: 0,
                        forecast: 0,
                        monthActual: 0, // ✅ Initialize monthActual
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
            const isForecastMonthData = (data) => {
                return data &&
                    typeof data.plan === 'number' &&
                    typeof data.actual === 'number';
            };
            // const handleMonthProgression = (forecastData: ForecastData): ForecastData => {
            //     const now = new Date();
            //     const currentMonthKey = formatMonthKey(now);
            //     const originalMonths = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
            //     const idealMonths = generateIdealMonthStructure();
            //     // Check if we need progression
            //     const hasCurrentMonth = originalMonths.includes(currentMonthKey);
            //     const lastDataMonth = originalMonths[originalMonths.length - 1];
            //     const lastIdealMonth = idealMonths[idealMonths.length - 1];
            //     // If data already has current month and last month matches, NO progression needed
            //     if (hasCurrentMonth && lastDataMonth === lastIdealMonth) {
            //         return forecastData;
            //     }
            //     // Find current month index in idealMonths
            //     const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
            //     // Detailed months start at current month
            //     const detailedStartIndex = currentMonthIndex !== -1 ? currentMonthIndex : 36;
            //     // If we get here, we need to add new months and carry values forward
            //     const updatedData: ForecastData = {};
            //     // Copy all existing data first
            //     idealMonths.forEach((monthKey, index) => {
            //         const existingData = forecastData[monthKey];
            //         const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
            //         if (existingData && isForecastMonthData(existingData)) {
            //             updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
            //         } else {
            //             updatedData[monthKey] = {
            //                 plan: 0,
            //                 actual: 0,
            //                 items: isDetailedMonth ? createEmptyItems() : {}
            //             };
            //         }
            //     });
            //     console.log("=== APPLYING VALUE PROPAGATION ===");
            //     // ✅ FIX: Only carry values from OLD months to NEW months
            //     // Carry values forward: previous month's actual becomes next month's plan
            //     for (let i = 1; i < idealMonths.length; i++) {
            //         const prevMonthKey = idealMonths[i - 1];
            //         const currMonthKey = idealMonths[i];
            //         const prevMonthData = updatedData[prevMonthKey];
            //         const currMonthData = updatedData[currMonthKey];
            //         if (prevMonthData && isForecastMonthData(prevMonthData) && 
            //             currMonthData && isForecastMonthData(currMonthData)) {
            //             // Only apply progression if the current month is NEW (doesn't exist in original data)
            //             // OR if the previous month is older than current month and current month is new
            //             const isPrevMonthOlder = prevMonthKey < currentMonthKey;
            //             const isCurrMonthNew = !originalMonths.includes(currMonthKey);
            //             if (isPrevMonthOlder && isCurrMonthNew) {
            //                 // Carry forward: previous month's actual becomes current month's plan
            //                 currMonthData.plan = prevMonthData.actual;
            //                 // Keep current month's actual as 0 (ready for new input)
            //                 currMonthData.actual = 0;
            //                 // Also carry forward items for detailed months
            //                 const monthIndex = idealMonths.indexOf(currMonthKey);
            //                 const isDetailedMonth = monthIndex >= detailedStartIndex && monthIndex < detailedStartIndex + 3;
            //                 if (isDetailedMonth && currMonthData.items) {
            //                     const prevItems = prevMonthData.items || {};
            //                     Object.keys(currMonthData.items).forEach(itemKey => {
            //                         const prevItem = prevItems[itemKey];
            //                         const currItem = currMonthData.items![itemKey];
            //                         if (prevItem && currItem) {
            //                             currItem.plan = prevItem.actual || 0;
            //                             currItem.actual = 0;
            //                         }
            //                     });
            //                 }
            //                 console.log(`📊 ${prevMonthKey} (actual: ${prevMonthData.actual}) → ${currMonthKey} (plan: ${currMonthData.plan})`);
            //             }
            //         }
            //     }
            //     // Preserve yearly data
            //     if (forecastData.yearly) {
            //         updatedData.yearly = JSON.parse(JSON.stringify(forecastData.yearly));
            //     } else {
            //         const lastMonthKey = idealMonths[idealMonths.length - 1];
            //         updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
            //     }
            //     return updatedData;
            // };
            const propagateMonthValues = (forecastData) => {
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
                const updatedData = {};
                idealMonths.forEach((monthKey, index) => {
                    const existingData = forecastData[monthKey];
                    const isDetailedMonth = index >= detailedStartIndex && index < detailedStartIndex + 3;
                    if (existingData && isForecastMonthData(existingData)) {
                        updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
                    }
                    else {
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
                            }
                            else {
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
                                    const currItem = currMonthData.items[itemKey];
                                    if (prevItem && currItem) {
                                        if (prevMonthActual > 0) {
                                            currItem.plan = prevMonthActual;
                                        }
                                        else {
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
                }
                else {
                    const lastMonthKey = idealMonths[idealMonths.length - 1];
                    updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
                }
                return updatedData;
            };
            const applyAutoProgression = async () => {
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
                    toast("success", `Monthly data progressed for ${progressedCount} revenue forecasts`);
                }
            };
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
                    // Log all projects with PMs for debugging
                    console.log('📋 All projects with PMs:');
                    projects.forEach(p => {
                        var _a;
                        const pmName = typeof p.ProjectManager === 'object'
                            ? (_a = p.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title
                            : p.ProjectManager;
                        console.log(`  - ${p.Title}: PM = ${pmName || 'None'}`);
                    });
                    // Get assigned projects based on user permissions
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
                }
                catch (error) {
                    console.error("Error loading projects:", error);
                    toast("error", "Failed to load projects");
                }
            };
            const getAssignedProjectValues = (key) => {
                const assignedProjects = getAssignedProjects(projects);
                const values = new Set();
                assignedProjects.forEach(project => {
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
            const loadForecasts = async () => {
                var _a;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,RevenueActual,Remarks,CreatedDate,UpdatedDate,Freeze&$expand=Project&$top=5000`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const data = await res.json();
                    forecasts = (((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || []).map(forecast => (Object.assign(Object.assign({}, forecast), { ForecastData: typeof forecast.ForecastData === 'string' ?
                            forecast.ForecastData :
                            JSON.stringify(forecast.ForecastData || {}), RevenueActual: typeof forecast.RevenueActual === 'string' ?
                            forecast.RevenueActual :
                            JSON.stringify(forecast.RevenueActual || {}), Freeze: forecast.Freeze || false })));
                    console.log("✅ Forecasts loaded:", forecasts.length);
                }
                catch (error) {
                    console.error("Error loading forecasts:", error);
                    toast("error", "Failed to load forecasts");
                }
            };
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
                const exportData = filteredData.length > 0 ? filteredData : forecasts;
                if (exportData.length === 0) {
                    toast("error", "No data available to export");
                    return;
                }
                try {
                    const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);
                    const excelData = exportData.map(forecast => {
                        var _a, _b;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        const oblValue = forecast.OBL || 0;
                        const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
                        const monthName = getMonthName(selectedMonth);
                        const selectedMonthForecastData = forecastData[selectedMonthKey] || { plan: 0, actual: 0, forecast: 0 };
                        const selectedMonthPlan = selectedMonthForecastData.plan || 0;
                        const selectedMonthCurrent = selectedMonthForecastData.actual || 0;
                        const selectedMonthActual = revenueActualData[selectedMonthKey] || 0;
                        // Further = forecast value from backend
                        const selectedMonthFurther = selectedMonthForecastData.forecast || 0;
                        // Forecast = Curr. Plan + Further
                        const selectedMonthForecast = selectedMonthCurrent + selectedMonthFurther;
                        const rowData = {
                            "Profit Centre": forecast.ProfitCentre || "-",
                            "SO/Project Ref": (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                            "Customer Name": (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                            "Project Name": projectTitle,
                            "Group Head": (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
                            "Project Manager": typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                            "OBL": oblValue,
                            // Order: Prev. Plan → Curr. Plan → Actual → Further → Forecast
                            [`${monthName} ${selectedYear} (Prev. Plan)`]: selectedMonthPlan,
                            [`${monthName} ${selectedYear} (Curr. Plan)`]: selectedMonthCurrent,
                            [`${monthName} ${selectedYear} (Actual)`]: selectedMonthActual,
                            [`${monthName} ${selectedYear} (Further)`]: selectedMonthFurther, // backend forecast field
                            [`${monthName} ${selectedYear} (Forecast)`]: selectedMonthForecast, // Curr. Plan + Further
                        };
                        allMonths.forEach(monthKey => {
                            const [year, month] = monthKey.split('-');
                            const monthLabel = getMonthName(month);
                            const displayLabel = `${monthLabel} ${year}`;
                            const forecastMonthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                            const planValue = forecastMonthData.plan || 0;
                            const currentValue = forecastMonthData.actual || 0;
                            rowData[`${displayLabel} (Prev. Plan)`] = planValue;
                            rowData[`${displayLabel} (Curr. Plan)`] = currentValue;
                        });
                        lastFiveYears.forEach(year => {
                            var _a;
                            const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0 };
                            rowData[`${year} (Prev. Plan)`] = yearData.plan;
                            rowData[`${year} (Curr. Plan)`] = yearData.actual;
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
                        if (header.includes('OBL') || header.includes('Balance'))
                            return { wch: 12 };
                        if (header.includes('(Prev. Plan)') || header.includes('(Curr. Plan)') || header.includes('(Further)') || header.includes('(Actual)') || header.includes('(Forecast)'))
                            return { wch: 10 };
                        return { wch: 12 };
                    });
                    ws['!cols'] = colWidths;
                    ws['!freeze'] = { x: 0, y: 1 };
                    XLSX.utils.book_append_sheet(wb, ws, "Revenue Forecast");
                    const monthName = getMonthName(selectedMonth);
                    const fileName = `Revenue_Forecast_${monthName}_${selectedYear}_Filtered.xlsx`;
                    XLSX.writeFile(wb, fileName);
                    toast("success", `Exported ${exportData.length} filtered records for ${monthName} ${selectedYear}`);
                }
                catch (error) {
                    toast("error", "Failed to export data to Excel");
                }
            };
            const generateHeadersStructure = (selectedYear, selectedMonth, exportData) => {
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
                    `${selectedMonthName} ${selectedYear} (Further)`, // backend forecast field
                    `${selectedMonthName} ${selectedYear} (Forecast)` // Curr. Plan + Further
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
                    monthlyHeaders.push(`${displayLabel} (Prev. Plan)`, `${displayLabel} (Curr. Plan)`);
                });
                const lastFiveYears = getLastFiveYears(exportData);
                const yearlyHeaders = [];
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
            const exportDetailedItemsToExcel = () => {
                var _a, _b, _c, _d, _e, _f, _g, _h, _j;
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
                    const detailedMonths = [];
                    for (let i = 0; i < 3; i++) {
                        const monthDate = new Date(currentDate);
                        monthDate.setMonth(monthDate.getMonth() + i);
                        const monthKey = formatMonthKey(monthDate);
                        if (monthKey) {
                            detailedMonths.push(monthKey);
                        }
                    }
                    const nextFifteenMonths = [];
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
                    const currentMonthName = getMonthName(((_a = detailedMonths[0]) === null || _a === void 0 ? void 0 : _a.split('-')[1]) || '') || ((_b = detailedMonths[0]) === null || _b === void 0 ? void 0 : _b.split('-')[1]) || '';
                    const nextMonth1Name = getMonthName(((_c = detailedMonths[1]) === null || _c === void 0 ? void 0 : _c.split('-')[1]) || '') || ((_d = detailedMonths[1]) === null || _d === void 0 ? void 0 : _d.split('-')[1]) || '';
                    const nextMonth2Name = getMonthName(((_e = detailedMonths[2]) === null || _e === void 0 ? void 0 : _e.split('-')[1]) || '') || ((_f = detailedMonths[2]) === null || _f === void 0 ? void 0 : _f.split('-')[1]) || '';
                    const currentMonthYear = ((_g = detailedMonths[0]) === null || _g === void 0 ? void 0 : _g.split('-')[0]) || '';
                    const nextMonth1Year = ((_h = detailedMonths[1]) === null || _h === void 0 ? void 0 : _h.split('-')[0]) || '';
                    const nextMonth2Year = ((_j = detailedMonths[2]) === null || _j === void 0 ? void 0 : _j.split('-')[0]) || '';
                    const excelData = [];
                    // First row headers - Order: Prev. Plan → Curr. Plan → Actual → Further → Forecast
                    const firstRowHeaders = [
                        "SO/Project Ref",
                        "Project Name",
                        "Customer Name",
                        "Profit Centre",
                        "Group Head",
                        "Project Manager",
                        "OBL",
                        `${selectedMonthName} ${selectedYear} (Prev. Plan)`,
                        `${selectedMonthName} ${selectedYear} (Curr. Plan)`,
                        `${selectedMonthName} ${selectedYear} (Actual)`,
                        `${selectedMonthName} ${selectedYear} (Further)`, // backend forecast field
                        `${selectedMonthName} ${selectedYear} (Forecast)`, // Curr. Plan + Further
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
                        "Item Description", "Prev. Plan", "Curr. Plan", "Actual", "Further", "Forecast",
                        "", "",
                        "Item Description", "Prev. Plan", "Curr. Plan", "Actual", "Further", "Forecast",
                        "", "",
                        "Item Description", "Prev. Plan", "Curr. Plan", "Actual", "Further", "Forecast",
                    ];
                    const remainingColumns = firstRowHeaders.length - secondRowHeaders.length;
                    for (let i = 0; i < remainingColumns; i++) {
                        secondRowHeaders.push("");
                    }
                    excelData.push(firstRowHeaders);
                    excelData.push(secondRowHeaders);
                    exportData.forEach(forecast => {
                        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
                        const oblValue = forecast.OBL || 0;
                        const detailedMonthsData = [];
                        for (let i = 0; i < 3; i++) {
                            const monthKey = detailedMonths[i];
                            if (monthKey) {
                                const monthForecastData = forecastData[monthKey] || { plan: 0, actual: 0, forecast: 0, items: createEmptyItems() };
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
                            }
                            else {
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
                            const monthData = forecastData[monthKey];
                            return sum + ((monthData === null || monthData === void 0 ? void 0 : monthData.actual) || 0);
                        }, 0);
                        const totalYearlyCurrent = forecastData.yearly ?
                            Object.values(forecastData.yearly).reduce((sum, yearData) => {
                                return sum + (yearData.actual || 0);
                            }, 0) : 0;
                        const balance = oblValue - (totalMonthlyCurrent + totalYearlyCurrent);
                        for (let itemNum = 1; itemNum <= 10; itemNum++) {
                            const itemKey = itemNum.toString();
                            const rowData = [
                                (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                                projectTitle,
                                (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                                forecast.ProfitCentre || "-",
                                (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
                                typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                                oblValue,
                                ((_c = detailedMonthsData[0]) === null || _c === void 0 ? void 0 : _c.plan) || 0,
                                ((_d = detailedMonthsData[0]) === null || _d === void 0 ? void 0 : _d.current) || 0,
                                ((_e = detailedMonthsData[0]) === null || _e === void 0 ? void 0 : _e.actual) || 0,
                                ((_f = detailedMonthsData[0]) === null || _f === void 0 ? void 0 : _f.further) || 0, // backend forecast field
                                ((_g = detailedMonthsData[0]) === null || _g === void 0 ? void 0 : _g.total) || 0, // Curr. Plan + Further
                            ];
                            const currentMonthItems = ((_h = detailedMonthsData[0]) === null || _h === void 0 ? void 0 : _h.items) || createEmptyItems();
                            const currentMonthItem = currentMonthItems[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
                            // Further = forecast value from backend
                            const furtherValue = currentMonthItem.forecast || 0;
                            // Forecast = Curr. Plan + Further
                            const itemTotal = (currentMonthItem.actual || 0) + furtherValue;
                            rowData.push(currentMonthItem.name || `Item ${itemNum}`, currentMonthItem.plan, currentMonthItem.actual, currentMonthItem.forecast || 0, // Further (backend forecast field)
                            furtherValue, // Further (same value)
                            itemTotal // Forecast = Curr. Plan + Further
                            );
                            rowData.push(((_j = detailedMonthsData[1]) === null || _j === void 0 ? void 0 : _j.plan) || 0, ((_k = detailedMonthsData[1]) === null || _k === void 0 ? void 0 : _k.current) || 0);
                            const nextMonth1Items = ((_l = detailedMonthsData[1]) === null || _l === void 0 ? void 0 : _l.items) || createEmptyItems();
                            const nextMonth1Item = nextMonth1Items[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0 };
                            const nextMonth1Further = nextMonth1Item.forecast || 0;
                            const nextMonth1Total = (nextMonth1Item.actual || 0) + nextMonth1Further;
                            rowData.push(nextMonth1Item.name || `Item ${itemNum}`, nextMonth1Item.plan, nextMonth1Item.actual, nextMonth1Item.forecast || 0, // Further (backend forecast field)
                            nextMonth1Further, // Further (same value)
                            nextMonth1Total // Forecast = Curr. Plan + Further
                            );
                            rowData.push(((_m = detailedMonthsData[2]) === null || _m === void 0 ? void 0 : _m.plan) || 0, ((_o = detailedMonthsData[2]) === null || _o === void 0 ? void 0 : _o.current) || 0);
                            const nextMonth2Items = ((_p = detailedMonthsData[2]) === null || _p === void 0 ? void 0 : _p.items) || createEmptyItems();
                            const nextMonth2Item = nextMonth2Items[itemKey] || { name: '', plan: 0, actual: 0, forecast: 0 };
                            const nextMonth2Further = nextMonth2Item.forecast || 0;
                            const nextMonth2Total = (nextMonth2Item.actual || 0) + nextMonth2Further;
                            rowData.push(nextMonth2Item.name || `Item ${itemNum}`, nextMonth2Item.plan, nextMonth2Item.actual, nextMonth2Item.forecast || 0, // Further (backend forecast field)
                            nextMonth2Further, // Further (same value)
                            nextMonth2Total // Forecast = Curr. Plan + Further
                            );
                            nextFifteenMonths.forEach(monthKey => {
                                const monthForecastData = forecastData[monthKey] || { plan: 0, actual: 0 };
                                rowData.push(monthForecastData.plan || 0, monthForecastData.actual || 0);
                            });
                            lastFiveYears.forEach(year => {
                                var _a;
                                const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0 };
                                rowData.push(yearData.plan, yearData.actual);
                            });
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
                    const numberStyle = Object.assign(Object.assign({}, dataStyle), { alignment: { horizontal: "right", vertical: "center" } });
                    const textStyle = Object.assign(Object.assign({}, dataStyle), { alignment: { horizontal: "left", vertical: "center" } });
                    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s)
                            ws[cellAddress].s = {};
                        ws[cellAddress].s = mainHeaderStyle;
                    }
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 1, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s)
                            ws[cellAddress].s = {};
                        ws[cellAddress].s = subHeaderStyle;
                    }
                    for (let row = range.s.r + 2; row <= range.e.r; row++) {
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                            if (!ws[cellAddress])
                                continue;
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
                                (col >= 33)) {
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
                        { wch: 12 }, // SO/Project Ref
                        { wch: 25 }, // Project Name
                        { wch: 20 }, // Customer Name
                        { wch: 12 }, // Profit Centre
                        { wch: 15 }, // Group Head
                        { wch: 15 }, // Project Manager
                        { wch: 10 }, // OBL
                        { wch: 12 }, // Selected Month Prev. Plan
                        { wch: 12 }, // Selected Month Curr. Plan
                        { wch: 12 }, // Selected Month Actual
                        { wch: 12 }, // Selected Month Further
                        { wch: 12 }, // Selected Month Forecast
                        { wch: 20 }, // Current Month Header
                        { wch: 15 }, // Item Description
                        { wch: 10 }, // Prev. Plan
                        { wch: 10 }, // Curr. Plan
                        { wch: 10 }, // Actual
                        { wch: 10 }, // Further (backend forecast)
                        { wch: 10 }, // Forecast (Curr. Plan + Further)
                        { wch: 12 }, // Next Month 1 Prev. Plan
                        { wch: 12 }, // Next Month 1 Curr. Plan
                        { wch: 20 }, // Next Month 1 Header
                        { wch: 15 }, // Item Description
                        { wch: 10 }, // Prev. Plan
                        { wch: 10 }, // Curr. Plan
                        { wch: 10 }, // Actual
                        { wch: 10 }, // Further (backend forecast)
                        { wch: 10 }, // Forecast (Curr. Plan + Further)
                        { wch: 12 }, // Next Month 2 Prev. Plan
                        { wch: 12 }, // Next Month 2 Curr. Plan
                        { wch: 20 }, // Next Month 2 Header
                        { wch: 15 }, // Item Description
                        { wch: 10 }, // Prev. Plan
                        { wch: 10 }, // Curr. Plan
                        { wch: 10 }, // Actual
                        { wch: 10 }, // Further (backend forecast)
                        { wch: 10 }, // Forecast (Curr. Plan + Further)
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
                            7, 8, 9, 10, 11, // Selected Month values (Prev Plan, Curr Plan, Actual, Further, Forecast)
                            21, 22, // Next Month 1 Prev Plan, Curr Plan
                            30, 31, // Next Month 2 Prev Plan, Curr Plan
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
                }
                catch (error) {
                    toast("error", "Failed to export detailed items data to Excel");
                }
            };
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
                const sortedYears = Array.from(years).sort();
                return sortedYears.slice(-5);
            };
            const exportHistoricalToExcel = () => {
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
                    const historicalHeaders = [];
                    historicalMonths.forEach(monthKey => {
                        const [year, month] = monthKey.split('-');
                        const monthLabel = getMonthName(month) || month;
                        const displayLabel = `${monthLabel} ${year}`;
                        historicalHeaders.push(`${displayLabel} (Curr. Plan)`, `${displayLabel} (Revenue Actual)`);
                    });
                    // Current month headers - 5 columns: Prev. Plan, Curr. Plan, Actual, Further, Forecast
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
                        var _a, _b;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
                        const oblValue = forecast.OBL || 0;
                        const rowData = {
                            "Profit Centre": forecast.ProfitCentre || "-",
                            "SO/Project Ref": (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                            "Customer Name": (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                            "Project Name": projectTitle,
                            "Group Head": (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
                            "Project Manager": typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                            "OBL": oblValue
                        };
                        // Add historical month data (Curr. Plan and Revenue Actual)
                        historicalMonths.forEach(monthKey => {
                            const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                            const currentValue = monthData.actual || 0;
                            const actualValue = revenueActualData[monthKey] || 0;
                            const [year, month] = monthKey.split('-');
                            const monthLabel = getMonthName(month) || month;
                            const displayLabel = `${monthLabel} ${year}`;
                            rowData[`${displayLabel} (Curr. Plan)`] = currentValue;
                            rowData[`${displayLabel} (Revenue Actual)`] = actualValue;
                        });
                        // Add current month data
                        const currentMonthData = forecastData[currentMonthKey] || { plan: 0, actual: 0, forecast: 0 };
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
                    const numberStyle = Object.assign(Object.assign({}, dataStyle), { alignment: { horizontal: "right", vertical: "center" } });
                    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                    // Apply header style
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (!ws[cellAddress])
                            continue;
                        if (!ws[cellAddress].s)
                            ws[cellAddress].s = {};
                        ws[cellAddress].s = headerStyle;
                    }
                    // Apply data styles
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
                    // Column widths
                    const colWidths = headers.map((header) => {
                        if (header.includes('Project Name'))
                            return { wch: 25 };
                        if (header.includes('Customer Name'))
                            return { wch: 20 };
                        if (header.includes('Profit Centre') || header.includes('SO/Project Ref'))
                            return { wch: 12 };
                        if (header.includes('Group Head') || header.includes('Project Manager'))
                            return { wch: 15 };
                        if (header.includes('OBL'))
                            return { wch: 12 };
                        if (header.includes('Curr. Plan') || header.includes('Revenue Actual') || header.includes('Prev. Plan') || header.includes('Actual') || header.includes('Further') || header.includes('Forecast'))
                            return { wch: 12 };
                        return { wch: 12 };
                    });
                    ws['!cols'] = colWidths;
                    ws['!freeze'] = { x: 7, y: 1 };
                    XLSX.utils.book_append_sheet(wb, ws, "Historical Revenue");
                    const fileName = `Revenue_Historical_36Months_${formatMonthLabel(new Date())}.xlsx`;
                    XLSX.writeFile(wb, fileName);
                    toast("success", `Exported ${exportData.length} records with 36 months historical data`);
                }
                catch (error) {
                    console.error("Error exporting historical data:", error);
                    toast("error", "Failed to export historical data");
                }
            };
            // COMPLETE REPLACEMENT FOR setupPaginationEventListeners
            const setupPaginationEventListeners = () => {
                console.log("🔧 Setting up pagination event listeners...");
                // Initialize page size from dropdown
                if (pageSizeSelect) {
                    pageSize = parseInt(pageSizeSelect.value) || 10;
                    console.log(`📊 Current pageSize: ${pageSize}`);
                }
                // Remove any existing listeners by cloning and replacing pagination buttons
                const paginationButtons = [firstPageBtn, prevPageBtn, nextPageBtn, lastPageBtn];
                paginationButtons.forEach(btn => {
                    var _a;
                    if (btn) {
                        const newBtn = btn.cloneNode(true);
                        (_a = btn.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newBtn, btn);
                    }
                });
                // Re-assign references after cloning
                const newFirstPageBtn = document.querySelector("#firstPage");
                const newPrevPageBtn = document.querySelector("#prevPage");
                const newNextPageBtn = document.querySelector("#nextPage");
                const newLastPageBtn = document.querySelector("#lastPage");
                if (newFirstPageBtn) {
                    newFirstPageBtn.addEventListener("click", () => goToPage(1));
                    Object.assign(firstPageBtn, newFirstPageBtn);
                }
                if (newPrevPageBtn) {
                    newPrevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
                    Object.assign(prevPageBtn, newPrevPageBtn);
                }
                if (newNextPageBtn) {
                    newNextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
                    Object.assign(nextPageBtn, newNextPageBtn);
                }
                if (newLastPageBtn) {
                    newLastPageBtn.addEventListener("click", () => goToPage(totalPages));
                    Object.assign(lastPageBtn, newLastPageBtn);
                }
                // Fix page size selector - Use event delegation on parent container
                // This way we don't need to worry about cloning and losing event listeners
                const paginationContainer = document.querySelector('.pagination-container');
                if (paginationContainer) {
                    // Remove any existing change listener on the page size select
                    // Use event delegation on the container
                    paginationContainer.addEventListener('change', function (e) {
                        const target = e.target;
                        if (target && target.id === 'pageSizeSelect') {
                            const newSize = parseInt(target.value);
                            console.log(`📊 Page size changed to: ${newSize}`);
                            changePageSize(newSize);
                        }
                    });
                }
                console.log(`📊 Page size selector initialized with value: ${pageSizeSelect === null || pageSizeSelect === void 0 ? void 0 : pageSizeSelect.value}`);
                console.log("✅ Pagination event listeners set up successfully");
            };
            // COMPLETE REPLACEMENT FOR changePageSize
            const changePageSize = (newSize) => {
                if (!newSize || newSize <= 0) {
                    return;
                }
                console.log(`📊 Changing page size from ${pageSize} to ${newSize}`);
                console.log(`📊 Current filteredData length: ${filteredData.length}`);
                console.log(`📊 Current page: ${currentPage}`);
                // Update page size
                pageSize = newSize;
                // Reset to page 1 when changing page size
                currentPage = 1;
                // Recalculate total pages using the new page size
                calculateTotalPages();
                // Update pagination UI
                updatePaginationControls();
                // CRITICAL: Force re-render the table with new page data
                renderCurrentPage();
                // Re-render totals header
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                setTimeout(() => {
                    renderTotalsHeader(displayMonths, currentMonthKey);
                }, 100);
                toast("info", `Showing ${pageSize} records per page`);
            };
            // COMPLETE REPLACEMENT FOR renderCurrentPage
            const renderCurrentPage = () => {
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
            const updatePaginationControls = () => {
                console.log(`📊 updatePaginationControls - currentPage: ${currentPage}, totalPages: ${totalPages}, filteredData length: ${filteredData.length}`);
                currentPageSpan.textContent = currentPage.toString();
                totalPagesSpan.textContent = totalPages.toString();
                if (filteredData.length === 0) {
                    startItemSpan.textContent = '0';
                    endItemSpan.textContent = '0';
                    totalItemsSpan.textContent = '0';
                }
                else {
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
            const goToPage = (page) => {
                console.log(`📊 goToPage called: ${page}, totalPages: ${totalPages}`);
                if (page < 1 || page > totalPages || page === currentPage)
                    return;
                if (filteredData.length === 0)
                    return;
                currentPage = page;
                updatePaginationControls();
                renderCurrentPage();
                const tableContainer = document.querySelector('.table-container-rev');
                if (tableContainer) {
                    tableContainer.scrollTop = 0;
                }
            };
            // COMPLETE REPLACEMENT FOR initializePagination
            const initializePagination = (data) => {
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
            const calculateTotalPages = () => {
                totalPages = Math.ceil(filteredData.length / pageSize);
                if (totalPages === 0)
                    totalPages = 1;
                console.log(`📄 Total pages: ${totalPages} (${filteredData.length} items / ${pageSize} per page)`);
            };
            // COMPLETE REPLACEMENT FOR updatePageNumbers
            const updatePageNumbers = () => {
                pageNumbersDiv.innerHTML = '';
                if (totalPages <= 1) {
                    // Only show page 1 if there's data
                    if (filteredData.length > 0) {
                        const pageBtn = document.createElement('button');
                        pageBtn.className = `page-number active`;
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
                    dotsBtn.className = 'page-number dots';
                    dotsBtn.textContent = '...';
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
                for (let i = startPage; i <= endPage; i++) {
                    const pageBtn = document.createElement('button');
                    pageBtn.className = `page-number ${i === currentPage ? 'active' : ''}`;
                    pageBtn.textContent = i.toString();
                    pageBtn.addEventListener('click', () => goToPage(i));
                    pageNumbersDiv.appendChild(pageBtn);
                }
                if (endPage < totalPages) {
                    const dotsBtn = document.createElement('button');
                    dotsBtn.className = 'page-number dots';
                    dotsBtn.textContent = '...';
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
            };
            const calculateTotals = (data, displayMonths, currentMonthKey) => {
                const totals = {
                    obl: 0,
                    monthly: {}
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
            const setFormReadOnly = (isFrozen) => {
                // Define which selectors should be forced readonly when frozen
                const forcedReadonlySelectors = [
                    '.simple-month-rev .main-current-input',
                    '.yearly-current-input-rev',
                    '.detailed-month-box .item-current-input',
                    '.detailed-month-box .item-description-input',
                    '#remarksInternal',
                    '#remarksSprintCell'
                ];
                const inputs = document.querySelectorAll(forcedReadonlySelectors.join(','));
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
                // ============================================================
                // MAIN-FORECAST-INPUT - ALWAYS READONLY (calculated from items)
                // ============================================================
                const mainForecastInputs = document.querySelectorAll('.main-forecast-input');
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
                    }
                    else {
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
                const itemForecastInputs = document.querySelectorAll('.item-forecast-input');
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
                const dateInputs = document.querySelectorAll('.forecast-date-input, .item-date-input');
                dateInputs.forEach(input => {
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
                // ⚠️ DO NOT HIDE OR DISABLE THE SAVE BUTTON HERE
                // The save button visibility is controlled in editForecast
                // ============================================================
                // Handle OBL input separately
                const oblInput = document.querySelector('#oblLastMonth');
                if (oblInput) {
                    const isExistingProject = editingId !== null;
                    if (isExistingProject || isFrozen) {
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
                // Update the freeze checkbox state with toast
                // Toast is now shown in editForecast to avoid duplication
            };
            const renderTotalsHeader = (displayMonths, currentMonthKey) => {
                const totalsHeader = document.querySelector("#totalsHeader");
                if (!totalsHeader)
                    return;
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
                const buildTotalsRow = (label, totals, css, showData = true) => {
                    if (!showData)
                        return '';
                    // Build month cells
                    const monthCells = displayMonths.map((m, index) => {
                        const isCurrent = m === currentMonthKey;
                        const t = totals.monthly[m] || { plan: 0, current: 0, actual: 0, forecast: 0, total: 0 };
                        if (isCurrent) {
                            // ✅ Current month: Prev. Plan, Curr. Plan, Actual, Further, Forecast
                            // Forecast = Actual + Further (calculated in calculateTotals)
                            return `
          <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle; background-color: #e8f5e9;">${t.actual.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.forecast.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col total-value" style="min-width:60px; text-align:center; vertical-align:middle; font-weight:bold; background-color: #f0f8ff;">${t.total.toLocaleString('en-IN')}</td>
        `;
                        }
                        else {
                            return `
          <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-IN')}</td>
        `;
                        }
                    }).join('');
                    return `
      <tr class="${css}">
        <td class="freeze-col" style="text-align:center; vertical-align:middle; min-width:40px; max-width:40px; width:40px;">${css === 'all-projects-total' ? '★' : '●'}</td>
        <td class="project-ref-col" style="text-align:center; vertical-align:middle; min-width:120px; max-width:120px; width:120px;"></td>
        <td class="project-col" style="text-align:left; vertical-align:middle; min-width:180px; max-width:180px; width:180px; font-weight:bold; background-color:${css === 'all-projects-total' ? '#d1ecf1' : '#d4edda'};">
          ${label}
        </td>
        <td class="pm-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
        <td class="gh-col" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
        <td class="customer-col" style="text-align:center; vertical-align:middle; min-width:130px; max-width:130px; width:130px;"></td>
        <td class="obl-col" style="text-align:right; vertical-align:middle; min-width:90px; max-width:90px; width:90px; font-weight:bold;">${totals.obl.toLocaleString('en-IN')}</td>
        ${monthCells}
      </tr>
    `;
                };
                // Build the totals header
                let html = `
    <tr>
      <th colspan="${totalColspan}" style="background:#f8f9fa; padding:8px; text-align:center; font-size:13px; font-weight:bold; color:#495057; border:1px solid #dee2e6;">
        📊 Collection Totals
      </th>
    </tr>
  `;
                // Always show ALL totals
                if (allData.length > 0) {
                    html += buildTotalsRow(`All Projects (${allData.length})`, allTotals, 'all-projects-total', true);
                }
                // Only show FILTERED totals if filters are applied AND data differs
                if (hasFilters && filtered.length > 0) {
                    html += buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total', true);
                }
                totalsHeader.innerHTML = html;
            };
            // Move attachEditTriggerListeners BEFORE renderTableWithMonthColumns
            const attachEditTriggerListeners = () => {
                document.querySelectorAll('.edit-trigger').forEach(link => {
                    link.addEventListener('click', (e) => {
                        e.preventDefault();
                        const id = parseInt(e.currentTarget.dataset.id);
                        editForecast(id);
                    });
                });
            };
            // COMPLETE REPLACEMENT FOR renderTableWithMonthColumns
            const renderTableWithMonthColumns = (data) => {
                console.log("📊 renderTableWithMonthColumns called with data length:", data.length);
                if (!data || data.length === 0) {
                    const displayMonths = getDisplayMonths();
                    const currentMonthKey = formatMonthKey(new Date());
                    const totalCols = 7 + (displayMonths.length * 2) + (displayMonths.includes(currentMonthKey) ? 3 : 0);
                    tableBody.innerHTML = `<tr><td colspan="${totalCols}" style="text-align:center;padding:20px;">No records found</td></tr>`;
                    const totalsHeader = document.querySelector("#totalsHeader");
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    return;
                }
                // FILTER OUT EMPTY/INVALID DATA - Skip records with no project title
                const validData = data.filter(forecast => {
                    var _a;
                    const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
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
                    const totalsHeader = document.querySelector("#totalsHeader");
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    return;
                }
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                renderTotalsHeader(displayMonths, currentMonthKey);
                const showProjectManagerColumn = true;
                const showGroupHeadColumn = true;
                // Main header
                const mainHeaderHTML = `
    <tr>  
      <th rowspan="2" class="freeze-col" style="text-align:center; min-width:40px; max-width:40px; width:40px;">#</th>
      <th rowspan="2" class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">Project Reference</th>
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
      <th rowspan="2" class="pm-col" style="min-width:100px; max-width:100px; width:100px;">Project Manager</th>
      <th rowspan="2" class="gh-col" style="min-width:100px; max-width:100px; width:100px;">Group Head</th>
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
      <th rowspan="2" class="obl-col" style="min-width:90px; max-width:90px; width:90px;">
        <div class="sort-header-container">
          <span class="header-label">OBL</span>
          <div class="sort-buttons-container">
            <button class="sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                    data-sort="obl" data-direction="asc" title="Sort Ascending">▲</button>
            <button class="sort-btn ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
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
          <th colspan="${colspan}" class="month-scroll-col month-column-header ${isCurrentMonth ? 'current-month-header' : ''}" style="min-width:${isCurrentMonth ? '110px' : '70px'};">
            <div class="sort-header-container">
              <span class="header-label">${monthLabel}${isCurrentDisplay}</span>
              <div class="sort-buttons-container">
                <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                        data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort Ascending">▲</button>
                <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
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
            <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:60px;">
              <div class="sort-header-container">
                <span>Prev. Plan</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:60px;">
              <div class="sort-header-container">
                <span>Curr. Plan</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="current" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:60px;">
              <div class="sort-header-container">
                <span>Actual</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'actual' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="actual" data-month-index="${monthIndex}">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'actual' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="actual" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:60px;">
              <div class="sort-header-container">
                <span>Further</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'forecast' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="forecast" data-month-index="${monthIndex}">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'forecast' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="forecast" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="month-scroll-col month-subheader current-month-header-table" style="min-width:60px;">
              <div class="sort-header-container">
                <span>Forecast</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'total' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="total" data-month-index="${monthIndex}">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'total' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="total" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
          `;
                    }
                    else {
                        return `
            <th class="month-scroll-col month-subheader" style="min-width:60px;">
              <div class="sort-header-container">
                <span>Prev. Plan</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
                          data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}">▼</button>
                </div>
              </div>
            </th>
            <th class="month-scroll-col month-subheader" style="min-width:60px;">
              <div class="sort-header-container">
                <span>Curr. Plan</span>
                <div class="sort-buttons-container">
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'sort-asc active' : ''}" 
                          data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}">▲</button>
                  <button class="sort-btn ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'sort-desc active' : ''}" 
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
                    var _a;
                    const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "Unknown Project";
                    const proj = projects.find(p => p.Title === projectTitle);
                    const forecastData = safeJsonParse(forecast.ForecastData, {});
                    const revenueActualData = safeJsonParse(forecast.RevenueActual, {});
                    const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-IN');
                    const isFrozen = forecast.Freeze || false;
                    const rowStyle = isFrozen ? 'background-color: #fff8e1;' : '';
                    // Check if user can edit this project
                    const canUserEdit = canEditProject(proj, forecast);
                    const canUserFreeze = canFreeze();
                    // Determine if checkbox should be disabled
                    const isCheckboxDisabled = !canUserFreeze;
                    // Build project reference link - only clickable if user can edit
                    const projectRef = getSafeString(proj === null || proj === void 0 ? void 0 : proj.ProjectReference);
                    const projectRefLink = projectRef !== "-"
                        ? (canUserEdit
                            ? `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectRef}</a>`
                            : `<span style="color:#6c757d; cursor:default;">${projectRef}</span>`)
                        : '<span style="color:#999;">-</span>';
                    // Build project name link - only clickable if user can edit
                    const projectNameLink = canUserEdit && projectTitle !== "Unknown Project"
                        ? `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectTitle}</a>`
                        : `<span style="color:${projectTitle === 'Unknown Project' ? '#999' : '#6c757d'}; cursor:default;">${projectTitle}</span>`;
                    // Add freeze badge if frozen
                    const freezeBadge = isFrozen ? ' 🔒' : '';
                    let rowHTML = `
      <tr style="${rowStyle}" data-forecast-id="${forecast.Id}"> 
        <td class="freeze-col" style="text-align:center; vertical-align:middle; min-width:40px; max-width:40px; width:40px;">
          <input type="checkbox" class="freeze-checkbox" 
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
        <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
        <td class="project-col" style="min-width:180px; max-width:180px; width:180px;">${projectNameLink}${freezeBadge}</td>
    `;
                    if (showProjectManagerColumn) {
                        rowHTML += `<td class="pm-col" style="min-width:100px; max-width:100px; width:100px;">${proj ? getSafeString(proj.ProjectManager) : "-"}</td>`;
                    }
                    if (showGroupHeadColumn) {
                        rowHTML += `<td class="gh-col" style="min-width:100px; max-width:100px; width:100px;">${proj ? getSafeString(proj.GroupHead) : "-"}</td>`;
                    }
                    rowHTML += `
        <td class="customer-col" style="min-width:130px; max-width:130px; width:130px;">${getSafeString(proj === null || proj === void 0 ? void 0 : proj.CustomerName)}</td>
        <td class="obl-col" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
    `;
                    displayMonths.forEach((monthKey, monthIndex) => {
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0, forecast: 0 };
                        const isCurrentMonth = monthKey === currentMonthKey;
                        const planValue = monthData.plan || 0;
                        const currentValue = monthData.actual || 0;
                        // const forecastValue = monthData.forecast || 0;
                        // const actualValue = revenueActualData[monthKey] || 0;
                        if (isCurrentMonth) {
                            // ✅ Forecast = RevenueActual + Further
                            const revenueActualValue = revenueActualData[monthKey] || 0; // From RevenueActual field
                            const furtherValue = monthData.forecast || 0; // Further field
                            const totalForecast = revenueActualValue + furtherValue; // Actual + Further
                            rowHTML += `
          <td class="month-scroll-col" style="min-width:60px;">${planValue.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px;">${currentValue.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px; background-color: #e8f5e9;">${revenueActualValue.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px;">${furtherValue.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col total-cell" style="min-width:60px; font-weight: bold; background-color: #f0f8ff;">${totalForecast.toLocaleString('en-IN')}</td>
        `;
                        }
                        else {
                            rowHTML += `
          <td class="month-scroll-col" style="min-width:60px;">${planValue.toLocaleString('en-IN')}</td>
          <td class="month-scroll-col" style="min-width:60px;">${currentValue.toLocaleString('en-IN')}</td>
        `;
                        }
                    });
                    tableBodyHTML += rowHTML + `</tr>`;
                });
                tableBody.innerHTML = tableBodyHTML;
                attachSortButtonListeners();
                attachFreezeEventListeners();
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
                            handleRevenueSort(sortColumn, direction, subColumn, monthIndex);
                        }
                    });
                });
            };
            const attachFreezeEventListeners = () => {
                console.log("🔧 Attaching freeze event listeners...");
                const checkboxes = document.querySelectorAll('.freeze-checkbox');
                const canUserFreeze = canFreeze();
                checkboxes.forEach(checkbox => {
                    const cb = checkbox;
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
                            const target = e.target;
                            const id = parseInt(target.dataset.id);
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
                                }
                                else {
                                    const errorText = await res.text();
                                    console.error("❌ Error response:", errorText);
                                    toast("error", "Failed to update freeze status");
                                    target.checked = !isChecked;
                                }
                            }
                            catch (error) {
                                console.error("Error updating freeze status:", error);
                                toast("error", "Failed to update freeze status");
                                target.checked = !isChecked;
                            }
                        });
                    }
                });
            };
            const getListEntityType = async () => {
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                        headers: { Accept: "application/json;odata=verbose" },
                    });
                    const data = await res.json();
                    return data.d.ListItemEntityTypeFullName;
                }
                catch (error) {
                    console.error("Error getting entity type:", error);
                    // Fallback to a common type name
                    return "SP.Data.Revenue_x0020_Forecast_x0020_ListListItem";
                }
            };
            const handleRevenueSort = (column, direction, subColumn, monthIndex) => {
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
                            const monthDataA = dataA[monthKeyToUse] || { plan: 0, actual: 0, forecast: 0 };
                            const monthDataB = dataB[monthKeyToUse] || { plan: 0, actual: 0, forecast: 0 };
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
                    const subColNames = {
                        'plan': 'Prev. Plan',
                        'current': 'Curr. Plan',
                        'forecast': 'Forecast',
                        'total': 'Further'
                    };
                    columnName = subColNames[subColumn] || subColumn;
                }
                else if (column === 'obl') {
                    columnName = 'OBL';
                }
                else if (column === 'project') {
                    columnName = 'Project';
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
            const showForm = (isEditMode = false, project, forecast) => {
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
                // Show/hide save button based on permissions
                const saveBtn = document.querySelector('.btn-primary[type="submit"]');
                if (saveBtn) {
                    const canSave = isEditMode ? canEditProject(project, forecast) : canCreateForecast();
                    saveBtn.style.display = canSave ? '' : 'none';
                }
                if (isEditMode) {
                    formSection === null || formSection === void 0 ? void 0 : formSection.classList.add('edit-mode');
                    if (projectSelect) {
                        projectSelect.disabled = true;
                    }
                    if (formTitle) {
                        formTitle.textContent = "Edit Forecast";
                    }
                }
                else {
                    formSection === null || formSection === void 0 ? void 0 : formSection.classList.remove('edit-mode');
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
            const hideForm = () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                formSection.classList.remove('edit-mode');
                projectSelect.disabled = false;
            };
            const resetForm = () => {
                editingId = null;
                form.reset();
                projectSelect.value = "";
                projectDetails.innerHTML = "";
                historicalContainer.innerHTML = "";
                firstRowContainer.innerHTML = "";
                remainingContainer.innerHTML = "";
                yearlyContainer.innerHTML = "";
                if (remarksInternal) {
                    remarksInternal.value = "";
                }
                if (remarksSprintCell) {
                    remarksSprintCell.value = "";
                }
                toggleForecastSections(false);
                const oblContainer = document.querySelector('.obl-section-container');
                oblContainer.className = 'obl-section-container hidden-section';
                oblLastMonthInput.removeAttribute('readonly');
                oblLastMonthInput.placeholder = "Enter OBL";
                setNumericValue(oblLastMonthInput, 0);
                const fieldNote = oblContainer.querySelector('.field-note');
                if (fieldNote)
                    fieldNote.textContent = "For new projects: Enter initial OBL value";
                oblValidationMessage.textContent = "✅ All values are within OBL limit";
                oblValidationMessage.className = 'validation-message';
                updateOBLValidationDisplay();
            };
            const blockAllCurrentInputs = () => {
                const currentInputs = document.querySelectorAll('input[data-type="actual"]:not([readonly])');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-rev:not([readonly])');
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
                const currentInputs = document.querySelectorAll('input.item-current-input:not([readonly])');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-rev:not([readonly])');
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
                // Editable Current Plan inputs
                // Long-term Forecast + 5-Year Forecast
                const currentInputs = document.querySelectorAll('input.item-current-input');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-rev');
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
                document.querySelectorAll('.detailed-month-box .main-current-input').forEach(input => {
                    input.setAttribute('readonly', 'true');
                    input.style.cursor = 'default';
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
                return true;
            };
            // const blockIncreasingMonthItemInputs = (monthKey: string, allowedInput?: HTMLInputElement): void => {
            // };
            // const unblockAllMonthItemInputs = (monthKey: string): void => {
            //   const itemInputs = document.querySelectorAll<HTMLInputElement>(`input.item-current-input[data-month="${monthKey}"]`);
            //   itemInputs.forEach(input => {
            //       input.style.backgroundColor = '';
            //       input.style.cursor = '';
            //       input.removeAttribute('data-block-increase');
            //       input.removeAttribute('readonly');
            //   });
            // };
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
                const oblContainer = document.querySelector('.obl-section-container');
                const fieldNote = oblContainer.querySelector('.field-note');
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
            const toggleForecastSections = (show) => {
                const sections = [
                    '.obl-section-container',
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
            const renderProjectDetails = (project) => {
                var _a;
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
                if (projectNameDisplay)
                    projectNameDisplay.textContent = project.Title || "-";
                // Update Project Ref (this should be the SECOND row)  
                const projectRef = document.querySelector('#projectRef');
                if (projectRef)
                    projectRef.textContent = project.ProjectReference || "-";
                // Update Profit Center (third row)
                const profitCenter = document.querySelector('#profitCenter');
                if (profitCenter)
                    profitCenter.textContent = project.PC || "-";
                // Update Project Manager (fourth row)
                const projectManager = document.querySelector('#projectManager');
                if (projectManager) {
                    projectManager.textContent = typeof project.ProjectManager === "object" ?
                        (_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title : project.ProjectManager || "-";
                }
                // Update Group Head (fifth row)
                const groupHead = document.querySelector('#groupHead');
                if (groupHead)
                    groupHead.textContent = project.GroupHead || "-";
                // Update Customer (sixth row)
                const customer = document.querySelector('#customer');
                if (customer)
                    customer.textContent = project.CustomerName || "-";
            };
            const getForecastMonths = () => {
                return generateIdealMonthStructure(); // This should generate 36+ months
            };
            // COMPLETE REPLACEMENT FOR renderForecastInterface
            const renderForecastInterface = (existingData, displayOBL, existingForecast) => {
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
                    }
                    else {
                        // Keep existing data, don't overwrite
                        // Just ensure items exists
                        const monthData = existingData[monthKey];
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
                        const formTitleEl = document.querySelector("#formTitle");
                        if (formTitleEl) {
                            if (isFrozen) {
                                formTitleEl.innerHTML =
                                    '🔒 Edit Forecast (Frozen - Further fields editable)';
                                formTitleEl.style.color = '#856404';
                            }
                            else {
                                formTitleEl.innerHTML = 'Edit Forecast';
                                formTitleEl.style.color = '';
                            }
                        }
                    }, 250);
                }
            };
            const setupHistoricalToggle = () => {
                var _a, _b;
                const toggleBtn = document.querySelector("#toggleHistoricalBtn");
                const historicalContainer = document.querySelector("#historicalContainer");
                const header = document.querySelector(".historical-toggle-header");
                if (!toggleBtn || !historicalContainer || !header)
                    return;
                // Set initial state - container should be HIDDEN by default
                let isVisible = false;
                // ===== BUTTON =====
                // Replace with new button
                const newBtn = document.createElement('button');
                newBtn.id = 'toggleHistoricalBtn';
                newBtn.className = 'btn-secondary';
                newBtn.textContent = 'Show';
                newBtn.type = 'button';
                newBtn.style.cssText = 'padding: 4px 16px; font-size: 13px; background: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; flex-shrink: 0; z-index: 10; position: relative;';
                (_a = toggleBtn.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newBtn, toggleBtn);
                // HIDE the container by default
                historicalContainer.style.display = 'none';
                // ===== HEADER =====
                // Replace header to remove all old listeners
                const newHeader = header.cloneNode(true);
                (_b = header.parentNode) === null || _b === void 0 ? void 0 : _b.replaceChild(newHeader, header);
                // ===== BUTTON CLICK (direct, no propagation issues) =====
                const btn = document.querySelector("#toggleHistoricalBtn");
                const container = document.querySelector("#historicalContainer");
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
                // ===== HEADER CLICK (only for the header, not the button) =====
                const headerEl = document.querySelector(".historical-toggle-header");
                if (headerEl) {
                    headerEl.addEventListener("click", function (e) {
                        // If the click is on the button or its children, ignore
                        if (e.target.closest("#toggleHistoricalBtn")) {
                            return;
                        }
                        e.preventDefault();
                        e.stopPropagation();
                        const btn = document.querySelector("#toggleHistoricalBtn");
                        const container = document.querySelector("#historicalContainer");
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
            const protectHistoricalElements = () => {
                document.querySelectorAll('.historical-month input, .historical-month select, .historical-month button').forEach(element => {
                    var _a;
                    const newElement = element.cloneNode(true);
                    (_a = element.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newElement, element);
                });
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
          <div class="month-box-rev historical-month">
            <div class="month-header-rev">${monthLabel} (Historical - Readonly)</div>
            
            <div class="month-block">
              <div class="main-inputs-container-rev">
                <div class="main-input-group-rev">
                  <label>Prev. Plan</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input ${planInputClass} number-input" />
                </div>
                <div class="main-input-group-rev">
                  <label>Curr. Plan</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        readonly class="main-current-input ${actualInputClass} number-input" />
                </div>
              </div>
            </div>

            <div class="items-section historical-items-section">
              <div class="items-header historical-items-header">
                <div>Item Description</div>
                <div>Prev. Plan</div>
                <div>Curr. Plan</div>
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
                   <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
      data-item="${itemNum}" data-month="${monthKey}" 
      class="${itemNameClass}" style="pointer-events: auto !important; cursor: text !important;" />
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
            const renderFirstRow = (months, existingData, existingForecast) => {
                let html = '';
                html += `<div class="detailed-months-row">`;
                // Get current month key
                const now = new Date();
                const currentMonthKey = formatMonthKey(now);
                // Find current month in the months array
                const currentMonthIndex = months.indexOf(currentMonthKey);
                // Detailed months: current month and next 2 months
                let detailedMonths = [];
                if (currentMonthIndex !== -1 && currentMonthIndex + 3 <= months.length) {
                    detailedMonths = months.slice(currentMonthIndex, currentMonthIndex + 3);
                }
                else if (currentMonthIndex !== -1) {
                    detailedMonths = months.slice(currentMonthIndex);
                }
                else {
                    detailedMonths = months.slice(0, 3);
                }
                console.log(`📊 Detailed months for rendering: ${detailedMonths.join(', ')}`);
                // ✅ Load RevenueActual data from separate column
                let revenueActualData = {};
                if (existingForecast) {
                    revenueActualData = safeJsonParse(existingForecast.RevenueActual, {});
                    console.log("📊 RevenueActual data:", revenueActualData);
                }
                // Current month is the first of the detailed months
                const currentMonth = detailedMonths[0];
                if (currentMonth) {
                    const monthDate = new Date(currentMonth + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[currentMonth] || {
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
                    html += `<div class="current-month-container">`;
                    html += renderDetailedMonthBox(currentMonth, monthLabel, data, true);
                    html += `</div>`;
                }
                // Next 2 months - also get monthActual from RevenueActual if available
                html += `<div class="next-months-container">`;
                for (let i = 1; i < 3 && i < detailedMonths.length; i++) {
                    const monthKey = detailedMonths[i];
                    const monthDate = new Date(monthKey + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[monthKey] || {
                        plan: 0,
                        actual: 0,
                        monthActual: 0,
                        items: {}
                    };
                    // ✅ Set monthActual from RevenueActual column for next months
                    if (revenueActualData[monthKey] !== undefined) {
                        data.monthActual = revenueActualData[monthKey];
                    }
                    html += `<div class="next-month-box">`;
                    html += renderDetailedMonthBox(monthKey, monthLabel, data, false);
                    html += `</div>`;
                }
                html += `</div>`;
                html += `</div>`;
                firstRowContainer.innerHTML = html;
            };
            const renderRemainingMonths = (months, existingData) => {
                let html = '';
                for (let row = 0; row < 3; row++) {
                    html += `<div class="remaining-row-rev">`;
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
          <div class="yearly-box-rev">
            <div class="yearly-header-rev">${year}</div>
            
            <div class="yearly-inputs-container-rev">
              <div class="yearly-input-group-rev">
                <label>Prev. Plan</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                      readonly class="yearly-plan-input-rev ${planInputClass} number-input" />
              </div>
              <div class="yearly-input-group-rev">
                <label>Curr. Plan</label>
                <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                      class="yearly-current-input-rev ${actualInputClass} number-input" />
              </div>
            </div>
          </div>
        `;
            };
            const renderDetailedMonthBox = (monthKey, monthLabel, data, isCurrentMonth = false) => {
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
                const planInputClass = data.plan > 0 ? 'input-with-value' : 'input-empty';
                const actualInputClass = data.actual > 0 ? 'input-with-value' : 'input-empty';
                const forecastInputClass = displayForecastValue > 0 ? 'input-with-value' : 'input-empty';
                const dateInputClass = displayDateValue ? 'input-with-value' : 'input-empty';
                const items = data.items || createEmptyItems();
                // Month Actual badge
                let monthActualBadge = '';
                if (isCurrentMonth) {
                    const hasMonthActual = displayMonthActual > 0;
                    monthActualBadge = `
            <span class="month-actual-badge ${hasMonthActual ? 'has-value' : ''}" style="
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
                <span>📊 Month Actual:</span>
                <span style="font-size: 16px; font-weight: 700;">${displayMonthActual.toLocaleString()}</span>
            </span>
        `;
                }
                return `
        <div class="detailed-month-box ${isCurrentMonth ? 'current-month-box' : 'next-month-box'}">
            <div class="month-header-rev" style="
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 8px 12px;
                background: #f8f9fa;
                border-bottom: 2px solid #007bff;
                border-radius: 4px 4px 0 0;
            ">
                <span style="font-size: 15px; font-weight: 700; color: #333;">
                    ${monthLabel}
                    ${isCurrentMonth ? '<span class="current-month-indicator" style="margin-left: 8px; font-size: 12px; color: #007bff; background: #e3f2fd; padding: 2px 10px; border-radius: 12px;">Current</span>' : ''}
                </span>
                ${monthActualBadge}
            </div>
            
            <div class="month-block" style="padding: 8px 12px;">
                <div class="enhanced-inputs-container" style="
                    display: grid;
                    grid-template-columns: ${isCurrentMonth ? '1fr 1fr 1fr 1fr' : '1fr 1fr'};
                    gap: 8px;
                ">
                    <div class="enhanced-input-group">
                        <label style="font-size: 11px; font-weight: 600; color: #666;">Prev. Plan</label>
                        <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                              readonly class="main-plan-input ${planInputClass} number-input" style="width: 100%; padding: 4px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center;" />
                    </div>
                    
                    <div class="enhanced-input-group">
                        <label style="font-size: 11px; font-weight: 600; color: #666;">Curr. Plan</label>
                        <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                              readonly class="main-current-input ${actualInputClass} number-input" style="width: 100%; padding: 4px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center;" data-calc-total="true" />
                    </div>
                    
                    ${isCurrentMonth ? `
                       <div class="enhanced-input-group">
        <label style="font-size: 11px; font-weight: 600; color: #666;">Further</label>
        <input type="text" value="${displayForecastValue}" data-type="forecast" data-month="${monthKey}"
              readonly class="main-forecast-input ${forecastInputClass} number-input" style="width: 100%; padding: 4px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center; background-color: #f5f5f5;" data-calc-total="true" />
    </div>
                        
                        <div class="enhanced-input-group">
                            <label style="font-size: 11px; font-weight: 600; color: #666;">Date</label>
                            <input type="date" value="${displayDateValue}" data-type="date" data-month="${monthKey}"
                                  class="forecast-date-input ${dateInputClass}" style="width: 100%; padding: 4px 6px; border: 1px solid #28a745; border-radius: 4px; background-color: #ffffff;" />
                        </div>
                    ` : ''}
                </div>
            </div>

            <div class="items-section" style="padding: 0 12px 12px 12px;">
                <div class="items-header" style="
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
                    <div>Item Description</div>
                    <div>Prev. Plan</div>
                    <div>Curr. Plan</div>
                    ${isCurrentMonth ? `
                        <div>Further</div>
                        <div>Date</div>
                    ` : ''}
                </div>
                
                ${Array.from({ length: 10 }, (_, i) => i + 1).map(itemNum => {
                    const item = items[itemNum] || { name: '', plan: 0, actual: 0, forecast: 0, date: '' };
                    const itemPlanValue = item.plan > 0 ? item.plan.toString() : '';
                    const itemActualValue = item.actual > 0 ? item.actual.toString() : '';
                    const itemForecastValue = (item.forecast || 0) > 0 ? (item.forecast || 0).toString() : '';
                    const itemDateValue = item.date || '';
                    const itemPlanClass = item.plan > 0 ? 'input-with-value' : 'input-empty';
                    const itemActualClass = item.actual > 0 ? 'input-with-value' : 'input-empty';
                    const itemForecastClass = (item.forecast || 0) > 0 ? 'input-with-value' : 'input-empty';
                    const itemDateClass = itemDateValue ? 'input-with-value' : 'input-empty';
                    const itemNameClass = item.name ? 'input-with-value' : 'input-empty';
                    return `
                        <div class="item-row ${isCurrentMonth ? 'revenue-row' : ''}" style="
                            display: grid;
                            grid-template-columns: ${isCurrentMonth ? '1fr 0.8fr 0.8fr 0.8fr 0.8fr' : '1fr 0.8fr 0.8fr'};
                            gap: 4px;
                            padding: 3px 4px;
                            border-bottom: 1px solid #f0f0f0;
                            align-items: center;
                        ">
                            <input type="text" placeholder="Item ${itemNum}" value="${item.name}" 
                                  data-item="${itemNum}" data-month="${monthKey}" 
                                  class="item-description-input ${itemNameClass}" 
                                  style="width: 100%; padding: 3px 6px; border: 1px solid #ddd; border-radius: 4px; font-size: 11px; pointer-events: auto !important; cursor: text !important; background-color: #ffffff !important;" />
                            <input type="text" value="${itemPlanValue}" 
                                  data-item-type="plan" data-item="${itemNum}" data-month="${monthKey}" 
                                  readonly class="item-plan-input ${itemPlanClass} number-input" style="width: 100%; padding: 3px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center; font-size: 11px;" />
                            
                            <input type="text" value="${itemActualValue}" 
                                  data-item-type="actual" data-item="${itemNum}" data-month="${monthKey}" 
                                  class="item-current-input ${itemActualClass} number-input" style="width: 100%; padding: 3px 6px; border: 1px solid #ddd; border-radius: 4px; text-align: center; font-size: 11px;" data-calc-total="true" />
                            
                            ${isCurrentMonth ? `
                                <input type="text" value="${itemForecastValue}" 
                                      data-item-type="forecast" data-item="${itemNum}" data-month="${monthKey}" 
                                      class="item-forecast-input ${itemForecastClass} number-input" style="width: 100%; padding: 3px 6px; border: 1px solid #28a745; border-radius: 4px; text-align: center; font-size: 11px; background-color: #ffffff;" data-calc-total="true" />
                                
                                <input type="date" value="${itemDateValue}" 
                                      data-item-type="date" data-item="${itemNum}" data-month="${monthKey}" 
                                      class="item-date-input ${itemDateClass}" style="width: 100%; padding: 3px 6px; border: 1px solid #28a745; border-radius: 4px; font-size: 11px; background-color: #ffffff;" />
                            ` : ''}
                        </div>
                    `;
                }).join('')}
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
          <div class="month-box-rev simple-month-rev">
            <div class="month-header-rev">${monthLabel}</div>
            
            <div class="main-inputs-container-rev">
              <div class="main-input-group-rev">
                <label>Prev. Plan</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}" 
                      readonly class="main-plan-input ${planInputClass} number-input" />
              </div>
              <div class="main-input-group-rev">
                <label>Curr. Plan</label>
                <input type="text" value="${displayActualValue}"  id="actualCurrent" data-type="actual" data-month="${monthKey}" 
                      class="main-current-input ${actualInputClass} number-input" />
              </div>
            </div>
          </div>
        `;
            };
            const attachForecastEventListeners = () => {
                document.querySelectorAll('.number-input').forEach(input => {
                    allowOnlyNumbers(input);
                });
                document.querySelectorAll('.number-input:not([readonly])').forEach(input => {
                    input.addEventListener('focus', function () {
                        this.setAttribute('data-old-value', getNumericValue(this).toString());
                    });
                });
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
                        if (currentValue === 0) {
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.main-forecast-input:not([readonly])').forEach(input => {
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
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
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
                        if (currentValue === 0) {
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.item-forecast-input:not([readonly])').forEach(input => {
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
                        if (currentValue === 0) {
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.yearly-current-input-rev:not([readonly])').forEach(input => {
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
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.forecast-date-input:not([readonly])').forEach(input => {
                    input.addEventListener('change', function () {
                        if (this.value === '') {
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                    });
                });
                document.querySelectorAll('.item-date-input:not([readonly])').forEach(input => {
                    input.addEventListener('change', function () {
                        if (this.value === '') {
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                    });
                });
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
                document.querySelectorAll('.main-forecast-input').forEach(input => {
                    // Ensure it's always editable
                    input.removeAttribute('readonly');
                    input.style.backgroundColor = '#ffffff';
                    input.style.pointerEvents = 'auto';
                    input.disabled = false;
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
                            this.className = this.className.replace('input-with-value', 'input-empty');
                        }
                        else {
                            this.className = this.className.replace('input-empty', 'input-with-value');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                // OBL input listener
                oblLastMonthInput.addEventListener('input', function () {
                    // Remove existing toast
                    const existingToast = document.querySelector('.obl-error-toast');
                    if (existingToast)
                        existingToast.remove();
                    // Update validation display
                    updateOBLValidationDisplay(this);
                    // Show appropriate toast based on current state
                    const oblValue = getNumericValue(oblLastMonthInput);
                    const totalCurrent = calculateTotalCurrentValue();
                    if (oblValue === 0 && totalCurrent === 0) {
                        showOBLToast('info', `ℹ️ OBL is zero. No budget limit applied.`);
                    }
                    else if (oblValue > 0 && totalCurrent <= oblValue) {
                        const remaining = oblValue - totalCurrent;
                        const percentage = Math.round((totalCurrent / oblValue) * 100);
                        if (totalCurrent === oblValue) {
                            showOBLToast('success', `✅ OBL FULLY UTILIZED: ${totalCurrent.toLocaleString()} = ${oblValue.toLocaleString()}`);
                        }
                        else if (percentage > 90) {
                            showOBLToast('warning', `⚠️ OBL nearly reached: ${percentage}% used. ${remaining.toLocaleString()} remaining.`);
                        }
                        else {
                            showOBLToast('success', `✅ OBL within limit: ${percentage}% used. ${remaining.toLocaleString()} remaining.`);
                        }
                    }
                    else if (oblValue > 0 && totalCurrent > oblValue) {
                        const exceedAmount = totalCurrent - oblValue;
                        showOBLToast('danger', `❌ OBL EXCEEDED by ${exceedAmount.toLocaleString()}! Total: ${totalCurrent.toLocaleString()} > OBL: ${oblValue.toLocaleString()}`);
                    }
                });
                // ===== GLOBAL INPUT LISTENER FOR ALL CURRENT INPUTS - FIXED =====
                // Use event delegation for better performance
                document.addEventListener('input', function (e) {
                    const target = e.target;
                    // Check if the input is one we care about
                    if (target.matches('.item-current-input, .main-current-input, .yearly-current-input-rev, #oblLastMonth')) {
                        // Small delay to let the value update
                        setTimeout(() => {
                            // Remove existing toast
                            const existingToast = document.querySelector('.obl-error-toast');
                            if (existingToast)
                                existingToast.remove();
                            // Calculate values
                            const oblValue = getNumericValue(oblLastMonthInput);
                            const totalCurrent = calculateTotalCurrentValue();
                            // Update display
                            totalCurrentValueSpan.textContent = totalCurrent.toLocaleString();
                            const remaining = oblValue - totalCurrent;
                            remainingValueSpan.textContent = remaining.toLocaleString();
                            // Update line fill
                            if (oblValue > 0) {
                                const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
                                oblLineFill.style.width = `${percentage}%`;
                                oblLineFill.className = 'line-fill';
                                if (percentage > 90)
                                    oblLineFill.classList.add('danger');
                                else if (percentage > 75)
                                    oblLineFill.classList.add('warning');
                            }
                            else {
                                oblLineFill.style.width = '0%';
                                oblLineFill.className = 'line-fill';
                            }
                            // Show toast based on state
                            if (oblValue === 0) {
                                if (totalCurrent > 0) {
                                    showOBLToast('warning', `⚠️ OBL is zero but values exist (${totalCurrent.toLocaleString()})`);
                                }
                                else {
                                    showOBLToast('info', `ℹ️ OBL is zero. No budget limit.`);
                                }
                            }
                            else if (totalCurrent > oblValue) {
                                showOBLToast('danger', `❌ OBL EXCEEDED by ${(totalCurrent - oblValue).toLocaleString()}! Total: ${totalCurrent.toLocaleString()} > OBL: ${oblValue.toLocaleString()}`);
                            }
                            else if (totalCurrent === oblValue) {
                                showOBLToast('success', `✅ OBL FULLY UTILIZED: ${totalCurrent.toLocaleString()} = ${oblValue.toLocaleString()}`);
                            }
                            else {
                                const pct = Math.round((totalCurrent / oblValue) * 100);
                                const remainingAmt = oblValue - totalCurrent;
                                if (pct > 90) {
                                    showOBLToast('warning', `⚠️ OBL nearly reached: ${pct}% used. ${remainingAmt.toLocaleString()} remaining.`);
                                }
                                else if (pct > 75) {
                                    showOBLToast('info', `📊 OBL usage: ${pct}%. ${remainingAmt.toLocaleString()} remaining.`);
                                }
                                else {
                                    showOBLToast('success', `✅ OBL within limit: ${pct}% used. ${remainingAmt.toLocaleString()} remaining.`);
                                }
                            }
                        }, 50);
                    }
                }, true);
                // Also trigger validation when any input changes
                document.querySelectorAll('.item-current-input:not([readonly]), .main-current-input:not([readonly]), .yearly-current-input-rev:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
                        // Small delay to allow value to be set
                        setTimeout(() => {
                            updateOBLValidationDisplay(this);
                        }, 50);
                    });
                });
                // ===== OBL REAL-TIME VALIDATION =====
                // Force OBL validation on any input change using event delegation
                // 1. Remove any existing OBL toast and show fresh on any input change
                document.addEventListener('input', function (e) {
                    const target = e.target;
                    // Check if the input is one we care about
                    if (target.matches('.item-current-input, .main-current-input, .yearly-current-input-rev, #oblLastMonth')) {
                        // Small delay to let the value update
                        setTimeout(() => {
                            // Remove existing toast
                            const existingToast = document.querySelector('.obl-error-toast');
                            if (existingToast)
                                existingToast.remove();
                            // Calculate values
                            const oblValue = getNumericValue(oblLastMonthInput);
                            const totalCurrent = calculateTotalCurrentValue();
                            // Update display
                            totalCurrentValueSpan.textContent = totalCurrent.toLocaleString();
                            const remaining = oblValue - totalCurrent;
                            remainingValueSpan.textContent = remaining.toLocaleString();
                            // Update line fill
                            if (oblValue > 0) {
                                const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
                                oblLineFill.style.width = `${percentage}%`;
                                oblLineFill.className = 'line-fill';
                                if (percentage > 90)
                                    oblLineFill.classList.add('danger');
                                else if (percentage > 75)
                                    oblLineFill.classList.add('warning');
                            }
                            else {
                                oblLineFill.style.width = '0%';
                                oblLineFill.className = 'line-fill';
                            }
                            // Show toast based on state
                            if (oblValue === 0) {
                                if (totalCurrent > 0) {
                                    showOBLToast('warning', `⚠️ OBL is zero but values exist (${totalCurrent.toLocaleString()})`);
                                }
                                else {
                                    showOBLToast('info', `ℹ️ OBL is zero. No budget limit.`);
                                }
                            }
                            else if (totalCurrent > oblValue) {
                                showOBLToast('danger', `❌ OBL EXCEEDED by ${(totalCurrent - oblValue).toLocaleString()}! Total: ${totalCurrent.toLocaleString()} > OBL: ${oblValue.toLocaleString()}`);
                            }
                            else if (totalCurrent === oblValue) {
                                showOBLToast('success', `✅ OBL FULLY UTILIZED: ${totalCurrent.toLocaleString()} = ${oblValue.toLocaleString()}`);
                            }
                            else {
                                const pct = Math.round((totalCurrent / oblValue) * 100);
                                const remainingAmt = oblValue - totalCurrent;
                                if (pct > 90) {
                                    showOBLToast('warning', `⚠️ OBL nearly reached: ${pct}% used. ${remainingAmt.toLocaleString()} remaining.`);
                                }
                                else if (pct > 75) {
                                    showOBLToast('info', `📊 OBL usage: ${pct}%. ${remainingAmt.toLocaleString()} remaining.`);
                                }
                                else {
                                    showOBLToast('success', `✅ OBL within limit: ${pct}% used. ${remainingAmt.toLocaleString()} remaining.`);
                                }
                            }
                        }, 50);
                    }
                }, true);
                document.querySelectorAll('.item-current-input:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
                        const monthKey = this.dataset.month;
                        if (monthKey) {
                            validateSingleMonthItems(monthKey, this);
                        }
                    });
                });
                const calculateMonthTotals = (monthKey) => {
                    const itemCurrentInputs = document.querySelectorAll(`input.item-current-input[data-month="${monthKey}"]`);
                    const itemForecastInputs = document.querySelectorAll(`input.item-forecast-input[data-month="${monthKey}"]`);
                    let totalCurrent = 0;
                    let totalForecast = 0;
                    itemCurrentInputs.forEach(input => {
                        totalCurrent += getNumericValue(input);
                    });
                    itemForecastInputs.forEach(input => {
                        totalForecast += getNumericValue(input);
                    });
                    // Update Curr. Plan main input
                    const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                    if (mainCurrentInput) {
                        setNumericValue(mainCurrentInput, totalCurrent);
                        mainCurrentInput.dispatchEvent(new Event('input'));
                    }
                    // ✅ Update main forecast input (READONLY - calculated from items)
                    const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${monthKey}"]`);
                    if (mainForecastInput) {
                        setNumericValue(mainForecastInput, totalForecast);
                        mainForecastInput.dispatchEvent(new Event('input'));
                    }
                };
                document.querySelectorAll('.item-current-input:not([readonly]), .item-forecast-input:not([readonly])').forEach(input => {
                    input.addEventListener('input', function () {
                        const monthKey = this.dataset.month;
                        if (monthKey) {
                            calculateMonthTotals(monthKey);
                            validateSingleMonthItems(monthKey, this);
                            updateOBLValidationDisplay(this);
                        }
                    });
                });
                const validationInputs = document.querySelectorAll(`
            .main-current-input:not([readonly]),
            .item-current-input:not([readonly]),
            .yearly-current-input-rev:not([readonly]),
            .main-forecast-input:not([readonly]),
            .item-forecast-input:not([readonly])
        `);
                validationInputs.forEach(input => {
                    input.addEventListener('input', function () {
                        setTimeout(() => {
                            updateOBLValidationDisplay(this);
                        }, 100);
                    });
                });
            };
            // ============================================================
            // Updated editForecast with RevenueActual loading
            // ============================================================
            const editForecast = async (id) => {
                editingId = id;
                const forecast = forecasts.find(f => f.Id === id);
                if (!forecast) {
                    toast("error", "Forecast not found");
                    return;
                }
                // Get project for permission check
                const project = projects.find(p => { var _a; return p.Id === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Id); });
                // Check if user can view this forecast
                if (!canRead()) {
                    toast("error", "❌ You don't have permission to view this forecast");
                    return;
                }
                // Check if project is assigned to user (for non-admin/non-GH)
                const assignedProjects = getAssignedProjects(projects);
                const isAssigned = assignedProjects.some(p => { var _a; return p.Id === ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Id); });
                const isGHHeadUser = isGHHead();
                const isFullControlUser = hasFullControl();
                if (!isFullControlUser && !isGHHeadUser && !isAssigned) {
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
                    }
                    else {
                        if (remarksInternal) {
                            remarksInternal.value = remarksData;
                        }
                    }
                }
                let forecastData = {};
                let revenueActualData = {};
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
                }
                catch (e) {
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
                    // ✅ SHOW SAVE BUTTON - always show if user can edit the project
                    const saveBtn = document.querySelector('.btn-primary[type="submit"]');
                    if (saveBtn) {
                        // Only hide if user doesn't have edit permission at all
                        saveBtn.style.display = canEdit ? '' : 'none';
                        // Change button text and style based on freeze state
                        if (isFrozen) {
                            saveBtn.textContent = '💾 Save Further Changes';
                            saveBtn.title = 'Save only Further field changes (Curr. Plan is locked)';
                            saveBtn.style.backgroundColor = '#28a745';
                            saveBtn.style.borderColor = '#28a745';
                            saveBtn.style.color = '#ffffff';
                            saveBtn.style.fontWeight = 'bold';
                            saveBtn.style.cursor = 'pointer';
                            saveBtn.setAttribute('type', 'submit');
                        }
                        else {
                            saveBtn.textContent = '💾 Save Revenue Plan';
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
                            formTitleEl.innerHTML = '🔒 Edit Forecast (Frozen - Further fields editable)';
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
                    // ✅ Show/hide OBL section based on freeze state
                    const oblContainer = document.querySelector('.obl-section-container');
                    if (oblContainer) {
                        if (isFrozen) {
                            oblContainer.style.opacity = '0.6';
                            oblContainer.style.pointerEvents = 'none';
                        }
                        else {
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
                    }
                    else if (!canEdit) {
                        showOBLToast('info', '🔒 You have view-only access to this project.');
                    }
                    else {
                        showOBLToast('success', '🔓 All fields are editable. Click "Save Revenue Plan" to save.');
                    }
                }, 300);
                toast("info", `Viewing forecast for ${forecast.Title}`);
            };
            // ============================================================
            // SAVE REVENUE PLAN / SAVE FURTHER CHANGES
            // ============================================================
            // ONE native submit handler for both buttons.
            //
            // Normal project:
            //   "Save Revenue Plan"
            //   -> saves editable Plan / Curr. Plan / Further / Date /
            //      detailed items / yearly data / RevenueActual
            //
            // Frozen project:
            //   "Save Further Changes"
            //   -> saves ONLY Further + Date + item Further + item Date
            //   -> preserves locked Curr. Plan / Actual / Yearly values
            // ============================================================
            form.addEventListener("submit", async (e) => {
                e.preventDefault();
                e.stopPropagation();
                console.log("============================================");
                console.log("💾 REVENUE FORM SUBMIT STARTED");
                console.log("💾 Editing ID:", editingId);
                // ------------------------------------------------------------
                // 1. Get selected project
                // ------------------------------------------------------------
                const projectId = parseInt(projectSelect.value, 10);
                const project = projects.find(p => p.Id === projectId);
                if (!project) {
                    toast("error", "Please select a project");
                    return;
                }
                // ------------------------------------------------------------
                // 2. Get existing forecast
                // ------------------------------------------------------------
                const existingForecast = editingId
                    ? forecasts.find(f => f.Id === editingId)
                    : undefined;
                // ------------------------------------------------------------
                // 3. Permission validation
                // ------------------------------------------------------------
                if (editingId) {
                    if (!canEditProject(project, existingForecast)) {
                        toast("error", "❌ You don't have permission to edit this forecast");
                        return;
                    }
                }
                else {
                    if (!canCreateForecast()) {
                        toast("error", "❌ You don't have permission to create forecasts");
                        return;
                    }
                }
                // ------------------------------------------------------------
                // 4. Determine frozen state
                // ------------------------------------------------------------
                const isFrozen = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.Freeze) === true;
                console.log("🔒 Is Frozen:", isFrozen);
                // ------------------------------------------------------------
                // 5. Confirm frozen save
                // ------------------------------------------------------------
                console.log("🔍 Before frozen confirmation");
                if (editingId && isFrozen) {
                    console.log("🔒 Frozen project - testing direct save");
                    const confirmSave = await Promise.resolve({
                        isConfirmed: true
                    });
                    console.log("✅ Frozen save confirmed:", confirmSave.isConfirmed);
                    if (!confirmSave.isConfirmed) {
                        toast("info", "Save cancelled");
                        return;
                    }
                }
                // ------------------------------------------------------------
                // 6. New/existing project validation
                // ------------------------------------------------------------
                const isNewProject = !editingId &&
                    !forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                if (!isNewProject &&
                    forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId && f.Id !== editingId; })) {
                    toast("error", "Forecast already exists for this project!");
                    return;
                }
                // ------------------------------------------------------------
                // 7. OBL validation
                // ------------------------------------------------------------
                if (!isFrozen) {
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
                }
                // ------------------------------------------------------------
                // 8. Disable Save button while saving
                // ------------------------------------------------------------
                const saveButton = form.querySelector('button[type="submit"]');
                if (saveButton) {
                    saveButton.disabled = true;
                    saveButton.dataset.saving = "true";
                }
                try {
                    // ========================================================
                    // 9. Load existing ForecastData
                    // ========================================================
                    let forecastData = {};
                    let revenueActualData = {};
                    if (existingForecast) {
                        const existingData = safeJsonParse(existingForecast.ForecastData, {});
                        // Deep clone so original SharePoint data is never mutated
                        forecastData = JSON.parse(JSON.stringify(existingData));
                        revenueActualData = safeJsonParse(existingForecast.RevenueActual, {});
                        console.log("📊 Existing RevenueActual:", revenueActualData);
                    }
                    // --------------------------------------------------------
                    // 10. Create fresh structure for new record
                    // --------------------------------------------------------
                    if (Object.keys(forecastData).length === 0) {
                        forecastData = createFreshForecastStructure();
                    }
                    // ========================================================
                    // 11. Get all storage months
                    // ========================================================
                    const allStorageMonths = getAllMonthsForStorage();
                    const currentMonthKey = formatMonthKey(new Date());
                    const currentMonthIndex = allStorageMonths.indexOf(currentMonthKey);
                    let detailedMonths = [];
                    if (currentMonthIndex !== -1) {
                        detailedMonths = allStorageMonths.slice(currentMonthIndex, Math.min(currentMonthIndex + 3, allStorageMonths.length));
                    }
                    else {
                        detailedMonths = allStorageMonths.slice(36, 39);
                    }
                    console.log("📅 Current month:", currentMonthKey);
                    console.log("📅 Detailed months:", detailedMonths);
                    // ========================================================
                    // 12. Collect monthly data from DOM
                    // ========================================================
                    allStorageMonths.forEach((monthKey) => {
                        const existingMonthData = forecastData[monthKey] || {
                            plan: 0,
                            actual: 0,
                            monthActual: 0,
                            forecast: 0,
                            date: "",
                            items: {}
                        };
                        const monthData = {
                            plan: existingMonthData.plan || 0,
                            actual: existingMonthData.actual || 0,
                            monthActual: existingMonthData.monthActual || 0,
                            forecast: existingMonthData.forecast || 0,
                            date: existingMonthData.date || "",
                            items: existingMonthData.items
                                ? JSON.parse(JSON.stringify(existingMonthData.items))
                                : {}
                        };
                        // ====================================================
                        // FROZEN PROJECT
                        // Only Further + Date are changed.
                        // ====================================================
                        if (isFrozen) {
                            console.log(`🔒 Frozen save: ${monthKey}`);
                            const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${monthKey}"]`);
                            const mainDateInput = document.querySelector(`input.forecast-date-input[data-month="${monthKey}"]`);
                            // Main Further / Forecast
                            if (mainForecastInput) {
                                monthData.forecast =
                                    getNumericValue(mainForecastInput);
                            }
                            // Main Date
                            if (mainDateInput) {
                                monthData.date =
                                    mainDateInput.value || monthData.date;
                            }
                            // ----------------------------------------------
                            // Item-level Further + Date
                            // ----------------------------------------------
                            if (!monthData.items) {
                                monthData.items = {};
                            }
                            for (let itemNum = 1; itemNum <= 10; itemNum++) {
                                const itemKey = itemNum.toString();
                                const itemForecastInput = document.querySelector(`.item-forecast-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                const itemDateInput = document.querySelector(`.item-date-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                // If the item already exists, update only
                                // editable Further / Date fields.
                                if (monthData.items[itemKey]) {
                                    if (itemForecastInput) {
                                        monthData.items[itemKey].forecast =
                                            getNumericValue(itemForecastInput);
                                    }
                                    if (itemDateInput) {
                                        monthData.items[itemKey].date =
                                            itemDateInput.value ||
                                                monthData.items[itemKey].date ||
                                                "";
                                    }
                                }
                            }
                            // IMPORTANT:
                            // plan and actual are NOT touched.
                        }
                        // ====================================================
                        // NORMAL PROJECT
                        // ====================================================
                        else {
                            console.log(`🔓 Normal save: ${monthKey}`);
                            const mainPlanInput = document.querySelector(`input.main-plan-input[data-month="${monthKey}"]`);
                            const mainCurrentInput = document.querySelector(`input.main-current-input[data-month="${monthKey}"]`);
                            const mainForecastInput = document.querySelector(`input.main-forecast-input[data-month="${monthKey}"]`);
                            const mainDateInput = document.querySelector(`input.forecast-date-input[data-month="${monthKey}"]`);
                            // ----------------------------------------------
                            // Main monthly values
                            // ----------------------------------------------
                            if (mainPlanInput) {
                                monthData.plan =
                                    getNumericValue(mainPlanInput);
                            }
                            if (mainCurrentInput) {
                                monthData.actual =
                                    getNumericValue(mainCurrentInput);
                            }
                            if (mainForecastInput) {
                                monthData.forecast =
                                    getNumericValue(mainForecastInput);
                            }
                            if (mainDateInput) {
                                monthData.date =
                                    mainDateInput.value || monthData.date;
                            }
                            // ----------------------------------------------
                            // Detailed item data
                            // ----------------------------------------------
                            if (detailedMonths.includes(monthKey)) {
                                if (!monthData.items) {
                                    monthData.items = {};
                                }
                                for (let itemNum = 1; itemNum <= 10; itemNum++) {
                                    const itemKey = itemNum.toString();
                                    const itemNameInput = document.querySelector(`.item-description-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    const itemPlanInput = document.querySelector(`.item-plan-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    const itemCurrentInput = document.querySelector(`.item-current-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    const itemForecastInput = document.querySelector(`.item-forecast-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    const itemDateInput = document.querySelector(`.item-date-input[data-item="${itemKey}"][data-month="${monthKey}"]`);
                                    if (itemNameInput ||
                                        itemPlanInput ||
                                        itemCurrentInput ||
                                        itemForecastInput ||
                                        itemDateInput) {
                                        const existingItem = monthData.items[itemKey] || {
                                            name: "",
                                            plan: 0,
                                            actual: 0,
                                            forecast: 0,
                                            date: ""
                                        };
                                        monthData.items[itemKey] = {
                                            name: itemNameInput
                                                ? itemNameInput.value || ""
                                                : existingItem.name || "",
                                            plan: itemPlanInput
                                                ? getNumericValue(itemPlanInput)
                                                : existingItem.plan || 0,
                                            actual: itemCurrentInput
                                                ? getNumericValue(itemCurrentInput)
                                                : existingItem.actual || 0,
                                            forecast: itemForecastInput
                                                ? getNumericValue(itemForecastInput)
                                                : existingItem.forecast || 0,
                                            date: itemDateInput
                                                ? itemDateInput.value || ""
                                                : existingItem.date || ""
                                        };
                                    }
                                }
                            }
                        }
                        // Save month back into ForecastData
                        forecastData[monthKey] = monthData;
                    });
                    // ========================================================
                    // 13. Yearly data
                    // ========================================================
                    if (!isFrozen) {
                        const yearlyPlanInputs = document.querySelectorAll(".yearly-plan-input-rev");
                        const existingYearly = forecastData.yearly || {};
                        if (yearlyPlanInputs.length > 0) {
                            const updatedYearlyData = Object.assign({}, existingYearly);
                            const years = new Set();
                            yearlyPlanInputs.forEach((input) => {
                                const year = input.dataset.year;
                                if (year) {
                                    years.add(year);
                                }
                            });
                            Array.from(years)
                                .sort()
                                .forEach((year) => {
                                const planInput = document.querySelector(`.yearly-plan-input-rev[data-year="${year}"]`);
                                const currentInput = document.querySelector(`.yearly-current-input-rev[data-year="${year}"]`);
                                if (planInput && currentInput) {
                                    updatedYearlyData[year] = {
                                        plan: getNumericValue(planInput),
                                        actual: getNumericValue(currentInput)
                                    };
                                }
                            });
                            forecastData.yearly =
                                updatedYearlyData;
                        }
                        else if (existingYearly &&
                            Object.keys(existingYearly).length > 0) {
                            forecastData.yearly =
                                existingYearly;
                        }
                    }
                    else {
                        // Frozen:
                        // preserve existing yearly data exactly.
                        console.log("🔒 Frozen: yearly data preserved");
                    }
                    // ========================================================
                    // 14. Apply progression only for normal projects
                    // ========================================================
                    let finalData = forecastData;
                    if (!isFrozen) {
                        console.log("🔓 Applying forecast progression...");
                        finalData =
                            propagateCollectionValues(forecastData);
                    }
                    else {
                        console.log("🔒 Frozen: progression skipped");
                    }
                    // ========================================================
                    // 15. RevenueActual
                    // ========================================================
                    if (!isFrozen) {
                        const currentMonthInput = document.querySelector(`input.main-current-input[data-month="${currentMonthKey}"]`);
                        if (currentMonthInput) {
                            const currentValue = getNumericValue(currentMonthInput);
                            revenueActualData[currentMonthKey] =
                                currentValue;
                            console.log(`📊 RevenueActual ${currentMonthKey}:`, currentValue);
                        }
                        Object.keys(revenueActualData).forEach((monthKey) => {
                            const month = finalData[monthKey];
                            if (month) {
                                month.monthActual =
                                    revenueActualData[monthKey];
                            }
                        });
                    }
                    else {
                        // Frozen:
                        // RevenueActual remains unchanged.
                        console.log("🔒 Frozen: RevenueActual preserved");
                    }
                    // ========================================================
                    // 16. Build SharePoint payload
                    // ========================================================
                    const digest = await getDigest();
                    const internalRemark = (remarksInternal === null || remarksInternal === void 0 ? void 0 : remarksInternal.value) || "";
                    const sprintCellRemark = (remarksSprintCell === null || remarksSprintCell === void 0 ? void 0 : remarksSprintCell.value) || "";
                    const combinedRemarks = internalRemark + "|||" + sprintCellRemark;
                    const forecastDataString = JSON.stringify(finalData);
                    let revenueActualString = JSON.stringify(revenueActualData);
                    // --------------------------------------------------------
                    // RevenueActual SharePoint field length protection
                    // --------------------------------------------------------
                    if (revenueActualString.length > 255) {
                        console.warn("⚠️ RevenueActual too long:", revenueActualString.length);
                        const cleanedRevenueActual = {};
                        Object.keys(revenueActualData).forEach((key) => {
                            const value = revenueActualData[key];
                            if (value > 0 &&
                                /^\d{4}-\d{2}$/.test(key)) {
                                cleanedRevenueActual[key] =
                                    value;
                            }
                        });
                        revenueActualString =
                            JSON.stringify(cleanedRevenueActual);
                    }
                    const payload = {
                        __metadata: {
                            type: "SP.Data.Revenue_x0020_Forecast_x0020_ListListItem"
                        },
                        Title: generateForecastTitle(project, finalData),
                        ProjectId: project.Id,
                        ProfitCentre: project.PC || "",
                        ForecastData: forecastDataString,
                        RevenueActual: revenueActualString,
                        Remarks: combinedRemarks,
                        UpdatedDate: new Date().toISOString()
                    };
                    // OBL is only set when creating a new project.
                    if (!editingId && isNewProject) {
                        payload.OBL =
                            getNumericValue(oblLastMonthInput);
                    }
                    // CreatedDate only for new record.
                    if (!editingId) {
                        payload.CreatedDate =
                            new Date().toISOString();
                    }
                    console.log("============================================");
                    console.log("📤 SHAREPOINT SAVE PAYLOAD");
                    console.log("📤 Editing ID:", editingId);
                    console.log("📤 Frozen:", isFrozen);
                    console.log("📤 ForecastData length:", forecastDataString.length);
                    console.log("📤 Payload:", payload);
                    // ========================================================
                    // 17. SharePoint URL
                    // ========================================================
                    const url = editingId
                        ? `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${editingId})`
                        : `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
                    // ========================================================
                    // 18. SharePoint headers
                    // ========================================================
                    const headers = {
                        Accept: "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": digest
                    };
                    // Existing record = MERGE
                    if (editingId) {
                        headers["IF-MATCH"] = "*";
                        headers["X-HTTP-Method"] = "MERGE";
                        console.log(`🔄 Updating SharePoint item ${editingId}`);
                    }
                    else {
                        console.log("➕ Creating new SharePoint item");
                    }
                    // ========================================================
                    // 19. SEND TO SHAREPOINT
                    // ========================================================
                    const response = await fetch(url, {
                        method: "POST",
                        headers,
                        body: JSON.stringify(payload)
                    });
                    console.log("📥 SharePoint response:", response.status, response.statusText);
                    // ========================================================
                    // 20. Handle SharePoint response
                    // ========================================================
                    if (!response.ok) {
                        const errorText = await response.text();
                        console.error("❌ SharePoint save failed:", response.status, errorText);
                        toast("error", `Error saving forecast: ${response.status}`);
                        return;
                    }
                    // ========================================================
                    // 21. SUCCESS
                    // ========================================================
                    console.log("✅ Revenue forecast saved successfully");
                    toast("success", isFrozen
                        ? "✅ Further changes saved successfully!"
                        : editingId
                            ? "✅ Revenue Plan updated successfully!"
                            : "✅ Revenue Plan saved successfully!");
                    // --------------------------------------------------------
                    // Reload DB data so UI contains the saved values
                    // --------------------------------------------------------
                    await loadForecasts();
                    // Rebuild pagination using freshly loaded data
                    initializePagination(filteredData.length
                        ? filteredData
                        : forecasts);
                    // Close form
                    hideForm();
                    // Reset form state
                    resetForm();
                    console.log("============================================");
                    console.log("✅ REVENUE SAVE COMPLETED");
                    console.log("============================================");
                }
                catch (error) {
                    console.error("❌ Error saving revenue forecast:", error);
                    toast("error", "Error saving forecast!");
                }
                finally {
                    // --------------------------------------------------------
                    // Re-enable Save button
                    // --------------------------------------------------------
                    if (saveButton) {
                        saveButton.disabled = false;
                        saveButton.dataset.saving = "false";
                    }
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
                // Get assigned projects for the current user
                // const assignedProjects = getAssignedProjects(projects);
                // ✅ Use getAssignedProjectValues helper
                const assignedProjectTitles = getAssignedProjectValues('Title');
                const assignedCustomers = getAssignedProjectValues('CustomerName');
                const assignedManagers = getAssignedProjectValues('ProjectManager');
                const assignedGroupHeads = getAssignedProjectValues('GroupHead');
                const assignedReferences = getAssignedProjectValues('ProjectReference');
                // ✅ FILTER: Project dropdown - ONLY show assigned projects
                filterProject.innerHTML = '<option value="">All Projects</option>' +
                    assignedProjectTitles.map(project => `<option value="${project}">${project}</option>`).join("");
                // ✅ FILTER: Customer dropdown - ONLY customers from assigned projects
                filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                    assignedCustomers.map(customer => `<option value="${customer}">${customer}</option>`).join("");
                // ✅ FILTER: Project Manager dropdown - ONLY managers from assigned projects
                filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                    assignedManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");
                // ✅ FILTER: Group Head dropdown - ONLY group heads from assigned projects
                filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
                    assignedGroupHeads.map(head => `<option value="${head}">${head}</option>`).join("");
                // ✅ FILTER: Project Reference dropdown - ONLY references from assigned projects
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
                let filtered = [...forecasts];
                // ✅ Filter by assigned projects for non-admin/non-GH users
                const assignedProjects = getAssignedProjects(projects);
                const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                const isGHHeadUser = isGHHead();
                const isFullControlUser = hasFullControl();
                console.log('🔍 filterTable - User roles:', { isGHHeadUser, isFullControlUser });
                console.log('🔍 Assigned project titles:', Array.from(assignedProjectTitles));
                console.log('🔍 Total forecasts before filter:', filtered.length);
                if (!isGHHeadUser && !isFullControlUser) {
                    const beforeCount = filtered.length;
                    filtered = filtered.filter(forecast => {
                        var _a;
                        const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        const isAssigned = projectTitle && assignedProjectTitles.has(projectTitle);
                        if (!isAssigned) {
                            console.log(`🔍 Filtering out: ${projectTitle} (not assigned to user)`);
                        }
                        return isAssigned;
                    });
                    console.log(`🔍 Filtered from ${beforeCount} to ${filtered.length} records (non-GH restriction)`);
                }
                // ✅ FILTER OUT EMPTY PROJECT TITLES - This removes the blank rows
                const beforeEmptyFilter = filtered.length;
                filtered = filtered.filter(forecast => {
                    var _a;
                    const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                    const isValid = projectTitle && projectTitle.trim() !== '';
                    if (!isValid) {
                        console.log(`⚠️ Filtering out record with empty project title:`, forecast);
                    }
                    return isValid;
                });
                console.log(`🔍 Removed ${beforeEmptyFilter - filtered.length} records with empty project titles`);
                // Apply other filters
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
                console.log('🔍 Final filtered count:', filtered.length);
                // Always use the current pageSize value from the dropdown
                if (pageSizeSelect) {
                    pageSize = parseInt(pageSizeSelect.value) || 10;
                }
                if (filtered.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:20px;">No records found for selected filters</td></tr>`;
                    const totalsHeader = document.querySelector("#totalsHeader");
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    filteredData = [];
                    currentPage = 1;
                    calculateTotalPages();
                    updatePaginationControls();
                    renderCurrentPage();
                }
                else {
                    // ✅ Set filteredData to the filtered array
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
            // const applyAutoProgression = async (): Promise<void> => { 
            //   if (forecasts.length === 0) { 
            //     return;
            //   }
            //   const digest = await getDigest();
            //   let progressedCount = 0;
            //   let errorCount = 0;
            //   for (const forecast of forecasts) {
            //     try {
            //       if (!forecast.ForecastData) { 
            //         continue;
            //       }
            //       const originalData = safeJsonParse(forecast.ForecastData, {});
            //       if (checkIfNeedsProgression(originalData)) {  
            //         const progressedData = propagateCollectionValues(originalData);
            //         const payload = {
            //           __metadata: { type: "SP.Data.Revenue_x0020_Forecast_x0020_ListListItem" },
            //           ForecastData: JSON.stringify(progressedData),
            //           UpdatedDate: new Date().toISOString()
            //         };
            //         const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${forecast.Id})`, {
            //           method: "POST",
            //           headers: {
            //             Accept: "application/json;odata=verbose",
            //             "Content-Type": "application/json;odata=verbose",
            //             "X-RequestDigest": digest,
            //             "IF-MATCH": "*",
            //             "X-HTTP-Method": "MERGE"
            //           },
            //           body: JSON.stringify(payload)
            //         });
            //         if (response.ok) {
            //           progressedCount++; 
            //         } else {
            //           errorCount++; 
            //         }
            //         await new Promise(resolve => setTimeout(resolve, 100));
            //       } else { 
            //       }
            //     } catch (error) {
            //       errorCount++; 
            //     }
            //   }
            //   if (progressedCount > 0) { 
            //     await loadForecasts();
            //     toast("success", `Monthly and yearly data progressed for ${progressedCount} revenue forecasts`);
            //   } else if (errorCount > 0) { 
            //     toast("warning", `Some errors occurred during monthly and yearly progression`);
            //   } else { 
            //   }
            // };
            const checkIfNeedsProgression = (forecastData) => {
                const now = new Date();
                const currentMonthKey = formatMonthKey(now);
                const months = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
                if (months.length === 0)
                    return true;
                // Get the LAST month in the data (most recent)
                const lastDataMonth = months[months.length - 1];
                // Get the ideal months structure
                const idealMonths = generateIdealMonthStructure();
                const lastIdealMonth = idealMonths[idealMonths.length - 1];
                // Check if the data already has the current month
                const hasCurrentMonth = months.includes(currentMonthKey);
                // If the data doesn't have the current month, we need progression
                if (!hasCurrentMonth) {
                    return true;
                }
                // If the last data month is older than the last ideal month, progression is needed
                const monthlyProgressionNeeded = lastDataMonth < lastIdealMonth;
                // Check yearly progression
                const currentYear = now.getFullYear();
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
                // ✅ monthActual stays inside the data - only 1 argument
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
                            actual: 0
                        };
                    }
                });
                updatedData.yearly = updatedYearlyData;
                return updatedData;
            };
            const importActualDatas = () => {
                const modalHTMLRev = `
          <div class="import-modal" style="display: none;">
            <div class="import-modal-content">
              <div class="import-modal-header">
                <h3>Import Monthly Actual Revenue Data</h3>
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
                  <p><small>Supported format: Excel with Project Reference, Month, and Actual Revenue columns</small></p>
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
                const existingModal = document.querySelector('.import-modal');
                if (existingModal) {
                    existingModal.remove();
                }
                document.body.insertAdjacentHTML("beforeend", modalHTMLRev);
                const modal = document.querySelector(".import-modal");
                setTimeout(() => {
                    modal.style.display = "flex";
                    modal.style.position = "fixed";
                    modal.style.top = "0";
                    modal.style.left = "0";
                    modal.style.width = "100%";
                    modal.style.height = "100%";
                    modal.style.backgroundColor = "rgba(0,0,0,0.5)";
                    modal.style.zIndex = "9999";
                    modal.style.justifyContent = "center";
                    modal.style.alignItems = "center";
                    const modalContent = modal.querySelector('.import-modal-content');
                    if (modalContent) {
                        modalContent.style.backgroundColor = "white";
                        modalContent.style.padding = "20px";
                        modalContent.style.borderRadius = "10px";
                        modalContent.style.maxWidth = "90%";
                        modalContent.style.maxHeight = "90%";
                        modalContent.style.overflow = "auto";
                        modalContent.style.boxShadow = "0 4px 20px rgba(0,0,0,0.2)";
                    }
                }, 10);
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
                const monthMap = {
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
                const parseMonth = (monthStr) => {
                    if (!monthStr)
                        return "";
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
                            }
                            else if (year.length === 2) {
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
                const getListEntityTypeForImport = async () => {
                    if (listEntityType)
                        return listEntityType;
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                        headers: { Accept: "application/json;odata=verbose" },
                    });
                    const data = await res.json();
                    listEntityType = data.d.ListItemEntityTypeFullName;
                    return listEntityType;
                };
                // ============================================================
                // Updated updateRevenueData - stores ONLY values > 0
                // ============================================================
                const updateRevenueData = (existingData, month, actualRevenue) => {
                    let finalData = {};
                    // Parse existing data
                    if (existingData && existingData.trim() !== "" && existingData !== "{}" && existingData !== "null" && existingData !== "undefined") {
                        try {
                            const parsed = JSON.parse(existingData);
                            finalData = Object.assign({}, parsed);
                        }
                        catch (parseError) {
                            finalData = {};
                        }
                    }
                    else {
                        finalData = {};
                    }
                    // Update the specific month
                    if (actualRevenue > 0) {
                        finalData[month] = actualRevenue;
                    }
                    else {
                        // If value is 0, remove it to keep data compact
                        delete finalData[month];
                    }
                    // Return compact JSON with only values > 0
                    const result = {};
                    Object.keys(finalData).forEach(key => {
                        if (finalData[key] > 0) {
                            result[key] = finalData[key];
                        }
                    });
                    return JSON.stringify(result);
                };
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
                                    if (row["Month"] || row["month"]) {
                                        let monthValue = row["Month"] || row["month"];
                                        if (typeof monthValue === "number") {
                                            const excelDate = new Date(Math.round((monthValue - 25569) * 86400 * 1000));
                                            const yyyy = excelDate.getFullYear();
                                            const mm = String(excelDate.getMonth() + 1).padStart(2, "0");
                                            monthValue = `${yyyy}-${mm}`;
                                        }
                                        const monthStr = monthValue.toString();
                                        row.parsedMonth = parseMonth(monthStr);
                                    }
                                    const hasProjectRef = row["Project Reference"] || row["project_reference"];
                                    const hasMonth = row.parsedMonth;
                                    if (hasProjectRef && hasMonth) {
                                        rows.push(row);
                                    }
                                    else {
                                    }
                                }
                                resolve(rows);
                            }
                            catch (err) {
                                reject(err);
                            }
                        };
                        reader.onerror = (error) => {
                            console.error("❌ FileReader error:", error);
                            reject(new Error("Failed to read file"));
                        };
                        reader.readAsArrayBuffer(file);
                    });
                };
                const validateImportData = (data) => {
                    if (data.length === 0) {
                        return false;
                    }
                    const first = data[0];
                    const keys = Object.keys(first).map((k) => k.toLowerCase());
                    const hasProjectRef = keys.some((k) => k.includes("project") && k.includes("reference"));
                    const hasMonth = keys.some((k) => k.includes("month"));
                    const hasRevenue = keys.some((k) => k.includes("actual") && (k.includes("revenue") || k.includes("rev")));
                    console.log("✅ Validation results:", { hasProjectRef, hasMonth, hasRevenue, dataLength: data.length });
                    return hasProjectRef && hasMonth;
                };
                const showImportPreview = (data) => {
                    const monthSummary = {};
                    data.forEach((row) => {
                        const month = row.parsedMonth;
                        if (month) {
                            monthSummary[month] = (monthSummary[month] || 0) + 1;
                        }
                    });
                    const monthSummaryText = Object.keys(monthSummary)
                        .map((month) => {
                        return `${month}: ${monthSummary[month]} records`;
                    })
                        .join(", ");
                    const previewHTML = `
                <div style="padding: 20px; background: #f8f9fa; border-radius: 8px;">
                    <h4 style="color: #28a745; margin-bottom: 15px;">✅ File Ready for Import</h4>
                    <p><strong>Found ${data.length} records to import:</strong></p>
                    <p><strong>Month Summary:</strong> ${monthSummaryText || "No months found"}</p>
                    <div class="preview-table-container" style="max-height: 400px; overflow-y: auto;">
                        <table class="preview-table" style="width: 100%; border-collapse: collapse;">
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
                        <p style="margin: 5px 0 0 0; font-size: 12px;"><strong>Note:</strong> Only the imported month will be updated, all other data preserved</p>
                    </div>
                </div>`;
                    importPreview.innerHTML = previewHTML;
                };
                const performImport = async () => {
                    confirmImport.disabled = true;
                    confirmImport.textContent = "Importing...";
                    importStatus.style.display = "block";
                    importStatus.textContent = "Updating Revenue Data...";
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
                        const forecast = forecasts.find((f) => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Title) === project.Title; });
                        if (!(forecast === null || forecast === void 0 ? void 0 : forecast.Id)) {
                            notFound++;
                            continue;
                        }
                        try {
                            const digest = await getDigest();
                            const currentForecastResponse = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${forecast.Id})?$select=Id,RevenueActual`, {
                                method: "GET",
                                headers: {
                                    Accept: "application/json;odata=verbose",
                                },
                            });
                            if (!currentForecastResponse.ok) {
                                throw new Error(`Failed to fetch current data: ${currentForecastResponse.status}`);
                            }
                            const currentData = await currentForecastResponse.json();
                            const currentRevenueData = currentData.d.RevenueActual || "{}";
                            const updatedRevenueData = updateRevenueData(currentRevenueData, month, actualRevenue);
                            try {
                                const before = JSON.parse(currentRevenueData);
                                const after = JSON.parse(updatedRevenueData);
                                const changedMonths = Object.keys(after).filter((m) => before[m] !== after[m]);
                                console.log(`🔍 CHANGED MONTHS: ${changedMonths.join(", ")}`);
                            }
                            catch (e) {
                            }
                            const payload = {
                                __metadata: { type: entityType },
                                RevenueActual: updatedRevenueData,
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
                            }
                            else {
                                error++;
                                const text = await res.text();
                                console.warn(`⚠️ Failed for ${projectRef} (${month}):`, text);
                            }
                            await new Promise((r) => setTimeout(r, 100));
                        }
                        catch (err) {
                            error++;
                            console.error(`❌ Error updating ${projectRef} (${month}):`, err);
                        }
                    }
                    skipped = importData.length - (success + error + notFound + invalidMonth);
                    importStatus.className = "import-status success";
                    importStatus.innerHTML = `
            <strong>Import Complete!</strong><br>
            ✅ Success: ${success} records imported<br>
            ❌ Errors: ${error} records<br>
            🔍 Project Not Found: ${notFound} records<br>
            📅 Invalid Month: ${invalidMonth} records<br>
            ⏭️ Skipped: ${skipped} records<br>
            <small>Only the imported month was updated, all other data preserved</small>
        `;
                    await loadForecasts();
                    initializePagination(forecasts);
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
                const closeModalHandler = () => modal.remove();
                closeModal.addEventListener("click", closeModalHandler);
                cancelImport.addEventListener("click", closeModalHandler);
                modal.addEventListener("click", (e) => {
                    if (e.target === modal)
                        closeModalHandler();
                });
                confirmImport.addEventListener("click", performImport);
            };
            const initialize = async () => {
                try {
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
                    console.log("🚀 Total forecasts loaded:", forecasts.length);
                    // Initialize page size from dropdown
                    if (pageSizeSelect) {
                        pageSize = parseInt(pageSizeSelect.value) || 10;
                        console.log(`📊 Initial page size set to: ${pageSize}`);
                    }
                    if (forecasts.length > 0) {
                        await applyAutoProgression();
                        await loadForecasts();
                    }
                    initializeFilters();
                    // Setup pagination event listeners
                    setupPaginationEventListeners();
                    // Call filterTable() to apply project filtering
                    filterTable();
                    // Update UI based on permissions
                    updateUIBasedOnPermissions();
                    // Only attach import button if user has permission
                    if (canImportData()) {
                        importActualBtns.addEventListener("click", importActualDatas);
                    }
                    console.log("✅ Application initialized successfully");
                    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                    console.log(`🔒 Freeze permission: ${canFreeze() ? 'ENABLED' : 'DISABLED'}`);
                }
                catch (error) {
                    console.error("❌ Failed to initialize application:", error);
                    toast("error", "Failed to initialize application");
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
                                    showForm(true);
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
            const exportHistoricalBtn = document.querySelector(".export-historical-btn");
            if (exportHistoricalBtn) {
                exportHistoricalBtn.addEventListener("click", exportHistoricalToExcel);
            }
            clearFiltersBtn.addEventListener("click", clearFilters);
            // Initialize the application      
            initialize();
        }, 100);
    }
};
//# sourceMappingURL=revenue.js.map
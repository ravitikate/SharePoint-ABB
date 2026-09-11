import Swal from "sweetalert2";
import * as XLSX from "xlsx-js-style";
import { saveAs } from "file-saver";
import "../styles/common.css";
import "../styles/procurement.css";
// Status order based on workflow
const statusOptions = [
    "RFQ Pending",
    "Offer Pending",
    "TC pending",
    "Order Closure pending",
    "PR pending",
    "PO pending",
    "Document Approval Pending",
    "MFC Pending",
    "Inspection Pending",
    "Dispatch Pending",
    "Delivery Pending",
    "Delivered",
];
const procurementContent = `
<div class="project-page fade-in">
  <section class="table-section card">
    <div class="page-header">
      <h1>Procurement Tracker</h1>
      <div>
        <button class="btn-primary show-form-btn">+ Add Procurement</button>
      </div>
    </div>
    
    <div class="table-actions">
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
        <label for="filterStatus">Status</label>
        <select id="filterStatus"><option value="">All Status</option></select>
      </div>
      <div class="tableColp4">
        <label for="filterMake">Make</label>
        <select id="filterMake"><option value="">All Makes</option></select>
      </div> 
      <div class="tableColp4"> 
        <label>Actions</label>
        <button class="btn-secondary clear-filters-btn">Reset Filter</button> 
      </div> 
      <div class="tableColp4"> 
        <label>&nbsp;</label>
        <button class="btn-secondary export-excel-btn">Export Excel</button>
      </div>
    </div>

    <div style="overflow-x:auto;">
      <table id="procurementTable" class="modern-table">
        <thead>
          <tr>
            <th rowspan="2">Project Reference</th>
            <th rowspan="2">Project</th>
            <th rowspan="2">Customer</th>
            <th rowspan="2">Project Manager</th>
            <th rowspan="2">Group Head</th>
            <th rowspan="2">Total Line Items</th>
            <th>RFQ Pending</th>
            <th>Offer Pending</th>
            <th>TC pending</th>
            <th>Order Closure pending</th>
            <th>PR pending</th>
            <th>PO pending</th>
            <th>Document Approval Pending</th>
            <th>MFC Pending</th>
            <th>Inspection Pending</th>
            <th>Dispatch Pending</th>
            <th>Delivery Pending</th>
            <th>Delivered</th>
            <th rowspan="2">Actions</th>
          </tr>
        </thead>
        <tbody id="procurementBody"></tbody>
      </table>
    </div>
    <div class="pagination"></div>
  </section>

  <section class="form-section card" style="display:none;">
    <div class="page-header">
      <h1 id="formTitle">Add Procurement</h1>
    </div>
    <form id="procurementForm">
      <!-- Project Info Section -->
      <div class="project-info-wrapper">
        <div class="project-info-left">
          <div class="detail-item project-select-container">
            <span class="detail-label">Select Project</span>
            <select id="projectSelect" class="detail-value project-select" required></select>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project</span>
            <span class="detail-value" id="projectNameDisplay">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project Reference</span>
            <span class="detail-value" id="projectRefDisplay">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project Center</span>
            <span class="detail-value" id="projectCenterDisplay">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project Manager</span>
            <span class="detail-value" id="projectManagerDisplay">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Group Head</span>
            <span class="detail-value" id="groupHeadDisplay">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Customer</span>
            <span class="detail-value" id="customerDisplay">-</span>
          </div>
        </div>
        <div id="existingDataAlert" class="alert alert-info" style="display: none;">
          <strong>ℹ️ Existing Data Found!</strong> This project already has procurement data. You can view and edit it below.
        </div>
      </div>

      <!-- Procurement Table -->
      <div class="procurement-inputs" id="procurementContainerTable" style="display:none;">
        <div class="table-container">
          <table class="procurement-table">
            <thead>
              <tr> 
                <th class="sticky-col col-1" rowspan="2">
                  <div class="description-header-content">
                    <input
                      type="text"
                      id="descriptionFilter"
                      class="description-filter"
                      placeholder="Filter Description..."
                      aria-label="Filter Description"
                      autocomplete="off"
                    >
                    <span class="description-header-title">Description</span>
                  </div>
                </th>
                <th class="sticky-col col-2" rowspan="2">FCM Qty</th>
                <th class="sticky-col col-3" rowspan="2">Engg Qty</th>
                <th class="sticky-col col-4" rowspan="2">Status</th>
                <th class="sticky-col col-5" rowspan="2">Make</th>
                <th class="sticky-col col-6" rowspan="2">Model</th> 
                <th colspan="3">RFQ / Engg</th>
                <th colspan="3">Vendor Offer / SCM</th>
                <th colspan="3">Technical Clearance / ENGG</th>
                <th colspan="3">Order Finalization / SCM</th>
                <th colspan="3">PR / PM</th>
                <th colspan="3">Purchase Order / SCM</th>
                <th colspan="3">Document Approval / ENGG</th>
                <th colspan="3">Manufacturing Clearance / ENGG</th>
                <th colspan="3">Inspection</th>
                <th colspan="3">Material Dispatch</th>
                <th colspan="3">Delivery at Site</th>
                <th rowspan="2">Remarks</th>
                <th rowspan="2">Actions</th>
              </tr>
              <tr>
                ${Array(11).fill("<th>Plan</th><th>Actual</th><th>Forecast</th>").join("")}
              </tr>
            </thead>
            <tbody id="procurementContainer"></tbody>
            <tfoot>
              <tr>
                <td colspan="6" class="sticky-footer">
                  <button type="button" class="add-procurement-row btn-secondary">+ Add Line Item</button>
                </td>
                <td colspan="49"></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
      
      <div class="form-actions">
        <button type="submit" class="btn-primary">💾 Save</button>
        <button type="button" class="btn-secondary back-to-table">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
`;
export default {
    content: procurementContent,
    style: `
/* ============================================ */
/* PROCUREMENT TABLE - FIXED HEADER & STICKY COLUMNS */
/* ============================================ */

.procurement-inputs {
  border: 1px solid #e2e8f0;
  border-radius: 12px;
  overflow: hidden;
  position: relative;
}

.table-container {
  overflow: auto;
  max-height: 600px;
  position: relative;
}

.procurement-table {
  border-collapse: separate;
  border-spacing: 0;
  width: 100%;
  min-width: 2200px;
  font-size: 13px;
}

.procurement-table th,
.procurement-table td {
  border: 1px solid #e1e1e1;
  padding: 6px 8px;
  text-align: left;
  white-space: nowrap;
  min-width: 80px;
}

.procurement-table th {
  background: #f3f2f1;
  font-weight: 600;
  text-align: center;
  position: sticky;
  top: 0;
  z-index: 10;
}
/* ============================================ */
/* DESCRIPTION FILTER - EDIT MODE               */
/* ============================================ */

.procurement-table .description-header-content {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  justify-content: center;
  gap: 6px;
  width: 100%;
  text-align: left;
}

.procurement-table .description-header-title {
  display: block;
  width: 100%;
  text-align: left;
  font-weight: 600;
}

.procurement-table .description-filter {
  display: block;
  width: 100%;
  min-width: 0;
  box-sizing: border-box;
  padding: 5px 7px;
  border: 1px solid #c8c8c8;
  border-radius: 4px;
  background: #ffffff;
  color: #212529;
  font-size: 12px;
  font-weight: 400;
  line-height: 1.2;
  text-align: left;
  outline: none;
}

.procurement-table .description-filter::placeholder {
  color: #6c757d;
  opacity: 1;
}

.procurement-table .description-filter:focus {
  border-color: #0078d4;
  box-shadow: 0 0 0 1px #0078d4;
}

.form-section.view-mode .description-filter {
  display: none;
}

.form-section.view-mode .description-header-content {
  gap: 0;
}

@media (max-width: 768px) {
  .procurement-table .description-filter {
    font-size: 10px;
    padding: 3px 5px;
  }
}


/* ============================================ */
/* STICKY COLUMNS - Header & Data Both */
/* ============================================ */

.procurement-table .sticky-col {
  position: sticky;
  z-index: 5;
  background: white;
}

/* Column 1: Description */
.procurement-table .sticky-col.col-1 {
  left: 0;
  min-width: 150px;
  z-index: 15;
}

/* Column 2: FCM Qty */
.procurement-table .sticky-col.col-2 {
  left: 150px;
  min-width: 80px;
  z-index: 15;
}

/* Column 3: Engg Qty */
.procurement-table .sticky-col.col-3 {
  left: 230px;
  min-width: 80px;
  z-index: 15;
}

/* Column 4: Status */
.procurement-table .sticky-col.col-4 {
  left: 310px;
  min-width: 120px;
  z-index: 15;
}

/* Column 5: Make */
.procurement-table .sticky-col.col-5 {
  left: 430px;
  min-width: 100px;
  z-index: 15;
}

/* Column 6: Model */
.procurement-table .sticky-col.col-6 {
  left: 530px;
  min-width: 100px;
  z-index: 15;
}

/* ============================================ */
/* HEADER STICKY COLUMNS - Higher Z-Index */
/* ============================================ */

.procurement-table thead .sticky-col {
  z-index: 20 !important;
  background: #f3f2f1;
}

.procurement-table thead .sticky-col.col-1,
.procurement-table thead .sticky-col.col-2,
.procurement-table thead .sticky-col.col-3,
.procurement-table thead .sticky-col.col-4,
.procurement-table thead .sticky-col.col-5,
.procurement-table thead .sticky-col.col-6 {
  z-index: 25 !important;
}

/* ============================================ */
/* STICKY FOOTER */
/* ============================================ */

.procurement-table tfoot td.sticky-footer {
  position: sticky;
  left: 0;
  z-index: 5;
  background: white;
  border-right: 2px solid #e1e1e1;
}

/* ============================================ */
/* HOVER EFFECTS */
/* ============================================ */

.procurement-table tbody tr:hover {
  background: #f5f5f5;
}

.procurement-table tbody tr:hover .sticky-col {
  background: #f5f5f5;
}

/* ============================================ */
/* VIEW MODE STYLES */
/* ============================================ */

.view-mode .procurement-table thead th {
  background: #e9ecef;
}

.view-mode .procurement-table tbody td {
  background: #f8f9fa;
}

.view-mode .procurement-table tbody .sticky-col {
  background: #f8f9fa !important;
}

.view-mode .procurement-table tbody tr:hover .sticky-col {
  background: #e9ecef !important;
}

.view-mode .procurement-table thead .sticky-col {
  background: #e9ecef !important;
}

/* ============================================ */
/* FORM CONTROLS INSIDE TABLE */
/* ============================================ */

.procurement-table .form-control {
  width: 100%;
  min-width: 60px;
  padding: 4px 6px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 12px;
  box-sizing: border-box;
}

.procurement-table .form-control:focus {
  border-color: #0078d4;
  outline: none;
  box-shadow: 0 0 3px rgba(0, 120, 212, 0.3);
}

.procurement-table input[type="date"] {
  min-width: 90px;
}

.procurement-table input[type="number"] {
  min-width: 60px;
  text-align: right;
}

/* ============================================ */
/* STATUS DISPLAY */
/* ============================================ */

.status-text-display {
  padding: 4px 8px;
  background: #eef6ff;
  border: 1px solid #ced4da;
  border-radius: 4px;
  font-weight: 600;
  color: #0078d4;
  display: inline-block;
  min-width: 100px;
  text-align: center;
  font-size: 12px;
}

.status-text-display.delivered {
  background: #e8f5e8;
  color: #1b5e20;
}

.status-auto-note {
  display: block;
  font-size: 9px;
  color: #6c757d;
  margin-top: 2px;
  font-style: italic;
}

/* ============================================ */
/* ADD/REMOVE ROW BUTTONS */
/* ============================================ */

.add-procurement-row {
  padding: 4px 12px;
  font-size: 12px;
}

.remove-row {
  background: #dc3545;
  color: white;
  border: none;
  border-radius: 4px;
  padding: 4px 8px;
  cursor: pointer;
  font-size: 12px;
}

.remove-row:hover {
  background: #c82333;
}

/* ============================================ */
/* ENGG QTY GREYED OUT */
/* ============================================ */

.engg-qty-disabled {
  background-color: #e9ecef !important;
  color: #6c757d !important;
  cursor: not-allowed !important;
  opacity: 0.7;
}

.engg-qty-disabled::placeholder {
  color: #adb5bd !important;
}

.engg-qty-wrapper {
  position: relative;
}

.engg-qty-wrapper .lock-icon {
  position: absolute;
  right: 6px;
  top: 50%;
  transform: translateY(-50%);
  color: #6c757d;
  font-size: 12px;
  pointer-events: none;
}

.engg-qty-wrapper .engg-qty-input {
  padding-right: 22px;
}

/* ============================================ */
/* STATUS BADGE COLORS */
/* ============================================ */

.status-badge {
  padding: 4px 8px;
  border-radius: 12px;
  font-size: 12px;
  font-weight: 600;
  text-align: center;
}

.status-badge[data-status="RFQ Pending"] {
  background-color: #e3f2fd;
  color: #1976d2;
}

.status-badge[data-status="Offer Pending"] {
  background-color: #fff3e0;
  color: #f57c00;
}

.status-badge[data-status="TC pending"] {
  background-color: #e8f5e8;
  color: #388e3c;
}

.status-badge[data-status="Order Closure pending"] {
  background-color: #fce4ec;
  color: #c2185b;
}

.status-badge[data-status="PR pending"] {
  background-color: #f3e5f5;
  color: #7b1fa2;
}

.status-badge[data-status="PO pending"] {
  background-color: #e0f2f1;
  color: #00796b;
}

.status-badge[data-status="Document Approval Pending"] {
  background-color: #fff8e1;
  color: #ff8f00;
}

.status-badge[data-status="MFC Pending"] {
  background-color: #e8f5e8;
  color: #2e7d32;
}

.status-badge[data-status="Inspection Pending"] {
  background-color: #fff3e0;
  color: #e65100;
}

.status-badge[data-status="Dispatch Pending"] {
  background-color: #e3f2fd;
  color: #0d47a1;
}

.status-badge[data-status="Delivery Pending"] {
  background-color: #fce4ec;
  color: #880e4f;
}

.status-badge[data-status="Delivered"] {
  background-color: #e8f5e8;
  color: #1b5e20;
}

/* ============================================ */
/* VIEW FIELD STYLES */
/* ============================================ */

.procurement-page .view-field {
  padding: 8px 12px;
  background: #f8f9fa;
  border: 1px solid #e9ecef;
  border-radius: 4px;
  min-height: 38px;
  display: flex;
  align-items: center;
}

.procurement-page .description-view {
  max-width: 200px;
  overflow: hidden;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  line-clamp: 2;
  -webkit-box-orient: vertical;
  cursor: pointer;
  transition: all 0.3s ease;
}

.procurement-page .description-view:hover {
  position: relative;
  z-index: 1000;
  background: white;
  border: 1px solid #0078d4;
  box-shadow: 0 4px 8px rgba(0,0,0,0.1);
  max-width: 400px;
  white-space: normal;
  -webkit-line-clamp: unset;
  line-clamp: unset;
  overflow: visible;
  word-wrap: break-word;
}

.procurement-page .view-mode .form-control {
  background-color: #f8f9fa !important;
  border-color: #e9ecef !important;
  cursor: not-allowed;
}

.procurement-page .remarks-view {
  max-width: 200px;
  word-wrap: break-word;
}

/* ============================================ */
/* CLICKABLE LINK STYLES */
/* ============================================ */

#procurementBody .clickable-link {
  color: #0078d4 !important;
  cursor: pointer !important;
  text-decoration: none !important;
  font-weight: 600 !important;
  border-bottom: 1px solid transparent !important;
  transition: all 0.2s ease !important;
  display: inline-block !important;
  padding: 2px 4px !important;
  border-radius: 3px !important;
}

#procurementBody .clickable-link:hover {
  color: #0056b3 !important;
  border-bottom: 1px solid #0056b3 !important;
  background-color: rgba(0, 120, 212, 0.08) !important;
}

#procurementBody .clickable-link:active {
  color: #003d7a !important;
  background-color: rgba(0, 120, 212, 0.15) !important;
}

#procurementBody td:first-child,
#procurementBody td:nth-child(2) {
  padding: 8px 12px !important;
}

#procurementBody .clickable-link {
  white-space: nowrap !important;
}

/* ============================================ */
/* PROJECT INFO WRAPPER */
/* ============================================ */

.project-info-wrapper {
  background: #f8f9fa;
  border-radius: 8px;
  border: 1px solid #e9ecef;
  padding: 20px 24px;
  margin-bottom: 25px;
  width: 100%;
  box-sizing: border-box;
}

.project-info-left {
  display: flex;
  flex-direction: column;
  gap: 0;
}

.project-info-left .detail-item {
  display: flex;
  flex-direction: row;
  align-items: center;
  padding: 8px 0;
  font-size: 14px;
  border-bottom: 1px solid #f0f0f0;
  text-align: left;
  justify-content: flex-start;
  width: 100%;
}

.project-info-left .detail-item:last-child {
  border-bottom: none;
}

.project-info-left .detail-label {
  font-weight: 700;
  white-space: nowrap;
  min-width: 140px;
  color: #495057;
  font-size: 13px;
}

.project-info-left .detail-value {
  font-weight: 500;
  color: #212529;
  font-size: 14px;
  padding-left: 10px;
}

.project-select-container .project-select {
  padding: 6px 12px;
  border: 1px solid #ced4da;
  border-radius: 4px;
  font-size: 14px;
  background-color: white;
  width: 100%;
  max-width: 400px;
}

.project-select-container .project-select:focus {
  border-color: #0078d4;
  outline: 0;
  box-shadow: 0 0 0 0.2rem rgba(0, 120, 212, 0.25);
}

.project-select-container .project-select:disabled {
  background-color: #e9ecef;
  cursor: not-allowed;
}

/* ============================================ */
/* ALERT */
/* ============================================ */

.alert-info {
  background-color: #e3f2fd;
  border-left: 4px solid #0078d4;
  padding: 12px 16px;
  border-radius: 4px;
  color: #0056b3;
  font-size: 13px;
  margin-top: 12px;
  width: 100%;
  box-sizing: border-box;
}

.alert-info strong {
  font-weight: 600;
}

/* ============================================ */
/* FORM ACTIONS */
/* ============================================ */

.form-actions {
  display: flex;
  gap: 15px;
  margin-top: 20px;
  padding-top: 20px;
  border-top: 1px solid #e9ecef;
  justify-content: flex-end;
}

.form-actions .btn-primary,
.form-actions .btn-secondary {
  padding: 10px 24px;
  font-size: 14px;
}

/* ============================================ */
/* TOTAL LINE ITEMS COLUMN */
/* ============================================ */

.total-line-items-col {
  text-align: center;
  font-weight: 600;
  background-color: #f0f7ff;
  min-width: 80px;
}

/* ============================================ */
/* RESPONSIVE */
/* ============================================ */

@media (max-width: 768px) {
  .project-info-left .detail-item {
    flex-direction: column;
    align-items: flex-start;
    padding: 10px 0;
  }
  
  .project-info-left .detail-label {
    min-width: auto;
    margin-bottom: 2px;
    font-size: 12px;
  }
  
  .project-info-left .detail-value {
    padding-left: 0;
    font-size: 13px;
  }
  
  .project-info-wrapper {
    padding: 15px;
  }

  .project-select-container .project-select {
    max-width: 100%;
  }

  .form-actions {
    flex-direction: column;
  }
  
  .form-actions button {
    width: 100%;
  }

  .procurement-table {
    min-width: 1600px;
    font-size: 11px;
  }

  .procurement-table .sticky-col.col-1 {
    left: 0;
    min-width: 100px;
  }

  .procurement-table .sticky-col.col-2 {
    left: 100px;
    min-width: 60px;
  }

  .procurement-table .sticky-col.col-3 {
    left: 160px;
    min-width: 60px;
  }

  .procurement-table .sticky-col.col-4 {
    left: 220px;
    min-width: 100px;
  }

  .procurement-table .sticky-col.col-5 {
    left: 320px;
    min-width: 80px;
  }

  .procurement-table .sticky-col.col-6 {
    left: 400px;
    min-width: 80px;
  }

  .procurement-table th,
  .procurement-table td {
    padding: 4px 6px;
    font-size: 11px;
    min-width: 60px;
  }

  .procurement-table input[type="date"] {
    min-width: 70px;
    font-size: 10px;
  }

  .procurement-table .form-control {
    font-size: 10px;
    min-width: 40px;
    padding: 2px 4px;
  }

  .status-text-display {
    font-size: 10px;
    min-width: 70px;
    padding: 2px 4px;
  }

  #procurementBody .clickable-link {
    white-space: normal !important;
    word-break: break-word !important;
  }

  .table-actions {
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 10px;
    padding: 10px;
  }

  .procurement-page .description-view {
    max-width: 120px;
  }

  .table-container {
    max-height: 400px;
  }
}

@media (max-width: 480px) {
  .procurement-table {
    min-width: 1400px;
    font-size: 10px;
  }

  .procurement-table th,
  .procurement-table td {
    padding: 3px 4px;
    min-width: 50px;
  }

  .procurement-table .sticky-col.col-1 {
    min-width: 80px;
  }

  .procurement-table .sticky-col.col-2 {
    left: 80px;
    min-width: 50px;
  }

  .procurement-table .sticky-col.col-3 {
    left: 130px;
    min-width: 50px;
  }

  .procurement-table .sticky-col.col-4 {
    left: 180px;
    min-width: 80px;
  }

  .procurement-table .sticky-col.col-5 {
    left: 260px;
    min-width: 60px;
  }

  .procurement-table .sticky-col.col-6 {
    left: 320px;
    min-width: 60px;
  }

  .procurement-table input[type="date"] {
    min-width: 60px;
    font-size: 9px;
  }
}
`,
    script: () => {
        setTimeout(async () => {
            var _a;
            // Configuration
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/DigitalTransformationTest";
            const listName = "ProcurementTracker";
            const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
            // State variables
            let fieldMappings = {
                ProcurementForecast: "ProcurementForecast",
                Project: "Project",
            };
            let listEntityTypeName = "";
            let projects = [];
            let editingId = null;
            let procurementData = [];
            let currentPage = 1;
            const pageSize = 10;
            // ============================================ //
            // STATUS ORDER - 12 Steps                      //
            // ============================================ //
            const isFilled = (v) => !!v && String(v).trim() !== "";
            const getStatusFromActualDates = (item) => {
                if (!item)
                    return "RFQ Pending";
                if (!isFilled(item.RfqActual))
                    return "RFQ Pending";
                if (!isFilled(item.OfferActual))
                    return "Offer Pending";
                if (!isFilled(item.TechActual))
                    return "TC pending";
                if (!isFilled(item.OrderActual))
                    return "Order Closure pending";
                if (!isFilled(item.PRActual))
                    return "PR pending";
                if (!isFilled(item.POActual))
                    return "PO pending";
                if (!isFilled(item.DocActual))
                    return "Document Approval Pending";
                if (!isFilled(item.MfgActual))
                    return "MFC Pending";
                if (!isFilled(item.InspActual))
                    return "Inspection Pending";
                if (!isFilled(item.DispActual))
                    return "Dispatch Pending";
                if (!isFilled(item.DeliveryActual))
                    return "Delivery Pending";
                return "Delivered";
            };
            // ============================================ //
            // ENGG QTY GREY OUT LOGIC                      //
            // ============================================ //
            const updateEnggQtyState = (row) => {
                const poActualInput = row.querySelector('input[name="po_actual"]');
                const enggQtyInput = row.querySelector('input[name="engg_qty"]');
                const enggQtyWrapper = row.querySelector('.engg-qty-wrapper');
                if (!poActualInput || !enggQtyInput)
                    return;
                const isPOActualFilled = isFilled(poActualInput.value);
                if (isPOActualFilled) {
                    enggQtyInput.disabled = true;
                    enggQtyInput.classList.add('engg-qty-disabled');
                    if (enggQtyWrapper) {
                        if (!enggQtyWrapper.querySelector('.lock-icon')) {
                            const lockIcon = document.createElement('span');
                            lockIcon.className = 'lock-icon';
                            lockIcon.textContent = '🔒';
                            lockIcon.title = 'Engg Qty is locked after PO Actual date is filled';
                            enggQtyWrapper.appendChild(lockIcon);
                        }
                    }
                }
                else {
                    enggQtyInput.disabled = false;
                    enggQtyInput.classList.remove('engg-qty-disabled');
                    if (enggQtyWrapper) {
                        const lockIcon = enggQtyWrapper.querySelector('.lock-icon');
                        if (lockIcon)
                            lockIcon.remove();
                    }
                }
            };
            const wireEnggQtyLock = (row) => {
                const poActualInput = row.querySelector('input[name="po_actual"]');
                if (poActualInput) {
                    poActualInput.addEventListener('change', () => updateEnggQtyState(row));
                    poActualInput.addEventListener('input', () => updateEnggQtyState(row));
                }
                updateEnggQtyState(row);
            };
            const recalcRowStatus = (row) => {
                const statusDisplay = row.querySelector('.status-text-display');
                if (!statusDisplay)
                    return;
                const getVal = (name) => {
                    var _a;
                    return ((_a = row.querySelector(`input[name="${name}"]`)) === null || _a === void 0 ? void 0 : _a.value) || "";
                };
                const newStatus = getStatusFromActualDates({
                    RfqActual: getVal("rfq_actual"),
                    OfferActual: getVal("offer_actual"),
                    TechActual: getVal("tech_actual"),
                    OrderActual: getVal("order_actual"),
                    PRActual: getVal("pr_actual"),
                    POActual: getVal("po_actual"),
                    DocActual: getVal("doc_actual"),
                    MfgActual: getVal("mfg_actual"),
                    InspActual: getVal("insp_actual"),
                    DispActual: getVal("disp_actual"),
                    DeliveryActual: getVal("delivery_actual"),
                });
                statusDisplay.textContent = newStatus;
                statusDisplay.className = 'status-text-display';
                if (newStatus === 'Delivered') {
                    statusDisplay.classList.add('delivered');
                }
                updateEnggQtyState(row);
            };
            const wireStatusAutoLink = (row) => {
                const actualInputNames = [
                    "rfq_actual",
                    "offer_actual",
                    "tech_actual",
                    "order_actual",
                    "pr_actual",
                    "po_actual",
                    "doc_actual",
                    "mfg_actual",
                    "insp_actual",
                    "disp_actual",
                    "delivery_actual",
                ];
                actualInputNames.forEach((name) => {
                    const input = row.querySelector(`input[name="${name}"]`);
                    if (input) {
                        input.addEventListener("change", () => {
                            recalcRowStatus(row);
                        });
                        input.addEventListener("input", () => {
                            recalcRowStatus(row);
                        });
                    }
                });
                recalcRowStatus(row);
                wireEnggQtyLock(row);
            };
            // DOM Elements
            const projectSelect = document.getElementById("projectSelect");
            const showFormBtn = document.querySelector(".show-form-btn");
            const exportExcelBtn = document.querySelector(".export-excel-btn");
            const formSection = document.querySelector(".form-section");
            const tableSection = document.querySelector(".table-section");
            const backBtn = document.querySelector(".back-to-table");
            const procurementContainer = document.getElementById("procurementContainer");
            const procurementContainerTable = document.getElementById("procurementContainerTable");
            const tableBody = document.getElementById("procurementBody");
            const paginationDiv = document.querySelector(".pagination");
            const form = document.getElementById("procurementForm");
            const formTitle = document.getElementById("formTitle");
            const existingDataAlert = document.getElementById("existingDataAlert");
            // Description filter for the Edit Procurement table
            const descriptionFilter = document.getElementById("descriptionFilter");
            // Project Info Display Elements
            const projectNameDisplay = document.getElementById("projectNameDisplay");
            const projectRefDisplay = document.getElementById("projectRefDisplay");
            const projectCenterDisplay = document.getElementById("projectCenterDisplay");
            const projectManagerDisplay = document.getElementById("projectManagerDisplay");
            const groupHeadDisplay = document.getElementById("groupHeadDisplay");
            const customerDisplay = document.getElementById("customerDisplay");
            // Filter Elements
            const filterProject = document.getElementById("filterProject");
            const filterCustomer = document.getElementById("filterCustomer");
            const filterProjectManager = document.getElementById("filterProjectManager");
            const filterGroupHead = document.getElementById("filterGroupHead");
            const filterProjectReference = document.getElementById("filterProjectReference");
            const filterStatus = document.getElementById("filterStatus");
            const filterMake = document.getElementById("filterMake");
            const clearFiltersBtn = document.querySelector(".table-actions .btn-secondary");
            setTimeout(() => {
                if (filterProject) {
                    filterProject.addEventListener("change", filterData);
                }
                if (filterCustomer) {
                    filterCustomer.addEventListener("change", filterData);
                }
                if (filterProjectManager) {
                    filterProjectManager.addEventListener("change", filterData);
                }
                if (filterGroupHead) {
                    filterGroupHead.addEventListener("change", filterData);
                }
                if (filterProjectReference) {
                    filterProjectReference.addEventListener("change", filterData);
                }
                if (filterStatus) {
                    filterStatus.addEventListener("change", filterData);
                }
                if (filterMake) {
                    filterMake.addEventListener("change", filterData);
                }
                if (clearFiltersBtn) {
                    clearFiltersBtn.addEventListener("click", clearFilters);
                }
            }, 500);
            // Utility Functions
            const toast = (msg, type = "success") => {
                Swal.fire({
                    toast: true,
                    position: "top-end",
                    icon: type,
                    title: msg,
                    showConfirmButton: false,
                    timer: 3000,
                    timerProgressBar: true,
                });
            };
            const showExportProgress = (message) => {
                Swal.fire({
                    title: "Exporting...",
                    text: message,
                    allowOutsideClick: false,
                    didOpen: () => {
                        Swal.showLoading();
                    },
                });
            };
            const hideExportProgress = () => {
                Swal.close();
            };
            const getRequestDigest = async () => {
                try {
                    const res = await fetch(`${siteUrl}/_api/contextinfo`, {
                        method: "POST",
                        headers: {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                        },
                    });
                    if (!res.ok)
                        throw new Error(`Failed to get request digest: ${res.status}`);
                    const data = await res.json();
                    return data.d.GetContextWebInformation.FormDigestValue;
                }
                catch (error) {
                    console.error("Error getting request digest:", error);
                    throw error;
                }
            };
            // Field discovery functions
            const discoverListFields = async () => {
                try {
                    const listRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')`, {
                        headers: { Accept: "application/json;odata=verbose" },
                    });
                    if (!listRes.ok)
                        throw new Error(`Failed to get list: ${listRes.status}`);
                    const listData = await listRes.json();
                    const entityType = listData.d.ListItemEntityTypeFullName;
                    const fieldsRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
                        headers: { Accept: "application/json;odata=verbose" },
                    });
                    if (!fieldsRes.ok)
                        throw new Error(`Failed to get fields: ${fieldsRes.status}`);
                    const fieldsData = await fieldsRes.json();
                    const fields = fieldsData.d.results;
                    return { fields, entityType };
                }
                catch (error) {
                    console.error("Error discovering fields:", error);
                    throw error;
                }
            };
            const createFieldMappings = (fields) => {
                const mappings = {
                    ProcurementForecast: "ProcurementForecast",
                    Project: "Project",
                };
                fields.forEach((field) => {
                    var _a;
                    const fieldTitle = (_a = field.Title) === null || _a === void 0 ? void 0 : _a.toLowerCase().replace(/\s+/g, "");
                    const internalName = field.InternalName;
                    if ([
                        "Id",
                        "ID",
                        "Title",
                        "Created",
                        "Modified",
                        "Author",
                        "Editor",
                    ].includes(internalName)) {
                        return;
                    }
                    if (field.TypeAsString === "Note" &&
                        (fieldTitle.includes("procurement") ||
                            fieldTitle.includes("forecast") ||
                            fieldTitle.includes("data"))) {
                        mappings.ProcurementForecast = internalName;
                    }
                    if (field.TypeAsString === "Lookup" &&
                        (fieldTitle.includes("project") ||
                            field.InternalName.includes("Project"))) {
                        mappings.Project = internalName;
                    }
                });
                return mappings;
            };
            const testFieldMappings = async () => {
                try {
                    const testQuery = `${apiBase}?$select=Id,${fieldMappings.Project}/Id&$top=1&$expand=${fieldMappings.Project}`;
                    const testRes = await fetch(testQuery, {
                        headers: { Accept: "application/json;odata=verbose" },
                    });
                    if (testRes.ok) {
                        return true;
                    }
                    else {
                        const error = await testRes.json();
                        console.error("Field mappings test: FAILED", error);
                        return false;
                    }
                }
                catch (error) {
                    console.error("Field mappings test error:", error);
                    return false;
                }
            };
            const renderProjectDetails = (project) => {
                var _a;
                if (!project) {
                    if (projectNameDisplay)
                        projectNameDisplay.textContent = "-";
                    if (projectRefDisplay)
                        projectRefDisplay.textContent = "-";
                    if (projectCenterDisplay)
                        projectCenterDisplay.textContent = "-";
                    if (projectManagerDisplay)
                        projectManagerDisplay.textContent = "-";
                    if (groupHeadDisplay)
                        groupHeadDisplay.textContent = "-";
                    if (customerDisplay)
                        customerDisplay.textContent = "-";
                    return;
                }
                const managerName = typeof project.ProjectManager === "object"
                    ? ((_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || ""
                    : project.ProjectManager || "";
                if (projectNameDisplay)
                    projectNameDisplay.textContent = project.Title || "-";
                if (projectRefDisplay)
                    projectRefDisplay.textContent = project.ProjectReference || "-";
                if (projectCenterDisplay)
                    projectCenterDisplay.textContent = project.PC || "-";
                if (projectManagerDisplay)
                    projectManagerDisplay.textContent = managerName || "-";
                if (groupHeadDisplay)
                    groupHeadDisplay.textContent = project.GroupHead || "-";
                if (customerDisplay)
                    customerDisplay.textContent = project.CustomerName || "-";
            };
            const initApp = async () => {
                try {
                    const { fields, entityType } = await discoverListFields();
                    listEntityTypeName = entityType;
                    fieldMappings = createFieldMappings(fields);
                    const mappingsValid = await testFieldMappings();
                    if (!mappingsValid) {
                        throw new Error("Field mappings are invalid. Please check the field names in your SharePoint list.");
                    }
                    await loadProjects();
                    await fetchData();
                    initializeFilters();
                    renderTable(procurementData);
                    attachFilterEventListeners();
                    toast("Procurement tracker initialized successfully", "success");
                }
                catch (error) {
                    console.error("Initialization error:", error);
                    toast("Failed to initialize application. Check console for details.", "error");
                }
            };
            const attachFilterEventListeners = () => {
                filterProject.addEventListener("change", filterData);
                filterCustomer.addEventListener("change", filterData);
                filterProjectManager.addEventListener("change", filterData);
                filterGroupHead.addEventListener("change", filterData);
                filterProjectReference.addEventListener("change", filterData);
                filterStatus.addEventListener("change", filterData);
                filterMake.addEventListener("change", filterData);
                const clearFiltersBtn = document.querySelector(".table-actions .btn-secondary:nth-child(1)");
                if (clearFiltersBtn) {
                    clearFiltersBtn.addEventListener("click", clearFilters);
                }
            };
            const loadProjects = async () => {
                var _a;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('Project List')/items?$select=Id,Title,CustomerName,GroupHead,PC,ProjectReference,ProjectManager/Title&$expand=ProjectManager`, { headers: { Accept: "application/json;odata=verbose" } });
                    if (!res.ok)
                        throw new Error(`Failed to load projects: ${res.status}`);
                    const data = await res.json();
                    const allProjects = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    const uniqueProjects = Array.from(new Map(allProjects.map((p) => [p.Title, p])).values());
                    projects = uniqueProjects;
                    projectSelect.innerHTML =
                        '<option value="">Select Project</option>' +
                            projects
                                .map((p) => `<option value="${p.Id}">${p.Title}</option>`)
                                .join("");
                    filterProject.innerHTML =
                        '<option value="">All Projects</option>' +
                            projects
                                .map((p) => `<option value="${p.Title}">${p.Title}</option>`)
                                .join("");
                }
                catch (error) {
                    console.error("Error loading projects:", error);
                    toast("Failed to load projects", "error");
                }
            };
            const parseProcurementForecast = (procurementForecastJson) => {
                if (!procurementForecastJson)
                    return null;
                try {
                    const parsed = JSON.parse(procurementForecastJson);
                    return Array.isArray(parsed) ? parsed : [parsed];
                }
                catch (error) {
                    console.error("Error parsing ProcurementForecast JSON:", error);
                    return null;
                }
            };
            const calculateProjectStatusCounts = (procurementItems) => {
                const counts = {};
                statusOptions.forEach((status) => {
                    counts[status] = 0;
                });
                procurementItems.forEach((item) => {
                    if (item.Status && counts.hasOwnProperty(item.Status)) {
                        counts[item.Status]++;
                    }
                });
                return counts;
            };
            const fetchData = async () => {
                try {
                    const selectFields = [
                        "Id",
                        "Title",
                        fieldMappings.ProcurementForecast,
                        `${fieldMappings.Project}/Id`,
                        `${fieldMappings.Project}/Title`,
                    ].filter((field) => field && !field.startsWith("_"));
                    const query = `${apiBase}?$select=${selectFields.join(",")}&$expand=${fieldMappings.Project}&$top=5000`;
                    const res = await fetch(query, {
                        headers: { Accept: "application/json;odata=verbose" },
                    });
                    if (!res.ok)
                        throw new Error(`HTTP error! status: ${res.status}`);
                    const data = await res.json();
                    procurementData = data.d.results || [];
                    initializeFilters();
                    renderTable(procurementData);
                }
                catch (error) {
                    console.error("Error fetching data:", error);
                    toast("Failed to fetch data", "error");
                }
            };
            const initializeFilters = () => {
                const uniqueCustomers = [
                    ...new Set(projects.map((p) => p.CustomerName).filter(Boolean)),
                ].sort();
                filterCustomer.innerHTML =
                    '<option value="">All Customers</option>' +
                        uniqueCustomers
                            .map((customer) => `<option value="${customer}">${customer}</option>`)
                            .join("");
                const uniqueManagers = [
                    ...new Set(projects
                        .map((p) => {
                        var _a;
                        if (typeof p.ProjectManager === "object")
                            return (_a = p.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title;
                        return p.ProjectManager;
                    })
                        .filter(Boolean)),
                ].sort();
                filterProjectManager.innerHTML =
                    '<option value="">All Managers</option>' +
                        uniqueManagers
                            .map((manager) => `<option value="${manager}">${manager}</option>`)
                            .join("");
                const uniqueGroupHeads = [
                    ...new Set(projects.map((p) => p.GroupHead).filter(Boolean)),
                ].sort();
                filterGroupHead.innerHTML =
                    '<option value="">All Group Heads</option>' +
                        uniqueGroupHeads
                            .map((head) => `<option value="${head}">${head}</option>`)
                            .join("");
                const uniqueReferences = [
                    ...new Set(projects.map((p) => p.ProjectReference).filter(Boolean)),
                ].sort();
                filterProjectReference.innerHTML =
                    '<option value="">All References</option>' +
                        uniqueReferences
                            .map((ref) => `<option value="${ref}">${ref}</option>`)
                            .join("");
                const uniqueStatuses = [];
                procurementData.forEach((item) => {
                    const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) ||
                        [];
                    procurementItems.forEach((procItem) => {
                        if (procItem.Status) {
                            uniqueStatuses.push(procItem.Status);
                        }
                    });
                });
                const filteredStatuses = [...new Set(uniqueStatuses)].sort();
                filterStatus.innerHTML =
                    '<option value="">All Status</option>' +
                        filteredStatuses
                            .map((status) => `<option value="${status}">${status}</option>`)
                            .join("");
                const uniqueMakes = [];
                procurementData.forEach((item) => {
                    const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) ||
                        [];
                    procurementItems.forEach((procItem) => {
                        if (procItem.Make) {
                            uniqueMakes.push(procItem.Make);
                        }
                    });
                });
                const filteredMakes = [...new Set(uniqueMakes)].sort();
                filterMake.innerHTML =
                    '<option value="">All Makes</option>' +
                        filteredMakes
                            .map((make) => `<option value="${make}">${make}</option>`)
                            .join("");
            };
            const filterData = () => {
                const project = filterProject.value;
                const customer = filterCustomer.value;
                const projectManager = filterProjectManager.value;
                const groupHead = filterGroupHead.value;
                const projectReference = filterProjectReference.value;
                const status = filterStatus.value;
                const make = filterMake.value;
                const filtered = procurementData.filter((item) => {
                    var _a, _b;
                    const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                    const projectData = projects.find((p) => p.Title === projectTitle);
                    const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) ||
                        [];
                    if (project && projectTitle !== project)
                        return false;
                    if (customer && (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) !== customer)
                        return false;
                    if (projectManager) {
                        const managerName = typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                            ? (_b = projectData.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                            : projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager;
                        if (managerName !== projectManager)
                            return false;
                    }
                    if (groupHead && (projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead) !== groupHead)
                        return false;
                    if (projectReference &&
                        (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) !== projectReference)
                        return false;
                    if (status &&
                        !procurementItems.some((item) => item.Status === status))
                        return false;
                    if (make && !procurementItems.some((item) => item.Make === make))
                        return false;
                    return true;
                });
                currentPage = 1;
                renderTable(filtered);
            };
            const clearFilters = () => {
                filterProject.value = "";
                filterCustomer.value = "";
                filterProjectManager.value = "";
                filterGroupHead.value = "";
                filterProjectReference.value = "";
                filterStatus.value = "";
                filterMake.value = "";
                currentPage = 1;
                renderTable(procurementData);
                toast("Filters cleared", "success");
            };
            const formatDate = (dateString) => {
                if (!dateString)
                    return "";
                try {
                    const date = new Date(dateString);
                    return isNaN(date.getTime())
                        ? dateString
                        : date.toLocaleDateString("en-GB");
                }
                catch (_a) {
                    return dateString;
                }
            };
            const formatDateForInput = (dateString) => {
                if (!dateString)
                    return "";
                try {
                    const date = new Date(dateString);
                    return isNaN(date.getTime()) ? "" : date.toISOString().split("T")[0];
                }
                catch (_a) {
                    return "";
                }
            };
            const formatDateForExcel = (dateString) => {
                if (!dateString)
                    return "";
                try {
                    const date = new Date(dateString);
                    if (isNaN(date.getTime()))
                        return "";
                    return date.toISOString().split("T")[0];
                }
                catch (_a) {
                    return "";
                }
            };
            const escapeHtml = (unsafe) => {
                const safeString = String(unsafe);
                return safeString
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;")
                    .replace(/"/g, "&quot;")
                    .replace(/'/g, "&#039;");
            };
            const openEditForm = (id) => {
                const record = procurementData.find((d) => d.Id === id);
                if (!record) {
                    toast("Record not found", "error");
                    return;
                }
                editingId = id;
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formTitle.textContent = "Edit Procurement Plan";
                existingDataAlert.style.display = "block";
                resetDescriptionFilter();
                formSection.classList.remove("view-mode");
                projectSelect.disabled = true;
                const projectId = record[`${fieldMappings.Project}Id`] ||
                    (record[fieldMappings.Project] && record[fieldMappings.Project].Id);
                if (projectId) {
                    projectSelect.value = projectId.toString();
                    const project = projects.find((p) => p.Id === projectId);
                    if (project) {
                        renderProjectDetails(project);
                    }
                    onProjectChange(projectId, true);
                }
                procurementContainer.innerHTML = "";
                const procurementItems = parseProcurementForecast(record[fieldMappings.ProcurementForecast]) ||
                    [];
                procurementItems.forEach((item) => {
                    procurementContainer.appendChild(createProcurementRow(item));
                });
                const addMoreBtn = document.querySelector(".add-procurement-row");
                if (addMoreBtn)
                    addMoreBtn.style.display = "block";
                const submitBtn = form.querySelector('button[type="submit"]');
                if (submitBtn) {
                    submitBtn.style.display = "inline-flex";
                    submitBtn.textContent = "💾 Save Procurement";
                }
                addImportButtonToEditMode(projectId, record);
            };
            const renderTable = (data) => {
                const start = (currentPage - 1) * pageSize;
                const pageData = data.slice(start, start + pageSize);
                tableBody.innerHTML = "";
                if (pageData.length > 0) {
                    pageData.forEach((item) => {
                        var _a, _b;
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        const projectData = projects.find((p) => p.Title === projectTitle);
                        const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) || [];
                        const statusCounts = calculateProjectStatusCounts(procurementItems);
                        const totalLineItems = procurementItems.length;
                        const row = document.createElement("tr");
                        row.innerHTML = `
              <td>
                <span class="clickable-link" data-id="${item.Id}" style="color: #0078d4; cursor: pointer; text-decoration: none; font-weight: 600; border-bottom: 1px solid transparent;">${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "")}</span>
              </td>
              <td>
                <span class="clickable-link" data-id="${item.Id}" style="color: #0078d4; cursor: pointer; text-decoration: none; font-weight: 600; border-bottom: 1px solid transparent;">${escapeHtml(projectTitle)}</span>
              </td>
              <td>${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "")}</td>
              <td>${escapeHtml(typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                            ? ((_b = projectData.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title) || ""
                            : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "")}</td>
              <td>${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead) || "")}</td>
              <td class="total-line-items-col" style="text-align: center; font-weight: 600; background-color: #f0f7ff;">${totalLineItems}</td>
              <td style="text-align: center; font-weight: bold; background-color: #e3f2fd;">${statusCounts["RFQ Pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #fff3e0;">${statusCounts["Offer Pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #e8f5e8;">${statusCounts["TC pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #fce4ec;">${statusCounts["Order Closure pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #f3e5f5;">${statusCounts["PR pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #e0f2f1;">${statusCounts["PO pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #fff8e1;">${statusCounts["Document Approval Pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #e8f5e8;">${statusCounts["MFC Pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #fff3e0;">${statusCounts["Inspection Pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #e3f2fd;">${statusCounts["Dispatch Pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #fce4ec;">${statusCounts["Delivery Pending"] || 0}</td>
              <td style="text-align: center; font-weight: bold; background-color: #e8f5e8;">${statusCounts["Delivered"] || 0}</td>
              <td class="actions-cell">
                <button class="view-btn" data-id="${item.Id}">👁️</button>
                <button class="delete-btn" data-id="${item.Id}">🗑️</button>
              </td>
            `;
                        tableBody.appendChild(row);
                    });
                    document.querySelectorAll(".clickable-link").forEach((link) => {
                        link.addEventListener("click", function (e) {
                            e.preventDefault();
                            const target = e.currentTarget;
                            const id = parseInt(target.dataset.id || "0");
                            if (id > 0) {
                                openEditForm(id);
                            }
                        });
                        const el = link;
                        el.addEventListener("mouseenter", function () {
                            this.style.color = "#0056b3";
                            this.style.borderBottom = "1px solid #0056b3";
                        });
                        el.addEventListener("mouseleave", function () {
                            this.style.color = "#0078d4";
                            this.style.borderBottom = "1px solid transparent";
                        });
                    });
                }
                else {
                    const noDataRow = document.createElement("tr");
                    noDataRow.innerHTML = `<td colspan="19" style="text-align:center;padding:20px;">No records found</td>`;
                    tableBody.appendChild(noDataRow);
                }
                document.querySelectorAll(".view-btn").forEach((btn) => {
                    btn.addEventListener("click", () => viewData(Number(btn.dataset.id)));
                });
                document.querySelectorAll(".delete-btn").forEach((btn) => {
                    btn.addEventListener("click", () => deleteData(Number(btn.dataset.id)));
                });
                renderPagination(Math.ceil(data.length / pageSize));
            };
            const renderPagination = (totalPages) => {
                paginationDiv.innerHTML = "";
                if (totalPages <= 1)
                    return;
                for (let i = 1; i <= totalPages; i++) {
                    const btn = document.createElement("button");
                    btn.textContent = i.toString();
                    btn.className = i === currentPage ? "active" : "";
                    btn.addEventListener("click", () => {
                        currentPage = i;
                        renderTable(procurementData);
                    });
                    paginationDiv.appendChild(btn);
                }
            };
            const deleteData = async (id) => {
                const result = await Swal.fire({
                    title: "Are you sure?",
                    text: "You won't be able to revert this!",
                    icon: "warning",
                    showCancelButton: true,
                    confirmButtonColor: "#0078d4",
                    cancelButtonColor: "#d33",
                    confirmButtonText: "Yes, delete it!",
                });
                if (result.isConfirmed) {
                    try {
                        const digest = await getRequestDigest();
                        const response = await fetch(`${apiBase}(${id})`, {
                            method: "DELETE",
                            headers: {
                                Accept: "application/json;odata=verbose",
                                "X-RequestDigest": digest,
                                "IF-MATCH": "*",
                            },
                        });
                        if (response.ok) {
                            toast("Deleted successfully", "success");
                            await fetchData();
                        }
                        else {
                            throw new Error(`Delete failed: ${response.status}`);
                        }
                    }
                    catch (error) {
                        console.error("Delete error:", error);
                        toast("Delete failed", "error");
                    }
                }
            };
            const viewData = (id) => {
                var _a;
                const record = procurementData.find((d) => d.Id === id);
                if (!record) {
                    toast("Record not found", "error");
                    return;
                }
                editingId = id;
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formTitle.textContent = "View Procurement";
                existingDataAlert.style.display = "none";
                formSection.classList.add("view-mode");
                const projectId = record[`${fieldMappings.Project}Id`] ||
                    (record[fieldMappings.Project] && record[fieldMappings.Project].Id);
                if (projectId) {
                    const project = projects.find((p) => p.Id === projectId);
                    if (project) {
                        projectSelect.value = projectId.toString();
                        renderProjectDetails(project);
                    }
                }
                projectSelect.disabled = true;
                procurementContainer.innerHTML = "";
                const procurementItems = parseProcurementForecast(record[fieldMappings.ProcurementForecast]) ||
                    [];
                procurementItems.forEach((item) => {
                    procurementContainer.appendChild(createViewProcurementRow(item));
                });
                const addMoreBtn = document.querySelector(".add-procurement-row");
                if (addMoreBtn)
                    addMoreBtn.style.display = "none";
                const submitBtn = form.querySelector('button[type="submit"]');
                if (submitBtn) {
                    submitBtn.style.display = "none";
                }
                const existingExportBtns = document.querySelectorAll(".export-project-excel-btn, .export-top-btn");
                existingExportBtns.forEach((btn) => btn.remove());
                const exportBtn = document.createElement("button");
                exportBtn.type = "button";
                exportBtn.className = "btn-primary export-project-excel-btn";
                exportBtn.textContent = "📊 Export BOM Excel";
                exportBtn.style.marginBottom = "15px";
                exportBtn.style.display = "block";
                exportBtn.addEventListener("click", () => exportProjectToExcel(projectId, record));
                const procurementContainerTable = document.getElementById("procurementContainerTable");
                (_a = procurementContainerTable.parentNode) === null || _a === void 0 ? void 0 : _a.insertBefore(exportBtn, procurementContainerTable);
                procurementContainerTable.style.display = "block";
            };
            // ============================================ //
            // DESCRIPTION FILTER - EDIT MODE             //
            // ============================================ //
            const filterEditTableByDescription = () => {
                const searchText = ((descriptionFilter === null || descriptionFilter === void 0 ? void 0 : descriptionFilter.value) || "").trim().toLowerCase();
                const rows = Array.from(procurementContainer.querySelectorAll("tr"));
                rows.forEach((row) => {
                    const descriptionInput = row.querySelector('input[name="description"]');
                    if (!descriptionInput)
                        return;
                    const description = descriptionInput.value.trim().toLowerCase();
                    row.style.display =
                        !searchText || description.includes(searchText) ? "" : "none";
                });
            };
            const resetDescriptionFilter = () => {
                if (descriptionFilter) {
                    descriptionFilter.value = "";
                }
                const rows = Array.from(procurementContainer.querySelectorAll("tr"));
                rows.forEach((row) => {
                    row.style.display = "";
                });
            };
            const createViewProcurementRow = (data) => {
                const row = document.createElement("tr");
                const getDisplayValue = (field) => {
                    if (!data)
                        return "";
                    const value = data[field];
                    if (value === null || value === undefined || value === "")
                        return "";
                    if (field.includes("Plan") ||
                        field.includes("Actual") ||
                        field.includes("Forecast")) {
                        return formatDate(String(value));
                    }
                    if ((field === "FCMQuantity" || field === "EnggQuantity") &&
                        typeof value === "number") {
                        return value.toString();
                    }
                    return String(value);
                };
                const statusValue = getDisplayValue("Status");
                const descriptionValue = escapeHtml(getDisplayValue("Description"));
                row.innerHTML = `
          <td class="sticky-col col-1">
            <div class="view-field description-view" title="${descriptionValue}">
              ${descriptionValue}
            </div>
          </td>
          <td class="sticky-col col-2"><div class="view-field">${getDisplayValue("FCMQuantity")}</div></td>
          <td class="sticky-col col-3"><div class="view-field">${getDisplayValue("EnggQuantity")}</div></td>
          <td class="sticky-col col-4"><div class="view-field status-badge" data-status="${escapeHtml(statusValue)}">${escapeHtml(statusValue)}</div></td>
          <td class="sticky-col col-5"><div class="view-field">${escapeHtml(getDisplayValue("Make"))}</div></td>
          <td class="sticky-col col-6"><div class="view-field">${escapeHtml(getDisplayValue("Model"))}</div></td>
          
          <td><div class="view-field">${getDisplayValue("RfqPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("RfqActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("RfqForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("OfferPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("OfferActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("OfferForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("TechPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("TechActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("TechForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("OrderPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("OrderActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("OrderForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("PRPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("PRActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("PRForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("POPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("POActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("POForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("DocPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("DocActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("DocForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("MfgPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("MfgActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("MfgForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("InspPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("InspActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("InspForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("DispPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("DispActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("DispForecast")}</div></td>
          
          <td><div class="view-field">${getDisplayValue("DeliveryPlan")}</div></td>
          <td><div class="view-field">${getDisplayValue("DeliveryActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("DeliveryForecast")}</div></td>
          
          <td><div class="view-field remarks-view">${escapeHtml(getDisplayValue("Remarks"))}</div></td>
          <td></td>
        `;
                return row;
            };
            // ============================================ //
            // IMPORT FUNCTIONS (simplified)                //
            // ============================================ //
            const collectValidationErrors = (data, headers) => {
                const errors = [];
                if (!data || data.length === 0) {
                    errors.push({
                        type: 'error',
                        message: 'No data found in the Excel file.'
                    });
                    return errors;
                }
                const requiredColumns = ['description', 'status', 'make'];
                const missingColumns = requiredColumns.filter(col => !headers.some(h => h.toLowerCase().includes(col) || h.toLowerCase() === col));
                if (missingColumns.length > 0) {
                    errors.push({
                        type: 'error',
                        message: `Missing required columns: ${missingColumns.join(', ')}`
                    });
                }
                return errors;
            };
            const showValidationErrors = (errors, container) => {
                const existingErrorList = container.querySelector('.error-list');
                if (existingErrorList)
                    existingErrorList.remove();
                if (errors.length === 0)
                    return;
                const errorList = document.createElement('div');
                errorList.className = 'error-list show';
                const criticalErrors = errors.filter(e => e.type === 'error');
                const warnings = errors.filter(e => e.type === 'warning');
                let html = `
          <div class="error-title">
            ⚠️ ${criticalErrors.length > 0 ? `${criticalErrors.length} Critical Error(s) Found` : 'Validation Issues Found'}
          </div>
          <ul class="error-items">
        `;
                criticalErrors.forEach(err => {
                    html += `<li class="error-critical">${err.message}</li>`;
                });
                warnings.forEach(err => {
                    html += `<li class="error-warning">${err.message}</li>`;
                });
                html += `</ul>`;
                if (criticalErrors.length > 0) {
                    html += `
            <div style="margin-top: 10px; padding: 8px 12px; background: #dc3545; color: white; border-radius: 4px; font-size: 13px; font-weight: 600;">
              ❌ ${criticalErrors.length} critical error(s) must be fixed before importing
            </div>
          `;
                }
                errorList.innerHTML = html;
                const previewSection = container.querySelector('#previewSection');
                if (previewSection) {
                    previewSection.insertAdjacentElement('afterend', errorList);
                }
                else {
                    container.appendChild(errorList);
                }
            };
            const parseExcelFile = (file) => {
                return new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        var _a;
                        try {
                            if (!((_a = e.target) === null || _a === void 0 ? void 0 : _a.result)) {
                                reject(new Error("Failed to read file"));
                                return;
                            }
                            const data = new Uint8Array(e.target.result);
                            const workbook = XLSX.read(data, {
                                type: "array",
                                cellDates: true,
                                raw: false,
                            });
                            const sheetName = workbook.SheetNames[0];
                            const worksheet = workbook.Sheets[sheetName];
                            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                                header: 1,
                                defval: "",
                            });
                            if (jsonData.length < 2) {
                                return resolve([]);
                            }
                            let headers;
                            let dataStartRow = 1;
                            if (jsonData.length >= 2 &&
                                Array.isArray(jsonData[0]) &&
                                Array.isArray(jsonData[1])) {
                                const firstRow = jsonData[0];
                                const secondRow = jsonData[1];
                                if (firstRow.some((cell) => cell && typeof cell === "string" && cell.includes("/")) &&
                                    secondRow.some((cell) => cell &&
                                        typeof cell === "string" &&
                                        (cell.toLowerCase().includes("desc") ||
                                            cell.toLowerCase().includes("status")))) {
                                    headers = secondRow;
                                    dataStartRow = 2;
                                }
                                else {
                                    headers = firstRow;
                                }
                            }
                            else {
                                headers = jsonData[0];
                            }
                            const rows = [];
                            const headerMap = {};
                            headers.forEach((header, idx) => {
                                if (header && typeof header === "string") {
                                    const cleanHeader = header.toString().trim().toLowerCase();
                                    if (cleanHeader.includes("desc") ||
                                        cleanHeader.includes("item")) {
                                        headerMap[cleanHeader] = "description";
                                    }
                                    else if (cleanHeader.includes("vendor") ||
                                        cleanHeader.includes("supplier")) {
                                        headerMap[cleanHeader] = "vendor";
                                    }
                                    else if (cleanHeader.includes("fcm") &&
                                        (cleanHeader.includes("qty") ||
                                            cleanHeader.includes("quantity"))) {
                                        headerMap[cleanHeader] = "fcmquantity";
                                    }
                                    else if (cleanHeader.includes("engg") &&
                                        (cleanHeader.includes("qty") ||
                                            cleanHeader.includes("quantity"))) {
                                        headerMap[cleanHeader] = "enggquantity";
                                    }
                                    else if (cleanHeader.includes("status")) {
                                        headerMap[cleanHeader] = "status";
                                    }
                                    else if (cleanHeader.includes("make") ||
                                        cleanHeader.includes("manufacturer")) {
                                        headerMap[cleanHeader] = "make";
                                    }
                                    else if (cleanHeader.includes("model") ||
                                        cleanHeader.includes("part")) {
                                        headerMap[cleanHeader] = "model";
                                    }
                                    else if (cleanHeader.includes("remark")) {
                                        headerMap[cleanHeader] = "remarks";
                                    }
                                    else {
                                        headerMap[cleanHeader] = cleanHeader;
                                    }
                                }
                            });
                            for (let i = dataStartRow; i < jsonData.length; i++) {
                                const rowData = jsonData[i];
                                if (!rowData ||
                                    rowData.every((cell) => cell === "" || cell === null || cell === undefined)) {
                                    continue;
                                }
                                const row = {};
                                headers.forEach((header, idx) => {
                                    if (header &&
                                        rowData[idx] !== undefined &&
                                        rowData[idx] !== null &&
                                        rowData[idx] !== "") {
                                        const cleanHeader = header
                                            .toString()
                                            .trim()
                                            .toLowerCase();
                                        const mappedHeader = headerMap[cleanHeader] || cleanHeader;
                                        let value = rowData[idx];
                                        if (typeof value === "object" && value instanceof Date) {
                                            value = value.toISOString().split("T")[0];
                                        }
                                        else if (typeof value === "number") {
                                            value = value.toString();
                                        }
                                        else if (typeof value === "string") {
                                            value = value.trim();
                                        }
                                        row[mappedHeader] = value;
                                    }
                                });
                                if (row.description || row.make || row.status || row.vendor) {
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
                        console.error("FileReader error:", error);
                        reject(new Error("Failed to read file"));
                    };
                    reader.onabort = () => {
                        reject(new Error("File reading was aborted"));
                    };
                    try {
                        reader.readAsArrayBuffer(file);
                    }
                    catch (error) {
                        reject(new Error(`Cannot read file: ${error}`));
                    }
                });
            };
            const showImportPreview = (data, errors, container) => {
                const errorRows = new Set(errors.filter(e => e.row).map(e => e.row));
                const previewHTML = `
          <div style="padding: 15px; background: #f8f9fa; border-radius: 8px;">
            <h4 style="color: #28a745; margin-bottom: 15px;">📋 Data Preview</h4>
            <p><strong>Found ${data.length} records:</strong></p>
            <div class="preview-table-container">
              <table class="preview-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Description</th>
                    <th>Status</th>
                    <th>Make</th>
                    <th>Model</th>
                    <th>FCM Qty</th>
                    <th>Engg Qty</th>
                    <th>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  ${data
                    .slice(0, 20)
                    .map((row, index) => {
                    const rowNum = index + 1;
                    const hasError = errorRows.has(rowNum);
                    return `
                        <tr class="${hasError ? 'row-invalid' : 'row-valid'}">
                          <td style="text-align: center;">${rowNum}</td>
                          <td>${row.description || "N/A"}</td>
                          <td>${row.status || "N/A"}</td>
                          <td>${row.make || "N/A"}</td>
                          <td>${row.model || "N/A"}</td>
                          <td style="text-align: right;">${row.fcmquantity || "N/A"}</td>
                          <td style="text-align: right;">${row.enggquantity || "N/A"}</td>
                          <td>${row.remarks || "N/A"}</td>
                        </tr>
                      `;
                })
                    .join("")}
                </tbody>
              </table>
            </div>
            ${data.length > 20 ? `<p style="margin-top: 10px; font-style: italic; color: #6c757d;">... and ${data.length - 20} more records</p>` : ""}
            <div style="margin-top: 15px; padding: 12px; background: #e7f3ff; border-radius: 5px;">
              <p style="margin: 0; color: #0056b3; font-size: 14px;">
                <strong>⚠️ Important:</strong> This will replace existing procurement data for this project.
              </p>
            </div>
          </div>
        `;
                const importPreview = container.querySelector('#importPreview');
                if (importPreview) {
                    importPreview.innerHTML = previewHTML;
                }
            };
            const addImportButtonToEditMode = (projectId, record) => {
                var _a;
                const existingImportBtns = document.querySelectorAll(".import-excel-edit-btn");
                existingImportBtns.forEach((btn) => btn.remove());
                const importBtn = document.createElement("button");
                importBtn.type = "button";
                importBtn.className = "btn-primary import-excel-edit-btn";
                importBtn.textContent = "📥 UPLOAD";
                importBtn.style.marginBottom = "15px";
                importBtn.style.marginRight = "10px";
                importBtn.style.display = "inline-block";
                importBtn.addEventListener("click", () => importExcelForEdit(projectId, record));
                const procurementContainerTable = document.getElementById("procurementContainerTable");
                (_a = procurementContainerTable.parentNode) === null || _a === void 0 ? void 0 : _a.insertBefore(importBtn, procurementContainerTable);
            };
            const importExcelForEdit = (projectId, record) => {
                const modalHTML = `
          <div class="import-modal">
            <div class="import-modal-content">
              <div class="import-modal-header">
                <h3>Import Procurement Data from Excel</h3>
                <button class="close-modal">&times;</button>
              </div>

              <div class="import-section">
                <h4>Step 1: Download Template</h4>
                <p>Download the template file with the correct structure for import.</p>
                <button class="btn-primary" id="downloadEditTemplate">⬇️ Download Template</button>
              </div>

              <div class="import-section">
                <h4>Step 2: Upload Excel File</h4>
                <div class="file-upload-area" id="fileUploadArea">
                  <div class="upload-icon">📁</div>
                  <p><strong>Click to upload or drag & drop Excel file</strong></p>
                  <p><small>Supported format: .xlsx, .xls (Max: 10MB)</small></p>
                  <input type="file" id="excelFile" class="file-input" accept=".xlsx,.xls">
                  <div class="file-info" id="fileInfo"></div>
                </div>
              </div>

              <div class="import-section" id="previewSection" style="display:none;">
                <h4>Step 3: Preview & Validate</h4>
                <div class="import-preview" id="importPreview"></div>
                <div class="import-actions">
                  <button class="btn-secondary" id="cancelImport">Cancel</button>
                  <button class="btn-primary" id="confirmImport" disabled>Import Data</button>
                </div>
                <div class="import-status" id="importStatus"></div>
              </div>
            </div>
          </div>
        `;
                document.body.insertAdjacentHTML("beforeend", modalHTML);
                const modal = document.querySelector(".import-modal");
                const fileUploadArea = document.getElementById("fileUploadArea");
                const excelFileInput = document.getElementById("excelFile");
                const fileInfo = document.getElementById("fileInfo");
                const previewSection = document.getElementById("previewSection");
                const confirmImport = document.getElementById("confirmImport");
                const cancelImport = document.getElementById("cancelImport");
                const closeModal = document.querySelector(".close-modal");
                const importStatus = document.getElementById("importStatus");
                const downloadTemplateBtn = document.getElementById("downloadEditTemplate");
                let importData = [];
                let importHeaders = [];
                fileUploadArea.addEventListener("click", () => excelFileInput.click());
                fileUploadArea.addEventListener("dragover", (e) => {
                    e.preventDefault();
                    fileUploadArea.classList.add("dragover");
                });
                fileUploadArea.addEventListener("dragleave", () => {
                    fileUploadArea.classList.remove("dragover");
                });
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
                    const errorList = modal.querySelector('.error-list');
                    if (errorList)
                        errorList.remove();
                    importStatus.className = 'import-status';
                    importStatus.style.display = 'none';
                    confirmImport.disabled = true;
                    if (!file.name.toLowerCase().match(/\.(xlsx|xls)$/)) {
                        fileInfo.textContent = 'Please select an Excel file (.xlsx or .xls)';
                        fileInfo.className = 'file-info error';
                        return;
                    }
                    if (file.size > 10 * 1024 * 1024) {
                        fileInfo.textContent = 'File size too large. Please use files under 10MB.';
                        fileInfo.className = 'file-info error';
                        return;
                    }
                    fileInfo.textContent = `Selected: ${file.name} (${(file.size / 1024).toFixed(2)} KB)`;
                    fileInfo.className = 'file-info';
                    fileInfo.style.color = '#0078d4';
                    try {
                        const parsedData = await parseExcelFile(file);
                        if (!parsedData || parsedData.length === 0) {
                            fileInfo.textContent = 'No valid data found in the Excel file.';
                            fileInfo.className = 'file-info error';
                            return;
                        }
                        if (parsedData.length > 0) {
                            importHeaders = Object.keys(parsedData[0]);
                        }
                        importData = parsedData;
                        const validationErrors = collectValidationErrors(parsedData, importHeaders);
                        showImportPreview(parsedData, validationErrors, modal);
                        if (validationErrors.length > 0) {
                            showValidationErrors(validationErrors, modal);
                            const hasCriticalErrors = validationErrors.some(e => e.type === 'error');
                            if (hasCriticalErrors) {
                                confirmImport.disabled = true;
                                importStatus.className = 'import-status error';
                                importStatus.style.display = 'block';
                                importStatus.textContent = '❌ Critical errors found. Please fix the issues and try again.';
                                return;
                            }
                        }
                        confirmImport.disabled = false;
                        previewSection.style.display = "block";
                        importStatus.className = 'import-status success';
                        importStatus.style.display = 'block';
                        importStatus.textContent = `✅ Found ${parsedData.length} valid row(s) to import.`;
                    }
                    catch (error) {
                        console.error("Excel parsing error:", error);
                        fileInfo.textContent = `Error reading Excel file: ${error}`;
                        fileInfo.className = 'file-info error';
                    }
                };
                downloadTemplateBtn.addEventListener("click", () => {
                    var _a;
                    const project = projects.find((p) => p.Id === projectId);
                    const currentProcurementItems = parseProcurementForecast(record[fieldMappings.ProcurementForecast]) || [];
                    const headers = [
                        "Description",
                        "Status",
                        "Make",
                        "Model",
                        "FCMQuantity",
                        "EnggQuantity",
                        "Remarks",
                        "RfqPlan",
                        "RfqActual",
                        "RfqForecast",
                        "OfferPlan",
                        "OfferActual",
                        "OfferForecast",
                        "TechPlan",
                        "TechActual",
                        "TechForecast",
                        "OrderPlan",
                        "OrderActual",
                        "OrderForecast",
                        "PRPlan",
                        "PRActual",
                        "PRForecast",
                        "POPlan",
                        "POActual",
                        "POForecast",
                        "DocPlan",
                        "DocActual",
                        "DocForecast",
                        "MfgPlan",
                        "MfgActual",
                        "MfgForecast",
                        "InspPlan",
                        "InspActual",
                        "InspForecast",
                        "DispPlan",
                        "DispActual",
                        "DispForecast",
                        "DeliveryPlan",
                        "DeliveryActual",
                        "DeliveryForecast",
                    ];
                    const wsData = [headers];
                    if (currentProcurementItems.length > 0) {
                        currentProcurementItems.forEach((item) => {
                            var _a, _b;
                            const row = [
                                item.Description || "",
                                item.Status || "RFQ Pending",
                                item.Make || "",
                                item.Model || "",
                                ((_a = item.FCMQuantity) === null || _a === void 0 ? void 0 : _a.toString()) || "",
                                ((_b = item.EnggQuantity) === null || _b === void 0 ? void 0 : _b.toString()) || "",
                                item.Remarks || "",
                                formatDateForExcel(item.RfqPlan) || "",
                                formatDateForExcel(item.RfqActual) || "",
                                formatDateForExcel(item.RfqForecast) || "",
                                formatDateForExcel(item.OfferPlan) || "",
                                formatDateForExcel(item.OfferActual) || "",
                                formatDateForExcel(item.OfferForecast) || "",
                                formatDateForExcel(item.TechPlan) || "",
                                formatDateForExcel(item.TechActual) || "",
                                formatDateForExcel(item.TechForecast) || "",
                                formatDateForExcel(item.OrderPlan) || "",
                                formatDateForExcel(item.OrderActual) || "",
                                formatDateForExcel(item.OrderForecast) || "",
                                formatDateForExcel(item.PRPlan) || "",
                                formatDateForExcel(item.PRActual) || "",
                                formatDateForExcel(item.PRForecast) || "",
                                formatDateForExcel(item.POPlan) || "",
                                formatDateForExcel(item.POActual) || "",
                                formatDateForExcel(item.POForecast) || "",
                                formatDateForExcel(item.DocPlan) || "",
                                formatDateForExcel(item.DocActual) || "",
                                formatDateForExcel(item.DocForecast) || "",
                                formatDateForExcel(item.MfgPlan) || "",
                                formatDateForExcel(item.MfgActual) || "",
                                formatDateForExcel(item.MfgForecast) || "",
                                formatDateForExcel(item.InspPlan) || "",
                                formatDateForExcel(item.InspActual) || "",
                                formatDateForExcel(item.InspForecast) || "",
                                formatDateForExcel(item.DispPlan) || "",
                                formatDateForExcel(item.DispActual) || "",
                                formatDateForExcel(item.DispForecast) || "",
                                formatDateForExcel(item.DeliveryPlan) || "",
                                formatDateForExcel(item.DeliveryActual) || "",
                                formatDateForExcel(item.DeliveryForecast) || "",
                            ];
                            wsData.push(row);
                        });
                    }
                    else {
                        const emptyRow = Array(headers.length).fill("");
                        emptyRow[1] = "RFQ Pending";
                        wsData.push(emptyRow);
                    }
                    const ws = XLSX.utils.aoa_to_sheet(wsData);
                    const headerStyle = {
                        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
                        fill: { patternType: "solid", fgColor: { rgb: "2F5597" } },
                        alignment: {
                            horizontal: "center",
                            vertical: "center",
                            wrapText: true,
                        },
                        border: {
                            top: { style: "medium", color: { rgb: "000000" } },
                            bottom: { style: "medium", color: { rgb: "000000" } },
                            left: { style: "medium", color: { rgb: "000000" } },
                            right: { style: "medium", color: { rgb: "000000" } },
                        },
                    };
                    for (let col = 0; col < headers.length; col++) {
                        const address = XLSX.utils.encode_cell({ r: 0, c: col });
                        ws[address].s = headerStyle;
                    }
                    ws["!cols"] = [
                        { wch: 30 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
                        { wch: 12 }, { wch: 12 }, { wch: 20 },
                        ...Array(33).fill({ wch: 10 })
                    ];
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Procurement Data");
                    const timestamp = new Date().toISOString().split("T")[0];
                    const fileName = `Procurement_Template_${((_a = project === null || project === void 0 ? void 0 : project.Title) === null || _a === void 0 ? void 0 : _a.replace(/[^a-zA-Z0-9]/g, "_")) || "Project"}_${timestamp}.xlsx`;
                    XLSX.writeFile(wb, fileName);
                    toast(`Template downloaded with ${currentProcurementItems.length} existing items`, "success");
                });
                const performImport = async () => {
                    confirmImport.disabled = true;
                    confirmImport.textContent = "Importing...";
                    importStatus.className = "import-status info";
                    importStatus.style.display = "block";
                    importStatus.textContent = "⏳ Processing procurement data...";
                    try {
                        const procurementItems = importData.map((row) => {
                            const item = {
                                Description: row.description || "",
                                FCMQuantity: row.fcmquantity
                                    ? parseFloat(row.fcmquantity)
                                    : null,
                                EnggQuantity: row.enggquantity
                                    ? parseFloat(row.enggquantity)
                                    : null,
                                Status: row.status || "RFQ Pending",
                                Make: row.make || row.manufacturer || "",
                                Model: row.model || row.partnumber || "",
                                Remarks: row.remarks || "",
                                RfqPlan: row.rfqplan || "",
                                RfqActual: row.rfqactual || "",
                                RfqForecast: row.rfqforecast || "",
                                OfferPlan: row.offerplan || "",
                                OfferActual: row.offeractual || "",
                                OfferForecast: row.offerforecast || "",
                                TechPlan: row.techplan || "",
                                TechActual: row.techactual || "",
                                TechForecast: row.techforecast || "",
                                OrderPlan: row.orderplan || "",
                                OrderActual: row.orderactual || "",
                                OrderForecast: row.orderforecast || "",
                                PRPlan: row.prplan || "",
                                PRActual: row.practual || "",
                                PRForecast: row.prforecast || "",
                                POPlan: row.poplan || "",
                                POActual: row.poactual || "",
                                POForecast: row.poforecast || "",
                                DocPlan: row.docplan || "",
                                DocActual: row.docactual || "",
                                DocForecast: row.docforecast || "",
                                MfgPlan: row.mfgplan || "",
                                MfgActual: row.mfgactual || "",
                                MfgForecast: row.mfgforecast || "",
                                InspPlan: row.inspplan || "",
                                InspActual: row.inspactual || "",
                                InspForecast: row.inspforecast || "",
                                DispPlan: row.dispplan || "",
                                DispActual: row.dispactual || "",
                                DispForecast: row.dispforecast || "",
                                DeliveryPlan: row.deliveryplan || "",
                                DeliveryActual: row.deliveryactual || "",
                                DeliveryForecast: row.deliveryforecast || "",
                            };
                            item.Status = getStatusFromActualDates(item);
                            return item;
                        });
                        await submitImportedData(procurementItems, projectId);
                        importStatus.className = "import-status success";
                        importStatus.textContent = `✅ Successfully imported ${procurementItems.length} procurement items!`;
                        setTimeout(() => {
                            modal.remove();
                            toast(`Successfully imported ${procurementItems.length} procurement items!`, "success");
                            resetForm();
                            fetchData();
                        }, 1000);
                    }
                    catch (error) {
                        console.error("Import error:", error);
                        importStatus.className = "import-status error";
                        importStatus.innerHTML = `<strong>❌ Import Failed:</strong> ${error}`;
                        toast("Failed to import and save data", "error");
                        confirmImport.disabled = false;
                        confirmImport.textContent = "Import Data";
                    }
                };
                const submitImportedData = async (procurementItems, projId) => {
                    try {
                        const digest = await getRequestDigest();
                        if (!projId) {
                            throw new Error("No project selected");
                        }
                        const existingRecord = procurementData.find((item) => item[`${fieldMappings.Project}Id`] === projId ||
                            (item[fieldMappings.Project] &&
                                item[fieldMappings.Project].Id === projId));
                        const itemData = {
                            __metadata: { type: listEntityTypeName },
                            [`${fieldMappings.Project}Id`]: projId,
                            [fieldMappings.ProcurementForecast]: JSON.stringify(procurementItems),
                        };
                        const headers = {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest,
                        };
                        let url = apiBase;
                        if (editingId || existingRecord) {
                            const recordId = editingId || existingRecord.Id;
                            url = `${apiBase}(${recordId})`;
                            headers["IF-MATCH"] = "*";
                            headers["X-HTTP-Method"] = "MERGE";
                        }
                        const response = await fetch(url, {
                            method: "POST",
                            headers: headers,
                            body: JSON.stringify(itemData),
                        });
                        if (!response.ok) {
                            const errorText = await response.text();
                            throw new Error(`HTTP error! status: ${response.status}, details: ${errorText}`);
                        }
                    }
                    catch (error) {
                        console.error("Error submitting imported data:", error);
                        throw error;
                    }
                };
                const closeModalHandler = () => {
                    const errorList = modal.querySelector('.error-list');
                    if (errorList)
                        errorList.remove();
                    modal.remove();
                };
                closeModal.addEventListener("click", closeModalHandler);
                cancelImport.addEventListener("click", closeModalHandler);
                modal.addEventListener("click", (e) => {
                    if (e.target === modal)
                        closeModalHandler();
                });
                confirmImport.addEventListener("click", performImport);
            };
            const onProjectChange = (selectedId, isEditing = false) => {
                const project = projects.find((p) => p.Id === selectedId);
                if (!project) {
                    renderProjectDetails(null);
                    return;
                }
                renderProjectDetails(project);
                procurementContainerTable.style.display = "block";
                if (!isEditing) {
                    const existingRecord = procurementData.find((item) => item[`${fieldMappings.Project}Id`] === selectedId ||
                        (item[fieldMappings.Project] &&
                            item[fieldMappings.Project].Id === selectedId));
                    if (existingRecord) {
                        existingDataAlert.style.display = "block";
                        formTitle.textContent = "Edit Procurement Plan";
                        editingId = existingRecord.Id;
                        const procurementItems = parseProcurementForecast(existingRecord[fieldMappings.ProcurementForecast]) || [];
                        procurementContainer.innerHTML = "";
                        procurementItems.forEach((item) => {
                            procurementContainer.appendChild(createProcurementRow(item));
                        });
                    }
                    else {
                        existingDataAlert.style.display = "none";
                        formTitle.textContent = "Add Procurement";
                        editingId = null;
                        procurementContainer.innerHTML = "";
                        procurementContainer.appendChild(createProcurementRow());
                    }
                }
            };
            const createProcurementRow = (data) => {
                const row = document.createElement("tr");
                const getFieldValue = (field) => {
                    if (!data)
                        return "";
                    const value = data[field];
                    if (value === null || value === undefined)
                        return "";
                    if (field.includes("Plan") ||
                        field.includes("Actual") ||
                        field.includes("Forecast")) {
                        return formatDateForInput(String(value));
                    }
                    if ((field === "FCMQuantity" || field === "EnggQuantity") &&
                        typeof value === "number") {
                        return value.toString();
                    }
                    return String(value);
                };
                const initialStatus = getStatusFromActualDates(data);
                const poActualValue = getFieldValue("POActual");
                const isPOActualFilled = isFilled(poActualValue);
                row.innerHTML = `
          <td class="sticky-col col-1"><input type="text" name="description" class="form-control description-input" placeholder="Description" value="${escapeHtml(getFieldValue("Description"))}" required></td>
          <td class="sticky-col col-2"><input type="number" name="fcm_qty" class="form-control number-input" placeholder="FCM Qty" value="${getFieldValue("FCMQuantity")}" step="0.01" min="0"></td>
          <td class="sticky-col col-3">
            <div class="engg-qty-wrapper">
              <input type="number" name="engg_qty" class="form-control number-input engg-qty-input" placeholder="Engg Qty" value="${getFieldValue("EnggQuantity")}" step="0.01" min="0" ${isPOActualFilled ? 'disabled' : ''}>
              ${isPOActualFilled ? '<span class="lock-icon">🔒</span>' : ''}
            </div>
            ${isPOActualFilled ? '<small style="color:#6c757d; font-size:9px; display:block;">Locked </small>' : ''}
          </td>
          <td class="sticky-col col-4">
            <span class="status-text-display ${initialStatus === 'Delivered' ? 'delivered' : ''}" name="status_display">${initialStatus}</span>
            
          </td>
          <td class="sticky-col col-5"><input type="text" name="make" class="form-control" placeholder="Make" value="${escapeHtml(getFieldValue("Make"))}"></td>
          <td class="sticky-col col-6"><input type="text" name="model" class="form-control" placeholder="Model" value="${escapeHtml(getFieldValue("Model"))}"></td>
          <!-- RFQ / Engg -->
          <td><input type="date" name="rfq_plan" class="form-control" value="${getFieldValue("RfqPlan")}"></td>
          <td><input type="date" name="rfq_actual" class="form-control" value="${getFieldValue("RfqActual")}"></td>
          <td><input type="date" name="rfq_forecast" class="form-control" value="${getFieldValue("RfqForecast")}"></td>
          
          <!-- Vendor Offer / SCM -->
          <td><input type="date" name="offer_plan" class="form-control" value="${getFieldValue("OfferPlan")}"></td>
          <td><input type="date" name="offer_actual" class="form-control" value="${getFieldValue("OfferActual")}"></td>
          <td><input type="date" name="offer_forecast" class="form-control" value="${getFieldValue("OfferForecast")}"></td>
          
          <!-- Technical Clearance / ENGG -->
          <td><input type="date" name="tech_plan" class="form-control" value="${getFieldValue("TechPlan")}"></td>
          <td><input type="date" name="tech_actual" class="form-control" value="${getFieldValue("TechActual")}"></td>
          <td><input type="date" name="tech_forecast" class="form-control" value="${getFieldValue("TechForecast")}"></td>
          
          <!-- Order Finalization / SCM -->
          <td><input type="date" name="order_plan" class="form-control" value="${getFieldValue("OrderPlan")}"></td>
          <td><input type="date" name="order_actual" class="form-control" value="${getFieldValue("OrderActual")}"></td>
          <td><input type="date" name="order_forecast" class="form-control" value="${getFieldValue("OrderForecast")}"></td>
          
          <!-- PR / PM -->
          <td><input type="date" name="pr_plan" class="form-control" value="${getFieldValue("PRPlan")}"></td>
          <td><input type="date" name="pr_actual" class="form-control" value="${getFieldValue("PRActual")}"></td>
          <td><input type="date" name="pr_forecast" class="form-control" value="${getFieldValue("PRForecast")}"></td>
          
          <!-- Purchase Order / SCM -->
          <td><input type="date" name="po_plan" class="form-control" value="${getFieldValue("POPlan")}"></td>
          <td><input type="date" name="po_actual" class="form-control" value="${getFieldValue("POActual")}"></td>
          <td><input type="date" name="po_forecast" class="form-control" value="${getFieldValue("POForecast")}"></td>
          
          <!-- Document Approval / ENGG -->
          <td><input type="date" name="doc_plan" class="form-control" value="${getFieldValue("DocPlan")}"></td>
          <td><input type="date" name="doc_actual" class="form-control" value="${getFieldValue("DocActual")}"></td>
          <td><input type="date" name="doc_forecast" class="form-control" value="${getFieldValue("DocForecast")}"></td>
          
          <!-- Manufacturing Clearance / ENGG -->
          <td><input type="date" name="mfg_plan" class="form-control" value="${getFieldValue("MfgPlan")}"></td>
          <td><input type="date" name="mfg_actual" class="form-control" value="${getFieldValue("MfgActual")}"></td>
          <td><input type="date" name="mfg_forecast" class="form-control" value="${getFieldValue("MfgForecast")}"></td>
          
          <!-- Inspection -->
          <td><input type="date" name="insp_plan" class="form-control" value="${getFieldValue("InspPlan")}"></td>
          <td><input type="date" name="insp_actual" class="form-control" value="${getFieldValue("InspActual")}"></td>
          <td><input type="date" name="insp_forecast" class="form-control" value="${getFieldValue("InspForecast")}"></td>
          
          <!-- Material Dispatch -->
          <td><input type="date" name="disp_plan" class="form-control" value="${getFieldValue("DispPlan")}"></td>
          <td><input type="date" name="disp_actual" class="form-control" value="${getFieldValue("DispActual")}"></td>
          <td><input type="date" name="disp_forecast" class="form-control" value="${getFieldValue("DispForecast")}"></td>
          
          <!-- Delivery at Site -->
          <td><input type="date" name="delivery_plan" class="form-control" value="${getFieldValue("DeliveryPlan")}"></td>
          <td><input type="date" name="delivery_actual" class="form-control" value="${getFieldValue("DeliveryActual")}"></td>
          <td><input type="date" name="delivery_forecast" class="form-control" value="${getFieldValue("DeliveryForecast")}"></td>

          <td><input type="text" name="remarks" class="form-control" placeholder="Remarks" value="${escapeHtml(getFieldValue("Remarks"))}"></td>
          <td><button type="button" class="remove-row btn-danger">🗑️</button></td>
        `;
                wireStatusAutoLink(row);
                return row;
            };
            form.addEventListener("submit", async (e) => {
                e.preventDefault();
                if (!projectSelect.value) {
                    toast("Please select a project", "error");
                    return;
                }
                const rows = Array.from(procurementContainer.querySelectorAll("tr"));
                if (rows.length === 0) {
                    toast("Please add at least one procurement item", "error");
                    return;
                }
                try {
                    const digest = await getRequestDigest();
                    const projectId = parseInt(projectSelect.value);
                    const procurementItems = rows.map((row) => {
                        const inputs = row.querySelectorAll("input, select");
                        const item = {
                            Description: "",
                            FCMQuantity: null,
                            EnggQuantity: null,
                            Status: "",
                            Make: "",
                            Model: "",
                            Remarks: "",
                            RfqPlan: "",
                            RfqActual: "",
                            RfqForecast: "",
                            OfferPlan: "",
                            OfferActual: "",
                            OfferForecast: "",
                            TechPlan: "",
                            TechActual: "",
                            TechForecast: "",
                            OrderPlan: "",
                            OrderActual: "",
                            OrderForecast: "",
                            PRPlan: "",
                            PRActual: "",
                            PRForecast: "",
                            POPlan: "",
                            POActual: "",
                            POForecast: "",
                            DocPlan: "",
                            DocActual: "",
                            DocForecast: "",
                            MfgPlan: "",
                            MfgActual: "",
                            MfgForecast: "",
                            InspPlan: "",
                            InspActual: "",
                            InspForecast: "",
                            DispPlan: "",
                            DispActual: "",
                            DispForecast: "",
                            DeliveryPlan: "",
                            DeliveryActual: "",
                            DeliveryForecast: "",
                        };
                        inputs.forEach((input) => {
                            const element = input;
                            const name = element.name;
                            const value = element.value;
                            if (!name)
                                return;
                            switch (name) {
                                case "description":
                                    item.Description = value;
                                    break;
                                case "fcm_qty":
                                    item.FCMQuantity = value ? parseFloat(value) : null;
                                    break;
                                case "engg_qty": {
                                    const enggInput = input;
                                    if (!enggInput.disabled) {
                                        item.EnggQuantity = value ? parseFloat(value) : null;
                                    }
                                    break;
                                }
                                case "make":
                                    item.Make = value;
                                    break;
                                case "model":
                                    item.Model = value;
                                    break;
                                case "remarks":
                                    item.Remarks = value;
                                    break;
                                case "rfq_plan":
                                    item.RfqPlan = value;
                                    break;
                                case "rfq_actual":
                                    item.RfqActual = value;
                                    break;
                                case "rfq_forecast":
                                    item.RfqForecast = value;
                                    break;
                                case "offer_plan":
                                    item.OfferPlan = value;
                                    break;
                                case "offer_actual":
                                    item.OfferActual = value;
                                    break;
                                case "offer_forecast":
                                    item.OfferForecast = value;
                                    break;
                                case "tech_plan":
                                    item.TechPlan = value;
                                    break;
                                case "tech_actual":
                                    item.TechActual = value;
                                    break;
                                case "tech_forecast":
                                    item.TechForecast = value;
                                    break;
                                case "order_plan":
                                    item.OrderPlan = value;
                                    break;
                                case "order_actual":
                                    item.OrderActual = value;
                                    break;
                                case "order_forecast":
                                    item.OrderForecast = value;
                                    break;
                                case "pr_plan":
                                    item.PRPlan = value;
                                    break;
                                case "pr_actual":
                                    item.PRActual = value;
                                    break;
                                case "pr_forecast":
                                    item.PRForecast = value;
                                    break;
                                case "po_plan":
                                    item.POPlan = value;
                                    break;
                                case "po_actual":
                                    item.POActual = value;
                                    break;
                                case "po_forecast":
                                    item.POForecast = value;
                                    break;
                                case "doc_plan":
                                    item.DocPlan = value;
                                    break;
                                case "doc_actual":
                                    item.DocActual = value;
                                    break;
                                case "doc_forecast":
                                    item.DocForecast = value;
                                    break;
                                case "mfg_plan":
                                    item.MfgPlan = value;
                                    break;
                                case "mfg_actual":
                                    item.MfgActual = value;
                                    break;
                                case "mfg_forecast":
                                    item.MfgForecast = value;
                                    break;
                                case "insp_plan":
                                    item.InspPlan = value;
                                    break;
                                case "insp_actual":
                                    item.InspActual = value;
                                    break;
                                case "insp_forecast":
                                    item.InspForecast = value;
                                    break;
                                case "disp_plan":
                                    item.DispPlan = value;
                                    break;
                                case "disp_actual":
                                    item.DispActual = value;
                                    break;
                                case "disp_forecast":
                                    item.DispForecast = value;
                                    break;
                                case "delivery_plan":
                                    item.DeliveryPlan = value;
                                    break;
                                case "delivery_actual":
                                    item.DeliveryActual = value;
                                    break;
                                case "delivery_forecast":
                                    item.DeliveryForecast = value;
                                    break;
                            }
                        });
                        item.Status = getStatusFromActualDates(item);
                        return item;
                    });
                    const existingRecord = procurementData.find((item) => item[`${fieldMappings.Project}Id`] === projectId ||
                        (item[fieldMappings.Project] &&
                            item[fieldMappings.Project].Id === projectId));
                    const itemData = {
                        __metadata: { type: listEntityTypeName },
                        [`${fieldMappings.Project}Id`]: projectId,
                        [fieldMappings.ProcurementForecast]: JSON.stringify(procurementItems),
                    };
                    const headers = {
                        Accept: "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": digest,
                    };
                    let url = apiBase;
                    if (editingId || existingRecord) {
                        const recordId = editingId || existingRecord.Id;
                        url = `${apiBase}(${recordId})`;
                        headers["IF-MATCH"] = "*";
                        headers["X-HTTP-Method"] = "MERGE";
                    }
                    const response = await fetch(url, {
                        method: "POST",
                        headers: headers,
                        body: JSON.stringify(itemData),
                    });
                    if (!response.ok) {
                        const errorText = await response.text();
                        throw new Error(`HTTP error! status: ${response.status}, details: ${errorText}`);
                    }
                    toast(`Procurement data ${editingId || existingRecord ? 'updated' : 'added'} successfully! ${procurementItems.length} item(s) saved`, "success");
                    resetForm();
                    await fetchData();
                }
                catch (error) {
                    console.error("Error saving procurement:", error);
                    toast("Failed to save procurement data", "error");
                }
            });
            const resetForm = () => {
                resetDescriptionFilter();
                editingId = null;
                form.reset();
                formSection.style.display = "none";
                tableSection.style.display = "block";
                procurementContainer.innerHTML = "";
                procurementContainerTable.style.display = "none";
                renderProjectDetails(null);
                existingDataAlert.style.display = "none";
                formTitle.textContent = "Add Procurement";
                projectSelect.disabled = false;
                projectSelect.value = "";
                formSection.classList.remove("view-mode");
                const actionBtns = document.querySelectorAll(".export-project-excel-btn, .export-top-btn, .import-excel-edit-btn");
                actionBtns.forEach((btn) => btn.remove());
                const addMoreBtn = document.querySelector(".add-procurement-row");
                if (addMoreBtn)
                    addMoreBtn.style.display = "block";
                const submitBtn = form.querySelector('button[type="submit"]');
                if (submitBtn) {
                    submitBtn.style.display = "inline-flex";
                }
            };
            const exportProjectToExcel = async (projectId, record) => {
                var _a;
                showExportProgress("Preparing project data for export...");
                try {
                    const currentRecord = procurementData.find((d) => d.Id === record.Id);
                    if (!currentRecord) {
                        hideExportProgress();
                        toast("Record not found", "error");
                        return;
                    }
                    const project = projects.find((p) => p.Id === projectId);
                    if (!project) {
                        hideExportProgress();
                        toast("Project not found", "error");
                        return;
                    }
                    const procurementItems = parseProcurementForecast(currentRecord[fieldMappings.ProcurementForecast]) || [];
                    if (procurementItems.length === 0) {
                        hideExportProgress();
                        toast("No procurement data to export for this project", "warning");
                        return;
                    }
                    const wb = XLSX.utils.book_new();
                    // Sheet 1: Project Summary
                    const projectSummaryData = [
                        ["PROJECT SUMMARY", ""],
                        ["", ""],
                        ["Project Information", ""],
                        ["Project Name", project.Title || ""],
                        ["Customer", project.CustomerName || ""],
                        ["Project Manager", typeof project.ProjectManager === "object" ? ((_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || "" : project.ProjectManager || ""],
                        ["Group Head", project.GroupHead || ""],
                        ["Project Reference", project.ProjectReference || ""],
                        ["Profit Centre", project.PC || ""],
                        ["", ""],
                        ["Procurement Summary", ""],
                        ["Total Procurement Items", procurementItems.length],
                        ["", ""],
                        ["Status-wise Breakdown", ""],
                    ];
                    const statusCounts = calculateProjectStatusCounts(procurementItems);
                    statusOptions.forEach((status) => {
                        projectSummaryData.push([status, statusCounts[status] || 0]);
                    });
                    projectSummaryData.push(["", ""]);
                    projectSummaryData.push(["TOTAL ITEMS", procurementItems.length]);
                    const projectSummaryWs = XLSX.utils.aoa_to_sheet(projectSummaryData);
                    XLSX.utils.book_append_sheet(wb, projectSummaryWs, "Cover Sheet");
                    // Sheet 2: Procurement Details
                    const headerRow1 = [
                        "Sr No", "Description Of Item", "Quantity", "Unit", "Make", "Model",
                        "RFQ / Engg", "", "",
                        "Offer / SCM", "", "",
                        "Technical Clearance / ENGG", "", "",
                        "PR / PM", "", "",
                        "Purchase Order / SCM", "", "",
                        "Manufacturing Clearance / ENGG", "", "",
                        "Status", "Remarks", "WEIGHTAGE"
                    ];
                    const headerRow2 = [
                        "Sr No", "Description Of Item", "Quantity", "Unit", "Make", "Model",
                        "Plan", "Actual", "Forecast",
                        "Plan", "Actual", "Forecast",
                        "Plan", "Actual", "Forecast",
                        "Plan", "Actual", "Forecast",
                        "Plan", "Actual", "Forecast",
                        "Plan", "Actual", "Forecast",
                        "Status", "Remarks", "WEIGHTAGE"
                    ];
                    const procurementWs = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2]);
                    procurementItems.forEach((procItem, index) => {
                        const row = [
                            index + 1,
                            procItem.Description || "",
                            procItem.FCMQuantity || procItem.EnggQuantity || "",
                            "",
                            procItem.Make || "",
                            procItem.Model || "",
                            formatDateForExcel(procItem.RfqPlan) || "",
                            formatDateForExcel(procItem.RfqActual) || "",
                            formatDateForExcel(procItem.RfqForecast) || "",
                            formatDateForExcel(procItem.OfferPlan) || "",
                            formatDateForExcel(procItem.OfferActual) || "",
                            formatDateForExcel(procItem.OfferForecast) || "",
                            formatDateForExcel(procItem.TechPlan) || "",
                            formatDateForExcel(procItem.TechActual) || "",
                            formatDateForExcel(procItem.TechForecast) || "",
                            formatDateForExcel(procItem.PRPlan) || "",
                            formatDateForExcel(procItem.PRActual) || "",
                            formatDateForExcel(procItem.PRForecast) || "",
                            formatDateForExcel(procItem.POPlan) || "",
                            formatDateForExcel(procItem.POActual) || "",
                            formatDateForExcel(procItem.POForecast) || "",
                            formatDateForExcel(procItem.MfgPlan) || "",
                            formatDateForExcel(procItem.MfgActual) || "",
                            formatDateForExcel(procItem.MfgForecast) || "",
                            procItem.Status || "",
                            procItem.Remarks || "",
                            "",
                        ];
                        XLSX.utils.sheet_add_aoa(procurementWs, [row], { origin: `A${index + 3}` });
                    });
                    XLSX.utils.book_append_sheet(wb, procurementWs, "BOM Standard Format");
                    const excelBuffer = XLSX.write(wb, {
                        bookType: "xlsx",
                        type: "array",
                    });
                    const data = new Blob([excelBuffer], {
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });
                    const timestamp = new Date().toISOString().split("T")[0];
                    const fileName = `Procurement_${project.Title.replace(/[^a-zA-Z0-9]/g, "_")}_${timestamp}.xlsx`;
                    saveAs(data, fileName);
                    hideExportProgress();
                    toast(`Excel file exported successfully!`, "success");
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Project export error:", error);
                    toast("Failed to export project Excel file", "error");
                }
            };
            const exportToExcel = async () => {
                showExportProgress("Preparing data for export...");
                try {
                    const project = filterProject.value;
                    const customer = filterCustomer.value;
                    const projectManager = filterProjectManager.value;
                    const groupHead = filterGroupHead.value;
                    const projectReference = filterProjectReference.value;
                    const status = filterStatus.value;
                    const make = filterMake.value;
                    const filteredData = procurementData.filter((item) => {
                        var _a, _b;
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        const projectData = projects.find((p) => p.Title === projectTitle);
                        const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) || [];
                        if (project && projectTitle !== project)
                            return false;
                        if (customer && (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) !== customer)
                            return false;
                        if (projectManager) {
                            const managerName = typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                                ? (_b = projectData.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                                : projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager;
                            if (managerName !== projectManager)
                                return false;
                        }
                        if (groupHead && (projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead) !== groupHead)
                            return false;
                        if (projectReference && (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) !== projectReference)
                            return false;
                        if (status && !procurementItems.some((item) => item.Status === status))
                            return false;
                        if (make && !procurementItems.some((item) => item.Make === make))
                            return false;
                        return true;
                    });
                    if (filteredData.length === 0) {
                        hideExportProgress();
                        toast("No data to export", "warning");
                        return;
                    }
                    const headers = [
                        "Project Reference", "Project", "Customer", "Project Manager", "Group Head",
                        "Total Line Items",
                        "RFQ Pending", "Offer Pending", "TC pending", "Order Closure pending",
                        "PR pending", "PO pending", "Document Approval Pending", "MFC Pending",
                        "Inspection Pending", "Dispatch Pending", "Delivery Pending", "Delivered"
                    ];
                    const dataRows = [];
                    filteredData.forEach((item) => {
                        var _a, _b;
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        const projectData = projects.find((p) => p.Title === projectTitle);
                        const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) || [];
                        const statusCounts = calculateProjectStatusCounts(procurementItems);
                        const totalLineItems = procurementItems.length;
                        const managerName = typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                            ? ((_b = projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title) || ""
                            : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "";
                        dataRows.push([
                            (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "",
                            projectTitle,
                            (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "",
                            managerName,
                            (projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead) || "",
                            totalLineItems,
                            statusCounts["RFQ Pending"] || 0,
                            statusCounts["Offer Pending"] || 0,
                            statusCounts["TC pending"] || 0,
                            statusCounts["Order Closure pending"] || 0,
                            statusCounts["PR pending"] || 0,
                            statusCounts["PO pending"] || 0,
                            statusCounts["Document Approval Pending"] || 0,
                            statusCounts["MFC Pending"] || 0,
                            statusCounts["Inspection Pending"] || 0,
                            statusCounts["Dispatch Pending"] || 0,
                            statusCounts["Delivery Pending"] || 0,
                            statusCounts["Delivered"] || 0
                        ]);
                    });
                    const wsData = [headers, ...dataRows];
                    const ws = XLSX.utils.aoa_to_sheet(wsData);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Procurement Summary");
                    const excelBuffer = XLSX.write(wb, {
                        bookType: "xlsx",
                        type: "array",
                    });
                    const data = new Blob([excelBuffer], {
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });
                    const timestamp = new Date().toISOString().split("T")[0];
                    const filterInfo = [
                        project ? `Project-${project}` : "",
                        customer ? `Customer-${customer}` : "",
                        status ? `Status-${status}` : "",
                        make ? `Make-${make}` : "",
                    ].filter(Boolean).join("_");
                    const fileName = filterInfo
                        ? `Procurement_Summary_${filterInfo}_${timestamp}.xlsx`
                        : `Procurement_Summary_${timestamp}.xlsx`;
                    saveAs(data, fileName);
                    hideExportProgress();
                    toast(`Excel exported successfully! ${filteredData.length} projects`, "success");
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Export error:", error);
                    toast("Failed to export Excel report", "error");
                }
            };
            // Event Listeners
            showFormBtn.addEventListener("click", () => {
                resetForm();
                formSection.style.display = "block";
                tableSection.style.display = "none";
                renderProjectDetails(null);
                procurementContainerTable.style.display = "block";
                procurementContainer.innerHTML = "";
                procurementContainer.appendChild(createProcurementRow());
                formTitle.textContent = "Add Procurement";
                existingDataAlert.style.display = "none";
                resetDescriptionFilter();
                projectSelect.disabled = false;
                projectSelect.value = "";
            });
            backBtn.addEventListener("click", resetForm);
            projectSelect.addEventListener("change", (e) => {
                const selectedId = parseInt(e.target.value);
                if (selectedId) {
                    const existingRecord = procurementData.find((item) => item[`${fieldMappings.Project}Id`] === selectedId ||
                        (item[fieldMappings.Project] &&
                            item[fieldMappings.Project].Id === selectedId));
                    if (existingRecord) {
                        existingDataAlert.style.display = "block";
                        formTitle.textContent = "Edit Procurement Plan";
                        editingId = existingRecord.Id;
                        const procurementItems = parseProcurementForecast(existingRecord[fieldMappings.ProcurementForecast]) || [];
                        procurementContainer.innerHTML = "";
                        procurementItems.forEach((item) => {
                            procurementContainer.appendChild(createProcurementRow(item));
                        });
                    }
                    else {
                        existingDataAlert.style.display = "none";
                        formTitle.textContent = "Add Procurement";
                        editingId = null;
                        procurementContainer.innerHTML = "";
                        procurementContainer.appendChild(createProcurementRow());
                    }
                    onProjectChange(selectedId, !!existingRecord);
                }
                else {
                    procurementContainerTable.style.display = "none";
                    renderProjectDetails(null);
                    existingDataAlert.style.display = "none";
                    formTitle.textContent = "Add Procurement";
                    procurementContainer.innerHTML = "";
                    procurementContainer.appendChild(createProcurementRow());
                }
            });
            (_a = document
                .querySelector(".add-procurement-row")) === null || _a === void 0 ? void 0 : _a.addEventListener("click", () => {
                procurementContainer.appendChild(createProcurementRow());
                filterEditTableByDescription();
            });
            descriptionFilter === null || descriptionFilter === void 0 ? void 0 : descriptionFilter.addEventListener("input", filterEditTableByDescription);
            procurementContainer.addEventListener("input", (e) => {
                const target = e.target;
                if (target.name === "description") {
                    filterEditTableByDescription();
                }
            });
            procurementContainer.addEventListener("click", (e) => {
                const target = e.target;
                if (target.classList.contains("remove-row")) {
                    const row = target.closest("tr");
                    if (row && procurementContainer.children.length > 1) {
                        row.remove();
                    }
                    else {
                        toast("At least one row is required", "warning");
                    }
                }
            });
            filterProject.addEventListener("change", filterData);
            filterCustomer.addEventListener("change", filterData);
            filterProjectManager.addEventListener("change", filterData);
            filterGroupHead.addEventListener("change", filterData);
            filterProjectReference.addEventListener("change", filterData);
            filterStatus.addEventListener("change", filterData);
            filterMake.addEventListener("change", filterData);
            clearFiltersBtn.addEventListener("click", clearFilters);
            exportExcelBtn.addEventListener("click", exportToExcel);
            await initApp();
        }, 100);
    },
};
//# sourceMappingURL=procurement.js.map
import "../styles/common.css";
import "../styles/dcs.css";
import * as XLSX from 'xlsx-js-style';
import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css";
// ============ SHARED AUTH + ROLE BASED ACCESS (same as SEIBIES / inventory) ============
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, isPMManager, isEngineeringUser, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canReadAllProjects, canCreateForProject, 
// canViewForm,  
getAssignedProjects, updateUIBasedOnPermissions, } from "./auth";
// =================== DCS OPTION LISTS (Engg Division / Status) ===================
const ENGG_DIVISIONS = [
    "Automation",
    "SCADA",
    "Instrumentation",
    "PMS",
    "HSE",
    "Project Management"
];
// Status options for the DESIGN OUTPUT table
const OUTPUT_STATUS_OPTIONS = [
    "Not Submitted",
    "Submitted",
    "Approved",
    "Approved with comments",
    "Rejected",
    "Retained for Information",
    "Document Deleted/cancelled"
];
// Status options for the DESIGN INPUT table
const INPUT_STATUS_OPTIONS = [
    "Pending",
    "Received",
    "Partially Received",
    "Revised inputs required"
];
const getStatusOptions = (tableType) => tableType === 'input' ? INPUT_STATUS_OPTIONS : OUTPUT_STATUS_OPTIONS;
const getDefaultStatus = (tableType) => tableType === 'input' ? "Pending" : "Not Submitted";
const escapeOptionText = (value) => String(value !== null && value !== void 0 ? value : "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
// Build <option> list for the Status select. If the stored value is a legacy status
// (e.g. "Under Review") that is no longer in the list, it is kept as a selectable
// option so existing data is never lost/changed silently.
const buildStatusOptionsHTML = (tableType, selected) => {
    const options = getStatusOptions(tableType);
    const current = (selected && selected.trim()) ? selected.trim() : getDefaultStatus(tableType);
    let html = options
        .map(opt => `<option value="${escapeOptionText(opt)}" ${opt === current ? "selected" : ""}>${escapeOptionText(opt)}</option>`)
        .join("\n");
    if (!options.includes(current)) {
        html += `\n<option value="${escapeOptionText(current)}" selected>${escapeOptionText(current)} (legacy)</option>`;
    }
    return html;
};
// Build <option> list for the Engg Division select (blank option first)
const buildEnggDivisionOptionsHTML = (selected) => {
    const current = (selected || "").trim();
    let html = `<option value="" ${current === "" ? "selected" : ""}>-- Select --</option>\n`;
    html += ENGG_DIVISIONS
        .map(div => `<option value="${escapeOptionText(div)}" ${div === current ? "selected" : ""}>${escapeOptionText(div)}</option>`)
        .join("\n");
    if (current && !ENGG_DIVISIONS.includes(current)) {
        html += `\n<option value="${escapeOptionText(current)}" selected>${escapeOptionText(current)}</option>`;
    }
    return html;
};
const dcsContent = `
<div class="dcs-page fade-in">
  <section class="table-section dcs-card" id="dcsTableSection">
    <div class="dcs-page-header stick">
      <h1>Document Control Index</h1>
      <div>
        <button class="dcs-btn-primary dcs-show-form-btn">
          + Document Control Index
        </button>
      </div>
    </div>
    
    
    <!-- ✅ TOOLBAR ABOVE THE TABLE (mock-up image 8)
         LEFT  = Filters (funnel icon + active count badge) + the compact
                 "Apply" pill that appears when the popup has unapplied changes
         RIGHT = Download Excel
         The page TITLE and the header button are untouched - this toolbar
         never sits next to the title. -->
    <div class="dcs-table-toolbar">
      <div class="dcs-toolbar-left">
        <button type="button" class="dcs-btn-secondary dcs-filter-trigger" id="dcsFilterTrigger" aria-haspopup="dialog" aria-expanded="false" title="Filter Document Control System">
          <svg class="dcs-filter-funnel" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8v5l-4 2v-7L3 5z"/></svg>
          <span>Filters</span><span class="dcs-filter-count" id="dcsFilterCount">0</span>
          <svg class="dcs-filter-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>
        </button>
        <!-- appears only while the popup has unapplied changes -->
        <button type="button" class="dcs-btn-primary dcs-filter-applypill" id="dcsFilterApplyPill" title="Apply the selected filters">Apply</button>
      </div>
      <div class="dcs-toolbar-right">
        <button type="button" class="dcs-btn-secondary dcs-export-excel-btn">&#128202; Download Excel</button>
      </div>
    </div>
    <!-- applied filters show up here as removable chips -->
    <div class="dcs-filter-summary" id="dcsFilterChips" hidden></div>

    <div class="dcs-table-scroll-wrapper dcs-main-table-scroll">
      <table id="dcsTable" class="dcs-modern-table">
        <thead class="dcs-sticky-main-header">
          <tr>
            <th class="dcs-sticky-main-col dcs-main-col-1 dcs-sortable-th" data-sort="projectReference">
              <span class="dcs-th-inner"><span class="dcs-th-text">Project Reference</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-sticky-main-col dcs-main-col-2 dcs-sortable-th" data-sort="customer">
              <span class="dcs-th-inner"><span class="dcs-th-text">Customer</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-sticky-main-col dcs-main-col-3 dcs-sortable-th" data-sort="project">
              <span class="dcs-th-inner"><span class="dcs-th-text">Project</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-sticky-main-col dcs-main-col-4 dcs-sortable-th" data-sort="projectManager">
              <span class="dcs-th-inner"><span class="dcs-th-text">Project Manager</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-count-th dcs-sortable-th" data-sort="total">
              <span class="dcs-th-inner"><span class="dcs-th-text">Total Documents</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-count-th dcs-sortable-th" data-sort="notSubmitted">
              <span class="dcs-th-inner"><span class="dcs-th-text">Not Submitted</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-count-th dcs-sortable-th" data-sort="notApproved">
              <span class="dcs-th-inner"><span class="dcs-th-text">Not Approved</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-count-th dcs-sortable-th dcs-th-wrap" data-sort="revised">
              <span class="dcs-th-inner"><span class="dcs-th-text">Revised Document not Submitted / Approved With comments</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-count-th dcs-sortable-th" data-sort="approved">
              <span class="dcs-th-inner"><span class="dcs-th-text">Approved</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-count-th dcs-sortable-th" data-sort="rejected">
              <span class="dcs-th-inner"><span class="dcs-th-text">Rejected</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th class="dcs-count-th dcs-sortable-th" data-sort="retained">
              <span class="dcs-th-inner"><span class="dcs-th-text">Retained for Information</span><span class="dcs-sort-arrows"><span class="dcs-arrow-up">&#9650;</span><span class="dcs-arrow-down">&#9660;</span></span></span>
            </th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody id="dcsTableBody"><tr class="dcs-table-loading-row"><td colspan="12"><div class="dcs-page-loader"><div class="dcs-page-loader-track"><div class="dcs-page-loader-bar"></div></div></div></td></tr></tbody>
      </table>
    </div>

    <!-- ✅ Loading state (bar loader - image 4.png): grey track + indigo bar,
         no text. Shown while DCS records are being fetched. -->
    <div class="dcs-pagination-container">
      <div class="dcs-pagination-left">
        <div class="dcs-page-size-selector">
          <label for="pageSizeSelect">Items per page:</label>
          <select id="pageSizeSelect" class="dcs-page-size-select">
            <option value="5">5</option>
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>
        <div class="dcs-pagination-info">Showing <span id="startItem">0</span> - <span id="endItem">0</span> of <span id="totalItems">0</span> items</div>
      </div>
      <div class="dcs-pagination-right">
        <button class="dcs-pagination-btn dcs-pg-icon" id="firstPage" disabled title="First page">&#171;</button>
        <button class="dcs-pagination-btn dcs-pg-icon" id="prevPage" disabled title="Previous page">&#8249;</button>
        <div class="dcs-pagination-pages" id="pageNumbers"></div>
        <span class="dcs-pg-of">of <span id="totalPages">1</span></span>
        <button class="dcs-pagination-btn dcs-pg-icon" id="nextPage" disabled title="Next page">&#8250;</button>
        <button class="dcs-pagination-btn dcs-pg-icon" id="lastPage" disabled title="Last page">&#187;</button>
      </div>
    </div>
    <div class="dcs-filter-overlay" id="dcsFilterOverlay" role="dialog" aria-modal="true" aria-labelledby="dcsFilterTitle">
      <div class="dcs-filter-modal">
        <div class="dcs-filter-modal-head"><h2 id="dcsFilterTitle">Filters</h2><button type="button" id="dcsFilterClose" aria-label="Close">&times;</button></div>
        <div class="dcs-filter-grid">
          <div><label for="filterProject">Project</label><select id="filterProject" class="dcs-form-control"><option value="">All Projects</option></select></div>
          <div><label for="filterCustomer">Customer</label><select id="filterCustomer" class="dcs-form-control"><option value="">All Customers</option></select></div>
          <div><label for="filterProjectManager">Project Manager</label><select id="filterProjectManager" class="dcs-form-control"><option value="">All Managers</option></select></div>
          <div><label for="filterProjectReference">Project Reference</label><select id="filterProjectReference" class="dcs-form-control"><option value="">All References</option></select></div>
        </div>
        <!-- Clear = the Reset action - it lives HERE, inside the popup.
             There is NO Reset button in the toolbar. -->
        <div class="dcs-filter-modal-foot"><button type="button" class="dcs-btn-secondary" id="dcsFilterCancel">Cancel</button><button type="button" class="dcs-btn-secondary dcs-filter-btn-clear" id="dcsFilterClear" title="Clear all filters" disabled>Clear</button><button type="button" class="dcs-btn-primary dcs-filter-btn-apply" id="dcsFilterApply" disabled>Apply</button></div>
      </div>
    </div>
  </section>

  <section class="form-section dcs-card" id="dcsFormSection" style="display:none;">
  <form id="dcsForm">
  <div class="dcs-page-header stick">
    <h1 id="dcsFormTitle">Document Control Index</h1>
       <div class="dcs-top-btn-group">
     
      
      <!-- ✅ Wrap the Save button so we can safely remove/re-add it -->
      <div class="dcs-top-save-wrapper">
        <button type="submit"
                form="dcsForm"
                class="dcs-btn-primary dcs-btn-save-top"
                title="Save Document Data">💾 Save Document Data</button>
      </div>
       <button type="button" class="dcs-btn-secondary dcs-top-back-btn" style="min-height: 36px; padding: 8px 16px;">
        ⬅ Back
      </button>
    </div>
  </div>
  
          
      <!-- MAIN WRAPPER - 2 column layout (matching SEIBIES edit section) -->
      <div class="dcs-project-main-wrapper">

        <!-- LEFT SIDE - Project Info -->
        <div class="dcs-project-left-side" id="dcsProjectDetails">
          <div class="dcs-detail-item dcs-project-select-container">
            <span class="dcs-detail-label">Project:</span>
            <span class="dcs-detail-value dcs-detail-value-select">
              <select id="dcsProjectSelect" class="dcs-form-control" required>
                <option value="">Select Project</option>
              </select>
            </span>
          </div>
          <div class="dcs-detail-item">
            <span class="dcs-detail-label">Project Reference:</span>
            <span class="dcs-detail-value" id="projectReference">-</span>
          </div>
          <div class="dcs-detail-item">
            <span class="dcs-detail-label">Customer:</span>
            <span class="dcs-detail-value" id="projectCustomer">-</span>
          </div>
          <div class="dcs-detail-item">
            <span class="dcs-detail-label">Profit Center:</span>
            <span class="dcs-detail-value" id="projectPC">-</span>
          </div>
          <div class="dcs-detail-item">
            <span class="dcs-detail-label">Group Head:</span>
            <span class="dcs-detail-value" id="projectGroupHead">-</span>
          </div>
          <div class="dcs-detail-item">
            <span class="dcs-detail-label">Project Manager:</span>
            <span class="dcs-detail-value" id="projectManager">-</span>
          </div>
                 
        </div>

        <!-- RIGHT SIDE - Excel import instructions -->
        <div class="dcs-project-right-side">
          <div id="dcsImportInstructions" class="dcs-import-instructions">
            <p><strong>Instructions for Excel Import:</strong></p>
            <ul>
              <li><strong>Step 1:</strong> Click "Download Current Data" to get the existing data</li>
              <li><strong>Step 2:</strong> Edit the downloaded Excel file:
                <ul style="padding-left: 15px; margin-top: 5px;">
                  <li>Add new rows for new documents</li>
                  <li>Modify existing data as needed</li>
                  <li>Delete rows to remove documents</li>
                  <li>For revisions: Set Rev No > 0 and place below main document</li>
                  <li>Engg Division: Automation, SCADA, Instrumentation, PMS, HSE or Project Management</li>
                  <li>Status (Output): Submitted, Approved, Approved with comments, Rejected, Not Submitted, Retained for Information, Document Deleted/cancelled</li>
                  <li>Status (Input): Received, Pending, Partially Received, Revised inputs required</li>
                </ul>
              </li>
              <li><strong>Step 3:</strong> Save the Excel file and click "Import Excel Data" to upload</li>
              <li><strong>Note:</strong> Document numbers will be auto-generated if left blank</li>
            </ul>
          </div>
        </div>
      </div>
      
      <div id="dcsExistingDataAlert" class="dcs-alert dcs-alert-info" style="display: none; margin: 15px 0;">
        <div style="display: flex; align-items: center; gap: 10px;">
          <span style="font-size: 16px;">ℹ️</span>
          <div>
            <strong>Existing Data Found!</strong> This project already has document data. You can view and edit it below.
          </div>
        </div>
      </div>
       
      <!-- Excel action buttons - below the info / instructions sections -->
      <div id="dcsImportButtonContainer" style="margin-bottom: 15px; display: none;">
        <div class="dcs-import-buttons">
          <input type="file" id="dcsImportExcelInput" accept=".xlsx,.xls" style="display: none;">
          <button type="button" class="dcs-btn-primary dcs-import-excel-btn">
            📥 Import Excel Data
          </button>
          <button type="button" class="dcs-btn-secondary dcs-export-project-excel-btn">
            📊 Export Template Excel
          </button>
          <button type="button" class="dcs-btn-secondary dcs-download-existing-btn">
            📋 Download Current Data
          </button>
          <button type="button" class="dcs-btn-secondary" id="dcsResetFormBtn">
            🔄 Reset Form
          </button>
        </div>
      </div>
      
      <div class="dcs-inputs" id="dcsContainerTable" style="overflow-x:auto; display:none;">
        
        <!-- Design Output Section -->
        <div class="dcs-section-header">
          <h3>DESIGN OUTPUT- FROM ABB</h3>
          <button type="button" class="dcs-btn-secondary dcs-add-output-document">
            + Add Output Document
          </button>
        </div>
        <table class="dcs-modern-table dcs-output-table dcs-revision-table">
          <thead>
            <tr>
              <th rowspan="2" style="width: 40px;">Sr No</th>
              <th rowspan="2" style="min-width: 150px;">Document No</th>
              <th rowspan="2" style="min-width: 180px;">Description</th>
              <th rowspan="2" style="width: 100px;">Category</th>
              <th rowspan="2" style="width: 130px;">Engg Division</th>
              <th rowspan="2" style="width: 80px;">Rev No</th>
              <th colspan="2" style="background-color: #e3f2fd;">Document Schedule</th>
              <th colspan="2" style="background-color: #fff3cd;">Document Submission</th>
              <th colspan="2" style="background-color: #d4edda;">Document Approval</th> 
              <th rowspan="2" style="width: 120px;">Status</th>
              <th rowspan="2" style="width: 100px;">Remarks</th>
              <th rowspan="2" style="width: 100px;">Actions</th>
            </tr>
            <tr>
              <th style="background-color: #e3f2fd; width: 150px;">Original Plan</th>
              <th style="background-color: #e3f2fd; width: 150px;">Revised Plan</th>
              <th style="background-color: #fff3cd; width: 150px;">Actual Submission</th>
              <th style="background-color: #fff3cd; width: 150px;">Transmittal No</th>
              <th style="background-color: #d4edda; width: 150px;">Date</th>
              <th style="background-color: #d4edda; width: 150px;">Transmittal No</th>
            </tr>
          </thead>
          <tbody id="dcsOutputTableBody">
            <!-- Default output document -->
            <tr class="dcs-document-group dcs-document-header" data-doc-id="default_output_1" data-rev-rows="1" data-is-default="true">
              <td class="dcs-sr-no">1</td>
              <td>
                <textarea class="dcs-form-control dcs-doc-no" rows="2" placeholder="Document No">DOC-OUT-001</textarea>
              </td>
              <td>
                <textarea class="dcs-form-control dcs-description-input" rows="3" placeholder="Enter description"></textarea>
              </td>
              <td>
                <select class="dcs-form-control dcs-category-select">
                  <option value="A" selected>A - Approval</option>
                  <option value="I">I - Information</option>
                </select>
              </td>
              <td>
                <select class="dcs-form-control dcs-engg-division-select">
                  ${buildEnggDivisionOptionsHTML("")}
                </select>
              </td>
              <td class="dcs-rev-cell">0</td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Original Plan"></td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Revised Plan"></td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Actual Submission"></td>
              <td>
                <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="Transmittal No"></textarea>
              </td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Approval Date"></td>
              <td>
                <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="Transmittal No"></textarea>
              </td>
              <td>
                <select class="dcs-form-control dcs-status-select">
                  ${buildStatusOptionsHTML('output')}
                </select>
              </td>
              <td>
                <textarea class="dcs-form-control dcs-remarks-input" rows="3" placeholder="Enter remarks"></textarea>
              </td>
              <td class="dcs-doc-actions">
                <button type="button" class="dcs-btn-success dcs-add-revision-btn" title="Add Revision1">
                  <span style="font-size: 16px;">➕</span>
                </button>
                <button type="button" class="dcs-btn-danger dcs-remove-doc" title="Remove Document1" style="display: none;">
                  <span style="font-size: 16px;">🗑️</span>
                </button>
              </td>
            </tr>
          </tbody>
        </table>
        
        <!-- Design Input Section -->
        <div class="dcs-section-header">
          <h3>DESIGN INPUT FROM CONSULTANT / CUSTOMER</h3>
          <button type="button" class="dcs-btn-secondary dcs-add-input-document">
            + Add Input Document
          </button>
        </div>
        <table class="dcs-modern-table dcs-input-table dcs-revision-table">
          <thead>
            <tr>
              <th rowspan="2" style="width: 40px;">Sr No</th>
              <th rowspan="2" style="min-width: 150px;">Document No</th>
              <th rowspan="2" style="min-width: 180px;">Description</th>
              <th rowspan="2" style="width: 100px;">Category</th>
              <th rowspan="2" style="width: 130px;">Engg Division</th>
              <th rowspan="2" style="width: 80px;">Rev No</th>
              <th colspan="2" style="background-color: #e8f5e9;">Document Schedule</th>
              <th colspan="2" style="background-color: #fff3cd;">Document Submission</th>
              <th colspan="2" style="background-color: #d4edda;">Document Approval</th> 
              <th rowspan="2" style="width: 120px;">Status</th>
              <th rowspan="2" style="width: 100px;">Remarks</th>
              <th rowspan="2" style="width: 100px;">Actions</th>
            </tr>
            <tr>
              <th style="background-color: #e8f5e9; width: 150px;">Original Plan</th>
              <th style="background-color: #e8f5e9; width: 150px;">Revised Plan</th>
              <th style="background-color: #fff3cd; width: 150px;">Actual Submission</th>
              <th style="background-color: #fff3cd; width: 150px;">Transmittal No</th>
              <th style="background-color: #d4edda; width: 150px;">Date</th>
              <th style="background-color: #d4edda; width: 150px;">Transmittal No</th>
            </tr>
          </thead>
          <tbody id="dcsInputTableBody">
            <!-- Default input document -->
            <tr class="dcs-document-group dcs-document-header" data-doc-id="default_input_1" data-rev-rows="1" data-is-default="true">
              <td class="dcs-sr-no">1</td>
              <td>
                <textarea class="dcs-form-control dcs-doc-no" rows="2" placeholder="Document No">DOC-001</textarea>
              </td>
              <td>
                <textarea class="dcs-form-control dcs-description-input" rows="3" placeholder="Enter description"></textarea>
              </td>
              <td>
                <select class="dcs-form-control dcs-category-select">
                  <option value="A" selected>A - Approval</option>
                  <option value="I">I - Information</option>
                </select>
              </td>
              <td>
                <select class="dcs-form-control dcs-engg-division-select">
                  ${buildEnggDivisionOptionsHTML("")}
                </select>
              </td>
              <td class="dcs-rev-cell">0</td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Original Plan"></td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Revised Plan"></td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Actual Submission"></td>
              <td>
                <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="Transmittal No"></textarea>
              </td>
              <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Approval Date"></td>
              <td>
                <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="Transmittal No"></textarea>
              </td>
              <td>
                <select class="dcs-form-control dcs-status-select">
                  ${buildStatusOptionsHTML('input')}
                </select>
              </td>
              <td>
                <textarea class="dcs-form-control dcs-remarks-input" rows="3" placeholder="Enter remarks"></textarea>
              </td>
              <td class="dcs-doc-actions">
                <button type="button" class="dcs-btn-success dcs-add-revision-btn" title="Add Revision2">
                  <span style="font-size: 16px;">➕</span>
                </button>
                <button type="button" class="dcs-btn-danger dcs-remove-doc" title="Remove Document2" style="display: none;">
                  <span style="font-size: 16px;">🗑️</span>
                </button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="dcs-form-actions">
        <button type="submit" class="dcs-btn-primary">
          💾 Save Document Data
        </button>
        <button type="button" class="dcs-btn-secondary dcs-back-to-table">
          ⬅ Back
        </button>
      </div>
    </form>
  </section>
</div>
`;
const DCSModule = {
    content: dcsContent,
    script: () => {
        setTimeout(async () => {
            if (window.__dtActivePageKey !== 'dci')
                return;
            // Configuration
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/Kriya";
            const listName = "DCSTracker";
            const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
            // State variables
            let fieldMappings = {
                DCSData: "DCSData",
                Project: "Project",
            };
            let listEntityTypeName = "";
            let projects = [];
            let dcsData = [];
            // Toast function with proper typing
            // ============================================================
            // ✅ SHARED TOAST — matches procurement.ts / other modules
            // ============================================================
            const toast = (message, type = "success") => {
                Swal.fire({
                    toast: true,
                    position: "top-end",
                    icon: type,
                    title: message,
                    showConfirmButton: false,
                    timer: 3000,
                    timerProgressBar: true,
                });
            };
            // ============================================
            // LOADING OVERLAY (same pattern as master.ts)
            // ============================================
            const createLoadingOverlay = (message = "Processing...") => {
                removeLoadingOverlay(); // never stack two loaders
                const overlay = document.createElement("div");
                overlay.className = "loading-overlay";
                overlay.setAttribute("aria-label", message);
                overlay.innerHTML = `
          <div class="loading-container">
            <div class="dcs-page-loader">
              <div class="dcs-page-loader-track"><div class="dcs-page-loader-bar"></div></div>
            </div>
          </div>`;
                if (!document.querySelector("#loading-styles")) {
                    const style = document.createElement("style");
                    style.id = "loading-styles";
                    style.textContent = `
            .loading-overlay{position:absolute;top:0;left:0;width:100%;height:100%;background:rgba(255,255,255,0.94);display:flex;justify-content:center;align-items:center;z-index:99999;backdrop-filter:blur(3px);border-radius:10px;overflow:hidden;box-sizing:border-box}
            .loading-container{padding:26px 30px;background:#fff;border:1px solid #e5e7eb;border-radius:10px;box-shadow:0 10px 32px rgba(16,24,40,.16);width:min(520px,85vw);box-sizing:border-box}`;
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
                const sections = Array.from(document.querySelectorAll('.table-section, .form-section'));
                const visible = sections.find(s => s.offsetParent !== null);
                return visible || document.querySelector('.table-section') || document.body;
            };
            // DOM Elements
            const projectSelect = document.getElementById("dcsProjectSelect");
            const showFormBtn = document.querySelector(".dcs-show-form-btn");
            const exportExcelBtn = document.querySelector(".dcs-export-excel-btn");
            const formSection = document.getElementById("dcsFormSection");
            const tableSection = document.getElementById("dcsTableSection");
            const backBtn = document.querySelector(".dcs-back-to-table");
            const outputTableBody = document.getElementById("dcsOutputTableBody");
            const inputTableBody = document.getElementById("dcsInputTableBody");
            const containerTable = document.getElementById("dcsContainerTable");
            const dcsForm = document.getElementById("dcsForm");
            const formTitle = document.getElementById("dcsFormTitle");
            const existingDataAlert = document.getElementById("dcsExistingDataAlert");
            const addInputDocBtn = document.querySelector(".dcs-add-input-document");
            const addOutputDocBtn = document.querySelector(".dcs-add-output-document");
            const importButtonContainer = document.getElementById("dcsImportButtonContainer");
            const resetFormBtn = document.getElementById("dcsResetFormBtn");
            // Filter Elements
            const filterProject = document.getElementById("filterProject");
            const filterCustomer = document.getElementById("filterCustomer");
            const filterProjectManager = document.getElementById("filterProjectManager");
            const filterProjectReference = document.getElementById("filterProjectReference");
            const filterTriggerBtn = document.getElementById("dcsFilterTrigger");
            const filterOverlay = document.getElementById("dcsFilterOverlay");
            const filterCloseBtn = document.getElementById("dcsFilterClose");
            const filterCancelBtn = document.getElementById("dcsFilterCancel");
            const filterClearBtn = document.getElementById("dcsFilterClear");
            const filterApplyBtn = document.getElementById("dcsFilterApply");
            const filterCount = document.getElementById("dcsFilterCount");
            const projectPC = document.getElementById("projectPC");
            const projectCustomer = document.getElementById("projectCustomer");
            const projectReference = document.getElementById("projectReference");
            const projectGroupHead = document.getElementById("projectGroupHead");
            const projectManager = document.getElementById("projectManager");
            // Pagination variables
            let currentPage = 1;
            let pageSize = 25; // ✅ default 25 rows per page
            let totalPages = 1;
            let filteredDCSData = [];
            // ---- Main table sorting state ----
            let sortKey = null;
            let sortDir = 'asc';
            const EMPTY_FILTERS = { project: '', customer: '', manager: '', reference: '' };
            const FILTER_LABELS = {
                project: 'Project',
                customer: 'Customer',
                manager: 'Project Manager',
                reference: 'Project Reference'
            };
            // ✅ Filters start CLEAN on every load - nothing is pre-selected.
            let appliedFilters = Object.assign({}, EMPTY_FILTERS);
            // toolbar badge / chips / floating Apply pill
            const filterChipsHost = document.getElementById('dcsFilterChips');
            const filterApplyPill = document.getElementById('dcsFilterApplyPill');
            const getDraftFilters = () => ({
                project: (filterProject === null || filterProject === void 0 ? void 0 : filterProject.value) || '',
                customer: (filterCustomer === null || filterCustomer === void 0 ? void 0 : filterCustomer.value) || '',
                manager: (filterProjectManager === null || filterProjectManager === void 0 ? void 0 : filterProjectManager.value) || '',
                reference: (filterProjectReference === null || filterProjectReference === void 0 ? void 0 : filterProjectReference.value) || ''
            });
            const sameFilters = (a, b) => Object.keys(EMPTY_FILTERS).every(k => (a[k] || '') === (b[k] || ''));
            const activeFilterCount = () => Object.values(appliedFilters).filter(Boolean).length;
            const hasDraftFilters = () => Object.values(getDraftFilters()).some(Boolean);
            const filtersAreDirty = () => !sameFilters(getDraftFilters(), appliedFilters);
            const restoreAppliedFilters = () => {
                if (filterProject)
                    filterProject.value = appliedFilters.project;
                if (filterCustomer)
                    filterCustomer.value = appliedFilters.customer;
                if (filterProjectManager)
                    filterProjectManager.value = appliedFilters.manager;
                if (filterProjectReference)
                    filterProjectReference.value = appliedFilters.reference;
            };
            // chips for the applied filters (each removable, plus "Clear all")
            const renderFilterChips = () => {
                if (!filterChipsHost)
                    return;
                const entries = Object.keys(appliedFilters)
                    .filter(k => !!appliedFilters[k])
                    .map(k => ({ key: k, value: appliedFilters[k] }));
                filterChipsHost.innerHTML = '';
                if (entries.length === 0) {
                    filterChipsHost.hidden = true;
                    return;
                }
                filterChipsHost.hidden = false;
                const label = document.createElement('span');
                label.className = 'dcs-filter-summary__label';
                label.textContent = 'Active filters:';
                filterChipsHost.appendChild(label);
                entries.forEach(({ key, value }) => {
                    var _a;
                    const chip = document.createElement('span');
                    chip.className = 'dcs-filter-chip';
                    chip.innerHTML = `<span>${FILTER_LABELS[key]}: <strong>${escapeHtml(value)}</strong></span>` +
                        `<button type="button" title="Remove this filter" aria-label="Remove ${FILTER_LABELS[key]} filter">&times;</button>`;
                    (_a = chip.querySelector('button')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
                        appliedFilters = Object.assign(Object.assign({}, appliedFilters), { [key]: '' });
                        restoreAppliedFilters();
                        syncFilterUI();
                        filterData();
                    });
                    filterChipsHost.appendChild(chip);
                });
                const clearAll = document.createElement('button');
                clearAll.type = 'button';
                clearAll.className = 'dcs-filter-chip__all';
                clearAll.textContent = 'Clear all';
                clearAll.addEventListener('click', () => resetFilters(false));
                filterChipsHost.appendChild(clearAll);
            };
            // single source of truth for badge / trigger state / popup + pill buttons
            const syncFilterUI = () => {
                const active = activeFilterCount();
                if (filterCount)
                    filterCount.textContent = String(active);
                filterTriggerBtn === null || filterTriggerBtn === void 0 ? void 0 : filterTriggerBtn.classList.toggle('is-active', active > 0);
                if (filterClearBtn) {
                    const canClear = active > 0 || hasDraftFilters();
                    filterClearBtn.disabled = !canClear;
                    filterClearBtn.classList.toggle('is-enabled', canClear);
                }
                if (filterApplyBtn) {
                    const canApply = filtersAreDirty();
                    filterApplyBtn.disabled = !canApply;
                    filterApplyBtn.classList.toggle('is-enabled', canApply);
                }
                filterApplyPill === null || filterApplyPill === void 0 ? void 0 : filterApplyPill.classList.toggle('is-visible', filtersAreDirty());
                renderFilterChips();
            };
            // Clear = the Reset action (popup only). Draft + applied both go empty
            // and the table re-renders immediately.
            const resetFilters = (closePopup) => {
                appliedFilters = Object.assign({}, EMPTY_FILTERS);
                restoreAppliedFilters();
                syncFilterUI();
                filterData();
                if (closePopup)
                    closeFilterModal(false);
                console.log('🔄 Filters cleared');
            };
            const openFilterModal = () => {
                restoreAppliedFilters(); // the draft always starts from what is applied
                filterOverlay === null || filterOverlay === void 0 ? void 0 : filterOverlay.classList.add('is-open');
                filterTriggerBtn === null || filterTriggerBtn === void 0 ? void 0 : filterTriggerBtn.setAttribute('aria-expanded', 'true');
                document.body.style.overflow = 'hidden';
                syncFilterUI();
            };
            const closeFilterModal = (restore = true) => {
                if (restore)
                    restoreAppliedFilters(); // Cancel / ✕ / backdrop discard the draft
                filterOverlay === null || filterOverlay === void 0 ? void 0 : filterOverlay.classList.remove('is-open');
                filterTriggerBtn === null || filterTriggerBtn === void 0 ? void 0 : filterTriggerBtn.setAttribute('aria-expanded', 'false');
                document.body.style.overflow = '';
                syncFilterUI();
            };
            const commitFilters = () => {
                appliedFilters = getDraftFilters();
                syncFilterUI();
                filterData();
            };
            const setupFilterEventListeners = () => {
                filterTriggerBtn === null || filterTriggerBtn === void 0 ? void 0 : filterTriggerBtn.addEventListener('click', openFilterModal);
                filterCloseBtn === null || filterCloseBtn === void 0 ? void 0 : filterCloseBtn.addEventListener('click', () => closeFilterModal());
                filterCancelBtn === null || filterCancelBtn === void 0 ? void 0 : filterCancelBtn.addEventListener('click', () => closeFilterModal());
                filterOverlay === null || filterOverlay === void 0 ? void 0 : filterOverlay.addEventListener('click', (event) => { if (event.target === filterOverlay)
                    closeFilterModal(); });
                // choosing a value in the popup only enables the buttons + shows the pill -
                // the table changes only when Apply (or the pill) is pressed
                [filterProject, filterCustomer, filterProjectManager, filterProjectReference].forEach(sel => sel === null || sel === void 0 ? void 0 : sel.addEventListener('change', syncFilterUI));
                filterClearBtn === null || filterClearBtn === void 0 ? void 0 : filterClearBtn.addEventListener('click', () => resetFilters(false));
                filterApplyBtn === null || filterApplyBtn === void 0 ? void 0 : filterApplyBtn.addEventListener('click', () => { commitFilters(); closeFilterModal(false); });
                filterApplyPill === null || filterApplyPill === void 0 ? void 0 : filterApplyPill.addEventListener('click', commitFilters);
                syncFilterUI();
            };
            const getRequestDigest = async () => {
                try {
                    const res = await fetch(`${siteUrl}/_api/contextinfo`, {
                        method: "POST",
                        headers: {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                        },
                        credentials: 'same-origin'
                    });
                    if (!res.ok) {
                        throw new Error(`Failed to get request digest: ${res.status} ${res.statusText}`);
                    }
                    const data = await res.json();
                    return data.d.GetContextWebInformation.FormDigestValue;
                }
                catch (error) {
                    console.error("Error getting request digest:", error);
                    toast("Failed to get security token. Please refresh the page.", "error");
                    throw error;
                }
            };
            const initializeEventListeners = () => {
                // 1. Show form button
                if (showFormBtn) {
                    showFormBtn.removeEventListener('click', handleShowForm);
                    showFormBtn.addEventListener('click', handleShowForm);
                }
                if (backBtn) {
                    backBtn.removeEventListener('click', handleBackToTable);
                    backBtn.addEventListener('click', () => { void handleBackToTable(); });
                }
                const topBackBtn = document.querySelector('.dcs-top-back-btn');
                if (topBackBtn) {
                    topBackBtn.removeEventListener('click', handleBackToTable);
                    topBackBtn.addEventListener('click', () => { void handleBackToTable(); });
                }
                // 3. Form submission
                if (dcsForm) {
                    dcsForm.removeEventListener('submit', handleFormSubmit);
                    dcsForm.addEventListener('submit', handleFormSubmit);
                }
                // 4. Project select change
                if (projectSelect) {
                    projectSelect.removeEventListener('change', handleProjectSelect);
                    projectSelect.addEventListener('change', handleProjectSelect);
                }
                setupImportExportFunctionality();
                // Show import instructions when project is selected
                if (projectSelect) {
                    projectSelect.addEventListener('change', () => {
                        const selectedId = parseInt(projectSelect.value);
                        if (selectedId) {
                            const importInstructions = document.getElementById('dcsImportInstructions');
                            if (importInstructions) {
                                importInstructions.style.display = 'block';
                            }
                        }
                    });
                }
                setupFilterEventListeners();
                // 6. Export button
                if (exportExcelBtn) {
                    exportExcelBtn.removeEventListener('click', exportToExcel);
                    exportExcelBtn.addEventListener('click', exportToExcel);
                }
                // 7. Reset form button
                if (resetFormBtn) {
                    resetFormBtn.removeEventListener('click', resetForm);
                    resetFormBtn.addEventListener('click', resetForm);
                }
            };
            // ============ ROLE BASED ACCESS UI ============
            // Same rules as the rest of the suite:
            //  - Add Project Documents  : Admin / GH Head / PM Manager (canCreateRecord)
            //  - Import Excel           : Admin / GH Head / PM Manager (canCreateRecord)
            //  - Export Excel (main)    : anyone who can read (canRead)
            //  - Save                   : Admin / GH Head / PM Manager / Engg (own) who can reach the form
            const applyRoleBasedUI = () => {
                const assigned = getAssignedProjects(projects);
                const canCreateAnyProject = canCreateRecord() ||
                    (isEngineeringUser() && assigned.length > 0);
                // Add + Import buttons
                if (showFormBtn) {
                    showFormBtn.style.display = canCreateAnyProject ? "inline-flex" : "none";
                }
                const importExcelBtn = document.querySelector('.dcs-import-excel-btn');
                if (importExcelBtn) {
                    importExcelBtn.style.display = canCreateAnyProject ? "inline-flex" : "none";
                }
                // Export / Save via the shared helper — unchanged
                updateUIBasedOnPermissions({
                    addBtn: '.dcs-show-form-btn',
                    exportBtn: '.dcs-export-excel-btn',
                    importBtn: '.dcs-import-excel-btn',
                    saveBtn: '.dcs-form-actions button[type="submit"]'
                });
            };
            const handleShowForm = () => {
                const assigned = getAssignedProjects(projects);
                const canOpenForm = canCreateRecord() ||
                    (isEngineeringUser() && assigned.length > 0);
                if (!canOpenForm) {
                    toast("You do not have permission to add project documents", "error");
                    return;
                }
                // Simple show/hide
                formSection.style.display = "block";
                tableSection.style.display = "none";
                // Reset everything
                projectSelect.value = "";
                resetProjectDetailsDisplay();
                // Show default documents WITHOUT delete buttons
                inputTableBody.innerHTML = createDocumentRowHTML('input', 1, `default_${Date.now()}_in`);
                outputTableBody.innerHTML = createDocumentRowHTML('output', 1, `default_${Date.now()}_out`);
                // Show import buttons for new projects too
                importButtonContainer.style.display = "flex";
                // Show import instructions
                const importInstructions = document.getElementById('dcsImportInstructions');
                if (importInstructions) {
                    importInstructions.style.display = 'block';
                }
                // Make sure buttons are visible
                if (addInputDocBtn)
                    addInputDocBtn.style.display = "inline-block";
                if (addOutputDocBtn)
                    addOutputDocBtn.style.display = "inline-block";
                // Show the container table
                containerTable.style.display = "block";
                // Attach listeners
                setTimeout(() => {
                    attachDocumentEventListeners();
                }, 100);
            };
            const getSaveButtons = () => Array.from(document.querySelectorAll('button.dcs-btn-primary[type="submit"][form="dcsForm"], .dcs-form-actions button.dcs-btn-primary[type="submit"]'));
            // ============================================================
            // ✅ BACK — SweetAlert confirmation, same look as procurement.ts
            // ============================================================
            let isLeavingForm = false;
            const handleBackToTable = async () => {
                if (isLeavingForm) {
                    console.log("⛔ back blocked: already leaving");
                    return;
                }
                isLeavingForm = true;
                try {
                    const isEditing = formSection.classList.contains("dcs-edit-mode");
                    // const isViewMode = formSection.classList.contains("dcs-view-mode");
                    // Only prompt when there's something to lose
                    if (isEditing) {
                        const confirmResult = await Swal.fire({
                            title: "Leave this page?",
                            html: "Any <strong>unsaved changes</strong> will be lost.<br>" +
                                "Do you want to go back to the document list?",
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
                            didOpen: () => {
                                const c = document.querySelector(".swal2-container");
                                if (c)
                                    c.style.zIndex = "2147483647";
                            },
                        });
                        if (!confirmResult.isConfirmed) {
                            console.log("🔵 back cancelled by user");
                            return;
                        }
                    }
                    // --- the rest is unchanged from your original logic ---
                    formSection.classList.remove("dcs-view-mode");
                    projectSelect.disabled = false;
                    const saveWrapper = document.querySelector(".dcs-top-save-wrapper");
                    if (saveWrapper) {
                        saveWrapper.style.display = "block";
                    }
                    getSaveButtons().forEach(btn => {
                        btn.style.display = "inline-flex";
                    });
                    formSection.style.display = "none";
                    tableSection.style.display = "block";
                    resetForm();
                }
                catch (err) {
                    console.error("Error while leaving the form:", err);
                    const message = err instanceof Error ? err.message : String(err);
                    toast(`Unexpected error: ${message}`, "error");
                }
                finally {
                    isLeavingForm = false;
                }
            };
            const handleFormSubmit = async (e) => {
                e.preventDefault();
                await saveAllDataToSingleJSON();
            };
            const handleProjectSelect = (e) => {
                const selectedId = parseInt(e.target.value);
                if (!selectedId)
                    return;
                const project = projects.find(p => p.Id === selectedId);
                // ✅ This call consumes the import — and enforces per-project ownership
                const canCreateHere = canCreateRecord() || canCreateForProject(project);
                if (!canCreateHere) {
                    toast("You can only add documents for projects assigned to you", "error");
                    e.target.value = "";
                    resetProjectDetailsDisplay();
                    containerTable.style.display = "none";
                    return;
                }
                onProjectChange(selectedId);
            };
            // ============================================ //
            // OUTER-STICKY HEADER (main table)             //
            //   The page scrolls vertically. The main      //
            //   page-header is sticky at top:0. The table  //
            //   header is sticky at top = main header      //
            //   height, so it 'touches' the main header    //
            //   and pins below it. No per-cell JS offset.  //
            // ============================================ //
            const updateMainTableStickyOffset = () => {
                const section = document.querySelector(".table-section");
                const header = section === null || section === void 0 ? void 0 : section.querySelector(".dcs-page-header");
                if (!section || !header)
                    return;
                // Fall back to a sensible minimum in case offsetHeight is 0 (e.g.
                // before fonts/images load), so the table header never sticks at
                // top:0 on top of the page header.
                const headerH = Math.max(header.offsetHeight || 0, 58);
                section.style.setProperty("--dcs-sticky-header-top", `${headerH}px`);
            };
            const initApp = async () => {
                try {
                    console.log('🚀 Initializing app...');
                    const initialLoading = createLoadingOverlay("Loading DCS data...");
                    loaderHost().appendChild(initialLoading);
                    // ✅ Initialize shared authentication + role context (same as SEIBIES / inventory)
                    await authManager.initialize({ siteUrl });
                    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                    console.log(`👤 User: ${getCurrentUserFullName()}`);
                    console.log(`📋 User Groups:`, getUserGroups());
                    console.log(`✅ Full Control: ${hasFullControl()} | Can Contribute: ${canContribute()} | Can Read: ${canRead()}`);
                    console.log(`✅ Can Read All Projects: ${canReadAllProjects()} | Can Freeze: ${canFreeze()}`);
                    console.log(`✅ Is GH Head: ${isGHHead()} | Is PM Manager: ${isPMManager()} | Is Engg User: ${isEngineeringUser()}`);
                    // Initialize SharePoint
                    await initializeSharePoint();
                    // Load projects
                    await loadProjects();
                    // Load DCS data
                    await fetchData();
                    console.log(`📊 Data loaded: ${projects.length} projects, ${dcsData.length} DCS records`);
                    // Setup UI
                    initializeEventListeners();
                    setupPaginationEventListeners();
                    attachDocumentEventListeners();
                    // Table head sticks right below the (sticky) page header
                    updateMainTableStickyOffset();
                    window.addEventListener('resize', updateMainTableStickyOffset);
                    // ✅ Role based UI (hide Add / Import for read-only roles etc.)
                    applyRoleBasedUI();
                    console.log('✅ App initialization complete');
                }
                catch (error) {
                    console.error("❌ Initialization error:", error);
                    toast("Failed to initialize application", "error");
                }
                finally {
                    removeLoadingOverlay();
                }
            };
            const setupPaginationEventListeners = () => {
                // Navigation buttons
                const firstPageBtn = document.getElementById("firstPage");
                const prevPageBtn = document.getElementById("prevPage");
                const nextPageBtn = document.getElementById("nextPage");
                const lastPageBtn = document.getElementById("lastPage");
                const pageSizeSelect = document.getElementById("pageSizeSelect");
                if (firstPageBtn)
                    firstPageBtn.addEventListener("click", () => goToPage(1));
                if (prevPageBtn)
                    prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
                if (nextPageBtn)
                    nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
                if (lastPageBtn)
                    lastPageBtn.addEventListener("click", () => goToPage(totalPages));
                // Page size selector
                if (pageSizeSelect) {
                    pageSizeSelect.addEventListener("change", (e) => {
                        const newSize = parseInt(e.target.value);
                        changePageSize(newSize);
                    });
                }
                // ---- Main table column sorting (delegated on the static thead) ----
                const mainThead = document.querySelector('#dcsTable thead');
                if (mainThead)
                    mainThead.addEventListener('click', handleSortClick);
                updateSortIndicators();
            };
            const loadProjects = async () => {
                var _a;
                try {
                    // ✅ $select must include every field that ROLE_OWN_FIELDS in auth.ts
                    // matches against, otherwise getAssignedProjects() returns [] for the
                    // KRIYA_Engg_* roles (and their DCS rows are filtered out entirely).
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('Project List')/items` +
                        `?$select=Id,Title,CustomerName,PC,ProjectReference,` +
                        `GroupHead1/Id,GroupHead1/Title,` +
                        `ProjectManager/Id,ProjectManager/Title,` +
                        `EnggLead/Id,EnggLead/Title,` +
                        `EnggSubLead1/Id,EnggSubLead1/Title,` +
                        `EnggSubLead2/Id,EnggSubLead2/Title,` +
                        `EngineeringManagers/Id,EngineeringManagers/Title` +
                        `&$expand=GroupHead1,ProjectManager,EnggLead,EnggSubLead1,EnggSubLead2,EngineeringManagers` +
                        `&$top=1000`, {
                        headers: { Accept: "application/json;odata=verbose" },
                        credentials: 'same-origin'
                    });
                    if (!res.ok) {
                        const txt = await res.text();
                        console.error("Project List error body:", txt);
                        throw new Error(`Failed to load projects: ${res.status}`);
                    }
                    const data = await res.json();
                    const allProjects = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    const validProjects = allProjects.filter(p => p.Title && p.Title.trim() !== "");
                    const mappedProjects = validProjects.map((p) => {
                        var _a, _b, _c, _d, _e, _f;
                        return ({
                            Id: p.Id,
                            Title: p.Title || "",
                            CustomerName: p.CustomerName || "",
                            GroupHead1: p.GroupHead1 || "",
                            GroupHead1Id: ((_a = p.GroupHead1) === null || _a === void 0 ? void 0 : _a.Id) || p.GroupHead1Id,
                            PC: p.PC || "",
                            ProjectReference: p.ProjectReference || "",
                            ProjectManager: p.ProjectManager || "",
                            ProjectManagerId: ((_b = p.ProjectManager) === null || _b === void 0 ? void 0 : _b.Id) || p.ProjectManagerId,
                            // ✅ Keep the engineering fields so getAssignedProjects() can match
                            EnggLead: p.EnggLead || "",
                            EnggLeadId: ((_c = p.EnggLead) === null || _c === void 0 ? void 0 : _c.Id) || p.EnggLeadId,
                            EnggSubLead1: p.EnggSubLead1 || "",
                            EnggSubLead1Id: ((_d = p.EnggSubLead1) === null || _d === void 0 ? void 0 : _d.Id) || p.EnggSubLead1Id,
                            EnggSubLead2: p.EnggSubLead2 || "",
                            EnggSubLead2Id: ((_e = p.EnggSubLead2) === null || _e === void 0 ? void 0 : _e.Id) || p.EnggSubLead2Id,
                            EngineeringManagers: p.EngineeringManagers || "",
                            EngineeringManagersId: ((_f = p.EngineeringManagers) === null || _f === void 0 ? void 0 : _f.Id) || p.EngineeringManagersId
                        });
                    });
                    const uniqueProjects = Array.from(new Map(mappedProjects.map((p) => [p.Title, p])).values());
                    projects = uniqueProjects;
                    const visibleProjects = getAssignedProjects(uniqueProjects);
                    console.log(`🔐 Visible projects for this role: ${visibleProjects.length} / ${uniqueProjects.length}`);
                    console.log('🔐 Visible project titles:', visibleProjects.map(p => p.Title));
                    projectSelect.innerHTML =
                        '<option value="">Select Project</option>' +
                            visibleProjects.map((p) => `<option value="${p.Id}">${p.Title}</option>`).join("");
                    filterProject.innerHTML =
                        '<option value="">All Projects</option>' +
                            visibleProjects.map((p) => `<option value="${p.Title}">${p.Title}</option>`).join("");
                }
                catch (error) {
                    console.error("Error loading projects:", error);
                    toast("Failed to load projects", "error");
                }
            };
            // ✅ Edit title like Procurement: "Edit Document Control Data - <Project> (<Ref>)"
            // ✅ Title shows only the project name (Ref); "Edit Document Control Data - " prefix removed
            const setEditTitle = (project) => {
                if (!formTitle)
                    return;
                // const name = project?.Title || "";
                const title = "Document Control Index";
                // const ref = project?.ProjectReference || "";
                formTitle.textContent = title
                    ? `${title}`
                    : "Edit Document Control Index";
                formSection.classList.add("dcs-edit-mode");
            };
            // Fill the project info rows (left panel of the form) for the given project
            const setProjectDetailsDisplay = (project) => {
                projectPC.textContent = (project === null || project === void 0 ? void 0 : project.PC) || "-";
                projectCustomer.textContent = (project === null || project === void 0 ? void 0 : project.CustomerName) || "-";
                projectReference.textContent = (project === null || project === void 0 ? void 0 : project.ProjectReference) || "-";
                projectGroupHead.textContent = project ? getGroupHeadDisplayName(project.GroupHead1) : "-";
                projectManager.textContent = project ? getProjectManagerDisplayName(project.ProjectManager) : "-";
            };
            // Reset the project info rows to "-" (the Project select itself is kept intact)
            const resetProjectDetailsDisplay = () => setProjectDetailsDisplay(undefined);
            const getGroupHeadDisplayName = (groupHead) => {
                if (!groupHead)
                    return "-";
                if (typeof groupHead === "object") {
                    return groupHead.Title || "-";
                }
                return groupHead;
            };
            const getProjectManagerDisplayName = (manager) => {
                if (!manager)
                    return "-";
                if (typeof manager === "object") {
                    return manager.Title || "-";
                }
                return manager;
            };
            // Parse JSON data from DCSData field
            const parseDCSData = (dcsDataJson) => {
                var _a;
                if (!dcsDataJson || dcsDataJson.trim() === "")
                    return null;
                try {
                    const data = JSON.parse(dcsDataJson);
                    // Check if it's the new format
                    if (data.documentStructure && (data.documentStructure.inputs || data.documentStructure.outputs)) {
                        return data;
                    }
                    // Check if it's the old format
                    if (data.inputs || data.outputs) {
                        return data;
                    }
                    // Check if it's the very old format
                    if (Array.isArray(data)) {
                        // This is an array, convert to new format
                        const inputs = [];
                        const outputs = [];
                        data.forEach((item) => {
                            if (item.isInput) {
                                inputs.push(item);
                            }
                            else {
                                outputs.push(item);
                            }
                        });
                        return {
                            projectId: ((_a = data[0]) === null || _a === void 0 ? void 0 : _a.projectId) || 0,
                            documentStructure: {
                                inputs,
                                outputs
                            }
                        };
                    }
                    return data;
                }
                catch (error) {
                    console.error("Error parsing DCSData JSON:", error);
                    return null;
                }
            };
            // ✅ Bar loader (image 4.png) - replaces the old spinner/text loader.
            //    While a fetch is in flight the table body shows this static row;
            //    renderTable() overwrites it with the real rows. No text on screen -
            //    all messages go to the console.
            const DCS_LOADING_ROW_HTML = `
        <tr class="dcs-table-loading-row">
          <td colspan="12">
            <div class="dcs-page-loader">
              <div class="dcs-page-loader-track"><div class="dcs-page-loader-bar"></div></div>
            </div>
          </td>
        </tr>`;
            const showDCSLoading = () => {
                const body = document.getElementById("dcsTableBody");
                if (body)
                    body.innerHTML = DCS_LOADING_ROW_HTML;
            };
            const hideDCSLoading = () => {
                /* renderTable() replaces the loading row - nothing to hide */
            };
            const fetchData = async () => {
                showDCSLoading();
                const loading = createLoadingOverlay("Loading DCS data...");
                loaderHost().appendChild(loading);
                try {
                    console.log('📥 Starting fetchData...');
                    // ✅ The 'Project' column is a LOOKUP column. 
                    //    We must expand it and select its Id and Title.
                    const selectFields = [
                        "Id",
                        "Title",
                        fieldMappings.DCSData,
                        `${fieldMappings.Project}/Id`,
                        `${fieldMappings.Project}/Title`,
                    ].join(",");
                    const expandField = fieldMappings.Project; // "Project"
                    const query = `${apiBase}?$select=${selectFields}&$expand=${expandField}&$top=5000`;
                    console.log('📡 Fetching from:', query);
                    const res = await fetch(query, {
                        headers: {
                            Accept: "application/json;odata=verbose",
                        },
                        credentials: 'same-origin'
                    });
                    if (!res.ok) {
                        const errorText = await res.text();
                        console.error("API Error:", errorText);
                        throw new Error(`HTTP error! status: ${res.status}`);
                    }
                    const data = await res.json();
                    const allData = data.d.results || [];
                    console.log(`📊 Raw data received: ${allData.length} items`);
                    // Helper: read the project title from the expanded lookup object
                    const getProjectTitle = (item) => { var _a, _b; return String((_b = (_a = item === null || item === void 0 ? void 0 : item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) !== null && _b !== void 0 ? _b : "").trim(); };
                    // Helper: read the numeric project Id from the expanded lookup object
                    const getProjectId = (item) => {
                        var _a;
                        const id = (_a = item === null || item === void 0 ? void 0 : item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Id;
                        return id ? Number(id) : null;
                    };
                    // Filter to keep only ONE record per project
                    const uniqueProjects = new Map();
                    const duplicates = [];
                    console.log('🔍 Processing unique projects...');
                    allData.forEach((item) => {
                        const projectId = getProjectId(item);
                        if (projectId) {
                            if (!uniqueProjects.has(projectId)) {
                                uniqueProjects.set(projectId, item);
                            }
                            else {
                                duplicates.push(item);
                                console.warn(`Duplicate found for project ${projectId}: ID ${item.Id}`);
                            }
                        }
                        else {
                            console.warn("No matching project ID found for record:", item.Id, "| Project object:", item[fieldMappings.Project], "| Item structure:", Object.keys(item));
                        }
                    });
                    const uniqueRows = Array.from(uniqueProjects.values());
                    // 🔐 ROLE SCOPE on rows (same as SEIBIES / inventory loadForecasts):
                    // Admin / Read-All groups keep every row; other roles only keep rows for
                    // their assigned projects (getAssignedProjects returns ALL for
                    // FULL_CONTROL / READ_ALL users).
                    const scopeTitles = new Set(getAssignedProjects(projects).map(p => p.Title));
                    dcsData = uniqueRows.filter(item => {
                        const t = getProjectTitle(item);
                        return t && scopeTitles.has(t);
                    });
                    console.log(`✅ Unique projects: ${uniqueRows.length} (visible to this role: ${dcsData.length}, duplicates: ${duplicates.length})`);
                    // If we found duplicates, we could delete them here - Admin (Full Control) only
                    if (duplicates.length > 0 && hasFullControl()) {
                        console.log(`🗑️ Cleaning up ${duplicates.length} duplicates (admin)...`);
                        await cleanupDuplicates(duplicates);
                    }
                    // ✅ Initialize filters with all data - and start CLEAN
                    //    (no default filter: badge 0, no chips, Apply/Clear disabled)
                    initializeFilters();
                    appliedFilters = Object.assign({}, EMPTY_FILTERS);
                    restoreAppliedFilters();
                    syncFilterUI();
                    // ✅ filterData() renders the filtered rows AND initialises pagination
                    filterData();
                    console.log('✅ fetchData completed successfully');
                }
                catch (error) {
                    // ✅ Fix: Type-safe error handling
                    console.error("❌ Error fetching data:", error);
                    let errorMessage = "Please check your SharePoint configuration";
                    if (error instanceof Error) {
                        errorMessage = error.message;
                    }
                    else if (typeof error === 'string') {
                        errorMessage = error;
                    }
                    const tableBody = document.getElementById("dcsTableBody");
                    if (tableBody) {
                        tableBody.innerHTML = `
        <tr>
          <td colspan="12" style="text-align:center;padding:40px;color:#dc3545;">
            <div style="font-size:48px;margin-bottom:10px;">❌</div>
            <h3>Error Loading Data</h3>
            <p>${errorMessage}</p>
            <button class="dcs-btn-primary" onclick="location.reload()" style="margin-top:20px;">
              ↻ Retry
            </button>
          </td>
        </tr>
      `;
                    }
                }
                finally {
                    hideDCSLoading();
                    removeLoadingOverlay();
                }
            };
            const cleanupDuplicates = async (duplicates) => {
                try {
                    const digest = await getRequestDigest();
                    for (const duplicate of duplicates) {
                        try {
                            await fetch(`${apiBase}(${duplicate.Id})`, {
                                method: "DELETE",
                                headers: {
                                    Accept: "application/json;odata=verbose",
                                    "X-RequestDigest": digest,
                                    "IF-MATCH": "*",
                                },
                                credentials: 'same-origin'
                            });
                        }
                        catch (error) {
                            console.error(`Failed to delete duplicate ${duplicate.Id}:`, error);
                        }
                    }
                }
                catch (error) {
                    console.error("Error in cleanupDuplicates:", error);
                }
            };
            const initializeFilters = () => {
                console.log('🎛️ Initializing filters...');
                // Project filter - show all projects
                const accessibleProjectTitles = projects
                    .map((p) => p.Title)
                    .filter((title) => !!title);
                filterProject.innerHTML = '<option value="">All Projects</option>' +
                    [...new Set(accessibleProjectTitles)]
                        .sort()
                        .map(project => `<option value="${project}">${project}</option>`)
                        .join("");
                // Customer filter - customers from all projects
                const accessibleCustomers = projects
                    .map(p => p.CustomerName)
                    .filter((customer) => !!customer);
                filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                    [...new Set(accessibleCustomers)]
                        .sort()
                        .map(customer => `<option value="${customer}">${customer}</option>`)
                        .join("");
                // Project Manager filter - all managers (no role filtering)
                let accessibleManagers = [];
                accessibleManagers = projects
                    .map(p => getProjectManagerDisplayName(p.ProjectManager))
                    .filter((manager) => !!manager);
                accessibleManagers = [...new Set(accessibleManagers)].sort();
                filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                    accessibleManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");
                // Project Reference filter - all references
                const accessibleReferences = projects
                    .map(p => p.ProjectReference)
                    .filter((ref) => !!ref);
                filterProjectReference.innerHTML = '<option value="">All References</option>' +
                    [...new Set(accessibleReferences)]
                        .sort()
                        .map(ref => `<option value="${ref}">${ref}</option>`)
                        .join("");
                console.log('✅ Filters initialized');
            };
            // Filter Data
            const filterData = () => {
                try {
                    console.log('🔍 filterData called');
                    const project = filterProject.value;
                    const customer = filterCustomer.value;
                    const projectManager = filterProjectManager.value;
                    const projectReference = filterProjectReference.value;
                    let filtered = dcsData.filter((item) => {
                        var _a, _b;
                        const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                        const projectData = projects.find((p) => p.Title === projectTitle);
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
                        if (projectReference && (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) !== projectReference)
                            return false;
                        return true;
                    });
                    // Initialize pagination with filtered data
                    initializePagination(filtered);
                }
                catch (error) {
                    console.error("❌ Error filtering data:", error);
                    // Even if filter fails, initialize pagination with raw data so it doesn't freeze
                    initializePagination(dcsData);
                }
            };
            // =================== MAIN TABLE COLUMN SORTING ===================
            // Value used to sort a record for the given column key
            const getSortValue = (item, key) => {
                var _a;
                const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
                const projectData = projects.find((p) => p.Title === projectTitle);
                const summary = getDocumentStatusSummary(item);
                switch (key) {
                    case "projectReference": return ((projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "").toLowerCase();
                    case "customer": return ((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "").toLowerCase();
                    case "project": return projectTitle.toLowerCase();
                    case "projectManager": return getProjectManagerDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager).toLowerCase();
                    case "total": return summary.total;
                    case "notSubmitted": return summary.notSubmitted;
                    case "notApproved": return summary.notApproved;
                    case "revised": return summary.revisedOrApprovedWithComments;
                    case "approved": return summary.approved;
                    case "rejected": return summary.rejected;
                    case "retained": return summary.retainedForInformation;
                    default: return "";
                }
            };
            // Sort a copy of the data; empty values always go last
            const applySort = (data) => {
                if (!sortKey)
                    return data;
                const dir = sortDir === "asc" ? 1 : -1;
                return [...data].sort((a, b) => {
                    const va = getSortValue(a, sortKey);
                    const vb = getSortValue(b, sortKey);
                    const aEmpty = va === "" || va === null || va === undefined;
                    const bEmpty = vb === "" || vb === null || vb === undefined;
                    if (aEmpty && bEmpty)
                        return 0;
                    if (aEmpty)
                        return 1;
                    if (bEmpty)
                        return -1;
                    if (va < vb)
                        return -1 * dir;
                    if (va > vb)
                        return 1 * dir;
                    return 0;
                });
            };
            // One delegated listener on the (static) thead - survives tbody re-renders
            const handleSortClick = (e) => {
                const th = e.target.closest('th[data-sort]');
                if (!th)
                    return;
                const key = th.getAttribute('data-sort') || '';
                if (sortKey === key) {
                    sortDir = sortDir === 'asc' ? 'desc' : 'asc';
                }
                else {
                    sortKey = key;
                    sortDir = 'asc';
                }
                initializePagination(filteredDCSData); // re-sorts + resets to page 1
                updateSortIndicators();
            };
            const updateSortIndicators = () => {
                document.querySelectorAll('#dcsTable thead th[data-sort]').forEach((th) => {
                    const el = th;
                    el.classList.remove('dcs-sorted-asc', 'dcs-sorted-desc');
                    if (el.getAttribute('data-sort') === sortKey) {
                        el.classList.add(sortDir === 'asc' ? 'dcs-sorted-asc' : 'dcs-sorted-desc');
                    }
                });
            };
            // Format date for input (YYYY-MM-DD)
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
            // Escape HTML to prevent XSS
            const escapeHtml = (unsafe) => {
                if (unsafe === null || unsafe === undefined)
                    return "";
                const safeString = String(unsafe);
                return safeString
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;")
                    .replace(/"/g, "&quot;")
                    .replace(/'/g, "&#039;");
            };
            // Normalise a status string for comparison (case/whitespace-insensitive)
            const normalizeStatus = (status) => String(status || "").trim().toLowerCase();
            // Counts DESIGN OUTPUT documents of a record by their current status.
            // Each main document is counted once (revisions are not counted separately).
            // Legacy statuses from older records are mapped to the nearest new bucket:
            //   Pending -> Not Submitted, Under Review -> Not Approved, Revised -> Approved with comments
            // "Document Deleted/cancelled" is included in Total only.
            const getDocumentStatusSummary = (item) => {
                const summary = {
                    total: 0,
                    notSubmitted: 0,
                    notApproved: 0,
                    revisedOrApprovedWithComments: 0,
                    approved: 0,
                    rejected: 0,
                    retainedForInformation: 0
                };
                try {
                    const parsed = parseDCSData(item === null || item === void 0 ? void 0 : item[fieldMappings.DCSData]);
                    if (!parsed)
                        return summary;
                    const outputs = getOutputsFromProjectData(parsed) || [];
                    outputs.forEach((doc) => {
                        summary.total++;
                        switch (normalizeStatus(doc.status)) {
                            case "not submitted":
                            case "pending": // legacy
                            case "": // no status recorded yet
                                summary.notSubmitted++;
                                break;
                            case "submitted":
                            case "under review": // legacy
                                summary.notApproved++;
                                break;
                            case "approved with comments":
                            case "revised": // legacy
                                summary.revisedOrApprovedWithComments++;
                                break;
                            case "approved":
                                summary.approved++;
                                break;
                            case "rejected":
                                summary.rejected++;
                                break;
                            case "retained for information":
                                summary.retainedForInformation++;
                                break;
                            default:
                                // "Document Deleted/cancelled" and any unknown status -> Total only
                                break;
                        }
                    });
                }
                catch (error) {
                    console.warn("Failed to build status summary for record", item === null || item === void 0 ? void 0 : item.Id, error);
                }
                return summary;
            };
            // Renders a count cell - same look as the Procurement main table
            // (centered, bold, light tinted background per column)
            const renderCountCell = (count, extraClass = "") => `<td class="dcs-count-cell ${extraClass}${count === 0 ? " dcs-count-zero" : ""}">${count}</td>`;
            const renderTable = (data) => {
                const tableBody = document.getElementById("dcsTableBody");
                tableBody.innerHTML = "";
                if (data.length > 0) {
                    data.forEach((item) => {
                        var _a;
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        const projectId = item[`${fieldMappings.Project}Id`] ||
                            (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
                        const projectData = projects.find((p) => p.Id === projectId);
                        const row = document.createElement("tr");
                        // Get display names
                        const projectManagerDisplay = getProjectManagerDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager);
                        // Document status counts (from the record's DCS JSON - output documents)
                        const summary = getDocumentStatusSummary(item);
                        // Project Reference - Edit link ONLY when this role may edit that project
                        // (Admin: all projects; GH Head / PM Manager / Engg Lead: their own
                        // assigned projects). Other roles see the reference as plain text and
                        // can still use the View button.
                        const projectRefText = escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "") || '<span style="color:#6c757d;">-</span>';
                        const canEditThis = canEditProject(projectData);
                        const projectRefCell = canEditThis
                            ? `<a href="#" class="dcs-edit-link dcs-edit-btn" data-id="${item.Id}" title="Edit document data">${projectRefText}</a>`
                            : `<span title="${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "")}">${projectRefText}</span>`;
                        row.innerHTML = `
              <td class="dcs-project-ref-cell dcs-sticky-main-col dcs-main-col-1" title="${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "")}">${projectRefCell}</td>
              <td class="dcs-sticky-main-col dcs-main-col-2" title="${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "")}">${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "")}</td>
              <td class="dcs-sticky-main-col dcs-main-col-3" title="${escapeHtml(projectTitle)}">${escapeHtml(projectTitle)}</td>
              <td class="dcs-sticky-main-col dcs-main-col-4" title="${escapeHtml(projectManagerDisplay)}">${escapeHtml(projectManagerDisplay)}</td>
              ${renderCountCell(summary.total, "dcs-count-total")}
              ${renderCountCell(summary.notSubmitted, "dcs-count-not-submitted")}
              ${renderCountCell(summary.notApproved, "dcs-count-not-approved")}
              ${renderCountCell(summary.revisedOrApprovedWithComments, "dcs-count-revised")}
              ${renderCountCell(summary.approved, "dcs-count-approved")}
              ${renderCountCell(summary.rejected, "dcs-count-rejected")}
              ${renderCountCell(summary.retainedForInformation, "dcs-count-retained")}
              <td class="dcs-actions-cell">
                <button class="dcs-view-btn" data-id="${item.Id}">👁️ View</button>
              </td>
            `;
                        tableBody.appendChild(row);
                    });
                }
                else {
                    const noDataRow = document.createElement("tr");
                    noDataRow.innerHTML = `<td colspan="12" style="text-align:center;padding:40px;color:#6c757d;">
            <div style="font-size:48px;margin-bottom:10px;">📋</div>
            <h3>No document data found</h3>
          </td>`;
                    tableBody.appendChild(noDataRow);
                }
                // Add event listeners
                // Project Reference link opens the edit form
                document.querySelectorAll(".dcs-edit-btn").forEach((btn) => {
                    btn.addEventListener("click", (e) => {
                        e.preventDefault();
                        editData(Number(btn.dataset.id));
                    });
                });
                document.querySelectorAll(".dcs-view-btn").forEach((btn) => {
                    btn.addEventListener("click", () => viewData(Number(btn.dataset.id)));
                });
            };
            // Get inputs / outputs from project data
            const getInputsFromProjectData = (data) => {
                if (!data)
                    return [];
                if ('documentStructure' in data && data.documentStructure) {
                    return data.documentStructure.inputs || [];
                }
                if ('inputs' in data) {
                    return data.inputs || [];
                }
                return [];
            };
            const getOutputsFromProjectData = (data) => {
                if (!data)
                    return [];
                if ('documentStructure' in data && data.documentStructure) {
                    return data.documentStructure.outputs || [];
                }
                if ('outputs' in data) {
                    return data.outputs || [];
                }
                return [];
            };
            const convertToNewFormat = (data) => {
                if ('documentStructure' in data) {
                    return data;
                }
                // Convert from old format to new format
                const legacyData = data;
                return {
                    projectId: legacyData.projectId,
                    documentStructure: {
                        inputs: legacyData.inputs || [],
                        outputs: legacyData.outputs || []
                    }
                };
            };
            const editData = (id) => {
                const record = dcsData.find((d) => d.Id === id);
                if (!record) {
                    toast("Record not found", "error");
                    return;
                }
                // 🔐 ROLE GATE: only users who may edit THIS project can open the edit form
                const gateProjectId = record[`${fieldMappings.Project}Id`] ||
                    (record[fieldMappings.Project] && record[fieldMappings.Project].Id);
                const gateProject = projects.find(p => p.Id === gateProjectId);
                if (!canEditProject(gateProject)) {
                    toast("You do not have permission to edit this project's documents", "error");
                    return;
                }
                // REMOVE view mode class
                formSection.classList.remove("dcs-view-mode");
                // Set editing mode
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formTitle.textContent = "";
                formSection.classList.remove("view-mode");
                projectSelect.disabled = true;
                const projectId = record[`${fieldMappings.Project}Id`] ||
                    (record[fieldMappings.Project] && record[fieldMappings.Project].Id);
                // Set title
                setEditTitle(projects.find(p => p.Id === projectId) || null);
                if (projectId) {
                    projectSelect.value = projectId.toString();
                    updateProjectDetails(projectId);
                }
                // Show the container table
                containerTable.style.display = "block";
                importButtonContainer.style.display = "flex";
                existingDataAlert.style.display = "none";
                // Parse and load the JSON data
                const parsedData = parseDCSData(record[fieldMappings.DCSData]);
                if (parsedData) {
                    const inputs = getInputsFromProjectData(parsedData);
                    const outputs = getOutputsFromProjectData(parsedData);
                    inputTableBody.innerHTML = '';
                    outputTableBody.innerHTML = '';
                    // Populate input documents
                    if (inputs.length > 0) {
                        inputs.forEach((doc, index) => {
                            const newRow = createEditableDocumentRow(doc, index + 1, true);
                            inputTableBody.appendChild(newRow);
                            if (doc.revisions && doc.revisions.length > 0) {
                                doc.revisions.forEach((revision) => {
                                    const revisionHTML = buildRevisionRowHTML(doc.docId, revision);
                                    inputTableBody.insertAdjacentHTML('beforeend', revisionHTML);
                                });
                                updateMergedCells(doc.docId);
                            }
                        });
                    }
                    else {
                        const defaultInputDoc = {
                            docId: `default_input_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                            docNo: "DOC-001", description: "", category: "A", enggDivision: "",
                            revNo: 0, originalPlan: "", revisedPlan: "", actualSubmission: "",
                            submissionTransmittal: "", approvalDate: "", approvalTransmittal: "",
                            status: "Pending", remarks: "", isInput: true, revisions: [], docType: 'input'
                        };
                        inputTableBody.appendChild(createEditableDocumentRow(defaultInputDoc, 1, true));
                    }
                    // Populate output documents
                    if (outputs.length > 0) {
                        outputs.forEach((doc, index) => {
                            const newRow = createEditableDocumentRow(doc, index + 1, false);
                            outputTableBody.appendChild(newRow);
                            if (doc.revisions && doc.revisions.length > 0) {
                                doc.revisions.forEach((revision) => {
                                    const revisionHTML = buildRevisionRowHTML(doc.docId, revision);
                                    outputTableBody.insertAdjacentHTML('beforeend', revisionHTML);
                                });
                                updateMergedCells(doc.docId);
                            }
                        });
                    }
                    else {
                        const defaultOutputDoc = {
                            docId: `default_output_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                            docNo: "DOC-OUT-001", description: "", category: "A", enggDivision: "",
                            revNo: 0, originalPlan: "", revisedPlan: "", actualSubmission: "",
                            submissionTransmittal: "", approvalDate: "", approvalTransmittal: "",
                            status: "Not Submitted", remarks: "", isInput: false, revisions: [], docType: 'output'
                        };
                        outputTableBody.appendChild(createEditableDocumentRow(defaultOutputDoc, 1, false));
                    }
                }
                else {
                    resetDocumentTables();
                }
                // Make sure buttons are visible
                if (addInputDocBtn)
                    addInputDocBtn.style.display = "inline-block";
                if (addOutputDocBtn)
                    addOutputDocBtn.style.display = "inline-block";
                // Show save button
                getSaveButtons().forEach(btn => {
                    btn.style.display = "inline-flex";
                    btn.innerHTML = '💾 Update Document Data';
                });
                containerTable.style.display = "block";
                importButtonContainer.style.display = "flex";
                // Show import instructions
                const importInstructions = document.getElementById('dcsImportInstructions');
                if (importInstructions) {
                    importInstructions.style.display = 'block';
                }
                // ✅ Re-attach document event listeners (cleaned up)
                setTimeout(() => {
                    attachDocumentEventListeners();
                    updateAllSerialNumbers();
                    ensureDocumentButtonsVisible();
                }, 100);
                applyRoleBasedUI();
            };
            // ✅ SHARED revision row builder - exactly 8 cells (7 fields + actions) so
            // revision rows line up with the 15-column grid when the header merges
            // Sr No/Doc No/Desc/Category/Division/Status/Remarks (Actions stays unmerged).
            // ✅ SHARED revision row builder - Must have 15 cells total to align with the header grid.
            // The merged columns (Sr No, Doc No, Desc, Category, Division, Status, Remarks) 
            // are represented by hidden <td>s so the unmerged cells (Dates, Transmittals, Actions) line up perfectly.
            const buildRevisionRowHTML = (docId, revision) => `
  <tr class="dcs-document-group" data-doc-id="${docId}">
    <!-- 1. Sr No (Merged - hidden) -->
    <td style="display: none;"></td>
    <!-- 2. Doc No (Merged - hidden) -->
    <td style="display: none;"></td>
    <!-- 3. Description (Merged - hidden) -->
    <td style="display: none;"></td>
    <!-- 4. Category (Merged - hidden) -->
    <td style="display: none;"></td>
    <!-- 5. Engg Division (Merged - hidden) -->
    <td style="display: none;"></td>
    
    <!-- 6. Rev No (Unmerged - visible) -->
    <td class="dcs-rev-cell">${revision.revNumber}</td>
    
    <!-- 7. Original Plan (Unmerged - visible) -->
    <td><input type="date" class="dcs-form-control dcs-date-input"
           value="${formatDateForInput(revision.originalPlan)}"
           placeholder="Original Plan"></td>
           
    <!-- 8. Revised Plan (Unmerged - visible) -->
    <td><input type="date" class="dcs-form-control dcs-date-input"
           value="${formatDateForInput(revision.revisedPlan)}"
           placeholder="Revised Plan"></td>
           
    <!-- 9. Actual Submission (Unmerged - visible) -->
    <td><input type="date" class="dcs-form-control dcs-date-input"
           value="${formatDateForInput(revision.actualSubmission)}"
           placeholder="Actual Submission"></td>
           
    <!-- 10. Submission Transmittal (Unmerged - visible) -->
    <td>
      <textarea class="dcs-form-control dcs-transmittal-input" rows="2"
            placeholder="Transmittal No">${escapeHtml(revision.submissionTransmittal)}</textarea>
    </td>
    
    <!-- 11. Approval Date (Unmerged - visible) -->
    <td><input type="date" class="dcs-form-control dcs-date-input"
           value="${formatDateForInput(revision.approvalDate)}"
           placeholder="Approval Date"></td>
           
    <!-- 12. Approval Transmittal (Unmerged - visible) -->
    <td>
      <textarea class="dcs-form-control dcs-transmittal-input" rows="2"
            placeholder="Transmittal No">${escapeHtml(revision.approvalTransmittal)}</textarea>
    </td>
    
    <!-- 13. Status (Merged - hidden) -->
    <td style="display: none;"></td>
    <!-- 14. Remarks (Merged - hidden) -->
    <td style="display: none;"></td>
    
    <!-- 15. Actions (Unmerged - visible) -->
    <td class="dcs-doc-actions">
      <button type="button" class="dcs-btn-danger dcs-remove-revision-btn"
              title="Remove Revision">➖</button>
    </td>
  </tr>
`;
            //       const updateFormActionsForEditMode = (): void => {
            //         const formActions = document.querySelector('.dcs-form-actions');
            //         if (!formActions) return;
            //         formActions.innerHTML = `
            //   <div style="display: flex; gap: 10px; flex-wrap: wrap;">
            //     <button type="submit" form="dcsForm" class="dcs-btn-primary">
            //       💾 Save Changes
            //     </button>
            //     <button type="button" class="dcs-btn-secondary dcs-back-to-table">
            //       ← Cancel
            //     </button>
            //   </div>
            // `;
            //         // ✅ Sync the top button with the new bottom button
            //         getSaveButtons().forEach(btn => {
            //           btn.style.display = "inline-flex";
            //         });
            //         // Reattach back button listener
            //         const backBtn = formActions.querySelector('.dcs-back-to-table');
            //         if (backBtn) {
            //           backBtn.addEventListener('click', handleBackToTable);
            //         }
            //       };
            const createEditableDocumentRow = (doc, srNo, isInput) => {
                var _a;
                const docId = doc.docId || `doc_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
                const prefix = isInput ? 'DOC' : 'DOC-OUT';
                const docNo = doc.docNo || `${prefix}-${String(srNo).padStart(3, '0')}`;
                // Set correct transmittal placeholders based on document type
                const submissionPlaceholder = isInput ? 'Transmittal No' : 'Transmittal No';
                const approvalPlaceholder = isInput ? 'Transmittal No' : 'Transmittal No';
                const row = document.createElement("tr");
                row.className = "dcs-document-group dcs-document-header";
                row.setAttribute("data-doc-id", docId);
                const totalRevRows = (((_a = doc.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0) + 1;
                row.setAttribute("data-rev-rows", totalRevRows.toString());
                const tableType = isInput ? 'input' : 'output';
                // Create the row with 15 columns total
                row.innerHTML = `
          <!-- Column 1: Sr No -->
          <td class="dcs-sr-no">${srNo}</td>
          <!-- Column 2: Document No -->
          <td>
            <textarea class="dcs-form-control dcs-doc-no" rows="2" placeholder="Document No">${escapeHtml(docNo)}</textarea>
          </td>
          <!-- Column 3: Description -->
          <td>
            <textarea class="dcs-form-control dcs-description-input" rows="3" placeholder="Enter description">${escapeHtml(doc.description)}</textarea>
          </td>
          <!-- Column 4: Category -->
          <td>
            <select class="dcs-form-control dcs-category-select">
              <option value="A" ${doc.category === "A" ? "selected" : ""}>A - Approval</option>
              <option value="I" ${doc.category === "I" ? "selected" : ""}>I - Information</option>
            </select>
          </td>
          <!-- Column 5: Engg Division -->
          <td>
            <select class="dcs-form-control dcs-engg-division-select">
              ${buildEnggDivisionOptionsHTML(doc.enggDivision || "")}
            </select>
          </td>
          <!-- Column 6: Rev No -->
          <td class="dcs-rev-cell">0</td>
          <!-- Column 7: Original Plan -->
          <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.originalPlan)}" placeholder="Original Plan"></td>
          <!-- Column 8: Revised Plan -->
          <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.revisedPlan)}" placeholder="Revised Plan"></td>
          <!-- Column 9: Actual Submission -->
          <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.actualSubmission)}" placeholder="Actual Submission"></td>
          <!-- Column 10: Submission Transmittal -->
          <td>
            <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${submissionPlaceholder}">${escapeHtml(doc.submissionTransmittal)}</textarea>
          </td>
          <!-- Column 11: Approval Date -->
          <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.approvalDate)}" placeholder="Approval Date"></td>
          <!-- Column 12: Approval Transmittal -->
          <td>
            <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${approvalPlaceholder}">${escapeHtml(doc.approvalTransmittal)}</textarea>
          </td>
          <!-- Column 13: Status -->
          <td>
            <select class="dcs-form-control dcs-status-select">
              ${buildStatusOptionsHTML(tableType, doc.status)}
            </select>
          </td>
          <!-- Column 14: Remarks -->
          <td>
            <textarea class="dcs-form-control dcs-remarks-input" rows="3" placeholder="Enter remarks">${escapeHtml(doc.remarks)}</textarea>
          </td>
          <!-- Column 15: Actions -->
          <td class="dcs-doc-actions">
            <button type="button" class="dcs-btn-success dcs-add-revision-btn" title="Add Revision">
              <span style="font-size: 16px;">➕</span>
            </button>
            <button type="button" class="dcs-btn-danger dcs-remove-doc" title="Remove Document">
              <span style="font-size: 16px;">🗑️</span>
            </button>
          </td>
        `;
                return row;
            };
            const viewData = (id) => {
                if (!canRead()) {
                    toast("You do not have permission to view document data", "error");
                    return;
                }
                const record = dcsData.find((d) => d.Id === id);
                if (!record) {
                    toast("Record not found", "error");
                    return;
                }
                const projectId = record[`${fieldMappings.Project}Id`] ||
                    (record[fieldMappings.Project] && record[fieldMappings.Project].Id);
                formSection.classList.add("dcs-view-mode");
                formTitle.textContent = "Document Control Index";
                formSection.classList.remove("dcs-edit-mode");
                formSection.style.display = "block";
                tableSection.style.display = "none";
                if (projectId) {
                    projectSelect.value = projectId.toString();
                }
                projectSelect.disabled = true;
                existingDataAlert.style.display = "none";
                importButtonContainer.style.display = "none";
                // ✅ Remove the Save button from the DOM completely
                const saveWrapper = document.querySelector(".dcs-top-save-wrapper");
                if (saveWrapper) {
                    saveWrapper.style.display = "none";
                }
                // Hide bottom save buttons too
                getSaveButtons().forEach(btn => {
                    btn.style.display = "none";
                });
                inputTableBody.innerHTML = '';
                outputTableBody.innerHTML = '';
                if (projectId) {
                    const project = projects.find(p => p.Id === projectId);
                    if (project) {
                        setProjectDetailsDisplay(project);
                    }
                    else {
                        toast("Project details not found", "error");
                    }
                }
                else {
                    toast("Project information missing", "error");
                }
                const parsedData = parseDCSData(record[fieldMappings.DCSData]);
                if (parsedData) {
                    containerTable.style.display = "block";
                    loadDataIntoReadonlyForm(parsedData);
                }
                else {
                    containerTable.style.display = "block";
                    inputTableBody.innerHTML = `
            <tr>
              <td colspan="14" style="text-align:center; padding:40px; color:#6c757d;">
                <div style="font-size: 48px; margin-bottom: 10px;">📄</div>
                <h3>No Document Data</h3>
                <p>This project doesn't have any document data yet.</p>
              </td>
            </tr>
          `;
                    outputTableBody.innerHTML = `
            <tr>
              <td colspan="14" style="text-align:center; padding:40px; color:#6c757d;">
                <div style="font-size: 48px; margin-bottom: 10px;">📄</div>
                <h3>No Document Data</h3>
                <p>This project doesn't have any document data yet.</p>
              </td>
            </tr>
          `;
                }
                setupViewModeActions(projectId);
                applyRoleBasedUI();
            };
            // Create readonly document row that matches edit mode structure exactly - FIXED (without placeholders, left aligned)
            const createReadonlyDocumentRow = (doc, srNo, tableType) => {
                const row = document.createElement("tr");
                row.className = "dcs-document-group dcs-document-header";
                row.setAttribute("data-doc-id", doc.docId);
                let scheduleBgColor;
                let submissionBgColor;
                let approvalBgColor;
                if (tableType === 'input') {
                    scheduleBgColor = '#e8f5e9'; // Green for input schedule
                    submissionBgColor = '#fff3cd'; // Yellow for submission
                    approvalBgColor = '#d4edda'; // Green for approval
                }
                else {
                    scheduleBgColor = '#e3f2fd'; // Blue for output schedule
                    submissionBgColor = '#fff3cd'; // Yellow for submission
                    approvalBgColor = '#d4edda'; // Green for approval
                }
                // Helper to format dates for display
                const formatDateForDisplay = (dateString) => {
                    if (!dateString)
                        return "";
                    try {
                        const date = new Date(dateString);
                        return isNaN(date.getTime()) ? dateString : date.toLocaleDateString("en-GB");
                    }
                    catch (_a) {
                        return dateString;
                    }
                };
                // Status badge styling
                const getStatusBadgeStyle = (status) => {
                    const statusStyles = {
                        // Output statuses
                        'Not Submitted': 'background: #e2e3e5; color: #383d41; border: 1px solid #d6d8db;',
                        'Submitted': 'background: #cce5ff; color: #004085; border: 1px solid #b8daff;',
                        'Approved': 'background: #d4edda; color: #155724; border: 1px solid #c3e6cb;',
                        'Approved with comments': 'background: #e2f0d9; color: #2e6b1f; border: 1px solid #b7dfa8;',
                        'Rejected': 'background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb;',
                        'Retained for Information': 'background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb;',
                        'Document Deleted/cancelled': 'background: #6c757d; color: #ffffff; border: 1px solid #5a6268;',
                        // Input statuses
                        'Pending': 'background: #fff3cd; color: #856404; border: 1px solid #ffeaa7;',
                        'Received': 'background: #d4edda; color: #155724; border: 1px solid #c3e6cb;',
                        'Partially Received': 'background: #ffe5b4; color: #8a5a00; border: 1px solid #ffd280;',
                        'Revised inputs required': 'background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb;',
                        // Legacy statuses (older records)
                        'Under Review': 'background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb;',
                        'Revised': 'background: #e2e3e5; color: #383d41; border: 1px solid #d6d8db;'
                    };
                    return statusStyles[status] || 'background: #f8f9fa; color: #6c757d; border: 1px solid #e9ecef;';
                };
                // Create the row with 14 columns (same as edit mode but without Actions column)
                row.innerHTML = `
          <!-- Column 1: Sr No (width: 40px) -->
          <td class="dcs-sr-no" style="width: 40px; text-align: center; vertical-align: middle; background: #f8f9fa; font-weight: 500; padding: 8px;">${srNo}</td>

          <!-- Column 2: Document No (min-width: 150px) -->
          <td style="min-width: 150px; vertical-align: top; padding: 8px;">
            <div class="dcs-view-field" style="
              background: #e8f4fd; 
              border: 1px solid #b8daff; 
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 40px;
              font-size: 13px;
              white-space: nowrap;
              word-wrap: break-word;
              text-align: left;
            ">
              ${escapeHtml(doc.docNo)}
            </div>
          </td>

          <!-- Column 3: Description (min-width: 180px) -->
          <td style="min-width: 180px; vertical-align: top; padding: 8px;">
            <div class="dcs-view-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 60px;
              font-size: 13px;
              white-space: nowrap;
              word-wrap: break-word;
              background: white;
              text-align: left;
            ">
              ${escapeHtml(doc.description) || '<span style="color: #6c757d; font-style: italic;">No description</span>'}
            </div>
          </td>

          <!-- Column 4: Category (width: 100px) -->
          <td style="width: 100px; text-align: center; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field" style="
              display: inline-block;
              padding: 4px 8px;
              border-radius: 4px;
              font-size: 12px;
              font-weight: 500;
              ${doc.category === 'A' ? 'background: #d4edda; color: #155724;' : 'background: #e2e3e5; color: #383d41;'}
            ">
              ${doc.category === 'A' ? 'A - Approval' : 'I - Information'}
            </div>
          </td>

          <!-- Column 5: Engg Division (width: 130px) -->
          <td style="width: 130px; text-align: center; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field" style="
              display: inline-block;
              padding: 4px 8px;
              border-radius: 4px;
              font-size: 12px;
              font-weight: 500;
              background: #e8f4fd;
              color: #0d6efd;
              border: 1px solid #b8daff;
            ">
              ${escapeHtml(doc.enggDivision || '') || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 6: Rev No (width: 80px) -->
          <td class="dcs-rev-cell" style="width: 80px; text-align: center; vertical-align: middle; padding: 8px; background: #f8f9fa; font-weight: 500;">0</td>

          <!-- Column 7: Original Plan (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${scheduleBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(doc.originalPlan) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 8: Revised Plan (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${scheduleBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(doc.revisedPlan) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 9: Actual Submission (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${submissionBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(doc.actualSubmission) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 10: Submission Transmittal (width: 150px) -->
          <td style="width: 150px; vertical-align: top; padding: 8px;">
            <div class="dcs-view-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 40px;
              font-size: 13px;
              background: ${submissionBgColor};
              white-space: nowrap;
              word-wrap: break-word;
              text-align: left;
            "> 
              ${escapeHtml(doc.submissionTransmittal) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 11: Approval Date (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${approvalBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(doc.approvalDate) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 12: Approval Transmittal (width: 150px) -->
          <td style="width: 150px; vertical-align: top; padding: 8px;">
            <div class="dcs-view-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 40px;
              font-size: 13px;
              background: ${approvalBgColor};
              white-space: nowrap;
              word-wrap: break-word;
              text-align: left;
            "> 
              ${escapeHtml(doc.approvalTransmittal) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>  

          <!-- Column 13: Status (width: 120px) -->
          <td style="width: 120px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field" style="text-align: center;">
              <span class="dcs-status-badge" data-status="${escapeHtml(doc.status)}" 
                    style="
                      display: inline-block;
                      padding: 4px 8px;
                      border-radius: 4px;
                      font-size: 12px;
                      font-weight: 500;
                      ${getStatusBadgeStyle(doc.status)}
                      min-width: 80px;
                    ">
                ${escapeHtml(doc.status)}
              </span>
            </div>
          </td>

          <!-- Column 14: Remarks (width: 150px) - Increased width since no Actions column -->
          <td style="width: 150px; vertical-align: top; padding: 8px;">
            <div class="dcs-view-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 60px;
              font-size: 13px;
              background: white;
              white-space: nowrap;
              word-wrap: break-word;
              text-align: left;
            ">
              ${escapeHtml(doc.remarks) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>
        `;
                return row;
            };
            // Create readonly revision row - FIXED to match edit mode structure (without placeholders, left aligned)
            const createReadonlyRevisionRow = (revision, doc, tableType) => {
                const revRow = document.createElement("tr");
                revRow.className = "dcs-document-group dcs-revision-row";
                let scheduleBgColor;
                let submissionBgColor;
                let approvalBgColor;
                if (tableType === 'input') {
                    scheduleBgColor = '#e8f5e9';
                    submissionBgColor = '#fff3cd';
                    approvalBgColor = '#d4edda';
                }
                else {
                    scheduleBgColor = '#e3f2fd';
                    submissionBgColor = '#fff3cd';
                    approvalBgColor = '#d4edda';
                }
                // Helper to format dates for display
                const formatDateForDisplay = (dateString) => {
                    if (!dateString)
                        return "";
                    try {
                        const date = new Date(dateString);
                        return isNaN(date.getTime()) ? dateString : date.toLocaleDateString("en-GB");
                    }
                    catch (_a) {
                        return dateString;
                    }
                };
                // Revision rows in view mode have 7 columns (Rev No + 6 date/transmittal columns)
                revRow.innerHTML = `
          <!-- Column 1: Rev No (width: 80px) -->
          <td class="dcs-rev-cell" style="width: 80px; text-align: center; vertical-align: middle; padding: 8px; background: #fff3cd; font-weight: 500;">
            ${revision.revNumber}
          </td>

          <!-- Column 2: Original Plan (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${scheduleBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(revision.originalPlan) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 3: Revised Plan (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${scheduleBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(revision.revisedPlan) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 4: Actual Submission (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${submissionBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(revision.actualSubmission) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 5: Submission Transmittal (width: 150px) -->
          <td style="width: 150px; vertical-align: top; padding: 8px;">
            <div class="dcs-view-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 40px;
              font-size: 13px;
              background: ${submissionBgColor};
              white-space: nowrap;
              word-wrap: break-word;
              text-align: left;
            ">
              
              ${escapeHtml(revision.submissionTransmittal) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 6: Approval Date (width: 150px) -->
          <td style="width: 150px; vertical-align: middle; padding: 8px;">
            <div class="dcs-view-field dcs-date-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 36px;
              font-size: 13px;
              background: ${approvalBgColor};
              display: flex;
              align-items: center;
              justify-content: flex-start;
            ">
              ${formatDateForDisplay(revision.approvalDate) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>

          <!-- Column 7: Approval Transmittal (width: 150px) -->
          <td style="width: 150px; vertical-align: top; padding: 8px;">
            <div class="dcs-view-field" style="
              border: 1px solid #dee2e6;
              border-radius: 4px;
              padding: 8px 10px;
              min-height: 40px;
              font-size: 13px;
              background: ${approvalBgColor};
              white-space: nowrap;
              word-wrap: break-word;
              text-align: left;
            ">
              
              ${escapeHtml(revision.approvalTransmittal) || '<span style="color: #6c757d;">-</span>'}
            </div>
          </td>
        `;
                return revRow;
            };
            // Load data into readonly form - UPDATED to match edit mode structure (without Actions)
            const loadDataIntoReadonlyForm = (projectData) => {
                // normalise legacy shapes so both formats render the same way
                const normalised = convertToNewFormat(projectData);
                const inputs = getInputsFromProjectData(normalised);
                const outputs = getOutputsFromProjectData(normalised);
                // Load input documents
                if (inputs.length > 0) {
                    inputs.forEach((doc, index) => {
                        const docRow = createReadonlyDocumentRow(doc, index + 1, 'input');
                        inputTableBody.appendChild(docRow);
                        // Add revision rows if they exist
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach((revision, revIndex) => {
                                const revRow = createReadonlyRevisionRow(revision, doc, 'input');
                                inputTableBody.appendChild(revRow);
                            });
                            // Set rowspan for merged cells in main document row
                            // In view mode without Actions: Sr No, Document No, Description, Category, Status, Remarks are merged
                            const cellsToMerge = [0, 1, 2, 3, 4, 12, 13]; // 0-indexed (14 columns total, index 0-13) - Sr No, Doc No, Description, Category, Engg Division, Status, Remarks
                            cellsToMerge.forEach(cellIndex => {
                                const cell = docRow.cells[cellIndex];
                                if (cell) {
                                    cell.rowSpan = doc.revisions.length + 1;
                                }
                            });
                        }
                    });
                }
                else {
                    inputTableBody.innerHTML = `
            <tr>
              <td colspan="14" style="text-align:center; padding:20px; color:#6c757d; background:#f8f9fa;">
                No input documents found
              </td>
            </tr>
          `;
                }
                // Load output documents
                if (outputs.length > 0) {
                    outputs.forEach((doc, index) => {
                        const docRow = createReadonlyDocumentRow(doc, index + 1, 'output');
                        outputTableBody.appendChild(docRow);
                        // Add revision rows if they exist
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach((revision, revIndex) => {
                                const revRow = createReadonlyRevisionRow(revision, doc, 'output');
                                outputTableBody.appendChild(revRow);
                            });
                            // Set rowspan for merged cells in main document row
                            const cellsToMerge = [0, 1, 2, 3, 4, 12, 13]; // 0-indexed (14 columns total, index 0-13) - Sr No, Doc No, Description, Category, Engg Division, Status, Remarks
                            cellsToMerge.forEach(cellIndex => {
                                const cell = docRow.cells[cellIndex];
                                if (cell) {
                                    cell.rowSpan = doc.revisions.length + 1;
                                }
                            });
                        }
                    });
                }
                else {
                    outputTableBody.innerHTML = `
            <tr>
              <td colspan="14" style="text-align:center; padding:20px; color:#6c757d; background:#f8f9fa;">
                No output documents found
              </td>
            </tr>
          `;
                }
                // Add responsive table container
                containerTable.style.overflowX = "auto";
                containerTable.style.display = "block";
            };
            // Update the export button event handler in setupViewModeActions
            const setupViewModeActions = (projectId) => {
                const formActions = document.querySelector('.dcs-form-actions');
                if (!formActions)
                    return;
                formActions.innerHTML = `
          <div style="display: flex; gap: 10px; flex-wrap: wrap;">
            <button type="button" class="dcs-btn-secondary dcs-back-to-table">
              ← Back to List
            </button>

            <button type="button" class="dcs-btn-primary dcs-export-document-btn" 
                    id="viewModeExportBtn">
              📤 Export Document
            </button>
          </div>
        `;
                // Add event listeners
                const backBtn = formActions.querySelector('.dcs-back-to-table');
                if (backBtn) {
                    backBtn.addEventListener('click', () => { void handleBackToTable(); });
                }
                // Export button - export current project only
                const exportBtn = document.getElementById('viewModeExportBtn');
                if (exportBtn) {
                    exportBtn.addEventListener('click', exportAllDocuments);
                }
            };
            const exportAllDocuments = async () => {
                showExportProgress("Preparing document data for export...");
                try {
                    // Get the current project ID from view mode
                    const projectSelect = document.getElementById("dcsProjectSelect");
                    const projectId = parseInt(projectSelect.value);
                    if (!projectId) {
                        hideExportProgress();
                        toast("No project selected", "warning");
                        return;
                    }
                    // Find project and data
                    const project = projects.find(p => p.Id === projectId);
                    if (!project) {
                        hideExportProgress();
                        toast("Project not found", "error");
                        return;
                    }
                    const record = dcsData.find((item) => {
                        const projId = item[`${fieldMappings.Project}Id`] ||
                            (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
                        return projId === projectId;
                    });
                    if (!record || !record[fieldMappings.DCSData]) {
                        hideExportProgress();
                        toast("No document data found", "warning");
                        return;
                    }
                    const parsedData = parseDCSData(record[fieldMappings.DCSData]);
                    if (!parsedData) {
                        hideExportProgress();
                        toast("No document data to export", "warning");
                        return;
                    }
                    const outputs = getOutputsFromProjectData(parsedData) || [];
                    const inputs = getInputsFromProjectData(parsedData) || [];
                    if (outputs.length === 0 && inputs.length === 0) {
                        hideExportProgress();
                        toast("No documents to export", "warning");
                        return;
                    }
                    showExportProgress("Creating Excel with two sheets...");
                    // Create workbook
                    const wb = XLSX.utils.book_new();
                    // =================== SHEET 1: COVER SHEET ===================
                    const coverSheetData = [];
                    // ========== ADD 3 NEW ROWS AT THE TOP ==========
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]); // Row 1
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]); // Row 2
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]); // Row 3
                    // ========== SECTION 1: APPROVAL BLOCK (11 columns A-K) ==========
                    const approvalBlockData = [
                        // Row 1-8: NAME/DATE grid (11 columns)
                        ["", "", "NAME", "", "", "NAME", "", "", "NAME", "", ""],
                        ["", "", "DATE", "", "", "DATE", "", "", "DATE", "", ""],
                        ["", "", "NAME", "", "", "NAME", "", "", "NAME", "", ""],
                        ["", "", "DATE", "", "", "DATE", "", "", "DATE", "", ""],
                        ["", "", "NAME", "", "", "NAME", "", "", "NAME", "", ""],
                        ["", "", "DATE", "", "", "DATE", "", "", "DATE", "", ""],
                        ["", "", "NAME", "", "", "NAME", "", "", "NAME", "", ""],
                        ["", "", "DATE", "", "", "DATE", "", "", "DATE", "", ""],
                        // Row 9: "0 For Approval"
                        ["0", "For Approval", "NAME", "", "", "NAME", "", "", "NAME", "", ""],
                        // Row 10: DATE row
                        ["", "", "DATE", "", "", "DATE", "", "", "DATE", "", ""],
                        // Row 11: Header row (A-K: 11 columns)
                        ["REV.", "DESCRIPTION", "PREPARED BY", "", "SIGN", "CHECKED BY", "", "SIGN", "APPROVED BY", "", "SIGN"],
                        // Empty rows as spacer
                        ["", "", "", "", "", "", "", "", "", "", ""],
                    ];
                    // Add approval block
                    coverSheetData.push(...approvalBlockData);
                    // ========== SECTION 2: PROJECT INFORMATION ==========
                    const abbProjectNo = project.ProjectReference || "";
                    const abbDocumentNo = "";
                    const prPoNo = "";
                    coverSheetData.push(["PROJECT TITLE:", "", project.Title || "", "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["DOCUMENT TITLE:", "", "DOCUMENT CONTROL INDEX", "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["ABB PROJECT NO:", "", abbProjectNo, "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["ABB DOCUMENT NO:", "", abbDocumentNo, "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["PR / PO NO:", "", prPoNo, "", "", "", "", "", "", "", ""]);
                    for (let i = 0; i < 2; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    }
                    // ========== SECTION 3: PMC, CUSTOMER, END USER, SUPPLIER ==========
                    const pmc = "";
                    coverSheetData.push(["PMC:", "", pmc, "", "", "", "", "", "", "", ""]);
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    }
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    const customerName = project.CustomerName || "";
                    const endUser = "";
                    const supplier = "";
                    coverSheetData.push(["CUSTOMER:", "", customerName, "", "", "", "", "", "", "", ""]);
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    }
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["END USER:", "", endUser, "", "", "", "", "", "", "", ""]);
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    }
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["SUPPLIER:", "", supplier, "", "", "", "", "", "", "", ""]);
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    }
                    // Create Cover Sheet worksheet
                    const coverSheetWs = XLSX.utils.aoa_to_sheet(coverSheetData);
                    // ========== STYLING ==========
                    const borderStyle = {
                        top: { style: "thin" },
                        bottom: { style: "thin" },
                        left: { style: "thin" },
                        right: { style: "thin" }
                    };
                    const fullDataRange = XLSX.utils.decode_range(coverSheetWs["!ref"] || "A1");
                    for (let R = 0; R < 3; ++R) {
                        for (let C = 0; C <= 10; ++C) {
                            const address = XLSX.utils.encode_cell({ r: R, c: C });
                            if (!coverSheetWs[address]) {
                                coverSheetWs[address] = { t: 's', v: '' };
                            }
                            if (!coverSheetWs[address].s)
                                coverSheetWs[address].s = {};
                            coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "FFFFFF" } };
                        }
                    }
                    for (let R = 0; R <= fullDataRange.e.r; ++R) {
                        for (let C = 0; C <= 10; ++C) {
                            const address = XLSX.utils.encode_cell({ r: R, c: C });
                            if (!coverSheetWs[address]) {
                                coverSheetWs[address] = { t: 's', v: '' };
                            }
                            if (!coverSheetWs[address].s)
                                coverSheetWs[address].s = {};
                            coverSheetWs[address].s.border = borderStyle;
                            const cellValue = coverSheetWs[address].v;
                            if (R === 14) {
                                coverSheetWs[address].s.font = { bold: true, sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "center", vertical: "center" };
                                coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "D9D9D9" } };
                            }
                            if (R === 12 && C === 0 && cellValue === "0") {
                                coverSheetWs[address].s.font = { bold: true, sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "center", vertical: "center" };
                                coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "E6F3FF" } };
                            }
                            if (R === 12 && C === 1 && cellValue === "For Approval") {
                                coverSheetWs[address].s.font = { bold: true, sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "center", vertical: "center" };
                                coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "E6F3FF" } };
                            }
                            if (R >= 3 && R < 13 && (cellValue === "NAME" || cellValue === "DATE")) {
                                coverSheetWs[address].s.alignment = { horizontal: "left", vertical: "center" };
                                coverSheetWs[address].s.font = { sz: 11 };
                            }
                            if (R >= 15 && R < 20 && C === 0 && cellValue) {
                                coverSheetWs[address].s.font = { bold: true, sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "left", vertical: "center" };
                                coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "F2F2F2" } };
                            }
                            if (R >= 15 && R < 20 && C === 0 && cellValue) {
                                coverSheetWs[address].s.font = { sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "left", vertical: "center" };
                            }
                        }
                    }
                    const lastRow = 48;
                    const lastCol = 10;
                    const darkGrayBorder = {
                        top: { style: "medium", color: { rgb: "404040" } },
                        bottom: { style: "medium", color: { rgb: "404040" } },
                        left: { style: "medium", color: { rgb: "404040" } },
                        right: { style: "medium", color: { rgb: "404040" } }
                    };
                    for (let C = 0; C <= lastCol; C++) {
                        const address = XLSX.utils.encode_cell({ r: 0, c: C });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { top: darkGrayBorder.top });
                    }
                    for (let C = 0; C <= lastCol; C++) {
                        const address = XLSX.utils.encode_cell({ r: lastRow, c: C });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { bottom: darkGrayBorder.bottom });
                    }
                    for (let R = 0; R <= lastRow; R++) {
                        const address = XLSX.utils.encode_cell({ r: R, c: 0 });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { left: darkGrayBorder.left });
                    }
                    for (let R = 0; R <= lastRow; R++) {
                        const address = XLSX.utils.encode_cell({ r: R, c: lastCol });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { right: darkGrayBorder.right });
                    }
                    const corners = [
                        { r: 0, c: 0 }, // Top-left
                        { r: 0, c: lastCol }, // Top-right
                        { r: lastRow, c: 0 }, // Bottom-left
                        { r: lastRow, c: lastCol } // Bottom-right
                    ];
                    corners.forEach(({ r, c }) => {
                        const address = XLSX.utils.encode_cell({ r, c });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        coverSheetWs[address].s.border = darkGrayBorder;
                    });
                    for (let R = 4; R < 13; R += 2) {
                        const dateAddressC = XLSX.utils.encode_cell({ r: R, c: 2 });
                        if (coverSheetWs[dateAddressC]) {
                            if (!coverSheetWs[dateAddressC].s)
                                coverSheetWs[dateAddressC].s = {};
                            coverSheetWs[dateAddressC].s.border = borderStyle;
                        }
                        const dateAddressF = XLSX.utils.encode_cell({ r: R, c: 5 });
                        if (coverSheetWs[dateAddressF]) {
                            if (!coverSheetWs[dateAddressF].s)
                                coverSheetWs[dateAddressF].s = {};
                            coverSheetWs[dateAddressF].s.border = borderStyle;
                        }
                        const dateAddressI = XLSX.utils.encode_cell({ r: R, c: 8 });
                        if (coverSheetWs[dateAddressI]) {
                            if (!coverSheetWs[dateAddressI].s)
                                coverSheetWs[dateAddressI].s = {};
                            coverSheetWs[dateAddressI].s.border = borderStyle;
                        }
                    }
                    // Apply borders to all cells in the merged range - SHIFTED by +3 rows
                    const mergeRangesToFix = [
                        { s: { r: 3, c: 2 }, e: { r: 3, c: 3 } },
                        { s: { r: 5, c: 2 }, e: { r: 5, c: 3 } },
                        { s: { r: 7, c: 2 }, e: { r: 7, c: 3 } },
                        { s: { r: 9, c: 2 }, e: { r: 9, c: 3 } },
                        { s: { r: 11, c: 2 }, e: { r: 11, c: 3 } },
                        { s: { r: 3, c: 5 }, e: { r: 3, c: 6 } },
                        { s: { r: 5, c: 5 }, e: { r: 5, c: 6 } },
                        { s: { r: 7, c: 5 }, e: { r: 7, c: 6 } },
                        { s: { r: 9, c: 5 }, e: { r: 9, c: 6 } },
                        { s: { r: 11, c: 5 }, e: { r: 11, c: 6 } },
                        { s: { r: 3, c: 8 }, e: { r: 3, c: 9 } },
                        { s: { r: 5, c: 8 }, e: { r: 5, c: 9 } },
                        { s: { r: 7, c: 8 }, e: { r: 7, c: 9 } },
                        { s: { r: 9, c: 8 }, e: { r: 9, c: 9 } },
                        { s: { r: 11, c: 8 }, e: { r: 11, c: 9 } },
                    ];
                    mergeRangesToFix.forEach(range => {
                        for (let r = range.s.r; r <= range.e.r; r++) {
                            for (let c = range.s.c; c <= range.e.c; c++) {
                                const address = XLSX.utils.encode_cell({ r: r, c: c });
                                if (!coverSheetWs[address]) {
                                    coverSheetWs[address] = { t: 's', v: '' };
                                }
                                if (!coverSheetWs[address].s)
                                    coverSheetWs[address].s = {};
                                coverSheetWs[address].s.border = borderStyle;
                            }
                        }
                    });
                    // ========== COLUMN WIDTHS ==========
                    coverSheetWs["!cols"] = [
                        { wch: 5 },
                        { wch: 11 },
                        { wch: 6 },
                        { wch: 6 },
                        { wch: 6 },
                        { wch: 6 },
                        { wch: 6 },
                        { wch: 6 },
                        { wch: 6 },
                        { wch: 6 },
                        { wch: 6 }
                    ];
                    // ========== MERGE CELLS ==========
                    const merges = [
                        { s: { r: 0, c: 0 }, e: { r: 2, c: 10 } },
                        { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } },
                        { s: { r: 5, c: 0 }, e: { r: 6, c: 0 } },
                        { s: { r: 7, c: 0 }, e: { r: 8, c: 0 } },
                        { s: { r: 9, c: 0 }, e: { r: 10, c: 0 } },
                        { s: { r: 11, c: 0 }, e: { r: 12, c: 0 } },
                        { s: { r: 3, c: 1 }, e: { r: 4, c: 1 } },
                        { s: { r: 5, c: 1 }, e: { r: 6, c: 1 } },
                        { s: { r: 7, c: 1 }, e: { r: 8, c: 1 } },
                        { s: { r: 9, c: 1 }, e: { r: 10, c: 1 } },
                        { s: { r: 11, c: 1 }, e: { r: 12, c: 1 } },
                        { s: { r: 3, c: 4 }, e: { r: 4, c: 4 } },
                        { s: { r: 5, c: 4 }, e: { r: 6, c: 4 } },
                        { s: { r: 7, c: 4 }, e: { r: 8, c: 4 } },
                        { s: { r: 9, c: 4 }, e: { r: 10, c: 4 } },
                        { s: { r: 11, c: 4 }, e: { r: 12, c: 4 } },
                        { s: { r: 3, c: 7 }, e: { r: 4, c: 7 } },
                        { s: { r: 5, c: 7 }, e: { r: 6, c: 7 } },
                        { s: { r: 7, c: 7 }, e: { r: 8, c: 7 } },
                        { s: { r: 9, c: 7 }, e: { r: 10, c: 7 } },
                        { s: { r: 11, c: 7 }, e: { r: 12, c: 7 } },
                        { s: { r: 3, c: 10 }, e: { r: 4, c: 10 } },
                        { s: { r: 5, c: 10 }, e: { r: 6, c: 10 } },
                        { s: { r: 7, c: 10 }, e: { r: 8, c: 10 } },
                        { s: { r: 9, c: 10 }, e: { r: 10, c: 10 } },
                        { s: { r: 11, c: 10 }, e: { r: 12, c: 10 } },
                        { s: { r: 13, c: 2 }, e: { r: 13, c: 3 } },
                        { s: { r: 13, c: 5 }, e: { r: 13, c: 6 } },
                        { s: { r: 13, c: 8 }, e: { r: 13, c: 9 } },
                        { s: { r: 14, c: 0 }, e: { r: 14, c: 10 } },
                    ];
                    for (let R = 12; R < 13; R++) {
                        if (R % 2 === 1) {
                            merges.push({ s: { r: R, c: 2 }, e: { r: R, c: 3 } });
                            merges.push({ s: { r: R, c: 5 }, e: { r: R, c: 6 } });
                            merges.push({ s: { r: R, c: 8 }, e: { r: R, c: 9 } });
                        }
                    }
                    // ========== PROJECT INFO MERGES ==========
                    const projectInfoStartRow = 15;
                    for (let i = 0; i < 5; i++) {
                        merges.push({ s: { r: projectInfoStartRow + i, c: 0 }, e: { r: projectInfoStartRow + i, c: 1 } });
                        merges.push({ s: { r: projectInfoStartRow + i, c: 2 }, e: { r: projectInfoStartRow + i, c: 10 } });
                    }
                    for (let i = 0; i < 2; i++) {
                        merges.push({ s: { r: projectInfoStartRow + 5 + i, c: 0 }, e: { r: projectInfoStartRow + 5 + i, c: 10 } });
                    }
                    // ========== PMC / CUSTOMER / END USER / SUPPLIER MERGES ==========
                    const pmcStartRow = projectInfoStartRow + 5 + 2;
                    const customerStartRow = pmcStartRow + 6 + 1;
                    const endUserStartRow = customerStartRow + 6 + 1;
                    const supplierStartRow = endUserStartRow + 6 + 1;
                    const sectionStartRows = [pmcStartRow, customerStartRow, endUserStartRow, supplierStartRow];
                    sectionStartRows.forEach(startRow => {
                        for (let row = 0; row < 6; row++) {
                            merges.push({ s: { r: startRow + row, c: 0 }, e: { r: startRow + row, c: 10 } });
                        }
                        if (startRow !== supplierStartRow) {
                            merges.push({ s: { r: startRow + 6, c: 0 }, e: { r: startRow + 6, c: 10 } });
                        }
                    });
                    coverSheetWs["!merges"] = merges;
                    const sections = [
                        { name: "PMC", startRow: pmcStartRow, endRow: pmcStartRow + 5 },
                        { name: "CUSTOMER", startRow: customerStartRow, endRow: customerStartRow + 5 },
                        { name: "END USER", startRow: endUserStartRow, endRow: endUserStartRow + 5 },
                        { name: "SUPPLIER", startRow: supplierStartRow, endRow: supplierStartRow + 5 }
                    ];
                    const separatorRows = [
                        pmcStartRow - 1,
                        customerStartRow - 1,
                        endUserStartRow - 1,
                        supplierStartRow - 1
                    ];
                    separatorRows.forEach(row => {
                        for (let C = 0; C <= 10; ++C) {
                            const address = XLSX.utils.encode_cell({ r: row, c: C });
                            if (!coverSheetWs[address]) {
                                coverSheetWs[address] = { t: 's', v: '' };
                            }
                            if (!coverSheetWs[address].s)
                                coverSheetWs[address].s = {};
                            coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "D9D9D9" } };
                        }
                    });
                    sections.forEach(section => {
                        for (let R = section.startRow; R <= section.endRow; ++R) {
                            for (let C = 0; C <= 10; ++C) {
                                const address = XLSX.utils.encode_cell({ r: R, c: C });
                                if (!coverSheetWs[address]) {
                                    coverSheetWs[address] = { t: 's', v: '' };
                                }
                                if (!coverSheetWs[address].s)
                                    coverSheetWs[address].s = {};
                                const isTopEdge = (R === section.startRow);
                                const isBottomEdge = (R === section.endRow);
                                const isLeftEdge = (C === 0);
                                const isRightEdge = (C === 10);
                                const borderStyle = {
                                    top: { style: "none" },
                                    bottom: { style: "none" },
                                    left: { style: "none" },
                                    right: { style: "none" }
                                };
                                if (isTopEdge)
                                    borderStyle.top = { style: "thin" };
                                if (isBottomEdge)
                                    borderStyle.bottom = { style: "thin" };
                                if (isLeftEdge)
                                    borderStyle.left = { style: "thin" };
                                if (isRightEdge)
                                    borderStyle.right = { style: "thin" };
                                coverSheetWs[address].s.border = borderStyle;
                                if (isTopEdge && C === 0) {
                                    const titleCell = coverSheetWs[address];
                                    if (titleCell.v) {
                                        titleCell.s.font = { bold: true, sz: 11 };
                                        titleCell.s.alignment = { horizontal: "left", vertical: "center" };
                                        titleCell.s.fill = { patternType: "solid", fgColor: { rgb: "FFFFFF" } };
                                    }
                                }
                                if (isTopEdge && C === 2) {
                                    const valueCell = coverSheetWs[address];
                                    if (valueCell.v) {
                                        valueCell.s.font = { sz: 11 };
                                        valueCell.s.alignment = { horizontal: "left", vertical: "center" };
                                        valueCell.s.fill = { patternType: "solid", fgColor: { rgb: "FFFFFF" } };
                                    }
                                }
                                coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "FFFFFF" } };
                            }
                        }
                    });
                    // Add Cover Sheet to workbook
                    XLSX.utils.book_append_sheet(wb, coverSheetWs, "Cover Sheet");
                    // Create the DCS Document sheet
                    const dcsWs = await createDCSDocumentSheet(outputs, inputs, project);
                    // Add DCS Document sheet to workbook
                    XLSX.utils.book_append_sheet(wb, dcsWs, "DCS Documents");
                    // Save the file
                    const excelBuffer = XLSX.write(wb, {
                        bookType: "xlsx",
                        type: "array",
                    });
                    const dataBlob = new Blob([excelBuffer], {
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });
                    const timestamp = new Date().toISOString().split("T")[0];
                    const safeProjectName = project.Title.replace(/[^a-zA-Z0-9-_ ]/g, '_').substring(0, 30);
                    const fileName = `DCS_${safeProjectName}_${timestamp}.xlsx`;
                    const link = document.createElement('a');
                    const url = window.URL.createObjectURL(dataBlob);
                    link.href = url;
                    link.download = fileName;
                    document.body.appendChild(link);
                    link.click();
                    setTimeout(() => {
                        document.body.removeChild(link);
                        window.URL.revokeObjectURL(url);
                    }, 100);
                    hideExportProgress();
                    toast(`Exported ${outputs.length + inputs.length} documents to Excel (2 sheets)`, "success");
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Export error:", error);
                    toast("Failed to export Excel report", "error");
                }
            };
            // Helper function to create DCS Document sheet (extracted from previous code)
            const createDCSDocumentSheet = async (outputs, inputs, project) => {
                // =================== CORRECTED HEADER STRUCTURE ===================
                const headerRow1 = [
                    "Sr No", "Document No", "Description", "Category", "Engg Division", "Rev No",
                    "Document Schedule", "Document Schedule",
                    "Document Submission", "Document Submission",
                    "Document Approval", "Document Approval",
                    "Status", "Remarks"
                ];
                const headerRow2 = [
                    "", "", "", "", "", "",
                    "Original Plan", "Revised Plan",
                    "Actual Submission", "Transmittal No",
                    "Date", "Transmittal No",
                    "", ""
                ];
                // Create worksheet with two-level headers
                const ws = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2]);
                // Prepare data rows
                const dataRows = [];
                const merges = [];
                let currentRow = 2; // Start after headers
                // =================== OUTPUT SECTION ===================
                if (outputs.length > 0) {
                    let outputSrNo = 1;
                    const outputTitleRow = new Array(headerRow1.length).fill("");
                    outputTitleRow[0] = "DESIGN OUTPUT - FROM ABB";
                    dataRows.push(outputTitleRow);
                    merges.push({
                        s: { r: currentRow, c: 0 },
                        e: { r: currentRow, c: headerRow1.length - 1 }
                    });
                    currentRow++;
                    outputs.forEach((doc) => {
                        const docStartRow = currentRow;
                        const mainRow = new Array(headerRow1.length).fill("");
                        mainRow[0] = outputSrNo;
                        mainRow[1] = doc.docNo || "";
                        mainRow[2] = doc.description || "";
                        mainRow[3] = doc.category === "A" ? "A - Approval" : "I - Information";
                        mainRow[4] = doc.enggDivision || "";
                        mainRow[5] = "0";
                        mainRow[6] = formatDateForExcel(doc.originalPlan);
                        mainRow[7] = formatDateForExcel(doc.revisedPlan);
                        mainRow[8] = formatDateForExcel(doc.actualSubmission);
                        mainRow[9] = doc.submissionTransmittal || "";
                        mainRow[10] = formatDateForExcel(doc.approvalDate);
                        mainRow[11] = doc.approvalTransmittal || "";
                        mainRow[12] = doc.status || "Not Submitted";
                        mainRow[13] = doc.remarks || "";
                        dataRows.push(mainRow);
                        currentRow++;
                        outputSrNo++;
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach((revision) => {
                                const revRow = new Array(headerRow1.length).fill("");
                                revRow[5] = revision.revNumber;
                                revRow[6] = formatDateForExcel(revision.originalPlan);
                                revRow[7] = formatDateForExcel(revision.revisedPlan);
                                revRow[8] = formatDateForExcel(revision.actualSubmission);
                                revRow[9] = revision.submissionTransmittal || "";
                                revRow[10] = formatDateForExcel(revision.approvalDate);
                                revRow[11] = revision.approvalTransmittal || "";
                                dataRows.push(revRow);
                                currentRow++;
                            });
                            const mergedColumns = [0, 1, 2, 3, 4, 12, 13];
                            mergedColumns.forEach(col => {
                                merges.push({
                                    s: { r: docStartRow, c: col },
                                    e: { r: currentRow - 1, c: col }
                                });
                            });
                        }
                    });
                }
                // =================== INPUT SECTION ===================
                if (inputs.length > 0) {
                    let inputSrNo = 1;
                    const emptyRow = new Array(headerRow1.length).fill("");
                    dataRows.push(emptyRow);
                    currentRow++;
                    const inputTitleRow = new Array(headerRow1.length).fill("");
                    inputTitleRow[0] = "DESIGN INPUT FROM CONSULTANT / CUSTOMER";
                    dataRows.push(inputTitleRow);
                    merges.push({
                        s: { r: currentRow, c: 0 },
                        e: { r: currentRow, c: headerRow1.length - 1 }
                    });
                    currentRow++;
                    inputs.forEach((doc) => {
                        const docStartRow = currentRow;
                        const mainRow = new Array(headerRow1.length).fill("");
                        mainRow[0] = inputSrNo;
                        mainRow[1] = doc.docNo || "";
                        mainRow[2] = doc.description || "";
                        mainRow[3] = doc.category === "A" ? "A - Approval" : "I - Information";
                        mainRow[4] = doc.enggDivision || "";
                        mainRow[5] = "0";
                        mainRow[6] = formatDateForExcel(doc.originalPlan);
                        mainRow[7] = formatDateForExcel(doc.revisedPlan);
                        mainRow[8] = formatDateForExcel(doc.actualSubmission);
                        mainRow[9] = doc.submissionTransmittal || "";
                        mainRow[10] = formatDateForExcel(doc.approvalDate);
                        mainRow[11] = doc.approvalTransmittal || "";
                        mainRow[12] = doc.status || "Pending";
                        mainRow[13] = doc.remarks || "";
                        dataRows.push(mainRow);
                        currentRow++;
                        inputSrNo++;
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach((revision) => {
                                const revRow = new Array(headerRow1.length).fill("");
                                revRow[5] = revision.revNumber;
                                revRow[6] = formatDateForExcel(revision.originalPlan);
                                revRow[7] = formatDateForExcel(revision.revisedPlan);
                                revRow[8] = formatDateForExcel(revision.actualSubmission);
                                revRow[9] = revision.submissionTransmittal || "";
                                revRow[10] = formatDateForExcel(revision.approvalDate);
                                revRow[11] = revision.approvalTransmittal || "";
                                dataRows.push(revRow);
                                currentRow++;
                            });
                            const mergedColumns = [0, 1, 2, 3, 4, 12, 13];
                            mergedColumns.forEach(col => {
                                merges.push({
                                    s: { r: docStartRow, c: col },
                                    e: { r: currentRow - 1, c: col }
                                });
                            });
                        }
                    });
                }
                // Add data to worksheet
                XLSX.utils.sheet_add_aoa(ws, dataRows, { origin: "A3" });
                // =================== STYLING ===================
                const outputScheduleColor = "E3F2FD";
                const inputScheduleColor = "E8F5E9";
                const submissionColor = "FFF3CD";
                const approvalColor = "D4EDDA";
                const dataRange = XLSX.utils.decode_range(ws["!ref"] || "A1");
                const totalRows = dataRange.e.r;
                const totalCols = dataRange.e.c;
                // =================== HEADER MERGES AND STYLING ===================
                merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, { s: { r: 0, c: 4 }, e: { r: 1, c: 4 } }, { s: { r: 0, c: 5 }, e: { r: 1, c: 5 } }, { s: { r: 0, c: 6 }, e: { r: 0, c: 7 } }, { s: { r: 0, c: 8 }, e: { r: 0, c: 9 } }, { s: { r: 0, c: 10 }, e: { r: 0, c: 11 } }, { s: { r: 0, c: 12 }, e: { r: 1, c: 12 } }, { s: { r: 0, c: 13 }, e: { r: 1, c: 13 } });
                for (let C = 0; C < headerRow1.length; C++) {
                    const addr1 = XLSX.utils.encode_cell({ r: 0, c: C });
                    if (ws[addr1]) {
                        ws[addr1].s = {
                            font: { bold: true, sz: 11 },
                            alignment: { horizontal: "center", vertical: "center", wrapText: true },
                            fill: { patternType: "solid", fgColor: { rgb: "C9E2F3" } },
                            border: {
                                top: { style: "medium", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } },
                                left: { style: C === 0 ? "medium" : "thin", color: { rgb: "000000" } },
                                right: { style: C === totalCols ? "medium" : "thin", color: { rgb: "000000" } }
                            }
                        };
                    }
                    const addr2 = XLSX.utils.encode_cell({ r: 1, c: C });
                    if (ws[addr2]) {
                        ws[addr2].s = {
                            font: { bold: true, sz: 10 },
                            alignment: { horizontal: "center", vertical: "center" },
                            fill: { patternType: "solid", fgColor: { rgb: "D9E1F2" } },
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "medium", color: { rgb: "000000" } },
                                left: { style: C === 0 ? "medium" : "thin", color: { rgb: "000000" } },
                                right: { style: C === totalCols ? "medium" : "thin", color: { rgb: "000000" } }
                            }
                        };
                    }
                }
                ws["!merges"] = merges;
                // =================== DATA CELLS STYLING ===================
                for (let R = 2; R <= totalRows; R++) {
                    for (let C = 0; C <= totalCols; C++) {
                        const address = XLSX.utils.encode_cell({ r: R, c: C });
                        if (!ws[address])
                            continue;
                        if (!ws[address].s)
                            ws[address].s = {};
                        const cellValue = ws[address].v;
                        const cellValueStr = typeof cellValue === 'string' ? cellValue : String(cellValue || '');
                        if (C === 0 && (cellValueStr === "DESIGN OUTPUT - FROM ABB" ||
                            cellValueStr === "DESIGN INPUT FROM CONSULTANT / CUSTOMER")) {
                            for (let titleC = 0; titleC <= totalCols; titleC++) {
                                const titleAddr = XLSX.utils.encode_cell({ r: R, c: titleC });
                                if (ws[titleAddr]) {
                                    ws[titleAddr].s = {
                                        font: { bold: true, sz: 12, color: { rgb: "000000" } },
                                        alignment: { horizontal: "left", vertical: "center" },
                                        fill: {
                                            patternType: "solid",
                                            fgColor: {
                                                rgb: cellValueStr.includes("OUTPUT") ? outputScheduleColor : inputScheduleColor
                                            }
                                        },
                                        border: {
                                            top: { style: "medium", color: { rgb: "000000" } },
                                            bottom: { style: "thin", color: { rgb: "000000" } },
                                            left: { style: titleC === 0 ? "medium" : "thin", color: { rgb: "000000" } },
                                            right: { style: titleC === totalCols ? "medium" : "thin", color: { rgb: "000000" } }
                                        }
                                    };
                                }
                            }
                            continue;
                        }
                        if (R >= 3) {
                            let sectionType = null;
                            for (let checkR = R; checkR >= 2; checkR--) {
                                const checkAddr = XLSX.utils.encode_cell({ r: checkR, c: 0 });
                                if (ws[checkAddr]) {
                                    const titleValue = ws[checkAddr].v;
                                    const titleValueStr = typeof titleValue === 'string' ? titleValue : String(titleValue || '');
                                    if (titleValueStr === "DESIGN OUTPUT - FROM ABB") {
                                        sectionType = "OUTPUT";
                                        break;
                                    }
                                    else if (titleValueStr === "DESIGN INPUT FROM CONSULTANT / CUSTOMER") {
                                        sectionType = "INPUT";
                                        break;
                                    }
                                }
                            }
                            if (sectionType) {
                                if (C === 6 || C === 7) {
                                    const scheduleColor = sectionType === "OUTPUT" ? outputScheduleColor : inputScheduleColor;
                                    ws[address].s.fill = { patternType: "solid", fgColor: { rgb: scheduleColor } };
                                }
                                else if (C === 8 || C === 9) {
                                    ws[address].s.fill = { patternType: "solid", fgColor: { rgb: submissionColor } };
                                }
                                else if (C === 10 || C === 11) {
                                    ws[address].s.fill = { patternType: "solid", fgColor: { rgb: approvalColor } };
                                }
                            }
                            if (C === 0 || C === 5 || C === 12) {
                                ws[address].s.alignment = { horizontal: "center", vertical: "center" };
                            }
                            else {
                                ws[address].s.alignment = {
                                    horizontal: (C === 3 || C === 4) ? "center" : "left",
                                    vertical: "center"
                                };
                            }
                            const prevRowAddr = XLSX.utils.encode_cell({ r: R - 1, c: 0 });
                            const nextRowAddr = XLSX.utils.encode_cell({ r: R + 1, c: 0 });
                            const prevCellValue = ws[prevRowAddr] ? ws[prevRowAddr].v : null;
                            const nextCellValue = ws[nextRowAddr] ? ws[nextRowAddr].v : null;
                            const prevCellValueStr = typeof prevCellValue === 'string' ? prevCellValue : String(prevCellValue || '');
                            const nextCellValueStr = typeof nextCellValue === 'string' ? nextCellValue : String(nextCellValue || '');
                            const isTopRow = R === 3 || (R > 3 && (prevCellValueStr.includes("DESIGN OUTPUT - FROM ABB") ||
                                prevCellValueStr.includes("DESIGN INPUT FROM CONSULTANT / CUSTOMER") ||
                                prevCellValueStr === ""));
                            const isBottomRow = R === totalRows || (R < totalRows && (nextCellValueStr.includes("DESIGN OUTPUT - FROM ABB") ||
                                nextCellValueStr.includes("DESIGN INPUT FROM CONSULTANT / CUSTOMER") ||
                                nextCellValueStr === ""));
                            ws[address].s.border = {
                                top: { style: isTopRow ? "medium" : "thin", color: { rgb: "000000" } },
                                bottom: { style: isBottomRow ? "medium" : "thin", color: { rgb: "000000" } },
                                left: { style: C === 0 ? "medium" : "thin", color: { rgb: "000000" } },
                                right: { style: C === totalCols ? "medium" : "thin", color: { rgb: "000000" } }
                            };
                        }
                    }
                }
                ws["!cols"] = [
                    { wch: 8 }, // A - Sr No
                    { wch: 20 }, // B - Document No
                    { wch: 35 }, // C - Description
                    { wch: 12 }, // D - Category
                    { wch: 18 }, // E - Engg Division
                    { wch: 10 }, // F - Rev No
                    { wch: 15 }, // G - Original Plan
                    { wch: 15 }, // H - Revised Plan
                    { wch: 18 }, // I - Actual Submission
                    { wch: 20 }, // J - Transmittal No
                    { wch: 15 }, // K - Approval Date
                    { wch: 20 }, // L - Transmittal No
                    { wch: 20 }, // M - Status
                    { wch: 25 } // N - Remarks
                ];
                return ws;
            };
            // Format date for Excel export (YYYY-MM-DD)
            const formatDateForExcel = (dateString) => {
                if (!dateString)
                    return "";
                try {
                    const date = new Date(dateString);
                    if (isNaN(date.getTime()))
                        return "";
                    const year = date.getFullYear();
                    const month = String(date.getMonth() + 1).padStart(2, '0');
                    const day = String(date.getDate()).padStart(2, '0');
                    return `${year}-${month}-${day}`;
                }
                catch (_a) {
                    return "";
                }
            };
            // Show export progress
            const showExportProgress = (message) => {
                const existing = document.querySelector('.dcs-export-progress');
                if (existing)
                    existing.remove();
                const progressEl = document.createElement("div");
                progressEl.className = "dcs-export-progress";
                progressEl.style.cssText = `
          position: fixed;
          top: 0;
          left: 0;
          width: 100%;
          height: 100%;
          background: rgba(0,0,0,0.5);
          display: flex;
          justify-content: center;
          align-items: center;
          z-index: 10001;
        `;
                progressEl.innerHTML = `
          <div style="background: white; padding: 30px; border-radius: 8px; text-align: center; min-width: 300px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <div style="font-size: 36px; margin-bottom: 15px;">⏳</div>
            <div style="font-size: 15px; font-weight: 600; color: #212529; margin-bottom: 10px;">${message}</div>
            <div style="font-size: 12px; color: #6c757d;">Please wait...</div>
          </div>
        `;
                document.body.appendChild(progressEl);
            };
            // Hide export progress
            const hideExportProgress = () => {
                const existing = document.querySelector('.dcs-export-progress');
                if (existing)
                    existing.remove();
            };
            // Setup export functionality
            const setupImportExportFunctionality = () => {
                if (exportExcelBtn) {
                    exportExcelBtn.removeEventListener("click", exportToExcel);
                    exportExcelBtn.addEventListener("click", exportToExcel);
                }
                const importInput = document.getElementById("dcsImportExcelInput");
                const importBtn = document.querySelector(".dcs-import-excel-btn");
                const exportTemplateBtn = document.querySelector(".dcs-export-project-excel-btn");
                const downloadExistingBtn = document.querySelector(".dcs-download-existing-btn");
                if (importBtn && importInput) {
                    importBtn.removeEventListener("click", () => importInput.click());
                    importBtn.addEventListener("click", () => importInput.click());
                    importInput.removeEventListener("change", handleExcelImport);
                    importInput.addEventListener("change", handleExcelImport);
                }
                if (exportTemplateBtn) {
                    exportTemplateBtn.removeEventListener("click", exportTemplateExcel);
                    exportTemplateBtn.addEventListener("click", exportTemplateExcel);
                }
                if (downloadExistingBtn) {
                    downloadExistingBtn.removeEventListener("click", downloadExistingData);
                    downloadExistingBtn.addEventListener("click", downloadExistingData);
                }
            };
            const exportToExcel = async () => {
                showExportProgress("Preparing DCS data for export...");
                try {
                    const filterProjectValue = filterProject.value;
                    const filterCustomerValue = filterCustomer.value;
                    const filterPMValue = filterProjectManager.value;
                    const filterRefValue = filterProjectReference.value;
                    // Apply current filters
                    const filteredProjects = projects.filter((p) => {
                        var _a;
                        if (filterProjectValue && p.Title !== filterProjectValue)
                            return false;
                        if (filterCustomerValue && p.CustomerName !== filterCustomerValue)
                            return false;
                        if (filterRefValue && p.ProjectReference !== filterRefValue)
                            return false;
                        if (filterPMValue) {
                            const managerName = typeof p.ProjectManager === "object"
                                ? (_a = p.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title
                                : p.ProjectManager;
                            if (managerName !== filterPMValue)
                                return false;
                        }
                        return true;
                    });
                    if (filteredProjects.length === 0) {
                        hideExportProgress();
                        alert("No data to export with current filters");
                        return;
                    }
                    showExportProgress(`Exporting ${filteredProjects.length} projects...`);
                    const wb = XLSX.utils.book_new();
                    // Build comprehensive export data
                    const exportData = [];
                    // Header row
                    exportData.push([
                        "Sr No", "Project Reference", "Customer", "Project", "Project Manager",
                        "Total Documents", "Not Submitted", "Not Approved",
                        "Revised Doc not Submitted/Approved with Comments", "Approved", "Rejected", "Retained for Information"
                    ]);
                    let srNo = 1;
                    filteredProjects.forEach((project) => {
                        const record = dcsData.find((item) => {
                            const itemProjectId = item[`${fieldMappings.Project}Id`] ||
                                (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
                            return itemProjectId === project.Id;
                        });
                        if (!record || !record[fieldMappings.DCSData]) {
                            exportData.push([
                                srNo++,
                                project.ProjectReference || "",
                                project.CustomerName || "",
                                project.Title || "",
                                getProjectManagerDisplayName(project.ProjectManager),
                                0, 0, 0, 0, 0, 0, 0
                            ]);
                            return;
                        }
                        const summary = getDocumentStatusSummary(record);
                        exportData.push([
                            srNo++,
                            project.ProjectReference || "",
                            project.CustomerName || "",
                            project.Title || "",
                            getProjectManagerDisplayName(project.ProjectManager),
                            summary.total,
                            summary.notSubmitted,
                            summary.notApproved,
                            summary.revisedOrApprovedWithComments,
                            summary.approved,
                            summary.rejected,
                            summary.retainedForInformation
                        ]);
                    });
                    // Create worksheet
                    const ws = XLSX.utils.aoa_to_sheet(exportData);
                    // Style the header row
                    const headerStyle = {
                        font: { bold: true, color: { rgb: "FFFFFF" } },
                        fill: { patternType: "solid", fgColor: { rgb: "0078D4" } },
                        alignment: { horizontal: "center", vertical: "center", wrapText: true },
                        border: {
                            top: { style: "thin" },
                            bottom: { style: "thin" },
                            left: { style: "thin" },
                            right: { style: "thin" }
                        }
                    };
                    const headerRange = XLSX.utils.decode_range("A1:L1");
                    for (let C = headerRange.s.c; C <= headerRange.e.c; C++) {
                        const address = XLSX.utils.encode_cell({ r: 0, c: C });
                        if (ws[address]) {
                            ws[address].s = headerStyle;
                        }
                    }
                    // Style data rows - highlight non-zero counts
                    for (let R = 1; R <= exportData.length - 1; R++) {
                        for (let C = 0; C <= 11; C++) {
                            const address = XLSX.utils.encode_cell({ r: R, c: C });
                            if (ws[address]) {
                                if (!ws[address].s)
                                    ws[address].s = {};
                                ws[address].s.border = {
                                    top: { style: "thin" },
                                    bottom: { style: "thin" },
                                    left: { style: "thin" },
                                    right: { style: "thin" }
                                };
                                ws[address].s.alignment = { vertical: "center" };
                                // Highlight count columns with colors
                                if (C >= 5) {
                                    const value = Number(ws[address].v) || 0;
                                    if (value > 0) {
                                        let color = "FFFFFF";
                                        if (C === 6)
                                            color = "FFE5E5"; // Not Submitted - light red
                                        else if (C === 7)
                                            color = "FFF3CD"; // Not Approved - light yellow
                                        else if (C === 8)
                                            color = "E2F0D9"; // Revised - light green
                                        else if (C === 9)
                                            color = "D4EDDA"; // Approved - green
                                        else if (C === 10)
                                            color = "F8D7DA"; // Rejected - light red
                                        else if (C === 11)
                                            color = "D1ECF1"; // Retained - light blue
                                        ws[address].s.fill = { patternType: "solid", fgColor: { rgb: color } };
                                    }
                                    ws[address].s.alignment = { horizontal: "center", vertical: "center" };
                                    ws[address].s.font = { bold: value > 0 };
                                }
                            }
                        }
                    }
                    // Set column widths
                    ws["!cols"] = [
                        { wch: 6 }, // Sr No
                        { wch: 20 }, // Project Reference
                        { wch: 25 }, // Customer
                        { wch: 35 }, // Project
                        { wch: 25 }, // Project Manager
                        { wch: 10 }, // Total Documents
                        { wch: 12 }, // Not Submitted
                        { wch: 12 }, // Not Approved
                        { wch: 16 }, // Revised
                        { wch: 10 }, // Approved
                        { wch: 10 }, // Rejected
                        { wch: 14 } // Retained
                    ];
                    XLSX.utils.book_append_sheet(wb, ws, "DCS Summary");
                    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
                    const dataBlob = new Blob([excelBuffer], {
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    });
                    const timestamp = new Date().toISOString().split("T")[0];
                    const now = new Date();
                    const timeStr = `${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
                    const fileName = `DCS_Export_${timestamp}_${timeStr}.xlsx`;
                    const link = document.createElement("a");
                    const url = window.URL.createObjectURL(dataBlob);
                    link.href = url;
                    link.download = fileName;
                    document.body.appendChild(link);
                    link.click();
                    setTimeout(() => {
                        document.body.removeChild(link);
                        window.URL.revokeObjectURL(url);
                    }, 100);
                    hideExportProgress();
                    console.log(`✅ Export complete: ${filteredProjects.length} projects`);
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Export error:", error);
                    alert("Failed to export data. Please try again.");
                }
            };
            // Export template Excel (empty template with structure)
            const exportTemplateExcel = async () => {
                showExportProgress("Generating Excel template...");
                try {
                    const wb = XLSX.utils.book_new();
                    // Template headers with two-level structure
                    const headerRow1 = [
                        "Sr No", "Document No", "Description", "Category", "Engg Division", "Rev No",
                        "Document Schedule", "Document Schedule",
                        "Document Submission", "Document Submission",
                        "Document Approval", "Document Approval",
                        "Status", "Remarks"
                    ];
                    const headerRow2 = [
                        "", "", "", "", "", "",
                        "Original Plan", "Revised Plan",
                        "Actual Submission", "Transmittal No",
                        "Date", "Transmittal No",
                        "", ""
                    ];
                    // Create worksheet with headers
                    const ws = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2]);
                    // Add sample row
                    const sampleRow = [
                        1, "DOC-001", "Sample document description", "A - Approval", "Automation", 0,
                        "", "", "", "", "", "", "Submitted", "Sample remarks"
                    ];
                    XLSX.utils.sheet_add_aoa(ws, [sampleRow], { origin: "A3" });
                    // Add DESIGN OUTPUT section title
                    const sectionRow = ["DESIGN OUTPUT - FROM ABB", "", "", "", "", "", "", "", "", "", "", "", "", ""];
                    XLSX.utils.sheet_add_aoa(ws, [sectionRow], { origin: "A3" });
                    // Style headers
                    for (let C = 0; C < headerRow1.length; C++) {
                        const addr1 = XLSX.utils.encode_cell({ r: 0, c: C });
                        if (ws[addr1]) {
                            ws[addr1].s = {
                                font: { bold: true, color: { rgb: "FFFFFF" } },
                                fill: { patternType: "solid", fgColor: { rgb: "4472C4" } },
                                alignment: { horizontal: "center", vertical: "center", wrapText: true },
                                border: {
                                    top: { style: "thin" }, bottom: { style: "thin" },
                                    left: { style: "thin" }, right: { style: "thin" }
                                }
                            };
                        }
                        const addr2 = XLSX.utils.encode_cell({ r: 1, c: C });
                        if (ws[addr2]) {
                            ws[addr2].s = {
                                font: { bold: true, color: { rgb: "FFFFFF" } },
                                fill: { patternType: "solid", fgColor: { rgb: "5B9BD5" } },
                                alignment: { horizontal: "center", vertical: "center" },
                                border: {
                                    top: { style: "thin" }, bottom: { style: "thin" },
                                    left: { style: "thin" }, right: { style: "thin" }
                                }
                            };
                        }
                    }
                    ws["!merges"] = [
                        { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
                        { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
                        { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
                        { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } },
                        { s: { r: 0, c: 4 }, e: { r: 1, c: 4 } },
                        { s: { r: 0, c: 5 }, e: { r: 1, c: 5 } },
                        { s: { r: 0, c: 6 }, e: { r: 0, c: 7 } },
                        { s: { r: 0, c: 8 }, e: { r: 0, c: 9 } },
                        { s: { r: 0, c: 10 }, e: { r: 0, c: 11 } },
                        { s: { r: 0, c: 12 }, e: { r: 1, c: 12 } },
                        { s: { r: 0, c: 13 }, e: { r: 1, c: 13 } }
                    ];
                    ws["!cols"] = [
                        { wch: 8 }, { wch: 20 }, { wch: 35 }, { wch: 12 }, { wch: 18 }, { wch: 10 },
                        { wch: 15 }, { wch: 15 }, { wch: 18 }, { wch: 20 }, { wch: 15 }, { wch: 20 },
                        { wch: 20 }, { wch: 25 }
                    ];
                    XLSX.utils.book_append_sheet(wb, ws, "DCS Template");
                    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
                    const dataBlob = new Blob([excelBuffer], {
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    });
                    const link = document.createElement("a");
                    const url = window.URL.createObjectURL(dataBlob);
                    link.href = url;
                    link.download = `DCS_Template_${new Date().toISOString().split("T")[0]}.xlsx`;
                    document.body.appendChild(link);
                    link.click();
                    setTimeout(() => {
                        document.body.removeChild(link);
                        window.URL.revokeObjectURL(url);
                    }, 100);
                    hideExportProgress();
                    toast("Template downloaded successfully", "success");
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Template export error:", error);
                    toast("Failed to generate template", "error");
                }
            };
            // Download existing data as Excel
            const downloadExistingData = async () => {
                const projectId = parseInt(projectSelect.value);
                if (!projectId) {
                    toast("Please select a project first", "warning");
                    return;
                }
                showExportProgress("Preparing existing data for download...");
                try {
                    const project = projects.find(p => p.Id === projectId);
                    if (!project) {
                        hideExportProgress();
                        toast("Project not found", "error");
                        return;
                    }
                    const record = dcsData.find((item) => {
                        const projId = item[`${fieldMappings.Project}Id`] ||
                            (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
                        return projId === projectId;
                    });
                    if (!record || !record[fieldMappings.DCSData]) {
                        hideExportProgress();
                        toast("No existing data for this project", "warning");
                        return;
                    }
                    const parsedData = parseDCSData(record[fieldMappings.DCSData]);
                    if (!parsedData) {
                        hideExportProgress();
                        toast("No document data found", "warning");
                        return;
                    }
                    const outputs = getOutputsFromProjectData(parsedData) || [];
                    const inputs = getInputsFromProjectData(parsedData) || [];
                    if (outputs.length === 0 && inputs.length === 0) {
                        hideExportProgress();
                        toast("No documents to download", "warning");
                        return;
                    }
                    const wb = XLSX.utils.book_new();
                    const ws = await createDCSDocumentSheet(outputs, inputs, project);
                    XLSX.utils.book_append_sheet(wb, ws, "DCS Documents");
                    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
                    const dataBlob = new Blob([excelBuffer], {
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    });
                    const safeProjectName = project.Title.replace(/[^a-zA-Z0-9-_ ]/g, '_').substring(0, 30);
                    const link = document.createElement("a");
                    const url = window.URL.createObjectURL(dataBlob);
                    link.href = url;
                    link.download = `DCS_${safeProjectName}_Existing_${new Date().toISOString().split("T")[0]}.xlsx`;
                    document.body.appendChild(link);
                    link.click();
                    setTimeout(() => {
                        document.body.removeChild(link);
                        window.URL.revokeObjectURL(url);
                    }, 100);
                    hideExportProgress();
                    toast(`Downloaded ${outputs.length + inputs.length} documents`, "success");
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Download error:", error);
                    toast("Failed to download data", "error");
                }
            };
            // Handle Excel import
            const handleExcelImport = async (e) => {
                var _a;
                const target = e.target;
                const file = (_a = target.files) === null || _a === void 0 ? void 0 : _a[0];
                if (!file)
                    return;
                // 🔐 Role gate: import writes data - create/edit roles only
                if (!canCreateRecord()) {
                    toast("You do not have permission to import document data", "error");
                    target.value = "";
                    return;
                }
                showExportProgress("Reading Excel file...");
                try {
                    const data = await file.arrayBuffer();
                    const workbook = XLSX.read(data);
                    const sheetName = workbook.SheetNames.includes("DCS Documents")
                        ? "DCS Documents"
                        : workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    if (jsonData.length < 3) {
                        hideExportProgress();
                        toast("Excel file has no data rows", "warning");
                        return;
                    }
                    showExportProgress("Processing imported data...");
                    const parsed = parseImportedExcelData(jsonData);
                    if (!parsed || (parsed.outputs.length === 0 && parsed.inputs.length === 0)) {
                        hideExportProgress();
                        toast("No valid document data found in Excel", "warning");
                        return;
                    }
                    // Clear existing tables
                    outputTableBody.innerHTML = '';
                    inputTableBody.innerHTML = '';
                    // Add imported output documents
                    if (parsed.outputs.length > 0) {
                        parsed.outputs.forEach((doc, index) => {
                            const row = createEditableDocumentRow(doc, index + 1, false);
                            outputTableBody.appendChild(row);
                            if (doc.revisions && doc.revisions.length > 0) {
                                doc.revisions.forEach((revision) => {
                                    const revHTML = buildRevisionRowHTML(doc.docId, revision);
                                    outputTableBody.insertAdjacentHTML('beforeend', revHTML);
                                });
                                updateMergedCells(doc.docId);
                            }
                        });
                    }
                    // Add imported input documents
                    if (parsed.inputs.length > 0) {
                        parsed.inputs.forEach((doc, index) => {
                            const row = createEditableDocumentRow(doc, index + 1, true);
                            inputTableBody.appendChild(row);
                            if (doc.revisions && doc.revisions.length > 0) {
                                doc.revisions.forEach((revision) => {
                                    const revHTML = buildRevisionRowHTML(doc.docId, revision);
                                    inputTableBody.insertAdjacentHTML('beforeend', revHTML);
                                });
                                updateMergedCells(doc.docId);
                            }
                        });
                    }
                    // Ensure at least one row per table
                    if (outputTableBody.children.length === 0) {
                        outputTableBody.innerHTML = createDocumentRowHTML('output', 1, `default_${Date.now()}_out`);
                    }
                    if (inputTableBody.children.length === 0) {
                        inputTableBody.innerHTML = createDocumentRowHTML('input', 1, `default_${Date.now()}_in`);
                    }
                    // Show container
                    containerTable.style.display = "block";
                    existingDataAlert.style.display = "none";
                    // Re-attach listeners
                    setTimeout(() => {
                        attachDocumentEventListeners();
                        updateAllSerialNumbers();
                        updateAllMergedCells();
                    }, 100);
                    hideExportProgress();
                    toast(`Imported ${parsed.outputs.length} output + ${parsed.inputs.length} input documents`, "success");
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Import error:", error);
                    toast("Failed to import Excel file", "error");
                }
                finally {
                    target.value = "";
                }
            };
            // Parse imported Excel data
            const parseImportedExcelData = (jsonData) => {
                try {
                    const outputs = [];
                    const inputs = [];
                    let currentSection = 'output';
                    let currentDoc = null;
                    for (let i = 2; i < jsonData.length; i++) {
                        const row = jsonData[i] || [];
                        const firstCell = String(row[0] || "").trim();
                        const revNoCell = row[5];
                        // Check for section titles
                        if (firstCell.includes("DESIGN OUTPUT")) {
                            currentSection = 'output';
                            currentDoc = null;
                            continue;
                        }
                        if (firstCell.includes("DESIGN INPUT")) {
                            currentSection = 'input';
                            currentDoc = null;
                            continue;
                        }
                        // Skip empty rows
                        if (!firstCell && !row[1] && !String(row[2] || "").trim()) {
                            continue;
                        }
                        const revNo = parseInt(String(revNoCell !== null && revNoCell !== void 0 ? revNoCell : "0")) || 0;
                        // Check if this is a revision row (revNo > 0 and no doc number)
                        const docNoCell = String(row[1] || "").trim();
                        const isRevisionRow = revNo > 0 && !docNoCell && currentDoc !== null;
                        if (isRevisionRow && currentDoc) {
                            const revision = {
                                revNumber: revNo,
                                originalPlan: parseExcelDate(row[6]),
                                revisedPlan: parseExcelDate(row[7]),
                                actualSubmission: parseExcelDate(row[8]),
                                submissionTransmittal: String(row[9] || "").trim(),
                                approvalDate: parseExcelDate(row[10]),
                                approvalTransmittal: String(row[11] || "").trim(),
                                status: currentDoc.status,
                                remarks: currentDoc.remarks
                            };
                            currentDoc.revisions.push(revision);
                            continue;
                        }
                        // This is a main document row
                        const srNo = parseInt(String(firstCell)) || 0;
                        if (srNo === 0 && !docNoCell)
                            continue;
                        const statusValue = String(row[12] || "").trim();
                        const defaultStatus = currentSection === 'input' ? 'Pending' : 'Not Submitted';
                        const validStatuses = currentSection === 'input' ? INPUT_STATUS_OPTIONS : OUTPUT_STATUS_OPTIONS;
                        const doc = {
                            docId: `${currentSection}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                            docNo: docNoCell,
                            description: String(row[2] || "").trim(),
                            category: String(row[3] || "A").trim().charAt(0).toUpperCase() || "A",
                            enggDivision: String(row[4] || "").trim(),
                            revNo: 0,
                            originalPlan: parseExcelDate(row[6]),
                            revisedPlan: parseExcelDate(row[7]),
                            actualSubmission: parseExcelDate(row[8]),
                            submissionTransmittal: String(row[9] || "").trim(),
                            approvalDate: parseExcelDate(row[10]),
                            approvalTransmittal: String(row[11] || "").trim(),
                            status: validStatuses.includes(statusValue) ? statusValue : (statusValue || defaultStatus),
                            remarks: String(row[13] || "").trim(),
                            isInput: currentSection === 'input',
                            revisions: [],
                            docType: currentSection
                        };
                        if (currentSection === 'output') {
                            outputs.push(doc);
                        }
                        else {
                            inputs.push(doc);
                        }
                        currentDoc = doc;
                    }
                    return { outputs, inputs };
                }
                catch (error) {
                    console.error("Error parsing imported data:", error);
                    return null;
                }
            };
            // Parse Excel date
            const parseExcelDate = (value) => {
                if (!value)
                    return "";
                // If it's already a date string
                if (typeof value === 'string') {
                    const trimmed = value.trim();
                    if (!trimmed)
                        return "";
                    // Try ISO format
                    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed))
                        return trimmed;
                    // Try other formats
                    const parsed = new Date(trimmed);
                    if (!isNaN(parsed.getTime())) {
                        const year = parsed.getFullYear();
                        const month = String(parsed.getMonth() + 1).padStart(2, '0');
                        const day = String(parsed.getDate()).padStart(2, '0');
                        return `${year}-${month}-${day}`;
                    }
                    return "";
                }
                // If it's an Excel serial date number
                if (typeof value === 'number') {
                    const excelEpoch = new Date(1899, 11, 30);
                    const date = new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
                    const year = date.getFullYear();
                    const month = String(date.getMonth() + 1).padStart(2, '0');
                    const day = String(date.getDate()).padStart(2, '0');
                    return `${year}-${month}-${day}`;
                }
                return "";
            };
            // ============================================
            // DOCUMENT TABLE FUNCTIONS
            // ============================================
            const createDocumentRowHTML = (type, srNo, docId) => {
                const prefix = type === 'input' ? 'DOC' : 'DOC-OUT';
                const docNo = `${prefix}-${String(srNo).padStart(3, '0')}`;
                const tableType = type;
                return `
          <tr class="dcs-document-group dcs-document-header" data-doc-id="${docId}" data-is-default="true">
            <td class="dcs-sr-no">${srNo}</td>
            <td>
              <textarea class="dcs-form-control dcs-doc-no" rows="2" placeholder="Document No">${docNo}</textarea>
            </td>
            <td>
              <textarea class="dcs-form-control dcs-description-input" rows="3" placeholder="Enter description"></textarea>
            </td>
            <td>
              <select class="dcs-form-control dcs-category-select">
                <option value="A" selected>A - Approval</option>
                <option value="I">I - Information</option>
              </select>
            </td>
            <td>
              <select class="dcs-form-control dcs-engg-division-select">
                ${buildEnggDivisionOptionsHTML("")}
              </select>
            </td>
            <td class="dcs-rev-cell">0</td>
            <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Original Plan"></td>
            <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Revised Plan"></td>
            <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Actual Submission"></td>
            <td>
              <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="Transmittal No"></textarea>
            </td>
            <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Approval Date"></td>
            <td>
              <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="Transmittal No"></textarea>
            </td>
            <td>
              <select class="dcs-form-control dcs-status-select">
                ${buildStatusOptionsHTML(tableType)}
              </select>
            </td>
            <td>
              <textarea class="dcs-form-control dcs-remarks-input" rows="3" placeholder="Enter remarks"></textarea>
            </td>
            <td class="dcs-doc-actions">
              <button type="button" class="dcs-btn-success dcs-add-revision-btn" title="Add Revision">
                <span style="font-size: 16px;">➕</span>
              </button>
              <button type="button" class="dcs-btn-danger dcs-remove-doc" title="Remove Document">
                <span style="font-size: 16px;">🗑️</span>
              </button>
            </td>
          </tr>
        `;
            };
            const attachDocumentEventListeners = () => {
                // Remove existing listeners by cloning approach is too heavy; use delegation on the tables
                if (addInputDocBtn) {
                    addInputDocBtn.removeEventListener('click', handleAddInputDocument);
                    addInputDocBtn.addEventListener('click', handleAddInputDocument);
                }
                if (addOutputDocBtn) {
                    addOutputDocBtn.removeEventListener('click', handleAddOutputDocument);
                    addOutputDocBtn.addEventListener('click', handleAddOutputDocument);
                }
                // Delegate add/remove actions inside the two tbody containers
                [inputTableBody, outputTableBody].forEach(body => {
                    if (!body)
                        return;
                    body.onclick = (e) => {
                        const target = e.target;
                        // Add Revision
                        const addRevBtn = target.closest('.dcs-add-revision-btn');
                        if (addRevBtn) {
                            const row = addRevBtn.closest('tr');
                            if (row)
                                handleAddRevision(row);
                            return;
                        }
                        // Remove Revision
                        const removeRevBtn = target.closest('.dcs-remove-revision-btn');
                        if (removeRevBtn) {
                            const row = removeRevBtn.closest('tr');
                            if (row)
                                handleRemoveRevision(row);
                            return;
                        }
                        // Remove Document
                        const removeDocBtn = target.closest('.dcs-remove-doc');
                        if (removeDocBtn) {
                            const row = removeDocBtn.closest('tr');
                            if (row)
                                handleRemoveDocument(row);
                            return;
                        }
                    };
                });
            };
            const handleAddInputDocument = () => {
                const rows = inputTableBody.querySelectorAll('tr.dcs-document-header');
                const srNo = rows.length + 1;
                const docId = `input_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
                const rowHTML = createDocumentRowHTML('input', srNo, docId);
                inputTableBody.insertAdjacentHTML('beforeend', rowHTML);
                ensureDocumentButtonsVisible();
                updateAllSerialNumbers();
                console.log('✅ Input document added');
            };
            const handleAddOutputDocument = () => {
                const rows = outputTableBody.querySelectorAll('tr.dcs-document-header');
                const srNo = rows.length + 1;
                const docId = `output_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
                const rowHTML = createDocumentRowHTML('output', srNo, docId);
                outputTableBody.insertAdjacentHTML('beforeend', rowHTML);
                ensureDocumentButtonsVisible();
                updateAllSerialNumbers();
                console.log('✅ Output document added');
            };
            const handleAddRevision = (mainRow) => {
                const docId = mainRow.getAttribute('data-doc-id');
                if (!docId)
                    return;
                const table = mainRow.closest('table');
                if (!table)
                    return;
                const tbody = mainRow.parentElement;
                const allRows = Array.from(tbody.querySelectorAll(`tr[data-doc-id="${docId}"]`));
                const mainIndex = allRows.indexOf(mainRow);
                // Count existing revision rows for this document
                const existingRevs = allRows.length - 1;
                const nextRevNo = existingRevs + 1;
                const revision = {
                    revNumber: nextRevNo,
                    originalPlan: "",
                    revisedPlan: "",
                    actualSubmission: "",
                    submissionTransmittal: "",
                    approvalDate: "",
                    approvalTransmittal: ""
                };
                const revHTML = buildRevisionRowHTML(docId, revision);
                // Insert after the last row of this document group
                const lastRow = allRows[allRows.length - 1];
                if (lastRow) {
                    lastRow.insertAdjacentHTML('afterend', revHTML);
                }
                else {
                    mainRow.insertAdjacentHTML('afterend', revHTML);
                }
                mainRow.setAttribute('data-rev-rows', String(allRows.length + 1));
                updateMergedCells(docId);
                console.log(`✅ Revision ${nextRevNo} added to document ${docId} (main index ${mainIndex})`);
            };
            const handleRemoveRevision = (revRow) => {
                const docId = revRow.getAttribute('data-doc-id');
                if (!docId)
                    return;
                const confirmed = confirm('Remove this revision?');
                if (!confirmed)
                    return;
                const tbody = revRow.parentElement;
                revRow.remove();
                const remaining = tbody.querySelectorAll(`tr[data-doc-id="${docId}"]`);
                const mainRow = remaining[0];
                if (mainRow) {
                    mainRow.setAttribute('data-rev-rows', String(remaining.length));
                    updateMergedCells(docId);
                }
                console.log(`🗑️ Revision removed from document ${docId}`);
            };
            const handleRemoveDocument = (mainRow) => {
                const docId = mainRow.getAttribute('data-doc-id');
                if (!docId)
                    return;
                const confirmed = confirm('Remove this document and all its revisions?');
                if (!confirmed)
                    return;
                const tbody = mainRow.parentElement;
                const rows = tbody.querySelectorAll(`tr[data-doc-id="${docId}"]`);
                rows.forEach(r => r.remove());
                ensureDocumentButtonsVisible();
                updateAllSerialNumbers();
                console.log(`🗑️ Document ${docId} removed`);
            };
            // Merge the Sr No / Doc No / Description / Category / Engg Division / Status / Remarks
            // cells of a document down over its revision rows (Actions column stays unmerged)
            const updateMergedCells = (docId) => {
                const allRows = document.querySelectorAll(`tr[data-doc-id="${docId}"]`);
                if (allRows.length <= 1) {
                    // single row - reset any previously merged cells
                    const single = allRows[0];
                    if (single) {
                        Array.from(single.cells).forEach(c => { c.rowSpan = 1; c.style.display = ''; });
                    }
                    return;
                }
                const mainRow = allRows[0];
                const revRows = Array.from(allRows).slice(1);
                const cellsToMerge = [0, 1, 2, 3, 4, 12, 13];
                cellsToMerge.forEach(idx => {
                    const mainCell = mainRow.cells[idx];
                    if (mainCell)
                        mainCell.rowSpan = allRows.length;
                });
                revRows.forEach(revRow => {
                    cellsToMerge.forEach(idx => {
                        const cell = revRow.cells[idx];
                        if (cell)
                            cell.style.display = 'none';
                    });
                });
            };
            const updateAllMergedCells = () => {
                const docIds = new Set();
                document.querySelectorAll('#dcsOutputTableBody tr[data-doc-id], #dcsInputTableBody tr[data-doc-id]').forEach(row => {
                    const id = row.getAttribute('data-doc-id');
                    if (id)
                        docIds.add(id);
                });
                docIds.forEach(id => updateMergedCells(id));
            };
            // Keep add/remove buttons visible on every document row (edit mode)
            const ensureDocumentButtonsVisible = () => {
                const totalOutputDocs = outputTableBody.querySelectorAll('tr.dcs-document-header').length;
                const totalInputDocs = inputTableBody.querySelectorAll('tr.dcs-document-header').length;
                outputTableBody.querySelectorAll('tr.dcs-document-header').forEach((row, index) => {
                    const removeBtn = row.querySelector('.dcs-remove-doc');
                    if (removeBtn) {
                        // keep the FIRST document when there is only one (cannot delete the only row)
                        removeBtn.style.display = (totalOutputDocs > 1 || index > 0) ? 'inline-flex' : 'none';
                    }
                });
                inputTableBody.querySelectorAll('tr.dcs-document-header').forEach((row, index) => {
                    const removeBtn = row.querySelector('.dcs-remove-doc');
                    if (removeBtn) {
                        removeBtn.style.display = (totalInputDocs > 1 || index > 0) ? 'inline-flex' : 'none';
                    }
                });
            };
            // Renumber Sr No of every document group (1,2,3...) in both tables
            const updateAllSerialNumbers = () => {
                [outputTableBody, inputTableBody].forEach(body => {
                    const headers = body.querySelectorAll('tr.dcs-document-header');
                    headers.forEach((row, index) => {
                        const srCell = row.querySelector('.dcs-sr-no');
                        if (srCell)
                            srCell.textContent = String(index + 1);
                    });
                });
            };
            // ============================================
            // SAVE FUNCTIONS
            // ============================================
            // ============================================================
            // ✅ SAVE — SweetAlert confirmation FIRST, then fetch.
            //    Same look & guard pattern as procurement.ts.
            // ============================================================
            let isSaving = false;
            const saveAllDataToSingleJSON = async () => {
                // Guard — only block while a save is in flight
                if (isSaving) {
                    console.log("⛔ saveAllDataToSingleJSON blocked: save already in progress");
                    return;
                }
                // 🔐 Role gate: writing data needs a create/edit role
                if (!canCreateRecord()) {
                    toast("You do not have permission to save document data", "error");
                    return;
                }
                const selectedProjectId = parseInt(projectSelect.value);
                if (!selectedProjectId) {
                    toast("Please select a project first", "warning");
                    return;
                }
                const project = projects.find(p => p.Id === selectedProjectId);
                if (!project) {
                    toast("Project not found", "error");
                    return;
                }
                isSaving = true;
                try {
                    const isEditing = formSection.classList.contains("dcs-edit-mode");
                    // ========== CONFIRMATION ==========
                    const confirmResult = await Swal.fire({
                        title: isEditing ? "Save Changes?" : "Save Document Data?",
                        html: isEditing
                            ? `Do you want to save the changes made to <strong>${escapeHtml(project.Title)}</strong>?`
                            : `Do you want to save the document data for <strong>${escapeHtml(project.Title)}</strong>?`,
                        icon: "question",
                        showCancelButton: true,
                        confirmButtonText: isEditing ? "Yes, Save Changes" : "Yes, Save",
                        cancelButtonText: "Cancel",
                        confirmButtonColor: "#605eee",
                        cancelButtonColor: "#6c757d",
                        reverseButtons: true,
                        focusCancel: true,
                        allowOutsideClick: false,
                        customClass: { popup: "proc-confirm-popup" },
                        didOpen: () => {
                            const c = document.querySelector(".swal2-container");
                            if (c)
                                c.style.zIndex = "2147483647";
                        },
                    });
                    if (!confirmResult.isConfirmed) {
                        console.log("🔵 save cancelled by user");
                        return;
                    }
                    // ========== ACTUAL SAVE ==========
                    showExportProgress("Saving document data...");
                    const saveButtons = getSaveButtons();
                    saveButtons.forEach(btn => { btn.disabled = true; });
                    try {
                        const inputs = collectTableData(inputTableBody, true);
                        const outputs = collectTableData(outputTableBody, false);
                        const structuredData = {
                            projectId: selectedProjectId,
                            projectTitle: project.Title,
                            lastUpdated: new Date().toISOString(),
                            documentStructure: { inputs, outputs },
                            summary: {
                                totalDocuments: inputs.length + outputs.length,
                                totalRevisions: inputs.reduce((sum, d) => { var _a; return sum + (((_a = d.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0) +
                                    outputs.reduce((sum, d) => { var _a; return sum + (((_a = d.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0),
                                inputDocuments: inputs.length,
                                outputDocuments: outputs.length,
                                byStatus: {},
                            },
                        };
                        [...inputs, ...outputs].forEach(doc => {
                            const status = doc.status || "Unknown";
                            structuredData.summary.byStatus[status] =
                                (structuredData.summary.byStatus[status] || 0) + 1;
                        });
                        const jsonString = JSON.stringify(structuredData);
                        const existing = dcsData.find((item) => {
                            const projId = item[`${fieldMappings.Project}Id`] ||
                                (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
                            return projId === selectedProjectId;
                        });
                        const digest = await getRequestDigest();
                        if (existing) {
                            const res = await fetch(`${apiBase}(${existing.Id})`, {
                                method: "POST",
                                headers: {
                                    Accept: "application/json;odata=verbose",
                                    "Content-Type": "application/json;odata=verbose",
                                    "X-RequestDigest": digest,
                                    "IF-MATCH": "*",
                                    "X-HTTP-Method": "MERGE",
                                },
                                credentials: "same-origin",
                                body: JSON.stringify({
                                    __metadata: { type: listEntityTypeName },
                                    [fieldMappings.DCSData]: jsonString,
                                }),
                            });
                            if (!res.ok)
                                throw new Error(`Failed to update record: ${res.status}`);
                            console.log("✅ Record updated:", existing.Id);
                            toast("Document data updated successfully", "success");
                        }
                        else {
                            const res = await fetch(apiBase, {
                                method: "POST",
                                headers: {
                                    Accept: "application/json;odata=verbose",
                                    "Content-Type": "application/json;odata=verbose",
                                    "X-RequestDigest": digest,
                                },
                                credentials: "same-origin",
                                body: JSON.stringify({
                                    __metadata: { type: listEntityTypeName },
                                    Title: project.Title,
                                    [`${fieldMappings.Project}Id`]: project.Id,
                                    [fieldMappings.DCSData]: jsonString,
                                }),
                            });
                            if (!res.ok)
                                throw new Error(`Failed to create record: ${res.status}`);
                            console.log("✅ Record created for project", selectedProjectId);
                            toast("Document data saved successfully", "success");
                        }
                        hideExportProgress();
                        await fetchData();
                        await handleBackToTable(); // no popup — the user just confirmed the save
                    }
                    catch (error) {
                        hideExportProgress();
                        console.error("❌ Save error:", error);
                        const msg = error instanceof Error ? error.message : "Failed to save document data";
                        toast(msg, "error");
                    }
                    finally {
                        saveButtons.forEach(btn => { btn.disabled = false; });
                    }
                }
                finally {
                    isSaving = false;
                }
            };
            // Collect rows of one table into DCSDocumentItem[]
            const collectTableData = (body, isInput) => {
                const docs = [];
                const headers = Array.from(body.querySelectorAll('tr.dcs-document-header'));
                headers.forEach((mainRow, index) => {
                    var _a, _b, _c, _d;
                    const docId = mainRow.getAttribute('data-doc-id') || `doc_${index}_${Date.now()}`;
                    const getVal = (sel) => {
                        const el = mainRow.querySelector(sel);
                        return el ? (el.value || "").trim() : "";
                    };
                    // Revisions: siblings of the main row that are not the main row
                    const tbody = mainRow.parentElement;
                    const groupRows = Array.from(tbody.querySelectorAll(`tr[data-doc-id="${docId}"]`));
                    const revRows = groupRows.slice(1);
                    const revisions = revRows.map((revRow, rIndex) => {
                        var _a, _b;
                        const revSel = (sel, idx) => {
                            const els = revRow.querySelectorAll(sel);
                            const el = els[idx];
                            return el ? (el.value || "").trim() : "";
                        };
                        return {
                            revNumber: parseInt(((_b = (_a = revRow.querySelector('.dcs-rev-cell')) === null || _a === void 0 ? void 0 : _a.textContent) === null || _b === void 0 ? void 0 : _b.trim()) || String(rIndex + 1)) || (rIndex + 1),
                            // revision rows are [revNo, originalPlan, revisedPlan, actualSubmission, submissionTx, approvalDate, approvalTx, actions]
                            originalPlan: revSel('.dcs-date-input', 0),
                            revisedPlan: revSel('.dcs-date-input', 1),
                            actualSubmission: revSel('.dcs-date-input', 2),
                            submissionTransmittal: revSel('.dcs-transmittal-input', 0),
                            approvalDate: revSel('.dcs-date-input', 3),
                            approvalTransmittal: revSel('.dcs-transmittal-input', 1)
                        };
                    });
                    docs.push({
                        docId,
                        docNo: getVal('.dcs-doc-no'),
                        description: getVal('.dcs-description-input'),
                        category: getVal('.dcs-category-select') || "A",
                        enggDivision: getVal('.dcs-engg-division-select'),
                        revNo: 0,
                        originalPlan: getVal('.dcs-date-input'),
                        revisedPlan: ((_a = mainRow.querySelectorAll('.dcs-date-input')[1]) === null || _a === void 0 ? void 0 : _a.value) || "",
                        actualSubmission: ((_b = mainRow.querySelectorAll('.dcs-date-input')[2]) === null || _b === void 0 ? void 0 : _b.value) || "",
                        submissionTransmittal: getVal('.dcs-transmittal-input'),
                        approvalDate: ((_c = mainRow.querySelectorAll('.dcs-date-input')[3]) === null || _c === void 0 ? void 0 : _c.value) || "",
                        approvalTransmittal: ((_d = mainRow.querySelectorAll('.dcs-transmittal-input')[1]) === null || _d === void 0 ? void 0 : _d.value) || "",
                        status: getVal('.dcs-status-select') || (isInput ? "Pending" : "Not Submitted"),
                        remarks: getVal('.dcs-remarks-input'),
                        isInput,
                        revisions,
                        docType: isInput ? 'input' : 'output'
                    });
                });
                return docs;
            };
            // Reset the two document tables to their default single row (edit mode)
            const resetDocumentTables = () => {
                inputTableBody.innerHTML = createDocumentRowHTML('input', 1, `default_input_${Date.now()}`);
                outputTableBody.innerHTML = createDocumentRowHTML('output', 1, `default_output_${Date.now()}`);
                ensureDocumentButtonsVisible();
                updateAllSerialNumbers();
            };
            const resetForm = () => {
                dcsForm.reset();
                projectSelect.disabled = false;
                formSection.classList.remove('dcs-view-mode', 'dcs-edit-mode');
                if (formTitle)
                    formTitle.textContent = "Document Control Index";
                resetProjectDetailsDisplay();
                existingDataAlert.style.display = 'none';
                importButtonContainer.style.display = 'none';
                containerTable.style.display = 'none';
                resetDocumentTables();
                getSaveButtons().forEach(btn => {
                    btn.style.display = 'inline-flex';
                    btn.innerHTML = '💾 Save Document Data';
                });
                console.log('🔄 Form reset');
            };
            // Update the project info panel when a project is picked in the form
            const updateProjectDetails = (projectId) => {
                const project = projects.find(p => p.Id === projectId);
                if (!project) {
                    resetProjectDetailsDisplay();
                    return;
                }
                setProjectDetailsDisplay(project);
                // Existing data? - load it so the user edits instead of creating a duplicate
                const record = dcsData.find((item) => {
                    const projId = item[`${fieldMappings.Project}Id`] ||
                        (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
                    return projId === projectId;
                });
                if (record && record[fieldMappings.DCSData]) {
                    existingDataAlert.style.display = 'block';
                    const parsedData = parseDCSData(record[fieldMappings.DCSData]);
                    if (parsedData) {
                        inputTableBody.innerHTML = '';
                        outputTableBody.innerHTML = '';
                        loadFormData(parsedData);
                    }
                }
                else {
                    existingDataAlert.style.display = 'none';
                }
            };
            // onProjectChange - alias kept for the select change handler
            const onProjectChange = (projectId) => updateProjectDetails(projectId);
            // Load parsed JSON into the editable tables
            const loadFormData = (projectData) => {
                const inputs = getInputsFromProjectData(projectData);
                const outputs = getOutputsFromProjectData(projectData);
                inputTableBody.innerHTML = '';
                outputTableBody.innerHTML = '';
                if (inputs.length > 0) {
                    inputs.forEach((doc, index) => {
                        inputTableBody.appendChild(createEditableDocumentRow(doc, index + 1, true));
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach(rev => {
                                inputTableBody.insertAdjacentHTML('beforeend', buildRevisionRowHTML(doc.docId, rev));
                            });
                            updateMergedCells(doc.docId);
                        }
                    });
                }
                else {
                    inputTableBody.innerHTML = createDocumentRowHTML('input', 1, `default_input_${Date.now()}`);
                }
                if (outputs.length > 0) {
                    outputs.forEach((doc, index) => {
                        outputTableBody.appendChild(createEditableDocumentRow(doc, index + 1, false));
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach(rev => {
                                outputTableBody.insertAdjacentHTML('beforeend', buildRevisionRowHTML(doc.docId, rev));
                            });
                            updateMergedCells(doc.docId);
                        }
                    });
                }
                else {
                    outputTableBody.innerHTML = createDocumentRowHTML('output', 1, `default_output_${Date.now()}`);
                }
                containerTable.style.display = 'block';
                importButtonContainer.style.display = 'flex';
                ensureDocumentButtonsVisible();
                updateAllSerialNumbers();
            };
            // ============================================
            // PAGINATION
            // ============================================
            const initializePagination = (data) => {
                // keep the sorted order (applySort is a no-op when nothing is sorted)
                filteredDCSData = applySort(data);
                const total = filteredDCSData.length;
                totalPages = Math.max(1, Math.ceil(total / pageSize));
                if (currentPage > totalPages)
                    currentPage = totalPages;
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, total);
                const pageData = filteredDCSData.slice(startIndex, endIndex);
                renderTable(pageData);
                updatePaginationUI(total);
            };
            const goToPage = (page) => {
                if (page < 1 || page > totalPages || page === currentPage)
                    return;
                currentPage = page;
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, filteredDCSData.length);
                renderTable(filteredDCSData.slice(startIndex, endIndex));
                updatePaginationUI(filteredDCSData.length);
            };
            const changePageSize = (newSize) => {
                // ✅ keep the FIRST visible record in view: work out its index with the
                //    OLD page size BEFORE switching, then open the page that contains it.
                const firstVisibleIndex = (currentPage - 1) * pageSize;
                pageSize = newSize;
                currentPage = Math.floor(firstVisibleIndex / pageSize) + 1;
                const total = filteredDCSData.length;
                totalPages = Math.max(1, Math.ceil(total / pageSize));
                if (currentPage > totalPages)
                    currentPage = totalPages;
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, total);
                renderTable(filteredDCSData.slice(startIndex, endIndex));
                updatePaginationUI(total);
            };
            const updatePaginationUI = (total) => {
                const startItem = document.getElementById("startItem");
                const endItem = document.getElementById("endItem");
                const totalItems = document.getElementById("totalItems");
                const totalPagesEl = document.getElementById("totalPages");
                const firstPageBtn = document.getElementById("firstPage");
                const prevPageBtn = document.getElementById("prevPage");
                const nextPageBtn = document.getElementById("nextPage");
                const lastPageBtn = document.getElementById("lastPage");
                const start = total === 0 ? 0 : (currentPage - 1) * pageSize + 1;
                const end = Math.min(currentPage * pageSize, total);
                if (startItem)
                    startItem.textContent = String(start);
                if (endItem)
                    endItem.textContent = String(end);
                if (totalItems)
                    totalItems.textContent = String(total);
                if (totalPagesEl)
                    totalPagesEl.textContent = String(totalPages);
                if (firstPageBtn)
                    firstPageBtn.disabled = currentPage === 1;
                if (prevPageBtn)
                    prevPageBtn.disabled = currentPage === 1;
                if (nextPageBtn)
                    nextPageBtn.disabled = currentPage === totalPages;
                if (lastPageBtn)
                    lastPageBtn.disabled = currentPage === totalPages;
                renderPageNumbers();
            };
            const renderPageNumbers = () => {
                const container = document.getElementById("pageNumbers");
                if (!container)
                    return;
                container.innerHTML = '';
                // ✅ One page is still a page: the reference pager (Inventory / Revenue /
                //    Daily Sprint) always draws the tile, so a single-page result reads
                //    "1 of 1" instead of leaving the tile strip empty. Only a truly
                //    empty result set has nothing to draw.
                if (filteredDCSData.length === 0)
                    return;
                const pages = [];
                const maxVisible = 5;
                if (totalPages <= maxVisible + 2) {
                    for (let i = 1; i <= totalPages; i++)
                        pages.push(i);
                }
                else if (currentPage <= 3) {
                    for (let i = 1; i <= maxVisible; i++)
                        pages.push(i);
                    pages.push('...');
                    pages.push(totalPages);
                }
                else if (currentPage >= totalPages - 2) {
                    pages.push(1);
                    pages.push('...');
                    for (let i = totalPages - maxVisible + 1; i <= totalPages; i++)
                        pages.push(i);
                }
                else {
                    pages.push(1);
                    pages.push('...');
                    for (let i = currentPage - 1; i <= currentPage + 1; i++)
                        pages.push(i);
                    pages.push('...');
                    pages.push(totalPages);
                }
                pages.forEach(p => {
                    if (p === '...') {
                        const dots = document.createElement('span');
                        dots.className = 'dcs-page-number dots';
                        dots.textContent = '...';
                        container.appendChild(dots);
                        return;
                    }
                    const btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'dcs-page-number' + (p === currentPage ? ' active' : '');
                    btn.textContent = String(p);
                    btn.addEventListener('click', () => goToPage(Number(p)));
                    container.appendChild(btn);
                });
            };
            const initializeSharePoint = async () => {
                var _a, _b, _c;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=Title,ItemCount`, {
                        headers: { Accept: "application/json;odata=verbose" },
                        credentials: 'same-origin'
                    });
                    if (!res.ok)
                        throw new Error(`List '${listName}' not found (${res.status})`);
                    const data = await res.json();
                    listEntityTypeName = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.ListItemEntityTypeFullName) || `SP.Data.${listName}ListItem`;
                    console.log(`✅ Connected to list '${listName}' (${(_c = (_b = data === null || data === void 0 ? void 0 : data.d) === null || _b === void 0 ? void 0 : _b.ItemCount) !== null && _c !== void 0 ? _c : 0} items)`);
                }
                catch (error) {
                    console.error("Error connecting to SharePoint:", error);
                    throw error;
                }
            };
            // Boot
            await initApp();
        }, 0);
    }
};
export default DCSModule;
//# sourceMappingURL=dci.js.map
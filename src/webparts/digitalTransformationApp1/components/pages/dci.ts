import "../styles/common.css";
import "../styles/dcs.css";
import * as XLSX from 'xlsx-js-style';

// ============ SHARED AUTH + ROLE BASED ACCESS (same as SEIBIES / inventory) ============
import {
  authManager,
  hasFullControl,
  canContribute,
  canRead,
  canFreeze,
  isGHHead,
  isPMManager,
  isEngineeringUser,
  getUserGroups,
  getCurrentUserFullName,
  getPermissionLevel,
  canEditProject,
  canCreateRecord,
  canReadAllProjects,
  // canViewForm,  
  getAssignedProjects,
  updateUIBasedOnPermissions,
} from "./auth";

// =================== DCS OPTION LISTS (Engg Division / Status) ===================
const ENGG_DIVISIONS: string[] = [
  "Automation",
  "SCADA",
  "Instrumentation",
  "PMS",
  "HSE",
  "Project Management"
];

// Status options for the DESIGN OUTPUT table
const OUTPUT_STATUS_OPTIONS: string[] = [
  "Not Submitted",
  "Submitted",
  "Approved",
  "Approved with comments",
  "Rejected",
  "Retained for Information",
  "Document Deleted/cancelled"
];

// Status options for the DESIGN INPUT table
const INPUT_STATUS_OPTIONS: string[] = [
  "Pending",
  "Received",
  "Partially Received",
  "Revised inputs required"
];

const getStatusOptions = (tableType: 'input' | 'output'): string[] =>
  tableType === 'input' ? INPUT_STATUS_OPTIONS : OUTPUT_STATUS_OPTIONS;

const getDefaultStatus = (tableType: 'input' | 'output'): string =>
  tableType === 'input' ? "Pending" : "Not Submitted";

const escapeOptionText = (value: string): string =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

// Build <option> list for the Status select. If the stored value is a legacy status
// (e.g. "Under Review") that is no longer in the list, it is kept as a selectable
// option so existing data is never lost/changed silently.
const buildStatusOptionsHTML = (tableType: 'input' | 'output', selected?: string): string => {
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
const buildEnggDivisionOptionsHTML = (selected?: string): string => {
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
      <h1>Document Control System</h1>
      <div>
        <button class="dcs-btn-primary dcs-show-form-btn">
          + Add Project Documents
        </button>
      </div>
    </div>
    
    
     <div class="dcs-table-actions">
      <div class="dcs-tableColp4">
        <label for="filterProject">Project</label>
        <select id="filterProject" class="dcs-form-control">
          <option value="">All Projects</option>
        </select>
      </div>
      <div class="dcs-tableColp4">
        <label for="filterCustomer">Customer</label>
        <select id="filterCustomer" class="dcs-form-control">
          <option value="">All Customers</option>
        </select>
      </div>
      <div class="dcs-tableColp4">
        <label for="filterProjectManager">Project Manager</label>
        <select id="filterProjectManager" class="dcs-form-control">
          <option value="">All Managers</option>
        </select>
      </div>
      <div class="dcs-tableColp4">
        <label for="filterProjectReference">Project Reference</label>
        <select id="filterProjectReference" class="dcs-form-control">
          <option value="">All References</option>
        </select>
      </div>
      <div class="dcs-tableColp4">  
        
          <button class="dcs-btn-secondary dcs-clear-filters-btn">Reset Filter</button>  
         
      </div>
      <div class="dcs-tableColp4">  
        
          <button class="dcs-btn-primary dcs-export-excel-btn"> Download Excel  </button>
         
      </div>
    </div>

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
        <tbody id="dcsTableBody">
          <!-- ... existing code ... -->
        </tbody>
      </table>
    </div>

    <!-- ✅ Loading state (same pattern as Inventory / Procurement): spinner + text,
         shown while DCS records are being fetched. -->
    <div class="dcs-loading" id="dcsLoading">
      <span class="dcs-loading-spinner"></span>
      <span class="dcs-loading-text">Loading DCS data...</span>
    </div>

    <div class="dcs-pagination-container">
      <div class="dcs-pagination-left">
        <button class="dcs-pagination-btn dcs-pg-icon" id="firstPage" disabled title="First page">&#171;</button>
        <button class="dcs-pagination-btn dcs-pg-icon" id="prevPage" disabled title="Previous page">&#8249;</button>

        <div class="dcs-pagination-info">
          Page <span id="currentPage">1</span> of <span id="totalPages">1</span>
        </div>

        <div class="dcs-pagination-pages" id="pageNumbers"></div>

        <button class="dcs-pagination-btn dcs-pg-icon" id="nextPage" disabled title="Next page">&#8250;</button>
        <button class="dcs-pagination-btn dcs-pg-icon" id="lastPage" disabled title="Last page">&#187;</button>
      </div>

      <div class="dcs-pagination-right">
        <div class="dcs-page-size-selector">
          <label for="pageSizeSelect">Show:</label>
          <select id="pageSizeSelect" class="dcs-page-size-select">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
          <span>per page</span>
        </div>

        <div class="dcs-total-badge">
          Showing <span class="dcs-total-range"><span id="startItem">0</span>-<span id="endItem">0</span></span>
          of <span class="dcs-total-count" id="totalItems">0</span> records
        </div>
      </div>
    </div>
  </section>

  <section class="form-section dcs-card" id="dcsFormSection" style="display:none;">
  <form id="dcsForm">
  <div class="dcs-page-header stick">
    <h1 id="dcsFormTitle">Add Project Documents</h1>
    <!-- ✅ TOP SAVE BUTTON - same type="submit" + .dcs-btn-primary
         as the bottom button, so both trigger the same submit handler
         and stay in sync (visibility, saving state). The header is
         position:sticky, so this button stays visible while scrolling
         through the long document tables. -->
    <button type="submit"
            form="dcsForm"
            class="dcs-btn-primary dcs-btn-save-top"
            title="Save Document Data">💾 Save Document Data</button>
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
          <div class="dcs-view-mode-note" id="dcsViewModeNote">
            <span style="font-size: 16px;">ℹ️</span>
            <div>
              <strong>View Mode:</strong> This is a read-only view of all document control data for this project.
              Click "Edit" on the main page to modify documents.
            </div>
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
          ⬅ Back to List
        </button>
      </div>
    </form>
  </section>
</div>
`;

// Declare SharePoint context for TypeScript
declare const _spPageContextInfo: any;

interface ProjectItem {
  Id: number;
  Title: string;
  CustomerName?: string;
  ProjectManagerId?: number;
  ProjectManager?: string | { Title: string; Email?: string };
  PC?: string;
  ProjectReference?: string;
  GroupHead1?: string | { Title: string; Email?: string };  // ✅ Changed to Person/Group
  GroupHead1Id?: number;  // ✅ Person/Group ID
}

interface RevisionItem {
  revNumber: number;
  originalPlan: string;
  revisedPlan: string;
  actualSubmission: string;
  submissionTransmittal: string;
  approvalDate: string;
  approvalTransmittal: string;
  status?: string;
  remarks?: string;
}

interface DCSDocumentItem {
  docId: string;
  docNo: string;
  description: string;
  category: string;
  enggDivision?: string;
  revNo: number;
  originalPlan: string;
  revisedPlan: string;
  actualSubmission: string;
  submissionTransmittal: string;
  approvalDate: string;
  approvalTransmittal: string;
  status: string;
  remarks: string;
  isInput: boolean;
  revisions: RevisionItem[];
  docType?: 'input' | 'output';
}

interface LegacyDCSProjectData {
  projectId: number;
  inputs?: DCSDocumentItem[];
  outputs?: DCSDocumentItem[];
}

interface DCSProjectData {
  projectId: number;
  projectTitle?: string;
  lastUpdated?: string;
  documentStructure: {
    inputs: DCSDocumentItem[];
    outputs: DCSDocumentItem[];
  };
  summary?: {
    totalDocuments: number;
    totalRevisions: number;
    inputDocuments: number;
    outputDocuments: number;
    byStatus?: {
      [key: string]: number;
    };
  };
}

interface FieldMappings {
  DCSData: string;
  Project: string;
}

type AnyDCSProjectData = DCSProjectData | LegacyDCSProjectData;

const DCSModule = {
  content: dcsContent,

  script: () => {
    setTimeout(async () => {
      // Configuration
      const siteUrl =
        typeof _spPageContextInfo !== "undefined"
          ? _spPageContextInfo.webAbsoluteUrl
          : window.location.origin + "/sites/DigitalTransformationTest";

      const listName = "DCSTracker";
      const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

      // State variables
      let fieldMappings: FieldMappings = {
        DCSData: "DCSData",
        Project: "Project",
      };

      let listEntityTypeName: string = "";
      let projects: ProjectItem[] = [];
      let editingId: number | null = null;
      let dcsData: any[] = [];

      // Toast function with proper typing
      const toast = (message: string, type: 'success' | 'error' | 'warning' = 'success'): void => {
        alert(`${type === 'success' ? '✅' : type === 'error' ? '❌' : '⚠️'} ${message}`);
      };

      // ============================================
      // LOADING OVERLAY (same pattern as master.ts)
      // ============================================

      const createLoadingOverlay = (message = "Processing..."): HTMLElement => {
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

      const removeLoadingOverlay = (): void => {
        document.querySelectorAll(".loading-overlay").forEach(el => el.remove());
      };

      /** mount the loader on the VISIBLE card (table section or form section) —
          the table card while data loads, the form card while saving */
      const loaderHost = (): HTMLElement => {
        const sections = Array.from(document.querySelectorAll<HTMLElement>('.table-section, .form-section'));
        const visible = sections.find(s => s.offsetParent !== null);
        return visible || (document.querySelector('.table-section') as HTMLElement | null) || document.body;
      };

      // DOM Elements
      const projectSelect = document.getElementById("dcsProjectSelect") as HTMLSelectElement;
      const showFormBtn = document.querySelector(".dcs-show-form-btn") as HTMLButtonElement;
      const exportExcelBtn = document.querySelector(".dcs-export-excel-btn") as HTMLButtonElement;
      const formSection = document.getElementById("dcsFormSection") as HTMLElement;
      const tableSection = document.getElementById("dcsTableSection") as HTMLElement;
      const backBtn = document.querySelector(".dcs-back-to-table") as HTMLButtonElement;
      const outputTableBody = document.getElementById("dcsOutputTableBody") as HTMLElement;
      const inputTableBody = document.getElementById("dcsInputTableBody") as HTMLElement;
      const containerTable = document.getElementById("dcsContainerTable") as HTMLElement;
      const dcsForm = document.getElementById("dcsForm") as HTMLFormElement;
      const formTitle = document.getElementById("dcsFormTitle") as HTMLHeadingElement;
      const existingDataAlert = document.getElementById("dcsExistingDataAlert") as HTMLDivElement;
      const addInputDocBtn = document.querySelector(".dcs-add-input-document") as HTMLButtonElement;
      const addOutputDocBtn = document.querySelector(".dcs-add-output-document") as HTMLButtonElement;
      const importButtonContainer = document.getElementById("dcsImportButtonContainer") as HTMLDivElement;
      const resetFormBtn = document.getElementById("dcsResetFormBtn") as HTMLButtonElement;

      // Filter Elements
      const filterProject = document.getElementById("filterProject") as HTMLSelectElement;
      const filterCustomer = document.getElementById("filterCustomer") as HTMLSelectElement;
      const filterProjectManager = document.getElementById("filterProjectManager") as HTMLSelectElement;
      const filterProjectReference = document.getElementById("filterProjectReference") as HTMLSelectElement;
      const clearFiltersBtn = document.querySelector(".dcs-clear-filters-btn") as HTMLButtonElement;

      const projectPC = document.getElementById("projectPC") as HTMLParagraphElement;
      const projectCustomer = document.getElementById("projectCustomer") as HTMLParagraphElement;
      const projectReference = document.getElementById("projectReference") as HTMLParagraphElement;
      const projectGroupHead = document.getElementById("projectGroupHead") as HTMLParagraphElement;
      const projectManager = document.getElementById("projectManager") as HTMLParagraphElement;

      // Pagination variables
      let currentPage = 1;
      let pageSize = 25; // ✅ default 25 rows per page
      let totalPages = 1;
      let filteredDCSData: any[] = [];

      // ---- Main table sorting state ----
      let sortKey: string | null = null;
      let sortDir: 'asc' | 'desc' = 'asc';

      const getRequestDigest = async (): Promise<string> => {
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
        } catch (error) {
          console.error("Error getting request digest:", error);
          toast("Failed to get security token. Please refresh the page.", "error");
          throw error;
        }
      };

      const initializeEventListeners = (): void => {
        // 1. Show form button
        if (showFormBtn) {
          showFormBtn.removeEventListener('click', handleShowForm);
          showFormBtn.addEventListener('click', handleShowForm);
        }

        // 2. Back to table button
        if (backBtn) {
          backBtn.removeEventListener('click', handleBackToTable);
          backBtn.addEventListener('click', handleBackToTable);
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

        // 5. Filter events
        const filters = [
          { el: filterProject, event: 'change', handler: filterData },
          { el: filterCustomer, event: 'change', handler: filterData },
          { el: filterProjectManager, event: 'change', handler: filterData },
          { el: filterProjectReference, event: 'change', handler: filterData },
          { el: clearFiltersBtn, event: 'click', handler: clearFilters }
        ];

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

        filters.forEach(({ el, event, handler }) => {
          if (el) {
            el.removeEventListener(event, handler);
            el.addEventListener(event, handler);
          }
        });

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
      const applyRoleBasedUI = (): void => {
        // Add button + import button only for create roles
        if (showFormBtn) {
          showFormBtn.style.display = canCreateRecord() ? "inline-flex" : "none";
        }
        const importExcelBtn = document.querySelector('.dcs-import-excel-btn') as HTMLElement | null;
        if (importExcelBtn) {
          importExcelBtn.style.display = canCreateRecord() ? "inline-flex" : "none";
        }
        // Export / Save via the shared helper (export = any read role, save = writers)
        updateUIBasedOnPermissions({
          addBtn: '.dcs-show-form-btn',
          exportBtn: '.dcs-export-excel-btn',
          importBtn: '.dcs-import-excel-btn',
          saveBtn: '.dcs-form-actions button[type="submit"]'
        });
      };

      const handleShowForm = (): void => {
        // 🔐 Role gate: only Admin / GH Head / PM Manager can create a new record
        if (!canCreateRecord()) {
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
        if (addInputDocBtn) addInputDocBtn.style.display = "inline-block";
        if (addOutputDocBtn) addOutputDocBtn.style.display = "inline-block";

        // Show the container table
        containerTable.style.display = "block";

        // Attach listeners
        setTimeout(() => {
          attachDocumentEventListeners();
        }, 100);
      };
const getSaveButtons = (): HTMLButtonElement[] =>
  Array.from(document.querySelectorAll<HTMLButtonElement>(
    'button.dcs-btn-primary[type="submit"][form="dcsForm"], .dcs-form-actions button.dcs-btn-primary[type="submit"]'
  ));
      const handleBackToTable = (): void => {
        // Remove view mode class - CSS will handle showing buttons again
        formSection.classList.remove("dcs-view-mode");

        // Enable project select
        projectSelect.disabled = false;

        // Show submit button
       getSaveButtons().forEach(btn => {
  btn.style.display = "inline-flex";
});
        // Reset editing state
        editingId = null;

        // Show/hide sections
        formSection.style.display = 'none';
        tableSection.style.display = 'block';

        // Reset form
        resetForm();
      };

      const handleFormSubmit = async (e: Event): Promise<void> => {
        e.preventDefault();
        await saveAllDataToSingleJSON();
      };

      const handleProjectSelect = (e: Event): void => {
        const selectedId = parseInt((e.target as HTMLSelectElement).value);
        if (selectedId) {
          onProjectChange(selectedId);
        }
      };

      // ============================================ //
      // OUTER-STICKY HEADER (main table)             //
      //   The page scrolls vertically. The main      //
      //   page-header is sticky at top:0. The table  //
      //   header is sticky at top = main header      //
      //   height, so it 'touches' the main header    //
      //   and pins below it. No per-cell JS offset.  //
      // ============================================ //
      const updateMainTableStickyOffset = (): void => {
        const section = document.querySelector(".table-section") as HTMLElement | null;
        const header = section?.querySelector(".dcs-page-header") as HTMLElement | null;
        if (!section || !header) return;
        // Fall back to a sensible minimum in case offsetHeight is 0 (e.g.
        // before fonts/images load), so the table header never sticks at
        // top:0 on top of the page header.
        const headerH = Math.max(header.offsetHeight || 0, 58);
        section.style.setProperty("--dcs-sticky-header-top", `${headerH}px`);
      };

      const initApp = async (): Promise<void> => {
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

          // ✅ Role based UI (hide Add / Import for read-only roles etc.)
          applyRoleBasedUI();

          console.log('✅ App initialization complete');

        } catch (error) {
          console.error("❌ Initialization error:", error);
          toast("Failed to initialize application", "error");
        } finally {
          removeLoadingOverlay();
        }
      };

      const setupPaginationEventListeners = (): void => {
        // Navigation buttons
        const firstPageBtn = document.getElementById("firstPage") as HTMLButtonElement;
        const prevPageBtn = document.getElementById("prevPage") as HTMLButtonElement;
        const nextPageBtn = document.getElementById("nextPage") as HTMLButtonElement;
        const lastPageBtn = document.getElementById("lastPage") as HTMLButtonElement;
        const pageSizeSelect = document.getElementById("pageSizeSelect") as HTMLSelectElement;

        if (firstPageBtn) firstPageBtn.addEventListener("click", () => goToPage(1));
        if (prevPageBtn) prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
        if (nextPageBtn) nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
        if (lastPageBtn) lastPageBtn.addEventListener("click", () => goToPage(totalPages));

        // Page size selector
        if (pageSizeSelect) {
          pageSizeSelect.addEventListener("change", (e) => {
            const newSize = parseInt((e.target as HTMLSelectElement).value);
            changePageSize(newSize);
          });
        }

        // ---- Main table column sorting (delegated on the static thead) ----
        const mainThead = document.querySelector('#dcsTable thead');
        if (mainThead) mainThead.addEventListener('click', handleSortClick);
        updateSortIndicators();
      };

      // Load Projects
      const loadProjects = async (): Promise<void> => {
        try {
          // ✅ Include GroupHead1 with expand
          const res = await fetch(
            `${siteUrl}/_api/web/lists/getbytitle('Project List')/items?$select=Id,Title,CustomerName,GroupHead1/Title,GroupHead1Id,PC,ProjectReference,ProjectManager/Title,ProjectManagerId&$expand=GroupHead1,ProjectManager&$top=1000`,
            {
              headers: { Accept: "application/json;odata=verbose" },
              credentials: 'same-origin'
            }
          );
          if (!res.ok) throw new Error(`Failed to load projects: ${res.status}`);

          const data = await res.json();
          const allProjects: any[] = data?.d?.results || [];

          // Filter out projects without a title
          const validProjects = allProjects.filter(p => p.Title && p.Title.trim() !== "");

          // Map to ProjectItem with proper Person/Group handling
          const mappedProjects: ProjectItem[] = validProjects.map((p: any) => ({
            Id: p.Id,
            Title: p.Title || "",
            CustomerName: p.CustomerName || "",
            GroupHead1: p.GroupHead1 || "",
            GroupHead1Id: p.GroupHead1Id,
            PC: p.PC || "",
            ProjectReference: p.ProjectReference || "",
            ProjectManager: p.ProjectManager || "",
            ProjectManagerId: p.ProjectManagerId
          }));

          // Get unique projects by Title
          const uniqueProjects = Array.from(
            new Map(mappedProjects.map((p: ProjectItem) => [p.Title, p])).values()
          );

          projects = uniqueProjects;

          // 🔐 ROLE SCOPE (same as SEIBIES / inventory): Admin + Read-All groups see
          // every project; GH Head / PM Manager / Engg roles only see the projects
          // assigned to them (getAssignedProjects already returns ALL for
          // FULL_CONTROL and READ_ALL users).
          const visibleProjects = getAssignedProjects(uniqueProjects);
          console.log(`🔐 Visible projects for this role: ${visibleProjects.length} / ${uniqueProjects.length}`);

          // Populate project select in form (visible scope only)
          projectSelect.innerHTML =
            '<option value="">Select Project</option>' +
            visibleProjects
              .map((p) => `<option value="${p.Id}">${p.Title}</option>`)
              .join("");

          // Populate project filter in table section (visible scope only)
          filterProject.innerHTML =
            '<option value="">All Projects</option>' +
            visibleProjects
              .map((p) => `<option value="${p.Title}">${p.Title}</option>`)
              .join("");

        } catch (error) {
          console.error("Error loading projects:", error);
          toast("Failed to load projects", "error");
        }
      };

      // ✅ Edit title like Procurement: "Edit Document Control Data - <Project> (<Ref>)"
      // ✅ Title shows only the project name (Ref); "Edit Document Control Data - " prefix removed
const setEditTitle = (project?: ProjectItem | null): void => {
  if (!formTitle) return;
  const name = project?.Title || "";
  const ref = project?.ProjectReference || "";
  formTitle.textContent = name
    ? `${name}${ref ? ` (${ref})` : ""}`
    : "Edit Document Control Data";
  formSection.classList.add("dcs-edit-mode");
};

      // Fill the project info rows (left panel of the form) for the given project
      const setProjectDetailsDisplay = (project: ProjectItem | undefined): void => {
        projectPC.textContent = project?.PC || "-";
        projectCustomer.textContent = project?.CustomerName || "-";
        projectReference.textContent = project?.ProjectReference || "-";
        projectGroupHead.textContent = project ? getGroupHeadDisplayName(project.GroupHead1) : "-";
        projectManager.textContent = project ? getProjectManagerDisplayName(project.ProjectManager) : "-";
      };

      // Reset the project info rows to "-" (the Project select itself is kept intact)
      const resetProjectDetailsDisplay = (): void => setProjectDetailsDisplay(undefined);

      const getGroupHeadDisplayName = (groupHead: string | { Title: string; Email?: string } | undefined): string => {
        if (!groupHead) return "-";
        if (typeof groupHead === "object") {
          return groupHead.Title || "-";
        }
        return groupHead;
      };

      const getProjectManagerDisplayName = (manager: string | { Title: string; Email?: string } | undefined): string => {
        if (!manager) return "-";
        if (typeof manager === "object") {
          return manager.Title || "-";
        }
        return manager;
      };

      // Parse JSON data from DCSData field
      const parseDCSData = (dcsDataJson: string): DCSProjectData | LegacyDCSProjectData | null => {
        if (!dcsDataJson || dcsDataJson.trim() === "") return null;

        try {
          const data = JSON.parse(dcsDataJson);

          // Check if it's the new format
          if (data.documentStructure && (data.documentStructure.inputs || data.documentStructure.outputs)) {
            return data as DCSProjectData;
          }

          // Check if it's the old format
          if (data.inputs || data.outputs) {
            return data as LegacyDCSProjectData;
          }

          // Check if it's the very old format
          if (Array.isArray(data)) {
            // This is an array, convert to new format
            const inputs: DCSDocumentItem[] = [];
            const outputs: DCSDocumentItem[] = [];

            data.forEach((item: any) => {
              if (item.isInput) {
                inputs.push(item);
              } else {
                outputs.push(item);
              }
            });

            return {
              projectId: data[0]?.projectId || 0,
              documentStructure: {
                inputs,
                outputs
              }
            } as DCSProjectData;
          }

          return data;
        } catch (error) {
          console.error("Error parsing DCSData JSON:", error);
          return null;
        }
      };

      // ✅ Loading state (same as Procurement / Inventory): spinner + text
      // while DCS records are being fetched. Hidden in fetchData's finally.
      const loadingIndicator = document.getElementById("dcsLoading") as HTMLElement | null;
      const showDCSLoading = (): void => {
        if (loadingIndicator) loadingIndicator.style.display = "flex";
      };
      const hideDCSLoading = (): void => {
        if (loadingIndicator) loadingIndicator.style.display = "none";
      };

      const fetchData = async (): Promise<void> => {
        showDCSLoading();
        const loading = createLoadingOverlay("Loading DCS data...");
        loaderHost().appendChild(loading);
        try {
          console.log('📥 Starting fetchData...');

          // Build select fields - try multiple patterns
          const selectFields = [
            "Id",
            "Title",
            fieldMappings.DCSData,
            `${fieldMappings.Project}Id`,
            `${fieldMappings.Project.replace(/Id$/, '')}Id`,
            `${fieldMappings.Project.replace(/Id$/, '')}/Id`,
            `${fieldMappings.Project.replace(/Id$/, '')}/Title`,
          ].filter((field, index, self) => self.indexOf(field) === index).join(",");

          const expandField = fieldMappings.Project.replace(/Id$/, '');

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
          const allData: any[] = data.d.results || [];

          console.log(`📊 Raw data received: ${allData.length} items`);

          // Helper function to extract project ID
          const getProjectId = (item: any): number | null => {
            // Try multiple patterns
            if (item[`${fieldMappings.Project}Id`]) {
              return Number(item[`${fieldMappings.Project}Id`]);
            }
            if (item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]) {
              return Number(item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]);
            }
            if (item[fieldMappings.Project.replace(/Id$/, '')]?.Id) {
              return Number(item[fieldMappings.Project.replace(/Id$/, '')].Id);
            }
            if (item.ProjectId) {
              return Number(item.ProjectId);
            }
            if (item.Project_ID) {
              return Number(item.Project_ID);
            }
            return null;
          };

          // Filter to keep only ONE record per project
          const uniqueProjects = new Map<number, any>();
          const duplicates: any[] = [];

          console.log('🔍 Processing unique projects...');

          allData.forEach((item: any) => {
            const projectId = getProjectId(item);

            if (projectId) {
              if (!uniqueProjects.has(projectId)) {
                uniqueProjects.set(projectId, item);
              } else {
                duplicates.push(item);
                console.warn(`Duplicate found for project ${projectId}: ID ${item.Id}`);
              }
            } else {
              console.warn("No project ID found for item:", item.Id, "Item structure:", Object.keys(item));
            }
          });

          const uniqueRows = Array.from(uniqueProjects.values());

          // 🔐 ROLE SCOPE on rows (same as SEIBIES / inventory loadForecasts):
          // Admin / Read-All groups keep every row; other roles only keep rows for
          // their assigned projects (getAssignedProjects returns ALL for
          // FULL_CONTROL / READ_ALL users).
          const scopeTitles = new Set(getAssignedProjects(projects).map(p => p.Title));
          dcsData = uniqueRows.filter(item => {
            const t = item[fieldMappings.Project]?.Title || "";
            return t && scopeTitles.has(t);
          });

          console.log(`✅ Unique projects: ${uniqueRows.length} (visible to this role: ${dcsData.length}, duplicates: ${duplicates.length})`);

          // If we found duplicates, we could delete them here - Admin (Full Control) only
          if (duplicates.length > 0 && hasFullControl()) {
            console.log(`🗑️ Cleaning up ${duplicates.length} duplicates (admin)...`);
            await cleanupDuplicates(duplicates);
          }

          // Initialize filters with all data
          initializeFilters();

          // Initialize pagination with all data
          initializePagination(dcsData);

          console.log('✅ fetchData completed successfully');

        } catch (error) {
          // ✅ Fix: Type-safe error handling
          console.error("❌ Error fetching data:", error);

          let errorMessage = "Please check your SharePoint configuration";
          if (error instanceof Error) {
            errorMessage = error.message;
          } else if (typeof error === 'string') {
            errorMessage = error;
          }

          const tableBody = document.getElementById("dcsTableBody") as HTMLElement;
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
        } finally {
          hideDCSLoading();
          removeLoadingOverlay();
        }
      };

      const cleanupDuplicates = async (duplicates: any[]): Promise<void> => {
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
            } catch (error) {
              console.error(`Failed to delete duplicate ${duplicate.Id}:`, error);
            }
          }
        } catch (error) {
          console.error("Error in cleanupDuplicates:", error);
        }
      };

      const initializeFilters = (): void => {
        console.log('🎛️ Initializing filters...');

        // Project filter - show all projects
        const accessibleProjectTitles = projects
          .map((p: ProjectItem) => p.Title)
          .filter((title): title is string => !!title);

        filterProject.innerHTML = '<option value="">All Projects</option>' +
          [...new Set(accessibleProjectTitles)]
            .sort()
            .map(project => `<option value="${project}">${project}</option>`)
            .join("");

        // Customer filter - customers from all projects
        const accessibleCustomers = projects
          .map(p => p.CustomerName)
          .filter((customer): customer is string => !!customer);

        filterCustomer.innerHTML = '<option value="">All Customers</option>' +
          [...new Set(accessibleCustomers)]
            .sort()
            .map(customer => `<option value="${customer}">${customer}</option>`)
            .join("");

        // Project Manager filter - all managers (no role filtering)
        let accessibleManagers: string[] = [];

        accessibleManagers = projects
          .map(p => getProjectManagerDisplayName(p.ProjectManager))
          .filter((manager): manager is string => !!manager);

        accessibleManagers = [...new Set(accessibleManagers)].sort();

        filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
          accessibleManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");

        // Project Reference filter - all references
        const accessibleReferences = projects
          .map(p => p.ProjectReference)
          .filter((ref): ref is string => !!ref);

        filterProjectReference.innerHTML = '<option value="">All References</option>' +
          [...new Set(accessibleReferences)]
            .sort()
            .map(ref => `<option value="${ref}">${ref}</option>`)
            .join("");

        console.log('✅ Filters initialized');
      };

      // Filter Data
      const filterData = (): void => {
        console.log('🔍 filterData called');

        const project = filterProject.value;
        const customer = filterCustomer.value;
        const projectManager = filterProjectManager.value;
        const projectReference = filterProjectReference.value;

        console.log('🎛️ Current filter values:', {
          project,
          customer,
          projectManager,
          projectReference
        });

        // Apply filters
        let filtered = dcsData.filter((item: any) => {
          const projectTitle = item[fieldMappings.Project]?.Title;
          const projectData = projects.find((p) => p.Title === projectTitle);

          if (project && projectTitle !== project) return false;
          if (customer && projectData?.CustomerName !== customer) return false;
          if (projectManager) {
            const managerName =
              typeof projectData?.ProjectManager === "object"
                ? projectData.ProjectManager?.Title
                : projectData?.ProjectManager;
            if (managerName !== projectManager) return false;
          }
          if (projectReference && projectData?.ProjectReference !== projectReference)
            return false;

          return true;
        });

        console.log('📊 After filters:', filtered.length, 'items');

        // Initialize pagination with filtered data
        initializePagination(filtered);

        console.log('✅ Filtering completed');
      };

      // Clear Filters
      const clearFilters = (): void => {
        console.log('🔄 Clearing filters...');

        filterProject.value = "";
        filterCustomer.value = "";
        filterProjectManager.value = "";
        filterProjectReference.value = "";

        // Reset to all data
        initializePagination(dcsData);
        toast("Filters cleared", "success");

        console.log('✅ Filters cleared');
      };

      // =================== MAIN TABLE COLUMN SORTING ===================
      // Value used to sort a record for the given column key
      const getSortValue = (item: any, key: string): string | number => {
        const projectTitle = item[fieldMappings.Project]?.Title || "";
        const projectData = projects.find((p) => p.Title === projectTitle);
        const summary = getDocumentStatusSummary(item);
        switch (key) {
          case "projectReference": return (projectData?.ProjectReference || "").toLowerCase();
          case "customer": return (projectData?.CustomerName || "").toLowerCase();
          case "project": return projectTitle.toLowerCase();
          case "projectManager": return getProjectManagerDisplayName(projectData?.ProjectManager).toLowerCase();
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
      const applySort = (data: any[]): any[] => {
        if (!sortKey) return data;
        const dir = sortDir === "asc" ? 1 : -1;
        return [...data].sort((a, b) => {
          const va = getSortValue(a, sortKey!);
          const vb = getSortValue(b, sortKey!);
          const aEmpty = va === "" || va === null || va === undefined;
          const bEmpty = vb === "" || vb === null || vb === undefined;
          if (aEmpty && bEmpty) return 0;
          if (aEmpty) return 1;
          if (bEmpty) return -1;
          if (va < vb) return -1 * dir;
          if (va > vb) return 1 * dir;
          return 0;
        });
      };

      // One delegated listener on the (static) thead - survives tbody re-renders
      const handleSortClick = (e: Event): void => {
        const th = (e.target as HTMLElement).closest('th[data-sort]') as HTMLElement | null;
        if (!th) return;
        const key = th.getAttribute('data-sort') || '';
        if (sortKey === key) {
          sortDir = sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          sortKey = key;
          sortDir = 'asc';
        }
        initializePagination(filteredDCSData); // re-sorts + resets to page 1
        updateSortIndicators();
      };

      const updateSortIndicators = (): void => {
        document.querySelectorAll('#dcsTable thead th[data-sort]').forEach((th) => {
          const el = th as HTMLElement;
          el.classList.remove('dcs-sorted-asc', 'dcs-sorted-desc');
          if (el.getAttribute('data-sort') === sortKey) {
            el.classList.add(sortDir === 'asc' ? 'dcs-sorted-asc' : 'dcs-sorted-desc');
          }
        });
      };

      // Format date for input (YYYY-MM-DD)
      const formatDateForInput = (dateString: string): string => {
        if (!dateString) return "";
        try {
          const date = new Date(dateString);
          return isNaN(date.getTime()) ? "" : date.toISOString().split("T")[0];
        } catch {
          return "";
        }
      };

      // Escape HTML to prevent XSS
      const escapeHtml = (unsafe: string | number): string => {
        if (unsafe === null || unsafe === undefined) return "";
        const safeString = String(unsafe);
        return safeString
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#039;");
      };

      // =================== MAIN TABLE DOCUMENT STATUS SUMMARY ===================
      interface DCSStatusSummary {
        total: number;
        notSubmitted: number;
        notApproved: number;
        revisedOrApprovedWithComments: number;
        approved: number;
        rejected: number;
        retainedForInformation: number;
      }

      // Normalise a status string for comparison (case/whitespace-insensitive)
      const normalizeStatus = (status: string | undefined | null): string =>
        String(status || "").trim().toLowerCase();

      // Counts DESIGN OUTPUT documents of a record by their current status.
      // Each main document is counted once (revisions are not counted separately).
      // Legacy statuses from older records are mapped to the nearest new bucket:
      //   Pending -> Not Submitted, Under Review -> Not Approved, Revised -> Approved with comments
      // "Document Deleted/cancelled" is included in Total only.
      const getDocumentStatusSummary = (item: any): DCSStatusSummary => {
        const summary: DCSStatusSummary = {
          total: 0,
          notSubmitted: 0,
          notApproved: 0,
          revisedOrApprovedWithComments: 0,
          approved: 0,
          rejected: 0,
          retainedForInformation: 0
        };

        try {
          const parsed = parseDCSData(item?.[fieldMappings.DCSData]);
          if (!parsed) return summary;

          const outputs: DCSDocumentItem[] = getOutputsFromProjectData(parsed) || [];

          outputs.forEach((doc: DCSDocumentItem) => {
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
        } catch (error) {
          console.warn("Failed to build status summary for record", item?.Id, error);
        }

        return summary;
      };

      // Renders a count cell - same look as the Procurement main table
      // (centered, bold, light tinted background per column)
      const renderCountCell = (count: number, extraClass: string = ""): string =>
        `<td class="dcs-count-cell ${extraClass}${count === 0 ? " dcs-count-zero" : ""}">${count}</td>`;

      const renderTable = (data: any[]): void => {
        const tableBody = document.getElementById("dcsTableBody") as HTMLElement;
        tableBody.innerHTML = "";

        if (data.length > 0) {
          data.forEach((item: any) => {
            const projectTitle = item[fieldMappings.Project]?.Title || "";
            const projectId = item[`${fieldMappings.Project}Id`] ||
              (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
            const projectData = projects.find((p) => p.Id === projectId);

            const row = document.createElement("tr");

            // Get display names
            const projectManagerDisplay = getProjectManagerDisplayName(projectData?.ProjectManager);

            // Document status counts (from the record's DCS JSON - output documents)
            const summary = getDocumentStatusSummary(item);

            // Project Reference - Edit link ONLY when this role may edit that project
            // (Admin: all projects; GH Head / PM Manager / Engg Lead: their own
            // assigned projects). Other roles see the reference as plain text and
            // can still use the View button.
            const projectRefText = escapeHtml(projectData?.ProjectReference || "") || '<span style="color:#6c757d;">-</span>';
            const canEditThis = canEditProject(projectData);
            const projectRefCell = canEditThis
              ? `<a href="#" class="dcs-edit-link dcs-edit-btn" data-id="${item.Id}" title="Edit document data">${projectRefText}</a>`
              : `<span title="${escapeHtml(projectData?.ProjectReference || "")}">${projectRefText}</span>`;

            row.innerHTML = `
              <td class="dcs-project-ref-cell dcs-sticky-main-col dcs-main-col-1" title="${escapeHtml(projectData?.ProjectReference || "")}">${projectRefCell}</td>
              <td class="dcs-sticky-main-col dcs-main-col-2" title="${escapeHtml(projectData?.CustomerName || "")}">${escapeHtml(projectData?.CustomerName || "")}</td>
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
        } else {
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
          btn.addEventListener("click", (e: Event) => {
            e.preventDefault();
            editData(Number((btn as HTMLElement).dataset.id));
          });
        });

        document.querySelectorAll(".dcs-view-btn").forEach((btn) => {
          btn.addEventListener("click", () =>
            viewData(Number((btn as HTMLElement).dataset.id))
          );
        });
      };

      // Get inputs / outputs from project data
      const getInputsFromProjectData = (data: AnyDCSProjectData | null): DCSDocumentItem[] => {
        if (!data) return [];

        if ('documentStructure' in data && data.documentStructure) {
          return data.documentStructure.inputs || [];
        }

        if ('inputs' in data) {
          return data.inputs || [];
        }

        return [];
      };

      const getOutputsFromProjectData = (data: AnyDCSProjectData | null): DCSDocumentItem[] => {
        if (!data) return [];

        if ('documentStructure' in data && data.documentStructure) {
          return data.documentStructure.outputs || [];
        }

        if ('outputs' in data) {
          return data.outputs || [];
        }

        return [];
      };

      const convertToNewFormat = (data: DCSProjectData | LegacyDCSProjectData): DCSProjectData => {
        if ('documentStructure' in data) {
          return data as DCSProjectData;
        }

        // Convert from old format to new format
        const legacyData = data as LegacyDCSProjectData;
        return {
          projectId: legacyData.projectId,
          documentStructure: {
            inputs: legacyData.inputs || [],
            outputs: legacyData.outputs || []
          }
        };
      };

      const editData = (id: number): void => {
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
        // Project cannot be changed while editing an existing record (like Procurement)
        projectSelect.disabled = true;

        // CRITICAL: Set the editing ID
        editingId = id;

        const projectId = record[`${fieldMappings.Project}Id`] ||
          (record[fieldMappings.Project] && record[fieldMappings.Project].Id);

        // Title: "Edit Document Control Data - <Project> (<Ref>)"
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
          // Use helper functions to get inputs and outputs
          const inputs = getInputsFromProjectData(parsedData);
          const outputs = getOutputsFromProjectData(parsedData);

          // Clear tables
          inputTableBody.innerHTML = '';
          outputTableBody.innerHTML = '';

          // Populate input documents
          if (inputs.length > 0) {
            inputs.forEach((doc: DCSDocumentItem, index: number) => {
              const newRow = createEditableDocumentRow(doc, index + 1, true);
              inputTableBody.appendChild(newRow);

              // Add revision rows if they exist (8 cells: 7 fields + actions)
              if (doc.revisions && doc.revisions.length > 0) {
                doc.revisions.forEach((revision: RevisionItem, revIndex: number) => {
                  const revisionHTML = buildRevisionRowHTML(doc.docId, revision);
                  inputTableBody.insertAdjacentHTML('beforeend', revisionHTML);
                });

                // Update merged cells
                updateMergedCells(doc.docId);
              }
            });
          } else {
            // Add one default input document
            const defaultInputDoc: DCSDocumentItem = {
              docId: `default_input_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
              docNo: "DOC-001",
              description: "",
              category: "A",
              enggDivision: "",
              revNo: 0,
              originalPlan: "",
              revisedPlan: "",
              actualSubmission: "",
              submissionTransmittal: "",
              approvalDate: "",
              approvalTransmittal: "",
              status: "Pending",
              remarks: "",
              isInput: true,
              revisions: [],
              docType: 'input'
            };
            const newRow = createEditableDocumentRow(defaultInputDoc, 1, true);
            inputTableBody.appendChild(newRow);
          }

          // Populate output documents
          if (outputs.length > 0) {
            outputs.forEach((doc: DCSDocumentItem, index: number) => {
              const newRow = createEditableDocumentRow(doc, index + 1, false);
              outputTableBody.appendChild(newRow);

              // Add revision rows if they exist (8 cells: 7 fields + actions)
              if (doc.revisions && doc.revisions.length > 0) {
                doc.revisions.forEach((revision: RevisionItem, revIndex: number) => {
                  const revisionHTML = buildRevisionRowHTML(doc.docId, revision);
                  outputTableBody.insertAdjacentHTML('beforeend', revisionHTML);
                });

                // Update merged cells
                updateMergedCells(doc.docId);
              }
            });
          } else {
            // Add one default output document
            const defaultOutputDoc: DCSDocumentItem = {
              docId: `default_output_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
              docNo: "DOC-OUT-001",
              description: "",
              category: "A",
              enggDivision: "",
              revNo: 0,
              originalPlan: "",
              revisedPlan: "",
              actualSubmission: "",
              submissionTransmittal: "",
              approvalDate: "",
              approvalTransmittal: "",
              status: "Not Submitted",
              remarks: "",
              isInput: false,
              revisions: [],
              docType: 'output'
            };
            const newRow = createEditableDocumentRow(defaultOutputDoc, 1, false);
            outputTableBody.appendChild(newRow);
          }
        } else {
          resetDocumentTables();
        }

        // Make sure buttons are visible
        if (addInputDocBtn) addInputDocBtn.style.display = "inline-block";
        if (addOutputDocBtn) addOutputDocBtn.style.display = "inline-block";

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
        // Update form actions for edit mode
        updateFormActionsForEditMode();

        // Re-attach document event listeners
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
      const buildRevisionRowHTML = (docId: string, revision: RevisionItem): string => `
        <tr class="dcs-document-group" data-doc-id="${docId}">
          <td class="dcs-rev-cell">${revision.revNumber}</td>
          <td><input type="date" class="dcs-form-control dcs-date-input"
                 value="${formatDateForInput(revision.originalPlan)}"
                 placeholder="Original Plan"></td>
          <td><input type="date" class="dcs-form-control dcs-date-input"
                 value="${formatDateForInput(revision.revisedPlan)}"
                 placeholder="Revised Plan"></td>
          <td><input type="date" class="dcs-form-control dcs-date-input"
                 value="${formatDateForInput(revision.actualSubmission)}"
                 placeholder="Actual Submission"></td>
          <td>
            <textarea class="dcs-form-control dcs-transmittal-input" rows="2"
                  placeholder="Transmittal No">${escapeHtml(revision.submissionTransmittal)}</textarea>
          </td>
          <td><input type="date" class="dcs-form-control dcs-date-input"
                 value="${formatDateForInput(revision.approvalDate)}"
                 placeholder="Approval Date"></td>
          <td>
            <textarea class="dcs-form-control dcs-transmittal-input" rows="2"
                  placeholder="Transmittal No">${escapeHtml(revision.approvalTransmittal)}</textarea>
          </td>
          <td class="dcs-doc-actions">
            <button type="button" class="dcs-btn-danger dcs-remove-revision-btn"
                    title="Remove Revision">➖</button>
          </td>
        </tr>
      `;

      const updateFormActionsForEditMode = (): void => {
        const formActions = document.querySelector('.dcs-form-actions');
        if (!formActions) return;

        formActions.innerHTML = `
  <div style="display: flex; gap: 10px; flex-wrap: wrap;">
    <button type="submit" form="dcsForm" class="dcs-btn-primary">
      💾 Save Changes
    </button>
    <button type="button" class="dcs-btn-secondary dcs-back-to-table">
      ← Cancel
    </button>
  </div>
`;

// ✅ Sync the top button with the new bottom button
getSaveButtons().forEach(btn => {
  btn.style.display = "inline-flex";
});

        // Reattach back button listener
        const backBtn = formActions.querySelector('.dcs-back-to-table');
        if (backBtn) {
          backBtn.addEventListener('click', handleBackToTable);
        }
      };

      const createEditableDocumentRow = (doc: DCSDocumentItem, srNo: number, isInput: boolean): HTMLTableRowElement => {
        const docId = doc.docId || `doc_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        const prefix = isInput ? 'DOC' : 'DOC-OUT';
        const docNo = doc.docNo || `${prefix}-${String(srNo).padStart(3, '0')}`;

        // Set correct transmittal placeholders based on document type
        const submissionPlaceholder = isInput ? 'Transmittal No' : 'Transmittal No';
        const approvalPlaceholder = isInput ? 'Transmittal No' : 'Transmittal No';

        const row = document.createElement("tr");
        row.className = "dcs-document-group dcs-document-header";
        row.setAttribute("data-doc-id", docId);

        const totalRevRows = (doc.revisions?.length || 0) + 1;
        row.setAttribute("data-rev-rows", totalRevRows.toString());

        const tableType: 'input' | 'output' = isInput ? 'input' : 'output';

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

      const viewData = (id: number): void => {
        // 🔐 ROLE GATE: view mode needs read access (any non-NONE role)
        if (!canRead()) {
          toast("You do not have permission to view document data", "error");
          return;
        }

        const record = dcsData.find((d) => d.Id === id);
        if (!record) {
          toast("Record not found", "error");
          return;
        }

        // Store the project ID in a data attribute for later use
        const projectId = record[`${fieldMappings.Project}Id`] ||
          (record[fieldMappings.Project] && record[fieldMappings.Project].Id);

        // Set view mode class - add this to CSS if not exists
        formSection.classList.add("dcs-view-mode");
        formTitle.textContent = "View Document Control Data";
        formSection.classList.remove("dcs-edit-mode");

        // Show form section, hide table
        formSection.style.display = "block";
        tableSection.style.display = "none";

        // CRITICAL: Set project select value BEFORE disabling
        if (projectId) {
          projectSelect.value = projectId.toString();
        }

        // Now disable project selection
        projectSelect.disabled = true;
        existingDataAlert.style.display = "none";
        importButtonContainer.style.display = "none";

        // Hide submit button
        getSaveButtons().forEach(btn => {
  btn.style.display = "none";
});

        // Clear tables
        inputTableBody.innerHTML = '';
        outputTableBody.innerHTML = '';

        // Update project details with the actual data
        if (projectId) {
          const project = projects.find(p => p.Id === projectId);
          if (project) {
            setProjectDetailsDisplay(project);
          } else {
            console.error("❌ Project not found in projects array for ID:", projectId);
            toast("Project details not found", "error");
          }
        } else {
          console.error("❌ No project ID found in record");
          toast("Project information missing", "error");
        }

        // Parse and load data into readonly fields
        const parsedData = parseDCSData(record[fieldMappings.DCSData]);

        if (parsedData) {
          containerTable.style.display = "block";

          // Load data into readonly form fields
          loadDataIntoReadonlyForm(parsedData);
        } else {
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
      const createReadonlyDocumentRow = (doc: DCSDocumentItem, srNo: number, tableType: 'input' | 'output'): HTMLTableRowElement => {
        const row = document.createElement("tr");
        row.className = "dcs-document-group dcs-document-header";
        row.setAttribute("data-doc-id", doc.docId);

        let scheduleBgColor: string;
        let submissionBgColor: string;
        let approvalBgColor: string;

        if (tableType === 'input') {
          scheduleBgColor = '#e8f5e9'; // Green for input schedule
          submissionBgColor = '#fff3cd'; // Yellow for submission
          approvalBgColor = '#d4edda'; // Green for approval
        } else {
          scheduleBgColor = '#e3f2fd'; // Blue for output schedule
          submissionBgColor = '#fff3cd'; // Yellow for submission
          approvalBgColor = '#d4edda'; // Green for approval
        }

        // Helper to format dates for display
        const formatDateForDisplay = (dateString: string): string => {
          if (!dateString) return "";
          try {
            const date = new Date(dateString);
            return isNaN(date.getTime()) ? dateString : date.toLocaleDateString("en-GB");
          } catch {
            return dateString;
          }
        };

        // Status badge styling
        const getStatusBadgeStyle = (status: string): string => {
          const statusStyles: { [key: string]: string } = {
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
      const createReadonlyRevisionRow = (revision: RevisionItem, doc: DCSDocumentItem, tableType: 'input' | 'output'): HTMLTableRowElement => {
        const revRow = document.createElement("tr");
        revRow.className = "dcs-document-group dcs-revision-row";

        let scheduleBgColor: string;
        let submissionBgColor: string;
        let approvalBgColor: string;

        if (tableType === 'input') {
          scheduleBgColor = '#e8f5e9';
          submissionBgColor = '#fff3cd';
          approvalBgColor = '#d4edda';
        } else {
          scheduleBgColor = '#e3f2fd';
          submissionBgColor = '#fff3cd';
          approvalBgColor = '#d4edda';
        }

        // Helper to format dates for display
        const formatDateForDisplay = (dateString: string): string => {
          if (!dateString) return "";
          try {
            const date = new Date(dateString);
            return isNaN(date.getTime()) ? dateString : date.toLocaleDateString("en-GB");
          } catch {
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
      const loadDataIntoReadonlyForm = (projectData: AnyDCSProjectData): void => {
        const inputs = getInputsFromProjectData(projectData);
        const outputs = getOutputsFromProjectData(projectData);

        // Load input documents
        if (inputs.length > 0) {
          inputs.forEach((doc: DCSDocumentItem, index: number) => {
            const docRow = createReadonlyDocumentRow(doc, index + 1, 'input');
            inputTableBody.appendChild(docRow);

            // Add revision rows if they exist
            if (doc.revisions && doc.revisions.length > 0) {
              doc.revisions.forEach((revision: RevisionItem, revIndex: number) => {
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
        } else {
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
          outputs.forEach((doc: DCSDocumentItem, index: number) => {
            const docRow = createReadonlyDocumentRow(doc, index + 1, 'output');
            outputTableBody.appendChild(docRow);

            // Add revision rows if they exist
            if (doc.revisions && doc.revisions.length > 0) {
              doc.revisions.forEach((revision: RevisionItem, revIndex: number) => {
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
        } else {
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
      const setupViewModeActions = (projectId: number | undefined): void => {
        const formActions = document.querySelector('.dcs-form-actions');
        if (!formActions) return;

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
          backBtn.addEventListener('click', handleBackToTable);
        }

        // Export button - export current project only
        const exportBtn = document.getElementById('viewModeExportBtn');
        if (exportBtn) {
          exportBtn.addEventListener('click', exportAllDocuments);
        }
      };

      const exportAllDocuments = async (): Promise<void> => {
        showExportProgress("Preparing document data for export...");

        try {
          // Get the current project ID from view mode
          const projectSelect = document.getElementById("dcsProjectSelect") as HTMLSelectElement;
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

          const record = dcsData.find((item: any) => {
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
          const coverSheetData: any[][] = [];

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
              if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
              coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "FFFFFF" } };
            }
          }
          for (let R = 0; R <= fullDataRange.e.r; ++R) {
            for (let C = 0; C <= 10; ++C) {
              const address = XLSX.utils.encode_cell({ r: R, c: C });
              if (!coverSheetWs[address]) {
                coverSheetWs[address] = { t: 's', v: '' };
              }
              if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
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
            if (!coverSheetWs[address]) coverSheetWs[address] = { t: 's', v: '' };
            if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
            coverSheetWs[address].s.border = {
              ...coverSheetWs[address].s.border,
              top: darkGrayBorder.top
            };
          }

          for (let C = 0; C <= lastCol; C++) {
            const address = XLSX.utils.encode_cell({ r: lastRow, c: C });
            if (!coverSheetWs[address]) coverSheetWs[address] = { t: 's', v: '' };
            if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
            coverSheetWs[address].s.border = {
              ...coverSheetWs[address].s.border,
              bottom: darkGrayBorder.bottom
            };
          }

          for (let R = 0; R <= lastRow; R++) {
            const address = XLSX.utils.encode_cell({ r: R, c: 0 });
            if (!coverSheetWs[address]) coverSheetWs[address] = { t: 's', v: '' };
            if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
            coverSheetWs[address].s.border = {
              ...coverSheetWs[address].s.border,
              left: darkGrayBorder.left
            };
          }

          for (let R = 0; R <= lastRow; R++) {
            const address = XLSX.utils.encode_cell({ r: R, c: lastCol });
            if (!coverSheetWs[address]) coverSheetWs[address] = { t: 's', v: '' };
            if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
            coverSheetWs[address].s.border = {
              ...coverSheetWs[address].s.border,
              right: darkGrayBorder.right
            };
          }

          const corners = [
            { r: 0, c: 0 }, // Top-left
            { r: 0, c: lastCol }, // Top-right
            { r: lastRow, c: 0 }, // Bottom-left
            { r: lastRow, c: lastCol } // Bottom-right
          ];

          corners.forEach(({ r, c }) => {
            const address = XLSX.utils.encode_cell({ r, c });
            if (!coverSheetWs[address]) coverSheetWs[address] = { t: 's', v: '' };
            if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
            coverSheetWs[address].s.border = darkGrayBorder;
          });

          for (let R = 4; R < 13; R += 2) {
            const dateAddressC = XLSX.utils.encode_cell({ r: R, c: 2 });
            if (coverSheetWs[dateAddressC]) {
              if (!coverSheetWs[dateAddressC].s) coverSheetWs[dateAddressC].s = {};
              coverSheetWs[dateAddressC].s.border = borderStyle;
            }
            const dateAddressF = XLSX.utils.encode_cell({ r: R, c: 5 });
            if (coverSheetWs[dateAddressF]) {
              if (!coverSheetWs[dateAddressF].s) coverSheetWs[dateAddressF].s = {};
              coverSheetWs[dateAddressF].s.border = borderStyle;
            }
            const dateAddressI = XLSX.utils.encode_cell({ r: R, c: 8 });
            if (coverSheetWs[dateAddressI]) {
              if (!coverSheetWs[dateAddressI].s) coverSheetWs[dateAddressI].s = {};
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
                if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
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
              if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
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
                if (!coverSheetWs[address].s) coverSheetWs[address].s = {};
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
                if (isTopEdge) borderStyle.top = { style: "thin" };
                if (isBottomEdge) borderStyle.bottom = { style: "thin" };
                if (isLeftEdge) borderStyle.left = { style: "thin" };
                if (isRightEdge) borderStyle.right = { style: "thin" };

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

        } catch (error: any) {
          hideExportProgress();
          console.error("Export error:", error);
          toast("Failed to export Excel report", "error");
        }
      };

      // Helper function to create DCS Document sheet (extracted from previous code)
      const createDCSDocumentSheet = async (
        outputs: DCSDocumentItem[],
        inputs: DCSDocumentItem[],
        project: any
      ): Promise<any> => {

        // =================== CORRECTED HEADER STRUCTURE ===================
        const headerRow1: string[] = [
          "Sr No", "Document No", "Description", "Category", "Engg Division", "Rev No",
          "Document Schedule", "Document Schedule",
          "Document Submission", "Document Submission",
          "Document Approval", "Document Approval",
          "Status", "Remarks"
        ];

        const headerRow2: string[] = [
          "", "", "", "", "", "",
          "Original Plan", "Revised Plan",
          "Actual Submission", "Transmittal No",
          "Date", "Transmittal No",
          "", ""
        ];

        // Create worksheet with two-level headers
        const ws = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2]);

        // Prepare data rows
        const dataRows: any[][] = [];
        const merges: XLSX.Range[] = [];
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

          outputs.forEach((doc: DCSDocumentItem) => {
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
              doc.revisions.forEach((revision: RevisionItem) => {
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

          inputs.forEach((doc: DCSDocumentItem) => {
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
              doc.revisions.forEach((revision: RevisionItem) => {
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
        merges.push(
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
        );

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
            if (!ws[address]) continue;

            if (!ws[address].s) ws[address].s = {};

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
              let sectionType: "OUTPUT" | "INPUT" | null = null;
              for (let checkR = R; checkR >= 2; checkR--) {
                const checkAddr = XLSX.utils.encode_cell({ r: checkR, c: 0 });
                if (ws[checkAddr]) {
                  const titleValue = ws[checkAddr].v;
                  const titleValueStr = typeof titleValue === 'string' ? titleValue : String(titleValue || '');
                  if (titleValueStr === "DESIGN OUTPUT - FROM ABB") {
                    sectionType = "OUTPUT";
                    break;
                  } else if (titleValueStr === "DESIGN INPUT FROM CONSULTANT / CUSTOMER") {
                    sectionType = "INPUT";
                    break;
                  }
                }
              }

              if (sectionType) {
                if (C === 6 || C === 7) {
                  const scheduleColor = sectionType === "OUTPUT" ? outputScheduleColor : inputScheduleColor;
                  ws[address].s.fill = { patternType: "solid", fgColor: { rgb: scheduleColor } };
                } else if (C === 8 || C === 9) {
                  ws[address].s.fill = { patternType: "solid", fgColor: { rgb: submissionColor } };
                } else if (C === 10 || C === 11) {
                  ws[address].s.fill = { patternType: "solid", fgColor: { rgb: approvalColor } };
                }
              }

              if (C === 0 || C === 5 || C === 12) {
                ws[address].s.alignment = { horizontal: "center", vertical: "center" };
              } else {
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

              const isTopRow = R === 3 || (R > 3 && (
                prevCellValueStr.includes("DESIGN OUTPUT - FROM ABB") ||
                prevCellValueStr.includes("DESIGN INPUT FROM CONSULTANT / CUSTOMER") ||
                prevCellValueStr === ""
              ));

              const isBottomRow = R === totalRows || (R < totalRows && (
                nextCellValueStr.includes("DESIGN OUTPUT - FROM ABB") ||
                nextCellValueStr.includes("DESIGN INPUT FROM CONSULTANT / CUSTOMER") ||
                nextCellValueStr === ""
              ));

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
          { wch: 8 },   // A - Sr No
          { wch: 20 },  // B - Document No
          { wch: 35 },  // C - Description
          { wch: 12 },  // D - Category
          { wch: 18 },  // E - Engg Division
          { wch: 10 },  // F - Rev No
          { wch: 15 },  // G - Original Plan
          { wch: 15 },  // H - Revised Plan
          { wch: 18 },  // I - Actual Submission
          { wch: 20 },  // J - Transmittal No
          { wch: 15 },  // K - Approval Date
          { wch: 20 },  // L - Transmittal No
          { wch: 20 },  // M - Status
          { wch: 25 }   // N - Remarks
        ];

        return ws;
      };

      // Format date for Excel export (YYYY-MM-DD)
      const formatDateForExcel = (dateString: string): string => {
        if (!dateString) return "";
        try {
          const date = new Date(dateString);
          if (isNaN(date.getTime())) return "";
          const year = date.getFullYear();
          const month = String(date.getMonth() + 1).padStart(2, '0');
          const day = String(date.getDate()).padStart(2, '0');
          return `${year}-${month}-${day}`;
        } catch {
          return "";
        }
      };

      // Show export progress
      const showExportProgress = (message: string): void => {
        const existing = document.querySelector('.dcs-export-progress');
        if (existing) existing.remove();

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
      const hideExportProgress = (): void => {
        const existing = document.querySelector('.dcs-export-progress');
        if (existing) existing.remove();
      };

      // Project Change Handler
      const onProjectChange = (selectedId: number, isEditing: boolean = false): void => {
        updateProjectDetails(selectedId);

        // Show form section and hide table
        formSection.style.display = "block";
        tableSection.style.display = "none";
        containerTable.style.display = "block";

        // Show import instructions
        const importInstructions = document.getElementById('dcsImportInstructions');
        if (importInstructions) {
          importInstructions.style.display = 'block';
        }

        // Check for existing data
        const existingRecord = dcsData.find(
          (item) =>
            item[`${fieldMappings.Project}Id`] === selectedId ||
            (item[fieldMappings.Project] &&
              item[fieldMappings.Project].Id === selectedId)
        );

        if (existingRecord && !isEditing) {
          existingDataAlert.style.display = "block";
          setEditTitle(projects.find(p => p.Id === selectedId) || null);
          importButtonContainer.style.display = "flex";

          const dcsProjectData = parseDCSData(existingRecord[fieldMappings.DCSData]);
          if (dcsProjectData) {
            loadDocumentsFromJSON(dcsProjectData);
          } else {
            resetDocumentTables();
          }
        } else {
          existingDataAlert.style.display = "none";
          formTitle.textContent = "Add Project Documents";
          formSection.classList.remove("dcs-edit-mode");

          importButtonContainer.style.display = "flex";

          resetDocumentTables();
        }

        // Re-attach document event listeners
        setTimeout(() => {
          attachDocumentEventListeners();
        }, 100);
      };

      const addRevisionToDocument = (docId: string): void => {
        const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`) as HTMLElement;
        if (!docHeader) {
          console.error(`Document header not found for ID: ${docId}`);
          return;
        }

        const tableBody = docHeader.closest('tbody');
        if (!tableBody) return;

        const existingRevRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);
        const newRevNumber = existingRevRows.length + 1;

        const currentRevRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');
        const newRevRows = currentRevRows + 1;

        docHeader.setAttribute('data-rev-rows', newRevRows.toString());

        // ✅ Exactly 8 cells (7 fields + actions) - aligned with the 15-column grid
        const newRevisionHTML = `
          <tr class="dcs-document-group" data-doc-id="${docId}" data-is-revision="true">
            <td class="dcs-rev-cell">${newRevNumber}</td>
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
            <td class="dcs-doc-actions">
              <button type="button" class="dcs-btn-danger dcs-remove-revision-btn" title="Remove Revision">➖</button>
            </td>
          </tr>
        `;

        let lastRow: Element | null = null;
        if (existingRevRows.length > 0) {
          lastRow = existingRevRows[existingRevRows.length - 1];
        } else {
          lastRow = docHeader;
        }

        lastRow.insertAdjacentHTML('afterend', newRevisionHTML);

        updateMergedCellsForDocument(docId);

        const newRevision = lastRow.nextElementSibling;
        if (newRevision) {
          const removeBtn = newRevision.querySelector('.dcs-remove-revision-btn');
          if (removeBtn) {
            removeBtn.addEventListener('click', (event: Event) => {
              event.preventDefault();
              event.stopPropagation();
              const button = event.currentTarget as HTMLButtonElement;
              removeRevisionFromDocument(docId, button);
            });
          }
        }
      };

      // Enhanced function to update merged cells - FIXED
      // ✅ Merge Sr No, Doc No, Desc, Category, Division, Status, Remarks only.
      // The Actions column (14) stays UNMERGED so revision rows (which carry the
      // per-revision remove button) align with the grid.
      const updateMergedCellsForDocument = (docId: string): void => {
        const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
        if (!docHeader) return;

        const revRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');

        const allCells = docHeader.querySelectorAll('td');
        allCells.forEach(cell => {
          cell.removeAttribute('rowspan');
          cell.classList.remove('dcs-merged-cell');
        });

        const cellsToMerge = [0, 1, 2, 3, 4, 12, 13];

        cellsToMerge.forEach(cellIndex => {
          const cell = docHeader.querySelector(`td:nth-child(${cellIndex + 1})`);
          if (cell) {
            cell.setAttribute('rowspan', revRows.toString());
            cell.classList.add('dcs-merged-cell');
          }
        });

        const revNoCell = docHeader.querySelector('.dcs-rev-cell');
        if (revNoCell) {
          revNoCell.removeAttribute('rowspan');
          revNoCell.classList.remove('dcs-merged-cell');
          revNoCell.textContent = '0';
        }
      };

      const updateMergedCells = (docId: string): void => {
        const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
        if (!docHeader) return;

        const revRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');

        const cells = docHeader.querySelectorAll('td');

        // ✅ Actions column (14) stays unmerged - matches 8-cell revision rows
        const cellsToMerge = [0, 1, 2, 3, 4, 12, 13];

        cellsToMerge.forEach(cellIndex => {
          if (cellIndex < cells.length) {
            const cell = cells[cellIndex];
            cell.setAttribute('rowspan', revRows.toString());
            cell.classList.add('dcs-merged-cell');
          }
        });
      };

      const removeRevisionFromDocument = (docId: string, buttonElement: HTMLButtonElement): void => {
        const revisionRow = buttonElement.closest('tr[data-doc-id]') as HTMLElement;
        if (!revisionRow) return;

        const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
        if (!docHeader) return;

        const allRevRows = document.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);

        if (allRevRows.length <= 1) {
          revisionRow.remove();
          docHeader.setAttribute('data-rev-rows', '1');
          updateMergedCellsForDocument(docId);
        } else {
          revisionRow.remove();

          const currentRevRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');
          docHeader.setAttribute('data-rev-rows', (currentRevRows - 1).toString());

          updateRevisionNumbers(docId);

          updateMergedCellsForDocument(docId);
        }
      };

      const updateRevisionNumbers = (docId: string): void => {
        const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
        if (!docHeader) return;

        const tableBody = docHeader.closest('tbody');
        if (!tableBody) return;

        const revisionRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);

        revisionRows.forEach((row, index) => {
          const revCell = row.querySelector('.dcs-rev-cell');
          if (revCell) {
            revCell.textContent = (index + 1).toString();
          }
        });
      };

      const handleAddRevision = (e: Event): void => {
        e.preventDefault();
        e.stopPropagation();

        const button = e.target as HTMLElement;
        const addRevBtn = button.closest('.dcs-add-revision-btn');

        if (addRevBtn) {
          const docHeader = addRevBtn.closest('.dcs-document-header') as HTMLElement;
          const docId = docHeader?.getAttribute('data-doc-id');
          if (docId) {
            addRevisionToDocument(docId);
          }
        }
      };

      const attachDocumentEventListeners = (): void => {
        if (addInputDocBtn) {
          addInputDocBtn.removeEventListener("click", handleAddInputDoc);
          addInputDocBtn.addEventListener("click", handleAddInputDoc);
        }

        if (addOutputDocBtn) {
          addOutputDocBtn.removeEventListener("click", handleAddOutputDoc);
          addOutputDocBtn.addEventListener("click", handleAddOutputDoc);
        }

        document.querySelectorAll('.dcs-add-revision-btn').forEach(btn => {
          btn.removeEventListener('click', handleAddRevision);
          btn.addEventListener('click', handleAddRevision);
        });

        document.querySelectorAll('.dcs-remove-doc').forEach(btn => {
          btn.removeEventListener('click', handleRemoveDocument);
          btn.addEventListener('click', handleRemoveDocument);
        });

        document.querySelectorAll('.dcs-remove-revision-btn').forEach(btn => {
          const doRemove = (e: Event) => {
            e.preventDefault();
            e.stopPropagation();
            const docHeader = btn.closest('tr[data-doc-id]');
            const docId = docHeader?.getAttribute('data-doc-id');
            if (docId) {
              removeRevisionFromDocument(docId, btn as HTMLButtonElement);
            }
          };
          btn.removeEventListener('click', doRemove);
          btn.addEventListener('click', doRemove);
        });
      };

      const handleAddInputDoc = (e?: Event): void => {
        if (e) {
          e.preventDefault();
          e.stopPropagation();
        }
        addNewDocumentToTable('input');
      };

      const handleAddOutputDoc = (e?: Event): void => {
        if (e) {
          e.preventDefault();
          e.stopPropagation();
        }
        addNewDocumentToTable('output');
      };

      const styleElement = document.createElement('style');
      styleElement.textContent = `
        .dcs-doc-actions {
          display: flex !important;
          gap: 5px !important;
          justify-content: center !important;
          visibility: visible !important;
          opacity: 1 !important;
        }
        .dcs-doc-actions button {
          display: inline-flex !important;
          align-items: center !important;
          padding: 6px 12px !important;
          visibility: visible !important;
          opacity: 1 !important;
        }
        .dcs-add-revision-btn, .dcs-remove-doc {
          min-width: 80px !important;
          height: 36px !important;
        }
      `;
      document.head.appendChild(styleElement);

      const ensureDocumentButtonsVisible = (): void => {
        const docActionCells = document.querySelectorAll('.dcs-doc-actions');
        docActionCells.forEach(cell => {
          if (cell instanceof HTMLElement) {
            cell.style.display = 'flex';
            cell.style.visibility = 'visible';
            cell.style.opacity = '1';
          }
        });

        const buttons = document.querySelectorAll('.dcs-doc-actions button');
        buttons.forEach(button => {
          if (button instanceof HTMLElement) {
            button.style.display = 'inline-flex';
            button.style.visibility = 'visible';
            button.style.opacity = '1';
          }
        });
      };

      const handleRemoveDocument = (e: Event): void => {
        e.preventDefault();
        e.stopPropagation();

        const button = e.target as HTMLElement;
        const removeBtn = button.closest('.dcs-remove-doc');

        if (removeBtn) {
          const docHeader = removeBtn.closest('.dcs-document-header') as HTMLElement;
          if (docHeader) {
            const docId = docHeader.getAttribute('data-doc-id');
            const isDefault = docHeader.getAttribute('data-is-default') === 'true';

            if (isDefault) {
              toast('Default document cannot be removed. Please clear the fields instead.', 'warning');
              return;
            }

            if (confirm('Remove this document and all its revisions?')) {
              const tableBody = docHeader.closest('tbody');
              if (!tableBody) return;

              const allRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]`);

              allRows.forEach(row => row.remove());

              updateAllSerialNumbers();

              attachDocumentEventListeners();
            }
          }
        }
      };

      const addNewDocumentToTable = (tableType: 'input' | 'output'): void => {
        const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
        if (!tableBody) {
          console.error(`Table body not found for ${tableType}`);
          return;
        }

        const existingDocs = tableBody.querySelectorAll('.dcs-document-header');
        const nextSerialNo = existingDocs.length + 1;

        const docId = `doc_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

        const newDocumentHTML = createDocumentRowHTML(tableType, nextSerialNo, docId);

        tableBody.insertAdjacentHTML('beforeend', newDocumentHTML);

        const newRow = tableBody.querySelector(`tr[data-doc-id="${docId}"]`);

        if (newRow) {
          updateAllSerialNumbers();
          attachEventListenersToRow(newRow as HTMLElement);
        }
      };

      const createDocumentRowHTML = (tableType: 'input' | 'output', srNo: number, docId: string): string => {
        const prefix = tableType === 'input' ? 'DOC' : 'DOC-OUT';
        const docNo = `${prefix}-${String(srNo).padStart(3, '0')}`;

        const submissionPlaceholder = tableType === 'input' ? 'Transmittal No' : 'Transmittal No';
        const approvalPlaceholder = tableType === 'input' ? 'Transmittal No' : 'Transmittal No';

        return `
          <tr class="dcs-document-group dcs-document-header" data-doc-id="${docId}" data-rev-rows="1">
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
              <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${submissionPlaceholder}"></textarea>
            </td>
            <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Approval Date"></td>
            <td>
              <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${approvalPlaceholder}"></textarea>
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

      // Helper function to attach event listeners to a specific row
      const attachEventListenersToRow = (row: HTMLElement): void => {
        const addRevisionBtn = row.querySelector('.dcs-add-revision-btn');
        if (addRevisionBtn) {
          addRevisionBtn.addEventListener('click', handleAddRevision);
        }

        const removeDocBtn = row.querySelector('.dcs-remove-doc');
        if (removeDocBtn) {
          removeDocBtn.addEventListener('click', handleRemoveDocument);
        }
      };

      const exportToExcel = async (): Promise<void> => {
        try {
          // Get filtered data based on current filters
          const project = filterProject.value;
          const customer = filterCustomer.value;
          const projectManager = filterProjectManager.value;
          const projectReference = filterProjectReference.value;

          const filteredData = dcsData.filter((item) => {
            const projectTitle = item[fieldMappings.Project]?.Title || "";
            const projectData = projects.find((p) => p.Title === projectTitle);

            if (project && projectTitle !== project) return false;
            if (customer && projectData?.CustomerName !== customer) return false;
            if (projectManager) {
              const managerName =
                typeof projectData?.ProjectManager === "object"
                  ? projectData.ProjectManager?.Title
                  : projectData?.ProjectManager;
              if (managerName !== projectManager) return false;
            }
            if (projectReference && projectData?.ProjectReference !== projectReference)
              return false;

            return true;
          });

          if (filteredData.length === 0) {
            hideExportProgress();
            toast("No data to export", "warning");
            return;
          }

          // Create header structure
          const headerRow = [
            "Project Reference",
            "Project Name",
            "Customer",
            "Project Manager",
            "Document No",
            "Description",
            "Category",
            "Engg Division",
            "Rev No",
            "Original Plan",
            "Revised Plan",
            "Actual Submission",
            "Submission Transmittal",
            "Approval Date",
            "Approval Transmittal",
            "Status",
            "Remarks",
            "Document Type"
          ];

          const wsData = [headerRow];

          // Build data rows
          filteredData.forEach((item) => {
            const projectTitle = item[fieldMappings.Project]?.Title || "";
            const projectData = projects.find((p) => p.Title === projectTitle);
            const parsedData = parseDCSData(item[fieldMappings.DCSData]);

            if (parsedData) {
              const inputs = getInputsFromProjectData(parsedData);
              const outputs = getOutputsFromProjectData(parsedData);

              // Input documents
              inputs.forEach((doc: DCSDocumentItem) => {
                const row = [
                  projectData?.ProjectReference || "",
                  projectTitle,
                  projectData?.CustomerName || "",
                  typeof projectData?.ProjectManager === "object"
                    ? projectData.ProjectManager?.Title || ""
                    : projectData?.ProjectManager || "",
                  doc.docNo,
                  doc.description,
                  doc.category === "A" ? "A - Approval" : "I - Information",
                  doc.enggDivision || "",
                  doc.revNo.toString(),
                  formatDateForExcel(doc.originalPlan),
                  formatDateForExcel(doc.revisedPlan),
                  formatDateForExcel(doc.actualSubmission),
                  doc.submissionTransmittal,
                  formatDateForExcel(doc.approvalDate),
                  doc.approvalTransmittal,
                  doc.status,
                  doc.remarks,
                  "Input"
                ];
                wsData.push(row);

                if (doc.revisions) {
                  doc.revisions.forEach((revision: RevisionItem) => {
                    const revRow = [
                      projectData?.ProjectReference || "",
                      projectTitle,
                      projectData?.CustomerName || "",
                      typeof projectData?.ProjectManager === "object"
                        ? projectData.ProjectManager?.Title || ""
                        : projectData?.ProjectManager || "",
                      doc.docNo,
                      doc.description,
                      doc.category === "A" ? "A - Approval" : "I - Information",
                      doc.enggDivision || "",
                      revision.revNumber.toString(),
                      formatDateForExcel(revision.originalPlan),
                      formatDateForExcel(revision.revisedPlan),
                      formatDateForExcel(revision.actualSubmission),
                      revision.submissionTransmittal,
                      formatDateForExcel(revision.approvalDate),
                      revision.approvalTransmittal,
                      doc.status,
                      doc.remarks,
                      "Input (Revision)"
                    ];
                    wsData.push(revRow);
                  });
                }
              });

              // Output documents
              outputs.forEach((doc: DCSDocumentItem) => {
                const row = [
                  projectData?.ProjectReference || "",
                  projectTitle,
                  projectData?.CustomerName || "",
                  typeof projectData?.ProjectManager === "object"
                    ? projectData.ProjectManager?.Title || ""
                    : projectData?.ProjectManager || "",
                  doc.docNo,
                  doc.description,
                  doc.category === "A" ? "A - Approval" : "I - Information",
                  doc.enggDivision || "",
                  doc.revNo.toString(),
                  formatDateForExcel(doc.originalPlan),
                  formatDateForExcel(doc.revisedPlan),
                  formatDateForExcel(doc.actualSubmission),
                  doc.submissionTransmittal,
                  formatDateForExcel(doc.approvalDate),
                  doc.approvalTransmittal,
                  doc.status,
                  doc.remarks,
                  "Output"
                ];
                wsData.push(row);

                if (doc.revisions) {
                  doc.revisions.forEach((revision: RevisionItem) => {
                    const revRow = [
                      projectData?.ProjectReference || "",
                      projectTitle,
                      projectData?.CustomerName || "",
                      typeof projectData?.ProjectManager === "object"
                        ? projectData.ProjectManager?.Title || ""
                        : projectData?.ProjectManager || "",
                      doc.docNo,
                      doc.description,
                      doc.category === "A" ? "A - Approval" : "I - Information",
                      doc.enggDivision || "",
                      revision.revNumber.toString(),
                      formatDateForExcel(revision.originalPlan),
                      formatDateForExcel(revision.revisedPlan),
                      formatDateForExcel(revision.actualSubmission),
                      revision.submissionTransmittal,
                      formatDateForExcel(revision.approvalDate),
                      revision.approvalTransmittal,
                      doc.status,
                      doc.remarks,
                      "Output (Revision)"
                    ];
                    wsData.push(revRow);
                  });
                }
              });
            }
          });

          // Export as CSV
          const csvContent = wsData.map(row =>
            row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
          ).join('\n');

          const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
          const link = document.createElement("a");
          const url = URL.createObjectURL(blob);
          link.setAttribute("href", url);

          const fileName = `DCS_Report_${new Date().toISOString().split('T')[0]}.csv`;
          link.setAttribute("download", fileName);
          link.style.visibility = 'hidden';
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);

          hideExportProgress();
          toast(`Exported: ${fileName}`, "success");

        } catch (error) {
          hideExportProgress();
          console.error("Export error:", error);
          toast("Failed to export Excel report", "error");
        }
      };

      const loadDocumentsFromJSON = (projectData: DCSProjectData | LegacyDCSProjectData): void => {
        inputTableBody.innerHTML = '';
        outputTableBody.innerHTML = '';

        const dcsProjectData = convertToNewFormat(projectData);

        const createDocumentWithRevisions = (doc: DCSDocumentItem, tableType: 'input' | 'output', index: number) => {
          const docRow = createEditableDocumentRow(doc, index + 1, tableType === 'input');
          const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
          tableBody.appendChild(docRow);

          if (doc.revisions && doc.revisions.length > 0) {
            doc.revisions.forEach((revision: RevisionItem, revIndex: number) => {
              // ✅ Shared 8-cell revision row (7 fields + actions)
              const revisionHTML = buildRevisionRowHTML(doc.docId, revision);
              tableBody.insertAdjacentHTML('beforeend', revisionHTML);
            });

            updateMergedCells(doc.docId);
          }
        };

        if (dcsProjectData.documentStructure.inputs && dcsProjectData.documentStructure.inputs.length > 0) {
          dcsProjectData.documentStructure.inputs.forEach((doc: DCSDocumentItem, index: number) => {
            createDocumentWithRevisions(doc, 'input', index);
          });
        } else {
          const defaultInputDoc: DCSDocumentItem = {
            docId: `default_input_1`,
            docNo: "DOC-001",
            description: "",
            category: "A",
            enggDivision: "",
            revNo: 0,
            originalPlan: "",
            revisedPlan: "",
            actualSubmission: "",
            submissionTransmittal: "",
            approvalDate: "",
            approvalTransmittal: "",
            status: "Pending",
            remarks: "",
            isInput: true,
            revisions: [],
            docType: 'input'
          };
          createDocumentWithRevisions(defaultInputDoc, 'input', 0);
        }

        if (dcsProjectData.documentStructure.outputs && dcsProjectData.documentStructure.outputs.length > 0) {
          dcsProjectData.documentStructure.outputs.forEach((doc: DCSDocumentItem, index: number) => {
            createDocumentWithRevisions(doc, 'output', index);
          });
        } else {
          const defaultOutputDoc: DCSDocumentItem = {
            docId: `default_output_1`,
            docNo: "DOC-OUT-001",
            description: "",
            category: "A",
            enggDivision: "",
            revNo: 0,
            originalPlan: "",
            revisedPlan: "",
            actualSubmission: "",
            submissionTransmittal: "",
            approvalDate: "",
            approvalTransmittal: "",
            status: "Not Submitted",
            remarks: "",
            isInput: false,
            revisions: [],
            docType: 'output'
          };
          createDocumentWithRevisions(defaultOutputDoc, 'output', 0);
        }

        setTimeout(() => {
          attachDocumentEventListeners();
          updateAllSerialNumbers();
        }, 100);
      };

      const resetForm = (): void => {
        dcsForm.reset();

        formSection.classList.remove("dcs-view-mode");

        formSection.style.display = "none";
        tableSection.style.display = "block";

        inputTableBody.innerHTML = "";
        outputTableBody.innerHTML = "";
        containerTable.style.display = "none";
        resetProjectDetailsDisplay();
        existingDataAlert.style.display = "none";
        importButtonContainer.style.display = "none";

        formTitle.textContent = "Add Project Documents";
        formSection.classList.remove("dcs-edit-mode");

        projectSelect.disabled = false;

        editingId = null;

        if (addInputDocBtn) addInputDocBtn.style.display = "inline-block";
        if (addOutputDocBtn) addOutputDocBtn.style.display = "inline-block";

        // ✅ Restore BOTH save buttons with the Save label
getSaveButtons().forEach(btn => {
  btn.style.display = "inline-flex";
  btn.innerHTML = '💾 Save Document Data';
});

        resetDocumentTables();

        const formActions = document.querySelector('.dcs-form-actions');
        if (formActions) {
          formActions.innerHTML = `
            <button type="submit" class="dcs-btn-primary">
              💾 Save Document Data
            </button>
            <button type="button" class="dcs-btn-secondary dcs-back-to-table">
              ⬅ Back to List
            </button>
          `;

          const newBackBtn = formActions.querySelector('.dcs-back-to-table');
          if (newBackBtn) {
            newBackBtn.addEventListener('click', handleBackToTable);
          }
        }
      };

      // ✅ (removed) duplicate `backBtn.addEventListener("click", resetForm);` -
      // the Back button is already wired to handleBackToTable (which calls resetForm).

      // Function to collect documents from a table - UPDATED for textarea fields
      const updateProjectDetails = (projectId: number): void => {
        const project = projects.find((p) => p.Id === projectId);
        if (project) {
          setProjectDetailsDisplay(project);
        } else {
          resetProjectDetailsDisplay();
        }
      };

      const collectDocumentsFromTable = (tableType: 'input' | 'output'): DCSDocumentItem[] => {
        const docs: DCSDocumentItem[] = [];
        const tableId = tableType === 'input' ? 'dcsInputTableBody' : 'dcsOutputTableBody';
        const tableBody = document.getElementById(tableId);

        if (!tableBody) return docs;

        const docHeaders = tableBody.querySelectorAll('.dcs-document-header');

        docHeaders.forEach((docHeader: Element, docIndex: number) => {
          const docId = docHeader.getAttribute('data-doc-id') || `doc_${Date.now()}_${docIndex}`;

          const getValue = (selector: string, defaultValue: string = ''): string => {
            const element = docHeader.querySelector(selector);
            if (element instanceof HTMLTextAreaElement || element instanceof HTMLInputElement) {
              return element.value.trim();
            } else if (element instanceof HTMLSelectElement) {
              return element.value;
            }
            return defaultValue;
          };

          const doc: DCSDocumentItem = {
            docId: docId,
            docNo: getValue('.dcs-doc-no', `${tableType === 'input' ? 'DOC' : 'DOC-OUT'}-${String(docIndex + 1).padStart(3, '0')}`),
            description: getValue('.dcs-description-input'),
            category: getValue('.dcs-category-select', 'A'),
            enggDivision: getValue('.dcs-engg-division-select', ''),
            revNo: 0,
            originalPlan: getValue('td:nth-child(7) input.dcs-date-input'),
            revisedPlan: getValue('td:nth-child(8) input.dcs-date-input'),
            actualSubmission: getValue('td:nth-child(9) input.dcs-date-input'),
            submissionTransmittal: getValue('td:nth-child(10) textarea.dcs-transmittal-input'),
            approvalDate: getValue('td:nth-child(11) input.dcs-date-input'),
            approvalTransmittal: getValue('td:nth-child(12) textarea.dcs-transmittal-input'),
            status: getValue('td:nth-child(13) select.dcs-status-select', getDefaultStatus(tableType)),
            remarks: getValue('td:nth-child(14) textarea.dcs-remarks-input'),
            isInput: tableType === 'input',
            docType: tableType,
            revisions: []
          };

          const allRevisionRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);

          allRevisionRows.forEach((row: Element, revIndex: number) => {
            const getRevValue = (cellIndex: number, defaultValue: string = ''): string => {
              const cell = row.querySelector(`td:nth-child(${cellIndex})`);
              if (cell) {
                const input = cell.querySelector('input');
                const textarea = cell.querySelector('textarea');
                if (input) return input.value.trim();
                if (textarea) return textarea.value.trim();
              }
              return defaultValue;
            };

            const revNumberCell = row.querySelector('.dcs-rev-cell');
            const revNumber = revNumberCell ? parseInt(revNumberCell.textContent || '1') : revIndex + 1;

            const revisionStatus = doc.status;
            const revisionRemarks = doc.remarks;

            const revision: RevisionItem = {
              revNumber: revNumber,
              originalPlan: getRevValue(2),
              revisedPlan: getRevValue(3),
              actualSubmission: getRevValue(4),
              submissionTransmittal: getRevValue(5),
              approvalDate: getRevValue(6),
              approvalTransmittal: getRevValue(7),
              status: revisionStatus,
              remarks: revisionRemarks
            };

            doc.revisions.push(revision);
          });

          doc.revisions.sort((a, b) => a.revNumber - b.revNumber);

          docs.push(doc);
        });

        return docs;
      };

      // Updated function to collect ALL documents in standardized format
      const collectAllDocumentsInJSONFormat = (): { inputs: DCSDocumentItem[]; outputs: DCSDocumentItem[] } => {
        return {
          inputs: collectDocumentsFromTable('input'),
          outputs: collectDocumentsFromTable('output')
        };
      };

      const saveAllDataToSingleJSON = async (): Promise<void> => {
// ✅ Disable BOTH save buttons while saving
const saveBtns = getSaveButtons();
saveBtns.forEach(btn => {
  btn.disabled = true;
  btn.dataset.saving = "true";
});

const saveLoading = createLoadingOverlay("Saving document data...");
loaderHost().appendChild(saveLoading);
try {
          if (!projectSelect.value) {
            toast("Please select a project", "error");
            return;
          }

          const projectId = parseInt(projectSelect.value);
          const projectTitle = projectSelect.options[projectSelect.selectedIndex]?.text || `Project ${projectId}`;

          const isEditing = editingId !== null;

          // 🔐 ROLE GATE (same as SEIBIES / inventory): creating a NEW record needs
          // create rights (Admin / GH Head / PM Manager); updating an EXISTING record
          // needs edit rights for that project (Admin / own-assigned contributors).
          const saveProject = projects.find((p) => p.Id === projectId);
          if (isEditing) {
            if (!canEditProject(saveProject)) {
              toast("You do not have permission to update this project", "error");
              return;
            }
          } else {
            if (!canCreateRecord()) {
              toast("You do not have permission to add project documents", "error");
              return;
            }
          }

          const { inputs, outputs } = collectAllDocumentsInJSONFormat();

          if (inputs.length === 0 && outputs.length === 0) {
            toast("Please add at least one document", "warning");
            return;
          }

          // Create the complete data structure in NEW format
          const dcsProjectData: DCSProjectData = {
            projectId: projectId,
            projectTitle: projectTitle,
            lastUpdated: new Date().toISOString(),
            documentStructure: {
              inputs: inputs,
              outputs: outputs
            },
            summary: {
              totalDocuments: inputs.length + outputs.length,
              totalRevisions: inputs.reduce((sum: number, doc: DCSDocumentItem) =>
                sum + (doc.revisions?.length || 0), 0) +
                outputs.reduce((sum: number, doc: DCSDocumentItem) =>
                  sum + (doc.revisions?.length || 0), 0),
              inputDocuments: inputs.length,
              outputDocuments: outputs.length
            }
          };

          // Save to SharePoint with retry mechanism
          await saveWithRetry(dcsProjectData, projectId, editingId);

          const totalDocs = inputs.length + outputs.length;
          const totalRevisions = inputs.reduce((sum: number, doc: DCSDocumentItem) =>
            sum + (doc.revisions?.length || 0), 0) +
            outputs.reduce((sum: number, doc: DCSDocumentItem) =>
              sum + (doc.revisions?.length || 0), 0);

          toast(`✅ Success! ${isEditing ? 'Updated' : 'Saved'} ${totalDocs} documents with ${totalRevisions} revisions`, "success");

          editingId = null;

          handleBackToTable();
          await fetchData();

        } catch (error: any) {
          console.error("Save error:", error);
        } finally {
  removeLoadingOverlay();
  // ✅ Re-enable BOTH save buttons
  saveBtns.forEach(btn => {
    btn.disabled = false;
    btn.dataset.saving = "false";
  });
}
      };

      // Enhanced save function for SharePoint
      const saveDCSDataToSharePoint = async (
        dcsProjectData: DCSProjectData,
        projectId: number,
        existingItemId: number | null = null
      ): Promise<void> => {
        try {
          const digest: string = await getRequestDigest();

          let existingRecord: any = null;
          if (existingItemId) {
            existingRecord = dcsData.find((item: any) => item.Id === existingItemId);
          } else {
            for (const item of dcsData) {
              let itemProjectId: number | null = null;

              if (item[`${fieldMappings.Project}Id`]) {
                itemProjectId = Number(item[`${fieldMappings.Project}Id`]);
              } else if (item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]) {
                itemProjectId = Number(item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]);
              } else if (item.ProjectId) {
                itemProjectId = Number(item.ProjectId);
              } else if (item.Project_ID) {
                itemProjectId = Number(item.Project_ID);
              } else if (item[fieldMappings.Project.replace(/Id$/, '')]?.Id) {
                itemProjectId = Number(item[fieldMappings.Project.replace(/Id$/, '')].Id);
              }

              if (itemProjectId === projectId) {
                existingRecord = item;
                break;
              }
            }
          }

          // Prepare the complete data structure
          const combinedData: DCSProjectData = {
            projectId: projectId,
            projectTitle: projectSelect.options[projectSelect.selectedIndex]?.text || `Project ${projectId}`,
            lastUpdated: new Date().toISOString(),
            documentStructure: {
              inputs: dcsProjectData.documentStructure.inputs || [],
              outputs: dcsProjectData.documentStructure.outputs || []
            },
            summary: {
              totalDocuments: (dcsProjectData.documentStructure.inputs?.length || 0) +
                (dcsProjectData.documentStructure.outputs?.length || 0),
              totalRevisions: (dcsProjectData.documentStructure.inputs?.reduce((sum: number, doc: DCSDocumentItem) =>
                sum + (doc.revisions?.length || 0), 0) || 0) +
                (dcsProjectData.documentStructure.outputs?.reduce((sum: number, doc: DCSDocumentItem) =>
                  sum + (doc.revisions?.length || 0), 0) || 0),
              inputDocuments: dcsProjectData.documentStructure.inputs?.length || 0,
              outputDocuments: dcsProjectData.documentStructure.outputs?.length || 0
            }
          };

          let lookupFieldName: string = `${fieldMappings.Project}Id`;
          const noteFieldName: string = fieldMappings.DCSData;

          interface SharePointItemMetadata {
            type: string;
          }

          const itemData: Record<string, any> = {
            __metadata: {
              type: listEntityTypeName
            } as SharePointItemMetadata,
            Title: `DCS Data - Project ${projectId}`,
          };

          const possibleLookupFields: string[] = [
            `${fieldMappings.Project}Id`,
            `${fieldMappings.Project.replace(/Id$/, '')}Id`,
            "ProjectId",
            "Project_ID",
            "ProjectLookupId",
            "ProjectReferenceId"
          ];

          if (existingRecord) {
            for (const field of possibleLookupFields) {
              if (existingRecord[field] !== undefined) {
                lookupFieldName = field;
                break;
              }
            }
          }

          itemData[lookupFieldName] = projectId;
          itemData[noteFieldName] = JSON.stringify(combinedData);

          console.log("Saving data:", {
            noteFieldName: noteFieldName,
            lookupFieldName: lookupFieldName,
            metadataType: itemData.__metadata.type,
            isUpdate: !!existingRecord,
            itemId: existingRecord?.Id
          });

          let url: string = apiBase;
          const headers: HeadersInit = {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest,
          };

          if (existingRecord) {
            console.log('Getting current ETag for item', existingRecord.Id);

            const getItemResponse: Response = await fetch(`${apiBase}(${existingRecord.Id})`, {
              headers: {
                'Accept': 'application/json;odata=verbose',
              },
              credentials: 'same-origin'
            });

            if (!getItemResponse.ok) {
              throw new Error(`Failed to get current item for ETag: ${getItemResponse.status}`);
            }

            const currentETag: string | null = getItemResponse.headers.get('ETag');

            console.log('Current ETag:', currentETag);

            url = `${apiBase}(${existingRecord.Id})`;

            if (currentETag) {
              headers["IF-MATCH"] = currentETag;
            } else {
              headers["IF-MATCH"] = "*";
            }

            headers["X-HTTP-Method"] = "MERGE";
          }

          console.log(`Making ${existingRecord ? 'UPDATE' : 'CREATE'} request to:`, url);

          const response: Response = await fetch(url, {
            method: "POST",
            headers: headers,
            body: JSON.stringify(itemData),
            credentials: 'same-origin'
          });

          if (!response.ok) {
            const errorText: string = await response.text();
            console.error('❌ Save error details:', {
              status: response.status,
              statusText: response.statusText,
              errorText: errorText,
              url: url,
              metadataType: itemData.__metadata.type,
              lookupFieldName: lookupFieldName,
              isUpdate: !!existingRecord
            });

            if (response.status === 409) {
              console.log('⚠️ 409 Conflict - Item was modified by someone else. Retrying...');
              try {
                await fetchData();
                const refreshedData: any = dcsData.find((item: any) => item.Id === existingRecord?.Id);
                if (refreshedData) {
                  console.log('Retrying with refreshed data...');
                  return await saveDCSDataToSharePoint(dcsProjectData, projectId, existingItemId);
                }
              } catch (retryError: any) {
                console.error('Retry failed:', retryError);
              }
            }

            if (errorText.includes("valid type") || errorText.includes("EntityTypeName") || errorText.includes("could not resolve")) {
              const alternativeTypes: string[] = [
                "SP.Data.DCSTracker_x0020_ListListItem",
                "SP.Data.DcsTrackerItem",
                "SP.Data.DCSTrackerListItem",
                "SP.Data.DCSTrackerItem",
                "SP.Data.DcsTrackerListItem"
              ];

              for (const altType of alternativeTypes) {
                if (altType !== listEntityTypeName) {
                  try {
                    console.log(`Trying alternative type: ${altType}`);

                    const altItemData: Record<string, any> = { ...itemData };
                    altItemData.__metadata.type = altType;

                    const retryResponse: Response = await fetch(url, {
                      method: "POST",
                      headers: headers,
                      body: JSON.stringify(altItemData),
                      credentials: 'same-origin'
                    });

                    if (retryResponse.ok) {
                      console.log(`✅ Success with type: ${altType}`);
                      listEntityTypeName = altType;
                      await fetchData();
                      return;
                    }
                  } catch (retryError: any) {
                    console.log(`Type ${altType} failed:`, retryError);
                  }
                }
              }
            }

            throw new Error(`Save failed: ${response.status} - ${response.statusText}`);
          }

          console.log('✅ Save successful!');
          await fetchData();

        } catch (error: any) {
          console.error('❌ Save error:', error);
          throw error;
        }
      };

      // CORRECTED: saveWithRetry function (no unused variables)
      const saveWithRetry = async (
        dcsProjectData: DCSProjectData,
        projectId: number,
        existingItemId: number | null = null,
        retries = 3
      ): Promise<void> => {
        let lastError: any = null;

        for (let attempt = 1; attempt <= retries; attempt++) {
          try {
            console.log(`Save attempt ${attempt} of ${retries}`);

            if (attempt > 1) {
              await fetchData();
            }

            await saveDCSDataToSharePoint(dcsProjectData, projectId, existingItemId);

            return;

          } catch (error: any) {
            lastError = error;

            const shouldRetry = (
              error.message.includes("409") ||
              error.message.includes("412") ||
              error.message.includes("ETag") ||
              error.message.includes("500") ||
              error.message.includes("Conflict")
            ) && attempt < retries;

            if (shouldRetry) {
              const delay = 1000 * Math.pow(2, attempt - 1);
              console.log(`Retry ${attempt} after ${delay}ms delay`);
              await new Promise(resolve => setTimeout(resolve, delay));
              continue;
            }

            break;
          }
        }

        throw lastError || new Error('Save failed after all retry attempts');
      };

      const initializePagination = (data: any[]): void => {
        console.log('📄 initializePagination called with:', data.length, 'items');

        // ✅ Re-apply the active column sort on every (re)load so sorting survives
        // filtering, Reset Filter and post-save refreshes
        filteredDCSData = applySort(data);

        console.log('📊 filteredDCSData set to:', filteredDCSData.length, 'items');

        calculateTotalPages();

        currentPage = 1;

        updatePaginationControls();

        renderCurrentPage();

        console.log('✅ Pagination initialized. Showing page', currentPage, 'of', totalPages);
      };

      const calculateTotalPages = (): void => {
        totalPages = Math.ceil(filteredDCSData.length / pageSize);
      };

      const updatePaginationControls = (): void => {
        const currentPageSpan = document.getElementById("currentPage") as HTMLSpanElement;
        const totalPagesSpan = document.getElementById("totalPages") as HTMLSpanElement;
        const startItemSpan = document.getElementById("startItem") as HTMLSpanElement;
        const endItemSpan = document.getElementById("endItem") as HTMLSpanElement;
        const totalItemsSpan = document.getElementById("totalItems") as HTMLSpanElement;

        if (currentPageSpan) currentPageSpan.textContent = currentPage.toString();
        if (totalPagesSpan) totalPagesSpan.textContent = totalPages.toString();

        const startItem = (currentPage - 1) * pageSize + 1;
        const endItem = Math.min(currentPage * pageSize, filteredDCSData.length);

        if (startItemSpan) startItemSpan.textContent = startItem.toString();
        if (endItemSpan) endItemSpan.textContent = endItem.toString();
        if (totalItemsSpan) totalItemsSpan.textContent = filteredDCSData.length.toString();

        const pageSizeSelect = document.getElementById("pageSizeSelect") as HTMLSelectElement;
        if (pageSizeSelect) pageSizeSelect.value = pageSize.toString();

        const firstPageBtn = document.getElementById("firstPage") as HTMLButtonElement;
        const prevPageBtn = document.getElementById("prevPage") as HTMLButtonElement;
        const nextPageBtn = document.getElementById("nextPage") as HTMLButtonElement;
        const lastPageBtn = document.getElementById("lastPage") as HTMLButtonElement;

        if (firstPageBtn) firstPageBtn.disabled = currentPage === 1;
        if (prevPageBtn) prevPageBtn.disabled = currentPage === 1;
        if (nextPageBtn) nextPageBtn.disabled = currentPage === totalPages;
        if (lastPageBtn) lastPageBtn.disabled = currentPage === totalPages;

        updatePageNumbers();
      };

      const updatePageNumbers = (): void => {
        const pageNumbersDiv = document.getElementById("pageNumbers") as HTMLDivElement;
        if (!pageNumbersDiv) return;

        pageNumbersDiv.innerHTML = '';

        const maxVisiblePages = 5;
        let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
        let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

        if (endPage - startPage + 1 < maxVisiblePages) {
          startPage = Math.max(1, endPage - maxVisiblePages + 1);
        }

        if (startPage > 1) {
          const dotsBtn = document.createElement('button');
          dotsBtn.className = 'dcs-page-number dots'; // ✅ was 'page-number dots' (unstyled)
          dotsBtn.textContent = '...';
          dotsBtn.disabled = true;
          pageNumbersDiv.appendChild(dotsBtn);
        }

        for (let i = startPage; i <= endPage; i++) {
          const pageBtn = document.createElement('button');
          pageBtn.className = `dcs-page-number ${i === currentPage ? 'active' : ''}`; // ✅ dcs- prefix
          pageBtn.textContent = i.toString();
          pageBtn.addEventListener('click', () => goToPage(i));
          pageNumbersDiv.appendChild(pageBtn);
        }

        if (endPage < totalPages) {
          const dotsBtn = document.createElement('button');
          dotsBtn.className = 'dcs-page-number dots'; // ✅ was 'page-number dots' (unstyled)
          dotsBtn.textContent = '...';
          dotsBtn.disabled = true;
          pageNumbersDiv.appendChild(dotsBtn);
        }
      };

      const goToPage = (page: number): void => {
        if (page < 1 || page > totalPages || page === currentPage) return;

        currentPage = page;
        updatePaginationControls();
        renderCurrentPage();

        const tableContainer = document.querySelector('.table-section') as HTMLElement;
        if (tableContainer) {
          tableContainer.scrollTop = 0;
        }
      };

      const changePageSize = (newSize: number): void => {
        if (newSize === pageSize) return;

        pageSize = newSize;

        const oldStartItem = (currentPage - 1) * pageSize;
        currentPage = Math.floor(oldStartItem / pageSize) + 1;

        calculateTotalPages();

        if (currentPage > totalPages) {
          currentPage = totalPages || 1;
        }

        updatePaginationControls();
        renderCurrentPage();
      };

      const renderCurrentPage = (): void => {
        const startIndex = (currentPage - 1) * pageSize;
        const endIndex = Math.min(startIndex + pageSize, filteredDCSData.length);
        const pageData = filteredDCSData.slice(startIndex, endIndex);

        console.log('📄 renderCurrentPage called');
        console.log('📊 Page data to render:', pageData.length, 'items');
        console.log('📈 Showing items', startIndex + 1, 'to', endIndex, 'of', filteredDCSData.length);

        renderTable(pageData);
      };

      const getCorrectListItemType = async (): Promise<void> => {
        try {
          // Construct the exact URL for your list
          const url = `${siteUrl}/_api/web/lists/getbytitle('DCSTracker')?$select=ListItemEntityTypeFullName`;

          const response = await fetch(url, {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
            },
            credentials: 'same-origin'
          });

          const data = await response.json();
          const correctTypeName = data.d.ListItemEntityTypeFullName;

          console.log('✅ Correct List Item Type Name:', correctTypeName);
          console.log('UPDATE your code - set listEntityTypeName to:', correctTypeName);

        } catch (error) {
          console.error('❌ Failed to get correct type name:', error);
        }
      };

      // Call this function when your page loads or from the browser console
      await getCorrectListItemType();

      // Add this function to debug SharePoint list fields
      const debugListFields = async (): Promise<void> => {
        try {
          const res = await fetch(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false`,
            {
              headers: { Accept: "application/json;odata=verbose" },
              credentials: 'same-origin'
            }
          );

          if (res.ok) {
            const data = await res.json();
            const fields = data.d.results;

            fields.forEach((field: any) => {
            });

            const lookupFields = fields.filter((f: any) => f.TypeAsString.includes("Lookup"));
            lookupFields.forEach((field: any) => {
            });

            const noteFields = fields.filter((f: any) => f.TypeAsString === "Note");
            noteFields.forEach((field: any) => {
            });
          }

        } catch (error) {
          console.error("Failed to debug fields:", error);
        }
      };

      // Call this during initialization
      await debugListFields();

      // Add import functionality functions
      const setupImportExportFunctionality = (): void => {
        const importExcelBtn = document.querySelector('.dcs-import-excel-btn') as HTMLButtonElement;
        const importExcelInput = document.getElementById('dcsImportExcelInput') as HTMLInputElement;
        const exportProjectExcelBtn = document.querySelector('.dcs-export-project-excel-btn') as HTMLButtonElement;
        const downloadExistingBtn = document.querySelector('.dcs-download-existing-btn') as HTMLButtonElement;

        if (importExcelBtn && importExcelInput) {
          importExcelBtn.addEventListener('click', () => {
            importExcelInput.click();
          });

          importExcelInput.addEventListener('change', handleExcelImport);
        }

        if (exportProjectExcelBtn) {
          exportProjectExcelBtn.addEventListener('click', exportTemplateExcel);
        }

        if (downloadExistingBtn) {
          downloadExistingBtn.addEventListener('click', downloadExistingData);
        }
      };

      // Handle Excel Import
      const handleExcelImport = async (event: Event): Promise<void> => {
        // 🔐 ROLE GATE: Excel import replaces project document data -> create roles only
        if (!canCreateRecord()) {
          toast("You do not have permission to import Excel data", "error");
          const inputReset = event.target as HTMLInputElement;
          if (inputReset) inputReset.value = "";
          return;
        }

        const input = event.target as HTMLInputElement;
        const file = input.files?.[0];

        if (!file) {
          toast('Please select a file', 'warning');
          return;
        }

        if (!file.name.match(/\.(xlsx|xls)$/i)) {
          toast('Please select an Excel file (.xlsx or .xls)', 'error');
          input.value = '';
          return;
        }

        try {
          showImportProgress('Reading Excel file...');

          const reader = new FileReader();
          reader.onload = (e) => {
            try {
              const data = e.target?.result;
              if (data) {
                const workbook = XLSX.read(data, {
                  type: 'binary',
                  cellDates: true // Parse dates automatically
                });

                const sheetName = workbook.SheetNames.find((name: string) =>
                  name.includes('DCS') || name.includes('Documents') || name.includes('Sheet')
                ) || workbook.SheetNames[0];

                const worksheet = workbook.Sheets[sheetName];

                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                  header: 1,
                  defval: '',
                  raw: false // Convert dates to strings
                }) as any[][];

                // Process the imported data
                processImportedData(jsonData);
              }
            } catch (error) {
              console.error('Error processing Excel:', error);
              toast('Failed to process Excel file. Please check the format.', 'error');
            } finally {
              hideImportProgress();
              input.value = '';
            }
          };

          reader.onerror = () => {
            hideImportProgress();
            toast('Failed to read file', 'error');
            input.value = '';
          };

          reader.readAsBinaryString(file);

        } catch (error) {
          hideImportProgress();
          console.error('Import error:', error);
          toast('Failed to import Excel file', 'error');
          input.value = '';
        }
      };

      // Process imported data from Excel - Enhanced to handle new documents and revisions
      const processImportedData = (data: any[][]): void => {
        let currentSection: 'input' | 'output' | null = null;
        let currentDoc: DCSDocumentItem | null = null;
        let currentDocRevisions: RevisionItem[] = [];

        const existingInputDocs = getDocumentsFromTable('input');
        const existingOutputDocs = getDocumentsFromTable('output');

        const newInputDocs: DCSDocumentItem[] = [...existingInputDocs];
        const newOutputDocs: DCSDocumentItem[] = [...existingOutputDocs];

        let hasEnggDivisionColumn = false;
        for (let i = 0; i < Math.min(10, data.length); i++) {
          const row = data[i];
          if (row && row.length > 0) {
            const firstCell = String(row[0] || '').trim().toLowerCase();
            if (firstCell === 'sr no') {
              const fifth = String(row[4] || '').trim().toLowerCase();
              hasEnggDivisionColumn = fifth.includes('engg') || fifth.includes('division');
              break;
            }
          }
        }
        const colOffset = hasEnggDivisionColumn ? 1 : 0;
        console.log(`Excel import layout: ${hasEnggDivisionColumn ? 'with' : 'without'} Engg Division column`);

        let dataStartRow = 0;
        for (let i = 0; i < Math.min(10, data.length); i++) {
          const row = data[i];
          if (row && row.length > 0) {
            const firstCell = String(row[0] || '').trim();
            if (firstCell.includes('DESIGN OUTPUT') || firstCell.includes('DESIGN INPUT')) {
              dataStartRow = i;
              break;
            }
          }
        }

        for (let i = dataStartRow; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const processedRow = row.map(cell => {
            if (cell === null || cell === undefined) return '';
            return String(cell).trim();
          });

          const firstCell = processedRow[0] || '';

          if (firstCell.includes('DESIGN OUTPUT')) {
            if (currentDoc && currentSection) {
              saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
            }
            currentSection = 'output';
            currentDoc = null;
            currentDocRevisions = [];
            continue;
          } else if (firstCell.includes('DESIGN INPUT')) {
            if (currentDoc && currentSection) {
              saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
            }
            currentSection = 'input';
            currentDoc = null;
            currentDocRevisions = [];
            continue;
          } else if (firstCell === '' || !firstCell) {
            if (currentDoc && currentSection) {
              saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
              currentDoc = null;
              currentDocRevisions = [];
            }
            continue;
          }

          const docNo = processedRow[1] || '';
          const description = processedRow[2] || '';
          const category = (processedRow[3] || 'A').charAt(0);
          const enggDivisionRaw = hasEnggDivisionColumn ? (processedRow[4] || '') : '';
          const enggDivision = ENGG_DIVISIONS.find(d => d.toLowerCase() === enggDivisionRaw.toLowerCase()) || enggDivisionRaw;
          const revNo = parseInt(processedRow[4 + colOffset] || '0') || 0;

          const originalPlan = parseExcelDate(processedRow[5 + colOffset] || '');
          const revisedPlan = parseExcelDate(processedRow[6 + colOffset] || '');
          const actualSubmission = parseExcelDate(processedRow[7 + colOffset] || '');
          const submissionTransmittal = processedRow[8 + colOffset] || '';
          const approvalDate = parseExcelDate(processedRow[9 + colOffset] || '');
          const approvalTransmittal = processedRow[10 + colOffset] || '';
          const statusRaw = processedRow[11 + colOffset] || '';
          const sectionStatusOptions = currentSection ? getStatusOptions(currentSection) : [];
          const status = sectionStatusOptions.find(o => o.toLowerCase() === statusRaw.toLowerCase())
            || statusRaw
            || (currentSection ? getDefaultStatus(currentSection) : 'Pending');
          const remarks = processedRow[12 + colOffset] || '';

          if (!docNo) {
            continue;
          }

          if (revNo === 0) {
            if (currentDoc && currentSection) {
              saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
            }

            const existingDocs = currentSection === 'input' ? newInputDocs : newOutputDocs;
            let existingDocIndex = -1;

            if (currentSection) {
              existingDocIndex = existingDocs.findIndex(doc =>
                doc.docNo === docNo && doc.isInput === (currentSection === 'input')
              );
            }

            if (existingDocIndex >= 0 && currentSection) {
              const existingDoc = existingDocs[existingDocIndex];
              currentDoc = {
                ...existingDoc,
                docNo: docNo,
                description: description || existingDoc.description,
                category: (category === 'A' || category === 'I') ? category : existingDoc.category,
                enggDivision: enggDivision || existingDoc.enggDivision || '',
                revNo: 0,
                originalPlan: originalPlan || existingDoc.originalPlan,
                revisedPlan: revisedPlan || existingDoc.revisedPlan,
                actualSubmission: actualSubmission || existingDoc.actualSubmission,
                submissionTransmittal: submissionTransmittal || existingDoc.submissionTransmittal,
                approvalDate: approvalDate || existingDoc.approvalDate,
                approvalTransmittal: approvalTransmittal || existingDoc.approvalTransmittal,
                status: status || existingDoc.status,
                remarks: remarks || existingDoc.remarks,
                revisions: [...existingDoc.revisions]
              };

              if (currentSection === 'input') {
                newInputDocs.splice(existingDocIndex, 1);
              } else {
                newOutputDocs.splice(existingDocIndex, 1);
              }
            } else {
              currentDoc = {
                docId: `imported_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                docNo: docNo,
                description: description,
                category: category === 'A' || category === 'I' ? category : 'A',
                enggDivision: enggDivision || '',
                revNo: 0,
                originalPlan: originalPlan,
                revisedPlan: revisedPlan,
                actualSubmission: actualSubmission,
                submissionTransmittal: submissionTransmittal,
                approvalDate: approvalDate,
                approvalTransmittal: approvalTransmittal,
                status: status || (currentSection ? getDefaultStatus(currentSection) : 'Pending'),
                remarks: remarks,
                isInput: currentSection === 'input',
                revisions: [],
                docType: currentSection || undefined
              };
            }

            currentDocRevisions = [];

          } else if (revNo > 0 && currentDoc) {
            const revision: RevisionItem = {
              revNumber: revNo,
              originalPlan: originalPlan,
              revisedPlan: revisedPlan,
              actualSubmission: actualSubmission,
              submissionTransmittal: submissionTransmittal,
              approvalDate: approvalDate,
              approvalTransmittal: approvalTransmittal,
              status: status || currentDoc.status,
              remarks: remarks || currentDoc.remarks
            };

            const existingRevisionIndex = currentDocRevisions.findIndex(r => r.revNumber === revNo);

            if (existingRevisionIndex >= 0) {
              currentDocRevisions[existingRevisionIndex] = revision;
            } else {
              currentDocRevisions.push(revision);
            }
          }
        }

        if (currentDoc && currentSection) {
          saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
        }

        updateTablesWithImportedData(newInputDocs, newOutputDocs);

        toast('Excel data imported successfully! New documents/revisions added.', 'success');
      };

      // Helper to save current document to appropriate collection
      const saveCurrentDocument = (
        doc: DCSDocumentItem,
        revisions: RevisionItem[],
        inputDocs: DCSDocumentItem[],
        outputDocs: DCSDocumentItem[]
      ): void => {
        doc.revisions = [...revisions];

        if (doc.isInput) {
          inputDocs.push(doc);
        } else {
          outputDocs.push(doc);
        }
      };

      // Get existing documents from table
      const getDocumentsFromTable = (tableType: 'input' | 'output'): DCSDocumentItem[] => {
        const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
        const docs: DCSDocumentItem[] = [];

        if (!tableBody) return docs;

        const docHeaders = tableBody.querySelectorAll('.dcs-document-header');

        docHeaders.forEach((docHeader: Element) => {
          const docId = docHeader.getAttribute('data-doc-id') || '';

          const getValue = (selector: string, defaultValue: string = ''): string => {
            const element = docHeader.querySelector(selector);
            if (element instanceof HTMLTextAreaElement || element instanceof HTMLInputElement) {
              return element.value.trim();
            } else if (element instanceof HTMLSelectElement) {
              return element.value;
            }
            return defaultValue;
          };

          const doc: DCSDocumentItem = {
            docId: docId,
            docNo: getValue('.dcs-doc-no'),
            description: getValue('.dcs-description-input'),
            category: getValue('.dcs-category-select', 'A'),
            enggDivision: getValue('.dcs-engg-division-select', ''),
            revNo: 0,
            originalPlan: getValue('td:nth-child(7) input.dcs-date-input'),
            revisedPlan: getValue('td:nth-child(8) input.dcs-date-input'),
            actualSubmission: getValue('td:nth-child(9) input.dcs-date-input'),
            submissionTransmittal: getValue('td:nth-child(10) textarea.dcs-transmittal-input'),
            approvalDate: getValue('td:nth-child(11) input.dcs-date-input'),
            approvalTransmittal: getValue('td:nth-child(12) textarea.dcs-transmittal-input'),
            status: getValue('td:nth-child(13) select.dcs-status-select', getDefaultStatus(tableType)),
            remarks: getValue('td:nth-child(14) textarea.dcs-remarks-input'),
            isInput: tableType === 'input',
            revisions: [],
            docType: tableType
          };

          const revisionRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);

          revisionRows.forEach((row: Element) => {
            const revCell = row.querySelector('.dcs-rev-cell');
            const revNumber = revCell ? parseInt(revCell.textContent || '1') : 1;

            const getRevValue = (cellIndex: number): string => {
              const cell = row.querySelector(`td:nth-child(${cellIndex})`);
              if (cell) {
                const input = cell.querySelector('input');
                const textarea = cell.querySelector('textarea');
                if (input) return input.value.trim();
                if (textarea) return textarea.value.trim();
              }
              return '';
            };

            const revision: RevisionItem = {
              revNumber: revNumber,
              originalPlan: getRevValue(2),
              revisedPlan: getRevValue(3),
              actualSubmission: getRevValue(4),
              submissionTransmittal: getRevValue(5),
              approvalDate: getRevValue(6),
              approvalTransmittal: getRevValue(7),
              status: doc.status,
              remarks: doc.remarks
            };

            doc.revisions.push(revision);
          });

          docs.push(doc);
        });

        return docs;
      };

      // Update tables with combined data
      const updateTablesWithImportedData = (inputDocs: DCSDocumentItem[], outputDocs: DCSDocumentItem[]): void => {
        inputTableBody.innerHTML = '';
        inputDocs.forEach((doc, index) => {
          addDocumentToTable(doc, 'input', index + 1);
        });

        outputTableBody.innerHTML = '';
        outputDocs.forEach((doc, index) => {
          addDocumentToTable(doc, 'output', index + 1);
        });

        updateAllSerialNumbers();

        setTimeout(() => {
          attachDocumentEventListeners();
          ensureDocumentButtonsVisible();
        }, 100);
      };

      // Updated addDocumentToTable with srNo parameter
      const addDocumentToTable = (doc: DCSDocumentItem, tableType: 'input' | 'output', srNo: number): void => {
        const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
        if (!tableBody) return;

        const docRow = createEditableDocumentRow(doc, srNo, tableType === 'input');
        tableBody.appendChild(docRow);

        if (doc.revisions && doc.revisions.length > 0) {
          doc.revisions.forEach((revision: RevisionItem) => {
            // ✅ Shared 8-cell revision row (7 fields + actions)
            const revisionHTML = buildRevisionRowHTML(doc.docId, revision);
            tableBody.insertAdjacentHTML('beforeend', revisionHTML);
          });

          updateMergedCells(doc.docId);
        }
      };

      // Parse Excel date (handles both string and Excel serial numbers)
      const parseExcelDate = (dateStr: string): string => {
        if (!dateStr || dateStr.trim() === '') return '';

        if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
          return dateStr;
        }

        if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
          const parts = dateStr.split('/');
          if (parts.length === 3) {
            const date1 = new Date(`${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`);
            const date2 = new Date(`${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`);

            if (!isNaN(date1.getTime())) {
              return date1.toISOString().split('T')[0];
            } else if (!isNaN(date2.getTime())) {
              return date2.toISOString().split('T')[0];
            }
          }
        }

        if (!isNaN(Number(dateStr)) && dateStr !== '') {
          const excelDate = Number(dateStr);
          const date = new Date((excelDate - 25569) * 86400 * 1000);
          if (!isNaN(date.getTime())) {
            return date.toISOString().split('T')[0];
          }
        }

        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) {
          return date.toISOString().split('T')[0];
        }

        return '';
      };

      const exportTemplateExcel = (): void => {
        try {
          const projectSelect = document.getElementById("dcsProjectSelect") as HTMLSelectElement;
          const projectId = parseInt(projectSelect.value);
          const project = projects.find(p => p.Id === projectId);

          if (!project) {
            toast('Please select a project first', 'warning');
            return;
          }

          const excelData: any[][] = [];

          excelData.push(
            ['Sr No', 'Document No', 'Description', 'Category', 'Engg Division', 'Rev No',
              'Original Plan', 'Revised Plan', 'Actual Submission', 'Transmittal No',
              'Approval Date', 'Transmittal No', 'Status', 'Remarks']
          );
          excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);

          excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '', '']);

          excelData.push([
            '1', 'DOC-OUT-001', 'Sample Output Document Description', 'A - Approval', 'Automation', '0',
            '', '', '', '', '', '', 'Not Submitted', 'Add output document details here'
          ]);

          excelData.push([
            '1', 'DOC-OUT-001', 'Sample Output Document Description', 'A - Approval', 'Automation', '1',
            '', '', '', '', '', '', 'Not Submitted', 'Add revision details here'
          ]);

          excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);

          excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '', '']);

          excelData.push([
            '1', 'DOC-001', 'Sample Input Document Description', 'A - Approval', 'Automation', '0',
            '', '', '', '', '', '', 'Pending', 'Add input document details here'
          ]);

          excelData.push([
            '1', 'DOC-001', 'Sample Input Document Description', 'A - Approval', 'Automation', '1',
            '', '', '', '', '', '', 'Pending', 'Add revision details here'
          ]);

          const ws = XLSX.utils.aoa_to_sheet(excelData);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'DCS Template');

          ws['!cols'] = [
            { wch: 8 },   // Sr No
            { wch: 20 },  // Document No
            { wch: 35 },  // Description
            { wch: 12 },  // Category
            { wch: 18 },  // Engg Division
            { wch: 8 },   // Rev No
            { wch: 15 },  // Original Plan
            { wch: 15 },  // Revised Plan
            { wch: 18 },  // Actual Submission
            { wch: 20 },  // Transmittal No
            { wch: 15 },  // Approval Date
            { wch: 20 },  // Transmittal No
            { wch: 20 },  // Status
            { wch: 25 }   // Remarks
          ];

          const timestamp = new Date().toISOString().split('T')[0];
          const safeProjectName = project.Title.replace(/[^a-zA-Z0-9-_ ]/g, '_').substring(0, 30);
          const fileName = `DCS_Template_${safeProjectName}_${timestamp}.xlsx`;

          XLSX.writeFile(wb, fileName);

          toast('Template exported successfully! Edit and import when ready.', 'success');

        } catch (error) {
          console.error('Export template error:', error);
          toast('Failed to export template', 'error');
        }
      };

      const downloadExistingData = async (): Promise<void> => {
        try {
          const projectSelect = document.getElementById("dcsProjectSelect") as HTMLSelectElement;
          const projectId = parseInt(projectSelect.value);

          if (!projectId || isNaN(projectId)) {
            toast('Please select a project first', 'warning');
            return;
          }

          const project = projects.find(p => p.Id === projectId);

          if (!project) {
            toast('Project not found', 'error');
            return;
          }

          let existingRecord: any = null;
          let foundData = false;

          for (const item of dcsData) {
            let itemProjectId: number | null = null;

            if (item[`${fieldMappings.Project}Id`]) {
              itemProjectId = Number(item[`${fieldMappings.Project}Id`]);
            } else if (item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]) {
              itemProjectId = Number(item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]);
            } else if (item[fieldMappings.Project]?.Id) {
              itemProjectId = Number(item[fieldMappings.Project].Id);
            } else if (item.ProjectId) {
              itemProjectId = Number(item.ProjectId);
            } else if (item.Project_ID) {
              itemProjectId = Number(item.Project_ID);
            }

            if (itemProjectId === projectId) {
              existingRecord = item;
              foundData = true;
              break;
            }
          }

          const excelData: any[][] = [];

          excelData.push(
            ['Sr No', 'Document No', 'Description', 'Category', 'Engg Division', 'Rev No',
              'Original Plan', 'Revised Plan', 'Actual Submission', 'Transmittal No',
              'Approval Date', 'Transmittal No', 'Status', 'Remarks']
          );
          excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);

          let hasAnyData = false;

          if (foundData && existingRecord && existingRecord[fieldMappings.DCSData]) {
            const parsedData = parseDCSData(existingRecord[fieldMappings.DCSData]);

            if (parsedData) {
              const inputs = getInputsFromProjectData(parsedData) || [];
              const outputs = getOutputsFromProjectData(parsedData) || [];

              hasAnyData = inputs.length > 0 || outputs.length > 0;

              if (outputs.length > 0) {
                excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '', '']);

                outputs.forEach((doc: DCSDocumentItem, index: number) => {
                  excelData.push([
                    (index + 1).toString(),
                    doc.docNo || `DOC-OUT-${String(index + 1).padStart(3, '0')}`,
                    doc.description || '',
                    doc.category === 'A' ? 'A - Approval' : 'I - Information',
                    doc.enggDivision || '',
                    '0',
                    formatDateForExcel(doc.originalPlan),
                    formatDateForExcel(doc.revisedPlan),
                    formatDateForExcel(doc.actualSubmission),
                    doc.submissionTransmittal || '',
                    formatDateForExcel(doc.approvalDate),
                    doc.approvalTransmittal || '',
                    doc.status || 'Not Submitted',
                    doc.remarks || ''
                  ]);

                  if (doc.revisions && doc.revisions.length > 0) {
                    doc.revisions.forEach((revision: RevisionItem) => {
                      excelData.push([
                        (index + 1).toString(),
                        doc.docNo || `DOC-OUT-${String(index + 1).padStart(3, '0')}`,
                        doc.description || '',
                        doc.category === 'A' ? 'A - Approval' : 'I - Information',
                        doc.enggDivision || '',
                        revision.revNumber.toString(),
                        formatDateForExcel(revision.originalPlan),
                        formatDateForExcel(revision.revisedPlan),
                        formatDateForExcel(revision.actualSubmission),
                        revision.submissionTransmittal || '',
                        formatDateForExcel(revision.approvalDate),
                        revision.approvalTransmittal || '',
                        revision.status || doc.status || 'Not Submitted',
                        revision.remarks || doc.remarks || ''
                      ]);
                    });
                  }
                });
              } else {
                excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '', '']);
                excelData.push([
                  '1', 'DOC-OUT-001', '', 'A - Approval', '', '0',
                  '', '', '', '', '', '', 'Not Submitted', ''
                ]);
              }

              excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);

              if (inputs.length > 0) {
                excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '', '']);

                inputs.forEach((doc: DCSDocumentItem, index: number) => {
                  excelData.push([
                    (index + 1).toString(),
                    doc.docNo || `DOC-${String(index + 1).padStart(3, '0')}`,
                    doc.description || '',
                    doc.category === 'A' ? 'A - Approval' : 'I - Information',
                    doc.enggDivision || '',
                    '0',
                    formatDateForExcel(doc.originalPlan),
                    formatDateForExcel(doc.revisedPlan),
                    formatDateForExcel(doc.actualSubmission),
                    doc.submissionTransmittal || '',
                    formatDateForExcel(doc.approvalDate),
                    doc.approvalTransmittal || '',
                    doc.status || 'Pending',
                    doc.remarks || ''
                  ]);

                  if (doc.revisions && doc.revisions.length > 0) {
                    doc.revisions.forEach((revision: RevisionItem) => {
                      excelData.push([
                        (index + 1).toString(),
                        doc.docNo || `DOC-${String(index + 1).padStart(3, '0')}`,
                        doc.description || '',
                        doc.category === 'A' ? 'A - Approval' : 'I - Information',
                        doc.enggDivision || '',
                        revision.revNumber.toString(),
                        formatDateForExcel(revision.originalPlan),
                        formatDateForExcel(revision.revisedPlan),
                        formatDateForExcel(revision.actualSubmission),
                        revision.submissionTransmittal || '',
                        formatDateForExcel(revision.approvalDate),
                        revision.approvalTransmittal || '',
                        revision.status || doc.status || 'Pending',
                        revision.remarks || doc.remarks || ''
                      ]);
                    });
                  }
                });
              } else {
                excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '', '']);
                excelData.push([
                  '1', 'DOC-001', '', 'A - Approval', '', '0',
                  '', '', '', '', '', '', 'Pending', ''
                ]);
              }
            }
          } else {
            hasAnyData = false;

            excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '', '']);
            excelData.push([
              '1', 'DOC-OUT-001', '', 'A - Approval', '', '0',
              '', '', '', '', '', '', 'Not Submitted', ''
            ]);

            excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);

            excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '', '']);
            excelData.push([
              '1', 'DOC-001', '', 'A - Approval', '', '0',
              '', '', '', '', '', '', 'Pending', ''
            ]);
          }

          const ws = XLSX.utils.aoa_to_sheet(excelData);

          ws['!cols'] = [
            { wch: 8 },   // Sr No
            { wch: 20 },  // Document No
            { wch: 35 },  // Description
            { wch: 12 },  // Category
            { wch: 18 },  // Engg Division
            { wch: 8 },   // Rev No
            { wch: 15 },  // Original Plan
            { wch: 15 },  // Revised Plan
            { wch: 18 },  // Actual Submission
            { wch: 20 },  // Transmittal No
            { wch: 15 },  // Approval Date
            { wch: 20 },  // Transmittal No
            { wch: 20 },  // Status
            { wch: 25 }   // Remarks
          ];

          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'DCS Data');

          const timestamp = new Date().toISOString().split('T')[0];
          const safeProjectName = project.Title.replace(/[^a-zA-Z0-9-_ ]/g, '_').substring(0, 30);
          const fileName = hasAnyData
            ? `DCS_Data_${safeProjectName}_${timestamp}.xlsx`
            : `DCS_Template_${safeProjectName}_${timestamp}.xlsx`;

          XLSX.writeFile(wb, fileName);

          const message = hasAnyData
            ? 'Current data downloaded. Edit in Excel and re-upload.'
            : 'Template downloaded. Fill data and upload.';

          toast(message, 'success');

        } catch (error: any) {
          console.error('Download error:', error);
          toast(`Failed to download: ${error.message}`, 'error');
        }
      };

      // Show import progress
      const showImportProgress = (message: string): void => {
        const existing = document.querySelector('.dcs-import-progress');
        if (existing) existing.remove();

        const progressEl = document.createElement("div");
        progressEl.className = "dcs-import-progress";
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

      // Hide import progress
      const hideImportProgress = (): void => {
        const existing = document.querySelector('.dcs-import-progress');
        if (existing) existing.remove();
      };

      const initializeSharePoint = async (): Promise<void> => {
        try {
          const listRes = await fetch(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
            {
              headers: { Accept: "application/json;odata=verbose" },
              credentials: 'same-origin'
            }
          );

          if (listRes.ok) {
            const listInfo = await listRes.json();
            listEntityTypeName = listInfo.d.ListItemEntityTypeFullName;
            console.log('✅ SharePoint initialized. List item type:', listEntityTypeName);
          } else {
            throw new Error(`List '${listName}' not found or inaccessible`);
          }

          // Get fields to determine correct names
          const fieldsRes = await fetch(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false`,
            {
              headers: { Accept: "application/json;odata=verbose" },
              credentials: 'same-origin'
            }
          );

          if (fieldsRes.ok) {
            const fieldsData = await fieldsRes.json();
            const fields = fieldsData.d.results;

            // Find DCSData field (Note type)
            const dcsDataField = fields.find((f: any) =>
              f.TypeAsString === "Note" &&
              (f.InternalName.includes("DCS") || f.Title.includes("DCS") || f.InternalName === "DCSData")
            ) || fields.find((f: any) => f.TypeAsString === "Note");

            // Find Project field (Lookup type)
            const projectField = fields.find((f: any) =>
              f.TypeAsString.includes("Lookup") &&
              (f.InternalName.includes("Project") || f.Title.includes("Project"))
            ) || fields.find((f: any) => f.TypeAsString.includes("Lookup"));

            fieldMappings.DCSData = dcsDataField?.InternalName || "DCSData";
            fieldMappings.Project = projectField?.InternalName || "Project";

            // If Project field ends with "Id", remove it for expansion
            if (fieldMappings.Project.endsWith("Id")) {
              fieldMappings.Project = fieldMappings.Project.replace(/Id$/, '');
            }

            console.log('✅ Field mappings:', fieldMappings);
          }

        } catch (error) {
          console.error("Error initializing SharePoint:", error);

          // Use safe defaults
          fieldMappings = {
            DCSData: "DCSData",
            Project: "Project",
          };
          listEntityTypeName = "SP.Data.DcsTrackerItem";

          console.log('⚠️ Using fallback values:', { fieldMappings, listEntityTypeName });
        }
      };

      const resetDocumentTables = (): void => {
        // Reset input table
        const defaultInputDoc: DCSDocumentItem = {
          docId: `default_input_1`,
          docNo: "DOC-001",
          description: "",
          category: "A",
          enggDivision: "",
          revNo: 0,
          originalPlan: "",
          revisedPlan: "",
          actualSubmission: "",
          submissionTransmittal: "",
          approvalDate: "",
          approvalTransmittal: "",
          status: "Pending",
          remarks: "",
          isInput: true,
          revisions: [],
          docType: 'input'
        };

        const inputRow = createEditableDocumentRow(defaultInputDoc, 1, true);
        inputTableBody.innerHTML = '';
        inputTableBody.appendChild(inputRow);

        // Reset output table
        const defaultOutputDoc: DCSDocumentItem = {
          docId: `default_output_1`,
          docNo: "DOC-OUT-001",
          description: "",
          category: "A",
          enggDivision: "",
          revNo: 0,
          originalPlan: "",
          revisedPlan: "",
          actualSubmission: "",
          submissionTransmittal: "",
          approvalDate: "",
          approvalTransmittal: "",
          status: "Not Submitted",
          remarks: "",
          isInput: false,
          revisions: [],
          docType: 'output'
        };

        const outputRow = createEditableDocumentRow(defaultOutputDoc, 1, false);
        outputTableBody.innerHTML = '';
        outputTableBody.appendChild(outputRow);

        // Ensure buttons are visible
        setTimeout(() => {
          ensureDocumentButtonsVisible();
          attachDocumentEventListeners();
        }, 100);
      };

      const updateAllSerialNumbers = (): void => {
        // Update input table
        const inputTable = document.getElementById('dcsInputTableBody');
        if (inputTable) {
          const inputDocs = inputTable.querySelectorAll('.dcs-document-header');
          inputDocs.forEach((doc, index) => {
            const srNoCell = doc.querySelector('.dcs-sr-no');
            if (srNoCell) {
              srNoCell.textContent = (index + 1).toString();
            }

            const docNoInput = doc.querySelector('.dcs-doc-no') as HTMLTextAreaElement;
            if (docNoInput) {
              const currentValue = docNoInput.value.trim();
              const expectedPattern = new RegExp(`^DOC-\\d{3}$`);
              if (!currentValue || expectedPattern.test(currentValue)) {
                docNoInput.value = `DOC-${String(index + 1).padStart(3, '0')}`;
              }
            }
          });
        }

        // Update output table
        const outputTable = document.getElementById('dcsOutputTableBody');
        if (outputTable) {
          const outputDocs = outputTable.querySelectorAll('.dcs-document-header');
          outputDocs.forEach((doc, index) => {
            const srNoCell = doc.querySelector('.dcs-sr-no');
            if (srNoCell) {
              srNoCell.textContent = (index + 1).toString();
            }

            const docNoInput = doc.querySelector('.dcs-doc-no') as HTMLTextAreaElement;
            if (docNoInput) {
              const currentValue = docNoInput.value.trim();
              const expectedPattern = new RegExp(`^DOC-OUT-\\d{3}$`);
              if (!currentValue || expectedPattern.test(currentValue)) {
                docNoInput.value = `DOC-OUT-${String(index + 1).padStart(3, '0')}`;
              }
            }
          });
        }
      };

      // ✅ (removed) the duplicate bottom-of-file event wiring block
      // (showFormBtn / backBtn / projectSelect / filters / exportExcelBtn handlers
      // and the standalone initializeEventListeners() call). All listeners are
      // attached exactly once, inside initApp() -> initializeEventListeners() +
      // setupPaginationEventListeners().

      // Initialize app
      await initApp();

      // Measure the outer page-header so the table header pins just below
      // it, then re-measure once fonts / layout have settled.
      updateMainTableStickyOffset();
      if (document.readyState === "complete") {
        setTimeout(updateMainTableStickyOffset, 100);
      } else {
        window.addEventListener("load", () => setTimeout(updateMainTableStickyOffset, 100), { once: true });
      }

    }, 100);
  }
};

// Export as default 
export default DCSModule;

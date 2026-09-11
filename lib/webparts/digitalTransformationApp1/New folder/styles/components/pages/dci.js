import "../styles/common.css";
import "../styles/dcs.css";
import * as XLSX from "xlsx-js-style";
const dcsStyles = `
/* Role-based styling */
.role-indicator {
    display: none !important;
}
.role-indicator {
  padding: 8px 15px;
  margin-bottom: 15px;
  border-radius: 5px;
  color: white;
  font-weight: bold;
  display: inline-block;
}

.role-badge {
  display: flex;
  align-items: center;
  gap: 8px;
}

.hidden-by-role {
  display: none !important;
}

/* Pagination styles */
.dcs-pagination-container {
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

.dcs-pagination-left {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}

.dcs-pagination-right {
  display: flex;
  align-items: center;
  gap: 20px;
  flex-wrap: wrap;
}

.dcs-page-size-selector {
  display: flex;
  align-items: center;
  gap: 8px;
}

.dcs-page-size-select {
  padding: 5px 10px;
  border: 1px solid #ced4da;
  border-radius: 4px;
  background-color: white;
  font-size: 14px;
  cursor: pointer;
}

.dcs-pagination-btn {
  padding: 6px 12px;
  background-color: #007bff;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  transition: background-color 0.2s;
}

.dcs-pagination-btn:hover:not(:disabled) {
  background-color: #0056b3;
}

.dcs-pagination-btn:disabled {
  background-color: #6c757d;
  cursor: not-allowed;
  opacity: 0.6;
}

.dcs-pagination-info {
  font-size: 14px;
  color: #495057;
  white-space: nowrap;
}

.dcs-pagination-pages {
  display: flex;
  gap: 5px;
}

.dcs-page-number {
  padding: 6px 12px;
  background-color: white;
  border: 1px solid #dee2e6;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  transition: all 0.2s;
}

.dcs-page-number:hover:not(.active) {
  background-color: #f8f9fa;
  border-color: #007bff;
}

.dcs-page-number.active {
  background-color: #007bff;
  color: white;
  border-color: #007bff;
}
  /* Add this to your existing dcsStyles */
.hidden-by-role {
  display: none !important;
}
`;
const dcsContent = `
<div class="dcs-page fade-in">
  <section class="table-section dcs-card" id="dcsTableSection">
    <div class="dcs-page-header">
      <h1>Document Control System</h1>
      <div>
        <button class="dcs-btn-primary dcs-show-form-btn">
          <span style="font-size: 18px; margin-right: 5px;">+</span> Add Project Documents
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
      <!-- REMOVED: hidden-by-role class - let JavaScript control visibility -->
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
        <label style="visibility: hidden;">Actions</label>
        <div style="display: flex; gap: 10px;">
          <button class="dcs-btn-secondary dcs-clear-filters-btn">Reset Filter</button> 
          <button class="dcs-btn-primary dcs-export-excel-btn">
            <span style="font-size: 16px; margin-right: 5px;">📊</span> Export Excel
          </button>
        </div>
      </div>
    </div>

    <div style="overflow-x:auto;">
      <table id="dcsTable" class="dcs-modern-table">
        <thead>
          <tr>
            <th>Project Reference</th>
            <th>Project</th>
            <th>Customer</th>
            <!-- REMOVED: hidden-by-role class - let JavaScript control visibility -->
            <th>Project Manager</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody id="dcsTableBody">
          <!-- ... existing code ... -->
        </tbody>
      </table>
    </div>
   <div class="dcs-pagination-container">
  <div class="dcs-pagination-left">
    <button class="dcs-pagination-btn" id="firstPage" disabled>First</button>
    <button class="dcs-pagination-btn" id="prevPage" disabled>Previous</button>
    
    <div class="dcs-pagination-info">
      Page <span id="currentPage">1</span> of <span id="totalPages">1</span>
    </div>
    
    <div class="dcs-pagination-pages" id="pageNumbers"></div>
    
    <button class="dcs-pagination-btn" id="nextPage" disabled>Next</button>
    <button class="dcs-pagination-btn" id="lastPage" disabled>Last</button>
  </div>
  
  <div class="dcs-pagination-right">
    <div class="dcs-page-size-selector">
      <label for="pageSizeSelect">Show:</label>
      <select id="pageSizeSelect" class="dcs-page-size-select">
        <option value="10">10</option>
        <option value="25">25</option>
        <option value="50">50</option>
        <option value="100">100</option>
      </select>
      <span>per page</span>
    </div>
    
    <div class="dcs-pagination-info">
      Showing <span id="startItem">0</span>-<span id="endItem">0</span> of <span id="totalItems">0</span> records
    </div>
  </div>
</div>
  </section>

  <section class="form-section dcs-card" id="dcsFormSection" style="display:none;">
    <div class="dcs-page-header">
      <h1 id="dcsFormTitle">Add Project Documents</h1>
    </div>
    <form id="dcsForm">
          
      <div class="dcs-form-row">
        <div class="dcs-form-group dcs-form-group-box dcs-project-select-container">
          <label>Project</label>
          <select id="dcsProjectSelect" class="dcs-form-control" required>
            <option value="">Select Project</option>
          </select> 
        </div>
        <div id="dcsProjectDetails" class="dcs-form-group dcs-form-group-box" style="flex: 2;">
          <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
            <div>
              <label style="font-size: 12px; color: #6c757d;">Profit Center</label>
              <p style="margin: 5px 0 0 0; font-weight: 500;" id="projectPC">-</p>
            </div>
            <div>
              <label style="font-size: 12px; color: #6c757d;">Customer</label>
              <p style="margin: 5px 0 0 0; font-weight: 500;" id="projectCustomer">-</p>
            </div>
            <div>
              <label style="font-size: 12px; color: #6c757d;">Project Reference</label>
              <p style="margin: 5px 0 0 0; font-weight: 500;" id="projectReference">-</p>
            </div>
            <div>
              <label style="font-size: 12px; color: #6c757d;">Group Head</label>
              <p style="margin: 5px 0 0 0; font-weight: 500;" id="projectGroupHead">-</p>
            </div>
            <div>
              <label style="font-size: 12px; color: #6c757d;">Project Manager</label>
              <p style="margin: 5px 0 0 0; font-weight: 500;" id="projectManager">-</p>
            </div>
          </div>
        </div>
      </div>
      
      <div id="dcsExistingDataAlert" class="dcs-alert dcs-alert-info" style="display: none; margin: 15px 0;">
        <div style="display: flex; align-items: center; gap: 10px;">
          <span style="font-size: 18px;">ℹ️</span>
          <div>
            <strong>Existing Data Found!</strong> This project already has document data. You can view and edit it below.
          </div>
        </div>
      </div>
       
<div id="dcsImportButtonContainer" style="margin-bottom: 15px; display: none;">
  <div class="dcs-import-export-container">
    <div class="dcs-import-buttons">
      <input type="file" id="dcsImportExcelInput" accept=".xlsx,.xls" style="display: none;">
      <button type="button" class="dcs-btn-primary dcs-import-excel-btn">
        <span style="font-size: 16px; margin-right: 5px;">📥</span> Import Excel Data
      </button>
      <button type="button" class="dcs-btn-secondary dcs-export-project-excel-btn">
        <span style="font-size: 16px; margin-right: 5px;">📊</span> Export Template Excel
      </button>
      <button type="button" class="dcs-btn-secondary dcs-download-existing-btn">
        <span style="font-size: 16px; margin-right: 5px;">📋</span> Download Current Data
      </button>
      <button type="button" class="dcs-btn-secondary" id="dcsResetFormBtn">
        <span style="font-size: 16px; margin-right: 5px;">🔄</span> Reset Form
      </button>
    </div>
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
          </ul>
        </li>
        <li><strong>Step 3:</strong> Save the Excel file and click "Import Excel Data" to upload</li>
        <li><strong>Note:</strong> Document numbers will be auto-generated if left blank</li>
      </ul>
    </div>
  </div>
</div>
      
      <div class="dcs-inputs" id="dcsContainerTable" style="overflow-x:auto; display:none;">
        
        <!-- Design Output Section -->
        <div class="dcs-section-header">
          <h3>DESIGN OUTPUT- FROM ABB</h3>
          <button type="button" class="dcs-btn-secondary dcs-add-output-document">
            <span style="font-size: 16px; margin-right: 5px;">+</span> Add Output Document
          </button>
        </div>
        <table class="dcs-modern-table dcs-output-table dcs-revision-table">
          <thead>
            <tr>
              <th rowspan="2" style="width: 40px;">Sr No</th>
              <th rowspan="2" style="min-width: 150px;">Document No</th>
              <th rowspan="2" style="min-width: 180px;">Description</th>
              <th rowspan="2" style="width: 100px;">Category</th>
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
                  <option value="Pending" selected>Pending</option>
                  <option value="Submitted">Submitted</option>
                  <option value="Under Review">Under Review</option>
                  <option value="Approved">Approved</option>
                  <option value="Rejected">Rejected</option>
                  <option value="Revised">Revised</option>
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
            <span style="font-size: 16px; margin-right: 5px;">+</span> Add Input Document
          </button>
        </div>
        <table class="dcs-modern-table dcs-input-table dcs-revision-table">
          <thead>
            <tr>
              <th rowspan="2" style="width: 40px;">Sr No</th>
              <th rowspan="2" style="min-width: 150px;">Document No</th>
              <th rowspan="2" style="min-width: 180px;">Description</th>
              <th rowspan="2" style="width: 100px;">Category</th>
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
                  <option value="Pending" selected>Pending</option>
                  <option value="Submitted">Submitted</option>
                  <option value="Under Review">Under Review</option>
                  <option value="Approved">Approved</option>
                  <option value="Rejected">Rejected</option>
                  <option value="Revised">Revised</option>
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
          <span style="font-size: 16px; margin-right: 5px;">💾</span> Save Document Data
        </button>
        <button type="button" class="dcs-btn-secondary dcs-back-to-table">
          <span style="font-size: 16px; margin-right: 5px;">⬅</span> Back to List
        </button>
      </div>
    </form>
  </section>
</div>
`;
const DCSModule = {
    content: dcsContent,
    style: dcsStyles,
    script: () => {
        setTimeout(async () => {
            // Configuration
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/DigitalTransformationTest";
            const listName = "DCSTracker";
            const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
            // State variables
            // State variables
            let fieldMappings = {
                DCSData: "DCSData",
                Project: "Project",
            };
            let listEntityTypeName = "";
            let projects = [];
            let editingId = null; // ← ADD THIS
            let dcsData = [];
            // Toast function with proper typing
            const toast = (message, type = 'success') => {
                alert(`${type === 'success' ? '✅' : type === 'error' ? '❌' : '⚠️'} ${message}`);
            };
            // DOM Elements
            const projectSelect = document.getElementById("dcsProjectSelect");
            const showFormBtn = document.querySelector(".dcs-show-form-btn");
            const exportExcelBtn = document.querySelector(".dcs-export-excel-btn");
            const formSection = document.getElementById("dcsFormSection");
            const tableSection = document.getElementById("dcsTableSection");
            const backBtn = document.querySelector(".dcs-back-to-table");
            const projectDetails = document.getElementById("dcsProjectDetails");
            const outputTableBody = document.getElementById("dcsOutputTableBody");
            const inputTableBody = document.getElementById("dcsInputTableBody");
            const containerTable = document.getElementById("dcsContainerTable");
            // const paginationDiv = document.querySelector(".dcs-pagination") as HTMLElement;
            const dcsForm = document.getElementById("dcsForm");
            const formTitle = document.getElementById("dcsFormTitle");
            const existingDataAlert = document.getElementById("dcsExistingDataAlert");
            const addInputDocBtn = document.querySelector(".dcs-add-input-document");
            const addOutputDocBtn = document.querySelector(".dcs-add-output-document");
            const importButtonContainer = document.getElementById("dcsImportButtonContainer");
            // const importExcelBtn = document.querySelector(".dcs-import-excel-btn") as HTMLButtonElement;
            // const exportProjectExcelBtn = document.querySelector(".dcs-export-project-excel-btn") as HTMLButtonElement;
            const resetFormBtn = document.getElementById("dcsResetFormBtn");
            // Filter Elements
            const filterProject = document.getElementById("filterProject");
            const filterCustomer = document.getElementById("filterCustomer");
            const filterProjectManager = document.getElementById("filterProjectManager");
            const filterProjectReference = document.getElementById("filterProjectReference");
            const clearFiltersBtn = document.querySelector(".dcs-clear-filters-btn");
            const projectPC = document.getElementById("projectPC");
            const projectCustomer = document.getElementById("projectCustomer");
            const projectReference = document.getElementById("projectReference");
            const projectGroupHead = document.getElementById("projectGroupHead");
            const projectManager = document.getElementById("projectManager");
            // Add these variables after other state variables
            let currentUser;
            let userGroups = [];
            let managedProjects = [];
            let managedGroups = [];
            // Add pagination variables
            let currentPage = 1;
            let pageSize = 10;
            let totalPages = 1;
            let filteredDCSData = [];
            const getUserContext = async () => {
                try {
                    // Get current user information
                    const userRes = await fetch(`${siteUrl}/_api/web/currentuser`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const userData = await userRes.json();
                    const loginName = userData.d.LoginName;
                    const displayName = userData.d.Title;
                    console.log(`👤 Current User: ${displayName} (${loginName})`);
                    // Get user's groups
                    const groupsRes = await fetch(`${siteUrl}/_api/web/currentuser/groups`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const groupsData = await groupsRes.json();
                    userGroups = groupsData.d.results.map((g) => g.Title);
                    console.log(`👥 User Groups:`, userGroups);
                    // Check roles based on group membership
                    const isSuperAdmin = userGroups.some(g => g.toLowerCase().includes("super admin") ||
                        g.toLowerCase() === "site owners" ||
                        g === "Super Admin");
                    const isGroupHead = userGroups.some(g => g.toLowerCase().includes("group head") ||
                        g === "Group Head");
                    const isProjectManager = userGroups.some(g => g.toLowerCase().includes("project manager") ||
                        g === "Project Manager");
                    console.log(`🎭 User Roles: SA=${isSuperAdmin}, GH=${isGroupHead}, PM=${isProjectManager}`);
                    // Reset managed arrays
                    managedGroups = [];
                    managedProjects = [];
                    // Group Head: get managed groups
                    if (isGroupHead) {
                        managedGroups = [displayName];
                        console.log(`📋 Group Head managed groups:`, managedGroups);
                    }
                    // Project Manager: get assigned projects
                    if (isProjectManager) {
                        console.log(`🔍 Fetching projects for Project Manager: ${displayName}`);
                        try {
                            const allProjectsRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('Project List')/items?$select=Id,Title,ProjectManager/Title,ProjectManager/Name&$expand=ProjectManager&$top=5000`, { headers: { Accept: "application/json;odata=verbose" } });
                            if (allProjectsRes.ok) {
                                const allProjectsData = await allProjectsRes.json();
                                const allProjects = allProjectsData.d.results || [];
                                console.log(`📊 Total projects in system: ${allProjects.length}`);
                                // Find projects assigned to current user
                                const assignedProjects = allProjects.filter((project) => {
                                    if (!project.ProjectManager) {
                                        console.log(`   ❌ Project "${project.Title}" has no Project Manager`);
                                        return false;
                                    }
                                    const managerDisplayName = project.ProjectManager.Title;
                                    const managerLoginName = project.ProjectManager.Name;
                                    console.log(`   🔍 Checking project: "${project.Title}"`);
                                    console.log(`       Manager Display Name: "${managerDisplayName}"`);
                                    console.log(`       Manager Login Name: "${managerLoginName}"`);
                                    console.log(`       Current User Display Name: "${displayName}"`);
                                    // Try multiple matching strategies
                                    const exactMatch = managerDisplayName === displayName;
                                    const loginMatch = managerLoginName === loginName;
                                    const containsMatch = managerDisplayName.includes(displayName) ||
                                        displayName.includes(managerDisplayName);
                                    console.log(`       Exact match: ${exactMatch}`);
                                    console.log(`       Login match: ${loginMatch}`);
                                    console.log(`       Contains match: ${containsMatch}`);
                                    const isMatch = exactMatch || loginMatch || containsMatch;
                                    console.log(`       Is assigned to current user: ${isMatch}`);
                                    return isMatch;
                                });
                                // Store project titles
                                managedProjects = assignedProjects.map((p) => p.Title);
                                console.log(`✅ Project Manager '${displayName}' assigned to ${managedProjects.length} projects:`, managedProjects);
                            }
                            else {
                                console.error(`❌ Failed to fetch projects: ${allProjectsRes.status}`);
                            }
                        }
                        catch (error) {
                            console.error("Error fetching assigned projects:", error);
                        }
                    }
                    return {
                        loginName,
                        displayName,
                        groups: userGroups,
                        isSuperAdmin,
                        isGroupHead,
                        isProjectManager,
                        userGroups,
                        managedProjects,
                        managedGroups
                    };
                }
                catch (error) {
                    console.error("Error getting user context:", error);
                    return {
                        loginName: "",
                        displayName: "",
                        groups: [],
                        isSuperAdmin: false,
                        isGroupHead: false,
                        isProjectManager: false,
                        managedProjects: [],
                        managedGroups: []
                    };
                }
            };
            const filterDataByRole = (data) => {
                console.log(`🚀 filterDataByRole called for user: ${currentUser.displayName}`);
                console.log(`Role: PM=${currentUser.isProjectManager}, GH=${currentUser.isGroupHead}, SA=${currentUser.isSuperAdmin}`);
                if (currentUser.isSuperAdmin) {
                    console.log('✅ Super Admin - returning all data');
                    return data;
                }
                const filteredData = data.filter(item => {
                    var _a, _b, _c, _d;
                    const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                    if (!projectTitle) {
                        console.log(`❌ Item ${item.Id} has no project title`);
                        return false;
                    }
                    const project = projects.find(p => p.Title === projectTitle);
                    if (!project) {
                        console.log(`❌ Project "${projectTitle}" not found in projects list`);
                        return false;
                    }
                    // Get the project manager name
                    let projectManagerName = '';
                    if (typeof project.ProjectManager === "object") {
                        projectManagerName = ((_b = project.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title) || '';
                    }
                    else {
                        projectManagerName = project.ProjectManager || '';
                    }
                    console.log(`🔍 Checking project: "${projectTitle}"`);
                    console.log(`   Project Manager: "${projectManagerName}"`);
                    console.log(`   Current User: "${currentUser.displayName}"`);
                    if (currentUser.isGroupHead) {
                        const canAccess = ((_c = currentUser.managedGroups) === null || _c === void 0 ? void 0 : _c.some(group => project.GroupHead === group)) || false;
                        console.log(`   Group Head access: ${canAccess}`);
                        return canAccess;
                    }
                    if (currentUser.isProjectManager) {
                        // CRITICAL: Compare exact names
                        const isAssigned = projectManagerName === currentUser.displayName;
                        console.log(`   Project Manager exact match: ${isAssigned} (${projectManagerName} === ${currentUser.displayName})`);
                        // Also check managedProjects list
                        const isInManagedList = ((_d = currentUser.managedProjects) === null || _d === void 0 ? void 0 : _d.includes(projectTitle)) || false;
                        console.log(`   In managedProjects list: ${isInManagedList}`);
                        const finalResult = isAssigned || isInManagedList;
                        console.log(`   FINAL RESULT for Project Manager: ${finalResult}`);
                        return finalResult;
                    }
                    console.log(`   ❌ User has no role - denying access`);
                    return false;
                });
                console.log(`📊 Filtered ${filteredData.length} of ${data.length} items`);
                console.log('✅ Filtered items:', filteredData.map(item => { var _a; return (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title; }));
                return filteredData;
            };
            // Add this test function and call it after initialization
            const testProjectManagerFilter = () => {
                console.log('🧪 TEST: Project Manager Filter Test');
                if (!currentUser.isProjectManager) {
                    console.log('❌ Current user is not a Project Manager');
                    return;
                }
                console.log(`🧪 Testing filter for: ${currentUser.displayName}`);
                // Test each project in the DCS data
                dcsData.forEach((item, index) => {
                    var _a, _b, _c;
                    const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                    const project = projects.find(p => p.Title === projectTitle);
                    if (project) {
                        const manager = typeof project.ProjectManager === "object"
                            ? (_b = project.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                            : project.ProjectManager;
                        const shouldShow = manager === currentUser.displayName;
                        const isInManaged = ((_c = currentUser.managedProjects) === null || _c === void 0 ? void 0 : _c.includes(projectTitle)) || false;
                        console.log(`${index + 1}. "${projectTitle}"`);
                        console.log(`   Manager: "${manager}"`);
                        console.log(`   Should show (exact match): ${shouldShow}`);
                        console.log(`   In managedProjects: ${isInManaged}`);
                        console.log(`   Will be shown: ${shouldShow || isInManaged}`);
                    }
                });
            };
            function createDocumentRow(tableType, srNo) {
                const docId = `doc_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
                const prefix = tableType === 'input' ? 'DOC' : 'DOC-OUT';
                const docNo = `${prefix}-${String(srNo).padStart(3, '0')}`;
                // FIX: Set correct transmittal placeholders based on table type
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
          <option value="Pending" selected>Pending</option>
          <option value="Submitted">Submitted</option>
          <option value="Under Review">Under Review</option>
          <option value="Approved">Approved</option>
          <option value="Rejected">Rejected</option>
          <option value="Revised">Revised</option>
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
            }
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
            const handleShowForm = () => {
                if (!checkCreatePermission()) {
                    toast("You don't have permission to create new documents", "error");
                    return;
                }
                // Simple show/hide
                formSection.style.display = "block";
                tableSection.style.display = "none";
                // Reset everything
                projectSelect.value = "";
                projectDetails.innerHTML = "";
                // Show default documents WITHOUT delete buttons
                inputTableBody.innerHTML = createDocumentRow('input', 1);
                outputTableBody.innerHTML = createDocumentRow('output', 1);
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
            const handleBackToTable = () => {
                // Remove view mode class - CSS will handle showing buttons again
                formSection.classList.remove("dcs-view-mode");
                // Enable project select
                projectSelect.disabled = false;
                // Show submit button
                const submitBtn = dcsForm.querySelector('button[type="submit"]');
                if (submitBtn)
                    submitBtn.style.display = "inline-flex";
                // Reset editing state
                editingId = null;
                // Show/hide sections
                formSection.style.display = 'none';
                tableSection.style.display = 'block';
                // Reset form
                resetForm();
            };
            const handleFormSubmit = async (e) => {
                e.preventDefault();
                await saveAllDataToSingleJSON();
            };
            const handleProjectSelect = (e) => {
                const selectedId = parseInt(e.target.value);
                if (selectedId) {
                    onProjectChange(selectedId);
                }
            };
            const initApp = async () => {
                var _a;
                try {
                    console.log('🚀 Initializing app...');
                    // Get user context first
                    currentUser = await getUserContext();
                    console.log('✅ User Context Loaded:', {
                        name: currentUser.displayName,
                        isPM: currentUser.isProjectManager,
                        assignedProjects: ((_a = currentUser.managedProjects) === null || _a === void 0 ? void 0 : _a.length) || 0
                    });
                    // Initialize SharePoint
                    await initializeSharePoint();
                    // Load projects
                    await loadProjects();
                    // Load DCS data
                    await fetchData();
                    console.log(`📊 Data loaded: ${projects.length} projects, ${dcsData.length} DCS records`);
                    // Debug project assignments
                    debugProjectAssignment();
                    // Test the filter
                    testProjectManagerFilter();
                    // Setup UI
                    initializeEventListeners();
                    setupPaginationEventListeners();
                    attachDocumentEventListeners();
                    updateUIForRole();
                    console.log('✅ App initialization complete');
                }
                catch (error) {
                    console.error("❌ Initialization error:", error);
                    toast("Failed to initialize application", "error");
                }
            };
            const checkCreatePermission = () => {
                if (currentUser.isSuperAdmin)
                    return true;
                if (currentUser.isGroupHead)
                    return true;
                if (currentUser.isProjectManager)
                    return true; // Project Managers can create
                return false;
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
            };
            // Add this debugging function
            const debugProjectAssignment = () => {
                console.log('🔍 DEBUG: Current User Info', {
                    displayName: currentUser.displayName,
                    isProjectManager: currentUser.isProjectManager,
                    managedProjects: currentUser.managedProjects
                });
                console.log('📋 All Projects with their Managers:');
                projects.forEach((project, index) => {
                    var _a;
                    const manager = typeof project.ProjectManager === "object"
                        ? (_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title
                        : project.ProjectManager;
                    console.log(`${index + 1}. ${project.Title} - Manager: ${manager}`);
                });
                console.log('📊 DCS Data being filtered:');
                dcsData.forEach((item, index) => {
                    var _a, _b;
                    const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                    const project = projects.find(p => p.Title === projectTitle);
                    const manager = project ? (typeof project.ProjectManager === "object"
                        ? (_b = project.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title
                        : project.ProjectManager) : 'No project found';
                    console.log(`${index + 1}. Project: ${projectTitle} - Manager: ${manager}`);
                });
            };
            // Load Projects
            const loadProjects = async () => {
                var _a;
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('Project List')/items?$select=Id,Title,CustomerName,GroupHead,PC,ProjectReference,ProjectManager/Title&$expand=ProjectManager&$top=1000`, {
                        headers: { Accept: "application/json;odata=verbose" },
                        credentials: 'same-origin'
                    });
                    if (!res.ok)
                        throw new Error(`Failed to load projects: ${res.status}`);
                    const data = await res.json();
                    const allProjects = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    // Filter out projects without a title
                    const validProjects = allProjects.filter(p => p.Title && p.Title.trim() !== "");
                    // Get unique projects by Title
                    const uniqueProjects = Array.from(new Map(validProjects.map((p) => [p.Title, p])).values());
                    projects = uniqueProjects;
                    // Populate project select in form
                    projectSelect.innerHTML =
                        '<option value="">Select Project</option>' +
                            projects
                                .map((p) => `<option value="${p.Id}">${p.Title}</option>`)
                                .join("");
                    // Populate project filter in table section
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
            const fetchData = async () => {
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
                    const allData = data.d.results || [];
                    console.log(`📊 Raw data received: ${allData.length} items`);
                    // Helper function to extract project ID
                    const getProjectId = (item) => {
                        var _a;
                        // Try multiple patterns
                        if (item[`${fieldMappings.Project}Id`]) {
                            return Number(item[`${fieldMappings.Project}Id`]);
                        }
                        if (item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]) {
                            return Number(item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]);
                        }
                        if ((_a = item[fieldMappings.Project.replace(/Id$/, '')]) === null || _a === void 0 ? void 0 : _a.Id) {
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
                            console.warn("No project ID found for item:", item.Id, "Item structure:", Object.keys(item));
                        }
                    });
                    dcsData = Array.from(uniqueProjects.values());
                    console.log(`✅ Unique projects: ${dcsData.length} (duplicates: ${duplicates.length})`);
                    // If we found duplicates, we could delete them here
                    if (duplicates.length > 0) {
                        console.log(`🗑️ Cleaning up ${duplicates.length} duplicates...`);
                        await cleanupDuplicates(duplicates);
                    }
                    // CRITICAL: Initialize with role-based filtering
                    console.log('🎭 Applying role-based filtering...');
                    // Get role-filtered data
                    const roleFilteredData = filterDataByRole(dcsData);
                    console.log(`📊 Before filtering: ${dcsData.length} items`);
                    console.log(`📊 After role filtering: ${roleFilteredData.length} items`);
                    // Debug: Show what was filtered
                    dcsData.forEach(item => {
                        var _a;
                        const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                        const project = projects.find(p => p.Title === projectTitle);
                        const isInFiltered = roleFilteredData.some(filteredItem => { var _a; return ((_a = filteredItem[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) === projectTitle; });
                        console.log(`${isInFiltered ? '✅' : '❌'} ${projectTitle} - Manager: ${project === null || project === void 0 ? void 0 : project.ProjectManager} - Show: ${isInFiltered}`);
                    });
                    // Initialize filters with accessible data only
                    initializeFilters();
                    // Initialize pagination with role-filtered data
                    initializePagination(roleFilteredData);
                    console.log('✅ fetchData completed successfully');
                }
                catch (error) {
                    console.error("❌ Error fetching data:", error);
                    const tableBody = document.getElementById("dcsTableBody");
                    if (tableBody) {
                        tableBody.innerHTML = `
        <tr>
          <td colspan="5" style="text-align:center;padding:40px;color:#dc3545;">
            <div style="font-size:48px;margin-bottom:10px;">❌</div>
            <h3>Error Loading Data</h3>
            <p>${error.message || "Please check your SharePoint configuration"}</p>
            <button class="dcs-btn-primary" onclick="location.reload()" style="margin-top:20px;">
              ↻ Retry
            </button>
          </td>
        </tr>
      `;
                    }
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
                var _a, _b, _c;
                console.log('🎛️ Initializing filters...');
                // Get role-filtered data for filter options
                const roleFilteredData = filterDataByRole(dcsData);
                console.log('📊 Role-filtered data for filters:', roleFilteredData.length, 'items');
                // Project filter - only show accessible projects
                const accessibleProjectTitles = roleFilteredData
                    .map((item) => { var _a; return (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title; })
                    .filter((title) => !!title);
                console.log('🏢 Accessible projects:', accessibleProjectTitles);
                filterProject.innerHTML = '<option value="">All Projects</option>' +
                    [...new Set(accessibleProjectTitles)]
                        .sort()
                        .map(project => `<option value="${project}">${project}</option>`)
                        .join("");
                // Customer filter - only customers from accessible projects
                const accessibleProjects = projects.filter(p => accessibleProjectTitles.includes(p.Title));
                const accessibleCustomers = accessibleProjects
                    .map(p => p.CustomerName)
                    .filter((customer) => !!customer);
                filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                    [...new Set(accessibleCustomers)]
                        .sort()
                        .map(customer => `<option value="${customer}">${customer}</option>`)
                        .join("");
                // **FIXED: Project Manager filter - show for Super Admin and Group Head ONLY**
                let accessibleManagers = [];
                // Helper function to extract manager name
                const getManagerName = (project) => {
                    if (!project.ProjectManager)
                        return undefined;
                    if (typeof project.ProjectManager === 'string') {
                        return project.ProjectManager;
                    }
                    return project.ProjectManager.Title;
                };
                if (currentUser.isSuperAdmin || currentUser.isGroupHead) {
                    // Super Admin and Group Head see all managers from accessible projects
                    accessibleManagers = projects
                        .filter(p => accessibleProjectTitles.includes(p.Title)) // Only from accessible projects
                        .map(getManagerName)
                        .filter((manager) => !!manager);
                    // Remove duplicates and sort
                    accessibleManagers = [...new Set(accessibleManagers)].sort();
                    filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                        accessibleManagers.map(manager => `<option value="${manager}">${manager}</option>`).join("");
                    // **SHOW the filter for Super Admin and Group Head**
                    const projectManagerFilterContainer = (_a = document.querySelector('#filterProjectManager')) === null || _a === void 0 ? void 0 : _a.closest('.dcs-tableColp4');
                    if (projectManagerFilterContainer) {
                        projectManagerFilterContainer.style.display = 'block';
                        projectManagerFilterContainer.classList.remove('hidden-by-role');
                    }
                }
                else if (currentUser.isProjectManager) {
                    // **HIDE the filter for Project Manager**
                    const projectManagerFilterContainer = (_b = document.querySelector('#filterProjectManager')) === null || _b === void 0 ? void 0 : _b.closest('.dcs-tableColp4');
                    if (projectManagerFilterContainer) {
                        projectManagerFilterContainer.style.display = 'none';
                        projectManagerFilterContainer.classList.add('hidden-by-role');
                    }
                    // Don't populate the filter for Project Managers
                    filterProjectManager.innerHTML = '<option value="">All Managers</option>';
                    console.log('👤 Hiding Project Manager filter for Project Manager login');
                }
                else {
                    // For other users, hide the filter
                    const projectManagerFilterContainer = (_c = document.querySelector('#filterProjectManager')) === null || _c === void 0 ? void 0 : _c.closest('.dcs-tableColp4');
                    if (projectManagerFilterContainer) {
                        projectManagerFilterContainer.style.display = 'none';
                        projectManagerFilterContainer.classList.add('hidden-by-role');
                    }
                }
                // Project Reference filter - only from accessible projects
                const accessibleReferences = accessibleProjects
                    .map(p => p.ProjectReference)
                    .filter((ref) => !!ref);
                filterProjectReference.innerHTML = '<option value="">All References</option>' +
                    [...new Set(accessibleReferences)]
                        .sort()
                        .map(ref => `<option value="${ref}">${ref}</option>`)
                        .join("");
                console.log('👥 Accessible managers for dropdown:', accessibleManagers);
                console.log('✅ Filters initialized');
            };
            // Filter Data
            const filterData = () => {
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
                // First apply regular filters
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
                console.log('📊 After regular filters:', filtered.length, 'items');
                // THEN apply role-based filtering (Project Managers only see their projects)
                filtered = filterDataByRole(filtered);
                console.log('📊 After role filtering:', filtered.length, 'items');
                // Initialize pagination with filtered data
                initializePagination(filtered);
                console.log('✅ Filtering completed');
            };
            // Clear Filters
            const clearFilters = () => {
                console.log('🔄 Clearing filters...');
                filterProject.value = "";
                filterCustomer.value = "";
                filterProjectManager.value = "";
                filterProjectReference.value = "";
                // Reset to all data with role filtering
                console.log('📊 Applying role filtering to all data...');
                const roleFilteredData = filterDataByRole(dcsData);
                console.log(`📊 Showing ${roleFilteredData.length} of ${dcsData.length} items`);
                initializePagination(roleFilteredData);
                toast("Filters cleared", "success");
                console.log('✅ Filters cleared');
            };
            const updateUIForRole = () => {
                var _a;
                const pageHeader = document.querySelector(".dcs-page-header");
                const roleIndicator = document.createElement("div");
                roleIndicator.className = "role-indicator";
                let roleText = "";
                if (currentUser.isSuperAdmin) {
                    roleText = "Super Admin - Full Access";
                    roleIndicator.style.backgroundColor = "#28a745"; // Green
                }
                else if (currentUser.isGroupHead) {
                    roleText = `Group Head - ${currentUser.displayName}`;
                    roleIndicator.style.backgroundColor = "#17a2b8"; // Blue
                }
                else if (currentUser.isProjectManager) {
                    roleText = `Project Manager - ${currentUser.displayName}`;
                    roleIndicator.style.backgroundColor = "#ffc107"; // Yellow
                }
                else {
                    roleText = "Limited Access";
                    roleIndicator.style.backgroundColor = "#6c757d"; // Gray
                }
                roleIndicator.innerHTML = `
    <span class="role-badge">👤 ${roleText}</span>
  `;
                // Insert role indicator after the header
                (_a = pageHeader.parentNode) === null || _a === void 0 ? void 0 : _a.insertBefore(roleIndicator, pageHeader.nextSibling);
                // Log role info for debugging
                console.log('📋 Current User Role:', roleText);
                console.log('🎭 Permissions:', {
                    isSuperAdmin: currentUser.isSuperAdmin,
                    isGroupHead: currentUser.isGroupHead,
                    isProjectManager: currentUser.isProjectManager
                });
                // Apply CSS classes for role-based UI elements
                setTimeout(() => {
                    applyRoleBasedUIClasses();
                }, 100);
            };
            const applyRoleBasedUIClasses = () => {
                var _a;
                console.log('🎭 Applying role-based UI classes for:', currentUser.displayName);
                console.log('Roles: SA=', currentUser.isSuperAdmin, ' GH=', currentUser.isGroupHead, ' PM=', currentUser.isProjectManager);
                // Reset all hidden-by-role classes first
                document.querySelectorAll('.hidden-by-role').forEach(element => {
                    element.classList.remove('hidden-by-role');
                });
                if (currentUser.isSuperAdmin) {
                    // Super Admin should see EVERYTHING
                    console.log('👑 Super Admin - showing all elements');
                    return;
                }
                // For Group Head and Project Manager, hide Project Manager column in table
                if (currentUser.isGroupHead || currentUser.isProjectManager) {
                    console.log('👥 Group Head/Project Manager - adjusting visibility');
                    // Hide Project Manager column in table
                    const tableHeaders = document.querySelectorAll('#dcsTable th');
                    tableHeaders.forEach(header => {
                        var _a;
                        const headerText = ((_a = header.textContent) === null || _a === void 0 ? void 0 : _a.toLowerCase()) || '';
                        if (headerText.includes('project manager') || headerText.includes('manager')) {
                            header.classList.add('hidden-by-role');
                        }
                    });
                }
                // **SPECIFICALLY for Project Manager: hide the filter dropdown**
                if (currentUser.isProjectManager) {
                    console.log('👤 Project Manager - hiding PM filter dropdown');
                    const projectManagerFilter = (_a = document.querySelector('#filterProjectManager')) === null || _a === void 0 ? void 0 : _a.closest('.dcs-tableColp4');
                    if (projectManagerFilter) {
                        projectManagerFilter.classList.add('hidden-by-role');
                    }
                }
                // **For Group Head: show the filter dropdown (already handled above, just log it)**
                if (currentUser.isGroupHead) {
                    console.log('🏢 Group Head - showing PM filter dropdown');
                }
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
            // Update the renderTable function to use pagination
            const renderTable = (data) => {
                const tableBody = document.getElementById("dcsTableBody");
                tableBody.innerHTML = "";
                if (data.length > 0) {
                    data.forEach((item) => {
                        var _a, _b;
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        const projectId = item[`${fieldMappings.Project}Id`] ||
                            (item[fieldMappings.Project] && item[fieldMappings.Project].Id);
                        const projectData = projects.find((p) => p.Id === projectId);
                        const row = document.createElement("tr");
                        // Determine if we should show Project Manager info based on role
                        const showProjectManager = !currentUser.isProjectManager;
                        row.innerHTML = `
        <td>${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "")}</td>
        <td>${escapeHtml(projectTitle)}</td>
        <td>${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "")}</td>
        ${showProjectManager ? `
          <td>${escapeHtml(typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                            ? ((_b = projectData.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title) || ""
                            : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "")}</td>
        ` : ''}
        <td class="dcs-actions-cell">
          ${checkEditPermission(item) ? `<button class="dcs-edit-btn" data-id="${item.Id}">✏️ Edit</button>` : ''}
          <button class="dcs-view-btn" data-id="${item.Id}">👁️ View</button>
        </td>
      `;
                        tableBody.appendChild(row);
                    });
                }
                else {
                    // Show appropriate message based on role
                    const message = currentUser.isProjectManager
                        ? `No projects assigned to you (${currentUser.displayName}). Contact administrator.`
                        : currentUser.isGroupHead
                            ? "No projects found in your groups"
                            : "No document data found";
                    const noDataRow = document.createElement("tr");
                    const colSpan = currentUser.isProjectManager ? 4 : 5; // Adjust colspan based on visible columns
                    noDataRow.innerHTML = `<td colspan="${colSpan}" style="text-align:center;padding:40px;color:#6c757d;">
      <div style="font-size:48px;margin-bottom:10px;">📋</div>
      <h3>${message}</h3>
    </td>`;
                    tableBody.appendChild(noDataRow);
                }
                // Add event listeners
                document.querySelectorAll(".dcs-edit-btn").forEach((btn) => {
                    btn.addEventListener("click", () => editData(Number(btn.dataset.id)));
                });
                document.querySelectorAll(".dcs-view-btn").forEach((btn) => {
                    btn.addEventListener("click", () => viewData(Number(btn.dataset.id)));
                });
            };
            // Render Pagination
            // const renderPagination = (totalPages: number): void => {
            //   paginationDiv.innerHTML = "";
            //   if (totalPages <= 1) return;
            //   for (let i = 1; i <= totalPages; i++) {
            //     const btn = document.createElement("button");
            //     btn.textContent = i.toString();
            //     btn.className = i === currentPage ? "active" : "";
            //     btn.addEventListener("click", () => {
            //       currentPage = i;
            //       renderTable(dcsData);
            //     });
            //     paginationDiv.appendChild(btn);
            //   }
            // };
            // let deldele ="";
            // Delete Data
            // const deleteData = async (id: number): Promise<void> => {
            //   if (!confirm("Are you sure you want to delete this record?")) {
            //     return;
            //   }
            //   try {
            //     const digest = await getRequestDigest();
            //     const response = await fetch(`${apiBase}(${id})`, {
            //       method: "DELETE",
            //       headers: {
            //         Accept: "application/json;odata=verbose",
            //         "X-RequestDigest": digest,
            //         "IF-MATCH": "*",
            //       },
            //       credentials: 'same-origin'
            //     });
            //     if (response.ok) {
            //       toast("Deleted successfully", "success");
            //       await fetchData();
            //     } else {
            //       throw new Error(`Delete failed: ${response.status}`);
            //     }
            //   } catch (error) {
            //     console.error("Delete error:", error);
            //     toast("Delete failed", "error");
            //   }
            // };
            // View Data Function
            // Add these helper functions
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
                // Check edit permission
                if (!checkEditPermission(record)) {
                    toast("You don't have permission to edit this record", "error");
                    return;
                }
                // REMOVE view mode class
                formSection.classList.remove("dcs-view-mode");
                // Set editing mode
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formTitle.textContent = "Edit Document Control Data";
                formSection.classList.remove("view-mode");
                projectSelect.disabled = false;
                // CRITICAL: Set the editing ID
                editingId = id;
                const projectId = record[`${fieldMappings.Project}Id`] ||
                    (record[fieldMappings.Project] && record[fieldMappings.Project].Id);
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
                        inputs.forEach((doc, index) => {
                            const newRow = createEditableDocumentRow(doc, index + 1, true);
                            inputTableBody.appendChild(newRow);
                            // Add revision rows if they exist
                            if (doc.revisions && doc.revisions.length > 0) {
                                doc.revisions.forEach((revision, revIndex) => {
                                    const submissionPlaceholder = 'Transmittal No';
                                    const approvalPlaceholder = 'Transmittal No';
                                    const revisionHTML = `
              <tr class="dcs-document-group" data-doc-id="${doc.docId}">
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
                        placeholder="${submissionPlaceholder}">${escapeHtml(revision.submissionTransmittal)}</textarea>
                </td>
                <td><input type="date" class="dcs-form-control dcs-date-input" 
                       value="${formatDateForInput(revision.approvalDate)}" 
                       placeholder="Approval Date"></td>       
                <td>
                  <textarea class="dcs-form-control dcs-transmittal-input" rows="2" 
                        placeholder="${approvalPlaceholder}">${escapeHtml(revision.approvalTransmittal)}</textarea>
                </td>
                <td colspan="7"></td>
                <td class="dcs-doc-actions">
                  <button type="button" class="dcs-btn-danger dcs-remove-revision-btn" 
                          title="Remove Revision">🗑️</button>
                </td>
              </tr>
            `;
                                    inputTableBody.insertAdjacentHTML('beforeend', revisionHTML);
                                });
                                // Update merged cells
                                updateMergedCells(doc.docId);
                            }
                        });
                    }
                    else {
                        // Add one default input document
                        const defaultInputDoc = {
                            docId: `default_input_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                            docNo: "DOC-001",
                            description: "",
                            category: "A",
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
                        outputs.forEach((doc, index) => {
                            const newRow = createEditableDocumentRow(doc, index + 1, false);
                            outputTableBody.appendChild(newRow);
                            // Add revision rows if they exist
                            if (doc.revisions && doc.revisions.length > 0) {
                                doc.revisions.forEach((revision, revIndex) => {
                                    const submissionPlaceholder = 'Transmittal No';
                                    const approvalPlaceholder = 'Transmittal No';
                                    const revisionHTML = `
              <tr class="dcs-document-group" data-doc-id="${doc.docId}">
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
                        placeholder="${submissionPlaceholder}">${escapeHtml(revision.submissionTransmittal)}</textarea>
                </td>
                <td><input type="date" class="dcs-form-control dcs-date-input" 
                       value="${formatDateForInput(revision.approvalDate)}" 
                       placeholder="Approval Date"></td>       
                <td>
                  <textarea class="dcs-form-control dcs-transmittal-input" rows="2" 
                        placeholder="${approvalPlaceholder}">${escapeHtml(revision.approvalTransmittal)}</textarea>
                </td>
                <td colspan="7"></td>
                <td class="dcs-doc-actions">
                  <button type="button" class="dcs-btn-danger dcs-remove-revision-btn" 
                          title="Remove Revision">🗑️</button>
                </td>
              </tr>
            `;
                                    outputTableBody.insertAdjacentHTML('beforeend', revisionHTML);
                                });
                                // Update merged cells
                                updateMergedCells(doc.docId);
                            }
                        });
                    }
                    else {
                        // Add one default output document
                        const defaultOutputDoc = {
                            docId: `default_output_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                            docNo: "DOC-OUT-001",
                            description: "",
                            category: "A",
                            revNo: 0,
                            originalPlan: "",
                            revisedPlan: "",
                            actualSubmission: "",
                            submissionTransmittal: "",
                            approvalDate: "",
                            approvalTransmittal: "",
                            status: "Pending",
                            remarks: "",
                            isInput: false,
                            revisions: [],
                            docType: 'output'
                        };
                        const newRow = createEditableDocumentRow(defaultOutputDoc, 1, false);
                        outputTableBody.appendChild(newRow);
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
                const submitBtn = dcsForm.querySelector('button[type="submit"]');
                if (submitBtn) {
                    submitBtn.style.display = "inline-flex";
                    submitBtn.innerHTML = '<span style="font-size: 16px; margin-right: 5px;">💾</span> Update Document Data';
                }
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
            };
            const updateFormActionsForEditMode = () => {
                const formActions = document.querySelector('.dcs-form-actions');
                if (!formActions)
                    return;
                formActions.innerHTML = `
    <div style="display: flex; gap: 10px; flex-wrap: wrap;">
      <button type="submit" class="dcs-btn-primary">
        <span style="font-size: 16px; margin-right: 5px;">💾</span> Save Changes
      </button>
      <button type="button" class="dcs-btn-secondary dcs-back-to-table">
        <span style="font-size: 16px; margin-right: 5px;">←</span> Cancel
      </button>
    </div>
  `;
                // Reattach back button listener
                const backBtn = formActions.querySelector('.dcs-back-to-table');
                if (backBtn) {
                    backBtn.addEventListener('click', handleBackToTable);
                }
            };
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
                // Create the row with 14 columns total
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
    <!-- Column 5: Rev No -->
    <td class="dcs-rev-cell">0</td>
    <!-- Column 6: Original Plan -->
    <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.originalPlan)}" placeholder="Original Plan"></td>
    <!-- Column 7: Revised Plan -->
    <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.revisedPlan)}" placeholder="Revised Plan"></td>
    <!-- Column 8: Actual Submission -->
    <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.actualSubmission)}" placeholder="Actual Submission"></td>
    <!-- Column 9: Submission Transmittal -->
    <td>
      <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${submissionPlaceholder}">${escapeHtml(doc.submissionTransmittal)}</textarea>
    </td>
    <!-- Column 10: Approval Date -->
    <td><input type="date" class="dcs-form-control dcs-date-input" value="${formatDateForInput(doc.approvalDate)}" placeholder="Approval Date"></td>
    <!-- Column 11: Approval Transmittal -->
    <td>
      <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${approvalPlaceholder}">${escapeHtml(doc.approvalTransmittal)}</textarea>
    </td>
    <!-- Column 12: Status -->
    <td>
      <select class="dcs-form-control dcs-status-select">
        <option value="Pending" ${doc.status === "Pending" ? "selected" : ""}>Pending</option>
        <option value="Submitted" ${doc.status === "Submitted" ? "selected" : ""}>Submitted</option>
        <option value="Under Review" ${doc.status === "Under Review" ? "selected" : ""}>Under Review</option>
        <option value="Approved" ${doc.status === "Approved" ? "selected" : ""}>Approved</option>
        <option value="Rejected" ${doc.status === "Rejected" ? "selected" : ""}>Rejected</option>
        <option value="Revised" ${doc.status === "Revised" ? "selected" : ""}>Revised</option>
      </select>
    </td>
    <!-- Column 13: Remarks -->
    <td>
      <textarea class="dcs-form-control dcs-remarks-input" rows="3" placeholder="Enter remarks">${escapeHtml(doc.remarks)}</textarea>
    </td>
    <!-- Column 14: Actions -->
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
                var _a;
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
                const submitBtn = dcsForm.querySelector('button[type="submit"]');
                if (submitBtn)
                    submitBtn.style.display = "none";
                // Clear tables
                inputTableBody.innerHTML = '';
                outputTableBody.innerHTML = '';
                // Update project details with the actual data
                if (projectId) {
                    const project = projects.find(p => p.Id === projectId);
                    if (project) {
                        projectPC.textContent = project.PC || "-";
                        projectCustomer.textContent = project.CustomerName || "-";
                        projectReference.textContent = project.ProjectReference || "-";
                        projectGroupHead.textContent = project.GroupHead || "-";
                        const managerName = typeof project.ProjectManager === "object"
                            ? ((_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || "-"
                            : project.ProjectManager || "-";
                        projectManager.textContent = managerName;
                        // Also update the project details box with better styling
                        const projectDetailsDiv = document.getElementById('dcsProjectDetails');
                        projectDetailsDiv.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px;">
          <div style="background: #e8f4fd; padding: 15px; border-radius: 8px; border-left: 4px solid #0d6efd;">
            <div style="font-size: 12px; color: #0d6efd; margin-bottom: 5px; font-weight: 600;">PROFIT CENTER</div>
            <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.PC || "-"}</div>
          </div>
          
          <div style="background: #d4edda; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745;">
            <div style="font-size: 12px; color: #28a745; margin-bottom: 5px; font-weight: 600;">CUSTOMER</div>
            <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.CustomerName || "-"}</div>
          </div>
          
          <div style="background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107;">
            <div style="font-size: 12px; color: #ffc107; margin-bottom: 5px; font-weight: 600;">PROJECT REFERENCE</div>
            <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.ProjectReference || "-"}</div>
          </div>
          
          <div style="background: #f8d7da; padding: 15px; border-radius: 8px; border-left: 4px solid #dc3545;">
            <div style="font-size: 12px; color: #dc3545; margin-bottom: 5px; font-weight: 600;">GROUP HEAD</div>
            <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.GroupHead || "-"}</div>
          </div>
          
          <div style="background: #d1ecf1; padding: 15px; border-radius: 8px; border-left: 4px solid #17a2b8;">
            <div style="font-size: 12px; color: #17a2b8; margin-bottom: 5px; font-weight: 600;">PROJECT MANAGER</div>
            <div style="font-size: 16px; font-weight: 600; color: #212529;">${managerName}</div>
          </div>
        </div>
         
      `;
                    }
                    else {
                        console.error("❌ Project not found in projects array for ID:", projectId);
                        toast("Project details not found", "error");
                    }
                }
                else {
                    console.error("❌ No project ID found in record");
                    toast("Project information missing", "error");
                }
                // Parse and load data into readonly fields
                const parsedData = parseDCSData(record[fieldMappings.DCSData]);
                if (parsedData) {
                    containerTable.style.display = "block";
                    // Load data into readonly form fields
                    loadDataIntoReadonlyForm(parsedData);
                }
                else {
                    containerTable.style.display = "block";
                    inputTableBody.innerHTML = `
      <tr>
        <td colspan="13" style="text-align:center; padding:40px; color:#6c757d;">
          <div style="font-size: 48px; margin-bottom: 10px;">📄</div>
          <h3>No Document Data</h3>
          <p>This project doesn't have any document data yet.</p>
        </td>
      </tr>
    `;
                    outputTableBody.innerHTML = `
      <tr>
        <td colspan="13" style="text-align:center; padding:40px; color:#6c757d;">
          <div style="font-size: 48px; margin-bottom: 10px;">📄</div>
          <h3>No Document Data</h3>
          <p>This project doesn't have any document data yet.</p>
        </td>
      </tr>
    `;
                }
                setupViewModeActions(projectId);
            };
            // Create readonly document row that matches edit mode structure exactly - FIXED (without placeholders, left aligned)
            const createReadonlyDocumentRow = (doc, srNo, tableType) => {
                const row = document.createElement("tr");
                row.className = "dcs-document-group dcs-document-header";
                row.setAttribute("data-doc-id", doc.docId);
                // Determine transmittal labels based on table type
                // let submissionLabel: string;
                // let approvalLabel: string;
                let scheduleBgColor;
                let submissionBgColor;
                let approvalBgColor;
                if (tableType === 'input') {
                    // submissionLabel = 'Transmittal No3';
                    // approvalLabel = 'Transmittal No4';
                    scheduleBgColor = '#e8f5e9'; // Green for input schedule
                    submissionBgColor = '#fff3cd'; // Yellow for submission
                    approvalBgColor = '#d4edda'; // Green for approval
                }
                else {
                    // submissionLabel = 'Transmittal No1';
                    // approvalLabel = 'Transmittal No2';
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
                        'Pending': 'background: #fff3cd; color: #856404; border: 1px solid #ffeaa7;',
                        'Submitted': 'background: #cce5ff; color: #004085; border: 1px solid #b8daff;',
                        'Under Review': 'background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb;',
                        'Approved': 'background: #d4edda; color: #155724; border: 1px solid #c3e6cb;',
                        'Rejected': 'background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb;',
                        'Revised': 'background: #e2e3e5; color: #383d41; border: 1px solid #d6d8db;'
                    };
                    return statusStyles[status] || 'background: #f8f9fa; color: #6c757d; border: 1px solid #e9ecef;';
                };
                // Create the row with 13 columns (same as edit mode but without Actions column)
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
        font-family: 'Courier New', monospace;
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
    
    <!-- Column 5: Rev No (width: 80px) -->
    <td class="dcs-rev-cell" style="width: 80px; text-align: center; vertical-align: middle; padding: 8px; background: #f8f9fa; font-weight: 500;">0</td>
    
    <!-- Column 6: Original Plan (width: 150px) -->
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
    
    <!-- Column 7: Revised Plan (width: 150px) -->
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
    
    <!-- Column 8: Actual Submission (width: 150px) -->
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
    
    <!-- Column 9: Submission Transmittal (width: 150px) -->
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
    
    <!-- Column 10: Approval Date (width: 150px) -->
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
    
    <!-- Column 11: Approval Transmittal (width: 150px) -->
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
    
    <!-- Column 12: Status (width: 120px) -->
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
    
    <!-- Column 13: Remarks (width: 150px) - Increased width since no Actions column -->
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
                // Determine transmittal labels based on table type
                // let submissionLabel: string;
                // let approvalLabel: string;
                let scheduleBgColor;
                let submissionBgColor;
                let approvalBgColor;
                if (tableType === 'input') {
                    // submissionLabel = 'Transmittal No3';
                    // approvalLabel = 'Transmittal No4';
                    scheduleBgColor = '#e8f5e9'; // Green for input schedule
                    submissionBgColor = '#fff3cd'; // Yellow for submission
                    approvalBgColor = '#d4edda'; // Green for approval
                }
                else {
                    // submissionLabel = 'Transmittal No1';
                    // approvalLabel = 'Transmittal No2';
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
                // Revision rows in edit mode only have 7 columns (Rev No + 6 date/transmittal columns + empty columns)
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
                const inputs = getInputsFromProjectData(projectData);
                const outputs = getOutputsFromProjectData(projectData);
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
                            const cellsToMerge = [0, 1, 2, 3, 11, 12]; // 0-indexed (13 columns total, index 0-12)
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
        <td colspan="13" style="text-align:center; padding:20px; color:#6c757d; background:#f8f9fa;">
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
                            const cellsToMerge = [0, 1, 2, 3, 11, 12]; // 0-indexed (13 columns total, index 0-12)
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
        <td colspan="13" style="text-align:center; padding:20px; color:#6c757d; background:#f8f9fa;">
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
        <span style="margin-right: 5px;">←</span> Back to List
      </button>
      
      <button type="button" class="dcs-btn-primary dcs-export-document-btn" 
              id="viewModeExportBtn">
        <span style="margin-right: 5px;">📤</span> Export Document
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
                        // Empty rows as spacer (removed 3 rows from original 4)
                        ["", "", "", "", "", "", "", "", "", "", ""],
                        // Rows 13, 14, 15 removed - only keeping 1 spacer row
                    ];
                    // Add approval block
                    coverSheetData.push(...approvalBlockData);
                    // ========== SECTION 2: PROJECT INFORMATION ==========
                    // Get dynamic project fields
                    const abbProjectNo = project.ProjectReference || "";
                    const abbDocumentNo = "";
                    const prPoNo = "";
                    // Project information rows - Labels merged A-B, values merged C-K
                    coverSheetData.push(["PROJECT TITLE:", "", project.Title || "", "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["DOCUMENT TITLE:", "", "DOCUMENT CONTROL INDEX", "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["ABB PROJECT NO:", "", abbProjectNo, "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["ABB DOCUMENT NO:", "", abbDocumentNo, "", "", "", "", "", "", "", ""]);
                    coverSheetData.push(["PR / PO NO:", "", prPoNo, "", "", "", "", "", "", "", ""]);
                    // Empty rows after project info (3 rows) - removed 1 row from original 3
                    for (let i = 0; i < 2; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    }
                    // ========== SECTION 3: PMC, CUSTOMER, END USER, SUPPLIER ==========
                    // ADDED PMC SECTION
                    const pmc = ""; // You can add a field for PMC if needed
                    // Customer Section - 6 rows total  
                    coverSheetData.push(["PMC:", "", pmc, "", "", "", "", "", "", "", ""]); // Row 1: Title
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]); // Rows 2-6: Empty
                    }
                    // Empty row between PMC and CUSTOMER
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    const customerName = project.CustomerName || "";
                    const endUser = "";
                    const supplier = "";
                    // Customer Section - 6 rows total  
                    coverSheetData.push(["CUSTOMER:", "", customerName, "", "", "", "", "", "", "", ""]); // Row 1: Title
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]); // Rows 2-6: Empty
                    }
                    // Empty row between CUSTOMER and END USER
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    // End User Section - 6 rows total
                    coverSheetData.push(["END USER:", "", endUser, "", "", "", "", "", "", "", ""]); // Row 1: Title
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]); // Rows 2-6: Empty
                    }
                    // Empty row between END USER and SUPPLIER
                    coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]);
                    // Supplier Section - 6 rows total
                    coverSheetData.push(["SUPPLIER:", "", supplier, "", "", "", "", "", "", "", ""]); // Row 1: Title
                    for (let i = 0; i < 5; i++) {
                        coverSheetData.push(["", "", "", "", "", "", "", "", "", "", ""]); // Rows 2-6: Empty
                    }
                    // Create Cover Sheet worksheet
                    const coverSheetWs = XLSX.utils.aoa_to_sheet(coverSheetData);
                    // ========== STYLING ==========
                    // Define border styles - SIMPLIFIED for SheetJS compatibility
                    const borderStyle = {
                        top: { style: "thin" },
                        bottom: { style: "thin" },
                        left: { style: "thin" },
                        right: { style: "thin" }
                    };
                    // Get the full data range
                    const fullDataRange = XLSX.utils.decode_range(coverSheetWs["!ref"] || "A1");
                    for (let R = 0; R < 3; ++R) {
                        for (let C = 0; C <= 10; ++C) {
                            const address = XLSX.utils.encode_cell({ r: R, c: C });
                            // Create cell if it doesn't exist
                            if (!coverSheetWs[address]) {
                                coverSheetWs[address] = { t: 's', v: '' };
                            }
                            if (!coverSheetWs[address].s)
                                coverSheetWs[address].s = {};
                            coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "FFFFFF" } };
                        }
                    }
                    // Apply borders and styling to ALL cells
                    for (let R = 0; R <= fullDataRange.e.r; ++R) {
                        for (let C = 0; C <= 10; ++C) { // Columns A-K (0-10)
                            const address = XLSX.utils.encode_cell({ r: R, c: C });
                            // Create cell if it doesn't exist
                            if (!coverSheetWs[address]) {
                                coverSheetWs[address] = { t: 's', v: '' };
                            }
                            if (!coverSheetWs[address].s)
                                coverSheetWs[address].s = {};
                            // Apply border to all cells
                            coverSheetWs[address].s.border = borderStyle;
                            // Apply specific styling
                            const cellValue = coverSheetWs[address].v;
                            // Header row (row 11, 0-indexed) - ADDED +3 OFFSET
                            if (R === 14) {
                                coverSheetWs[address].s.font = { bold: true, sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "center", vertical: "center" };
                                coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "D9D9D9" } };
                            }
                            // "0 For Approval" row (row 9) - ADDED +3 OFFSET
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
                            // NAME and DATE cells in approval block (rows 0-10) - ADDED +3 OFFSET
                            if (R >= 3 && R < 13 && (cellValue === "NAME" || cellValue === "DATE")) {
                                coverSheetWs[address].s.alignment = { horizontal: "left", vertical: "center" };
                                coverSheetWs[address].s.font = { sz: 11 };
                            }
                            // Original rows 15-19 become rows 12-16 after removing rows 13,14,15
                            if (R >= 15 && R < 20 && C === 0 && cellValue) {
                                coverSheetWs[address].s.font = { bold: true, sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "left", vertical: "center" };
                                coverSheetWs[address].s.fill = { patternType: "solid", fgColor: { rgb: "F2F2F2" } };
                            }
                            // Project information values (original rows 15-19, now adjusted) - ADDED +3 OFFSET
                            if (R >= 15 && R < 20 && C === 0 && cellValue) {
                                coverSheetWs[address].s.font = { sz: 11 };
                                coverSheetWs[address].s.alignment = { horizontal: "left", vertical: "center" };
                            }
                        }
                    }
                    const lastRow = 48; // Row 49 in 0-index is 48
                    const lastCol = 10; // Column K is 10 (A=0, B=1, ..., K=10)
                    // Define dark gray color for outer border
                    const darkGrayBorder = {
                        top: { style: "medium", color: { rgb: "404040" } },
                        bottom: { style: "medium", color: { rgb: "404040" } },
                        left: { style: "medium", color: { rgb: "404040" } },
                        right: { style: "medium", color: { rgb: "404040" } }
                    };
                    // Top border (Row 0, all columns A-K)
                    for (let C = 0; C <= lastCol; C++) {
                        const address = XLSX.utils.encode_cell({ r: 0, c: C });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        // Only keep top border for outer edge
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { top: darkGrayBorder.top });
                    }
                    // Bottom border (Row 48, all columns A-K)
                    for (let C = 0; C <= lastCol; C++) {
                        const address = XLSX.utils.encode_cell({ r: lastRow, c: C });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        // Only keep bottom border for outer edge
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { bottom: darkGrayBorder.bottom });
                    }
                    // Left border (Column A, all rows 0-48)
                    for (let R = 0; R <= lastRow; R++) {
                        const address = XLSX.utils.encode_cell({ r: R, c: 0 });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        // Only keep left border for outer edge
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { left: darkGrayBorder.left });
                    }
                    // Right border (Column K, all rows 0-48)
                    for (let R = 0; R <= lastRow; R++) {
                        const address = XLSX.utils.encode_cell({ r: R, c: lastCol });
                        if (!coverSheetWs[address])
                            coverSheetWs[address] = { t: 's', v: '' };
                        if (!coverSheetWs[address].s)
                            coverSheetWs[address].s = {};
                        // Only keep right border for outer edge
                        coverSheetWs[address].s.border = Object.assign(Object.assign({}, coverSheetWs[address].s.border), { right: darkGrayBorder.right });
                    }
                    // Corners - ensure all borders are applied
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
                    // Add merges for NAME/DATE rows - SHIFTED +3
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
                    // Apply gray background to separator rows
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
                // First row: Main headers (will be merged for groups)
                const headerRow1 = [
                    "Sr No", "Document No", "Description", "Category", "Rev No",
                    "Document Schedule", "Document Schedule", // Will merge columns F-G
                    "Document Submission", "Document Submission", // Will merge columns H-I
                    "Document Approval", "Document Approval", // Will merge columns J-K
                    "Status", "Remarks"
                ];
                // Second row: Sub-headers (for Document Schedule, Submission, Approval groups)
                const headerRow2 = [
                    "", "", "", "", "", // Sr No to Rev No (empty because merged with row above)
                    "Original Plan", "Revised Plan", // Under Document Schedule
                    "Actual Submission", "Transmittal No", // Under Document Submission
                    "Date", "Transmittal No", // Under Document Approval
                    "", "" // Status and Remarks (empty because merged with row above)
                ];
                // Create worksheet with two-level headers
                const ws = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2]);
                // Prepare data rows
                const dataRows = [];
                const merges = [];
                let currentRow = 2; // Start after headers
                // =================== OUTPUT SECTION ===================
                if (outputs.length > 0) {
                    let outputSrNo = 1; // Separate Sr No for Output section
                    // Add OUTPUT section title (LEFT ALIGNED)
                    const outputTitleRow = new Array(headerRow1.length).fill("");
                    outputTitleRow[0] = "DESIGN OUTPUT - FROM ABB";
                    dataRows.push(outputTitleRow);
                    merges.push({
                        s: { r: currentRow, c: 0 },
                        e: { r: currentRow, c: headerRow1.length - 1 }
                    });
                    currentRow++;
                    // Add each OUTPUT document with revisions
                    outputs.forEach((doc) => {
                        const docStartRow = currentRow;
                        // Main document row (Revision 0)
                        const mainRow = new Array(headerRow1.length).fill("");
                        mainRow[0] = outputSrNo; // Separate Sr No for Output
                        mainRow[1] = doc.docNo || "";
                        mainRow[2] = doc.description || "";
                        mainRow[3] = doc.category === "A" ? "A - Approval" : "I - Information";
                        mainRow[4] = "0"; // Main document Rev No
                        mainRow[5] = formatDateForExcel(doc.originalPlan); // Original Plan
                        mainRow[6] = formatDateForExcel(doc.revisedPlan); // Revised Plan
                        mainRow[7] = formatDateForExcel(doc.actualSubmission); // Actual Submission
                        mainRow[8] = doc.submissionTransmittal || ""; // Transmittal No
                        mainRow[9] = formatDateForExcel(doc.approvalDate); // Approval Date
                        mainRow[10] = doc.approvalTransmittal || ""; // Transmittal No
                        mainRow[11] = doc.status || "Pending"; // Status
                        mainRow[12] = doc.remarks || ""; // Remarks
                        dataRows.push(mainRow);
                        currentRow++;
                        outputSrNo++;
                        // Add revision rows if they exist
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach((revision) => {
                                const revRow = new Array(headerRow1.length).fill("");
                                revRow[4] = revision.revNumber; // Only Rev No column filled
                                revRow[5] = formatDateForExcel(revision.originalPlan);
                                revRow[6] = formatDateForExcel(revision.revisedPlan);
                                revRow[7] = formatDateForExcel(revision.actualSubmission);
                                revRow[8] = revision.submissionTransmittal || "";
                                revRow[9] = formatDateForExcel(revision.approvalDate);
                                revRow[10] = revision.approvalTransmittal || "";
                                dataRows.push(revRow);
                                currentRow++;
                            });
                            // Add merges for cells that should be merged across revisions
                            const mergedColumns = [0, 1, 2, 3, 11, 12]; // Sr No, Doc No, Description, Category, Status, Remarks
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
                    let inputSrNo = 1; // Separate Sr No for Input section
                    // Add INPUT section title with one empty row separator
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
                    // Add each INPUT document with revisions
                    inputs.forEach((doc) => {
                        const docStartRow = currentRow;
                        // Main document row (Revision 0)
                        const mainRow = new Array(headerRow1.length).fill("");
                        mainRow[0] = inputSrNo; // Separate Sr No for Input
                        mainRow[1] = doc.docNo || "";
                        mainRow[2] = doc.description || "";
                        mainRow[3] = doc.category === "A" ? "A - Approval" : "I - Information";
                        mainRow[4] = "0"; // Main document Rev No
                        mainRow[5] = formatDateForExcel(doc.originalPlan); // Original Plan
                        mainRow[6] = formatDateForExcel(doc.revisedPlan); // Revised Plan
                        mainRow[7] = formatDateForExcel(doc.actualSubmission); // Actual Submission
                        mainRow[8] = doc.submissionTransmittal || ""; // Transmittal No
                        mainRow[9] = formatDateForExcel(doc.approvalDate); // Approval Date
                        mainRow[10] = doc.approvalTransmittal || ""; // Transmittal No
                        mainRow[11] = doc.status || "Pending"; // Status
                        mainRow[12] = doc.remarks || ""; // Remarks
                        dataRows.push(mainRow);
                        currentRow++;
                        inputSrNo++;
                        // Add revision rows if they exist
                        if (doc.revisions && doc.revisions.length > 0) {
                            doc.revisions.forEach((revision) => {
                                const revRow = new Array(headerRow1.length).fill("");
                                revRow[4] = revision.revNumber; // Only Rev No column filled
                                revRow[5] = formatDateForExcel(revision.originalPlan);
                                revRow[6] = formatDateForExcel(revision.revisedPlan);
                                revRow[7] = formatDateForExcel(revision.actualSubmission);
                                revRow[8] = revision.submissionTransmittal || "";
                                revRow[9] = formatDateForExcel(revision.approvalDate);
                                revRow[10] = revision.approvalTransmittal || "";
                                dataRows.push(revRow);
                                currentRow++;
                            });
                            // Add merges for cells that should be merged across revisions
                            const mergedColumns = [0, 1, 2, 3, 11, 12]; // Sr No, Doc No, Description, Category, Status, Remarks
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
                const outputScheduleColor = "E3F2FD"; // Blue for Output
                const inputScheduleColor = "E8F5E9"; // Green for Input
                const submissionColor = "FFF3CD"; // Yellow
                const approvalColor = "D4EDDA"; // Green
                // Get the full data range
                const dataRange = XLSX.utils.decode_range(ws["!ref"] || "A1");
                const totalRows = dataRange.e.r;
                const totalCols = dataRange.e.c;
                // =================== HEADER MERGES AND STYLING ===================
                // Add merges for ALL two-level headers
                merges.push(
                // Merge Sr No header (rows 0-1, column A)
                { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, 
                // Merge Document No header (rows 0-1, column B)
                { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, 
                // Merge Description header (rows 0-1, column C)
                { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, 
                // Merge Category header (rows 0-1, column D)
                { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, 
                // Merge Rev No header (rows 0-1, column E)
                { s: { r: 0, c: 4 }, e: { r: 1, c: 4 } }, 
                // Merge Document Schedule main header (columns F-G, row 0)
                { s: { r: 0, c: 5 }, e: { r: 0, c: 6 } }, 
                // Merge Document Submission main header (columns H-I, row 0)
                { s: { r: 0, c: 7 }, e: { r: 0, c: 8 } }, 
                // Merge Document Approval main header (columns J-K, row 0)
                { s: { r: 0, c: 9 }, e: { r: 0, c: 10 } }, 
                // Merge Status header (rows 0-1, column L)
                { s: { r: 0, c: 11 }, e: { r: 1, c: 11 } }, 
                // Merge Remarks header (rows 0-1, column M)
                { s: { r: 0, c: 12 }, e: { r: 1, c: 12 } });
                // Apply header styling
                for (let C = 0; C < headerRow1.length; C++) {
                    // Row 0 (first header row) - ALL CENTERED VERTICALLY AND HORIZONTALLY
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
                    // Row 1 (second header row) - ALL CENTERED VERTICALLY AND HORIZONTALLY
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
                        // Initialize style if not exists
                        if (!ws[address].s)
                            ws[address].s = {};
                        // Get cell value for section detection
                        const cellValue = ws[address].v;
                        const cellValueStr = typeof cellValue === 'string' ? cellValue : String(cellValue || '');
                        // =================== SECTION TITLE ROWS ===================
                        // Check if this is a section title row
                        if (C === 0 && (cellValueStr === "DESIGN OUTPUT - FROM ABB" ||
                            cellValueStr === "DESIGN INPUT FROM CONSULTANT / CUSTOMER")) {
                            // Style the entire title row
                            for (let titleC = 0; titleC <= totalCols; titleC++) {
                                const titleAddr = XLSX.utils.encode_cell({ r: R, c: titleC });
                                if (ws[titleAddr]) {
                                    ws[titleAddr].s = {
                                        font: { bold: true, sz: 12, color: { rgb: "000000" } },
                                        alignment: { horizontal: "left", vertical: "center" }, // LEFT ALIGNED for titles
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
                        // =================== DATA ROWS ===================
                        if (R >= 3) {
                            // Find the nearest section title row above current row
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
                            // Apply background colors based on column and section
                            if (sectionType) {
                                if (C === 5 || C === 6) { // Document Schedule columns
                                    const scheduleColor = sectionType === "OUTPUT" ? outputScheduleColor : inputScheduleColor;
                                    ws[address].s.fill = { patternType: "solid", fgColor: { rgb: scheduleColor } };
                                }
                                else if (C === 7 || C === 8) { // Document Submission columns
                                    ws[address].s.fill = { patternType: "solid", fgColor: { rgb: submissionColor } };
                                }
                                else if (C === 9 || C === 10) { // Document Approval columns
                                    ws[address].s.fill = { patternType: "solid", fgColor: { rgb: approvalColor } };
                                }
                            }
                            // Apply CENTER alignment for specific columns
                            // Sr No, Rev No, and Status should be CENTERED both horizontally and vertically
                            if (C === 0 || C === 4 || C === 11) { // Sr No (0), Rev No (4), Status (11)
                                ws[address].s.alignment = { horizontal: "center", vertical: "center" };
                            }
                            // Other columns: left align text, center align vertically
                            else {
                                ws[address].s.alignment = {
                                    horizontal: C === 3 ? "center" : "left", // Center Category, left align others
                                    vertical: "center"
                                };
                            }
                            // Set borders
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
                // Set column widths
                ws["!cols"] = [
                    { wch: 8 }, // A - Sr No
                    { wch: 20 }, // B - Document No
                    { wch: 35 }, // C - Description
                    { wch: 12 }, // D - Category
                    { wch: 10 }, // E - Rev No
                    { wch: 15 }, // F - Original Plan
                    { wch: 15 }, // G - Revised Plan
                    { wch: 18 }, // H - Actual Submission
                    { wch: 20 }, // I - Transmittal No
                    { wch: 15 }, // J - Approval Date
                    { wch: 20 }, // K - Transmittal No
                    { wch: 12 }, // L - Status
                    { wch: 25 } // M - Remarks
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
      <div style="font-size: 16px; font-weight: 500; margin-bottom: 10px;">${message}</div>
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
            // Project Change Handler
            const onProjectChange = (selectedId, isEditing = false) => {
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
                const existingRecord = dcsData.find((item) => item[`${fieldMappings.Project}Id`] === selectedId ||
                    (item[fieldMappings.Project] &&
                        item[fieldMappings.Project].Id === selectedId));
                if (existingRecord && !isEditing) {
                    // If existing data found
                    existingDataAlert.style.display = "block";
                    formTitle.textContent = "Edit Document Control Data";
                    importButtonContainer.style.display = "flex";
                    const dcsProjectData = parseDCSData(existingRecord[fieldMappings.DCSData]);
                    if (dcsProjectData) {
                        loadDocumentsFromJSON(dcsProjectData);
                    }
                    else {
                        resetDocumentTables();
                    }
                }
                else {
                    // If no existing data (adding new)
                    existingDataAlert.style.display = "none";
                    formTitle.textContent = "Add Project Documents";
                    // IMPORTANT: Show import buttons even for new projects
                    // Users might want to import template data or start from an existing Excel
                    importButtonContainer.style.display = "flex";
                    resetDocumentTables();
                }
                // Re-attach document event listeners
                setTimeout(() => {
                    attachDocumentEventListeners();
                }, 100);
            };
            const addRevisionToDocument = (docId) => {
                const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
                if (!docHeader) {
                    console.error(`Document header not found for ID: ${docId}`);
                    return;
                }
                // Get table body and determine document type
                const tableBody = docHeader.closest('tbody');
                if (!tableBody)
                    return;
                // Determine if it's input or output table
                const isInputTable = tableBody.id === 'dcsInputTableBody';
                // Use correct placeholders based on table type
                let submissionPlaceholder = 'Transmittal No';
                let approvalPlaceholder = 'Transmittal No';
                if (isInputTable) {
                    submissionPlaceholder = 'Transmittal No';
                    approvalPlaceholder = 'Transmittal No';
                }
                // Get all revision rows for this document
                const existingRevRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);
                const newRevNumber = existingRevRows.length + 1;
                // Get current rowspan value
                const currentRevRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');
                const newRevRows = currentRevRows + 1;
                // Update the data-rev-rows attribute
                docHeader.setAttribute('data-rev-rows', newRevRows.toString());
                // Create revision row with proper data attributes
                const newRevisionHTML = `
    <tr class="dcs-document-group" data-doc-id="${docId}" data-is-revision="true">
      <!-- Rev No -->
      <td class="dcs-rev-cell">${newRevNumber}</td>
      
      <!-- Original Plan -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Original Plan"></td>
      
      <!-- Revised Plan -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Revised Plan"></td>
      
      <!-- Actual Submission -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Actual Submission"></td>
      
      <!-- Submission Transmittal -->
      <td>
        <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${submissionPlaceholder}"></textarea>
      </td>
      
      <!-- Approval Date -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Approval Date"></td>
      
      <!-- Approval Transmittal -->
      <td>
        <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${approvalPlaceholder}"></textarea>
      </td>
      
      <!-- Status (empty - merged from main row) -->
      <td></td>
      
      <!-- Remarks (empty - merged from main row) -->
      <td></td>
      
      <!-- Actions -->
      <td class="dcs-doc-actions">
        <button type="button" class="dcs-btn-danger dcs-remove-revision-btn" title="Remove Revision">🗑️</button>
      </td>
    </tr>
  `;
                // Insert after the last row of this document
                let lastRow = null;
                if (existingRevRows.length > 0) {
                    lastRow = existingRevRows[existingRevRows.length - 1];
                }
                else {
                    lastRow = docHeader;
                }
                lastRow.insertAdjacentHTML('afterend', newRevisionHTML);
                // Update merged cells in the header
                updateMergedCellsForDocument(docId);
                // Attach remove listener to the new revision
                const newRevision = lastRow.nextElementSibling;
                if (newRevision) {
                    const removeBtn = newRevision.querySelector('.dcs-remove-revision-btn');
                    if (removeBtn) {
                        removeBtn.addEventListener('click', function (e) {
                            e.preventDefault();
                            e.stopPropagation();
                            removeRevisionFromDocument(docId, this);
                        });
                    }
                }
            };
            // Enhanced function to update merged cells - FIXED
            const updateMergedCellsForDocument = (docId) => {
                const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
                if (!docHeader)
                    return;
                const revRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');
                // Reset all rowspans first
                const allCells = docHeader.querySelectorAll('td');
                allCells.forEach(cell => {
                    cell.removeAttribute('rowspan');
                    cell.classList.remove('dcs-merged-cell');
                });
                // Cells that should have rowspan (cover revision rows too):
                // 0: Sr No, 1: Document No, 2: Description, 3: Category, 
                // 11: Status (index 11), 12: Remarks (index 12), 13: Actions (index 13)
                const cellsToMerge = [0, 1, 2, 3, 11, 12, 13];
                cellsToMerge.forEach(cellIndex => {
                    const cell = docHeader.querySelector(`td:nth-child(${cellIndex + 1})`);
                    if (cell) {
                        cell.setAttribute('rowspan', revRows.toString());
                        cell.classList.add('dcs-merged-cell');
                    }
                });
                // Rev No cell (column 5, index 4) should NOT be merged
                const revNoCell = docHeader.querySelector('.dcs-rev-cell');
                if (revNoCell) {
                    revNoCell.removeAttribute('rowspan');
                    revNoCell.classList.remove('dcs-merged-cell');
                    revNoCell.textContent = '0'; // Main document has rev number 0
                }
            };
            const updateMergedCells = (docId) => {
                const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
                if (!docHeader)
                    return;
                const revRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');
                // Get all cells in the header row
                const cells = docHeader.querySelectorAll('td');
                // Column indices that need rowspan (0-indexed):
                // 0: Sr No, 1: Document No, 2: Description, 3: Category, 
                // 10: Status, 11: Remarks, 12: Actions
                const cellsToMerge = [0, 1, 2, 3, 11, 12, 13];
                cellsToMerge.forEach(cellIndex => {
                    if (cellIndex < cells.length) {
                        const cell = cells[cellIndex];
                        cell.setAttribute('rowspan', revRows.toString());
                        cell.classList.add('dcs-merged-cell');
                    }
                });
            };
            const removeRevisionFromDocument = (docId, buttonElement) => {
                const revisionRow = buttonElement.closest('tr[data-doc-id]');
                if (!revisionRow)
                    return;
                // Get all revision rows for this document
                const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
                if (!docHeader)
                    return;
                const allRevRows = document.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);
                if (allRevRows.length <= 1) {
                    // If removing the only revision, just remove it
                    revisionRow.remove();
                    // Update the document header to have no revisions
                    docHeader.setAttribute('data-rev-rows', '1');
                    // Reset merged cells
                    updateMergedCellsForDocument(docId);
                }
                else {
                    // Remove this revision
                    revisionRow.remove();
                    // Update the document header
                    const currentRevRows = parseInt(docHeader.getAttribute('data-rev-rows') || '1');
                    docHeader.setAttribute('data-rev-rows', (currentRevRows - 1).toString());
                    // Update revision numbers
                    updateRevisionNumbers(docId);
                    // Update merged cells
                    updateMergedCellsForDocument(docId);
                }
            };
            const updateRevisionNumbers = (docId) => {
                const docHeader = document.querySelector(`.dcs-document-header[data-doc-id="${docId}"]`);
                if (!docHeader)
                    return;
                const tableBody = docHeader.closest('tbody');
                if (!tableBody)
                    return;
                const revisionRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);
                revisionRows.forEach((row, index) => {
                    const revCell = row.querySelector('.dcs-rev-cell');
                    if (revCell) {
                        revCell.textContent = (index + 1).toString();
                    }
                });
            };
            const handleAddRevision = (e) => {
                e.preventDefault();
                e.stopPropagation();
                const button = e.target;
                const addRevBtn = button.closest('.dcs-add-revision-btn');
                if (addRevBtn) {
                    const docHeader = addRevBtn.closest('.dcs-document-header');
                    const docId = docHeader === null || docHeader === void 0 ? void 0 : docHeader.getAttribute('data-doc-id');
                    if (docId) {
                        addRevisionToDocument(docId);
                    }
                }
            };
            const attachDocumentEventListeners = () => {
                // Attach to Add Input Document button
                if (addInputDocBtn) {
                    addInputDocBtn.removeEventListener("click", handleAddInputDoc);
                    addInputDocBtn.addEventListener("click", handleAddInputDoc);
                }
                // Attach to Add Output Document button
                if (addOutputDocBtn) {
                    addOutputDocBtn.removeEventListener("click", handleAddOutputDoc);
                    addOutputDocBtn.addEventListener("click", handleAddOutputDoc);
                }
                // Attach listeners to ALL existing revision buttons
                document.querySelectorAll('.dcs-add-revision-btn').forEach(btn => {
                    btn.removeEventListener('click', handleAddRevision);
                    btn.addEventListener('click', handleAddRevision);
                });
                // Attach listeners to ALL existing remove document buttons
                document.querySelectorAll('.dcs-remove-doc').forEach(btn => {
                    btn.removeEventListener('click', handleRemoveDocument);
                    btn.addEventListener('click', handleRemoveDocument);
                });
                // Attach listeners to ALL existing remove revision buttons
                document.querySelectorAll('.dcs-remove-revision-btn').forEach(btn => {
                    btn.removeEventListener('click', function (e) {
                        e.preventDefault();
                        e.stopPropagation();
                        const docHeader = btn.closest('tr[data-doc-id]');
                        const docId = docHeader === null || docHeader === void 0 ? void 0 : docHeader.getAttribute('data-doc-id');
                        if (docId) {
                            removeRevisionFromDocument(docId, btn);
                        }
                    });
                    btn.addEventListener('click', function (e) {
                        e.preventDefault();
                        e.stopPropagation();
                        const docHeader = btn.closest('tr[data-doc-id]');
                        const docId = docHeader === null || docHeader === void 0 ? void 0 : docHeader.getAttribute('data-doc-id');
                        if (docId) {
                            removeRevisionFromDocument(docId, btn);
                        }
                    });
                });
            };
            const handleAddInputDoc = (e) => {
                if (e) {
                    e.preventDefault();
                    e.stopPropagation();
                }
                addNewDocumentToTable('input');
            };
            const handleAddOutputDoc = (e) => {
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
            const ensureDocumentButtonsVisible = () => {
                // Make sure all document action cells are visible
                const docActionCells = document.querySelectorAll('.dcs-doc-actions');
                docActionCells.forEach(cell => {
                    if (cell instanceof HTMLElement) {
                        cell.style.display = 'flex';
                        cell.style.visibility = 'visible';
                        cell.style.opacity = '1';
                    }
                });
                // Make sure all buttons within are visible
                const buttons = document.querySelectorAll('.dcs-doc-actions button');
                buttons.forEach(button => {
                    if (button instanceof HTMLElement) {
                        button.style.display = 'inline-flex';
                        button.style.visibility = 'visible';
                        button.style.opacity = '1';
                    }
                });
            };
            const handleRemoveDocument = (e) => {
                e.preventDefault();
                e.stopPropagation();
                const button = e.target;
                const removeBtn = button.closest('.dcs-remove-doc');
                if (removeBtn) {
                    const docHeader = removeBtn.closest('.dcs-document-header');
                    if (docHeader) {
                        const docId = docHeader.getAttribute('data-doc-id');
                        const isDefault = docHeader.getAttribute('data-is-default') === 'true';
                        if (isDefault) {
                            toast('Default document cannot be removed. Please clear the fields instead.', 'warning');
                            return;
                        }
                        if (confirm('Remove this document and all its revisions?')) {
                            // Find the table body
                            const tableBody = docHeader.closest('tbody');
                            if (!tableBody)
                                return;
                            // Get all rows with this docId (document header + all revision rows)
                            const allRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]`);
                            // Remove all rows
                            allRows.forEach(row => row.remove());
                            // Update serial numbers for the entire table
                            updateAllSerialNumbers();
                            // Re-attach event listeners
                            attachDocumentEventListeners();
                        }
                    }
                }
            };
            const addNewDocumentToTable = (tableType) => {
                const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
                if (!tableBody) {
                    console.error(`Table body not found for ${tableType}`);
                    return;
                }
                const existingDocs = tableBody.querySelectorAll('.dcs-document-header');
                const nextSerialNo = existingDocs.length + 1;
                // Generate unique docId
                const docId = `doc_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
                // Create new document row HTML
                const newDocumentHTML = createDocumentRowHTML(tableType, nextSerialNo, docId);
                // Append to table
                tableBody.insertAdjacentHTML('beforeend', newDocumentHTML);
                // Get the newly added row
                const newRow = tableBody.querySelector(`tr[data-doc-id="${docId}"]`);
                if (newRow) {
                    // Update serial numbers
                    updateAllSerialNumbers();
                    // Attach event listeners to the new row specifically
                    attachEventListenersToRow(newRow);
                }
            };
            const createDocumentRowHTML = (tableType, srNo, docId) => {
                const prefix = tableType === 'input' ? 'DOC' : 'DOC-OUT';
                const docNo = `${prefix}-${String(srNo).padStart(3, '0')}`;
                // Set correct transmittal placeholders based on table type
                const submissionPlaceholder = tableType === 'input' ? 'Transmittal No' : 'Transmittal No';
                const approvalPlaceholder = tableType === 'input' ? 'Transmittal No' : 'Transmittal No';
                return `
    <tr class="dcs-document-group dcs-document-header" data-doc-id="${docId}" data-rev-rows="1">
      <!-- Column 1: Sr No -->
      <td class="dcs-sr-no">${srNo}</td>
      <!-- Column 2: Document No -->
      <td>
        <textarea class="dcs-form-control dcs-doc-no" rows="2" placeholder="Document No">${docNo}</textarea>
      </td>
      <!-- Column 3: Description -->
      <td>
        <textarea class="dcs-form-control dcs-description-input" rows="3" placeholder="Enter description"></textarea>
      </td>
      <!-- Column 4: Category -->
      <td>
        <select class="dcs-form-control dcs-category-select">
          <option value="A" selected>A - Approval</option>
          <option value="I">I - Information</option>
        </select>
      </td>
      <!-- Column 5: Rev No -->
      <td class="dcs-rev-cell">0</td>
      <!-- Column 6: Original Plan -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Original Plan"></td>
      <!-- Column 7: Revised Plan -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Revised Plan"></td>
      <!-- Column 8: Actual Submission -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Actual Submission"></td>
      <!-- Column 9: Submission Transmittal -->
      <td>
        <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${submissionPlaceholder}"></textarea>
      </td>
      <!-- Column 10: Approval Date -->
      <td><input type="date" class="dcs-form-control dcs-date-input" placeholder="Approval Date"></td>
      <!-- Column 11: Approval Transmittal -->
      <td>
        <textarea class="dcs-form-control dcs-transmittal-input" rows="2" placeholder="${approvalPlaceholder}"></textarea>
      </td>
      <!-- Column 12: Status -->
      <td>
        <select class="dcs-form-control dcs-status-select">
          <option value="Pending" selected>Pending</option>
          <option value="Submitted">Submitted</option>
          <option value="Under Review">Under Review</option>
          <option value="Approved">Approved</option>
          <option value="Rejected">Rejected</option>
          <option value="Revised">Revised</option>
        </select>
      </td>
      <!-- Column 13: Remarks -->
      <td>
        <textarea class="dcs-form-control dcs-remarks-input" rows="3" placeholder="Enter remarks"></textarea>
      </td>
      <!-- Column 14: Actions -->
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
            const attachEventListenersToRow = (row) => {
                // Add revision button
                const addRevisionBtn = row.querySelector('.dcs-add-revision-btn');
                if (addRevisionBtn) {
                    addRevisionBtn.addEventListener('click', handleAddRevision);
                }
                // Remove document button
                const removeDocBtn = row.querySelector('.dcs-remove-doc');
                if (removeDocBtn) {
                    removeDocBtn.addEventListener('click', handleRemoveDocument);
                }
            };
            const checkEditPermission = (item) => {
                var _a, _b, _c;
                if (currentUser.isSuperAdmin) {
                    return true;
                }
                const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                if (!projectTitle)
                    return false;
                const project = projects.find(p => p.Title === projectTitle);
                if (!project)
                    return false;
                // Group Head can edit projects in their groups
                if (currentUser.isGroupHead) {
                    return ((_b = currentUser.managedGroups) === null || _b === void 0 ? void 0 : _b.some(group => project.GroupHead === group)) || false;
                }
                // Project Manager can edit only their assigned projects
                if (currentUser.isProjectManager) {
                    const projectManager = typeof project.ProjectManager === "object"
                        ? (_c = project.ProjectManager) === null || _c === void 0 ? void 0 : _c.Title
                        : project.ProjectManager;
                    return projectManager === currentUser.displayName;
                }
                return false;
            };
            const exportToExcel = async () => {
                // const progressEl = showExportProgress("Preparing data for export...");
                try {
                    // Get filtered data based on current filters
                    const project = filterProject.value;
                    const customer = filterCustomer.value;
                    const projectManager = filterProjectManager.value;
                    const projectReference = filterProjectReference.value;
                    const filteredData = dcsData.filter((item) => {
                        var _a, _b;
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
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
                        var _a;
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || "";
                        const projectData = projects.find((p) => p.Title === projectTitle);
                        const parsedData = parseDCSData(item[fieldMappings.DCSData]);
                        if (parsedData) {
                            const inputs = getInputsFromProjectData(parsedData);
                            const outputs = getOutputsFromProjectData(parsedData);
                            // Input documents
                            inputs.forEach((doc) => {
                                var _a;
                                const row = [
                                    (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "",
                                    projectTitle,
                                    (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "",
                                    typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                                        ? ((_a = projectData.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || ""
                                        : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "",
                                    doc.docNo,
                                    doc.description,
                                    doc.category === "A" ? "A - Approval" : "I - Information",
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
                                // Add revision rows
                                if (doc.revisions) {
                                    doc.revisions.forEach((revision) => {
                                        var _a;
                                        const revRow = [
                                            (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "",
                                            projectTitle,
                                            (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "",
                                            typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                                                ? ((_a = projectData.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || ""
                                                : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "",
                                            doc.docNo,
                                            doc.description,
                                            doc.category === "A" ? "A - Approval" : "I - Information",
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
                            outputs.forEach((doc) => {
                                var _a;
                                const row = [
                                    (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "",
                                    projectTitle,
                                    (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "",
                                    typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                                        ? ((_a = projectData.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || ""
                                        : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "",
                                    doc.docNo,
                                    doc.description,
                                    doc.category === "A" ? "A - Approval" : "I - Information",
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
                                // Add revision rows
                                if (doc.revisions) {
                                    doc.revisions.forEach((revision) => {
                                        var _a;
                                        const revRow = [
                                            (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "",
                                            projectTitle,
                                            (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "",
                                            typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                                                ? ((_a = projectData.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || ""
                                                : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "",
                                            doc.docNo,
                                            doc.description,
                                            doc.category === "A" ? "A - Approval" : "I - Information",
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
                    const csvContent = wsData.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');
                    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                    const link = document.createElement("a");
                    const url = URL.createObjectURL(blob);
                    link.setAttribute("href", url);
                    // Generate filename
                    const fileName = `DCS_Report_${new Date().toISOString().split('T')[0]}.csv`;
                    link.setAttribute("download", fileName);
                    link.style.visibility = 'hidden';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    hideExportProgress();
                    toast(`Exported: ${fileName}`, "success");
                }
                catch (error) {
                    hideExportProgress();
                    console.error("Export error:", error);
                    toast("Failed to export Excel report", "error");
                }
            };
            // Form Submission
            dcsForm.addEventListener("submit", async (e) => {
                e.preventDefault();
                await saveAllDataToSingleJSON();
            });
            const loadDocumentsFromJSON = (projectData) => {
                // Clear existing documents
                inputTableBody.innerHTML = '';
                outputTableBody.innerHTML = '';
                // Convert to new format if needed
                const dcsProjectData = convertToNewFormat(projectData);
                // Helper to create document row with revisions
                const createDocumentWithRevisions = (doc, tableType, index) => {
                    // Create main document row
                    const docRow = createEditableDocumentRow(doc, index + 1, tableType === 'input');
                    const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
                    tableBody.appendChild(docRow);
                    // Add revision rows if they exist
                    if (doc.revisions && doc.revisions.length > 0) {
                        doc.revisions.forEach((revision, revIndex) => {
                            const revisionHTML = `
          <tr class="dcs-document-group" data-doc-id="${doc.docId}">
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
            <td colspan="7"></td>
            <td class="dcs-doc-actions">
              <button type="button" class="dcs-btn-danger dcs-remove-revision-btn" 
                      title="Remove Revision">🗑️</button>
            </td>
          </tr>
        `;
                            tableBody.insertAdjacentHTML('beforeend', revisionHTML);
                        });
                        // Update merged cells
                        updateMergedCells(doc.docId);
                    }
                };
                // Load input documents
                if (dcsProjectData.documentStructure.inputs && dcsProjectData.documentStructure.inputs.length > 0) {
                    dcsProjectData.documentStructure.inputs.forEach((doc, index) => {
                        createDocumentWithRevisions(doc, 'input', index);
                    });
                }
                else {
                    // Add default input document
                    const defaultInputDoc = {
                        docId: `default_input_1`,
                        docNo: "DOC-001",
                        description: "",
                        category: "A",
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
                // Load output documents
                if (dcsProjectData.documentStructure.outputs && dcsProjectData.documentStructure.outputs.length > 0) {
                    dcsProjectData.documentStructure.outputs.forEach((doc, index) => {
                        createDocumentWithRevisions(doc, 'output', index);
                    });
                }
                else {
                    // Add default output document
                    const defaultOutputDoc = {
                        docId: `default_output_1`,
                        docNo: "DOC-OUT-001",
                        description: "",
                        category: "A",
                        revNo: 0,
                        originalPlan: "",
                        revisedPlan: "",
                        actualSubmission: "",
                        submissionTransmittal: "",
                        approvalDate: "",
                        approvalTransmittal: "",
                        status: "Pending",
                        remarks: "",
                        isInput: false,
                        revisions: [],
                        docType: 'output'
                    };
                    createDocumentWithRevisions(defaultOutputDoc, 'output', 0);
                }
                // Attach event listeners
                setTimeout(() => {
                    attachDocumentEventListeners();
                    updateAllSerialNumbers();
                }, 100);
            };
            const resetForm = () => {
                // Reset everything
                dcsForm.reset();
                // Remove view mode
                formSection.classList.remove("dcs-view-mode");
                // Hide form, show table
                formSection.style.display = "none";
                tableSection.style.display = "block";
                // Clear content
                inputTableBody.innerHTML = "";
                outputTableBody.innerHTML = "";
                containerTable.style.display = "none";
                projectDetails.innerHTML = "";
                existingDataAlert.style.display = "none";
                importButtonContainer.style.display = "none";
                // Reset form title
                formTitle.textContent = "Add Project Documents";
                // Enable project select
                projectSelect.disabled = false;
                // Reset editing state
                editingId = null;
                // Show all form buttons
                if (addInputDocBtn)
                    addInputDocBtn.style.display = "inline-block";
                if (addOutputDocBtn)
                    addOutputDocBtn.style.display = "inline-block";
                const submitBtn = dcsForm.querySelector('button[type="submit"]');
                if (submitBtn) {
                    submitBtn.style.display = "inline-flex";
                    submitBtn.innerHTML = '<span style="font-size: 16px; margin-right: 5px;">💾</span> Save Document Data';
                }
                // Reset to default documents
                resetDocumentTables();
                // Reset form actions to default
                const formActions = document.querySelector('.dcs-form-actions');
                if (formActions) {
                    formActions.innerHTML = `
      <button type="submit" class="dcs-btn-primary">
        <span style="font-size: 16px; margin-right: 5px;">💾</span> Save Document Data
      </button>
      <button type="button" class="dcs-btn-secondary dcs-back-to-table">
        <span style="font-size: 16px; margin-right: 5px;">⬅</span> Back to List
      </button>
    `;
                    // Reattach back button listener
                    const newBackBtn = formActions.querySelector('.dcs-back-to-table');
                    if (newBackBtn) {
                        newBackBtn.addEventListener('click', handleBackToTable);
                    }
                }
            };
            backBtn.addEventListener("click", resetForm);
            // Function to collect documents from a table - UPDATED for textarea fields
            // Add this function to update project details
            const updateProjectDetails = (projectId) => {
                var _a;
                const project = projects.find((p) => p.Id === projectId);
                if (project) {
                    const projectDetailsDiv = document.getElementById('dcsProjectDetails');
                    // Update the simple text displays
                    projectPC.textContent = project.PC || "-";
                    projectCustomer.textContent = project.CustomerName || "-";
                    projectReference.textContent = project.ProjectReference || "-";
                    projectGroupHead.textContent = project.GroupHead || "-";
                    const managerName = typeof project.ProjectManager === "object"
                        ? ((_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title) || "-"
                        : project.ProjectManager || "-";
                    projectManager.textContent = managerName;
                    // Also update the project details box with better styling
                    projectDetailsDiv.innerHTML = `
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px;">
        <div style="background: #e8f4fd; padding: 15px; border-radius: 8px; border-left: 4px solid #0d6efd;">
          <div style="font-size: 12px; color: #0d6efd; margin-bottom: 5px; font-weight: 600;">PROFIT CENTER</div>
          <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.PC || "-"}</div>
        </div>
        
        <div style="background: #d4edda; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745;">
          <div style="font-size: 12px; color: #28a745; margin-bottom: 5px; font-weight: 600;">CUSTOMER</div>
          <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.CustomerName || "-"}</div>
        </div>
        
        <div style="background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107;">
          <div style="font-size: 12px; color: #ffc107; margin-bottom: 5px; font-weight: 600;">PROJECT REFERENCE</div>
          <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.ProjectReference || "-"}</div>
        </div>
        
        <div style="background: #f8d7da; padding: 15px; border-radius: 8px; border-left: 4px solid #dc3545;">
          <div style="font-size: 12px; color: #dc3545; margin-bottom: 5px; font-weight: 600;">GROUP HEAD</div>
          <div style="font-size: 16px; font-weight: 600; color: #212529;">${project.GroupHead || "-"}</div>
        </div>
        
        <div style="background: #d1ecf1; padding: 15px; border-radius: 8px; border-left: 4px solid #17a2b8;">
          <div style="font-size: 12px; color: #17a2b8; margin-bottom: 5px; font-weight: 600;">PROJECT MANAGER</div>
          <div style="font-size: 16px; font-weight: 600; color: #212529;">${managerName}</div>
        </div>
      </div>
      
      <div style="margin-top: 15px; padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px dashed #dee2e6;">
        <div style="display: flex; align-items: center; gap: 10px; color: #6c757d;">
          <span style="font-size: 16px;">ℹ️</span>
          <div style="font-size: 13px;">
            <strong>View Mode:</strong> This is a read-only view of all document control data for this project.
            Click "Edit" on the main page to modify documents.
          </div>
        </div>
      </div>
    `;
                }
            };
            const collectDocumentsFromTable = (tableType) => {
                const docs = [];
                const tableId = tableType === 'input' ? 'dcsInputTableBody' : 'dcsOutputTableBody';
                const tableBody = document.getElementById(tableId);
                if (!tableBody)
                    return docs;
                // Get all document headers
                const docHeaders = tableBody.querySelectorAll('.dcs-document-header');
                docHeaders.forEach((docHeader, docIndex) => {
                    const docId = docHeader.getAttribute('data-doc-id') || `doc_${Date.now()}_${docIndex}`;
                    // Get all form elements from main document row
                    const getValue = (selector, defaultValue = '') => {
                        const element = docHeader.querySelector(selector);
                        if (element instanceof HTMLTextAreaElement || element instanceof HTMLInputElement) {
                            return element.value.trim();
                        }
                        else if (element instanceof HTMLSelectElement) {
                            return element.value;
                        }
                        return defaultValue;
                    };
                    // Get main document data
                    const doc = {
                        docId: docId,
                        docNo: getValue('.dcs-doc-no', `${tableType === 'input' ? 'DOC' : 'DOC-OUT'}-${String(docIndex + 1).padStart(3, '0')}`),
                        description: getValue('.dcs-description-input'),
                        category: getValue('.dcs-category-select', 'A'),
                        revNo: 0, // Main document always has revNo 0
                        originalPlan: getValue('td:nth-child(6) input.dcs-date-input'),
                        revisedPlan: getValue('td:nth-child(7) input.dcs-date-input'),
                        actualSubmission: getValue('td:nth-child(8) input.dcs-date-input'),
                        submissionTransmittal: getValue('td:nth-child(9) textarea.dcs-transmittal-input'),
                        approvalDate: getValue('td:nth-child(10) input.dcs-date-input'),
                        approvalTransmittal: getValue('td:nth-child(11) textarea.dcs-transmittal-input'),
                        status: getValue('td:nth-child(12) select.dcs-status-select', 'Pending'),
                        remarks: getValue('td:nth-child(13) textarea.dcs-remarks-input'),
                        isInput: tableType === 'input',
                        docType: tableType,
                        revisions: []
                    };
                    // Find all revision rows for this document (both with and without data-is-revision attribute)
                    const allRevisionRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);
                    allRevisionRows.forEach((row, revIndex) => {
                        // Helper to get revision field values
                        const getRevValue = (cellIndex, defaultValue = '') => {
                            const cell = row.querySelector(`td:nth-child(${cellIndex})`);
                            if (cell) {
                                const input = cell.querySelector('input');
                                const textarea = cell.querySelector('textarea');
                                if (input)
                                    return input.value.trim();
                                if (textarea)
                                    return textarea.value.trim();
                            }
                            return defaultValue;
                        };
                        // Get revision number from first cell
                        const revNumberCell = row.querySelector('.dcs-rev-cell');
                        const revNumber = revNumberCell ? parseInt(revNumberCell.textContent || '1') : revIndex + 1;
                        // Get revision status and remarks from main document (since they're merged)
                        const revisionStatus = doc.status;
                        const revisionRemarks = doc.remarks;
                        const revision = {
                            revNumber: revNumber,
                            originalPlan: getRevValue(2), // Column 2 in revision row = Original Plan
                            revisedPlan: getRevValue(3), // Column 3 in revision row = Revised Plan
                            actualSubmission: getRevValue(4), // Column 4 in revision row = Actual Submission
                            submissionTransmittal: getRevValue(5), // Column 5 in revision row = Submission Transmittal
                            approvalDate: getRevValue(6), // Column 6 in revision row = Approval Date
                            approvalTransmittal: getRevValue(7), // Column 7 in revision row = Approval Transmittal
                            status: revisionStatus,
                            remarks: revisionRemarks
                        };
                        doc.revisions.push(revision);
                    });
                    // Sort revisions by revision number
                    doc.revisions.sort((a, b) => a.revNumber - b.revNumber);
                    docs.push(doc);
                });
                return docs;
            };
            // Updated function to collect ALL documents in standardized format
            const collectAllDocumentsInJSONFormat = () => {
                return {
                    inputs: collectDocumentsFromTable('input'),
                    outputs: collectDocumentsFromTable('output')
                };
            };
            const saveAllDataToSingleJSON = async () => {
                var _a;
                try {
                    if (!projectSelect.value) {
                        toast("Please select a project", "error");
                        return;
                    }
                    const projectId = parseInt(projectSelect.value);
                    const projectTitle = ((_a = projectSelect.options[projectSelect.selectedIndex]) === null || _a === void 0 ? void 0 : _a.text) || `Project ${projectId}`;
                    // Check if we're editing an existing record
                    const isEditing = editingId !== null;
                    // Collect data in standardized format
                    const { inputs, outputs } = collectAllDocumentsInJSONFormat();
                    if (inputs.length === 0 && outputs.length === 0) {
                        toast("Please add at least one document", "warning");
                        return;
                    }
                    // Create the complete data structure in NEW format
                    const dcsProjectData = {
                        projectId: projectId,
                        projectTitle: projectTitle,
                        lastUpdated: new Date().toISOString(),
                        documentStructure: {
                            inputs: inputs,
                            outputs: outputs
                        },
                        summary: {
                            totalDocuments: inputs.length + outputs.length,
                            totalRevisions: inputs.reduce((sum, doc) => { var _a; return sum + (((_a = doc.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0) +
                                outputs.reduce((sum, doc) => { var _a; return sum + (((_a = doc.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0),
                            inputDocuments: inputs.length,
                            outputDocuments: outputs.length
                        }
                    };
                    // Save to SharePoint with retry mechanism
                    await saveWithRetry(dcsProjectData, projectId, editingId);
                    const totalDocs = inputs.length + outputs.length;
                    const totalRevisions = inputs.reduce((sum, doc) => { var _a; return sum + (((_a = doc.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0) +
                        outputs.reduce((sum, doc) => { var _a; return sum + (((_a = doc.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0);
                    toast(`✅ Success! ${isEditing ? 'Updated' : 'Saved'} ${totalDocs} documents with ${totalRevisions} revisions`, "success");
                    // Reset editing state
                    editingId = null;
                    // Return to table view
                    handleBackToTable();
                    await fetchData();
                }
                catch (error) {
                    console.error("Save error:", error);
                    // toast(`❌ Save failed: ${error.message}`, "error");
                }
            };
            // Enhanced save function for SharePoint
            const saveDCSDataToSharePoint = async (dcsProjectData, projectId, existingItemId = null) => {
                var _a, _b, _c, _d, _e, _f, _g, _h;
                try {
                    const digest = await getRequestDigest();
                    // Find existing record
                    let existingRecord = null;
                    if (existingItemId) {
                        existingRecord = dcsData.find((item) => item.Id === existingItemId);
                    }
                    else {
                        // Try to find by project ID
                        for (const item of dcsData) {
                            let itemProjectId = null;
                            // Check multiple possible field names
                            if (item[`${fieldMappings.Project}Id`]) {
                                itemProjectId = Number(item[`${fieldMappings.Project}Id`]);
                            }
                            else if (item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]) {
                                itemProjectId = Number(item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]);
                            }
                            else if (item.ProjectId) {
                                itemProjectId = Number(item.ProjectId);
                            }
                            else if (item.Project_ID) {
                                itemProjectId = Number(item.Project_ID);
                            }
                            else if ((_a = item[fieldMappings.Project.replace(/Id$/, '')]) === null || _a === void 0 ? void 0 : _a.Id) {
                                itemProjectId = Number(item[fieldMappings.Project.replace(/Id$/, '')].Id);
                            }
                            if (itemProjectId === projectId) {
                                existingRecord = item;
                                break;
                            }
                        }
                    }
                    // Prepare the complete data structure
                    const combinedData = {
                        projectId: projectId,
                        projectTitle: ((_b = projectSelect.options[projectSelect.selectedIndex]) === null || _b === void 0 ? void 0 : _b.text) || `Project ${projectId}`,
                        lastUpdated: new Date().toISOString(),
                        documentStructure: {
                            inputs: dcsProjectData.documentStructure.inputs || [],
                            outputs: dcsProjectData.documentStructure.outputs || []
                        },
                        summary: {
                            totalDocuments: (((_c = dcsProjectData.documentStructure.inputs) === null || _c === void 0 ? void 0 : _c.length) || 0) +
                                (((_d = dcsProjectData.documentStructure.outputs) === null || _d === void 0 ? void 0 : _d.length) || 0),
                            totalRevisions: (((_e = dcsProjectData.documentStructure.inputs) === null || _e === void 0 ? void 0 : _e.reduce((sum, doc) => { var _a; return sum + (((_a = doc.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0)) || 0) +
                                (((_f = dcsProjectData.documentStructure.outputs) === null || _f === void 0 ? void 0 : _f.reduce((sum, doc) => { var _a; return sum + (((_a = doc.revisions) === null || _a === void 0 ? void 0 : _a.length) || 0); }, 0)) || 0),
                            inputDocuments: ((_g = dcsProjectData.documentStructure.inputs) === null || _g === void 0 ? void 0 : _g.length) || 0,
                            outputDocuments: ((_h = dcsProjectData.documentStructure.outputs) === null || _h === void 0 ? void 0 : _h.length) || 0
                        }
                    };
                    // Determine the correct lookup field name
                    let lookupFieldName = `${fieldMappings.Project}Id`;
                    const noteFieldName = fieldMappings.DCSData;
                    // Use Record type which is more flexible for dynamic properties
                    const itemData = {
                        __metadata: {
                            type: listEntityTypeName
                        },
                        Title: `DCS Data - Project ${projectId}`,
                    };
                    // Try different lookup field patterns
                    const possibleLookupFields = [
                        `${fieldMappings.Project}Id`,
                        `${fieldMappings.Project.replace(/Id$/, '')}Id`,
                        "ProjectId",
                        "Project_ID",
                        "ProjectLookupId",
                        "ProjectReferenceId"
                    ];
                    // Use the project field from existing record if available
                    if (existingRecord) {
                        for (const field of possibleLookupFields) {
                            if (existingRecord[field] !== undefined) {
                                lookupFieldName = field;
                                break;
                            }
                        }
                    }
                    // Set the lookup field
                    itemData[lookupFieldName] = projectId;
                    // Set the DCS data
                    itemData[noteFieldName] = JSON.stringify(combinedData);
                    console.log("Saving data:", {
                        noteFieldName: noteFieldName,
                        lookupFieldName: lookupFieldName,
                        metadataType: itemData.__metadata.type,
                        isUpdate: !!existingRecord,
                        itemId: existingRecord === null || existingRecord === void 0 ? void 0 : existingRecord.Id
                    });
                    let url = apiBase;
                    const headers = {
                        Accept: "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": digest,
                    };
                    if (existingRecord) {
                        // UPDATE existing item - GET CURRENT ETAG FIRST
                        console.log('Getting current ETag for item', existingRecord.Id);
                        // First, fetch the current item to get its ETag
                        const getItemResponse = await fetch(`${apiBase}(${existingRecord.Id})`, {
                            headers: {
                                'Accept': 'application/json;odata=verbose',
                            },
                            credentials: 'same-origin'
                        });
                        if (!getItemResponse.ok) {
                            throw new Error(`Failed to get current item for ETag: ${getItemResponse.status}`);
                        }
                        // We only need the ETag from headers, not the response body
                        const currentETag = getItemResponse.headers.get('ETag');
                        console.log('Current ETag:', currentETag);
                        url = `${apiBase}(${existingRecord.Id})`;
                        // Use the correct ETag for the IF-MATCH header
                        if (currentETag) {
                            headers["IF-MATCH"] = currentETag;
                        }
                        else {
                            headers["IF-MATCH"] = "*"; // Fallback
                        }
                        headers["X-HTTP-Method"] = "MERGE";
                    }
                    console.log(`Making ${existingRecord ? 'UPDATE' : 'CREATE'} request to:`, url);
                    const response = await fetch(url, {
                        method: "POST",
                        headers: headers,
                        body: JSON.stringify(itemData),
                        credentials: 'same-origin'
                    });
                    if (!response.ok) {
                        const errorText = await response.text();
                        console.error('❌ Save error details:', {
                            status: response.status,
                            statusText: response.statusText,
                            errorText: errorText,
                            url: url,
                            metadataType: itemData.__metadata.type,
                            lookupFieldName: lookupFieldName,
                            isUpdate: !!existingRecord
                        });
                        // Handle 409 Conflict specifically
                        if (response.status === 409) {
                            console.log('⚠️ 409 Conflict - Item was modified by someone else. Retrying...');
                            // Try to get fresh data and retry
                            try {
                                await fetchData(); // Refresh data
                                // Find the item again after refresh
                                const refreshedData = dcsData.find((item) => item.Id === (existingRecord === null || existingRecord === void 0 ? void 0 : existingRecord.Id));
                                if (refreshedData) {
                                    console.log('Retrying with refreshed data...');
                                    // Retry the save
                                    return await saveDCSDataToSharePoint(dcsProjectData, projectId, existingItemId);
                                }
                            }
                            catch (retryError) {
                                console.error('Retry failed:', retryError);
                            }
                        }
                        // Try alternative metadata type if the current one fails
                        if (errorText.includes("valid type") || errorText.includes("EntityTypeName") || errorText.includes("could not resolve")) {
                            // Try alternative metadata types
                            const alternativeTypes = [
                                "SP.Data.DCSTracker_x0020_ListListItem", // Current
                                "SP.Data.DcsTrackerItem",
                                "SP.Data.DCSTrackerListItem",
                                "SP.Data.DCSTrackerItem",
                                "SP.Data.DcsTrackerListItem"
                            ];
                            for (const altType of alternativeTypes) {
                                if (altType !== listEntityTypeName) {
                                    try {
                                        console.log(`Trying alternative type: ${altType}`);
                                        const altItemData = Object.assign({}, itemData);
                                        altItemData.__metadata.type = altType;
                                        const retryResponse = await fetch(url, {
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
                                    }
                                    catch (retryError) {
                                        console.log(`Type ${altType} failed:`, retryError);
                                    }
                                }
                            }
                        }
                        throw new Error(`Save failed: ${response.status} - ${response.statusText}`);
                    }
                    console.log('✅ Save successful!');
                    await fetchData();
                }
                catch (error) {
                    console.error('❌ Save error:', error);
                    // if (error.message.includes("409") || error.message.includes("Conflict")) {
                    //   toast("Update conflict. Please refresh the page and try again.", "error");
                    // } else if (error.message.includes("412")) {
                    //   toast("Version conflict. Please refresh and try again.", "error");
                    // } else if (error.message.includes("does not exist") || error.message.includes("Invalid field")) {
                    //   toast("Configuration error. Please contact administrator.", "error");
                    // } else {
                    //   toast(`Save failed: ${error.message}`, "error");
                    // }  
                    throw error;
                }
            };
            // CORRECTED: saveWithRetry function (no unused variables)
            const saveWithRetry = async (dcsProjectData, projectId, existingItemId = null, retries = 3) => {
                let lastError = null;
                for (let attempt = 1; attempt <= retries; attempt++) {
                    try {
                        console.log(`Save attempt ${attempt} of ${retries}`);
                        // For 409 conflicts on retry, refresh data first
                        if (attempt > 1) {
                            await fetchData(); // Refresh data before retry
                        }
                        await saveDCSDataToSharePoint(dcsProjectData, projectId, existingItemId);
                        return; // Success
                    }
                    catch (error) {
                        lastError = error;
                        // Check for specific errors that warrant retry
                        const shouldRetry = (error.message.includes("409") || // Conflict
                            error.message.includes("412") || // Precondition failed
                            error.message.includes("ETag") ||
                            error.message.includes("500") || // Server error
                            error.message.includes("Conflict")) && attempt < retries;
                        if (shouldRetry) {
                            // Exponential backoff delay
                            const delay = 1000 * Math.pow(2, attempt - 1);
                            console.log(`Retry ${attempt} after ${delay}ms delay`);
                            await new Promise(resolve => setTimeout(resolve, delay));
                            continue; // Try again
                        }
                        // Not a retryable error or out of retries
                        break;
                    }
                }
                // If we get here, all retries failed
                throw lastError || new Error('Save failed after all retry attempts');
            };
            const initializePagination = (data) => {
                console.log('📄 initializePagination called with:', data.length, 'items');
                // Store the filtered data
                filteredDCSData = data;
                console.log('📊 filteredDCSData set to:', filteredDCSData.length, 'items');
                // Calculate pages
                calculateTotalPages();
                // Reset to first page
                currentPage = 1;
                // Update UI controls
                updatePaginationControls();
                // Render the table
                renderCurrentPage();
                console.log('✅ Pagination initialized. Showing page', currentPage, 'of', totalPages);
            };
            const calculateTotalPages = () => {
                totalPages = Math.ceil(filteredDCSData.length / pageSize);
            };
            const updatePaginationControls = () => {
                const currentPageSpan = document.getElementById("currentPage");
                const totalPagesSpan = document.getElementById("totalPages");
                const startItemSpan = document.getElementById("startItem");
                const endItemSpan = document.getElementById("endItem");
                const totalItemsSpan = document.getElementById("totalItems");
                if (currentPageSpan)
                    currentPageSpan.textContent = currentPage.toString();
                if (totalPagesSpan)
                    totalPagesSpan.textContent = totalPages.toString();
                const startItem = (currentPage - 1) * pageSize + 1;
                const endItem = Math.min(currentPage * pageSize, filteredDCSData.length);
                if (startItemSpan)
                    startItemSpan.textContent = startItem.toString();
                if (endItemSpan)
                    endItemSpan.textContent = endItem.toString();
                if (totalItemsSpan)
                    totalItemsSpan.textContent = filteredDCSData.length.toString();
                // Update page size selector
                const pageSizeSelect = document.getElementById("pageSizeSelect");
                if (pageSizeSelect)
                    pageSizeSelect.value = pageSize.toString();
                // Enable/disable navigation buttons
                const firstPageBtn = document.getElementById("firstPage");
                const prevPageBtn = document.getElementById("prevPage");
                const nextPageBtn = document.getElementById("nextPage");
                const lastPageBtn = document.getElementById("lastPage");
                if (firstPageBtn)
                    firstPageBtn.disabled = currentPage === 1;
                if (prevPageBtn)
                    prevPageBtn.disabled = currentPage === 1;
                if (nextPageBtn)
                    nextPageBtn.disabled = currentPage === totalPages;
                if (lastPageBtn)
                    lastPageBtn.disabled = currentPage === totalPages;
                updatePageNumbers();
            };
            const updatePageNumbers = () => {
                const pageNumbersDiv = document.getElementById("pageNumbers");
                if (!pageNumbersDiv)
                    return;
                pageNumbersDiv.innerHTML = '';
                const maxVisiblePages = 5;
                let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
                let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
                if (endPage - startPage + 1 < maxVisiblePages) {
                    startPage = Math.max(1, endPage - maxVisiblePages + 1);
                }
                // Add "..." before first page if needed
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
                // Add "..." after last page if needed
                if (endPage < totalPages) {
                    const dotsBtn = document.createElement('button');
                    dotsBtn.className = 'page-number dots';
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
                // Scroll to top of table
                const tableContainer = document.querySelector('.table-section');
                if (tableContainer) {
                    tableContainer.scrollTop = 0;
                }
            };
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
            };
            const renderCurrentPage = () => {
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, filteredDCSData.length);
                const pageData = filteredDCSData.slice(startIndex, endIndex);
                console.log('📄 renderCurrentPage called');
                console.log('📊 Page data to render:', pageData.length, 'items');
                console.log('📈 Showing items', startIndex + 1, 'to', endIndex, 'of', filteredDCSData.length);
                pageData.forEach((item, index) => {
                    var _a;
                    const projectTitle = (_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title;
                    const project = projects.find(p => p.Title === projectTitle);
                    console.log(`${index + 1}. ${projectTitle} - Manager: ${project === null || project === void 0 ? void 0 : project.ProjectManager}`);
                });
                renderTable(pageData);
            };
            const getCorrectListItemType = async () => {
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
                    // You can immediately update your variable
                    // listEntityTypeName = correctTypeName;
                }
                catch (error) {
                    console.error('❌ Failed to get correct type name:', error);
                }
            };
            // Call this function when your page loads or from the browser console
            await getCorrectListItemType();
            // Add this function to debug SharePoint list fields
            const debugListFields = async () => {
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false`, {
                        headers: { Accept: "application/json;odata=verbose" },
                        credentials: 'same-origin'
                    });
                    if (res.ok) {
                        const data = await res.json();
                        const fields = data.d.results;
                        fields.forEach((field) => {
                        });
                        // Find lookup fields
                        const lookupFields = fields.filter((f) => f.TypeAsString.includes("Lookup"));
                        lookupFields.forEach((field) => {
                        });
                        // Find note fields
                        const noteFields = fields.filter((f) => f.TypeAsString === "Note");
                        noteFields.forEach((field) => {
                        });
                    }
                }
                catch (error) {
                    console.error("Failed to debug fields:", error);
                }
            };
            // Call this during initialization
            await debugListFields();
            // Add import functionality functions
            const setupImportExportFunctionality = () => {
                const importExcelBtn = document.querySelector('.dcs-import-excel-btn');
                const importExcelInput = document.getElementById('dcsImportExcelInput');
                const exportProjectExcelBtn = document.querySelector('.dcs-export-project-excel-btn');
                const downloadExistingBtn = document.querySelector('.dcs-download-existing-btn');
                // Import Excel button click
                if (importExcelBtn && importExcelInput) {
                    importExcelBtn.addEventListener('click', () => {
                        importExcelInput.click();
                    });
                    importExcelInput.addEventListener('change', handleExcelImport);
                }
                // Export Template button
                if (exportProjectExcelBtn) {
                    exportProjectExcelBtn.addEventListener('click', exportTemplateExcel);
                }
                // Download Existing Data button
                if (downloadExistingBtn) {
                    downloadExistingBtn.addEventListener('click', downloadExistingData);
                }
            };
            // Handle Excel Import
            const handleExcelImport = async (event) => {
                var _a;
                const input = event.target;
                const file = (_a = input.files) === null || _a === void 0 ? void 0 : _a[0];
                if (!file) {
                    toast('Please select a file', 'warning');
                    return;
                }
                // Validate file type
                if (!file.name.match(/\.(xlsx|xls)$/i)) {
                    toast('Please select an Excel file (.xlsx or .xls)', 'error');
                    input.value = '';
                    return;
                }
                try {
                    showImportProgress('Reading Excel file...');
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        var _a;
                        try {
                            const data = (_a = e.target) === null || _a === void 0 ? void 0 : _a.result;
                            if (data) {
                                const workbook = XLSX.read(data, {
                                    type: 'binary',
                                    cellDates: true // Parse dates automatically
                                });
                                // Look for DCS Documents sheet or use first sheet
                                const sheetName = workbook.SheetNames.find(name => name.includes('DCS') || name.includes('Documents') || name.includes('Sheet')) || workbook.SheetNames[0];
                                const worksheet = workbook.Sheets[sheetName];
                                // Convert to array of arrays with proper type casting
                                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                                    header: 1,
                                    defval: '', // default value for empty cells
                                    raw: false // Convert dates to strings
                                });
                                // Process the imported data
                                processImportedData(jsonData);
                            }
                        }
                        catch (error) {
                            console.error('Error processing Excel:', error);
                            toast('Failed to process Excel file. Please check the format.', 'error');
                        }
                        finally {
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
                }
                catch (error) {
                    hideImportProgress();
                    console.error('Import error:', error);
                    toast('Failed to import Excel file', 'error');
                    input.value = '';
                }
            };
            // Process imported data from Excel - Enhanced to handle new documents and revisions
            const processImportedData = (data) => {
                // DON'T clear existing tables - we're merging/adding to existing data
                // Keep existing documents that are not in the Excel file
                let currentSection = null;
                let currentDoc = null;
                let currentDocRevisions = [];
                // Get existing documents to preserve them
                const existingInputDocs = getDocumentsFromTable('input');
                const existingOutputDocs = getDocumentsFromTable('output');
                // Collections for new/modified documents
                const newInputDocs = [...existingInputDocs];
                const newOutputDocs = [...existingOutputDocs];
                // Find where data starts
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
                // Process rows starting from dataStartRow
                for (let i = dataStartRow; i < data.length; i++) {
                    const row = data[i];
                    if (!row || row.length === 0)
                        continue;
                    // Convert all cells to strings
                    const processedRow = row.map(cell => {
                        if (cell === null || cell === undefined)
                            return '';
                        return String(cell).trim();
                    });
                    const firstCell = processedRow[0] || '';
                    // Check for section headers
                    if (firstCell.includes('DESIGN OUTPUT')) {
                        if (currentDoc && currentSection) {
                            saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
                        }
                        currentSection = 'output';
                        currentDoc = null;
                        currentDocRevisions = [];
                        continue;
                    }
                    else if (firstCell.includes('DESIGN INPUT')) {
                        if (currentDoc && currentSection) {
                            saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
                        }
                        currentSection = 'input';
                        currentDoc = null;
                        currentDocRevisions = [];
                        continue;
                    }
                    else if (firstCell === '' || !firstCell) {
                        // Empty row, save current document
                        if (currentDoc && currentSection) {
                            saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
                            currentDoc = null;
                            currentDocRevisions = [];
                        }
                        continue;
                    }
                    // Process data row
                    const docNo = processedRow[1] || '';
                    const description = processedRow[2] || '';
                    const category = (processedRow[3] || 'A').charAt(0);
                    const revNo = parseInt(processedRow[4] || '0') || 0;
                    // Parse dates
                    const originalPlan = parseExcelDate(processedRow[5] || '');
                    const revisedPlan = parseExcelDate(processedRow[6] || '');
                    const actualSubmission = parseExcelDate(processedRow[7] || '');
                    const submissionTransmittal = processedRow[8] || '';
                    const approvalDate = parseExcelDate(processedRow[9] || '');
                    const approvalTransmittal = processedRow[10] || '';
                    const status = processedRow[11] || 'Pending';
                    const remarks = processedRow[12] || '';
                    // Skip if no document number
                    if (!docNo) {
                        continue;
                    }
                    if (revNo === 0) {
                        // This is a main document (not a revision)
                        if (currentDoc && currentSection) {
                            saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
                        }
                        // Check if this document already exists
                        const existingDocs = currentSection === 'input' ? newInputDocs : newOutputDocs;
                        let existingDocIndex = -1;
                        if (currentSection) {
                            existingDocIndex = existingDocs.findIndex(doc => doc.docNo === docNo && doc.isInput === (currentSection === 'input'));
                        }
                        if (existingDocIndex >= 0 && currentSection) {
                            // Update existing document
                            const existingDoc = existingDocs[existingDocIndex];
                            currentDoc = Object.assign(Object.assign({}, existingDoc), { docNo: docNo, description: description || existingDoc.description, category: (category === 'A' || category === 'I') ? category : existingDoc.category, revNo: 0, originalPlan: originalPlan || existingDoc.originalPlan, revisedPlan: revisedPlan || existingDoc.revisedPlan, actualSubmission: actualSubmission || existingDoc.actualSubmission, submissionTransmittal: submissionTransmittal || existingDoc.submissionTransmittal, approvalDate: approvalDate || existingDoc.approvalDate, approvalTransmittal: approvalTransmittal || existingDoc.approvalTransmittal, status: status || existingDoc.status, remarks: remarks || existingDoc.remarks, revisions: [...existingDoc.revisions] // Keep existing revisions
                             });
                            // Remove the old one
                            if (currentSection === 'input') {
                                newInputDocs.splice(existingDocIndex, 1);
                            }
                            else {
                                newOutputDocs.splice(existingDocIndex, 1);
                            }
                        }
                        else {
                            // Create new document
                            currentDoc = {
                                docId: `imported_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                                docNo: docNo,
                                description: description,
                                category: category === 'A' || category === 'I' ? category : 'A',
                                revNo: 0,
                                originalPlan: originalPlan,
                                revisedPlan: revisedPlan,
                                actualSubmission: actualSubmission,
                                submissionTransmittal: submissionTransmittal,
                                approvalDate: approvalDate,
                                approvalTransmittal: approvalTransmittal,
                                status: status || 'Pending',
                                remarks: remarks,
                                isInput: currentSection === 'input',
                                revisions: [],
                                docType: currentSection || undefined
                            };
                        }
                        currentDocRevisions = [];
                    }
                    else if (revNo > 0 && currentDoc) {
                        // This is a revision of the current document
                        const revision = {
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
                        // Check if this revision already exists
                        const existingRevisionIndex = currentDocRevisions.findIndex(r => r.revNumber === revNo);
                        if (existingRevisionIndex >= 0) {
                            // Update existing revision
                            currentDocRevisions[existingRevisionIndex] = revision;
                        }
                        else {
                            // Add new revision
                            currentDocRevisions.push(revision);
                        }
                    }
                }
                // Save the last document
                if (currentDoc && currentSection) {
                    saveCurrentDocument(currentDoc, currentDocRevisions, newInputDocs, newOutputDocs);
                }
                // Now update the tables with combined data
                updateTablesWithImportedData(newInputDocs, newOutputDocs);
                toast('Excel data imported successfully! New documents/revisions added.', 'success');
            };
            // Helper to save current document to appropriate collection
            const saveCurrentDocument = (doc, revisions, inputDocs, outputDocs) => {
                doc.revisions = [...revisions];
                if (doc.isInput) {
                    inputDocs.push(doc);
                }
                else {
                    outputDocs.push(doc);
                }
            };
            // Get existing documents from table
            const getDocumentsFromTable = (tableType) => {
                const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
                const docs = [];
                if (!tableBody)
                    return docs;
                const docHeaders = tableBody.querySelectorAll('.dcs-document-header');
                docHeaders.forEach((docHeader) => {
                    const docId = docHeader.getAttribute('data-doc-id') || '';
                    // Get values from form elements
                    const getValue = (selector, defaultValue = '') => {
                        const element = docHeader.querySelector(selector);
                        if (element instanceof HTMLTextAreaElement || element instanceof HTMLInputElement) {
                            return element.value.trim();
                        }
                        else if (element instanceof HTMLSelectElement) {
                            return element.value;
                        }
                        return defaultValue;
                    };
                    const doc = {
                        docId: docId,
                        docNo: getValue('.dcs-doc-no'),
                        description: getValue('.dcs-description-input'),
                        category: getValue('.dcs-category-select', 'A'),
                        revNo: 0,
                        originalPlan: getValue('td:nth-child(6) input.dcs-date-input'),
                        revisedPlan: getValue('td:nth-child(7) input.dcs-date-input'),
                        actualSubmission: getValue('td:nth-child(8) input.dcs-date-input'),
                        submissionTransmittal: getValue('td:nth-child(9) textarea.dcs-transmittal-input'),
                        approvalDate: getValue('td:nth-child(10) input.dcs-date-input'),
                        approvalTransmittal: getValue('td:nth-child(11) textarea.dcs-transmittal-input'),
                        status: getValue('td:nth-child(12) select.dcs-status-select', 'Pending'),
                        remarks: getValue('td:nth-child(13) textarea.dcs-remarks-input'),
                        isInput: tableType === 'input',
                        revisions: [],
                        docType: tableType
                    };
                    // Get revision rows for this document
                    const revisionRows = tableBody.querySelectorAll(`tr[data-doc-id="${docId}"]:not(.dcs-document-header)`);
                    revisionRows.forEach((row) => {
                        const revCell = row.querySelector('.dcs-rev-cell');
                        const revNumber = revCell ? parseInt(revCell.textContent || '1') : 1;
                        const getRevValue = (cellIndex) => {
                            const cell = row.querySelector(`td:nth-child(${cellIndex})`);
                            if (cell) {
                                const input = cell.querySelector('input');
                                const textarea = cell.querySelector('textarea');
                                if (input)
                                    return input.value.trim();
                                if (textarea)
                                    return textarea.value.trim();
                            }
                            return '';
                        };
                        const revision = {
                            revNumber: revNumber,
                            originalPlan: getRevValue(2),
                            revisedPlan: getRevValue(3),
                            actualSubmission: getRevValue(4),
                            submissionTransmittal: getRevValue(5),
                            approvalDate: getRevValue(6),
                            approvalTransmittal: getRevValue(7),
                            status: doc.status, // Revisions inherit status from main doc
                            remarks: doc.remarks // Revisions inherit remarks from main doc
                        };
                        doc.revisions.push(revision);
                    });
                    docs.push(doc);
                });
                return docs;
            };
            // Update tables with combined data
            const updateTablesWithImportedData = (inputDocs, outputDocs) => {
                // Clear and rebuild input table
                inputTableBody.innerHTML = '';
                inputDocs.forEach((doc, index) => {
                    addDocumentToTable(doc, 'input', index + 1);
                });
                // Clear and rebuild output table
                outputTableBody.innerHTML = '';
                outputDocs.forEach((doc, index) => {
                    addDocumentToTable(doc, 'output', index + 1);
                });
                // Update serial numbers
                updateAllSerialNumbers();
                // Attach event listeners
                setTimeout(() => {
                    attachDocumentEventListeners();
                    ensureDocumentButtonsVisible();
                }, 100);
            };
            // Updated addDocumentToTable with srNo parameter
            const addDocumentToTable = (doc, tableType, srNo) => {
                const tableBody = tableType === 'input' ? inputTableBody : outputTableBody;
                if (!tableBody)
                    return;
                const docRow = createEditableDocumentRow(doc, srNo, tableType === 'input');
                tableBody.appendChild(docRow);
                // Add revision rows
                if (doc.revisions && doc.revisions.length > 0) {
                    doc.revisions.forEach((revision) => {
                        const submissionPlaceholder = tableType === 'input' ? 'Transmittal No' : 'Transmittal No';
                        const approvalPlaceholder = tableType === 'input' ? 'Transmittal No' : 'Transmittal No';
                        const revisionHTML = `
        <tr class="dcs-document-group" data-doc-id="${doc.docId}">
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
                  placeholder="${submissionPlaceholder}">${escapeHtml(revision.submissionTransmittal)}</textarea>
          </td>
          <td><input type="date" class="dcs-form-control dcs-date-input" 
                 value="${formatDateForInput(revision.approvalDate)}" 
                 placeholder="Approval Date"></td>       
          <td>
            <textarea class="dcs-form-control dcs-transmittal-input" rows="2" 
                  placeholder="${approvalPlaceholder}">${escapeHtml(revision.approvalTransmittal)}</textarea>
          </td>
          <td colspan="7"></td>
          <td class="dcs-doc-actions">
            <button type="button" class="dcs-btn-danger dcs-remove-revision-btn" 
                    title="Remove Revision">🗑️</button>
          </td>
        </tr>
      `;
                        tableBody.insertAdjacentHTML('beforeend', revisionHTML);
                    });
                    // Update merged cells
                    updateMergedCells(doc.docId);
                }
            };
            // Parse Excel date (handles both string and Excel serial numbers)
            const parseExcelDate = (dateStr) => {
                if (!dateStr || dateStr.trim() === '')
                    return '';
                // If it's already in YYYY-MM-DD format
                if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
                    return dateStr;
                }
                // If it's in DD/MM/YYYY or MM/DD/YYYY format
                if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
                    const parts = dateStr.split('/');
                    if (parts.length === 3) {
                        // Try both MM/DD/YYYY and DD/MM/YYYY
                        const date1 = new Date(`${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`);
                        const date2 = new Date(`${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`);
                        if (!isNaN(date1.getTime())) {
                            return date1.toISOString().split('T')[0];
                        }
                        else if (!isNaN(date2.getTime())) {
                            return date2.toISOString().split('T')[0];
                        }
                    }
                }
                // If it's an Excel serial number (numeric)
                if (!isNaN(Number(dateStr)) && dateStr !== '') {
                    const excelDate = Number(dateStr);
                    // Excel date system starts from 1899-12-30
                    const date = new Date((excelDate - 25569) * 86400 * 1000);
                    if (!isNaN(date.getTime())) {
                        return date.toISOString().split('T')[0];
                    }
                }
                // Try to parse as regular date
                const date = new Date(dateStr);
                if (!isNaN(date.getTime())) {
                    return date.toISOString().split('T')[0];
                }
                return '';
            };
            const exportTemplateExcel = () => {
                try {
                    const projectSelect = document.getElementById("dcsProjectSelect");
                    const projectId = parseInt(projectSelect.value);
                    const project = projects.find(p => p.Id === projectId);
                    if (!project) {
                        toast('Please select a project first', 'warning');
                        return;
                    }
                    // Create template data structure (SAME FORMAT as downloadExistingData)
                    const excelData = [];
                    // =================== HEADER ROWS ===================
                    excelData.push(['Sr No', 'Document No', 'Description', 'Category', 'Rev No',
                        'Original Plan', 'Revised Plan', 'Actual Submission', 'Transmittal No',
                        'Approval Date', 'Transmittal No', 'Status', 'Remarks']);
                    excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']); // Empty row
                    // =================== OUTPUT SECTION ===================
                    excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '']);
                    // Sample Output document
                    excelData.push([
                        '1', 'DOC-OUT-001', 'Sample Output Document Description', 'A - Approval', '0',
                        '', '', '', '', '', '', 'Pending', 'Add output document details here'
                    ]);
                    // Sample Output revision
                    excelData.push([
                        '1', 'DOC-OUT-001', 'Sample Output Document Description', 'A - Approval', '1',
                        '', '', '', '', '', '', 'Pending', 'Add revision details here'
                    ]);
                    // Add spacer row
                    excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
                    // =================== INPUT SECTION ===================
                    excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '']);
                    // Sample Input document
                    excelData.push([
                        '1', 'DOC-001', 'Sample Input Document Description', 'A - Approval', '0',
                        '', '', '', '', '', '', 'Pending', 'Add input document details here'
                    ]);
                    // Sample Input revision
                    excelData.push([
                        '1', 'DOC-001', 'Sample Input Document Description', 'A - Approval', '1',
                        '', '', '', '', '', '', 'Pending', 'Add revision details here'
                    ]);
                    const ws = XLSX.utils.aoa_to_sheet(excelData);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, 'DCS Template');
                    // Set column widths (SAME as downloadExistingData)
                    ws['!cols'] = [
                        { wch: 8 }, // Sr No
                        { wch: 20 }, // Document No
                        { wch: 35 }, // Description
                        { wch: 12 }, // Category
                        { wch: 8 }, // Rev No
                        { wch: 15 }, // Original Plan
                        { wch: 15 }, // Revised Plan
                        { wch: 18 }, // Actual Submission
                        { wch: 20 }, // Transmittal No
                        { wch: 15 }, // Approval Date
                        { wch: 20 }, // Transmittal No
                        { wch: 12 }, // Status
                        { wch: 25 } // Remarks
                    ];
                    // Generate filename
                    const timestamp = new Date().toISOString().split('T')[0];
                    const safeProjectName = project.Title.replace(/[^a-zA-Z0-9-_ ]/g, '_').substring(0, 30);
                    const fileName = `DCS_Template_${safeProjectName}_${timestamp}.xlsx`;
                    // Save file
                    XLSX.writeFile(wb, fileName);
                    toast('Template exported successfully! Edit and import when ready.', 'success');
                }
                catch (error) {
                    console.error('Export template error:', error);
                    toast('Failed to export template', 'error');
                }
            };
            const downloadExistingData = async () => {
                var _a;
                try {
                    const projectSelect = document.getElementById("dcsProjectSelect");
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
                    // Find existing data for this project
                    let existingRecord = null;
                    let foundData = false;
                    for (const item of dcsData) {
                        let itemProjectId = null;
                        // Try different field patterns
                        if (item[`${fieldMappings.Project}Id`]) {
                            itemProjectId = Number(item[`${fieldMappings.Project}Id`]);
                        }
                        else if (item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]) {
                            itemProjectId = Number(item[`${fieldMappings.Project.replace(/Id$/, '')}Id`]);
                        }
                        else if ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Id) {
                            itemProjectId = Number(item[fieldMappings.Project].Id);
                        }
                        else if (item.ProjectId) {
                            itemProjectId = Number(item.ProjectId);
                        }
                        else if (item.Project_ID) {
                            itemProjectId = Number(item.Project_ID);
                        }
                        if (itemProjectId === projectId) {
                            existingRecord = item;
                            foundData = true;
                            break;
                        }
                    }
                    // Prepare Excel data in TEMPLATE FORMAT (same as exportTemplateExcel)
                    const excelData = [];
                    // =================== HEADER ROWS (EXACT SAME AS TEMPLATE) ===================
                    excelData.push(['Sr No', 'Document No', 'Description', 'Category', 'Rev No',
                        'Original Plan', 'Revised Plan', 'Actual Submission', 'Transmittal No',
                        'Approval Date', 'Transmittal No', 'Status', 'Remarks']);
                    excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']); // Empty row
                    let hasAnyData = false;
                    if (foundData && existingRecord && existingRecord[fieldMappings.DCSData]) {
                        const parsedData = parseDCSData(existingRecord[fieldMappings.DCSData]);
                        if (parsedData) {
                            const inputs = getInputsFromProjectData(parsedData) || [];
                            const outputs = getOutputsFromProjectData(parsedData) || [];
                            hasAnyData = inputs.length > 0 || outputs.length > 0;
                            // =================== OUTPUT SECTION ===================
                            if (outputs.length > 0) {
                                excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '']);
                                outputs.forEach((doc, index) => {
                                    // Main document row (Rev No = 0)
                                    excelData.push([
                                        (index + 1).toString(),
                                        doc.docNo || `DOC-OUT-${String(index + 1).padStart(3, '0')}`,
                                        doc.description || '',
                                        doc.category === 'A' ? 'A - Approval' : 'I - Information',
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
                                    // Revision rows (Rev No > 0)
                                    if (doc.revisions && doc.revisions.length > 0) {
                                        doc.revisions.forEach((revision) => {
                                            excelData.push([
                                                (index + 1).toString(), // Same Sr No as parent
                                                doc.docNo || `DOC-OUT-${String(index + 1).padStart(3, '0')}`,
                                                doc.description || '',
                                                doc.category === 'A' ? 'A - Approval' : 'I - Information',
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
                            }
                            else {
                                // Add empty output section if no data exists
                                excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '']);
                                excelData.push([
                                    '1', 'DOC-OUT-001', '', 'A - Approval', '0',
                                    '', '', '', '', '', '', 'Pending', ''
                                ]);
                            }
                            // Add spacer row
                            excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
                            // =================== INPUT SECTION ===================
                            if (inputs.length > 0) {
                                excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '']);
                                inputs.forEach((doc, index) => {
                                    // Main document row (Rev No = 0)
                                    excelData.push([
                                        (index + 1).toString(),
                                        doc.docNo || `DOC-${String(index + 1).padStart(3, '0')}`,
                                        doc.description || '',
                                        doc.category === 'A' ? 'A - Approval' : 'I - Information',
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
                                    // Revision rows (Rev No > 0)
                                    if (doc.revisions && doc.revisions.length > 0) {
                                        doc.revisions.forEach((revision) => {
                                            excelData.push([
                                                (index + 1).toString(), // Same Sr No as parent
                                                doc.docNo || `DOC-${String(index + 1).padStart(3, '0')}`,
                                                doc.description || '',
                                                doc.category === 'A' ? 'A - Approval' : 'I - Information',
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
                            }
                            else {
                                // Add empty input section if no data exists
                                excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '']);
                                excelData.push([
                                    '1', 'DOC-001', '', 'A - Approval', '0',
                                    '', '', '', '', '', '', 'Pending', ''
                                ]);
                            }
                        }
                    }
                    else {
                        // No existing data - create empty template
                        hasAnyData = false;
                        // =================== OUTPUT SECTION ===================
                        excelData.push(['DESIGN OUTPUT - FROM ABB', '', '', '', '', '', '', '', '', '', '', '', '']);
                        excelData.push([
                            '1', 'DOC-OUT-001', '', 'A - Approval', '0',
                            '', '', '', '', '', '', 'Pending', ''
                        ]);
                        // Add spacer row
                        excelData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
                        // =================== INPUT SECTION ===================
                        excelData.push(['DESIGN INPUT FROM CONSULTANT / CUSTOMER', '', '', '', '', '', '', '', '', '', '', '', '']);
                        excelData.push([
                            '1', 'DOC-001', '', 'A - Approval', '0',
                            '', '', '', '', '', '', 'Pending', ''
                        ]);
                    }
                    // Create worksheet with EXACT same format as template
                    const ws = XLSX.utils.aoa_to_sheet(excelData);
                    // Set column widths (SAME AS TEMPLATE)
                    ws['!cols'] = [
                        { wch: 8 }, // Sr No
                        { wch: 20 }, // Document No
                        { wch: 35 }, // Description
                        { wch: 12 }, // Category
                        { wch: 8 }, // Rev No
                        { wch: 15 }, // Original Plan
                        { wch: 15 }, // Revised Plan
                        { wch: 18 }, // Actual Submission
                        { wch: 20 }, // Transmittal No
                        { wch: 15 }, // Approval Date
                        { wch: 20 }, // Transmittal No
                        { wch: 12 }, // Status
                        { wch: 25 } // Remarks
                    ];
                    // Create workbook
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, 'DCS Data');
                    // Generate filename
                    const timestamp = new Date().toISOString().split('T')[0];
                    const safeProjectName = project.Title.replace(/[^a-zA-Z0-9-_ ]/g, '_').substring(0, 30);
                    const fileName = hasAnyData
                        ? `DCS_Data_${safeProjectName}_${timestamp}.xlsx`
                        : `DCS_Template_${safeProjectName}_${timestamp}.xlsx`;
                    // Save file
                    XLSX.writeFile(wb, fileName);
                    const message = hasAnyData
                        ? 'Current data downloaded. Edit in Excel and re-upload.'
                        : 'Template downloaded. Fill data and upload.';
                    toast(message, 'success');
                }
                catch (error) {
                    console.error('Download error:', error);
                    toast(`Failed to download: ${error.message}`, 'error');
                }
            };
            // Show import progress
            const showImportProgress = (message) => {
                const existing = document.querySelector('.dcs-import-progress');
                if (existing)
                    existing.remove();
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
      <div style="font-size: 16px; font-weight: 500; margin-bottom: 10px;">${message}</div>
      <div style="font-size: 12px; color: #6c757d;">Please wait...</div>
    </div>
  `;
                document.body.appendChild(progressEl);
            };
            // Hide import progress
            const hideImportProgress = () => {
                const existing = document.querySelector('.dcs-import-progress');
                if (existing)
                    existing.remove();
            };
            const initializeSharePoint = async () => {
                try {
                    const listRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                        headers: { Accept: "application/json;odata=verbose" },
                        credentials: 'same-origin'
                    });
                    if (listRes.ok) {
                        const listInfo = await listRes.json();
                        listEntityTypeName = listInfo.d.ListItemEntityTypeFullName;
                        console.log('✅ SharePoint initialized. List item type:', listEntityTypeName);
                    }
                    else {
                        throw new Error(`List '${listName}' not found or inaccessible`);
                    }
                    // Get fields to determine correct names
                    const fieldsRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false`, {
                        headers: { Accept: "application/json;odata=verbose" },
                        credentials: 'same-origin'
                    });
                    if (fieldsRes.ok) {
                        const fieldsData = await fieldsRes.json();
                        const fields = fieldsData.d.results;
                        // Find DCSData field (Note type)
                        const dcsDataField = fields.find((f) => f.TypeAsString === "Note" &&
                            (f.InternalName.includes("DCS") || f.Title.includes("DCS") || f.InternalName === "DCSData")) || fields.find((f) => f.TypeAsString === "Note");
                        // Find Project field (Lookup type)
                        const projectField = fields.find((f) => f.TypeAsString.includes("Lookup") &&
                            (f.InternalName.includes("Project") || f.Title.includes("Project"))) || fields.find((f) => f.TypeAsString.includes("Lookup"));
                        fieldMappings.DCSData = (dcsDataField === null || dcsDataField === void 0 ? void 0 : dcsDataField.InternalName) || "DCSData";
                        fieldMappings.Project = (projectField === null || projectField === void 0 ? void 0 : projectField.InternalName) || "Project";
                        // If Project field ends with "Id", remove it for expansion
                        if (fieldMappings.Project.endsWith("Id")) {
                            fieldMappings.Project = fieldMappings.Project.replace(/Id$/, '');
                        }
                        console.log('✅ Field mappings:', fieldMappings);
                    }
                }
                catch (error) {
                    console.error("Error initializing SharePoint:", error);
                    // Use safe defaults
                    fieldMappings = {
                        DCSData: "DCSData",
                        Project: "Project",
                    };
                    // Use the type from your logs as fallback
                    listEntityTypeName = "SP.Data.DcsTrackerItem";
                    console.log('⚠️ Using fallback values:', { fieldMappings, listEntityTypeName });
                }
            };
            async function getListItemType() {
                try {
                    const siteUrl = _spPageContextInfo.webAbsoluteUrl;
                    const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('DCSTracker')?$select=ListItemEntityTypeFullName`, {
                        headers: { 'Accept': 'application/json;odata=verbose' },
                        credentials: 'same-origin'
                    });
                    const data = await response.json();
                    console.log('✅ Exact List Item Type:', data.d.ListItemEntityTypeFullName);
                    return data.d.ListItemEntityTypeFullName;
                }
                catch (error) {
                    console.error('Failed:', error);
                }
            }
            getListItemType();
            const resetDocumentTables = () => {
                // Reset input table
                const defaultInputDoc = {
                    docId: `default_input_1`,
                    docNo: "DOC-001",
                    description: "",
                    category: "A",
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
                const defaultOutputDoc = {
                    docId: `default_output_1`,
                    docNo: "DOC-OUT-001",
                    description: "",
                    category: "A",
                    revNo: 0,
                    originalPlan: "",
                    revisedPlan: "",
                    actualSubmission: "",
                    submissionTransmittal: "",
                    approvalDate: "",
                    approvalTransmittal: "",
                    status: "Pending",
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
            const updateAllSerialNumbers = () => {
                // Update input table
                const inputTable = document.getElementById('dcsInputTableBody');
                if (inputTable) {
                    const inputDocs = inputTable.querySelectorAll('.dcs-document-header');
                    inputDocs.forEach((doc, index) => {
                        const srNoCell = doc.querySelector('.dcs-sr-no');
                        if (srNoCell) {
                            srNoCell.textContent = (index + 1).toString();
                        }
                        // Update document number in the textarea
                        const docNoInput = doc.querySelector('.dcs-doc-no');
                        if (docNoInput) {
                            // Only update if it's not manually changed by user
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
                        // Update document number in the textarea
                        const docNoInput = doc.querySelector('.dcs-doc-no');
                        if (docNoInput) {
                            // Only update if it's not manually changed by user
                            const currentValue = docNoInput.value.trim();
                            const expectedPattern = new RegExp(`^DOC-OUT-\\d{3}$`);
                            if (!currentValue || expectedPattern.test(currentValue)) {
                                docNoInput.value = `DOC-OUT-${String(index + 1).padStart(3, '0')}`;
                            }
                        }
                    });
                }
            };
            if (resetFormBtn) {
                resetFormBtn.addEventListener("click", resetForm);
            }
            // Event Listeners
            showFormBtn.addEventListener("click", () => {
                resetForm();
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formSection.style.animation = "fadeIn 0.3s ease";
            });
            backBtn.addEventListener("click", () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                tableSection.style.animation = "fadeIn 0.3s ease";
            });
            projectSelect.addEventListener("change", (e) => {
                const selectedId = parseInt(e.target.value);
                if (selectedId) {
                    onProjectChange(selectedId);
                }
            });
            // Filter events
            filterProject.addEventListener("change", filterData);
            filterCustomer.addEventListener("change", filterData);
            filterProjectManager.addEventListener("change", filterData);
            filterProjectReference.addEventListener("change", filterData);
            clearFiltersBtn.addEventListener("click", clearFilters);
            // Export button
            if (exportExcelBtn) {
                exportExcelBtn.addEventListener("click", exportToExcel);
            }
            initializeEventListeners();
            // Initialize app
            await initApp();
        }, 100);
    }
};
// Export as default
export default DCSModule;
//# sourceMappingURL=dci.js.map
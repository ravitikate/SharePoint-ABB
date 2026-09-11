// project.ts - Complete file with Engineering Managers field
import Swal from 'sweetalert2';
import 'sweetalert2/dist/sweetalert2.min.css';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import "../styles/common.css";
import { authManager, canContribute, canRead, hasFullControl, canEditProject, getAssignedProjects, updateUIBasedOnPermissions } from './auth';
// ============================================
// PERMISSION HELPERS
// ============================================
const hasEditPermission = (project) => {
    return canEditProject(project);
};
const canCreate = () => {
    return canContribute();
};
// ============================================
// MANAGER LOOKUP COMPONENT
// ============================================
class ManagerLookup {
    constructor(container, siteUrl, options) {
        this.container = container;
        this.siteUrl = siteUrl;
        this.onSelect = options === null || options === void 0 ? void 0 : options.onSelect;
        this.onClear = options === null || options === void 0 ? void 0 : options.onClear;
        this.searchInput = container.querySelector('.manager-search-input');
        this.searchBtn = container.querySelector('.search-manager-btn');
        this.suggestionsDiv = container.querySelector('.manager-suggestions');
        this.selectedDisplay = container.querySelector('.selected-manager');
        this.selectedNameSpan = container.querySelector('.selected-manager-name');
        this.removeBtn = container.querySelector('.remove-manager');
        this.hiddenInput = container.querySelector('.manager-id-input');
        this.bindEvents();
    }
    bindEvents() {
        this.searchBtn.addEventListener('click', () => this.search());
        this.searchInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.search();
            }
        });
        this.searchInput.addEventListener('input', () => {
            const term = this.searchInput.value.trim();
            if (term.length >= 2) {
                this.search();
            }
            else {
                this.suggestionsDiv.style.display = 'none';
            }
        });
        this.searchInput.addEventListener('focus', () => {
            const term = this.searchInput.value.trim();
            if (term.length >= 2) {
                this.search();
            }
        });
        this.removeBtn.addEventListener('click', () => this.clear());
        document.addEventListener('click', (e) => {
            if (!this.container.contains(e.target)) {
                this.suggestionsDiv.style.display = 'none';
            }
        });
    }
    async search() {
        const searchTerm = this.searchInput.value.trim();
        if (searchTerm.length < 2) {
            Swal.fire({ title: 'Please enter at least 2 characters', icon: 'warning' });
            return;
        }
        try {
            const managers = await this.searchManagers(searchTerm);
            this.displaySuggestions(managers);
        }
        catch (error) {
            console.error('Error searching managers:', error);
            Swal.fire({ title: 'Search failed', icon: 'error' });
        }
    }
    async searchManagers(searchTerm) {
        const url = `${this.siteUrl}/_api/web/siteusers?$filter=PrincipalType eq 1 and (substringof('${searchTerm}', Title) eq true or substringof('${searchTerm}', Email) eq true)&$top=20&$orderby=Title`;
        const res = await fetch(url, {
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
            }
        });
        if (!res.ok) {
            throw new Error(`HTTP error! status: ${res.status}`);
        }
        const data = await res.json();
        const users = data.d.results || [];
        return users
            .filter((user) => {
            const title = user.Title || '';
            return !title.includes('System Account') &&
                !title.includes('App@sharepoint') &&
                !title.includes('SharingLinks') &&
                !title.includes('NT AUTHORITY') &&
                !title.includes('Membership');
        })
            .map((user) => ({
            id: user.Id,
            title: user.Title || '',
            email: user.Email || ''
        }))
            .sort((a, b) => a.title.localeCompare(b.title));
    }
    displaySuggestions(managers) {
        this.suggestionsDiv.innerHTML = '';
        if (managers.length === 0) {
            const noResults = document.createElement('div');
            noResults.className = 'manager-suggestion-item';
            noResults.textContent = 'No managers found';
            this.suggestionsDiv.appendChild(noResults);
        }
        else {
            managers.forEach(manager => {
                const suggestion = document.createElement('div');
                suggestion.className = 'manager-suggestion-item';
                suggestion.innerHTML = `
          <strong>${manager.title}</strong>
          ${manager.email ? `<br><small>${manager.email}</small>` : ''}
        `;
                suggestion.dataset.id = manager.id.toString();
                suggestion.addEventListener('click', () => this.select(manager));
                this.suggestionsDiv.appendChild(suggestion);
            });
        }
        this.suggestionsDiv.style.display = 'block';
    }
    select(manager) {
        this.searchInput.style.display = 'none';
        this.searchBtn.style.display = 'none';
        this.selectedNameSpan.textContent = `${manager.title}${manager.email ? ` (${manager.email})` : ''}`;
        this.selectedDisplay.style.display = 'flex';
        this.hiddenInput.value = manager.id.toString();
        this.suggestionsDiv.style.display = 'none';
        if (this.onSelect) {
            this.onSelect(manager);
        }
    }
    clear() {
        this.searchInput.style.display = 'block';
        this.searchBtn.style.display = 'block';
        this.selectedDisplay.style.display = 'none';
        this.hiddenInput.value = '';
        this.searchInput.value = '';
        this.suggestionsDiv.style.display = 'none';
        if (this.onClear) {
            this.onClear();
        }
    }
    getSelectedId() {
        return this.hiddenInput.value || null;
    }
    getSelectedName() {
        return this.selectedNameSpan.textContent || '';
    }
    setSelected(manager) {
        if (manager && manager.id) {
            this.select(manager);
        }
        else {
            this.clear();
        }
    }
}
// ============================================
// MAIN PROJECT MANAGEMENT COMPONENT
// ============================================
const projectContent = `
<div class="project-page fade-in">
  <section class="table-section card">
    <div class="page-header">
      <h1>Project Management</h1>
      <div>
        <button class="btn-primary show-form-btn">+ Add New Project</button>
        <button class="btn-primary export-btn">Export Excel</button>
        <button class="btn-primary import-btn">Import Excel</button>
        <input type="file" class="excel-input" accept=".xlsx,.xls" style="display:none;" />
      </div>
    </div>
    <div class="table-actions">
      <input type="text" class="search-input" placeholder="Search Projects..." />
      <button class="btn-secondary reset-search-btn">Reset Filter</button>
    </div>
    <div style="overflow-x:auto;">
      <table class="modern-table">
        <thead>
          <tr>
            <th>Profit Centre</th>
            <th>Project Reference</th>
            <th>Customer Name</th>
            <th>Project Name</th>
            <th>GH</th>
            <th>Project Manager</th>
            <th>Engg Lead</th>
            <th>Engg Sub Lead 1</th>
            <th>Engg Sub Lead 2</th>
            <th>Engineering Managers</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody class="project-table-body"></tbody>
      </table>
    </div>
    <div class="pagination-controls">
      <div class="pagination-info">
        <span class="record-count"></span>
      </div>
      <div class="pagination-size">
        <label>Rows per page:</label>
        <select class="page-size-select">
          <option value="10">10</option>
          <option value="25" selected>25</option>
          <option value="50">50</option>
          <option value="100">100</option>
        </select>
      </div>
      <div class="pagination"></div>
    </div>
  </section>

  <section class="form-section card" style="display:none;">
    <div class="form-header"><h2>➕ Add New Project</h2></div>
    <form id="projectForm">
      <div class="form-row">
        <div class="form-group"><label>Profit Centre</label><input type="text" id="pc" required /></div>
        <div class="form-group"><label>Project Reference</label><input type="text" id="reference" required /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Customer Name</label><input type="text" id="customer" required /></div>
        <div class="form-group"><label>Project Name</label><input type="text" id="name" required /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>GH</label><input type="text" id="head" required /></div>
        <div class="form-group">
          <label>Project Manager</label>
          <div class="manager-lookup-container" id="projectManagerLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search for manager..." />
              <button type="button" class="search-manager-btn">Search</button>
            </div>
            <div class="manager-suggestions" style="display:none;"></div>
            <div class="selected-manager" style="display:none;">
              <span class="selected-manager-name"></span>
              <button type="button" class="remove-manager">×</button>
            </div>
            <input type="hidden" class="manager-id-input" />
          </div>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Engg Lead</label>
          <div class="manager-lookup-container" id="enggLeadLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search for engg lead..." />
              <button type="button" class="search-manager-btn">Search</button>
            </div>
            <div class="manager-suggestions" style="display:none;"></div>
            <div class="selected-manager" style="display:none;">
              <span class="selected-manager-name"></span>
              <button type="button" class="remove-manager">×</button>
            </div>
            <input type="hidden" class="manager-id-input" />
          </div>
        </div>
        <div class="form-group">
          <label>Engg Sub Lead 1</label>
          <div class="manager-lookup-container" id="enggSubLead1Lookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search for engg sub lead..." />
              <button type="button" class="search-manager-btn">Search</button>
            </div>
            <div class="manager-suggestions" style="display:none;"></div>
            <div class="selected-manager" style="display:none;">
              <span class="selected-manager-name"></span>
              <button type="button" class="remove-manager">×</button>
            </div>
            <input type="hidden" class="manager-id-input" />
          </div>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Engg Sub Lead 2</label>
          <div class="manager-lookup-container" id="enggSubLead2Lookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search for engg sub lead..." />
              <button type="button" class="search-manager-btn">Search</button>
            </div>
            <div class="manager-suggestions" style="display:none;"></div>
            <div class="selected-manager" style="display:none;">
              <span class="selected-manager-name"></span>
              <button type="button" class="remove-manager">×</button>
            </div>
            <input type="hidden" class="manager-id-input" />
          </div>
        </div>
        <div class="form-group">
          <label>Engineering Managers</label>
          <div class="manager-lookup-container" id="enggManagersLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search for engineering manager..." />
              <button type="button" class="search-manager-btn">Search</button>
            </div>
            <div class="manager-suggestions" style="display:none;"></div>
            <div class="selected-manager" style="display:none;">
              <span class="selected-manager-name"></span>
              <button type="button" class="remove-manager">×</button>
            </div>
            <input type="hidden" class="manager-id-input" />
          </div>
        </div>
      </div>
      <div class="form-actions">
        <button type="submit" class="btn-primary">Save Project</button>
        <button type="button" class="btn-secondary back-to-table">Back to List</button>
      </div>
    </form>
  </section>
</div>
`;
export default {
    content: projectContent,
    style: `
.project-page { font-family:'Segoe UI', Arial,sans-serif; color:#333; animation:fadeIn 0.4s ease-in; }

.card { background:#fff; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,0.08); padding:20px; margin-bottom:20px; transition:all 0.3s ease; }

/* Table Styles with text wrapping */
.modern-table { width:100%; border-collapse:collapse; font-size:14px; }
.modern-table th { 
  background:#f3f2f1; 
  text-align:left; 
  padding:12px; 
  font-weight:600; 
  border-bottom:2px solid #e1e1e1; 
  white-space:nowrap; 
}
.modern-table td { 
  padding:10px; 
  border-bottom:1px solid #eee; 
  max-width:200px; 
  word-wrap:break-word;
  white-space:normal;
  vertical-align:middle;
}
.modern-table tr:hover { background:#fafafa; }

/* Style for manager name with email on new line */
.manager-cell {
  display:flex;
  flex-direction:column;
  gap:2px;
}
.manager-name {
  font-weight:500;
  color:#333;
}
.manager-email {
  font-size:11px;
  color:#666;
  word-break:break-all;
}

.form-row { display:flex; flex-wrap:wrap; gap:20px; margin-bottom:20px; }
.form-group { flex:1 1 calc(50% - 10px); display:flex; flex-direction:column; }
.form-group label { font-weight:500; color:#444; margin-bottom:6px; }
.form-group input, .form-group select { width:100%; padding:10px 12px; border-radius:6px; border:1px solid #ccc; font-size:14px; line-height:1.4; box-sizing:border-box; transition:all 0.2s ease; }
.form-group input:focus, .form-group select:focus { border-color:#0078d4; box-shadow:0 0 3px rgba(0,120,212,0.4); outline:none; }
.form-actions { text-align:center; margin-top:25px; }
.btn-primary { background:#0078d4; color:#fff; border:none; padding:10px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; transition:all 0.2s ease; }
.btn-primary:hover { background:#005a9e; }
.btn-secondary { background:#e1e1e1; color:#333; border:none; padding:10px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; transition:all 0.2s ease; }
.btn-secondary:hover { background:#ccc; }
.form-header h2 { color:#0078d4; margin-bottom:20px; font-size:18px; font-weight:600; }
.table-actions { display:flex; justify-content:space-between; margin-bottom:10px; flex-wrap:wrap; gap:10px; }
.search-input { padding:8px 12px; border:1px solid #ccc; border-radius:6px; flex:1; min-width:180px; }

/* Pagination Controls - Aligned in single row */
.pagination-controls {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-top: 15px;
  padding-top: 15px;
  border-top: 1px solid #eee;
  flex-wrap: wrap;
  gap: 10px;
}

.pagination-info {
  font-size: 14px;
  color: #666;
  white-space: nowrap;
}

.pagination-size {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 14px;
  color: #666;
  white-space: nowrap;
}

.pagination-size select {
  padding: 6px 10px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 14px;
  background: white;
  cursor: pointer;
}

.pagination-size select:focus {
  outline: none;
  border-color: #0078d4;
}

.pagination {
  display: flex;
  gap: 5px;
  flex-wrap: wrap;
  align-items: center;
}

.pagination button {
  padding: 6px 12px;
  border: 1px solid #ddd;
  background: #fff;
  border-radius: 4px;
  cursor: pointer;
  transition: all 0.2s ease;
  font-size: 14px;
  color: #333;
}

.pagination button:hover {
  background: #f0f0f0;
}

.pagination button.active {
  background: #0078d4;
  color: #fff;
  border-color: #0078d4;
}

.pagination button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.toast { position: fixed; bottom: 20px; right: 20px; padding: 12px 20px; color: #fff; border-radius: 6px; box-shadow: 0 2px 6px rgba(0,0,0,0.2); opacity: 0; transition: opacity 0.3s; z-index: 10000; }
.toast.success { background: #28a745; }
.toast.error { background: #dc3545; }
.actions-cell { display: flex; gap: 5px; }
.edit-btn, .delete-btn { width: 28px; height: 28px; border: none; border-radius: 4px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; padding: 0; transition: all 0.2s ease; }
.edit-btn { background-color: #ff9800; color: #fff; }
.delete-btn { background-color: #dc3545; color: #fff; }
.edit-btn:hover { background-color: #e68900; }
.delete-btn:hover { background-color: #c82333; }
.edit-btn:disabled, .delete-btn:disabled { opacity: 0.5; cursor: not-allowed; }

.manager-lookup-container {
  position: relative;
  width: 100%;
}

.manager-lookup { 
  display: flex; 
  gap: 10px; 
  margin-bottom: 5px; 
  align-items: center;
}
.manager-lookup input { 
  flex: 1; 
  padding: 10px 12px; 
  border: 1px solid #ccc; 
  border-radius: 6px; 
  font-size: 14px; 
  min-width: 0;
}
.manager-lookup input:focus { 
  border-color: #0078d4; 
  outline: none; 
  box-shadow: 0 0 3px rgba(0,120,212,0.4); 
}
.manager-lookup button { 
  background: #0078d4; 
  color: white; 
  border: none; 
  border-radius: 4px; 
  padding: 10px 15px; 
  cursor: pointer; 
  font-size: 14px; 
  font-weight: 500; 
  white-space: nowrap; 
  flex-shrink: 0;
}
.manager-lookup button:hover { 
  background: #005a9e; 
}

.manager-suggestions { 
  position: absolute;
  top: 100%;
  left: 0;
  right: 0;
  background: white; 
  border: 1px solid #ddd; 
  border-radius: 4px; 
  max-height: 200px; 
  overflow-y: auto; 
  z-index: 1000; 
  width: 100%; 
  box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
  margin-top: 2px;
}
.manager-suggestion-item { 
  padding: 10px; 
  cursor: pointer; 
  border-bottom: 1px solid #eee; 
  transition: background 0.2s ease;
}
.manager-suggestion-item:hover { 
  background: #f3f2f1; 
}
.manager-suggestion-item.selected { 
  background: #e1f5fe; 
}
.manager-suggestion-item:last-child {
  border-bottom: none;
}

.selected-manager { 
  display: flex; 
  align-items: center; 
  justify-content: space-between; 
  background: #f3f2f1; 
  padding: 10px 12px; 
  border-radius: 6px; 
  margin-top: 5px; 
  border: 1px solid #e1e1e1; 
  width: 100%;
  box-sizing: border-box;
}
.selected-manager span { 
  font-weight: 500; 
  color: #333;
  word-break: break-all;
}
.remove-manager { 
  background: none; 
  border: none; 
  color: #d13438; 
  cursor: pointer; 
  font-size: 20px; 
  padding: 0 8px; 
  line-height: 1; 
  transition: color 0.2s ease;
  flex-shrink: 0;
}
.remove-manager:hover { 
  color: #a4262c; 
}

@media (max-width:768px){
  .form-group { flex:1 1 100%; }
  .btn-primary,.btn-secondary { width:100%; justify-content:center; margin-top:10px; }
  .manager-lookup { flex-direction: column; }
  .manager-lookup button { width: 100%; }
  .modern-table { font-size: 12px; }
  .modern-table th, .modern-table td { padding: 6px; }
  .pagination-controls {
    flex-direction: column;
    align-items: stretch;
    gap: 10px;
  }
  .pagination-info {
    text-align: center;
  }
  .pagination-size {
    justify-content: center;
  }
  .pagination {
    justify-content: center;
  }
}
`,
    script: () => {
        setTimeout(async () => {
            const listName = "Project List";
            // ============================================
            // DOM ELEMENTS
            // ============================================
            const tableBody = document.querySelector('.project-table-body');
            const showFormBtn = document.querySelector('.show-form-btn');
            const formSection = document.querySelector('.form-section');
            const tableSection = document.querySelector('.table-section');
            const form = document.querySelector('#projectForm');
            const backBtn = document.querySelector('.back-to-table');
            const submitBtn = form.querySelector('button[type="submit"]');
            const exportBtn = document.querySelector('.export-btn');
            const importBtn = document.querySelector('.import-btn');
            const excelInput = document.querySelector('.excel-input');
            const searchInput = document.querySelector('.search-input');
            const resetSearchBtn = document.querySelector('.reset-search-btn');
            const formHeader = document.querySelector('.form-header h2');
            const paginationDiv = document.querySelector('.pagination');
            const pageSizeSelect = document.querySelector('.page-size-select');
            const recordCountSpan = document.querySelector('.record-count');
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/DigitalTransformationTest";
            const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
            let projectData = [];
            let filteredData = [];
            let editingId = null;
            let currentPage = 1;
            let pageSize = 25;
            // ============================================
            // INITIALIZE AUTH
            // ============================================
            try {
                await authManager.initialize({ siteUrl, listName: listName });
                console.log('🔑 Auth initialized successfully');
            }
            catch (error) {
                console.error('❌ Auth initialization failed:', error);
            }
            // ============================================
            // MANAGER LOOKUP INSTANCES
            // ============================================
            const projectManagerLookup = new ManagerLookup(document.querySelector('#projectManagerLookup'), siteUrl);
            const enggLeadLookup = new ManagerLookup(document.querySelector('#enggLeadLookup'), siteUrl);
            const enggSubLead1Lookup = new ManagerLookup(document.querySelector('#enggSubLead1Lookup'), siteUrl);
            const enggSubLead2Lookup = new ManagerLookup(document.querySelector('#enggSubLead2Lookup'), siteUrl);
            const enggManagersLookup = new ManagerLookup(document.querySelector('#enggManagersLookup'), siteUrl);
            const managerLookups = {
                projectManager: projectManagerLookup,
                enggLead: enggLeadLookup,
                enggSubLead1: enggSubLead1Lookup,
                enggSubLead2: enggSubLead2Lookup,
                enggManagers: enggManagersLookup
            };
            // ============================================
            // UTILITY FUNCTIONS
            // ============================================
            const toast = (msg, type = 'success') => {
                const toastEl = document.createElement("div");
                toastEl.className = `toast ${type}`;
                toastEl.textContent = msg;
                document.body.appendChild(toastEl);
                requestAnimationFrame(() => toastEl.style.opacity = "1");
                setTimeout(() => { toastEl.style.opacity = "0"; setTimeout(() => toastEl.remove(), 500); }, 2500);
            };
            const getDigest = async () => {
                if (typeof _spPageContextInfo !== "undefined")
                    return _spPageContextInfo.formDigestValue;
                const res = await fetch(`${siteUrl}/_api/contextinfo`, {
                    method: "POST",
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose"
                    }
                });
                const data = await res.json();
                return data.d.GetContextWebInformation.FormDigestValue;
            };
            const getEntityType = async () => {
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                        headers: {
                            'Accept': 'application/json;odata=verbose'
                        }
                    });
                    const data = await res.json();
                    return data.d.ListItemEntityTypeFullName;
                }
                catch (error) {
                    console.error('Error getting entity type:', error);
                    return 'SP.Data.Project_x0020_ListListItem';
                }
            };
            const getManagerDisplayName = (project, field = 'ProjectManager') => {
                if (project[field] && project[field].Title) {
                    return {
                        name: project[field].Title || '',
                        email: project[field].Email || ''
                    };
                }
                return { name: '', email: '' };
            };
            const clearAllLookups = () => {
                Object.values(managerLookups).forEach(lookup => lookup.clear());
            };
            // ============================================
            // FETCH PROJECTS - FIXED QUERY
            // ============================================
            const fetchProjects = async () => {
                try {
                    const selectFields = [
                        'Id',
                        'PC',
                        'ProjectReference',
                        'CustomerName',
                        'Title',
                        'GroupHead',
                        'ProjectManagerId',
                        'EnggLeadId',
                        'EnggSubLead1Id',
                        'EnggSubLead2Id',
                        'EngineeringManagersId'
                    ].join(',');
                    const expandFields = [
                        'ProjectManager',
                        'EnggLead',
                        'EnggSubLead1',
                        'EnggSubLead2',
                        'EngineeringManagers'
                    ].join(',');
                    const query = `$select=${selectFields}&$expand=${expandFields}&$orderby=Id desc`;
                    console.log('📤 Fetch query:', query);
                    const res = await fetch(`${apiBase}?${query}`, {
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose'
                        }
                    });
                    if (!res.ok) {
                        console.error('❌ First fetch failed with status:', res.status);
                        const fallbackQuery = `$select=${selectFields}&$orderby=Id desc`;
                        console.log('📤 Fallback query:', fallbackQuery);
                        const res2 = await fetch(`${apiBase}?${fallbackQuery}`, {
                            headers: {
                                'Accept': 'application/json;odata=verbose',
                                'Content-Type': 'application/json;odata=verbose'
                            }
                        });
                        if (!res2.ok) {
                            const errorText = await res2.text();
                            console.error('❌ Fallback fetch error:', errorText);
                            throw new Error(`HTTP error! status: ${res2.status}`);
                        }
                        const data2 = await res2.json();
                        projectData = data2.d.results || [];
                        console.log('📊 Projects fetched (fallback):', projectData.length);
                        await fetchManagerDetails(projectData);
                    }
                    else {
                        const data = await res.json();
                        projectData = data.d.results || [];
                        console.log('📊 Projects fetched:', projectData.length);
                        const hasExpandedData = projectData.some(p => p.ProjectManager && p.ProjectManager.Title);
                        if (!hasExpandedData) {
                            console.log('📊 Expanded data not available, fetching manager details separately...');
                            await fetchManagerDetails(projectData);
                        }
                    }
                    filteredData = getAssignedProjects(projectData);
                    console.log(`📋 Showing ${filteredData.length} projects (filtered by role)`);
                    currentPage = 1;
                    renderTable(filteredData);
                    updateUIBasedOnPermissions({
                        importBtn: '.import-btn',
                        exportBtn: '.export-btn',
                        addBtn: '.show-form-btn'
                    });
                }
                catch (err) {
                    console.error("Error fetching projects:", err);
                    toast("Error fetching projects. Please check console for details.", "error");
                }
            };
            const fetchManagerDetails = async (projects) => {
                for (const project of projects) {
                    const managerFields = [
                        { id: 'ProjectManagerId', field: 'ProjectManager' },
                        { id: 'EnggLeadId', field: 'EnggLead' },
                        { id: 'EnggSubLead1Id', field: 'EnggSubLead1' },
                        { id: 'EnggSubLead2Id', field: 'EnggSubLead2' },
                        { id: 'EngineeringManagersId', field: 'EngineeringManagers' }
                    ];
                    for (const { id, field } of managerFields) {
                        const managerId = project[id];
                        if (managerId && !project[field]) {
                            try {
                                const managerRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${managerId})?$select=Id,Title,Email`, {
                                    headers: {
                                        'Accept': 'application/json;odata=verbose',
                                        'Content-Type': 'application/json;odata=verbose'
                                    }
                                });
                                if (managerRes.ok) {
                                    const managerData = await managerRes.json();
                                    project[field] = {
                                        Id: managerData.d.Id,
                                        Title: managerData.d.Title,
                                        Email: managerData.d.Email || ''
                                    };
                                }
                            }
                            catch (error) {
                                console.warn(`Could not fetch manager details for ${field} ID ${managerId}:`, error);
                            }
                        }
                    }
                }
            };
            // ============================================
            // SAVE PROJECT
            // ============================================
            const saveProject = async (payload, isUpdate = false) => {
                const duplicate = projectData.find(p => {
                    var _a;
                    return ((_a = p.Title) === null || _a === void 0 ? void 0 : _a.toLowerCase()) === payload.Title.toLowerCase() &&
                        (!isUpdate || p.Id !== editingId);
                });
                if (duplicate) {
                    toast("Project with this name already exists", "error");
                    return;
                }
                try {
                    const digest = await getDigest();
                    const entityType = await getEntityType();
                    const url = isUpdate ? `${apiBase}(${editingId})` : apiBase;
                    const headers = {
                        "Accept": "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": digest
                    };
                    if (isUpdate) {
                        if (editingId === null) {
                            toast("No project selected", "error");
                            return;
                        }
                        headers["X-HTTP-Method"] = "MERGE";
                        headers["IF-MATCH"] = "*";
                    }
                    const finalPayload = {
                        __metadata: { type: entityType },
                        PC: payload.PC,
                        ProjectReference: payload.ProjectReference,
                        CustomerName: payload.CustomerName,
                        Title: payload.Title,
                        GroupHead: payload.GroupHead
                    };
                    const managerFields = [
                        { field: 'ProjectManagerId', value: payload.ProjectManagerId },
                        { field: 'EnggLeadId', value: payload.EnggLeadId },
                        { field: 'EnggSubLead1Id', value: payload.EnggSubLead1Id },
                        { field: 'EnggSubLead2Id', value: payload.EnggSubLead2Id },
                        { field: 'EngineeringManagersId', value: payload.EngineeringManagersId }
                    ];
                    for (const { field, value } of managerFields) {
                        if (value && value !== '' && value !== 'null' && value !== 'undefined') {
                            finalPayload[field] = parseInt(value);
                        }
                        else {
                            finalPayload[field] = null;
                        }
                    }
                    console.log("📤 Saving payload:", JSON.stringify(finalPayload, null, 2));
                    const res = await fetch(url, {
                        method: "POST",
                        headers,
                        body: JSON.stringify(finalPayload)
                    });
                    if (!res.ok) {
                        const errorText = await res.text();
                        console.error("❌ Save error response:", errorText);
                        toast(isUpdate ? "Update failed" : "Add failed", "error");
                    }
                    else {
                        toast(isUpdate ? "✅ Project updated successfully" : "✅ Project added successfully", "success");
                        editingId = null;
                        setTimeout(() => fetchProjects(), 1000);
                    }
                }
                catch (err) {
                    console.error("❌ Save error:", err);
                    toast(isUpdate ? "Update failed" : "Add failed", "error");
                }
            };
            // ============================================
            // DELETE PROJECT - USING NATIVE CONFIRM
            // ============================================
            const deleteProject = async (id) => {
                console.log('🔍 deleteProject called with ID:', id);
                if (!hasFullControl()) {
                    console.warn('❌ User does not have full control permission');
                    toast("You don't have permission to delete projects", "error");
                    return;
                }
                const project = projectData.find(p => p.Id === id);
                if (!project) {
                    console.error('❌ Project not found with ID:', id);
                    toast("Project not found", "error");
                    return;
                }
                const projectName = project.Title || 'this project';
                console.log('🗑️ Preparing to delete project:', projectName);
                const confirmed = confirm(`⚠️ Delete Project\n\nAre you sure you want to delete "${projectName}"?\n\nThis action cannot be undone!`);
                if (!confirmed) {
                    toast("Delete cancelled", "error");
                    return;
                }
                console.log('✅ User confirmed deletion');
                try {
                    toast(`⏳ Deleting "${projectName}"...`, "success");
                    const digest = await getDigest();
                    const response = await fetch(`${apiBase}(${id})`, {
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest,
                            "X-HTTP-Method": "DELETE",
                            "IF-MATCH": "*"
                        }
                    });
                    console.log('📥 Delete response status:', response.status);
                    if (response.ok) {
                        toast(`✅ Project "${projectName}" deleted successfully`, "success");
                        await fetchProjects();
                    }
                    else {
                        let errorMessage = 'Failed to delete project';
                        try {
                            const errorData = await response.json();
                            if (errorData.error && errorData.error.message) {
                                errorMessage = errorData.error.message;
                            }
                        }
                        catch (e) {
                            errorMessage = response.statusText || errorMessage;
                        }
                        console.error('❌ Delete error:', errorMessage);
                        toast(`❌ Delete failed: ${errorMessage}`, "error");
                    }
                }
                catch (err) {
                    console.error("❌ Delete error:", err);
                    toast("Delete request failed. Please check your connection and try again.", "error");
                }
            };
            // ============================================
            // RENDER TABLE
            // ============================================
            const renderTable = (data) => {
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, data.length);
                const paginatedData = data.slice(startIndex, endIndex);
                if (paginatedData.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="11" style="text-align:center;padding:40px;color:#999;">No projects found</td></tr>`;
                    updatePaginationInfo(0);
                    renderPagination(0);
                    return;
                }
                const hasFullControlPermission = hasFullControl();
                console.log('🔑 Has full control for delete:', hasFullControlPermission);
                tableBody.innerHTML = paginatedData.map(p => {
                    const pm = getManagerDisplayName(p, 'ProjectManager');
                    const enggLead = getManagerDisplayName(p, 'EnggLead');
                    const enggSubLead1 = getManagerDisplayName(p, 'EnggSubLead1');
                    const enggSubLead2 = getManagerDisplayName(p, 'EnggSubLead2');
                    const enggManagers = getManagerDisplayName(p, 'EngineeringManagers');
                    const canEdit = hasEditPermission(p);
                    const editBtnDisabled = !canEdit ? 'disabled' : '';
                    const renderManagerCell = (manager) => {
                        if (!manager.name && !manager.email)
                            return '';
                        return `
              <div class="manager-cell">
                <span class="manager-name">${manager.name}</span>
                ${manager.email ? `<span class="manager-email">${manager.email}</span>` : ''}
              </div>
            `;
                    };
                    const deleteButtonHtml = hasFullControlPermission
                        ? `<button class="delete-btn" data-id="${p.Id}" title="Delete Project">🗑️</button>`
                        : '';
                    return `
            <tr data-project-id="${p.Id}">
              <td>${p.PC || ""}</td>
              <td>${p.ProjectReference || ""}</td>
              <td>${p.CustomerName || ""}</td>
              <td>${p.Title || ""}</td>
              <td>${p.GroupHead || ""}</td>
              <td>${renderManagerCell(pm)}</td>
              <td>${renderManagerCell(enggLead)}</td>
              <td>${renderManagerCell(enggSubLead1)}</td>
              <td>${renderManagerCell(enggSubLead2)}</td>
              <td>${renderManagerCell(enggManagers)}</td>
              <td class="actions-cell">
                <button class="edit-btn" data-id="${p.Id}" title="Edit Project" ${editBtnDisabled}>✏️</button>
                ${deleteButtonHtml}
              </td>
            </tr>
          `;
                }).join("");
                attachTableHandlers();
                updatePaginationInfo(data.length);
                const totalPages = Math.ceil(data.length / pageSize);
                renderPagination(totalPages);
            };
            const updatePaginationInfo = (totalRecords) => {
                if (totalRecords === 0) {
                    recordCountSpan.textContent = 'No records found';
                    return;
                }
                const startIndex = (currentPage - 1) * pageSize + 1;
                const endIndex = Math.min(currentPage * pageSize, totalRecords);
                recordCountSpan.textContent = `Showing ${startIndex} to ${endIndex} of ${totalRecords} records`;
            };
            const renderPagination = (totalPages) => {
                paginationDiv.innerHTML = '';
                if (totalPages <= 1) {
                    return;
                }
                const prevBtn = document.createElement('button');
                prevBtn.textContent = '‹';
                prevBtn.disabled = currentPage === 1;
                prevBtn.addEventListener('click', () => {
                    if (currentPage > 1) {
                        currentPage--;
                        renderTable(filteredData);
                    }
                });
                paginationDiv.appendChild(prevBtn);
                const maxVisiblePages = 5;
                let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
                let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
                if (endPage - startPage < maxVisiblePages - 1) {
                    startPage = Math.max(1, endPage - maxVisiblePages + 1);
                }
                if (startPage > 1) {
                    const firstBtn = document.createElement('button');
                    firstBtn.textContent = '1';
                    firstBtn.addEventListener('click', () => {
                        currentPage = 1;
                        renderTable(filteredData);
                    });
                    paginationDiv.appendChild(firstBtn);
                    if (startPage > 2) {
                        const ellipsis = document.createElement('span');
                        ellipsis.textContent = '…';
                        ellipsis.style.padding = '0 5px';
                        paginationDiv.appendChild(ellipsis);
                    }
                }
                for (let i = startPage; i <= endPage; i++) {
                    const btn = document.createElement('button');
                    btn.textContent = i.toString();
                    btn.className = i === currentPage ? 'active' : '';
                    btn.addEventListener('click', () => {
                        currentPage = i;
                        renderTable(filteredData);
                    });
                    paginationDiv.appendChild(btn);
                }
                if (endPage < totalPages) {
                    if (endPage < totalPages - 1) {
                        const ellipsis = document.createElement('span');
                        ellipsis.textContent = '…';
                        ellipsis.style.padding = '0 5px';
                        paginationDiv.appendChild(ellipsis);
                    }
                    const lastBtn = document.createElement('button');
                    lastBtn.textContent = totalPages.toString();
                    lastBtn.addEventListener('click', () => {
                        currentPage = totalPages;
                        renderTable(filteredData);
                    });
                    paginationDiv.appendChild(lastBtn);
                }
                const nextBtn = document.createElement('button');
                nextBtn.textContent = '›';
                nextBtn.disabled = currentPage === totalPages;
                nextBtn.addEventListener('click', () => {
                    if (currentPage < totalPages) {
                        currentPage++;
                        renderTable(filteredData);
                    }
                });
                paginationDiv.appendChild(nextBtn);
            };
            const attachTableHandlers = () => {
                tableBody.removeEventListener('click', handleTableClick);
                tableBody.addEventListener('click', handleTableClick);
                console.log('✅ Table event listeners attached');
            };
            const handleTableClick = (event) => {
                const target = event.target;
                const deleteBtn = target.closest('.delete-btn');
                if (deleteBtn) {
                    event.preventDefault();
                    event.stopPropagation();
                    const id = Number(deleteBtn.getAttribute('data-id'));
                    console.log('🗑️ Delete clicked for ID:', id);
                    if (!id || isNaN(id)) {
                        console.error('Invalid project ID');
                        toast("Invalid project ID", "error");
                        return;
                    }
                    deleteProject(id);
                    return;
                }
                const editBtn = target.closest('.edit-btn');
                if (editBtn) {
                    event.preventDefault();
                    event.stopPropagation();
                    const id = Number(editBtn.getAttribute('data-id'));
                    console.log('✏️ Edit clicked for ID:', id);
                    if (!id || isNaN(id)) {
                        console.error('Invalid project ID');
                        toast("Invalid project ID", "error");
                        return;
                    }
                    handleEditClickById(id);
                    return;
                }
            };
            const handleEditClickById = async (id) => {
                console.log('✏️ Edit requested for project ID:', id);
                const proj = projectData.find(x => x.Id === id);
                if (!proj) {
                    console.error('❌ Project not found with ID:', id);
                    toast("Project not found", "error");
                    return;
                }
                if (!hasEditPermission(proj)) {
                    toast("You don't have permission to edit this project", "error");
                    return;
                }
                document.querySelector('#pc').value = proj.PC || '';
                document.querySelector('#reference').value = proj.ProjectReference || '';
                document.querySelector('#customer').value = proj.CustomerName || '';
                document.querySelector('#name').value = proj.Title || '';
                document.querySelector('#head').value = proj.GroupHead || '';
                clearAllLookups();
                const setLookup = (lookup, managerData) => {
                    if (managerData && managerData.Id && managerData.Title) {
                        lookup.setSelected({
                            id: managerData.Id,
                            title: managerData.Title,
                            email: managerData.Email || ''
                        });
                    }
                };
                setLookup(projectManagerLookup, proj.ProjectManager);
                setLookup(enggLeadLookup, proj.EnggLead);
                setLookup(enggSubLead1Lookup, proj.EnggSubLead1);
                setLookup(enggSubLead2Lookup, proj.EnggSubLead2);
                setLookup(enggManagersLookup, proj.EngineeringManagers);
                editingId = proj.Id;
                formHeader.textContent = "✏️ Update Project";
                submitBtn.textContent = "Update Project";
                formSection.style.display = "block";
                tableSection.style.display = "none";
                showFormBtn.style.display = "none";
            };
            // ============================================
            // SEARCH AND FILTER
            // ============================================
            const applySearchAndFilter = () => {
                const query = searchInput.value.toLowerCase();
                const searchedData = projectData.filter(p => {
                    const pm = getManagerDisplayName(p, 'ProjectManager');
                    const enggLead = getManagerDisplayName(p, 'EnggLead');
                    const enggSubLead1 = getManagerDisplayName(p, 'EnggSubLead1');
                    const enggSubLead2 = getManagerDisplayName(p, 'EnggSubLead2');
                    const enggManagers = getManagerDisplayName(p, 'EngineeringManagers');
                    return ((p.PC || '').toLowerCase().includes(query) ||
                        (p.ProjectReference || '').toLowerCase().includes(query) ||
                        (p.CustomerName || '').toLowerCase().includes(query) ||
                        (p.Title || '').toLowerCase().includes(query) ||
                        (p.GroupHead || '').toLowerCase().includes(query) ||
                        (pm.name || '').toLowerCase().includes(query) ||
                        (pm.email || '').toLowerCase().includes(query) ||
                        (enggLead.name || '').toLowerCase().includes(query) ||
                        (enggLead.email || '').toLowerCase().includes(query) ||
                        (enggSubLead1.name || '').toLowerCase().includes(query) ||
                        (enggSubLead1.email || '').toLowerCase().includes(query) ||
                        (enggSubLead2.name || '').toLowerCase().includes(query) ||
                        (enggSubLead2.email || '').toLowerCase().includes(query) ||
                        (enggManagers.name || '').toLowerCase().includes(query) ||
                        (enggManagers.email || '').toLowerCase().includes(query));
                });
                filteredData = getAssignedProjects(searchedData);
                currentPage = 1;
                renderTable(filteredData);
            };
            searchInput.addEventListener('input', () => {
                applySearchAndFilter();
            });
            resetSearchBtn.addEventListener('click', () => {
                searchInput.value = '';
                applySearchAndFilter();
            });
            pageSizeSelect.addEventListener('change', () => {
                pageSize = parseInt(pageSizeSelect.value);
                currentPage = 1;
                renderTable(filteredData);
            });
            // ============================================
            // FORM SUBMIT
            // ============================================
            form.addEventListener('submit', async (e) => {
                e.preventDefault();
                if (!canCreate()) {
                    toast("You don't have permission to create projects", "error");
                    return;
                }
                const payload = {
                    PC: document.querySelector('#pc').value,
                    ProjectReference: document.querySelector('#reference').value,
                    CustomerName: document.querySelector('#customer').value,
                    Title: document.querySelector('#name').value,
                    GroupHead: document.querySelector('#head').value,
                    ProjectManagerId: projectManagerLookup.getSelectedId(),
                    EnggLeadId: enggLeadLookup.getSelectedId(),
                    EnggSubLead1Id: enggSubLead1Lookup.getSelectedId(),
                    EnggSubLead2Id: enggSubLead2Lookup.getSelectedId(),
                    EngineeringManagersId: enggManagersLookup.getSelectedId()
                };
                await saveProject(payload, editingId !== null);
                form.reset();
                clearAllLookups();
                formSection.style.display = "none";
                tableSection.style.display = "block";
                showFormBtn.style.display = "inline-block";
                submitBtn.textContent = "Save Project";
                editingId = null;
            });
            // ============================================
            // SHOW/HIDE FORM
            // ============================================
            showFormBtn.addEventListener('click', () => {
                if (!canCreate()) {
                    toast("You don't have permission to create projects", "error");
                    return;
                }
                form.reset();
                clearAllLookups();
                editingId = null;
                formSection.style.display = "block";
                tableSection.style.display = "none";
                showFormBtn.style.display = "none";
                submitBtn.textContent = "Add Project";
                formHeader.textContent = "➕ Add New Project";
            });
            backBtn.addEventListener('click', () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                showFormBtn.style.display = "inline-block";
                editingId = null;
                form.reset();
                clearAllLookups();
                submitBtn.textContent = "Save Project";
            });
            // ============================================
            // EXCEL EXPORT
            // ============================================
            exportBtn.addEventListener('click', () => {
                if (!canRead()) {
                    toast("You don't have permission to export", "error");
                    return;
                }
                if (!projectData.length) {
                    toast("No data to export", "error");
                    return;
                }
                const wsData = projectData.map(p => {
                    const pm = getManagerDisplayName(p, 'ProjectManager');
                    const enggLead = getManagerDisplayName(p, 'EnggLead');
                    const enggSubLead1 = getManagerDisplayName(p, 'EnggSubLead1');
                    const enggSubLead2 = getManagerDisplayName(p, 'EnggSubLead2');
                    const enggManagers = getManagerDisplayName(p, 'EngineeringManagers');
                    return {
                        'Profit Centre': p.PC || '',
                        'Project Reference': p.ProjectReference || '',
                        'Customer Name': p.CustomerName || '',
                        'Project Name': p.Title || '',
                        'GH': p.GroupHead || '',
                        'Project Manager': pm.name + (pm.email ? ` (${pm.email})` : ''),
                        'Engg Lead': enggLead.name + (enggLead.email ? ` (${enggLead.email})` : ''),
                        'Engg Sub Lead 1': enggSubLead1.name + (enggSubLead1.email ? ` (${enggSubLead1.email})` : ''),
                        'Engg Sub Lead 2': enggSubLead2.name + (enggSubLead2.email ? ` (${enggSubLead2.email})` : ''),
                        'Engineering Managers': enggManagers.name + (enggManagers.email ? ` (${enggManagers.email})` : '')
                    };
                });
                const ws = XLSX.utils.json_to_sheet(wsData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Projects");
                const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
                saveAs(new Blob([buf], { type: "application/octet-stream" }), `projects_${new Date().toISOString().slice(0, 10)}.xlsx`);
                toast("Excel exported successfully", "success");
            });
            // ============================================
            // EXCEL IMPORT
            // ============================================
            importBtn.addEventListener('click', () => {
                if (!canCreate()) {
                    toast("You don't have permission to import", "error");
                    return;
                }
                excelInput.click();
            });
            excelInput.addEventListener('change', async (e) => {
                var _a;
                const file = (_a = e.target.files) === null || _a === void 0 ? void 0 : _a[0];
                if (!file)
                    return;
                Swal.fire({
                    title: 'Importing...',
                    text: 'Please wait while importing data',
                    allowOutsideClick: false,
                    showConfirmButton: false,
                    willOpen: () => {
                        Swal.showLoading();
                    }
                });
                try {
                    const data = await file.arrayBuffer();
                    const workbook = XLSX.read(data);
                    const sheetName = workbook.SheetNames[0];
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    let successCount = 0;
                    let errorCount = 0;
                    for (const row of jsonData) {
                        try {
                            const payload = {
                                PC: row["Profit Centre"] || '',
                                ProjectReference: row["Project Reference"] || '',
                                CustomerName: row["Customer Name"] || '',
                                Title: row["Project Name"] || '',
                                GroupHead: row["GH"] || ''
                            };
                            const managerFields = [
                                { excelField: "Project Manager", payloadField: "ProjectManagerId" },
                                { excelField: "Engg Lead", payloadField: "EnggLeadId" },
                                { excelField: "Engg Sub Lead 1", payloadField: "EnggSubLead1Id" },
                                { excelField: "Engg Sub Lead 2", payloadField: "EnggSubLead2Id" },
                                { excelField: "Engineering Managers", payloadField: "EngineeringManagersId" }
                            ];
                            for (const { excelField, payloadField } of managerFields) {
                                const managerName = row[excelField];
                                if (managerName && managerName.trim() !== "") {
                                    const emailMatch = managerName.match(/\(([^)]+)\)/);
                                    const searchTerm = emailMatch ? emailMatch[1] : managerName;
                                    const managers = await searchManagers(searchTerm);
                                    if (managers.length > 0) {
                                        payload[payloadField] = managers[0].id;
                                    }
                                }
                            }
                            await saveProject(payload, false);
                            successCount++;
                        }
                        catch (error) {
                            errorCount++;
                            console.error("Import error for row:", row, error);
                        }
                    }
                    Swal.fire({
                        title: 'Import Complete',
                        text: `Successfully imported ${successCount} projects. ${errorCount} failed.`,
                        icon: 'success',
                        confirmButtonColor: '#0078d4'
                    });
                    fetchProjects();
                }
                catch (error) {
                    console.error("Import error:", error);
                    Swal.fire({
                        title: 'Import Failed',
                        text: 'Error reading Excel file',
                        icon: 'error',
                        confirmButtonColor: '#0078d4'
                    });
                }
                excelInput.value = "";
            });
            // ============================================
            // SEARCH MANAGERS (for import)
            // ============================================
            const searchManagers = async (searchTerm = '') => {
                try {
                    let url = `${siteUrl}/_api/web/siteusers?$filter=PrincipalType eq 1`;
                    if (searchTerm) {
                        url += ` and (substringof('${searchTerm}', Title) eq true or substringof('${searchTerm}', Email) eq true)`;
                    }
                    url += `&$top=20&$orderby=Title`;
                    const res = await fetch(url, {
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose'
                        }
                    });
                    if (!res.ok) {
                        throw new Error(`HTTP error! status: ${res.status}`);
                    }
                    const data = await res.json();
                    const users = data.d.results || [];
                    return users
                        .filter((user) => {
                        const title = user.Title || '';
                        return !title.includes('System Account') &&
                            !title.includes('App@sharepoint') &&
                            !title.includes('SharingLinks') &&
                            !title.includes('NT AUTHORITY') &&
                            !title.includes('Membership');
                    })
                        .map((user) => ({
                        id: user.Id,
                        title: user.Title || '',
                        email: user.Email || ''
                    }))
                        .sort((a, b) => a.title.localeCompare(b.title));
                }
                catch (error) {
                    console.error("Error searching managers:", error);
                    return [];
                }
            };
            // ============================================
            // INITIALIZE
            // ============================================
            fetchProjects();
        }, 300);
    }
};
//# sourceMappingURL=project.js.map
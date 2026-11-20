// project.ts - Complete file with SCM & Controlling fields added
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import "../styles/common.css";
import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css";
import { authManager, canCreate, canRead, canEditProject, getAssignedProjects, 
// updateUIBasedOnPermissions,
ROLE_GROUP_NAMES } from './auth';
// ============================================
// PERMISSION HELPERS
// ============================================
// const hasEditPermission = (project?: any): boolean => {
//   return canEditProject(project);
// };
const hasEditPermission = (_project) => {
    return canEditProject(undefined); // undefined ⇒ no per-project bypass ⇒ Admin only
};
// ============================================
// MANAGER LOOKUP COMPONENT
// ============================================
class ManagerLookup {
    constructor(container, siteUrl, groupName, options) {
        this.container = container;
        this.siteUrl = siteUrl;
        this.groupName = groupName;
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
            // Show inline hint instead of Swal
            this.suggestionsDiv.innerHTML =
                '<div class="manager-suggestion-item">Type at least 2 characters to search…</div>';
            this.suggestionsDiv.style.display = 'block';
            return;
        }
        try {
            const managers = await this.searchManagers(searchTerm);
            this.displaySuggestions(managers);
        }
        catch (error) {
            console.error('Error searching managers:', error);
            this.suggestionsDiv.innerHTML =
                '<div class="manager-suggestion-item" style="color:#d13438;">Search failed. Please try again.</div>';
            this.suggestionsDiv.style.display = 'block';
        }
    }
    /** Group-aware search */
    async searchManagers(searchTerm) {
        var _a;
        const safeTerm = searchTerm.replace(/'/g, "''");
        if (this.groupName) {
            const safeGroup = this.groupName.replace(/'/g, "''");
            const groupUrl = `${this.siteUrl}/_api/web/sitegroups/getbyname('${safeGroup}')/users` +
                `?$filter=PrincipalType eq 1 and (substringof('${safeTerm}', Title) eq true or substringof('${safeTerm}', Email) eq true)` +
                `&$top=20&$orderby=Title&$select=Id,Title,Email`;
            const res = await fetch(groupUrl, {
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose'
                }
            });
            if (res.ok) {
                const data = await res.json();
                const users = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || data.value || [];
                return users
                    .filter((u) => !this.isServiceAccount(u.Title || ''))
                    .map((u) => ({ id: u.Id, title: u.Title || '', email: u.Email || '' }))
                    .sort((a, b) => a.title.localeCompare(b.title));
            }
            console.warn(`Group "${this.groupName}" not found or not readable; falling back to site-user search.`);
        }
        const url = `${this.siteUrl}/_api/web/siteusers` +
            `?$filter=PrincipalType eq 1 and (substringof('${safeTerm}', Title) eq true or substringof('${safeTerm}', Email) eq true)` +
            `&$top=20&$orderby=Title&$select=Id,Title,Email`;
        const res = await fetch(url, {
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
            }
        });
        if (!res.ok)
            throw new Error(`HTTP error! status: ${res.status}`);
        const data = await res.json();
        const users = data.d.results || [];
        return users
            .filter((user) => !this.isServiceAccount(user.Title || ''))
            .map((user) => ({ id: user.Id, title: user.Title || '', email: user.Email || '' }))
            .sort((a, b) => a.title.localeCompare(b.title));
    }
    isServiceAccount(title) {
        return !title ||
            title.includes('System Account') ||
            title.includes('App@sharepoint') ||
            title.includes('SharingLinks') ||
            title.includes('NT AUTHORITY') ||
            title.includes('Membership');
    }
    displaySuggestions(managers) {
        this.suggestionsDiv.innerHTML = '';
        if (managers.length === 0) {
            const noResults = document.createElement('div');
            noResults.className = 'manager-suggestion-item';
            noResults.textContent = `No users found${this.groupName ? ` in "${this.groupName}"` : ''}`;
            this.suggestionsDiv.appendChild(noResults);
        }
        else {
            managers.forEach(manager => {
                const suggestion = document.createElement('div');
                suggestion.className = 'manager-suggestion-item';
                suggestion.innerHTML = `
          <strong>${manager.title}</strong>
          ${manager.email ? `<small>${manager.email}</small>` : ''}
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
        const fg = this.container.closest('.form-group');
        if (fg) {
            fg.classList.remove('has-error');
            const err = fg.querySelector('.field-error');
            if (err)
                err.remove();
        }
        if (this.onSelect)
            this.onSelect(manager);
    }
    clear() {
        this.searchInput.style.display = 'block';
        this.searchBtn.style.display = 'block';
        this.selectedDisplay.style.display = 'none';
        this.hiddenInput.value = '';
        this.searchInput.value = '';
        this.suggestionsDiv.style.display = 'none';
        if (this.onClear)
            this.onClear();
    }
    getSelectedId() {
        return this.hiddenInput.value || null;
    }
    getSelectedName() {
        return this.selectedNameSpan.textContent || '';
    }
    setSelected(manager) {
        if (manager && manager.id)
            this.select(manager);
        else
            this.clear();
    }
}
// ============================================
// MAIN PROJECT MANAGEMENT COMPONENT
// ============================================
const projectContent = `
<div class="project-page fade-in">
<h1 class="project-title">Upload projects</h1>
<div class="page-header">
<div class="page-header-actions">
  <button class="btn-primary show-form-btn admin-only-btn" hidden>+ Add New Project</button>
  <button class="btn-primary export-btn">⬇ Download Data</button>
  <button class="btn-primary import-btn admin-only-btn" hidden>Upload Multiple Projects</button>
</div>
    </div>
    <div class="table-actions-project">
      <div class="table-actions-row-project">
        <div class="tableColp4-project">
          <label for="projectSearch">Search Projects</label>
          <input type="text" id="projectSearch" class="search-input-project" placeholder="Search by project name, reference, profit centre, customer or any person..." />
        </div>
        <button type="button" class="reset-search-btn-project">Reset Filter</button>
      </div>
    </div>
  <section class="table-section card">
    
    <!-- Top Pagination -->
    <div class="pagination-controls top-pagination" style="margin-top: 0; margin-bottom: 16px; border-top: none; border-bottom: 1px solid #eee; padding-top: 0; padding-bottom: 15px;">
      <div class="pagination-left">
        <div class="pagination-size">
          <label for="pjPageSizeTop">Items per page:</label>
          <select class="page-size-select top-page-size" id="pjPageSizeTop">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>
        <div class="pagination-info">
          Showing <span class="pj-pg-range"><span class="record-start">0</span> - <span class="record-end">0</span></span> of <span class="pj-pg-count record-count">0</span> items
        </div>
      </div>
      <div class="pagination-right">
        <button type="button" class="pj-page-btn pj-pg-icon first-page" title="First page" disabled>«</button>
        <button type="button" class="pj-page-btn pj-pg-icon prev-page" title="Previous page" disabled>‹</button>
        <div class="pagination pj-page-numbers page-numbers"></div>
        <span class="pj-pg-of">of <span class="total-pages">1</span></span>
        <button type="button" class="pj-page-btn pj-pg-icon next-page" title="Next page" disabled>›</button>
        <button type="button" class="pj-page-btn pj-pg-icon last-page" title="Last page" disabled>»</button>
      </div>
    </div>

    <div class="pj-table-scroll">
      <table class="modern-table">
        <colgroup>
          <col class="pj-col-pc" />
          <col class="pj-col-ref" />
          <col class="pj-col-customer" />
          <col class="pj-col-project" />
          <col class="pj-col-person" />
          <col class="pj-col-person" />
          <col class="pj-col-person" />
          <col class="pj-col-person" />
          <col class="pj-col-person" />
          <col class="pj-col-person-lg" />
          <col class="pj-col-person" />
          <col class="pj-col-person" />
          <col class="pj-col-actions" />
        </colgroup>
        <thead>
          <tr>
            <th class="pj-th pj-th-text">Profit Centre</th>
            <th class="pj-th pj-th-text">Project Reference</th>
            <th class="pj-th pj-th-text">Customer Name</th>
            <th class="pj-th pj-th-project">Project Name</th>
            <th class="pj-th pj-th-person">Group Head</th>
            <th class="pj-th pj-th-person">Project Manager</th>
            <th class="pj-th pj-th-person">Engg Lead</th>
            <th class="pj-th pj-th-person">Engg Sub Lead 1</th>
            <th class="pj-th pj-th-person">Engg Sub Lead 2</th>
            <th class="pj-th pj-th-person">Engineering Managers</th>
            <th class="pj-th pj-th-person">SCM</th>
            <th class="pj-th pj-th-person">Controlling</th>
            <th class="pj-th pj-th-actions"></th>
          </tr>
        </thead>
        <tbody class="project-table-body"></tbody>
      </table>
    </div>
    
    <!-- Bottom Pagination -->
    <div class="pagination-controls bottom-pagination">
      <div class="pagination-left">
        <div class="pagination-size">
          <label for="pjPageSize">Items per page:</label>
          <select class="page-size-select bottom-page-size" id="pjPageSize">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>
        <div class="pagination-info">
          Showing <span class="pj-pg-range"><span class="record-start">0</span> - <span class="record-end">0</span></span> of <span class="pj-pg-count record-count">0</span> items
        </div>
      </div>
      <div class="pagination-right">
        <button type="button" class="pj-page-btn pj-pg-icon first-page" title="First page" disabled>«</button>
        <button type="button" class="pj-page-btn pj-pg-icon prev-page" title="Previous page" disabled>‹</button>
        <div class="pagination pj-page-numbers page-numbers"></div>
        <span class="pj-pg-of">of <span class="total-pages">1</span></span>
        <button type="button" class="pj-page-btn pj-pg-icon next-page" title="Next page" disabled>›</button>
        <button type="button" class="pj-page-btn pj-pg-icon last-page" title="Last page" disabled>»</button>
      </div>
    </div>
  </section>

  <section class="form-section card" style="display:none;">
    <!-- Top Bar: Back Button and Save/Update Button -->
    <div class="form-topbar">
      <div class="form-topbar-spacer"></div>
      <div class="form-topbar-actions">
        <button type="button" class="btn-secondary back-to-table">Back to List</button>
        <button type="button" class="btn-primary form-update-btn">Save Project</button>
      </div>
    </div>

    <form id="projectForm" novalidate>
      <div class="form-row">
        <div class="form-group">
          <label>Profit Centre <span class="req">*</span></label>
          <input type="text" id="pc" required />
        </div>
        <div class="form-group">
          <label>Project Reference <span class="req">*</span></label>
          <input type="text" id="reference" required />
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Customer Name <span class="req">*</span></label>
          <input type="text" id="customer" required />
        </div>
        <div class="form-group">
          <label>Project Name <span class="req">*</span></label>
          <input type="text" id="name" required />
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Group Head <span class="req">*</span></label>
          <div class="manager-lookup-container" id="groupHeadLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search group heads..." />
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
          <label>Project Manager <span class="req">*</span></label>
          <div class="manager-lookup-container" id="projectManagerLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search project managers..." />
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
              <input type="text" class="manager-search-input" placeholder="Search engg leads..." />
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
              <input type="text" class="manager-search-input" placeholder="Search engg sub leads..." />
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
              <input type="text" class="manager-search-input" placeholder="Search engg sub leads..." />
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
              <input type="text" class="manager-search-input" placeholder="Search engg managers..." />
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
          <label>SCM</label>
          <div class="manager-lookup-container" id="scmLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search SCM team..." />
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
          <label>Controlling</label>
          <div class="manager-lookup-container" id="controllingLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search controlling team..." />
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
    </form>
  </section>
</div>
`;
export default {
    content: projectContent,
    style: `
  @font-face {
  font-family: 'ABBvoiceOffice';
  src: url('../../assets/ABBvoiceOffice_Rg.ttf') format('truetype');
  font-weight: 600;
  font-style: normal;
  font-display: swap;
}

.project-page { font-family:'ABBvoiceOffice'; }

.dt-body .project-page > .project-title {
  margin-bottom: 14px !important;
  padding-left: 28px !important;
}

.project-page .page-header {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 10px;
  margin-bottom: 18px;
}
.project-page .page-header h1 { margin: 0; color: #323130; }
.project-page .page-header > div {
  display: flex;
  align-items: center;
  justify-content: flex-start;
  flex-wrap: nowrap;
  gap: 8px;
  width: 100%;
  max-width: 100%;
  overflow-x: auto;
  padding-bottom: 3px;
}
.project-page .page-header > div::-webkit-scrollbar { height: 5px; }
.project-page .page-header > div::-webkit-scrollbar-track { background: transparent; }
.project-page .page-header > div::-webkit-scrollbar-thumb { background: #c5c5d2; border-radius: 8px; }
.project-page .page-header .btn-primary {
  flex: 0 0 auto;
  white-space: nowrap;
}

.card { background:#fff; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,0.08); padding:20px; margin-bottom:20px; transition:all 0.3s ease; position:relative; }

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

.project-page .modern-table { min-width: 900px; }

.project-page .project-ref-link {
  color: #0078d4;
  text-decoration: none;
  font-weight: 500;
  cursor: pointer;
  border-bottom: 1px dashed transparent;
  transition: color 0.15s ease, border-color 0.15s ease;
}
.project-page .project-ref-link:hover {
  color: #005a9e;
  border-bottom-color: #005a9e;
}
.project-page .project-ref-link-disabled {
  color: #8a94a0;
  cursor: not-allowed;
  border-bottom-color: transparent;
}
.project-page .project-ref-link-disabled:hover {
  color: #8a94a0;
  border-bottom-color: transparent;
}

.manager-cell { display:flex; flex-direction:column; gap:2px; }
.manager-name { font-weight:500; color:#333; }
.manager-email { font-size:11px; color:#666; word-break:break-all; }

.form-row { display:flex; flex-wrap:wrap; gap:20px; margin-bottom:20px; }
.form-group { flex:1 1 calc(50% - 10px); display:flex; flex-direction:column; }
.form-group label { font-weight:500; color:#444; margin-bottom:6px; }
.form-group input, .form-group select { width:100%; padding:10px 12px; border-radius:6px; border:1px solid #ccc; font-size:14px; line-height:1.4; box-sizing:border-box; transition:all 0.2s ease; }
.form-group input:focus, .form-group select:focus { border-color:#0078d4; box-shadow:0 0 3px rgba(0,120,212,0.4); outline:none; }
.form-group input[readonly] { background:#f3f2f1; color:#605e5c; cursor:not-allowed; }

.btn-primary { background:#0078d4; color:#fff; border:none; padding:10px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; transition:all 0.2s ease; }
.btn-primary:hover { background:#005a9e; }
.btn-secondary { background:#e1e1e1; color:#333; border:none; padding:10px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; transition:all 0.2s ease; }
.btn-secondary:hover { background:#ccc; }

.form-topbar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 18px;
  padding-bottom: 12px;
  border-bottom: 1px solid #eef0f2;
  gap: 12px;
}
.form-topbar-spacer { flex: 1 1 auto; }
.form-topbar-actions {
  display: inline-flex;
  gap: 10px;
  flex: 0 0 auto;
}
.form-topbar .btn-primary,
.form-topbar .btn-secondary {
  min-height: 38px;
  padding: 0 20px;
}

.form-group label .req {
  color: #d13438;
  font-weight: 700;
  margin-left: 2px;
}
.form-group.has-error input,
.form-group.has-error .manager-lookup input,
.form-group.has-error .selected-manager {
  border-color: #d13438 !important;
  box-shadow: 0 0 0 2px rgba(209, 52, 56, 0.18) !important;
}
.form-group .field-error {
  color: #d13438;
  font-size: 12px;
  font-weight: 600;
  margin-top: 4px;
  line-height: 1.3;
}

.table-actions-project {
  display: flex;
  flex-direction: column;
  gap: 12px;
  margin-bottom: 16px;
  padding: 15px;
  background: #f8f9fa;
  border: 1px solid #e9ecef;
  border-radius: 8px;
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
}

.table-actions-row-project {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 15px;
  align-items: end;
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
}

.tableColp4-project {
  display: flex;
  flex-direction: column;
  gap: 8px;
  min-width: 0;
}

.tableColp4-project label {
  font-size: 12px;
  font-weight: 600;
  color: #495057;
  line-height: 1.2;
}

.search-input-project {
  width: 100%;
  height: 40px;
  box-sizing: border-box;
  padding: 0 12px;
  border: 1px solid #ced4da;
  border-radius: 6px;
  font-size: 13px;
  color: #212529;
  background: #fff;
  transition: border-color 0.2s ease, box-shadow 0.2s ease;
}
.search-input-project::placeholder { color: #8a94a0; }
.search-input-project:hover { border-color: #adb5bd; }
.search-input-project:focus {
  border-color: #0078d4;
  box-shadow: 0 0 0 2px rgba(0, 120, 212, 0.2);
  outline: none;
}

.reset-search-btn-project {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  height: 40px;
  min-height: 40px;
  box-sizing: border-box;
  padding: 0 18px;
  border: none;
  border-radius: 6px;
  background-color: #dc3545;
  color: #fff;
  font-size: 13px;
  font-weight: 500;
  line-height: 1;
  cursor: pointer;
  white-space: nowrap;
  transition: background-color 0.2s ease, transform 0.15s ease;
}
.reset-search-btn-project:hover { background-color: #c82333; transform: translateY(-1px); }
.reset-search-btn-project:active { transform: translateY(0); }
.reset-search-btn-project:focus-visible { outline: 2px solid #0078d4; outline-offset: 2px; }

.pagination-controls {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-top: 15px;
  padding-top: 15px;
  border-top: 1px solid #eee;
  flex-wrap: wrap;
  gap: 10px 14px;
  width: 100%;
  box-sizing: border-box;
}

.pagination-info {
  font-size: 14px;
  color: #666;
  white-space: nowrap;
  flex: 1 1 220px;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
}

.pagination-size {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 14px;
  color: #666;
  white-space: nowrap;
  flex: 0 0 auto;
}

.pagination-size select {
  padding: 6px 10px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 14px;
  background: white;
  cursor: pointer;
}

.pagination-size select:focus { outline: none; border-color: #0078d4; }

.pagination {
  display: flex;
  gap: 5px;
  flex-wrap: wrap;
  align-items: center;
  flex: 1 1 260px;
  justify-content: flex-end;
  min-width: 0;
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
  flex: 0 0 auto;
}

.pagination button:hover { background: #f0f0f0; }
.pagination button.active { background: #0078d4; color: #fff; border-color: #0078d4; }
.pagination button:disabled { opacity: 0.5; cursor: not-allowed; }

.toast { position: fixed; bottom: 20px; right: 20px; padding: 12px 20px; color: #fff; border-radius: 6px; box-shadow: 0 2px 6px rgba(0,0,0,0.2); opacity: 0; transition: opacity 0.3s; z-index: 10000; }
.toast.success { background: #28a745; }
.toast.error { background: #dc3545; }

.project-page .modern-table col.pj-col-actions { width: 90px; }
.project-page .modern-table th.pj-th-actions { text-align: center; }
.project-page .modern-table td.pj-td-actions { text-align: center; padding: 8px; }
.project-page .row-delete-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  border: none;
  border-radius: 6px;
  background: #dc3545;
  color: #fff;
  cursor: pointer;
  transition: background-color 0.2s ease, transform 0.15s ease;
  font-size: 15px;
  line-height: 1;
}
.project-page .row-delete-btn:hover { background: #c82333; transform: translateY(-1px); }
.project-page .row-delete-btn:disabled { opacity: 0.4; cursor: not-allowed; transform: none; }
.project-page .row-delete-btn svg { width: 15px; height: 15px; fill: currentColor; }

.manager-lookup-container { position: relative; width: 100%; }
.manager-lookup { display: flex; gap: 10px; margin-bottom: 5px; align-items: center; }
.manager-lookup input { flex: 1; padding: 10px 12px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px; min-width: 0; }
.manager-lookup input:focus { border-color: #0078d4; outline: none; box-shadow: 0 0 3px rgba(0,120,212,0.4); }
.manager-lookup button { background: #0078d4; color: white; border: none; border-radius: 4px; padding: 10px 15px; cursor: pointer; font-size: 14px; font-weight: 500; white-space: nowrap; flex-shrink: 0; }
.manager-lookup button:hover { background: #005a9e; }

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
.manager-suggestion-item { padding: 10px; cursor: pointer; border-bottom: 1px solid #eee; transition: background 0.2s ease; }
.manager-suggestion-item:hover { background: #f3f2f1; }
.manager-suggestion-item.selected { background: #e1f5fe; }
.manager-suggestion-item:last-child { border-bottom: none; }

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
.selected-manager span { font-weight: 500; color: #333; word-break: break-all; }
.remove-manager { background: none; border: none; color: #d13438; cursor: pointer; font-size: 20px; padding: 0 8px; line-height: 1; transition: color 0.2s ease; flex-shrink: 0; }
.remove-manager:hover { color: #a4262c; }

.project-page { --pj-stick-offset: 0px; }
.project-page .pj-table-scroll { position: relative; overflow: visible; }
.project-page .pj-table-scroll .modern-table {
  border-collapse: separate;
  border-spacing: 0;
}
.project-page .pj-table-scroll .modern-table thead th {
  position: sticky;
  top: var(--pj-stick-offset, 0px);
  z-index: 3;
  background-color: #f3f2f1;
}

/* ============ IMPORT MODAL ============ */
.pimp-modal-overlay {
  position: fixed; inset: 0; background: rgba(0,0,0,0.5);
  display: none; align-items: center; justify-content: center; z-index: 100000;
  padding: 20px; box-sizing: border-box;
}
.pimp-modal-overlay.pimp-open { display: flex; }
.pimp-modal {
  background: #fff; width: 100%; max-width: 860px; max-height: 90vh;
  border-radius: 10px; overflow-y: auto; box-shadow: 0 8px 30px rgba(0,0,0,0.3);
  font-family: 'Segoe UI', Arial, sans-serif;
}
.pimp-modal-header {
  display: flex; justify-content: space-between; align-items: center;
  padding: 16px 20px; border-bottom: 1px solid #e1e1e1; position: sticky; top: 0;
  background: #fff; border-radius: 10px 10px 0 0;
}
.pimp-modal-header h3 { margin: 0; color: #0078d4; font-size: 18px; font-weight: 600; }
.pimp-close { background: none; border: none; font-size: 26px; color: #888; cursor: pointer; line-height: 1; }
.pimp-close:hover { color: #d13438; }
.pimp-modal-body { padding: 20px; }
.pimp-step { margin-bottom: 24px; }
.pimp-step h4 { margin: 0 0 10px; color: #333; font-size: 14px; font-weight: 600; border-left: 4px solid #0078d4; padding-left: 10px; }
.pimp-file-area {
  border: 2px dashed #ccc; border-radius: 8px; padding: 28px; text-align: center;
  cursor: pointer; color: #666; transition: all .2s;
}
.pimp-file-area:hover, .pimp-file-area.pimp-dragover { border-color: #0078d4; background: #f0f8ff; }
.pimp-upload-icon { font-size: 34px; }
.pimp-file-area small { color: #999; }
.pimp-file-info { margin-top: 10px; font-weight: 600; color: #333; }
.pimp-preview { margin-top: 14px; }
.pimp-preview-note { background:#e7f3ff; border-radius:6px; padding:12px 14px; font-size:13px; color:#0056b3; margin-bottom:12px; }
.pimp-table-wrap { max-height: 320px; overflow: auto; border: 1px solid #ddd; border-radius: 6px; }
.pimp-table { width: 100%; border-collapse: collapse; font-size: 12px; min-width: 640px; }
.pimp-table th { background: #0078d4; color: #fff; padding: 8px; text-align: left; position: sticky; top: 0; }
.pimp-table td { padding: 7px 8px; border-bottom: 1px solid #eee; }
.pimp-table tr:nth-child(even) { background: #fafafa; }
.pimp-badge { display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; }
.pimp-badge-add { background:#d4edda; color:#155724; }
.pimp-badge-err { background:#f8d7da; color:#721c24; }
.pimp-actions { display:flex; justify-content:flex-end; gap:10px; margin-top:14px; flex-wrap:wrap; }
.pimp-status { margin-top: 12px; font-size: 13px; white-space: pre-line; }
.pimp-status.error { color:#d13438; }
.pimp-status.success { color:#107c10; }
.pimp-btn { border:none; border-radius:6px; padding:10px 20px; cursor:pointer; font-size:14px; font-weight:600; }
.pimp-btn-primary { background:#0078d4; color:#fff; }
.pimp-btn-primary:hover { background:#005a9e; }
.pimp-btn-secondary { background:#e1e1e1; color:#333; }
.pimp-btn-secondary:hover { background:#ccc; }
.pimp-btn-primary:disabled { opacity:.5; cursor:not-allowed; }

@media (max-width:768px){
  .form-group { flex:1 1 100%; }
  .manager-lookup { flex-direction: column; }
  .manager-lookup button { width: 100%; }
  .modern-table { font-size: 12px; }
  .modern-table th, .modern-table td { padding: 6px; }
  .pagination-controls { gap: 8px; }
  .pagination button { padding: 5px 9px; font-size: 13px; }
  .project-page .page-header .btn-primary { padding: 9px 12px; font-size: 12.5px; }
  .table-actions-project { gap: 12px; padding: 10px; }
  .project-page .modern-table { min-width: 640px; }
  .project-page .pj-table-scroll { overflow-x: auto; }
  .project-page .pj-table-scroll .modern-table thead th { position: static; }
  .form-topbar { flex-direction: column; align-items: stretch; }
  .form-topbar-actions { justify-content: flex-end; }
}

@media (max-width:480px){
  .project-page .page-header .btn-primary { padding: 8px 10px; font-size: 12px; }
  .project-page .modern-table { min-width: 560px; }
  .pagination-size label { display: none; }
  .pagination button { padding: 5px 7px; }
}

.project-page {
  --lilac-primary: #605EEE;
  --lilac-primary-dark: #4C49E0;
  --lilac-primary-soft: #EEEEFE;
  --lilac-ink: #1F2430;
  --lilac-muted: #667085;
  --lilac-line: #E4E7EC;
  --lilac-pill: 999px;
  --lilac-shadow-sm: 0 1px 2px rgba(16, 24, 40, .06);
}
.admin-only-btn[hidden] {
  display: none !important;
}
.project-page .pagination-controls {
  display: flex !important;
  flex-direction: row !important;
  flex-wrap: wrap !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 12px !important;
  width: 100% !important;
  margin: 16px 0 0 0 !important;
  padding: 10px 14px !important;
  background: #fff !important;
  border: 1px solid var(--lilac-line) !important;
  border-top: 1px solid var(--lilac-line) !important;
  border-radius: 12px !important;
  box-shadow: var(--lilac-shadow-sm) !important;
  box-sizing: border-box !important;
}
.project-page .pagination-left,
.project-page .pagination-right {
  display: flex !important;
  align-items: center !important;
  gap: 8px !important;
  flex-wrap: wrap !important;
  min-width: 0 !important;
}
.project-page .pagination-left { gap: 14px !important; }
.project-page .pagination-right { justify-content: flex-end !important; }

.project-page .pagination-size {
  display: inline-flex !important;
  align-items: center !important;
  gap: 8px !important;
  flex: 0 0 auto !important;
  white-space: nowrap !important;
}
.project-page .pagination-size label {
  display: inline-block !important;
  font-size: 12.5px !important;
  font-weight: 500 !important;
  color: #344054 !important;
  margin: 0 !important;
}
.project-page .pagination-size select {
  height: 32px !important;
  padding: 0 8px !important;
  border: 1px solid var(--lilac-line) !important;
  border-radius: 8px !important;
  background: #fff !important;
  color: var(--lilac-ink) !important;
  font-size: 12.5px !important;
  font-weight: 600 !important;
  cursor: pointer !important;
}
.project-page .pagination-size select:focus {
  outline: none !important;
  border-color: var(--lilac-primary) !important;
  box-shadow: 0 0 0 3px rgba(96, 94, 238, .22) !important;
}

.project-page .pagination-info {
  display: inline-flex !important;
  align-items: center !important;
  gap: 6px !important;
  flex: 0 1 auto !important;
  font-size: 12.5px !important;
  font-weight: 400 !important;
  color: var(--lilac-muted) !important;
  white-space: nowrap !important;
  overflow: visible !important;
  text-overflow: clip !important;
}
.project-page .pagination-info .pj-pg-range {
  background: #EEF2F6 !important;
  color: #344054 !important;
  font-weight: 700 !important;
  border-radius: 999px !important;
  padding: 2px 10px !important;
}
.project-page .pagination-info .pj-pg-count {
  background: var(--lilac-primary-soft) !important;
  color: var(--lilac-primary-dark) !important;
  font-weight: 700 !important;
  border-radius: 999px !important;
  padding: 2px 10px !important;
}

.project-page .pj-page-btn {
  height: 32px !important;
  min-width: 32px !important;
  width: 32px !important;
  padding: 0 !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  background: #fff !important;
  border: 1px solid var(--lilac-line) !important;
  border-radius: 8px !important;
  color: #475467 !important;
  font-size: 14px !important;
  font-weight: 600 !important;
  font-family: Georgia, "Times New Roman", serif !important;
  line-height: 1 !important;
  cursor: pointer !important;
  flex: 0 0 auto !important;
  transition: background .15s ease, border-color .15s ease, color .15s ease;
}
.project-page .pj-page-btn:hover:not(:disabled) {
  background: var(--lilac-primary-soft) !important;
  border-color: var(--lilac-primary) !important;
  color: var(--lilac-primary-dark) !important;
}
.project-page .pj-page-btn:disabled {
  color: #C4C9D2 !important;
  cursor: not-allowed !important;
  opacity: 1 !important;
}
.project-page .pj-page-btn {
  width: auto !important;
  min-width: 26px !important;
  height: 32px !important;
  min-height: 32px !important;
  padding: 4px 5px !important;
  background: transparent !important;
  border: 0 !important;
  border-radius: 4px !important;
  color: #475467 !important;
  font-family: Georgia, "Times New Roman", serif !important;
  font-size: 16px !important;
  line-height: 1 !important;
  box-shadow: none !important;
  transform: none !important;
}
.project-page .pj-page-btn:hover:not(:disabled) {
  background: transparent !important;
  border: 0 !important;
  color: #344054 !important;
  box-shadow: none !important;
  transform: none !important;
}
.project-page .pj-page-btn:disabled {
  background: transparent !important;
  border: 0 !important;
  color: #98a2b3 !important;
  box-shadow: none !important;
  transform: none !important;
}

.project-page .pagination {
  display: flex !important;
  gap: 4px !important;
  flex: 0 0 auto !important;
  flex-wrap: wrap !important;
  align-items: center !important;
  justify-content: flex-start !important;
  margin: 0 !important;
  min-width: 0 !important;
}
.project-page .pagination .pj-page-number {
  min-width: 32px !important;
  height: 32px !important;
  padding: 0 8px !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  background: #fff !important;
  border: 1px solid var(--lilac-line) !important;
  border-radius: 8px !important;
  color: #475467 !important;
  font-size: 12.5px !important;
  font-weight: 600 !important;
  line-height: 1 !important;
  cursor: pointer !important;
  box-sizing: border-box !important;
  transition: background .15s ease, border-color .15s ease, color .15s ease;
}
.project-page .pagination .pj-page-number:hover:not(.active):not(:disabled) {
  background: var(--lilac-primary-soft) !important;
  border-color: var(--lilac-primary) !important;
  color: var(--lilac-primary-dark) !important;
}
.project-page .pagination .pj-page-number.active {
  background: var(--lilac-primary) !important;
  border-color: var(--lilac-primary) !important;
  color: #fff !important;
  box-shadow: none !important;
}
.project-page .pagination .pj-page-number.dots,
.project-page .pagination .pj-page-number:disabled {
  background: transparent !important;
  border-color: transparent !important;
  color: #98A2B3 !important;
  cursor: default !important;
  min-width: 20px !important;
  padding: 0 !important;
}
.project-page .pj-pg-of {
  font-size: 12.5px !important;
  font-weight: 600 !important;
  color: #475467 !important;
  white-space: nowrap !important;
}

.project-page .pj-loading-overlay {
  position: absolute !important;
  inset: 0 !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  background: rgba(255, 255, 255, .9) !important;
  border-radius: 10px !important;
  z-index: 9 !important;
}
.project-page .pj-loader-card {
  width: min(520px, 84%) !important;
  padding: 22px 26px !important;
  background: #fff !important;
  border: 1px solid #E6E9EC !important;
  border-radius: 12px !important;
  box-shadow: 0 10px 26px rgba(15, 45, 85, .12) !important;
}
.project-page .pj-page-loader { display: block !important; width: 100% !important; }
.project-page .pj-page-loader__track {
  position: relative !important;
  width: 100% !important;
  height: 6px !important;
  overflow: hidden !important;
  background: #DCDCDC !important;
  border-radius: 999px !important;
}
.project-page .pj-page-loader__bar {
  position: absolute !important;
  top: 0 !important;
  left: 0 !important;
  width: 50% !important;
  height: 100% !important;
  background: #6664FF !important;
  border-radius: 999px !important;
  animation: pjPageLoaderSlide 1.25s cubic-bezier(.65, .05, .36, 1) infinite !important;
}
@keyframes pjPageLoaderSlide {
  0%   { transform: translateX(-100%); }
  100% { transform: translateX(200%); }
}
.project-page .table-loader-overlay { display: none !important; }

.project-page .modern-table col.pj-col-pc        { width: 100px; }
.project-page .modern-table col.pj-col-ref       { width: 150px; }
.project-page .modern-table col.pj-col-customer  { width: 180px; }
.project-page .modern-table col.pj-col-project   { width: 240px; }
.project-page .modern-table col.pj-col-person    { width: 155px; }
.project-page .modern-table col.pj-col-person-lg { width: 175px; }

.project-page .modern-table thead th {
  padding: 10px 12px !important;
  font-size: 12px !important;
  font-weight: 700 !important;
  line-height: 1.25 !important;
  color: #323130 !important;
  white-space: nowrap !important;
  vertical-align: middle !important;
  letter-spacing: .01em !important;
  border-bottom: 1px solid #e1e1e1 !important;
  box-shadow: inset 0 -1px 0 #e1e1e1 !important;
}

.project-page .modern-table tbody td {
  padding: 10px 12px !important;
  height: 56px !important;
  vertical-align: middle !important;
  max-width: none !important;
  white-space: normal !important;
  overflow-wrap: anywhere !important;
  border-bottom: 1px solid #eef0f2 !important;
}
.project-page .modern-table tbody td.pj-td-text     { min-width: 100px !important; color: #4b5563 !important; }
.project-page .modern-table tbody td.pj-td-ref      { min-width: 120px !important; }
.project-page .modern-table tbody td.pj-td-project  { min-width: 180px !important; font-weight: 600 !important; color: var(--lilac-ink) !important; }
.project-page .modern-table tbody td.pj-td-person   { min-width: 140px !important; }
.project-page .modern-table tbody tr:hover td { background: #fafbfc !important; }

.project-page .modern-table .manager-cell {
  display: flex !important;
  flex-direction: column !important;
  gap: 2px !important;
  min-width: 0 !important;
  max-width: 100% !important;
  line-height: 1.3 !important;
}
.project-page .modern-table .manager-name {
  font-weight: 600 !important;
  font-size: 13px !important;
  color: #323130 !important;
  overflow-wrap: anywhere !important;
}
.project-page .modern-table .manager-email {
  font-size: 11px !important;
  color: #7a8391 !important;
  word-break: normal !important;
  overflow-wrap: anywhere !important;
  line-height: 1.25 !important;
}
.project-page .modern-table .pj-empty {
  display: inline-block !important;
  color: #b9c0c9 !important;
  font-size: 13px !important;
  line-height: 1.3 !important;
  user-select: none !important;
  cursor: default !important;
}
.project-page .modern-table .project-ref-link { font-weight: 600 !important; }

.project-page .modern-table tbody td:last-child,
.project-page .modern-table thead th:last-child { padding-right: 16px !important; }

@media (max-width: 768px) {
  .project-page .pagination-controls {
    flex-direction: column !important;
    align-items: stretch !important;
    justify-content: flex-start !important;
    gap: 10px !important;
  }
  .project-page .pagination-left,
  .project-page .pagination-right { width: 100% !important; justify-content: flex-start !important; }
  .project-page .modern-table thead th,
  .project-page .modern-table tbody td { padding: 7px 8px !important; }
  .project-page .modern-table tbody td { height: auto !important; }
  .project-page .modern-table .manager-name { font-size: 12px !important; }
  .project-page .modern-table .manager-email { font-size: 10.5px !important; }
  .project-page .modern-table col.pj-col-person    { width: 135px; }
  .project-page .modern-table col.pj-col-person-lg { width: 150px; }
  .project-page .modern-table col.pj-col-project   { width: 200px; }
}

@media (max-width: 480px) {
  .project-page .pagination-controls { padding: 10px 12px !important; }
  .project-page .pagination .pj-page-number,
  .project-page .pj-page-btn { min-width: 28px !important; width: 28px !important; height: 28px !important; font-size: 11.5px !important; }
  .project-page .pagination .pj-page-number { padding: 0 6px !important; }
}

.project-page { padding: 20px !important; box-sizing: border-box !important; }
@media (max-width: 480px) {
  .project-page { padding: 10px !important; }
}

.project-page { --pj-stick-left: 0px; }

.project-page .pj-table-scroll {
  position: relative !important;
  overflow: visible !important;
  max-height: none !important;
  min-width: 2020px !important;
  background: #fff !important;
}

.project-page .pj-table-scroll .modern-table {
  table-layout: fixed !important;
  width: 100% !important;
  min-width: 2020px !important;
  max-width: none !important;
}
.project-page .pj-table-scroll .modern-table col:last-child,
.project-page .pj-table-scroll .modern-table thead th:last-child,
.project-page .pj-table-scroll .modern-table tbody td:last-child {
  width: auto !important;
  min-width: 0 !important;
  max-width: none !important;
}

.project-page .pj-table-scroll .modern-table thead th {
  position: sticky !important;
  top: var(--pj-stick-offset, 0px) !important;
  z-index: 20 !important;
  background-color: #f3f2f1 !important;
}

.project-page .pj-table-scroll .modern-table thead th:nth-child(-n+6),
.project-page .pj-table-scroll .modern-table tbody td:nth-child(-n+6):not([colspan]) {
  position: sticky !important;
  z-index: 5 !important;
  box-sizing: border-box !important;
}
.project-page .pj-table-scroll .modern-table tbody td:nth-child(-n+6):not([colspan]) {
  background: #fff !important;
}
.project-page .pj-table-scroll .modern-table thead th:nth-child(-n+6) {
  z-index: 30 !important;
  background-color: #f3f2f1 !important;
}
.project-page .pj-table-scroll .modern-table tbody tr:hover td:nth-child(-n+6):not([colspan]) {
  background: #fafbfc !important;
}

.project-page .pj-table-scroll .modern-table th:nth-child(1),
.project-page .pj-table-scroll .modern-table td:nth-child(1):not([colspan]) { left: calc(var(--pj-stick-left, 0px) + 0px);     width: 100px !important; min-width: 100px !important; max-width: 100px !important; }
.project-page .pj-table-scroll .modern-table th:nth-child(2),
.project-page .pj-table-scroll .modern-table td:nth-child(2):not([colspan]) { left: calc(var(--pj-stick-left, 0px) + 100px); width: 150px !important; min-width: 150px !important; max-width: 150px !important; }
.project-page .pj-table-scroll .modern-table th:nth-child(3),
.project-page .pj-table-scroll .modern-table td:nth-child(3):not([colspan]) { left: calc(var(--pj-stick-left, 0px) + 250px); width: 180px !important; min-width: 180px !important; max-width: 180px !important; }
.project-page .pj-table-scroll .modern-table th:nth-child(4),
.project-page .pj-table-scroll .modern-table td:nth-child(4):not([colspan]) { left: calc(var(--pj-stick-left, 0px) + 430px); width: 240px !important; min-width: 240px !important; max-width: 240px !important; }
.project-page .pj-table-scroll .modern-table th:nth-child(5),
.project-page .pj-table-scroll .modern-table td:nth-child(5):not([colspan]) { left: calc(var(--pj-stick-left, 0px) + 670px); width: 155px !important; min-width: 155px !important; max-width: 155px !important; }
.project-page .pj-table-scroll .modern-table th:nth-child(6),
.project-page .pj-table-scroll .modern-table td:nth-child(6):not([colspan]) { left: calc(var(--pj-stick-left, 0px) + 825px); width: 155px !important; min-width: 155px !important; max-width: 155px !important; }

.project-page .pj-table-scroll .modern-table th:nth-child(6),
.project-page .pj-table-scroll .modern-table td:nth-child(6):not([colspan]) {
  border-right: 2px solid #E4E7EC !important;
  box-shadow: 2px 0 6px -2px rgba(16, 24, 40, .12) !important;
}

.project-page .pj-loading-overlay { z-index: 40 !important; }

@media (max-width: 768px) {
  .project-page .pj-table-scroll .modern-table tbody td:nth-child(n+3):not([colspan]) {
    position: static !important;
    left: auto !important;
  }
  .project-page .pj-table-scroll .modern-table thead th:nth-child(n+3) { left: auto !important; }
  .project-page .pj-table-scroll .modern-table th:nth-child(2),
  .project-page .pj-table-scroll .modern-table tbody td:nth-child(2):not([colspan]) {
    border-right: 2px solid #E4E7EC !important;
    box-shadow: 2px 0 6px -2px rgba(16, 24, 40, .12) !important;
  }
  .project-page .pj-table-scroll .modern-table th:nth-child(6),
  .project-page .pj-table-scroll .modern-table tbody td:nth-child(6):not([colspan]) {
    border-right: 0 !important;
    box-shadow: none !important;
  }
}
`,
    script: () => {
        setTimeout(async () => {
            // ============================================
            // LOAD SWEETALERT2 (SharePoint-safe CDN loader)
            // ============================================
            const listName = "Project List";
            // ============================================
            // DOM ELEMENTS
            // ============================================
            const tableBody = document.querySelector('.project-table-body');
            const showFormBtn = document.querySelector('.show-form-btn');
            const formSection = document.querySelector('.form-section');
            const tableSection = document.querySelector('.table-section');
            const form = document.querySelector('#projectForm');
            const backBtns = document.querySelectorAll('.back-to-table');
            const formUpdateBtn = document.querySelector('.form-update-btn');
            const exportBtn = document.querySelector('.export-btn');
            const importBtn = document.querySelector('.import-btn');
            const searchInput = document.querySelector('.search-input-project');
            const resetSearchBtn = document.querySelector('.reset-search-btn-project');
            const pageHeaderActions = document.querySelector('.page-header-actions');
            const tableActionsProject = document.querySelector('.table-actions-project');
            const pageSizeSelects = document.querySelectorAll('.page-size-select');
            const paginationWrappers = document.querySelectorAll('.pagination-controls');
            // ============================================
            // TABLE LOADING STATE
            // ============================================
            const showTableLoader = (text) => console.log(`⏳ ${text || 'Loading projects…'}`);
            const hideTableLoader = () => { };
            const createLoadingOverlay = (message = "Processing...") => {
                removeLoadingOverlay();
                const overlay = document.createElement("div");
                overlay.className = "pj-loading-overlay";
                // Use a transparent, non-intrusive background so it doesn't flash white
                overlay.style.cssText = "position:fixed;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,0.15);z-index:99999;";
                overlay.innerHTML = `
          <div class="pj-loader-card" style="background:#fff; padding: 20px 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); text-align: center;">
            <div class="pj-page-loader">
              <div class="pj-page-loader__track"><div class="pj-page-loader__bar"></div></div>
            </div>
            <div style="margin-top:15px; font-size:14px; font-weight:600; color:#333;">${message}</div>
          </div>`;
                return overlay;
            };
            const removeLoadingOverlay = () => {
                document.querySelectorAll(".pj-loading-overlay, .loading-overlay").forEach(el => el.remove());
            };
            const loaderHost = () => {
                const sections = Array.from(document.querySelectorAll('.table-section, .form-section'));
                const visible = sections.find(s => s.offsetParent !== null);
                return visible || document.querySelector('.table-section') || document.body;
            };
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/Kriya";
            const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
            let projectData = [];
            let filteredData = [];
            let editingId = null;
            let currentPage = 1;
            let pageSize = 25;
            let allEmployees = [];
            let allEmployeesLoaded = false;
            // ============================================
            // INIT AUTH
            // ============================================
            const initialLoading = createLoadingOverlay("Loading project data...");
            loaderHost().appendChild(initialLoading);
            try {
                await authManager.initialize({ siteUrl, listName: listName });
            }
            catch (error) {
                console.error('❌ Auth initialization failed:', error);
            }
            removeLoadingOverlay();
            // ============================================
            // MANAGER LOOKUP INSTANCES
            // ============================================
            const groupHeadLookup = new ManagerLookup(document.querySelector('#groupHeadLookup'), siteUrl, ROLE_GROUP_NAMES.GroupHead1);
            const projectManagerLookup = new ManagerLookup(document.querySelector('#projectManagerLookup'), siteUrl, ROLE_GROUP_NAMES.ProjectManager);
            const enggLeadLookup = new ManagerLookup(document.querySelector('#enggLeadLookup'), siteUrl, ROLE_GROUP_NAMES.EnggLead);
            const enggSubLead1Lookup = new ManagerLookup(document.querySelector('#enggSubLead1Lookup'), siteUrl, ROLE_GROUP_NAMES.EnggSubLead1);
            const enggSubLead2Lookup = new ManagerLookup(document.querySelector('#enggSubLead2Lookup'), siteUrl, ROLE_GROUP_NAMES.EnggSubLead2);
            const enggManagersLookup = new ManagerLookup(document.querySelector('#enggManagersLookup'), siteUrl, ROLE_GROUP_NAMES.EngineeringManagers);
            const scmLookup = new ManagerLookup(document.querySelector('#scmLookup'), siteUrl, ROLE_GROUP_NAMES.SCM);
            const controllingLookup = new ManagerLookup(document.querySelector('#controllingLookup'), siteUrl, ROLE_GROUP_NAMES.Controlling);
            const managerLookups = {
                groupHead: groupHeadLookup,
                projectManager: projectManagerLookup,
                enggLead: enggLeadLookup,
                enggSubLead1: enggSubLead1Lookup,
                enggSubLead2: enggSubLead2Lookup,
                enggManagers: enggManagersLookup,
                scm: scmLookup,
                controlling: controllingLookup
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
                    headers: { "Accept": "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose" }
                });
                const data = await res.json();
                return data.d.GetContextWebInformation.FormDigestValue;
            };
            const getEntityType = async () => {
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                        headers: { 'Accept': 'application/json;odata=verbose' }
                    });
                    const data = await res.json();
                    return data.d.ListItemEntityTypeFullName;
                }
                catch (error) {
                    return 'SP.Data.Project_x0020_ListListItem';
                }
            };
            const getManagerDisplayName = (project, field = 'ProjectManager') => {
                const person = project[field];
                if (!person)
                    return { name: '', email: '' };
                if (typeof person === 'string')
                    return { name: person, email: '' };
                if (person.Title || person.Name)
                    return { name: person.Title || person.Name || '', email: person.Email || '' };
                return { name: '', email: '' };
            };
            const clearAllLookups = () => Object.values(managerLookups).forEach(l => l.clear());
            // ============================================
            // VALIDATION HELPERS
            // ============================================
            const setFieldError = (container, message) => {
                if (!container)
                    return;
                container.classList.add('has-error');
                let err = container.querySelector('.field-error');
                if (!err) {
                    err = document.createElement('div');
                    err.className = 'field-error';
                    container.appendChild(err);
                }
                err.textContent = message;
            };
            const clearFieldError = (container) => {
                if (!container)
                    return;
                container.classList.remove('has-error');
                const err = container.querySelector('.field-error');
                if (err)
                    err.remove();
            };
            const clearAllFieldErrors = () => {
                form.querySelectorAll('.form-group').forEach(fg => clearFieldError(fg));
            };
            const validateMandatoryFields = (payload) => {
                let ok = true;
                const textChecks = [
                    { key: 'PC', inputId: '#pc', label: 'Profit Centre' },
                    { key: 'ProjectReference', inputId: '#reference', label: 'Project Reference' },
                    { key: 'CustomerName', inputId: '#customer', label: 'Customer Name' },
                    { key: 'Title', inputId: '#name', label: 'Project Name' }
                ];
                textChecks.forEach(({ key, inputId, label }) => {
                    var _a;
                    const input = document.querySelector(inputId);
                    const fg = input === null || input === void 0 ? void 0 : input.closest('.form-group');
                    const val = ((_a = payload[key]) !== null && _a !== void 0 ? _a : '').toString().trim();
                    if (!val) {
                        ok = false;
                        if (fg)
                            setFieldError(fg, `${label} is required.`);
                    }
                    else if (fg)
                        clearFieldError(fg);
                });
                const lookupChecks = [
                    { key: 'GroupHead1Id', containerSelector: '#groupHeadLookup', label: 'Group Head' },
                    { key: 'ProjectManagerId', containerSelector: '#projectManagerLookup', label: 'Project Manager' }
                ];
                lookupChecks.forEach(({ key, containerSelector, label }) => {
                    var _a;
                    const container = document.querySelector(containerSelector);
                    const fg = container === null || container === void 0 ? void 0 : container.closest('.form-group');
                    const val = ((_a = payload[key]) !== null && _a !== void 0 ? _a : '').toString().trim();
                    if (!val || val === 'null' || val === 'undefined') {
                        ok = false;
                        if (fg)
                            setFieldError(fg, `${label} is required.`);
                    }
                    else if (fg)
                        clearFieldError(fg);
                });
                if (!ok) {
                    Swal.fire({
                        title: 'Missing required fields',
                        html: 'Please complete all fields marked with <span style="color:#d13438;font-weight:700;">*</span> before saving.',
                        icon: 'error',
                        confirmButtonColor: '#0078d4'
                    });
                }
                return ok;
            };
            const findDuplicateProjectReference = (reference, ignoreId = null) => {
                const needle = (reference !== null && reference !== void 0 ? reference : '').toString().trim().toLowerCase();
                if (!needle)
                    return null;
                return projectData.find(p => {
                    var _a;
                    if (ignoreId !== null && p.Id === ignoreId)
                        return false;
                    const ref = ((_a = p.ProjectReference) !== null && _a !== void 0 ? _a : '').toString().trim().toLowerCase();
                    return ref !== '' && ref === needle;
                }) || null;
            };
            // ============================================
            // FETCH MANAGER DETAILS (Fallback)
            // ============================================
            const fetchManagerDetails = async (projects) => {
                for (const project of projects) {
                    const managerFields = [
                        { id: 'GroupHead1Id', field: 'GroupHead1' },
                        { id: 'ProjectManagerId', field: 'ProjectManager' },
                        { id: 'EnggLeadId', field: 'EnggLead' },
                        { id: 'EnggSubLead1Id', field: 'EnggSubLead1' },
                        { id: 'EnggSubLead2Id', field: 'EnggSubLead2' },
                        { id: 'EngineeringManagersId', field: 'EngineeringManagers' },
                        { id: 'SCMId', field: 'SCM' },
                        { id: 'ControllingId', field: 'Controlling' }
                    ];
                    for (const { id, field } of managerFields) {
                        const managerId = project[id];
                        if (managerId && !project[field]) {
                            try {
                                const managerRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${managerId})?$select=Id,Title,Email`, { headers: { 'Accept': 'application/json;odata=verbose' } });
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
            // FETCH PROJECTS
            // ============================================
            const fetchProjects = async (isBackground = false) => {
                if (!isBackground) {
                    showTableLoader("Loading projects…");
                    const loading = createLoadingOverlay("Loading projects…");
                    loaderHost().appendChild(loading);
                }
                try {
                    const selectFields = [
                        'Id', 'PC', 'ProjectReference', 'CustomerName', 'Title',
                        'GroupHead1Id', 'ProjectManagerId', 'EnggLeadId', 'EnggSubLead1Id', 'EnggSubLead2Id',
                        'EngineeringManagersId', 'SCMId', 'ControllingId'
                    ].join(',');
                    const expandFields = [
                        'GroupHead1/Id', 'GroupHead1/Title', 'GroupHead1/Email',
                        'ProjectManager/Id', 'ProjectManager/Title', 'ProjectManager/Email',
                        'EnggLead/Id', 'EnggLead/Title', 'EnggLead/Email',
                        'EnggSubLead1/Id', 'EnggSubLead1/Title', 'EnggSubLead1/Email',
                        'EnggSubLead2/Id', 'EnggSubLead2/Title', 'EnggSubLead2/Email',
                        'EngineeringManagers/Id', 'EngineeringManagers/Title', 'EngineeringManagers/Email',
                        'SCM/Id', 'SCM/Title', 'SCM/Email',
                        'Controlling/Id', 'Controlling/Title', 'Controlling/Email'
                    ].join(',');
                    const query = `$select=${selectFields}&$expand=${expandFields}&$top=5000&$orderby=Id desc`;
                    const res = await fetch(`${apiBase}?${query}`, {
                        headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }
                    });
                    let rawProjects = [];
                    if (!res.ok) {
                        const res2 = await fetch(`${apiBase}?$select=${selectFields}&$top=5000&$orderby=Id desc`, {
                            headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }
                        });
                        if (res2.ok) {
                            const data2 = await res2.json();
                            rawProjects = data2.d.results || [];
                        }
                    }
                    else {
                        const data = await res.json();
                        rawProjects = data.d.results || [];
                    }
                    const hasExpandedData = rawProjects.some(p => p.ProjectManager && p.ProjectManager.Title);
                    if (!hasExpandedData) {
                        await fetchManagerDetails(rawProjects);
                    }
                    projectData = rawProjects;
                    filteredData = getAssignedProjects(projectData);
                    if (!isBackground)
                        currentPage = 1;
                    renderTable(filteredData);
                    // ------------------------------------------------------------
                    // 🔐 Role-based button visibility
                    // ------------------------------------------------------------
                    const isAdmin = hasEditPermission(); // Admin-only helper
                    const canExport = canRead();
                    // Reveal Add + Upload only for Admin using the 'hidden' attribute
                    document.querySelectorAll('.admin-only-btn').forEach(el => {
                        if (isAdmin) {
                            el.removeAttribute('hidden');
                            el.disabled = false;
                        }
                        else {
                            el.setAttribute('hidden', 'true');
                            el.disabled = true;
                        }
                    });
                    // Download Data — visible to every reader
                    const exportBtnEl = document.querySelector('.export-btn');
                    if (exportBtnEl) {
                        exportBtnEl.style.display = canExport ? '' : 'none';
                        exportBtnEl.disabled = !canExport;
                    }
                }
                catch (err) {
                    console.error("Error fetching projects:", err);
                    if (!isBackground)
                        toast("Error fetching projects.", "error");
                }
                finally {
                    if (!isBackground) {
                        hideTableLoader();
                        removeLoadingOverlay();
                    }
                }
            };
            // ============================================
            // SAVE PROJECT
            // ============================================
            const saveProject = async (payload, isUpdate) => {
                if (!validateMandatoryFields(payload))
                    return false;
                // 🔐 ADMIN-ONLY EDIT — block non-admin UPDATE attempts at the lowest
                //    layer too, so a stale form or deep-link cannot bypass the gate.
                if (isUpdate && !hasEditPermission()) {
                    toast("Only an administrator can edit a project", "error");
                    return false;
                }
                const duplicate = findDuplicateProjectReference(payload.ProjectReference, isUpdate ? editingId : null);
                if (duplicate) {
                    const refInput = document.querySelector('#reference');
                    const refFg = refInput === null || refInput === void 0 ? void 0 : refInput.closest('.form-group');
                    if (refFg)
                        setFieldError(refFg, 'This Project Reference already exists.');
                    await Swal.fire({
                        title: 'Duplicate Project Reference',
                        html: `The Project Reference <b>"${payload.ProjectReference}"</b> is already used by another project` +
                            (duplicate.Title ? `:<br/><br/><b>${duplicate.Title}</b>` : '.') +
                            `<br/><br/>Please enter a unique reference.`,
                        icon: 'error', confirmButtonColor: '#0078d4'
                    });
                    return false;
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
                        headers["X-HTTP-Method"] = "MERGE";
                        headers["IF-MATCH"] = "*";
                    }
                    const finalPayload = {
                        __metadata: { type: entityType },
                        PC: payload.PC,
                        ProjectReference: payload.ProjectReference,
                        CustomerName: payload.CustomerName,
                        Title: payload.Title
                    };
                    const managerFields = [
                        { field: 'GroupHead1Id', value: payload.GroupHead1Id },
                        { field: 'ProjectManagerId', value: payload.ProjectManagerId },
                        { field: 'EnggLeadId', value: payload.EnggLeadId },
                        { field: 'EnggSubLead1Id', value: payload.EnggSubLead1Id },
                        { field: 'EnggSubLead2Id', value: payload.EnggSubLead2Id },
                        { field: 'EngineeringManagersId', value: payload.EngineeringManagersId },
                        { field: 'SCMId', value: payload.SCMId },
                        { field: 'ControllingId', value: payload.ControllingId }
                    ];
                    for (const { field, value } of managerFields) {
                        if (value && value !== '' && value !== 'null' && value !== 'undefined')
                            finalPayload[field] = parseInt(value);
                        else
                            finalPayload[field] = null;
                    }
                    const res = await fetch(url, { method: "POST", headers, body: JSON.stringify(finalPayload) });
                    if (!res.ok) {
                        toast(isUpdate ? "Update failed" : "Add failed", "error");
                        return false;
                    }
                    // --- INSTANT LOCAL UPDATE ---
                    const mk = (lookup) => {
                        var _a;
                        return lookup.getSelectedId()
                            ? {
                                Id: parseInt(lookup.getSelectedId()),
                                Title: lookup.getSelectedName().split(' (')[0],
                                Email: ((_a = lookup.getSelectedName().match(/\((.*?)\)/)) === null || _a === void 0 ? void 0 : _a[1]) || ''
                            }
                            : null;
                    };
                    if (isUpdate && editingId !== null) {
                        const idx = projectData.findIndex(p => p.Id === editingId);
                        if (idx !== -1) {
                            projectData[idx] = Object.assign(Object.assign(Object.assign({}, projectData[idx]), finalPayload), { GroupHead1: mk(groupHeadLookup), ProjectManager: mk(projectManagerLookup), EnggLead: mk(enggLeadLookup), EnggSubLead1: mk(enggSubLead1Lookup), EnggSubLead2: mk(enggSubLead2Lookup), EngineeringManagers: mk(enggManagersLookup), SCM: mk(scmLookup), Controlling: mk(controllingLookup) });
                        }
                    }
                    else {
                        const newItem = Object.assign(Object.assign({ Id: Date.now() }, finalPayload), { GroupHead1: mk(groupHeadLookup), ProjectManager: mk(projectManagerLookup), EnggLead: mk(enggLeadLookup), EnggSubLead1: mk(enggSubLead1Lookup), EnggSubLead2: mk(enggSubLead2Lookup), EngineeringManagers: mk(enggManagersLookup), SCM: mk(scmLookup), Controlling: mk(controllingLookup) });
                        projectData.unshift(newItem);
                    }
                    applySearchAndFilter();
                    setTimeout(() => fetchProjects(true), 2000);
                    toast(isUpdate ? "✅ Project updated successfully" : "✅ Project added successfully", "success");
                    return true;
                }
                catch (err) {
                    console.error("❌ Save error:", err);
                    toast(isUpdate ? "Update failed" : "Add failed", "error");
                    return false;
                }
            };
            // ============================================
            // DELETE PROJECT
            // ============================================
            const deleteProject = async (id, projectName) => {
                // 🔐 ADMIN-ONLY: delete is restricted with edit
                if (!hasEditPermission()) {
                    toast("Only an administrator can delete a project", "error");
                    return;
                }
                const confirmed = await Swal.fire({
                    title: 'Delete Project',
                    html: `This will delete all data related to <b>"${projectName}"</b> this project from Kriya. Continue ?.`,
                    icon: 'warning',
                    showCancelButton: true,
                    confirmButtonText: 'Yes, delete',
                    cancelButtonText: 'Cancel',
                    confirmButtonColor: '#dc3545',
                    cancelButtonColor: '#6c757d'
                });
                if (!confirmed.isConfirmed)
                    return;
                const delLoading = createLoadingOverlay(`Deleting "${projectName}"...`);
                document.body.appendChild(delLoading);
                try {
                    const digest = await getDigest();
                    const res = await fetch(`${apiBase}(${id})`, {
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest,
                            "X-HTTP-Method": "DELETE",
                            "IF-MATCH": "*"
                        }
                    });
                    if (res.ok) {
                        projectData = projectData.filter(p => p.Id !== id);
                        applySearchAndFilter();
                        toast(`✅ Project "${projectName}" deleted successfully`, "success");
                        setTimeout(() => fetchProjects(true), 2000);
                    }
                    else {
                        toast("Delete failed. Please try again.", "error");
                    }
                }
                catch (err) {
                    toast("Delete request failed. Please check your connection and try again.", "error");
                }
                finally {
                    removeLoadingOverlay();
                }
            };
            if (editingId !== null && !hasEditPermission()) {
                toast("Only an administrator can edit a project", "error");
                return;
            }
            // ============================================
            // RENDER TABLE
            // ============================================
            const renderTable = (data) => {
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, data.length);
                const paginatedData = data.slice(startIndex, endIndex);
                if (paginatedData.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="13" style="text-align:center;padding:40px;color:#999;">No projects found</td></tr>`;
                    updatePaginationUI(0);
                    return;
                }
                const trashIcon = `<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M9 3v1H4v2h1v13a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V6h1V4h-5V3H9zm0 5h2v11H9V8zm4 0h2v11h-2V8z"/></svg>`;
                tableBody.innerHTML = paginatedData.map(p => {
                    const gh = getManagerDisplayName(p, 'GroupHead1');
                    const pm = getManagerDisplayName(p, 'ProjectManager');
                    const enggLead = getManagerDisplayName(p, 'EnggLead');
                    const enggSubLead1 = getManagerDisplayName(p, 'EnggSubLead1');
                    const enggSubLead2 = getManagerDisplayName(p, 'EnggSubLead2');
                    const enggManagers = getManagerDisplayName(p, 'EngineeringManagers');
                    const scm = getManagerDisplayName(p, 'SCM');
                    const controlling = getManagerDisplayName(p, 'Controlling');
                    const canEdit = hasEditPermission(p);
                    const renderManagerCell = (manager) => {
                        if (!manager.name && !manager.email)
                            return '<span class="pj-empty" title="Not assigned">—</span>';
                        return `<div class="manager-cell">
                <span class="manager-name">${manager.name}</span>
                ${manager.email ? `<span class="manager-email">${manager.email}</span>` : ''}
              </div>`;
                    };
                    return `
            <tr data-project-id="${p.Id}">
              <td class="pj-td pj-td-text">${p.PC || '<span class="pj-empty">—</span>'}</td>
              <td class="pj-td pj-td-ref">
                <a href="#"
                   class="project-ref-link${canEdit ? '' : ' project-ref-link-disabled'}"
                   data-id="${p.Id}"
                   title="${canEdit ? 'Open edit form' : 'You do not have permission to edit this project'}"
                   ${canEdit ? '' : 'aria-disabled="true"'}>${p.ProjectReference || '<span class="pj-empty">—</span>'}</a>
              </td>
              <td class="pj-td pj-td-text">${p.CustomerName || '<span class="pj-empty">—</span>'}</td>
              <td class="pj-td pj-td-project">${p.Title || '<span class="pj-empty">—</span>'}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(gh)}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(pm)}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(enggLead)}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(enggSubLead1)}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(enggSubLead2)}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(enggManagers)}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(scm)}</td>
              <td class="pj-td pj-td-person">${renderManagerCell(controlling)}</td>
              <td class="pj-td pj-td-actions">
  ${canEdit
                        ? `<button type="button" class="row-delete-btn" data-id="${p.Id}" data-name="${(p.Title || '').replace(/"/g, '&quot;')}" title="Delete project">${trashIcon}</button>`
                        : ''}
</td>
            </tr>
          `;
                }).join("");
                attachTableHandlers();
                updatePaginationUI(data.length);
            };
            const updatePaginationUI = (totalRecords) => {
                const totalPages = Math.max(1, Math.ceil(totalRecords / pageSize));
                const startIndex = totalRecords === 0 ? 0 : (currentPage - 1) * pageSize + 1;
                const endIndex = Math.min(currentPage * pageSize, totalRecords);
                paginationWrappers.forEach(wrapper => {
                    wrapper.querySelector('.record-start').textContent = String(startIndex);
                    wrapper.querySelector('.record-end').textContent = String(endIndex);
                    wrapper.querySelector('.record-count').textContent = String(totalRecords);
                    wrapper.querySelector('.total-pages').textContent = String(totalPages);
                    const firstPageBtn = wrapper.querySelector('.first-page');
                    const prevPageBtn = wrapper.querySelector('.prev-page');
                    const nextPageBtn = wrapper.querySelector('.next-page');
                    const lastPageBtn = wrapper.querySelector('.last-page');
                    const noData = filteredData.length === 0 || totalPages <= 0;
                    if (firstPageBtn)
                        firstPageBtn.disabled = noData || currentPage === 1;
                    if (prevPageBtn)
                        prevPageBtn.disabled = noData || currentPage === 1;
                    if (nextPageBtn)
                        nextPageBtn.disabled = noData || currentPage >= totalPages;
                    if (lastPageBtn)
                        lastPageBtn.disabled = noData || currentPage >= totalPages;
                    const paginationDiv = wrapper.querySelector('.page-numbers');
                    paginationDiv.innerHTML = '';
                    if (noData)
                        return;
                    const addTile = (label, page, active) => {
                        const btn = document.createElement('button');
                        btn.type = 'button';
                        btn.className = `pj-page-number${active ? ' active' : ''}`;
                        btn.textContent = label;
                        if (active)
                            btn.setAttribute('aria-current', 'page');
                        btn.addEventListener('click', () => goToPage(page));
                        paginationDiv.appendChild(btn);
                    };
                    const addDots = () => {
                        const dots = document.createElement('button');
                        dots.type = 'button';
                        dots.className = 'pj-page-number dots';
                        dots.textContent = '…';
                        dots.disabled = true;
                        paginationDiv.appendChild(dots);
                    };
                    const maxVisiblePages = 5;
                    let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
                    const endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
                    if (endPage - startPage + 1 < maxVisiblePages)
                        startPage = Math.max(1, endPage - maxVisiblePages + 1);
                    if (startPage > 1)
                        addDots();
                    for (let i = startPage; i <= endPage; i++)
                        addTile(i.toString(), i, i === currentPage);
                    if (endPage < totalPages)
                        addDots();
                });
            };
            const goToPage = (page) => {
                const totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
                if (page < 1 || page > totalPages || page === currentPage)
                    return;
                currentPage = page;
                renderTable(filteredData);
            };
            paginationWrappers.forEach(wrapper => {
                var _a, _b, _c, _d;
                (_a = wrapper.querySelector('.first-page')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => goToPage(1));
                (_b = wrapper.querySelector('.prev-page')) === null || _b === void 0 ? void 0 : _b.addEventListener('click', () => goToPage(currentPage - 1));
                (_c = wrapper.querySelector('.next-page')) === null || _c === void 0 ? void 0 : _c.addEventListener('click', () => goToPage(currentPage + 1));
                (_d = wrapper.querySelector('.last-page')) === null || _d === void 0 ? void 0 : _d.addEventListener('click', () => goToPage(Math.max(1, Math.ceil(filteredData.length / pageSize))));
            });
            pageSizeSelects.forEach(select => {
                select.addEventListener('change', (e) => {
                    const firstVisibleIndex = Math.max(0, (currentPage - 1) * pageSize);
                    pageSize = parseInt(e.target.value) || 25;
                    pageSizeSelects.forEach(s => s.value = pageSize.toString());
                    const totalPages = Math.max(1, Math.ceil(filteredData.length / pageSize));
                    currentPage = Math.min(totalPages, Math.floor(firstVisibleIndex / pageSize) + 1);
                    renderTable(filteredData);
                });
            });
            const attachTableHandlers = () => {
                tableBody.removeEventListener('click', handleTableClick);
                tableBody.addEventListener('click', handleTableClick);
            };
            const handleTableClick = (event) => {
                const target = event.target;
                const delBtn = target.closest('.row-delete-btn');
                if (delBtn) {
                    event.preventDefault();
                    event.stopPropagation();
                    const id = Number(delBtn.getAttribute('data-id'));
                    const name = delBtn.getAttribute('data-name') || 'this project';
                    if (!id || isNaN(id))
                        return;
                    deleteProject(id, name);
                    return;
                }
                const refLink = target.closest('.project-ref-link');
                if (refLink) {
                    event.preventDefault();
                    event.stopPropagation();
                    if (refLink.classList.contains('project-ref-link-disabled')) {
                        toast("You don't have permission to edit this project", "error");
                        return;
                    }
                    const id = Number(refLink.getAttribute('data-id'));
                    if (!id || isNaN(id))
                        return;
                    handleEditClickById(id);
                    return;
                }
            };
            // ============================================
            // FORM MODE HELPERS
            // ============================================
            const enterAddMode = () => {
                form.reset();
                clearAllLookups();
                clearAllFieldErrors();
                editingId = null;
                document.querySelector('#reference').readOnly = false;
                formUpdateBtn.textContent = 'Save Project';
                // Hide toolbar and search
                pageHeaderActions.style.display = "none";
                tableActionsProject.style.display = "none";
                formSection.style.display = "block";
                tableSection.style.display = "none";
                showFormBtn.style.display = "none";
            };
            const enterEditMode = (proj) => {
                clearAllFieldErrors();
                document.querySelector('#pc').value = proj.PC || '';
                const refEl = document.querySelector('#reference');
                refEl.value = proj.ProjectReference || '';
                refEl.readOnly = true;
                document.querySelector('#customer').value = proj.CustomerName || '';
                document.querySelector('#name').value = proj.Title || '';
                clearAllLookups();
                const setLookup = (lookup, managerData) => {
                    if (managerData && managerData.Id && managerData.Title) {
                        lookup.setSelected({ id: managerData.Id, title: managerData.Title, email: managerData.Email || '' });
                    }
                };
                setLookup(groupHeadLookup, proj.GroupHead1);
                setLookup(projectManagerLookup, proj.ProjectManager);
                setLookup(enggLeadLookup, proj.EnggLead);
                setLookup(enggSubLead1Lookup, proj.EnggSubLead1);
                setLookup(enggSubLead2Lookup, proj.EnggSubLead2);
                setLookup(enggManagersLookup, proj.EngineeringManagers);
                setLookup(scmLookup, proj.SCM);
                setLookup(controllingLookup, proj.Controlling);
                editingId = proj.Id;
                formUpdateBtn.textContent = 'Update Project';
                // Hide toolbar and search
                pageHeaderActions.style.display = "none";
                tableActionsProject.style.display = "none";
                formSection.style.display = "block";
                tableSection.style.display = "none";
                showFormBtn.style.display = "none";
            };
            const handleEditClickById = (id) => {
                const proj = projectData.find(x => x.Id === id);
                if (!proj) {
                    toast("Project not found", "error");
                    return;
                }
                if (!hasEditPermission(proj)) {
                    toast("You don't have permission to edit this project", "error");
                    return;
                }
                enterEditMode(proj);
            };
            const exitForm = () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                showFormBtn.style.display = "inline-block";
                editingId = null;
                form.reset();
                clearAllLookups();
                clearAllFieldErrors();
                document.querySelector('#reference').readOnly = false;
                formUpdateBtn.textContent = 'Save Project';
                // Restore toolbar and search
                pageHeaderActions.style.display = "flex";
                tableActionsProject.style.display = "flex";
            };
            // ============================================
            // BUILD PAYLOAD FROM FORM
            // ============================================
            const buildPayload = () => ({
                PC: document.querySelector('#pc').value.trim(),
                ProjectReference: document.querySelector('#reference').value.trim(),
                CustomerName: document.querySelector('#customer').value.trim(),
                Title: document.querySelector('#name').value.trim(),
                GroupHead1Id: groupHeadLookup.getSelectedId(),
                ProjectManagerId: projectManagerLookup.getSelectedId(),
                EnggLeadId: enggLeadLookup.getSelectedId(),
                EnggSubLead1Id: enggSubLead1Lookup.getSelectedId(),
                EnggSubLead2Id: enggSubLead2Lookup.getSelectedId(),
                EngineeringManagersId: enggManagersLookup.getSelectedId(),
                SCMId: scmLookup.getSelectedId(),
                ControllingId: controllingLookup.getSelectedId()
            });
            // ============================================
            // SAVE / UPDATE FLOW
            // ============================================
            const performSave = async () => {
                clearAllFieldErrors();
                const payload = buildPayload();
                const isUpdate = editingId !== null;
                if (!canCreate()) {
                    toast("You don't have permission to save projects", "error");
                    return;
                }
                // 🔐 ADMIN-ONLY EDIT — non-admins may not update an existing project
                if (editingId !== null && !hasEditPermission()) {
                    toast("Only an administrator can edit a project", "error");
                    return;
                }
                // ✅ POPUP ALWAYS SHOWS NOW
                const result = await Swal.fire({
                    title: isUpdate ? 'Update Project?' : 'Save Project?',
                    text: 'Do you want to save the changes you made?',
                    icon: 'question',
                    showCancelButton: true,
                    confirmButtonText: 'Yes',
                    cancelButtonText: 'Cancel',
                    confirmButtonColor: '#0078d4',
                    cancelButtonColor: '#6c757d'
                });
                if (!result.isConfirmed)
                    return;
                const saveLoading = createLoadingOverlay(isUpdate ? "Updating project..." : "Saving project...");
                document.body.appendChild(saveLoading);
                let ok = false;
                try {
                    ok = await saveProject(payload, isUpdate);
                }
                finally {
                    removeLoadingOverlay();
                }
                if (ok)
                    exitForm();
            };
            formUpdateBtn.addEventListener('click', async (e) => {
                e.preventDefault();
                await performSave();
            });
            form.addEventListener('submit', (e) => { e.preventDefault(); });
            // ============================================
            // BACK TO LIST — dirty-check
            // ============================================
            backBtns.forEach(btn => {
                btn.addEventListener('click', async (e) => {
                    e.preventDefault();
                    // For back, we just ask if they want to discard or cancel, no need to force popup if empty
                    const result = await Swal.fire({
                        title: 'Go Back?',
                        text: 'Any unsaved changes will be lost. Are you sure?',
                        icon: 'warning',
                        showCancelButton: true,
                        confirmButtonText: 'Yes, go back',
                        cancelButtonText: 'Cancel',
                        confirmButtonColor: '#0078d4',
                        cancelButtonColor: '#6c757d'
                    });
                    if (result.isConfirmed) {
                        exitForm();
                    }
                });
            });
            // ============================================
            // SHOW FORM (ADD)
            // ============================================
            showFormBtn.addEventListener('click', () => {
                // 🔐 ADMIN-ONLY: creating projects is restricted
                if (!hasEditPermission()) {
                    toast("Only an administrator can add projects", "error");
                    return;
                }
                if (!canCreate()) {
                    toast("You don't have permission to create projects", "error");
                    return;
                }
                enterAddMode();
            });
            // ============================================
            // Clear inline errors as user types
            // ============================================
            ['#pc', '#reference', '#customer', '#name'].forEach(sel => {
                const el = document.querySelector(sel);
                if (!el)
                    return;
                el.addEventListener('input', () => {
                    const fg = el.closest('.form-group');
                    if (fg)
                        clearFieldError(fg);
                });
            });
            // ============================================
            // SEARCH / FILTER
            // ============================================
            const applySearchAndFilter = () => {
                const query = searchInput.value.toLowerCase();
                const searchedData = projectData.filter(p => {
                    const gh = getManagerDisplayName(p, 'GroupHead1');
                    const pm = getManagerDisplayName(p, 'ProjectManager');
                    const el = getManagerDisplayName(p, 'EnggLead');
                    const es1 = getManagerDisplayName(p, 'EnggSubLead1');
                    const es2 = getManagerDisplayName(p, 'EnggSubLead2');
                    const em = getManagerDisplayName(p, 'EngineeringManagers');
                    const scm = getManagerDisplayName(p, 'SCM');
                    const ctrl = getManagerDisplayName(p, 'Controlling');
                    return ((p.PC || '').toLowerCase().includes(query) ||
                        (p.ProjectReference || '').toLowerCase().includes(query) ||
                        (p.CustomerName || '').toLowerCase().includes(query) ||
                        (p.Title || '').toLowerCase().includes(query) ||
                        [gh, pm, el, es1, es2, em, scm, ctrl].some(m => (m.name || '').toLowerCase().includes(query) || (m.email || '').toLowerCase().includes(query)));
                });
                filteredData = getAssignedProjects(searchedData);
                currentPage = 1;
                renderTable(filteredData);
            };
            searchInput.addEventListener('input', applySearchAndFilter);
            resetSearchBtn.addEventListener('click', () => { searchInput.value = ''; applySearchAndFilter(); });
            // ============================================
            // EXCEL EXPORT
            // ============================================
            // ============================================
            // EXCEL EXPORT — scoped to visible projects only
            // ============================================
            exportBtn.addEventListener('click', () => {
                if (!canRead()) {
                    toast("You don't have permission to export", "error");
                    return;
                }
                // 🔐 Export only what this user is allowed to see.
                //    Admin (canRead all) → every project.
                //    Everyone else        → getAssignedProjects(projectData) only.
                const visibleData = getAssignedProjects(projectData);
                if (!visibleData.length) {
                    toast("No data to export", "error");
                    return;
                }
                const wsData = visibleData.map(p => {
                    const gh = getManagerDisplayName(p, 'GroupHead1');
                    const pm = getManagerDisplayName(p, 'ProjectManager');
                    const el = getManagerDisplayName(p, 'EnggLead');
                    const es1 = getManagerDisplayName(p, 'EnggSubLead1');
                    const es2 = getManagerDisplayName(p, 'EnggSubLead2');
                    const em = getManagerDisplayName(p, 'EngineeringManagers');
                    const scm = getManagerDisplayName(p, 'SCM');
                    const ctrl = getManagerDisplayName(p, 'Controlling');
                    return {
                        'Profit Centre': p.PC || '',
                        'Project Reference': p.ProjectReference || '',
                        'Customer Name': p.CustomerName || '',
                        'Project Name': p.Title || '',
                        'Group Head': gh.name + (gh.email ? ` (${gh.email})` : ''),
                        'Project Manager': pm.name + (pm.email ? ` (${pm.email})` : ''),
                        'Engg Lead': el.name + (el.email ? ` (${el.email})` : ''),
                        'Engg Sub Lead 1': es1.name + (es1.email ? ` (${es1.email})` : ''),
                        'Engg Sub Lead 2': es2.name + (es2.email ? ` (${es2.email})` : ''),
                        'Engineering Managers': em.name + (em.email ? ` (${em.email})` : ''),
                        'SCM': scm.name + (scm.email ? ` (${scm.email})` : ''),
                        'Controlling': ctrl.name + (ctrl.email ? ` (${ctrl.email})` : '')
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
            // EXCEL IMPORT & TEMPLATE
            // ============================================
            const IMPORT_HEADERS = [
                'Profit Centre', 'Project Reference', 'Customer Name', 'Project Name',
                'Group Head', 'Project Manager', 'Engg Lead', 'Engg Sub Lead 1',
                'Engg Sub Lead 2', 'Engineering Managers', 'SCM', 'Controlling'
            ];
            const ROLE_COLS = [
                { header: 'Group Head', field: 'GroupHead1Id' },
                { header: 'Project Manager', field: 'ProjectManagerId' },
                { header: 'Engg Lead', field: 'EnggLeadId' },
                { header: 'Engg Sub Lead 1', field: 'EnggSubLead1Id' },
                { header: 'Engg Sub Lead 2', field: 'EnggSubLead2Id' },
                { header: 'Engineering Managers', field: 'EngineeringManagersId' },
                { header: 'SCM', field: 'SCMId' },
                { header: 'Controlling', field: 'ControllingId' }
            ];
            // ============================================================
            // IMPORT HELPERS — header validation + row preview
            // ============================================================
            /** Normalize a header cell for comparison (trim, lowercase, collapse spaces). */
            const normHeader = (s) => String(s !== null && s !== void 0 ? s : '').trim().toLowerCase().replace(/\s+/g, ' ');
            /**
             * Compare the uploaded sheet's header row against the template.
             * Missing required columns  → REJECT the file.
             * Extra columns             → accept, but list them.
             * Column order does not matter.
             */
            const validateImportHeaders = (sheetHeaders) => {
                const expected = IMPORT_HEADERS.map(normHeader);
                const got = sheetHeaders.map(normHeader);
                const gotSet = new Set(got);
                const expectedSet = new Set(expected);
                const missing = IMPORT_HEADERS.filter((h, i) => !gotSet.has(expected[i]));
                const unexpected = sheetHeaders.filter((h, i) => {
                    const n = got[i];
                    return n !== '' && !expectedSet.has(n);
                });
                if (missing.length) {
                    return {
                        ok: false, missing, unexpected,
                        message: `This file does not match the Project import template.\n\n` +
                            `Missing required column${missing.length > 1 ? 's' : ''}:\n` +
                            `  • ${missing.join('\n  • ')}\n\n` +
                            `Expected columns (in any order):\n` +
                            `  ${IMPORT_HEADERS.join(', ')}\n\n` +
                            `Download a fresh template with the "Download Template" button and paste your data into it.`
                    };
                }
                return {
                    ok: true, missing: [], unexpected,
                    message: unexpected.length
                        ? `File accepted. These extra columns will be ignored: ${unexpected.join(', ')}`
                        : 'File accepted.'
                };
            };
            /** Show a multi-line error inside the modal (falls back to Swal if missing). */
            const showImportError = (message) => {
                const statusEl = document.querySelector('.pimp-status');
                if (!statusEl) {
                    void Swal.fire({ icon: 'error', title: 'Invalid file', text: message });
                    return;
                }
                statusEl.className = 'pimp-status error';
                statusEl.style.whiteSpace = 'pre-line';
                statusEl.textContent = message;
                statusEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            };
            const showImportSuccess = (message) => {
                const statusEl = document.querySelector('.pimp-status');
                if (!statusEl)
                    return;
                statusEl.className = 'pimp-status success';
                statusEl.style.whiteSpace = 'pre-line';
                statusEl.textContent = message;
            };
            const setImportConfirmEnabled = (enabled) => {
                const btn = document.querySelector('.pimp-btn-primary');
                if (btn)
                    btn.disabled = !enabled;
            };
            /**
             * Blank every role-lookup field on a payload. Called for any row that
             * fails validation, so the preview shows empty role cells on error rows.
             */
            const blankRoleCellsOnError = (payload) => {
                const out = Object.assign({}, payload);
                ROLE_COLS.forEach(rc => { out[rc.field] = ''; });
                return out;
            };
            /** Turn one raw Excel row (keyed by the template headers) into a payload. */
            const mapImportRowToPayload = (raw) => {
                var _a, _b, _c, _d;
                const payload = {
                    PC: String((_a = raw['Profit Centre']) !== null && _a !== void 0 ? _a : '').trim(),
                    ProjectReference: String((_b = raw['Project Reference']) !== null && _b !== void 0 ? _b : '').trim(),
                    CustomerName: String((_c = raw['Customer Name']) !== null && _c !== void 0 ? _c : '').trim(),
                    Title: String((_d = raw['Project Name']) !== null && _d !== void 0 ? _d : '').trim()
                };
                for (const rc of ROLE_COLS) {
                    const emp = matchEmployee(raw[rc.header]);
                    payload[rc.field] = emp ? emp.id : '';
                }
                return payload;
            };
            /** Per-row validation. Returns a list of human-readable problems. */
            const validateImportRow = (payload, existingRefsLower) => {
                const missing = [];
                if (!payload.PC)
                    missing.push('Profit Centre');
                if (!payload.ProjectReference)
                    missing.push('Project Reference');
                if (!payload.CustomerName)
                    missing.push('Customer Name');
                if (!payload.Title)
                    missing.push('Project Name');
                if (!payload.GroupHead1Id)
                    missing.push('Group Head');
                if (!payload.ProjectManagerId)
                    missing.push('Project Manager');
                if (missing.length)
                    return [`Missing: ${missing.join(', ')}`];
                const ref = String(payload.ProjectReference).trim().toLowerCase();
                if (existingRefsLower.has(ref)) {
                    return [`Project Reference "${payload.ProjectReference}" already exists`];
                }
                return [];
            };
            let currentPreviewRows = [];
            /** Render the preview table + error-download button. */
            const renderImportPreview = (rows) => {
                var _a;
                currentPreviewRows = rows;
                const previewHost = document.querySelector('.pimp-preview');
                if (!previewHost)
                    return;
                if (!rows.length) {
                    previewHost.innerHTML = '';
                    return;
                }
                const esc = (v) => String(v !== null && v !== void 0 ? v : '')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;');
                const errCount = rows.filter(r => r.__status === 'error').length;
                const head = `
          <thead>
            <tr>
              <th>#</th>
              ${IMPORT_HEADERS.map(h => `<th>${h}</th>`).join('')}
              <th>Status</th>
            </tr>
          </thead>`;
                const body = rows.map((row, i) => {
                    const isErr = row.__status === 'error';
                    const badge = isErr
                        ? `<span class="pimp-badge pimp-badge-err" title="${esc(row.__error)}">Error</span>`
                        : `<span class="pimp-badge pimp-badge-add">Add</span>`;
                    const cells = IMPORT_HEADERS.map(header => {
                        const isRoleCol = ROLE_COLS.some(rc => rc.header === header);
                        if (isErr && isRoleCol) {
                            return `<td><span class="pimp-cell-blank"></span></td>`;
                        }
                        let value = '';
                        if (header === 'Profit Centre')
                            value = row.PC;
                        else if (header === 'Project Reference')
                            value = row.ProjectReference;
                        else if (header === 'Customer Name')
                            value = row.CustomerName;
                        else if (header === 'Project Name')
                            value = row.Title;
                        else {
                            const rc = ROLE_COLS.find(r => r.header === header);
                            if (rc && row[rc.field]) {
                                const emp = allEmployees.find(e => e.id === Number(row[rc.field]));
                                value = emp ? emp.title : '';
                            }
                        }
                        if (!String(value !== null && value !== void 0 ? value : '').trim()) {
                            return isErr
                                ? `<td><span class="pimp-cell-blank"></span></td>`
                                : `<td><span class="pj-empty">—</span></td>`;
                        }
                        return `<td>${esc(value)}</td>`;
                    }).join('');
                    return `
            <tr class="${isErr ? 'pimp-row-error' : ''}">
              <td>${i + 1}</td>
              ${cells}
              <td>${badge}</td>
            </tr>`;
                }).join('');
                const errorBtn = errCount
                    ? `<div style="margin:12px 0 4px 0;text-align:right;">
               <button type="button" class="pimp-btn pimp-btn-error download-import-errors-btn">
                 ⬇ Download Error List (${errCount} row${errCount === 1 ? '' : 's'})
               </button>
             </div>`
                    : '';
                previewHost.innerHTML = `
          <div class="pimp-preview-note">
            Review before importing. Errored rows have their role cells cleared and will NOT be written.
          </div>
          ${errorBtn}
          <div class="pimp-table-wrap">
            <table class="pimp-table">${head}<tbody>${body}</tbody></table>
          </div>`;
                (_a = previewHost.querySelector('.download-import-errors-btn')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => downloadImportErrors());
            };
            const downloadImportErrors = () => {
                const errRows = currentPreviewRows.filter(r => r.__status === 'error');
                if (!errRows.length) {
                    toast('No error rows to download', 'error');
                    return;
                }
                const header = [...IMPORT_HEADERS, '__Row #', '__Reason'];
                const aoa = [header];
                errRows.forEach((row, idx) => {
                    const line = [];
                    line.push(row.PC || '');
                    line.push(row.ProjectReference || '');
                    line.push(row.CustomerName || '');
                    line.push(row.Title || '');
                    for (const rc of ROLE_COLS) {
                        // errors always come in with blank role cells — write them blank
                        const emp = allEmployees.find(e => e.id === Number(row[rc.field]));
                        line.push(emp ? emp.title : '');
                    }
                    // row index in the original sheet is (idx in preview) + 2 (header row + 1-based)
                    line.push(idx + 2);
                    line.push(row.__error || '');
                    aoa.push(line);
                });
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                ws['!cols'] = header.map(h => {
                    if (h === '__Row #')
                        return { wch: 8 };
                    if (h === '__Reason')
                        return { wch: 60 };
                    return { wch: 22 };
                });
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Import Errors');
                const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                const stamp = new Date().toISOString().slice(0, 10);
                saveAs(new Blob([buf], { type: 'application/octet-stream' }), `projects_import_errors_${stamp}.xlsx`);
                toast(`⬇ Downloaded ${errRows.length} error row${errRows.length === 1 ? '' : 's'}`, 'success');
            };
            /**
             * Handle an uploaded file:
             *   1. Read the workbook.
             *   2. Validate headers → on mismatch, show error IN THE MODAL and stop.
             *   3. Parse + validate each row → on any error, blank that row's role cells.
             *   4. Render the preview.
             */
            /** Handle an uploaded file: header check → per-row validation → preview. */
            const handleImportFile = async (file) => {
                const statusEl = document.querySelector('.pimp-status');
                if (statusEl) {
                    statusEl.textContent = '';
                    statusEl.className = 'pimp-status';
                }
                setImportConfirmEnabled(false);
                currentPreviewRows = [];
                let wb;
                try {
                    const buf = await file.arrayBuffer();
                    wb = XLSX.read(buf);
                }
                catch (_a) {
                    showImportError('Could not read the file. Make sure it is a valid .xlsx or .xls file.');
                    return;
                }
                const ws = wb.Sheets[wb.SheetNames[0]];
                if (!ws) {
                    showImportError('The workbook has no sheets.');
                    return;
                }
                const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
                const rawHeaders = (aoa[0] || [])
                    .map(h => String(h !== null && h !== void 0 ? h : ''))
                    .filter(h => h.trim() !== '');
                if (!rawHeaders.length) {
                    showImportError('The first row of the sheet is empty — expected the template header row.');
                    return;
                }
                const headerCheck = validateImportHeaders(rawHeaders);
                if (!headerCheck.ok) {
                    showImportError(headerCheck.message);
                    setImportConfirmEnabled(false);
                    return;
                }
                setImportConfirmEnabled(true);
                const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: false });
                if (!rawRows.length) {
                    showImportError('No data rows found below the header. Add at least one project row.');
                    setImportConfirmEnabled(false);
                    return;
                }
                await ensureEmployeesLoaded();
                const existingRefsLower = new Set(projectData
                    .map(p => { var _a; return String((_a = p.ProjectReference) !== null && _a !== void 0 ? _a : '').trim().toLowerCase(); })
                    .filter(Boolean));
                const previewRows = [];
                let addCount = 0, errCount = 0;
                for (const raw of rawRows) {
                    const payload = mapImportRowToPayload(raw);
                    const problems = validateImportRow(payload, existingRefsLower);
                    if (problems.length) {
                        previewRows.push(Object.assign(Object.assign({}, blankRoleCellsOnError(payload)), { __status: 'error', __error: problems.join(' | ') }));
                        errCount++;
                    }
                    else {
                        previewRows.push(Object.assign(Object.assign({}, payload), { __status: 'added', __error: '' }));
                        addCount++;
                        existingRefsLower.add(String(payload.ProjectReference).trim().toLowerCase());
                    }
                }
                renderImportPreview(previewRows);
                if (errCount) {
                    showImportError(`Preview ready — ${addCount} valid, ${errCount} error${errCount === 1 ? '' : 's'}. ` +
                        `Errored rows have their role cells cleared. Use "Download Error List" to save them.`);
                }
                else {
                    showImportSuccess(`Preview ready — all ${addCount} row${addCount === 1 ? '' : 's'} valid. Click Import to save them.`);
                }
            };
            /** Bind the modal's file input + drag-drop to handleImportFile. */
            const bindImportFileInput = () => {
                const fileInput = document.querySelector('.pimp-file-input');
                if (!fileInput) {
                    console.warn('⚠️ .pimp-file-input not found — import modal markup may be missing.');
                    return;
                }
                fileInput.addEventListener('change', () => {
                    var _a;
                    const f = (_a = fileInput.files) === null || _a === void 0 ? void 0 : _a[0];
                    if (f)
                        void handleImportFile(f);
                });
                const dropArea = document.querySelector('.pimp-file-area');
                if (!dropArea)
                    return;
                ['dragenter', 'dragover'].forEach(ev => dropArea.addEventListener(ev, (e) => {
                    e.preventDefault();
                    dropArea.classList.add('pimp-dragover');
                }));
                ['dragleave', 'drop'].forEach(ev => dropArea.addEventListener(ev, (e) => {
                    e.preventDefault();
                    dropArea.classList.remove('pimp-dragover');
                }));
                dropArea.addEventListener('drop', (e) => {
                    var _a;
                    const dt = e.dataTransfer;
                    const f = (_a = dt === null || dt === void 0 ? void 0 : dt.files) === null || _a === void 0 ? void 0 : _a[0];
                    if (f)
                        void handleImportFile(f);
                });
            };
            const isServiceAccount = (title) => !title || title.includes('System Account') || title.includes('App@sharepoint') ||
                title.includes('SharingLinks') || title.includes('NT AUTHORITY') || title.includes('Membership');
            const ensureEmployeesLoaded = async () => {
                var _a;
                if (allEmployeesLoaded)
                    return;
                try {
                    let users = [];
                    const res = await fetch(`${siteUrl}/_api/web/siteusers?$filter=PrincipalType eq 1&$top=5000&$orderby=Title`, {
                        headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }
                    });
                    if (res.ok) {
                        const data = await res.json();
                        users = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    }
                    allEmployees = users
                        .filter((u) => !isServiceAccount(u.Title || ''))
                        .map((u) => ({ id: u.Id, title: u.Title || '', email: u.Email || '' }))
                        .sort((a, b) => a.title.localeCompare(b.title));
                    allEmployeesLoaded = true;
                }
                catch (error) {
                    console.error('Error loading employees:', error);
                }
            };
            const matchEmployee = (value) => {
                if (value === null || value === undefined)
                    return null;
                const v = String(value).trim();
                if (!v)
                    return null;
                const m = v.match(/\(([^)]+)\)/);
                const emailInParens = m ? m[1].trim() : '';
                const searchEmails = [];
                if (emailInParens)
                    searchEmails.push(emailInParens);
                if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v))
                    searchEmails.push(v);
                for (const em of searchEmails) {
                    const f = allEmployees.find(e => e.email && e.email.toLowerCase() === em.toLowerCase());
                    if (f)
                        return f;
                }
                const byName = allEmployees.find(e => e.title && e.title.toLowerCase() === v.toLowerCase());
                if (byName)
                    return byName;
                const low = v.toLowerCase();
                return allEmployees.find(e => (e.title && e.title.toLowerCase().includes(low)) || (e.email && e.email.toLowerCase().includes(low))) || null;
            };
            const writeProjectRow = async (payload) => {
                try {
                    const ref = (payload.ProjectReference || '').trim().toLowerCase();
                    const existing = projectData.find(p => (p.ProjectReference || '').trim().toLowerCase() === ref);
                    if (existing) {
                        return { ok: false, mode: 'error', error: 'Project Reference already exists' };
                    }
                    const digest = await getDigest();
                    const entityType = await getEntityType();
                    const headers = {
                        'Accept': 'application/json;odata=verbose',
                        'Content-Type': 'application/json;odata=verbose',
                        'X-RequestDigest': digest
                    };
                    const finalPayload = {
                        __metadata: { type: entityType },
                        PC: payload.PC || '',
                        ProjectReference: payload.ProjectReference || '',
                        CustomerName: payload.CustomerName || '',
                        Title: payload.Title || ''
                    };
                    for (const rc of ROLE_COLS) {
                        const val = payload[rc.field];
                        if (val && val !== '' && val !== 'null' && val !== 'undefined')
                            finalPayload[rc.field] = parseInt(val, 10);
                    }
                    const res = await fetch(apiBase, { method: 'POST', headers, body: JSON.stringify(finalPayload) });
                    if (!res.ok) {
                        return { ok: false, mode: 'error', error: `HTTP ${res.status}` };
                    }
                    return { ok: true, mode: 'added' };
                }
                catch (err) {
                    return { ok: false, mode: 'error', error: String(err) };
                }
            };
            const downloadTemplate = async () => {
                await ensureEmployeesLoaded();
                if (!allEmployees.length) {
                    toast('Could not load employee list for the template', 'error');
                    return;
                }
                const headers = IMPORT_HEADERS.slice();
                const ws = XLSX.utils.aoa_to_sheet([headers]);
                ws['!cols'] = headers.map((h) => {
                    if (h === 'Project Name' || h === 'Customer Name')
                        return { wch: 30 };
                    if (h === 'Project Reference')
                        return { wch: 22 };
                    if (h === 'Profit Centre')
                        return { wch: 16 };
                    if (h === 'Engineering Managers')
                        return { wch: 36 };
                    if (h.includes('Engg') || h === 'SCM' || h === 'Controlling' || h === 'Group Head' || h === 'Project Manager')
                        return { wch: 34 };
                    return { wch: 18 };
                });
                for (let r = 0; r < 10; r++)
                    XLSX.utils.sheet_add_aoa(ws, [[]], { origin: -1 });
                const empAoa = [['Employee Name', 'Email']];
                allEmployees.forEach(e => empAoa.push([e.title || '', e.email || '']));
                const empWs = XLSX.utils.aoa_to_sheet(empAoa);
                empWs['!cols'] = [{ wch: 40 }, { wch: 34 }];
                const inst = XLSX.utils.aoa_to_sheet([
                    ['How to use this template'],
                    [''],
                    ['1. Fill in the "Projects" sheet. One row per project.'],
                    ['2. Required columns: Profit Centre, Project Reference, Customer Name, Project Name, Group Head, Project Manager.'],
                    ['3. Role columns: copy a person from the "Employee List" sheet (or type name/email).'],
                    ['4. The system will reject rows where the Project Reference already exists in the system.'],
                    ['5. Do NOT rename, reorder, or delete the header row of the "Projects" sheet.'],
                    [''],
                    ['The "Employee List" sheet lists every employee already added in SharePoint.']
                ]);
                inst['!cols'] = [{ wch: 110 }];
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Projects');
                XLSX.utils.book_append_sheet(wb, empWs, 'Employee List');
                XLSX.utils.book_append_sheet(wb, inst, 'Instructions');
                const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                saveAs(new Blob([buf], { type: 'application/octet-stream' }), `Project_Import_Template_${new Date().toISOString().slice(0, 10)}.xlsx`);
                toast('Template downloaded', 'success');
            };
            // ============================================================
            // ============================================================
            // Modal construction + .import-btn wiring
            // ============================================================
            const ensureImportModal = () => {
                var _a, _b, _c;
                let overlay = document.querySelector('.pimp-modal-overlay');
                if (overlay)
                    return overlay;
                overlay = document.createElement('div');
                overlay.className = 'pimp-modal-overlay';
                overlay.setAttribute('role', 'dialog');
                overlay.setAttribute('aria-modal', 'true');
                overlay.innerHTML = `
          <div class="pimp-modal">
            <div class="pimp-modal-header">
              <h3>Upload Multiple Projects</h3>
              <button type="button" class="pimp-close" title="Close">×</button>
            </div>
            <div class="pimp-modal-body">
              <div class="pimp-step">
                <h4>1. Download the template</h4>
                <button type="button" class="pimp-btn pimp-btn-secondary pimp-download-template-btn">
                  ⬇ Download Template
                </button>
              </div>

              <div class="pimp-step">
                <h4>2. Choose your Excel file</h4>
                <div class="pimp-file-area">
                  <div class="pimp-upload-icon">📂</div>
                  <div>Click to choose a file, or drag &amp; drop here</div>
                  <small>Accepted formats: .xlsx, .xls</small>
                  <input type="file" class="pimp-file-input" accept=".xlsx,.xls" style="display:none;" />
                  <div class="pimp-file-info"></div>
                </div>
              </div>

              <div class="pimp-preview"></div>
              <div class="pimp-status"></div>

              <div class="pimp-actions">
                <button type="button" class="pimp-btn pimp-btn-secondary pimp-cancel-btn">Cancel</button>
                <button type="button" class="pimp-btn pimp-btn-primary" disabled>Import</button>
              </div>
            </div>
          </div>`;
                document.body.appendChild(overlay);
                const close = () => { overlay.classList.remove('pimp-open'); };
                (_a = overlay.querySelector('.pimp-close')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', close);
                (_b = overlay.querySelector('.pimp-cancel-btn')) === null || _b === void 0 ? void 0 : _b.addEventListener('click', close);
                overlay.addEventListener('click', (e) => { if (e.target === overlay)
                    close(); });
                const fileInput = overlay.querySelector('.pimp-file-input');
                const fileArea = overlay.querySelector('.pimp-file-area');
                fileArea.addEventListener('click', (e) => {
                    if (e.target === fileInput)
                        return;
                    fileInput.click();
                });
                (_c = overlay.querySelector('.pimp-download-template-btn')) === null || _c === void 0 ? void 0 : _c.addEventListener('click', () => { void downloadTemplate(); });
                const confirmBtn = overlay.querySelector('.pimp-btn-primary');
                confirmBtn.addEventListener('click', async () => {
                    const toImport = currentPreviewRows.filter(r => r.__status === 'added');
                    if (!toImport.length) {
                        toast('No valid rows to import', 'error');
                        return;
                    }
                    confirmBtn.disabled = true;
                    const loading = createLoadingOverlay(`Importing ${toImport.length} project(s)…`);
                    document.body.appendChild(loading);
                    let ok = 0, fail = 0;
                    try {
                        for (const r of toImport) {
                            const payload = {
                                PC: r.PC, ProjectReference: r.ProjectReference,
                                CustomerName: r.CustomerName, Title: r.Title
                            };
                            ROLE_COLS.forEach(rc => { payload[rc.field] = r[rc.field] || ''; });
                            const res = await writeProjectRow(payload);
                            if (res.ok)
                                ok++;
                            else
                                fail++;
                        }
                        await fetchProjects(true);
                        toast(`✅ Imported ${ok} project${ok === 1 ? '' : 's'}` +
                            (fail ? ` — ${fail} failed (see console)` : ''), fail ? 'error' : 'success');
                        close();
                    }
                    finally {
                        removeLoadingOverlay();
                        confirmBtn.disabled = false;
                    }
                });
                return overlay;
            };
            importBtn.addEventListener('click', () => {
                if (!canCreate()) {
                    toast("You don't have permission to import projects", 'error');
                    return;
                }
                const overlay = ensureImportModal();
                overlay.classList.add('pimp-open');
                overlay.querySelector('.pimp-preview').innerHTML = '';
                const statusEl = overlay.querySelector('.pimp-status');
                statusEl.textContent = '';
                statusEl.className = 'pimp-status';
                overlay.querySelector('.pimp-file-info').textContent = '';
                setImportConfirmEnabled(false);
                currentPreviewRows = [];
                const fileInput = overlay.querySelector('.pimp-file-input');
                fileInput.value = '';
                fileInput.onchange = () => {
                    var _a;
                    const f = (_a = fileInput.files) === null || _a === void 0 ? void 0 : _a[0];
                    if (!f)
                        return;
                    overlay.querySelector('.pimp-file-info').textContent = f.name;
                    void handleImportFile(f);
                };
                const dropArea = overlay.querySelector('.pimp-file-area');
                dropArea.ondragover = (e) => { e.preventDefault(); dropArea.classList.add('pimp-dragover'); };
                dropArea.ondragleave = () => dropArea.classList.remove('pimp-dragover');
                dropArea.ondrop = (e) => {
                    var _a, _b;
                    e.preventDefault();
                    dropArea.classList.remove('pimp-dragover');
                    const f = (_b = (_a = e.dataTransfer) === null || _a === void 0 ? void 0 : _a.files) === null || _b === void 0 ? void 0 : _b[0];
                    if (!f)
                        return;
                    overlay.querySelector('.pimp-file-info').textContent = f.name;
                    void handleImportFile(f);
                };
            });
            bindImportFileInput();
            // ============================================
            // INITIALIZE
            // ============================================
            fetchProjects();
        }, 300);
    }
};
//# sourceMappingURL=project.js.map
// project.ts - Complete file with SCM & Controlling fields added
import Swal from 'sweetalert2';
import 'sweetalert2/dist/sweetalert2.min.css';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import "../styles/common.css";
import {
  authManager,
  canCreate,
  // canDelete,
  canRead,
  canEditProject,
  getAssignedProjects,
  updateUIBasedOnPermissions
} from './auth';
declare const _spPageContextInfo: any;

// ============================================
// PERMISSION HELPERS
// ============================================

const hasEditPermission = (project?: any): boolean => {
  return canEditProject(project);
};

// ============================================
// MANAGER LOOKUP COMPONENT
// ============================================

class ManagerLookup {
  private container: HTMLElement;
  private searchInput: HTMLInputElement;
  private searchBtn: HTMLButtonElement;
  private suggestionsDiv: HTMLElement;
  private selectedDisplay: HTMLElement;
  private selectedNameSpan: HTMLElement;
  private removeBtn: HTMLButtonElement;
  private hiddenInput: HTMLInputElement;
  private siteUrl: string;
  private onSelect?: (manager: any) => void;
  private onClear?: () => void;

  constructor(container: HTMLElement, siteUrl: string, options?: {
    onSelect?: (manager: any) => void;
    onClear?: () => void;
  }) {
    this.container = container;
    this.siteUrl = siteUrl;
    this.onSelect = options?.onSelect;
    this.onClear = options?.onClear;

    this.searchInput = container.querySelector('.manager-search-input') as HTMLInputElement;
    this.searchBtn = container.querySelector('.search-manager-btn') as HTMLButtonElement;
    this.suggestionsDiv = container.querySelector('.manager-suggestions') as HTMLElement;
    this.selectedDisplay = container.querySelector('.selected-manager') as HTMLElement;
    this.selectedNameSpan = container.querySelector('.selected-manager-name') as HTMLSpanElement;
    this.removeBtn = container.querySelector('.remove-manager') as HTMLButtonElement;
    this.hiddenInput = container.querySelector('.manager-id-input') as HTMLInputElement;

    this.bindEvents();
  }

  private bindEvents(): void {
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
      } else {
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
      if (!this.container.contains(e.target as Node)) {
        this.suggestionsDiv.style.display = 'none';
      }
    });
  }

  private async search(): Promise<void> {
    const searchTerm = this.searchInput.value.trim();
    if (searchTerm.length < 2) {
      Swal.fire({ title: 'Please enter at least 2 characters', icon: 'warning' });
      return;
    }

    try {
      const managers = await this.searchManagers(searchTerm);
      this.displaySuggestions(managers);
    } catch (error) {
      console.error('Error searching managers:', error);
      Swal.fire({ title: 'Search failed', icon: 'error' });
    }
  }

  private async searchManagers(searchTerm: string): Promise<any[]> {
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
      .filter((user: any) => {
        const title = user.Title || '';
        return !title.includes('System Account') &&
          !title.includes('App@sharepoint') &&
          !title.includes('SharingLinks') &&
          !title.includes('NT AUTHORITY') &&
          !title.includes('Membership');
      })
      .map((user: any) => ({
        id: user.Id,
        title: user.Title || '',
        email: user.Email || ''
      }))
      .sort((a: any, b: any) => a.title.localeCompare(b.title));
  }

  private displaySuggestions(managers: any[]): void {
    this.suggestionsDiv.innerHTML = '';

    if (managers.length === 0) {
      const noResults = document.createElement('div');
      noResults.className = 'manager-suggestion-item';
      noResults.textContent = 'No managers found';
      this.suggestionsDiv.appendChild(noResults);
    } else {
      managers.forEach(manager => {
        const suggestion = document.createElement('div');
        suggestion.className = 'manager-suggestion-item';
        suggestion.innerHTML = `
          <strong>${manager.title}</strong>
          ${manager.email ? `
<small>${manager.email}</small>` : ''}
        `;
        suggestion.dataset.id = manager.id.toString();
        suggestion.addEventListener('click', () => this.select(manager));
        this.suggestionsDiv.appendChild(suggestion);
      });
    }

    this.suggestionsDiv.style.display = 'block';
  }

  public select(manager: any): void {
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

  public clear(): void {
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

  public getSelectedId(): string | null {
    return this.hiddenInput.value || null;
  }

  public getSelectedName(): string {
    return this.selectedNameSpan.textContent || '';
  }

  public setSelected(manager: any): void {
    if (manager && manager.id) {
      this.select(manager);
    } else {
      this.clear();
    }
  }
}

// ============================================
// MAIN PROJECT MANAGEMENT COMPONENT
// ============================================

const projectContent = `
<div class="project-page fade-in">
<div class="page-header">
      <h1 class="left-side">Project Management</h1>
      <div>
        <button class="btn-primary show-form-btn">+ Add New Project</button>
        <button class="btn-primary export-btn">⬇ Download Project </button>
         <button class="btn-primary import-btn">Upload Project</button>
         <button class="btn-primary download-template-btn">⬇ Template</button>
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

    <div class="pj-table-scroll">
      <table class="modern-table">
        <thead>
          <tr>
            <th>Profit Centre</th>
            <th>Project Reference</th>
            <th>Customer Name</th>
            <th>Project Name</th>
            <th>Group Head</th>
            <th>Project Manager</th>
            <th>Engg Lead</th>
            <th>Engg Sub Lead 1</th>
            <th>Engg Sub Lead 2</th>
            <th>Engineering Managers</th>
            <th>SCM</th>
            <th>Controlling</th>
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
        <div class="form-group">
          <label>Group Head</label>
          <div class="manager-lookup-container" id="groupHeadLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search for group head..." />
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
      <div class="form-row">
        <div class="form-group">
          <label>SCM</label>
          <div class="manager-lookup-container" id="scmLookup">
            <div class="manager-lookup">
              <input type="text" class="manager-search-input" placeholder="Search for SCM..." />
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
              <input type="text" class="manager-search-input" placeholder="Search for Controlling..." />
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
  @font-face {
  font-family: 'ABBvoiceOffice';
  src: url('../../assets/ABBvoiceOffice_Rg.ttf') format('truetype');
  font-weight: 600;
  font-style: normal;
  font-display: swap;
}

.project-page { font-family:'ABBvoiceOffice'; }

/* Page header: "Project Management" title on the LEFT side, and the
   action-button row BELOW it, starting from the LEFT.
   Buttons ALWAYS stay on ONE ROW (never wrap): if the space is too
   narrow the row scrolls horizontally instead of wrapping. */
.project-page .page-header {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 10px;
  margin-bottom: 18px;
}
.project-page .page-header h1 { margin: 0; color: #323130; font-size: 22px; font-weight: 600; }
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

.card { background:#fff; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,0.08); padding:20px; margin-bottom:20px; transition:all 0.3s ease; position:relative; /* host for the loading overlay (master.ts pattern) */ }

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

/* 12 columns: keep a readable width and let the wrapper scroll on narrow screens */
.project-page .modern-table { min-width: 900px; }

/* ============================================================ */
/* Clickable Project Reference cell                              */
/* The reference is now a link that opens the edit form (same     */
/* action the old pencil button performed). A disabled look is    */
/* applied when the current user cannot edit that project.        */
/* ============================================================ */
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
.project-page .project-ref-link:focus-visible {
  outline: 2px solid #0078d4;
  outline-offset: 2px;
  border-radius: 2px;
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
.form-actions { display:flex; flex-wrap:wrap; justify-content:center; align-items:center; gap:10px; margin-top:25px; }
.btn-primary { background:#0078d4; color:#fff; border:none; padding:10px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; transition:all 0.2s ease; }
.btn-primary:hover { background:#005a9e; }
.btn-secondary { background:#e1e1e1; color:#333; border:none; padding:10px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; transition:all 0.2s ease; }
.btn-secondary:hover { background:#ccc; }
.form-header h2 { color:#0078d4; margin-bottom:20px; font-size:18px; font-weight:600; }
/* ============================================================ */
/* FILTER BAR - suite structure (same pattern as Procurement /    */
/* Inventory / SEIBIES): namespaced -project classes => no cross-  */
/* module CSS collision; label above the control; one row grid     */
/* with bottom-aligned controls; 40px controls + suite focus ring. */
/* ============================================================ */
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

/* Reset Filter - suite red action button, same 40px height as the search box */
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

/* Pagination Controls - Aligned in single row */
.pagination-controls {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-top: 15px;
  padding-top: 15px;
  border-top: 1px solid #eee;
  flex-wrap: nowrap;         /* info + size + page buttons always on ONE row */
  gap: 10px 14px;
  overflow-x: auto;          /* with 1000+ projects the bar scrolls, it never wraps */
}
.pagination-controls::-webkit-scrollbar { height: 5px; }
.pagination-controls::-webkit-scrollbar-track { background: transparent; }
.pagination-controls::-webkit-scrollbar-thumb { background: #c5c5d2; border-radius: 8px; }

.pagination-info {
  font-size: 14px;
  color: #666;
  white-space: nowrap;
  flex: 1 1 auto;            /* long "Showing 1..25 of 1000" truncates with … first */
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

.pagination-size select:focus {
  outline: none;
  border-color: #0078d4;
}

.pagination {
  display: flex;
  gap: 5px;
  flex-wrap: nowrap;         /* page numbers never drop to a second line */
  align-items: center;
  flex: 0 0 auto;
  max-width: 100%;
  overflow-x: auto;
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

/* ============================================================ */
/* STICKY TABLE HEADER - OUTER (FULL PAGE) SCROLL                 */
/* NO inner wrapper scroll: the page/content scrolls normally     */
/* and the table header row pins itself at the top of the scroll  */
/* area - right under the page/app main header.                   */
/* If your shell has a FIXED top bar that overlaps the top of the */
/* scrolling area, set --pj-stick-offset to that bar's height.    */
/* ============================================================ */
.project-page { --pj-stick-offset: 0px; }
.project-page .pj-table-scroll { position: relative; overflow: visible; }
.project-page .pj-table-scroll .modern-table {
  border-collapse: separate;   /* sticky header cells keep clean borders */
  border-spacing: 0;
}
.project-page .pj-table-scroll .modern-table thead th {
  position: sticky;
  top: var(--pj-stick-offset, 0px);
  z-index: 3;
  background-color: #f3f2f1;   /* opaque so rows scroll underneath it */
}

/* ============================================================ */
/* TABLE LOADING / PROCESSING OVERLAY                            */
/* Veil + spinner card shown while project data is being loaded  */
/* from SharePoint (first load, after save / delete / import).   */
/* ============================================================ */

.project-page .table-loader-overlay {
  position: absolute;
  top: 0; left: 0; right: 0; bottom: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(255, 255, 255, 0.75);
  z-index: 5;
  opacity: 0;
  pointer-events: none;
  transition: opacity 0.2s ease;
}
.project-page .table-loader-overlay.show {
  opacity: 1;
  pointer-events: auto;
}
.project-page .table-loader-box {
  display: flex;
  align-items: center;
  gap: 12px;
  background: #fff;
  border: 1px solid #e9ecef;
  border-radius: 8px;
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.12);
  padding: 12px 20px;
  font-size: 14px;
  font-weight: 500;
  color: #323130;
  white-space: nowrap;
}
.project-spinner {
  width: 22px;
  height: 22px;
  box-sizing: border-box;
  border: 3px solid #e1e6eb;
  border-top-color: #0078d4;
  border-radius: 50%;
  animation: project-spin 0.75s linear infinite;
  flex-shrink: 0;
}
@keyframes project-spin {
  to { transform: rotate(360deg); }
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
.pimp-badge-upd { background:#fff3cd; color:#856404; }
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
  .form-actions .btn-primary,
  .form-actions .btn-secondary { width: 100%; }
  .manager-lookup { flex-direction: column; }
  .manager-lookup button { width: 100%; }
  .modern-table { font-size: 12px; }
  .modern-table th, .modern-table td { padding: 6px; }
  .pagination-controls { gap: 8px; }
  .pagination button { padding: 5px 9px; font-size: 13px; }
  /* Title keeps its LEFT position; button row below it, left aligned.
     Buttons only get slightly more compact so they still fit on ONE row. */
  .project-page .page-header .btn-primary { padding: 9px 12px; font-size: 12.5px; }
  /* Filter bar keeps search + reset on ONE row (label above the search box) */
  .table-actions-project { gap: 12px; padding: 10px; }
  .project-page .modern-table { min-width: 640px; }
  /* Narrow screens only: horizontal pan inside the table wrapper (sticky off there) */
  .project-page .pj-table-scroll { overflow-x: auto; }
  .project-page .pj-table-scroll .modern-table thead th { position: static; }
}

@media (max-width:480px){
  /* Compact the 4 header buttons so they fit on the single row below the title */
  .project-page .page-header .btn-primary { padding: 8px 10px; font-size: 12px; }
  .project-page .modern-table { min-width: 560px; }
  .form-actions { flex-direction: column; }
  /* Pagination keeps ONE row: hide the text label, keep the selector + pages */
  .pagination-size label { display: none; }
  .pagination button { padding: 5px 7px; }
}
`,
  script: () => {
    setTimeout(async () => {
      const listName = "Project List";

      // ============================================
      // DOM ELEMENTS
      // ============================================
      const tableBody = document.querySelector('.project-table-body') as HTMLElement;
      const showFormBtn = document.querySelector('.show-form-btn') as HTMLElement;
      const formSection = document.querySelector('.form-section') as HTMLElement;
      const tableSection = document.querySelector('.table-section') as HTMLElement;
      const form = document.querySelector('#projectForm') as HTMLFormElement;
      const backBtn = document.querySelector('.back-to-table') as HTMLElement;
      const submitBtn = form.querySelector('button[type="submit"]') as HTMLButtonElement;
      const exportBtn = document.querySelector('.export-btn') as HTMLElement;
      const importBtn = document.querySelector('.import-btn') as HTMLElement;
      const templateBtn = document.querySelector('.download-template-btn') as HTMLElement;
      const searchInput = document.querySelector('.search-input-project') as HTMLInputElement;
      const resetSearchBtn = document.querySelector('.reset-search-btn-project') as HTMLButtonElement;
      const formHeader = document.querySelector('.form-header h2') as HTMLElement;
      const paginationDiv = document.querySelector('.pagination') as HTMLElement;
      const pageSizeSelect = document.querySelector('.page-size-select') as HTMLSelectElement;
      const recordCountSpan = document.querySelector('.record-count') as HTMLSpanElement;

      // ============================================
      // TABLE LOADING OVERLAY (spinner while data loads)
      // ============================================
      const tableScrollWrap = document.querySelector('.pj-table-scroll') as HTMLElement;
      const loaderOverlay = document.createElement('div');
      loaderOverlay.className = 'table-loader-overlay';
      loaderOverlay.innerHTML = `
        <div class="table-loader-box">
          <span class="project-spinner"></span>
          <span class="table-loader-text">Loading projects…</span>
        </div>
      `;
      if (tableScrollWrap) tableScrollWrap.appendChild(loaderOverlay);

      const showTableLoader = (text?: string) => {
        const t = loaderOverlay.querySelector('.table-loader-text') as HTMLElement | null;
        if (text && t) t.textContent = text;
        loaderOverlay.classList.add('show');
      };
      const hideTableLoader = () => loaderOverlay.classList.remove('show');

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


      const siteUrl = typeof _spPageContextInfo !== "undefined"
        ? _spPageContextInfo.webAbsoluteUrl
        : window.location.origin + "/sites/DigitalTransformationTest";

      const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      let projectData: any[] = [];
      let filteredData: any[] = [];
      let editingId: number | null = null;
      let currentPage = 1;
      let pageSize = 25;
      let allEmployees: any[] = [];          // cached list of all SharePoint users (for template Employee List + import matching)
      let allEmployeesLoaded = false;

      // ============================================
      // INITIALIZE AUTH
      // ============================================
      const initialLoading = createLoadingOverlay("Loading project data...");
      loaderHost().appendChild(initialLoading);
      try {
        await authManager.initialize({ siteUrl, listName: listName });
        console.log('🔑 Auth initialized successfully');
      } catch (error) {
        console.error('❌ Auth initialization failed:', error);
      }

      // ============================================
      // MANAGER LOOKUP INSTANCES
      // ============================================
      const groupHeadLookup = new ManagerLookup(
        document.querySelector('#groupHeadLookup') as HTMLElement,
        siteUrl
      );
      const projectManagerLookup = new ManagerLookup(
        document.querySelector('#projectManagerLookup') as HTMLElement,
        siteUrl
      );
      const enggLeadLookup = new ManagerLookup(
        document.querySelector('#enggLeadLookup') as HTMLElement,
        siteUrl
      );
      const enggSubLead1Lookup = new ManagerLookup(
        document.querySelector('#enggSubLead1Lookup') as HTMLElement,
        siteUrl
      );
      const enggSubLead2Lookup = new ManagerLookup(
        document.querySelector('#enggSubLead2Lookup') as HTMLElement,
        siteUrl
      );
      const enggManagersLookup = new ManagerLookup(
        document.querySelector('#enggManagersLookup') as HTMLElement,
        siteUrl
      );
      const scmLookup = new ManagerLookup(
        document.querySelector('#scmLookup') as HTMLElement,
        siteUrl
      );
      const controllingLookup = new ManagerLookup(
        document.querySelector('#controllingLookup') as HTMLElement,
        siteUrl
      );

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
      const toast = (msg: string, type: 'success' | 'error' = 'success') => {
        const toastEl = document.createElement("div");
        toastEl.className = `toast ${type}`;
        toastEl.textContent = msg;
        document.body.appendChild(toastEl);
        requestAnimationFrame(() => toastEl.style.opacity = "1");
        setTimeout(() => { toastEl.style.opacity = "0"; setTimeout(() => toastEl.remove(), 500); }, 2500);
      };

      const getDigest = async () => {
        if (typeof _spPageContextInfo !== "undefined") return _spPageContextInfo.formDigestValue;
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

      const getEntityType = async (): Promise<string> => {
        try {
          const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
            headers: {
              'Accept': 'application/json;odata=verbose'
            }
          });
          const data = await res.json();
          return data.d.ListItemEntityTypeFullName;
        } catch (error) {
          console.error('Error getting entity type:', error);
          return 'SP.Data.Project_x0020_ListListItem';
        }
      };

      const getManagerDisplayName = (project: any, field: string = 'ProjectManager'): { name: string, email: string } => {
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
      // FETCH PROJECTS
      // ============================================
      const fetchProjects = async () => {
        showTableLoader("Loading projects…");
        const loading = createLoadingOverlay("Loading projects…");
        loaderHost().appendChild(loading);
        try {
          const selectFields = [
            'Id',
            'PC',
            'ProjectReference',
            'CustomerName',
            'Title',
            'GroupHead1Id',
            'ProjectManagerId',
            'EnggLeadId',
            'EnggSubLead1Id',
            'EnggSubLead2Id',
            'EngineeringManagersId',
            'SCMId',
            'ControllingId'
          ].join(',');

          const expandFields = [
            'GroupHead1',
            'ProjectManager',
            'EnggLead',
            'EnggSubLead1',
            'EnggSubLead2',
            'EngineeringManagers',
            'SCM',
            'Controlling'
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
          } else {
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

        } catch (err) {
          console.error("Error fetching projects:", err);
          toast("Error fetching projects. Please check console for details.", "error");
        } finally {
          hideTableLoader();
          removeLoadingOverlay();
        }
      };

      const fetchManagerDetails = async (projects: any[]) => {
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
                const managerRes = await fetch(
                  `${siteUrl}/_api/web/getuserbyid(${managerId})?$select=Id,Title,Email`,
                  {
                    headers: {
                      'Accept': 'application/json;odata=verbose',
                      'Content-Type': 'application/json;odata=verbose'
                    }
                  }
                );
                if (managerRes.ok) {
                  const managerData = await managerRes.json();
                  project[field] = {
                    Id: managerData.d.Id,
                    Title: managerData.d.Title,
                    Email: managerData.d.Email || ''
                  };
                }
              } catch (error) {
                console.warn(`Could not fetch manager details for ${field} ID ${managerId}:`, error);
              }
            }
          }
        }
      };

      // ============================================
      // SAVE PROJECT
      // ============================================
      const saveProject = async (payload: any, isUpdate = false) => {
        const duplicate = projectData.find(p =>
          p.Title?.toLowerCase() === payload.Title.toLowerCase() &&
          (!isUpdate || p.Id !== editingId)
        );
        if (duplicate) {
          toast("Project with this name already exists", "error");
          return;
        }

        try {
          const digest = await getDigest();
          const entityType = await getEntityType();
          const url = isUpdate ? `${apiBase}(${editingId})` : apiBase;
          
          const headers: any = {
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

          const finalPayload: any = {
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
            if (value && value !== '' && value !== 'null' && value !== 'undefined') {
              finalPayload[field] = parseInt(value);
            } else {
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
          } else {
            toast(isUpdate ? "✅ Project updated successfully" : "✅ Project added successfully", "success");
            editingId = null;
            setTimeout(() => fetchProjects(), 1000);
          }
        } catch (err) {
          console.error("❌ Save error:", err);
          toast(isUpdate ? "Update failed" : "Add failed", "error");
        }
      };

      // ============================================
      // DELETE PROJECT
      // ============================================
      // const deleteProject = async (id: number) => {
      //   console.log('🔍 deleteProject called with ID:', id);
        
      //   if (!canDelete()) {
      //     console.warn('❌ User does not have permission to delete projects');
      //     toast("You don't have permission to delete projects", "error");
      //     return;
      //   }

      //   const project = projectData.find(p => p.Id === id);
      //   if (!project) {
      //     console.error('❌ Project not found with ID:', id);
      //     toast("Project not found", "error");
      //     return;
      //   }
        
      //   const projectName = project.Title || 'this project';
      //   console.log('🗑️ Preparing to delete project:', projectName);

      //   const confirmed = confirm(`⚠️ Delete Project\n\nAre you sure you want to delete "${projectName}"?\n\nThis action cannot be undone!`);
        
      //   if (!confirmed) {
      //     toast("Delete cancelled", "error");
      //     return;
      //   }

      //   console.log('✅ User confirmed deletion');

      //   const deleteLoading = createLoadingOverlay(`Deleting "${projectName}"...`);
      //   loaderHost().appendChild(deleteLoading);
      //   try {
      //     toast(`⏳ Deleting "${projectName}"...`, "success");
          
      //     const digest = await getDigest();
      //     const response = await fetch(`${apiBase}(${id})`, {
      //       method: "POST",
      //       headers: {
      //         "Accept": "application/json;odata=verbose",
      //         "Content-Type": "application/json;odata=verbose",
      //         "X-RequestDigest": digest,
      //         "X-HTTP-Method": "DELETE",
      //         "IF-MATCH": "*"
      //       }
      //     });

      //     console.log('📥 Delete response status:', response.status);

      //     if (response.ok) {
      //       toast(`✅ Project "${projectName}" deleted successfully`, "success");
      //       await fetchProjects();
      //     } else {
      //       let errorMessage = 'Failed to delete project';
      //       try {
      //         const errorData = await response.json();
      //         if (errorData.error && errorData.error.message) {
      //           errorMessage = errorData.error.message;
      //         }
      //       } catch (e) {
      //         errorMessage = response.statusText || errorMessage;
      //       }
            
      //       console.error('❌ Delete error:', errorMessage);
      //       toast(`❌ Delete failed: ${errorMessage}`, "error");
      //     }
      //   } catch (err) {
      //     console.error("❌ Delete error:", err);
      //     toast("Delete request failed. Please check your connection and try again.", "error");
      //   } finally {
      //     removeLoadingOverlay();
      //   }
      // };

      // ============================================
      // RENDER TABLE
      // ============================================
      const renderTable = (data: any[]) => {
        const startIndex = (currentPage - 1) * pageSize;
        const endIndex = Math.min(startIndex + pageSize, data.length);
        const paginatedData = data.slice(startIndex, endIndex);

        if (paginatedData.length === 0) {
          tableBody.innerHTML = `<tr><td colspan="12" style="text-align:center;padding:40px;color:#999;">No projects found</td></tr>`;
          updatePaginationInfo(0);
          renderPagination(0);
          return;
        }

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

          const renderManagerCell = (manager: { name: string, email: string }) => {
            if (!manager.name && !manager.email) return '';
            return `
              <div class="manager-cell">
                <span class="manager-name">${manager.name}</span>
                ${manager.email ? `<span class="manager-email">${manager.email}</span>` : ''}
              </div>
            `;
          };

          return `
            <tr data-project-id="${p.Id}">
              <td>${p.PC || ""}</td>
              <td>
                <a href="#"
                   class="project-ref-link${canEdit ? '' : ' project-ref-link-disabled'}"
                   data-id="${p.Id}"
                   title="${canEdit ? 'Open edit form' : 'You do not have permission to edit this project'}"
                   ${canEdit ? '' : 'aria-disabled="true"'}>${p.ProjectReference || ""}</a>
              </td>
              <td>${p.CustomerName || ""}</td>
              <td>${p.Title || ""}</td>
              <td>${renderManagerCell(gh)}</td>
              <td>${renderManagerCell(pm)}</td>
              <td>${renderManagerCell(enggLead)}</td>
              <td>${renderManagerCell(enggSubLead1)}</td>
              <td>${renderManagerCell(enggSubLead2)}</td>
              <td>${renderManagerCell(enggManagers)}</td>
              <td>${renderManagerCell(scm)}</td>
              <td>${renderManagerCell(controlling)}</td>
            </tr>
          `;
        }).join("");

        attachTableHandlers();
        updatePaginationInfo(data.length);
        const totalPages = Math.ceil(data.length / pageSize);
        renderPagination(totalPages);
      };

      const updatePaginationInfo = (totalRecords: number) => {
        if (totalRecords === 0) {
          recordCountSpan.textContent = 'No records found';
          return;
        }
        const startIndex = (currentPage - 1) * pageSize + 1;
        const endIndex = Math.min(currentPage * pageSize, totalRecords);
        recordCountSpan.textContent = `Showing ${startIndex} to ${endIndex} of ${totalRecords} records`;
      };

      const renderPagination = (totalPages: number) => {
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

      const handleTableClick = (event: Event) => {
        const target = event.target as HTMLElement;

        // Clicking the Project Reference opens the edit form (same path the
        // pencil button used to take). Permission is still enforced inside
        // handleEditClickById; we also short-circuit the disabled look here.
        const refLink = target.closest('.project-ref-link') as HTMLAnchorElement | null;
        if (refLink) {
          event.preventDefault();
          event.stopPropagation();

          if (refLink.classList.contains('project-ref-link-disabled')) {
            toast("You don't have permission to edit this project", "error");
            return;
          }

          const id = Number(refLink.getAttribute('data-id'));
          console.log('✏️ Reference clicked for ID:', id);

          if (!id || isNaN(id)) {
            console.error('Invalid project ID');
            toast("Invalid project ID", "error");
            return;
          }
          handleEditClickById(id);
          return;
        }
      };

      const handleEditClickById = async (id: number) => {
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

        (document.querySelector('#pc') as HTMLInputElement).value = proj.PC || '';
        (document.querySelector('#reference') as HTMLInputElement).value = proj.ProjectReference || '';
        (document.querySelector('#customer') as HTMLInputElement).value = proj.CustomerName || '';
        (document.querySelector('#name') as HTMLInputElement).value = proj.Title || '';

        clearAllLookups();

        const setLookup = (lookup: ManagerLookup, managerData: any) => {
          if (managerData && managerData.Id && managerData.Title) {
            lookup.setSelected({
              id: managerData.Id,
              title: managerData.Title,
              email: managerData.Email || ''
            });
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
          const gh = getManagerDisplayName(p, 'GroupHead1');
          const pm = getManagerDisplayName(p, 'ProjectManager');
          const enggLead = getManagerDisplayName(p, 'EnggLead');
          const enggSubLead1 = getManagerDisplayName(p, 'EnggSubLead1');
          const enggSubLead2 = getManagerDisplayName(p, 'EnggSubLead2');
          const enggManagers = getManagerDisplayName(p, 'EngineeringManagers');
          const scm = getManagerDisplayName(p, 'SCM');
          const controlling = getManagerDisplayName(p, 'Controlling');

          return (
            (p.PC || '').toLowerCase().includes(query) ||
            (p.ProjectReference || '').toLowerCase().includes(query) ||
            (p.CustomerName || '').toLowerCase().includes(query) ||
            (p.Title || '').toLowerCase().includes(query) ||
            (gh.name || '').toLowerCase().includes(query) ||
            (gh.email || '').toLowerCase().includes(query) ||
            (pm.name || '').toLowerCase().includes(query) ||
            (pm.email || '').toLowerCase().includes(query) ||
            (enggLead.name || '').toLowerCase().includes(query) ||
            (enggLead.email || '').toLowerCase().includes(query) ||
            (enggSubLead1.name || '').toLowerCase().includes(query) ||
            (enggSubLead1.email || '').toLowerCase().includes(query) ||
            (enggSubLead2.name || '').toLowerCase().includes(query) ||
            (enggSubLead2.email || '').toLowerCase().includes(query) ||
            (enggManagers.name || '').toLowerCase().includes(query) ||
            (enggManagers.email || '').toLowerCase().includes(query) ||
            (scm.name || '').toLowerCase().includes(query) ||
            (scm.email || '').toLowerCase().includes(query) ||
            (controlling.name || '').toLowerCase().includes(query) ||
            (controlling.email || '').toLowerCase().includes(query)
          );
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
      form.addEventListener('submit', async e => {
        e.preventDefault();

        if (!canCreate()) {
          toast("You don't have permission to create projects", "error");
          return;
        }

        const payload: any = {
          PC: (document.querySelector('#pc') as HTMLInputElement).value,
          ProjectReference: (document.querySelector('#reference') as HTMLInputElement).value,
          CustomerName: (document.querySelector('#customer') as HTMLInputElement).value,
          Title: (document.querySelector('#name') as HTMLInputElement).value,
          GroupHead1Id: groupHeadLookup.getSelectedId(),
          ProjectManagerId: projectManagerLookup.getSelectedId(),
          EnggLeadId: enggLeadLookup.getSelectedId(),
          EnggSubLead1Id: enggSubLead1Lookup.getSelectedId(),
          EnggSubLead2Id: enggSubLead2Lookup.getSelectedId(),
          EngineeringManagersId: enggManagersLookup.getSelectedId(),
          SCMId: scmLookup.getSelectedId(),
          ControllingId: controllingLookup.getSelectedId()
        };

        const saveLoading = createLoadingOverlay(editingId !== null ? "Updating project..." : "Saving project...");
        loaderHost().appendChild(saveLoading);
        try {
          await saveProject(payload, editingId !== null);
        } finally {
          removeLoadingOverlay();
        }

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
          const gh = getManagerDisplayName(p, 'GroupHead1');
          const pm = getManagerDisplayName(p, 'ProjectManager');
          const enggLead = getManagerDisplayName(p, 'EnggLead');
          const enggSubLead1 = getManagerDisplayName(p, 'EnggSubLead1');
          const enggSubLead2 = getManagerDisplayName(p, 'EnggSubLead2');
          const enggManagers = getManagerDisplayName(p, 'EngineeringManagers');
          const scm = getManagerDisplayName(p, 'SCM');
          const controlling = getManagerDisplayName(p, 'Controlling');
          
          return {
            'Profit Centre': p.PC || '',
            'Project Reference': p.ProjectReference || '',
            'Customer Name': p.CustomerName || '',
            'Project Name': p.Title || '',
            'Group Head': gh.name + (gh.email ? ` (${gh.email})` : ''),
            'Project Manager': pm.name + (pm.email ? ` (${pm.email})` : ''),
            'Engg Lead': enggLead.name + (enggLead.email ? ` (${enggLead.email})` : ''),
            'Engg Sub Lead 1': enggSubLead1.name + (enggSubLead1.email ? ` (${enggSubLead1.email})` : ''),
            'Engg Sub Lead 2': enggSubLead2.name + (enggSubLead2.email ? ` (${enggSubLead2.email})` : ''),
            'Engineering Managers': enggManagers.name + (enggManagers.email ? ` (${enggManagers.email})` : ''),
            'SCM': scm.name + (scm.email ? ` (${scm.email})` : ''),
            'Controlling': controlling.name + (controlling.email ? ` (${controlling.email})` : '')
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
      // EXCEL IMPORT & TEMPLATE  (mirrors the Revenue module UX)
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

      const isServiceAccount = (title: string): boolean => {
        return !title || title.includes('System Account') ||
          title.includes('App@sharepoint') || title.includes('SharingLinks') ||
          title.includes('NT AUTHORITY') || title.includes('Membership');
      };

      // Load + cache every real user of the site once.
      const ensureEmployeesLoaded = async (): Promise<void> => {
        if (allEmployeesLoaded) return;
        try {
          let users: any[] = [];
          const res = await fetch(`${siteUrl}/_api/web/siteusers?$filter=PrincipalType eq 1&$top=5000&$orderby=Title`, {
            headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }
          });
          if (res.ok) {
            const data = await res.json();
            users = data.d?.results || [];
          }
          allEmployees = users
            .filter((u: any) => !isServiceAccount(u.Title || ''))
            .map((u: any) => ({ id: u.Id, title: u.Title || '', email: u.Email || '' }))
            .sort((a: any, b: any) => a.title.localeCompare(b.title));
          allEmployeesLoaded = true;
          console.log('👥 Employees loaded:', allEmployees.length);
        } catch (error) {
          console.error('Error loading employees:', error);
        }
      };

      // Best-effort match a cell value (Name, email, "Name (email)", or partial) to a site user.
      const matchEmployee = (value: any): any => {
        if (value === null || value === undefined) return null;
        const v = String(value).trim();
        if (!v) return null;
        const m = v.match(/\(([^)]+)\)/);
        const emailInParens = m ? m[1].trim() : '';
        const searchEmails: string[] = [];
        if (emailInParens) searchEmails.push(emailInParens);
        if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) searchEmails.push(v);   // whole cell is just an email

        for (const em of searchEmails) {
          const f = allEmployees.find(e => e.email && e.email.toLowerCase() === em.toLowerCase());
          if (f) return f;
        }
        const byName = allEmployees.find(e => e.title && e.title.toLowerCase() === v.toLowerCase());
        if (byName) return byName;
        const low = v.toLowerCase();
        return allEmployees.find(e =>
          (e.title && e.title.toLowerCase().includes(low)) ||
          (e.email && e.email.toLowerCase().includes(low))
        ) || null;
      };

      // Low-level add-or-update (match key = Project Reference).
      const writeProjectRow = async (payload: any): Promise<{ ok: boolean; mode: 'added' | 'updated' | 'error'; error?: string }> => {
        try {
          const ref = (payload.ProjectReference || '').trim().toLowerCase();
          let existing: any = null;
          if (ref) existing = projectData.find(p => (p.ProjectReference || '').trim().toLowerCase() === ref);

          const digest = await getDigest();
          const entityType = await getEntityType();
          const headers: any = {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': digest
          };
          let url = apiBase;
          if (existing) {
            url = `${apiBase}(${existing.Id})`;
            headers['X-HTTP-Method'] = 'MERGE';
            headers['IF-MATCH'] = '*';
          }

          const finalPayload: any = {
            __metadata: { type: entityType },
            PC: payload.PC || '',
            ProjectReference: payload.ProjectReference || '',
            CustomerName: payload.CustomerName || '',
            Title: payload.Title || ''
          };
          // Only write a role when a real employee id was resolved.
          // If a cell is blank (or an employee name could not be matched) we do NOT send the
          // field, so on an UPDATE the existing person is kept and on a NEW row it stays empty.
          for (const rc of ROLE_COLS) {
            const val = payload[rc.field];
            if (val && val !== '' && val !== 'null' && val !== 'undefined') finalPayload[rc.field] = parseInt(val, 10);
          }

          const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(finalPayload) });
          if (!res.ok) {
            console.error('❌ Import write error:', await res.text());
            return { ok: false, mode: 'error', error: `HTTP ${res.status}` };
          }
          return { ok: true, mode: existing ? 'updated' : 'added' };
        } catch (err) {
          console.error('❌ Import write exception:', err);
          return { ok: false, mode: 'error', error: String(err) };
        }
      };

      // ============================================
      // DOWNLOAD TEMPLATE  (includes an Employee List reference sheet)
      // ============================================
      const downloadTemplate = async (): Promise<void> => {
        await ensureEmployeesLoaded();
        if (!allEmployees.length) {
          toast('Could not load employee list for the template', 'error');
          return;
        }

        // Main sheet = Projects (first sheet so Excel opens on it)
        const headers = IMPORT_HEADERS.slice();
        const ws = XLSX.utils.aoa_to_sheet([headers]);
        ws['!cols'] = headers.map((h: string) => {
          if (h === 'Project Name' || h === 'Customer Name') return { wch: 30 };
          if (h === 'Project Reference') return { wch: 22 };
          if (h === 'Profit Centre') return { wch: 16 };
          if (h === 'Engineering Managers') return { wch: 36 };
          if (h.includes('Engg') || h === 'SCM' || h === 'Controlling' ||
              h === 'Group Head' || h === 'Project Manager') return { wch: 34 };
          return { wch: 18 };
        });

        // Add ~10 blank starter rows so it is clearly a fill-in sheet
        for (let r = 0; r < 10; r++) {
          XLSX.utils.sheet_add_aoa(ws, [[]], { origin: -1 });
        }  

        // Employee List sheet: every real user so names/emails can be copied into the role cells
        const empAoa: any[][] = [['Employee Name', 'Email']];
        allEmployees.forEach(e => empAoa.push([e.title || '', e.email || '']));
        const empWs = XLSX.utils.aoa_to_sheet(empAoa);
        empWs['!cols'] = [{ wch: 40 }, { wch: 34 }];

        // Instructions sheet
        const inst = XLSX.utils.aoa_to_sheet([
          ['How to use this template'],
          [''],
          ['1. Fill in the "Projects" sheet. One row per project.'],
          ['2. Role columns (Group Head 1, Project Manager, Engg Lead, Engg Sub Lead 1, Engg Sub Lead 2, Engineering Managers, SCM, Controlling):'],
          ['   copy a person from the "Employee List" sheet (or just type their name or email). Any of Name / email works.'],
          ['3. A project is matched and UPDATED by its "Project Reference". If no existing project has that reference, a NEW project is created.'],
          ['4. Do NOT rename, reorder, or delete the header row of the "Projects" sheet.'],
          [''],
          ['Required columns: Profit Centre, Project Reference, Customer Name, Project Name.'],
          ['You may leave role columns empty.'],
          ['The "Employee List" sheet lists every employee already added in SharePoint and is provided for your reference.']
        ]);
        inst['!cols'] = [{ wch: 100 }];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Projects');
        XLSX.utils.book_append_sheet(wb, empWs, 'Employee List');
        XLSX.utils.book_append_sheet(wb, inst, 'Instructions');
        const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        saveAs(new Blob([buf], { type: 'application/octet-stream' }), `Project_Import_Template_${new Date().toISOString().slice(0, 10)}.xlsx`);
        toast('Template downloaded (Projects + Employee List + Instructions sheets)', 'success');
      };

      // ============================================
      // IMPORT MODAL (Download -> Upload -> Preview -> Confirm)
      // ============================================
      const openImportModal = async (): Promise<void> => {
        await ensureEmployeesLoaded();
        const overlayId = 'pimp-overlay';
        document.getElementById(overlayId)?.remove();

        const overlay = document.createElement('div');
        overlay.className = 'pimp-modal-overlay';
        overlay.id = overlayId;
        overlay.innerHTML = `
          <div class="pimp-modal">
            <div class="pimp-modal-header">
              <h3>Import Projects</h3>
              <button type="button" class="pimp-close" title="Close">&times;</button>
            </div>
            <div class="pimp-modal-body">
              <div class="pimp-step">
                <h4>Step 1 — Download the template (includes every employee in the "Employee List" sheet)</h4>
                <div class="pimp-actions" style="justify-content:flex-start;">
                  <button type="button" class="pimp-btn pimp-btn-primary" id="pimp-download">⬇ Download Template</button>
                </div>
              </div>
              <div class="pimp-step">
                <h4>Step 2 — Upload your filled Excel file</h4>
                <div class="pimp-file-area" id="pimp-file-area">
                  <div class="pimp-upload-icon">📁</div>
                  <p>Click to upload or drag &amp; drop an Excel file</p>
                  <p><small>Supported: .xlsx / .xls. Header row must match the template.</small></p>
                  <input type="file" id="pimp-file" accept=".xlsx,.xls" style="display:none;" />
                  <div class="pimp-file-info" id="pimp-file-info"></div>
                </div>
              </div>
              <div class="pimp-step" id="pimp-preview-step" style="display:none;">
                <h4>Step 3 — Preview &amp; Import</h4>
                <div class="pimp-preview" id="pimp-preview"></div>
                <div class="pimp-actions">
                  <button type="button" class="pimp-btn pimp-btn-secondary" id="pimp-cancel">Cancel</button>
                  <button type="button" class="pimp-btn pimp-btn-primary" id="pimp-confirm">Import Data</button>
                </div>
                <div class="pimp-status" id="pimp-status"></div>
              </div>
            </div>
          </div>`;
        document.body.appendChild(overlay);
        setTimeout(() => overlay.classList.add('pimp-open'), 10);

        const fileArea = overlay.querySelector('#pimp-file-area') as HTMLElement;
        const fileInput = overlay.querySelector('#pimp-file') as HTMLInputElement;
        const fileInfo = overlay.querySelector('#pimp-file-info') as HTMLElement;
        const previewStep = overlay.querySelector('#pimp-preview-step') as HTMLElement;
        const previewDiv = overlay.querySelector('#pimp-preview') as HTMLElement;
        const statusDiv = overlay.querySelector('#pimp-status') as HTMLElement;
        const confirmBtn = overlay.querySelector('#pimp-confirm') as HTMLButtonElement;
        const downloadBtn = overlay.querySelector('#pimp-download') as HTMLButtonElement;

        let parsed: any[] = [];   // { rec, ref, name, existing, inFile, unresolved } processed rows

        const esc = (s: any): string => String(s ?? '')
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

        const close = () => { overlay.remove(); };
        overlay.querySelector('.pimp-close')?.addEventListener('click', close);
        overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
        overlay.querySelector('#pimp-cancel')?.addEventListener('click', close);
        downloadBtn.addEventListener('click', downloadTemplate);

        fileArea.addEventListener('click', () => fileInput.click());
        fileArea.addEventListener('dragover', (e) => { e.preventDefault(); fileArea.classList.add('pimp-dragover'); });
        fileArea.addEventListener('dragleave', () => fileArea.classList.remove('pimp-dragover'));
        fileArea.addEventListener('drop', (e) => {
          e.preventDefault(); fileArea.classList.remove('pimp-dragover');
          const files = e.dataTransfer?.files;
          if (files && files[0]) { fileInput.files = files; handleFile(files[0]); }
        });
        fileInput.addEventListener('change', (e) => {
          const files = (e.target as HTMLInputElement).files;
          if (files && files[0]) handleFile(files[0]);
        });

        const handleFile = async (file: File): Promise<void> => {
          if (!/\.(xlsx|xls)$/i.test(file.name)) {
            statusDiv.textContent = 'Please select an Excel file (.xlsx or .xls)';
            statusDiv.className = 'pimp-status error';
            return;
          }
          fileInfo.textContent = `Selected: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
          try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer);

            // Find the sheet that actually has the project headers (skip Employee List / Instructions)
            let dataRows: any[][] = [];
            let dataName = '';
            for (const name of workbook.SheetNames) {
              const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 }) as any[][];
              const head = (rows[0] || []).map((h: any) => String(h ?? '').trim());
              if (head.includes('Project Name')) { dataRows = rows; dataName = name; break; }
            }
            if (!dataRows.length && workbook.SheetNames.length) {
              dataRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 }) as any[][];
              dataName = workbook.SheetNames[0];
            }
            if (dataRows.length < 2) {
              statusDiv.textContent = 'No data rows found in the sheet "' + dataName + '".';
              statusDiv.className = 'pimp-status error';
              return;
            }

            const headers = (dataRows[0] || []).map((h: any) => String(h ?? '').trim());
            const colPos: { [label: string]: number } = {};
            IMPORT_HEADERS.forEach(lb => { colPos[lb] = headers.indexOf(lb); });
            const extract = (row: any[]): { [label: string]: any } => {
              const rec: any = {};
              for (const lb of IMPORT_HEADERS) {
                const i = colPos[lb];
                rec[lb] = i >= 0 ? row[i] : undefined;
              }
              return rec;
            };

            parsed = [];
            const refSet = new Set<string>();
            let invalid = 0;
            for (let r = 1; r < dataRows.length; r++) {
              const row = dataRows[r];
              if (!row) continue;
              const hasAny = row.some(c => c !== null && c !== undefined && String(c).trim() !== '');
              if (!hasAny) continue;
              const rec = extract(row);
              const ref = (rec['Project Reference'] ?? '').toString().trim();
              const name = (rec['Project Name'] ?? '').toString().trim();
              if (!ref || !name) { invalid++; continue; }
              const refKey = ref.toLowerCase();
              const existing = projectData.find(p => (p.ProjectReference || '').trim().toLowerCase() === refKey);
              const inFile = refSet.has(refKey);
              refSet.add(refKey);
              const unresolved: { header: string; value: string }[] = [];
              for (const rc of ROLE_COLS) {
                const raw = rec[rc.header];
                if (raw !== null && raw !== undefined && String(raw).trim() !== '') {
                  const emp = matchEmployee(raw);
                  if (!emp) unresolved.push({ header: rc.header, value: String(raw).trim() });
                }
              }
              parsed.push({ rec, ref, name, existing, inFile, unresolved, excelRow: r + 1 });   // r is 0-based in dataRows; header is row 1
            }

            if (!parsed.length) {
              statusDiv.textContent = 'No valid rows found. Make sure each row has a Project Reference and Project Name.';
              statusDiv.className = 'pimp-status error';
              return;
            }

            const addCount = parsed.filter(p => !p.existing && !p.inFile).length;
            const updCount = parsed.filter(p => !!p.existing).length;
            const skipCount = parsed.filter(p => p.inFile && !p.existing).length;
            const unresolvedTotal = parsed.reduce((s, p) => s + p.unresolved.length, 0);
            const issueRows = parsed.filter(p => p.unresolved.length > 0);

            // Show every row (scrollable). Row # = Excel row (header is row 1).
            const rowsHtml = parsed.map((p, i) => {
              const roleCell = p.unresolved.length
                ? p.unresolved.map((u: any) => `<div style="color:#d13438;line-height:1.4;">✗ ${esc(u.header)} = ${esc(u.value)}</div>`).join('')
                : '<span style="color:#107c10;">✓ all matched</span>';
              const action = p.existing
                ? '<span class="pimp-badge pimp-badge-upd">Update</span>'
                : (p.inFile ? '<span class="pimp-badge pimp-badge-add">Dup in file (skip)</span>' : '<span class="pimp-badge pimp-badge-add">Add New</span>');
              return `<tr${p.unresolved.length ? ' style="background:#fff5f5;"' : ''}>
                <td>${p.excelRow}</td>
                <td>${esc(p.ref)}</td>
                <td>${esc(p.name)}</td>
                <td>${action}</td>
                <td>${roleCell}</td>
              </tr>`;
            }).join('');

            const issuesHtml = issueRows.length
              ? `<div style="margin-top:12px;padding:10px 12px;background:#fdecea;border:1px solid #f5c6cb;border-radius:6px;font-size:12px;color:#a71d2a;">
                   <strong>⚠️ Rows with unmatched employee names</strong> — the value entered does not match any person in SharePoint, so that role cell will be skipped. On an updated project the current person is kept; on a new project it stays empty. Please check these before importing:
                   ${issueRows.map(p => {
                     const details = p.unresolved.map((u: any) => `${u.header} = "${u.value}"`).join('; ');
                     return `<div style="margin-top:6px;">• Excel row <b>${p.excelRow}</b> — <b>${esc(p.ref)}</b> / ${esc(p.name)}: <span style="color:#a71d2a;">${esc(details)}</span></div>`;
                   }).join('')}
                 </div>`
              : '';

            previewDiv.innerHTML = `
              <div class="pimp-preview-note">
                <strong>Ready for import.</strong> Found <strong>${parsed.length}</strong> project row(s):
                <strong style="color:#107c10;">${addCount}</strong> new  · 
                <strong style="color:#856404;">${updCount}</strong> will be updated (matched by Project Reference)
                ${skipCount ? `  ·  <strong style="color:#c0392b;">${skipCount}</strong> duplicate reference within the file (skipped)` : ''}
                ${invalid ? `  ·  ${invalid} incomplete row(s) skipped` : ''}
                ${unresolvedTotal ? `<br/><span style="color:#d13438;">⚠️ ${unresolvedTotal} employee name(s) could not be matched. See the red rows / list below before importing.</span>` : ''}
              </div>
              <div class="pimp-table-wrap" style="max-height:360px;">
                <table class="pimp-table">
                  <thead><tr><th style="width:60px;">Row</th><th>Project Reference</th><th>Project Name</th><th style="width:120px;">Action</th><th>Role Match</th></tr></thead>
                  <tbody>${rowsHtml}</tbody>
                </table>
              </div>
              ${issueRows.length ? issuesHtml : ''}
            `;
            previewStep.style.display = 'block';
            statusDiv.textContent = '';
            statusDiv.className = 'pimp-status';
          } catch (err) {
            console.error('Error reading Excel:', err);
            statusDiv.textContent = 'Could not read the Excel file.';
            statusDiv.className = 'pimp-status error';
          }
        };

        confirmBtn.addEventListener('click', async () => {
          if (!parsed.length) return;
          confirmBtn.disabled = true;
          confirmBtn.textContent = 'Importing…';
          statusDiv.className = 'pimp-status';
          statusDiv.textContent = 'Importing rows…';

          let added = 0, updated = 0, errors = 0, skipped = 0;
          for (const p of parsed) {
            if (p.inFile && !p.existing) { skipped++; continue; }
            const payload: any = {
              PC: (p.rec['Profit Centre'] ?? '').toString(),
              ProjectReference: p.ref,
              CustomerName: (p.rec['Customer Name'] ?? '').toString(),
              Title: p.name
            };
            for (const rc of ROLE_COLS) {
              const raw = p.rec[rc.header];
              if (raw !== null && raw !== undefined && String(raw).trim() !== '') {
                const emp = matchEmployee(raw);
                if (emp) payload[rc.field] = emp.id;
              }
            }
            const r = await writeProjectRow(payload);
            if (r.ok) { if (r.mode === 'added') added++; else updated++; }
            else errors++;
            statusDiv.textContent = `Importing…  (added: ${added}, updated: ${updated}, errors: ${errors})`;
            await new Promise(res => setTimeout(res, 120));
          }

          statusDiv.className = 'pimp-status success';
          statusDiv.textContent =
            `Import complete:\n✅ ${added} added\n🔄 ${updated} updated\n⏭️ ${skipped} duplicate-in-file\n❌ ${errors} errors`;
          confirmBtn.textContent = 'Done';
          confirmBtn.disabled = true;

          await fetchProjects();
          Swal.fire({ title: 'Import Complete', text: `${added} added, ${updated} updated, ${errors} errors.`, icon: 'success', confirmButtonColor: '#0078d4' });
        });
      };

      // ============================================
      // DOWNLOAD TEMPLATE TOOLBAR BUTTON (guarded so a missing button never crashes)
      // ============================================
      if (templateBtn) {
        templateBtn.addEventListener('click', async () => {
          if (!canRead()) { toast("You don't have permission to download the template", "error"); return; }
          const tplLoading = createLoadingOverlay("Preparing template...");
          loaderHost().appendChild(tplLoading);
          try {
            await downloadTemplate();
          } finally {
            removeLoadingOverlay();
          }
        });
      }

      importBtn.addEventListener('click', () => {
        if (!canCreate()) { toast("You don't have permission to import", "error"); return; }
        openImportModal();
      });

      // ============================================
      // INITIALIZE
      // ============================================
      fetchProjects();

    }, 300);
  }
};
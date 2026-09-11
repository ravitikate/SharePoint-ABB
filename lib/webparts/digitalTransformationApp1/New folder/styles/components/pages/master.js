import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import "../styles/common.css";
// Loading overlay component
const createLoadingOverlay = (message = "Processing...") => {
    const overlay = document.createElement('div');
    overlay.className = 'loading-overlay';
    overlay.innerHTML = `
    <div class="loading-container">
      <div class="loading-spinner"></div>
      <div class="loading-text">${message}</div>
      <div class="loading-progress"></div>
    </div>
  `;
    // Add styles if not already present
    if (!document.querySelector('#loading-styles')) {
        const style = document.createElement('style');
        style.id = 'loading-styles';
        style.textContent = `
      .loading-overlay {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(255, 255, 255, 0.95);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 99999;
        backdrop-filter: blur(3px);
      }
      
      .loading-container {
        text-align: center;
        padding: 40px;
        background: white;
        border-radius: 12px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.15);
        min-width: 300px;
        border: 1px solid #e8e8e8;
      }
      
      .loading-spinner {
        width: 50px;
        height: 50px;
        border: 4px solid #f3f3f3;
        border-top: 4px solid #0078d4;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin: 0 auto 20px;
      }
      
      .loading-text {
        font-size: 16px;
        color: #333;
        margin-bottom: 15px;
        font-weight: 500;
      }
      
      .loading-progress {
        height: 4px;
        background: #f0f0f0;
        border-radius: 2px;
        overflow: hidden;
        position: relative;
      }
      
      .loading-progress::after {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, #0078d4 0%, #00a2ff 100%);
        animation: progress 1.5s ease-in-out infinite;
      }
      
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
      
      @keyframes progress {
        0% { left: -100%; }
        100% { left: 100%; }
      }
    `;
        document.head.appendChild(style);
    }
    return overlay;
};
const updateLoadingMessage = (overlay, message) => {
    const textElement = overlay.querySelector('.loading-text');
    if (textElement) {
        textElement.textContent = message;
    }
};
const removeLoadingOverlay = () => {
    const overlay = document.querySelector('.loading-overlay');
    if (overlay) {
        overlay.remove();
    }
};
// Toast notification system
const showToast = (message, type = 'info', duration = 3000) => {
    const toast = document.createElement('div');
    toast.className = `custom-toast toast-${type}`;
    toast.innerHTML = `
    <div class="toast-icon">
      ${type === 'success' ? '✅' : type === 'error' ? '❌' : type === 'warning' ? '⚠️' : 'ℹ️'}
    </div>
    <div class="toast-content">${message}</div>
    <button class="toast-close">&times;</button>
  `;
    // Add toast styles if not already present
    if (!document.querySelector('#toast-styles')) {
        const style = document.createElement('style');
        style.id = 'toast-styles';
        style.textContent = `
      .custom-toast {
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 16px 20px;
        background: white;
        border-radius: 8px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.15);
        display: flex;
        align-items: center;
        gap: 12px;
        z-index: 10000;
        min-width: 300px;
        max-width: 400px;
        border-left: 4px solid;
        animation: slideInRight 0.3s ease-out;
        transform: translateX(120%);
        opacity: 0;
        transition: transform 0.3s, opacity 0.3s;
      }
      
      .custom-toast.show {
        transform: translateX(0);
        opacity: 1;
      }
      
      .custom-toast.hide {
        transform: translateX(120%);
        opacity: 0;
      }
      
      .toast-success {
        border-left-color: #28a745;
        background: linear-gradient(135deg, #d4edda 0%, #ffffff 100%);
      }
      
      .toast-error {
        border-left-color: #dc3545;
        background: linear-gradient(135deg, #f8d7da 0%, #ffffff 100%);
      }
      
      .toast-info {
        border-left-color: #17a2b8;
        background: linear-gradient(135deg, #d1ecf1 0%, #ffffff 100%);
      }
      
      .toast-warning {
        border-left-color: #ffc107;
        background: linear-gradient(135deg, #fff3cd 0%, #ffffff 100%);
      }
      
      .toast-icon {
        font-size: 20px;
      }
      
      .toast-content {
        flex: 1;
        font-size: 14px;
        color: #333;
      }
      
      .toast-close {
        background: none;
        border: none;
        font-size: 18px;
        cursor: pointer;
        color: #999;
        padding: 0;
        width: 24px;
        height: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
      }
      
      .toast-close:hover {
        background: rgba(0, 0, 0, 0.1);
      }
      
      @keyframes slideInRight {
        from {
          transform: translateX(120%);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
    `;
        document.head.appendChild(style);
    }
    document.body.appendChild(toast);
    // Show toast
    requestAnimationFrame(() => {
        toast.classList.add('show');
    });
    // Auto dismiss
    const dismiss = () => {
        toast.classList.add('hide');
        setTimeout(() => {
            if (toast.parentNode) {
                toast.remove();
            }
        }, 300);
    };
    // Close button
    const closeBtn = toast.querySelector('.toast-close');
    if (closeBtn) {
        closeBtn.addEventListener('click', dismiss);
    }
    setTimeout(dismiss, duration);
    return { dismiss, toast };
};
const projectContent = `
<div class="project-page fade-in">
  <section class="table-section card master-dashboard">
    <div class="page-header"> 
      <div class="header-actions">
        <button class="btn-primary import-master-btn">📤 Import Master Excel</button>
        <input type="file" class="master-excel-input" accept=".xlsx,.xls" style="display:none;" />
        <div class="download-template">
          <button class="btn-secondary download-template-btn">📥 Download Template</button>
        </div>
      </div>
    </div>
    

    <div class="table-actions">
      <input type="text" class="search-input" placeholder="🔍 Search by Project Reference..." />
      <button class="btn-secondary reset-search-btn">Reset Filter</button>
    </div>
    
    <div style="overflow-x:auto;">
      <table class="modern-table master-table">
        <thead>
          <tr>
            <th>#</th>
            <th>Project Reference</th>
            <th>Project Name</th>
            <th>Receivables</th>
            <th>OBL</th>
            <th>Balance to Invoice</th>
            <th>SEI/BIES </th>
            <th>Inventory</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody class="project-table-body"></tbody>
      </table>
    </div>
    
    <div class="pagination-container">
      <div class="pagination"></div>
    </div>
    
    <div class="import-summary card" style="display:none;">
      <h3>📊 Import Summary</h3>
      <div class="summary-grid">
        <div class="summary-item success">
          <span class="summary-label">Success</span>
          <span class="summary-value" id="successCount">0</span>
        </div>
        <div class="summary-item warning">
          <span class="summary-label">New Projects</span>
          <span class="summary-value" id="newProjectCount">0</span>
        </div>
        <div class="summary-item error">
          <span class="summary-label">Errors</span>
          <span class="summary-value" id="errorCount">0</span>
        </div>
        <div class="summary-item info">
          <span class="summary-label">Total Processed</span>
          <span class="summary-value" id="totalCount">0</span>
        </div>
      </div>
    </div>
  </section>

  <section class="edit-section card" style="display:none;">
    <div class="form-header"><h2>✏️ Edit Master Data</h2></div>
    <form id="masterEditForm">
      <div class="form-group">
        <label>Project Reference</label>
        <input type="text" id="editProjectReference" readonly class="readonly-field" />
      </div>
      
      <div class="form-group">
        <label>Project Name</label>
        <input type="text" id="editProjectName" readonly class="readonly-field" required />
      </div>
      
      <div class="form-section-header">Collection Data</div>
      <div class="form-row">
        <div class="form-group"><label>NormalDue</label><input type="number" id="editNormalDue" step="0.0001" /></div>
        <div class="form-group"><label>NormalOverdue</label><input type="number" id="editNormalOverdue" step="0.0001" /></div>
        <div class="form-group"><label>RetentionDue</label><input type="number" id="editRetentionDue" step="0.0001" /></div>
        <div class="form-group"><label>RetentionOverdue</label><input type="number" id="editRetentionOverdue" step="0.0001" /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>GstTds</label><input type="number" id="editGstTds" step="0.0001" /></div>
        <div class="form-group"><label>WHT</label><input type="number" id="editWHT" step="0.0001" /></div>
        <div class="form-group"><label>ReceivableValue</label><input type="number" id="editReceivableValue" step="0.0001" /></div>
      </div>
      
      <div class="form-section-header">Other Data</div>
      <div class="form-row">
        <div class="form-group"><label>Revenue Forecast</label><input type="number" id="editRevenueForecast" step="0.0001" /></div>
        <div class="form-group"><label>Invoicing Planning</label><input type="number" id="editInvoicingPlanning" step="0.0001" /></div>
        <div class="form-group"><label>SEIBIES</label><input type="number" id="editSEIBIES" step="0.0001" /></div>
        <div class="form-group"><label>Inventory</label><input type="number" id="editInventory" step="0.0001" /></div>
      </div>
      
      <div class="form-actions">
        <button type="submit" class="btn-primary">💾 Update Data</button>
        <button type="button" class="btn-secondary back-to-table">← Back to List</button>
      </div>
    </form>
  </section>
</div>
`;
export default {
    content: projectContent,
    style: ` 
    /* Add to existing styles */
    .form-group input[type="number"]::-webkit-inner-spin-button,
    .form-group input[type="number"]::-webkit-outer-spin-button {
      opacity: 1;
      margin: 0;
    }
    
    .form-group input[type="number"]::placeholder {
      color: #999;
      font-size: 12px;
    }
    
    .project-page { font-family:'Segoe UI', Arial,sans-serif; color:#333; animation:fadeIn 0.4s ease-in; }
    
    .card { background:#fff; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.08); padding:20px; margin-bottom:20px; transition:all 0.3s ease; border:1px solid #e8e8e8; }
    
    .master-dashboard { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); }
    
    .page-header { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 2px solid #0078d4; }
    .page-header h1 { color: #0056b3; margin: 0; font-size: 28px; font-weight: 600; }
    .header-actions { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
    
    .btn-primary { background: linear-gradient(135deg, #0078d4 0%, #005a9e 100%); color: white; border: none; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; display: inline-flex; align-items: center; gap: 8px; transition: all 0.3s ease; box-shadow: 0 2px 4px rgba(0,120,212,0.2); }
    .btn-primary:hover { background: linear-gradient(135deg, #005a9e 0%, #004578 100%); transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,120,212,0.3); }
    .btn-secondary { background: #6c757d; color: white; border: none; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 500; display: inline-flex; align-items: center; gap: 8px; transition: all 0.3s ease; }
    .btn-secondary:hover { background: #5a6268; transform: translateY(-1px); }
    
    .import-instructions { background: #f8f9fa; border-left: 4px solid #0078d4; }
    .import-instructions h3 { color: #0056b3; margin-top: 0; margin-bottom: 15px; }
    .instructions-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 5px; }
    .instruction-col h4 { color: #495057; margin-top: 0; margin-bottom: 10px; font-size: 16px; }
    .instruction-col ul { margin: 0; padding-left: 5px; }
    .instruction-col li { margin-bottom: 5px; color: #666; }
    
    .master-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 14px; }
    .master-table th { background: linear-gradient(135deg, #f3f2f1 0%, #e1e1e1 100%); text-align: left; padding: 15px; font-weight: 600; border-bottom: 2px solid #0078d4; color: #0056b3; position: sticky; top: 0; z-index: 10; }
    .master-table td { padding: 12px 15px; border-bottom: 1px solid #e8e8e8; vertical-align: top; }
    .master-table tr:hover { background: #f8f9fa; }
    .master-table tr:nth-child(even) { background: #fafafa; }
    .master-table tr:nth-child(even):hover { background: #f5f5f5; }
    
    .table-actions { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; flex-wrap: wrap; gap: 10px; }
    .search-input { padding: 10px 15px; border: 2px solid #e1e1e1; border-radius: 8px; flex: 1; min-width: 250px; font-size: 14px; transition: all 0.3s ease; }
    .search-input:focus { border-color: #0078d4; box-shadow: 0 0 0 3px rgba(0,120,212,0.1); outline: none; }
    
    .import-summary { background: #f8f9fa; border-top: 3px solid #28a745; }
    .import-summary h3 { margin-top: 0; color: #28a745; }
    .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
    .summary-item { padding: 15px; border-radius: 8px; text-align: center; }
    .summary-item.success { background: #d4edda; border: 1px solid #c3e6cb; }
    .summary-item.warning { background: #fff3cd; border: 1px solid #ffeaa7; }
    .summary-item.error { background: #f8d7da; border: 1px solid #f5c6cb; }
    .summary-item.info { background: #d1ecf1; border: 1px solid #bee5eb; }
    .summary-label { display: block; font-size: 14px; color: #666; margin-bottom: 5px; }
    .summary-value { display: block; font-size: 24px; font-weight: 600; }
    
    .edit-section { max-width: 900px; margin: 0 auto; }
    .form-section-header { font-size: 16px; font-weight: 600; color: #0078d4; margin: 20px 0 15px 0; padding-bottom: 8px; border-bottom: 2px solid #e8e8e8; }
    .form-row { display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 15px; }
    .form-group { flex: 1 1 calc(25% - 20px); display: flex; flex-direction: column; min-width: 200px; }
    .form-group label { font-weight: 500; color: #495057; margin-bottom: 8px; font-size: 14px; }
    .form-group input { width: 100%; padding: 12px; border: 2px solid #e1e1e1; border-radius: 6px; font-size: 14px; transition: all 0.3s ease; }
    .form-group input:focus { border-color: #0078d4; box-shadow: 0 0 0 3px rgba(0,120,212,0.1); outline: none; }
    .readonly-field { background: #f8f9fa; color: #6c757d; cursor: not-allowed; }
    
    .form-actions { text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #e8e8e8; }
    
    .collection-badge { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500; margin: 2px; }
    .collection-badge.due { background: #cfe2ff; color: #084298; }
    .collection-badge.overdue { background: #f8d7da; color: #842029; }
    .collection-badge.receivable { background: #d1e7dd; color: #0f5132; }
    .collection-badge.gst { background: #f8f1d7f0; color: #842029; }
    .collection-badge.wht { background: #a9d2b8b3; color: #0f5132; }
     
    .data-cell { font-weight: 500; color: #495057; }
    .data-cell.positive { color: #198754; }
    .data-cell.negative { color: #dc3545; }
    
    .actions-cell { display: flex; gap: 8px; }
    .edit-btn, .delete-btn { width: 36px; height: 36px; border: none; border-radius: 6px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; padding: 0; transition: all 0.3s ease; font-size: 16px; }
    .edit-btn { background: linear-gradient(135deg, #ffc107 0%, #e0a800 100%); color: #fff; }
    .delete-btn { background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); color: #fff; }
    .edit-btn:hover { background: linear-gradient(135deg, #e0a800 0%, #d39e00 100%); transform: translateY(-2px); }
    .delete-btn:hover { background: linear-gradient(135deg, #c82333 0%, #bd2130 100%); transform: translateY(-2px); }
    
    .download-template { margin-left: auto; }
    
    .pagination-container { margin-top: 20px; }
    .pagination { display: flex; gap: 8px; flex-wrap: wrap; justify-content: center; }
    .pagination button { padding: 8px 16px; border: 2px solid #dee2e6; background: white; border-radius: 6px; cursor: pointer; transition: all 0.3s ease; }
    .pagination button:hover { border-color: #0078d4; background: #f8f9fa; }
    .pagination button.active { background: #0078d4; color: white; border-color: #0078d4; }
    
    @media (max-width: 768px) {
      .page-header { flex-direction: column; align-items: stretch; }
      .header-actions { flex-direction: column; }
      .instructions-grid { grid-template-columns: 1fr; }
      .summary-grid { grid-template-columns: 1fr; }
      .form-group { flex: 1 1 100%; }
      .master-table { font-size: 12px; }
      .master-table th, .master-table td { padding: 8px; }
    }
  `,
    script: () => {
        setTimeout(() => {
            // SharePoint list names
            const projectListName = "Project List";
            const collectionListName = "Collection List";
            const revenueListName = "Revenue Forecast List";
            const invoicingListName = "InvoicingPlanning";
            const seibiesListName = "SEIBIES";
            const inventoryListName = "Inventory";
            const tableBody = document.querySelector('.project-table-body');
            const importBtn = document.querySelector('.import-master-btn');
            const excelInput = document.querySelector('.master-excel-input');
            const searchInput = document.querySelector('.search-input');
            const resetSearchBtn = document.querySelector('.reset-search-btn');
            const summaryDiv = document.querySelector('.import-summary');
            const editSection = document.querySelector('.edit-section');
            const tableSection = document.querySelector('.table-section');
            const editForm = document.querySelector('#masterEditForm');
            const backBtn = document.querySelector('.back-to-table');
            const downloadTemplateBtn = document.querySelector('.download-template-btn');
            const paginationDiv = document.querySelector('.pagination');
            // Summary counters
            const successCountEl = document.querySelector('#successCount');
            const errorCountEl = document.querySelector('#errorCount');
            const totalCountEl = document.querySelector('#totalCount');
            // Global variables with proper typing
            let masterData = [];
            let currentEditingProject = null;
            let currentPage = 1;
            const pageSize = 10;
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/DigitalTransformationTest";
            const getDigest = async () => {
                try {
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
                }
                catch (error) {
                    console.error("Error getting digest:", error);
                    throw error;
                }
            };
            // Enhanced fetchMasterData with loading overlay
            const fetchMasterData = async () => {
                // Show loading overlay
                const loadingOverlay = createLoadingOverlay("Loading master data...");
                document.body.appendChild(loadingOverlay);
                try {
                    updateLoadingMessage(loadingOverlay, "Fetching project list...");
                    // Fetch projects
                    const projects = await fetchListData(projectListName, [
                        'ProjectReference',
                        'Project_x0020_Reference',
                        'ProjectName',
                        'Project_Name',
                        'Title'
                    ]);
                    // Initialize master data with projects only first
                    masterData = projects.map((project) => {
                        const projectRef = project.ProjectReference ||
                            project.Project_x0020_Reference ||
                            project.ProjectName ||
                            project.Project_Name ||
                            project.Title ||
                            '';
                        return {
                            id: project.Id,
                            projectReference: projectRef,
                            projectName: project.Title || '',
                            collection: {
                                normalDue: 0,
                                normalOverdue: 0,
                                retentionDue: 0,
                                retentionOverdue: 0,
                                gstTds: 0,
                                wht: 0,
                                receivableValue: 0
                            },
                            revenueForecast: 0,
                            invoicingPlanning: 0,
                            seibies: 0,
                            inventory: 0
                        };
                    });
                    // ========== FETCH COLLECTION DATA ==========
                    updateLoadingMessage(loadingOverlay, "Fetching collection data...");
                    const collections = await fetchListData(collectionListName, [
                        'NormalDue',
                        'NormalOverdue',
                        'RetentionDue',
                        'RetentionOverdue',
                        'GstTds',
                        'WHT',
                        'ReceivableValue',
                        'OBL',
                        'OBLData'
                    ]);
                    // Match collection data
                    masterData.forEach(project => {
                        let collection = null;
                        const matchingStrategies = [
                            (item) => item.Title && item.Title.includes(project.projectReference),
                            (item) => item.Title === project.projectReference,
                            (item) => item.Project && String(item.Project).includes(project.projectReference)
                        ];
                        for (const strategy of matchingStrategies) {
                            collection = collections.find(strategy);
                            if (collection)
                                break;
                        }
                        if (collection) {
                            // Helper function to extract numeric value
                            const getValue = (item, possibleFieldNames) => {
                                for (const fieldName of possibleFieldNames) {
                                    if (item[fieldName] !== undefined && item[fieldName] !== null && item[fieldName] !== '') {
                                        let value = item[fieldName];
                                        if (typeof value === 'object' && value !== null) {
                                            value = value.Title || value.Value || value.Name || value;
                                        }
                                        const numValue = parseFloat(value);
                                        if (!isNaN(numValue)) {
                                            return numValue;
                                        }
                                    }
                                }
                                return 0;
                            };
                            project.collection.normalDue = getValue(collection, ['NormalDue', 'Normal_Due', 'Normal_x0020_Due', 'Normal']);
                            project.collection.normalOverdue = getValue(collection, ['NormalOverdue', 'Normal_Overdue', 'Normal_x0020_Overdue', 'Overdue']);
                            project.collection.retentionDue = getValue(collection, ['RetentionDue', 'Retention_Due', 'Retention_x0020_Due', 'Retention']);
                            project.collection.retentionOverdue = getValue(collection, ['RetentionOverdue', 'Retention_Overdue', 'Retention_x0020_Overdue', 'RetentionOver']);
                            project.collection.gstTds = getValue(collection, ['GstTds', 'GSTTDS', 'Gst_Tds', 'GST_TDS', 'Gst_x0020_Tds', 'GST', 'TDS']);
                            project.collection.wht = getValue(collection, ['WHT', 'WithholdingTax', 'Withholding_Tax', 'Withholding']);
                            project.collection.receivableValue = getValue(collection, ['ReceivableValue', 'Receivable_Value', 'Receivable_x0020_Value', 'Receivable', 'OBL', 'Value', 'Amount', 'Total']);
                        }
                    });
                    // ========== FETCH OTHER LISTS DATA ==========
                    updateLoadingMessage(loadingOverlay, "Fetching revenue forecast data...");
                    await fetchAndMatchOBLData(revenueListName, 'revenueForecast');
                    updateLoadingMessage(loadingOverlay, "Fetching invoicing planning data...");
                    await fetchAndMatchOBLData(invoicingListName, 'invoicingPlanning');
                    updateLoadingMessage(loadingOverlay, "Fetching SEIBIES data...");
                    await fetchAndMatchOBLData(seibiesListName, 'seibies');
                    updateLoadingMessage(loadingOverlay, "Fetching inventory data...");
                    await fetchAndMatchOBLData(inventoryListName, 'inventory');
                    // ========== FINAL DATA PROCESSING ==========
                    updateLoadingMessage(loadingOverlay, "Rendering table...");
                    // Check if we have any non-zero data
                    const hasData = masterData.some(project => {
                        const hasCollectionData = Object.values(project.collection).some(val => val !== 0);
                        const hasOtherData = project.revenueForecast !== 0 ||
                            project.invoicingPlanning !== 0 ||
                            project.seibies !== 0 ||
                            project.inventory !== 0;
                        return hasCollectionData || hasOtherData;
                    });
                    // Render the table
                    renderTable(masterData);
                    // Remove loading overlay
                    removeLoadingOverlay();
                    // Show appropriate toast message
                    if (hasData) {
                        const dataCount = masterData.filter(p => Object.values(p.collection).some(v => v !== 0) ||
                            p.revenueForecast !== 0 ||
                            p.invoicingPlanning !== 0 ||
                            p.seibies !== 0 ||
                            p.inventory !== 0).length;
                        showToast(`Loaded ${dataCount} projects with data out of ${masterData.length} total`, "success");
                    }
                    else if (masterData.length > 0) {
                        showToast(`Found ${masterData.length} projects but no data found in lists. Import data first.`, "info");
                    }
                    else {
                        showToast("No projects found. Check Project List.", "info");
                    }
                }
                catch (error) {
                    console.error("Error fetching master data:", error);
                    removeLoadingOverlay();
                    showToast("Error loading master data. Check console for details.", "error");
                    masterData = [];
                    renderTable(masterData);
                }
            };
            // Helper function to fetch and match OBL data
            const fetchAndMatchOBLData = async (listName, targetField) => {
                const items = await fetchListData(listName, [
                    'OBL',
                    'OBLValue',
                    'Value',
                    'Amount',
                    'Total',
                    'Forecast',
                    'Revenue',
                    'Inventory',
                    'InventoryActual'
                ]);
                masterData.forEach(project => {
                    let matchedItem = null;
                    const matchingStrategies = [
                        (item) => item.Title && item.Title.includes(project.projectReference),
                        (item) => item.Title === project.projectReference,
                        (item) => item.Project && String(item.Project).includes(project.projectReference),
                        (item) => {
                            const allValues = Object.values(item).map((v) => String(v).toLowerCase());
                            return allValues.some(val => val.includes(project.projectReference.toLowerCase()));
                        }
                    ];
                    for (const strategy of matchingStrategies) {
                        matchedItem = items.find(strategy);
                        if (matchedItem)
                            break;
                    }
                    if (matchedItem) {
                        const possibleFields = ['OBL', 'Obl', 'obl', 'OBLValue', 'Value', 'Amount', 'Total', 'Revenue', 'Forecast', 'Inventory', 'InventoryActual', 'Data'];
                        for (const field of possibleFields) {
                            if (matchedItem[field] !== undefined && matchedItem[field] !== null && matchedItem[field] !== '') {
                                let value = matchedItem[field];
                                if (typeof value === 'object' && value !== null) {
                                    value = value.Title || value.Value || value.Name || value;
                                }
                                const numValue = parseFloat(value);
                                if (!isNaN(numValue)) {
                                    switch (targetField) {
                                        case 'revenueForecast':
                                            project.revenueForecast = numValue;
                                            break;
                                        case 'invoicingPlanning':
                                            project.invoicingPlanning = numValue;
                                            break;
                                        case 'seibies':
                                            project.seibies = numValue;
                                            break;
                                        case 'inventory':
                                            project.inventory = numValue;
                                            break;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                });
            };
            const fetchListData = async (listName, specificFields) => {
                try {
                    const exists = await checkListExists(listName);
                    if (!exists) {
                        return [];
                    }
                    const fields = await getListFieldsWithNames(listName);
                    if (fields.length === 0) {
                        return [];
                    }
                    let selectFields = [];
                    let expandFields = [];
                    if (!specificFields || specificFields.length === 0) {
                        selectFields = fields.map(f => f.internalName);
                    }
                    else {
                        selectFields = ['Title', 'Id'];
                        for (const desiredField of specificFields) {
                            const actualName = findActualFieldName(fields, desiredField);
                            if (actualName && !selectFields.includes(actualName)) {
                                const field = fields.find(f => f.internalName === actualName);
                                if (field && field.type === 'Lookup') {
                                    selectFields.push(`${actualName}Id`);
                                    expandFields.push(actualName);
                                }
                                else {
                                    selectFields.push(actualName);
                                }
                            }
                        }
                    }
                    const uniqueFields = [...new Set(selectFields)];
                    if (uniqueFields.length === 0) {
                        return [];
                    }
                    let query = `$select=${encodeURIComponent(uniqueFields.join(','))}`;
                    if (expandFields.length > 0) {
                        query += `&$expand=${encodeURIComponent(expandFields.join(','))}`;
                    }
                    query += `&$orderby=Id desc&$top=1000`;
                    const url = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?${query}`;
                    const res = await fetch(url, {
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose'
                        }
                    });
                    if (!res.ok) {
                        const errorText = await res.text();
                        console.error(`Failed to fetch ${listName}:`, errorText);
                        return [];
                    }
                    const data = await res.json();
                    return data.d.results || [];
                }
                catch (error) {
                    console.error(`Error fetching ${listName}:`, error);
                    return [];
                }
            };
            const getListFieldsWithNames = async (listName) => {
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/fields?$select=InternalName,Title,TypeAsString,Hidden&$filter=Hidden eq false`, {
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose'
                        }
                    });
                    if (res.ok) {
                        const data = await res.json();
                        const filteredFields = data.d.results.filter((f) => !f.InternalName.startsWith('_') &&
                            f.TypeAsString !== 'Attachments' &&
                            f.TypeAsString !== 'Computed' &&
                            !f.InternalName.includes('ColorTag') &&
                            f.InternalName !== 'ID' &&
                            f.InternalName !== 'ContentType');
                        return filteredFields.map((f) => ({
                            internalName: f.InternalName,
                            title: f.Title,
                            type: f.TypeAsString
                        }));
                    }
                    return [];
                }
                catch (error) {
                    console.error(`Error getting fields for ${listName}:`, error);
                    return [];
                }
            };
            const findActualFieldName = (fields, desiredName) => {
                const lowerDesired = desiredName.toLowerCase().trim();
                const exactMatch = fields.find(f => f.internalName.toLowerCase() === lowerDesired ||
                    f.title.toLowerCase() === lowerDesired);
                if (exactMatch)
                    return exactMatch.internalName;
                const cleanDesired = desiredName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                const cleanMatch = fields.find(f => f.internalName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === cleanDesired);
                if (cleanMatch)
                    return cleanMatch.internalName;
                const variations = [
                    desiredName,
                    `${desiredName}_x0020_`,
                    `${desiredName}_x0020_0`,
                    `${desiredName}0`,
                    desiredName.replace(/ /g, '_x0020_'),
                    desiredName.replace(/[^a-zA-Z0-9]/g, '')
                ];
                for (const variation of variations) {
                    const match = fields.find(f => f.internalName.toLowerCase() === variation.toLowerCase());
                    if (match)
                        return match.internalName;
                }
                const partialMatch = fields.find(f => f.internalName.toLowerCase().includes(lowerDesired) ||
                    f.title.toLowerCase().includes(lowerDesired));
                if (partialMatch)
                    return partialMatch.internalName;
                return null;
            };
            const checkListExists = async (listName) => {
                try {
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')`, {
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose'
                        }
                    });
                    return res.ok;
                }
                catch (error) {
                    return false;
                }
            };
            const renderTable = (data) => {
                if (!tableBody) {
                    console.error("Table body not found");
                    return;
                }
                const startIndex = (currentPage - 1) * pageSize;
                const paginatedData = data.slice(startIndex, startIndex + pageSize);
                if (paginatedData.length === 0) {
                    tableBody.innerHTML = `
            <tr>
              <td colspan="9" style="text-align: center; padding: 40px; color: #666;">
                <h3>No projects found</h3>
                <p>Check if Project List has items, or import master data.</p>
              </td>
            </tr>
          `;
                    return;
                }
                const formatNumber = (value) => {
                    return Math.round(value !== null && value !== void 0 ? value : 0).toLocaleString('en-IN');
                };
                tableBody.innerHTML = paginatedData.map((item, index) => {
                    const rowNum = startIndex + index + 1;
                    const collection = item.collection;
                    return `
      <tr>
        <td>${rowNum}</td>
        <td><strong>${item.projectReference || 'N/A'}</strong></td>
        <td>${item.projectName || '-'}</td>

        <td>
          <div class="collection-data">
            <div><span class="collection-badge due">NormalDue: ${formatNumber(collection.normalDue)}</span></div>
            <div><span class="collection-badge overdue">NormalOverdue: ${formatNumber(collection.normalOverdue)}</span></div>
            <div><span class="collection-badge due">RetentionDue: ${formatNumber(collection.retentionDue)}</span></div>
            <div><span class="collection-badge overdue">RetentionOverdue: ${formatNumber(collection.retentionOverdue)}</span></div>
            <div><span class="collection-badge gst"><small>GstTds: ${formatNumber(collection.gstTds)}</small></div>
            <div><span class="collection-badge wht"><small>WHT: ${formatNumber(collection.wht)}</small></div>
            <div><span class="collection-badge receivable">Receivable: ${formatNumber(collection.receivableValue)}</span></div>
          </div>
        </td>

        <td class="data-cell ${item.revenueForecast > 0 ? 'positive' : ''}">
          ${formatNumber(item.revenueForecast)}
        </td>

        <td class="data-cell ${item.invoicingPlanning > 0 ? 'positive' : ''}">
          ${formatNumber(item.invoicingPlanning)}
        </td>

        <td class="data-cell ${item.seibies > 0 ? 'positive' : ''}">
          ${formatNumber(item.seibies)}
        </td>

        <td class="data-cell ${item.inventory > 0 ? 'positive' : ''}">
          ${formatNumber(item.inventory)}
        </td>

        <td class="actions-cell">
          <button class="edit-btn" data-id="${item.id}" data-ref="${item.projectReference}" title="Edit">✏️</button>
        </td>
      </tr>`;
                }).join('');
                attachEditHandlers();
                const totalPages = Math.ceil(data.length / pageSize);
                renderPagination(totalPages);
            };
            const attachEditHandlers = () => {
                document.querySelectorAll('.edit-btn').forEach(btn => {
                    btn.addEventListener('click', async () => {
                        const projectRef = btn.dataset.ref;
                        const project = masterData.find(p => p.projectReference === projectRef);
                        if (project) {
                            currentEditingProject = project;
                            document.querySelector('#editProjectReference').value = project.projectReference;
                            document.querySelector('#editProjectName').value = project.projectName;
                            document.querySelector('#editNormalDue').value = project.collection.normalDue.toString();
                            document.querySelector('#editNormalOverdue').value = project.collection.normalOverdue.toString();
                            document.querySelector('#editRetentionDue').value = project.collection.retentionDue.toString();
                            document.querySelector('#editRetentionOverdue').value = project.collection.retentionOverdue.toString();
                            document.querySelector('#editGstTds').value = project.collection.gstTds.toString();
                            document.querySelector('#editWHT').value = project.collection.wht.toString();
                            document.querySelector('#editReceivableValue').value = project.collection.receivableValue.toString();
                            document.querySelector('#editRevenueForecast').value = project.revenueForecast.toString();
                            document.querySelector('#editInvoicingPlanning').value = project.invoicingPlanning.toString();
                            document.querySelector('#editSEIBIES').value = project.seibies.toString();
                            document.querySelector('#editInventory').value = project.inventory.toString();
                            if (editSection)
                                editSection.style.display = "block";
                            if (tableSection)
                                tableSection.style.display = "none";
                            formatNumberInputs();
                        }
                    });
                });
            };
            const upsertListItem = async (listName, data, lookupTitle) => {
                var _a, _b, _c;
                try {
                    const digest = await getDigest();
                    const fields = await getListFieldsWithNames(listName);
                    if (fields.length === 0) {
                        console.error(`No fields found for ${listName}`);
                        return { success: false, id: null };
                    }
                    const findRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$filter=Title eq '${encodeURIComponent(lookupTitle)}'&$top=1`, {
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose'
                        }
                    });
                    let itemId = null;
                    let itemExists = false;
                    if (findRes.ok) {
                        const findData = await findRes.json();
                        if (findData.d.results.length > 0) {
                            itemId = findData.d.results[0].Id;
                            itemExists = true;
                        }
                    }
                    const listType = `SP.Data.${listName.replace(/\s+/g, '_x0020_')}ListItem`;
                    const requestBody = {
                        __metadata: { type: listType }
                    };
                    for (const [key, value] of Object.entries(data)) {
                        const fieldValue = value;
                        if (key === 'project' || key === 'Project') {
                            const projectFieldName = findActualFieldName(fields, 'project') ||
                                findActualFieldName(fields, 'Project');
                            if (projectFieldName) {
                                const fieldType = (_a = fields.find(f => f.internalName === projectFieldName)) === null || _a === void 0 ? void 0 : _a.type;
                                if (fieldType === 'Lookup') {
                                    const projectNameMatch = lookupTitle.match(/^[^-]+ - (.+)$/);
                                    const projectName = projectNameMatch ? projectNameMatch[1] : String(fieldValue);
                                    const projectLookupRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(projectListName)}')/items?$filter=Title eq '${encodeURIComponent(projectName)}'&$select=Id&$top=1`, {
                                        headers: {
                                            'Accept': 'application/json;odata=verbose',
                                            'Content-Type': 'application/json;odata=verbose'
                                        }
                                    });
                                    if (projectLookupRes.ok) {
                                        const projectData = await projectLookupRes.json();
                                        if (projectData.d.results.length > 0) {
                                            const projectId = projectData.d.results[0].Id;
                                            requestBody[`${projectFieldName}Id`] = projectId;
                                        }
                                    }
                                }
                                else if (fieldType === 'Text' || fieldType === 'Note') {
                                    requestBody[projectFieldName] = String(fieldValue);
                                }
                            }
                        }
                        else if (key === 'Title') {
                            requestBody['Title'] = String(fieldValue);
                        }
                        else if (key === 'OBL' || key === 'NormalDue' || key === 'NormalOverdue' ||
                            key === 'RetentionDue' || key === 'RetentionOverdue' ||
                            key === 'GstTds' || key === 'WHT' || key === 'ReceivableValue') {
                            const fieldName = findActualFieldName(fields, key);
                            if (fieldName) {
                                const fieldType = (_b = fields.find(f => f.internalName === fieldName)) === null || _b === void 0 ? void 0 : _b.type;
                                if (fieldType === 'Number' || fieldType === 'Currency') {
                                    requestBody[fieldName] = typeof fieldValue === 'string' ? parseFloat(fieldValue) || 0 : Number(fieldValue) || 0;
                                }
                                else if (fieldType === 'Text') {
                                    requestBody[fieldName] = String(fieldValue);
                                }
                            }
                        }
                        else {
                            const fieldName = findActualFieldName(fields, key);
                            if (fieldName) {
                                const fieldType = (_c = fields.find(f => f.internalName === fieldName)) === null || _c === void 0 ? void 0 : _c.type;
                                if (fieldType === 'Number' || fieldType === 'Currency') {
                                    requestBody[fieldName] = typeof fieldValue === 'string' ? parseFloat(fieldValue) || 0 : Number(fieldValue) || 0;
                                }
                                else if (fieldType === 'Boolean') {
                                    requestBody[fieldName] = Boolean(fieldValue);
                                }
                                else {
                                    requestBody[fieldName] = String(fieldValue);
                                }
                            }
                        }
                    }
                    if (!requestBody['Title'] && data['Title']) {
                        requestBody['Title'] = String(data['Title']);
                    }
                    let response;
                    let createdId = null;
                    if (itemExists && itemId) {
                        response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`, {
                            method: "POST",
                            headers: {
                                "Accept": "application/json;odata=verbose",
                                "Content-Type": "application/json;odata=verbose",
                                "X-RequestDigest": digest,
                                "X-HTTP-Method": "MERGE",
                                "IF-MATCH": "*"
                            },
                            body: JSON.stringify(requestBody)
                        });
                        if (response.ok)
                            createdId = itemId;
                    }
                    else {
                        response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`, {
                            method: "POST",
                            headers: {
                                "Accept": "application/json;odata=verbose",
                                "Content-Type": "application/json;odata=verbose",
                                "X-RequestDigest": digest
                            },
                            body: JSON.stringify(requestBody)
                        });
                        if (response.ok) {
                            try {
                                const result = await response.json();
                                createdId = result.d.Id;
                            }
                            catch (e) {
                                console.log('Could not parse create response');
                            }
                        }
                    }
                    if (response.ok) {
                        return { success: true, id: createdId, isNew: !itemExists };
                    }
                    else {
                        const errorText = await response.text();
                        console.error(`❌ Failed in ${listName}:`, errorText);
                        return { success: false, id: null, error: errorText };
                    }
                }
                catch (error) {
                    console.error(`Error in ${listName}:`, error);
                    return { success: false, id: null, error: error };
                }
            };
            const formatNumberInputs = () => {
                const numberInputs = document.querySelectorAll('input[type="number"]');
                numberInputs.forEach(input => {
                    input.addEventListener('blur', (e) => {
                        const target = e.target;
                        if (target.value) {
                            const num = parseFloat(target.value);
                            if (!isNaN(num)) {
                                target.value = num.toFixed(4);
                            }
                        }
                    });
                });
            };
            // Enhanced Excel Import with detailed progress
            importBtn.addEventListener('click', () => excelInput.click());
            excelInput.addEventListener('change', async (e) => {
                var _a;
                const file = (_a = e.target.files) === null || _a === void 0 ? void 0 : _a[0];
                if (!file)
                    return;
                // Create processing overlay
                const processingOverlay = createLoadingOverlay("Processing Excel file...");
                document.body.appendChild(processingOverlay);
                try {
                    const data = await file.arrayBuffer();
                    const workbook = XLSX.read(data);
                    const sheetName = workbook.SheetNames[0];
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    // Reset counters
                    let successCount = 0;
                    let newEntriesInOtherLists = 0;
                    let errorCount = 0;
                    const totalCount = jsonData.length;
                    const formatToFourDecimals = (value) => {
                        if (value === undefined || value === null || value === "")
                            return 0;
                        const num = parseFloat(value);
                        if (isNaN(num))
                            return 0;
                        return parseFloat(num.toFixed(4));
                    };
                    // Process each row with progress updates
                    for (let i = 0; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        updateLoadingMessage(processingOverlay, `Processing row ${i + 1} of ${totalCount}...`);
                        try {
                            const excelProjectRef = row["Project Reference"] || row["A"] || '';
                            if (!excelProjectRef) {
                                console.warn("Skipping row: No project reference found");
                                errorCount++;
                                continue;
                            }
                            // 1. Find project in Project List
                            let projectName = '';
                            let projectExists = false;
                            const projectRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(projectListName)}')/items?$select=Id,Title,ProjectReference,PC,CustomerName,GroupHead&$filter=ProjectReference eq '${encodeURIComponent(excelProjectRef)}'&$top=1`, {
                                headers: {
                                    'Accept': 'application/json;odata=verbose',
                                    'Content-Type': 'application/json;odata=verbose'
                                }
                            });
                            if (projectRes.ok) {
                                const projectData = await projectRes.json();
                                if (projectData.d.results.length > 0) {
                                    const project = projectData.d.results[0];
                                    if (project.Title && project.Title !== project.ProjectReference && project.Title !== excelProjectRef) {
                                        projectName = project.Title;
                                    }
                                    else if (project.CustomerName && project.CustomerName !== 'N/A') {
                                        projectName = project.CustomerName;
                                    }
                                    else {
                                        projectName = excelProjectRef;
                                    }
                                    projectExists = true;
                                }
                                else {
                                    errorCount++;
                                    continue;
                                }
                            }
                            else {
                                errorCount++;
                                continue;
                            }
                            if (projectExists && projectName) {
                                const combinedTitle = `${excelProjectRef} - ${projectName}`;
                                // 2. Process Collection List
                                const collectionData = {
                                    Title: combinedTitle,
                                    project: projectName
                                };
                                if (row["NormalDue"] !== undefined && row["NormalDue"] !== null && row["NormalDue"] !== "") {
                                    collectionData.NormalDue = formatToFourDecimals(row["NormalDue"]);
                                }
                                if (row["NormalOverdue"] !== undefined && row["NormalOverdue"] !== null && row["NormalOverdue"] !== "") {
                                    collectionData.NormalOverdue = formatToFourDecimals(row["NormalOverdue"]);
                                }
                                if (row["RetentionDue"] !== undefined && row["RetentionDue"] !== null && row["RetentionDue"] !== "") {
                                    collectionData.RetentionDue = formatToFourDecimals(row["RetentionDue"]);
                                }
                                if (row["RetentionOverdue"] !== undefined && row["RetentionOverdue"] !== null && row["RetentionOverdue"] !== "") {
                                    collectionData.RetentionOverdue = formatToFourDecimals(row["RetentionOverdue"]);
                                }
                                if (row["GstTds"] !== undefined && row["GstTds"] !== null && row["GstTds"] !== "") {
                                    collectionData.GstTds = formatToFourDecimals(row["GstTds"]);
                                }
                                if (row["WHT"] !== undefined && row["WHT"] !== null && row["WHT"] !== "") {
                                    collectionData.WHT = formatToFourDecimals(row["WHT"]);
                                }
                                if (row["ReceivableValue"] !== undefined && row["ReceivableValue"] !== null && row["ReceivableValue"] !== "") {
                                    collectionData.ReceivableValue = formatToFourDecimals(row["ReceivableValue"]);
                                }
                                const collectionResult = await upsertListItem(collectionListName, collectionData, combinedTitle);
                                if (collectionResult.success) {
                                    if (collectionResult.isNew)
                                        newEntriesInOtherLists++;
                                }
                                // 3. Process other lists
                                const processOtherList = async (listName, columnName, value) => {
                                    if (value !== undefined && value !== null && value !== "") {
                                        const numericValue = formatToFourDecimals(value);
                                        if (numericValue !== 0) {
                                            const listData = {
                                                Title: combinedTitle,
                                                project: projectName,
                                                OBL: numericValue
                                            };
                                            const result = await upsertListItem(listName, listData, combinedTitle);
                                            if (result.success && result.isNew) {
                                                newEntriesInOtherLists++;
                                                return true;
                                            }
                                        }
                                    }
                                    return false;
                                };
                                await processOtherList(revenueListName, "Revenue Forecast OBL", row["Revenue Forecast OBL"]);
                                await processOtherList(invoicingListName, "Invoicing Planning OBL", row["Invoicing Planning OBL"]);
                                await processOtherList(seibiesListName, "SEIBIES OBL", row["SEIBIES OBL"]);
                                await processOtherList(inventoryListName, "Inventory OBL", row["Inventory OBL"]);
                                successCount++;
                            }
                            else {
                                errorCount++;
                            }
                        }
                        catch (error) {
                            console.error("Error processing row:", row, error);
                            errorCount++;
                        }
                    }
                    // Update summary UI
                    if (successCountEl)
                        successCountEl.textContent = successCount.toString();
                    if (errorCountEl)
                        errorCountEl.textContent = errorCount.toString();
                    if (totalCountEl)
                        totalCountEl.textContent = totalCount.toString();
                    let newEntriesEl = document.querySelector('#newEntriesCount');
                    if (!newEntriesEl) {
                        const summaryGrid = document.querySelector('.summary-grid');
                        if (summaryGrid) {
                            const newEntriesDiv = document.createElement('div');
                            newEntriesDiv.className = 'summary-item info';
                            newEntriesDiv.innerHTML = `
                <span class="summary-label">New Entries</span>
                <span class="summary-value" id="newEntriesCount">0</span>
              `;
                            summaryGrid.appendChild(newEntriesDiv);
                            newEntriesEl = document.querySelector('#newEntriesCount');
                        }
                    }
                    if (newEntriesEl) {
                        newEntriesEl.textContent = newEntriesInOtherLists.toString();
                    }
                    if (summaryDiv)
                        summaryDiv.style.display = 'block';
                    // Remove processing overlay
                    removeLoadingOverlay();
                    // Show success toast
                    if (successCount > 0) {
                        showToast(`✅ Excel import completed! Successfully imported ${successCount} of ${totalCount} projects.`, "success", 5000);
                    }
                    else {
                        showToast(`⚠️ Import completed with ${errorCount} errors. Check that projects exist in Project List.`, "warning", 5000);
                    }
                    // Refresh data
                    await fetchMasterData();
                }
                catch (error) {
                    console.error("Import error:", error);
                    removeLoadingOverlay();
                    showToast("❌ Failed to process Excel file. Please try again.", "error");
                }
                excelInput.value = "";
            });
            // Download template
            downloadTemplateBtn.addEventListener('click', () => {
                const templateData = [
                    {
                        "Project Reference": "REF-001",
                        "NormalDue": 100000,
                        "NormalOverdue": 50000,
                        "RetentionDue": 20000,
                        "RetentionOverdue": 10000,
                        "GstTds": 18000,
                        "WHT": 10000,
                        "ReceivableValue": 150000,
                        "Revenue Forecast OBL": 500000,
                        "Invoicing Planning OBL": 300000,
                        "SEIBIES OBL": 200000,
                        "Inventory OBL": 100000
                    }
                ];
                const ws = XLSX.utils.json_to_sheet(templateData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Template");
                const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
                saveAs(new Blob([buf], { type: "application/octet-stream" }), "Master_Data_Template.xlsx");
                showToast("📥 Template downloaded successfully", "success");
            });
            // Edit form submission with loading
            editForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const formatToFourDecimals = (value) => {
                    if (!value || value.trim() === '')
                        return 0;
                    const num = parseFloat(value);
                    if (isNaN(num))
                        return 0;
                    return parseFloat(num.toFixed(4));
                };
                const formData = {
                    projectReference: document.querySelector('#editProjectReference').value.trim(),
                    projectName: document.querySelector('#editProjectName').value.trim(),
                    normalDue: formatToFourDecimals(document.querySelector('#editNormalDue').value).toString(),
                    normalOverdue: formatToFourDecimals(document.querySelector('#editNormalOverdue').value).toString(),
                    retentionDue: formatToFourDecimals(document.querySelector('#editRetentionDue').value).toString(),
                    retentionOverdue: formatToFourDecimals(document.querySelector('#editRetentionOverdue').value).toString(),
                    gstTds: formatToFourDecimals(document.querySelector('#editGstTds').value).toString(),
                    wht: formatToFourDecimals(document.querySelector('#editWHT').value).toString(),
                    receivableValue: formatToFourDecimals(document.querySelector('#editReceivableValue').value).toString(),
                    revenueForecast: formatToFourDecimals(document.querySelector('#editRevenueForecast').value).toString(),
                    invoicingPlanning: formatToFourDecimals(document.querySelector('#editInvoicingPlanning').value).toString(),
                    seibies: formatToFourDecimals(document.querySelector('#editSEIBIES').value).toString(),
                    inventory: formatToFourDecimals(document.querySelector('#editInventory').value).toString()
                };
                // Show loading overlay
                const updatingOverlay = createLoadingOverlay("Updating project data...");
                document.body.appendChild(updatingOverlay);
                try {
                    const digest = await getDigest();
                    const projectId = currentEditingProject === null || currentEditingProject === void 0 ? void 0 : currentEditingProject.id;
                    if (!projectId) {
                        throw new Error("No project ID found");
                    }
                    // 1. Update Project List
                    if (formData.projectName !== (currentEditingProject === null || currentEditingProject === void 0 ? void 0 : currentEditingProject.projectName)) {
                        try {
                            await fetch(`${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(projectListName)}')/items(${projectId})`, {
                                method: "POST",
                                headers: {
                                    "Accept": "application/json;odata=verbose",
                                    "Content-Type": "application/json;odata=verbose",
                                    "X-RequestDigest": digest,
                                    "X-HTTP-Method": "MERGE",
                                    "IF-MATCH": "*"
                                },
                                body: JSON.stringify({
                                    __metadata: { type: "SP.Data.Project_x0020_ListListItem" },
                                    Title: formData.projectName,
                                    ProjectReference: formData.projectReference
                                })
                            });
                        }
                        catch (error) {
                            console.error("Error updating Project List:", error);
                        }
                    }
                    // 2. Update Collection List
                    const fields = await getListFieldsWithNames(collectionListName);
                    const combinedTitle = `${formData.projectReference} - ${formData.projectName}`;
                    const collectionUpdateData = {
                        Title: combinedTitle
                    };
                    // Map fields
                    const mapField = (excelName, value) => {
                        const fieldName = findActualFieldName(fields, excelName);
                        if (fieldName) {
                            collectionUpdateData[fieldName] = parseFloat(value);
                        }
                    };
                    mapField('NormalDue', formData.normalDue);
                    mapField('NormalOverdue', formData.normalOverdue);
                    mapField('RetentionDue', formData.retentionDue);
                    mapField('RetentionOverdue', formData.retentionOverdue);
                    mapField('GstTds', formData.gstTds);
                    mapField('WHT', formData.wht);
                    mapField('ReceivableValue', formData.receivableValue);
                    await upsertListItem(collectionListName, collectionUpdateData, combinedTitle);
                    // 3. Update other lists
                    updateLoadingMessage(updatingOverlay, "Updating other lists...");
                    const updateOtherList = async (listName, value) => {
                        const numValue = parseFloat(value);
                        if (!isNaN(numValue)) {
                            const otherListData = {
                                Title: combinedTitle,
                                OBL: numValue
                            };
                            await upsertListItem(listName, otherListData, combinedTitle);
                            return true;
                        }
                        return false;
                    };
                    await Promise.all([
                        updateOtherList(revenueListName, formData.revenueForecast),
                        updateOtherList(invoicingListName, formData.invoicingPlanning),
                        updateOtherList(seibiesListName, formData.seibies),
                        updateOtherList(inventoryListName, formData.inventory)
                    ]);
                    // Remove loading overlay
                    removeLoadingOverlay();
                    // Show success toast
                    showToast("✅ Project data updated successfully", "success");
                    // Refresh the master data
                    await fetchMasterData();
                }
                catch (error) {
                    console.error("Error updating project data:", error);
                    removeLoadingOverlay();
                    showToast("❌ Failed to update project data. Please try again.", "error");
                }
                // Return to table view
                if (editSection)
                    editSection.style.display = "none";
                if (tableSection)
                    tableSection.style.display = "block";
                currentEditingProject = null;
                editForm.reset();
            });
            // Back button
            backBtn.addEventListener('click', () => {
                if (editSection)
                    editSection.style.display = "none";
                if (tableSection)
                    tableSection.style.display = "block";
                currentEditingProject = null;
                editForm.reset();
            });
            // Search functionality
            searchInput.addEventListener('input', () => {
                currentPage = 1;
                const query = searchInput.value.toLowerCase();
                const filteredData = masterData.filter(p => (p.projectReference || '').toLowerCase().includes(query) ||
                    (p.projectName || '').toLowerCase().includes(query));
                renderTable(filteredData);
            });
            resetSearchBtn.addEventListener('click', () => {
                searchInput.value = '';
                currentPage = 1;
                renderTable(masterData);
            });
            const renderPagination = (totalPages) => {
                if (!paginationDiv) {
                    console.warn("Pagination div not found");
                    return;
                }
                paginationDiv.innerHTML = '';
                for (let i = 1; i <= totalPages; i++) {
                    const btn = document.createElement('button');
                    btn.textContent = i.toString();
                    btn.className = i === currentPage ? 'active' : '';
                    btn.addEventListener('click', () => {
                        currentPage = i;
                        const query = searchInput.value.toLowerCase();
                        const filteredData = masterData.filter(p => (p.projectReference || '').toLowerCase().includes(query) ||
                            (p.projectName || '').toLowerCase().includes(query));
                        renderTable(filteredData);
                    });
                    paginationDiv.appendChild(btn);
                }
            };
            // Initialize with loading
            setTimeout(() => {
                fetchMasterData();
            }, 300);
        });
    }
};
//# sourceMappingURL=master.js.map
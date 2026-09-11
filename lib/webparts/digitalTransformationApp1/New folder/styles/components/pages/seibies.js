import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/seibies.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canViewForm, getAssignedProjects, isPMManager, isEngineeringUser, 
// getAssignedProjectValues,
updateUIBasedOnPermissions } from "./auth";
let currentSortConfig = {
    column: '',
    direction: 'asc'
};
export default {
    content: `
<div class="seibies-page fade-in">
  <section class="table-section-seibies card">
    <div class="page-header-seibies">
      <h1>SEI-BIES Liquidation</h1>
    </div>
    
    <div class="table-actions-seibies">
      <div class="table-actions-row-seibies">
        <div class="tableColp4-seibies">
          <label for="filterProjectSeibies">Project</label>
          <select id="filterProjectSeibies"><option value="">All Projects</option></select>
        </div>
        <div class="tableColp4-seibies">
          <label for="filterCustomerSeibies">Customer</label>
          <select id="filterCustomerSeibies"><option value="">All Customers</option></select>
        </div>
        <div class="tableColp4-seibies">
          <label for="filterProjectManagerSeibies">Project Manager</label>
          <select id="filterProjectManagerSeibies"><option value="">All Managers</option></select>
        </div>
        <div class="tableColp4-seibies">
          <label for="filterGroupHeadSeibies">Group Head</label>
          <select id="filterGroupHeadSeibies"><option value="">All Group Heads</option></select>
        </div>
        <div class="tableColp4-seibies">
          <label for="filterProjectReferenceSeibies">Project Reference</label>
          <select id="filterProjectReferenceSeibies"><option value="">All References</option></select>
        </div>
        <div class="tableColp4-seibies"> 
          <label>&nbsp;</label>
          <button class="btn-secondary-seibies clear-filters-btn-seibies" title="Clear All Filters">Reset Filter</button> 
        </div>
      </div>
      
      <div class="table-actions-row-seibies">
        <div class="tableColp4-seibies">
          <label for="filterYearSeibies">Year</label>
          <select id="filterYearSeibies"><option value="">All Years</option></select>
        </div>
        <div class="tableColp4-seibies">
          <label for="filterMonthSeibies">Month</label>
          <select id="filterMonthSeibies"><option value="">All Months</option></select>
        </div>
        <div class="tableColp4-seibies"> 
          <label>&nbsp;</label>
          <button class="btn-secondary-seibies export-btn-seibies">📊 Export Excel</button>
        </div> 
      </div>
    </div>
       
    <div class="table-container-seibies">
      <div style="overflow-x:auto;">
        <table class="month-columns-table-seibies">
          <thead id="totalsHeaderSeibies"></thead>
          <thead id="monthColumnsHeaderSeibies"></thead>
          <tbody class="seibies-table-body"></tbody>
        </table>
      </div>
      
      <div class="pagination-container-seibies">
        <div class="pagination-row-seibies">
          <div class="pagination-left-seibies">
            <button class="pagination-btn-seibies" id="firstPageSeibies" disabled>First</button>
            <button class="pagination-btn-seibies" id="prevPageSeibies" disabled>Previous</button>
            
            <div class="pagination-info-seibies">
              Page <span id="currentPageSeibies">1</span> of <span id="totalPagesSeibies">1</span>
            </div>
            
            <div class="pagination-pages-seibies" id="pageNumbersSeibies"></div>
            
            <button class="pagination-btn-seibies" id="nextPageSeibies" disabled>Next</button>
            <button class="pagination-btn-seibies" id="lastPageSeibies" disabled>Last</button>
          </div>
        </div>
        
        <div class="pagination-row-seibies pagination-bottom-row-seibies">
          <div class="pagination-right-seibies">
            <div class="page-size-selector-seibies">
              <label for="pageSizeSelectSeibies">Show:</label>
              <select id="pageSizeSelectSeibies" class="page-size-select-seibies">
                <option value="10">10</option>
                <option value="25">25</option>
                <option value="50">50</option>
                <option value="100">100</option>
              </select>
              <span>per page</span>
            </div>
            
            <div class="pagination-info-seibies">
              Showing <span id="startItemSeibies">0</span>-<span id="endItemSeibies">0</span> of <span id="totalItemsSeibies">0</span> records
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>

  <section class="form-section-seibies card" style="display:none;">
    <div class="page-header-seibies">
      <h1 id="formTitleSeibies">Add / Edit SEIBIES Forecast</h1>
    </div>
    <form id="seibiesForm">
      
      <!-- Project Selection (hidden in edit mode) -->
      <div class="form-row-seibies hidden-section">
        <div class="form-group-seibies project-select-container-seibies">
          <label>Project</label>
          <select id="projectSelectSeibies" required></select>
        </div>
      </div>

      <!-- MAIN WRAPPER - 2 Column Layout (Matching Revenue) -->
      <div class="project-main-wrapper-seibies">

        <!-- LEFT SIDE - Project Details (6 Rows) -->
        <div class="project-left-side-seibies" id="projectDetailsSeibies">
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Project:</span>
            <span class="detail-value-seibies" id="projectNameDisplaySeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Project Ref:</span>
            <span class="detail-value-seibies" id="projectRefSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Profit Center:</span>
            <span class="detail-value-seibies" id="profitCenterSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Project Manager:</span>
            <span class="detail-value-seibies" id="projectManagerSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Group Head:</span>
            <span class="detail-value-seibies" id="groupHeadSeibies">-</span>
          </div>
          <div class="detail-item-seibies">
            <span class="detail-label-seibies">Customer:</span>
            <span class="detail-value-seibies" id="customerSeibies">-</span>
          </div>
        </div>

        <!-- RIGHT SIDE - OBL Section (No Remarks) -->
        <div class="project-right-side-seibies">
          <!-- OBL Section -->
          <div class="obl-section-container-seibies hidden-section" id="oblContainerSeibies">
            <div class="obl-content-seibies">
              <div class="obl-left-seibies">
                <div class="obl-label-group-seibies">
                  <label>SEIBIES </label>
                </div>
                <div class="obl-value-display-seibies">
                  <input type="text" id="oblLastMonthSeibies" class="obl-input-seibies number-input-seibies" placeholder="Enter SEIBIES" readonly />
                </div>
              </div>
              <div class="obl-right-seibies">
                <div class="validation-summary-seibies">
                  <div class="validation-item-seibies">
                    <span class="validation-label-seibies">Total Current Value:</span>
                    <span class="validation-value-seibies" id="totalCurrentValueSeibies">0</span>
                  </div>
                  <div class="validation-item-seibies">
                    <span class="validation-label-seibies">Remaining Balance:</span>
                    <span class="validation-value-seibies" id="remainingValueSeibies">0</span>
                  </div>
                  <div class="validation-line-seibies">
                    <div class="line-fill-seibies" id="oblLineFillSeibies"></div>
                  </div>
                </div>
              </div>
            </div>
            <div class="validation-message-seibies" id="oblValidationMessageSeibies">
              ✅ All values are within OBL limit
            </div>
          </div>
        </div>

      </div>

      <!-- Forecast Sections -->
      <div class="forecast-section-seibies hidden-section" id="historicalForecastSectionSeibies">
        <div class="historical-toggle-header-seibies">
          <h3>📊 Historical Data (Last 3 Months - Readonly)</h3>
          <button type="button" id="toggleHistoricalBtnSeibies" class="btn-secondary-seibies">Show</button>
        </div>
        <div id="historicalContainerSeibies" class="historical-container-seibies" style="display:none;"></div>
      </div>
      
      <div class="forecast-section-seibies hidden-section" id="detailedForecastSectionSeibies">
        <h3>Detailed Forecast (Current + Next 2 Months)</h3>
        <div id="firstRowContainerSeibies" class="first-row-container-seibies"></div>
      </div>
      
      <div class="forecast-section-seibies hidden-section" id="longtermForecastSectionSeibies">
        <h3>Long-term Forecast (Next 15 Months)</h3>
        <div id="remainingContainerSeibies" class="remaining-container-seibies"></div>
      </div>

      <div class="forecast-section-seibies hidden-section" id="yearlyForecastSectionSeibies">
        <h3>5-Year Forecast</h3>
        <div id="yearlyContainerSeibies" class="yearly-container-seibies"></div>
      </div>
      
      <div class="form-actions-seibies">
        <button type="submit" class="btn-primary-seibies">💾 Save SEIBIES Forecast</button>
        <button type="button" class="btn-secondary-seibies back-to-table-seibies">⬅ Back</button>
      </div>
    </form>
  </section>
</div>
  `,
    style: `
  .role-indicator-seibies {
    display: none;
  }

  /* ============================================ */
  /* PAGINATION - IMPROVED STYLING               */
  /* ============================================ */

  .pagination-container-seibies {
    display: flex;
    flex-direction: column;
    margin-top: 20px;
    padding: 16px 20px;
    background-color: #f8f9fa;
    border-radius: 8px;
    border: 1px solid #dee2e6;
    gap: 12px;
  }

  .pagination-row-seibies {
    display: flex;
    justify-content: center;
    align-items: center;
    width: 100%;
  }

  .pagination-bottom-row-seibies {
    border-top: 1px solid #e9ecef;
    padding-top: 12px;
    margin-top: 4px;
  }

  .pagination-left-seibies {
    display: flex;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
    justify-content: center;
  }

  .pagination-right-seibies {
    display: flex;
    align-items: center;
    gap: 20px;
    flex-wrap: wrap;
    justify-content: center;
    width: 100%;
  }

  .pagination-btn-seibies {
    padding: 6px 14px;
    background-color: #007bff;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-size: 13px;
    transition: background-color 0.2s, opacity 0.2s;
    white-space: nowrap;
    min-width: 60px;
  }

  .pagination-btn-seibies:hover:not(:disabled) {
    background-color: #0056b3;
  }

  .pagination-btn-seibies:disabled {
    background-color: #6c757d;
    cursor: not-allowed;
    opacity: 0.6;
  }

  .pagination-info-seibies {
    font-size: 14px;
    color: #495057;
    white-space: nowrap;
    padding: 0 8px;
  }

  .pagination-pages-seibies {
    display: flex;
    gap: 4px;
  }

  .page-number-seibies {
    padding: 5px 12px;
    background-color: white;
    border: 1px solid #dee2e6;
    border-radius: 4px;
    cursor: pointer;
    font-size: 13px;
    transition: all 0.2s;
    min-width: 34px;
    text-align: center;
  }

  .page-number-seibies:hover:not(.active-seibies):not(.dots-seibies) {
    background-color: #f8f9fa;
    border-color: #007bff;
  }

  .page-number-seibies.active-seibies {
    background-color: #007bff;
    color: white;
    border-color: #007bff;
  }

  .page-number-seibies.dots-seibies {
    background-color: transparent;
    border: none;
    cursor: default;
    color: #6c757d;
    padding: 5px 8px;
  }

  .page-size-selector-seibies {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 13px;
    color: #495057;
  }

  .page-size-selector-seibies label {
    margin: 0;
    font-weight: 500;
  }

  .page-size-select-seibies {
    padding: 4px 8px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    background-color: white;
    font-size: 13px;
    cursor: pointer;
    min-width: 55px;
  }

  .page-size-select-seibies:focus {
    outline: none;
    border-color: #007bff;
    box-shadow: 0 0 0 2px rgba(0, 123, 255, 0.25);
  }

  .pagination-right-seibies .pagination-info-seibies {
    font-size: 13px;
    color: #6c757d;
  }

  @media (max-width: 768px) {
    .pagination-container-seibies {
      padding: 12px;
      gap: 10px;
    }
    
    .pagination-left-seibies {
      gap: 4px;
    }
    
    .pagination-btn-seibies {
      padding: 5px 10px;
      font-size: 12px;
      min-width: 50px;
    }
    
    .page-number-seibies {
      padding: 4px 9px;
      font-size: 12px;
      min-width: 30px;
    }
    
    .pagination-info-seibies {
      font-size: 12px;
      padding: 0 4px;
    }
    
    .page-size-selector-seibies {
      font-size: 12px;
    }
    
    .pagination-right-seibies {
      gap: 12px;
    }
  }

  @media (max-width: 480px) {
    .pagination-left-seibies {
      gap: 3px;
    }
    
    .pagination-btn-seibies {
      padding: 4px 8px;
      font-size: 11px;
      min-width: 40px;
    }
    
    .page-number-seibies {
      padding: 3px 7px;
      font-size: 11px;
      min-width: 26px;
    }
    
    .pagination-info-seibies {
      font-size: 11px;
    }
  }

  /* ============================================ */
  /* HISTORICAL TOGGLE HEADER - MATCHING REVENUE  */
  /* ============================================ */

  .historical-toggle-header-seibies {
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

  .historical-toggle-header-seibies:hover {
    background: #e9ecef !important;
  }

  .historical-toggle-header-seibies h3 {
    margin: 0 !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    color: #495057 !important;
  }

  #toggleHistoricalBtnSeibies {
    padding: 4px 16px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    background: #28a745 !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    cursor: pointer !important;
    transition: background 0.2s !important;
    flex-shrink: 0 !important;
  }

  #toggleHistoricalBtnSeibies:hover {
    opacity: 0.85 !important;
  }

  .historical-container-seibies {
    transition: all 0.3s ease !important;
  }

  /* ============================================ */
  /* PROJECT MAIN WRAPPER - 2 COLUMN LAYOUT       */
  /* ============================================ */

  .project-main-wrapper-seibies {
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
  .project-left-side-seibies {
    padding: 8px 16px !important;
    background: #ffffff !important;
    border-right: 1px solid #e9ecef !important;
    display: flex !important;
    flex-direction: column !important;
    gap: 0 !important;
    justify-content: flex-start !important;
    align-items: flex-start !important;
  }

  .project-left-side-seibies .detail-item-seibies {
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

  .project-left-side-seibies .detail-label-seibies {
    font-weight: 700 !important;
    white-space: nowrap !important;
    min-width: 100px !important;
  }

  .project-left-side-seibies .detail-value-seibies {
    white-space: nowrap !important;
  }

  /* RIGHT SIDE */
  .project-right-side-seibies {
    padding: 8px 16px !important;
    background: #ffffff !important;
    display: flex !important;
    flex-direction: column !important;
    gap: 6px !important;
    justify-content: flex-start !important;
  }

  /* ============================================ */
  /* OBL SECTION - MATCHING REVENUE               */
  /* ============================================ */

  .obl-section-container-seibies {
    margin: 2px 0 0 0 !important;
    padding: 6px 12px !important;
    background: #f8f9fa !important;
    border-radius: 4px !important;
    border: 1px solid #e9ecef !important;
  }

  .obl-content-seibies {
    display: grid !important;
    grid-template-columns: 1fr 1.5fr !important;
    gap: 12px !important;
    align-items: start !important;
  }

  .obl-left-seibies {
    display: flex !important;
    flex-direction: row !important;
  }

  .obl-label-group-seibies {
    display: flex !important;
    align-items: center !important;
    gap: 4px !important;
    margin-bottom: 2px !important;
    text-align: left !important;
  }

  .obl-label-group-seibies label {
    font-size: 12px !important;
    font-weight: 700 !important;
    color: #495057 !important;
    margin: 0 !important;
  }

  .info-tooltip-seibies {
    font-size: 12px !important;
    color: #007bff !important;
    cursor: help !important;
  }

  .obl-value-display-seibies {
    display: flex !important;
    flex-direction: column !important;
  }

  .obl-input-seibies {
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

  .field-note-seibies {
    font-size: 9px !important;
    color: #6c757d !important;
    margin-top: 2px !important;
    font-style: italic !important;
    text-align: left !important;
  }

  .obl-right-seibies {
    display: flex !important;
    align-items: center !important;
  }

  .validation-summary-seibies {
    background: #fff !important;
    padding: 4px 12px !important;
    border-radius: 4px !important;
    border: 1px solid #e9ecef !important;
    width: 100% !important;
  }

  .validation-item-seibies {
    display: flex !important;
    justify-content: space-between !important;
    margin-bottom: 1px !important;
    font-size: 11px !important;
    text-align: left !important;
  }

  .validation-label-seibies {
    font-weight: 600 !important;
    color: #666 !important;
    font-size: 10px !important;
    text-align: left !important;
  }

  .validation-value-seibies {
    font-weight: 700 !important;
    font-size: 11px !important;
    color: #333 !important;
    text-align: right !important;
  }

  .validation-line-seibies {
    height: 4px !important;
    background: #e9ecef !important;
    border-radius: 2px !important;
    margin-top: 3px !important;
    overflow: hidden !important;
  }

  .line-fill-seibies {
    height: 100% !important;
    background: #28a745 !important;
    border-radius: 2px !important;
    transition: all 0.3s ease !important;
    width: 0% !important;
  }

  .line-fill-seibies.warning-seibies {
    background: #ffc107 !important;
  }

  .line-fill-seibies.danger-seibies {
    background: #dc3545 !important;
  }

  .validation-message-seibies {
    display: block !important;
    padding: 8px 12px;
    border-radius: 4px;
    margin-top: 5px;
    font-size: 12px;
    font-weight: 500;
  }

  .validation-message-seibies.danger-seibies {
    background-color: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
  }

  .validation-message-seibies.warning-seibies {
    background-color: #fff3cd;
    color: #856404;
    border: 1px solid #ffc107;
  }

  /* ============================================ */
  /* RESPONSIVE                                   */
  /* ============================================ */

  @media (max-width: 992px) {
    .project-main-wrapper-seibies {
      grid-template-columns: 1fr !important;
    }
    
    .project-left-side-seibies {
      border-right: none !important;
      border-bottom: 1px solid #e9ecef !important;
    }
    
    .project-right-side-seibies {
      padding: 8px 14px !important;
    }
  }

  @media (max-width: 768px) {
    .obl-content-seibies {
      grid-template-columns: 1fr !important;
      gap: 8px !important;
    }
    
    .obl-input-seibies {
      max-width: 100% !important;
    }
    
    .project-left-side-seibies .detail-label-seibies {
      min-width: 80px !important;
    }
  }

  /* ============================================ */
  /* MAIN TABLE - MONTH COLUMNS - MATCHING REVENUE */
  /* ============================================ */

  .month-columns-table-seibies {
    border-collapse: separate;
    border-spacing: 0;
    width: 100%;
    min-width: 100%;
  }

  .month-columns-table-seibies th,
  .month-columns-table-seibies td {
    border: 1px solid #68e0f0 !important;
    padding: 6px 4px;
    text-align: center;
    white-space: normal !important;
    word-wrap: break-word !important;
    word-break: break-word !important;
    overflow-wrap: break-word !important;
  }

  .month-columns-table-seibies thead th {
    background: #68e0f0 !important;
    font-weight: bold;
    border-bottom: 2px solid #007bff !important;
  }

  /* ============================================ */
  /* NON-CURRENT MONTH - WHITE HEADER BORDERS     */
  /* ============================================ */

  .month-columns-table-seibies .non-current-month-header {
    border-left: 1px solid #ffffff !important;
    border-right: 1px solid #ffffff !important;
  }

  .month-columns-table-seibies .non-current-month-subheader {
    border-left: 1px solid #ffffff !important;
    border-right: 1px solid #ffffff !important;
  }

  /* ============================================ */
  /* NON-CURRENT MONTH - CYAN DATA BORDERS        */
  /* ============================================ */

  .month-columns-table-seibies .non-current-month-cell {
    border-left: 1px solid #68e0f0 !important;
    border-right: 1px solid #68e0f0 !important;
  }

  .totals-header-table-seibies .non-current-month-cell {
    border-left: 1px solid #68e0f0 !important;
    border-right: 1px solid #68e0f0 !important;
  }

  /* ----- STICKY COLUMNS - NO FREEZE COLUMN ----- */

  /* Project Reference - First Sticky Column - NOW CLICKABLE */
  .project-ref-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 0 !important;
    min-width: 120px !important;
    max-width: 120px !important;
    width: 120px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .project-ref-col .edit-trigger {
    color: #0078d4 !important;
    text-decoration: underline !important;
    cursor: pointer !important;
  }

  .project-ref-col .edit-trigger:hover {
    color: #0056b3 !important;
  }

  .month-columns-table-seibies thead .project-ref-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  /* Project Column - Second Sticky Column - CLICKABLE */
  .project-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 120px !important;
    min-width: 180px !important;
    max-width: 180px !important;
    width: 180px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .project-col .edit-trigger {
    color: #0078d4 !important;
    text-decoration: underline !important;
    cursor: pointer !important;
  }

  .project-col .edit-trigger:hover {
    color: #0056b3 !important;
  }

  .month-columns-table-seibies thead .project-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  /* Project Manager - Third Sticky Column */
  .pm-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 300px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-seibies thead .pm-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  /* Group Head - Fourth Sticky Column */
  .gh-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 400px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-seibies thead .gh-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  /* Customer - Fifth Sticky Column */
  .customer-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 500px !important;
    min-width: 130px !important;
    max-width: 130px !important;
    width: 130px !important;
    text-align: left !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-seibies thead .customer-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  /* SEI-BIES (OBL) - Sixth Sticky Column */
  .obl-col {
    position: sticky !important;
    z-index: 2 !important;
    background: white !important;
    left: 630px !important;
    min-width: 90px !important;
    max-width: 90px !important;
    width: 90px !important;
    text-align: right !important;
    padding: 4px 6px !important;
    font-size: 11px !important;
  }

  .month-columns-table-seibies thead .obl-col {
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  /* Month Scroll Columns */
  .month-scroll-col {
    min-width: 60px !important;
    width: 60px !important;
  }

  /* ----- ZEBRA STRIPING & ROW HOVER - REVENUE STYLES ----- */
  .month-columns-table-seibies tbody tr:nth-child(even) {
    background-color: #f8f9fa !important;
  }

  .month-columns-table-seibies tbody tr:hover {
    background-color: #e3f2fd !important;
  }

  /* ----- TOTALS HEADER TABLE - REVENUE STYLES ----- */
  .totals-header-table-seibies {
    margin-bottom: 10px;
    background-color: #f8f9fa;
    border-collapse: separate;
    border-spacing: 0;
    table-layout: fixed !important;
    width: 100% !important;
  }

  .totals-header-table-seibies th,
  .totals-header-table-seibies td {
    width: auto !important;
    min-width: 0 !important;
    padding: 6px 4px !important;
    font-size: 12px;
    font-weight: bold;
    text-align: center;
    vertical-align: middle;
    border: 1px solid #dee2e6;
    white-space: normal !important;
    word-wrap: break-word !important;
    word-break: break-word !important;
    overflow-wrap: break-word !important;
  }

  .totals-header-table-seibies .totals-label-seibies {
    font-weight: 600;
    color: #495057;
    background-color: #f1f3f4 !important;
  }

  .totals-header-table-seibies .all-projects-total-seibies td {
    background-color: #d1ecf1 !important;
    color: #0c5460;
  }

  .totals-header-table-seibies .filtered-total-seibies td {
    background-color: #d4edda !important;
    color: #155724;
  }

  .totals-header-table-seibies .total-value-seibies {
    font-weight: bold;
    font-size: 13px;
  }

  /* Totals header sticky columns */
  .totals-header-table-seibies .project-ref-col {
    left: 0 !important;
    min-width: 120px !important;
    max-width: 120px !important;
    width: 120px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-seibies .project-col {
    left: 120px !important;
    min-width: 180px !important;
    max-width: 180px !important;
    width: 180px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-seibies .pm-col {
    left: 300px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-seibies .gh-col {
    left: 400px !important;
    min-width: 100px !important;
    max-width: 100px !important;
    width: 100px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-seibies .customer-col {
    left: 500px !important;
    min-width: 130px !important;
    max-width: 130px !important;
    width: 130px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-seibies .obl-col {
    left: 630px !important;
    min-width: 90px !important;
    max-width: 90px !important;
    width: 90px !important;
    position: sticky !important;
    z-index: 5 !important;
    background: #e9ecef !important;
  }

  .totals-header-table-seibies .month-scroll-col {
    min-width: 60px !important;
    width: 60px !important;
  }

  /* Totals Row Colors */
  tr.filtered-total-seibies {
    background-color: #d4edda !important;
  }
  tr.all-projects-total-seibies {
    background-color: #d1ecf1 !important;
  }

  /* ----- MONTH COLUMN HEADERS - REVENUE STYLES ----- */
  .month-column-header-seibies {
    background: #0078d4 !important;
    color: white !important;
    font-weight: 600;
    font-size: 11px;
    white-space: nowrap !important;
  }

  .month-subheader-seibies {
    background: #e3f2fd !important;
    color: #0078d4 !important;
    font-weight: 600;
    font-size: 10px;
    white-space: nowrap !important;
  }

  /* CURRENT MONTH HEADER - BLUE */
  .current-month-header-seibies {
    background: #2196f3 !important;
    color: white !important;
  }

  /* CURRENT MONTH SUB-HEADER - ORANGE */
  .current-month-header-table-seibies {
    background: #ff9800 !important;
    color: white !important;
  }

  /* ----- SORT HEADERS ----- */
  .sort-header-container-seibies {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
  }

  .sort-buttons-container-seibies {
    display: flex;
    flex-direction: column;
    gap: 1px;
  }

  .sort-btn-seibies {
    background: none;
    border: none;
    cursor: pointer;
    font-size: 8px;
    padding: 1px 2px;
    border-radius: 2px;
    color: #ccc;
    transition: all 0.2s ease;
  }

  .sort-btn-seibies:hover {
    background-color: rgba(255, 255, 255, 0.2);
    color: white;
  }

  .sort-btn-seibies.active-seibies {
    color: #ffeb3b;
  }

  .sort-btn-seibies.sort-asc-seibies.active-seibies {
    color: #4caf50;
  }

  .sort-btn-seibies.sort-desc-seibies.active-seibies {
    color: #f44336;
  }

  .header-label-seibies {
    margin-right: 4px;
  }

  /* ----- Clickable Edit Link ----- */
  .edit-trigger {
    color: #0078d4 !important;
    text-decoration: underline !important;
    cursor: pointer !important;
  }

  .edit-trigger:hover {
    color: #0056b3 !important;
  }

  /* ----- RESPONSIVE - STICKY COLUMN ADJUSTMENTS ----- */
  @media (max-width: 1400px) {
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
    .obl-col {
      left: 520px !important;
      min-width: 70px !important;
      max-width: 70px !important;
      width: 70px !important;
    }

    .totals-header-table-seibies .project-ref-col {
      left: 0 !important;
      min-width: 100px !important;
      max-width: 100px !important;
      width: 100px !important;
    }
    .totals-header-table-seibies .project-col {
      left: 100px !important;
      min-width: 150px !important;
      max-width: 150px !important;
      width: 150px !important;
    }
    .totals-header-table-seibies .pm-col {
      left: 250px !important;
      min-width: 80px !important;
      max-width: 80px !important;
      width: 80px !important;
    }
    .totals-header-table-seibies .gh-col {
      left: 330px !important;
      min-width: 80px !important;
      max-width: 80px !important;
      width: 80px !important;
    }
    .totals-header-table-seibies .customer-col {
      left: 410px !important;
      min-width: 110px !important;
      max-width: 110px !important;
      width: 110px !important;
    }
    .totals-header-table-seibies .obl-col {
      left: 520px !important;
      min-width: 70px !important;
      max-width: 70px !important;
      width: 70px !important;
    }

    .table-actions-row-seibies {
      grid-template-columns: repeat(3, 1fr);
    }
    .historical-container-seibies,
    .first-row-container-seibies { grid-template-columns: repeat(2, 1fr); }
    .remaining-row-seibies { grid-template-columns: repeat(3, 1fr); }
  }

  @media (max-width: 768px) {
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
    .obl-col {
      left: 370px !important;
      min-width: 55px !important;
      max-width: 55px !important;
      width: 55px !important;
    }

    .month-columns-table-seibies .month-scroll-col {
      min-width: 50px;
      width: 50px;
    }

    .table-container-seibies {
      overflow-x: auto;
    }
  }

  /* ----- SCROLLBAR STYLING ----- */
  .table-container-seibies::-webkit-scrollbar {
    width: 8px;
    height: 8px;
  }

  .table-container-seibies::-webkit-scrollbar-track {
    background: #f1f1f1;
  }

  .table-container-seibies::-webkit-scrollbar-thumb {
    background: #888;
    border-radius: 4px;
  }

  .table-container-seibies::-webkit-scrollbar-thumb:hover {
    background: #555;
  }

  /* ----- Core Layout & Components (existing) ----- */
  .seibies-page {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  }

  .hidden-section {
    display: none !important;
  }

  .visible-section-seibies {
    display: block !important;     
  }

  .forecast-section-seibies.visible-section-seibies {
    display: block;
  }

  .obl-section-container-seibies.visible-section-seibies {
    display: block;
  }

  .card { 
    background:#fff; 
    border-radius:10px; 
    padding:20px; 
    margin-bottom:20px; 
    box-shadow:0 2px 8px rgba(0,0,0,0.08);
  }

  .detailed-month-seibies {
    position: relative;
  }

  .current-month-indicator-seibies {
    position: absolute;
    top: 8px;
    right: 8px;
    background: #0078d4;
    color: white;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 10px;
    font-weight: bold;
  }

  .enhanced-inputs-container-seibies {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
    margin-bottom: 15px;
  }

  .enhanced-input-group-seibies {
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .enhanced-input-group-seibies.full-width-seibies {
    grid-column: 1 / -1;
  }

  .enhanced-input-group-seibies label {
    font-size: 11px;
    color: #666;
    margin-bottom: 4px;
    font-weight: 600;
    text-align: center;
  }

  .enhanced-input-group-seibies input,
  .enhanced-input-group-seibies select {
    width: 100%;
    padding: 6px 8px;
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    font-size: 12px;
    background: #fff;
    transition: all 0.2s ease;
    color: #000 !important;
  }

  .enhanced-input-group-seibies input:focus,
  .enhanced-input-group-seibies select:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
    outline: none;
  }

  .forecast-date-input-seibies {
    background: #fff !important;
    color: #333 !important;
    border: 1px solid #ddd !important;
  }

  .main-inputs-container-seibies {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 15px;
    margin-bottom: 15px;
  }

  .main-input-group-seibies {
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .main-input-group-seibies label {
    font-size: 12px;
    color: #666;
    margin-bottom: 6px;
    font-weight: 600;
  }

  .main-input-group-seibies input {
    width: 100%;
    padding: 8px 10px;
    color:#000 !important;
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    font-size: 14px;
    background: #fff;
    transition: all 0.2s ease;
  }

  .main-input-group-seibies input:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
    outline: none;
  }

  .month-box-seibies {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 15px;
    background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    position: relative;
  }

  .month-header-seibies {
    text-align: center;
    font-weight: 600;
    color: #0078d4;
    margin-bottom: 15px;
    padding-bottom: 8px;
    border-bottom: 1px solid #eee;
  }

  .historical-month-seibies {
    background: linear-gradient(135deg, #f5f5f5 0%, #e8e8e8 100%) !important;
    border: 1px solid #d0d0d0 !important;
    pointer-events: none !important;
    user-select: none !important;
    opacity: 0.8;
  }

  .historical-month-seibies * {
    pointer-events: none !important;
    user-select: none !important;
  }

  .historical-month-seibies::before {
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

  .historical-month-seibies .month-header-seibies {
    background: #6c757d !important;
    color: white !important;
    padding: 8px;
    margin: -15px -15px 15px -15px;
    border-radius: 8px 8px 0 0;
    border-bottom: 1px solid #ccc;
  }

  .historical-month-seibies input {
    background-color: #f8f9fa !important;
    color: #6c757d !important;
    border-color: #ced4da !important;
    cursor: not-allowed !important;
  }

  .items-section-seibies {
    margin-top: 15px;
    padding-top: 15px;
    border-top: 2px solid #e0e0e0;
  }

  .items-header-seibies {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 8px;
    font-size: 11px;
    color: #666;
    text-align: center;
    font-weight: 600;
    padding-bottom: 4px;
    border-bottom: 1px solid #eee;
  }

  .item-row-seibies {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 6px;
    align-items: center;
  }

  .item-row-seibies input {
    padding: 5px 8px;
    border: 1px solid #ddd;
    border-radius: 3px;
    font-size: 12px;
    background: #fff;
    transition: all 0.2s ease;
  }

  .item-row-seibies input[type="text"] { text-align: left; }
  .item-row-seibies input[type="number"] { text-align: right; }

  .item-row-seibies input:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 1px rgba(0,120,212,0.2);
    outline: none;
  }

  .historical-month-seibies .items-header-seibies {
    grid-template-columns: 2fr 1fr 1fr !important;
  }

  .historical-month-seibies .item-row-seibies {
    grid-template-columns: 2fr 1fr 1fr !important;
  }

  .historical-month-seibies .items-header-seibies div:nth-child(4),
  .historical-month-seibies .items-header-seibies div:nth-child(5),
  .historical-month-seibies .item-row-seibies input:nth-child(4),
  .historical-month-seibies .item-row-seibies input:nth-child(5) {
    display: none;
  }

  .historical-month-seibies .main-inputs-container-seibies {
    grid-template-columns: 1fr 1fr !important;
  }

  .historical-month-seibies .main-input-group-seibies:nth-child(3),
  .historical-month-seibies .main-input-group-seibies:nth-child(4) {
    display: none !important;
  }

  .historical-items-section-seibies .historical-items-header-seibies {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 8px;
    font-size: 11px;
    color: #666;
    text-align: center;
    font-weight: 600;
    padding-bottom: 4px;
    border-bottom: 1px solid #eee;
  }

  .historical-items-section-seibies .historical-item-row-seibies {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 6px;
    align-items: center;
  }

  .historical-items-section-seibies .historical-items-header-seibies div:nth-child(4),
  .historical-items-section-seibies .historical-items-header-seibies div:nth-child(5),
  .historical-items-section-seibies .historical-item-row-seibies input:nth-child(4),
  .historical-items-section-seibies .historical-item-row-seibies input:nth-child(5) {
    display: none;
  }

  .historical-month-seibies .main-inputs-container-seibies {
    grid-template-columns: 1fr 1fr;
  }

  .historical-month-seibies .main-input-group-seibies:nth-child(3),
  .historical-month-seibies .main-input-group-seibies:nth-child(4) {
    display: none;
  }

  .items-section-seibies {
    max-height: 400px;
    overflow-y: auto;
  }

  .item-row-seibies {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 6px;
    align-items: center;
  }

  .historical-month-seibies .item-row-seibies {
    grid-template-columns: 2fr 1fr 1fr;
  }

  .readonly-field-seibies {
    background-color: #f5f5f5 !important;
    color: #666 !important;
    border-color: #ddd !important;
  }

  .historical-month-seibies input,
  input[readonly] {
    background-color: #f8f9fa !important;
    color: #6c757d !important;
    border-color: #ced4da !important;
    cursor: not-allowed !important;
  }

  .main-plan-input-seibies, 
  .item-plan-input-seibies {
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

  .yearly-container-seibies {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 15px;
    margin-top: 15px;
  }

  .yearly-box-seibies {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 15px;
    background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    position: relative;
  }

  .yearly-header-seibies {
    text-align: center;
    font-weight: 600;
    color: #0078d4;
    margin-bottom: 15px;
    padding-bottom: 8px;
    border-bottom: 1px solid #eee;
  }

  .yearly-inputs-container-seibies {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 15px;
  }

  .yearly-input-group-seibies {
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .yearly-input-group-seibies label {
    font-size: 12px;
    color: #666;
    margin-bottom: 6px;
    font-weight: 600;
  }

  .yearly-input-group-seibies input {
    width: 100%;
    color:#000 !important;
    padding: 8px 10px;
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    font-size: 14px;
    background: #fff;
    transition: all 0.2s ease;
  }

  .yearly-input-group-seibies input:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
    outline: none;
  }

  .yearly-plan-input-seibies {
    background-color: #f8f9fa !important;
    color: #6c757d !important;
    border-color: #ced4da !important;
  }

  input.yearly-current-input-seibies {
    background: #fff !important;
  }

  .month-total-invalid-seibies {
    background: red;
    padding: 8px;
    color: white;
  }

  .month-total-valid-seibies {
    background: #28a745;
    padding: 10px;
    color: white;
  }

  input.main-current-input-seibies {
    background: #fff !important;
  }

  input.item-current-input-seibies {
    background: #fff !important;
  }

  #itemsColor-seibies {
    background: #fff !important;
  }

  .detailed-months-row-seibies {
    display: flex;
    gap: 20px;
    margin-bottom: 20px;
    width: 100%;
  }

  .current-month-container-seibies {
    flex: 0 0 45%;
    max-width: 45%;
  }

  input.item-date-input-seibies {
    background-color: white !important;
  }

  input.item-forecast-input-seibies {
    background-color: white !important;
  }

  .next-months-container-seibies {
    flex: 0 0 55%;
    max-width: 55%;
    display: flex;
    gap: 15px;
  }

  .next-month-box-seibies {
    flex: 1;
    min-width: 0;
  }

  .detailed-month-box-seibies {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 15px;
    background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    position: relative;
    height: fit-content;
  }

  .current-month-box-seibies {
    background: linear-gradient(135deg, #ffffff 0%, #f0f8ff 100%);
    border: 2px solid #0078d4;
  }

  .next-month-box-seibies .detailed-month-box-seibies {
    background: linear-gradient(135deg, #ffffff 0%, #f9f9f9 100%);
    border: 1px solid #ddd;
  }

  .current-month-indicator-seibies {
    position: absolute;
    top: 8px;
    right: 8px;
    background: #0078d4;
    color: white;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 10px;
    font-weight: bold;
  }

  .enhanced-inputs-container-seibies {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr 1fr;
    gap: 12px;
    margin-bottom: 15px;
  }

  .enhanced-input-group-seibies {
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .enhanced-input-group-seibies label {
    font-size: 11px;
    color: #666;
    margin-bottom: 4px;
    font-weight: 600;
    text-align: center;
  }

  .enhanced-input-group-seibies input,
  .enhanced-input-group-seibies select {
    width: 100%;
    padding: 6px 8px;
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    font-size: 12px;
    background: #fff;
    transition: all 0.2s ease;
  }

  .enhanced-input-group-seibies input:focus,
  .enhanced-input-group-seibies select:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 2px rgba(0,120,212,0.2);
    outline: none;
  }

  .forecast-date-input-seibies {
    background: #fff !important;
    color: #333 !important;
    border: 1px solid #ddd !important;
  }

  .current-month-header-table-seibies { 
    background: #ff9800 !important;
    color: white !important;
  }

  .current-month-total-seibies {
    background: #ffb74d !important;
    font-weight: bold;
    border-left: 2px solid #ff9800 !important;
  }

  .current-month-cell-seibies {
    background: #fff3e0 !important;
    border-left: 2px solid #ff9800 !important;
  }
  .current-month-cell-right-seibies { 
    border-right: 2px solid #ff9800 !important;
  }

  .current-month-header-seibies {
    background: #2196f3 !important;
    position: relative;
  }

  .current-month-header-seibies::after {
    position: absolute;
    right: 5px;
    top: 50%;
    transform: translateY(-50%);
    font-size: 12px;
  }

  .items-section-seibies {
    margin-top: 15px;
    padding-top: 15px;
    border-top: 2px solid #e0e0e0;
  }

  .items-header-seibies {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr 1fr 2fr;
    gap: 8px;
    margin-bottom: 8px;
    font-size: 11px;
    color: #666;
    text-align: center;
    font-weight: 600;
    padding-bottom: 4px;
    border-bottom: 1px solid #eee;
  }

  .item-row-seibies {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr 1fr 1fr;
    gap: 8px;
    margin-bottom: 6px;
    align-items: center;
  }

  .item-row-seibies input {
    padding: 5px 8px;
    border: 1px solid #ddd;
    border-radius: 3px;
    font-size: 12px;
    background: #fff;
    transition: all 0.2s ease;
  }

  .item-row-seibies input[type="text"] { text-align: left; }
  .item-row-seibies input[type="number"] { text-align: right; }
  .item-row-seibies input[type="date"] { text-align: center; }

  .item-row-seibies input:focus {
    border-color: #0078d4;
    box-shadow: 0 0 0 1px rgba(0,120,212,0.2);
    outline: none;
  }

  .next-month-box-seibies input[data-type="forecast"],
  .next-month-box-seibies input[data-type="date"],
  .next-month-box-seibies input[data-item-type="forecast"],
  .next-month-box-seibies input[data-item-type="date"] {
    background-color: #f8f9fa !important;
    color: #6c757d !important;
    border-color: #ced4da !important;
    cursor: not-allowed !important;
  }

  .simple-month-seibies {
    background: #fff;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, .1);  
    padding: 15px;
    position: relative;
  }

  .next-months-container-seibies .enhanced-inputs-container-seibies {
    grid-template-columns: 1fr 1fr;
  }

  .next-months-container-seibies .enhanced-input-group-seibies:nth-child(3),
  .next-months-container-seibies .enhanced-input-group-seibies:nth-child(4) {
    display: none;
  }

  .next-months-container-seibies .items-header-seibies {
    grid-template-columns: 2fr 1fr 1fr;
  }

  .next-months-container-seibies .item-row-seibies {
    grid-template-columns: 2fr 1fr 1fr;
  }

  .next-months-container-seibies .items-header-seibies div:nth-child(4),
  .next-months-container-seibies .items-header-seibies div:nth-child(5),
  .next-months-container-seibies .item-row-seibies input:nth-child(4),
  .next-months-container-seibies .item-row-seibies input:nth-child(5) {
    display: none;
  }

  .input-invalid-seibies {
    border-color: #dc3545 !important;
    background-color: #f8f9fa !important;
  }

  .input-warning-seibies {
    border-color: #ffc107 !important;
    background-color: #fffbf0 !important;
  }

  .btn-loading-seibies {
    opacity: 0.7;
    cursor: not-allowed !important;
  }

  .btn-loading-seibies:hover {
    transform: none !important;
  }

  .obl-section-container-seibies.critical-seibies {
    border-left: 4px solid #dc3545;
  }

  .obl-section-container-seibies.warning-seibies {
    border-left: 4px solid #ffc107;
  }

  .obl-section-container-seibies.normal-seibies {
    border-left: 4px solid #28a745;
  }

  .obl-exceed-toast {
    font-size: 14px;
    min-width: 400px;
    background: #fee;
    border: 1px solid #f5c6cb;
  }

  .swal2-popup.swal2-toast {
    font-size: 14px !important;
  }

  body.import-modal-open {
    overflow: hidden !important;
  }

  .import-modal-seibies {
    display: none;
    position: fixed !important;
    top: 0 !important;
    left: 0 !important;
    width: 100% !important;
    height: 100% !important;
    background-color: rgba(0, 0, 0, 0.7) !important;
    z-index: 999999 !important;
    justify-content: center !important;
    align-items: center !important;
  }

  .import-modal-content-seibies {
    background-color: white !important;
    padding: 20px !important;
    border-radius: 10px !important;
    width: 90% !important;
    max-width: 800px !important;
    max-height: 90vh !important;
    overflow: auto !important;
    box-shadow: 0 5px 30px rgba(0, 0, 0, 0.3) !important;
    position: relative !important;
  }

  /* ----- Buttons & Actions ----- */
  .btn-primary-seibies {
    background: #0078d4;
    color: white;
    padding: 8px 12px;
    border: none;
    border-radius: 5px;
    font-size: 14px;
    cursor: pointer;
    transition: all 0.2s ease;
  }

  .btn-primary-seibies:hover:not(:disabled) {
    background: #106ebe;
    transform: translateY(-1px);
  }

  .btn-primary-seibies:disabled {
    background: #a19f9d;
    cursor: not-allowed;
  }

  .btn-secondary-seibies {
    background: #6c757d;
    color: white;
    padding: 8px 12px;
    border: none;
    border-radius: 5px;
    font-size: 14px;
    cursor: pointer;
    transition: all 0.2s ease;
  }

  .btn-secondary-seibies:hover:not(:disabled) {
    background: #5a6268;
    transform: translateY(-1px);
  }

  .clear-filters-btn-seibies {
    background: #dc3545 !important;
  }

  .clear-filters-btn-seibies:hover {
    background: #c82333 !important;
  }

  .export-btn-seibies {
    background: #17a2b8 !important;
  }

  .export-btn-seibies:hover {
    background: #138496 !important;
  }

  .edit-action-btn-seibies {
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

  .edit-action-btn-seibies:hover {
    background: linear-gradient(135deg, #106ebe 0%, #005a9e 100%);
    transform: translateY(-1px);
    box-shadow: 0 2px 4px rgba(0,0,0,0.15);
  }

  /* ----- Forms & Inputs ----- */
  .form-row-seibies { 
    display:flex; 
    flex-wrap:wrap; 
    gap:20px; 
    margin-bottom:20px;
  }   

  .form-group-seibies { 
    flex:1 1 250px; 
    display:flex; 
    flex-direction:column;     
  }

  .form-group-seibies label {
    font-weight:600;
    margin-bottom:5px;
    color:#333;
  }

  .form-group-seibies input,
  .form-group-seibies select {
    padding:8px 12px;
    border:1px solid #ddd;
    border-radius:5px;
    font-size:14px;
  }

  .project-details-grid-seibies {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 15px;
    margin-top: 15px;
    padding: 15px;
    background: #f8f9fa;
    border-radius: 8px;
    border: 1px solid #e9ecef;
  }

  .detail-item-seibies {
    display: flex;
    flex-direction: column;
  }

  .detail-item-seibies strong {
    font-size: 12px;
    color: #666;
    margin-bottom: 2px;
  }

  .detail-item-seibies span {
    font-size: 14px;
    color: #333;
  }

  .edit-mode-seibies .project-select-container-seibies select {
    background-color: #f8f9fa !important;
    color: #666 !important;
    border-color: #d0d0d0 !important;
    cursor: not-allowed !important;
  }
`,
    script: () => {
        setTimeout(async () => {
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/DigitalTransformationTest";
            const listName = "SEIBIES";
            const projectListName = "Project List";
            // DOM Elements
            const tableBody = document.querySelector(".seibies-table-body");
            const monthColumnsHeader = document.querySelector("#monthColumnsHeaderSeibies");
            const totalsHeader = document.querySelector("#totalsHeaderSeibies");
            const formSection = document.querySelector(".form-section-seibies");
            const tableSection = document.querySelector(".table-section-seibies");
            const showFormBtn = document.querySelector(".show-form-btn-seibies");
            const projectSelect = document.querySelector("#projectSelectSeibies");
            const historicalContainer = document.querySelector("#historicalContainerSeibies");
            const firstRowContainer = document.querySelector("#firstRowContainerSeibies");
            const remainingContainer = document.querySelector("#remainingContainerSeibies");
            const yearlyContainer = document.querySelector("#yearlyContainerSeibies");
            const backBtn = document.querySelector(".back-to-table-seibies");
            const form = document.querySelector("#seibiesForm");
            const exportBtn = document.querySelector(".export-btn-seibies");
            const oblLastMonthInput = document.querySelector("#oblLastMonthSeibies");
            const totalCurrentValueSpan = document.querySelector("#totalCurrentValueSeibies");
            const remainingValueSpan = document.querySelector("#remainingValueSeibies");
            const oblLineFill = document.querySelector("#oblLineFillSeibies");
            const oblValidationMessage = document.querySelector("#oblValidationMessageSeibies");
            const formTitle = document.querySelector("#formTitleSeibies");
            // Filter Elements
            const filterYear = document.querySelector("#filterYearSeibies");
            const filterMonth = document.querySelector("#filterMonthSeibies");
            const filterProject = document.querySelector("#filterProjectSeibies");
            const filterCustomer = document.querySelector("#filterCustomerSeibies");
            const filterProjectManager = document.querySelector("#filterProjectManagerSeibies");
            const filterGroupHead = document.querySelector("#filterGroupHeadSeibies");
            const filterProjectReference = document.querySelector("#filterProjectReferenceSeibies");
            // Action Buttons
            const clearFiltersBtn = document.querySelector(".clear-filters-btn-seibies");
            // Pagination Elements
            const firstPageBtn = document.querySelector("#firstPageSeibies");
            const prevPageBtn = document.querySelector("#prevPageSeibies");
            const nextPageBtn = document.querySelector("#nextPageSeibies");
            const lastPageBtn = document.querySelector("#lastPageSeibies");
            const currentPageSpan = document.querySelector("#currentPageSeibies");
            const totalPagesSpan = document.querySelector("#totalPagesSeibies");
            const pageNumbersDiv = document.querySelector("#pageNumbersSeibies");
            const startItemSpan = document.querySelector("#startItemSeibies");
            const endItemSpan = document.querySelector("#endItemSeibies");
            const totalItemsSpan = document.querySelector("#totalItemsSeibies");
            const pageSizeSelect = document.querySelector("#pageSizeSelectSeibies");
            let projects = [];
            let forecasts = [];
            let editingId = null;
            let allMonths = [];
            // Pagination variables
            let pageSize = 10;
            let currentPage = 1;
            let totalPages = 1;
            let filteredData = [];
            // ========== UTILITY FUNCTIONS ==========
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
                return `${project.Title} - ${currentYear} SEIBIES Forecast - ${totalValue.toLocaleString()}`;
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
                return generateIdealMonthStructure();
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
                        this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                    }
                    else {
                        this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
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
            const getNumericValue = (input) => {
                const value = input.value.trim();
                if (value === '' || value === '0')
                    return 0;
                return parseInt(value) || 0;
            };
            const setNumericValue = (input, value) => {
                if (value === 0) {
                    input.value = '';
                    input.className = input.className.replace('input-with-value-seibies', 'input-empty-seibies');
                }
                else {
                    input.value = value.toString();
                    input.className = input.className.replace('input-empty-seibies', 'input-with-value-seibies');
                }
            };
            const showOBLErrorToast = (message, isWarning = false) => {
                var _a;
                const existingToast = document.querySelector('.obl-error-toast-seibies');
                if (existingToast)
                    existingToast.remove();
                const toast = document.createElement('div');
                toast.className = `obl-error-toast-seibies ${isWarning ? 'warning-seibies' : ''}`;
                toast.innerHTML = `
          ${message}
          <button class="toast-close-seibies">&times;</button>
        `;
                document.body.appendChild(toast);
                setTimeout(() => {
                    if (toast.parentNode) {
                        toast.remove();
                    }
                }, 5000);
                (_a = toast.querySelector('.toast-close-seibies')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', () => {
                    toast.remove();
                });
            };
            // ========== MONTH STRUCTURE FUNCTIONS ==========
            const generateIdealMonthStructure = () => {
                const months = [];
                const currentDate = new Date();
                for (let i = -3; i < 0; i++) {
                    const historicalDate = new Date(currentDate);
                    historicalDate.setMonth(historicalDate.getMonth() + i);
                    months.push(formatMonthKey(historicalDate));
                }
                for (let i = 0; i < 3; i++) {
                    const detailedDate = new Date(currentDate);
                    detailedDate.setMonth(detailedDate.getMonth() + i);
                    months.push(formatMonthKey(detailedDate));
                }
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
                const idealMonths = generateIdealMonthStructure();
                const freshData = {};
                const currentMonthKey = formatMonthKey(new Date());
                idealMonths.forEach((monthKey) => {
                    const isCurrentMonth = monthKey === currentMonthKey;
                    freshData[monthKey] = {
                        plan: 0,
                        actual: 0,
                        forecast: 0,
                        date: isCurrentMonth ? '' : undefined
                    };
                });
                const lastMonthKey = idealMonths[idealMonths.length - 1];
                freshData.yearly = createFreshYearlyStructure(lastMonthKey);
                return freshData;
            };
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
                if (originalMonths[0] === idealMonths[0]) {
                    return forecastData;
                }
                const updatedData = {};
                idealMonths.forEach((monthKey) => {
                    const existingData = forecastData[monthKey];
                    if (existingData && isForecastMonthData(existingData)) {
                        updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
                    }
                    else {
                        updatedData[monthKey] = {
                            plan: 0,
                            actual: 0
                        };
                    }
                });
                const currentMonthIndex = idealMonths.indexOf(currentMonthKey);
                idealMonths.forEach((monthKey, index) => {
                    const monthData = updatedData[monthKey];
                    if (monthData && isForecastMonthData(monthData)) {
                        const isCurrentOrFuture = index >= currentMonthIndex;
                        if (isCurrentOrFuture) {
                            const oldActual = monthData.actual;
                            monthData.plan = oldActual;
                            monthData.actual = 0;
                        }
                    }
                });
                if (forecastData.yearly) {
                    updatedData.yearly = JSON.parse(JSON.stringify(forecastData.yearly));
                }
                else {
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
                    // GH Heads will get ALL projects
                    // PM Managers will get only their projects
                    // Engineering roles will get only their projects
                    // READ_ALL users will get ALL projects
                    const assignedProjects = getAssignedProjects(projects);
                    console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);
                    console.log(`👑 Is GH Head: ${isGHHead()}`);
                    console.log(`📋 User Groups:`, getUserGroups());
                    // ✅ Populate project select dropdown with assigned projects
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
                    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,SeibiesActual,CreatedDate,UpdatedDate&$expand=Project&$top=5000`, {
                        headers: { Accept: "application/json;odata=verbose" }
                    });
                    const data = await res.json();
                    // ✅ Get assigned projects (GH Heads get ALL)
                    const assignedProjects = getAssignedProjects(projects);
                    const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                    const allForecasts = (((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || []).map(forecast => (Object.assign(Object.assign({}, forecast), { ForecastData: typeof forecast.ForecastData === 'string' ?
                            forecast.ForecastData :
                            JSON.stringify(forecast.ForecastData || {}), SeibiesActual: typeof forecast.SeibiesActual === 'string' ?
                            forecast.SeibiesActual :
                            JSON.stringify(forecast.SeibiesActual || {}) })));
                    // ✅ Filter forecasts based on assigned projects
                    // GH Heads will see ALL forecasts because assignedProjects contains ALL projects
                    forecasts = allForecasts.filter(forecast => {
                        var _a;
                        const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        return projectTitle && assignedProjectTitles.has(projectTitle);
                    });
                    console.log(`✅ Loaded ${forecasts.length} SEIBIES forecasts`);
                    console.log(`👑 Is GH Head: ${isGHHead()}`);
                    console.log(`📊 Total forecasts: ${forecasts.length}`);
                }
                catch (error) {
                    console.error("Error loading SEIBIES forecasts:", error);
                    toast("error", "Failed to load SEIBIES forecasts");
                }
            };
            // ========== EXPORT FUNCTIONS ==========
            const getMonthName = (monthNumber) => {
                const months = [
                    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
                ];
                return months[parseInt(monthNumber) - 1] || "Month";
            };
            const generateHeadersStructure = (selectedYear, selectedMonth, exportData) => {
                const baseHeaders = [
                    "Profit Centre",
                    "SO/Project Ref",
                    "Customer Name",
                    "Project Name",
                    "Group Head",
                    "Project Manager",
                    "SEI-BIES",
                ];
                const selectedMonthName = getMonthName(selectedMonth);
                const selectedMonthHeaders = [
                    `${selectedMonthName} ${selectedYear} (P)`,
                    `${selectedMonthName} ${selectedYear} (C)`,
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
                    monthlyHeaders.push(`${displayLabel} (P)`, `${displayLabel} (C)`);
                });
                const lastFiveYears = getLastFiveYears(exportData);
                const yearlyHeaders = [];
                lastFiveYears.forEach(year => {
                    yearlyHeaders.push(`${year} (P)`, `${year} (C)`);
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
            const exportToExcelWithFullStyle = () => {
                const selectedYear = filterYear.value;
                const selectedMonth = filterMonth.value;
                if (!selectedYear || !selectedMonth) {
                    toast("error", "Please select year and month to export");
                    return;
                }
                const exportData = filteredData.length > 0 ? filteredData : forecasts;
                if (exportData.length === 0) {
                    toast("error", "No SEIBIES data available to export");
                    return;
                }
                try {
                    const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);
                    const excelData = exportData.map(forecast => {
                        var _a, _b;
                        const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                        const proj = projects.find(p => p.Title === projectTitle);
                        const forecastData = safeJsonParse(forecast.ForecastData, {});
                        const oblValue = forecast.OBL || 0;
                        const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
                        const monthName = getMonthName(selectedMonth);
                        const selectedMonthForecastData = forecastData[selectedMonthKey] || { plan: 0, actual: 0 };
                        const selectedMonthPlan = selectedMonthForecastData.plan || 0;
                        const selectedMonthCurrent = selectedMonthForecastData.actual || 0;
                        const rowData = {
                            "Profit Centre": forecast.ProfitCentre || "-",
                            "SO/Project Ref": (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || "-",
                            "Customer Name": (proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-",
                            "Project Name": projectTitle,
                            "Group Head": (proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-",
                            "Project Manager": typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-",
                            "SEI-BIES": oblValue,
                            [`${monthName} ${selectedYear} (P)`]: selectedMonthPlan,
                            [`${monthName} ${selectedYear} (C)`]: selectedMonthCurrent,
                        };
                        allMonths.forEach(monthKey => {
                            const [year, month] = monthKey.split('-');
                            const monthLabel = getMonthName(month);
                            const displayLabel = `${monthLabel} ${year}`;
                            const forecastMonthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                            const planValue = forecastMonthData.plan || 0;
                            const currentValue = forecastMonthData.actual || 0;
                            rowData[`${displayLabel} (P)`] = planValue;
                            rowData[`${displayLabel} (C)`] = currentValue;
                        });
                        lastFiveYears.forEach(year => {
                            var _a;
                            const yearData = ((_a = forecastData.yearly) === null || _a === void 0 ? void 0 : _a[year]) || { plan: 0, actual: 0 };
                            rowData[`${year} (P)`] = yearData.plan;
                            rowData[`${year} (C)`] = yearData.actual;
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
                        if (header.includes('SEI-BIES') || header.includes('Balance'))
                            return { wch: 12 };
                        if (header.includes('(P)') || header.includes('(C)'))
                            return { wch: 10 };
                        return { wch: 12 };
                    });
                    ws['!cols'] = colWidths;
                    ws['!freeze'] = { x: 0, y: 1 };
                    XLSX.utils.book_append_sheet(wb, ws, "SEIBIES Forecast");
                    const monthName = getMonthName(selectedMonth);
                    const fileName = `SEIBIES_Forecast_${monthName}_${selectedYear}_Filtered.xlsx`;
                    XLSX.writeFile(wb, fileName);
                    toast("success", `Exported ${exportData.length} filtered SEIBIES records for ${monthName} ${selectedYear}`);
                }
                catch (error) {
                    toast("error", "Failed to export SEIBIES data to Excel");
                }
            };
            // ========== PAGINATION FUNCTIONS ==========
            const initializePagination = (data) => {
                filteredData = [...data];
                calculateTotalPages();
                currentPage = 1;
                updatePaginationControls();
                renderCurrentPage();
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                setTimeout(() => {
                    renderTotalsHeader(displayMonths, currentMonthKey);
                }, 100);
            };
            const calculateTotalPages = () => {
                totalPages = Math.ceil(filteredData.length / pageSize);
                if (totalPages === 0)
                    totalPages = 1;
            };
            const changePageSize = (newSize) => {
                if (newSize === pageSize)
                    return;
                pageSize = newSize;
                const oldStartItem = (currentPage - 1) * pageSize;
                currentPage = Math.floor(oldStartItem / pageSize) + 1;
                calculateTotalPages();
                if (currentPage > totalPages) {
                    currentPage = totalPages || 1;
                }
                updatePaginationControls();
                renderCurrentPage();
                toast("info", `Showing ${pageSize} SEIBIES records per page`);
            };
            const updatePaginationControls = () => {
                currentPageSpan.textContent = currentPage.toString();
                totalPagesSpan.textContent = totalPages.toString();
                const startItem = (currentPage - 1) * pageSize + 1;
                const endItem = Math.min(currentPage * pageSize, filteredData.length);
                startItemSpan.textContent = startItem.toString();
                endItemSpan.textContent = endItem.toString();
                totalItemsSpan.textContent = filteredData.length.toString();
                if (pageSizeSelect) {
                    pageSizeSelect.value = pageSize.toString();
                }
                firstPageBtn.disabled = currentPage === 1 || filteredData.length === 0;
                prevPageBtn.disabled = currentPage === 1 || filteredData.length === 0;
                nextPageBtn.disabled = currentPage === totalPages || filteredData.length === 0;
                lastPageBtn.disabled = currentPage === totalPages || filteredData.length === 0;
                updatePageNumbers();
            };
            const updatePageNumbers = () => {
                pageNumbersDiv.innerHTML = '';
                if (filteredData.length === 0)
                    return;
                const maxVisiblePages = 5;
                let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
                let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
                if (endPage - startPage + 1 < maxVisiblePages) {
                    startPage = Math.max(1, endPage - maxVisiblePages + 1);
                }
                if (startPage > 1) {
                    const dotsBtn = document.createElement('button');
                    dotsBtn.className = 'page-number-seibies dots-seibies';
                    dotsBtn.textContent = '...';
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
                for (let i = startPage; i <= endPage; i++) {
                    const pageBtn = document.createElement('button');
                    pageBtn.className = `page-number-seibies ${i === currentPage ? 'active-seibies' : ''}`;
                    pageBtn.textContent = i.toString();
                    pageBtn.addEventListener('click', () => goToPage(i));
                    pageNumbersDiv.appendChild(pageBtn);
                }
                if (endPage < totalPages) {
                    const dotsBtn = document.createElement('button');
                    dotsBtn.className = 'page-number-seibies dots-seibies';
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
            };
            const setupPaginationEventListeners = () => {
                if (pageSizeSelect) {
                    pageSizeSelect.addEventListener("change", (e) => {
                        const newSize = parseInt(e.target.value);
                        changePageSize(newSize);
                    });
                }
            };
            // ========== TABLE RENDERING FUNCTIONS ==========
            const calculateTotals = (data, displayMonths, currentMonthKey) => {
                const totals = {
                    obl: 0,
                    monthly: {}
                };
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
            const renderTotalsHeader = (displayMonths, currentMonthKey) => {
                if (!totalsHeader)
                    return;
                totalsHeader.innerHTML = '';
                const allData = forecasts;
                const filtered = filteredData.length ? filteredData : allData;
                const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
                const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);
                const buildTotalsRow = (label, totals, css) => {
                    return `
            <tr class="${css}">
              <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;"></td>
              <td class="project-col" style="min-width:180px; max-width:180px; width:180px; text-align:left; font-weight:bold; background-color:${css === 'all-projects-total-seibies' ? '#d1ecf1' : '#d4edda'};">
                ${label}
              </td>
              <td class="pm-col" style="min-width:100px; max-width:100px; width:100px;"></td>
              <td class="gh-col" style="min-width:100px; max-width:100px; width:100px;"></td>
              <td class="customer-col" style="min-width:130px; max-width:130px; width:130px;"></td>
              <td class="obl-col" style="min-width:90px; max-width:90px; width:90px; text-align:right; font-weight:bold;">${totals.obl.toLocaleString('en-IN')}</td>
              ${displayMonths.map(m => {
                        const isCurrentMonth = m === currentMonthKey;
                        const t = totals.monthly[m] || { plan: 0, current: 0 };
                        const cellClass = isCurrentMonth ? '' : 'non-current-month-cell';
                        return `
                  <td class="month-scroll-col ${cellClass}" style="min-width:60px;">${t.plan.toLocaleString('en-IN')}</td>
                  <td class="month-scroll-col ${cellClass}" style="min-width:60px;">${t.current.toLocaleString('en-IN')}</td>
                `;
                    }).join('')}
            </tr>
          `;
                };
                totalsHeader.innerHTML = `
          <table class="totals-header-table-seibies" style="width: 100%;">
            ${buildTotalsRow(`All (${allData.length})`, allTotals, 'all-projects-total-seibies')}
            ${buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total-seibies')}
          </table>
        `;
            };
            const renderCurrentPage = () => {
                if (filteredData.length === 0) {
                    tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No SEIBIES records found</td></tr>`;
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    return;
                }
                const startIndex = (currentPage - 1) * pageSize;
                const endIndex = Math.min(startIndex + pageSize, filteredData.length);
                const pageData = filteredData.slice(startIndex, endIndex);
                renderTableWithMonthColumns(pageData);
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                setTimeout(() => {
                    renderTotalsHeader(displayMonths, currentMonthKey);
                }, 100);
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
            const renderTableWithMonthColumns = (data) => {
                if (!data.length) {
                    tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No SEIBIES records found</td></tr>`;
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    return;
                }
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                renderTotalsHeader(displayMonths, currentMonthKey);
                // Build main table header - Project Reference and Project are both clickable
                const mainHeaderHTML = `
          <tr>
            <!-- Project Reference - FIRST - now clickable -->
            <th rowspan="2" class="project-ref-col" style="min-width:120px; max-width:120px; width:120px; background:#68e0f0;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Project Reference</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectRef' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="projectRef" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectRef' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="projectRef" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Project Column with Sort - clickable to edit -->
            <th rowspan="2" class="project-col" style="min-width:180px; max-width:180px; width:180px; background:#68e0f0;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Project</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="project" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'project' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="project" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Project Manager -->
            <th rowspan="2" class="pm-col" style="min-width:100px; max-width:100px; width:100px; background:#68e0f0;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Project Manager</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="projectManager" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'projectManager' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="projectManager" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Group Head -->
            <th rowspan="2" class="gh-col" style="min-width:100px; max-width:100px; width:100px; background:#68e0f0;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Group Head</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="groupHead" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'groupHead' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="groupHead" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- Customer -->
            <th rowspan="2" class="customer-col" style="min-width:130px; max-width:130px; width:130px; background:#68e0f0;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">Customer</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="customer" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'customer' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="customer" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            <!-- SEI-BIES Column -->
            <th rowspan="2" class="obl-col" style="min-width:90px; max-width:90px; width:90px; background:#68e0f0;">
              <div class="sort-header-container-seibies">
                <span class="header-label-seibies">SEI-BIES</span>
                <div class="sort-buttons-container-seibies">
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                          data-sort="obl" data-direction="asc" title="Sort Ascending">▲</button>
                  <button class="sort-btn-seibies ${currentSortConfig.column === 'obl' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                          data-sort="obl" data-direction="desc" title="Sort Descending">▼</button>
                </div>
              </div>
            </th>
            
            ${displayMonths.map((month, monthIndex) => {
                    const isCurrentMonth = month === currentMonthKey;
                    const monthLabel = formatMonthLabel(new Date(month + "-01"));
                    const displayLabel = isCurrentMonth ? `${monthLabel} (Current Month)` : monthLabel;
                    const headerClass = isCurrentMonth ? 'current-month-header-seibies' : 'non-current-month-header';
                    return `
                <th colspan="2" class="month-scroll-col month-column-header-seibies ${headerClass}" style="min-width:60px; color:white; ${!isCurrentMonth ? 'border-left: 1px solid #ffffff !important; border-right: 1px solid #ffffff !important;' : ''}">
                  <div class="sort-header-container-seibies">
                    <span class="header-label-seibies">${displayLabel}</span>
                    <div class="sort-buttons-container-seibies">
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="asc" data-month-index="${monthIndex}" title="Sort by Month Total Ascending">▲</button>
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && !currentSortConfig.subColumn && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="desc" data-month-index="${monthIndex}" title="Sort by Month Total Descending">▼</button>
                    </div>
                  </div>
                </th>
              `;
                }).join('')}
          </tr>
          
          <tr>
            ${displayMonths.map((month, monthIndex) => {
                    const isCurrentMonth = month === currentMonthKey;
                    const subHeaderClass = isCurrentMonth ? 'current-month-header-table-seibies' : 'non-current-month-subheader';
                    return `
                <th class="month-scroll-col month-subheader-seibies ${subHeaderClass}" style="min-width:60px; ${!isCurrentMonth ? 'border-left: 1px solid #ffffff !important; border-right: 1px solid #ffffff !important;' : ''}">
                  <div class="sort-header-container-seibies">
                    <span>Plan</span>
                    <div class="sort-buttons-container-seibies">
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Ascending">▲</button>
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Descending">▼</button>
                    </div>
                  </div>
                </th>
                <th class="month-scroll-col month-subheader-seibies ${subHeaderClass}" style="min-width:60px; ${!isCurrentMonth ? 'border-left: 1px solid #ffffff !important; border-right: 1px solid #ffffff !important;' : ''}">
                  <div class="sort-header-container-seibies">
                    <span>Current</span>
                    <div class="sort-buttons-container-seibies">
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Ascending">▲</button>
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'current' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="current" data-month-index="${monthIndex}" title="Sort Current Descending">▼</button>
                    </div>
                  </div>
                </th>
              `;
                }).join('')}
          </tr>
        `;
                monthColumnsHeader.innerHTML = mainHeaderHTML;
                let tableBodyHTML = '';
                data.forEach((forecast) => {
                    var _a, _b;
                    const projectTitle = ((_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title) || "-";
                    const proj = projects.find(p => p.Title === projectTitle);
                    const forecastData = safeJsonParse(forecast.ForecastData, {});
                    const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-IN');
                    // Project Reference - now clickable link
                    const projectRefLink = (proj === null || proj === void 0 ? void 0 : proj.ProjectReference) ?
                        `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${proj.ProjectReference}</a>` :
                        "-";
                    // Project name link - clickable
                    const projectNameLink = `<a href="#" class="edit-trigger" data-id="${forecast.Id}" style="color:#0078d4; text-decoration:underline; cursor:pointer;">${projectTitle}</a>`;
                    let rowHTML = `
            <tr data-forecast-id="${forecast.Id}">
              <!-- Project Reference - clickable link -->
              <td class="project-ref-col" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
              
              <!-- Project - clickable link -->
              <td class="project-col" style="min-width:180px; max-width:180px; width:180px;">${projectNameLink}</td>
              
              <!-- Project Manager -->
              <td class="pm-col" style="min-width:100px; max-width:100px; width:100px;">${typeof (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) === "object" ? (_b = proj.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title : (proj === null || proj === void 0 ? void 0 : proj.ProjectManager) || "-"}</td>
              
              <!-- Group Head -->
              <td class="gh-col" style="min-width:100px; max-width:100px; width:100px;">${(proj === null || proj === void 0 ? void 0 : proj.GroupHead) || "-"}</td>
              
              <!-- Customer -->
              <td class="customer-col" style="min-width:130px; max-width:130px; width:130px;">${(proj === null || proj === void 0 ? void 0 : proj.CustomerName) || "-"}</td>
              
              <!-- SEI-BIES -->
              <td class="obl-col" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
          `;
                    displayMonths.forEach((monthKey) => {
                        const isCurrentMonth = monthKey === currentMonthKey;
                        const monthData = forecastData[monthKey] || { plan: 0, actual: 0 };
                        const planValue = monthData.plan || 0;
                        const currentValue = monthData.actual || 0;
                        const cellClass = isCurrentMonth ? '' : 'non-current-month-cell';
                        rowHTML += `
              <td class="month-scroll-col ${cellClass}" style="min-width:60px;">${planValue.toLocaleString('en-IN')}</td>
              <td class="month-scroll-col ${cellClass}" style="min-width:60px;">${currentValue.toLocaleString('en-IN')}</td>
            `;
                    });
                    rowHTML += `</tr>`;
                    tableBodyHTML += rowHTML;
                });
                tableBody.innerHTML = tableBodyHTML;
                attachSortButtonListeners();
                attachEditTriggerListeners();
            };
            const attachSortButtonListeners = () => {
                document.querySelectorAll('.sort-btn-seibies[data-sort]').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const button = e.currentTarget;
                        const sortColumn = button.dataset.sort;
                        const direction = button.dataset.direction;
                        const subColumn = button.dataset.subcolumn;
                        const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;
                        if (sortColumn && direction) {
                            handleSeibiesSort(sortColumn, direction, subColumn, monthIndex);
                        }
                    });
                });
            };
            const handleSeibiesSort = (column, direction, subColumn, monthIndex) => {
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
                        case 'projectRef':
                            const projA = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projB = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            valueA = (projA === null || projA === void 0 ? void 0 : projA.ProjectReference) || "";
                            valueB = (projB === null || projB === void 0 ? void 0 : projB.ProjectReference) || "";
                            break;
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
                        case 'customer':
                            const projACust = projects.find(p => { var _a; return p.Title === ((_a = a.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            const projBCust = projects.find(p => { var _a; return p.Title === ((_a = b.Project) === null || _a === void 0 ? void 0 : _a.Title); });
                            valueA = (projACust === null || projACust === void 0 ? void 0 : projACust.CustomerName) || "";
                            valueB = (projBCust === null || projBCust === void 0 ? void 0 : projBCust.CustomerName) || "";
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
                else if (column === 'projectRef') {
                    columnName = 'Project Reference';
                }
                else if (column === 'obl') {
                    columnName = 'SEI-BIES';
                }
                else if (column === 'project') {
                    columnName = 'Project';
                }
                else if (column === 'customer') {
                    columnName = 'Customer';
                }
                else if (column === 'projectManager') {
                    columnName = 'Project Manager';
                }
                else if (column === 'groupHead') {
                    columnName = 'Group Head';
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
                initializePagination(forecasts);
                const displayMonths = getDisplayMonths();
                const currentMonthKey = formatMonthKey(new Date());
                setTimeout(() => {
                    renderTotalsHeader(displayMonths, currentMonthKey);
                }, 100);
                toast("info", "All SEIBIES filters and sorting cleared");
            };
            // ========== RENDER PROJECT DETAILS ==========
            const renderProjectDetails = (project) => {
                var _a;
                const projectNameDisplay = document.querySelector('#projectNameDisplaySeibies');
                if (projectNameDisplay)
                    projectNameDisplay.textContent = project.Title || "-";
                const projectRef = document.querySelector('#projectRefSeibies');
                if (projectRef)
                    projectRef.textContent = project.ProjectReference || "-";
                const profitCenter = document.querySelector('#profitCenterSeibies');
                if (profitCenter)
                    profitCenter.textContent = project.PC || "-";
                const projectManager = document.querySelector('#projectManagerSeibies');
                if (projectManager) {
                    projectManager.textContent = typeof project.ProjectManager === "object" ?
                        (_a = project.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title : project.ProjectManager || "-";
                }
                const groupHead = document.querySelector('#groupHeadSeibies');
                if (groupHead)
                    groupHead.textContent = project.GroupHead || "-";
                const customer = document.querySelector('#customerSeibies');
                if (customer)
                    customer.textContent = project.CustomerName || "-";
            };
            // ========== HISTORICAL TOGGLE FUNCTION - HIDE BY DEFAULT ==========
            const setupHistoricalToggle = () => {
                var _a, _b;
                const toggleBtn = document.querySelector("#toggleHistoricalBtnSeibies");
                const historicalContainer = document.querySelector("#historicalContainerSeibies");
                const header = document.querySelector(".historical-toggle-header-seibies");
                if (!toggleBtn || !historicalContainer || !header)
                    return;
                // Set initial state - container should be HIDDEN by default
                let isVisible = false;
                historicalContainer.style.display = "none";
                toggleBtn.textContent = 'Show';
                toggleBtn.style.background = '#28a745';
                // Replace button to remove old listeners
                const newBtn = document.createElement('button');
                newBtn.id = 'toggleHistoricalBtnSeibies';
                newBtn.className = 'btn-secondary-seibies';
                newBtn.textContent = 'Show';
                newBtn.type = 'button';
                newBtn.style.cssText = 'padding: 4px 16px; font-size: 13px; background: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; flex-shrink: 0; z-index: 10; position: relative;';
                (_a = toggleBtn.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newBtn, toggleBtn);
                // Replace header to remove all old listeners
                const newHeader = header.cloneNode(true);
                (_b = header.parentNode) === null || _b === void 0 ? void 0 : _b.replaceChild(newHeader, header);
                // Button click handler
                const btn = document.querySelector("#toggleHistoricalBtnSeibies");
                const container = document.querySelector("#historicalContainerSeibies");
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
                // Header click handler (not the button)
                const headerEl = document.querySelector(".historical-toggle-header-seibies");
                if (headerEl) {
                    headerEl.addEventListener("click", function (e) {
                        // If click is on the button or its children, ignore
                        if (e.target.closest("#toggleHistoricalBtnSeibies")) {
                            return;
                        }
                        e.preventDefault();
                        e.stopPropagation();
                        const btn = document.querySelector("#toggleHistoricalBtnSeibies");
                        const container = document.querySelector("#historicalContainerSeibies");
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
            // ========== FORM MANAGEMENT FUNCTIONS ==========
            const showForm = (isEditMode = false, project) => {
                // ✅ Debug logging
                console.log(`📝 showForm called: isEditMode=${isEditMode}, project=${project === null || project === void 0 ? void 0 : project.Title}`);
                console.log(`👑 Is GH Head: ${isGHHead()}`);
                console.log(`📋 Permission Level: ${getPermissionLevel()}`);
                console.log(`🔑 Can Edit: ${canEditProject(project)}`);
                console.log(`📝 Can Create: ${canCreateRecord()}`);
                console.log(`👁️ Can View: ${canViewForm(isEditMode, project)}`);
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
                    toast("error", "❌ You don't have permission to create SEIBIES forecasts");
                    return;
                }
                formSection.style.display = "block";
                tableSection.style.display = "none";
                if (showFormBtn)
                    showFormBtn.style.display = "none";
                // ✅ Show/hide save button
                const saveBtn = document.querySelector('.btn-primary-seibies[type="submit"]');
                if (saveBtn) {
                    const canSave = isEditMode ? canEditProject(project) : canCreateRecord();
                    saveBtn.style.display = canSave ? '' : 'none';
                    console.log(`💾 Save button ${canSave ? 'visible' : 'hidden'}`);
                }
                if (isEditMode) {
                    formSection.classList.add('edit-mode-seibies');
                    projectSelect.disabled = true;
                    formTitle.textContent = "Edit SEIBIES Forecast";
                }
                else {
                    formSection.classList.remove('edit-mode-seibies');
                    projectSelect.disabled = false;
                    formTitle.textContent = "Add SEIBIES Forecast";
                }
            };
            const hideForm = () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                if (showFormBtn)
                    showFormBtn.style.display = "inline-block";
                formSection.classList.remove('edit-mode-seibies');
                projectSelect.disabled = false;
            };
            const resetForm = () => {
                editingId = null;
                form.reset();
                projectSelect.value = "";
                // Clear project details display
                const projectNameDisplay = document.querySelector('#projectNameDisplaySeibies');
                if (projectNameDisplay)
                    projectNameDisplay.textContent = "-";
                const projectRef = document.querySelector('#projectRefSeibies');
                if (projectRef)
                    projectRef.textContent = "-";
                const profitCenter = document.querySelector('#profitCenterSeibies');
                if (profitCenter)
                    profitCenter.textContent = "-";
                const projectManager = document.querySelector('#projectManagerSeibies');
                if (projectManager)
                    projectManager.textContent = "-";
                const groupHead = document.querySelector('#groupHeadSeibies');
                if (groupHead)
                    groupHead.textContent = "-";
                const customer = document.querySelector('#customerSeibies');
                if (customer)
                    customer.textContent = "-";
                historicalContainer.innerHTML = "";
                firstRowContainer.innerHTML = "";
                remainingContainer.innerHTML = "";
                yearlyContainer.innerHTML = "";
                toggleForecastSections(false);
                const oblContainer = document.querySelector('.obl-section-container-seibies');
                if (oblContainer) {
                    oblContainer.className = 'obl-section-container-seibies hidden-section';
                }
                oblLastMonthInput.removeAttribute('readonly');
                oblLastMonthInput.placeholder = "Enter OBL";
                setNumericValue(oblLastMonthInput, 0);
                oblValidationMessage.textContent = "✅ All values are within OBL limit";
                oblValidationMessage.className = 'validation-message-seibies';
                updateOBLValidationDisplay();
            };
            // ========== VALIDATION FUNCTIONS ==========
            const calculateTotalCurrentValue = () => {
                let total = 0;
                const editableMonths = allMonths.slice(3);
                editableMonths.forEach(monthKey => {
                    const currentInputs = document.querySelectorAll(`input.main-current-input-seibies[data-month="${monthKey}"]:not([readonly])`);
                    currentInputs.forEach(input => {
                        total += getNumericValue(input);
                    });
                });
                const yearlyCurrentInputs = document.querySelectorAll(`input.yearly-current-input-seibies:not([readonly])`);
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
                    oblLineFill.className = 'line-fill-seibies';
                    if (totalCurrent > 0) {
                        showOBLErrorToast('⚠️ OBL is zero but current values exist. Please set OBL value.', true);
                        blockAllCurrentInputs();
                        return false;
                    }
                    else {
                        const existingToast = document.querySelector('.obl-error-toast-seibies');
                        if (existingToast)
                            existingToast.remove();
                        unblockAllCurrentInputs();
                        return true;
                    }
                }
                const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
                oblLineFill.style.width = `${percentage}%`;
                oblLineFill.className = 'line-fill-seibies';
                if (percentage > 90) {
                    oblLineFill.classList.add('danger-seibies');
                }
                else if (percentage > 75) {
                    oblLineFill.classList.add('warning-seibies');
                }
                oblValidationMessage.className = 'validation-message-seibies';
                if (totalCurrent > oblValue) {
                    const exceedAmount = totalCurrent - oblValue;
                    showOBLErrorToast(`❌ Total current value exceeds OBL by ${exceedAmount.toLocaleString()}! Reduce values to continue.`);
                    oblValidationMessage.classList.add('danger-seibies');
                    blockIncreasingCurrentInputs(changedInput);
                    return false;
                }
                else if (percentage > 90) {
                    showOBLErrorToast('⚠️ OBL limit nearly reached! Consider reducing values.', true);
                    oblValidationMessage.classList.add('warning-seibies');
                    unblockAllCurrentInputs();
                    return true;
                }
                else {
                    const existingToast = document.querySelector('.obl-error-toast-seibies');
                    if (existingToast)
                        existingToast.remove();
                    oblValidationMessage.textContent = '✅ All values are within OBL limit';
                    unblockAllCurrentInputs();
                    return true;
                }
            };
            const blockAllCurrentInputs = () => {
                const currentInputs = document.querySelectorAll('input[data-type="actual"]:not([readonly])');
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-seibies:not([readonly])');
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
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-seibies:not([readonly])');
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
                const yearlyCurrentInputs = document.querySelectorAll('input.yearly-current-input-seibies');
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
            const updateOBLValidationDisplay = (changedInput) => {
                validateOBLLimit(changedInput);
            };
            // ========== FORECAST FORM FUNCTIONS ==========
            const toggleForecastSections = (show) => {
                const sections = [
                    '.obl-section-container-seibies',
                    '#historicalForecastSectionSeibies',
                    '#detailedForecastSectionSeibies',
                    '#longtermForecastSectionSeibies',
                    '#yearlyForecastSectionSeibies'
                ];
                sections.forEach(selector => {
                    const element = document.querySelector(selector);
                    if (element) {
                        if (show) {
                            element.classList.remove('hidden-section');
                            element.classList.add('visible-section-seibies');
                        }
                        else {
                            element.classList.remove('visible-section-seibies');
                            element.classList.add('hidden-section');
                        }
                    }
                });
            };
            // ========== REMAINING FUNCTIONS ==========
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
                const oblContainer = document.querySelector('.obl-section-container-seibies');
                const fieldNote = oblContainer === null || oblContainer === void 0 ? void 0 : oblContainer.querySelector('.field-note-seibies');
                if (isNewProject) {
                    if (oblContainer) {
                        oblContainer.className = 'obl-section-container-seibies visible-section-seibies new-project-seibies';
                    }
                    oblLastMonthInput.removeAttribute('readonly');
                    oblLastMonthInput.placeholder = "Enter initial OBL for new project";
                    if (fieldNote)
                        fieldNote.textContent = "For new projects: Enter initial OBL value";
                    oblValidationMessage.textContent = "Enter initial OBL value for new project";
                    oblValidationMessage.className = 'validation-message-seibies';
                    const defaultOBL = project.PC ? 100000 : 0;
                    setNumericValue(oblLastMonthInput, defaultOBL);
                }
                else {
                    if (oblContainer) {
                        oblContainer.className = 'obl-section-container-seibies visible-section-seibies existing-project-seibies';
                    }
                    oblLastMonthInput.setAttribute('readonly', 'true');
                    const oblValue = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                    setNumericValue(oblLastMonthInput, oblValue);
                    oblLastMonthInput.placeholder = "OBL from database";
                    if (fieldNote)
                        fieldNote.textContent = "OBL value from database record";
                    if (oblValue > 0) {
                        oblLastMonthInput.value = oblValue.toString();
                        oblLastMonthInput.className = oblLastMonthInput.className.replace('input-empty-seibies', 'input-with-value-seibies');
                    }
                }
                renderForecastInterface(forecastData, existingForecast);
            };
            const getForecastMonths = () => {
                return generateIdealMonthStructure();
            };
            const renderForecastInterface = (existingData, existingForecast) => {
                allMonths = getForecastMonths();
                const displayOBL = (existingForecast === null || existingForecast === void 0 ? void 0 : existingForecast.OBL) || 0;
                setNumericValue(oblLastMonthInput, displayOBL);
                renderHistoricalMonths(allMonths.slice(0, 3), existingData);
                renderFirstRow(allMonths.slice(3, 6), existingData);
                renderRemainingMonths(allMonths.slice(6), existingData);
                renderYearlyForecast(existingData.yearly || createFreshYearlyStructure(getLastMonthKey(existingData)));
                protectHistoricalElements();
                updateOBLValidationDisplay();
                attachForecastEventListeners();
                // Setup historical toggle after rendering - HIDE BY DEFAULT
                setTimeout(() => {
                    setupHistoricalToggle();
                }, 50);
            };
            // ========== RENDER FUNCTIONS FOR FORECAST SECTIONS ==========
            const renderHistoricalMonths = (months, existingData) => {
                let html = '';
                months.forEach((monthKey, index) => {
                    const monthDate = new Date(monthKey + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[monthKey] || { plan: 0, actual: 0 };
                    html += renderHistoricalMonthBox(monthKey, monthLabel, data, index);
                });
                historicalContainer.innerHTML = html;
            };
            const renderHistoricalMonthBox = (monthKey, monthLabel, data, index) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="month-box-seibies historical-month-seibies">
            <div class="month-header-seibies">${monthLabel} (Historical - Readonly)</div>

            <div class="month-block-seibies">
              <div class="main-inputs-container-seibies">
                <div class="main-input-group-seibies">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input-seibies ${planInputClass} number-input-seibies" />
                </div>
                <div class="main-input-group-seibies">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        readonly class="main-current-input-seibies ${actualInputClass} number-input-seibies" />
                </div>
              </div>
            </div>
          </div>
        `;
            };
            const renderFirstRow = (months, existingData) => {
                let html = '';
                html += `<div class="detailed-months-row-seibies">`;
                const currentMonth = months[0];
                if (currentMonth) {
                    const monthDate = new Date(currentMonth + "-01");
                    const monthLabel = formatMonthLabel(monthDate);
                    const data = existingData[currentMonth] || { plan: 0, actual: 0 };
                    html += `<div class="current-month-container-seibies">`;
                    html += renderDetailedMonthBox(currentMonth, monthLabel, data, true);
                    html += `</div>`;
                }
                html += `<div class="next-months-container-seibies">`;
                for (let i = 1; i <= 2; i++) {
                    if (i < months.length) {
                        const monthKey = months[i];
                        const monthDate = new Date(monthKey + "-01");
                        const monthLabel = formatMonthLabel(monthDate);
                        const data = existingData[monthKey] || { plan: 0, actual: 0 };
                        html += `<div class="next-month-box-seibies">`;
                        html += renderDetailedMonthBox(monthKey, monthLabel, data, false);
                        html += `</div>`;
                    }
                }
                html += `</div>`;
                html += `</div>`;
                firstRowContainer.innerHTML = html;
            };
            const renderRemainingMonths = (months, existingData) => {
                let html = '';
                for (let row = 0; row < 3; row++) {
                    html += `<div class="remaining-row-seibies">`;
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
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="yearly-box-seibies">
            <div class="yearly-header-seibies">${year}</div>

            <div class="yearly-inputs-container-seibies">
              <div class="yearly-input-group-seibies">
                <label>Previous</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-year="${year}"
                      readonly class="yearly-plan-input-seibies ${planInputClass} number-input-seibies" />
              </div>
              <div class="yearly-input-group-seibies">
                <label>Current</label>
                <input type="text" value="${displayActualValue}" data-type="actual" data-year="${year}"
                      class="yearly-current-input-seibies ${actualInputClass} number-input-seibies" />
              </div>
            </div>
          </div>
        `;
            };
            const renderDetailedMonthBox = (monthKey, monthLabel, data, isCurrentMonth = false) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const displayForecastValue = data.forecast || 0;
                const displayDateValue = data.date || '';
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const forecastInputClass = displayForecastValue > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const dateInputClass = displayDateValue ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="detailed-month-box-seibies ${isCurrentMonth ? 'current-month-box-seibies' : 'next-month-box-seibies'}">
            <div class="month-header-seibies">
              ${monthLabel}
              ${isCurrentMonth ? '<span class="current-month-indicator-seibies">Current</span>' : ''}
            </div>

            <div class="month-block-seibies">
              <div class="enhanced-inputs-container-seibies">
                <div class="enhanced-input-group-seibies">
                  <label>Previous</label>
                  <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                        readonly class="main-plan-input-seibies ${planInputClass} number-input-seibies" />
                </div>

                <div class="enhanced-input-group-seibies">
                  <label>Current</label>
                  <input type="text" value="${displayActualValue}" data-type="actual" data-month="${monthKey}"
                        class="main-current-input-seibies ${actualInputClass} number-input-seibies" />
                </div>

                ${isCurrentMonth ? `
                  <div class="enhanced-input-group-seibies">
                    <label>Forecast</label>
                    <input type="text" value="${displayForecastValue}" data-type="forecast" data-month="${monthKey}"
                          class="main-forecast-input-seibies ${forecastInputClass} number-input-seibies" />
                  </div>

                  <div class="enhanced-input-group-seibies">
                    <label>Date</label>
                    <input type="date" value="${displayDateValue}" data-type="date" data-month="${monthKey}"
                          class="forecast-date-input-seibies ${dateInputClass}" />
                  </div>
                ` : ''}
              </div>
            </div>
          </div>
        `;
            };
            const renderSimpleMonthBox = (monthKey, monthLabel, data) => {
                const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
                const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
                const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
                return `
          <div class="month-box-seibies simple-month-seibies">
            <div class="month-header-seibies">${monthLabel}</div>

            <div class="main-inputs-container-seibies">
              <div class="main-input-group-seibies">
                <label>Previous</label>
                <input type="text" value="${displayPlanValue}" data-type="plan" data-month="${monthKey}"
                      readonly class="main-plan-input-seibies ${planInputClass} number-input-seibies" />
              </div>
              <div class="main-input-group-seibies">
                <label>Current</label>
                <input type="text" value="${displayActualValue}" id="actualCurrent" data-type="actual" data-month="${monthKey}"
                      class="main-current-input-seibies ${actualInputClass} number-input-seibies" />
              </div>
            </div>
          </div>
        `;
            };
            const protectHistoricalElements = () => {
                document.querySelectorAll('.historical-month-seibies input, .historical-month-seibies select, .historical-month-seibies button').forEach(element => {
                    var _a;
                    const newElement = element.cloneNode(true);
                    (_a = element.parentNode) === null || _a === void 0 ? void 0 : _a.replaceChild(newElement, element);
                });
                document.querySelectorAll('.historical-month-seibies').forEach(section => {
                    section.addEventListener('click', (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        toast("warning", "Historical SEIBIES data cannot be modified");
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
            const attachForecastEventListeners = () => {
                document.querySelectorAll('.number-input-seibies').forEach(input => {
                    allowOnlyNumbers(input);
                });
                document.querySelectorAll('.number-input-seibies:not([readonly])').forEach(input => {
                    input.addEventListener('focus', function () {
                        this.setAttribute('data-old-value', getNumericValue(this).toString());
                    });
                });
                document.querySelectorAll('.main-current-input-seibies:not([readonly])').forEach(input => {
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
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                        }
                        else {
                            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.main-forecast-input-seibies:not([readonly])').forEach(input => {
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
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                        }
                        else {
                            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.yearly-current-input-seibies:not([readonly])').forEach(input => {
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
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                        }
                        else {
                            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                        }
                        updateOBLValidationDisplay(this);
                    });
                });
                document.querySelectorAll('.forecast-date-input-seibies:not([readonly])').forEach(input => {
                    input.addEventListener('change', function () {
                        if (this.value === '') {
                            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
                        }
                        else {
                            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
                        }
                    });
                });
                const validationInputs = document.querySelectorAll(`
          .main-current-input-seibies:not([readonly]),
          .yearly-current-input-seibies:not([readonly]),
          .main-forecast-input-seibies:not([readonly])
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
                    toast("error", "SEIBIES forecast not found");
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
                    toast("warning", "Corrupted SEIBIES forecast data detected. Loaded empty structure.");
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
                    // Apply freeze state to all form fields
                    setFormReadOnly(isReadOnly);
                    // Show/hide save button
                    const saveBtn = document.querySelector('.btn-primary-seibies[type="submit"]');
                    if (saveBtn) {
                        saveBtn.style.display = isReadOnly ? 'none' : '';
                    }
                    // Update form title
                    const formTitleEl = document.querySelector("#formTitleSeibies");
                    if (formTitleEl) {
                        if (isFrozen) {
                            formTitleEl.innerHTML = '🔒 Edit SEIBIES Forecast (Frozen - Read Only)';
                            formTitleEl.style.color = '#856404';
                        }
                        else if (!canEdit) {
                            formTitleEl.innerHTML = '🔒 Edit SEIBIES Forecast (Read Only)';
                            formTitleEl.style.color = '#856404';
                        }
                        else {
                            formTitleEl.innerHTML = 'Edit SEIBIES Forecast';
                            formTitleEl.style.color = '';
                        }
                    }
                }, 300);
                toast("info", `Editing SEIBIES forecast for ${forecast.Title}`);
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
                    if (!canEditProject(project)) {
                        toast("error", "❌ You don't have permission to edit this SEIBIES forecast");
                        return;
                    }
                }
                else {
                    if (!canCreateRecord()) {
                        toast("error", "❌ You don't have permission to create SEIBIES forecasts");
                        return;
                    }
                }
                const isNewProject = !editingId && !forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId; });
                if (!isNewProject && forecasts.some(f => { var _a; return ((_a = f.Project) === null || _a === void 0 ? void 0 : _a.Id) === projectId && f.Id !== editingId; })) {
                    toast("error", "SEIBIES forecast already exists for this project!");
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
                try {
                    const forecastData = {};
                    const allStorageMonths = getAllMonthsForStorage();
                    allStorageMonths.forEach(monthKey => {
                        const mainCurrentInput = document.querySelector(`input.main-current-input-seibies[data-month="${monthKey}"]`);
                        const mainPlanInput = document.querySelector(`input.main-plan-input-seibies[data-month="${monthKey}"]`);
                        const mainForecastInput = document.querySelector(`input.main-forecast-input-seibies[data-month="${monthKey}"]`);
                        const mainDateInput = document.querySelector(`input.forecast-date-input-seibies[data-month="${monthKey}"]`);
                        if (mainCurrentInput && mainPlanInput) {
                            const monthData = {
                                plan: getNumericValue(mainPlanInput),
                                actual: getNumericValue(mainCurrentInput)
                            };
                            if (mainForecastInput) {
                                monthData.forecast = getNumericValue(mainForecastInput);
                            }
                            if (mainDateInput && mainDateInput.value) {
                                monthData.date = mainDateInput.value;
                            }
                            forecastData[monthKey] = monthData;
                        }
                    });
                    const yearlyData = {};
                    const yearlyPlanInputs = document.querySelectorAll('.yearly-plan-input-seibies');
                    if (yearlyPlanInputs.length > 0) {
                        const years = new Set();
                        yearlyPlanInputs.forEach(input => {
                            const year = input.dataset.year;
                            if (year)
                                years.add(year);
                        });
                        Array.from(years).sort().forEach(year => {
                            const planInput = document.querySelector(`input.yearly-plan-input-seibies[data-year="${year}"]`);
                            const currentInput = document.querySelector(`input.yearly-current-input-seibies[data-year="${year}"]`);
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
                        __metadata: { type: "SP.Data.SEIBIESListItem" },
                        Title: generateForecastTitle(project, propagatedData),
                        ProjectId: project.Id,
                        ProfitCentre: project.PC || "",
                        ForecastData: JSON.stringify(propagatedData),
                        UpdatedDate: new Date().toISOString()
                    };
                    if (!editingId && isNewProject) {
                        const oblValueToSave = getNumericValue(oblLastMonthInput);
                        payload.OBL = oblValueToSave;
                    }
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
                        toast("success", editingId ? "SEIBIES Forecast Updated!" : "SEIBIES Forecast Added!");
                        await loadForecasts();
                        initializePagination(forecasts);
                        hideForm();
                        resetForm();
                    }
                    else {
                        const errorText = await res.text();
                        console.error("SharePoint Error:", errorText);
                        toast("error", `Error saving SEIBIES forecast: ${res.status}`);
                    }
                }
                catch (error) {
                    console.error("Error saving SEIBIES forecast:", error);
                    toast("error", "Error saving SEIBIES forecast!");
                }
            });
            // ========== SET FORM READONLY ==========
            const setFormReadOnly = (isReadOnly) => {
                const editableSelectors = [
                    '.main-current-input-seibies',
                    '.yearly-current-input-seibies',
                    '.main-forecast-input-seibies',
                    '.forecast-date-input-seibies',
                    '.main-plan-input-seibies',
                    '.yearly-plan-input-seibies'
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
                const oblInput = document.querySelector('#oblLastMonthSeibies');
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
                // Show toast notification
                if (isReadOnly) {
                    toast("info", '🔒 This project is read-only.');
                }
                else {
                    toast("info", '🔓 Project is now editable.');
                }
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
                            actual: 0,
                        };
                    }
                });
                updatedData.yearly = updatedYearlyData;
                return updatedData;
            };
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
                // ✅ Use getAssignedProjects for all filter dropdowns
                // GH Heads will see ALL projects
                const assignedProjects = getAssignedProjects(projects);
                const assignedProjectTitles = assignedProjects.map(p => p.Title);
                // Project filter - show assigned projects (GH Heads see all)
                const forecastProjectTitles = forecasts
                    .map(f => { var _a; return (_a = f.Project) === null || _a === void 0 ? void 0 : _a.Title; })
                    .filter((title) => !!title)
                    .filter(title => assignedProjectTitles.includes(title));
                filterProject.innerHTML = '<option value="">All Projects</option>' +
                    [...new Set(forecastProjectTitles)]
                        .sort()
                        .map(project => `<option value="${project}">${project}</option>`)
                        .join("");
                const forecastProjects = projects.filter(p => forecastProjectTitles.includes(p.Title));
                const forecastCustomers = forecastProjects
                    .map(p => p.CustomerName)
                    .filter((customer) => !!customer);
                filterCustomer.innerHTML = '<option value="">All Customers</option>' +
                    [...new Set(forecastCustomers)]
                        .sort()
                        .map(customer => `<option value="${customer}">${customer}</option>`)
                        .join("");
                const forecastManagers = forecastProjects
                    .map(p => { var _a; return typeof p.ProjectManager === "object" ? (_a = p.ProjectManager) === null || _a === void 0 ? void 0 : _a.Title : p.ProjectManager; })
                    .filter((manager) => !!manager);
                filterProjectManager.innerHTML = '<option value="">All Managers</option>' +
                    [...new Set(forecastManagers)]
                        .sort()
                        .map(manager => `<option value="${manager}">${manager}</option>`)
                        .join("");
                const forecastGroupHeads = forecastProjects
                    .map(p => p.GroupHead)
                    .filter((head) => !!head);
                filterGroupHead.innerHTML = '<option value="">All Group Heads</option>' +
                    [...new Set(forecastGroupHeads)]
                        .sort()
                        .map(head => `<option value="${head}">${head}</option>`)
                        .join("");
                const forecastReferences = forecastProjects
                    .map(p => p.ProjectReference)
                    .filter((ref) => !!ref);
                filterProjectReference.innerHTML = '<option value="">All References</option>' +
                    [...new Set(forecastReferences)]
                        .sort()
                        .map(ref => `<option value="${ref}">${ref}</option>`)
                        .join("");
            };
            const filterTable = () => {
                const year = filterYear.value;
                const month = filterMonth.value;
                const project = filterProject.value;
                const customer = filterCustomer.value;
                const projectManager = filterProjectManager.value;
                const groupHead = filterGroupHead.value;
                const projectReference = filterProjectReference.value;
                // ✅ Use getAssignedProjects for permission-based filtering
                const assignedProjects = getAssignedProjects(projects);
                const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));
                let filtered = [...forecasts];
                // ✅ GH Heads see ALL projects (getAssignedProjects returns all)
                // Full Control see ALL projects
                // READ_ALL users see ALL projects (read-only)
                // PM Managers see only their projects
                // Engineering users see only their projects
                const isGHHeadUser = isGHHead();
                const isFullControlUser = hasFullControl();
                const isReadAllUser = getPermissionLevel() === 'read_all';
                // Only filter if user is NOT GH Head, NOT Full Control, and NOT READ_ALL
                if (!isGHHeadUser && !isFullControlUser && !isReadAllUser) {
                    filtered = filtered.filter(forecast => {
                        var _a;
                        const projectTitle = (_a = forecast.Project) === null || _a === void 0 ? void 0 : _a.Title;
                        return projectTitle && assignedProjectTitles.has(projectTitle);
                    });
                }
                // Apply additional filters
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
                    tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No SEIBIES records found for selected filters</td></tr>`;
                    const totalsHeader = document.querySelector("#totalsHeaderSeibies");
                    if (totalsHeader)
                        totalsHeader.innerHTML = '';
                    initializePagination([]);
                }
                else {
                    initializePagination(filtered);
                    const displayMonths = getDisplayMonths();
                    const currentMonthKey = formatMonthKey(new Date());
                    setTimeout(() => {
                        renderTotalsHeader(displayMonths, currentMonthKey);
                    }, 100);
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
                        if (checkIfNeedsProgression(originalData)) {
                            const progressedData = propagateCollectionValues(originalData);
                            const payload = {
                                __metadata: { type: "SP.Data.SEIBIESListItem" },
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
                    toast("success", `Monthly and yearly data progressed for ${progressedCount} SEIBIES forecasts`);
                }
                else if (errorCount > 0) {
                    toast("warning", `Some errors occurred during SEIBIES monthly and yearly progression`);
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
                const monthlyProgressionNeeded = firstDataMonth !== firstIdealMonth;
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
            // ========== INITIALIZE ==========
            // 8. In initialize - use all auth functions
            const initialize = async () => {
                try {
                    // ✅ Initialize shared authentication
                    await authManager.initialize({ siteUrl });
                    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                    console.log(`👤 User: ${getCurrentUserFullName()}`);
                    console.log(`📋 User Groups:`, getUserGroups());
                    console.log(`✅ Has Full Control: ${hasFullControl()}`);
                    console.log(`✅ Can Contribute: ${canContribute()}`);
                    console.log(`✅ Can Freeze: ${canFreeze()}`);
                    console.log(`✅ Is GH Head: ${isGHHead()}`);
                    console.log(`✅ Is PM Manager: ${isPMManager()}`); // ✅ NEW
                    console.log(`✅ Is Engineering User: ${isEngineeringUser()}`); // ✅ NEW
                    await loadProjects();
                    await loadForecasts();
                    if (forecasts.length > 0) {
                        await applyAutoProgression();
                        await loadForecasts();
                    }
                    initializeFilters();
                    setupPaginationEventListeners();
                    // ✅ Call filterTable() to apply project filtering
                    filterTable();
                    // ✅ Update UI based on permissions
                    updateUIBasedOnPermissions({
                        exportBtn: '.export-btn-seibies',
                        addBtn: '.show-form-btn-seibies',
                        saveBtn: '.btn-primary-seibies[type="submit"]'
                    });
                    console.log("✅ SEIBIES application initialized successfully");
                }
                catch (error) {
                    console.error("Failed to initialize SEIBIES application:", error);
                    toast("error", "Failed to initialize SEIBIES application");
                }
            };
            // ========== EVENT LISTENERS ==========
            if (showFormBtn) {
                showFormBtn.addEventListener("click", () => {
                    resetForm();
                    showForm(false);
                });
            }
            if (backBtn)
                backBtn.addEventListener("click", hideForm);
            if (projectSelect) {
                projectSelect.addEventListener("change", (e) => {
                    if (editingId) {
                        toast("warning", "Cannot change project in SEIBIES edit mode");
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
                                    title: "Existing SEIBIES Forecast Found",
                                    text: `A SEIBIES forecast already exists for ${project.Title}. Would you like to edit it?`,
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
            }
            if (firstPageBtn)
                firstPageBtn.addEventListener("click", () => goToPage(1));
            if (prevPageBtn)
                prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
            if (nextPageBtn)
                nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
            if (lastPageBtn)
                lastPageBtn.addEventListener("click", () => goToPage(totalPages));
            if (filterYear)
                filterYear.addEventListener("change", filterTable);
            if (filterMonth)
                filterMonth.addEventListener("change", filterTable);
            if (filterProject)
                filterProject.addEventListener("change", filterTable);
            if (filterCustomer)
                filterCustomer.addEventListener("change", filterTable);
            if (filterProjectManager)
                filterProjectManager.addEventListener("change", filterTable);
            if (filterGroupHead)
                filterGroupHead.addEventListener("change", filterTable);
            if (filterProjectReference)
                filterProjectReference.addEventListener("change", filterTable);
            if (exportBtn)
                exportBtn.addEventListener("click", exportToExcelWithFullStyle);
            if (clearFiltersBtn)
                clearFiltersBtn.addEventListener("click", clearFilters);
            initialize();
        }, 100);
    }
};
//# sourceMappingURL=seibies.js.map
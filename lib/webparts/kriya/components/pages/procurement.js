import Swal from "sweetalert2";
import "sweetalert2/dist/sweetalert2.min.css";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/procurement.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canReadAllProjects, getAssignedProjects, isPMManager, canCreateForProject, isEngineeringUser, updateUIBasedOnPermissions } from "./auth";
// interface ValidationError {
//   type: 'error' | 'warning' | 'info';
//   message: string;
//   row?: number;
//   column?: string;
// }
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
    <!-- ✅ PAGE HEADER - UNCHANGED (title LEFT, "+ Add Procurement" RIGHT).
         Same markup as before: <h1> + the .page-header-left wrapper that
         styles/procurement.css positions on the right (order:2; margin-left:auto).
         Nothing here is re-styled by the upgrade - the ONLY change is the button
         FORMAT, because the button keeps its original .btn-primary class and that
         is now the shared lilac pill recipe. -->
    <div class="page-header stick">
      <div class="page-header-left">
        <button class="btn-primary show-form-btn">+ New</button>
      </div>
      <h1>Procurement Plan/BOM</h1>
    </div>

    <!-- =====================================================================
         ✅ TOOLBAR ABOVE THE TABLE (mock-up image 8)
         LEFT  = Filters (funnel icon + active count badge) + the compact
                 "Apply" pill that appears when the popup has unapplied changes
         RIGHT = Download Excel
         There is NO Reset button here - the popup's Clear button resets every
         filter (the old .table-actions row of dropdowns + red Reset is gone).
         The five dropdowns moved into the Filters popup below and keep their
         original ids (#filterProject, #filterCustomer, #filterProjectManager,
         #filterGroupHead, #filterProjectReference), so initializeFilters(),
         filterData() and the Excel export keep working unchanged.
         ===================================================================== -->
    <div class="proc-toolbar">
      <div class="proc-toolbar__left">
        <button type="button" class="btn-secondary proc-filter-trigger" id="procFilterTrigger"
                aria-haspopup="dialog" aria-expanded="false" title="Filter Procurement">
          <svg class="proc-filter-funnel" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8v5l-4 2v-7L3 5z"/></svg>
          <span class="proc-filter-label">Filters</span>
          <span class="proc-filter-count" id="procFilterCount">0</span>
          <svg class="proc-filter-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>
        </button>

        <!-- compact Apply pill - shown only while the popup holds changes that
             have not been applied yet -->
        <button type="button" class="btn-primary proc-filter-applypill" id="procFloatingApply">Apply</button>
      </div>

      <div class="proc-toolbar__right">
        <button type="button" class="btn-secondary export-excel-btn">📊 Download Excel</button>
      </div>
    </div>

    <!-- chips summarising the filters that are currently applied -->
   <!--  <div class="proc-filter-summary" id="procFilterSummary" style="display:none;"></div>-->

    <!-- ✅ Main table: OUTER-STICKY header -->
    <!--    The page scrolls vertically (not an inner box), so the header  -->
    <!--    runs up, touches the main header and pins below it.             -->
    <div class="table-scroll-wrapper main-table-scroll">
      <table id="procurementTable" class="modern-table">
        <thead class="sticky-main-header">
          <tr>
            <th class="sticky-main-col main-col-1 proc-sortable-th" data-sort="reference" title="Sort by Project Reference"><span class="proc-th-inner"><span class="proc-th-text">Project Reference</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="reference" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="reference" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="sticky-main-col main-col-2 proc-sortable-th" data-sort="customer" title="Sort by Customer"><span class="proc-th-inner"><span class="proc-th-text">Customer</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="customer" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="customer" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="sticky-main-col main-col-3 proc-sortable-th" data-sort="project" title="Sort by Project"><span class="proc-th-inner"><span class="proc-th-text">Project</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="project" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="project" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="sticky-main-col main-col-4 proc-sortable-th" data-sort="grouphead" title="Sort by Group Head"><span class="proc-th-inner"><span class="proc-th-text">Group Head</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="grouphead" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="grouphead" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="sticky-main-col main-col-5 proc-sortable-th" data-sort="manager" title="Sort by Project Manager"><span class="proc-th-inner"><span class="proc-th-text">Project Manager</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="manager" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="manager" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="total" title="Sort by Total Line Items"><span class="proc-th-inner"><span class="proc-th-text">Total Line Items</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="total" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="total" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="rfq" title="Sort by RFQ Pending"><span class="proc-th-inner"><span class="proc-th-text">RFQ Pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="rfq" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="rfq" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="offer" title="Sort by Offer Pending"><span class="proc-th-inner"><span class="proc-th-text">Offer Pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="offer" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="offer" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="tc" title="Sort by TC pending"><span class="proc-th-inner"><span class="proc-th-text">TC pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="tc" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="tc" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="orderclosure" title="Sort by Order Closure pending"><span class="proc-th-inner"><span class="proc-th-text">Order Closure pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="orderclosure" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="orderclosure" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="pr" title="Sort by PR pending"><span class="proc-th-inner"><span class="proc-th-text">PR pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="pr" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="pr" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="po" title="Sort by PO pending"><span class="proc-th-inner"><span class="proc-th-text">PO pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="po" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="po" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="docapproval" title="Sort by Document Approval Pending"><span class="proc-th-inner"><span class="proc-th-text">Document Approval Pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="docapproval" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="docapproval" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="mfc" title="Sort by MFC Pending"><span class="proc-th-inner"><span class="proc-th-text">MFC Pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="mfc" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="mfc" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="inspection" title="Sort by Inspection Pending"><span class="proc-th-inner"><span class="proc-th-text">Inspection Pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="inspection" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="inspection" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="dispatch" title="Sort by Dispatch Pending"><span class="proc-th-inner"><span class="proc-th-text">Dispatch Pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="dispatch" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="dispatch" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="delivery" title="Sort by Delivery Pending"><span class="proc-th-inner"><span class="proc-th-text">Delivery Pending</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="delivery" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="delivery" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="delivered" title="Sort by Delivered"><span class="proc-th-inner"><span class="proc-th-text">Delivered</span><span class="proc-sort-buttons-container"><button type="button" class="proc-sort-btn proc-sort-asc" data-sort="delivered" data-direction="asc" title="Sort Ascending">&#9650;</button><button type="button" class="proc-sort-btn proc-sort-desc" data-sort="delivered" data-direction="desc" title="Sort Descending">&#9660;</button></span></span></th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody id="procurementBody"></tbody>
      </table>
    </div>

    <!-- ✅ Loading state: the same indeterminate BAR as the other modules
         (grey track + indigo bar, no spinner, no text). It sits in the same
         place as before - below the table, above the pagination bar. -->
    <div class="proc-loading" id="procurementLoading">
      <div class="proc-page-loader">
        <div class="proc-page-loader__track">
          <div class="proc-page-loader__bar"></div>
        </div>
      </div>
    </div>

    <!-- ✅ ONE-LINE PAGINATION BAR (mock-up layout)
         LEFT  = Items per page: [n]   Showing 1 - 25 of 60 items
         RIGHT = «   ‹   [page tiles] of N   ›   »
         Every id used by the script is unchanged (#procFirstPage, #procPrevPage,
         #procNextPage, #procLastPage, #procTotalPages, #procPageNumbers,
         #procStartItem, #procEndItem, #procTotalItems, #pageSizeSelect) - the
         "Page X of Y" pill is gone, it is now the "of N" after the tiles. -->
    <div class="proc-pg-container">
      <div class="proc-pg-left">
        <div class="proc-pg-size">
          <label for="pageSizeSelect">Items per page:</label>
          <select id="pageSizeSelect" class="proc-pg-size-select">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>

        <div class="proc-pg-total">
          Showing <span id="procStartItem">0</span> - <span id="procEndItem">0</span> of <span id="procTotalItems">0</span> items
        </div>
      </div>

      <div class="proc-pg-right">
        <button type="button" class="proc-pg-icon" id="procFirstPage" title="First page">«</button>
        <button type="button" class="proc-pg-icon" id="procPrevPage" title="Previous page">‹</button>

        <div class="proc-pg-pages" id="procPageNumbers"></div>

        <span class="proc-pg-of">of <span id="procTotalPages">1</span></span>

        <button type="button" class="proc-pg-icon" id="procNextPage" title="Next page">›</button>
        <button type="button" class="proc-pg-icon" id="procLastPage" title="Last page">»</button>
      </div>
    </div>
  </section>

  <section class="form-section card" style="display:none;">
  <form id="procurementForm">
  <div class="page-header">
    <h1 id="formTitle">Procurement Plan/BOM</h1>
    <!-- ✅ TOP SAVE BUTTON - same type="submit" + .btn-primary
         as the bottom button, so both trigger the same submit handler
         and stay in sync (visibility, saving state). The header is
         position:sticky, so this button stays visible while scrolling
         through the long procurement table. -->
        
    <button type="button"
        class="btn-primary proc-btn-save-top proc-save-btn"
        title="Save Procurement Plan">💾 Save</button>
         <button type="button" class="btn-secondary back-to-table">⬅ Back</button>
  </div>

      <!-- Project Info Section -->
      <div class="project-info-wrapper">
        <div class="project-info-left">
          <div class="detail-item project-select-container">
            <span class="detail-label">Select Project</span>
            <select id="projectSelect" class="detail-value project-select" required></select>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project Reference</span>
            <span class="detail-value" id="projectRefDisplay">-</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Project Center</span>
            <span class="detail-value" id="projectCenterDisplay">-</span>
          </div>
        </div>
        <div class="project-info-right">
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
        </div> 
      </div>

      <!-- Procurement Table -->
      <div class="procurement-inputs" id="procurementContainerTable" style="display:none;">
        <!-- 🔍 Table search bar (replaces the old "Filter Description..." box
             that used to sit inside the Description header cell).
             It searches EVERY column of the line-item table below. -->
        <div class="proc-table-toolbar">
          <div class="proc-search-wrapper">
            <span class="proc-search-icon">🔍</span>
            <input
              type="text"
              id="procTableSearch"
              class="proc-table-search"
              placeholder="Search all columns (description, status, make, model, dates, remarks)..."
              aria-label="Search all columns of the procurement table"
              autocomplete="off"
            >
            <button type="button" class="proc-search-clear" id="procTableSearchClear" title="Clear search" aria-label="Clear search">&times;</button>
          </div>
          <span class="proc-search-count" id="procTableSearchCount"></span>
        </div>

        <div class="table-container">
          <table class="procurement-table">
            <thead>
              <tr>
                <th class="sticky-col col-1" rowspan="2">Description</th>
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
                <td colspan="35"></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      <div class="form-actions">
       <button type="button" class="btn-primary proc-save-btn">💾 Save</button>
        <button type="button" class="btn-secondary back-to-table">⬅ Back</button>
      </div>
    </form>
  </section>

  <!-- =====================================================================
       ✅ FILTERS POPUP (mock-up image 7): header · fields · Cancel / Clear / Apply
       The five dropdowns keep their ORIGINAL ids, so initializeFilters(),
       filterData(), clearFilters() and the Excel export need no change.
       Clear = the Reset action (it lives here, not in the toolbar).
       ===================================================================== -->
  <div class="proc-filter-overlay" id="procFilterOverlay" role="dialog" aria-modal="true" aria-labelledby="procFilterTitle">
    <div class="proc-filter-modal">
      <div class="proc-filter-modal__head">
        <h2 id="procFilterTitle">Filters</h2>
        <button type="button" class="proc-filter-modal__close" id="procFilterClose" title="Close" aria-label="Close">✕</button>
      </div>

      <div class="proc-filter-modal__body">
        <div class="proc-filter-grid">
          <div class="proc-filter-field proc-filter-field--wide">
            <label for="filterProject">Project</label>
            <select id="filterProject"><option value="">All Projects</option></select>
          </div>
          <div class="proc-filter-field">
            <label for="filterCustomer">Customer</label>
            <select id="filterCustomer"><option value="">All Customers</option></select>
          </div>
          <div class="proc-filter-field">
            <label for="filterProjectManager">Project Manager</label>
            <select id="filterProjectManager"><option value="">All Managers</option></select>
          </div>
          <div class="proc-filter-field">
            <label for="filterGroupHead">Group Head</label>
            <select id="filterGroupHead"><option value="">All Group Heads</option></select>
          </div>
          <div class="proc-filter-field">
            <label for="filterProjectReference">Project Reference</label>
            <select id="filterProjectReference"><option value="">All References</option></select>
          </div>
        </div>
      </div>

      <div class="proc-filter-modal__foot">
        <button type="button" class="proc-filter-btn-text" id="procFilterCancel">Cancel</button>
        <button type="button" class="proc-filter-btn-clear" id="procFilterClear" title="Reset all filters" disabled>Clear</button>
        <button type="button" class="proc-filter-btn-apply" id="procFilterApply" disabled>Apply</button>
      </div>
    </div>
  </div>
</div>
`;
export default {
    content: procurementContent,
    script: () => {
        setTimeout(async () => {
            if (window.__dtActivePageKey !== 'procurement')
                return;
            // Configuration
            const siteUrl = typeof _spPageContextInfo !== "undefined"
                ? _spPageContextInfo.webAbsoluteUrl
                : window.location.origin + "/sites/Kriya";
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
            let currentTableData = []; // data currently shown (respects filters) - used by pagination
            let currentPage = 1;
            let pageSize = 25; // ✅ default 25 per page (same as DCS)
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
            // OVERDUE ACTUAL DATE (Plan crossed, Actual empty)
            // -> red border on the Actual date input.
            // Border is removed as soon as Actual is filled.
            // ============================================ //
            const STAGE_PREFIXES = [
                "rfq",
                "offer",
                "tech",
                "order",
                "pr",
                "po",
                "doc",
                "mfg",
                "insp",
                "disp",
                "delivery",
            ];
            const getTodayISO = () => {
                const now = new Date();
                return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
            };
            // ✅ A user may edit/save a project when they have explicit edit rights
            //    OR when the project is assigned to them (own-project roles) OR they
            //    are a read-all role that is allowed to edit.
            const canEditOrAssignedProject = (project) => {
                if (!project)
                    return false;
                if (canEditProject(project))
                    return true;
                if (canReadAllProjects() && canContribute())
                    return true; // read-all roles that may contribute
                return getAssignedProjects(projects).some((p) => p.Id === project.Id);
            };
            // Plan date is "crossed" when it is strictly before today
            const isPlanCrossed = (planValue) => {
                if (!isFilled(planValue))
                    return false;
                const plan = new Date(planValue);
                if (isNaN(plan.getTime()))
                    return false;
                const planISO = `${plan.getFullYear()}-${String(plan.getMonth() + 1).padStart(2, "0")}-${String(plan.getDate()).padStart(2, "0")}`;
                return planISO < getTodayISO();
            };
            const isActualOverdue = (planValue, actualValue) => isPlanCrossed(planValue) && !isFilled(actualValue);
            const updateOverdueHighlights = (row) => {
                STAGE_PREFIXES.forEach((prefix) => {
                    const planInput = row.querySelector(`input[name="${prefix}_plan"]`);
                    const actualInput = row.querySelector(`input[name="${prefix}_actual"]`);
                    if (!actualInput)
                        return;
                    const overdue = isActualOverdue((planInput === null || planInput === void 0 ? void 0 : planInput.value) || "", actualInput.value || "");
                    if (overdue) {
                        actualInput.classList.add("actual-overdue");
                        actualInput.title = `Plan date (${(planInput === null || planInput === void 0 ? void 0 : planInput.value) || ""}) has passed - please update the Actual date`;
                    }
                    else {
                        actualInput.classList.remove("actual-overdue");
                        actualInput.removeAttribute("title");
                    }
                });
            };
            const wireOverdueHighlights = (row) => {
                STAGE_PREFIXES.forEach((prefix) => {
                    ["plan", "actual"].forEach((kind) => {
                        const input = row.querySelector(`input[name="${prefix}_${kind}"]`);
                        if (!input)
                            return;
                        input.addEventListener("change", () => updateOverdueHighlights(row));
                        input.addEventListener("input", () => updateOverdueHighlights(row));
                    });
                });
                updateOverdueHighlights(row);
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
                wireOverdueHighlights(row);
            };
            // DOM Elements
            const projectSelect = document.getElementById("projectSelect");
            const showFormBtn = document.querySelector(".show-form-btn");
            const exportExcelBtn = document.querySelector(".export-excel-btn");
            const formSection = document.querySelector(".form-section");
            const tableSection = document.querySelector(".table-section");
            // const backBtn = document.querySelector(".back-to-table") as HTMLButtonElement;
            const procurementContainer = document.getElementById("procurementContainer");
            const procurementContainerTable = document.getElementById("procurementContainerTable");
            const tableBody = document.getElementById("procurementBody");
            // const form = document.getElementById("procurementForm") as HTMLFormElement;
            const formTitle = document.getElementById("formTitle");
            const existingDataAlert = document.getElementById("existingDataAlert");
            // ✅ Search box for the Edit / View Procurement table.
            // (Moved OUT of the Description header into its own search bar above the
            // table - it searches EVERY column, not just Description.)
            const tableSearchInput = document.getElementById("procTableSearch");
            const tableSearchClearBtn = document.getElementById("procTableSearchClear");
            const tableSearchCount = document.getElementById("procTableSearchCount");
            // Project Info Display Elements
            const projectRefDisplay = document.getElementById("projectRefDisplay");
            const projectCenterDisplay = document.getElementById("projectCenterDisplay");
            const projectManagerDisplay = document.getElementById("projectManagerDisplay");
            const groupHeadDisplay = document.getElementById("groupHeadDisplay");
            const customerDisplay = document.getElementById("customerDisplay");
            // Filter Elements (they now live inside the Filters popup)
            const filterProject = document.getElementById("filterProject");
            const filterCustomer = document.getElementById("filterCustomer");
            const filterProjectManager = document.getElementById("filterProjectManager");
            const filterGroupHead = document.getElementById("filterGroupHead");
            const filterProjectReference = document.getElementById("filterProjectReference");
            // ✅ The Reset/clear action now lives INSIDE the Filters popup (Clear),
            //    so the old toolbar button - and its reference - are gone.
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
            // ============================================
            // LOADING OVERLAY - BAR ONLY (same as Inventory / SEIBIES / Invoicing)
            //   createLoadingOverlay(message) returns the element to append and
            //   removeLoadingOverlay() cleans up, so every call site keeps working.
            //   The circular spinner + "Loading ..." text are gone: the grey track +
            //   indigo bar (styled in styles/procurement.css) are the only thing on
            //   screen, and the step message is logged to the console instead.
            // ============================================
            const createLoadingOverlay = (message = "Processing...") => {
                removeLoadingOverlay(); // never stack two loaders
                if (message)
                    console.log(`⏳ ${message}`);
                const overlay = document.createElement("div");
                overlay.className = "loading-overlay proc-loading-overlay";
                overlay.innerHTML = `
          <div class="proc-loader-panel">
            <div class="proc-page-loader">
              <div class="proc-page-loader__track">
                <div class="proc-page-loader__bar"></div>
              </div>
            </div>
          </div>`;
                return overlay;
            };
            const removeLoadingOverlay = () => {
                document.querySelectorAll(".loading-overlay, .proc-loading-overlay").forEach(el => el.remove());
            };
            /** mount the loader on the VISIBLE card (table section or form section) —
                the table card while data loads, the form card while saving */
            const loaderHost = () => {
                const sections = Array.from(document.querySelectorAll('.table-section, .form-section'));
                const visible = sections.find(s => s.offsetParent !== null);
                return visible || document.querySelector('.table-section') || document.body;
            };
            // ============================================ //
            // ROLE-BASED ACCESS (same as SEIBIES / auth.ts) //
            //   • a role may OPEN an EXISTING procurement    //
            //     plan when it can edit it, when it is a     //
            //     READ_ALL role (opens read-only if not      //
            //     editable), or when the project is assigned //
            //     to it.                                     //
            //   • creating a NEW plan requires create rights //
            //     (FULL_CONTROL / CONTRIBUTE).               //
            // ============================================ //
            const canViewProcurementForm = (isEdit = false, project) => {
                if (!canRead())
                    return false;
                if (isEdit) {
                    return canEditOrAssignedProject(project);
                }
                // Create path — "+ New" runs BEFORE a project is selected, so `project`
                // may be undefined. Allow any role that can create for at least one
                // project; the strict per-project gate runs later in onProjectChange.
                return canCreateRecord() || getAssignedProjects(projects).length > 0;
            };
            // Show / hide the Save button based on the current role + mode
            const applySaveButtonVisibility = (isEditMode, project) => {
                const canSave = isEditMode
                    ? canEditOrAssignedProject(project) // ✅ assigned projects can save
                    : canCreateRecord() || !!project && canEditOrAssignedProject(project);
                getSaveButtons().forEach(btn => {
                    btn.style.display = canSave ? "inline-flex" : "none";
                });
            };
            const getSaveButtons = () => Array.from(document.querySelectorAll(".proc-save-btn"));
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
                const systemNames = [
                    "Id",
                    "ID",
                    "Title",
                    "Created",
                    "Modified",
                    "Author",
                    "Editor",
                ];
                // Only real, user-visible, writable columns are candidates
                const candidates = (fields || []).filter((field) => {
                    const internalName = field.InternalName || "";
                    if (!internalName || internalName.startsWith("_"))
                        return false;
                    if (systemNames.includes(internalName))
                        return false;
                    if (field.Hidden === true)
                        return false;
                    if (field.ReadOnlyField === true)
                        return false;
                    if (field.FromBaseType === true)
                        return false;
                    return true;
                });
                const hasExact = (internal, type) => candidates.some((f) => f.InternalName === internal && f.TypeAsString === type);
                // --- ProcurementForecast (Note / multi-line text) ---
                if (!hasExact("ProcurementForecast", "Note")) {
                    const noteField = candidates.find((field) => {
                        if (field.TypeAsString !== "Note")
                            return false;
                        const t = (field.Title || "").toLowerCase().replace(/\s+/g, "");
                        const n = (field.InternalName || "").toLowerCase();
                        return (t.includes("procurement") ||
                            n.includes("procurement") ||
                            t.includes("forecast") ||
                            n.includes("forecast") ||
                            t.includes("data"));
                    });
                    if (noteField)
                        mappings.ProcurementForecast = noteField.InternalName;
                }
                // --- Project (Lookup) ---
                if (!hasExact("Project", "Lookup")) {
                    const lookupField = candidates.find((field) => field.TypeAsString === "Lookup" &&
                        (field.Title || "").toLowerCase().replace(/\s+/g, "") ===
                            "project") ||
                        candidates.find((field) => {
                            if (field.TypeAsString !== "Lookup")
                                return false;
                            const t = (field.Title || "").toLowerCase().replace(/\s+/g, "");
                            const n = (field.InternalName || "").toLowerCase();
                            return t.includes("project") || n.includes("project");
                        });
                    if (lookupField)
                        mappings.Project = lookupField.InternalName;
                }
                console.log("Procurement field mappings:", mappings);
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
                        let errorText = "";
                        try {
                            errorText = await testRes.text();
                        }
                        catch (_a) {
                            /* ignore */
                        }
                        console.error(`Field mappings test: FAILED (${testRes.status}) for query: ${testQuery}`, errorText);
                        return false;
                    }
                }
                catch (error) {
                    console.error("Field mappings test error:", error);
                    return false;
                }
            };
            const renderProjectDetails = (project) => {
                if (!project) {
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
                const managerName = getProjectManagerDisplayName(project.ProjectManager);
                const groupHeadName = getGroupHeadDisplayName(project.GroupHead1);
                if (projectRefDisplay)
                    projectRefDisplay.textContent = project.ProjectReference || "-";
                if (projectCenterDisplay)
                    projectCenterDisplay.textContent = project.PC || "-";
                if (projectManagerDisplay)
                    projectManagerDisplay.textContent = managerName || "-";
                if (groupHeadDisplay)
                    groupHeadDisplay.textContent = groupHeadName || "-";
                if (customerDisplay)
                    customerDisplay.textContent = project.CustomerName || "-";
            };
            const initApp = async () => {
                try {
                    const initialLoading = createLoadingOverlay("Loading Procurement data...");
                    loaderHost().appendChild(initialLoading);
                    // ✅ Initialize shared authentication (same as SEIBIES)
                    await authManager.initialize({ siteUrl });
                    console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
                    console.log(`👤 User: ${getCurrentUserFullName()}`);
                    console.log(`📋 User Groups:`, getUserGroups());
                    console.log(`✅ Has Full Control: ${hasFullControl()}`);
                    console.log(`✅ Can Contribute: ${canContribute()}`);
                    console.log(`✅ Can Freeze: ${canFreeze()}`);
                    console.log(`✅ Is GH Head: ${isGHHead()}`);
                    console.log(`✅ Is PM Manager: ${isPMManager()}`);
                    console.log(`✅ Is Engineering User: ${isEngineeringUser()}`);
                    const { fields, entityType } = await discoverListFields();
                    listEntityTypeName = entityType;
                    fieldMappings = createFieldMappings(fields);
                    const mappingsValid = await testFieldMappings();
                    if (!mappingsValid) {
                        throw new Error("Field mappings are invalid. Please check the field names in your SharePoint list.");
                    }
                    await loadProjects();
                    const dataLoaded = await fetchData();
                    if (!dataLoaded)
                        return;
                    // toast("Procurement plan initialized successfully", "success");
                    // ✅ Apply role-based UI visibility (same as SEIBIES)
                    updateUIBasedOnPermissions({
                        exportBtn: ".export-excel-btn",
                        addBtn: ".show-form-btn",
                        saveBtn: ".proc-save-btn" // ✅ matches your current buttons
                    });
                }
                catch (error) {
                    console.error("Initialization error:", error);
                    const message = error instanceof Error ? error.message : String(error);
                    showTableMessage(`Failed to load data: ${message}`);
                    toast(`Failed to initialize application: ${message}`, "error");
                    hideProcurementLoading();
                }
                finally {
                    removeLoadingOverlay();
                }
            };
            // Shows a message row inside the main table (so an error is never silent)
            const showTableMessage = (message) => {
                if (!tableBody)
                    return;
                tableBody.innerHTML = `<tr><td colspan="19" style="text-align:center;padding:20px;color:#c62828;">${escapeHtml(message)}</td></tr>`;
            };
            const loadProjects = async () => {
                var _a;
                try {
                    const listPath = `${siteUrl}/_api/web/lists/getbytitle('Project List')/items`;
                    const headers = { Accept: "application/json;odata=verbose" };
                    // ✅ $select + $expand must include every field ROLE_OWN_FIELDS in auth.ts
                    // matches against, otherwise getAssignedProjects() returns [] for the
                    // KRIYA_Engg_* roles.
                    const fullQuery = `${listPath}?$select=Id,Title,CustomerName,PC,ProjectReference,` +
                        `GroupHead1/Title,GroupHead1Id,` +
                        `ProjectManager/Title,ProjectManagerId,` +
                        `EnggLead/Title,EnggLeadId,` +
                        `EnggSubLead1/Title,EnggSubLead1Id,` +
                        `EnggSubLead2/Title,EnggSubLead2Id,` +
                        `EngineeringManagers/Title,EngineeringManagersId` +
                        `&$expand=GroupHead1,ProjectManager,EnggLead,EnggSubLead1,EnggSubLead2,EngineeringManagers` +
                        `&$top=5000`;
                    let res = await fetch(fullQuery, { headers });
                    // Fallback 1 - engineering fields as IDs only (no $expand)
                    // Still carries the Engg*Id columns so we can resolve the names
                    // via getuserbyid just like GroupHead1Id / ProjectManagerId.
                    let rawProjects = [];
                    if (!res.ok) {
                        console.warn(`Project List expanded query failed (${res.status}); retrying with ID-only query`);
                        const idQuery = `${listPath}?$select=Id,Title,CustomerName,PC,ProjectReference,` +
                            `GroupHead1Id,ProjectManagerId,` +
                            `EnggLeadId,EnggSubLead1Id,EnggSubLead2Id,EngineeringManagersId` +
                            `&$top=5000`;
                        res = await fetch(idQuery, { headers });
                    }
                    // Fallback 2 - bare fetch, then filter out hidden/system columns client-side
                    if (!res.ok) {
                        console.warn(`Project List ID query failed (${res.status}); retrying with full fetch`);
                        res = await fetch(`${listPath}?$top=5000`, { headers });
                    }
                    if (!res.ok)
                        throw new Error(`Failed to load projects: ${res.status}`);
                    const data = await res.json();
                    rawProjects = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    // Memoised user-id → name resolver (used by the fallback paths)
                    const userCache = new Map();
                    const getUserName = async (id) => {
                        var _a;
                        if (!id)
                            return "";
                        if (userCache.has(id))
                            return userCache.get(id);
                        try {
                            const uRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${id})?$select=Id,Title,Email`, { headers });
                            if (uRes.ok) {
                                const uData = await uRes.json();
                                const name = ((_a = uData.d) === null || _a === void 0 ? void 0 : _a.Title) || "";
                                userCache.set(id, name);
                                return name;
                            }
                        }
                        catch (e) {
                            console.warn(`Could not fetch user ${id}:`, e);
                        }
                        userCache.set(id, "");
                        return "";
                    };
                    // Field name resolver — works whether the response uses the expanded
                    // shape ({ Title }) or the flat ID shape (EnngLeadId + no object).
                    const nameOf = (val) => {
                        if (!val)
                            return "";
                        if (typeof val === "object")
                            return val.Title || "";
                        return String(val);
                    };
                    const mappedProjects = await Promise.all(rawProjects.map(async (p) => {
                        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m;
                        return ({
                            Id: p.Id,
                            Title: p.Title || "",
                            CustomerName: p.CustomerName || "",
                            PC: p.PC || "",
                            ProjectReference: p.ProjectReference || "",
                            GroupHead1: p.GroupHead1 ? nameOf(p.GroupHead1) : await getUserName(p.GroupHead1Id),
                            GroupHead1Id: (_b = (_a = p.GroupHead1) === null || _a === void 0 ? void 0 : _a.Id) !== null && _b !== void 0 ? _b : p.GroupHead1Id,
                            ProjectManager: p.ProjectManager ? nameOf(p.ProjectManager) : await getUserName(p.ProjectManagerId),
                            ProjectManagerId: (_d = (_c = p.ProjectManager) === null || _c === void 0 ? void 0 : _c.Id) !== null && _d !== void 0 ? _d : p.ProjectManagerId,
                            EnggLead: p.EnggLead ? nameOf(p.EnggLead) : await getUserName(p.EnggLeadId),
                            EnggLeadId: (_f = (_e = p.EnggLead) === null || _e === void 0 ? void 0 : _e.Id) !== null && _f !== void 0 ? _f : p.EnggLeadId,
                            EnggSubLead1: p.EnggSubLead1 ? nameOf(p.EnggSubLead1) : await getUserName(p.EnggSubLead1Id),
                            EnggSubLead1Id: (_h = (_g = p.EnggSubLead1) === null || _g === void 0 ? void 0 : _g.Id) !== null && _h !== void 0 ? _h : p.EnggSubLead1Id,
                            EnggSubLead2: p.EnggSubLead2 ? nameOf(p.EnggSubLead2) : await getUserName(p.EnggSubLead2Id),
                            EnggSubLead2Id: (_k = (_j = p.EnggSubLead2) === null || _j === void 0 ? void 0 : _j.Id) !== null && _k !== void 0 ? _k : p.EnggSubLead2Id,
                            EngineeringManagers: p.EngineeringManagers ? nameOf(p.EngineeringManagers) : await getUserName(p.EngineeringManagersId),
                            EngineeringManagersId: (_m = (_l = p.EngineeringManagers) === null || _l === void 0 ? void 0 : _l.Id) !== null && _m !== void 0 ? _m : p.EnggSubLead2Id
                        });
                    }));
                    const uniqueProjects = Array.from(new Map(mappedProjects.map((p) => [p.Title, p])).values());
                    projects = uniqueProjects;
                    console.log(`Loaded ${projects.length} projects from Project List`);
                    if (projects.length > 0) {
                        console.log('📋 Sample project data:', {
                            Title: projects[0].Title,
                            ProjectManager: projects[0].ProjectManager,
                            GroupHead1: projects[0].GroupHead1,
                            EnggLead: projects[0].EnggLead,
                            EnggSubLead1: projects[0].EnggSubLead1,
                            EnggSubLead2: projects[0].EnggSubLead2,
                            EngineeringManagers: projects[0].EngineeringManagers
                        });
                    }
                    // ✅ Role access (same as before): dropdowns only show assigned projects
                    const assignedProjects = getAssignedProjects(projects);
                    console.log(`📋 Total: ${projects.length}, Assigned: ${assignedProjects.length}`);
                    console.log('📋 Assigned titles:', assignedProjects.map((p) => p.Title));
                    projectSelect.innerHTML =
                        '<option value="">Select Project</option>' +
                            assignedProjects
                                .map((p) => `<option value="${p.Id}">${p.Title}</option>`)
                                .join("");
                    filterProject.innerHTML =
                        '<option value="">All Projects</option>' +
                            [...new Set(assignedProjects.map((p) => p.Title))]
                                .sort()
                                .map((project) => `<option value="${project}">${project}</option>`)
                                .join("");
                }
                catch (error) {
                    console.error("Error loading projects:", error);
                    toast("Failed to load projects", "error");
                }
            };
            // Helper functions for Person/Group display
            const getGroupHeadDisplayName = (groupHead) => {
                if (!groupHead)
                    return "";
                if (typeof groupHead === "object" && groupHead !== null) {
                    return groupHead.Title || "";
                }
                // ✅ cast: the object branch was handled above, so this is the plain string
                return groupHead;
            };
            const getProjectManagerDisplayName = (manager) => {
                if (!manager)
                    return "";
                if (typeof manager === "object" && manager !== null) {
                    return manager.Title || "";
                }
                // ✅ cast: the object branch was handled above, so this is the plain string
                return manager;
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
            // ✅ Loading state (same pattern as Inventory module):
            // the indeterminate BAR (grey track + indigo bar, no text) while records are fetched.
            const loadingIndicator = document.getElementById("procurementLoading");
            const showProcurementLoading = () => {
                if (loadingIndicator)
                    loadingIndicator.style.display = "flex";
            };
            const hideProcurementLoading = () => {
                if (loadingIndicator)
                    loadingIndicator.style.display = "none";
            };
            const fetchData = async () => {
                var _a;
                showProcurementLoading();
                const loading = createLoadingOverlay("Loading Procurement data...");
                loaderHost().appendChild(loading);
                try {
                    const headers = { Accept: "application/json;odata=verbose" };
                    const projectField = fieldMappings.Project;
                    const forecastField = fieldMappings.ProcurementForecast;
                    const selectFields = [
                        "Id",
                        "Title",
                        forecastField,
                        `${projectField}Id`,
                        `${projectField}/Id`,
                        `${projectField}/Title`,
                    ].filter((field) => field && !field.startsWith("_"));
                    const query = `${apiBase}?$select=${selectFields.join(",")}&$expand=${projectField}&$top=5000`;
                    let res = await fetch(query, { headers });
                    // ✅ Fallback: if the expanded query is rejected by SharePoint,
                    // load the raw items instead of showing an empty table.
                    if (!res.ok) {
                        let errorText = "";
                        try {
                            errorText = await res.text();
                        }
                        catch (_b) {
                            /* ignore */
                        }
                        console.warn(`Procurement query failed (${res.status}) for: ${query}`, errorText);
                        res = await fetch(`${apiBase}?$top=5000`, { headers });
                    }
                    if (!res.ok)
                        throw new Error(`HTTP error! status: ${res.status}`);
                    const data = await res.json();
                    procurementData = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    console.log(`Loaded ${procurementData.length} procurement records from ${listName}`);
                    // ✅ Role access (same as SEIBIES): a non-READ_ALL role only sees the
                    // procurement of projects assigned to it (GH Heads / PM Managers get ALL).
                    if (!canReadAllProjects() && projects.length > 0) {
                        const assignedProjectTitles = new Set(getAssignedProjects(projects).map((p) => p.Title));
                        procurementData = procurementData.filter((item) => {
                            var _a;
                            const projectData = resolveProjectForItem(item);
                            const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || (projectData === null || projectData === void 0 ? void 0 : projectData.Title) || "";
                            return projectTitle && assignedProjectTitles.has(projectTitle);
                        });
                        console.log(`Role-filtered to ${procurementData.length} procurement records`);
                    }
                    initializeFilters();
                    syncFilterUIFromSelects(); // toolbar badge / chips / Apply state
                    filterData(); // re-apply whatever the user has applied
                    return true;
                }
                catch (error) {
                    console.error("Error fetching data:", error);
                    const message = error instanceof Error ? error.message : String(error);
                    showTableMessage(`Failed to fetch data: ${message}`);
                    toast(`Failed to fetch data: ${message}`, "error");
                    return false;
                }
                finally {
                    hideProcurementLoading();
                    removeLoadingOverlay();
                }
            };
            const initializeFilters = () => {
                // ✅ Role access (same as SEIBIES): dropdowns derive from the projects this
                // role can see (GH Heads / PM Managers / READ_ALL get all projects).
                const assignedProjects = getAssignedProjects(projects);
                const filterSource = canReadAllProjects() ? projects : assignedProjects;
                const uniqueCustomers = [
                    ...new Set(filterSource.map((p) => p.CustomerName).filter(Boolean)),
                ].sort();
                filterCustomer.innerHTML =
                    '<option value="">All Customers</option>' +
                        uniqueCustomers
                            .map((customer) => `<option value="${customer}">${customer}</option>`)
                            .join("");
                // ✅ ProjectManager - Person/Group type
                const uniqueManagers = [
                    ...new Set(filterSource
                        .map((p) => getProjectManagerDisplayName(p.ProjectManager))
                        .filter(Boolean)),
                ].sort();
                filterProjectManager.innerHTML =
                    '<option value="">All Managers</option>' +
                        uniqueManagers
                            .map((manager) => `<option value="${manager}">${manager}</option>`)
                            .join("");
                // ✅ GroupHead1 - Person/Group type
                const uniqueGroupHeads = [
                    ...new Set(filterSource
                        .map((p) => getGroupHeadDisplayName(p.GroupHead1))
                        .filter(Boolean)),
                ].sort();
                filterGroupHead.innerHTML =
                    '<option value="">All Group Heads</option>' +
                        uniqueGroupHeads
                            .map((head) => `<option value="${head}">${head}</option>`)
                            .join("");
                const uniqueReferences = [
                    ...new Set(filterSource.map((p) => p.ProjectReference).filter(Boolean)),
                ].sort();
                filterProjectReference.innerHTML =
                    '<option value="">All References</option>' +
                        uniqueReferences
                            .map((ref) => `<option value="${ref}">${ref}</option>`)
                            .join("");
            };
            // ✅ The actual filtering - unchanged logic, now triggered by the popup's
            //    Apply button (or the floating Apply pill) instead of every keystroke.
            const filterData = () => {
                const project = filterProject.value;
                const customer = filterCustomer.value;
                const projectManager = filterProjectManager.value;
                const groupHead = filterGroupHead.value;
                const projectReference = filterProjectReference.value;
                const filtered = procurementData.filter((item) => {
                    var _a;
                    const projectData = resolveProjectForItem(item);
                    const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || (projectData === null || projectData === void 0 ? void 0 : projectData.Title) || "";
                    if (project && projectTitle !== project)
                        return false;
                    if (customer && (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) !== customer)
                        return false;
                    // ✅ ProjectManager - Person/Group type
                    if (projectManager) {
                        const managerName = getProjectManagerDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager);
                        if (managerName !== projectManager)
                            return false;
                    }
                    // ✅ GroupHead1 - Person/Group type
                    if (groupHead) {
                        const groupHeadName = getGroupHeadDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead1);
                        if (groupHeadName !== groupHead)
                            return false;
                    }
                    if (projectReference &&
                        (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) !== projectReference)
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
                currentPage = 1;
                renderTable(procurementData);
                toast("Filters cleared", "success");
            };
            // =====================================================================
            // ✅ FILTER ENGINE + FILTER POPUP  (same behaviour as Inventory / SEIBIES)
            //
            //    appliedFilters               = what the TABLE is filtered by
            //    the five <select> values     = the DRAFT inside the popup
            //      • Apply              → snapshot the draft into appliedFilters and
            //                             run the existing filterData()
            //      • Clear              → clearFilters() (this module's own logic)
            //      • Cancel             → a real cancel: the applied values go back
            //      • ✕ / Esc / backdrop → close but KEEP the draft, so the compact
            //                             "Apply" pill appears (mock-up image 8)
            //    A missing element never breaks the flow (every node is optional).
            // =====================================================================
            const filterTriggerBtn = document.querySelector("#procFilterTrigger");
            const filterOverlay = document.querySelector("#procFilterOverlay");
            const filterModalCloseBtn = document.querySelector("#procFilterClose");
            const filterCancelBtn = document.querySelector("#procFilterCancel");
            const filterClearModalBtn = document.querySelector("#procFilterClear");
            const filterApplyModalBtn = document.querySelector("#procFilterApply");
            const floatingApplyBtn = document.querySelector("#procFloatingApply");
            const filterCountBadge = document.querySelector("#procFilterCount");
            const filterSummaryBox = document.querySelector("#procFilterSummary");
            let appliedFilters = {};
            const FILTER_FIELDS = ['project', 'customer', 'manager', 'groupHead', 'reference'];
            const FILTER_SELECTS = {
                project: filterProject,
                customer: filterCustomer,
                manager: filterProjectManager,
                groupHead: filterGroupHead,
                reference: filterProjectReference
            };
            const FILTER_LABELS = {
                project: 'Project',
                customer: 'Customer',
                manager: 'Manager',
                groupHead: 'Group Head',
                reference: 'Reference'
            };
            const isFilterModalOpen = () => !!filterOverlay && filterOverlay.classList.contains('is-open');
            const readDraftFilters = () => {
                const out = {};
                FILTER_FIELDS.forEach(field => {
                    const el = FILTER_SELECTS[field];
                    if (el && el.value)
                        out[field] = el.value;
                });
                return out;
            };
            const sameFilters = (a, b) => FILTER_FIELDS.every(field => (a[field] || '') === (b[field] || ''));
            const countFilters = (source) => FILTER_FIELDS.filter(field => !!source[field]).length;
            const syncFilterFieldStyles = () => {
                FILTER_FIELDS.forEach(field => {
                    const el = FILTER_SELECTS[field];
                    if (el)
                        el.classList.toggle('is-set', !!el.value);
                });
            };
            // Apply/Clear enable only when there is something to do; the floating Apply
            // pill shows only while there are unapplied changes and the popup is closed
            const syncFilterButtons = () => {
                const draft = readDraftFilters();
                const dirty = !sameFilters(draft, appliedFilters);
                const anyDraft = countFilters(draft) > 0;
                const anyApplied = countFilters(appliedFilters) > 0;
                if (filterApplyModalBtn) {
                    filterApplyModalBtn.disabled = !dirty;
                    filterApplyModalBtn.classList.toggle('is-enabled', dirty);
                }
                if (filterClearModalBtn) {
                    const canClear = anyDraft || anyApplied;
                    filterClearModalBtn.disabled = !canClear;
                    filterClearModalBtn.classList.toggle('is-enabled', canClear);
                }
                if (floatingApplyBtn) {
                    floatingApplyBtn.classList.toggle('is-visible', dirty && !isFilterModalOpen());
                }
            };
            const syncFilterTrigger = () => {
                const n = countFilters(appliedFilters);
                if (filterCountBadge)
                    filterCountBadge.textContent = String(n);
                if (filterTriggerBtn)
                    filterTriggerBtn.classList.toggle('is-active', n > 0);
            };
            // chips above the table: one per applied filter, each removable
            const renderFilterChips = () => {
                if (!filterSummaryBox)
                    return;
                const activeFields = FILTER_FIELDS.filter(field => !!appliedFilters[field]);
                if (activeFields.length === 0) {
                    filterSummaryBox.style.display = 'none';
                    filterSummaryBox.innerHTML = '';
                    return;
                }
                filterSummaryBox.style.display = 'flex';
                filterSummaryBox.innerHTML =
                    '<span class="proc-filter-summary__label">Active filters:</span>' +
                        activeFields.map(field => `<span class="proc-filter-chip">${FILTER_LABELS[field]}: ${appliedFilters[field]}<button type="button" data-clear-filter="${field}" title="Remove this filter">&times;</button></span>`).join('') +
                        '<button type="button" class="proc-filter-chip__all" id="procFilterChipsClearAll">Clear all</button>';
                filterSummaryBox.querySelectorAll('[data-clear-filter]').forEach(btn => {
                    btn.addEventListener('click', () => {
                        const field = btn.dataset.clearFilter;
                        const el = FILTER_SELECTS[field];
                        if (el)
                            el.value = '';
                        delete appliedFilters[field];
                        syncFilterFieldStyles();
                        syncFilterTrigger();
                        renderFilterChips();
                        syncFilterButtons();
                        filterData();
                    });
                });
                const clearAllBtn = filterSummaryBox.querySelector('#procFilterChipsClearAll');
                if (clearAllBtn) {
                    clearAllBtn.addEventListener('click', () => {
                        clearFilters(); // this module's own logic (selects + table + toast)
                        syncFilterUIFromSelects();
                    });
                }
            };
            // after a Clear / Reset the selects are the source of truth again
            const syncFilterUIFromSelects = () => {
                appliedFilters = readDraftFilters();
                syncFilterFieldStyles();
                syncFilterTrigger();
                renderFilterChips();
                syncFilterButtons();
            };
            const openFilterModal = () => {
                if (!filterOverlay)
                    return;
                // the popup always opens showing what is currently applied
                FILTER_FIELDS.forEach(field => {
                    const el = FILTER_SELECTS[field];
                    if (el)
                        el.value = appliedFilters[field] || '';
                });
                syncFilterFieldStyles();
                syncFilterButtons();
                filterOverlay.classList.add('is-open');
                if (filterTriggerBtn) {
                    filterTriggerBtn.classList.add('is-open');
                    filterTriggerBtn.setAttribute('aria-expanded', 'true');
                }
                document.body.style.overflow = 'hidden';
            };
            const closeFilterModal = (restoreApplied = true) => {
                if (!filterOverlay)
                    return;
                if (restoreApplied) {
                    // real cancel: discard the draft, show the applied values again
                    FILTER_FIELDS.forEach(field => {
                        const el = FILTER_SELECTS[field];
                        if (el)
                            el.value = appliedFilters[field] || '';
                    });
                }
                filterOverlay.classList.remove('is-open');
                if (filterTriggerBtn) {
                    filterTriggerBtn.classList.remove('is-open');
                    filterTriggerBtn.setAttribute('aria-expanded', 'false');
                }
                document.body.style.overflow = '';
                syncFilterFieldStyles();
                syncFilterButtons();
            };
            const applyFilterModal = () => {
                appliedFilters = readDraftFilters();
                syncFilterTrigger();
                renderFilterChips();
                syncFilterButtons();
                filterData(); // real filtering + totals + pagination
                closeFilterModal(false); // keep the values that were just applied
                const n = countFilters(appliedFilters);
                toast(n > 0 ? `${n} procurement filter(s) applied` : "Procurement filters cleared", "success");
            };
            const clearFilterDraft = () => {
                clearFilters(); // this module's own clear logic, popup stays open
                syncFilterUIFromSelects();
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
            // ✅ Edit title like SEIBIES: "Edit Procurement Plan - <Project> (<Ref>)"
            // ✅ Title shows only the project name (Ref); "Edit Procurement Plan - " prefix removed
            const setEditTitle = (project) => {
                if (!formTitle)
                    return;
                const name = "Procurement Plan/BOM";
                // const ref = project?.ProjectReference || "";
                formTitle.textContent = name
                    ? `${name}`
                    : "Edit Procurement Plan/BOM";
                formSection.classList.add("edit-mode");
            };
            const openEditForm = (id) => {
                const record = procurementData.find((d) => d.Id === id);
                if (!record) {
                    toast("Record not found", "error");
                    return;
                }
                // ✅ Role access (same as SEIBIES): only permitted roles may open the edit form
                const guardProjectId = record[`${fieldMappings.Project}Id`] ||
                    (record[fieldMappings.Project] && record[fieldMappings.Project].Id);
                const guardProject = guardProjectId !== undefined && guardProjectId !== null
                    ? projects.find((p) => p.Id === guardProjectId)
                    : undefined;
                if (!canViewProcurementForm(true, guardProject)) {
                    toast("❌ You don't have permission to view this procurement plan", "error");
                    return;
                }
                editingId = id;
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formTitle.textContent = "";
                existingDataAlert.style.display = "block";
                resetTableSearch();
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
                    setEditTitle(project);
                }
                else {
                    setEditTitle(null);
                }
                applySaveButtonVisibility(true, guardProject);
                window.scrollTo({ top: 0, behavior: "auto" });
                procurementContainer.innerHTML = "";
                const procurementItems = parseProcurementForecast(record[fieldMappings.ProcurementForecast]) ||
                    [];
                procurementItems.forEach((item) => {
                    procurementContainer.appendChild(createProcurementRow(item));
                });
                const addMoreBtn = document.querySelector(".add-procurement-row");
                if (addMoreBtn)
                    addMoreBtn.style.display = "block";
                getSaveButtons().forEach(btn => {
                    btn.textContent = "💾 Save";
                });
                // Re-apply role-based visibility (overrides the inline default)
                applySaveButtonVisibility(true, guardProject);
                addImportButtonToEditMode(projectId, record);
            };
            // ============================================ //
            // OUTER-STICKY HEADER (main table)             //
            //   The page scrolls vertically. The main      //
            //   header is sticky at top:0. The table       //
            //   header is sticky at top = main header      //
            //   height, so it 'touches' the main header    //
            //   and pins below it. No per-cell JS offset.  //
            // ============================================ //
            const updateMainTableStickyOffset = () => {
                // const section = document.querySelector(".table-section") as HTMLElement | null;
                // const header = section?.querySelector(".page-header") as HTMLElement | null;
                // if (!section || !header) return;
                // Fall back to a sensible minimum in case offsetHeight is 0 (e.g.
                // before fonts/images load), so the table header never sticks at
                // top:0 on top of the page header.
                // const headerH = Math.max(header.offsetHeight || 0, 58);
                // section.style.setProperty("--sticky-header-top", `${headerH}px`);
            };
            // Re-measure after the page fully loads (fonts / layout settled) so the
            // offset matches the real page-header height and both stay pinned.
            if (document.readyState === "complete") {
                setTimeout(updateMainTableStickyOffset, 100);
            }
            else {
                window.addEventListener("load", () => setTimeout(updateMainTableStickyOffset, 100), { once: true });
            }
            // ============================================ //
            // PAGINATION INFO (mock-up bar)                //
            //   Fills the "Showing X - Y of Z items" text  //
            //   in the pagination bar.                     //
            // ============================================ //
            const updatePaginationInfo = (total) => {
                const startItemSpan = document.getElementById("procStartItem");
                const endItemSpan = document.getElementById("procEndItem");
                const totalItemsSpan = document.getElementById("procTotalItems");
                if (total === 0) {
                    if (startItemSpan)
                        startItemSpan.textContent = "0";
                    if (endItemSpan)
                        endItemSpan.textContent = "0";
                    if (totalItemsSpan)
                        totalItemsSpan.textContent = "0";
                    return;
                }
                const startItem = (currentPage - 1) * pageSize + 1;
                const endItem = Math.min(currentPage * pageSize, total);
                if (startItemSpan)
                    startItemSpan.textContent = startItem.toString();
                if (endItemSpan)
                    endItemSpan.textContent = endItem.toString();
                if (totalItemsSpan)
                    totalItemsSpan.textContent = total.toString();
            };
            // ============================================ //
            // MAIN TABLE RENDER                            //
            // - Only Project Reference is the edit link    //
            // - Row actions handled via event delegation   //
            //   (see tableBody click listener below)       //
            // ============================================ //
            // ✅ Find the Project List entry for a procurement item — by lookup Id
            // first (most reliable), then by lookup Title.
            const resolveProjectForItem = (item) => {
                var _a;
                const lookup = item === null || item === void 0 ? void 0 : item[fieldMappings.Project];
                const lookupId = (_a = (lookup && typeof lookup === "object" ? lookup.Id : undefined)) !== null && _a !== void 0 ? _a : item === null || item === void 0 ? void 0 : item[`${fieldMappings.Project}Id`];
                const lookupTitle = lookup && typeof lookup === "object" ? lookup.Title : undefined;
                let projectData;
                if (lookupId !== undefined && lookupId !== null) {
                    projectData = projects.find((p) => p.Id === Number(lookupId));
                }
                if (!projectData && lookupTitle) {
                    projectData = projects.find((p) => p.Title === lookupTitle);
                }
                return projectData;
            };
            const renderTable = (data) => {
                currentTableData = data;
                // ✅ Sort first (same as DCS): sortKey/sortDir + blanks-last
                const sortedData = applySort(data);
                const start = (currentPage - 1) * pageSize;
                const pageData = sortedData.slice(start, start + pageSize);
                tableBody.innerHTML = "";
                if (pageData.length > 0) {
                    pageData.forEach((item) => {
                        var _a;
                        const projectData = resolveProjectForItem(item);
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || (projectData === null || projectData === void 0 ? void 0 : projectData.Title) || "";
                        const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) || [];
                        const statusCounts = calculateProjectStatusCounts(procurementItems);
                        const totalLineItems = procurementItems.length;
                        const managerName = getProjectManagerDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager);
                        const groupHeadName = getGroupHeadDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead1);
                        // ✅ Only Project Reference is clickable. If a project has no reference,
                        // fall back to "#Id" so every row still has an edit link.
                        const projectRefLabel = (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || `#${item.Id}`;
                        const row = document.createElement("tr");
                        row.innerHTML = `
              <td class="sticky-main-col main-col-1">
                <span class="clickable-link" data-id="${item.Id}" title="Click to edit">${escapeHtml(projectRefLabel)}</span>
              </td>
              <td class="sticky-main-col main-col-2" title="${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "")}">${escapeHtml((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "")}</td>
              <td class="sticky-main-col main-col-3" title="${escapeHtml(projectTitle)}"><span class="project-name-text">${escapeHtml(projectTitle)}</span></td>
              <td class="sticky-main-col main-col-4" title="${escapeHtml(groupHeadName)}">${escapeHtml(groupHeadName)}</td>
              <td class="sticky-main-col main-col-5" title="${escapeHtml(managerName)}">${escapeHtml(managerName)}</td>
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
                <button type="button" class="view-btn proc-view-btn" data-id="${item.Id}" title="View">👁️ View</button>
              </td>
            `;
                        tableBody.appendChild(row);
                    });
                }
                else {
                    const noDataRow = document.createElement("tr");
                    noDataRow.innerHTML = `<td colspan="19" style="text-align:center;padding:20px;">No records found</td>`;
                    tableBody.appendChild(noDataRow);
                }
                renderPagination(Math.ceil(data.length / pageSize));
                updatePaginationInfo(data.length);
                updateSortIndicators();
                updateMainTableStickyOffset();
            };
            // ✅ Event delegation for the main table (edit link / view).
            // Attached once, so it keeps working after every re-render.
            tableBody.addEventListener("click", (e) => {
                const target = e.target;
                const viewBtn = target.closest(".view-btn");
                if (viewBtn) {
                    e.preventDefault();
                    e.stopPropagation();
                    const id = parseInt(viewBtn.dataset.id || "0", 10);
                    if (id > 0)
                        viewData(id);
                    return;
                }
                const link = target.closest(".clickable-link");
                if (link) {
                    e.preventDefault();
                    const id = parseInt(link.dataset.id || "0", 10);
                    if (id > 0)
                        openEditForm(id);
                }
            });
            // ============================================ //
            // PAGINATION (mock-up bar)                     //
            //   « ‹ [windowed tiles] of N › »              //
            //   Markup lives in procurementContent; this   //
            //   only refreshes state + tiles.              //
            // ============================================ //
            const calcTotalPages = (total) => Math.max(1, Math.ceil(total / pageSize));
            const renderPagination = (totalPages) => {
                const safeTotal = Math.max(1, totalPages);
                // ✅ every node is OPTIONAL, so a missing element never aborts the update.
                //    NOTE: the old "Page X of Y" pill is gone - the total now lives in the
                //    "of N" span at the end of the pager.
                const totalPagesSpan = document.getElementById("procTotalPages");
                if (totalPagesSpan)
                    totalPagesSpan.textContent = safeTotal.toString();
                const firstPageBtn = document.getElementById("procFirstPage");
                const prevPageBtn = document.getElementById("procPrevPage");
                const nextPageBtn = document.getElementById("procNextPage");
                const lastPageBtn = document.getElementById("procLastPage");
                const onFirst = currentPage <= 1;
                const onLast = currentPage >= safeTotal;
                if (firstPageBtn)
                    firstPageBtn.disabled = onFirst;
                if (prevPageBtn)
                    prevPageBtn.disabled = onFirst;
                if (nextPageBtn)
                    nextPageBtn.disabled = onLast;
                if (lastPageBtn)
                    lastPageBtn.disabled = onLast;
                const pageNumbersDiv = document.getElementById("procPageNumbers");
                if (!pageNumbersDiv)
                    return;
                pageNumbersDiv.innerHTML = "";
                const maxVisiblePages = 5;
                let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
                let endPage = Math.min(safeTotal, startPage + maxVisiblePages - 1);
                if (endPage - startPage + 1 < maxVisiblePages) {
                    startPage = Math.max(1, endPage - maxVisiblePages + 1);
                }
                if (startPage > 1) {
                    const dotsBtn = document.createElement("button");
                    dotsBtn.className = "proc-pg-number dots";
                    dotsBtn.textContent = "...";
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
                for (let i = startPage; i <= endPage; i++) {
                    const pageBtn = document.createElement("button");
                    pageBtn.className = `proc-pg-number ${i === currentPage ? "active" : ""}`;
                    pageBtn.textContent = i.toString();
                    pageBtn.addEventListener("click", () => {
                        currentPage = i;
                        renderTable(currentTableData);
                    });
                    pageNumbersDiv.appendChild(pageBtn);
                }
                if (endPage < safeTotal) {
                    const dotsBtn = document.createElement("button");
                    dotsBtn.className = "proc-pg-number dots";
                    dotsBtn.textContent = "...";
                    dotsBtn.disabled = true;
                    pageNumbersDiv.appendChild(dotsBtn);
                }
            };
            // Pagination nav buttons (First / Prev / Next / Last)
            const goToProcPage = (page) => {
                const total = currentTableData.length ? currentTableData : procurementData;
                const safeTotal = calcTotalPages(total.length);
                if (page < 1 || page > safeTotal || page === currentPage)
                    return;
                currentPage = page;
                renderTable(total);
            };
            // ============================================ //
            // ITEMS PER PAGE - keeps the FIRST VISIBLE ROW //
            //   (same fix as the Invoicing module)          //
            //   25 -> 10 while on page 3 must not jump:     //
            //   the row that was on top stays on top, so    //
            //   the page number is derived from it AFTER    //
            //   the new size is known.                      //
            // ============================================ //
            const changePageSize = (newSize) => {
                if (!newSize || newSize === pageSize)
                    return;
                const total = currentTableData.length ? currentTableData : procurementData;
                const firstVisibleIndex = (currentPage - 1) * pageSize;
                const totalPages = Math.max(1, Math.ceil(total.length / newSize));
                pageSize = newSize;
                currentPage = Math.max(1, Math.min(totalPages, Math.floor(firstVisibleIndex / pageSize) + 1));
                renderTable(total);
            };
            // ============================================ //
            // COLUMN SORTING - all sortable columns        //
            //   Every header that carries data-sort has    //
            //   ONE ⇅ circle: the top triangle sorts       //
            //   ascending, the bottom one descending       //
            //   (clicking the header itself still toggles) //
            //   Same behaviour as DCS: case-insensitive    //
            //   text sort, numeric for the count columns,  //
            //   blanks always last, resets to page 1.      //
            // ============================================ //
            let sortKey = null;
            let sortDir = 'asc';
            const getSortValue = (item, key) => {
                var _a;
                const projectData = resolveProjectForItem(item);
                switch (key) {
                    case "reference": return ((projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "").toLowerCase();
                    case "customer": return ((projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "").toLowerCase();
                    case "project": return (((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || (projectData === null || projectData === void 0 ? void 0 : projectData.Title) || "").toLowerCase();
                    case "grouphead": return getGroupHeadDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead1).toLowerCase();
                    case "manager": return getProjectManagerDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager).toLowerCase();
                    default: {
                        // ✅ Numeric count columns (Total Line Items + 12 statuses)
                        const items = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) || [];
                        if (key === "total")
                            return items.length;
                        const counts = calculateProjectStatusCounts(items);
                        switch (key) {
                            case "rfq": return counts["RFQ Pending"] || 0;
                            case "offer": return counts["Offer Pending"] || 0;
                            case "tc": return counts["TC pending"] || 0;
                            case "orderclosure": return counts["Order Closure pending"] || 0;
                            case "pr": return counts["PR pending"] || 0;
                            case "po": return counts["PO pending"] || 0;
                            case "docapproval": return counts["Document Approval Pending"] || 0;
                            case "mfc": return counts["MFC Pending"] || 0;
                            case "inspection": return counts["Inspection Pending"] || 0;
                            case "dispatch": return counts["Dispatch Pending"] || 0;
                            case "delivery": return counts["Delivery Pending"] || 0;
                            case "delivered": return counts["Delivered"] || 0;
                            default: return "";
                        }
                    }
                }
            };
            // Sort a copy of the data; empty values always go last.
            // Numeric-aware (same as DCS count columns): if both values parse as
            // numbers, compare numerically (2 < 10) instead of as text ("10" < "2").
            // A per-sort cache keeps getSortValue to one call per row (fast JSON).
            const applySort = (data) => {
                if (!sortKey)
                    return data;
                const dir = sortDir === "asc" ? 1 : -1;
                const valueCache = new Map();
                const valOf = (item) => {
                    if (!valueCache.has(item))
                        valueCache.set(item, getSortValue(item, sortKey));
                    return valueCache.get(item);
                };
                return [...data].sort((a, b) => {
                    const va = valOf(a);
                    const vb = valOf(b);
                    const aEmpty = va === "" || va === null || va === undefined;
                    const bEmpty = vb === "" || vb === null || vb === undefined;
                    if (aEmpty && bEmpty)
                        return 0;
                    if (aEmpty)
                        return 1;
                    if (bEmpty)
                        return -1;
                    const na = Number(va);
                    const nb = Number(vb);
                    if (!isNaN(na) && !isNaN(nb))
                        return (na - nb) * dir;
                    if (va < vb)
                        return -1 * dir;
                    if (va > vb)
                        return 1 * dir;
                    return 0;
                });
            };
            // One delegated listener on the (static) thead - survives tbody re-renders.
            // ✅ Clicking one of the two triangles sets that direction directly;
            //    clicking the header text keeps the old toggle behaviour.
            const handleSortClick = (e) => {
                const target = e.target;
                const th = target.closest('th[data-sort]');
                if (!th)
                    return;
                const key = th.getAttribute('data-sort') || '';
                if (!key)
                    return;
                const arrow = target.closest('.proc-sort-btn');
                if (arrow) {
                    const dir = (arrow.getAttribute('data-direction') || 'asc');
                    sortKey = key;
                    sortDir = dir;
                }
                else if (sortKey === key) {
                    sortDir = sortDir === 'asc' ? 'desc' : 'asc';
                }
                else {
                    sortKey = key;
                    sortDir = 'asc';
                }
                currentPage = 1; // same as DCS: sorting resets to page 1
                renderTable(currentTableData);
                updateSortIndicators();
            };
            const updateSortIndicators = () => {
                document.querySelectorAll('#procurementTable thead th[data-sort]').forEach((th) => {
                    const el = th;
                    const isActive = el.getAttribute('data-sort') === sortKey;
                    el.classList.remove('proc-sorted-asc', 'proc-sorted-desc');
                    if (isActive) {
                        el.classList.add(sortDir === 'asc' ? 'proc-sorted-asc' : 'proc-sorted-desc');
                    }
                    // ✅ the two triangles inside the circle light up individually
                    el.querySelectorAll('.proc-sort-btn').forEach((btn) => {
                        const dir = btn.getAttribute('data-direction');
                        btn.classList.toggle('proc-active', isActive && dir === sortDir);
                    });
                });
            };
            const mainThead = document.querySelector('#procurementTable thead.sticky-main-header');
            if (mainThead)
                mainThead.addEventListener('click', handleSortClick);
            const viewData = (id) => {
                var _a;
                // ✅ Role access (same as SEIBIES): view requires at least read access
                if (!canRead()) {
                    toast("❌ You don't have permission to view procurement data", "error");
                    return;
                }
                const record = procurementData.find((d) => d.Id === id);
                if (!record) {
                    toast("Record not found", "error");
                    return;
                }
                editingId = id;
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formTitle.textContent = "Procurement Plan/BOM";
                existingDataAlert.style.display = "none";
                formSection.classList.remove("edit-mode");
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
                getSaveButtons().forEach(btn => {
                    btn.style.display = "none";
                });
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
            // ============================================================ //
            // 🔍 SEARCH / FILTER - EDIT (& VIEW) PROCUREMENT TABLE          //
            //   ✅ CHANGED: the search box in the Description header now     //
            //      searches EVERY column of the line-item table, not only    //
            //      the Description column:                                   //
            //        • Description, FCM Qty, Engg Qty, Make, Model, Remarks  //
            //        • Status (auto-calculated text)                         //
            //        • ALL 33 Plan / Actual / Forecast dates - matched both  //
            //          as stored (YYYY-MM-DD) and as displayed (DD/MM/YYYY)  //
            //   Multiple words = AND search ("pump delivered" -> rows that    //
            //   contain both anywhere in the row). No UI / layout change.     //
            // ============================================================ //
            // Per-row searchable text cache (rebuilt whenever a row is edited)
            const rowSearchCache = new WeakMap();
            const buildRowSearchText = (row) => {
                const parts = [];
                // 1) Editable cells (inputs / selects) - edit mode
                const controls = row.querySelectorAll("input, select, textarea");
                controls.forEach((el) => {
                    const type = el.type || "";
                    if (type === "button" || type === "submit" || type === "hidden")
                        return;
                    const value = (el.value || "").trim();
                    if (!value)
                        return;
                    parts.push(value);
                    // Dates: also index the displayed DD/MM/YYYY form so a user can
                    // type "31/12/2025" as well as "2025-12-31".
                    if (type === "date") {
                        const display = formatDate(value);
                        if (display && display !== value)
                            parts.push(display);
                    }
                    // Selects: index the visible option text too
                    if (el instanceof HTMLSelectElement) {
                        const opt = el.options[el.selectedIndex];
                        if (opt && opt.text)
                            parts.push(opt.text);
                    }
                });
                // 2) Read-only cells (status chip, view-mode fields, plain text)
                const texts = row.querySelectorAll(".status-text-display, .view-field, .status-badge");
                texts.forEach((el) => {
                    const t = (el.textContent || "").trim();
                    if (t)
                        parts.push(t);
                });
                return parts.join(" | ").toLowerCase();
            };
            const getRowSearchText = (row) => {
                let text = rowSearchCache.get(row);
                if (text === undefined) {
                    text = buildRowSearchText(row);
                    rowSearchCache.set(row, text);
                }
                return text;
            };
            const invalidateRowSearchText = (row) => {
                if (row)
                    rowSearchCache.delete(row);
            };
            // Small "x of y line items" hint next to the search box
            const updateSearchCount = (shown, total) => {
                if (!tableSearchCount)
                    return;
                const query = ((tableSearchInput === null || tableSearchInput === void 0 ? void 0 : tableSearchInput.value) || "").trim();
                tableSearchCount.textContent = query
                    ? `${shown} of ${total} line item${total === 1 ? "" : "s"} match`
                    : "";
            };
            // ✅ Searches ALL columns of the Edit / View procurement table.
            const filterEditTableRows = () => {
                const query = ((tableSearchInput === null || tableSearchInput === void 0 ? void 0 : tableSearchInput.value) || "").trim().toLowerCase();
                const rows = Array.from(procurementContainer.querySelectorAll("tr"));
                if (tableSearchClearBtn) {
                    tableSearchClearBtn.style.display = query ? "flex" : "none";
                }
                if (!query) {
                    rows.forEach((row) => {
                        row.style.display = "";
                    });
                    updateSearchCount(rows.length, rows.length);
                    return;
                }
                // multi-word search: every term must be found somewhere in the row
                const terms = query.split(/\s+/).filter(Boolean);
                let shown = 0;
                rows.forEach((row) => {
                    const haystack = getRowSearchText(row);
                    const matches = terms.every((term) => haystack.includes(term));
                    row.style.display = matches ? "" : "none";
                    if (matches)
                        shown++;
                });
                updateSearchCount(shown, rows.length);
            };
            const resetTableSearch = () => {
                if (tableSearchInput) {
                    tableSearchInput.value = "";
                }
                if (tableSearchClearBtn) {
                    tableSearchClearBtn.style.display = "none";
                }
                if (tableSearchCount) {
                    tableSearchCount.textContent = "";
                }
                const rows = Array.from(procurementContainer.querySelectorAll("tr"));
                rows.forEach((row) => {
                    invalidateRowSearchText(row);
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
                // Red highlight for overdue Actual cells in view mode too
                const actualClass = (planField, actualField) => {
                    var _a, _b;
                    return isActualOverdue(String((_a = data === null || data === void 0 ? void 0 : data[planField]) !== null && _a !== void 0 ? _a : ""), String((_b = data === null || data === void 0 ? void 0 : data[actualField]) !== null && _b !== void 0 ? _b : ""))
                        ? " actual-overdue"
                        : "";
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
          <td><div class="view-field${actualClass("RfqPlan", "RfqActual")}">${getDisplayValue("RfqActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("RfqForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("OfferPlan")}</div></td>
          <td><div class="view-field${actualClass("OfferPlan", "OfferActual")}">${getDisplayValue("OfferActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("OfferForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("TechPlan")}</div></td>
          <td><div class="view-field${actualClass("TechPlan", "TechActual")}">${getDisplayValue("TechActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("TechForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("OrderPlan")}</div></td>
          <td><div class="view-field${actualClass("OrderPlan", "OrderActual")}">${getDisplayValue("OrderActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("OrderForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("PRPlan")}</div></td>
          <td><div class="view-field${actualClass("PRPlan", "PRActual")}">${getDisplayValue("PRActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("PRForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("POPlan")}</div></td>
          <td><div class="view-field${actualClass("POPlan", "POActual")}">${getDisplayValue("POActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("POForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("DocPlan")}</div></td>
          <td><div class="view-field${actualClass("DocPlan", "DocActual")}">${getDisplayValue("DocActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("DocForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("MfgPlan")}</div></td>
          <td><div class="view-field${actualClass("MfgPlan", "MfgActual")}">${getDisplayValue("MfgActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("MfgForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("InspPlan")}</div></td>
          <td><div class="view-field${actualClass("InspPlan", "InspActual")}">${getDisplayValue("InspActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("InspForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("DispPlan")}</div></td>
          <td><div class="view-field${actualClass("DispPlan", "DispActual")}">${getDisplayValue("DispActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("DispForecast")}</div></td>

          <td><div class="view-field">${getDisplayValue("DeliveryPlan")}</div></td>
          <td><div class="view-field${actualClass("DeliveryPlan", "DeliveryActual")}">${getDisplayValue("DeliveryActual")}</div></td>
          <td><div class="view-field">${getDisplayValue("DeliveryForecast")}</div></td>

          <td><div class="view-field remarks-view">${escapeHtml(getDisplayValue("Remarks"))}</div></td>
          <td></td>
        `;
                return row;
            };
            // ============================================ //
            // IMPORT FUNCTIONS (simplified)                //
            // ============================================ //
            // const collectValidationErrors = (data: any[], headers: string[]): ValidationError[] => {
            //   const errors: ValidationError[] = [];
            //   if (!data || data.length === 0) {
            //     errors.push({
            //       type: 'error',
            //       message: 'No data found in the Excel file.'
            //     });
            //     return errors;
            //   }
            //   const requiredColumns = ['description', 'status', 'make'];
            //   const missingColumns = requiredColumns.filter(col =>
            //     !headers.some(h => h.toLowerCase().includes(col) || h.toLowerCase() === col)
            //   );
            //   if (missingColumns.length > 0) {
            //     errors.push({
            //       type: 'error',
            //       message: `Missing required columns: ${missingColumns.join(', ')}`
            //     });
            //   }
            //   return errors;
            // };
            // const showValidationErrors = (errors: ValidationError[], container: HTMLElement): void => {
            //   const existingErrorList = container.querySelector('.error-list');
            //   if (existingErrorList) existingErrorList.remove();
            //   if (errors.length === 0) return;
            //   const errorList = document.createElement('div');
            //   errorList.className = 'error-list show';
            //   const criticalErrors = errors.filter(e => e.type === 'error');
            //   const warnings = errors.filter(e => e.type === 'warning');
            //   let html = `
            //     <div class="error-title">
            //       ⚠️ ${criticalErrors.length > 0 ? `${criticalErrors.length} Critical Error(s) Found` : 'Validation Issues Found'}
            //     </div>
            //     <ul class="error-items">
            //   `;
            //   criticalErrors.forEach(err => {
            //     html += `<li class="error-critical">${err.message}</li>`;
            //   });
            //   warnings.forEach(err => {
            //     html += `<li class="error-warning">${err.message}</li>`;
            //   });
            //   html += `</ul>`;
            //   if (criticalErrors.length > 0) {
            //     html += `
            //       <div style="margin-top: 10px; padding: 8px 12px; background: #dc3545; color: white; border-radius: 4px; font-size: 13px; font-weight: 600;">
            //         ❌ ${criticalErrors.length} critical error(s) must be fixed before importing
            //       </div>
            //     `;
            //   }
            //   errorList.innerHTML = html;
            //   const previewSection = container.querySelector('#previewSection');
            //   if (previewSection) {
            //     previewSection.insertAdjacentElement('afterend', errorList);
            //   } else {
            //     container.appendChild(errorList);
            //   }
            // };
            // const parseExcelFile = (file: File): Promise<any[]> => {
            //   return new Promise((resolve, reject) => {
            //     const reader = new FileReader();
            //     reader.onload = (e) => {
            //       try {
            //         if (!e.target?.result) {
            //           reject(new Error("Failed to read file"));
            //           return;
            //         }
            //         const data = new Uint8Array(e.target.result as ArrayBuffer);
            //         const workbook = XLSX.read(data, {
            //           type: "array",
            //           cellDates: true,
            //           raw: false,
            //         });
            //         const sheetName = workbook.SheetNames[0];
            //         const worksheet = workbook.Sheets[sheetName];
            //         const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            //           header: 1,
            //           defval: "",
            //         });
            //         if (jsonData.length < 2) {
            //           return resolve([]);
            //         }
            //         let headers: string[];
            //         let dataStartRow = 1;
            //         if (
            //           jsonData.length >= 2 &&
            //           Array.isArray(jsonData[0]) &&
            //           Array.isArray(jsonData[1])
            //         ) {
            //           const firstRow = jsonData[0] as string[];
            //           const secondRow = jsonData[1] as string[];
            //           if (
            //             firstRow.some(
            //               (cell) =>
            //                 cell && typeof cell === "string" && cell.includes("/")
            //             ) &&
            //             secondRow.some(
            //               (cell) =>
            //                 cell &&
            //                 typeof cell === "string" &&
            //                 (cell.toLowerCase().includes("desc") ||
            //                   cell.toLowerCase().includes("status"))
            //             )
            //           ) {
            //             headers = secondRow;
            //             dataStartRow = 2;
            //           } else {
            //             headers = firstRow;
            //           }
            //         } else {
            //           headers = jsonData[0] as string[];
            //         }
            //         const rows: any[] = [];
            //         const headerMap: { [key: string]: string } = {};
            //         headers.forEach((header) => {
            //           if (header && typeof header === "string") {
            //             const cleanHeader = header.toString().trim().toLowerCase();
            //             if (
            //               cleanHeader.includes("desc") ||
            //               cleanHeader.includes("item")
            //             ) {
            //               headerMap[cleanHeader] = "description";
            //             } else if (
            //               cleanHeader.includes("vendor") ||
            //               cleanHeader.includes("supplier")
            //             ) {
            //               headerMap[cleanHeader] = "vendor";
            //             } else if (
            //               cleanHeader.includes("fcm") &&
            //               (cleanHeader.includes("qty") ||
            //                 cleanHeader.includes("quantity"))
            //             ) {
            //               headerMap[cleanHeader] = "fcmquantity";
            //             } else if (
            //               cleanHeader.includes("engg") &&
            //               (cleanHeader.includes("qty") ||
            //                 cleanHeader.includes("quantity"))
            //             ) {
            //               headerMap[cleanHeader] = "enggquantity";
            //             } else if (cleanHeader.includes("status")) {
            //               headerMap[cleanHeader] = "status";
            //             } else if (
            //               cleanHeader.includes("make") ||
            //               cleanHeader.includes("manufacturer")
            //             ) {
            //               headerMap[cleanHeader] = "make";
            //             } else if (
            //               cleanHeader.includes("model") ||
            //               cleanHeader.includes("part")
            //             ) {
            //               headerMap[cleanHeader] = "model";
            //             } else if (cleanHeader.includes("remark")) {
            //               headerMap[cleanHeader] = "remarks";
            //             } else {
            //               headerMap[cleanHeader] = cleanHeader;
            //             }
            //           }
            //         });
            //         for (let i = dataStartRow; i < jsonData.length; i++) {
            //           const rowData = jsonData[i] as any[];
            //           if (
            //             !rowData ||
            //             rowData.every(
            //               (cell) =>
            //                 cell === "" || cell === null || cell === undefined
            //             )
            //           ) {
            //             continue;
            //           }
            //           const row: any = {};
            //           headers.forEach((header, idx) => {
            //             if (
            //               header &&
            //               rowData[idx] !== undefined &&
            //               rowData[idx] !== null &&
            //               rowData[idx] !== ""
            //             ) {
            //               const cleanHeader = header
            //                 .toString()
            //                 .trim()
            //                 .toLowerCase();
            //               const mappedHeader =
            //                 headerMap[cleanHeader] || cleanHeader;
            //               let value = rowData[idx];
            //               if (typeof value === "object" && value instanceof Date) {
            //                 value = value.toISOString().split("T")[0];
            //               } else if (typeof value === "number") {
            //                 value = value.toString();
            //               } else if (typeof value === "string") {
            //                 value = value.trim();
            //               }
            //               row[mappedHeader] = value;
            //             }
            //           });
            //           if (row.description || row.make || row.status || row.vendor) {
            //             rows.push(row);
            //           }
            //         }
            //         resolve(rows);
            //       } catch (err) {
            //         reject(err);
            //       }
            //     };
            //     reader.onerror = (error) => {
            //       console.error("FileReader error:", error);
            //       reject(new Error("Failed to read file"));
            //     };
            //     reader.onabort = () => {
            //       reject(new Error("File reading was aborted"));
            //     };
            //     try {
            //       reader.readAsArrayBuffer(file);
            //     } catch (error) {
            //       reject(new Error(`Cannot read file: ${error}`));
            //     }
            //   });
            // };
            // const showImportPreview = (data: any[], errors: ValidationError[], container: HTMLElement): void => {
            //   const errorRows = new Set(errors.filter(e => e.row).map(e => e.row));
            //   const previewHTML = `
            //     <div style="padding: 15px; background: #f8f9fa; border-radius: 8px;">
            //       <h4 style="color: #28a745; margin-bottom: 15px;">📋 Data Preview</h4>
            //       <p><strong>Found ${data.length} records:</strong></p>
            //       <div class="preview-table-container">
            //         <table class="preview-table">
            //           <thead>
            //             <tr>
            //               <th>#</th>
            //               <th>Description</th>
            //               <th>Status</th>
            //               <th>Make</th>
            //               <th>Model</th>
            //               <th>FCM Qty</th>
            //               <th>Engg Qty</th>
            //               <th>Remarks</th>
            //             </tr>
            //           </thead>
            //           <tbody>
            //             ${data
            //               .slice(0, 20)
            //               .map((row, index) => {
            //                 const rowNum = index + 1;
            //                 const hasError = errorRows.has(rowNum);
            //                 return `
            //                   <tr class="${hasError ? 'row-invalid' : 'row-valid'}">
            //                     <td style="text-align: center;">${rowNum}</td>
            //                     <td>${row.description || "N/A"}</td>
            //                     <td>${row.status || "N/A"}</td>
            //                     <td>${row.make || "N/A"}</td>
            //                     <td>${row.model || "N/A"}</td>
            //                     <td style="text-align: right;">${row.fcmquantity || "N/A"}</td>
            //                     <td style="text-align: right;">${row.enggquantity || "N/A"}</td>
            //                     <td>${row.remarks || "N/A"}</td>
            //                   </tr>
            //                 `;
            //               })
            //               .join("")}
            //           </tbody>
            //         </table>
            //       </div>
            //       ${data.length > 20 ? `<p style="margin-top: 10px; font-style: italic; color: #6c757d;">... and ${data.length - 20} more records</p>` : ""}
            //       <div style="margin-top: 15px; padding: 12px; background: #e7f3ff; border-radius: 5px;">
            //         <p style="margin: 0; color: #0056b3; font-size: 14px;">
            //           <strong>⚠️ Important:</strong> This will replace existing procurement data for this project.
            //         </p>
            //       </div>
            //     </div>
            //   `;
            //   const importPreview = container.querySelector('#importPreview') as HTMLElement;
            //   if (importPreview) {
            //     importPreview.innerHTML = previewHTML;
            //   }
            // };
            const addImportButtonToEditMode = (_projectId, _record) => {
                var _a;
                const container = document.getElementById("procurementContainerTable");
                if (!container)
                    return;
                let toolbar = document.getElementById("procFullToolbar");
                if (!toolbar) {
                    toolbar = document.createElement("div");
                    toolbar.id = "procFullToolbar";
                    toolbar.style.cssText = "display:flex;gap:10px;align-items:center;padding:8px 0 12px;flex-wrap:wrap;";
                    toolbar.innerHTML = `
      <button type="button" class="btn-primary import-excel-edit-btn" title="Upload the full line-item table from Excel">
        📥 Upload (all columns)
      </button>
      <button type="button" class="btn-secondary export-excel-edit-btn" title="Download current data as a fillable template">
        📥 Download Template
      </button>
      <input type="file" class="import-excel-edit-input" accept=".xlsx,.xls" style="display:none;" />
      <span class="proc-full-status" id="procFullStatus" style="font-size:12px;color:#6c757d;"></span>
    `;
                    (_a = container.parentNode) === null || _a === void 0 ? void 0 : _a.insertBefore(toolbar, container);
                }
                toolbar.style.display = "flex";
                setupProcFullToolbar();
            };
            let procFullToolbarWired = false;
            const setProcFullStatus = (msg, tone = 'info') => {
                const el = document.getElementById('procFullStatus');
                if (!el)
                    return;
                el.textContent = msg;
                el.style.color = tone === 'error' ? '#c62828' : tone === 'success' ? '#107c10' : '#6c757d';
            };
            const setupProcFullToolbar = () => {
                const uploadBtn = document.querySelector('.import-excel-edit-btn');
                const downloadBtn = document.querySelector('.export-excel-edit-btn');
                const fileInput = document.querySelector('.import-excel-edit-input');
                if (!uploadBtn && !downloadBtn)
                    return;
                if (procFullToolbarWired) {
                    if (fileInput)
                        fileInput.value = '';
                    return;
                }
                procFullToolbarWired = true;
                // ---- UPLOAD (OPENS OS FILE PICKER IMMEDIATELY — NO POPUP) ----
                uploadBtn === null || uploadBtn === void 0 ? void 0 : uploadBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    setProcFullStatus('');
                    fileInput === null || fileInput === void 0 ? void 0 : fileInput.click(); // ← direct OS file picker
                });
                // ---- DOWNLOAD TEMPLATE ----
                downloadBtn === null || downloadBtn === void 0 ? void 0 : downloadBtn.addEventListener('click', () => {
                    try {
                        const rows = collectCurrentProcRows();
                        const aoa = buildProcTemplateAoa(rows);
                        const ws = XLSX.utils.aoa_to_sheet(aoa);
                        ws['!cols'] = PROC_FULL_COLUMNS.map((c) => ({
                            wch: c.field === 'Description' ? 34 : c.field === 'Remarks' ? 28 : c.kind === 'date' ? 12 : 14
                        }));
                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, 'Procurement');
                        const projectId = Number((projectSelect === null || projectSelect === void 0 ? void 0 : projectSelect.value) || 0);
                        const proj = projects.find((p) => p.Id === projectId);
                        const ref = ((proj === null || proj === void 0 ? void 0 : proj.ProjectReference) || (proj === null || proj === void 0 ? void 0 : proj.Title) || 'procurement')
                            .replace(/[\\/:*?"<>|]/g, '_').slice(0, 40);
                        const stamp = new Date().toISOString().slice(0, 10);
                        XLSX.writeFile(wb, `Procurement_Template_${ref}_${stamp}.xlsx`);
                        setProcFullStatus(`✅ Template downloaded (${rows.length} row${rows.length === 1 ? '' : 's'})`, 'success');
                        toast('Template downloaded', 'success');
                    }
                    catch (err) {
                        console.error('Procurement template download failed:', err);
                        setProcFullStatus('Template download failed', 'error');
                        toast('Download failed', 'error');
                    }
                });
                // ---- FILE CHOSEN → validate + replace table ----
                fileInput === null || fileInput === void 0 ? void 0 : fileInput.addEventListener('change', async () => {
                    var _a;
                    const file = (_a = fileInput.files) === null || _a === void 0 ? void 0 : _a[0];
                    if (!file)
                        return;
                    try {
                        setProcFullStatus('Reading file…');
                        const result = await readProcFullExcelWithValidation(file);
                        if (result.missingColumns.length > 0) {
                            setProcFullStatus(`Missing required column(s): ${result.missingColumns.join(', ')}`, 'error');
                            if (result.errors.length)
                                promptDownloadErrors(result.errors);
                            return;
                        }
                        if (result.validRows.length === 0 && result.errors.length === 0) {
                            setProcFullStatus('No data rows found in the file.', 'error');
                            return;
                        }
                        const confirmed = await Swal.fire({
                            title: 'Replace the whole table?',
                            html: `<b>${result.validRows.length}</b> valid row${result.validRows.length === 1 ? '' : 's'} will replace the current line items.` +
                                (result.errors.length
                                    ? `<br/><br/><span style="color:#c62828;"><b>${result.errors.length}</b> row${result.errors.length === 1 ? '' : 's'} failed validation and will be skipped.</span>`
                                    : '') +
                                `<br/><br/>Nothing is written to SharePoint until you press <b>Save</b>.`,
                            icon: 'warning',
                            showCancelButton: true,
                            confirmButtonText: 'Yes, replace',
                            cancelButtonText: 'Cancel',
                            confirmButtonColor: '#0078d4',
                            cancelButtonColor: '#6c757d'
                        });
                        if (!confirmed.isConfirmed) {
                            setProcFullStatus('Upload cancelled - nothing changed.');
                            return;
                        }
                        procurementContainer.innerHTML = '';
                        result.validRows.forEach((r) => procurementContainer.appendChild(createProcurementRow(r)));
                        setProcFullStatus(`✅ ${result.validRows.length} row${result.validRows.length === 1 ? '' : 's'} loaded. Press Save to commit.` +
                            (result.errors.length ? ` (${result.errors.length} skipped — see error file.)` : ''), result.errors.length ? 'info' : 'success');
                        toast('Table uploaded - press Save to commit', 'success');
                        if (result.errors.length > 0) {
                            promptDownloadErrors(result.errors);
                        }
                    }
                    catch (err) {
                        console.error('Procurement upload failed:', err);
                        setProcFullStatus(`Upload failed: ${err instanceof Error ? err.message : String(err)}`, 'error');
                        toast('Upload failed', 'error');
                    }
                    finally {
                        if (fileInput)
                            fileInput.value = '';
                    }
                });
            };
            const procFullToolbar = document.getElementById("procFullToolbar");
            if (procFullToolbar)
                procFullToolbar.style.display = "none";
            const PROC_FULL_COLUMNS = [
                { header: 'Description', field: 'Description', kind: 'text' },
                { header: 'FCM Qty', field: 'FCMQuantity', kind: 'number' },
                { header: 'Engg Qty', field: 'EnggQuantity', kind: 'number' },
                { header: 'Status', field: 'Status', kind: 'status' },
                { header: 'Make', field: 'Make', kind: 'text' },
                { header: 'Model', field: 'Model', kind: 'text' },
                { header: 'RFQ Plan', field: 'RfqPlan', kind: 'date' },
                { header: 'RFQ Actual', field: 'RfqActual', kind: 'date' },
                { header: 'RFQ Forecast', field: 'RfqForecast', kind: 'date' },
                { header: 'Offer Plan', field: 'OfferPlan', kind: 'date' },
                { header: 'Offer Actual', field: 'OfferActual', kind: 'date' },
                { header: 'Offer Forecast', field: 'OfferForecast', kind: 'date' },
                { header: 'Tech Plan', field: 'TechPlan', kind: 'date' },
                { header: 'Tech Actual', field: 'TechActual', kind: 'date' },
                { header: 'Tech Forecast', field: 'TechForecast', kind: 'date' },
                { header: 'Order Plan', field: 'OrderPlan', kind: 'date' },
                { header: 'Order Actual', field: 'OrderActual', kind: 'date' },
                { header: 'Order Forecast', field: 'OrderForecast', kind: 'date' },
                { header: 'PR Plan', field: 'PRPlan', kind: 'date' },
                { header: 'PR Actual', field: 'PRActual', kind: 'date' },
                { header: 'PR Forecast', field: 'PRForecast', kind: 'date' },
                { header: 'PO Plan', field: 'POPlan', kind: 'date' },
                { header: 'PO Actual', field: 'POActual', kind: 'date' },
                { header: 'PO Forecast', field: 'POForecast', kind: 'date' },
                { header: 'Document Plan', field: 'DocPlan', kind: 'date' },
                { header: 'Document Actual', field: 'DocActual', kind: 'date' },
                { header: 'Document Forecast', field: 'DocForecast', kind: 'date' },
                { header: 'MFC Plan', field: 'MfgPlan', kind: 'date' },
                { header: 'MFC Actual', field: 'MfgActual', kind: 'date' },
                { header: 'MFC Forecast', field: 'MfgForecast', kind: 'date' },
                { header: 'Inspection Plan', field: 'InspPlan', kind: 'date' },
                { header: 'Inspection Actual', field: 'InspActual', kind: 'date' },
                { header: 'Inspection Forecast', field: 'InspForecast', kind: 'date' },
                { header: 'Dispatch Plan', field: 'DispPlan', kind: 'date' },
                { header: 'Dispatch Actual', field: 'DispActual', kind: 'date' },
                { header: 'Dispatch Forecast', field: 'DispForecast', kind: 'date' },
                { header: 'Delivery Plan', field: 'DeliveryPlan', kind: 'date' },
                { header: 'Delivery Actual', field: 'DeliveryActual', kind: 'date' },
                { header: 'Delivery Forecast', field: 'DeliveryForecast', kind: 'date' },
                { header: 'Remarks', field: 'Remarks', kind: 'text' },
            ];
            // =====================================================================
            // ✅ FULL-COLUMN TOOLBAR — Template download, Upload, Error report
            // ---------------------------------------------------------------------
            // • Download Template → exports the CURRENT line items (all 39 columns)
            //   so the user has a ready-to-edit file.
            // • Upload → validates each row against the template; replaces the
            //   entire table with rows that pass. If any row fails, an error
            //   .xlsx is generated and offered for download.
            // • Never touches SharePoint — Save does that (existing flow).
            // =====================================================================
            // ---- ERROR FILE ----
            const promptDownloadErrors = (errors) => {
                Swal.fire({
                    title: 'Upload had errors',
                    html: `<b>${errors.length}</b> row${errors.length === 1 ? '' : 's'} could not be imported.<br/>Download the error report to see what to fix.`,
                    icon: 'warning',
                    showCancelButton: true,
                    confirmButtonText: '⬇ Download error file',
                    cancelButtonText: 'Close',
                    confirmButtonColor: '#dc3545',
                    cancelButtonColor: '#6c757d'
                }).then((r) => {
                    if (r.isConfirmed)
                        downloadProcErrors(errors);
                });
            };
            const downloadProcErrors = (errors) => {
                try {
                    const header = [...PROC_FULL_COLUMNS.map((c) => c.header), '__Row Number', '__Errors'];
                    const aoa = [header];
                    errors.forEach((e) => {
                        const rowVals = PROC_FULL_COLUMNS.map((c) => {
                            var _a;
                            const raw = (_a = e.raw) === null || _a === void 0 ? void 0 : _a[c.field];
                            if (raw === null || raw === undefined)
                                return '';
                            if (c.kind === 'date')
                                return String(raw).trim() ? formatDateForExcel(String(raw)) : '';
                            if (c.kind === 'number')
                                return typeof raw === 'number' ? raw : Number(raw) || '';
                            return String(raw);
                        });
                        aoa.push([...rowVals, e.rowNumber, e.errors.join(' | ')]);
                    });
                    const ws = XLSX.utils.aoa_to_sheet(aoa);
                    ws['!cols'] = [
                        ...PROC_FULL_COLUMNS.map((c) => ({ wch: c.field === 'Description' ? 34 : 14 })),
                        { wch: 8 },
                        { wch: 50 }
                    ];
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, 'Errors');
                    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
                    XLSX.writeFile(wb, `Procurement_Upload_Errors_${stamp}.xlsx`);
                    toast(`Error report downloaded (${errors.length} row${errors.length === 1 ? '' : 's'})`, 'success');
                }
                catch (err) {
                    console.error('Error file download failed:', err);
                    toast('Could not download error report', 'error');
                }
            };
            // ---- COLLECT CURRENT ROWS ----
            const collectCurrentProcRows = () => {
                const rows = [];
                procurementContainer.querySelectorAll('tr').forEach((tr) => {
                    const row = tr;
                    const v = (name) => {
                        var _a;
                        const el = row.querySelector(`[name="${name}"]`);
                        return el ? String((_a = el.value) !== null && _a !== void 0 ? _a : '') : '';
                    };
                    const num = (name) => {
                        const s = v(name).trim();
                        if (!s)
                            return null;
                        const n = Number(s);
                        return Number.isFinite(n) ? n : null;
                    };
                    const item = {
                        Description: v('description').trim(),
                        FCMQuantity: num('fcm_qty'),
                        EnggQuantity: num('engg_qty'),
                        Status: v('status') || 'RFQ Pending',
                        Make: v('make').trim(),
                        Model: v('model').trim(),
                        Remarks: v('remarks').trim(),
                        RfqPlan: v('rfq_plan'), RfqActual: v('rfq_actual'), RfqForecast: v('rfq_forecast'),
                        OfferPlan: v('offer_plan'), OfferActual: v('offer_actual'), OfferForecast: v('offer_forecast'),
                        TechPlan: v('tech_plan'), TechActual: v('tech_actual'), TechForecast: v('tech_forecast'),
                        OrderPlan: v('order_plan'), OrderActual: v('order_actual'), OrderForecast: v('order_forecast'),
                        PRPlan: v('pr_plan'), PRActual: v('pr_actual'), PRForecast: v('pr_forecast'),
                        POPlan: v('po_plan'), POActual: v('po_actual'), POForecast: v('po_forecast'),
                        DocPlan: v('doc_plan'), DocActual: v('doc_actual'), DocForecast: v('doc_forecast'),
                        MfgPlan: v('mfg_plan'), MfgActual: v('mfg_actual'), MfgForecast: v('mfg_forecast'),
                        InspPlan: v('insp_plan'), InspActual: v('insp_actual'), InspForecast: v('insp_forecast'),
                        DispPlan: v('disp_plan'), DispActual: v('disp_actual'), DispForecast: v('disp_forecast'),
                        DeliveryPlan: v('delivery_plan'), DeliveryActual: v('delivery_actual'), DeliveryForecast: v('delivery_forecast')
                    };
                    const hasData = item.Description || item.FCMQuantity !== null || item.EnggQuantity !== null ||
                        item.Make || item.Model || item.Remarks ||
                        PROC_FULL_COLUMNS.some((c) => c.kind === 'date' && String(item[c.field] || '').trim() !== '');
                    if (hasData)
                        rows.push(item);
                });
                return rows;
            };
            // ---- BUILD TEMPLATE (header + current data) ----
            const buildProcTemplateAoa = (rows) => {
                const header = PROC_FULL_COLUMNS.map((c) => c.header);
                const body = rows.map((r) => PROC_FULL_COLUMNS.map((c) => {
                    const raw = r[c.field];
                    if (raw === null || raw === undefined)
                        return '';
                    if (c.kind === 'date') {
                        const s = String(raw).trim();
                        return s ? formatDateForExcel(s) : '';
                    }
                    if (c.kind === 'number')
                        return typeof raw === 'number' ? raw : Number(raw) || '';
                    return String(raw);
                }));
                return [header, ...body];
            };
            const readProcFullExcelWithValidation = async (file) => {
                var _a, _b, _c, _d, _e;
                const buf = await file.arrayBuffer();
                const wb = XLSX.read(buf, { type: 'array', cellDates: true });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
                if (aoa.length < 2) {
                    return { validRows: [], errors: [], missingColumns: [] };
                }
                // 1) Header check
                const headerRow = (aoa[0] || []).map((h) => String(h !== null && h !== void 0 ? h : '').trim().toLowerCase());
                const idxOf = new Map();
                headerRow.forEach((h, i) => { if (h)
                    idxOf.set(h, i); });
                const missingColumns = [];
                const colIdx = {};
                PROC_FULL_COLUMNS.forEach((c) => {
                    const i = idxOf.get(c.header.toLowerCase());
                    if (i === undefined)
                        missingColumns.push(c.header);
                    else
                        colIdx[c.header] = i;
                });
                if (missingColumns.length) {
                    return { validRows: [], errors: [], missingColumns };
                }
                // 2) Row-by-row validation
                const validRows = [];
                const errors = [];
                const todayISO = getTodayISO();
                for (let r = 1; r < aoa.length; r++) {
                    const raw = aoa[r] || [];
                    const pick = (h) => raw[colIdx[h]];
                    const rowErrors = [];
                    // Description is required
                    const description = String((_a = pick('Description')) !== null && _a !== void 0 ? _a : '').trim();
                    if (!description)
                        rowErrors.push('Description is required');
                    // Status must be a known value
                    const statusRaw = String((_b = pick('Status')) !== null && _b !== void 0 ? _b : '').trim() || 'RFQ Pending';
                    if (!statusOptions.includes(statusRaw)) {
                        rowErrors.push(`Status "${statusRaw}" is not one of: ${statusOptions.join(', ')}`);
                    }
                    // Quantities must be numeric (or blank)
                    const parseQty = (v, name) => {
                        if (v === null || v === undefined || v === '')
                            return null;
                        const n = Number(String(v).replace(/[,\s]/g, ''));
                        if (!Number.isFinite(n))
                            rowErrors.push(`${name} must be a number`);
                        return Number.isFinite(n) ? n : null;
                    };
                    // Dates must be valid (or blank)
                    const parseDate = (v, name) => {
                        if (v === null || v === undefined || v === '')
                            return '';
                        const nd = normDate(v);
                        if (!nd)
                            rowErrors.push(`${name} is not a valid date`);
                        return nd;
                    };
                    const item = {
                        Description: description,
                        FCMQuantity: parseQty(pick('FCM Qty'), 'FCM Qty'),
                        EnggQuantity: parseQty(pick('Engg Qty'), 'Engg Qty'),
                        Status: statusRaw,
                        Make: String((_c = pick('Make')) !== null && _c !== void 0 ? _c : '').trim(),
                        Model: String((_d = pick('Model')) !== null && _d !== void 0 ? _d : '').trim(),
                        Remarks: String((_e = pick('Remarks')) !== null && _e !== void 0 ? _e : '').trim(),
                        RfqPlan: parseDate(pick('RFQ Plan'), 'RFQ Plan'),
                        RfqActual: parseDate(pick('RFQ Actual'), 'RFQ Actual'),
                        RfqForecast: parseDate(pick('RFQ Forecast'), 'RFQ Forecast'),
                        OfferPlan: parseDate(pick('Offer Plan'), 'Offer Plan'),
                        OfferActual: parseDate(pick('Offer Actual'), 'Offer Actual'),
                        OfferForecast: parseDate(pick('Offer Forecast'), 'Offer Forecast'),
                        TechPlan: parseDate(pick('Tech Plan'), 'Tech Plan'),
                        TechActual: parseDate(pick('Tech Actual'), 'Tech Actual'),
                        TechForecast: parseDate(pick('Tech Forecast'), 'Tech Forecast'),
                        OrderPlan: parseDate(pick('Order Plan'), 'Order Plan'),
                        OrderActual: parseDate(pick('Order Actual'), 'Order Actual'),
                        OrderForecast: parseDate(pick('Order Forecast'), 'Order Forecast'),
                        PRPlan: parseDate(pick('PR Plan'), 'PR Plan'),
                        PRActual: parseDate(pick('PR Actual'), 'PR Actual'),
                        PRForecast: parseDate(pick('PR Forecast'), 'PR Forecast'),
                        POPlan: parseDate(pick('PO Plan'), 'PO Plan'),
                        POActual: parseDate(pick('PO Actual'), 'PO Actual'),
                        POForecast: parseDate(pick('PO Forecast'), 'PO Forecast'),
                        DocPlan: parseDate(pick('Document Plan'), 'Document Plan'),
                        DocActual: parseDate(pick('Document Actual'), 'Document Actual'),
                        DocForecast: parseDate(pick('Document Forecast'), 'Document Forecast'),
                        MfgPlan: parseDate(pick('MFC Plan'), 'MFC Plan'),
                        MfgActual: parseDate(pick('MFC Actual'), 'MFC Actual'),
                        MfgForecast: parseDate(pick('MFC Forecast'), 'MFC Forecast'),
                        InspPlan: parseDate(pick('Inspection Plan'), 'Inspection Plan'),
                        InspActual: parseDate(pick('Inspection Actual'), 'Inspection Actual'),
                        InspForecast: parseDate(pick('Inspection Forecast'), 'Inspection Forecast'),
                        DispPlan: parseDate(pick('Dispatch Plan'), 'Dispatch Plan'),
                        DispActual: parseDate(pick('Dispatch Actual'), 'Dispatch Actual'),
                        DispForecast: parseDate(pick('Dispatch Forecast'), 'Dispatch Forecast'),
                        DeliveryPlan: parseDate(pick('Delivery Plan'), 'Delivery Plan'),
                        DeliveryActual: parseDate(pick('Delivery Actual'), 'Delivery Actual'),
                        DeliveryForecast: parseDate(pick('Delivery Forecast'), 'Delivery Forecast')
                    };
                    // Skip fully-empty rows
                    const hasAnyData = item.Description || item.Make || item.Model || item.Remarks ||
                        item.FCMQuantity !== null || item.EnggQuantity !== null ||
                        PROC_FULL_COLUMNS.some((c) => c.kind === 'date' && String(item[c.field] || '').trim() !== '');
                    if (!hasAnyData)
                        continue;
                    // Actual dates in the future (soft validation)
                    STAGE_PREFIXES.forEach((prefix) => {
                        const actualField = `${prefix.charAt(0).toUpperCase() + prefix.slice(1)}Actual`;
                        const val = String(item[actualField] || '');
                        if (val && val > todayISO) {
                            rowErrors.push(`${actualField} cannot be in the future (${val})`);
                        }
                    });
                    if (rowErrors.length) {
                        errors.push({ rowNumber: r + 1, raw: item, errors: rowErrors });
                    }
                    else {
                        validRows.push(item);
                    }
                }
                return { validRows, errors, missingColumns };
            };
            // ---- date / number helpers (kept) ----
            const normDate = (v) => {
                if (v === null || v === undefined || v === '')
                    return '';
                if (v instanceof Date && !isNaN(v.getTime())) {
                    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
                }
                const s = String(v).trim();
                const m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
                if (m) {
                    const dd = m[1].padStart(2, '0');
                    const mm = m[2].padStart(2, '0');
                    const yy = m[3].length === 2 ? `20${m[3]}` : m[3];
                    return `${yy}-${mm}-${dd}`;
                }
                if (/^\d{4}-\d{2}-\d{2}/.test(s))
                    return s.slice(0, 10);
                const d = new Date(s);
                if (!isNaN(d.getTime())) {
                    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                }
                return '';
            };
            // ============================================ //
            // IMPORT MODAL (EDIT MODE) — FIXED             //
            //   1. Removes any leftover modal before open  //
            //      (previously every click appended ANOTHER//
            //      modal and listeners bound to the first).//
            //   2. Inline styles: centered, fixed, on top  //
            //      even without .import-* CSS.             //
            //   3. ✕ / Cancel / backdrop all close THIS    //
            //      modal (scoped via `modal` variable).    //
            // ============================================ //
            // ============================================ //
            // EDIT / VIEW ROW BUILDERS                     //
            // ============================================ //
            const STAGE_FIELD_PREFIX = {
                rfq: "Rfq",
                offer: "Offer",
                tech: "Tech",
                order: "Order",
                pr: "PR",
                po: "PO",
                doc: "Doc",
                mfg: "Mfg",
                insp: "Insp",
                disp: "Disp",
                delivery: "Delivery",
            };
            const createProcurementRow = (data) => {
                const row = document.createElement("tr");
                row.className = "procurement-row";
                const text = (v) => v === null || v === undefined ? "" : escapeHtml(String(v));
                const num = (v) => v === null || v === undefined || v === "" ? "" : String(v);
                // ---- 11 stage groups × Plan / Actual / Forecast = 33 date cells ----
                const stageCells = STAGE_PREFIXES.map((prefix) => {
                    const fieldPrefix = STAGE_FIELD_PREFIX[prefix];
                    const plan = formatDateForInput(String((data === null || data === void 0 ? void 0 : data[`${fieldPrefix}Plan`]) || ""));
                    const actual = formatDateForInput(String((data === null || data === void 0 ? void 0 : data[`${fieldPrefix}Actual`]) || ""));
                    const forecast = formatDateForInput(String((data === null || data === void 0 ? void 0 : data[`${fieldPrefix}Forecast`]) || ""));
                    // ✅ Forecasts must be today or later — enforced by `min` + JS validation
                    const todayISO = getTodayISO();
                    return `
            <td><input type="date" name="${prefix}_plan" class="date-input plan-input" value="${plan}"></td>
            <td><input type="date" name="${prefix}_actual" class="date-input actual-input" value="${actual}"></td>
            <td><input type="date" name="${prefix}_forecast" class="date-input forecast-input"
                      min="${todayISO}" value="${forecast}"></td>
          `;
                }).join("");
                row.innerHTML = `
          <td class="sticky-col col-1">
            <textarea name="description" class="description-input" rows="2" placeholder="Enter description">${text(data === null || data === void 0 ? void 0 : data.Description)}</textarea>
          </td>
          <td class="sticky-col col-2">
            <input type="number" name="fcm_qty" class="qty-input" min="0" step="1" placeholder="0" value="${num(data === null || data === void 0 ? void 0 : data.FCMQuantity)}">
          </td>
          <td class="sticky-col col-3">
            <div class="engg-qty-wrapper">
              <input type="number" name="engg_qty" class="qty-input" min="0" step="1" placeholder="0" value="${num(data === null || data === void 0 ? void 0 : data.EnggQuantity)}">
            </div>
          </td>
          <td class="sticky-col col-4">
            <span class="status-text-display">${text((data === null || data === void 0 ? void 0 : data.Status) || getStatusFromActualDates(data))}</span>
          </td>
          <td class="sticky-col col-5">
            <input type="text" name="make" class="text-input" placeholder="Make" value="${text(data === null || data === void 0 ? void 0 : data.Make)}">
          </td>
          <td class="sticky-col col-6">
            <input type="text" name="model" class="text-input" placeholder="Model" value="${text(data === null || data === void 0 ? void 0 : data.Model)}">
          </td>
          ${stageCells}
          <td>
            <textarea name="remarks" class="remarks-input" rows="2" placeholder="Remarks">${text(data === null || data === void 0 ? void 0 : data.Remarks)}</textarea>
          </td>
          <td class="actions-cell">
            <button type="button" class="remove-row" title="Delete this line item">🗑️ Delete</button>
          </td>
        `;
                // ✅ Status is calculated from the Actual dates, Engg Qty locks after PO
                //    Actual, and a crossed Plan date highlights the empty Actual field.
                wireStatusAutoLink(row);
                // ✅ Forecast dates must be today or later (revert typed past dates)
                wireForecastMinValidation(row);
                // Invalidate the search cache whenever a field changes
                row.addEventListener("input", () => invalidateRowSearchText(row));
                row.addEventListener("change", () => invalidateRowSearchText(row));
                return row;
            };
            // "+ Add Line Item" - appends an empty row and keeps the search in sync
            procurementContainer.addEventListener("click", (e) => {
                const target = e.target;
                const deleteBtn = target.closest(".remove-row");
                if (deleteBtn) {
                    e.preventDefault();
                    const row = deleteBtn.closest("tr");
                    if (!row)
                        return;
                    const rows = procurementContainer.querySelectorAll("tr");
                    if (rows.length <= 1) {
                        toast("At least one line item is required", "warning");
                        return;
                    }
                    row.remove();
                    filterEditTableRows();
                    toast("Line item removed", "success");
                }
            });
            const addProcurementRow = () => {
                procurementContainer.appendChild(createProcurementRow());
                filterEditTableRows();
            };
            // Shows / hides the line-item table (the search bar + table + footer)
            const toggleProcurementSections = (show) => {
                if (procurementContainerTable) {
                    procurementContainerTable.style.display = show ? "block" : "none";
                }
                const addMoreBtn = document.querySelector(".add-procurement-row");
                if (addMoreBtn) {
                    addMoreBtn.style.display = show ? "inline-flex" : "none";
                    addMoreBtn.style.visibility = "visible";
                }
            };
            // ============================================ //
            // PROJECT CHANGE                               //
            // ============================================ //
            const onProjectChange = async (projectId, isEditMode = false) => {
                const project = projects.find((p) => p.Id === projectId);
                if (!project) {
                    renderProjectDetails(null);
                    toggleProcurementSections(false);
                    return;
                }
                renderProjectDetails(project);
                if (!isEditMode) {
                    // ✅ Role access (project-scoped):
                    //    FULL_CONTROL / CONTRIBUTE  -> allowed on any project
                    //    CONTRIBUTE_OWN             -> allowed only when the current user
                    //                                  is the person in the project's own role
                    //                                  field (EnggLead / EnggSubLead1 / ...)
                    //    READ / READ_ALL / NONE     -> not allowed to create
                    const allowedToCreateForThisProject = canCreateRecord() || canCreateForProject(project);
                    if (!allowedToCreateForThisProject) {
                        toast("❌ You don't have permission to add procurement for this project", "error");
                        toggleProcurementSections(false);
                        return;
                    }
                    existingDataAlert.style.display = "none";
                    applySaveButtonVisibility(false, project);
                }
                // A project only has ONE procurement record: load it (if it exists) so
                // editing always works on the stored data instead of a blank table.
                const existing = procurementData.find((item) => {
                    const projectData = resolveProjectForItem(item);
                    return (projectData === null || projectData === void 0 ? void 0 : projectData.Id) === projectId;
                });
                procurementContainer.innerHTML = "";
                if (existing) {
                    const items = parseProcurementForecast(existing[fieldMappings.ProcurementForecast]) ||
                        [];
                    items.forEach((item) => {
                        procurementContainer.appendChild(createProcurementRow(item));
                    });
                    existingDataAlert.style.display = "block";
                    editingId = existing.Id;
                }
                else if (!isEditMode) {
                    procurementContainer.appendChild(createProcurementRow());
                }
                // ✅ Keep Save visibility in sync with the newly selected project
                //    (edit mode uses the project too, so always recompute)
                applySaveButtonVisibility(isEditMode, project);
                toggleProcurementSections(true);
                resetTableSearch();
            };
            // ============================================ //
            // SHOW / HIDE THE FORM                         //
            // ============================================ //
            const showForm = () => {
                if (!canViewProcurementForm(false)) {
                    toast("❌ You don't have permission to add procurement", "error");
                    return;
                }
                editingId = null;
                formSection.style.display = "block";
                tableSection.style.display = "none";
                formTitle.textContent = "New";
                existingDataAlert.style.display = "none";
                formSection.classList.remove("edit-mode", "view-mode");
                projectSelect.disabled = false;
                projectSelect.value = "";
                renderProjectDetails(null);
                procurementContainer.innerHTML = "";
                toggleProcurementSections(false);
                resetTableSearch();
                // Remove anything that only belongs to view / edit mode
                document
                    .querySelectorAll(".import-excel-edit-btn, .export-project-excel-btn")
                    .forEach((btn) => btn.remove());
                getSaveButtons().forEach((btn) => {
                    btn.textContent = "💾 Save Procurement Plan";
                });
                applySaveButtonVisibility(false, null);
                window.scrollTo({ top: 0, behavior: "auto" });
            };
            const hideForm = () => {
                editingId = null;
                formSection.style.display = "none";
                tableSection.style.display = "block";
                formSection.classList.remove("edit-mode", "view-mode");
                projectSelect.disabled = false;
                projectSelect.value = "";
                procurementContainer.innerHTML = "";
                toggleProcurementSections(false);
                resetTableSearch();
                document
                    .querySelectorAll(".import-excel-edit-btn, .export-project-excel-btn")
                    .forEach((btn) => btn.remove());
                getSaveButtons().forEach((btn) => {
                    btn.textContent = "💾 Save Procurement Plan";
                });
                window.scrollTo({ top: 0, behavior: "auto" });
            };
            // ============================================ //
            // COLLECT THE FORM DATA                        //
            // ============================================ //
            const collectProcurementItems = () => {
                const rows = Array.from(procurementContainer.querySelectorAll("tr"));
                return rows.map((row) => {
                    const val = (name) => {
                        var _a;
                        return ((_a = row.querySelector(`[name="${name}"]`)) === null || _a === void 0 ? void 0 : _a.value.trim()) || "";
                    };
                    const dateFields = {};
                    STAGE_PREFIXES.forEach((prefix) => {
                        const fieldPrefix = STAGE_FIELD_PREFIX[prefix];
                        dateFields[`${fieldPrefix}Plan`] = val(`${prefix}_plan`);
                        dateFields[`${fieldPrefix}Actual`] = val(`${prefix}_actual`);
                        dateFields[`${fieldPrefix}Forecast`] = val(`${prefix}_forecast`);
                    });
                    const fcm = val("fcm_qty");
                    const engg = val("engg_qty");
                    const item = Object.assign({ Description: val("description"), FCMQuantity: fcm === "" ? null : Number(fcm), EnggQuantity: engg === "" ? null : Number(engg), Status: "", Make: val("make"), Model: val("model"), Remarks: val("remarks") }, dateFields);
                    // ✅ Status is always derived from the Actual dates (same rule the row uses)
                    item.Status = getStatusFromActualDates(item);
                    return item;
                });
            };
            const isForecastInPast = (value) => {
                if (!isFilled(value))
                    return false;
                const d = new Date(value);
                if (isNaN(d.getTime()))
                    return false;
                const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
                return iso < getTodayISO();
            };
            const validateProcurementItems = (items) => {
                if (items.length === 0)
                    return "Please add at least one line item.";
                const emptyRow = items.findIndex((item) => !item.Description);
                if (emptyRow !== -1) {
                    return `Row ${emptyRow + 1}: Description is required.`;
                }
                const badQty = items.findIndex((item) => (item.FCMQuantity !== null && Number.isNaN(item.FCMQuantity)) ||
                    (item.EnggQuantity !== null && Number.isNaN(item.EnggQuantity)));
                if (badQty !== -1) {
                    return `Row ${badQty + 1}: FCM Qty / Engg Qty must be a number.`;
                }
                // ✅ Every Forecast date must be today or later
                const forecastKeys = [
                    "RfqForecast",
                    "OfferForecast",
                    "TechForecast",
                    "OrderForecast",
                    "PRForecast",
                    "POForecast",
                    "DocForecast",
                    "MfgForecast",
                    "InspForecast",
                    "DispForecast",
                    "DeliveryForecast",
                ];
                for (let i = 0; i < items.length; i++) {
                    const item = items[i];
                    for (const key of forecastKeys) {
                        const value = String(item[key] || "");
                        if (isForecastInPast(value)) {
                            return `Row ${i + 1}: "${String(key).replace(/Forecast$/, "")} Forecast" (${value}) cannot be in the past. Please use today or a future date.`;
                        }
                    }
                }
                return null;
            };
            // ============================================ //
            // SAVE (create or update one procurement plan) //
            // ============================================ //
            // ============================================ //
            // SAVE (create or update one procurement plan) //
            //   • Shows the confirmation popup FIRST      //
            //   • Then performs the fetch                  //
            //   • Guard flag lives inside ONE try/finally  //
            //     so it can NEVER get stuck at true        //
            // ============================================ //
            // ============================================ //
            // EXPORT - SUMMARY TABLE (toolbar Download Excel)
            // ============================================ //
            const SUMMARY_HEADERS = [
                "Project Reference",
                "Customer",
                "Project",
                "Group Head",
                "Project Manager",
                "Total Line Items",
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
            const summaryRowFor = (item) => {
                var _a;
                const projectData = resolveProjectForItem(item);
                const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || (projectData === null || projectData === void 0 ? void 0 : projectData.Title) || "";
                const items = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) || [];
                const counts = calculateProjectStatusCounts(items);
                return [
                    (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "",
                    (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "",
                    projectTitle,
                    getGroupHeadDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead1),
                    getProjectManagerDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager),
                    items.length,
                    counts["RFQ Pending"] || 0,
                    counts["Offer Pending"] || 0,
                    counts["TC pending"] || 0,
                    counts["Order Closure pending"] || 0,
                    counts["PR pending"] || 0,
                    counts["PO pending"] || 0,
                    counts["Document Approval Pending"] || 0,
                    counts["MFC Pending"] || 0,
                    counts["Inspection Pending"] || 0,
                    counts["Dispatch Pending"] || 0,
                    counts["Delivery Pending"] || 0,
                    counts["Delivered"] || 0,
                ];
            };
            const exportToExcel = () => {
                // ✅ exports exactly what the table shows right now (filters included)
                const source = currentTableData.length ? currentTableData : procurementData;
                if (source.length === 0) {
                    toast("There is nothing to export", "warning");
                    return;
                }
                showExportProgress("Building the summary sheet...");
                try {
                    const aoa = [
                        SUMMARY_HEADERS,
                        ...source.map(summaryRowFor),
                    ];
                    const ws = XLSX.utils.aoa_to_sheet(aoa);
                    // header row styling (same look as the other modules)
                    SUMMARY_HEADERS.forEach((_, col) => {
                        const ref = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (ws[ref]) {
                            ws[ref].s = {
                                font: { bold: true, color: { rgb: "FFFFFFFF" } },
                                fill: { fgColor: { rgb: "FF356BFF" } },
                                alignment: { horizontal: "center", vertical: "center", wrapText: true },
                                border: {
                                    top: { style: "thin", color: { rgb: "FFDDE6FF" } },
                                    bottom: { style: "thin", color: { rgb: "FFDDE6FF" } },
                                    left: { style: "thin", color: { rgb: "FFDDE6FF" } },
                                    right: { style: "thin", color: { rgb: "FFDDE6FF" } },
                                },
                            };
                        }
                    });
                    ws["!cols"] = SUMMARY_HEADERS.map((h) => ({
                        wch: Math.max(14, Math.min(30, h.length + 6)),
                    }));
                    // Freeze the header row
                    ws["!freeze"] = { x: 0, y: 1 };
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Procurement Summary");
                    const filtersActive = countFilters(appliedFilters) > 0;
                    const stamp = new Date().toISOString().split("T")[0];
                    const fileName = `Procurement_Summary${filtersActive ? "_filtered" : ""}_${stamp}.xlsx`;
                    XLSX.writeFile(wb, fileName);
                    toast(`Exported ${source.length} record${source.length === 1 ? "" : "s"} to Excel`, "success");
                }
                catch (error) {
                    console.error("Error exporting summary:", error);
                    toast("Failed to export the summary sheet", "error");
                }
                finally {
                    hideExportProgress();
                }
            };
            // ============================================ //
            // EXPORT - ONE PROJECT'S BOM (view mode)       //
            // ============================================ //
            const STAGE_LABELS = [
                "RFQ / Engg",
                "Vendor Offer / SCM",
                "Technical Clearance / ENGG",
                "Order Finalization / SCM",
                "PR / PM",
                "Purchase Order / SCM",
                "Document Approval / ENGG",
                "Manufacturing Clearance / ENGG",
                "Inspection",
                "Material Dispatch",
                "Delivery at Site",
            ];
            const exportProjectToExcel = (projectId, record) => {
                const project = projects.find((p) => p.Id === projectId);
                const items = parseProcurementForecast(record[fieldMappings.ProcurementForecast]) || [];
                if (items.length === 0) {
                    toast("This project has no procurement line items to export", "warning");
                    return;
                }
                showExportProgress("Building the BOM sheet...");
                try {
                    const headerRow = [
                        "Description",
                        "FCM Qty",
                        "Engg Qty",
                        "Status",
                        "Make",
                        "Model",
                    ];
                    STAGE_LABELS.forEach((label) => {
                        headerRow.push(`${label} - Plan`, `${label} - Actual`, `${label} - Forecast`);
                    });
                    headerRow.push("Remarks");
                    const aoa = [headerRow];
                    items.forEach((item) => {
                        var _a, _b;
                        const row = [
                            item.Description || "",
                            (_a = item.FCMQuantity) !== null && _a !== void 0 ? _a : "",
                            (_b = item.EnggQuantity) !== null && _b !== void 0 ? _b : "",
                            item.Status || getStatusFromActualDates(item),
                            item.Make || "",
                            item.Model || "",
                        ];
                        STAGE_PREFIXES.forEach((prefix) => {
                            const fieldPrefix = STAGE_FIELD_PREFIX[prefix];
                            row.push(formatDateForExcel(String((item === null || item === void 0 ? void 0 : item[`${fieldPrefix}Plan`]) || "")), formatDateForExcel(String((item === null || item === void 0 ? void 0 : item[`${fieldPrefix}Actual`]) || "")), formatDateForExcel(String((item === null || item === void 0 ? void 0 : item[`${fieldPrefix}Forecast`]) || "")));
                        });
                        row.push(item.Remarks || "");
                        aoa.push(row);
                    });
                    const ws = XLSX.utils.aoa_to_sheet(aoa);
                    headerRow.forEach((_, col) => {
                        const ref = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (ws[ref]) {
                            ws[ref].s = {
                                font: { bold: true, color: { rgb: "FFFFFFFF" } },
                                fill: { fgColor: { rgb: "FF605EEE" } },
                                alignment: { horizontal: "center", vertical: "center", wrapText: true },
                            };
                        }
                    });
                    ws["!cols"] = headerRow.map((h) => ({
                        wch: Math.max(12, Math.min(28, h.length + 6)),
                    }));
                    // Freeze the header row
                    ws["!freeze"] = { x: 0, y: 1 };
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Procurement BOM");
                    const safeName = ((project === null || project === void 0 ? void 0 : project.ProjectReference) || (project === null || project === void 0 ? void 0 : project.Title) || "Project")
                        .toString()
                        .replace(/[^\w\-]+/g, "_");
                    const fileName = `Procurement_BOM_${safeName}.xlsx`;
                    XLSX.writeFile(wb, fileName);
                    toast("Procurement BOM exported successfully", "success");
                }
                catch (error) {
                    console.error("Error exporting BOM:", error);
                    toast("Failed to export the BOM sheet", "error");
                }
                finally {
                    hideExportProgress();
                }
            };
            // ========== PAGINATION EVENT LISTENERS ==========
            //   Bound ONCE here (the tiles of the pager get their own listener each
            //   time renderPagination() repaints them).
            const setupPaginationEventListeners = () => {
                const firstPageBtn = document.getElementById("procFirstPage");
                const prevPageBtn = document.getElementById("procPrevPage");
                const nextPageBtn = document.getElementById("procNextPage");
                const lastPageBtn = document.getElementById("procLastPage");
                if (firstPageBtn)
                    firstPageBtn.addEventListener("click", () => goToProcPage(1));
                if (prevPageBtn)
                    prevPageBtn.addEventListener("click", () => goToProcPage(currentPage - 1));
                if (nextPageBtn)
                    nextPageBtn.addEventListener("click", () => goToProcPage(currentPage + 1));
                if (lastPageBtn) {
                    lastPageBtn.addEventListener("click", () => {
                        const total = currentTableData.length ? currentTableData : procurementData;
                        goToProcPage(calcTotalPages(total.length));
                    });
                }
                // ✅ Items per page - keeps the first visible row (changePageSize)
                const sizeSelect = document.getElementById("pageSizeSelect");
                if (sizeSelect) {
                    sizeSelect.addEventListener("change", (e) => {
                        changePageSize(parseInt(e.target.value, 10));
                    });
                }
            };
            // ✅ One reload path for every "the data may have changed" case
            //    (Back button, after a save): projects + procurement, then re-fill
            //    the dropdowns and re-run the filter the user still has applied.
            const reloadProcurementData = async () => {
                const loading = createLoadingOverlay("Reloading procurement data...");
                loaderHost().appendChild(loading);
                try {
                    await loadProjects();
                    await fetchData(); // ends with initializeFilters + syncFilterUIFromSelects + filterData
                }
                catch (error) {
                    console.error("❌ Reload failed:", error);
                    toast("Reload failed - showing current data", "error");
                }
                finally {
                    removeLoadingOverlay();
                }
            };
            window.reloadProcurementData = reloadProcurementData;
            // ========== EVENT LISTENERS ==========
            showFormBtn === null || showFormBtn === void 0 ? void 0 : showFormBtn.addEventListener("click", showForm);
            exportExcelBtn === null || exportExcelBtn === void 0 ? void 0 : exportExcelBtn.addEventListener("click", exportToExcel);
            // ============================================================
            // ✅ ONE delegated click listener for Save + Back.
            //    Because the buttons live inside `.form-section` (which is
            //    created from the template), and because the template may
            //    not be mounted yet when this script runs, we bind on
            //    `document` and match the target by class. This survives
            //    every re-render of the form.
            // ============================================================
            // let backInProgress = false;
            // ============================================================
            // ✅ ONE delegated click listener for Save + Back.
            //    Bound on `document` so it survives form re-renders and
            //    doesn't depend on the buttons existing at script start.
            //    Each flow has its OWN guard inside its OWN try/finally,
            //    so neither can ever get stuck at `true`.
            // ============================================================
            let isSaving = false;
            let isLeavingForm = false;
            // ============================================================
            // ✅ SAVE — confirmation popup, then fetch.
            //    • Swal is forced above every module overlay (didOpen).
            //    • Any stale loader is removed before the dialog opens.
            //    • One try/finally ⇒ the guard can never get stuck.
            // ============================================================
            const saveProcurement = async (e) => {
                var _a;
                (_a = e === null || e === void 0 ? void 0 : e.preventDefault) === null || _a === void 0 ? void 0 : _a.call(e);
                // Guard — only block while a save is in flight
                if (isSaving) {
                    console.log("⛔ saveProcurement blocked: save already in progress");
                    return;
                }
                // --- Pre-checks (nothing shown yet) ---
                const projectId = parseInt(projectSelect.value, 10);
                if (!projectId) {
                    toast("Please select a project", "warning");
                    return;
                }
                const project = projects.find((p) => p.Id === projectId);
                const allowed = editingId ? canEditProject(project) : canCreateRecord();
                if (!allowed) {
                    toast(`❌ You don't have permission to ${editingId ? "edit" : "create"} this procurement plan`, "error");
                    return;
                }
                const items = collectProcurementItems();
                const validationError = validateProcurementItems(items);
                if (validationError) {
                    toast(validationError, "error");
                    return;
                }
                isSaving = true;
                try {
                    const projectName = escapeHtml((project === null || project === void 0 ? void 0 : project.Title) || "this project");
                    // ✅ Kill any leftover loader BEFORE opening the dialog
                    removeLoadingOverlay();
                    // ========== 1) CONFIRMATION ==========
                    console.log("🟣 Opening Save confirmation…");
                    const confirmResult = await Swal.fire({
                        title: editingId ? "Save Changes?" : "Save Procurement Plan?",
                        html: editingId
                            ? `Do you want to save the changes made to <strong>${projectName}</strong>?`
                            : `Do you want to save the procurement plan for <strong>${projectName}</strong>?`,
                        icon: "question",
                        showCancelButton: true,
                        confirmButtonText: editingId ? "Yes, Save Changes" : "Yes, Save Plan",
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
                    // ========== 2) SAVE ==========
                    const loading = createLoadingOverlay("Saving procurement plan...");
                    loaderHost().appendChild(loading);
                    try {
                        const digest = await getRequestDigest();
                        const body = {
                            __metadata: { type: listEntityTypeName },
                            [fieldMappings.ProcurementForecast]: JSON.stringify(items),
                            [`${fieldMappings.Project}Id`]: projectId,
                        };
                        const headers = {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest,
                        };
                        let url = apiBase;
                        let method = "POST";
                        if (editingId) {
                            url = `${apiBase}(${editingId})`;
                            headers["X-HTTP-Method"] = "MERGE";
                            headers["IF-MATCH"] = "*";
                        }
                        const res = await fetch(url, {
                            method,
                            headers,
                            body: JSON.stringify(body),
                        });
                        if (!res.ok) {
                            let detail = "";
                            try {
                                detail = await res.text();
                            }
                            catch (_b) {
                                /* ignore */
                            }
                            throw new Error(`Save failed (${res.status}) ${detail}`);
                        }
                        toast(editingId
                            ? "Procurement plan updated successfully"
                            : "Procurement plan saved successfully", "success");
                        await reloadProcurementData();
                    }
                    catch (error) {
                        console.error("Error saving procurement plan:", error);
                        const message = error instanceof Error ? error.message : String(error);
                        toast(`Failed to save procurement plan: ${message}`, "error");
                    }
                    finally {
                        removeLoadingOverlay();
                    }
                }
                catch (outerError) {
                    console.error("Unexpected error in saveProcurement:", outerError);
                    const message = outerError instanceof Error ? outerError.message : String(outerError);
                    toast(`Unexpected error: ${message}`, "error");
                }
                finally {
                    isSaving = false;
                }
            };
            // ============================================================
            // ✅ BACK — confirmation popup, isolated for direct binding.
            // ============================================================
            const leaveForm = async () => {
                if (isLeavingForm) {
                    console.log("⛔ back blocked: already leaving");
                    return;
                }
                isLeavingForm = true;
                try {
                    // ✅ Kill any leftover loader BEFORE opening the dialog
                    removeLoadingOverlay();
                    console.log("🟣 Opening Leave confirmation…");
                    const confirmResult = await Swal.fire({
                        title: "Leave this page?",
                        html: "Any <strong>unsaved changes</strong> will be lost.<br>" +
                            "Do you want to go back to the procurement table?",
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
                    hideForm();
                    await reloadProcurementData();
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
            // ============================================ //
            // FORECAST DATE VALIDATION                     //
            //   Forecasts must be today or in the future. //
            //   Rejects past dates immediately on input   //
            //   and reverts the field to the last valid   //
            //   value, with a small toast.                //
            // ============================================ //
            // ============================================================
            // ✅ DIRECT button bindings for Save + Back.
            //    Runs ONCE — the buttons live in the static template so
            //    they survive every tbody re-render.
            //    Idempotent: re-running won't double-bind.
            // ============================================================
            const bindSaveAndBackButtons = () => {
                // SAVE — every .proc-save-btn (top + bottom)
                document.querySelectorAll(".proc-save-btn").forEach((btn) => {
                    if (btn.dataset.saveBound === "1")
                        return;
                    btn.dataset.saveBound = "1";
                    btn.addEventListener("click", async (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        console.log("🖱️ Save clicked → opening confirmation");
                        await saveProcurement(e);
                    });
                });
                // BACK — every .back-to-table (top + bottom)
                document.querySelectorAll(".back-to-table").forEach((btn) => {
                    if (btn.dataset.backBound === "1")
                        return;
                    btn.dataset.backBound = "1";
                    btn.addEventListener("click", async (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        console.log("🖱️ Back clicked → opening confirmation");
                        await leaveForm();
                    });
                });
            };
            const wireForecastMinValidation = (row) => {
                const forecastInputs = row.querySelectorAll('input.date-input.forecast-input');
                forecastInputs.forEach((input) => {
                    // Keep a snapshot of the last valid value
                    let lastValid = input.value || "";
                    const validate = () => {
                        // Refresh min in case the page has been open across midnight
                        input.min = getTodayISO();
                        if (isForecastInPast(input.value)) {
                            input.classList.add("forecast-invalid");
                            toast("Forecast dates cannot be in the past. Please pick today or a future date.", "warning");
                            // Revert to the last known-good value
                            input.value = lastValid;
                            // Give the toast time to show, then clear the error style
                            setTimeout(() => input.classList.remove("forecast-invalid"), 1200);
                        }
                        else {
                            input.classList.remove("forecast-invalid");
                            lastValid = input.value || "";
                        }
                    };
                    input.addEventListener("change", validate);
                    input.addEventListener("blur", validate);
                });
            };
            projectSelect.addEventListener("change", (e) => {
                const projectId = parseInt(e.target.value, 10);
                if (projectId) {
                    onProjectChange(projectId);
                }
                else {
                    renderProjectDetails(null);
                    toggleProcurementSections(false);
                }
            });
            // "+ Add Line Item" in the table footer
            const addRowBtn = document.querySelector(".add-procurement-row");
            if (addRowBtn) {
                addRowBtn.addEventListener("click", (e) => {
                    e.preventDefault();
                    addProcurementRow();
                });
            }
            // ✅ Search bar above the line-item table (searches EVERY column)
            if (tableSearchInput) {
                tableSearchInput.addEventListener("input", filterEditTableRows);
                tableSearchInput.addEventListener("search", filterEditTableRows);
                tableSearchInput.addEventListener("keydown", (e) => {
                    if (e.key === "Escape")
                        resetTableSearch();
                });
            }
            if (tableSearchClearBtn) {
                tableSearchClearBtn.addEventListener("click", (e) => {
                    e.preventDefault();
                    resetTableSearch();
                    tableSearchInput === null || tableSearchInput === void 0 ? void 0 : tableSearchInput.focus();
                });
            }
            // ✅ The filter selects no longer do anything on their own - they are inside
            //    the popup and only Apply (or the floating Apply pill) filters the table.
            //    Picking a value just enables Apply / marks the field as set.
            FILTER_FIELDS.forEach(field => {
                const el = FILTER_SELECTS[field];
                if (el) {
                    el.addEventListener("change", () => {
                        syncFilterFieldStyles();
                        syncFilterButtons();
                    });
                }
            });
            if (filterTriggerBtn) {
                filterTriggerBtn.addEventListener("click", () => {
                    isFilterModalOpen() ? closeFilterModal() : openFilterModal();
                });
            }
            // ✕ keeps the draft (the floating Apply pill then offers to apply it)
            if (filterModalCloseBtn)
                filterModalCloseBtn.addEventListener("click", () => closeFilterModal(false));
            if (filterCancelBtn)
                filterCancelBtn.addEventListener("click", () => closeFilterModal());
            if (filterApplyModalBtn)
                filterApplyModalBtn.addEventListener("click", () => { if (!filterApplyModalBtn.disabled)
                    applyFilterModal(); });
            if (filterClearModalBtn)
                filterClearModalBtn.addEventListener("click", () => { if (!filterClearModalBtn.disabled)
                    clearFilterDraft(); });
            if (floatingApplyBtn)
                floatingApplyBtn.addEventListener("click", applyFilterModal);
            if (filterOverlay) {
                filterOverlay.addEventListener("mousedown", (e) => {
                    if (e.target === filterOverlay)
                        closeFilterModal(false); // keep the draft
                });
            }
            document.addEventListener("keydown", (e) => {
                // Esc = light dismiss: keep the draft, the floating Apply pill stays available
                if (e.key === "Escape" && isFilterModalOpen())
                    closeFilterModal(false);
            });
            setupPaginationEventListeners();
            bindSaveAndBackButtons();
            // ========== INITIAL LOADS ==========
            await initApp();
            console.log("✅ Procurement Tracker Module Initialized - Default page size: 25 per page");
        }, 100);
    }
};
//# sourceMappingURL=procurement.js.map
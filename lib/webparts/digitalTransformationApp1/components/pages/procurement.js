import Swal from "sweetalert2";
import * as XLSX from "xlsx-js-style";
import { saveAs } from "file-saver";
import "../styles/common.css";
import "../styles/procurement.css";
import { authManager, hasFullControl, canContribute, canRead, canFreeze, isGHHead, getUserGroups, getCurrentUserFullName, getPermissionLevel, canEditProject, canCreateRecord, canReadAllProjects, getAssignedProjects, isPMManager, isEngineeringUser, updateUIBasedOnPermissions } from "./auth";
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
    <div class="page-header stick">
      <div class="page-header-left">
        <button class="btn-primary show-form-btn">+ Add Procurement</button>
      </div>
      <h1>Procurement Tracker</h1>
    </div>
    
    <!-- 🔍 Filter bar: row 1 = label-less dropdowns + Reset (like the reference) -->
    <div class="table-actions">
      <div class="filter-row">
        <div class="tableColp4">
          <select id="filterProject"><option value="">All Projects</option></select>
        </div>
        <div class="tableColp4">
          <select id="filterCustomer"><option value="">All Customers</option></select>
        </div>
        <div class="tableColp4">
          <select id="filterProjectManager"><option value="">All Managers</option></select>
        </div>
        <div class="tableColp4">
          <select id="filterGroupHead"><option value="">All Group Heads</option></select>
        </div>
        <div class="tableColp4">
          <select id="filterProjectReference"><option value="">All References</option></select>
        </div>
        <div class="tableColp4 action-col">
          <button type="button" class="btn-secondary clear-filters-btn">Reset Filter</button>
        </div>

        <div class="tableColp4 filter-actions-row">
           <button type="button" class="btn-secondary export-excel-btn">📊 Download Excel</button>
        </div>
      </div>

      <!-- 🔘 Row 2: action buttons -->
     

    <!-- ✅ Main table: OUTER-STICKY header -->
    <!--    The page scrolls vertically (not an inner box), so the header  -->
    <!--    runs up, touches the main header and pins below it.             -->
    <div class="table-scroll-wrapper main-table-scroll">
      <table id="procurementTable" class="modern-table">
        <thead class="sticky-main-header">
          <tr>
            <th class="sticky-main-col main-col-1 proc-sortable-th" data-sort="reference" title="Sort by Project Reference"><span class="proc-th-inner"><span class="proc-th-text">Project Reference</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="sticky-main-col main-col-2 proc-sortable-th" data-sort="customer" title="Sort by Customer"><span class="proc-th-inner"><span class="proc-th-text">Customer</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="sticky-main-col main-col-3 proc-sortable-th" data-sort="project" title="Sort by Project"><span class="proc-th-inner"><span class="proc-th-text">Project</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="sticky-main-col main-col-4 proc-sortable-th" data-sort="grouphead" title="Sort by Group Head"><span class="proc-th-inner"><span class="proc-th-text">Group Head</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="sticky-main-col main-col-5 proc-sortable-th" data-sort="manager" title="Sort by Project Manager"><span class="proc-th-inner"><span class="proc-th-text">Project Manager</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="total" title="Sort by Total Line Items"><span class="proc-th-inner"><span class="proc-th-text">Total Line Items</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="rfq" title="Sort by RFQ Pending"><span class="proc-th-inner"><span class="proc-th-text">RFQ Pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="offer" title="Sort by Offer Pending"><span class="proc-th-inner"><span class="proc-th-text">Offer Pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="tc" title="Sort by TC pending"><span class="proc-th-inner"><span class="proc-th-text">TC pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="orderclosure" title="Sort by Order Closure pending"><span class="proc-th-inner"><span class="proc-th-text">Order Closure pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="pr" title="Sort by PR pending"><span class="proc-th-inner"><span class="proc-th-text">PR pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="po" title="Sort by PO pending"><span class="proc-th-inner"><span class="proc-th-text">PO pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="docapproval" title="Sort by Document Approval Pending"><span class="proc-th-inner"><span class="proc-th-text">Document Approval Pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="mfc" title="Sort by MFC Pending"><span class="proc-th-inner"><span class="proc-th-text">MFC Pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="inspection" title="Sort by Inspection Pending"><span class="proc-th-inner"><span class="proc-th-text">Inspection Pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="dispatch" title="Sort by Dispatch Pending"><span class="proc-th-inner"><span class="proc-th-text">Dispatch Pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="delivery" title="Sort by Delivery Pending"><span class="proc-th-inner"><span class="proc-th-text">Delivery Pending</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th class="proc-sortable-th proc-count-th" data-sort="delivered" title="Sort by Delivered"><span class="proc-th-inner"><span class="proc-th-text">Delivered</span><span class="proc-sort-arrows"><span class="proc-arrow-up">&#9650;</span><span class="proc-arrow-down">&#9660;</span></span></span></th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody id="procurementBody"></tbody>
      </table>
    </div>

    <!-- ✅ Loading state (same pattern as DCS / Inventory): spinner + text,
         shown while procurement records are being fetched - same position
         as DCS: below the table (header), above the pagination bar. -->
    <div class="proc-loading" id="procurementLoading">
      <span class="proc-loading-spinner"></span>
      <span class="proc-loading-text">Loading Procurement data...</span>
    </div>

    <!-- ✅ DCS-style pagination bar (same structure as the DCS module):
         LEFT  = « ‹ [Page X of Y] [page tiles] › »
         RIGHT = Show [size] per page + "Showing X-Y of N records" badge. -->
    <div class="proc-pg-container">
      <div class="proc-pg-left">
        <button type="button" class="proc-pg-icon" id="procFirstPage" title="First page">«</button>
        <button type="button" class="proc-pg-icon" id="procPrevPage" title="Previous page">‹</button>

        <div class="proc-pg-info">
          Page <span id="procCurrentPage">1</span> of <span id="procTotalPages">1</span>
        </div>

        <div class="proc-pg-pages" id="procPageNumbers"></div>

        <button type="button" class="proc-pg-icon" id="procNextPage" title="Next page">›</button>
        <button type="button" class="proc-pg-icon" id="procLastPage" title="Last page">»</button>
      </div>

      <div class="proc-pg-right">
        <div class="proc-pg-size">
          <label for="pageSizeSelect">Show:</label>
          <select id="pageSizeSelect" class="proc-pg-size-select">
            <option value="10">10</option>
            <option value="25" selected>25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
          <span>per page</span>
        </div>

        <div class="proc-pg-total" id="paginationInfo">
          Showing <span class="proc-pg-range"><span id="procStartItem">0</span>-<span id="procEndItem">0</span></span>
          of <span class="proc-pg-count" id="procTotalItems">0</span> records
        </div>
      </div>
    </div>
  </section>

  <section class="form-section card" style="display:none;">
  <form id="procurementForm">
  <div class="page-header">
    <h1 id="formTitle">Add Procurement</h1>
    <!-- ✅ TOP SAVE BUTTON - same type="submit" + .btn-primary
         as the bottom button, so both trigger the same submit handler
         and stay in sync (visibility, saving state). The header is
         position:sticky, so this button stays visible while scrolling
         through the long procurement table. -->
    <button type="submit"
            form="procurementForm"
            class="btn-primary proc-btn-save-top"
            title="Save Procurement Plan">💾 Save Procurement Plan</button>
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
          <strong>ℹ️ Existing Data Found!</strong> This project already has procurement data. You can view and edit it below.
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
    script: () => {
        setTimeout(async () => {
            var _a, _b, _c, _d, _e;
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
            const backBtn = document.querySelector(".back-to-table");
            const procurementContainer = document.getElementById("procurementContainer");
            const procurementContainerTable = document.getElementById("procurementContainerTable");
            const tableBody = document.getElementById("procurementBody");
            const form = document.getElementById("procurementForm");
            const formTitle = document.getElementById("formTitle");
            const existingDataAlert = document.getElementById("existingDataAlert");
            // ✅ Search box for the Edit / View Procurement table.
            // (Moved OUT of the Description header into its own search bar above the
            // table - it searches EVERY column, not just Description.)
            const tableSearchInput = document.getElementById("procTableSearch");
            const tableSearchClearBtn = document.getElementById("procTableSearchClear");
            const tableSearchCount = document.getElementById("procTableSearchCount");
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
            const clearFiltersBtn = document.querySelector(".table-actions .clear-filters-btn");
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
            // ============================================
            // LOADING OVERLAY (same pattern as master.ts)
            // ============================================
            const createLoadingOverlay = (message = "Processing...") => {
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
                    if (canEditProject(project))
                        return true;
                    if (canReadAllProjects())
                        return true;
                    if (project &&
                        getAssignedProjects(projects).some((p) => p.Id === project.Id))
                        return true;
                    return false;
                }
                return canCreateRecord();
            };
            // Show / hide the Save button based on the current role + mode
            const applySaveButtonVisibility = (isEditMode, project) => {
                // ✅ Applies to BOTH the top button (header) and the bottom button (form-actions)
                const canSave = isEditMode
                    ? canEditProject(project || undefined)
                    : canCreateRecord();
                getSaveButtons().forEach(btn => {
                    btn.style.display = canSave ? "inline-flex" : "none";
                });
            };
            const getSaveButtons = () => Array.from(document.querySelectorAll('button.btn-primary[type="submit"][form="procurementForm"], .form-actions button.btn-primary[type="submit"]'));
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
                const managerName = getProjectManagerDisplayName(project.ProjectManager);
                const groupHeadName = getGroupHeadDisplayName(project.GroupHead1);
                if (projectNameDisplay)
                    projectNameDisplay.textContent = project.Title || "-";
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
                    attachFilterEventListeners();
                    if (!dataLoaded)
                        return;
                    initializeFilters();
                    renderTable(procurementData);
                    toast("Procurement tracker initialized successfully", "success");
                    // ✅ Apply role-based UI visibility (same as SEIBIES)
                    updateUIBasedOnPermissions({
                        exportBtn: ".export-excel-btn",
                        addBtn: ".show-form-btn",
                        saveBtn: 'button[type="submit"]'
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
            const attachFilterEventListeners = () => {
                filterProject.addEventListener("change", filterData);
                filterCustomer.addEventListener("change", filterData);
                filterProjectManager.addEventListener("change", filterData);
                filterGroupHead.addEventListener("change", filterData);
                filterProjectReference.addEventListener("change", filterData);
                const clearFiltersBtn = document.querySelector(".table-actions .clear-filters-btn");
                if (clearFiltersBtn) {
                    clearFiltersBtn.addEventListener("click", clearFilters);
                }
            };
            const loadProjects = async () => {
                var _a;
                try {
                    const listPath = `${siteUrl}/_api/web/lists/getbytitle('Project List')/items`;
                    const headers = { Accept: "application/json;odata=verbose" };
                    // ✅ Preferred query: GroupHead1 / ProjectManager as Person fields (expanded)
                    let res = await fetch(`${listPath}?$select=Id,Title,CustomerName,GroupHead1/Title,GroupHead1Id,PC,ProjectReference,ProjectManager/Title,ProjectManagerId&$expand=GroupHead1,ProjectManager&$top=5000`, { headers });
                    // ✅ Fallback: if the expand fails (e.g. those columns are plain text, or
                    // a column name differs) retrieve the plain columns instead of failing.
                    if (!res.ok) {
                        console.warn(`Project List expanded query failed (${res.status}); retrying without $expand`);
                        res = await fetch(`${listPath}?$select=Id,Title,CustomerName,GroupHead1,PC,ProjectReference,ProjectManager&$top=5000`, { headers });
                    }
                    if (!res.ok) {
                        console.warn(`Project List plain query failed (${res.status}); retrying with all columns`);
                        res = await fetch(`${listPath}?$top=5000`, { headers });
                    }
                    if (!res.ok)
                        throw new Error(`Failed to load projects: ${res.status}`);
                    const data = await res.json();
                    const allProjects = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                    // Map to ProjectItem with proper Person/Group handling
                    const mappedProjects = allProjects.map((p) => ({
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
                    const uniqueProjects = Array.from(new Map(mappedProjects.map((p) => [p.Title, p])).values());
                    projects = uniqueProjects;
                    console.log(`Loaded ${projects.length} projects from Project List`);
                    // ✅ Role access (same as SEIBIES): dropdowns only show assigned projects
                    const assignedProjects = getAssignedProjects(projects);
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
                return groupHead;
            };
            const getProjectManagerDisplayName = (manager) => {
                if (!manager)
                    return "";
                if (typeof manager === "object" && manager !== null) {
                    return manager.Title || "";
                }
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
            // spinner + "Loading Procurement data..." while records are being fetched.
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
                    renderTable(procurementData);
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
                const name = (project === null || project === void 0 ? void 0 : project.Title) || "";
                const ref = (project === null || project === void 0 ? void 0 : project.ProjectReference) || "";
                formTitle.textContent = name
                    ? `${name}${ref ? ` (${ref})` : ""}`
                    : "Edit Procurement Plan";
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
                    btn.textContent = "💾 Save Procurement";
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
            // PAGINATION INFO (DCS-style)                  //
            //   Fills the "Showing X-Y of N records" badge //
            //   (range + count chips in the size group).   //
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
            // PAGINATION (DCS-style)                       //
            //   « ‹ windowed page tiles › » + info pill.   //
            //   Markup lives in procurementContent; this   //
            //   only refreshes state + tiles.              //
            // ============================================ //
            const calcTotalPages = (total) => Math.max(1, Math.ceil(total / pageSize));
            const renderPagination = (totalPages) => {
                const safeTotal = Math.max(1, totalPages);
                const currentPageSpan = document.getElementById("procCurrentPage");
                const totalPagesSpan = document.getElementById("procTotalPages");
                if (currentPageSpan)
                    currentPageSpan.textContent = currentPage.toString();
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
            (_a = document.getElementById("procFirstPage")) === null || _a === void 0 ? void 0 : _a.addEventListener("click", () => goToProcPage(1));
            (_b = document.getElementById("procPrevPage")) === null || _b === void 0 ? void 0 : _b.addEventListener("click", () => goToProcPage(currentPage - 1));
            (_c = document.getElementById("procNextPage")) === null || _c === void 0 ? void 0 : _c.addEventListener("click", () => goToProcPage(currentPage + 1));
            (_d = document.getElementById("procLastPage")) === null || _d === void 0 ? void 0 : _d.addEventListener("click", () => {
                const total = currentTableData.length ? currentTableData : procurementData;
                goToProcPage(calcTotalPages(total.length));
            });
            // ============================================ //
            // COLUMN SORTING - first 5 columns             //
            //   (Project Reference / Customer / Project /  //
            //   Group Head / Project Manager)              //
            //   Same behaviour as DCS: click toggles       //
            //   asc -> desc, case-insensitive text sort,   //
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
                currentPage = 1; // same as DCS: sorting resets to page 1
                renderTable(currentTableData);
                updateSortIndicators();
            };
            const updateSortIndicators = () => {
                document.querySelectorAll('#procurementTable thead th[data-sort]').forEach((th) => {
                    const el = th;
                    el.classList.remove('proc-sorted-asc', 'proc-sorted-desc');
                    if (el.getAttribute('data-sort') === sortKey) {
                        el.classList.add(sortDir === 'asc' ? 'proc-sorted-asc' : 'proc-sorted-desc');
                    }
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
                formTitle.textContent = "View Procurement";
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
            const importExcelForEdit = (projectId, record) => {
                // ✅ FIX 1: remove any leftover modal from a previous open (idempotent -
                // previously every click appended ANOTHER unstyled modal to the page).
                document.querySelectorAll(".import-modal").forEach((m) => m.remove());
                // ✅ FIX 2: build the modal with INLINE STYLES so it works even though
                // procurement.css has no .import-modal classes (same approach as the
                // working Collection import modal). Previously the unstyled modal was
                // appended at the very bottom of the page - invisible on click.
                const modal = document.createElement("div");
                modal.className = "import-modal";
                modal.style.cssText = `
          position: fixed;
          top: 0;
          left: 0;
          width: 100%;
          height: 100%;
          background: rgba(0, 0, 0, 0.6);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 999999;
          overflow-y: auto;
          padding: 20px;
        `;
                modal.innerHTML = `
          <div class="import-modal-content" style="
            background: white;
            border-radius: 12px;
            padding: 30px;
            max-width: 850px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
            position: relative;
          ">
            <div class="import-modal-header" style="
              display: flex;
              justify-content: space-between;
              align-items: center;
              border-bottom: 2px solid #007bff;
              padding-bottom: 15px;
              margin-bottom: 25px;
            ">
              <h3 style="margin: 0; color: #495057; font-size: 20px;">Import Procurement Data from Excel</h3>
              <button class="close-modal" type="button" style="
                background: none;
                border: none;
                font-size: 28px;
                cursor: pointer;
                color: #6c757d;
                padding: 0 10px;
                line-height: 1;
              ">&times;</button>
            </div>

            <div class="import-section" style="margin-bottom: 25px;">
              <h4 style="color: #495057; margin-bottom: 12px; font-size: 15px;">Step 1: Download Template</h4>
              <p style="color: #6c757d; font-size: 13px; margin: 0 0 12px 0;">Download the template file with the correct structure for import.</p>
              <button class="btn-primary" id="downloadEditTemplate" type="button" style="
                background: #0078d4;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                cursor: pointer;
                font-size: 14px;
              ">⬇️ Download Template</button>
            </div>

            <div class="import-section" style="margin-bottom: 25px;">
              <h4 style="color: #495057; margin-bottom: 12px; font-size: 15px;">Step 2: Upload Excel File</h4>
              <div class="file-upload-area" id="fileUploadArea" style="
                border: 2px dashed #ced4da;
                border-radius: 10px;
                padding: 40px 20px;
                text-align: center;
                cursor: pointer;
                transition: all 0.3s ease;
                background: #f8f9fa;
              ">
                <div style="font-size: 44px; margin-bottom: 10px;">📁</div>
                <p style="margin: 0; color: #495057; font-size: 15px;"><strong>Click to upload or drag &amp; drop Excel file</strong></p>
                <p style="margin: 5px 0 0 0; font-size: 12px; color: #6c757d;">Supported format: .xlsx, .xls (Max: 10MB)</p>
                <input type="file" id="excelFile" class="file-input" accept=".xlsx,.xls" style="display: none;">
                <div class="file-info" id="fileInfo" style="margin-top: 10px; font-size: 14px; color: #28a745;"></div>
              </div>
            </div>

            <div class="import-section" id="previewSection" style="display: none; margin-bottom: 25px;">
              <h4 style="color: #495057; margin-bottom: 12px; font-size: 15px;">Step 3: Preview &amp; Validate</h4>
              <div class="import-preview" id="importPreview"></div>
              <div class="import-actions" style="display: flex; gap: 10px; margin-top: 15px;">
                <button class="btn-secondary" id="cancelImport" type="button" style="
                  background: #6c757d;
                  color: white;
                  border: none;
                  padding: 10px 20px;
                  border-radius: 6px;
                  cursor: pointer;
                  font-size: 14px;
                ">Cancel</button>
                <button class="btn-primary" id="confirmImport" type="button" disabled style="
                  background: #28a745;
                  color: white;
                  border: none;
                  padding: 10px 20px;
                  border-radius: 6px;
                  cursor: pointer;
                  font-size: 14px;
                  opacity: 0.6;
                ">Import Data</button>
              </div>
              <div class="import-status" id="importStatus" style="
                margin-top: 10px;
                padding: 12px;
                border-radius: 6px;
                display: none;
              "></div>
            </div>
          </div>
        `;
                document.body.appendChild(modal);
                // Click on the dark backdrop closes the modal
                modal.addEventListener("click", (e) => {
                    if (e.target === modal)
                        modal.remove();
                });
                const fileUploadArea = document.getElementById("fileUploadArea");
                const excelFileInput = document.getElementById("excelFile");
                const fileInfo = document.getElementById("fileInfo");
                const previewSection = document.getElementById("previewSection");
                const confirmImport = document.getElementById("confirmImport");
                const cancelImport = document.getElementById("cancelImport");
                const closeModal = document.querySelector(".close-modal");
                const importStatus = document.getElementById("importStatus");
                const downloadTemplateBtn = document.getElementById("downloadEditTemplate");
                const setConfirmEnabled = (enabled) => {
                    confirmImport.disabled = !enabled;
                    confirmImport.style.opacity = enabled ? "1" : "0.6";
                    confirmImport.style.cursor = enabled ? "pointer" : "not-allowed";
                };
                let importData = [];
                let importHeaders = [];
                fileUploadArea.addEventListener("click", () => excelFileInput.click());
                fileUploadArea.addEventListener("dragover", (e) => {
                    e.preventDefault();
                    fileUploadArea.style.borderColor = "#0078d4";
                    fileUploadArea.style.background = "#e7f3ff";
                });
                fileUploadArea.addEventListener("dragleave", () => {
                    fileUploadArea.style.borderColor = "#ced4da";
                    fileUploadArea.style.background = "#f8f9fa";
                });
                fileUploadArea.addEventListener("drop", (e) => {
                    var _a;
                    e.preventDefault();
                    fileUploadArea.style.borderColor = "#ced4da";
                    fileUploadArea.style.background = "#f8f9fa";
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
                    setConfirmEnabled(false);
                    if (!file.name.toLowerCase().match(/\.(xlsx|xls)$/)) {
                        fileInfo.textContent = 'Please select an Excel file (.xlsx or .xls)';
                        fileInfo.className = 'file-info error';
                        fileInfo.style.color = '#dc3545';
                        return;
                    }
                    if (file.size > 10 * 1024 * 1024) {
                        fileInfo.textContent = 'File size too large. Please use files under 10MB.';
                        fileInfo.className = 'file-info error';
                        fileInfo.style.color = '#dc3545';
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
                            fileInfo.style.color = '#dc3545';
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
                                setConfirmEnabled(false);
                                importStatus.className = 'import-status error';
                                importStatus.style.display = 'block';
                                importStatus.style.background = '#f8d7da';
                                importStatus.style.color = '#721c24';
                                importStatus.style.border = '1px solid #f5c6cb';
                                importStatus.textContent = '❌ Critical errors found. Please fix the issues and try again.';
                                return;
                            }
                        }
                        setConfirmEnabled(true);
                        previewSection.style.display = "block";
                        importStatus.className = 'import-status success';
                        importStatus.style.display = 'block';
                        importStatus.style.background = '#d4edda';
                        importStatus.style.color = '#155724';
                        importStatus.style.border = '1px solid #c3e6cb';
                        importStatus.textContent = `✅ Found ${parsedData.length} valid row(s) to import.`;
                    }
                    catch (error) {
                        console.error("Excel parsing error:", error);
                        fileInfo.textContent = `Error reading Excel file: ${error}`;
                        fileInfo.className = 'file-info error';
                        fileInfo.style.color = '#dc3545';
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
                    setConfirmEnabled(false);
                    confirmImport.textContent = "Importing...";
                    importStatus.className = "import-status info";
                    importStatus.style.display = "block";
                    importStatus.style.background = "#e7f3ff";
                    importStatus.style.color = "#0056b3";
                    importStatus.style.border = "1px solid #b8daff";
                    importStatus.textContent = "⏳ Processing procurement data...";
                    try {
                        const procurementItems = importData.map((row) => {
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
                            inputsMapper(row, item);
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
                        setConfirmEnabled(true);
                        confirmImport.textContent = "Import Data";
                    }
                };
                // Maps the parsed Excel row onto a ProcurementForecastItem
                // (same field mapping as the original implementation)
                const inputsMapper = (row, item) => {
                    item.Description = row.description || "";
                    item.FCMQuantity = row.fcmquantity ? parseFloat(row.fcmquantity) : null;
                    item.EnggQuantity = row.enggquantity ? parseFloat(row.enggquantity) : null;
                    item.Make = row.make || row.manufacturer || "";
                    item.Model = row.model || row.partnumber || "";
                    item.Remarks = row.remarks || "";
                    item.RfqPlan = row.rfqplan || "";
                    item.RfqActual = row.rfqactual || "";
                    item.RfqForecast = row.rfqforecast || "";
                    item.OfferPlan = row.offerplan || "";
                    item.OfferActual = row.offeractual || "";
                    item.OfferForecast = row.offerforecast || "";
                    item.TechPlan = row.techplan || "";
                    item.TechActual = row.techactual || "";
                    item.TechForecast = row.techforecast || "";
                    item.OrderPlan = row.orderplan || "";
                    item.OrderActual = row.orderactual || "";
                    item.OrderForecast = row.orderforecast || "";
                    item.PRPlan = row.prplan || "";
                    item.PRActual = row.practual || "";
                    item.PRForecast = row.prforecast || "";
                    item.POPlan = row.poplan || "";
                    item.POActual = row.poactual || "";
                    item.POForecast = row.poforecast || "";
                    item.DocPlan = row.docplan || "";
                    item.DocActual = row.docactual || "";
                    item.DocForecast = row.docforecast || "";
                    item.MfgPlan = row.mfgplan || "";
                    item.MfgActual = row.mfgactual || "";
                    item.MfgForecast = row.mfgforecast || "";
                    item.InspPlan = row.inspplan || "";
                    item.InspActual = row.inspactual || "";
                    item.InspForecast = row.inspforecast || "";
                    item.DispPlan = row.dispplan || "";
                    item.DispActual = row.dispactual || "";
                    item.DispForecast = row.dispforecast || "";
                    item.DeliveryPlan = row.deliveryplan || "";
                    item.DeliveryActual = row.deliveryactual || "";
                    item.DeliveryForecast = row.deliveryforecast || "";
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
                        setEditTitle(project);
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
                        formSection.classList.remove("edit-mode");
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
                // Initial overdue class for an Actual input (Plan crossed & Actual empty)
                const actualCls = (planField, actualField) => isActualOverdue(getFieldValue(planField), getFieldValue(actualField))
                    ? " actual-overdue"
                    : "";
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
          <td><input type="date" name="rfq_actual" class="form-control${actualCls("RfqPlan", "RfqActual")}" value="${getFieldValue("RfqActual")}"></td>
          <td><input type="date" name="rfq_forecast" class="form-control" value="${getFieldValue("RfqForecast")}"></td>

          <!-- Vendor Offer / SCM -->
          <td><input type="date" name="offer_plan" class="form-control" value="${getFieldValue("OfferPlan")}"></td>
          <td><input type="date" name="offer_actual" class="form-control${actualCls("OfferPlan", "OfferActual")}" value="${getFieldValue("OfferActual")}"></td>
          <td><input type="date" name="offer_forecast" class="form-control" value="${getFieldValue("OfferForecast")}"></td>

          <!-- Technical Clearance / ENGG -->
          <td><input type="date" name="tech_plan" class="form-control" value="${getFieldValue("TechPlan")}"></td>
          <td><input type="date" name="tech_actual" class="form-control${actualCls("TechPlan", "TechActual")}" value="${getFieldValue("TechActual")}"></td>
          <td><input type="date" name="tech_forecast" class="form-control" value="${getFieldValue("TechForecast")}"></td>

          <!-- Order Finalization / SCM -->
          <td><input type="date" name="order_plan" class="form-control" value="${getFieldValue("OrderPlan")}"></td>
          <td><input type="date" name="order_actual" class="form-control${actualCls("OrderPlan", "OrderActual")}" value="${getFieldValue("OrderActual")}"></td>
          <td><input type="date" name="order_forecast" class="form-control" value="${getFieldValue("OrderForecast")}"></td>

          <!-- PR / PM -->
          <td><input type="date" name="pr_plan" class="form-control" value="${getFieldValue("PRPlan")}"></td>
          <td><input type="date" name="pr_actual" class="form-control${actualCls("PRPlan", "PRActual")}" value="${getFieldValue("PRActual")}"></td>
          <td><input type="date" name="pr_forecast" class="form-control" value="${getFieldValue("PRForecast")}"></td>

          <!-- Purchase Order / SCM -->
          <td><input type="date" name="po_plan" class="form-control" value="${getFieldValue("POPlan")}"></td>
          <td><input type="date" name="po_actual" class="form-control${actualCls("POPlan", "POActual")}" value="${getFieldValue("POActual")}"></td>
          <td><input type="date" name="po_forecast" class="form-control" value="${getFieldValue("POForecast")}"></td>

          <!-- Document Approval / ENGG -->
          <td><input type="date" name="doc_plan" class="form-control" value="${getFieldValue("DocPlan")}"></td>
          <td><input type="date" name="doc_actual" class="form-control${actualCls("DocPlan", "DocActual")}" value="${getFieldValue("DocActual")}"></td>
          <td><input type="date" name="doc_forecast" class="form-control" value="${getFieldValue("DocForecast")}"></td>

          <!-- Manufacturing Clearance / ENGG -->
          <td><input type="date" name="mfg_plan" class="form-control" value="${getFieldValue("MfgPlan")}"></td>
          <td><input type="date" name="mfg_actual" class="form-control${actualCls("MfgPlan", "MfgActual")}" value="${getFieldValue("MfgActual")}"></td>
          <td><input type="date" name="mfg_forecast" class="form-control" value="${getFieldValue("MfgForecast")}"></td>

          <!-- Inspection -->
          <td><input type="date" name="insp_plan" class="form-control" value="${getFieldValue("InspPlan")}"></td>
          <td><input type="date" name="insp_actual" class="form-control${actualCls("InspPlan", "InspActual")}" value="${getFieldValue("InspActual")}"></td>
          <td><input type="date" name="insp_forecast" class="form-control" value="${getFieldValue("InspForecast")}"></td>

          <!-- Material Dispatch -->
          <td><input type="date" name="disp_plan" class="form-control" value="${getFieldValue("DispPlan")}"></td>
          <td><input type="date" name="disp_actual" class="form-control${actualCls("DispPlan", "DispActual")}" value="${getFieldValue("DispActual")}"></td>
          <td><input type="date" name="disp_forecast" class="form-control" value="${getFieldValue("DispForecast")}"></td>

          <!-- Delivery at Site -->
          <td><input type="date" name="delivery_plan" class="form-control" value="${getFieldValue("DeliveryPlan")}"></td>
          <td><input type="date" name="delivery_actual" class="form-control${actualCls("DeliveryPlan", "DeliveryActual")}" value="${getFieldValue("DeliveryActual")}"></td>
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
                // ✅ Disable BOTH save buttons while saving
                const saveBtns = getSaveButtons();
                saveBtns.forEach(btn => {
                    btn.disabled = true;
                    btn.dataset.saving = "true";
                });
                const saveLoading = createLoadingOverlay("Saving procurement data...");
                loaderHost().appendChild(saveLoading);
                try {
                    const digest = await getRequestDigest();
                    const projectId = parseInt(projectSelect.value);
                    const submittingProject = projects.find((p) => p.Id === projectId);
                    // ✅ Role access on save (same as SEIBIES)
                    if (editingId) {
                        if (!canEditProject(submittingProject)) {
                            toast("❌ You don't have permission to edit this procurement plan", "error");
                            return;
                        }
                    }
                    else {
                        if (!canCreateRecord()) {
                            toast("❌ You don't have permission to create procurement plans", "error");
                            return;
                        }
                    }
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
                finally {
                    removeLoadingOverlay();
                    // ✅ Re-enable BOTH save buttons
                    saveBtns.forEach(btn => {
                        btn.disabled = false;
                        btn.dataset.saving = "false";
                    });
                }
            });
            const resetForm = () => {
                resetTableSearch();
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
                formSection.classList.remove("edit-mode");
                const actionBtns = document.querySelectorAll(".export-project-excel-btn, .export-top-btn, .import-excel-edit-btn");
                actionBtns.forEach((btn) => btn.remove());
                const addMoreBtn = document.querySelector(".add-procurement-row");
                if (addMoreBtn)
                    addMoreBtn.style.display = "block";
                getSaveButtons().forEach(btn => {
                    btn.style.display = "inline-flex";
                });
                updateMainTableStickyOffset();
            };
            const exportProjectToExcel = async (projectId, record) => {
                // ✅ Role guard (same as SEIBIES): export needs at least read access
                if (!canRead()) {
                    toast("❌ You don't have permission to export data", "error");
                    return;
                }
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
                    const managerName = getProjectManagerDisplayName(project.ProjectManager);
                    const groupHeadName = getGroupHeadDisplayName(project.GroupHead1);
                    // Sheet 1: Project Summary
                    const projectSummaryData = [
                        ["PROJECT SUMMARY", ""],
                        ["", ""],
                        ["Project Information", ""],
                        ["Project Name", project.Title || ""],
                        ["Customer", project.CustomerName || ""],
                        ["Project Manager", managerName],
                        ["Group Head", groupHeadName],
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
                // ✅ Role guard (same as SEIBIES): export needs at least read access
                if (!canRead()) {
                    toast("❌ You don't have permission to export data", "error");
                    return;
                }
                showExportProgress("Preparing data for export...");
                try {
                    const project = filterProject.value;
                    const customer = filterCustomer.value;
                    const projectManager = filterProjectManager.value;
                    const groupHead = filterGroupHead.value;
                    const projectReference = filterProjectReference.value;
                    const filteredData = procurementData.filter((item) => {
                        var _a, _b;
                        const projectData = resolveProjectForItem(item);
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || (projectData === null || projectData === void 0 ? void 0 : projectData.Title) || "";
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
                        if (groupHead) {
                            const groupHeadName = getGroupHeadDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead1);
                            if (groupHeadName !== groupHead)
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
                    const headers = [
                        "Project Reference", "Customer", "Project", "Group Head", "Project Manager",
                        "Total Line Items",
                        "RFQ Pending", "Offer Pending", "TC pending", "Order Closure pending",
                        "PR pending", "PO pending", "Document Approval Pending", "MFC Pending",
                        "Inspection Pending", "Dispatch Pending", "Delivery Pending", "Delivered"
                    ];
                    const dataRows = [];
                    filteredData.forEach((item) => {
                        var _a, _b;
                        const projectData = resolveProjectForItem(item);
                        const projectTitle = ((_a = item[fieldMappings.Project]) === null || _a === void 0 ? void 0 : _a.Title) || (projectData === null || projectData === void 0 ? void 0 : projectData.Title) || "";
                        const procurementItems = parseProcurementForecast(item[fieldMappings.ProcurementForecast]) || [];
                        const statusCounts = calculateProjectStatusCounts(procurementItems);
                        const totalLineItems = procurementItems.length;
                        const managerName = typeof (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === "object"
                            ? ((_b = projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) === null || _b === void 0 ? void 0 : _b.Title) || ""
                            : (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectManager) || "";
                        const groupHeadName = getGroupHeadDisplayName(projectData === null || projectData === void 0 ? void 0 : projectData.GroupHead1);
                        dataRows.push([
                            (projectData === null || projectData === void 0 ? void 0 : projectData.ProjectReference) || "",
                            (projectData === null || projectData === void 0 ? void 0 : projectData.CustomerName) || "",
                            projectTitle,
                            groupHeadName,
                            managerName,
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
                        projectManager ? `Manager-${projectManager}` : "",
                        groupHead ? `Group-${groupHead}` : "",
                        projectReference ? `Reference-${projectReference}` : "",
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
                // ✅ Role access (same as SEIBIES): create requires create rights
                if (!canViewProcurementForm(false)) {
                    toast("❌ You don't have permission to create procurement plans", "error");
                    return;
                }
                resetForm();
                formSection.style.display = "block";
                tableSection.style.display = "none";
                renderProjectDetails(null);
                procurementContainerTable.style.display = "block";
                procurementContainer.innerHTML = "";
                procurementContainer.appendChild(createProcurementRow());
                formTitle.textContent = "Add Procurement";
                existingDataAlert.style.display = "none";
                resetTableSearch();
                projectSelect.disabled = false;
                projectSelect.value = "";
                applySaveButtonVisibility(false);
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
                        setEditTitle(projects.find((p) => p.Id === selectedId) || null);
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
                        formSection.classList.remove("edit-mode");
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
                    formSection.classList.remove("edit-mode");
                    procurementContainer.innerHTML = "";
                    procurementContainer.appendChild(createProcurementRow());
                }
            });
            (_e = document
                .querySelector(".add-procurement-row")) === null || _e === void 0 ? void 0 : _e.addEventListener("click", () => {
                procurementContainer.appendChild(createProcurementRow());
                filterEditTableRows();
            });
            tableSearchInput === null || tableSearchInput === void 0 ? void 0 : tableSearchInput.addEventListener("input", filterEditTableRows);
            // Clear (×) button inside the search box
            tableSearchClearBtn === null || tableSearchClearBtn === void 0 ? void 0 : tableSearchClearBtn.addEventListener("click", () => {
                if (tableSearchInput)
                    tableSearchInput.value = "";
                filterEditTableRows();
                tableSearchInput === null || tableSearchInput === void 0 ? void 0 : tableSearchInput.focus();
            });
            // Esc clears the search
            tableSearchInput === null || tableSearchInput === void 0 ? void 0 : tableSearchInput.addEventListener("keydown", (e) => {
                if (e.key === "Escape") {
                    if (tableSearchInput)
                        tableSearchInput.value = "";
                    filterEditTableRows();
                }
            });
            // ✅ Any edit in ANY column refreshes that row's searchable text and
            // re-applies the active search (previously only the Description input
            // triggered a re-filter).
            procurementContainer.addEventListener("input", (e) => {
                var _a;
                const row = (_a = e.target) === null || _a === void 0 ? void 0 : _a.closest("tr");
                invalidateRowSearchText(row);
                if (tableSearchInput && tableSearchInput.value.trim() !== "") {
                    filterEditTableRows();
                }
            });
            // Date pickers / selects fire "change" (not always "input")
            procurementContainer.addEventListener("change", (e) => {
                var _a;
                const row = (_a = e.target) === null || _a === void 0 ? void 0 : _a.closest("tr");
                invalidateRowSearchText(row);
                if (tableSearchInput && tableSearchInput.value.trim() !== "") {
                    filterEditTableRows();
                }
            });
            procurementContainer.addEventListener("click", (e) => {
                const target = e.target;
                if (target.classList.contains("remove-row")) {
                    const row = target.closest("tr");
                    if (row && procurementContainer.children.length > 1) {
                        invalidateRowSearchText(row);
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
            clearFiltersBtn.addEventListener("click", clearFilters);
            exportExcelBtn.addEventListener("click", exportToExcel);
            // ✅ Page size selector (10 / 25 / 50 / 100) - like SEIBIES
            const pageSizeSelect = document.getElementById("pageSizeSelect");
            if (pageSizeSelect) {
                pageSizeSelect.value = pageSize.toString();
                pageSizeSelect.addEventListener("change", (e) => {
                    const newSize = parseInt(e.target.value, 10);
                    if (!newSize || newSize === pageSize)
                        return;
                    pageSize = newSize;
                    currentPage = 1;
                    renderTable(currentTableData.length ? currentTableData : procurementData);
                    toast(`Showing ${pageSize} records per page`, "success");
                });
            }
            // Re-apply the outer-sticky offset when the window is resized
            window.addEventListener("resize", () => {
                clearTimeout(window._procurementStickyResizeTimer);
                window._procurementStickyResizeTimer = setTimeout(updateMainTableStickyOffset, 150);
            });
            // Re-apply when coming back from the form (table is visible again)
            updateMainTableStickyOffset();
            await initApp();
        }, 100);
    },
};
//# sourceMappingURL=procurement.js.map
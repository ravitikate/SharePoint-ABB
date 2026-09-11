declare const _spPageContextInfo: any;
import Swal from "sweetalert2";
import * as XLSX from 'xlsx-js-style';
import "../styles/common.css";
import "../styles/seibies.css";
import { 
  authManager,
  hasFullControl,
  canContribute,
  canRead,
  canFreeze,
  isGHHead,
  getUserGroups,
  getCurrentUserFullName,
  getPermissionLevel,
  canEditProject,
  canCreateRecord,
  canReadAllProjects,
  getAssignedProjects,
  isPMManager,
  isEngineeringUser,
  updateUIBasedOnPermissions 
} from "./auth";

interface ProjectItem {
  Id: number;
  Title: string;
  CustomerName?: string;
  ProjectManagerId?: number;
  ProjectManager?: string | { Title: string; Email?: string };
  PC?: string;
  ProjectReference?: string;
  GroupHead1?: string | { Title: string; Email?: string };
  GroupHead1Id?: number;
}

interface RawProjectItem {
  Id: number;
  Title: string;
  CustomerName?: string;
  GroupHead1Id?: number;
  PC?: string;
  ProjectReference?: string;
  ProjectManagerId?: number;
  GroupHead1?: { Title: string; Email?: string };
  ProjectManager?: { Title: string; Email?: string };
  [key: string]: any;
}

interface SortConfig {
  column: string;
  direction: 'asc' | 'desc';
  subColumn?: string;
  monthIndex?: number;
}

let currentSortConfig: SortConfig = {
  column: '',
  direction: 'asc'
};

interface SeibiesForecastItem {
  Id: number;
  Title: string;
  Project?: {
    Id: number;
    Title: string;
  };
  ProfitCentre?: string;
  OBL?: number;
  ForecastData?: string;
  SeibiesActual?: string;
  CreatedDate?: string;
  UpdatedDate?: string;
  [key: string]: any;
}

interface SharePointResponse<T> {
  d: {
    results: T[];
  };
}

interface ForecastMonthData {
  plan: number;
  actual: number;
  forecast?: number;
  date?: string;
}

interface YearlyForecastData {
  [year: string]: {
    plan: number;
    actual: number;
    forecast?: number;
  };
}

interface ForecastData {
  [monthKey: string]: ForecastMonthData | YearlyForecastData | undefined;
  yearly?: YearlyForecastData;
}

export default {
  content: `
<div class="seibies-page fade-in">

  <!-- ============================================ -->
  <!-- LOADER - covers ONLY this section              -->
  <!-- ============================================ -->
  <div id="revenuePageLoader" class="revenue-loader" style="position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;background:rgba(248,249,250,0.92);z-index:999;">
    <div class="revenue-loader-box">
      <div class="revenue-loader-spinner"></div>
      <div class="revenue-loader-text">Loading SEI-BIES Data</div>
      <div class="revenue-loader-subtext">Please wait... fetching SEIBIES projects and forecasts</div>
      <div class="revenue-loader-bar"><span></span></div>
    </div>
  </div>
  <section class="table-section-seibies card-seibies">
    <div class="page-header">
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
          <label> </label>
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
          <label> </label>
          <button class="btn-secondary-seibies export-btn-seibies">Download Excel</button>
        </div> 
      </div>
    </div>
       
    <div class="table-container-seibies">
      <div class="table-scroll-wrapper-seibies">
        <table class="month-columns-table-seibies">
          <thead id="totalsHeaderSeibies" class="sticky-totals-header"></thead>
          <thead id="monthColumnsHeaderSeibies" class="sticky-main-header"></thead>
          <tbody class="seibies-table-body">
            <tr class="seibies-table-loading-row">
              <td colspan="25">
                <span class="seibies-table-loading-spinner"></span>Loading SEIBIES data…
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      
      <div class="pagination-container-seibies">
        <div class="pagination-row-seibies">
          <div class="pagination-left-seibies">
            <button class="pagination-btn-seibies seibies-pg-icon" id="firstPageSeibies" title="First page">«</button>
            <button class="pagination-btn-seibies seibies-pg-icon" id="prevPageSeibies" title="Previous page">‹</button>

            <div class="pagination-info-seibies seibies-pg-info">
              Page <span id="currentPageSeibies">1</span> of <span id="totalPagesSeibies">1</span>
            </div>

            <div class="pagination-pages-seibies" id="pageNumbersSeibies"></div>

            <button class="pagination-btn-seibies seibies-pg-icon" id="nextPageSeibies" title="Next page">›</button>
            <button class="pagination-btn-seibies seibies-pg-icon" id="lastPageSeibies" title="Last page">»</button>
          </div>
        </div>

        <div class="pagination-row-seibies pagination-bottom-row-seibies">
          <div class="pagination-right-seibies">
            <div class="page-size-selector-seibies seibies-pg-size">
              <label for="pageSizeSelectSeibies">Show:</label>
              <select id="pageSizeSelectSeibies" class="page-size-select-seibies">
                <option value="10">10</option>
                <option value="25" selected>25</option>
                <option value="50">50</option>
                <option value="100">100</option>
              </select>
              <span>per page</span>
            </div>

            <div class="pagination-info-seibies seibies-pg-total">
              Showing <span class="seibies-pg-range"><span id="startItemSeibies">0</span>-<span id="endItemSeibies">0</span></span>
              of <span class="seibies-pg-count" id="totalItemsSeibies">0</span> records
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>

  <section class="form-section-seibies card-seibies" style="display:none;">
  <form id="seibiesForm">
  <div class="page-header"> 
    <h1 id="formTitleSeibies">Add / Edit SEIBIES Forecast</h1>
    <!-- ✅ TOP SAVE BUTTON - same type="submit" + .btn-primary-seibies
         as the bottom button, so both trigger the same submit handler
         and stay in sync (visibility, saving state). The header is
         position:sticky, so this button stays visible while scrolling
         through the long forecast form. -->
    <button type="submit"
            form="seibiesForm"
            class="btn-primary-seibies btn-save-top-seibies"
            title="Save SEIBIES Forecast">💾 Save SEIBIES Forecast</button>
  </div>
  
      
      <!-- Project Selection (hidden in edit mode) -->
      <div class="form-row-seibies hidden-section">
        <div class="form-group-seibies project-select-container-seibies">
          <label>Project</label>
          <select id="projectSelectSeibies" required></select>
        </div>
      </div>

      <!-- MAIN WRAPPER - 2 Column Layout -->
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

        <!-- RIGHT SIDE - OBL Section -->
        <div class="project-right-side-seibies">
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
              ✅ All values are within SEI-BIES limit
            </div>
          </div>
        </div>

      </div>

      <!-- Forecast Sections -->
      <div class="forecast-section-seibies hidden-section" id="historicalForecastSectionSeibies">
        <div class="historical-toggle-header-seibies">
          <h3>📊 Historical Data (Last 3 Months - Readonly)</h3>
          <button type="button" id="toggleHistoricalBtnSeibies" class="toggle-arrow-btn-seibies" aria-expanded="false" title="Show historical data">
            <span class="toggle-arrow-seibies">▼</span>
            <span class="toggle-label-seibies">Show</span>
          </button>
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
  script: () => {
    setTimeout(async () => {
      const siteUrl = typeof _spPageContextInfo !== "undefined" 
        ? _spPageContextInfo.webAbsoluteUrl 
        : window.location.origin + "/sites/DigitalTransformationTest";
      const listName = "SEIBIES";
      const projectListName = "Project List";

      // DOM Elements
      const tableBody = document.querySelector(".seibies-table-body") as HTMLElement;
      const monthColumnsHeader = document.querySelector("#monthColumnsHeaderSeibies") as HTMLElement;
      const totalsHeader = document.querySelector("#totalsHeaderSeibies") as HTMLElement;
      const formSection = document.querySelector(".form-section-seibies") as HTMLElement;
      const tableSection = document.querySelector(".table-section-seibies") as HTMLElement;
      const showFormBtn = document.querySelector(".show-form-btn-seibies") as HTMLButtonElement;
      const projectSelect = document.querySelector("#projectSelectSeibies") as HTMLSelectElement;
      const historicalContainer = document.querySelector("#historicalContainerSeibies") as HTMLElement;
      const firstRowContainer = document.querySelector("#firstRowContainerSeibies") as HTMLElement;
      const remainingContainer = document.querySelector("#remainingContainerSeibies") as HTMLElement;
      const yearlyContainer = document.querySelector("#yearlyContainerSeibies") as HTMLElement;
      const backBtn = document.querySelector(".back-to-table-seibies") as HTMLButtonElement;
      const form = document.querySelector("#seibiesForm") as HTMLFormElement;
      const exportBtn = document.querySelector(".export-btn-seibies") as HTMLButtonElement;
      const oblLastMonthInput = document.querySelector("#oblLastMonthSeibies") as HTMLInputElement;
      const totalCurrentValueSpan = document.querySelector("#totalCurrentValueSeibies") as HTMLSpanElement;
      const remainingValueSpan = document.querySelector("#remainingValueSeibies") as HTMLSpanElement;
      const oblLineFill = document.querySelector("#oblLineFillSeibies") as HTMLDivElement;
      const oblValidationMessage = document.querySelector("#oblValidationMessageSeibies") as HTMLDivElement;
      const formTitle = document.querySelector("#formTitleSeibies") as HTMLElement;

      // Filter Elements
      const filterYear = document.querySelector("#filterYearSeibies") as HTMLSelectElement;
      const filterMonth = document.querySelector("#filterMonthSeibies") as HTMLSelectElement;
      const filterProject = document.querySelector("#filterProjectSeibies") as HTMLSelectElement;
      const filterCustomer = document.querySelector("#filterCustomerSeibies") as HTMLSelectElement;
      const filterProjectManager = document.querySelector("#filterProjectManagerSeibies") as HTMLSelectElement;
      const filterGroupHead = document.querySelector("#filterGroupHeadSeibies") as HTMLSelectElement;
      const filterProjectReference = document.querySelector("#filterProjectReferenceSeibies") as HTMLSelectElement;

      // Action Buttons
      const clearFiltersBtn = document.querySelector(".clear-filters-btn-seibies") as HTMLButtonElement;

      // Pagination Elements
      const firstPageBtn = document.querySelector("#firstPageSeibies") as HTMLButtonElement;
      const prevPageBtn = document.querySelector("#prevPageSeibies") as HTMLButtonElement;
      const nextPageBtn = document.querySelector("#nextPageSeibies") as HTMLButtonElement;
      const lastPageBtn = document.querySelector("#lastPageSeibies") as HTMLButtonElement;
      const currentPageSpan = document.querySelector("#currentPageSeibies") as HTMLSpanElement;
      const totalPagesSpan = document.querySelector("#totalPagesSeibies") as HTMLSpanElement;
      const pageNumbersDiv = document.querySelector("#pageNumbersSeibies") as HTMLDivElement;
      const startItemSpan = document.querySelector("#startItemSeibies") as HTMLSpanElement;
      const endItemSpan = document.querySelector("#endItemSeibies") as HTMLSpanElement;
      const totalItemsSpan = document.querySelector("#totalItemsSeibies") as HTMLSpanElement;

      const pageSizeSelect = document.querySelector("#pageSizeSelectSeibies") as HTMLSelectElement;

      let projects: ProjectItem[] = [];
      let forecasts: SeibiesForecastItem[] = [];
      let editingId: number | null = null;
      let allMonths: string[] = [];

      // Pagination variables
      let pageSize = 25;
      let currentPage = 1;
      let totalPages = 1;
      let filteredData: SeibiesForecastItem[] = [];

      // ========== UTILITY FUNCTIONS ==========

      const toast = (icon: any, title: string) => Swal.fire({ toast: true, icon, title, position: "top-end", timer: 2000, showConfirmButton: false });

      /**
       * ✅ SINGLE SOURCE OF TRUTH for every save button in the SEIBIES form.
       *    Both the top button (in the sticky project details panel) and the
       *    bottom button (in .form-actions-seibies) carry the .btn-primary-seibies
       *    class + type="submit", so this helper always returns BOTH.
       *    Any code that shows / hides / disables / re-labels a save button MUST
       *    use this helper so the two buttons never drift out of sync.
       */
      const getSaveButtons = (): HTMLButtonElement[] =>
        Array.from(document.querySelectorAll<HTMLButtonElement>('button.btn-primary-seibies[type="submit"]'));

      // ============================================
      // SECTION LOADER
      // ============================================

      const getRevenueLoaderMarkup = (): string => `
  <div id="revenuePageLoader" class="revenue-loader" style="position:absolute;top:0;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;background:rgba(248,249,250,0.92);z-index:999;">
    <div class="revenue-loader-box">
      <div class="revenue-loader-spinner"></div>
      <div class="revenue-loader-text">Loading Revenue Data</div>
      <div class="revenue-loader-subtext">Please wait... fetching SEIBIES projects and forecasts</div>
      <div class="revenue-loader-bar"><span></span></div>
    </div>
  </div>`;

      const showRevenueLoader = (text?: string, subtext?: string): void => {
        let loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;

        if (!loader) {
          const host = document.querySelector('.seibies-page') || document.body;
          host.insertAdjacentHTML('afterbegin', getRevenueLoaderMarkup());
          loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;
        }
        if (!loader) return;

        loader.classList.remove('revenue-loader-hidden');
        loader.style.display = 'flex';
        loader.style.opacity = '1';
        loader.style.visibility = 'visible';

        loader.parentElement?.classList.add('revenue-loading');
        positionLoaderOverlay();
        window.setTimeout(positionLoaderOverlay, 0);

        const textEl = loader.querySelector('.revenue-loader-text');
        if (textEl && text) textEl.textContent = text;

        const subEl = loader.querySelector('.revenue-loader-subtext');
        if (subEl && subtext) subEl.textContent = subtext;
      };

      const positionLoaderOverlay = (): void => {
        const loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;
        if (!loader || !loader.parentElement) return;

        const rect = loader.parentElement.getBoundingClientRect();
        const top = Math.max(0, rect.top);
        const height = Math.max(window.innerHeight - top, rect.height - top, 320);

        loader.style.setProperty('--loader-top', `${top}px`);
        loader.style.height = `${height}px`;
      };

      window.addEventListener('resize', positionLoaderOverlay);
      window.addEventListener('scroll', positionLoaderOverlay, { passive: true });

      const hideRevenueLoader = (): void => {
        const loader = document.querySelector('#revenuePageLoader') as HTMLElement | null;
        document.querySelector('.seibies-page')?.classList.remove('revenue-loading');
        if (!loader) return;

        loader.classList.add('revenue-loader-hidden');

        window.setTimeout(() => {
          if (loader.parentNode) loader.parentNode.removeChild(loader);
        }, 350);
      };

      const reloadSeibiesData = async (): Promise<void> => {
        showRevenueLoader('Reloading Data', 'Refreshing projects and forecasts...');
        try {
          await loadProjects();
          await loadForecasts();
          initializePagination(forecasts);
        } catch (error) {
          console.error('❌ Reload failed:', error);
          toast("error", "Reload failed - showing current data");
        } finally {
          hideRevenueLoader();
        }
      };

      (window as any).reloadSeibiesData = reloadSeibiesData;
      (window as any).hideRevenueLoader = hideRevenueLoader;

      const getDigest = async (): Promise<string> => {
        const res = await fetch(`${siteUrl}/_api/contextinfo`, {
          method: "POST",
          headers: { Accept: "application/json;odata=verbose" }
        });
        const data = await res.json();
        return data.d.GetContextWebInformation.FormDigestValue;
      };

      const safeJsonParse = (jsonString: string | undefined, fallback: any = {}): any => {
        try {
          if (!jsonString || jsonString.trim() === '') {
            return fallback;
          }

          const cleanedString = jsonString.trim();

          if (cleanedString === '""' || cleanedString === "null") {
            return fallback;
          }

          return JSON.parse(cleanedString);
        } catch (error) {
          console.error("JSON parse error:", error, "Input:", jsonString);
          return fallback;
        }
      };

      const formatMonthKey = (date: Date): string => {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
      };

      const formatMonthLabel = (date: Date): string => {
        return date.toLocaleString("default", { month: "short", year: "2-digit" });
      };

      const generateForecastTitle = (project: ProjectItem, forecastData: ForecastData): string => {
        const currentYear = new Date().getFullYear();
        let totalValue = 0;

        Object.keys(forecastData).forEach(key => {
          if (key.match(/^\d{4}-\d{2}$/)) {
            const monthData = forecastData[key] as ForecastMonthData;
            if (monthData) {
              totalValue += monthData.actual;
            }
          }
        });

        return `${project.Title} - ${currentYear} SEIBIES Forecast - ${totalValue.toLocaleString()}`;
      };

      const getDisplayMonths = (): string[] => {
        const now = new Date();
        const months: string[] = [];

        for (let i = 0; i < 3; i++) {
          const month = new Date(now.getFullYear(), now.getMonth() + i, 1);
          months.push(formatMonthKey(month));
        }

        return months;
      };

      const getAllMonthsForStorage = (): string[] => {
        return generateIdealMonthStructure();
      };

      const allowOnlyNumbers = (input: HTMLInputElement): void => {
        input.addEventListener('input', function() {
          let value = this.value.replace(/[^\d]/g, '');

          if (value.length > 1 && value.startsWith('0')) {
            value = value.replace(/^0+/, '');
            if (value === '') value = '0';
          }

          this.value = value;

          if (value === '' || value === '0') {
            this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
          } else {
            this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
          }
        });

        input.addEventListener('keydown', function(e) {
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

        input.addEventListener('paste', function(e) {
          const clipboardData = e.clipboardData || (window as any).clipboardData;
          const pastedData = clipboardData?.getData('Text') || '';

          if (/[^\d]/.test(pastedData)) {
            e.preventDefault();
          }
        });
      };

      const getNumericValue = (input: HTMLInputElement): number => {
        const value = input.value.trim();
        if (value === '' || value === '0') return 0;
        return parseInt(value) || 0;
      };

      const setNumericValue = (input: HTMLInputElement, value: number): void => {
        if (value === 0) {
          input.value = '';
          input.className = input.className.replace('input-with-value-seibies', 'input-empty-seibies');
        } else {
          input.value = value.toString();
          input.className = input.className.replace('input-empty-seibies', 'input-with-value-seibies');
        }
      };

      const showOBLErrorToast = (message: string, isWarning: boolean = false): void => {
        const existingToast = document.querySelector('.obl-error-toast-seibies');
        if (existingToast) existingToast.remove();

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

        toast.querySelector('.toast-close-seibies')?.addEventListener('click', () => {
          toast.remove();
        });
      };

      // ========== MONTH STRUCTURE FUNCTIONS ==========

      const generateIdealMonthStructure = (): string[] => {
        const months: string[] = [];
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

      const getLastMonthKey = (existingData: ForecastData): string | undefined => {
        const monthKeys = Object.keys(existingData).filter(key => key.match(/^\d{4}-\d{2}$/));

        if (monthKeys.length === 0) {
          return undefined;
        }

        const sortedMonths = monthKeys.sort();
        const lastMonth = sortedMonths[sortedMonths.length - 1];
        return lastMonth;
      };

      const createFreshYearlyStructure = (lastMonthKey?: string): YearlyForecastData => {
        const yearlyData: YearlyForecastData = {};

        let startYear: number;

        if (lastMonthKey) {
          startYear = parseInt(lastMonthKey.split('-')[0]);
        } else {
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

      const createFreshForecastStructure = (): ForecastData => {
        const idealMonths = generateIdealMonthStructure();
        const freshData: ForecastData = {};

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

      const isForecastMonthData = (data: any): data is ForecastMonthData => {
        return data &&
               typeof data.plan === 'number' &&
               typeof data.actual === 'number';
      };

      const handleMonthProgression = (forecastData: ForecastData): ForecastData => {
        const now = new Date();
        const currentMonthKey = formatMonthKey(now);
        const originalMonths = Object.keys(forecastData).filter(key => key.match(/^\d{4}-\d{2}$/)).sort();
        const idealMonths = generateIdealMonthStructure();

        if (originalMonths[0] === idealMonths[0]) {
          return forecastData;
        }

        const updatedData: ForecastData = {};

        idealMonths.forEach((monthKey) => {
          const existingData = forecastData[monthKey];
          if (existingData && isForecastMonthData(existingData)) {
            updatedData[monthKey] = JSON.parse(JSON.stringify(existingData));
          } else {
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
        } else {
          const lastMonthKey = idealMonths[idealMonths.length - 1];
          updatedData.yearly = createFreshYearlyStructure(lastMonthKey);
        }

        return updatedData;
      };

      const propagateMonthValues = (forecastData: ForecastData): ForecastData => {
        if (Object.keys(forecastData).length === 0) {
          return forecastData;
        }

        return handleMonthProgression(forecastData);
      };

      // ========== DATA LOADING FUNCTIONS ==========

const loadProjects = async (): Promise<void> => {
  try {
    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items?$select=Id,Title,CustomerName,GroupHead1/Title,GroupHead1Id,PC,ProjectReference,ProjectManager/Title,ProjectManagerId&$expand=GroupHead1,ProjectManager&$top=5000`, {
      headers: { Accept: "application/json;odata=verbose" }
    });
    
    if (!res.ok) {
      const fallbackRes = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${projectListName}')/items?$select=Id,Title,CustomerName,GroupHead1Id,PC,ProjectReference,ProjectManagerId&$top=5000`, {
        headers: { Accept: "application/json;odata=verbose" }
      });
      
      if (!fallbackRes.ok) {
        throw new Error(`HTTP error! status: ${fallbackRes.status}`);
      }
      
      const fallbackData = await fallbackRes.json();
      const rawProjects: RawProjectItem[] = fallbackData.d?.results || [];
      
      const cleanedProjects = await Promise.all(rawProjects.map(async (p: RawProjectItem) => {
        let projectManager = "";
        let groupHead1 = "";
        
        if (p.ProjectManagerId) {
          try {
            const pmRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${p.ProjectManagerId})?$select=Id,Title,Email`, {
              headers: { Accept: "application/json;odata=verbose" }
            });
            if (pmRes.ok) {
              const pmData = await pmRes.json();
              projectManager = pmData.d.Title || "";
            }
          } catch (e) {
            console.warn('Could not fetch ProjectManager:', e);
          }
        }
        
        if (p.GroupHead1Id) {
          try {
            const ghRes = await fetch(`${siteUrl}/_api/web/getuserbyid(${p.GroupHead1Id})?$select=Id,Title,Email`, {
              headers: { Accept: "application/json;odata=verbose" }
            });
            if (ghRes.ok) {
              const ghData = await ghRes.json();
              groupHead1 = ghData.d.Title || "";
            }
          } catch (e) {
            console.warn('Could not fetch GroupHead1:', e);
          }
        }
        
        return {
          ...p,
          CustomerName: p.CustomerName || "",
          GroupHead1: groupHead1,
          GroupHead1Id: p.GroupHead1Id,
          PC: p.PC || "",
          ProjectReference: p.ProjectReference || "",
          ProjectManager: projectManager,
          ProjectManagerId: p.ProjectManagerId
        };
      }));
      
      const projectMap = new Map<string, ProjectItem>();
      cleanedProjects.forEach((p: ProjectItem) => {
        projectMap.set(p.Title, p);
      });
      
      projects = Array.from(projectMap.values());
      console.log(`📋 Total projects from SharePoint (fallback): ${projects.length}`);
      
      const assignedProjects = getAssignedProjects(projects);
      
      projectSelect.innerHTML = `<option value="">Select Project</option>` + 
        assignedProjects.map((p: ProjectItem) => `<option value="${p.Id}">${p.Title}</option>`).join("");
      
      filterProject.innerHTML = '<option value="">All Projects</option>' + 
        [...new Set(assignedProjects.map((p: ProjectItem) => p.Title))]
          .sort()
          .map(project => `<option value="${project}">${project}</option>`)
          .join("");
      
      return;
    }
    
    const data = await res.json();
    const rawProjects: RawProjectItem[] = data.d?.results || [];

    const cleanedProjects = rawProjects.map((p: RawProjectItem) => ({
      ...p,
      CustomerName: p.CustomerName || "",
      GroupHead1: p.GroupHead1 || "",
      GroupHead1Id: p.GroupHead1Id,
      PC: p.PC || "",
      ProjectReference: p.ProjectReference || "",
      ProjectManager: p.ProjectManager || "",
      ProjectManagerId: p.ProjectManagerId
    }));

    const projectMap = new Map<string, ProjectItem>();
    cleanedProjects.forEach((p: ProjectItem) => {
      projectMap.set(p.Title, p);
    });

    projects = Array.from(projectMap.values());

    if (projects.length > 0) {
      console.log('📋 Sample project data:', JSON.stringify({
        Title: projects[0].Title,
        ProjectManager: projects[0].ProjectManager,
        ProjectManagerId: projects[0].ProjectManagerId,
        GroupHead1: projects[0].GroupHead1,
        GroupHead1Id: projects[0].GroupHead1Id
      }, null, 2));
    }

    const assignedProjects = getAssignedProjects(projects);
    console.log(`📋 Total projects: ${projects.length}, Assigned projects: ${assignedProjects.length}`);

    projectSelect.innerHTML = `<option value="">Select Project</option>` + 
      assignedProjects.map((p: ProjectItem) => `<option value="${p.Id}">${p.Title}</option>`).join("");

    filterProject.innerHTML = '<option value="">All Projects</option>' + 
      [...new Set(assignedProjects.map((p: ProjectItem) => p.Title))]
        .sort()
        .map(project => `<option value="${project}">${project}</option>`)
        .join("");

  } catch (error) {
    console.error("Error loading projects:", error);
    toast("error", "Failed to load projects");
  }
};

const loadForecasts = async (): Promise<void> => {
  try {
    const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Project/Id,Project/Title,ProfitCentre,OBL,ForecastData,SeibiesActual,CreatedDate,UpdatedDate&$expand=Project&$top=5000`, {
      headers: { Accept: "application/json;odata=verbose" }
    });
    const data: SharePointResponse<SeibiesForecastItem> = await res.json();

    const assignedProjects = getAssignedProjects(projects);
    const assignedProjectTitles = new Set(assignedProjects.map(p => p.Title));

    const allForecasts = (data.d?.results || []).map(forecast => ({
      ...forecast,
      ForecastData: typeof forecast.ForecastData === 'string' ?
        forecast.ForecastData :
        JSON.stringify(forecast.ForecastData || {}),
      SeibiesActual: typeof forecast.SeibiesActual === 'string' ?
        forecast.SeibiesActual :
        JSON.stringify(forecast.SeibiesActual || {})
    }));

    forecasts = allForecasts.filter(forecast => {
      const projectTitle = forecast.Project?.Title;
      return projectTitle && assignedProjectTitles.has(projectTitle);
    });

    console.log(`✅ Loaded ${forecasts.length} SEIBIES forecasts`);
    console.log(`👑 Is GH Head: ${isGHHead()}`);

  } catch (error) {
    console.error("Error loading SEIBIES forecasts:", error);
    toast("error", "Failed to load SEIBIES forecasts");
  }
};

const getGroupHeadDisplayName = (groupHead: string | { Title: string; Email?: string } | undefined): string => {
  if (!groupHead) return "-";
  if (typeof groupHead === "object" && groupHead !== null) {
    return groupHead.Title || "-";
  }
  return groupHead;
};

      // ========== EXPORT FUNCTIONS ==========

      const getMonthName = (monthNumber: string): string => {
        const months = [
          "Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ];
        return months[parseInt(monthNumber) - 1] || "Month";
      };

      const generateHeadersStructure = (selectedYear: string, selectedMonth: string, exportData: SeibiesForecastItem[]) => {
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

        const allMonthKeys = new Set<string>();

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

        const monthlyHeaders: string[] = [];
        next17Months.forEach(monthKey => {
          const [year, month] = monthKey.split('-');
          const monthLabel = getMonthName(month);
          const displayLabel = `${monthLabel} ${year}`;
          monthlyHeaders.push(`${displayLabel} (P)`, `${displayLabel} (C)`);
        });

        const lastFiveYears = getLastFiveYears(exportData);
        const yearlyHeaders: string[] = [];
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

      const getLastFiveYears = (data: SeibiesForecastItem[]): string[] => {
        const years = new Set<string>();

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

      const exportToExcelWithFullStyle = (): void => {
        if (!canRead()) {
          toast("error", "❌ You don't have permission to export data");
          return;
        }

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

        showRevenueLoader('Exporting', 'Preparing the Excel workbook...');

        try {
          const { headers, allMonths, lastFiveYears } = generateHeadersStructure(selectedYear, selectedMonth, exportData);

          const excelData = exportData.map(forecast => {
            const projectTitle = forecast.Project?.Title || "-";
            const proj = projects.find(p => p.Title === projectTitle);

            const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
            const oblValue = forecast.OBL || 0;

            const selectedMonthKey = `${selectedYear}-${selectedMonth.padStart(2, '0')}`;
            const monthName = getMonthName(selectedMonth);

            const selectedMonthForecastData = forecastData[selectedMonthKey] as ForecastMonthData || { plan: 0, actual: 0 };
            const selectedMonthPlan = selectedMonthForecastData.plan || 0;
            const selectedMonthCurrent = selectedMonthForecastData.actual || 0;

            let groupHeadDisplay = "-";
            if (proj?.GroupHead1) {
              groupHeadDisplay = typeof proj.GroupHead1 === "object" && proj.GroupHead1 !== null
                ? proj.GroupHead1.Title || ""
                : proj.GroupHead1 || "";
            }
            if (!groupHeadDisplay) groupHeadDisplay = "-";

            let projectManagerDisplay = "-";
            if (proj?.ProjectManager) {
              projectManagerDisplay = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
                ? proj.ProjectManager.Title || ""
                : proj.ProjectManager || "";
            }
            if (!projectManagerDisplay) projectManagerDisplay = "-";

            const rowData: any = {
              "Profit Centre": forecast.ProfitCentre || "-",
              "SO/Project Ref": proj?.ProjectReference || "-",
              "Customer Name": proj?.CustomerName || "-",
              "Project Name": projectTitle,
              "Group Head": groupHeadDisplay,
              "Project Manager": projectManagerDisplay,
              "SEI-BIES": oblValue,
              [`${monthName} ${selectedYear} (P)`]: selectedMonthPlan,
              [`${monthName} ${selectedYear} (C)`]: selectedMonthCurrent,
            };

            allMonths.forEach(monthKey => {
              const [year, month] = monthKey.split('-');
              const monthLabel = getMonthName(month);
              const displayLabel = `${monthLabel} ${year}`;

              const forecastMonthData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
              const planValue = forecastMonthData.plan || 0;
              const currentValue = forecastMonthData.actual || 0;

              rowData[`${displayLabel} (P)`] = planValue;
              rowData[`${displayLabel} (C)`] = currentValue;
            });

            lastFiveYears.forEach(year => {
              const yearData = forecastData.yearly?.[year] || { plan: 0, actual: 0 };
              rowData[`${year} (P)`] = yearData.plan;
              rowData[`${year} (C)`] = yearData.actual;
            });

            const totalMonthlyCurrent = Object.keys(forecastData)
              .filter(key => key.match(/^\d{4}-\d{2}$/))
              .reduce((sum, monthKey) => {
                const monthData = forecastData[monthKey] as ForecastMonthData;
                return sum + (monthData?.actual || 0);
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
            fill: { fgColor: { rgb: "FF0078D4" } },
            font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "FF000000" } },
              left: { style: "thin", color: { rgb: "FF000000" } },
              bottom: { style: "thin", color: { rgb: "FF000000" } },
              right: { style: "thin", color: { rgb: "FF000000" } }
            }
          };

          const dataStyle = {
            font: { name: "Arial", sz: 10 },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "FFD0D0D0" } },
              left: { style: "thin", color: { rgb: "FFD0D0D0" } },
              bottom: { style: "thin", color: { rgb: "FFD0D0D0" } },
              right: { style: "thin", color: { rgb: "FFD0D0D0" } }
            }
          };

          const numberStyle = {
            ...dataStyle,
            alignment: { horizontal: "right", vertical: "center" }
          };

          const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            if (!ws[cellAddress]) continue;

            if (!ws[cellAddress].s) {
              ws[cellAddress].s = {};
            }
            ws[cellAddress].s = headerStyle;
          }

          for (let row = range.s.r + 1; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
              const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
              if (!ws[cellAddress]) continue;

              const isNumberColumn = col >= 6;
              const styleToApply = isNumberColumn ? numberStyle : dataStyle;

              if (!ws[cellAddress].s) {
                ws[cellAddress].s = {};
              }
              ws[cellAddress].s = styleToApply;
            }
          }

          const colWidths = headers.map((header: string) => {
            if (header.includes('Project Name')) return { wch: 25 };
            if (header.includes('Customer Name')) return { wch: 20 };
            if (header.includes('Profit Centre') || header.includes('SO/Project Ref')) return { wch: 12 };
            if (header.includes('Group Head') || header.includes('Project Manager')) return { wch: 15 };
            if (header.includes('SEI-BIES') || header.includes('Balance')) return { wch: 12 };
            if (header.includes('(P)') || header.includes('(C)')) return { wch: 10 };
            return { wch: 12 };
          });
          ws['!cols'] = colWidths;

          ws['!freeze'] = { x: 0, y: 1 };

          XLSX.utils.book_append_sheet(wb, ws, "SEIBIES Forecast");

          const monthName = getMonthName(selectedMonth);
          const fileName = `SEIBIES_Forecast_${monthName}_${selectedYear}_Filtered.xlsx`;

          XLSX.writeFile(wb, fileName);
          hideRevenueLoader();

          toast("success", `Exported ${exportData.length} filtered SEIBIES records for ${monthName} ${selectedYear}`);

        } catch (error) {
          hideRevenueLoader();
          console.error("Export error:", error);
          toast("error", "Failed to export SEIBIES data to Excel");
        }
      };

      // ========== PAGINATION FUNCTIONS ==========

      const initializePagination = (data: SeibiesForecastItem[]): void => {
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

      const calculateTotalPages = (): void => {
        totalPages = Math.ceil(filteredData.length / pageSize);
        if (totalPages === 0) totalPages = 1;
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

        toast("info", `Showing ${pageSize} SEIBIES records per page`);
      };

      const updatePaginationControls = (): void => {
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

      const updatePageNumbers = (): void => {
        pageNumbersDiv.innerHTML = '';

        if (filteredData.length === 0) return;

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

      const goToPage = (page: number): void => {
        if (page < 1 || page > totalPages || page === currentPage) return;

        currentPage = page;
        updatePaginationControls();
        renderCurrentPage();
      };

      const setupPaginationEventListeners = (): void => {
        if (firstPageBtn) firstPageBtn.addEventListener("click", () => goToPage(1));
        if (prevPageBtn) prevPageBtn.addEventListener("click", () => goToPage(currentPage - 1));
        if (nextPageBtn) nextPageBtn.addEventListener("click", () => goToPage(currentPage + 1));
        if (lastPageBtn) lastPageBtn.addEventListener("click", () => goToPage(totalPages));

        if (pageSizeSelect) {
          pageSizeSelect.addEventListener("change", (e) => {
            const newSize = parseInt((e.target as HTMLSelectElement).value);
            changePageSize(newSize);
          });
        }
      };

      // ========== TABLE RENDERING FUNCTIONS ==========

      const calculateTotals = (
        data: SeibiesForecastItem[],
        displayMonths: string[],
        currentMonthKey: string
      ): any => {
        const totals = {
          obl: 0,
          monthly: {} as {
            [key: string]: { plan: number; current: number }
          }
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

      const applyStickyHeaderOffsets = (): void => {
        const table = document.querySelector('.month-columns-table-seibies') as HTMLTableElement | null;
        if (!table) return;

        const headerRows = table.querySelectorAll<HTMLTableRowElement>('thead tr');
        let offset = 0;

        headerRows.forEach(row => {
          row.querySelectorAll<HTMLElement>('th, td').forEach(cell => {
            cell.style.position = 'sticky';
            cell.style.top = `${offset}px`;
          });
          offset += row.getBoundingClientRect().height;
        });
      };

      const renderTotalsHeader = (displayMonths: string[], currentMonthKey: string): void => {
        if (!totalsHeader) return;

        totalsHeader.innerHTML = '';

        const allData = forecasts;
        const filtered = filteredData.length ? filteredData : allData;

        const allTotals = calculateTotals(allData, displayMonths, currentMonthKey);
        const filteredTotals = calculateTotals(filtered, displayMonths, currentMonthKey);

        const baseColumns = 6;
        const monthColumns = displayMonths.length * 2;
        const totalColspan = baseColumns + monthColumns;

        const buildTotalsRow = (label: string, totals: any, css: string) => {
          const monthCells = displayMonths.map(m => {
            const t = totals.monthly[m] || { plan: 0, current: 0 };
            return `
              <td class="month-scroll-col-seibies" style="min-width:60px; text-align:center; vertical-align:middle;">${t.plan.toLocaleString('en-IN')}</td>
              <td class="month-scroll-col-seibies" style="min-width:60px; text-align:center; vertical-align:middle;">${t.current.toLocaleString('en-IN')}</td>
            `;
          }).join('');

          return `
            <tr class="${css}">
              <td class="project-ref-col-seibies" style="text-align:center; vertical-align:middle; min-width:120px; max-width:120px; width:120px;">${css === 'all-projects-total-seibies' ? '★' : '●'}</td>
              <td class="customer-col-seibies" style="text-align:center; vertical-align:middle; min-width:130px; max-width:130px; width:130px;"></td>
              <td class="project-col-seibies" style="text-align:left; vertical-align:middle; min-width:180px; max-width:180px; width:180px; font-weight:bold; background-color:${css === 'all-projects-total-seibies' ? '#d1ecf1' : '#d4edda'};">
                ${label}
              </td>
              <td class="gh-col-seibies" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
              <td class="pm-col-seibies" style="text-align:center; vertical-align:middle; min-width:100px; max-width:100px; width:100px;"></td>
              <td class="obl-col-seibies" style="text-align:right; vertical-align:middle; min-width:90px; max-width:90px; width:90px; font-weight:bold;">${totals.obl.toLocaleString('en-IN')}</td>
              ${monthCells}
            </tr>
          `;
        };

        let html = `
          <tr class="totals-title-row-seibies">
            <th colspan="${totalColspan}" style="background:#f8f9fa; padding:8px; text-align:center; font-size:13px; font-weight:bold; color:#495057; border:1px solid #dee2e6;">
              📊 SEI-BIES Totals
            </th>
          </tr>
        `;

        if (allData.length > 0) {
          html += buildTotalsRow(`All Projects (${allData.length})`, allTotals, 'all-projects-total-seibies');
        }

        if (filtered.length > 0) {
          html += buildTotalsRow(`Filtered (${filtered.length})`, filteredTotals, 'filtered-total-seibies');
        }

        totalsHeader.innerHTML = html;
        applyStickyHeaderOffsets();
      };

      const renderCurrentPage = (): void => {
        if (filteredData.length === 0) {
          tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No SEIBIES records found</td></tr>`;
          if (totalsHeader) totalsHeader.innerHTML = '';
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

      const attachEditTriggerListeners = (): void => {
        document.querySelectorAll('.edit-trigger-seibies').forEach(link => {
          link.addEventListener('click', (e) => {
            e.preventDefault();
            const id = parseInt((e.currentTarget as HTMLElement).dataset.id!);
            editForecast(id);
          });
        });
      };

      const renderTableWithMonthColumns = (data: SeibiesForecastItem[]): void => {
        if (!data.length) {
          tableBody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px;">No SEIBIES records found</td></tr>`;
          if (totalsHeader) totalsHeader.innerHTML = '';
          return;
        }

        const displayMonths = getDisplayMonths();
        const currentMonthKey = formatMonthKey(new Date());

        renderTotalsHeader(displayMonths, currentMonthKey);

        const mainHeaderHTML = `
          <tr>
            <th rowspan="2" class="project-ref-col-seibies" style="min-width:120px; max-width:120px; width:120px;">
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
            
            <th rowspan="2" class="customer-col-seibies" style="min-width:130px; max-width:130px; width:130px;">
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
            
            <th rowspan="2" class="project-col-seibies" style="min-width:180px; max-width:180px; width:180px;">
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
            
            <th rowspan="2" class="gh-col-seibies" style="min-width:100px; max-width:100px; width:100px;">
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
            
            <th rowspan="2" class="pm-col-seibies" style="min-width:100px; max-width:100px; width:100px;">
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
            
            <th rowspan="2" class="obl-col-seibies" style="min-width:90px; max-width:90px; width:90px;">
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
              const headerClass = isCurrentMonth ? 'current-month-header-seibies' : '';
              return `
                <th colspan="2" class="month-scroll-col-seibies month-column-header-seibies ${headerClass}" style="min-width:60px;">
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
              const subHeaderClass = isCurrentMonth ? 'current-month-header-table-seibies' : '';
              return `
                <th class="month-scroll-col-seibies month-subheader-seibies ${subHeaderClass}" style="min-width:60px;">
                  <div class="sort-header-container-seibies">
                    <span>Previous</span>
                    <div class="sort-buttons-container-seibies">
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'asc' ? 'sort-asc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="asc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Ascending">▲</button>
                      <button class="sort-btn-seibies ${currentSortConfig.column === 'month' && currentSortConfig.monthIndex === monthIndex && currentSortConfig.subColumn === 'plan' && currentSortConfig.direction === 'desc' ? 'sort-desc-seibies active-seibies' : ''}" 
                              data-sort="month" data-direction="desc" data-subcolumn="plan" data-month-index="${monthIndex}" title="Sort Plan Descending">▼</button>
                    </div>
                  </div>
                </th>
                <th class="month-scroll-col-seibies month-subheader-seibies ${subHeaderClass}" style="min-width:60px;">
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
          const projectTitle = forecast.Project?.Title || "-";
          const proj = projects.find(p => p.Title === projectTitle);

          const forecastData: ForecastData = safeJsonParse(forecast.ForecastData, {});
          const formattedOBL = Math.round(forecast.OBL || 0).toLocaleString('en-IN');

          const projectRefText = proj?.ProjectReference || `#${forecast.Id}`;
          const projectRefLink = `<a href="#" class="edit-trigger-seibies" data-id="${forecast.Id}" title="Open / edit this forecast">${projectRefText}</a>`;

          const projectName = projectTitle;

          let projectManagerDisplay = "-";
          if (proj?.ProjectManager) {
            projectManagerDisplay = typeof proj.ProjectManager === "object" && proj.ProjectManager !== null
              ? proj.ProjectManager.Title || "-"
              : proj.ProjectManager || "-";
          }

          let rowHTML = `
            <tr data-forecast-id="${forecast.Id}">
              <td class="project-ref-col-seibies" style="min-width:120px; max-width:120px; width:120px;">${projectRefLink}</td>
              
              <td class="customer-col-seibies" style="min-width:130px; max-width:130px; width:130px;">${proj?.CustomerName || "-"}</td>
              
              <td class="project-col-seibies" style="min-width:180px; max-width:180px; width:180px;">${projectName}</td>
              
              <td class="gh-col-seibies" style="min-width:100px; max-width:100px; width:100px;">${getGroupHeadDisplayName(proj?.GroupHead1)}</td>
              
              <td class="pm-col-seibies" style="min-width:100px; max-width:100px; width:100px;">${projectManagerDisplay}</td>
              
              <td class="obl-col-seibies" style="min-width:90px; max-width:90px; width:90px; text-align:right;">${formattedOBL}</td>
          `;

          displayMonths.forEach((monthKey) => {
            const monthData = forecastData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };
            const planValue = monthData.plan || 0;
            const currentValue = monthData.actual || 0;

            rowHTML += `
              <td class="month-scroll-col-seibies" style="min-width:60px;">${planValue.toLocaleString('en-IN')}</td>
              <td class="month-scroll-col-seibies" style="min-width:60px;">${currentValue.toLocaleString('en-IN')}</td>
            `;
          });

          rowHTML += `</tr>`;
          tableBodyHTML += rowHTML;
        });

        tableBody.innerHTML = tableBodyHTML;

        attachSortButtonListeners();
        attachEditTriggerListeners();
        applyStickyHeaderOffsets();
      };

      const attachSortButtonListeners = (): void => {
        document.querySelectorAll('.sort-btn-seibies[data-sort]').forEach(btn => {
          btn.addEventListener('click', (e) => {
            const button = e.currentTarget as HTMLButtonElement;
            const sortColumn = button.dataset.sort;
            const direction = button.dataset.direction as 'asc' | 'desc';
            const subColumn = button.dataset.subcolumn;
            const monthIndex = button.dataset.monthIndex ? parseInt(button.dataset.monthIndex) : undefined;

            if (sortColumn && direction) {
              handleSeibiesSort(sortColumn, direction, subColumn, monthIndex);
            }
          });
        });
      };

      const handleSeibiesSort = (column: string, direction: 'asc' | 'desc', subColumn?: string, monthIndex?: number): void => {
        currentSortConfig = {
          column,
          direction,
          subColumn,
          monthIndex
        };

        const dataToSort = filteredData.length > 0 ? filteredData : forecasts;

        if (dataToSort.length === 0) return;

        const sortedData = [...dataToSort].sort((a, b) => {
          let valueA: any = 0;
          let valueB: any = 0;

          switch (column) {
            case 'projectRef':
              const projA = projects.find(p => p.Title === a.Project?.Title);
              const projB = projects.find(p => p.Title === b.Project?.Title);
              valueA = projA?.ProjectReference || "";
              valueB = projB?.ProjectReference || "";
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
              const monthDataA = dataA[monthKeyToUse] as ForecastMonthData || { plan: 0, actual: 0 };
              const monthDataB = dataB[monthKeyToUse] as ForecastMonthData || { plan: 0, actual: 0 };

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
              } else {
                valueA = monthDataA.actual || 0;
                valueB = monthDataB.actual || 0;
              }
              break;

            case 'project':
              valueA = a.Project?.Title || "";
              valueB = b.Project?.Title || "";
              break;

            case 'customer':
              const projACust = projects.find(p => p.Title === a.Project?.Title);
              const projBCust = projects.find(p => p.Title === b.Project?.Title);
              valueA = projACust?.CustomerName || "";
              valueB = projBCust?.CustomerName || "";
              break;

            case 'projectManager':
              const projMA = projects.find(p => p.Title === a.Project?.Title);
              const projMB = projects.find(p => p.Title === b.Project?.Title);
              if (projMA?.ProjectManager) {
                valueA = typeof projMA.ProjectManager === "object" && projMA.ProjectManager !== null
                  ? projMA.ProjectManager.Title || ""
                  : projMA.ProjectManager || "";
              } else {
                valueA = "";
              }
              if (projMB?.ProjectManager) {
                valueB = typeof projMB.ProjectManager === "object" && projMB.ProjectManager !== null
                  ? projMB.ProjectManager.Title || ""
                  : projMB.ProjectManager || "";
              } else {
                valueB = "";
              }
              break;

            case 'groupHead':
              const projGA = projects.find(p => p.Title === a.Project?.Title);
              const projGB = projects.find(p => p.Title === b.Project?.Title);
              if (projGA?.GroupHead1) {
                valueA = typeof projGA.GroupHead1 === "object" && projGA.GroupHead1 !== null
                  ? projGA.GroupHead1.Title || ""
                  : projGA.GroupHead1 || "";
              } else {
                valueA = "";
              }
              if (projGB?.GroupHead1) {
                valueB = typeof projGB.GroupHead1 === "object" && projGB.GroupHead1 !== null
                  ? projGB.GroupHead1.Title || ""
                  : projGB.GroupHead1 || "";
              } else {
                valueB = "";
              }
              break;

            default:
              return 0;
          }

          if (typeof valueA === 'number' && typeof valueB === 'number') {
            return direction === 'asc' ? valueA - valueB : valueB - valueA;
          } else {
            const strA = String(valueA).toLowerCase();
            const strB = String(valueB).toLowerCase();
            return direction === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
          }
        });

        if (filteredData.length > 0) {
          filteredData = sortedData;
        } else {
          forecasts = sortedData;
        }

        currentPage = 1;
        updatePaginationControls();
        renderCurrentPage();

        const sortDirectionText = direction === 'asc' ? 'Ascending' : 'Descending';
        let columnName = '';

        if (subColumn) {
          columnName = subColumn.replace('-', ' ').replace(/\b\w/g, l => l.toUpperCase());
        } else if (column === 'projectRef') {
          columnName = 'Project Reference';
        } else if (column === 'obl') {
          columnName = 'SEI-BIES';
        } else if (column === 'project') {
          columnName = 'Project';
        } else if (column === 'customer') {
          columnName = 'Customer';
        } else if (column === 'projectManager') {
          columnName = 'Project Manager';
        } else if (column === 'groupHead') {
          columnName = 'Group Head';
        } else {
          const displayMonths = getDisplayMonths();
          const monthIndexToUse = monthIndex !== undefined ? monthIndex : 0;
          const monthKey = displayMonths[monthIndexToUse];
          const monthDate = new Date(monthKey + "-01");
          columnName = formatMonthLabel(monthDate);
        }

        toast("info", `Sorted by ${columnName} (${sortDirectionText})`);
      };

      const clearFilters = (): void => {
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

      const renderProjectDetails = (project: ProjectItem): void => {
        const projectNameDisplay = document.querySelector('#projectNameDisplaySeibies');
        if (projectNameDisplay) projectNameDisplay.textContent = project.Title || "-";
        
        const projectRef = document.querySelector('#projectRefSeibies');
        if (projectRef) projectRef.textContent = project.ProjectReference || "-";
        
        const profitCenter = document.querySelector('#profitCenterSeibies');
        if (profitCenter) profitCenter.textContent = project.PC || "-";
        
        const projectManager = document.querySelector('#projectManagerSeibies');
        if (projectManager) {
          projectManager.textContent = typeof project.ProjectManager === "object" ? 
            project.ProjectManager?.Title : project.ProjectManager || "-";
        }
        
        const groupHead = document.querySelector('#groupHeadSeibies');
        if (groupHead) {
          groupHead.textContent = getGroupHeadDisplayName(project.GroupHead1);
        }
        
        const customer = document.querySelector('#customerSeibies');
        if (customer) customer.textContent = project.CustomerName || "-";
      };

      // ========== HISTORICAL TOGGLE ==========

      const setupHistoricalToggle = (): void => {
        const toggleBtn = document.querySelector("#toggleHistoricalBtnSeibies") as HTMLButtonElement | null;
        const historicalContainer = document.querySelector("#historicalContainerSeibies") as HTMLElement | null;
        const header = document.querySelector(".historical-toggle-header-seibies") as HTMLElement | null;

        if (!toggleBtn || !historicalContainer || !header) return;

        let isVisible = false;

        const updateHistoricalToggle = (): void => {
          const arrow = toggleBtn.querySelector(".toggle-arrow-seibies") as HTMLElement | null;
          const label = toggleBtn.querySelector(".toggle-label-seibies") as HTMLElement | null;

          if (isVisible) {
            historicalContainer.style.display = "grid";
            toggleBtn.setAttribute("aria-expanded", "true");
            toggleBtn.title = "Hide historical data";
            if (arrow) arrow.textContent = "▲";
            if (label) label.textContent = "Hide";
          } else {
            historicalContainer.style.display = "none";
            toggleBtn.setAttribute("aria-expanded", "false");
            toggleBtn.title = "Show historical data";
            if (arrow) arrow.textContent = "▼";
            if (label) label.textContent = "Show";
          }
        };

        const toggleHistorical = (event?: Event): void => {
          if (event) {
            event.preventDefault();
            event.stopPropagation();
          }

          isVisible = !isVisible;
          updateHistoricalToggle();
        };

        toggleBtn.onclick = (event: MouseEvent): void => {
          toggleHistorical(event);
        };

        header.onclick = (event: MouseEvent): void => {
          const target = event.target as HTMLElement;

          if (target.closest("#toggleHistoricalBtnSeibies")) {
            return;
          }

          toggleHistorical(event);
        };

        updateHistoricalToggle();
      };

      // ========== FORM MANAGEMENT FUNCTIONS ==========

      const canViewForecastForm = (isEdit: boolean = false, project?: ProjectItem): boolean => {
        if (!canRead()) return false;

        if (isEdit) {
          if (canEditProject(project)) return true;
          if (canReadAllProjects()) return true;
          if (project && getAssignedProjects(projects).some(p => p.Id === project.Id)) return true;
          return false;
        }

        return canCreateRecord();
      };

      const showForm = (isEditMode: boolean = false, project?: ProjectItem): void => {
        console.log(`📝 showForm called: isEditMode=${isEditMode}, project=${project?.Title}`);

        if (!canViewForecastForm(isEditMode, project)) {
          if (!canRead()) {
            toast("error", "❌ You don't have permission to access this form");
            return;
          }
          if (isEditMode) {
            toast("error", "❌ You don't have permission to view this forecast");
            return;
          }
          toast("error", "❌ You don't have permission to create SEIBIES forecasts");
          return;
        }

        formSection.style.display = "block";
        tableSection.style.display = "none";
        if (showFormBtn) showFormBtn.style.display = "none";

        // ✅ Show/hide BOTH save buttons together (top + bottom)
        const canSave = isEditMode ? canEditProject(project) : canCreateRecord();
        const saveBtns = getSaveButtons();
        saveBtns.forEach(btn => {
          btn.style.display = canSave ? '' : 'none';
        });
        console.log(`💾 Save buttons ${canSave ? 'visible' : 'hidden'} (${saveBtns.length} button(s))`);

        if (isEditMode) {
          formSection.classList.add('edit-mode-seibies');
          projectSelect.disabled = true;

          if (formTitle) {
            const projectName = project?.Title || "";
            const projectReference = project?.ProjectReference || "";

            formTitle.textContent = projectName
              ? `${projectName}${projectReference ? ` (${projectReference})` : ""}`
              : "Edit SEIBIES Forecast";
            formTitle.style.color = '';
          }
        } else {
          formSection.classList.remove('edit-mode-seibies');
          projectSelect.disabled = false;
          if (formTitle) {
            formTitle.textContent = "Add SEIBIES Forecast";
            formTitle.style.color = '';
          }
        }
      };

      const hideForm = (): void => {
        formSection.style.display = "none";
        tableSection.style.display = "block";
        if (showFormBtn) showFormBtn.style.display = "inline-block";
        formSection.classList.remove('edit-mode-seibies');
        projectSelect.disabled = false;
      };

      const resetForm = (): void => {
        editingId = null;
        form.reset();
        projectSelect.value = "";
        
        const projectNameDisplay = document.querySelector('#projectNameDisplaySeibies');
        if (projectNameDisplay) projectNameDisplay.textContent = "-";
        const projectRef = document.querySelector('#projectRefSeibies');
        if (projectRef) projectRef.textContent = "-";
        const profitCenter = document.querySelector('#profitCenterSeibies');
        if (profitCenter) profitCenter.textContent = "-";
        const projectManager = document.querySelector('#projectManagerSeibies');
        if (projectManager) projectManager.textContent = "-";
        const groupHead = document.querySelector('#groupHeadSeibies');
        if (groupHead) groupHead.textContent = "-";
        const customer = document.querySelector('#customerSeibies');
        if (customer) customer.textContent = "-";
        
        historicalContainer.innerHTML = "";
        firstRowContainer.innerHTML = "";
        remainingContainer.innerHTML = "";
        yearlyContainer.innerHTML = "";

        toggleForecastSections(false);

        const oblContainer = document.querySelector('.obl-section-container-seibies') as HTMLElement;
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

      const calculateTotalCurrentValue = (): number => {
        let total = 0;

        const editableMonths = allMonths.slice(3);

        editableMonths.forEach(monthKey => {
          const currentInputs = document.querySelectorAll<HTMLInputElement>(`input.main-current-input-seibies[data-month="${monthKey}"]:not([readonly])`);
          currentInputs.forEach(input => {
            total += getNumericValue(input);
          });
        });

        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>(`input.yearly-current-input-seibies:not([readonly])`);
        yearlyCurrentInputs.forEach(input => {
          total += getNumericValue(input);
        });

        return total;
      };

      const validateOBLLimit = (changedInput?: HTMLInputElement): boolean => {
        const oblValue = getNumericValue(oblLastMonthInput);
        const totalCurrent = calculateTotalCurrentValue();

        totalCurrentValueSpan.textContent = totalCurrent.toLocaleString();

        const remaining = Math.abs(oblValue) - totalCurrent;

        remainingValueSpan.textContent = remaining.toLocaleString();

        if (oblValue === 0) {
          oblLineFill.style.width = `0%`;
          oblLineFill.className = 'line-fill-seibies';

          if (totalCurrent > 0) {
            showOBLErrorToast('⚠️ OBL is zero but current values exist. Please set OBL value.', true);
            oblValidationMessage.textContent = '⚠️ OBL is zero but current values exist. Please set OBL value.';
            oblValidationMessage.className = 'validation-message-seibies danger-seibies';
            blockAllCurrentInputs();
            return false;
          } else {
            const existingToast = document.querySelector('.obl-error-toast-seibies');
            if (existingToast) existingToast.remove();
            oblValidationMessage.textContent = '✅ All values are within SEI-BIES limit';
            oblValidationMessage.className = 'validation-message-seibies success-seibies';
            unblockAllCurrentInputs();
            return true;
          }
        }

        const percentage = Math.min((totalCurrent / oblValue) * 100, 100);
        oblLineFill.style.width = `${percentage}%`;

        oblLineFill.className = 'line-fill-seibies';
        if (percentage > 90) {
          oblLineFill.classList.add('danger-seibies');
        } else if (percentage > 75) {
          oblLineFill.classList.add('warning-seibies');
        }

        oblValidationMessage.className = 'validation-message-seibies';

        if (totalCurrent > oblValue) {
          const exceedAmount = totalCurrent - oblValue;
          showOBLErrorToast(`❌ Total current value exceeds SEI-BIES by ${exceedAmount.toLocaleString()}! Reduce values to continue.`);
          oblValidationMessage.textContent = `❌ Total current value exceeds SEI-BIES by ${exceedAmount.toLocaleString()}! Reduce values to continue.`;
          oblValidationMessage.classList.add('danger-seibies');
          blockIncreasingCurrentInputs(changedInput);
          return false;
        } else if (percentage > 90) {
          showOBLErrorToast('⚠️ SEI-BIES limit nearly reached! Consider reducing values.', true);
          oblValidationMessage.textContent = '⚠️ SEI-BIES limit nearly reached! Consider reducing values.';
          oblValidationMessage.classList.add('warning-seibies');
          unblockAllCurrentInputs();
          return true;
        } else {
          const existingToast = document.querySelector('.obl-error-toast-seibies');
          if (existingToast) existingToast.remove();

          oblValidationMessage.textContent = '✅ All values are within SEI-BIES limit';
          oblValidationMessage.classList.add('success-seibies');
          unblockAllCurrentInputs();
          return true;
        }
      };

      const blockAllCurrentInputs = (): void => {
        const currentInputs = document.querySelectorAll<HTMLInputElement>('input[data-type="actual"]:not([readonly])');
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>('input.yearly-current-input-seibies:not([readonly])');

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

      const blockIncreasingCurrentInputs = (allowedInput?: HTMLInputElement): void => {
        const currentInputs = document.querySelectorAll<HTMLInputElement>('input[data-type="actual"]:not([readonly])');
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>('input.yearly-current-input-seibies:not([readonly])');

        currentInputs.forEach(input => {
          if (input === allowedInput) {
            input.style.backgroundColor = '';
            input.style.cursor = '';
            input.removeAttribute('readonly');
          } else {
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
          } else {
            input.setAttribute('data-block-increase', 'true');
            input.style.backgroundColor = '#fff3cd';
            input.style.cursor = 'not-allowed';
          }
        });
      };

      const unblockAllCurrentInputs = (): void => {
        const currentInputs = document.querySelectorAll<HTMLInputElement>('input[data-type="actual"]:not([data-month="0"])');
        const yearlyCurrentInputs = document.querySelectorAll<HTMLInputElement>('input.yearly-current-input-seibies');

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

      const allowInputIncrease = (input: HTMLInputElement, newValue: number): boolean => {
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

      const updateOBLValidationDisplay = (changedInput?: HTMLInputElement): void => {
        validateOBLLimit(changedInput);
      };

      // ========== FORECAST FORM FUNCTIONS ==========

      const toggleForecastSections = (show: boolean): void => {
        const sections = [
          '.obl-section-container-seibies',
          '#historicalForecastSectionSeibies',
          '#detailedForecastSectionSeibies',
          '#longtermForecastSectionSeibies',
          '#yearlyForecastSectionSeibies'
        ];

        sections.forEach(selector => {
          const element = document.querySelector(selector) as HTMLElement;
          if (element) {
            if (show) {
              element.classList.remove('hidden-section');
              element.classList.add('visible-section-seibies');
            } else {
              element.classList.remove('visible-section-seibies');
              element.classList.add('hidden-section');
            }
          }
        });
      };

      // ========== REMAINING FUNCTIONS ==========

      const onProjectChange = (projectId: number, existingForecast?: SeibiesForecastItem): void => {
        const project = projects.find(p => p.Id === projectId);
        if (!project) return;

        toggleForecastSections(true);

        renderProjectDetails(project);

        let forecastData: ForecastData = {};
        if (existingForecast) {
          forecastData = safeJsonParse(existingForecast.ForecastData, {});
        }

        const isNewProject = !existingForecast;

        const oblContainer = document.querySelector('.obl-section-container-seibies') as HTMLElement;
        const fieldNote = oblContainer?.querySelector('.field-note-seibies') as HTMLElement;

        if (isNewProject) {
          if (oblContainer) {
            oblContainer.className = 'obl-section-container-seibies visible-section-seibies new-project-seibies';
          }
          oblLastMonthInput.removeAttribute('readonly');
          oblLastMonthInput.placeholder = "Enter initial OBL for new project";
          if (fieldNote) fieldNote.textContent = "For new projects: Enter initial OBL value";
          oblValidationMessage.textContent = "Enter initial OBL value for new project";
          oblValidationMessage.className = 'validation-message-seibies';

          const defaultOBL = project.PC ? 0 : 0;
          setNumericValue(oblLastMonthInput, defaultOBL);
        } else {
          if (oblContainer) {
            oblContainer.className = 'obl-section-container-seibies visible-section-seibies existing-project-seibies';
          }
          oblLastMonthInput.setAttribute('readonly', 'true');

          const oblValue = existingForecast?.OBL || 0;
          setNumericValue(oblLastMonthInput, oblValue);

          oblLastMonthInput.placeholder = "OBL from database";
          if (fieldNote) fieldNote.textContent = "OBL value from database record";

          if (oblValue > 0) {
            oblLastMonthInput.value = oblValue.toString();
            oblLastMonthInput.className = oblLastMonthInput.className.replace('input-empty-seibies', 'input-with-value-seibies');
          }
        }

        renderForecastInterface(forecastData, existingForecast);
      };

      const getForecastMonths = (): string[] => {
        return generateIdealMonthStructure();
      };

      const renderForecastInterface = (existingData: ForecastData, existingForecast?: SeibiesForecastItem): void => {
        allMonths = getForecastMonths();

        const displayOBL = existingForecast?.OBL || 0;
        setNumericValue(oblLastMonthInput, displayOBL);

        renderHistoricalMonths(allMonths.slice(0, 3), existingData);
        renderFirstRow(allMonths.slice(3, 6), existingData);
        renderRemainingMonths(allMonths.slice(6), existingData);
        renderYearlyForecast(existingData.yearly || createFreshYearlyStructure(getLastMonthKey(existingData)));

        protectHistoricalElements();

        updateOBLValidationDisplay();
        attachForecastEventListeners();

        setTimeout(() => {
          setupHistoricalToggle();
        }, 50);
      };

      // ========== RENDER FUNCTIONS FOR FORECAST SECTIONS ==========

      const renderHistoricalMonths = (months: string[], existingData: ForecastData): void => {
        let html = '';

        months.forEach((monthKey, index) => {
          const monthDate = new Date(monthKey + "-01");
          const monthLabel = formatMonthLabel(monthDate);
          const data = existingData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };

          html += renderHistoricalMonthBox(monthKey, monthLabel, data, index);
        });

        historicalContainer.innerHTML = html;
      };

      const renderHistoricalMonthBox = (monthKey: string, monthLabel: string, data: ForecastMonthData, index: number): string => {
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

      const renderFirstRow = (months: string[], existingData: ForecastData): void => {
        let html = '';

        html += `<div class="detailed-months-row-seibies">`;

        const currentMonth = months[0];
        if (currentMonth) {
          const monthDate = new Date(currentMonth + "-01");
          const monthLabel = formatMonthLabel(monthDate);
          const data = existingData[currentMonth] as ForecastMonthData || { plan: 0, actual: 0 };

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
            const data = existingData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };

            html += `<div class="next-month-box-seibies">`;
            html += renderDetailedMonthBox(monthKey, monthLabel, data, false);
            html += `</div>`;
          }
        }
        html += `</div>`;

        html += `</div>`;

        firstRowContainer.innerHTML = html;
      };

      const renderRemainingMonths = (months: string[], existingData: ForecastData): void => {
        let html = '';

        for (let row = 0; row < 3; row++) {
          html += `<div class="remaining-row-seibies">`;
          for (let col = 0; col < 5; col++) {
            const index = (row * 5) + col;
            if (index >= months.length) break;

            const monthKey = months[index];
            const monthDate = new Date(monthKey + "-01");
            const monthLabel = formatMonthLabel(monthDate);
            const data = existingData[monthKey] as ForecastMonthData || { plan: 0, actual: 0 };

            html += renderSimpleMonthBox(monthKey, monthLabel, data);
          }
          html += `</div>`;
        }

        remainingContainer.innerHTML = html;
      };

      const renderYearlyForecast = (yearlyData: YearlyForecastData): void => {
        let html = '';

        const years = Object.keys(yearlyData).sort();

        years.forEach(year => {
          const data = yearlyData[year] || { plan: 0, actual: 0 };
          html += renderYearlyBox(year, data);
        });

        yearlyContainer.innerHTML = html;
      };

      const renderYearlyBox = (year: string, data: { plan: number; actual: number }): string => {
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

      const getCurrentMonthMaxDate = (): string => {
        const now = new Date();
        const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        return `${lastDay.getFullYear()}-${String(lastDay.getMonth() + 1).padStart(2, "0")}-${String(lastDay.getDate()).padStart(2, "0")}`;
      };

      /**
       * ✅ renderDetailedMonthBox
       *    - Current month  : Previous / Current / Date  (Further input REMOVED)
       *    - Next 2 months  : Previous / Current
       *    The data model still carries monthData.forecast so previously-saved
       *    "Further" values are preserved on save (just not shown/editable here).
       */
      const renderDetailedMonthBox = (monthKey: string, monthLabel: string, data: ForecastMonthData, isCurrentMonth: boolean = false): string => {
        const displayPlanValue = data.plan > 0 ? data.plan.toString() : '';
        const displayActualValue = data.actual > 0 ? data.actual.toString() : '';
        const displayDateValue = data.date || '';

        const planInputClass = data.plan > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
        const actualInputClass = data.actual > 0 ? 'input-with-value-seibies' : 'input-empty-seibies';
        const dateInputClass = displayDateValue ? 'input-with-value-seibies' : 'input-empty-seibies';

        return `
          <div class="detailed-month-box-seibies ${isCurrentMonth ? 'current-month-box-seibies' : 'next-month-box-seibies'}">
            <div class="month-header-seibies">
              ${monthLabel}
              ${isCurrentMonth ? '<span class="current-month-indicator-seibies">Current</span>' : ''}
            </div>

            <div class="month-block-seibies">
              <div class="enhanced-inputs-container-seibies ${isCurrentMonth ? 'current-month-inputs-seibies' : ''}">
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
                    <label>Date</label>
                    <input type="date" value="${displayDateValue}" data-type="date" data-month="${monthKey}" max="${getCurrentMonthMaxDate()}" title="Select a date up to the end of the current month"
                          class="forecast-date-input-seibies ${dateInputClass}" />
                  </div>
                ` : ''}
              </div>
            </div>
          </div>
        `;
      };

      const renderSimpleMonthBox = (monthKey: string, monthLabel: string, data: ForecastMonthData): string => {
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

      const protectHistoricalElements = (): void => {
        document.querySelectorAll('.historical-month-seibies input, .historical-month-seibies select, .historical-month-seibies button').forEach(element => {
          const newElement = element.cloneNode(true) as HTMLElement;
          element.parentNode?.replaceChild(newElement, element);
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

      const attachForecastEventListeners = (): void => {
        document.querySelectorAll<HTMLInputElement>('.number-input-seibies').forEach(input => {
          allowOnlyNumbers(input);
        });

        document.querySelectorAll<HTMLInputElement>('.number-input-seibies:not([readonly])').forEach(input => {
          input.addEventListener('focus', function() {
            this.setAttribute('data-old-value', getNumericValue(this).toString());
          });
        });

        document.querySelectorAll<HTMLInputElement>('.main-current-input-seibies:not([readonly])').forEach(input => {
          input.addEventListener('input', function() {
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
            } else {
              this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
            }

            updateOBLValidationDisplay(this);
          });
        });

        document.querySelectorAll<HTMLInputElement>('.yearly-current-input-seibies:not([readonly])').forEach(input => {
          input.addEventListener('input', function() {
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
            } else {
              this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
            }

            updateOBLValidationDisplay(this);
          });
        });

        document.querySelectorAll<HTMLInputElement>('.forecast-date-input-seibies:not([readonly])').forEach(input => {
          input.addEventListener('change', function() {
            const maxAllowed = this.getAttribute('max') || getCurrentMonthMaxDate();
            if (this.value && this.value > maxAllowed) {
              this.value = '';
              this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
              toast("warning", `Date cannot be beyond the current month (max ${maxAllowed})`);
              return;
            }

            if (this.value === '') {
              this.className = this.className.replace('input-with-value-seibies', 'input-empty-seibies');
            } else {
              this.className = this.className.replace('input-empty-seibies', 'input-with-value-seibies');
            }
          });
        });

        const validationInputs = document.querySelectorAll<HTMLInputElement>(`
          .main-current-input-seibies:not([readonly]),
          .yearly-current-input-seibies:not([readonly])
        `);

        validationInputs.forEach(input => {
          input.addEventListener('input', function() {
            setTimeout(() => {
              updateOBLValidationDisplay(this);
            }, 100);
          });
        });
      };

const editForecast = async (id: number): Promise<void> => {
  editingId = id;
  const forecast = forecasts.find(f => f.Id === id);
  
  if (!forecast) {
    toast("error", "SEIBIES forecast not found");
    return;
  }

  if (!canRead()) {
    toast("error", "❌ You don't have permission to view this forecast");
    return;
  }

  const project = projects.find(p => p.Id === forecast.Project?.Id);
  const canEdit = canEditProject(project);
  const isFrozen = forecast.Freeze || false;

  showForm(true, project);

  let forecastData: ForecastData = {};

  try {
    let forecastDataString = forecast.ForecastData || "{}";
    forecastDataString = forecastDataString.trim();

    if (!forecastDataString || forecastDataString === '""' || forecastDataString === "null") {
      forecastDataString = "{}";
    }

    forecastData = JSON.parse(forecastDataString);

  } catch (e) {
    forecastData = createFreshForecastStructure();
    toast("warning", "Corrupted SEIBIES forecast data detected. Loaded empty structure.");
  }

  forecastData = propagateMonthValues(forecastData);

  if (forecast.Project) {
    projectSelect.value = String(forecast.Project.Id);
    onProjectChange(forecast.Project.Id, forecast);
  }

  setTimeout(() => {
    const isReadOnly = isFrozen || !canEdit;
    
    setFormReadOnly(isReadOnly);

    // ✅ Show/hide BOTH save buttons together (top + bottom)
    const saveBtns = getSaveButtons();
    saveBtns.forEach(btn => {
      btn.style.display = isReadOnly ? 'none' : '';
    });

    const formTitleEl = document.querySelector("#formTitleSeibies") as HTMLElement;
if (formTitleEl) {
  const projectName = project?.Title || "";
  const projectReference = project?.ProjectReference || "";
  const projectTitle = projectName
    ? `${projectName}${projectReference ? ` (${projectReference})` : ""}`
    : "";

  if (isFrozen) {
    formTitleEl.innerHTML = `🔒 ${projectTitle || "SEIBIES Forecast"} (Frozen - Read Only)`;
    formTitleEl.style.color = '#856404';
  } else if (!canEdit) {
    formTitleEl.innerHTML = `🔒 ${projectTitle || "SEIBIES Forecast"} (Read Only)`;
    formTitleEl.style.color = '#856404';
  } else {
    formTitleEl.innerHTML = projectTitle || "";
    formTitleEl.style.color = '';
  }
}
  }, 300);

  toast("info", `Editing SEIBIES forecast for ${forecast.Title}`);
};

      // ========== FORM SUBMISSION ==========

      form.addEventListener("submit", async (e: Event) => {
        e.preventDefault();

        const projectId = parseInt(projectSelect.value);
        const project = projects.find(p => p.Id === projectId);

        if (!project) {
          toast("error", "Please select a project");
          return;
        }

        if (editingId) {
          if (!canEditProject(project)) {
            toast("error", "❌ You don't have permission to edit this SEIBIES forecast");
            return;
          }
        } else {
          if (!canCreateRecord()) {
            toast("error", "❌ You don't have permission to create SEIBIES forecasts");
            return;
          }
        }

        const isNewProject = !editingId && !forecasts.some(f => f.Project?.Id === projectId);

        if (!isNewProject && forecasts.some(f => f.Project?.Id === projectId && f.Id !== editingId)) {
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

        // ✅ Disable BOTH save buttons while saving
        const saveBtns = getSaveButtons();
        saveBtns.forEach(btn => {
          btn.disabled = true;
          btn.dataset.saving = "true";
        });

        showRevenueLoader(editingId ? 'Updating' : 'Saving',
          editingId ? 'Updating the SEIBIES forecast on SharePoint...' : 'Saving the SEIBIES forecast to SharePoint...');

        try {
          const forecastData: ForecastData = {};

          const allStorageMonths = getAllMonthsForStorage();

          allStorageMonths.forEach(monthKey => {
            const mainCurrentInput = document.querySelector<HTMLInputElement>(`input.main-current-input-seibies[data-month="${monthKey}"]`);
            const mainPlanInput = document.querySelector<HTMLInputElement>(`input.main-plan-input-seibies[data-month="${monthKey}"]`);
            const mainDateInput = document.querySelector<HTMLInputElement>(`input.forecast-date-input-seibies[data-month="${monthKey}"]`);

            if (mainCurrentInput && mainPlanInput) {
              const monthData: ForecastMonthData = {
                plan: getNumericValue(mainPlanInput),
                actual: getNumericValue(mainCurrentInput)
              };

              // ✅ Preserve any previously-saved "Further" value (input removed from UI,
              //    but the field still exists on the existing forecast data)
              const existingData = safeJsonParse(forecasts.find(f => f.Id === editingId)?.ForecastData, {});
              const prevForecast = (existingData[monthKey] as ForecastMonthData)?.forecast;
              if (typeof prevForecast === 'number' && prevForecast > 0) {
                monthData.forecast = prevForecast;
              }

              if (mainDateInput && mainDateInput.value) {
                monthData.date = mainDateInput.value;
              }

              forecastData[monthKey] = monthData;
            }
          });

          const yearlyData: YearlyForecastData = {};
          const yearlyPlanInputs = document.querySelectorAll<HTMLInputElement>('.yearly-plan-input-seibies');

          if (yearlyPlanInputs.length > 0) {
            const years = new Set<string>();
            yearlyPlanInputs.forEach(input => {
              const year = input.dataset.year;
              if (year) years.add(year);
            });

            Array.from(years).sort().forEach(year => {
              const planInput = document.querySelector<HTMLInputElement>(`input.yearly-plan-input-seibies[data-year="${year}"]`);
              const currentInput = document.querySelector<HTMLInputElement>(`input.yearly-current-input-seibies[data-year="${year}"]`);

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

          const payload: any = {
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

          const headers: any = {
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
            hideRevenueLoader();
          } else {
            hideRevenueLoader();
            const errorText = await res.text();
            console.error("SharePoint Error:", errorText);
            toast("error", `Error saving SEIBIES forecast: ${res.status}`);
          }
        } catch (error) {
          hideRevenueLoader();
          console.error("Error saving SEIBIES forecast:", error);
          toast("error", "Error saving SEIBIES forecast!");
        } finally {
          // ✅ Re-enable BOTH save buttons
          saveBtns.forEach(btn => {
            btn.disabled = false;
            btn.dataset.saving = "false";
          });
        }
      });

      // ========== SET FORM READONLY ==========

const setFormReadOnly = (isReadOnly: boolean): void => {
  const editableSelectors = [
    '.main-current-input-seibies',
    '.yearly-current-input-seibies',
    '.forecast-date-input-seibies'
  ];

  const alwaysReadonlySelectors = [
    '.main-plan-input-seibies',
    '.yearly-plan-input-seibies'
  ];

  const inputs = document.querySelectorAll<HTMLInputElement>(editableSelectors.join(','));

  inputs.forEach(input => {
    if (isReadOnly) {
      input.setAttribute('readonly', 'true');
      input.style.backgroundColor = '#f5f5f5';
      input.style.color = '#6c757d';
      input.style.cursor = 'not-allowed';
      input.style.opacity = '0.7';
      input.style.pointerEvents = 'none';
    } else {
      input.removeAttribute('readonly');
      input.style.backgroundColor = '';
      input.style.color = '';
      input.style.cursor = '';
      input.style.opacity = '1';
      input.style.pointerEvents = '';
    }
  });

  const readonlyInputs = document.querySelectorAll<HTMLInputElement>(alwaysReadonlySelectors.join(','));

  readonlyInputs.forEach(input => {
    input.setAttribute('readonly', 'true');
    if (isReadOnly) {
      input.style.backgroundColor = '#f5f5f5';
      input.style.color = '#6c757d';
      input.style.cursor = 'not-allowed';
      input.style.opacity = '0.7';
      input.style.pointerEvents = 'none';
    } else {
      input.style.backgroundColor = '';
      input.style.color = '';
      input.style.cursor = '';
      input.style.opacity = '';
      input.style.pointerEvents = '';
    }
  });

  const oblInput = document.querySelector('#oblLastMonthSeibies') as HTMLInputElement;
  if (oblInput) {
    oblInput.setAttribute('readonly', 'true');
    if (isReadOnly) {
      oblInput.style.backgroundColor = '#f5f5f5';
      oblInput.style.cursor = 'not-allowed';
    } else {
      oblInput.style.backgroundColor = '';
      oblInput.style.cursor = '';
    }
  }

  if (isReadOnly) {
    toast("info", '🔒 This project is read-only.');
  } else {
    toast("info", '🔓 Project is now editable.');
  }
};

      const propagateCollectionValues = (forecastData: ForecastData): ForecastData => {
        const updatedData: ForecastData = { ...forecastData };

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

      const propagateYearlyValues = (forecastData: ForecastData, isNewMonth: boolean = false): ForecastData => {
        if (!isNewMonth) {
          return forecastData;
        }

        const updatedData: ForecastData = { ...forecastData };

        if (!updatedData.yearly || Object.keys(updatedData.yearly).length === 0) {
          return updatedData;
        }

        const yearlyData = updatedData.yearly;
        const updatedYearlyData: YearlyForecastData = {};

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

      // ========== EVENT LISTENERS ==========

      showFormBtn?.addEventListener("click", () => {
        editingId = null;
        resetForm();
        showForm(false);
      });

      backBtn.addEventListener("click", () => {
        hideForm();
        resetForm();
        editingId = null;
        loadForecasts();
      });

      projectSelect.addEventListener("change", (e: Event) => {
        const projectId = parseInt((e.target as HTMLSelectElement).value);

        if (projectId) {
          onProjectChange(projectId);
        } else {
          toggleForecastSections(false);
        }
      });

      clearFiltersBtn.addEventListener("click", clearFilters);

      exportBtn?.addEventListener("click", exportToExcelWithFullStyle);

      // ========== INITIAL LOADS ==========

      showRevenueLoader();

      (window as any).__revenueLoaderDeadline = Date.now() + 30000;
      window.setTimeout(() => {
        const loader = document.querySelector('#revenuePageLoader');
        if (loader && Date.now() > Number((window as any).__revenueLoaderDeadline || 0)) {
          console.warn('⏱ Loader deadline exceeded - forcing hide');
          hideRevenueLoader();
        }
      }, 31000);

      try {
        await authManager.initialize({ siteUrl });

        console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
        console.log(`👤 User: ${getCurrentUserFullName()}`);
        console.log(`📋 User Groups:`, getUserGroups());
        console.log(`✅ Full Control: ${hasFullControl()} | Can Contribute: ${canContribute()} | Can Read: ${canRead()}`);
        console.log(`✅ Can Read All Projects: ${canReadAllProjects()} | Can Freeze: ${canFreeze()}`);
        console.log(`✅ Is GH Head: ${isGHHead()} | Is PM Manager: ${isPMManager()} | Is Engg User: ${isEngineeringUser()}`);

        await loadProjects();
        await loadForecasts();

        setupPaginationEventListeners();

        initializePagination(forecasts);

        updateUIBasedOnPermissions({
          exportBtn: ".export-btn-seibies",
          addBtn: ".show-form-btn-seibies",
          saveBtn: '.btn-primary-seibies[type="submit"]'
        });

      } catch (error) {
        console.error('❌ Error initializing SEIBIES module:', error);
        toast("error", "Failed to initialize application");
      } finally {
        hideRevenueLoader();
      }

      console.log("✅ SEIBIES Forecast Module Initialized - Default page size: 25 per page");

    }, 100);
  }
}; 
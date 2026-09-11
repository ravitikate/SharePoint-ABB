// src/webparts/digitalTransformationApp1/DigitalTransformationApp1WebPart.ts
// 🔧 ACTIVE-TAB FIX:
//   The active-tab highlight is now driven from mountPage(), the single choke
//   point through which every successful page open flows. Previously the
//   highlight could only be set by a document-level capture click listener in
//   the Sidebar, so opening Master (or any page) via the async admin gate /
//   the boot property / the dashboard fallback / the "Go to Dashboard" button
//   opened the page but never highlighted its tab.
//
//   mountPage() now calls setActiveSidebarPage(pageKey, false) — which clears
//   every other tab and marks the one matching the page actually opened.
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import sidebarModule, { getSidebarContent, setActiveSidebarPage } from './components/Sidebar';
import headerModule from './components/Header';
import footerModule from './components/Footer';
import dashboardPage from './components/pages/dashboard';
import projectPage from './components/pages/project';
import dciPage from './components/pages/dci';
import bomPage from './components/pages/bom';
import procurementPage from './components/pages/procurement';
import revenuePage from './components/pages/revenue';
import collectionPage from './components/pages/collection';
import knowledgePage from './components/pages/knowledge';
import invoicingPage from './components/pages/invoicing';
import inventoryPage from './components/pages/inventory';
import seibiesPage from './components/pages/seibies';
import masterPage from './components/pages/master';
import sprintPage from './components/pages/dailysprint';
// Import auth system 
import { authManager, hasFullControl, 
// canContribute, 
// canRead,
getCurrentUserFullName, getUserGroups, getPermissionLevel } from './components/pages/auth';
const pages = {
    dashboard: dashboardPage,
    master: masterPage,
    project: projectPage,
    sprint: sprintPage,
    dci: dciPage,
    bom: bomPage,
    procurement: procurementPage,
    revenue: revenuePage,
    collection: collectionPage,
    knowledge: knowledgePage,
    invoicing: invoicingPage,
    inventory: inventoryPage,
    seibies: seibiesPage
};
/** 🔧 Pages that only Full-Control / admin-group users may open.
 *  Row- and action-level role access for the other pages is enforced inside
 *  each page module (same auth.ts helpers). */
const ADMIN_ONLY_PAGES = new Set(['master']);
export default class DigitalTransformationApp1WebPart extends BaseClientSideWebPart {
    constructor() {
        super(...arguments);
        this.isAuthInitialized = false;
        /** 🔧 SINGLE-FLIGHT GUARD (fixes the double-initialisation seen in the console log):
         *  SharePoint / the workbench can call render() more than once (double mount,
         *  property-pane refresh, canvas re-render). Without a guard two async renders
         *  run concurrently → the full auth/sidebar/page sequence logs twice, <style>
         *  tags and module scripts pile up, and listeners get attached twice.
         *  Subsequent render() calls while one is in flight now just await the first. */
        this.renderInFlight = null;
        /** 🔧 Tracks the page currently shown inside the already-mounted shell so a
         *  re-render (property change / canvas remount) only reloads when the
         *  requested page really differs. */
        this.activePageKey = '';
    }
    render() {
        if (this.renderInFlight)
            return this.renderInFlight;
        this.renderInFlight = this.doRender().finally(() => {
            this.renderInFlight = null;
        });
        return this.renderInFlight;
    }
    /** 🔧 Idempotent style injection: same-id style tags are removed before the new
     *  one is appended, so repeated renders never leave duplicate <style> blocks. */
    appendStyleTag(css, id) {
        if (id) {
            const existing = this.domElement.querySelector(`style#${id}`);
            if (existing)
                existing.remove();
        }
        const tag = document.createElement('style');
        if (id)
            tag.id = id;
        tag.innerHTML = css;
        this.domElement.appendChild(tag);
    }
    async doRender() {
        try {
            /** HIDE EVERYTHING FROM SHAREPOINT (FULL APP MODE) */
            this.hideSharePointUI();
            // Initialize auth system first
            await this.initializeAuth();
            // 🔧 BOOT-ONCE: SharePoint / the workbench can call render() repeatedly
            // (property-pane change, canvas remount, page refresh in edit mode). If the
            // shell already exists, do NOT rebuild the header/sidebar/styles or re-run
            // the module scripts a second (third…) time — that was producing the
            // duplicated logs + duplicated listeners/styles seen in the console.
            const existingShell = this.domElement.querySelector('.dt-app');
            if (existingShell) {
                const targetPage = this.properties.page || 'dashboard';
                if (targetPage !== this.activePageKey) {
                    console.log(`♻️ App shell already mounted – switching to "${targetPage}"`);
                    this.loadPage(targetPage);
                }
                else {
                    console.log(`♻️ App shell already mounted – keeping current page "${this.activePageKey}"`);
                }
                return;
            }
            /** Get dynamic sidebar content based on user role */
            const sidebarHtml = await getSidebarContent(this.context);
            /** MAIN HTML */
            this.domElement.innerHTML = `
        <div class="dt-app">
          <header class="dt-header">${headerModule.content}</header>
          <div class="dt-main">
            <aside class="dt-sidebar">${sidebarHtml}</aside>
            <main class="dt-body">Welcome to Digital Transformation</main>
          </div>
          <footer class="dt-footer">${footerModule.content}</footer>
        </div>
      `;
            this.applyGlobalStyles();
            /** Load component styles/scripts */
            [headerModule, footerModule].forEach((mod, i) => {
                if (mod.style) {
                    this.appendStyleTag(mod.style, `dt-module-style-${i}`);
                }
                if (mod.script) {
                    try {
                        mod.script();
                    }
                    catch (error) {
                        console.error('❌ Component script error (header/footer):', error);
                    }
                }
            });
            // Add sidebar styles
            this.appendStyleTag(sidebarModule.style, 'dt-sidebar-style');
            if (sidebarModule.script) {
                try {
                    sidebarModule.script();
                }
                catch (error) {
                    console.error('❌ Sidebar script error:', error);
                }
            }
            this.attachSidebarEvents();
            this.attachSidebarToggle();
            this.loadPage(this.properties.page || 'dashboard');
        }
        catch (error) {
            console.error('❌ Failed to render application:', error);
            this.domElement.innerHTML = `
        <div style="display:flex;align-items:center;justify-content:center;height:100vh;background:#f5f5f9;">
          <div style="text-align:center;padding:40px;background:#fff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);max-width:420px;">
            <h2 style="color:#dc3545;margin:0 0 12px 0;">❌ Application failed to load</h2>
            <p style="color:#666;font-size:14px;margin:0;word-break:break-word;">
              ${error instanceof Error ? error.message : String(error)}
            </p>
            <p style="color:#999;font-size:13px;margin:12px 0 0 0;">Please refresh the page. If the problem persists, check the browser console (F12).</p>
          </div>
        </div>
      `;
        }
    }
    /** ---------------------------
     *   INITIALIZE AUTH SYSTEM
     * --------------------------- */
    async initializeAuth() {
        if (this.isAuthInitialized)
            return;
        const siteUrl = this.context.pageContext.web.absoluteUrl;
        try {
            // Initialize auth with site URL
            await authManager.initialize({ siteUrl });
            this.isAuthInitialized = true;
            console.log('🔑 Auth initialized successfully');
            console.log(`👤 User: ${getCurrentUserFullName()}`);
            console.log(`📋 Groups:`, getUserGroups());
            console.log(`🔑 Permission Level: ${getPermissionLevel().toUpperCase()}`);
        }
        catch (error) {
            console.error('Failed to initialize auth:', error);
        }
    }
    /** ---------------------------
     *   FULL APP MODE
     *   REMOVE ALL SHAREPOINT UI
     * --------------------------- */
    hideSharePointUI() {
        const hideList = [
            "#SuiteNavWrapper",
            ".ms-compositeHeader",
            ".ms-CommandBar",
            ".spAppBar",
            "#spCommandBar",
            "header[data-automationid='SiteHeader']",
            "div[data-automationid='PageHeader']",
            "div[data-automationid='PageTitle']",
            "div[data-automationid='PageBreadcrumb']",
            "footer",
            ".od-SuiteNav",
            ".CanvasZone",
            ".CanvasSection",
            ".spLeftNav",
            "div[data-automationid='PageNavigation']"
        ];
        hideList.forEach(sel => {
            const el = document.querySelector(sel);
            if (el)
                el.style.display = "none";
        });
        /** Make Page Background White / Clean */
        const canvas = document.querySelector(".SPPageChrome");
        if (canvas) {
            canvas.style.margin = "0";
            canvas.style.padding = "0";
            canvas.style.background = "#f5f5f9";
        }
        /** Make WebPart FULL SCREEN */
        const wpContainer = this.domElement.closest(".ControlZone");
        if (wpContainer) {
            wpContainer.style.position = "fixed";
            wpContainer.style.top = "0";
            wpContainer.style.left = "0";
            wpContainer.style.width = "100vw";
            wpContainer.style.height = "100vh";
            wpContainer.style.margin = "0";
            wpContainer.style.padding = "0";
            wpContainer.style.zIndex = "99999";
            wpContainer.style.background = "#f5f5f9";
            wpContainer.style.overflow = "hidden";
        }
        /** Remove Scrolling of SharePoint Page */
        document.body.style.margin = "0";
        document.body.style.padding = "0";
        document.body.style.overflow = "hidden";
        document.documentElement.style.overflow = "hidden";
        console.log("✔ SharePoint hidden — FULL APP MODE ENABLED");
    }
    applyGlobalStyles() {
        this.appendStyleTag(`
      html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
        overflow: hidden;
      }
      .dt-app {
        display: flex;
        flex-direction: column;
        width: 100vw;
        height: 100vh;
        background: #f5f5f9;
        overflow: hidden;
      }
      .dt-main {
        display: flex;
        flex: 1;
        overflow: hidden;
      }
      .dt-sidebar {
        width: 230px;
        overflow-y: auto;
        background: #ffffff;
        border-right: 1px solid #eee;
      }
      .dt-body {
        flex: 1;
        overflow-y: auto;
        padding: 2px;
        background: #fafafa;
      }
      .dt-header, .dt-footer {
        flex-shrink: 0;
        background: #fff;
      }
    `, 'dt-global-style');
    }
    /** 🔧 Unified page mount: removes the previous page style, injects the new
     *  content + style, runs the page script inside try/catch and resets the
     *  scroll position of the body so navigation always starts at the top.
     *
     *  🔧 ACTIVE-TAB FIX: we set the active sidebar tab HERE (before injecting
     *  content) because mountPage is the single guaranteed path for every page
     *  open. setActiveSidebarPage clears all other tabs and activates the one
     *  matching the page being opened. persist=false keeps sessionStorage
     *  untouched for internal navigations. */
    mountPage(body, page, pageKey) {
        // Track currently displayed page
        this.activePageKey = pageKey;
        // Keep sidebar highlight synchronized with the page
        setActiveSidebarPage(pageKey, false);
        // 🔧 DEFAULT-TAB GUARANTEE: set the highlight directly on the sidebar links
        // (covers first open on Dashboard + any async sidebar tab (re)build).
        this.highlightSidebarTab(pageKey);
        // Render page
        body.innerHTML = page.content;
        // Remove previous page style
        const prevStyle = this.domElement.querySelector('#page-style');
        if (prevStyle) {
            prevStyle.remove();
        }
        // Add new page style
        if (page.style) {
            const styleEl = document.createElement('style');
            styleEl.id = 'page-style';
            styleEl.innerHTML = page.style;
            this.domElement.appendChild(styleEl);
        }
        // Reset scroll position
        if (body.scrollTop !== undefined) {
            body.scrollTop = 0;
        }
        // Run page script
        if (page.script) {
            try {
                page.script();
            }
            catch (error) {
                console.error(`❌ Script error on page "${pageKey}":`, error);
            }
        }
    }
    /** 🔧 DEFAULT/ACTIVE TAB: highlight the sidebar link for the page actually
     *  open (Dashboard by default on first open). Sets the class DIRECTLY on the
     *  links so it works even when the Sidebar module's own highlight pass can't
     *  find them yet (async getSidebarContent / role-based tab rebuild) — retries
     *  briefly until the links exist. On the very first visit it also stores the
     *  default page in sessionStorage so a refresh keeps it highlighted. */
    highlightSidebarTab(pageKey, attempt = 0) {
        // A newer navigation may have happened while we were waiting — never
        // steal the highlight from a tab the user already clicked.
        if (this.activePageKey !== pageKey)
            return;
        const links = this.domElement.querySelectorAll('.dt-sidebar a[data-page]');
        if (links.length === 0) {
            // Sidebar tabs not in the DOM yet — retry for up to ~1 second
            if (attempt < 20) {
                window.setTimeout(() => this.highlightSidebarTab(pageKey, attempt + 1), 50);
            }
            return;
        }
        let found = false;
        links.forEach(link => {
            const isActive = (link.dataset.page || '') === pageKey;
            link.classList.toggle('active', isActive);
            if (isActive) {
                link.setAttribute('aria-current', 'page');
                found = true;
            }
            else {
                link.setAttribute('aria-current', 'false');
            }
        });
        // First visit: nothing stored yet → persist the default (dashboard)
        try {
            if (found && !window.sessionStorage.getItem('dt-sidebar-active-page')) {
                window.sessionStorage.setItem('dt-sidebar-active-page', pageKey);
            }
        }
        catch (error) {
            // Ignore storage errors
        }
    }
    /** 🔧 Shared access-denied screen (used by Master + any admin-only page),
     *  with the dashboard button re-bound on every mount. */
    showAccessDenied(body, pageKey) {
        this.activePageKey = pageKey;
        body.innerHTML = `
      <div style="text-align: center; padding: 50px; background: #fff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
        <h2 style="color: #dc3545;">🔒 Access Denied</h2>
        <p style="color: #666; font-size: 16px; margin: 20px 0;">
          You don't have permission to access the ${pageKey === 'master' ? 'Master Data' : pageKey} page.
        </p>
        <p style="color: #999; font-size: 14px; margin-bottom: 20px;">
          Current User: ${getCurrentUserFullName()} | 
          Permission Level: ${getPermissionLevel().toUpperCase()}
        </p>
        <button data-page="dashboard" style="padding: 12px 30px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">
          Go to Dashboard
        </button>
      </div>
    `;
        const goToDashboardBtn = body.querySelector('[data-page="dashboard"]');
        if (goToDashboardBtn) {
            goToDashboardBtn.addEventListener('click', () => {
                this.loadPage('dashboard');
            });
        }
    }
    loadPage(pageKey) {
        const body = this.domElement.querySelector('.dt-body');
        if (!body)
            return;
        // Check if user can access master page (admin-only)
        if (pageKey === 'master' || ADMIN_ONLY_PAGES.has(pageKey)) {
            this.checkAndLoadMasterPage(body, pageKey);
            return;
        }
        const page = pages[pageKey];
        if (!page) {
            // 🔧 Unknown page → never leave the shell on a dead screen
            console.warn(`⚠️ Unknown page "${pageKey}" – falling back to dashboard`);
            this.mountPage(body, pages['dashboard'], 'dashboard');
            return;
        }
        this.mountPage(body, page, pageKey);
    }
    /** 🔧 Admin-only gate (Master page): Full Control OR an admin group. */
    canAccessMaster() {
        const userGroups = getUserGroups();
        const isAdminGroup = userGroups.some(group => group === 'KRIYA_Admin_Full' ||
            group === 'Site Collection Administrators' ||
            group === 'Super Admin');
        return hasFullControl() || isAdminGroup;
    }
    /** 🔧 Used by the sidebar so denied pages never even start to load. */
    isPageAllowed(pageKey) {
        if (ADMIN_ONLY_PAGES.has(pageKey))
            return this.canAccessMaster();
        // All other pages are readable by any authenticated user; row/action level
        // role access (create/edit/freeze/export) is enforced inside each module.
        return true;
    }
    async checkAndLoadMasterPage(body, pageKey) {
        // Ensure auth is initialized
        if (!this.isAuthInitialized) {
            await this.initializeAuth();
        }
        const hasAccess = this.canAccessMaster();
        if (hasAccess) {
            console.log(`✅ User ${getCurrentUserFullName()} has access to Master Data page`);
            const page = pages['master'];
            if (page) {
                this.mountPage(body, page, 'master');
            }
        }
        else {
            console.log(`❌ User ${getCurrentUserFullName()} denied access to Master Data page`);
            this.showAccessDenied(body, pageKey);
        }
    }
    attachSidebarEvents() {
        const links = this.domElement.querySelectorAll('.dt-sidebar a[data-page]');
        links.forEach(link => {
            link.addEventListener('click', (e) => {
                e.preventDefault();
                const clickedLink = e.currentTarget;
                const page = clickedLink.dataset.page;
                if (!page)
                    return;
                // ==========================================
                // ACTIVE SIDEBAR TAB
                // ==========================================
                links.forEach(item => {
                    const itemLink = item;
                    itemLink.classList.remove('active');
                    itemLink.setAttribute('aria-current', 'false');
                });
                // Highlight clicked tab immediately
                clickedLink.classList.add('active');
                clickedLink.setAttribute('aria-current', 'page');
                // Store selected page
                try {
                    window.sessionStorage.setItem('dt-sidebar-active-page', page);
                }
                catch (error) {
                    // Ignore storage errors
                }
                // ==========================================
                // PAGE ACCESS CHECK
                // ==========================================
                if (!this.isPageAllowed(page)) {
                    const body = this.domElement.querySelector('.dt-body');
                    if (body) {
                        this.showAccessDenied(body, page);
                    }
                    return;
                }
                // ==========================================
                // LOAD PAGE
                // ==========================================
                this.loadPage(page);
            });
        });
    }
    attachSidebarToggle() {
        const toggleBtn = this.domElement.querySelector('.toggle-sidebar');
        const sidebar = this.domElement.querySelector('.dt-sidebar');
        if (toggleBtn && sidebar) {
            toggleBtn.addEventListener('click', () => sidebar.classList.toggle('collapsed'));
        }
    }
    get dataVersion() {
        return Version.parse('1.0');
    }
}
//# sourceMappingURL=DigitalTransformationApp1WebPart.js.map
// src/webparts/Kriya/KriyaWebPart.ts
// 🔧 ACTIVE-TAB FIX:
//   The active-tab highlight is now driven from mountPage(), the single choke
//   point through which every successful page open flows.
//
// 🔧 USER-READY FIX:
//   notifyUserReady() now sends the full user identity — name + login + email +
//   groups + isAdmin + isGH + canContribute + permissionLevel — so the header
//   can gate the "Add User" button on KRIYA_Admin_Full membership.
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
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
import './components/styles/pageHeaders.css';
// Import auth system
import { authManager, hasFullControl, canContribute, isGHHead, getCurrentUserFullName, getUserGroups, getPermissionLevel } from './components/pages/auth';
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
/** 🔧 Pages that only Full-Control / admin-group users may open. */
const ADMIN_ONLY_PAGES = new Set(['master']);
export default class KriyaWebPart extends BaseClientSideWebPart {
    constructor() {
        super(...arguments);
        this.isAuthInitialized = false;
        /** 🔧 SINGLE-FLIGHT GUARD: prevents double-initialisation if render()
         *  is called more than once by SharePoint / the workbench. */
        this.renderInFlight = null;
        /** 🔧 Tracks the page currently shown inside the already-mounted shell. */
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
    /** 🔧 Idempotent style injection. */
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
            this.hideSharePointUI();
            await this.initializeAuth();
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
                this.notifyUserReady();
                return;
            }
            const sidebarHtml = await getSidebarContent(this.context);
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
            // First fire — user payload may not be read yet, but harmless.
            this.notifyUserReady();
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
            // Second fire — this one always lands after the header's listener is attached.
            window.setTimeout(() => this.notifyUserReady(), 0);
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
    /** 🔧 Full user identity broadcaster.
     *  Sends name + login + email + groups + isAdmin + isGH + canContribute +
     *  permissionLevel so the header can decide whether to show the
     *  "Add User" button (visible only for KRIYA_Admin_Full). */
    notifyUserReady() {
        let name = '';
        let login = '';
        let email = '';
        let groups = [];
        let isAdmin = false;
        let isGH = false;
        let canWrite = false;
        let permissionLevel = '';
        try {
            name = getCurrentUserFullName() || '';
        }
        catch ( /* auth not ready */_a) { /* auth not ready */ }
        try {
            groups = getUserGroups() || [];
        }
        catch ( /* auth not ready */_b) { /* auth not ready */ }
        try {
            isAdmin = hasFullControl();
        }
        catch ( /* auth not ready */_c) { /* auth not ready */ }
        try {
            isGH = isGHHead();
        }
        catch ( /* auth not ready */_d) { /* auth not ready */ }
        try {
            canWrite = canContribute();
        }
        catch ( /* auth not ready */_e) { /* auth not ready */ }
        try {
            permissionLevel = getPermissionLevel() || '';
        }
        catch ( /* auth not ready */_f) { /* auth not ready */ }
        const sp = window._spPageContextInfo || {};
        if (!name) {
            name = sp.userDisplayName || '';
        }
        if (!login) {
            login = sp.userLoginName || '';
        }
        if (!email) {
            email = sp.userEmail || '';
        }
        try {
            if (!login) {
                login = this.context.pageContext.user.loginName || '';
            }
            if (!email) {
                email = this.context.pageContext.user.email || '';
            }
            if (!name) {
                name = this.context.pageContext.user.displayName || '';
            }
        }
        catch ( /* context not ready */_g) { /* context not ready */ }
        if (!name && !login) {
            return;
        }
        try {
            document.dispatchEvent(new CustomEvent('dt-user-ready', {
                detail: {
                    name,
                    login,
                    email,
                    groups,
                    isAdmin,
                    isGH,
                    canContribute: canWrite,
                    permissionLevel
                }
            }));
            // tslint:disable-next-line:no-console
            console.log('📣 dt-user-ready fired:', {
                name, login, groups: groups.length, isAdmin, isGH, canWrite, permissionLevel
            });
        }
        catch (e) {
            console.warn('⚠️ Could not fire dt-user-ready:', e);
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
        const canvas = document.querySelector(".SPPageChrome");
        if (canvas) {
            canvas.style.margin = "0";
            canvas.style.padding = "0";
            canvas.style.background = "#f5f5f9";
        }
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
    mountPage(body, page, pageKey) {
        this.activePageKey = pageKey;
        setActiveSidebarPage(pageKey, false);
        this.highlightSidebarTab(pageKey);
        body.innerHTML = page.content;
        const prevStyle = this.domElement.querySelector('#page-style');
        if (prevStyle) {
            prevStyle.remove();
        }
        if (page.style) {
            const styleEl = document.createElement('style');
            styleEl.id = 'page-style';
            styleEl.innerHTML = page.style;
            this.domElement.appendChild(styleEl);
        }
        if (body.scrollTop !== undefined) {
            body.scrollTop = 0;
        }
        if (page.script) {
            try {
                page.script();
            }
            catch (error) {
                console.error(`❌ Script error on page "${pageKey}":`, error);
            }
        }
    }
    highlightSidebarTab(pageKey, attempt = 0) {
        if (this.activePageKey !== pageKey)
            return;
        const links = this.domElement.querySelectorAll('.dt-sidebar a[data-page]');
        if (links.length === 0) {
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
        try {
            if (found && !window.sessionStorage.getItem('dt-sidebar-active-page')) {
                window.sessionStorage.setItem('dt-sidebar-active-page', pageKey);
            }
        }
        catch (_a) {
            // Ignore storage errors
        }
    }
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
        if (pageKey === 'master' || ADMIN_ONLY_PAGES.has(pageKey)) {
            this.checkAndLoadMasterPage(body, pageKey);
            return;
        }
        const page = pages[pageKey];
        if (!page) {
            console.warn(`⚠️ Unknown page "${pageKey}" – falling back to dashboard`);
            this.mountPage(body, pages['dashboard'], 'dashboard');
            return;
        }
        this.mountPage(body, page, pageKey);
    }
    canAccessMaster() {
        const userGroups = getUserGroups();
        const isAdminGroup = userGroups.some(group => group === 'KRIYA_Admin_Full' ||
            group === 'Site Collection Administrators' ||
            group === 'Super Admin');
        return hasFullControl() || isAdminGroup;
    }
    isPageAllowed(pageKey) {
        if (ADMIN_ONLY_PAGES.has(pageKey))
            return this.canAccessMaster();
        return true;
    }
    async checkAndLoadMasterPage(body, pageKey) {
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
                links.forEach(item => {
                    const itemLink = item;
                    itemLink.classList.remove('active');
                    itemLink.setAttribute('aria-current', 'false');
                });
                clickedLink.classList.add('active');
                clickedLink.setAttribute('aria-current', 'page');
                try {
                    window.sessionStorage.setItem('dt-sidebar-active-page', page);
                }
                catch (_a) {
                    // Ignore storage errors
                }
                if (!this.isPageAllowed(page)) {
                    const body = this.domElement.querySelector('.dt-body');
                    if (body) {
                        this.showAccessDenied(body, page);
                    }
                    return;
                }
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
//# sourceMappingURL=KriyaWebPart.js.map
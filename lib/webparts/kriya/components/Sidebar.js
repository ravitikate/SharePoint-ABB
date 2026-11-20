import { authManager, hasFullControl, canContribute, isGHHead, getUserGroups } from './pages/auth';
/* ============================================================
   ROLE MATRIX (from your Excel screenshot)
   ------------------------------------------------------------
   Group Name            | Sidebar tabs allowed
   ------------------------------------------------------------
   KRIYA_GH_Heads        | ALL (Project Specific, Contribute)
   KRIYA_PM_Managers     | ALL (PM Specific, Contribute)
   KRIYA_Engg_Lead       | ALL (DCI/Proc = Contribute, Finance = Read Only)
   KRIYA_Engg_Sub1       | ALL (same as Engg Lead)
   KRIYA_Engg_Sub2       | ALL (same as Engg Lead)
   KRIYA_Engg_Mgr        | ALL (same as Engg Lead)
   KRIYA_SC_Team         | ONLY DCI + PROCUREMENT (Read)
   KRIYA_Controlling     | ONLY REVENUE + COLLECTION + INVOICING
                         |      + INVENTORY + SEI/BIES (Read)
   Admin (Full Control)  | EVERYTHING incl. MASTER DATA
   ------------------------------------------------------------
   NOTE 1: Table Row1/Row2 labels look swapped
   (Project Managers -> KRIYA_GH_Heads, Group Heads -> KRIYA_PM_Managers).
   Code below uses the GROUP NAME column, so it works either way
   since both groups get full access.
   NOTE 2: "Project Specific / PM Specific / All Projects" is DATA
   filtering inside each page, not sidebar. Sidebar only shows/hides
   tabs. Use canEditPage() + your project filter in each page.
   ============================================================ */
export const GROUPS = {
    GH_HEADS: 'KRIYA_GH_Heads',
    PM_MANAGERS: 'KRIYA_PM_Managers',
    ENGG_LEAD: 'KRIYA_Engg_Lead',
    ENGG_SUB1: 'KRIYA_Engg_Sub1',
    ENGG_SUB2: 'KRIYA_Engg_Sub2',
    ENGG_MGR: 'KRIYA_Engg_Mgr',
    SC_TEAM: 'KRIYA_SC_Team',
    CONTROLLING: 'KRIYA_Controlling'
};
// Groups that see the full sidebar (except Master Data)
const ALL_PROJECT_GROUPS = [
    GROUPS.GH_HEADS,
    GROUPS.PM_MANAGERS,
    GROUPS.ENGG_LEAD,
    GROUPS.ENGG_SUB1,
    GROUPS.ENGG_SUB2,
    GROUPS.ENGG_MGR
];
// Full-contribute groups (can edit every module in their scope)
const FULL_CONTRIBUTE_GROUPS = [
    GROUPS.GH_HEADS,
    GROUPS.PM_MANAGERS
];
// Engg groups (Contribute on Dashboard/DCI/Procurement, Read Only on finance modules)
const ENGG_GROUPS = [
    GROUPS.ENGG_LEAD,
    GROUPS.ENGG_SUB1,
    GROUPS.ENGG_SUB2,
    GROUPS.ENGG_MGR
];
// Order here = order in the sidebar. Kept exactly as your original code.
const PAGE_DEFS = [
    { page: 'dashboard', label: 'DASHBOARD', iconKey: 'dashboard' },
    { page: 'master', label: 'MASTER DATA', iconKey: 'master' }, // admin only
    { page: 'project', label: 'PROJECT', iconKey: 'project' },
    { page: 'sprint', label: 'DAILY SPRINT', iconKey: 'dailysprint' },
    { page: 'revenue', label: 'REVENUE PLAN', iconKey: 'revenue' },
    { page: 'collection', label: 'COLLECTION PLAN', iconKey: 'collection' },
    { page: 'dci', label: 'DCI', iconKey: 'dci' },
    { page: 'procurement', label: 'PROCUREMENT PLAN', iconKey: 'procurement' },
    { page: 'invoicing', label: 'INVOICING PLAN', iconKey: 'invoicing' },
    { page: 'inventory', label: 'INVENTORY LIQUIDATION', iconKey: 'inventory' },
    { page: 'seibies', label: 'SEI BIES LIQUIDATION', iconKey: 'seibies' },
    { page: 'knowledge', label: 'KNOWLEDGE BASE', iconKey: 'knowledge' }
];
// Standard full sidebar (everyone in ALL_PROJECT_GROUPS) — everything except Master
const STANDARD_PAGES = [
    'dashboard', 'project', 'sprint', 'revenue', 'collection',
    'dci', 'procurement', 'invoicing', 'inventory', 'seibies', 'knowledge'
];
// SCM team: ONLY DCI + Procurement (per your requirement + Excel blank = no Dashboard)
// 👉 If you later want Dashboard for SCM too, change to: ['dashboard', 'dci', 'procurement']
const SCM_PAGES = ['dci', 'procurement'];
// Controlling team: ONLY the 5 finance tabs (per your requirement + Excel blank = no Dashboard)
// 👉 If you later want Dashboard for Controlling too, add 'dashboard' to this array.
const CONTROLLING_PAGES = ['revenue', 'collection', 'invoicing', 'inventory', 'seibies'];
// Engg groups: editable vs read-only split (from Excel "Read Only" columns)
const ENGG_EDITABLE_PAGES = ['dashboard', 'project', 'sprint', 'dci', 'procurement', 'knowledge'];
const ENGG_READONLY_PAGES = ['revenue', 'collection', 'invoicing', 'inventory', 'seibies'];
const icon = (path) => `<svg viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${path}</svg>`;
const ICONS = {
    dashboard: icon(`<rect x="3" y="3" width="7" height="7" rx="1"/>
     <rect x="14" y="3" width="7" height="7" rx="1"/>
     <rect x="3" y="14" width="7" height="7" rx="1"/>
     <rect x="14" y="14" width="7" height="7" rx="1"/>`),
    master: icon(`<ellipse cx="12" cy="5" rx="9" ry="3"/>
     <path d="M3 5v6c0 1.66 4 3 9 3s9-1.34 9-3V5"/>
     <path d="M3 11v6c0 1.66 4 3 9 3s9-1.34 9-3v-6"/>`),
    dailysprint: icon(`<rect x="3" y="4" width="18" height="18" rx="2"/>
     <line x1="16" y1="2" x2="16" y2="6"/>
     <line x1="8" y1="2" x2="8" y2="6"/>
     <line x1="3" y1="10" x2="21" y2="10"/>
     <polyline points="8 15 11 18 16 13"/>`),
    project: icon(`<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>`),
    dci: icon(`<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
     <polyline points="14 2 14 8 20 8"/>
     <line x1="16" y1="13" x2="8" y2="13"/>
     <line x1="16" y1="17" x2="8" y2="17"/>
     <line x1="10" y1="9" x2="8" y2="9"/>`),
    procurement: icon(`<circle cx="9" cy="21" r="1"/>
     <circle cx="20" cy="21" r="1"/>
     <path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/>`),
    revenue: icon(`<path d="M12 1v22"/>
     <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>`),
    collection: icon(`<rect x="2" y="6" width="20" height="12" rx="2"/>
     <circle cx="12" cy="12" r="2.5"/>
     <path d="M6 10v4M18 10v4"/>`),
    invoicing: icon(`<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
     <polyline points="14 2 14 8 20 8"/>
     <path d="M12 18v-6"/>
     <path d="M9.5 13.5h4a1.5 1.5 0 0 0 0-3h-3a1.5 1.5 0 0 0 0 3"/>
     <path d="M9.5 15.5h5"/>`),
    inventory: icon(`<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/>
     <polyline points="3.27 6.96 12 12.01 20.73 6.96"/>
     <line x1="12" y1="22.08" x2="12" y2="12"/>`),
    seibies: icon(`<line x1="18" y1="20" x2="18" y2="10"/>
     <line x1="12" y1="20" x2="12" y2="4"/>
     <line x1="6" y1="20" x2="6" y2="14"/>
     <line x1="3" y1="20" x2="21" y2="20"/>`),
    knowledge: icon(`<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/>
     <path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>`)
};
const navItem = (page, label, iconKey) => `
  <li>
    <a href="#" data-page="${page}" aria-current="false" aria-label="${label}" title="${label}">
      <span class="icon" data-tooltip="${label}">${ICONS[iconKey]}</span>
      <span class="icon-label">${label}</span>
    </a>
  </li>
`;
/* ============================================================
   ACTIVE TAB HIGHLIGHTING
   ============================================================ */
/** Page shown when nothing has been selected yet (for users who HAVE dashboard access). */
const DEFAULT_PAGE = 'dashboard';
/** sessionStorage key so the highlight survives a page refresh. */
const ACTIVE_PAGE_KEY = 'dt-sidebar-active-page';
/** Fired after the active item changes, in case other components care. */
export const SIDEBAR_PAGE_CHANGED_EVENT = 'dt-sidebar-page-changed';
let activePage = DEFAULT_PAGE;
let activeStateWired = false;
/** Allowed pages for the current user (set by getSidebarContent, used to validate highlight/routing). */
let cachedAllowedPages = null;
const getSidebarRoot = () => document.querySelector('.dt-sidebar');
/** The page key currently highlighted. */
export const getActiveSidebarPage = () => activePage;
/**
 * Highlight the nav item for `page` and clear every other one.
 * Safe to call before the sidebar is in the DOM — the value is
 * remembered and applied by initSidebarActiveState().
 */
export const setActiveSidebarPage = (page, persist = true) => {
    if (!page) {
        return;
    }
    activePage = page;
    if (persist) {
        try {
            window.sessionStorage.setItem(ACTIVE_PAGE_KEY, page);
        }
        catch (e) {
            /* private mode / storage disabled — highlight still works for this session */
        }
    }
    const root = getSidebarRoot();
    if (!root) {
        return;
    }
    const links = root.querySelectorAll('a[data-page]');
    for (let i = 0; i < links.length; i++) {
        const link = links[i];
        const isActive = link.getAttribute('data-page') === page;
        if (isActive) {
            link.classList.add('active');
            link.setAttribute('aria-current', 'page');
        }
        else {
            link.classList.remove('active');
            link.setAttribute('aria-current', 'false');
        }
    }
    try {
        root.dispatchEvent(new CustomEvent(SIDEBAR_PAGE_CHANGED_EVENT, { detail: { page: page }, bubbles: true }));
    }
    catch (e) {
        /* older browsers without CustomEvent constructor — highlight already applied */
    }
};
/** Read the page to highlight on first paint: URL hash → sessionStorage → default, validated against allowed pages. */
const resolveInitialPage = () => {
    // Make sure we know what this user is allowed to see BEFORE picking a page.
    if (!cachedAllowedPages) {
        try {
            cachedAllowedPages = getAllowedPagesSync();
        }
        catch (e) {
            /* auth not ready — fall through, we'll re-validate below */
        }
    }
    const allowed = cachedAllowedPages;
    // 1) Try URL hash first
    const hash = (window.location.hash || '').replace(/^#\/?/, '').split('?')[0].trim();
    if (hash && (!allowed || allowed.length === 0 || allowed.indexOf(hash) !== -1)) {
        return hash;
    }
    // 2) Try last stored page — but only if still allowed
    try {
        const stored = window.sessionStorage.getItem(ACTIVE_PAGE_KEY);
        if (stored && (!allowed || allowed.length === 0 || allowed.indexOf(stored) !== -1)) {
            return stored;
        }
    }
    catch (e) { /* ignore */ }
    // 3) Fall back to the user's FIRST allowed page (DCI for SCM, Revenue for Controlling, Dashboard for Project users)
    if (allowed && allowed.length > 0) {
        return allowed[0];
    }
    // 4) Only when we truly have no auth info do we default to dashboard
    return DEFAULT_PAGE;
};
/**
 * Wire up the highlight. Call once after the sidebar HTML is injected.
 *
 * Uses a capture-phase delegated listener so the highlight is applied
 * even if the host WebPart's own click handler calls stopPropagation()
 * or preventDefault() for routing. Routing behaviour is untouched.
 */
export const initSidebarActiveState = (initialPage) => {
    const root = getSidebarRoot();
    if (!root) {
        return;
    }
    // If getSidebarContent hasn't run yet (auth not cached), try to resolve allowed pages now.
    if (!cachedAllowedPages) {
        try {
            cachedAllowedPages = getAllowedPagesSync();
        }
        catch (e) { /* auth not ready yet — highlight proceeds without validation */ }
    }
    if (!activeStateWired) {
        document.addEventListener('click', (evt) => {
            const target = evt.target;
            if (!target || !target.closest) {
                return;
            }
            const link = target.closest('.dt-sidebar a[data-page]');
            if (!link) {
                return;
            }
            // clicking a tab navigates away — never leave a bubble behind
            hideSidebarTooltip();
            const page = link.getAttribute('data-page');
            if (page) {
                setActiveSidebarPage(page);
            }
        }, true // capture
        );
        // Keep the highlight in sync if the app routes via the URL hash.
        // Blocked hashes (e.g. #dashboard for SCM user) are ignored so the
        // highlight stays on a tab the user is actually allowed to see.
        window.addEventListener('hashchange', () => {
            const h = (window.location.hash || '').replace(/^#\/?/, '').split('?')[0].trim();
            if (!h) {
                return;
            }
            if (cachedAllowedPages && cachedAllowedPages.length > 0 && cachedAllowedPages.indexOf(h) === -1) {
                return; // not allowed — keep current highlight
            }
            setActiveSidebarPage(h);
        });
        activeStateWired = true;
    }
    // collapsed-rail tooltips (idempotent — safe on every init)
    wireSidebarTooltips();
    setActiveSidebarPage(initialPage || resolveInitialPage(), false);
};
/* ============================================================
   COLLAPSED-RAIL TOOLTIPS
   ------------------------------------------------------------
   The bubble used to be a :hover::after pseudo-element on .icon with
   left: calc(100% + 14px), i.e. ~14px OUTSIDE the 60px rail. The rail is a
   scroll container (overflow-y: auto) and, per CSS Overflow Level 3, the
   visible/clip values of overflow compute to auto/hidden when the other axis
   is neither visible nor clip — so .dt-sidebar.collapsed's overflow-x: visible
   computed to auto and the rail clipped the bubble at its own edge.

   A pseudo-element can never escape an ancestor scroll container, so the bubble
   is now a real element appended to <body> and placed with position: fixed from
   the hovered row's bounding rect. Nothing clips it, it keeps working while the
   rail is scrolled, and hovering the whole row works (not just the 22px icon).
   ============================================================ */
const TOOLTIP_CLASS = 'dt-sidebar-tooltip';
const TOOLTIP_OFFSET_PX = 16;
const NATIVE_TITLE_FALLBACK_ATTR = 'data-title-fallback';
let tooltipEl = null;
let tooltipLink = null;
let tooltipsWired = false;
/** Expanded rails already show the text label — tooltips are for the collapsed rail. */
const isRailCollapsed = (el) => {
    const rail = el.closest('.dt-sidebar');
    return !!rail && rail.classList.contains('collapsed');
};
const closestNavLink = (target) => {
    const el = target;
    if (!el || !el.closest) {
        return null;
    }
    return el.closest('.dt-sidebar a[data-page]');
};
const iconOfLink = (link) => link.querySelector('.icon[data-tooltip]');
/**
 * Move the link's native title into a data attribute the first time the pointer
 * touches the row: otherwise the OS bubble and the styled one show together.
 * This is also the no-JS fallback — if this script never runs, title stays put.
 */
const stashNativeTitle = (link) => {
    const title = link.getAttribute('title');
    if (title && !link.getAttribute(NATIVE_TITLE_FALLBACK_ATTR)) {
        link.setAttribute(NATIVE_TITLE_FALLBACK_ATTR, title);
        link.removeAttribute('title');
    }
};
/** Remove the bubble. Safe to call at any time (no bubble → no-op). */
export const hideSidebarTooltip = () => {
    if (tooltipEl && tooltipEl.parentNode) {
        tooltipEl.parentNode.removeChild(tooltipEl);
    }
    tooltipEl = null;
    tooltipLink = null;
};
const placeTooltip = (bubble, icon) => {
    const rect = icon.getBoundingClientRect();
    bubble.style.top = Math.round(rect.top + rect.height / 2) + 'px';
    bubble.style.left = Math.round(rect.right + TOOLTIP_OFFSET_PX) + 'px';
};
const showSidebarTooltip = (link) => {
    const icon = iconOfLink(link);
    const label = icon ? icon.getAttribute('data-tooltip') : null;
    if (!icon || !label || !isRailCollapsed(link)) {
        hideSidebarTooltip();
        return;
    }
    // same row still hovered → follow it, don't rebuild (no flicker moving svg ↔ label)
    if (tooltipEl && tooltipLink === link) {
        placeTooltip(tooltipEl, icon);
        return;
    }
    hideSidebarTooltip();
    tooltipLink = link;
    const bubble = document.createElement('div');
    bubble.className = TOOLTIP_CLASS;
    bubble.setAttribute('role', 'presentation');
    bubble.textContent = label;
    placeTooltip(bubble, icon);
    document.body.appendChild(bubble);
    const reveal = () => { bubble.classList.add('is-visible'); };
    if (typeof window.requestAnimationFrame === 'function') {
        window.requestAnimationFrame(reveal);
    }
    else {
        reveal();
    }
    tooltipEl = bubble;
};
/**
 * Wire the collapsed-rail tooltips. Delegated on document, so the listeners
 * survive the WebPart re-rendering the sidebar HTML. Idempotent.
 */
export const wireSidebarTooltips = () => {
    if (tooltipsWired) {
        return;
    }
    if (!getSidebarRoot()) {
        return;
    }
    document.addEventListener('mouseover', (evt) => {
        const link = closestNavLink(evt.target);
        if (!link) {
            return;
        }
        stashNativeTitle(link);
        showSidebarTooltip(link);
    });
    document.addEventListener('mouseout', (evt) => {
        const link = closestNavLink(evt.target);
        if (!link) {
            return;
        }
        const next = evt.relatedTarget;
        // moving between the row, the icon and the label must not flicker the bubble
        if (next && link.contains(next)) {
            return;
        }
        hideSidebarTooltip();
    });
    // Keyboard: focusing a tab explains it as well (aria-label keeps the name either way).
    document.addEventListener('focusin', (evt) => {
        const link = closestNavLink(evt.target);
        if (link) {
            showSidebarTooltip(link);
        }
    });
    document.addEventListener('focusout', () => { hideSidebarTooltip(); });
    // Anything that moves the rail under the cursor invalidates the fixed position.
    window.addEventListener('scroll', hideSidebarTooltip, true);
    window.addEventListener('resize', hideSidebarTooltip);
    document.addEventListener('click', hideSidebarTooltip, true);
    tooltipsWired = true;
    console.log('[Sidebar] Collapsed-rail tooltips wired (fixed bubble escapes the rail clip)');
};
/**
 * Initialize authentication for the sidebar
 * This should be called before any permission checks
 */
export const initializeSidebarAuth = async (context) => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    await authManager.initialize({ siteUrl });
};
/**
 * Get current user's groups using the shared auth module
 */
export const getCurrentUserGroups = async () => {
    const groups = getUserGroups();
    console.log('Logged-in user belongs to groups:');
    groups.forEach((group) => {
        console.log(`Name: ${group}`);
    });
};
/**
 * Check if user is in a specific group using the shared auth module (case-insensitive)
 */
export const isUserInGroup = (groupName) => {
    const groups = getUserGroups();
    return groups.some(group => group.toLowerCase() === groupName.toLowerCase());
};
/** Convenience role checks */
export const isSCTeam = () => isUserInGroup(GROUPS.SC_TEAM);
export const isControllingTeam = () => isUserInGroup(GROUPS.CONTROLLING);
export const isEnggUser = () => ENGG_GROUPS.some(g => isUserInGroup(g));
export const isFullContributeUser = () => FULL_CONTRIBUTE_GROUPS.some(g => isUserInGroup(g));
export const isProjectUser = () => ALL_PROJECT_GROUPS.some(g => isUserInGroup(g));
/**
 * Compute the sidebar pages the current user may SEE.
 * Must be called AFTER authManager.initialize() (i.e. after initializeSidebarAuth).
 * Users in multiple groups get the UNION of all their groups' pages.
 */
export const getAllowedPagesSync = () => {
    // Admin sees everything including Master Data
    if (hasFullControl()) {
        const all = [];
        for (let i = 0; i < PAGE_DEFS.length; i++) {
            all.push(PAGE_DEFS[i].page);
        }
        return all;
    }
    const inSC = isSCTeam();
    const inControlling = isControllingTeam();
    const inProject = isProjectUser();
    if (!inSC && !inControlling && !inProject) {
        // User is in none of the KRIYA groups — show nothing (caller renders "No access").
        // If you'd rather show Dashboard by default, change this to: return ['dashboard'];
        return [];
    }
    const map = {};
    let i = 0;
    if (inProject) {
        for (i = 0; i < STANDARD_PAGES.length; i++) {
            map[STANDARD_PAGES[i]] = true;
        }
    }
    if (inSC) {
        for (i = 0; i < SCM_PAGES.length; i++) {
            map[SCM_PAGES[i]] = true;
        }
    }
    if (inControlling) {
        for (i = 0; i < CONTROLLING_PAGES.length; i++) {
            map[CONTROLLING_PAGES[i]] = true;
        }
    }
    // Preserve sidebar order; Master Data stays admin-only
    const ordered = [];
    for (i = 0; i < PAGE_DEFS.length; i++) {
        const pg = PAGE_DEFS[i].page;
        if (pg === 'master') {
            continue;
        }
        if (map[pg]) {
            ordered.push(pg);
        }
    }
    return ordered;
};
/** Safe wrapper that never throws (falls back to cached pages, then dashboard). */
export const getAllowedPagesSyncSafe = () => {
    try {
        return getAllowedPagesSync();
    }
    catch (e) {
        return cachedAllowedPages || [DEFAULT_PAGE];
    }
};
/** First tab the user is allowed to see — use as landing page after login / for blocked routes. */
export const getFirstAllowedPage = () => {
    const allowed = cachedAllowedPages || getAllowedPagesSyncSafe();
    return allowed.length > 0 ? allowed[0] : DEFAULT_PAGE;
};
/**
 * Router guard — call this in your WebPart routing before rendering a page:
 *
 *   import { isPageAllowed, getFirstAllowedPage } from './sidebar';
 *   if (!isPageAllowed(page)) { page = getFirstAllowedPage(); }
 *
 * This blocks direct URL-hash access (e.g. SCM user typing #revenue).
 */
export const isPageAllowed = (page) => {
    if (!page) {
        return false;
    }
    try {
        if (hasFullControl()) {
            return true;
        }
    }
    catch (e) { /* auth not ready */ }
    const allowed = cachedAllowedPages || getAllowedPagesSyncSafe();
    return allowed.indexOf(page) !== -1;
};
/**
 * Edit permission per page (Contribute vs Read Only from the Excel matrix):
 * - Admin / GH_Heads / PM_Managers → can edit everything in their scope
 * - Engg groups → can edit Dashboard/DCI/Procurement (+ Project/Sprint/Knowledge),
 *                 READ ONLY on Revenue/Collection/Invoicing/Inventory/SEI-BIES
 * - SC_Team / Controlling → READ ONLY everywhere (Permission Level = Read)
 * - Master Data → Admin only
 *
 * Usage in a page (hide/disable Save buttons):
 *
 *   import { canEditPage } from '../sidebar';
 *   const readOnly = !canEditPage('revenue');
 *
 * NOTE: "Project Specific / PM Specific / All Projects" row filtering
 * (show only my projects) must still be done in each page with your
 * project-assignment list — sidebar only controls tab visibility.
 */
export const canEditPage = (page) => {
    try {
        if (hasFullControl()) {
            return true;
        }
    }
    catch (e) { /* auth not ready — fall through */ }
    if (page === 'master') {
        return false;
    } // only admin (handled above)
    // Full contribute groups can edit every module in their scope
    if (isFullContributeUser()) {
        return true;
    }
    // Engg users: editable on Dashboard/DCI/Procurement family, read-only on finance family
    if (isEnggUser()) {
        if (ENGG_EDITABLE_PAGES.indexOf(page) !== -1) {
            return true;
        }
        if (ENGG_READONLY_PAGES.indexOf(page) !== -1) {
            return false;
        }
        // Unknown page key — fall through to canContribute()
    }
    // SC / Controlling are Read-only everywhere per matrix.
    // (If the user were ALSO in an edit-granting group, we'd have returned true above.)
    if (isSCTeam() || isControllingTeam()) {
        // User ONLY in read groups (or Engg viewing a finance page) → deny edit.
        // If user is in NO edit group at all, deny; otherwise defer to canContribute().
        if (!isEnggUser() && !isFullContributeUser()) {
            return false;
        }
        // Engg user on an unknown page key + also in a read group → defer to canContribute()
    }
    try {
        return canContribute();
    }
    catch (e) {
        return false;
    }
};
// Generate sidebar content based on user role
export const getSidebarContent = async (context) => {
    await initializeSidebarAuth(context);
    const isAdmin = hasFullControl();
    const isGH = isGHHead();
    const canUserContribute = canContribute();
    const userFullName = authManager.getCurrentUserFullName();
    const userGroups = getUserGroups();
    const currentUser = context.pageContext.user;
    console.log('Logged-in User:');
    console.log('Display Name:', currentUser.displayName);
    console.log('Email:', currentUser.email);
    console.log('Login Name:', currentUser.loginName);
    console.log('User ID:', context.pageContext.legacyPageContext.userId);
    await getCurrentUserGroups();
    console.log('All Site Groups (from auth):', userGroups);
    console.log('User Info:', {
        name: currentUser.displayName,
        email: currentUser.email,
        login: currentUser.loginName,
        id: context.pageContext.legacyPageContext.userId,
        fullName: userFullName,
        isAdmin: isAdmin,
        isGH: isGH,
        canContribute: canUserContribute,
        isSC: isSCTeam(),
        isControlling: isControllingTeam(),
        isEngg: isEnggUser()
    });
    console.log('Is Kirya Admin?:', isAdmin);
    console.log('Is GH Head?:', isGH);
    // ---- Role-based tab filtering (Excel matrix) ----
    const allowed = getAllowedPagesSync();
    cachedAllowedPages = allowed;
    console.log('✅ Allowed sidebar pages for this user:', allowed);
    if (allowed.length === 0) {
        return `
      <ul class="dt-sidebar-list">
        <li style="padding:16px 12px;color:#b00020;font-size:13px;">
          No module access assigned.<br/>Please contact your administrator.
        </li>
      </ul>
    `;
    }
    // data-page values kept as-is so existing page routing still matches
    let sidebarItems = '';
    for (let i = 0; i < PAGE_DEFS.length; i++) {
        const def = PAGE_DEFS[i];
        if (allowed.indexOf(def.page) === -1) {
            continue;
        }
        if (def.page === 'master' && !isAdmin) {
            continue;
        } // double-guard: Master = admin only
        sidebarItems += navItem(def.page, def.label, def.iconKey);
    }
    // 🔧 FIXED: return only the <ul> - the WebPart already wraps this in
    // <aside class="dt-sidebar">. Previously this returned ANOTHER
    // <div class="dt-sidebar"> inside the aside → duplicated sidebar
    // container (double width/padding/shadow + duplicate styling).
    return `
    <ul class="dt-sidebar-list">
      ${sidebarItems}
    </ul>
  `;
};
export const sidebarStyle = `
@font-face {
  font-family: 'ABBvoiceOffice';
  src: url('../../assets/ABBvoiceOffice_Rg.ttf') format('truetype');
  font-weight: 600;
  font-style: normal;
  font-display: swap;
}

html, body, .page, .page * { font-family: 'ABBvoiceOffice'; }

/* ============================================
   SIDEBAR SHELL — light gray like the second image
   ============================================ */
.dt-sidebar {
  background: #f4f5f7;
  border-right: 1px solid #e6e8eb;
  box-shadow: none;
  width: 225px;
  padding: 8px 6px;
  transition: width 0.3s ease;
  height: 100%;
  max-height: 100%;
  overflow-x: hidden;
  overflow-y: auto;
  font-family: ABBvoiceOffice;
  box-sizing: border-box; 
}

.dt-sidebar.collapsed {
  width: 60px;
  padding: 8px 6px;
  overflow-x: visible;
  overflow-y: auto;
}

.dt-sidebar ul { list-style: none; padding: 0; margin: 0; }
.dt-sidebar li { margin: 4px 0; }

.dt-sidebar a {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 12px;
  color: #4d5f71;
  text-decoration: none;
  border-radius: 8px;
  position: relative;
  white-space: nowrap;
  font-weight: 500;
  font-size: 13px;
  transition: background 0.18s ease, color 0.18s ease;
}

.dt-sidebar a:hover {
  background: #e9ebef;
  color: #2c3e50;
}

/* ============================================
   ACTIVE / SELECTED TAB — soft purple-blue
   (same as second image)
   ============================================ */
.dt-sidebar a.active {
  background: linear-gradient(90deg, #eef0ff 0%, #e2e5ff 100%);
  color: #4a4fd6;
  font-weight: 700;
}

.dt-sidebar a.active:hover {
  background: linear-gradient(90deg, #e6e9ff 0%, #d8dcff 100%);
  color: #3e43c7;
}

/* Left accent bar — same purple-blue */
.dt-sidebar a.active::before {
  content: '';
  position: absolute;
  left: 0;
  top: 50%;
  transform: translateY(-50%);
  width: 3px;
  height: 60%;
  background: #4a4fd6;
  border-radius: 0 3px 3px 0;
}

/* Icon tint matches the active color */
.dt-sidebar a.active .icon svg { stroke: #4a4fd6; }
.dt-sidebar a:hover  .icon svg { stroke: #2c3e50; }

/* ============================================
   COLLAPSED RAIL — same soft purple highlight
   ============================================ */
.dt-sidebar.collapsed a.active {
  background: linear-gradient(180deg, #eef0ff 0%, #e2e5ff 100%);
  color: #4a4fd6;
  border-radius: 10px;
}

.dt-sidebar.collapsed a.active::before {
  left: -6px;
  height: 70%;
  background: #4a4fd6;
}

/* Focus ring */
.dt-sidebar a:focus-visible {
  outline: 2px solid #4a4fd6;
  outline-offset: -2px;
}

.dt-sidebar.collapsed a {
  justify-content: center;
  padding: 8px;
}
.dt-sidebar.collapsed .icon-label { display: none; }

/* Icon sizing */
.icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  position: relative;
  flex-shrink: 0;
  width: 22px;
  height: 22px;
}
.icon svg {
  display: block;
  width: 20px;
  height: 20px;
  transition: stroke 0.18s ease;
}

/* Collapsed-rail tooltip bubble. */
/* Rendered as a real <div> appended to <body> by wireSidebarTooltips() and pinned
   with position: fixed — a pseudo-element inside .dt-sidebar can never escape the
   rail's own scroll container (overflow-y: auto forces the visible overflow-x to
   compute to auto), which is why the old ::after bubble never showed up. The look
   is unchanged: #333 bubble, 11px, radius 4, arrow on the rail side. */
.dt-sidebar-tooltip {
  position: fixed;
  z-index: 99999;
  background: #333;
  color: #fff;
  font-size: 11px;
  line-height: 1.2;
  padding: 4px 8px;
  border-radius: 4px;
  white-space: nowrap;
  pointer-events: none;
  opacity: 0;
  transform: translateY(-50%) translateX(-4px);
  transition: opacity 0.12s ease, transform 0.12s ease;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.18);
}
.dt-sidebar-tooltip.is-visible {
  opacity: 1;
  transform: translateY(-50%) translateX(0);
}
.dt-sidebar-tooltip::before {
  content: '';
  position: absolute;
  left: -8px;
  top: 50%;
  transform: translateY(-50%);
  border-width: 5px;
  border-style: solid;
  border-color: transparent #333 transparent transparent;
}

/* Thin scrollbar */
.dt-sidebar::-webkit-scrollbar { width: 6px; }
.dt-sidebar::-webkit-scrollbar-track { background: transparent; }
.dt-sidebar::-webkit-scrollbar-thumb {
  background: #c5c5d2;
  border-radius: 8px;
}
.dt-sidebar::-webkit-scrollbar-thumb:hover { background: #a8a8b8; }

.dt-sidebar {
  scrollbar-width: thin;
  scrollbar-color: #c5c5d2 transparent;
}
`;
export const sidebarScript = (initialPage) => {
    // Guard the requested landing page — e.g. don't land an SCM user on dashboard
    let requested = initialPage;
    if (requested &&
        cachedAllowedPages &&
        cachedAllowedPages.length > 0 &&
        cachedAllowedPages.indexOf(requested) === -1) {
        // tslint:disable-next-line:no-console
        console.warn(`[Sidebar] Initial page "${requested}" not allowed for this user — falling back to "${cachedAllowedPages[0]}"`);
        requested = cachedAllowedPages[0];
    }
    initSidebarActiveState(requested);
    console.log('Sidebar loaded with role-based visibility + active tab highlighting');
};
export default {
    getContent: getSidebarContent,
    style: sidebarStyle,
    script: sidebarScript,
    setActivePage: setActiveSidebarPage,
    getActivePage: getActiveSidebarPage,
    isPageAllowed: isPageAllowed,
    getFirstAllowedPage: getFirstAllowedPage,
    getAllowedPages: getAllowedPagesSyncSafe,
    canEditPage: canEditPage
};
//# sourceMappingURL=Sidebar.js.map
import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface ISidebarProps {
    context: WebPartContext;
}
export declare const GROUPS: {
    GH_HEADS: string;
    PM_MANAGERS: string;
    ENGG_LEAD: string;
    ENGG_SUB1: string;
    ENGG_SUB2: string;
    ENGG_MGR: string;
    SC_TEAM: string;
    CONTROLLING: string;
};
/** Fired after the active item changes, in case other components care. */
export declare const SIDEBAR_PAGE_CHANGED_EVENT = "dt-sidebar-page-changed";
/** The page key currently highlighted. */
export declare const getActiveSidebarPage: () => string;
/**
 * Highlight the nav item for `page` and clear every other one.
 * Safe to call before the sidebar is in the DOM — the value is
 * remembered and applied by initSidebarActiveState().
 */
export declare const setActiveSidebarPage: (page: string, persist?: boolean) => void;
/**
 * Wire up the highlight. Call once after the sidebar HTML is injected.
 *
 * Uses a capture-phase delegated listener so the highlight is applied
 * even if the host WebPart's own click handler calls stopPropagation()
 * or preventDefault() for routing. Routing behaviour is untouched.
 */
export declare const initSidebarActiveState: (initialPage?: string) => void;
/**
 * Initialize authentication for the sidebar
 * This should be called before any permission checks
 */
export declare const initializeSidebarAuth: (context: WebPartContext) => Promise<void>;
/**
 * Get current user's groups using the shared auth module
 */
export declare const getCurrentUserGroups: () => Promise<void>;
/**
 * Check if user is in a specific group using the shared auth module (case-insensitive)
 */
export declare const isUserInGroup: (groupName: string) => boolean;
/** Convenience role checks */
export declare const isSCTeam: () => boolean;
export declare const isControllingTeam: () => boolean;
export declare const isEnggUser: () => boolean;
export declare const isFullContributeUser: () => boolean;
export declare const isProjectUser: () => boolean;
/**
 * Compute the sidebar pages the current user may SEE.
 * Must be called AFTER authManager.initialize() (i.e. after initializeSidebarAuth).
 * Users in multiple groups get the UNION of all their groups' pages.
 */
export declare const getAllowedPagesSync: () => string[];
/** Safe wrapper that never throws (falls back to cached pages, then dashboard). */
export declare const getAllowedPagesSyncSafe: () => string[];
/** First tab the user is allowed to see — use as landing page after login / for blocked routes. */
export declare const getFirstAllowedPage: () => string;
/**
 * Router guard — call this in your WebPart routing before rendering a page:
 *
 *   import { isPageAllowed, getFirstAllowedPage } from './sidebar';
 *   if (!isPageAllowed(page)) { page = getFirstAllowedPage(); }
 *
 * This blocks direct URL-hash access (e.g. SCM user typing #revenue).
 */
export declare const isPageAllowed: (page: string) => boolean;
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
export declare const canEditPage: (page: string) => boolean;
export declare const getSidebarContent: (context: WebPartContext) => Promise<string>;
export declare const sidebarStyle = "\n@font-face {\n  font-family: 'ABBvoiceOffice';\n  src: url('../../assets/ABBvoiceOffice_Rg.ttf') format('truetype');\n  font-weight: 600;\n  font-style: normal;\n  font-display: swap;\n}\n\n/* Apply ABB font to entire application */\nhtml,\nbody,  \n.page,\n.page * {\n  font-family: 'ABBvoiceOffice';\n}\n\n.dt-sidebar {\n  background: #ffffff;\n  box-shadow: 0 1px 6px -2px #eee, 0 2px 6px 1px #eee, 0 1px 4px 2px #eee;\n  width: 225px;\n  padding: 8px 6px 8px 8px;\n  transition: width 0.3s ease;\n  height: 100%;\n  max-height: 100%;\n  overflow-x: hidden;\n  overflow-y: auto;\n  font-family: ABBvoiceOffice;\n  box-sizing: border-box;\n}\n\n.dt-sidebar.collapsed {\n  width: 60px;\n  padding: 8px 6px;\n  overflow-x: visible;\n  overflow-y: auto;\n}\n\n.dt-sidebar ul {\n  list-style: none;\n  padding: 0;\n  margin: 0;\n}\n\n.dt-sidebar li {\n  margin: 2px 0;\n}\n\n.dt-sidebar a {\n  display: flex;\n  align-items: center;\n  gap: 10px;\n  padding: 8px 12px;\n  color: #4d5f71;\n  text-decoration: none;\n  border-radius: 8px;\n  position: relative;\n  white-space: nowrap;\n  font-weight: 400;\n  transition: background 0.18s ease, color 0.18s ease;\n}\n\n.dt-sidebar a:hover {\n  background: #f0f0ff;\n  color: #696cff;\n}\n\n/* ============================================\n   ACTIVE / SELECTED TAB\n   ============================================ */\n\n.dt-sidebar a.active {\n  background: #e7e7ff;\n  color: #696cff;\n  font-weight: 700;\n}\n\n/* Keep the active item looking selected even while hovering something else */\n.dt-sidebar a.active:hover {\n  background: #dedeff;\n  color: #5a5dff;\n}\n\n/* Left accent bar */\n.dt-sidebar a.active::before {\n  content: '';\n  position: absolute;\n  left: 0;\n  top: 50%;\n  transform: translateY(-50%);\n  width: 3px;\n  height: 60%;\n  background: #696cff;\n  border-radius: 0 3px 3px 0;\n}\n\n/* Tint the inline SVG (CSS beats the stroke=\"#0078d4\" presentation attribute) */\n.dt-sidebar a.active .icon svg {\n  stroke: #696cff;\n}\n\n.dt-sidebar a:hover .icon svg {\n  stroke: #696cff;\n}\n\n/* Collapsed rail: no room for a label, so lean on the pill + accent bar */\n.dt-sidebar.collapsed a.active {\n  background: #e7e7ff;\n}\n\n.dt-sidebar.collapsed a.active::before {\n  left: -6px;\n  height: 70%;\n}\n\n/* Keyboard focus ring */\n.dt-sidebar a:focus-visible {\n  outline: 2px solid #696cff;\n  outline-offset: -2px;\n}\n\n.dt-sidebar.collapsed a {\n  justify-content: center;\n  padding: 8px;\n}\n\n.dt-sidebar.collapsed .icon-label {\n  display: none;\n}\n\n.icon {\n  display: inline-flex;\n  align-items: center;\n  justify-content: center;\n  cursor: pointer;\n  position: relative;\n  flex-shrink: 0;\n  width: 22px;\n  height: 22px;\n}\n\n.icon svg {\n  display: block;\n  width: 20px;\n  height: 20px;\n  transition: stroke 0.18s ease;\n}\n\n/* Collapsed: tooltip to the right so it is not clipped by vertical scroll */\n.dt-sidebar.collapsed .icon[data-tooltip]:hover::after {\n  content: attr(data-tooltip);\n  position: absolute;\n  left: calc(100% + 14px);\n  top: 50%;\n  transform: translateY(-50%);\n  background: #333;\n  color: #fff;\n  font-size: 11px;\n  padding: 4px 8px;\n  border-radius: 4px;\n  white-space: nowrap;\n  z-index: 9999;\n  pointer-events: none;\n}\n\n.dt-sidebar.collapsed .icon[data-tooltip]:hover::before {\n  content: '';\n  position: absolute;\n  left: calc(100% + 4px);\n  top: 50%;\n  transform: translateY(-50%);\n  border-width: 5px;\n  border-style: solid;\n  border-color: transparent #333 transparent transparent;\n  z-index: 9999;\n  pointer-events: none;\n}\n\n/* Thin scrollbar \u2014 only appears when the list overflows */\n.dt-sidebar::-webkit-scrollbar {\n  width: 6px;\n}\n\n.dt-sidebar::-webkit-scrollbar-track {\n  background: transparent;\n}\n\n.dt-sidebar::-webkit-scrollbar-thumb {\n  background: #c5c5d2;\n  border-radius: 8px;\n}\n\n.dt-sidebar::-webkit-scrollbar-thumb:hover {\n  background: #a8a8b8;\n}\n\n.dt-sidebar {\n  scrollbar-width: thin;\n  scrollbar-color: #c5c5d2 transparent;\n}\n";
export declare const sidebarScript: (initialPage?: string) => void;
declare const _default: {
    getContent: (context: WebPartContext) => Promise<string>;
    style: string;
    script: (initialPage?: string | undefined) => void;
    setActivePage: (page: string, persist?: boolean) => void;
    getActivePage: () => string;
    isPageAllowed: (page: string) => boolean;
    getFirstAllowedPage: () => string;
    getAllowedPages: () => string[];
    canEditPage: (page: string) => boolean;
};
export default _default;
//# sourceMappingURL=Sidebar.d.ts.map
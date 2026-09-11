import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IDigitalTransformationApp1WebPartProps {
    page: string;
}
export default class DigitalTransformationApp1WebPart extends BaseClientSideWebPart<IDigitalTransformationApp1WebPartProps> {
    private isAuthInitialized;
    /** 🔧 SINGLE-FLIGHT GUARD (fixes the double-initialisation seen in the console log):
     *  SharePoint / the workbench can call render() more than once (double mount,
     *  property-pane refresh, canvas re-render). Without a guard two async renders
     *  run concurrently → the full auth/sidebar/page sequence logs twice, <style>
     *  tags and module scripts pile up, and listeners get attached twice.
     *  Subsequent render() calls while one is in flight now just await the first. */
    private renderInFlight;
    /** 🔧 Tracks the page currently shown inside the already-mounted shell so a
     *  re-render (property change / canvas remount) only reloads when the
     *  requested page really differs. */
    private activePageKey;
    render(): Promise<void>;
    /** 🔧 Idempotent style injection: same-id style tags are removed before the new
     *  one is appended, so repeated renders never leave duplicate <style> blocks. */
    private appendStyleTag;
    private doRender;
    /** ---------------------------
     *   INITIALIZE AUTH SYSTEM
     * --------------------------- */
    private initializeAuth;
    /** ---------------------------
     *   FULL APP MODE
     *   REMOVE ALL SHAREPOINT UI
     * --------------------------- */
    private hideSharePointUI;
    private applyGlobalStyles;
    /** 🔧 Unified page mount: removes the previous page style, injects the new
     *  content + style, runs the page script inside try/catch and resets the
     *  scroll position of the body so navigation always starts at the top.
     *
     *  🔧 ACTIVE-TAB FIX: we set the active sidebar tab HERE (before injecting
     *  content) because mountPage is the single guaranteed path for every page
     *  open. setActiveSidebarPage clears all other tabs and activates the one
     *  matching the page being opened. persist=false keeps sessionStorage
     *  untouched for internal navigations. */
    private mountPage;
    /** 🔧 DEFAULT/ACTIVE TAB: highlight the sidebar link for the page actually
     *  open (Dashboard by default on first open). Sets the class DIRECTLY on the
     *  links so it works even when the Sidebar module's own highlight pass can't
     *  find them yet (async getSidebarContent / role-based tab rebuild) — retries
     *  briefly until the links exist. On the very first visit it also stores the
     *  default page in sessionStorage so a refresh keeps it highlighted. */
    private highlightSidebarTab;
    /** 🔧 Shared access-denied screen (used by Master + any admin-only page),
     *  with the dashboard button re-bound on every mount. */
    private showAccessDenied;
    private loadPage;
    /** 🔧 Admin-only gate (Master page): Full Control OR an admin group. */
    private canAccessMaster;
    /** 🔧 Used by the sidebar so denied pages never even start to load. */
    private isPageAllowed;
    private checkAndLoadMasterPage;
    private attachSidebarEvents;
    private attachSidebarToggle;
    protected get dataVersion(): Version;
}
//# sourceMappingURL=DigitalTransformationApp1WebPart.d.ts.map
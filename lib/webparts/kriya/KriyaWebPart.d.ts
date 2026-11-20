import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import './components/styles/pageHeaders.css';
export interface IKriyaWebPartProps {
    page: string;
}
export default class KriyaWebPart extends BaseClientSideWebPart<IKriyaWebPartProps> {
    private isAuthInitialized;
    /** 🔧 SINGLE-FLIGHT GUARD: prevents double-initialisation if render()
     *  is called more than once by SharePoint / the workbench. */
    private renderInFlight;
    /** 🔧 Tracks the page currently shown inside the already-mounted shell. */
    private activePageKey;
    render(): Promise<void>;
    /** 🔧 Idempotent style injection. */
    private appendStyleTag;
    private doRender;
    /** 🔧 Full user identity broadcaster.
     *  Sends name + login + email + groups + isAdmin + isGH + canContribute +
     *  permissionLevel so the header can decide whether to show the
     *  "Add User" button (visible only for KRIYA_Admin_Full). */
    private notifyUserReady;
    /** ---------------------------
     *   INITIALIZE AUTH SYSTEM
     * --------------------------- */
    private initializeAuth;
    /** ---------------------------
     *   FULL APP MODE
     * --------------------------- */
    private hideSharePointUI;
    private applyGlobalStyles;
    private mountPage;
    private highlightSidebarTab;
    private showAccessDenied;
    private loadPage;
    private canAccessMaster;
    private isPageAllowed;
    private checkAndLoadMasterPage;
    private attachSidebarEvents;
    private attachSidebarToggle;
    protected get dataVersion(): Version;
}
//# sourceMappingURL=KriyaWebPart.d.ts.map
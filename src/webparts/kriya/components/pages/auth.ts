// auth.ts
// ============================================
// SHARED ROLE-BASED ACCESS CONTROL SYSTEM
// ============================================
//
// CORRECTED PERMISSION MODEL
// ============================================
//   KRIYA_Admin_Full  -> FULL_CONTROL   : view ALL + edit + create + delete
//   KRIYA_GH_Heads    -> CONTRIBUTE     : view OWN (GroupHead1) + edit OWN + create
//   KRIYA_PM_Managers -> CONTRIBUTE     : view OWN (ProjectManager) + edit OWN + create
//   KRIYA_Engg_Lead   -> CONTRIBUTE_OWN : view OWN (EnggLead) + edit OWN (no create)
//   KRIYA_Engg_Sub1   -> CONTRIBUTE_OWN : view OWN (EnggSubLead1) + edit OWN (no create)
//   KRIYA_Engg_Sub2   -> READ           : view OWN (EnggSubLead2)
//   KRIYA_Engg_Mgr    -> READ           : view OWN (EngineeringManagers)
//   KRIYA_SC_Team     -> READ_ALL       : view ALL
//   KRIYA_Controlling -> READ_ALL       : view ALL
//
// "OWN" = the user is the person assigned to that project's role field.
// ============================================

// Declare SharePoint global variable
declare const _spPageContextInfo: any;

// ============================================
// PERMISSION LEVELS DEFINITION
// ============================================

export const PERMISSION_LEVELS = {
  FULL_CONTROL: 'full_control' as const,
  CONTRIBUTE: 'contribute' as const,
  CONTRIBUTE_OWN: 'contribute_own' as const,
  READ_ALL: 'read_all' as const,
  READ: 'read' as const,
  NONE: 'none' as const
} as const;

export type PermissionLevel = typeof PERMISSION_LEVELS[keyof typeof PERMISSION_LEVELS];

// ============================================
// GROUP TO PERMISSION MAPPING
// ============================================

export const GROUP_PERMISSIONS: Record<string, PermissionLevel> = {
  'KRIYA_Admin_Full': PERMISSION_LEVELS.FULL_CONTROL,
  'KRIYA_GH_Heads': PERMISSION_LEVELS.CONTRIBUTE,
  'KRIYA_PM_Managers': PERMISSION_LEVELS.CONTRIBUTE,
  'KRIYA_Engg_Lead': PERMISSION_LEVELS.CONTRIBUTE_OWN,
  'KRIYA_Engg_Sub1': PERMISSION_LEVELS.CONTRIBUTE_OWN,
  'KRIYA_Engg_Sub2': PERMISSION_LEVELS.CONTRIBUTE_OWN,
  'KRIYA_Engg_Mgr': PERMISSION_LEVELS.CONTRIBUTE_OWN,
  'KRIYA_SC_Team': PERMISSION_LEVELS.READ_ALL,
  'KRIYA_Controlling': PERMISSION_LEVELS.READ_ALL,
};

// ============================================
// ROLE → SHAREPOINT GROUP NAME MAP
// ============================================
// Single source of truth for the SharePoint group titles used to populate
// the role-based people-picker dropdowns (Group Head, Project Manager,
// Engg Lead, Engg Sub Leads, Engineering Managers, SCM, Controlling).
//
// The values MUST match the actual SharePoint group titles exactly.
export const ROLE_GROUP_NAMES = {
  GroupHead1: 'KRIYA_GH_Heads',
  ProjectManager: 'KRIYA_PM_Managers',
  EnggLead: 'KRIYA_Engg_Lead',
  EnggSubLead1: 'KRIYA_Engg_Sub1',
  EnggSubLead2: 'KRIYA_Engg_Sub2',
  EngineeringManagers: 'KRIYA_Engg_Mgr',
  SCM: 'KRIYA_SC_Team',
  Controlling: 'KRIYA_Controlling'
} as const;

export type RoleGroupKey = keyof typeof ROLE_GROUP_NAMES;

// ============================================
// ROLE -> FIELD(S) THAT DEFINE A USER'S OWN PROJECTS
// ============================================
// A project is "owned" / "assigned" to a user when the user's name is the
// value of one of these fields on that project. Edit this map to tweak the
// meaning of "Own Assigned Project" per role (a role may own several fields).
export const ROLE_OWN_FIELDS: Record<string, string[]> = {
  'KRIYA_GH_Heads': ['GroupHead1'],
  'KRIYA_PM_Managers': ['ProjectManager'],
  'KRIYA_Engg_Lead': ['EnggLead'],
  'KRIYA_Engg_Sub1': ['EnggSubLead1'],
  'KRIYA_Engg_Sub2': ['EnggSubLead2'],
  'KRIYA_Engg_Mgr': ['EngineeringManagers'],
};

// Engineering groups (used for isEngineeringUser())
const ENGINEERING_GROUPS = [
  'KRIYA_Engg_Lead',
  'KRIYA_Engg_Sub1',
  'KRIYA_Engg_Sub2',
  'KRIYA_Engg_Mgr',
];

// ============================================
// INTERFACES
// ============================================

export interface AuthConfig {
  siteUrl: string;
  listName?: string;
  projectListName?: string;
  isCollection?: boolean;
}

export interface UserContext {
  userGroups: string[];
  userPermissionLevel: PermissionLevel;
  currentUserFullName: string;
  currentUserLogin: string;
  isSiteAdmin: boolean;
}

// ============================================
// AUTH MANAGER CLASS
// ============================================

class AuthManager {
  private static instance: AuthManager;
  private context: UserContext;
  private _config: AuthConfig | null;
  private _initialized: boolean;

  private constructor() {
    this.context = {
      userGroups: [],
      userPermissionLevel: PERMISSION_LEVELS.NONE,
      currentUserFullName: '',
      currentUserLogin: '',
      isSiteAdmin: false
    };
    this._config = null;
    this._initialized = false;
  }

  // ============================================
  // SINGLETON INSTANCE
  // ============================================

  public static getInstance(): AuthManager {
    if (!AuthManager.instance) {
      AuthManager.instance = new AuthManager();
    }
    return AuthManager.instance;
  }

  // ============================================
  // INITIALIZATION
  // ============================================

  public async initialize(config: AuthConfig): Promise<void> {
    if (this._initialized) {
      console.log('🔑 Auth already initialized');
      return;
    }

    this._config = config;

    try {
      // Get current user login
      const currentUserLogin = typeof _spPageContextInfo !== 'undefined'
        ? _spPageContextInfo.userLoginName
        : '';

      // Get user full name
      const currentUserFullName = await this.fetchCurrentUserFullName(config.siteUrl);
      console.log(`✅ Current user full name from SharePoint: "${currentUserFullName}"`);

      // Get user groups
      const userGroups = await this.fetchUserGroups(config.siteUrl);

      // Check site admin
      const isSiteAdmin = typeof _spPageContextInfo !== 'undefined'
        ? _spPageContextInfo.isSiteAdmin
        : false;

      // Determine permission level
      const userPermissionLevel = this.determinePermissionLevel(userGroups, isSiteAdmin);

      // Update context
      this.context = {
        userGroups,
        userPermissionLevel,
        currentUserFullName,
        currentUserLogin,
        isSiteAdmin
      };

      this._initialized = true;

      console.log(`👤 User: ${currentUserFullName} (${currentUserLogin})`);
      console.log(`📋 User Groups:`, userGroups);
      console.log(`🔑 Permission: ${userPermissionLevel.toUpperCase()}`);
    } catch (error) {
      console.error('Error initializing authentication:', error);
      this.context.userPermissionLevel = PERMISSION_LEVELS.NONE;
      this._initialized = true;
    }
  }

  // ============================================
  // PRIVATE HELPER METHODS
  // ============================================

  private async fetchCurrentUserFullName(siteUrl: string): Promise<string> {
    try {
      const response = await fetch(`${siteUrl}/_api/web/currentuser`, {
        headers: { Accept: 'application/json;odata=verbose' }
      });
      const data = await response.json();
      return data.d?.Title || '';
    } catch (error) {
      console.error('Error getting user name:', error);
      return '';
    }
  }

  private async fetchUserGroups(siteUrl: string): Promise<string[]> {
    try {
      const response = await fetch(`${siteUrl}/_api/web/currentuser/groups`, {
        headers: { Accept: 'application/json;odata=verbose' }
      });
      const data = await response.json();
      const groups = data.d?.results || [];
      return groups.map((group: any) => group.Title || group.LoginName);
    } catch (error) {
      console.error('Error getting user groups:', error);
      return [];
    }
  }

  // Highest privilege wins. Order: FULL_CONTROL > CONTRIBUTE > CONTRIBUTE_OWN > READ_ALL > READ > NONE
  private determinePermissionLevel(
    userGroups: string[],
    isSiteAdmin: boolean
  ): PermissionLevel {
    // Site admins get full control
    if (isSiteAdmin) {
      return PERMISSION_LEVELS.FULL_CONTROL;
    }

    let highestLevel: PermissionLevel = PERMISSION_LEVELS.NONE;

    for (const group of userGroups) {
      const permission = GROUP_PERMISSIONS[group];

      if (!permission) continue;

      // FULL_CONTROL has highest priority - return immediately
      if (permission === PERMISSION_LEVELS.FULL_CONTROL) {
        return PERMISSION_LEVELS.FULL_CONTROL;
      }

      // CONTRIBUTE - second highest priority
      if (permission === PERMISSION_LEVELS.CONTRIBUTE) {
        highestLevel = PERMISSION_LEVELS.CONTRIBUTE;
        continue;
      }

      // CONTRIBUTE_OWN - third highest priority (only if we haven't set CONTRIBUTE)
      if (permission === PERMISSION_LEVELS.CONTRIBUTE_OWN &&
          highestLevel !== PERMISSION_LEVELS.CONTRIBUTE) {
        highestLevel = PERMISSION_LEVELS.CONTRIBUTE_OWN;
        continue;
      }

      // READ_ALL - fourth highest (only if we haven't set any edit level)
      if (permission === PERMISSION_LEVELS.READ_ALL &&
          highestLevel !== PERMISSION_LEVELS.CONTRIBUTE &&
          highestLevel !== PERMISSION_LEVELS.CONTRIBUTE_OWN) {
        highestLevel = PERMISSION_LEVELS.READ_ALL;
        continue;
      }

      // READ - lowest priority (only if we haven't set anything higher)
      if (permission === PERMISSION_LEVELS.READ &&
          highestLevel === PERMISSION_LEVELS.NONE) {
        highestLevel = PERMISSION_LEVELS.READ;
      }
    }

    return highestLevel;
  }

  // The field(s) that define "own assigned" projects for the current user,
  // derived from the user's group memberships.
  private getOwnedRoleFields(): string[] {
    const fields = new Set<string>();
    for (const group of this.context.userGroups) {
      const own = ROLE_OWN_FIELDS[group];
      if (own) {
        own.forEach(field => fields.add(field));
      }
    }
    return Array.from(fields);
  }

  // Extract a display name from a Person/Group value (object or string).
  private getFieldValueName(value: any): string {
    if (!value) return '';
    if (typeof value === 'object' && value !== null) {
      return value.Title || value.Email || '';
    }
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return String(value);
    return '';
  }

  // True if the current user is the assigned person for any of the given fields.
  private projectMatchesUser(project: any, fields: string[]): boolean {
    if (!project) return false;
    const userName = (this.context.currentUserFullName || '').trim().toLowerCase();
    if (!userName) return false;

    for (const field of fields) {
      const name = this.getFieldValueName(project[field]);
      if (name && name.trim().toLowerCase() === userName) {
        return true;
      }
    }
    return false;
  }

  // ============================================
  // PUBLIC PERMISSION CHECK METHODS
  // ============================================

  public hasFullControl(): boolean {
    return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL;
  }

  // Can create new records (Add New / Upload / Import).
  public canCreate(): boolean {
    return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL ||
           this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE;
  }

  // Can edit at least the projects assigned to them.
  public canEditOwn(): boolean {
    return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL ||
           this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE ||
           this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE_OWN;
  }

  // Back-compat alias: "contribute" = create rights.
  public canContribute(): boolean {
    return this.canCreate();
  }

  // Can delete records (Admin only).
  public canDelete(): boolean {
    return this.hasFullControl();
  }

  public canRead(): boolean {
    return this.context.userPermissionLevel !== PERMISSION_LEVELS.NONE;
  }

  public canReadAllProjects(): boolean {
    return this.context.userPermissionLevel === PERMISSION_LEVELS.READ_ALL ||
           this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL;
  }

  public canFreeze(): boolean {
    return this.hasFullControl();
  }

  public isGHHead(): boolean {
    return this.context.userGroups.some(group => group === 'KRIYA_GH_Heads');
  }

  public isPMManager(): boolean {
    return this.context.userGroups.some(group => group === 'KRIYA_PM_Managers');
  }

  public isEngineeringUser(): boolean {
    return this.context.userGroups.some(group => ENGINEERING_GROUPS.includes(group));
  }

  public getUserGroups(): string[] {
    return [...this.context.userGroups];
  }

  public getCurrentUserFullName(): string {
    return this.context.currentUserFullName;
  }

  public getCurrentUserLogin(): string {
    return this.context.currentUserLogin;
  }

  public getPermissionLevel(): PermissionLevel {
    return this.context.userPermissionLevel;
  }

  public isInitialized(): boolean {
    return this._initialized;
  }

  public getConfig(): AuthConfig | null {
    return this._config;
  }

  // ============================================
  // PROJECT-BASED PERMISSION CHECKS
  // ============================================

  public canEditProject(
    project?: any
  ): boolean {
    // Full Control users can edit everything
    if (this.hasFullControl()) return true;

    // Only edit-capable roles can edit (CONTRIBUTE / CONTRIBUTE_OWN)
    if (!this.canEditOwn()) return false;

    // Otherwise they can only edit their OWN assigned projects.
    return this.projectMatchesUser(project, this.getOwnedRoleFields());
  }

  public canCreateRecord(): boolean {
    return this.canCreate();
  }

   public canCreateForProject(project?: any): boolean {
    if (this.hasFullControl()) return true;

    const level = this.context.userPermissionLevel;

    if (level === PERMISSION_LEVELS.CONTRIBUTE) return true;

    if (level === PERMISSION_LEVELS.CONTRIBUTE_OWN) {
      if (!project) return false;
      return this.projectMatchesUser(project, this.getOwnedRoleFields());
    }

    return false;
  }

  public canViewForm(
    isEdit: boolean = false,
    project?: any
  ): boolean {
    if (!this.canRead()) return false;

    if (isEdit) {
      return this.canEditProject(project);
    }

    return this.canCreate();
  }

  // ============================================
  // GET ASSIGNED PROJECTS
  // ============================================

  public getAssignedProjects<T extends {
    Id: number;
    Title: string;
    ProjectManager?: string | { Title: string; Email?: string };
    EnggLead?: string | { Title: string };
    EnggSubLead1?: string | { Title: string };
    EnggSubLead2?: string | { Title: string };
    EngineeringManagers?: string | { Title: string };
    GroupHead1?: string | { Title: string };
  }>(
    allProjects: T[]
  ): T[] {
    // ✅ Admin (FULL_CONTROL) sees all projects
    if (this.hasFullControl()) {
      console.log('✅ Full Control - showing all projects');
      return allProjects;
    }

    // ✅ READ_ALL users (SC_Team, Controlling) see all projects read-only
    if (this.canReadAllProjects()) {
      console.log('✅ Read All - showing all projects (read-only)');
      return allProjects;
    }

    const fields = this.getOwnedRoleFields();
    const userFullName = this.context.currentUserFullName.trim();
    console.log(`🔍 Checking projects for user: ${userFullName} (own fields: ${fields.join(', ')})`);

    if (!fields.length) {
      console.log(`⚠️ User ${userFullName} has no owned role fields - showing no projects`);
      return [];
    }

    const assigned = allProjects.filter(project => {
      for (const field of fields) {
        const managerData = project[field as keyof typeof project];
        if (!managerData) continue;

        let managerName = '';
        if (typeof managerData === 'object' && managerData !== null) {
          managerName = (managerData as any).Title || '';
        } else if (typeof managerData === 'string') {
          managerName = managerData;
        }

        if (managerName && managerName.toLowerCase().trim() === userFullName.toLowerCase().trim()) {
          console.log(`✅ User matches ${field}: ${managerName}`);
          return true;
        }
      }
      return false;
    });

    console.log(`📋 User ${userFullName} found ${assigned.length} assigned project(s) out of ${allProjects.length}`);
    return assigned;
  }

  public getAssignedProjectValues<T extends {
    Id: number;
    Title: string;
    ProjectManager?: string | { Title: string };
    [key: string]: any;
  }>(
    allProjects: T[],
    key: string
  ): string[] {
    const assigned = this.getAssignedProjects(allProjects);
    const values = new Set<string>();

    for (const project of assigned) {
      let value = project[key];

      if (!value) continue;

      // Handle manager fields with object values
      if (typeof value === 'object' && value.Title) {
        value = value.Title;
      }

      if (typeof value === 'string' && value.trim()) {
        values.add(value.trim());
      }
    }

    return Array.from(values).sort();
  }

  // ============================================
  // UI UPDATE METHODS
  // ============================================

  public updateUIBasedOnPermissions(selectors: {
  importBtn?: string;
  exportBtn?: string;
  addBtn?: string;
  saveBtn?: string;
  freezeCheckboxes?: string;
} = {}): void {
  // ✅ Role-level capabilities
  const canCreateAny = this.canCreateAny();  // show "+ New" for CONTRIBUTE_OWN too
  const canImport    = this.canCreate();     // bulk import: true create rights only
  const canExport    = this.canRead();
  const canFreeze    = this.canFreeze();
  const canWrite     = this.canCreate() || this.canEditOwn(); // Save button
  const isReadOnly   =
    this.context.userPermissionLevel === PERMISSION_LEVELS.READ ||
    this.context.userPermissionLevel === PERMISSION_LEVELS.READ_ALL;

  // Import button — global create only
  if (selectors.importBtn) {
    const button = document.querySelector(selectors.importBtn) as HTMLElement;
    if (button) {
      button.style.display = canImport ? '' : 'none';
    }
  }

  // Export button — read access
  if (selectors.exportBtn) {
    const button = document.querySelector(selectors.exportBtn) as HTMLElement;
    if (button) {
      button.style.display = canExport ? '' : 'none';
    }
  }

  // ✅ Add ("+ New") button — any role that can create for at least one project
  if (selectors.addBtn) {
    const button = document.querySelector(selectors.addBtn) as HTMLElement;
    if (button) {
      button.style.display = canCreateAny ? '' : 'none';
    }
  }

  // Save button — create OR edit-own
  if (selectors.saveBtn) {
    const button = document.querySelector(selectors.saveBtn) as HTMLElement;
    if (button) {
      button.style.display = canWrite ? '' : 'none';
    }
  }

  // Freeze checkboxes — admin only
  if (selectors.freezeCheckboxes) {
    document.querySelectorAll(selectors.freezeCheckboxes).forEach(element => {
      const checkbox = element as HTMLInputElement;
      checkbox.disabled = !canFreeze;
      checkbox.style.pointerEvents = canFreeze ? 'auto' : 'none';
    });
  }

  // Read-only indicator
  if (isReadOnly) {
    const header = document.querySelector('.page-header h1') as HTMLElement;
    if (header && !header.textContent?.includes('(Read Only)')) {
      header.textContent += ' (Read Only)';
    }
  }
}

  // ============================================
  // SHAREPOINT UTILITIES
  // ============================================

  public async getDigest(siteUrl: string): Promise<string> {
    const response = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: { Accept: 'application/json;odata=verbose' }
    });
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  public async getListEntityType(siteUrl: string, listName: string): Promise<string> {
    try {
      const response = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        {
          headers: { Accept: 'application/json;odata=verbose' }
        }
      );
      const data = await response.json();
      return data.d.ListItemEntityTypeFullName;
    } catch (error) {
      console.error('Error getting entity type:', error);
      return `SP.Data.${listName.replace(/ /g, '_x0020_')}ListItem`;
    }
  }

  // ============================================
  // RESET AUTH (for testing)
  // ============================================

  public reset(): void {
    this._initialized = false;
    this.context = {
      userGroups: [],
      userPermissionLevel: PERMISSION_LEVELS.NONE,
      currentUserFullName: '',
      currentUserLogin: '',
      isSiteAdmin: false
    };
    this._config = null;
  }

  // Can create for at least one project (the actual project check happens
// later, per-selected-project, via canCreateForProject).
public canCreateAny(): boolean {
  return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL ||
         this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE ||
         this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE_OWN;
}
}

// ============================================
// EXPORT SINGLETON INSTANCE
// ============================================

export const authManager = AuthManager.getInstance();

// ============================================
// CONVENIENCE EXPORT FUNCTIONS
// ============================================

export const getAuth = (): AuthManager => authManager;

// Permission check functions
export const hasFullControl = (): boolean => authManager.hasFullControl();
export const canCreate = (): boolean => authManager.canCreate();
export const canEditOwn = (): boolean => authManager.canEditOwn();
export const canContribute = (): boolean => authManager.canContribute();
export const canDelete = (): boolean => authManager.canDelete();
export const canRead = (): boolean => authManager.canRead();
export const canReadAllProjects = (): boolean => authManager.canReadAllProjects();
export const canFreeze = (): boolean => authManager.canFreeze();
export const isGHHead = (): boolean => authManager.isGHHead();
export const isPMManager = (): boolean => authManager.isPMManager();
export const isEngineeringUser = (): boolean => authManager.isEngineeringUser();

// User info functions
export const getUserGroups = (): string[] => authManager.getUserGroups();
export const getCurrentUserFullName = (): string => authManager.getCurrentUserFullName();
export const getCurrentUserLogin = (): string => authManager.getCurrentUserLogin();
export const getPermissionLevel = (): PermissionLevel => authManager.getPermissionLevel();

// Project-based permission functions
export const canEditProject = (
  project?: any
): boolean => authManager.canEditProject(project);

export const canCreateRecord = (): boolean => authManager.canCreateRecord();
export const canCreateForProject = (project?: any): boolean =>
  authManager.canCreateForProject(project);

export const canViewForm = (
  isEdit: boolean = false,
  project?: any
): boolean => authManager.canViewForm(isEdit, project);

// Project filtering functions
export const getAssignedProjects = <T extends {
  Id: number;
  Title: string;
  ProjectManager?: string | { Title: string };
  EnggLead?: string | { Title: string };
  EnggSubLead1?: string | { Title: string };
  EnggSubLead2?: string | { Title: string };
  EngineeringManagers?: string | { Title: string };
  GroupHead1?: string | { Title: string };
}>(
  allProjects: T[]
): T[] => authManager.getAssignedProjects(allProjects);

export const getAssignedProjectValues = <T extends {
  Id: number;
  Title: string;
  ProjectManager?: string | { Title: string };
  [key: string]: any;
}>(
  allProjects: T[],
  key: string
): string[] => authManager.getAssignedProjectValues(allProjects, key);

// UI update function
export const updateUIBasedOnPermissions = (selectors?: {
  importBtn?: string;
  exportBtn?: string;
  addBtn?: string;
  saveBtn?: string;
  freezeCheckboxes?: string;
}): void => authManager.updateUIBasedOnPermissions(selectors);

// SharePoint utility functions
export const getDigest = (siteUrl: string): Promise<string> =>
  authManager.getDigest(siteUrl);

export const getListEntityType = (siteUrl: string, listName: string): Promise<string> =>
  authManager.getListEntityType(siteUrl, listName);
export const canCreateAny = (): boolean => authManager.canCreateAny();

// ============================================
// DEFAULT EXPORT
// ============================================

export default {
  authManager,     
  PERMISSION_LEVELS,
  GROUP_PERMISSIONS,
  ROLE_GROUP_NAMES,
  ROLE_OWN_FIELDS,
  hasFullControl,
  canCreate,
  canEditOwn,
  canContribute,
  canDelete,
  canRead,
  canReadAllProjects,
  canFreeze,
  isGHHead,
  isPMManager,
  isEngineeringUser,
  getUserGroups,
  getCurrentUserFullName,
  getCurrentUserLogin,
  getPermissionLevel,
  canEditProject,
  canCreateRecord,
  canCreateForProject,   // ← added
  canViewForm,
  getAssignedProjects,
  getAssignedProjectValues,
  updateUIBasedOnPermissions,
  getDigest,
  getListEntityType
};
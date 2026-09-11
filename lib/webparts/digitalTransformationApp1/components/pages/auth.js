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
// ============================================
// PERMISSION LEVELS DEFINITION
// ============================================
export const PERMISSION_LEVELS = {
    FULL_CONTROL: 'full_control',
    CONTRIBUTE: 'contribute',
    CONTRIBUTE_OWN: 'contribute_own',
    READ_ALL: 'read_all',
    READ: 'read',
    NONE: 'none'
};
// ============================================
// GROUP TO PERMISSION MAPPING
// ============================================
export const GROUP_PERMISSIONS = {
    'KRIYA_Admin_Full': PERMISSION_LEVELS.FULL_CONTROL,
    'KRIYA_GH_Heads': PERMISSION_LEVELS.CONTRIBUTE,
    'KRIYA_PM_Managers': PERMISSION_LEVELS.CONTRIBUTE,
    'KRIYA_Engg_Lead': PERMISSION_LEVELS.CONTRIBUTE_OWN,
    'KRIYA_Engg_Sub1': PERMISSION_LEVELS.CONTRIBUTE_OWN,
    'KRIYA_Engg_Sub2': PERMISSION_LEVELS.READ,
    'KRIYA_Engg_Mgr': PERMISSION_LEVELS.READ,
    'KRIYA_SC_Team': PERMISSION_LEVELS.READ_ALL,
    'KRIYA_Controlling': PERMISSION_LEVELS.READ_ALL,
};
// ============================================
// ROLE -> FIELD(S) THAT DEFINE A USER'S OWN PROJECTS
// ============================================
// A project is "owned" / "assigned" to a user when the user's name is the
// value of one of these fields on that project. Edit this map to tweak the
// meaning of "Own Assigned Project" per role (a role may own several fields).
export const ROLE_OWN_FIELDS = {
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
// AUTH MANAGER CLASS
// ============================================
class AuthManager {
    constructor() {
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
    static getInstance() {
        if (!AuthManager.instance) {
            AuthManager.instance = new AuthManager();
        }
        return AuthManager.instance;
    }
    // ============================================
    // INITIALIZATION
    // ============================================
    async initialize(config) {
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
        }
        catch (error) {
            console.error('Error initializing authentication:', error);
            this.context.userPermissionLevel = PERMISSION_LEVELS.NONE;
            this._initialized = true;
        }
    }
    // ============================================
    // PRIVATE HELPER METHODS
    // ============================================
    async fetchCurrentUserFullName(siteUrl) {
        var _a;
        try {
            const response = await fetch(`${siteUrl}/_api/web/currentuser`, {
                headers: { Accept: 'application/json;odata=verbose' }
            });
            const data = await response.json();
            return ((_a = data.d) === null || _a === void 0 ? void 0 : _a.Title) || '';
        }
        catch (error) {
            console.error('Error getting user name:', error);
            return '';
        }
    }
    async fetchUserGroups(siteUrl) {
        var _a;
        try {
            const response = await fetch(`${siteUrl}/_api/web/currentuser/groups`, {
                headers: { Accept: 'application/json;odata=verbose' }
            });
            const data = await response.json();
            const groups = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
            return groups.map((group) => group.Title || group.LoginName);
        }
        catch (error) {
            console.error('Error getting user groups:', error);
            return [];
        }
    }
    // Highest privilege wins. Order: FULL_CONTROL > CONTRIBUTE > CONTRIBUTE_OWN > READ_ALL > READ > NONE
    determinePermissionLevel(userGroups, isSiteAdmin) {
        // Site admins get full control
        if (isSiteAdmin) {
            return PERMISSION_LEVELS.FULL_CONTROL;
        }
        let highestLevel = PERMISSION_LEVELS.NONE;
        for (const group of userGroups) {
            const permission = GROUP_PERMISSIONS[group];
            if (!permission)
                continue;
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
    getOwnedRoleFields() {
        const fields = new Set();
        for (const group of this.context.userGroups) {
            const own = ROLE_OWN_FIELDS[group];
            if (own) {
                own.forEach(field => fields.add(field));
            }
        }
        return Array.from(fields);
    }
    // Extract a display name from a Person/Group value (object or string).
    getFieldValueName(value) {
        if (!value)
            return '';
        if (typeof value === 'object' && value !== null) {
            return value.Title || value.Email || '';
        }
        if (typeof value === 'string')
            return value;
        if (typeof value === 'number')
            return String(value);
        return '';
    }
    // True if the current user is the assigned person for any of the given fields.
    projectMatchesUser(project, fields) {
        if (!project)
            return false;
        const userName = (this.context.currentUserFullName || '').trim().toLowerCase();
        if (!userName)
            return false;
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
    hasFullControl() {
        return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL;
    }
    // Can create new records (Add New / Upload / Import).
    canCreate() {
        return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL ||
            this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE;
    }
    // Can edit at least the projects assigned to them.
    canEditOwn() {
        return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL ||
            this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE ||
            this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE_OWN;
    }
    // Back-compat alias: "contribute" = create rights.
    canContribute() {
        return this.canCreate();
    }
    // Can delete records (Admin only).
    canDelete() {
        return this.hasFullControl();
    }
    canRead() {
        return this.context.userPermissionLevel !== PERMISSION_LEVELS.NONE;
    }
    canReadAllProjects() {
        return this.context.userPermissionLevel === PERMISSION_LEVELS.READ_ALL ||
            this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL;
    }
    canFreeze() {
        return this.hasFullControl();
    }
    isGHHead() {
        return this.context.userGroups.some(group => group === 'KRIYA_GH_Heads');
    }
    isPMManager() {
        return this.context.userGroups.some(group => group === 'KRIYA_PM_Managers');
    }
    isEngineeringUser() {
        return this.context.userGroups.some(group => ENGINEERING_GROUPS.includes(group));
    }
    getUserGroups() {
        return [...this.context.userGroups];
    }
    getCurrentUserFullName() {
        return this.context.currentUserFullName;
    }
    getCurrentUserLogin() {
        return this.context.currentUserLogin;
    }
    getPermissionLevel() {
        return this.context.userPermissionLevel;
    }
    isInitialized() {
        return this._initialized;
    }
    getConfig() {
        return this._config;
    }
    // ============================================
    // PROJECT-BASED PERMISSION CHECKS
    // ============================================
    canEditProject(project) {
        // Full Control users can edit everything
        if (this.hasFullControl())
            return true;
        // Only edit-capable roles can edit (CONTRIBUTE / CONTRIBUTE_OWN)
        if (!this.canEditOwn())
            return false;
        // Otherwise they can only edit their OWN assigned projects.
        return this.projectMatchesUser(project, this.getOwnedRoleFields());
    }
    canCreateRecord() {
        return this.canCreate();
    }
    canViewForm(isEdit = false, project) {
        if (!this.canRead())
            return false;
        if (isEdit) {
            return this.canEditProject(project);
        }
        return this.canCreate();
    }
    // ============================================
    // GET ASSIGNED PROJECTS
    // ============================================
    getAssignedProjects(allProjects) {
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
                const managerData = project[field];
                if (!managerData)
                    continue;
                let managerName = '';
                if (typeof managerData === 'object' && managerData !== null) {
                    managerName = managerData.Title || '';
                }
                else if (typeof managerData === 'string') {
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
    getAssignedProjectValues(allProjects, key) {
        const assigned = this.getAssignedProjects(allProjects);
        const values = new Set();
        for (const project of assigned) {
            let value = project[key];
            if (!value)
                continue;
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
    updateUIBasedOnPermissions(selectors = {}) {
        var _a;
        const canCreate = this.canCreate();
        const canExport = this.canRead();
        const canFreeze = this.canFreeze();
        const canWrite = this.canCreate() || this.canEditOwn();
        const isReadOnly = this.context.userPermissionLevel === PERMISSION_LEVELS.READ ||
            this.context.userPermissionLevel === PERMISSION_LEVELS.READ_ALL;
        // Import button
        if (selectors.importBtn) {
            const button = document.querySelector(selectors.importBtn);
            if (button) {
                button.style.display = canCreate ? '' : 'none';
            }
        }
        // Export button
        if (selectors.exportBtn) {
            const button = document.querySelector(selectors.exportBtn);
            if (button) {
                button.style.display = canExport ? '' : 'none';
            }
        }
        // Add button
        if (selectors.addBtn) {
            const button = document.querySelector(selectors.addBtn);
            if (button) {
                button.style.display = canCreate ? '' : 'none';
            }
        }
        // Save button (visible when the user can create OR edit own projects)
        if (selectors.saveBtn) {
            const button = document.querySelector(selectors.saveBtn);
            if (button) {
                button.style.display = canWrite ? '' : 'none';
            }
        }
        // Freeze checkboxes
        if (selectors.freezeCheckboxes) {
            document.querySelectorAll(selectors.freezeCheckboxes).forEach(element => {
                const checkbox = element;
                checkbox.disabled = !canFreeze;
                checkbox.style.pointerEvents = canFreeze ? 'auto' : 'none';
            });
        }
        // Show read-only indicator if needed
        if (isReadOnly) {
            const header = document.querySelector('.page-header h1');
            if (header && !((_a = header.textContent) === null || _a === void 0 ? void 0 : _a.includes('(Read Only)'))) {
                header.textContent += ' (Read Only)';
            }
        }
    }
    // ============================================
    // SHAREPOINT UTILITIES
    // ============================================
    async getDigest(siteUrl) {
        const response = await fetch(`${siteUrl}/_api/contextinfo`, {
            method: 'POST',
            headers: { Accept: 'application/json;odata=verbose' }
        });
        const data = await response.json();
        return data.d.GetContextWebInformation.FormDigestValue;
    }
    async getListEntityType(siteUrl, listName) {
        try {
            const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                headers: { Accept: 'application/json;odata=verbose' }
            });
            const data = await response.json();
            return data.d.ListItemEntityTypeFullName;
        }
        catch (error) {
            console.error('Error getting entity type:', error);
            return `SP.Data.${listName.replace(/ /g, '_x0020_')}ListItem`;
        }
    }
    // ============================================
    // RESET AUTH (for testing)
    // ============================================
    reset() {
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
}
// ============================================
// EXPORT SINGLETON INSTANCE
// ============================================
export const authManager = AuthManager.getInstance();
// ============================================
// CONVENIENCE EXPORT FUNCTIONS
// ============================================
export const getAuth = () => authManager;
// Permission check functions
export const hasFullControl = () => authManager.hasFullControl();
export const canCreate = () => authManager.canCreate();
export const canEditOwn = () => authManager.canEditOwn();
export const canContribute = () => authManager.canContribute();
export const canDelete = () => authManager.canDelete();
export const canRead = () => authManager.canRead();
export const canReadAllProjects = () => authManager.canReadAllProjects();
export const canFreeze = () => authManager.canFreeze();
export const isGHHead = () => authManager.isGHHead();
export const isPMManager = () => authManager.isPMManager();
export const isEngineeringUser = () => authManager.isEngineeringUser();
// User info functions
export const getUserGroups = () => authManager.getUserGroups();
export const getCurrentUserFullName = () => authManager.getCurrentUserFullName();
export const getCurrentUserLogin = () => authManager.getCurrentUserLogin();
export const getPermissionLevel = () => authManager.getPermissionLevel();
// Project-based permission functions
export const canEditProject = (project) => authManager.canEditProject(project);
export const canCreateRecord = () => authManager.canCreateRecord();
export const canViewForm = (isEdit = false, project) => authManager.canViewForm(isEdit, project);
// Project filtering functions
export const getAssignedProjects = (allProjects) => authManager.getAssignedProjects(allProjects);
export const getAssignedProjectValues = (allProjects, key) => authManager.getAssignedProjectValues(allProjects, key);
// UI update function
export const updateUIBasedOnPermissions = (selectors) => authManager.updateUIBasedOnPermissions(selectors);
// SharePoint utility functions
export const getDigest = (siteUrl) => authManager.getDigest(siteUrl);
export const getListEntityType = (siteUrl, listName) => authManager.getListEntityType(siteUrl, listName);
// ============================================
// DEFAULT EXPORT
// ============================================
export default {
    authManager,
    PERMISSION_LEVELS,
    GROUP_PERMISSIONS,
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
    canViewForm,
    getAssignedProjects,
    getAssignedProjectValues,
    updateUIBasedOnPermissions,
    getDigest,
    getListEntityType
};
//# sourceMappingURL=auth.js.map
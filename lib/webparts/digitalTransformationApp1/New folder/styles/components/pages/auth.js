// auth.ts
// ============================================
// SHARED ROLE-BASED ACCESS CONTROL SYSTEM
// ============================================
// ============================================
// PERMISSION LEVELS DEFINITION
// ============================================
export const PERMISSION_LEVELS = {
    FULL_CONTROL: 'full_control',
    CONTRIBUTE: 'contribute',
    READ: 'read',
    READ_ALL: 'read_all',
    NONE: 'none'
};
// ============================================
// GROUP TO PERMISSION MAPPING
// ============================================
export const GROUP_PERMISSIONS = {
    'KRIYA_Admin_Full': PERMISSION_LEVELS.FULL_CONTROL,
    'KRIYA_GH_Heads': PERMISSION_LEVELS.CONTRIBUTE,
    'KRIYA_PM_Managers': PERMISSION_LEVELS.CONTRIBUTE,
    'KRIYA_Engg_Lead': PERMISSION_LEVELS.READ,
    'KRIYA_Engg_Sub1': PERMISSION_LEVELS.READ,
    'KRIYA_Engg_Sub2': PERMISSION_LEVELS.READ,
    'KRIYA_Engg_Mgr': PERMISSION_LEVELS.READ,
    'KRIYA_SC_Team': PERMISSION_LEVELS.READ_ALL,
    'KRIYA_Controlling': PERMISSION_LEVELS.READ_ALL,
};
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
            const currentUser = typeof _spPageContextInfo !== 'undefined'
                ? _spPageContextInfo.userLoginName
                : '';
            // Get user full name
            const currentUserFullName = await this.fetchCurrentUserFullName(config.siteUrl);
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
                currentUserLogin: currentUser,
                isSiteAdmin
            };
            this._initialized = true;
            console.log(`👤 User: ${currentUserFullName} (${currentUser})`);
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
            // READ_ALL - third highest priority (only if we haven't set CONTRIBUTE)
            if (permission === PERMISSION_LEVELS.READ_ALL &&
                highestLevel !== PERMISSION_LEVELS.CONTRIBUTE) {
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
    isInContributeGroup(userGroups) {
        const contributeGroups = [
            'KRIYA_GH_Heads',
            'KRIYA_PM_Managers'
        ];
        return userGroups.some(group => contributeGroups.includes(group));
    }
    isInReadGroup(userGroups) {
        const readGroups = [
            'KRIYA_Engg_Lead',
            'KRIYA_Engg_Sub1',
            'KRIYA_Engg_Sub2',
            'KRIYA_Engg_Mgr'
        ];
        return userGroups.some(group => readGroups.includes(group));
    }
    getProjectManagerName(project) {
        if (!(project === null || project === void 0 ? void 0 : project.ProjectManager))
            return '';
        return typeof project.ProjectManager === 'object'
            ? project.ProjectManager.Title
            : project.ProjectManager;
    }
    // ============================================
    // PUBLIC PERMISSION CHECK METHODS
    // ============================================
    hasFullControl() {
        return this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL;
    }
    canContribute() {
        return this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE ||
            this.context.userPermissionLevel === PERMISSION_LEVELS.FULL_CONTROL;
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
        return this.isInReadGroup(this.context.userGroups);
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
        // READ_ALL users (SC_Team, Controlling) - Read only all projects
        if (this.context.userPermissionLevel === PERMISSION_LEVELS.READ_ALL)
            return false;
        // READ users (Engineering roles) - Read only their projects
        if (this.context.userPermissionLevel === PERMISSION_LEVELS.READ)
            return false;
        // Contribute users need to be in specific groups with project-based access
        if (this.context.userPermissionLevel === PERMISSION_LEVELS.CONTRIBUTE) {
            // Check if user is in any contribute group
            if (!this.isInContributeGroup(this.context.userGroups)) {
                return false;
            }
            // GH Heads can edit all projects
            if (this.context.userGroups.includes('KRIYA_GH_Heads')) {
                return true;
            }
            // PM Managers can ONLY edit their own projects
            const projectManager = this.getProjectManagerName(project);
            if (projectManager) {
                return projectManager.toLowerCase() === this.context.currentUserFullName.toLowerCase();
            }
            return false;
        }
        return false;
    }
    canCreateRecord() {
        return this.canContribute();
    }
    canViewForm(isEdit = false, project) {
        if (!this.canRead())
            return false;
        if (isEdit) {
            return this.canEditProject(project);
        }
        return this.canCreateRecord();
    }
    // ============================================
    // GET ASSIGNED PROJECTS
    // ============================================
    getAssignedProjects(allProjects) {
        // Full Control users see all projects
        if (this.hasFullControl()) {
            console.log('✅ Full Control - showing all projects');
            return allProjects;
        }
        // READ_ALL users see all projects but read-only
        if (this.canReadAllProjects()) {
            console.log('✅ Read All - showing all projects (read-only)');
            return allProjects;
        }
        // GH Heads see all projects (they can edit)
        if (this.context.userGroups.some(group => group === 'KRIYA_GH_Heads')) {
            console.log('✅ GH Head - showing all projects');
            return allProjects;
        }
        // READ users (Engineering) see only their projects (read-only)
        if (this.context.userPermissionLevel === PERMISSION_LEVELS.READ) {
            const userFullName = this.context.currentUserFullName.trim();
            const assigned = allProjects.filter(project => {
                const projectManager = this.getProjectManagerName(project);
                return projectManager.toLowerCase() === userFullName.toLowerCase();
            });
            console.log(`📋 Engineering User (${userFullName}) found ${assigned.length} assigned projects (read-only) out of ${allProjects.length}`);
            return assigned;
        }
        // PM Managers: Filter by Project Manager (they can edit their projects)
        if (this.context.userGroups.some(group => group === 'KRIYA_PM_Managers')) {
            const userFullName = this.context.currentUserFullName.trim();
            const assigned = allProjects.filter(project => {
                const projectManager = this.getProjectManagerName(project);
                return projectManager.toLowerCase() === userFullName.toLowerCase();
            });
            console.log(`📋 PM Manager (${userFullName}) found ${assigned.length} assigned projects out of ${allProjects.length}`);
            console.log('📋 Assigned projects:', assigned.map(project => project.Title));
            return assigned;
        }
        // Default: no projects
        console.log('⚠️ No projects assigned or insufficient permissions');
        return [];
    }
    getAssignedProjectValues(allProjects, key) {
        const assigned = this.getAssignedProjects(allProjects);
        const values = new Set();
        for (const project of assigned) {
            let value = project[key];
            if (!value)
                continue;
            // Handle ProjectManager specially
            if (key === 'ProjectManager' && typeof value === 'object') {
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
        const canContribute = this.canContribute();
        const canExport = this.canRead();
        const canFreeze = this.canFreeze();
        const isReadOnly = this.context.userPermissionLevel === PERMISSION_LEVELS.READ ||
            this.context.userPermissionLevel === PERMISSION_LEVELS.READ_ALL;
        // Import button
        if (selectors.importBtn) {
            const button = document.querySelector(selectors.importBtn);
            if (button) {
                button.style.display = canContribute ? '' : 'none';
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
                button.style.display = canContribute ? '' : 'none';
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
export const canContribute = () => authManager.canContribute();
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
    hasFullControl,
    canContribute,
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
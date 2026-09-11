export declare const PERMISSION_LEVELS: {
    readonly FULL_CONTROL: "full_control";
    readonly CONTRIBUTE: "contribute";
    readonly READ: "read";
    readonly READ_ALL: "read_all";
    readonly NONE: "none";
};
export type PermissionLevel = typeof PERMISSION_LEVELS[keyof typeof PERMISSION_LEVELS];
export declare const GROUP_PERMISSIONS: Record<string, PermissionLevel>;
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
declare class AuthManager {
    private static instance;
    private context;
    private _config;
    private _initialized;
    private constructor();
    static getInstance(): AuthManager;
    initialize(config: AuthConfig): Promise<void>;
    private fetchCurrentUserFullName;
    private fetchUserGroups;
    private determinePermissionLevel;
    private isInContributeGroup;
    private isInReadGroup;
    private getProjectManagerName;
    hasFullControl(): boolean;
    canContribute(): boolean;
    canRead(): boolean;
    canReadAllProjects(): boolean;
    canFreeze(): boolean;
    isGHHead(): boolean;
    isPMManager(): boolean;
    isEngineeringUser(): boolean;
    getUserGroups(): string[];
    getCurrentUserFullName(): string;
    getCurrentUserLogin(): string;
    getPermissionLevel(): PermissionLevel;
    isInitialized(): boolean;
    getConfig(): AuthConfig | null;
    canEditProject(project?: {
        ProjectManager?: string | {
            Title: string;
        };
    }): boolean;
    canCreateRecord(): boolean;
    canViewForm(isEdit?: boolean, project?: {
        ProjectManager?: string | {
            Title: string;
        };
    }): boolean;
    getAssignedProjects<T extends {
        Id: number;
        Title: string;
        ProjectManager?: string | {
            Title: string;
        };
    }>(allProjects: T[]): T[];
    getAssignedProjectValues<T extends {
        Id: number;
        Title: string;
        ProjectManager?: string | {
            Title: string;
        };
        [key: string]: any;
    }>(allProjects: T[], key: string): string[];
    updateUIBasedOnPermissions(selectors?: {
        importBtn?: string;
        exportBtn?: string;
        addBtn?: string;
        saveBtn?: string;
        freezeCheckboxes?: string;
    }): void;
    getDigest(siteUrl: string): Promise<string>;
    getListEntityType(siteUrl: string, listName: string): Promise<string>;
    reset(): void;
}
export declare const authManager: AuthManager;
export declare const getAuth: () => AuthManager;
export declare const hasFullControl: () => boolean;
export declare const canContribute: () => boolean;
export declare const canRead: () => boolean;
export declare const canReadAllProjects: () => boolean;
export declare const canFreeze: () => boolean;
export declare const isGHHead: () => boolean;
export declare const isPMManager: () => boolean;
export declare const isEngineeringUser: () => boolean;
export declare const getUserGroups: () => string[];
export declare const getCurrentUserFullName: () => string;
export declare const getCurrentUserLogin: () => string;
export declare const getPermissionLevel: () => PermissionLevel;
export declare const canEditProject: (project?: {
    ProjectManager?: string | {
        Title: string;
    };
}) => boolean;
export declare const canCreateRecord: () => boolean;
export declare const canViewForm: (isEdit?: boolean, project?: {
    ProjectManager?: string | {
        Title: string;
    };
}) => boolean;
export declare const getAssignedProjects: <T extends {
    Id: number;
    Title: string;
    ProjectManager?: string | {
        Title: string;
    } | undefined;
}>(allProjects: T[]) => T[];
export declare const getAssignedProjectValues: <T extends {
    [key: string]: any;
    Id: number;
    Title: string;
    ProjectManager?: string | {
        Title: string;
    } | undefined;
}>(allProjects: T[], key: string) => string[];
export declare const updateUIBasedOnPermissions: (selectors?: {
    importBtn?: string;
    exportBtn?: string;
    addBtn?: string;
    saveBtn?: string;
    freezeCheckboxes?: string;
}) => void;
export declare const getDigest: (siteUrl: string) => Promise<string>;
export declare const getListEntityType: (siteUrl: string, listName: string) => Promise<string>;
declare const _default: {
    authManager: AuthManager;
    PERMISSION_LEVELS: {
        readonly FULL_CONTROL: "full_control";
        readonly CONTRIBUTE: "contribute";
        readonly READ: "read";
        readonly READ_ALL: "read_all";
        readonly NONE: "none";
    };
    GROUP_PERMISSIONS: Record<string, PermissionLevel>;
    hasFullControl: () => boolean;
    canContribute: () => boolean;
    canRead: () => boolean;
    canReadAllProjects: () => boolean;
    canFreeze: () => boolean;
    isGHHead: () => boolean;
    isPMManager: () => boolean;
    isEngineeringUser: () => boolean;
    getUserGroups: () => string[];
    getCurrentUserFullName: () => string;
    getCurrentUserLogin: () => string;
    getPermissionLevel: () => PermissionLevel;
    canEditProject: (project?: {
        ProjectManager?: string | {
            Title: string;
        } | undefined;
    } | undefined) => boolean;
    canCreateRecord: () => boolean;
    canViewForm: (isEdit?: boolean, project?: {
        ProjectManager?: string | {
            Title: string;
        } | undefined;
    } | undefined) => boolean;
    getAssignedProjects: <T extends {
        Id: number;
        Title: string;
        ProjectManager?: string | {
            Title: string;
        } | undefined;
    }>(allProjects: T[]) => T[];
    getAssignedProjectValues: <T_1 extends {
        [key: string]: any;
        Id: number;
        Title: string;
        ProjectManager?: string | {
            Title: string;
        } | undefined;
    }>(allProjects: T_1[], key: string) => string[];
    updateUIBasedOnPermissions: (selectors?: {
        importBtn?: string | undefined;
        exportBtn?: string | undefined;
        addBtn?: string | undefined;
        saveBtn?: string | undefined;
        freezeCheckboxes?: string | undefined;
    } | undefined) => void;
    getDigest: (siteUrl: string) => Promise<string>;
    getListEntityType: (siteUrl: string, listName: string) => Promise<string>;
};
export default _default;
//# sourceMappingURL=auth.d.ts.map
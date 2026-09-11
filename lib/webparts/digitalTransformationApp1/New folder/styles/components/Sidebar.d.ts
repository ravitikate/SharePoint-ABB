import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface ISidebarProps {
    context: WebPartContext;
}
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
 * Check if user is in a specific group using the shared auth module
 */
export declare const isUserInGroup: (groupName: string) => boolean;
export declare const getSidebarContent: (context: WebPartContext) => Promise<string>;
export declare const sidebarStyle = "\n.dt-sidebar {\n  background: #ffffff;\n  box-shadow: 0 1px 6px -2px #eee, 0 2px 6px 1px #eee, 0 1px 4px 2px #eee;\n  width: 220px;\n  padding: 10px;\n  transition: width 0.3s;\n  height: 100vh;\n  overflow-y: auto;\n  font-family: Arial, sans-serif;\n}\n\n.dt-sidebar.collapsed {\n  width: 60px;\n}\n\n.dt-sidebar ul {\n  list-style: none;\n  padding: 0;\n  margin: 0;\n}\n\n.dt-sidebar li {\n  margin: 5px 0;\n}\n\n.dt-sidebar a {\n  display: flex;\n  align-items: center;\n  gap: 10px;\n  padding: 10px 15px;\n  color: #4d5f71;\n  text-decoration: none;\n  border-radius: 8px;\n  position: relative;\n}\n\n.dt-sidebar a:hover {\n  background: #e7e7ff;\n  color:#696cff\n}\n\n.dt-sidebar.collapsed .icon-label {\n  display: none;\n}\n\n.icon {\n  display: inline-flex;\n  align-items: center;\n  justify-content: center;\n  cursor: pointer;\n  position: relative;\n}\n\n/* TOOLTIP */\n.icon[data-tooltip]:hover::after {\n  content: attr(data-tooltip);\n  position: absolute;\n  bottom: 125%;\n  left: 50%;\n  transform: translateX(-50%);\n  background: #333;\n  color: #fff;\n  font-size: 11px;\n  padding: 4px 6px;\n  border-radius: 4px;\n  white-space: nowrap;\n  z-index: 9999;\n  opacity: 1;\n}\n\n.icon[data-tooltip]:hover::before {\n  content: '';\n  position: absolute;\n  bottom: 115%;\n  left: 50%;\n  transform: translateX(-50%);\n  border-width: 5px;\n  border-style: solid;\n  border-color: #333 transparent transparent transparent;\n  z-index: 9999;\n  opacity: 1;\n}\n\nsvg {\n  display: block;\n  width: 20px;\n  height: 20px;\n}\n";
export declare const sidebarScript: () => void;
declare const _default: {
    getContent: (context: WebPartContext) => Promise<string>;
    style: string;
    script: () => void;
};
export default _default;
//# sourceMappingURL=Sidebar.d.ts.map
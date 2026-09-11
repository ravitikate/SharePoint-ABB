// src/webparts/digitalTransformationApp1/components/Sidebar.ts
import { authManager, hasFullControl, canContribute, isGHHead, getUserGroups } from './pages/auth';
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
    console.log("Logged-in user belongs to groups:");
    groups.forEach((group) => {
        console.log(`Name: ${group}`);
    });
};
/**
 * Check if user is in a specific group using the shared auth module
 */
export const isUserInGroup = (groupName) => {
    const groups = getUserGroups();
    return groups.some(group => group.toLowerCase() === groupName.toLowerCase());
};
// Generate sidebar content based on user role
export const getSidebarContent = async (context) => {
    // ✅ Initialize auth first (if not already initialized)
    await initializeSidebarAuth(context);
    // ✅ Get user info from authManager
    const isAdmin = hasFullControl();
    const isGH = isGHHead();
    const canUserContribute = canContribute();
    const userFullName = authManager.getCurrentUserFullName();
    const userGroups = getUserGroups();
    const currentUser = context.pageContext.user;
    console.log("Logged-in User:");
    console.log("Display Name:", currentUser.displayName);
    console.log("Email:", currentUser.email);
    console.log("Login Name:", currentUser.loginName);
    console.log("User ID:", context.pageContext.legacyPageContext.userId);
    // ✅ Get all groups the user belongs to from authManager
    await getCurrentUserGroups();
    console.log("All Site Groups (from auth):", userGroups);
    console.log("User Info:", {
        name: currentUser.displayName,
        email: currentUser.email,
        login: currentUser.loginName,
        id: context.pageContext.legacyPageContext.userId,
        fullName: userFullName,
        isAdmin: isAdmin,
        isGH: isGH,
        canContribute: canUserContribute
    });
    console.log("Is Kirya Admin?:", isAdmin);
    console.log("Is GH Head?:", isGH);
    // ✅ BASE SIDEBAR ITEMS (Visible to all users with read permission)
    let sidebarItems = `
    <li><a href="#" data-page="dashboard"><span class="icon" data-tooltip="Dashboard"><svg viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1" stroke="#0078d4" stroke-width="2" fill="none"/><rect x="14" y="3" width="7" height="7" rx="1" stroke="#0078d4" stroke-width="2" fill="none"/><rect x="3" y="14" width="7" height="7" rx="1" stroke="#0078d4" stroke-width="2" fill="none"/><rect x="14" y="14" width="7" height="7" rx="1" stroke="#0078d4" stroke-width="2" fill="none"/></svg></span><span class="icon-label">Dashboard</span></a></li>
  `;
    // ✅ CONDITIONALLY ADD "MASTER DATA" ITEM FOR ADMIN ONLY
    if (isAdmin) {
        console.log("✅ User IS a Kirya Admin - Adding Master Data menu item");
        sidebarItems += `
      <li><a href="#" data-page="master"><span class="icon" data-tooltip="Master"><svg viewBox="0 0 24 24"><polygon points="12 2 2 7 12 12 22 7 12 2" stroke="#0078d4" stroke-width="2" fill="none"/><polyline points="2 17 12 22 22 17" stroke="#0078d4" stroke-width="2" fill="none"/><polyline points="2 12 12 17 22 12" stroke="#0078d4" stroke-width="2" fill="none"/></svg></span><span class="icon-label">Master Data</span></a></li>
    `;
    }
    else {
        console.log("❌ User is NOT a Kirya Admin - Master Data menu item hidden");
    }
    // ✅ ADD REMAINING ITEMS (Visible to all users with read permission)
    sidebarItems += `  
    <li><a href="#" data-page="project"><span class="icon" data-tooltip="Project"><svg viewBox="0 0 24 24"><polygon points="12 2 2 7 12 12 22 7 12 2" stroke="#0078d4" stroke-width="2" fill="none"/><polyline points="2 17 12 22 22 17" stroke="#0078d4" stroke-width="2" fill="none"/><polyline points="2 12 12 17 22 12" stroke="#0078d4" stroke-width="2" fill="none"/></svg></span><span class="icon-label">Project</span></a></li>
    
    <li><a href="#" data-page="dci"><span class="icon" data-tooltip="DCI"><svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="#0078d4" stroke-width="2" fill="none"/><polyline points="14 2 14 8 20 8" stroke="#0078d4" stroke-width="2"/><line x1="16" y1="13" x2="8" y2="13" stroke="#0078d4" stroke-width="2"/><line x1="16" y1="17" x2="8" y2="17" stroke="#0078d4" stroke-width="2"/><polyline points="10 9 9 9 8 9" stroke="#0078d4" stroke-width="2"/></svg></span><span class="icon-label">DCI</span></a></li>
    
    <li><a href="#" data-page="procurement"><span class="icon" data-tooltip="Procurement"><svg viewBox="0 0 24 24"><path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z" stroke="#0078d4" stroke-width="2" fill="none"/><line x1="3" y1="6" x2="21" y2="6" stroke="#0078d4" stroke-width="2"/><path d="M16 10a4 4 0 0 1-8 0" stroke="#0078d4" stroke-width="2"/></svg></span><span class="icon-label">Procurement</span></a></li>
    
    <li><a href="#" data-page="revenue"><span class="icon" data-tooltip="Revenue"><svg viewBox="0 0 24 24"><path d="M12 1v22" stroke="#0078d4" stroke-width="2"/><path d="M17 5H9.5a3.5 3.5 0 1 0 0 7h5a3.5 3.5 0 1 1 0 7H6" stroke="#0078d4" stroke-width="2"/></svg></span><span class="icon-label">Revenue</span></a></li>
    
    <li><a href="#" data-page="collection"><span class="icon" data-tooltip="Collection"><svg viewBox="0 0 24 24"><path d="M20 12V8H6a2 2 0 0 1-2-2V3a1 1 0 0 0-1-1" stroke="#0078d4" stroke-width="2" fill="none"/><path d="M4 21h16a2 2 0 0 0 2-2v-5H6a2 2 0 0 1-2-2v-1a1 1 0 0 0-1-1" stroke="#0078d4" stroke-width="2" fill="none"/><circle cx="8" cy="15" r="1" stroke="#0078d4" stroke-width="2" fill="none"/><circle cx="16" cy="15" r="1" stroke="#0078d4" stroke-width="2" fill="none"/></svg></span><span class="icon-label">Collection</span></a></li>
     
    <li><a href="#" data-page="invoicing"><span class="icon" data-tooltip="Invoicing Planning"><svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="#0078d4" stroke-width="2" fill="none"/><polyline points="14 2 14 8 20 8" stroke="#0078d4" stroke-width="2"/><line x1="16" y1="13" x2="8" y2="13" stroke="#0078d4" stroke-width="2"/><line x1="16" y1="17" x2="8" y2="17" stroke="#0078d4" stroke-width="2"/><polyline points="10 9 9 9 8 9" stroke="#0078d4" stroke-width="2"/><path d="M12 17v-5" stroke="#0078d4" stroke-width="2"/></svg></span><span class="icon-label">Invoicing Planning</span></a></li>
    
    <li><a href="#" data-page="inventory"><span class="icon" data-tooltip="Inventory"><svg viewBox="0 0 24 24"><path d="M3 6h18" stroke="#0078d4" stroke-width="2"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" stroke="#0078d4" stroke-width="2" fill="none"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" stroke="#0078d4" stroke-width="2" fill="none"/><line x1="10" y1="11" x2="10" y2="17" stroke="#0078d4" stroke-width="2"/><line x1="14" y1="11" x2="14" y2="17" stroke="#0078d4" stroke-width="2"/></svg></span><span class="icon-label">Inventory</span></a></li>
    
    <li><a href="#" data-page="seibies"><span class="icon" data-tooltip="SEI–BIES"><svg viewBox="0 0 24 24"><path d="M12 2L2 7l10 5 10-5-10-5z" stroke="#0078d4" stroke-width="2" fill="none"/><path d="M2 17l10 5 10-5" stroke="#0078d4" stroke-width="2" fill="none"/><path d="M2 12l10 5 10-5" stroke="#0078d4" stroke-width="2" fill="none"/><path d="M12 22V12" stroke="#0078d4" stroke-width="2"/></svg></span><span class="icon-label">SEI–BIES</span></a></li>
    
    <li><a href="#" data-page="knowledge"><span class="icon" data-tooltip="Knowledge Base"><svg viewBox="0 0 24 24"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z" stroke="#0078d4" stroke-width="2" fill="none"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z" stroke="#0078d4" stroke-width="2" fill="none"/></svg></span><span class="icon-label">Knowledge Base</span></a></li>
  `;
    // ✅ RETURN THE FULL HTML CONTENT
    return `
    <div class="dt-sidebar">
      <ul class="dt-sidebar-list">
        ${sidebarItems}
      </ul>
    </div>
  `;
};
// Styles remain the same
export const sidebarStyle = `
.dt-sidebar {
  background: #ffffff;
  box-shadow: 0 1px 6px -2px #eee, 0 2px 6px 1px #eee, 0 1px 4px 2px #eee;
  width: 220px;
  padding: 10px;
  transition: width 0.3s;
  height: 100vh;
  overflow-y: auto;
  font-family: Arial, sans-serif;
}

.dt-sidebar.collapsed {
  width: 60px;
}

.dt-sidebar ul {
  list-style: none;
  padding: 0;
  margin: 0;
}

.dt-sidebar li {
  margin: 5px 0;
}

.dt-sidebar a {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 15px;
  color: #4d5f71;
  text-decoration: none;
  border-radius: 8px;
  position: relative;
}

.dt-sidebar a:hover {
  background: #e7e7ff;
  color:#696cff
}

.dt-sidebar.collapsed .icon-label {
  display: none;
}

.icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  position: relative;
}

/* TOOLTIP */
.icon[data-tooltip]:hover::after {
  content: attr(data-tooltip);
  position: absolute;
  bottom: 125%;
  left: 50%;
  transform: translateX(-50%);
  background: #333;
  color: #fff;
  font-size: 11px;
  padding: 4px 6px;
  border-radius: 4px;
  white-space: nowrap;
  z-index: 9999;
  opacity: 1;
}

.icon[data-tooltip]:hover::before {
  content: '';
  position: absolute;
  bottom: 115%;
  left: 50%;
  transform: translateX(-50%);
  border-width: 5px;
  border-style: solid;
  border-color: #333 transparent transparent transparent;
  z-index: 9999;
  opacity: 1;
}

svg {
  display: block;
  width: 20px;
  height: 20px;
}
`;
export const sidebarScript = () => console.log("Sidebar loaded with role-based visibility");
// For backward compatibility with your imports
export default {
    getContent: getSidebarContent,
    style: sidebarStyle,
    script: sidebarScript
};
//# sourceMappingURL=Sidebar.js.map
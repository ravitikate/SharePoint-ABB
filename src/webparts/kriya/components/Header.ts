// src/webparts/Kriya/components/Header.ts
//
// The Add User button is NOT inside the header in the DOM. It is a
// fixed-position element attached to <body> so that no collapsed header
// height, overflow rule, flex behaviour, or workbench quirk can ever
// affect it. The script positions it in the top-right corner of the
// viewport and shows it only for members of KRIYA_Admin_Full.

import logoPath from '../assets/abb.png';
import kriyaLogoPath from '../assets/kriya_logo.png'; // ✅ Import the new logo

const ADD_USER_LINK = 'https://abb.sharepoint.com/sites/Kriya/_layouts/15/user.aspx';
const ADMIN_GROUP_CANONICAL = 'KRIYA_Admin_Full';
const BTN_ID = 'dtAddUserBtn';

// ------------------------------------------------------------------
// HEADER HTML — note that the button is NOT in here. It is created
// dynamically and appended to document.body by the script below.
// ------------------------------------------------------------------
const headerContent = `
  <div class="dt-header-content">
    <div class="left-section">
      <button class="toggle-sidebar" aria-label="Toggle sidebar">&#9776;</button>
      <img src="${logoPath}" alt="ABB" class="dt-logo" />
      <div class="brand-block">
        <span class="brand-title">Energy Industries</span>
        <span class="brand-sub">India</span>
      </div>
      <div class="app-divider"></div>
      <!-- ✅ Replaced text with the new logo image -->
      <img src="${kriyaLogoPath}" alt="Kriya" class="kriya-logo" />
    </div>
    <nav class="dt-nav">
      <div class="user-avatar dt-user-avatar" data-dt-avatar="1" title="Loading...">--</div>
    </nav>
  </div>
`;

// ------------------------------------------------------------------
// HEADER STYLES
// ------------------------------------------------------------------
const headerStyle = `
  /* ---------- header shell ---------- */
  .dt-header {
    flex-shrink: 0;
    flex-grow: 0;
    min-height: 56px;
    width: 100%;
    padding: 0;
    background: #ffffff;
    border-bottom: 1px solid #ececec;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
    position: relative;
    z-index: 100;
    overflow: visible;
  }

  .dt-header-content {
    display: flex;
    align-items: center;
    justify-content: space-between;
    width: 100%;
    height: 56px;
    min-height: 56px;
    padding: 0 20px;
    box-sizing: border-box;
    color: #4d5f71;
  }

  .left-section {
    display: flex;
    align-items: center;
    gap: 12px;
    flex-shrink: 0;
    min-width: 0;
  }

  .toggle-sidebar {
    cursor: pointer;
    font-size: 22px;
    background: none;
    border: none;
    color: #4d5f71;
    padding: 4px 6px;
    border-radius: 4px;
    transition: background 0.18s ease;
  }
  .toggle-sidebar:hover { background: #f2f4f7; }

  .dt-logo { width: 52px; height: auto; display: block; }

  .brand-block { display: flex; flex-direction: column; line-height: 1.1; margin-left: 4px; }
  .brand-title { font-size: 15px; font-weight: 700; color: #1a1a1a; }
  .brand-sub   { font-size: 10px; color: #8a94a6; font-weight: 500; }

  .app-divider { width: 1px; height: 26px; background: #e0e0e0; margin: 0 14px; }

  /* ✅ New Kriya Logo Style */
  .kriya-logo { 
    height: 38px; 
    width: auto; 
    display: block; 
    object-fit: contain;
  }

  .dt-nav {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-shrink: 1;
    overflow: visible;
    justify-content: flex-end;
  }

  .user-avatar {
    width: 36px;
    height: 36px;
    border-radius: 50%;
    background: #e6e8eb;
    color: #4d5f71;
    font-size: 13px;
    font-weight: 700;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    margin-left: 4px;
    text-transform: uppercase;
    user-select: none;
    transition: background 0.18s ease;
    flex-shrink: 0;
  }
  .user-avatar:hover { background: #d8dbe0; }

  /* ============================================================
     ADD USER BUTTON — fixed to the viewport, attached to <body>
     ============================================================ */
  #dtAddUserBtn {
    position: fixed !important;
    top: 10px !important;
    right: 70px !important;
    left: auto !important;
    bottom: auto !important;
    display: inline-flex !important;
    align-items: center;
    gap: 6px;
    padding: 6px 12px;
    font-size: 12.5px;
    font-weight: 600;
    color: #ffffff !important;
    background: #e01a23;
    border: 1px solid #c5171f;
    border-radius: 6px;
    text-decoration: none;
    cursor: pointer;
    white-space: nowrap;
    z-index: 2147483647 !important;
    box-shadow: 0 2px 6px rgba(224, 26, 35, 0.30);
    transition: background 0.18s ease, box-shadow 0.18s ease;
    font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", Arial, sans-serif;
  }
  #dtAddUserBtn[hidden] { display: none !important; }
  #dtAddUserBtn:hover {
    background: #c5171f;
    box-shadow: 0 3px 10px rgba(224, 26, 35, 0.45);
  }
  #dtAddUserBtn .dt-add-user-icon {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 16px;
    height: 16px;
    border-radius: 50%;
    background: rgba(255, 255, 255, 0.22);
    font-size: 13px;
    font-weight: 700;
    line-height: 1;
  }

  /* ---------- responsive ---------- */
  @media (max-width: 900px) {
    .dt-header-content { padding: 0 16px; }
    .app-divider { margin: 0 10px; }
    #dtAddUserBtn { right: 60px !important; }
  }

  @media (max-width: 600px) {
    .dt-header-content { height: 52px; padding: 0 12px; }
    .brand-sub { display: none; }
    .brand-title { font-size: 14px; }
    .kriya-logo { height: 30px; } /* ✅ Shrinks on mobile */
    .dt-logo { width: 44px; }
    .app-divider { margin: 0 8px; }
    .user-avatar { width: 32px; height: 32px; font-size: 12px; }
    #dtAddUserBtn { right: 52px !important; padding: 6px 8px; }
    #dtAddUserBtn .dt-add-user-label { display: none; }
  }

  @media (max-width: 400px) {
    .app-divider, .brand-block { display: none; }
    .dt-logo { width: 38px; }
    .kriya-logo { height: 24px; } /* ✅ Shrinks further on small screens */
    #dtAddUserBtn { right: 48px !important; }
  }
`;

// ============================================================
// INITIALS
// ============================================================
const getInitials = (name: string): string => {
  const cleaned = (name || '')
    .replace(/@.*$/, '')
    .replace(/[._\-]+/g, ' ')
    .replace(/[^A-Za-z\s]/g, '')
    .trim();

  const parts = cleaned.split(/\s+/).filter(Boolean);
  if (!parts.length) return '?';
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
};

const AVATAR_SELECTOR = '.dt-user-avatar, #dtUserAvatar';

const getAvatarEls = (): HTMLElement[] =>
  Array.from(document.querySelectorAll<HTMLElement>(AVATAR_SELECTOR));

const paintAll = (initials: string, fullName: string): void => {
  const els = getAvatarEls();
  if (!els.length) return;
  els.forEach(el => {
    el.textContent = initials;
    el.title = fullName || initials;
  });
};

const NAME_EVENT = 'dt-user-ready';

// ============================================================
// SHARED STATE — survives multiple header script mounts
// ============================================================
interface HeaderState {
  fullName: string;
  login: string;
  groups: string[];
  isAdmin: boolean | null;
  permissionLevel: string;
  buttonShouldShow: boolean;
}

const STATE_KEY = '__dtHeaderState';

const getSharedState = (): HeaderState => {
  const w = window as any;
  if (!w[STATE_KEY]) {
    w[STATE_KEY] = {
      fullName: '',
      login: '',
      groups: [],
      isAdmin: null,
      permissionLevel: '',
      buttonShouldShow: false
    };
  }
  return w[STATE_KEY] as HeaderState;
};

const paintFromName = (fullName: string): void => {
  if (!fullName) return;
  getSharedState().fullName = fullName;
  paintAll(getInitials(fullName), fullName);
};

// ============================================================
// ADMIN CHECK
// ============================================================
const normalise = (s: string): string =>
  (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

const groupsContainAdmin = (groups: string[]): boolean => {
  const target = normalise(ADMIN_GROUP_CANONICAL);   // "kriyaadminfull"
  return groups.some(g => normalise(g) === target);
};

// ============================================================
// BUTTON — attached to document.body, positioned with position:fixed
// ============================================================

/** Ensure the button element exists on document.body. */
const ensureBodyButton = (): HTMLAnchorElement => {
  let btn = document.getElementById(BTN_ID) as HTMLAnchorElement | null;

  if (!btn) {
    btn = document.createElement('a');
    btn.id = BTN_ID;
    btn.href = ADD_USER_LINK;
    btn.target = '_blank';
    btn.rel = 'noopener noreferrer';
    btn.title = 'Add User';
    btn.setAttribute('hidden', '');
    btn.innerHTML =
      '<span class="dt-add-user-icon">+</span>' +
      '<span class="dt-add-user-label">Add User</span>';
    document.body.appendChild(btn);
  } else if (btn.parentElement !== document.body) {
    // If it somehow ended up inside the header, move it to body
    document.body.appendChild(btn);
  }

  return btn;
};

const applyButtonVisibility = (): void => {
  const state = getSharedState();

  const shouldShow =
    state.permissionLevel === 'full_control' ||
    state.isAdmin === true ||
    (state.groups.length > 0 && groupsContainAdmin(state.groups));

  state.buttonShouldShow = shouldShow;

  const btn = ensureBodyButton();

  if (shouldShow) {
    btn.removeAttribute('hidden');
    // Force every relevant property so nothing can override it
    btn.style.setProperty('position', 'fixed', 'important');
    btn.style.setProperty('top', '10px', 'important');
    btn.style.setProperty('right', '70px', 'important');
    btn.style.setProperty('left', 'auto', 'important');
    btn.style.setProperty('bottom', 'auto', 'important');
    btn.style.setProperty('display', 'inline-flex', 'important');
    btn.style.setProperty('visibility', 'visible', 'important');
    btn.style.setProperty('opacity', '1', 'important');
    btn.style.setProperty('z-index', '2147483647', 'important');
  } else {
    btn.setAttribute('hidden', '');
    btn.style.setProperty('display', 'none', 'important');
  }
};

const evaluateAdminAndToggle = (opts: {
  login?: string;
  groups?: string[];
  isAdmin?: boolean;
  permissionLevel?: string;
}): void => {
  const state = getSharedState();
  if (opts.login) state.login = opts.login;
  if (opts.groups && opts.groups.length) state.groups = opts.groups;
  if (typeof opts.isAdmin === 'boolean') state.isAdmin = opts.isAdmin;
  if (opts.permissionLevel) state.permissionLevel = opts.permissionLevel;

  applyButtonVisibility();
};

const paintFromContext = (): void => {
  const sp: any = (window as any)._spPageContextInfo;
  if (!sp) return;
  const n = sp.userDisplayName || sp.userLoginName || sp.userEmail || '';
  paintFromName(n);
  if (sp.userLoginName) evaluateAdminAndToggle({ login: String(sp.userLoginName) });
};

const repaintFromMemory = (): boolean => {
  const state = getSharedState();
  if (!state.fullName) return false;
  const els = getAvatarEls();
  if (!els.length) return false;
  paintAll(getInitials(state.fullName), state.fullName);
  return true;
};

// ============================================================
// HEADER SCRIPT
// ============================================================
const headerScript = (): void => {
  console.log('%c======== Header script starting ========', 'color:#0078d4;font-weight:bold');

  // Create the Add User button on document.body right away
  ensureBodyButton();

  // Read whatever context is available
  paintFromContext();

  // Apply the current shared state immediately
  applyButtonVisibility();

  // Fast retry — 150 ms × 30 (≈ 4.5 s)
  let fastTries = 0;
  const fastRetry = (): void => {
    fastTries++;
    const els = getAvatarEls();
    const allPainted = els.length > 0 && els.every(el => {
      const t = (el.textContent || '').trim();
      return t !== '' && t !== '--' && t !== '?';
    });
    if (!allPainted) {
      if (!repaintFromMemory()) paintFromContext();
    }
    applyButtonVisibility();
    if (fastTries < 30) window.setTimeout(fastRetry, 150);
  };
  window.setTimeout(fastRetry, 150);

  // Watchdog every 500 ms — keeps the avatar and the button in sync forever
  window.setInterval(() => {
    const els = getAvatarEls();
    if (els.length) {
      const anyPlaceholder = els.some(el => {
        const t = (el.textContent || '').trim();
        return t === '' || t === '--' || t === '?';
      });
      if (anyPlaceholder) {
        if (!repaintFromMemory()) paintFromContext();
      }
    }
    applyButtonVisibility();
  }, 500);

  // Payload listener from KriyaWebPart
  document.addEventListener(NAME_EVENT, (evt: Event) => {
    const detail = (evt as CustomEvent).detail || {};
    const name            = detail.name  || '';
    const login           = detail.login || '';
    const email           = detail.email || '';
    const groups          = Array.isArray(detail.groups) ? detail.groups : [];
    const isAdmin         = typeof detail.isAdmin === 'boolean' ? detail.isAdmin : undefined;
    const permissionLevel = detail.permissionLevel || '';

    console.log('📥 Header received dt-user-ready:',
      name, '| login:', login,
      '| groups:', groups.length,
      '| isAdmin:', isAdmin,
      '| permission:', permissionLevel || '(none)');

    if (name) paintFromName(name);
    else if (email) paintFromName(email);

    evaluateAdminAndToggle({ login, groups, isAdmin, permissionLevel });

    // Reapply a few times to defeat any late DOM churn
    window.setTimeout(applyButtonVisibility, 0);
    window.setTimeout(applyButtonVisibility, 100);
    window.setTimeout(applyButtonVisibility, 300);
    window.setTimeout(applyButtonVisibility, 800);
  });
};

export default {
  content: headerContent,
  style: headerStyle,
  script: headerScript
};
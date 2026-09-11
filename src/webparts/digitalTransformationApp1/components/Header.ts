// src/webparts/digitalTransformationApp1/components/header.ts
import logoPath from '../assets/abb.png';

const headerContent = `
  <div class="dt-header-content">
    <div class="left-section">
      <button class="toggle-sidebar">&#9776;</button>
      <img src="${logoPath}" alt="Logo" class="dt-logo" />
    </div>
    <nav class="dt-nav">
      <!-- <a href="#" class="dt-nav-link">Home</a>
      <a href="#" class="dt-nav-link">Contact</a>
      <a href="#" class="dt-nav-link">Help</a> -->
    </nav>
  </div>
`;

const headerStyle = `
  /* Header container */
  .dt-header, .dt-footer {
      flex-shrink: 0;
      width: 98%;
      padding-right: 20px !important;
  }

  .dt-header-content {
    display: flex;
    align-items: center;
    justify-content: space-between;
    width: 100%;
    padding: 0 15px; /* reduced padding to prevent overflow */
    box-sizing: border-box;
    color: #4d5f71;
    height: 60px;
    background: #ffffff;
    box-shadow: 2px 2px 4px rgba(0,0,0,0.05);
  }

  /* Left section: toggle + logo */
  .left-section {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-shrink: 0; /* do not shrink */
  }

  .toggle-sidebar {
    cursor: pointer;
    font-size: 24px;
    background: none;
    border: none;
    color: #4d5f71;
  }

  .dt-logo {
    width: 55px;
    height: auto;
  }

  /* Navigation */
  .dt-nav {
    display: flex;
    gap: 15px;
    flex-shrink: 1; /* allow shrinking if needed */
    overflow: hidden;
    justify-content: flex-end;
  }

  .dt-nav-link {
    color: #4d5f71;
    text-decoration: none;
    font-weight: bold;
    white-space: nowrap; /* prevent breaking */
    flex-shrink: 1; /* allow each link to shrink */
    min-width: 0; /* allow shrinking beyond content width */
    text-overflow: ellipsis; /* optional: show ... if too small */
    overflow: hidden;
  }

  .dt-nav-link:hover {
    text-decoration: underline;
  }

  /* Small screens: reduce gap */
  @media (max-width: 700px) {
    .dt-nav {
      gap: 8px;
    }
    .dt-nav-link {
      font-size: 14px;
    }
  }

  @media (max-width: 500px) {
    .dt-nav {
      gap: 5px;
      font-size: 13px;
    }
  }

  @media (max-width: 400px) {
    .dt-nav {
      display: none; /* hide nav completely on tiny screens */
    }
  }
`;

const headerScript = () => console.log("Header loaded");

export default {
  content: headerContent,
  style: headerStyle,
  script: headerScript
};    

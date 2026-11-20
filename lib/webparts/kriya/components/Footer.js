// src/webparts/Kriya/components/Footer.ts
// ============================================================
// FINANCIAL YEAR (Indian FY: 1 Apr – 31 Mar)
// ============================================================
const getCurrentFY = () => {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1; // 1–12
    const startYear = month >= 4 ? year : year - 1;
    const endYearShort = String(startYear + 1).slice(-2);
    return `${startYear}-${endYearShort}`;
};
// ============================================================
// AS-ON DATE — shown on the left of the footer
// (defaults to today; the dashboard may override via an event)
// ============================================================
const pad2 = (n) => String(n).padStart(2, '0');
const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtAsOn = (isoDate) => {
    if (!isoDate)
        return '';
    const d = new Date(`${isoDate}T00:00:00`);
    if (isNaN(d.getTime()))
        return isoDate;
    return `${pad2(d.getDate())}-${MONTH_ABBR[d.getMonth()]}-${d.getFullYear()}`;
};
const todayIso = new Date().toISOString().slice(0, 10);
const footerContent = `
  <div class="dt-footer-content">
    <div class="dt-footer-left">
      Data as on <b id="dtFooterAsOn">${fmtAsOn(todayIso)}</b>
      <span class="dt-footer-sep">|</span>
      All figures in INR
    </div>
    <div class="dt-footer-right">
      &copy; ${getCurrentFY()} Kriya
    </div>
  </div>
`;
const footerStyle = `
  .dt-footer {
    flex-shrink: 0;
    flex-grow: 0;
    width: 100%;
    background: #ffffff;
    border-top: 1px solid #ececec;
    padding: 8px 20px;
    min-height: 36px;
    box-sizing: border-box;
  }

  .dt-footer-content {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
    font-size: 12px;
    font-weight: 500;
    color: #8a93a6;
    letter-spacing: 0.2px;
    line-height: 1.3;
    width: 100%;
  }

  .dt-footer-left {
    display: flex;
    align-items: center;
    gap: 6px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .dt-footer-left b {
    color: #5b6478;
    font-weight: 600;
    margin: 0 2px;
  }

  .dt-footer-sep {
    color: #c8cdd8;
    margin: 0 6px;
  }

  .dt-footer-right {
    display: flex;
    align-items: center;
    white-space: nowrap;
    color: #8a93a6;
  }

  /* Tablet */
  @media (max-width: 900px) {
    .dt-footer {
      padding: 7px 16px;
    }
    .dt-footer-content {
      font-size: 11.5px;
    }
  }

  /* Mobile */
  @media (max-width: 600px) {
    .dt-footer {
      padding: 6px 12px;
    }
    .dt-footer-content {
      font-size: 10.5px;
      gap: 8px;
    }
    .dt-footer-sep {
      margin: 0 4px;
    }
  }

  /* Very small mobile — stack the two halves */
  @media (max-width: 420px) {
    .dt-footer-content {
      flex-direction: column;
      align-items: flex-start;
      gap: 2px;
    }
    .dt-footer-right {
      align-self: flex-end;
    }
  }
`;
const footerScript = () => {
    console.log('Footer loaded');
    // Optional: allow the dashboard to update the "Data as on" date via a
    // custom event. The dashboard fires this whenever it recomputes the
    // as-on date, so the footer stays in sync with the KPI snapshot.
    document.addEventListener('dt-as-on-changed', (evt) => {
        const detail = evt.detail || {};
        const iso = detail.iso || '';
        if (!iso)
            return;
        const el = document.getElementById('dtFooterAsOn');
        if (el) {
            el.textContent = fmtAsOn(iso);
        }
    });
};
export default {
    content: footerContent,
    style: footerStyle,
    script: footerScript
};
//# sourceMappingURL=Footer.js.map
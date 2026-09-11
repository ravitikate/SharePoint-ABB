const getCurrentFY = () => {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1; // 1–12
    // Indian FY: 1 Apr – 31 Mar
    const startYear = month >= 4 ? year : year - 1;
    const endYearShort = String(startYear + 1).slice(-2);
    return `${startYear}-${endYearShort}`;
};
const footerContent = `<p>© ${getCurrentFY()} Kriya</p>`;
const footerStyle = `
  .dt-footer {
    background: #ffffff;
    padding: 2px 12px;
    box-shadow: 0 1px 2px #eee;
    margin-top: 2px;
    text-align: center;
    line-height: 1.2;
    min-height: 0;
  }
  .dt-footer p {
    margin: 0;
    padding: 0;
    font-size: 12px;
    line-height: 1.2;
  }
`;
const footerScript = () => console.log("Footer loaded");
export default {
    content: footerContent,
    style: footerStyle,
    script: footerScript
};
//# sourceMappingURL=Footer.js.map
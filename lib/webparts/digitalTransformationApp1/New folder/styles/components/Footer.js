// src/webparts/digitalTransformationApp1/components/footer.ts
const footerContent = `<p>© 2025 Digital Transformation App</p>`;
const footerStyle = `
  .dt-footer {
    background: #ffffff;
    padding: 10px 20px;
     box-shadow: 2px 2px 2px 2px #eee;
     margin-top: 5px;
    text-align: center;
  }
`;
const footerScript = () => console.log("Footer loaded");
export default {
    content: footerContent,
    style: footerStyle,
    script: footerScript
};
//# sourceMappingURL=Footer.js.map
// src/webparts/Kriya/components/pages/bom.ts
const bomContent = `
  <div class="bom-page">
    <h1>bom</h1>
    <p>Welcome to the bom page.</p>
  </div>
`;
export default {
    content: bomContent,
    style: `
    .bom-page { font-family: Arial, sans-serif; padding: 20px; }
    .bom-page h1 { color: #0078d4; }
    .bom-page p { font-size: 16px; }
  `,
    script: () => console.log("bom loaded")
};
//# sourceMappingURL=bom.js.map
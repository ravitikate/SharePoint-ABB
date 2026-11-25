// src/webparts/Kriya/components/pages/knowledge.ts
const knowledgeContent = `
  <div class="knowledge-page">
    <h1>Knowledge Base</h1>
    <p>Under development.</p>
  </div>
`;
export default {
    content: knowledgeContent,
    style: `
    .knowledge-page { font-family: Arial, sans-serif; padding: 20px; }
    .knowledge-page h1 { color: #0078d4; }
    .knowledge-page p { font-size: 16px; }
  `,
    script: () => console.log("knowledge loaded")
};
//# sourceMappingURL=knowledge.js.map
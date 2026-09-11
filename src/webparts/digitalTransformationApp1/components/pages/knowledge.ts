 // src/webparts/digitalTransformationApp1/components/pages/knowledge.ts
const knowledgeContent = `
  <div class="knowledge-page">
    <h1>knowledge</h1>
    <p>Welcome to the knowledge page.</p>
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

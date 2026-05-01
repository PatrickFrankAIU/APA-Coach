const react = require("@vitejs/plugin-react");

function commonJsSourceBridge() {
  return {
    name: "apa-coach-commonjs-source-bridge",
    transform(code, id) {
      const fileId = id.split("?")[0];

      if (fileId.endsWith("/src/checks/checkApaFormatting.js")) {
        return {
          code: code.replace(
            /\nmodule\.exports = \{[\s\S]*?\n\};\s*$/u,
            `
export {
  checkApaFormatting,
  checkTitlePage,
  checkReferencesPage,
  checkReferencesFormatting,
  checkMargins,
  checkLineSpacingForRole,
  checkParagraphSpacingForRole,
  checkFirstLineIndents,
  checkAlignment,
};
`,
          ),
          map: null,
        };
      }

      if (fileId.endsWith("/src/docx/extractDocxFormatting.js")) {
        const browserCode = code.replace(
          'const { XMLParser } = require("fast-xml-parser");',
          'import { XMLParser } from "fast-xml-parser";',
        );

        return {
          code: browserCode.replace(
            /\nmodule\.exports = \{[\s\S]*?\n\};\s*$/u,
            `
export {
  extractDocxFormattingFromXml,
};
`,
          ),
          map: null,
        };
      }

      return null;
    },
  };
}

module.exports = {
  base: "/APA-Coach/",
  plugins: [react(), commonJsSourceBridge()],
};

const react = require("@vitejs/plugin-react");
const { VitePWA } = require("vite-plugin-pwa");

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
  checkPageNumbering,
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
  resolveHeaderFiles,
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
  plugins: [
    react(),
    commonJsSourceBridge(),
    VitePWA({
      registerType: "autoUpdate",
      workbox: {
        globPatterns: ["**/*.{js,css,html,png,svg,ico}"],
      },
      manifest: {
        name: "APA Coach",
        short_name: "APA Coach",
        version: "0.7.3",
        description: "Check APA formatting in your Word documents — right in the browser.",
        start_url: "/APA-Coach/",
        scope: "/APA-Coach/",
        display: "standalone",
        background_color: "#1f2933",
        theme_color: "#2a3441",
        icons: [
          { src: "pwa-192.png", sizes: "192x192", type: "image/png" },
          { src: "pwa-512.png", sizes: "512x512", type: "image/png", purpose: "any maskable" },
        ],
      },
    }),
  ],
};

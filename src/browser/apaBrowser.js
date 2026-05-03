import JSZip from "jszip";
import { extractDocxFormattingFromXml, resolveHeaderFiles } from "../docx/extractDocxFormatting.js";
import { checkApaFormatting } from "../checks/checkApaFormatting.js";

async function readZipText(zip, entryName) {
  const entry = zip.file(entryName);
  return entry ? entry.async("text") : null;
}

export async function analyzeDocxFile(file) {
  const buffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await readZipText(zip, "word/document.xml");
  const stylesXml = await readZipText(zip, "word/styles.xml");
  const relsXml = await readZipText(zip, "word/_rels/document.xml.rels");

  const headerFiles = resolveHeaderFiles(relsXml, documentXml);
  const headerXmlsByType = {};
  for (const [type, file] of Object.entries(headerFiles)) {
    headerXmlsByType[type] = await readZipText(zip, file);
  }

  const extracted = extractDocxFormattingFromXml(documentXml, stylesXml, relsXml, headerXmlsByType);
  const checks = checkApaFormatting(extracted);
  const passed = checks.filter((check) => check.status === "pass").length;
  const failed = checks.filter((check) => check.status === "fail").length;
  const review = checks.filter((check) => check.status === "review").length;

  return {
    file: file.name,
    generatedAt: new Date().toISOString(),
    extracted,
    checks,
    summary: {
      totalChecks: checks.length,
      passed,
      failed,
      review,
    },
  };
}

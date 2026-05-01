import JSZip from "jszip";
import { extractDocxFormattingFromXml } from "../docx/extractDocxFormatting.js";
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
  const extracted = extractDocxFormattingFromXml(documentXml, stylesXml);
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

import JSZip from "jszip";
import { extractDocxFormattingFromXml, resolveHeaderFiles } from "../docx/extractDocxFormatting.js";
import { checkApaFormatting } from "../checks/checkApaFormatting.js";
import { verifyReferenceLinks } from "../checks/verifyReferenceLinks.js";

async function readZipText(zip, entryName) {
  const entry = zip.file(entryName);
  return entry ? entry.async("text") : null;
}

const MAX_FILE_SIZE_BYTES = 20 * 1024 * 1024; // 20 MB

export async function analyzeDocxFile(file, onProgress) {
  if (file.size > MAX_FILE_SIZE_BYTES) {
    throw new Error("This file is too large. APA Coach can analyze documents up to 20 MB. Please check that you are submitting a standard Word document.");
  }
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
  const { checks, referenceGroups } = checkApaFormatting(extracted);

  onProgress?.("Verifying references…");
  const verifyCheck = await verifyReferenceLinks(referenceGroups);
  if (verifyCheck) checks.push(verifyCheck);

  const passed = checks.filter((check) => check.status === "pass").length;
  const failed = checks.filter((check) => check.status === "fail").length;
  const warn = checks.filter((check) => check.status === "warn").length;
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
      warn,
      review,
    },
  };
}

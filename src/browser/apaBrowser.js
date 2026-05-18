import JSZip from "jszip";
import { extractDocxFormattingFromXml, resolveHeaderFiles } from "../docx/extractDocxFormatting.js";
import { checkApaFormatting } from "../checks/checkApaFormatting.js";
import { verifyReferenceLinks } from "../checks/verifyReferenceLinks.js";

async function readZipText(zip, entryName) {
  const entry = zip.file(entryName);
  return entry ? entry.async("text") : null;
}

const MAX_FILE_SIZE_BYTES = 20 * 1024 * 1024; // 20 MB

function makeOfflineSkippedCheck() {
  return {
    rule: "Reference link verification",
    status: "skipped",
    passed: false,
    expected: "Each DOI should resolve to the source it claims to cite. Each URL should load and lead to the cited page.",
    expectedText: "Each DOI should resolve to the source it claims to cite. Each URL should load and lead to the cited page.",
    foundText: "No internet connection — reference links were not checked. Your formatting results above are complete.",
    found: "Skipped (offline)",
    applicable: 0,
    checked: 0,
    matched: 0,
    failed: 0,
    unknown: 0,
    details: [],
    howToFix: [],
    resources: [],
    missingItems: [],
  };
}

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

  if (!navigator.onLine) {
    checks.push(makeOfflineSkippedCheck());
  } else {
    onProgress?.("Verifying references…");
    const offlineController = new AbortController();
    const handleOffline = () => offlineController.abort();
    window.addEventListener("offline", handleOffline, { once: true });
    let verifyCheck;
    try {
      verifyCheck = await verifyReferenceLinks(referenceGroups, offlineController.signal);
    } finally {
      window.removeEventListener("offline", handleOffline);
    }
    if (!navigator.onLine) {
      checks.push(makeOfflineSkippedCheck());
    } else if (verifyCheck) {
      checks.push(verifyCheck);
    }
  }

  const passed = checks.filter((check) => check.status === "pass").length;
  const failed = checks.filter((check) => check.status === "fail").length;
  const warn = checks.filter((check) => check.status === "warn").length;
  const review = checks.filter((check) => check.status === "review").length;
  const skipped = checks.filter((check) => check.status === "skipped").length;

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
      skipped,
    },
  };
}

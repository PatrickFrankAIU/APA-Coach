#!/usr/bin/env node
// Captures a stable snapshot of every check result for a fixture .docx.
// Used to diff behavior before/after a change to matching logic.
// Usage: node scripts/capture-citation-snapshot.js <input.docx> <output.json>

const path = require("path");
const fs = require("fs");
const { extractDocxFormatting } = require("../src/docx/extractDocxFormatting");
const { checkApaFormatting } = require("../src/checks/checkApaFormatting");

// Kept in sync with scripts/test-citation-regression.js — see the comment there
// for why "Reference link verification" is excluded (currently a no-op, since
// this script never exercises the browser entry point that adds it).
const EXCLUDED_RULES = new Set(["Reference link verification"]);

function main() {
  const [inputPath, outputPath] = process.argv.slice(2);
  if (!inputPath || !outputPath) {
    console.error("Usage: node scripts/capture-citation-snapshot.js <input.docx> <output.json>");
    process.exitCode = 1;
    return;
  }

  const extracted = extractDocxFormatting(inputPath);
  const { checks } = checkApaFormatting(extracted);

  const snapshot = checks
    .filter((check) => !EXCLUDED_RULES.has(check.rule))
    .map((check) => ({
      rule: check.rule,
      status: check.status,
      applicable: check.applicable,
      checked: check.checked,
      matched: check.matched,
      failed: check.failed,
      unknown: check.unknown,
      details: check.details || [],
    }))
    .sort((a, b) => a.rule.localeCompare(b.rule));

  fs.writeFileSync(outputPath, JSON.stringify(snapshot, null, 2) + "\n");
  console.log(`Wrote ${snapshot.length} check snapshots to ${path.resolve(outputPath)}`);
}

main();

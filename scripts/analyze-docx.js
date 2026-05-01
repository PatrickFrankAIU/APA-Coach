#!/usr/bin/env node

const path = require("path");
const { extractDocxFormatting } = require("../src/docx/extractDocxFormatting");
const { checkApaFormatting } = require("../src/checks/checkApaFormatting");

function printUsage() {
  console.error("Usage: node scripts/analyze-docx.js path/to/file.docx [--verbose]");
}

function printReadableReport(report, options = {}) {
  console.log("\nAPA Formatting Report");
  console.log("=====================");
  console.log(`File: ${report.file}`);
  console.log(`APA Coach found ${report.summary.failed} issues that should be corrected before submission.`);

  if (options.verbose) {
    console.log(`Paragraphs analyzed: ${report.extracted.paragraphs.length}`);
    console.log(`Passed checks: ${report.summary.passed}/${report.summary.totalChecks}`);
    console.log(`Review checks: ${report.summary.review}`);
    console.log(`Failed checks: ${report.summary.failed}`);
  }

  for (const check of report.checks) {
    if (!options.verbose && check.status === "review" && check.applicable === 0 && check.rule.startsWith("Heading ")) {
      continue;
    }

    const mark = check.status.toUpperCase();
    console.log(`\n[${mark}] ${check.rule}`);
    console.log(`Found: ${check.foundText || check.found}`);
    console.log(`APA expects: ${check.expectedText || check.expected}`);

    if (options.verbose) {
      console.log(`Counts: applicable ${check.applicable}, checked ${check.checked}, matched ${check.matched}, review ${check.unknown}`);
    }

    if (check.status === "fail" && check.howToFix && check.howToFix.length > 0) {
      console.log("How to fix:");
      for (const step of check.howToFix) {
        console.log(`- ${step}`);
      }
    }

    if (options.verbose && check.details.length > 0) {
      console.log("Details:");
      for (const detail of check.details) {
        console.log(`- ${detail}`);
      }
    }
  }
}

async function main() {
  const args = process.argv.slice(2);
  const verbose = args.includes("--verbose");
  const inputPath = args.find((arg) => arg !== "--verbose");

  if (!inputPath) {
    printUsage();
    process.exitCode = 1;
    return;
  }

  if (path.extname(inputPath).toLowerCase() !== ".docx") {
    console.error("Input file must have a .docx extension.");
    process.exitCode = 1;
    return;
  }

  try {
    const extracted = extractDocxFormatting(inputPath);
    const checks = checkApaFormatting(extracted);
    const passed = checks.filter((check) => check.status === "pass").length;
    const failed = checks.filter((check) => check.status === "fail").length;
    const review = checks.filter((check) => check.status === "review").length;

    const report = {
      file: path.resolve(inputPath),
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

    console.log(JSON.stringify(report, null, 2));
    printReadableReport(report, { verbose });
  } catch (error) {
    console.error(`Failed to analyze document: ${error.message}`);
    process.exitCode = 1;
  }
}

main();

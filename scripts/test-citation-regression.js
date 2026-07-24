#!/usr/bin/env node
// Regression fixture for citation/reference matching determinism.
//
// Asserts exact per-check counts for two sample papers against frozen
// snapshots, and re-runs the full pipeline N times per fixture to assert the
// result never varies. Originally added because a beta tester saw citation
// counts differ between two runs of BC_Test_3_July_10.docx; investigation
// found the checker itself deterministic (see notes/citation-determinism.md).
// testingperfect.docx was added as a clean-paper counterpart so a change that
// only causes false positives on well-formed papers doesn't slip through.
//
// Some checks (e.g. "Citation alphabetical order", "Citation title format")
// are only emitted when the paper contains an applicable case, so the two
// fixtures do not have identical check sets — the comparison below is keyed
// by rule name per fixture, not by array position or a shared rule list.
//
// "Reference link verification" is deliberately excluded from every snapshot:
// it depends on live network calls and is only ever added by the browser
// entry point (src/browser/apaBrowser.js), which this script does not
// exercise — checkApaFormatting() never produces it. The exclusion below is
// a defensive no-op today, kept so this fixture stays stable if that ever
// changes.

const path = require("path");
const fs = require("fs");
const { extractDocxFormatting } = require("../src/docx/extractDocxFormatting");
const { checkApaFormatting } = require("../src/checks/checkApaFormatting");

const EXCLUDED_RULES = new Set(["Reference link verification"]);
const DETERMINISM_RUNS = 10;

const FIXTURES = [
  {
    docx: path.join(__dirname, "..", "samples", "BC_Test_3_July_10.docx"),
    snapshot: path.join(__dirname, "__fixtures__", "citation-postfix.json"),
  },
  {
    docx: path.join(__dirname, "..", "samples", "testingperfect.docx"),
    snapshot: path.join(__dirname, "__fixtures__", "testingperfect-postfix.json"),
  },
];

function snapshotOf(checks) {
  return checks
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
}

function runOnce(docxPath) {
  const extracted = extractDocxFormatting(docxPath);
  const { checks } = checkApaFormatting(extracted);
  return snapshotOf(checks);
}

function checkFixture(fixture) {
  const { docx, snapshot } = fixture;
  const label = path.basename(docx);
  let failures = 0;

  if (!fs.existsSync(docx)) {
    console.error(`Fixture file not found: ${docx}`);
    return 1;
  }
  if (!fs.existsSync(snapshot)) {
    console.error(`Expected snapshot not found: ${snapshot}`);
    console.error(`Run: node scripts/capture-citation-snapshot.js ${path.relative(process.cwd(), docx)} ${path.relative(process.cwd(), snapshot)}`);
    return 1;
  }

  const expected = JSON.parse(fs.readFileSync(snapshot, "utf8"));
  const expectedByRule = new Map(expected.map((c) => [c.rule, c]));

  // 1. Exact-match assertion against the frozen snapshot (keyed by rule name,
  //    so a check present in one fixture's snapshot but not the other's is fine).
  const actual = runOnce(docx);
  const actualByRule = new Map(actual.map((c) => [c.rule, c]));

  for (const [rule, exp] of expectedByRule) {
    const act = actualByRule.get(rule);
    if (!act) {
      console.error(`[${label}] MISSING CHECK: "${rule}" is in the snapshot but was not produced.`);
      failures += 1;
      continue;
    }
    const expStr = JSON.stringify(exp);
    const actStr = JSON.stringify(act);
    if (expStr !== actStr) {
      console.error(`[${label}] MISMATCH: "${rule}"`);
      console.error(`  expected: ${expStr}`);
      console.error(`  actual:   ${actStr}`);
      failures += 1;
    }
  }
  for (const rule of actualByRule.keys()) {
    if (!expectedByRule.has(rule)) {
      console.error(`[${label}] UNEXPECTED NEW CHECK: "${rule}" was produced but is not in the snapshot.`);
      failures += 1;
    }
  }

  // 2. Determinism assertion: re-run the full pipeline N times, assert identical output.
  const runs = [actual, ...Array.from({ length: DETERMINISM_RUNS - 1 }, () => runOnce(docx))];
  const distinct = new Set(runs.map((r) => JSON.stringify(r)));
  if (distinct.size !== 1) {
    console.error(`[${label}] NONDETERMINISM DETECTED: ${DETERMINISM_RUNS} runs produced ${distinct.size} distinct results.`);
    failures += 1;
  }

  if (failures === 0) {
    console.log(`OK [${label}]: ${expected.length} checks match the frozen snapshot; ${DETERMINISM_RUNS}/${DETERMINISM_RUNS} runs identical.`);
  }

  return failures;
}

function main() {
  let totalFailures = 0;
  for (const fixture of FIXTURES) {
    totalFailures += checkFixture(fixture);
  }

  if (totalFailures > 0) {
    console.error(`\n${totalFailures} regression failure(s).`);
    process.exitCode = 1;
  }
}

main();

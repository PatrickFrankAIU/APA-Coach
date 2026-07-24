#!/usr/bin/env node
// Regression fixture for citation/reference matching determinism.
//
// Asserts exact per-check counts for samples/BC_Test_3_July_10.docx against a
// frozen snapshot (scripts/__fixtures__/citation-postfix.json), and re-runs the
// full pipeline N times to assert the result never varies. This fixture exists
// because a beta tester saw the citation counts differ between two runs of this
// file; investigation found the checker itself deterministic (identical input
// always produces identical output — see notes/citation-determinism.md), so this
// guards that property going forward.

const path = require("path");
const fs = require("fs");
const { extractDocxFormatting } = require("../src/docx/extractDocxFormatting");
const { checkApaFormatting } = require("../src/checks/checkApaFormatting");

const FIXTURE_DOCX = path.join(__dirname, "..", "samples", "BC_Test_3_July_10.docx");
const SNAPSHOT_PATH = path.join(__dirname, "__fixtures__", "citation-postfix.json");
const DETERMINISM_RUNS = 10;

function snapshotOf(checks) {
  return checks
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

function runOnce() {
  const extracted = extractDocxFormatting(FIXTURE_DOCX);
  const { checks } = checkApaFormatting(extracted);
  return snapshotOf(checks);
}

function main() {
  if (!fs.existsSync(FIXTURE_DOCX)) {
    console.error(`Fixture file not found: ${FIXTURE_DOCX}`);
    process.exitCode = 1;
    return;
  }
  if (!fs.existsSync(SNAPSHOT_PATH)) {
    console.error(`Expected snapshot not found: ${SNAPSHOT_PATH}`);
    console.error("Run: node scripts/capture-citation-snapshot.js samples/BC_Test_3_July_10.docx scripts/__fixtures__/citation-postfix.json");
    process.exitCode = 1;
    return;
  }

  const expected = JSON.parse(fs.readFileSync(SNAPSHOT_PATH, "utf8"));
  const expectedByRule = new Map(expected.map((c) => [c.rule, c]));

  let failures = 0;

  // 1. Exact-match assertion against the frozen snapshot.
  const actual = runOnce();
  const actualByRule = new Map(actual.map((c) => [c.rule, c]));

  for (const [rule, exp] of expectedByRule) {
    const act = actualByRule.get(rule);
    if (!act) {
      console.error(`MISSING CHECK: "${rule}" is in the snapshot but was not produced.`);
      failures += 1;
      continue;
    }
    const expStr = JSON.stringify(exp);
    const actStr = JSON.stringify(act);
    if (expStr !== actStr) {
      console.error(`MISMATCH: "${rule}"`);
      console.error(`  expected: ${expStr}`);
      console.error(`  actual:   ${actStr}`);
      failures += 1;
    }
  }
  for (const rule of actualByRule.keys()) {
    if (!expectedByRule.has(rule)) {
      console.error(`UNEXPECTED NEW CHECK: "${rule}" was produced but is not in the snapshot.`);
      failures += 1;
    }
  }

  // 2. Determinism assertion: re-run the full pipeline N times, assert identical output.
  const runs = [actual, ...Array.from({ length: DETERMINISM_RUNS - 1 }, () => runOnce())];
  const distinct = new Set(runs.map((r) => JSON.stringify(r)));
  if (distinct.size !== 1) {
    console.error(`NONDETERMINISM DETECTED: ${DETERMINISM_RUNS} runs produced ${distinct.size} distinct results.`);
    failures += 1;
  }

  if (failures > 0) {
    console.error(`\n${failures} regression failure(s).`);
    process.exitCode = 1;
    return;
  }

  console.log(`OK: ${expected.length} checks match the frozen snapshot; ${DETERMINISM_RUNS}/${DETERMINISM_RUNS} runs identical.`);
}

main();

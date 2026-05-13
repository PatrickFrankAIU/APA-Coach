# APA Coach Rules

- Do NOT use LLMs in application logic
- All APA checks must be deterministic
- Prefer simple, readable code for teaching purposes
- Each APA rule must be implemented as a separate function
- Output must be structured JSON before rendering

---

## How to add a new check

### 1. Write the check function in `src/checks/checkApaFormatting.js`

Every check is a plain function that receives the `extracted` object (the output of `extractDocxFormatting`) and returns a result object. The simplest checks look like this:

```js
function checkMyNewRule(extracted) {
  const rule = "My new rule";  // must match the label shown in the UI
  const expected = "APA expects ...";

  // Your logic here. Use extracted.paragraphs, extracted.margins, etc.
  const failed = false;
  const status = failed ? "fail" : "pass";

  return {
    rule,
    status,                       // "pass" | "fail" | "review"
    passed: status === "pass",
    expected,
    expectedText: expected,       // human-readable version of the expectation
    foundText: failed             // one sentence summarizing what was found
      ? "What was wrong."
      : "What looked correct.",
    applicable: 1,
    checked: 1,
    matched: failed ? 0 : 1,
    failed: failed ? 1 : 0,
    unknown: 0,
    found: status === "pass" ? "Detected OK" : "Issue detected",
    applicableParagraphs: 0,
    details: ["One or more diagnostic sentences shown in the expanded card."],
    howToFix: status === "fail" ? ["Step 1.", "Step 2."] : [],
    resources: status === "fail" ? [{ label: "Link text", url: "https://..." }] : [],
  };
}
```

**Key fields:**
- `status` — `"pass"` (green), `"fail"` (red), or `"review"` (orange/unverifiable)
- `howToFix` — step-by-step Word instructions; only shown when `status === "fail"`
- `resources` — help links; only shown when `status === "fail"`
- `applicableParagraphs` — number of paragraphs the check applied to (0 for document-level checks)

For checks that iterate over paragraphs, use the helper `getParagraphsByRole(extracted, role)` (roles: `"body"`, `"heading"`, `"titlePage"`, `"referencesEntry"`, `"blank"`). Use `finishCheck(...)` to build the return value — it computes `status`, `matched`, etc. automatically from your `failures` and `unknowns` arrays. Look at `checkAlignment` or `checkFirstLineIndents` for a worked example.

### 2. Add how-to-fix steps in `getHowToFix`

If your check can fail, add a branch to the `getHowToFix(rule)` function that returns an array of step-by-step fix instructions. Match on the same `rule` string you used in step 1.

### 3. Register the check in `checkApaFormatting`

Add your function call to the `checks` array inside `checkApaFormatting`:

```js
const checks = [
  // ... existing checks ...
  checkMyNewRule(extracted),   // ← add here
];
```

If your check only applies when a References page exists, add it inside the `referencesHeading ? [...]` block.

### 4. Export the function

Add it to the `module.exports` at the bottom of the file:

```js
module.exports = {
  // ... existing exports ...
  checkMyNewRule,
};
```

### 5. Add a case to the README checks table

Update the table in `README.md` with a row for your new check so users know what it covers.

---

### What `extracted` contains

The `extracted` object is produced by `src/docx/extractDocxFormatting.js`. The fields most useful for writing checks:

| Field | Description |
|---|---|
| `extracted.paragraphs` | Array of all paragraphs. Each has `.role`, `.text`, `.index`, and `.formatting` (see below). |
| `extracted.margins` | `{ top, right, bottom, left }` — each is a resolved value `{ value, source, known }` |
| `extracted.pageNumbering` | Header/page-number info: `{ defaultHeader, firstPageHeader, titlePgEnabled }` |

Each paragraph's `.formatting` object has resolved fields for common properties:

| Key | Unit | What it measures |
|---|---|---|
| `lineSpacing` | multiplier (2.0 = double) | Line spacing |
| `spaceBefore` / `spaceAfter` | points | Paragraph spacing before/after |
| `firstLineIndent` | inches | First-line indent |
| `alignment` | string (`"left"`, `"center"`, etc.) | Paragraph alignment |
| `font` | string | Font family name |
| `fontSize` | points | Font size |

Every resolved field is an object `{ value, source, known }`:
- `known: false` means the value couldn't be determined — treat as unverifiable, not a failure
- `source` is one of `"direct"`, `"style"`, `"inherited"`, `"default"`, `"inferredDefault"`

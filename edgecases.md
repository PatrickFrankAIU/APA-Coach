Thank you for the opportunity to test the APA Coach from a library perspective. For the test document, we used the attached CC by 4.0 article (see attribution below). We inserted several errors, many of which were successfully detected by APA Coach.

 

Following are the areas that it missed.

 

In-text citations:

Used the term inline – The APA Manual uses the term  in-text.
Overlooked missing commas in the in-text citations e.g., (Li 2024), (Liu 2024), (Tang 2024), (Yang 2024)
Corrections required: (Li, 2024), (Liu, 2024), (Tang, 2024), (Yang, 2024),
Although it detected the missing period after et al., it did not detect the missing comma e.g., (Schwartz et al 2022), (Sloan et al 2024), (Dwivedi et al 2023)
Corrections required: (Schwartz et al., 2022), (Sloan et al., 2024), (Dwivedi et al., 2023).
Overlooked a comma we inserted after a date in the test doc i.e., Dempere et al (2023,)
Correction required: Dempere et al. (2023)
Overlooked the numbering that preceded headings in the test doc, e.g., 1 Ethical use of ChatGPT in education—Best practices to combat AI-induced plagiarism
Overlooked the use of sentence case instead of title case and bold print in the headings e.g., 1 Ethical use of ChatGPT in education—Best practices to combat AI-induced plagiarism
Example of corrections required: Ethical Use of ChatGPT in Education—Best Practices to Combat AI-Induced Plagiarism
Did not point out that all the headings labelled 1, 2, 3, 4, should have been level one headings
Example of corrections required:
Bold, Title Case, and Centered

Did not point out that the headings labelled 3.1, 3.2, 3.3, 3.4, 3.5 should have been level 2 headings (Bold, Title Case, against the Left Margin) e.g., point number 3.1 Current educational strategies to counter unethical use of LLMs
Example of corrections required:
Current Educational Strategies to Counter Unethical Use of LLMs

Did not point out that all headings labelled 3.1.1, 3.1.2, 3.1.3, 3.1.4, 3.1.5, 3.1.6 and 3.5.1, 3.5.2, 3.5.3, 3.5.4 should have been level 3 headings (Bold, Title Case, against the left margin) e.g., point number 3.1.1 Regulating AI usage within curricula
Example of corrections required:
Regulating AI Usage Within Curricula

Stated that all 14 items cited in-text had not been cited in the references.
All 14 items were included in the references.
Example of note required: All the in-text citations contained formatting errors.
Stated that 11 items listed in the reference had not been cited in-text.
All items listed in the references were cited in-text
Example of note required: All the in-text citations contained formatting errors.
 

References:

Did not point out that the references started immediately after (on the same page as) the text
Example of instruction required: References must start on a new clean page
Did not point out that all the references were numbered
Example of instruction required: The APA Manual does not use numbers or bullets in the references.
Stated that no DOIs were included in the references. Eleven of the 14  references contained DOIs
However, should have noted that all 11 DOIs were incorrectly formatted.
Example of instruction: All DOIs must be preceded by https://doi.org/  e.g., instead of 10.1007/s10639-024-12871-0 should have been https://doi.org/10.1007/s10639-024-12871-0
Stated that most of the items used title case instead of sentence case for the titles. I could only find 1 item that had that error.
Overlooked the use of lower case in the first letter of the first word of every subtitle in the references e.g., ChatGPT adoption and its influence on faculty well-being: an empirical research in higher education
Example of corrections required: ChatGPT adoption and its influence on faculty well-being: An empirical research in higher education.  
Stated that italics were wrongly used in the references. I could not find any titles that were italicized.
In the test doc, only journal titles were (correctly) italicized
Did not detect the use of “available at” in some references.
Example of instruction required: The APA Manual does not use that phrase “available at.”
Did not detect that some reference used “accessed” and a date.
Example of instruction required: The APA Manual does not use “accessed.”

---

## Fixes Applied

All 19 items were resolved across four implementation slices (branch `fix/bctest-edgecases`).

### Slice A — Quick detectors

**Item 1 — “inline” should be “in-text”**
Renamed the rule label from “Inline citations” to “In-text citations” throughout: `checkApaFormatting.js`, `main.jsx` (`CHECK_CATEGORY` map and UI label text), and `README.md` checks table. Internal function names (`checkInlineCitations`) were left unchanged.

**Items 2, 3, 4 — Missing and stray commas in citations**
Added `checkCitationComma` in `checkApaFormatting.js`. The check scans parenthetical citations for a missing comma between the author group and the year (`(Li 2024)` → `(Li, 2024)`), handles the `et al` (no period) case (`(Schwartz et al 2022)`), and also catches a trailing comma after the year (`Dempere et al (2023,)`). Registered in `checkApaFormatting`, exported, and added to `CHECK_CATEGORY` in `main.jsx` and `README.md`.

**Item 14 — Bare DOIs not detected / format not flagged**
Updated `hasDOIOrURL` in `checkApaFormatting.js` to match bare DOIs (`\b10\.\d{4,}/\S+`) so `checkReferenceDOIs` no longer falsely reports missing DOIs. Extended `checkReferenceDOIFormat` to flag bare DOIs as misformatted (must be preceded by `https://doi.org/`). The two checks interact intentionally: DOI/URL passes (DOI is present), DOI format fails (DOI needs the full URL prefix).

**Items 18, 19 — “Available at” and “accessed [date]” phrases**
Added `checkReferenceForbiddenPhrases` in `checkApaFormatting.js`. Scans reference entries for `available at` (case-insensitive) and `accessed` followed by a date. Registered, exported, added to `CHECK_CATEGORY` in `main.jsx` and `README.md`.

---

### Slice B — Reference-parser hardening

**Items 10, 11 — All 14 refs falsely uncited / all 11 citations falsely unmatched**
Root cause: `extractReferenceKey` in `checkApaFormatting.js` could not recover a clean surname from mangled author strings like `CambFierroJ.` or `LiM.`. Fixed by matching a lowercase→uppercase boundary regex (`/^(.*?[a-zà-ÿ])(?=[A-ZÀ-Ÿ]\.)/`) against the raw `beforeYear` string (retaining the trailing period so the lookahead fires correctly). This allowed all 14 reference keys and all 12 citation keys to match, eliminating all 26 false reports.

**Item 15 — Title case over-reported (~10 false positives, ~1 real)**
Rewrote the title-case detection logic in `checkReferenceTitleCapitalization` in `checkApaFormatting.js`. Changed the heuristic from “any two capitalized words” to “two or more *consecutive* capitalized words” and added a skip for words that follow sentence-ending punctuation (`.`, `?`, `!`). Reduced false positives from ~10 to ~1 real case.

**Item 16 — Lowercase subtitle after colon not flagged**
Extended `checkReferenceTitleCapitalization` to detect a lowercase first word after a colon in the title (e.g., `well-being: an empirical` → should be `An`). The check looks for `: [a-z]` patterns in the parsed title span.

**Item 17 — Italics falsely flagged on article titles**
Root cause: `parseReferenceEntry` in `checkApaFormatting.js` was splitting the title from the journal name at `. ` (period-space), but U+00A0 (non-breaking space) between the title and journal name caused the split to miss, pulling the journal name into the title span. Fixed by changing the split pattern from `/\. | \./` to `/\.\s|\s\./` so `\s` matches non-breaking space. Once the boundary was correct, `getSpanItalicState` found no italic article titles and all five false positives cleared.

---

### Slice C — Heading overhaul

**Item 5 — Heading numbering not flagged**
Added `checkHeadingNumbering` in `checkApaFormatting.js`. Flags headings that carry a leading section number (`1`, `3.1`, `3.1.1`, …). Gated on `hasNumberedHeadings` so the check only runs when numbered headings are actually present, ensuring zero impact on properly styled documents.

**Items 6, 7, 8, 9 — Sentence case, missing bold, wrong level/alignment**
Added `checkHeadingTitleCase`, `checkHeadingBold`, and `checkHeadingLevelFormat` in `checkApaFormatting.js`. These checks rely on a new `isNumberedHeadingText` detector and `parseHeadingNumber` helper added to `extractDocxFormatting.js`, which recognize paragraphs beginning with a section number (`^\d+(\.\d+)*\s+`) as headings even when no Word heading style is applied. The numbering depth determines the APA level: one part = L1, two parts = L2, three parts = L3. `checkHeadingLevelFormat` flags L1 headings that are not centered and L2/L3 headings that are not flush left. All four checks registered, exported, added to `CHECK_CATEGORY` in `main.jsx` and `README.md`.

---

### Slice D — Structural reference checks

**Item 12 — References not on a new page**
Added `checkReferencesStartNewPage` in `checkApaFormatting.js`. Checks for a hard page break before the References heading by inspecting (a) `pageBreakBefore` in the heading's own paragraph properties, and (b) `endsWithPageBreak` on the immediately preceding paragraph (a `w:br w:type=”page”` run). Both signals are surfaced as new per-paragraph fields (`pageBreakBefore`, `endsWithPageBreak`) extracted in `extractDocxFormatting.js`. To avoid false positives when References is inside a Word content control (sdt) — where `collectParagraphNodes` does not preserve the interleave order between `body.p` and `body.sdt` elements — the check returns `review` rather than `pass` or `fail` when either the heading or the preceding paragraph came from an sdt. Documents using Word's built-in bibliography tool will see a `review` result with a note explaining the limitation and advising against the built-in tool. A `fromSdt` flag was added to each paragraph in `extractDocxFormatting.js` to support this distinction.

**Item 13 — References are numbered**
Added `checkReferencesNumbered` in `checkApaFormatting.js`. Flags reference entry paragraphs that carry a `w:numPr` list-numbering element (automatic Word list numbering or bullets). The `numPr: { numId, ilvl }` signal is surfaced as a new per-paragraph field extracted in `extractDocxFormatting.js`. BCTestArticle's references carried `numId=22`, confirmed by raw XML inspection before implementation. Both new checks registered, exported, added to `CHECK_CATEGORY` in `main.jsx` and `README.md`.

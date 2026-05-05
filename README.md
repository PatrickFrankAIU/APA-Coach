# APA Coach

A client-side tool for checking APA 7 formatting in Word documents. Upload a `.docx` file and get an instant report. Runs entirely within the browser — the student's paper never leaves their device.

> **Beta:** APA Coach is in active development. Checks may miss issues or flag things incorrectly. Always review your paper manually before submitting.

Contact pfrank@aiuniv.edu with questions, comments, or bug reports.

Licensed under the [GNU General Public License v3.0](LICENSE). You are free to use, modify, and distribute this software, but any distributed version — including modified versions — must also be released under the GPL v3.0 with source code available. It may not be used as the basis for proprietary closed-source software.

---

## What it does

APA Coach reads the formatting metadata inside a `.docx` file and checks it against APA 7th edition requirements. It produces a report organized into three categories:

- **Required fixes** — formatting errors that should be corrected before submission
- **Optional review** — items that may be correct but couldn't be fully verified from the file alone
- **Passed** — items that meet APA expectations

Each result card explains what was found, what APA expects, and (for failures) step-by-step instructions for fixing the issue in Microsoft Word.

---

## Checks

| Check | What it verifies |
|---|---|
| Page numbering | Plain page number in the upper-right header on every page, starting at 1 on the title page; no "Page" or "Pg" label |
| Title page | Detects a standard APA title page at the start of the document |
| References page | Detects a References page near the end of the document |
| Inline citations | At least one (Author, Year) citation found in the body |
| References heading alignment | "References" heading is centered |
| References formatting | Hanging indents on reference entries; flags broken or bare-URL entries |
| Uncited references | Every reference entry has a matching inline citation |
| Unmatched citations | Every inline citation has a matching reference entry |
| Reference DOI/URL | Each reference includes a visible DOI or URL |
| Reference short link | Reference URLs link to a specific page, not just a domain homepage |
| Unapproved source | References do not use sources on [AIU's list of 145+ unapproved domains](https://careered.libguides.com/AIUS/unacceptablewebsites) |
| Margins | 1-inch margins on all four sides |
| Body line spacing | Double spacing throughout body paragraphs |
| Heading line spacing | Double spacing on heading paragraphs |
| References line spacing | Double spacing on reference entries |
| Body paragraph spacing | 0 pt before and after body paragraphs |
| Heading paragraph spacing | 0 pt before and after headings |
| Body first-line indents | 0.5-inch first-line indent on body paragraphs |
| Body alignment | Left alignment on body text |
| Font | 12pt Times New Roman throughout; flags mixed sizes or families |

---

## How it works

APA Coach runs entirely in the browser using three layers:

1. **Extraction** (`src/docx/extractDocxFormatting.js`) — A `.docx` file is a ZIP archive containing XML. [JSZip](https://stuk.github.io/jszip/) unpacks it and [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) reads `word/document.xml`, `word/styles.xml`, and the header XML files referenced in `word/_rels/document.xml.rels`. The extractor walks every paragraph and resolves its formatting (margins, spacing, indents, alignment) by merging direct formatting, applied styles, and Word defaults into a single resolved value with a known source.

2. **Checking** (`src/checks/checkApaFormatting.js`) — Each APA rule is implemented as a separate deterministic function. Checks produce a structured result object with a status (`fail`, `review`, or `pass`), human-readable found/expected text, diagnostic details, and how-to-fix steps. No AI or heuristic guessing — only values that can be read from the file are evaluated; anything else is flagged as unverifiable rather than assumed.

3. **UI** (`src/browser/main.jsx`) — A React interface renders the structured report. Results are grouped by status, with color-coded badges and expandable fix instructions.

---

## Privacy

Files are processed locally in your browser using the [File API](https://developer.mozilla.org/en-US/docs/Web/API/File_API). Nothing is uploaded, stored, or transmitted. Closing the tab clears everything.

---

## Running locally

```bash
npm install
npm run dev
```

Then open `http://localhost:5173/APA-Coach/` and upload a `.docx` file.

To build for production:

```bash
npm run build
```

### CLI

A Node script is available for testing from the command line:

```bash
node scripts/analyze-docx.js path/to/file.docx
node scripts/analyze-docx.js path/to/file.docx --verbose
```

The CLI prints a JSON report followed by a human-readable summary. `--verbose` includes paragraph-level counts and diagnostic details.

---

## Tech stack

- [Vite](https://vitejs.dev/) + [React](https://react.dev/) — build tooling and UI
- [Node.js](https://nodejs.org/) — runtime for the CLI script and local development
- [JSZip](https://stuk.github.io/jszip/) — in-browser `.docx` unpacking
- [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) — XML parsing
- [adm-zip](https://github.com/cthackers/adm-zip) — `.docx` unpacking in the Node CLI

---

## Project rules

See [AGENTS.md](AGENTS.md). The short version: no LLMs in application logic, all checks are deterministic, each APA rule is a separate function, output is structured JSON before rendering.

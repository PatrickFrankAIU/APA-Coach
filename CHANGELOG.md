# Changelog

## [0.5.0] - 2026-05-03

### Added

- **Page numbering check** — Detects whether a plain page number appears in the document header. Verifies that the default header contains a `PAGE` field, that the title-page header also shows a number when a separate first-page header is configured, and that no "Page" or "Pg" label precedes the number. Handles page number fields wrapped in Word's inline content-control (`sdt`) elements as well as literal numbers on first-page headers.
- **Beta notice banner** — Amber warning banner added to the app info card reminding users that APA Coach is in active development and that manual review is still required.

## [0.4.0] - 2026-05-02

### Added

- **Unapproved source check** — New red fail card that detects references linking to any of 145 domains on AIU's list of websites not approved for academic use (e.g., Wikipedia, Chegg, Course Hero, Quora, ChatGPT-writing tools). Checks both visible URLs in reference text and hidden hyperlinks.

### Fixed

**Citation extraction**
- Fixed narrative two-author citations ("Stallings and Brown (2018)") being extracted as only the last name before the year ("Brown"). The narrative regex now consumes "and CoAuthor" segments and anchors on the first author.
- Fixed missing space before "et al." ("Nurdiyantoet al.") — the "et" concatenated onto the name is now stripped when the text immediately after the name is "al."
- Fixed author initials glued to the last name in references without a separating comma ("Laudon K.C." → "Laudon"), preventing false "uncited reference" hits.
- Fixed title-as-author parenthetical citations ("Enterprise portal, n.d.") being truncated to only the first capitalized word. Author/title extraction now takes everything before the year rather than requiring each word to start with a capital letter.
- Fixed citation key extraction index offset bug: when a multi-citation parenthetical (e.g., "Smith, 2023; Jones, 2024") is split on semicolons, segments with a leading space had their year index applied to the trimmed string, corrupting the extracted author name ("Jones, 2" instead of "Jones"). Year regex now matches on the already-trimmed segment.
- Fixed years in their own Word run being silently dropped. `fast-xml-parser` parses `<w:t>2026</w:t>` as a JavaScript number when the content is purely numeric; `getRunText` now converts numeric text nodes to strings.

**Citation matching**
- Added note to Unmatched citations how-to-fix explaining that "et al." citations must match the first author's last name in the reference entry.
- Suppressed redundant "Uncited references" and "Unmatched citations" yellow cards when no inline citations are found — the "Inline citations" red fail card already covers this case.

## [0.3.0] - 2026-05-02

### Changed

**Visual examples on fail cards**
- Added a visual example to the References formatting card showing a hanging indent
- Added a before/after visual to the References heading alignment card (left vs. centered)
- Added a before/after visual to the Body first-line indents card
- Added a before/after visual to the References line spacing card (single vs. double)
- Added a before/after visual to the Body line spacing card (single vs. double)
- Added a before/after visual to the Heading line spacing card using headings and body text
- Added a before/after visual to the Body paragraph spacing card (extra vs. no spacing)
- Added a before/after visual to the Heading paragraph spacing card using headings and body text
- Added a three-way margin diagram to the Margins card (too narrow / too wide / correct)
- Added a centered title page mockup to the Title page card

**Additional help links**
- References heading alignment: Microsoft Support (Align text), GeeksforGeeks (Text Alignment in MS Word)
- References formatting: Microsoft Support (hanging indent), Scribbr (hanging indent)
- References page: Scribbr (APA Reference Pages)
- References line spacing: Microsoft Support (Change the line spacing in Word)
- Body first-line indents: Microsoft Support (Indent the first line of a paragraph)
- Body paragraph spacing: Microsoft Support (Change spacing between paragraphs)
- Heading paragraph spacing: Microsoft Support (Change spacing between paragraphs)
- Margins: Microsoft Support (Change margins)
- Title page: APA Style (Title Page Setup)

**Text and formatting fixes**
- Fixed "References heading alignment" how-to-fix steps being swallowed by the generic alignment rule
- Updated how-to-fix steps for References heading alignment to be clearer
- Fixed number formatting to always show at least one decimal place (e.g. "1.0x" instead of "1x")
- Fixed "All heading paragraphs use…" to read "All headings use…"
- Fixed broken sentence in summarizeAllValueGroups after paragraph/paragraphs wording update
- Added "paragraph"/"paragraphs" before verb in value group summaries for clarity

**App info card**
- Added GitHub repository link above the "Please report" line
- AIU logo now links to aiuniv.edu
- Shortened "Please report any problems with this app to" to "Please report any problems to"

## [0.2.0] - 2026-05-02

### Changed

- Added an application information card with the version number, last-updated date, and support contact.
- Moved the application information card to the bottom of the page after a report appears.
- Added AIU branding to the app header.

## [0.1.0] — 2026-05-01

Initial release of APA Coach.

### Features

**Core analysis**
- Client-side `.docx` analysis — files never leave the browser
- Node CLI script (`scripts/analyze-docx.js`) for local testing with `--verbose` flag
- Structured JSON report output before rendering

**APA checks implemented**
- Title page detection
- Margins (1 inch on all sides)
- Body line spacing (double)
- Heading line spacing (double)
- References line spacing (double)
- Body paragraph spacing (0 pt before/after)
- Heading paragraph spacing (0 pt before/after)
- Body first-line indents (0.5 inch)
- Body alignment (left)
- References page detection
- References heading alignment (centered)
- References entry formatting (hanging indents, consistency, broken entries)

**Report UI**
- Summary cards with pass/fail/review counts
- Fail, Review, and Pass sections with section headers and intro banners
- Color-coded status badges (red/amber/green)
- Expandable "How to fix" instructions on each failed check
- Smooth-scroll navigation from summary cards to sections
- "↑ Back to top" links at the end of each section

**How to fix guidance**
- Step-by-step Word instructions for every failing check
- Visual cues for toolbar buttons (e.g., ⇵☰ for Line Spacing)
- Hanging indent instructions via both the Paragraph dialog and the ruler

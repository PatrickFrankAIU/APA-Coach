# Changelog

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

# APA Coach Coverage Assessment
_Last updated: July 3, 2026 (v1.4.1)_

## Well Covered

- **Document formatting** — margins, font, line spacing, paragraph spacing, indents, alignment, page numbering
- **Headings** — bold, title case, no numbering, level alignment (centered vs. left)
- **In-text citations** — ampersand, et al., comma, n.d., page format, multiple sources, year suffix, secondary citations, personal communications
- **Reference content** — authors, year, title capitalization, italics (with journal name abbreviation note), punctuation (volume/issue/page/article number formats), DOI format, forbidden phrases, unapproved sources, short links
- **Reference structure** — hanging indent, new page, heading alignment, no numbering, citation/reference matching, alphabetical order _(added v1.4.1)_

---

## Notable Gaps

### High value, relatively straightforward to add

1. **Abstract** — APA requires one for many papers (150–250 words, no first-line indent, single paragraph, double-spaced). We don't check for its presence or format at all.

2. **Title page completeness** — We detect whether a title page exists but don't verify required fields: title (centered, bold, title case), author name, institutional affiliation, course name and number, instructor name, assignment due date. Field detection is fuzzy (lines aren't labeled) so false positive risk needs care.

### Moderate effort, sample file needed first

3. **Block quotes** — Direct quotes of 40+ words must be formatted as an indented block (0.5") without quotation marks. Detection is feasible if we can see how Word encodes block quotes in a sample; semantic detection (knowing whether un-indented long text *should* be a block quote) is not viable.

4. **Tables and figures** — APA has specific rules for table titles (bold, above the table), figure captions (plain text, below the figure), and notes. Requires parsing the DOCX table XML structure, which is substantially different from paragraph parsing. A sample with correctly and incorrectly formatted tables would significantly reduce implementation complexity.

### Higher complexity, lower priority

5. **Heading levels 3–5** — We verify levels 1 (centered bold) and 2 (left-aligned bold) only. Levels 3 (bold italic, flush left), 4 (bold italic, indented, run-in), and 5 (italic, indented, run-in) are unchecked. Requires confident heading-level identification, which is currently heuristic.

6. **Specific reference types in depth** — Book chapters (editor format, page range), theses/dissertations, conference papers, government reports each have type-specific APA rules beyond what we currently verify. Our reference classifier would need substantial hardening before type-specific rules could be reliably applied.

7. **Abbreviated journal title detection** — We added an instructional note in v1.4.1 but don't detect abbreviations. Heuristic (words ending in period mid-string) works for standard multi-word abbreviations like "Int. J. Inf. Manag." but misses Frontiers-style colon-format entries and abbreviations without periods.

const EXPECTED_MARGIN_INCHES = 1;
const EXPECTED_LINE_SPACING = 2;
const EXPECTED_PARAGRAPH_SPACING_POINTS = 0;
const EXPECTED_FIRST_LINE_INDENT_INCHES = 0.5;
const EXPECTED_ALIGNMENT = "left";
const TOLERANCE = 0.01;
const MISSING_HANGING_INDENT_ISSUE = "One or more references are missing a 0.5-inch hanging indent.";
const BARE_URL_ISSUE = "One or more entries appear to be bare URLs with no author, title, or year. APA references must include the author, publication year, title, and source — a URL alone is not a complete reference.";
const HANGING_INDENT_RESOURCE = {
  label: "Microsoft Support: Create a hanging indent in Word",
  url: "https://support.microsoft.com/en-us/office/create-a-hanging-indent-in-word-7bdfb86a-c714-41a8-ac7a-3782a91ccad5",
};
const HANGING_INDENT_SCRIBBR_RESOURCE = {
  label: "Hanging Indent Instructions (Scribbr)",
  url: "https://www.scribbr.com/citing-sources/hanging-indent/",
};
const MARGINS_MSFT_RESOURCE = {
  label: "Microsoft Support: Change margins",
  url: "https://support.microsoft.com/en-us/office/change-margins-da21a474-99d8-4e54-b12d-a8a14ea7ce02",
};
const PARA_SPACING_MSFT_RESOURCE = {
  label: "Microsoft Support: Change spacing between paragraphs",
  url: "https://support.microsoft.com/en-us/office/change-spacing-between-paragraphs-ee4c7016-7cb8-405e-90a1-6601e657f3ce",
};
const LINE_SPACING_MSFT_RESOURCE = {
  label: "Microsoft Support: Change the line spacing in Word",
  url: "https://support.microsoft.com/en-us/office/change-the-line-spacing-in-word-04ada056-b8ef-4b84-87dd-5d7c28a85712",
};
const FIRST_LINE_INDENT_MSFT_RESOURCE = {
  label: "Microsoft Support: Indent the first line of a paragraph",
  url: "https://support.microsoft.com/en-us/office/indent-the-first-line-of-a-paragraph-b3721167-e1c8-40c3-8a97-3f046fc72d6d",
};
const APA_DOIS_URLS_RESOURCE = {
  label: "APA Style DOIs and URLs",
  url: "https://apastyle.apa.org/style-grammar-guidelines/references/dois-urls",
};
const APA_REFERENCES_PAGE_SCRIBBR_RESOURCE = {
  label: "Scribbr Guide to APA Reference Pages",
  url: "https://www.scribbr.com/apa-style/apa-reference-page/",
};
const APA_TITLE_PAGE_RESOURCE = {
  label: "APA Title Page Setup",
  url: "https://apastyle.apa.org/style-grammar-guidelines/paper-format/title-page",
};
const PAGE_NUMBER_MSFT_RESOURCE = {
  label: "Microsoft Support: Insert page numbers",
  url: "https://support.microsoft.com/en-us/office/insert-page-numbers-9f366518-0500-4b45-903d-987d3827c007",
};
const ALIGN_TEXT_MSFT_RESOURCE = {
  label: "Microsoft Support: Align text",
  url: "https://support.microsoft.com/en-us/office/align-text-left-or-right-center-text-or-justify-text-on-a-page-70da744d-0f4d-472e-916d-1c42d94dc33f",
};
const ALIGN_TEXT_GFG_RESOURCE = {
  label: "GeeksforGeeks: Text Alignment in MS Word",
  url: "https://www.geeksforgeeks.org/ms-word/text-alignment-in-ms-word/",
};
const INLINE_CITATIONS_APA_RESOURCE = {
  label: "APA Style: In-text citations",
  url: "https://apastyle.apa.org/style-grammar-guidelines/citations",
};
const INLINE_CITATIONS_OWL_RESOURCE = {
  label: "Purdue OWL: In-Text Citations — The Basics",
  url: "https://owl.purdue.edu/owl/research_and_citation/apa_style/apa_formatting_and_style_guide/in_text_citations_the_basics.html",
};

const APA_REFERENCE_FORMAT_RESOURCE = {
  label: "APA Style: Reference Examples",
  url: "https://apastyle.apa.org/style-grammar-guidelines/references/examples",
};
const APA_AUTHORS_RESOURCE = {
  label: "APA Style: Author Format",
  url: "https://apastyle.apa.org/style-grammar-guidelines/references/author-format",
};
const APA_DOI_FORMAT_RESOURCE = {
  label: "APA Style: DOIs and URLs",
  url: "https://apastyle.apa.org/style-grammar-guidelines/references/dois-urls",
};
const APA_IN_TEXT_FORMAT_RESOURCE = {
  label: "APA Style: In-Text Citation Format",
  url: "https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles",
};

const UNAPPROVED_DOMAINS = [
  "123helpme.com","a1-termpaper.com","academic-papers.blogspot.com","academized.com",
  "activant.in","advancedwriters.com","affordablepapers.com","algebra.com",
  "allessayhelp.net","allinterview.com","answerbag.com","answers.com",
  "antiessays.com","antistudy.com","articlesbase.com","articlesfactory.com",
  "ask.com","askanewquestion.com","askmehelpdesk.com","assignmenthelppro.co",
  "bartleby.com","bestessays.com","besttermpaper.com","blurtit.com",
  "bookrags.com","brainly.com","brainmass.com","buenastareas.com",
  "bukisa.com","business-bliss.com","business-plan.com","cheathouse.com",
  "chegg.com","chuckiii.com","cliffsnotes.com","collegetermpapers.com",
  "copy.ai","coursehero.com","cram.com","customessaymeister.com",
  "customwriting.com","customwritings.com","cyberessays.com","docshare.com",
  "domypapers.com","dreamessays.com","eddusaver.com","eduessays.com",
  "edugenie.net","effectivepapers.com","ehow.com","eliteacademicessays.com",
  "enotes.com","essay24.com","essayailab.com","essaycapital.com",
  "essayexperts.com","essayhub.com","essaykitchen.com","essayok.net",
  "essaypro.com","essaysforstudent.com","essaytown.com","examcollection.com",
  "exampleessays.com","exclusivepapers.com","fastpapers.com","foxessays.com",
  "freeadvice.com","freeessay.com","freeonlineresearchpapers.com","gotessays.com",
  "homework.ecrater.com","homework.study.com","how2pass.com","hubpages.com",
  "indiacelebrating.com","informationheadquarter.com","investopedia.com","ipl.org",
  "ivythesis.typepad.com","javacodee.blogspot.com","justanswer.com","keenessays.com",
  "lawbirdie.com","lawteacher.net","lotsofessays.com","majortests.com",
  "markedbyteachers.com","mastersthesiswriting.com","megaessays.com","midterm.us",
  "myassignmenthelp.com","myspace.com","mysupergeek.com","netmba.com",
  "otherpapers.com","papercamp.com","paperhelp.org","paperial.com",
  "paperlap.com","papersowl.com","phd-dissertations.com","premium-papers.com",
  "prepmypaper.com","prowritershub.com","quickessayrelief.com","quizlet.com",
  "quora.com","radessays.com","rentacoder.com","researchpaper.com",
  "researchpapers247.com","reviewessays.com","rushtermpapers.com","scienceline.ucsb.edu",
  "slickpapers.com","smodin.io","stackoverflow.co","statisticsbrain.com",
  "studentshare.org","studocu.com","studybay.com","studybounty.com",
  "studymode.com","studymoose.com","superiorpapers.com","sweetstudy.com",
  "termpaperaccess.com","termpaperwarehouse.com","thebalance.com","thekumachan.com",
  "transtutors.com","tutorsonspot.com","ukessays.com","wedoyouressay.com",
  "wikibooks.org","wikipedia.org","wikiversity.org","wisegeek.com",
  "wordpress.com","writeessaytoday.com","writemyessays.net","writework.com",
  "zeepedia.com",
];

function isClose(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= TOLERANCE;
}

function formatNumber(value, suffix) {
  return typeof value === "number" ? `${value.toFixed(2).replace(/0$/, "")}${suffix}` : "not found";
}

function formatSource(source) {
  if (source === "inferredDefault") {
    return "inferred default";
  }

  return source;
}

function formatStudentSource(source) {
  if (source === "direct") {
    return "set directly in the document";
  }

  if (source === "style") {
    return "from a Word style";
  }

  if (source === "inherited") {
    return "inherited from a Word style";
  }

  if (source === "default") {
    return "from Word's default formatting";
  }

  if (source === "inferredDefault") {
    return "from Word's built-in default formatting";
  }

  return "from an unknown source";
}

function formatResolved(resolved, suffix) {
  return `${formatNumber(resolved.value, suffix)} (${formatSource(resolved.source)})`;
}

function formatStudentResolved(resolved, suffix, label) {
  if (!resolved.known) {
    return `unknown ${label}`;
  }

  const value = typeof resolved.value === "number" ? formatNumber(resolved.value, suffix) : resolved.value;
  return `${value} ${label} ${formatStudentSource(resolved.source)}`;
}

function formatPlainValue(resolved, suffix) {
  if (!resolved.known) {
    return "unknown";
  }

  const source = formatStudentSource(resolved.source);
  return `${formatNumber(resolved.value, suffix)} ${source}`;
}

function formatSpacingValueWithSource(resolved) {
  if (!resolved.known) {
    return "unknown";
  }

  return `${formatNumber(resolved.value, " pt")} ${formatStudentSource(resolved.source)}`;
}

function summarizeParagraphs(paragraphs) {
  const shown = paragraphs.slice(0, 5).map((paragraph) => paragraph.index).join(", ");
  const extra = paragraphs.length > 5 ? `, and ${paragraphs.length - 5} more` : "";

  return `paragraphs ${shown}${extra}`;
}

function finishCheck(rule, expectedText, foundText, applicableParagraphs, failures, unknowns, details, howToFix = [], resources = []) {
  const checked = applicableParagraphs.length - unknowns.length;
  const matched = checked - failures.length;
  let status = "pass";
  if (applicableParagraphs.length === 0) {
    status = "review";
  } else if (failures.length > 0) {
    status = "fail";
  } else if (unknowns.length > 0) {
    status = "review";
  }

  return {
    rule,
    status,
    passed: status === "pass",
    expected: expectedText,
    expectedText,
    foundText,
    applicable: applicableParagraphs.length,
    checked,
    matched,
    failed: failures.length,
    unknown: unknowns.length,
    found:
      applicableParagraphs.length === 0
        ? "No applicable paragraphs found"
        : `Applicable: ${applicableParagraphs.length}, checked: ${checked}, matched: ${matched}, unknown: ${unknowns.length}`,
    applicableParagraphs: applicableParagraphs.length,
    details: applicableParagraphs.length === 0 ? ["No applicable paragraphs were found for this check."] : details,
    howToFix: status === "fail" ? howToFix : [],
    resources: status === "fail" ? resources : [],
  };
}

function groupByDescription(items, describe) {
  const groups = new Map();

  for (const item of items) {
    const description = describe(item);
    if (!groups.has(description)) {
      groups.set(description, []);
    }

    groups.get(description).push(item);
  }

  return Array.from(groups.entries()).map(([description, paragraphs]) => {
    return `${description} for ${paragraphs.length} ${paragraphs.length === 1 ? "paragraph" : "paragraphs"} (${summarizeParagraphs(
      paragraphs,
    )}).`;
  });
}

function getParagraphsByRole(extracted, role) {
  return extracted.paragraphs.filter((paragraph) => paragraph.role === role);
}

// Returns 'italic' | 'not-italic' | 'mixed' | 'not-found'.
// Searches for searchText inside the paragraph's run array (attached by Phase 1).
function getSpanItalicState(runs, searchText) {
  if (!runs || runs.length === 0 || !searchText || searchText.trim().length === 0) return "not-found";
  const fullText = runs.map((r) => r.text).join("");
  const idx = fullText.indexOf(searchText);
  if (idx === -1) return "not-found";
  const end = idx + searchText.length;
  let pos = 0;
  const states = [];
  for (const run of runs) {
    const runEnd = pos + run.text.length;
    if (run.text.length > 0 && runEnd > idx && pos < end) {
      states.push(run.italic);
    }
    pos = runEnd;
  }
  if (states.length === 0) return "not-found";
  if (states.every((s) => s)) return "italic";
  if (states.every((s) => !s)) return "not-italic";
  return "mixed";
}

// Parses the source part of a journal reference into its components.
function parseJournalSourcePart(sourcePart) {
  if (!sourcePart) return null;
  const m = sourcePart.match(/^(.*?),\s*(\d+)\s*(\(\s*\d+\s*\))?\s*(?:,\s*(.+))?$/);
  if (!m) return null;
  return {
    journalName: m[1].trim(),
    volume: m[2],
    issueWithParens: m[3] ? m[3].trim() : null,
    issue: m[3] ? m[3].replace(/[()]/g, "").trim() : null,
    pages: m[4] ? m[4].replace(/\.\s*$/, "").trim() : null,
  };
}

function getBodyText(extracted, referencesHeading) {
  return extracted.paragraphs
    .filter((p) => p.role !== "blank" && (!referencesHeading || p.index < referencesHeading.index))
    .map((p) => p.text)
    .join("\n");
}

function splitUnknowns(paragraphs, fieldName) {
  return paragraphs.filter((paragraph) => !paragraph.formatting[fieldName].known);
}

function summarizeValueGroups(paragraphs, fieldName, suffix, valueLabel) {
  const groups = new Map();

  for (const paragraph of paragraphs) {
    const resolved = paragraph.formatting[fieldName];
    const key = `${resolved.value}|${resolved.source}`;
    if (!groups.has(key)) {
      groups.set(key, {
        count: 0,
        text: formatStudentResolved(resolved, suffix, valueLabel),
      });
    }

    groups.get(key).count += 1;
  }

  return Array.from(groups.values())
    .map((group) => `${group.count} ${group.count === 1 ? "paragraph uses" : "paragraphs use"} ${group.text}`)
    .join("; ");
}

function summarizeAllValueGroups(paragraphs, fieldName, suffix, valueLabel) {
  const grouped = summarizeValueGroups(paragraphs, fieldName, suffix, valueLabel);
  return grouped.replace(/\b\d+ paragraph uses /g, "").replace(/\b\d+ paragraphs use /g, "");
}

function summarizeKnownCheck(label, expected, applicableParagraphs, failures, unknowns, fieldName, suffix, valueLabel) {
  if (applicableParagraphs.length === 0) {
    return `APA Coach did not find any ${label.toLowerCase()} to check.`;
  }

  const matched = applicableParagraphs.length - failures.length - unknowns.length;
  const lowerLabel = label.toLowerCase();

  if (failures.length > 0 && unknowns.length > 0) {
    if (matched === 0) {
      return `No readable ${lowerLabel} match APA style. ${summarizeValueGroups(
        failures,
        fieldName,
        suffix,
        valueLabel,
      )}. APA Coach could not determine the value for ${unknowns.length} ${
        unknowns.length === 1 ? "paragraph" : "paragraphs"
      }.`;
    }

    return `${matched} of ${applicableParagraphs.length} ${lowerLabel} match APA style. ${summarizeValueGroups(
      failures,
      fieldName,
      suffix,
      valueLabel,
    )}. APA Coach could not determine the value for ${unknowns.length} ${
      unknowns.length === 1 ? "paragraph" : "paragraphs"
    }.`;
  }

  if (failures.length > 0) {
    if (matched === 0) {
      if (failures.length === applicableParagraphs.length) {
        return `All ${lowerLabel} use ${summarizeAllValueGroups(failures, fieldName, suffix, valueLabel)}.`;
      }

      return `All readable ${lowerLabel} use non-APA values: ${summarizeValueGroups(
        failures,
        fieldName,
        suffix,
        valueLabel,
      )}.`;
    }

    return `${matched} of ${applicableParagraphs.length} ${lowerLabel} match APA style. ${summarizeValueGroups(
      failures,
      fieldName,
      suffix,
      valueLabel,
    )}.`;
  }

  if (unknowns.length > 0) {
    if (matched === 0) {
      return `APA Coach could not determine this formatting for all ${lowerLabel}.`;
    }

    return `${matched} of ${applicableParagraphs.length} ${lowerLabel} match APA style. APA Coach could not determine the value for ${unknowns.length} ${
      unknowns.length === 1 ? "paragraph" : "paragraphs"
    }.`;
  }

  return `All ${lowerLabel} use ${expected}.`;
}

function getHowToFix(rule) {
  if (rule === "Unconverted markup symbols") {
    return [
      "These symbols are markdown formatting codes — they work in apps like Notion or GitHub but not in Word.",
      "Remove the asterisks or underscores and apply formatting directly in Word instead.",
      "To italicize: select the text, delete the surrounding * symbols, then press Ctrl+I (Cmd+I on Mac).",
      "To bold: select the text, delete the surrounding ** symbols, then press Ctrl+B (Cmd+B on Mac).",
    ];
  }

  if (rule === "Unapproved source") {
    return [
      "This source is on AIU's list of websites not approved for academic use.",
      "Replace it with a peer-reviewed article, textbook, or credible institutional source.",
      "Search your library database (ProQuest, EBSCOhost) for a credible alternative on the same topic.",
    ];
  }

  if (rule === "Font") {
    return [
      "Select all text in the document (Ctrl+A / Cmd+A).",
      "In the Home toolbar, set the font to Times New Roman and size to 12.",
      "Check headings: APA headings should also be 12pt (bold or bold-italic, not larger).",
      "Avoid mixing fonts — use one font family throughout the entire document.",
    ];
  }

  if (rule === "Title page") {
    return [
      "Remove the existing cover page or title-page template.",
      "Create a simple APA title page at the beginning of the document.",
      "Center the title page text.",
      "Include the title, author, course, instructor, and date.",
    ];
  }

  if (rule === "Margins") {
    return [
      "In Microsoft Word, open Layout > Margins.",
      "Choose Normal, or set Top, Bottom, Left, and Right to 1 inch.",
    ];
  }

  if (rule.includes("line spacing")) {
    return [
      "Select the affected text in Microsoft Word.",
      "Click the Line Spacing button (⇵☰) in the Home toolbar.",
      "Choose 2.0 for double spacing.",
    ];
  }

  if (rule.includes("paragraph spacing")) {
    return [
      "Select the affected text in Microsoft Word.",
      "Click the Line Spacing button (⇵☰) in the Home toolbar to open its menu.",
      "Click 'Line Spacing Options...'",
      "In the Paragraph window, find the Spacing section: set Before to 0 pt and After to 0 pt.",
      "Click OK.",
    ];
  }

  if (rule.includes("first-line indents")) {
    return [
      "Select the affected body paragraphs in Microsoft Word.",
      "Open Home > Paragraph settings.",
      "Under Indentation, set Special to First line and By to 0.5 inch.",
    ];
  }

  if (rule.includes("alignment") && rule !== "References heading alignment") {
    return [
      "Select the affected body text in Microsoft Word.",
      "Open Home > Paragraph.",
      "Choose Align Left.",
    ];
  }

  if (rule === "In-text citations") {
    return [
      "Add an in-text citation every time you use information from a source.",
      "Parenthetical format: place (Author, Year) at the end of the sentence before the period.",
      "Narrative format: use Author (Year) at the start of the sentence.",
    ];
  }

  if (rule === "Citation comma") {
    return [
      "APA citations require a comma between the author name and the year.",
      "Parenthetical: (Li, 2024) not (Li 2024); (Smith et al., 2022) not (Smith et al 2022).",
      "Narrative: Smith et al. (2023) not Smith et al. (2023,) — no comma after the year.",
    ];
  }

  if (rule === "Uncited references") {
    return [
      "For each reference listed, find where you used that source and add an in-text citation.",
      "Parenthetical format: place (Author, Year) at the end of the sentence before the period.",
      "Narrative format: use Author (Year) at the start of the sentence.",
      "Note: abbreviated organization names (e.g., APA) may not be detected automatically — check those manually.",
    ];
  }

  if (rule === "Unmatched citations") {
    return [
      "For each citation listed, add a matching entry to your References page.",
      "Each APA reference starts with the author's last name and the publication year.",
      "Use a citation generator or your library database to format the full reference entry.",
      "Note: if your citation uses 'et al.' (e.g., Smith et al., 2022), make sure the reference entry starts with the same first author's last name.",
      "Note: abbreviated organization names (e.g., APA) may not be detected automatically — check those manually.",
    ];
  }

  if (rule === "Reference short link") {
    return [
      "Go to the website and navigate to the specific article or page you used.",
      "Copy the full URL from your browser's address bar — it should include a path after the domain.",
      "Replace the short domain link in your reference with the full URL.",
      "If you can't find the original page, use a search engine to locate the article and copy its direct URL.",
    ];
  }

  if (rule === "Reference DOI/URL") {
    return [
      "For each reference listed, find and add its DOI or URL.",
      "DOI format: https://doi.org/10.xxxx/xxxxx — find it on the article's page or via doi.org.",
      "If no DOI exists, add the URL of the publisher's page or database where you accessed the source.",
      "Note: some older books and sources do not have a DOI or URL — those may be flagged incorrectly.",
    ];
  }

  if (rule === "Page numbering") {
    return [
      "In Microsoft Word, double-click near the top of the page to open the Header.",
      "Press Tab twice to move the cursor to the right margin (the Header has built-in center and right-flush tab stops).",
      "Go to Insert > Page Number > Current Position > Plain Number.",
      "Make sure page numbering starts at 1 on the title page.",
      "Do not type 'Page' or 'Pg' before the number — use a plain number only.",
    ];
  }

  if (rule === "References page") {
    return [
      "Add a new page at the end of the document titled 'References'.",
      "Center the word 'References' at the top of the page.",
      "List all sources in alphabetical order.",
      "Use a hanging indent for each reference entry.",
    ];
  }

  if (rule === "References heading alignment") {
    return [
      "Select the \"References\" heading in Microsoft Word.",
      "In the Home toolbar, click the Center Align button.",
    ];
  }

  if (rule === "Reference authors") {
    return [
      "List author names in Last, F. M. format — last name first, then initials with a period after each initial.",
      "Use only initials, not full first or middle names: write \"Smith, J. A.\" not \"Smith, Jane Ann\".",
      "Each initial must be followed by a period and a space: \"J. R.\" not \"JR\" or \"J R\".",
      "Last names use normal title case — capitalize only the first letter: \"Hallett\" not \"HALLETT\".",
      "Separate two authors with \", & \" — use \"&\" not \"and\".",
      "For 3–20 authors, list all of them separated by commas, with \"&\" before the last.",
      "For 21+ authors, list the first 19, then an ellipsis (…), then the last author.",
      "Do not use \"et al.\" in a reference entry — \"et al.\" is only for in-text citations.",
    ];
  }

  if (rule === "Reference title capitalization") {
    return [
      "Reference titles use sentence case: capitalize only the first word, the first word after a colon, and proper nouns.",
      "Example — correct: \"The effects of social media on academic performance\"",
      "Example — incorrect: \"The Effects of Social Media on Academic Performance\"",
      "Proper nouns (names of people, places, organizations, brands) stay capitalized: \"Facebook\", \"United States\", \"COVID-19\".",
      "Note: This check may flag titles containing correctly capitalized proper nouns (e.g., \"African American\"), and may miss titles where every word is capitalized with no lowercase words in between. When in doubt, review manually.",
    ];
  }

  if (rule === "Reference year format") {
    return [
      "Place the publication year in parentheses immediately after the author(s): Smith, J. A. (2020).",
      "Follow the closing parenthesis with a period: (2020). Title…",
      "If there is no date, write (n.d.) — lowercase, with periods after each letter.",
    ];
  }

  if (rule === "Reference italics") {
    return [
      "Book, report, and standalone work titles should be italicized.",
      "Journal article and book chapter titles should NOT be italicized — remove any italic formatting.",
      "For journal articles: italicize only the journal name and volume number. The issue number in parentheses and page range should not be italic.",
      "In Word, select the text and press Ctrl+I (Cmd+I on Mac) to toggle italic.",
    ];
  }

  if (rule === "Reference punctuation") {
    return [
      "Each reference must end with the URL or DOI on its last line.",
      "Journal page ranges do not use \"p.\" or \"pp.\" — write the page range directly: 45–67, not pp. 45–67.",
      "Write volume and issue with no space: 15(3) not 15 (3).",
      "Use an en dash (–) for page ranges, not a hyphen (-).",
    ];
  }

  if (rule === "Reference DOI format") {
    return [
      "DOIs must be formatted as full URLs: https://doi.org/10.xxxx/xxxxx",
      "A bare DOI (e.g., 10.1007/xxxxx) must be preceded by https://doi.org/ — write https://doi.org/10.1007/xxxxx",
      "Do not use the old \"doi:\" prefix — replace it with \"https://doi.org/\".",
      "Do not use \"http://dx.doi.org/\" — use \"https://doi.org/\" instead.",
    ];
  }

  if (rule === "Reference forbidden phrases") {
    return [
      "APA does not use \"Available at\" before a URL — write the URL directly without that phrase.",
      "APA does not require an access date for most sources — remove \"accessed\" and the date that follows.",
      "Example: instead of \"Available at: https://example.com (accessed July 10, 2024)\" write just \"https://example.com\"",
    ];
  }

  if (rule === "Citation ampersand") {
    return [
      "Inside parenthetical citations, use \"&\" between author names: (Smith & Jones, 2020).",
      "In narrative citations (outside parentheses), spell out \"and\": Smith and Jones (2020) found…",
    ];
  }

  if (rule === "Citation et al. format") {
    return [
      "Write \"et al.\" with a period after \"al\" and no period after \"et\": et al.",
      "Do not write \"et. al.\", \"et al\", \"etal\", or \"et. al\".",
      "\"et al.\" is used in citations when a work has 3 or more authors.",
    ];
  }

  if (rule === "Citation no-date format") {
    return [
      "When a source has no date, write \"n.d.\" — lowercase, with a period after each letter.",
      "Do not write \"N.D.\", \"n.d\", \"nd\", or \"no date\".",
      "Example: (Smith, n.d.) or Smith (n.d.).",
    ];
  }

  if (rule === "Citation page format") {
    return [
      "Use \"p.\" for a single page and \"pp.\" for a page range in citations.",
      "Do not use \"pg.\" or \"pgs.\" — these are not APA format.",
      "Example: (Smith, 2020, p. 45) or (Smith, 2020, pp. 45–47).",
    ];
  }

  if (rule === "Citation multiple sources") {
    return [
      "When citing multiple sources in one parenthetical, separate them with a semicolon: (Smith, 2020; Jones, 2021).",
      "List multiple sources in alphabetical order by the first author's last name.",
    ];
  }

  if (rule === "Citation year suffix") {
    return [
      "When you cite two works by the same author(s) published in the same year, add a letter suffix: (Smith, 2020a) and (Smith, 2020b).",
      "Make sure the same suffix appears in both the in-text citation and the reference entry.",
    ];
  }

  if (rule === "Personal communication") {
    return [
      "Personal communications (interviews, emails, conversations) are cited in the text only — do not add them to the References list.",
      "Format: (F. M. Last, personal communication, Month Day, Year) — use the person's full initials and last name.",
      "Example: (J. A. Smith, personal communication, March 15, 2024).",
    ];
  }

  if (rule === "Reference hanging indent") {
    return [
      "If any entries are bare URLs: replace each URL-only line with a full APA reference. A complete reference includes the author's last name and initials, publication year in parentheses, title of the work, and the URL or DOI.",
      "Example: Smith, J. A. (2023). Title of article. Site Name. https://www.example.com/article",
      "A hanging indent means the first line is flush left and all following lines are indented 0.5\". References must also be double-spaced.",
      "Select all reference entries in Microsoft Word.",
      "Click the Line Spacing button (⇵☰) in the Home toolbar to open its menu.",
      "Click 'Line Spacing Options...'",
      "In the Paragraph window: under Indentation, set Special to Hanging and By to 0.5\". Under Spacing, set Line spacing to Double.",
      "Click OK.",
    ];
  }

  return [];
}

function checkPageNumbering(extracted) {
  const rule = "Page numbering";
  const expected = "APA requires a plain page number in the upper right corner of every page, starting with 1 on the title page.";
  const expectedText = expected;

  const pn = extracted.pageNumbering;
  if (!pn) {
    return {
      rule, status: "review", passed: false,
      expected, expectedText,
      foundText: "APA Coach could not check page numbering in this document.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "Page numbering unknown",
      applicableParagraphs: 0,
      details: ["Page numbering information could not be extracted from this document."],
      howToFix: [], resources: [],
    };
  }

  const issues = [];
  const details = [];

  if (!pn.defaultHeader || !pn.defaultHeader.hasNumber) {
    issues.push("No page number was found in the document header. APA requires a page number in the upper right corner of every page.");
  }

  if (pn.titlePgEnabled && (!pn.firstPageHeader || !pn.firstPageHeader.hasNumber)) {
    issues.push("The title page header is missing a page number. Page numbering must start at 1 on the title page.");
  }

  const headersToCheck = [pn.defaultHeader, pn.titlePgEnabled ? pn.firstPageHeader : null].filter(Boolean);
  if (headersToCheck.some((h) => h.hasPageLabel)) {
    issues.push('The page number should be a plain number only — do not include "Page" or "Pg" before the number.');
  }

  if (pn.defaultHeader && pn.defaultHeader.hasNumber && !pn.defaultHeader.tabBeforeNumber) {
    details.push(
      "The page number may not be positioned in the upper right corner. In Word, press Tab twice in the Header to reach the right-flush position before inserting the page number.",
    );
  }

  const status = issues.length > 0 ? "fail" : "pass";

  return {
    rule, status, passed: status === "pass",
    expected, expectedText,
    foundText: status === "pass"
      ? "Page numbers appear in the header on all pages, starting at 1 on the title page."
      : issues.join(" "),
    applicable: 1, checked: 1,
    matched: status === "pass" ? 1 : 0,
    failed: status === "fail" ? 1 : 0,
    unknown: 0,
    found: status === "pass" ? "Page numbering detected" : "Page numbering issue detected",
    applicableParagraphs: 0,
    details: issues.concat(details),
    howToFix: status === "fail" ? getHowToFix(rule) : [],
    resources: status === "fail" ? [PAGE_NUMBER_MSFT_RESOURCE] : [],
  };
}

function checkTitlePage(extracted) {
  const titlePageParagraphs = getParagraphsByRole(extracted, "titlePage");
  const firstTextParagraph = extracted.paragraphs.find((paragraph) => paragraph.role !== "blank");
  const beginsWithBodyText = firstTextParagraph ? firstTextParagraph.role === "body" : false;
  const status = titlePageParagraphs.length === 0 && beginsWithBodyText ? "fail" : "pass";

  return {
    rule: "Title page",
    status,
    passed: status === "pass",
    expected: "APA expects a title page with centered text including title, author, course, instructor, and date.",
    expectedText: "APA expects a title page with centered text including title, author, course, instructor, and date.",
    foundText:
      status === "fail"
        ? "APA Coach did not detect a standard APA title page at the beginning of the document."
        : "APA Coach detected possible title page text at the beginning of the document.",
    applicable: 1,
    checked: 1,
    matched: status === "pass" ? 1 : 0,
    failed: status === "fail" ? 1 : 0,
    unknown: 0,
    found: status === "pass" ? "Title page detected" : "Title page not detected",
    applicableParagraphs: titlePageParagraphs.length,
    details:
      status === "fail"
        ? ["No titlePage paragraphs were detected before the first body paragraph."]
        : [`Detected ${titlePageParagraphs.length} possible title page paragraph(s).`],
    howToFix: status === "fail" ? getHowToFix("Title page") : [],
    resources: status === "fail" ? [APA_TITLE_PAGE_RESOURCE] : [],
  };
}

function summarizeParagraphSpacingFoundText(label, applicableParagraphs, failures) {
  const lowerLabel = label.toLowerCase();
  const isHeading = label === "Heading";

  if (applicableParagraphs.length === 0) {
    return `APA Coach did not find any ${lowerLabel} paragraphs to check.`;
  }

  if (failures.length === 0) {
    return `All ${lowerLabel} paragraphs have 0 pt before and 0 pt after.`;
  }

  return isHeading
    ? "Extra spacing around headings was detected. Some headings use 'Spacing Before' or 'Spacing After,' which is not allowed in APA."
    : "Extra spacing between paragraphs was detected. Some paragraphs use 'Spacing Before' or 'Spacing After,' which is not allowed in APA.";
}

function summarizeParagraphSpacingWithUnknowns(label, applicableParagraphs, failures, partialUnknowns) {
  const isHeading = label === "Heading";

  if (failures.length > 0) {
    return isHeading
      ? "Extra spacing around headings was detected. Some headings use 'Spacing Before' or 'Spacing After,' which is not allowed in APA."
      : "Extra spacing between paragraphs was detected. Some paragraphs use 'Spacing Before' or 'Spacing After,' which is not allowed in APA.";
  }

  return isHeading
    ? "Heading paragraph spacing could not be fully verified. This may need a quick check in Word."
    : "Paragraph spacing could not be fully verified. This may need a quick check in Word.";
}

function summarizeMarginsForStudents(entries, failures, unknowns) {
  if (failures.length === 0 && unknowns.length === 0) {
    return "APA Coach found 1-inch margins on the top, right, bottom, and left sides.";
  }

  const incorrectText = failures.map(
    ([side, margin]) => `${side} is ${formatNumber(margin.value, " in")} ${formatStudentSource(margin.source)}`,
  );
  const correctText = entries
    .filter(([, margin]) => margin.known && isClose(margin.value, EXPECTED_MARGIN_INCHES))
    .map(([side]) => side);
  const unknownText = unknowns.map(([side]) => side);
  const parts = [];

  if (incorrectText.length > 0) {
    parts.push(`Incorrect margins: ${incorrectText.join("; ")}.`);
  }

  if (correctText.length > 0) {
    parts.push(`Correct margins: ${formatList(correctText)}.`);
  }

  if (unknownText.length > 0) {
    parts.push(`Unknown margins: ${formatList(unknownText)}.`);
  }

  return parts.join(" ");
}

function formatList(items) {
  if (items.length <= 1) {
    return items.join("");
  }

  if (items.length === 2) {
    return `${items[0]} and ${items[1]}`;
  }

  return `${items.slice(0, -1).join(", ")}, and ${items[items.length - 1]}`;
}

function checkMargins(extracted) {
  const margins = extracted.margins;
  const entries = [
    ["top", margins.topInches],
    ["right", margins.rightInches],
    ["bottom", margins.bottomInches],
    ["left", margins.leftInches],
  ];
  const unknowns = entries.filter(([, margin]) => !margin.known);
  const failures = entries.filter(([, margin]) => margin.known && !isClose(margin.value, EXPECTED_MARGIN_INCHES));
  const checked = entries.length - unknowns.length;
  const matched = checked - failures.length;
  const status = failures.length > 0 ? "fail" : unknowns.length > 0 ? "review" : "pass";

  return {
    rule: "Margins",
    status,
    passed: status === "pass",
    expected: "1 inch on all sides",
    expectedText: "APA expects 1-inch margins on all sides.",
    foundText: summarizeMarginsForStudents(entries, failures, unknowns),
    applicable: entries.length,
    checked,
    matched,
    failed: failures.length,
    unknown: unknowns.length,
    found: `Applicable: ${entries.length}, checked: ${checked}, matched: ${matched}, unknown: ${unknowns.length}`,
    details: failures
      .map(([side, margin]) => `${side} margin is ${formatNumber(margin.value, " in")} (${formatSource(margin.source)}).`)
      .concat(unknowns.map(([side]) => `${side} margin is unknown.`)),
    howToFix: status === "fail" ? getHowToFix("Margins") : [],
    resources: status === "fail" ? [MARGINS_MSFT_RESOURCE] : [],
  };
}

function checkLineSpacingForParagraphs(paragraphs, rule, label) {
  const unknowns = splitUnknowns(paragraphs, "lineSpacingMultiple");
  const failures = paragraphs.filter(
    (paragraph) =>
      paragraph.formatting.lineSpacingMultiple.known &&
      !isClose(paragraph.formatting.lineSpacingMultiple.value, EXPECTED_LINE_SPACING),
  );
  const details = [];

  if (unknowns.length > 0) {
    details.push(
      `Line spacing: Unknown for ${unknowns.length} ${label.toLowerCase()} ${
        unknowns.length === 1 ? "paragraph" : "paragraphs"
      }. APA Coach could not determine whether spacing is inherited from a style.`,
    );
  }

  details.push(
    ...groupByDescription(failures, (paragraph) => {
      return `Line spacing is ${formatResolved(paragraph.formatting.lineSpacingMultiple, "x")}`;
    }),
  );

  return finishCheck(
    rule,
    "APA expects double spacing.",
    summarizeKnownCheck(
      label === "Heading" ? "Headings" : `${label} paragraphs`,
      "double spacing",
      paragraphs,
      failures,
      unknowns,
      "lineSpacingMultiple",
      "",
      "spacing",
    ),
    paragraphs,
    failures,
    unknowns,
    details,
    getHowToFix(rule),
    [LINE_SPACING_MSFT_RESOURCE],
  );
}

function checkLineSpacingForRole(extracted, role, label) {
  return checkLineSpacingForParagraphs(getParagraphsByRole(extracted, role), `${label} line spacing`, label);
}

function checkParagraphSpacingForRole(extracted, role, label) {
  const applicableParagraphs = getParagraphsByRole(extracted, role);
  // Fail if any *known* side is non-zero — don't require both sides to be known.
  const failures = applicableParagraphs.filter((paragraph) => {
    const before = paragraph.formatting.spacingBeforePoints;
    const after = paragraph.formatting.spacingAfterPoints;
    return (before.known && !isClose(before.value, EXPECTED_PARAGRAPH_SPACING_POINTS)) ||
           (after.known && !isClose(after.value, EXPECTED_PARAGRAPH_SPACING_POINTS));
  });
  const failureSet = new Set(failures);
  // Unknowns: not already a failure, but at least one side can't be determined.
  // Exception: if one side is explicitly 0 (correct) and the other is simply unset
  // (source "unknown" with no explicit value), treat as passing. This covers the common
  // Word workflow of Ctrl+A → set "Space After" to 0, leaving "Space Before" blank.
  const partialUnknowns = applicableParagraphs.filter((paragraph) => {
    const before = paragraph.formatting.spacingBeforePoints;
    const after = paragraph.formatting.spacingAfterPoints;
    if (failureSet.has(paragraph)) return false;
    if (before.known && after.known) return false;
    // One side is unknown — pass if the other side is explicitly correct
    const knownSideCorrect =
      (before.known && isClose(before.value, EXPECTED_PARAGRAPH_SPACING_POINTS)) ||
      (after.known && isClose(after.value, EXPECTED_PARAGRAPH_SPACING_POINTS));
    return !knownSideCorrect;
  });
  const details = [];

  if (partialUnknowns.length > 0) {
    details.push(
      `Paragraph spacing: Some ${label.toLowerCase()} paragraphs have incomplete spacing information. APA Coach reports known before and after values separately when only one side can be determined.`,
    );
    details.push(
      ...groupByDescription(partialUnknowns, (paragraph) => {
        const before = formatPlainValue(paragraph.formatting.spacingBeforePoints, " pt");
        const after = formatPlainValue(paragraph.formatting.spacingAfterPoints, " pt");
        return `Paragraph spacing is before ${before}, after ${after}`;
      }),
    );
  }

  details.push(
    ...groupByDescription(failures, (paragraph) => {
      const before = formatResolved(paragraph.formatting.spacingBeforePoints, " pt");
      const after = formatResolved(paragraph.formatting.spacingAfterPoints, " pt");
      return `Paragraph spacing is before ${before}, after ${after}`;
    }),
  );

  const foundText =
    partialUnknowns.length > 0
      ? summarizeParagraphSpacingWithUnknowns(label, applicableParagraphs, failures, partialUnknowns)
      : summarizeParagraphSpacingFoundText(label, applicableParagraphs, failures);

  return finishCheck(
    `${label} paragraph spacing`,
    label === "Heading"
      ? "APA requires headings to have 0 pt before and 0 pt after."
      : "APA requires 0 pt before and 0 pt after paragraphs.",
    foundText,
    applicableParagraphs,
    failures,
    partialUnknowns,
    details,
    getHowToFix(`${label} paragraph spacing`),
    [PARA_SPACING_MSFT_RESOURCE],
  );
}

function looksLikeMisformattedHeading(paragraph) {
  const text = paragraph.text.trim();
  const bold = paragraph.runs && paragraph.runs.length > 0 && paragraph.runs.every(r => r.bold);
  const centered = paragraph.formatting.alignment.known && paragraph.formatting.alignment.value === "center";
  return (bold || centered) && text.length > 0 && text.length <= 100 && !/[.!?;:]$/.test(text);
}

// APA 7 Level 1 headings are centered and bold — correct alignment, just missing proper Word style.
function looksLikeApaLevel1Heading(paragraph) {
  const alignment = paragraph.formatting.alignment;
  return looksLikeMisformattedHeading(paragraph) &&
    alignment.known && alignment.value === "center";
}

// APA 7: abstract body text is flush left — no first-line indent required.
function isAbstractBodyParagraph(paragraph, allParagraphs) {
  const idx = allParagraphs.indexOf(paragraph);
  for (let i = idx - 1; i >= 0; i--) {
    const p = allParagraphs[i];
    if (p.text.trim() === "") continue;
    if (/^abstract$/i.test(p.text.trim())) return true;
    break;
  }
  return false;
}

function failingParaPreview(paragraph) {
  const t = paragraph.text.trim();
  const snippet = t.length > 80 ? t.slice(0, 80) + "…" : t;
  return looksLikeMisformattedHeading(paragraph) ? `Heading: "${snippet}"` : `Paragraph: "${snippet}"`;
}

function checkFirstLineIndents(extracted, referencesHeading) {
  const allParagraphs = extracted.paragraphs;
  const refParagraphs = new Set(referencesHeading ? getReferenceEntryParagraphs(allParagraphs, referencesHeading) : []);
  // Exclude: abstract body (no indent per APA 7), Level 1 headings (centered+bold: correctly no indent), reference entries (covered by hanging-indent check)
  const applicableParagraphs = getParagraphsByRole(extracted, "body").filter(
    (p) => !isAbstractBodyParagraph(p, allParagraphs) && !looksLikeMisformattedHeading(p) && !refParagraphs.has(p)
      && p.style.id !== "ListParagraph",
  );
  const unknowns = splitUnknowns(applicableParagraphs, "firstLineIndentInches");
  const checked = applicableParagraphs.length - unknowns.length;
  const failures = applicableParagraphs.filter((paragraph) => {
    const indent = paragraph.formatting.firstLineIndentInches;
    if (indent.known && isClose(indent.value, EXPECTED_FIRST_LINE_INDENT_INCHES)) return false;
    // A leading tab character satisfies APA's 0.5" indent requirement (APA 7, §2.23)
    if (/^\t/.test(paragraph.text)) return false;
    return indent.known;
  });
  const matched = checked - failures.length;
  const headingFailures = failures.filter(looksLikeMisformattedHeading);

  const details = [];

  if (unknowns.length > 0) {
    details.push(
      `First-line indent: Unknown for ${unknowns.length} body ${
        unknowns.length === 1 ? "paragraph" : "paragraphs"
      }. APA Coach could not determine whether indentation is inherited from a style.`,
    );
  }

  details.push(
    ...groupByDescription(failures, (paragraph) => {
      return `First-line indent is ${formatResolved(paragraph.formatting.firstLineIndentInches, " in")}`;
    }),
  );

  let status;
  let foundText;

  if (applicableParagraphs.length === 0) {
    status = "review";
    foundText = "APA Coach did not find any body paragraphs to check.";
  } else if (failures.length === 0 && unknowns.length === 0) {
    status = "pass";
    foundText = "All body paragraphs use a 0.5-inch first-line indent.";
  } else if (failures.length > 0) {
    status = "fail";
    const n = failures.length;
    if (headingFailures.length === failures.length) {
      foundText = `${n} section ${n === 1 ? "heading appears" : "headings appear"} to be formatted as body text without a heading style — these should use Word's built-in heading styles, not manual bold/center formatting.`;
    } else if (headingFailures.length > 0) {
      foundText = `${n} ${n === 1 ? "item is" : "items are"} missing the required 0.5-inch first-line indent, including ${headingFailures.length} that ${headingFailures.length === 1 ? "looks like a heading" : "look like headings"} formatted as body text.`;
    } else {
      foundText = `${n} body ${n === 1 ? "paragraph is" : "paragraphs are"} missing the required 0.5-inch first-line indent.`;
    }
  } else {
    status = "review";
    foundText = "These paragraphs appear correct, but may not be using Word's built-in first-line indent.";
  }

  const missingItems = failures.map(failingParaPreview);

  return {
    rule: "Body first-line indents",
    status,
    passed: status === "pass",
    expected: "APA requires a 0.5-inch first-line indent.",
    expectedText: "APA requires a 0.5-inch first-line indent.",
    foundText,
    applicable: applicableParagraphs.length,
    checked,
    matched,
    failed: failures.length,
    unknown: unknowns.length,
    found:
      applicableParagraphs.length === 0
        ? "No applicable paragraphs found"
        : `Applicable: ${applicableParagraphs.length}, checked: ${checked}, matched: ${matched}, unknown: ${unknowns.length}`,
    applicableParagraphs: applicableParagraphs.length,
    details: applicableParagraphs.length === 0 ? ["No applicable paragraphs were found for this check."] : details,
    howToFix: status === "fail" ? getHowToFix("Body first-line indents") : [],
    resources: status === "fail" ? [FIRST_LINE_INDENT_MSFT_RESOURCE] : [],
    missingItems: missingItems.length > 0 ? missingItems : [],
    missingItemsLabel: missingItems.length > 0 ? "Items without a first-line indent:" : "",
  };
}

function isBodyTitle(paragraph, position) {
  const text = paragraph.text.trim();
  const alignment = paragraph.formatting.alignment;
  const indent = paragraph.formatting.firstLineIndentInches;
  return (
    position <= 1 &&
    text.length > 0 &&
    text.length <= 140 &&
    alignment.known &&
    alignment.value === "center" &&
    (!indent.known || isClose(indent.value, 0))
  );
}

function checkAlignment(extracted) {
  const bodyParagraphs = getParagraphsByRole(extracted, "body");
  const applicableParagraphs = bodyParagraphs.filter(
    // APA 7 Level 1 headings are correctly centered+bold — exclude them from alignment failures.
    // isBodyTitle handles title-page centered text.
    (paragraph, position) => !isBodyTitle(paragraph, position) && !looksLikeApaLevel1Heading(paragraph),
  );
  const unknowns = splitUnknowns(applicableParagraphs, "alignment");
  const checked = applicableParagraphs.length - unknowns.length;
  const failures = applicableParagraphs.filter((paragraph) => {
    const alignment = paragraph.formatting.alignment;
    return alignment.known && alignment.value !== EXPECTED_ALIGNMENT;
  });
  const matched = checked - failures.length;
  const details = [];

  if (unknowns.length > 0) {
    details.push(
      `Alignment: Unknown for ${unknowns.length} body ${
        unknowns.length === 1 ? "paragraph" : "paragraphs"
      }. APA Coach could not determine whether alignment is inherited from a style.`,
    );
  }

  details.push(
    ...groupByDescription(failures, (paragraph) => {
      const alignment = paragraph.formatting.alignment;
      return `Alignment is ${alignment.value} (${formatSource(alignment.source)})`;
    }),
  );

  let status;
  let foundText;

  const headingFailures = failures.filter(looksLikeMisformattedHeading);

  if (applicableParagraphs.length === 0) {
    status = "review";
    foundText = "APA Coach did not find any body paragraphs to check.";
  } else if (failures.length === 0 && unknowns.length === 0) {
    status = "pass";
    foundText = "All body paragraphs use left alignment.";
  } else if (failures.length > 0) {
    status = "fail";
    const n = failures.length;
    if (headingFailures.length === failures.length) {
      foundText = `${n} section ${n === 1 ? "heading appears" : "headings appear"} to be formatted as body text with manual center alignment — these should use Word's built-in heading styles instead.`;
    } else if (headingFailures.length > 0) {
      foundText = `${n} body ${n === 1 ? "paragraph is" : "paragraphs are"} not left aligned, including ${headingFailures.length} that ${headingFailures.length === 1 ? "looks like a heading" : "look like headings"} formatted as body text.`;
    } else {
      foundText = `${n} body ${n === 1 ? "paragraph is" : "paragraphs are"} not left aligned. APA requires left alignment for all body text.`;
    }
  } else {
    status = "review";
    foundText = "These paragraphs appear correct, but alignment could not be fully verified.";
  }

  const missingItems = failures.map(failingParaPreview);

  return {
    rule: "Body alignment",
    status,
    passed: status === "pass",
    expected: "APA expects body text to be left aligned.",
    expectedText: "APA expects body text to be left aligned.",
    foundText,
    applicable: applicableParagraphs.length,
    checked,
    matched,
    failed: failures.length,
    unknown: unknowns.length,
    found:
      applicableParagraphs.length === 0
        ? "No applicable paragraphs found"
        : `Applicable: ${applicableParagraphs.length}, checked: ${checked}, matched: ${matched}, unknown: ${unknowns.length}`,
    applicableParagraphs: applicableParagraphs.length,
    details: applicableParagraphs.length === 0 ? ["No applicable paragraphs were found for this check."] : details,
    howToFix: status === "fail" ? getHowToFix("Body alignment") : [],
    missingItems: missingItems.length > 0 ? missingItems : [],
    missingItemsLabel: missingItems.length > 0 ? "Items not left aligned:" : "",
  };
}

function isReferencesHeading(paragraph) {
  const text = paragraph.text.trim().replace(/^[^a-zA-Z]+|[^a-zA-Z]+$/g, "").toLowerCase();
  return text === "references" || text === "reference";
}

function findReferencesHeadingNearEnd(paragraphs) {
  // Use last non-blank paragraph as ceiling to ignore trailing blank pages
  const nonBlank = paragraphs.filter((p) => p.text.trim().length > 0);
  if (nonBlank.length === 0) return null;
  const lastNonBlankIndex = nonBlank[nonBlank.length - 1].index;
  const threshold = Math.floor(lastNonBlankIndex * 0.6);

  // Look for an explicit "References" heading
  const heading = paragraphs.find((p) => p.index > threshold && isReferencesHeading(p));
  if (heading) return heading;

  // Fallback: find first run of reference-entry-style paragraphs after the threshold
  // (handles documents where the References heading is missing)
  const firstEntry = paragraphs.find((p) => p.index > threshold && looksLikeReferenceEntryStart(p));
  if (!firstEntry) return null;

  // Synthesize a virtual heading one index before the first entry so that
  // getReferenceEntryParagraphs (index > heading.index) captures all entries
  return { index: firstEntry.index - 1, text: "", synthetic: true };
}

function getReferenceEntryParagraphs(paragraphs, referencesHeading) {
  return paragraphs
    .filter((paragraph) => paragraph.index > referencesHeading.index)
    .filter((paragraph) => paragraph.text.trim().length > 0);
}

function looksLikeReferenceEntryStart(paragraph) {
  return /\((\d{4}[a-z]?|n\.d\.)[,)]/i.test(paragraph.text);
}

function extractReferenceKey(text) {
  const yearMatch = text.match(/\((\d{4}[a-z]?|n\.d\.)[,)]/);
  if (!yearMatch) return null;
  const year = yearMatch[1];
  const beforeYear = text.substring(0, yearMatch.index).trim().replace(/\.\s*$/, "").trim();
  const commaIdx = beforeYear.indexOf(",");
  // Strip trailing initials glued without comma (e.g., "Laudon K.C." → "Laudon")
  const lastName = (commaIdx > 0 ? beforeYear.substring(0, commaIdx) : beforeYear)
    .trim()
    .replace(/\s+[A-Z]\.(?:[A-Z]\.)*$/, "")
    .trim();
  if (!lastName || lastName.length < 2) return null;
  const preview = text.length > 80 ? text.substring(0, 80) + "…" : text;
  return { lastName, year, preview };
}

// Parses a reference paragraph into structural segments. Heuristic and deterministic —
// when confidence is low, kindGuess is "unknown" so downstream checks return "review".
function parseReferenceEntry(paragraph) {
  const text = (paragraph && paragraph.text ? paragraph.text : "").trim();
  const result = {
    raw: text,
    authorsRaw: null,
    year: null,
    title: null,
    sourcePart: null,
    doiOrUrl: null,
    kindGuess: "unknown",
  };

  if (!text) return result;

  const yearMatch = text.match(/\((\d{4}[a-z]?|n\.d\.)(?:,\s*[^)]+)?\)/);
  if (!yearMatch) return result;

  result.year = yearMatch[1];

  const beforeYear = text.slice(0, yearMatch.index).trim().replace(/[,;\s]+$/, "").trim();
  if (beforeYear) result.authorsRaw = beforeYear;

  let afterYear = text.slice(yearMatch.index + yearMatch[0].length).trim();
  afterYear = afterYear.replace(/^[.\s]+/, "");

  // Match doi: prefix before bare https:// so "doi:https://..." is captured whole
  const urlMatch = afterYear.match(/\bdoi:\s*\S+/i) || afterYear.match(/\bhttps?:\/\/\S+/i);
  if (urlMatch) {
    result.doiOrUrl = urlMatch[0].replace(/[.,;)]+$/, "");
    afterYear = (afterYear.slice(0, urlMatch.index) + afterYear.slice(urlMatch.index + urlMatch[0].length))
      .trim()
      .replace(/[.,;\s]+$/, "");
  }

  // Split title from source on ". " (normal) or " ." (student typo: space before period)
  const splitMatch = afterYear.match(/\. | \./);
  if (splitMatch) {
    result.title = afterYear.slice(0, splitMatch.index).trim();
    result.sourcePart = afterYear.slice(splitMatch.index + splitMatch[0].length).trim().replace(/\.\s*$/, "");
  } else {
    result.title = afterYear.replace(/\.\s*$/, "").trim();
  }

  if (!result.title) result.title = null;
  if (!result.sourcePart) result.sourcePart = null;

  result.kindGuess = classifyReferenceKind(result);
  return result;
}

function classifyReferenceKind(parsed) {
  const source = parsed.sourcePart || "";
  const title = parsed.title || "";

  // Journal article: "Volume(Issue), pp" or "Volume, pp" pattern in source
  if (/\b\d+\s*\(\s*\d+\s*\)/.test(source) || /,\s*\d+\s*(?:[–—-]\s*\d+)\b/.test(source)) {
    return "journal-article";
  }

  // Book chapter: "In Editor (Ed.), Book Title" or bare "In Book Title. Publisher."
  if (/\bIn\s+[^.]*\(Eds?\.\)/i.test(source) || /\bIn\s+[^.]*\(Eds?\.\)/i.test(title)) {
    return "book-chapter";
  }
  if (/^\s*In\s+[A-Z]/.test(source)) {
    return "book-chapter";
  }

  // Webpage: URL present, no source signals — but doi.org links are journal articles
  if (parsed.doiOrUrl && !source) {
    return /\bdoi\.org\/10\./i.test(parsed.doiOrUrl) ? "journal-article" : "webpage";
  }

  // Book: has a publisher-like tail and no journal numbers
  if (source && !/\d/.test(source)) {
    return "book";
  }

  // Standalone number only = page count for a report/book
  if (/^\d+$/.test(source.trim())) {
    return "book";
  }

  return "unknown";
}

function extractInlineCitationKeys(bodyText) {
  const keys = [];
  const seen = new Set();

  const parenRe = /\(([^()]{2,300})\)/g;
  let m;
  while ((m = parenRe.exec(bodyText)) !== null) {
    const content = m[1];
    if (!/\d{4}|n\.d\./.test(content)) continue;
    if (/\bpersonal\s+communication\b/i.test(content)) continue;
    if (/\bas\s+cited\s+in\b/i.test(content)) continue;
    const segments = content.split(";");
    for (const seg of segments) {
      const segTrimmed = seg.trim();
      const yearM = segTrimmed.match(/\b(\d{4}[a-z]?)\b/) || segTrimmed.match(/(n\.d\.)/);
      if (!yearM) continue;
      const year = yearM[1];
      // Extract everything before the year as the author/title portion
      const beforeYear = segTrimmed.slice(0, yearM.index).replace(/[,\s]+$/, "").trim();
      if (!beforeYear) continue;
      // Strip "et al." — also handles missing space (e.g., "Smithet al." → "Smith")
      const withoutEtAl = beforeYear.replace(/\s*et\s+al\.?$/i, "").trim();
      // For multi-author paren citations, take only the first author.
      // Split on "&"/"and" first (handles 2-author), then on the first comma
      // (handles 3+ authors like "A, B, & C" emitted by Word's references feature).
      const firstAuthorChunk = withoutEtAl.split(/\s*&\s*|\s+and\s+/i)[0].trim();
      const lastName = firstAuthorChunk.split(",")[0].trim() || firstAuthorChunk || withoutEtAl;
      const key = `${lastName.toLowerCase()}|${year.replace(/[a-z]$/, "")}`;
      if (!seen.has(key)) {
        seen.add(key);
        keys.push({ lastName, year, source: segTrimmed });
      }
    }
  }

  const narrativeRe = /\b([A-ZÀ-Ÿ][A-Za-zÀ-ÿ''-]+)(?:\s+(?:and|&)\s+[A-ZÀ-Ÿ][A-Za-zÀ-ÿ''-]+)*(?:\s+et\s+al\.?)?\s+\((\d{4}[a-z]?)\)/g;
  while ((m = narrativeRe.exec(bodyText)) !== null) {
    const lastName = m[1];
    const year = m[2];
    const key = `${lastName.toLowerCase()}|${year.replace(/[a-z]$/, "")}`;
    if (!seen.has(key)) {
      seen.add(key);
      keys.push({ lastName, year, source: `${lastName} (${year})` });
    }
  }

  return keys;
}

function isReferenceCited(key, bodyText, citationKeys) {
  const { lastName, year } = key;
  const escaped = lastName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const yearPat = year === "n.d." ? "n\\.d\\." : year.replace(/[a-z]$/, "") + "[a-z]?";
  const strict = new RegExp(`(?<![a-zA-Z])${escaped}(?:\\s+et\\s+al\\.?)?[^\\n.]{0,40}${yearPat}`, "i");
  if (strict.test(bodyText)) return true;
  // Fallback: name and year appear within 100 chars (catches typos like "toGrandNode(2020)")
  const nameIdx = bodyText.search(new RegExp(escaped, "i"));
  if (nameIdx !== -1) {
    const window = bodyText.substring(Math.max(0, nameIdx - 10), nameIdx + lastName.length + 100);
    if (new RegExp(yearPat).test(window)) return true;
  }
  // Fuzzy fallback: check extracted citation keys for a near-match (catches "Hallet" vs "HALLETT")
  if (citationKeys) {
    const yearBase = year.replace(/[a-z]$/, "");
    return citationKeys.some(
      (ck) => namesMatch(ck.lastName, lastName) && ck.year.replace(/[a-z]$/, "") === yearBase,
    );
  }
  return false;
}

function normalizeName(name) {
  return name.toLowerCase().replace(/[-‑‐]/g, "");
}

function editDistance(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, (_, i) => [i]);
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
    }
  }
  return dp[m][n];
}

function namesMatch(a, b) {
  const na = normalizeName(a), nb = normalizeName(b);
  if (na === nb) return true;
  // Allow 1-edit difference for names of 5+ chars (catches single-letter typos and all-caps mismatches)
  return na.length >= 5 && nb.length >= 5 && editDistance(na, nb) <= 1;
}

function isCitationMatched(citKey, referenceKeys) {
  const yearBase = citKey.year.replace(/[a-z]$/, "");
  return referenceKeys.some(
    (refKey) =>
      namesMatch(citKey.lastName, refKey.lastName) &&
      refKey.year.replace(/[a-z]$/, "") === yearBase,
  );
}

function hasDOIOrURL(text) {
  return /https?:\/\/|doi\.org|\bdoi:\s*10\.|\b10\.\d{4,}\//i.test(text);
}

function groupHasVisibleDOIOrURL(group) {
  return group.some((p) => hasDOIOrURL(p.text));
}

function groupHasHiddenHyperlinkURL(group) {
  return group.some(
    (p) =>
      p.hyperlinkUrls &&
      p.hyperlinkUrls.some((url) => /https?:\/\//i.test(url)) &&
      !hasDOIOrURL(p.text),
  );
}

function isDomainOnlyUrl(url) {
  try {
    const parsed = new URL(url.trim());
    return parsed.pathname === "/" || parsed.pathname === "";
  } catch {
    return false;
  }
}

function extractVisibleUrls(text) {
  return (text.match(/https?:\/\/[^\s,)<>]+/gi) || []);
}

function groupHasDomainOnlyUrl(group) {
  const combined = group.map((p) => p.text).join(" ");
  if (extractVisibleUrls(combined).some(isDomainOnlyUrl)) return true;
  return group.some((p) => p.hyperlinkUrls && p.hyperlinkUrls.some(isDomainOnlyUrl));
}

function groupReferenceEntries(referenceParagraphs) {
  const groups = [];
  for (const paragraph of referenceParagraphs) {
    if (looksLikeReferenceEntryStart(paragraph)) {
      groups.push([paragraph]);
    } else if (groups.length > 0) {
      groups[groups.length - 1].push(paragraph);
    }
  }
  return groups;
}

function mergeReferenceGroup(group) {
  if (group.length === 1) return group[0];
  let text = group[0].text;
  const runs = [...(group[0].runs || [])];
  for (let i = 1; i < group.length; i++) {
    const cont = group[i].text;
    // Preserve hyphenated word breaks (e.g. "E-" + "Commerce" → "E-Commerce")
    // otherwise join with a space
    if (/\-$/.test(text.trimEnd()) && /^[A-Za-z]/.test(cont.trimStart())) {
      text = text.trimEnd() + cont.trimStart();
    } else {
      text = text.trimEnd() + " " + cont.trimStart();
    }
    for (const run of (group[i].runs || [])) runs.push(run);
  }
  return { ...group[0], text, runs, isMerged: true, mergedFrom: group };
}

function getMergedReferenceEntries(referenceParagraphs) {
  return groupReferenceEntries(referenceParagraphs).map(mergeReferenceGroup);
}

function looksLikeIncompleteEnd(text) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  return /[,;]$/.test(trimmed) || !/[.!?]$/.test(trimmed);
}

function looksLikeContinuationStart(text) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  return /^[a-z]/.test(trimmed) || /^(https?:\/\/|www\.|doi:)/.test(trimmed);
}

function hasBrokenReferenceEntries(referenceParagraphs) {
  for (let i = 0; i < referenceParagraphs.length - 1; i++) {
    const current = referenceParagraphs[i];
    const next = referenceParagraphs[i + 1];
    if (
      looksLikeReferenceEntryStart(current) &&
      looksLikeIncompleteEnd(current.text) &&
      !looksLikeReferenceEntryStart(next) &&
      looksLikeContinuationStart(next.text)
    ) {
      return true;
    }
  }
  return false;
}

function hasExpectedHangingIndent(paragraph) {
  const hangingIndent = paragraph.formatting.hangingIndentInches;
  return hangingIndent.known && isClose(hangingIndent.value, EXPECTED_FIRST_LINE_INDENT_INCHES);
}

function isCentered(paragraph) {
  const alignment = paragraph.formatting.alignment;
  return alignment.known && alignment.value === "center";
}

function looksLikeBareUrl(paragraph) {
  return /^https?:\/\//i.test(paragraph.text.trim());
}

function findReferencesFormattingIssues(referencesHeading, referenceParagraphs) {
  const issues = [];
  const bareUrlParagraphs = referenceParagraphs.filter(looksLikeBareUrl);
  const properEntries = referenceParagraphs.filter((p) => !looksLikeBareUrl(p));
  const shortParagraphs = properEntries.filter((paragraph) => paragraph.text.trim().length < 40);
  const missingHangingIndentParagraphs = properEntries.filter((paragraph) => !hasExpectedHangingIndent(paragraph));

  if (referenceParagraphs.length === 0) {
    issues.push("No reference entries were detected after the References heading.");
  }

  if (bareUrlParagraphs.length > 0) {
    issues.push(BARE_URL_ISSUE);
  }

  if (shortParagraphs.length >= 2) {
    issues.push("Some reference entries appear incomplete or separated incorrectly.");
  }

  if (missingHangingIndentParagraphs.length > 0) {
    issues.push(MISSING_HANGING_INDENT_ISSUE);
  }

  if (hasBrokenReferenceEntries(properEntries)) {
    issues.push("Some reference entries appear to be broken across separate paragraphs.");
  }

  return issues;
}

function checkReferencesPage(extracted) {
  const referencesHeading = findReferencesHeadingNearEnd(extracted.paragraphs);
  const status = referencesHeading ? "pass" : "fail";
  const isSynthetic = referencesHeading && referencesHeading.synthetic;

  return {
    rule: "References page",
    status,
    passed: status === "pass",
    expected: "APA expects a References page with sources listed in alphabetical order using hanging indents.",
    expectedText: "APA expects a References page with sources listed in alphabetical order using hanging indents.",
    foundText:
      status === "fail"
        ? "APA Coach did not detect a References page at the end of the document."
        : isSynthetic
          ? "Reference entries were detected, but APA Coach did not find a 'References' heading above them."
          : "A References page was detected at the end of the document.",
    applicable: 1,
    checked: 1,
    matched: status === "pass" ? 1 : 0,
    failed: status === "fail" ? 1 : 0,
    unknown: 0,
    found: status === "pass" ? "References page detected" : "References page not detected",
    applicableParagraphs: referencesHeading ? 1 : 0,
    details:
      status === "pass"
        ? [`Detected References heading at paragraph ${referencesHeading.index}.`]
        : ["No References heading was detected near the end of the document."],
    howToFix: status === "fail" ? getHowToFix("References page") : [],
    resources: status === "fail" ? [APA_REFERENCES_PAGE_SCRIBBR_RESOURCE] : [],
  };
}

function checkReferencesHeadingAlignment(referencesHeading) {
  const centered = isCentered(referencesHeading);
  const status = centered ? "pass" : "fail";

  return {
    rule: "References heading alignment",
    status,
    passed: status === "pass",
    expected: "APA expects the 'References' heading to be centered.",
    expectedText: "APA expects the 'References' heading to be centered.",
    foundText: centered
      ? "The 'References' heading is centered."
      : "The 'References' heading is not centered.",
    applicable: 1,
    checked: 1,
    matched: centered ? 1 : 0,
    failed: centered ? 0 : 1,
    unknown: 0,
    found: centered ? "Heading is centered" : "Heading is not centered",
    applicableParagraphs: 1,
    details: centered
      ? ["The References heading uses center alignment."]
      : ["The References heading does not use center alignment."],
    howToFix: status === "fail" ? getHowToFix("References heading alignment") : [],
    resources: status === "fail" ? [ALIGN_TEXT_MSFT_RESOURCE, ALIGN_TEXT_GFG_RESOURCE] : [],
  };
}

function checkReferencesFormatting(extracted, referencesHeading) {
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const issues = findReferencesFormattingIssues(referencesHeading, referenceParagraphs);
  const status = issues.length > 0 ? "fail" : "pass";
  const resources =
    status === "fail" && issues.includes(MISSING_HANGING_INDENT_ISSUE)
      ? [HANGING_INDENT_RESOURCE, HANGING_INDENT_SCRIBBR_RESOURCE]
      : [];

  return {
    rule: "Reference hanging indent",
    status,
    passed: status === "pass",
    expected: "APA requires each reference entry to use a 0.5-inch hanging indent and double spacing.",
    expectedText: "APA requires each reference entry to use a 0.5-inch hanging indent and double spacing.",
    foundText:
      status === "fail"
        ? issues.join(" ")
        : "Reference entries appear to use basic APA formatting.",
    applicable: referenceParagraphs.length,
    checked: referenceParagraphs.length,
    matched: status === "pass" ? referenceParagraphs.length : Math.max(referenceParagraphs.length - issues.length, 0),
    failed: issues.length,
    unknown: 0,
    found:
      status === "pass"
        ? "Reference hanging indents detected"
        : `Formatting issues detected: ${issues.length}`,
    applicableParagraphs: referenceParagraphs.length,
    details: issues,
    howToFix: status === "fail" ? getHowToFix("Reference hanging indent") : [],
    resources,
  };
}

function checkReferencesLineSpacing(extracted, referencesHeading) {
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const unknowns = splitUnknowns(referenceParagraphs, "lineSpacingMultiple");
  const failures = referenceParagraphs.filter(
    (paragraph) =>
      paragraph.formatting.lineSpacingMultiple.known &&
      !isClose(paragraph.formatting.lineSpacingMultiple.value, EXPECTED_LINE_SPACING),
  );

  let foundText;
  if (referenceParagraphs.length === 0) {
    foundText = "APA Coach did not find any reference entries to check.";
  } else if (failures.length > 0) {
    foundText =
      "Some references are not double spaced. A hanging indent does not replace double spacing—both are required.";
  } else if (unknowns.length > 0) {
    foundText = "Reference line spacing could not be fully verified. This may need a quick check.";
  } else {
    foundText = "All reference entries use double spacing.";
  }

  const details = [];

  if (unknowns.length > 0) {
    details.push(
      `Line spacing: Unknown for ${unknowns.length} reference ${
        unknowns.length === 1 ? "paragraph" : "paragraphs"
      }. APA Coach could not determine whether spacing is inherited from a style.`,
    );
  }

  details.push(
    ...groupByDescription(failures, (paragraph) => {
      return `Line spacing is ${formatResolved(paragraph.formatting.lineSpacingMultiple, "x")}`;
    }),
  );

  return finishCheck(
    "References line spacing",
    "APA requires references to be double spaced with a hanging indent.",
    foundText,
    referenceParagraphs,
    failures,
    unknowns,
    details,
    getHowToFix("References line spacing"),
    [LINE_SPACING_MSFT_RESOURCE],
  );
}

function checkInlineCitations(extracted, referencesHeading) {
  const rule = "In-text citations";
  const expected = "Each source used in the paper should have an in-text citation.";

  const allText = extracted.paragraphs
    .filter((p) => p.role !== "blank" && (!referencesHeading || p.index < referencesHeading.index))
    .map((p) => p.text)
    .join(" ");
  const citationKeys = extractInlineCitationKeys(allText);

  if (citationKeys.length > 0) {
    return {
      rule, status: "pass", passed: true,
      expected, expectedText: expected,
      foundText: `APA Coach found ${citationKeys.length} in-text citation(s) in the document.`,
      applicable: citationKeys.length, checked: citationKeys.length, matched: citationKeys.length,
      failed: 0, unknown: 0, found: `${citationKeys.length} in-text citation(s) found`,
      applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const substantiveCount = extracted.paragraphs.filter(
    (p) => p.role !== "blank" && p.role !== "titlePage" && (!referencesHeading || p.index < referencesHeading.index),
  ).length;

  if (substantiveCount <= 2) {
    return {
      rule, status: "review", passed: false,
      expected, expectedText: expected,
      foundText: "APA Coach did not find enough content to check for in-text citations.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "Not enough content", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "fail", passed: false,
    expected, expectedText: expected,
    foundText: "APA Coach did not find any in-text citations. If you used sources, add (Author, Year) citations throughout the body.",
    applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
    found: "No in-text citations found", applicableParagraphs: 0,
    details: [], howToFix: getHowToFix(rule), resources: [INLINE_CITATIONS_APA_RESOURCE, INLINE_CITATIONS_OWL_RESOURCE],
  };
}

function checkUncitedReferences(extracted, referencesHeading) {
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);
  const bodyParagraphs = getParagraphsByRole(extracted, "body").filter(
    (p) => p.index < referencesHeading.index,
  );
  const bodyText = bodyParagraphs.map((p) => p.text).join(" ");
  const referenceKeys = entryParagraphs.map((p) => extractReferenceKey(p.text)).filter(Boolean);

  if (referenceKeys.length === 0) {
    return {
      rule: "Uncited references",
      status: "review",
      passed: false,
      expected: "Each reference should have at least one matching in-text citation.",
      expectedText: "Each reference should have at least one matching in-text citation.",
      foundText: "APA Coach could not parse the reference entries to check for in-text citations.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "Could not parse references",
      applicableParagraphs: 0,
      details: [], howToFix: [], resources: [],
    };
  }

  const citationKeys = extractInlineCitationKeys(bodyText);
  const bodyHasCitations = citationKeys.length > 0;
  const uncited = referenceKeys.filter((key) => !isReferenceCited(key, bodyText, citationKeys));
  const status = uncited.length === 0 ? "pass" : bodyHasCitations ? "fail" : "review";

  return {
    rule: "Uncited references",
    status,
    passed: status === "pass",
    expected: "Each reference should have at least one matching in-text citation.",
    expectedText: "Each reference should have at least one matching in-text citation.",
    foundText:
      status === "pass"
        ? `All ${referenceKeys.length} references appear to have a matching in-text citation.`
        : !bodyHasCitations
          ? "APA Coach could not find any in-text citations in the body. If your paper uses citations, check that they follow APA format: (Author, Year)."
          : `${uncited.length} of ${referenceKeys.length} reference${referenceKeys.length === 1 ? "" : "s"} appear to have no matching in-text citation.`,
    applicable: referenceKeys.length,
    checked: referenceKeys.length,
    matched: referenceKeys.length - uncited.length,
    failed: uncited.length,
    unknown: 0,
    found: status === "pass" ? "All references cited" : `${uncited.length} reference(s) appear uncited`,
    applicableParagraphs: entryParagraphs.length,
    details: uncited.map((key) => `No citation found for: "${key.preview}"`),
    howToFix: status === "fail" ? getHowToFix("Uncited references") : [],
    resources: [],
    missingItems: status === "fail" ? uncited.map((key) => key.preview) : [],
    missingItemsLabel: "References that appear uncited:",
  };
}

function checkUnmatchedCitations(extracted, referencesHeading) {
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);
  const bodyParagraphs = getParagraphsByRole(extracted, "body").filter(
    (p) => p.index < referencesHeading.index,
  );
  const bodyText = bodyParagraphs.map((p) => p.text).join(" ");
  const citationKeys = extractInlineCitationKeys(bodyText);
  const referenceKeys = entryParagraphs.map((p) => extractReferenceKey(p.text)).filter(Boolean);

  if (citationKeys.length === 0) {
    return {
      rule: "Unmatched citations",
      status: "review",
      passed: false,
      expected: "Each in-text citation should have a matching entry in the References list.",
      expectedText: "Each in-text citation should have a matching entry in the References list.",
      foundText: "APA Coach could not find any in-text citations in the body to check.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No in-text citations found",
      applicableParagraphs: 0,
      details: [], howToFix: [], resources: [],
    };
  }

  if (referenceKeys.length === 0) {
    return {
      rule: "Unmatched citations",
      status: "review",
      passed: false,
      expected: "Each in-text citation should have a matching entry in the References list.",
      expectedText: "Each in-text citation should have a matching entry in the References list.",
      foundText: "APA Coach could not parse the References list to check against in-text citations.",
      applicable: citationKeys.length, checked: 0, matched: 0, failed: 0, unknown: citationKeys.length,
      found: "Could not parse references",
      applicableParagraphs: 0,
      details: [], howToFix: [], resources: [],
    };
  }

  const unmatched = citationKeys.filter((cit) => !isCitationMatched(cit, referenceKeys));
  const status = unmatched.length === 0 ? "pass" : "fail";

  return {
    rule: "Unmatched citations",
    status,
    passed: status === "pass",
    expected: "Each inline citation should have a matching entry in the References list.",
    expectedText: "Each inline citation should have a matching entry in the References list.",
    foundText:
      status === "pass"
        ? `All ${citationKeys.length} in-text citations appear to have a matching reference.`
        : `${unmatched.length} of ${citationKeys.length} in-text citation${citationKeys.length === 1 ? "" : "s"} appear to have no matching reference entry.`,
    applicable: citationKeys.length,
    checked: citationKeys.length,
    matched: citationKeys.length - unmatched.length,
    failed: unmatched.length,
    unknown: 0,
    found: status === "pass" ? "All citations matched" : `${unmatched.length} citation(s) unmatched`,
    applicableParagraphs: entryParagraphs.length,
    details: unmatched.map((cit) => `No reference found for: "${cit.source}"`),
    howToFix: status === "fail" ? getHowToFix("Unmatched citations") : [],
    resources: [],
    missingItems: status === "fail" ? unmatched.map((cit) => cit.source) : [],
    missingItemsLabel: "Citations with no matching reference:",
  };
}

function checkReferenceDOIs(extracted, referencesHeading) {
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule: "Reference DOI/URL",
      status: "review",
      passed: false,
      expected: "APA 7th edition requires a DOI or URL for most sources.",
      expectedText: "APA 7th edition requires a DOI or URL for most sources.",
      foundText: "APA Coach could not find any reference entries to check for DOIs.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references found",
      applicableParagraphs: 0,
      details: [], howToFix: [], resources: [],
    };
  }

  const groups = groupReferenceEntries(referenceParagraphs);
  const total = groups.length;

  const missingGroups = groups.filter((g) => !groupHasVisibleDOIOrURL(g));
  const hiddenLinkGroups = missingGroups.filter((g) => groupHasHiddenHyperlinkURL(g));
  const trulyMissingGroups = missingGroups.filter((g) => !groupHasHiddenHyperlinkURL(g));
  const status = missingGroups.length === 0 ? "pass" : "fail";

  const missingItems = [
    ...hiddenLinkGroups.map((g) => {
      const preview = g[0].text.length > 70 ? g[0].text.substring(0, 70) + "…" : g[0].text;
      return `${preview} — URL is hidden in a hyperlink; paste it as visible text`;
    }),
    ...trulyMissingGroups.map((g) => (g[0].text.length > 80 ? g[0].text.substring(0, 80) + "…" : g[0].text)),
  ];

  return {
    rule: "Reference DOI/URL",
    status,
    passed: status === "pass",
    expected: "APA 7th edition requires a DOI or URL for most sources.",
    expectedText: "APA 7th edition requires a DOI or URL for most sources.",
    foundText:
      status === "pass"
        ? `All ${total} references include a DOI or URL.`
        : `${missingGroups.length} of ${total} reference${total === 1 ? "" : "s"} appear to be missing a visible DOI or URL.`,
    applicable: total,
    checked: total,
    matched: total - missingGroups.length,
    failed: missingGroups.length,
    unknown: 0,
    found: status === "pass" ? "All references have DOI/URL" : `${missingGroups.length} reference(s) missing DOI/URL`,
    applicableParagraphs: entryParagraphs.length,
    details: missingItems.map((item) => `Issue: "${item}"`),
    howToFix: status === "fail" ? getHowToFix("Reference DOI/URL") : [],
    resources: status === "fail" ? [APA_DOIS_URLS_RESOURCE] : [],
    missingItems: status === "fail" ? missingItems : [],
    missingItemsLabel: "References missing a visible DOI or URL:",
  };
}

function checkUnconvertedMarkup(extracted, referencesHeading) {
  const rule = "Unconverted markup symbols";
  const expected = "All formatting should use Word's built-in styles — not markdown symbols like *asterisks* or **double asterisks**.";

  // For reference paragraphs, use merged entries so split references show as one unit
  const refParaSet = new Set();
  const mergedRefParagraphs = [];
  if (referencesHeading) {
    const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
    const merged = getMergedReferenceEntries(referenceParagraphs);
    for (const m of merged) {
      mergedRefParagraphs.push(m);
      if (m.mergedFrom) { for (const src of m.mergedFrom) refParaSet.add(src); }
      else refParaSet.add(m);
    }
  }
  const nonRefParagraphs = extracted.paragraphs.filter((p) => p.text.trim().length > 0 && !refParaSet.has(p));
  const applicableParagraphs = [...nonRefParagraphs, ...mergedRefParagraphs];

  const failures = [];
  const details = [];
  const missingItems = [];

  for (const p of applicableParagraphs) {
    const text = p.text;
    // Match paired asterisks or double-asterisks around at least one letter
    // Require a letter immediately after the opening marker to avoid lone * used as a bullet
    if (/\*\*[A-Za-z][^*\n]*\*\*/.test(text) || /\*[A-Za-z][^*\n]*\*/.test(text)) {
      failures.push(p);
      const preview = text.trim().length > 70 ? text.trim().slice(0, 70) + "…" : text.trim();
      details.push(`Markdown asterisks found: "${preview}"`);
      missingItems.push(`"${preview}"`);
    }
  }

  const foundText =
    failures.length > 0
      ? `${failures.length} paragraph${failures.length === 1 ? "" : "s"} contain${failures.length === 1 ? "s" : ""} markdown asterisks that were not converted to Word formatting.`
      : "No unconverted markdown symbols found.";

  return {
    ...finishCheck(rule, expected, foundText, applicableParagraphs, failures, [], details,
      getHowToFix(rule), []),
    missingItems: missingItems.length > 0 ? missingItems : [],
    missingItemsLabel: missingItems.length > 0 ? "Paragraphs with markup symbols:" : "",
  };
}

function checkFonts(extracted) {
  const rule = "Font";
  const expected = "APA recommends 12-point Times New Roman throughout the document.";

  const explicitSizes = new Set();
  const explicitFamilies = new Set();

  for (const p of extracted.paragraphs) {
    if (p.role === "blank" || !p.fonts) continue;
    for (const s of p.fonts.sizes) explicitSizes.add(s);
    for (const f of p.fonts.families) explicitFamilies.add(f);
  }

  // Fall back to document defaults only when paragraphs have no explicit font info
  const allSizes = explicitSizes.size > 0 ? explicitSizes : new Set(
    extracted.fontDefaults?.sizePoints != null ? [extracted.fontDefaults.sizePoints] : [],
  );
  const allFamilies = explicitFamilies.size > 0 ? explicitFamilies : new Set(
    extracted.fontDefaults?.family ? [extracted.fontDefaults.family] : [],
  );

  if (allSizes.size === 0 && allFamilies.size === 0) {
    return {
      rule, status: "review", passed: false,
      expected, expectedText: expected,
      foundText: "APA Coach could not detect font information in this document.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No font info", applicableParagraphs: 0,
      details: [], howToFix: [], resources: [],
    };
  }

  const sortedSizes = [...allSizes].sort((a, b) => a - b);
  const oversized = sortedSizes.filter((s) => s > 12);
  const issues = [];

  if (oversized.length > 0) {
    issues.push("Font size exceeds 12pt.");
  }
  if (allSizes.size > 1) {
    issues.push(`Multiple font sizes detected: ${sortedSizes.map((s) => s + "pt").join(", ")}.`);
  }
  if (allFamilies.size > 1) {
    issues.push(`Multiple fonts detected: ${[...allFamilies].join(", ")}.`);
  }

  const status = issues.length > 0 ? "fail" : "pass";
  const sizeLabel = sortedSizes.length === 1 ? sortedSizes[0] + "pt" : sortedSizes.map((s) => s + "pt").join("/");
  const familyLabel = allFamilies.size === 1 ? [...allFamilies][0] : `${allFamilies.size} fonts`;

  return {
    rule, status, passed: status === "pass",
    expected, expectedText: expected,
    foundText:
      status === "pass"
        ? `Font appears consistent: ${sizeLabel} ${familyLabel}.`
        : issues.join(" "),
    applicable: allSizes.size + allFamilies.size,
    checked: allSizes.size + allFamilies.size,
    matched: status === "pass" ? allSizes.size + allFamilies.size : 0,
    failed: issues.length, unknown: 0,
    found: status === "pass" ? "Font consistent" : `${issues.length} font issue(s) detected`,
    applicableParagraphs: extracted.paragraphs.filter((p) => p.role !== "blank").length,
    details: issues,
    howToFix: status === "fail" ? getHowToFix("Font") : [],
    resources: [],
  };
}

function checkReferenceShortLinks(extracted, referencesHeading) {
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const groups = groupReferenceEntries(referenceParagraphs);

  if (groups.length === 0) {
    return {
      rule: "Reference short link",
      status: "review", passed: false,
      expected: "Each reference URL should link directly to the specific article or page, not just the website homepage.",
      expectedText: "Each reference URL should link directly to the specific article or page, not just the website homepage.",
      foundText: "APA Coach could not find any reference entries to check for short links.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references found", applicableParagraphs: 0,
      details: [], howToFix: [], resources: [], missingItems: [], missingItemsLabel: "",
    };
  }

  const shortLinkGroups = groups.filter(groupHasDomainOnlyUrl);
  const status = shortLinkGroups.length === 0 ? "pass" : "fail";
  const total = groups.length;

  const missingItems = shortLinkGroups.map((g) =>
    g[0].text.length > 80 ? g[0].text.substring(0, 80) + "…" : g[0].text,
  );

  return {
    rule: "Reference short link",
    status, passed: status === "pass",
    expected: "Each reference URL should link directly to the specific article or page, not just the website homepage.",
    expectedText: "Each reference URL should link directly to the specific article or page, not just the website homepage.",
    foundText:
      status === "pass"
        ? `All ${total} references with URLs appear to link to a specific page.`
        : `${shortLinkGroups.length} of ${total} reference${total === 1 ? "" : "s"} appear to use a domain-only URL (e.g., https://www.example.com) instead of a direct link to the article.`,
    applicable: total, checked: total, matched: total - shortLinkGroups.length,
    failed: shortLinkGroups.length, unknown: 0,
    found: status === "pass" ? "No short links detected" : `${shortLinkGroups.length} reference(s) use a domain-only URL`,
    applicableParagraphs: referenceParagraphs.length,
    details: missingItems.map((item) => `Short link: "${item}"`),
    howToFix: status === "fail" ? getHowToFix("Reference short link") : [],
    resources: [],
    missingItems: status === "fail" ? missingItems : [],
    missingItemsLabel: "References with domain-only URLs:",
  };
}

function urlMatchesUnapproved(url) {
  try {
    const hostname = new URL(url.trim()).hostname.toLowerCase();
    return UNAPPROVED_DOMAINS.some((d) => hostname === d || hostname.endsWith("." + d));
  } catch {
    return false;
  }
}

function groupHasUnapprovedSource(group) {
  const combined = group.map((p) => p.text).join(" ");
  if (extractVisibleUrls(combined).some(urlMatchesUnapproved)) return true;
  return group.some((p) => p.hyperlinkUrls && p.hyperlinkUrls.some(urlMatchesUnapproved));
}

function checkUnapprovedSources(extracted, referencesHeading) {
  const rule = "Unapproved source";
  const expected = "References must use sources approved for academic use at AIU.";
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const groups = groupReferenceEntries(referenceParagraphs);

  if (groups.length === 0) {
    return {
      rule, status: "review", passed: false,
      expected, expectedText: expected,
      foundText: "APA Coach could not find any reference entries to check.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references found", applicableParagraphs: 0,
      details: [], howToFix: [], resources: [], missingItems: [], missingItemsLabel: "",
    };
  }

  const badGroups = groups.filter(groupHasUnapprovedSource);
  const status = badGroups.length === 0 ? "pass" : "fail";
  const missingItems = badGroups.map((g) =>
    g[0].text.length > 80 ? g[0].text.substring(0, 80) + "…" : g[0].text,
  );

  return {
    rule, status, passed: status === "pass",
    expected, expectedText: expected,
    foundText: status === "pass"
      ? `All ${groups.length} references appear to use approved sources.`
      : `${badGroups.length} of ${groups.length} reference${groups.length === 1 ? "" : "s"} use a source that is not approved for academic use at AIU.`,
    applicable: groups.length, checked: groups.length,
    matched: groups.length - badGroups.length, failed: badGroups.length, unknown: 0,
    found: status === "pass" ? "No unapproved sources detected" : `${badGroups.length} unapproved source(s) found`,
    applicableParagraphs: referenceParagraphs.length,
    details: missingItems.map((item) => `Unapproved source: "${item}"`),
    howToFix: status === "fail" ? getHowToFix(rule) : [],
    resources: [],
    missingItems: status === "fail" ? missingItems : [],
    missingItemsLabel: "References using unapproved sources:",
  };
}

// ─── Phase 2: Reference facet checks ────────────────────────────────────────

function checkReferenceAuthors(extracted, referencesHeading) {
  const rule = "Reference authors";
  const expected = 'APA requires authors in "Last, F. M." format with "&" (not "and") before the final author.';
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule, status: "review", passed: false, expected, expectedText: expected,
      foundText: "APA Coach could not find reference entries to check author formatting.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references found", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const failures = [];
  const unknowns = [];
  const details = [];
  const missingItems = [];

  for (const p of entryParagraphs) {
    const parsed = parseReferenceEntry(p);
    if (!parsed.authorsRaw) { unknowns.push(p); continue; }
    const authors = parsed.authorsRaw;
    const preview = authors.length > 70 ? authors.slice(0, 70) + "…" : authors;
    const refLabel = parsed.authorsRaw.split(",")[0] + (parsed.year ? ` (${parsed.year})` : "");
    let failed = false;
    let failureDetail = null;

    if (/,\s+and\s+[A-Z]/i.test(authors)) {
      failures.push(p);
      failureDetail = `Use "&" instead of "and" between authors: "${preview}"`;
      details.push(failureDetail);
      failed = true;
    }
    if (/\bet\.?\s*al\.?\b/i.test(authors)) {
      if (!failed) failures.push(p);
      failureDetail = `"et al." should not appear in a reference entry — list all authors or use the 21+ format.`;
      details.push(failureDetail);
      failed = true;
    }

    // Check each individual author token.
    // APA separates authors with ", " (and ", & " before the last), so the boundary between
    // two authors always looks like: <initial>., <LastName>... or <initial>. & <LastName>...
    // Lookbehind keeps the period with the token so "S." doesn't become "S" (false positive).
    const authorTokens = authors.split(/(?<=\.)(?:,\s*(?:&\s*)?|\s*&\s*)(?=[A-Z])/);
    for (const token of authorTokens) {
      const trimmed = token.trim();
      if (!trimmed) continue;

      // Expect "Last, I." or "Last, I. M." — must have a comma
      const commaIdx = trimmed.indexOf(",");
      if (commaIdx === -1) continue; // can't check without comma

      const lastName = trimmed.slice(0, commaIdx).trim();
      const initialsRaw = trimmed.slice(commaIdx + 1).trim();

      // "First Last" format: last name contains a space, suggesting the whole author
      // string is in First Last order rather than Last, F. order
      if (/\s/.test(lastName)) {
        if (!failed) failures.push(p);
        failureDetail = `Authors appear to be listed first-name first — use "Last, F." format instead.`;
        details.push(failureDetail);
        failed = true;
        break;
      }

      // All-caps last name: e.g. HALLETT
      if (lastName.length > 1 && lastName === lastName.toUpperCase() && /^[A-Z]/.test(lastName)) {
        if (!failed) failures.push(p);
        const titleCased = lastName[0] + lastName.slice(1).toLowerCase();
        failureDetail = `Last name is all caps — write "${titleCased}" not "${lastName}"`;
        details.push(failureDetail);
        failed = true;
        break;
      }

      // Last name doesn't start with uppercase — allow lowercase particles (da, de, di, van, von, le, la, etc.)
      if (lastName.length > 0 && !/^[A-Z]/.test(lastName)) {
        const hasParticle = /^(?:da|de|del|della|di|do|dos|du|van|van der|von|le|la|l'|al|bin|binti|bt)\s+[A-Z]/i.test(lastName);
        if (!hasParticle) {
          if (!failed) failures.push(p);
          failureDetail = `Last name should start with a capital letter: "${lastName}" in "${preview}"`;
          details.push(failureDetail);
          failed = true;
          break;
        }
      }

      // Full first name instead of initial: a word of 2+ letters with no trailing period
      // e.g. "Scharon" or "Jane" — match a word that is 2+ alpha chars
      if (/\b[A-Za-z]{2,}(?!\.)/.test(initialsRaw)) {
        // Allow "Jr." or "Sr." suffixes — ignore those
        const withoutSuffix = initialsRaw.replace(/\b(Jr|Sr|II|III|IV)\b\.?/gi, "").trim();
        if (/\b[A-Za-z]{2,}(?!\.)/.test(withoutSuffix)) {
          const fullNameMatch = withoutSuffix.match(/\b([A-Za-z]{2,})(?!\.)/);
          const fullName = fullNameMatch ? fullNameMatch[1] : initialsRaw.trim();
          const suggested = `${lastName}, ${fullName[0].toUpperCase()}.`;
          if (!failed) failures.push(p);
          failureDetail = `Full name used instead of initial — write "${suggested}" not "${lastName}, ${fullName}"`;
          details.push(failureDetail);
          failed = true;
          break;
        }
      }

      // Missing period after an initial: a single uppercase letter not followed by a period
      // e.g. "J R" or "J," or trailing "J"
      if (/\b[A-Z](?!\.)/.test(initialsRaw)) {
        if (!failed) failures.push(p);
        failureDetail = `Initial is missing a period — use "J." not "J": "${preview}"`;
        details.push(failureDetail);
        failed = true;
        break;
      }

      // Missing space between initials: "J.R." should be "J. R."
      if (/\.[A-Z]/.test(initialsRaw)) {
        if (!failed) failures.push(p);
        failureDetail = `Missing space between initials — use "J. R." not "J.R.": "${preview}"`;
        details.push(failureDetail);
        failed = true;
        break;
      }
    }

    if (failed && failureDetail) {
      missingItems.push(`${refLabel}: ${failureDetail}`);
    }
  }

  const foundText =
    failures.length > 0
      ? `${failures.length} reference${failures.length === 1 ? "" : "s"} appear to have author formatting issues.`
      : unknowns.length > 0
        ? "APA Coach could not verify author format for some references — please review manually."
        : "Reference author formatting appears correct.";

  return {
    ...finishCheck(rule, expected, foundText, entryParagraphs, failures, unknowns, details,
      getHowToFix(rule), [APA_AUTHORS_RESOURCE]),
    missingItems: missingItems.length > 0 ? missingItems : [],
    missingItemsLabel: missingItems.length > 0 ? "Author formatting issues:" : "",
  };
}

function checkReferenceYear(extracted, referencesHeading) {
  const rule = "Reference year format";
  const expected = "APA requires the year in parentheses followed by a period: (2020). Title…";
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule, status: "review", passed: false, expected, expectedText: expected,
      foundText: "No reference entries found.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const failures = [];
  const unknowns = [];
  const details = [];
  const missingItems = [];

  for (const p of entryParagraphs) {
    const yearM = p.text.match(/\((\d{4}[a-z]?|n\.d\.)(?:,\s*[^)]+)?\)/);
    if (!yearM) { unknowns.push(p); continue; }
    const afterParen = p.text.slice(yearM.index + yearM[0].length);
    if (afterParen.trim().length > 0 && !/^\s*\./.test(afterParen)) {
      failures.push(p);
      const preview = p.text.length > 70 ? p.text.slice(0, 70) + "…" : p.text;
      details.push(`Missing period after year "(${yearM[1]})": "${preview}"`);
      const parsed = parseReferenceEntry(p);
      const authorLabel = parsed.authorsRaw
        ? `${parsed.authorsRaw.split(",")[0]}${parsed.year ? ` (${parsed.year})` : ""}: `
        : "";
      missingItems.push(`${authorLabel}"${preview}"`);
    }
  }

  const foundText =
    failures.length > 0
      ? `${failures.length} reference${failures.length === 1 ? "" : "s"} appear to be missing a period after the year.`
      : unknowns.length > 0
        ? "APA Coach could not verify year format for some references."
        : "Reference year formatting appears correct.";

  return {
    ...finishCheck(rule, expected, foundText, entryParagraphs, failures, unknowns, details,
      getHowToFix(rule), [APA_REFERENCE_FORMAT_RESOURCE]),
    missingItems: missingItems.length > 0 ? missingItems : [],
    missingItemsLabel: missingItems.length > 0 ? "References with year format issues:" : "",
  };
}

function checkReferenceTitleCapitalization(extracted, referencesHeading) {
  const rule = "Reference title capitalization";
  const expected = "APA requires sentence case for reference titles — capitalize only the first word, the first word after a colon, and proper nouns.";
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule, status: "review", passed: false, expected, expectedText: expected,
      foundText: "No reference entries found.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
      missingItems: [], missingItemsLabel: "",
    };
  }

  const failures = [];
  const unknowns = [];
  const details = [];
  const missingItems = [];

  for (const p of entryParagraphs) {
    const parsed = parseReferenceEntry(p);
    if (!parsed.title) { unknowns.push(p); continue; }
    // Webpage titles vary too much to check reliably; all other kinds use sentence case
    if (parsed.kindGuess === "webpage") continue;

    const title = parsed.title;
    const words = title.split(/\s+/);

    let afterColon = false;
    let capsCount = 0;

    for (let i = 0; i < words.length; i++) {
      const word = words[i];
      // Strip leading/trailing punctuation for the capital check
      const bare = word.replace(/^["""''([\-–—]+|["""'')\].,!?;:]+$/g, "");
      const isSkipped = i === 0 || afterColon;
      afterColon = word.endsWith(":") || word.endsWith("—");
      if (isSkipped) continue;
      // Acronyms (all-caps, 2+ chars) are legitimate in any case style
      if (bare.length > 1 && bare === bare.toUpperCase()) continue;
      if (/^[A-Z]/.test(bare)) capsCount++;
    }

    if (capsCount >= 2) {
      failures.push(p);
      const preview = title.length > 60 ? title.slice(0, 60) + "…" : title;
      details.push(`Title appears to use Title Case instead of sentence case: "${preview}"`);
      const authorLabel = parsed.authorsRaw
        ? `${parsed.authorsRaw.split(",")[0]}${parsed.year ? ` (${parsed.year})` : ""}: `
        : "";
      missingItems.push(`${authorLabel}"${preview}"`);
    }
  }

  const foundText =
    failures.length > 0
      ? `${failures.length} reference title${failures.length === 1 ? "" : "s"} appear to use Title Case instead of sentence case.`
      : unknowns.length > 0
        ? "APA Coach could not verify title capitalization for some references — please review manually."
        : "Reference title capitalization appears correct.";

  return {
    ...finishCheck(rule, expected, foundText, entryParagraphs, failures, unknowns, details,
      getHowToFix(rule), [APA_REFERENCE_FORMAT_RESOURCE]),
    missingItems: missingItems.length > 0 ? missingItems : [],
    missingItemsLabel: missingItems.length > 0 ? "Titles to review:" : "",
  };
}

function checkReferenceItalics(extracted, referencesHeading) {
  const rule = "Reference italics";
  const expected = "APA uses italics in references, but the rules depend on the source type:";
  const expectedItems = [
    "Books and reports: the title should be in italics.",
    { text: "Book *chapters*, if used: do not italicize the chapter title.", sub: true },
    "Journal articles: do not italicize the article title. Italicize the journal name and the volume number — but not the issue number (the number in parentheses).",
  ];
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule, status: "review", passed: false, expected, expectedText: expected, expectedItems,
      foundText: "APA Coach could not find reference entries to check italic formatting.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references found", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const failures = [];
  const unknowns = [];
  const details = [];

  for (const p of entryParagraphs) {
    const parsed = parseReferenceEntry(p);
    const kind = parsed.kindGuess;
    let checkedAnything = false;

    // Title italics — applies to journal-article, book, book-chapter
    if (parsed.title && (kind === "journal-article" || kind === "book" || kind === "book-chapter")) {
      const titleState = getSpanItalicState(p.runs, parsed.title);
      if (titleState !== "not-found") {
        checkedAnything = true;
        const preview = parsed.title.length > 60 ? parsed.title.slice(0, 60) + "…" : parsed.title;
        if ((kind === "journal-article" || kind === "book-chapter") && (titleState === "italic" || titleState === "mixed")) {
          failures.push(p);
          details.push(`Article/chapter title should not be italicized: "${preview}"`);
        } else if (kind === "book" && (titleState === "not-italic" || titleState === "mixed")) {
          failures.push(p);
          details.push(`Book title should be italicized: "${preview}"`);
        }
      }
    }

    // Journal name + volume/issue italics — journal-article only
    if (kind === "journal-article") {
      const journalData = parseJournalSourcePart(parsed.sourcePart);
      if (journalData) {
        if (journalData.journalName) {
          const nameState = getSpanItalicState(p.runs, journalData.journalName);
          if (nameState !== "not-found") {
            checkedAnything = true;
            if (nameState === "not-italic" || nameState === "mixed") {
              failures.push(p);
              details.push(`Journal name should be italicized: "${journalData.journalName}"`);
            }
          }
        }
        if (journalData.volume) {
          // Include issue in search to avoid matching year digits (e.g. "21" in "(2021)")
          const volSearchText = journalData.volume + (journalData.issueWithParens || "");
          const volState = getSpanItalicState(p.runs, volSearchText);
          if (volState !== "not-found") {
            checkedAnything = true;
            if (volState === "not-italic") {
              failures.push(p);
              details.push(`Volume number "${journalData.volume}" should be italicized.`);
            }
          }
        }
        if (journalData.issueWithParens) {
          const issueState = getSpanItalicState(p.runs, journalData.issueWithParens);
          if (issueState !== "not-found") {
            checkedAnything = true;
            if (issueState === "italic") {
              failures.push(p);
              details.push(`Issue number "${journalData.issueWithParens}" should not be italicized — only the volume number is italic.`);
            }
          }
        }
      }
    }

    if (!checkedAnything && kind === "unknown") unknowns.push(p);
  }

  const failedParagraphs = [...new Set(failures)];
  const missingItems = failedParagraphs.map((p) => {
    const t = p.text.trim();
    return t.length > 80 ? t.slice(0, 80) + "…" : t;
  });

  const foundText =
    failures.length > 0
      ? `${failures.length} italic formatting issue${failures.length === 1 ? "" : "s"} found across reference entries.`
      : unknowns.length === entryParagraphs.length
        ? "APA Coach could not verify italic formatting — reference types could not be determined."
        : unknowns.length > 0
          ? `Italic formatting appears correct for ${entryParagraphs.length - unknowns.length} reference${entryParagraphs.length - unknowns.length === 1 ? "" : "s"}; ${unknowns.length} could not be verified.`
          : "Reference italic formatting appears correct.";

  return {
    ...finishCheck(rule, expected, foundText, entryParagraphs, failures, unknowns, details,
      getHowToFix(rule), [APA_REFERENCE_FORMAT_RESOURCE]),
    expectedItems,
    missingItems: missingItems.length > 0 ? missingItems : [],
    missingItemsLabel: missingItems.length > 0 ? "References with italic issues:" : "",
  };
}

function checkReferencePunctuation(extracted, referencesHeading) {
  const rule = "Reference punctuation";
  const expected = "For journal articles, APA has two punctuation rules for the volume, issue, and page range:";
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule, status: "review", passed: false, expected, expectedText: expected,
      foundText: "No reference entries found.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const failures = [];
  const unknowns = [];
  const details = [];

  for (const p of entryParagraphs) {
    const parsed = parseReferenceEntry(p);
    if (parsed.kindGuess === "unknown") { unknowns.push(p); continue; }

    if (parsed.kindGuess === "journal-article" && parsed.sourcePart) {
      // No p./pp. in journal page ranges
      if (/\bpp?\.\s*\d/i.test(parsed.sourcePart)) {
        failures.push(p);
        const src = parsed.sourcePart.length > 60 ? parsed.sourcePart.slice(0, 60) + "…" : parsed.sourcePart;
        details.push(`Journal page ranges should not use "p." or "pp.": "${src}"`);
        continue;
      }
      // No space between volume and issue: 15(3) not 15 (3)
      if (/\d\s+\(\s*\d/.test(parsed.sourcePart)) {
        failures.push(p);
        const src = parsed.sourcePart.length > 60 ? parsed.sourcePart.slice(0, 60) + "…" : parsed.sourcePart;
        details.push(`Write volume and issue with no space: "15(3)" not "15 (3)": "${src}"`);
      }
    }
  }

  const foundText =
    failures.length > 0
      ? `${failures.length} reference${failures.length === 1 ? "" : "s"} appear to have punctuation issues.`
      : unknowns.length > 0
        ? "APA Coach could not verify punctuation for some entries — please review manually."
        : "Reference punctuation appears correct.";

  const expectedItems = [
    "Write the volume and issue together with no space: *15*(3) — not *15* (3).",
    "Write the page range directly after the issue: *15*(3), 45–67 — do not add \"p.\" or \"pp.\" before the numbers.",
  ];

  return {
    ...finishCheck(rule, expected, foundText, entryParagraphs, failures, unknowns, details,
      getHowToFix(rule), [APA_REFERENCE_FORMAT_RESOURCE]),
    expectedItems,
  };
}

function checkReferenceDOIFormat(extracted, referencesHeading) {
  const rule = "Reference DOI format";
  const expected = 'APA requires DOIs as full URLs: https://doi.org/10.xxxx — not "doi:" or "http://dx.doi.org/".';
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule, status: "review", passed: false, expected, expectedText: expected,
      foundText: "No reference entries found.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const failures = [];
  const details = [];

  for (const p of entryParagraphs) {
    const text = p.text;
    if (/\bdoi:\s*10\./i.test(text)) {
      failures.push(p);
      details.push(`Use "https://doi.org/10.xxx" not "doi:10.xxx" — found in: "${text.slice(0, 70)}"`);
      continue;
    }
    if (/\bhttp:\/\/(?:dx\.)?doi\.org\//i.test(text)) {
      failures.push(p);
      details.push(`Use "https://doi.org/" not "http://dx.doi.org/" — found in: "${text.slice(0, 70)}"`);
      continue;
    }
    // Bare DOI: 10.xxxx/... not already preceded by https://doi.org/ or doi:
    const bareDoiMatch = text.match(/\b(10\.\d{4,}\/\S+)/);
    if (bareDoiMatch) {
      const before = text.slice(0, text.indexOf(bareDoiMatch[1]));
      if (!/https?:\/\/(?:dx\.)?doi\.org\/$/i.test(before) && !/doi:\s*$/i.test(before)) {
        const doi = bareDoiMatch[1].replace(/[.,;)]+$/, "");
        failures.push(p);
        details.push(`Bare DOI must be a full URL — replace "${doi}" with "https://doi.org/${doi}"`);
      }
    }
  }

  const foundText =
    failures.length > 0
      ? `${failures.length} reference${failures.length === 1 ? "" : "s"} use an outdated or incorrect DOI format.`
      : "Reference DOI formatting appears correct.";

  return finishCheck(rule, expected, foundText, entryParagraphs, failures, [], details,
    getHowToFix(rule), [APA_DOI_FORMAT_RESOURCE]);
}

function checkReferenceForbiddenPhrases(extracted, referencesHeading) {
  const rule = "Reference forbidden phrases";
  const expected = 'APA references do not use "Available at" or "accessed" — write the URL directly.';
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = getMergedReferenceEntries(referenceParagraphs);

  if (entryParagraphs.length === 0) {
    return {
      rule, status: "review", passed: false, expected, expectedText: expected,
      foundText: "No reference entries found.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No references", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const failures = [];
  const details = [];

  for (const p of entryParagraphs) {
    const text = p.text;
    if (/\bavailable\s+at\b/i.test(text)) {
      failures.push(p);
      details.push(`Remove "Available at" — write the URL directly: "${text.slice(0, 80)}…"`);
    } else if (/\baccessed\b/i.test(text)) {
      failures.push(p);
      details.push(`Remove "accessed" and the access date — APA does not use this: "${text.slice(0, 80)}…"`);
    }
  }

  const foundText = failures.length === 0
    ? `All ${entryParagraphs.length} references use correct URL phrasing.`
    : `${failures.length} reference${failures.length === 1 ? "" : "s"} use non-APA phrases ("Available at" or "accessed").`;

  return finishCheck(rule, expected, foundText, entryParagraphs, failures, [], details,
    getHowToFix(rule), []);
}

// ─── Phase 2.5: Citation facet checks ───────────────────────────────────────

// Checks body text (before references) for personal communication citations and
// returns their raw strings so checkUnmatchedCitations can exempt them.
function extractPersonalCommunicationStrings(extracted, referencesHeading) {
  const bodyText = getBodyText(extracted, referencesHeading);
  const results = [];
  const re = /\(([^)]*\bpersonal\s+communication\b[^)]*)\)/gi;
  let m;
  while ((m = re.exec(bodyText)) !== null) results.push(m[1]);
  return results;
}

function checkPersonalCommunications(extracted, referencesHeading) {
  const rule = "Personal communication";
  const expected = "Personal communications use the format: (F. M. Last, personal communication, Month Day, Year).";
  const bodyText = getBodyText(extracted, referencesHeading);

  const re = /\(([^)]*\bpersonal\s+communication\b[^)]*)\)/gi;
  const found = [];
  let m;
  while ((m = re.exec(bodyText)) !== null) found.push({ inner: m[1], raw: m[0] });

  if (found.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: "No personal communication citations found.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No personal communications", applicableParagraphs: 0,
      details: [], howToFix: [], resources: [],
    };
  }

  const failures = [];
  const details = [];

  for (const { inner, raw } of found) {
    const issues = [];
    // Must have initials + last name before "personal communication"
    const beforePhrase = inner.replace(/\bpersonal\s+communication\b.*/i, "").trim().replace(/,\s*$/, "");
    if (!/^[A-Z]\.\s/.test(beforePhrase)) {
      issues.push("include the person's full initials and last name (e.g., J. A. Smith)");
    }
    // Must have a specific date (month + day or month alone + year) after the phrase
    const afterPhrase = inner.replace(/.*\bpersonal\s+communication\b,?\s*/i, "");
    if (!/\b(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+\d/i.test(afterPhrase)) {
      issues.push("include the specific date (e.g., March 15, 2024)");
    }
    if (issues.length > 0) {
      failures.push(raw);
      details.push(`Personal communication "${raw}" — please ${issues.join(" and ")}.`);
    }
  }

  const total = found.length;
  const passed = total - failures.length;
  const status = failures.length > 0 ? "fail" : "pass";

  return {
    rule, status, passed: status === "pass", expected, expectedText: expected,
    foundText:
      status === "pass"
        ? `${total} personal communication citation${total === 1 ? "" : "s"} found and formatted correctly.`
        : `${failures.length} personal communication citation${failures.length === 1 ? "" : "s"} appear to have formatting issues.`,
    applicable: total, checked: total, matched: passed, failed: failures.length, unknown: 0,
    found: `${total} personal communication(s) found`,
    applicableParagraphs: 0, details,
    howToFix: status === "fail" ? getHowToFix(rule) : [], resources: [],
  };
}

function checkCitationAmpersandUsage(extracted, referencesHeading) {
  const rule = "Citation ampersand";
  const expected = 'Use "&" inside parenthetical citations; spell out "and" in narrative citations.';
  const bodyText = getBodyText(extracted, referencesHeading);
  const issues = [];

  // Wrong: "and" inside a parenthetical citation — (Smith and Jones, YYYY)
  const parenAndRe = /\(([^()]{2,200})\)/g;
  let m;
  while ((m = parenAndRe.exec(bodyText)) !== null) {
    const inner = m[1];
    if (!/\b\d{4}[a-z]?\b|n\.d\./.test(inner)) continue;
    if (/\bpersonal\s+communication\b/i.test(inner)) continue;
    if (/\band\s+[A-Z]/i.test(inner)) {
      const preview = m[0].length > 60 ? m[0].slice(0, 60) + "…)" : m[0];
      issues.push({ type: "and-in-parens", preview });
    }
  }

  // Wrong: "&" in narrative before (Year) — Author & Author (YYYY)
  const narrativeAmpRe = /([A-Z][a-zA-Z'-]+)\s+&\s+([A-Z][a-zA-Z'-]+)\s+\((?:\d{4}[a-z]?|n\.d\.)/g;
  while ((m = narrativeAmpRe.exec(bodyText)) !== null) {
    const preview = bodyText.slice(m.index, m.index + 60) + (m[0].length > 60 ? "…" : "");
    issues.push({ type: "amp-in-narrative", preview });
  }

  if (issues.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: "Ampersand usage in citations appears correct.",
      applicable: issues.length, checked: issues.length, matched: issues.length,
      failed: 0, unknown: 0, found: "Ampersand usage OK",
      applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  const details = issues.map(({ type, preview }) =>
    type === "and-in-parens"
      ? `Use "&" not "and" inside parenthetical citation: "${preview}"`
      : `Use "and" not "&" in narrative citation: "${preview}"`,
  );

  return {
    rule, status: "fail", passed: false, expected, expectedText: expected,
    foundText: `${issues.length} citation${issues.length === 1 ? "" : "s"} appear to use the wrong form of "and"/"&".`,
    applicable: issues.length, checked: issues.length, matched: 0,
    failed: issues.length, unknown: 0,
    found: `${issues.length} ampersand issue(s)`,
    applicableParagraphs: 0, details,
    howToFix: getHowToFix(rule), resources: [APA_IN_TEXT_FORMAT_RESOURCE],
  };
}

function checkCitationEtAl(extracted, referencesHeading) {
  const rule = "Citation et al. format";
  const expected = 'Write "et al." with a period after "al" only — not "et. al.", "et al", or "etal".';
  const bodyText = getBodyText(extracted, referencesHeading);
  const issues = [];

  // Malformed variants near a year (to avoid flagging random prose)
  const badEtAlForms = [
    { re: /\bet\.\s+al\.?\b/g, label: 'Use "et al." not "et. al."' },
    { re: /\betal\b/g, label: 'Use "et al." not "etal"' },
    { re: /\bet\s+al\b(?!\.)/g, label: 'Write "et al." with a period: "et al."' },
  ];

  for (const { re, label } of badEtAlForms) {
    let m;
    while ((m = re.exec(bodyText)) !== null) {
      const context = bodyText.slice(Math.max(0, m.index - 20), m.index + m[0].length + 20);
      if (/\d{4}|n\.d\./.test(context)) {
        issues.push(`${label} — found near: "…${context.trim()}…"`);
      }
    }
  }

  if (issues.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: '"et al." formatting appears correct.',
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "et al. OK", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "fail", passed: false, expected, expectedText: expected,
    foundText: `${issues.length} instance${issues.length === 1 ? "" : "s"} of malformed "et al." detected.`,
    applicable: issues.length, checked: issues.length, matched: 0,
    failed: issues.length, unknown: 0, found: `${issues.length} et al. issue(s)`,
    applicableParagraphs: 0, details: issues,
    howToFix: getHowToFix(rule), resources: [APA_IN_TEXT_FORMAT_RESOURCE],
  };
}

function checkCitationComma(extracted, referencesHeading) {
  const rule = "Citation comma";
  const expected = "APA citations require a comma between the author name and the year: (Author, Year).";
  const bodyText = getBodyText(extracted, referencesHeading);
  const issues = [];
  const seen = new Set();

  const parenRe = /\(([^()]{2,200})\)/g;
  let m;
  while ((m = parenRe.exec(bodyText)) !== null) {
    const content = m[1].trim();
    if (!/\d{4}|n\.d\./.test(content)) continue;
    if (/\bpersonal\s+communication\b/i.test(content)) continue;
    if (/\bas\s+cited\s+in\b/i.test(content)) continue;

    // Missing comma: author-like content starts with a capital letter,
    // and a non-comma char is followed directly by a space and the year.
    if (/^[A-ZÀ-Ÿ]/.test(content) && /[^,\s]\s+\d{4}[a-z]?/.test(content)) {
      const key = `comma:${m[0]}`;
      if (!seen.has(key)) {
        seen.add(key);
        issues.push(`Missing comma before year — found: "${m[0]}"`);
      }
    }

    // Stray trailing comma after year: (Author, 2023,) or (2023,)
    if (/\d{4}[a-z]?,\s*$/.test(content)) {
      const context = bodyText.slice(Math.max(0, m.index - 20), m.index + m[0].length + 5).trim();
      const key = `stray:${m[0]}`;
      if (!seen.has(key)) {
        seen.add(key);
        issues.push(`Trailing comma after year — remove the comma near: "…${context}…"`);
      }
    }
  }

  if (issues.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: "Citation comma formatting appears correct.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "Citation commas OK", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "fail", passed: false, expected, expectedText: expected,
    foundText: `${issues.length} citation${issues.length === 1 ? "" : "s"} appear to have a missing or extra comma.`,
    applicable: issues.length, checked: issues.length, matched: 0,
    failed: issues.length, unknown: 0, found: `${issues.length} citation comma issue(s)`,
    applicableParagraphs: 0, details: issues,
    howToFix: getHowToFix(rule), resources: [APA_IN_TEXT_FORMAT_RESOURCE],
  };
}

function checkCitationNoDate(extracted, referencesHeading) {
  const rule = "Citation no-date format";
  const expected = 'Write "n.d." — lowercase with periods after each letter — when a source has no date.';
  const bodyText = getBodyText(extracted, referencesHeading);
  const issues = [];

  // Wrong forms of n.d. near a citation context (inside parens)
  const badNdForms = [
    { re: /\bN\.D\.\b/g, label: 'Write "n.d." (lowercase) not "N.D."' },
    { re: /\bn\.d\b(?!\.)/g, label: 'Write "n.d." with a period after "d"' },
    { re: /\bno\s+date\b/gi, label: 'Write "n.d." not "no date" in a citation' },
    { re: /\bn\.?\s*d\b(?!\.)/g, label: 'Write "n.d." — both letters need periods' },
  ];

  for (const { re, label } of badNdForms) {
    let m;
    while ((m = re.exec(bodyText)) !== null) {
      // Only flag when inside or near a parenthetical citation context
      const surrounding = bodyText.slice(Math.max(0, m.index - 30), m.index + m[0].length + 30);
      if (/[()]/.test(surrounding)) {
        const preview = surrounding.trim().replace(/\s+/g, " ");
        issues.push(`${label}: "…${preview}…"`);
      }
    }
  }

  if (issues.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: '"n.d." formatting appears correct.',
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "n.d. OK", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "fail", passed: false, expected, expectedText: expected,
    foundText: `${issues.length} instance${issues.length === 1 ? "" : "s"} of incorrectly formatted "n.d." detected.`,
    applicable: issues.length, checked: issues.length, matched: 0,
    failed: issues.length, unknown: 0, found: `${issues.length} n.d. issue(s)`,
    applicableParagraphs: 0, details: issues,
    howToFix: getHowToFix(rule), resources: [APA_IN_TEXT_FORMAT_RESOURCE],
  };
}

function checkCitationPageFormat(extracted, referencesHeading) {
  const rule = "Citation page format";
  const expected = 'APA uses "p." for single page and "pp." for a page range in citations — not "pg.", "pgs.", etc.';
  const bodyText = getBodyText(extracted, referencesHeading);
  const issues = [];

  // Scan inside parenthetical citation context for wrong page prefix
  const parenRe = /\(([^()]{2,200})\)/g;
  let m;
  while ((m = parenRe.exec(bodyText)) !== null) {
    const inner = m[1];
    if (!/\b\d{4}[a-z]?\b|n\.d\./.test(inner)) continue;
    if (/\bpgs?\.\s*\d/i.test(inner)) {
      const preview = m[0].length > 70 ? m[0].slice(0, 70) + "…)" : m[0];
      issues.push(`Use "p." or "pp." not "pg."/"pgs.": "${preview}"`);
    }
  }

  if (issues.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: "Citation page number formatting appears correct.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "Citation page format OK", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "fail", passed: false, expected, expectedText: expected,
    foundText: `${issues.length} citation${issues.length === 1 ? "" : "s"} appear to use an incorrect page prefix.`,
    applicable: issues.length, checked: issues.length, matched: 0,
    failed: issues.length, unknown: 0, found: `${issues.length} page format issue(s)`,
    applicableParagraphs: 0, details: issues,
    howToFix: getHowToFix(rule), resources: [APA_IN_TEXT_FORMAT_RESOURCE],
  };
}

function checkCitationMultipleSources(extracted, referencesHeading) {
  const rule = "Citation multiple sources";
  const expected = "When citing multiple sources in one parenthetical, separate them with semicolons: (Smith, 2020; Jones, 2021).";
  const bodyText = getBodyText(extracted, referencesHeading);
  const issues = [];

  // Look for (Author, YYYY, Author, YYYY) — two years separated by comma + author, not semicolon
  const parenRe = /\(([^()]{2,300})\)/g;
  let m;
  while ((m = parenRe.exec(bodyText)) !== null) {
    const inner = m[1];
    if (!/\b\d{4}[a-z]?\b|n\.d\./.test(inner)) continue;
    if (/\bpersonal\s+communication\b/i.test(inner)) continue;
    // If there's a year followed by comma+space+capital letter (suggests another author, not page/issue)
    if (/\d{4}[a-z]?,\s+[A-Z]/.test(inner) && !inner.includes(";")) {
      const preview = m[0].length > 80 ? m[0].slice(0, 80) + "…)" : m[0];
      issues.push(`Use ";" to separate multiple citations: "${preview}"`);
    }
  }

  if (issues.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: "Multiple-source citation formatting appears correct.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "Multiple source citations OK", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "fail", passed: false, expected, expectedText: expected,
    foundText: `${issues.length} parenthetical citation${issues.length === 1 ? "" : "s"} appear to use a comma instead of a semicolon between sources.`,
    applicable: issues.length, checked: issues.length, matched: 0,
    failed: issues.length, unknown: 0, found: `${issues.length} multiple-source issue(s)`,
    applicableParagraphs: 0, details: issues,
    howToFix: getHowToFix(rule), resources: [APA_IN_TEXT_FORMAT_RESOURCE],
  };
}

function checkCitationYearSuffixes(extracted, referencesHeading) {
  const rule = "Citation year suffix";
  const expected = 'When citing two works by the same author in the same year, add a letter suffix: (Smith, 2020a) and (Smith, 2020b). The suffix must also appear in the reference entry.';
  const bodyText = getBodyText(extracted, referencesHeading);

  // Find all citation year suffixes in body text
  const suffixRe = /\b(\d{4})([a-z])\b/g;
  const citationSuffixes = new Map(); // "Smith-2020" → ["a", "b"]
  const parenRe = /\(([^()]{2,300})\)/g;
  let m;
  while ((m = parenRe.exec(bodyText)) !== null) {
    const inner = m[1];
    if (/\bpersonal\s+communication\b/i.test(inner)) continue;
    for (const seg of inner.split(";")) {
      const yearSuffixM = seg.match(/\b(\d{4})([a-z])\b/);
      if (!yearSuffixM) continue;
      const year = yearSuffixM[1];
      const suffix = yearSuffixM[2];
      const beforeYear = seg.slice(0, seg.search(/\b\d{4}/)).replace(/[,\s]+$/, "").trim();
      if (!beforeYear) continue;
      const key = `${beforeYear}-${year}`;
      if (!citationSuffixes.has(key)) citationSuffixes.set(key, new Set());
      citationSuffixes.get(key).add(suffix);
    }
  }

  if (citationSuffixes.size === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: "No year-suffix citations detected.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No year suffixes found", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  // Build reference key set for cross-checking
  const referenceParagraphs = referencesHeading
    ? getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading).filter(looksLikeReferenceEntryStart)
    : [];
  const refKeys = new Set(
    referenceParagraphs.map((p) => {
      const k = extractReferenceKey(p.text);
      return k ? `${k.lastName}-${k.year}` : null;
    }).filter(Boolean),
  );

  const failures = [];
  const details = [];

  for (const [key, suffixSet] of citationSuffixes) {
    for (const suffix of suffixSet) {
      const fullKey = `${key}${suffix}`;
      // Check if a reference entry with this year+suffix exists
      if (refKeys.size > 0 && !refKeys.has(fullKey)) {
        failures.push(key);
        details.push(`Citation suffix "${key}${suffix}" has no matching reference entry — add a reference with year "${key.split("-")[1]}${suffix}".`);
      }
    }
  }

  const total = citationSuffixes.size;
  const status = failures.length > 0 ? "fail" : "pass";

  return {
    rule, status, passed: status === "pass", expected, expectedText: expected,
    foundText:
      status === "pass"
        ? `Year-suffix citations detected and reference entries appear to match.`
        : `${failures.length} year-suffix citation${failures.length === 1 ? "" : "s"} appear to have no matching reference entry.`,
    applicable: total, checked: total, matched: total - failures.length,
    failed: failures.length, unknown: 0,
    found: status === "pass" ? "Year suffixes matched" : `${failures.length} suffix(es) unmatched`,
    applicableParagraphs: 0, details,
    howToFix: status === "fail" ? getHowToFix(rule) : [], resources: [],
  };
}

// Detects secondary citations ("as cited in") and returns them as "review" warnings.
function checkSecondaryCitations(extracted, referencesHeading) {
  const rule = "Secondary citations";
  const expected = 'APA discourages secondary citations ("as cited in"). Locate the original source when possible.';
  const bodyText = getBodyText(extracted, referencesHeading);
  const re = /\bas\s+cited\s+in\b[^)]{0,100}/gi;
  const found = [];
  let m;
  while ((m = re.exec(bodyText)) !== null) {
    found.push(bodyText.slice(m.index, m.index + Math.min(80, m[0].length + 20)).trim());
  }

  if (found.length === 0) {
    return {
      rule, status: "pass", passed: true, expected, expectedText: expected,
      foundText: "No secondary citations detected.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No secondary citations", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "review", passed: false, expected, expectedText: expected,
    foundText: `${found.length} secondary citation${found.length === 1 ? "" : "s"} ("as cited in") detected. APA discourages these — try to find and cite the original source.`,
    applicable: found.length, checked: found.length, matched: 0,
    failed: 0, unknown: found.length,
    found: `${found.length} secondary citation(s)`,
    applicableParagraphs: 0,
    details: found.map((s) => `Secondary citation: "…${s}…"`),
    howToFix: [], resources: [APA_IN_TEXT_FORMAT_RESOURCE],
  };
}

function checkApaFormatting(extracted) {
  const referencesHeading = findReferencesHeadingNearEnd(extracted.paragraphs);
  const bodyParagraphs = getParagraphsByRole(extracted, "body").filter(
    (paragraph) => !referencesHeading || paragraph.index < referencesHeading.index,
  );

  const bodyText = extracted.paragraphs
    .filter((p) => p.role !== "blank" && (!referencesHeading || p.index < referencesHeading.index))
    .map((p) => p.text)
    .join(" ");
  const hasCitations = extractInlineCitationKeys(bodyText).length > 0;

  const checks = [
    checkPageNumbering(extracted),
    checkTitlePage(extracted),
    checkReferencesPage(extracted),
    checkInlineCitations(extracted, referencesHeading),
    checkPersonalCommunications(extracted, referencesHeading),
    checkSecondaryCitations(extracted, referencesHeading),
    checkCitationAmpersandUsage(extracted, referencesHeading),
    checkCitationEtAl(extracted, referencesHeading),
    checkCitationComma(extracted, referencesHeading),
    checkCitationNoDate(extracted, referencesHeading),
    checkCitationPageFormat(extracted, referencesHeading),
    checkCitationMultipleSources(extracted, referencesHeading),
    checkCitationYearSuffixes(extracted, referencesHeading),
    ...(referencesHeading ? [
      ...(referencesHeading.synthetic ? [] : [checkReferencesHeadingAlignment(referencesHeading)]),
      checkReferencesFormatting(extracted, referencesHeading),
      ...(hasCitations ? [
        checkUncitedReferences(extracted, referencesHeading),
        checkUnmatchedCitations(extracted, referencesHeading),
      ] : []),
      checkReferenceDOIs(extracted, referencesHeading),
      checkReferenceDOIFormat(extracted, referencesHeading),
      checkReferenceForbiddenPhrases(extracted, referencesHeading),
      checkReferenceShortLinks(extracted, referencesHeading),
      checkUnapprovedSources(extracted, referencesHeading),
      checkReferenceAuthors(extracted, referencesHeading),
      checkReferenceYear(extracted, referencesHeading),
      checkReferenceTitleCapitalization(extracted, referencesHeading),
      checkReferenceItalics(extracted, referencesHeading),
      checkReferencePunctuation(extracted, referencesHeading),
    ] : []),
    checkMargins(extracted),
    checkLineSpacingForParagraphs(bodyParagraphs, "Body line spacing", "Body"),
    checkLineSpacingForRole(extracted, "heading", "Heading"),
    ...(referencesHeading ? [checkReferencesLineSpacing(extracted, referencesHeading)] : []),
    checkParagraphSpacingForRole(extracted, "body", "Body"),
    checkParagraphSpacingForRole(extracted, "heading", "Heading"),
    checkFirstLineIndents(extracted, referencesHeading),
    checkAlignment(extracted),
    checkFonts(extracted),
    checkUnconvertedMarkup(extracted, referencesHeading),
  ];

  const referenceGroups = referencesHeading
    ? groupReferenceEntries(getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading))
    : [];

  return { checks, referenceGroups };
}

module.exports = {
  checkApaFormatting,
  checkPageNumbering,
  checkTitlePage,
  checkReferencesPage,
  checkReferencesHeadingAlignment,
  checkReferencesFormatting,
  checkMargins,
  checkLineSpacingForRole,
  checkReferencesLineSpacing,
  checkParagraphSpacingForRole,
  checkFirstLineIndents,
  checkAlignment,
  checkInlineCitations,
  checkUncitedReferences,
  checkUnmatchedCitations,
  checkReferenceDOIs,
  checkReferenceShortLinks,
  checkUnapprovedSources,
  checkFonts,
  checkUnconvertedMarkup,
  parseReferenceEntry,
  classifyReferenceKind,
  checkReferenceAuthors,
  checkReferenceYear,
  checkReferenceTitleCapitalization,
  checkReferenceItalics,
  checkReferencePunctuation,
  checkReferenceDOIFormat,
  checkReferenceForbiddenPhrases,
  checkPersonalCommunications,
  checkSecondaryCitations,
  checkCitationAmpersandUsage,
  checkCitationEtAl,
  checkCitationComma,
  checkCitationNoDate,
  checkCitationPageFormat,
  checkCitationMultipleSources,
  checkCitationYearSuffixes,
};

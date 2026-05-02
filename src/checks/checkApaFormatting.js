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
const ALIGN_TEXT_MSFT_RESOURCE = {
  label: "Microsoft Support: Align text",
  url: "https://support.microsoft.com/en-us/office/align-text-left-or-right-center-text-or-justify-text-on-a-page-70da744d-0f4d-472e-916d-1c42d94dc33f",
};
const ALIGN_TEXT_GFG_RESOURCE = {
  label: "GeeksforGeeks: Text Alignment in MS Word",
  url: "https://www.geeksforgeeks.org/ms-word/text-alignment-in-ms-word/",
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
  if (rule === "Unapproved source") {
    return [
      "This source is on AIU's list of websites not approved for academic use.",
      "Replace it with a peer-reviewed article, textbook, or credible institutional source.",
      "Search your AIU library database (ProQuest, EBSCOhost) for a credible alternative on the same topic.",
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

  if (rule === "Inline citations") {
    return [
      "Add an inline citation every time you use information from a source.",
      "Parenthetical format: place (Author, Year) at the end of the sentence before the period.",
      "Narrative format: use Author (Year) at the start of the sentence.",
    ];
  }

  if (rule === "Uncited references") {
    return [
      "For each reference listed, find where you used that source and add an inline citation.",
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

  if (rule === "References formatting") {
    return [
      "If any entries are bare URLs: replace each URL-only line with a full APA reference. A complete reference includes the author's last name and initials, publication year in parentheses, title of the work, and the URL or DOI.",
      "Example: Smith, J. A. (2023). Title of article. Site Name. https://www.example.com/article",
      "A hanging indent means the first line is flush left and all following lines are indented 0.5\".",
      "Select all reference entries in Microsoft Word.",
      "Click the Line Spacing button (⇵☰) in the Home toolbar to open its menu.",
      "Click 'Line Spacing Options...'",
      "In the Paragraph window, under Indentation, set Special to Hanging and By to 0.5\".",
      "Click OK.",
    ];
  }

  return [];
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
  );
}

function checkLineSpacingForRole(extracted, role, label) {
  return checkLineSpacingForParagraphs(getParagraphsByRole(extracted, role), `${label} line spacing`, label);
}

function checkParagraphSpacingForRole(extracted, role, label) {
  const applicableParagraphs = getParagraphsByRole(extracted, role);
  const partialUnknowns = applicableParagraphs.filter((paragraph) => {
    return (
      !paragraph.formatting.spacingBeforePoints.known || !paragraph.formatting.spacingAfterPoints.known
    );
  });
  const failures = applicableParagraphs.filter((paragraph) => {
    const before = paragraph.formatting.spacingBeforePoints;
    const after = paragraph.formatting.spacingAfterPoints;

    return (
      before.known &&
      after.known &&
      (!isClose(before.value, EXPECTED_PARAGRAPH_SPACING_POINTS) ||
        !isClose(after.value, EXPECTED_PARAGRAPH_SPACING_POINTS))
    );
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

function checkFirstLineIndents(extracted) {
  const applicableParagraphs = getParagraphsByRole(extracted, "body");
  const unknowns = splitUnknowns(applicableParagraphs, "firstLineIndentInches");
  const checked = applicableParagraphs.length - unknowns.length;
  const failures = applicableParagraphs.filter((paragraph) => {
    const indent = paragraph.formatting.firstLineIndentInches;
    return indent.known && !isClose(indent.value, EXPECTED_FIRST_LINE_INDENT_INCHES);
  });
  const matched = checked - failures.length;

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
  } else if (failures.length > 0 && matched === 0) {
    status = "fail";
    const n = failures.length;
    foundText =
      n === applicableParagraphs.length
        ? "All body paragraphs are missing the required first-line indent."
        : `${n} body ${n === 1 ? "paragraph is" : "paragraphs are"} missing the required first-line indent.`;
  } else if (failures.length > 0) {
    status = "review";
    foundText =
      "Some paragraphs may not be using Word's built-in first-line indent. " +
      "They may appear indented but could be using tabs or spaces instead.";
  } else {
    status = "review";
    foundText = "These paragraphs appear correct, but may not be using Word's built-in first-line indent.";
  }

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
    (paragraph, position) => !isBodyTitle(paragraph, position),
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

  if (applicableParagraphs.length === 0) {
    status = "review";
    foundText = "APA Coach did not find any body paragraphs to check.";
  } else if (failures.length === 0 && unknowns.length === 0) {
    status = "pass";
    foundText = "All body paragraphs use left alignment.";
  } else if (failures.length > 0 && matched === 0) {
    status = "fail";
    foundText = "Body paragraphs are not left aligned. APA requires left alignment for all body text.";
  } else if (failures.length > 0) {
    status = "review";
    foundText = "Some body paragraphs may not be left aligned.";
  } else {
    status = "review";
    foundText = "These paragraphs appear correct, but may not be using Word's built-in alignment setting.";
  }

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

function extractInlineCitationKeys(bodyText) {
  const keys = [];
  const seen = new Set();

  const parenRe = /\(([^()]{2,300})\)/g;
  let m;
  while ((m = parenRe.exec(bodyText)) !== null) {
    const content = m[1];
    if (!/\d{4}|n\.d\./.test(content)) continue;
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
      // For multi-author paren citations, take only the first author
      const lastName = withoutEtAl.split(/\s*&\s*|\s+and\s+/i)[0].trim() || withoutEtAl;
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
  return /https?:\/\/|doi\.org|\bdoi:\s*10\./i.test(text);
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
    rule: "References formatting",
    status,
    passed: status === "pass",
    expected: "APA expects references to use hanging indents and consistent formatting.",
    expectedText: "APA expects references to use hanging indents and consistent formatting.",
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
        ? "References formatting detected"
        : `Formatting issues detected: ${issues.length}`,
    applicableParagraphs: referenceParagraphs.length,
    details: issues,
    howToFix: status === "fail" ? getHowToFix("References formatting") : [],
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
    "APA requires references to be double spaced.",
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
  const rule = "Inline citations";
  const expected = "Each source used in the paper should have an inline citation.";

  const allText = extracted.paragraphs
    .filter((p) => p.role !== "blank" && (!referencesHeading || p.index < referencesHeading.index))
    .map((p) => p.text)
    .join(" ");
  const citationKeys = extractInlineCitationKeys(allText);

  if (citationKeys.length > 0) {
    return {
      rule, status: "pass", passed: true,
      expected, expectedText: expected,
      foundText: `APA Coach found ${citationKeys.length} inline citation(s) in the document.`,
      applicable: citationKeys.length, checked: citationKeys.length, matched: citationKeys.length,
      failed: 0, unknown: 0, found: `${citationKeys.length} inline citation(s) found`,
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
      foundText: "APA Coach did not find enough content to check for inline citations.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "Not enough content", applicableParagraphs: 0, details: [], howToFix: [], resources: [],
    };
  }

  return {
    rule, status: "fail", passed: false,
    expected, expectedText: expected,
    foundText: "APA Coach did not find any inline citations. If you used sources, add (Author, Year) citations throughout the body.",
    applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
    found: "No inline citations found", applicableParagraphs: 0,
    details: [], howToFix: getHowToFix(rule), resources: [],
  };
}

function checkUncitedReferences(extracted, referencesHeading) {
  const referenceParagraphs = getReferenceEntryParagraphs(extracted.paragraphs, referencesHeading);
  const entryParagraphs = referenceParagraphs.filter(looksLikeReferenceEntryStart);
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
      expected: "Each reference should have at least one matching inline citation.",
      expectedText: "Each reference should have at least one matching inline citation.",
      foundText: "APA Coach could not parse the reference entries to check for inline citations.",
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
    expected: "Each reference should have at least one matching inline citation.",
    expectedText: "Each reference should have at least one matching inline citation.",
    foundText:
      status === "pass"
        ? `All ${referenceKeys.length} references appear to have a matching inline citation.`
        : !bodyHasCitations
          ? "APA Coach could not find any inline citations in the body. If your paper uses citations, check that they follow APA format: (Author, Year)."
          : `${uncited.length} of ${referenceKeys.length} reference${referenceKeys.length === 1 ? "" : "s"} appear to have no matching inline citation.`,
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
  const entryParagraphs = referenceParagraphs.filter(looksLikeReferenceEntryStart);
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
      expected: "Each inline citation should have a matching entry in the References list.",
      expectedText: "Each inline citation should have a matching entry in the References list.",
      foundText: "APA Coach could not find any inline citations in the body to check.",
      applicable: 0, checked: 0, matched: 0, failed: 0, unknown: 0,
      found: "No inline citations found",
      applicableParagraphs: 0,
      details: [], howToFix: [], resources: [],
    };
  }

  if (referenceKeys.length === 0) {
    return {
      rule: "Unmatched citations",
      status: "review",
      passed: false,
      expected: "Each inline citation should have a matching entry in the References list.",
      expectedText: "Each inline citation should have a matching entry in the References list.",
      foundText: "APA Coach could not parse the References list to check against inline citations.",
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
        ? `All ${citationKeys.length} inline citations appear to have a matching reference.`
        : `${unmatched.length} of ${citationKeys.length} inline citation${citationKeys.length === 1 ? "" : "s"} appear to have no matching reference entry.`,
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
  const entryParagraphs = referenceParagraphs.filter(looksLikeReferenceEntryStart);

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

  return [
    checkTitlePage(extracted),
    checkReferencesPage(extracted),
    checkInlineCitations(extracted, referencesHeading),
    ...(referencesHeading ? [
      ...(referencesHeading.synthetic ? [] : [checkReferencesHeadingAlignment(referencesHeading)]),
      checkReferencesFormatting(extracted, referencesHeading),
      ...(hasCitations ? [
        checkUncitedReferences(extracted, referencesHeading),
        checkUnmatchedCitations(extracted, referencesHeading),
      ] : []),
      checkReferenceDOIs(extracted, referencesHeading),
      checkReferenceShortLinks(extracted, referencesHeading),
      checkUnapprovedSources(extracted, referencesHeading),
    ] : []),
    checkMargins(extracted),
    checkLineSpacingForParagraphs(bodyParagraphs, "Body line spacing", "Body"),
    checkLineSpacingForRole(extracted, "heading", "Heading"),
    ...(referencesHeading ? [checkReferencesLineSpacing(extracted, referencesHeading)] : []),
    checkParagraphSpacingForRole(extracted, "body", "Body"),
    checkParagraphSpacingForRole(extracted, "heading", "Heading"),
    checkFirstLineIndents(extracted),
    checkAlignment(extracted),
    checkFonts(extracted),
  ];
}

module.exports = {
  checkApaFormatting,
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
};

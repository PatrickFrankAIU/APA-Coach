const EXPECTED_MARGIN_INCHES = 1;
const EXPECTED_LINE_SPACING = 2;
const EXPECTED_PARAGRAPH_SPACING_POINTS = 0;
const EXPECTED_FIRST_LINE_INDENT_INCHES = 0.5;
const EXPECTED_ALIGNMENT = "left";
const TOLERANCE = 0.01;
const MISSING_HANGING_INDENT_ISSUE = "One or more references are missing a 0.5-inch hanging indent.";
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
      "A hanging indent means the first line is flush left and all following lines are indented 0.5\".",
      "Select all reference entries in Microsoft Word.",
      "Click the Line Spacing button (⇵☰) in the Home toolbar to open its menu.",
      "Click 'Line Spacing Options...'",
      "In the Paragraph window, under Indentation, set Special to Hanging and By to 0.5\".",
      "Click OK.",
      "Alternative: show the ruler (View > Ruler), then drag the lower triangle (◭) to the 0.5\" mark.",
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
  return paragraph.text.trim().toLowerCase() === "references";
}

function findReferencesHeadingNearEnd(paragraphs) {
  const firstNearEndIndex = Math.floor(paragraphs.length * 0.75);

  return paragraphs.find((paragraph) => paragraph.index > firstNearEndIndex && isReferencesHeading(paragraph)) || null;
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
  const yearMatch = text.match(/\((\d{4}[a-z]?|n\.d\.)\)/);
  if (!yearMatch) return null;
  const year = yearMatch[1];
  const beforeYear = text.substring(0, yearMatch.index).trim().replace(/\.\s*$/, "").trim();
  const commaIdx = beforeYear.indexOf(",");
  const lastName = (commaIdx > 0 ? beforeYear.substring(0, commaIdx) : beforeYear).trim();
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
    if (!/\d{4}/.test(content)) continue;
    const segments = content.split(";");
    for (const seg of segments) {
      const yearM = seg.match(/\b(\d{4}[a-z]?)\b/);
      if (!yearM) continue;
      const year = yearM[1];
      const nameM = seg.trim().match(/^([A-ZÀ-Ÿ][A-Za-zÀ-ÿ''-]+(?:\s+[A-ZÀ-Ÿ][A-Za-zÀ-ÿ''-]+)*)/);
      if (!nameM) continue;
      const lastName = nameM[1].replace(/\s+et$/, "").trim();
      const key = `${lastName.toLowerCase()}|${year.replace(/[a-z]$/, "")}`;
      if (!seen.has(key)) {
        seen.add(key);
        keys.push({ lastName, year, source: seg.trim() });
      }
    }
  }

  const narrativeRe = /\b([A-ZÀ-Ÿ][A-Za-zÀ-ÿ''-]+)(?:\s+et\s+al\.?)?\s+\((\d{4}[a-z]?)\)/g;
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

function isReferenceCited(key, bodyText) {
  const { lastName, year } = key;
  const escaped = lastName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const yearPat = year === "n.d." ? "n\\.d\\." : year.replace(/[a-z]$/, "") + "[a-z]?";
  const re = new RegExp(`\\b${escaped}(?:\\s+et\\s+al\\.?)?[^\\n.]{0,40}${yearPat}\\b`, "i");
  return re.test(bodyText);
}

function isCitationMatched(citKey, referenceKeys) {
  const normalLast = citKey.lastName.toLowerCase();
  const yearBase = citKey.year.replace(/[a-z]$/, "");
  return referenceKeys.some(
    (refKey) =>
      refKey.lastName.toLowerCase() === normalLast &&
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

function findReferencesFormattingIssues(referencesHeading, referenceParagraphs) {
  const issues = [];
  const shortParagraphs = referenceParagraphs.filter((paragraph) => paragraph.text.trim().length < 40);
  const missingHangingIndentParagraphs = referenceParagraphs.filter((paragraph) => !hasExpectedHangingIndent(paragraph));

  if (referenceParagraphs.length === 0) {
    issues.push("No reference entries were detected after the References heading.");
  }

  if (shortParagraphs.length >= 2) {
    issues.push("Some reference entries appear incomplete or separated incorrectly.");
  }

  if (missingHangingIndentParagraphs.length > 0) {
    issues.push(MISSING_HANGING_INDENT_ISSUE);
  }

  if (hasBrokenReferenceEntries(referenceParagraphs)) {
    issues.push("Some reference entries appear to be broken across separate paragraphs.");
  }

  return issues;
}

function checkReferencesPage(extracted) {
  const referencesHeading = findReferencesHeadingNearEnd(extracted.paragraphs);
  const status = referencesHeading ? "pass" : "fail";

  return {
    rule: "References page",
    status,
    passed: status === "pass",
    expected: "APA expects a References page with sources listed in alphabetical order using hanging indents.",
    expectedText: "APA expects a References page with sources listed in alphabetical order using hanging indents.",
    foundText:
      status === "pass"
        ? "A References page was detected at the end of the document."
        : "APA Coach did not detect a References page at the end of the document.",
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
  const uncited = referenceKeys.filter((key) => !isReferenceCited(key, bodyText));
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

function checkApaFormatting(extracted) {
  const referencesHeading = findReferencesHeadingNearEnd(extracted.paragraphs);
  const bodyParagraphs = getParagraphsByRole(extracted, "body").filter(
    (paragraph) => !referencesHeading || paragraph.index < referencesHeading.index,
  );

  return [
    checkTitlePage(extracted),
    checkReferencesPage(extracted),
    ...(referencesHeading ? [
      checkReferencesHeadingAlignment(referencesHeading),
      checkReferencesFormatting(extracted, referencesHeading),
      checkUncitedReferences(extracted, referencesHeading),
      checkUnmatchedCitations(extracted, referencesHeading),
      checkReferenceDOIs(extracted, referencesHeading),
    ] : []),
    checkMargins(extracted),
    checkLineSpacingForParagraphs(bodyParagraphs, "Body line spacing", "Body"),
    checkLineSpacingForRole(extracted, "heading", "Heading"),
    ...(referencesHeading ? [checkReferencesLineSpacing(extracted, referencesHeading)] : []),
    checkParagraphSpacingForRole(extracted, "body", "Body"),
    checkParagraphSpacingForRole(extracted, "heading", "Heading"),
    checkFirstLineIndents(extracted),
    checkAlignment(extracted),
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
  checkUncitedReferences,
  checkUnmatchedCitations,
  checkReferenceDOIs,
};

const { XMLParser } = require("fast-xml-parser");

const TWIPS_PER_INCH = 1440;
const INFERRED_SOURCE = "inferredDefault";

function toArray(value) {
  if (!value) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

function valueWithSource(value, source) {
  return {
    value,
    source: value === null || value === undefined ? "unknown" : source,
    known: value !== null && value !== undefined,
  };
}

function twipsToInches(value) {
  if (value === undefined || value === null || value === "") {
    return null;
  }

  const number = Number(value);
  return Number.isFinite(number) ? number / TWIPS_PER_INCH : null;
}

function twipsToPoints(value) {
  const inches = twipsToInches(value);
  return inches === null ? null : inches * 72;
}

function lineValueToMultiple(line, lineRule) {
  if (line === undefined || line === null || line === "") {
    return null;
  }

  const number = Number(line);
  if (!Number.isFinite(number)) {
    return null;
  }

  if (!lineRule || lineRule === "auto") {
    return number / 240;
  }

  return null;
}

function readDocxXml(zip, entryName) {
  const entry = zip.getEntry(entryName);
  if (!entry) {
    return null;
  }

  return entry.getData().toString("utf8");
}

function parseXml(xml) {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    removeNSPrefix: true,
  });

  return parser.parse(xml);
}

function findSectionProperties(document) {
  const body = document.document && document.document.body;
  if (!body) {
    return null;
  }

  if (body.sectPr) {
    return body.sectPr;
  }

  const paragraphs = toArray(body.p);
  for (let index = paragraphs.length - 1; index >= 0; index -= 1) {
    const paragraph = paragraphs[index];
    if (paragraph.pPr && paragraph.pPr.sectPr) {
      return paragraph.pPr.sectPr;
    }
  }

  return null;
}

function extractMargins(sectionProperties) {
  const margins = sectionProperties && sectionProperties.pgMar;

  return {
    topInches: valueWithSource(margins ? twipsToInches(margins.top) : null, "direct"),
    rightInches: valueWithSource(margins ? twipsToInches(margins.right) : null, "direct"),
    bottomInches: valueWithSource(margins ? twipsToInches(margins.bottom) : null, "direct"),
    leftInches: valueWithSource(margins ? twipsToInches(margins.left) : null, "direct"),
  };
}

function getParagraphText(paragraph) {
  const runs = toArray(paragraph.r);
  const parts = [];

  for (const run of runs) {
    const textNodes = toArray(run.t);
    for (const text of textNodes) {
      if (typeof text === "string") {
        parts.push(text);
      } else if (text && typeof text["#text"] === "string") {
        parts.push(text["#text"]);
      }
    }
  }

  return parts.join("");
}

function getStyleId(paragraphProperties) {
  return paragraphProperties && paragraphProperties.pStyle ? paragraphProperties.pStyle.val || null : null;
}

function buildStyleMap(stylesDocument) {
  const styles = stylesDocument && stylesDocument.styles ? toArray(stylesDocument.styles.style) : [];
  const paragraphStyles = {};
  let defaultParagraphStyleId = null;

  for (const style of styles) {
    if (style.type !== "paragraph" || !style.styleId) {
      continue;
    }

    paragraphStyles[style.styleId] = {
      id: style.styleId,
      name: style.name ? style.name.val || null : null,
      basedOn: style.basedOn ? style.basedOn.val || null : null,
      properties: style.pPr || {},
    };

    if (style.default === "1" || style.default === true) {
      defaultParagraphStyleId = style.styleId;
    }
  }

  return {
    paragraphStyles,
    defaultParagraphStyleId,
  };
}

function extractDocumentDefaults(stylesDocument) {
  const defaults = stylesDocument && stylesDocument.styles ? stylesDocument.styles.docDefaults : null;
  const paragraphProperties = defaults && defaults.pPrDefault ? defaults.pPrDefault.pPr || {} : {};
  const defaultLineSpacing = findProperty(paragraphProperties, "lineSpacing");
  const defaultSpacingBefore = findProperty(paragraphProperties, "spacingBefore");
  const defaultSpacingAfter = findProperty(paragraphProperties, "spacingAfter");
  const defaultFirstLineIndent = findProperty(paragraphProperties, "firstLineIndent");
  const defaultHangingIndent = findProperty(paragraphProperties, "hangingIndent");
  const defaultAlignment = findProperty(paragraphProperties, "alignment");

  return {
    paragraphProperties,
    formatting: {
      lineSpacingMultiple: valueWithSource(
        defaultLineSpacing ? lineValueToMultiple(defaultLineSpacing.line, defaultLineSpacing.lineRule) : null,
        "default",
      ),
      spacingBeforePoints: valueWithSource(
        defaultSpacingBefore !== undefined ? twipsToPoints(defaultSpacingBefore) : null,
        "default",
      ),
      spacingAfterPoints: valueWithSource(
        defaultSpacingAfter !== undefined ? twipsToPoints(defaultSpacingAfter) : null,
        "default",
      ),
      firstLineIndentInches: valueWithSource(
        defaultFirstLineIndent !== undefined ? twipsToInches(defaultFirstLineIndent) : 0,
        defaultFirstLineIndent !== undefined ? "default" : INFERRED_SOURCE,
      ),
      hangingIndentInches: valueWithSource(
        defaultHangingIndent !== undefined ? twipsToInches(defaultHangingIndent) : 0,
        defaultHangingIndent !== undefined ? "default" : INFERRED_SOURCE,
      ),
      alignment: valueWithSource(
        defaultAlignment !== undefined ? defaultAlignment : "left",
        defaultAlignment !== undefined ? "default" : INFERRED_SOURCE,
      ),
    },
  };
}

function findProperty(paragraphProperties, propertyName) {
  if (!paragraphProperties) {
    return undefined;
  }

  if (propertyName === "lineSpacing") {
    return paragraphProperties.spacing && paragraphProperties.spacing.line !== undefined
      ? {
          line: paragraphProperties.spacing.line,
          lineRule: paragraphProperties.spacing.lineRule || null,
        }
      : undefined;
  }

  if (propertyName === "spacingBefore") {
    return paragraphProperties.spacing && paragraphProperties.spacing.before !== undefined
      ? paragraphProperties.spacing.before
      : undefined;
  }

  if (propertyName === "spacingAfter") {
    return paragraphProperties.spacing && paragraphProperties.spacing.after !== undefined
      ? paragraphProperties.spacing.after
      : undefined;
  }

  if (propertyName === "firstLineIndent") {
    return paragraphProperties.ind && paragraphProperties.ind.firstLine !== undefined
      ? paragraphProperties.ind.firstLine
      : undefined;
  }

  if (propertyName === "hangingIndent") {
    return paragraphProperties.ind && paragraphProperties.ind.hanging !== undefined
      ? paragraphProperties.ind.hanging
      : undefined;
  }

  if (propertyName === "alignment") {
    return paragraphProperties.jc && paragraphProperties.jc.val !== undefined ? paragraphProperties.jc.val : undefined;
  }

  return undefined;
}

function resolveFromStyleChain(styleId, styleMap, propertyName) {
  let currentStyleId = styleId;
  let isParentStyle = false;
  const visited = new Set();

  while (currentStyleId && !visited.has(currentStyleId)) {
    visited.add(currentStyleId);
    const style = styleMap.paragraphStyles[currentStyleId];
    if (!style) {
      return null;
    }

    const rawValue = findProperty(style.properties, propertyName);
    if (rawValue !== undefined) {
      return {
        rawValue,
        source: isParentStyle ? "inherited" : "style",
      };
    }

    currentStyleId = style.basedOn;
    isParentStyle = true;
  }

  return null;
}

function resolveRawParagraphProperty(paragraphProperties, styleId, styleMap, defaults, propertyName) {
  const directValue = findProperty(paragraphProperties, propertyName);
  if (directValue !== undefined) {
    return {
      rawValue: directValue,
      source: "direct",
    };
  }

  const styleValue = resolveFromStyleChain(styleId || styleMap.defaultParagraphStyleId, styleMap, propertyName);
  if (styleValue) {
    return styleValue;
  }

  const defaultValue = findProperty(defaults.paragraphProperties, propertyName);
  if (defaultValue !== undefined) {
    return {
      rawValue: defaultValue,
      source: "default",
    };
  }

  if (propertyName === "alignment") {
    return {
      rawValue: "left",
      source: INFERRED_SOURCE,
    };
  }

  if (propertyName === "firstLineIndent" || propertyName === "hangingIndent") {
    return {
      rawValue: "0",
      source: INFERRED_SOURCE,
    };
  }

  return {
    rawValue: null,
    source: "unknown",
  };
}

function resolveParagraphValue(paragraphProperties, styleId, styleMap, defaults, propertyName, convert) {
  const resolved = resolveRawParagraphProperty(paragraphProperties, styleId, styleMap, defaults, propertyName);
  const value = resolved.rawValue === null ? null : convert(resolved.rawValue);

  return valueWithSource(value, resolved.source);
}

function resolveStyleInfo(styleId, styleMap) {
  const style = styleId ? styleMap.paragraphStyles[styleId] : null;

  return {
    id: styleId,
    name: style ? style.name : null,
    basedOn: style ? style.basedOn : null,
  };
}

function isHeading(styleInfo) {
  const styleId = (styleInfo.id || "").toLowerCase();
  const styleName = (styleInfo.name || "").toLowerCase();

  return styleId.startsWith("heading") || styleName.startsWith("heading");
}

function isTitleCaseWord(word) {
  if (!/[A-Za-z]/.test(word)) {
    return true;
  }

  return /^[A-Z][a-z0-9'’-]*$/.test(word) || /^[A-Z0-9]+$/.test(word);
}

function isLikelyHeadingText(text) {
  const trimmed = text.trim();
  if (!trimmed || trimmed.length > 80 || /[.!?;:]$/.test(trimmed)) {
    return false;
  }

  const words = trimmed.split(/\s+/);
  if (words.length > 10) {
    return false;
  }

  const titleCaseWords = words.filter(isTitleCaseWord).length;
  return titleCaseWords / words.length >= 0.8;
}

function isLikelyTitlePageParagraph(paragraph, nonBlankPosition) {
  const styleId = (paragraph.style.id || "").toLowerCase();
  const styleName = (paragraph.style.name || "").toLowerCase();

  if (styleId.includes("title") || styleName.includes("title")) {
    return true;
  }

  if (nonBlankPosition > 6) {
    return false;
  }

  const text = paragraph.text.trim();
  const alignment = paragraph.formatting.alignment.value;
  const firstLineIndent = paragraph.formatting.firstLineIndentInches.value;
  const hasBodyIndent = typeof firstLineIndent === "number" && firstLineIndent > 0;

  return text.length > 0 && text.length <= 140 && alignment === "center" && !hasBodyIndent;
}

function classifyParagraph(paragraph, nonBlankPosition) {
  if (paragraph.text.trim().length === 0) {
    return "blank";
  }

  if (isLikelyTitlePageParagraph(paragraph, nonBlankPosition)) {
    return "titlePage";
  }

  if (isHeading(paragraph.style) || isLikelyHeadingText(paragraph.text)) {
    return "heading";
  }

  return "body";
}

function extractParagraphFormatting(paragraph, index, styleMap, defaults, nonBlankPosition) {
  const paragraphProperties = paragraph.pPr || {};
  const explicitStyleId = getStyleId(paragraphProperties);
  const styleId = explicitStyleId || styleMap.defaultParagraphStyleId;
  const text = getParagraphText(paragraph);
  const style = resolveStyleInfo(styleId, styleMap);
  const extracted = {
    index: index + 1,
    text,
    textPreview: text.slice(0, 80),
    style,
    hasExplicitStyle: Boolean(explicitStyleId),
    role: "body",
    formatting: {
      lineSpacingMultiple: resolveParagraphValue(
        paragraphProperties,
        styleId,
        styleMap,
        defaults,
        "lineSpacing",
        (value) => lineValueToMultiple(value.line, value.lineRule),
      ),
      lineSpacingRule: resolveParagraphValue(
        paragraphProperties,
        styleId,
        styleMap,
        defaults,
        "lineSpacing",
        (value) => value.lineRule || null,
      ),
      spacingBeforePoints: resolveParagraphValue(
        paragraphProperties,
        styleId,
        styleMap,
        defaults,
        "spacingBefore",
        twipsToPoints,
      ),
      spacingAfterPoints: resolveParagraphValue(
        paragraphProperties,
        styleId,
        styleMap,
        defaults,
        "spacingAfter",
        twipsToPoints,
      ),
      firstLineIndentInches: resolveParagraphValue(
        paragraphProperties,
        styleId,
        styleMap,
        defaults,
        "firstLineIndent",
        twipsToInches,
      ),
      hangingIndentInches: resolveParagraphValue(
        paragraphProperties,
        styleId,
        styleMap,
        defaults,
        "hangingIndent",
        twipsToInches,
      ),
      alignment: resolveParagraphValue(paragraphProperties, styleId, styleMap, defaults, "alignment", (value) => value),
    },
  };

  extracted.role = classifyParagraph(extracted, nonBlankPosition);

  return extracted;
}

function extractParagraphs(document, styleMap, defaults) {
  const body = document.document && document.document.body;
  const paragraphs = body ? toArray(body.p) : [];
  let nonBlankPosition = 0;

  return paragraphs.map((paragraph, index) => {
    if (getParagraphText(paragraph).trim().length > 0) {
      nonBlankPosition += 1;
    }

    return extractParagraphFormatting(paragraph, index, styleMap, defaults, nonBlankPosition);
  });
}

function extractDocxFormattingFromXml(documentXml, stylesXml = null) {
  if (!documentXml) {
    throw new Error("Invalid .docx file: word/document.xml was not found.");
  }

  const document = parseXml(documentXml);
  const stylesDocument = stylesXml ? parseXml(stylesXml) : null;
  const styleMap = buildStyleMap(stylesDocument);
  const defaults = extractDocumentDefaults(stylesDocument);
  const sectionProperties = findSectionProperties(document);

  return {
    margins: extractMargins(sectionProperties),
    styles: {
      hasStylesXml: Boolean(stylesXml),
      defaultParagraphStyleId: styleMap.defaultParagraphStyleId,
    },
    documentDefaults: defaults.formatting,
    paragraphs: extractParagraphs(document, styleMap, defaults),
  };
}

function extractDocxFormatting(filePath) {
  const nodeRequire = module.require.bind(module);
  const fs = nodeRequire("fs");
  const AdmZip = nodeRequire("adm-zip");

  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const zip = new AdmZip(filePath);
  const documentXml = readDocxXml(zip, "word/document.xml");
  const stylesXml = readDocxXml(zip, "word/styles.xml");

  return extractDocxFormattingFromXml(documentXml, stylesXml);
}

module.exports = {
  extractDocxFormatting,
  extractDocxFormattingFromXml,
};

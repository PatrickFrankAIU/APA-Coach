// Session-scoped cache: avoids re-fetching the same DOI or URL within a session.
// Keyed as "doi:<doi>" or "url:<normalized-url>". Clears on page reload.
const cache = new Map();

const STOPWORDS = new Set(["the", "a", "an", "of", "in", "on", "for", "and", "or", "to", "with"]);
const DOI_PATTERN = /\b(10\.\d{4,}\/[^\s,)>"]+)/i;

function normalizeText(text) {
  return text.toLowerCase().replace(/[^\w\s]/g, " ").replace(/\s+/g, " ").trim();
}

function contentWords(text) {
  return normalizeText(text)
    .split(" ")
    .filter((w) => w.length > 1 && !STOPWORDS.has(w));
}

// Recall-biased overlap: measures what fraction of the CrossRef title's words appear in the
// reference text. Works well when the query (CrossRef title) is shorter than the document
// (full reference entry text).
function tokenSetRatio(query, document) {
  const qWords = new Set(contentWords(query));
  const dWords = new Set(contentWords(document));
  if (qWords.size === 0) return 0;
  let hits = 0;
  for (const w of qWords) {
    if (dWords.has(w)) hits++;
  }
  return hits / qWords.size;
}

function extractDoi(group) {
  const text = group.map((p) => p.text).join(" ");
  const match = text.match(DOI_PATTERN);
  return match ? match[1].replace(/[.,;:]+$/, "") : null;
}

function extractUrl(group) {
  const text = group.map((p) => p.text).join(" ");
  const urls = (text.match(/https?:\/\/[^\s,)<>"]+/gi) || []).map((u) =>
    u.replace(/[.,;]+$/, ""),
  );
  const nonDoi = urls.find((u) => !u.toLowerCase().includes("doi.org"));
  if (nonDoi) return nonDoi;
  for (const p of group) {
    if (p.hyperlinkUrls) {
      for (const url of p.hyperlinkUrls) {
        if (/^https?:\/\//i.test(url) && !url.toLowerCase().includes("doi.org")) {
          return url;
        }
      }
    }
  }
  return null;
}

function extractYear(group) {
  const match = group[0].text.match(/\((\d{4})[a-z]?\)/);
  return match ? match[1] : null;
}

function extractLastName(group) {
  const text = group[0].text;
  const yearMatch = text.match(/\((\d{4})[a-z]?\)/);
  if (!yearMatch) return null;
  const before = text.substring(0, yearMatch.index).trim();
  return before.split(",")[0].trim() || null;
}

function groupPreview(group) {
  const text = group[0].text;
  return text.length > 70 ? text.substring(0, 70) + "…" : text;
}

// Runs tasks with at most `limit` concurrent promises, preserving result order.
async function withConcurrency(tasks, limit) {
  const results = new Array(tasks.length);
  let next = 0;

  async function worker() {
    while (next < tasks.length) {
      const i = next++;
      results[i] = await tasks[i]();
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(limit, tasks.length) }, () => worker()),
  );
  return results;
}

async function fetchCrossRef(doi, signal, attempt = 1) {
  const response = await fetch(`https://api.crossref.org/works/${encodeURIComponent(doi)}`, {
    headers: { "User-Agent": "APA-Coach/1.0 (mailto:pfrank@aiuniv.edu)" },
    signal: AbortSignal.any([signal, AbortSignal.timeout(8000)]),
  });
  if (response.status === 429 && attempt === 1) {
    await new Promise((r) => setTimeout(r, 1000));
    return fetchCrossRef(doi, signal, 2);
  }
  return response;
}

async function verifySingleDoi(doi, refText, refYear, refLastName, signal) {
  const cacheKey = `doi:${doi}`;
  if (cache.has(cacheKey)) return cache.get(cacheKey);

  // Malformed DOI — skip CrossRef, treat as unresolved.
  if (!/^10\.\d{4,}\/.+/.test(doi)) {
    const result = { outcome: "unresolved" };
    cache.set(cacheKey, result);
    return result;
  }

  let response;
  try {
    response = await fetchCrossRef(doi, signal);
  } catch (err) {
    // Global deadline fired — don't cache so a re-upload can retry this DOI.
    if (signal?.aborted) return { outcome: "couldnt_verify" };
    // Per-request timeout or network error — real failure, worth caching.
    const result = { outcome: "unresolved" };
    cache.set(cacheKey, result);
    return result;
  }

  if (response.status === 404) {
    const result = { outcome: "unresolved" };
    cache.set(cacheKey, result);
    return result;
  }
  if (!response.ok) {
    // Covers 429 after retry and other server errors.
    const result = { outcome: "couldnt_verify" };
    cache.set(cacheKey, result);
    return result;
  }

  let data;
  try {
    data = await response.json();
  } catch {
    const result = { outcome: "couldnt_verify" };
    cache.set(cacheKey, result);
    return result;
  }

  const work = data.message;
  const crossrefTitle = (Array.isArray(work.title) ? work.title[0] : work.title) || "";
  const crossrefAuthorFamily = work.author?.[0]?.family || "";
  const crossrefYear = String(
    work["published-print"]?.["date-parts"]?.[0]?.[0] ??
      work["published-online"]?.["date-parts"]?.[0]?.[0] ??
      work["created"]?.["date-parts"]?.[0]?.[0] ??
      "",
  );

  const titleMatch = tokenSetRatio(crossrefTitle, refText) >= 0.75;

  // Allow substring match in either direction to handle name variations and transliterations.
  const normCrossref = crossrefAuthorFamily.toLowerCase().replace(/[^a-z]/g, "");
  const normRef = (refLastName || "").toLowerCase().replace(/[^a-z]/g, "");
  const authorMatch =
    !normCrossref ||
    !normRef ||
    normCrossref.includes(normRef) ||
    normRef.includes(normCrossref);

  const yearMatch = !crossrefYear || !refYear || crossrefYear === refYear;

  const result =
    titleMatch && authorMatch && yearMatch
      ? { outcome: "verified" }
      : { outcome: "mismatch", crossrefTitle };
  cache.set(cacheKey, result);
  return result;
}

async function verifySingleUrl(url, signal) {
  const cacheKey = `url:${url.toLowerCase().replace(/\/$/, "")}`;
  if (cache.has(cacheKey)) return cache.get(cacheKey);

  // Liveness-only: browser CORS policy prevents reading the response body or status code
  // for cross-origin URLs. mode:'no-cors' gives an opaque response when the host answers,
  // but opaque responses expose no status code — a 200 and a 404 are indistinguishable.
  // To extend this to full content verification (og:title, JSON-LD, body-text matching),
  // route fetches through a server-side proxy that returns the page metadata.
  let result;
  try {
    await fetch(url, {
      mode: "no-cors",
      signal: AbortSignal.any([signal, AbortSignal.timeout(6000)]),
    });
    // Opaque response: the host answered, but status code and content are inaccessible.
    result = { outcome: "couldnt_verify" };
  } catch (err) {
    if (signal?.aborted) {
      // Global deadline fired — don't cache so a re-upload can retry this URL.
      return { outcome: "couldnt_verify" };
    }
    // Per-request timeout (6 s) or network error — real failure, worth caching.
    result = { outcome: "unresolved" };
  }

  cache.set(cacheKey, result);
  return result;
}

export async function verifyReferenceLinks(referenceGroups) {
  if (!referenceGroups || referenceGroups.length === 0) return null;

  const controller = new AbortController();
  const { signal } = controller;
  // 13 s total budget: two full rounds of 5 concurrent 6 s URL fetches, plus a hair of overage.
  const deadlineTimer = setTimeout(() => controller.abort(), 13_000);

  const tasks = referenceGroups.map((group) => async () => {
    // If the deadline already fired before this task starts, skip the fetch entirely.
    if (signal.aborted) {
      return { preview: groupPreview(group), outcome: "couldnt_verify" };
    }

    const doi = extractDoi(group);
    const refText = group.map((p) => p.text).join(" ");
    const refYear = extractYear(group);
    const refLastName = extractLastName(group);

    if (doi) {
      const res = await verifySingleDoi(doi, refText, refYear, refLastName, signal);
      return { preview: groupPreview(group), doi, ...res };
    }

    const url = extractUrl(group);
    if (url) {
      const res = await verifySingleUrl(url, signal);
      return { preview: groupPreview(group), url, ...res };
    }

    return null; // Print source — no DOI or URL to verify.
  });

  let raw;
  try {
    raw = await withConcurrency(tasks, 5);
  } finally {
    clearTimeout(deadlineTimer);
  }
  const results = raw.filter(Boolean);

  if (results.length === 0) return null;

  const mismatches = results.filter((r) => r.outcome === "mismatch");
  const unresolved = results.filter((r) => r.outcome === "unresolved");
  const verified = results.filter((r) => r.outcome === "verified");
  // DOI couldnt_verify = CrossRef was unreachable or errored (unusual, worth flagging).
  // URL couldnt_verify = opaque no-cors response — the best the browser can confirm.
  //   Treating these as effectively passing avoids a permanent yellow card on every paper
  //   with URL-only references, since content verification requires a server-side proxy.
  const doiCouldntVerify = results.filter((r) => r.outcome === "couldnt_verify" && r.doi);
  const urlResponded = results.filter((r) => r.outcome === "couldnt_verify" && !r.doi);

  const status =
    mismatches.length > 0
      ? "warn"
      : unresolved.length > 0 || doiCouldntVerify.length > 0
        ? "review"
        : "pass";

  const issueCount = mismatches.length + unresolved.length + doiCouldntVerify.length;
  const total = results.length;

  let foundText;
  if (status === "pass") {
    if (verified.length === total) {
      foundText = `All ${total} reference link${total === 1 ? "" : "s"} verified via CrossRef.`;
    } else if (urlResponded.length > 0 && verified.length === 0) {
      foundText = `All ${total} reference URL${total === 1 ? "" : "s"} responded. Link content cannot be read from the browser — open each link to confirm it leads to the source you intended.`;
    } else {
      foundText = `${verified.length} DOI${verified.length === 1 ? "" : "s"} verified; ${urlResponded.length} URL${urlResponded.length === 1 ? "" : "s"} responded.`;
    }
  } else if (mismatches.length > 0 && unresolved.length === 0 && doiCouldntVerify.length === 0) {
    foundText = `${mismatches.length} of ${total} reference${mismatches.length === 1 ? "" : "s"} may point to the wrong source.`;
  } else {
    foundText = `${issueCount} of ${total} reference link${issueCount === 1 ? "" : "s"} could not be confirmed. Some sites block automated checks with paywalls or robot verification — open each flagged link to confirm it works and leads to the source you intended.`;
  }

  const missingItems = [
    ...mismatches.map(
      (r) =>
        `${r.preview} — DOI resolves to a different source than your reference describes; verify you have the correct DOI`,
    ),
    ...unresolved.map(
      (r) => `${r.preview} — link didn't load; check for typos or broken links`,
    ),
    ...doiCouldntVerify.map(
      (r) =>
        `${r.preview} — couldn't reach CrossRef to verify this DOI; check the doi.org link manually`,
    ),
  ];

  const howToFix = [];
  if (mismatches.length > 0) {
    howToFix.push(
      "For references with a mismatched DOI: re-look up the correct DOI using CrossRef (search.crossref.org) or your library database, then copy it directly from the article's landing page.",
    );
  }
  if (unresolved.length > 0) {
    howToFix.push(
      "For broken or dead links: open each URL in your browser. If the page has moved, find the current location or search for an archived version at web.archive.org.",
    );
  }
  if (doiCouldntVerify.length > 0) {
    howToFix.push(
      "For unverified DOIs: visit https://doi.org/<your-doi> directly in your browser and confirm it leads to the correct source.",
    );
  }

  return {
    rule: "Reference link verification",
    status,
    passed: status === "pass",
    expected:
      "Each DOI should resolve to the source it claims to cite. Each URL should load and lead to the cited page.",
    expectedText:
      "Each DOI should resolve to the source it claims to cite. Each URL should load and lead to the cited page.",
    foundText,
    applicable: total,
    checked: total,
    matched: verified.length + urlResponded.length,
    failed: mismatches.length,
    unknown: unresolved.length + doiCouldntVerify.length,
    found: status === "pass" ? "All links verified" : `${issueCount} link(s) need attention`,
    applicableParagraphs: referenceGroups.length,
    details: missingItems,
    howToFix,
    resources: [],
    missingItems: status !== "pass" ? missingItems : [],
    missingItemsLabel: "References with link issues:",
  };
}

import React, { useEffect, useId, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import apaCoachLogoUrl from "../../apa-coach-logo.svg";
import packageInfo from "../../package.json";
import { analyzeDocxFile } from "./apaBrowser.js";
import "./styles.css";

const APP_INFO = {
  version: packageInfo.version,
  lastUpdated: "May 23, 2026",
  supportEmail: "pfrank@aiuniv.edu",
  resources: [
    {
      label: "AIUS APA Guide",
      url: "https://careered.libguides.com/AIUS/APA",
    },
  ],
};

function useInstallPrompt() {
  const [prompt, setPrompt] = useState(null);

  useEffect(() => {
    function handler(e) {
      e.preventDefault();
      setPrompt(e);
    }
    window.addEventListener("beforeinstallprompt", handler);
    return () => window.removeEventListener("beforeinstallprompt", handler);
  }, []);

  function triggerInstall() {
    if (!prompt) return;
    prompt.prompt();
    prompt.userChoice.then((choice) => {
      if (choice.outcome === "accepted") {
        window.goatcounter?.count({ path: "pwa-installed", title: "PWA Installed", event: true });
      }
      setPrompt(null);
    });
  }

  return { canInstall: !!prompt, triggerInstall };
}

const STATUS_ORDER = {
  fail: 0,
  warn: 1,
  review: 2,
  pass: 3,
  skipped: 4,
};

const CHECK_CATEGORY = {
  // Paper formatting
  "Page numbering": "Paper formatting",
  "Title page": "Paper formatting",
  "Margins": "Paper formatting",
  "Body line spacing": "Paper formatting",
  "Heading line spacing": "Paper formatting",
  "Body paragraph spacing": "Paper formatting",
  "Heading paragraph spacing": "Paper formatting",
  "Body first-line indents": "Paper formatting",
  "Body alignment": "Paper formatting",
  "Font": "Paper formatting",
  "Unconverted markup symbols": "Paper formatting",
  // References
  "References page": "References",
  "References heading alignment": "References",
  "References line spacing": "References",
  "Reference hanging indent": "References",
  "Reference DOI/URL": "References",
  "Reference short link": "References",
  "Unapproved source": "References",
  "Reference authors": "References",
  "Reference year format": "References",
  "Reference title capitalization": "References",
  "Reference italics": "References",
  "Reference punctuation": "References",
  "Reference DOI format": "References",
  "Reference link verification": "References",
  // Citations
  "In-text citations": "Citations",
  "Uncited references": "Citations",
  "Unmatched citations": "Citations",
  "Personal communication": "Citations",
  "Citation ampersand": "Citations",
  "Citation et al. format": "Citations",
  "Citation comma": "Citations",
  "Citation no-date format": "Citations",
  "Citation page format": "Citations",
  "Citation multiple sources": "Citations",
  "Citation year suffix": "Citations",
  "Secondary citations": "Citations",
};

const CATEGORY_ORDER = ["Paper formatting", "References", "Citations"];

function sortChecks(checks) {
  return [...checks].sort((a, b) => STATUS_ORDER[a.status] - STATUS_ORDER[b.status]);
}

function categorySlug(label) {
  return label.toLowerCase().replace(/\s+/g, "-");
}

function getCategoryLinksForStatuses(checks, statuses) {
  const groups = new Map();
  for (const check of checks) {
    if (!statuses.includes(check.status)) continue;
    const cat = CHECK_CATEGORY[check.rule] ?? "Other";
    if (!groups.has(cat)) groups.set(cat, { count: 0, primaryStatus: check.status });
    const g = groups.get(cat);
    g.count++;
    if (STATUS_ORDER[check.status] < STATUS_ORDER[g.primaryStatus]) g.primaryStatus = check.status;
  }
  const links = [];
  for (const cat of [...CATEGORY_ORDER, "Other"]) {
    if (groups.has(cat)) {
      const { count, primaryStatus } = groups.get(cat);
      links.push({ label: cat, count, href: `#${primaryStatus}-${categorySlug(cat)}` });
    }
  }
  return links;
}

function CheckList({ checks, status }) {
  const categories = useMemo(() => {
    const groups = new Map();
    for (const check of checks) {
      const cat = CHECK_CATEGORY[check.rule] ?? "Other";
      if (!groups.has(cat)) groups.set(cat, []);
      groups.get(cat).push(check);
    }
    const ordered = [];
    for (const cat of [...CATEGORY_ORDER, "Other"]) {
      if (groups.has(cat)) ordered.push({ label: cat, checks: groups.get(cat) });
    }
    return ordered;
  }, [checks]);

  return (
    <>
      {categories.map(({ label, checks: catChecks }) => (
        <React.Fragment key={label}>
          <div id={`${status}-${categorySlug(label)}`} className="check-category-label">
            <span>{label}</span>
            <a href="#top" className="check-category-top">Back to top</a>
          </div>
          {catChecks.map((check) => (
            <CheckCard key={check.rule} check={check} />
          ))}
        </React.Fragment>
      ))}
    </>
  );
}

function SummaryCell({ label, links, colorClass }) {
  return (
    <div className={`summary-cell summary-cell--${colorClass}`}>
      <dt>{label}</dt>
      <dd>
        {links.length === 0 ? (
          <span className="summary-cell-zero">—</span>
        ) : (
          <ul className="summary-cat-links" role="list">
            {links.map(({ label: catLabel, count, href }) => (
              <li key={catLabel}>
                <a href={href} className="summary-cat-link">
                  <span className="summary-cat-name">{catLabel}</span>
                  <span className="summary-cat-count">{count}</span>
                </a>
              </li>
            ))}
          </ul>
        )}
      </dd>
    </div>
  );
}

function Summary({ report, checks }) {
  const { failed, warn = 0, review } = report.summary;
  const issueCount = failed + warn;
  const isReady = issueCount === 0;
  const issueText = issueCount === 1 ? "issue" : "issues";

  const failLinks = useMemo(() => getCategoryLinksForStatuses(checks, ["fail", "warn"]), [checks]);
  const reviewLinks = useMemo(() => getCategoryLinksForStatuses(checks, ["review"]), [checks]);
  const passLinks = useMemo(() => getCategoryLinksForStatuses(checks, ["pass"]), [checks]);

  return (
    <section className="summary" aria-labelledby="summary-heading">
      <div>
        <p className="eyebrow">Report summary</p>
        <h2 id="summary-heading">
          {isReady ? "You're ready to submit" : `${issueCount} ${issueText} to fix`}
        </h2>
        <p className="summary-copy">
          {isReady && review > 0
            ? `${review} optional ${review === 1 ? "item" : "items"} to review.`
            : `APA Coach checked ${report.summary.totalChecks} items in ${report.file}.`}
        </p>
        <button
          className="print-report-btn"
          type="button"
          onClick={() => {
            const stem = report.file.replace(/\.docx$/i, "");
            const orig = document.title;
            document.title = `APACoach-${stem}`;
            window.addEventListener("afterprint", () => { document.title = orig; }, { once: true });
            window.print();
          }}
        >
          Print / Save as PDF
        </button>
      </div>
      <dl className="summary-grid" aria-label="Check totals">
        <SummaryCell label="Failed" links={failLinks} colorClass="fail" />
        <SummaryCell label="Review" links={reviewLinks} colorClass="review" />
        <SummaryCell label="Passed" links={passLinks} colorClass="pass" />
      </dl>
    </section>
  );
}

const ITEMS_PREVIEW_COUNT = 5;

function CheckCard({ check }) {
  const fixId = useId();
  const itemsId = useId();
  const [open, setOpen] = useState(false);
  const [itemsOpen, setItemsOpen] = useState(false);
  const hasFixes = check.howToFix && check.howToFix.length > 0;
  const hasResources = check.resources && check.resources.length > 0;
  const allItems = check.missingItems || [];
  const hasMoreItems = allItems.length > ITEMS_PREVIEW_COUNT;
  const visibleItems = hasMoreItems && !itemsOpen ? allItems.slice(0, ITEMS_PREVIEW_COUNT) : allItems;

  return (
    <article className={`check-card ${check.status}`}>
      <div className="check-header">
        <span className="status-badge">{check.status}</span>
        <h3>{check.rule}</h3>
      </div>
      <p className="found">{check.foundText || check.found}</p>
      <p className="expected">{check.expectedText || check.expected}</p>
      {check.expectedItems && check.expectedItems.length > 0 ? (
        <ul className="expected-items">
          {check.expectedItems.map((item, i) => {
            const text = typeof item === "string" ? item : item.text;
            const sub = typeof item === "object" && item.sub;
            const parts = text.split(/\*([^*]+)\*/);
            const content = parts.map((part, j) => j % 2 === 1 ? <em key={j}>{part}</em> : part);
            return <li key={i} className={sub ? "expected-item-sub" : undefined}>{content}</li>;
          })}
        </ul>
      ) : null}
      {hasFixes ? (
        <div className="fixes">
          {check.rule === "Page numbering" ? (
            <div className="page-numbering-example">
              <p className="page-numbering-example-label">Page number must appear in the upper right corner of every page:</p>
              <div className="page-numbering-demo">
                <div className="page-numbering-col">
                  <span className="page-numbering-tag page-numbering-tag--wrong">✗ Missing</span>
                  <div className="page-numbering-page">
                    <div className="page-numbering-header page-numbering-header--empty" />
                    <div className="page-numbering-body">
                      <div className="page-numbering-line" />
                      <div className="page-numbering-line" />
                      <div className="page-numbering-line page-numbering-line--short" />
                    </div>
                  </div>
                </div>
                <div className="page-numbering-col">
                  <span className="page-numbering-tag page-numbering-tag--right">✓ Correct</span>
                  <div className="page-numbering-page">
                    <div className="page-numbering-header">
                      <span className="page-numbering-number">1</span>
                    </div>
                    <div className="page-numbering-body">
                      <div className="page-numbering-line" />
                      <div className="page-numbering-line" />
                      <div className="page-numbering-line page-numbering-line--short" />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Title page" ? (
            <div className="title-page-example">
              <p className="title-page-example-label">Example APA title page layout:</p>
              <div className="title-page-demo">
                <p className="title-page-title">The Full Title of Your Paper</p>
                <div className="title-page-spacer" />
                <p className="title-page-meta">Author Name</p>
                <p className="title-page-meta">Department, University Name</p>
                <p className="title-page-meta">Course Number: Course Name</p>
                <p className="title-page-meta">Instructor Name</p>
                <p className="title-page-meta">Month Day, Year</p>
              </div>
            </div>
          ) : null}
          {check.rule === "Margins" ? (
            <div className="margins-example">
              <p className="margins-example-label">All four margins must be 1 inch:</p>
              <div className="margins-demo">
                <div className="margins-col">
                  <span className="margins-tag margins-tag--wrong">✗ Wrong</span>
                  <div className="margins-page">
                    <div className="margins-content margins-content--tight">
                      <div className="margins-line" /><div className="margins-line" /><div className="margins-line margins-line--short" />
                    </div>
                  </div>
                  <span className="margins-caption">Too narrow</span>
                </div>
                <div className="margins-col">
                  <span className="margins-tag margins-tag--wrong">✗ Wrong</span>
                  <div className="margins-page">
                    <div className="margins-content margins-content--wide">
                      <div className="margins-line" /><div className="margins-line" /><div className="margins-line margins-line--short" />
                    </div>
                  </div>
                  <span className="margins-caption">Too wide</span>
                </div>
                <div className="margins-col">
                  <span className="margins-tag margins-tag--right">✓ Correct</span>
                  <div className="margins-page">
                    <div className="margins-content margins-content--correct">
                      <div className="margins-line" /><div className="margins-line" /><div className="margins-line margins-line--short" />
                    </div>
                  </div>
                  <span className="margins-caption">1 inch on all sides</span>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Heading paragraph spacing" ? (
            <div className="para-spacing-example">
              <p className="para-spacing-example-label">Headings must have 0 pt spacing before and after:</p>
              <div className="para-spacing-demo">
                <div className="para-spacing-col">
                  <span className="para-spacing-tag para-spacing-tag--wrong">✗ Extra spacing</span>
                  <div className="para-spacing-page">
                    <p className="line-spacing-heading-para line-spacing-entry--double para-spacing-para--extra">Method</p>
                    <p className="line-spacing-body-para line-spacing-entry--double para-spacing-para--extra">Lorem ipsum dolor sit amet, consectetur adipiscing.</p>
                    <p className="line-spacing-heading-para line-spacing-entry--double para-spacing-para--extra">Results</p>
                    <p className="line-spacing-body-para line-spacing-entry--double para-spacing-para--extra">Sed do eiusmod tempor incididunt ut labore.</p>
                  </div>
                </div>
                <div className="para-spacing-col">
                  <span className="para-spacing-tag para-spacing-tag--right">✓ No extra spacing</span>
                  <div className="para-spacing-page">
                    <p className="line-spacing-heading-para line-spacing-entry--double para-spacing-para--none">Method</p>
                    <p className="line-spacing-body-para line-spacing-entry--double para-spacing-para--none">Lorem ipsum dolor sit amet, consectetur adipiscing.</p>
                    <p className="line-spacing-heading-para line-spacing-entry--double para-spacing-para--none">Results</p>
                    <p className="line-spacing-body-para line-spacing-entry--double para-spacing-para--none">Sed do eiusmod tempor incididunt ut labore.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Body paragraph spacing" ? (
            <div className="para-spacing-example">
              <p className="para-spacing-example-label">Paragraphs must have 0 pt spacing before and after:</p>
              <div className="para-spacing-demo">
                <div className="para-spacing-col">
                  <span className="para-spacing-tag para-spacing-tag--wrong">✗ Extra spacing</span>
                  <div className="para-spacing-page">
                    <p className="para-spacing-para para-spacing-para--extra">Lorem ipsum dolor sit amet, consectetur adipiscing elit ut labore.</p>
                    <p className="para-spacing-para para-spacing-para--extra">Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
                    <p className="para-spacing-para para-spacing-para--extra">Ut enim ad minim veniam quis nostrud exercitation ullamco.</p>
                  </div>
                </div>
                <div className="para-spacing-col">
                  <span className="para-spacing-tag para-spacing-tag--right">✓ No extra spacing</span>
                  <div className="para-spacing-page">
                    <p className="para-spacing-para para-spacing-para--none">Lorem ipsum dolor sit amet, consectetur adipiscing elit ut labore.</p>
                    <p className="para-spacing-para para-spacing-para--none">Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
                    <p className="para-spacing-para para-spacing-para--none">Ut enim ad minim veniam quis nostrud exercitation ullamco.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Body line spacing" ? (
            <div className="line-spacing-example">
              <p className="line-spacing-example-label">Body text must be double-spaced:</p>
              <div className="line-spacing-demo">
                <div className="line-spacing-col">
                  <span className="line-spacing-tag line-spacing-tag--wrong">✗ Single spacing</span>
                  <div className="line-spacing-page">
                    <p className="line-spacing-body-para line-spacing-entry--single">Lorem ipsum dolor sit amet, consectetur adipiscing elit.</p>
                    <p className="line-spacing-body-para line-spacing-entry--single">Sed do eiusmod tempor incididunt ut labore et dolore.</p>
                  </div>
                </div>
                <div className="line-spacing-col">
                  <span className="line-spacing-tag line-spacing-tag--right">✓ Double spacing</span>
                  <div className="line-spacing-page">
                    <p className="line-spacing-body-para line-spacing-entry--double">Lorem ipsum dolor sit amet, consectetur adipiscing elit.</p>
                    <p className="line-spacing-body-para line-spacing-entry--double">Sed do eiusmod tempor incididunt ut labore et dolore.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Heading line spacing" ? (
            <div className="line-spacing-example">
              <p className="line-spacing-example-label">Headings must be double-spaced:</p>
              <div className="line-spacing-demo">
                <div className="line-spacing-col">
                  <span className="line-spacing-tag line-spacing-tag--wrong">✗ Single spacing</span>
                  <div className="line-spacing-page">
                    <p className="line-spacing-heading-para line-spacing-entry--single">Method</p>
                    <p className="line-spacing-body-para line-spacing-entry--single">Lorem ipsum dolor sit amet, consectetur adipiscing.</p>
                    <p className="line-spacing-heading-para line-spacing-entry--single">Results</p>
                    <p className="line-spacing-body-para line-spacing-entry--single">Sed do eiusmod tempor incididunt ut labore.</p>
                  </div>
                </div>
                <div className="line-spacing-col">
                  <span className="line-spacing-tag line-spacing-tag--right">✓ Double spacing</span>
                  <div className="line-spacing-page">
                    <p className="line-spacing-heading-para line-spacing-entry--double">Method</p>
                    <p className="line-spacing-body-para line-spacing-entry--double">Lorem ipsum dolor sit amet, consectetur adipiscing.</p>
                    <p className="line-spacing-heading-para line-spacing-entry--double">Results</p>
                    <p className="line-spacing-body-para line-spacing-entry--double">Sed do eiusmod tempor incididunt ut labore.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "References line spacing" ? (
            <div className="line-spacing-example">
              <p className="line-spacing-example-label">References must be double-spaced:</p>
              <div className="line-spacing-demo">
                <div className="line-spacing-col">
                  <span className="line-spacing-tag line-spacing-tag--wrong">✗ Single spacing</span>
                  <div className="line-spacing-page">
                    <p className="line-spacing-entry line-spacing-entry--single">Smith, J. A. (2023). <em>Book title here.</em> Publisher.</p>
                    <p className="line-spacing-entry line-spacing-entry--single">Jones, B. C. (2022). Article title. <em>Journal Name, 10</em>(2), 45–67.</p>
                  </div>
                </div>
                <div className="line-spacing-col">
                  <span className="line-spacing-tag line-spacing-tag--right">✓ Double spacing</span>
                  <div className="line-spacing-page">
                    <p className="line-spacing-entry line-spacing-entry--double">Smith, J. A. (2023). <em>Book title here.</em> Publisher.</p>
                    <p className="line-spacing-entry line-spacing-entry--double">Jones, B. C. (2022). Article title. <em>Journal Name, 10</em>(2), 45–67.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Body first-line indents" ? (
            <div className="first-line-indent-example">
              <p className="first-line-indent-example-label">Each paragraph must begin with a 0.5" indent:</p>
              <div className="first-line-indent-demo">
                <div className="first-line-indent-col">
                  <span className="first-line-indent-tag first-line-indent-tag--wrong">✗ Wrong</span>
                  <div className="first-line-indent-page">
                    <p className="first-line-indent-para first-line-indent-para--none">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor.</p>
                    <p className="first-line-indent-para first-line-indent-para--none">Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi.</p>
                  </div>
                </div>
                <div className="first-line-indent-col">
                  <span className="first-line-indent-tag first-line-indent-tag--right">✓ Correct</span>
                  <div className="first-line-indent-page">
                    <p className="first-line-indent-para first-line-indent-para--indented">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor.</p>
                    <p className="first-line-indent-para first-line-indent-para--indented">Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Reference hanging indent" ? (
            <div className="hanging-indent-example">
              <p className="hanging-indent-example-label">Example:</p>
              <div className="hanging-indent-demo">
                <p className="hanging-indent-entry">
                  Smith, J. A. (2023). <em>The title of a really long book that wraps onto a second line to show the indent effect.</em> Publisher Name.
                </p>
                <p className="hanging-indent-entry">
                  Jones, B. C., &amp; Lee, D. E. (2021). Another reference entry that also wraps to demonstrate the spacing between entries. Publisher.
                </p>
              </div>
            </div>
          ) : null}
          {check.rule === "Reference title capitalization" ? (
            <div className="citation-example">
              <p className="citation-example-label">Reference titles use sentence case — only the first word, the first word after a colon, and proper nouns are capitalized:</p>
              <div className="citation-example-demo">
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--wrong">✗ Wrong (Title Case)</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text">Smith, J. A. (2023). <em>The Effects of Social Media on Student Academic Performance.</em> Publisher.</p>
                  </div>
                </div>
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--right">✓ Correct (sentence case)</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text">Smith, J. A. (2023). <em>The effects of social media on student academic performance.</em> Publisher.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Font" ? (
            <div className="citation-example">
              <p className="citation-example-label">Use 12-point Times New Roman consistently throughout:</p>
              <div className="citation-example-demo">
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--wrong">✗ Wrong</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text font-example--wrong-heading">Results</p>
                    <p className="citation-body-text font-example--wrong-body">The experiment showed significant effects on participant response times across all conditions tested.</p>
                  </div>
                </div>
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--right">✓ Correct</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text font-example--right-heading">Results</p>
                    <p className="citation-body-text font-example--right-body">The experiment showed significant effects on participant response times across all conditions tested.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "In-text citations" ? (
            <div className="citation-example">
              <p className="citation-example-label">Cite every source you use in the body of your paper:</p>
              <div className="citation-example-demo">
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--wrong">✗ No citation</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text citation-body-text--highlighted">Research shows that sleep deprivation significantly impairs cognitive performance.</p>
                  </div>
                </div>
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--right">✓ Citation added</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text">Research shows that sleep deprivation significantly impairs cognitive performance <span className="citation-inline">(Walker, 2017)</span>.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Uncited references" ? (
            <div className="citation-example">
              <p className="citation-example-label">Each reference needs a matching in-text citation:</p>
              <div className="citation-example-demo">
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--wrong">✗ Missing citation</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text">The experiment showed significant effects on memory retention and recall ability.</p>
                    <div className="citation-divider" />
                    <p className="citation-ref-entry citation-ref-entry--highlighted">Smith, J. A. (2023). <em>Memory and cognition.</em> Publisher.</p>
                  </div>
                </div>
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--right">✓ Citation present</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text">The experiment showed significant effects on memory retention and recall ability <span className="citation-inline">(Smith, 2023)</span>.</p>
                    <div className="citation-divider" />
                    <p className="citation-ref-entry">Smith, J. A. (2023). <em>Memory and cognition.</em> Publisher.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Unmatched citations" ? (
            <div className="citation-example">
              <p className="citation-example-label">Each in-text citation needs a matching reference entry:</p>
              <div className="citation-example-demo">
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--wrong">✗ No reference</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text">Research suggests this is common <span className="citation-inline citation-inline--highlighted">(Jones, 2022)</span>.</p>
                    <div className="citation-divider" />
                    <p className="citation-ref-entry">Smith, J. A. (2023). <em>Memory and cognition.</em> Publisher.</p>
                    <p className="citation-ref-missing">Jones (2022) — not found</p>
                  </div>
                </div>
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--right">✓ Reference present</span>
                  <div className="citation-example-page">
                    <p className="citation-body-text">Research suggests this is common <span className="citation-inline">(Jones, 2022)</span>.</p>
                    <div className="citation-divider" />
                    <p className="citation-ref-entry">Jones, B. C. (2022). Article title. <em>Journal, 5</em>(1), 10–20.</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Reference short link" ? (
            <div className="citation-example">
              <p className="citation-example-label">The URL must link to the specific article, not the website homepage:</p>
              <div className="citation-example-demo">
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--wrong">✗ Domain only</span>
                  <div className="citation-example-page">
                    <p className="citation-ref-entry">Intel. (2023). <em>Thermal design in modern processors.</em> <span className="citation-ref-entry--highlighted">https://www.intel.com</span></p>
                  </div>
                </div>
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--right">✓ Full article URL</span>
                  <div className="citation-example-page">
                    <p className="citation-ref-entry">Intel. (2023). <em>Thermal design in modern processors.</em> https://www.intel.com/gaming/cpu-cooler.html</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "Reference DOI/URL" ? (
            <div className="citation-example">
              <p className="citation-example-label">Each reference should include a DOI or URL:</p>
              <div className="citation-example-demo">
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--wrong">✗ No DOI/URL</span>
                  <div className="citation-example-page">
                    <p className="citation-ref-entry citation-ref-entry--highlighted">Smith, J. A. (2023). <em>Book title.</em> Publisher.</p>
                  </div>
                </div>
                <div className="citation-example-col">
                  <span className="citation-example-tag citation-example-tag--right">✓ DOI included</span>
                  <div className="citation-example-page">
                    <p className="citation-ref-entry">Smith, J. A. (2023). <em>Book title.</em> Publisher. https://doi.org/10.xxxx/xxxxx</p>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {check.rule === "References heading alignment" ? (
            <div className="heading-alignment-example">
              <p className="heading-alignment-example-label">The heading must be centered:</p>
              <div className="heading-alignment-demo">
                <div className="heading-alignment-col">
                  <span className="heading-alignment-tag heading-alignment-tag--wrong">✗ Wrong</span>
                  <div className="heading-alignment-page">
                    <p className="heading-alignment-heading heading-alignment-heading--left">References</p>
                    <div className="heading-alignment-lines">
                      <span /><span /><span className="heading-alignment-line--short" />
                    </div>
                  </div>
                </div>
                <div className="heading-alignment-col">
                  <span className="heading-alignment-tag heading-alignment-tag--right">✓ Correct</span>
                  <div className="heading-alignment-page">
                    <p className="heading-alignment-heading heading-alignment-heading--center">References</p>
                    <div className="heading-alignment-lines">
                      <span /><span /><span className="heading-alignment-line--short" />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ) : null}
          {allItems.length > 0 ? (
            <div className="missing-items">
              <p className="missing-items-label">{check.missingItemsLabel}</p>
              <ul>
                {visibleItems.map((item, i) => (
                  <li key={i}>{item}</li>
                ))}
              </ul>
              {hasMoreItems ? (
                <button
                  className="fix-toggle"
                  type="button"
                  aria-expanded={itemsOpen}
                  aria-controls={itemsId}
                  onClick={() => setItemsOpen((v) => !v)}
                >
                  {itemsOpen ? "Show fewer" : `Show all ${allItems.length}`}
                </button>
              ) : null}
            </div>
          ) : null}
          <button
            className="fix-toggle"
            type="button"
            aria-expanded={open}
            aria-controls={fixId}
            onClick={() => setOpen((value) => !value)}
          >
            {open ? "Hide how to fix" : "Show how to fix"}
          </button>
          <div id={fixId} hidden={!open}>
            <ol>
              {check.howToFix.filter((s) => !s.startsWith("Note:")).map((step) => (
                <li key={step}>{step}</li>
              ))}
            </ol>
            {check.howToFix.filter((s) => s.startsWith("Note:")).map((note) => (
              <p key={note} className="fix-note"><em>{note}</em></p>
            ))}
            {hasResources ? (
              <div className="check-resources">
                <h4>Additional help</h4>
                <ul>
                  {check.resources.map((resource) => (
                    <li key={resource.url}>
                      <a href={resource.url} target="_blank" rel="noreferrer">
                        {resource.label}
                      </a>
                    </li>
                  ))}
                </ul>
              </div>
            ) : null}
          </div>
        </div>
      ) : null}
    </article>
  );
}

function Report({ report }) {
  const sortedChecks = useMemo(() => sortChecks(report.checks), [report]);
  const failChecks = sortedChecks.filter((c) => c.status === "fail");
  const warnChecks = sortedChecks.filter((c) => c.status === "warn");
  const reviewChecks = sortedChecks.filter((c) => c.status === "review");
  const passChecks = sortedChecks.filter((c) => c.status === "pass");
  const skippedChecks = sortedChecks.filter((c) => c.status === "skipped");

  const printDate = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });

  return (
    <main className="report" aria-live="polite">
      <div className="print-report-header" aria-hidden="true">
        <p className="print-report-title">APA Formatting Report</p>
        <p className="print-report-meta">{report.file} &mdash; Checked {printDate} &mdash; APA Coach v{APP_INFO.version}</p>
      </div>
      <Summary report={report} checks={report.checks} />
      <section className="checks" aria-label="APA checks">
        {failChecks.length > 0 && (
          <section id="fail-section" className="check-group" aria-labelledby="fail-heading">
            <h3 id="fail-heading" className="check-group-heading">Failed</h3>
            <p className="fail-intro">These items should be corrected before you submit your paper.</p>
            <CheckList checks={failChecks} status="fail" />
          </section>
        )}
        {warnChecks.length > 0 && (
          <section id="warn-section" className="check-group" aria-labelledby="warn-heading">
            <h3 id="warn-heading" className="check-group-heading">Warnings</h3>
            <p className="warn-intro">These items may indicate a problem. Open each link and confirm it leads to the source you intended.</p>
            <CheckList checks={warnChecks} status="warn" />
          </section>
        )}
        {reviewChecks.length > 0 && (
          <section id="review-section" className="check-group" aria-labelledby="review-heading">
            <h3 id="review-heading" className="check-group-heading">Review</h3>
            <p className="review-intro">
              These items may not require changes, but are worth double-checking.
            </p>
            <CheckList checks={reviewChecks} status="review" />
          </section>
        )}
        {passChecks.length > 0 && (
          <section id="pass-section" className="check-group" aria-labelledby="pass-heading">
            <h3 id="pass-heading" className="check-group-heading">Passed</h3>
            <p className="pass-intro">These items meet APA expectations.</p>
            <CheckList checks={passChecks} status="pass" />
            <a href="#top" className="back-to-top">↑ Back to top</a>
          </section>
        )}
        {skippedChecks.length > 0 && (
          <section id="skipped-section" className="check-group" aria-labelledby="skipped-heading">
            <h3 id="skipped-heading" className="check-group-heading">Not checked (Offline)</h3>
            <CheckList checks={skippedChecks} status="skipped" />
            <a href="#top" className="back-to-top">↑ Back to top</a>
          </section>
        )}
      </section>
    </main>
  );
}

function DocxDropZone({ fileName, isAnalyzing, onFileSelected, compact }) {
  const inputRef = useRef(null);
  const [isDragging, setIsDragging] = useState(false);

  function openFilePicker() {
    inputRef.current?.click();
  }

  function handleInputChange(event) {
    const file = event.target.files && event.target.files[0];
    onFileSelected(file);
    event.target.value = "";
  }

  function handleDragOver(event) {
    event.preventDefault();
    setIsDragging(true);
  }

  function handleDragLeave(event) {
    if (!event.currentTarget.contains(event.relatedTarget)) {
      setIsDragging(false);
    }
  }

  function handleDrop(event) {
    event.preventDefault();
    setIsDragging(false);
    onFileSelected(event.dataTransfer.files && event.dataTransfer.files[0]);
  }

  function handleKeyDown(event) {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      openFilePicker();
    }
  }

  if (compact) {
    return (
      <div
        className={`drop-zone-compact${isDragging ? " dragging" : ""}`}
        role="button"
        tabIndex="0"
        aria-label="Upload a .docx file"
        onClick={openFilePicker}
        onKeyDown={handleKeyDown}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <input
          ref={inputRef}
          className="file-input"
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          onChange={handleInputChange}
        />
        <p className="drop-zone-compact-hint">Drop a .docx file here, or</p>
        <button
          className="choose-file-button"
          type="button"
          disabled={isAnalyzing}
          onClick={(event) => {
            event.stopPropagation();
            openFilePicker();
          }}
        >
          Check another paper
        </button>
      </div>
    );
  }

  return (
    <div
      className={`drop-zone${isDragging ? " dragging" : ""}`}
      role="button"
      tabIndex="0"
      aria-label="Upload a .docx file"
      onClick={openFilePicker}
      onKeyDown={handleKeyDown}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <input
        ref={inputRef}
        className="file-input"
        type="file"
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        onChange={handleInputChange}
      />
      <p className="drop-title">Upload .docx file</p>
      <p className="drop-copy">
        {fileName || "Drag and drop your .docx file here, or use the button below."}
      </p>
      <button
        className="choose-file-button"
        type="button"
        disabled={isAnalyzing}
        onClick={(event) => {
          event.stopPropagation();
          openFilePicker();
        }}
      >
        Choose .docx file
      </button>
    </div>
  );
}

function AppInfoCard({ canInstall, triggerInstall }) {
  return (
    <aside className="app-info-card" aria-label="Application information">
      <div className="beta-notice" role="note">
        APA Coach is in active development. Checks may miss issues or flag things incorrectly. Always review your paper manually before submitting.
      </div>
      <dl className="app-info-list">
        <div>
          <dt>Version</dt>
          <dd>{APP_INFO.version}</dd>
        </div>
        <div>
          <dt>Last updated</dt>
          <dd>{APP_INFO.lastUpdated}</dd>
        </div>
      </dl>
      <p>
        This project is actively maintained on{" "}
        <a href="https://github.com/PatrickFrankAIU/APA-Coach" target="_blank" rel="noreferrer">GitHub</a>.
      </p>
      <p>
        Please report any problems to{" "}
        <a href={`mailto:${APP_INFO.supportEmail}`}>{APP_INFO.supportEmail}</a>
      </p>
      <p><em>If you report a problem, please attach a copy of your paper so I can see the issue myself. Thanks!</em></p>
      <div className="app-resources">
        <h2>Useful resources</h2>
        <ul>
          {APP_INFO.resources.map((resource) => (
            <li key={resource.url}>
              <a href={resource.url} target="_blank" rel="noreferrer">
                {resource.label}
              </a>
            </li>
          ))}
        </ul>
      </div>
      {canInstall && (
        <div className="install-section">
          <h2>Install APA Coach</h2>
          <p>Add APA Coach to your device for quick access. Formatting checks work offline; reference link verification requires a connection to the internet.</p>
          <button className="install-button" type="button" onClick={triggerInstall}>
            Install App
          </button>
        </div>
      )}
    </aside>
  );
}

const SESSION_KEY = "apa-coach-session";

function App() {
  const [report, setReport] = useState(null);
  const [error, setError] = useState("");
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analyzePhase, setAnalyzePhase] = useState("Analyzing document…");
  const [fileName, setFileName] = useState("");
  const analysisToken = useRef(0);
  const { canInstall, triggerInstall } = useInstallPrompt();

  useEffect(() => {
    try {
      const saved = sessionStorage.getItem(SESSION_KEY);
      if (saved) {
        const { report: savedReport, fileName: savedFileName } = JSON.parse(saved);
        setReport(savedReport);
        setFileName(savedFileName);
      }
    } catch {
      sessionStorage.removeItem(SESSION_KEY);
    }
  }, []);

  useEffect(() => {
    if (report) {
      sessionStorage.setItem(SESSION_KEY, JSON.stringify({ report, fileName }));
    } else {
      sessionStorage.removeItem(SESSION_KEY);
    }
  }, [report, fileName]);

  async function analyzeSelectedFile(file) {
    const token = ++analysisToken.current;
    setReport(null);
    setError("");

    if (!file) {
      setIsAnalyzing(false);
      return;
    }

    if (!file.name.toLowerCase().endsWith(".docx")) {
      setFileName("");
      setError("APA Coach can only analyze Word documents saved as .docx files.");
      setIsAnalyzing(false);
      return;
    }

    setFileName(file.name);
    try {
      setIsAnalyzing(true);
      setAnalyzePhase("Analyzing document…");
      const result = await analyzeDocxFile(file, (phase) => setAnalyzePhase(phase));
      if (token !== analysisToken.current) return;
      setReport(result);
      window.goatcounter?.count({ path: "paper-checked", title: "Paper Checked", event: true });
    } catch (analysisError) {
      if (token !== analysisToken.current) return;
      setError(analysisError.message || "APA Coach could not read this .docx file.");
    } finally {
      if (token === analysisToken.current) setIsAnalyzing(false);
    }
  }

  const isCompact = !!(report || isAnalyzing);

  return (
    <div id="top" className="app-shell">
      {isCompact ? (
        <header className="app-header-compact">
          <img className="brand-logo" src={apaCoachLogoUrl} alt="APA Coach" />
          <div className="compact-file-info">
            <p className="compact-file-label">{isAnalyzing ? analyzePhase : "Analyzing"}</p>
            <p className="compact-file-name">{fileName}</p>
          </div>
          <DocxDropZone fileName={fileName} isAnalyzing={isAnalyzing} onFileSelected={analyzeSelectedFile} compact />
        </header>
      ) : (
        <header className="app-header">
          <div>
            <div className="brand-block">
              <img className="brand-logo" src={apaCoachLogoUrl} alt="APA Coach" />
            </div>
            <h1>Check APA 7 Formatting</h1>
            <p>
              Submit a Word document to verify its compatibility with APA 7. Files are not uploaded or stored — your paper stays in your browser. DOIs and URLs in your references are checked against public databases to verify they resolve correctly. Formatting checks work offline; reference link verification requires a connection to the internet.
            </p>
          </div>
          <DocxDropZone fileName={fileName} isAnalyzing={isAnalyzing} onFileSelected={analyzeSelectedFile} />
        </header>
      )}

      {!report ? <AppInfoCard canInstall={canInstall} triggerInstall={triggerInstall} /> : null}
      {isAnalyzing ? <p className="notice" role="status" aria-live="polite">{analyzePhase}</p> : null}
      {error ? (
        <p className="error" role="alert">
          {error}
        </p>
      ) : null}
      {report ? <Report report={report} /> : null}
      {report ? <AppInfoCard canInstall={canInstall} triggerInstall={triggerInstall} /> : null}
      <footer className="app-footer">
        <p>APA Coach &copy; 2026 Patrick Frank</p>
        <p>GPL-3.0 Licensed</p>
      </footer>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);

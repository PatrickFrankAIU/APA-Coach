import React, { useId, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import aiuLogoUrl from "../../aiuslogo.png";
import packageInfo from "../../package.json";
import { analyzeDocxFile } from "./apaBrowser.js";
import "./styles.css";

const APP_INFO = {
  version: packageInfo.version,
  lastUpdated: "May 2, 2026",
  supportEmail: "pfrank@aiuniv.edu",
  resources: [
    {
      label: "AIU APA Guide",
      url: "https://careered.libguides.com/AIUS/APA",
    },
  ],
};

const STATUS_ORDER = {
  fail: 0,
  review: 1,
  pass: 2,
};

function sortChecks(checks) {
  return [...checks].sort((a, b) => STATUS_ORDER[a.status] - STATUS_ORDER[b.status]);
}

function SummaryCell({ count, label, href }) {
  const Tag = href ? "a" : "div";
  return (
    <Tag className={`summary-cell${href ? " summary-cell-link" : ""}`} href={href || undefined}>
      <dt>{label}</dt>
      <dd>{count}</dd>
    </Tag>
  );
}

function Summary({ report }) {
  const { failed, review, passed } = report.summary;
  const isReady = failed === 0;
  const issueText = failed === 1 ? "issue" : "issues";

  return (
    <section className="summary" aria-labelledby="summary-heading">
      <div>
        <p className="eyebrow">Report summary</p>
        <h2 id="summary-heading">
          {isReady ? "You're ready to submit" : `${failed} ${issueText} to fix`}
        </h2>
        <p className="summary-copy">
          {isReady && review > 0
            ? `${review} optional ${review === 1 ? "item" : "items"} to review. Your paper stays in this browser tab.`
            : `APA Coach checked ${report.summary.totalChecks} items in ${report.file}. Your paper stays in this browser tab.`}
        </p>
      </div>
      <dl className="summary-grid" aria-label="Check totals">
        <SummaryCell count={failed} label="Failed" href={failed > 0 ? "#fail-section" : null} />
        <SummaryCell count={review} label="Review" href={review > 0 ? "#review-section" : null} />
        <SummaryCell count={passed} label="Passed" href={passed > 0 ? "#pass-section" : null} />
      </dl>
    </section>
  );
}

function CheckCard({ check }) {
  const fixId = useId();
  const [open, setOpen] = useState(false);
  const hasFixes = check.howToFix && check.howToFix.length > 0;
  const hasResources = check.resources && check.resources.length > 0;

  return (
    <article className={`check-card ${check.status}`}>
      <div className="check-header">
        <span className="status-badge">{check.status}</span>
        <h3>{check.rule}</h3>
      </div>
      <p className="found">{check.foundText || check.found}</p>
      <p className="expected">{check.expectedText || check.expected}</p>
      {hasFixes ? (
        <div className="fixes">
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
          {check.rule === "References formatting" ? (
            <div className="hanging-indent-example">
              <p className="hanging-indent-example-label">Example of a hanging indent:</p>
              <div className="hanging-indent-demo">
                <p className="hanging-indent-entry">
                  Smith, J. A. (2023). <em>The title of a really long book that wraps onto a second line to show the indent effect.</em> Publisher Name.
                </p>
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
          {check.rule === "Inline citations" ? (
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
              <p className="citation-example-label">Each reference needs a matching inline citation:</p>
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
              <p className="citation-example-label">Each inline citation needs a matching reference entry:</p>
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
                    <p className="citation-ref-entry">Intel. (2023). <em>Thermal design in modern processors.</em> https://www.intel.com/content/www/us/en/gaming/resources/cpu-cooler.html</p>
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
          {check.missingItems && check.missingItems.length > 0 ? (
            <div className="missing-items">
              <p className="missing-items-label">{check.missingItemsLabel}</p>
              <ul>
                {check.missingItems.map((item, i) => (
                  <li key={i}>{item}</li>
                ))}
              </ul>
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
              {check.howToFix.map((step) => (
                <li key={step}>{step}</li>
              ))}
            </ol>
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
  const reviewChecks = sortedChecks.filter((c) => c.status === "review");
  const passChecks = sortedChecks.filter((c) => c.status === "pass");

  return (
    <main className="report" aria-live="polite">
      <Summary report={report} />
      <section className="checks" aria-label="APA checks">
        {failChecks.length > 0 && (
          <section id="fail-section" className="check-group" aria-labelledby="fail-heading">
            <h3 id="fail-heading" className="check-group-heading">Issues to fix (Required)</h3>
            <p className="fail-intro">These items should be corrected before you submit your paper.</p>
            {failChecks.map((check) => (
              <CheckCard key={check.rule} check={check} />
            ))}
            <a href="#top" className="back-to-top">↑ Back to top</a>
          </section>
        )}
        {reviewChecks.length > 0 && (
          <section id="review-section" className="check-group" aria-labelledby="review-heading">
            <h3 id="review-heading" className="check-group-heading">Optional checks (Review)</h3>
            <p className="review-intro">
              These items may not require changes, but are worth double-checking.
            </p>
            {reviewChecks.map((check) => (
              <CheckCard key={check.rule} check={check} />
            ))}
            <a href="#top" className="back-to-top">↑ Back to top</a>
          </section>
        )}
        {passChecks.length > 0 && (
          <section id="pass-section" className="check-group" aria-labelledby="pass-heading">
            <h3 id="pass-heading" className="check-group-heading">Looks good (Passed)</h3>
            <p className="pass-intro">These items meet APA expectations.</p>
            {passChecks.map((check) => (
              <CheckCard key={check.rule} check={check} />
            ))}
            <a href="#top" className="back-to-top">↑ Back to top</a>
          </section>
        )}
      </section>
    </main>
  );
}

function DocxDropZone({ fileName, isAnalyzing, onFileSelected }) {
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

function AppInfoCard() {
  return (
    <aside className="app-info-card" aria-label="Application information">
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
    </aside>
  );
}

function App() {
  const [report, setReport] = useState(null);
  const [error, setError] = useState("");
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [fileName, setFileName] = useState("");

  async function analyzeSelectedFile(file) {
    setReport(null);
    setError("");

    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith(".docx")) {
      setFileName("");
      setError("APA Coach can only analyze Word documents saved as .docx files.");
      return;
    }

    setFileName(file.name);

    try {
      setIsAnalyzing(true);
      setReport(await analyzeDocxFile(file));
    } catch (analysisError) {
      setError(analysisError.message || "APA Coach could not read this .docx file.");
    } finally {
      setIsAnalyzing(false);
    }
  }

  return (
    <div id="top" className="app-shell">
      <header className="app-header">
        <div>
          <a href="https://www.aiuniv.edu/" target="_blank" rel="noreferrer">
            <img className="brand-logo" src={aiuLogoUrl} alt="AIU" />
          </a>
          <p className="eyebrow">AIU APA Coach</p>
          <h1>Check APA Format</h1>
          <p>
            Submit a Word document to verify its APA formatting. Files are not uploaded, stored, or saved.
          </p>
        </div>
        <DocxDropZone fileName={fileName} isAnalyzing={isAnalyzing} onFileSelected={analyzeSelectedFile} />
      </header>

      {!report ? <AppInfoCard /> : null}
      {isAnalyzing ? (
        <p className="notice" role="status">
          Analyzing document...
        </p>
      ) : null}
      {error ? (
        <p className="error" role="alert">
          {error}
        </p>
      ) : null}
      {report ? <Report report={report} /> : null}
      {report ? <AppInfoCard /> : null}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);

import React, { useId, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { analyzeDocxFile } from "./apaBrowser.js";
import "./styles.css";

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
          <p className="eyebrow">APA Coach</p>
          <h1>Check APA Format</h1>
          <p>
            Upload a Word document to verify its APA formatting. Files are not uploaded, stored, or saved.
          </p>
        </div>
        <DocxDropZone fileName={fileName} isAnalyzing={isAnalyzing} onFileSelected={analyzeSelectedFile} />
      </header>

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
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);

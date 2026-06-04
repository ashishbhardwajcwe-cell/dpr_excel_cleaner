import { useEffect, useRef, useState } from "react";
import {
  readWorkbook,
  cleanWorkbook,
  buildCleanedXlsx,
  writeXlsx,
} from "./lib/cleaner.js";
import { scanForIRCCodes, loadIRCIndex } from "./lib/irc.js";
import { toMarkdown, toJSON, downloadBlob } from "./lib/exporter.js";
import "./App.css";

const DEFAULT_OPTIONS = {
  collapseBlankRows: true,
  dropEmptyMidColumns: true,
  dropEmptyMidRows: false,
};

export default function App() {
  const [stage, setStage] = useState("idle");
  const [file, setFile] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [ircCodes, setIrcCodes] = useState(new Map());
  const [ircLibrary, setIrcLibrary] = useState(null);
  const [error, setError] = useState("");
  const [options, setOptions] = useState(DEFAULT_OPTIONS);
  const [activeSheet, setActiveSheet] = useState(0);
  const inputRef = useRef(null);

  useEffect(() => {
    loadIRCIndex().then(setIrcLibrary);
  }, []);

  async function processFile(f) {
    setFile(f);
    setStage("analyzing");
    setError("");
    setSheets([]);
    setIrcCodes(new Map());
    setActiveSheet(0);
    try {
      const wb = await readWorkbook(f);
      const cleaned = cleanWorkbook(wb, options);
      const codes = scanForIRCCodes(cleaned);
      setSheets(cleaned);
      setIrcCodes(codes);
      setStage("done");
    } catch (err) {
      setError(err.message || "Processing failed");
      setStage("error");
    }
  }

  async function reprocess(nextOptions) {
    if (!file) return;
    setOptions(nextOptions);
    setStage("analyzing");
    try {
      const wb = await readWorkbook(file);
      const cleaned = cleanWorkbook(wb, nextOptions);
      const codes = scanForIRCCodes(cleaned);
      setSheets(cleaned);
      setIrcCodes(codes);
      setStage("done");
    } catch (err) {
      setError(err.message || "Processing failed");
      setStage("error");
    }
  }

  function handleDrop(e) {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f) processFile(f);
  }

  function handleSelect(e) {
    const f = e.target.files[0];
    if (f) processFile(f);
    e.target.value = "";
  }

  function baseName() {
    return (file?.name || "dpr").replace(/\.(xlsx|xls|xlsm|csv)$/i, "");
  }

  function downloadXlsx() {
    const wb = buildCleanedXlsx(sheets);
    writeXlsx(wb, `${baseName()}_cleaned.xlsx`);
  }

  function downloadMarkdown() {
    const md = toMarkdown(sheets, {
      fileName: file?.name,
      generatedAt: new Date().toISOString(),
      ircCodes,
    });
    downloadBlob(md, `${baseName()}_cleaned.md`, "text/markdown;charset=utf-8");
  }

  function downloadJSON() {
    const j = toJSON(sheets, {
      fileName: file?.name,
      generatedAt: new Date().toISOString(),
      ircCodes,
    });
    downloadBlob(j, `${baseName()}_cleaned.json`, "application/json;charset=utf-8");
  }

  function reset() {
    setStage("idle");
    setFile(null);
    setSheets([]);
    setIrcCodes(new Map());
    setError("");
    setOptions(DEFAULT_OPTIONS);
    setActiveSheet(0);
  }

  const kept = sheets.filter((s) => !s.empty);
  const skipped = sheets.filter((s) => s.empty);
  const totalOrigRows = sheets.reduce((a, s) => a + (s.originalRowCount || 0), 0);
  const totalKeptRows = kept.reduce((a, s) => a + (s.cleanedRowCount || 0), 0);
  const totalRemovedCols = sheets.reduce((a, s) => a + (s.removedColumnIndexes?.length || 0), 0);
  const totalCollapsed = sheets.reduce((a, s) => a + (s.collapsedBlankRowRuns || 0), 0);

  return (
    <div className="app">
      <header className="app-header">
        <div className="app-logo">DPR</div>
        <div>
          <h1>DPR Excel Cleaner</h1>
          <p className="muted">
            Strip phantom cells, trailing empties and whitespace junk from DPR workbooks —
            without losing any actual data. Outputs a cleaned <code>.xlsx</code>, a Claude-ready
            Markdown report, and a structured JSON payload.
          </p>
        </div>
      </header>

      {stage === "idle" && (
        <section
          className="drop"
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDrop}
          onClick={() => inputRef.current?.click()}
        >
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={handleSelect}
            hidden
          />
          <div className="drop-icon">+</div>
          <p className="drop-title">Drop your DPR workbook here</p>
          <p className="drop-sub">
            or click to browse — <code>.xlsx</code> / <code>.xls</code> / <code>.xlsm</code> · processed
            entirely in your browser, nothing uploaded.
          </p>
        </section>
      )}

      {stage === "analyzing" && (
        <section className="status">
          <div className="spinner" aria-hidden="true" />
          <p className="status-title">Scanning sheets…</p>
          <p className="muted">{file?.name}</p>
        </section>
      )}

      {stage === "error" && (
        <section className="error">
          <p className="error-title">Processing failed</p>
          <p>{error}</p>
          <button onClick={reset} className="btn">Try another file</button>
        </section>
      )}

      {stage === "done" && (
        <section className="results">
          <div className="stats">
            <Stat label="Sheets" value={sheets.length} />
            <Stat label="With data" value={kept.length} accent="ok" />
            <Stat label="Empty" value={skipped.length} accent={skipped.length ? "warn" : ""} />
            <Stat label="Rows kept" value={totalKeptRows.toLocaleString()} />
            <Stat label="Original rows" value={totalOrigRows.toLocaleString()} accent="muted" />
            <Stat label="Empty cols removed" value={totalRemovedCols} />
            <Stat label="Blank-row runs collapsed" value={totalCollapsed} />
            <Stat label="IRC codes found" value={ircCodes.size} accent={ircCodes.size ? "ok" : ""} />
          </div>

          <div className="actions">
            <button className="btn primary" onClick={downloadXlsx}>Download cleaned .xlsx</button>
            <button className="btn" onClick={downloadMarkdown}>Download Markdown (Claude-ready)</button>
            <button className="btn" onClick={downloadJSON}>Download JSON</button>
            <button className="btn ghost" onClick={reset}>New file</button>
          </div>

          <CleaningOptions
            options={options}
            onChange={reprocess}
          />

          {ircCodes.size > 0 && (
            <details className="panel" open>
              <summary>IRC codes referenced in this DPR ({ircCodes.size})</summary>
              <table className="tbl">
                <thead>
                  <tr><th>Code</th><th>Occurrences</th><th>First location</th><th>In library</th></tr>
                </thead>
                <tbody>
                  {Array.from(ircCodes.entries()).map(([code, occs]) => {
                    const inLib = ircLibrary?.files?.some((f) => f.code === code) || false;
                    return (
                      <tr key={code}>
                        <td><code>{code}</code></td>
                        <td>{occs.length}</td>
                        <td className="muted">
                          {occs[0].sheet} · r{occs[0].row + 1} c{occs[0].col + 1}
                        </td>
                        <td>{inLib ? <span className="pill ok">loaded</span> : <span className="pill muted">missing</span>}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {ircLibrary && (
                <p className="muted small">
                  IRC library: version {ircLibrary.version ?? "?"} · {ircLibrary.files?.length ?? 0} reference files loaded from <code>/irc-codes/</code>.
                </p>
              )}
              {!ircLibrary && (
                <p className="muted small">
                  No <code>public/irc-codes/INDEX.json</code> found. Drop your IRC v63/v64 <code>.md</code> files
                  into <code>public/irc-codes/</code> and list them in <code>INDEX.json</code> to enable cross-referencing.
                </p>
              )}
            </details>
          )}

          {kept.length > 0 && (
            <div className="panel">
              <p className="panel-title">Per-sheet summary</p>
              <table className="tbl">
                <thead>
                  <tr>
                    <th>Sheet</th>
                    <th>Orig rows</th>
                    <th>Kept rows</th>
                    <th>Orig cols</th>
                    <th>Kept cols</th>
                    <th>Empty cols removed</th>
                    <th>Blank runs collapsed</th>
                  </tr>
                </thead>
                <tbody>
                  {kept.map((s, i) => (
                    <tr
                      key={s.name}
                      className={i === activeSheet ? "active" : ""}
                      onClick={() => setActiveSheet(i)}
                      style={{ cursor: "pointer" }}
                    >
                      <td><code>{s.name}</code></td>
                      <td>{s.originalRowCount.toLocaleString()}</td>
                      <td>{s.cleanedRowCount.toLocaleString()}</td>
                      <td>{s.originalColCount}</td>
                      <td>{s.cleanedColCount}</td>
                      <td>{s.removedColumnIndexes.length}</td>
                      <td>{s.collapsedBlankRowRuns}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {skipped.length > 0 && (
            <div className="panel">
              <p className="panel-title">Empty sheets ({skipped.length})</p>
              <div className="chips">
                {skipped.map((s) => (
                  <span key={s.name} className="chip">{s.name}</span>
                ))}
              </div>
            </div>
          )}

          {kept[activeSheet] && (
            <SheetPreview sheet={kept[activeSheet]} />
          )}

          <div className="hint">
            <strong>Next step:</strong> upload the cleaned <code>.xlsx</code> (or paste the Markdown export)
            into your DPR Analyzer (Ver63 / Ver64). Cell positions, sheet names, units and quantities are preserved,
            so IRC clause citations resolve correctly.
          </div>
        </section>
      )}

      <footer className="foot muted small">
        100% client-side · your workbook never leaves your browser · deploy to Netlify by pushing this repo and connecting it.
      </footer>
    </div>
  );
}

function Stat({ label, value, accent }) {
  return (
    <div className={`stat stat-${accent || ""}`}>
      <div className="stat-label">{label}</div>
      <div className="stat-value">{value}</div>
    </div>
  );
}

function CleaningOptions({ options, onChange }) {
  return (
    <details className="panel">
      <summary>Cleaning options</summary>
      <div className="opt-grid">
        <label>
          <input
            type="checkbox"
            checked={options.collapseBlankRows}
            onChange={(e) => onChange({ ...options, collapseBlankRows: e.target.checked })}
          />
          <span>Collapse runs of blank rows down to 1 (preserves separators)</span>
        </label>
        <label>
          <input
            type="checkbox"
            checked={options.dropEmptyMidColumns}
            onChange={(e) => onChange({ ...options, dropEmptyMidColumns: e.target.checked })}
          />
          <span>Remove fully-empty columns (including phantom ones)</span>
        </label>
        <label>
          <input
            type="checkbox"
            checked={options.dropEmptyMidRows}
            onChange={(e) => onChange({ ...options, dropEmptyMidRows: e.target.checked })}
          />
          <span>Remove <em>all</em> blank rows (more aggressive — may merge BOQ sections)</span>
        </label>
      </div>
      <p className="muted small">
        Data inside non-empty cells is never modified beyond whitespace normalization
        (trim, collapse runs of spaces, normalize line endings). Numbers and dates are preserved.
      </p>
    </details>
  );
}

function SheetPreview({ sheet }) {
  const MAX_ROWS = 200;
  const MAX_COLS = 30;
  const rows = sheet.rows.slice(0, MAX_ROWS);
  const truncRows = sheet.rows.length > MAX_ROWS;
  const cols = sheet.cleanedColCount;
  const truncCols = cols > MAX_COLS;
  const showCols = Math.min(cols, MAX_COLS);

  return (
    <div className="panel">
      <p className="panel-title">
        Preview: <code>{sheet.name}</code>{" "}
        <span className="muted small">
          showing {rows.length}/{sheet.cleanedRowCount} rows × {showCols}/{cols} cols
          {(truncRows || truncCols) && " (preview only — full data is in the export)"}
        </span>
      </p>
      <div className="preview-scroll">
        <table className="preview-tbl">
          <tbody>
            {rows.map((row, ri) => (
              <tr key={ri} className={ri === sheet.headerRowIndex ? "header-row" : ""}>
                <td className="rownum">{ri + 1}</td>
                {row.slice(0, showCols).map((cell, ci) => (
                  <td key={ci}>{cell === null ? "" : String(cell)}</td>
                ))}
                {truncCols && <td className="muted">…</td>}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const MAX_OUTPUT_CHARS = 120000;
const MAX_ROWS_PER_SHEET = 600;
const MAX_COLS = 60;
const MIN_DATA_ROWS = 2;

function readExcelForCleaning(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = window.XLSX.read(data, { type: "array" });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    r.onerror = () => reject(new Error("File read failed"));
    r.readAsArrayBuffer(file);
  });
}

function findLastMeaningfulCol(rows, maxCols) {
  let last = 0;
  for (const row of rows) {
    for (let c = Math.min(row.length, maxCols) - 1; c >= last; c--) {
      const v = row[c];
      if (v !== null && v !== undefined && String(v).trim() !== "") {
        last = Math.max(last, c + 1);
        break;
      }
    }
  }
  return Math.min(last, maxCols);
}

function sheetStats(ws) {
  const json = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const dataRows = json.filter(row => row.some(c => c !== null && c !== undefined && String(c).trim() !== ""));
  return { totalRows: json.length, dataRows };
}

function cleanSheet(ws, sheetName) {
  const json = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const dataRows = json.filter(row => row.some(c => c !== null && c !== undefined && String(c).trim() !== ""));

  if (dataRows.length < MIN_DATA_ROWS) {
    return { skipped: true, reason: "too few data rows (" + dataRows.length + ")" };
  }

  const lastCol = findLastMeaningfulCol(dataRows, MAX_COLS);

  const trimmed = dataRows.slice(0, MAX_ROWS_PER_SHEET).map((row, rowIdx) => {
    const out = [];
    for (let c = 0; c < lastCol; c++) {
      const v = row[c];
      out.push(v !== null && v !== undefined ? String(v).trim() : "");
    }
    return out;
  });

  const truncated = dataRows.length > MAX_ROWS_PER_SHEET;
  const colsTruncated = findLastMeaningfulCol(dataRows, Infinity) > MAX_COLS;

  return {
    skipped: false,
    rows: trimmed,
    originalRows: dataRows.length,
    originalCols: findLastMeaningfulCol(dataRows, Infinity),
    keptRows: trimmed.length,
    keptCols: lastCol,
    truncated,
    colsTruncated,
  };
}

function buildCleanWorkbook(results) {
  const wb = window.XLSX.utils.book_new();
  for (const r of results) {
    if (r.skipped) continue;
    const ws = window.XLSX.utils.aoa_to_sheet(r.rows);
    window.XLSX.utils.book_append_sheet(wb, ws, r.sheetName.slice(0, 31));
  }
  return wb;
}

function buildTextPreview(results) {
  let txt = "";
  let chars = 0;
  for (const r of results) {
    if (r.skipped) continue;
    const header = `\n=== Sheet: ${r.sheetName} (${r.keptRows} rows × ${r.keptCols} cols)${r.truncated ? " [ROWS TRUNCATED]" : ""}${r.colsTruncated ? " [COLS TRUNCATED]" : ""} ===\n`;
    txt += header;
    chars += header.length;
    for (const row of r.rows) {
      const line = row.join(",") + "\n";
      if (chars + line.length > MAX_OUTPUT_CHARS) {
        txt += "[...remaining content truncated for size limit...]\n";
        return txt;
      }
      txt += line;
      chars += line.length;
    }
  }
  return txt;
}

export default function App() {
  const [stage, setStage] = useState("idle"); // idle | analyzing | done | error
  const [file, setFile] = useState(null);
  const [results, setResults] = useState([]);
  const [summary, setSummary] = useState(null);
  const [error, setError] = useState("");
  const [textPreview, setTextPreview] = useState("");
  const [showPreview, setShowPreview] = useState(false);
  const inputRef = useRef(null);

  async function processFile(f) {
    setFile(f);
    setStage("analyzing");
    setError("");
    setResults([]);
    setSummary(null);
    setTextPreview("");
    setShowPreview(false);

    try {
      const wb = await readExcelForCleaning(f);
      const sheetResults = [];

      for (const sn of wb.SheetNames) {
        const ws = wb.Sheets[sn];
        const res = cleanSheet(ws, sn);
        sheetResults.push({ sheetName: sn, ...res });
      }

      const kept = sheetResults.filter(r => !r.skipped);
      const skipped = sheetResults.filter(r => r.skipped);
      const totalOrigCols = sheetResults.reduce((a, r) => a + (r.originalCols || 0), 0);
      const totalKeptCols = kept.reduce((a, r) => a + (r.keptCols || 0), 0);
      const totalOrigRows = sheetResults.reduce((a, r) => a + (r.originalRows || 0), 0);
      const totalKeptRows = kept.reduce((a, r) => a + (r.keptRows || 0), 0);

      setResults(sheetResults);
      setSummary({
        totalSheets: wb.SheetNames.length,
        keptSheets: kept.length,
        skippedSheets: skipped.length,
        totalOrigRows,
        totalKeptRows,
        totalOrigCols,
        totalKeptCols,
      });

      const preview = buildTextPreview(sheetResults);
      setTextPreview(preview);
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

  function downloadClean() {
    const wb = buildCleanWorkbook(results);
    const base = (file?.name || "cleaned").replace(/\.xlsx?$/i, "");
    window.XLSX.writeFile(wb, base + "_cleaned.xlsx");
  }

  function downloadTextPreview() {
    const blob = new Blob([textPreview], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = (file?.name || "dpr").replace(/\.xlsx?$/i, "") + "_text_preview.txt";
    a.click();
    URL.revokeObjectURL(url);
  }

  function reset() {
    setStage("idle");
    setFile(null);
    setResults([]);
    setSummary(null);
    setError("");
    setTextPreview("");
    setShowPreview(false);
  }

  const kept = results.filter(r => !r.skipped);
  const skipped = results.filter(r => r.skipped);

  return (
    <>
      <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css" />
      <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js" />

      <h2 className="sr-only">DPR Excel Cleaner — pre-process large Excel DPRs before uploading to Ver63</h2>

      <div style={{ padding: "1.5rem 0 2rem" }}>

        {/* Header */}
        <div style={{ marginBottom: "1.5rem" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "6px" }}>
            <div style={{
              width: 36, height: 36, borderRadius: "var(--border-radius-md)",
              background: "var(--color-background-info)", display: "flex",
              alignItems: "center", justifyContent: "center"
            }}>
              <i className="ti ti-table-import" style={{ fontSize: 18, color: "var(--color-text-info)" }} aria-hidden="true" />
            </div>
            <div>
              <p style={{ margin: 0, fontWeight: 500, fontSize: 17 }}>DPR Excel Cleaner</p>
              <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-secondary)" }}>Pre-processor for Ver63 — strips phantom columns, trims oversized sheets</p>
            </div>
          </div>
        </div>

        {/* Upload zone */}
        {stage === "idle" && (
          <div
            onDragOver={(e) => e.preventDefault()}
            onDrop={handleDrop}
            onClick={() => inputRef.current?.click()}
            style={{
              border: "0.5px dashed var(--color-border-primary)",
              borderRadius: "var(--border-radius-lg)",
              padding: "2.5rem 1.5rem",
              textAlign: "center",
              cursor: "pointer",
              background: "var(--color-background-secondary)",
            }}
          >
            <input ref={inputRef} type="file" accept=".xlsx,.xls" onChange={handleSelect} style={{ display: "none" }} />
            <i className="ti ti-file-spreadsheet" style={{ fontSize: 40, color: "var(--color-text-secondary)", display: "block", marginBottom: 12 }} aria-hidden="true" />
            <p style={{ margin: "0 0 6px", fontWeight: 500 }}>Drop your DPR Excel here</p>
            <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-secondary)" }}>or click to browse — .xlsx / .xls only</p>
          </div>
        )}

        {/* Analyzing */}
        {stage === "analyzing" && (
          <div style={{ textAlign: "center", padding: "3rem 1rem" }}>
            <i className="ti ti-loader-2" style={{ fontSize: 36, color: "var(--color-text-secondary)", display: "block", marginBottom: 12, animation: "spin 1s linear infinite" }} aria-hidden="true" />
            <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
            <p style={{ margin: 0, fontWeight: 500 }}>Scanning sheets...</p>
            <p style={{ margin: "4px 0 0", fontSize: 13, color: "var(--color-text-secondary)" }}>{file?.name}</p>
          </div>
        )}

        {/* Error */}
        {stage === "error" && (
          <div style={{
            background: "var(--color-background-danger)",
            border: "0.5px solid var(--color-border-danger)",
            borderRadius: "var(--border-radius-md)", padding: "1rem 1.25rem", marginBottom: 16
          }}>
            <p style={{ margin: "0 0 4px", fontWeight: 500, color: "var(--color-text-danger)" }}>Processing failed</p>
            <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-danger)" }}>{error}</p>
            <button onClick={reset} style={{ marginTop: 10 }}>Try another file</button>
          </div>
        )}

        {/* Results */}
        {stage === "done" && summary && (
          <div>
            {/* Stats row */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: 10, marginBottom: "1.5rem" }}>
              {[
                { label: "Total sheets", val: summary.totalSheets },
                { label: "Sheets kept", val: summary.keptSheets, ok: true },
                { label: "Sheets skipped", val: summary.skippedSheets, warn: summary.skippedSheets > 0 },
                { label: "Rows kept", val: summary.totalKeptRows.toLocaleString() },
                { label: "Max cols/sheet", val: MAX_COLS },
              ].map(s => (
                <div key={s.label} style={{
                  background: "var(--color-background-secondary)",
                  borderRadius: "var(--border-radius-md)", padding: "0.75rem 1rem"
                }}>
                  <p style={{ margin: "0 0 2px", fontSize: 12, color: "var(--color-text-secondary)" }}>{s.label}</p>
                  <p style={{
                    margin: 0, fontWeight: 500, fontSize: 20,
                    color: s.ok ? "var(--color-text-success)" : s.warn ? "var(--color-text-warning)" : "var(--color-text-primary)"
                  }}>{s.val}</p>
                </div>
              ))}
            </div>

            {/* Actions */}
            <div style={{ display: "flex", gap: 10, marginBottom: "1.5rem", flexWrap: "wrap" }}>
              <button onClick={downloadClean} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <i className="ti ti-download" aria-hidden="true" />
                Download cleaned .xlsx
              </button>
              <button onClick={downloadTextPreview} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <i className="ti ti-file-text" aria-hidden="true" />
                Download text preview
              </button>
              <button onClick={() => setShowPreview(!showPreview)} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <i className={`ti ti-eye${showPreview ? "-off" : ""}`} aria-hidden="true" />
                {showPreview ? "Hide" : "Show"} preview
              </button>
              <button onClick={reset} style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6 }}>
                <i className="ti ti-refresh" aria-hidden="true" />
                New file
              </button>
            </div>

            {/* Kept sheets */}
            {kept.length > 0 && (
              <div style={{ marginBottom: "1.5rem" }}>
                <p style={{ margin: "0 0 8px", fontWeight: 500, fontSize: 14 }}>
                  <i className="ti ti-check" style={{ color: "var(--color-text-success)", marginRight: 6 }} aria-hidden="true" />
                  Kept sheets ({kept.length})
                </p>
                <div style={{
                  border: "0.5px solid var(--color-border-tertiary)",
                  borderRadius: "var(--border-radius-md)", overflow: "hidden"
                }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ background: "var(--color-background-secondary)" }}>
                        {["Sheet name", "Orig rows", "Kept rows", "Orig cols", "Kept cols", "Notes"].map(h => (
                          <th key={h} style={{ padding: "7px 10px", textAlign: "left", fontWeight: 500, borderBottom: "0.5px solid var(--color-border-tertiary)", color: "var(--color-text-secondary)" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {kept.map((r, i) => (
                        <tr key={r.sheetName} style={{ background: i % 2 === 0 ? "transparent" : "var(--color-background-secondary)" }}>
                          <td style={{ padding: "6px 10px", fontFamily: "monospace", fontSize: 12 }}>{r.sheetName}</td>
                          <td style={{ padding: "6px 10px", color: "var(--color-text-secondary)" }}>{r.originalRows?.toLocaleString()}</td>
                          <td style={{ padding: "6px 10px", color: r.truncated ? "var(--color-text-warning)" : "var(--color-text-primary)", fontWeight: r.truncated ? 500 : 400 }}>{r.keptRows}</td>
                          <td style={{ padding: "6px 10px", color: "var(--color-text-secondary)" }}>{r.originalCols?.toLocaleString()}</td>
                          <td style={{ padding: "6px 10px", color: r.colsTruncated ? "var(--color-text-warning)" : "var(--color-text-primary)", fontWeight: r.colsTruncated ? 500 : 400 }}>{r.keptCols}</td>
                          <td style={{ padding: "6px 10px", fontSize: 11, color: "var(--color-text-secondary)" }}>
                            {[r.truncated && "rows capped at " + MAX_ROWS_PER_SHEET, r.colsTruncated && "cols capped at " + MAX_COLS].filter(Boolean).join("; ") || "—"}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Skipped sheets */}
            {skipped.length > 0 && (
              <div style={{ marginBottom: "1.5rem" }}>
                <p style={{ margin: "0 0 8px", fontWeight: 500, fontSize: 14, color: "var(--color-text-secondary)" }}>
                  <i className="ti ti-minus" style={{ marginRight: 6 }} aria-hidden="true" />
                  Skipped sheets ({skipped.length}) — empty or near-empty
                </p>
                <div style={{
                  display: "flex", flexWrap: "wrap", gap: 6
                }}>
                  {skipped.map(r => (
                    <span key={r.sheetName} style={{
                      padding: "3px 10px", borderRadius: "var(--border-radius-md)",
                      background: "var(--color-background-secondary)",
                      border: "0.5px solid var(--color-border-tertiary)",
                      fontSize: 12, fontFamily: "monospace", color: "var(--color-text-tertiary)"
                    }}>{r.sheetName}</span>
                  ))}
                </div>
              </div>
            )}

            {/* Instructions */}
            <div style={{
              background: "var(--color-background-info)",
              border: "0.5px solid var(--color-border-info)",
              borderRadius: "var(--border-radius-md)", padding: "1rem 1.25rem", marginBottom: "1.5rem"
            }}>
              <p style={{ margin: "0 0 6px", fontWeight: 500, fontSize: 14, color: "var(--color-text-info)" }}>
                <i className="ti ti-info-circle" style={{ marginRight: 6 }} aria-hidden="true" />
                Next step
              </p>
              <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-info)", lineHeight: 1.6 }}>
                Download the cleaned .xlsx above, then upload it into Ver63 (DPR Analyzer Pro). All sheet names and cell positions are preserved, so clause citations will reference the correct location.
              </p>
            </div>

            {/* Text preview */}
            {showPreview && (
              <div>
                <p style={{ margin: "0 0 8px", fontWeight: 500, fontSize: 14 }}>Text preview (first ~120K chars)</p>
                <pre style={{
                  background: "var(--color-background-secondary)",
                  border: "0.5px solid var(--color-border-tertiary)",
                  borderRadius: "var(--border-radius-md)",
                  padding: "0.75rem 1rem",
                  fontSize: 11, fontFamily: "monospace",
                  overflowX: "auto", whiteSpace: "pre-wrap",
                  maxHeight: 400, overflowY: "auto",
                  lineHeight: 1.5
                }}>{textPreview.slice(0, 12000)}{textPreview.length > 12000 ? "\n\n[... preview truncated — download full text file for complete content ...]" : ""}</pre>
              </div>
            )}
          </div>
        )}
      </div>
    </>
  );
}

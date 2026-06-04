function fmtCell(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return String(v);
}

function mdEscape(s) {
  return s.replace(/\|/g, "\\|").replace(/\n/g, "<br>");
}

export function toMarkdown(sheets, meta = {}) {
  const lines = [];
  lines.push(`# DPR — Cleaned workbook`);
  if (meta.fileName) lines.push(`Source file: \`${meta.fileName}\``);
  if (meta.generatedAt) lines.push(`Generated: ${meta.generatedAt}`);
  lines.push("");
  lines.push(`Sheets: **${sheets.length}** (${sheets.filter((s) => !s.empty).length} with data, ${sheets.filter((s) => s.empty).length} empty)`);
  lines.push("");

  if (meta.ircCodes && meta.ircCodes.size > 0) {
    lines.push(`## IRC codes referenced`);
    lines.push("");
    lines.push(`| Code | Occurrences | Locations |`);
    lines.push(`| --- | --- | --- |`);
    for (const [code, occs] of meta.ircCodes.entries()) {
      const locs = occs.slice(0, 5).map((o) => `${o.sheet} r${o.row + 1}c${o.col + 1}`).join("; ");
      const more = occs.length > 5 ? ` (+${occs.length - 5} more)` : "";
      lines.push(`| ${mdEscape(code)} | ${occs.length} | ${mdEscape(locs)}${more} |`);
    }
    lines.push("");
  }

  for (const s of sheets) {
    lines.push(`---`);
    lines.push("");
    lines.push(`## Sheet: ${s.name}`);
    if (s.empty) {
      lines.push("");
      lines.push(`_(empty — skipped)_`);
      lines.push("");
      continue;
    }
    lines.push("");
    lines.push(`Dimensions: ${s.cleanedRowCount} rows × ${s.cleanedColCount} cols  `);
    lines.push(`Original: ${s.originalRowCount} rows × ${s.originalColCount} cols  `);
    const removed = [];
    if (s.removedColumnIndexes.length) removed.push(`${s.removedColumnIndexes.length} empty col(s) removed`);
    if (s.collapsedBlankRowRuns) removed.push(`${s.collapsedBlankRowRuns} blank row(s) collapsed`);
    if (removed.length) lines.push(`Cleanup: ${removed.join(", ")}  `);
    lines.push("");

    const headerIdx = s.headerRowIndex >= 0 ? s.headerRowIndex : 0;
    const headerRow = s.rows[headerIdx] || [];
    const headers = headerRow.map((h, i) => {
      const t = fmtCell(h).trim();
      return t === "" ? `col${i + 1}` : t;
    });

    if (headerIdx > 0) {
      lines.push(`**Pre-header rows (${headerIdx}):**`);
      lines.push("");
      lines.push("```");
      for (let i = 0; i < headerIdx; i++) {
        lines.push(s.rows[i].map(fmtCell).join(" | "));
      }
      lines.push("```");
      lines.push("");
    }

    lines.push(`| ${headers.map(mdEscape).join(" | ")} |`);
    lines.push(`| ${headers.map(() => "---").join(" | ")} |`);
    for (let i = headerIdx + 1; i < s.rows.length; i++) {
      const row = s.rows[i].map((c) => mdEscape(fmtCell(c)));
      while (row.length < headers.length) row.push("");
      lines.push(`| ${row.slice(0, headers.length).join(" | ")} |`);
    }
    lines.push("");
  }

  return lines.join("\n");
}

export function toJSON(sheets, meta = {}) {
  const out = {
    schemaVersion: 1,
    generator: "dpr-excel-cleaner",
    generatedAt: meta.generatedAt || new Date().toISOString(),
    fileName: meta.fileName || null,
    workbook: {
      sheetCount: sheets.length,
      sheets: sheets.map((s) => ({
        name: s.name,
        empty: !!s.empty,
        headerRowIndex: s.headerRowIndex,
        originalRowCount: s.originalRowCount,
        originalColCount: s.originalColCount,
        cleanedRowCount: s.cleanedRowCount,
        cleanedColCount: s.cleanedColCount,
        removedColumnIndexes: s.removedColumnIndexes,
        collapsedBlankRowRuns: s.collapsedBlankRowRuns,
        rows: s.rows,
      })),
    },
    ircCodes: meta.ircCodes
      ? Array.from(meta.ircCodes.entries()).map(([code, occs]) => ({ code, occurrences: occs }))
      : [],
  };
  return JSON.stringify(out, null, 2);
}

export function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

import * as XLSX from "xlsx";

export function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {
          type: "array",
          cellDates: true,
          cellNF: false,
          cellStyles: false,
        });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    r.onerror = () => reject(new Error("File read failed"));
    r.readAsArrayBuffer(file);
  });
}

export function isEmptyCell(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === "string" && v.trim() === "") return true;
  return false;
}

export function normalizeCell(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) {
    if (Number.isNaN(v.getTime())) return null;
    return v.toISOString().slice(0, 10);
  }
  if (typeof v === "number") {
    if (!Number.isFinite(v)) return null;
    return v;
  }
  if (typeof v === "boolean") return v;
  const s = String(v).replace(/\r\n?/g, "\n").replace(/[ \t]+/g, " ").trim();
  return s === "" ? null : s;
}

function maxColCount(rows) {
  let m = 0;
  for (const r of rows) if (r && r.length > m) m = r.length;
  return m;
}

function findBoundingBox(rows) {
  let lastRow = -1;
  let lastCol = -1;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (!isEmptyCell(row[c])) {
        if (r > lastRow) lastRow = r;
        if (c > lastCol) lastCol = c;
      }
    }
  }
  return lastRow >= 0 ? { lastRow, lastCol } : null;
}

function trimToBoundingBox(rows, bbox) {
  const out = [];
  for (let r = 0; r <= bbox.lastRow; r++) {
    const row = rows[r] || [];
    const trimmed = [];
    for (let c = 0; c <= bbox.lastCol; c++) {
      trimmed.push(c < row.length ? row[c] : null);
    }
    out.push(trimmed);
  }
  return out;
}

function findEmptyColumns(rows, width) {
  const empty = [];
  for (let c = 0; c < width; c++) {
    let allEmpty = true;
    for (const row of rows) {
      if (!isEmptyCell(row[c])) {
        allEmpty = false;
        break;
      }
    }
    if (allEmpty) empty.push(c);
  }
  return empty;
}

function dropColumns(rows, indexes) {
  const set = new Set(indexes);
  return rows.map((r) => r.filter((_, i) => !set.has(i)));
}

function dropFullyEmptyRows(rows) {
  return rows.filter((r) => r.some((c) => !isEmptyCell(c)));
}

function collapseBlankRuns(rows) {
  const out = [];
  let lastBlank = false;
  for (const r of rows) {
    const blank = r.every((c) => isEmptyCell(c));
    if (blank && lastBlank) continue;
    out.push(r);
    lastBlank = blank;
  }
  while (out.length && out[out.length - 1].every((c) => isEmptyCell(c))) {
    out.pop();
  }
  return out;
}

function inferHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i];
    const filled = row.filter((c) => !isEmptyCell(c)).length;
    const strings = row.filter((c) => typeof c === "string").length;
    if (filled >= 2 && strings >= Math.max(2, Math.floor(filled * 0.6))) {
      return i;
    }
  }
  return rows.findIndex((r) => r.some((c) => !isEmptyCell(c)));
}

export function cleanSheet(ws, name, options = {}) {
  const opts = {
    collapseBlankRows: true,
    dropEmptyMidColumns: true,
    dropEmptyMidRows: false,
    ...options,
  };

  const raw = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    blankrows: true,
    raw: false,
    rawNumbers: true,
  });

  const originalRowCount = raw.length;
  const originalColCount = maxColCount(raw);

  const bbox = findBoundingBox(raw);
  if (!bbox) {
    return {
      name,
      empty: true,
      rows: [],
      headerRowIndex: -1,
      originalRowCount,
      originalColCount,
      cleanedRowCount: 0,
      cleanedColCount: 0,
      removedColumnIndexes: [],
      collapsedBlankRowRuns: 0,
      notes: ["sheet has no non-empty cells"],
    };
  }

  let rows = trimToBoundingBox(raw, bbox);
  rows = rows.map((row) => row.map(normalizeCell));

  let removedCols = [];
  if (opts.dropEmptyMidColumns) {
    const w = rows[0]?.length || 0;
    removedCols = findEmptyColumns(rows, w);
    if (removedCols.length) rows = dropColumns(rows, removedCols);
  }

  let collapsed = 0;
  if (opts.collapseBlankRows) {
    const before = rows.length;
    rows = collapseBlankRuns(rows);
    collapsed = before - rows.length;
  }

  if (opts.dropEmptyMidRows) {
    rows = dropFullyEmptyRows(rows);
  }

  const headerRowIndex = inferHeaderRow(rows);

  return {
    name,
    empty: rows.length === 0,
    rows,
    headerRowIndex,
    originalRowCount,
    originalColCount,
    cleanedRowCount: rows.length,
    cleanedColCount: rows[0]?.length || 0,
    removedColumnIndexes: removedCols,
    collapsedBlankRowRuns: collapsed,
    notes: [],
  };
}

export function cleanWorkbook(wb, options) {
  return wb.SheetNames.map((name) => cleanSheet(wb.Sheets[name], name, options));
}

export function buildCleanedXlsx(sheets) {
  const wb = XLSX.utils.book_new();
  for (const s of sheets) {
    if (s.empty) continue;
    const ws = XLSX.utils.aoa_to_sheet(s.rows);
    XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0, 31));
  }
  return wb;
}

export function writeXlsx(wb, filename) {
  XLSX.writeFile(wb, filename);
}

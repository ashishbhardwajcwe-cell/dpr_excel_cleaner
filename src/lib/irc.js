const IRC_PATTERN = /\bIRC[\s:.-]*(?:SP[\s:.-]*)?(\d+)(?:[-/](\d{4}))?\b/gi;

export function normalizeIRCCode(raw) {
  const m = /IRC[\s:.-]*(SP[\s:.-]*)?(\d+)(?:[-/](\d{4}))?/i.exec(raw);
  if (!m) return null;
  const sp = m[1] ? "SP:" : "";
  const num = m[2];
  const year = m[3] ? `-${m[3]}` : "";
  return `IRC:${sp}${num}${year}`;
}

export function scanForIRCCodes(sheets) {
  const codes = new Map();
  for (const s of sheets) {
    if (s.empty) continue;
    for (let r = 0; r < s.rows.length; r++) {
      const row = s.rows[r];
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (typeof v !== "string") continue;
        const matches = v.matchAll(IRC_PATTERN);
        for (const m of matches) {
          const code = normalizeIRCCode(m[0]);
          if (!code) continue;
          if (!codes.has(code)) codes.set(code, []);
          codes.get(code).push({ sheet: s.name, row: r, col: c, snippet: v.slice(0, 120) });
        }
      }
    }
  }
  return codes;
}

export async function loadIRCIndex(basePath = "/irc-codes/INDEX.json") {
  try {
    const r = await fetch(basePath);
    if (!r.ok) return null;
    const j = await r.json();
    if (!j || !Array.isArray(j.files)) return null;
    return j;
  } catch {
    return null;
  }
}

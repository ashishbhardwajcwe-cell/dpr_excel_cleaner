# DPR Excel Cleaner

A pre-processor for DPR (Detailed Project Report) Excel workbooks. It strips
phantom cells, trailing empty rows/columns, and whitespace junk **without
losing any actual data**, then exports:

1. **Cleaned `.xlsx`** — same sheet names, same cell positions (minus the junk),
   ready to feed into your downstream analyzer (Ver63 / Ver64).
2. **Markdown report** — Claude-ready tables, one section per sheet, IRC codes
   listed up front. Drop this straight into a Claude conversation.
3. **Structured JSON** — same content as the Markdown, machine-readable. Useful
   if you want to wire it to the Claude API directly.

Runs entirely in the browser. No file is ever uploaded to a server.

## Local development

```bash
npm install
npm run dev
```

Then open the printed URL.

## Deploy to Netlify (no config needed)

1. Push this repo to GitHub.
2. On Netlify → **Add new site → Import an existing project** → pick this repo.
3. Netlify auto-detects the settings from `netlify.toml`:
   - Build command: `npm run build`
   - Publish directory: `dist`
4. Deploy. Done.

Or one-click via the Netlify deploy button (replace `OWNER/REPO`):

```
https://app.netlify.com/start/deploy?repository=https://github.com/OWNER/REPO
```

## What it does to your workbook

For each sheet:

| Operation | Default | Notes |
| --- | --- | --- |
| Trim to data bounding box | always on | Strips trailing empty rows/cols Excel left behind |
| Normalize whitespace inside cells | always on | Trims, collapses runs of spaces, normalizes `\r\n` |
| Remove fully-empty columns | on | Including phantom columns Excel reports |
| Collapse runs of blank rows to 1 | on | Preserves section separators |
| Drop **all** blank rows | off | More aggressive — may merge BOQ sections |

Numbers, dates, units, and quantities are **never modified**. Cell positions
within the trimmed bounding box are preserved so IRC clause citations from the
downstream analyzer still resolve correctly.

## IRC code cross-referencing

Drop your IRC standard `.md` files into `public/irc-codes/` and list them in
`public/irc-codes/INDEX.json`:

```json
{
  "version": "64",
  "files": [
    { "code": "IRC:36-2010", "file": "IRC-36-2010.md", "title": "..." },
    { "code": "IRC:SP:84-2019", "file": "IRC-SP-84-2019.md", "title": "..." }
  ]
}
```

After cleaning, the app:

- Scans all cell content for IRC code patterns (`IRC:36-2010`, `IRC SP:48`, …).
- Lists every found code, with occurrence counts and first-location pointers.
- Flags which codes have a matching `.md` file in your library and which don't.
- Includes the same list at the top of the Markdown export so Claude sees the
  IRC context immediately.

## Output formats — what Claude sees

The Markdown export starts with workbook metadata and the IRC code summary,
then one section per sheet with a real Markdown table for the data. This is
the format Claude parses most reliably — far better than dumping CSV or raw
Excel into a prompt.

The JSON export contains the same data plus per-sheet cleanup stats (rows
collapsed, cols removed, header row index, etc.) for programmatic pipelines.

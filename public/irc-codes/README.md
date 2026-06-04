# IRC reference codes

Drop your IRC standard `.md` files into this folder before building. Examples:

```
public/irc-codes/
├── IRC-36-2010.md
├── IRC-37-2018.md
├── IRC-73-1980.md
├── IRC-SP-48-1998.md
├── IRC-SP-84-2019.md
└── INDEX.json            # optional — list of files for the loader
```

If `INDEX.json` is present the cleaner will offer a sidebar listing each loaded
code; otherwise the cleaner falls back to whatever pattern matches it can find
inside the workbook (e.g. `IRC:36-2010`, `IRC SP:48`, etc.).

Cell-level IRC mentions discovered during cleaning are surfaced in the
Markdown / JSON export so the downstream analyzer (Ver63 / Ver64) can resolve
them against your local IRC library.

Format of `INDEX.json`:

```json
{
  "version": "64",
  "files": [
    { "code": "IRC:36-2010", "file": "IRC-36-2010.md", "title": "..." },
    { "code": "IRC:37-2018", "file": "IRC-37-2018.md", "title": "..." }
  ]
}
```

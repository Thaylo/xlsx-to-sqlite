---
name: xlsx-to-sqlite
description: >
  Convert Excel .xlsx files into queryable SQLite databases via streaming —
  works on files far too large for Excel, pandas, or openpyxl to open (hundreds
  of MB to multi-GB), with flat memory use and zero dependencies. Use this
  whenever the user wants a spreadsheet turned into a database, wants to run
  SQL over spreadsheet data, complains that an Excel file is huge, slow,
  crashing, or "won't open", or asks to extract/analyze data trapped in a big
  .xlsx dump — even if they never say the words "SQLite" or "database".
---

# xlsx → SQLite, at any size

A .xlsx file is a zip of XML. The worksheet XML is typically 3–10x larger than
the file on disk, which is why "just open it" fails for big dumps: Excel, pandas
and openpyxl (default mode) all try to hold that XML in memory as objects. The
bundled script instead streams the XML and writes SQLite in batches — memory
stays flat whether the file has 10 thousand rows or 10 million (validated on a
real-world 800k+ row, 1.4 GB export: ~9,000 rows/s, constant RAM).

## Workflow

### 1. Look before you convert

```bash
unzip -l file.xlsx | sort -k1 -rn | head        # sheet XML sizes = real data size
python3 scripts/xlsx_to_sqlite.py file.xlsx --peek
```

`--peek` lists every sheet and prints its first 5 rows. Use it to answer three
questions that decide the flags you'll pass:

- **Where are the headers?** Real dumps often have title/junk rows first. If
  the header is (say) the 3rd non-empty row, pass `--header-row 3`. If there is
  no header at all, pass `--no-header`.
- **Which sheets matter?** Convert everything by default; `--sheet NAME`
  (repeatable) to cherry-pick.
- **Does the output fit on disk?** Expect the SQLite file to be roughly the
  size of the *uncompressed* sheet XML shown by `unzip -l` (≈10x the .xlsx
  size). The script checks free space and refuses rather than fill the disk —
  a full disk mid-write loses the work *and* can destabilize the machine. If
  space is tight, free some first; suggest candidates (caches, node_modules,
  docker) but let the user decide what dies.

### 2. Convert

```bash
python3 scripts/xlsx_to_sqlite.py file.xlsx            # all sheets -> file.sqlite
python3 scripts/xlsx_to_sqlite.py file.xlsx -o out.sqlite --sheet "Vendas 2024"
python3 scripts/xlsx_to_sqlite.py "https://…/big.xlsx" -o out.sqlite   # remote
```

**Remote files convert while downloading.** Pass an http(s) URL (Google
Drive/Docs share links are auto-resolved, including the big-file virus-scan
confirmation) and the converter streams it with HTTP Range requests: the .xlsx
never touches local disk — only the SQLite output does — so a 2 GB export
needs ~2 GB of disk, not ~4. Each 8 MB block is an independent request retried
with backoff, so flaky connections cost one block, not the whole download.
`--peek` on a URL is nearly free (a few MB), so still peek first. Total bytes
transferred ≈ file size — the wins are disk, overlap, and retry granularity,
not bandwidth. Servers without Range support fail fast as `E06`; private
Drive files as `E07` (download those manually, then convert the local copy).

The script needs only the Python standard library — never install pandas or
openpyxl for this task; a 2 GB file would take 30+ minutes and gigabytes of RAM
that way. What it does for you:

- one table per sheet, names sanitized to safe SQL identifiers
  (`"Preço Unitário (R$)"` → `preco_unitario_r`, duplicates deduped)
- Excel serial dates → ISO 8601 text (`45366` → `2024-03-15`), detected from
  the workbook's styles, honoring the 1904 epoch on Mac-origin files
- sparse rows handled by cell reference, so a missing cell becomes NULL in the
  *right column* — values never shift left
- column affinities (INTEGER/REAL/TEXT) sniffed from the first 2000 rows
- resilient to quirks: sharedStrings or inline strings, missing `r=`
  attributes, rows wider than the header (columns are added on the fly)

It refuses to overwrite an existing output (`--force` to override) and prints
progress with a rows/s rate — a long-running conversion should be run in the
background, then verified when it reports DONE.

### 3. Verify — the numbers, not the vibes

The script prints per-table row counts (cross-checked against the DB) and
sample rows. Before declaring success, confirm against the source of truth:

```bash
sqlite3 out.sqlite "SELECT COUNT(*) FROM tablename"
```

The count must equal the sheet's data rows (the `<dimension ref="A1:H500001">`
attribute visible in `--peek`'s size line, minus header). Then eyeball 2–3
sample rows against what `--peek` showed — especially date columns (should read
`2024-03-15`, not `45366`) and columns that were sparse.

### 4. Index for the queries that will come

The conversion doesn't create indexes (pointless write cost if nobody filters).
Look at the columns and index the obvious filter/join candidates — dates,
categories, foreign-key-ish ids, low-cardinality dimensions:

```bash
sqlite3 out.sqlite 'CREATE INDEX idx_date ON vendas(date); ANALYZE;'
```

### 5. Hand off

Tell the user how to use the result — path, size, table names, row counts, and
one working example query. Suggest `sqlite3` CLI or DB Browser for SQLite for
GUIs, and warn that `SELECT *` without `LIMIT` on a table with big text columns
re-creates the original problem.

## When something looks wrong

The converter emits stable diagnostic codes — `WARN [W01..W07]` for structural
anomalies (it still completes) and `ERROR [E01..E05]` for hard failures. Treat
any WARN as a to-do, not noise: the most common cause is a sheet holding
several logical tables (stacked blocks, per-block headers, side-by-side
ranges, title rows), which converts *physically* fine into one table that is
*logically* wrong. Read `references/error-codes.md` for the full code table
with a recovery playbook per code (splitting blocks with SQL, `--header-row`,
`--no-header` raw imports, dropping artifact columns) — and relay the code to
the user when reporting.

Read `references/xlsx-internals.md` for the file-format details (cell types,
date serials, sharedStrings, namespaces) before hand-rolling any fix — most
"corrupt" files are just a quirk the reference explains. Non-.xlsx inputs are
out of scope for the script: `.xls` (old binary) and `.xlsb` need a converter
first (`ssconvert`, LibreOffice `soffice --headless --convert-to xlsx`), and
`.csv` goes straight into sqlite3's `.import`.

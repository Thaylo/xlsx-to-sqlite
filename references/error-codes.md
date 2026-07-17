# Diagnostic codes — what they mean and how to recover

The converter never fails silently on messy layouts: it emits `WARN [Wnn]`
lines for structural anomalies (conversion still completes) and `ERROR [Enn]`
for hard failures (exit code ≠ 0). Every code is stable and greppable —
`grep -E 'WARN \[W|ERROR \[E'` on the output gives you the full picture.

The single most dangerous input is a **sheet that holds several logical tables**
(stacked blocks separated by blank rows, per-block headers, side-by-side
ranges, title rows, subtotal footers). "One table per sheet" is the physical
truth of the conversion — these codes tell you when the *logical* truth
differs, so you can split afterwards instead of shipping a garbage table.

## Errors (conversion refused or aborted)

| Code | Meaning | Recovery |
|------|---------|----------|
| E01 | Input missing, not a zip, or missing workbook parts | Confirm the path. `.xls`/`.xlsb` are different binary formats: convert first with `soffice --headless --convert-to xlsx file.xls` (LibreOffice) or `ssconvert` (Gnumeric). A `.csv` renamed to `.xlsx` also lands here — import CSVs with sqlite3's `.import` instead. |
| E02 | `--sheet NAME` not found | The message lists available sheet names — they are case- and space-sensitive; quote them. |
| E03 | Not enough free disk (output ≈ uncompressed XML size × 1.2) | Free disk space first. `--ignore-space` overrides only when you know better (e.g. output on another volume). |
| E04 | Output file already exists | Pick another `-o` path, or `--force` to overwrite deliberately. |
| E05 | Sheet has no rows (reported as WARN; sheet skipped, run continues) | Nothing to do — but if you expected data, check you're looking at the right sheet. |

## Warnings (conversion completed — inspect before trusting)

| Code | Signal | Likely cause | Playbook |
|------|--------|--------------|----------|
| W01 | Blank gaps of >3 rows inside the data (boundary rows are listed) | **Stacked tables** in one sheet | Confirm with `--peek`. Import stays one table; split with SQL using the reported boundaries: rows are inserted in sheet order, so `CREATE TABLE block2 AS SELECT * FROM t WHERE rowid > N`. If blocks have different schemas, re-run with `--no-header` and split on the `col_N` raw table instead. |
| W02 | The header row's values reappear as a data row (rows listed) | Stacked tables with **per-block headers**, or a repeated header every N rows (print layouts) | The repeated header rows were imported as data — delete them (`DELETE FROM t WHERE col1 = 'Header1' AND col2 = 'Header2'`) after splitting blocks per W01. Their positions mark block boundaries. |
| W03 | Data rows extend beyond the header width; `col_N` columns were added | **Side-by-side tables**, or stray notes right of the range | Inspect the extra columns: `SELECT DISTINCT col_9, col_10 FROM t LIMIT 20`. A second logical table → copy those columns out into their own table and drop them; stray junk → just drop them. |
| W04 | Gaps inside the header row itself | Merged title cells, or a header that skips columns | Verify with `--peek` that the chosen row really is the header; a better one may sit lower (`--header-row N`). Gap columns got `col_N` names — rename with `ALTER TABLE t RENAME COLUMN col_3 TO better_name`. |
| W05 | Columns >99% empty (names listed) | Layout artifacts: spacer columns, a lone comment far right | Usually safe to `ALTER TABLE t DROP COLUMN x` — but check the few non-null values first: `SELECT rowid, x FROM t WHERE x IS NOT NULL LIMIT 5`; a subtotal or footnote hiding there can matter. |
| W06 | Declared sheet dimension disagrees with imported row count | Trailing blank rows, or a stale/lying `dimension` attribute | Trust the imported count. Cross-check against the source: `--peek` shows the dimension; if the source system reports an expected row count, compare against that, not the xlsx metadata. |
| W07 | First row has a single cell but data rows are wide | A **title row** above the real header — the title became the only column name | Re-run with `--header-row 2` (or higher — `--peek` shows where the real header sits). |

## Rules of thumb for agents

- Warnings are not failures: row-level data is intact; only the *shape* may be
  wrong. Never re-run blindly — inspect with `--peek` and SQL first.
- W01 + W02 together is near-certain stacked tables: split before analysis,
  or aggregates will silently mix blocks.
- When the layout is too chaotic for header inference, `--no-header` gives a
  faithful raw import (`col_1..col_N`, every row) — SQL is a better scalpel
  than re-parsing XML.
- Quote the code (e.g. "W02") when reporting to the user; it makes the issue
  searchable in this table.

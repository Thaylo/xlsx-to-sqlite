# xlsx-to-sqlite

An [Agent Skill](https://agentskills.io) that lets AI coding agents convert
huge Excel dumps — hundreds of MB to multiple GB — into queryable SQLite
databases, with flat memory use and zero dependencies beyond the Python
standard library.

A `.xlsx` file is a zip of XML, and the worksheet XML is typically 3–10x the
file's size on disk. That's why big exports crash Excel, hang pandas, and OOM
openpyxl: they all try to hold that XML in memory. The converter here streams
it instead, so a 10-million-row file costs the same RAM as a 10-thousand-row
one. Validated on a real-world 800k+ row / 1.4 GB single-sheet export:
~9,000 rows/s at constant memory.

## For AI agents (the skill)

Install into any agent runtime that supports Agent Skills, e.g. Claude Code:

```bash
git clone https://github.com/Thaylo/xlsx-to-sqlite ~/.claude/skills/xlsx-to-sqlite
```

From then on, prompts like *"this excel export won't open, turn it into
something I can query"* trigger a workflow that inspects the file first
(`--peek`), checks free disk before writing gigabytes, converts via the
streaming script, verifies row counts against the source, and adds indexes for
likely queries. `SKILL.md` is the agent-facing playbook;
`references/xlsx-internals.md` documents the file-format traps (sparse rows,
date serials, sharedStrings, the 1904 epoch) for when a file misbehaves.

## For humans (the script standalone)

No agent required — it's a plain CLI:

```bash
python3 scripts/xlsx_to_sqlite.py dump.xlsx --peek          # look before you leap
python3 scripts/xlsx_to_sqlite.py dump.xlsx                 # all sheets -> dump.sqlite
python3 scripts/xlsx_to_sqlite.py dump.xlsx --sheet "Sales" -o sales.sqlite
```

What you get:

- one table per sheet, column names sanitized to SQL-safe identifiers
  (`"Unit Price (R$)"` → `unit_price_r`, duplicates deduped, accents stripped)
- Excel date serials converted to ISO 8601 text (`45366` → `2024-03-15`),
  detected from the workbook's styles, including the Mac 1904-epoch variant
- sparse rows mapped by cell reference — missing cells become NULL in the
  right column, values never shift left
- INTEGER/REAL/TEXT affinities sniffed from the data
- resilience to writer quirks: inline strings or sharedStrings, missing `r=`
  attributes, rows wider than the header, phonetic (furigana) runs
- a free-disk check before writing (the output is roughly the size of the
  uncompressed XML — about 10x the .xlsx), refusal to overwrite without
  `--force`, progress with rows/s, and a verification summary with sample rows

Useful flags: `--header-row N` (junk rows above the header), `--no-header`,
`--sheet NAME` (repeatable), `--batch N`, `--ignore-space`.

## Tests

```bash
python3 tests/test_conversion.py
```

Generates synthetic fixtures (inline-string dump with sparse cells; multi-sheet
workbook with sharedStrings, date serials and booleans) and asserts row counts,
NULL placement, date conversion, and identifier sanitization.

## License

MIT — see [LICENSE](LICENSE).

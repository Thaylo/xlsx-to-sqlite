# xlsx-to-sqlite

Convert huge Excel dumps — hundreds of MB to multiple GB — into queryable
SQLite databases. Works both as an [Agent Skill](https://agentskills.io) for AI
coding agents and as a standalone CLI. No dependencies: Python standard
library only.

A `.xlsx` is a zip of XML that expands to 3–10x its size on disk — that's why
big exports crash Excel and OOM pandas/openpyxl, which hold it all in memory.
This converter streams it instead: a 10-million-row file costs the same RAM as
a 10-thousand-row one. Validated on a real-world 800k+ row / 1.4 GB export at
~9,000 rows/s with constant memory.

## Requirements

Python 3.8+ — nothing else.

- **Debian/Ubuntu**: usually preinstalled; otherwise `sudo apt install python3`
- **macOS**: preinstalled (or `brew install python`)
- **Windows 10+**: `winget install Python.Python.3.12` (or the Microsoft
  Store), then use `py` wherever the examples say `python3`

## Quick start (CLI)

```bash
git clone https://github.com/Thaylo/xlsx-to-sqlite
cd xlsx-to-sqlite

python3 scripts/xlsx_to_sqlite.py dump.xlsx --peek   # inspect sheets + first rows
python3 scripts/xlsx_to_sqlite.py dump.xlsx          # convert -> dump.sqlite
python3 scripts/xlsx_to_sqlite.py dump.xlsx --sheet "Sales" -o sales.sqlite

# remote files convert WHILE downloading (HTTP Range streaming) — the .xlsx
# never touches your disk; Google Drive share links are auto-resolved:
python3 scripts/xlsx_to_sqlite.py "https://drive.google.com/file/d/FILE_ID/view" -o data.sqlite
```

Then query with the `sqlite3` CLI or any GUI (e.g.
[DB Browser for SQLite](https://sqlitebrowser.org/) — Linux/macOS/Windows).

Useful flags: `--header-row N` (junk rows above the header), `--no-header`,
`--sheet NAME` (repeatable), `--force` (overwrite), `--ignore-space` (skip the
free-disk check).

## Install as an Agent Skill

Give the skill to an agent runtime that supports Agent Skills (e.g. Claude
Code) by cloning into its skills folder:

```bash
# Linux / macOS
git clone https://github.com/Thaylo/xlsx-to-sqlite ~/.claude/skills/xlsx-to-sqlite

# Windows (PowerShell)
git clone https://github.com/Thaylo/xlsx-to-sqlite "$env:USERPROFILE\.claude\skills\xlsx-to-sqlite"
```

From then on, prompts like *"this excel export won't open, turn it into
something I can query"* make the agent inspect the file, check free disk,
stream-convert, verify row counts against the source, and add useful indexes.
`SKILL.md` is the agent-facing playbook; `references/xlsx-internals.md`
documents the file-format traps for when a file misbehaves.

## What you get

- one table per sheet; column names sanitized to SQL-safe identifiers
  (`"Unit Price ($)"` → `unit_price`, duplicates deduped, accents stripped)
- Excel date serials → ISO 8601 text (`45366` → `2024-03-15`), detected from
  the workbook styles, including the Mac 1904-epoch variant
- sparse rows mapped by cell reference: missing cells become NULL in the right
  column — values never shift left
- INTEGER/REAL/TEXT column affinities sniffed from the data
- handles sharedStrings and inline strings, missing `r=` attributes, rows
  wider than the header, phonetic (furigana) runs
- free-disk check before writing (output ≈ 10x the .xlsx size), refusal to
  overwrite without `--force`, progress with rows/s, verification summary
- structural diagnostics with stable codes — stacked tables in one sheet,
  repeated headers, title rows, layout artifacts — each with a recovery
  playbook in [references/error-codes.md](references/error-codes.md)

## Tests

```bash
python3 tests/test_conversion.py
```

## License

MIT — see [LICENSE](LICENSE).

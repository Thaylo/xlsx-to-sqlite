# xlsx internals — what you need when a file misbehaves

A `.xlsx` is a zip. The members that matter:

| member | role |
|---|---|
| `xl/workbook.xml` | sheet names, order, `date1904` flag |
| `xl/_rels/workbook.xml.rels` | maps each sheet's `r:id` to its XML file |
| `xl/worksheets/sheetN.xml` | the actual cells (this is the big one) |
| `xl/sharedStrings.xml` | every distinct string, referenced by index |
| `xl/styles.xml` | number formats — the only way to know a cell is a date |

All elements live in the namespace
`http://schemas.openxmlformats.org/spreadsheetml/2006/main`.

## Cells

```xml
<row r="2">
  <c r="A2"><v>42</v></c>                          <!-- number (or date!) -->
  <c r="B2" t="s"><v>17</v></c>                    <!-- sharedStrings[17] -->
  <c r="C2" t="inlineStr"><is><t>text</t></is></c> <!-- literal string -->
  <c r="D2" t="b"><v>1</v></c>                     <!-- boolean -->
  <c r="E2" t="str"><v>result</v></c>              <!-- formula's cached string -->
  <c r="F2" t="e"><v>#DIV/0!</v></c>               <!-- error value -->
</row>
```

Traps, in descending order of how often they bite:

- **Sparse rows.** Empty cells are simply absent. Position comes from the
  `r="C2"` reference, never from element order — mapping cells to columns by
  order silently shifts values left past every gap.
- **Dates are numbers.** `45366` is a date if and only if the cell's style
  (`s="N"` → `styles.xml` `cellXfs[N]` → `numFmtId`) is a date format. Builtin
  date ids: 14–22, 27–36, 45–47, 50–58; custom formats are date-ish when the
  code contains d/m/y/h/s outside quoted/bracketed sections. Serial → date:
  `datetime(1899,12,30) + timedelta(days=serial)` — or `datetime(1904,1,1)`
  when `<workbookPr date1904="1"/>` (files from old Mac Excel).
- **Missing `r=` attributes.** Some writers (streaming exporters) omit cell and
  row references; the reader must then count positions itself.
- **Rich text.** A sharedStrings entry or inline string can be split across
  many `<r><t>…</t></r>` runs — concatenate every `<t>`, but skip `<rPh>`
  (phonetic furigana) subtrees or Japanese files get polluted text.
- **`t="str"` vs `t="s"`.** `str` means "formula whose cached result is a
  string" — the value is literal, not a sharedStrings index. Confusing them
  turns text into garbage numbers or IndexErrors.
- **Huge sheets have no sharedStrings.** Machine-generated dumps often use
  `inlineStr` everywhere (that's what a 1.4 GB single-sheet dump looks like).
  Conversely, hand-made files put nearly everything through sharedStrings —
  which must be loaded fully before reading any sheet.
- **`dimension` lies sometimes.** `<dimension ref="A1:H500001"/>` is a hint,
  not a guarantee; trust the actual rows you streamed, and use dimension only
  as a cross-check.

## Memory discipline for iterparse

`ET.iterparse` keeps every parsed element attached to the root until you drop
it. The pattern that keeps memory flat on multi-GB XML:

```python
for event, elem in ET.iterparse(stream, events=("start", "end")):
    if event == "start":
        root = root or elem
        continue
    if elem.tag == ROW:
        process(elem)
        elem.clear()
        root.clear()      # detach processed rows from the tree
```

Forgetting `root.clear()` looks fine in testing and OOMs in production — the
cleared elements themselves still accumulate as (empty) children of the root.

## SQLite loading speed

Per-row INSERTs with default journaling are ~100x too slow for millions of
rows. During a bulk load into a fresh file: `PRAGMA journal_mode=OFF`,
`PRAGMA synchronous=OFF`, batched `executemany` (10k rows) with one commit per
batch. Crash-safety doesn't matter — the source file still exists; rerun.
Restore `journal_mode=WAL` at the end so readers get a well-behaved database.

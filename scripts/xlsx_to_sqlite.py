#!/usr/bin/env python3
"""Stream large .xlsx files into SQLite with flat memory. Stdlib only.

Usage:
  python3 xlsx_to_sqlite.py input.xlsx                 # convert all sheets
  python3 xlsx_to_sqlite.py input.xlsx --peek          # inspect before converting
  python3 xlsx_to_sqlite.py input.xlsx -o out.sqlite --sheet "Sales 2024"

Why streaming: a .xlsx is a zip of XML; the worksheet XML is often 3-10x the
file size once decompressed. Loading it whole (pandas/openpyxl default mode)
needs RAM proportional to that. This script parses the XML as a stream and
inserts in batches, so memory stays flat regardless of file size.

Diagnostics: structural anomalies (stacked tables, repeated headers, layout
artifacts) are reported as `WARN [Wnn]` lines and hard failures as
`ERROR [Enn]` — the code table with recovery playbooks lives in
references/error-codes.md.
"""
import argparse
import os
import posixpath
import re
import shutil
import sqlite3
import sys
import time
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PNS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Built-in Excel number-format ids that render as dates/times.
BUILTIN_DATE_FMTS = (
    set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))
)
SNIFF_ROWS = 2000
GAP_ROWS = 3  # this many consecutive blank rows inside data smells like stacked tables


# ---------- workbook metadata ----------

def discover_sheets(zf):
    """Return [(sheet_name, zip_member_path)] in workbook order."""
    rels = {}
    with zf.open("xl/_rels/workbook.xml.rels") as f:
        for rel in ET.parse(f).getroot():
            target = rel.get("Target", "").lstrip("/")
            if not target.startswith("xl/"):
                target = posixpath.normpath(posixpath.join("xl", target))
            rels[rel.get("Id")] = target
    sheets = []
    with zf.open("xl/workbook.xml") as f:
        root = ET.parse(f).getroot()
        for sh in root.iter(NS + "sheet"):
            rid = sh.get(RNS + "id")
            if rid in rels:
                sheets.append((sh.get("name"), rels[rid]))
    return sheets


def uses_1904_epoch(zf):
    with zf.open("xl/workbook.xml") as f:
        pr = ET.parse(f).getroot().find(NS + "workbookPr")
    return pr is not None and pr.get("date1904") in ("1", "true")


def load_shared_strings(zf):
    """sharedStrings.xml holds every distinct string; cells reference by index."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    strings, root = [], None
    with zf.open("xl/sharedStrings.xml") as f:
        for event, si in ET.iterparse(f, events=("start", "end")):
            if event == "start":
                if root is None:
                    root = si
                continue
            if si.tag != NS + "si":
                continue
            # drop phonetic (furigana) runs so they don't pollute the text
            for junk in si.findall(NS + "rPh") + si.findall(NS + "phoneticPr"):
                si.remove(junk)
            strings.append("".join(t.text or "" for t in si.iter(NS + "t")))
            si.clear()
            if len(strings) % 100_000 == 0:
                root.clear()
    return strings


def _is_date_code(code):
    if re.search(r"\[(h+|m+|s+)\]", code, re.I):  # elapsed time like [h]:mm
        return True
    stripped = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code)
    return bool(re.search(r"[dmhys]", stripped, re.I))


def load_date_styles(zf):
    """Set of cell style indexes (the s= attribute) whose format is a date."""
    if "xl/styles.xml" not in zf.namelist():
        return set()
    with zf.open("xl/styles.xml") as f:
        root = ET.parse(f).getroot()
    custom = {}
    for nf in root.iter(NS + "numFmt"):
        custom[int(nf.get("numFmtId"))] = nf.get("formatCode", "")
    date_styles = set()
    cellxfs = root.find(NS + "cellXfs")
    if cellxfs is None:
        return date_styles
    for i, xf in enumerate(cellxfs.findall(NS + "xf")):
        fmt = int(xf.get("numFmtId", "0"))
        if fmt in BUILTIN_DATE_FMTS or (fmt in custom and _is_date_code(custom[fmt])):
            date_styles.add(str(i))
    return date_styles


# ---------- cell decoding ----------

def col_index(ref):
    idx = 0
    for ch in ref:
        if ch.isdigit():
            break
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def serial_to_text(num, epoch):
    """Excel stores dates as day-counts from an epoch; emit ISO 8601 text."""
    try:
        days = int(num)
        secs = round((num - days) * 86400)
        dt = epoch + timedelta(days=days, seconds=secs)
    except (OverflowError, ValueError):
        return num
    if 0 <= num < 1:
        return dt.strftime("%H:%M:%S")
    if dt.hour == dt.minute == dt.second == 0:
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def cell_value(c, shared, date_styles, epoch):
    t = c.get("t")
    if t == "s":
        v = c.find(NS + "v")
        return shared[int(v.text)] if v is not None and v.text else None
    if t == "inlineStr":
        is_el = c.find(NS + "is")
        if is_el is None:
            return None
        for junk in is_el.findall(NS + "rPh") + is_el.findall(NS + "phoneticPr"):
            is_el.remove(junk)
        return "".join(node.text or "" for node in is_el.iter(NS + "t")) or None
    v = c.find(NS + "v")
    if v is None or v.text is None:
        return None
    if t in ("str", "e"):
        return v.text
    if t == "b":
        return int(v.text)
    s = v.text
    try:
        num = int(s)
    except ValueError:
        try:
            num = float(s)
        except ValueError:
            return s
    if c.get("s") in date_styles:
        return serial_to_text(num, epoch)
    return num


def iter_rows(zf, member, shared, date_styles, epoch, meta=None):
    """Yield (row_number, {col_index: value}) per row, streaming. Never trust
    cell order alone: sparse rows omit empty cells, so positions come from r=."""
    with zf.open(member) as stream:
        root, rownum = None, 0
        for event, elem in ET.iterparse(stream, events=("start", "end")):
            if event == "start":
                if root is None:
                    root = elem
                continue
            if elem.tag == NS + "dimension" and meta is not None:
                meta["dimension"] = elem.get("ref", "")
                continue
            if elem.tag != NS + "row":
                continue
            r = elem.get("r")
            rownum = int(r) if r and r.isdigit() else rownum + 1
            cells, last = {}, -1
            for c in elem.iter(NS + "c"):
                ref = c.get("r")
                idx = col_index(ref) if ref else last + 1  # some writers omit r=
                last = idx
                val = cell_value(c, shared, date_styles, epoch)
                if val is not None:
                    cells[idx] = val
            yield rownum, cells
            elem.clear()
            root.clear()


# ---------- naming and typing ----------

def sanitize(name, i, seen):
    s = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^0-9a-zA-Z]+", "_", s).strip("_").lower() or f"col_{i + 1}"
    if s[0].isdigit():
        s = "c_" + s
    base, n = s, 2
    while s in seen:
        s, n = f"{base}_{n}", n + 1
    seen.add(s)
    return s


def sniff_types(rows, ncols):
    kinds = [set() for _ in range(ncols)]
    for _, row in rows:
        for i, v in row.items():
            if i < ncols and v is not None:
                kinds[i].add(float if isinstance(v, float) else type(v))
    out = []
    for k in kinds:
        if str in k or not k:
            out.append("TEXT")
        elif float in k:
            out.append("REAL")
        else:
            out.append("INTEGER")
    return out


# ---------- conversion ----------

def convert_sheet(conn, zf, sheet_name, member, shared, date_styles, epoch, opts,
                  taken_tables, diagnostics):
    table = sanitize(sheet_name, 0, taken_tables)
    meta = {}
    rows = iter_rows(zf, member, shared, date_styles, epoch, meta)

    def diag(code, msg):
        line = f"WARN [{code}] sheet '{sheet_name}': {msg}"
        diagnostics.append(line)
        print("  " + line, flush=True)

    headers_src, buffer = None, []
    skipped_before_header = seen_rows = header_rownum = 0
    for rownum, cells in rows:
        if not cells:
            continue
        seen_rows += 1
        if opts.no_header:
            headers_src = {}
            buffer.append((rownum, cells))
            break
        if seen_rows < opts.header_row:
            skipped_before_header += 1
            continue
        headers_src = cells
        header_rownum = rownum
        break
    else:
        print(f"  WARN [E05] sheet '{sheet_name}': no rows, skipped", flush=True)
        return None

    # buffer a sniff window to size the table and pick column affinities
    for rc in rows:
        buffer.append(rc)
        if len(buffer) >= SNIFF_ROWS:
            break
    width = max(
        ([max(headers_src, default=-1) + 1] if headers_src else [0])
        + [max(c, default=-1) + 1 for _, c in buffer]
    )

    # --- structural sanity checks on the header (codes: references/error-codes.md)
    if headers_src and not opts.no_header:
        span = max(headers_src) - min(headers_src) + 1
        if len(headers_src) < span:
            diag("W04", f"header row {header_rownum} has {span - len(headers_src)} "
                        "gap(s) inside it — merged cells or title layout; the gap "
                        "columns got col_N names")
        if opts.header_row == 1 and len(headers_src) == 1 and buffer:
            avg = sum(len(c) for _, c in buffer[:5]) / len(buffer[:5])
            if avg >= 3:
                only = next(iter(headers_src.values()))
                diag("W07", f"first non-empty row has a single cell "
                            f"({str(only)[:40]!r}) but data rows have ~{avg:.0f} "
                            "cells — likely a title above the real header; "
                            "re-run with --header-row 2 (check with --peek)")

    seen_cols = set()
    if opts.no_header:
        headers = [sanitize(None, i, seen_cols) for i in range(width)]
    else:
        headers = [sanitize(headers_src.get(i), i, seen_cols) for i in range(width)]
    types = sniff_types(buffer, width)
    cols_sql = ", ".join(f'"{h}" {t}' for h, t in zip(headers, types))
    conn.execute(f'CREATE TABLE "{table}" ({cols_sql})')
    insert = f'INSERT INTO "{table}" VALUES ({",".join("?" * width)})'

    t0, total, batch, last_report = time.time(), 0, [], time.time()
    counts = {}          # per-column non-null tally, for W05
    gaps, repeats = [], []
    scan_last = header_rownum or None
    header_sig = (frozenset(headers_src.values()) or None) if headers_src else None
    first_hkey = min(headers_src) if headers_src else None

    def scan(rownum, cells):
        """Track stacked-table smells: blank-row gaps and repeated header rows."""
        nonlocal scan_last
        if scan_last is not None and rownum - scan_last > GAP_ROWS and len(gaps) < 10:
            gaps.append((scan_last, rownum))
        scan_last = rownum
        if (header_sig and cells.get(first_hkey) == headers_src[first_hkey]
                and len(cells) == len(headers_src)
                and frozenset(cells.values()) == header_sig and len(repeats) < 10):
            repeats.append(rownum)

    def flush():
        nonlocal total
        if batch:
            conn.executemany(insert, batch)
            conn.commit()
            total += len(batch)
            batch.clear()

    def widen(new_width):
        nonlocal width, insert
        diag("W03", f"data extends beyond the header ({width} -> {new_width} "
                    "columns) — extra col_N columns added; side-by-side table "
                    "or stray cells right of the range?")
        for i in range(width, new_width):
            headers.append(sanitize(None, i, seen_cols))
            conn.execute(f'ALTER TABLE "{table}" ADD COLUMN "{headers[i]}"')
        width = new_width
        insert = f'INSERT INTO "{table}" VALUES ({",".join("?" * width)})'

    def push(rownum, cells):
        nonlocal last_report
        if not cells:
            return
        scan(rownum, cells)
        for i in cells:
            counts[i] = counts.get(i, 0) + 1
        if max(cells) >= width:
            flush()
            widen(max(cells) + 1)
        batch.append(tuple(cells.get(i) for i in range(width)))
        if len(batch) >= opts.batch:
            flush()
            if time.time() - last_report > 5:
                el = time.time() - t0
                print(f"  {table}: {total:,} rows  {el:.0f}s  ({total / el:,.0f} rows/s)", flush=True)
                last_report = time.time()

    for rownum, cells in buffer:
        push(rownum, cells)
    for rownum, cells in rows:
        push(rownum, cells)
    flush()

    # --- post-load structural diagnostics
    if gaps:
        shown = ", ".join(f"{a}->{b}" for a, b in gaps[:3])
        diag("W01", f"{len(gaps)} blank gap(s) of >{GAP_ROWS} rows inside the data "
                    f"(e.g. rows {shown}) — the sheet may contain stacked tables")
    if repeats:
        diag("W02", f"the header row repeats at row(s) {repeats} — stacked tables "
                    "with per-block headers; those rows were imported as data")
    if total >= 1000:
        hollow = [headers[i] for i in range(width) if counts.get(i, 0) < total * 0.01]
        if hollow:
            diag("W05", f"column(s) {hollow} are >99% empty — layout artifacts? "
                        "consider dropping them")
    m = re.search(r":[A-Z]+(\d+)$", meta.get("dimension", ""))
    if m:
        declared = int(m.group(1)) - (header_rownum or 0)
        if declared > 0 and abs(declared - total) > max(2, declared * 0.01):
            diag("W06", f"sheet dimension declares ~{declared:,} data rows but "
                        f"{total:,} were imported — trailing blanks or a lying "
                        "dimension; trust the imported count, cross-check the source")

    dbcount = conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
    status = "OK" if dbcount == total else f"MISMATCH inserted={total}"
    print(f"  sheet '{sheet_name}' -> table {table}: {dbcount:,} rows "
          f"[{status}] in {time.time() - t0:.0f}s", flush=True)
    if skipped_before_header:
        print(f"  note: skipped {skipped_before_header} row(s) above the header", flush=True)
    return table, dbcount, headers


def peek(zf, sheets, shared, date_styles, epoch, n):
    for name, member in sheets:
        size = zf.getinfo(member).file_size
        meta, shown = {}, 0
        print(f"\n=== sheet '{name}' ({size / 1e6:,.0f} MB uncompressed XML)")
        for rownum, cells in iter_rows(zf, member, shared, date_styles, epoch, meta):
            if shown >= n:
                break
            shown += 1
            width = max(cells, default=-1) + 1
            vals = [str(cells.get(j, ""))[:28] for j in range(min(width, 12))]
            print(f"  row {rownum}: {vals}{' …' if width > 12 else ''}")
        if meta.get("dimension"):
            print(f"  dimension: {meta['dimension']}")


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("xlsx")
    p.add_argument("-o", "--output")
    p.add_argument("--sheet", action="append", help="convert only this sheet (repeatable)")
    p.add_argument("--peek", action="store_true", help="print sheets + first rows, no conversion")
    p.add_argument("--no-header", action="store_true", help="first row is data; name columns col_N")
    p.add_argument("--header-row", type=int, default=1,
                   help="1-based index (among non-empty rows) of the header row; rows above are skipped")
    p.add_argument("--batch", type=int, default=10_000)
    p.add_argument("--sample", type=int, default=3, help="sample rows to print per table at the end")
    p.add_argument("--force", action="store_true", help="overwrite an existing output file")
    p.add_argument("--ignore-space", action="store_true", help="skip the free-disk check")
    opts = p.parse_args()

    try:
        zf = zipfile.ZipFile(opts.xlsx)
        sheets = discover_sheets(zf)
    except FileNotFoundError:
        sys.exit(f"ERROR [E01] {opts.xlsx}: file not found")
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as e:
        sys.exit(f"ERROR [E01] {opts.xlsx}: not a readable .xlsx ({e!r}); .xls/.xlsb "
                 "need conversion first — see references/error-codes.md")
    if opts.sheet:
        wanted = set(opts.sheet)
        sheets = [s for s in sheets if s[0] in wanted]
        missing = wanted - {s[0] for s in sheets}
        if missing:
            sys.exit(f"ERROR [E02] sheet(s) not found: {sorted(missing)}; available: "
                     f"{[s[0] for s in discover_sheets(zf)]}")

    epoch = datetime(1904, 1, 1) if uses_1904_epoch(zf) else datetime(1899, 12, 30)
    date_styles = load_date_styles(zf)
    shared = load_shared_strings(zf)

    if opts.peek:
        peek(zf, sheets, shared, date_styles, epoch, 5)
        return

    out = opts.output or os.path.splitext(opts.xlsx)[0] + ".sqlite"
    if os.path.exists(out) and not opts.force:
        sys.exit(f"ERROR [E04] refusing to overwrite {out} (use --force)")

    # The DB lands roughly the size of the uncompressed XML. Check before
    # writing gigabytes: running a disk to 0 mid-conversion hurts.
    est = sum(zf.getinfo(m).file_size for _, m in sheets)
    if "xl/sharedStrings.xml" in zf.namelist():
        est += 2 * zf.getinfo("xl/sharedStrings.xml").file_size
    free = shutil.disk_usage(os.path.dirname(os.path.abspath(out))).free
    print(f"estimated output ~{est / 1e9:.1f} GB, free disk {free / 1e9:.1f} GB")
    if not opts.ignore_space and free < est * 1.2 + 2e8:
        sys.exit("ERROR [E03] not enough free disk for a safe conversion "
                 "(need ~1.2x the uncompressed sheet size); free space or use --ignore-space")

    if os.path.exists(out):
        os.remove(out)
    conn = sqlite3.connect(out)
    conn.executescript(
        "PRAGMA journal_mode=OFF; PRAGMA synchronous=OFF; PRAGMA temp_store=MEMORY;"
    )
    results, taken, diagnostics = [], set(), []
    for name, member in sheets:
        r = convert_sheet(conn, zf, name, member, shared, date_styles, epoch, opts,
                          taken, diagnostics)
        if r:
            results.append(r)
    conn.execute("PRAGMA journal_mode=WAL")

    print(f"\nDONE -> {out} ({os.path.getsize(out) / 1e9:.2f} GB)")
    for table, count, headers in results:
        print(f"  {table}: {count:,} rows, columns: {headers}")
        for row in conn.execute(f'SELECT * FROM "{table}" LIMIT {opts.sample}'):
            print("    " + " | ".join(str(v)[:60] for v in row))
    if diagnostics:
        print(f"diagnostics: {len(diagnostics)} warning(s) — codes explained in "
              "references/error-codes.md")
    else:
        print("diagnostics: clean")
    conn.close()


if __name__ == "__main__":
    main()

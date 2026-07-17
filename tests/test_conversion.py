#!/usr/bin/env python3
"""End-to-end tests for scripts/xlsx_to_sqlite.py. Stdlib only.

Builds two synthetic workbooks exercising the format traps that corrupt data
silently (sparse cells, date serials, sharedStrings, dirty headers), converts
them, and asserts on the resulting databases.
"""
import http.server
import importlib.util
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import threading
import zipfile
from datetime import date, timedelta

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(REPO, "scripts", "xlsx_to_sqlite.py")
NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
RNS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
PKG = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
EPOCH = date(1899, 12, 30)
ROWS = 2000


def write_xlsx(path, sheet_names, sheet_xmls, shared_xml=None, styles_xml=None):
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(len(sheet_names)))
    if shared_xml:
        overrides += ('<Override PartName="/xl/sharedStrings.xml" ContentType='
                      '"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
    if styles_xml:
        overrides += ('<Override PartName="/xl/styles.xml" ContentType='
                      '"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    sheets = "".join(f'<sheet name="{n}" sheetId="{i+1}" r:id="rId{i+1}"/>'
                     for i, n in enumerate(sheet_names))
    rels = "".join(
        f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
        for i in range(len(sheet_names)))
    k = len(sheet_names)
    if shared_xml:
        rels += (f'<Relationship Id="rId{k+1}" Type="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')
    if styles_xml:
        rels += (f'<Relationship Id="rId{k+2}" Type="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",
                    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/'
                    'package/2006/content-types">'
                    '<Default Extension="rels" ContentType='
                    '"application/vnd.openxmlformats-package.relationships+xml"/>'
                    '<Default Extension="xml" ContentType="application/xml"/>'
                    '<Override PartName="/xl/workbook.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                    f'{overrides}</Types>')
        zf.writestr("_rels/.rels",
                    f'<?xml version="1.0"?><Relationships {PKG}><Relationship Id="rId1" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
                    'officeDocument" Target="xl/workbook.xml"/></Relationships>')
        zf.writestr("xl/workbook.xml",
                    f'<?xml version="1.0"?><workbook {NS} {RNS}><sheets>{sheets}</sheets></workbook>')
        zf.writestr("xl/_rels/workbook.xml.rels",
                    f'<?xml version="1.0"?><Relationships {PKG}>{rels}</Relationships>')
        for i, xml in enumerate(sheet_xmls):
            zf.writestr(f"xl/worksheets/sheet{i+1}.xml", xml)
        if shared_xml:
            zf.writestr("xl/sharedStrings.xml", shared_xml)
        if styles_xml:
            zf.writestr("xl/styles.xml", styles_xml)


def sheet(rows):
    return (f'<?xml version="1.0"?><worksheet {NS}>'
            f'<sheetData>{"".join(rows)}</sheetData></worksheet>')


def convert(xlsx, out, *extra):
    r = subprocess.run([sys.executable, SCRIPT, xlsx, "-o", out, "--force", *extra],
                       capture_output=True, text=True)
    assert r.returncode == 0, f"converter failed:\n{r.stdout}\n{r.stderr}"
    return sqlite3.connect(out), r.stdout


def test_inline_sparse_dirty_headers(tmp):
    """Inline strings, sparse Notes column, headers needing sanitization."""
    headers = ["Order ID", "Product", "Qty", "Unit Price ($)", "Résumé Notes"]
    rows = ['<row r="1">' + "".join(
        f'<c r="{chr(65+j)}1" t="inlineStr"><is><t>{h}</t></is></c>'
        for j, h in enumerate(headers)) + "</row>"]
    for i in range(1, ROWS + 1):
        r = i + 1
        cells = [f'<c r="A{r}"><v>{i}</v></c>',
                 f'<c r="B{r}" t="inlineStr"><is><t>item-{i}</t></is></c>',
                 f'<c r="C{r}"><v>{(i % 9) + 1}</v></c>',
                 f'<c r="D{r}"><v>{round(i * 0.07, 2)}</v></c>']
        if i % 2 == 1:  # sparse: even rows have no Notes cell at all
            cells.append(f'<c r="E{r}" t="inlineStr"><is><t>note {i}</t></is></c>')
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    xlsx = os.path.join(tmp, "orders.xlsx")
    write_xlsx(xlsx, ["Orders"], [sheet(rows)])
    conn, log = convert(xlsx, os.path.join(tmp, "orders.sqlite"))
    assert "diagnostics: clean" in log, log

    cols = [r[1] for r in conn.execute("PRAGMA table_info(orders)")]
    assert cols == ["order_id", "product", "qty", "unit_price", "resume_notes"], cols
    n = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    assert n == ROWS, n
    row = conn.execute("SELECT product, qty FROM orders WHERE order_id=777").fetchone()
    assert row == ("item-777", (777 % 9) + 1), row
    # the trap: even ids must have NULL notes, odd ids their own note — no shifting
    assert conn.execute("SELECT resume_notes FROM orders WHERE order_id=42").fetchone()[0] is None
    assert conn.execute("SELECT resume_notes FROM orders WHERE order_id=43").fetchone()[0] == "note 43"
    shifted = conn.execute(
        "SELECT COUNT(*) FROM orders WHERE resume_notes IS NOT NULL "
        "AND resume_notes != 'note ' || order_id").fetchone()[0]
    assert shifted == 0, f"{shifted} shifted note values"
    print("PASS inline strings + sparse cells + header sanitization")


def test_sharedstrings_dates_bools(tmp):
    """Two sheets, sharedStrings, date serials with a date style, booleans."""
    strings, index = [], {}

    def s(text):
        if text not in index:
            index[text] = len(strings)
            strings.append(text)
        return index[text]

    rows1 = ['<row r="1">' + "".join(
        f'<c r="{chr(65+j)}1" t="s"><v>{s(h)}</v></c>'
        for j, h in enumerate(["ID", "Name", "Signup Date", "Active"])) + "</row>"]
    for i in range(1, 501):
        r = i + 1
        serial = (date(2024, 1, 1) + timedelta(days=i % 365) - EPOCH).days
        rows1.append(f'<row r="{r}"><c r="A{r}"><v>{i}</v></c>'
                     f'<c r="B{r}" t="s"><v>{s(f"Person {i % 50}")}</v></c>'
                     f'<c r="C{r}" s="1"><v>{serial}</v></c>'
                     f'<c r="D{r}" t="b"><v>{1 if i % 3 == 0 else 0}</v></c></row>')
    rows2 = ['<row r="1">' + "".join(
        f'<c r="{chr(65+j)}1" t="s"><v>{s(h)}</v></c>'
        for j, h in enumerate(["ID", "Company"])) + "</row>"]
    for j in range(1, 101):
        r = j + 1
        rows2.append(f'<row r="{r}"><c r="A{r}"><v>{j}</v></c>'
                     f'<c r="B{r}" t="s"><v>{s(f"Company {j} Ltd")}</v></c></row>')

    sst = (f'<?xml version="1.0"?><sst {NS} count="{len(strings)}" '
           f'uniqueCount="{len(strings)}">'
           + "".join(f"<si><t>{t}</t></si>" for t in strings) + "</sst>")
    styles = (f'<?xml version="1.0"?><styleSheet {NS}>'
              '<cellXfs count="2"><xf numFmtId="0"/>'
              '<xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>')
    xlsx = os.path.join(tmp, "crm.xlsx")
    write_xlsx(xlsx, ["Contacts 2024", "Companies"], [sheet(rows1), sheet(rows2)],
               shared_xml=sst, styles_xml=styles)
    conn, _ = convert(xlsx, os.path.join(tmp, "crm.sqlite"))

    tables = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")}
    assert tables == {"contacts_2024", "companies"}, tables
    assert conn.execute("SELECT COUNT(*) FROM contacts_2024").fetchone()[0] == 500
    assert conn.execute("SELECT COUNT(*) FROM companies").fetchone()[0] == 100
    d = conn.execute("SELECT signup_date FROM contacts_2024 WHERE id=100").fetchone()[0]
    assert d == (date(2024, 1, 1) + timedelta(days=100)).isoformat(), d
    a99, a100 = (conn.execute(
        f"SELECT active FROM contacts_2024 WHERE id={i}").fetchone()[0] for i in (99, 100))
    assert (a99, a100) == (1, 0), (a99, a100)
    print("PASS sharedStrings + date serials + booleans + multi-sheet")


def _cells(r, values):
    out = []
    for j, v in enumerate(values):
        if v is not None:
            out.append(f'<c r="{chr(65+j)}{r}" t="inlineStr"><is><t>{v}</t></is></c>')
    return f'<row r="{r}">' + "".join(out) + "</row>"


def test_stacked_tables_diagnostics(tmp):
    """Two blocks with repeated headers separated by blank rows -> W01 + W02."""
    rows = [_cells(1, ["ID", "Name", "Val"])]
    for i in range(2, 22):
        rows.append(_cells(i, [f"a{i}", f"n{i}", f"v{i}"]))
    rows.append(_cells(30, ["ID", "Name", "Val"]))  # second block, rows 22-29 blank
    for i in range(31, 41):
        rows.append(_cells(i, [f"b{i}", f"m{i}", f"w{i}"]))
    xlsx = os.path.join(tmp, "stacked.xlsx")
    write_xlsx(xlsx, ["Report"], [sheet(rows)])
    conn, log = convert(xlsx, os.path.join(tmp, "stacked.sqlite"))
    assert "[W01]" in log, log     # blank gap detected
    assert "[W02]" in log, log     # repeated header detected
    n = conn.execute("SELECT COUNT(*) FROM report").fetchone()[0]
    assert n == 31, n              # 20 + repeated header + 10, nothing lost
    print("PASS stacked-tables diagnostics (W01 + W02)")


def test_title_row_diagnostic(tmp):
    """A one-cell title above the real header -> W07; --header-row 2 fixes it."""
    rows = [_cells(1, ["Quarterly Report 2024"]),
            _cells(2, ["ID", "Name", "Val"])]
    for i in range(3, 13):
        rows.append(_cells(i, [f"a{i}", f"n{i}", f"v{i}"]))
    xlsx = os.path.join(tmp, "titled.xlsx")
    write_xlsx(xlsx, ["Summary"], [sheet(rows)])
    _, log = convert(xlsx, os.path.join(tmp, "titled_bad.sqlite"))
    assert "[W07]" in log and "--header-row 2" in log, log
    conn, log2 = convert(xlsx, os.path.join(tmp, "titled_ok.sqlite"), "--header-row", "2")
    cols = [r[1] for r in conn.execute("PRAGMA table_info(summary)")]
    assert cols == ["id", "name", "val"], cols
    assert conn.execute("SELECT COUNT(*) FROM summary").fetchone()[0] == 10
    assert "diagnostics: clean" in log2, log2
    print("PASS title-row diagnostic (W07) + --header-row recovery")


class _RangeHandler(http.server.BaseHTTPRequestHandler):
    """Minimal HTTP server honoring Range requests (like real file hosts)."""
    payload = b""

    def do_GET(self):
        data = self.payload
        m = re.match(r"bytes=(\d+)-(\d*)$", self.headers.get("Range") or "")
        if m:
            start = int(m.group(1))
            end = min(int(m.group(2)) if m.group(2) else len(data) - 1, len(data) - 1)
            chunk = data[start:end + 1]
            self.send_response(206)
            self.send_header("Content-Range", f"bytes {start}-{end}/{len(data)}")
        else:
            chunk = data
            self.send_response(200)
        self.send_header("Content-Length", str(len(chunk)))
        self.end_headers()
        self.wfile.write(chunk)

    def log_message(self, *a):
        pass


class _NoRangeHandler(_RangeHandler):
    """Server that ignores Range — must be rejected with E06, not half-read."""
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Length", str(len(self.payload)))
        self.end_headers()
        self.wfile.write(self.payload)


def _serve(handler, payload):
    handler.payload = payload
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, f"http://127.0.0.1:{srv.server_address[1]}/file.xlsx"


def _small_orders_xlsx(tmp, n=300):
    rows = [_cells(1, ["ID", "Name", "Val"])]
    rows += [_cells(i, [f"a{i}", f"n{i}", f"v{i}"]) for i in range(2, n + 2)]
    path = os.path.join(tmp, "remote_src.xlsx")
    write_xlsx(path, ["Orders"], [sheet(rows)])
    return path, n


def test_remote_streaming(tmp):
    """Convert straight from a URL: ranged download + conversion in one pass."""
    path, n = _small_orders_xlsx(tmp)
    payload = open(path, "rb").read()
    srv, url = _serve(_RangeHandler, payload)
    try:
        conn, log = convert(url, os.path.join(tmp, "remote.sqlite"))
        assert "remote input:" in log and "network:" in log, log
        assert conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0] == n

        # multi-block path: tiny blocks force many ranged GETs + LRU eviction
        spec = importlib.util.spec_from_file_location("x2s", SCRIPT)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        hf = mod.HttpFile(url, block_size=1024, cache_blocks=3)
        with zipfile.ZipFile(hf) as zremote, zipfile.ZipFile(path) as zlocal:
            assert zremote.namelist() == zlocal.namelist()
            member = "xl/worksheets/sheet1.xml"
            assert zremote.read(member) == zlocal.read(member)
        assert hf.fetched > 5 * 1024, hf.fetched  # really came in many pieces
    finally:
        srv.shutdown()
    print("PASS remote streaming (ranged download -> sqlite, multi-block)")


def test_remote_no_range_support(tmp):
    """A server without Range support must fail fast with E06."""
    path, _ = _small_orders_xlsx(tmp)
    srv, url = _serve(_NoRangeHandler, open(path, "rb").read())
    try:
        r = subprocess.run([sys.executable, SCRIPT, url, "-o",
                            os.path.join(tmp, "never.sqlite")],
                           capture_output=True, text=True)
        assert r.returncode != 0 and "[E06]" in (r.stdout + r.stderr), r.stderr
    finally:
        srv.shutdown()
    print("PASS E06 on servers without range support")


def test_compress_roundtrip(tmp):
    """--compress: prose columns become zlib BLOBs, unz() restores them
    byte-for-byte, short columns stay plain, and the file shrinks."""
    prose = "the quick brown fox jumps over the lazy dog. " * 12  # ~540 B
    rows = [_cells(1, ["ID", "Tag", "Body"])]
    for i in range(2, 402):
        rows.append(_cells(i, [f"a{i}", f"t{i % 7}", f"{prose}#{i}"]))
    xlsx = os.path.join(tmp, "blog.xlsx")
    write_xlsx(xlsx, ["Posts"], [sheet(rows)])

    plain_db = os.path.join(tmp, "plain.sqlite")
    comp_db = os.path.join(tmp, "comp.sqlite")
    convert(xlsx, plain_db)
    conn, log = convert(xlsx, comp_db, "--compress")
    assert "compressing column(s) ['body']" in log, log

    assert conn.execute("SELECT typeof(body) FROM posts LIMIT 1").fetchone()[0] == "blob"
    assert conn.execute("SELECT typeof(tag) FROM posts LIMIT 1").fetchone()[0] == "text"
    assert conn.execute("SELECT tbl, col, codec FROM _compressed_columns").fetchall() \
        == [("posts", "body", "zlib")]
    import zlib as _z
    blob = conn.execute("SELECT body FROM posts WHERE id='a77'").fetchone()[0]
    assert _z.decompress(blob).decode() == f"{prose}#77"
    assert os.path.getsize(comp_db) < os.path.getsize(plain_db) / 3

    r = subprocess.run([sys.executable, os.path.join(REPO, "scripts", "zquery.py"),
                        comp_db, "SELECT unz(body) FROM posts WHERE id='a99'"],
                       capture_output=True, text=True)
    assert r.returncode == 0 and f"{prose}#99" in r.stdout, r.stdout + r.stderr

    # compress_db.py: same result starting from the already-plain database
    comp2 = os.path.join(tmp, "comp2.sqlite")
    r = subprocess.run([sys.executable, os.path.join(REPO, "scripts", "compress_db.py"),
                        plain_db, comp2], capture_output=True, text=True)
    assert r.returncode == 0 and "compressing ['body']" in r.stdout, r.stdout + r.stderr
    c2 = sqlite3.connect(comp2)
    blob = c2.execute("SELECT body FROM posts WHERE id='a77'").fetchone()[0]
    assert _z.decompress(blob).decode() == f"{prose}#77"
    assert c2.execute("SELECT COUNT(*) FROM posts").fetchone()[0] == 400
    print("PASS --compress + zquery + compress_db roundtrip")


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as tmp:
        test_inline_sparse_dirty_headers(tmp)
        test_sharedstrings_dates_bools(tmp)
        test_stacked_tables_diagnostics(tmp)
        test_title_row_diagnostic(tmp)
        test_remote_streaming(tmp)
        test_remote_no_range_support(tmp)
        test_compress_roundtrip(tmp)
    print("ALL TESTS PASS")

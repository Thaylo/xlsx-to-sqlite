#!/usr/bin/env python3
"""End-to-end tests for scripts/xlsx_to_sqlite.py. Stdlib only.

Builds two synthetic workbooks exercising the format traps that corrupt data
silently (sparse cells, date serials, sharedStrings, dirty headers), converts
them, and asserts on the resulting databases.
"""
import os
import sqlite3
import subprocess
import sys
import tempfile
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


def convert(xlsx, out):
    r = subprocess.run([sys.executable, SCRIPT, xlsx, "-o", out, "--force"],
                       capture_output=True, text=True)
    assert r.returncode == 0, f"converter failed:\n{r.stdout}\n{r.stderr}"
    return sqlite3.connect(out)


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
    conn = convert(xlsx, os.path.join(tmp, "orders.sqlite"))

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
    conn = convert(xlsx, os.path.join(tmp, "crm.sqlite"))

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


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as tmp:
        test_inline_sparse_dirty_headers(tmp)
        test_sharedstrings_dates_bools(tmp)
    print("ALL TESTS PASS")

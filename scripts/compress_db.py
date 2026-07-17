#!/usr/bin/env python3
"""Losslessly shrink an existing SQLite database by storing its big prose
columns as zlib BLOBs. Produces a new file; the source is never touched.

Usage:
  python3 compress_db.py source.sqlite dest.sqlite
  python3 compress_db.py source.sqlite dest.sqlite --cols articles.body,articles.notes

Auto-detection: TEXT columns averaging >256 bytes over the first 2000 rows.
Expect ~2x on unique prose (entropy-bound), more on repetitive text. Reads:
zquery.py or a one-line unz() UDF. Indexes are recreated on the new file;
compressed columns are recorded in the _compressed_columns table.
"""
import argparse
import os
import sqlite3
import sys
import zlib

MIN_AVG = 256
BATCH = 5000


def detect(conn, table, cols):
    sample = conn.execute(f'SELECT * FROM "{table}" LIMIT 2000').fetchall()
    picked = []
    for i, col in enumerate(cols):
        vals = [r[i] for r in sample if isinstance(r[i], str)]
        if len(vals) >= 10 and sum(map(len, vals)) / len(vals) > MIN_AVG:
            picked.append(col)
    return picked


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("source")
    p.add_argument("dest")
    p.add_argument("--cols", help="comma-separated table.column list overriding auto-detection")
    p.add_argument("--level", type=int, default=6, help="zlib level 1-9 (default 6)")
    opts = p.parse_args()
    if os.path.exists(opts.dest):
        sys.exit(f"refusing to overwrite {opts.dest}")

    src = sqlite3.connect(f"file:{opts.source}?mode=ro", uri=True)
    dst = sqlite3.connect(opts.dest)
    dst.executescript("PRAGMA page_size=8192; PRAGMA journal_mode=OFF; "
                      "PRAGMA synchronous=OFF; PRAGMA temp_store=MEMORY;")

    manual = {}
    if opts.cols:
        for tc in opts.cols.split(","):
            t, c = tc.strip().split(".")
            manual.setdefault(t, []).append(c)

    tables = [r[0] for r in src.execute(
        "SELECT name FROM sqlite_master WHERE type='table' "
        "AND name NOT LIKE 'sqlite_%' AND name != '_compressed_columns'")]
    dst.execute("CREATE TABLE _compressed_columns (tbl TEXT, col TEXT, codec TEXT)")

    for table in tables:
        info = src.execute(f'PRAGMA table_info("{table}")').fetchall()
        cols = [r[1] for r in info]
        comp = manual.get(table) if manual else detect(src, table, cols)
        comp = [c for c in (comp or []) if c in cols]
        comp_idx = {cols.index(c) for c in comp}
        decls = ", ".join(
            f'"{name}" {"BLOB" if i in comp_idx else (decl or "")}'.rstrip()
            for i, (_, name, decl, *_1) in enumerate(info))
        dst.execute(f'CREATE TABLE "{table}" ({decls})')
        for c in comp:
            dst.execute("INSERT INTO _compressed_columns VALUES (?, ?, 'zlib')", (table, c))
        print(f"{table}: compressing {comp or 'nothing'}", flush=True)

        ins = f'INSERT INTO "{table}" VALUES ({",".join("?" * len(cols))})'
        batch, n = [], 0
        for row in src.execute(f'SELECT * FROM "{table}"'):
            if comp_idx:
                row = tuple(
                    zlib.compress(v.encode("utf-8"), opts.level)
                    if i in comp_idx and isinstance(v, str) else v
                    for i, v in enumerate(row))
            batch.append(row)
            if len(batch) >= BATCH:
                dst.executemany(ins, batch)
                dst.commit()
                n += len(batch)
                batch.clear()
                if n % 100_000 == 0:
                    print(f"  {n:,} rows", flush=True)
        if batch:
            dst.executemany(ins, batch)
            dst.commit()
            n += len(batch)
        print(f"  {table}: {n:,} rows copied", flush=True)

    for (sql,) in src.execute(
            "SELECT sql FROM sqlite_master WHERE type='index' AND sql IS NOT NULL"):
        dst.execute(sql)
    dst.commit()
    dst.execute("PRAGMA journal_mode=WAL")
    dst.close()

    a, b = os.path.getsize(opts.source), os.path.getsize(opts.dest)
    print(f"DONE: {a / 1e9:.2f} GB -> {b / 1e9:.2f} GB ({(1 - b / a) * 100:.0f}% smaller)")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Query a SQLite database that has zlib-compressed columns (--compress /
compress_db.py). Registers an unz() SQL function that transparently
decompresses — it passes plain values through untouched, so unz(col) is
always safe on any column.

Usage:
  python3 zquery.py db.sqlite "SELECT id, unz(article_text) FROM t LIMIT 3"
  python3 zquery.py db.sqlite            # interactive: one query per line

In your own code, the whole trick is one line:
  conn.create_function("unz", 1, lambda v: zlib.decompress(v).decode()
                       if isinstance(v, (bytes, bytearray)) else v)
"""
import sqlite3
import sys
import zlib


def unz(v):
    if isinstance(v, (bytes, bytearray)):
        try:
            return zlib.decompress(v).decode("utf-8")
        except (zlib.error, UnicodeDecodeError):
            return v  # a genuine BLOB that isn't ours — hand it back untouched
    return v


def run(conn, query):
    try:
        cur = conn.execute(query)
    except sqlite3.Error as e:
        print(f"error: {e}", file=sys.stderr)
        return
    if cur.description:
        print("\t".join(d[0] for d in cur.description))
    for row in cur:
        print("\t".join("" if v is None else str(v) for v in row))
    conn.commit()


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    conn = sqlite3.connect(sys.argv[1])
    conn.create_function("unz", 1, unz)
    meta = conn.execute("SELECT tbl, col FROM _compressed_columns").fetchall() \
        if conn.execute("SELECT 1 FROM sqlite_master WHERE name='_compressed_columns'").fetchone() else []
    if meta:
        print(f"compressed columns (wrap in unz()): "
              f"{', '.join('%s.%s' % m for m in meta)}", file=sys.stderr)
    if len(sys.argv) > 2:
        run(conn, sys.argv[2])
        return
    for line in sys.stdin:
        if line.strip():
            run(conn, line)


if __name__ == "__main__":
    main()

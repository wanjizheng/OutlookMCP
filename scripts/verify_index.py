"""Read-only integrity check and non-content index inventory."""

import argparse
import json
import sqlite3
from pathlib import Path


def main():
    p = argparse.ArgumentParser()
    p.add_argument("database", type=Path)
    a = p.parse_args()
    c = sqlite3.connect(a.database.resolve().as_uri() + "?mode=ro", uri=True)
    try:
        c.execute("PRAGMA query_only=ON")
        check = c.execute("PRAGMA quick_check").fetchall()
        if check != [("ok",)]:
            raise RuntimeError("SQLite integrity check failed")
        print(
            json.dumps(
                {
                    "integrity": "ok",
                    "emails": c.execute("SELECT count(*) FROM emails").fetchone()[0],
                    "chunks": c.execute("SELECT count(*) FROM email_chunks").fetchone()[
                        0
                    ],
                }
            )
        )
    finally:
        c.close()


if __name__ == "__main__":
    main()

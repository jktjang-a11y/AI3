# -*- coding: utf-8 -*-
"""전국 법정동 지역 엑셀을 로컬 SQLite DB(local_db/region_codes.sqlite3)로 변환한다.

사용: python tools/import_region_codes.py ["전국 법정동 지역.xls"]
"""

from __future__ import annotations

import sqlite3
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SOURCE = ROOT / "전국 법정동 지역.xls"
DB_PATH = ROOT / "local_db" / "region_codes.sqlite3"


def main() -> None:
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SOURCE
    engine = "xlrd" if source.suffix.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(source, engine=engine)
    for col in ["REG", "시도", "시군구"]:
        if col not in df.columns:
            raise SystemExit(f"필수 컬럼 누락: {col}")

    df = df[["REG", "시도", "시군구"]].copy()
    df["REG"] = df["REG"].astype(str).str.extract(r"(\d+)")[0].str.zfill(5)
    df = df[df["REG"].str.len() == 5]
    df = df[df["시도"].notna() & df["시군구"].notna()].copy()
    df["시도"] = df["시도"].astype(str).str.strip()
    df["시군구"] = df["시군구"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["REG", "시도", "시군구"]).sort_values(["시도", "시군구"])

    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    if DB_PATH.exists():
        DB_PATH.unlink()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("CREATE TABLE region_codes (REG TEXT NOT NULL, 시도 TEXT NOT NULL, 시군구 TEXT NOT NULL, PRIMARY KEY (REG, 시도, 시군구))")
        conn.executemany("INSERT INTO region_codes (REG, 시도, 시군구) VALUES (?, ?, ?)", df.itertuples(index=False, name=None))
        conn.execute("CREATE TABLE metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
        conn.executemany(
            "INSERT INTO metadata (key, value) VALUES (?, ?)",
            [("source_file", source.name), ("imported_at", datetime.now().isoformat(timespec="seconds")), ("row_count", str(len(df)))],
        )
    print(f"완료: {len(df):,}건 -> {DB_PATH}")


if __name__ == "__main__":
    main()

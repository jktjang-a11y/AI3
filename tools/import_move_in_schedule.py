# -*- coding: utf-8 -*-
"""아파트 입주예정물량 엑셀을 보고서용 로컬 SQLite DB로 변환한다."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd


SIGUNGU_RE = re.compile(r"(?:광주광역시|전라남도|전남광주통합특별시)\s+([가-힣]+(?:시|군|구))")


def text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def integer(value):
    try:
        if value is None or pd.isna(value):
            return None
        return int(round(float(value)))
    except Exception:
        return None


def month_text(value) -> str:
    try:
        dt = pd.to_datetime(value)
        return dt.strftime("%Y-%m")
    except Exception:
        return ""


def derive_sigungu(address: str, original: str) -> tuple[str, str]:
    match = SIGUNGU_RE.search(address)
    if match:
        actual = match.group(1)
        note = ""
        if original and actual != original:
            note = f"원본 시군구({original})와 주소 기준 시군구({actual})가 달라 주소 기준 적용"
        return actual, note
    return original, "주소에서 시군구를 확인하지 못해 원본 시군구 적용"


def load_rows(workbook_path: Path) -> list[dict]:
    book = pd.ExcelFile(workbook_path)
    rows = []
    for sheet in book.sheet_names:
        if "APT" not in sheet.upper() or "입주" not in sheet or sheet == "변경내역":
            continue
        df = pd.read_excel(workbook_path, sheet_name=sheet, header=1)
        if "주택명" not in df.columns:
            continue
        for idx, row in df.iterrows():
            project = text(row.get("주택명"))
            if not project or project == "총 계":
                continue
            address = text(row.get("공급위치"))
            original_sigungu = text(row.get("시군구"))
            sigungu, location_note = derive_sigungu(address, original_sigungu)
            total_raw = integer(row.get("전체 세대수"))
            supplied = integer(row.get("공급규모")) or 0
            if total_raw is None:
                total = supplied
                basis = "공급규모 대체"
            else:
                total = total_raw
                basis = "전체 세대수"
            notes = [n for n in [location_note] if n]
            if total == 0 and supplied > 0:
                notes.append("취소분·추가공급 행으로 전체 입주물량 합계에서 제외")
            planned_month = month_text(row.get("입주예정월"))
            rows.append({
                "source_sheet": sheet,
                "source_row": int(idx) + 3,
                "housing_manage_no": text(row.get("주택관리번호")),
                "announcement_no": text(row.get("공고번호")),
                "project_name": project,
                "region_name_original": text(row.get("공급지역명")),
                "sigungu_original": original_sigungu,
                "sigungu": sigungu,
                "address": address,
                "planned_month": planned_month,
                "planned_year": int(planned_month[:4]) if planned_month else None,
                "total_households": total,
                "supplied_households": supplied,
                "household_basis": basis,
                "builder": text(row.get("건설업체명_시공사")),
                "developer": text(row.get("사업주체명_시행사")),
                "homepage": text(row.get("모집공고홈페이지주소")) or text(row.get("홈페이지주소")),
                "note": "; ".join(notes),
            })
    return rows


def write_db(workbook_path: Path, db_path: Path) -> dict:
    rows = load_rows(workbook_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256(workbook_path.read_bytes()).hexdigest()
    with sqlite3.connect(db_path) as conn:
        conn.executescript("""
        DROP TABLE IF EXISTS move_in_projects;
        DROP TABLE IF EXISTS metadata;
        CREATE TABLE metadata (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE move_in_projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_workbook TEXT NOT NULL,
            source_sheet TEXT NOT NULL,
            source_row INTEGER,
            housing_manage_no TEXT,
            announcement_no TEXT,
            project_name TEXT NOT NULL,
            region_name_original TEXT,
            sigungu_original TEXT,
            sigungu TEXT NOT NULL,
            address TEXT,
            planned_month TEXT NOT NULL,
            planned_year INTEGER NOT NULL,
            total_households INTEGER NOT NULL,
            supplied_households INTEGER NOT NULL,
            household_basis TEXT NOT NULL,
            builder TEXT,
            developer TEXT,
            homepage TEXT,
            note TEXT
        );
        CREATE INDEX idx_move_in_year_sigungu ON move_in_projects(planned_year, sigungu);
        CREATE INDEX idx_move_in_month ON move_in_projects(planned_month);
        """)
        cols = [
            "source_workbook", "source_sheet", "source_row", "housing_manage_no", "announcement_no",
            "project_name", "region_name_original", "sigungu_original", "sigungu", "address",
            "planned_month", "planned_year", "total_households", "supplied_households",
            "household_basis", "builder", "developer", "homepage", "note",
        ]
        payload = [[workbook_path.name] + [row[c] for c in cols[1:]] for row in rows]
        conn.executemany(
            f"INSERT INTO move_in_projects ({','.join(cols)}) VALUES ({','.join(['?'] * len(cols))})",
            payload,
        )
        metadata = {
            "schema_version": "1",
            "source_workbook": workbook_path.name,
            "source_path_at_import": str(workbook_path),
            "source_sha256": digest,
            "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "household_rule": "전체 세대수 우선, 결측 시 공급규모 대체, 명시적 0은 합계 제외",
            "location_rule": "공급위치 주소의 시군구 우선, 주소 확인 불가 시 원본 시군구 사용",
        }
        conn.executemany("INSERT INTO metadata(key,value) VALUES (?,?)", metadata.items())

        totals = conn.execute("""
            SELECT planned_year, sigungu, COUNT(*) AS project_rows,
                   SUM(CASE WHEN total_households > 0 THEN 1 ELSE 0 END) AS counted_projects,
                   SUM(total_households) AS total_households
            FROM move_in_projects
            GROUP BY planned_year, sigungu
            ORDER BY planned_year, sigungu
        """).fetchall()
    return {"database": str(db_path), "rows": len(rows), "totals": totals}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--db", type=Path, default=Path("local_db/apt_move_in_schedule.sqlite3"))
    args = parser.parse_args()
    print(json.dumps(write_db(args.workbook.resolve(), args.db.resolve()), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

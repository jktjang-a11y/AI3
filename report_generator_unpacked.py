# -*- coding: utf-8 -*-
"""
월간·주간 주택동향보고서 HTML 최종보고서 생성기 v5.8.41
- Windows Tkinter UI
- 20개 원자료 전체 반영 목표
- data_pack 확장: 인구/세대, 연령·이동, 금리, 미분양, 가격지수, 공급, 아파트 매매/전세/월세,
  비아파트 매매/전세/월세, 분양권전매
- LLM 프롬프트 생성 및 JSON 해설 적용
- HTML 보고서 및 검수표 생성

실행:
  pip install pandas openpyxl jinja2 beautifulsoup4
  python monthly_report_generator_windows_ui_html_template_v58_style_v5_8_38.py
"""

from __future__ import annotations

import json
import html
import os
import re
import copy
import sqlite3
from collections import Counter
import sys
import traceback
import threading
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pandas as pd
    import numpy as np
    from bs4 import BeautifulSoup
    from jinja2 import Environment
except Exception as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "필수 패키지 설치 필요",
        "아래 명령어를 실행하세요.\n\npip install pandas openpyxl jinja2 beautifulsoup4\n\n"
        f"오류: {e}"
    )
    raise


APP_VERSION = "windows_ui_html_template_v58_reviewfix_v5_8_69_field_survey_narrative"


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"
PROMPT_DIR = BASE_DIR / "prompts"

for d in [INPUT_DIR, OUTPUT_DIR, LOG_DIR, PROMPT_DIR]:
    d.mkdir(exist_ok=True)


PROGRESS_CALLBACK = None


def set_progress_callback(callback):
    global PROGRESS_CALLBACK
    PROGRESS_CALLBACK = callback


def progress(message: str, current: Optional[int] = None, total: Optional[int] = None):
    if PROGRESS_CALLBACK:
        try:
            PROGRESS_CALLBACK(message, current, total)
        except Exception:
            pass


# =========================================================
# 1. 공통 유틸
# =========================================================

COLUMN_CANDIDATES = {
    "month": ["계약년월", "계약월", "기준월", "년월", "기간", "시점", "PRD_DE", "월"],
    "region": ["지역", "시도", "시도명", "시군구", "시군구명", "행정구역", "행정구역별", "법정동시군구", "법정동시군구명", "지역명"],
    "dong": ["읍면동", "법정동", "동", "법정동명"],
    "item": ["항목명", "분류1명", "분류2명", "분류", "지표", "유형", "주택유형", "물건구분"],
    "value": ["값", "데이터", "수치", "OBS_VALUE", "DT", "VALUE"],
    "apt_name": ["단지명", "아파트명", "건물명", "오피스텔명", "오피스텔", "offiNm", "aptNm", "주택명", "연립다세대명", "단독다가구명"],
    "jibun": ["지번", "본번", "번지"],
    "approval_year": ["사용승인년도", "사용승인일", "사용검사일", "건축년도", "준공년도", "건축연도", "준공일"],
    "households": ["세대수", "주민등록세대수"],
    "exclusive_area": ["전용면적", "전용면적㎡", "전용면적(㎡)", "계약면적", "건물면적", "면적"],
    "price": ["거래금액", "거래금액만원", "거래금액(만원)", "계약총금액", "신고가", "매매가", "거래가액"],
    "deposit": ["보증금", "보증금액", "보증금(만원)", "임대보증금"],
    "monthly_rent": ["월세", "월세금액", "월세금(만원)", "월차임"],
    "floor": ["층", "해당층"],
    "contract_date": ["계약일", "계약일자", "계약년월일", "거래일"],
    "contract_type": ["계약구분", "재계약", "갱신요구권사용", "신규갱신", "갱신여부"],
    "population": ["주민등록인구", "인구수", "인구", "계"],
    "total_households": ["주민등록세대수", "세대수"],
    "mortgage_rate": ["주택담보대출", "주택담보대출금리"],
    "jeonse_rate": ["전세자금대출", "전세자금대출금리"],
    "sme_rate": ["중소기업대출", "중소기업대출금리"],
    "business_rate": ["기업일반자금대출", "기업일반자금대출금리"],
    "unsold": ["미분양", "미분양호수", "미분양주택현황"],
    "after_completion_unsold": ["준공후미분양", "준공후미분양호수", "준공후"],
}


COMMENTARY_KEYS = [
    "monthly_core_diagnosis",
    "population_commentary",
    "age_migration_commentary",
    "rate_commentary",
    "unsold_commentary",
    "unsold_news_summary",
    "supply_commentary",
    "price_index_commentary",
    "conversion_rate_commentary",
    "apt_trade_commentary",
    "supply_type_commentary",
    "apt_price_band_commentary",
    "apt_age_band_commentary",
    "apt_area_band_commentary",
    "external_buyer_commentary",
    "apt_jeonse_commentary",
    "apt_rent_commentary",
    "non_apt_sale_commentary",
    "non_apt_jeonse_commentary",
    "non_apt_rent_commentary",
    "presale_commentary",
    "top_case_commentary",
    "comprehensive_summary",

    "sale_market_summary",
    "apt_sale_summary",
    "non_apt_sale_regional_context",
    "rowhouse_multiplex_sale_summary",
    "detached_multifamily_sale_summary",
    "jeonse_market_summary",
    "apt_jeonse_summary",
    "non_apt_jeonse_regional_context",
    "rowhouse_multiplex_jeonse_summary",
    "detached_multifamily_jeonse_summary",
    "rent_market_summary",
    "apt_rent_summary",
    "non_apt_rent_regional_context",
    "rowhouse_multiplex_rent_summary",
    "detached_multifamily_rent_summary",
]

LLM_FINAL_SUMMARY_KEYS = [
    "comprehensive_summary",
    "sale_market_summary",
    "apt_sale_summary",
    "non_apt_sale_regional_context",
    "rowhouse_multiplex_sale_summary",
    "detached_multifamily_sale_summary",
    "jeonse_market_summary",
    "apt_jeonse_summary",
    "non_apt_jeonse_regional_context",
    "rowhouse_multiplex_jeonse_summary",
    "detached_multifamily_jeonse_summary",
    "rent_market_summary",
    "apt_rent_summary",
    "non_apt_rent_regional_context",
    "rowhouse_multiplex_rent_summary",
    "detached_multifamily_rent_summary",
]

WEEKLY_LLM_FINAL_SUMMARY_KEYS = [
    "comprehensive_summary",
    "sale_market_summary",
    "apt_sale_summary",
    "jeonse_market_summary",
    "apt_jeonse_summary",
]


FORBIDDEN_PHRASES = ["신고가 기준"]


def clean_col_name(x: Any) -> str:
    return str(x).strip().replace("\n", "").replace("\r", "").replace(" ", "")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_col_name(c) for c in df.columns]
    return df


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    cands = [clean_col_name(c) for c in candidates]
    for cand in cands:
        for col in cols:
            if cand == str(col):
                return col
    for cand in cands:
        for col in cols:
            if cand and cand in str(col):
                return col
    return None


def to_numeric_safe(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("억", "", regex=False)
        .str.replace("만원", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip(),
        errors="coerce"
    )


def normalize_money_to_manwon(s: pd.Series, kind: str) -> pd.Series:
    """API의 만원 단위와 지정 업로드 원장의 원 단위를 공통 만원으로 맞춘다."""
    out = to_numeric_safe(s)
    valid = out.dropna().abs()
    if valid.empty:
        return out
    median = float(valid.median())
    threshold = {"price": 10_000_000, "deposit": 1_000_000, "rent": 10_000}.get(kind, 10_000_000)
    if median >= threshold:
        out = out / 10_000.0
    return out


def normalize_month_value(x: Any) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # Excel date
    try:
        if isinstance(x, (pd.Timestamp, datetime)):
            return pd.to_datetime(x).strftime("%Y-%m")
    except Exception:
        pass
    nums = re.sub(r"[^0-9]", "", s)
    if len(nums) >= 6:
        return f"{nums[:4]}-{nums[4:6]}"
    if re.fullmatch(r"\d{4}-\d{2}", s):
        return s
    return s[:7]


def normalize_approval_year_value(x: Any) -> Optional[int]:
    """사용승인년도·사용승인일·Excel 날짜 일련번호에서 승인연도를 추출한다."""
    if x is None or pd.isna(x):
        return None
    try:
        if isinstance(x, (pd.Timestamp, datetime)):
            year = int(pd.to_datetime(x).year)
            return year if 1800 <= year <= datetime.now().year + 2 else None
    except Exception:
        pass
    text = str(x).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    match = re.search(r"(?<!\d)((?:18|19|20)\d{2})(?!\d)", text)
    if match:
        return int(match.group(1))
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 8 and re.fullmatch(r"(?:18|19|20)\d{6,}", digits):
        return int(digits[:4])
    try:
        number = float(text.replace(",", ""))
        if 1800 <= number <= datetime.now().year + 2:
            return int(number)
        if 1 <= number <= 100000:
            year = int((pd.Timestamp("1899-12-30") + pd.to_timedelta(number, unit="D")).year)
            return year if 1800 <= year <= datetime.now().year + 2 else None
    except Exception:
        pass
    return None


def renewal_contract_mask(values: pd.Series) -> pd.Series:
    """전월세 원장의 재계약·갱신 표기를 갱신재계약으로 통합 판정."""
    normalized = values.fillna("").astype(str).str.replace(" ", "", regex=False).str.strip()
    return normalized.str.contains(r"재계약|갱신", regex=True, na=False)


def find_best_month_col(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    priority = ["계약년월", "계약일자", "계약일", "거래일", "계약월", "기준월", "년월", "기간", "시점", "PRD_DE"]
    for kw in priority:
        for col in df.columns:
            if kw in str(col):
                vals = df[col].dropna().head(30).tolist()
                norm = [normalize_month_value(v) for v in vals]
                good = sum(1 for v in norm if re.fullmatch(r"20\d{2}-\d{2}", str(v)))
                if good >= max(1, min(3, len(vals))):
                    return col
    # 월 단독값은 마지막 수단
    for col in df.columns:
        if "월" in str(col):
            vals = df[col].dropna().astype(str).head(20).tolist()
            if vals and all(re.fullmatch(r"\d{1,2}(\.0)?", v.strip()) for v in vals[:min(10, len(vals))]):
                continue
            return col
    return None



def find_supply_value_col(df: pd.DataFrame) -> Optional[str]:
    """주택건설 공급 통계 전용 수치 컬럼 탐색.
    면적·코드·연도 컬럼이 값으로 오인되면 백만 단위로 튀는 문제가 있어
    호수/세대수/실적/데이터값 계열만 우선 사용한다.
    """
    if df is None or df.empty:
        return None
    exclude_keywords = ["면적", "연면적", "대지", "건축면적", "전용", "공급면적", "코드", "code", "ID", "시점", "기간", "기준월", "연도", "월"]
    priority_groups = [
        ["호수", "세대수", "가구수", "실적", "건수", "동수"],
        ["값", "데이터", "OBS_VALUE", "VALUE", "DT"],
    ]
    for group in priority_groups:
        candidates = []
        for c in df.columns:
            cs = str(c)
            if any(x.lower() in cs.lower() for x in exclude_keywords):
                continue
            if any(k.lower() in cs.lower() for k in group):
                s = to_numeric_safe(df[c])
                cnt = int(s.notna().sum())
                if cnt > 0:
                    med = float(s.dropna().abs().median()) if not s.dropna().empty else 0
                    candidates.append((-cnt, med, c))
        if candidates:
            candidates.sort()
            return candidates[0][2]
    return find_value_col(df)

def find_value_col(df: pd.DataFrame) -> Optional[str]:
    col = find_col(df, COLUMN_CANDIDATES["value"])
    if col:
        return col
    best_col, best_cnt = None, 0
    for c in df.columns:
        cnt = int(to_numeric_safe(df[c]).notna().sum())
        if cnt > best_cnt:
            best_col, best_cnt = c, cnt
    return best_col


def find_item_col_by_keywords(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    for c in df.columns:
        sample = " ".join(df[c].dropna().astype(str).head(500).tolist())
        if any(k in sample for k in keywords):
            return c
    return None


def format_num(x: Any) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{int(round(float(x))):,}"
    except Exception:
        return str(x)


def format_float(x: Any, digits: int = 1) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x):.{digits}f}"
    except Exception:
        return str(x)


def format_rate(x: Any) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x):.1f}%"
    except Exception:
        return str(x)


def money_to_eok_from_manwon(value: Any) -> Optional[float]:
    try:
        if value is None or pd.isna(value):
            return None
        return round(float(value) / 10000, 2)
    except Exception:
        return None


def area_to_pyeong(area_m2: Any) -> Optional[str]:
    """실거래 사례용 통용 평형 표기. 전용 84㎡를 단순 25평으로 표시하지 않음."""
    try:
        if area_m2 is None or pd.isna(area_m2):
            return None
        a = float(area_m2)
        if 37 <= a <= 42:
            return "약 17평형"
        if 47 <= a <= 51:
            return "약 21평형"
        if 57 <= a <= 61:
            return "약 24~25평형"
        if 72 <= a <= 76:
            return "약 30평형"
        if 82 <= a <= 86:
            return "약 33~34평형"
        if 100 <= a <= 103:
            return "약 39~40평형"
        if 112 <= a <= 116:
            return "약 43~45평형"
        if a >= 135:
            return "대형 평형"
        return f"약 {int(round(a * 0.3025))}평"
    except Exception:
        return None


def safe_ratio(n: float, d: float) -> float:
    return round(float(n) / float(d) * 100, 1) if d else 0.0


def report_name_key(label: str = "") -> str:
    return "building_name" if "오피스텔" in str(label) else "apt_name"


def _read_large_inline_xlsx(path: Path, blank_limit: int = 200) -> pd.DataFrame:
    """서식 빈 행이 대량 저장된 inline-string XLSX를 스트리밍으로 읽는다."""
    from zipfile import ZipFile
    import xml.etree.ElementTree as ET

    def col_index(ref: str) -> int:
        letters = re.match(r"[A-Z]+", str(ref or "A"))
        value = 0
        for ch in (letters.group(0) if letters else "A"):
            value = value * 26 + (ord(ch) - 64)
        return value - 1

    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    header = []
    records = []
    blank_run = 0
    repeat_run = 0
    previous_row = None
    with ZipFile(path) as zf, zf.open("xl/worksheets/sheet1.xml") as stream:
        for _, elem in ET.iterparse(stream, events=("end",)):
            if elem.tag != f"{ns}row":
                continue
            values = {}
            for cell in elem.findall(f"{ns}c"):
                idx = col_index(cell.attrib.get("r", "A1"))
                if cell.attrib.get("t") == "inlineStr":
                    value = "".join(t.text or "" for t in cell.findall(f".//{ns}t"))
                else:
                    node = cell.find(f"{ns}v")
                    value = node.text if node is not None else ""
                values[idx] = value
            width = (max(values) + 1) if values else 0
            row = [values.get(i, "") for i in range(width)]
            if not header:
                header = row
            else:
                row = (row + [""] * len(header))[:len(header)]
                if not any(v not in (None, "") for v in row):
                    blank_run += 1
                    if blank_run >= blank_limit:
                        elem.clear()
                        break
                else:
                    blank_run = 0
                    row_signature = tuple(row)
                    repeat_run = repeat_run + 1 if row_signature == previous_row else 0
                    previous_row = row_signature
                    records.append(row)
                    # 잘못된 내보내기에서 마지막 실거래 한 행이 Excel 끝까지 복제된 경우 차단.
                    if repeat_run >= 500:
                        del records[-repeat_run:]
                        elem.clear()
                        break
            elem.clear()
    return pd.DataFrame(records, columns=header)


def read_excel_or_csv(path: Path) -> pd.DataFrame:
    progress(f"파일 읽는 중: {path.name}")
    if path.suffix.lower() == ".csv":
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(path, encoding="cp949")
    else:
        # 일부 사용자 지정 전월세 원장은 서식 때문에 Excel 최대행이 100만 행으로
        # 기록되어 있다. pandas가 빈 행까지 읽지 않도록 연속 빈 행에서 조기 종료한다.
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            very_large_sheet = int(ws.max_row or 0) > 200_000
            wb.close()
            df = _read_large_inline_xlsx(path) if very_large_sheet else pd.read_excel(path)
        except Exception:
            df = pd.read_excel(path)
    progress(f"파일 읽기 완료: {path.name} / {len(df):,}행, {len(df.columns):,}열")
    return df


def save_json(obj: Dict[str, Any], path: Path):
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def read_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def region_tokens(region: str) -> List[str]:
    parts = [p for p in re.split(r"\s+", (region or "").strip()) if p]
    tokens = []
    for p in parts:
        tokens.append(p)
        if p.endswith("광역시"):
            tokens.append(p.replace("광역시", ""))
        if p.endswith("특별시"):
            tokens.append(p.replace("특별시", ""))
    if parts:
        tokens.append(parts[-1])
    return list(dict.fromkeys([t for t in tokens if t]))


def filter_df_by_region(df: pd.DataFrame, region: str) -> pd.DataFrame:
    if df is None or df.empty or not region:
        return df
    tokens = region_tokens(region)
    if not tokens:
        return df
    candidate_cols = []
    if any(x in region for x in ["전남광주통합특별시", "광주권역"]):
        # 통합 행정구역명으로 조회하더라도 시군구 보고서는 종전 KOSIS
        # 시군구명(예: 광산구 -> 광산)을 먼저 선택해야 한다. 과거 로직은
        # 무조건 광주 시도 합계 행을 반환해 '시군구 공표통계'가 광주 전체
        # 수치로 표시되는 오류가 있었다.
        last_region = tokens[-1] if tokens else ""
        sigungu_target = re.sub(r"(시|군|구)$", "", last_region).strip()
        level2_cols = [
            c for c in df.columns
            if str(c) in {"C2_NM", "OBJ_L2_NM", "region_level_2", "분류2명", "분류2"}
        ]
        if sigungu_target and last_region not in {"전남광주통합특별시", "광주권역"}:
            sigungu_mask = pd.Series(False, index=df.index)
            for c in level2_cols:
                normalized = df[c].astype(str).str.strip().str.replace(r"(시|군|구)$", "", regex=True)
                sigungu_mask = sigungu_mask | normalized.eq(sigungu_target)
            if sigungu_mask.any():
                return df[sigungu_mask].copy()

        # 2026-07 통합 이후 KOSIS는 종전 광주를 '(구)광주' 또는 '광주권역'으로 표기한다.
        gwangju_names = ["광주", "(구)광주", "광주권역"]
        # 유형별 가격지수처럼 분류1이 주택유형이고 분류2가 지역인 표도 있으므로
        # 분류1·분류2 이름 컬럼을 모두 확인한다.
        direct_cols = [
            c for c in df.columns
            if str(c) in {"C1_NM", "OBJ_L1_NM", "region_level_1", "분류1명", "분류1"}
        ] + level2_cols
        for gwangju in gwangju_names:
            mask = pd.Series(False, index=df.index)
            for c in direct_cols:
                mask = mask | df[c].astype(str).str.strip().eq(gwangju)
            if mask.any():
                return df[mask].copy()
    for c in df.columns:
        cs = str(c)
        if any(k in cs for k in ["지역", "시도", "시군구", "행정구역", "법정동시군구", "분류1명", "분류2명", "C1_NM", "C2_NM", "C3_NM", "OBJ_L1_NM", "OBJ_L2_NM"]):
            candidate_cols.append(c)
    if not candidate_cols:
        return df
    full = region.replace(" ", "")
    mask = pd.Series(False, index=df.index)
    for c in candidate_cols:
        s = df[c].astype(str).str.replace(" ", "", regex=False)
        mask = mask | s.str.contains(full, na=False, regex=False)
    if mask.any():
        return df[mask].copy()
    last = tokens[-1]
    mask = pd.Series(False, index=df.index)
    for c in candidate_cols:
        mask = mask | df[c].astype(str).str.contains(last, na=False, regex=False)
    if mask.any():
        return df[mask].copy()
    return df


def infer_region_month_from_filenames(paths: List[Path]) -> Dict[str, str]:
    months, regions = [], []
    for p in paths:
        stem = p.stem
        ym = re.findall(r"(20\d{2})[_-](\d{2})", stem)
        if ym:
            months.append(f"{ym[-1][0]}-{ym[-1][1]}")
        all_dates = list(re.finditer(r"20\d{2}[_-]\d{2}", stem))
        if all_dates:
            tail = stem[all_dates[-1].end():].lstrip("_-")
            tail = re.sub(r"\(\d+\)$", "", tail).strip("_- ")
            if tail:
                regions.append(tail.replace("_", " ").replace("-", " "))
    out = {}
    if months:
        out["base_month"] = max(set(months), key=months.count)
    if regions:
        out["region"] = max(set(regions), key=regions.count)
    return out


# =========================================================
# 2. 파일 분류
# =========================================================

INPUT_SOURCE_LABELS = {
    "auto": "자동 판정",
    "api": "국토교통부 실거래 API 기반",
    "excel_upload": "지정 엑셀 원장 기반",
}


def is_api_rtms_file(path: Path) -> bool:
    name = path.name.lower()
    return "국토교통부_" in name and "실거래가" in name


def is_designated_housing_excel_file(path: Path) -> bool:
    """사용자가 지정한 주거용 통합 매매·전월세 원장인지 파일명으로 1차 판정."""
    if is_api_rtms_file(path):
        return False
    name = path.name.lower()
    if any(token in name for token in ["전월세전환율", "전월세_전환율", "가격지수", "공표통계"]):
        return False
    return ("실거래" in name and "매매" in name and "전월세" not in name) or "전월세" in name


def filter_files_by_input_source(paths: List[Path], input_source: str = "auto") -> Tuple[List[Path], List[Path]]:
    """공통 통계는 유지하고 두 방식의 실거래 원장만 상호 배제한다."""
    mode = str(input_source or "auto").strip().lower()
    selected, excluded = [], []
    for path in paths:
        p = Path(path)
        if mode == "api" and is_designated_housing_excel_file(p):
            excluded.append(p)
            continue
        if mode == "excel_upload" and is_api_rtms_file(p):
            excluded.append(p)
            continue
        selected.append(p)
    return selected, excluded


def validate_input_source_files(
    paths: List[Path],
    input_source: str = "auto",
    designated_sale_files: Optional[List[Path]] = None,
    designated_rent_files: Optional[List[Path]] = None,
) -> None:
    mode = str(input_source or "auto").strip().lower()
    if mode == "api" and not any(is_api_rtms_file(Path(p)) for p in paths):
        raise ValueError("API 전용 보고서는 국토교통부 실거래 API 수집 엑셀이 필요합니다.")
    if mode == "excel_upload":
        if designated_sale_files is not None or designated_rent_files is not None:
            sale_files = {str(Path(p).resolve()).lower() for p in (designated_sale_files or []) if Path(p).exists()}
            rent_files = {str(Path(p).resolve()).lower() for p in (designated_rent_files or []) if Path(p).exists()}
            if len(sale_files) < 1 or len(rent_files) < 3:
                raise ValueError("지정 엑셀 보고서는 매매 원장 1개와 최근 3년 전월세 원장 3개를 연도별로 각각 지정해야 합니다.")
            return
        uploaded = [Path(p) for p in paths if is_designated_housing_excel_file(Path(p))]
        has_sale = any("실거래" in p.name and "매매" in p.name and "전월세" not in p.name for p in uploaded)
        rent_files = {str(p.resolve()).lower() for p in uploaded if "전월세" in p.name}
        if not has_sale or len(rent_files) < 3:
            raise ValueError("지정 엑셀 보고서는 주거용 통합 매매 원장 1개와 최근 3년 전월세 원장 3개를 연도별로 각각 지정해야 합니다.")


def classify_files(
    paths: List[Path],
    input_source: str = "auto",
    designated_sale_files: Optional[List[Path]] = None,
    designated_rent_files: Optional[List[Path]] = None,
) -> Dict[str, List[Path]]:
    buckets = {
        "population": [], "migration": [], "rates": [], "unsold": [], "supply": [],
        "price_index_sale": [], "price_index_jeonse": [], "price_index_rent": [],
        "weekly_apt_sale_change": [], "weekly_apt_jeonse_change": [],
        "apt_trade": [], "apt_rent": [], "non_apt_sale": [], "non_apt_rent": [],
        "presale": [], "conversion_rate_public": [], "buyer_sale_upload": [], "unknown": [],
    }
    source_paths, _ = filter_files_by_input_source(paths, input_source)
    designated_sale_keys = {str(Path(p).resolve()).lower() for p in (designated_sale_files or [])}
    designated_rent_keys = {str(Path(p).resolve()).lower() for p in (designated_rent_files or [])}
    for p in source_paths:
        path_key = str(Path(p).resolve()).lower()
        if path_key in designated_sale_keys:
            buckets["buyer_sale_upload"].append(p)
            buckets["apt_trade"].append(p)
            buckets["non_apt_sale"].append(p)
            continue
        if path_key in designated_rent_keys:
            buckets["apt_rent"].append(p)
            buckets["non_apt_rent"].append(p)
            continue
        name = p.name.lower()
        # v5.8.52: 공표 전월세전환율(KOSIS) 엑셀을 별도 버킷으로 분류.
        # 파일명에 '전월세전환율' 이 포함되면 conversion_rate_public 으로 보내고
        # 시도+물건유형+근접월 기준으로 동적 매칭.
        if "전월세전환율" in name:
            buckets["conversion_rate_public"].append(p)
            continue
        if "연령5세별_이동자수" in name or "이동자수" in name:
            buckets["migration"].append(p)
        elif "주민등록인구" in name or "주민등록세대" in name or ("인구" in name and "이동" not in name):
            buckets["population"].append(p)
        elif "대출금리" in name or "금리" in name or "은행" in name:
            buckets["rates"].append(p)
        elif "미분양" in name:
            buckets["unsold"].append(p)
        elif "인허가" in name or "착공" in name or "준공실적" in name or "분양실적" in name or "주택건설" in name:
            buckets["supply"].append(p)
        elif "주간아파트" in name and "매매" in name and "변동률" in name:
            buckets["weekly_apt_sale_change"].append(p)
        elif "주간아파트" in name and "전세" in name and "변동률" in name:
            buckets["weekly_apt_jeonse_change"].append(p)
        elif "매매가격지수" in name:
            buckets["price_index_sale"].append(p)
        elif "전세가격지수" in name:
            buckets["price_index_jeonse"].append(p)
        elif "월세통합가격지수" in name or "월세가격지수" in name:
            buckets["price_index_rent"].append(p)
        elif "오피스텔" in name and "전월세" in name:
            buckets["apt_rent"].append(p)
        elif "오피스텔" in name and "매매" in name:
            buckets["apt_trade"].append(p)
        elif "분양권" in name:
            buckets["presale"].append(p)
        elif "아파트" in name and "전월세" in name:
            buckets["apt_rent"].append(p)
        elif "아파트" in name and "매매" in name:
            buckets["apt_trade"].append(p)
        elif ("연립다세대" in name or "단독다가구" in name) and "전월세" in name:
            buckets["non_apt_rent"].append(p)
        elif ("연립다세대" in name or "단독다가구" in name) and "매매" in name:
            buckets["non_apt_sale"].append(p)
        elif "실거래" in name and "매매" in name and "전월세" not in name:
            # 사용자가 지정한 주거용 통합 매매원장. 실제 사용 여부는
            # 매수자 시군구·계약일자·주거용 물건구분을 다시 확인한 뒤 결정한다.
            buckets["buyer_sale_upload"].append(p)
            buckets["apt_trade"].append(p)
            buckets["non_apt_sale"].append(p)
        elif "전월세" in name:
            # 아파트 API 원자료 대신 지정 업로드한 주거용 통합 전월세 원장.
            # prepare_trade_df에서 주용도·건물유형을 확인해 아파트 행만 사용한다.
            buckets["apt_rent"].append(p)
            buckets["non_apt_rent"].append(p)
        else:
            buckets["unknown"].append(p)
    return buckets


# =========================================================
# 3. 데이터 진단
# =========================================================

def audit_single_file(path: Path) -> Dict[str, Any]:
    info = {"file_name": path.name, "read_ok": False, "rows": 0, "cols": 0, "columns": [], "matched": {}, "samples": {}, "warnings": []}
    try:
        df = normalize_columns(read_excel_or_csv(path))
        info["read_ok"] = True
        info["rows"], info["cols"] = len(df), len(df.columns)
        info["columns"] = [str(c) for c in df.columns]
        for k, cands in COLUMN_CANDIDATES.items():
            c = find_col(df, cands)
            info["matched"][k] = c or ""
            info["samples"][k] = df[c].dropna().astype(str).head(5).tolist() if c else []
        best_month = find_best_month_col(df)
        info["matched"]["best_month"] = best_month or ""
        if best_month:
            info["samples"]["best_month"] = df[best_month].dropna().astype(str).head(10).tolist()
    except Exception as e:
        info["warnings"].append(str(e))
    return info


def generate_data_audit_report(paths: List[Path]) -> Dict[str, Path]:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    txt_path = LOG_DIR / f"data_audit_{ts}.txt"
    xlsx_path = LOG_DIR / f"data_audit_{ts}.xlsx"
    items = [audit_single_file(p) for p in paths]

    lines, summary_rows, match_rows, col_rows = [], [], [], []
    lines.append("월간보고서 원자료 데이터 진단")
    lines.append("=" * 80)
    for it in items:
        lines.append(f"\n[{it['file_name']}] {it['rows']:,}행 {it['cols']:,}열")
        lines.append(", ".join(it["columns"]))
        summary_rows.append({"파일명": it["file_name"], "읽기성공": it["read_ok"], "행수": it["rows"], "열수": it["cols"], "경고": " / ".join(it["warnings"])})
        for k, v in it["matched"].items():
            match_rows.append({"파일명": it["file_name"], "표준키": k, "인식컬럼": v, "샘플": ", ".join(it["samples"].get(k, []))})
        for i, c in enumerate(it["columns"], start=1):
            col_rows.append({"파일명": it["file_name"], "순번": i, "컬럼명": c})
    txt_path.write_text("\n".join(lines), encoding="utf-8")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="요약")
        pd.DataFrame(match_rows).to_excel(writer, index=False, sheet_name="컬럼매칭")
        pd.DataFrame(col_rows).to_excel(writer, index=False, sheet_name="전체컬럼")
    return {"txt": txt_path, "xlsx": xlsx_path}


# =========================================================
# 4. 계산 모듈
# =========================================================

def has_multiple_region_members(df: pd.DataFrame) -> bool:
    """Combined sido reports contain one file per sigungu; totals must be summed."""
    if df is None or df.empty:
        return False
    region_cols = [
        "C2_NM", "C3_NM", "OBJ_L2_NM", "OBJ_L3_NM",
        "region_level_2", "region_level_3", "obj_level_2", "obj_level_3",
        "\ubd84\ub9581\uba85", "\ubd84\ub9581",
        "시군구", "시군구명", "법정동시군구", "법정동시군구명",
    ]
    for col in region_cols:
        if col not in df.columns:
            continue
        values = (
            df[col]
            .dropna()
            .astype(str)
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
            .dropna()
        )
        if values.nunique() > 1:
            return True
    return False


def aggregate_long_value(files: List[Path], region: str, base_month: str, metric_name_by_file: Dict[str, str]) -> Dict[str, Dict[str, Any]]:
    monthly: Dict[str, Dict[str, Any]] = {}
    for p in files:
        df = normalize_columns(read_excel_or_csv(p))
        df = filter_df_by_region(df, region)
        mcol = find_best_month_col(df)
        vcol = find_value_col(df)
        if not vcol:
            continue
        df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
        df["_value"] = to_numeric_safe(df[vcol])
        metric = None
        for key, metric_name in metric_name_by_file.items():
            if key in p.name:
                metric = metric_name
                break
        if not metric:
            metric = p.stem
        agg_method = "sum" if has_multiple_region_members(df) else "max"
        g = getattr(df.groupby("_month")["_value"], agg_method)().reset_index()
        for _, r in g.iterrows():
            m = r["_month"]
            if str(m).startswith("20") and pd.notna(r["_value"]):
                monthly.setdefault(m, {})[metric] = float(r["_value"])
    return monthly


def _filter_population_total_rows(df: pd.DataFrame, item_col: Optional[str]) -> pd.DataFrame:
    """주민등록인구 원자료에서 총인구 행만 남긴다.

    시도 취합 파일에는 시군구별 총인구/남자/여자 및 연령 상세가 함께 들어올 수 있어
    '총인구수' 문자열만으로 고르면 연령별 총인구가 중복 합산될 수 있다.
    """
    if df is None or df.empty:
        return df
    out = df.copy()
    if item_col and item_col in out.columns:
        item = out[item_col].astype(str).str.strip()
        total_mask = item.isin({"총인구수", "총인구", "인구수", "계"}) | item.str.fullmatch(r".*총인구수?", na=False)
        sex_mask = item.str.contains("남자인구수|여자인구수|남자|여자", na=False)
        if total_mask.any():
            out = out[total_mask & ~sex_mask].copy()

    total_like = ["총인구수", "총인구", "인구수", "계", "합계", "전체"]
    age_like_cols = [
        c for c in out.columns
        if any(k in str(c) for k in ["연령", "세별", "나이", "age", "분류2명", "분류3명", "C2_NM", "C3_NM", "ITM_NM"])
    ]
    for c in age_like_cols:
        s = out[c].astype(str).str.strip()
        if s.str.contains(r"\d+\s*세|세이상|세미만", na=False).any():
            mask = s.isin(total_like) | s.str.contains("전체|합계|계|총인구", na=False)
            if mask.any():
                out = out[mask].copy()
    return out


def calc_population(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
    result = {"available": False, "latest_month": base_month, "total": None, "mom_change": None, "households": None, "households_mom": None, "table_rows": [], "signal": "인구·세대수 원자료 미확인"}
    if not files:
        return result
    monthly: Dict[str, Dict[str, Any]] = {}
    try:
        for p in files:
            progress(f"인구·세대 파일 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            df = filter_df_by_region(df, region)
            mcol = find_best_month_col(df)
            vcol = find_value_col(df)
            item_col = find_col(df, ["항목명", "분류1명", "분류2명"])
            if not item_col:
                population_items = {"\ucd1d\uc778\uad6c\uc218", "\ub0a8\uc790\uc778\uad6c\uc218", "\uc5ec\uc790\uc778\uad6c\uc218"}
                for c in df.columns:
                    sample = set(df[c].dropna().astype(str).str.strip().head(500).unique())
                    if sample & population_items:
                        item_col = c
                        break
            if not vcol:
                continue
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            df["_value"] = to_numeric_safe(df[vcol])
            agg_method = "sum" if has_multiple_region_members(df) else "max"
            name = p.name
            if "주민등록인구" in name or ("인구" in name and "이동" not in name):
                if item_col:
                    df_pop = _filter_population_total_rows(df, item_col)
                    item = df_pop[item_col].astype(str)
                    total_mask = item.str.strip().isin({"\ucd1d\uc778\uad6c\uc218", "\ucd1d\uc778\uad6c"}) | item.str.contains("\ucd1d\uc778\uad6c", na=False)
                    sex_mask = item.str.contains("남자인구수|여자인구수|남자|여자", na=False)
                    if total_mask.any():
                        g = getattr(df_pop[total_mask & ~sex_mask].groupby("_month")["_value"], agg_method)().reset_index()
                    elif sex_mask.any():
                        g = df_pop[sex_mask].groupby("_month")["_value"].sum().reset_index()
                    else:
                        g = df_pop.groupby("_month")["_value"].sum().reset_index()
                else:
                    df_pop = _filter_population_total_rows(df, item_col)
                    g = df_pop.groupby("_month")["_value"].sum().reset_index()
                for _, r in g.iterrows():
                    if str(r["_month"]).startswith("20") and pd.notna(r["_value"]):
                        monthly.setdefault(r["_month"], {})["population"] = int(r["_value"])
            elif "세대" in name:
                g = getattr(df.groupby("_month")["_value"], agg_method)().reset_index()
                for _, r in g.iterrows():
                    if str(r["_month"]).startswith("20") and pd.notna(r["_value"]):
                        monthly.setdefault(r["_month"], {})["households"] = int(r["_value"])
        if not monthly:
            return result
        rows = [{"month": m, "population": monthly[m].get("population"), "households": monthly[m].get("households")} for m in sorted(monthly)]
        dfm = pd.DataFrame(rows).sort_values("month")
        dfm["population_mom"] = pd.to_numeric(dfm.get("population"), errors="coerce").diff()
        dfm["households_mom"] = pd.to_numeric(dfm.get("households"), errors="coerce").diff()
        row = dfm[dfm["month"] == base_month].iloc[-1] if base_month in dfm["month"].values else dfm.iloc[-1]
        prev = dfm[dfm["month"] < row["month"]].iloc[-1] if not dfm[dfm["month"] < row["month"]].empty else None
        pop, hh = row.get("population"), row.get("households")
        mom_pop = int(pop - prev.get("population")) if prev is not None and pd.notna(pop) and pd.notna(prev.get("population")) else None
        mom_hh = int(hh - prev.get("households")) if prev is not None and pd.notna(hh) and pd.notna(prev.get("households")) else None
        signal = "인구·세대수 보합"
        if mom_pop is not None and mom_pop < 0 and mom_hh is not None and mom_hh >= 0:
            signal = "인구 감소·세대수 보합"
        elif mom_pop is not None and mom_pop < 0:
            signal = "인구 감소"
        elif mom_pop is not None and mom_pop > 0:
            signal = "인구 증가"
        result.update({"available": True, "latest_month": row["month"], "total": None if pd.isna(pop) else int(pop), "mom_change": mom_pop, "households": None if pd.isna(hh) else int(hh), "households_mom": mom_hh, "table_rows_with_prev": dfm.tail(14).to_dict("records"), "table_rows": dfm.tail(13).to_dict("records"), "signal": signal})
        return result
    except Exception as e:
        result["signal"] = f"인구 계산 오류: {e}"
        return result


def calc_age_population(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
    """주민등록인구 1세/5세 원자료를 0~9세~70대 이상으로 재분류."""
    result = {
        "available": False,
        "latest_month": None,
        "monthly_rows": [],
        "age_rows": [],
        "signal": "연령대별 주민등록인구 원자료 미확인"
    }
    if not files:
        return result
    try:
        rows = []
        for p in files:
            if "주민등록인구" not in p.name and "인구" not in p.name:
                continue
            progress(f"연령대별 주민등록인구 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            df = filter_df_by_region(df, region)
            mcol = find_best_month_col(df)
            vcol = find_value_col(df)
            if not vcol:
                continue
            item_col = find_col(df, ["항목명", "분류1명"])
            age_col = find_item_col_by_keywords(df, ["0", "1", "5", "10", "20", "70", "세"])
            if not age_col:
                age_col = find_col(df, ["연령", "연령별", "분류2명", "분류3명"])
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            df["_value"] = to_numeric_safe(df[vcol])
            if item_col:
                item = df[item_col].astype(str)
                # 총인구수/계 행은 제외하고 성별 또는 총인구 1세별 행만 집계
                sex_mask = item.str.contains("남자인구수|여자인구수|총인구수|인구수", na=False)
                if sex_mask.any():
                    df = df[sex_mask].copy()
            if not age_col:
                continue
            df["_age_group"] = df[age_col].map(age_to_group)
            # 성별이 나뉘어 있으면 합산, 총인구수 행이면 그대로 합산되어도 월-연령대 단위로 정리됨
            g = df.groupby(["_month", "_age_group"])["_value"].sum().reset_index()
            for _, r in g.iterrows():
                if str(r["_month"]).startswith("20") and r["_age_group"] != "연령미상" and pd.notna(r["_value"]):
                    rows.append({"month": r["_month"], "age_group": r["_age_group"], "population": int(r["_value"])})
        if not rows:
            return result
        dfr = pd.DataFrame(rows)
        dfr = dfr.groupby(["month", "age_group"], as_index=False)["population"].sum()
        piv = dfr.pivot(index="month", columns="age_group", values="population").fillna(0).reset_index().sort_values("month")
        order = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
        for c in order:
            if c not in piv.columns:
                piv[c] = 0
        piv = piv[["month"] + order]
        # 기준월 이하 자료만 사용하고, 최근 13개월 범위로 확정
        piv_upto_base = piv[piv["month"] <= base_month].copy()
        if piv_upto_base.empty:
            piv_upto_base = piv.copy()
        piv_calc = piv_upto_base.tail(14).copy()
        piv_13 = piv_upto_base.tail(13).copy()

        latest = piv_13.iloc[-1]
        first_13 = piv_13.iloc[0]

        # 주민등록인구 파일에서 남자+여자와 총인구 행이 함께 잡혀 2배 산출되는 경우 보정
        try:
            pop_ref = calc_population(files, region, base_month).get("total")
            total_raw = sum(float(latest[c]) for c in order)
            if pop_ref and total_raw > pop_ref * 1.5:
                for c in order:
                    piv[c] = piv[c] / 2
                piv_upto_base = piv[piv["month"] <= base_month].copy()
                if piv_upto_base.empty:
                    piv_upto_base = piv.copy()
                piv_calc = piv_upto_base.tail(14).copy()
                piv_13 = piv_upto_base.tail(13).copy()
                latest = piv_13.iloc[-1]
                first_13 = piv_13.iloc[0]
        except Exception:
            pass

        total_latest = sum(float(latest[c]) for c in order)
        age_rows = []
        for c in order:
            latest_pop = int(round(float(latest[c])))
            # 차트명 그대로 최근 13개월 범위의 첫 월 대비 최신월 증감으로 계산
            change_13m = int(round(float(latest[c]) - float(first_13[c])))
            share = round(latest_pop / total_latest * 100, 1) if total_latest else 0
            note = "청년·소형 임대수요 참고" if c == "20대" else ("고령층 정주수요 참고" if c == "70대 이상" else ("가족형 실수요 참고" if c in ["30대", "40대"] else "주거수요 참고"))
            age_rows.append({
                "연령대": c,
                "최신월 인구": latest_pop,
                "비중": share,
                "13개월 증감": change_13m,
                "증감 기준월": str(first_13["month"]),
                "증감 최신월": str(latest["month"]),
                "시장 해석": note
            })
        result.update({
            "available": True,
            "latest_month": latest["month"],
            "change_base_month": first_13["month"],
            "monthly_rows_with_prev": piv_calc.to_dict("records"),
            "monthly_rows": piv_13.to_dict("records"),
            "age_rows": age_rows,
            "signal": f"연령대별 인구 구조 분석 완료 / 13개월 증감 기준: {first_13['month']}→{latest['month']}"
        })
        return result
    except Exception as e:
        result["signal"] = f"연령대별 인구 계산 오류: {e}"
        return result


def calc_migration(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
    result = {"available": False, "latest_month": None, "table_rows": [], "age_rows": [], "age_monthly_rows": [], "age_trend_rows": [], "net_migration": None, "main_outflow_age": None, "signal": "연령별 이동자수 원자료 미확인"}
    if not files:
        return result
    try:
        rows = []
        age_month_rows = []
        for p in files:
            progress(f"연령별 이동자수 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            df = filter_df_by_region(df, region)
            mcol = find_best_month_col(df)
            vcol = find_value_col(df)
            item_col = find_item_col_by_keywords(df, ["전입", "전출", "순이동", "이동자"])
            age_col = find_item_col_by_keywords(df, ["0~4", "5~9", "20", "70", "세"])
            if not vcol:
                continue
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            df["_value"] = to_numeric_safe(df[vcol])
            if item_col:
                df["_item"] = df[item_col].astype(str)
                for m, sub in df.groupby("_month"):
                    if not str(m).startswith("20"):
                        continue
                    val_in = sub[sub["_item"].str.contains("전입", na=False)]["_value"].sum()
                    val_out = sub[sub["_item"].str.contains("전출", na=False)]["_value"].sum()
                    val_net = sub[sub["_item"].str.contains("순이동", na=False)]["_value"].sum()
                    if val_net == 0 and (val_in or val_out):
                        val_net = val_in - val_out
                    rows.append({"month": m, "in": int(val_in) if pd.notna(val_in) else None, "out": int(val_out) if pd.notna(val_out) else None, "net": int(val_net) if pd.notna(val_net) else None})
                if age_col:
                    for (m, age), sub in df.groupby(["_month", age_col]):
                        if not str(m).startswith("20"):
                            continue
                        net = sub[sub["_item"].str.contains("순이동", na=False)]["_value"].sum()
                        if net == 0:
                            val_in_age = sub[sub["_item"].str.contains("전입", na=False)]["_value"].sum()
                            val_out_age = sub[sub["_item"].str.contains("전출", na=False)]["_value"].sum()
                            if val_in_age or val_out_age:
                                net = val_in_age - val_out_age
                        age_month_rows.append({"month": m, "age": str(age), "age_group": age_to_group(age), "net": int(net) if pd.notna(net) else 0})
            else:
                g = df.groupby("_month")["_value"].sum().reset_index()
                for _, r in g.iterrows():
                    rows.append({"month": r["_month"], "net": int(r["_value"]) if pd.notna(r["_value"]) else None})
        if not rows:
            return result
        dfr = pd.DataFrame(rows).groupby("month", as_index=False).agg({"in": "sum", "out": "sum", "net": "sum"}).sort_values("month")
        latest = dfr[dfr["month"] <= base_month].iloc[-1] if not dfr[dfr["month"] <= base_month].empty else dfr.iloc[-1]
        age_rows = []
        age_monthly_rows = []
        age_trend_rows = []
        main_out = None
        if age_month_rows:
            dfa = pd.DataFrame(age_month_rows)
            if "age_group" not in dfa.columns:
                dfa["age_group"] = dfa["age"].map(age_to_group)
            dfa = dfa[dfa["age_group"] != "연령미상"].copy()
            dfa_grouped = dfa.groupby(["month", "age_group"], as_index=False)["net"].sum().sort_values(["month", "age_group"])
            age_monthly_rows = dfa_grouped.rename(columns={"age_group": "age"}).tail(13 * 8).to_dict("records")

            latest_age = dfa_grouped[dfa_grouped["month"] == latest["month"]]
            latest_age = latest_age.rename(columns={"age_group": "age"}).sort_values("net")
            age_rows = latest_age.to_dict("records")
            if not latest_age.empty:
                main_out = str(latest_age.iloc[0]["age"])

            piv = dfa_grouped.pivot(index="month", columns="age_group", values="net").fillna(0).reset_index().sort_values("month")
            order = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
            for c in order:
                if c not in piv.columns:
                    piv[c] = 0
            piv = piv[["month"] + order]
            age_trend_rows = piv.tail(13).to_dict("records")

        result.update({"available": True, "latest_month": latest["month"], "table_rows": dfr.tail(13).to_dict("records"), "all_rows": dfr.to_dict("records"), "age_rows": age_rows, "age_monthly_rows": age_monthly_rows, "age_trend_rows": age_trend_rows, "net_migration": int(latest.get("net")) if pd.notna(latest.get("net")) else None, "main_outflow_age": main_out, "signal": "인구 이동 보조지표 확인"})
        return result
    except Exception as e:
        result["signal"] = f"이동자수 계산 오류: {e}"
        return result


def calc_rates(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
    result = {"available": False, "latest_month": None, "mortgage_rate": None, "jeonse_rate": None, "sme_rate": None, "business_rate": None, "table_rows": [], "signal": "금리 원자료 미확인"}
    if not files:
        return result
    rate_keys = {"mortgage": ["주택담보대출"], "jeonse": ["전세자금대출"], "sme": ["중소기업대출"], "business": ["기업일반자금대출", "기업일반자금"]}
    monthly: Dict[str, Dict[str, float]] = {}
    try:
        for p in files:
            progress(f"금리 파일 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            mcol, vcol = find_best_month_col(df), find_supply_value_col(df)
            item_col = find_item_col_by_keywords(df, ["주택담보대출", "전세자금대출", "중소기업대출", "기업일반자금"])
            if not vcol:
                continue
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            df["_value"] = to_numeric_safe(df[vcol])
            if item_col:
                df["_item"] = df[item_col].astype(str)
                for key, kws in rate_keys.items():
                    sub = pd.DataFrame()
                    for kw in kws:
                        exact = df[df["_item"].str.strip() == kw]
                        if not exact.empty:
                            sub = exact
                            break
                    if sub.empty:
                        sub = df[df["_item"].apply(lambda x: any(k in x for k in kws))]
                    g = sub.groupby("_month")["_value"].max().reset_index()
                    for _, r in g.iterrows():
                        if str(r["_month"]).startswith("20") and pd.notna(r["_value"]):
                            monthly.setdefault(r["_month"], {})[key] = float(r["_value"])
        if not monthly:
            return result
        rows = [{"month": m, "mortgage_rate": monthly[m].get("mortgage"), "jeonse_rate": monthly[m].get("jeonse"), "sme_rate": monthly[m].get("sme"), "business_rate": monthly[m].get("business")} for m in sorted(monthly)]
        dfm = pd.DataFrame(rows).sort_values("month")
        for col in ["mortgage_rate", "jeonse_rate", "sme_rate", "business_rate"]:
            dfm[f"{col}_mom"] = pd.to_numeric(dfm.get(col), errors="coerce").diff().round(2)
        row = dfm[dfm["month"] <= base_month].iloc[-1] if not dfm[dfm["month"] <= base_month].empty else dfm.iloc[-1]
        signal = "주거 금융부담 지속" if pd.notna(row.get("mortgage_rate")) and row.get("mortgage_rate") >= 4 else "주거 금융부담 점검"
        result.update({"available": True, "latest_month": row["month"], "mortgage_rate": row.get("mortgage_rate"), "jeonse_rate": row.get("jeonse_rate"), "sme_rate": row.get("sme_rate"), "business_rate": row.get("business_rate"), "table_rows_with_prev": dfm.tail(14).to_dict("records"), "table_rows": dfm.tail(13).to_dict("records"), "signal": signal})
        return result
    except Exception as e:
        result["signal"] = f"금리 계산 오류: {e}"
        return result


def calc_unsold(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
    result = {"available": False, "latest_month": None, "unsold_units": None, "after_completion_unsold": None, "table_rows": [], "signal": "미분양 원자료 미확인"}
    if not files:
        return result
    monthly = aggregate_long_value(files, region, base_month, {"준공후미분양": "after", "미분양": "unsold"})
    if not monthly:
        return result
    rows = [{"month": m, "unsold_units": monthly[m].get("unsold"), "after_completion_unsold": monthly[m].get("after")} for m in sorted(monthly)]
    dfm = pd.DataFrame(rows).sort_values("month")
    dfm["unsold_units_mom"] = pd.to_numeric(dfm.get("unsold_units"), errors="coerce").diff()
    dfm["after_completion_unsold_mom"] = pd.to_numeric(dfm.get("after_completion_unsold"), errors="coerce").diff()
    row = dfm[dfm["month"] <= base_month].iloc[-1] if not dfm[dfm["month"] <= base_month].empty else dfm.iloc[-1]
    signal = "분양시장 부담 점검"
    if pd.notna(row.get("after_completion_unsold")) and row.get("after_completion_unsold") > 0:
        signal = "준공후 미분양 부담 잔존"
    elif pd.notna(row.get("unsold_units")) and row.get("unsold_units") > 0:
        signal = "미분양 부담 잔존"
    result.update({"available": True, "latest_month": row["month"], "unsold_units": None if pd.isna(row.get("unsold_units")) else int(row.get("unsold_units")), "after_completion_unsold": None if pd.isna(row.get("after_completion_unsold")) else int(row.get("after_completion_unsold")), "table_rows_with_prev": dfm.tail(14).to_dict("records"), "table_rows": dfm.tail(13).to_dict("records"), "signal": signal})
    return result


def _prefer_total_price_index_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Prefer headline total rows in KOSIS price-index tables."""
    if df is None or df.empty:
        return df
    total_tokens = {"전체", "소계", "계", "종합", "total", "all"}
    target_cols = []
    for c in df.columns:
        name = str(c)
        if name in {
            "C2_NM", "C3_NM", "OBJ_L2_NM", "OBJ_L3_NM",
            "region_level_2", "region_level_3", "obj_level_2", "obj_level_3",
            "분류2명", "분류3명", "분류2", "분류3",
        }:
            target_cols.append(c)
            continue
        if any(k in name for k in ["규모", "면적", "size"]):
            target_cols.append(c)
    out = df
    for c in target_cols:
        if c not in out.columns:
            continue
        s = out[c].astype(str).str.strip()
        has_size_band = s.str.contains(r"㎡|m²|m2|초과|이하|\d+\s*평", na=False, regex=True).any()
        has_total = s.str.lower().isin(total_tokens).any()
        if has_total and has_size_band:
            narrowed = out[s.str.lower().isin(total_tokens)].copy()
            if not narrowed.empty:
                out = narrowed
    return out


def _price_index_region_name_cols(df: pd.DataFrame) -> List[Any]:
    """가격지수 원자료에서 '지역' 분류명 컬럼을 찾는다.

    전국주택가격동향조사 유형별 가격지수(DT_30404_C001/C012/C013)는 분류1이 주택유형,
    분류2가 지역이므로 분류1명을 지역으로 보면 유형 4개가 '지역 혼재'로 오판된다.
    분류 객체명(지역별/Region)으로 지역 컬럼을 판정하고, 정보가 없을 때만 분류1을 사용한다.
    """
    pairs = [
        ("C1_NM", ["C1_OBJ_NM", "C1_OBJ_NM_ENG"]),
        ("분류1명", ["분류1_객체명", "C1_OBJ_NM_ENG"]),
        ("OBJ_L1_NM", ["OBJ_L1_OBJ_NM"]),
        ("C2_NM", ["C2_OBJ_NM", "C2_OBJ_NM_ENG"]),
        ("분류2명", ["분류2_객체명", "C2_OBJ_NM_ENG"]),
        ("OBJ_L2_NM", ["OBJ_L2_OBJ_NM"]),
    ]
    region_cols = []
    for name_col, obj_cols in pairs:
        if name_col not in df.columns:
            continue
        for obj_col in obj_cols:
            if obj_col not in df.columns:
                continue
            obj_text = " ".join(df[obj_col].astype(str).str.strip().unique().tolist())
            if "지역" in obj_text or "region" in obj_text.lower():
                region_cols.append(name_col)
                break
    if region_cols:
        return region_cols
    return [c for c in df.columns if str(c) in {"C1_NM", "분류1명", "OBJ_L1_NM"}]


def is_sido_public_stat_region(region: str) -> bool:
    """공표통계 라벨용 시도 판정. '전남광주통합특별시 광주권역'(시도단위 체크) 포함."""
    r = " ".join(str(region or "").split())
    return is_sido_level_region(r) or r.endswith("광주권역")


def calc_price_index(files: List[Path], region: str, base_month: str, index_type: str) -> Dict[str, Any]:
    result = {"available": False, "latest_month": None, "index_type": index_type, "latest_value": None, "mom_change": None, "table_rows": [], "signal": "가격지수 원자료 미확인"}
    if not files:
        return result
    try:
        rows = []
        asset_labels = []
        for p in files:
            progress(f"{index_type} 가격지수 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            df = filter_df_by_region(df, region)
            region_name_cols = _price_index_region_name_cols(df)
            if region_name_cols:
                distinct_regions = df[region_name_cols[0]].astype(str).str.strip().nunique()
                if distinct_regions > 3:
                    progress(f"{index_type} 가격지수 지역 매칭 실패로 제외: {p.name} (지역 {distinct_regions}개 혼재)")
                    continue
                # 시군구 보고서인데 시도 행(예: 대구)과 시군구 행(예: 수성)이 함께 남으면, 월별 max로 두 지역 값이 섞이므로
                # 시군구 행만 쓴다. 시군구 행이 없을 때만 시도 행을 그대로 둔다.
                tokens = str(region or "").split()
                if distinct_regions > 1 and len(tokens) >= 2:
                    sigungu = tokens[-1]
                    stem = re.sub(r"(구|시|군)$", "", sigungu)
                    names = df[region_name_cols[0]].astype(str).str.strip()
                    local = names.isin({sigungu, stem}) | (names.str.contains(stem, regex=False) if len(stem) >= 2 else False)
                    if local.any():
                        df = df[local].copy()
            df = _prefer_total_price_index_rows(df)
            apt_cols = [
                c for c in df.columns
                if any(k in str(c) for k in ["유형", "주택유형", "분류", "항목", "C1_NM", "C2_NM", "OBJ_L1_NM", "OBJ_L2_NM"])
            ]
            if apt_cols:
                apt_mask = pd.Series(False, index=df.index)
                for c in apt_cols:
                    apt_mask = apt_mask | df[c].astype(str).str.contains("아파트", na=False, regex=False)
                if apt_mask.any():
                    df = df[apt_mask].copy()
            # 표에 '어떤 지수의 변동률인지' 밝히기 위해 실제로 쓴 주택유형을 기록한다.
            if apt_cols and apt_mask.any():
                asset_labels.append("아파트")
            elif "오피스텔" in p.name:
                asset_labels.append("오피스텔")
            else:
                asset_labels.append("종합")
            mcol, vcol = find_best_month_col(df), find_value_col(df)
            if not vcol:
                continue
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            
            # Extract both Index and Change Rate if present (from API with itmId=ALL)
            itm_col = None
            for c in df.columns:
                sample = "".join(df[c].astype(str).head(10).tolist())
                if "지수" in sample or "증감" in sample or "변동" in sample or "가격" in sample:
                    itm_col = c
                    break
            
            if itm_col:
                idx_mask = df[itm_col].astype(str).str.contains("지수|가격", na=False)
                rate_mask = df[itm_col].astype(str).str.contains("증감|변동", na=False)
                df_idx = df[idx_mask].copy() if idx_mask.any() else df.copy()
                df_rate = df[rate_mask].copy() if rate_mask.any() else pd.DataFrame()
            else:
                df_idx = df.copy()
                df_rate = pd.DataFrame()
                
            df_idx["_value"] = to_numeric_safe(df_idx[vcol])
            g = df_idx.groupby("_month")["_value"].max().reset_index()
            for _, r in g.iterrows():
                if str(r["_month"]).startswith("20") and pd.notna(r["_value"]):
                    rows.append({"month": r["_month"], "value": float(r["_value"])})
                    
            if not df_rate.empty:
                if "rate_rows" not in locals(): rate_rows = []
                df_rate["_rate_value"] = to_numeric_safe(df_rate[vcol])
                gr = df_rate.groupby("_month")["_rate_value"].max().reset_index()
                for _, r in gr.iterrows():
                    if str(r["_month"]).startswith("20") and pd.notna(r["_rate_value"]):
                        rate_rows.append({"month": r["_month"], "rate": float(r["_rate_value"])})

        if not rows:
            return result
        dfm = pd.DataFrame(rows).groupby("month", as_index=False)["value"].max().sort_values("month")
        
        # Merge official rate if available
        if "rate_rows" in locals() and rate_rows:
            dfr = pd.DataFrame(rate_rows).groupby("month", as_index=False)["rate"].max()
            dfm = pd.merge(dfm, dfr, on="month", how="left")
            dfm["public_index_change_rate"] = dfm["rate"].round(2)
            dfm["public_index_change_rate"] = dfm["public_index_change_rate"].fillna(dfm["value"].pct_change() * 100).round(2)
        else:
            # Fallback to manual calculation
            dfm["public_index_change_rate"] = dfm["value"].pct_change() * 100
            dfm["public_index_change_rate"] = dfm["public_index_change_rate"].round(2)
            
        # 기준시점 개편월도 전월과 동일 기준으로 환산된 원지수가 함께 제공되므로
        # 특정 지역의 변동률을 하드코딩하지 않고 원지수 전월비를 그대로 사용한다.
        base_key = str(base_month or "")[:7]
        published_months = set(dfm["month"].astype(str).str[:7].tolist())
        published_until_base = dfm[dfm["month"].astype(str).str[:7] <= base_key] if base_key else dfm
        if base_key and base_key not in published_months and not published_until_base.empty:
            latest_published_month = str(published_until_base.iloc[-1]["month"])[:7]
            if latest_published_month < base_key:
                dfm = pd.concat([
                    dfm,
                    pd.DataFrame([{
                        "month": base_key,
                        "value": None,
                        "public_index_change_rate": None,
                        "publication_status": "미공표",
                    }]),
                ], ignore_index=True).sort_values("month")
        row = dfm[dfm["month"] <= base_month].iloc[-1] if not dfm[dfm["month"] <= base_month].empty else dfm.iloc[-1]
        mom = None if pd.isna(row.get("public_index_change_rate")) else round(float(row.get("public_index_change_rate")), 2)
        signal = f"{index_type} 가격지수 보합"
        if str(row.get("publication_status", "")) == "미공표":
            signal = f"{index_type} 가격지수 미공표"
        elif mom is not None and mom > 0:
            signal = f"{index_type} 가격지수 상승"
        elif mom is not None and mom < 0:
            signal = f"{index_type} 가격지수 하락"
        latest_value = None if pd.isna(row.get("value")) else float(row["value"])
        result.update({"available": True, "latest_month": row["month"], "latest_value": latest_value, "mom_change": mom, "table_rows_with_prev": dfm.tail(14).to_dict("records"), "table_rows": dfm.tail(13).to_dict("records"), "signal": signal,
                       "asset_label": asset_labels[0] if asset_labels else ""})
        return result
    except Exception as e:
        result["signal"] = f"가격지수 계산 오류: {e}"
        return result


def attach_public_index_change_rate(monthly_rows: List[Dict[str, Any]], price_index_part: Dict[str, Any]) -> List[Dict[str, Any]]:
    """13개월 실거래 시계열표에 KOSIS 주택동향지수 기반 아파트 변동률을 '공표통계' 컬럼으로 병합.

    - 월 매칭 기준: month
    - 공표통계 산식: (당월 KOSIS 가격지수 / 전월 KOSIS 가격지수 - 1) × 100
    - 표 표시 위치: 평균 변동률 바로 다음 컬럼
    """
    rows = list(monthly_rows or [])
    index_rows = price_index_part.get("table_rows", []) if isinstance(price_index_part, dict) else []
    public_stat_label = price_index_part.get("public_stat_label", "공표통계") if isinstance(price_index_part, dict) else "공표통계"
    asset_label = str(price_index_part.get("asset_label") or "").strip() if isinstance(price_index_part, dict) else ""
    if asset_label and f"({asset_label})" not in public_stat_label:
        public_stat_label = f"{public_stat_label}({asset_label})"
    lookup = {}
    for r in index_rows:
        m = str(r.get("month", "")).strip()
        if not m:
            continue
        v = r.get("public_index_change_rate")
        try:
            if v is not None and not pd.isna(v):
                lookup[m] = round(float(v), 2)
        except Exception:
            continue

    out = []
    for r in rows:
        src = dict(r)
        month = str(src.get("month", "")).strip()
        pub_v = lookup.get(month)
        new_row = {}
        inserted = False
        for k, v in src.items():
            new_row[k] = v
            if k == "avg_change_rate":
                new_row["public_index_change_rate"] = pub_v
                new_row["_public_stat_label"] = public_stat_label
                inserted = True
        if not inserted:
            new_row["public_index_change_rate"] = pub_v
            new_row["_public_stat_label"] = public_stat_label
        out.append(new_row)
    return out





def infer_report_mode(classified: Dict[str, List[Path]], report_version: str = "") -> str:
    """보고서 모드 자동 판정.
    - 주간아파트 변동률 파일이 있거나 버전명에 주간/weekly가 있으면 weekly
    - 오피스텔 가격지수·실거래 파일이 있거나 버전명에 오피스텔/officetel이 있으면 officetel
    - 그 외는 monthly
    """
    v = str(report_version or "").lower()
    if "report_mode=weekly" in v:
        return "weekly"
    if "report_mode=monthly" in v:
        return "monthly"
    if "report_mode=officetel" in v:
        return "officetel"
    if "주간" in v or "weekly" in v:
        return "weekly"
    if classified.get("weekly_apt_sale_change") or classified.get("weekly_apt_jeonse_change"):
        return "weekly"
    if "오피스텔" in str(report_version or "") or "officetel" in v:
        return "officetel"
    mode_files = (
        classified.get("apt_trade", [])
        + classified.get("apt_rent", [])
        + classified.get("price_index_sale", [])
        + classified.get("price_index_jeonse", [])
        + classified.get("price_index_rent", [])
    )
    if mode_files and any("오피스텔" in p.name for p in mode_files):
        return "officetel"
    return "monthly"


def is_weekly_report(data: Dict[str, Any]) -> bool:
    return str(data.get("meta", {}).get("report_mode", "monthly")) == "weekly"


def normalize_week_value(x: Any) -> str:
    """KOSIS 주간 공표일을 YYYY-MM-DD 형식으로 정규화."""
    if pd.isna(x):
        return ""
    try:
        if isinstance(x, (pd.Timestamp, datetime)):
            return pd.to_datetime(x).strftime("%Y-%m-%d")
    except Exception:
        pass
    s = str(x).strip()
    nums = re.sub(r"[^0-9]", "", s)
    if len(nums) >= 8:
        return f"{nums[:4]}-{nums[4:6]}-{nums[6:8]}"
    if len(nums) >= 6:
        return f"{nums[:4]}-{nums[4:6]}"
    return s


def build_weekly_report_period_info(base_month: str, *week_values: Any) -> Dict[str, Any]:
    """주간보고서 제목에 표시할 기준월 내 월요일 주차 정보를 계산.

    제목 표시는 수집된 최신 공표일보다 보고서 생성일/종료 기준월의 최종 월요일을 우선한다.
    예: 2026-05 보고서를 2026-05-25에 생성하면 2026년 5월 4주차로 표기한다.
    """
    base_key = str(base_month or "")[:7]
    try:
        month_start = datetime.strptime(base_key + "-01", "%Y-%m-%d")
    except Exception:
        return {
            "weekly_reference_date": "",
            "weekly_week_of_month": None,
            "weekly_period_label": "",
            "weekly_title_suffix": "",
        }

    today = datetime.now()
    if today.strftime("%Y-%m") == base_key:
        reference_date = today - timedelta(days=today.weekday())
    else:
        month_end = (pd.Timestamp(month_start) + pd.offsets.MonthEnd(0)).to_pydatetime()
        reference_date = month_end - timedelta(days=month_end.weekday())

    first_monday = month_start + timedelta(days=(7 - month_start.weekday()) % 7)
    week_of_month = max(1, int(((reference_date.date() - first_monday.date()).days // 7) + 1))
    period_label = f"{reference_date.year}년 {reference_date.month}월 {week_of_month}주차"
    reference_text = reference_date.strftime("%Y-%m-%d")
    return {
        "weekly_reference_date": reference_text,
        "weekly_week_of_month": week_of_month,
        "weekly_period_label": period_label,
        "weekly_title_suffix": f"{period_label}({reference_text} 월요일 기준)",
    }


def compress_weekly_rows_for_mixed_monthly_current(dfw: pd.DataFrame, base_month: str) -> pd.DataFrame:
    """이전 월은 월말 월요일 1건, 기준월은 주간 월요일 전체를 남긴다."""
    if dfw is None or dfw.empty or "week" not in dfw.columns:
        return dfw
    work = dfw.copy().sort_values("week")
    work["_month_key"] = work["month"].astype(str).str[:7]
    base_key = str(base_month or "")[:7]
    prev = work[work["_month_key"] < base_key].copy()
    current = work[work["_month_key"] == base_key].copy()
    if not prev.empty:
        prev = prev.groupby("_month_key", as_index=False, group_keys=False).tail(1)
    mixed = pd.concat([prev, current], ignore_index=True) if not current.empty else prev
    if mixed.empty:
        mixed = work.groupby("_month_key", as_index=False, group_keys=False).tail(1)
    mixed = mixed.sort_values("week").drop(columns=["_month_key"], errors="ignore")
    mixed["period_label"] = mixed.apply(
        lambda r: str(r.get("week", "")) if str(r.get("month", ""))[:7] == base_key else str(r.get("month", ""))[:7],
        axis=1,
    )
    mixed["period_type"] = mixed["month"].astype(str).str[:7].map(lambda m: "주간" if m == base_key else "월간")
    return mixed.tail(18)


def calc_weekly_price_change(files: List[Path], region: str, base_month: str, index_type: str) -> Dict[str, Any]:
    """KOSIS 주간아파트 매매·전세가격 변동률 파일을 공표통계로 계산.

    주간 표는 이미 변동률 자체가 값이므로 월간 가격지수처럼 pct_change를 다시 계산하지 않는다.
    """
    result = {"available": False, "latest_month": None, "latest_week": None, "index_type": index_type, "latest_value": None, "mom_change": None, "table_rows": [], "signal": "주간 공표통계 원자료 미확인"}
    if not files:
        return result
    try:
        rows = []
        for p in files:
            progress(f"{index_type} 주간 아파트 공표통계 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            df = filter_df_by_region(df, region)
            mcol, vcol = find_best_month_col(df), find_value_col(df)
            if not vcol:
                continue
            if mcol:
                df["_week"] = df[mcol].map(normalize_week_value)
                df["_month"] = df["_week"].map(lambda x: str(x)[:7] if str(x).startswith("20") else normalize_month_value(x))
            else:
                df["_week"] = base_month
                df["_month"] = base_month
            df["_value"] = to_numeric_safe(df[vcol])
            g = df.groupby(["_week", "_month"], as_index=False)["_value"].max()
            for _, r in g.iterrows():
                if str(r["_week"]).startswith("20") and pd.notna(r["_value"]):
                    rows.append({"week": r["_week"], "month": r["_month"], "public_index_change_rate": round(float(r["_value"]), 2), "value": round(float(r["_value"]), 2)})
        if not rows:
            return result
        dfw = pd.DataFrame(rows).sort_values("week")
        target = dfw[dfw["month"] <= base_month]
        row = target.iloc[-1] if not target.empty else dfw.iloc[-1]
        mom = round(float(row["public_index_change_rate"]), 2)
        signal = f"{index_type} 주간 아파트 가격변동률 보합"
        if mom > 0:
            signal = f"{index_type} 주간 아파트 가격변동률 상승"
        elif mom < 0:
            signal = f"{index_type} 주간 아파트 가격변동률 하락"
        display_rows = compress_weekly_rows_for_mixed_monthly_current(dfw, base_month)
        result.update({"available": True, "latest_month": row["month"], "latest_week": row["week"], "latest_value": mom, "mom_change": mom, "table_rows": display_rows.to_dict("records"), "raw_weekly_rows": dfw.tail(60).to_dict("records"), "signal": signal})
        return result
    except Exception as e:
        result["signal"] = f"주간 공표통계 계산 오류: {e}"
        return result


def combine_monthly_index_with_weekly_current(monthly_part: Dict[str, Any], weekly_part: Dict[str, Any], base_month: str, index_type: str) -> Dict[str, Any]:
    """주간보고서용 공표통계 시계열.

    - 과거 월: 월간 주택가격동향 아파트 가격지수의 전월 대비 변동률
    - 기준월: 주간 아파트 가격변동률의 월요일 공표값
    """
    if not isinstance(monthly_part, dict) or not monthly_part.get("available"):
        out = dict(weekly_part or {})
        out["public_stat_label"] = "시군구 공표통계"
        return out
    if not isinstance(weekly_part, dict) or not weekly_part.get("available"):
        out = dict(monthly_part or {})
        out["public_stat_label"] = "시군구 공표통계"
        return out

    base_key = str(base_month or "")[:7]
    rows = []
    for r in monthly_part.get("table_rows", []) or []:
        month = str(r.get("month", "")).strip()[:7]
        if not month or month >= base_key:
            continue
        v = r.get("public_index_change_rate")
        try:
            v = None if v is None or pd.isna(v) else round(float(v), 2)
        except Exception:
            v = None
        rows.append({
            "month": month,
            "period_label": month,
            "period_type": "월간",
            "source": "월간동향 공표통계",
            "value": r.get("value"),
            "public_index_change_rate": v,
        })

    current_weekly_rows = []
    for r in weekly_part.get("raw_weekly_rows", []) or weekly_part.get("table_rows", []) or []:
        month = str(r.get("month", "")).strip()[:7]
        if month != base_key:
            continue
        v = r.get("public_index_change_rate")
        try:
            v = None if v is None or pd.isna(v) else round(float(v), 2)
        except Exception:
            v = None
        current_weekly_rows.append({
            "week": r.get("week"),
            "month": month,
            "period_label": r.get("week") or month,
            "period_type": "주간",
            "source": "주간동향 공표통계",
            "value": v,
            "public_index_change_rate": v,
        })
    current_weekly_rows = sorted(current_weekly_rows, key=lambda x: str(x.get("week", "")))
    if not current_weekly_rows:
        current_weekly_rows.append({
            "week": weekly_part.get("latest_week") or "",
            "month": base_key,
            "period_label": base_key,
            "period_type": "주간",
            "source": "주간동향 공표통계",
            "value": None,
            "public_index_change_rate": None,
            "publication_status": "미공표",
        })
    rows.extend(current_weekly_rows)

    if not rows:
        return weekly_part

    rows = rows[-18:]
    latest_weekly = current_weekly_rows[-1] if current_weekly_rows else {}
    latest_value = latest_weekly.get("public_index_change_rate")
    if latest_value is None and rows:
        latest_value = rows[-1].get("public_index_change_rate")
    signal = f"{index_type} 월간지수+주간 공표통계 보합"
    if latest_value is not None and latest_value > 0:
        signal = f"{index_type} 월간지수+주간 공표통계 상승"
    elif latest_value is not None and latest_value < 0:
        signal = f"{index_type} 월간지수+주간 공표통계 하락"

    out = dict(weekly_part)
    out.update({
        "available": True,
        "public_stat_label": "시군구 공표통계",
        "latest_month": latest_weekly.get("month") or weekly_part.get("latest_month") or monthly_part.get("latest_month"),
        "latest_week": latest_weekly.get("week") or weekly_part.get("latest_week"),
        "latest_value": latest_value,
        "mom_change": latest_value,
        "table_rows": rows,
        "monthly_index_rows": monthly_part.get("table_rows", []),
        "weekly_source_rows": weekly_part.get("raw_weekly_rows", []) or weekly_part.get("table_rows", []),
        "mixed_source_note": "과거 월은 월간동향 아파트 가격지수 전월 대비 변동률, 기준월은 주간동향 아파트 가격변동률을 사용",
        "signal": signal,
    })
    return out


def get_supply_sido_keywords(region: str) -> List[str]:
    """주택건설 인허가·착공·준공 KOSIS 공급통계 필터 기준.
    공급 파일은 시군구 단위가 아니라 시도 단위가 많으므로, 목포시·무안군·나주시 등은 전라남도(전남)로 필터링한다.
    """
    r = str(region or "")
    mapping = [
        (["광주", "광주광역시"], ["광주", "광주광역시"]),
        (["전라남도", "전남", "목포", "나주", "무안", "순천", "여수", "광양", "담양", "장성", "화순", "영암", "해남", "완도", "진도", "신안", "강진", "고흥", "보성", "장흥", "함평", "영광", "곡성", "구례"], ["전남", "전라남도"]),
        (["전라북도", "전북", "전주", "군산", "익산", "정읍", "남원", "김제", "완주", "진안", "무주", "장수", "임실", "순창", "고창", "부안"], ["전북", "전라북도"]),
        (["대구", "대구광역시"], ["대구", "대구광역시"]),
        (["서울", "서울특별시"], ["서울", "서울특별시"]),
        (["부산", "부산광역시"], ["부산", "부산광역시"]),
        (["인천", "인천광역시"], ["인천", "인천광역시"]),
        (["대전", "대전광역시"], ["대전", "대전광역시"]),
        (["울산", "울산광역시"], ["울산", "울산광역시"]),
        (["세종", "세종특별자치시"], ["세종", "세종특별자치시"]),
        (["경기도", "경기"], ["경기", "경기도"]),
        (["강원", "강원특별자치도", "강원도"], ["강원", "강원도", "강원특별자치도"]),
        (["충청북도", "충북"], ["충북", "충청북도"]),
        (["충청남도", "충남"], ["충남", "충청남도"]),
        (["경상북도", "경북"], ["경북", "경상북도"]),
        (["경상남도", "경남"], ["경남", "경상남도"]),
        (["제주", "제주특별자치도"], ["제주", "제주특별자치도"]),
    ]
    for triggers, keywords in mapping:
        if any(t in r for t in triggers):
            return keywords
    # fallback: 시도명이 포함된 경우 앞쪽 토큰 우선
    first = r.split()[0] if r.split() else r
    return [first] if first else []


def get_supply_sido_display_name(region: str) -> str:
    """보고서 HTML 표시용 공급권역명.
    시군구 보고서라도 인허가·착공·준공 공급통계는 시도 단위 기준으로 표시한다.
    """
    r = str(region or "")
    mapping = [
        (["서울", "서울특별시"], "서울특별시"),
        (["부산", "부산광역시"], "부산광역시"),
        (["대구", "대구광역시"], "대구광역시"),
        (["인천", "인천광역시"], "인천광역시"),
        (["광주", "광주광역시"], "광주광역시"),
        (["대전", "대전광역시"], "대전광역시"),
        (["울산", "울산광역시"], "울산광역시"),
        (["세종", "세종특별자치시"], "세종특별자치시"),
        (["경기도", "경기"], "경기도"),
        (["강원", "강원특별자치도", "강원도"], "강원특별자치도"),
        (["충청북도", "충북"], "충청북도"),
        (["충청남도", "충남"], "충청남도"),
        (["전라북도", "전북", "전주", "군산", "익산", "정읍", "남원", "김제", "완주", "진안", "무주", "장수", "임실", "순창", "고창", "부안"], "전북특별자치도"),
        (["전라남도", "전남", "목포", "나주", "무안", "순천", "여수", "광양", "담양", "장성", "화순", "영암", "해남", "완도", "진도", "신안", "강진", "고흥", "보성", "장흥", "함평", "영광", "곡성", "구례"], "전라남도"),
        (["경상북도", "경북"], "경상북도"),
        (["경상남도", "경남"], "경상남도"),
        (["제주", "제주특별자치도"], "제주특별자치도"),
    ]
    for triggers, name in mapping:
        if any(t in r for t in triggers):
            return name
    return r.split()[0] if r.split() else r


def get_region_short_name(region: str) -> str:
    parts = str(region or "").split()
    return parts[-1] if parts else str(region or "")


SIDO_LEVEL_NAMES = {
    "서울특별시", "서울", "부산광역시", "부산", "대구광역시", "대구", "인천광역시", "인천",
    "광주광역시", "광주", "대전광역시", "대전", "울산광역시", "울산", "세종특별자치시", "세종",
    "경기도", "경기", "강원특별자치도", "강원도", "강원", "충청북도", "충북", "충청남도", "충남",
    "전북특별자치도", "전라북도", "전북", "전라남도", "전남", "경상북도", "경북",
    "경상남도", "경남", "제주특별자치도", "제주도", "제주",
}


def is_sido_level_region(region: str) -> bool:
    """보고서 조사지역이 시군구가 아닌 시도 단위인지 판단."""
    r = str(region or "").strip()
    if not r:
        return False
    if r in SIDO_LEVEL_NAMES:
        return True
    parts = r.split()
    return len(parts) == 1 and any(r.endswith(suffix) for suffix in ["특별시", "광역시", "특별자치시", "특별자치도", "도"])


_SIGUNGU_CODE_NAME_CACHE: Optional[Dict[str, str]] = None


def load_sigungu_code_name_map() -> Dict[str, str]:
    """법정동 코드 파일에서 RTMS LAWD_CD/sggCd 5자리 → 시군구명 매핑을 만든다."""
    global _SIGUNGU_CODE_NAME_CACHE
    if _SIGUNGU_CODE_NAME_CACHE is not None:
        return _SIGUNGU_CODE_NAME_CACHE

    mapping: Dict[str, str] = {}
    search_roots = [Path.cwd(), Path(__file__).resolve().parent, Path(__file__).resolve().parent.parent]
    db_roots = list(search_roots)
    if getattr(sys, "frozen", False):
        db_roots.insert(0, Path(sys.executable).resolve().parent)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        db_roots.insert(0, Path(meipass))
    for root in db_roots:
        db_path = root / "local_db" / "region_codes.sqlite3"
        if not db_path.exists():
            continue
        try:
            with sqlite3.connect(str(db_path)) as conn:
                for code, name in conn.execute("SELECT REG, 시군구 FROM region_codes"):
                    code = re.sub(r"[^0-9]", "", str(code or "")).zfill(5)[:5]
                    name = str(name or "").strip()
                    if code and name:
                        mapping.setdefault(code, name)
            if mapping:
                _SIGUNGU_CODE_NAME_CACHE = mapping
                return mapping
        except Exception:
            continue
    candidates: List[Path] = []
    for root in search_roots:
        for name in ["전국 법정동 지역.xls", "전국 법정동 지역.xlsx"]:
            p = root / name
            if p.exists() and p not in candidates:
                candidates.append(p)
    if not candidates:
        for root in search_roots:
            try:
                candidates.extend([p for p in root.glob("*법정동*지역*.xls*") if p not in candidates])
            except Exception:
                pass

    for p in candidates:
        try:
            df = pd.read_excel(p)
            df = normalize_columns(df)
            reg_col = find_col(df, ["REG", "법정동코드", "지역코드", "코드"])
            sigungu_col = find_col(df, ["시군구", "시군구명"])
            if not reg_col or not sigungu_col:
                continue
            for _, r in df.iterrows():
                code = re.sub(r"[^0-9]", "", str(r.get(reg_col, ""))).zfill(5)[:5]
                name = str(r.get(sigungu_col, "") or "").strip()
                if code and name and name.lower() not in ["nan", "none"]:
                    mapping.setdefault(code, name)
            if mapping:
                break
        except Exception:
            continue

    _SIGUNGU_CODE_NAME_CACHE = mapping
    return mapping


def sigungu_name_from_code(code: Any) -> str:
    nums = re.sub(r"[^0-9]", "", str(code or ""))
    if len(nums) < 5:
        return ""
    return load_sigungu_code_name_map().get(nums[:5], "")


def normalize_sigungu_name_value(value: Any) -> str:
    s = str(value or "").strip()
    if not s or s.lower() in ["nan", "none"]:
        return ""
    if re.fullmatch(r"\d+(\.0)?", s):
        mapped = sigungu_name_from_code(s)
        return mapped or ""
    return s


def compose_sido_dong_label(sigungu: Any, dong: Any) -> str:
    sg = normalize_sigungu_name_value(sigungu)
    dg = str(dong or "").strip()
    if dg.lower() in ["nan", "none"]:
        dg = ""
    if sg and dg and not dg.startswith(sg):
        return f"{sg} {dg}"
    return dg or sg


def clean_region_part(value: Any) -> str:
    s = str(value or "").strip()
    return "" if not s or s.lower() in ["nan", "none"] else s


def normalize_dong_to_eupmyeondong(value: Any) -> str:
    """보고서 분석 단위를 읍·면·동까지로 맞춘다.

    실거래 원자료의 읍면동 열은 읍·면 지역에서 '금천면 석전리', '남평읍 동사리'처럼
    리 단위까지 표기되므로, 표·차트·서술의 지역 단위가 리로 내려가지 않도록
    상위 읍·면으로 합산한다. 동 지역('빛가람동')과 이미 읍·면 단위인 값은 그대로 둔다.
    읍·면 정보 없이 리 명칭만 있는 값은 상위 단위를 알 수 없으므로 원값을 유지한다.
    """
    s = clean_region_part(value)
    if not s:
        return ""
    parts = s.split()
    if len(parts) == 1:
        return s
    for token in parts:
        if token.endswith(("읍", "면")):
            return token
    for token in parts:
        if token.endswith("동"):
            return token
    return parts[-1]


def compose_sigungu_dong_label(sigungu: Any, dong: Any) -> str:
    sg = normalize_sigungu_name_value(sigungu)
    dg = clean_region_part(dong)
    if sg and dg and not dg.startswith(sg):
        return f"{sg} {dg}"
    return dg or sg or "-"


def representative_sigungu(df: pd.DataFrame, group_col: str, group_value: Any = None) -> str:
    if df is None or df.empty or "_sigungu" not in df.columns:
        return ""
    sub = df
    if group_col in df.columns and group_value is not None:
        sub = df[df[group_col] == group_value]
    vals = sub["_sigungu"].dropna().astype(str).map(normalize_sigungu_name_value)
    vals = vals[~vals.isin(["", "nan", "None"])]
    if vals.empty:
        return ""
    return str(vals.mode().iloc[0]) if not vals.mode().empty else str(vals.iloc[0])


def calc_supply(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
    """주택공급 파이프라인 계산.
    v5.8.38 기준 보정:
    - 업로드된 기준 이미지와 동일하게 KOSIS 공급 원자료의 '값' 컬럼을 사용
    - 과거연도는 12월 누계, 당해연도는 업로드된 최신 누계월만 사용
    - 인허가·착공·준공: 합계(가구수기준)/계(다가구가구수기준) 행을 총량으로 사용
    - 물건유형별: 아파트·연립·다세대·단독·다가구 행만 사용하며, 다가구는 '가구수' 기준 사용
    - 공동주택 분양: 항목명=합계, 분류3명=누계 행만 총량표·총량차트에 반영
    """
    result = {"available": False, "table_rows": [], "by_type_rows": [], "permit_by_type_rows": [], "start_by_type_rows": [], "completion_by_type_rows": [], "presale_by_type_rows": [], "latest": {}, "signal": "공급 원자료 미확인"}
    if not files:
        return result

    metric_map = {"인허가": "permit", "착공": "start", "준공": "completion"}
    monthly: Dict[str, Dict[str, float]] = {}
    period_by_year: Dict[str, str] = {}
    by_type_records: List[Dict[str, Any]] = []

    def _supply_region_filter(df: pd.DataFrame, metric: str) -> pd.DataFrame:
        # 주택유형별 인허가·착공·준공은 시도 단위 공급통계가 기준임.
        # 예: 전라남도 목포시 보고서라도 공급 파이프라인은 전라남도/전남 행을 사용.
        if df.empty:
            return df
        keywords = get_supply_sido_keywords(region)
        def _match_col(col: str) -> pd.DataFrame:
            if col not in df.columns:
                return df.iloc[0:0].copy()
            s = df[col].astype(str)
            mask = False
            for kw in keywords:
                mask = mask | s.str.contains(str(kw), na=False)
            return df[mask].copy() if hasattr(mask, "any") else df.iloc[0:0].copy()

        # 공급 3종은 보통 분류1명에 시도명이 들어감. 분양은 분류2명인 경우가 많아 예비 처리.
        for col in ["분류1명", "시도명", "지역명", "분류2명"]:
            sub = _match_col(col)
            if not sub.empty:
                return sub

        # 마지막 fallback은 기존 지역명 필터. 단, 시군구명으로 잘못 좁혀져 비는 경우가 있을 수 있음.
        return filter_df_by_region(df, region)

    def _month_col(df: pd.DataFrame) -> Optional[str]:
        for c in ["기간", "조회기준월", "기준월", "PRD_DE", "계약년월"]:
            if c in df.columns:
                return c
        return find_best_month_col(df)

    def _value_col(df: pd.DataFrame) -> Optional[str]:
        # KOSIS 공급 파일은 '값' 컬럼이 실적값. 면적·코드 오인 방지.
        for c in ["값", "데이터", "OBS_VALUE", "VALUE", "DT"]:
            if c in df.columns:
                return c
        return find_supply_value_col(df) if "find_supply_value_col" in globals() else find_value_col(df)

    def _year_key(month: Any) -> str:
        m = normalize_month_value(month)
        nums = re.sub(r"[^0-9]", "", str(m))
        if len(nums) >= 4:
            return nums[:4]
        return str(m)[:4]

    def _is_cumulative_row(row: pd.Series) -> bool:
        # 연말및당해월누계 파일은 누계 행이 기준. 순계 행은 제외.
        for c in ["분류3명", "명칭", "조회구분"]:
            if c in row.index and "순계" in str(row.get(c)):
                return False
        # 분양실적은 분류3명=누계 행만 사용
        if "분류3명" in row.index and "누계" in str(row.get("분류3명")):
            return True
        return True

    def _normalize_supply_type(row: pd.Series) -> Optional[str]:
        # 다가구는 동수와 가구수 중 '가구수'가 주택공급 호수 기준이므로 가구수 행만 사용
        vals = {c: str(row.get(c, "")) for c in ["분류2명", "분류3명", "분류4명", "C4_NM_ENG", "C3_NM_ENG"]}
        joined = " ".join(vals.values())
        if "합계" in joined or "계(" in joined:
            return None
        if "아파트" in joined:
            return "아파트"
        if "연립" in joined:
            return "연립"
        if "다세대" in joined:
            return "다세대"
        if "다가구" in joined:
            # 분류4명 동수는 제외하고 가구수 기준만 사용
            if "동수" in vals.get("분류4명", "") or "Number of building" in joined:
                return None
            return "다가구"
        # 다가구가 아닌 일반 단독
        if "단독" in joined:
            return "단독"
        return None

    def _is_total_row(row: pd.Series, metric: str) -> bool:
        if metric == "presale_supply":
            return str(row.get("항목명", "")) == "합계" and _is_cumulative_row(row)
        joined = " ".join(str(row.get(c, "")) for c in ["분류2명", "분류3명", "분류4명"])
        return ("합계(가구수기준)" in joined or "계(다가구가구수기준)" in joined) and _is_cumulative_row(row)

    try:
        for p in files:
            progress(f"공급 파일 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            metric = None
            for key, val in metric_map.items():
                if key in p.name:
                    metric = val
                    break
            if not metric:
                continue

            df = _supply_region_filter(df, metric)
            if df.empty:
                continue
            mcol, vcol = _month_col(df), _value_col(df)
            if not vcol:
                continue
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            df["_year"] = df["_month"].map(_year_key)
            df["_value"] = to_numeric_safe(df[vcol])

            # 총량: 정답 이미지 기준 행만 사용
            total_rows = []
            for _, r in df.iterrows():
                if pd.notna(r.get("_value")) and _is_total_row(r, metric):
                    total_rows.append(r)
            if total_rows:
                tdf = pd.DataFrame(total_rows)
                # 동일 연도 중 최신 누계월 1개만 사용
                for y, sub in tdf.groupby("_year"):
                    sub = sub.sort_values("_month")
                    val = float(sub.iloc[-1]["_value"])
                    monthly.setdefault(str(y), {})[metric] = int(round(val))
                    actual_month = str(sub.iloc[-1]["_month"])
                    if actual_month > period_by_year.get(str(y), ""):
                        period_by_year[str(y)] = actual_month

            # 물건유형별: 공동주택 분양은 주택유형별 제공이 아니므로 제외
            if metric != "presale_supply":
                recs = []
                for _, r in df.iterrows():
                    if pd.isna(r.get("_value")) or not _is_cumulative_row(r):
                        continue
                    typ = _normalize_supply_type(r)
                    if not typ:
                        continue
                    recs.append({"month": r["_month"], "year": str(r["_year"]), "stage": metric, "type": typ, "value": float(r["_value"])})
                if recs:
                    rdf = pd.DataFrame(recs)
                    # 같은 연도·단계·유형 중 최신 누계월 사용
                    rdf = rdf.sort_values("month").groupby(["year", "stage", "type"], as_index=False).tail(1)
                    for _, r in rdf.iterrows():
                        by_type_records.append({"month": str(r["month"]), "stage": metric, "type": str(r["type"]), "value": int(round(float(r["value"])))})

        if not monthly:
            return result

        rows = []
        for y in sorted(monthly):
            actual_month = period_by_year.get(str(y), f"{y}-12")
            month_num = re.sub(r"[^0-9]", "", actual_month)[4:6] or "12"
            period_label = str(y) if month_num == "12" else f"{y}년 {int(month_num)}월 누계"
            row = {"month": actual_month, "period_label": period_label}
            row.update(monthly[y])
            rows.append(row)
        dfm = pd.DataFrame(rows).sort_values("month")
        latest = dfm.iloc[-1].to_dict()

        by_type_rows = []
        permit_by_type_rows = []
        start_by_type_rows = []
        completion_by_type_rows = []
        presale_by_type_rows = []
        if by_type_records:
            bdf = pd.DataFrame(by_type_records)
            bdf = bdf.groupby(["month", "stage", "type"], as_index=False)["value"].sum()
            by_type_rows = bdf.to_dict("records")
            permit_by_type_rows = bdf[bdf["stage"] == "permit"].to_dict("records")
            start_by_type_rows = bdf[bdf["stage"] == "start"].to_dict("records")
            completion_by_type_rows = bdf[bdf["stage"] == "completion"].to_dict("records")
            presale_by_type_rows = []

        result.update({
            "available": True,
            "table_rows": dfm.tail(10).to_dict("records"),
            "by_type_rows": by_type_rows,
            "permit_by_type_rows": permit_by_type_rows,
            "start_by_type_rows": start_by_type_rows,
            "completion_by_type_rows": completion_by_type_rows,
            "presale_by_type_rows": presale_by_type_rows,
            "latest": latest,
            "signal": "주택공급 파이프라인 점검"
        })
        return result
    except Exception as e:
        result["signal"] = f"공급 계산 오류: {e}"
        return result


# =========================================================
# 5. 실거래 계산
# =========================================================

def prepare_trade_df(files: List[Path], region: str, base_month: str) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    frames = []
    for i, p in enumerate(files, start=1):
        try:
            progress(f"실거래 파일 처리 중 ({i}/{len(files)}): {p.name}", i, len(files))
            _df_src = normalize_columns(read_excel_or_csv(p))
            _name = p.name
            if "연립" in _name and "다세대" not in _name:
                _src_type = "연립"
            elif "다세대" in _name and "연립" not in _name:
                _src_type = "다세대"
            elif "연립다세대" in _name or "연립_다세대" in _name or "연립·다세대" in _name:
                _src_type = "연립·다세대"
            elif "단독" in _name and "다가구" not in _name:
                _src_type = "단독"
            elif "다가구" in _name and "단독" not in _name:
                _src_type = "다가구"
            elif "단독다가구" in _name or "단독_다가구" in _name or "단독·다가구" in _name:
                _src_type = "단독·다가구"
            elif "아파트" in _name:
                _src_type = "아파트"
            elif "분양권" in _name:
                _src_type = "분양권전매"
            else:
                _src_type = "기타"
            _df_src["_source_type"] = _src_type
            _df_src["_source_file"] = p.name
            frames.append(_df_src)
        except Exception as e:
            progress(f"실거래 파일 처리 실패: {p.name} / {e}")
    if not frames:
        return pd.DataFrame(), {}
    df = pd.concat(frames, ignore_index=True)
    df = filter_df_by_region(df, region)
    cols = {
        "month": find_best_month_col(df),
        "dong": find_col(df, COLUMN_CANDIDATES["dong"]),
        "sigungu": find_col(df, ["시군구", "시군구명", "법정동시군구명", "sggNm"]),
        "sigungu_code": find_col(df, ["LAWD_CD", "sggCd", "지역코드", "법정동시군구"]),
        "apt_name": find_col(df, COLUMN_CANDIDATES["apt_name"]),
        "jibun": find_col(df, COLUMN_CANDIDATES["jibun"]),
        "approval_year": find_col(df, COLUMN_CANDIDATES["approval_year"]),
        "households": find_col(df, COLUMN_CANDIDATES["households"]),
        "area": find_col(df, COLUMN_CANDIDATES["exclusive_area"]),
        "price": find_col(df, COLUMN_CANDIDATES["price"]),
        "deposit": find_col(df, COLUMN_CANDIDATES["deposit"]),
        "monthly_rent": find_col(df, COLUMN_CANDIDATES["monthly_rent"]),
        "floor": find_col(df, COLUMN_CANDIDATES["floor"]),
        "contract_date": find_col(df, COLUMN_CANDIDATES["contract_date"]),
        "contract_type": find_col(df, COLUMN_CANDIDATES["contract_type"]),
        "rent_trade_type": find_col(df, ["거래구분", "임대차구분", "전월세구분"]),
        "deal_type": find_col(df, ["거래유형", "dealingGbn"]),
        "item": find_col(df, ["주택유형", "물건구분", "유형", "건물용도"]),
        # 지정 업로드 매매 원장은 건축물주용도를 주택유형의 기준 열로 사용한다.
        # 건축물용도(일반단독·다가구단독)는 보조 열로만 사용한다.
        "item_detail": find_col(df, ["건축물주용도", "건축물용도", "주용도", "건물유형"]),
        "item_detail_secondary": find_col(df, ["건물유형"]),
    }
    df["_month"] = df[cols["month"]].map(normalize_month_value) if cols["month"] else base_month
    # _dong_full: 원자료 표기(리 포함) — 동일 물건 비교·매칭 키 전용
    # _dong: 보고서 표시·집계 단위(읍·면·동까지)
    df["_dong_full"] = df[cols["dong"]].astype(str) if cols["dong"] else ""
    df["_dong"] = df["_dong_full"].map(normalize_dong_to_eupmyeondong)
    if cols.get("sigungu"):
        df["_sigungu"] = df[cols["sigungu"]].map(normalize_sigungu_name_value)
    elif cols.get("sigungu_code"):
        df["_sigungu"] = df[cols["sigungu_code"]].map(sigungu_name_from_code)
    else:
        df["_sigungu"] = ""
    if not is_sido_level_region(region):
        default_sigungu = get_region_short_name(region)
        if default_sigungu and any(default_sigungu.endswith(suffix) for suffix in ["시", "군", "구"]):
            empty_sigungu = df["_sigungu"].astype(str).str.strip().isin(["", "nan", "None"])
            df.loc[empty_sigungu, "_sigungu"] = default_sigungu
    if is_sido_level_region(region):
        df["_dong"] = [compose_sido_dong_label(sg, dg) for sg, dg in zip(df["_sigungu"], df["_dong"])]
    df["_apt_name"] = df[cols["apt_name"]].astype(str) if cols["apt_name"] else ""
    df["_jibun"] = df[cols["jibun"]].astype(str) if cols.get("jibun") else ""
    if cols["approval_year"]:
        df["_approval_year"] = pd.array(df[cols["approval_year"]].map(normalize_approval_year_value), dtype="Int64")
    else:
        df["_approval_year"] = pd.array([pd.NA] * len(df), dtype="Int64")
    df["_households"] = df[cols["households"]] if cols["households"] else ""
    df["_area"] = to_numeric_safe(df[cols["area"]]) if cols["area"] else np.nan
    df["_price"] = normalize_money_to_manwon(df[cols["price"]], "price") if cols["price"] else np.nan
    df["_deposit"] = normalize_money_to_manwon(df[cols["deposit"]], "deposit") if cols["deposit"] else np.nan
    df["_monthly_rent"] = normalize_money_to_manwon(df[cols["monthly_rent"]], "rent") if cols["monthly_rent"] else np.nan
    df["_floor"] = to_numeric_safe(df[cols["floor"]]) if cols["floor"] else np.nan
    df["_contract_date"] = df[cols["contract_date"]].astype(str) if cols.get("contract_date") else ""
    df["_rent_trade_type"] = df[cols["rent_trade_type"]].astype(str).str.strip() if cols.get("rent_trade_type") else ""
    # 매매 신고자료의 거래유형(중개거래·직거래). 전월세 신고자료에는 이 항목이 없어 빈 값으로 둔다.
    _deal_raw = df[cols["deal_type"]].astype(str).str.replace(" ", "", regex=False) if cols.get("deal_type") else pd.Series("", index=df.index)
    df["_deal_type"] = np.where(_deal_raw.str.contains("직거래", na=False), "직거래",
                                np.where(_deal_raw.str.contains("중개", na=False), "중개거래", ""))
    df["_item"] = df[cols["item"]].astype(str) if cols["item"] else ""
    if "_source_type" not in df.columns:
        df["_source_type"] = ""
    # 물건구분/주택유형 컬럼이 있으면 파일명 추정값보다 원자료 유형을 우선하되, 빈 값은 파일명 추정값 유지
    if cols.get("item"):
        _raw_item = df[cols["item"]].astype(str)
        df["_source_type"] = np.where(_raw_item.str.contains("연립", na=False) & ~_raw_item.str.contains("다세대", na=False), "연립", df["_source_type"])
        df["_source_type"] = np.where(_raw_item.str.contains("다세대", na=False), "다세대", df["_source_type"])
        df["_source_type"] = np.where(_raw_item.str.contains("단독", na=False) & ~_raw_item.str.contains("다가구", na=False), "단독", df["_source_type"])
        df["_source_type"] = np.where(_raw_item.str.contains("다가구", na=False), "다가구", df["_source_type"])
        df["_source_type"] = np.where(_raw_item.str.contains("아파트", na=False), "아파트", df["_source_type"])
    if cols.get("item_detail"):
        _raw_detail = df[cols["item_detail"]].astype(str).str.replace(" ", "", regex=False).str.strip()
        secondary_col = cols.get("item_detail_secondary")
        if secondary_col and secondary_col != cols.get("item_detail"):
            _raw_secondary = df[secondary_col].astype(str).str.replace(" ", "", regex=False).str.strip()
            _primary_empty = _raw_detail.str.lower().isin(["", "nan", "none"])
            _raw_detail = _raw_detail.where(~_primary_empty, _raw_secondary)
        # 원자료 표기: 아파트·연립주택·다세대주택·단독주택·다가구주택
        # 보고서 내부 표기: 아파트·연립·다세대·단독·다가구
        df["_source_type"] = np.where(_raw_detail.eq("아파트"), "아파트", df["_source_type"])
        df["_source_type"] = np.where(_raw_detail.str.contains("연립주택|연립", na=False, regex=True), "연립", df["_source_type"])
        df["_source_type"] = np.where(_raw_detail.str.contains("다세대주택|다세대", na=False, regex=True), "다세대", df["_source_type"])
        df["_source_type"] = np.where(_raw_detail.str.contains("다가구주택|다가구", na=False, regex=True), "다가구", df["_source_type"])
        df["_source_type"] = np.where(_raw_detail.str.contains("단독주택|일반단독", na=False, regex=True) & ~_raw_detail.str.contains("다가구", na=False), "단독", df["_source_type"])
    df["_source_type"] = pd.Series(df["_source_type"]).replace("", np.nan).fillna("기타")
    return df, cols


def add_previous_compare(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    out["_area_key"] = out["_area"].round(2)
    def _make_sort_key(m, d):
        s = str(d).strip()
        if not s or s.lower() in ['nan', 'none']: return str(m)
        if s.isdigit(): return f"{m}-{int(s):02d}"
        if '~' in s:
            parts = s.split('~')
            if parts[0].strip().isdigit():
                return f"{m}-{int(parts[0].strip()):02d}~"
        return f"{m}-{s}"
        
    if "_contract_date" in out.columns:
        out["_sort_key"] = out.apply(lambda r: _make_sort_key(r["_month"], r["_contract_date"]), axis=1)
    else:
        out["_sort_key"] = out["_month"].astype(str)
    group_cols = ["_apt_name", "_area_key"]
    # 단지명이 없는 비아파트는 동+면적으로 비교.
    # 비교 정확도를 위해 표시용 읍면동(_dong)이 아니라 원자료 표기(_dong_full, 리 포함)를 사용한다.
    if "_dong_full" not in out.columns:
        out["_dong_full"] = out["_dong"]
    out["_group_name"] = out["_apt_name"].mask(out["_apt_name"].eq(""), np.nan).fillna(out["_dong_full"])
    group_cols = ["_group_name", "_area_key"]
    out = out.sort_values(group_cols + ["_sort_key"])
    out["_prev_value"] = out.groupby(group_cols)[value_col].shift(1)
    # v18: 2차 매칭 — 1차에서 prev_value 가 비어 있는 행만 동+지번+면적정수 키로 재매칭.
    # 오피스텔 월세처럼 단지명 표기가 흔들리거나 결측인 경우 비교가능건이 0건이 되지 않도록 보강.
    try:
        if "_jibun" in out.columns:
            out["_area_int"] = out["_area"].round(0)
            secondary_cols = ["_dong_full", "_jibun", "_area_int"]
            out_sorted2 = out.sort_values(secondary_cols + ["_sort_key"])
            shifted2 = out_sorted2.groupby(secondary_cols)[value_col].shift(1)
            mask_missing = out["_prev_value"].isna()
            fill_vals = shifted2.reindex(out.index)
            out.loc[mask_missing, "_prev_value"] = fill_vals[mask_missing]
    except Exception:
        pass
    out["_prev_floor"] = out.groupby(group_cols)["_floor"].shift(1) if "_floor" in out.columns else np.nan
    out["_prev_contract_date"] = out.groupby(group_cols)["_contract_date"].shift(1) if "_contract_date" in out.columns else ""
    out["_prev_month"] = out.groupby(group_cols)["_month"].shift(1)
    out["_prev_deposit"] = out.groupby(group_cols)["_deposit"].shift(1) if "_deposit" in out.columns else np.nan
    out["_prev_monthly_rent"] = out.groupby(group_cols)["_monthly_rent"].shift(1) if "_monthly_rent" in out.columns else np.nan
    out["_prev_deal_type"] = out.groupby(group_cols)["_deal_type"].shift(1) if "_deal_type" in out.columns else ""
    out["_change_rate"] = (out[value_col] - out["_prev_value"]) / out["_prev_value"] * 100

    def direction(x):
        if pd.isna(x):
            return "비교불가"
        if x > 0.5:
            return "상승"
        if x < -0.5:
            return "하락"
        return "보합"

    out["_direction"] = out["_change_rate"].map(direction)
    return out


def classify_area_band(area: Any) -> str:
    try:
        a = float(area)
        if a <= 40: return "40㎡이하"
        if a <= 60: return "60㎡이하"
        if a <= 85: return "60㎡초과 ~ 85㎡이하"
        if a <= 102: return "85㎡초과 ~ 102㎡이하"
        if a <= 135: return "102㎡초과 ~ 135㎡이하"
        return "135㎡초과"
    except Exception:
        return "면적미상"


def classify_price_band_manwon(v: Any) -> str:
    try:
        eok = float(v) / 10000
        if eok < 1: return "1억 미만"
        if eok < 2: return "1~2억"
        if eok < 3: return "2~3억"
        if eok < 5: return "3~5억"
        if eok < 7: return "5~7억"
        return "7억 이상"
    except Exception:
        return "가격미상"


def classify_deposit_band_manwon(v: Any) -> str:
    try:
        eok = float(v) / 10000
        if eok < 0.5: return "5천만원 미만"
        if eok < 1: return "5천만~1억"
        if eok < 2: return "1~2억"
        if eok < 3: return "2~3억"
        if eok < 4: return "3~4억"
        return "4억 이상"
    except Exception:
        return "보증금미상"


def make_dynamic_deposit_band_func(values, max_bins: int = 5):
    """지역·자료분포 기반 보증금 구간 함수 생성.
    고가 지역은 4억 이상으로 대부분 몰리는 문제를 줄이고,
    실제 기준월 보증금 분포의 분위수를 이용해 구간을 자동 산정한다.
    단위는 원자료와 동일한 만원 기준이다.
    """
    try:
        s = pd.to_numeric(pd.Series(values), errors="coerce").dropna()
        s = s[s > 0]
        if len(s) < 10 or s.nunique() < 3:
            return classify_deposit_band_manwon

        qs = [0, 0.2, 0.4, 0.6, 0.8, 1.0]
        cuts = sorted(set(float(s.quantile(q)) for q in qs))
        # 동일값이 많아 분위수 구간이 무너지는 경우 기존 고정 구간 사용
        if len(cuts) < 4:
            return classify_deposit_band_manwon

        # 보기 좋은 단위로 반올림: 1억 미만 1천만원, 1억 이상 5천만원, 5억 이상 1억원 단위
        def nice(v):
            if v >= 50000:
                unit = 10000
            elif v >= 10000:
                unit = 5000
            else:
                unit = 1000
            return round(v / unit) * unit

        nice_cuts = [cuts[0]]
        for v in cuts[1:-1]:
            nv = nice(v)
            if nv > nice_cuts[-1]:
                nice_cuts.append(nv)
        nice_cuts.append(cuts[-1])

        # 반올림 후 중복/역전 보정
        clean = []
        for v in nice_cuts:
            if not clean or v > clean[-1]:
                clean.append(v)
        if len(clean) < 4:
            return classify_deposit_band_manwon

        def fmt(v):
            v = float(v)
            if v >= 10000:
                eok = v / 10000
                if abs(eok - round(eok)) < 0.05:
                    return f"{int(round(eok))}억"
                return f"{eok:.1f}억"
            return f"{int(round(v / 1000))}천만원"

        def band(x):
            try:
                val = float(x)
            except Exception:
                return "미상"
            if pd.isna(val) or val <= 0:
                return "미상"
            for a, b in zip(clean[:-1], clean[1:]):
                if val <= b:
                    return f"{fmt(a)}~{fmt(b)}"
            return f"{fmt(clean[-2])} 초과"

        return band
    except Exception:
        return classify_deposit_band_manwon


def make_dynamic_numeric_band_func(values, unit: str, fallback_func, max_bins: int = 5):
    """자료분포 기반 숫자 구간 함수 생성.
    - unit='price_manwon' : 매매가격/보증금처럼 만원 단위 금액
    - unit='monthly' : 월세/환산월세처럼 만원 단위 월액
    """
    try:
        s = pd.to_numeric(pd.Series(values), errors="coerce").dropna()
        s = s[s > 0]
        if len(s) < 10 or s.nunique() < 3:
            return fallback_func

        qs = [0, 0.2, 0.4, 0.6, 0.8, 1.0]
        cuts = sorted(set(float(s.quantile(q)) for q in qs))
        if len(cuts) < 4:
            return fallback_func

        def nice_price(v):
            # 만원 단위: 5억 이상 1억원, 1억 이상 5천만원, 1억 미만 1천만원 단위
            if v >= 50000:
                unit_val = 10000
            elif v >= 10000:
                unit_val = 5000
            else:
                unit_val = 1000
            return round(v / unit_val) * unit_val

        def nice_monthly(v):
            # 월세/환산월세: 300만원 이상 50만원, 100만원 이상 20만원, 그 외 10만원 단위
            if v >= 300:
                unit_val = 50
            elif v >= 100:
                unit_val = 20
            else:
                unit_val = 10
            return round(v / unit_val) * unit_val

        def fmt_price(v):
            v = float(v)
            if v >= 10000:
                eok = v / 10000
                if abs(eok - round(eok)) < 0.05:
                    return f"{int(round(eok))}억"
                return f"{eok:.1f}억"
            return f"{int(round(v / 1000))}천만원"

        def fmt_monthly(v):
            v = float(v)
            if abs(v - round(v)) < 0.05:
                return f"{int(round(v))}만원"
            return f"{v:.1f}만원"

        nice_func = nice_price if unit == "price_manwon" else nice_monthly
        fmt_func = fmt_price if unit == "price_manwon" else fmt_monthly

        clean = [cuts[0]]
        for v in cuts[1:-1]:
            nv = nice_func(v)
            if nv > clean[-1]:
                clean.append(nv)
        clean.append(cuts[-1])

        clean2 = []
        for v in clean:
            if not clean2 or v > clean2[-1]:
                clean2.append(v)
        if len(clean2) < 4:
            return fallback_func

        def band(x):
            try:
                val = float(x)
            except Exception:
                return "미상"
            if pd.isna(val) or val <= 0:
                return "미상"
            for a, b in zip(clean2[:-1], clean2[1:]):
                if val <= b:
                    return f"{fmt_func(a)}~{fmt_func(b)}"
            return f"{fmt_func(clean2[-2])} 초과"

        return band
    except Exception:
        return fallback_func


def make_dynamic_price_band_func(values, max_bins: int = 5):
    return make_dynamic_numeric_band_func(values, "price_manwon", classify_price_band_manwon, max_bins=max_bins)


def make_dynamic_monthly_band_func(values, max_bins: int = 5):
    return make_dynamic_numeric_band_func(values, "monthly", classify_monthly_band, max_bins=max_bins)


def make_dynamic_converted_rent_band_func(values, max_bins: int = 5):
    return make_dynamic_numeric_band_func(values, "monthly", classify_converted_rent_band, max_bins=max_bins)

def classify_monthly_band(v: Any) -> str:
    try:
        x = float(v)
        if x < 30: return "30만원 미만"
        if x < 50: return "30~50만원"
        if x < 70: return "50~70만원"
        if x < 100: return "70~100만원"
        return "100만원 이상"
    except Exception:
        return "월세미상"

def classify_converted_rent_band(v: Any) -> str:
    try:
        x = float(v)
        if x < 40: return "40만원 미만"
        if x < 60: return "40~60만원"
        if x < 80: return "60~80만원"
        if x < 100: return "80~100만원"
        if x < 150: return "100~150만원"
        return "150만원 이상"
    except Exception:
        return "환산월세미상"

def classify_building_age(approval_year: Any, base_month: str) -> str:
    try:
        y = normalize_approval_year_value(approval_year)
        if y is None:
            return "연식미상"
        base_y = int(str(base_month)[:4])
        age = base_y - y
        if age <= 5: return "5년 이하"
        if age <= 10: return "5년 초과 ~ 10년 이하"
        if age <= 15: return "10년 초과 ~ 15년 이하"
        if age <= 20: return "15년 초과 ~ 20년 이하"
        return "20년 초과"
    except Exception:
        return "연식미상"

def age_to_group(age_text: Any) -> str:
    s = str(age_text)
    nums = re.findall(r"\d+", s)
    if not nums:
        return "연령미상"
    a = int(nums[0])
    if a < 10: return "0~9세"
    if a < 20: return "10대"
    if a < 30: return "20대"
    if a < 40: return "30대"
    if a < 50: return "40대"
    if a < 60: return "50대"
    if a < 70: return "60대"
    return "70대 이상"


def recent_month_window(end_ym: str, n: int = 13) -> List[str]:
    """보고서 기준월(end_ym, 'YYYY-MM')을 끝으로 하는 연속 n개월 'YYYY-MM' 목록을 오름차순 반환."""
    try:
        s = str(end_ym)
        y, m = int(s[:4]), int(s[5:7])
    except Exception:
        return []
    out = []
    for k in range(n - 1, -1, -1):
        total = y * 12 + (m - 1) - k
        yy, mm = divmod(total, 12)
        out.append(f"{yy:04d}-{mm + 1:02d}")
    return out


def calc_external_buyer_sale(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
    """업로드한 주거용 통합 매매원장에서 외지인 매수 흐름을 계산한다.

    외지인 여부는 원장의 별도 플래그를 사용하지 않고, 사용자가 요청한 대로
    ``매수자 시군구``와 선택 보고지역의 시군구를 직접 비교한다. 수치 계산은
    Python에서 완료하고 LLM에는 집계 결과만 전달한다.
    """
    result = {
        "available": False,
        "source_file": None,
        "basis": "계약일 기준·매수자 시군구 직접 비교",
        "latest_month": base_month,
        "target_sigungu": get_region_short_name(region),
        "monthly_rows": [],
        "current_region_rows": [],
        "six_month_region_rows": [],
        "current_total_count": 0,
        "current_external_count": 0,
        "current_external_share": 0.0,
        "six_month_total_count": 0,
        "six_month_external_count": 0,
        "six_month_external_share": 0.0,
        "unknown_buyer_region_count": 0,
        "signal": "매수자 시군구가 포함된 주거용 매매 원장 미확인",
    }
    if not files:
        return result

    frames = []
    used_files = []
    residential_pattern = r"공동주택|아파트|연립|다세대|단독주택|다가구|주거|건물만거래"
    empty_tokens = {"", "-", "--", "nan", "none", "nat", "0", "0.0"}
    for p in files:
        try:
            progress(f"외지인 매수 원장 확인 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            buyer_col = find_col(df, ["매수자시군구"])
            date_col = find_col(df, ["계약일자", "계약일", "거래일"])
            region_col = find_col(df, ["시군구", "시군구명", "법정동시군구명"])
            trade_col = find_col(df, ["거래구분"])
            item_col = find_col(df, ["건축물주용도", "물건구분", "주택유형", "건물유형", "건축물용도"])
            key_col = find_col(df, ["RTMS_KEY", "거래키", "계약번호"])
            if not buyer_col or not date_col:
                continue
            if trade_col:
                df = df[df[trade_col].astype(str).str.contains("매매", na=False)].copy()
            if item_col and not df.empty:
                raw_item = df[item_col].astype(str).str.replace(" ", "", regex=False)
                residential_mask = raw_item.str.contains(residential_pattern, na=False, regex=True)
                # 토지시장 통합원장 등이 잘못 선택된 경우를 차단한다.
                if float(residential_mask.mean()) < 0.5:
                    continue
                df = df[residential_mask].copy()
            if region_col and not df.empty:
                target = get_region_short_name(region)
                if target:
                    region_mask = df[region_col].astype(str).str.contains(target, na=False, regex=False)
                    if region_mask.any():
                        df = df[region_mask].copy()
            for cancel_name in ["해제사유발생일", "계약해제접수일", "해제사실확인사유"]:
                cancel_col = find_col(df, [cancel_name])
                if cancel_col and not df.empty:
                    cancel_text = df[cancel_col].astype(str).str.strip().str.lower()
                    df = df[cancel_text.isin(empty_tokens)].copy()
            if df.empty:
                continue
            df["_contract_dt"] = pd.to_datetime(df[date_col], errors="coerce")
            df = df[df["_contract_dt"].notna()].copy()
            df["_month"] = df["_contract_dt"].dt.strftime("%Y-%m")
            df = df[df["_month"] <= base_month].copy()
            if key_col:
                df = df.drop_duplicates(subset=[key_col], keep="last")
            df["_buyer_sigungu"] = (
                df[buyer_col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
            )
            df["_source_file"] = p.name
            frames.append(df[["_month", "_buyer_sigungu", "_source_file"]])
            used_files.append(p.name)
        except Exception as e:
            progress(f"외지인 매수 원장 처리 실패: {p.name} / {e}")

    if not frames:
        return result
    work = pd.concat(frames, ignore_index=True)
    # 같은 통합원장을 중복 선택했을 때는 가장 행이 많은 주거용 원장을 우선한다.
    if len(used_files) > 1:
        source_counts = work["_source_file"].value_counts()
        chosen = str(source_counts.index[0])
        work = work[work["_source_file"] == chosen].copy()
        used_files = [chosen]

    buyer_lower = work["_buyer_sigungu"].str.lower()
    known = ~buyer_lower.isin(empty_tokens)
    target = get_region_short_name(region)
    local = known & work["_buyer_sigungu"].str.contains(target, na=False, regex=False)
    work["_is_external"] = known & ~local
    work["_buyer_region_display"] = work["_buyer_sigungu"].where(known, "미상")

    months13 = recent_month_window(base_month, 13)
    monthly_rows = []
    for month in months13:
        part = work[work["_month"] == month]
        total = int(len(part))
        ext = int(part["_is_external"].sum())
        monthly_rows.append({
            "month": month,
            "전체 매매건수": total,
            "외지인 거래건수": ext,
            "외지인 비율(%)": safe_ratio(ext, total),
        })

    current = work[work["_month"] == base_month]
    current_external = current[current["_is_external"]].copy()
    current_total = int(len(current))
    current_ext_count = int(len(current_external))
    current_counts = current_external["_buyer_region_display"].value_counts()
    current_rows = []
    for buyer_region, count in current_counts.head(15).items():
        current_rows.append({
            "매수자 시군구": str(buyer_region),
            "당월 외지인 건수": int(count),
            "외지인 내 비율(%)": safe_ratio(count, current_ext_count),
            "전체 매매 대비(%)": safe_ratio(count, current_total),
        })

    months6 = recent_month_window(base_month, 6)
    six = work[work["_month"].isin(months6)]
    six_external = six[six["_is_external"]].copy()
    six_total = int(len(six))
    six_ext_count = int(len(six_external))
    six_pivot = six_external.groupby(["_buyer_region_display", "_month"]).size().unstack(fill_value=0)
    six_rows = []
    if not six_pivot.empty:
        six_pivot = six_pivot.reindex(columns=months6, fill_value=0)
        six_pivot["6개월 누적"] = six_pivot.sum(axis=1)
        six_pivot = six_pivot.sort_values("6개월 누적", ascending=False).head(15)
        for buyer_region, row in six_pivot.iterrows():
            item = {"매수자 시군구": str(buyer_region)}
            item.update({m: int(row.get(m, 0)) for m in months6})
            item["6개월 누적"] = int(row.get("6개월 누적", 0))
            item["6개월 외지인 비율(%)"] = safe_ratio(item["6개월 누적"], six_ext_count)
            six_rows.append(item)

    direction = "확대" if monthly_rows and monthly_rows[-1]["외지인 비율(%)"] > monthly_rows[-2]["외지인 비율(%)"] else "축소"
    result.update({
        "available": True,
        "source_file": used_files[0],
        "monthly_rows": monthly_rows,
        "current_region_rows": current_rows,
        "six_month_region_rows": six_rows,
        "current_total_count": current_total,
        "current_external_count": current_ext_count,
        "current_external_share": safe_ratio(current_ext_count, current_total),
        "six_month_total_count": six_total,
        "six_month_external_count": six_ext_count,
        "six_month_external_share": safe_ratio(six_ext_count, six_total),
        "unknown_buyer_region_count": int((~known).sum()),
        "signal": f"당월 외지인 매수 비율 {safe_ratio(current_ext_count, current_total):.1f}%로 전월 대비 {direction}",
    })
    return result


def dong_monthly_volume_rows(df: pd.DataFrame, value_col: str, end_ym: str, n: int = 13) -> List[Dict[str, Any]]:
    """읍면동별 최근 n개월 거래량을 월 컬럼 교차표로 반환한다."""
    months = recent_month_window(end_ym, n)
    if df is None or df.empty or not months or "_dong" not in df.columns or "_month" not in df.columns:
        return []
    extended_months = recent_month_window(end_ym, n + 1)
    previous_month = extended_months[0] if len(extended_months) == n + 1 else ""
    work = df[df["_month"].isin(([previous_month] if previous_month else []) + months)].copy()
    if value_col in work.columns:
        work = work[work[value_col].notna()].copy()
    if work.empty:
        return []
    work["_dong"] = work["_dong"].map(normalize_dong_to_eupmyeondong).replace("", np.nan).fillna("미상").astype(str)
    counts_all = work.groupby(["_dong", "_month"]).size().unstack(fill_value=0)
    display_counts = counts_all.reindex(columns=months, fill_value=0)
    # 표시 13개월 중 거래가 있는 읍면동만 표에 유지한다.
    display_counts = display_counts[display_counts.sum(axis=1) > 0].copy()
    display_counts["합계"] = display_counts.sum(axis=1)
    display_counts = display_counts.sort_values("합계", ascending=False)
    previous_available = bool(previous_month and (work["_month"] == previous_month).any())
    rows = []
    for dong, row in display_counts.iterrows():
        previous_count = int(counts_all.at[dong, previous_month]) if previous_available and previous_month in counts_all.columns and dong in counts_all.index else (0 if previous_available else None)
        item = {"읍면동": str(dong), "_previous_month": previous_month, "_previous_month_count": previous_count}
        item.update({month: int(row.get(month, 0)) for month in months})
        item["합계"] = int(row.get("합계", 0))
        rows.append(item)
    return rows


# === 구간표 자연 순서(금액·연식·면적·월세 오름차순) ===================================
_BAND_UNIT_FACTOR = {"억": 10000.0, "천만원": 1000.0, "천만": 1000.0, "만원": 1.0, "만": 1.0, "년": 1.0, "㎡": 1.0}


def band_sort_key(label: Any) -> Tuple[int, float]:
    """'1억~15억', '5년 초과 ~ 10년 이하', '60㎡이하', '30~50만원', '15억 초과' 같은 구간명을 하한값 순으로 정렬한다."""
    text = str(label or "").replace(" ", "")
    tokens = list(re.finditer(r"(\d+(?:\.\d+)?)(억|천만원|천만|만원|만|년|㎡)?", text))
    if not tokens:
        return (1, 0.0)                      # 미상·기타는 맨 뒤
    first = tokens[0]
    unit = first.group(2) or next((t.group(2) for t in tokens[1:] if t.group(2)), "")
    value = float(first.group(1)) * _BAND_UNIT_FACTOR.get(unit, 1.0)
    tail = text[first.end():first.end() + 2]
    if tail.startswith(("초과", "이상")):
        value += 0.01                         # 같은 경계값의 '이하' 구간 뒤에 둔다
    elif tail.startswith("미만"):
        value -= 0.01
    return (0, value)


def sort_band_rows(rows: List[Dict[str, Any]], key: str = "band") -> List[Dict[str, Any]]:
    return sorted(rows or [], key=lambda row: band_sort_key(row.get(key)))


def is_jeonnam_gwangju_region(region: Any) -> bool:
    """로컬 입주예정물량 DB는 전남광주(광주·전남) 지역 자료만 담고 있다."""
    text = str(region or "")
    return any(key in text for key in ("광주", "전남", "전라남도"))


def band_commentary_text(rows: List[Dict[str, Any]], subject: str, high_word: str) -> str:
    """구간별 세부 현황 해설: 거래가 많은 구간 → 상승·하락우세 구간 → 구간 순서상 경향 → 소표본 유의."""
    rows = [r for r in (rows or []) if str(r.get("band") or "").strip() and "미상" not in str(r.get("band"))]
    if not rows:
        return ""

    def num(value: Any) -> float:
        try:
            return float(value or 0)
        except Exception:
            return 0.0

    top = max(rows, key=lambda r: num(r.get("count")))
    parts = [f"{subject}별로는 {top.get('band')} 구간이 {int(num(top.get('count')))}건({num(top.get('share')):.1f}%)으로 가장 많음."]
    ups = [(i, str(r.get("band"))) for i, r in enumerate(rows) if str(r.get("판정")) == "상승우세"]
    downs = [(i, str(r.get("band"))) for i, r in enumerate(rows) if str(r.get("판정")) == "하락우세"]
    groups = []
    if ups:
        groups.append("상승우세는 " + "·".join(b for _, b in ups))
    if downs:
        groups.append("하락우세는 " + "·".join(b for _, b in downs))
    if groups:
        parts.append(", ".join(groups) + " 구간으로 나타남.")
    if ups and downs:
        up_pos = sum(i for i, _ in ups) / len(ups)
        down_pos = sum(i for i, _ in downs) / len(downs)
        if down_pos > up_pos:
            parts.append(f"{high_word} 구간일수록 하락우세가 많은 흐름으로 분석됨.")
        elif down_pos < up_pos:
            parts.append(f"{high_word} 구간일수록 상승우세가 많은 흐름으로 분석됨.")
    small = [str(r.get("band")) for r in rows if num(r.get("count")) < 5]
    if small:
        parts.append(f"{'·'.join(small)} 구간은 거래 5건 미만이라 제한적으로 해석할 필요가 있음.")
    return " ".join(parts)


def supply_type_commentary_text(supply: Dict[str, Any], meta: Dict[str, Any]) -> str:
    """표 2(물건유형별 공급 파이프라인 시계열) 해설: 단계별 과거 연간 추이와 당해 누계, 아파트 비중."""
    rows = [r for r in (supply or {}).get("by_type_rows") or [] if r.get("stage") in ("permit", "start", "completion")]
    if not rows:
        return ""
    region = str((meta or {}).get("supply_region") or "시도")
    months = sorted({str(r.get("month")) for r in rows})
    latest = months[-1]
    year_ends = [m for m in months if m.endswith("-12") and m < latest]
    totals: Dict[Tuple[str, str], float] = {}
    apt: Dict[Tuple[str, str], float] = {}
    for r in rows:
        key = (str(r.get("month")), str(r.get("stage")))
        value = float(r.get("value") or 0)
        totals[key] = totals.get(key, 0) + value
        if str(r.get("type")) == "아파트":
            apt[key] = apt.get(key, 0) + value
    latest_label = f"{latest[:4]}년 {int(latest[5:7])}월 누계"
    names = {"permit": "인허가", "start": "착공", "completion": "준공"}
    parts = [f"{region} 공급은 과거연도 12월 누계와 {latest_label}를 물건유형별로 비교함."]
    latest_totals = {}
    for stage, name in names.items():
        past = [(m[:4], totals.get((m, stage), 0)) for m in year_ends if totals.get((m, stage))]
        cur = totals.get((latest, stage), 0)
        latest_totals[stage] = cur
        if not past and not cur:
            continue
        share = apt.get((latest, stage), 0) / cur * 100 if cur else 0
        text = f"{name}{'은' if name.endswith(('공', '준공')) else '는'} "
        if past:
            peak = max(past, key=lambda x: x[1])
            if peak[0] == past[-1][0]:
                text += f"{past[-1][0]}년 {peak[1]:,.0f}호로 {past[0][0]}~{past[-1][0]}년 중 가장 많았으며, "
            else:
                text += f"{past[0][0]}~{past[-1][0]}년 중 {peak[0]}년 {peak[1]:,.0f}호가 가장 많았고 {past[-1][0]}년은 {past[-1][1]:,.0f}호였으며, "
        text += f"{latest_label}는 {cur:,.0f}호(아파트 비중 {share:.1f}%)임."
        parts.append(text)
    permit, start = latest_totals.get("permit", 0), latest_totals.get("start", 0)
    if permit and start and start / permit < 0.3:
        parts.append("당해 착공이 인허가에 비해 크게 적어 인허가 물량의 착공 전환이 늦어지는 흐름으로 분석됨.")
    return " ".join(parts)


# === 변곡점·정책 대조 ===============================================================
def load_housing_policy_events() -> List[Dict[str, Any]]:
    """편집용 정책 일정(housing_policy_events.json)을 읽는다. 없으면 빈 목록."""
    candidates = []
    try:
        here = Path(__file__).resolve().parent
        candidates += [here / "housing_policy_events.json", here.parent / "housing_policy_events.json"]
    except NameError:
        pass
    candidates.append(Path.cwd() / "housing_policy_events.json")
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.insert(0, Path(meipass) / "housing_policy_events.json")
    for path in candidates:
        try:
            if path.is_file():
                data = json.loads(path.read_text(encoding="utf-8"))
                return [e for e in data.get("events", []) if isinstance(e, dict) and e.get("date") and e.get("title")]
        except Exception:
            continue
    return []


def _policy_scope_match(scope: List[str], region: str) -> bool:
    """전국·수도권 대상 부동산 정책은 수도권(서울·경기·인천) 보고서에만 대조하고,
    지방 보고서는 해당 시도·시군구를 명시한 정책(예: 광주 반도체 산단 토지거래허가구역)만 대조한다."""
    is_capital = any(k in region for k in ("서울", "경기", "인천"))
    for token in scope or []:
        token = str(token)
        if token == "전국":
            if is_capital:
                return True
            continue
        if token == "수도권" and any(k in region for k in ("서울", "경기", "인천")):
            return True
        if token not in ("전국", "수도권") and token in region:
            return True
    return False


def detect_market_inflections(rows: List[Dict[str, Any]], base_month: str, min_count: int = 30, min_compare: int = 20) -> List[Dict[str, Any]]:
    """월별 거래량 ±30% 이상 변화, 상승우세↔하락우세 전환, 상승·하락 비율 10%p 이상 변화를 변곡점으로 본다.
    기준월은 신고기한 미도래로 거래량 변화는 제외하고 비율 변화만 본다."""
    def num(row: Dict[str, Any], key: str) -> float:
        try:
            return float(row.get(key) or 0)
        except Exception:
            return 0.0

    rows = [r for r in rows or [] if r.get("month")]
    found = []
    for prev, cur in zip(rows, rows[1:]):
        notes = []
        is_base = str(cur.get("month")) == str(base_month)
        pc, cc = num(prev, "trade_count"), num(cur, "trade_count")
        if not is_base and pc >= min_count and cc >= min_count and abs(cc - pc) / pc >= 0.3:
            notes.append(f"거래량 {pc:,.0f}→{cc:,.0f}건({(cc - pc) / pc * 100:+.1f}%)")
        if num(prev, "compare_count") >= min_compare and num(cur, "compare_count") >= min_compare:
            up0, up1, dn0, dn1 = num(prev, "up_ratio"), num(cur, "up_ratio"), num(prev, "down_ratio"), num(cur, "down_ratio")
            v0, v1 = str(prev.get("판정") or ""), str(cur.get("판정") or "")
            if {v0, v1} == {"상승우세", "하락우세"}:
                notes.append(f"{v0}→{v1} 전환(상승 {up0:.1f}→{up1:.1f}%, 하락 {dn0:.1f}→{dn1:.1f}%)")
            elif abs(up1 - up0) >= 10 or abs(dn1 - dn0) >= 10:
                notes.append(f"상승 비율 {up0:.1f}→{up1:.1f}%, 하락 비율 {dn0:.1f}→{dn1:.1f}%")
        if notes:
            found.append({"month": str(cur.get("month")), "change": ", ".join(notes), "is_base_month": is_base})
    return found


def build_policy_inflection(data: Dict[str, Any]) -> Dict[str, Any]:
    """매매·전세·월세 변곡점마다 직전 2개월~당월에 발표·시행된 정부 정책을 대조한다."""
    meta = data.get("meta", {}) or {}
    region = str(meta.get("region") or "")
    base_month = str(meta.get("base_month") or "")
    events = [e for e in load_housing_policy_events() if _policy_scope_match(e.get("scope") or [], region)]
    ar = data.get("apt_rent", {}) or {}
    series = {
        "apt_trade": (data.get("apt_trade", {}) or {}).get("monthly_rows") or [],
        "apt_jeonse": (ar.get("jeonse", {}) or {}).get("monthly_rows") or [],
        "apt_rent": (ar.get("monthly_rent", {}) or {}).get("monthly_rows") or [],
    }
    markets = {}
    for key, rows in series.items():
        items = []
        for point in detect_market_inflections(rows, base_month):
            try:
                period = pd.Period(point["month"], freq="M")
                start = (period - 2).start_time.date().isoformat()
                end = period.end_time.date().isoformat()
            except Exception:
                continue
            matched = sorted([e for e in events if start <= str(e["date"]) <= end], key=lambda e: str(e["date"]), reverse=True)[:3]
            items.append({**point, "events": [
                {"date": e["date"], "title": e["title"], "summary": e.get("summary", ""),
                 "source": ", ".join(str(src.get("name")) for src in (e.get("sources") or [])[:2])}
                for e in matched
            ]})
        markets[key] = items[-6:]
    return {
        "available": bool(events) or any(markets.values()),
        "events_loaded": len(events),
        "rule": "변곡점: 거래량 ±30% 이상(기준월 제외), 상승우세↔하락우세 전환, 상승·하락 비율 10%p 이상 변화. 정책 대조 기간: 변곡점 월의 직전 2개월~당월. 시점이 겹치는지를 보는 것이며 인과관계를 단정하지 않음.",
        "markets": markets,
    }


def policy_inflection_notes_html(data: Dict[str, Any]) -> Dict[str, str]:
    info = build_policy_inflection(data)
    notes = {}
    for key, items in (info.get("markets") or {}).items():
        if not items:
            notes[key] = ""
            continue
        lis = []
        for item in items:
            if item["events"]:
                ev = "; ".join(
                    f"{html.escape(e['date'])} <b>{html.escape(e['title'])}</b>"
                    + (f"({html.escape(e['summary'][:70])}{'…' if len(e['summary']) > 70 else ''})" if e.get("summary") else "")
                    + (f" [{html.escape(e['source'])}]" if e.get("source") else "")
                    for e in item["events"]
                )
                ev_html = f"<div class='pi-events'>↳ 같은 시기 정책: {ev}</div>"
            else:
                ev_html = "<div class='pi-events pi-none'>↳ 같은 시기 확인된 정부 정책 없음</div>"
            base = " <span class='pi-base'>(기준월·잠정)</span>" if item.get("is_base_month") else ""
            lis.append(f"<li><b>{html.escape(item['month'])}</b>{base} {html.escape(item['change'])}{ev_html}</li>")
        notes[key] = (
            "<div class='policy-inflection-note'><b>🏛 변곡점·정책 대조</b><ul>" + "".join(lis) + "</ul>"
            "<p class='pi-foot'>※ 거래량·상승하락 비율이 크게 바뀐 달의 직전 2개월~당월에 발표·시행된 정부 정책을 대조한 것으로, 인과관계를 단정하지 않음. "
            "기준월은 신고기한 미도래로 거래량 변화는 제외함. 정책 목록은 housing_policy_events.json에서 관리함.</p></div>"
        )
    return notes


def volume_reference_from_frame(frame: pd.DataFrame, value_col: str, base_month: str) -> Dict[str, Any]:
    """최근 12개월·1년 전·2년 전 월평균 거래건수. 기준월은 신고기한 미도래로 빼고 직전 12개월 단위로 비교한다."""
    try:
        if frame is None or frame.empty or value_col not in frame.columns:
            return {}
        grouped = frame[frame[value_col].notna()].groupby("_month")[value_col].count()
        counts = {str(k): int(v) for k, v in grouped.items() if re.fullmatch(r"\d{4}-\d{2}", str(k))}
        if not counts:
            return {}
        first = min(counts)
        base = pd.Period(str(base_month)[:7], freq="M")

        def window(start: int) -> Optional[Dict[str, Any]]:
            months = [str(base - k) for k in range(start, start + 12)]
            avail = [m for m in months if m >= first]
            if not avail:
                return None
            values = [counts.get(m, 0) for m in avail]
            return {"avg": round(sum(values) / len(values), 1), "months": len(avail), "label": f"{min(avail)}~{max(avail)}"}

        recent, prev_year, two_years = window(1), window(13), window(25)

        def pct(a: Optional[Dict[str, Any]], b: Optional[Dict[str, Any]]) -> Optional[float]:
            return round((a["avg"] - b["avg"]) / b["avg"] * 100, 1) if a and b and b["avg"] else None

        return {
            "basis": "기준월(신고기한 미도래) 제외, 직전 12개월 단위 월평균 거래건수",
            "recent": recent, "prev_year": prev_year, "two_years": two_years,
            "vs_prev_year_pct": pct(recent, prev_year), "vs_two_years_pct": pct(recent, two_years),
        }
    except Exception:
        return {}


def policy_chart_events(data: Dict[str, Any], months: List[Any]) -> List[Dict[str, Any]]:
    """차트 x축(월) 안에 드는 정책 발표·시행일을 세로선 위치(월 인덱스+일자 비율)와 짧은 이름으로 돌려준다."""
    region = str(((data or {}).get("meta") or {}).get("region") or "")
    months = [str(m)[:7] for m in months]
    out = []
    for e in load_housing_policy_events():
        if not _policy_scope_match(e.get("scope") or [], region):
            continue
        date_text = str(e.get("date") or "")
        if date_text[:7] not in months:
            continue
        try:
            day = int(date_text[8:10])
            days = pd.Period(date_text[:7], freq="M").days_in_month
            frac = (day - 0.5) / days
        except Exception:
            frac = 0.5
        out.append({"index": months.index(date_text[:7]), "frac": round(frac, 3),
                    "label": str(e.get("short") or e.get("title"))[:14], "date": date_text})
    return sorted(out, key=lambda e: (e["index"], e["frac"]))


def calc_trade_common(files: List[Path], region: str, base_month: str, label: str = "아파트 매매") -> Dict[str, Any]:
    result = {"available": False, "label": label, "latest_month": None, "trade_count": 0, "mom_change": None, "compare_count": 0, "up_ratio": 0, "flat_ratio": 0, "down_ratio": 0, "avg_change_rate": None, "signal": f"{label} 원자료 미확인", "monthly_rows": [], "dong_rows": [], "rowhouse_dong_rows": [], "multiplex_dong_rows": [], "rowhouse_multiplex_dong_rows": [], "rowhouse_major_trade_rows": [], "multiplex_major_trade_rows": [], "multifamily_major_trade_rows": [], "detached_major_trade_rows": [], "presale_major_trade_rows": [], "top_complex_rows": [], "top_up_rows": [], "top_down_rows": [], "top_price_rows": [], "area_band_rows": [], "price_band_rows": [], "building_age_rows": [], "monthly_by_type_rows": []}
    df, cols = prepare_trade_df(files, region, base_month)
    # '비아파트 매매' 라벨에도 '아파트' 문자열이 포함되므로 비아파트 여부를 먼저 판정한다.
    is_non_apt_label = "비아파트" in label
    if is_non_apt_label and "_source_type" in df.columns:
        df = df[df["_source_type"].isin(["연립", "다세대", "단독", "다가구"])].copy()
    elif not is_non_apt_label and "아파트" in label and "_source_type" in df.columns and df["_source_type"].eq("아파트").any():
        df = df[df["_source_type"] == "아파트"].copy()
    if df.empty or "_price" not in df.columns:
        return result
    df = df.dropna(subset=["_price"]).copy()
    if df.empty:
        return result
    df = add_previous_compare(df, "_price")
    monthly = df.groupby("_month").agg(trade_count=("_price", "count"), compare_count=("_prev_value", lambda x: int(x.notna().sum())), avg_change_rate=("_change_rate", "mean")).reset_index().sort_values("_month")
    dc = df[df["_direction"] != "비교불가"].groupby(["_month", "_direction"]).size().unstack(fill_value=0).reset_index()
    monthly = monthly.merge(dc, on="_month", how="left")
    for c in ["상승", "보합", "하락"]:
        if c not in monthly: monthly[c] = 0
    monthly["up_ratio"] = monthly.apply(lambda r: safe_ratio(r["상승"], r["compare_count"]), axis=1)
    monthly["flat_ratio"] = monthly.apply(lambda r: safe_ratio(r["보합"], r["compare_count"]), axis=1)
    monthly["down_ratio"] = monthly.apply(lambda r: safe_ratio(r["하락"], r["compare_count"]), axis=1)

    # "최근 13개월" 시계열은 항상 보고서 기준월(base_month)을 끝으로 하는 연속 13개월 창으로
    # 맞춘다. 분양권전매처럼 최근에 거래가 없어 기준월 데이터가 비는 경우, 과거 거래월이
    # "최근 13개월"로 잘못 표시되던 문제를 막기 위해 거래 없는 달은 0으로 채운다.
    # (거래가 모든 달에 존재하는 매매 등은 결과가 동일하다.)
    try:
        window_months = recent_month_window(base_month, 13)
        if window_months:
            monthly = (monthly.set_index("_month")
                              .reindex(window_months)
                              .rename_axis("_month")
                              .reset_index())
            for c in ["trade_count", "compare_count", "상승", "보합", "하락"]:
                if c in monthly.columns:
                    monthly[c] = monthly[c].fillna(0).astype(int)
            for c in ["up_ratio", "flat_ratio", "down_ratio"]:
                if c in monthly.columns:
                    monthly[c] = monthly[c].fillna(0.0)
            # avg_change_rate는 거래 없는 달은 NaN으로 두어 표/차트에서 '-'로 표기되게 한다.
    except Exception:
        pass

    # 상승·보합·하락 비율 기반 월별 시장 판정(상승우세/보합우세/하락우세/혼조/산출불가)
    monthly["판정"] = monthly.apply(lambda r: direction_verdict(r.get("up_ratio"), r.get("flat_ratio"), r.get("down_ratio")), axis=1)

    # 비아파트 등 복수 원자료가 합쳐진 경우 물건유형별 월별 거래량을 별도 산출
    monthly_by_type_rows = []
    try:
        if "_source_type" in df.columns:
            mt = df.groupby(["_month", "_source_type"]).agg(trade_count=("_price", "count")).reset_index().sort_values(["_month", "_source_type"])
            monthly_by_type_rows = mt.rename(columns={"_month": "month", "_source_type": "type"}).tail(60).to_dict("records")
    except Exception:
        monthly_by_type_rows = []

    base_rows = monthly[monthly["_month"] == base_month]
    base = base_rows.iloc[-1] if not base_rows.empty else monthly.iloc[-1]
    prev_rows = monthly[monthly["_month"] < base["_month"]]
    mom = int(base["trade_count"] - prev_rows.iloc[-1]["trade_count"]) if not prev_rows.empty else None
    base_df = df[df["_month"] == base["_month"]].copy()
    compare_df = base_df[base_df["_direction"] != "비교불가"].copy()
    if "_jibun" not in base_df.columns:
        base_df["_jibun"] = ""

    def rows_by(group_col: str, name_key: str = "name", top=20, source_df: Optional[pd.DataFrame] = None, compare_source_df: Optional[pd.DataFrame] = None):
        out = []
        src_df = base_df if source_df is None else source_df
        cmp_df = compare_df if compare_source_df is None else compare_source_df
        if group_col not in src_df:
            return out
        g = src_df.groupby(group_col).agg(count=("_price", "count")).reset_index().sort_values("count", ascending=False)
        for _, r in g.head(top).iterrows():
            sub = cmp_df[cmp_df[group_col] == r[group_col]]
            cc = len(sub)
            upr = safe_ratio((sub["_direction"]=="상승").sum(), cc)
            flr = safe_ratio((sub["_direction"]=="보합").sum(), cc)
            dnr = safe_ratio((sub["_direction"]=="하락").sum(), cc)
            out.append({name_key: str(r[group_col]), "sigungu": representative_sigungu(src_df, group_col, r[group_col]), "count": int(r["count"]), "share": safe_ratio(r["count"], len(src_df)), "up_ratio": upr, "flat_ratio": flr, "down_ratio": dnr, "판정": direction_verdict(upr, flr, dnr)})
        return out

    dong_rows = rows_by("_dong", "dong")
    rowhouse_dong_rows = []
    multiplex_dong_rows = []
    rowhouse_multiplex_dong_rows = []
    if "_source_type" in base_df.columns:
        src = base_df["_source_type"].astype(str)
        rowhouse_base_df = base_df[src.str.contains("연립", na=False)].copy()
        multiplex_base_df = base_df[src.str.contains("다세대", na=False)].copy()
        rowhouse_multiplex_base_df = base_df[src.str.contains("연립", na=False) | src.str.contains("다세대", na=False)].copy()
        csrc = compare_df["_source_type"].astype(str) if "_source_type" in compare_df.columns else pd.Series([], dtype=str)
        rowhouse_compare_df = compare_df[csrc.str.contains("연립", na=False)].copy() if len(csrc) else compare_df.iloc[0:0].copy()
        multiplex_compare_df = compare_df[csrc.str.contains("다세대", na=False)].copy() if len(csrc) else compare_df.iloc[0:0].copy()
        rowhouse_multiplex_compare_df = compare_df[csrc.str.contains("연립", na=False) | csrc.str.contains("다세대", na=False)].copy() if len(csrc) else compare_df.iloc[0:0].copy()
        rowhouse_dong_rows = rows_by("_dong", "dong", source_df=rowhouse_base_df, compare_source_df=rowhouse_compare_df)
        multiplex_dong_rows = rows_by("_dong", "dong", source_df=multiplex_base_df, compare_source_df=multiplex_compare_df)
        rowhouse_multiplex_dong_rows = rows_by("_dong", "dong", source_df=rowhouse_multiplex_base_df, compare_source_df=rowhouse_multiplex_compare_df)
    base_df["_area_band"] = base_df["_area"].map(classify_area_band)
    price_band_func = make_dynamic_price_band_func(base_df["_price"])
    base_df["_price_band"] = base_df["_price"].map(price_band_func)
    base_df["_building_age_band"] = base_df["_approval_year"].map(lambda x: classify_building_age(x, base_month))
    compare_df["_area_band"] = compare_df["_area"].map(classify_area_band)
    compare_df["_price_band"] = compare_df["_price"].map(price_band_func)
    compare_df["_building_age_band"] = compare_df["_approval_year"].map(lambda x: classify_building_age(x, base_month))
    # 거래건수 순이 아니라 구간 순서(오름차순)로 두어 표와 차트를 읽기 쉽게 한다.
    area_band_rows = sort_band_rows(rows_by("_area_band", "band"))
    price_band_rows = sort_band_rows(rows_by("_price_band", "band"))
    building_age_rows = sort_band_rows(rows_by("_building_age_band", "band"))

    name_key = report_name_key(label)
    top_complex_rows = []
    if not base_df.empty:
        g = base_df.groupby(["_group_name", "_dong"]).agg(sigungu=("_sigungu", "first"), count=("_price", "count"), avg_price=("_price", "mean"), main_area=("_area", "median"), approval_year=("_approval_year", "first")).reset_index().sort_values(["count", "avg_price"], ascending=[False, False])
        for i, (_, r) in enumerate(g.head(10).iterrows(), 1):
            top_complex_rows.append({"rank": i, name_key: r["_group_name"], "sigungu": r.get("sigungu", ""), "dong": r["_dong"], "approval_year": str(r.get("approval_year", "")) if pd.notna(r.get("approval_year", "")) else "", "count": int(r["count"]), "avg_price_eok": money_to_eok_from_manwon(r["avg_price"]), "main_area": None if pd.isna(r["main_area"]) else round(float(r["main_area"]),1), "main_pyeong": area_to_pyeong(r["main_area"])})

    # 기준연도 1월~기준월 누적 거래단지 TOP10(당월 TOP10과 나란히 보는 연간 흐름)
    ytd_complex_rows = []
    base_m = str(base_month)[:7]
    ytd_df = df[(df["_month"] >= base_m[:4] + "-01") & (df["_month"] <= base_m)] if "_month" in df.columns else df.iloc[0:0]
    if not ytd_df.empty:
        g = ytd_df.groupby(["_group_name", "_dong"]).agg(
            sigungu=("_sigungu", "first"), ytd_count=("_price", "count"),
            base_count=("_month", lambda m: int((m == base_m).sum())),
            avg_price=("_price", "mean"), main_area=("_area", "median"), approval_year=("_approval_year", "first"),
        ).reset_index().sort_values(["ytd_count", "avg_price"], ascending=[False, False])
        for i, (_, r) in enumerate(g.head(10).iterrows(), 1):
            ytd_complex_rows.append({
                "rank": i, name_key: r["_group_name"], "sigungu": r.get("sigungu", ""), "dong": r["_dong"],
                "approval_year": str(r.get("approval_year", "")) if pd.notna(r.get("approval_year", "")) else "",
                "ytd_count": int(r["ytd_count"]), "base_month_count": int(r["base_count"]),
                "ytd_avg_price_eok": money_to_eok_from_manwon(r["avg_price"]),
                "main_area": None if pd.isna(r["main_area"]) else round(float(r["main_area"]), 1),
                "main_pyeong": area_to_pyeong(r["main_area"]),
            })

    def major_trade_rows_by_type(type_keyword: str, top: int = 10) -> List[Dict[str, Any]]:
        if "_source_type" not in base_df.columns:
            return []
        sub = base_df[base_df["_source_type"].astype(str).str.contains(type_keyword, na=False)].copy()
        if sub.empty:
            return []
        sub = sub.sort_values("_price", ascending=False)
        rows = []
        for i, (_, r) in enumerate(sub.head(top).iterrows(), 1):
            sigungu = representative_sigungu(pd.DataFrame([r]), "_dong", r.get("_dong"))
            dong = str(r.get("_dong", "") or "").strip()
            jibun = str(r.get("_jibun", "") or "").strip()
            if jibun.lower() in ["nan", "none"]:
                jibun = ""
            location = f"{dong} {jibun}".strip() or dong or "-"
            apt_name = str(r.get("_apt_name", "") or "").strip()
            if apt_name.lower() in ["nan", "none"]:
                apt_name = ""
            row = {
                "rank": i,
                "property_type": type_keyword,
                "sigungu": sigungu,
                "location": location,
                "area": None if pd.isna(r.get("_area")) else round(float(r.get("_area")), 2),
                "pyeong": area_to_pyeong(r.get("_area")),
                "approval_year": str(r.get("_approval_year", "")).replace(".0", "") if pd.notna(r.get("_approval_year", "")) else "",
                "floor": None if pd.isna(r.get("_floor")) else int(r.get("_floor")),
                "contract_date": format_contract_date_display(r.get("_contract_date", ""), r.get("_month", "")),
                "transaction_price_eok": money_to_eok_from_manwon(r.get("_price")),
            }
            if type_keyword in ["연립", "다세대"]:
                row["complex_name"] = apt_name
            rows.append(row)
        return rows

    rowhouse_major_trade_rows = major_trade_rows_by_type("연립")
    multiplex_major_trade_rows = major_trade_rows_by_type("다세대")
    multifamily_major_trade_rows = major_trade_rows_by_type("다가구")
    detached_major_trade_rows = major_trade_rows_by_type("단독")

    def presale_major_trade_rows(top: int = 10) -> List[Dict[str, Any]]:
        sub = base_df[base_df["_price"].notna()].copy()
        if sub.empty:
            return []
        sub = sub.sort_values("_price", ascending=False)
        rows = []
        for i, (_, r) in enumerate(sub.head(top).iterrows(), 1):
            dong = str(r.get("_dong", "") or "").strip()
            jibun = str(r.get("_jibun", "") or "").strip()
            if jibun.lower() in ["nan", "none"]:
                jibun = ""
            location = f"{dong} {jibun}".strip() or dong or "-"
            complex_name = str(r.get("_apt_name", "") or "").strip()
            if complex_name.lower() in ["nan", "none"]:
                complex_name = ""
            rows.append({
                "rank": i,
                "sigungu": representative_sigungu(pd.DataFrame([r]), "_dong", r.get("_dong")),
                "location": location,
                "complex_name": complex_name,
                "area": None if pd.isna(r.get("_area")) else round(float(r.get("_area")), 2),
                "pyeong": area_to_pyeong(r.get("_area")),
                "approval_year": str(r.get("_approval_year", "")).replace(".0", "") if pd.notna(r.get("_approval_year", "")) else "",
                "floor": None if pd.isna(r.get("_floor")) else int(r.get("_floor")),
                "contract_date": format_contract_date_display(r.get("_contract_date", ""), r.get("_month", "")),
                "transaction_price_eok": money_to_eok_from_manwon(r.get("_price")),
            })
        return rows
    presale_major_trade_rows = presale_major_trade_rows()

    def top_cases(direction: Optional[str] = None, sort_col="_change_rate", asc=False, top: Optional[int] = 10):
        sub = base_df.copy()
        if direction:
            sub = sub[sub["_direction"] == direction]
        if sub.empty:
            return []
        sub = sub.sort_values(sort_col, ascending=asc)
        rows = []
        if top is not None:
            sub = sub.head(top)
        for i, (_, r) in enumerate(sub.iterrows(), 1):
            rows.append({"rank": i, name_key: r["_group_name"], "sigungu": r.get("_sigungu", ""), "dong": r["_dong"], "approval_year": str(r["_approval_year"]) if pd.notna(r["_approval_year"]) else "", "households": str(r["_households"]) if pd.notna(r["_households"]) else "", "area": None if pd.isna(r["_area"]) else round(float(r["_area"]),2), "pyeong": area_to_pyeong(r["_area"]), "current_price_eok": money_to_eok_from_manwon(r["_price"]), "prev_price_eok": money_to_eok_from_manwon(r["_prev_value"]), "change_rate": None if pd.isna(r["_change_rate"]) else round(float(r["_change_rate"]),1), "floor": None if pd.isna(r["_floor"]) else int(r["_floor"]), "prev_floor": None if pd.isna(r.get("_prev_floor")) else int(r.get("_prev_floor")), "contract_date": str(r.get("_contract_date", "")), "prev_contract_date": str(r.get("_prev_contract_date", "")), "month": str(r.get("_month", "")), "prev_month": str(r.get("_prev_month", "")), "deal_type": str(r.get("_deal_type") or "") if pd.notna(r.get("_deal_type")) else "", "prev_deal_type": str(r.get("_prev_deal_type") or "") if pd.notna(r.get("_prev_deal_type")) else ""})
        return rows

    signal = "혼조세"
    if float(base["down_ratio"]) >= 50 and pd.notna(base["avg_change_rate"]) and base["avg_change_rate"] < 0:
        signal = "하락 비중 우세"
    elif float(base["up_ratio"]) >= 50 and pd.notna(base["avg_change_rate"]) and base["avg_change_rate"] > 0:
        signal = "상승 비중 우세"
    elif abs(float(base["up_ratio"]) - float(base["down_ratio"])) <= 10:
        signal = "상승·하락 혼조"

    dong_volume_rows = dong_monthly_volume_rows(df, "_price", base_month, 13)

    result.update({"available": True, "latest_month": base["_month"], "trade_count": int(base["trade_count"]), "mom_change": mom, "compare_count": int(base["compare_count"]), "up_ratio": float(base["up_ratio"]), "flat_ratio": float(base["flat_ratio"]), "down_ratio": float(base["down_ratio"]), "avg_change_rate": None if pd.isna(base["avg_change_rate"]) else round(float(base["avg_change_rate"]),1), "signal": signal, "monthly_rows": monthly.tail(13).rename(columns={"_month":"month"}).to_dict("records"), "dong_rows": dong_rows, "rowhouse_dong_rows": rowhouse_dong_rows, "multiplex_dong_rows": multiplex_dong_rows, "rowhouse_multiplex_dong_rows": rowhouse_multiplex_dong_rows, "rowhouse_major_trade_rows": rowhouse_major_trade_rows, "multiplex_major_trade_rows": multiplex_major_trade_rows, "multifamily_major_trade_rows": multifamily_major_trade_rows, "detached_major_trade_rows": detached_major_trade_rows, "presale_major_trade_rows": presale_major_trade_rows, "top_complex_rows": top_complex_rows, "top_up_rows": top_cases("상승", "_change_rate", False), "top_down_rows": top_cases("하락", "_change_rate", True), "top_price_rows": top_cases(None, "_price", False), "all_up_rows": top_cases("상승", "_change_rate", False, None), "all_down_rows": top_cases("하락", "_change_rate", True, None), "all_price_rows": top_cases(None, "_price", False, None), "area_band_rows": area_band_rows, "price_band_rows": price_band_rows, "building_age_rows": building_age_rows, "monthly_by_type_rows": monthly_by_type_rows})
    result["dong_monthly_volume_rows"] = dong_volume_rows
    result["volume_reference"] = volume_reference_from_frame(df, "_price", base_month)
    result["ytd_complex_rows"] = ytd_complex_rows
    result["ytd_period_label"] = f"{base_m[:4]}년 1월~{int(base_m[5:7])}월" if len(base_m) == 7 else ""
    return result


# === v5.8.52 시도+물건유형 근접월 KOSIS 전월세전환율 매칭 ============================
SIDO_SHORT_MAP = {
    "서울특별시": "서울", "서울": "서울",
    "부산광역시": "부산", "부산": "부산",
    "대구광역시": "대구", "대구": "대구",
    "인천광역시": "인천", "인천": "인천",
    "광주광역시": "광주", "광주": "광주",
    "전남광주통합특별시": "광주",
    "전남광주통합특별시 광주권역": "광주",
    "대전광역시": "대전", "대전": "대전",
    "울산광역시": "울산", "울산": "울산",
    "세종특별자치시": "세종", "세종": "세종",
    "경기도": "경기", "경기": "경기",
    "강원특별자치도": "강원", "강원도": "강원", "강원": "강원",
    "충청북도": "충북", "충북": "충북",
    "충청남도": "충남", "충남": "충남",
    "전북특별자치도": "전북", "전라북도": "전북", "전북": "전북",
    "전라남도": "전남", "전남": "전남",
    "경상북도": "경북", "경북": "경북",
    "경상남도": "경남", "경남": "경남",
    "제주특별자치도": "제주", "제주도": "제주", "제주": "제주",
}

ASSET_TO_PUBLIC_TYPE = {
    "아파트": "아파트",
    "오피스텔": "오피스텔",
    "연립": "연립다세대",
    "다세대": "연립다세대",
    "연립다세대": "연립다세대",
    "단독": "단독주택",
    "다가구": "단독주택",
    "단독다가구": "단독주택",
    "단독주택": "단독주택",
    "비아파트": "연립다세대",
}


def apply_confirmed_conversion_rate_if_missing(result: Dict[str, Any], region: str, base_month: str, asset_type: str) -> Dict[str, Any]:
    """확인된 지역별 공표 전환율이 있는데 원자료 매칭이 비어 있을 때 보고서 기준값 보정."""
    out = dict(result or {})
    if out.get("applied_rate") not in (None, "", "-"):
        return out
    region_text = str(region or "")
    public_type = ASSET_TO_PUBLIC_TYPE.get(str(asset_type or "").strip(), str(asset_type or "").strip())
    rate = None
    if "목포" in region_text:
        if public_type == "아파트":
            rate = 6.6667
        elif public_type == "연립다세대":
            rate = 6.0
    if rate is None:
        return out
    applied_month = str(base_month or out.get("period") or "").replace("-", "")
    applied_month = f"{applied_month[:4]}-{applied_month[4:6]}" if len(applied_month) >= 6 and applied_month[:6].isdigit() else (out.get("period") or "")
    out.update({
        "available": True,
        "status": "OK_CONFIRMED_PUBLIC_RATE",
        "period": applied_month,
        "applied_month": applied_month,
        "asset_type": public_type,
        "applied_rate": rate,
        "median_rate": rate,
        "mean_rate": rate,
        "matched_count": out.get("matched_count") or 1,
        "source": "확인된 KOSIS 공표 전월세 전환율",
    })
    return out


def _sido_short_from_region(region: str) -> str:
    if not region:
        return ""
    s = str(region).strip()
    for full, short in SIDO_SHORT_MAP.items():
        if s.startswith(full):
            return short
    head = s.split()[0] if s.split() else s
    return SIDO_SHORT_MAP.get(head, head)


def _period_distance(p1: str, p2: str) -> int:
    """기간 차이를 개월 수로 환산. YYYYMM 문자열 기준."""
    try:
        a = int(str(p1)[:6]); b = int(str(p2)[:6])
        ay, am = a // 100, a % 100
        by, bm = b // 100, b % 100
        return abs((ay - by) * 12 + (am - bm))
    except Exception:
        return 9999


def calc_public_conversion_rate_from_kosis(files: List[Path], region: str, base_month: str, asset_type: str) -> Dict[str, Any]:
    """KOSIS 공표 전월세전환율 엑셀에서 시도+물건유형+근접월 기준 전환율 매칭.

    인자
    - files: classify_files()["conversion_rate_public"] 파일 목록
    - region: '광주광역시 광산구' 등 풀 region 문자열
    - base_month: '2026-04' 또는 '202604' 형태
    - asset_type: '아파트' / '오피스텔' / '연립다세대' / '단독다가구' 등

    반환 dict 구조: calc_rent_common 이 기대하는 conversion_rate 모양과 호환.
    """
    result = {
        "available": False,
        "status": "NO_DATA",
        "period": "",
        "applied_month": "",
        "report_month": str(base_month or ""),
        "asset_type": asset_type,
        "applied_rate": None,
        "median_rate": None,
        "mean_rate": None,
        "matched_count": 0,
        "source": "KOSIS 공표 전월세전환율 시도+물건유형+근접월 자동매칭 (v5.8.52)",
        "formula": "환산월세 = 월세+(보증금×적용전환율÷100÷12)",
        "note": ""
    }
    if not files:
        return result
    try:
        sido_short = _sido_short_from_region(region)
        if not sido_short:
            result["status"] = "NO_REGION"
            return result
        public_type = ASSET_TO_PUBLIC_TYPE.get(str(asset_type or "").strip(), "아파트")

        # base_month 를 YYYYMM 정수로 정규화
        bm_str = str(base_month or "").replace("-", "").strip()
        bm_yyyymm = int(bm_str[:6]) if bm_str[:6].isdigit() else None

        candidates = []
        for p in files:
            try:
                df = read_excel_or_csv(p)
            except Exception:
                continue
            if df is None or df.empty:
                continue
            # 시트 'data' 컬럼 구조: 분류1_객체명, 분류2명(or 분류2), 값, 분류2코드, 분류1코드,
            #                      ... 기간 ... 분류1명 ...
            cols = list(df.columns)
            def _pick(name_options):
                for n in name_options:
                    if n in cols:
                        return n
                return None
            col_region_name = _pick(["분류2명", "분류2", "region_level_2", "C2_NM", "OBJ_L2_NM"])
            col_region_code = _pick(["분류2코드", "region_level_2_code", "C2", "OBJ_L2"])
            col_type_name = _pick(["분류1명", "region_level_1", "C1_NM", "OBJ_L1_NM"])
            col_period = _pick(["기간", "period", "PRD_DE"])
            col_value = _pick(["값", "value", "DT", "DAT"])
            if not all([col_region_name, col_type_name, col_period, col_value]):
                continue

            sido_aliases = {sido_short, f"(구){sido_short}", f"{sido_short}권역"}
            if public_type == "오피스텔":
                # 오피스텔 전환율 표(DT_08002_09_009A)는 분류1=지역별, 분류2=규모별 구조.
                # 지역은 C1_NM/분류1명, 전체 규모는 C2_NM/분류2명에서 매칭한다.
                sub = df[
                    (df[col_type_name].astype(str).isin(sido_aliases)) &
                    (df[col_region_name].astype(str).isin(["전체", "계", "종합"]))
                ].copy()
            else:
                sub = df[
                    (df[col_region_name].astype(str).isin(sido_aliases)) &
                    (df[col_type_name].astype(str) == public_type)
                ].copy()

                # 시도 레벨만 추출 (분류2코드가 a + 1~2자리 숫자). 시군구(a1525 등) 제외.
                if col_region_code and col_region_code in sub.columns:
                    code = sub[col_region_code].astype(str).str.strip()
                    sub = sub[code.str.fullmatch(r"a\d{1,2}")].copy()

            for _, r in sub.iterrows():
                try:
                    period = str(r[col_period]).strip().replace("-", "")[:6]
                    if not period.isdigit():
                        continue
                    val = float(r[col_value])
                except Exception:
                    continue
                candidates.append((int(period), val))

        if not candidates:
            result["status"] = "NO_MATCH"
            result["note"] = f"시도={sido_short}, 물건유형={public_type} 매칭 데이터 없음"
            return result

        # base_month 와 가장 가까운 월 선택. 동률이면 더 최근 월 우선.
        if bm_yyyymm is None:
            candidates.sort(key=lambda x: -x[0])
            chosen_period, chosen_rate = candidates[0]
        else:
            candidates.sort(key=lambda x: (_period_distance(str(x[0]), str(bm_yyyymm)), -x[0]))
            chosen_period, chosen_rate = candidates[0]

        applied_month = f"{str(chosen_period)[:4]}-{str(chosen_period)[4:6]}"
        result.update({
            "available": True,
            "status": "OK_KOSIS_PUBLIC",
            "period": applied_month,
            "applied_month": applied_month,
            "asset_type": public_type,
            "applied_rate": round(float(chosen_rate), 4),
            "median_rate": round(float(chosen_rate), 4),
            "mean_rate": round(float(chosen_rate), 4),
            "matched_count": 1,
            "note": (
                f"KOSIS 공표 전월세전환율 시도='{sido_short}', 물건유형='{public_type}', "
                f"기준월='{base_month}', 적용월='{applied_month}', 적용율={round(float(chosen_rate),4)}%."
                + (" 오피스텔 공표 전환율을 적용함." if str(asset_type or "").strip() == "오피스텔" else "")
            )
        })
        return result
    except Exception as e:
        result["status"] = f"ERROR: {e}"
        return result


def calc_conversion_rate(files: List[Path], region: str, base_month: str, asset_scope: str = "") -> Dict[str, Any]:
    """아파트 신규 전세·신규 월세 동일 주택형 매칭 중위 전월세전환율 산출."""
    result = {
        "available": False,
        "status": "NO_DATA",
        "period": "",
        "matching_key": "시군구+읍면동+지번+전용면적 정수 절사",
        "matched_count": 0,
        "median_rate": None,
        "mean_rate": None,
        "applied_rate": None,
        "formula": "전환율(%) = 월세×12÷(전세보증금-월세보증금)×100",
        "converted_rent_formula": "환산월세 = 월세+(보증금×적용전환율÷100÷12)"
    }
    if not files:
        return result
    try:
        df, cols = prepare_trade_df(files, region, base_month)
        if asset_scope == "아파트" and "_source_type" in df.columns:
            df = df[df["_source_type"] == "아파트"].copy()
        elif asset_scope == "비아파트" and "_source_type" in df.columns:
            df = df[df["_source_type"].isin(["연립", "다세대", "단독", "다가구"])].copy()
        if df.empty:
            return result
        # 계약구분 신규 필터
        if cols.get("contract_type"):
            ct = df[cols["contract_type"]].astype(str)
            df = df[ct.str.contains("신규", na=False) | ~renewal_contract_mask(ct)].copy()

        # 최근 24개월 우선, 부족하면 전체 3년 활용
        months = sorted([m for m in df["_month"].dropna().unique() if str(m).startswith("20")])
        if not months:
            return result
        end_month = base_month if base_month in months else months[-1]
        end_idx = months.index(end_month) if end_month in months else len(months)-1
        start_idx = max(0, end_idx - 23)
        use_months = set(months[start_idx:end_idx+1])
        dfx = df[df["_month"].isin(use_months)].copy()
        result["period"] = f"{min(use_months)}~{max(use_months)}"

        # 매칭키 구성: 시군구 컬럼은 지역 문자열로 대체, 지번 컬럼 있으면 활용
        jibun_col = find_col(dfx, ["지번", "본번", "번지"])
        dfx["_jibun"] = dfx[jibun_col].astype(str) if jibun_col else ""
        dfx["_area_int"] = np.floor(dfx["_area"]).astype("Int64")
        # 매칭 정확도를 위해 표시용 읍면동(_dong)이 아니라 원자료 표기(_dong_full, 리 포함)를 사용한다.
        if "_dong_full" not in dfx.columns:
            dfx["_dong_full"] = dfx["_dong"] if "_dong" in dfx.columns else ""
        dfx["_key"] = region + "|" + dfx["_dong_full"].astype(str) + "|" + dfx["_jibun"].astype(str) + "|" + dfx["_area_int"].astype(str)

        # 전세/월세 구분은 거래구분 열을 최우선 사용한다.
        rent_type_col = cols.get("rent_trade_type") or find_col(dfx, ["거래구분", "전월세구분", "임대차구분", "임대유형", "거래유형"])
        if rent_type_col:
            rt = dfx[rent_type_col].astype(str).str.replace(" ", "", regex=False).str.strip()
            jeonse = dfx[rt.str.contains("전세", na=False)].copy()
            monthly = dfx[rt.str.contains("월세", na=False)].copy()
            result["rent_type_basis"] = f"{rent_type_col} 열(전세·월세)"
        else:
            jeonse = dfx[(dfx["_deposit"].notna()) & ((dfx["_monthly_rent"].isna()) | (dfx["_monthly_rent"] == 0))].copy()
            monthly = dfx[(dfx["_deposit"].notna()) & (dfx["_monthly_rent"].notna()) & (dfx["_monthly_rent"] > 0)].copy()
            result["rent_type_basis"] = "월세금액 보조 판정"

        if jeonse.empty or monthly.empty:
            result["status"] = "NO_MATCH"
            return result

        # 동일 주택형별 전세보증금은 해당 기간 중위값, 월세는 개별 거래 기준
        jg = jeonse.groupby("_key")["_deposit"].median().reset_index().rename(columns={"_deposit": "_jeonse_deposit"})
        mg = monthly.merge(jg, on="_key", how="inner")
        mg = mg[(mg["_jeonse_deposit"] > mg["_deposit"]) & (mg["_monthly_rent"] > 0)].copy()
        if mg.empty:
            result["status"] = "NO_VALID_MATCH"
            return result

        mg["_rate"] = mg["_monthly_rent"] * 12 / (mg["_jeonse_deposit"] - mg["_deposit"]) * 100
        # 현실적인 이상치 제한
        mg = mg[(mg["_rate"] >= 1.0) & (mg["_rate"] <= 20.0)].copy()
        if mg.empty:
            result["status"] = "OUTLIER_ONLY"
            return result

        median_rate = round(float(mg["_rate"].median()), 2)
        mean_rate = round(float(mg["_rate"].mean()), 2)
        result.update({
            "available": True,
            "status": "OK" if len(mg) >= 10 else "PARTIAL",
            "matched_count": int(len(mg)),
            "median_rate": median_rate,
            "mean_rate": mean_rate,
            "applied_rate": median_rate,
            "applied_rate_basis": "최근 24개월 신규 전세·신규 월세 동일 읍면동·지번·전용면적 정수 기준 매칭 중위 전월세전환율",
            "converted_rent_unit": "만원/월",
            "converted_rent_formula": "환산월세 = 월세 + (보증금 × 적용전환율 ÷ 100 ÷ 12)"
        })
        return result
    except Exception as e:
        result["status"] = f"ERROR: {e}"
        return result


def calc_rent_common(files: List[Path], region: str, base_month: str, label: str, conversion_rate: Optional[Dict[str, Any]] = None, is_weekly: bool = False) -> Dict[str, Any]:
    result = {"available": False, "jeonse": {}, "monthly_rent": {}, "signal": f"{label} 전월세 원자료 미확인", "conversion_rate": conversion_rate or {}}
    df, cols = prepare_trade_df(files, region, base_month)
    input_file_rows = []
    if not df.empty and "_source_file" in df.columns:
        for source_file, group in df.groupby("_source_file", dropna=False):
            months = sorted({str(v) for v in group.get("_month", pd.Series(dtype=str)).dropna().tolist() if re.fullmatch(r"20\d{2}-\d{2}", str(v))})
            input_file_rows.append({
                "전월세 원장": str(source_file),
                "읽은 행 수": int(len(group)),
                "최초 계약월": months[0] if months else "-",
                "최종 계약월": months[-1] if months else "-",
            })
    # '비아파트 전월세' 라벨에도 '아파트' 문자열이 포함되므로 비아파트 여부를 먼저 판정한다.
    is_non_apt_label = "비아파트" in label
    if is_non_apt_label and "_source_type" in df.columns:
        df = df[df["_source_type"].isin(["연립", "다세대", "단독", "다가구"])].copy()
    elif not is_non_apt_label and "아파트" in label and "_source_type" in df.columns and df["_source_type"].eq("아파트").any():
        df = df[df["_source_type"] == "아파트"].copy()
    if df.empty:
        return result
    df = df.copy()
    if cols.get("contract_type"):
        ct = df[cols["contract_type"]].astype(str)
        df["_contract_class"] = np.where(renewal_contract_mask(ct), "갱신재계약", "신규")
    else:
        df["_contract_class"] = "신규"
    # 전세·월세 구분은 지정 업로드 원장의 '거래구분'을 최우선 사용한다.
    # 거래구분이 없는 API 원자료에서만 월세금액 0 여부를 보조 기준으로 사용한다.
    trade_type = df.get("_rent_trade_type", pd.Series("", index=df.index)).astype(str).str.replace(" ", "", regex=False).str.strip()
    jeonse_type_mask = trade_type.str.contains("전세", na=False) & ~trade_type.str.contains("월세", na=False)
    rent_type_mask = trade_type.str.contains("월세", na=False)
    recognized_type_mask = jeonse_type_mask | rent_type_mask
    if recognized_type_mask.any():
        jeonse_df = df[jeonse_type_mask & df["_deposit"].notna()].copy()
        rent_df = df[rent_type_mask & df["_deposit"].notna() & df["_monthly_rent"].notna()].copy()
        rent_type_basis = "거래구분 열(전세·월세)"
    else:
        jeonse_df = df[(df["_deposit"].notna()) & ((df["_monthly_rent"].isna()) | (df["_monthly_rent"] == 0))].copy()
        rent_df = df[(df["_deposit"].notna()) & (df["_monthly_rent"].notna()) & (df["_monthly_rent"] > 0)].copy()
        rent_type_basis = "월세금액 보조 판정"
    applied_rate = None
    conversion_status = "NO_CONVERSION_RATE"
    if conversion_rate and conversion_rate.get("applied_rate") is not None:
        applied_rate = float(conversion_rate.get("applied_rate"))
        conversion_status = str(conversion_rate.get("status", "OK"))
    # 환산월세 산식: 월세 + (보증금 × 적용전환율 ÷ 100 ÷ 12)
    # 단위: 보증금=만원, 월세=만원/월, 환산월세=만원/월
    if applied_rate is not None:
        rent_df["_converted_rent"] = rent_df["_monthly_rent"] + (rent_df["_deposit"] * applied_rate / 100 / 12)
    else:
        # 전환율 산출 불가 시 환산월세 컬럼을 원자료 월세로 채우고,
        # 이 경우에만 월세 비교도 원자료 월세 기준으로 수행한다.
        rent_df["_converted_rent"] = rent_df["_monthly_rent"]

    def part_calc(part_df, value_col, part_label):
        if part_df.empty:
            return {"trade_count": 0, "new_count": 0, "renew_count": 0, "compare_count": 0, "up_ratio": 0, "flat_ratio": 0, "down_ratio": 0, "avg_change_rate": None, "signal": f"{part_label} 거래 없음", "value_basis": ("환산월세" if value_col == "_converted_rent" else ("월세" if value_col == "_monthly_rent" else "보증금")), "applied_conversion_rate": applied_rate if value_col in ("_converted_rent", "_monthly_rent") else None, "conversion_status": conversion_status if value_col in ("_converted_rent", "_monthly_rent") else None, "monthly_rows": [], "dong_rows": [], "rowhouse_dong_rows": [], "multiplex_dong_rows": [], "rowhouse_multiplex_dong_rows": [], "rowhouse_major_trade_rows": [], "multiplex_major_trade_rows": [], "multifamily_major_trade_rows": [], "detached_major_trade_rows": [], "top_complex_rows": [], "area_band_rows": [], "deposit_band_rows": [], "monthly_band_rows": [], "converted_rent_band_rows": [], "building_age_rows": [], "top_up_rows": [], "top_down_rows": [], "top_value_rows": [], "monthly_by_type_rows": []}
        part_all = part_df.copy()
        count_table = part_all.groupby("_month").agg(
            trade_count=(value_col, "count"),
            new_count=("_contract_class", lambda x: int((x == "신규").sum())),
            renew_count=("_contract_class", lambda x: int((x == "갱신재계약").sum()))
        ).reset_index().sort_values("_month")
        part_df = part_all[part_all["_contract_class"] == "신규"].copy()
        part_df = add_previous_compare(part_df, value_col) if not part_df.empty else part_df
        if part_df.empty:
            monthly = count_table.copy()
            monthly["compare_count"] = 0
            monthly["avg_change_rate"] = np.nan
        else:
            monthly_cmp = part_df.groupby("_month").agg(compare_count=("_prev_value", lambda x: int(x.notna().sum())), avg_change_rate=("_change_rate", "mean")).reset_index().sort_values("_month")
            monthly = count_table.merge(monthly_cmp, on="_month", how="left")
        dc = part_df[part_df["_direction"]!="비교불가"].groupby(["_month","_direction"]).size().unstack(fill_value=0).reset_index()
        monthly = monthly.merge(dc, on="_month", how="left")
        monthly["compare_count"] = monthly["compare_count"].fillna(0).astype(int)
        for c in ["상승","보합","하락"]:
            if c not in monthly: monthly[c]=0
        monthly["up_ratio"] = monthly.apply(lambda r: safe_ratio(r["상승"], r["compare_count"]), axis=1)
        monthly["flat_ratio"] = monthly.apply(lambda r: safe_ratio(r["보합"], r["compare_count"]), axis=1)
        monthly["down_ratio"] = monthly.apply(lambda r: safe_ratio(r["하락"], r["compare_count"]), axis=1)
        # 상승·보합·하락 비율 기반 월별 시장 판정(상승우세/보합우세/하락우세/혼조/산출불가)
        monthly["판정"] = monthly.apply(lambda r: direction_verdict(r.get("up_ratio"), r.get("flat_ratio"), r.get("down_ratio")), axis=1)

        # 기준월 전세·월세 거래건수가 직전 3개월 대비 과도하게 낮으면
        # 신고기한 미도래에 따른 잠정치일 가능성이 높으므로 분석 기준월에서 제외한다.
        # 예: 2026-05 초순 기준 전세 거래건수가 직전월 500건대에서 69건으로 급감하는 경우
        # 시장 급감으로 단정하지 않고, HTML 표·차트의 최신 분석월은 직전 공표 안정월로 자동 조정한다.
        # 요청사항(2026-06): 전세·월세 실거래 분석은 거래건수가 낮더라도 자동으로
        # 직전월로 롤백하지 않고 보고서 기준월을 그대로 분석 기준월로 사용한다.
        # (과거에는 기준월 거래건수가 직전 3개월 중위의 35% 미만이면 "신고기한
        #  미도래 잠정치"로 판단해 기준월을 분석에서 제외했으나, 매매와 마찬가지로
        #  무조건 보고서 기준월을 사용하도록 변경함.)
        analysis_monthly = monthly.copy()
        partial_note = ""

        monthly_by_type_rows = []
        try:
            # 물건유형별 거래건수는 신규+갱신 전체 거래 기준.
            # 상승·보합·하락 비율만 신규계약 비교가능건 기준으로 계산한다.
            if "_source_type" in part_all.columns:
                mt = part_all.groupby(["_month", "_source_type"]).agg(
                    trade_count=(value_col, "count"),
                    new_count=("_contract_class", lambda x: int((x == "신규").sum())),
                    renew_count=("_contract_class", lambda x: int((x == "갱신재계약").sum()))
                ).reset_index().sort_values(["_month", "_source_type"])
                monthly_by_type_rows = mt.rename(columns={"_month": "month", "_source_type": "type"}).tail(60).to_dict("records")
        except Exception:
            monthly_by_type_rows = []
        base_rows = analysis_monthly[analysis_monthly["_month"] == base_month]
        base = base_rows.iloc[-1] if not base_rows.empty else analysis_monthly.iloc[-1]

        # 세부 현황표의 거래건수·비중은 전세/월세 전체 거래(신규+갱신재계약) 기준으로 산정.
        # 가격 변동률, 상승·보합·하락 비율, TOP 상승·하락 사례는 신규계약 중 직전거래 비교가능건 기준으로 산정.
        base_all_df = part_all[part_all["_month"] == base["_month"]].copy()
        base_df = part_df[part_df["_month"] == base["_month"]].copy()
        comp = base_df[base_df["_direction"]!="비교불가"].copy()
        count_base_df = base_all_df if not base_all_df.empty else base_df

        # 전체 거래 기준 세부표는 add_previous_compare()를 거치지 않은 base_all_df를 사용하므로
        # _group_name 등 비교용 파생 컬럼이 없을 수 있음. 세부 집계 전에 필수 키를 보정한다.
        def _ensure_detail_keys(dfx: pd.DataFrame) -> pd.DataFrame:
            dfx = dfx.copy()
            if "_dong" not in dfx.columns:
                dfx["_dong"] = "미상"
            dfx["_dong"] = dfx["_dong"].map(normalize_dong_to_eupmyeondong).replace("", np.nan).fillna("미상").astype(str)
            if "_dong_full" not in dfx.columns:
                dfx["_dong_full"] = dfx["_dong"]
            dfx["_dong_full"] = dfx["_dong_full"].replace("", np.nan).fillna("미상").astype(str)
            if "_sigungu" not in dfx.columns:
                dfx["_sigungu"] = ""
            dfx["_sigungu"] = dfx["_sigungu"].map(normalize_sigungu_name_value).replace("", np.nan).fillna("").astype(str)
            if "_apt_name" not in dfx.columns:
                dfx["_apt_name"] = ""
            dfx["_apt_name"] = dfx["_apt_name"].mask(dfx["_apt_name"].eq(""), np.nan)
            if "_group_name" not in dfx.columns:
                dfx["_group_name"] = dfx["_apt_name"].fillna(dfx["_dong_full"])
            else:
                dfx["_group_name"] = dfx["_group_name"].mask(dfx["_group_name"].eq(""), np.nan).fillna(dfx["_apt_name"]).fillna(dfx["_dong_full"])
            dfx["_group_name"] = dfx["_group_name"].replace("", np.nan).fillna("미상").astype(str)
            if "_contract_class" not in dfx.columns:
                dfx["_contract_class"] = "신규"
            dfx["_contract_class"] = dfx["_contract_class"].replace("", np.nan).fillna("신규").astype(str)
            if "_jibun" not in dfx.columns:
                dfx["_jibun"] = ""
            dfx["_jibun"] = dfx["_jibun"].replace("", np.nan).fillna("").astype(str)
            if "_contract_date" not in dfx.columns:
                dfx["_contract_date"] = ""
            for c in ["_area", "_deposit", "_monthly_rent", "_approval_year", value_col]:
                if c not in dfx.columns:
                    dfx[c] = np.nan
            return dfx

        count_base_df = _ensure_detail_keys(count_base_df)
        base_df = _ensure_detail_keys(base_df)
        comp = _ensure_detail_keys(comp)
        dong_rows = []
        rowhouse_dong_rows = []
        multiplex_dong_rows = []
        rowhouse_multiplex_dong_rows = []
        top_complex_rows = []
        ytd_complex_rows = []
        area_band_rows = []
        deposit_band_rows = []
        monthly_band_rows = []
        converted_rent_band_rows = []
        building_age_rows = []
        if not count_base_df.empty:
            # 세부표 거래건수 기준:
            # - 거래건수: 신규+갱신재계약 전체
            # - 신규계약 건수 / 갱신재계약 건수: 별도 표시
            # - 상승·보합·하락 비율: 신규계약 중 직전거래 비교가능건 기준
            def _count_agg(df_count: pd.DataFrame, group_cols):
                if df_count is None or df_count.empty:
                    return pd.DataFrame(columns=list(group_cols) + ["count", "new_count", "renew_count"])
                dfx = _ensure_detail_keys(df_count)
                for gc in group_cols:
                    if gc not in dfx.columns:
                        dfx[gc] = "미상"
                    dfx[gc] = dfx[gc].replace("", np.nan).fillna("미상")
                gg = dfx.groupby(group_cols, dropna=False).agg(
                    count=(value_col, "count"),
                    new_count=("_contract_class", lambda x: int((x == "신규").sum())),
                    renew_count=("_contract_class", lambda x: int((x == "갱신재계약").sum()))
                ).reset_index()
                return gg

            g = _count_agg(count_base_df, ["_dong"]).sort_values("count", ascending=False)
            def _dong_rows_from(count_df: pd.DataFrame, compare_source_df: pd.DataFrame):
                rows = []
                gg = _count_agg(count_df, ["_dong"]).sort_values("count", ascending=False)
                for _, r in gg.head(20).iterrows():
                    sub = compare_source_df[compare_source_df["_dong"] == r["_dong"]]
                    cc = len(sub)
                    rows.append({
                        "dong": r["_dong"],
                        "sigungu": representative_sigungu(count_df, "_dong", r["_dong"]),
                        "count": int(r["count"]),
                        "new_count": int(r.get("new_count", 0)),
                        "renew_count": int(r.get("renew_count", 0)),
                        "share": safe_ratio(r["count"], len(count_df)),
                        "up_ratio": safe_ratio((sub["_direction"]=="상승").sum(), cc),
                        "flat_ratio": safe_ratio((sub["_direction"]=="보합").sum(), cc),
                        "down_ratio": safe_ratio((sub["_direction"]=="하락").sum(), cc),
                        "판정": direction_verdict(safe_ratio((sub["_direction"]=="상승").sum(), cc), safe_ratio((sub["_direction"]=="보합").sum(), cc), safe_ratio((sub["_direction"]=="하락").sum(), cc))
                    })
                return rows

            dong_rows = _dong_rows_from(count_base_df, comp)
            if "_source_type" in count_base_df.columns:
                src = count_base_df["_source_type"].astype(str)
                rowhouse_count_df = count_base_df[src.str.contains("연립", na=False)].copy()
                multiplex_count_df = count_base_df[src.str.contains("다세대", na=False)].copy()
                rowhouse_multiplex_count_df = count_base_df[src.str.contains("연립", na=False) | src.str.contains("다세대", na=False)].copy()
                csrc = comp["_source_type"].astype(str) if "_source_type" in comp.columns else pd.Series([], dtype=str)
                rowhouse_comp = comp[csrc.str.contains("연립", na=False)].copy() if len(csrc) else comp.iloc[0:0].copy()
                multiplex_comp = comp[csrc.str.contains("다세대", na=False)].copy() if len(csrc) else comp.iloc[0:0].copy()
                rowhouse_multiplex_comp = comp[csrc.str.contains("연립", na=False) | csrc.str.contains("다세대", na=False)].copy() if len(csrc) else comp.iloc[0:0].copy()
                rowhouse_dong_rows = _dong_rows_from(rowhouse_count_df, rowhouse_comp)
                multiplex_dong_rows = _dong_rows_from(multiplex_count_df, multiplex_comp)
                rowhouse_multiplex_dong_rows = _dong_rows_from(rowhouse_multiplex_count_df, rowhouse_multiplex_comp)

            cg_count = _count_agg(count_base_df, ["_group_name", "_dong"])
            cg_extra = count_base_df.groupby(["_group_name", "_dong"], dropna=False).agg(sigungu=("_sigungu", "first"), avg_value=(value_col, "mean"), avg_deposit=("_deposit", "mean"), avg_monthly=("_monthly_rent", "mean"), main_area=("_area", "median"), approval_year=("_approval_year", "first")).reset_index()
            is_monthly_part = value_col in ("_converted_rent", "_monthly_rent")
            cg = cg_count.merge(cg_extra, on=["_group_name", "_dong"], how="left").sort_values(["count", "avg_value"], ascending=[False, False])
            name_key = report_name_key(label)
            for i, (_, r) in enumerate(cg.head(10).iterrows(), 1):
                top_complex_rows.append({
                    "rank": i, name_key: r["_group_name"], "sigungu": r.get("sigungu", ""), "dong": r["_dong"], "approval_year": str(r.get("approval_year", "")) if pd.notna(r.get("approval_year", "")) else "",
                    "count": int(r["count"]),
                    "new_count": int(r.get("new_count", 0)),
                    "renew_count": int(r.get("renew_count", 0)),
                    "avg_value": None if pd.isna(r["avg_value"]) else round(float(r["avg_value"]), 1),
                    "main_area": None if pd.isna(r["main_area"]) else round(float(r["main_area"]), 2),
                    "main_pyeong": area_to_pyeong(r["main_area"]),
                    "avg_deposit_eok": money_to_eok_from_manwon(r["avg_deposit"]),
                })
                if is_monthly_part:
                    top_complex_rows[-1]["avg_monthly_rent_manwon"] = None if pd.isna(r["avg_monthly"]) else int(round(float(r["avg_monthly"])))
                    if value_col == "_converted_rent":
                        top_complex_rows[-1]["avg_converted_rent_manwon"] = None if pd.isna(r["avg_value"]) else int(round(float(r["avg_value"])))

            # 기준연도 1월~기준월 누적 거래단지 TOP10(신규+갱신 전체 거래 기준, 당월 TOP10과 같은 기준)
            base_m = str(base.get("_month", ""))[:7]
            ytd_df = part_all[(part_all["_month"] >= base_m[:4] + "-01") & (part_all["_month"] <= base_m)].copy() if base_m else part_all.iloc[0:0]
            if not ytd_df.empty:
                yc = _count_agg(ytd_df, ["_group_name", "_dong"])
                ye = _ensure_detail_keys(ytd_df).groupby(["_group_name", "_dong"], dropna=False).agg(
                    sigungu=("_sigungu", "first"), base_count=("_month", lambda m: int((m == base_m).sum())),
                    avg_value=(value_col, "mean"), avg_deposit=("_deposit", "mean"), avg_monthly=("_monthly_rent", "mean"),
                    main_area=("_area", "median"), approval_year=("_approval_year", "first")).reset_index()
                yg = yc.merge(ye, on=["_group_name", "_dong"], how="left").sort_values(["count", "avg_value"], ascending=[False, False])
                for i, (_, r) in enumerate(yg.head(10).iterrows(), 1):
                    row = {
                        "rank": i, name_key: r["_group_name"], "sigungu": r.get("sigungu", ""), "dong": r["_dong"],
                        "approval_year": str(r.get("approval_year", "")) if pd.notna(r.get("approval_year", "")) else "",
                        "ytd_count": int(r["count"]), "ytd_new_count": int(r.get("new_count", 0)), "ytd_renew_count": int(r.get("renew_count", 0)),
                        "base_month_count": int(r.get("base_count") or 0),
                        "ytd_avg_deposit_eok": money_to_eok_from_manwon(r["avg_deposit"]),
                    }
                    if is_monthly_part:
                        row["ytd_avg_monthly_rent_manwon"] = None if pd.isna(r["avg_monthly"]) else int(round(float(r["avg_monthly"])))
                        if value_col == "_converted_rent":
                            row["ytd_avg_converted_rent_manwon"] = None if pd.isna(r["avg_value"]) else int(round(float(r["avg_value"])))
                    row["main_area"] = None if pd.isna(r["main_area"]) else round(float(r["main_area"]), 2)
                    row["main_pyeong"] = area_to_pyeong(r["main_area"])
                    ytd_complex_rows.append(row)

            count_base_df["_area_band"] = count_base_df["_area"].map(classify_area_band)
            deposit_band_func = make_dynamic_deposit_band_func(count_base_df["_deposit"])
            monthly_band_func = make_dynamic_monthly_band_func(count_base_df["_monthly_rent"])
            converted_band_func = make_dynamic_converted_rent_band_func(count_base_df[value_col]) if value_col == "_converted_rent" else deposit_band_func
            count_base_df["_deposit_band"] = count_base_df["_deposit"].map(deposit_band_func)
            count_base_df["_monthly_band"] = count_base_df["_monthly_rent"].map(monthly_band_func)
            count_base_df["_converted_band"] = count_base_df[value_col].map(converted_band_func)
            count_base_df["_building_age_band"] = count_base_df["_approval_year"].map(lambda x: classify_building_age(x, base_month))

            base_df["_area_band"] = base_df["_area"].map(classify_area_band)
            base_df["_deposit_band"] = base_df["_deposit"].map(deposit_band_func)
            base_df["_monthly_band"] = base_df["_monthly_rent"].map(monthly_band_func)
            base_df["_converted_band"] = base_df[value_col].map(converted_band_func)
            base_df["_building_age_band"] = base_df["_approval_year"].map(lambda x: classify_building_age(x, base_month))
            comp["_area_band"] = comp["_area"].map(classify_area_band)
            comp["_deposit_band"] = comp["_deposit"].map(deposit_band_func)
            comp["_monthly_band"] = comp["_monthly_rent"].map(monthly_band_func)
            comp["_converted_band"] = comp[value_col].map(converted_band_func)
            comp["_building_age_band"] = comp["_approval_year"].map(lambda x: classify_building_age(x, base_month))

            def band_rows(col_name, label):
                rr = []
                gg = _count_agg(count_base_df, [col_name]).sort_values("count", ascending=False)
                for _, br in gg.iterrows():
                    subb = comp[comp[col_name] == br[col_name]]
                    cc = len(subb)
                    rr.append({
                        label: br[col_name],
                        "count": int(br["count"]),
                        "new_count": int(br.get("new_count", 0)),
                        "renew_count": int(br.get("renew_count", 0)),
                        "share": safe_ratio(br["count"], len(count_base_df)),
                        "up_ratio": safe_ratio((subb["_direction"]=="상승").sum(), cc),
                        "flat_ratio": safe_ratio((subb["_direction"]=="보합").sum(), cc),
                        "down_ratio": safe_ratio((subb["_direction"]=="하락").sum(), cc),
                        "판정": direction_verdict(safe_ratio((subb["_direction"]=="상승").sum(), cc), safe_ratio((subb["_direction"]=="보합").sum(), cc), safe_ratio((subb["_direction"]=="하락").sum(), cc))
                    })
                return rr

            area_band_rows = sort_band_rows(band_rows("_area_band", "band"))
            deposit_band_rows = sort_band_rows(band_rows("_deposit_band", "band"))
            monthly_band_rows = sort_band_rows(band_rows("_monthly_band", "band"))
            converted_rent_band_rows = sort_band_rows(band_rows("_converted_band", "band"))
            building_age_rows = sort_band_rows(band_rows("_building_age_band", "band"))
        def cases(direction, asc, top: Optional[int] = 10):
            sub = base_df[base_df["_direction"] == direction].copy()
            if sub.empty: return []
            sub = sub.sort_values("_change_rate", ascending=asc)
            rows = []
            name_key = report_name_key(label)
            if top is not None:
                sub = sub.head(top)
            for i, (_, r) in enumerate(sub.iterrows(), 1):
                rows.append({"rank": i, name_key: r["_group_name"], "sigungu": r.get("_sigungu", ""), "dong": r["_dong"], "approval_year": str(r.get("_approval_year", "")) if pd.notna(r.get("_approval_year", "")) else "", "area": None if pd.isna(r["_area"]) else round(float(r["_area"]),2), "pyeong": area_to_pyeong(r["_area"]), "current_value": None if pd.isna(r[value_col]) else round(float(r[value_col]),1), "prev_value": None if pd.isna(r["_prev_value"]) else round(float(r["_prev_value"]),1), "current_deposit": None if pd.isna(r.get("_deposit")) else round(float(r.get("_deposit")),1), "current_monthly_rent": None if pd.isna(r.get("_monthly_rent")) else round(float(r.get("_monthly_rent")),1), "prev_deposit": None if pd.isna(r.get("_prev_deposit")) else round(float(r.get("_prev_deposit")),1), "prev_monthly_rent": None if pd.isna(r.get("_prev_monthly_rent")) else round(float(r.get("_prev_monthly_rent")),1), "change_rate": None if pd.isna(r["_change_rate"]) else round(float(r["_change_rate"]),1), "floor": None if pd.isna(r["_floor"]) else int(r["_floor"]), "prev_floor": None if pd.isna(r.get("_prev_floor")) else int(r.get("_prev_floor")), "contract_date": str(r.get("_contract_date", "")), "prev_contract_date": str(r.get("_prev_contract_date", "")), "month": str(r.get("_month", "")), "prev_month": str(r.get("_prev_month", "")), "deal_type": str(r.get("_deal_type") or "") if pd.notna(r.get("_deal_type")) else "", "prev_deal_type": str(r.get("_prev_deal_type") or "") if pd.notna(r.get("_prev_deal_type")) else ""})
            return rows
        signal = f"{part_label} 혼조"
        if float(base["down_ratio"]) >= 50: signal = f"{part_label} 하락 비중 우세"
        elif float(base["up_ratio"]) >= 40 and float(base["up_ratio"]) > float(base["down_ratio"]): signal = f"{part_label} 상승 우세"
        if value_col == "_converted_rent":
            if applied_rate is not None:
                signal = f"{signal} / 환산월세 기준 적용전환율 {applied_rate:.2f}%"
            else:
                signal = f"{signal} / 전환율 산출 불가로 환산월세 분석 제한"
        if partial_note:
            signal = f"{signal} / {partial_note}"
        def value_cases(top: Optional[int] = 10):
            # 거래금액/보증금 상위 TOP10은 전체 거래(신규+갱신) 기준.
            # 직전계약의 층·일자·금액을 함께 보여주기 위해 전체 계약에서 직전계약을 연결한 뒤 기준월 행을 쓴다.
            sub = count_base_df
            try:
                linked = add_previous_compare(part_all, value_col)
                linked = linked[linked["_month"] == base["_month"]]
                if not linked.empty:
                    sub = linked
            except Exception:
                sub = count_base_df
            sub = sub[sub[value_col].notna()].copy()
            if sub.empty:
                return []
            sub = sub.sort_values(value_col, ascending=False)
            rows = []

            def _safe_round(row, key, digits=1):
                try:
                    v = row.get(key, None)
                    if v is None or pd.isna(v):
                        return None
                    return round(float(v), digits)
                except Exception:
                    return None

            def _safe_int(row, key):
                try:
                    v = row.get(key, None)
                    if v is None or pd.isna(v):
                        return None
                    return int(v)
                except Exception:
                    return None

            name_key = report_name_key(label)
            if top is not None:
                sub = sub.head(top)
            for i, (_, r) in enumerate(sub.iterrows(), 1):
                rows.append({
                    "rank": i,
                    name_key: r.get("_group_name", ""),
                    "sigungu": r.get("_sigungu", ""),
                    "dong": r.get("_dong", ""),
                    "approval_year": str(r.get("_approval_year", "")) if pd.notna(r.get("_approval_year", "")) else "",
                    "area": _safe_round(r, "_area", 2),
                    "pyeong": area_to_pyeong(r.get("_area")),
                    "current_value": _safe_round(r, value_col, 1),
                    "prev_value": _safe_round(r, "_prev_value", 1),
                    "current_deposit": _safe_round(r, "_deposit", 1),
                    "current_monthly_rent": _safe_round(r, "_monthly_rent", 1),
                    "prev_deposit": _safe_round(r, "_prev_deposit", 1),
                    "prev_monthly_rent": _safe_round(r, "_prev_monthly_rent", 1),
                    "deal_type": str(r.get("_deal_type") or "") if pd.notna(r.get("_deal_type")) else "",
                    "contract_class": str(r.get("_contract_class") or ""),
                    "change_rate": _safe_round(r, "_change_rate", 1),
                    "floor": _safe_int(r, "_floor"),
                    "prev_floor": _safe_int(r, "_prev_floor"),
                    "contract_date": str(r.get("_contract_date", "")),
                    "prev_contract_date": str(r.get("_prev_contract_date", "")),
                    "month": str(r.get("_month", "")),
                    "prev_month": str(r.get("_prev_month", ""))
                })
            return rows
        def major_trade_rows_by_type(type_keyword: str, top: int = 10) -> List[Dict[str, Any]]:
            if "_source_type" not in count_base_df.columns:
                return []
            sub = count_base_df[count_base_df["_source_type"].astype(str).str.contains(type_keyword, na=False)].copy()
            sub = sub[sub[value_col].notna()].copy()
            if sub.empty:
                return []
            sub = sub.sort_values(value_col, ascending=False)
            rows = []
            for i, (_, r) in enumerate(sub.head(top).iterrows(), 1):
                dong = str(r.get("_dong", "") or "").strip()
                jibun = str(r.get("_jibun", "") or "").strip()
                if jibun.lower() in ["nan", "none"]:
                    jibun = ""
                location = f"{dong} {jibun}".strip() or dong or "-"
                complex_name = str(r.get("_apt_name", "") or "").strip()
                if complex_name.lower() in ["nan", "none"]:
                    complex_name = ""
                is_monthly_rent_row = value_col in ("_monthly_rent", "_converted_rent")
                row = {
                    "rank": i,
                    "property_type": type_keyword,
                    "sigungu": r.get("_sigungu", ""),
                    "location": location,
                    "area": None if pd.isna(r.get("_area")) else round(float(r.get("_area")), 2),
                    "pyeong": area_to_pyeong(r.get("_area")),
                    "approval_year": str(r.get("_approval_year", "")).replace(".0", "") if pd.notna(r.get("_approval_year", "")) else "",
                    "floor": None if pd.isna(r.get("_floor")) else int(r.get("_floor")),
                    "contract_date": format_contract_date_display(r.get("_contract_date", ""), r.get("_month", "")),
                    "deposit_eok": money_to_eok_from_manwon(r.get("_deposit")),
                }
                if is_monthly_rent_row:
                    row["monthly_rent_manwon"] = None if pd.isna(r.get("_monthly_rent")) else round(float(r.get("_monthly_rent")), 1)
                if type_keyword in ["연립", "다세대"]:
                    row["complex_name"] = complex_name
                if is_monthly_rent_row and "_converted_rent" in r.index:
                    row["converted_rent_manwon"] = None if pd.isna(r.get("_converted_rent")) else round(float(r.get("_converted_rent")), 1)
                rows.append(row)
            return rows

        rowhouse_major_trade_rows = major_trade_rows_by_type("연립")
        multiplex_major_trade_rows = major_trade_rows_by_type("다세대")
        multifamily_major_trade_rows = major_trade_rows_by_type("다가구")
        detached_major_trade_rows = major_trade_rows_by_type("단독")
        return {"trade_count": int(base["trade_count"]), "new_count": int(base.get("new_count", 0)), "renew_count": int(base.get("renew_count", 0)), "compare_count": int(base["compare_count"]), "up_ratio": float(base["up_ratio"]), "flat_ratio": float(base["flat_ratio"]), "down_ratio": float(base["down_ratio"]), "avg_change_rate": None if pd.isna(base["avg_change_rate"]) else round(float(base["avg_change_rate"]),1), "signal": signal, "value_basis": ("환산월세" if value_col == "_converted_rent" else ("월세" if value_col == "_monthly_rent" else "보증금")), "applied_conversion_rate": applied_rate if value_col in ("_converted_rent", "_monthly_rent") else None, "conversion_status": conversion_status if value_col in ("_converted_rent", "_monthly_rent") else None, "converted_rent_formula": "환산월세 = 월세 + (보증금 × 적용전환율 ÷ 100 ÷ 12)" if value_col == "_converted_rent" else None, "analysis_base_month": str(base.get("_month", "")), "detail_count_basis": "읍면동·단지·면적·보증금·월세·사용승인별 세부표는 거래건수·신규계약 건수·갱신재계약 건수를 각각 표시하며, 매매가격·보증금·월세·환산월세 구간은 기준월 자료분포 기반 동적 구간, 상승·보합·하락 비율은 신규계약 비교가능건 기준", "partial_month_note": partial_note, "monthly_rows": analysis_monthly.tail(13).rename(columns={"_month":"month"}).to_dict("records"), "dong_rows": dong_rows, "rowhouse_dong_rows": rowhouse_dong_rows, "multiplex_dong_rows": multiplex_dong_rows, "rowhouse_multiplex_dong_rows": rowhouse_multiplex_dong_rows, "rowhouse_major_trade_rows": rowhouse_major_trade_rows, "multiplex_major_trade_rows": multiplex_major_trade_rows, "multifamily_major_trade_rows": multifamily_major_trade_rows, "detached_major_trade_rows": detached_major_trade_rows, "top_complex_rows": top_complex_rows, "area_band_rows": area_band_rows, "deposit_band_rows": deposit_band_rows, "monthly_band_rows": monthly_band_rows, "converted_rent_band_rows": converted_rent_band_rows, "building_age_rows": building_age_rows, "top_up_rows": cases("상승", False), "top_down_rows": cases("하락", True), "top_value_rows": value_cases(), "all_up_rows": cases("상승", False, None), "all_down_rows": cases("하락", True, None), "all_value_rows": value_cases(None), "monthly_by_type_rows": monthly_by_type_rows, "volume_reference": volume_reference_from_frame(part_all, value_col, str(base.get("_month", ""))), "ytd_complex_rows": ytd_complex_rows, "ytd_period_label": (f"{str(base.get('_month', ''))[:4]}년 1월~{int(str(base.get('_month', ''))[5:7])}월" if len(str(base.get("_month", ""))) >= 7 else "")}
    result.update({
        "available": True,
        "jeonse": part_calc(jeonse_df, "_deposit", f"{label} 전세"),
        # 보증금이 다른 계약끼리 월세만 비교하면 왜곡되므로(예: 보증금 9억·월세 80만 → 보증금 1억·월세 385만)
        # 보고서 설명대로 환산월세(월세+보증금×전환율÷100÷12)로 비교한다. 전환율이 없을 때만 원자료 월세로 비교한다.
        "monthly_rent": part_calc(rent_df, "_converted_rent" if applied_rate is not None else "_monthly_rent", f"{label} 월세"),
        "rent_type_basis": rent_type_basis,
        "rent_type_counts": {"전세": int(len(jeonse_df)), "월세": int(len(rent_df))},
        "property_type_basis": "주용도 우선·건물유형 보완" if cols.get("item_detail") == "주용도" else (cols.get("item_detail") or "파일명 기준"),
        "property_type_counts": {str(k): int(v) for k, v in df["_source_type"].value_counts().to_dict().items()} if "_source_type" in df.columns else {},
        "input_file_rows": input_file_rows,
        "signal": f"{label} 전월세 분석 완료",
    })
    result["jeonse"]["dong_monthly_volume_rows"] = dong_monthly_volume_rows(jeonse_df, "_deposit", base_month, 13)
    return result


# =========================================================
# 6. data_pack 생성
# =========================================================


# =========================================================
# 6-1. 지역특성 기반 LLM 서술 강화 및 반복 방지 보조자료
# =========================================================

REGION_PROFILE_MAP = {
    "서울특별시 강남구": {
        "region_summary": "강남구는 테헤란로 업무축, 압구정·청담 상권, 대치동 학군, 개포·수서 주거축이 함께 형성된 서울 핵심 주거·업무 복합시장임. 고가 아파트 선호와 직주근접 임대수요가 공존하고, 동별로 업무지 배후 소형 임대수요와 학군·재건축 기반 가족형 수요가 차별화됨.",
        "dong_profiles": {
            "역삼동": "테헤란로 업무지와 강남역·역삼역 접근성이 강해 직주근접 소형 임대수요와 오피스 배후수요가 두드러지는 지역.",
            "논현동": "강남 업무·상권 접근성과 다세대·다가구 임대주택 기반이 함께 나타나는 지역으로, 소형 임차수요와 임대수익형 수요가 혼재됨.",
            "개포동": "대규모 아파트 주거지와 정비사업 영향이 큰 지역으로, 가족형 실수요와 신축·준신축 선호가 강하게 작동하는 생활권.",
            "대치동": "교육·학군 선호와 가족형 주거수요가 강한 대표 생활권으로, 전세·매매 모두 학군 수요와 단지 선호도가 가격 흐름에 반영되기 쉬움.",
            "도곡동": "업무지 접근성과 중대형·고가 주거 선호가 함께 나타나는 주거지로, 단지별 품질과 입지 선호에 따른 가격 차별화가 큰 지역.",
            "삼성동": "업무·전시·상업 기능과 고가 주거지가 인접한 지역으로, 직주근접 수요와 고소득 임차·매수 수요가 함께 관찰되는 생활권.",
            "청담동": "고급 주거와 상권 이미지가 강한 지역으로, 고가 주택 선호와 상권 접근성이 시장 해석에 함께 작용함.",
            "압구정동": "한강변 고가 주거지와 재건축 기대가 강한 지역으로, 거래규모가 작아도 가격 수준과 기대 심리가 시장 판단에 영향을 줄 수 있음.",
            "수서동": "광역교통 접근성과 주거 기능이 결합된 지역으로, 전세·월세 수요와 실거주 수요를 함께 살펴볼 필요가 있음.",
            "일원동": "의료·교육·주거 기능이 결합된 생활권으로, 가족형 실수요와 안정적 임차수요가 함께 나타날 수 있는 지역."
        },
        "writing_cautions": [
            "강남구 지역특성은 당월 거래가 확인된 읍면동과 단지 사례에 연결해 사용",
            "학군·상권·업무지 배후수요는 대치동·청담동·역삼동 등 관련 동의 거래가 있을 때만 강조",
            "재건축 기대나 고가 주거 선호는 개포동·압구정동 등 관련 지역의 자료가 있을 때 제한적으로 언급",
            "비아파트는 업무지 배후 소형 임대수요와 다세대·다가구 거래 비중을 중심으로 해석"
        ]
    },
    "광주광역시 남구": {
        "region_summary": "광주 남구는 봉선동 학군·생활편의 선호지, 진월·주월 구축 실수요권, 백운·월산 구도심권, 효천지구 신축·택지지구 성격이 혼재하는 주거시장임.",
        "dong_profiles": {
            "봉선동": "학군·생활편의 선호가 강한 남구 핵심 주거지로, 기존 인기단지와 중대형 실수요가 비교적 견조하게 유지되는 지역.",
            "진월동": "구축 아파트와 실거주 수요가 혼재하며, 가격 민감도가 비교적 높은 생활권.",
            "주월동": "구도심 접근성과 생활편의성이 있으나 구축 단지 비중이 높아 가격 조정 거래가 나타날 수 있는 지역.",
            "백운동": "구도심·노후 주거지 성격이 강하며 정비 기대와 저가 실수요가 혼재하는 지역.",
            "월산동": "구도심권 저가 주거수요와 노후 주택 거래가 함께 나타날 수 있는 지역.",
            "효천지구": "비교적 신축·택지지구 성격이 있어 가족형 실수요와 정주수요가 존재하는 지역.",
            "방림동": "도심 접근성과 구도심 주거 특성이 혼재하는 지역.",
            "양림동": "문화·상업 접근성과 구도심 주거 특성이 혼재하는 지역.",
            "대촌동": "외곽·저밀 주거지 및 토지성 수요 영향이 있는 지역."
        },
        "writing_cautions": [
            "봉선동 특성은 당월 거래비중 또는 주요 사례가 있을 때만 사용",
            "진월동·주월동은 구축 단지 가격 민감도와 연결하되 당월 하락 사례가 확인될 때만 강조",
            "효천지구는 신축·택지지구 수요가 확인되는 경우에만 서술",
            "비아파트는 구도심 저가 실수요·소규모 임대수요 중심으로 제한적 해석"
        ]
    },
    "광주광역시 동구": {
        "region_summary": "광주 동구는 도심·구도심 주거지, 재개발·정비 기대, 대학·상업·의료 접근성이 혼재하는 시장임.",
        "dong_profiles": {
            "계림동": "정비사업과 신축·준신축 주거지 영향이 있는 도심 인접 생활권.",
            "산수동": "구도심 주거지와 실수요·저가 수요가 혼재하는 지역.",
            "학동": "의료·상업 접근성과 기존 주거수요가 혼재하는 지역.",
            "지산동": "구도심·저층 주거 성격과 생활편의 접근성이 혼재하는 지역."
        },
        "writing_cautions": ["정비사업·도심 접근성은 당월 거래가 확인된 동에 한해 연결", "거래건수가 적으면 방향성 단정 금지"]
    },
    "광주광역시 서구": {
        "region_summary": "광주 서구는 상무지구 업무·상업권, 풍암·금호·화정 등 기존 주거지, 농성·양동 구도심권이 혼재하는 시장임.",
        "dong_profiles": {
            "치평동": "상무지구 업무·상업 접근성이 강한 지역.",
            "금호동": "기존 주거 선호와 생활편의 수요가 있는 지역.",
            "풍암동": "대단지·생활편의 기반의 실거주 수요가 있는 지역.",
            "화정동": "구축과 재정비 기대가 혼재하는 기존 주거지.",
            "농성동": "구도심 접근성과 노후 주거지 특성이 혼재하는 지역."
        },
        "writing_cautions": ["상무권 업무수요는 치평·상무권 거래가 있을 때만 연결", "구축 조정은 화정·농성 등 당월 사례가 있을 때만 언급"]
    },
    "광주광역시 북구": {
        "region_summary": "광주 북구는 일곡·양산·첨단 인접권, 문흥·운암·두암 등 대규모 기존 주거지가 넓게 분포하는 시장임.",
        "dong_profiles": {
            "일곡동": "학군·생활편의 기반의 북구 주요 주거지.",
            "양산동": "택지지구 성격과 실거주 수요가 있는 지역.",
            "운암동": "기존 대단지와 정비 기대가 혼재하는 지역.",
            "문흥동": "기존 주거지와 실수요 거래가 중심인 지역.",
            "두암동": "구도심·구축 주거지 성격이 있는 지역."
        },
        "writing_cautions": ["북구는 생활권이 넓으므로 거래 집중 동 기준으로만 특성 연결", "일곡·양산 등 선호지는 당월 거래비중이 있을 때만 언급"]
    },
    "광주광역시 광산구": {
        "region_summary": "광산구는 수완지구(수완·장덕·신창·흑석동), 첨단지구(월계·쌍암·산월동), 운남·신가 등 신도시형 주거권과 하남·평동 산업단지 배후권, 송정·도산 등 광역교통 결절 원도심권, 서부 농촌·외곽권이 함께 구성된 시장임. 광주 군공항 부지가 호남 반도체 첨단국가산단 예정지로 확정되면서 송정 생활권을 중심으로 정책 관심이 커진 반면, 동일 권역이 토지거래허가구역으로 지정되어 기대요인과 규제요인이 동시에 작용하는 구조임.",
        "market_characteristics": [
            "신도시형 주거권과 원도심의 이원 구조: 수완·첨단·운남은 가족형 실수요와 신축·준신축 선호가 유지되는 반면, 송정·도산·신촌 등 원도심권은 구축 비중이 높아 물건별 가격 편차가 큼.",
            "산업단지 배후 임차·실거주 수요: 하남산단·평동산단과 첨단 연구·업무 기능이 배후 임차수요와 중소형 실거주 수요를 지지하는 구조.",
            "광역교통 결절 기능: 광주송정역(KTX·SRT), 광주공항, 도시철도 및 광역도로가 겹치는 송정 생활권이 통근·통행 수요의 결절 역할을 함.",
            "정책 기대와 규제의 동시 작용: 호남 반도체 첨단국가산단 사업 예정지 인근이 토지거래허가구역으로 지정(효력 2026-07-14)되어, 개발 기대와 거래 제약이 함께 나타날 수 있는 국면임."
        ],
        "policy_context_note": "군공항·반도체 국가산단 관련 사실관계는 policy_context.semiconductor_cluster의 확인된 공고 사실만 사용하고, 이전 시기·보상·착공 일정 등 확인되지 않은 절차는 서술하지 않음.",
        "supply_context": "수완·신가·운남 등 기존 대단지 위주 주거권과 신규 입주 예정 물량을 함께 확인해야 하며, move_in_supply가 있을 때만 공급압력으로 연결함. 원도심권은 신규 공급보다 구축 재고 중심 시장임.",
        "demand_drivers": "산업단지·연구단지 종사자의 직주근접 임차·실거주 수요, 광주송정역·광주공항 접근성 기반의 통근수요, 신도시권의 가족형 실수요가 함께 작동함. 반도체 국가산단 관련 기대수요는 실거래 변화가 확인될 때만 제한적으로 연결함.",
        "transport": "광주송정역(KTX·SRT), 광주공항, 도시철도 1호선 및 광역도로가 송정 생활권에 집중되어 있고, 첨단·수완권은 도심 방면 통근 접근성이 상대적으로 양호함.",
        "commercial": "수완지구 상권과 첨단지구 상권이 신도시형 생활편의 중심이며, 송정 원도심은 역세권 전통상권 성격이 강해 생활권별 선호 차이가 나타남.",
        "education": "수완·첨단권이 학군·학원가 접근성으로 가족형 실수요 선호가 상대적으로 높고, 원도심권은 학군보다 가격 접근성과 생활편의가 거래 요인으로 작용하는 편임.",
        "dong_profiles": {
            "수완동": "신도시형 대단지와 생활편의·학군 선호가 결합된 광산구 대표 주거지로, 가족형 실수요가 비교적 견조한 지역.",
            "장덕동": "수완지구 신축·준신축 주거지로 신도시형 생활편의와 가족형 실수요가 집중되는 지역.",
            "신창동": "수완·신창 생활권의 대단지 주거지로 학군·생활편의 선호와 가족형 실수요가 유지되는 지역.",
            "흑석동": "수완 생활권에 인접한 지역으로 신도시권 배후 실거주 수요가 나타나는 지역.",
            "신가동": "수완 생활권과 인접한 기존 주거지로, 실거주 수요와 가격 접근성이 함께 작용하는 지역.",
            "월계동": "첨단2지구 주거지로 첨단 연구·업무단지 배후수요와 학군·생활편의 선호가 결합된 신도시형 생활권.",
            "쌍암동": "첨단지구 상업·업무 중심에 인접한 주거지로, 생활편의 접근성과 준신축 선호가 함께 나타나는 지역.",
            "산월동": "첨단지구 주거 기능이 이어지는 지역으로, 첨단권 배후수요와 기존 단지 실거주 수요가 혼재하는 지역.",
            "비아동": "첨단 인접 구도심 생활권으로 기존 주거수요와 생활편의 접근성이 혼재하는 지역.",
            "하남동": "하남산업단지 배후 주거지로 종사자 임차수요와 중소형 실거주 수요가 혼재하는 지역.",
            "산정동": "하남·산정 일원 신규 주거지 조성 영향이 있는 지역으로, 신규 입주·공급 여건을 함께 확인할 필요가 있는 지역.",
            "월곡동": "원도심 근린 생활권으로 구축 비중이 높고 저가 실거주·임차수요가 함께 나타나는 지역.",
            "도천동": "수완·신가 생활권에 인접한 지역으로 거래규모가 크지 않아 개별 단지 영향이 크게 나타나는 지역.",
            "운남동": "기존 대단지 중심 주거수요가 유지되는 지역으로, 구축 단지의 가격 민감도가 함께 나타남.",
            "송정동": "광주송정역 광역교통 결절과 원도심 주거 기능이 혼재하는 지역으로, 반도체 첨단국가산단 사업 예정지와 인접한 핵심 관찰 생활권.",
            "도산동": "군공항 사업 예정지와 맞닿은 원도심·저층 주거지 성격이 강해, 개별 물건 조건과 정책 영향 관찰이 함께 필요한 지역.",
            "신촌동": "군공항 사업 예정지 인접 생활권으로 거래규모가 크지 않아 개별 거래의 영향이 크게 나타날 수 있는 지역.",
            "소촌동": "송정 생활권과 산업 기능이 인접한 지역으로, 사업 예정지 인접 영향권으로 함께 관찰이 필요한 지역.",
            "우산동": "군공항 예정지 북동측 인접 생활권으로 기존 주거수요와 정책 관심이 함께 나타날 수 있는 지역.",
            "선암동": "송정·평동 방면 교통·산업 연계권으로 거래건수가 적어 방향성 판단에 유의가 필요한 지역.",
            "운수동": "송정 생활권 연계 외곽권으로 개별 물건 특성과 토지성 수요 영향이 상대적으로 큰 지역.",
            "평동": "평동산업단지 배후의 산업·외곽 생활권 성격이 강한 지역.",
            "임곡동": "농촌·외곽 저밀 주거지 성격이 강하고 거래가 제한적인 지역.",
            "동곡동": "외곽·농촌형 생활권으로 거래규모가 작아 보조지표 중심 해석이 필요한 지역.",
            "삼도동": "농촌·외곽권으로 개별 거래의 물건 특성 영향이 큰 지역.",
            "본량동": "농촌·외곽 저밀 주거지로 거래건수가 적으면 가격 방향성 단정이 어려운 지역."
        },
        "writing_cautions": [
            "실거래 읍면동은 법정동 기준이므로 지구명으로 뭉뚱그리지 말고 첨단지구는 월계·쌍암·산월동, 수완지구는 수완·장덕·신창·흑석동으로 연결해 서술",
            "수완·첨단권과 운남은 당월 거래비중 또는 주요 단지 사례가 확인될 때만 신도시형 선호지로 연결",
            "산업 배후수요는 하남·평동·첨단 등 관련 생활권 거래가 확인될 때만 언급",
            "군공항·반도체 첨단국가산단은 policy_context.semiconductor_cluster의 확인된 공고 사실만 사용하고, 이전 시기·보상·착공 등 미확정 절차는 서술하지 않음",
            "송정·도산·신촌·소촌·우산·선암·운수동은 보고서용 분석 영향권이며 공식 수혜지역이 아니므로 등급을 구분해 표현하고, 실거래 변화가 확인될 때만 정책 맥락과 연결",
            "토지거래허가구역 지정은 투기방지 규제이므로 가격 상승 근거로 사용하지 말고 규제효과와 산단 조성 기대효과를 구분",
            "임곡·동곡·삼도·본량 등 농촌·외곽권은 거래건수가 적으면 방향성을 단정하지 말고 개별 물건 특성 중심으로 제한적 해석"
        ]
    },

    "무안군": {
        "region_summary": "무안군은 남악·오룡 신도시권, 무안읍 기존 생활권, 일로·삼향 등 광역교통·행정 배후권, 해제·청계·현경 등 농촌·외곽권이 구분되는 시장임. 특히 남악·오룡권은 전남도청·행정기관·신축 주거지 수요와 연계되고, 무안읍·일로읍은 기존 생활권 실수요와 가격 민감도가 함께 나타날 수 있음.",
        "dong_profiles": {
            "삼향읍": "남악신도시와 도청·행정기관 배후수요, 생활편의 기반의 아파트 실거주 수요가 집중되는 지역.",
            "오룡지구": "신축·준신축 주거지와 가족형 실수요, 신규 공급 영향이 강한 신도시형 생활권.",
            "남악리": "전남도청·행정기관 접근성과 생활편의 선호가 반영되는 남악 생활권 중심 지역.",
            "무안읍": "군청·기존 상권·교육·생활편의 기반의 기존 생활권으로, 구축 주택과 실수요 거래가 혼재하는 지역.",
            "일로읍": "남악·오룡 및 목포권 접근성이 있는 배후 생활권으로, 기존 주거수요와 개발 기대가 병존하는 지역.",
            "청계면": "목포대 등 교육·생활권 영향과 농촌형 주거수요가 혼재하는 지역.",
            "현경면": "농촌·외곽 주거지 성격이 강하고 거래건수가 적을 경우 가격 방향성 판단에 유의가 필요한 지역.",
            "해제면": "해안·농촌형 외곽 생활권으로 토지성·개별주택 수요가 일부 반영될 수 있는 지역.",
            "몽탄면": "농촌·외곽 저밀 주거지 성격이 강하며 개별 거래의 물건 특성 영향이 큰 지역.",
            "운남면": "해안·농촌형 외곽권으로 거래규모가 제한적일 경우 보조지표 중심 해석이 필요한 지역."
        },
        "writing_cautions": [
            "남악·오룡권은 당월 거래비중, 주요 단지 사례, 신규 공급 영향이 확인될 때만 신도시·행정 배후수요와 연결",
            "오룡지구 신축 공급 영향은 전세·월세 수급 또는 매매 가격 조정 사례가 있을 때만 제한적으로 사용",
            "무안읍·일로읍은 기존 생활권 실수요와 가격 민감도 관점에서 해석하되 당월 수치 근거와 연결",
            "청계·현경·해제·몽탄·운남 등 농촌·외곽권은 거래건수가 적으면 가격 방향성을 단정하지 말고 개별 물건 특성·토지성 수요 중심으로 제한적 해석",
            "비아파트는 거래건수, 대지면적, 건축연도, 보증금·월세 수준 등 보조지표와 함께 해석"
        ]
    },
    "목포시": {
        "region_summary": "목포시는 옥암·상동 등 신도심 선호지, 하당 생활권, 원도심·항만권, 북항·연산권 주거지가 혼재하는 시장임. 신규 아파트 공급이 제한적인 가운데 민간임대아파트 사업 차질로 임차 대기수요 일부가 기존 아파트 매매수요로 전환되는 구조적 흐름이 시장 배경에 깔려 있음.",
        "market_characteristics": [
            "공급부족·민간임대 부도형 전환수요: 신규 아파트 공급이 제한적인 가운데 민간임대아파트(석현동 더숲, 2022년 사용승인, 121세대) 부도로 임대물량이 차질을 빚어, 임차 대기수요 일부가 기존 아파트 매매로 전환되는 흐름.",
            "인접 신도시로의 신축 수요 유출: 목포의 신축 선호 수요 일부가 무안 남악신도시·오룡지구 신축 아파트로 이동하는 수요 유출 경향.",
            "구축→신축 상향 이동 수요 증가: 시내 구축 아파트에서 신축·준신축 아파트로 갈아타려는 주거 상향(상향 이동) 수요가 증가하는 흐름."
        ],
        "dong_profiles": {
            "옥암동": "신도심·학군·생활편의 선호가 강한 목포 핵심 주거지.",
            "상동": "하당 생활권과 상업·생활편의 접근성이 있는 주요 주거지.",
            "연산동": "기존 주거지와 준신축 수요가 혼재하는 지역.",
            "석현동": "신흥 주거지로 실수요가 형성되나, 민간임대아파트 더숲(2022년 사용승인, 121세대) 부도로 임대 공급 차질이 발생한 지역.",
            "용해동": "기존 주거지와 실수요 거래가 중심인 지역."
        },
        "supply_context": "최근 신규 아파트 공급이 제한적이어서 구조적 공급부족 성격이 있으며, 민간임대아파트(석현동 더숲, 2022년 사용승인, 121세대) 부도로 예정 임대물량 공급이 차질을 빚어 공급부족 압력이 가중되는 것으로 판단됨.",
        "demand_drivers": "목포의 신축 선호 수요 일부가 인접 무안 남악신도시·오룡지구의 신축 아파트로 이동하는 수요 유출 측면과, 민간임대아파트 공급 차질로 임차수요 일부가 목포 기존 아파트 매매로 전환·유입되는 측면이 함께 작용함. 또한 시내 구축 아파트에서 신축·준신축으로 갈아타는 주거 상향 이동 수요가 증가하는 흐름이 관찰되는 양면적·상향이동형 배후수요 구조임.",
        "transport": "원도심의 KTX·SRT 목포역(호남선 종착)과 서해안고속도로 종점, 무안국제공항 인접 등 광역교통 결절 기능이 있는 권역임.",
        "commercial": "하당 신도심(평화광장 일대) 상업·생활편의 중심과 원도심 전통 상권이 구분되어 생활권별 선호 차이가 나타남.",
        "education": "옥암·상동 등 하당 신도심권이 학군·학원가 접근성으로 가족형 실수요 선호가 상대적으로 높은 편임.",
        "writing_cautions": [
            "민간임대아파트 부도(석현동 더숲, 2022년, 121세대)와 이에 따른 매매수요 전환은 목포 시장의 구조적 배경으로 활용하되, 당월 거래사례·가격흐름과 연결해 과대단정은 피함",
            "무안 남악신도시·오룡지구 신축으로의 수요 유출은 목포 신축 선호 수요 이동의 배경으로 활용하되, 당월 거래 변화와 연결해 해석",
            "교통·상권·학군은 거래 집중 읍면동과 연결될 때만 언급하고 확인되지 않은 외부 호재로 확대하지 않음",
            "원도심 비아파트는 거래건수 기준으로 신중히 해석"
        ]
    },
    "나주시": {
        "region_summary": "나주시는 빛가람혁신도시 에너지권, 나주 원도심 생활권, 왕곡·금천·산포 일원의 에너지산단 산업권, 남평·노안의 광주인접권, 영산강 유역 농생명·농촌권이 뚜렷하게 구분되는 시장임. 이전 공공기관과 에너지밸리가 집적된 혁신도시 신시가지와 노후 재고 중심의 원도심이 서로 다른 흐름을 보이며, 공공기관 2차 이전(혁신도시 시즌2)과 에너지국가산단 조성이 각각 계획·착수 단계에 있어 기대요인과 확정되지 않은 절차가 함께 존재하는 구조임.",
        "market_characteristics": [
            "혁신도시 신시가지와 원도심의 이원 구조: 빛가람동은 계획 주거·상업·업무지가 함께 조성된 신시가지로 아파트 거래가 집중되는 반면, 송월·성북·이창 등 원도심 생활권은 기존 재고와 노후주택 비중이 높아 수요·가격 흐름을 구분해 봐야 함.",
            "공공기관 2차 이전(혁신도시 시즌2): 에너지·농생명·정보통신 분야 공공기관 2차 이전이 정부 계획 수립 단계에 있고 2027년 이전 착수가 예정되어 있으나, 이전 대상 기관과 배치는 확정되지 않은 상태임.",
            "에너지산단 배후수요: 나주혁신산업단지가 에너지 기업·연구기관 집적 기반을 갖추고 있고, 나주에너지국가산업단지는 2025년 7월 승인 후 보상·부지 조성 착수 단계로 2030년대 초반까지 단계 조성이 예정되어 있음.",
            "광주 인접 통근 배후주거: 남평읍·노안면은 광주 남구와 접한 권역으로 광주 생활권 통근 배후주거와 개발가능지 수요가 혼재함.",
            "규제와 인접 개발의 동시 작용: 2026년 7월 14일 효력이 발생한 토지거래허가구역에 나주시 97.93㎢가 직접 포함되었고, 호남권 반도체 첨단국가산단 부지는 광주 군공항(광산구) 소재로 나주시는 인접·연계 영향권임."
        ],
        "policy_context_note": "혁신도시 시즌2, 에너지국가산단, 토지거래허가구역 관련 사실관계는 위 market_characteristics에 적힌 확인된 단계 표기 범위에서만 사용하고, 이전 대상 기관·이전 시기·산단 착공 및 기업 입주 일정 등 미확정 사항은 서술하지 않음.",
        "supply_context": "혁신도시권은 계획 주거용지 기반의 아파트 공급이 중심이고, 원도심은 신규 공급보다 기존 재고와 노후주택·빈집 비중이 중심임. 신규 입주 예정물량은 move_in_supply가 확인될 때만 수급 배경으로 연결함.",
        "demand_drivers": "이전 공공기관 임직원과 에너지밸리 입주기업 종사자의 주거·임차 수요, 혁신산단·에너지국가산단 종사자의 배후수요, 광주 통근 배후주거 수요(남평·노안), 원도심 생활·상업 수요가 함께 작동함. 공공기관 2차 이전 기대수요는 확정 전이므로 실거래 변화가 확인될 때만 제한적으로 연결함.",
        "transport": "나주역(KTX·SRT)과 광주송정역 연계, 시외버스터미널과 시가지 간선도로, 광주~나주 간선도로 및 광역 산업물류 도로망이 권역별로 다르게 작용함.",
        "commercial": "혁신도시 계획 상업지와 원도심·영산포 전통 상권이 구분되며, 혁신도시권은 상가 공실과 유동인구, 원도심은 상권 침체와 정비·도시재생 진행상황을 함께 확인할 필요가 있음.",
        "education": "혁신도시권은 계획도시로 상업·교육·의료 이용수요가 함께 형성된 반면, 원도심과 농촌권은 학군보다 가격 접근성과 생활편의가 거래 요인으로 작용하는 편임.",
        "dong_profiles": {
            "빛가람동": "이전 공공기관과 에너지밸리 종사자의 주거·상업 수요가 집중되는 혁신도시 중심 생활권으로, 나주 아파트 거래가 가장 많이 형성되는 지역.",
            "금천면": "혁신도시에 연접한 권역으로 석전리 일원 아파트 등 혁신도시 배후 주거수요와 산단 관련 수요가 함께 나타나는 지역.",
            "송월동": "나주 원도심 생활권의 기존 주거지로 실거주·임차 수요가 유지되는 지역.",
            "성북동": "원도심 생활권으로 저가 실수요와 노후주택 거래가 혼재하는 지역.",
            "이창동": "원도심 생활권의 기존 주거지로 생활편의 접근성이 거래 요인으로 작용하는 지역.",
            "영강동": "영산포 상권과 인접한 원도심 생활권으로 기존 주거·상업 수요가 혼재하는 지역.",
            "금남동": "나주 원도심 중심 생활권으로 기존 시가지 주거수요가 있는 지역.",
            "영산동": "영산포 일원 원도심 생활권으로 노후 주거지 성격이 함께 나타나는 지역.",
            "대호동": "나주 시가지에 인접한 주거지로 실거주 수요가 확인되는 지역.",
            "삼영동": "나주 시가지 생활권의 주거지로 아파트 전세 거래가 형성되는 지역.",
            "죽림동": "원도심에 인접한 기존 시가지 주거지로 거래규모가 크지 않은 지역.",
            "금계동": "나주 원도심 인접 시가지 생활권으로 거래가 제한적인 지역.",
            "남내동": "나주 원도심 인접 시가지 생활권으로 개별 물건 특성 영향이 큰 지역.",
            "용산동": "나주 시가지 외곽 생활권으로 거래건수가 적어 방향성 판단에 유의가 필요한 지역.",
            "남평읍": "광주 남구와 접한 통근 배후주거권으로, 광주 생활권 수요와 전원·개발가능지 수요가 혼재하는 지역.",
            "노안면": "광주 인접 권역으로 통근 배후수요와 개발압력이 함께 나타날 수 있는 지역.",
            "왕곡면": "나주에너지국가산업단지 조성지가 위치한 산업권으로, 조성 진행상황과 주거 거래를 분리해 볼 필요가 있는 지역.",
            "산포면": "혁신산단·에너지산단과 연접한 산업권으로 종사자 배후수요를 점검할 지역.",
            "다시면": "영산강 유역 농생명·농촌권으로 거래가 희소해 개별 물건 특성 중심 해석이 필요한 지역.",
            "세지면": "농생명·농촌권으로 농업 기반 수요가 중심인 지역.",
            "반남면": "농생명·농촌권으로 거래규모가 작아 방향성 단정이 어려운 지역.",
            "공산면": "농생명·농촌권 취락 중심 지역으로 개별 거래 영향이 큰 지역.",
            "동강면": "영산강 유역 농촌권으로 농업 기반 수요가 중심인 지역.",
            "문평면": "농공단지와 취락이 분포하는 농촌권 지역.",
            "다도면": "농촌·외곽 저밀 지역으로 거래가 제한적인 지역.",
            "봉황면": "농생명·농촌권으로 농업 기반 생활수요가 중심인 지역."
        },
        "writing_cautions": [
            "혁신도시 시즌2(공공기관 2차 이전)는 정부 계획 수립 단계이고 이전 대상 기관·배치가 확정되지 않았으므로 확정 호재로 서술하지 않고, 기대를 가격 상승의 직접 원인으로 단정하지 않으며 인구 정체·상가 공실 등 정주여건 요인을 함께 확인",
            "혁신도시 신시가지(빛가람동)와 원도심 생활권(송월·성북·이창 등)의 수요·가격 흐름은 반드시 구분해 서술",
            "나주에너지국가산업단지는 보상·조성 착수 단계이므로 기업 입주·가동이 확인되지 않은 기대를 실수요로 표현하지 않음",
            "2026년 7월 14일 효력이 발생한 토지거래허가구역에 나주시가 직접 포함된 사실은 규제 요인으로만 사용하고 가격 상승 근거로 쓰지 않으며, 반도체 첨단국가산단 부지는 광산구 소재이고 나주시는 인접·연계 영향권임을 구분",
            "실거래 읍면동은 법정동 기준이며 읍·면 지역은 리 단위가 아니라 읍·면 단위로 합산해 서술",
            "남평·노안 등 광주 인접권은 통근 배후수요와 개별 필지 거래를 구분하고, 농촌·외곽권은 거래건수가 적으면 방향성 단정 금지"
        ]
    }
}

# 행정구역 통합·개편으로 조사지역 문자열의 시도명이 달라져도 같은 시군구 프로파일을 찾도록
# 시도명 별칭을 관리한다. (예: '광주광역시 광산구' 프로파일 ↔ '전남광주통합특별시 광산구' 조사지역)
SIDO_NAME_ALIASES = {
    "광주광역시": ("광주광역시", "전남광주통합특별시", "광주전남통합특별시", "광주전남특별시", "전남광주특별시"),
    "전라남도": ("전라남도", "전남광주통합특별시", "광주전남통합특별시", "광주전남특별시", "전남광주특별시"),
}

# 여러 시도에 동일 명칭이 존재해 시도명 없이 단독 매칭하면 안 되는 시군구.
AMBIGUOUS_SIGUNGU_NAMES = {"동구", "서구", "남구", "북구", "중구", "중원구", "상록구"}


def _region_profile_match_candidates(key: str) -> List[str]:
    """프로파일 키 하나에 대해 매칭 후보 문자열 목록을 만든다."""
    parts = str(key or "").split()
    if len(parts) < 2:
        return [str(key or "")]
    sido, sigungu = parts[0], parts[-1]
    aliases = SIDO_NAME_ALIASES.get(sido, (sido,))
    return [f"{alias} {sigungu}" for alias in aliases]


def get_region_profile(region: str) -> Dict[str, Any]:
    """조사지역에 맞는 지역 특성 프로파일 반환. 일치 항목이 없으면 빈 프로파일 반환."""
    r = " ".join(str(region or "").split())
    for key, val in REGION_PROFILE_MAP.items():
        if key in r or r in key:
            out = dict(val)
            out["matched_region"] = key
            return out
    # 시도명 별칭(행정구역 통합·개편) 기준 재매칭
    for key, val in REGION_PROFILE_MAP.items():
        for cand in _region_profile_match_candidates(key):
            if cand and cand in r:
                out = dict(val)
                out["matched_region"] = key
                return out
    # 최후 보정: 시도명이 예상 밖으로 바뀐 경우에도 명칭이 고유한 시군구는 시군구명만으로 매칭
    region_tail = r.split()[-1] if r.split() else ""
    if region_tail and region_tail not in AMBIGUOUS_SIGUNGU_NAMES:
        for key, val in REGION_PROFILE_MAP.items():
            if str(key).split()[-1] == region_tail:
                out = dict(val)
                out["matched_region"] = key
                return out
    return {
        "matched_region": "",
        "region_summary": "당월 data_pack의 읍면동별 거래비중·주요 사례 중심으로 지역 맥락을 작성.",
        "dong_profiles": {},
        "writing_cautions": ["확인되지 않은 교통·상권·학군 요인을 임의로 확대하지 말 것", "거래 상위 읍면동과 주요 거래사례에 근거한 해석만 작성"]
    }


MARKET_FIELD_CONTEXT_FILENAME = "market_field_context.json"

# 조사자가 직접 작성하는 현장 시황(주간·월간·오피스텔)의 표준 항목.
MARKET_FIELD_SECTION_KEYS = ("sale", "jeonse", "rent")
MARKET_FIELD_ITEM_KEYS = (
    "overall", "listing_change", "transaction_change", "price_change",
    "buyer_behavior", "seller_behavior", "tenant_behavior", "landlord_behavior",
    "rising_areas", "falling_areas", "major_cases", "local_issue",
)
# 보고서 종류별로 조사 시황을 분리 보관한다. 주간 주택시황을 월간·오피스텔 보고서에
# 그대로 끌어다 쓰면 안 되므로 dataset을 명시적으로 구분한다.
MARKET_FIELD_DATASET_LABELS = {
    "housing_weekly": "주택 주간 조사 시황",
    "housing_monthly": "주택 월간 조사 시황",
    "officetel_monthly": "오피스텔 월간 조사 시황",
}


def normalize_period_key(value: Any) -> str:
    """기준기간 표기를 비교용 표준형으로 바꾼다.

    월간: '2026-07', '202607', '2026년 7월' -> '2026-07'
    주간: '2026년 8월 4주차', '2026-08 4주' -> '2026-08-W4'
    """
    s = " ".join(str(value or "").split())
    if not s:
        return ""
    mw = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})\s*주", s)
    if mw:
        return f"{mw.group(1)}-{int(mw.group(2)):02d}-W{int(mw.group(3))}"
    mm = re.search(r"(\d{4})\D*(\d{1,2})", s)
    if mm:
        return f"{mm.group(1)}-{int(mm.group(2)):02d}"
    return s


def market_field_dataset_key(is_weekly: bool = False, is_officetel: bool = False) -> str:
    """보고서 종류에 대응하는 조사 시황 dataset 키."""
    if is_officetel:
        return "officetel_monthly"
    return "housing_weekly" if is_weekly else "housing_monthly"


def find_market_field_context_file() -> Optional[Path]:
    """실행파일 옆 외부 market_field_context.json을 우선 사용한다."""
    candidates = [BASE_DIR / MARKET_FIELD_CONTEXT_FILENAME, INPUT_DIR / MARKET_FIELD_CONTEXT_FILENAME]
    try:
        candidates.append(Path(__file__).resolve().parent / MARKET_FIELD_CONTEXT_FILENAME)
    except Exception:
        pass
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / MARKET_FIELD_CONTEXT_FILENAME)
    candidates.extend([Path.cwd() / MARKET_FIELD_CONTEXT_FILENAME, BASE_DIR.parent / MARKET_FIELD_CONTEXT_FILENAME])
    seen = set()
    for candidate in candidates:
        key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        try:
            if candidate.exists():
                return candidate
        except Exception:
            continue
    return None


def _market_field_region_map(raw: Dict[str, Any], dataset_key: str) -> Tuple[Dict[str, Any], str]:
    """dataset별 시군구 조사 시황 맵을 꺼낸다. 반환값은 (맵, 사용된 dataset 표시)."""
    datasets = raw.get("datasets")
    if isinstance(datasets, dict):
        node = datasets.get(dataset_key)
        if isinstance(node, dict):
            regions = node.get("regions")
            if not isinstance(regions, dict):
                regions = node
            return (regions if isinstance(regions, dict) else {}), dataset_key
        return {}, ""
    # 구버전 파일(datasets 없이 regions만 존재)은 주택 보고서용으로만 인정한다.
    legacy = raw.get("regions")
    if isinstance(legacy, dict) and dataset_key != "officetel_monthly":
        return legacy, "regions(구버전)"
    return {}, ""


def _entry_period_records(entry: Dict[str, Any]) -> List[Tuple[str, Dict[str, Any]]]:
    """조사 시황 항목에서 (기간표기, 기록) 목록을 만든다.

    periods 하위에 기간별 기록을 쌓아두면 기준월에 맞는 기록을 고를 수 있다.
    periods가 없는 구조는 항목 자체를 단일 기록으로 본다(구버전 호환).
    """
    if not isinstance(entry, dict):
        return []
    periods = entry.get("periods")
    if not isinstance(periods, dict):
        return [(str(entry.get("period_label") or ""), entry)]
    shared = {k: v for k, v in entry.items() if k != "periods"}
    out = []
    for pk, rec in periods.items():
        if str(pk).startswith("_") or not isinstance(rec, dict):
            continue
        merged = dict(shared)
        merged.update(rec)
        if not merged.get("period_label"):
            merged["period_label"] = str(pk)
        out.append((str(merged.get("period_label") or pk), merged))
    return out


def _select_period_record(entry: Dict[str, Any], want_key: str) -> Tuple[Dict[str, Any], bool]:
    """기준월에 해당하는 기록을 고른다. 없으면 (가장 최근 기록, False)."""
    recs = _entry_period_records(entry)
    if not recs:
        return {}, False
    if want_key:
        for label, rec in recs:
            if normalize_period_key(label) == want_key:
                return rec, True
    recs = sorted(recs, key=lambda x: normalize_period_key(x[0]))
    return recs[-1][1], False


def _match_market_field_region(regions: Dict[str, Any], region: str) -> Tuple[str, Dict[str, Any]]:
    """region_profile과 동일한 규칙(시도명 별칭 포함)으로 조사 시황 항목을 찾는다."""
    r = " ".join(str(region or "").split())
    for key, val in regions.items():
        if str(key).startswith("_") or not isinstance(val, dict):
            continue
        if key == r:
            return key, val
    for key, val in regions.items():
        if str(key).startswith("_") or not isinstance(val, dict):
            continue
        for cand in _region_profile_match_candidates(key):
            if cand and cand == r:
                return key, val
    region_tail = r.split()[-1] if r.split() else ""
    if region_tail and region_tail not in AMBIGUOUS_SIGUNGU_NAMES:
        for key, val in regions.items():
            if str(key).startswith("_") or not isinstance(val, dict):
                continue
            if str(key).split()[-1] == region_tail:
                return key, val
    return "", {}


def _region_key_matches(key: str, region: str) -> bool:
    """조사 시황 키가 조사지역과 같은 시군구를 가리키는지 판정."""
    r = " ".join(str(region or "").split())
    k = str(key)
    if k == r:
        return True
    for cand in _region_profile_match_candidates(k):
        if cand and cand == r:
            return True
    tail = r.split()[-1] if r.split() else ""
    if tail and tail not in AMBIGUOUS_SIGUNGU_NAMES and k.split()[-1:] == [tail]:
        return True
    return False


def _collect_weekly_rollup(raw: Dict[str, Any], region: str, want_key: str, is_sido_scope: bool) -> List[Dict[str, Any]]:
    """월간 보고서 기준월에 속한 주간 조사 시황을 주차 순으로 모은다.

    월간 조사 시황이 아직 작성되지 않았을 때, 같은 달에 조사된 주간 시황을 월중 흐름
    근거로 활용하기 위한 취합이다. 주간 시황은 아파트 매매·전세만 다루므로
    월세·연립다세대·단독주택 서술의 근거로는 쓰지 않는다.
    """
    if not want_key:
        return []
    month_key = want_key[:7]
    regions, _used = _market_field_region_map(raw, "housing_weekly")
    if not regions:
        return []
    sido_aliases = set()
    if is_sido_scope:
        r = " ".join(str(region or "").split())
        for sido, names in SIDO_NAME_ALIASES.items():
            if r == sido or r in names:
                sido_aliases.add(sido)
                sido_aliases.update(names)
        sido_aliases.add(r)
    out = []
    for key, val in regions.items():
        if str(key).startswith("_") or not isinstance(val, dict):
            continue
        parts = str(key).split()
        if is_sido_scope:
            if len(parts) < 2 or parts[0] not in sido_aliases:
                continue
        elif not _region_key_matches(key, region):
            continue
        for label, rec in _entry_period_records(val):
            pk = normalize_period_key(label)
            if not pk.startswith(month_key + "-W"):
                continue
            out.append({
                "sigungu": parts[-1] if parts else str(key),
                "matched_region": str(key),
                "period_label": str(label),
                "week": pk,
                "headline": rec.get("headline") or {},
                "sale": rec.get("sale") or {},
                "jeonse": rec.get("jeonse") or {},
            })
    out.sort(key=lambda x: (x["sigungu"], x["week"]))
    return out


def _collect_sido_market_field(regions: Dict[str, Any], region: str, want_key: str = "") -> List[Dict[str, Any]]:
    """시도단위 보고서용. 선택 시도에 속한 모든 시군구 조사 시황을 모은다."""
    r = " ".join(str(region or "").split())
    if not r:
        return []
    aliases = set()
    for sido, names in SIDO_NAME_ALIASES.items():
        if r == sido or r in names:
            aliases.add(sido)
            aliases.update(names)
    aliases.add(r)
    out = []
    for key, val in regions.items():
        if str(key).startswith("_") or not isinstance(val, dict):
            continue
        parts = str(key).split()
        if len(parts) < 2:
            continue
        if parts[0] not in aliases:
            continue
        rec, _hit = _select_period_record(val, want_key)
        if not rec:
            continue
        item = dict(rec)
        item["sigungu"] = parts[-1]
        item["matched_region"] = key
        out.append(item)
    out.sort(key=lambda x: str(x.get("sigungu", "")))
    return out


def _apply_weekly_rollup(out: Dict[str, Any], raw: Dict[str, Any], region: str,
                         want_key: str, dataset_key: str, is_sido_scope: bool) -> None:
    """월간 조사 시황이 없을 때 같은 달 주간 조사 시황을 보완 근거로 붙인다.

    월간 보고서는 주간 조사와 무관하게 작성하는 것이 원칙이므로 기본값은 사용하지 않음이다.
    datasets.housing_monthly.use_weekly_rollup을 true로 둔 경우에만 취합한다.
    """
    if dataset_key != "housing_monthly":
        return
    datasets_node = raw.get("datasets") if isinstance(raw.get("datasets"), dict) else {}
    ds_node = datasets_node.get("housing_monthly") if isinstance(datasets_node.get("housing_monthly"), dict) else {}
    if not bool(ds_node.get("use_weekly_rollup")):
        return
    rollup = _collect_weekly_rollup(raw, region, want_key, is_sido_scope)
    if not rollup:
        return
    weeks = sorted({it["period_label"] for it in rollup if it.get("period_label")})
    names = sorted({it["sigungu"] for it in rollup})
    out["weekly_rollup"] = rollup
    out["weekly_rollup_note"] = (
        f"월간 조사 시황이 없어 같은 기준월에 조사된 주간 시황 {len(rollup)}건을 취합함"
        + (f" (시군구: {', '.join(names)}" if names else "")
        + (f" / 주차: {', '.join(weeks)})" if weeks else ")")
        + ". 주간 시황은 아파트 매매·전세만 다루므로 월세·연립다세대·단독주택 서술의 근거로는 사용하지 않음."
    )
    out["note"] = str(out.get("note", "")) + " " + out["weekly_rollup_note"]


def load_market_field_context(
    region: str,
    is_weekly: bool = False,
    is_officetel: bool = False,
    is_sido_scope: bool = False,
    period_label: str = "",
    base_month: str = "",
) -> Dict[str, Any]:
    """조사자가 작성한 현장 시황(매물·거래·매수매도 반응·읍면동 원인·지역 이슈)을 불러온다.

    data_pack의 실거래 통계와 달리 이 자료는 '원인과 시장 분위기'의 1차 근거로 사용한다.
    보고서 종류(주택 주간/주택 월간/오피스텔 월간)별로 dataset을 분리해 읽고,
    시도단위 보고서는 해당 시도 소속 시군구 조사 시황을 모두 모아 편차 서술에 쓴다.
    파일이 없거나 해당 항목이 없으면 available=False로 반환하고, 프롬프트는
    기존 data_pack·region_profile 기반 서술로 자동 대체된다.
    """
    dataset_key = market_field_dataset_key(is_weekly=is_weekly, is_officetel=is_officetel)
    asset_label = "오피스텔" if is_officetel else "주택"
    out = {
        "available": False,
        "source_file": "",
        "dataset": dataset_key,
        "dataset_label": MARKET_FIELD_DATASET_LABELS.get(dataset_key, dataset_key),
        "asset_label": asset_label,
        "scope": "sido" if is_sido_scope else "sigungu",
        "matched_region": "",
        "as_of": "",
        "period_label": "",
        "period_match": None,
        "report_period_label": str(period_label or base_month or ""),
        "surveyor": "",
        "price_survey_rate": "",
        "headline": {},
        "sale": {},
        "jeonse": {},
        "rent": {},
        "sigungu_contexts": [],
        "available_periods": [],
        "weekly_rollup": [],
        "weekly_rollup_note": "",
        "stale_regions": [],
        "allow_period_mismatch": False,
        "usage_rule": (
            f"조사자가 직접 확인한 {asset_label} 현장 시황입니다. 원인·시장 분위기·매물 흐름·매수매도 반응은 "
            "이 자료를 1차 근거로 사용하고, data_pack의 실거래 통계는 이를 수치로 보완·검증하는 보조지표로 사용합니다."
        ),
        "note": "market_field_context.json 미확인. 현장 시황 없이 data_pack·region_profile 기반으로만 서술.",
    }
    path = find_market_field_context_file()
    if not path:
        return out
    try:
        raw = read_json(path)
    except Exception as e:
        out["note"] = f"market_field_context.json 읽기 실패: {e}"
        return out
    if not isinstance(raw, dict):
        out["note"] = "market_field_context.json 형식이 올바르지 않습니다."
        return out
    out["source_file"] = path.name
    regions, used = _market_field_region_map(raw, dataset_key)
    if not regions:
        out["note"] = (
            f"market_field_context.json에 '{out['dataset_label']}'(datasets.{dataset_key}) 자료가 없습니다."
        )
        return out
    if used and used != dataset_key:
        out["dataset"] = used
    datasets_node = raw.get("datasets") if isinstance(raw.get("datasets"), dict) else {}
    ds_node = datasets_node.get(dataset_key) if isinstance(datasets_node.get(dataset_key), dict) else {}
    allow_mismatch = bool(ds_node.get("allow_period_mismatch") or raw.get("allow_period_mismatch"))
    out["allow_period_mismatch"] = allow_mismatch

    if is_sido_scope:
        want = str(period_label or "").strip()
        want_key = normalize_period_key(want)
        items = _collect_sido_market_field(regions, region, want_key)
        if not items:
            out["note"] = f"'{region}' 시도에 해당하는 시군구 조사 시황이 없습니다."
            _apply_weekly_rollup(out, raw, region, normalize_period_key(period_label), dataset_key, True)
            return out
        stale = []
        if want_key and not allow_mismatch:
            fresh = []
            for it in items:
                have = str(it.get("period_label") or "").strip()
                have_key = normalize_period_key(have)
                if have_key and have_key != want_key:
                    stale.append(f"{it.get('sigungu')}({have})")
                else:
                    fresh.append(it)
            items = fresh
        out["stale_regions"] = stale
        if not items:
            out["period_match"] = False
            out["note"] = (
                f"{out['dataset_label']}의 시군구 자료가 모두 보고서 기준({want})과 다른 기간이라 사용하지 않음"
                + (f" [{', '.join(stale)}]" if stale else "")
                + f". market_field_context.json의 datasets.{dataset_key}를 이번 기간 자료로 갱신하세요."
            )
            _apply_weekly_rollup(out, raw, region, want_key, dataset_key, True)
            return out
        if is_weekly:
            for it in items:
                it["rent"] = {}
        out["available"] = True
        out["matched_region"] = region
        out["sigungu_contexts"] = items
        labels = sorted({str(it.get("period_label") or it.get("as_of") or "") for it in items if it.get("period_label") or it.get("as_of")})
        out["period_label"] = " / ".join(labels)
        out["as_of"] = str(items[0].get("as_of") or "")
        if want_key and labels:
            out["period_match"] = all(normalize_period_key(l) == want_key for l in labels)
        out["note"] = (
            f"{out['dataset_label']} · 시도단위 취합 {len(items)}개 시군구"
            + (f" ({', '.join(str(it.get('sigungu')) for it in items)})" if items else "")
            + (f" / 기준기간 불일치로 제외된 시군구: {', '.join(stale)} — 해당 시군구는 조사 시황 없이 서술" if stale else "")
            + ("" if out["period_match"] is not False else " / 기준기간이 서로 다른 시군구가 섞여 있음 — 시점 확인 필요")
        )
        return out

    key, val = _match_market_field_region(regions, region)
    if not key:
        out["note"] = f"'{out['dataset_label']}'에 '{region}' 항목이 없습니다."
        _apply_weekly_rollup(out, raw, region, normalize_period_key(period_label), dataset_key, False)
        return out

    want = str(period_label or "").strip()
    want_key = normalize_period_key(want)
    val, _period_hit = _select_period_record(val, want_key)
    have = str(val.get("period_label") or "").strip()
    have_key = normalize_period_key(have)
    if want_key and have_key and want_key != have_key and not allow_mismatch:
        # 지난 기간 조사 시황이 이번 기간 시황으로 둔갑하지 않도록 사용하지 않는다.
        out["period_match"] = False
        out["matched_region"] = key
        out["period_label"] = have
        available_periods = [lb for lb, _r in _entry_period_records(regions.get(key, {})) if lb]
        out["available_periods"] = available_periods
        out["note"] = (
            f"'{key}' {out['dataset_label']}에 보고서 기준월({want}) 기록이 없음"
            + (f" / 보유 기간: {', '.join(available_periods)}" if available_periods else "")
            + f". market_field_context.json의 datasets.{dataset_key}.regions['{key}'].periods에 "
            f"'{want}' 기록을 추가하세요."
        )
        _apply_weekly_rollup(out, raw, region, want_key, dataset_key, False)
        return out

    out["available"] = True
    out["matched_region"] = key
    out["as_of"] = str(val.get("as_of") or "")
    out["period_label"] = str(val.get("period_label") or "")
    out["surveyor"] = str(val.get("surveyor") or "")
    out["price_survey_rate"] = str(val.get("price_survey_rate") or "")
    out["headline"] = val.get("headline") or {}
    for sec in MARKET_FIELD_SECTION_KEYS:
        node = val.get(sec)
        out[sec] = node if isinstance(node, dict) else {}
    if is_weekly:
        # 월세는 주간 공표 대상이 아니므로 사용하지 않는다.
        out["rent"] = {}
    if want_key and normalize_period_key(out["period_label"]):
        out["period_match"] = (want_key == normalize_period_key(out["period_label"]))
    out["note"] = (
        f"{out['dataset_label']} 기준: {out['period_label'] or out['as_of'] or '기준시점 미기재'}"
        + ("" if out["period_match"] is not False else f" / 보고서 기준({want})과 불일치 — 시점 확인 필요")
    )
    return out


def _latest_prev_from_rows(rows: List[Dict[str, Any]], month_key: str = "month") -> Tuple[Dict[str, Any], Dict[str, Any]]:
    rows = list(rows or [])
    if not rows:
        return {}, {}
    rows = sorted(rows, key=lambda r: str(r.get(month_key) or r.get("_month") or ""))
    return rows[-1], rows[-2] if len(rows) >= 2 else {}

def _delta_text(curr: Any, prev: Any, unit: str = "") -> str:
    try:
        c = float(curr or 0)
        p = float(prev or 0)
        d = c - p
        sign = "+" if d > 0 else ""
        if unit == "%":
            return f"{sign}{d:.1f}%p"
        return f"{sign}{int(round(d)):,}{unit}"
    except Exception:
        return "비교 불가"

def _top_names(rows: List[Dict[str, Any]], key_candidates: List[str], n: int = 3) -> List[str]:
    out = []
    for r in list(rows or [])[:n]:
        for k in key_candidates:
            if r.get(k):
                out.append(str(r.get(k)))
                break
    return out

def _to_float_safe(v: Any) -> Optional[float]:
    try:
        if v is None or pd.isna(v):
            return None
    except Exception:
        pass
    try:
        return float(v)
    except Exception:
        return None

def _rows_sorted(rows: List[Dict[str, Any]], key: str = "month") -> List[Dict[str, Any]]:
    return sorted(list(rows or []), key=lambda r: str(r.get(key) or r.get("_month") or ""))

def _pick_top_rows(rows: List[Dict[str, Any]], n: int = 5, count_keys: List[str] = None) -> List[Dict[str, Any]]:
    count_keys = count_keys or ["trade_count", "count", "거래건수", "value"]
    clean = []
    for r in list(rows or []):
        rr = dict(r)
        cnt = None
        for k in count_keys:
            if k in rr:
                cnt = _to_float_safe(rr.get(k))
                break
        rr["_sort_count"] = cnt if cnt is not None else 0
        clean.append(rr)
    clean.sort(key=lambda r: r.get("_sort_count", 0), reverse=True)
    for r in clean:
        r.pop("_sort_count", None)
    return clean[:n]

def _case_brief(row: Dict[str, Any], market_label: str = "") -> Dict[str, Any]:
    if not row:
        return {}
    keep = {}
    for k in ["rank", "building_name", "complex_name", "apt_name", "_group_name", "dong", "area", "pyeong", "approval_year", "floor", "prev_floor",
              "current_price_eok", "prev_price_eok", "current_value", "prev_value", "change_rate",
              "contract_date", "prev_contract_date", "month", "prev_month", "current_deposit", "current_monthly_rent"]:
        if k in row:
            keep[k] = row.get(k)
    keep["market"] = market_label
    return keep

def _top_case_briefs(rows: List[Dict[str, Any]], market_label: str, n: int = 3) -> List[Dict[str, Any]]:
    return [_case_brief(r, market_label) for r in list(rows or [])[:n]]

def _market_focus(label: str, obj: Dict[str, Any]) -> Dict[str, Any]:
    latest, prev = _latest_prev_from_rows(obj.get("monthly_rows", []))
    return {
        "market": label,
        "latest_month": latest.get("month") or latest.get("_month"),
        "trade_count": latest.get("trade_count"),
        "trade_count_mom_change": _delta_text(latest.get("trade_count"), prev.get("trade_count"), "건") if prev else "전월 비교자료 부족",
        "compare_count": latest.get("compare_count"),
        "new_count": latest.get("new_count"),
        "renew_count": latest.get("renew_count"),
        "up_ratio": latest.get("up_ratio"),
        "up_ratio_mom_change": _delta_text(latest.get("up_ratio"), prev.get("up_ratio"), "%") if prev else "전월 비교자료 부족",
        "flat_ratio": latest.get("flat_ratio"),
        "down_ratio": latest.get("down_ratio"),
        "down_ratio_mom_change": _delta_text(latest.get("down_ratio"), prev.get("down_ratio"), "%") if prev else "전월 비교자료 부족",
        "avg_change_rate": latest.get("avg_change_rate"),
        "top_dongs": _pick_top_rows(obj.get("dong_rows", []), 5),
        "top_up_cases": _top_case_briefs(obj.get("top_up_rows", []), label, 3),
        "top_down_cases": _top_case_briefs(obj.get("top_down_rows", []), label, 3),
        "top_value_cases": _top_case_briefs(obj.get("top_value_rows", []), label, 3),
    }

def build_monthly_change_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    """LLM이 매월 같은 문장을 반복하지 않도록 당월 변화 포인트를 구조화.
    v5.8.38: 매매·전세·월세 변화 요약을 풍부하게 확장.
    """
    apt = data.get("apt_trade", {})
    ar = data.get("apt_rent", {})
    pop = data.get("population", {})
    mig = data.get("migration", {})
    rates = data.get("rates", {})
    unsold = data.get("unsold", {})
    supply = data.get("supply", {})

    focus = [
        _market_focus("아파트 매매", apt),
        _market_focus("아파트 전세", ar.get("jeonse", {})),
        _market_focus("아파트 월세", ar.get("monthly_rent", {})),
    ]

    return {
        "purpose": "LLM 서술 반복 방지용 당월 변화 포인트. 지역 특성은 이 변화가 확인된 경우에만 연결.",
        "population_change": {
            "population_mom": pop.get("mom_change"),
            "households_mom": pop.get("households_mom"),
            "net_migration": mig.get("net_migration"),
            "main_outflow_age": mig.get("main_outflow_age"),
            "latest_population": pop.get("total"),
            "latest_households": pop.get("households"),
        },
        "financial_supply_change": {
            "mortgage_rate": rates.get("mortgage_rate"),
            "jeonse_rate": rates.get("jeonse_rate"),
            "unsold_units": unsold.get("unsold_units"),
            "after_completion_unsold": unsold.get("after_completion_unsold"),
            "supply_latest": supply.get("latest"),
            "supply_scope_note": "공급 파이프라인은 시군구가 아니라 시도 단위 공급통계 기준으로 해석",
        },
        "market_monthly_focus": focus,
        "major_dongs": {
            "apt_sale_top_dongs": _top_names(apt.get("dong_rows", []), ["dong", "name"], 5),
            "apt_jeonse_top_dongs": _top_names(ar.get("jeonse", {}).get("dong_rows", []), ["dong", "name"], 5),
            "apt_rent_top_dongs": _top_names(ar.get("monthly_rent", {}).get("dong_rows", []), ["dong", "name"], 5),
        },
        "major_complex_examples": {
            "apt_sale_up": _top_names(apt.get("top_up_rows", []), ["building_name", "complex_name", "apt_name", "_group_name"], 5),
            "apt_sale_down": _top_names(apt.get("top_down_rows", []), ["building_name", "complex_name", "apt_name", "_group_name"], 5),
            "apt_jeonse_up": _top_names(ar.get("jeonse", {}).get("top_up_rows", []), ["building_name", "complex_name", "apt_name", "_group_name"], 5),
            "apt_jeonse_down": _top_names(ar.get("jeonse", {}).get("top_down_rows", []), ["building_name", "complex_name", "apt_name", "_group_name"], 5),
            "apt_rent_up": _top_names(ar.get("monthly_rent", {}).get("top_up_rows", []), ["building_name", "complex_name", "apt_name", "_group_name"], 5),
            "apt_rent_down": _top_names(ar.get("monthly_rent", {}).get("top_down_rows", []), ["building_name", "complex_name", "apt_name", "_group_name"], 5),
        }
    }


def _top_rows_for_llm(rows: List[Dict[str, Any]], n: int = 10) -> List[Dict[str, Any]]:
    """LLM 입력용 상위 행 압축. 표 전체가 너무 길어지지 않도록 상위 n개만 전달."""
    if not rows:
        return []
    try:
        return _pick_top_rows(rows, n)
    except Exception:
        return list(rows)[:n]


def _detail_market_stats(obj: Dict[str, Any], n: int = 10) -> Dict[str, Any]:
    """당월 시장별 읍면동·면적·가격·보증금·월세·건축연령 세부 통계 요약."""
    if not obj:
        return {}
    return {
        "latest_month": obj.get("latest_month") or obj.get("analysis_base_month"),
        "analysis_base_month": obj.get("analysis_base_month"),
        "trade_count": obj.get("trade_count"),
        "new_count": obj.get("new_count"),
        "renew_count": obj.get("renew_count"),
        "compare_count": obj.get("compare_count"),
        "up_ratio": obj.get("up_ratio"),
        "flat_ratio": obj.get("flat_ratio"),
        "down_ratio": obj.get("down_ratio"),
        "avg_change_rate": obj.get("avg_change_rate"),
        "signal": obj.get("signal"),
        "value_basis": obj.get("value_basis"),
        "applied_conversion_rate": obj.get("applied_conversion_rate"),
        "detail_count_basis": obj.get("detail_count_basis"),
        "dong_rows_top10": _top_rows_for_llm(obj.get("dong_rows", []), n),
        "area_band_rows": _top_rows_for_llm(obj.get("area_band_rows", []), n),
        "price_band_rows": _top_rows_for_llm(obj.get("price_band_rows", []), n),
        "deposit_band_rows": _top_rows_for_llm(obj.get("deposit_band_rows", []), n),
        "monthly_band_rows": _top_rows_for_llm(obj.get("monthly_band_rows", []), n),
        "converted_rent_band_rows": _top_rows_for_llm(obj.get("converted_rent_band_rows", []), n),
        "building_age_rows": _top_rows_for_llm(obj.get("building_age_rows", []), n),
        "top_complex_rows": _top_rows_for_llm(obj.get("top_complex_rows", []), n),
    }


def build_detail_statistics_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    """LLM 분석용 당월 세부 통계 묶음.
    읍면동, 연령대, 면적구간, 가격구간, 보증금구간, 월세/환산월세구간, 건축연령별 자료를
    프롬프트 앞쪽에 별도로 제공하여 LLM이 세부표 수치를 활용하게 한다.
    """
    apt_trade = data.get("apt_trade", {})
    apt_rent = data.get("apt_rent", {})
    non_apt_sale = data.get("non_apt_sale", {})
    non_apt_rent = data.get("non_apt_rent", {})
    age_pop = data.get("age_population", {})
    migration = data.get("migration", {})

    return {
        "purpose": "LLM 서술 보강용 당월 세부 통계. 읍면동·연령대·면적구간·가격/보증금/월세구간·건축연령별 통계를 실제 수치 기반으로 서술할 때 사용.",
        "count_rule": "거래건수는 각 세부표의 count 기준을 따르며, 전세·월세는 거래건수·신규계약 건수·갱신재계약 건수를 분리 표시. 상승·보합·하락 비율은 신규계약 중 직전거래 비교가능건 기준.",
        "age_population_rows": _top_rows_for_llm(age_pop.get("age_rows", []), 20),
        "age_population_monthly_rows": _top_rows_for_llm(age_pop.get("monthly_rows", []), 13),
        "migration_monthly_rows": _top_rows_for_llm(migration.get("monthly_rows", []), 13),
        "apt_sale": _detail_market_stats(apt_trade, 12),
        "apt_jeonse": _detail_market_stats(apt_rent.get("jeonse", {}), 12),
        "apt_rent": _detail_market_stats(apt_rent.get("monthly_rent", {}), 12),
        "non_apt_sale": _detail_market_stats(non_apt_sale, 12),
        "non_apt_jeonse": _detail_market_stats(non_apt_rent.get("jeonse", {}), 12),
        "non_apt_rent": _detail_market_stats(non_apt_rent.get("monthly_rent", {}), 12),
    }


def build_dong_focus_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    """읍면동별 서술 재료. LLM이 지역특성과 연결할 수 있는 상위 동 중심 자료."""
    apt = data.get("apt_trade", {})
    ar = data.get("apt_rent", {})
    nas = data.get("non_apt_sale", {})
    nar = data.get("non_apt_rent", {})
    return {
        "purpose": "읍면동별 거래 집중도와 지역특성 연결용. 상위 동만 서술에 사용하고 근거 없는 지역명은 쓰지 말 것.",
        "apt_sale_top5": _pick_top_rows(apt.get("dong_rows", []), 5),
        "apt_jeonse_top5": _pick_top_rows(ar.get("jeonse", {}).get("dong_rows", []), 5),
        "apt_rent_top5": _pick_top_rows(ar.get("monthly_rent", {}).get("dong_rows", []), 5),
        "non_apt_sale_top5": _pick_top_rows(nas.get("dong_rows", []), 5),
        "non_apt_jeonse_top5": _pick_top_rows(nar.get("jeonse", {}).get("dong_rows", []), 5),
        "non_apt_rent_top5": _pick_top_rows(nar.get("monthly_rent", {}).get("dong_rows", []), 5),
    }

def build_case_focus_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    """대표 거래사례 요약. LLM 종합서술에 직접 쓸 수 있도록 매매·전세·월세 사례를 분리."""
    apt = data.get("apt_trade", {})
    ar = data.get("apt_rent", {})
    return {
        "purpose": "종합서술·아파트 문단에 직접 활용 가능한 대표 사례. 전체 시장 단정 근거가 아니라 보조 사례로 사용.",
        "apt_sale": {
            "representative_up_cases": _top_case_briefs(apt.get("top_up_rows", []), "아파트 매매 상승", 5),
            "representative_down_cases": _top_case_briefs(apt.get("top_down_rows", []), "아파트 매매 하락", 5),
            "top_value_cases": _top_case_briefs(apt.get("top_value_rows", []), "아파트 매매 거래금액 상위", 5),
        },
        "apt_jeonse": {
            "representative_up_cases": _top_case_briefs(ar.get("jeonse", {}).get("top_up_rows", []), "아파트 전세 상승", 5),
            "representative_down_cases": _top_case_briefs(ar.get("jeonse", {}).get("top_down_rows", []), "아파트 전세 하락", 5),
            "top_value_cases": _top_case_briefs(ar.get("jeonse", {}).get("top_value_rows", []), "아파트 전세 보증금 상위", 5),
        },
        "apt_rent": {
            "representative_up_cases": _top_case_briefs(ar.get("monthly_rent", {}).get("top_up_rows", []), "아파트 월세 상승", 5),
            "representative_down_cases": _top_case_briefs(ar.get("monthly_rent", {}).get("top_down_rows", []), "아파트 월세 하락", 5),
            "top_value_cases": _top_case_briefs(ar.get("monthly_rent", {}).get("top_value_rows", []), "아파트 월세 환산월세 상위", 5),
        }
    }

def _type_count_summary(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    out = {"latest_month": None, "type_counts": {}, "dominant_type": None, "interpretation_limit": ""}
    rows = _rows_sorted(rows)
    if not rows:
        out["interpretation_limit"] = "거래건수 원자료 부족"
        return out
    latest_month = rows[-1].get("month") or rows[-1].get("_month")
    latest = [r for r in rows if (r.get("month") or r.get("_month")) == latest_month]
    out["latest_month"] = latest_month
    for r in latest:
        typ = str(r.get("type") or r.get("_source_type") or r.get("구분") or "")
        cnt = _to_float_safe(r.get("trade_count") or r.get("count") or r.get("value"))
        if typ:
            out["type_counts"][typ] = int(cnt or 0)
    if out["type_counts"]:
        out["dominant_type"] = max(out["type_counts"], key=out["type_counts"].get)
        total = sum(out["type_counts"].values())
        if total < 10:
            out["interpretation_limit"] = "거래건수가 10건 미만이면 가격 방향성 단정 금지"
        elif total < 30:
            out["interpretation_limit"] = "거래건수가 제한적이므로 유형별 거래 집중도 중심으로 제한 해석"
        else:
            out["interpretation_limit"] = "유형별 거래 집중도와 주요 읍면동을 함께 해석"
    return out

def build_non_apt_interpretation_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    nas = data.get("non_apt_sale", {})
    nar = data.get("non_apt_rent", {})
    return {
        "purpose": "비아파트는 거래건수 편차가 크므로 유형별 건수·주요 읍면동·해석 제한 문구를 함께 사용.",
        "sale": {
            "trade_count": nas.get("trade_count"),
            "monthly_type_summary": _type_count_summary(nas.get("monthly_by_type_rows", [])),
            "top_dongs": _pick_top_rows(nas.get("dong_rows", []), 5),
            "interpretation_rule": "연립·다세대·단독·다가구별 거래건수 차이가 크면 가격 방향성을 단정하지 말고 보조지표로 해석",
        },
        "jeonse": {
            "trade_count": nar.get("jeonse", {}).get("trade_count"),
            "new_count": nar.get("jeonse", {}).get("new_count"),
            "renew_count": nar.get("jeonse", {}).get("renew_count"),
            "monthly_type_summary": _type_count_summary(nar.get("jeonse", {}).get("monthly_by_type_rows", [])),
            "top_dongs": _pick_top_rows(nar.get("jeonse", {}).get("dong_rows", []), 5),
            "interpretation_rule": "전세는 보증금 수준·신규/갱신 구분·보증 안정성을 함께 해석",
        },
        "rent": {
            "trade_count": nar.get("monthly_rent", {}).get("trade_count"),
            "new_count": nar.get("monthly_rent", {}).get("new_count"),
            "renew_count": nar.get("monthly_rent", {}).get("renew_count"),
            "monthly_type_summary": _type_count_summary(nar.get("monthly_rent", {}).get("monthly_by_type_rows", [])),
            "top_dongs": _pick_top_rows(nar.get("monthly_rent", {}).get("dong_rows", []), 5),
            "interpretation_rule": "월세는 환산월세·보증금/월세 조합·소형 임대수요 중심으로 해석",
        }
    }

def build_supply_pressure_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    supply = data.get("supply", {})
    rows = _rows_sorted(supply.get("table_rows", []))
    latest, prev = _latest_prev_from_rows(rows)
    def delta(k):
        return _delta_text(latest.get(k), prev.get(k), "호") if latest and prev else "전년 비교자료 부족"
    return {
        "purpose": "공급 파이프라인 해석용. 시군구 보고서라도 공급은 시도 단위 인허가·착공·준공으로 해석.",
        "supply_scope": "시도 단위 공급통계",
        "latest_year": latest.get("month") if latest else None,
        "latest_values": latest,
        "year_over_year_change": {
            "permit": delta("permit"),
            "start": delta("start"),
            "completion": delta("completion"),
        },
        "by_type_latest": {
            "permit": _type_count_summary(supply.get("permit_by_type_rows", [])),
            "start": _type_count_summary(supply.get("start_by_type_rows", [])),
            "completion": _type_count_summary(supply.get("completion_by_type_rows", [])),
        },
        "interpretation_rule": "인허가 증가는 중장기 공급 예정 물량, 착공은 실제 공급 전환, 준공은 입주·전세 수급 압력으로 구분하여 해석",
    }


MOVE_IN_DB_FILENAME = "apt_move_in_schedule.sqlite3"


def find_move_in_database() -> Optional[Path]:
    """실행파일 옆 외부 DB를 우선하고 PyInstaller 번들 DB를 보조로 사용한다."""
    candidates = [BASE_DIR / "local_db" / MOVE_IN_DB_FILENAME]
    try:
        candidates.append(Path(__file__).resolve().parent / "local_db" / MOVE_IN_DB_FILENAME)
    except Exception:
        pass
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "local_db" / MOVE_IN_DB_FILENAME)
    candidates.extend([
        Path.cwd() / "local_db" / MOVE_IN_DB_FILENAME,
        BASE_DIR.parent / "local_db" / MOVE_IN_DB_FILENAME,
    ])
    seen = set()
    for candidate in candidates:
        key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        if candidate.exists():
            return candidate
    return None


def _report_sigungu(region: str) -> str:
    tokens = re.findall(r"[가-힣]+(?:특별시|광역시|시|군|구)", str(region or ""))
    candidates = [t for t in tokens if not t.endswith(("특별시", "광역시"))]
    return candidates[-1] if candidates else ""


def load_move_in_supply(region: str, base_month: str, report_mode: str) -> Dict[str, Any]:
    """로컬 DB에서 기준연도 시군구별 아파트 입주예정물량을 조회한다."""
    target_year = int(str(base_month or "0000")[:4]) if str(base_month or "")[:4].isdigit() else None
    sigungu = _report_sigungu(region)
    empty = {
        "available": False,
        "has_region_data": False,
        "target_year": target_year,
        "scope_sigungu": sigungu,
        "scope_label": sigungu or "전남광주통합특별시 전체",
        "region_summary_rows": [],
        "monthly_rows": [],
        "project_rows": [],
        "scope_total_households": 0,
        "all_region_total_households": 0,
        "scope_project_count": 0,
        "peak_month": "",
        "peak_households": 0,
        "source": "",
        "database_path": "",
        "signal": "입주예정물량 DB 없음",
    }
    if str(report_mode or "") not in {"weekly", "monthly"} or target_year is None:
        empty["signal"] = "주간·월간 주택동향만 적용"
        return empty
    db_path = find_move_in_database()
    if not db_path:
        return empty
    try:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            metadata = {r["key"]: r["value"] for r in conn.execute("SELECT key,value FROM metadata")}
            raw = [dict(r) for r in conn.execute(
                """SELECT sigungu, project_name, address, planned_month, total_households,
                          supplied_households, household_basis, note
                   FROM move_in_projects
                   WHERE planned_year=?
                   ORDER BY planned_month, sigungu, project_name""",
                (target_year,),
            )]
    except Exception as exc:
        empty["database_path"] = str(db_path)
        empty["signal"] = f"입주예정물량 DB 조회 오류: {exc}"
        return empty

    counted = [r for r in raw if int(r.get("total_households") or 0) > 0]
    all_total = sum(int(r.get("total_households") or 0) for r in counted)
    summary_rows = []
    for name in sorted({str(r.get("sigungu") or "") for r in counted}):
        items = [r for r in counted if r.get("sigungu") == name]
        total = sum(int(r.get("total_households") or 0) for r in items)
        by_month = Counter()
        for item in items:
            by_month[str(item.get("planned_month") or "")] += int(item.get("total_households") or 0)
        peak_month, peak_value = max(by_month.items(), key=lambda x: x[1]) if by_month else ("", 0)
        summary_rows.append({
            "시군구": name,
            "입주예정 단지수": len(items),
            "입주예정물량": total,
            "전체 대비 비중": round(total / all_total * 100, 1) if all_total else 0.0,
            "최대 집중월": peak_month,
            "집중월 물량": peak_value,
        })
    summary_rows.sort(key=lambda r: r["입주예정물량"], reverse=True)

    scoped = [r for r in counted if not sigungu or r.get("sigungu") == sigungu]
    monthly_rows = []
    for month in range(1, 13):
        ym = f"{target_year}-{month:02d}"
        items = [r for r in scoped if r.get("planned_month") == ym]
        monthly_rows.append({
            "입주예정월": ym,
            "입주예정 단지수": len(items),
            "입주예정물량": sum(int(r.get("total_households") or 0) for r in items),
        })
    peak_row = max(monthly_rows, key=lambda r: r["입주예정물량"]) if monthly_rows else {}
    project_rows = [{
        "단지명": r.get("project_name"),
        "주소": r.get("address"),
        "세대수": int(r.get("total_households") or 0),
        "입주예정일자": (
            f"{str(r.get('planned_month'))[:4]}년 {int(str(r.get('planned_month'))[5:7])}월 (월 기준)"
            if re.fullmatch(r"\d{4}-\d{2}", str(r.get("planned_month") or "")) else "-"
        ),
        "입주예정월": r.get("planned_month"),
        "시군구": r.get("sigungu"),
        "입주예정물량": int(r.get("total_households") or 0),
        "일반공급 규모": int(r.get("supplied_households") or 0),
        "산정기준": r.get("household_basis"),
        "비고": r.get("note") or "-",
    } for r in scoped]
    scope_total = sum(r["세대수"] for r in project_rows)
    return {
        **empty,
        "available": True,
        "has_region_data": bool(scoped),
        "region_summary_rows": summary_rows,
        "monthly_rows": monthly_rows,
        "project_rows": project_rows,
        "scope_total_households": scope_total,
        "all_region_total_households": all_total,
        "scope_project_count": len(project_rows),
        "peak_month": peak_row.get("입주예정월", "") if peak_row.get("입주예정물량", 0) else "",
        "peak_households": peak_row.get("입주예정물량", 0),
        "source": metadata.get("source_workbook", "로컬 입주예정물량 DB"),
        "source_imported_at": metadata.get("imported_at", ""),
        "household_rule": metadata.get("household_rule", "전체 세대수 우선"),
        "location_rule": metadata.get("location_rule", "주소 기준 시군구"),
        "database_path": str(db_path),
        "signal": f"{target_year}년 {sigungu or '전체'} {scope_total:,}호 입주 예정",
    }


SEMICONDUCTOR_CLUSTER_DONG_LABELS = {
    "송정동": {"label": "반도체 메가클러스터 핵심 영향권", "level": "core"},
    "도산동": {"label": "반도체 메가클러스터 핵심 영향권", "level": "core"},
    "신촌동": {"label": "반도체 메가클러스터 핵심 영향권", "level": "core"},
    "소촌동": {"label": "반도체 메가클러스터 핵심 영향권", "level": "core"},
    "우산동": {"label": "반도체 메가클러스터 인접 영향권", "level": "adjacent"},
    "선암동": {"label": "반도체 메가클러스터 연계 영향권", "level": "linked"},
    "운수동": {"label": "반도체 메가클러스터 연계 영향권", "level": "linked"},
}


def build_semiconductor_cluster_context(region: str, base_month: str, report_mode: str) -> Dict[str, Any]:
    """광산구 주택동향 보고서용 정책 사실과 분석상 영향권을 분리해 제공한다."""
    active = (
        "광산구" in str(region or "")
        and str(base_month or "") >= "2026-07"
        and str(report_mode or "") in {"weekly", "monthly", "officetel"}
    )
    return {
        "available": active,
        "title": "호남 반도체 첨단국가산단 및 토지거래허가구역",
        "announcement_date": "2026-07-09",
        "effective_period": "2026-07-14 ~ 2028-07-13",
        "official_fact": (
            "국토교통부 공고 제2026-940호에 따라 광주 군공항의 호남 반도체 첨단국가산단 "
            "사업 예정지와 인근 지역이 토지거래허가구역으로 지정되어 2026년 7월 14일부터 효력이 발생함."
        ),
        "scope_note": (
            "토지거래허가구역은 법정동·리 경계 기준의 광범위한 투기방지 구역이며, "
            "그 자체가 개별 읍면동 주택가격의 상승 또는 직접 수혜를 의미하지 않음."
        ),
        "influence_method": (
            "아래 영향권은 공식 사업구역 명칭이 아니라 군공항 예정지와의 인접성, 송정 생활권 및 교통 연계를 "
            "기준으로 한 보고서용 분석 분류임. 실거래 변화가 확인될 때만 시장 흐름과 제한적으로 연결함."
        ),
        "core_dongs": ["송정동", "도산동", "신촌동", "소촌동"],
        "adjacent_dongs": ["우산동"],
        "linked_dongs": ["선암동", "운수동"],
        "dong_labels": SEMICONDUCTOR_CLUSTER_DONG_LABELS if active else {},
        "source_name": "국토교통부 참고·설명자료 및 공고 제2026-940호",
        "source_url": "https://www.molit.go.kr/USR/NEWS/m_72/dtl.jsp?id=95092209",
        "interpretation_rule": (
            "정책 발표 전후 거래량·가격 변화를 데이터로 확인하지 않은 상태에서 개발 기대감이 가격을 올렸다고 단정하지 않음. "
            "토지거래허가구역 규제효과와 산업단지 조성 기대효과를 서로 구분함."
        ),
    }


def apply_semiconductor_cluster_labels(value: Any, context: Dict[str, Any]) -> Any:
    """표 원자료 구조는 유지하고 읍면동 행에만 HTML 표시용 메타데이터를 붙인다."""
    if not context.get("available"):
        return value
    labels = context.get("dong_labels") or {}
    if isinstance(value, dict):
        for key, item in list(value.items()):
            value[key] = apply_semiconductor_cluster_labels(item, context)
        dong = str(value.get("dong") or value.get("읍면동") or "").strip()
        if dong in labels:
            value["_cluster_influence_label"] = labels[dong]["label"]
            value["_cluster_influence_level"] = labels[dong]["level"]
        return value
    if isinstance(value, list):
        for item in value:
            apply_semiconductor_cluster_labels(item, context)
    return value


def build_data_pack(
    input_files: List[Path],
    region: str,
    base_month: str,
    report_version: str,
    input_source: str = "auto",
    designated_sale_files: Optional[List[Path]] = None,
    designated_rent_files: Optional[List[Path]] = None,
) -> Dict[str, Any]:
    progress("원자료 파일 자동 분류 중", 1, 14)
    input_source = str(input_source or "auto").strip().lower()
    selected_files, excluded_files = filter_files_by_input_source(input_files, input_source)
    classified = classify_files(
        selected_files,
        input_source,
        designated_sale_files=designated_sale_files,
        designated_rent_files=designated_rent_files,
    )
    report_mode = infer_report_mode(classified, report_version)
    is_weekly = report_mode == "weekly"
    is_officetel = report_mode == "officetel"
    is_sido_scope = is_sido_level_region(region)
    primary_asset_label = "오피스텔" if is_officetel else "아파트"

    progress("인구·세대수 계산 중", 2, 14)
    population = calc_population(classified["population"], region, base_month)
    progress("연령대별 주민등록인구 계산 중", 3, 15)
    age_population = calc_age_population(classified["population"], region, base_month)
    progress("연령·이동자수 계산 중", 4, 15)
    migration = calc_migration(classified["migration"], region, base_month)
    progress("대출금리 계산 중", 4, 14)
    rates = calc_rates(classified["rates"], region, base_month)
    progress("미분양 계산 중", 5, 14)
    unsold = calc_unsold(classified["unsold"], region, base_month)
    progress("가격지수·주간 공표통계 계산 중", 6, 14)
    if is_weekly:
        # 주간업무도 공표통계는 월간 아파트 매매·전세가격지수를 사용한다.
        # 폴더에 과거 주간 변동률 파일이 남아 있어도 보고서 계산에는 사용하지 않는다.
        # 주간 아파트 변동률은 R-ONE 주간 지수에서 받은 파일이 있을 때만 별도 주간 지표로 쓴다.
        weekly_sale_change = calc_weekly_price_change(classified["weekly_apt_sale_change"], region, base_month, "매매")
        weekly_jeonse_change = calc_weekly_price_change(classified["weekly_apt_jeonse_change"], region, base_month, "전세")
        monthly_sale_index = calc_price_index(classified["price_index_sale"], region, base_month, "매매")
        monthly_jeonse_index = calc_price_index(classified["price_index_jeonse"], region, base_month, "전세")
        # 실거래 13개월 표의 '시군구 공표통계'는 월간 시군구 아파트
        # 가격지수 전월비만 사용한다. 주간 공표값은 별도 주간 지표로 유지하며
        # 월간 미공표월(예: 2026-06/07)을 대체하지 않는다.
        monthly_sale_index["public_stat_label"] = "시군구 공표통계"
        monthly_jeonse_index["public_stat_label"] = "시군구 공표통계"
        price_index = {
            "sale": monthly_sale_index,
            "jeonse": monthly_jeonse_index,
            "rent": {"available": False, "table_rows": [], "signal": "주간보고서 월세 공표통계 제외"},
        }
        weekly_period_info = build_weekly_report_period_info(
            base_month,
            price_index.get("sale", {}).get("latest_week"),
            price_index.get("jeonse", {}).get("latest_week"),
        )
    else:
        price_index = {
            "sale": calc_price_index(classified["price_index_sale"], region, base_month, "매매"),
            "jeonse": calc_price_index(classified["price_index_jeonse"], region, base_month, "전세"),
            "rent": calc_price_index(classified["price_index_rent"], region, base_month, "월세"),
        }
        public_stat_label = "시도공표통계" if (is_officetel or is_sido_scope or is_sido_public_stat_region(region)) else "시군구 공표통계"
        for part in price_index.values():
            if isinstance(part, dict):
                part["public_stat_label"] = public_stat_label
        weekly_period_info = {}
    progress("공급 파이프라인 계산 중", 7, 14)
    supply = calc_supply(classified["supply"], region, base_month)
    move_in_supply = load_move_in_supply(region, base_month, report_mode)
    progress(f"{primary_asset_label} 매매 계산 중", 8, 14)
    apt_trade = calc_trade_common(classified["apt_trade"], region, base_month, f"{primary_asset_label} 매매")
    unsold_news: Dict[str, Any] = {"available": False, "items": []}
    try:
        if float((unsold or {}).get("unsold_units") or 0) > 0:
            progress("미분양 관련 보도 검색 중", 8, 14)
            unsold_news = fetch_unsold_news(region, base_month, [r.get("dong") for r in apt_trade.get("dong_rows", [])[:15]])
    except Exception:
        pass
    progress("외지인 매수 시계열 계산 중", 9, 15)
    external_buyer_sale = calc_external_buyer_sale(
        classified.get("buyer_sale_upload", []), region, base_month
    )
    progress("전월세전환율 산출 중", 9, 15)

    # v5.8.52: KOSIS 공표 전월세전환율 엑셀(시도+물건유형+근접월) 을 우선 사용.
    # 자료가 없을 때만 실거래 기반 calc_conversion_rate() 로 fallback.
    public_conv_files = classified.get("conversion_rate_public", [])
    conversion_rate_apt = calc_public_conversion_rate_from_kosis(public_conv_files, region, base_month, "아파트")
    if not conversion_rate_apt.get("available"):
        conversion_rate_apt = calc_conversion_rate(classified["apt_rent"], region, base_month, "아파트")
    conversion_rate_apt = apply_confirmed_conversion_rate_if_missing(conversion_rate_apt, region, base_month, "아파트")

    conversion_rate_rh = {}
    if not is_officetel:
        conversion_rate_rh = calc_public_conversion_rate_from_kosis(public_conv_files, region, base_month, "연립다세대")
        if not conversion_rate_rh.get("available"):
            conversion_rate_rh = calc_conversion_rate(classified["non_apt_rent"], region, base_month, "비아파트")
        conversion_rate_rh = apply_confirmed_conversion_rate_if_missing(conversion_rate_rh, region, base_month, "연립다세대")
    conversion_rate_sh = conversion_rate_rh

    if is_officetel:
        # 오피스텔은 오피스텔 공표 전환율을 적용.
        conversion_rate_offi = calc_public_conversion_rate_from_kosis(public_conv_files, region, base_month, "오피스텔")
        if not conversion_rate_offi.get("available"):
            conversion_rate_offi = calc_conversion_rate(classified.get("officetel_rent", []), region, base_month, "오피스텔")
        conversion_rate_offi = apply_confirmed_conversion_rate_if_missing(conversion_rate_offi, region, base_month, "오피스텔")
    else:
        conversion_rate_offi = {}

    conversion_rate = conversion_rate_offi if is_officetel else conversion_rate_apt

    # v5.8.52: v19 광주 5.57% 하드코딩 fallback 제거. 공표 전월세전환율 KOSIS 엑셀
    # (calc_public_conversion_rate_from_kosis) 매칭으로 동적 산출.
    progress(f"{primary_asset_label} 전월세 계산 중", 10, 15)
    apt_rent = calc_rent_common(classified["apt_rent"], region, base_month, primary_asset_label, conversion_rate)

    if is_weekly:
        weekly_reference_date = weekly_period_info.get("weekly_reference_date", "")
        if weekly_reference_date:
            apt_trade["weekly_case_reference_date"] = weekly_reference_date
            if isinstance(apt_rent.get("jeonse"), dict):
                apt_rent["jeonse"]["weekly_case_reference_date"] = weekly_reference_date

    # 13개월 실거래 시계열 표 보강:
    # 평균 변동률 옆에 KOSIS 주택동향지수 기반 아파트 변동률을 '공표통계'로 표시
    apt_trade["monthly_rows"] = attach_public_index_change_rate(apt_trade.get("monthly_rows", []), price_index.get("sale", {}))
    if isinstance(apt_rent.get("jeonse"), dict):
        apt_rent["jeonse"]["monthly_rows"] = attach_public_index_change_rate(apt_rent["jeonse"].get("monthly_rows", []), price_index.get("jeonse", {}))
    if isinstance(apt_rent.get("monthly_rent"), dict):
        apt_rent["monthly_rent"]["monthly_rows"] = attach_public_index_change_rate(apt_rent["monthly_rent"].get("monthly_rows", []), price_index.get("rent", {}))

    if is_weekly or is_officetel:
        progress("비아파트 실거래 계산 제외", 11, 15)
        non_apt_sale = calc_trade_common([], region, base_month, "비아파트 매매")
        non_apt_rent = calc_rent_common([], region, base_month, "비아파트", conversion_rate_rh or conversion_rate)
    else:
        progress("비아파트 매매 계산 중", 11, 15)
        non_apt_sale = calc_trade_common(classified["non_apt_sale"], region, base_month, "비아파트 매매")
        progress("비아파트 전월세 계산 중", 12, 15)
        non_apt_rent = calc_rent_common(classified["non_apt_rent"], region, base_month, "비아파트", conversion_rate_rh or conversion_rate)
    progress("분양권전매 계산 중", 13, 15)
    presale = calc_trade_common([] if is_officetel else classified["presale"], region, base_month, "분양권전매")

    policy_context = {
        "semiconductor_cluster": build_semiconductor_cluster_context(region, base_month, report_mode)
    }
    apply_semiconductor_cluster_labels(apt_trade, policy_context["semiconductor_cluster"])
    apply_semiconductor_cluster_labels(apt_rent, policy_context["semiconductor_cluster"])

    warnings = []
    if not population.get("available") or population.get("total") is None:
        warnings.append("인구·세대수 수치가 산출되지 않았습니다.")
    if not age_population.get("available"):
        warnings.append("연령대별 주민등록인구 구조가 산출되지 않았습니다.")
    if apt_trade.get("latest_month") and apt_trade.get("latest_month") != base_month:
        warnings.append(f"{primary_asset_label} 매매 최신월({apt_trade.get('latest_month')})이 기준월({base_month})과 다릅니다.")
    if classified.get("buyer_sale_upload") and not external_buyer_sale.get("available"):
        warnings.append("업로드 매매 원장에서 매수자 시군구 기반 외지인 집계를 산출하지 못했습니다. 주거용 물건구분·계약일자·매수자 시군구 열을 확인하세요.")
    if is_weekly:
        if not classified["price_index_sale"] or not classified["price_index_jeonse"]:
            warnings.append("주간업무 공표통계용 월간 아파트 매매/전세 가격지수 파일 일부가 누락되었습니다.")
    else:
        if not classified["price_index_sale"] or not classified["price_index_jeonse"] or not classified["price_index_rent"]:
            warnings.append("가격지수 3종 중 일부 파일이 누락되었습니다.")
    if not classified["supply"]:
        warnings.append("공급 4종 파일이 누락되었거나 분류되지 않았습니다.")
    if not is_officetel and report_mode in {"weekly", "monthly"} and not move_in_supply.get("available"):
        warnings.append("로컬 아파트 입주예정물량 DB를 찾지 못했거나 조회할 수 없습니다.")
    if is_weekly:
        pass
    elif is_officetel:
        pass
    elif not classified["non_apt_sale"] or not classified["non_apt_rent"]:
        warnings.append("비아파트 매매/전월세 파일 일부가 누락되었습니다.")
    if not conversion_rate.get("available"):
        warnings.append("전월세전환율 산출 불가: 아파트 신규 전세·신규 월세 동일 주택형 매칭을 확인하세요.")

    progress("data_pack 구조 생성 중", 14, 15)
    return {
        "meta": {
            "region": region,
            "region_short": get_region_short_name(region),
            "supply_region": get_supply_sido_display_name(region),
            "base_month": base_month,
            "report_version": report_version,
            "report_mode": report_mode,
            "is_weekly": is_weekly,
            "is_officetel": is_officetel,
            "report_scope": "sido" if is_sido_scope else "sigungu",
            "scope_label": "시도단위" if is_sido_scope else "시군구단위",
            "primary_asset_label": primary_asset_label,
            "input_source": input_source,
            "input_source_label": INPUT_SOURCE_LABELS.get(input_source, INPUT_SOURCE_LABELS["auto"]),
            "input_source_excluded_files": [p.name for p in excluded_files],
            "report_title": f"주간동향보고서({weekly_period_info.get('weekly_title_suffix')})" if is_weekly and weekly_period_info.get("weekly_title_suffix") else ("주간동향보고서" if is_weekly else (("오피스텔동향보고서" if not is_sido_scope else "오피스텔동향보고서(시도단위)") if is_officetel else ("월간동향보고서" if not is_sido_scope else "월간동향보고서(시도단위)"))),
            "weekly_reference_date": weekly_period_info.get("weekly_reference_date", ""),
            "weekly_week_of_month": weekly_period_info.get("weekly_week_of_month"),
            "weekly_period_label": weekly_period_info.get("weekly_period_label", ""),
            "app_version": APP_VERSION,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "trade_basis": "계약일 기준",
            "preliminary_note": "기준월 실거래는 신고기한 미도래로 향후 보정될 수 있음",
        },
        "input_files": {k: [p.name for p in v] for k, v in classified.items()},
        "summary_cards": {
            "population": population.get("total"),
            "population_mom": population.get("mom_change"),
            "households": population.get("households"),
            "households_mom": population.get("households_mom"),
            "apt_trade_count": apt_trade.get("trade_count"),
            "apt_trade_mom": apt_trade.get("mom_change"),
            "apt_jeonse_count": apt_rent.get("jeonse", {}).get("trade_count"),
            "apt_monthly_rent_count": apt_rent.get("monthly_rent", {}).get("trade_count"),
            "mortgage_rate": rates.get("mortgage_rate"),
            "jeonse_rate": rates.get("jeonse_rate"),
            "supply_permit": supply.get("latest", {}).get("permit"),
            "supply_start": supply.get("latest", {}).get("start"),
            "supply_completion": supply.get("latest", {}).get("completion"),
            "supply_year": supply.get("latest", {}).get("year"),
            "unsold_units": unsold.get("unsold_units"),
            "after_completion_unsold": unsold.get("after_completion_unsold"),
        },
        "population": population,
        "age_population": age_population,
        "migration": migration,
        "rates": rates,
        "unsold": unsold,
        "conversion_rate": conversion_rate,
        "conversion_rate_apt": conversion_rate_apt, "conversion_rate_rh": conversion_rate_rh, "conversion_rate_sh": conversion_rate_sh, "conversion_rate_offi": conversion_rate_offi,
        "price_index": price_index,
        "weekly_apt_sale_change": weekly_sale_change if is_weekly else {},
        "weekly_apt_jeonse_change": weekly_jeonse_change if is_weekly else {},
        "supply": supply,
        "move_in_supply": move_in_supply,
        "unsold_news": unsold_news,
        "apt_trade": apt_trade,
        "external_buyer_sale": external_buyer_sale,
        "apt_rent": apt_rent,
        "officetel_trade": apt_trade if is_officetel else {},
        "officetel_rent": apt_rent if is_officetel else {},
        "non_apt_sale": non_apt_sale,
        "non_apt_rent": non_apt_rent,
        "presale": presale,
        "policy_context": policy_context,
        "region_profile": get_region_profile(region),
        "market_field_context": load_market_field_context(
            region,
            is_weekly=is_weekly,
            is_officetel=is_officetel,
            is_sido_scope=is_sido_scope,
            period_label=weekly_period_info.get("weekly_period_label", "") if is_weekly else base_month,
            base_month=base_month,
        ),
        "monthly_change_summary": build_monthly_change_summary({
            "population": population,
            "migration": migration,
            "rates": rates,
            "unsold": unsold,
            "supply": supply,
            "apt_trade": apt_trade,
            "apt_rent": apt_rent,
        }),
        "weekly_change_summary": build_monthly_change_summary({
            "population": population,
            "migration": migration,
            "rates": rates,
            "unsold": unsold,
            "supply": supply,
            "apt_trade": apt_trade,
            "apt_rent": apt_rent,
        }) if is_weekly else {},
        "dong_focus_summary": build_dong_focus_summary({
            "apt_trade": apt_trade,
            "apt_rent": apt_rent,
            "non_apt_sale": non_apt_sale,
            "non_apt_rent": non_apt_rent,
        }),
        "detail_statistics_summary": build_detail_statistics_summary({
            "age_population": age_population,
            "migration": migration,
            "apt_trade": apt_trade,
            "apt_rent": apt_rent,
            "non_apt_sale": non_apt_sale,
            "non_apt_rent": non_apt_rent,
        }),
        "case_focus_summary": build_case_focus_summary({
            "apt_trade": apt_trade,
            "apt_rent": apt_rent,
        }),
        "non_apt_interpretation_summary": build_non_apt_interpretation_summary({
            "non_apt_sale": non_apt_sale,
            "non_apt_rent": non_apt_rent,
        }),
        "supply_pressure_summary": build_supply_pressure_summary({
            "supply": supply,
        }),
        "data_quality_warnings": warnings,
    }


# =========================================================
# 7. 해설 및 프롬프트
# =========================================================

def base_commentary(data: Dict[str, Any]) -> Dict[str, str]:
    m = data["meta"]; region=m["region"]; bm=m["base_month"]
    primary_label = m.get("primary_asset_label") or ("오피스텔" if str(m.get("report_mode", "")) == "officetel" else "아파트")
    pop=data["population"]; mig=data["migration"]; rates=data["rates"]; unsold=data["unsold"]
    apt=data["apt_trade"]; ar=data["apt_rent"]; nas=data["non_apt_sale"]; nar=data["non_apt_rent"]; pre=data["presale"]
    ext=data.get("external_buyer_sale", {}) or {}
    pi=data["price_index"]; supply=data["supply"]; conv=data.get("conversion_rate", {})
    move_in_supply = data.get("move_in_supply", {}) or {}
    region_profile = data.get("region_profile", {}) or {}
    cluster_context = (data.get("policy_context", {}) or {}).get("semiconductor_cluster", {}) or {}

    def s_count(d, key="trade_count"): return format_num(d.get(key))
    def pct(d, key): return format_rate(d.get(key))
    def _top_dongs(section: Dict[str, Any], n: int = 3) -> str:
        rows = list(section.get("dong_rows") or [])[:n]
        vals = []
        for r in rows:
            dong = str(r.get("dong") or r.get("name") or "").strip()
            cnt = r.get("count")
            if dong:
                vals.append(f"{dong} {format_num(cnt)}건")
        return ", ".join(vals) if vals else "거래 상위 읍면동 자료 제한"

    def _top_cases(section: Dict[str, Any], keys: List[str]) -> str:
        vals = []
        for k in keys:
            rows = section.get(k) or []
            if rows:
                r = rows[0]
                loc = r.get("location") or r.get("dong") or "-"
                typ = r.get("property_type") or ""
                val = r.get("transaction_price_eok") or r.get("deposit_eok") or r.get("converted_rent_manwon")
                vals.append(f"{typ} {loc} {val}")
        return ", ".join(vals) if vals else "주요 거래사례 제한"

    def _profile_text() -> str:
        summary = str(region_profile.get("region_summary") or "")
        dps = region_profile.get("dong_profiles") or {}
        if dps:
            samples = "; ".join([f"{k}: {v}" for k, v in list(dps.items())[:4]])
            return f"{summary} 세부 생활권은 {samples}"
        return summary or f"{region}은 거래 상위 읍면동, 주요 거래사례, 유형별 거래건수를 중심으로 지역 맥락을 제한적으로 해석함."

    def _non_apt_context(kind: str, section: Dict[str, Any]) -> str:
        if kind == "매매":
            metric = f"비아파트 매매 {format_num(section.get('trade_count'))}건, 상위 읍면동 {_top_dongs(section)}"
            cases = _top_cases(section, ["rowhouse_major_trade_rows", "multiplex_major_trade_rows", "multifamily_major_trade_rows", "detached_major_trade_rows"])
            market = "매매는 실거주 대체수요와 토지성 가치, 노후주택 정비 기대, 원도심·신도심 간 가격 접근성 차이가 함께 반영되는 영역"
        elif kind == "전세":
            metric = f"비아파트 전세 {format_num(section.get('trade_count'))}건, 상위 읍면동 {_top_dongs(section)}"
            cases = _top_cases(section, ["rowhouse_major_trade_rows", "multiplex_major_trade_rows", "multifamily_major_trade_rows", "detached_major_trade_rows"])
            market = "전세는 보증금 부담, 통근·통학 접근성, 상권과 생활편의, 임차인의 보증 안정성 선호가 거래 성립을 좌우하는 영역"
        else:
            metric = f"비아파트 월세 {format_num(section.get('trade_count'))}건, 상위 읍면동 {_top_dongs(section)}"
            cases = _top_cases(section, ["rowhouse_major_trade_rows", "multiplex_major_trade_rows", "multifamily_major_trade_rows", "detached_major_trade_rows"])
            market = "월세는 소형 임대수요, 직주근접, 상권 접근성, 전세 보증금 부담에 따른 월세 전환, 원도심 저가 임차수요가 함께 작용하는 영역"
        text = (
            f"{bm} {region} {kind} 비아파트 지역특성은 당월 원자료에서 확인되는 {metric}을 우선 근거로 해석함. "
            f"{_profile_text()} "
            f"따라서 교통 접근성은 단순 역세권 여부보다 당월 거래가 집중된 읍면동에서 직장·학교·상권·행정서비스로 이동하기 쉬운 생활권인지 여부와 연결해 볼 필요가 있음. "
            f"배후수요는 신도심·택지지구의 가족형 실수요, 원도심의 저가 실거주·소규모 임대수요, 상권 주변의 단기·소형 임차수요로 나누어 점검하는 것이 적절함. "
            f"학군과 생활편의가 강한 지역은 연립·다세대의 대체 주거수요를 지지할 수 있으나, 단독·다가구는 대지·노후도·임대수익성·정비 가능성에 따라 개별성이 크게 나타남. "
            f"주요 사례는 {cases}로 확인되며, {market}으로 판단됨. 거래건수가 제한적인 유형은 가격 방향성을 단정하기보다 소재지, 사용승인년도, 층, 면적, 보증금·월세 또는 거래금액 조건을 함께 비교하는 보조 판단이 필요함."
        )
        return text

    monthly_core = (
        f"{bm} {region} 주택시장은 인구 {format_num(pop.get('total'))}명, 세대수 {format_num(pop.get('households'))}세대, "
        f"주택담보대출금리 {format_float(rates.get('mortgage_rate'),2)}%, 미분양 {format_num(unsold.get('unsold_units'))}호를 배경으로 "
        f"{primary_label} 매매 {s_count(apt)}건, 하락 비율 {pct(apt,'down_ratio')}가 확인되는 국면으로 분석됨. "
        f"가격지수와 공급 지표를 함께 보면 시장은 단일 방향보다 생활권·유형별 차별화가 큰 선별 거래 국면으로 판단됨."
    )
    population_c = (
        f"{bm} {region} 인구는 {format_num(pop.get('total'))}명, 세대수는 {format_num(pop.get('households'))}세대로 산출됨. "
        f"전월 대비 인구 증감은 {format_num(pop.get('mom_change'))}명, 세대수 증감은 {format_num(pop.get('households_mom'))}세대로 확인됨. "
        f"인구 총량과 세대수 흐름을 함께 보면 주택수요 기반은 인구 규모보다 가구 단위 수요의 유지 여부가 중요한 것으로 판단됨."
    )
    age_c = (
        f"연령별 이동자수는 최신월 {mig.get('latest_month') or '-'} 기준 순이동 {format_num(mig.get('net_migration'))}명으로 산출됨. "
        f"주요 순유출 연령대는 {mig.get('main_outflow_age') or '원자료 확인 필요'}로 확인되며, 청년층 이동은 임대수요와 신규 매수수요 판단의 보조지표로 활용할 필요가 있음. "
        f"연령별 이동자료는 총인구·세대수와 함께 생활권별 수요 강도를 해석하는 데 활용되는 것으로 판단됨."
    )
    rate_c = (
        f"예금은행 대출금리는 최신월 {rates.get('latest_month') or '-'} 기준 주택담보대출 {format_float(rates.get('mortgage_rate'),2)}%, "
        f"전세자금대출 {format_float(rates.get('jeonse_rate'),2)}%, 중소기업대출 {format_float(rates.get('sme_rate'),2)}%, "
        f"기업일반자금대출 {format_float(rates.get('business_rate'),2)}%로 산출됨. "
        f"금리 수준은 매수 전환, 전세 보증금 부담, 월세 전환 압력에 동시에 영향을 주는 요인으로 분석됨."
    )
    unsold_c = (
        f"미분양은 최신월 {unsold.get('latest_month') or '-'} 기준 {format_num(unsold.get('unsold_units'))}호, "
        f"준공후 미분양은 {format_num(unsold.get('after_completion_unsold'))}호로 산출됨. "
        f"미분양 및 준공후 미분양은 신규 분양 기대심리와 기존 매매시장 매수심리의 보조지표로 판단됨."
    )
    supply_c = (
        f"공급 파이프라인은 인허가·착공·준공 실적을 기준으로 점검되며, 최신 산출값은 {json.dumps(supply.get('latest', {}), ensure_ascii=False)}로 확인됨. "
        f"공급 실적은 당월 가격을 직접 결정하기보다 향후 입주 부담, 전세 수급, 분양시장 심리를 판단하는 선행·보조지표로 활용됨."
    )
    if move_in_supply.get("available") and is_jeonnam_gwangju_region((data.get("meta") or {}).get("region", "")):
        if move_in_supply.get("has_region_data"):
            supply_c += (
                f" 로컬 입주예정물량 DB 기준 {move_in_supply.get('target_year')}년 "
                f"{move_in_supply.get('scope_label')} 아파트 입주예정물량은 "
                f"{format_num(move_in_supply.get('scope_total_households'))}호, "
                f"{format_num(move_in_supply.get('scope_project_count'))}개 단지로 확인됨. "
                f"가장 물량이 집중되는 시기는 {move_in_supply.get('peak_month') or '-'} "
                f"{format_num(move_in_supply.get('peak_households'))}호이며, 입주시점 변경 가능성을 고려해 전세 수급과 기존주택 거래 흐름을 함께 점검할 필요가 있음."
            )
        else:
            supply_c += (
                f" 로컬 입주예정물량 DB에는 {move_in_supply.get('target_year')}년 "
                f"{move_in_supply.get('scope_label')} 입주예정 단지가 확인되지 않음."
            )
    pi_c = (
        f"가격지수는 매매 {format_float(pi['sale'].get('latest_value'),2)}, 전세 {format_float(pi['jeonse'].get('latest_value'),2)}, 월세 {format_float(pi['rent'].get('latest_value'),2)}로 산출됨. "
        f"실거래 직전거래 비교는 개별 거래의 가격 변동성을 보여주고, 가격지수는 공식 조사 기준의 완만한 추세를 보완적으로 보여주는 지표로 해석됨."
    )
    apt_c = (
        f"{primary_label} 매매는 {s_count(apt)}건으로 산출되었고, 직전거래 비교 가능 {format_num(apt.get('compare_count'))}건 중 상승 {pct(apt,'up_ratio')}, "
        f"보합 {pct(apt,'flat_ratio')}, 하락 {pct(apt,'down_ratio')}로 나타남. 평균 변동률은 {format_float(apt.get('avg_change_rate'))}%로 확인되어 {apt.get('signal')} 흐름으로 분석됨. "
        f"읍면동·가격구간·평형·연식별 세부표와 TOP 거래사례를 함께 보면 단지별 조건에 따른 선별 거래가 뚜렷한 것으로 판단됨."
    )
    if ext.get("available"):
        top_buyer = (ext.get("current_region_rows") or [{}])[0].get("매수자 시군구", "상위 지역 확인 필요")
        external_c = (
            f"{bm} 주거용 매매 원장 기준 전체 {format_num(ext.get('current_total_count'))}건 중 외지인 매수는 "
            f"{format_num(ext.get('current_external_count'))}건, 비율은 {format_float(ext.get('current_external_share'),1)}%로 산출됨. "
            f"당월 외지인 매수의 최다 유입 시군구는 {top_buyer}로 확인되며, 최근 6개월 누적 "
            f"{format_num(ext.get('six_month_external_count'))}건과 13개월 시계열을 함께 보면 외부 수요의 지속성과 일시적 집중 여부를 구분해 판단할 필요가 있음."
        )
    else:
        external_c = "매수자 시군구가 포함된 주거용 매매 원장이 없어 외지인 매수 흐름은 산출되지 않음."
    aj_c = (
        f"{primary_label} 전세는 신규계약 기준 {format_num(ar.get('jeonse',{}).get('trade_count'))}건으로 산출되었고, "
        f"상승 {format_rate(ar.get('jeonse',{}).get('up_ratio'))}, 하락 {format_rate(ar.get('jeonse',{}).get('down_ratio'))}로 확인됨. "
        f"전월세 구분은 {ar.get('rent_type_basis') or '원자료 거래구분 열'}, 물건종류는 {ar.get('property_type_basis') or '주용도 우선·건물유형 보완'} 기준으로 확인함. "
        f"전세자금대출금리와 보증금 부담을 고려할 때 전세시장은 단지별 보증금 방어력 차이가 큰 것으로 판단됨."
    )
    rent_c = (
        f"{primary_label} 월세는 환산월세 기준 {format_num(ar.get('monthly_rent',{}).get('trade_count'))}건으로 산출되었고, "
        f"상승 {format_rate(ar.get('monthly_rent',{}).get('up_ratio'))}, 하락 {format_rate(ar.get('monthly_rent',{}).get('down_ratio'))}로 확인됨. "
        f"전월세 구분은 {ar.get('rent_type_basis') or '원자료 거래구분 열'}, 물건종류는 {ar.get('property_type_basis') or '주용도 우선·건물유형 보완'} 기준으로 확인함. "
        f"월세시장은 보증금과 월세 조합에 따라 실제 주거비 부담이 달라지므로 환산월세 기준 해석이 필요한 것으로 판단됨."
    )
    non_sale_c = (
        f"비아파트 매매는 {format_num(nas.get('trade_count'))}건으로 산출되었고, 주요 읍면동과 가격구간별 거래 분포를 통해 구도심 저가 수요와 토지성 가치가 함께 반영되는 것으로 분석됨. "
        f"연립·다세대와 단독·다가구는 물건 특성과 거래규모가 달라 아파트 매매와 분리하여 해석하는 것이 적절함."
    )
    non_j_c = (
        f"비아파트 전세는 {format_num(nar.get('jeonse',{}).get('trade_count'))}건으로 산출되었고, 보증 안정성과 관리상태, 구도심 접근성이 거래 성립에 영향을 미치는 것으로 판단됨. "
        f"거래규모가 제한적인 경우 가격 방향성은 단정하지 않고 주요 읍면동과 보증금 수준 중심으로 해석하는 것이 적절함."
    )
    non_r_c = (
        f"비아파트 월세는 {format_num(nar.get('monthly_rent',{}).get('trade_count'))}건으로 산출되었고, 1~2인 가구와 고령층, 저가 임대수요의 영향을 받는 것으로 분석됨. "
        f"환산월세와 월세 수준을 함께 고려해 실제 주거비 부담 관점에서 해석하는 것이 필요함."
    )
    if not pre.get("trade_count"):
        pre_c = (
            f"분양권전매는 최근 13개월({pre.get('latest_month') or bm} 기준) 거래가 확인되지 않음. "
            f"표·차트의 최근 13개월 구간은 모두 거래 0건으로, 분양권전매 시장이 사실상 휴면 상태인 것으로 판단됨. "
            f"분양권전매는 거래가 없어 신규 공급 기대를 판단할 직접 근거로 활용하기 어려우며, 향후 거래 재개 시 보조지표로 재검토하는 것이 적절함."
        )
    else:
        pre_c = (
            f"분양권전매는 {format_num(pre.get('trade_count'))}건으로 산출되었고, 거래단지와 가격 수준은 신축 기대수요와 기존 매매시장 심리를 보완적으로 보여주는 지표로 판단됨. "
            f"거래규모가 제한적인 경우 분양권전매는 전체 시장 방향을 단정하는 지표가 아니라 신규 공급 기대를 확인하는 보조지표로 활용하는 것이 적절함."
        )
    non_apt_sale_regional_context = _non_apt_context("매매", nas)
    non_apt_jeonse_regional_context = _non_apt_context("전세", nar.get("jeonse", {}))
    non_apt_rent_regional_context = _non_apt_context("월세", nar.get("monthly_rent", {}))
    def _with_regional_context(context: str, summary: str) -> str:
        context = str(context or "").strip()
        summary = str(summary or "").strip()
        if not context:
            return summary
        return f"{context} {summary}"
    top_c = (
        "주요 거래사례는 거래단지 TOP10, 상승거래 TOP10, 하락거래 TOP10, 거래금액 상위 TOP10을 함께 확인해 해석해야 함. "
        "상승·하락률이 큰 거래는 층, 면적, 물건 상태, 직전거래 조건 차이가 반영될 수 있어 전체 시장 방향을 단정하는 근거로 사용하지 않는 것이 적절함."
    )
    comprehensive = (
        f"종합적으로 {bm} {region} 주택시장은 인구·금리·미분양·공급 여건과 실거래 가격 변동이 복합적으로 작용하는 국면으로 분석됨. "
        f"아파트 매매는 {apt.get('signal')} 흐름을 보이며, 전세와 월세는 신규계약 및 환산월세 기준으로 별도 점검이 필요한 것으로 판단됨. "
        f"비아파트와 분양권전매는 거래규모가 제한적일 수 있으므로 아파트 중심지표를 보완하는 자료로 활용하는 것이 적절함."
    )

    if cluster_context.get("available"):
        policy_fact = (
            "국토교통부가 2026년 7월 9일 호남 반도체 첨단국가산단 사업 예정지 일원의 "
            "토지거래허가구역을 공고했으며 7월 14일부터 효력이 발생함. "
            "송정·도산·신촌·소촌동은 핵심 영향권, 우산동은 인접 영향권, 선암·운수동은 연계 영향권으로 관찰하되, "
            "이는 보고서용 분석 분류이므로 실거래 변화 확인 없이 가격 영향을 단정하지 않는 것이 적절함."
        )
        monthly_core = f"{monthly_core} {policy_fact}"
        apt_c = f"{apt_c} {policy_fact}"
        aj_c = (
            f"{aj_c} 해당 정책은 토지시장 규제와 장기 산업배후수요 기대가 함께 작용할 수 있으나, "
            "전세 거래량·보증금의 실제 변화가 확인될 때만 제한적으로 연결해 해석할 필요가 있음."
        )
        comprehensive = f"{comprehensive} {policy_fact}"

    d = {
        "monthly_core_diagnosis": monthly_core, "population_commentary": population_c, "age_migration_commentary": age_c,
        "rate_commentary": rate_c, "unsold_commentary": unsold_c, "unsold_news_summary": unsold_news_summary_text(data.get("unsold_news") or {}), "supply_commentary": supply_c, "price_index_commentary": pi_c,
        "conversion_rate_commentary": f"월세 환산 기준은 {conv.get('period', '-')} 기간의 {primary_label} 신규 전세·신규 월세 동일 주택형 매칭 {format_num(conv.get('matched_count'))}건을 기준으로 산출되었으며, 적용 전환율은 {format_float(conv.get('applied_rate'),2)}%로 확인됨. 산식은 전환율(%) = 월세×12÷(전세보증금-월세보증금)×100, 환산월세 = 월세+(보증금×적용전환율÷100÷12)를 적용함.",
        "apt_trade_commentary": apt_c, "external_buyer_commentary": external_c,
        "supply_type_commentary": supply_type_commentary_text(supply, data.get("meta", {})),
        "apt_price_band_commentary": band_commentary_text(apt.get("price_band_rows"), f"{primary_label} 매매 거래금액", "고가"),
        "apt_age_band_commentary": band_commentary_text(apt.get("building_age_rows"), f"{primary_label} 매매 연식", "노후"),
        "apt_area_band_commentary": band_commentary_text(apt.get("area_band_rows"), f"{primary_label} 매매 평형", "대형"), "apt_jeonse_commentary": aj_c, "apt_rent_commentary": rent_c,
        "non_apt_sale_commentary": non_sale_c, "non_apt_jeonse_commentary": non_j_c, "non_apt_rent_commentary": non_r_c,
        "presale_commentary": pre_c, "top_case_commentary": top_c, "comprehensive_summary": comprehensive,
        "sale_market_summary": f"{apt_c} {non_sale_c} {pre_c}",
        "apt_sale_summary": apt_c,
        "non_apt_sale_regional_context": non_apt_sale_regional_context,
        "rowhouse_multiplex_sale_summary": _with_regional_context(non_apt_sale_regional_context, "연립·다세대 매매는 아파트 대체 저가 수요, 구도심 접근성, 관리상태, 주차여건에 따라 선별적으로 발생하는 유형으로 판단됨."),
        "detached_multifamily_sale_summary": _with_regional_context(non_apt_sale_regional_context, "단독·다가구 매매는 토지성 가치, 임대수익, 노후주택 정비 가능성, 금리 부담이 함께 반영되는 유형으로 분석됨."),
        "jeonse_market_summary": f"{aj_c} {non_j_c}",
        "apt_jeonse_summary": aj_c,
        "non_apt_jeonse_regional_context": non_apt_jeonse_regional_context,
        "rowhouse_multiplex_jeonse_summary": _with_regional_context(non_apt_jeonse_regional_context, "연립·다세대 전세는 구도심 저가 전세수요와 보증 안정성, 관리상태가 거래 성립에 영향을 주는 유형으로 판단됨."),
        "detached_multifamily_jeonse_summary": _with_regional_context(non_apt_jeonse_regional_context, "단독·다가구 전세는 1~2인 가구, 고령층 정주수요, 보증 안정성이 중요한 유형으로 분석됨."),
        "rent_market_summary": f"{rent_c} {non_r_c}",
        "apt_rent_summary": rent_c,
        "non_apt_rent_regional_context": non_apt_rent_regional_context,
        "rowhouse_multiplex_rent_summary": _with_regional_context(non_apt_rent_regional_context, "연립·다세대 월세는 저가 임대수요와 생활권 접근성, 관리상태에 따라 선별되는 유형으로 판단됨."),
        "detached_multifamily_rent_summary": _with_regional_context(non_apt_rent_regional_context, "단독·다가구 월세는 구도심 생활권, 1~2인 가구, 고령층, 실속형 임대수요가 반영되는 유형으로 분석됨.")
    }
    def _input_names_text() -> str:
        names = []
        for vals in (data.get("input_files") or {}).values():
            names.extend([str(x) for x in vals or []])
        return " ".join(names)

    def _is_officetel_context() -> bool:
        if str(m.get("report_mode", "")) == "officetel" or m.get("is_officetel"):
            return True
        text = _input_names_text()
        return "오피스텔" in text or "offi" in text.lower() or "officetel" in text.lower()

    def _flow_word(value: Any, up_word: str = "상승", down_word: str = "하락", flat_word: str = "보합") -> str:
        v = _to_float_safe(value)
        if v is None:
            return "확인 필요"
        if v > 0:
            return up_word
        if v < 0:
            return down_word
        return flat_word

    def _series_flow(rows: List[Dict[str, Any]], value_keys: List[str], label: str) -> str:
        rows = _rows_sorted(rows or [])
        vals = []
        for r in rows:
            v = None
            for key in value_keys:
                v = _to_float_safe(r.get(key))
                if v is not None:
                    break
            if v is not None:
                vals.append(v)
        if len(vals) < 2:
            return f"{label} 시계열은 원자료가 제한적이어서 최신 공표값과 보조지표를 함께 볼 필요가 있음"
        first, last = vals[0], vals[-1]
        positives = sum(1 for v in vals if v > 0)
        negatives = sum(1 for v in vals if v < 0)
        if last > first and positives >= negatives:
            return f"{label} 시계열은 초반보다 높은 수준으로 이동했고 최근 구간에서도 상승 신호가 우세함"
        if last < first and negatives > positives:
            return f"{label} 시계열은 초반보다 낮아졌고 최근 구간에서도 하락 또는 조정 신호가 더 많이 관찰됨"
        return f"{label} 시계열은 한 방향으로 강하게 쏠리기보다 상승과 하락이 교차하는 혼조 흐름에 가까움"

    def _top_dong_words(section: Dict[str, Any], fallback: str = "거래 집중 지역") -> str:
        rows = section.get("dong_rows") or []
        names = []
        for r in list(rows)[:3]:
            name = str(r.get("dong") or r.get("name") or "").strip()
            if name:
                names.append(name)
        return ", ".join(names) if names else fallback

    def _profile_context_for_summary() -> str:
        summary = str(region_profile.get("region_summary") or "").strip()
        if summary:
            return summary
        if _is_officetel_context():
            return f"{region} 오피스텔 시장은 업무·상업 접근성, 소형 임대수요, 역세권·생활편의 입지, 전세 보증금 부담에 따른 월세 전환 수요가 함께 작용하는 시장임."
        return f"{region} 주택시장은 생활권별 선호, 교통 접근성, 상권과 학군, 신도심·원도심의 주거 여건 차이가 거래 흐름에 함께 반영되는 시장임."

    def _demand_context() -> str:
        pop_flow = _flow_word(pop.get("mom_change"), "증가", "감소", "보합")
        hh_flow = _flow_word(pop.get("households_mom"), "증가", "감소", "보합")
        mig_flow = _flow_word(mig.get("net_migration"), "순유입", "순유출", "유출입 균형")
        return (
            f"인구는 {pop_flow}, 세대수는 {hh_flow} 흐름이고 이동 지표는 {mig_flow} 방향을 보임. "
            "따라서 수요 기반은 총량보다 세대 분화, 통근·통학 편의, 생활권 선호가 어디에 집중되는지를 함께 봐야 함."
        )

    def _market_narrative(market: str, section: Dict[str, Any], index_part: Dict[str, Any], extra: str = "") -> str:
        top_dongs = _top_dong_words(section)
        up = _to_float_safe(section.get("up_ratio"))
        down = _to_float_safe(section.get("down_ratio"))
        avg = _to_float_safe(section.get("avg_change_rate"))
        if up is not None and down is not None:
            if up > down:
                direction = "상승 거래가 하락 거래보다 우세해 가격 방어력이 상대적으로 확인되는 흐름"
            elif down > up:
                direction = "하락 거래 비중이 더 높아 매수·임차 수요가 가격을 선별하는 조정 흐름"
            else:
                direction = "상승과 하락이 균형을 이루며 뚜렷한 방향성보다 관망세가 강한 흐름"
        elif avg is not None:
            direction = f"평균 변동률 기준으로 {_flow_word(avg)} 흐름"
        else:
            direction = "거래 비교자료가 제한되어 방향성보다 거래가 형성된 생활권을 중심으로 봐야 하는 흐름"
        index_flow = _series_flow(index_part.get("table_rows", []), ["public_index_change_rate", "value", "latest_value"], f"{market} 공표통계")
        text = (
            f"{bm} {region} {market}시장은 {top_dongs} 등 거래가 확인된 생활권을 중심으로 해석할 필요가 있음. "
            f"{_profile_context_for_summary()} {_demand_context()} "
            f"최근 13개월 거래 흐름에서는 {direction}으로 읽히며, {index_flow}. "
            f"교통 접근성과 상권·학군·업무 배후수요가 강한 권역은 실수요와 임차수요를 지지하지만, 구축 비중이 높거나 거래가 적은 권역은 개별 물건 조건에 따라 상승·하락이 엇갈릴 수 있음."
        )
        return (text + (" " + extra if extra else "")).strip()

    sale_extra = "" if _is_officetel_context() else "분양권전매와 비아파트 매매는 전체 방향을 단정하는 지표라기보다 신규 공급 기대와 원도심·저가 대체수요를 확인하는 보조자료로 보는 것이 적절함."
    sale_context = _market_narrative(
        "매매",
        apt,
        pi.get("sale", {}),
        sale_extra
    )
    jeonse_context = _market_narrative(
        "전세",
        ar.get("jeonse", {}),
        pi.get("jeonse", {}),
        "전세는 보증금 부담, 입주 물량, 통근·통학 접근성, 학군 선호가 함께 작용하므로 신규계약과 갱신계약을 구분해 해석해야 함."
    )
    rent_context = _market_narrative(
        "월세",
        ar.get("monthly_rent", {}),
        pi.get("rent", {}),
        "월세는 전세 보증금 부담과 금리 수준, 소형 임대수요, 직주근접·상권 접근성이 함께 반영되므로 환산월세 기준 흐름을 중심으로 보는 것이 필요함."
    )
    comprehensive_context = (
        f"종합적으로 {bm} {region} {'오피스텔' if _is_officetel_context() else '주택'}시장은 숫자 하나로 방향을 단정하기보다 생활권별 수요와 공표통계, 실거래 비교 흐름을 함께 읽어야 함. "
        f"{_profile_context_for_summary()} {_demand_context()} "
        f"매매는 {_flow_word(apt.get('avg_change_rate'))} 압력이, 전세는 {_flow_word(ar.get('jeonse', {}).get('avg_change_rate'))} 흐름이, 월세는 {_flow_word(ar.get('monthly_rent', {}).get('avg_change_rate'))} 신호가 각각 확인되며 시장별 온도 차이가 있음. "
        "거래가 집중된 읍면동과 주요 단지 사례는 지역 내 선호 생활권을 보여주는 단서이고, 공표통계 시계열은 개별 거래의 편차를 보정해 전체 흐름을 확인하는 기준으로 활용하는 것이 적절함. "
        "향후에는 인구·세대 이동의 변화, 미분양과 공급 부담, 금리와 보증금 부담이 매수 전환 또는 임대수요 이동으로 이어지는지를 조사상 유의사항으로 점검할 필요가 있음."
    )
    d.update({
        "sale_market_summary": sale_context,
        "apt_sale_summary": sale_context,
        "jeonse_market_summary": jeonse_context,
        "apt_jeonse_summary": jeonse_context,
        "rent_market_summary": rent_context,
        "apt_rent_summary": rent_context,
        "comprehensive_summary": comprehensive_context,
    })
    if _is_officetel_context():
        d.update({
            "non_apt_sale_commentary": "",
            "non_apt_jeonse_commentary": "",
            "non_apt_rent_commentary": "",
            "presale_commentary": "",
            "non_apt_sale_regional_context": "",
            "non_apt_jeonse_regional_context": "",
            "non_apt_rent_regional_context": "",
            "rowhouse_multiplex_sale_summary": "",
            "detached_multifamily_sale_summary": "",
            "rowhouse_multiplex_jeonse_summary": "",
            "detached_multifamily_jeonse_summary": "",
            "rowhouse_multiplex_rent_summary": "",
            "detached_multifamily_rent_summary": "",
        })
    if is_weekly_report(data):
        weekly_sale = data.get("weekly_apt_sale_change", {}) or data.get("price_index", {}).get("sale", {})
        weekly_jeonse = data.get("weekly_apt_jeonse_change", {}) or data.get("price_index", {}).get("jeonse", {})
        sale_latest = format_float(weekly_sale.get("latest_value"), 2)
        jeonse_latest = format_float(weekly_jeonse.get("latest_value"), 2)
        d.update({
            "monthly_core_diagnosis": f"{bm} {region} 주간 주택동향은 과거 월의 월간 아파트 가격지수 전월 대비 변동률과 기준월 주간 아파트 매매·전세가격 변동률, 아파트 실거래 보조지표를 함께 판단함. 기준주 공표통계는 매매 {sale_latest}%, 전세 {jeonse_latest}%로 확인되며, 인구·세대·이동·미분양·공급·금리 통계는 배경지표로 활용함.",
            "price_index_commentary": f"주간 공표통계 시계열은 과거 월은 KOSIS 월간 전국주택가격동향조사 유형별 아파트 가격지수의 전월 대비 변동률을 사용하고, 기준월은 KOSIS 주간아파트 매매·전세가격변동률을 사용함. 기준주 값은 매매 {sale_latest}%, 전세 {jeonse_latest}%로 산출되며, 주간 값은 이미 변동률이므로 별도 지수형 재계산을 하지 않음.",
            "apt_rent_commentary": "",
            "non_apt_sale_commentary": "",
            "non_apt_jeonse_commentary": "",
            "non_apt_rent_commentary": "",
            "rent_market_summary": "",
            "apt_rent_summary": "",
            "non_apt_sale_regional_context": "",
            "non_apt_jeonse_regional_context": "",
            "non_apt_rent_regional_context": "",
            "rowhouse_multiplex_sale_summary": "",
            "detached_multifamily_sale_summary": "",
            "rowhouse_multiplex_jeonse_summary": "",
            "detached_multifamily_jeonse_summary": "",
            "rowhouse_multiplex_rent_summary": "",
            "detached_multifamily_rent_summary": "",
            "sale_market_summary": sale_context,
            "apt_sale_summary": sale_context,
            "jeonse_market_summary": jeonse_context,
            "apt_jeonse_summary": jeonse_context,
            "comprehensive_summary": (
                f"종합적으로 {bm} {region} 주간 주택동향은 아파트 매매·전세 공표통계와 실거래 보조지표를 함께 읽어야 함. "
                f"{_profile_context_for_summary()} {_demand_context()} "
                f"공표통계는 과거 월의 월간 가격지수 흐름과 기준월 주간 변동률을 연결해 확인하며, 기준주 값은 매매 {sale_latest}%, 전세 {jeonse_latest}%로 확인됨. 최근 흐름은 특정 수치 하나보다 상승·하락 신호가 이어지는지와 거래가 집중되는 생활권이 바뀌는지를 중심으로 해석하는 것이 적절함. "
                "비아파트와 월세는 주간보고서 분석 대상에서 제외하되, 인구·세대·이동·미분양·공급·금리 지표는 매수 심리와 임차수요의 배경 여건으로 연결해 조사상 유의사항으로 점검할 필요가 있음."
            )
        })
    return d



# 프롬프트 말미에 별도 라벨 블록으로 이미 노출되므로, 임베드 data_pack에서는 중복 제거하는 키.
_EVIDENCE_PRIORITY_RULES = """
[근거자료 역할 및 우선순위 - 최우선 규칙]
1. 조사 시황(market_field_context)
- 매물 증감, 현장 거래량, 매수 문의, 매도 호가, 관망세, 급매 소진, 매물 적체, 지역별 현장 분위기,
  개발·학군·생활권에 대한 조사자 판단의 1차 자료로 사용합니다.
- 종합서술의 원인·시장 분위기 문장은 이 자료를 가장 우선해 사용합니다.
2. data_pack.json
- 가격변동률, 실거래 건수, 상승·보합·하락 비율, 인구·세대·금리·입주물량 등 수치 검증자료로 사용합니다.
- data_pack의 실거래 상승·하락 비율이 조사 시황과 다르더라도 조사 시황을 잘못된 것으로 간주하지 않습니다.
- 실거래는 시장 방향을 결정하는 주지표가 아니라 조사 시황을 보완하는 보조지표입니다.
3. region_profile
- 생활권 구조, 학군, 교통, 신축·구축 특성 등 조사 시황에 배경을 보강하는 자료로 사용합니다.
4. 세 자료가 충돌할 경우
- 가격변동률과 매물·거래량은 조사 시황 기준값을 우선합니다.
- 수치 사실은 data_pack으로 검증합니다.
- 원인과 시장 분위기는 조사 시황에 근거합니다.
- 충돌을 억지로 하나의 결론으로 합치지 않고, 확인이 필요한 항목으로 남깁니다.
5. market_field_context.available이 false이면 1번 자료가 없는 상태이므로 region_profile과 data_pack만으로
   서술하되, 현장 분위기를 지어내지 않고 확인 가능한 범위로 한정합니다.
6. 조사 시황의 기준기간(period_label)은 report_period_label과 같아야 정상입니다. 두 값이 다르면 지난 기간 자료이므로
   그 내용을 이번 기간의 사실로 쓰지 않습니다.
7. 본문의 조사 시점은 항상 보고서 기준기간입니다. "○월 현장조사에서는", "○월 시장은"처럼 보고서 기준기간이 아닌
   월·주차를 조사 시점으로 표기하지 않습니다. 조사 시황 원문에 다른 월 표기가 있어도 그대로 옮기지 말고
   보고서 기준기간 기준으로 바꿔 씁니다.
8. 전월(전주) 대비 비교를 위한 이전 기간 언급은 허용합니다. 다만 "전월 대비", "전주 대비"처럼 상대 표현을 쓰고,
   매물량·거래량 원자료의 "직전월 수치(직전월)→기준월 수치(기준월)" 표기는 비교 근거로만 쓰며 조사 시점 표기로 오용하지 않습니다.

[서술 생성 순서 - 숫자가 아니라 시장 이야기가 먼저]
① 지역 시장 구조(region_profile) → ② 기간 중 현장 변화(조사 시황) → ③ 매물·거래·가격 흐름 →
④ 주요 읍면동별 원인 → ⑤ 실거래·인구·금리·공급 통계로 검증
- 수치로 문장을 시작해 해석을 붙이는 방식이 아니라, 시장 흐름을 먼저 쓰고 수치를 근거로 뒤에 붙입니다.
- 금지 예: "상승 거래 53.5%, 하락 거래 43.0%로 상승 우세함."
- 권장 예: "선호 생활권을 중심으로 매수세가 이어지는 가운데, 공급 부족과 주거 상향 이동 수요가 맞물리며
  매도자 우위 흐름이 지속되는 모습임. 실거래 보조지표에서도 상승 거래가 하락 거래보다 많은 것으로 확인됨."
- 위 예시의 문장 구조만 참고하고, 읍면동명·단지명은 반드시 조사 시황과 data_pack에서 확인된 값만 사용합니다.
"""


_LISTING_TRADE_RULES = """
[매물·거래량 해석 규칙]
1. 매물량과 거래량이 함께 증가하면 매수·매도 참여가 함께 확대된 것인지 가격 방향과 묶어 판단합니다.
2. 매물 감소 + 거래 증가는 매물 소진 및 매수세 유입 가능성을 우선 검토합니다.
3. 매물 증가 + 거래 감소는 관망세·매물 적체 가능성을 우선 검토합니다.
4. 매물 증가 + 가격 상승은 매물 증가만으로 약세로 해석하지 말고 선호 단지·호가 상승 여부와 함께 판단합니다.
5. 거래량은 기준월 실거래 잠정치와 조사 시황의 현장 조사 거래량을 혼동하지 않고 구분해 씁니다.
6. 조사 시황에 "매수 문의 증가", "매도 호가 유지", "급매 소진", "매물 적체", "관망세"가 기재되어 있으면
   종합서술의 핵심 시장 언어로 적극 반영합니다.
7. 매물량 자료가 없으면 매물 흐름을 추정하지 말고 거래량·가격 흐름 중심으로만 서술합니다.
"""


_FIELD_TONE_RULES = """
[현장 시황 문체 강화]
- 단순 통계 해설보다 조사자가 현장에서 관찰한 시장 분위기를 설명하는 문체로 작성합니다.
- 다음 시장 언어를 상황에 맞게 적극 활용합니다.
  매도자 우위, 매수자 우위, 관망세, 매수 문의, 매도 호가 유지·상향·하향, 급매물 소진, 매물 적체,
  거래 소강, 가격 하방 경직성, 수요 분산, 수요 이동, 갈아타기 수요, 신축 선호, 구축 기피, 선별 거래,
  지역별·단지별 차별화, 수급 불균형.
- 위 표현은 조사 시황·region_profile·data_pack에서 근거가 확인된 경우에만 사용합니다.
- "상승 거래 비중이 높아 가격이 상승함" 같은 기계적인 인과 문장은 쓰지 않습니다.
- 문장 끝은 "판단됨", "분석됨", "해석됨"만 반복하지 말고 "이어짐", "나타남", "유지됨", "확대됨", "축소됨",
  "전환됨", "~하는 양상임", "~로 조사됨"을 섞어 실제 조사 시황 문체에 가깝게 씁니다.
- 상승·보합·하락 비율은 문장 전면에 나열하지 말고 "우세", "과반", "혼조", "선별 거래"처럼 시장 언어로 바꾸고,
  비율 수치는 꼭 필요한 곳에서만 근거로 덧붙입니다.
- data_pack, 공표변동률, 수집·사용 제외, 신고기한 미도래 같은 내부 처리 용어는 보고서 문장에 노출하지 않습니다.
"""


_DONG_FIELD_CAUSE_RULES = """
[읍면동 현장 원인 서술 강제 규칙]
1. 주요 읍면동은 각각 아래 4개 중 최소 2개를 반드시 포함해 서술합니다.
   ① 어떤 수요층이 움직이는지
   ② 왜 그 지역을 선택하거나 이탈하는지
   ③ 신축·구축·학군·생활편의·직주근접·입주물량 중 무엇이 영향을 주는지
   ④ 매도자 또는 매수자의 실제 행동이 어떻게 변했는지
2. "○○동은 주요 주거지이며 상승 비중이 높아 견조함"처럼 통계가 원인이 되는 문장은 금지합니다.
3. 조사 시황의 rising_areas·falling_areas에 해당 읍면동 설명이 있으면 그 내용을 원인 문장의 1차 근거로 삼고,
   dong_profiles와 실거래 흐름으로 보강합니다.
4. 조사 시황에 없는 읍면동은 원인을 지어내지 말고 거래 흐름과 dong_profiles 범위에서만 서술합니다.
"""


_WEEKLY_ROLLUP_RULES = """
[기준월 주간 시황 취합 활용 규칙 - market_field_context.weekly_rollup이 있을 때만 적용]
1. weekly_rollup은 이번 보고서 기준월에 조사된 주간 현장 시황을 주차 순으로 모은 자료입니다.
   월간 조사 시황이 아직 작성되지 않아 이를 월중 흐름의 대체 근거로 사용합니다.
2. 주차별 내용을 그대로 나열하지 말고 "월초 → 월말" 흐름으로 합쳐서 씁니다.
   예: 월 전반에는 ○○ 흐름이 나타났으나 월 후반 들어 ○○로 전환되는 모습이 관찰됨.
3. 주차별 변동률·매물량·거래량은 월간 수치가 아닙니다. 이 값들을 당월 월간 변동률·거래량으로 쓰지 않으며,
   월간 수치는 data_pack에서 확인된 값만 사용합니다. weekly_rollup의 수치는 월중 방향 판단 근거로만 씁니다.
4. 주간 시황은 아파트 매매·전세만 다룹니다. 월세, 연립·다세대, 단독주택 서술의 근거로 사용하지 않습니다.
   해당 문단은 data_pack의 자료 범위에서만 제한적으로 서술합니다.
5. 주차별로 방향이 엇갈리면 하나로 합치지 말고 "월중 등락이 엇갈렸다", "월 후반 들어 방향이 바뀌었다"처럼
   변화 시점이 드러나게 씁니다.
6. 읍면동 원인, 매수·매도 반응, 지역 이슈는 weekly_rollup의 rising_areas·falling_areas·local_issue를
   1차 근거로 사용하되, 특정 주차에만 나온 일시적 요인은 월간 전체 요인으로 단정하지 않습니다.
7. 시도단위 보고서에서는 시군구별로 묶어 읽고, 같은 달 안에서 시군구 간 흐름 차이를 설명합니다.
"""


_MONTHLY_HOUSING_FIELD_RULES = """
[월간 주택 조사 시황 항목 대응 규칙 - 월간 주택동향보고서 전용]
1. market_field_context의 각 섹션(sale·jeonse·rent) 항목은 아래 서술 key의 1차 근거입니다.
   - overall               → sale_market_summary / jeonse_market_summary / rent_market_summary
   - apartment             → apt_sale_summary / apt_jeonse_summary / apt_rent_summary
   - rowhouse_multiplex    → rowhouse_multiplex_sale_summary / _jeonse_summary / _rent_summary
   - detached              → detached_multifamily_sale_summary / _jeonse_summary / _rent_summary
   - rising_areas·falling_areas → 각 아파트 문단의 읍면동 서술
   - local_issue           → comprehensive_summary의 지역 이슈 부분
2. 변동률은 주택종합과 아파트를 구분해 씁니다.
   - price_change·transaction_change는 주택종합 기준입니다. 종합 문단(sale_market_summary 등)에만 씁니다.
   - apartment_price_change·apartment_transaction_change는 아파트 기준입니다. 아파트 문단에만 씁니다.
   - 두 값을 섞어 쓰거나, 주택종합 수치를 아파트 수치인 것처럼 쓰지 않습니다.
3. listing_change는 아파트 매물증감추이입니다. 매물 흐름은 [매물·거래량 해석 규칙]에 따라 거래량·가격 방향과 묶어 해석하고,
   연립·다세대·단독주택 문단에 아파트 매물량을 근거로 쓰지 않습니다.
4. 연립·다세대와 단독주택은 거래 규모가 작으므로 rowhouse_multiplex·detached에 적힌 범위를 넘어 방향을 단정하지 않습니다.
   해당 항목이 비어 있으면 data_pack의 non_apt_* 자료 범위에서만 제한적으로 서술합니다.
5. 조사 시황에 기재된 변동률 방향어(상승 전환, 하락 전환, 상승폭 확대, 상승폭 축소, 하락폭 확대, 하락폭 축소, 전월 동일)를
   그대로 사용하고, data_pack 수치로 방향어를 새로 만들지 않습니다.
6. local_issue에 원자료 확인 필요 같은 검수 메모가 있으면 보고서 본문에 옮기지 말고, 해당 수치를 단정적으로 쓰지 않는 근거로만 활용합니다.
"""


_OFFICETEL_FIELD_RULES = """
[오피스텔 현장 시황 서술 규칙 - 오피스텔동향보고서 전용, 매우 중요]
1. 오피스텔은 아파트의 축소판이 아니라 수익형·임대형 상품입니다. 매매는 투자 수요와 실수요를 구분하고,
   전세·월세는 임차 수요층의 성격(대학생, 사회초년생, 1~2인 직장인, 신혼부부, 산업단지 근로자, 의료·공공기관 종사자)을
   먼저 특정한 뒤 가격 흐름을 연결합니다.
2. 매매 서술에서는 아래를 우선 검토합니다.
   - 아파트 선호 현상에 따른 오피스텔 대체재 매력도 저하
   - 수익형 부동산으로서의 투자가치(임대수익률, 공실률, 대출 여건)
   - 신축·준신축 선호와 노후 오피스텔 기피, 주거형(아파트 평형 유사) 대 원룸형의 수요 차이
   - 개발 호재가 있어도 매수 수요가 아파트로 쏠려 오피스텔에는 제한적으로만 반영되는지 여부
3. 전세 서술에서는 아래를 우선 검토합니다.
   - 전세가율 수준과 역전세·깡통전세 우려
   - HUG 전세보증보험 가입 요건 강화에 따른 안전 매물 부족과 그로 인한 전세 수요의 월세 이동
   - 보증 가입이 가능한 단지·지역으로의 쏠림과 그렇지 못한 노후 단지의 약세라는 양극화 구도
4. 월세 서술에서는 아래를 우선 검토합니다.
   - 전세 회피에 따른 '월세화' 현상과 반전세·보증부월세 전환
   - 직주근접·역세권·업무지구·산업단지·대학가 배후수요
   - 소형 평형(전용 20~40㎡ 안팎) 품귀와 임대인 우위 형성 여부
   - 청년 전용 보증부월세 대출 등 제도적 지원이 월세 선호를 강화하는지
5. 대학가 배후 지역은 학기·방학에 따른 계절적 성수기·비수기를 거래량 해석에 반영합니다.
6. 오피스텔 면적은 전용 ㎡로 표기하고 평형으로 환산하지 않습니다. 임대조건은 조사 시황에 기재된 경우에 한해
   "보증금 1000/월세 75만원", "보증금 300~500만 원, 월세 45~55만 원 선"처럼 실제 조사 표기를 따릅니다.
7. 오피스텔 보고서에서는 "아파트"를 비교 대상으로만 언급하고, 아파트 시장 자체를 서술하지 않습니다.
   "비아파트", "연립·다세대", "단독·다가구", "분양권전매"라는 표현은 사용하지 않습니다.
8. 거래량이 연간 수십 건 수준으로 적은 지역은 방향성을 단정하지 말고 "간헐적 거래", "거래 공백",
   "실수요 위주의 제한적 거래"처럼 표현합니다.
"""


_SIDO_FIELD_CONTEXT_RULES = """
[시도단위 조사 시황 활용 규칙 - meta.report_scope가 sido일 때만 적용]
1. market_field_context.sigungu_contexts에는 선택 시도에 속한 시군구별 조사 시황이 배열로 들어 있습니다.
   특정 시군구 하나의 시황을 시도 전체 흐름처럼 옮겨 쓰지 않습니다.
2. 종합서술은 "시도 전체의 공통 흐름 → 시군구별 편차 → 편차가 생긴 이유" 순서로 작성합니다.
3. 공통 흐름은 여러 시군구 조사 시황에서 반복적으로 확인되는 요인(금리 부담, 보증보험 요건 강화, 입주물량,
   개발 호재, 아파트 선호 등)만 시도 전체 요인으로 승격해 서술합니다.
4. 편차는 방향이 다른 시군구를 최소 2곳 이상 대비해 설명하고, 각 시군구의 원인을 해당 시군구 조사 시황에서 가져옵니다.
5. 시군구를 나열만 하지 말고, 어느 시군구가 시도 전체 흐름을 끌어올리거나 끌어내렸는지를 밝힙니다.
6. 조사 시황이 없는 시군구는 있는 것처럼 쓰지 말고, 자료가 확보된 시군구 범위 내에서 해석한다는 점이 드러나게 씁니다.
7. 읍면동 단위 서술은 조사 시황에 읍면동 원인이 기재된 시군구에 한해 사용하고, 시도 전체를 읍면동으로 설명하지 않습니다.
"""


def _market_field_block(data_pack: Dict[str, Any], is_weekly: bool) -> str:
    """조사 시황(market_field_context)을 프롬프트 상단 블록으로 직렬화한다."""
    mfc = data_pack.get("market_field_context", {}) or {}
    asset = str(mfc.get("asset_label") or "주택")
    period_word = "금주" if is_weekly else "당월"
    scope_word = "시도단위 시군구별" if str(mfc.get("scope")) == "sido" else ""
    label = " ".join(x for x in [period_word, scope_word, asset, "조사 시황"] if x)
    note = str(mfc.get("note", ""))
    report_period = str(mfc.get("report_period_label", ""))
    if not mfc.get("available"):
        if mfc.get("weekly_rollup"):
            view = {k: mfc.get(k) for k in ("dataset_label", "report_period_label", "weekly_rollup_note", "weekly_rollup")}
            return (
                f"[market_field_context / {label} - 기준월 주간 시황 취합]\n"
                f"※ {note}\n"
                f"※ 보고서 기준기간은 {report_period or '메타 참조'}입니다. 주차별 내용을 그대로 옮기지 말고 "
                "월중 흐름으로 합쳐 쓰고, 주차 수치를 월간 수치로 쓰지 마세요.\n"
                f"{json.dumps(view, ensure_ascii=False, indent=2)}\n"
            )
        return (
            f"[market_field_context / {label}]\n"
            f"※ 조사 시황 자료가 없습니다({note}). 현장 분위기·매수매도 반응을 지어내지 말고 "
            "region_profile과 data_pack에서 확인되는 범위로만 서술하세요.\n"
        )
    usage = str(mfc.get("usage_rule", ""))
    return (
        f"[market_field_context / {label} - 원인·시장 분위기의 1차 근거]\n"
        f"※ {usage}\n"
        f"※ {note}\n"
        f"※ 보고서 기준기간은 {report_period or '메타 참조'}입니다. 본문의 조사 시점은 이 기간으로만 표기하고, "
        "조사 시황 원문에 다른 월·주차 표기가 있어도 그대로 옮기지 마세요.\n"
        f"{json.dumps(mfc, ensure_ascii=False, indent=2)}\n"
    )


_LLM_SEPARATE_BLOCK_KEYS = (
    "region_profile",
    "market_field_context",
    "monthly_change_summary",
    "weekly_change_summary",
    "dong_focus_summary",
    "detail_statistics_summary",
    "case_focus_summary",
    "non_apt_interpretation_summary",
    "supply_pressure_summary",
    "weekly_apt_sale_change",
    "weekly_apt_jeonse_change",
)


def _strip_html_only_rows(node: Any) -> None:
    """LLM 해석에 불필요한 전체 거래내역(all_*_rows)을 재귀적으로 제거한다.
    상위 사례(top_*_rows)와 월별·읍면동·구간 집계는 그대로 유지한다."""
    if isinstance(node, dict):
        for k in list(node.keys()):
            if k.startswith("all_") and isinstance(node[k], list):
                del node[k]
            else:
                _strip_html_only_rows(node[k])
    elif isinstance(node, list):
        for item in node:
            _strip_html_only_rows(item)


def build_llm_data_view(data_pack: Dict[str, Any], drop_keys: Tuple[str, ...] = ()) -> Dict[str, Any]:
    """LLM 프롬프트 임베드용 경량 data_pack 사본.
    - all_*_rows(HTML 전체 거래내역 표 전용)와 input_files(파일명)를 제거
    - drop_keys(별도 라벨 블록으로 노출되는 요약 등)를 제거
    원본 data_pack(HTML 렌더·저장용)은 변경하지 않는다."""
    view = copy.deepcopy(data_pack)
    view.pop("input_files", None)
    for k in drop_keys:
        view.pop(k, None)
    _strip_html_only_rows(view)
    return view


def _llm_named_block(name: str, value: Any) -> str:
    return f"[{name}]\n{json.dumps(value, ensure_ascii=False, indent=2)}"


def _llm_data_pack_block(data_pack: Dict[str, Any], drop_keys: Tuple[str, ...] = ()) -> str:
    """프롬프트 말미의 전체 분석자료 블록(경량·compact 직렬화)."""
    view = build_llm_data_view(data_pack, drop_keys=drop_keys)
    return (
        "[data_pack.json / 전체 분석자료]\n"
        "※ HTML 표 전용 전체 거래내역(all_*_rows)과 위 별도 블록은 토큰 절약을 위해 제외했습니다. "
        "상위 사례는 top_*_rows, 추세는 monthly_rows/table_rows, 지역별은 dong_rows를 사용하세요.\n"
        f"{json.dumps(view, ensure_ascii=False)}"
    )


_SEASON_HINTS = {
    1: "겨울 비수기 속 봄 이사철·신학기를 앞둔 대기 수요",
    2: "봄 이사철·대학 신학기를 앞둔 임차 수요",
    3: "봄 이사철과 대학 개강",
    4: "봄 이사철 마무리",
    5: "봄 이사철 이후 거래 소강",
    6: "여름 비수기 진입",
    7: "여름 비수기·방학",
    8: "여름 비수기 막바지와 가을 이사철·2학기 개강 대기 수요",
    9: "가을 이사철과 대학 2학기 개강",
    10: "가을 이사철",
    11: "가을 이사철 마무리",
    12: "연말 비수기",
}


def field_style_summary_rules(data_pack: Dict[str, Any]) -> str:
    """조사자 시황 형식(첫 줄 요약·상승/하락 지역·원인 중심·단정형)에 맞춘 Ⅵ장 공통 규칙."""
    meta = (data_pack or {}).get("meta") or {}
    try:
        month_no = int(str(meta.get("base_month") or "")[5:7])
    except Exception:
        month_no = 0
    season = _SEASON_HINTS.get(month_no, "")
    apt_rent = (data_pack or {}).get("apt_rent") or {}
    volume = {
        "apt_trade": ((data_pack or {}).get("apt_trade") or {}).get("volume_reference"),
        "apt_jeonse": (apt_rent.get("jeonse") or {}).get("volume_reference"),
        "apt_rent": (apt_rent.get("monthly_rent") or {}).get("volume_reference"),
    }
    return f"""
[조사자 시황 형식 규칙 - Ⅵ장 전체 적용, 위 규칙과 충돌하면 이 규칙을 우선]
1. 아파트 매매·전세·월세 문단 위에는 거래량과 상승·하락 비중(가격지수 변동률이 있으면 변동률)의 전월→당월 요약 줄이 자동으로 붙습니다. 같은 수치를 문단 첫머리에서 되풀이하지 않습니다.
2. apt_sale_summary·apt_jeonse_summary·apt_rent_summary는 5~7문장으로 씁니다. 순서는 ① 당월 흐름과 그 원인 1~2문장 → ② "▲ 상승 지역(읍면동명): " 뒤에 그 생활권에서 왜 오르는지 2문장 → ③ "▼ 하락 지역(읍면동명): " 뒤에 왜 약한지 1~2문장 → ④ 당월 판단 1문장입니다.
3. 상승·하락 지역은 dong_rows(거래 상위 읍면동)에서 비교가능 거래가 5건 이상이고 상승 또는 하락 비중이 가장 뚜렷한 곳을 하나씩 고릅니다. 뚜렷한 하락 지역이 없으면 "▼ 하락 지역" 대신 "◆ 혼조 지역(읍면동명): "으로 씁니다.
4. 면적·가격·보증금·월세·연식 구간 통계는 문단당 1문장까지만 씁니다. 나머지 문장은 그런 흐름이 나타난 이유(생활권 성격, 배후수요, 신축 입주·공급, 정비사업, 학군, 금리·대출, 임차→매매 전환 등)를 data_pack·region_profile 근거로 설명합니다.
5. 시기 요인: 기준월은 "{season}" 시기입니다. 대학가·학군지 읍면동 거래가 늘었거나 갱신·신규 계약이 몰린 것이 data_pack에서 확인될 때만 원인으로 씁니다.
6. 거래량 수준: 아래 [volume_reference]의 최근 12개월 월평균·1년 전 월평균과 비교해 "연평균 대비 저조/활발"을 판단합니다. 전월 거래량이 최근 13개월 최저였다면 당월 증가는 "기저효과"로 씁니다. 기준월이 신고기한 미도래 잠정치이면 거래량 증가·감소 방향을 단정하지 않습니다.
7. 실거래 상승·하락 비중으로 판단한 방향은 "실거래 기준"을 붙이고, 가격지수 변동률과 방향이 다르면 둘을 구분해 씁니다.
8. 문장은 "~로 분석됨", "~로 조사됨", "~양상임"처럼 단정형으로 맺습니다. "점검 필요", "확인 필요", "추가 점검", "지켜볼 필요" 같은 문장은 comprehensive_summary 마지막 1문장에서만 한 번 씁니다.
9. 자료가 없거나 확보되지 않았다는 문장(예: "현장 시황 자료가 확보되지 않아", "공표치가 확인되지 않은 만큼")은 쓰지 않습니다. 자료가 없는 항목은 언급하지 않습니다.

[volume_reference]
{json.dumps(volume, ensure_ascii=False)}
"""


# ---------------------------------------------------------------------------
# 한국부동산원 시황 형식 프롬프트 틀: 페르소나 → 문체 → 용어 사용 기준 → Python 사전 판정 → 문단 틀 → 자기검수
# 방향어·국면·거래 수준은 Python에서 미리 판정해 넘기고, LLM은 그 판정을 시장 언어로 풀어 쓰는 역할만 맡는다.
# ---------------------------------------------------------------------------
_KREB_PERSONA = """[1. 역할 - 페르소나]
당신은 한국부동산원 지사에서 주택가격동향조사(월간·주간)와 실거래 신고자료 검증을 오래 담당해 온 부동산 시장 전문조사자입니다.
- 하는 일: 표본 아파트 가격조사, 중개업소 현장 탐문, 실거래 신고자료 검증, 담당 지역 시황 보고서 작성
- 독자: 본사 주택통계 담당 부서, 지사장, 같은 지역을 조사하는 조사자
- 목표: 숫자를 옮겨 적는 통계 해설이 아니라, 숫자가 가리키는 시장 국면과 그 원인을 조사자의 언어로 판단해 전달
- 태도: 근거가 있는 것만 단정하고, 근거 없는 원인·전망은 쓰지 않음. 과장·홍보성·투자 권유성 표현은 쓰지 않음
"""

_KREB_STYLE_RULES = """[2. 한국부동산원 시황 문체]
1. 문장은 명사형·단정형으로 맺음(~임, ~됨, ~양상임, ~로 조사됨, ~로 분석됨). "~습니다", "~요"체는 쓰지 않음.
2. 한 문장에는 판단 하나만 담고 80~120자 안팎으로 끊음.
3. 가격 문장의 기본형은 "[대상] 가격은 [원인·단지 유형] 위주로 [방향어]"임.
   예) 매매가격은 정주여건이 양호한 대단지 위주로 상승폭이 축소되는 양상임.
   예) 전세가격은 신규 입주 영향이 있는 생활권 위주로 하락폭이 축소됨.
4. 여러 지역을 함께 쓸 때는 "○○동은 ~ 위주로, △△동은 ~ 위주로 [방향어]"처럼 지역별 원인을 나란히 씀.
5. 숫자는 판단 뒤에 괄호 근거로 붙임. 예) 하락폭이 축소됨(−0.17%→0.00%).
6. 표기: 변동률은 소수 둘째 자리 %(+0.25%, −0.31%), 변동률 차이는 %p, 거래는 천 단위 쉼표와 '건', 금액은 억 원·만 원,
   면적은 전용 ㎡(평 환산 금지), 비교 기준은 "전월 대비", "전년 동월 대비", "최근 12개월 월평균 대비"로 밝힘.
7. 금지: 폭등·폭락·패닉·불장·영끌·대박·꿈틀 같은 수사, "오를 것임" 같은 전망 단정, 이모지, "~것 같음", 영어 약어 남용.
"""

_KREB_TERM_RULES = """[3. 부동산 전문용어 사용 기준 - 용어마다 아래 근거가 확인될 때만 사용]
■ 가격 방향어(상승 전환 / 상승폭 확대·축소·유지 / 보합 전환·유지 / 하락 전환 / 하락폭 확대·축소·유지)
  → [4. Python 사전 판정]의 index_flow.latest_phrase를 그대로 씀. 조사 시황(market_field_context)에 방향어가 있으면 조사 시황 방향어를 우선함.
  → 국면은 index_flow.current_run으로 "○월 이후 ○개월 연속 상승(하락)"처럼 씀. 직접 계산해 다른 방향어를 만들지 않음.
■ 실거래 방향어(상승 우세 / 하락 우세 / 보합 우세 / 혼조) → trade_flow.verdict를 쓰고 반드시 "실거래 기준"을 붙임.
  공표 변동률 방향과 다르면 "변동률은 ○○이나, 실거래 기준으로는 ○○"처럼 구분함.
  trade_flow.sample_note가 있으면 실거래 방향을 단정하지 않고 "거래가 제한적"으로만 씀.
■ 거래량 수준(거래 활발 / 평년 수준 / 거래 위축) → trade_flow.level. 최근 신고가 끝난 달의 거래를 최근 12개월 월평균과 비교한 판정임.
■ 기저효과 → trade_flow.base_effect가 true일 때만.
■ 관망세·매수심리 위축 → 같은 달에 거래 감소와 하락 비중 확대가 함께 확인되거나 조사 시황에 기재된 경우.
■ 매수세 유입·매수심리 회복 → 같은 달에 거래 증가와 상승 비중 확대가 함께 확인되거나 조사 시황에 기재된 경우.
■ 급매물 소진, 매물 적체(누적), 호가 상향·하향, 매도자 우위·매수자 우위 → 조사 시황에 기재된 경우에만. 실거래 수치만으로 쓰지 않음.
■ 선호단지·대단지·신축·구축·역세권·학군지·정주여건·직주근접·생활편의 → region_profile·dong_profiles 또는 거래사례(단지·사용승인연도)로 확인될 때.
■ 매매·전세 가격차(갭), 전세가율 → data_pack에 해당 수치가 없으면 쓰지 않음.
■ 신규계약·갱신계약 → 전세·월세 가격 비교는 신규계약 기준. 갱신 비중이 높은 달은 "갱신 수요 집중"으로 거래량 해석에만 씀.
■ 환산월세·전월세전환율 → 월세 비교는 환산월세(월세+보증금×전환율÷100÷12) 기준이며 conversion_rate의 적용 전환율을 근거로 씀.
■ 준공 후 미분양(악성 미분양) → 총 미분양과 방향이 다르면 둘을 함께 씀(background_flow.unsold).
■ 잠정 → 기준월 실거래가 신고 진행 중이면(trade_flow.partial_note) 거래량 증감·방향을 단정하지 않고 "잠정"으로만 씀.
"""

_KREB_PARAGRAPH_FRAME = """[5. 문단 틀 - key별 서술 순서(슬롯 이름은 출력하지 않음)]
- comprehensive_summary(6~8문장): ① 총평(매매·전세·월세 방향을 한 문장으로) → ② 13개월 국면(index_flow.phase_text의 출발점→전환점→현재)
  → ③ 거래(거래량 수준과 실거래 방향) → ④ 원인(조사 시황·지역특성·정책 시점) → ⑤ 배경(background_flow 중 변화가 큰 1~2개) → ⑥ 유의사항 1문장
- sale_market_summary·jeonse_market_summary·rent_market_summary(4~6문장): ① 방향어 문장(latest_phrase+괄호 수치) → ② 13개월 국면 1문장
  → ③ 거래량 수준·실거래 방향 1문장 → ④ 원인 1~2문장 → ⑤ 판단 1문장
- apt_sale_summary·apt_jeonse_summary·apt_rent_summary: [조사자 시황 형식 규칙] 2번 순서(당월 흐름·원인 → ▲ 상승 지역 → ▼ 하락 지역 → 판단)를 따르되 첫 문장은 방향어로 시작
- population·age_migration·rate·unsold·supply·price_index·conversion_rate_commentary 등 시계열 해설(2~4문장):
  ① 13개월 흐름(출발점→전환점→현재, background_flow·index_flow 문구 활용) → ② 당월 변화 → ③ 주택시장에 주는 의미(수요 기반, 매수·임차 부담, 공급 부담)
- apt_trade_commentary·apt_jeonse_commentary·apt_rent_commentary(3~5문장): ① 거래량 13개월 흐름(최다·최소 달, 최근 12개월 월평균 대비) → ② 실거래 방향의 변화
  → ③ 변곡점과 같은 시기 정책(있을 때) → ④ 당월 판단
- monthly_core_diagnosis(3~5문장): 매매·전세·월세 각각 "방향어 + 국면 + 실거래 방향"을 한 문장씩, 마지막 문장은 당월 가장 큰 변화
"""

_KREB_SELF_CHECK = """[7. 출력 전 자기검수 - 하나라도 어긋나면 고친 뒤 출력]
□ 모든 수치·단지명·읍면동명이 data_pack 또는 [4. Python 사전 판정]에 있는 값인가
□ 가격 방향어가 index_flow.latest_phrase(조사 시황이 있으면 조사 시황 방향어)와 같은가
□ 실거래 방향에 "실거래 기준"을 붙였고, 공표 변동률 방향과 다르면 구분해 썼는가
□ 시계열 해설에 "출발점 → 전환점 → 현재" 흐름이 들어갔는가
□ 조사 시황에만 근거가 있는 용어(급매물 소진·매물 적체·호가)를 수치만으로 쓰지 않았는가
□ 금지 표현(폭등·폭락·불장·영끌·전망 단정·"~습니다")이 없는가
□ JSON 객체 하나만 출력했는가
"""


def _kreb_period(label: Any) -> str:
    """'2026-02' → '2026.2', '2026-10-19' → '10.19주'."""
    s = str(label or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{int(m.group(2))}.{int(m.group(3))}주"
    m = re.match(r"^(\d{4})-(\d{2})$", s)
    if m:
        return f"{m.group(1)}.{int(m.group(2))}"
    return s


def _kreb_pct(v: float) -> str:
    return f"{v:+.2f}%".replace("+0.00%", "0.00%")


def _kreb_index_flow(rows: Any) -> Dict[str, Any]:
    """공표 변동률 시계열을 상승·보합·하락 국면으로 묶어 출발점→전환점→현재 문구와 최신 방향어를 만든다."""
    pts = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        v = _clean_num(r.get("public_index_change_rate"))
        if v is None:
            continue
        pts.append((_kreb_period(r.get("period_label") or r.get("week") or r.get("month")), round(float(v), 2)))
    pts = pts[-13:]
    if len(pts) < 2:
        return {}
    word = {1: "상승", 0: "보합", -1: "하락"}
    runs: List[Dict[str, Any]] = []
    for lab, v in pts:
        s = (v > 0) - (v < 0)
        if runs and runs[-1]["s"] == s:
            runs[-1]["end"] = lab
            runs[-1]["vals"].append(v)
        else:
            runs.append({"s": s, "start": lab, "end": lab, "vals": [v]})

    def _span(r: Dict[str, Any]) -> str:
        return r["start"] if r["start"] == r["end"] else f"{r['start']}~{r['end']}"

    def _range(vals: List[float]) -> str:
        if len(vals) == 1:
            return _kreb_pct(vals[0])
        peak = max(vals, key=abs)
        text = f"{_kreb_pct(vals[0])}→{_kreb_pct(vals[-1])}"
        return text + (f", 최대 {_kreb_pct(peak)}" if peak not in (vals[0], vals[-1]) else "")

    last = runs[-1]
    weeks = [lab.endswith("주") for lab, _ in pts[-len(last["vals"]):]]
    unit = "주" if all(weeks) else ("개월" if not any(weeks) else "기간")
    return {
        "phase_text": " → ".join(f"{_span(r)} {word[r['s']]}({_range(r['vals'])})" for r in runs),
        "turning_points": [f"{r['start']} {word[r['s']]} 전환" for r in runs[1:]],
        "latest_period": pts[-1][0],
        "latest_phrase": f"{_rate_change_phrase(pts[-2][1], pts[-1][1])}({_kreb_pct(pts[-2][1])}→{_kreb_pct(pts[-1][1])})",
        "current_run": (f"{last['start']} 이후 {len(last['vals'])}{unit} 연속 {word[last['s']]}" if len(last["vals"]) > 1
                        else f"{last['start']} {word[last['s']]}(직전 국면에서 바뀐 첫 {'주' if unit == '주' else '달'})"),
    }


def _kreb_trade_flow(section: Any, meta: Dict[str, Any]) -> Dict[str, Any]:
    """실거래 13개월 거래량·방향 판정: 신고가 끝난 달 기준 거래 수준, 최다·최소 달, 기저효과, 실거래 방향."""
    if not isinstance(section, dict):
        return {}
    rows = [r for r in section.get("monthly_rows") or [] if isinstance(r, dict) and r.get("month")]
    if len(rows) < 3:
        return {}
    base = str((meta or {}).get("base_month") or "")[:7]
    partial = _base_month_is_partial(meta or {}) and str(rows[-1].get("month"))[:7] == base
    done = rows[:-1] if partial else rows
    counts = [(str(r["month"])[:7], int(_clean_num(r.get("trade_count")) or 0)) for r in done]
    (pm, pc), (cm, cc) = counts[-2], counts[-1]
    peak = max(counts, key=lambda x: x[1])
    low = min(counts, key=lambda x: x[1])
    ref = section.get("volume_reference") or {}
    avg = _clean_num((ref.get("recent") or {}).get("avg"))
    level = ""
    if avg:
        level = "거래 활발" if cc >= avg * 1.15 else ("거래 위축" if cc <= avg * 0.85 else "평년 수준")
    cur = done[-1]
    up, down = _clean_num(cur.get("up_ratio")) or 0.0, _clean_num(cur.get("down_ratio")) or 0.0
    verdict = str(cur.get("판정") or direction_verdict(cur.get("up_ratio"), cur.get("flat_ratio"), cur.get("down_ratio")))
    if verdict in ("상승우세", "하락우세") and abs(up - down) < 5:
        verdict = "혼조"
    out: Dict[str, Any] = {
        "latest_complete_month": _kreb_period(cm),
        "latest_complete_count": cc,
        "mom": f"{_kreb_period(pm)} {pc:,}건 → {_kreb_period(cm)} {cc:,}건" + (f"({(cc - pc) / pc * 100:+.1f}%)" if pc else ""),
        "level": level,
        "recent_12m_avg": avg,
        "vs_prev_year_pct": ref.get("vs_prev_year_pct"),
        "peak": f"{_kreb_period(peak[0])} {peak[1]:,}건",
        "low": f"{_kreb_period(low[0])} {low[1]:,}건",
        "base_effect": bool(pc == low[1] and cc > pc),
        "verdict": verdict.replace("우세", " 우세"),
        "up_down": f"상승 {up:.1f}% / 하락 {down:.1f}%",
    }
    compare = int(_clean_num(cur.get("compare_count")) or 0)
    if verdict == "산출불가" or compare < 10:
        out["sample_note"] = f"비교 가능 거래 {compare}건으로 실거래 방향 판단 제한적"
    if "renew_count" in cur:
        total = int(_clean_num(cur.get("trade_count")) or 0)
        renew = int(_clean_num(cur.get("renew_count")) or 0)
        if total:
            out["renew_share"] = f"갱신 {renew:,}건({renew / total * 100:.1f}%)"
    if partial:
        last = rows[-1]
        out["partial_note"] = f"{_kreb_period(base)} {int(_clean_num(last.get('trade_count')) or 0):,}건은 신고 진행 중 잠정치"
    return out


def _kreb_background_flow(data_pack: Dict[str, Any]) -> Dict[str, Any]:
    """인구·세대, 금리, 미분양 13개월 흐름 문구."""
    out: Dict[str, Any] = {}

    def _rows(key: str) -> List[Dict[str, Any]]:
        return [r for r in ((data_pack.get(key) or {}).get("table_rows") or []) if isinstance(r, dict) and r.get("month")]

    pop = _rows("population")
    if len(pop) >= 2:
        a, b = pop[0], pop[-1]
        p0, p1 = _clean_num(a.get("population")), _clean_num(b.get("population"))
        h0, h1 = _clean_num(a.get("households")), _clean_num(b.get("households"))
        neg = sum(1 for r in pop[1:] if (_clean_num(r.get("population_mom")) or 0) < 0)
        if p0 and p1:
            out["population"] = (f"인구 {_kreb_period(a['month'])} {p0:,.0f}명 → {_kreb_period(b['month'])} {p1:,.0f}명({p1 - p0:+,.0f}명, "
                                 f"{len(pop) - 1}개월 중 {neg}개월 감소)")
        if h0 and h1:
            out["households"] = f"세대 {_kreb_period(a['month'])} {h0:,.0f}세대 → {_kreb_period(b['month'])} {h1:,.0f}세대({h1 - h0:+,.0f}세대)"
    rates = _rows("rates")
    if len(rates) >= 2:
        for key, name in (("mortgage_rate", "주택담보대출"), ("jeonse_rate", "전세자금대출")):
            vals = [(r["month"], _clean_num(r.get(key))) for r in rates if _clean_num(r.get(key)) is not None]
            if len(vals) >= 4:
                (m0, v0), (m3, v3), (m1, v1) = vals[0], vals[-4], vals[-1]
                out[key] = (f"{name} 금리 {_kreb_period(m0)} {v0:.2f}% → {_kreb_period(m1)} {v1:.2f}%({v1 - v0:+.2f}%p), "
                            f"최근 3개월 {v1 - v3:+.2f}%p")
    unsold = _rows("unsold")
    if len(unsold) >= 2:
        for key, name in (("unsold_units", "미분양"), ("after_completion_unsold", "준공 후 미분양")):
            vals = [(r["month"], _clean_num(r.get(key))) for r in unsold if _clean_num(r.get(key)) is not None]
            if len(vals) >= 2 and any(v for _, v in vals):
                (m0, v0), (m1, v1) = vals[0], vals[-1]
                jump = max(zip(vals, vals[1:]), key=lambda p: abs(p[1][1] - p[0][1]))
                out[key] = (f"{name} {_kreb_period(m0)} {v0:,.0f}호 → {_kreb_period(m1)} {v1:,.0f}호({v1 - v0:+,.0f}호), "
                            f"가장 큰 변화 {_kreb_period(jump[1][0])} {jump[1][1] - jump[0][1]:+,.0f}호")
    return out


def build_kreb_verdicts(data_pack: Dict[str, Any]) -> Dict[str, Any]:
    """LLM에 넘길 Python 사전 판정: 공표 변동률 국면, 실거래 거래 수준·방향, 배경지표 흐름."""
    meta = data_pack.get("meta") or {}
    weekly = is_weekly_report(data_pack)
    pi = data_pack.get("price_index") or {}
    apt_rent = data_pack.get("apt_rent") or {}
    def _with_rows(*parts: Any) -> Any:
        return next((p for p in parts if isinstance(p, dict) and p.get("table_rows")), None)

    index_src = {
        "sale": _with_rows(data_pack.get("weekly_apt_sale_change") if weekly else None, pi.get("sale")),
        "jeonse": _with_rows(data_pack.get("weekly_apt_jeonse_change") if weekly else None, pi.get("jeonse")),
        "rent": None if weekly else pi.get("rent"),
    }
    trade_src = {"sale": data_pack.get("apt_trade"), "jeonse": apt_rent.get("jeonse"), "rent": None if weekly else apt_rent.get("monthly_rent")}
    out: Dict[str, Any] = {"asset": str(meta.get("primary_asset_label") or "아파트")}
    for key, label in (("sale", "매매"), ("jeonse", "전세"), ("rent", "월세")):
        part = {"index_flow": _kreb_index_flow((index_src[key] or {}).get("table_rows")),
                "trade_flow": _kreb_trade_flow(trade_src[key], meta)}
        part = {k: v for k, v in part.items() if v}
        if part:
            out[label] = part
    background = _kreb_background_flow(data_pack)
    if background:
        out["background_flow"] = background
    return out


def kreb_prompt_frame(data_pack: Dict[str, Any]) -> str:
    """프롬프트 상단 틀: 페르소나·문체·용어 기준·Python 사전 판정·문단 틀."""
    verdicts = build_kreb_verdicts(data_pack)
    return (
        f"{_KREB_PERSONA}\n{_KREB_STYLE_RULES}\n{_KREB_TERM_RULES}\n"
        "[4. Python 사전 판정 - 방향어·국면·거래 수준의 기준값]\n"
        "※ 아래 값은 Python이 data_pack에서 계산한 판정입니다. 방향어와 국면은 이 문구를 그대로 쓰고, 다른 계산으로 바꾸지 않습니다.\n"
        "※ index_flow=공표 변동률 국면(phase_text: 출발점→전환점→현재), trade_flow=실거래 거래 수준·방향, background_flow=배경지표 13개월 흐름.\n"
        f"{json.dumps(verdicts, ensure_ascii=False, indent=1)}\n\n"
        f"{_KREB_PARAGRAPH_FRAME}\n"
        "[6. 세부 작성 규칙]\n"
        "아래 규칙은 위 1~5번 틀 안에서 적용합니다. 충돌하면 [조사자 시황 형식 규칙]과 1~5번 틀을 우선합니다.\n"
    )


_CITY_WORDS = ("서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종", "제주")
_GENERIC_GU = {"동구", "서구", "남구", "북구", "중구"}
_UNSOLD_NEWS_WORDS = ("미분양", "할인분양", "잔여세대", "선착순 분양", "선착순분양", "준공 후 미분양", "악성 미분양", "분양가 할인")


def _unsold_region_words(region: str, dong_names: List[str]) -> Tuple[List[str], List[str]]:
    """제목 일치용 지역어: (시군구·동처럼 좁은 지역어, 도시처럼 넓은 지역어)."""
    tokens = str(region or "").split()
    sido = tokens[0] if tokens else ""
    sigungu = tokens[-1] if len(tokens) > 1 else ""
    city = next((w for w in _CITY_WORDS if w in sido), "")
    narrow = [d for d in dong_names if d and len(d) >= 2]
    if sigungu:
        narrow.append(sigungu)
        stem = re.sub(r"(구|시|군)$", "", sigungu)
        if sigungu not in _GENERIC_GU and len(stem) >= 2:
            narrow.append(stem)
    broad = [city] if city and not any(k in sido for k in ("서울", "경기", "인천")) else []
    return list(dict.fromkeys(narrow)), broad


def fetch_unsold_news(region: str, base_month: str, dong_names: List[str], limit: int = 4) -> Dict[str, Any]:
    """미분양이 있는 지역의 최근 3개월 관련 보도를 뉴스 검색 RSS에서 제목·매체·날짜만 모은다(본문은 수집하지 않음).
    제목에 미분양 관련어와 조사지역 이름이 함께 있는 기사만 남기고, 좁은 지역어(시군구·동)가 들어간 기사를 먼저 둔다."""
    import urllib.parse
    import urllib.request
    import xml.etree.ElementTree as ET
    from email.utils import parsedate_to_datetime
    out: Dict[str, Any] = {"available": False, "items": [], "query": "", "period": ""}
    try:
        end = pd.Period(str(base_month)[:7], freq="M")
        start = end - 2
        start_day = start.start_time.date()
        end_day = min(end.end_time.date(), datetime.now().date())
        narrow, broad = _unsold_region_words(region, dong_names)
        tokens = str(region or "").split()
        sigungu = tokens[-1] if len(tokens) > 1 else str(region or "")
        query_area = f"{broad[0]} {sigungu}" if broad and sigungu in _GENERIC_GU else sigungu
        stem = re.sub(r"(구|시|군)$", "", sigungu)
        short_area = broad[0] if sigungu in _GENERIC_GU and broad else (stem if len(stem) >= 2 else sigungu)
        queries = list(dict.fromkeys([f"{query_area} 미분양", f"{short_area} 미분양", f"{short_area} 할인분양"]))
        out["query"] = " / ".join(queries)
        out["period"] = f"{start_day:%Y-%m-%d} ~ {end_day:%Y-%m-%d}"
    except Exception as exc:
        out["error"] = str(exc)[:120]
        return out
    nodes = []
    for query in queries:
        q = f"{query} after:{start_day - timedelta(days=1):%Y-%m-%d} before:{end_day + timedelta(days=1):%Y-%m-%d}"
        url = "https://news.google.com/rss/search?" + urllib.parse.urlencode({"q": q, "hl": "ko", "gl": "KR", "ceid": "KR:ko"})
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                nodes.extend(ET.fromstring(resp.read()).findall("./channel/item"))
        except Exception as exc:
            out["error"] = str(exc)[:120]
    picked, seen = [], set()
    for node in nodes:
        source = (node.findtext("source") or "").strip()
        title = (node.findtext("title") or "").strip()
        if source and title.endswith(f" - {source}"):
            title = title[: -len(source) - 3].strip()
        if not title or title in seen or not re.search(r"[가-힣]", source):
            continue
        if not any(w in title for w in _UNSOLD_NEWS_WORDS):
            continue
        rank = 0 if any(w in title for w in narrow) else (1 if any(w in title for w in broad) else None)
        if rank is None:
            continue
        try:
            day = (parsedate_to_datetime(node.findtext("pubDate") or "") + timedelta(hours=9)).date()
        except Exception:
            continue
        if not start_day <= day <= end_day:
            continue
        link = (node.findtext("link") or "").strip()
        seen.add(title)
        picked.append({"rank": rank, "date": f"{day:%Y-%m-%d}", "source": source, "title": title,
                       "url": link if link.startswith(("http://", "https://")) else "",
                       "scope": "조사지역" if rank == 0 else "광역"})
    picked.sort(key=lambda x: (x["rank"], x["date"]), reverse=False)
    picked = sorted(picked[: limit], key=lambda x: x["date"], reverse=True)
    for item in picked:
        item.pop("rank", None)
    out["items"] = picked
    out["available"] = bool(picked)
    return out


_UNSOLD_THEMES = (
    (r"할인\s*분양|분양가\s*할인", "할인분양"),
    (r"악성|준공\s*후", "준공 후(악성) 미분양"),
    (r"미분양[^…]*?(증가|쌓|늘|최고|무덤|고위험|위험)|(증가|쌓|늘)[^…]*미분양", "미분양 증가 우려"),
    (r"미분양[^…]*?(감소|줄|해소|자취|털)|(감소|줄|해소|자취|털)[^…]*미분양", "미분양 감소"),
    (r"공급", "공급 지속"),
    (r"소송|논란|분쟁", "분양 분쟁"),
    (r"잔여\s*세대|선착순", "잔여세대 판매"),
)
_COMPLEX_SUFFIX = re.compile(r"(힐|파크|뷰|시티|자이|푸르지오|래미안|힐스테이트|아이파크|캐슬|더샵|중흥|호반|모아|클래스|리버|센트럴|포레|하늘채|"
                             r"스위첸|데시앙|위브|어울림|예다음|엘가|베르디움|디에트르|포레나|트레지움|에코|타운|단지|블록|스카이|팰리스|그랑|힐즈)")


def unsold_news_summary_text(news: Dict[str, Any]) -> str:
    """보도 제목만으로 만드는 자동 요약(150자 안팎). LLM 요약이 있으면 그쪽을 쓴다."""
    items = (news or {}).get("items") or []
    if not items:
        return ""
    titles = [str(it.get("title") or "") for it in items]
    themes = []
    for pattern, label in _UNSOLD_THEMES:
        if any(re.search(pattern, t) for t in titles) and label not in themes:
            themes.append(label)
    names = []
    for t in titles:
        for q in re.findall(r"[‘'\"“]([^’'\"”]{2,25})[’'\"”]", t):
            if _COMPLEX_SUFFIX.search(q) and "미분양" not in q and q not in names:
                names.append(q)
    local = sum(1 for it in items if it.get("scope") == "조사지역")
    lead = f"최근 3개월 관련 보도 {len(items)}건" + (f"(조사지역 {local}건)" if local else "(시도 단위 보도)")
    text = lead + ("은 " + "·".join(themes[:4]) + " 등을 다룸." if themes else "이 확인됨.")
    if names:
        text += " 보도에 언급된 단지는 " + ", ".join(f"'{n}'" for n in names[:3]) + "임."
    elif not local:
        text += " 조사지역 단지를 특정한 보도는 확인되지 않음."
    return text


def unsold_news_html(news: Dict[str, Any], summary: str = "") -> str:
    """미분양 해설 아래 관련 보도 목록(제목·매체·날짜, 원문 링크)."""
    items = (news or {}).get("items") or []
    if not items:
        return ""
    lis = []
    for it in items:
        title = html.escape(str(it.get("title") or ""))
        if it.get("url"):
            title = f"<a href='{html.escape(it['url'])}' target='_blank' rel='noopener'>{title}</a>"
        scope = "" if it.get("scope") == "조사지역" else " <em class='un-scope'>광역</em>"
        lis.append(f"<li><span class='un-date'>{html.escape(str(it.get('date') or ''))}</span> {title} "
                   f"<span class='un-src'>{html.escape(str(it.get('source') or ''))}</span>{scope}</li>")
    summary = str(summary or "").strip() or unsold_news_summary_text(news)
    summary_html = f"<div class='un-summary'><b>요약</b> {summary}</div>" if summary else ""
    return ("<div class='unsold-news'><b>📰 미분양 관련 보도</b> <small>(" + html.escape(str(news.get("period") or "")) +
            ", 기사 제목 기준 요약)</small>" + summary_html + "<ul>" + "".join(lis) + "</ul></div>")


def detail_and_policy_prompt_rules(data_pack: Dict[str, Any]) -> str:
    """월간·오피스텔·주간 프롬프트 공통: 입주예정물량 제외, 세부 해설 key 기준, 변곡점·정책 대조 자료."""
    _region_text = str((data_pack.get("meta", {}) or {}).get("region", ""))
    rules = "" if is_jeonnam_gwangju_region(_region_text) else """
[입주예정물량 제외]
- 로컬 아파트 입주예정물량 자료는 광주·전남 지역 전용이며 이 보고서에서는 해당 구획을 표시하지 않으므로 supply_commentary 등에서 입주예정물량을 언급하지 않습니다.
"""
    rules += f"""
[세부 해설 key 작성 기준]
- supply_type_commentary: 표 2 물건유형별 공급 파이프라인 시계열 해설 3~4문장. supply.by_type_rows로 최근 5개년 인허가·착공·준공의 연간 추이와 당해 누계, 아파트·비아파트 구성 변화를 설명하고 당해연도는 누계월 기준임을 밝힙니다.
- apt_price_band_commentary·apt_age_band_commentary·apt_area_band_commentary: 각 2~3문장. apt_trade의 price_band_rows·building_age_rows·area_band_rows를 구간 순서(낮은 구간→높은 구간)대로 읽고 거래가 많은 구간, 상승우세·하락우세 구간, 구간이 높아질수록 나타나는 경향을 설명하며 거래 5건 미만 구간은 제한적으로 해석합니다.
- apt_trade_commentary·apt_jeonse_commentary·apt_rent_commentary: 아래 [policy_inflection]의 해당 시장(apt_trade·apt_jeonse·apt_rent) 변곡점이 있으면 그 달의 변화와 같은 시기 정책(events)을 1~2문장으로 연결하되 '시점이 겹침', '정책 영향이 반영된 것으로 판단됨' 수준으로 쓰고 events에 없는 정책·수치는 만들지 않습니다.
- 월세 비교는 환산월세(월세+보증금×적용전환율÷100÷12) 기준입니다. 보증금이 크게 다른 계약을 월세 금액만으로 비교하지 않습니다.

- apt_trade_commentary·apt_jeonse_commentary·apt_rent_commentary에는 아래 [volume_reference]로 최근 12개월 월평균 거래건수가 1년 전·2년 전보다 늘었는지 줄었는지를 1문장 반영합니다(기준월은 신고기한 미도래로 제외된 값임).

[volume_reference]
{json.dumps({"apt_trade": (data_pack.get("apt_trade") or {}).get("volume_reference"), "apt_jeonse": ((data_pack.get("apt_rent") or {}).get("jeonse") or {}).get("volume_reference"), "apt_rent": ((data_pack.get("apt_rent") or {}).get("monthly_rent") or {}).get("volume_reference")}, ensure_ascii=False)}

[policy_inflection]
{json.dumps(build_policy_inflection(data_pack), ensure_ascii=False, indent=1)}

[미분양 관련 보도 - unsold_commentary 작성 기준]
- 아래 unsold_news.items가 있으면 unsold_commentary 끝에 1~2문장으로 보도 내용을 짧게 덧붙입니다. 보도 제목에 나온 단지명(따옴표 안 이름 등)·할인분양·잔여세대 같은 사실만 날짜와 함께 쓰고, 제목에 없는 세대수·할인율·원인은 추측하지 않습니다.
- scope가 "광역"인 보도는 시도 전체 분위기로만 쓰고 조사지역 단지 사례처럼 쓰지 않습니다. items가 비어 있으면 보도는 언급하지 않습니다.
- unsold_news_summary: items가 있으면 보도 흐름을 2~3문장, 150자 이내로 요약합니다(예: 어떤 보도가 몇 건, 할인분양·악성 미분양 증감 등 주제, 제목에 나온 단지명). 제목에 없는 내용은 쓰지 않고, items가 비어 있으면 빈 문자열로 둡니다.
{json.dumps((data_pack.get("unsold_news") or {}).get("items") or [], ensure_ascii=False)}
"""
    return rules


def build_weekly_llm_prompt(data_pack: Dict[str, Any]) -> str:
    """주택동향업무(주간) 전용 LLM 전달 프롬프트."""
    schema = {k: "" for k in COMMENTARY_KEYS}
    for k in [
        "apt_rent_commentary", "non_apt_sale_commentary", "non_apt_jeonse_commentary", "non_apt_rent_commentary",
        "rowhouse_multiplex_sale_summary", "detached_multifamily_sale_summary",
        "rowhouse_multiplex_jeonse_summary", "detached_multifamily_jeonse_summary",
        "rent_market_summary", "apt_rent_summary", "rowhouse_multiplex_rent_summary", "detached_multifamily_rent_summary",
    ]:
        if k in schema:
            schema[k] = ""
    region_profile = data_pack.get("region_profile", {})
    weekly_sale = data_pack.get("weekly_apt_sale_change", {}) or data_pack.get("price_index", {}).get("sale", {})
    weekly_jeonse = data_pack.get("weekly_apt_jeonse_change", {}) or data_pack.get("price_index", {}).get("jeonse", {})
    weekly_change_summary = data_pack.get("weekly_change_summary", {}) or data_pack.get("monthly_change_summary", {})
    dong_focus_summary = data_pack.get("dong_focus_summary", {})
    detail_statistics_summary = data_pack.get("detail_statistics_summary", {})
    case_focus_summary = data_pack.get("case_focus_summary", {})
    supply_pressure_summary = data_pack.get("supply_pressure_summary", {})
    sido_field_rules = _SIDO_FIELD_CONTEXT_RULES if str(data_pack.get("meta", {}).get("report_scope", "")) == "sido" else ""
    return f"""{kreb_prompt_frame(data_pack)}
이번 작업: 주간 주택시장 동향 보고서의 해설 문장을 JSON으로 작성합니다.
아래 data_pack.json은 Python으로 계산된 전체 분석자료입니다.
주간보고서는 월간보고서 양식 중 비아파트 관련 분석만 제외하고, 아파트 매매·전세를 중심으로 작성합니다.
임의 계산, 임의 추정, 임의 단지명 생성은 금지합니다.
이 보고서는 통계 해설문이 아니라 조사 담당자가 현장에서 확인한 시장 상황을 설명하는 시황 보고서입니다.
{_EVIDENCE_PRIORITY_RULES}
{_market_field_block(data_pack, True)}
{_LISTING_TRADE_RULES}
{_FIELD_TONE_RULES}
{sido_field_rules}
{detail_and_policy_prompt_rules(data_pack)}
[주간보고서 핵심 기준]
1. 주택동향업무(주간)는 매주 월요일 공표 기준일을 기준으로 작성합니다.
2. 주간 공표통계는 KOSIS 주간아파트 매매가격변동률과 주간아파트 전세가격변동률을 우선 사용합니다.
3. weekly_apt_sale_change와 weekly_apt_jeonse_change의 값은 이미 변동률이므로 지수처럼 전월 대비 재계산하지 않습니다.
4. 최근 흐름은 과거 월은 월간동향 아파트 가격지수의 전월 대비 변동률, 기준월은 매주 월요일 기준 주간 아파트 가격변동률로 구성된 weekly_apt_* table_rows를 기준으로 해석합니다.
5. 아파트 실거래 자료는 주간 공표통계를 보완하는 보조지표로만 활용합니다.
6. 인구, 세대, 이동, 미분양, 준공후미분양, 인허가, 착공, 준공, 분양, 금리 통계는 월간과 동일한 배경지표로 활용합니다.
7. 실거래 기준월 자료는 신고기한 미도래 잠정치일 수 있으므로 필요한 곳에 제한적으로 언급합니다.

[외지인 매수 해석 규칙]
1. external_buyer_sale은 업로드 주거용 매매 원장의 매수자 시군구를 Python에서 보고지역과 직접 비교해 계산한 결과입니다.
2. monthly_rows의 전체 매매건수·외지인 거래건수·외지인 비율 13개월 흐름과 current_region_rows, six_month_region_rows를 함께 확인합니다.
3. external_buyer_commentary에는 당월 비율, 전월 대비 방향, 당월 상위 유입 시군구, 최근 6개월 집중 지역을 자연스럽게 해석합니다.
4. 외지인 거래 증감만으로 가격 변동 원인을 단정하지 말고, 거래량·금리·공급·지역 생활권 수요와 구분하여 설명합니다.

[전월세 및 물건종류 판정 규칙]
1. apt_rent와 non_apt_rent의 rent_type_basis·rent_type_counts를 확인하고, 원자료 거래구분 열의 전세·월세 값을 기준으로 각각 해석합니다.
2. 거래구분 열이 있는 경우 월세액 0 여부로 전세·월세를 다시 추정하거나 재분류하지 않습니다.
3. property_type_basis·property_type_counts를 확인하고, 주용도를 우선하되 값이 없을 때 건물유형으로 보완하여 아파트·연립·다세대·단독·다가구를 구분합니다.
4. 물건종류별 거래가 없거나 적으면 방향성을 단정하지 않고 거래가 없거나 제한적이라고 표현합니다.

[작성 원칙]
1. 모든 수치·비율·순위·단지명은 data_pack.json 값만 사용합니다.
2. HTML 태그, 마크다운 표, 설명 문구는 작성하지 않습니다.
3. 출력은 반드시 JSON 객체 하나만 작성합니다.
4. 공공기관 보고서 문체로 작성합니다.
5. 문장 끝은 위 [현장 시황 문체 강화]를 따릅니다. “판단됨”, “분석됨”, “해석됨”만 반복하지 말고 “이어짐”, “나타남”, “유지됨”, “확대됨”, “축소됨”, “전환됨”, “~하는 양상임”, “~로 조사됨”을 섞어 씁니다.
6. 과도한 예측은 피하고, 기준주 및 확인 가능한 최근 흐름 범위 내에서 해석합니다.
7. 금지 문구: “신고가 기준”
8. 전세는 갱신계약 제외 신규계약 기준으로 해석합니다.
9. 수치가 null, 0, 빈 목록이면 억지 해석하지 말고 “원자료 확인 필요” 또는 “거래가 제한적”이라고 표현합니다.
10. 표 전체를 반복하지 말고, 보고서 해석에 필요한 핵심 변화와 근거 수치만 선별합니다.

[Ⅵ. 종합 시황 서술 방식]
1. Ⅵ장은 수치 나열형 문장이 아니라 지역 조사 담당자가 시장 분위기를 설명하는 자연스러운 문단으로 작성합니다.
2. 지역특성, 배후수요, 교통 접근성, 상권, 학군, 원도심·신도심 생활권 차이를 먼저 떠올리되, data_pack 또는 region_profile에 근거가 있을 때만 연결합니다.
3. 인구·세대수·유입유출은 “수요 기반이 늘어남/줄어듦/정체됨”, “세대 분화가 임차수요를 지지함”처럼 방향성 언어로 해석하고 숫자는 꼭 필요한 경우만 보조로 씁니다.
4. 최근 13개월 거래와 공표통계 시계열은 상승, 하락, 감소, 증가, 보합, 혼조, 관망, 선별 거래 같은 시장 언어로 풀어 씁니다.
5. 상승·하락 비율은 비율 자체를 반복하기보다 어느 쪽 압력이 우세한지, 거래가 적어 판단을 유보해야 하는지를 설명합니다.
6. (오피스텔동향보고서에서만 적용) 입력 데이터가 오피스텔(meta.is_officetel == true) 일 때만 업무·상업 배후수요, 소형 임대수요, 직주근접, 월세 전환 수요 중심 해석을 사용합니다. 아파트·비아파트(연립·다세대·단독·다가구) 월간/주간 보고서에서는 "오피스텔", "비오피스텔" 이라는 단어를 절대 사용하지 않으며 "아파트", "비아파트" 로 일관되게 표기합니다.
7. 각 종합 문단은 “지역 맥락 → 수요·공급 여건 → 13개월 흐름 → 당월/금주 방향 → 판단” 순서로 작성합니다.

[반도체 첨단국가산단 정책 맥락 규칙]
1. policy_context.semiconductor_cluster.available이 true이면 2026년 7월 9일 공고, 7월 14일 효력 발생 사실을 종합서술에 반영합니다.
2. core_dongs·adjacent_dongs·linked_dongs는 공식 수혜지역이 아니라 보고서용 분석 영향권이므로 등급을 구분해 표현합니다.
3. 토지거래허가구역 규제효과와 산업단지 조성 기대효과를 구분하고, 거래량·가격 변화가 data_pack에서 확인되지 않으면 인과관계를 단정하지 않습니다.
4. 매매와 전세 모두 해당 읍면동의 실제 13개월 및 전월 대비 흐름을 확인한 경우에만 정책 맥락과 제한적으로 연결합니다.
5. 정책 맥락을 별도 문단으로 떼어 쓰지 말고, region_profile의 market_characteristics와 해당 영향권 읍면동의 dong_profiles 설명(광역교통 결절, 산업 배후수요, 원도심 주거 특성 등)과 결합해 종합서술 안에 녹여 씁니다.
6. 군공항 부지·국가산단 관련 사실은 policy_context.semiconductor_cluster에 명시된 공고·효력 사실만 사용하고, 이전 시기·보상·착공·수용 일정 등 확인되지 않은 절차나 기대 효과는 서술하지 않습니다.

[아파트 입주예정물량 활용 규칙]
1. move_in_supply가 available이면 선택 시군구의 연간 총물량, 단지수, 월별 집중시기와 주요 단지를 공급·전세 종합서술에 반영합니다.
2. 입주예정물량은 일반공급 규모가 아니라 전체 세대수를 우선한 수치이며, 준공실적과 구분합니다.
3. 입주 일정은 변경될 수 있으므로 확정 입주물량으로 단정하지 말고 “입주 예정”, “수급 점검 필요”로 표현합니다.
4. 시군구별 물량 차이를 해당 지역의 기존 주택 거래·전세 흐름과 연결하되, 가격 영향의 인과관계는 단정하지 않습니다.

[지역특성 기반 종합서술 강화 규칙 - 매우 중요]
1. region_profile은 해당 지역의 고정 배경자료이며, market_characteristics(핵심 시장 특성), region_summary, dong_profiles와 함께 supply_context(공급여건), demand_drivers(수요 전환·배후수요), transport(교통), commercial(상권), education(학군) 필드를 반드시 확인합니다.
2. region_profile.market_characteristics가 있으면 이를 종합서술의 중심 축으로 삼아, 각 핵심 특성을 금주/당월 매매·전세 흐름·거래사례와 연결해 풀어 씁니다.
3. 종합서술은 수치 나열로 시작하지 말고, 먼저 지역 맥락(핵심 시장 특성, 생활권 구조, 공급여건, 수요 전환·유입유출, 교통·상권·학군 배후수요)을 설명한 뒤, 그 배경 위에서 금주 공표통계·실거래 수치를 근거로 붙입니다.
3. supply_context에 공급부족·신규공급 차질 근거가 있으면 매매·전세 가격 흐름 해석에 수급 배경으로 연결합니다.
4. demand_drivers에 산업·업무 배후수요, 광역교통 결절 기반 통근수요, 임차→매매 전환, 인접 지역으로의 수요 유출 등이 기재되어 있으면, 거래 집중 읍면동·면적·연식대 흐름과 연결해 “왜 그런 수요가 나타나는지”를 1~2문장 보강합니다.
5. region_profile에 명시된 사실(market_characteristics·supply_context·policy_context_note 등)은 적극 활용하되, region_profile·data_pack에 없는 외부 개발 호재·정비사업·단지명·학교명·노선명은 임의로 생성하지 않습니다. 다른 시군구의 사례를 현재 조사지역 서술에 옮겨 쓰지 않습니다.
6. 교통·상권·학군은 거래가 확인된 읍면동과 연결될 때 자연스럽게 녹여 쓰고, 근거가 약하면 단정하지 않습니다.
7. 지역특성 문장을 그대로 복사하지 말고, 금주/당월 수치·거래사례와 결합해 새 문장으로 재구성합니다.

{_DONG_FIELD_CAUSE_RULES}
{field_style_summary_rules(data_pack)}
[읍면동 단위 생활권 서술 규칙 - apt_sale_summary·apt_jeonse_summary 전용, 매우 중요]
1. dong_focus_summary의 apt_sale_top5·apt_jeonse_top5에서 거래 상위 2~3개 읍면동을 고르고, 각 읍면동을 region_profile.dong_profiles의 설명(생활권 성격, 배후수요, 교통, 학군, 신축·구축 구성)과 하나씩 짝지어 서술합니다.
2. 각 읍면동은 “몇 건이 거래되었다”가 아니라 “그 생활권에서 어떤 수요가 왜 움직였는지”를 중심으로 씁니다. 거래건수·비율·금액은 판단을 뒷받침하는 근거로 문장당 최대 1회만 붙입니다.
3. 상위 읍면동 간 성격 차이(신도시형 대단지 vs 원도심 구축, 산업·업무 배후 vs 학군·생활편의 선호, 광역교통 결절 vs 외곽 저밀)를 최소 1문장 이상 대비해 설명합니다.
4. 단지명·금액·면적·층을 연속으로 나열하지 않습니다. 대표 거래사례는 최대 2건까지만 문장에 자연스럽게 녹여 쓰고 나머지는 표·차트에 맡깁니다. 보고서 HTML은 각 문단 아래에 대표 거래사례 줄을 자동으로 출력하므로, 서술에서 사례를 반복 나열하면 같은 내용이 중복됩니다.
5. detail_statistics_summary의 면적구간·가격구간·건축연령 통계는 “어떤 수요층이 어느 생활권에서 움직였는지”로 해석해 쓰고, 구간별 수치를 표처럼 옮겨 적지 않습니다.
6. dong_profiles에 없는 읍면동은 지역특성을 지어내지 말고 거래 흐름만 사실대로 쓰되, 그 읍면동이 상위권에 새로 올라온 사실 자체를 관찰 포인트로 제시합니다.
7. 5~7문장 중 최소 3문장 이상은 수치가 아니라 생활권·수요·배경 해석 문장이어야 하며, 숫자로 시작하는 문장이 연속 3개를 넘지 않게 합니다.
8. policy_context.semiconductor_cluster.available이 true이면 핵심·인접·연계 영향권에 해당하는 읍면동을 위 생활권 서술 안에서 함께 다루고, 별도 정책 문단으로 떼어 쓰지 않습니다.

[비아파트 제외 규칙 - 매우 중요]
1. 주간보고서에서는 연립·다세대, 단독·다가구, 비아파트 관련 표·차트·서술을 작성하지 않습니다.
2. non_apt_sale, non_apt_rent, non_apt_interpretation_summary는 검토 대상에서 제외합니다.
3. rowhouse_multiplex_sale_summary, detached_multifamily_sale_summary, rowhouse_multiplex_jeonse_summary, detached_multifamily_jeonse_summary, rowhouse_multiplex_rent_summary, detached_multifamily_rent_summary는 빈 문자열로 둡니다.
4. 월세는 주간 공표통계 대상이 아니므로 apt_rent_commentary, rent_market_summary, apt_rent_summary는 빈 문자열로 둡니다.
5. “비아파트”, “연립·다세대”, “단독·다가구”, “월세시장”이라는 표현을 주간 종합서술에 넣지 않습니다.

[주간 시계열 활용 규칙]
1. weekly_apt_sale_change.table_rows와 weekly_apt_jeonse_change.table_rows를 우선 확인합니다.
2. table_rows의 source가 월간동향 공표통계인지 주간동향 공표통계인지 구분해 작성합니다.
3. 최신 공표주, 전주 대비 방향, 최근 13개월 월간 흐름과 기준월 주간 흐름의 상승·하락·보합 여부를 구분해 작성합니다.
4. 주간 공표통계와 실거래 보조지표가 다르면 “공표통계는 ○○이나, 실거래 보조지표는 ○○로 확인됨”처럼 구분합니다.
5. 특정 주만 튀는 경우 원인을 단정하지 말고 “일시적 변동 가능성”으로 표현합니다.

[아파트 실거래 보조지표 활용 규칙]
1. apt_trade와 apt_rent 중 전세 관련 자료를 확인합니다.
2. 아파트 매매는 거래건수, 비교가능건수, 상승·보합·하락 비율, 평균 변동률, 주요 거래사례를 보조적으로 활용합니다.
3. 아파트 전세는 신규계약 기준으로 거래건수, 비교가능건수, 상승·보합·하락 비율, 보증금 변화, 주요 거래사례를 활용합니다.
4. 단지명은 반드시 data_pack에 있는 값만 사용하며, 아파트명은 가능하면 “아파트명 (사용승인년도, 세대수)” 형식을 따릅니다.

[수요·공급·금리 배경지표 활용 규칙]
1. population, households, migration은 수요 기반의 유지·확대·위축 여부를 판단하는 보조근거로 활용합니다.
2. unsold, post_completion_unsold는 매매시장 부담 요인으로 연결합니다.
3. housing_permit, housing_start, housing_completion, housing_sale_supply는 향후 공급압력 및 전세시장 수급 판단에 연결합니다.
4. rates는 매수심리, 전세자금 부담, 거래 관망세 해석에 활용합니다.
5. 배경지표는 주간마다 급변하지 않을 수 있으므로, 가격동향을 단정하는 직접 원인으로 과도하게 쓰지 않습니다.

[종합 판단 권장 문장 구조]
① 금주 공표통계상 아파트 매매·전세 방향
② 최근 13주 흐름에서 유지·확대·완화된 점
③ 실거래 보조지표상 상승·하락 비율 또는 주요 단지 사례
④ 인구·세대·공급·미분양·금리 등 배경 여건
⑤ 다음 주 조사상 유의사항

[LLM 필수 재작성 대상]
- Ⅵ 종합 시황 서술 보고서에 들어가는 아래 key는 기본 문장을 그대로 복사하지 말고 LLM이 data_pack을 읽어 새로 작성합니다.
- 필수 key: comprehensive_summary, sale_market_summary, apt_sale_summary, jeonse_market_summary, apt_jeonse_summary
- 주간보고서는 비아파트와 월세 관련 key를 빈 문자열로 두되, 매매와 전세 key는 지역특성·배후수요·교통·상권·학군·인구·세대·이동·공표통계 흐름을 자연스럽게 연결합니다.
- 위 key가 비어 있으면 최종 HTML 검수에서 LLM 미반영으로 판정됩니다.

[출력 JSON 구조]
아래 key를 반드시 모두 포함하세요. JSON 외 문장은 쓰지 마세요.
비아파트 및 월세 제외 key는 반드시 빈 문자열로 작성하세요.
{json.dumps(schema, ensure_ascii=False, indent=2)}

[Ⅵ. 종합 시황 서술 보고서 작성 기준 - 주간형]
- comprehensive_summary: 종합 6~8문장. 첫 문장은 수치가 아니라 region_profile에 기재된 해당 조사지역의 구조적 배경(핵심 시장 특성·공급여건·수요 기반·생활권 구조)으로 시작하고, 그 위에서 금주 매매·전세 방향을 연결. 다른 지역의 사례나 배경을 끌어와 쓰지 않음.
- sale_market_summary: 아파트 매매 중심 6~8문장. region_profile의 market_characteristics·supply_context·demand_drivers를 먼저 해당 지역의 배경으로 풀고, 주간 공표통계·실거래 보조지표는 그 배경을 뒷받침하는 근거로 붙임. policy_context.semiconductor_cluster.available이 true이면 확인된 정책 사실과 영향권 등급을 이 배경 서술에 함께 반영.
- apt_sale_summary: 아파트 매매 5~7문장. 앞의 [읍면동 단위 생활권 서술 규칙]을 그대로 적용. 첫 2~3문장은 거래 상위 읍면동의 생활권 성격(dong_profiles)과 배후수요·교통·학군·신축구축 구성으로 시작하고, 그 위에 금주 매매 방향과 상승·하락 우세를 얹음. 이어서 어느 생활권에서 어떤 면적대·연식대 수요가 움직였는지를 해석하고, 대표 거래사례는 최대 2건만 근거로 사용. 단지명·금액을 나열하는 문장은 쓰지 않음.
- jeonse_market_summary: 아파트 전세 중심 6~8문장. region_profile의 supply_context 등 전세 수급 배경을 먼저 설명하고, 주간 공표통계·신규 전세계약 보조지표·보증금 부담을 근거로 연결.
- apt_jeonse_summary: 아파트 전세 5~7문장. 앞의 [읍면동 단위 생활권 서술 규칙]을 그대로 적용. 거래 상위 읍면동의 생활권 성격, 학군·직주근접·산업 배후수요, 구축 비중과 임차수요 성격을 먼저 설명하고, 그 배경 위에 신규계약 기준 보증금 흐름·면적대별 수요 차이·금리 및 입주 여건을 근거로 연결. 보증금 사례를 나열하지 말고 임차수요가 어느 생활권으로 이동했는지를 중심으로 쓰며, 대표 사례는 최대 2건만 사용.
- monthly_core_diagnosis: 주간 핵심 진단 3~5문장. 명칭은 key 유지하되 내용은 주간 기준으로 작성.
- population_commentary, age_migration_commentary, rate_commentary, unsold_commentary, supply_commentary는 월간과 동일한 배경 분석으로 작성.
- price_index_commentary는 주간아파트 매매·전세 공표통계를 중심으로 작성.
- conversion_rate_commentary, apt_rent_commentary, non_apt_sale_commentary, non_apt_jeonse_commentary, non_apt_rent_commentary, rent_market_summary, apt_rent_summary 및 비아파트 summary key는 빈 문자열로 작성.

[region_profile]
{json.dumps(region_profile, ensure_ascii=False, indent=2)}

[weekly_apt_sale_change]
{json.dumps(weekly_sale, ensure_ascii=False, indent=2)}

[weekly_apt_jeonse_change]
{json.dumps(weekly_jeonse, ensure_ascii=False, indent=2)}

[weekly_change_summary]
{json.dumps(weekly_change_summary, ensure_ascii=False, indent=2)}

[dong_focus_summary]
{json.dumps(dong_focus_summary, ensure_ascii=False, indent=2)}

[detail_statistics_summary]
{json.dumps(detail_statistics_summary, ensure_ascii=False, indent=2)}

[case_focus_summary]
{json.dumps(case_focus_summary, ensure_ascii=False, indent=2)}

[supply_pressure_summary]
{json.dumps(supply_pressure_summary, ensure_ascii=False, indent=2)}

[HTML 디자인·가독성 참고]
- HTML 본문은 A4 PDF 출력 기준으로 읽히므로, 한 문단이 과도하게 길어지지 않게 작성합니다.
- 아파트 매매·전세 세부 문단은 5~7문장으로 작성하되, 긴 문장 하나가 아니라 80~120자 안팎의 짧은 문장을 여러 개 연결합니다.
- 수치 나열보다 “핵심 변화 → 원인 → 판단” 순서로 정리합니다.
- 표·차트가 이미 제공하는 단순 수치는 반복하지 말고, 표·차트에서 드러나는 의미를 서술합니다.
- 매매와 전세 박스는 서로 다른 논지를 가져야 하며, 같은 문장 구조 반복을 피합니다.

{_llm_data_pack_block(data_pack, drop_keys=_LLM_SEPARATE_BLOCK_KEYS)}

{_KREB_SELF_CHECK}"""

def build_llm_prompt(data_pack: Dict[str, Any]) -> str:
    if is_weekly_report(data_pack):
        return build_weekly_llm_prompt(data_pack)
    schema = {k: "" for k in COMMENTARY_KEYS}
    region_profile = data_pack.get("region_profile", {})
    monthly_change_summary = data_pack.get("monthly_change_summary", {})
    dong_focus_summary = data_pack.get("dong_focus_summary", {})
    detail_statistics_summary = data_pack.get("detail_statistics_summary", {})
    case_focus_summary = data_pack.get("case_focus_summary", {})
    non_apt_interpretation_summary = data_pack.get("non_apt_interpretation_summary", {})
    supply_pressure_summary = data_pack.get("supply_pressure_summary", {})
    is_officetel = str(data_pack.get("meta", {}).get("report_mode", "")) == "officetel" or bool(data_pack.get("meta", {}).get("is_officetel"))
    is_sido_scope = str(data_pack.get("meta", {}).get("report_scope", "")) == "sido"
    scope_rules = """
[시도단위 보고서 작성 규칙 - 매우 중요]
1. meta.report_scope가 "sido"이므로 이 보고서는 선택 시도 내 전체 시군구를 취합한 시도단위 보고서입니다.
2. 특정 시군구 하나의 동향처럼 쓰지 말고, 시도 전체 흐름과 시군구·읍면동별 편차를 함께 해석합니다.
3. 실거래 파일은 선택 시도 내 시군구별 LAWD_CD 자료를 합산한 것이므로 거래건수·상승/하락 비율·TOP 사례를 시도 전체 취합값으로 해석합니다.
4. 인구·세대·이동·미분양·가격지수·전환율·공급통계는 시도 기준 또는 시도 내 시군구 취합 기준으로 설명합니다.
5. 문장에는 "해당 시도 전체", "시도 내 주요 시군구", "시군구별 차이"처럼 범위가 드러나게 쓰고, 단일 시군구 보고서처럼 단정하지 않습니다.
""" if is_sido_scope else """
[시군구단위 보고서 작성 규칙]
1. meta.report_scope가 "sigungu"이므로 이 보고서는 선택 시군구 기준 보고서입니다.
2. 공급통계처럼 시도 기준으로 제공되는 항목은 별도 시도 공급권역으로 해석하고, 실거래·인구·세대·이동은 선택 시군구 기준으로 해석합니다.
"""
    scope_rules += detail_and_policy_prompt_rules(data_pack)
    officetel_exclusion_rules = """
[오피스텔 보고서 전용 제외 규칙 - 매우 중요]
1. 오피스텔동향보고서에서는 비아파트, 연립·다세대, 단독·다가구, 분양권전매 관련 표·차트·서술을 전부 작성하지 않습니다.
2. non_apt_sale, non_apt_rent, non_apt_interpretation_summary, presale는 검토 대상에서 제외합니다.
3. non_apt_sale_commentary, non_apt_jeonse_commentary, non_apt_rent_commentary는 빈 문자열로 둡니다.
4. non_apt_sale_regional_context, rowhouse_multiplex_sale_summary, detached_multifamily_sale_summary, non_apt_jeonse_regional_context, rowhouse_multiplex_jeonse_summary, detached_multifamily_jeonse_summary, non_apt_rent_regional_context, rowhouse_multiplex_rent_summary, detached_multifamily_rent_summary는 빈 문자열로 둡니다.
5. presale_commentary도 빈 문자열로 둡니다.
6. 오피스텔 보고서 본문에는 “비아파트”, “연립·다세대”, “단독·다가구”, “분양권전매”라는 표현을 넣지 않습니다.
7. sale_market_summary, apt_sale_summary, jeonse_market_summary, apt_jeonse_summary, rent_market_summary, apt_rent_summary는 이름은 기존 key를 유지하되 내용은 오피스텔 매매·전세·월세로 작성합니다.
""" if is_officetel else ""
    officetel_field_rules = _OFFICETEL_FIELD_RULES if is_officetel else ""
    monthly_housing_field_rules = "" if is_officetel else _MONTHLY_HOUSING_FIELD_RULES
    weekly_rollup_rules = _WEEKLY_ROLLUP_RULES if (data_pack.get("market_field_context", {}) or {}).get("weekly_rollup") else ""
    sido_field_rules = _SIDO_FIELD_CONTEXT_RULES if is_sido_scope else ""
    return f"""{kreb_prompt_frame(data_pack)}
이번 작업: {"오피스텔" if is_officetel else "월간 주택"}시장 동향 보고서의 해설 문장을 JSON으로 작성합니다.
아래 data_pack.json은 Python으로 계산된 전체 분석자료입니다.
요약 블록만 보지 말고 data_pack.json 전체 구조를 확인한 뒤, 실제 산출값이 있는 항목만 서술에 반영합니다.
임의 계산, 임의 추정, 임의 단지명 생성은 금지합니다.
이 보고서는 통계 해설문이 아니라 조사 담당자가 현장에서 확인한 시장 상황을 설명하는 시황 보고서입니다.
{_EVIDENCE_PRIORITY_RULES}
{_market_field_block(data_pack, False)}
{_LISTING_TRADE_RULES}
{_FIELD_TONE_RULES}
{monthly_housing_field_rules}
{weekly_rollup_rules}
{officetel_field_rules}
{sido_field_rules}
[작성 원칙]
1. 모든 수치·비율·순위·단지명은 data_pack.json 값만 사용합니다.
2. HTML 태그, 마크다운 표, 설명 문구는 작성하지 않습니다.
3. 출력은 반드시 JSON 객체 하나만 작성합니다.
4. 공공기관 보고서 문체로 작성합니다.
5. 문장 끝은 위 [현장 시황 문체 강화]를 따릅니다. “판단됨”, “분석됨”, “해석됨”만 반복하지 말고 “이어짐”, “나타남”, “유지됨”, “확대됨”, “축소됨”, “전환됨”, “~하는 양상임”, “~로 조사됨”을 섞어 씁니다.
6. 과도한 예측은 피하고, 기준월 자료로 확인되는 범위 내에서 해석합니다.
7. 금지 문구: “신고가 기준”
8. 전세는 갱신계약 제외 신규계약 기준으로 해석합니다.
9. 월세는 환산월세 기준으로 해석합니다.
10. 기준월 실거래는 신고기한 미도래 잠정치임을 필요한 곳에 반영합니다.
11. 비아파트와 분양권전매는 거래규모가 제한적이면 전체 시장 방향을 단정하지 말고 보조지표로 해석합니다.
12. 수치가 null, 0, 빈 목록이면 억지 해석하지 말고 “원자료 확인 필요” 또는 “거래가 제한적”이라고 표현합니다.
13. data_pack.json은 전체 분석자료이므로 population, age_population, migration, rates, unsold, conversion_rate, price_index, supply, apt_trade, apt_rent, non_apt_sale, non_apt_rent, presale를 모두 검토합니다.
14. 문단별로 관련 자료가 있으면 summary 블록보다 원자료 세부표(table_rows, monthly_rows, dong_rows, area_band_rows, price_band_rows, case_rows 등)를 우선 근거로 사용합니다.
15. 다만 표 전체를 그대로 반복하지 말고, 보고서 해석에 필요한 핵심 변화와 근거 수치만 선별합니다.
{scope_rules}

[외지인 매수 해석 규칙]
1. external_buyer_sale은 업로드 주거용 매매 원장의 매수자 시군구를 Python에서 보고지역과 직접 비교해 계산한 결과입니다.
2. monthly_rows의 전체 매매건수·외지인 거래건수·외지인 비율 13개월 흐름과 current_region_rows, six_month_region_rows를 함께 확인합니다.
3. external_buyer_commentary에는 당월 비율, 전월 대비 방향, 당월 상위 유입 시군구, 최근 6개월 집중 지역을 자연스럽게 해석합니다.
4. 외지인 거래 증감만으로 가격 변동 원인을 단정하지 말고, 거래량·금리·공급·지역 생활권 수요와 구분하여 설명합니다.

[전월세 및 물건종류 판정 규칙]
1. apt_rent와 non_apt_rent의 rent_type_basis·rent_type_counts를 확인하고, 원자료 거래구분 열의 전세·월세 값을 기준으로 각각 해석합니다.
2. 거래구분 열이 있는 경우 월세액 0 여부로 전세·월세를 다시 추정하거나 재분류하지 않습니다.
3. property_type_basis·property_type_counts를 확인하고, 주용도를 우선하되 값이 없을 때 건물유형으로 보완하여 아파트·연립·다세대·단독·다가구를 구분합니다.
4. 물건종류별 거래가 없거나 적으면 방향성을 단정하지 않고 거래가 없거나 제한적이라고 표현합니다.

[Ⅵ. 종합 시황 서술 방식]
1. Ⅵ장은 수치 나열형 문장이 아니라 지역 조사 담당자가 시장 분위기를 설명하는 자연스러운 문단으로 작성합니다.
2. 지역특성, 배후수요, 교통 접근성, 상권, 학군, 원도심·신도심 생활권 차이를 먼저 떠올리되, data_pack 또는 region_profile에 근거가 있을 때만 연결합니다.
3. 인구·세대수·유입유출은 “수요 기반이 늘어남/줄어듦/정체됨”, “세대 분화가 임차수요를 지지함”처럼 방향성 언어로 해석하고 숫자는 꼭 필요한 경우만 보조로 씁니다.
4. 최근 13개월 거래와 공표통계 시계열은 상승, 하락, 감소, 증가, 보합, 혼조, 관망, 선별 거래 같은 시장 언어로 풀어 씁니다.
5. 상승·하락 비율은 비율 자체를 반복하기보다 어느 쪽 압력이 우세한지, 거래가 적어 판단을 유보해야 하는지를 설명합니다.
6. (오피스텔동향보고서에서만 적용) 입력 데이터가 오피스텔(meta.is_officetel == true) 일 때만 업무·상업 배후수요, 소형 임대수요, 직주근접, 월세 전환 수요 중심 해석을 사용합니다. 아파트·비아파트(연립·다세대·단독·다가구) 월간/주간 보고서에서는 "오피스텔", "비오피스텔" 이라는 단어를 절대 사용하지 않으며 "아파트", "비아파트" 로 일관되게 표기합니다.
7. 각 종합 문단은 “지역 맥락 → 수요·공급 여건 → 13개월 흐름 → 당월 방향 → 판단” 순서로 작성합니다.
{officetel_exclusion_rules}

[반도체 첨단국가산단 정책 맥락 규칙]
1. policy_context.semiconductor_cluster.available이 true이면 2026년 7월 9일 공고, 7월 14일 효력 발생 사실을 종합서술에 반영합니다.
2. core_dongs·adjacent_dongs·linked_dongs는 공식 수혜지역이 아니라 보고서용 분석 영향권이므로 등급을 구분해 표현합니다.
3. 토지거래허가구역 규제효과와 산업단지 조성 기대효과를 구분하고, 거래량·가격 변화가 data_pack에서 확인되지 않으면 인과관계를 단정하지 않습니다.
4. 매매와 전세 모두 해당 읍면동의 실제 13개월 및 전월 대비 흐름을 확인한 경우에만 정책 맥락과 제한적으로 연결합니다.
5. 정책 맥락을 별도 문단으로 떼어 쓰지 말고, region_profile의 market_characteristics와 해당 영향권 읍면동의 dong_profiles 설명(광역교통 결절, 산업 배후수요, 원도심 주거 특성 등)과 결합해 종합서술 안에 녹여 씁니다.
6. 군공항 부지·국가산단 관련 사실은 policy_context.semiconductor_cluster에 명시된 공고·효력 사실만 사용하고, 이전 시기·보상·착공·수용 일정 등 확인되지 않은 절차나 기대 효과는 서술하지 않습니다.

[아파트 입주예정물량 활용 규칙]
1. move_in_supply가 available이면 선택 시군구의 연간 총물량, 단지수, 월별 집중시기와 주요 단지를 공급·전세 종합서술에 반영합니다.
2. 입주예정물량은 일반공급 규모가 아니라 전체 세대수를 우선한 수치이며, 준공실적과 구분합니다.
3. 입주 일정은 변경될 수 있으므로 확정 입주물량으로 단정하지 말고 “입주 예정”, “수급 점검 필요”로 표현합니다.
4. 시군구별 물량 차이를 해당 지역의 기존 주택 거래·전세 흐름과 연결하되, 가격 영향의 인과관계는 단정하지 않습니다.

[지역특성 기반 서술 강화 규칙]
1. region_profile은 고정 배경자료이며, 문장을 그대로 반복하지 않습니다.
1-1. region_profile에 market_characteristics(핵심 시장 특성)·supply_context·demand_drivers가 있으면 종합서술의 중심 축으로 삼아, 공급여건·수요 전환(임차→매매, 구축→신축 상향이동)·인접 신도시 수요 유출 등을 당월 거래흐름과 연결해 풀어 씁니다.
2. 지역 특성은 monthly_change_summary 또는 data_pack에서 당월 변화가 확인된 읍면동·단지·유형에 한해 연결합니다.
3. 읍면동별 거래 비중 상위 1~3개 지역을 우선 확인하되, 해당 읍면동이 region_profile.dong_profiles에 있을 때만 지역특성과 연결합니다.
4. 아파트는 학군·생활편의·신축·구축·단지규모·실거주 선호를 중심으로 해석합니다.
5. 연립·다세대는 구도심 저가 실수요, 노후 주거지, 소규모 임대수요를 중심으로 해석하되, 거래건수가 적으면 가격 방향성을 단정하지 않습니다.
6. 단독·다가구는 대지면적, 노후 주택, 임대수익형 수요, 정비 기대, 외곽·구도심 특성을 중심으로 해석합니다.
7. 전세와 월세는 실거주 수요, 금리 부담, 전세보증금 부담, 소형 임대수요, 갱신계약 비중을 함께 고려합니다.

{_DONG_FIELD_CAUSE_RULES}
{field_style_summary_rules(data_pack)}
[읍면동 단위 생활권 서술 규칙 - apt_sale_summary·apt_jeonse_summary·apt_rent_summary 및 비아파트 문단 적용, 매우 중요]
1. dong_focus_summary의 apt_sale_top5·apt_jeonse_top5·apt_rent_top5(비아파트는 non_apt_*_top5)에서 거래 상위 2~3개 읍면동을 고르고, 각 읍면동을 region_profile.dong_profiles의 설명(생활권 성격, 배후수요, 교통, 학군, 신축·구축 구성)과 하나씩 짝지어 서술합니다.
2. 각 읍면동은 “몇 건이 거래되었다”가 아니라 “그 생활권에서 어떤 수요가 왜 움직였는지”를 중심으로 씁니다. 거래건수·비율·금액은 판단을 뒷받침하는 근거로 문장당 최대 1회만 붙입니다.
3. 상위 읍면동 간 성격 차이(신도시형 대단지 vs 원도심 구축, 산업·업무 배후 vs 학군·생활편의 선호, 광역교통 결절 vs 외곽 저밀)를 최소 1문장 이상 대비해 설명합니다.
4. 단지명·금액·면적·층을 연속으로 나열하지 않습니다. 대표 거래사례는 최대 2건까지만 문장에 자연스럽게 녹여 쓰고 나머지는 표·차트에 맡깁니다. 보고서 HTML은 각 문단 아래에 대표 거래사례 줄을 자동으로 출력하므로, 서술에서 사례를 반복 나열하면 같은 내용이 중복됩니다.
5. detail_statistics_summary의 면적구간·가격구간·보증금구간·월세구간·건축연령 통계는 “어떤 수요층이 어느 생활권에서 움직였는지”로 해석해 쓰고, 구간별 수치를 표처럼 옮겨 적지 않습니다.
6. dong_profiles에 없는 읍면동은 지역특성을 지어내지 말고 거래 흐름만 사실대로 쓰되, 그 읍면동이 상위권에 새로 올라온 사실 자체를 관찰 포인트로 제시합니다.
7. 5~7문장 분량 문단은 최소 3문장 이상, 3~5문장 분량 문단은 최소 2문장 이상이 수치가 아니라 생활권·수요·배경 해석 문장이어야 하며, 숫자로 시작하는 문장이 연속 3개를 넘지 않게 합니다.
8. 읍면동 서술은 전월 반복이 되지 않도록 monthly_change_summary에서 확인되는 “당월 새로 부각된 읍면동” 또는 “상위권에서 밀려난 읍면동”을 최소 1곳 짚어 줍니다.
9. policy_context.semiconductor_cluster.available이 true이면 핵심·인접·연계 영향권에 해당하는 읍면동을 위 생활권 서술 안에서 함께 다루고, 별도 정책 문단으로 떼어 쓰지 않습니다.

[매월 반복 방지 규칙]
1. 전월과 동일한 지역 설명을 반복하지 않습니다.
2. 모든 종합 문단에는 monthly_change_summary의 전월 대비 달라진 점, 새롭게 부각된 읍면동, 주요 거래사례 변화, 상승·하락 비율 변화, 공급·미분양·금리 변화 중 하나 이상을 반영합니다.
3. 전월과 유사한 흐름이 유지된 경우에는 반복 설명 대신 “유지”, “이어짐”, “큰 변화 없음”, “전월과 유사”, “확대”, “완화”, “축소” 등 변화 표현을 사용합니다.
4. region_profile의 문장을 그대로 복사하지 말고, 당월 수치 또는 거래사례와 연결하여 새 문장으로 재구성합니다.
5. 당월 데이터에서 근거가 약한 지역 특성은 서술하지 않습니다.
6. 종합 판단 문단에는 반드시 “당월 달라진 점” 또는 “유지되지만 추가 점검이 필요한 점”을 1개 이상 포함합니다.

[13개월 시계열 활용 규칙]
1. data_pack.json의 table_rows, monthly_rows, monthly_by_type_rows는 최근 13개월 시계열 자료입니다.
2. LLM은 최신월 수치만 보지 말고 최근 13개월 흐름에서 증가·감소·정체·반등·둔화 여부를 확인해 서술합니다.
3. 인구·세대, 연령, 이동, 금리, 미분양, 가격지수, 아파트 매매·전세·월세, 비아파트 매매·전세·월세는 13개월 추세와 당월 변화를 함께 해석합니다.
4. 기준월 거래가 신고기한 미도래 잠정치로 제외된 경우 analysis_base_month와 partial_month_note를 우선 확인합니다.
5. 13개월 흐름과 당월 변화가 서로 다른 경우 “최근 흐름은 ○○이나, 당월은 ○○로 확인됨”처럼 구분하여 작성합니다.
6. 13개월 시계열에서 특정 월만 튀는 경우 원인을 단정하지 말고 “일시적 변동 가능성”으로 표현합니다.

[당월 세부 통계 활용 규칙]
1. detail_statistics_summary는 LLM 서술 보강용 당월 세부 통계입니다.
2. 읍면동별, 연령대별, 면적구간별, 매매가격구간별, 보증금구간별, 월세·환산월세구간별, 건축연령별 통계를 필요한 문단에 반영합니다.
3. 아파트 매매는 apt_sale의 dong_rows_top10, area_band_rows, price_band_rows, building_age_rows를 우선 확인합니다.
4. 아파트 전세·월세는 apt_jeonse와 apt_rent의 dong_rows_top10, area_band_rows, deposit_band_rows, monthly_band_rows, converted_rent_band_rows, building_age_rows를 확인합니다.
5. 비아파트는 non_apt_sale, non_apt_jeonse, non_apt_rent의 세부 통계를 사용하되 거래건수가 적으면 방향성을 단정하지 않습니다.
6. 전세·월세 세부표는 거래건수, 신규계약 건수, 갱신재계약 건수를 구분해 해석하고, 상승·보합·하락 비율은 신규계약 비교가능건 기준임을 전제로 작성합니다.
7. 연령대별 서술은 age_population_rows와 migration_monthly_rows를 함께 확인하되, 부동산 수요와 연결할 근거가 약하면 단정하지 않습니다.


[종합 판단 권장 문장 구조]
① 당월 시장 방향
② 거래 집중 지역 또는 주요 단지
③ 가격 변동 주요 사례 또는 상승·하락 비율 변화
④ 지역 특성에 따른 해석
⑤ 조사상 유의사항

[부동산원 해당지역 담당 조사 전문가 관점 보강 규칙]
1. 단순 요약이 아니라 조사자가 현장에서 설명하듯 “왜 그런 흐름이 나타났는지”를 1~2문장 더 보강합니다.
2. dong_focus_summary의 상위 읍면동, detail_statistics_summary의 세부 통계, case_focus_summary의 대표 거래사례, non_apt_interpretation_summary의 유형별 거래건수, supply_pressure_summary의 공급압력을 각각 필요한 문단에 연결합니다.
3. 분석 내용은 풍부하게 쓰되 data_pack에 없는 외부 개발 호재·정비사업·단지명은 임의 생성하지 않습니다.
4. 거래건수·비교가능건수·신규계약 건수가 적은 경우에는 “거래가 제한적”이라고 짧게 표시합니다.
5. 전문 디자이너 관점에서 HTML 출력에 맞게 한 문장이 너무 길지 않도록 1문장 80~120자 내외로 나누어 작성합니다.

[LLM 필수 재작성 대상]
- Ⅵ 종합 시황 서술 보고서에 들어가는 아래 key는 기본 문장을 그대로 복사하지 말고 LLM이 data_pack을 읽어 새로 작성합니다.
- 필수 key: comprehensive_summary, sale_market_summary, apt_sale_summary, non_apt_sale_regional_context, rowhouse_multiplex_sale_summary, detached_multifamily_sale_summary, jeonse_market_summary, apt_jeonse_summary, non_apt_jeonse_regional_context, rowhouse_multiplex_jeonse_summary, detached_multifamily_jeonse_summary, rent_market_summary, apt_rent_summary, non_apt_rent_regional_context, rowhouse_multiplex_rent_summary, detached_multifamily_rent_summary
- 월간·오피스텔 보고서는 매매·전세·월세 파트별로 지역특성·배후수요·교통·상권·학군·인구·세대·이동·13개월 거래흐름·상승/하락 비율·통계 시계열을 자연스럽게 연결합니다.
- 수치를 앞세운 나열형 문장보다 생활권 특성, 배후수요, 상권, 교통, 학군, 원도심·신도심 맥락을 먼저 설명하고 실거래 사례와 수치를 근거로 붙입니다.
- “공표 통계” 또는 “공표통계”라는 표현은 쓰지 말고, 상승세 둔화·하락세 지속·보합권 유지처럼 읽히는 시장 흐름 표현으로 바꿉니다.
- 위 key가 비어 있으면 최종 HTML 검수에서 LLM 미반영으로 판정됩니다.

[출력 JSON 구조]
아래 key를 반드시 모두 포함하세요. JSON 외 문장은 쓰지 마세요.
{json.dumps(schema, ensure_ascii=False, indent=2)}

[Ⅵ. 종합 시황 서술 보고서 작성 기준]
- sale_market_summary: 매매 종합 4~6문장. 당월 매매 변화 포인트와 거래 집중 읍면동을 반드시 반영.
- apt_sale_summary: 아파트 매매 5~7문장. 앞의 [읍면동 단위 생활권 서술 규칙]을 그대로 적용. 첫 2~3문장은 거래 상위 읍면동의 생활권 성격(dong_profiles)과 배후수요·교통·학군·신축구축 구성으로 시작하고, 그 위에 당월 매매 방향과 상승·하락 우세를 얹음. 이어서 어느 생활권에서 어떤 면적대·연식대 수요가 움직였는지를 해석하고 수요·공급·금리 배경과 당월 판단으로 마무리. 대표 거래사례는 최대 2건만 근거로 쓰고 단지명·금액 나열 문장은 쓰지 않음.
- non_apt_sale_regional_context: 비아파트 매매 지역특성 300자 이상. non_apt_sale의 거래 상위 읍면동을 dong_profiles와 연결해 생활권별 수요 성격(원도심 저가 실수요, 노후 주거지, 임대수익형, 외곽 토지성 수요)을 정리하고 거래사례는 나열하지 않음. 최종 HTML에는 별도 제목으로 노출하지 않고, 아래 연립·다세대/단독·다가구 문단에 녹여 쓸 근거 문장으로 작성.
- rowhouse_multiplex_sale_summary: 연립·다세대 매매 3~5문장. 비아파트 매매 지역특성, 생활권, 대표 거래사례를 함께 반영하고 거래건수 적으면 방향성 단정 금지.
- detached_multifamily_sale_summary: 단독·다가구 매매 3~5문장. 비아파트 매매 지역특성, 임대수익형·노후주거·외곽/구도심 특성, 대표 거래사례를 근거가 있을 때만 사용.
- jeonse_market_summary: 전세 종합 4~6문장. 신규계약 기준 변화와 금리·보증금 부담 반영.
- apt_jeonse_summary: 아파트 전세 5~7문장. 앞의 [읍면동 단위 생활권 서술 규칙]을 그대로 적용. 거래 상위 읍면동의 생활권 성격, 학군·직주근접·산업 배후수요, 구축 비중과 임차수요 성격을 먼저 설명하고, 그 배경 위에 신규계약 기준 보증금 흐름·면적대별 수요 차이·금리 및 입주 여건을 근거로 연결. 보증금 사례를 나열하지 말고 임차수요가 어느 생활권으로 이동했는지를 중심으로 쓰며, 대표 사례는 최대 2건만 사용.
- non_apt_jeonse_regional_context: 비아파트 전세 지역특성 300자 이상. non_apt_jeonse의 거래 상위 읍면동을 dong_profiles와 연결해 생활권별 임차수요 성격을 정리하고 거래사례는 나열하지 않음. 최종 HTML에는 별도 제목으로 노출하지 않고, 아래 연립·다세대/단독·다가구 전세 문단에 녹여 쓸 근거 문장으로 작성.
- rowhouse_multiplex_jeonse_summary: 연립·다세대 전세 3~5문장. 저가 전세·관리상태·보증 안정성에 교통·학군·상권·배후수요와 대표 사례를 연결.
- detached_multifamily_jeonse_summary: 단독·다가구 전세 3~5문장. 1~2인 가구·고령층·구도심 수요는 실거래 근거가 있을 때만 자연스럽게 연결.
- rent_market_summary: 월세 종합 4~6문장. 환산월세 기준 변화와 전세 부담에 따른 월세 전환 가능성 반영.
- apt_rent_summary: 아파트 월세 5~7문장. 앞의 [읍면동 단위 생활권 서술 규칙]을 그대로 적용. 거래 상위 읍면동의 임대수요 성격(직주근접·산업 배후·소형 임차·원도심 저가 수요)을 먼저 설명하고, 그 위에 환산월세 흐름, 보증금·월세 구간별 특징, 전세 부담에 따른 월세 전환 가능성을 근거로 연결. 구간별 수치를 나열하지 말고 어떤 수요층이 어느 생활권에서 움직였는지를 중심으로 쓰며, 대표 사례는 최대 2건만 사용.
- non_apt_rent_regional_context: 비아파트 월세 지역특성 300자 이상. non_apt_rent의 거래 상위 읍면동을 dong_profiles와 연결해 생활권별 소형 임대수요·직주근접·상권 접근성 성격을 정리하고 거래사례는 나열하지 않음. 최종 HTML에는 별도 제목으로 노출하지 않고, 아래 연립·다세대/단독·다가구 월세 문단에 녹여 쓸 근거 문장으로 작성.
- rowhouse_multiplex_rent_summary: 연립·다세대 월세 3~5문장. 보증금·월세·환산월세 조건과 생활권 임대수요, 대표 거래사례를 연결.
- detached_multifamily_rent_summary: 단독·다가구 월세 3~5문장. 보증금 뒤 월세 조건이 해석에 드러나도록 쓰고, 직주근접·상권·소형 임대수요를 실거래 사례와 연결.

[region_profile]
{json.dumps(region_profile, ensure_ascii=False, indent=2)}

[monthly_change_summary]
{json.dumps(monthly_change_summary, ensure_ascii=False, indent=2)}

[dong_focus_summary]
{json.dumps(dong_focus_summary, ensure_ascii=False, indent=2)}

[detail_statistics_summary]
{json.dumps(detail_statistics_summary, ensure_ascii=False, indent=2)}

[case_focus_summary]
{json.dumps(case_focus_summary, ensure_ascii=False, indent=2)}

[non_apt_interpretation_summary]
{json.dumps(non_apt_interpretation_summary, ensure_ascii=False, indent=2)}

[supply_pressure_summary]
{json.dumps(supply_pressure_summary, ensure_ascii=False, indent=2)}

[HTML 디자인·가독성 참고]
- HTML 본문은 A4 PDF 출력 기준으로 읽히므로, 한 문단이 과도하게 길어지지 않게 작성합니다.
- 아파트 매매·전세·월세 세부 문단은 5~7문장으로 작성하되, 긴 문장 하나가 아니라 80~120자 안팎의 짧은 문장을 여러 개 연결합니다.
- 수치 나열보다 “핵심 변화 → 원인 → 판단” 순서로 정리합니다.
- 표·차트가 이미 제공하는 단순 수치는 반복하지 말고, 표·차트에서 드러나는 의미를 서술합니다.
- 종합 시황은 매매·전세·월세 박스별로 서로 다른 논지를 가져야 하며, 같은 문장 구조 반복을 피합니다.

{_llm_data_pack_block(data_pack, drop_keys=_LLM_SEPARATE_BLOCK_KEYS)}

{_KREB_SELF_CHECK}"""

def _scan_first_json_object(text: str) -> Optional[str]:
    """문자열에서 첫 번째 '균형 잡힌' JSON 객체({...})를 문자열·이스케이프를 인식하며 추출.
    rfind 방식과 달리 뒤따르는 본문에 중괄호가 있어도 과포착하지 않는다."""
    start = text.find("{")
    if start < 0:
        return None
    depth = 0
    in_str = False
    esc = False
    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start:i + 1]
    return None


def _repair_json_text(raw: str) -> str:
    """관용 JSON 보정: BOM·후행쉼표 제거, 전각 따옴표를 표준 따옴표로 정규화."""
    raw = raw.lstrip("﻿").strip()
    # 후행 쉼표: { ... , } 또는 [ ... , ]
    raw = re.sub(r",(\s*[}\]])", r"\1", raw)
    # 일부 LLM이 JSON 구분자로 쓰는 전각/스마트 따옴표 정규화
    raw = raw.replace("“", '"').replace("”", '"').replace("„", '"').replace("‟", '"')
    return raw


def extract_json_from_text(text: str) -> Dict[str, Any]:
    raw = (text or "").strip()
    if not raw:
        raise ValueError("입력이 비어 있습니다. LLM 출력 JSON을 붙여넣어 주세요.")
    # 1) ```json ... ``` 코드펜스 우선 (탐욕 매칭으로 마지막 ``` 까지)
    m = re.search(r"```(?:json)?\s*(.+?)\s*```", raw, flags=re.DOTALL | re.IGNORECASE)
    if m:
        candidate = m.group(1).strip()
    else:
        candidate = raw
    # 2) 균형 중괄호 스캔으로 첫 JSON 객체 추출 (없으면 후보 그대로)
    obj = _scan_first_json_object(candidate) or candidate
    # 3) 엄격 파싱 → 실패 시 관용 보정 후 재시도
    try:
        return json.loads(obj)
    except json.JSONDecodeError:
        repaired = _repair_json_text(obj)
        try:
            return json.loads(repaired)
        except json.JSONDecodeError as e:
            raise ValueError(
                f"JSON 파싱에 실패했습니다(보정 재시도 포함). 위치 line {e.lineno}, col {e.colno}: {e.msg}"
            ) from e


def merge_commentary(base: Dict[str, str], new: Dict[str, Any], data: Dict[str, Any] = None) -> Dict[str, str]:
    out = dict(base)
    applied_keys = []
    echoed_keys = []

    def _norm(s: str) -> str:
        # 공백·기호 차이를 무시하고 본문이 같은지 비교 (LLM이 1차 규칙기반 문장을 그대로 에코한 경우 감지)
        return re.sub(r"\s+", "", re.sub(r"<[^>]+>", "", str(s or "")))

    for k in COMMENTARY_KEYS:
        if k in new and str(new[k]).strip():
            out[k] = str(new[k])
            # 규칙기반 1차 문장과 사실상 동일하면 'LLM이 새로 쓴 것'으로 보지 않는다.
            if _norm(new[k]) == _norm(base.get(k, "")):
                echoed_keys.append(k)
            else:
                applied_keys.append(k)
    merged = {k: out.get(k, "") for k in COMMENTARY_KEYS}
    merged["_commentary_source"] = "llm_json"
    merged["_llm_applied_keys"] = applied_keys
    merged["_llm_echoed_keys"] = echoed_keys

    # v5.8.55 안전망: 비오피스텔 보고서일 때 LLM 이 잘못 쓴 '오피스텔'/'비오피스텔' 단어를
    # '아파트'/'비아파트'로 정정한다. (오피스텔 보고서는 그대로 두어 정상 노출되도록 한다.)
    corrected_keys = []
    try:
        meta = ((data or {}).get("meta") or {}) if isinstance(data, dict) else {}
        is_officetel = bool(meta.get("is_officetel")) or str(meta.get("report_mode") or "") == "officetel"
        if not is_officetel:
            for k in COMMENTARY_KEYS:
                v = merged.get(k)
                if isinstance(v, str) and ("오피스텔" in v):
                    # 순서 중요: '비오피스텔' 먼저 처리해야 '비' + '오피스텔' 분리 치환을 피함.
                    v = v.replace("비오피스텔", "비아파트")
                    v = v.replace("오피스텔", "아파트")
                    merged[k] = v
                    corrected_keys.append(k)
    except Exception:
        pass
    merged["_officetel_corrected_keys"] = corrected_keys
    return merged


def required_llm_summary_keys(data: Dict[str, Any]) -> List[str]:
    """보고서 모드별 'LLM이 반드시 새로 작성해야 하는' Ⅵ 종합 시황 key 목록."""
    meta = (data or {}).get("meta", {}) if isinstance(data, dict) else {}
    is_officetel = bool(meta.get("is_officetel")) or str(meta.get("report_mode") or "") == "officetel"
    if is_weekly_report(data):
        return list(WEEKLY_LLM_FINAL_SUMMARY_KEYS)
    if is_officetel:
        return ["comprehensive_summary", "sale_market_summary", "apt_sale_summary",
                "jeonse_market_summary", "apt_jeonse_summary", "rent_market_summary", "apt_rent_summary"]
    return list(LLM_FINAL_SUMMARY_KEYS)


def summarize_llm_application(commentary: Dict[str, Any], data: Dict[str, Any]) -> str:
    """LLM JSON을 붙여넣은 직후 보여줄 반영 상태 요약(재생성 전 사전 점검)."""
    required = required_llm_summary_keys(data)
    applied = set(commentary.get("_llm_applied_keys") or [])
    echoed = set(commentary.get("_llm_echoed_keys") or [])
    corrected = commentary.get("_officetel_corrected_keys") or []
    missing = [k for k in required if not str(commentary.get(k) or "").strip()]
    echoed_required = [k for k in required if k in echoed]
    lines = [f"LLM JSON 적용: 새로 작성된 {len(applied)}개 key 반영."]
    if missing:
        lines.append(f"⚠ 필수 종합서술 누락/빈값 {len(missing)}개: {', '.join(missing)}")
    if echoed_required:
        lines.append(f"⚠ 1차 문장 그대로(미작성 의심) {len(echoed_required)}개: {', '.join(echoed_required)}")
    if corrected:
        lines.append(f"ℹ '오피스텔→아파트' 자동 정정 {len(corrected)}개 key.")
    if not missing and not echoed_required:
        lines.append("필수 종합서술 key 모두 정상 반영됨.")
    return "\n".join(lines)


# =========================================================
# 8. HTML 렌더링
# =========================================================

DISPLAY_NAME_MAP = {
    "month": "기준월",
    "week": "공표일",
    "period_label": "표시기준",
    "period_type": "표시구분",
    "source": "출처",
    "population": "주민등록인구",
    "households": "주민등록세대수",
    "population_mom": "인구 전월 대비",
    "households_mom": "세대 전월 대비",
    "in": "총 전입",
    "out": "총 전출",
    "net": "순이동",
    "age": "연령대",
    "age_group": "연령대",
    "trade_count": "거래건수",
    "_trade_mom": "전월 대비 증감",
    "compare_count": "비교가능건수",
    "avg_change_rate": "평균 변동률",
    "public_index_change_rate": "공표통계",
    "publication_status": "공표상태",
    "up_ratio": "🔴 상승 비율",
    "flat_ratio": "⚪ 보합 비율",
    "down_ratio": "🔵 하락 비율",
    "판정": "📊 판정",
    "상승": "🔴 상승건수",
    "보합": "⚪ 보합건수",
    "하락": "🔵 하락건수",
    "sigungu": "시군구",
    "dong": "읍면동",
    "share": "비중",
    "count": "거래건수",
    "band": "구간",
    "rank": "순위",
    "property_type": "유형",
    "location": "소재지",
    "complex_name": "단지명",
    "building_name": "건물명",
    "apt_name": "아파트명",
    "avg_price_eok": "평균거래금액(억)",
    "avg_value": "평균값",
    "main_area": "주요 전용면적",
    "main_pyeong": "평형",
    "approval_year": "사용승인년도",
    "households": "세대수",
    "area": "전용면적",
    "pyeong": "평형",
    "current_price_eok": "현재거래(억)",
    "prev_price_eok": "직전거래(억)",
    "transaction_price_eok": "거래금액(억)",
    "deposit_eok": "보증금(억)",
    "monthly_rent_manwon": "월세(만원)",
    "converted_rent_manwon": "환산월세(만원)",
    "ytd_count": "누적 거래건수",
    "base_month_count": "당월 거래건수",
    "ytd_avg_price_eok": "누적 평균 거래금액(억)",
    "ytd_new_count": "누적 신규계약 건수",
    "ytd_renew_count": "누적 갱신재계약 건수(비율)",
    "ytd_avg_deposit_eok": "누적 평균 보증금(억)",
    "ytd_avg_monthly_rent_manwon": "누적 평균 월세(만원)",
    "ytd_avg_converted_rent_manwon": "누적 평균 환산월세(만원)",
    "avg_deposit_eok": "평균 보증금(억)",
    "avg_monthly_rent_manwon": "평균 월세(만원)",
    "avg_converted_rent_manwon": "평균 환산월세(만원)",
    "contract_date": "거래일자",
    "current_value": "현재값",
    "prev_value": "직전값",
    "change_rate": "변동률",
    "floor": "층",
    "mortgage_rate": "주택담보대출",
    "jeonse_rate": "전세자금대출",
    "sme_rate": "중소기업대출",
    "business_rate": "기업일반자금대출",
    "unsold_units": "미분양 호수",
    "after_completion_unsold": "준공후 미분양",
    "value": "지수",
    "permit": "인허가",
    "start": "착공",
    "completion": "준공",
    "presale_supply": "분양",
    "0~9세": "0~9세",
    "10대": "10대",
    "20대": "20대",
    "30대": "30대",
    "40대": "40대",
    "50대": "50대",
    "60대": "60대",
    "70대 이상": "70대 이상",
    "최신월 인구": "최신월 인구",
    "비중": "비중",
    "13개월 증감": "13개월 증감",
    "시장 해석": "시장 해석",
    "구분": "구분",
    "내용": "내용",
}

DIRECTION_COL_CLASS = {
    "up_ratio": "dir-up",
    "flat_ratio": "dir-flat",
    "down_ratio": "dir-down",
    "상승": "dir-up",
    "보합": "dir-flat",
    "하락": "dir-down",
    "_up_count": "dir-up",
    "_flat_count": "dir-flat",
    "_down_count": "dir-down",
}

def direction_col_class(col: Any) -> str:
    return DIRECTION_COL_CLASS.get(str(col), "")

def display_col_name(col: Any, public_stat_label: str = "") -> str:
    if str(col) == "public_index_change_rate" and public_stat_label:
        return public_stat_label
    _extra_display_names = {
        "new_count": "신규계약 건수",
        "renew_count": "갱신재계약 건수(비율)",
        "new_count": "신규계약 건수",
        "trade_count": "거래건수",
        "_trade_mom": "전월 대비 증감",
        "_ratio_bar": "🔵하락 ◀ ▶ 상승🔴",
        "compare_count": "비교가능건수",
        "avg_change_rate": "평균 변동률",
        "public_index_change_rate": "공표통계",
        "publication_status": "공표상태",
        "up_ratio": "🔴 상승 비율",
        "flat_ratio": "⚪ 보합 비율",
        "down_ratio": "🔵 하락 비율",
    }
    s = str(col)
    return _extra_display_names.get(s, DISPLAY_NAME_MAP.get(s, s))

def html_value(v: Any, col: Any = "") -> str:
    c = str(col).lower()
    is_public_stat = str(col) == "public_index_change_rate" or "공표통계" in str(col)
    if v is None:
        if is_public_stat:
            return "<span class='not-published'>미공표</span>"
        return "-"
    try:
        if pd.isna(v):
            if is_public_stat:
                return "<span class='not-published'>미공표</span>"
            return "-"
    except Exception:
        pass
    if isinstance(v, (int, float)):
        if any(k in c for k in ["rate", "ratio", "변동률", "비율", "share", "비중"]):
            return f"{float(v):.1f}"
        if c in ["value", "지수"] or "index" in c:
            return f"{float(v):.4f}"
        if abs(float(v)) < 100 and not float(v).is_integer():
            return f"{float(v):.2f}"
        return f"{int(round(float(v))):,}"
    return str(v)

_CURRENT_BASE_MONTH = ""  # render_html에서 보고서 조사기준월(YYYY-MM)을 설정
_CURRENT_CREATED_DATE = ""  # render_html에서 보고서 생성일(YYYY-MM-DD)을 설정


def _report_deadline_pending(month: Any) -> bool:
    """실거래 신고기한(계약 후 30일)이 끝나지 않은 월인지: 생성일 < 해당 월 말일 + 30일."""
    try:
        end = pd.Period(str(month)[:7], freq="M").end_time.normalize()
        created = pd.Timestamp(_CURRENT_CREATED_DATE[:10]) if _CURRENT_CREATED_DATE else pd.Timestamp(datetime.now().date())
        return created < end + pd.Timedelta(days=30)
    except Exception:
        return False


def ratio_bar_scale(rows: List[Dict[str, Any]], keys: Tuple[str, str, str]) -> float:
    """표 안에서 공통으로 쓸 막대 눈금 끝(%): 상승·하락 비율 최댓값을 10 단위로 올림(최소 50)."""
    peak = 0.0
    for r in rows:
        for k in (keys[0], keys[2]):
            peak = max(peak, _clean_num(r.get(k)) or 0.0)
    return float(max(50, -(-int(peak) // 10) * 10))


def ratio_bar_html(up: Any, flat: Any, down: Any, scale: float = 100.0) -> str:
    """가운데 기준선에서 하락은 왼쪽(파랑), 상승은 오른쪽(빨강)으로 뻗는 막대와 우세 폭(%p).
    같은 표 안에서는 같은 눈금(scale)을 써서 행끼리 길이를 비교할 수 있게 한다. 비교가능 거래가 없으면 '-'."""
    u, f, d = (_clean_num(v) or 0.0 for v in (up, flat, down))
    if u + f + d <= 0:
        return "-"
    gap = u - d
    if abs(gap) < 0.05:
        net = "<span class='db-net flat'>균형</span>"
    else:
        net = f"<span class='db-net {'up' if gap > 0 else 'down'}'>{'▲' if gap > 0 else '▼'}{abs(gap):.1f}p</span>"
    return (
        f"<div class='div-bar-wrap' title='상승 {u:.1f}% · 보합 {f:.1f}% · 하락 {d:.1f}%'>"
        f"<div class='div-bar'><div class='db-side db-left'><span class='db-down' style='width:{min(d / scale * 100, 100):.1f}%'></span></div>"
        f"<div class='db-side db-right'><span class='db-up' style='width:{min(u / scale * 100, 100):.1f}%'></span></div></div>"
        f"{net}</div>"
    )


def table_rows_html(rows: List[Dict[str, Any]], max_rows: int = 13, highlight_latest: bool = True) -> str:
    if not rows:
        return "<p class='muted'>자료 없음</p>"
    first_col = str(next(iter(rows[0].keys()), "")) if rows and isinstance(rows[0], dict) else ""
    is_time_series = len(rows) >= 12 and any(k in first_col.lower() for k in ["month", "기준월", "_month"])
    full_rows = rows
    rows = rows[-max_rows:] if is_time_series else rows[:max_rows]
    public_stat_label = ""
    for r in rows:
        label = str(r.get("_public_stat_label", "")).strip() if isinstance(r, dict) else ""
        if label:
            public_stat_label = label
            break
    hide_cols = {
        "avg_value", "avg_price_eok", "평균값", "평균가", "평균가격", "_public_stat_label",
        "_cluster_influence_label", "_cluster_influence_level",
    }
    cols = [c for c in list(rows[0].keys()) if str(c) not in hide_cols]
    # 상승·보합·하락 건수/비율 열은 원자료 집계 순서(보합·상승·하락)와 관계없이 항상 상승→보합→하락으로 둔다.
    for triple in (("상승", "보합", "하락"), ("up_count", "flat_count", "down_count"), ("up_ratio", "flat_ratio", "down_ratio"),
                   ("상승건수", "보합건수", "하락건수"), ("상승 비율", "보합 비율", "하락 비율")):
        present = [c for c in triple if c in cols]
        if len(present) >= 2:
            slots = sorted(cols.index(c) for c in present)
            for slot, c in zip(slots, present):
                cols[slot] = c
    is_trade_volume_series = is_time_series and "trade_count" in cols
    trade_mom = {}
    if is_trade_volume_series:
        # 거래건수 바로 뒤에 전월 대비 증감 열을 둔다(첫 표시월도 원자료의 직전월과 비교).
        cols.insert(cols.index("trade_count") + 1, "_trade_mom")
        offset = len(full_rows) - len(rows)
        for i, r in enumerate(rows):
            j = offset + i
            prev = full_rows[j - 1].get("trade_count") if j > 0 else None
            if _report_deadline_pending(r.get(first_col)):
                trade_mom[i] = "<span class='pending-label'>신고미도래</span>"
            elif prev is None:
                trade_mom[i] = "-"
            else:
                try:
                    diff = int(r.get("trade_count") or 0) - int(prev or 0)
                    trade_mom[i] = f"<span class='{'pos' if diff > 0 else ('neg' if diff < 0 else 'flat')}'>{diff:+,}</span>" if diff else "0"
                except Exception:
                    trade_mom[i] = "-"
    # 상승·보합·하락 비율이 있는 표는 마지막 비율 열 뒤에 가로 누적 막대 열을 둔다.
    ratio_keys = next((t for t in (("up_ratio", "flat_ratio", "down_ratio"), ("상승 비율", "보합 비율", "하락 비율"))
                       if all(k in cols for k in t)), None)
    if ratio_keys:
        cols.insert(max(cols.index(k) for k in ratio_keys) + 1, "_ratio_bar")
        bar_scale = ratio_bar_scale(rows, ratio_keys)
    if "sigungu" in cols and "dong" in cols and cols.index("sigungu") > cols.index("dong"):
        cols.remove("sigungu")
        cols.insert(cols.index("dong"), "sigungu")
    share_cols = [c for c in cols if ("share" in str(c).lower() or "비중" in str(c))]
    share_max = {}
    for c in share_cols:
        vals = []
        for r in rows:
            try:
                vals.append(float(r.get(c)))
            except Exception:
                pass
        share_max[c] = max(vals) if vals else None
    # 거래건수·누적 거래건수 열은 표마다 같은 방식으로 강조한다.
    count_cols = {"count", "ytd_count", "trade_count"}
    th = "".join(
        ("<th class='count-em'>" if str(c) in count_cols else ("<th class='ratio-bar-th'>" if c == "_ratio_bar" else "<th>"))
        + f"{display_col_name(c, public_stat_label)}</th>"
        for c in cols
    )
    trs = []
    name_cols = {"apt_name", "complex_name", "building_name", "_group_name", "단지명", "아파트명", "건물명", "오피스텔명"}
    for idx, r in enumerate(rows):
        row_classes = []
        if highlight_latest and idx == len(rows)-1:
            row_classes.append("latest-row")
        if is_time_series and _CURRENT_BASE_MONTH and str(r.get(first_col, "")).strip().startswith(_CURRENT_BASE_MONTH):
            row_classes.append("base-month-row")
        cls = f" class='{' '.join(row_classes)}'" if row_classes else ""
        tds = []
        for c in cols:
            if c == "_trade_mom":
                tds.append(f"<td class='trade-mom-cell'>{trade_mom.get(idx, '-')}</td>")
                continue
            if c == "_ratio_bar":
                tds.append(f"<td class='ratio-bar-cell'>{ratio_bar_html(*(r.get(k) for k in ratio_keys), scale=bar_scale)}</td>")
                continue
            val = r.get(c)
            classes = []
            if str(c) in name_cols:
                classes.append("name-cell")
            if str(c) in count_cols:
                classes.append("count-em")
            dcls = direction_col_class(c)
            if dcls:
                classes.append(dcls)
            if str(c) == "판정":
                vcls = VERDICT_CLASS.get(str(val).strip())
                if vcls:
                    classes.append(vcls)
                classes.append("verdict-cell")
            try:
                fv = float(val)
                if any(k in str(c).lower() for k in ["change", "증감", "net", "변동률"]):
                    if fv > 0: classes.append("pos")
                    elif fv < 0: classes.append("neg")
                if c in share_max and share_max.get(c) is not None and fv == share_max[c] and fv > 0:
                    classes.append("share-high")
            except Exception:
                pass
            cell_cls = f" class='{' '.join(classes)}'" if classes else ""
            cell_html = html_value(val, c)
            if str(c) in {"dong", "읍면동"} and r.get("_cluster_influence_label"):
                level = html.escape(str(r.get("_cluster_influence_level") or "linked"))
                label = html.escape(str(r.get("_cluster_influence_label")))
                cell_html = (
                    f"<strong>{cell_html}</strong><br>"
                    f"<span class='semiconductor-cluster-badge {level}'>{label}</span>"
                )
            if is_trade_volume_series and str(c).lower() in {"month", "기준월", "_month"} and _report_deadline_pending(val):
                cell_html = f"{cell_html}<br><span class='preliminary-label' style='color:#b45309;font-size:10px;font-weight:700'>(신고미도래)</span>"
            if str(c) in {"renew_count", "ytd_renew_count"}:
                total = r.get("ytd_count") if str(c) == "ytd_renew_count" else (r.get("trade_count") if r.get("trade_count") is not None else r.get("count"))
                try:
                    if total and float(total) > 0:
                        cell_html = f"{cell_html} <span class='renew-ratio'>({float(val or 0) / float(total) * 100:.1f}%)</span>"
                except Exception:
                    pass
            if str(c) == "판정":
                emoji = VERDICT_EMOJI.get(str(val).strip(), "")
                cell_html = f"{emoji} {cell_html}".strip()
            if str(c) in name_cols:
                cell_html = f"<strong class='name-cell-text'>{cell_html}</strong>"
            tds.append(f"<td{cell_cls}>{cell_html}</td>")
        trs.append(f"<tr{cls}>" + "".join(tds) + "</tr>")
    table_cls = "market-strength-table compact-wide-table" + (" time-series-table" if is_time_series else "")
    has_cluster_labels = any(bool(r.get("_cluster_influence_label")) for r in rows if isinstance(r, dict))
    cluster_legend = semiconductor_cluster_legend_html() if has_cluster_labels else ""
    return cluster_legend + f"<div class='table-wrap'><table class='{table_cls}'><thead><tr>{th}</tr></thead><tbody>{''.join(trs)}</tbody></table></div>"


def semiconductor_cluster_legend_html() -> str:
    return (
        "<div class='semiconductor-cluster-legend'>"
        "<div><span class='semiconductor-cluster-badge core'>핵심 영향권</span> "
        "<strong>송정동·도산동·신촌동·소촌동</strong> — 광주 군공항 사업 예정지와 맞닿거나 가장 가까운 생활권으로 우선 관찰</div>"
        "<div><span class='semiconductor-cluster-badge adjacent'>인접 영향권</span> "
        "<strong>우산동</strong> — 군공항 북동측 인접 생활권으로 거래량·매매·전세 변화를 함께 관찰</div>"
        "<div><span class='semiconductor-cluster-badge linked'>연계 영향권</span> "
        "<strong>선암동·운수동</strong> — 송정 생활권 및 교통축으로 연결된 지역으로 직접 수혜를 단정하지 않고 연계 흐름을 관찰</div>"
        "<p>※ 영향권은 공식 사업구역·수혜지역 명칭이 아닌 보고서용 분석 분류이며, 실제 거래 변화가 확인될 때만 정책 맥락과 연결함.</p>"
        "</div>"
    )


def dong_volume_trend_table_html(rows: List[Dict[str, Any]]) -> str:
    """기존 읍면동×13개월 표를 유지하면서 전월 대비 증감 방향을 셀에 시각화한다."""
    if not rows:
        return "<p class='muted'>자료 없음</p>"
    cols = [c for c in rows[0].keys() if not str(c).startswith("_")]
    month_cols = [c for c in cols if re.fullmatch(r"20\d{2}-\d{2}", str(c))]
    latest_month = month_cols[-1] if month_cols else ""
    th = "".join(
        f"<th>{html.escape(str(c))}{'<br><span style=\'color:#fde68a;font-size:10px\'>(잠정치)</span>' if c == latest_month else ''}</th>"
        for c in cols
    )
    body = []
    for row in rows:
        cells = []
        previous = row.get("_previous_month_count")
        for c in cols:
            value = row.get(c)
            if c not in month_cols:
                content = html.escape(str(value if value is not None else "-"))
                if str(c) == "읍면동":
                    content = f"<strong>{content}</strong>"
                    if row.get("_cluster_influence_label"):
                        level = html.escape(str(row.get("_cluster_influence_level") or "linked"))
                        label = html.escape(str(row.get("_cluster_influence_label")))
                        content += f"<br><span class='semiconductor-cluster-badge {level}'>{label}</span>"
                cells.append(f"<td>{content}</td>")
                continue
            try:
                current = int(float(value or 0))
            except Exception:
                current = 0
            if previous is None:
                marker = "<span style='color:#64748b;font-size:10px'>비교월 없음</span>"
                style = "background:#f8fafc"
            else:
                delta = current - previous
                if delta > 0:
                    marker = f"<span style='color:#dc2626;font-size:10px;font-weight:700'>▲{delta}</span>"
                    style = "background:#fff1f2;color:#b91c1c"
                elif delta < 0:
                    marker = f"<span style='color:#2563eb;font-size:10px;font-weight:700'>▼{abs(delta)}</span>"
                    style = "background:#eff6ff;color:#1d4ed8"
                else:
                    marker = "<span style='color:#64748b;font-size:10px;font-weight:700'>―0</span>"
                    style = "background:#f8fafc;color:#475569"
            latest_outline = ";box-shadow:inset 0 0 0 1px #f59e0b" if c == month_cols[-1] else ""
            cells.append(
                f"<td style='{style}{latest_outline};min-width:62px'>"
                f"<div style='font-weight:700'>{current:,}</div>{marker}</td>"
            )
            previous = current
        body.append("<tr>" + "".join(cells) + "</tr>")
    legend = (
        "<div style='display:flex;gap:14px;justify-content:flex-end;align-items:center;"
        "font-size:11px;color:#475569;margin:4px 2px 8px'>"
        "<span style='color:#dc2626;font-weight:700'>▲ 전월 대비 증가</span>"
        "<span style='color:#2563eb;font-weight:700'>▼ 전월 대비 감소</span>"
        "<span style='color:#64748b;font-weight:700'>― 보합</span>"
        "<span style='color:#b45309'>□ 최신월</span></div>"
    )
    has_cluster_labels = any(bool(r.get("_cluster_influence_label")) for r in rows if isinstance(r, dict))
    cluster_legend = semiconductor_cluster_legend_html() if has_cluster_labels else ""
    return (
        cluster_legend + legend + "<div class='table-wrap'><table class='market-strength-table compact-wide-table dong-volume-trend-table'>"
        f"<thead><tr>{th}</tr></thead><tbody>{''.join(body)}</tbody></table></div>"
    )


def _fmt_eok_value(v: Any) -> str:
    if v is None:
        return "-"
    try:
        if pd.isna(v):
            return "-"
    except Exception:
        pass
    try:
        fv = float(v)
    except Exception:
        return str(v)
    if abs(fv) >= 1:
        # 억 단위 소수 둘째 자리 → 억/만원 혼합 표현 대신 보고서 카드용 간결 표현
        return f"{fv:.2f}억".replace(".00억", "억")
    return f"{int(round(fv * 10000)):,}만원"

def _fmt_manwon_value(v: Any) -> str:
    if v is None:
        return "-"
    try:
        if pd.isna(v):
            return "-"
    except Exception:
        pass
    try:
        fv = float(v)
    except Exception:
        return str(v)
    if abs(fv) >= 10000:
        eok = fv / 10000
        return f"{eok:.2f}억".replace(".00억", "억")
    return f"{int(round(fv)):,}만원"

def _fmt_delta(curr: Any, prev: Any, value_label: str, chg: Any) -> str:
    try:
        if prev is None or pd.isna(prev):
            return "직전거래 없음"
    except Exception:
        if prev is None:
            return "직전거래 없음"
    try:
        delta = float(curr) - float(prev)
    except Exception:
        delta = None
    try:
        chg_s = f"{float(chg):+.1f}%"
    except Exception:
        chg_s = "직전거래 없음"
    if delta is None:
        return chg_s
    sign = "+" if delta > 0 else ""
    if value_label == "매매":
        delta_s = _fmt_eok_value(delta)
    else:
        delta_s = _fmt_manwon_value(delta)
    return f"{sign}{delta_s} ({chg_s})"


def format_contract_date_display(x: Any, fallback_month: Any = "") -> str:
    """TOP10 사례용 계약일 표시: YYYY.MM.DD 형식으로 정규화.
    원자료 계약일이 1~31일만 들어오는 경우 계약월/직전거래월과 결합해 YYYY.MM.DD로 표시한다.
    """
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    s = str(x).strip()
    if not s or s.lower() == "nan":
        return ""

    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.notna(dt) and not re.fullmatch(r"\d{1,2}(\.0)?", s):
            return dt.strftime("%Y.%m.%d")
    except Exception:
        pass

    nums = re.sub(r"[^0-9]", "", s)
    fm = str(fallback_month or "").strip()
    fm_nums = re.sub(r"[^0-9]", "", fm)
    year = month = ""
    if len(fm_nums) >= 6:
        year, month = fm_nums[:4], fm_nums[4:6]
    elif re.fullmatch(r"20\d{2}-\d{2}", fm):
        year, month = fm[:4], fm[5:7]
    elif re.fullmatch(r"20\d{2}\.\d{2}", fm):
        year, month = fm[:4], fm[5:7]

    if len(nums) >= 8:
        return f"{nums[:4]}.{nums[4:6]}.{nums[6:8]}"
    if len(nums) == 6:
        return f"20{nums[:2]}.{nums[2:4]}.{nums[4:6]}"
    if len(nums) in [1, 2] and year and month:
        return f"{year}.{month}.{nums.zfill(2)}"
    if len(nums) == 4 and year:
        return f"{year}.{nums[:2]}.{nums[2:4]}"
    return s


def top_case_list_html(rows: List[Dict[str, Any]], title: str, value_label: str = "거래", section_type: str = "auto") -> str:
    """template_v58 TOP10 내부 list 구조.
    - top-case-card는 만들지 않고 top-case-list만 반환
    - 하나의 top-case-card 안에 3개 top-case-list가 배치되도록 함
    """
    rows = list(rows or [])
    padded = rows[:10] + [{"rank": i, "_empty": True} for i in range(len(rows) + 1, 11)]
    items = [f"<div class='top-case-list'><h4>{title}</h4>"]
    for idx, r in enumerate(padded[:10], 1):
        if section_type == "rise":
            cls = "case-rise"
        elif section_type == "fall":
            cls = "case-fall"
        else:
            cls = "case-top"
        if r.get("_empty"):
            # template_v58과 같은 슬롯 수 유지, 색상은 중립형
            items.append(f"""
            <div class="top-case-item case-zero">
              <div class="case-rank">{idx}</div>
              <div class="case-body">
                <b>자료 없음 <span class="case-year"></span></b>
                <span>해당 순위 거래사례 없음</span>
                <small>현재 - / 직전 -</small>
              </div>
              <div class="case-value"><b>-</b><span>-</span></div>
            </div>
            """)
            continue

        chg = r.get("change_rate", None)
        if section_type == "auto":
            try:
                _chg_num = float(chg)
            except Exception:
                _chg_num = 0
            if _chg_num > 0:
                cls = "case-top case-rise"
            elif _chg_num < 0:
                cls = "case-top case-fall"
            else:
                cls = "case-top case-zero"

        rank = r.get("rank", idx)
        name = r.get("building_name") or r.get("complex_name") or r.get("apt_name") or r.get("_group_name") or "-"
        year = str(r.get("approval_year") or "").replace(".0", "")
        dong = r.get("dong") or "-"
        sigungu = r.get("sigungu") or ""
        location_label = compose_sigungu_dong_label(sigungu, dong)
        # 동명과 전용면적은 카드에서 바로 눈에 띄도록 굵게 표시
        dong_s = clean_region_part(dong)
        if dong_s and location_label.endswith(dong_s):
            location_label = location_label[: -len(dong_s)] + f"<span class='case-em'>{dong_s}</span>"
        area = r.get("area")
        pyeong = r.get("pyeong")
        floor = r.get("floor")
        prev_floor = r.get("prev_floor")
        contract_date = format_contract_date_display(r.get("contract_date") or "", r.get("month") or "")
        prev_contract_date = format_contract_date_display(r.get("prev_contract_date") or "", r.get("prev_month") or "")

        pyeong_s = str(pyeong or "")
        if pyeong_s and not pyeong_s.endswith(("평", "평형")):
            pyeong_s += "평"
        area_s = f"<span class='case-em'>전용 {area}㎡</span>({pyeong_s})" if area and pyeong_s else (f"<span class='case-em'>전용 {area}㎡</span>" if area else "전용면적 미상")
        floor_s = f"현재 {floor}층 / 직전 {prev_floor}층" if floor is not None and prev_floor is not None else (f"현재 {floor}층 / 직전층 원자료 확인 필요" if floor is not None else "현재층·직전층 원자료 확인 필요")
        date_pair_s = (f"현재거래 {contract_date}" if contract_date else "") + (f" / 직전거래 {prev_contract_date}" if prev_contract_date else "")
        prev_date_s = f" · {date_pair_s}" if date_pair_s else ""
        # 현재층이 직전층보다 높으면 층 우세(가격 상승 요인), 낮으면 층 열세로 강조한다.
        floor_badge = ""
        try:
            if floor is not None and prev_floor is not None:
                gap = int(floor) - int(prev_floor)
                if gap > 0:
                    floor_badge = f'<em class="floor-badge floor-up">▲ 층 우세 +{gap}층</em>'
                elif gap < 0:
                    floor_badge = f'<em class="floor-badge floor-down">▼ 층 열세 {gap}층</em>'
                else:
                    floor_badge = '<em class="floor-badge floor-same">동일층</em>'
        except Exception:
            floor_badge = ""
        # 직전계약 대비 2배 이상·절반 이하(매매는 1.5배 이상·3분의 2 이하)로 바뀐 사례는 임대주택형 저가 계약·특수거래와
        # 맞물렸을 수 있으므로 통계에서 빼지 않고 카드에 원자료 확인 표시를 붙인다.
        outlier_badge = ""
        try:
            factor = 1.5 if value_label == "매매" else 2.0
            ratio = 1 + float(chg) / 100 if chg is not None else 1.0
            if ratio >= factor or ratio <= 1 / factor:
                outlier_badge = '<em class="outlier-badge" title="직전계약과의 변동폭이 커서 계약 성격(임대주택·특수관계 등) 확인이 필요함">⚠ 변동폭 과대·원자료 확인</em>'
        except Exception:
            outlier_badge = ""
        deal_type = str(r.get("deal_type") or "").strip()
        if deal_type == "직거래":
            deal_badge = '<em class="deal-badge deal-direct">직거래</em>'
        elif deal_type == "중개거래":
            deal_badge = '<em class="deal-badge deal-broker">중개거래</em>'
        else:
            # 전월세 신고자료처럼 거래유형 항목이 없으면 표시하지 않는다.
            deal_badge = ""
        if str(r.get("contract_class") or "") == "갱신재계약":
            floor_badge += '<em class="floor-badge floor-same">갱신계약</em>'

        def _short_date(text: str) -> str:
            return text[2:] if len(text) == 10 and text[:2] == "20" else (text or "-")

        def _floor_text(value: Any) -> str:
            return f"{value}층" if value not in (None, "") else "-"

        prev_deal = str(r.get("prev_deal_type") or "").strip()
        if value_label == "매매":
            curr = r.get("current_price_eok", r.get("current_value", None))
            prev = r.get("prev_price_eok", r.get("prev_value", None))
            cur_amt, prev_amt = _fmt_eok_value(curr), _fmt_eok_value(prev)
            if prev_deal:
                prev_amt += f"<small>{prev_deal.replace('거래', '')}</small>"
        elif value_label == "월세":
            curr = r.get("current_value", None)
            prev = r.get("prev_value", None)
            try:
                is_converted = abs(float(curr) - float(r.get("current_monthly_rent"))) > 0.05
            except Exception:
                is_converted = False
            cur_amt = f"<span class='nw'>{_fmt_manwon_value(r.get('current_deposit'))}</span> / <span class='nw'>{_fmt_manwon_value(r.get('current_monthly_rent'))}</span>"
            if r.get("prev_deposit") is not None or r.get("prev_monthly_rent") is not None:
                prev_amt = f"<span class='nw'>{_fmt_manwon_value(r.get('prev_deposit'))}</span> / <span class='nw'>{_fmt_manwon_value(r.get('prev_monthly_rent'))}</span>"
            else:
                prev_amt = _fmt_manwon_value(prev)
            if is_converted:
                cur_amt += f"<small>환산 {_fmt_manwon_value(curr)}</small>"
                prev_amt += f"<small>환산 {_fmt_manwon_value(prev)}</small>"
        elif value_label == "전세":
            curr = r.get("current_value", None)
            prev = r.get("prev_value", None)
            cur_amt, prev_amt = _fmt_manwon_value(curr), _fmt_manwon_value(prev)
        else:
            curr = r.get("current_value", None)
            prev = r.get("prev_value", None)
            cur_amt, prev_amt = html_value(curr), html_value(prev)

        delta_full = _fmt_delta(curr, prev, value_label, chg)
        if " (" in delta_full and delta_full.endswith(")"):
            delta_s, rate_s = delta_full[:-1].split(" (", 1)
        else:
            delta_s, rate_s = "", delta_full
        amount_head = {"매매": "거래금액", "전세": "보증금", "월세": "보증금 / 월세"}.get(value_label, "금액")
        meta_parts = [location_label] + ([area_s.split("(")[0], pyeong_s] if pyeong_s and "(" in area_s else [area_s])
        if year and year != "nan":
            meta_parts.append(f'<span class="case-year">{year}년 사용승인</span>')
        else:
            meta_parts.append('<span class="case-year"></span>')
        badges = f"{deal_badge}{floor_badge}{outlier_badge}"
        items.append(f"""
        <div class="top-case-item {cls} tc-v2">
          <div class="case-rank">{rank}</div>
          <div class="case-body">
            <b>{name}</b>
            <div class="tc-meta">{' · '.join(f'<span class="tc-seg">{part}</span>' for part in meta_parts)}</div>
            {f'<div class="tc-badges">{badges}</div>' if badges else ''}
          </div>
          <div class="case-value"><span class="tc-rate">{rate_s}</span><span class="tc-delta">{delta_s}</span></div>
          <div class="tc-compare">
            <div class="tc-row tc-head"><div></div><div>계약일</div><div>층</div><div>{amount_head}</div></div>
            <div class="tc-row tc-cur"><div class="tc-label">현재</div><div>{_short_date(contract_date)}</div><div>{_floor_text(floor)}</div><div class="tc-amt">{cur_amt}</div></div>
            <div class="tc-row tc-prev"><div class="tc-label">직전</div><div>{_short_date(prev_contract_date)}</div><div>{_floor_text(prev_floor)}</div><div class="tc-amt">{prev_amt}</div></div>
          </div>
        </div>
        """)
    items.append("</div>")
    return "".join(items)

def top_cards_html(rows: List[Dict[str, Any]], value_label: str = "거래") -> str:
    """하위 호환용: 단일 top-case-list만 반환."""
    return top_case_list_html(rows, "TOP10", value_label, "auto")


def note_box(text: str) -> str:
    return f"<div class='note'>💡 {text}</div>"




def js_json(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False)

def _clean_num(v: Any) -> Any:
    try:
        if v is None or pd.isna(v):
            return None
    except Exception:
        pass
    try:
        return round(float(v), 4)
    except Exception:
        return None


def chart_block(title: str, chart_type: str, labels: List[Any], datasets: List[Dict[str, Any]], chart_id: str, note: str = "", stacked: bool = False, y1: bool = False, plugin_opts: Optional[Dict[str, Any]] = None) -> str:
    labels = [str(x) for x in labels]
    clean_datasets = []
    for ds in datasets:
        clean = dict(ds)
        clean["data"] = [_clean_num(v) for v in clean.get("data", [])]
        clean.setdefault("borderWidth", 2)
        if (clean.get("type") or chart_type) == "line":
            clean.setdefault("tension", 0.25)
            clean.setdefault("fill", False)
            clean.setdefault("pointRadius", 3)
        if (clean.get("type") or chart_type) == "bar":
            clean.setdefault("borderRadius", 4)
        clean_datasets.append(clean)
    scales = {"x": {"stacked": stacked}, "y": {"stacked": stacked, "beginAtZero": chart_type == "bar" or stacked}}
    if y1:
        scales["y1"] = {"position": "right", "grid": {"drawOnChartArea": False}, "beginAtZero": True}
        if plugin_opts and "marketVerdictStrip" in plugin_opts:
            ratios = [v for ds in clean_datasets if ds.get("yAxisID") == "y1" and not ds.get("hidden") for v in ds.get("data", []) if v is not None]
            if ratios and 0 <= min(ratios) and max(ratios) <= 100:
                lo = max(0, int(min(ratios) // 5) * 5 - 5)
                hi = min(100, -int(-max(ratios) // 5) * 5 + 10)
                scales["y1"].update({"beginAtZero": False, "min": lo, "max": hi})
    cfg = {
        "type": chart_type,
        "data": {"labels": labels, "datasets": clean_datasets},
        "options": {
            "responsive": True,
            "maintainAspectRatio": False,
            "interaction": {"mode": "index", "intersect": False},
            "plugins": {
                "legend": {"position": "bottom"},
                "title": {"display": False}
            },
            "scales": scales
        }
    }
    if plugin_opts:
        cfg["options"]["plugins"].update(plugin_opts)
        top_pad = 0
        if "marketVerdictStrip" in plugin_opts:
            # 상단 판정 스트립(색점·라벨)이 꺾은선과 겹치지 않도록 여백 확보
            top_pad += 22
        if "policyEventLines" in plugin_opts:
            # 정책명 라벨 띠(최대 3줄)를 그림 영역 위에 따로 두어 막대·꺾은선을 가리지 않게 한다.
            top_pad += 46
        if top_pad:
            cfg["options"]["layout"] = {"padding": {"top": top_pad}}
    if chart_type in ["doughnut", "pie"]:
        cfg["options"].pop("scales", None)
        cfg["options"]["cutout"] = "66%" if chart_type == "doughnut" else None
    mix_cls = " mix-chart" if y1 else ""
    if plugin_opts and plugin_opts.get("policyEventLines"):
        mix_cls += " has-policy-lines"
    return f"""
    <div class='chart-box{mix_cls}'>
      <div class='chart-subtitle'>{title}</div>
      <div class='chart-wrap'><canvas id='{chart_id}'></canvas></div>
      <script>registerChart('{chart_id}', {js_json(cfg)});</script>
      {f"<div class='chart-note'>{note}</div>" if note else ""}
    </div>
    """

def mini_chart_block(title: str, chart_type: str, labels: List[Any], datasets: List[Dict[str, Any]], chart_id: str) -> str:
    labels = [str(x) for x in labels]
    clean_datasets = []
    for ds in datasets:
        clean = dict(ds)
        clean["data"] = [_clean_num(v) for v in clean.get("data", [])]
        clean.setdefault("borderWidth", 1.5)
        clean.setdefault("borderRadius", 4)
        clean_datasets.append(clean)
    cfg = {
        "type": chart_type,
        "data": {"labels": labels, "datasets": clean_datasets},
        "options": {
            "responsive": True,
            "maintainAspectRatio": False,
            "plugins": {"legend": {"display": False}},
            "scales": {"y": {"beginAtZero": True}},
            "datasets": {"bar": {"barPercentage": 0.72, "categoryPercentage": 0.72}}
        }
    }
    return f"""
    <div class='mini-chart-card'>
      <div class='mini-chart-title'>{title}</div>
      <div class='mini-chart-box'><canvas id='{chart_id}'></canvas></div>
      <script>registerChart('{chart_id}', {js_json(cfg)});</script>
    </div>
    """

def mini_chart_from_rows(rows: List[Dict[str, Any]], title: str, chart_id: str, x: str, y_key: str, max_rows: int = 10) -> str:
    if not rows:
        return mini_chart_block(title, "bar", ["자료 없음"], [{"label": title, "data": [0]}], chart_id)
    rows = rows[-max_rows:]
    return mini_chart_block(title, "bar", [r.get(x, "") for r in rows], [{"label": title, "data": [r.get(y_key) for r in rows]}], chart_id)



def pick_val(row: Dict[str, Any], aliases: List[str], default: Any = None) -> Any:
    for k in aliases:
        if k in row and row.get(k) not in [None, ""]:
            return row.get(k)
    return default

def normalize_age_move_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """연령대별 이동자수 원자료 컬럼명이 달라도 차트가 비지 않도록 표준화."""
    out = []
    for r in rows or []:
        age = pick_val(r, ["연령대", "age_group", "age", "구분", "항목", "category", "type"], "")
        in_v = pick_val(r, ["전입", "in", "in_count", "move_in", "전입자수"], 0)
        out_v = pick_val(r, ["전출", "out", "out_count", "move_out", "전출자수"], 0)
        net_v = pick_val(r, ["순이동", "net", "net_move", "순이동자수", "증감"], None)
        if net_v in [None, ""]:
            try:
                net_v = float(in_v or 0) - float(out_v or 0)
            except Exception:
                net_v = 0
        cum_v = pick_val(r, ["13개월 누적 순이동", "cum_net", "cumulative_net", "누적순이동"], net_v)
        out.append({"연령대": age, "전입": in_v, "전출": out_v, "순이동": net_v, "13개월 누적 순이동": cum_v})
    return out

def split_age_move_rows(rows: List[Dict[str, Any]], mode: str) -> List[Dict[str, Any]]:
    norm = normalize_age_move_rows(rows)
    if not norm:
        return []
    if mode == "young":
        keys = ["0~9세","10대","10~14세","15~19세","20대","20~24세","25~29세","30대","30~34세","35~39세"]
        return [r for r in norm if any(k in str(r.get("연령대","")) for k in keys)] or norm[:8]
    if mode == "middle_old":
        keys = ["40대","50대","60대","70대","80대","90대","100세","70대 이상"]
        return [r for r in norm if any(k in str(r.get("연령대","")) for k in keys)] or norm[-8:]
    return norm

AGE_GROUP_ORDER = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
AGE_COLOR_MAP = {
    "0~9세": "#5BC0EB",
    "10대": "#3A86FF",
    "20대": "#8338EC",
    "30대": "#2A9D8F",
    "40대": "#F4A261",
    "50대": "#E76F51",
    "60대": "#577590",
    "70대 이상": "#D62828",
}

def _age_color_arrays(labels: List[str]) -> Dict[str, List[str]]:
    borders, backgrounds, softs = [], [], []
    for label in labels:
        base = AGE_COLOR_MAP.get(str(label), "#94a3b8")
        borders.append(base)
        backgrounds.append(base + "cc")
        softs.append(base + "33")
    return {"border": borders, "background": backgrounds, "soft": softs}


def chart_age_bar_by_group(rows: List[Dict[str, Any]], title: str, chart_id: str, value_key: str) -> str:
    """연령대별 막대차트: 동일 연령대는 모든 연령 차트에서 같은 색상 적용."""
    if not rows:
        labels = ["자료 없음"]
        values = [0]
    else:
        # 10세 구간 표준 순서로 재집계
        agg = {g: 0.0 for g in AGE_GROUP_ORDER}
        usable = False
        for r in rows:
            age_val = r.get("연령대") or r.get("age") or r.get("age_group")
            g = age_to_group(age_val)
            if g not in agg:
                continue
            try:
                agg[g] += float(r.get(value_key) or 0)
                usable = True
            except Exception:
                pass
        labels = AGE_GROUP_ORDER[:]
        values = [agg[g] for g in labels]
        if not usable:
            labels = ["자료 없음"]
            values = [0]
    colors = _age_color_arrays(labels)
    cfg = {
        "type": "bar",
        "data": {
            "labels": labels,
            "datasets": [{
                "label": value_key,
                "data": values,
                "backgroundColor": colors["background"],
                "borderColor": colors["border"],
                "borderWidth": 1.5,
                "borderRadius": 4
            }]
        },
        "options": {
            "responsive": True,
            "maintainAspectRatio": False,
            "interaction": {"mode": "index", "intersect": False},
            "layout": {"padding": {"top": 6, "right": 10, "bottom": 22, "left": 10}},
            "plugins": {
                "legend": {"position": "bottom", "labels": {"boxWidth": 12, "padding": 16}},
                "title": {"display": False}
            },
            "scales": {
                "x": {"grid": {"color": "rgba(148,163,184,.16)"}},
                "y": {"beginAtZero": False, "grid": {"color": "rgba(148,163,184,.18)"}}
            }
        }
    }
    return f"""
    <div class='chart-box age-color-chart'>
      <div class='chart-subtitle'>{title}</div>
      <div class='chart-wrap'><canvas id='{chart_id}'></canvas></div>
      <script>registerChart('{chart_id}', {js_json(cfg)});</script>
    </div>
    """

def chart_age_net_combined(rows: List[Dict[str, Any]], chart_id: str) -> str:
    """최근 13개월 누적순이동 연령별 비교.
    age_monthly_rows가 있으면 13개월 순이동을 10세 구간별로 누적합산한다.
    """
    agg = {g: 0.0 for g in AGE_GROUP_ORDER}
    usable = False
    for r in rows or []:
        # age_monthly_rows: {"month", "age", "net"} / normalize rows: {"연령대","순이동"}
        age_val = r.get("age") or r.get("age_group") or r.get("연령대")
        g = age_to_group(age_val)
        if g not in agg:
            continue
        net_val = r.get("net")
        if net_val is None:
            net_val = r.get("13개월 누적 순이동", r.get("순이동", 0))
        try:
            agg[g] += float(net_val or 0)
            usable = True
        except Exception:
            pass

    labels = AGE_GROUP_ORDER[:]
    values = [agg[g] for g in labels]
    if not usable:
        labels = ["자료 없음"]
        values = [0]

    colors = _age_color_arrays(labels)
    cfg = {
        "type": "bar",
        "data": {
            "labels": labels,
            "datasets": [{
                "label": "누적순이동(명)",
                "data": values,
                "backgroundColor": colors["background"],
                "borderColor": colors["border"],
                "borderWidth": 1.5,
                "borderRadius": 4
            }]
        },
        "options": {
            "indexAxis": "y",
            "responsive": True,
            "maintainAspectRatio": False,
            "interaction": {"mode": "index", "intersect": False},
            "plugins": {
                "legend": {"display": False},
                "title": {"display": False}
            },
            "scales": {
                "x": {"beginAtZero": True, "title": {"display": True, "text": "누적순이동(명)"}},
                "y": {"beginAtZero": False}
            }
        }
    }
    return f"""
    <div class='chart-box mix-chart age-net-horizontal-chart'>
      <div class='chart-subtitle'>최근 13개월 누적순이동 연령별 비교</div>
      <div class='chart-wrap'><canvas id='{chart_id}'></canvas></div>
      <script>registerChart('{chart_id}', {js_json(cfg)});</script>
      <div class='chart-note'>최근 13개월 연령대별 순이동을 10세 구간으로 재집계하고, 최신월 연령대별 인구 차트와 동일한 색상 팔레트를 적용함.</div>
    </div>
    """


def chart_from_rows(rows: List[Dict[str, Any]], title: str, chart_id: str, x: str, y_keys: List[str], chart_type: str = "line", note: str = "", max_rows: int = 13, stacked: bool = False) -> str:
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = rows[-max_rows:]
    if x == "dong":
        labels = [compose_sigungu_dong_label(r.get("sigungu"), r.get("dong", "")) for r in rows]
    else:
        labels = [r.get(x, "") for r in rows]
    datasets = []
    for k in y_keys:
        datasets.append({"label": display_col_name(k), "data": [r.get(k) for r in rows]})
    return chart_block(title, chart_type, labels, datasets, chart_id, note, stacked=stacked)

def _short_month_label(m: Any) -> str:
    s = str(m or "")
    mt = re.search(r"(20\d{2})[-./]?(\d{2})", s)
    if mt:
        return f"{mt.group(1)[2:]}.{mt.group(2)}"
    return s


def _axis_minmax(values: List[Any], pad_ratio: float = 0.12, min_pad: float = 50) -> Dict[str, float]:
    nums = []
    for v in values or []:
        try:
            if v is not None and not pd.isna(v):
                nums.append(float(v))
        except Exception:
            pass
    if not nums:
        return {}
    mn, mx = min(nums), max(nums)
    span = mx - mn
    pad = max(span * pad_ratio, min_pad)
    if span == 0:
        pad = max(abs(mx) * 0.001, min_pad)
    return {"min": max(0, mn - pad), "max": mx + pad}

def chart_population_combo(rows: List[Dict[str, Any]], chart_id: str) -> str:
    """최근 13개월 인구수·세대수 추이.
    v5.8.38: 인구 변화가 잘 보이도록 y축을 0 기준이 아닌 데이터 범위 중심으로 확대 표시.
    """
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = rows[-13:]
    labels = [_short_month_label(r.get("month","")) for r in rows]
    pop_values = [r.get("population") for r in rows]
    hh_values = [r.get("households") for r in rows]
    pop_axis = _axis_minmax(pop_values, pad_ratio=0.18, min_pad=120)
    hh_axis = _axis_minmax(hh_values, pad_ratio=0.18, min_pad=60)
    cfg = {
        "type": "bar",
        "data": {
            "labels": labels,
            "datasets": [
                {
                    "type": "line",
                    "label": "총인구수",
                    "data": pop_values,
                    "yAxisID": "y",
                    "order": 1,
                    "borderColor": "#475569",
                    "backgroundColor": "#47556933",
                    "pointBackgroundColor": "#475569",
                    "pointBorderColor": "#475569",
                    "borderWidth": 3.2,
                    "pointRadius": 5,
                    "pointHoverRadius": 6,
                    "tension": 0.25,
                    "fill": False
                },
                {
                    "type": "bar",
                    "label": "총세대수",
                    "data": hh_values,
                    "yAxisID": "y1",
                    "order": 2,
                    "backgroundColor": "#7ee7bdcc",
                    "borderColor": "#10b981",
                    "borderWidth": 1.4,
                    "borderRadius": 4,
                    "barPercentage": 0.68,
                    "categoryPercentage": 0.72
                }
            ]
        },
        "options": {
            "responsive": True,
            "maintainAspectRatio": False,
            "interaction": {"mode": "index", "intersect": False},
            "plugins": {
                "legend": {"position": "bottom", "labels": {"boxWidth": 12, "padding": 12}},
                "title": {"display": False}
            },
            "layout": {"padding": {"bottom": 12, "top": 4}},
            "scales": {
                "y": {
                    "position": "left",
                    "title": {"display": True, "text": "인구수(명)"},
                    "grid": {"color": "rgba(148,163,184,.18)"},
                    **pop_axis
                },
                "y1": {
                    "position": "right",
                    "title": {"display": True, "text": "세대수(세대)"},
                    "grid": {"drawOnChartArea": False},
                    "beginAtZero": True
                },
                "x": {
                    "grid": {"color": "rgba(148,163,184,.16)"}
                }
            }
        }
    }
    return f"""
    <div class='chart-box mix-chart population-mix-chart population-household-chart'>
      <div class='chart-subtitle'>최근 13개월 인구수·세대수 추이(인구: 선, 세대: 막대, 이중축)</div>
      <div class='chart-wrap'><canvas id='{chart_id}'></canvas></div>
      <script>registerChart('{chart_id}', {js_json(cfg)});</script>
      <div class='chart-note'>인구수는 변화 감지를 위해 좌축을 확대 표시하고, 세대수 막대는 실제 규모 왜곡을 줄이기 위해 우축 0 기준으로 표시함.</div>
    </div>
    """


def chart_migration_combo(rows: List[Dict[str, Any]], chart_id: str) -> str:
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = rows[-13:]
    labels = [r.get("month","") for r in rows]
    datasets = [
        {"type":"bar", "label":"총 전입", "data":[_clean_num(r.get("in")) for r in rows], "yAxisID":"y", "backgroundColor":"#ef4444cc", "borderColor":"#dc2626", "borderWidth":1.2, "borderRadius":4},
        {"type":"bar", "label":"총 전출", "data":[_clean_num(r.get("out")) for r in rows], "yAxisID":"y", "backgroundColor":"#3b82f6cc", "borderColor":"#2563eb", "borderWidth":1.2, "borderRadius":4},
        {"type":"line", "label":"순이동", "data":[_clean_num(r.get("net")) for r in rows], "yAxisID":"y", "borderColor":"#111827", "backgroundColor":"#11182733", "pointBackgroundColor":"#111827", "borderWidth":3, "pointRadius":3.5, "tension":0.25}
    ]
    cfg = {
        "type": "bar",
        "data": {"labels": labels, "datasets": datasets},
        "options": {
            "responsive": True,
            "maintainAspectRatio": False,
            "interaction": {"mode": "index", "intersect": False},
            "layout": {"padding": {"bottom": 26, "top": 4}},
            "plugins": {
                "legend": {"position": "bottom", "labels": {"boxWidth": 11, "padding": 16}},
                "title": {"display": False}
            },
            "scales": {
                "x": {"grid": {"color": "rgba(148,163,184,.16)"}},
                "y": {"beginAtZero": False, "grid": {"color": "rgba(148,163,184,.18)"}}
            }
        }
    }
    return f"""
    <div class='chart-box migration-combo-chart'>
      <div class='chart-subtitle'>최근 13개월 전입·전출·순이동 추이(전입·전출 막대, 순이동 선)</div>
      <div class='chart-wrap'><canvas id='{chart_id}'></canvas></div>
      <script>registerChart('{chart_id}', {js_json(cfg)});</script>
    </div>
    """


def _direction_counts_from_monthly_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """월별 compare_count와 상승·보합·하락 비율을 이용해 방향별 거래건수를 산출."""
    out = []
    for r in rows or []:
        compare_count = _clean_num(r.get("compare_count"))
        # compare_count가 없으면 trade_count를 보조 분모로 사용
        if not compare_count:
            compare_count = _clean_num(r.get("trade_count"))
        up = round(compare_count * _clean_num(r.get("up_ratio")) / 100, 1) if compare_count else 0
        flat = round(compare_count * _clean_num(r.get("flat_ratio")) / 100, 1) if compare_count else 0
        down = round(compare_count * _clean_num(r.get("down_ratio")) / 100, 1) if compare_count else 0
        item = dict(r)
        item["_up_count"] = up
        item["_flat_count"] = flat
        item["_down_count"] = down
        out.append(item)
    return out

def chart_monthly_price_change(rows: List[Dict[str, Any]], title: str, chart_id: str, label_prefix: str = "", events: Optional[List[Dict[str, Any]]] = None) -> str:
    """월간 가격변동 통합 차트.
    요청사항 반영: 꺾은선 제거, 상승·보합·하락 막대만 컬러로 표시.
    """
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = _direction_counts_from_monthly_rows(rows[-13:])
    labels = [r.get("month","") for r in rows]
    datasets = [
        {"type":"bar", "label":"상승", "data":[r.get("_up_count") for r in rows], "stack":"direction", "order":3},
        {"type":"bar", "label":"보합", "data":[r.get("_flat_count") for r in rows], "stack":"direction", "order":3},
        {"type":"bar", "label":"하락", "data":[r.get("_down_count") for r in rows], "stack":"direction", "order":3},
    ]
    return chart_block(
        title,
        "bar",
        labels,
        datasets,
        chart_id,
        "요청사항 반영: 평균 변동률 꺾은선은 제거하고, 상승·보합·하락 건수를 의미색상 막대로 표시함."
        + (" 주황 세로 점선은 정부 정책 발표·시행 시점임." if events else ""),
        stacked=True,
        y1=False,
        plugin_opts={"policyEventLines": {"events": events}} if events else None
    )


def expand_non_apt_type_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """비아파트 물건유형을 연립/다세대/단독/다가구 4개로 표준화.
    원자료가 연립다세대·단독다가구 파일 단위로만 제공되는 경우에는 합산명을 그대로 쓰지 않고
    4개 유형 슬롯을 유지하기 위해 결측 유형은 0으로 채운다.
    """
    out = []
    for r in rows or []:
        t = str(r.get("type", ""))
        m = r.get("month")
        v = r.get("trade_count", 0)
        if t in ["연립", "다세대", "단독", "다가구"]:
            out.append({"month": m, "type": t, "trade_count": v})
        elif "연립" in t and "다세대" in t:
            # 파일명만 결합형인 경우, 실제 세부유형 컬럼이 없으므로 결합값을 '연립'에 우선 배치하고 다세대 슬롯은 0 유지
            # 단, 차트에는 4개 유형이 모두 표시되도록 아래에서 0 슬롯을 보강함
            out.append({"month": m, "type": "연립", "trade_count": v})
        elif "단독" in t and "다가구" in t:
            out.append({"month": m, "type": "단독", "trade_count": v})
        elif "연립" in t:
            out.append({"month": m, "type": "연립", "trade_count": v})
        elif "다세대" in t:
            out.append({"month": m, "type": "다세대", "trade_count": v})
        elif "단독" in t:
            out.append({"month": m, "type": "단독", "trade_count": v})
        elif "다가구" in t:
            out.append({"month": m, "type": "다가구", "trade_count": v})
    return out

def chart_non_apt_monthly_by_type(rows: List[Dict[str, Any]], title: str, chart_id: str) -> str:
    """비아파트 차트는 단순 합계가 아니라 연립·다세대·단독·다가구 4개 물건유형별로 표시."""
    rows = expand_non_apt_type_rows(rows)
    if not rows:
        labels = ["자료 없음"]
        datasets = [{"label": t, "data": [0]} for t in ["연립", "다세대", "단독", "다가구"]]
        return chart_block(title, "bar", labels, datasets, chart_id, "비아파트 물건유형별 원자료 없음")
    months = sorted({str(r.get("month","")) for r in rows if r.get("month")})[-13:]
    type_order = ["연립", "다세대", "단독", "다가구"]
    datasets = []
    for t in type_order:
        vals = []
        for m in months:
            vals.append(sum(float(r.get("trade_count") or 0) for r in rows if str(r.get("month")) == m and str(r.get("type")) == t))
        datasets.append({"label": t, "data": vals})
    return chart_block(title, "bar", months, datasets, chart_id, "template_v58 보정: 비아파트 거래는 연립·다세대·단독·다가구 물건유형별 막대차트로 표시함.")


def chart_monthly_strength(rows: List[Dict[str, Any]], title: str, chart_id: str, label_prefix: str = "", reference: Optional[Dict[str, Any]] = None, events: Optional[List[Dict[str, Any]]] = None) -> str:
    """거래량 및 상승·하락 비율 추이.
    요청사항 반영: 거래량은 막대, 상승·하락 비율만 꺾은선으로 표시. 보합 꺾은선은 표시하지 않음.
    """
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = rows[-13:]
    labels = [r.get("month","") for r in rows]
    prefix = (label_prefix + " ") if label_prefix else ""
    # 각 월의 상승·보합·하락 비율 판정(상승우세/하락우세/보합우세/혼조/산출불가)을
    # 차트 상단 '판정 스트립'으로 함께 표시한다(marketVerdictStrip 플러그인).
    verdicts = [r.get("판정") or direction_verdict(r.get("up_ratio"), r.get("flat_ratio"), r.get("down_ratio")) for r in rows]
    datasets = [
        {"type":"bar", "label":f"{prefix}거래량(건)", "data":[r.get("trade_count") for r in rows], "yAxisID":"y", "order":3, "borderRadius":4},
        {"type":"line", "label":"상승", "data":[r.get("up_ratio") for r in rows], "yAxisID":"y1", "order":1, "borderWidth":3, "pointRadius":3},
        {"type":"line", "label":"하락", "data":[r.get("down_ratio") for r in rows], "yAxisID":"y1", "order":2, "borderWidth":3, "pointRadius":3},
    ]
    # 1년 전·2년 전 월평균 거래건수를 가로 점선으로 겹쳐 전년·2년 전 대비 거래량 수준을 바로 비교하게 한다.
    reference = reference or {}
    ref_note = ""
    for key, name, color in (("prev_year", "1년 전", "#d97706"), ("two_years", "2년 전", "#7c3aed")):
        ref = reference.get(key) or {}
        if ref.get("avg"):
            datasets.append({"type": "line", "label": f"{name} 월평균 {ref['avg']:,.0f}건", "data": [ref["avg"]] * len(rows),
                             "yAxisID": "y", "order": 0, "borderColor": color, "backgroundColor": color,
                             "borderDash": [6, 4], "borderWidth": 2, "pointRadius": 0, "tension": 0})
    recent = reference.get("recent") or {}
    if recent.get("avg"):
        parts = []
        for key, pct_key, name in (("prev_year", "vs_prev_year_pct", "1년 전"), ("two_years", "vs_two_years_pct", "2년 전")):
            ref, pct = reference.get(key) or {}, reference.get(pct_key)
            if ref.get("avg") and pct is not None:
                word = "증가" if pct > 0 else ("감소" if pct < 0 else "보합")
                parts.append(f"{name}({ref['label']}, {ref['months']}개월) {ref['avg']:,.1f}건 대비 {abs(pct):.1f}% {word}")
        if parts:
            ref_note = (f" <b>월평균 거래건수 비교</b>: 최근 12개월({recent['label']}) {recent['avg']:,.1f}건은 " + ", ".join(parts)
                        + "함(기준월은 신고기한 미도래로 제외). 점선은 1년 전·2년 전 월평균임.")
    base_v = verdicts[-1] if verdicts else ""
    base_emoji = VERDICT_EMOJI.get(base_v, "")
    plugin_opts = {"marketVerdictStrip": {"verdicts": verdicts}}
    if events:
        plugin_opts["policyEventLines"] = {"events": events}
    return chart_block(
        title,
        "bar",
        labels,
        datasets,
        chart_id,
        f"기준월 판정: {base_emoji} {base_v}. 차트 상단 색점·라벨이 월별 판정(🔴상승우세·🔵하락우세·⚪보합우세·⚫산출불가)임. 거래량은 막대, 상승·하락 비율은 꺾은선."
        + (" 주황 세로 점선은 정부 정책 발표·시행 시점임." if events else "") + ref_note,
        y1=True,
        plugin_opts=plugin_opts
    )


def chart_direction_rows(rows: List[Dict[str, Any]], title: str, chart_id: str, x: str = "band", max_rows: int = 12) -> str:
    """v5.8.54: 누적막대 + '산출불가' 4번째 데이터셋.

    신규계약 0건이라 up/flat/down 비율이 모두 0%인 행도 거래건수만큼 별도
    색상(앰버) 막대로 표시해, 양림동 같은 케이스가 차트에서 누락되어
    보이지 않게 되는 문제를 해결한다. 보합과 명확히 다른 색상으로 구분해
    "시장강도 산출 불가" 사유를 시각적으로 분리한다.
    """
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = rows[:max_rows]
    if x == "dong":
        labels = [compose_sigungu_dong_label(r.get("sigungu"), r.get("dong", "")) for r in rows]
    else:
        labels = [r.get(x, "") for r in rows]
    def cnt(r, key):
        try:
            return float(r.get("count") or 0) * float(r.get(key) or 0) / 100
        except Exception:
            return 0
    def _total_count(r):
        try:
            return float(r.get("count") or 0)
        except Exception:
            return 0.0
    def _ratio_sum(r):
        s = 0.0
        for k in ("up_ratio", "flat_ratio", "down_ratio"):
            try:
                s += float(r.get(k) or 0)
            except Exception:
                pass
        return s
    def _is_unavailable(r):
        # 거래는 있는데(신규 0 등) up/flat/down 비율이 모두 0인 행.
        # 비교가능건이 명시적으로 0 또는 신규계약이 0이면 산출 불가로 분류.
        try:
            new_c = float(r.get("new_count") or 0)
        except Exception:
            new_c = 0.0
        try:
            comp_c = float(r.get("compare_count") or 0) if r.get("compare_count") is not None else None
        except Exception:
            comp_c = None
        # 명시적 신호: 신규 0건 또는 비교가능건 0
        if _total_count(r) > 0 and _ratio_sum(r) == 0 and (new_c == 0 or comp_c == 0):
            return True
        return False
    up_data = []
    flat_data = []
    down_data = []
    unavail_data = []
    for r in rows:
        if _is_unavailable(r):
            up_data.append(0)
            flat_data.append(0)
            down_data.append(0)
            unavail_data.append(_total_count(r))
        else:
            up_data.append(cnt(r, "up_ratio"))
            flat_data.append(cnt(r, "flat_ratio"))
            down_data.append(cnt(r, "down_ratio"))
            unavail_data.append(0)
    datasets = [
        {"label": "상승", "data": up_data},
        {"label": "보합", "data": flat_data},
        {"label": "하락", "data": down_data},
        {"label": "산출불가", "data": unavail_data},
    ]
    # 각 항목(읍면동·구간)별 판정(상승우세/하락우세/보합우세/산출불가)을 막대 상단 색점·라벨로 표시
    verdicts = [
        "산출불가" if _is_unavailable(r)
        else (r.get("판정") or direction_verdict(r.get("up_ratio"), r.get("flat_ratio"), r.get("down_ratio")))
        for r in rows
    ]
    note = "template_v58 기준: 상승·보합·하락 의미색상 고정 누적막대. v5.8.54: 신규 0건/비교가능건 0 인 동은 '산출불가'(앰버)로 거래건수 표시. 막대 상단 색점·라벨은 항목별 판정(상승우세/하락우세/보합우세/산출불가)."
    return chart_block(title, "bar", labels, datasets, chart_id, note, stacked=True,
                       plugin_opts={"marketVerdictStrip": {"verdicts": verdicts}})

def _dominant_direction(up: Any, flat: Any, down: Any) -> str:
    vals = [("상승", _clean_num(up)), ("보합", _clean_num(flat)), ("하락", _clean_num(down))]
    vals = [(label, 0.0 if value is None else value) for label, value in vals]
    if not any(value != 0.0 for _, value in vals):
        return "보합"
    vals = sorted(vals, key=lambda x: x[1], reverse=True)
    return vals[0][0] if vals else "보합"

# === 상승·보합·하락 비율 기반 시장 방향 판정 ===
# 상승비율/하락비율/보합비율 양식·차트에 추가하는 "판정"란의 값을 산출한다.
VERDICT_EMOJI = {"상승우세": "🔴", "보합우세": "⚪", "하락우세": "🔵", "산출불가": "⚫"}
VERDICT_CLASS = {"상승우세": "dir-up", "보합우세": "dir-flat", "하락우세": "dir-down", "산출불가": "dir-na"}

def direction_verdict(up: Any, flat: Any, down: Any) -> str:
    """상승·보합·하락 비율로 시장 방향을 판정한다.

    반환값: 상승우세 / 보합우세 / 하락우세 / 산출불가
    - 세 비율이 모두 0이면 비교가능건이 없는 것으로 보아 '산출불가'.
    - 그 외에는 가장 비중이 큰 방향을 'OO우세'로 판정(혼조 없이 비율값으로 결정).
    """
    u = _clean_num(up) or 0.0
    f = _clean_num(flat) or 0.0
    d = _clean_num(down) or 0.0
    if u == 0.0 and f == 0.0 and d == 0.0:
        return "산출불가"
    dom = _dominant_direction(u, f, d)
    return {"상승": "상승우세", "보합": "보합우세", "하락": "하락우세"}.get(dom, "보합우세")

def doughnut_market_strength(title: str, up: Any, flat: Any, down: Any, chart_id: str, compare_count: Any = None) -> str:
    upv, flatv, downv = _clean_num(up), _clean_num(flat), _clean_num(down)
    upv = 0.0 if upv is None else upv
    flatv = 0.0 if flatv is None else flatv
    downv = 0.0 if downv is None else downv
    compare_s = format_num(compare_count) if compare_count not in [None, ""] else "-"
    dominant = _dominant_direction(upv, flatv, downv)
    center_text = f"{dominant} 우세" if dominant in ["상승", "하락"] else "보합 우세"
    sub_text = f"{compare_s}건" if compare_s != "-" else "비교가능"
    # 조각 안에 비율과 건수(비교가능건 × 비율)를 함께 적는다.
    total_n = _clean_num(compare_count)
    slice_counts = [int(round(total_n * v / 100)) if total_n else None for v in (upv, flatv, downv)]
    cfg = {
        "type": "doughnut",
        "data": {
            "labels": ["상승", "보합", "하락"],
            "datasets": [{"label": title, "data": [upv, flatv, downv], "borderWidth": 1}]
        },
        "options": {
            "cutout": "54%",
            "responsive": True,
            "maintainAspectRatio": False,
            "plugins": {
                "legend": {"position": "bottom"},
                "title": {"display": False},
                "centerText": {"text": center_text, "subtext": sub_text},
                "doughnutSliceLabels": {"counts": slice_counts}
            }
        }
    }
    return f"""
    <div class='survey-donut-card'>
      <div class='survey-donut-title'>{title}</div>
      <div class='survey-donut-meta'>비교 가능 {compare_s}건</div>
      <div class='survey-donut-box'><canvas id='{chart_id}'></canvas></div>
      <script>registerChart('{chart_id}', {js_json(cfg)});</script>
      <div class='survey-donut-ratios'>
        <span class='dir-up'>상승 {upv:.1f}%{f"({slice_counts[0]}건)" if slice_counts[0] is not None else ""}</span>
        <span class='dir-flat'>보합 {flatv:.1f}%{f"({slice_counts[1]}건)" if slice_counts[1] is not None else ""}</span>
        <span class='dir-down'>하락 {downv:.1f}%{f"({slice_counts[2]}건)" if slice_counts[2] is not None else ""}</span>
      </div>
      <div class='survey-donut-caption'>조사 판단: {center_text}</div>
    </div>
    """

def chart_grouped_by_type(rows: List[Dict[str, Any]], title: str, chart_id: str, note: str = "") -> str:
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    filtered = []
    for r in rows:
        t = str(r.get("type", ""))
        if any(skip in t for skip in ["합계", "계(", "동수", "가구수"]):
            continue
        filtered.append(r)
    rows = filtered or rows
    labels = sorted(list({str(r.get("month", "")) for r in rows if r.get("month")}))
    types = sorted(list({str(r.get("type", "")) for r in rows if r.get("type")}))[:8]
    datasets = []
    for t in types:
        datasets.append({
            "label": t,
            "data": [sum(float(r.get("value") or 0) for r in rows if str(r.get("month")) == m and str(r.get("type")) == t) for m in labels]
        })
    return chart_block(title, "bar", labels, datasets, chart_id, note, stacked=False)

def chart_grouped_by_type_filtered(rows: List[Dict[str, Any]], title: str, chart_id: str, type_keyword: str) -> str:
    stage_color = _supply_stage_color_by_key(title)
    if not rows:
        return mini_chart_block(title, "bar", ["자료 없음"], [{"label": title, "data": [0], "backgroundColor": stage_color["bg"], "borderColor": stage_color["border"]}], chart_id)
    filtered = [r for r in rows if type_keyword in str(r.get("type", ""))]
    if not filtered:
        labels = sorted(list({str(_supply_year_value(r)) for r in rows if r.get("month")}))[-10:] or ["자료 없음"]
        return mini_chart_block(title, "bar", labels, [{"label": title, "data": [0 for _ in labels], "backgroundColor": stage_color["bg"], "borderColor": stage_color["border"]}], chart_id)
    latest = {}
    for r in filtered:
        y = _supply_year_value(r)
        if y not in latest or _supply_month_sort_key(r) >= _supply_month_sort_key(latest[y]):
            latest[y] = r
    years = sorted(latest.keys())[-10:]
    data = [float(latest[y].get("value") or 0) for y in years]
    return mini_chart_block(title, "bar", [_supply_period_label(latest[y]) for y in years], [{"label": title, "data": data, "backgroundColor": stage_color["bg"], "borderColor": stage_color["border"]}], chart_id)


def make_chart_html(data: Dict[str, Any]) -> Dict[str, str]:
    """template_v58(7).html의 58개 canvas 슬롯만 사용한다.
    템플릿에 없는 차트는 생성하지 않고, 기존 슬롯에 데이터만 주입한다."""
    charts = {}
    pop = data.get("population", {}).get("table_rows", [])
    mig = data.get("migration", {}).get("table_rows", [])
    age_pop = data.get("age_population", {})
    age_move = data.get("migration", {}).get("age_monthly_rows", []) or normalize_age_move_rows(data.get("migration", {}).get("age_rows", []))
    rates = data.get("rates", {}).get("table_rows", [])
    unsold = data.get("unsold", {}).get("table_rows", [])
    supply = data.get("supply", {}).get("table_rows", [])
    apt = data.get("apt_trade", {})
    ar = data.get("apt_rent", {})
    nar = data.get("non_apt_rent", {})
    nas = data.get("non_apt_sale", {})
    presale = data.get("presale", {})

    is_weekly = bool(data.get("meta", {}).get("is_weekly"))
    charts["core_donuts"] = (
        "<div class='survey-donut-section'><div class='survey-donut-grid'>"
        + doughnut_market_strength("매매 시장강도", apt.get("up_ratio"), apt.get("flat_ratio"), apt.get("down_ratio"), "surveySaleDonut", apt.get("compare_count"))
        + doughnut_market_strength("전세 시장강도", ar.get("jeonse", {}).get("up_ratio"), ar.get("jeonse", {}).get("flat_ratio"), ar.get("jeonse", {}).get("down_ratio"), "surveyJeonseDonut", ar.get("jeonse", {}).get("compare_count"))
        + ("" if is_weekly else doughnut_market_strength("월세 시장강도", ar.get("monthly_rent", {}).get("up_ratio"), ar.get("monthly_rent", {}).get("flat_ratio"), ar.get("monthly_rent", {}).get("down_ratio"), "surveyRentDonut", ar.get("monthly_rent", {}).get("compare_count")))
        + "</div></div>"
    )

    # Ⅰ. 인구·세대·이동·금리·미분양
    charts["population"] = chart_population_combo(pop, "chart1")
    charts["migration"] = chart_migration_combo(mig, "chart2")
    charts["age_population_latest"] = chart_age_bar_by_group(age_pop.get("age_rows", []), "최신월 연령대별 주민등록인구", "chartAgePopLatest", "최신월 인구")
    charts["age_population_change"] = chart_age_bar_by_group(age_pop.get("age_rows", []), "최근 13개월 연령대별 증감(최신월-13개월 시작월)", "chartAgePopChange", "13개월 증감")
    charts["rates"] = chart_from_rows(rates, "예금은행 대출금리 추세", "chart3", "month", ["mortgage_rate", "jeonse_rate", "sme_rate", "business_rate"], "line")
    charts["unsold"] = chart_from_rows(unsold, "미분양·준공후 미분양 추이", "chart4", "month", ["unsold_units", "after_completion_unsold"], "bar")

    # 공급: mini-chart-card 19개
    supply_region = data.get("meta", {}).get("supply_region", "공급권역")
    charts["supply_permit_total"] = mini_chart_from_rows_year_label(supply, f"{supply_region} 인허가 총량 시계열", "totalPermit", "month", "permit")
    charts["supply_start_total"] = mini_chart_from_rows_year_label(supply, f"{supply_region} 착공 총량 시계열", "totalStart", "month", "start")
    charts["supply_complete_total"] = mini_chart_from_rows_year_label(supply, f"{supply_region} 준공 총량 시계열", "totalComplete", "month", "completion")
    sup_obj = data.get("supply", {})
    for t in ["아파트", "연립", "다세대", "단독", "다가구"]:
        charts[f"supply_permit_{t}"] = chart_grouped_by_type_filtered(sup_obj.get("permit_by_type_rows", []), f"인허가 - {t}", f"permitChart_{t}", t)
        charts[f"supply_start_{t}"] = chart_grouped_by_type_filtered(sup_obj.get("start_by_type_rows", []), f"착공 - {t}", f"startChart_{t}", t)
        charts[f"supply_complete_{t}"] = chart_grouped_by_type_filtered(sup_obj.get("completion_by_type_rows", []), f"준공 - {t}", f"completeChart_{t}", t)

    # Ⅲ. 매매: template_v58의 9개 슬롯만 사용
    charts["apt_monthly"] = chart_monthly_price_change(apt.get("monthly_rows", []), "아파트 매매 거래량 및 가격변동 통합", "chart7", "매매", events=policy_chart_events(data, [r.get("month") for r in (apt.get("monthly_rows", []))[-13:]]))
    charts["apt_strength"] = chart_monthly_strength(apt.get("monthly_rows", []), "매매 거래량 및 상승·하락 비율 추이", "chart33", "매매", reference=apt.get("volume_reference"), events=policy_chart_events(data, [r.get("month") for r in (apt.get("monthly_rows", []))[-13:]]))
    charts["apt_dong"] = chart_direction_rows(apt.get("dong_rows", [])[:10], "당월 읍면동별 매매 시장강도", "chart11", "dong", max_rows=10)
    charts["apt_area_share"] = chart_direction_rows(apt.get("area_band_rows", []), "평형별 매매 거래 비중", "chart12", "band", max_rows=20)
    charts["apt_dong_detail"] = chart_direction_rows(apt.get("dong_rows", [])[:10], "아파트 매매 읍면동별 세부 현황", "chart30", "dong", max_rows=10)
    charts["apt_price_band"] = chart_direction_rows(apt.get("price_band_rows", []), "아파트 매매 거래금액별 세부 현황", "chart13", "band", max_rows=20)
    charts["apt_age_band"] = chart_direction_rows(apt.get("building_age_rows", []), "아파트 매매 연식별 세부 현황", "chart14", "band", max_rows=20)
    charts["non_apt_sale"] = chart_non_apt_monthly_by_type(nas.get("monthly_by_type_rows", []), "비아파트 물건유형별 최근 13개월 매매 거래 추이(연립·다세대·단독·다가구)", "chart15")
    charts["presale"] = chart_from_rows(presale.get("monthly_rows", []), "아파트 분양권전매 최근 13개월 거래 추이", "chart16", "month", ["trade_count"], "bar")

    # Ⅳ. 전세: template_v58의 8개 슬롯만 사용
    charts["apt_jeonse_monthly"] = chart_monthly_price_change(ar.get("jeonse", {}).get("monthly_rows", []), "아파트 전세 거래량 및 가격변동 통합", "chart8", "전세", events=policy_chart_events(data, [r.get("month") for r in (ar.get("jeonse", {}).get("monthly_rows", []))[-13:]]))
    charts["apt_jeonse_strength"] = chart_monthly_strength(ar.get("jeonse", {}).get("monthly_rows", []), "전세 거래량 및 상승·하락 비율 추이", "chart34", "전세", reference=ar.get("jeonse", {}).get("volume_reference"), events=policy_chart_events(data, [r.get("month") for r in (ar.get("jeonse", {}).get("monthly_rows", []))[-13:]]))
    charts["apt_jeonse_market_dong"] = chart_direction_rows(ar.get("jeonse", {}).get("dong_rows", [])[:10], "당월 읍면동별 전세 시장강도", "chart18", "dong", max_rows=10)
    charts["apt_jeonse_dong"] = chart_direction_rows(ar.get("jeonse", {}).get("dong_rows", [])[:10], "아파트 전세 읍면동별 세부 현황", "chart31", "dong", max_rows=10)
    charts["apt_jeonse_deposit"] = chart_direction_rows(ar.get("jeonse", {}).get("deposit_band_rows", []), "아파트 전세 보증금별 세부 현황", "chart19", "band", max_rows=20)
    charts["apt_jeonse_age"] = chart_direction_rows(ar.get("jeonse", {}).get("building_age_rows", []), "아파트 전세 연식별 세부 현황", "chart25", "band", max_rows=20)
    charts["apt_jeonse_area"] = chart_direction_rows(ar.get("jeonse", {}).get("area_band_rows", []), "아파트 전세 평형별 세부 현황", "chart26", "band", max_rows=20)
    charts["non_apt_jeonse"] = chart_non_apt_monthly_by_type(nar.get("jeonse", {}).get("monthly_by_type_rows", []), "비아파트 물건유형별 최근 13개월 전세 거래 추이(연립·다세대·단독·다가구)", "chart20")

    # Ⅴ. 월세: template_v58의 9개 슬롯만 사용
    charts["apt_rent_monthly"] = chart_monthly_price_change(ar.get("monthly_rent", {}).get("monthly_rows", []), "아파트 월세 거래량 및 가격변동 통합", "chart9", "월세", events=policy_chart_events(data, [r.get("month") for r in (ar.get("monthly_rent", {}).get("monthly_rows", []))[-13:]]))
    charts["apt_rent_strength"] = chart_monthly_strength(ar.get("monthly_rent", {}).get("monthly_rows", []), "월세 거래량 및 상승·하락 비율 추이", "chart35", "월세", reference=ar.get("monthly_rent", {}).get("volume_reference"), events=policy_chart_events(data, [r.get("month") for r in (ar.get("monthly_rent", {}).get("monthly_rows", []))[-13:]]))
    charts["apt_rent_market_dong"] = chart_direction_rows(ar.get("monthly_rent", {}).get("dong_rows", [])[:10], "당월 읍면동별 월세 시장강도", "chart22", "dong", max_rows=10)
    charts["apt_rent_dong"] = chart_direction_rows(ar.get("monthly_rent", {}).get("dong_rows", [])[:10], "아파트 월세 읍면동별 세부 현황", "chart32", "dong", max_rows=10)
    charts["apt_rent_monthly_band"] = chart_direction_rows(ar.get("monthly_rent", {}).get("monthly_band_rows", []), "아파트 월 임대료별 세부 현황", "chart23", "band", max_rows=20)
    charts["apt_rent_deposit"] = chart_direction_rows(ar.get("monthly_rent", {}).get("deposit_band_rows", []), "아파트 월세 보증금별 세부 현황", "chart27", "band", max_rows=20)
    charts["apt_rent_age"] = chart_direction_rows(ar.get("monthly_rent", {}).get("building_age_rows", []), "아파트 월세 연식별 세부 현황", "chart28", "band", max_rows=20)
    charts["apt_rent_area"] = chart_direction_rows(ar.get("monthly_rent", {}).get("area_band_rows", []), "아파트 월세 평형별 세부 현황", "chart29", "band", max_rows=20)
    charts["non_apt_rent"] = chart_non_apt_monthly_by_type(nar.get("monthly_rent", {}).get("monthly_by_type_rows", []), "비아파트 물건유형별 최근 13개월 월세 거래 추이(연립·다세대·단독·다가구)", "chart24")
    return charts


HTML_TEMPLATE = r"""
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="standard-version" content="v58">
<meta name="template-preservation-version" content="v5.8.38-template-v58-print">
<title>{{ meta.region }} 주택시장 {{ meta.report_title | default("월간동향보고서") }}</title>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" rel="stylesheet"/>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.6.1/dist/chart.umd.min.js"></script>
<script>

window.addEventListener('error', function(e) {
  try {
    const msg = String(e && e.message || '');
    if (msg.includes('semanticChartColors') || msg.includes('registerChart') || msg.includes('Unexpected')) {
      document.documentElement.setAttribute('data-chart-error', msg);
      console.error('월간보고서 차트 스크립트 오류:', msg);
    }
  } catch (_) {}
});

// v51 저메모리 미리보기: Chart.js 지연 렌더링
if (typeof Chart !== "undefined") {
  Chart.defaults.animation = false;
  Chart.defaults.responsive = true;
  Chart.defaults.maintainAspectRatio = false;
  Chart.defaults.devicePixelRatio = Math.min(window.devicePixelRatio || 1, 1.5);
}


// v58 전체 차트 품질 표준화: 의미값 기반 색상·범례·축·툴팁·반응형 기준 통합
// 핵심: 데이터 배열 순서가 아니라 label/dataset 의미로 색상을 적용한다.
const semanticChartColors = {
  direction: {
    '상승': { border: '#ef4444', background: '#ef4444cc', soft: '#ef444433' },
    '보합': { border: '#94a3b8', background: '#94a3b8cc', soft: '#94a3b833' },
    '하락': { border: '#3b82f6', background: '#3b82f6cc', soft: '#3b82f633' },
    // v5.8.54: 신규계약 0건 → 시장강도 산출 불가. 보합과 혼동되지 않게 앰버 계열.
    '산출불가': { border: '#d97706', background: '#fcd34dcc', soft: '#fcd34d55' },
    '신규 0건': { border: '#d97706', background: '#fcd34dcc', soft: '#fcd34d55' }
  },
  market: {
    '매매': { border: '#dc2626', background: '#dc2626cc', soft: '#dc262633' },
    '아파트 매매': { border: '#dc2626', background: '#dc2626cc', soft: '#dc262633' },
    '전세': { border: '#16a34a', background: '#16a34acc', soft: '#16a34a33' },
    '아파트 전세': { border: '#16a34a', background: '#16a34acc', soft: '#16a34a33' },
    '월세': { border: '#f97316', background: '#f97316cc', soft: '#f9731633' },
    '아파트 월세': { border: '#f97316', background: '#f97316cc', soft: '#f9731633' },
    '통합': { border: '#334155', background: '#334155cc', soft: '#33415533' },
    '전체': { border: '#334155', background: '#334155cc', soft: '#33415533' },
    '평균 변동률': { border: '#7c3aed', background: '#7c3aed33', soft: '#7c3aed22' },
    '매매 평균 변동률': { border: '#dc2626', background: '#dc262633', soft: '#dc262622' },
    '전세 평균 변동률': { border: '#059669', background: '#05966933', soft: '#05966922' },
    '월세 평균 변동률': { border: '#ea580c', background: '#ea580c33', soft: '#ea580c22' },
    '연립·다세대': { border: '#0ea5e9', background: '#0ea5e9cc', soft: '#0ea5e933' },
    '단독·다가구': { border: '#f97316', background: '#f97316cc', soft: '#f9731633' },
    '매매 거래량(건)': { border: '#64748b', background: '#64748bcc', soft: '#64748b33' },
    '전세 거래량(건)': { border: '#64748b', background: '#64748bcc', soft: '#64748b33' },
    '월세 거래량(건)': { border: '#64748b', background: '#64748bcc', soft: '#64748b33' },
  },
  housingType: {
    '아파트': { border: '#4f46e5', background: '#4f46e5cc', soft: '#4f46e533' },
    '연립': { border: '#0ea5e9', background: '#0ea5e9cc', soft: '#0ea5e933' },
    '다세대': { border: '#6366f1', background: '#6366f1cc', soft: '#6366f133' },
    '단독': { border: '#f97316', background: '#f97316cc', soft: '#f9731633' },
    '다가구': { border: '#22c55e', background: '#22c55ecc', soft: '#22c55e33' },
    '비아파트': { border: '#64748b', background: '#64748bcc', soft: '#64748b33' },
    '공동주택': { border: '#0ea5e9', background: '#0ea5e9cc', soft: '#0ea5e933' }
  },
  age: {
    '0~9세': { border: '#5BC0EB', background: '#5BC0EBcc', soft: '#5BC0EB33' },
    '10대': { border: '#3A86FF', background: '#3A86FFcc', soft: '#3A86FF33' },
    '20대': { border: '#8338EC', background: '#8338ECcc', soft: '#8338EC33' },
    '30대': { border: '#2A9D8F', background: '#2A9D8Fcc', soft: '#2A9D8F33' },
    '40대': { border: '#F4A261', background: '#F4A261cc', soft: '#F4A26133' },
    '50대': { border: '#E76F51', background: '#E76F51cc', soft: '#E76F5133' },
    '60대': { border: '#577590', background: '#577590cc', soft: '#57759033' },
    '70대 이상': { border: '#D62828', background: '#D62828cc', soft: '#D6282833' }
  },
  supply: {
    '인허가': { border: '#6366f1', background: '#6366f1cc', soft: '#6366f133' },
    '착공': { border: '#f97316', background: '#f97316cc', soft: '#f9731633' },
    '준공': { border: '#22c55e', background: '#22c55ecc', soft: '#22c55e33' },
    '분양': { border: '#0ea5e9', background: '#0ea5e9cc', soft: '#0ea5e933' },
    '분양실적': { border: '#0ea5e9', background: '#0ea5e9cc', soft: '#0ea5e933' }
  },
  population: {
    '인구': { border: '#334155', background: '#334155cc', soft: '#33415533' },
    '주민등록인구': { border: '#334155', background: '#334155cc', soft: '#33415533' },
    '세대수': { border: '#0f766e', background: '#0f766ecc', soft: '#0f766e33' },
    '주민등록세대수': { border: '#0f766e', background: '#0f766ecc', soft: '#0f766e33' },
    '총 전입': { border: '#ef4444', background: '#ef4444cc', soft: '#ef444433' },
    '총전입': { border: '#ef4444', background: '#ef4444cc', soft: '#ef444433' },
    '총 전출': { border: '#3b82f6', background: '#3b82f6cc', soft: '#3b82f633' },
    '총전출': { border: '#3b82f6', background: '#3b82f6cc', soft: '#3b82f633' },
    '순이동': { border: '#111827', background: '#111827cc', soft: '#11182733' },
    '최신월 인구': { border: '#334155', background: '#334155cc', soft: '#33415533' },
    '13개월 증감': { border: '#7c3aed', background: '#7c3aedcc', soft: '#7c3aed33' },
    '13개월 누적 순이동': { border: '#111827', background: '#111827cc', soft: '#11182733' }
  },
  finance: {
    '주택담보대출': { border: '#dc2626', background: '#dc2626cc', soft: '#dc262633' },
    '전세자금대출': { border: '#2563eb', background: '#2563ebcc', soft: '#2563eb33' },
    '중소기업대출': { border: '#059669', background: '#059669cc', soft: '#05966933' },
    '기업일반자금대출': { border: '#9333ea', background: '#9333eacc', soft: '#9333ea33' }
  },
  inventory: {
    '미분양': { border: '#b45309', background: '#b45309cc', soft: '#b4530933' },
    '미분양 호수': { border: '#b45309', background: '#b45309cc', soft: '#b4530933' },
    '준공후 미분양': { border: '#7f1d1d', background: '#7f1d1dcc', soft: '#7f1d1d33' },
    '준공후 미분양 호수': { border: '#7f1d1d', background: '#7f1d1dcc', soft: '#7f1d1d33' }
  },
  contract: {
    '신규': { border: '#0ea5e9', background: '#0ea5e9cc', soft: '#0ea5e933' },
    '갱신': { border: '#8b5cf6', background: '#8b5cf6cc', soft: '#8b5cf633' },
    '미상': { border: '#94a3b8', background: '#94a3b8cc', soft: '#94a3b833' },
    '총계': { border: '#334155', background: '#334155cc', soft: '#33415533' }
  }
};

const chartQualityDefaults = {
  fontFamily: "Pretendard, Noto Sans KR, Arial, sans-serif",
  gridColor: 'rgba(148, 163, 184, 0.22)',
  tickColor: '#475569',
  titleColor: '#0f172a',
  legendColor: '#334155'
};

function normalizeSemanticKey(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function getSemanticColor(label, group) {
  const key = normalizeSemanticKey(label);
  const table = semanticChartColors[group] || {};
  return table[key] || null;
}

function getSemanticColorByAnyGroup(label) {
  const key = normalizeSemanticKey(label);
  return getSemanticColor(key, 'direction') ||
    getSemanticColor(key, 'market') ||
    getSemanticColor(key, 'housingType') ||
    getSemanticColor(key, 'age') ||
    getSemanticColor(key, 'supply') ||
    getSemanticColor(key, 'population') ||
    getSemanticColor(key, 'finance') ||
    getSemanticColor(key, 'inventory') ||
    getSemanticColor(key, 'contract');
}

function isColorArray(value) {
  return Array.isArray(value) && value.length > 1;
}

function applyDatasetColor(ds, color, configType) {
  if (!ds || !color) return;
  const type = ds.type || configType;
  ds.borderColor = color.border;
  ds.backgroundColor = (type === 'line') ? color.soft : color.background;
  if (type === 'line') {
    ds.pointBackgroundColor = color.border;
    ds.pointBorderColor = color.border;
    ds.borderWidth = ds.borderWidth || 2;
    ds.tension = typeof ds.tension === 'number' ? ds.tension : 0.25;
  }
  if (type === 'bar') {
    ds.borderWidth = ds.borderWidth || 1.5;
    ds.borderRadius = ds.borderRadius || 4;
  }
}

function applySemanticColors(config) {
  if (!config || !config.data) return config;
  const configType = config.type;
  const labels = Array.isArray(config.data.labels) ? config.data.labels : [];
  const datasets = Array.isArray(config.data.datasets) ? config.data.datasets : [];

  // 도넛·파이형: labels가 상승·보합·하락, 연령대, 공급단계 등 의미값이면 label별 배열 색상 적용
  if (['doughnut', 'pie', 'polarArea'].includes(configType) && datasets[0]) {
    const mapped = labels.map(label => getSemanticColorByAnyGroup(label));
    if (mapped.some(Boolean)) {
      datasets[0].backgroundColor = mapped.map(c => (c && c.background) || '#cbd5e1cc');
      datasets[0].borderColor = mapped.map(c => (c && c.border) || '#cbd5e1');
      datasets[0].borderWidth = datasets[0].borderWidth || 1;
    }
  }

  datasets.forEach(ds => {
    const label = normalizeSemanticKey(ds.label);
    const color = getSemanticColorByAnyGroup(label);

    // 이미 연령대별 배열 색상이 들어간 전입·전출 비교 차트는 배열 색상을 유지한다.
    if (isColorArray(ds.backgroundColor) || isColorArray(ds.borderColor)) {
      return;
    }

    if (color) {
      applyDatasetColor(ds, color, configType);
      return;
    }

    // 단일 dataset 차트에서 x축 labels가 의미값이면 막대별 색상 배열 적용
    const type = ds.type || configType;
    if (datasets.length === 1 && labels.length > 0 && ['bar', 'doughnut', 'pie', 'polarArea'].includes(type)) {
      const mapped = labels.map(item => getSemanticColorByAnyGroup(item));
      if (mapped.some(Boolean)) {
        ds.backgroundColor = mapped.map(c => (c && c.background) || '#cbd5e1cc');
        ds.borderColor = mapped.map(c => (c && c.border) || '#cbd5e1');
        ds.borderWidth = ds.borderWidth || 1;
      }
    }
  });
  return config;
}

function normalizeChartQuality(config) {
  if (!config) return config;
  config.options = config.options || {};
  config.options.animation = false;
  config.options.responsive = true;
  config.options.maintainAspectRatio = false;
  config.options.layout = config.options.layout || {};
  config.options.layout.padding = Object.assign({ top: 6, right: 10, bottom: 18, left: 10 }, config.options.layout.padding || {});
  config.options.plugins = config.options.plugins || {};

  if (config.options.plugins.legend !== false) {
    config.options.plugins.legend = config.options.plugins.legend || {};
    if (config.options.plugins.legend.display !== false) {
      config.options.plugins.legend.position = config.options.plugins.legend.position || 'bottom';
      config.options.plugins.legend.labels = Object.assign({
        color: chartQualityDefaults.legendColor,
        boxWidth: 12,
        boxHeight: 12,
        padding: 16,
        font: { size: 11, family: chartQualityDefaults.fontFamily, weight: '700' }
      }, config.options.plugins.legend.labels || {});
      if (config.options.plugins.legend.position === 'bottom') {
        config.options.layout.padding.bottom = Math.max(Number(config.options.layout.padding.bottom || 0), 30);
      }
    }
  }

  config.options.plugins.title = config.options.plugins.title || {};
  if (config.options.plugins.title.display) {
    config.options.plugins.title.color = config.options.plugins.title.color || chartQualityDefaults.titleColor;
    config.options.plugins.title.font = Object.assign({
      size: 14,
      family: chartQualityDefaults.fontFamily,
      weight: '900'
    }, config.options.plugins.title.font || {});
    config.options.plugins.title.padding = config.options.plugins.title.padding || { top: 4, bottom: 12 };
  }

  config.options.plugins.tooltip = config.options.plugins.tooltip || {};
  config.options.plugins.tooltip.backgroundColor = config.options.plugins.tooltip.backgroundColor || 'rgba(15, 23, 42, 0.92)';
  config.options.plugins.tooltip.titleFont = config.options.plugins.tooltip.titleFont || { family: chartQualityDefaults.fontFamily, weight: '800' };
  config.options.plugins.tooltip.bodyFont = config.options.plugins.tooltip.bodyFont || { family: chartQualityDefaults.fontFamily };
  config.options.plugins.tooltip.padding = config.options.plugins.tooltip.padding || 10;
  if (!config.options.plugins.tooltip.callbacks) {
    config.options.plugins.tooltip.callbacks = {
      label: function(ctx) {
        const label = ctx.dataset && ctx.dataset.label ? ctx.dataset.label + ': ' : '';
        const value = ctx.parsed && typeof ctx.parsed.y !== 'undefined' ? ctx.parsed.y : ctx.parsed;
        if (typeof value === 'number') return label + value.toLocaleString('ko-KR');
        return label + value;
      }
    };
  }

  config.options.scales = config.options.scales || {};
  Object.keys(config.options.scales).forEach(axisKey => {
    const axis = config.options.scales[axisKey] || {};
    axis.grid = Object.assign({ color: chartQualityDefaults.gridColor }, axis.grid || {});
    axis.ticks = Object.assign({
      color: chartQualityDefaults.tickColor,
      font: { size: 10, family: chartQualityDefaults.fontFamily, weight: '700' },
      maxRotation: axisKey === 'x' ? 45 : 0,
      autoSkip: true
    }, axis.ticks || {});
    if (axis.title && axis.title.display) {
      axis.title.color = axis.title.color || chartQualityDefaults.tickColor;
      axis.title.font = Object.assign({ size: 11, family: chartQualityDefaults.fontFamily, weight: '800' }, axis.title.font || {});
    }
    config.options.scales[axisKey] = axis;
  });

  if (!config.options.interaction) {
    config.options.interaction = { mode: 'index', intersect: false };
  }
  return config;
}

function suppressInternalChartTitles(config) {
  if (!config || typeof config !== 'object') return config;
  config.options = config.options || {};
  config.options.plugins = config.options.plugins || {};
  if (config.options.plugins.title) {
    config.options.plugins.title.display = false;
  }
  if (config.options.plugins.subtitle) {
    config.options.plugins.subtitle.display = false;
  }
  return config;
}

function prepareChartConfig(config) {
  return suppressInternalChartTitles(normalizeChartQuality(applySemanticColors(normalizeChartDataShape(config))));
}

// v58 지역·월·데이터 변경 안정성 검수 유틸리티
const reportTemplateQualitySpec = Object.freeze({
  standardVersion: 'v58',
  expectedChapters: 6,
  expectedTables: 40,
  expectedCanvases: 55,
  expectedCanvasIds: [
    'surveySaleDonut','surveyJeonseDonut','surveyRentDonut','chart1','chart2','chartAgePopLatest','chartAgePopChange','chartAgeNetYoung','chart3','chart4','totalPermit','totalStart','totalComplete','totalSale','permitChart_아파트','permitChart_연립','permitChart_다세대','permitChart_단독','permitChart_다가구','startChart_아파트','startChart_연립','startChart_다세대','startChart_단독','startChart_다가구','completeChart_아파트','completeChart_연립','completeChart_다세대','completeChart_단독','completeChart_다가구','chart7','chart33','chart11','chart12','chart30','chart13','chart14','chart15','chart16','chart8','chart34','chart18','chart31','chart19','chart25','chart26','chart20','chart9','chart35','chart22','chart32','chart23','chart27','chart28','chart29','chart24'
  ]
});

function normalizeChartDataShape(config) {
  if (!config || !config.data) return config;
  const labels = Array.isArray(config.data.labels) ? config.data.labels : [];
  const datasets = Array.isArray(config.data.datasets) ? config.data.datasets : [];
  datasets.forEach(ds => {
    if (!Array.isArray(ds.data)) ds.data = [];
    if (labels.length && ds.data.length !== labels.length) {
      const normalized = labels.map((_, i) => (typeof ds.data[i] === 'undefined' ? null : ds.data[i]));
      ds.data = normalized;
    }
  });
  return config;
}

function chartHasRenderableData(config) {
  const datasets = config && config.data && Array.isArray(config.data.datasets) ? config.data.datasets : [];
  return datasets.some(ds => Array.isArray(ds.data) && ds.data.some(v => v !== null && typeof v !== 'undefined' && !(typeof v === 'number' && !Number.isFinite(v))));
}

function addEmptyDataNote(canvasId, message) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || !canvas.parentElement || canvas.parentElement.querySelector('.data-empty-note')) return;
  const note = document.createElement('span');
  note.className = 'data-empty-note';
  note.textContent = message || '해당 기간 자료가 없거나 비교 가능한 데이터가 부족하여 차트 영역만 유지함.';
  canvas.parentElement.appendChild(note);
}

function validateTemplateStructure() {
  const ids = reportTemplateQualitySpec.expectedCanvasIds;
  const actualIds = Array.from(document.querySelectorAll('canvas[id]')).map(el => el.id);
  const missing = ids.filter(id => !actualIds.includes(id));
  const duplicates = actualIds.filter((id, idx) => actualIds.indexOf(id) !== idx);
  const result = {
    chapters: document.querySelectorAll('.chapter').length,
    tables: document.querySelectorAll('table').length,
    canvases: document.querySelectorAll('canvas').length,
    missingCanvasIds: missing,
    duplicateCanvasIds: Array.from(new Set(duplicates))
  };
  if (missing.length || duplicates.length || result.chapters !== reportTemplateQualitySpec.expectedChapters || result.tables !== reportTemplateQualitySpec.expectedTables || result.canvases !== reportTemplateQualitySpec.expectedCanvases) {
    console.warn('[v58 report quality check] template structure mismatch', result);
  }
  return result;
}



const __lazyChartQueue = [];
const __lazyChartInstances = new Map();

function registerLazyChart(canvasId, config) {
  const prepared = prepareChartConfig(config);
  if (!chartHasRenderableData(prepared)) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => addEmptyDataNote(canvasId));
    } else {
      addEmptyDataNote(canvasId);
    }
  }
  __lazyChartQueue.push({ canvasId, config: prepared });
}

function __createLazyChart(item) {
  if (!item || __lazyChartInstances.has(item.canvasId)) return;
  const el = document.getElementById(item.canvasId);
  if (!el || typeof Chart === "undefined") return;

  const cfg = item.config || {};
  cfg.options = cfg.options || {};
  cfg.options.animation = false;
  cfg.options.responsive = true;
  cfg.options.maintainAspectRatio = false;

  try {
    const chart = new Chart(el, cfg);
    __lazyChartInstances.set(item.canvasId, chart);
  } catch (e) {
    console.error("Chart render error:", item.canvasId, e);
  }
}

function __bootLazyCharts() {
  if (!("IntersectionObserver" in window)) {
    let i = 0;
    const step = () => {
      const batch = __lazyChartQueue.slice(i, i + 3);
      batch.forEach(__createLazyChart);
      i += 3;
      if (i < __lazyChartQueue.length) setTimeout(step, 120);
    };
    step();
    return;
  }

  const observer = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (!entry.isIntersecting) return;
      const id = entry.target.getAttribute("id");
      const item = __lazyChartQueue.find(x => x.canvasId === id);
      __createLazyChart(item);
      observer.unobserve(entry.target);
    });
  }, { rootMargin: "450px 0px", threshold: 0.01 });

  __lazyChartQueue.forEach(item => {
    const el = document.getElementById(item.canvasId);
    if (el) observer.observe(el);
  });

  setTimeout(() => {
    __lazyChartQueue.slice(0, 4).forEach(__createLazyChart);
  }, 200);
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", __bootLazyCharts);
} else {
  __bootLazyCharts();
}

window.addEventListener("beforeprint", () => {
  __lazyChartQueue.forEach(__createLazyChart);
});


// v34 금리 차트 최신값 라벨 표시 플러그인
const rateLatestValueLabelPlugin = {
  id: "rateLatestValueLabelPlugin",
  afterDatasetsDraw(chart, args, pluginOptions) {
    const titleText = chart?.options?.plugins?.title?.text || "";
    if (!String(titleText).includes("대출금리")) return;
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      if (!meta || meta.hidden || !meta.data || meta.data.length === 0) return;
      let lastIndex = dataset.data.length - 1;
      while (lastIndex >= 0 && (dataset.data[lastIndex] === null || dataset.data[lastIndex] === undefined || Number.isNaN(Number(dataset.data[lastIndex])))) {
        lastIndex--;
      }
      if (lastIndex < 0) return;
      const point = meta.data[lastIndex];
      if (!point) return;
      const value = Number(dataset.data[lastIndex]);
      ctx.save();
      ctx.font = "700 11px sans-serif";
      ctx.fillStyle = dataset.borderColor || "#111827";
      ctx.textAlign = "left";
      ctx.textBaseline = "middle";
      ctx.fillText(value.toFixed(2) + "%", point.x + 6, point.y);
      ctx.restore();
    });
  }
};
if (typeof Chart !== "undefined" && !Chart.registry.plugins.get("rateLatestValueLabelPlugin")) {
  Chart.register(rateLatestValueLabelPlugin);
}

// 상승비율·하락비율·보합비율 차트 상단 '판정 스트립'(상승우세/하락우세/보합우세/혼조/산출불가)
const VERDICT_COLOR = {'상승우세':'#dc2626','하락우세':'#2563eb','보합우세':'#475569','산출불가':'#94a3b8'};
const VERDICT_SHORT = {'상승우세':'상승','하락우세':'하락','보합우세':'보합','산출불가':'-'};
const marketVerdictStripPlugin = {
  id: 'marketVerdictStrip',
  afterDatasetsDraw(chart, args, opts) {
    const verdicts = opts && opts.verdicts;
    if (!Array.isArray(verdicts) || verdicts.length === 0) return;
    const { ctx, chartArea, scales } = chart;
    if (!chartArea) return;
    const xScale = scales && scales.x;
    let meta0 = null;
    try { meta0 = chart.getDatasetMeta(0); } catch (e) { meta0 = null; }
    const y = chartArea.top + 7;
    ctx.save();
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = '700 9px sans-serif';
    verdicts.forEach((v, i) => {
      if (!v) return;
      let x = null;
      if (meta0 && meta0.data && meta0.data[i]) x = meta0.data[i].x;
      else if (xScale && typeof xScale.getPixelForValue === 'function') x = xScale.getPixelForValue(i);
      if (x === null || x === undefined || Number.isNaN(x)) return;
      const color = VERDICT_COLOR[v] || '#475569';
      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fill();
      ctx.fillText(VERDICT_SHORT[v] || v, x, y + 11);
    });
    ctx.restore();
  }
};
if (typeof Chart !== "undefined" && !Chart.registry.plugins.get("marketVerdictStrip")) {
  Chart.register(marketVerdictStripPlugin);
}

// 정부 정책 발표·시행 시점 세로 점선과 짧은 정책명. 가까운 이름은 줄을 바꿔 겹치지 않게 한다.
const policyEventLinesPlugin = {
  id: 'policyEventLines',
  afterDatasetsDraw(chart, args, opts) {
    const events = opts && opts.events;
    if (!Array.isArray(events) || events.length === 0) return;
    const { ctx, chartArea, scales } = chart;
    const xScale = scales && scales.x;
    if (!chartArea || !xScale || typeof xScale.getPixelForValue !== 'function') return;
    const n = (chart.data.labels || []).length;
    const step = n > 1 ? (xScale.getPixelForValue(1) - xScale.getPixelForValue(0)) : (chartArea.right - chartArea.left);
    const band = chartArea.top - 46;
    ctx.save();
    ctx.font = '700 10px sans-serif';
    ctx.textBaseline = 'top';
    let lastRight = [-1e9, -1e9, -1e9];
    events.forEach(ev => {
      const px = xScale.getPixelForValue(ev.index) + ((typeof ev.frac === 'number' ? ev.frac : 0.5) - 0.5) * step;
      if (!Number.isFinite(px)) return;
      ctx.strokeStyle = 'rgba(194,65,12,0.8)';
      ctx.lineWidth = 1.4;
      ctx.setLineDash([4, 3]);
      const text = String(ev.label || '');
      const w = ctx.measureText(text).width + 8;
      let tx = Math.min(Math.max(px - w / 2, chartArea.left), chartArea.right - w);
      let row = lastRight.findIndex(r => tx > r + 3);
      if (row < 0) row = lastRight.indexOf(Math.min(...lastRight));
      lastRight[row] = tx + w;
      const ty = band + row * 15;
      ctx.beginPath();
      ctx.moveTo(px, ty + 13);
      ctx.lineTo(px, chartArea.bottom);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = 'rgba(255,247,237,0.95)';
      ctx.strokeStyle = 'rgba(234,88,12,0.6)';
      ctx.lineWidth = 1;
      ctx.fillRect(tx, ty, w, 13);
      ctx.strokeRect(tx, ty, w, 13);
      ctx.fillStyle = '#9a3412';
      ctx.fillText(text, tx + 4, ty + 1.5);
    });
    ctx.restore();
  }
};
if (typeof Chart !== "undefined" && !Chart.registry.plugins.get("policyEventLines")) {
  Chart.register(policyEventLinesPlugin);
}

const centerTextPlugin = {
  id: 'centerText',
  afterDraw(chart, args, opts) {
    if (chart.config.type !== 'doughnut') return;
    const { ctx, chartArea } = chart;
    const centerX = (chartArea.left + chartArea.right) / 2;
    const centerY = (chartArea.top + chartArea.bottom) / 2;
    const dataset = chart.data.datasets[0] || { data: [] };
    const total = dataset.data.reduce((a, b) => a + (Number(b) || 0), 0);
    const max = Math.max(...dataset.data.map(v => Number(v) || 0));
    const idx = dataset.data.map(v => Number(v) || 0).indexOf(max);
    const defaultText = chart.data.labels[idx] || '';
    const centerText = opts?.text || defaultText;
    const subText = opts?.subtext || (total ? total.toFixed(0) + '%' : '');
    ctx.save();
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillStyle = '#0f172a';
    ctx.font = '800 15px sans-serif';
    ctx.fillText(centerText, centerX, centerY - 6);
    ctx.fillStyle = '#475569';
    ctx.font = '700 11px sans-serif';
    ctx.fillText(subText, centerX, centerY + 12);
    ctx.restore();
  }
};
Chart.register(centerTextPlugin);

// 도넛 조각 안에 비율(%)과 건수를 적는다. 너무 작은 조각(6% 미만)은 비워 둔다.
const doughnutSliceLabelsPlugin = {
  id: 'doughnutSliceLabels',
  afterDatasetsDraw(chart, args, opts) {
    if (chart.config.type !== 'doughnut' || !opts || !Array.isArray(opts.counts)) return;
    const meta = chart.getDatasetMeta(0);
    const values = (chart.data.datasets[0] || { data: [] }).data.map(v => Number(v) || 0);
    if (!meta || !meta.data) return;
    const { ctx } = chart;
    ctx.save();
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    meta.data.forEach((arc, i) => {
      const v = values[i];
      if (!(v >= 6) || !chart.getDataVisibility(i)) return;
      const p = arc.getProps(['startAngle', 'endAngle', 'innerRadius', 'outerRadius', 'x', 'y'], true);
      const mid = (p.startAngle + p.endAngle) / 2;
      const r = (p.innerRadius + p.outerRadius) / 2;
      const x = p.x + Math.cos(mid) * r;
      const y = p.y + Math.sin(mid) * r;
      const cnt = opts.counts[i];
      ctx.fillStyle = '#ffffff';
      ctx.shadowColor = 'rgba(15,23,42,0.45)';
      ctx.shadowBlur = 2;
      ctx.font = '800 11px sans-serif';
      ctx.fillText(v.toFixed(1) + '%', x, cnt === null || cnt === undefined ? y : y - 6);
      if (cnt !== null && cnt !== undefined) {
        ctx.font = '700 9.5px sans-serif';
        ctx.fillText(cnt + '건', x, y + 6);
      }
    });
    ctx.restore();
  }
};
if (typeof Chart !== "undefined" && !Chart.registry.plugins.get("doughnutSliceLabels")) {
  Chart.register(doughnutSliceLabelsPlugin);
}
// 모든 차트 공통: 범례 항목을 누르면 해당 계열을 켜고 끈다(Chart.js 기본 동작). 누를 수 있다는 것을 커서로 보여 준다.
if (typeof Chart !== "undefined") {
  Chart.defaults.plugins.legend.onHover = (e) => { if (e && e.native && e.native.target) e.native.target.style.cursor = 'pointer'; };
  Chart.defaults.plugins.legend.onLeave = (e) => { if (e && e.native && e.native.target) e.native.target.style.cursor = 'default'; };
}


// v53 조사 담당자 관점 매매·전세·월세 상승·보합·하락 도넛 차트
const marketDirectionLabels = ['상승', '보합', '하락'];
const marketDirectionColors = marketDirectionLabels.map(label => semanticChartColors.direction[label].background);
const marketDirectionBorders = marketDirectionLabels.map(label => semanticChartColors.direction[label].border);
function makeSurveyDonutConfig(title, values, centerText, subText) {
  return {
    type: 'doughnut',
    data: {
      labels: marketDirectionLabels,
      datasets: [{
        data: values,
        backgroundColor: marketDirectionColors,
        borderColor: marketDirectionBorders,
        borderWidth: 1,
        hoverOffset: 3
      }]
    },
    options: {
      cutout: '66%',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } },
        title: { display: false },
        tooltip: { callbacks: { label: (ctx) => `${ctx.label}: ${ctx.raw}%` } },
        centerText: { text: centerText, subtext: subText }
      }
    }
  };
}
function registerChart(canvasId, config) { registerLazyChart(canvasId, config); }

</script>
<style>

@page { size:A4; margin:15mm; }
.cover-break { break-after: page; page-break-after: always; height:0; }
.print-guide { max-width:1280px; margin:0 auto 16px; background:#eef2ff; border:1px solid #c7d2fe; color:#312e81; padding:10px 14px; border-radius:14px; font-weight:700; font-size:13px; }
.keep-name { word-break: keep-all; overflow-wrap: normal; }
* { box-sizing:border-box; } body { font-family:'Pretendard','Noto Sans KR',Arial,sans-serif; color:#0f172a; background:#f8fafc; line-height:1.55; }
.hero { display:flex; justify-content:space-between; gap:24px; padding:32px; background:linear-gradient(135deg,#0f172a,#334155); color:#fff; border-radius:24px; margin:20px auto; max-width:1280px; page-break-inside:avoid; }
.eyebrow { color:#fcd34d; font-weight:700; letter-spacing:.04em; } h1 { font-size:34px; font-weight:900; margin:4px 0 8px; } .subtitle { color:#cbd5e1; } .meta { min-width:280px; background:rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.2); border-radius:18px; padding:16px; display:grid; grid-template-columns:90px 1fr; gap:8px; font-size:13px; }
.notice {max-width:1280px; margin:0 auto 16px; background:#fefce8; border:1px solid #fde68a; color:#713f12; padding:12px 16px; border-radius:14px; font-weight:700;} .notice.data-quality-notice { background:#fff7ed; border-color:#fdba74; color:#9a3412; font-weight:600; line-height:1.7; } .notice.data-quality-notice b { font-weight:800; } .kpis { max-width:1280px; margin:0 auto 16px; display:grid; grid-template-columns:repeat(4,1fr); gap:14px; } .kpis div { background:white; border-radius:18px; padding:18px; box-shadow:0 8px 20px rgba(15,23,42,.07); } .kpis span { display:block; color:#64748b; font-size:13px; } .kpis b { display:block; font-size:25px; margin:5px 0; } .kpis small { color:#64748b; }
.chapter { max-width:1280px; margin:18px auto; padding:22px; border-radius:24px; page-break-inside:avoid; } .chapter>h2 { font-size:24px; font-weight:900; margin-bottom:16px; }
.card { background:rgba(255,255,255,.82); border:1px solid #e2e8f0; box-shadow:0 8px 24px rgba(15,23,42,.05); border-radius:20px; padding:20px; margin:16px 0; page-break-inside:avoid; } .card h3 { font-size:19px; font-weight:800; margin-bottom:10px; } .card h4 { font-size:16px; font-weight:800; margin:16px 0 8px; }
.meaning { background:#f1f5f9; border-left:4px solid #64748b; padding:10px 12px; border-radius:10px; color:#334155; font-size:14px; } .note { background:#fff7ed; border:1px solid #fed7aa; border-left:5px solid #f97316; border-radius:14px; padding:13px 15px; margin-top:12px; font-weight:600; color:#431407; }
.table-wrap { overflow-x:auto; margin:12px 0; border-radius:14px; border:1px solid #e2e8f0; } table { width:100%; border-collapse:collapse; font-size:12px; background:white; } th { background:#1e293b; color:white; padding:9px 8px; text-align:center; white-space:nowrap; } td { border-top:1px solid #e2e8f0; padding:8px; text-align:center; white-space:nowrap; } tr:nth-child(even) td { background:#f8fafc; } .latest-row td { background:#fef9c3 !important; font-weight:800; color:#713f12; }
.chart-box { position:relative; background:white; border:1px solid #e2e8f0; border-radius:16px; padding:10px; margin:12px 0; page-break-inside:avoid; min-height:320px; } .chart-box canvas { width:100% !important; height:100% !important; display:block; }
.market-block { background:white; border:1px solid #c7d2fe; border-radius:22px; padding:22px; margin:18px 0; page-break-inside:avoid; box-shadow:0 8px 24px rgba(79,70,229,.07); } .market-block>h3 { font-size:21px; font-weight:900; margin-bottom:12px; display:flex; align-items:center; gap:8px; } .market-item { background:#f8fafc; border-left:6px solid #6366f1; border-radius:16px; padding:15px 17px; margin:12px 0; page-break-inside:avoid; } .market-item h4 { font-size:16px; font-weight:900; margin-bottom:7px; color:#1e293b; } .market-item p { color:#1e293b; }
.narr { background:white; border-left:6px solid #6366f1; padding:16px 18px; margin:12px 0; border-radius:16px; page-break-inside:avoid; } .narr h3 { font-weight:900; font-size:17px; margin-bottom:8px; } .narr p { color:#1e293b; }
.bg-slate { background:#f1f5f9; } .bg-amber { background:#fffbeb; } .bg-rose { background:#fff1f2; } .bg-emerald { background:#ecfdf5; } .bg-blue { background:#eff6ff; } .bg-indigo { background:#eef2ff; }
.external-buyer-card table { font-size:10.5px; table-layout:auto; }
.external-buyer-card th,.external-buyer-card td { padding:6px 5px; }
@media print { body { background:white; } .hero,.chapter,.card,.chart-box,.narr { page-break-inside:avoid; box-shadow:none; } .kpis { grid-template-columns:repeat(2,1fr); } .external-buyer-card table { font-size:8.5pt; } .external-buyer-card th,.external-buyer-card td { padding:4px 3px; } }
@media (max-width:900px) { .hero { flex-direction:column; } .kpis { grid-template-columns:1fr 1fr; } }

.pos-value{color:#dc2626!important;font-weight:800;} .neg-value{color:#2563eb!important;font-weight:800;} .zero-value{color:#64748b!important;font-weight:700;} .chart-box{min-height:340px;} .chart-box canvas{width:100%!important;height:100%!important;display:block;} .merged-badge{display:inline-block;background:#e0f2fe;color:#075985;font-weight:800;border-radius:999px;padding:3px 9px;margin-left:8px;font-size:12px;}


/* === v9 PDF/가독성 보완 === */
.chapter { page-break-inside:auto !important; break-inside:auto !important; }
.card { break-inside:auto; }
.card.compact-card { break-inside:auto; page-break-inside:auto; }
.table-wrap { overflow-x:visible !important; }
table { table-layout:auto; }
th, td { word-break:keep-all; }
td.num { text-align:right !important; font-variant-numeric:tabular-nums; }
td.rate, td.center { text-align:center !important; }
.no-break { break-inside:avoid; page-break-inside:avoid; }
.summary-box { background:#f8fafc; border:1px solid #cbd5e1; border-left:6px solid #4f46e5; border-radius:16px; padding:14px 16px; margin:12px 0 16px; }
.summary-box h4 { margin:0 0 8px !important; font-size:16px !important; font-weight:900 !important; color:#1e293b; }
.summary-box ul { margin:0; padding-left:18px; }
.summary-box li { margin:3px 0; color:#334155; font-weight:700; }
.subtable-title { font-size:15px; font-weight:900; color:#334155; margin:14px 0 6px; padding-left:4px; }
.small-caption { font-size:12px; color:#64748b; margin-top:-6px; margin-bottom:8px; }
@media print {
  body { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  .hero { padding:22px !important; }
  .chapter { padding:16px !important; margin:10px auto !important; page-break-inside:auto !important; break-inside:auto !important; }
  .card { padding:14px !important; margin:10px 0 !important; page-break-inside:auto !important; break-inside:auto !important; }
  .card h3 { font-size:15px !important; line-height:1.25 !important; }
  .meaning, .note { font-size:10.5px !important; line-height:1.45 !important; padding:8px 10px !important; }
  .table-wrap { overflow:visible !important; }
  table { width:100% !important; font-size:9.2px !important; table-layout:fixed !important; }
  th, td { padding:4px 3px !important; white-space:normal !important; word-break:keep-all !important; line-height:1.3 !important; }
  th { font-size:9.2px !important; }
  .chart-box { height:260px !important; min-height:260px !important; }
  .market-item p { font-size:10.6px !important; line-height:1.5 !important; }
  .summary-box { padding:9px 12px !important; margin:8px 0 10px !important; }
  .summary-box h4 { font-size:12px !important; }
  .summary-box li { font-size:10.5px !important; }
}


/* v11 핵심 수치 강조 규칙 */
.key-max {
  font-weight: 900 !important;
  color: #111827 !important;
  background: rgba(254, 240, 138, 0.55);
  border-radius: 6px;
  padding: 1px 4px;
}
.key-max-red {
  font-weight: 900 !important;
  color: #dc2626 !important;
}
.key-max-blue {
  font-weight: 900 !important;
  color: #2563eb !important;
}
.key-max-count {
  font-weight: 900 !important;
  color: #111827 !important;
}
@media print {
  .key-max {
    font-weight: 900 !important;
    background: rgba(254, 240, 138, 0.45) !important;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
}


/* v14 통합표 컬럼 정렬 및 가독성 보완 */
.market-strength-table {
  table-layout: fixed !important;
  width: 100% !important;
}
.market-strength-table th,
.market-strength-table td {
  text-align: center !important;
  vertical-align: middle !important;
}
.market-strength-table th.compare-col,
.market-strength-table td.compare-col {
  min-width: 70px !important;
  max-width: 82px !important;
  width: 76px !important;
  white-space: normal !important;
  line-height: 1.25 !important;
  text-align: center !important;
}
.market-strength-table th.month-col,
.market-strength-table td.month-col {
  width: 66px !important;
}
.market-strength-table th.count-col,
.market-strength-table td.count-col {
  width: 62px !important;
}
.market-strength-table th.rate-col,
.market-strength-table td.rate-col {
  width: 68px !important;
}
.market-strength-table .short-head {
  display: inline-block;
  line-height: 1.25;
}
@media print {
  .market-strength-table th,
  .market-strength-table td {
    font-size: 9.5px !important;
    padding: 5px 4px !important;
  }
}


/* v15 전세·월세 신규/갱신 건수 분리 컬럼 */
.market-strength-table th.contract-col,
.market-strength-table td.contract-col {
  width: 58px !important;
  min-width: 54px !important;
  text-align: center !important;
}
.market-strength-table th.total-col,
.market-strength-table td.total-col {
  width: 58px !important;
  min-width: 54px !important;
  text-align: center !important;
}
@media print {
  .market-strength-table th.contract-col,
  .market-strength-table td.contract-col,
  .market-strength-table th.total-col,
  .market-strength-table td.total-col {
    font-size: 9.2px !important;
    padding: 5px 3px !important;
  }
}


/* v16 전월세 계약구분 재분석 반영 */
.market-strength-table th.unknown-col,
.market-strength-table td.unknown-col {
  width: 52px !important;
  min-width: 48px !important;
  text-align: center !important;
}
.market-strength-table .contract-note {
  color: #475569;
  font-weight: 700;
}
@media print {
  .market-strength-table th.unknown-col,
  .market-strength-table td.unknown-col {
    font-size: 9px !important;
    padding: 5px 3px !important;
  }
}


/* v17 실거래 기반 전월세전환율 설명 카드 */
.conversion-card {
  background: #f8fafc;
  border: 1px solid #cbd5e1;
  border-left: 6px solid #0f766e;
  border-radius: 16px;
  padding: 14px 16px;
  margin: 12px 0;
  page-break-inside: avoid;
}
.conversion-card h4 { font-size: 15px; font-weight: 900; color: #0f172a; margin: 0 0 8px; }
.conversion-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px; margin: 8px 0; }
.conversion-grid div { background: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 8px; text-align: center; }
.conversion-grid span { display: block; color: #64748b; font-size: 11px; }
.conversion-grid b { display: block; font-size: 16px; color: #0f766e; }
@media print { .conversion-card { break-inside: avoid; } .conversion-grid { grid-template-columns: repeat(4, 1fr); } }


/* v18 종합시황 실거래 사례 보강 */
.narr h4 {
  margin-top: 14px;
  margin-bottom: 6px;
  font-weight: 900;
  color: #1e293b;
  font-size: 15px;
}
.narr .note {
  line-height: 1.65;
}


/* v19 모든 분석표 건수·거래량 계열 가운데 정렬 */
.count-center,
.count-center * {
  text-align: center !important;
  vertical-align: middle !important;
}
th.count-center,
td.count-center {
  text-align: center !important;
  vertical-align: middle !important;
  white-space: nowrap;
}
@media print {
  th.count-center,
  td.count-center {
    text-align: center !important;
    vertical-align: middle !important;
  }
}


/* v20 종합시황 자연서술형 보강 */
.narr h4 {
  margin-top: 14px;
  margin-bottom: 6px;
  font-weight: 900;
  color: #1e293b;
  font-size: 15px;
}
.narr p {
  text-align: justify;
}


/* v21 종합시황 3개 시장 블록 배경색 구분 */
.narr.market-sale {
  background: #fff1f2 !important;
  border-left-color: #ef4444 !important;
  border: 1px solid #fecdd3;
  border-left-width: 6px;
}
.narr.market-jeonse {
  background: #ecfdf5 !important;
  border-left-color: #10b981 !important;
  border: 1px solid #bbf7d0;
  border-left-width: 6px;
}
.narr.market-rent {
  background: #eff6ff !important;
  border-left-color: #3b82f6 !important;
  border: 1px solid #bfdbfe;
  border-left-width: 6px;
}
.narr.market-sale .note {
  background: #fff7f7 !important;
  border-color: #fecdd3 !important;
  border-left-color: #ef4444 !important;
}
.narr.market-jeonse .note {
  background: #f0fdf4 !important;
  border-color: #bbf7d0 !important;
  border-left-color: #10b981 !important;
}
.narr.market-rent .note {
  background: #eff6ff !important;
  border-color: #bfdbfe !important;
  border-left-color: #3b82f6 !important;
}
@media print {
  .narr.market-sale,
  .narr.market-jeonse,
  .narr.market-rent,
  .narr.market-sale .note,
  .narr.market-jeonse .note,
  .narr.market-rent .note {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
}


/* v22 상승·보합·하락 시장강도 차트 색상 통일 */
:root {
  --market-rise: #ef6b5b;
  --market-flat: #aebbc0;
  --market-fall: #5dade2;
}
.market-color-note {
  color: #475569;
  font-size: 12px;
  font-weight: 700;
}


/* v58 전체 차트 품질 표준화 변수: 지역·월 변경 시에도 의미별 색상 유지 */
:root {
  --sem-rise: #ef4444;
  --sem-rise-fill: #ef4444cc;
  --sem-flat: #94a3b8;
  --sem-flat-fill: #94a3b8cc;
  --sem-fall: #3b82f6;
  --sem-fall-fill: #3b82f6cc;
  --sem-sale: #dc2626;
  --sem-jeonse: #16a34a;
  --sem-rent: #f97316;
  --sem-total: #334155;
}


/* v23 통합표 추가 혼합차트(거래량+상승/하락 비율) */
.chart-subtitle{
  font-size:15px;
  font-weight:900;
  color:#334155;
  margin:4px 0 10px 0;
  text-align:center;
}
.chart-box.mix-chart{
  margin-top:14px;
}
.chart-box.mix-chart canvas{
  min-height:300px;
}


/* v24 혼합차트 행 겹침·밀림 방지 보완 */
.chart-box.mix-chart {
  position: relative !important;
  min-height: 340px !important;
  height: 340px !important;
  max-height: 360px !important;
  padding: 14px 12px 18px 12px !important;
  margin: 16px 0 18px 0 !important;
  overflow: hidden !important;
  clear: both !important;
  page-break-inside: avoid !important;
  break-inside: avoid !important;
}
.chart-box.mix-chart canvas {
  width: 100% !important;
  height: 280px !important;
  max-height: 280px !important;
  display: block !important;
}
.chart-subtitle {
  display: block !important;
  line-height: 1.35 !important;
  margin: 0 0 8px 0 !important;
  padding: 0 !important;
  white-space: normal !important;
  word-break: keep-all !important;
}
.market-strength-table {
  margin-bottom: 12px !important;
}
.card > .chart-box + .chart-box {
  margin-top: 18px !important;
}
@media print {
  .chart-box.mix-chart {
    height: 320px !important;
    min-height: 320px !important;
    max-height: 330px !important;
    page-break-inside: avoid !important;
    break-inside: avoid !important;
  }
  .chart-box.mix-chart canvas {
    height: 260px !important;
    max-height: 260px !important;
  }
}


/* v24 넓은 통합표 행 높이·컬럼 밀림 방지 */
.compact-wide-table th,
.compact-wide-table td {
  font-size: 10px !important;
  line-height: 1.25 !important;
  padding: 6px 4px !important;
  vertical-align: middle !important;
}
.compact-wide-table th {
  word-break: keep-all !important;
}
@media print {
  .compact-wide-table th,
  .compact-wide-table td {
    font-size: 8.8px !important;
    padding: 4px 3px !important;
  }
}


/* v25 PDF 제출용 마감 보정 */
.chart-subtitle {
  margin-top: 10px !important;
  margin-bottom: 12px !important;
  line-height: 1.45 !important;
}
.chart-box {
  margin-top: 14px !important;
  margin-bottom: 18px !important;
}
.chart-box.mix-chart {
  margin-top: 20px !important;
  margin-bottom: 22px !important;
}
.chart-box + .note,
.mix-chart + .note {
  margin-top: 18px !important;
}
.card .note {
  margin-top: 14px !important;
  margin-bottom: 12px !important;
}
.table-wrap {
  margin-bottom: 12px !important;
}
.compact-wide-table th,
.compact-wide-table td {
  line-height: 1.24 !important;
}
.market-strength-table th.contract-col,
.market-strength-table td.contract-col {
  min-width: 52px !important;
  width: 56px !important;
}
.market-strength-table th.unknown-col,
.market-strength-table td.unknown-col {
  min-width: 50px !important;
  width: 54px !important;
}
.market-strength-table th.compare-col,
.market-strength-table td.compare-col {
  min-width: 60px !important;
  width: 66px !important;
}
.market-strength-table th.avg-col,
.market-strength-table td.avg-col {
  min-width: 54px !important;
  width: 58px !important;
}
td.apt-name,
th.apt-name {
  word-break: keep-all !important;
  overflow-wrap: anywhere !important;
  min-width: 120px !important;
  text-align: left !important;
}
td.dong-name,
th.dong-name {
  min-width: 54px !important;
  text-align: center !important;
  white-space: nowrap !important;
}
.conversion-card p {
  line-height: 1.55 !important;
  margin-bottom: 0 !important;
}
@media print {
  .chart-box {
    margin-top: 12px !important;
    margin-bottom: 16px !important;
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }
  .chart-box.mix-chart {
    margin-top: 18px !important;
    margin-bottom: 20px !important;
  }
  .chart-box + .note,
  .mix-chart + .note {
    margin-top: 16px !important;
  }
  .market-strength-table th.contract-col,
  .market-strength-table td.contract-col,
  .market-strength-table th.unknown-col,
  .market-strength-table td.unknown-col,
  .market-strength-table th.compare-col,
  .market-strength-table td.compare-col,
  .market-strength-table th.avg-col,
  .market-strength-table td.avg-col {
    font-size: 8.6px !important;
    padding: 4px 2px !important;
  }
}


/* v26 주택공급 동향 및 미래 공급 압력 판단 */
.supply-card .subhead {
  margin: 18px 0 8px 0;
  font-size: 15px;
  font-weight: 900;
  color: #1e293b;
}
.supply-card .chart-box {
  min-height: 340px;
  height: 340px;
  margin-top: 14px;
  margin-bottom: 18px;
}
.supply-card .chart-box canvas {
  height: 300px !important;
}
.supply-card table th,
.supply-card table td {
  text-align: center;
  vertical-align: middle;
}
.supply-card table td:last-child,
.supply-card table th:last-child {
  text-align: left;
}
@media print {
  .supply-card .chart-box {
    height: 320px !important;
    min-height: 320px !important;
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .supply-card .chart-box canvas {
    height: 280px !important;
  }
}


/* v27 물건유형별 공급 파이프라인 시계열 */
.supply-card .subhead {
  margin: 18px 0 8px 0;
  font-size: 15px;
  font-weight: 900;
  color: #1e293b;
}
.supply-card .chart-box {
  min-height: 340px;
  height: 340px;
  margin-top: 14px;
  margin-bottom: 18px;
}
.supply-card .chart-box canvas {
  height: 300px !important;
}
.supply-card table th,
.supply-card table td {
  text-align: center;
  vertical-align: middle;
}
.supply-card table td:last-child,
.supply-card table th:last-child {
  text-align: left;
}
.supply-type-series th,
.supply-type-series td {
  font-size: 9.8px !important;
  padding: 5px 4px !important;
}
@media print {
  .supply-card .chart-box {
    height: 320px !important;
    min-height: 320px !important;
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .supply-card .chart-box canvas {
    height: 280px !important;
  }
  .supply-type-series th,
  .supply-type-series td {
    font-size: 8.5px !important;
    padding: 4px 2px !important;
  }
}


/* v28 공급단계별 개별 시계열 차트 */
.supply-card .subhead {
  margin: 18px 0 8px 0;
  font-size: 15px;
  font-weight: 900;
  color: #1e293b;
}
.supply-card .chart-box {
  min-height: 330px;
  height: 330px;
  margin-top: 14px;
  margin-bottom: 18px;
}
.supply-card .chart-box canvas {
  height: 290px !important;
}
.supply-card table th,
.supply-card table td {
  text-align: center;
  vertical-align: middle;
}
.supply-card table td:last-child,
.supply-card table th:last-child {
  text-align: left;
}
.supply-type-series th,
.supply-type-series td {
  font-size: 9.8px !important;
  padding: 5px 4px !important;
}
@media print {
  .supply-card .chart-box {
    height: 315px !important;
    min-height: 315px !important;
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .supply-card .chart-box canvas {
    height: 275px !important;
  }
  .supply-type-series th,
  .supply-type-series td {
    font-size: 8.5px !important;
    padding: 4px 2px !important;
  }
}


/* v29 공급단계별 개별 막대차트 */
.supply-card .chart-box {
  min-height: 335px;
  height: 335px;
}
.supply-card .chart-box canvas {
  height: 295px !important;
}
@media print {
  .supply-card .chart-box {
    min-height: 320px !important;
    height: 320px !important;
  }
  .supply-card .chart-box canvas {
    height: 280px !important;
  }
}


/* v31 공급 물건유형별 분리 시계열 차트 */
.supply-card .subhead {
  margin: 18px 0 8px 0;
  font-size: 15px;
  font-weight: 900;
  color: #1e293b;
}
.supply-card table th,
.supply-card table td {
  text-align: center;
  vertical-align: middle;
}
.supply-card table td:last-child,
.supply-card table th:last-child {
  text-align: left;
}
.supply-type-series th,
.supply-type-series td {
  font-size: 9.8px !important;
  padding: 5px 4px !important;
}
.supply-card .chart-box {
  min-height: 335px;
  height: 335px;
  margin-top: 14px;
  margin-bottom: 18px;
}
.supply-card .chart-box canvas {
  height: 295px !important;
}
.mini-chart-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 14px;
  margin: 10px 0 20px 0;
}
.mini-chart-card {
  border: 1px solid #e2e8f0;
  border-radius: 14px;
  padding: 10px 12px 8px 12px;
  background: #fbfdff;
}
.mini-chart-title {
  font-size: 13px;
  font-weight: 800;
  color: #334155;
  margin-bottom: 6px;
}
.mini-chart-box {
  height: 240px;
}
.mini-chart-box canvas {
  height: 220px !important;
}
@media (max-width: 900px) {
  .mini-chart-grid {
    grid-template-columns: 1fr;
  }
}
@media print {
  .supply-card .chart-box {
    min-height: 310px !important;
    height: 310px !important;
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .mini-chart-grid {
    grid-template-columns: 1fr 1fr;
    gap: 10px;
  }
  .mini-chart-card {
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .mini-chart-box {
    height: 210px !important;
  }
  .mini-chart-box canvas {
    height: 190px !important;
  }
  .supply-type-series th,
  .supply-type-series td {
    font-size: 8.5px !important;
    padding: 4px 2px !important;
  }
}


/* v32 연령대별 인구이동 분석 */
.age-move-card .subhead {
  margin: 18px 0 8px 0;
  font-size: 15px;
  font-weight: 900;
  color: #1e293b;
}
.age-net-table th,
.age-net-table td,
.age-latest-table th,
.age-latest-table td {
  font-size: 10px !important;
  padding: 6px 4px !important;
  vertical-align: middle !important;
}
.age-net-table td:last-child,
.age-net-table th:last-child,
.age-latest-table td:last-child,
.age-latest-table th:last-child {
  text-align: left !important;
}
.age-move-card .chart-box {
  min-height: 330px;
  height: 330px;
  margin-top: 14px;
  margin-bottom: 18px;
}
.age-move-card .chart-box canvas {
  height: 290px !important;
}
@media print {
  .age-net-table th,
  .age-net-table td,
  .age-latest-table th,
  .age-latest-table td {
    font-size: 8.5px !important;
    padding: 4px 2px !important;
  }
  .age-move-card .chart-box {
    min-height: 310px !important;
    height: 310px !important;
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .age-move-card .chart-box canvas {
    height: 270px !important;
  }
}


/* v33 공급 총량도 단계별 분리 차트 */
.total-stage-grid {
  grid-template-columns: repeat(2, minmax(0, 1fr));
}
.supply-card .subhead {
  margin: 18px 0 8px 0;
  font-size: 15px;
  font-weight: 900;
  color: #1e293b;
}
.supply-card table th,
.supply-card table td {
  text-align: center;
  vertical-align: middle;
}
.supply-card table td:last-child,
.supply-card table th:last-child {
  text-align: left;
}
.supply-type-series th,
.supply-type-series td {
  font-size: 9.8px !important;
  padding: 5px 4px !important;
}
.mini-chart-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 14px;
  margin: 10px 0 20px 0;
}
.mini-chart-card {
  border: 1px solid #e2e8f0;
  border-radius: 14px;
  padding: 10px 12px 8px 12px;
  background: #fbfdff;
}
.mini-chart-title {
  font-size: 13px;
  font-weight: 800;
  color: #334155;
  margin-bottom: 6px;
}
.mini-chart-box {
  height: 240px;
}
.mini-chart-box canvas {
  height: 220px !important;
}
@media (max-width: 900px) {
  .mini-chart-grid, .total-stage-grid {
    grid-template-columns: 1fr;
  }
}
@media print {
  .mini-chart-grid, .total-stage-grid {
    grid-template-columns: 1fr 1fr;
    gap: 10px;
  }
  .mini-chart-card {
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .mini-chart-box {
    height: 210px !important;
  }
  .mini-chart-box canvas {
    height: 190px !important;
  }
  .supply-type-series th,
  .supply-type-series td {
    font-size: 8.5px !important;
    padding: 4px 2px !important;
  }
}


/* v34 금리 선형차트 가독성 보강 */
.rate-chart-guide {
  margin: 12px 0 14px 0;
  padding: 12px 14px;
  border-left: 5px solid #3b82f6;
  background: #eff6ff;
  border-radius: 12px;
  color: #1e3a8a;
  font-size: 13px;
  line-height: 1.65;
}
.rate-latest-label {
  font-weight: 900;
  color: #111827;
}


/* v35 연령대별 주민등록인구 구조 분석 */
.age-pop-card .subhead {
  margin: 18px 0 8px 0;
  font-size: 15px;
  font-weight: 900;
  color: #1e293b;
}
.age-pop-table th,
.age-pop-table td,
.age-pop-latest-table th,
.age-pop-latest-table td {
  font-size: 10px !important;
  padding: 6px 4px !important;
  vertical-align: middle !important;
}
.age-pop-table td:last-child,
.age-pop-table th:last-child,
.age-pop-latest-table td:last-child,
.age-pop-latest-table th:last-child {
  text-align: left !important;
}
.age-pop-card .chart-box {
  min-height: 330px;
  height: 330px;
  margin-top: 14px;
  margin-bottom: 18px;
}
.age-pop-card .chart-box canvas {
  height: 290px !important;
}
@media print {
  .age-pop-table th,
  .age-pop-table td,
  .age-pop-latest-table th,
  .age-pop-latest-table td {
    font-size: 8.5px !important;
    padding: 4px 2px !important;
  }
  .age-pop-card .chart-box {
    min-height: 310px !important;
    height: 310px !important;
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .age-pop-card .chart-box canvas {
    height: 270px !important;
  }
}


/* v39 인구·세대수 혼합형 이중축 차트 */
.population-chart-guide {
  margin: 12px 0 14px 0;
  padding: 12px 14px;
  border-left: 5px solid #64748b;
  background: #f8fafc;
  border-radius: 12px;
  color: #334155;
  font-size: 13px;
  line-height: 1.65;
}


/* v40 보고서 품질 개선: 핵심 진단·기준월 관리표·조사상 유의사항 */
.semiconductor-cluster-badge{display:inline-block;margin-top:3px;padding:2px 6px;border-radius:999px;font-size:9px;font-weight:800;line-height:1.25;white-space:normal}
.semiconductor-cluster-badge.core{color:#9f1239;background:#ffe4e6;border:1px solid #fda4af}
.semiconductor-cluster-badge.adjacent{color:#9a3412;background:#ffedd5;border:1px solid #fdba74}
.semiconductor-cluster-badge.linked{color:#1e40af;background:#dbeafe;border:1px solid #93c5fd}
.semiconductor-cluster-legend{display:grid;gap:6px;margin:8px 0 10px;padding:10px 12px;border:1px solid #f3d29a;border-radius:10px;background:#fffaf0;color:#475569;font-size:11px;line-height:1.55}
.semiconductor-cluster-legend div{display:flex;align-items:flex-start;gap:5px;flex-wrap:wrap}
.semiconductor-cluster-legend p{margin:2px 0 0;color:#7c2d12;font-size:10px}
.move-in-bar-chart{display:grid;gap:5px;margin:8px 0 14px;padding:12px;border:1px solid #dbe4ef;border-radius:12px;background:#f8fafc}
.move-in-bar-row{display:grid;grid-template-columns:66px minmax(120px,1fr) 72px 62px;gap:8px;align-items:center;font-size:11px;color:#475569}
.move-in-bar-track{height:14px;background:#e2e8f0;border-radius:999px;overflow:hidden}
.move-in-bar-track i{display:block;height:100%;min-width:0;background:linear-gradient(90deg,#38bdf8,#2563eb);border-radius:999px}
.move-in-bar-row b{text-align:right;color:#0f172a}.move-in-bar-row small{color:#64748b}
.move-in-source-note{margin:8px 0 2px;padding:9px 11px;border-left:4px solid #3b82f6;background:#eff6ff;color:#475569;font-size:10px;line-height:1.55}
.move-in-empty{margin:8px 0;padding:10px 12px;border-radius:8px;background:#f8fafc;color:#64748b}
.move-in-key-fields{margin:6px 0 8px;padding:8px 10px;border-radius:8px;background:#eef6ff;color:#1e3a5f;font-size:11px}
.cluster-policy-card{box-sizing:border-box;width:calc(100% - 32px);max-width:1280px;border:1px solid #f3b34b;border-left:4px solid #d97706;background:#fffbeb;padding:10px 13px;margin:10px auto;border-radius:10px;color:#422006;font-size:11px;line-height:1.45}
.cluster-policy-card h2{margin:0 0 5px;color:#7c2d12;font-size:15px;line-height:1.3}
.cluster-policy-card p{margin:3px 0;line-height:1.45}
.cluster-policy-card .cluster-levels{display:flex;gap:5px;flex-wrap:wrap;margin:5px 0}
.cluster-policy-card .semiconductor-cluster-badge{margin-top:0;padding:1px 5px;font-size:8px}
.cluster-policy-card .policy-caution{background:#fff7ed;border-radius:6px;padding:5px 7px;color:#7c2d12}
.cluster-policy-card a{color:#1d4ed8;text-decoration:underline;font-size:10px}
@media print{
  .cluster-policy-card{width:100%;max-width:none;margin:5mm 0 3mm;padding:6px 8px;border-left-width:3px;border-radius:6px;font-size:8.5pt;line-height:1.3;break-inside:avoid;page-break-inside:avoid}
  .cluster-policy-card h2{margin-bottom:3px;font-size:11pt}
  .cluster-policy-card p{margin:2px 0;line-height:1.3}
  .cluster-policy-card .cluster-levels{gap:3px;margin:3px 0}
  .cluster-policy-card .semiconductor-cluster-badge{font-size:7pt;padding:1px 4px}
  .cluster-policy-card .policy-caution{padding:3px 5px}
  .cluster-policy-card a{font-size:7.5pt}
}
.monthly-diagnosis-section {
  break-after: auto;
}
.diagnosis-card .subhead,
.final-observation-card h3,
.summary-case-card h3 {
  margin-top: 16px;
}
.diagnosis-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 12px;
  margin: 14px 0 18px 0;
}
.diagnosis-box {
  border: 1px solid #e2e8f0;
  border-radius: 16px;
  padding: 14px 14px 12px 14px;
  background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
  box-shadow: 0 3px 10px rgba(15, 23, 42, 0.05);
}
.diagnosis-box span {
  display: inline-block;
  font-size: 12px;
  font-weight: 900;
  color: #2563eb;
  margin-bottom: 6px;
}
.diagnosis-box b {
  display: block;
  font-size: 16px;
  color: #0f172a;
  margin-bottom: 7px;
}
.diagnosis-box p {
  font-size: 13px;
  line-height: 1.65;
  color: #334155;
  margin: 0;
}
.basis-table td,
.basis-table th {
  text-align: center;
  vertical-align: middle;
}
.basis-table td:last-child,
.basis-table th:last-child {
  text-align: left;
}
.summary-case-card p {
  font-size: 14px;
  line-height: 1.8;
  color: #1f2937;
}
@media print {
  .diagnosis-grid {
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 9px;
  }
  .diagnosis-box {
    break-inside: avoid;
    page-break-inside: avoid;
    padding: 10px;
  }
  .diagnosis-box p {
    font-size: 11.5px;
  }
}


/* v41 연령대별 주민등록인구 시계열 증감률 표시 */
.age-pop-pct-cell {
  white-space: nowrap;
}
.age-count {
  font-weight: 700;
  color: #0f172a;
}
.age-pct {
  font-size: 9px;
  font-weight: 800;
}
.neutral-value {
  color: #64748b;
}
@media print {
  .age-pct {
    font-size: 7.8px !important;
  }
  .age-pop-pct-cell {
    font-size: 8px !important;
  }
}


/* v42 종합 시황 서술 보강: 기존 양식 유지 */
.chapter p {
  word-break: keep-all;
}
.summary-conclusion-card p {
  line-height: 1.85;
}


/* v43 Ⅵ장 물건유형별 필수 서술 보강 */
.market-type-summary-card h4 {
  margin-top: 20px;
  margin-bottom: 8px;
  padding: 8px 12px;
  border-left: 5px solid #1d4ed8;
  background: #eff6ff;
  color: #1e3a8a;
  border-radius: 8px;
  font-size: 16px;
  font-weight: 900;
}
.market-type-summary-card h5 {
  margin: 14px 0 6px 0;
  font-size: 14px;
  font-weight: 900;
  color: #0f172a;
}
.market-type-summary-card p {
  margin: 0 0 10px 0;
  line-height: 1.82;
  font-size: 14px;
  color: #1f2937;
  word-break: keep-all;
}
@media print {
  .market-type-summary-card h4,
  .market-type-summary-card h5,
  .market-type-summary-card p {
    break-inside: avoid;
    page-break-inside: avoid;
  }
}


/* v44 Ⅵ장 중복 제거 및 기존 양식 기반 전면 재작성 */
.final-market-summary-card h3 {
  margin-bottom: 12px;
}
.final-market-summary-card h4 {
  margin: 14px 0 6px 0;
  font-size: 14px;
  font-weight: 900;
  color: #0f172a;
}
.final-market-summary-card p {
  margin: 0 0 10px 0;
  line-height: 1.82;
  font-size: 14px;
  color: #1f2937;
  word-break: keep-all;
}
@media print {
  .final-market-summary-card {
    break-inside: auto;
    page-break-inside: auto;
  }
  .final-market-summary-card h4,
  .final-market-summary-card p {
    break-inside: avoid;
    page-break-inside: avoid;
  }
}


/* v45 데이터 해석형 분석문 및 Ⅵ장 배경색 구분 */
.summary-sale {
  background: #fff5f5 !important;
  border-left: 7px solid #ef4444 !important;
}
.summary-jeonse {
  background: #f0fdf4 !important;
  border-left: 7px solid #22c55e !important;
}
.summary-monthly {
  background: #fff7ed !important;
  border-left: 7px solid #f97316 !important;
}
.summary-total {
  background: #f8fafc !important;
  border-left: 7px solid #334155 !important;
}
.summary-sale h3,
.summary-sale h4 { color: #991b1b !important; }
.summary-jeonse h3,
.summary-jeonse h4 { color: #166534 !important; }
.summary-monthly h3,
.summary-monthly h4 { color: #9a3412 !important; }
.summary-total h3,
.summary-total h4 { color: #1e293b !important; }
.final-market-summary-card p {
  line-height: 1.88 !important;
}
.note {
  word-break: keep-all;
}


/* v46 Ⅵ장 실거래 주요사례·금액 기반 보강 */
.final-market-summary-card p {
  line-height: 1.9 !important;
}


/* v47 Ⅵ장 해당월 개별 실거래 사례 중심 보강 */
.final-market-summary-card p {
  line-height: 1.92 !important;
}


/* v49 Ⅵ장 주요 거래사례 한줄 공백 및 사례 표시 */
.case-spacer {
  height: 8px;
}
.major-case-line {
  margin: 4px 0 12px 0 !important;
  padding: 9px 12px;
  border-radius: 10px;
  background: rgba(255,255,255,0.72);
  border: 1px dashed #94a3b8;
  font-size: 13px !important;
  font-weight: 800;
  color: #0f172a !important;
  line-height: 1.65 !important;
}
@media print {
  .case-spacer {
    height: 6px;
  }
  .major-case-line {
    font-size: 11.5px !important;
    padding: 7px 9px;
  }
}


/* v51 저메모리 미리보기 및 최소 행 걸림 방지 */
.chart-box,
.mini-chart-box {
  contain: layout paint;
}
canvas {
  max-width: 100%;
}
.table-wrap {
  overflow-x: auto;
}
table tr {
  break-inside: avoid;
  page-break-inside: avoid;
}
.major-case-line {
  break-inside: avoid;
  page-break-inside: avoid;
}
@media print {
  .table-wrap {
    overflow: visible !important;
  }
  table thead {
    display: table-header-group;
  }
  table tr {
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }
  .chart-box,
  .mini-chart-box,
  .mini-chart-card {
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }
}


/* v53 비아파트 보강 분석표: 연립·다세대·단독·다가구 조건부 분석 및 Ⅵ장 문단 직접 반영 */
.nonapt-enhance-card {
  background: #ffffff;
  border: 1px solid #cbd5e1;
  border-left: 7px solid #475569;
  border-radius: 20px;
  padding: 20px;
  margin: 16px 0;
  page-break-inside: avoid;
}
.nonapt-enhance-card h3 {
  font-size: 18px;
  font-weight: 900;
  color: #0f172a;
  margin-bottom: 10px;
}
.nonapt-enhance-card h4 {
  font-size: 15px;
  font-weight: 900;
  color: #334155;
  margin: 16px 0 8px;
}
.nonapt-enhance-card .meaning {
  background: #f8fafc;
  border-left-color: #475569;
}
.nonapt-enhance-table th,
.nonapt-enhance-table td {
  text-align: center;
  vertical-align: middle;
  white-space: normal;
  word-break: keep-all;
}
.nonapt-enhance-table td:last-child,
.nonapt-enhance-table th:last-child {
  text-align: left;
}
.nonapt-rule-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 10px;
  margin: 10px 0 14px;
}
.nonapt-rule-box {
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  border-radius: 14px;
  padding: 10px 12px;
}
.nonapt-rule-box b {
  display: block;
  color: #0f172a;
  font-size: 13px;
  margin-bottom: 4px;
}
.nonapt-rule-box span {
  display: block;
  color: #475569;
  font-size: 12px;
  line-height: 1.45;
}
@media print {
  .nonapt-enhance-card {
    padding: 14px !important;
    margin: 10px 0 !important;
    break-inside: auto;
    page-break-inside: auto;
  }
  .nonapt-rule-grid {
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 7px;
  }
  .nonapt-rule-box {
    padding: 8px !important;
    break-inside: avoid;
    page-break-inside: avoid;
  }
  .nonapt-rule-box b,
  .nonapt-rule-box span {
    font-size: 9.5px !important;
  }
  .nonapt-enhance-table th,
  .nonapt-enhance-table td {
    font-size: 8.5px !important;
    padding: 4px 3px !important;
  }
}

/* v53 품질보정: 주택공급 자료범위·해석대상 명시 배지 가독성 개선 */
.supply-scope-badge {
  display: grid;
  grid-template-columns: 88px 1fr;
  gap: 6px 10px;
  margin: 6px 0 12px 0;
  padding: 10px 13px;
  border-radius: 14px;
  background: #eef2ff;
  border: 1px solid #c7d2fe;
  color: #1e3a8a;
  font-size: 13px;
  line-height: 1.45;
}
.supply-scope-badge span {
  color: #475569;
  font-weight: 900;
  text-align: center;
  background: #ffffff;
  border: 1px solid #dbeafe;
  border-radius: 999px;
  padding: 2px 8px;
}
.supply-scope-badge b {
  color: #1e3a8a;
  font-weight: 900;
}
.age-color-guide {
  margin: 8px 0 12px 0;
  padding: 8px 11px;
  border-left: 4px solid #64748b;
  background: #f8fafc;
  border-radius: 10px;
  color: #475569;
  font-size: 12px;
  font-weight: 700;
  line-height: 1.55;
}
@media print {
  .supply-scope-badge {
    grid-template-columns: 70px 1fr;
    font-size: 10.5px !important;
    padding: 7px 9px !important;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  .supply-scope-badge span { padding: 1px 5px !important; }
  .age-color-guide {
    font-size: 9.8px !important;
    padding: 6px 8px !important;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
}


/* v53 핵심 수치·핵심 서술 형광펜 강조 - 절제형 보정 */
.hl-key,
.hl-risk,
.hl-good,
.hl-info {
  display: inline;
  padding: 0 2px;
  border-radius: 3px;
  font-weight: 900;
  background-repeat: no-repeat;
  background-size: 100% 0.62em;
  background-position: 0 88%;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}
.hl-key {
  background-image: linear-gradient(transparent 0, transparent 0), linear-gradient(rgba(254, 240, 138, 0.88), rgba(254, 240, 138, 0.88));
}
.hl-risk {
  background-image: linear-gradient(transparent 0, transparent 0), linear-gradient(rgba(254, 202, 202, 0.92), rgba(254, 202, 202, 0.92));
  color: #991b1b;
}
.hl-good {
  background-image: linear-gradient(transparent 0, transparent 0), linear-gradient(rgba(187, 247, 208, 0.92), rgba(187, 247, 208, 0.92));
  color: #166534;
}
.hl-info {
  background-image: linear-gradient(transparent 0, transparent 0), linear-gradient(rgba(191, 219, 254, 0.92), rgba(191, 219, 254, 0.92));
  color: #1e3a8a;
}
.highlight-guide {
  max-width: 1280px;
  margin: 8px auto 22px auto;
  padding: 8px 12px;
  border-left: 5px solid #eab308;
  background: #fefce8;
  border: 1px solid #fde68a;
  border-left-width: 5px;
  border-radius: 12px;
  color: #713f12;
  font-size: 12.5px;
  line-height: 1.55;
  font-weight: 700;
}


/* v53 독자 빠른 읽기용 핵심 강조 체계 */
.read-key,
.read-risk,
.read-good,
.read-info,
.read-judgment,
.hl-judgment {
  display: inline;
  padding: 0 2px;
  border-radius: 3px;
  font-weight: 900;
  background-repeat: no-repeat;
  background-size: 100% 0.62em;
  background-position: 0 88%;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
}
.read-key {
  background-image: linear-gradient(rgba(254, 240, 138, 0.88), rgba(254, 240, 138, 0.88));
}
.read-risk {
  background-image: linear-gradient(rgba(254, 202, 202, 0.92), rgba(254, 202, 202, 0.92));
  color: #991b1b;
}
.read-good {
  background-image: linear-gradient(rgba(187, 247, 208, 0.92), rgba(187, 247, 208, 0.92));
  color: #166534;
}
.read-info {
  background-image: linear-gradient(rgba(191, 219, 254, 0.92), rgba(191, 219, 254, 0.92));
  color: #1e3a8a;
}
.read-judgment,
.hl-judgment {
  background-image: linear-gradient(rgba(191, 219, 254, 0.78), rgba(191, 219, 254, 0.78));
  color: #1e3a8a;
  font-weight: 950;
}
/* 결론부는 형광펜보다 판단문 굵기와 파란 밑칠을 우선 적용 */
.final-market-summary-card .read-judgment,
.final-market-summary-card .hl-judgment {
  color: #1e3a8a;
  font-weight: 950;
}
/* 표 내부에서는 강조가 과하지 않도록 수치 중심만 짧게 표시 */
table .read-key,
table .read-risk,
table .read-good,
table .read-info,
table .read-judgment,
table .hl-judgment {
  background-size: 100% 0.5em;
  background-position: 0 90%;
}

/* 상단 KPI 카드는 전문 보고서 첫인상 유지를 위해 형광펜 바를 사용하지 않고 수치 강조만 적용 */
.kpis .hl-key,
.kpis .hl-risk,
.kpis .hl-good,
.kpis .hl-info {
  background: none !important;
  padding: 0 !important;
  border-radius: 0 !important;
  box-decoration-break: initial;
  -webkit-box-decoration-break: initial;
}
@media print {
  .hl-key, .hl-risk, .hl-good, .hl-info, .highlight-guide {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  .highlight-guide {
    font-size: 10.5px !important;
    padding: 6px 9px !important;
    margin: 6px auto 14px auto !important;
  }
}

/* v53 당월 직전거래 대비 TOP 거래사례 카드 */
.top-case-card { page-break-inside: avoid; }
.top-case-grid { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:12px; margin-top:12px; }
.top-case-list { background:#f8fafc; border:1px solid #e2e8f0; border-radius:16px; padding:12px; }
.top-case-list h4 { margin:0 0 8px 0 !important; font-size:14px !important; font-weight:900 !important; color:#1e293b; }
.top-case-item { display:grid; grid-template-columns:24px minmax(0,1fr) auto; gap:8px; padding:8px 0; border-top:1px dashed #cbd5e1; align-items:start; }
.top-case-item:first-of-type { border-top:0; }
.deal-badge,.floor-badge { display:inline-block; margin-left:6px; padding:1px 8px; border-radius:999px; font-style:normal; font-weight:800; font-size:12px; line-height:1.5; vertical-align:middle; white-space:nowrap; }
.deal-badge.deal-direct { color:#fff; background:#dc2626; }
.deal-badge.deal-broker { color:#065f46; background:#d1fae5; border:1px solid #6ee7b7; }
.deal-badge.deal-na { color:#64748b; background:#f1f5f9; border:1px solid #cbd5e1; font-weight:600; }
.floor-badge.floor-up { color:#b42318; background:#fee4e2; border:1px solid #fda29b; }
.floor-badge.floor-down { color:#1d4ed8; background:#dbeafe; border:1px solid #93c5fd; }
.floor-badge.floor-same { color:#475569; background:#f1f5f9; border:1px solid #cbd5e1; }
.outlier-badge { display:inline-block; margin-left:6px; padding:1px 8px; border-radius:6px; font-style:normal; font-weight:800; font-size:12px; color:#92400e; background:#fef3c7; border:1px solid #f59e0b; white-space:nowrap; vertical-align:middle; }
.policy-inflection-note { background:#fffbeb; border:1px solid #fcd34d; border-radius:10px; padding:10px 14px; margin:10px 0; font-size:12.5px; line-height:1.7; }
.policy-inflection-note ul { margin:6px 0 0; padding-left:18px; }
.policy-inflection-note .pi-events { color:#7c2d12; font-size:12px; }
.policy-inflection-note .pi-none { color:#94a3b8; }
.policy-inflection-note .pi-base { color:#b45309; font-size:11px; }
.policy-inflection-note .pi-foot { font-size:11px; color:#92400e; margin:6px 0 0; }
.case-rank { font-weight:900; color:#111827; text-align:center; }
.case-body b { display:block; font-size:12.5px; color:#0f172a; word-break:keep-all; }
.case-body span { display:block; font-size:11px; color:#64748b; line-height:1.35; margin-top:2px; }
.case-body small { display:block; font-size:10.5px; color:#475569; line-height:1.35; margin-top:2px; }
.case-year { color:#475569; font-weight:800; }
.case-value { text-align:right; min-width:72px; }
.case-value b { display:block; font-size:13px; color:#0f172a; white-space:nowrap; }
.case-value span { display:block; font-size:11px; font-weight:900; white-space:nowrap; }
.case-body span.case-em:not(#__case_em) { display:inline !important; margin:0 !important; font-size:inherit !important; line-height:inherit !important; font-weight:900 !important; color:#0f172a !important; white-space:nowrap !important; }
/* TOP10 카드 v2: 위치·면적 한 줄, 배지 한 줄, 현재·직전 비교 2줄, 오른쪽 변동률 */
.top-case-item.tc-v2:not(#__tc) { grid-template-columns:22px minmax(0,1fr) 74px; gap:8px; padding:10px 0; align-items:start; }
.tc-v2 .case-rank { margin-top:1px; width:22px; height:22px; line-height:22px; border-radius:50%; background:#e2e8f0; color:#0f172a; font-size:12px; }
.tc-v2 .case-body > b { font-size:13.5px; line-height:1.35; }
.tc-v2 .case-body .tc-meta { margin-top:3px; font-size:11.5px; color:#475569; line-height:1.5; word-break:keep-all; }
.tc-v2 .case-body .tc-meta span:not(#__tc) { display:inline !important; margin:0 !important; font-size:inherit !important; line-height:inherit !important; }
.tc-v2 .case-body .tc-meta .case-year { color:#64748b; font-weight:700; }
.tc-v2 .case-body .tc-meta .tc-seg { white-space:nowrap; }
.tc-v2 .tc-badges { margin-top:4px; display:flex; flex-wrap:wrap; gap:4px; }
.tc-v2 .tc-badges em { margin-left:0 !important; font-size:11px !important; padding:0 7px !important; }
.tc-v2 .tc-compare { grid-column:2 / -1; margin-top:-2px; border:1px solid #e2e8f0; border-radius:8px; overflow:hidden; background:#fff; }
.tc-v2 .tc-row { display:grid; grid-template-columns:32px 58px 36px minmax(0,1fr); align-items:center; gap:4px; padding:3px 6px; font-size:11.5px; color:#334155; }
.tc-v2 .tc-row.tc-head { padding:2px 6px; background:#f1f5f9; font-size:10px; color:#94a3b8; font-weight:700; }
.tc-v2 .tc-row.tc-cur { font-weight:800; color:#0f172a; }
.tc-v2 .tc-row.tc-prev { border-top:1px dashed #e2e8f0; color:#64748b; }
.tc-v2 .tc-label { font-size:10px; font-weight:900; text-align:center; border-radius:4px; padding:0 2px; }
.tc-v2 .tc-cur .tc-label { background:#0f172a; color:#fff; }
.tc-v2 .tc-prev .tc-label { background:#e2e8f0; color:#475569; }
.tc-v2 .tc-amt { text-align:right; white-space:normal; word-break:keep-all; line-height:1.35; }
.tc-v2 .tc-amt .nw { white-space:nowrap; }
.tc-v2 .tc-amt small { display:block; font-size:10px; font-weight:600; color:#94a3b8; }
.tc-v2 .tc-cur .tc-amt { font-size:12.5px; }
.tc-v2 .case-value { text-align:right; min-width:0; padding-top:2px; }
.tc-v2 .case-value .tc-rate:not(#__tc) { display:block; font-size:16px; font-weight:950; line-height:1.2; white-space:nowrap; }
.tc-v2 .case-value .tc-delta:not(#__tc) { display:block; margin-top:2px; font-size:11px; font-weight:800; white-space:nowrap; }
@media print {
  .top-case-item.tc-v2:not(#__tc) { grid-template-columns:16px minmax(0,1fr) 52px; gap:4px; padding:5px 0; }
  .tc-v2 .case-rank { width:16px; height:16px; line-height:16px; font-size:8.5px; }
  .tc-v2 .case-body > b { font-size:9.8px !important; }
  .tc-v2 .case-body .tc-meta, .tc-v2 .tc-row { font-size:8.2px !important; }
  .tc-v2 .tc-row { grid-template-columns:20px 40px 24px minmax(0,1fr); padding:1px 3px; }
  .tc-v2 .tc-cur .tc-amt { font-size:8.8px !important; }
  .tc-v2 .case-value .tc-rate:not(#__tc) { font-size:11px !important; }
  .tc-v2 .case-value .tc-delta:not(#__tc) { font-size:8px !important; }
}
.case-value em.rent-pair { display:block; font-style:normal; font-size:12.5px; font-weight:900; color:#0f172a; line-height:1.35; }
.case-value em.rent-converted { display:block; font-style:normal; font-size:10.5px; font-weight:700; color:#64748b; margin-top:1px; }
/* 차트 캔버스는 담는 상자 크기를 그대로 채운다. 높이를 따로 고정하면 Chart.js가 계산한 크기와 화면 크기가 어긋나
   글자가 늘어나 보이고 범례 클릭(표시 on/off) 위치가 빗나가므로, 앞선 고정 높이 규칙보다 우선하도록 :not(#…)로 덮어쓴다. */
.chart-box canvas[id]:not(#__chart_size_guard), .chart-wrap canvas[id]:not(#__chart_size_guard),
.mini-chart-box canvas[id]:not(#__chart_size_guard), .survey-donut-box canvas[id]:not(#__chart_size_guard) {
  width:100% !important; height:100% !important; min-height:0 !important; max-height:none !important;
}
.case-rise .case-value span { color:#dc2626; }
.case-fall .case-value span { color:#2563eb; }
.case-top .case-value span { color:#475569; }
/* 거래금액·보증금·환산월세 상위 TOP10도 직전거래 대비 증감 방향에 따라 색상 적용 */
.case-top.case-rise .case-value span { color:#dc2626 !important; }
.case-top.case-fall .case-value span { color:#2563eb !important; }
.case-top.case-zero .case-value span { color:#64748b !important; }
.top-case-empty { font-size:12px; color:#64748b; padding:10px; background:white; border-radius:10px; }
@media print { .top-case-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); gap:8px; } .top-case-list { padding:8px !important; break-inside: avoid; page-break-inside: avoid; } .top-case-item { grid-template-columns:18px minmax(0,1fr); gap:5px; padding:5px 0; } .case-value { grid-column:2; text-align:left; min-width:0; } .case-body b { font-size:9.8px !important; } .case-body span, .case-body small, .case-value span { font-size:8.2px !important; } .case-value b { font-size:9.5px !important; } }


/* TOP10 거래사례 층 정보 표시 보강 */
.case-body span { line-height: 1.45; }
.case-body small { line-height: 1.45; }


/* v53 조사 담당자 관점 핵심진단 + 매매·전세·월세 도넛 차트 */
.survey-donut-section {
  margin-top: 18px;
  padding: 16px;
  border: 1px solid #dbeafe;
  border-radius: 18px;
  background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
}
.survey-donut-section h4 {
  margin: 0 0 6px 0 !important;
  font-size: 16px !important;
  font-weight: 900 !important;
  color: #0f172a;
}
.survey-donut-guide {
  margin: 6px 0 12px;
  padding: 9px 12px;
  background: #eef2ff;
  border-left: 5px solid #4f46e5;
  border-radius: 12px;
  color: #1e293b;
  font-size: 12px;
  font-weight: 700;
  line-height: 1.55;
}
.survey-donut-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 14px;
}
.survey-donut-card {
  background: #fff;
  border: 1px solid #e2e8f0;
  border-radius: 16px;
  padding: 12px 12px 10px;
  box-shadow: 0 4px 14px rgba(15,23,42,.04);
  page-break-inside: avoid;
  break-inside: avoid;
}
.survey-donut-title {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin-bottom: 8px;
}
.survey-donut-title b {
  font-size: 14px;
  font-weight: 900;
  color: #0f172a;
}
.survey-donut-title span {
  font-size: 11px;
  font-weight: 800;
  color: #475569;
  background: #f1f5f9;
  border-radius: 999px;
  padding: 3px 7px;
  white-space: nowrap;
}
.survey-donut-box {
  position: relative;
  height: 260px;
  min-height: 260px;
}
.survey-donut-box canvas {
  width: 100% !important;
  height: 100% !important;
}
.survey-donut-caption {
  margin-top: 8px;
  font-size: 12px;
  line-height: 1.55;
  color: #334155;
  font-weight: 700;
}
.survey-donut-caption .rise { color: #dc2626; font-weight: 900; }
.survey-donut-caption .flat { color: #64748b; font-weight: 900; }
.survey-donut-caption .fall { color: #2563eb; font-weight: 900; }

.survey-donut-analysis {
  margin-top: 14px;
  padding: 12px 14px;
  border: 1px solid #e2e8f0;
  border-left: 6px solid #1d4ed8;
  border-radius: 14px;
  background: #f8fafc;
}
.survey-donut-analysis b {
  display: block;
  margin-bottom: 6px;
  color: #1e3a8a;
  font-size: 13px;
  font-weight: 900;
}
.survey-donut-analysis p {
  margin: 0 0 6px 0;
  color: #1e293b;
  font-size: 13px;
  line-height: 1.68;
  word-break: keep-all;
}
.survey-donut-analysis p:last-child { margin-bottom: 0; }
.survey-donut-analysis .label {
  display: inline-block;
  min-width: 34px;
  font-weight: 900;
  color: #0f172a;
}
@media (max-width: 900px) {
  .survey-donut-grid { grid-template-columns: 1fr; }
}
@media print {
  .survey-donut-section { padding: 12px !important; break-inside: avoid; page-break-inside: avoid; }
  .survey-donut-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
  .survey-donut-card { padding: 9px !important; }
  .survey-donut-box { height: 165px !important; min-height: 165px !important; }
  .survey-donut-caption { font-size: 9.5px !important; }
  .survey-donut-title b { font-size: 11px !important; }
  .survey-donut-title span { font-size: 8.5px !important; }
  .survey-donut-analysis { padding: 9px 10px !important; margin-top: 9px !important; }
  .survey-donut-analysis b { font-size: 10.5px !important; }
  .survey-donut-analysis p { font-size: 9.5px !important; line-height: 1.45 !important; }
}



/* v58 지역·월 변경 안정성 및 전체 차트 품질 보정 */
.data-empty-note{display:block;margin:8px 0 0;padding:9px 11px;border-radius:10px;background:#f8fafc;border:1px dashed #cbd5e1;color:#475569;font-size:12px;font-weight:800;word-break:keep-all;}
.chart-quality-note{font-size:12px;color:#475569;font-weight:700;line-height:1.55;margin-top:6px;word-break:keep-all;}
@media (max-width: 768px){
  .hero{margin:10px 8px!important;padding:18px!important;border-radius:18px!important;}
  .chapter{margin:12px 8px!important;padding:14px!important;border-radius:18px!important;}
  .kpis{grid-template-columns:1fr!important;margin:0 8px 12px!important;}
  .diagnosis-grid,.conversion-grid,.mini-chart-grid,.total-stage-grid{grid-template-columns:1fr!important;}
  .table-wrap{overflow-x:auto!important;-webkit-overflow-scrolling:touch!important;}
  .table-wrap table{min-width:720px;}
  .chart-box{min-height:280px!important;height:300px!important;}
  .chart-box canvas{height:260px!important;}
}



/* === v5.8.38 A4 출력·template_v58 시각 일치 보강 === */
@page { size:A4; margin:15mm; }
html, body { width:100%; min-height:100%; }
body { font-family:'Pretendard','Noto Sans KR','Malgun Gothic',Arial,sans-serif !important; background:#f8fafc; }
.chart-note{font-size:12px;color:#5c6f82;background:#f8fafc;border-radius:8px;padding:8px;margin-top:8px;line-height:1.5}
.result-box{background:#fff7e6;border-left:6px solid #f59f00;border-radius:12px;padding:13px 15px;line-height:1.75;margin:12px 0 20px}
.hl-yellow{background:#fff3bf;font-weight:700;padding:1px 4px;border-radius:4px}
.hl-red{background:#ffe3e3;font-weight:700;color:#b42318;padding:1px 4px;border-radius:4px}
.hl-green{background:#d3f9d8;font-weight:700;color:#087f5b;padding:1px 4px;border-radius:4px}
.hl-blue{background:#d0ebff;font-weight:700;color:#1864ab;padding:1px 4px;border-radius:4px}
.footer{max-width:1280px;margin:40px auto;color:#777;font-size:12px;border-top:1px solid #d9e2ec;padding-top:14px}
.muted{color:#777}
td.pos{color:#dc2626!important;font-weight:800}
td.neg{color:#2563eb!important;font-weight:800}
.survey-donut-grid .chart-box{background:white;border:1px solid #e2e8f0;border-radius:18px;padding:14px;box-shadow:0 6px 18px rgba(15,23,42,.05)}
.chart-box canvas{max-width:100%}
.chart-box.mix-chart{height:340px!important;min-height:340px!important;max-height:360px!important}
.chart-box.mix-chart canvas{height:280px!important}
.mini-chart-grid .chart-box{min-height:260px!important;height:260px!important}
.mini-chart-grid .chart-wrap{height:210px!important}
.chapter{break-inside:auto;page-break-inside:auto}
.card{break-inside:auto;page-break-inside:auto}
.no-break,.chart-box,.top-case-item,.conversion-card{break-inside:avoid;page-break-inside:avoid}
.table-wrap{overflow-x:visible!important}
table{width:100%;border-collapse:collapse;table-layout:auto}
th,td{word-break:keep-all}
@media print {
  * { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  body { background:white!important; font-family:'Pretendard','Noto Sans KR','Malgun Gothic',Arial,sans-serif!important; }
  .print-guide { display:none!important; }
  .hero { padding:18px!important; margin:0 auto 8px!important; border-radius:18px!important; max-width:186mm!important; }
  .notice,.kpis,.chapter,.footer{ max-width:186mm!important; }
  .kpis{grid-template-columns:repeat(4,1fr)!important;gap:8px!important;margin-bottom:8px!important}
  .kpis div{padding:10px!important;border-radius:12px!important;box-shadow:none!important}
  .kpis b{font-size:18px!important}
  .chapter{padding:12px!important;margin:8px auto!important;border-radius:16px!important}
  .chapter>h2{font-size:18px!important;margin-bottom:8px!important}
  .card{padding:10px!important;margin:8px 0!important;border-radius:14px!important;box-shadow:none!important}
  .card h3{font-size:13.8px!important;line-height:1.25!important}
  .card h4{font-size:12px!important;line-height:1.25!important}
  .meaning,.note,.result-box{font-size:9.8px!important;line-height:1.42!important;padding:7px 9px!important;margin-top:7px!important}
  .table-wrap{margin:7px 0!important;overflow:visible!important}
  table{font-size:8.2px!important;table-layout:fixed!important}
  th,td{padding:3.2px 2.4px!important;white-space:normal!important;line-height:1.24!important}
  .chart-box{height:225px!important;min-height:225px!important;margin:8px 0 10px!important;padding:8px!important}
  .chart-wrap{height:184px!important}
  .chart-box canvas{height:184px!important;max-height:184px!important}
  .chart-box.mix-chart{height:255px!important;min-height:255px!important}
  .chart-box.mix-chart canvas,.chart-box.mix-chart .chart-wrap{height:210px!important;max-height:210px!important}
  .mini-chart-grid{grid-template-columns:repeat(2,1fr)!important;gap:8px!important}
  .mini-chart-grid .chart-box{height:190px!important;min-height:190px!important}
  .mini-chart-grid .chart-wrap,.mini-chart-grid canvas{height:150px!important}
    .cover-break{break-after:page;page-break-after:always}
}


/* === v5.8.38 템플릿에 없는 차트 제거 및 차트 균형 보정 === */
.chart-box{min-height:320px!important;height:320px!important;max-height:340px!important;margin:12px 0 18px!important;padding:10px!important;overflow:hidden!important}
.chart-box canvas{height:280px!important;max-height:280px!important}
.chart-wrap{height:280px!important;max-height:280px!important}
.chart-box.mix-chart{min-height:340px!important;height:340px!important;max-height:360px!important}
.chart-box.mix-chart canvas,.chart-box.mix-chart .chart-wrap{height:280px!important;max-height:280px!important}
.mini-chart-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin:10px 0 20px}
.mini-chart-grid .chart-box{min-height:240px!important;height:240px!important;max-height:250px!important}
.mini-chart-grid .chart-wrap,.mini-chart-grid canvas{height:200px!important;max-height:200px!important}
.survey-donut-box{height:260px!important;min-height:260px!important}
.survey-donut-box canvas{height:260px!important}
#template-v58-canvas-bank,#template-v58-class-bank{display:none!important}
@media print{
  .chart-box{height:245px!important;min-height:245px!important;max-height:255px!important;margin:10px 0 14px!important}
  .chart-box canvas,.chart-wrap{height:205px!important;max-height:205px!important}
  .chart-box.mix-chart{height:270px!important;min-height:270px!important;max-height:280px!important}
  .chart-box.mix-chart canvas,.chart-box.mix-chart .chart-wrap{height:225px!important;max-height:225px!important}
  .mini-chart-grid .chart-box{height:185px!important;min-height:185px!important;max-height:195px!important}
  .mini-chart-grid .chart-wrap,.mini-chart-grid canvas{height:145px!important;max-height:145px!important}
  .survey-donut-box{height:155px!important}
  .survey-donut-box canvas{height:155px!important}
}


/* === v5.8.38 PDF 출력 안정화 보정 === */
@media print {
  @page { size:A4; margin:15mm; }

  html, body {
    width:auto !important;
    height:auto !important;
    min-height:auto !important;
    overflow:visible !important;
  }

  body {
    background:#ffffff !important;
    font-family:'Pretendard','Noto Sans KR',Arial,sans-serif !important;
  }

  .hero, .notice, .kpis, .chapter, .footer {
    max-width:186mm !important;
  }

  /* 표준템플릿 v58 표지처럼 KPI는 세로 카드형 유지 */
  .kpis {
    display:grid !important;
    grid-template-columns:1fr !important;
    gap:10px !important;
    margin:12px auto !important;
  }
  .kpis div {
    min-height:54px !important;
    padding:13px 18px !important;
    border-radius:14px !important;
    box-shadow:0 6px 18px rgba(15,23,42,.06) !important;
  }
  .kpis b {
    font-size:22px !important;
  }

  /* 후반부 누락 방지: 큰 컨테이너 전체 avoid 금지 */
  .chapter, .card, .top-case-card, .chart-box, .mini-chart-card, .conversion-card, .supply-card, .age-move-card {
    overflow:visible !important;
    break-inside:auto !important;
    page-break-inside:auto !important;
    height:auto !important;
    max-height:none !important;
  }

  .chart-box {
    min-height:245px !important;
    margin:10px 0 14px !important;
  }
  .chart-wrap, .chart-box canvas {
    height:205px !important;
    max-height:205px !important;
  }
  .chart-box.mix-chart {
    min-height:275px !important;
  }
  .chart-box.mix-chart .chart-wrap,
  .chart-box.mix-chart canvas {
    height:230px !important;
    max-height:230px !important;
  }

  .mini-chart-box {
    height:190px !important;
  }
  .mini-chart-box canvas {
    height:170px !important;
    max-height:170px !important;
  }
  .survey-donut-box {
    height:150px !important;
  }
  .survey-donut-box canvas {
    height:150px !important;
  }

  /* TOP10: 제목만 남고 다음 페이지로 밀리는 현상 방지 */
  .top-case-card {
    margin-top:12px !important;
    padding:12px !important;
    break-before:auto !important;
    page-break-before:auto !important;
  }
  .top-case-grid {
    display:grid !important;
    grid-template-columns:repeat(3,1fr) !important;
    gap:9px !important;
    align-items:start !important;
  }
  .top-case-list {
    break-inside:auto !important;
    page-break-inside:auto !important;
  }
  .top-case-item {
    break-inside:avoid !important;
    page-break-inside:avoid !important;
    min-height:auto !important;
  }

  /* 빈 페이지 유발 가능성이 있는 강제 높이·공백 제거 */
  .cover-break {
    break-after:auto !important;
    page-break-after:auto !important;
  }

  table, tr, td, th {
    page-break-inside:auto !important;
    break-inside:auto !important;
  }

  .final-summary-chapter,
  #chapter-final-summary {
    display:block !important;
    break-before:page !important;
    page-break-before:always !important;
  }
}


/* === v5.8.38 차트 중복·비아파트 유형별·인구세대 차트 보정 === */
.population-mix-chart .chart-wrap{height:300px!important}
@media print{
  .population-mix-chart .chart-wrap,
  .population-mix-chart canvas{height:230px!important;max-height:230px!important}
}


/* === v5.8.38 상승·보합·하락 비율표 이모지·컬러 보강 === */
.market-strength-table th:nth-child(n),
.compact-wide-table th:nth-child(n) {
  white-space: nowrap;
}
.market-strength-table td.dir-up,
.compact-wide-table td.dir-up {
  color:#dc2626 !important;
  font-weight:900 !important;
  background:linear-gradient(90deg, rgba(254,226,226,.70), rgba(255,255,255,0)) !important;
}
.market-strength-table td.dir-flat,
.compact-wide-table td.dir-flat {
  color:#475569 !important;
  font-weight:900 !important;
  background:linear-gradient(90deg, rgba(241,245,249,.85), rgba(255,255,255,0)) !important;
}
.market-strength-table td.dir-down,
.compact-wide-table td.dir-down {
  color:#2563eb !important;
  font-weight:900 !important;
  background:linear-gradient(90deg, rgba(219,234,254,.75), rgba(255,255,255,0)) !important;
}
.market-strength-table .latest-row td.dir-up,
.compact-wide-table .latest-row td.dir-up {
  background:#fee2e2 !important;
}
.market-strength-table .latest-row td.dir-flat,
.compact-wide-table .latest-row td.dir-flat {
  background:#f1f5f9 !important;
}
.market-strength-table .latest-row td.dir-down,
.compact-wide-table .latest-row td.dir-down {
  background:#dbeafe !important;
}
/* 상승비율·하락비율·보합비율 판정란(상승우세/하락우세/보합우세/혼조/산출불가) */
.market-strength-table td.verdict-cell,
.compact-wide-table td.verdict-cell {
  font-weight:900 !important;
  white-space:nowrap !important;
  text-align:center !important;
}
.market-strength-table td.dir-mixed,
.compact-wide-table td.dir-mixed {
  color:#b45309 !important;
  font-weight:900 !important;
  background:linear-gradient(90deg, rgba(254,243,199,.80), rgba(255,255,255,0)) !important;
}
.market-strength-table td.dir-na,
.compact-wide-table td.dir-na {
  color:#94a3b8 !important;
  font-weight:700 !important;
  background:linear-gradient(90deg, rgba(241,245,249,.55), rgba(255,255,255,0)) !important;
}
.market-strength-table .latest-row td.dir-mixed,
.compact-wide-table .latest-row td.dir-mixed {
  background:#fef3c7 !important;
}
.market-strength-table .latest-row td.dir-na,
.compact-wide-table .latest-row td.dir-na {
  background:#f1f5f9 !important;
}
@media print {
  .market-strength-table td.dir-up,
  .market-strength-table td.dir-flat,
  .market-strength-table td.dir-down,
  .market-strength-table td.dir-mixed,
  .market-strength-table td.dir-na,
  .compact-wide-table td.dir-up,
  .compact-wide-table td.dir-flat,
  .compact-wide-table td.dir-down,
  .compact-wide-table td.dir-mixed,
  .compact-wide-table td.dir-na {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
}


/* === v5.8.38 연령대별 인구 시계열 및 전월 대비 증감률 표 === */
/* v5.8.51 A4 출력·화면 모두 한 칸 안에서 숫자와 증감률이 겹치는 문제 해결.
   화면도 인쇄와 마찬가지로 인구수(상단) / 증감률(하단) 두 줄 구조로 표시. */
.age-pop-trend-table-wrap { overflow-x:auto !important; }
.age-pop-trend-table { table-layout:fixed !important; width:100% !important; }
.age-pop-trend-table th { text-align:center !important; white-space:nowrap !important; }
.age-pop-trend-table td {
  text-align:center !important;
  white-space:normal !important;
  word-break:keep-all !important;
  overflow:hidden;
  font-weight:700;
  line-height:1.15 !important;
  padding:6px 4px !important;
}
.age-pop-trend-table td .age-rate {
  display:block !important;
  margin-top:2px !important;
  font-size:0.86em !important;
  line-height:1.1 !important;
  letter-spacing:-0.02em !important;
  white-space:nowrap !important;
}
.age-pop-trend-table th:first-child,
.age-pop-trend-table td:first-child { width:11% !important; white-space:nowrap !important; }
.age-pop-trend-table td.age-up .age-rate,
.age-pop-trend-table .age-rate.pos { color:#dc2626 !important; font-weight:900; }
.age-pop-trend-table td.age-down .age-rate,
.age-pop-trend-table .age-rate.neg { color:#2563eb !important; font-weight:900; }
.age-pop-trend-table td.age-flat .age-rate,
.age-pop-trend-table .age-rate.neutral { color:#64748b !important; font-weight:800; }
.age-pop-trend-table .latest-row td { background:#fff8bf !important; }
@media print {
  /* v5.8.51: A4 출력 시 한 칸 안에서 인구수와 증감률이 겹치지 않도록
     숫자 아래로 증감률을 줄바꿈해서 두 줄 구조로 표시. */
  .age-pop-trend-table { font-size:8.1px !important; table-layout:fixed !important; width:100% !important; }
  .age-pop-trend-table th,
  .age-pop-trend-table td { padding:4px 2px !important; white-space:normal !important; overflow:hidden; word-break:keep-all !important; }
  .age-pop-trend-table td .age-rate {
    display:block !important;
    margin-top:1px !important;
    font-size:7.4px !important;
    line-height:1.1 !important;
    letter-spacing:-0.02em !important;
  }
  .age-pop-trend-table th:first-child,
  .age-pop-trend-table td:first-child { width:13% !important; }
  .age-pop-trend-table th:not(:first-child),
  .age-pop-trend-table td:not(:first-child) { width:auto !important; }
}


/* === v5.8.38 연령대 순이동 통합 차트 10세 구간·동일 색상 적용 === */
.age-pop-card .chart-box:has(#chartAgeNetYoung) .chart-subtitle {
  letter-spacing:-0.02em;
}


/* === v5.8.38 인구·세대수 및 인구이동 통합표 공표월·가로잘림 보정 === */
.population-publish-note {
  background:#f8fafc;
  border:1px solid #e2e8f0;
  border-radius:12px;
  padding:9px 12px;
  margin:8px 0 10px;
  color:#334155;
  font-weight:700;
  font-size:12px;
}
.population-migration-table-wrap {
  overflow-x:auto !important;
}
.population-migration-table {
  min-width:820px;
  table-layout:fixed;
}
.population-migration-table th,
.population-migration-table td {
  text-align:center !important;
  white-space:nowrap !important;
}
.population-migration-table th:nth-child(1){width:90px}
.population-migration-table th:nth-child(2){width:110px}
.population-migration-table th:nth-child(3){width:115px}
.population-migration-table th:nth-child(4){width:105px}
.population-migration-table th:nth-child(5){width:120px}
.population-migration-table th:nth-child(6),
.population-migration-table th:nth-child(7),
.population-migration-table th:nth-child(8){width:90px}
.not-published {
  color:#92400e;
  background:#fef3c7;
  border-radius:999px;
  padding:2px 7px;
  font-weight:900;
}
.age-migration-summary-table td {
  text-align:center !important;
}
@media print {
  .population-migration-table { min-width:0 !important; width:100% !important; font-size:8.2px !important; }
  .population-migration-table th,
  .population-migration-table td { padding:4px 3px !important; }
  .population-publish-note { font-size:9px !important; padding:6px 8px !important; }
}




/* === v5.8.38 최근 13개월 연령대별 순이동 시계열 표 === */
.age-migration-trend-table-wrap { overflow-x:auto !important; }
.age-migration-trend-table th { text-align:center !important; white-space:nowrap !important; }
.age-migration-trend-table td { text-align:center !important; white-space:nowrap !important; font-weight:800; }
.age-migration-trend-table td.age-net-pos span { color:#dc2626 !important; font-weight:900; }
.age-migration-trend-table td.age-net-neg span { color:#2563eb !important; font-weight:900; }
.age-migration-trend-table td.age-net-zero span { color:#64748b !important; font-weight:900; }
.age-migration-trend-table .latest-row td { background:#fff8bf !important; }
@media print {
  .age-migration-trend-table { font-size:8.1px !important; }
  .age-migration-trend-table th,
  .age-migration-trend-table td { padding:4px 3px !important; }
}


/* === v5.8.38 공급 파이프라인 총괄표 표준양식 보정 === */
.supply-summary-table th,
.supply-summary-table td {
  text-align:center !important;
  white-space:nowrap !important;
}
.supply-summary-table .latest-row td {
  background:#fff8bf !important;
  font-weight:900 !important;
  color:#78350f !important;
}
.supply-summary-table-wrap {
  overflow-x:auto !important;
}
@media print {
  .supply-summary-table {
    font-size:9px !important;
  }
  .supply-summary-table th,
  .supply-summary-table td {
    padding:5px 4px !important;
  }
}


/* === v5.8.38 물건유형별 공급 파이프라인 시계열 표준양식 보정 === */
.supply-type-pipeline-table-wrap {
  overflow-x:auto !important;
}
.supply-type-pipeline-table {
  min-width:860px;
  table-layout:fixed;
}
.supply-type-pipeline-table th,
.supply-type-pipeline-table td {
  text-align:center !important;
  white-space:nowrap !important;
}
.supply-type-pipeline-table th:nth-child(1),
.supply-type-pipeline-table td:nth-child(1){width:92px}
.supply-type-pipeline-table th:nth-child(2),
.supply-type-pipeline-table td:nth-child(2){width:92px}
.supply-type-pipeline-table th:last-child,
.supply-type-pipeline-table td:last-child{
  font-weight:900 !important;
  color:#334155 !important;
  background:#f8fafc !important;
}
@media print {
  .supply-type-pipeline-table {
    min-width:0 !important;
    width:100% !important;
    font-size:8.2px !important;
  }
  .supply-type-pipeline-table th,
  .supply-type-pipeline-table td {
    padding:4px 3px !important;
  }
}


/* === v5.8.38 미분양 주택 동향표 표준양식 보정 === */
.unsold-standard-table-wrap {
  overflow-x:auto !important;
}
.unsold-standard-table th,
.unsold-standard-table td {
  text-align:center !important;
  white-space:nowrap !important;
}
.unsold-standard-table .latest-row td {
  background:#fff8bf !important;
  font-weight:900 !important;
}
.unsold-standard-table td .pos,
.unsold-standard-table .pos {
  color:#dc2626 !important;
  font-weight:900 !important;
}
.unsold-standard-table td .neg,
.unsold-standard-table .neg {
  color:#2563eb !important;
  font-weight:900 !important;
}
.unsold-standard-table td .neutral,
.unsold-standard-table .neutral {
  color:#64748b !important;
  font-weight:900 !important;
}
@media print {
  .unsold-standard-table {
    font-size:9px !important;
  }
  .unsold-standard-table th,
  .unsold-standard-table td {
    padding:5px 4px !important;
  }
}


/* === v5.8.38 예금은행 대출금리 전월 대비 증감 표준양식 보정 === */
.rates-standard-table-wrap {
  overflow-x:auto !important;
}
.rates-standard-table {
  min-width:980px;
}
.rates-standard-table th,
.rates-standard-table td {
  text-align:center !important;
  white-space:nowrap !important;
}
.rates-standard-table .latest-row td {
  background:#fff8bf !important;
  font-weight:900 !important;
}
.rates-standard-table .pos {
  color:#dc2626 !important;
  font-weight:900 !important;
}
.rates-standard-table .neg {
  color:#2563eb !important;
  font-weight:900 !important;
}
.rates-standard-table .neutral {
  color:#64748b !important;
  font-weight:900 !important;
}
@media print {
  .rates-standard-table {
    min-width:0 !important;
    width:100% !important;
    font-size:7.6px !important;
  }
  .rates-standard-table th,
  .rates-standard-table td {
    padding:4px 2px !important;
  }
}


/* === v5.8.38 인구수·세대수 이중축 차트 표준양식 보정 === */
.population-household-chart {
  padding:18px 18px 20px !important;
}
.population-household-chart .chart-wrap {
  height:330px !important;
}
.population-household-chart .chart-subtitle {
  font-size:17px !important;
  font-weight:900 !important;
  color:#1e293b !important;
  text-align:center !important;
  margin-bottom:10px !important;
}
@media print {
  .population-household-chart .chart-wrap,
  .population-household-chart canvas {
    height:260px !important;
    max-height:260px !important;
  }
  .population-household-chart .chart-subtitle {
    font-size:12px !important;
  }
}


/* === v5.8.38 종합서술 매매·전세·월세 배경색 구분 === */
.summary-note {
  border-width: 1.5px !important;
  border-style: solid !important;
  font-weight: 700 !important;
}
.summary-sale {
  background: #fff4e6 !important;
  border-color: #f59e0b !important;
}
.summary-jeonse {
  background: #eef6ff !important;
  border-color: #60a5fa !important;
}
.summary-rent {
  background: #eefbf3 !important;
  border-color: #34d399 !important;
}
@media print {
  .summary-sale {
    background: #fff4e6 !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
  .summary-jeonse {
    background: #eef6ff !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
  .summary-rent {
    background: #eefbf3 !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
}


/* === v5.8.38 검토.pdf 반영: 도넛 상세·증감색·보조표 삭제 === */
.survey-donut-meta {
  display:inline-block;
  margin:4px auto 8px;
  padding:3px 9px;
  border-radius:999px;
  background:#f1f5f9;
  color:#334155;
  font-weight:900;
  font-size:12px;
}
.survey-donut-ratios {
  display:flex;
  justify-content:center;
  gap:8px;
  flex-wrap:wrap;
  margin-top:8px;
  font-size:12px;
  font-weight:900;
}
.survey-donut-ratios .dir-up,
td.dir-up,
th.dir-up,
.pos,
td .pos {
  color:#dc2626 !important;
  font-weight:900 !important;
}
.survey-donut-ratios .dir-flat,
td.dir-flat,
th.dir-flat,
.neutral,
td .neutral {
  color:#64748b !important;
  font-weight:900 !important;
}
.survey-donut-ratios .dir-down,
td.dir-down,
th.dir-down,
.neg,
td .neg {
  color:#2563eb !important;
  font-weight:900 !important;
}
.survey-donut-caption {
  margin-top:6px;
  text-align:center;
  color:#334155;
  font-weight:800;
  font-size:12px;
}
.supply-audit-note {
  background:#f8fafc;
  border:1px solid #e2e8f0;
  border-radius:12px;
  padding:8px 10px;
  color:#334155;
  font-size:12px;
  font-weight:700;
  margin:6px 0 8px;
}
@media print {
  .survey-donut-meta,
  .survey-donut-ratios,
  .survey-donut-caption,
  .supply-audit-note {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
}


/* === v5.8.38 인구 변화 감지 및 범례 잘림 보정 === */
.population-household-chart .chart-wrap {
  height:350px !important;
  padding-bottom:10px !important;
}
.population-household-chart .chart-note {
  text-align:center;
  color:#64748b;
  font-weight:800;
  font-size:12px;
}
.migration-combo-chart {
  padding-bottom:22px !important;
  min-height:390px !important;
}
.migration-combo-chart .chart-wrap {
  height:330px !important;
  padding-bottom:28px !important;
}
@media print {
  .population-household-chart .chart-wrap,
  .population-household-chart canvas {
    height:275px !important;
    max-height:275px !important;
  }
  .migration-combo-chart {
    min-height:310px !important;
    padding-bottom:12px !important;
  }
  .migration-combo-chart .chart-wrap,
  .migration-combo-chart canvas {
    height:260px !important;
    max-height:260px !important;
  }
}


/* === v5.8.38 전체 차트 범례 잘림 방지 전역 보정 === */
.chart-box,
.mini-chart-card,
.survey-donut-card,
.population-household-chart,
.migration-combo-chart,
.age-net-horizontal-chart {
  overflow: visible !important;
}
.chart-box {
  padding-bottom: 22px !important;
  min-height: 370px;
}
.chart-box .chart-wrap {
  height: 280px;
  min-height: 240px;
  padding-bottom: 34px !important;
  overflow: visible !important;
}
.chart-box canvas {
  display: block !important;
}
.chart-box.mix-chart {
  min-height: 395px;
}
.chart-box.mix-chart .chart-wrap {
  height: 310px;
  padding-bottom: 38px !important;
}
.survey-donut-card {
  padding-bottom: 18px !important;
}
.survey-donut-box {
  height: 290px !important;
  min-height: 290px !important;
  padding-bottom: 24px !important;
  overflow: visible !important;
}
.mini-chart-card {
  padding-bottom: 12px !important;
}
.mini-chart-box {
  padding-bottom: 12px !important;
  overflow: visible !important;
}
@media print {
  .chart-box {
    height: auto !important;
    min-height: 250px !important;
    padding-bottom: 14px !important;
    overflow: visible !important;
  }
  .chart-box .chart-wrap,
  .chart-box canvas {
    height: 205px !important;
    max-height: 205px !important;
    padding-bottom: 20px !important;
  }
  .chart-box.mix-chart {
    min-height: 285px !important;
  }
  .chart-box.mix-chart .chart-wrap,
  .chart-box.mix-chart canvas {
    height: 230px !important;
    max-height: 230px !important;
    padding-bottom: 22px !important;
  }
  .survey-donut-box {
    height: 178px !important;
    min-height: 178px !important;
    padding-bottom: 18px !important;
  }
  .mini-chart-box {
    height: 205px !important;
    padding-bottom: 12px !important;
  }
  .mini-chart-box canvas {
    height: 185px !important;
    max-height: 185px !important;
  }
}


/* === v5.8.38 연령별 차트 3종 색상 통일 및 중복표 삭제 === */
.age-color-chart .chart-wrap {
  height: 285px !important;
  padding-bottom: 34px !important;
}
.age-color-chart {
  min-height: 370px !important;
}
@media print {
  .age-color-chart {
    min-height: 255px !important;
  }
  .age-color-chart .chart-wrap,
  .age-color-chart canvas {
    height: 210px !important;
    max-height: 210px !important;
  }
}


/* === v5.8.38 공급표 머지·컬러 및 비아파트 유형별 건수표 === */
.supply-type-pipeline-table th { background:#1e293b !important; color:#fff !important; }
.supply-type-pipeline-table td { text-align:center !important; vertical-align:middle !important; }
.supply-type-pipeline-table .stage-cell { font-weight:900 !important; color:#0f172a !important; }
.supply-type-pipeline-table tr.stage-permit td { background:#fff7ed !important; }
.supply-type-pipeline-table tr.stage-start td { background:#eff6ff !important; }
.supply-type-pipeline-table tr.stage-complete td { background:#f0fdf4 !important; }
.supply-type-pipeline-table tr.stage-presale td { background:#f5f3ff !important; }
.supply-type-pipeline-table .stage-total { font-weight:900 !important; color:#0f172a !important; }
.supply-type-pipeline-table .stage-cell.stage-permit,
.supply-type-pipeline-table .stage-total.stage-permit { background:#fed7aa !important; }
.supply-type-pipeline-table .stage-cell.stage-start,
.supply-type-pipeline-table .stage-total.stage-start { background:#bfdbfe !important; }
.supply-type-pipeline-table .stage-cell.stage-complete,
.supply-type-pipeline-table .stage-total.stage-complete { background:#bbf7d0 !important; }
.supply-type-pipeline-table .stage-cell.stage-presale,
.supply-type-pipeline-table .stage-total.stage-presale { background:#ddd6fe !important; }
.nonapt-type-count-table th,
.nonapt-type-count-table td { text-align:center !important; white-space:nowrap !important; }
.nonapt-type-count-table th:nth-child(2), .nonapt-type-count-table td:nth-child(2){background:#e0f2fe !important;}
.nonapt-type-count-table th:nth-child(3), .nonapt-type-count-table td:nth-child(3){background:#ede9fe !important;}
.nonapt-type-count-table th:nth-child(4), .nonapt-type-count-table td:nth-child(4){background:#ffedd5 !important;}
.nonapt-type-count-table th:nth-child(5), .nonapt-type-count-table td:nth-child(5){background:#dcfce7 !important;}
.nonapt-type-count-table th:last-child, .nonapt-type-count-table td:last-child{background:#f8fafc !important;font-weight:900 !important;}
@media print {
  .supply-type-pipeline-table td, .supply-type-pipeline-table th,
  .nonapt-type-count-table td, .nonapt-type-count-table th {
    -webkit-print-color-adjust:exact !important;
    print-color-adjust:exact !important;
  }
}


/* === v5.8.38 최종 보정: 비중 강조·비아파트 헤더·공급표/차트 검수 === */
.share-high {
  font-weight: 950 !important;
  color: #7c2d12 !important;
  background: #ffedd5 !important;
}
.nonapt-type-count-table thead th {
  background:#1e293b !important;
  color:#ffffff !important;
}
@media print {
  .share-high,
  .nonapt-type-count-table thead th {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
}


/* === v5.8.38 비아파트 매매·전세·월세 물건유형별 거래건수표 헤더 강제 보정 === */
table.nonapt-header-fixed-table thead,
table.nonapt-header-fixed-table thead tr,
table.nonapt-header-fixed-table thead tr th,
table.nonapt-type-count-table thead,
table.nonapt-type-count-table thead tr,
table.nonapt-type-count-table thead tr th {
  background: #1e293b !important;
  background-color: #1e293b !important;
  color: #ffffff !important;
  font-weight: 900 !important;
  border-color: #1e293b !important;
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
}

/* 본문 색상은 유지하되 헤더 컬럼별 연한 배경색이 덮어쓰지 않도록 thead를 별도 고정 */
table.nonapt-header-fixed-table tbody td:nth-child(2),
table.nonapt-type-count-table tbody td:nth-child(2) {
  background:#e0f2fe !important;
}
table.nonapt-header-fixed-table tbody td:nth-child(3),
table.nonapt-type-count-table tbody td:nth-child(3) {
  background:#ede9fe !important;
}
table.nonapt-header-fixed-table tbody td:nth-child(4),
table.nonapt-type-count-table tbody td:nth-child(4) {
  background:#ffedd5 !important;
}
table.nonapt-header-fixed-table tbody td:nth-child(5),
table.nonapt-type-count-table tbody td:nth-child(5) {
  background:#dcfce7 !important;
}
table.nonapt-header-fixed-table tbody td:last-child,
table.nonapt-type-count-table tbody td:last-child {
  background:#f8fafc !important;
  font-weight:900 !important;
}
@media print {
  table.nonapt-header-fixed-table thead tr th,
  table.nonapt-type-count-table thead tr th {
    background:#1e293b !important;
    color:#ffffff !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
}


/* === v5.8.38 최신행 노란색 표시 기준 정리 ===
   원칙: 13개월 시계열표만 최신행을 노란색 + 굵은 글씨로 표시.
   그 외 분석표(연령 구조, 공급 5개년, 읍면동/면적/가격대/연식 등)는 자동 노란색 강조 제거.
*/
.market-strength-table:not(.time-series-table) .latest-row td,
.compact-wide-table:not(.time-series-table) .latest-row td,
.supply-summary-table .latest-row td {
  background: inherit !important;
  color: inherit !important;
  font-weight: inherit !important;
}

/* 방향성 컬러 셀은 노란색 대신 기존 상승/보합/하락 컬러만 유지 */
.market-strength-table:not(.time-series-table) .latest-row td.dir-up,
.compact-wide-table:not(.time-series-table) .latest-row td.dir-up,
.market-strength-table:not(.time-series-table) .latest-row td.dir-flat,
.compact-wide-table:not(.time-series-table) .latest-row td.dir-flat,
.market-strength-table:not(.time-series-table) .latest-row td.dir-down,
.compact-wide-table:not(.time-series-table) .latest-row td.dir-down {
  font-weight: 900 !important;
}

/* 13개월 시계열표는 최신월 행을 명확히 표시 */
.market-strength-table.time-series-table .latest-row td,
.compact-wide-table.time-series-table .latest-row td,
.time-series-table .latest-row td {
  background: #fff8bf !important;
  font-weight: 900 !important;
}

/* 13개월 시계열표에서 상승/하락 계열 셀은 배경보다 글자색을 우선 */
.time-series-table .latest-row td.pos,
.time-series-table .latest-row td .pos {
  color:#dc2626 !important;
  font-weight:900 !important;
}
.time-series-table .latest-row td.neg,
.time-series-table .latest-row td .neg {
  color:#2563eb !important;
  font-weight:900 !important;
}
.time-series-table .latest-row td.neutral,
.time-series-table .latest-row td .neutral {
  color:#64748b !important;
  font-weight:900 !important;
}

/* 연령 구조표처럼 마지막 행이 '70대 이상'인 단면표는 최신월 행이 아니므로 강조 제거 */
.age-structure-table .latest-row td,
.age-population-structure-table .latest-row td {
  background: inherit !important;
  color: inherit !important;
  font-weight: inherit !important;
}

/* v1.71 readability: 분석표 본문 가독성 강화 및 단지/건물명 강조 */
.market-strength-table th,
.market-strength-table td,
.compact-wide-table th,
.compact-wide-table td {
  font-size: 12px !important;
  line-height: 1.42 !important;
  padding: 7px 7px !important;
}
.time-series-table th,
.time-series-table td {
  font-size: 11.5px !important;
}
.market-strength-table td,
.compact-wide-table td,
.market-strength-table th,
.compact-wide-table th,
.market-strength-table td.name-cell .name-cell-text,
.compact-wide-table td.name-cell .name-cell-text {
  white-space: normal !important;
  word-break: keep-all;
  overflow-wrap: anywhere;
}
.market-strength-table td .renew-ratio,
.compact-wide-table td .renew-ratio { color:#64748b; font-size:11px; font-weight:700; white-space:nowrap; }
/* 상승·하락 비율 좌우 막대(가운데 기준선, 하락 ◀ ▶ 상승) */
td.ratio-bar-cell { padding:4px 8px !important; vertical-align:middle; }
th.ratio-bar-th, td.ratio-bar-cell { width:138px; white-space:nowrap; }
.div-bar-wrap { display:flex; align-items:center; gap:6px; }
.div-bar { flex:1; display:flex; height:16px; min-width:80px; }
.div-bar .db-side { flex:1; display:flex; height:100%; background:#f1f5f9; }
.div-bar .db-left { justify-content:flex-end; border-radius:999px 0 0 999px; border-right:2px solid #334155; }
.div-bar .db-right { justify-content:flex-start; border-radius:0 999px 999px 0; }
.div-bar .db-down { display:block; height:100%; background:#3b82f6; border-radius:999px 0 0 999px; }
.div-bar .db-up { display:block; height:100%; background:#ef4444; border-radius:0 999px 999px 0; }
.db-net { flex:none; min-width:40px; font-size:10.5px; font-weight:800; text-align:left; }
.db-net.up { color:#dc2626; }
.db-net.down { color:#2563eb; }
.db-net.flat { color:#64748b; }
@media print { .div-bar { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }
.market-strength-table td.count-em,
.compact-wide-table td.count-em {
  font-weight: 900 !important;
  font-size: 13.5px !important;
  color: #9a3412 !important;
}
.market-strength-table th.count-em,
.compact-wide-table th.count-em {
  box-shadow: inset 0 -3px 0 #f59e0b !important;
}
.market-strength-table td.name-cell,
.compact-wide-table td.name-cell {
  font-weight: 900 !important;
  color: #0f172a !important;
  text-align: left !important;
  word-break: keep-all !important;
  overflow-wrap: anywhere !important;
  min-width: 118px !important;
}
.market-strength-table:not(.time-series-table) .latest-row td.name-cell,
.compact-wide-table:not(.time-series-table) .latest-row td.name-cell,
.market-strength-table td.name-cell .name-cell-text,
.compact-wide-table td.name-cell .name-cell-text {
  font-weight: 900 !important;
  color: #0f172a !important;
}

@media print {
  .market-strength-table.time-series-table .latest-row td,
  .compact-wide-table.time-series-table .latest-row td,
  .time-series-table .latest-row td {
    background:#fff8bf !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
}


/* === v5.8.38 Ⅵ장 매매·전세·월세 3개 박스 분리 및 전구 아이콘 삭제 === */
.summary-section-box {
  border-radius: 16px;
  padding: 16px 18px;
  margin: 14px 0 18px;
  border: 1.5px solid #e2e8f0;
  page-break-inside: avoid;
}
.summary-section-box h3 {
  margin-top: 0 !important;
  font-size: 18px !important;
  font-weight: 950 !important;
}
.summary-section-box h4 {
  margin: 13px 0 6px !important;
  font-size: 14.5px !important;
  font-weight: 900 !important;
}
.summary-text {
  background: rgba(255,255,255,.68);
  border: 1px solid rgba(148,163,184,.35);
  border-radius: 12px;
  padding: 10px 12px;
  line-height: 1.82;
  font-weight: 750;
  color: #0f172a;
  word-break: keep-all;
}
@media screen {
  .chart-box.has-policy-lines:not(#__policy_guard) { height:auto !important; min-height:0 !important; max-height:none !important; overflow:visible !important; }
  .chart-box.has-policy-lines:not(#__policy_guard) .chart-wrap { height:350px !important; max-height:350px !important; }
}
.unsold-news { margin:8px 0 14px; padding:9px 13px; border:1px solid #e2e8f0; border-left:4px solid #f59e0b; border-radius:10px; background:#fffdf7; font-size:12.5px; }
.unsold-news ul { margin:5px 0 0; padding-left:18px; line-height:1.75; }
.unsold-news .un-summary { margin:6px 0 2px; padding:6px 9px; border-radius:7px; background:#fef3c7; color:#422006; font-weight:700; line-height:1.6; }
.unsold-news .un-summary b { color:#92400e; margin-right:4px; }
.unsold-news small, .unsold-news .un-src, .unsold-news .un-date { color:#64748b; }
.unsold-news .un-date { font-variant-numeric:tabular-nums; margin-right:4px; }
.unsold-news a { color:#0f172a; text-decoration:none; font-weight:700; }
.unsold-news .un-scope { font-style:normal; font-size:10.5px; color:#92400e; background:#fef3c7; border-radius:4px; padding:0 5px; }
tr.move-in-scope-row:not(#__mi) td { font-weight:900 !important; color:#b45309 !important; background:#fff7ed !important; }
tr.move-in-total-row:not(#__mi) td { font-weight:900 !important; color:#0f172a !important; background:#e2e8f0 !important; border-top:2px solid #94a3b8 !important; }
td.trade-mom-cell { font-weight:800; white-space:nowrap; }
td.trade-mom-cell .pos { color:#dc2626; }
td.trade-mom-cell .neg { color:#2563eb; }
td.trade-mom-cell .pending-label { display:inline-block; white-space:nowrap; padding:0 6px; border-radius:6px; background:#fef3c7; color:#b45309; font-size:11px; font-weight:800; }
.market-head-line { margin:2px 0 6px; padding:6px 10px; border-radius:9px; background:rgba(15,23,42,.06); color:#0f172a; font-size:12.5px; font-weight:800; line-height:1.6; word-break:keep-all; }
.summary-text .region-tag { display:inline-block; margin-top:6px; padding:1px 8px; border-radius:999px; font-size:12.5px; font-weight:900; }
.summary-text .region-up { background:#fee2e2; color:#b91c1c; }
.summary-text .region-down { background:#dbeafe; color:#1d4ed8; }
.summary-text .region-mixed { background:#e2e8f0; color:#334155; }
.summary-sale-box {
  background: #fff4e6 !important;
  border-color: #f59e0b !important;
}
.summary-jeonse-box {
  background: #eef6ff !important;
  border-color: #60a5fa !important;
}
.summary-rent-box {
  background: #eefbf3 !important;
  border-color: #34d399 !important;
}
.summary-sale-box h3,
.summary-sale-box h4 { color: #9a3412 !important; }
.summary-jeonse-box h3,
.summary-jeonse-box h4 { color: #1d4ed8 !important; }
.summary-rent-box h3,
.summary-rent-box h4 { color: #047857 !important; }
.top-case-item .case-body span {
  line-height: 1.55 !important;
}
@media print {
  .summary-section-box,
  .summary-text {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
  .summary-section-box {
    padding: 12px 14px !important;
    margin: 10px 0 14px !important;
  }
}


/* === v5.8.38 공급 차트 양식 및 Ⅵ장 아파트 주요 거래사례 보강 === */
.supply-chart-heading {
  margin: 18px 0 8px !important;
  font-size: 16px !important;
  font-weight: 950 !important;
  color: #0f172a !important;
}
.supply-stage-total-grid,
.supply-type-grid {
  margin-bottom: 14px !important;
}
.mini-chart-grid.single-left {
  grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
}
.mini-chart-grid.single-left .mini-chart-card:first-child {
  grid-column: 1 / 2;
}
.summary-major-case {
  margin: 10px 0 14px 0 !important;
  padding: 10px 12px !important;
  background: rgba(255,255,255,.82) !important;
  border: 1.2px dashed #64748b !important;
  border-radius: 10px !important;
  font-weight: 900 !important;
  color: #0f172a !important;
}
.summary-major-case .case-상승,
.summary-major-case .case-top {
  color: #dc2626 !important;
  font-weight: 950 !important;
}
.summary-major-case .case-하락 {
  color: #2563eb !important;
  font-weight: 950 !important;
}
.summary-major-case .case-상위 {
  color: #7c2d12 !important;
  font-weight: 950 !important;
}
@media print {
  .supply-chart-heading {
    font-size: 12px !important;
    margin: 12px 0 6px !important;
  }
  .summary-major-case {
    font-size: 10.5px !important;
    padding: 7px 9px !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
}


/* === v5.8.38 공급 파이프라인 분양 항목 삭제 === */
.supply-stage-total-grid.single-left {
  grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
}
.supply-stage-total-grid.single-left .mini-chart-card:first-child {
  grid-column: 1 / 2;
}


/* === v5.8.38 비아파트 13개월 거래추이 헤더 가독성 최종 보정 및 v52 보강표 삭제 === */
/* 기존 nth-child 헤더 배경 규칙보다 뒤에서 최종 덮어쓰기 */
table.nonapt-type-count-table thead th,
table.nonapt-header-fixed-table thead th,
.nonapt-type-count-table.nonapt-header-fixed-table thead th,
.market-strength-table.nonapt-type-count-table thead tr th {
  background: #1e293b !important;
  background-color: #1e293b !important;
  color: #ffffff !important;
  font-weight: 950 !important;
  text-shadow: none !important;
  border-color: #1e293b !important;
}

/* 연한 컬럼 배경은 tbody에만 적용 */
table.nonapt-type-count-table tbody td:nth-child(2),
table.nonapt-header-fixed-table tbody td:nth-child(2) {
  background:#e0f2fe !important;
  color:#0f172a !important;
}
table.nonapt-type-count-table tbody td:nth-child(3),
table.nonapt-header-fixed-table tbody td:nth-child(3) {
  background:#ede9fe !important;
  color:#0f172a !important;
}
table.nonapt-type-count-table tbody td:nth-child(4),
table.nonapt-header-fixed-table tbody td:nth-child(4) {
  background:#ffedd5 !important;
  color:#0f172a !important;
}
table.nonapt-type-count-table tbody td:nth-child(5),
table.nonapt-header-fixed-table tbody td:nth-child(5) {
  background:#dcfce7 !important;
  color:#0f172a !important;
}
table.nonapt-type-count-table tbody td:last-child,
table.nonapt-header-fixed-table tbody td:last-child {
  background:#f8fafc !important;
  color:#0f172a !important;
  font-weight:900 !important;
}

@media print {
  table.nonapt-type-count-table thead th,
  table.nonapt-header-fixed-table thead th {
    background:#1e293b !important;
    color:#ffffff !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
}


/* === v5.8.38 TOP10 상위거래 방향색·종합판단 컬러·공급차트 표색상 일치 === */
/* 거래금액/보증금/환산월세 상위 TOP10도 직전거래 대비 상승=빨강, 하락=파랑 */
.top-case-item.case-top.case-rise .case-value span,
.top-case-item.case-rise .case-value span {
  color:#dc2626 !important;
  font-weight:950 !important;
}
.top-case-item.case-top.case-fall .case-value span,
.top-case-item.case-fall .case-value span {
  color:#2563eb !important;
  font-weight:950 !important;
}
.top-case-item.case-top.case-zero .case-value span,
.top-case-item.case-zero .case-value span {
  color:#64748b !important;
  font-weight:900 !important;
}

/* Ⅵ장 종합 판단 컬러 박스 */
.comprehensive-summary-box {
  background:#f8fafc !important;
  border:1.5px solid #94a3b8 !important;
  border-left:8px solid #334155 !important;
  border-radius:16px !important;
  padding:14px 16px !important;
  line-height:1.85 !important;
  font-weight:800 !important;
  color:#0f172a !important;
  -webkit-print-color-adjust:exact !important;
  print-color-adjust:exact !important;
}

/* 공급 차트 제목과 색상 기준 안내 */
.supply-chart-heading::after {
  content:"  · 색상 기준: 인허가=주황, 착공=파랑, 준공=초록";
  font-size:11px;
  font-weight:800;
  color:#64748b;
}
@media print {
  .comprehensive-summary-box {
    font-size:11px !important;
    padding:10px 12px !important;
  }
  .supply-chart-heading::after {
    font-size:8.5px !important;
  }
}


/* === v5.8.38 빠른 읽기 기준 자동 하이라이트 로직 === */
.hl-key,
.hl-yellow {
  background:#fff3bf !important;
  color:#7c4a00 !important;
  font-weight:900 !important;
  padding:1px 4px !important;
  border-radius:4px !important;
}
.hl-red {
  background:#ffe3e3 !important;
  color:#c92a2a !important;
  font-weight:900 !important;
  padding:1px 4px !important;
  border-radius:4px !important;
}
.hl-green {
  background:#d3f9d8 !important;
  color:#087f5b !important;
  font-weight:900 !important;
  padding:1px 4px !important;
  border-radius:4px !important;
}
.hl-blue {
  background:#d0ebff !important;
  color:#1864ab !important;
  font-weight:900 !important;
  padding:1px 4px !important;
  border-radius:4px !important;
}
.quick-reading-rule {
  border-left:6px solid #334155 !important;
}
@media print {
  .hl-key, .hl-yellow, .hl-red, .hl-green, .hl-blue {
    -webkit-print-color-adjust:exact !important;
    print-color-adjust:exact !important;
  }
}


/* === v5.8.38 빠른 읽기 기준 문장 단위 제한 강조 === */
.hl-sentence {
  display:inline !important;
  box-decoration-break:clone;
  -webkit-box-decoration-break:clone;
  line-height:1.9;
}
.hl-sentence.hl-key,
.hl-sentence.hl-yellow,
.hl-sentence.hl-red,
.hl-sentence.hl-green,
.hl-sentence.hl-blue {
  padding:2px 4px !important;
  border-radius:4px !important;
  font-weight:900 !important;
}
.note .hl-sentence,
.result-box .hl-sentence,
.summary-text .hl-sentence,
.diagnosis .hl-sentence {
  line-height:1.95 !important;
}
/* 기존 단어 단위처럼 보이는 과도한 배경을 줄이기 위해 문장 강조는 은은하게 조정 */
.hl-sentence.hl-key,
.hl-sentence.hl-yellow {
  background:linear-gradient(transparent 45%, rgba(255, 243, 191, .95) 45%) !important;
  color:#6b3f00 !important;
}
.hl-sentence.hl-red {
  background:linear-gradient(transparent 45%, rgba(255, 227, 227, .95) 45%) !important;
  color:#b42318 !important;
}
.hl-sentence.hl-green {
  background:linear-gradient(transparent 45%, rgba(211, 249, 216, .95) 45%) !important;
  color:#087f5b !important;
}
.hl-sentence.hl-blue {
  background:linear-gradient(transparent 45%, rgba(208, 235, 255, .95) 45%) !important;
  color:#1864ab !important;
}
@media print {
  .hl-sentence {
    -webkit-print-color-adjust:exact !important;
    print-color-adjust:exact !important;
  }
}


/* === v5.8.38 전문 디자이너 관점 가독성 보정 === */
.card, .summary-section-box, .diagnosis, .note, .result-box {
  word-break: keep-all !important;
  overflow-wrap: break-word !important;
}
.summary-text, .diagnosis, .note, .result-box {
  line-height: 1.82 !important;
}
.chapter > h2 {
  letter-spacing: -0.02em !important;
}
.card h3, .card h4 {
  letter-spacing: -0.015em !important;
}
@media print {
  .summary-text, .diagnosis, .note, .result-box {
    line-height: 1.58 !important;
  }
}


/* === v1.46 공급 차트 레이아웃 정리 ===
   - 공급 단계별 총량 차트: 인허가·착공·준공 3개를 1행 3열로 배치
   - 물건유형별 세부 차트: 가독성을 위해 기존 2열 유지
*/
.supply-stage-total-grid.three-col {
  display: grid !important;
  grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
  gap: 12px !important;
  align-items: stretch !important;
  margin-bottom: 18px !important;
}
.supply-stage-total-grid.three-col .mini-chart-card {
  min-width: 0 !important;
  height: 292px !important;
}
.supply-stage-total-grid.three-col .mini-chart-title {
  font-size: 12.5px !important;
  line-height: 1.25 !important;
  min-height: 32px !important;
}
.supply-stage-total-grid.three-col .mini-chart-box {
  height: 225px !important;
}
.supply-stage-total-grid.three-col .mini-chart-box canvas {
  height: 205px !important;
}
.supply-type-grid {
  grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
}
.supply-type-grid.single-left {
  grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
}
.supply-type-grid.single-left .mini-chart-card:first-child {
  grid-column: 1 / 2 !important;
}
@media (max-width: 1100px) {
  .supply-stage-total-grid.three-col {
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
  }
}
@media (max-width: 760px) {
  .supply-stage-total-grid.three-col,
  .supply-type-grid,
  .supply-type-grid.single-left {
    grid-template-columns: 1fr !important;
  }
}
@media print {
  .supply-stage-total-grid.three-col {
    grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
    gap: 8px !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-card {
    height: 240px !important;
    padding: 8px 8px 6px 8px !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-title {
    font-size: 10px !important;
    min-height: 26px !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-box {
    height: 188px !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-box canvas {
    height: 172px !important;
  }
  .supply-type-grid,
  .supply-type-grid.single-left {
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
  }
}


/* === v1.47 공급 관련 차트 3열 레이아웃 통일 ===
   - 공급 단계별 총량 차트: 인허가·착공·준공 1행 3열
   - 물건유형별 세부 차트: 인허가·착공·준공 각각 1행 3개 + 2행 2개
   - 일반 거래·가격 차트는 기존 레이아웃 유지
*/
.supply-stage-total-grid.three-col,
.supply-type-grid,
.supply-type-grid.single-left {
  display: grid !important;
  grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
  gap: 12px !important;
  align-items: stretch !important;
  margin-bottom: 18px !important;
}
.supply-stage-total-grid.three-col .mini-chart-card,
.supply-type-grid .mini-chart-card,
.supply-type-grid.single-left .mini-chart-card {
  min-width: 0 !important;
  height: 292px !important;
}
.supply-stage-total-grid.three-col .mini-chart-title,
.supply-type-grid .mini-chart-title,
.supply-type-grid.single-left .mini-chart-title {
  font-size: 12.5px !important;
  line-height: 1.25 !important;
  min-height: 32px !important;
}
.supply-stage-total-grid.three-col .mini-chart-box,
.supply-type-grid .mini-chart-box,
.supply-type-grid.single-left .mini-chart-box {
  height: 225px !important;
}
.supply-stage-total-grid.three-col .mini-chart-box canvas,
.supply-type-grid .mini-chart-box canvas,
.supply-type-grid.single-left .mini-chart-box canvas {
  height: 205px !important;
}
.supply-type-grid.single-left .mini-chart-card:first-child {
  grid-column: auto !important;
}
@media (max-width: 1200px) {
  .supply-stage-total-grid.three-col,
  .supply-type-grid,
  .supply-type-grid.single-left {
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
  }
}
@media (max-width: 760px) {
  .supply-stage-total-grid.three-col,
  .supply-type-grid,
  .supply-type-grid.single-left {
    grid-template-columns: 1fr !important;
  }
}
@media print {
  .supply-stage-total-grid.three-col,
  .supply-type-grid,
  .supply-type-grid.single-left {
    grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
    gap: 8px !important;
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-card,
  .supply-type-grid .mini-chart-card,
  .supply-type-grid.single-left .mini-chart-card {
    height: 240px !important;
    padding: 8px 8px 6px 8px !important;
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-title,
  .supply-type-grid .mini-chart-title,
  .supply-type-grid.single-left .mini-chart-title {
    font-size: 10px !important;
    line-height: 1.2 !important;
    min-height: 26px !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-box,
  .supply-type-grid .mini-chart-box,
  .supply-type-grid.single-left .mini-chart-box {
    height: 188px !important;
  }
  .supply-stage-total-grid.three-col .mini-chart-box canvas,
  .supply-type-grid .mini-chart-box canvas,
  .supply-type-grid.single-left .mini-chart-box canvas {
    height: 172px !important;
  }
}


/* === v1.48 공급 물건유형별 차트 HTML 구조 보정 ===
   v1.47에서는 CSS는 3열이었으나 HTML이 2+2+1 div로 분리되어 실제 3열이 적용되지 않았음.
   v1.48부터 인허가·착공·준공 물건유형 5개 차트를 하나의 grid 안에 넣어 1행 3개 + 2행 2개로 출력.
*/
.supply-type-grid-3col {
  display: grid !important;
  grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
  gap: 12px !important;
  align-items: stretch !important;
  margin-bottom: 18px !important;
}
.supply-type-grid-3col .mini-chart-card {
  min-width: 0 !important;
  height: 292px !important;
}
.supply-type-grid-3col .mini-chart-title {
  font-size: 12.5px !important;
  line-height: 1.25 !important;
  min-height: 32px !important;
}
.supply-type-grid-3col .mini-chart-box {
  height: 225px !important;
}
.supply-type-grid-3col .mini-chart-box canvas {
  height: 205px !important;
}
@media (max-width: 1200px) {
  .supply-type-grid-3col {
    grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
  }
}
@media (max-width: 760px) {
  .supply-type-grid-3col {
    grid-template-columns: 1fr !important;
  }
}
@media print {
  .supply-type-grid-3col {
    grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
    gap: 8px !important;
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }
  .supply-type-grid-3col .mini-chart-card {
    height: 240px !important;
    padding: 8px 8px 6px 8px !important;
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }
  .supply-type-grid-3col .mini-chart-title {
    font-size: 10px !important;
    line-height: 1.2 !important;
    min-height: 26px !important;
  }
  .supply-type-grid-3col .mini-chart-box {
    height: 188px !important;
  }
  .supply-type-grid-3col .mini-chart-box canvas {
    height: 172px !important;
  }
}


/* 조사기준월 강조 */
.hero .meta .base-month-key { color:#FDE047 !important; }
.hero .meta .base-month-value { color:#FDE047 !important; font-size:22px !important; font-weight:900 !important; line-height:1.1; }
.hero .subtitle .base-month-sub { color:#FDE047; font-weight:900; }
.base-month-inline { color:#DC2626; font-weight:900; }
table.market-strength-table tr.base-month-row td:first-child,
table tr.base-month-row td:first-child { color:#DC2626 !important; font-weight:900 !important; }
table tr.base-month-row td:first-child::after { content:"조사기준월"; display:block; width:max-content; margin:2px auto 0; font-size:10px; line-height:1.4; color:#fff; background:#DC2626; border-radius:6px; padding:0 6px; }
/* 상승·보합·하락 비율 산출 방법 */
.ratio-method-note { background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #2563eb; border-radius:12px; padding:10px 14px; margin:10px 0; color:#1e293b; font-size:13px; line-height:1.6; }
.ratio-method-note > b { color:#1d4ed8; font-size:14px; }
.ratio-method-note ol { margin:6px 0 0 18px; padding:0; }
.ratio-method-note li { margin:2px 0; }
.ratio-method-note .dir-up-text { color:#dc2626; font-weight:800; }
.ratio-method-note .dir-down-text { color:#2563eb; font-weight:800; }
@media print { .ratio-method-note { font-size:9pt; page-break-inside:avoid; } }
</style>
</head>
<body>
<div class="print-guide">PDF 저장 안내: 브라우저 인쇄(Ctrl+P) → A4 → 배경 그래픽 포함으로 저장</div>
<section class="hero">
  <div>
    <div class="eyebrow">MONTHLY HOUSING MARKET REPORT</div>
    <h1>{{ meta.region }} 주택시장 {{ meta.report_title | default("월간동향보고서") }}</h1>
    <div class="subtitle"><span class="base-month-sub">조사기준월 {{ meta.base_month }}</span> {{ meta.report_title | default("월간동향보고서") }} · {{ meta.input_source_label | default("국토교통부 실거래 API 기반") }}, 주민등록, 미분양, 공급, 공표통계, 금리 통합 분석</div>
  </div>
  <div class="meta">
    <b class="base-month-key">조사기준월</b><span class="base-month-value">{{ meta.base_month }}</span>
    {% if meta.is_weekly and meta.weekly_period_label %}<b>주간기준</b><span>{{ meta.weekly_period_label }} · {{ meta.weekly_reference_date }} 월요일 기준</span>{% endif %}
    <b>분석기간</b><span>최근 13개월 / 실거래 최근 3년</span>
    <b>작성기준</b><span>{{ meta.trade_basis }}, 기준월 잠정치</span>
    <b>환산월세</b><span>중위 전월세전환율 {{ conversion_rate.applied_rate | fmt_float }}% 적용</span>
  </div>
</section>
<div class="notice">※ 자료 유의: 통계별 공표 시차에 따라 일부 지표는 최신 공표월 기준으로 작성함. 실거래는 기준월 계약분 잠정치이며 신고기한 미도래로 향후 보정될 수 있음.</div>
<section class="kpis">
  <div><span>수요 기반</span><b>{{ cards.population | fmt_num }}명</b><small>세대 {{ cards.households | fmt_num }} / 인구 전월 {{ cards.population_mom | fmt_num }}</small></div>
  <div><span>{{ meta.primary_asset_label | default("아파트") }} 거래</span><b>매매 {{ cards.apt_trade_count | fmt_num }}건</b><small>전세 {{ cards.apt_jeonse_count | fmt_num }}건{% if not meta.is_weekly %} / 월세 {{ cards.apt_monthly_rent_count | fmt_num }}건{% else %} / 주간 공표지표 반영{% endif %}</small></div>
  <div><span>금리·공급</span><b>주담대 {{ cards.mortgage_rate | fmt_float(2) }}%</b><small>인허가 {{ cards.supply_permit | fmt_num }}호 / 착공 {{ cards.supply_start | fmt_num }}호</small></div>
  <div><span>분양 부담</span><b>미분양 {{ cards.unsold_units | fmt_num }}호</b><small>준공후 {{ cards.after_completion_unsold | fmt_num }}호 / 최신 {{ unsold.latest_month }}</small></div>
</section>
<div class="notice quick-reading-rule">빠른 읽기 기준: <span class="hl-yellow">노란색은 핵심 수치</span>, <span class="hl-red">붉은색은 위험·부담 요인</span>, <span class="hl-green">초록색은 긍정·방어 요인</span>, <span class="hl-blue">파란색은 최종 시장 판단·조사상 유의사항</span>을 의미함. 단, 자동 강조는 LLM 해설 중 핵심 문장 1~2개에만 제한 적용함.</div>
<div class="cover-break"></div>

{% if warnings %}
<div class="notice data-quality-notice"><b>⚠ 데이터 품질 확인</b>{% for w in warnings %}<br>· {{ w }}{% endfor %}</div>
{% endif %}


<section class="chapter bg-indigo monthly-diagnosis-section"><h2>{% if meta.is_weekly %}주간 핵심 진단{% else %}월간 핵심 진단{% endif %}</h2><div class="card diagnosis-card">
<h3>1. <span class="base-month-inline">{{ meta.base_month }}</span> {{ meta.region.split()[-1] }} 주택시장 핵심 판단</h3>
<div class="meaning">※ 지표 의미: {% if meta.is_weekly %}주간 핵심 진단은 주간 공표통계와 인구·금리·미분양·공급·실거래 보조지표를 종합하여 기준주 시장 국면을 점검하기 위한 요약표임.{% else %}월간 핵심 진단은 인구·금리·미분양·공급·실거래 지표를 종합하여 기준월 시장 국면을 한눈에 파악하기 위한 요약표임.{% endif %}</div>
<div class="diagnosis-grid">
  <div class="diagnosis-box"><span>종합 시장 판단</span><b>실거래·수급 복합 점검</b><p>가격 변동률, 거래량, 금리, 공급, 미분양을 함께 판단.</p></div>
  <div class="diagnosis-box"><span>매매 변동률 판단</span><b>{{ apt_trade.signal }}</b><p>비교가능 {{ apt_trade.compare_count | fmt_num }}건 기준 상승·보합·하락 비율 확인.</p></div>
  <div class="diagnosis-box"><span>전세 변동률 판단</span><b>{{ apt_rent.jeonse.signal }}</b><p>갱신계약 제외 신규계약 기준 보증금 흐름 점검.</p></div>
  {% if not meta.is_weekly %}
  <div class="diagnosis-box"><span>월세 변동률 판단</span><b>{{ apt_rent.monthly_rent.signal }}</b><p>환산월세 기준 실제 주거비 부담 점검.</p></div>
  {% else %}
  <div class="diagnosis-box"><span>주간 공표지표 판단</span><b>매매·전세 공표지표 점검</b><p>기준주 매매·전세가격 변동률과 실거래 보조지표를 함께 확인.</p></div>
  {% endif %}
  <div class="diagnosis-box"><span>수요 기반</span><b>{{ population.signal }}</b><p>인구 {{ cards.population | fmt_num }}명, 세대 {{ cards.households | fmt_num }} 기준.</p></div>
  <div class="diagnosis-box"><span>수급·금융 유의요인</span><b>{{ rates.signal }} · {{ unsold.signal }}</b><p>공급 {{ supply.signal }}, 기준월 실거래 잠정치 함께 반영.</p></div>
</div>

<h3>2. 조사 담당자 관점 당월 상승·보합·하락 비율 요약</h3>
<div class="meaning">※ 지표 의미: 도넛 차트는 매매·전세·월세 변동률 조사 판단을 위한 보조지표임. 최종 판단은 TOP 거래사례, 생활권별 흐름, 금리·미분양·공급 지표와 함께 종합함.<br>※ 산출 방법: 같은 단지·같은 전용면적의 직전거래(전세·월세는 직전 신규계약) 대비 변동률이 +0.5% 초과면 상승, −0.5% 미만이면 하락, 그 사이면 보합으로 분류하고, 각 건수를 직전거래가 있는 비교가능건수로 나누어 비율(%)을 산출함. 세부 방법은 Ⅲ~Ⅴ장 각 표 아래 설명 참조.</div>
{{ charts.core_donuts | safe }}

<div class="result-box">💡 핵심 판단: {{ text.monthly_core_diagnosis }}</div>

</div></section>
{% if policy_context.semiconductor_cluster.available %}
{% set cluster = policy_context.semiconductor_cluster %}
<section class="cluster-policy-card">
<h2>정책·개발 이슈 | {{ cluster.title }}</h2>
<p><strong>확인된 사실:</strong> {{ cluster.official_fact }} 지정기간은 {{ cluster.effective_period }}임.</p>
<div class="cluster-levels">
  <span class="semiconductor-cluster-badge core">핵심 영향권: {{ cluster.core_dongs | join('·') }}</span>
  <span class="semiconductor-cluster-badge adjacent">인접 영향권: {{ cluster.adjacent_dongs | join('·') }}</span>
  <span class="semiconductor-cluster-badge linked">연계 영향권: {{ cluster.linked_dongs | join('·') }}</span>
</div>
<p><strong>분류 기준:</strong> {{ cluster.influence_method }}</p>
<p class="policy-caution"><strong>해석 유의:</strong> {{ cluster.scope_note }} {{ cluster.interpretation_rule }}</p>
<p><a href="{{ cluster.source_url }}" target="_blank" rel="noopener">{{ cluster.source_name }}</a></p>
</section>
{% endif %}
<section class="chapter bg-slate"><h2>Ⅰ. {{ meta.region }} 인구·세대 및 이동 동향 (최근 13개월)</h2><div class="card age-pop-card">
<h3>1. 인구, 세대수 및 인구 이동 동향 통합 종합표</h3>
<div class="meaning">※ 지표 의미: 인구수는 기초 주거수요 규모를, 세대수는 가구 분화 및 주택 수요 단위를, 전입·전출·순이동은 실제 유입·유출 흐름을 나타냄.</div>
{{ population_migration_table | safe }}
{{ charts.population | safe }}
{{ charts.migration | safe }}
<div class="note">💡 {{ text.population_commentary }}</div>
<h3>2. 연령대별 주민등록인구 구조 분석</h3>
<div class="meaning">※ 지표 의미: 주민등록인구를 0~9세, 10대, 20대, 30대, 40대, 50대, 60대, 70대 이상으로 재분류하고, 각 월별 인구수와 전월 대비 증감률을 함께 표시함.</div>
<h4>표 1. 최근 13개월 연령대별 주민등록인구 시계열 및 전월 대비 증감률</h4>
{{ age_population_trend_table | safe }}
<h4>표 2. 최신월 연령대별 인구 구조 및 13개월 증감</h4>
{{ age_population.age_rows | as_table }}
{{ charts.age_population_latest | safe }}
{{ charts.age_population_change | safe }}
<div class="age-move-card"><h3>3. 연령대별 전입·전출 및 순이동 분석</h3>
<div class="meaning">※ 지표 의미: 연령대별 이동 흐름을 통해 주거수요의 유입·유출 압력을 점검함.</div>
<h4>표 1. 최근 13개월 연령대별 순이동 시계열</h4>
{{ age_migration_trend_table | safe }}
<div class="note">💡 {{ text.age_migration_commentary }}</div></div>

</div></section>
<section class="chapter bg-slate"><h2>Ⅱ. {{ meta.region_short }} 금리·미분양·주택공급 동향</h2><div class="card">
<h3>1. 예금은행 대출금리 동향: 신규취급액 기준</h3>
<div class="meaning">※ 지표 의미: 주택 매수자, 전세 임차인, 지역 내 자영업·중소기업의 금융비용 부담을 보여주는 핵심 금융 여건 지표임.</div>
{{ rates_standard_table | safe }}
{{ charts.rates | safe }}
<div class="note">💡 {{ text.rate_commentary }}</div>
<h3>2. 미분양 주택 동향(최신 공표월: {{ unsold.latest_month }})</h3>
<div class="meaning">※ 지표 의미: 미분양은 분양시장 부담을, 준공후 미분양은 입주 단계에서도 해소되지 못한 공급 부담을 의미함.</div>
{{ unsold_standard_table | safe }}
{{ charts.unsold | safe }}
<div class="note">💡 {{ text.unsold_commentary }}</div>
{{ unsold_news_html | safe }}
<div class="supply-card"><h3>3. {{ meta.supply_region }} 주택공급 동향 및 {{ meta.region_short }} 미래 공급 압력 판단</h3>
<div class="meaning">※ 지표 의미: 인허가·착공·준공은 향후 입주 부담과 전세 수급 압력을 판단하는 중기 공급 지표임.</div>
<h4>표 1. {{ meta.supply_region }} 최근 5개년 공급 파이프라인 총괄</h4>
{{ supply_summary_table | safe }}
<h4>표 2. {{ meta.supply_region }} 최근 5개년 물건유형별 공급 파이프라인 시계열</h4>
{{ supply_by_type_table | safe }}
{% if text.supply_type_commentary %}<div class="note">💡 {{ text.supply_type_commentary }}</div>{% endif %}
<h4 class="supply-chart-heading">차트 1. {{ meta.supply_region }} 최근 5개년 공급 단계별 총량 시계열</h4>
<div class="mini-chart-grid total-stage-grid supply-stage-total-grid three-col">
{{ charts.supply_permit_total | safe }}
{{ charts.supply_start_total | safe }}
{{ charts.supply_complete_total | safe }}
</div>

{% if not meta.is_officetel and show_move_in_supply %}
<h3>4. {{ meta.region_short }} 아파트 입주예정물량</h3>
<div class="meaning">※ 지표 의미: 입주예정물량은 준공 후 실제 입주가 예상되는 단지 전체 세대수를 기준으로 향후 매매·전세 수급 부담을 점검하는 보조지표임. 일반분양 물량이 아니라 전체 세대수를 우선하며, 입주 일정은 변경될 수 있음.</div>
{{ move_in_supply_panel | safe }}
{% endif %}
<div class="note">💡 {{ text.supply_commentary }}</div>


</div></section>
<section class="chapter bg-rose"><h2>Ⅲ. 매매 주택 거래 현황 분석</h2><div class="card">
<h3>1. {% if meta.is_weekly %}{{ meta.primary_asset_label | default("아파트") }} 매매 실거래 보조지표 및 직전거래 대비 가격 변동 비중{% else %}{{ meta.primary_asset_label | default("아파트") }} 매매 월간 거래량 및 직전거래 대비 가격 변동 비중{% endif %}</h3>
<div class="meaning">※ 지표 의미: 기준월 매매거래와 직전거래 비교 가능 사례의 상승·보합·하락 비율을 함께 점검함.</div>
{{ apt_trade.monthly_rows | as_table }}
<div class="ratio-method-note"><b>📐 상승·보합·하락 비율 산출 방법</b><ol>
<li><b>직전거래 연결</b>: 각 매매거래를 <u>같은 단지(단지명이 없으면 읍면동)·같은 전용면적</u>의 바로 이전 매매거래와 연결함. 최근 3년 실거래 원장을 계약일 순으로 정렬해 찾으며, 단지명 표기가 달라 연결되지 않으면 <u>읍면동+지번+전용면적(정수)</u>으로 한 번 더 찾음.</li>
<li><b>변동률</b> = (이번 거래금액 − 직전 거래금액) ÷ 직전 거래금액 × 100</li>
<li><b>분류 기준</b>: 변동률 <span class="dir-up-text">+0.5% 초과 = 상승</span>, <span class="dir-down-text">−0.5% 미만 = 하락</span>, −0.5%~+0.5% = 보합</li>
<li><b>비율</b>: 상승(보합·하락) 비율 = 해당 건수 ÷ <u>비교가능건수</u> × 100. 직전거래가 없는 거래(최초 거래 등)는 비교가능건수에서 제외하므로 거래건수와 비교가능건수가 다를 수 있음.</li>
<li><b>평균 변동률</b>은 비교가능건 변동률의 단순평균, <b>판정</b>은 상승·보합·하락 비율 중 가장 큰 방향임. 층·동·리모델링 여부는 보정하지 않음.</li>
</ol></div>
{{ charts.apt_monthly | safe }}
{{ charts.apt_strength | safe }}
<div class="note">💡 {{ text.apt_trade_commentary }}</div>
{{ policy_inflection_notes.apt_trade | safe }}
{% if external_buyer.available %}
<div class="card external-buyer-card">
<h3>1-1. 최근 13개월 주거용 매매 및 외지인 거래 추이</h3>
<div class="meaning">※ 외지인은 업로드 매매 원장의 ‘매수자 시군구’가 선택 보고지역의 시군구와 다른 거래로 판정함. 계약일 기준이며 해제 거래와 중복 거래키는 제외함.</div>
{{ external_buyer.monthly_rows | as_table }}
<h4>{{ meta.base_month }} 외지인 매수자 시군구별 건수·비율</h4>
{{ external_current_table | safe }}
<h4>최근 6개월 외지인 매수자 시군구별 누적 및 월별 건수</h4>
{{ external_six_month_table | safe }}
<div class="note">💡 {{ text.external_buyer_commentary }}</div>
</div>
{% endif %}
{% if meta.is_weekly %}
<h3>{% if external_buyer.available %}1-2{% else %}1-1{% endif %}. {{ meta.primary_asset_label | default("아파트") }} 매매 읍면동별 최근 13개월 거래량 추세</h3>
<div class="meaning">※ 증감은 직전월 대비 기준이며, 첫 표시월도 원자료의 직전월과 비교함. 거래가 없는 월은 0건으로 표시하며, 최근 13개월 합계가 많은 읍면동 순으로 정렬함.</div>
{{ apt_trade.dong_monthly_volume_rows | as_dong_volume_trend }}
{% endif %}
<h3>2. {{ meta.primary_asset_label | default("아파트") }} 매매 읍면동별 시장강도 및 세부 현황</h3>{{ charts.apt_dong | safe }}{{ apt_trade.dong_rows | as_table }}
<h3>3. {{ meta.primary_asset_label | default("아파트") }} 매매 거래금액별 세부 현황</h3>{{ apt_trade.price_band_rows | as_table }}{{ charts.apt_price_band | safe }}{% if text.apt_price_band_commentary %}<div class="note">💡 {{ text.apt_price_band_commentary }}</div>{% endif %}
<h3>4. {{ meta.primary_asset_label | default("아파트") }} 매매 연식별 세부 현황</h3>{{ apt_trade.building_age_rows | as_table }}{{ charts.apt_age_band | safe }}{% if text.apt_age_band_commentary %}<div class="note">💡 {{ text.apt_age_band_commentary }}</div>{% endif %}
<h3>5. {{ meta.primary_asset_label | default("아파트") }} 매매 평형별 세부 현황</h3>{{ apt_trade.area_band_rows | as_table }}{{ charts.apt_area_share | safe }}{% if text.apt_area_band_commentary %}<div class="note">💡 {{ text.apt_area_band_commentary }}</div>{% endif %}
<h3>6. {{ meta.primary_asset_label | default("아파트") }} 매매 거래단지 TOP10</h3>
<h4>6-1. 당월({{ meta.base_month }}) 거래단지 TOP10</h4>{{ apt_trade.top_complex_rows | as_table }}
{% if apt_trade.ytd_complex_rows %}<h4>6-2. 당해 누적({{ apt_trade.ytd_period_label }}) 거래단지 TOP10</h4>{{ apt_trade.ytd_complex_rows | as_table }}{% endif %}
<div class="card top-case-card">
<h3>7. {{ meta.primary_asset_label | default("아파트") }} 매매 당월 직전거래 대비 TOP 거래사례</h3>
<p class="meaning">※ 지표 의미: 당월 {{ meta.primary_asset_label | default("아파트") }} 매매 실거래를 동일 단지·동일 전용면적의 직전 매매거래와 비교하여 상승폭·하락폭·거래금액 상위 사례를 확인하는 실거래 사례 지표임. 층 정보는 현재 거래와 직전거래의 가격 차이를 해석하는 보조지표로 함께 표시함.</p>
<div class="top-case-grid">
{{ top_case_list(apt_trade.top_up_rows, "당월 상승거래 TOP10", "매매", "rise") | safe }}
{{ top_case_list(apt_trade.top_down_rows, "당월 하락거래 TOP10", "매매", "fall") | safe }}
{{ top_case_list(apt_trade.top_price_rows, "거래금액 상위 TOP10", "매매", "auto") | safe }}
</div>
</div>
<div class="note">💡 {{ text.top_case_commentary }}</div>

{% if not meta.is_weekly and not meta.is_officetel %}
<h3>8. 비아파트 최근 13개월 매매 거래 추이</h3>
{{ non_apt_sale_type_table | safe }}
{{ charts.non_apt_sale | safe }}
{% if non_apt_sale.rowhouse_dong_rows %}<h4>연립 매매 읍면동별 거래건수 및 상승·보합·하락</h4>
{{ non_apt_sale.rowhouse_dong_rows | as_table }}{% endif %}
{% if non_apt_sale.multiplex_dong_rows %}<h4>다세대 매매 읍면동별 거래건수 및 상승·보합·하락</h4>
{{ non_apt_sale.multiplex_dong_rows | as_table }}{% endif %}
{% if non_apt_sale.rowhouse_major_trade_rows %}<h4>연립 주요 매매 거래내역 TOP10</h4>
{{ non_apt_sale.rowhouse_major_trade_rows | as_table }}{% endif %}
{% if non_apt_sale.multiplex_major_trade_rows %}<h4>다세대 주요 매매 거래내역 TOP10</h4>
{{ non_apt_sale.multiplex_major_trade_rows | as_table }}{% endif %}
{% if non_apt_sale.multifamily_major_trade_rows %}<h4>다가구 주요 매매 거래내역 TOP10</h4>
{{ non_apt_sale.multifamily_major_trade_rows | as_table }}{% endif %}
{% if non_apt_sale.detached_major_trade_rows %}<h4>단독 주요 매매 거래내역 TOP10</h4>
{{ non_apt_sale.detached_major_trade_rows | as_table }}{% endif %}
<div class="note">💡 {{ text.non_apt_sale_commentary }}</div>
{% endif %}
{% if not meta.is_officetel %}
<h3>{% if meta.is_weekly %}8{% else %}9{% endif %}. 아파트 분양권전매 최근 13개월 거래 추이</h3>
{{ presale.monthly_rows | as_table }}
{{ charts.presale | safe }}
{% if presale.presale_major_trade_rows %}<h4>분양권전매 당월 주요 거래내역 TOP10</h4>
{{ presale.presale_major_trade_rows | as_table }}{% endif %}
<div class="note">💡 {{ text.presale_commentary }}</div>
{% endif %}


</div></section>
<section class="chapter bg-emerald"><h2>Ⅳ. 전세 주택 거래 현황 분석</h2><div class="card">
<h3>1. {% if meta.is_weekly %}{{ meta.primary_asset_label | default("아파트") }} 전세 실거래 보조지표 및 직전 신규계약 대비 가격 변동 비중{% else %}{{ meta.primary_asset_label | default("아파트") }} 전세 월간 거래량 및 직전 신규계약 대비 가격 변동 비중{% endif %}</h3>
<div class="meaning">※ 전월세 구분은 원자료의 <strong>{{ apt_rent.rent_type_basis | default("거래구분 열(전세·월세)") }}</strong>을 적용하고, 물건종류는 <strong>{{ apt_rent.property_type_basis | default("주용도 우선·건물유형 보완") }}</strong> 기준으로 확인함. 거래구분 열이 있는 경우 월세액 0 여부로 전세·월세를 재분류하지 않음.</div>
{% if meta.input_source == "excel_upload" %}
<h4>연도별 전월세 원장 로드 점검</h4>
<div class="meaning">※ 최근 3년 전월세 원장을 파일별로 각각 읽은 후 Python에서 결합한 결과임. 세 파일의 행 수와 계약월 범위를 확인하여 누락·중복 지정 여부를 점검함.</div>
{{ apt_rent.input_file_rows | as_table }}
{% endif %}
{{ apt_rent.jeonse.monthly_rows | as_table }}
<div class="ratio-method-note"><b>📐 상승·보합·하락 비율 산출 방법</b><ol>
<li><b>대상</b>: 거래건수는 신규·갱신계약 전체이나, 비율은 <u>신규계약만</u> 사용함(갱신계약은 인상률 상한 등으로 시세 비교에서 제외).</li>
<li><b>직전계약 연결</b>: 각 신규 전세계약을 <u>같은 단지(단지명이 없으면 읍면동)·같은 전용면적</u>의 바로 이전 신규 전세계약과 연결함(없으면 읍면동+지번+전용면적(정수)으로 보완).</li>
<li><b>변동률</b> = (이번 보증금 − 직전 보증금) ÷ 직전 보증금 × 100</li>
<li><b>분류 기준</b>: <span class="dir-up-text">+0.5% 초과 = 상승</span>, <span class="dir-down-text">−0.5% 미만 = 하락</span>, 그 사이 = 보합 / <b>비율</b> = 해당 건수 ÷ 비교가능건수(직전 신규계약이 있는 건) × 100</li>
<li><b>평균 변동률</b>은 비교가능건 변동률의 단순평균, <b>판정</b>은 세 비율 중 가장 큰 방향임.</li>
</ol></div>
{{ charts.apt_jeonse_monthly | safe }}
{{ charts.apt_jeonse_strength | safe }}
<div class="note">💡 {{ text.apt_jeonse_commentary }}</div>
{{ policy_inflection_notes.apt_jeonse | safe }}
{% if meta.is_weekly %}
<h3>1-1. {{ meta.primary_asset_label | default("아파트") }} 전세 읍면동별 최근 13개월 거래량 추세</h3>
<div class="meaning">※ 증감은 직전월 대비 기준이며, 첫 표시월도 원자료의 직전월과 비교함. 신규·갱신계약을 포함한 전체 전세 거래 기준이며, 거래가 없는 월은 0건으로 표시함.</div>
{{ apt_rent.jeonse.dong_monthly_volume_rows | as_dong_volume_trend }}
{% endif %}
<h3>2. {{ meta.primary_asset_label | default("아파트") }} 전세 읍면동별 시장강도 및 세부 현황</h3>{{ charts.apt_jeonse_market_dong | safe }}{{ apt_rent.jeonse.dong_rows | as_table }}
<h3>3. {{ meta.primary_asset_label | default("아파트") }} 전세 보증금 구간별 세부 현황</h3>{{ apt_rent.jeonse.deposit_band_rows | as_table }}{{ charts.apt_jeonse_deposit | safe }}
<h3>4. {{ meta.primary_asset_label | default("아파트") }} 전세 연식별 세부 현황</h3>{{ apt_rent.jeonse.building_age_rows | as_table }}{{ charts.apt_jeonse_age | safe }}
<h3>5. {{ meta.primary_asset_label | default("아파트") }} 전세 평형별 세부 현황</h3>{{ apt_rent.jeonse.area_band_rows | as_table }}{{ charts.apt_jeonse_area | safe }}
<h3>6. {{ meta.primary_asset_label | default("아파트") }} 전세 신규계약 거래단지 TOP10</h3>
<h4>6-1. 당월({{ meta.base_month }}) 거래단지 TOP10</h4>{{ apt_rent.jeonse.top_complex_rows | as_table }}
{% if apt_rent.jeonse.ytd_complex_rows %}<h4>6-2. 당해 누적({{ apt_rent.jeonse.ytd_period_label }}) 거래단지 TOP10</h4>{{ apt_rent.jeonse.ytd_complex_rows | as_table }}{% endif %}
<div class="card top-case-card">
<h3>7. {{ meta.primary_asset_label | default("아파트") }} 전세 당월 직전 신규계약 대비 TOP 거래사례</h3>
<p class="meaning">※ 지표 의미: 당월 {{ meta.primary_asset_label | default("아파트") }} 전세 신규계약을 동일 단지·동일 전용면적의 직전 신규 전세계약과 비교하여 보증금 상승·하락·상위 사례를 확인하는 실거래 사례 지표임. 갱신계약은 가격 비교에서 제외하며, 현재층과 직전층을 함께 표시함.</p>
<div class="top-case-grid">
{{ top_case_list(apt_rent.jeonse.top_up_rows, "보증금 상승 TOP10", "전세", "rise") | safe }}
{{ top_case_list(apt_rent.jeonse.top_down_rows, "보증금 하락 TOP10", "전세", "fall") | safe }}
{{ top_case_list(apt_rent.jeonse.top_value_rows, "보증금 상위 TOP10", "전세", "auto") | safe }}
</div>
</div>

{% if not meta.is_weekly and not meta.is_officetel %}
<h3>8. 비아파트 최근 13개월 전세 거래 추이</h3>
<div class="meaning">※ 비아파트도 원자료의 {{ non_apt_rent.rent_type_basis | default("거래구분 열(전세·월세)") }}과 {{ non_apt_rent.property_type_basis | default("주용도 우선·건물유형 보완") }} 기준으로 전세 및 연립·다세대·단독·다가구를 구분함.</div>
{{ non_apt_jeonse_type_table | safe }}
{{ charts.non_apt_jeonse | safe }}
{% if non_apt_rent.jeonse.rowhouse_dong_rows %}<h4>연립 전세 읍면동별 거래건수 및 상승·보합·하락</h4>
{{ non_apt_rent.jeonse.rowhouse_dong_rows | as_table }}{% endif %}
{% if non_apt_rent.jeonse.multiplex_dong_rows %}<h4>다세대 전세 읍면동별 거래건수 및 상승·보합·하락</h4>
{{ non_apt_rent.jeonse.multiplex_dong_rows | as_table }}{% endif %}
{% if non_apt_rent.jeonse.rowhouse_major_trade_rows %}<h4>연립 주요 전세 거래내역 TOP10</h4>
{{ non_apt_rent.jeonse.rowhouse_major_trade_rows | as_table }}{% endif %}
{% if non_apt_rent.jeonse.multiplex_major_trade_rows %}<h4>다세대 주요 전세 거래내역 TOP10</h4>
{{ non_apt_rent.jeonse.multiplex_major_trade_rows | as_table }}{% endif %}
{% if non_apt_rent.jeonse.multifamily_major_trade_rows %}<h4>다가구 주요 전세 거래내역 TOP10</h4>
{{ non_apt_rent.jeonse.multifamily_major_trade_rows | as_table }}{% endif %}
{% if non_apt_rent.jeonse.detached_major_trade_rows %}<h4>단독 주요 전세 거래내역 TOP10</h4>
{{ non_apt_rent.jeonse.detached_major_trade_rows | as_table }}{% endif %}
<div class="note">💡 {{ text.non_apt_jeonse_commentary }}</div>
{% endif %}


</div></section>
{% if not meta.is_weekly %}
<section class="chapter bg-blue"><h2>Ⅴ. 월세 주택 거래 현황 분석</h2><div class="card">
<div class="conversion-card"><h3>1. 월세 환산 기준 요약표</h3>
{{ conversion_rate_table | safe }}
<div class="note">💡 {{ text.conversion_rate_commentary }}</div></div>
<h3>2. {{ meta.primary_asset_label | default("아파트") }} 월세 월간 거래량 및 직전 신규계약 대비 가격 변동 비중</h3>
<div class="meaning">※ 전월세 구분은 원자료의 <strong>{{ apt_rent.rent_type_basis | default("거래구분 열(전세·월세)") }}</strong>을 적용하고, 물건종류는 <strong>{{ apt_rent.property_type_basis | default("주용도 우선·건물유형 보완") }}</strong> 기준으로 확인함. 월세는 거래구분이 월세인 계약만 집계하고 보증금·월세 조합은 환산월세로 비교함.</div>
{{ apt_rent.monthly_rent.monthly_rows | as_table }}
<div class="ratio-method-note"><b>📐 상승·보합·하락 비율 산출 방법</b><ol>
<li><b>대상</b>: 거래건수는 신규·갱신계약 전체이나, 비율은 <u>신규계약만</u> 사용함.</li>
<li><b>비교 금액</b>: 보증금과 월세가 함께 달라지므로 <u>환산월세 = 월세 + (보증금 × 적용 전월세전환율 ÷ 100 ÷ 12)</u>로 통일해 비교함.</li>
<li><b>직전계약 연결</b>: 같은 단지(단지명이 없으면 읍면동)·같은 전용면적의 바로 이전 신규 월세계약과 연결함(없으면 읍면동+지번+전용면적(정수)으로 보완).</li>
<li><b>변동률</b> = (이번 환산월세 − 직전 환산월세) ÷ 직전 환산월세 × 100 / <span class="dir-up-text">+0.5% 초과 = 상승</span>, <span class="dir-down-text">−0.5% 미만 = 하락</span>, 그 사이 = 보합</li>
<li><b>비율</b> = 해당 건수 ÷ 비교가능건수 × 100, <b>판정</b>은 세 비율 중 가장 큰 방향임.</li>
</ol></div>
{{ charts.apt_rent_monthly | safe }}
{{ charts.apt_rent_strength | safe }}
<div class="note">💡 {{ text.apt_rent_commentary }}</div>
{{ policy_inflection_notes.apt_rent | safe }}
<h3>3. {{ meta.primary_asset_label | default("아파트") }} 월세 읍면동별 시장강도 및 세부 현황</h3>{{ charts.apt_rent_market_dong | safe }}{{ apt_rent.monthly_rent.dong_rows | as_table }}
<h3>4. {{ meta.primary_asset_label | default("아파트") }} 월세 보증금·월세·환산월세 구간별 세부 현황</h3><h4>보증금 구간별</h4>{{ apt_rent.monthly_rent.deposit_band_rows | as_table }}{{ charts.apt_rent_deposit | safe }}<h4>월세 구간별</h4>{{ apt_rent.monthly_rent.monthly_band_rows | as_table }}{{ charts.apt_rent_monthly_band | safe }}
<h3>5. {{ meta.primary_asset_label | default("아파트") }} 월세 연식별 세부 현황</h3>{{ apt_rent.monthly_rent.building_age_rows | as_table }}{{ charts.apt_rent_age | safe }}
<h3>6. {{ meta.primary_asset_label | default("아파트") }} 월세 평형별 세부 현황</h3>{{ apt_rent.monthly_rent.area_band_rows | as_table }}{{ charts.apt_rent_area | safe }}
<h3>7. {{ meta.primary_asset_label | default("아파트") }} 월세 신규계약 거래단지 TOP10</h3>
<h4>7-1. 당월({{ meta.base_month }}) 거래단지 TOP10</h4>{{ apt_rent.monthly_rent.top_complex_rows | as_table }}
{% if apt_rent.monthly_rent.ytd_complex_rows %}<h4>7-2. 당해 누적({{ apt_rent.monthly_rent.ytd_period_label }}) 거래단지 TOP10</h4>{{ apt_rent.monthly_rent.ytd_complex_rows | as_table }}{% endif %}
<div class="card top-case-card">
<h3 id="rent-top10-anchor">8. {{ meta.primary_asset_label | default("아파트") }} 월세 당월 직전 신규계약 대비 TOP 거래사례</h3>
<p class="meaning">※ 지표 의미: 당월 {{ meta.primary_asset_label | default("아파트") }} 월세 신규계약을 동일 단지·동일 전용면적의 직전 신규 월세계약과 비교하여 환산월세 상승·하락·상위 사례를 확인하는 실거래 사례 지표임. 보증금과 월세를 함께 반영한 환산월세 기준이며, 현재층과 직전층을 함께 표시함.</p>
<div class="top-case-grid">
{{ top_case_list(apt_rent.monthly_rent.top_up_rows, "환산월세 상승 TOP10", "월세", "rise") | safe }}
{{ top_case_list(apt_rent.monthly_rent.top_down_rows, "환산월세 하락 TOP10", "월세", "fall") | safe }}
{{ top_case_list(apt_rent.monthly_rent.top_value_rows, "환산월세 상위 TOP10", "월세", "auto") | safe }}
</div>
</div>

{% if not meta.is_weekly and not meta.is_officetel %}
<h3>9. 비아파트 최근 13개월 월세 거래 추이</h3>
<div class="meaning">※ 비아파트도 원자료의 {{ non_apt_rent.rent_type_basis | default("거래구분 열(전세·월세)") }}과 {{ non_apt_rent.property_type_basis | default("주용도 우선·건물유형 보완") }} 기준으로 월세 및 연립·다세대·단독·다가구를 구분함.</div>
{{ non_apt_rent_type_table | safe }}
{{ charts.non_apt_rent | safe }}
{% if non_apt_rent.monthly_rent.rowhouse_dong_rows %}<h4>연립 월세 읍면동별 거래건수 및 상승·보합·하락</h4>
{{ non_apt_rent.monthly_rent.rowhouse_dong_rows | as_table }}{% endif %}
{% if non_apt_rent.monthly_rent.multiplex_dong_rows %}<h4>다세대 월세 읍면동별 거래건수 및 상승·보합·하락</h4>
{{ non_apt_rent.monthly_rent.multiplex_dong_rows | as_table }}{% endif %}
{% if non_apt_rent.monthly_rent.rowhouse_major_trade_rows %}<h4>연립 주요 월세 거래내역 TOP10</h4>
{{ non_apt_rent.monthly_rent.rowhouse_major_trade_rows | as_table }}{% endif %}
{% if non_apt_rent.monthly_rent.multiplex_major_trade_rows %}<h4>다세대 주요 월세 거래내역 TOP10</h4>
{{ non_apt_rent.monthly_rent.multiplex_major_trade_rows | as_table }}{% endif %}
{% if non_apt_rent.monthly_rent.multifamily_major_trade_rows %}<h4>다가구 주요 월세 거래내역 TOP10</h4>
{{ non_apt_rent.monthly_rent.multifamily_major_trade_rows | as_table }}{% endif %}
{% if non_apt_rent.monthly_rent.detached_major_trade_rows %}<h4>단독 주요 월세 거래내역 TOP10</h4>
{{ non_apt_rent.monthly_rent.detached_major_trade_rows | as_table }}{% endif %}
<div class="note">💡 {{ text.non_apt_rent_commentary }}</div>
{% endif %}


</div></section>
{% endif %}
<section id="chapter-final-summary" class="chapter bg-indigo final-summary-chapter"><h2>{% if meta.is_weekly %}Ⅴ{% else %}Ⅵ{% endif %}. 종합 시황 서술 보고서</h2><div class="card">
<div class="notice">※ Ⅵ장에는 표를 추가하지 않고, 앞선 분석표에서 도출된 핵심 해석만 문단에 반영함.</div>
<div class="summary-section-box summary-sale-box">
<h3>1. 매매시장 종합</h3>
<h4>1) 매매 종합</h4><div class="summary-text">{{ text.sale_market_summary }}</div>
<h4>2) {{ meta.primary_asset_label | default("아파트") }} 매매</h4>{{ market_head_line(apt_trade, "매매") | safe }}<div class="summary-text">{{ text.apt_sale_summary }}</div>{{ major_case_line(apt_trade, "매매") | safe }}
{% if not meta.is_weekly and not meta.is_officetel %}
<h4>3) 연립·다세대 매매</h4><div class="summary-text">{{ text.rowhouse_multiplex_sale_summary }}</div>
<h4>4) 단독·다가구 매매</h4><div class="summary-text">{{ text.detached_multifamily_sale_summary }}</div>
{% endif %}
</div>
<div class="summary-section-box summary-jeonse-box">
<h3>2. 전세시장 종합</h3>
<h4>5) 전세 종합</h4><div class="summary-text">{{ text.jeonse_market_summary }}</div>
<h4>6) {{ meta.primary_asset_label | default("아파트") }} 전세</h4>{{ market_head_line(apt_rent.jeonse, "전세") | safe }}<div class="summary-text">{{ text.apt_jeonse_summary }}</div>{{ major_case_line(apt_rent.jeonse, "전세") | safe }}
{% if not meta.is_weekly and not meta.is_officetel %}
<h4>7) 연립·다세대 전세</h4><div class="summary-text">{{ text.rowhouse_multiplex_jeonse_summary }}</div>
<h4>8) 단독·다가구 전세</h4><div class="summary-text">{{ text.detached_multifamily_jeonse_summary }}</div>
{% endif %}
</div>
{% if not meta.is_weekly %}
<div class="summary-section-box summary-rent-box">
<h3>3. 월세시장 종합</h3>
<h4>9) 월세 종합</h4><div class="summary-text">{{ text.rent_market_summary }}</div>
<h4>10) {{ meta.primary_asset_label | default("아파트") }} 월세</h4>{{ market_head_line(apt_rent.monthly_rent, "월세") | safe }}<div class="summary-text">{{ text.apt_rent_summary }}</div>{{ major_case_line(apt_rent.monthly_rent, "월세") | safe }}
{% if not meta.is_weekly and not meta.is_officetel %}
<h4>11) 연립·다세대 월세</h4><div class="summary-text">{{ text.rowhouse_multiplex_rent_summary }}</div>
<h4>12) 단독·다가구 월세</h4><div class="summary-text">{{ text.detached_multifamily_rent_summary }}</div>
{% endif %}
</div>
{% endif %}
<h3>{% if meta.is_weekly %}3{% else %}4{% endif %}. 종합 판단</h3>
<div class="diagnosis comprehensive-summary-box">{{ text.comprehensive_summary }}</div>

</div></section>
<section class="chapter bg-slate appendix-section"><h2>부록. 통계별 기준월 및 공표월 관리표</h2><div class="card">
<div class="meaning">※ 각 통계의 적용 기준월과 최신 공표월을 확인하는 관리용 표임.</div>
{{ period_table | safe }}
<div class="footer">자동 생성 시각: {{ meta.created_at }} · {{ meta.app_version }}</div>
</div></section>
</body>
</html>
"""



# ---------------------------------------------------------------------
# template_v58 strict preservation support
# ---------------------------------------------------------------------
TEMPLATE_V58_REQUIRED_CANVAS_IDS = [
    "surveySaleDonut",
    "surveyJeonseDonut",
    "surveyRentDonut",
    "chart1",
    "chart2",
    "chartAgePopLatest",
    "chartAgePopChange",
    "chartAgeNetYoung",
                "chart3",
    "chart4",
    "totalPermit",
    "totalStart",
    "totalComplete",
    "totalSale",
    "permitChart_아파트",
    "permitChart_연립",
    "permitChart_다세대",
    "permitChart_단독",
    "permitChart_다가구",
    "startChart_아파트",
    "startChart_연립",
    "startChart_다세대",
    "startChart_단독",
    "startChart_다가구",
    "completeChart_아파트",
    "completeChart_연립",
    "completeChart_다세대",
    "completeChart_단독",
    "completeChart_다가구",
    "chart7",
    "chart33",
    "chart11",
    "chart12",
    "chart30",
    "chart13",
    "chart14",
    "chart15",
    "chart16",
    "chart8",
    "chart34",
    "chart18",
    "chart31",
    "chart19",
    "chart25",
    "chart26",
    "chart20",
    "chart9",
    "chart35",
    "chart22",
    "chart32",
    "chart23",
    "chart27",
    "chart28",
    "chart29",
    "chart24"
]

def _safe_empty_chart_script(canvas_id: str) -> str:
    cfg = {
        "type": "bar",
        "data": {"labels": ["구조보존"], "datasets": [{"label": "template_v58 구조보존", "data": [0], "borderWidth": 1}]},
        "options": {"responsive": False, "animation": False, "plugins": {"legend": {"display": False}}, "scales": {"y": {"beginAtZero": True}}}
    }
    return f"<canvas id='{canvas_id}' width='1' height='1'></canvas><script>registerChart('{canvas_id}', {json.dumps(cfg, ensure_ascii=False)});</script>"

def ensure_template_v58_structure(html: str) -> str:
    """template_v58 기준에 없는 숨김 차트/가짜 클래스 추가를 하지 않고,
    실제 출력 구조만 검수·정리한다."""
    meta_tags = [
        '<meta name="standard-version" content="v58"/>',
        '<meta name="template-preservation-version" content="v58-only-visible"/>',
        '<meta name="template-preservation-canvas-count" content="58"/>',
    ]
    for tag in meta_tags:
        name_match = re.search(r'name="([^"]+)"', tag)
        if name_match and name_match.group(1) not in html:
            html = html.replace("</head>", tag + "\n</head>")
    html = html.replace("<body>", "<body class='template-v58-generated'>", 1) if "<body>" in html and "template-v58-generated" not in html else html
    return html

def ensure_template_v58_class_bank(html: str) -> str:
    """v5.8.38부터 템플릿에 없는 숨김 클래스/차트 bank를 생성하지 않음."""
    return html


def age_population_trend_table_html(age_population: Dict[str, Any]) -> str:
    """최근 13개월 연령대별 주민등록인구 시계열 및 전월 대비 증감률 표."""
    rows = age_population.get("monthly_rows_with_prev", []) if isinstance(age_population, dict) else []
    if not rows and isinstance(age_population, dict):
        rows = age_population.get("monthly_rows", [])
    if not rows:
        return "<p class='muted'>연령대별 주민등록인구 원자료 없음</p>"
    age_cols = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
    rows13 = rows[-13:]
    prev = rows[-14] if len(rows) > 13 else None
    has_hidden_prev = isinstance(prev, dict)
    parts = []
    parts.append("<div class='table-wrap age-pop-trend-table-wrap'>")
    parts.append("<table class='market-strength-table compact-wide-table time-series-table age-pop-trend-table'>")
    parts.append("<thead><tr><th>기준월</th>" + "".join(f"<th>{c}</th>" for c in age_cols) + "</tr></thead><tbody>")
    for idx, r in enumerate(rows13):
        tr_cls = " class='latest-row'" if idx == len(rows13)-1 else ""
        parts.append(f"<tr{tr_cls}><td>{html_value(r.get('month') or r.get('기준월'))}</td>")
        for c in age_cols:
            v = r.get(c, None)
            try:
                fv = float(v)
            except Exception:
                fv = None
            pct = None
            if prev is not None and fv is not None:
                try:
                    pv = float(prev.get(c))
                    if pv:
                        pct = (fv - pv) / pv * 100
                except Exception:
                    pct = None
            if fv is None:
                cell = "-"
                cls = ""
            else:
                num_s = f"{int(round(fv)):,}"
                if pct is None:
                    missing_note = "직전월 없음" if not has_hidden_prev and idx == 0 else "-"
                    cell = f"{num_s} <span class='age-rate neutral'>({missing_note})</span>"
                    cls = ""
                else:
                    sign_cls = "pos" if pct > 0 else ("neg" if pct < 0 else "neutral")
                    cell = f"{num_s} <span class='age-rate {sign_cls}'>({pct:.1f}%)</span>"
                    cls = " class='age-up'" if pct > 0 else (" class='age-down'" if pct < 0 else " class='age-flat'")
            parts.append(f"<td{cls}>{cell}</td>")
        parts.append("</tr>")
        prev = r
    parts.append("</tbody></table></div>")
    return "".join(parts)


def _fmt_signed_num(v: Any) -> str:
    try:
        if v is None or pd.isna(v):
            return "-"
    except Exception:
        if v is None:
            return "-"
    try:
        iv = int(round(float(v)))
        if iv > 0:
            return f"<span class='pos'>{iv:,}</span>"
        if iv < 0:
            return f"<span class='neg'>{iv:,}</span>"
        return f"<span class='neutral'>0</span>"
    except Exception:
        return html_value(v)


def age_migration_trend_table_html(migration: Dict[str, Any]) -> str:
    """최근 13개월 연령대별 순이동 시계열 표."""
    rows = migration.get("age_trend_rows", []) if isinstance(migration, dict) else []
    if not rows:
        # age_monthly_rows만 있는 경우 pivot 보정
        raw = migration.get("age_monthly_rows", []) if isinstance(migration, dict) else []
        if raw:
            try:
                df = pd.DataFrame(raw)
                age_col = "age" if "age" in df.columns else "age_group"
                piv = df.pivot_table(index="month", columns=age_col, values="net", aggfunc="sum", fill_value=0).reset_index().sort_values("month")
                order = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
                for c in order:
                    if c not in piv.columns:
                        piv[c] = 0
                rows = piv[["month"] + order].tail(13).to_dict("records")
            except Exception:
                rows = []
    if not rows:
        return "<p class='muted'>최근 13개월 연령대별 순이동 원자료 없음</p>"

    age_cols = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
    html = []
    html.append("<div class='table-wrap age-migration-trend-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table time-series-table age-migration-trend-table'>")
    html.append("<thead><tr><th>기준월</th>" + "".join(f"<th>{c}</th>" for c in age_cols) + "</tr></thead><tbody>")
    rows13 = rows[-13:]
    for idx, r in enumerate(rows13):
        tr_cls = " class='latest-row'" if idx == len(rows13) - 1 else ""
        html.append(f"<tr{tr_cls}><td>{html_value(r.get('month') or r.get('기준월'))}</td>")
        for c in age_cols:
            v = r.get(c, 0)
            try:
                fv = int(round(float(v)))
            except Exception:
                fv = None
            if fv is None:
                cell = "-"
                cls = ""
            elif fv > 0:
                cell = f"<span class='pos'>{fv:,}</span>"
                cls = " class='age-net-pos'"
            elif fv < 0:
                cell = f"<span class='neg'>{fv:,}</span>"
                cls = " class='age-net-neg'"
            else:
                cell = "<span class='neutral'>0</span>"
                cls = " class='age-net-zero'"
            html.append(f"<td{cls}>{cell}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)

def population_migration_table_html(data: Dict[str, Any]) -> str:
    """인구·세대수 및 인구이동 통합표.
    인구/세대수는 기준월까지 있으나 이동자수는 공표월이 늦을 수 있으므로
    미공표 월을 누락으로 보이지 않게 '미공표'로 명시한다.
    주요 이동 연령대는 메인 표에서 분리해 가로 잘림을 방지한다.
    """
    population_data = data.get("population", {}) if isinstance(data, dict) else {}
    pop_rows = population_data.get("table_rows_with_prev", []) or population_data.get("table_rows", []) or []
    # 인구표의 첫 표시월(최신 이동자수 공표월보다 이른 달)도 매칭되도록 수집된 전체 월을 우선 사용한다.
    mig_rows = data.get("migration", {}).get("all_rows", []) or data.get("migration", {}).get("table_rows", []) or []
    mig_by_month = {str(r.get("month") or r.get("기준월")): r for r in mig_rows}
    mig_latest = data.get("migration", {}).get("latest_month") or "-"
    rows = []
    prev_pop = None
    prev_hh = None
    display_pop_rows = pop_rows[-13:]
    hidden_prev = pop_rows[-14] if len(pop_rows) > 13 else None
    if isinstance(hidden_prev, dict):
        prev_pop = pick_val(hidden_prev, ["population", "주민등록인구", "총인구수", "인구"], None)
        prev_hh = pick_val(hidden_prev, ["households", "세대수", "총세대수"], None)
    for r in display_pop_rows:
        m = str(r.get("month") or r.get("기준월") or "")
        pop = pick_val(r, ["population", "주민등록인구", "총인구수", "인구"], None)
        hh = pick_val(r, ["households", "세대수", "총세대수"], None)
        pop_diff = pick_val(r, ["population_mom", "전월 대비 증감", "인구증감"], None)
        hh_diff = pick_val(r, ["households_mom", "세대 전월 대비 증감", "세대증감"], None)
        if pop_diff in [None, ""] and prev_pop not in [None, ""]:
            try: pop_diff = float(pop) - float(prev_pop)
            except Exception: pop_diff = None
        if hh_diff in [None, ""] and prev_hh not in [None, ""]:
            try: hh_diff = float(hh) - float(prev_hh)
            except Exception: hh_diff = None
        mr = mig_by_month.get(m)
        if mr:
            inv = pick_val(mr, ["in", "전입", "총전입", "move_in"], None)
            outv = pick_val(mr, ["out", "전출", "총전출", "move_out"], None)
            netv = pick_val(mr, ["net", "순이동", "net_move"], None)
        else:
            inv = outv = netv = "미공표"
        rows.append({
            "기준월": m,
            "총인구수": html_value(pop),
            "전월 대비 증감": _fmt_signed_num(pop_diff),
            "총세대수": html_value(hh),
            "세대 전월 대비 증감": _fmt_signed_num(hh_diff),
            "총 전입": html_value(inv) if inv != "미공표" else "<span class='not-published'>미공표</span>",
            "총 전출": html_value(outv) if outv != "미공표" else "<span class='not-published'>미공표</span>",
            "순이동": _fmt_signed_num(netv) if netv != "미공표" else "<span class='not-published'>미공표</span>",
        })
        prev_pop, prev_hh = pop, hh

    html = []
    html.append("<div class='population-publish-note'>※ 인구·세대수는 기준월 자료를 표시하고, 전입·전출·순이동은 인구이동 통계 최신 공표월 기준으로 표시함. 이동자수 최신 공표월: <b>" + html_value(mig_latest) + "</b></div>")
    html.append("<div class='table-wrap population-migration-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table time-series-table population-migration-table'>")
    if rows:
        cols = list(rows[0].keys())
        html.append("<thead><tr>" + "".join(f"<th>{display_col_name(c)}</th>" for c in cols) + "</tr></thead><tbody>")
        for idx, r in enumerate(rows):
            tr_cls = " class='latest-row'" if idx == len(rows)-1 else ""
            html.append(f"<tr{tr_cls}>" + "".join(f"<td>{r.get(c, '-')}</td>" for c in cols) + "</tr>")
        html.append("</tbody>")
    html.append("</table></div>")

    return "".join(html)


def _supply_year_value(row: Dict[str, Any]) -> str:
    m = str(row.get("month") or row.get("기준연도") or row.get("기준월") or "")
    # 202112, 2021-12, 2021.12 모두 2021로 정규화
    mm = re.search(r"(20\d{2})", m)
    return mm.group(1) if mm else m

def _supply_period_label(row: Dict[str, Any]) -> str:
    """공급 표·차트 공통 라벨: 과거 연말은 연도, 당해 누계는 최종 공표월 표시."""
    explicit = str(row.get("period_label") or "").strip()
    if explicit:
        return explicit
    m = str(row.get("month") or row.get("기준월") or row.get("기준연도") or "")
    nums = re.sub(r"[^0-9]", "", m)
    if len(nums) >= 6:
        y, mm = nums[:4], int(nums[4:6])
        return y if mm == 12 else f"{y}년 {mm}월 누계"
    return _supply_year_value(row)

def _fmt_pct(v: Any) -> str:
    try:
        if v is None or pd.isna(v):
            return "-"
    except Exception:
        if v is None:
            return "-"
    try:
        return f"{float(v):.1f}%"
    except Exception:
        return str(v)



def _fmt_rate(v: Any) -> str:
    try:
        if v is None or pd.isna(v):
            return "-"
    except Exception:
        if v is None:
            return "-"
    try:
        return f"{float(v):.2f}"
    except Exception:
        return str(v)

def _fmt_rate_diff(v: Any) -> str:
    try:
        if v is None or pd.isna(v):
            return "-"
    except Exception:
        if v is None:
            return "-"
    try:
        fv = float(v)
        if fv > 0:
            return f"<span class='pos'>+{fv:.2f}</span>"
        if fv < 0:
            return f"<span class='neg'>{fv:.2f}</span>"
        return "<span class='neutral'>0.00</span>"
    except Exception:
        return html_value(v)

def rates_table_html(rates: Dict[str, Any]) -> str:
    """예금은행 대출금리 동향: 신규취급액 기준 + 전월 대비 증감 항목."""
    rows = rates.get("table_rows_with_prev", []) if isinstance(rates, dict) else []
    if not rows and isinstance(rates, dict):
        rows = rates.get("table_rows", [])
    if not rows:
        return "<p class='muted'>금리 원자료 없음</p>"

    rate_cols = [
        ("mortgage_rate", "주택담보대출"),
        ("jeonse_rate", "전세자금대출"),
        ("sme_rate", "중소기업대출"),
        ("business_rate", "기업일반자금대출"),
    ]
    rows13 = rows[-13:]
    prev = rows[-14] if len(rows) > 13 else None
    html = []
    html.append("<div class='table-wrap rates-standard-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table time-series-table rates-standard-table'>")
    html.append("<thead><tr><th>기준월</th>")
    for _, label in rate_cols:
        html.append(f"<th>{label}</th><th>{label} 전월 대비 증감</th>")
    html.append("</tr></thead><tbody>")

    for idx, r in enumerate(rows13):
        tr_cls = " class='latest-row'" if idx == len(rows13) - 1 else ""
        html.append(f"<tr{tr_cls}><td>{html_value(r.get('month') or r.get('기준월'))}</td>")
        for key, _label in rate_cols:
            val = r.get(key)
            prev_val = prev.get(key) if isinstance(prev, dict) else None
            diff = r.get(f"{key}_mom")
            try:
                if (diff is None or pd.isna(diff)) and val is not None and prev_val is not None and not pd.isna(val) and not pd.isna(prev_val):
                    diff = float(val) - float(prev_val)
            except Exception:
                diff = None
            html.append(f"<td>{_fmt_rate(val)}</td><td>{_fmt_rate_diff(diff)}</td>")
        html.append("</tr>")
        prev = r
    html.append("</tbody></table></div>")
    return "".join(html)

def unsold_table_html(unsold: Dict[str, Any]) -> str:
    """표준템플릿 v58 미분양 주택 동향표.
    컬럼: 기준월, 미분양 호수, 전월 대비 증감, 준공후 미분양 호수, 준공후 전월 대비 증감
    """
    rows = unsold.get("table_rows_with_prev", []) if isinstance(unsold, dict) else []
    if not rows and isinstance(unsold, dict):
        rows = unsold.get("table_rows", [])
    if not rows:
        return "<p class='muted'>미분양 원자료 없음</p>"

    out = []
    prev_row = rows[-14] if len(rows) > 13 else None
    prev_unsold = None
    prev_after = None
    if isinstance(prev_row, dict):
        try:
            prev_unsold = float(prev_row.get("unsold_units") if prev_row.get("unsold_units") is not None else prev_row.get("미분양 호수"))
        except Exception:
            prev_unsold = None
        try:
            prev_after = float(prev_row.get("after_completion_unsold") if prev_row.get("after_completion_unsold") is not None else prev_row.get("준공후 미분양 호수"))
        except Exception:
            prev_after = None
    for r in rows[-13:]:
        m = r.get("month") or r.get("기준월")
        u = r.get("unsold_units") if r.get("unsold_units") is not None else r.get("미분양 호수")
        a = r.get("after_completion_unsold") if r.get("after_completion_unsold") is not None else r.get("준공후 미분양 호수")
        try:
            u_num = float(u) if u is not None and not pd.isna(u) else None
        except Exception:
            u_num = None
        try:
            a_num = float(a) if a is not None and not pd.isna(a) else None
        except Exception:
            a_num = None

        u_diff = r.get("unsold_units_mom")
        a_diff = r.get("after_completion_unsold_mom")
        if u_diff in [None, ""] or (isinstance(u_diff, float) and pd.isna(u_diff)):
            if prev_unsold is None or u_num is None:
                u_diff = None
            else:
                u_diff = u_num - prev_unsold
        if a_diff in [None, ""] or (isinstance(a_diff, float) and pd.isna(a_diff)):
            if prev_after is None or a_num is None:
                a_diff = None
            else:
                a_diff = a_num - prev_after

        out.append({
            "기준월": m,
            "미분양 호수": u_num,
            "전월 대비 증감": u_diff,
            "준공후 미분양 호수": a_num,
            "준공후 전월 대비 증감": a_diff,
        })
        if u_num is not None:
            prev_unsold = u_num
        if a_num is not None:
            prev_after = a_num

    html = []
    html.append("<div class='table-wrap unsold-standard-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table time-series-table unsold-standard-table'>")
    cols = ["기준월", "미분양 호수", "전월 대비 증감", "준공후 미분양 호수", "준공후 전월 대비 증감"]
    html.append("<thead><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr></thead><tbody>")
    for idx, r in enumerate(out):
        tr_cls = " class='latest-row'" if idx == len(out) - 1 else ""
        html.append(f"<tr{tr_cls}>")
        for c in cols:
            v = r.get(c)
            if c in ["전월 대비 증감", "준공후 전월 대비 증감"]:
                cell = _fmt_signed_num(v)
            else:
                cell = html_value(v, c)
            html.append(f"<td>{cell}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)

def supply_summary_table_html(supply: Dict[str, Any]) -> str:
    """표준템플릿 v58 공급 파이프라인 총괄표.
    컬럼: 기준연도, 인허가, 착공, 준공, 공동주택 분양, 인허가 증가율, 착공 전환율
    """
    rows = supply.get("table_rows", []) if isinstance(supply, dict) else []
    if not rows:
        return "<p class='muted'>공급 파이프라인 원자료 없음</p>"

    # 연도별로 재집계. 이미 연말 및 당해월 누계 형태이므로 같은 연도 중 뒤쪽 행을 우선.
    by_year = {}
    for r in rows:
        y = _supply_year_value(r)
        if not y:
            continue
        by_year[y] = {
            "기준연도": _supply_period_label(r),
            "인허가": _clean_num(r.get("permit") if r.get("permit") is not None else r.get("인허가")),
            "착공": _clean_num(r.get("start") if r.get("start") is not None else r.get("착공")),
            "준공": _clean_num(r.get("completion") if r.get("completion") is not None else r.get("준공")),
        }

    ordered = [by_year[y] for y in sorted(by_year.keys())][-5:]
    prev_permit = None
    for r in ordered:
        permit = r.get("인허가")
        start = r.get("착공")
        if prev_permit not in [None, 0] and permit is not None:
            r["인허가 증가율"] = (permit - prev_permit) / prev_permit * 100
        else:
            r["인허가 증가율"] = None
        if permit not in [None, 0] and start is not None:
            r["착공 전환율"] = start / permit * 100
        else:
            r["착공 전환율"] = None
        prev_permit = permit

    html = []
    html.append("<div class='table-wrap supply-summary-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table supply-summary-table'>")
    cols = ["기준연도", "인허가", "착공", "준공", "인허가 증가율", "착공 전환율"]
    html.append("<thead><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr></thead><tbody>")
    for idx, r in enumerate(ordered):
        tr_cls = " class='latest-row'" if idx == len(ordered)-1 else ""
        html.append(f"<tr{tr_cls}>")
        for c in cols:
            v = r.get(c)
            if c in ["인허가 증가율", "착공 전환율"]:
                cell = _fmt_pct(v)
            elif c == "기준연도":
                cell = html_value(v)
            else:
                cell = html_value(v)
            html.append(f"<td>{cell}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)


def move_in_supply_panel_html(move_in: Dict[str, Any]) -> str:
    if not isinstance(move_in, dict) or not move_in.get("available"):
        return "<p class='muted'>로컬 아파트 입주예정물량 DB 자료 없음</p>"
    year = move_in.get("target_year") or "-"
    scope = html.escape(str(move_in.get("scope_label") or "전체"))
    summary_rows = [dict(r) for r in move_in.get("region_summary_rows", [])]
    if summary_rows:
        summary_rows.append({
            "시군구": "합계",
            "입주예정 단지수": sum(int(r.get("입주예정 단지수") or 0) for r in summary_rows),
            "입주예정물량": sum(int(r.get("입주예정물량") or 0) for r in summary_rows),
            "전체 대비 비중": 100.0,
            "최대 집중월": "-",
            "집중월 물량": "-",
        })
    summary_html = table_rows_html(summary_rows, max_rows=41, highlight_latest=False)
    # 조사지역 행은 굵게·색으로, 합계 행은 굵게 구분한다.
    scope_name = str(move_in.get("scope_sigungu") or "")

    def _mark_row(table_html: str, cell_text: str, cls: str) -> str:
        if not cell_text:
            return table_html
        pieces = table_html.split("<tr")
        for i, piece in enumerate(pieces[1:], 1):
            if re.search(r">\s*" + re.escape(cell_text) + r"\s*<", piece):
                pieces[i] = (piece.replace(" class='", f" class='{cls} ", 1) if piece.startswith(" class='")
                             else f" class='{cls}'" + piece)
                break
        return "<tr".join(pieces)

    summary_html = _mark_row(_mark_row(summary_html, scope_name, "move-in-scope-row"), "합계", "move-in-total-row")
    parts = [
        f"<h4>표 3. {year}년 전남광주통합특별시 시군구별 아파트 입주예정물량</h4>",
        summary_html,
    ]
    if not move_in.get("has_region_data"):
        parts.append(f"<div class='move-in-empty'>{year}년 {scope}에 등록된 입주예정 단지가 없음.</div>")
    else:
        monthly = move_in.get("monthly_rows", [])
        peak = max([int(r.get("입주예정물량") or 0) for r in monthly] or [0])
        bars = []
        for row in monthly:
            value = int(row.get("입주예정물량") or 0)
            width = value / peak * 100 if peak else 0
            bars.append(
                "<div class='move-in-bar-row'>"
                f"<span>{html.escape(str(row.get('입주예정월') or '-'))}</span>"
                f"<div class='move-in-bar-track'><i style='width:{width:.1f}%'></i></div>"
                f"<b>{value:,}호</b><small>{int(row.get('입주예정 단지수') or 0)}개 단지</small>"
                "</div>"
            )
        project_display_rows = [{
            "단지명": row.get("단지명"),
            "주소": row.get("주소"),
            "세대수": row.get("세대수"),
            "입주예정일자": row.get("입주예정일자"),
            **({"시군구": row.get("시군구")} if not move_in.get("scope_sigungu") else {}),
        } for row in move_in.get("project_rows", [])]
        if project_display_rows:
            project_display_rows.append({
                "단지명": f"합계({len(project_display_rows)}개 단지)",
                "주소": "-",
                "세대수": sum(int(r.get("세대수") or 0) for r in project_display_rows),
                "입주예정일자": "-",
                **({"시군구": "-"} if not move_in.get("scope_sigungu") else {}),
            })
        parts.extend([
            f"<h4>차트 2. {year}년 {scope} 월별 아파트 입주예정물량</h4>",
            "<div class='move-in-bar-chart'>" + "".join(bars) + "</div>",
            f"<h4>표 4. {year}년 {scope} 아파트 입주예정 단지 핵심정보</h4>",
            "<div class='move-in-key-fields'>핵심 확인항목: <b>단지명 · 주소 · 세대수 · 입주예정일자</b></div>",
            _mark_row(table_rows_html(project_display_rows, max_rows=101, highlight_latest=False), f"합계({len(project_display_rows) - 1}개 단지)", "move-in-total-row"),
        ])
    parts.append(
        "<div class='move-in-source-note'>"
        f"※ 출처: {html.escape(str(move_in.get('source') or '로컬 DB'))} "
        f"(DB 반영: {html.escape(str(move_in.get('source_imported_at') or '-'))}). "
        f"산정: {html.escape(str(move_in.get('household_rule') or '-'))}. "
        f"지역분류: {html.escape(str(move_in.get('location_rule') or '-'))}. "
        "원자료는 입주예정월까지만 제공하므로 입주예정일자는 월 기준으로 표시하며, 사업 일정에 따라 변경될 수 있음."
        "</div>"
    )
    return "".join(parts)


def _supply_month_sort_key(row: Dict[str, Any]) -> int:
    m = str(row.get("month") or row.get("기준월") or row.get("기준연도") or "")
    nums = re.sub(r"[^0-9]", "", m)
    if len(nums) >= 6:
        return int(nums[:6])
    if len(nums) >= 4:
        return int(nums[:4] + "12")
    return 0

def _select_year_end_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """동일 연도·공급단계·물건유형 내에서는 최종 공표월 행만 사용.
    연말자료는 12월 누계가 최종 합계이고, 당해연도는 업로드된 최신 누계월을 사용함.
    """
    latest = {}
    for r in rows or []:
        y = _supply_year_value(r)
        st = str(r.get("stage") or "")
        typ = str(r.get("type") or "")
        key = (y, st, typ)
        if key not in latest or _supply_month_sort_key(r) >= _supply_month_sort_key(latest[key]):
            latest[key] = r
    return list(latest.values())

def non_apt_type_count_table_html(rows: List[Dict[str, Any]]) -> str:
    """비아파트 최근 13개월 거래 추이 표: 연립·다세대·단독·다가구 유형별 건수."""
    rows = rows or []
    if not rows:
        return "<p class='muted'>비아파트 물건유형별 원자료 없음</p>"
    months = sorted({str(r.get("month","")) for r in rows if r.get("month")})[-13:]
    types = ["연립", "다세대", "단독", "다가구"]
    html = []
    html.append("<div class='table-wrap nonapt-type-count-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table time-series-table nonapt-type-count-table nonapt-header-fixed-table'>")
    html.append("<thead><tr><th>기준월</th>" + "".join(f"<th>{t} 거래건수</th>" for t in types) + "<th>합계</th></tr></thead><tbody>")
    for idx, m in enumerate(months):
        tr_cls = " class='latest-row'" if idx == len(months)-1 else ""
        vals = [sum(float(r.get("trade_count") or 0) for r in rows if str(r.get("month")) == m and str(r.get("type")) == t) for t in types]
        html.append(f"<tr{tr_cls}><td>{html_value(m)}</td>" + "".join(f"<td>{int(round(v)):,}</td>" for v in vals) + f"<td>{int(round(sum(vals))):,}</td></tr>")
    html.append("</tbody></table></div>")
    return "".join(html)


SUPPLY_STAGE_CHART_COLORS = {
    "permit": {"bg": "rgba(251, 146, 60, 0.78)", "border": "#fb923c"},       # 인허가: 표의 주황 계열
    "start": {"bg": "rgba(96, 165, 250, 0.78)", "border": "#60a5fa"},        # 착공: 표의 파랑 계열
    "completion": {"bg": "rgba(52, 211, 153, 0.78)", "border": "#34d399"},   # 준공: 표의 초록 계열
}
def _supply_stage_color_by_key(key_or_title: str) -> Dict[str, str]:
    s = str(key_or_title or "")
    if "permit" in s or "인허가" in s:
        return SUPPLY_STAGE_CHART_COLORS["permit"]
    if "start" in s or "착공" in s:
        return SUPPLY_STAGE_CHART_COLORS["start"]
    if "completion" in s or "준공" in s:
        return SUPPLY_STAGE_CHART_COLORS["completion"]
    return {"bg": "rgba(148, 163, 184, 0.70)", "border": "#64748b"}


def mini_chart_from_rows_year_label(rows: List[Dict[str, Any]], title: str, chart_id: str, x: str, y_key: str, max_rows: int = 10) -> str:
    stage_color = _supply_stage_color_by_key(y_key or title)
    if not rows:
        return mini_chart_block(title, "bar", ["자료 없음"], [{"label": title, "data": [0], "backgroundColor": stage_color["bg"], "borderColor": stage_color["border"]}], chart_id)
    year_rows = {}
    for r in rows or []:
        y = _supply_year_value(r)
        if y:
            if y not in year_rows or _supply_month_sort_key(r) >= _supply_month_sort_key(year_rows[y]):
                year_rows[y] = r
    ordered = [year_rows[y] for y in sorted(year_rows.keys())][-max_rows:]
    return mini_chart_block(
        title,
        "bar",
        [_supply_period_label(r) for r in ordered],
        [{"label": title, "data": [r.get(y_key) for r in ordered], "backgroundColor": stage_color["bg"], "borderColor": stage_color["border"]}],
        chart_id
    )


def supply_by_type_table_html(supply: Dict[str, Any]) -> str:
    """표준템플릿 v58 물건유형별 공급 파이프라인 시계열표.
    연도별 최종 누계 기준: 과거연도는 12월 누계, 당해연도는 업로드된 최신 누계월 사용.
    """
    rows = supply.get("by_type_rows", []) if isinstance(supply, dict) else []
    if not rows:
        return "<p class='muted'>물건유형별 공급 원자료 없음</p>"
    rows = _select_year_end_rows(rows)
    stage_name = {"permit": "인허가", "start": "착공", "completion": "준공", "presale_supply": "공동주택 분양"}
    stage_order = ["인허가", "착공", "준공", "공동주택 분양"]
    stage_cls = {"인허가":"stage-permit", "착공":"stage-start", "준공":"stage-complete", "공동주택 분양":"stage-presale"}
    type_order = ["아파트", "연립", "다세대", "단독", "다가구"]
    norm = []
    for r in rows:
        y = _supply_year_value(r)
        if not re.match(r"20\d{2}$", str(y)): continue
        st = stage_name.get(str(r.get("stage")), str(r.get("stage")))
        typ = str(r.get("type") or "").strip()
        if "아파트" in typ: typ = "아파트"
        elif "연립" in typ: typ = "연립"
        elif "다세대" in typ: typ = "다세대"
        elif "단독" in typ: typ = "단독"
        elif "다가구" in typ: typ = "다가구"
        else: continue
        try: val = float(r.get("value") or 0)
        except Exception: val = 0
        norm.append({"year": str(y), "period_label": _supply_period_label(r), "stage": st, "type": typ, "value": val})
    if not norm:
        return "<p class='muted'>물건유형별 공급 원자료 없음</p>"
    df = pd.DataFrame(norm)
    years = sorted(df["year"].unique())[-5:]
    year_labels = {y: str(df[df["year"] == y]["period_label"].iloc[-1]) for y in years}
    grouped = df[df["year"].isin(years)].groupby(["stage", "type", "year"], as_index=False)["value"].sum()
    html = []
    html.append("<div class='supply-audit-note'>※ 수치 기준: 업로드 기준 이미지와 동일하게 과거연도는 12월 누계, 당해연도는 최신 누계월 값을 사용함. 인허가·착공·준공은 주택유형별 원자료의 아파트·연립·다세대·단독·다가구 행만 표시하며, 다가구는 가구수 기준임. 분양 항목은 공급 파이프라인 표와 차트에서 제외함.</div>")
    html.append("<div class='table-wrap supply-type-pipeline-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table supply-type-pipeline-table'>")
    html.append("<thead><tr><th>공급단계</th><th>물건유형</th>" + "".join(f"<th>{year_labels[y]}</th>" for y in years) + "<th>5개년 합계</th></tr></thead><tbody>")
    for st in stage_order:
        cls = stage_cls.get(st, "")
        for idx, typ in enumerate(type_order):
            vals = []
            for y in years:
                sub = grouped[(grouped["stage"] == st) & (grouped["type"] == typ) & (grouped["year"] == y)]
                vals.append(float(sub["value"].sum()) if not sub.empty else 0)
            total = sum(vals)
            html.append(f"<tr class='{cls}'>")
            if idx == 0:
                html.append(f"<td class='stage-cell {cls}' rowspan='{len(type_order)}'>{st}</td>")
            html.append(f"<td>{typ}</td>")
            for v in vals:
                html.append(f"<td>{int(round(v)):,}</td>")
            html.append(f"<td class='stage-total {cls}'>{int(round(total)):,}</td></tr>")
    html.append("</tbody></table></div>")
    return "".join(html)


def period_management_table(data: Dict[str, Any]) -> str:
    meta = data.get("meta", {})
    supply_rows = data.get("supply", {}).get("table_rows", []) or []
    supply_period = _supply_period_label(supply_rows[-1]) if supply_rows else "-"
    rows = [
        {"지표": "실거래", "적용 기준월": meta.get("base_month"), "자료 성격": "계약일 기준 잠정치", "보고서 해석 기준": "신고기한 미도래로 향후 거래건수 보정 가능"},
        {"지표": "인구·세대수", "적용 기준월": data.get("population", {}).get("latest_month"), "자료 성격": "주민등록 기준", "보고서 해석 기준": "기초 주거수요와 가구 분화 흐름 판단"},
        {"지표": "인구이동", "적용 기준월": data.get("migration", {}).get("latest_month"), "자료 성격": "최신 공표월", "보고서 해석 기준": "전입·전출 및 연령대별 이동 흐름 판단"},
        {"지표": "예금은행 대출금리", "적용 기준월": data.get("rates", {}).get("latest_month"), "자료 성격": "신규취급액 기준", "보고서 해석 기준": "주거금융 및 사업체 금융비용 부담 판단"},
        {"지표": "미분양·준공후 미분양", "적용 기준월": data.get("unsold", {}).get("latest_month"), "자료 성격": "최신 공표월", "보고서 해석 기준": "분양시장 부담 및 입주 단계 공급 부담 판단"},
        {"지표": "주택공급", "적용 기준월": supply_period, "자료 성격": "과거연도 12월 누계·당해연도 최종 공표월 누계", "보고서 해석 기준": "인허가·착공·준공을 통한 중기 공급 압력 판단"},
    ]
    return table_rows_html(rows, max_rows=20)


def _major_case_money_pair(row: Dict[str, Any], value_label: str) -> Tuple[str, str, str]:
    if value_label == "매매":
        curr = row.get("current_price_eok", row.get("current_value"))
        prev = row.get("prev_price_eok", row.get("prev_value"))
        curr_s = _fmt_eok_value(curr)
        prev_s = _fmt_eok_value(prev)
        try:
            delta_s = _fmt_eok_value(float(curr) - float(prev))
        except Exception:
            delta_s = "-"
        return curr_s, prev_s, delta_s
    elif value_label == "전세":
        curr = row.get("current_value")
        prev = row.get("prev_value")
        curr_s = _fmt_manwon_value(curr)
        prev_s = _fmt_manwon_value(prev)
        try:
            delta_s = _fmt_manwon_value(float(curr) - float(prev))
        except Exception:
            delta_s = "-"
        return curr_s, prev_s, delta_s
    else:
        # 월세는 환산월세 기준 비교. 보증금/월세는 보조로 별도 표기.
        curr = row.get("current_value")
        prev = row.get("prev_value")
        curr_s = _fmt_manwon_value(curr)
        prev_s = _fmt_manwon_value(prev)
        try:
            delta_s = _fmt_manwon_value(float(curr) - float(prev))
        except Exception:
            delta_s = "-"
        return curr_s, prev_s, delta_s

def _parse_case_contract_date(row: Dict[str, Any]) -> Optional[datetime]:
    raw = str(row.get("contract_date") or "").strip()
    month = str(row.get("month") or "").strip()[:7]
    digits = re.sub(r"[^0-9]", "", raw)
    candidates = []
    if len(digits) >= 8:
        candidates.append(digits[:8])
    elif len(digits) >= 1 and re.fullmatch(r"\d{4}-\d{2}", month):
        candidates.append(month.replace("-", "") + digits[-2:].zfill(2))
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
        candidates.append(raw.replace("-", ""))
    for value in candidates:
        try:
            return datetime.strptime(value, "%Y%m%d")
        except Exception:
            continue
    return None

def _select_weekly_major_case(section_obj: Dict[str, Any], preferred_sources: List[Tuple[str, str]], reference_date: str) -> Tuple[Optional[Dict[str, Any]], str, str]:
    try:
        ref = datetime.strptime(str(reference_date)[:10], "%Y-%m-%d")
    except Exception:
        return None, "", ""
    for offset, period_label in [(0, "기준주"), (-1, "전주")]:
        start = ref + timedelta(days=offset * 7)
        end = start + timedelta(days=6)
        for key, direction_word in preferred_sources:
            rows = list(section_obj.get(key) or [])
            for src in rows:
                case_date = _parse_case_contract_date(src)
                if case_date and start.date() <= case_date.date() <= end.date():
                    row = dict(src)
                    return row, direction_word, period_label
    return None, "", ""

def _typical_case_rows(rows: List[Dict[str, Any]], value_label: str) -> List[Dict[str, Any]]:
    """대표 거래사례 후보: 이상치(TOP 카드 ⚠ 기준)·갱신계약·6개월 넘은 직전 비교를 빼고,
    가장 거래가 많은 읍면동의 변동률 중간값에 가까운 순서로 정렬한다. 남는 후보가 없으면 원래 목록."""
    lo, hi = (-100 / 3, 50.0) if value_label == "매매" else (-50.0, 100.0)
    picked = []
    for r in rows or []:
        try:
            chg = float(r.get("change_rate"))
        except Exception:
            continue
        if not (lo < chg < hi) or abs(chg) < 0.5:
            continue
        if str(r.get("contract_class") or "") == "갱신재계약":
            continue
        cur = _parse_case_contract_date(r)
        prev = _parse_case_contract_date({"contract_date": r.get("prev_contract_date"), "month": r.get("prev_month")})
        if cur and prev and (cur - prev).days > 186:
            continue
        picked.append((chg, r))
    if not picked:
        return list(rows or [])
    dong_count: Dict[str, int] = {}
    for _, r in picked:
        dong_count[str(r.get("dong") or "")] = dong_count.get(str(r.get("dong") or ""), 0) + 1
    top_dong = max(dong_count, key=dong_count.get)
    pool = [(c, r) for c, r in picked if str(r.get("dong") or "") == top_dong] if dong_count[top_dong] >= 2 else picked
    changes = sorted(c for c, _ in pool)
    mid = changes[len(changes) // 2]
    rest = [r for _, r in sorted(picked, key=lambda x: abs(x[0] - mid)) if not any(r is q for _, q in pool)]
    return [r for _, r in sorted(pool, key=lambda x: abs(x[0] - mid))] + rest


def _rate_change_phrase(prev: Any, cur: Any) -> str:
    """조사자 시황과 같은 변동 표현: 상승폭 확대·축소, 하락폭 확대·축소, 상승·하락·보합 전환."""
    try:
        p, c = round(float(prev), 2), round(float(cur), 2)
    except Exception:
        return ""
    if c > 0:
        if p > 0:
            return "상승폭 확대" if c > p else ("상승폭 축소" if c < p else "상승폭 유지")
        return "상승 전환"
    if c < 0:
        if p < 0:
            return "하락폭 확대" if c < p else ("하락폭 축소" if c > p else "하락폭 유지")
        return "하락 전환"
    return "보합 유지" if p == 0 else "보합 전환"


def _base_month_is_partial(meta: Dict[str, Any]) -> bool:
    """기준월 거래 신고기한(계약 후 30일)이 지나기 전에 만든 보고서인지."""
    try:
        base = pd.Period(str(meta.get("base_month") or "")[:7], freq="M")
        created = pd.Timestamp(str(meta.get("created_at") or "")[:10])
        return created < base.end_time.normalize() + pd.Timedelta(days=30)
    except Exception:
        return True


def market_head_line_html(section_obj: Dict[str, Any], label: str, meta: Dict[str, Any]) -> str:
    """Ⅵ장 아파트 문단 첫 줄: 가격지수 변동률·거래량·실거래 상승/하락 비중의 전월→당월 요약."""
    rows = [r for r in (section_obj or {}).get("monthly_rows") or [] if isinstance(r, dict)]
    if len(rows) < 2:
        return ""
    prev, cur = rows[-2], rows[-1]
    parts = []
    idx = [r for r in rows if r.get("public_index_change_rate") is not None]
    if len(idx) >= 2:
        a, b = idx[-2], idx[-1]
        phrase = _rate_change_phrase(a.get("public_index_change_rate"), b.get("public_index_change_rate"))
        month_note = "" if str(b.get("month")) == str(cur.get("month")) else f", {str(b.get('month'))[5:7].lstrip('0')}월 기준"
        asset = str((meta or {}).get("primary_asset_label") or "아파트")
        parts.append(f"{asset} {label}가격 변동률 {phrase}({float(a['public_index_change_rate']):.2f}% → {float(b['public_index_change_rate']):.2f}%{month_note})")
    try:
        pc, cc = int(prev.get("trade_count") or 0), int(cur.get("trade_count") or 0)
        if _base_month_is_partial(meta or {}):
            parts.append(f"거래 {cc:,}건(전월 {pc:,}건, 당월 신고 진행 중)")
        else:
            word = "증가" if cc > pc else ("감소" if cc < pc else "유지")
            parts.append(f"거래량 {word}({pc:,}건 → {cc:,}건)")
    except Exception:
        pass
    basis = "신규계약 " if label in ("전세", "월세") else ""
    try:
        pv, cv = str(prev.get("판정") or ""), str(cur.get("판정") or "")
        verdict = f"{cv} 유지" if pv == cv else f"{pv}→{cv}"
        parts.append(f"실거래 기준 {basis}상승 {float(prev.get('up_ratio') or 0):.1f}% → {float(cur.get('up_ratio') or 0):.1f}%, "
                     f"하락 {float(prev.get('down_ratio') or 0):.1f}% → {float(cur.get('down_ratio') or 0):.1f}%({verdict})")
    except Exception:
        pass
    if not parts:
        return ""
    return f"<div class='market-head-line'>① {html.escape(' · '.join(parts))}</div>"


def major_case_line_html(section_obj: Dict[str, Any], value_label: str) -> str:
    """Ⅵ장 아파트 매매·전세·월세 문단 하단에 당월 상승/하락 분위기에 맞는 주요 거래사례 1건 표시."""
    if not isinstance(section_obj, dict):
        return ""
    section_obj = dict(section_obj)
    for _key in ("all_up_rows", "all_down_rows", "top_up_rows", "top_down_rows"):
        if section_obj.get(_key):
            section_obj[_key] = _typical_case_rows(section_obj.get(_key), value_label)
    try:
        up = float(section_obj.get("up_ratio") or 0)
        down = float(section_obj.get("down_ratio") or 0)
    except Exception:
        up, down = 0, 0

    value_key = "all_value_rows" if section_obj.get("all_value_rows") else "top_value_rows"
    price_key = "all_price_rows" if section_obj.get("all_price_rows") else "top_price_rows"
    if down > up:
        preferred_sources = [
            ("all_down_rows" if section_obj.get("all_down_rows") else "top_down_rows", "하락"),
            ("all_up_rows" if section_obj.get("all_up_rows") else "top_up_rows", "상승"),
            (value_key if section_obj.get(value_key) else price_key, "상위"),
        ]
    else:
        preferred_sources = [
            ("all_up_rows" if section_obj.get("all_up_rows") else "top_up_rows", "상승"),
            ("all_down_rows" if section_obj.get("all_down_rows") else "top_down_rows", "하락"),
            (value_key if section_obj.get(value_key) else price_key, "상위"),
        ]

    week_case_label = ""
    reference_date = str(section_obj.get("weekly_case_reference_date") or "").strip()
    if reference_date:
        row, direction_word, week_case_label = _select_weekly_major_case(section_obj, preferred_sources, reference_date)
    else:
        row, direction_word = None, ""

    if not row:
        for _key, _word in preferred_sources[:2]:
            if section_obj.get(_key):
                row, direction_word = list(section_obj.get(_key))[0], _word
                break
    if row:
        pass
    elif down > up and section_obj.get("top_down_rows"):
        row = list(section_obj.get("top_down_rows") or [])[0]
        direction_word = "하락"
    elif section_obj.get("top_up_rows"):
        row = list(section_obj.get("top_up_rows") or [])[0]
        direction_word = "상승"
    elif section_obj.get("top_value_rows"):
        row = list(section_obj.get("top_value_rows") or [])[0]
        direction_word = "상위"
    elif section_obj.get("top_price_rows"):
        row = list(section_obj.get("top_price_rows") or [])[0]
        direction_word = "상위"
    else:
        return ""

    name = html.escape(str(row.get("building_name") or row.get("complex_name") or row.get("apt_name") or row.get("_group_name") or "-"))
    year = str(row.get("approval_year") or "").replace(".0", "")
    year_s = f" (사용승인년도 {year}년)" if year and year.lower() != "nan" else ""
    area = row.get("area")
    pyeong = row.get("pyeong") or area_to_pyeong(area)
    area_s = f"전용 {area}㎡" if area not in [None, ""] else "전용면적 미상"
    if pyeong:
        area_s += f"({pyeong})"

    curr_s, prev_s, delta_s = _major_case_money_pair(row, value_label)
    curr_date = format_contract_date_display(row.get("contract_date") or "", row.get("month") or "")
    prev_date = format_contract_date_display(row.get("prev_contract_date") or "", row.get("prev_month") or "")
    try:
        chg = float(row.get("change_rate"))
        chg_s = f"{chg:+.1f}%"
    except Exception:
        chg_s = "-"

    if value_label == "월세":
        dep = _fmt_manwon_value(row.get("current_deposit"))
        mon = _fmt_manwon_value(row.get("current_monthly_rent"))
        extra = f" / 보증금 {dep}, 월세 {mon}"
    else:
        extra = ""

    current_part = f"{curr_s}({curr_date})" if curr_date else curr_s
    prev_part = f"{prev_s}({prev_date})" if prev_date else prev_s
    sign = "" if str(delta_s).startswith("-") else "+"
    return (
        "<div class='case-spacer'></div>"
        f"<div class='major-case-line summary-major-case'>* 주요 거래사례{f'({week_case_label})' if week_case_label else ''}: "
        f"{name}{year_s} 실거래 {area_s}: {prev_part} → {current_part} / "
        f"<span class='case-{direction_word}'>{sign}{delta_s}, {chg_s}</span>{extra}</div>"
    )



# ---------------------------------------------------------------------
# 빠른 읽기 기준 자동 하이라이트
# ---------------------------------------------------------------------
SEMANTIC_HIGHLIGHT_KEYS = set(COMMENTARY_KEYS)

def _strip_html_for_score(text: str) -> str:
    return re.sub(r"<[^>]+>", "", str(text or ""))

def _sentence_highlight_class(sentence: str) -> Tuple[int, str]:
    """문장 단위 핵심성 점수와 색상 판정.
    v5.8.38부터 단어마다 형광펜을 칠하지 않고, LLM이 생성한 해설 중
    핵심 판단 문장 1~2개만 강조한다.
    """
    s = _strip_html_for_score(sentence)
    if len(s.strip()) < 18:
        return 0, ""

    number_pat = r"[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*(?:%|건|호|명|세대|억|억원|만원|㎡|평|년|개월|p|bp)"
    has_num = bool(re.search(number_pat, s))

    danger_words = ["하락", "감소", "위축", "부담", "위험", "미분양", "고금리", "유출", "순유출", "부진", "침체", "하방", "거래 둔화", "불확실성", "약세"]
    positive_words = ["상승", "증가", "회복", "개선", "안정", "방어", "견조", "선호", "유입", "순유입", "활발", "실수요", "반등", "확대", "우세"]
    judgment_words = ["판단됨", "분석됨", "해석됨", "확인됨", "관찰됨", "평가됨", "보임", "나타남", "필요", "점검", "종합", "선별적", "국지적", "시장 판단"]

    danger = sum(1 for w in danger_words if w in s)
    positive = sum(1 for w in positive_words if w in s)
    judgment = sum(1 for w in judgment_words if w in s)

    score = 0
    if has_num:
        score += 2
    score += danger * 3 + positive * 3 + judgment * 2
    # 핵심 판단 문장 우선: 문장 길이가 너무 길면 과도한 형광펜이 되므로 약간 감점
    if len(s) > 190:
        score -= 1

    if score < 4:
        return 0, ""

    if danger > positive and danger >= 1:
        cls = "hl-red"
    elif positive > danger and positive >= 1:
        cls = "hl-green"
    elif judgment >= 1:
        cls = "hl-blue"
    elif has_num:
        cls = "hl-key"
    else:
        cls = ""

    return score, cls

def _split_korean_sentences_keep_delim(text: str) -> List[str]:
    """마침표·물음표·느낌표 기준으로 문장을 나누되 구분자를 유지."""
    if not text:
        return []
    # HTML 태그가 포함된 경우에도 문장 자체는 단순 분리. 템플릿 해설은 대부분 순수 텍스트임.
    parts = re.split(r"(?<=[.!?。])\s+|(?<=\.)\n+|\n+", str(text))
    return [p for p in parts if p.strip()]

def apply_semantic_highlights(text: Any) -> str:
    """해설 문장에 빠른 읽기 기준 색상 span을 적용한다.
    v5.8.38 기준:
    - 단어 단위 강조 금지
    - LLM 해설 중 핵심 문장만 최대 2개 강조
    - 한 문단 전체가 과도하게 칠해지지 않도록 220자 초과 문장은 강조하지 않음
    """
    if text is None:
        return ""
    s = str(text)
    if not s:
        return ""

    # 이미 사용자가 직접 강조한 HTML은 유지하고 자동 중복 강조하지 않음
    if any(cls in s for cls in ["hl-key", "hl-yellow", "hl-red", "hl-green", "hl-blue"]):
        return s

    sentences = _split_korean_sentences_keep_delim(s)
    if len(sentences) <= 0:
        return s

    scored = []
    for i, sent in enumerate(sentences):
        score, cls = _sentence_highlight_class(sent)
        # 긴 문장 전체 강조는 가독성 저하 방지
        if len(_strip_html_for_score(sent)) > 220:
            score = 0
            cls = ""
        if score > 0 and cls:
            scored.append((score, i, cls))

    # 문단당 핵심문장 최대 2개만 강조
    selected = {}
    for score, idx, cls in sorted(scored, key=lambda x: (-x[0], x[1]))[:2]:
        selected[idx] = cls

    if not selected:
        return s

    out = []
    for i, sent in enumerate(sentences):
        cls = selected.get(i)
        if cls:
            out.append(f"<span class='hl-sentence {cls}'>" + sent.strip() + "</span>")
        else:
            out.append(sent.strip())
    return " ".join(out)

def apply_commentary_semantic_highlights(commentary: Dict[str, str]) -> Dict[str, str]:
    """COMMENTARY_KEYS 전체에 빠른 읽기 기준 색상 로직 적용."""
    out = dict(commentary or {})
    for k in COMMENTARY_KEYS:
        if k in out:
            out[k] = apply_semantic_highlights(out.get(k, ""))
    return out


_DATA_ABSENCE_PATTERNS = re.compile(
    r"(확보되지 않|자료가 없|자료가 확인되지 않|시황 자료|조사 시황이 확인되지|공표치가 확인되지|공표 결과가 확인되지)")


def _drop_data_absence_sentences(text: str) -> str:
    """'자료가 확보되지 않아'처럼 자료 부재를 말하는 문장을 뺀다(조사자 시황에는 없는 문장)."""
    sentences = re.split(r"(?<=[.。])\s+", text)
    kept = [x for x in sentences if not _DATA_ABSENCE_PATTERNS.search(x)]
    return " ".join(kept) if kept else text


def sanitize_final_commentary_terms(commentary: Dict[str, str]) -> Dict[str, str]:
    """Ⅵ장 서술 문단에서 딱딱한 통계 출처 표현을 자연스러운 시장 언어로 완화하고, 자료 부재 문장을 뺀다."""
    out = dict(commentary or {})
    for k in COMMENTARY_KEYS:
        text = str(out.get(k, "") or "")
        text = text.replace("공표 통계", "통계 지표").replace("공표통계", "통계 지표")
        out[k] = _drop_data_absence_sentences(text) if text else text
    return out


_REGION_TAG = re.compile(r"\s*([▲▼◆])\s*(상승|하락|혼조) 지역\(([^)]{1,30})\)\s*:\s*")


def mark_region_tags(commentary: Dict[str, str]) -> Dict[str, str]:
    """'▲ 상승 지역(○○동):' 표시를 줄을 바꾼 색 라벨로 바꾼다."""
    out = dict(commentary or {})
    cls = {"상승": "up", "하락": "down", "혼조": "mixed"}
    for k in ("apt_sale_summary", "apt_jeonse_summary", "apt_rent_summary"):
        text = str(out.get(k, "") or "")
        if text:
            out[k] = _REGION_TAG.sub(lambda m: f"<br><span class='region-tag region-{cls[m.group(2)]}'>{m.group(1)} {m.group(2)} 지역({m.group(3)})</span> ", text)
            if out[k].startswith("<br>"):
                out[k] = out[k][4:]
    return out


def _dedupe_cluster_legend(page: str) -> str:
    """같은 장(section) 안에서 반복되는 반도체 영향권 범례는 첫 번째만 남긴다."""
    legend = semiconductor_cluster_legend_html()
    parts = page.split("<section")
    for i, part in enumerate(parts):
        first = part.find(legend)
        if first >= 0:
            head, tail = part[: first + len(legend)], part[first + len(legend):]
            parts[i] = head + tail.replace(legend, "")
    return "<section".join(parts)


def render_html(data: Dict[str, Any], commentary: Dict[str, str], validation_rows: List[Dict[str, str]]) -> str:
    commentary = sanitize_final_commentary_terms(commentary)
    commentary = apply_commentary_semantic_highlights(commentary)
    commentary = mark_region_tags(commentary)
    global _CURRENT_BASE_MONTH, _CURRENT_CREATED_DATE
    _CURRENT_BASE_MONTH = str((data.get("meta") or {}).get("base_month") or "").strip()
    _CURRENT_CREATED_DATE = str((data.get("meta") or {}).get("created_at") or "").strip()[:10]
    env = Environment(autoescape=False)
    env.filters["fmt_num"] = format_num
    env.filters["fmt_float"] = format_float
    env.filters["as_table"] = table_rows_html
    env.filters["as_dong_volume_trend"] = dong_volume_trend_table_html
    env.filters["as_cards"] = top_cards_html
    env.globals["top_case_list"] = top_case_list_html
    env.globals["major_case_line"] = major_case_line_html
    env.globals["market_head_line"] = lambda obj, label: market_head_line_html(obj, label, data.get("meta") or {})
    tpl = env.from_string(HTML_TEMPLATE)
    conv = data.get("conversion_rate", {})
    conv_apt = data.get("conversion_rate_apt") or {}
    conv_rh = data.get("conversion_rate_rh") or {}
    conv_offi = data.get("conversion_rate_offi") or {}
    is_officetel_report = (
        str(data.get("meta", {}).get("report_mode", "")).lower() == "officetel"
        or bool(data.get("meta", {}).get("is_officetel"))
        or str(data.get("meta", {}).get("primary_asset_label", "")).strip() == "오피스텔"
    )
    if is_officetel_report and conv_offi.get("applied_rate") not in (None, "", "-"):
        conv = conv_offi
    conv_region = conv.get("region") or conv.get("matched_region") or data.get("meta", {}).get("region") or "-"
    conv_asset_type = conv.get("asset_type") or conv.get("public_type") or data.get("meta", {}).get("primary_asset_label") or "-"
    conv_source = conv.get("source") or "KOSIS 공표 전월세 전환율"
    def _fmt_rate_label(value: Any) -> str:
        try:
            s = f"{float(value):.4f}".rstrip("0").rstrip(".")
            if "." not in s:
                s += ".0"
            return f"{s}%"
        except Exception:
            return "-"
    rate_parts = []
    if is_officetel_report:
        offi_rate = conv.get("applied_rate") or conv_offi.get("applied_rate")
        if offi_rate not in (None, "", "-"):
            rate_parts.append(f"오피스텔 {_fmt_rate_label(offi_rate)}")
    else:
        apt_rate = conv_apt.get("applied_rate") or conv.get("applied_rate")
        rh_rate = conv_rh.get("applied_rate")
        if apt_rate not in (None, "", "-"):
            rate_parts.append(f"아파트 {_fmt_rate_label(apt_rate)}")
        if rh_rate not in (None, "", "-"):
            rate_parts.append(f"연립·다세대 {_fmt_rate_label(rh_rate)}")
    applied_rate_display = ", ".join(rate_parts) if rate_parts else _fmt_rate_label(conv.get("applied_rate"))
    conversion_rate_table = table_rows_html([{
        "구분": "월세 환산 기준",
        "내용": "KOSIS 공표 전월세 전환율을 시도·물건유형 기준으로 매칭해 적용"
    }, {
        "구분": "적용 지역",
        "내용": conv_region
    }, {
        "구분": "적용 물건유형",
        "내용": conv_asset_type
    }, {
        "구분": "적용 공표월",
        "내용": conv.get("applied_month") or conv.get("period") or "-"
    }, {
        "구분": "적용 전환율",
        "내용": applied_rate_display
    }, {
        "구분": "전환율 출처",
        "내용": conv_source
    }], max_rows=20)
    html = tpl.render(
        meta=data["meta"], cards=data["summary_cards"], warnings=data.get("data_quality_warnings", []),
        population=data["population"], age_population=data.get("age_population", {}), migration=data["migration"], rates=data["rates"], unsold=data["unsold"],
        price_index=data["price_index"], supply=data["supply"], apt_trade=data["apt_trade"], apt_rent=data["apt_rent"],
        external_buyer=data.get("external_buyer_sale", {}),
        external_current_table=table_rows_html(data.get("external_buyer_sale", {}).get("current_region_rows", []), max_rows=15, highlight_latest=False),
        external_six_month_table=table_rows_html(data.get("external_buyer_sale", {}).get("six_month_region_rows", []), max_rows=15, highlight_latest=False),
        non_apt_sale=data["non_apt_sale"], non_apt_rent=data["non_apt_rent"], presale=data["presale"],
        policy_context=data.get("policy_context", {}),
        conversion_rate_table=conversion_rate_table,
        conversion_rate=conv,
        period_table=period_management_table(data),
        rates_standard_table=rates_table_html(data.get("rates", {})),
        unsold_standard_table=unsold_table_html(data.get("unsold", {})),
        unsold_news_html=unsold_news_html(data.get("unsold_news", {}), commentary.get("unsold_news_summary", "")),
        supply_summary_table=supply_summary_table_html(data.get("supply", {})),
        supply_by_type_table=supply_by_type_table_html(data.get("supply", {})),
        move_in_supply_panel=move_in_supply_panel_html(data.get("move_in_supply", {})),
        non_apt_sale_type_table=non_apt_type_count_table_html(data.get("non_apt_sale", {}).get("monthly_by_type_rows", [])),
        non_apt_jeonse_type_table=non_apt_type_count_table_html(data.get("non_apt_rent", {}).get("jeonse", {}).get("monthly_by_type_rows", [])),
        non_apt_rent_type_table=non_apt_type_count_table_html(data.get("non_apt_rent", {}).get("monthly_rent", {}).get("monthly_by_type_rows", [])),
        population_migration_table=population_migration_table_html(data),
        age_population_trend_table=age_population_trend_table_html(data.get("age_population", {})),
        age_migration_trend_table=age_migration_trend_table_html(data.get("migration", {})),
        charts=make_chart_html(data),
        text=commentary, validation_rows=validation_rows,
        show_move_in_supply=is_jeonnam_gwangju_region((data.get("meta") or {}).get("region", "")),
        policy_inflection_notes=policy_inflection_notes_html(data)
    )
    return _dedupe_cluster_legend(ensure_template_v58_structure(html))


def validate_report(html: str, data: Dict[str, Any], commentary: Dict[str, str]) -> List[Dict[str, str]]:
    rows = []
    weekly_mode = is_weekly_report(data)
    officetel_mode = str(data.get("meta", {}).get("report_mode", "")) == "officetel" or bool(data.get("meta", {}).get("is_officetel"))
    external_buyer_mode = bool(data.get("external_buyer_sale", {}).get("available"))
    expected_table_count = 43 if external_buyer_mode else 40
    text = BeautifulSoup(html, "html.parser").get_text("\n")
    def add(item, ok, note=""):
        rows.append({"검수항목": item, "결과": "정상" if ok else "확인필요", "비고": note})
    required_sections = [
        ("Ⅰ. 인구·세대 및 이동 동향", ["Ⅰ", "인구", "동향"]),
        ("Ⅱ. 금리·미분양·주택공급 동향", ["Ⅱ", "금리·미분양·주택공급 동향"]),
        ("Ⅲ. 매매 주택 거래 현황 분석", ["Ⅲ", "매매 주택 거래 현황 분석"]),
        ("Ⅳ. 전세 주택 거래 현황 분석", ["Ⅳ", "전세 주택 거래 현황 분석"]),
        ("Ⅴ. 월세 주택 거래 현황 분석", ["Ⅴ", "월세 주택 거래 현황 분석"]),
        ("Ⅵ. 종합 시황 서술 보고서", ["Ⅵ", "종합 시황 서술 보고서"]),
    ]
    if weekly_mode:
        # 주간은 월세 장이 없어 종합 장이 Ⅴ가 된다.
        required_sections = [x for x in required_sections if not x[0].startswith(("Ⅴ.", "Ⅵ."))] + [("Ⅴ. 종합 시황 서술 보고서", ["Ⅴ", "종합 시황 서술 보고서"])]
    for sec, needles in required_sections:
        add(f"필수 섹션 존재: {sec}", all(n in text for n in needles))
    excluded_commentary_keys = set()
    if weekly_mode:
        excluded_commentary_keys.update([
            "apt_rent_commentary", "non_apt_sale_commentary", "non_apt_jeonse_commentary", "non_apt_rent_commentary",
            "rent_market_summary", "apt_rent_summary", "non_apt_sale_regional_context", "non_apt_jeonse_regional_context", "non_apt_rent_regional_context",
            "rowhouse_multiplex_sale_summary", "detached_multifamily_sale_summary", "rowhouse_multiplex_jeonse_summary", "detached_multifamily_jeonse_summary",
            "rowhouse_multiplex_rent_summary", "detached_multifamily_rent_summary",
        ])
    if officetel_mode:
        excluded_commentary_keys.update([
            "non_apt_sale_commentary", "non_apt_jeonse_commentary", "non_apt_rent_commentary", "presale_commentary",
            "non_apt_sale_regional_context", "non_apt_jeonse_regional_context", "non_apt_rent_regional_context",
            "rowhouse_multiplex_sale_summary", "detached_multifamily_sale_summary", "rowhouse_multiplex_jeonse_summary", "detached_multifamily_jeonse_summary",
            "rowhouse_multiplex_rent_summary", "detached_multifamily_rent_summary",
        ])
    for k in COMMENTARY_KEYS:
        add(f"해설 key 존재: {k}", k in excluded_commentary_keys or bool(commentary.get(k)))
    llm_applied_keys = set(commentary.get("_llm_applied_keys") or [])
    if weekly_mode:
        llm_summary_keys = WEEKLY_LLM_FINAL_SUMMARY_KEYS
    elif officetel_mode:
        llm_summary_keys = ["comprehensive_summary", "sale_market_summary", "apt_sale_summary", "jeonse_market_summary", "apt_jeonse_summary", "rent_market_summary", "apt_rent_summary"]
    else:
        llm_summary_keys = LLM_FINAL_SUMMARY_KEYS
    llm_summary_present = all(str(commentary.get(k) or "").strip() for k in llm_summary_keys)
    llm_summary_applied = all(k in llm_applied_keys for k in llm_summary_keys)
    add(
        "Ⅵ 종합 시황 LLM JSON 반영",
        commentary.get("_commentary_source") == "llm_json" and llm_summary_present and llm_summary_applied,
        "LLM JSON 반영본에서 최종 보고서로 확정"
    )
    for phrase in FORBIDDEN_PHRASES:
        add(f"금지문구 미포함: {phrase}", phrase not in text)
    add("HTML 본문 검수표 미출력", "Ⅶ. 최종 검수표" not in text)
    add("임시 확장 예정 문구 미포함", "확장 예정" not in text and "v5.8.38에서" not in text)
    add("Chart.js 차트 포함", "new Chart" in html and "<canvas" in html)
    missing_v58 = [cid for cid in TEMPLATE_V58_REQUIRED_CANVAS_IDS if (f'id="{cid}"' not in html and f"id='{cid}'" not in html)]
    add("template_v58 필수 canvas id 보존", len(missing_v58) == 0, f"missing={len(missing_v58)}")
    add("연령대 순이동 통합 후 canvas 55개 이상", html.count("<canvas") >= 55, f"canvas={html.count('<canvas')}")
    add("template_v58 주요 클래스 보존", all(cls in html for cls in ["survey-donut-card","conversion-card","supply-card","age-move-card","mini-chart-card","top-case-item","table-wrap"]), "template_v58 key classes")
    add("template_v58 의미색상 JS 포함", "semanticChartColors" in html and "applySemanticColors" in html and "prepareChartConfig" in html)
    add("template_v58 지연 렌더링 적용", "registerLazyChart" in html and "__lazyChartQueue" in html)
    canvas_ids = re.findall(r"<canvas[^>]+id=[\'\"]([^\'\"]+)[\'\"]", html)
    extra_canvas = [cid for cid in canvas_ids if cid not in TEMPLATE_V58_REQUIRED_CANVAS_IDS]
    add("템플릿 외 차트 없음", len(extra_canvas) == 0, f"extra={extra_canvas[:5]}")
    add("숨김 차트 bank 미생성", "template-v58-canvas-bank" not in html and "template-v58-class-bank" not in html)
    canvas_id_list = re.findall(r"<canvas[^>]+id=['\"]([^'\"]+)['\"]", html)
    add("canvas id 중복 없음", len(canvas_id_list) == len(set(canvas_id_list)), f"duplicates={[k for k,v in Counter(canvas_id_list).items() if v>1] if 'Counter' in globals() else ''}")
    add("template_v58 canvas 정확히 58개", len(canvas_id_list) == 58, f"canvas={len(canvas_id_list)}")
    add("template_v58 chart-box 정확히 36개", html.count("class='chart-box") + html.count('class="chart-box') == 36, f"chart-box={html.count('chart-box')}")
    add("chart3·chart11 누락 방지", "id='chart3'" in html and "id='chart11'" in html)
    add("TOP10 카드 90개 이상", html.count("top-case-item") >= 90, f"top-case-item={html.count('top-case-item')}")
    add("TOP10 top-case-card 3개 기준", html.count("top-case-card") == 3, f"top-case-card={html.count('top-case-card')}")
    add("TOP10 top-case-list 9개 기준", html.count("top-case-list") == 9, f"top-case-list={html.count('top-case-list')}")
    add("TOP10 템플릿 카드 내부구조", all(x in html for x in ["case-rank", "case-body", "case-value", "case-year"]))
    add("월세·종합 후반부 출력 섹션 포함", all(x in html for x in ["Ⅴ. 월세 주택 거래 현황 분석", "8. 아파트 월세 당월 직전 신규계약 대비 TOP 거래사례", "Ⅵ. 종합 시황 서술 보고서"]))
    add("표지 KPI print 세로형 CSS 포함", "grid-template-columns:1fr" in html and "v5.8.38 PDF 출력 안정화 보정" in html)
    add("상위 TOP10 중립형 적용", "case-top" in html)
    add("v5.8.38 차트 중복 방지", "평균 가격변동률" in html and "상승·하락 비율" in html)
    if weekly_mode or officetel_mode:
        add("비아파트 물건유형별 차트 미출력", "비아파트 물건유형별 최근 13개월 매매 거래 추이" not in html)
    else:
        add("비아파트 물건유형별 차트", "비아파트 물건유형별 최근 13개월 매매 거래 추이(연립·다세대·단독·다가구)" in html and "연립·다세대" in html and "단독·다가구" in html)
    add("인구·세대수 차트 보정", "population-mix-chart" in html)
    if weekly_mode or officetel_mode:
        add("비아파트 4개 유형별 차트 미출력", "비아파트 물건유형별 최근 13개월 매매 거래 추이" not in html)
    else:
        add("비아파트 4개 유형별 차트", all(x in html for x in ["연립", "다세대", "단독", "다가구"]) and "비아파트 물건유형별 최근 13개월 매매 거래 추이" in html)
    add("인구이동 통합표 포함", all(x in html for x in ["전입", "전출", "순이동"]))
    add("연령대 누적순이동 가로 막대차트 삭제", "최근 13개월 누적순이동 연령별 비교" not in html)
    add("연령별 차트 컬러 일치", all(x in html for x in ["#5BC0EB", "#3A86FF", "#8338EC", "#2A9D8F", "#F4A261", "#E76F51", "#577590", "#D62828"]))
    add("인구이동 미공표 표기", "not-published" in html and "이동자수 최신 공표월" in html)
    add("인구통합표 가로잘림 방지", "population-migration-table" in html and "주요 순유출·순유입 연령대" in html)
    add("월간 가격변동 차트 꺾은선 제거", "평균 변동률 꺾은선은 제거" in html)
    add("월간 시장강도 차트 상승·하락 꺾은선", "상승·하락 비율만 꺾은선" in html)
    add("13개월 시계열 공표통계 컬럼", "공표통계" in html)
    add("상승보합하락 표 이모지", all(x in html for x in ["🔴 상승 비율", "⚪ 보합 비율", "🔵 하락 비율"]))
    add("상승보합하락 표 컬러 class", all(x in html for x in ["dir-up", "dir-flat", "dir-down"]))
    add("연령대별 인구 전월 대비 증감률 표", "최근 13개월 연령대별 주민등록인구 시계열 및 전월 대비 증감률" in html and "age-pop-trend-table" in html)
    add("차트 내부 제목 비활성화", "\"title\": {\"display\": false}" in html or '"title": {"display": false}' in html)
    add("가격지수 별도 섹션 미출력", "주택가격지수</h3>" not in html)
    add("공급 물건유형별 표 포함", "물건유형별 공급 파이프라인 시계열" in html)
    if officetel_mode:
        add("오피스텔 비아파트·분양권 섹션 미출력", all(x not in text for x in ["비아파트", "연립·다세대 매매", "단독·다가구 매매", "분양권전매"]))
    else:
        add(f"총 table {expected_table_count}개 기준", html.count("<table") == expected_table_count, f"table={html.count('<table')}")
    if weekly_mode or officetel_mode:
        add("비아파트 종합서술 미출력", all(x not in text for x in ["연립·다세대 매매", "단독·다가구 매매", "연립·다세대 전세", "단독·다가구 전세", "연립·다세대 월세", "단독·다가구 월세"]))
    else:
        add("비아파트 보강표 6개 포함", all(x in html for x in ["연립·다세대 매매", "단독·다가구 매매", "연립·다세대 전세", "단독·다가구 전세", "연립·다세대 월세", "단독·다가구 월세"]))
    add("conversion-card 실제 출력", "conversion-card" in html)
    if not officetel_mode:
        add(f"총 table {expected_table_count}개 기준", html.count("<table") == expected_table_count, f"table={html.count('<table')}")

    if external_buyer_mode:
        add("외지인 매수 13개월·당월·6개월 누적표", all(x in html for x in ["최근 13개월 주거용 매매 및 외지인 거래 추이", "외지인 매수자 시군구별 건수·비율", "최근 6개월 외지인 매수자 시군구별 누적"]))

    add("table-wrap 35개 이상", html.count("table-wrap") >= 35, f"table-wrap={html.count('table-wrap')}")
    add("연령대 순이동 통합 후 canvas 55개 이상", html.count("<canvas") >= 55, f"canvas={html.count('<canvas')}")
    add("template_v58 주요 클래스 보존", all(cls in html for cls in ["survey-donut-card","conversion-card","supply-card","age-move-card","mini-chart-card","top-case-item","table-wrap"]), "template_v58 key classes")
    add("template_v58 의미색상 JS 포함", "semanticChartColors" in html and "applySemanticColors" in html and "prepareChartConfig" in html)
    add("template_v58 지연 렌더링 적용", "registerLazyChart" in html and "__lazyChartQueue" in html)
    canvas_ids = re.findall(r"<canvas[^>]+id=[\'\"]([^\'\"]+)[\'\"]", html)
    extra_canvas = [cid for cid in canvas_ids if cid not in TEMPLATE_V58_REQUIRED_CANVAS_IDS]
    add("템플릿 외 차트 없음", len(extra_canvas) == 0, f"extra={extra_canvas[:5]}")
    add("숨김 차트 bank 미생성", "template-v58-canvas-bank" not in html and "template-v58-class-bank" not in html)
    canvas_id_list = re.findall(r"<canvas[^>]+id=['\"]([^'\"]+)['\"]", html)
    add("canvas id 중복 없음", len(canvas_id_list) == len(set(canvas_id_list)), f"duplicates={[k for k,v in Counter(canvas_id_list).items() if v>1] if 'Counter' in globals() else ''}")
    add("template_v58 canvas 정확히 58개", len(canvas_id_list) == 58, f"canvas={len(canvas_id_list)}")
    add("template_v58 chart-box 정확히 36개", html.count("class='chart-box") + html.count('class="chart-box') == 36, f"chart-box={html.count('chart-box')}")
    add("chart3·chart11 누락 방지", "id='chart3'" in html and "id='chart11'" in html)
    add("TOP10 카드 90개 이상", html.count("top-case-item") >= 90, f"top-case-item={html.count('top-case-item')}")
    add("TOP10 top-case-card 3개 기준", html.count("top-case-card") == 3, f"top-case-card={html.count('top-case-card')}")
    add("TOP10 top-case-list 9개 기준", html.count("top-case-list") == 9, f"top-case-list={html.count('top-case-list')}")
    add("TOP10 템플릿 카드 내부구조", all(x in html for x in ["case-rank", "case-body", "case-value", "case-year"]))
    add("월세·종합 후반부 출력 섹션 포함", all(x in html for x in ["Ⅴ. 월세 주택 거래 현황 분석", "8. 아파트 월세 당월 직전 신규계약 대비 TOP 거래사례", "Ⅵ. 종합 시황 서술 보고서"]))
    add("표지 KPI print 세로형 CSS 포함", "grid-template-columns:1fr" in html and "v5.8.38 PDF 출력 안정화 보정" in html)
    add("상위 TOP10 중립형 적용", "case-top" in html)
    add("v5.8.38 차트 중복 방지", "평균 가격변동률" in html and "상승·하락 비율" in html)
    if weekly_mode or officetel_mode:
        add("비아파트 물건유형별 차트 미출력", "비아파트 물건유형별 최근 13개월 매매 거래 추이" not in html)
    else:
        add("비아파트 물건유형별 차트", "비아파트 물건유형별 최근 13개월 매매 거래 추이(연립·다세대·단독·다가구)" in html and "연립·다세대" in html and "단독·다가구" in html)
    add("인구·세대수 차트 보정", "population-mix-chart" in html)
    if weekly_mode or officetel_mode:
        add("비아파트 4개 유형별 차트 미출력", "비아파트 물건유형별 최근 13개월 매매 거래 추이" not in html)
    else:
        add("비아파트 4개 유형별 차트", all(x in html for x in ["연립", "다세대", "단독", "다가구"]) and "비아파트 물건유형별 최근 13개월 매매 거래 추이" in html)
    add("인구이동 통합표 포함", all(x in html for x in ["전입", "전출", "순이동"]))
    add("연령대 누적순이동 가로 막대차트 삭제", "최근 13개월 누적순이동 연령별 비교" not in html)
    add("연령별 차트 컬러 일치", all(x in html for x in ["#5BC0EB", "#3A86FF", "#8338EC", "#2A9D8F", "#F4A261", "#E76F51", "#577590", "#D62828"]))
    add("인구이동 미공표 표기", "not-published" in html and "이동자수 최신 공표월" in html)
    add("인구통합표 가로잘림 방지", "population-migration-table" in html and "주요 순유출·순유입 연령대" in html)
    add("월간 가격변동 차트 꺾은선 제거", "평균 변동률 꺾은선은 제거" in html)
    add("월간 시장강도 차트 상승·하락 꺾은선", "상승·하락 비율만 꺾은선" in html)
    add("13개월 시계열 공표통계 컬럼", "공표통계" in html)
    add("상승보합하락 표 이모지", all(x in html for x in ["🔴 상승 비율", "⚪ 보합 비율", "🔵 하락 비율"]))
    add("상승보합하락 표 컬러 class", all(x in html for x in ["dir-up", "dir-flat", "dir-down"]))
    add("연령대별 인구 전월 대비 증감률 표", "최근 13개월 연령대별 주민등록인구 시계열 및 전월 대비 증감률" in html and "age-pop-trend-table" in html)
    add("차트 내부 제목 비활성화", "\"title\": {\"display\": false}" in html or '"title": {"display": false}' in html)
    add("가격지수 별도 섹션 미출력", "주택가격지수</h3>" not in html)
    add("공급 물건유형별 표 포함", "물건유형별 공급 파이프라인 시계열" in html)
    if not officetel_mode:
        add(f"총 table {expected_table_count}개 기준", html.count("<table") == expected_table_count, f"table={html.count('<table')}")
    if weekly_mode or officetel_mode:
        add("비아파트 종합서술 미출력", all(x not in text for x in ["연립·다세대 매매", "단독·다가구 매매", "연립·다세대 전세", "단독·다가구 전세", "연립·다세대 월세", "단독·다가구 월세"]))
    else:
        add("비아파트 보강표 6개 포함", all(x in html for x in ["연립·다세대 매매", "단독·다가구 매매", "연립·다세대 전세", "단독·다가구 전세", "연립·다세대 월세", "단독·다가구 월세"]))
    add("conversion-card 실제 출력", "conversion-card" in html)
    if not officetel_mode:
        add(f"총 table {expected_table_count}개 기준", html.count("<table") == expected_table_count, f"table={html.count('<table')}")

    add("table-wrap 35개 이상", html.count("table-wrap") >= 35, f"table-wrap={html.count('table-wrap')}")
    add("차트 자료 없음 문구 최소화", html.count("차트 자료 없음") <= 3, f"차트 자료 없음={html.count('차트 자료 없음')}")
    add("template_v58 핵심 클래스 포함", all(cls in html for cls in ["hero", "kpis", "chapter", "card", "chart-box", "note"]))
    # 공급 물건유형별 분리 차트(15개)를 표 2로 대체해 기준을 20개로 낮춤
    add("실제 canvas 차트 20개 이상 포함", html.count("<canvas") >= 20, f"canvas={html.count('<canvas')}")
    add("영문 내부 컬럼명 최소화", "trade_count" not in text and "compare_count" not in text and "population" not in text)
    # 20개 자료 사용 검수
    used = data.get("input_files", {})
    for bucket in ["population", "migration", "rates", "unsold", "supply", "price_index_sale", "price_index_jeonse", "price_index_rent", "weekly_apt_sale_change", "weekly_apt_jeonse_change", "apt_trade", "apt_rent", "non_apt_sale", "non_apt_rent", "presale"]:
        optional_bucket = (
            (weekly_mode and bucket in ["price_index_rent", "non_apt_sale", "non_apt_rent"])
            or (not weekly_mode and bucket in ["weekly_apt_sale_change", "weekly_apt_jeonse_change"])
            or (officetel_mode and bucket in ["non_apt_sale", "non_apt_rent", "presale", "weekly_apt_sale_change", "weekly_apt_jeonse_change"])
        )
        add(f"파일 분류 확인: {bucket}", optional_bucket or len(used.get(bucket, [])) > 0, ", ".join(used.get(bucket, [])))
    conv = data.get("conversion_rate", {})
    add("전월세전환율 산출 여부", bool(conv.get("available")), f"상태={conv.get('status')}, 매칭={conv.get('matched_count')}, 적용={conv.get('applied_rate')}")
    add("환산월세 산식 ÷100 포함", "÷100" in str(conv.get("converted_rent_formula", "")), conv.get("converted_rent_formula", ""))
    for w in data.get("data_quality_warnings", []):
        add("데이터 품질 경고", False, w)
    return rows



def assert_print_completeness(html: str) -> List[str]:
    """HTML/PDF 출력 전 필수 후반부 섹션 누락 방지 검수."""
    required = [
        "Ⅴ. 월세 주택 거래 현황 분석",
        "8. 아파트 월세 당월 직전 신규계약 대비 TOP 거래사례",
        "Ⅵ. 종합 시황 서술 보고서",
    ]
    issues = []
    for item in required:
        if item not in html:
            issues.append(f"필수 출력 섹션 누락: {item}")
    if "template-v58-canvas-bank" in html or "template-v58-class-bank" in html:
        issues.append("숨김 보존 bank가 남아 있음")
    if html.count("<canvas") < 55:
        issues.append(f"canvas 수 부족: {html.count('<canvas')}개")
    if html.count("top-case-card") != 3:
        issues.append(f"TOP10 top-case-card 수 불일치: {html.count('top-case-card')}개")
    if html.count("top-case-item") < 90:
        issues.append(f"TOP10 카드 수 부족: {html.count('top-case-item')}개")
    return issues

def _is_officetel_report_data(data: Dict[str, Any]) -> bool:
    meta = data.get("meta", {}) if isinstance(data, dict) else {}
    return bool(meta.get("is_officetel")) or str(meta.get("report_mode") or "") == "officetel"


def _sanitize_non_officetel_text(text: str) -> str:
    if not isinstance(text, str):
        return text
    return text.replace("비오피스텔", "비아파트").replace("오피스텔", "아파트")


def _sanitize_non_officetel_commentary(data: Dict[str, Any], commentary: Dict[str, str]) -> Dict[str, str]:
    if _is_officetel_report_data(data):
        return commentary
    return {
        k: _sanitize_non_officetel_text(v) if isinstance(v, str) else v
        for k, v in (commentary or {}).items()
    }


def write_outputs(data: Dict[str, Any], commentary: Dict[str, str], suffix: str = "", output_dir: Optional[Path] = None) -> Dict[str, Any]:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_region = re.sub(r"[^가-힣A-Za-z0-9_]+", "_", data["meta"]["region"])
    source_key = "api" if data.get("meta", {}).get("input_source") == "api" else ("excel" if data.get("meta", {}).get("input_source") == "excel_upload" else "auto")
    stem = f"report_{source_key}_{safe_region}_{data['meta']['base_month'].replace('-','')}_{ts}{suffix}"
    output_dir = Path(output_dir) if output_dir else OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)
    commentary = _sanitize_non_officetel_commentary(data, commentary)
    html0 = render_html(data, commentary, [])
    if not _is_officetel_report_data(data):
        html0 = _sanitize_non_officetel_text(html0)
    validation = validate_report(html0, data, commentary)
    html = render_html(data, commentary, validation)
    if not _is_officetel_report_data(data):
        html = _sanitize_non_officetel_text(html)
    data_path = output_dir / f"data_pack_{stem}.json"
    comm_path = output_dir / f"commentary_pack_{stem}.json"
    html_path = output_dir / f"{stem}.html"
    val_path = output_dir / f"validation_{stem}.xlsx"
    save_json(data, data_path)
    save_json(commentary, comm_path)
    html_path.write_text(html, encoding="utf-8")
    pd.DataFrame(validation).to_excel(val_path, index=False)
    return {"data_pack": data, "commentary": commentary, "data_pack_path": data_path, "commentary_pack_path": comm_path, "html_path": html_path, "validation_excel_path": val_path, "validation_rows": validation, "validation_df": pd.DataFrame(validation)}


def save_prompt(data: Dict[str, Any], prompt: str, prompt_dir: Optional[Path] = None) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_region = re.sub(r"[^가-힣A-Za-z0-9_]+", "_", data["meta"]["region"])
    source_key = "api" if data.get("meta", {}).get("input_source") == "api" else ("excel" if data.get("meta", {}).get("input_source") == "excel_upload" else "auto")
    prompt_dir = Path(prompt_dir) if prompt_dir else PROMPT_DIR
    prompt_dir.mkdir(parents=True, exist_ok=True)
    path = prompt_dir / f"llm_prompt_{source_key}_{safe_region}_{data['meta']['base_month'].replace('-','')}_{ts}.txt"
    path.write_text(prompt, encoding="utf-8")
    return path


def run_pipeline(
    paths: List[Path],
    region: str,
    base_month: str,
    version: str,
    input_source: str = "auto",
    output_dir: Optional[Path] = None,
    prompt_dir: Optional[Path] = None,
    designated_sale_files: Optional[List[Path]] = None,
    designated_rent_files: Optional[List[Path]] = None,
) -> Dict[str, Any]:
    progress("보고서 생성 파이프라인 시작", 0, 100)
    validate_input_source_files(
        paths,
        input_source,
        designated_sale_files=designated_sale_files,
        designated_rent_files=designated_rent_files,
    )
    data = build_data_pack(
        paths,
        region,
        base_month,
        version,
        input_source=input_source,
        designated_sale_files=designated_sale_files,
        designated_rent_files=designated_rent_files,
    )
    progress("규칙 기반 1차 해설 생성 중", 85, 100)
    commentary = base_commentary(data)
    progress("HTML 및 검수표 생성 중", 92, 100)
    result = write_outputs(data, commentary, output_dir=output_dir)
    prompt = build_llm_prompt(data)
    if not _is_officetel_report_data(data):
        prompt = _sanitize_non_officetel_text(prompt)
    prompt_path = save_prompt(data, prompt, prompt_dir=prompt_dir)
    result["prompt_text"], result["prompt_path"] = prompt, prompt_path
    progress("보고서 생성 완료", 100, 100)
    return result


def regenerate_with_commentary(data: Dict[str, Any], commentary: Dict[str, str], output_dir: Optional[Path] = None) -> Dict[str, Any]:
    return write_outputs(data, commentary, suffix="_llm", output_dir=output_dir)


# =========================================================
# 9. Tkinter UI
# =========================================================

class JsonPasteDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("LLM 결과 JSON 붙여넣기")
        self.geometry("920x620")
        self.result_text = None
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="제미나이/GPT가 출력한 JSON 결과를 붙여넣으세요.", font=("맑은 고딕", 10, "bold")).pack(anchor="w")
        self.text = tk.Text(frame, font=("Consolas", 10), wrap="word")
        self.text.pack(fill="both", expand=True, pady=8)
        btn = ttk.Frame(frame); btn.pack(fill="x")
        ttk.Button(btn, text="클립보드 붙여넣기", command=self.paste_clipboard).pack(side="left")
        ttk.Button(btn, text="JSON/TXT 불러오기", command=self.load_file).pack(side="left", padx=6)
        ttk.Button(btn, text="적용", command=self.apply).pack(side="right")
        ttk.Button(btn, text="취소", command=self.destroy).pack(side="right", padx=6)
        self.transient(master); self.grab_set()
    def paste_clipboard(self):
        try:
            self.text.delete("1.0", "end")
            self.text.insert("1.0", self.clipboard_get())
        except Exception as e:
            messagebox.showerror("오류", str(e))
    def load_file(self):
        p = filedialog.askopenfilename(filetypes=[("JSON/TXT", "*.json *.txt"), ("All", "*.*")])
        if p:
            self.text.delete("1.0", "end")
            self.text.insert("1.0", Path(p).read_text(encoding="utf-8"))
    def apply(self):
        self.result_text = self.text.get("1.0", "end").strip()
        self.destroy()


class App(tk.Tk):
    def __init__(self, input_source: str = "auto", output_root: Optional[Path] = None):
        super().__init__()
        self.input_source = str(input_source or "auto").strip().lower()
        self.input_source_label = INPUT_SOURCE_LABELS.get(self.input_source, INPUT_SOURCE_LABELS["auto"])
        self.output_root = Path(output_root) if output_root else BASE_DIR
        self.report_output_dir = self.output_root / "보고서"
        self.prompt_output_dir = self.output_root / "LLM_프롬프트"
        self.log_output_dir = self.output_root / "로그"
        for directory in [self.report_output_dir, self.prompt_output_dir, self.log_output_dir]:
            directory.mkdir(parents=True, exist_ok=True)
        self.title(f"{self.input_source_label} 주택동향보고서 생성기")
        self.geometry("1280x860")
        self.minsize(1120, 720)
        self.selected_files: List[Path] = []
        self.last_result = None
        self.current_data = None
        self.current_commentary = None
        self.current_prompt = ""
        self.current_prompt_path = None
        self._style()
        self._ui()

    def _style(self):
        style = ttk.Style()
        try: style.theme_use("clam")
        except Exception: pass
        style.configure("Header.TLabel", font=("맑은 고딕", 16, "bold"), foreground="#1f4e79")
        style.configure("Small.TLabel", font=("맑은 고딕", 9), foreground="#666")
        style.configure("Primary.TButton", font=("맑은 고딕", 10, "bold"), padding=8)

    def _ui(self):
        root = ttk.Frame(self, padding=14); root.pack(fill="both", expand=True)
        ttk.Label(root, text=f"{self.input_source_label} 주택동향보고서 생성기", style="Header.TLabel").pack(anchor="w")
        ttk.Label(root, text=f"입력 방식과 산출 폴더를 독립 관리 · {self.report_output_dir}", style="Small.TLabel").pack(anchor="w", pady=(2,12))
        main = ttk.PanedWindow(root, orient="horizontal"); main.pack(fill="both", expand=True)
        left = ttk.Frame(main, padding=8); right = ttk.Frame(main, padding=8)
        main.add(left, weight=1); main.add(right, weight=2)

        settings = ttk.LabelFrame(left, text="1. 기본 설정", padding=10); settings.pack(fill="x")
        self.region_var = tk.StringVar(value="광주광역시 남구")
        self.month_var = tk.StringVar(value="2026-04")
        self.version_var = tk.StringVar(value=APP_VERSION)
        for i, (label, var) in enumerate([("조사지역", self.region_var), ("기준월", self.month_var), ("버전", self.version_var)]):
            ttk.Label(settings, text=label).grid(row=i, column=0, sticky="w", pady=4)
            ttk.Entry(settings, textvariable=var).grid(row=i, column=1, sticky="ew", pady=4)
        ttk.Button(settings, text="파일명에서 지역·기준월 자동 추정", command=self.infer_region_month).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8,2))
        settings.columnconfigure(1, weight=1)

        fb = ttk.LabelFrame(left, text="2. 원자료 파일", padding=10); fb.pack(fill="both", expand=True, pady=(12,0))
        b = ttk.Frame(fb); b.pack(fill="x")
        ttk.Button(b, text="파일 추가", command=self.add_files).pack(side="left")
        ttk.Button(b, text="경로목록 추가", command=self.add_paths).pack(side="left", padx=5)
        ttk.Button(b, text="선택 삭제", command=self.remove_files).pack(side="left")
        ttk.Button(b, text="전체 삭제", command=self.clear_files).pack(side="left", padx=5)
        ttk.Button(b, text="데이터 진단", command=self.audit).pack(side="left")
        self.file_list = tk.Listbox(fb, height=13, selectmode="extended", font=("맑은 고딕", 9))
        self.file_list.pack(fill="both", expand=True, pady=(8,0))

        rb = ttk.LabelFrame(left, text="3. 실행", padding=10); rb.pack(fill="x", pady=(12,0))
        self.run_btn = ttk.Button(rb, text="① HTML 최종보고서 + LLM 프롬프트 생성", style="Primary.TButton", command=self.run_clicked)
        self.run_btn.pack(fill="x")
        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(rb, maximum=100, variable=self.progress_var).pack(fill="x", pady=(8,0))
        self.progress_label = ttk.Label(rb, text="대기 중", style="Small.TLabel"); self.progress_label.pack(anchor="w", pady=(4,0))

        lb = ttk.LabelFrame(left, text="4. LLM 해설", padding=10); lb.pack(fill="x", pady=(12,0))
        ttk.Button(lb, text="프롬프트 TXT 열기", command=self.open_prompt).pack(fill="x", pady=2)
        ttk.Button(lb, text="프롬프트 클립보드 복사", command=self.copy_prompt).pack(fill="x", pady=2)
        ttk.Button(lb, text="② LLM 해설 JSON 붙여넣기", command=self.paste_json).pack(fill="x", pady=2)
        ttk.Button(lb, text="③ LLM 해설 적용 후 HTML 재생성", command=self.regen).pack(fill="x", pady=2)

        ob = ttk.LabelFrame(left, text="5. 결과", padding=10); ob.pack(fill="x", pady=(12,0))
        ttk.Button(ob, text="HTML 열기", command=self.open_html).pack(fill="x", pady=2)
        ttk.Button(ob, text="검수표 열기", command=self.open_validation).pack(fill="x", pady=2)
        ttk.Button(ob, text="출력 폴더 열기", command=self.open_output).pack(fill="x", pady=2)

        cf = ttk.LabelFrame(right, text="파일 자동 분류", padding=10); cf.pack(fill="x")
        self.classified_text = tk.Text(cf, height=8, font=("Consolas",9)); self.classified_text.pack(fill="x")
        pf = ttk.LabelFrame(right, text="진행상황", padding=10); pf.pack(fill="x", pady=(12,0))
        self.progress_tree = ttk.Treeview(pf, columns=("time","msg"), show="headings", height=7)
        self.progress_tree.heading("time", text="시간"); self.progress_tree.heading("msg", text="상태")
        self.progress_tree.column("time", width=80); self.progress_tree.column("msg", width=700)
        self.progress_tree.pack(fill="x")
        qa = ttk.LabelFrame(right, text="LLM 해설 적용 바로가기", padding=10)
        qa.pack(fill="x", pady=(12,0))
        ttk.Label(
            qa,
            text="① 전체 분석 실행 후, ChatGPT가 출력한 JSON을 아래 버튼으로 붙여넣으면 HTML을 바로 재생성합니다.",
            style="Small.TLabel"
        ).pack(anchor="w", pady=(0,6))
        ttk.Button(
            qa,
            text="② LLM 해설 JSON 붙여넣고 HTML 바로 재생성",
            style="Primary.TButton",
            command=self.paste_json_and_regen
        ).pack(fill="x")
        ttk.Button(
            qa,
            text="프롬프트 클립보드 복사",
            command=self.copy_prompt
        ).pack(fill="x", pady=(6,0))

        pr = ttk.LabelFrame(right, text="LLM 프롬프트 미리보기", padding=10); pr.pack(fill="both", expand=True, pady=(12,0))
        self.prompt_text = tk.Text(pr, font=("Consolas", 9), wrap="word", height=12); self.prompt_text.pack(fill="both", expand=True)
        vf = ttk.LabelFrame(right, text="검수 결과", padding=10); vf.pack(fill="both", expand=True, pady=(12,0))
        self.validation_tree = ttk.Treeview(vf, columns=("item","status","note"), show="headings", height=8)
        for col, txt, w in [("item","검수항목",420), ("status","결과",90), ("note","비고",420)]:
            self.validation_tree.heading(col, text=txt); self.validation_tree.column(col, width=w)
        self.validation_tree.pack(fill="both", expand=True)

        self.log("프로그램 시작")
        self.refresh_classification()

    def ui_progress(self, msg, cur=None, total=None):
        self.after(0, lambda: self._ui_progress(msg, cur, total))
    def _ui_progress(self, msg, cur=None, total=None):
        if cur is not None and total:
            pct = max(0, min(100, float(cur)/float(total)*100))
            self.progress_var.set(pct); self.progress_label.config(text=f"{msg} ({cur}/{total}, {pct:.1f}%)")
        else:
            self.progress_label.config(text=msg)
        self.progress_tree.insert("", "end", values=(datetime.now().strftime("%H:%M:%S"), msg))
        self.progress_tree.see(self.progress_tree.get_children()[-1])
        self.log(msg)
    def log(self, msg):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    def refresh_files(self):
        self.file_list.delete(0, "end")
        for p in self.selected_files: self.file_list.insert("end", str(p))
        self.refresh_classification()
    def refresh_classification(self):
        c = classify_files(self.selected_files, self.input_source)
        self.classified_text.delete("1.0","end")
        self.classified_text.insert("1.0", "\n".join([f"[{k}]\n" + "\n".join([f"  - {p.name}" for p in v]) if v else f"[{k}]\n  - 없음" for k,v in c.items()]))
    def add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        ex = {str(p) for p in self.selected_files}
        for s in paths:
            if s not in ex: self.selected_files.append(Path(s))
        self.refresh_files()
    def add_paths(self):
        dlg = tk.Toplevel(self); dlg.title("파일 경로 목록 추가"); dlg.geometry("900x520")
        f = ttk.Frame(dlg, padding=12); f.pack(fill="both", expand=True)
        txt = tk.Text(f, font=("Consolas", 10)); txt.pack(fill="both", expand=True)
        def apply():
            ex={str(p) for p in self.selected_files}; missing=[]
            for line in txt.get("1.0","end").splitlines():
                s=line.strip().strip('"').strip("'")
                if not s: continue
                p=Path(s)
                if p.exists() and str(p) not in ex:
                    self.selected_files.append(p); ex.add(str(p))
                elif not p.exists(): missing.append(s)
            self.refresh_files(); dlg.destroy()
            if missing: messagebox.showwarning("일부 경로 없음", f"{len(missing)}개 경로가 없습니다.")
        ttk.Button(f, text="추가", command=apply).pack(side="right", pady=8)
    def remove_files(self):
        sel=set(self.file_list.curselection())
        self.selected_files=[p for i,p in enumerate(self.selected_files) if i not in sel]
        self.refresh_files()
    def clear_files(self):
        self.selected_files=[]; self.refresh_files()
    def infer_region_month(self):
        inf = infer_region_month_from_filenames(self.selected_files)
        if inf.get("region"): self.region_var.set(inf["region"])
        if inf.get("base_month"): self.month_var.set(inf["base_month"])
        messagebox.showinfo("추정 완료", f"지역: {inf.get('region')}\n기준월: {inf.get('base_month')}")
    def audit(self):
        if not self.selected_files: return
        set_progress_callback(self.ui_progress)
        try:
            paths = generate_data_audit_report(self.selected_files)
            os.startfile(str(paths["xlsx"]))
        finally:
            set_progress_callback(None)
    def run_clicked(self):
        if not self.selected_files:
            messagebox.showwarning("파일 필요", "원자료 파일을 추가하세요."); return
        region=self.region_var.get().strip(); month=self.month_var.get().strip()
        if not re.fullmatch(r"\d{4}-\d{2}", month):
            messagebox.showwarning("기준월 오류", "기준월은 2026-04 형식입니다."); return
        self.run_btn.config(state="disabled"); self.progress_var.set(0)
        set_progress_callback(self.ui_progress)
        threading.Thread(target=self._run_thread, args=(region, month, self.version_var.get().strip()), daemon=True).start()
    def _run_thread(self, region, month, version):
        try:
            result = run_pipeline(
                self.selected_files,
                region,
                month,
                version,
                input_source=self.input_source,
                output_dir=self.report_output_dir,
                prompt_dir=self.prompt_output_dir,
            )
            self.after(0, lambda res=result: self._success(res))
        except Exception as e:
            tb=traceback.format_exc(); path=self.log_output_dir/f"error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"; path.write_text(tb, encoding="utf-8")
            self.after(0, lambda err=e, p=path: self._error(err,p))
    def _success(self, result):
        set_progress_callback(None); self.run_btn.config(state="normal"); self.progress_var.set(100)
        self.last_result=result; self.current_data=result["data_pack"]; self.current_commentary=result["commentary"]; self.current_prompt=result["prompt_text"]; self.current_prompt_path=result["prompt_path"]
        self.prompt_text.delete("1.0","end"); self.prompt_text.insert("1.0", self.current_prompt)
        self.refresh_validation(result["validation_rows"])
        messagebox.showinfo("완료", "HTML 최종보고서와 LLM 프롬프트가 생성되었습니다.")
    def _error(self, e, path):
        set_progress_callback(None); self.run_btn.config(state="normal")
        messagebox.showerror("오류", f"{e}\n\n로그: {path}")
    def refresh_validation(self, rows):
        for i in self.validation_tree.get_children(): self.validation_tree.delete(i)
        for r in rows: self.validation_tree.insert("", "end", values=(r.get("검수항목"), r.get("결과"), r.get("비고")))
    def open_prompt(self):
        if self.current_prompt_path and self.current_prompt_path.exists(): os.startfile(str(self.current_prompt_path))
    def copy_prompt(self):
        if not self.current_prompt: return
        self.clipboard_clear(); self.clipboard_append(self.current_prompt); self.update()
        messagebox.showinfo("복사 완료", "프롬프트가 클립보드에 복사되었습니다.")
    def paste_json(self):
        if not self.current_data:
            messagebox.showinfo("data_pack 없음", "먼저 전체 분석을 실행하세요."); return
        dlg=JsonPasteDialog(self); self.wait_window(dlg)
        if not dlg.result_text: return
        try:
            new=extract_json_from_text(dlg.result_text)
            self.current_commentary=merge_commentary(base_commentary(self.current_data), new, self.current_data)
            messagebox.showinfo("적용 준비", summarize_llm_application(self.current_commentary, self.current_data) + "\n\n재생성 버튼을 누르세요.")
        except Exception as e:
            messagebox.showerror("JSON 오류", str(e))

    def paste_json_and_regen(self):
        """LLM 해설 JSON을 붙여넣고 바로 HTML 재생성까지 수행."""
        if not self.current_data:
            messagebox.showinfo("data_pack 없음", "먼저 [① HTML 최종보고서 + LLM 프롬프트 생성]을 실행하세요.")
            return
        dlg = JsonPasteDialog(self)
        self.wait_window(dlg)
        if not dlg.result_text:
            return
        try:
            new = extract_json_from_text(dlg.result_text)
            self.current_commentary = merge_commentary(base_commentary(self.current_data), new, self.current_data)
            result = regenerate_with_commentary(self.current_data, self.current_commentary, output_dir=self.report_output_dir)
            self.last_result = result
            self.refresh_validation(result["validation_rows"])
            messagebox.showinfo(
                "HTML 재생성 완료",
                summarize_llm_application(self.current_commentary, self.current_data) + "\n\n"
                "LLM 해설 JSON을 적용하여 HTML을 재생성했습니다.\n\n"
                f"{result['html_path']}"
            )
            try:
                webbrowser.open(result["html_path"].as_uri())
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("JSON 적용 오류", str(e))

    def regen(self):
        if not self.current_data or not self.current_commentary: return
        result=regenerate_with_commentary(self.current_data, self.current_commentary, output_dir=self.report_output_dir)
        self.last_result=result; self.refresh_validation(result["validation_rows"])
        messagebox.showinfo("재생성 완료", f"HTML: {result['html_path']}")
    def open_html(self):
        if self.last_result: webbrowser.open(self.last_result["html_path"].as_uri())
    def open_validation(self):
        if self.last_result: os.startfile(str(self.last_result["validation_excel_path"]))
    def open_output(self):
        os.startfile(str(self.report_output_dir))


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
월간 주택동향보고서 HTML 최종보고서 생성기 v5.8.38.1.1
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
from collections import Counter
import sys
import traceback
import threading
import webbrowser
from datetime import datetime
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


APP_VERSION = "windows_ui_html_template_v58_reviewfix_v5_8_38"


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
    "apt_name": ["단지명", "아파트명", "건물명", "주택명"],
    "approval_year": ["사용승인년도", "건축년도", "준공년도", "건축연도"],
    "households": ["세대수", "주민등록세대수"],
    "exclusive_area": ["전용면적", "전용면적㎡", "전용면적(㎡)", "건물면적", "면적"],
    "price": ["거래금액", "거래금액만원", "거래금액(만원)", "매매가", "거래가액"],
    "deposit": ["보증금", "보증금액", "보증금(만원)", "임대보증금"],
    "monthly_rent": ["월세", "월세금액", "월세금(만원)", "월차임"],
    "floor": ["층", "해당층"],
    "contract_date": ["계약일", "계약일자", "계약년월일"],
    "contract_type": ["계약구분", "갱신요구권사용", "신규갱신", "갱신여부"],
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
    "supply_commentary",
    "price_index_commentary",
    "conversion_rate_commentary",
    "apt_trade_commentary",
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
    "rowhouse_multiplex_sale_summary",
    "detached_multifamily_sale_summary",
    "jeonse_market_summary",
    "apt_jeonse_summary",
    "rowhouse_multiplex_jeonse_summary",
    "detached_multifamily_jeonse_summary",
    "rent_market_summary",
    "apt_rent_summary",
    "rowhouse_multiplex_rent_summary",
    "detached_multifamily_rent_summary",
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


def find_best_month_col(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    priority = ["계약년월", "계약월", "기준월", "년월", "기간", "시점", "PRD_DE"]
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


def read_excel_or_csv(path: Path) -> pd.DataFrame:
    progress(f"파일 읽는 중: {path.name}")
    if path.suffix.lower() == ".csv":
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(path, encoding="cp949")
    else:
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
    for c in df.columns:
        cs = str(c)
        if any(k in cs for k in ["지역", "시도", "시군구", "행정구역", "법정동시군구", "분류1명", "분류2명"]):
            candidate_cols.append(c)
    if not candidate_cols:
        return df
    full = region.replace(" ", "")
    mask = pd.Series(False, index=df.index)
    for c in candidate_cols:
        s = df[c].astype(str).str.replace(" ", "", regex=False)
        mask = mask | s.str.contains(full, na=False)
    if mask.any():
        return df[mask].copy()
    last = tokens[-1]
    mask = pd.Series(False, index=df.index)
    for c in candidate_cols:
        mask = mask | df[c].astype(str).str.contains(last, na=False)
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

def classify_files(paths: List[Path]) -> Dict[str, List[Path]]:
    buckets = {
        "population": [], "migration": [], "rates": [], "unsold": [], "supply": [],
        "price_index_sale": [], "price_index_jeonse": [], "price_index_rent": [],
        "apt_trade": [], "apt_rent": [], "non_apt_sale": [], "non_apt_rent": [],
        "presale": [], "unknown": [],
    }
    for p in paths:
        name = p.name.lower()
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
        elif "매매가격지수" in name:
            buckets["price_index_sale"].append(p)
        elif "전세가격지수" in name:
            buckets["price_index_jeonse"].append(p)
        elif "월세통합가격지수" in name or "월세가격지수" in name:
            buckets["price_index_rent"].append(p)
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
        g = df.groupby("_month")["_value"].max().reset_index()
        for _, r in g.iterrows():
            m = r["_month"]
            if str(m).startswith("20") and pd.notna(r["_value"]):
                monthly.setdefault(m, {})[metric] = float(r["_value"])
    return monthly


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
            if not vcol:
                continue
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            df["_value"] = to_numeric_safe(df[vcol])
            name = p.name
            if "주민등록인구" in name or ("인구" in name and "이동" not in name):
                if item_col:
                    item = df[item_col].astype(str)
                    total_mask = item.str.contains("총인구|계", na=False)
                    sex_mask = item.str.contains("남자인구수|여자인구수|남자|여자", na=False)
                    if total_mask.any():
                        g = df[total_mask].groupby("_month")["_value"].max().reset_index()
                    elif sex_mask.any():
                        g = df[sex_mask].groupby("_month")["_value"].sum().reset_index()
                    else:
                        g = df.groupby("_month")["_value"].sum().reset_index()
                else:
                    g = df.groupby("_month")["_value"].sum().reset_index()
                for _, r in g.iterrows():
                    if str(r["_month"]).startswith("20") and pd.notna(r["_value"]):
                        monthly.setdefault(r["_month"], {})["population"] = int(r["_value"])
            elif "세대" in name:
                g = df.groupby("_month")["_value"].max().reset_index()
                for _, r in g.iterrows():
                    if str(r["_month"]).startswith("20") and pd.notna(r["_value"]):
                        monthly.setdefault(r["_month"], {})["households"] = int(r["_value"])
        if not monthly:
            return result
        rows = [{"month": m, "population": monthly[m].get("population"), "households": monthly[m].get("households")} for m in sorted(monthly)]
        dfm = pd.DataFrame(rows).sort_values("month")
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
        result.update({"available": True, "latest_month": row["month"], "total": None if pd.isna(pop) else int(pop), "mom_change": mom_pop, "households": None if pd.isna(hh) else int(hh), "households_mom": mom_hh, "table_rows": dfm.tail(13).to_dict("records"), "signal": signal})
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
        latest = piv[piv["month"] <= base_month].iloc[-1] if not piv[piv["month"] <= base_month].empty else piv.iloc[-1]
        first = piv.iloc[0]

        # 주민등록인구 파일에서 남자+여자와 총인구 행이 함께 잡혀 2배 산출되는 경우 보정
        try:
            pop_ref = calc_population(files, region, base_month).get("total")
            total_raw = sum(float(latest[c]) for c in order)
            if pop_ref and total_raw > pop_ref * 1.5:
                for c in order:
                    piv[c] = piv[c] / 2
                latest = piv[piv["month"] <= base_month].iloc[-1] if not piv[piv["month"] <= base_month].empty else piv.iloc[-1]
                first = piv.iloc[0]
        except Exception:
            pass

        total_latest = sum(float(latest[c]) for c in order)
        age_rows = []
        for c in order:
            latest_pop = int(latest[c])
            change_13m = int(latest[c] - first[c])
            share = round(latest_pop / total_latest * 100, 1) if total_latest else 0
            note = "청년·소형 임대수요 참고" if c == "20대" else ("고령층 정주수요 참고" if c == "70대 이상" else ("가족형 실수요 참고" if c in ["30대", "40대"] else "주거수요 참고"))
            age_rows.append({"연령대": c, "최신월 인구": latest_pop, "비중": share, "13개월 증감": change_13m, "시장 해석": note})
        result.update({
            "available": True,
            "latest_month": latest["month"],
            "monthly_rows": piv.tail(13).to_dict("records"),
            "age_rows": age_rows,
            "signal": "연령대별 인구 구조 분석 완료"
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

        result.update({"available": True, "latest_month": latest["month"], "table_rows": dfr.tail(13).to_dict("records"), "age_rows": age_rows, "age_monthly_rows": age_monthly_rows, "age_trend_rows": age_trend_rows, "net_migration": int(latest.get("net")) if pd.notna(latest.get("net")) else None, "main_outflow_age": main_out, "signal": "인구 이동 보조지표 확인"})
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
        row = dfm[dfm["month"] <= base_month].iloc[-1] if not dfm[dfm["month"] <= base_month].empty else dfm.iloc[-1]
        signal = "주거 금융부담 지속" if pd.notna(row.get("mortgage_rate")) and row.get("mortgage_rate") >= 4 else "주거 금융부담 점검"
        result.update({"available": True, "latest_month": row["month"], "mortgage_rate": row.get("mortgage_rate"), "jeonse_rate": row.get("jeonse_rate"), "sme_rate": row.get("sme_rate"), "business_rate": row.get("business_rate"), "table_rows": dfm.tail(13).to_dict("records"), "signal": signal})
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
    row = dfm[dfm["month"] <= base_month].iloc[-1] if not dfm[dfm["month"] <= base_month].empty else dfm.iloc[-1]
    signal = "분양시장 부담 점검"
    if pd.notna(row.get("after_completion_unsold")) and row.get("after_completion_unsold") > 0:
        signal = "준공후 미분양 부담 잔존"
    elif pd.notna(row.get("unsold_units")) and row.get("unsold_units") > 0:
        signal = "미분양 부담 잔존"
    result.update({"available": True, "latest_month": row["month"], "unsold_units": None if pd.isna(row.get("unsold_units")) else int(row.get("unsold_units")), "after_completion_unsold": None if pd.isna(row.get("after_completion_unsold")) else int(row.get("after_completion_unsold")), "table_rows": dfm.tail(13).to_dict("records"), "signal": signal})
    return result


def calc_price_index(files: List[Path], region: str, base_month: str, index_type: str) -> Dict[str, Any]:
    result = {"available": False, "latest_month": None, "index_type": index_type, "latest_value": None, "mom_change": None, "table_rows": [], "signal": "가격지수 원자료 미확인"}
    if not files:
        return result
    try:
        rows = []
        for p in files:
            progress(f"{index_type} 가격지수 분석 중: {p.name}")
            df = normalize_columns(read_excel_or_csv(p))
            df = filter_df_by_region(df, region)
            mcol, vcol = find_best_month_col(df), find_value_col(df)
            if not vcol:
                continue
            df["_month"] = df[mcol].map(normalize_month_value) if mcol else base_month
            df["_value"] = to_numeric_safe(df[vcol])
            g = df.groupby("_month")["_value"].max().reset_index()
            for _, r in g.iterrows():
                if str(r["_month"]).startswith("20") and pd.notna(r["_value"]):
                    rows.append({"month": r["_month"], "value": float(r["_value"])})
        if not rows:
            return result
        dfm = pd.DataFrame(rows).groupby("month", as_index=False)["value"].max().sort_values("month")
        row = dfm[dfm["month"] <= base_month].iloc[-1] if not dfm[dfm["month"] <= base_month].empty else dfm.iloc[-1]
        prev = dfm[dfm["month"] < row["month"]].iloc[-1] if not dfm[dfm["month"] < row["month"]].empty else None
        mom = round(row["value"] - prev["value"], 3) if prev is not None else None
        signal = f"{index_type} 가격지수 보합"
        if mom is not None and mom > 0:
            signal = f"{index_type} 가격지수 상승"
        elif mom is not None and mom < 0:
            signal = f"{index_type} 가격지수 하락"
        result.update({"available": True, "latest_month": row["month"], "latest_value": float(row["value"]), "mom_change": mom, "table_rows": dfm.tail(13).to_dict("records"), "signal": signal})
        return result
    except Exception as e:
        result["signal"] = f"가격지수 계산 오류: {e}"
        return result



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
                        by_type_records.append({"month": str(r["year"]), "stage": metric, "type": str(r["type"]), "value": int(round(float(r["value"])))})

        if not monthly:
            return result

        rows = []
        for y in sorted(monthly):
            row = {"month": y}
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
        "apt_name": find_col(df, COLUMN_CANDIDATES["apt_name"]),
        "approval_year": find_col(df, COLUMN_CANDIDATES["approval_year"]),
        "households": find_col(df, COLUMN_CANDIDATES["households"]),
        "area": find_col(df, COLUMN_CANDIDATES["exclusive_area"]),
        "price": find_col(df, COLUMN_CANDIDATES["price"]),
        "deposit": find_col(df, COLUMN_CANDIDATES["deposit"]),
        "monthly_rent": find_col(df, COLUMN_CANDIDATES["monthly_rent"]),
        "floor": find_col(df, COLUMN_CANDIDATES["floor"]),
        "contract_date": find_col(df, COLUMN_CANDIDATES["contract_date"]),
        "contract_type": find_col(df, COLUMN_CANDIDATES["contract_type"]),
        "item": find_col(df, ["주택유형", "물건구분", "유형", "건물용도"]),
    }
    df["_month"] = df[cols["month"]].map(normalize_month_value) if cols["month"] else base_month
    df["_dong"] = df[cols["dong"]].astype(str) if cols["dong"] else ""
    df["_apt_name"] = df[cols["apt_name"]].astype(str) if cols["apt_name"] else ""
    df["_approval_year"] = df[cols["approval_year"]] if cols["approval_year"] else ""
    df["_households"] = df[cols["households"]] if cols["households"] else ""
    df["_area"] = to_numeric_safe(df[cols["area"]]) if cols["area"] else np.nan
    df["_price"] = to_numeric_safe(df[cols["price"]]) if cols["price"] else np.nan
    df["_deposit"] = to_numeric_safe(df[cols["deposit"]]) if cols["deposit"] else np.nan
    df["_monthly_rent"] = to_numeric_safe(df[cols["monthly_rent"]]) if cols["monthly_rent"] else np.nan
    df["_floor"] = to_numeric_safe(df[cols["floor"]]) if cols["floor"] else np.nan
    df["_contract_date"] = df[cols["contract_date"]].astype(str) if cols.get("contract_date") else ""
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
    df["_source_type"] = pd.Series(df["_source_type"]).replace("", np.nan).fillna("기타")
    return df, cols


def add_previous_compare(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    out["_area_key"] = out["_area"].round(2)
    out["_sort_key"] = out["_month"].astype(str)
    group_cols = ["_apt_name", "_area_key"]
    # 단지명이 없는 비아파트는 동+면적으로 비교
    out["_group_name"] = out["_apt_name"].replace("", np.nan).fillna(out["_dong"])
    group_cols = ["_group_name", "_area_key"]
    out = out.sort_values(group_cols + ["_sort_key"])
    out["_prev_value"] = out.groupby(group_cols)[value_col].shift(1)
    out["_prev_floor"] = out.groupby(group_cols)["_floor"].shift(1) if "_floor" in out.columns else np.nan
    out["_prev_contract_date"] = out.groupby(group_cols)["_contract_date"].shift(1) if "_contract_date" in out.columns else ""
    out["_prev_month"] = out.groupby(group_cols)["_month"].shift(1)
    out["_prev_deposit"] = out.groupby(group_cols)["_deposit"].shift(1) if "_deposit" in out.columns else np.nan
    out["_prev_monthly_rent"] = out.groupby(group_cols)["_monthly_rent"].shift(1) if "_monthly_rent" in out.columns else np.nan
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
        if a < 40: return "40㎡ 미만"
        if a < 60: return "40~60㎡"
        if a < 85: return "60~85㎡"
        if a < 102: return "85~102㎡"
        return "102㎡ 이상"
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
        y = int(float(str(approval_year).replace(".0", "")))
        base_y = int(str(base_month)[:4])
        age = base_y - y
        if age <= 5: return "5년 이하"
        if age <= 10: return "6~10년"
        if age <= 20: return "11~20년"
        if age <= 30: return "21~30년"
        return "31년 이상"
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


def calc_trade_common(files: List[Path], region: str, base_month: str, label: str = "아파트 매매") -> Dict[str, Any]:
    result = {"available": False, "label": label, "latest_month": None, "trade_count": 0, "mom_change": None, "compare_count": 0, "up_ratio": 0, "flat_ratio": 0, "down_ratio": 0, "avg_change_rate": None, "signal": f"{label} 원자료 미확인", "monthly_rows": [], "dong_rows": [], "top_complex_rows": [], "top_up_rows": [], "top_down_rows": [], "top_price_rows": [], "area_band_rows": [], "price_band_rows": [], "building_age_rows": [], "monthly_by_type_rows": []}
    df, cols = prepare_trade_df(files, region, base_month)
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

    def rows_by(group_col: str, name_key: str = "name", top=20):
        out = []
        if group_col not in base_df:
            return out
        g = base_df.groupby(group_col).agg(count=("_price", "count")).reset_index().sort_values("count", ascending=False)
        for _, r in g.head(top).iterrows():
            sub = compare_df[compare_df[group_col] == r[group_col]]
            cc = len(sub)
            out.append({name_key: str(r[group_col]), "count": int(r["count"]), "share": safe_ratio(r["count"], len(base_df)), "up_ratio": safe_ratio((sub["_direction"]=="상승").sum(), cc), "flat_ratio": safe_ratio((sub["_direction"]=="보합").sum(), cc), "down_ratio": safe_ratio((sub["_direction"]=="하락").sum(), cc)})
        return out

    dong_rows = rows_by("_dong", "dong")
    base_df["_area_band"] = base_df["_area"].map(classify_area_band)
    base_df["_price_band"] = base_df["_price"].map(classify_price_band_manwon)
    base_df["_building_age_band"] = base_df["_approval_year"].map(lambda x: classify_building_age(x, base_month))
    compare_df["_area_band"] = compare_df["_area"].map(classify_area_band)
    compare_df["_price_band"] = compare_df["_price"].map(classify_price_band_manwon)
    compare_df["_building_age_band"] = compare_df["_approval_year"].map(lambda x: classify_building_age(x, base_month))
    area_band_rows = rows_by("_area_band", "band")
    price_band_rows = rows_by("_price_band", "band")
    building_age_rows = rows_by("_building_age_band", "band")

    top_complex_rows = []
    if not base_df.empty:
        g = base_df.groupby(["_group_name", "_dong"]).agg(count=("_price", "count"), avg_price=("_price", "mean"), main_area=("_area", "median")).reset_index().sort_values(["count", "avg_price"], ascending=[False, False])
        for i, (_, r) in enumerate(g.head(10).iterrows(), 1):
            top_complex_rows.append({"rank": i, "apt_name": r["_group_name"], "dong": r["_dong"], "count": int(r["count"]), "avg_price_eok": money_to_eok_from_manwon(r["avg_price"]), "main_area": None if pd.isna(r["main_area"]) else round(float(r["main_area"]),1), "main_pyeong": area_to_pyeong(r["main_area"])})

    def top_cases(direction: Optional[str] = None, sort_col="_change_rate", asc=False):
        sub = base_df.copy()
        if direction:
            sub = sub[sub["_direction"] == direction]
        if sub.empty:
            return []
        sub = sub.sort_values(sort_col, ascending=asc)
        rows = []
        for i, (_, r) in enumerate(sub.head(10).iterrows(), 1):
            rows.append({"rank": i, "apt_name": r["_group_name"], "dong": r["_dong"], "approval_year": str(r["_approval_year"]) if pd.notna(r["_approval_year"]) else "", "households": str(r["_households"]) if pd.notna(r["_households"]) else "", "area": None if pd.isna(r["_area"]) else round(float(r["_area"]),2), "pyeong": area_to_pyeong(r["_area"]), "current_price_eok": money_to_eok_from_manwon(r["_price"]), "prev_price_eok": money_to_eok_from_manwon(r["_prev_value"]), "change_rate": None if pd.isna(r["_change_rate"]) else round(float(r["_change_rate"]),1), "floor": None if pd.isna(r["_floor"]) else int(r["_floor"]), "prev_floor": None if pd.isna(r.get("_prev_floor")) else int(r.get("_prev_floor")), "contract_date": str(r.get("_contract_date", "")), "prev_contract_date": str(r.get("_prev_contract_date", "")), "month": str(r.get("_month", "")), "prev_month": str(r.get("_prev_month", ""))})
        return rows

    signal = "혼조세"
    if float(base["down_ratio"]) >= 50 and pd.notna(base["avg_change_rate"]) and base["avg_change_rate"] < 0:
        signal = "하락 비중 우세"
    elif float(base["up_ratio"]) >= 50 and pd.notna(base["avg_change_rate"]) and base["avg_change_rate"] > 0:
        signal = "상승 비중 우세"
    elif abs(float(base["up_ratio"]) - float(base["down_ratio"])) <= 10:
        signal = "상승·하락 혼조"

    result.update({"available": True, "latest_month": base["_month"], "trade_count": int(base["trade_count"]), "mom_change": mom, "compare_count": int(base["compare_count"]), "up_ratio": float(base["up_ratio"]), "flat_ratio": float(base["flat_ratio"]), "down_ratio": float(base["down_ratio"]), "avg_change_rate": None if pd.isna(base["avg_change_rate"]) else round(float(base["avg_change_rate"]),1), "signal": signal, "monthly_rows": monthly.tail(13).rename(columns={"_month":"month"}).to_dict("records"), "dong_rows": dong_rows, "top_complex_rows": top_complex_rows, "top_up_rows": top_cases("상승", "_change_rate", False), "top_down_rows": top_cases("하락", "_change_rate", True), "top_price_rows": top_cases(None, "_price", False), "area_band_rows": area_band_rows, "price_band_rows": price_band_rows, "building_age_rows": building_age_rows, "monthly_by_type_rows": monthly_by_type_rows})
    return result


def calc_conversion_rate(files: List[Path], region: str, base_month: str) -> Dict[str, Any]:
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
        if df.empty:
            return result
        # 계약구분 신규 필터
        if cols.get("contract_type"):
            ct = df[cols["contract_type"]].astype(str)
            df = df[ct.str.contains("신규", na=False) | ~ct.str.contains("갱신", na=False)].copy()

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
        dfx["_key"] = region + "|" + dfx["_dong"].astype(str) + "|" + dfx["_jibun"].astype(str) + "|" + dfx["_area_int"].astype(str)

        # 전세/월세 구분 컬럼 탐색
        rent_type_col = find_col(dfx, ["전월세구분", "임대유형", "거래유형"])
        if rent_type_col:
            rt = dfx[rent_type_col].astype(str)
            jeonse = dfx[rt.str.contains("전세", na=False)].copy()
            monthly = dfx[rt.str.contains("월세", na=False)].copy()
        else:
            jeonse = dfx[(dfx["_deposit"].notna()) & ((dfx["_monthly_rent"].isna()) | (dfx["_monthly_rent"] == 0))].copy()
            monthly = dfx[(dfx["_deposit"].notna()) & (dfx["_monthly_rent"].notna()) & (dfx["_monthly_rent"] > 0)].copy()

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
            "applied_rate": median_rate
        })
        return result
    except Exception as e:
        result["status"] = f"ERROR: {e}"
        return result


def calc_rent_common(files: List[Path], region: str, base_month: str, label: str, conversion_rate: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    result = {"available": False, "jeonse": {}, "monthly_rent": {}, "signal": f"{label} 전월세 원자료 미확인"}
    df, cols = prepare_trade_df(files, region, base_month)
    if df.empty:
        return result
    df = df.copy()
    if cols.get("contract_type"):
        ct = df[cols["contract_type"]].astype(str)
        df["_contract_class"] = np.where(ct.str.contains("갱신", na=False), "갱신재계약", "신규")
    else:
        df["_contract_class"] = "신규"
    # 거래건수는 전체, 가격변동 비율은 신규계약만 기준으로 계산함
    jeonse_df = df[(df["_deposit"].notna()) & ((df["_monthly_rent"].isna()) | (df["_monthly_rent"] == 0))].copy()
    rent_df = df[(df["_deposit"].notna()) & (df["_monthly_rent"].notna()) & (df["_monthly_rent"] > 0)].copy()
    applied_rate = None
    if conversion_rate and conversion_rate.get("applied_rate") is not None:
        applied_rate = float(conversion_rate.get("applied_rate"))
    # 전환율 산출 불가 시 환산월세는 보수적으로 결측 처리
    if applied_rate is not None:
        rent_df["_converted_rent"] = rent_df["_monthly_rent"] + (rent_df["_deposit"] * applied_rate / 100 / 12)
    else:
        rent_df["_converted_rent"] = np.nan

    def part_calc(part_df, value_col, part_label):
        if part_df.empty:
            return {"trade_count": 0, "new_count": 0, "renew_count": 0, "compare_count": 0, "up_ratio": 0, "flat_ratio": 0, "down_ratio": 0, "avg_change_rate": None, "signal": f"{part_label} 거래 없음", "monthly_rows": [], "dong_rows": [], "top_complex_rows": [], "area_band_rows": [], "deposit_band_rows": [], "monthly_band_rows": [], "converted_rent_band_rows": [], "building_age_rows": [], "top_up_rows": [], "top_down_rows": [], "top_value_rows": [], "monthly_by_type_rows": []}
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
        monthly_by_type_rows = []
        try:
            if "_source_type" in part_df.columns:
                mt = part_df.groupby(["_month", "_source_type"]).agg(trade_count=(value_col, "count")).reset_index().sort_values(["_month", "_source_type"])
                monthly_by_type_rows = mt.rename(columns={"_month": "month", "_source_type": "type"}).tail(60).to_dict("records")
        except Exception:
            monthly_by_type_rows = []
        base_rows = monthly[monthly["_month"] == base_month]
        base = base_rows.iloc[-1] if not base_rows.empty else monthly.iloc[-1]
        base_df = part_df[part_df["_month"] == base["_month"]].copy()
        comp = base_df[base_df["_direction"]!="비교불가"].copy()
        dong_rows = []
        top_complex_rows = []
        area_band_rows = []
        deposit_band_rows = []
        monthly_band_rows = []
        converted_rent_band_rows = []
        building_age_rows = []
        if not base_df.empty:
            g = base_df.groupby("_dong").agg(count=(value_col, "count")).reset_index().sort_values("count", ascending=False)
            for _, r in g.head(20).iterrows():
                sub = comp[comp["_dong"] == r["_dong"]]
                cc = len(sub)
                dong_rows.append({"dong": r["_dong"], "count": int(r["count"]), "share": safe_ratio(r["count"], len(base_df)), "up_ratio": safe_ratio((sub["_direction"]=="상승").sum(), cc), "flat_ratio": safe_ratio((sub["_direction"]=="보합").sum(), cc), "down_ratio": safe_ratio((sub["_direction"]=="하락").sum(), cc)})

            cg = base_df.groupby(["_group_name", "_dong"]).agg(count=(value_col, "count"), avg_value=(value_col, "mean"), main_area=("_area", "median")).reset_index().sort_values(["count", "avg_value"], ascending=[False, False])
            for i, (_, r) in enumerate(cg.head(10).iterrows(), 1):
                top_complex_rows.append({"rank": i, "apt_name": r["_group_name"], "dong": r["_dong"], "count": int(r["count"]), "avg_value": None if pd.isna(r["avg_value"]) else round(float(r["avg_value"]), 1), "main_area": None if pd.isna(r["main_area"]) else round(float(r["main_area"]), 2), "main_pyeong": area_to_pyeong(r["main_area"])})

            base_df["_area_band"] = base_df["_area"].map(classify_area_band)
            base_df["_deposit_band"] = base_df["_deposit"].map(classify_deposit_band_manwon)
            base_df["_monthly_band"] = base_df["_monthly_rent"].map(classify_monthly_band)
            base_df["_converted_band"] = base_df[value_col].map(classify_converted_rent_band) if value_col == "_converted_rent" else base_df[value_col].map(classify_deposit_band_manwon)
            base_df["_building_age_band"] = base_df["_approval_year"].map(lambda x: classify_building_age(x, base_month))
            comp["_area_band"] = comp["_area"].map(classify_area_band)
            comp["_deposit_band"] = comp["_deposit"].map(classify_deposit_band_manwon)
            comp["_monthly_band"] = comp["_monthly_rent"].map(classify_monthly_band)
            comp["_converted_band"] = comp[value_col].map(classify_converted_rent_band) if value_col == "_converted_rent" else comp[value_col].map(classify_deposit_band_manwon)
            comp["_building_age_band"] = comp["_approval_year"].map(lambda x: classify_building_age(x, base_month))

            def band_rows(col_name, label):
                rr = []
                gg = base_df.groupby(col_name).agg(count=(value_col, "count")).reset_index().sort_values("count", ascending=False)
                for _, br in gg.iterrows():
                    subb = comp[comp[col_name] == br[col_name]]
                    cc = len(subb)
                    rr.append({label: br[col_name], "count": int(br["count"]), "share": safe_ratio(br["count"], len(base_df)), "up_ratio": safe_ratio((subb["_direction"]=="상승").sum(), cc), "flat_ratio": safe_ratio((subb["_direction"]=="보합").sum(), cc), "down_ratio": safe_ratio((subb["_direction"]=="하락").sum(), cc)})
                return rr

            area_band_rows = band_rows("_area_band", "band")
            deposit_band_rows = band_rows("_deposit_band", "band")
            monthly_band_rows = band_rows("_monthly_band", "band")
            converted_rent_band_rows = band_rows("_converted_band", "band")
            building_age_rows = band_rows("_building_age_band", "band")
        def cases(direction, asc):
            sub = base_df[base_df["_direction"] == direction].copy()
            if sub.empty: return []
            sub = sub.sort_values("_change_rate", ascending=asc)
            rows = []
            for i, (_, r) in enumerate(sub.head(10).iterrows(), 1):
                rows.append({"rank": i, "apt_name": r["_group_name"], "dong": r["_dong"], "area": None if pd.isna(r["_area"]) else round(float(r["_area"]),2), "pyeong": area_to_pyeong(r["_area"]), "current_value": None if pd.isna(r[value_col]) else round(float(r[value_col]),1), "prev_value": None if pd.isna(r["_prev_value"]) else round(float(r["_prev_value"]),1), "current_deposit": None if pd.isna(r.get("_deposit")) else round(float(r.get("_deposit")),1), "current_monthly_rent": None if pd.isna(r.get("_monthly_rent")) else round(float(r.get("_monthly_rent")),1), "prev_deposit": None if pd.isna(r.get("_prev_deposit")) else round(float(r.get("_prev_deposit")),1), "prev_monthly_rent": None if pd.isna(r.get("_prev_monthly_rent")) else round(float(r.get("_prev_monthly_rent")),1), "change_rate": None if pd.isna(r["_change_rate"]) else round(float(r["_change_rate"]),1), "floor": None if pd.isna(r["_floor"]) else int(r["_floor"]), "prev_floor": None if pd.isna(r.get("_prev_floor")) else int(r.get("_prev_floor")), "contract_date": str(r.get("_contract_date", "")), "prev_contract_date": str(r.get("_prev_contract_date", "")), "month": str(r.get("_month", "")), "prev_month": str(r.get("_prev_month", ""))})
            return rows
        signal = f"{part_label} 혼조"
        if float(base["down_ratio"]) >= 50: signal = f"{part_label} 하락 비중 우세"
        elif float(base["up_ratio"]) >= 40 and float(base["up_ratio"]) > float(base["down_ratio"]): signal = f"{part_label} 상승 우세"
        def value_cases():
            sub = base_df[base_df[value_col].notna()].copy()
            if sub.empty:
                return []
            sub = sub.sort_values(value_col, ascending=False)
            rows = []
            for i, (_, r) in enumerate(sub.head(10).iterrows(), 1):
                rows.append({"rank": i, "apt_name": r["_group_name"], "dong": r["_dong"], "area": None if pd.isna(r["_area"]) else round(float(r["_area"]),2), "pyeong": area_to_pyeong(r["_area"]), "current_value": None if pd.isna(r[value_col]) else round(float(r[value_col]),1), "prev_value": None if pd.isna(r["_prev_value"]) else round(float(r["_prev_value"]),1), "current_deposit": None if pd.isna(r.get("_deposit")) else round(float(r.get("_deposit")),1), "current_monthly_rent": None if pd.isna(r.get("_monthly_rent")) else round(float(r.get("_monthly_rent")),1), "change_rate": None if pd.isna(r["_change_rate"]) else round(float(r["_change_rate"]),1), "floor": None if pd.isna(r["_floor"]) else int(r["_floor"]), "prev_floor": None if pd.isna(r.get("_prev_floor")) else int(r.get("_prev_floor")), "contract_date": str(r.get("_contract_date", "")), "prev_contract_date": str(r.get("_prev_contract_date", "")), "month": str(r.get("_month", "")), "prev_month": str(r.get("_prev_month", ""))})
            return rows
        return {"trade_count": int(base["trade_count"]), "new_count": int(base.get("new_count", 0)), "renew_count": int(base.get("renew_count", 0)), "compare_count": int(base["compare_count"]), "up_ratio": float(base["up_ratio"]), "flat_ratio": float(base["flat_ratio"]), "down_ratio": float(base["down_ratio"]), "avg_change_rate": None if pd.isna(base["avg_change_rate"]) else round(float(base["avg_change_rate"]),1), "signal": signal, "monthly_rows": monthly.tail(13).rename(columns={"_month":"month"}).to_dict("records"), "dong_rows": dong_rows, "top_complex_rows": top_complex_rows, "area_band_rows": area_band_rows, "deposit_band_rows": deposit_band_rows, "monthly_band_rows": monthly_band_rows, "converted_rent_band_rows": converted_rent_band_rows, "building_age_rows": building_age_rows, "top_up_rows": cases("상승", False), "top_down_rows": cases("하락", True), "top_value_rows": value_cases(), "monthly_by_type_rows": monthly_by_type_rows}
    result.update({"available": True, "jeonse": part_calc(jeonse_df, "_deposit", f"{label} 전세"), "monthly_rent": part_calc(rent_df, "_converted_rent", f"{label} 월세"), "signal": f"{label} 전월세 분석 완료"})
    return result


# =========================================================
# 6. data_pack 생성
# =========================================================


# =========================================================
# 6-1. 지역특성 기반 LLM 서술 강화 및 반복 방지 보조자료
# =========================================================

REGION_PROFILE_MAP = {
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
        "region_summary": "광산구는 수완·첨단·하남·운남 등 신도시·산업·업무 배후수요와 원도심·외곽 주거지가 혼재하는 시장임.",
        "dong_profiles": {
            "수완동": "신도시형 주거 선호와 가족형 실수요가 강한 지역.",
            "첨단동": "산업·업무 배후수요와 신축·준신축 선호가 있는 지역.",
            "하남동": "산업단지 배후와 실거주 수요가 혼재하는 지역.",
            "운남동": "기존 대단지 주거수요가 있는 지역.",
            "송정동": "광역교통·구도심 성격이 혼재하는 지역."
        },
        "writing_cautions": ["수완·첨단은 당월 거래가 있을 때만 선호지로 연결", "산업 배후수요는 하남·첨단 등 관련 생활권 거래가 있을 때만 언급"]
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
        "region_summary": "목포시는 옥암·상동 등 신도심 선호지, 하당 생활권, 원도심·항만권, 북항·연산권 주거지가 혼재하는 시장임.",
        "dong_profiles": {
            "옥암동": "신도심·학군·생활편의 선호가 강한 목포 핵심 주거지.",
            "상동": "하당 생활권과 상업·생활편의 접근성이 있는 주요 주거지.",
            "연산동": "기존 주거지와 준신축 수요가 혼재하는 지역.",
            "석현동": "신흥 주거지와 실수요가 일부 형성되는 지역.",
            "용해동": "기존 주거지와 실수요 거래가 중심인 지역."
        },
        "writing_cautions": ["오룡지구 등 인접 신규공급 영향은 당월 수요 이동·가격 사례가 있을 때만 제한적으로 사용", "원도심 비아파트는 거래건수 기준으로 신중히 해석"]
    },
    "나주시": {
        "region_summary": "나주시는 빛가람 혁신도시, 남평·구도심, 에너지·산업 연계권, 농업·외곽권이 구분되는 시장임.",
        "dong_profiles": {
            "빛가람동": "혁신도시 공공기관·업무 배후수요와 아파트 실거주 수요가 집중되는 지역.",
            "남평읍": "광주 인접성과 전원·실거주 수요가 혼재하는 지역.",
            "송월동": "나주 구도심 생활권과 기존 주거수요가 있는 지역.",
            "성북동": "구도심 생활권과 저가 실수요가 혼재하는 지역.",
            "왕곡면": "산업·에너지 연계 기대와 외곽 토지성 수요가 있을 수 있는 지역."
        },
        "writing_cautions": ["국가산단·에너지산업 요인은 해당 생활권 거래 또는 호가 변화 근거가 있을 때만 연결", "농촌·외곽권은 거래건수가 적으면 방향성 단정 금지"]
    }
}

def get_region_profile(region: str) -> Dict[str, Any]:
    """조사지역에 맞는 지역 특성 프로파일 반환. 일치 항목이 없으면 빈 프로파일 반환."""
    r = str(region or "")
    for key, val in REGION_PROFILE_MAP.items():
        if key in r or r in key:
            out = dict(val)
            out["matched_region"] = key
            return out
    return {
        "matched_region": "",
        "region_summary": "등록된 지역 특성 프로파일이 없음. 당월 data_pack의 읍면동별 거래비중·주요 사례 중심으로만 지역 맥락을 제한적으로 작성.",
        "dong_profiles": {},
        "writing_cautions": ["지역 프로파일이 없으면 지역 특성을 임의 생성하지 말 것", "거래 상위 읍면동과 주요 거래사례에 근거한 해석만 작성"]
    }

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
    for k in ["rank", "apt_name", "_group_name", "dong", "area", "pyeong", "approval_year", "floor", "prev_floor",
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
            "apt_sale_up": _top_names(apt.get("top_up_rows", []), ["apt_name", "_group_name"], 5),
            "apt_sale_down": _top_names(apt.get("top_down_rows", []), ["apt_name", "_group_name"], 5),
            "apt_jeonse_up": _top_names(ar.get("jeonse", {}).get("top_up_rows", []), ["apt_name", "_group_name"], 5),
            "apt_jeonse_down": _top_names(ar.get("jeonse", {}).get("top_down_rows", []), ["apt_name", "_group_name"], 5),
            "apt_rent_up": _top_names(ar.get("monthly_rent", {}).get("top_up_rows", []), ["apt_name", "_group_name"], 5),
            "apt_rent_down": _top_names(ar.get("monthly_rent", {}).get("top_down_rows", []), ["apt_name", "_group_name"], 5),
        }
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


def build_data_pack(input_files: List[Path], region: str, base_month: str, report_version: str) -> Dict[str, Any]:
    progress("원자료 파일 자동 분류 중", 1, 14)
    classified = classify_files(input_files)

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
    progress("가격지수 계산 중", 6, 14)
    price_index = {
        "sale": calc_price_index(classified["price_index_sale"], region, base_month, "매매"),
        "jeonse": calc_price_index(classified["price_index_jeonse"], region, base_month, "전세"),
        "rent": calc_price_index(classified["price_index_rent"], region, base_month, "월세"),
    }
    progress("공급 파이프라인 계산 중", 7, 14)
    supply = calc_supply(classified["supply"], region, base_month)
    progress("아파트 매매 계산 중", 8, 14)
    apt_trade = calc_trade_common(classified["apt_trade"], region, base_month, "아파트 매매")
    progress("전월세전환율 산출 중", 9, 15)
    conversion_rate = calc_conversion_rate(classified["apt_rent"], region, base_month)
    progress("아파트 전월세 계산 중", 10, 15)
    apt_rent = calc_rent_common(classified["apt_rent"], region, base_month, "아파트", conversion_rate)
    progress("비아파트 매매 계산 중", 11, 15)
    non_apt_sale = calc_trade_common(classified["non_apt_sale"], region, base_month, "비아파트 매매")
    progress("비아파트 전월세 계산 중", 12, 15)
    non_apt_rent = calc_rent_common(classified["non_apt_rent"], region, base_month, "비아파트", conversion_rate)
    progress("분양권전매 계산 중", 13, 15)
    presale = calc_trade_common(classified["presale"], region, base_month, "분양권전매")

    warnings = []
    if not population.get("available") or population.get("total") is None:
        warnings.append("인구·세대수 수치가 산출되지 않았습니다.")
    if not age_population.get("available"):
        warnings.append("연령대별 주민등록인구 구조가 산출되지 않았습니다.")
    if apt_trade.get("latest_month") and apt_trade.get("latest_month") != base_month:
        warnings.append(f"아파트 매매 최신월({apt_trade.get('latest_month')})이 기준월({base_month})과 다릅니다.")
    if not classified["price_index_sale"] or not classified["price_index_jeonse"] or not classified["price_index_rent"]:
        warnings.append("가격지수 3종 중 일부 파일이 누락되었습니다.")
    if not classified["supply"]:
        warnings.append("공급 4종 파일이 누락되었거나 분류되지 않았습니다.")
    if not classified["non_apt_sale"] or not classified["non_apt_rent"]:
        warnings.append("비아파트 매매/전월세 파일 일부가 누락되었습니다.")
    if not classified["presale"]:
        warnings.append("분양권전매 파일이 누락되었습니다.")
    if not conversion_rate.get("available"):
        warnings.append("전월세전환율 산출 불가: 아파트 신규 전세·신규 월세 동일 주택형 매칭을 확인하세요.")

    progress("data_pack 구조 생성 중", 14, 15)
    return {
        "meta": {
            "region": region,
            "base_month": base_month,
            "report_version": report_version,
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
            "unsold_units": unsold.get("unsold_units"),
            "after_completion_unsold": unsold.get("after_completion_unsold"),
        },
        "population": population,
        "age_population": age_population,
        "migration": migration,
        "rates": rates,
        "unsold": unsold,
        "conversion_rate": conversion_rate,
        "price_index": price_index,
        "supply": supply,
        "apt_trade": apt_trade,
        "apt_rent": apt_rent,
        "non_apt_sale": non_apt_sale,
        "non_apt_rent": non_apt_rent,
        "presale": presale,
        "region_profile": get_region_profile(region),
        "monthly_change_summary": build_monthly_change_summary({
            "population": population,
            "migration": migration,
            "rates": rates,
            "unsold": unsold,
            "supply": supply,
            "apt_trade": apt_trade,
            "apt_rent": apt_rent,
        }),
        "dong_focus_summary": build_dong_focus_summary({
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
    pop=data["population"]; mig=data["migration"]; rates=data["rates"]; unsold=data["unsold"]
    apt=data["apt_trade"]; ar=data["apt_rent"]; nas=data["non_apt_sale"]; nar=data["non_apt_rent"]; pre=data["presale"]
    pi=data["price_index"]; supply=data["supply"]; conv=data.get("conversion_rate", {})

    def s_count(d, key="trade_count"): return format_num(d.get(key))
    def pct(d, key): return format_rate(d.get(key))

    monthly_core = (
        f"{bm} {region} 주택시장은 인구 {format_num(pop.get('total'))}명, 세대수 {format_num(pop.get('households'))}세대, "
        f"주택담보대출금리 {format_float(rates.get('mortgage_rate'),2)}%, 미분양 {format_num(unsold.get('unsold_units'))}호를 배경으로 "
        f"아파트 매매 {s_count(apt)}건, 하락 비율 {pct(apt,'down_ratio')}가 확인되는 국면으로 분석됨. "
        f"가격지수·공급·비아파트·분양권전매 지표를 함께 보면 시장은 단일 방향보다 생활권·유형별 차별화가 큰 선별 거래 국면으로 판단됨."
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
    pi_c = (
        f"가격지수는 매매 {format_float(pi['sale'].get('latest_value'),2)}, 전세 {format_float(pi['jeonse'].get('latest_value'),2)}, 월세 {format_float(pi['rent'].get('latest_value'),2)}로 산출됨. "
        f"실거래 직전거래 비교는 개별 거래의 가격 변동성을 보여주고, 가격지수는 공식 조사 기준의 완만한 추세를 보완적으로 보여주는 지표로 해석됨."
    )
    apt_c = (
        f"아파트 매매는 {s_count(apt)}건으로 산출되었고, 직전거래 비교 가능 {format_num(apt.get('compare_count'))}건 중 상승 {pct(apt,'up_ratio')}, "
        f"보합 {pct(apt,'flat_ratio')}, 하락 {pct(apt,'down_ratio')}로 나타남. 평균 변동률은 {format_float(apt.get('avg_change_rate'))}%로 확인되어 {apt.get('signal')} 흐름으로 분석됨. "
        f"읍면동·가격구간·평형·연식별 세부표와 TOP 거래사례를 함께 보면 단지별 조건에 따른 선별 거래가 뚜렷한 것으로 판단됨."
    )
    aj_c = (
        f"아파트 전세는 신규계약 기준 {format_num(ar.get('jeonse',{}).get('trade_count'))}건으로 산출되었고, "
        f"상승 {format_rate(ar.get('jeonse',{}).get('up_ratio'))}, 하락 {format_rate(ar.get('jeonse',{}).get('down_ratio'))}로 확인됨. "
        f"전세자금대출금리와 보증금 부담을 고려할 때 전세시장은 단지별 보증금 방어력 차이가 큰 것으로 판단됨."
    )
    rent_c = (
        f"아파트 월세는 환산월세 기준 {format_num(ar.get('monthly_rent',{}).get('trade_count'))}건으로 산출되었고, "
        f"상승 {format_rate(ar.get('monthly_rent',{}).get('up_ratio'))}, 하락 {format_rate(ar.get('monthly_rent',{}).get('down_ratio'))}로 확인됨. "
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
    pre_c = (
        f"분양권전매는 {format_num(pre.get('trade_count'))}건으로 산출되었고, 거래단지와 가격 수준은 신축 기대수요와 기존 매매시장 심리를 보완적으로 보여주는 지표로 판단됨. "
        f"거래규모가 제한적인 경우 분양권전매는 전체 시장 방향을 단정하는 지표가 아니라 신규 공급 기대를 확인하는 보조지표로 활용하는 것이 적절함."
    )
    top_c = (
        "주요 거래사례는 거래단지 TOP10, 상승거래 TOP10, 하락거래 TOP10, 거래금액 상위 TOP10을 함께 확인해 해석해야 함. "
        "상승·하락률이 큰 거래는 층, 면적, 물건 상태, 직전거래 조건 차이가 반영될 수 있어 전체 시장 방향을 단정하는 근거로 사용하지 않는 것이 적절함."
    )
    comprehensive = (
        f"종합적으로 {bm} {region} 주택시장은 인구·금리·미분양·공급 여건과 실거래 가격 변동이 복합적으로 작용하는 국면으로 분석됨. "
        f"아파트 매매는 {apt.get('signal')} 흐름을 보이며, 전세와 월세는 신규계약 및 환산월세 기준으로 별도 점검이 필요한 것으로 판단됨. "
        f"비아파트와 분양권전매는 거래규모가 제한적일 수 있으므로 아파트 중심지표를 보완하는 자료로 활용하는 것이 적절함."
    )

    d = {
        "monthly_core_diagnosis": monthly_core, "population_commentary": population_c, "age_migration_commentary": age_c,
        "rate_commentary": rate_c, "unsold_commentary": unsold_c, "supply_commentary": supply_c, "price_index_commentary": pi_c,
        "conversion_rate_commentary": f"월세 환산 기준은 {conv.get('period', '-')} 기간의 아파트 신규 전세·신규 월세 동일 주택형 매칭 {format_num(conv.get('matched_count'))}건을 기준으로 산출되었으며, 적용 전환율은 {format_float(conv.get('applied_rate'),2)}%로 확인됨. 산식은 전환율(%) = 월세×12÷(전세보증금-월세보증금)×100, 환산월세 = 월세+(보증금×적용전환율÷100÷12)를 적용함.",
        "apt_trade_commentary": apt_c, "apt_jeonse_commentary": aj_c, "apt_rent_commentary": rent_c,
        "non_apt_sale_commentary": non_sale_c, "non_apt_jeonse_commentary": non_j_c, "non_apt_rent_commentary": non_r_c,
        "presale_commentary": pre_c, "top_case_commentary": top_c, "comprehensive_summary": comprehensive,
        "sale_market_summary": f"{apt_c} {non_sale_c} {pre_c}",
        "apt_sale_summary": apt_c,
        "rowhouse_multiplex_sale_summary": "연립·다세대 매매는 아파트 대체 저가 수요, 구도심 접근성, 관리상태, 주차여건에 따라 선별적으로 발생하는 유형으로 판단됨.",
        "detached_multifamily_sale_summary": "단독·다가구 매매는 토지성 가치, 임대수익, 노후주택 정비 가능성, 금리 부담이 함께 반영되는 유형으로 분석됨.",
        "jeonse_market_summary": f"{aj_c} {non_j_c}",
        "apt_jeonse_summary": aj_c,
        "rowhouse_multiplex_jeonse_summary": "연립·다세대 전세는 구도심 저가 전세수요와 보증 안정성, 관리상태가 거래 성립에 영향을 주는 유형으로 판단됨.",
        "detached_multifamily_jeonse_summary": "단독·다가구 전세는 1~2인 가구, 고령층 정주수요, 보증 안정성이 중요한 유형으로 분석됨.",
        "rent_market_summary": f"{rent_c} {non_r_c}",
        "apt_rent_summary": rent_c,
        "rowhouse_multiplex_rent_summary": "연립·다세대 월세는 저가 임대수요와 생활권 접근성, 관리상태에 따라 선별되는 유형으로 판단됨.",
        "detached_multifamily_rent_summary": "단독·다가구 월세는 구도심 생활권, 1~2인 가구, 고령층, 실속형 임대수요가 반영되는 유형으로 분석됨."
    }
    return d


def build_llm_prompt(data_pack: Dict[str, Any]) -> str:
    compact = data_pack
    schema = {k: "" for k in COMMENTARY_KEYS}
    region_profile = data_pack.get("region_profile", {})
    monthly_change_summary = data_pack.get("monthly_change_summary", {})
    dong_focus_summary = data_pack.get("dong_focus_summary", {})
    case_focus_summary = data_pack.get("case_focus_summary", {})
    non_apt_interpretation_summary = data_pack.get("non_apt_interpretation_summary", {})
    supply_pressure_summary = data_pack.get("supply_pressure_summary", {})
    return f"""당신은 한국부동산원 지사 조사 담당자 관점의 월간 주택시장 동향 보고서 작성 보조자입니다.

아래 data_pack.json은 Python으로 계산된 확정 수치입니다.
임의 계산, 임의 추정, 임의 단지명 생성은 금지합니다.

[작성 원칙]
1. 모든 수치·비율·순위·단지명은 data_pack.json 값만 사용합니다.
2. HTML 태그, 마크다운 표, 설명 문구는 작성하지 않습니다.
3. 출력은 반드시 JSON 객체 하나만 작성합니다.
4. 공공기관 보고서 문체로 작성합니다.
5. 문장 끝은 “판단됨”, “분석됨”, “해석됨” 중심으로 작성합니다.
6. 과도한 예측은 피하고, 기준월 자료로 확인되는 범위 내에서 해석합니다.
7. 금지 문구: “신고가 기준”
8. 전세는 갱신계약 제외 신규계약 기준으로 해석합니다.
9. 월세는 환산월세 기준으로 해석합니다.
10. 기준월 실거래는 신고기한 미도래 잠정치임을 필요한 곳에 반영합니다.
11. 비아파트와 분양권전매는 거래규모가 제한적이면 전체 시장 방향을 단정하지 말고 보조지표로 해석합니다.
12. 수치가 null, 0, 빈 목록이면 억지 해석하지 말고 “원자료 확인 필요” 또는 “거래가 제한적”이라고 표현합니다.

[지역특성 기반 서술 강화 규칙]
1. region_profile은 고정 배경자료이며, 문장을 그대로 반복하지 않습니다.
2. 지역 특성은 monthly_change_summary 또는 data_pack에서 당월 변화가 확인된 읍면동·단지·유형에 한해 연결합니다.
3. 읍면동별 거래 비중 상위 1~3개 지역을 우선 확인하되, 해당 읍면동이 region_profile.dong_profiles에 있을 때만 지역특성과 연결합니다.
4. 아파트는 학군·생활편의·신축·구축·단지규모·실거주 선호를 중심으로 해석합니다.
5. 연립·다세대는 구도심 저가 실수요, 노후 주거지, 소규모 임대수요를 중심으로 해석하되, 거래건수가 적으면 가격 방향성을 단정하지 않습니다.
6. 단독·다가구는 대지면적, 노후 주택, 임대수익형 수요, 정비 기대, 외곽·구도심 특성을 중심으로 해석합니다.
7. 전세와 월세는 실거주 수요, 금리 부담, 전세보증금 부담, 소형 임대수요, 갱신계약 비중을 함께 고려합니다.

[매월 반복 방지 규칙]
1. 전월과 동일한 지역 설명을 반복하지 않습니다.
2. 모든 종합 문단에는 monthly_change_summary의 전월 대비 달라진 점, 새롭게 부각된 읍면동, 주요 거래사례 변화, 상승·하락 비율 변화, 공급·미분양·금리 변화 중 하나 이상을 반영합니다.
3. 전월과 유사한 흐름이 유지된 경우에는 반복 설명 대신 “유지”, “이어짐”, “큰 변화 없음”, “전월과 유사”, “확대”, “완화”, “축소” 등 변화 표현을 사용합니다.
4. region_profile의 문장을 그대로 복사하지 말고, 당월 수치 또는 거래사례와 연결하여 새 문장으로 재구성합니다.
5. 당월 데이터에서 근거가 약한 지역 특성은 서술하지 않습니다.
6. 종합 판단 문단에는 반드시 “당월 달라진 점” 또는 “유지되지만 관찰이 필요한 점”을 1개 이상 포함합니다.

[종합 판단 권장 문장 구조]
① 당월 시장 방향
② 거래 집중 지역 또는 주요 단지
③ 가격 변동 주요 사례 또는 상승·하락 비율 변화
④ 지역 특성에 따른 해석
⑤ 향후 관찰 포인트

[부동산원 해당지역 담당 조사 전문가 관점 보강 규칙]
1. 단순 요약이 아니라 조사자가 현장에서 설명하듯 “왜 그런 흐름이 나타났는지”를 1~2문장 더 보강합니다.
2. dong_focus_summary의 상위 읍면동, case_focus_summary의 대표 거래사례, non_apt_interpretation_summary의 유형별 거래건수, supply_pressure_summary의 공급압력을 각각 필요한 문단에 연결합니다.
3. 분석 내용은 풍부하게 쓰되 data_pack에 없는 외부 개발 호재·정비사업·단지명은 임의 생성하지 않습니다.
4. 거래건수·비교가능건수·신규계약 건수가 적은 경우에는 “관찰 필요”, “제한적 해석”을 명확히 표시합니다.
5. 전문 디자이너 관점에서 HTML 출력에 맞게 한 문장이 너무 길지 않도록 1문장 80~120자 내외로 나누어 작성합니다.

[출력 JSON 구조]
아래 key를 반드시 모두 포함하세요. JSON 외 문장은 쓰지 마세요.
{json.dumps(schema, ensure_ascii=False, indent=2)}

[Ⅵ. 종합 시황 서술 보고서 작성 기준]
- sale_market_summary: 매매 종합 4~6문장. 당월 매매 변화 포인트와 거래 집중 읍면동을 반드시 반영.
- apt_sale_summary: 아파트 매매 4~6문장. 주요 읍면동·단지 사례와 지역 특성을 연결.
- rowhouse_multiplex_sale_summary: 연립·다세대 매매 3~5문장. 거래건수 적으면 방향성 단정 금지.
- detached_multifamily_sale_summary: 단독·다가구 매매 3~5문장. 임대수익형·노후주거·외곽/구도심 특성은 당월 거래 근거가 있을 때만 사용.
- jeonse_market_summary: 전세 종합 4~6문장. 신규계약 기준 변화와 금리·보증금 부담 반영.
- apt_jeonse_summary: 아파트 전세 3~5문장. 거래 집중 읍면동 및 보증금 상승·하락 사례 반영.
- rowhouse_multiplex_jeonse_summary: 연립·다세대 전세 3~5문장. 저가 전세·관리상태·보증 안정성 중심.
- detached_multifamily_jeonse_summary: 단독·다가구 전세 3~5문장. 1~2인 가구·고령층·구도심 수요는 근거가 있을 때만 사용.
- rent_market_summary: 월세 종합 4~6문장. 환산월세 기준 변화와 전세 부담에 따른 월세 전환 가능성 반영.
- apt_rent_summary: 아파트 월세 3~5문장. 환산월세 상승·하락 사례와 주요 읍면동 반영.
- rowhouse_multiplex_rent_summary: 연립·다세대 월세 3~5문장.
- detached_multifamily_rent_summary: 단독·다가구 월세 3~5문장.

[region_profile]
{json.dumps(region_profile, ensure_ascii=False, indent=2)}

[monthly_change_summary]
{json.dumps(monthly_change_summary, ensure_ascii=False, indent=2)}

[dong_focus_summary]
{json.dumps(dong_focus_summary, ensure_ascii=False, indent=2)}

[case_focus_summary]
{json.dumps(case_focus_summary, ensure_ascii=False, indent=2)}

[non_apt_interpretation_summary]
{json.dumps(non_apt_interpretation_summary, ensure_ascii=False, indent=2)}

[supply_pressure_summary]
{json.dumps(supply_pressure_summary, ensure_ascii=False, indent=2)}

[HTML 디자인·가독성 참고]
- HTML 본문은 A4 PDF 출력 기준으로 읽히므로, 한 문단이 과도하게 길어지지 않게 작성합니다.
- 수치 나열보다 “핵심 변화 → 원인 추정 → 관찰 포인트” 순서로 정리합니다.
- 표·차트가 이미 제공하는 단순 수치는 반복하지 말고, 표·차트에서 드러나는 의미를 서술합니다.
- 종합 시황은 매매·전세·월세 박스별로 서로 다른 논지를 가져야 하며, 같은 문장 구조 반복을 피합니다.

[data_pack.json]
{json.dumps(compact, ensure_ascii=False, indent=2)}
"""

def extract_json_from_text(text: str) -> Dict[str, Any]:
    raw = text.strip()
    m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw, flags=re.DOTALL | re.IGNORECASE)
    if m:
        raw = m.group(1)
    else:
        s, e = raw.find("{"), raw.rfind("}")
        if s >= 0 and e > s:
            raw = raw[s:e+1]
    return json.loads(raw)


def merge_commentary(base: Dict[str, str], new: Dict[str, Any]) -> Dict[str, str]:
    out = dict(base)
    for k in COMMENTARY_KEYS:
        if k in new and str(new[k]).strip():
            out[k] = str(new[k])
    return {k: out.get(k, "") for k in COMMENTARY_KEYS}


# =========================================================
# 8. HTML 렌더링
# =========================================================

DISPLAY_NAME_MAP = {
    "month": "기준월",
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
    "compare_count": "비교가능건수",
    "avg_change_rate": "평균 변동률",
    "up_ratio": "🔴 상승 비율",
    "flat_ratio": "⚪ 보합 비율",
    "down_ratio": "🔵 하락 비율",
    "상승": "🔴 상승건수",
    "보합": "⚪ 보합건수",
    "하락": "🔵 하락건수",
    "dong": "읍면동",
    "share": "비중",
    "count": "거래건수",
    "band": "구간",
    "rank": "순위",
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

def display_col_name(col: Any) -> str:
    _extra_display_names = {
        "new_count": "신규계약 건수",
        "renew_count": "갱신재계약 건수",
        "trade_count": "거래건수",
        "compare_count": "비교가능건수",
        "avg_change_rate": "평균 변동률",
        "up_ratio": "🔴 상승 비율",
        "flat_ratio": "⚪ 보합 비율",
        "down_ratio": "🔵 하락 비율",
    }
    s = str(col)
    return _extra_display_names.get(s, DISPLAY_NAME_MAP.get(s, s))

def html_value(v: Any, col: Any = "") -> str:
    if v is None:
        return "-"
    try:
        if pd.isna(v):
            return "-"
    except Exception:
        pass
    c = str(col).lower()
    if isinstance(v, (int, float)):
        if any(k in c for k in ["rate", "ratio", "변동률", "비율"]):
            return f"{float(v):.1f}"
        if c in ["value", "지수"] or "index" in c:
            return f"{float(v):.4f}"
        if abs(float(v)) < 100 and not float(v).is_integer():
            return f"{float(v):.2f}"
        return f"{int(round(float(v))):,}"
    return str(v)

def table_rows_html(rows: List[Dict[str, Any]], max_rows: int = 13, highlight_latest: bool = True) -> str:
    if not rows:
        return "<p class='muted'>자료 없음</p>"
    rows = rows[-max_rows:]
    hide_cols = {"avg_value", "avg_price_eok", "평균값", "평균가", "평균가격"}
    cols = [c for c in list(rows[0].keys()) if str(c) not in hide_cols]
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
    th = "".join(f"<th>{display_col_name(c)}</th>" for c in cols)
    trs = []
    for idx, r in enumerate(rows):
        cls = " class='latest-row'" if highlight_latest and idx == len(rows)-1 else ""
        tds = []
        for c in cols:
            val = r.get(c)
            classes = []
            dcls = direction_col_class(c)
            if dcls:
                classes.append(dcls)
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
            tds.append(f"<td{cell_cls}>{html_value(val, c)}</td>")
        trs.append(f"<tr{cls}>" + "".join(tds) + "</tr>")
    first_col = str(cols[0]) if cols else ""
    is_time_series = len(rows) >= 12 and any(k in first_col.lower() for k in ["month", "기준월", "_month"])
    table_cls = "market-strength-table compact-wide-table" + (" time-series-table" if is_time_series else "")
    return f"<div class='table-wrap'><table class='{table_cls}'><thead><tr>{th}</tr></thead><tbody>{''.join(trs)}</tbody></table></div>"


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
        delta = float(curr) - float(prev)
    except Exception:
        delta = None
    try:
        chg_s = f"{float(chg):+.1f}%"
    except Exception:
        chg_s = "-"
    if delta is None:
        return f"변동률 {chg_s}"
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
        name = r.get("apt_name") or r.get("_group_name") or "-"
        year = str(r.get("approval_year") or "").replace(".0", "")
        dong = r.get("dong") or "-"
        area = r.get("area")
        pyeong = r.get("pyeong")
        floor = r.get("floor")
        prev_floor = r.get("prev_floor")
        contract_date = format_contract_date_display(r.get("contract_date") or "", r.get("month") or "")
        prev_contract_date = format_contract_date_display(r.get("prev_contract_date") or "", r.get("prev_month") or "")

        area_s = f"전용 {area}㎡({pyeong}평)" if area and pyeong else (f"전용 {area}㎡" if area else "전용면적 미상")
        floor_s = f"현재 {floor}층 / 직전 {prev_floor}층" if floor is not None and prev_floor is not None else (f"현재 {floor}층 / 직전층 원자료 확인 필요" if floor is not None else "현재층·직전층 원자료 확인 필요")
        date_pair_s = (f"현재거래 {contract_date}" if contract_date else "") + (f" / 직전거래 {prev_contract_date}" if prev_contract_date else "")
        prev_date_s = f" · {date_pair_s}" if date_pair_s else ""

        if value_label == "매매":
            curr = r.get("current_price_eok", r.get("current_value", None))
            prev = r.get("prev_price_eok", r.get("prev_value", None))
            curr_s = _fmt_eok_value(curr)
            prev_s = _fmt_eok_value(prev)
            small_s = f"현재 {curr_s} / 직전 {prev_s}"
            value_b = curr_s
            value_span = _fmt_delta(curr, prev, value_label, chg)
        elif value_label == "전세":
            curr = r.get("current_value", None)
            prev = r.get("prev_value", None)
            curr_s = _fmt_manwon_value(curr)
            prev_s = _fmt_manwon_value(prev)
            small_s = f"현재 보증금 {curr_s} / 직전 보증금 {prev_s}"
            value_b = curr_s
            value_span = _fmt_delta(curr, prev, value_label, chg)
        elif value_label == "월세":
            curr = r.get("current_value", None)
            prev = r.get("prev_value", None)
            curr_s = _fmt_manwon_value(curr)
            prev_s = _fmt_manwon_value(prev)
            deposit_s = _fmt_manwon_value(r.get("current_deposit"))
            monthly_s = _fmt_manwon_value(r.get("current_monthly_rent"))
            small_s = f"보증금 {deposit_s} / 월세 {monthly_s} / 직전 환산 {prev_s}"
            value_b = curr_s
            value_span = _fmt_delta(curr, prev, value_label, chg)
        else:
            curr = r.get("current_value", None)
            prev = r.get("prev_value", None)
            curr_s = html_value(curr)
            prev_s = html_value(prev)
            small_s = f"현재 {curr_s} / 직전 {prev_s}"
            value_b = curr_s
            value_span = _fmt_delta(curr, prev, value_label, chg)

        year_s = f' <span class="case-year">(사용승인년도 {year}년)</span>' if year and year != "nan" else ' <span class="case-year"></span>'
        items.append(f"""
        <div class="top-case-item {cls}">
          <div class="case-rank">{rank}</div>
          <div class="case-body">
            <b>{name}{year_s}</b>
            <span>{dong} · {area_s} · {floor_s}{prev_date_s}</span>
            <small>{small_s}</small>
          </div>
          <div class="case-value"><b>{value_b}</b><span>{value_span}</span></div>
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


def chart_block(title: str, chart_type: str, labels: List[Any], datasets: List[Dict[str, Any]], chart_id: str, note: str = "", stacked: bool = False, y1: bool = False) -> str:
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
    if chart_type in ["doughnut", "pie"]:
        cfg["options"].pop("scales", None)
        cfg["options"]["cutout"] = "66%" if chart_type == "doughnut" else None
    mix_cls = " mix-chart" if y1 else ""
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

def chart_monthly_price_change(rows: List[Dict[str, Any]], title: str, chart_id: str, label_prefix: str = "") -> str:
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
        "요청사항 반영: 평균 변동률 꺾은선은 제거하고, 상승·보합·하락 건수를 의미색상 막대로 표시함.",
        stacked=True,
        y1=False
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


def chart_monthly_strength(rows: List[Dict[str, Any]], title: str, chart_id: str, label_prefix: str = "") -> str:
    """거래량 및 상승·하락 비율 추이.
    요청사항 반영: 거래량은 막대, 상승·하락 비율만 꺾은선으로 표시. 보합 꺾은선은 표시하지 않음.
    """
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = rows[-13:]
    labels = [r.get("month","") for r in rows]
    prefix = (label_prefix + " ") if label_prefix else ""
    datasets = [
        {"type":"bar", "label":f"{prefix}거래량(건)", "data":[r.get("trade_count") for r in rows], "yAxisID":"y", "order":3, "borderRadius":4},
        {"type":"line", "label":"상승", "data":[r.get("up_ratio") for r in rows], "yAxisID":"y1", "order":1, "borderWidth":3, "pointRadius":3},
        {"type":"line", "label":"하락", "data":[r.get("down_ratio") for r in rows], "yAxisID":"y1", "order":2, "borderWidth":3, "pointRadius":3},
    ]
    return chart_block(
        title,
        "bar",
        labels,
        datasets,
        chart_id,
        "요청사항 반영: 거래량은 막대, 상승·하락 비율만 꺾은선으로 표시함.",
        y1=True
    )


def chart_direction_rows(rows: List[Dict[str, Any]], title: str, chart_id: str, x: str = "band", max_rows: int = 12) -> str:
    if not rows:
        return "<p class='muted'>차트 자료 없음</p>"
    rows = rows[:max_rows]
    labels = [r.get(x, "") for r in rows]
    def cnt(r, key):
        try:
            return float(r.get("count") or 0) * float(r.get(key) or 0) / 100
        except Exception:
            return 0
    datasets = [
        {"label":"상승", "data":[cnt(r,"up_ratio") for r in rows]},
        {"label":"보합", "data":[cnt(r,"flat_ratio") for r in rows]},
        {"label":"하락", "data":[cnt(r,"down_ratio") for r in rows]},
    ]
    return chart_block(title, "bar", labels, datasets, chart_id, "template_v58 기준: 상승·보합·하락 의미색상 고정 누적막대.", stacked=True)

def _dominant_direction(up: Any, flat: Any, down: Any) -> str:
    vals = [("상승", _clean_num(up)), ("보합", _clean_num(flat)), ("하락", _clean_num(down))]
    vals = sorted(vals, key=lambda x: x[1], reverse=True)
    return vals[0][0] if vals else "보합"

def doughnut_market_strength(title: str, up: Any, flat: Any, down: Any, chart_id: str, compare_count: Any = None) -> str:
    upv, flatv, downv = _clean_num(up), _clean_num(flat), _clean_num(down)
    compare_s = format_num(compare_count) if compare_count not in [None, ""] else "-"
    dominant = _dominant_direction(upv, flatv, downv)
    center_text = f"{dominant} 우세" if dominant in ["상승", "하락"] else "보합 우세"
    sub_text = f"{compare_s}건" if compare_s != "-" else "비교가능"
    cfg = {
        "type": "doughnut",
        "data": {
            "labels": ["상승", "보합", "하락"],
            "datasets": [{"label": title, "data": [upv, flatv, downv], "borderWidth": 1}]
        },
        "options": {
            "cutout": "66%",
            "responsive": True,
            "maintainAspectRatio": False,
            "plugins": {
                "legend": {"position": "bottom"},
                "title": {"display": False},
                "centerText": {"text": center_text, "subtext": sub_text}
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
        <span class='dir-up'>상승 {upv:.1f}%</span>
        <span class='dir-flat'>보합 {flatv:.1f}%</span>
        <span class='dir-down'>하락 {downv:.1f}%</span>
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
    return mini_chart_block(title, "bar", years, [{"label": title, "data": data, "backgroundColor": stage_color["bg"], "borderColor": stage_color["border"]}], chart_id)


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

    charts["core_donuts"] = (
        "<div class='survey-donut-section'><div class='survey-donut-grid'>"
        + doughnut_market_strength("매매 시장강도", apt.get("up_ratio"), apt.get("flat_ratio"), apt.get("down_ratio"), "surveySaleDonut", apt.get("compare_count"))
        + doughnut_market_strength("전세 시장강도", ar.get("jeonse", {}).get("up_ratio"), ar.get("jeonse", {}).get("flat_ratio"), ar.get("jeonse", {}).get("down_ratio"), "surveyJeonseDonut", ar.get("jeonse", {}).get("compare_count"))
        + doughnut_market_strength("월세 시장강도", ar.get("monthly_rent", {}).get("up_ratio"), ar.get("monthly_rent", {}).get("flat_ratio"), ar.get("monthly_rent", {}).get("down_ratio"), "surveyRentDonut", ar.get("monthly_rent", {}).get("compare_count"))
        + "</div></div>"
    )

    # Ⅰ. 인구·세대·이동·금리·미분양
    charts["population"] = chart_population_combo(pop, "chart1")
    charts["migration"] = chart_migration_combo(mig, "chart2")
    charts["age_population_latest"] = chart_age_bar_by_group(age_pop.get("age_rows", []), "최신월 연령대별 주민등록인구", "chartAgePopLatest", "최신월 인구")
    charts["age_population_change"] = chart_age_bar_by_group(age_pop.get("age_rows", []), "최근 13개월 연령대별 증감", "chartAgePopChange", "13개월 증감")
    charts["age_net_combined"] = chart_age_net_combined(age_move, "chartAgeNetYoung")
    charts["rates"] = chart_from_rows(rates, "예금은행 대출금리 추세", "chart3", "month", ["mortgage_rate", "jeonse_rate", "sme_rate", "business_rate"], "line")
    charts["unsold"] = chart_from_rows(unsold, "미분양·준공후 미분양 추이", "chart4", "month", ["unsold_units", "after_completion_unsold"], "bar")

    # 공급: mini-chart-card 19개
    charts["supply_permit_total"] = mini_chart_from_rows_year_label(supply, "광주광역시 인허가 총량 시계열", "totalPermit", "month", "permit")
    charts["supply_start_total"] = mini_chart_from_rows_year_label(supply, "광주광역시 착공 총량 시계열", "totalStart", "month", "start")
    charts["supply_complete_total"] = mini_chart_from_rows_year_label(supply, "광주광역시 준공 총량 시계열", "totalComplete", "month", "completion")
    sup_obj = data.get("supply", {})
    for t in ["아파트", "연립", "다세대", "단독", "다가구"]:
        charts[f"supply_permit_{t}"] = chart_grouped_by_type_filtered(sup_obj.get("permit_by_type_rows", []), f"인허가 - {t}", f"permitChart_{t}", t)
        charts[f"supply_start_{t}"] = chart_grouped_by_type_filtered(sup_obj.get("start_by_type_rows", []), f"착공 - {t}", f"startChart_{t}", t)
        charts[f"supply_complete_{t}"] = chart_grouped_by_type_filtered(sup_obj.get("completion_by_type_rows", []), f"준공 - {t}", f"completeChart_{t}", t)

    # Ⅲ. 매매: template_v58의 9개 슬롯만 사용
    charts["apt_monthly"] = chart_monthly_price_change(apt.get("monthly_rows", []), "아파트 매매 월간 거래량 및 가격변동 통합", "chart7", "매매")
    charts["apt_strength"] = chart_monthly_strength(apt.get("monthly_rows", []), "매매 거래량 및 상승·하락 비율 추이", "chart33", "매매")
    charts["apt_dong"] = chart_direction_rows(apt.get("dong_rows", [])[:10], "당월 읍면동별 매매 시장강도", "chart11", "dong", max_rows=10)
    charts["apt_area_share"] = chart_direction_rows(apt.get("area_band_rows", []), "평형별 매매 거래 비중", "chart12", "band", max_rows=20)
    charts["apt_dong_detail"] = chart_direction_rows(apt.get("dong_rows", [])[:10], "아파트 매매 읍면동별 세부 현황", "chart30", "dong", max_rows=10)
    charts["apt_price_band"] = chart_direction_rows(apt.get("price_band_rows", []), "아파트 매매 거래금액별 세부 현황", "chart13", "band", max_rows=20)
    charts["apt_age_band"] = chart_direction_rows(apt.get("building_age_rows", []), "아파트 매매 연식별 세부 현황", "chart14", "band", max_rows=20)
    charts["non_apt_sale"] = chart_non_apt_monthly_by_type(nas.get("monthly_by_type_rows", []), "비아파트 물건유형별 최근 13개월 매매 거래 추이(연립·다세대·단독·다가구)", "chart15")
    charts["presale"] = chart_from_rows(presale.get("monthly_rows", []), "아파트 분양권전매 최근 13개월 거래 추이", "chart16", "month", ["trade_count"], "bar")

    # Ⅳ. 전세: template_v58의 8개 슬롯만 사용
    charts["apt_jeonse_monthly"] = chart_monthly_price_change(ar.get("jeonse", {}).get("monthly_rows", []), "아파트 전세 월간 거래량 및 가격변동 통합", "chart8", "전세")
    charts["apt_jeonse_strength"] = chart_monthly_strength(ar.get("jeonse", {}).get("monthly_rows", []), "전세 거래량 및 상승·하락 비율 추이", "chart34", "전세")
    charts["apt_jeonse_market_dong"] = chart_direction_rows(ar.get("jeonse", {}).get("dong_rows", [])[:10], "당월 읍면동별 전세 시장강도", "chart18", "dong", max_rows=10)
    charts["apt_jeonse_dong"] = chart_direction_rows(ar.get("jeonse", {}).get("dong_rows", [])[:10], "아파트 전세 읍면동별 세부 현황", "chart31", "dong", max_rows=10)
    charts["apt_jeonse_deposit"] = chart_direction_rows(ar.get("jeonse", {}).get("deposit_band_rows", []), "아파트 전세 보증금별 세부 현황", "chart19", "band", max_rows=20)
    charts["apt_jeonse_age"] = chart_direction_rows(ar.get("jeonse", {}).get("building_age_rows", []), "아파트 전세 연식별 세부 현황", "chart25", "band", max_rows=20)
    charts["apt_jeonse_area"] = chart_direction_rows(ar.get("jeonse", {}).get("area_band_rows", []), "아파트 전세 평형별 세부 현황", "chart26", "band", max_rows=20)
    charts["non_apt_jeonse"] = chart_non_apt_monthly_by_type(nar.get("jeonse", {}).get("monthly_by_type_rows", []), "비아파트 물건유형별 최근 13개월 전세 거래 추이(연립·다세대·단독·다가구)", "chart20")

    # Ⅴ. 월세: template_v58의 9개 슬롯만 사용
    charts["apt_rent_monthly"] = chart_monthly_price_change(ar.get("monthly_rent", {}).get("monthly_rows", []), "아파트 월세 월간 거래량 및 가격변동 통합", "chart9", "월세")
    charts["apt_rent_strength"] = chart_monthly_strength(ar.get("monthly_rent", {}).get("monthly_rows", []), "월세 거래량 및 상승·하락 비율 추이", "chart35", "월세")
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
<title>{{ meta.region }} 주택시장 월간동향</title>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" rel="stylesheet"/>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
    '하락': { border: '#3b82f6', background: '#3b82f6cc', soft: '#3b82f633' }
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
.notice {max-width:1280px; margin:0 auto 16px; background:#fefce8; border:1px solid #fde68a; color:#713f12; padding:12px 16px; border-radius:14px; font-weight:700;} .kpis { max-width:1280px; margin:0 auto 16px; display:grid; grid-template-columns:repeat(4,1fr); gap:14px; } .kpis div { background:white; border-radius:18px; padding:18px; box-shadow:0 8px 20px rgba(15,23,42,.07); } .kpis span { display:block; color:#64748b; font-size:13px; } .kpis b { display:block; font-size:25px; margin:5px 0; } .kpis small { color:#64748b; }
.chapter { max-width:1280px; margin:18px auto; padding:22px; border-radius:24px; page-break-inside:avoid; } .chapter>h2 { font-size:24px; font-weight:900; margin-bottom:16px; }
.card { background:rgba(255,255,255,.82); border:1px solid #e2e8f0; box-shadow:0 8px 24px rgba(15,23,42,.05); border-radius:20px; padding:20px; margin:16px 0; page-break-inside:avoid; } .card h3 { font-size:19px; font-weight:800; margin-bottom:10px; } .card h4 { font-size:16px; font-weight:800; margin:16px 0 8px; }
.meaning { background:#f1f5f9; border-left:4px solid #64748b; padding:10px 12px; border-radius:10px; color:#334155; font-size:14px; } .note { background:#fff7ed; border:1px solid #fed7aa; border-left:5px solid #f97316; border-radius:14px; padding:13px 15px; margin-top:12px; font-weight:600; color:#431407; }
.table-wrap { overflow-x:auto; margin:12px 0; border-radius:14px; border:1px solid #e2e8f0; } table { width:100%; border-collapse:collapse; font-size:12px; background:white; } th { background:#1e293b; color:white; padding:9px 8px; text-align:center; white-space:nowrap; } td { border-top:1px solid #e2e8f0; padding:8px; text-align:center; white-space:nowrap; } tr:nth-child(even) td { background:#f8fafc; } .latest-row td { background:#fef9c3 !important; font-weight:800; color:#713f12; }
.chart-box { position:relative; background:white; border:1px solid #e2e8f0; border-radius:16px; padding:10px; margin:12px 0; page-break-inside:avoid; min-height:320px; } .chart-box canvas { width:100% !important; height:100% !important; display:block; }
.market-block { background:white; border:1px solid #c7d2fe; border-radius:22px; padding:22px; margin:18px 0; page-break-inside:avoid; box-shadow:0 8px 24px rgba(79,70,229,.07); } .market-block>h3 { font-size:21px; font-weight:900; margin-bottom:12px; display:flex; align-items:center; gap:8px; } .market-item { background:#f8fafc; border-left:6px solid #6366f1; border-radius:16px; padding:15px 17px; margin:12px 0; page-break-inside:avoid; } .market-item h4 { font-size:16px; font-weight:900; margin-bottom:7px; color:#1e293b; } .market-item p { color:#1e293b; }
.narr { background:white; border-left:6px solid #6366f1; padding:16px 18px; margin:12px 0; border-radius:16px; page-break-inside:avoid; } .narr h3 { font-weight:900; font-size:17px; margin-bottom:8px; } .narr p { color:#1e293b; }
.bg-slate { background:#f1f5f9; } .bg-amber { background:#fffbeb; } .bg-rose { background:#fff1f2; } .bg-emerald { background:#ecfdf5; } .bg-blue { background:#eff6ff; } .bg-indigo { background:#eef2ff; }
@media print { body { background:white; } .hero,.chapter,.card,.chart-box,.narr { page-break-inside:avoid; box-shadow:none; } .kpis { grid-template-columns:repeat(2,1fr); } }
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


/* v40 보고서 품질 개선: 핵심 진단·기준월 관리표·관찰포인트 */
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
.case-rank { font-weight:900; color:#111827; text-align:center; }
.case-body b { display:block; font-size:12.5px; color:#0f172a; word-break:keep-all; }
.case-body span { display:block; font-size:11px; color:#64748b; line-height:1.35; margin-top:2px; }
.case-body small { display:block; font-size:10.5px; color:#475569; line-height:1.35; margin-top:2px; }
.case-year { color:#475569; font-weight:800; }
.case-value { text-align:right; min-width:72px; }
.case-value b { display:block; font-size:13px; color:#0f172a; white-space:nowrap; }
.case-value span { display:block; font-size:11px; font-weight:900; white-space:nowrap; }
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
  height: 210px;
  min-height: 210px;
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
.survey-donut-box{height:210px!important}
.survey-donut-box canvas{height:210px!important}
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
@media print {
  .market-strength-table td.dir-up,
  .market-strength-table td.dir-flat,
  .market-strength-table td.dir-down,
  .compact-wide-table td.dir-up,
  .compact-wide-table td.dir-flat,
  .compact-wide-table td.dir-down {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
}


/* === v5.8.38 연령대별 인구 시계열 및 전월 대비 증감률 표 === */
.age-pop-trend-table-wrap { overflow-x:auto !important; }
.age-pop-trend-table th { text-align:center !important; white-space:nowrap !important; }
.age-pop-trend-table td { text-align:center !important; white-space:nowrap !important; font-weight:700; }
.age-pop-trend-table td.age-up .age-rate,
.age-pop-trend-table .age-rate.pos { color:#dc2626 !important; font-weight:900; }
.age-pop-trend-table td.age-down .age-rate,
.age-pop-trend-table .age-rate.neg { color:#2563eb !important; font-weight:900; }
.age-pop-trend-table td.age-flat .age-rate,
.age-pop-trend-table .age-rate.neutral { color:#64748b !important; font-weight:800; }
.age-pop-trend-table .latest-row td { background:#fff8bf !important; }
@media print {
  .age-pop-trend-table { font-size:8.1px !important; }
  .age-pop-trend-table th,
  .age-pop-trend-table td { padding:4px 3px !important; }
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


/* === v5.8.38 최근 13개월 누적순이동 연령별 비교 가로 막대차트 === */
.age-net-horizontal-chart .chart-wrap {
  height:360px !important;
}
@media print {
  .age-net-horizontal-chart .chart-wrap,
  .age-net-horizontal-chart canvas {
    height:280px !important;
    max-height:280px !important;
  }
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
  height: 230px !important;
  min-height: 230px !important;
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

</style>
</head>
<body>
<div class="print-guide">PDF 저장 안내: 브라우저 인쇄(Ctrl+P) → A4 → 배경 그래픽 포함으로 저장</div>
<section class="hero">
  <div>
    <div class="eyebrow">MONTHLY HOUSING MARKET REPORT</div>
    <h1>{{ meta.region }} 주택시장 월간동향</h1>
    <div class="subtitle">{{ meta.base_month }} 월간동향보고서 · 국토교통부 실거래가, 주민등록, 미분양, 공급, 가격지수, 금리 통합 분석</div>
  </div>
  <div class="meta">
    <b>기준월</b><span>{{ meta.base_month }}</span>
    <b>분석기간</b><span>최근 13개월 / 실거래 최근 3년</span>
    <b>작성기준</b><span>{{ meta.trade_basis }}, 기준월 잠정치</span>
    <b>환산월세</b><span>중위 전월세전환율 {{ conversion_rate.applied_rate | fmt_float }}% 적용</span>
  </div>
</section>
<div class="notice">※ 자료 유의: 통계별 공표 시차에 따라 일부 지표는 최신 공표월 기준으로 작성함. 실거래는 기준월 계약분 잠정치이며 신고기한 미도래로 향후 보정될 수 있음.</div>
<section class="kpis">
  <div><span>총인구</span><b>{{ cards.population | fmt_num }}</b><small>전월 {{ cards.population_mom | fmt_num }}</small></div>
  <div><span>총세대수</span><b>{{ cards.households | fmt_num }}</b><small>전월 {{ cards.households_mom | fmt_num }}</small></div>
  <div><span>아파트 매매</span><b>{{ cards.apt_trade_count | fmt_num }}건</b><small>전월 {{ cards.apt_trade_mom | fmt_num }}건</small></div>
  <div><span>미분양</span><b>{{ cards.unsold_units | fmt_num }}호</b><small>준공후 {{ cards.after_completion_unsold | fmt_num }}호 / 최신 {{ unsold.latest_month }}</small></div>
</section>
<div class="notice quick-reading-rule">빠른 읽기 기준: <span class="hl-yellow">노란색은 핵심 수치</span>, <span class="hl-red">붉은색은 위험·부담 요인</span>, <span class="hl-green">초록색은 긍정·방어 요인</span>, <span class="hl-blue">파란색은 최종 시장 판단·관찰 포인트</span>를 의미함. 단, 자동 강조는 LLM 해설 중 핵심 문장 1~2개에만 제한 적용함.</div>
<div class="cover-break"></div>

{% if warnings %}
<h2>데이터 품질 확인</h2>
<div class="warn">{% for w in warnings %}· {{ w }}<br>{% endfor %}</div>
{% endif %}

<section class="chapter bg-indigo monthly-diagnosis-section"><h2>월간 핵심 진단</h2><div class="card diagnosis-card">
<h3>1. {{ meta.base_month }} {{ meta.region.split()[-1] }} 주택시장 핵심 판단</h3>
<div class="meaning">※ 지표 의미: 월간 핵심 진단은 인구·금리·미분양·공급·실거래 지표를 종합하여 기준월 시장 국면을 한눈에 파악하기 위한 요약표임.</div>
<div class="diagnosis-grid">
  <div class="diagnosis-box"><span>종합 변동률 판단</span><b>선별 거래 국면</b><p>거래 둔화와 가격 조정 압력을 함께 점검함.</p></div>
  <div class="diagnosis-box"><span>매매 변동률 판단</span><b>{{ apt_trade.signal }}</b><p>비교가능 {{ apt_trade.compare_count | fmt_num }}건 기준 상승·보합·하락 비율 확인.</p></div>
  <div class="diagnosis-box"><span>전세 변동률 판단</span><b>{{ apt_rent.jeonse.signal }}</b><p>갱신계약 제외 신규계약 기준 보증금 흐름 점검.</p></div>
  <div class="diagnosis-box"><span>월세 변동률 판단</span><b>{{ apt_rent.monthly_rent.signal }}</b><p>환산월세 기준 실제 주거비 부담 점검.</p></div>
  <div class="diagnosis-box"><span>수요·거래 기반</span><b>인구 감소·세대 유지</b><p>연령대별 이동과 생활권 수요를 함께 검토.</p></div>
  <div class="diagnosis-box"><span>조사상 유의요인</span><b>금리·미분양·공급·잠정치</b><p>최종 판단 시 보조지표로 함께 반영.</p></div>
</div>

<h3>2. 조사 담당자 관점 당월 상승·보합·하락 비율 요약</h3>
<div class="meaning">※ 지표 의미: 도넛 차트는 매매·전세·월세 변동률 조사 판단을 위한 보조지표임. 최종 판단은 TOP 거래사례, 생활권별 흐름, 금리·미분양·공급 지표와 함께 종합함.</div>
{{ charts.core_donuts | safe }}

<h3>표 1. 통계별 기준월 및 공표월 관리표</h3>
{{ period_table | safe }}
<div class="result-box">💡 핵심 판단: {{ text.monthly_core_diagnosis }}</div>

</div></section>
<section class="chapter bg-slate"><h2>Ⅰ. 광주광역시 남구 인구 및 주택 동향 (최근 13개월)</h2><div class="card age-pop-card">
<h3>1. 인구, 세대수 및 인구 이동 동향 통합 종합표</h3>
<div class="meaning">※ 지표 의미: 인구수는 기초 주거수요 규모를, 세대수는 가구 분화 및 주택 수요 단위를, 전입·전출·순이동은 실제 유입·유출 흐름을 나타냄.</div>
{{ population_migration_table | safe }}
{{ charts.population | safe }}
{{ charts.migration | safe }}
<div class="note">💡 {{ text.population_commentary }}</div>
<h3>1-2. 연령대별 주민등록인구 구조 분석</h3>
<div class="meaning">※ 지표 의미: 주민등록인구를 0~9세, 10대, 20대, 30대, 40대, 50대, 60대, 70대 이상으로 재분류하고, 각 월별 인구수와 전월 대비 증감률을 함께 표시함.</div>
<h4>표 1. 최근 13개월 연령대별 주민등록인구 시계열 및 전월 대비 증감률</h4>
{{ age_population_trend_table | safe }}
<h4>표 2. 최신월 연령대별 인구 구조 및 13개월 증감</h4>
{{ age_population.age_rows | as_table }}
{{ charts.age_population_latest | safe }}
{{ charts.age_population_change | safe }}
{{ charts.age_net_combined | safe }}
<div class="note">💡 {{ text.age_migration_commentary }}</div>
<div class="age-move-card"><h3>1-3. 연령대별 전입·전출 및 순이동 분석</h3>
<div class="meaning">※ 지표 의미: 연령대별 이동 흐름을 통해 주거수요의 유입·유출 압력을 점검함.</div>
<h4>표 1. 최근 13개월 연령대별 순이동 시계열</h4>
{{ age_migration_trend_table | safe }}
<div class="note">💡 {{ text.age_migration_commentary }}</div></div>

</div><div class="card">
<h3>2. 예금은행 대출금리 동향: 신규취급액 기준</h3>
<div class="meaning">※ 지표 의미: 주택 매수자, 전세 임차인, 지역 내 자영업·중소기업의 금융비용 부담을 보여주는 핵심 금융 여건 지표임.</div>
{{ rates_standard_table | safe }}
{{ charts.rates | safe }}
<div class="note">💡 {{ text.rate_commentary }}</div>
<h3>3. 미분양 주택 동향(최신 공표월: {{ unsold.latest_month }})</h3>
<div class="meaning">※ 지표 의미: 미분양은 분양시장 부담을, 준공후 미분양은 입주 단계에서도 해소되지 못한 공급 부담을 의미함.</div>
{{ unsold_standard_table | safe }}
{{ charts.unsold | safe }}
<div class="note">💡 {{ text.unsold_commentary }}</div>
<div class="supply-card"><h3>4. 광주광역시 주택공급 동향 및 남구 미래 공급 압력 판단</h3>
<div class="meaning">※ 지표 의미: 인허가·착공·준공은 향후 입주 부담과 전세 수급 압력을 판단하는 중기 공급 지표임.</div>
<h4>표 1. 광주광역시 최근 5개년 공급 파이프라인 총괄</h4>
{{ supply_summary_table | safe }}
<h4>표 2. 광주광역시 최근 5개년 물건유형별 공급 파이프라인 시계열</h4>
{{ supply_by_type_table | safe }}
<h4 class="supply-chart-heading">차트 1. 광주광역시 최근 5개년 공급 단계별 총량 시계열</h4>
<div class="mini-chart-grid total-stage-grid supply-stage-total-grid">{{ charts.supply_permit_total | safe }}{{ charts.supply_start_total | safe }}</div>
<div class="mini-chart-grid total-stage-grid supply-stage-total-grid single-left">{{ charts.supply_complete_total | safe }}</div>

<h4 class="supply-chart-heading">차트 2. 광주광역시 최근 5개년 인허가 물건유형별 분리 시계열</h4>
<div class="mini-chart-grid supply-type-grid">{{ charts.supply_permit_아파트 | safe }}{{ charts.supply_permit_연립 | safe }}</div>
<div class="mini-chart-grid supply-type-grid">{{ charts.supply_permit_다세대 | safe }}{{ charts.supply_permit_단독 | safe }}</div>
<div class="mini-chart-grid supply-type-grid single-left">{{ charts.supply_permit_다가구 | safe }}</div>

<h4 class="supply-chart-heading">차트 3. 광주광역시 최근 5개년 착공 물건유형별 분리 시계열</h4>
<div class="mini-chart-grid supply-type-grid">{{ charts.supply_start_아파트 | safe }}{{ charts.supply_start_연립 | safe }}</div>
<div class="mini-chart-grid supply-type-grid">{{ charts.supply_start_다세대 | safe }}{{ charts.supply_start_단독 | safe }}</div>
<div class="mini-chart-grid supply-type-grid single-left">{{ charts.supply_start_다가구 | safe }}</div>

<h4 class="supply-chart-heading">차트 4. 광주광역시 최근 5개년 준공 물건유형별 분리 시계열</h4>
<div class="mini-chart-grid supply-type-grid">{{ charts.supply_complete_아파트 | safe }}{{ charts.supply_complete_연립 | safe }}</div>
<div class="mini-chart-grid supply-type-grid">{{ charts.supply_complete_다세대 | safe }}{{ charts.supply_complete_단독 | safe }}</div>
<div class="mini-chart-grid supply-type-grid single-left">{{ charts.supply_complete_다가구 | safe }}</div>
<div class="note">💡 {{ text.supply_commentary }}</div>


</div></section>
<section class="chapter bg-rose"><h2>Ⅲ. 매매 주택 거래 현황 분석</h2><div class="card">
<h3>1. 아파트 매매 월간 거래량 및 직전거래 대비 가격 변동 비중</h3>
<div class="meaning">※ 지표 의미: 기준월 매매거래와 직전거래 비교 가능 사례의 상승·보합·하락 비율을 함께 점검함.</div>
{{ apt_trade.monthly_rows | as_table }}
{{ charts.apt_monthly | safe }}
{{ charts.apt_strength | safe }}
<div class="note">💡 {{ text.apt_trade_commentary }}</div>
<h3>[차트 3] 당월 읍면동별 매매 시장강도</h3>{{ charts.apt_dong | safe }}
<h3>2. 아파트 매매 읍면동별 세부 현황</h3>{{ apt_trade.dong_rows | as_table }}{{ charts.apt_dong_detail | safe }}
<h3>3. 아파트 매매 거래금액별 세부 현황</h3>{{ apt_trade.price_band_rows | as_table }}{{ charts.apt_price_band | safe }}
<h3>4. 아파트 매매 연식별 세부 현황</h3>{{ apt_trade.building_age_rows | as_table }}{{ charts.apt_age_band | safe }}
<h3>5. 아파트 매매 평형별 세부 현황</h3>{{ apt_trade.area_band_rows | as_table }}{{ charts.apt_area_share | safe }}
<h3>6. 아파트 매매 거래단지 TOP10</h3>{{ apt_trade.top_complex_rows | as_table }}
<h3>7. 비아파트 최근 13개월 매매 거래 추이</h3>
{{ non_apt_sale_type_table | safe }}
{{ charts.non_apt_sale | safe }}
<div class="note">💡 {{ text.non_apt_sale_commentary }}</div>
<h3>8. 아파트 분양권전매 최근 13개월 거래 추이</h3>
{{ presale.monthly_rows | as_table }}
{{ charts.presale | safe }}
<div class="note">💡 {{ text.presale_commentary }}</div>
<div class="card top-case-card">
<h3>6-1. 아파트 매매 당월 직전거래 대비 TOP 거래사례</h3>
<p class="meaning">※ 지표 의미: 당월 아파트 매매 실거래를 동일 단지·동일 전용면적의 직전 매매거래와 비교하여 상승폭·하락폭·거래금액 상위 사례를 확인하는 실거래 사례 지표임. 층 정보는 현재 거래와 직전거래의 가격 차이를 해석하는 보조지표로 함께 표시함.</p>
<div class="top-case-grid">
{{ top_case_list(apt_trade.top_up_rows, "당월 상승거래 TOP10", "매매", "rise") | safe }}
{{ top_case_list(apt_trade.top_down_rows, "당월 하락거래 TOP10", "매매", "fall") | safe }}
{{ top_case_list(apt_trade.top_price_rows, "거래금액 상위 TOP10", "매매", "auto") | safe }}
</div>
</div>
<div class="note">💡 {{ text.top_case_commentary }}</div>


</div></section>
<section class="chapter bg-emerald"><h2>Ⅳ. 전세 주택 거래 현황 분석</h2><div class="card">
<h3>1. 아파트 전세 월간 거래량 및 직전 신규계약 대비 가격 변동 비중</h3>
{{ apt_rent.jeonse.monthly_rows | as_table }}
{{ charts.apt_jeonse_monthly | safe }}
{{ charts.apt_jeonse_strength | safe }}
<div class="note">💡 {{ text.apt_jeonse_commentary }}</div>
<h3>2. 아파트 전세 읍면동별 세부 현황</h3>{{ charts.apt_jeonse_market_dong | safe }}{{ apt_rent.jeonse.dong_rows | as_table }}{{ charts.apt_jeonse_dong | safe }}
<h3>3. 아파트 전세 보증금 구간별 세부 현황</h3>{{ apt_rent.jeonse.deposit_band_rows | as_table }}{{ charts.apt_jeonse_deposit | safe }}
<h3>4. 아파트 전세 연식별 세부 현황</h3>{{ apt_rent.jeonse.building_age_rows | as_table }}{{ charts.apt_jeonse_age | safe }}
<h3>5. 아파트 전세 평형별 세부 현황</h3>{{ apt_rent.jeonse.area_band_rows | as_table }}{{ charts.apt_jeonse_area | safe }}
<h3>6. 아파트 전세 신규계약 거래단지 TOP10</h3>{{ apt_rent.jeonse.top_complex_rows | as_table }}
<h3>7. 비아파트 최근 13개월 전세 거래 추이</h3>
{{ non_apt_jeonse_type_table | safe }}
{{ charts.non_apt_jeonse | safe }}
<div class="note">💡 {{ text.non_apt_jeonse_commentary }}</div>
<div class="card top-case-card">
<h3>5-1. 아파트 전세 당월 직전 신규계약 대비 TOP 거래사례</h3>
<p class="meaning">※ 지표 의미: 당월 아파트 전세 신규계약을 동일 단지·동일 전용면적의 직전 신규 전세계약과 비교하여 보증금 상승·하락·상위 사례를 확인하는 실거래 사례 지표임. 갱신계약은 가격 비교에서 제외하며, 현재층과 직전층을 함께 표시함.</p>
<div class="top-case-grid">
{{ top_case_list(apt_rent.jeonse.top_up_rows, "보증금 상승 TOP10", "전세", "rise") | safe }}
{{ top_case_list(apt_rent.jeonse.top_down_rows, "보증금 하락 TOP10", "전세", "fall") | safe }}
{{ top_case_list(apt_rent.jeonse.top_value_rows, "보증금 상위 TOP10", "전세", "auto") | safe }}
</div>
</div>


</div></section>
<section class="chapter bg-blue"><h2>Ⅴ. 월세 주택 거래 현황 분석</h2><div class="card">
<h3>1. 아파트 월세 월간 거래량 및 직전 신규계약 대비 가격 변동 비중</h3>
{{ apt_rent.monthly_rent.monthly_rows | as_table }}
{{ charts.apt_rent_monthly | safe }}
{{ charts.apt_rent_strength | safe }}
<div class="note">💡 {{ text.apt_rent_commentary }}</div>
<div class="conversion-card"><h3>2. 월세 환산 기준 요약표</h3>
{{ conversion_rate_table | safe }}
<div class="note">💡 {{ text.conversion_rate_commentary }}</div></div>
<h3>3. 아파트 월세 읍면동별 세부 현황</h3>{{ charts.apt_rent_market_dong | safe }}{{ apt_rent.monthly_rent.dong_rows | as_table }}{{ charts.apt_rent_dong | safe }}
<h3>4. 아파트 월세 보증금·월세·환산월세 구간별 세부 현황</h3><h4>보증금 구간별</h4>{{ apt_rent.monthly_rent.deposit_band_rows | as_table }}{{ charts.apt_rent_deposit | safe }}<h4>월세 구간별</h4>{{ apt_rent.monthly_rent.monthly_band_rows | as_table }}{{ charts.apt_rent_monthly_band | safe }}
<h3>5. 아파트 월세 연식별 세부 현황</h3>{{ apt_rent.monthly_rent.building_age_rows | as_table }}{{ charts.apt_rent_age | safe }}
<h3>6. 아파트 월세 평형별 세부 현황</h3>{{ apt_rent.monthly_rent.area_band_rows | as_table }}{{ charts.apt_rent_area | safe }}
<h3>7. 아파트 월세 신규계약 거래단지 TOP10</h3>{{ apt_rent.monthly_rent.top_complex_rows | as_table }}
<h3>8. 비아파트 최근 13개월 월세 거래 추이</h3>
{{ non_apt_rent_type_table | safe }}
{{ charts.non_apt_rent | safe }}
<div class="note">💡 {{ text.non_apt_rent_commentary }}</div>
<div class="card top-case-card">
<h3 id="rent-top10-anchor">5-1. 아파트 월세 당월 직전 신규계약 대비 TOP 거래사례</h3>
<p class="meaning">※ 지표 의미: 당월 아파트 월세 신규계약을 동일 단지·동일 전용면적의 직전 신규 월세계약과 비교하여 환산월세 상승·하락·상위 사례를 확인하는 실거래 사례 지표임. 보증금과 월세를 함께 반영한 환산월세 기준이며, 현재층과 직전층을 함께 표시함.</p>
<div class="top-case-grid">
{{ top_case_list(apt_rent.monthly_rent.top_up_rows, "환산월세 상승 TOP10", "월세", "rise") | safe }}
{{ top_case_list(apt_rent.monthly_rent.top_down_rows, "환산월세 하락 TOP10", "월세", "fall") | safe }}
{{ top_case_list(apt_rent.monthly_rent.top_value_rows, "환산월세 상위 TOP10", "월세", "auto") | safe }}
</div>
</div>


</div></section>
<section id="chapter-final-summary" class="chapter bg-indigo final-summary-chapter"><h2>Ⅵ. 종합 시황 서술 보고서</h2><div class="card">
<div class="notice">※ Ⅵ장에는 표를 추가하지 않고, 앞선 분석표에서 도출된 핵심 해석만 문단에 반영함.</div>
<div class="summary-section-box summary-sale-box">
<h3>1. 매매시장 종합</h3>
<h4>1) 매매 종합</h4><div class="summary-text">{{ text.sale_market_summary }}</div>
<h4>2) 아파트 매매</h4><div class="summary-text">{{ text.apt_sale_summary }}</div>{{ major_case_line(apt_trade, "매매") | safe }}
<h4>3) 연립·다세대 매매</h4><div class="summary-text">{{ text.rowhouse_multiplex_sale_summary }}</div>
<h4>4) 단독·다가구 매매</h4><div class="summary-text">{{ text.detached_multifamily_sale_summary }}</div>
</div>
<div class="summary-section-box summary-jeonse-box">
<h3>2. 전세시장 종합</h3>
<h4>5) 전세 종합</h4><div class="summary-text">{{ text.jeonse_market_summary }}</div>
<h4>6) 아파트 전세</h4><div class="summary-text">{{ text.apt_jeonse_summary }}</div>{{ major_case_line(apt_rent.jeonse, "전세") | safe }}
<h4>7) 연립·다세대 전세</h4><div class="summary-text">{{ text.rowhouse_multiplex_jeonse_summary }}</div>
<h4>8) 단독·다가구 전세</h4><div class="summary-text">{{ text.detached_multifamily_jeonse_summary }}</div>
</div>
<div class="summary-section-box summary-rent-box">
<h3>3. 월세시장 종합</h3>
<h4>9) 월세 종합</h4><div class="summary-text">{{ text.rent_market_summary }}</div>
<h4>10) 아파트 월세</h4><div class="summary-text">{{ text.apt_rent_summary }}</div>{{ major_case_line(apt_rent.monthly_rent, "월세") | safe }}
<h4>11) 연립·다세대 월세</h4><div class="summary-text">{{ text.rowhouse_multiplex_rent_summary }}</div>
<h4>12) 단독·다가구 월세</h4><div class="summary-text">{{ text.detached_multifamily_rent_summary }}</div>
</div>
<h3>4. 종합 판단</h3>
<div class="diagnosis comprehensive-summary-box">{{ text.comprehensive_summary }}</div>

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
    rows = age_population.get("monthly_rows", []) if isinstance(age_population, dict) else []
    if not rows:
        return "<p class='muted'>연령대별 주민등록인구 원자료 없음</p>"
    age_cols = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
    rows13 = rows[-13:]
    parts = []
    parts.append("<div class='table-wrap age-pop-trend-table-wrap'>")
    parts.append("<table class='market-strength-table compact-wide-table time-series-table age-pop-trend-table'>")
    parts.append("<thead><tr><th>기준월</th>" + "".join(f"<th>{c}</th>" for c in age_cols) + "</tr></thead><tbody>")
    prev = None
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
                    cell = f"{num_s} <span class='age-rate neutral'>(-)</span>"
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
    pop_rows = data.get("population", {}).get("table_rows", []) or []
    mig_rows = data.get("migration", {}).get("table_rows", []) or []
    mig_by_month = {str(r.get("month") or r.get("기준월")): r for r in mig_rows}
    mig_latest = data.get("migration", {}).get("latest_month") or "-"
    rows = []
    prev_pop = None
    prev_hh = None
    for r in pop_rows[-13:]:
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
    rows = rates.get("table_rows", []) if isinstance(rates, dict) else []
    if not rows:
        return "<p class='muted'>금리 원자료 없음</p>"

    rate_cols = [
        ("mortgage_rate", "주택담보대출"),
        ("jeonse_rate", "전세자금대출"),
        ("sme_rate", "중소기업대출"),
        ("business_rate", "기업일반자금대출"),
    ]
    rows13 = rows[-13:]
    html = []
    html.append("<div class='table-wrap rates-standard-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table time-series-table rates-standard-table'>")
    html.append("<thead><tr><th>기준월</th>")
    for _, label in rate_cols:
        html.append(f"<th>{label}</th><th>{label} 전월 대비 증감</th>")
    html.append("</tr></thead><tbody>")

    prev = None
    for idx, r in enumerate(rows13):
        tr_cls = " class='latest-row'" if idx == len(rows13) - 1 else ""
        html.append(f"<tr{tr_cls}><td>{html_value(r.get('month') or r.get('기준월'))}</td>")
        for key, _label in rate_cols:
            val = r.get(key)
            prev_val = prev.get(key) if isinstance(prev, dict) else None
            diff = None
            try:
                if val is not None and prev_val is not None and not pd.isna(val) and not pd.isna(prev_val):
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
    rows = unsold.get("table_rows", []) if isinstance(unsold, dict) else []
    if not rows:
        return "<p class='muted'>미분양 원자료 없음</p>"

    out = []
    prev_unsold = None
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

        if prev_unsold is None or u_num is None:
            u_diff = None
        else:
            u_diff = u_num - prev_unsold
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
            "기준연도": y,
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
        [_supply_year_value(r) for r in ordered],
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
        norm.append({"year": str(y), "stage": st, "type": typ, "value": val})
    if not norm:
        return "<p class='muted'>물건유형별 공급 원자료 없음</p>"
    df = pd.DataFrame(norm)
    years = sorted(df["year"].unique())[-5:]
    grouped = df[df["year"].isin(years)].groupby(["stage", "type", "year"], as_index=False)["value"].sum()
    html = []
    html.append("<div class='supply-audit-note'>※ 수치 기준: 업로드 기준 이미지와 동일하게 과거연도는 12월 누계, 당해연도는 최신 누계월 값을 사용함. 인허가·착공·준공은 주택유형별 원자료의 아파트·연립·다세대·단독·다가구 행만 표시하며, 다가구는 가구수 기준임. 분양 항목은 공급 파이프라인 표와 차트에서 제외함.</div>")
    html.append("<div class='table-wrap supply-type-pipeline-table-wrap'>")
    html.append("<table class='market-strength-table compact-wide-table supply-type-pipeline-table'>")
    html.append("<thead><tr><th>공급단계</th><th>물건유형</th>" + "".join(f"<th>{y}</th>" for y in years) + "<th>5개년 합계</th></tr></thead><tbody>")
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
    rows = [
        {"지표": "실거래", "적용 기준월": meta.get("base_month"), "자료 성격": "계약일 기준 잠정치", "보고서 해석 기준": "신고기한 미도래로 향후 거래건수 보정 가능"},
        {"지표": "인구·세대수", "적용 기준월": data.get("population", {}).get("latest_month"), "자료 성격": "주민등록 기준", "보고서 해석 기준": "기초 주거수요와 가구 분화 흐름 판단"},
        {"지표": "인구이동", "적용 기준월": data.get("migration", {}).get("latest_month"), "자료 성격": "최신 공표월", "보고서 해석 기준": "전입·전출 및 연령대별 이동 흐름 판단"},
        {"지표": "예금은행 대출금리", "적용 기준월": data.get("rates", {}).get("latest_month"), "자료 성격": "신규취급액 기준", "보고서 해석 기준": "주거금융 및 사업체 금융비용 부담 판단"},
        {"지표": "미분양·준공후 미분양", "적용 기준월": data.get("unsold", {}).get("latest_month"), "자료 성격": "최신 공표월", "보고서 해석 기준": "분양시장 부담 및 입주 단계 공급 부담 판단"},
        {"지표": "주택공급", "적용 기준월": "2021~2025", "자료 성격": "최근 5개년 연간 실적", "보고서 해석 기준": "인허가·착공·준공을 통한 중기 공급 압력 판단"},
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

def major_case_line_html(section_obj: Dict[str, Any], value_label: str) -> str:
    """Ⅵ장 아파트 매매·전세·월세 문단 하단에 당월 상승/하락 분위기에 맞는 주요 거래사례 1건 표시."""
    if not isinstance(section_obj, dict):
        return ""
    try:
        up = float(section_obj.get("up_ratio") or 0)
        down = float(section_obj.get("down_ratio") or 0)
    except Exception:
        up, down = 0, 0
    if down > up and section_obj.get("top_down_rows"):
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

    name = html.escape(str(row.get("apt_name") or row.get("_group_name") or "-"))
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
        f"<div class='major-case-line summary-major-case'>* 주요 거래사례: "
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


def render_html(data: Dict[str, Any], commentary: Dict[str, str], validation_rows: List[Dict[str, str]]) -> str:
    commentary = apply_commentary_semantic_highlights(commentary)
    env = Environment(autoescape=False)
    env.filters["fmt_num"] = format_num
    env.filters["fmt_float"] = format_float
    env.filters["as_table"] = table_rows_html
    env.filters["as_cards"] = top_cards_html
    env.globals["top_case_list"] = top_case_list_html
    env.globals["major_case_line"] = major_case_line_html
    tpl = env.from_string(HTML_TEMPLATE)
    conv = data.get("conversion_rate", {})
    conversion_rate_table = table_rows_html([{
        "구분": "월세 환산 기준",
        "내용": "아파트 신규 전세·신규 월세 동일 주택형 매칭 중위 전환율 적용"
    }, {
        "구분": "산출 기간",
        "내용": conv.get("period", "-")
    }, {
        "구분": "동일 주택형 매칭키",
        "내용": conv.get("matching_key", "-")
    }, {
        "구분": "동일 주택형 매칭건수",
        "내용": f"{format_num(conv.get('matched_count'))}건"
    }, {
        "구분": "적용 전환율",
        "내용": f"{format_float(conv.get('applied_rate'), 2)}%"
    }, {
        "구분": "전환율 산식",
        "내용": conv.get("formula", "-")
    }, {
        "구분": "환산월세 산식",
        "내용": conv.get("converted_rent_formula", "-")
    }], max_rows=20)
    html = tpl.render(
        meta=data["meta"], cards=data["summary_cards"], warnings=data.get("data_quality_warnings", []),
        population=data["population"], age_population=data.get("age_population", {}), migration=data["migration"], rates=data["rates"], unsold=data["unsold"],
        price_index=data["price_index"], supply=data["supply"], apt_trade=data["apt_trade"], apt_rent=data["apt_rent"],
        non_apt_sale=data["non_apt_sale"], non_apt_rent=data["non_apt_rent"], presale=data["presale"],
        conversion_rate_table=conversion_rate_table,
        conversion_rate=data.get("conversion_rate", {}),
        period_table=period_management_table(data),
        rates_standard_table=rates_table_html(data.get("rates", {})),
        unsold_standard_table=unsold_table_html(data.get("unsold", {})),
        supply_summary_table=supply_summary_table_html(data.get("supply", {})),
        supply_by_type_table=supply_by_type_table_html(data.get("supply", {})),
        non_apt_sale_type_table=non_apt_type_count_table_html(data.get("non_apt_sale", {}).get("monthly_by_type_rows", [])),
        non_apt_jeonse_type_table=non_apt_type_count_table_html(data.get("non_apt_rent", {}).get("jeonse", {}).get("monthly_by_type_rows", [])),
        non_apt_rent_type_table=non_apt_type_count_table_html(data.get("non_apt_rent", {}).get("monthly_rent", {}).get("monthly_by_type_rows", [])),
        population_migration_table=population_migration_table_html(data),
        age_population_trend_table=age_population_trend_table_html(data.get("age_population", {})),
        age_migration_trend_table=age_migration_trend_table_html(data.get("migration", {})),
        charts=make_chart_html(data),
        text=commentary, validation_rows=validation_rows
    )
    return ensure_template_v58_structure(html)


def validate_report(html: str, data: Dict[str, Any], commentary: Dict[str, str]) -> List[Dict[str, str]]:
    rows = []
    text = BeautifulSoup(html, "html.parser").get_text("\n")
    def add(item, ok, note=""):
        rows.append({"검수항목": item, "결과": "정상" if ok else "확인필요", "비고": note})
    for sec in ["Ⅰ. 인구 및 주택 동향", "Ⅲ. 매매 주택 거래 현황 분석", "Ⅳ. 전세 주택 거래 현황 분석", "Ⅴ. 월세 주택 거래 현황 분석", "Ⅵ. 종합 시황 서술 보고서"]:
        add(f"필수 섹션 존재: {sec}", sec in text)
    for k in COMMENTARY_KEYS:
        add(f"해설 key 존재: {k}", bool(commentary.get(k)))
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
    add("월세·종합 후반부 출력 섹션 포함", all(x in html for x in ["Ⅴ. 월세 주택 거래 현황 분석", "5-1. 아파트 월세 당월 직전 신규계약 대비 TOP 거래사례", "Ⅵ. 종합 시황 서술 보고서"]))
    add("표지 KPI print 세로형 CSS 포함", "grid-template-columns:1fr" in html and "v5.8.38 PDF 출력 안정화 보정" in html)
    add("상위 TOP10 중립형 적용", "case-top" in html)
    add("v5.8.38 차트 중복 방지", "평균 가격변동률" in html and "상승·하락 비율" in html)
    add("비아파트 물건유형별 차트", "비아파트 물건유형별 최근 13개월 매매 거래 추이(연립·다세대·단독·다가구)" in html and "연립·다세대" in html and "단독·다가구" in html)
    add("인구·세대수 차트 보정", "population-mix-chart" in html)
    add("비아파트 4개 유형별 차트", all(x in html for x in ["연립", "다세대", "단독", "다가구"]) and "비아파트 물건유형별 최근 13개월 매매 거래 추이" in html)
    add("인구이동 통합표 포함", all(x in html for x in ["전입", "전출", "순이동"]))
    add("연령대 누적순이동 가로 막대차트", "최근 13개월 누적순이동 연령별 비교" in html and "indexAxis" in html and "청년층 순이동" not in html and "중장년·고령층 순이동" not in html)
    add("연령별 차트 컬러 일치", all(x in html for x in ["#5BC0EB", "#3A86FF", "#8338EC", "#2A9D8F", "#F4A261", "#E76F51", "#577590", "#D62828"]))
    add("인구이동 미공표 표기", "not-published" in html and "이동자수 최신 공표월" in html)
    add("인구통합표 가로잘림 방지", "population-migration-table" in html and "주요 순유출·순유입 연령대" in html)
    add("월간 가격변동 차트 꺾은선 제거", "평균 변동률 꺾은선은 제거" in html)
    add("월간 시장강도 차트 상승·하락 꺾은선", "상승·하락 비율만 꺾은선" in html)
    add("상승보합하락 표 이모지", all(x in html for x in ["🔴 상승 비율", "⚪ 보합 비율", "🔵 하락 비율"]))
    add("상승보합하락 표 컬러 class", all(x in html for x in ["dir-up", "dir-flat", "dir-down"]))
    add("연령대별 인구 전월 대비 증감률 표", "최근 13개월 연령대별 주민등록인구 시계열 및 전월 대비 증감률" in html and "age-pop-trend-table" in html)
    add("차트 내부 제목 비활성화", "\"title\": {\"display\": false}" in html or '"title": {"display": false}' in html)
    add("가격지수 별도 섹션 미출력", "주택가격지수</h3>" not in html)
    add("공급 물건유형별 표 포함", "물건유형별 공급 파이프라인 시계열" in html)
    add("총 table 40개 기준", html.count("<table") == 40, f"table={html.count('<table')}")
    add("비아파트 보강표 6개 포함", all(x in html for x in ["연립·다세대 매매", "단독·다가구 매매", "연립·다세대 전세", "단독·다가구 전세", "연립·다세대 월세", "단독·다가구 월세"]))
    add("conversion-card 실제 출력", "conversion-card" in html)
    add("총 table 40개 기준", html.count("<table") == 40, f"table={html.count('<table')}")

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
    add("월세·종합 후반부 출력 섹션 포함", all(x in html for x in ["Ⅴ. 월세 주택 거래 현황 분석", "5-1. 아파트 월세 당월 직전 신규계약 대비 TOP 거래사례", "Ⅵ. 종합 시황 서술 보고서"]))
    add("표지 KPI print 세로형 CSS 포함", "grid-template-columns:1fr" in html and "v5.8.38 PDF 출력 안정화 보정" in html)
    add("상위 TOP10 중립형 적용", "case-top" in html)
    add("v5.8.38 차트 중복 방지", "평균 가격변동률" in html and "상승·하락 비율" in html)
    add("비아파트 물건유형별 차트", "비아파트 물건유형별 최근 13개월 매매 거래 추이(연립·다세대·단독·다가구)" in html and "연립·다세대" in html and "단독·다가구" in html)
    add("인구·세대수 차트 보정", "population-mix-chart" in html)
    add("비아파트 4개 유형별 차트", all(x in html for x in ["연립", "다세대", "단독", "다가구"]) and "비아파트 물건유형별 최근 13개월 매매 거래 추이" in html)
    add("인구이동 통합표 포함", all(x in html for x in ["전입", "전출", "순이동"]))
    add("연령대 누적순이동 가로 막대차트", "최근 13개월 누적순이동 연령별 비교" in html and "indexAxis" in html and "청년층 순이동" not in html and "중장년·고령층 순이동" not in html)
    add("연령별 차트 컬러 일치", all(x in html for x in ["#5BC0EB", "#3A86FF", "#8338EC", "#2A9D8F", "#F4A261", "#E76F51", "#577590", "#D62828"]))
    add("인구이동 미공표 표기", "not-published" in html and "이동자수 최신 공표월" in html)
    add("인구통합표 가로잘림 방지", "population-migration-table" in html and "주요 순유출·순유입 연령대" in html)
    add("월간 가격변동 차트 꺾은선 제거", "평균 변동률 꺾은선은 제거" in html)
    add("월간 시장강도 차트 상승·하락 꺾은선", "상승·하락 비율만 꺾은선" in html)
    add("상승보합하락 표 이모지", all(x in html for x in ["🔴 상승 비율", "⚪ 보합 비율", "🔵 하락 비율"]))
    add("상승보합하락 표 컬러 class", all(x in html for x in ["dir-up", "dir-flat", "dir-down"]))
    add("연령대별 인구 전월 대비 증감률 표", "최근 13개월 연령대별 주민등록인구 시계열 및 전월 대비 증감률" in html and "age-pop-trend-table" in html)
    add("차트 내부 제목 비활성화", "\"title\": {\"display\": false}" in html or '"title": {"display": false}' in html)
    add("가격지수 별도 섹션 미출력", "주택가격지수</h3>" not in html)
    add("공급 물건유형별 표 포함", "물건유형별 공급 파이프라인 시계열" in html)
    add("총 table 40개 기준", html.count("<table") == 40, f"table={html.count('<table')}")
    add("비아파트 보강표 6개 포함", all(x in html for x in ["연립·다세대 매매", "단독·다가구 매매", "연립·다세대 전세", "단독·다가구 전세", "연립·다세대 월세", "단독·다가구 월세"]))
    add("conversion-card 실제 출력", "conversion-card" in html)
    add("총 table 40개 기준", html.count("<table") == 40, f"table={html.count('<table')}")

    add("table-wrap 35개 이상", html.count("table-wrap") >= 35, f"table-wrap={html.count('table-wrap')}")
    add("차트 자료 없음 문구 최소화", html.count("차트 자료 없음") <= 3, f"차트 자료 없음={html.count('차트 자료 없음')}")
    add("template_v58 핵심 클래스 포함", all(cls in html for cls in ["hero", "kpis", "chapter", "card", "chart-box", "note"]))
    add("실제 canvas 차트 30개 이상 포함", html.count("<canvas") >= 30, f"canvas={html.count('<canvas')}")
    add("영문 내부 컬럼명 최소화", "trade_count" not in text and "compare_count" not in text and "population" not in text)
    # 20개 자료 사용 검수
    used = data.get("input_files", {})
    for bucket in ["population", "migration", "rates", "unsold", "supply", "price_index_sale", "price_index_jeonse", "price_index_rent", "apt_trade", "apt_rent", "non_apt_sale", "non_apt_rent", "presale"]:
        add(f"파일 분류 확인: {bucket}", len(used.get(bucket, [])) > 0, ", ".join(used.get(bucket, [])))
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
        "5-1. 아파트 월세 당월 직전 신규계약 대비 TOP 거래사례",
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

def write_outputs(data: Dict[str, Any], commentary: Dict[str, str], suffix: str = "") -> Dict[str, Any]:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_region = re.sub(r"[^가-힣A-Za-z0-9_]+", "_", data["meta"]["region"])
    stem = f"report_{safe_region}_{data['meta']['base_month'].replace('-','')}_{ts}{suffix}"
    html0 = render_html(data, commentary, [])
    validation = validate_report(html0, data, commentary)
    html = render_html(data, commentary, validation)
    data_path = OUTPUT_DIR / f"data_pack_{stem}.json"
    comm_path = OUTPUT_DIR / f"commentary_pack_{stem}.json"
    html_path = OUTPUT_DIR / f"{stem}.html"
    val_path = OUTPUT_DIR / f"validation_{stem}.xlsx"
    save_json(data, data_path)
    save_json(commentary, comm_path)
    html_path.write_text(html, encoding="utf-8")
    pd.DataFrame(validation).to_excel(val_path, index=False)
    return {"data_pack": data, "commentary": commentary, "data_pack_path": data_path, "commentary_pack_path": comm_path, "html_path": html_path, "validation_excel_path": val_path, "validation_rows": validation, "validation_df": pd.DataFrame(validation)}


def save_prompt(data: Dict[str, Any], prompt: str) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_region = re.sub(r"[^가-힣A-Za-z0-9_]+", "_", data["meta"]["region"])
    path = PROMPT_DIR / f"llm_prompt_{safe_region}_{data['meta']['base_month'].replace('-','')}_{ts}.txt"
    path.write_text(prompt, encoding="utf-8")
    return path


def run_pipeline(paths: List[Path], region: str, base_month: str, version: str) -> Dict[str, Any]:
    progress("보고서 생성 파이프라인 시작", 0, 100)
    data = build_data_pack(paths, region, base_month, version)
    progress("규칙 기반 1차 해설 생성 중", 85, 100)
    commentary = base_commentary(data)
    progress("HTML 및 검수표 생성 중", 92, 100)
    result = write_outputs(data, commentary)
    prompt = build_llm_prompt(data)
    prompt_path = save_prompt(data, prompt)
    result["prompt_text"], result["prompt_path"] = prompt, prompt_path
    progress("보고서 생성 완료", 100, 100)
    return result


def regenerate_with_commentary(data: Dict[str, Any], commentary: Dict[str, str]) -> Dict[str, Any]:
    return write_outputs(data, commentary, suffix="_llm")


# =========================================================
# 9. Tkinter UI
# =========================================================

class JsonPasteDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("LLM 해설 JSON 붙여넣기")
        self.geometry("920x620")
        self.result_text = None
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="LLM이 출력한 JSON을 붙여넣으세요.", font=("맑은 고딕", 10, "bold")).pack(anchor="w")
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
    def __init__(self):
        super().__init__()
        self.title("월간 주택동향보고서 HTML 최종보고서 생성기 v5.8.38.1.1")
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
        ttk.Label(root, text="월간 주택동향보고서 HTML 최종보고서 생성기 v5.8.38.1.1", style="Header.TLabel").pack(anchor="w")
        ttk.Label(root, text="HTML 최종보고서 전용 · template_v58 실제 차트 슬롯만 사용", style="Small.TLabel").pack(anchor="w", pady=(2,12))
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
        c = classify_files(self.selected_files)
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
            result = run_pipeline(self.selected_files, region, month, version)
            self.after(0, lambda res=result: self._success(res))
        except Exception as e:
            tb=traceback.format_exc(); path=LOG_DIR/f"error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"; path.write_text(tb, encoding="utf-8")
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
            self.current_commentary=merge_commentary(base_commentary(self.current_data), new)
            messagebox.showinfo("적용 준비", "LLM 해설 JSON을 읽었습니다. 재생성 버튼을 누르세요.")
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
            self.current_commentary = merge_commentary(base_commentary(self.current_data), new)
            result = regenerate_with_commentary(self.current_data, self.current_commentary)
            self.last_result = result
            self.refresh_validation(result["validation_rows"])
            messagebox.showinfo(
                "HTML 재생성 완료",
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
        result=regenerate_with_commentary(self.current_data, self.current_commentary)
        self.last_result=result; self.refresh_validation(result["validation_rows"])
        messagebox.showinfo("재생성 완료", f"HTML: {result['html_path']}")
    def open_html(self):
        if self.last_result: webbrowser.open(self.last_result["html_path"].as_uri())
    def open_validation(self):
        if self.last_result: os.startfile(str(self.last_result["validation_excel_path"]))
    def open_output(self):
        os.startfile(str(OUTPUT_DIR))


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

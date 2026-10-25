# -*- coding: utf-8 -*-
"""
KOSIS 통계 버튼형 검증 수집기 - 최종본 (파일명 YYYY_MM 통일)
주요 기능
- KOSIS API KEY 자동 저장/불러오기
- 법정동 코드 엑셀 기반 시도/시군구 콤보박스
- 통계별로 시군구 5자리 / 시도 2자리 자동 적용
- 월간/연간/주간 통계 지원
- 월간: 1개월씩 수집 후 연도별 엑셀 저장
- 주간: 연도 범위 조회 후 종료 기준월까지의 최근 발표 주차만 반영
- requests 실패 시 urllib fallback
- err 30(데이터 없음)은 빈 데이터 처리

지역코드 적용 원칙
- 시군구 단위 통계: REG 5자리 사용
- 시도 단위 통계: REG 앞 2자리 사용
"""

import json
import os
import queue
import shutil
import random
import re
import ssl
import sys
import threading
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from dateutil.relativedelta import relativedelta
except ImportError:
    raise SystemExit("python-dateutil 패키지가 필요합니다. pip install python-dateutil")

try:
    import xlrd  # noqa: F401
except ImportError:
    xlrd = None


APP_TITLE = "부동산 AI 동향 리포터 v1.71 | 제작자: 한국부동산원 장규태"
APP_VERSION = "v122 오피스텔 HTML 보고서 오류 보정"
APP_AUTHOR = "한국부동산원 장규태"
APP_UPDATED = "2026-05-08"
BG_COLOR = "#F5F7FB"
CARD_COLOR = "#FFFFFF"
TITLE_COLOR = "#0F172A"
TEXT_SUB = "#64748B"
BORDER_COLOR = "#E2E8F0"
PRIMARY_BTN = "#2563EB"
PRIMARY_BTN_ACTIVE = "#1D4ED8"
SECONDARY_BTN = "#F1F5F9"
SECONDARY_BTN_ACTIVE = "#E2E8F0"
DASHBOARD_NAV = "#0F2742"
DASHBOARD_NAV_SUB = "#D8E3F0"
SUCCESS_COLOR = PRIMARY_BTN
WARNING_COLOR = "#D97706"
BASE_URL = "https://kosis.kr/openapi/Param/statisticsParameterData.do"
STATISTICS_LIST_URL = "https://kosis.kr/openapi/statisticsList.do"
META_URL = "https://kosis.kr/openapi/statisticsData.do"
DEFAULT_TIMEOUT = (10, 60)
DEFAULT_WORKERS = 3
REQUEST_SLEEP = 2.5
KOSIS_RETRY_ATTEMPTS = 6
KOSIS_RETRY_BACKOFF = 2.0
def user_global_config_dir() -> Path:
    """실행 폴더가 매번 달라도 API 키와 사용자 설정을 유지하기 위한 전역 설정 폴더."""
    root = Path(os.environ.get("APPDATA", str(Path.home()))) / "RealEstateApiReportUnified"
    try:
        root.mkdir(parents=True, exist_ok=True)
    except Exception:
        root = Path.home() / ".realestate_api_report_unified"
        root.mkdir(parents=True, exist_ok=True)
    return root

CONFIG_PATH = user_global_config_dir() / "kosis_app_config.json"
DEFAULT_API_KEY = "__KOSIS_API_KEY_REDACTED__"
DEFAULT_RONE_API_KEY = "__API_KEY_REDACTED__"  # 한국부동산원 R-ONE OpenAPI 인증키(사용자 발급). 화면에는 별표로만 표시

def app_base_path() -> Path:
    """실행 기준 폴더를 반환합니다.

    - 일반 Python 실행: .py 파일이 있는 폴더
    - PyInstaller EXE 실행: EXE가 있는 폴더
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def bundled_resource_path(filename: str) -> Path:
    """PyInstaller EXE/통합실행기/작업폴더 기준 리소스 경로를 순차 탐색합니다.
    통합 실행기에서는 내부 모듈이 _integrated_modules_api_report 폴더에 생성되므로,
    법정동 코드 파일이 실행기와 같은 폴더에 있을 경우 기존 방식으로는 찾지 못할 수 있습니다.
    """
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / filename)

    module_dir = app_base_path()
    candidates.extend([
        module_dir / filename,
        module_dir.parent / filename,       # 통합 실행기 루트
        Path.cwd() / filename,              # 현재 작업폴더
        Path.home() / filename,             # 사용자 홈 폴더
    ])

    # 사용자가 전역 설정에 저장한 법정동 파일 경로도 후보로 사용
    try:
        cfg = load_app_config()
        saved = cfg.get("region_file", "")
        if saved:
            candidates.insert(0, Path(saved))
    except Exception:
        pass

    for p in candidates:
        try:
            if p and Path(p).exists():
                return Path(p)
        except Exception:
            pass

    # 못 찾더라도 기존 UI에서 직접 선택할 수 있도록 가장 가능성 높은 통합 실행기 루트 경로 반환
    return module_dir.parent / filename

DEFAULT_REGION_FILENAME = "전국 법정동 지역.xls"
DEFAULT_REGION_FILE = bundled_resource_path(DEFAULT_REGION_FILENAME)
REGION_DB_FILENAME = "region_codes.sqlite3"

# 조사기준월 콤보박스 선택 범위(값은 YYYY-MM, 화면 표시는 YYYY년 M월)
SURVEY_MONTHS = [f"{year}-{month:02d}" for year in (2026, 2027) for month in range(1, 13)]
DEFAULT_SURVEY_MONTH = "2026-08"

# 조사지역 기본값(저장된 선택이 없을 때 사용)
DEFAULT_SIDO = "전남광주통합특별시"
DEFAULT_SIGUNGU = "광산구"
DEFAULT_REGION_CODE = "12330"


def survey_month_label(value: str) -> str:
    year, month = value.split("-")
    return f"{year}년 {int(month)}월"


def launch_dir() -> Path:
    """사용자가 실행한 프로그램(EXE 또는 실행 스크립트)이 있는 폴더."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    try:
        return Path(sys.argv[0]).resolve().parent
    except Exception:
        return Path.cwd()


def default_output_dir() -> Path:
    """실행 폴더 아래 output 폴더를 자동 생성해 출력 폴더로 고정한다."""
    out = launch_dir() / "output"
    try:
        out.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return out


def region_db_path() -> Path | None:
    """법정동 코드 로컬 DB(local_db/region_codes.sqlite3) 위치 탐색."""
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "local_db" / REGION_DB_FILENAME)
    for root in [launch_dir(), app_base_path(), app_base_path().parent, Path.cwd()]:
        candidates.append(root / "local_db" / REGION_DB_FILENAME)
    for c in candidates:
        try:
            if c.exists() and c.is_file():
                return c
        except Exception:
            continue
    return None


def load_region_codes_from_db() -> pd.DataFrame | None:
    path = region_db_path()
    if path is None:
        return None
    import sqlite3
    with sqlite3.connect(str(path)) as conn:
        return pd.read_sql_query("SELECT REG, 시도, 시군구 FROM region_codes", conn)
MOLIT_OPENAPI_URL = "http://stat.molit.go.kr/portal/openapi/service/rest/getList.do"
MOLIT_DEFAULT_KEY = "__MOLIT_API_KEY_REDACTED__"
RTMS_SERVICE_KEY = "__RTMS_SERVICE_KEY_REDACTED__"
RTMS_BASE_URL = "https://apis.data.go.kr/1613000"
MOLIT_RTMS_RECENT_MONTHS = 36  # 실거래 최근 3년치(36개월) 수집

SIMPLE_RUN_FOLDER = "선택통계실행"


# KOSIS/부동산원 주택가격동향조사 지역코드 보정표
# - 전국주택가격동향조사/주간아파트 표는 법정동 REG 5자리와 다른 c### 지역코드를 사용함
# - 화면 콤보박스의 시군구명 기준으로 아래 코드로 변환하여 전달
KAB_SIGUNGU_CODE_MAP = {
    # 전라남도 시·군: 예) 목포시 -> c501
    "목포시": "c501", "여수시": "c502", "순천시": "c503", "나주시": "c504", "광양시": "c505",
    "담양군": "c506", "곡성군": "c507", "구례군": "c508", "고흥군": "c509", "보성군": "c510",
    "화순군": "c511", "장흥군": "c512", "강진군": "c513", "해남군": "c514", "영암군": "c515",
    "무안군": "c516", "함평군": "c517", "영광군": "c518", "장성군": "c519", "완도군": "c520",
    "진도군": "c521", "신안군": "c522",
}

WEEKLY_APT_REGION_CODES = " ".join([
    "a7010101", "a7010102", "a7010103", "a7010201", "a7010202", "a7010203",
    "a7010204", "a7010205", "a7010206", "a7010207", "a7010208", "a7010301",
    "a7010302", "a7010303", "a7020101", "a7020102", "a7020103", "a7020104",
    "a7020105", "a7020106", "a7020107", "a7020201", "a7020202", "a7020203",
    "a7020204", "a8010201", "a8010202", "a8010301", "a8010302", "a8010303",
    "a8020201", "a8020202", "a8020203", "a8020301", "a8020302", "a8020303",
    "a8020304", "a8030101", "a8030102", "a8030103", "a8030201", "a8030202",
    "a8030501", "a8030502", "a8030503", "a8030504", "a8060201", "a8060202",
    "a8060203",
])


# 전국주택가격동향조사 유형별 가격지수(DT_30404_C001/C012/C013, 기준시점 2026.6=100)는
# objL1=ALL, objL2=ALL로 전체 지역을 받은 뒤 KOSIS 내부 권역코드 기준으로 후처리합니다.
# 화면에서 선택한 법정동 시도/시군구를 기준으로 "해당 시도 + 해당 시군구"만 남깁니다.
KAB_PRICE_INDEX_SIDO_PREFIX_MAP = {
    "11": "a7",  # 서울특별시
    "26": "b1",  # 부산광역시
    "27": "b2",  # 대구광역시
    "28": "a9",  # 인천광역시
    "29": "b3",  # 광주광역시
    "30": "b4",  # 대전광역시
    "31": "b5",  # 울산광역시
    "36": "b6",  # 세종특별자치시
    "41": "a8",  # 경기도
    "42": "c1",  # 강원도
    "51": "c1",  # 강원특별자치도
    "43": "c2",  # 충청북도
    "44": "c3",  # 충청남도
    "45": "c4",  # 전라북도
    "52": "c4",  # 전북특별자치도
    "46": "c5",  # 전라남도
    "47": "c6",  # 경상북도
    "48": "c7",  # 경상남도
    "50": "c8",  # 제주특별자치도
}


# KOSIS 통계 반영이 행정구역 개편보다 늦는 경우가 있어 KOSIS 조회에만 구코드를 사용합니다.
# 국토부 실거래 LAWD_CD는 현재 법정동 코드가 필요하므로 이 보정표를 적용하지 않습니다.
KOSIS_LEGACY_REGION_CODE_MAP = {
    ("전남광주통합특별시", "목포시"): "46110",
    ("전남광주통합특별시", "여수시"): "46130",
    ("전남광주통합특별시", "순천시"): "46150",
    ("전남광주통합특별시", "광양시"): "46230",
    ("전남광주통합특별시", "동구"): "29110",
    ("전남광주통합특별시", "서구"): "29140",
    ("전남광주통합특별시", "남구"): "29155",
    ("전남광주통합특별시", "북구"): "29170",
    ("전남광주통합특별시", "광산구"): "29200",  # 2026-07 기준 KOSIS 주민등록/이동자수는 광주광역시 광산구 체계
    ("전남광주통합특별시", "나주시"): "46170",  # 2026-07 기준 KOSIS 주민등록/이동자수는 전라남도 나주시 체계
    ("전남광주통합특별시", "담양군"): "46710",
    ("전남광주통합특별시", "곡성군"): "46720",
    ("전남광주통합특별시", "구례군"): "46730",
    ("전남광주통합특별시", "고흥군"): "46770",
    ("전남광주통합특별시", "보성군"): "46780",
    ("전남광주통합특별시", "화순군"): "46790",
    ("전남광주통합특별시", "장흥군"): "46800",
    ("전남광주통합특별시", "강진군"): "46810",
    ("전남광주통합특별시", "해남군"): "46820",
    ("전남광주통합특별시", "영암군"): "46830",
    ("전남광주통합특별시", "무안군"): "46840",
    ("전남광주통합특별시", "함평군"): "46860",
    ("전남광주통합특별시", "영광군"): "46870",
    ("전남광주통합특별시", "장성군"): "46880",
    ("전남광주통합특별시", "완도군"): "46890",
    ("전남광주통합특별시", "진도군"): "46900",
    ("전남광주통합특별시", "신안군"): "46910",
    "12110": "46110", "12130": "46130", "12150": "46150", "12170": "46170", "12190": "46230",
    "12210": "29110",
    "12240": "29140",
    "12270": "29155",
    "12300": "29170",
    "12330": "29200",
    "12170": "46170",
    "12710": "46710", "12720": "46720", "12730": "46730", "12740": "46770",
    "12750": "46780", "12760": "46790", "12770": "46800", "12780": "46810",
    "12790": "46820", "12800": "46830", "12810": "46840", "12820": "46860",
    "12830": "46870", "12840": "46880", "12850": "46890", "12860": "46900", "12870": "46910",
}
KOSIS_LEGACY_SIDO_CODE_MAP = {
    ("전남광주통합특별시", "동구"): "29",
    ("전남광주통합특별시", "서구"): "29",
    ("전남광주통합특별시", "남구"): "29",
    ("전남광주통합특별시", "북구"): "29",
    ("전남광주통합특별시", "광산구"): "29",
    ("전남광주통합특별시", "나주시"): "46",
    **{("전남광주통합특별시", name): "46" for name in [
        "목포시", "여수시", "순천시", "광양시", "담양군", "곡성군", "구례군", "고흥군",
        "보성군", "화순군", "장흥군", "강진군", "해남군", "영암군", "무안군", "함평군",
        "영광군", "장성군", "완도군", "진도군", "신안군",
    ]},
    "12": "46",
}

GWANGJU_AREA_SIGUNGU_NAMES = {"동구", "서구", "남구", "북구", "광산구"}


MIGRATION_SUBLIST_IDS = [
    ("A34_1", "연령(각세별) 이동자수"),
    ("A34_2", "이동건수(전입신고건수)"),
    ("A34_3", "성/연령(5세)별 이동자수"),
    ("A_1_004", "전입사유별이동"),
]


DATASETS = {
    "unsold_sigungu": {
        "label": "미분양주택현황_최근15개월_전월대비산출용",
        "status": "verified",
        "source_type": "molit_unsold",
        "orgId": "116",
        "tblId": "DT_MLTM_2082",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "filename": "미분양주택현황_시군구_최근15개월_전월대비산출용_{suffix}.xlsx",
        "params": {},
        "region_param_key": None,
        "region_code_length": 5,
        "require_region": True,
        "note": "국토교통부 OpenAPI form_id=2082, style_num=128 기준. 표시 13개월의 전월 대비 증감 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
        "molit_form_id": "2082",
        "molit_style_num": "128",
        "post_filter_region": False,
    },
    "post_completion_unsold": {
        "label": "준공후미분양현황_최근15개월_전월대비산출용",
        "status": "verified",
        "source_type": "molit_unsold",
        "orgId": "116",
        "tblId": "DT_MLTM_5328",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "filename": "준공후미분양현황_시군구_최근15개월_전월대비산출용_{suffix}.xlsx",
        "params": {},
        "region_param_key": None,
        "region_code_length": 5,
        "require_region": True,
        "note": "국토교통부 OpenAPI form_id=5328, style_num=1 기준. 표시 13개월의 전월 대비 증감 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
        "molit_form_id": "5328",
        "molit_style_num": "1",
        "post_filter_region": False,
        "molit_value_col": "호",
        "molit_item_id": "POST_COMPLETION_UNSOLD",
        "molit_extra_filters": {"규모": "계", "부문": "계"},
    },
    "housing_permit": {
        "label": "주택유형별_주택건설_인허가실적_월별누계",
        "status": "verified",
        "orgId": "116",
        "tblId": "DT_MLTM_1948",
        "prdSe": "M",
        "lookback_years": 6,
        "filename": "주택유형별_주택건설_인허가실적_최근6년_연말및당해월누계_{suffix}.xlsx",
        "params": {
            "itmId": "13103871090T1",
            "objL1": "ALL", "objL2": "ALL", "objL3": "ALL", "objL4": "ALL",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "cumulative_point_request": True,
        "cumulative_years": 6,
        "latest_month_backtrack": 24,
        "note": "전국 고정. 월별 누계 통계이므로 해당년도 포함 최근 6년 기준으로 과거 연도는 매년 12월 누계, 당해연도는 선택 종료월 누계를 단건 조회",
    },
    "housing_completion": {
        "label": "주택유형별_주택건설_준공실적_월별누계",
        "status": "verified",
        "orgId": "116",
        "tblId": "DT_MLTM_5373",
        "prdSe": "M",
        "lookback_years": 6,
        "filename": "주택유형별_주택건설_준공실적_최근6년_연말및당해월누계_{suffix}.xlsx",
        "params": {
            "itmId": "13103766973T1",
            "objL1": "ALL", "objL2": "ALL", "objL3": "ALL", "objL4": "ALL",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "cumulative_point_request": True,
        "cumulative_years": 6,
        "latest_month_backtrack": 24,
        "note": "전국 고정. 월별 누계 통계이므로 해당년도 포함 최근 6년 기준으로 과거 연도는 매년 12월 누계, 당해연도는 선택 종료월 누계를 단건 조회",
    },
    "housing_start": {
        "label": "주택유형별_주택건설_착공실적_월별누계",
        "status": "verified",
        "orgId": "116",
        "tblId": "DT_MLTM_5387",
        "prdSe": "M",
        "lookback_years": 6,
        "filename": "주택유형별_주택건설_착공실적_최근6년_연말및당해월누계_{suffix}.xlsx",
        "params": {
            "itmId": "13103766969T1",
            "objL1": "ALL", "objL2": "ALL", "objL3": "ALL", "objL4": "ALL",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "cumulative_point_request": True,
        "cumulative_years": 6,
        "latest_month_backtrack": 24,
        "note": "전국 고정. 월별 누계 통계이므로 해당년도 포함 최근 6년 기준으로 과거 연도는 매년 12월 누계, 당해연도는 선택 종료월 누계를 단건 조회",
    },
    "housing_sale_supply": {
        "label": "주택건설_분양실적_공동주택_월별누계",
        "status": "verified",
        "orgId": "116",
        "tblId": "DT_MLTM_5557",
        "prdSe": "M",
        "lookback_years": 6,
        "filename": "주택건설_분양실적_공동주택_최근6년_연말및당해월누계_{suffix}.xlsx",
        "params": {
            "itmId": "13103133605T1 13103133605T2 13103133605T3 13103133605T4",
            "objL1": "ALL", "objL2": "ALL", "objL3": "ALL", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "cumulative_point_request": True,
        "cumulative_years": 6,
        "latest_month_backtrack": 24,
        "note": "전국 고정. 월별 누계 통계이므로 해당년도 포함 최근 6년 기준으로 과거 연도는 매년 12월 누계, 당해연도는 선택 종료월 누계를 단건 조회",
    },
    "rtms_apt_trade": {
        "label": "국토교통부_아파트_매매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade",
        "recent_months": 36,
        "filename": "국토교통부_아파트_매매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_apt_rent": {
        "label": "국토교통부_아파트_전월세_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcAptRent/getRTMSDataSvcAptRent",
        "recent_months": 36,
        "filename": "국토교통부_아파트_전월세_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_rh_trade": {
        "label": "국토교통부_연립다세대_매매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",
        "recent_months": 36,
        "filename": "국토교통부_연립다세대_매매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_rh_rent": {
        "label": "국토교통부_연립다세대_전월세_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcRHRent/getRTMSDataSvcRHRent",
        "recent_months": 36,
        "filename": "국토교통부_연립다세대_전월세_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_sh_trade": {
        "label": "국토교통부_단독다가구_매매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade",
        "recent_months": 36,
        "filename": "국토교통부_단독다가구_매매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_sh_rent": {
        "label": "국토교통부_단독다가구_전월세_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcSHRent/getRTMSDataSvcSHRent",
        "recent_months": 36,
        "filename": "국토교통부_단독다가구_전월세_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_silv_trade": {
        "label": "국토교통부_아파트_분양권전매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcSilvTrade/getRTMSDataSvcSilvTrade",
        "recent_months": 36,
        "filename": "국토교통부_아파트_분양권전매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_offi_trade": {
        "label": "국토교통부_오피스텔_매매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcOffiTrade/getRTMSDataSvcOffiTrade",
        "recent_months": 36,
        "filename": "국토교통부_오피스텔_매매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_offi_rent": {
        "label": "국토교통부_오피스텔_전월세_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcOffiRent/getRTMSDataSvcOffiRent",
        "recent_months": 36,
        "filename": "국토교통부_오피스텔_전월세_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_land_trade": {
        "label": "국토교통부_토지_매매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcLandTrade/getRTMSDataSvcLandTrade",
        "recent_months": 36,
        "filename": "국토교통부_토지_매매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_nrg_trade": {
        "label": "국토교통부_상업업무용_매매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcNrgTrade/getRTMSDataSvcNrgTrade",
        "recent_months": 36,
        "filename": "국토교통부_상업업무용_매매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "rtms_indu_trade": {
        "label": "국토교통부_공장창고등_매매_실거래가_최근3년",
        "status": "verified",
        "source_type": "molit_rtms",
        "endpoint": "RTMSDataSvcInduTrade/getRTMSDataSvcInduTrade",
        "recent_months": 36,
        "filename": "국토교통부_공장창고등_매매_실거래가_최근3년_{suffix}.xlsx",
        "require_region": True,
        "note": "국토교통부 실거래가 OpenAPI. LAWD_CD 시군구 5자리, DEAL_YMD 월별 최근 3년(36개월)",
    },
    "population_sigungu": {
        "label": "시군구별_주민등록인구_최근15개월_전월대비산출용",
        "status": "verified",
        "orgId": "101",
        "tblId": "DT_1B04006",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "filename": "시군구별_주민등록인구_최근15개월_전월대비산출용_{suffix}.xlsx",
        "params": {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        "region_param_key": "objL1",
        "region_code_length": 5,
        "require_region": True,
        "note": "행정구역(시군구)별/1세별 주민등록인구. 표시 13개월의 전월 대비 증감·증감률 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "household_sigungu": {
        "label": "시군구별_주민등록세대수_최근15개월_전월대비산출용",
        "status": "verified",
        "orgId": "101",
        "tblId": "DT_1B040B3",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "filename": "시군구별_주민등록세대수_최근15개월_전월대비산출용_{suffix}.xlsx",
        "params": {"objL1": "ALL"},
        "region_param_key": "objL1",
        "region_code_length": 5,
        "require_region": True,
        "custom_param_probe": "household_sigungu",
        "note": "행정구역(시군구)별 주민등록세대수. 표시 13개월의 전월 대비 증감 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "house_price_type_sale": {
        "label": "전국주택가격동향조사_유형별매매가격지수_전국_최근15개월_변동률산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_30404_C012",  # 2026.6=100 개편표(B012는 2026-06 공표 후 중단)
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "single_period_request": True,
        "filename": "전국주택가격동향조사_유형별매매가격지수_전국_최근15개월_변동률산출용_{suffix}.xlsx",
        "params": {
            "itmId": "ALL",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "post_filter_region": True,
        "kab_price_region_filter": True,
        "note": "전국주택가격동향조사 유형별 매매가격지수. 표시 13개월의 전월 대비 변동률 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "house_price_type_jeonse": {
        "label": "전국주택가격동향조사_유형별전세가격지수_전국_최근15개월_변동률산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_30404_C013",  # 2026.6=100 개편표(B013은 2026-06 공표 후 중단)
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "single_period_request": True,
        "filename": "전국주택가격동향조사_유형별전세가격지수_전국_최근15개월_변동률산출용_{suffix}.xlsx",
        "params": {
            "itmId": "ALL",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "post_filter_region": True,
        "kab_price_region_filter": True,
        "note": "전국주택가격동향조사 유형별 전세가격지수. 표시 13개월의 전월 대비 변동률 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "house_price_type_wolse": {
        "label": "전국주택가격동향조사_유형별월세가격지수",
        "status": "hold",
        "orgId": "408",
        "tblId": "DT_30404_C002",
        "prdSe": "M",
        "lookback_years": 3,
        "filename": "전국주택가격동향조사_유형별월세가격지수_{suffix}.xlsx",
        "params": {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        "region_param_key": "objL2",
        "region_code_length": 5,
        "note": "KOSIS 표 전용 차원축 매핑 필요. 공통 요청식으로는 오류가 발생해 보류 처리",
        "note": "전국주택가격동향조사 유형별 월세가격지수",
    },
    "house_price_type_wolse_total": {
        "label": "전국주택가격동향조사_유형별월세통합가격지수_전국_최근15개월_변동률산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_30404_C001",  # 2026.6=100 개편표(B001은 2026-06 공표 후 중단)
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "single_period_request": True,
        "filename": "전국주택가격동향조사_유형별월세통합가격지수_전국_최근15개월_변동률산출용_{suffix}.xlsx",
        "params": {
            "itmId": "rent ",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "post_filter_region": True,
        "kab_price_region_filter": True,
        "note": "전국주택가격동향조사 유형별 월세통합가격지수. 표시 13개월의 전월 대비 변동률 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "rent_conversion_rate_public": {
        "label": "전국주택가격동향조사_전월세전환율_시도_물건유형별_최근15개월_전월대비산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_30404_N0010",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "single_period_request": True,
        "filename": "전국주택가격동향조사_전월세전환율_시도_물건유형별_최근15개월_{suffix}.xlsx",
        "params": {
            "itmId": "T1 ",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "note": "전국주택가격동향조사 전월세전환율. 주택 월간 보고서의 시도 단위·물건유형별 최근 13개월 공표 전환율을 월세 환산에 적용하며, 기준월 자료가 없으면 기준월 이전 가장 가까운 공표월 전환율을 사용",
    },
    "officetel_rent_conversion_rate_public": {
        "label": "오피스텔_전월세전환율_시도_최근15개월_전월대비산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_08002_09_009A",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "single_period_request": True,
        "filename": "오피스텔_전월세전환율_시도_최근15개월_{suffix}.xlsx",
        "params": {
            "itmId": "T001 ",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "note": "오피스텔 전월세전환율. KOSIS orgId=408, tblId=DT_08002_09_009A, itmId=T001, objL1=ALL, objL2=ALL 기준 공표 전환율을 오피스텔 월세 환산에 적용",
    },
    "weekly_apt_sale_change": {
        "label": "주간아파트_매매가격변동률",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_304004_WEEK_001_B",
        "prdSe": "D",
        "lookback_years": 1,
        "recent_months": 13,
        "single_period_request": True,
        "filename": "주간아파트_매매가격변동률_최근13개월_{suffix}.xlsx",
        "params": {
            "itmId": "T001",
            "objL1": WEEKLY_APT_REGION_CODES, "objL2": "", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "require_region": True,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "weekly_apt_region_post_filter": True,
        "weekly_mixed_monthly_current": True,
        "note": "매주 월요일 공표 기준일 최근 13개월. KOSIS DT_304004_WEEK_001_B의 전국 지역별 시군구 objL1 목록을 조회한 뒤 선택 시군구만 후처리 추출",
    },
    "weekly_apt_jeonse_change": {
        "label": "주간아파트_전세가격변동률",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_304004_WEEK_003_B",
        "prdSe": "D",
        "lookback_years": 1,
        "recent_months": 13,
        "single_period_request": True,
        "filename": "주간아파트_전세가격변동률_최근13개월_{suffix}.xlsx",
        "params": {
            "itmId": "T001",
            "objL1": WEEKLY_APT_REGION_CODES, "objL2": "", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "require_region": True,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "weekly_apt_region_post_filter": True,
        "weekly_mixed_monthly_current": True,
        "note": "매주 월요일 공표 기준일 최근 13개월. KOSIS DT_304004_WEEK_003_B의 전국 지역별 시군구 objL1 목록을 조회한 뒤 선택 시군구만 후처리 추출",
    },
    "apt_real_sale_index": {
        "label": "아파트_매매_실거래가격지수",
        "status": "hold",
        "orgId": "408",
        "tblId": "DT_KAB_11672_S1",
        "prdSe": "M",
        "lookback_years": 3,
        "filename": "아파트_매매_실거래가격지수_{suffix}.xlsx",
        "params": {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        "region_param_key": "objL2",
        "region_code_length": 2,
        "note": "KOSIS 표 전용 차원축 매핑 필요. 공통 요청식으로는 오류가 발생해 [매핑필요]로 보류 처리",
    },
    "apt_real_jeonse_index": {
        "label": "아파트_전세_실거래가격지수",
        "status": "hold",
        "orgId": "408",
        "tblId": "DT_KAB_11672_S23",
        "prdSe": "M",
        "lookback_years": 3,
        "filename": "아파트_전세_실거래가격지수_{suffix}.xlsx",
        "params": {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        "region_param_key": "objL2",
        "region_code_length": 2,
        "note": "KOSIS 표 전용 차원축 매핑 필요. 공통 요청식으로는 오류가 발생해 [매핑필요]로 보류 처리",
    },
    "officetel_sale_index": {
        "label": "오피스텔_매매가격지수_최근15개월_변동률산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_08002_01_001A",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "filename": "오피스텔_매매가격지수_전국_최근15개월_변동률산출용_{suffix}.xlsx",
        "params": {
            "itmId": "T001",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "note": "오피스텔 매매가격지수. 표시 13개월의 전월 대비 변동률 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "officetel_jeonse_index": {
        "label": "오피스텔_전세가격지수_최근15개월_변동률산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_08002_02_002A",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "filename": "오피스텔_전세가격지수_전국_최근15개월_변동률산출용_{suffix}.xlsx",
        "params": {
            "itmId": "T001",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "note": "오피스텔 전세가격지수. 표시 13개월의 전월 대비 변동률 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "officetel_wolse_index": {
        "label": "오피스텔_월세가격지수_최근15개월_변동률산출용",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_08002_03_003A",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "filename": "오피스텔_월세가격지수_전국_최근15개월_변동률산출용_{suffix}.xlsx",
        "params": {
            "itmId": "T001",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "note": "오피스텔 월세가격지수. 표시 13개월의 전월 대비 변동률 계산을 위해 원자료는 14개월 수집 후 보고서에는 최근 13개월만 표시",
    },
    "land_price_change_sigungu": {
        "label": "지가변동률_시군구용도이용상황별_최근13개월",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_PLCAHTUSE",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 13,
        "filename": "지가변동률_시군구용도이용상황별_최근13개월_{suffix}.xlsx",
        "params": {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        "region_param_key": "objL1",
        "region_code_length": 5,
        "require_region": True,
        "note": "선택 시군구 5자리 코드를 적용하여 기준월 포함 최근 13개월의 시군구별·용도지역별·이용상황별 지가변동률을 수집",
    },
    "land_price_change_eupmyeondong": {
        "label": "지가변동률_읍면동별",
        "status": "verified",
        "orgId": "408",
        "tblId": "DT_31501N_010",
        "prdSe": "M",
        "lookback_years": 3,
        "filename": "지가변동률_읍면동별_{suffix}.xlsx",
        "params": {"objL1": "ALL", "itmId": "ALL"},
        "region_param_key": "objL1",
        "region_code_length": 5,
        "note": "읍면동별 지가변동률",
    },
    "marriage_sigungu": {
        "label": "시군구별_월별혼인건수",
        "status": "hold",
        "orgId": "101",
        "tblId": "DT_1B83A35",
        "prdSe": "M",
        "lookback_years": 3,
        "filename": "시군구별_월별혼인건수_{suffix}.xlsx",
        "params": {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        "region_param_key": "objL2",
        "region_code_length": 5,
        "note": "KOSIS 파라미터 재확인 필요. 현재 공통 요청식으로는 오류가 발생해 보류 처리",
    },
    "migration_sigungu": {
        "label": "시군구_연령5세별_이동자수_최근15개월_전월대비산출용",
        "status": "verified",
        "orgId": "101",
        "tblId": "DT_1B26001_A03",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "single_period_request": True,
        "filename": "시군구_연령5세별_이동자수_최근15개월_전월대비산출용_{suffix}.xlsx",
        "params": {
            "itmId": "T10 T20 T25 T30 T31 T32 T40 T50",
            "objL1": "ALL", "objL2": "ALL", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": "objL1",
        "region_code_length": 5,
        "require_region": True,
        "note": "시군구/연령(5세)별 이동자수. KOSIS DT_1B26001_A03 / objL1=선택 시군구 5자리 / objL2=ALL / 표시 13개월의 직전월 흐름 참고를 위해 원자료는 14개월 수집",
    },
    "one_person_household_residence": {
        "label": "시군구별_1인가구_거처종류별",
        "status": "hold",
        "orgId": "101",
        "tblId": "DT_1PL1501",
        "prdSe": "Y",
        "lookback_years": 5,
        "filename": "시군구별_1인가구_거처종류별_{suffix}.xlsx",
        "params": {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        "region_param_key": "objL2",
        "region_code_length": 5,
        "note": "KOSIS 파라미터 재확인 필요. 현재 공통 요청식으로는 오류가 발생해 보류 처리",
    },
    "deposit_bank_loan_rate_new": {
        "label": "예금은행_대출금리_신규취급액기준_전국_최근15개월_전월대비산출용",
        "status": "verified",
        "orgId": "301",
        "tblId": "DT_121Y006",
        "prdSe": "M",
        "lookback_years": 2,
        "recent_months": 15,
        "single_period_request": True,
        "filename": "예금은행_대출금리_신규취급액기준_전국_최근15개월_전월대비산출용_{suffix}.xlsx",
        "params": {
            "itmId": "13103134553999",
            "objL1": "ALL", "objL2": "", "objL3": "", "objL4": "",
            "objL5": "", "objL6": "", "objL7": "", "objL8": "",
        },
        "region_param_key": None,
        "region_code_length": 0,
        "no_region_filter": True,
        "sido_total_single_fetch": True,
        "note": "예금은행 대출금리(신규취급액 기준). KOSIS DT_121Y006 / orgId=301 / itmId=13103134553999 / 표시 13개월의 전월 대비 증감 계산을 위해 원자료는 14개월 수집",
    },
}

HOLD_DATASET_KEYS = {"marriage_sigungu", "one_person_household_residence"}
HIDDEN_GROUP_NAMES = set()

# 화면에 표시할 통계 키: 업무그룹에 포함된 개발 대상 통계만 노출

# ---------------------------------------------------------------------------
# 한국부동산원 R-ONE OpenAPI 수집
# KOSIS에 늦게 옮겨 실리거나(주간아파트는 2026.6 기준 개편 이후 KOSIS 갱신 중단) 같은 값을 그대로 싣는
# 한국부동산원 작성 통계는 R-ONE에서 직접 받는다. 결과는 KOSIS 응답과 같은 열 이름(C1_NM·C2_NM·PRD_DE·DT 등)으로
# 바꿔 기존 저장·보고서 코드를 그대로 쓰며, R-ONE 키가 없거나 조회가 실패하면 KOSIS로 받는다.
# 읍면동 지가변동률(DT_31501N_010)과 실거래가격지수는 R-ONE OpenAPI에 없어 KOSIS를 유지한다.
# ---------------------------------------------------------------------------
RONE_DATA_URL = "https://www.reb.or.kr/r-one/openapi/SttsApiTblData.do"
RONE_PAGE_SIZE = 1000

RONE_SPECS = {
    "house_price_type_sale": {"kind": "type_region", "cycle": "MM", "itm_id": "sales", "itm_nm": "매매가격", "tbl_nm": "유형별 매매가격지수", "unit": "2026.6=100.0",
                              "tables": [("A_2024_00016", "종합", "0"), ("A_2024_00045", "아파트", "1")]},
    "house_price_type_jeonse": {"kind": "type_region", "cycle": "MM", "itm_id": "sales", "itm_nm": "전세가격", "tbl_nm": "유형별 전세가격지수", "unit": "2026.6=100.0",
                                "tables": [("A_2024_00019", "종합", "0"), ("A_2024_00050", "아파트", "1")]},
    "house_price_type_wolse_total": {"kind": "type_region", "cycle": "MM", "itm_id": "rent", "itm_nm": "월세통합가격", "tbl_nm": "유형별 월세통합가격지수", "unit": "2026.6=100.0",
                                     "tables": [("A_2024_00021", "종합", "0"), ("A_2024_00054", "아파트", "1")]},
    "rent_conversion_rate_public": {"kind": "conversion", "cycle": "MM", "itm_id": "T1", "itm_nm": "지역별 전월세전환율", "tbl_nm": "지역별 전월세전환율", "unit": "%",
                                    "tables": [("A_2024_00155", "종합", "0"), ("A_2024_00156", "아파트", "1"), ("A_2024_00157", "연립다세대", "2"), ("A_2024_00158", "단독주택", "3")]},
    "officetel_rent_conversion_rate_public": {"kind": "officetel", "cycle": "MM", "itm_id": "T001", "itm_nm": "전월세전환율", "tbl_nm": "오피스텔 전월세전환율", "unit": "%",
                                              "tables": [("T241163133546529", "", "")]},
    "officetel_sale_index": {"kind": "officetel", "cycle": "MM", "itm_id": "T001", "itm_nm": "매매가격지수", "tbl_nm": "오피스텔 매매가격지수", "unit": "지수",
                             "tables": [("A_2024_00615", "", "")]},
    "officetel_jeonse_index": {"kind": "officetel", "cycle": "MM", "itm_id": "T001", "itm_nm": "전세가격지수", "tbl_nm": "오피스텔 전세가격지수", "unit": "지수",
                               "tables": [("A_2024_00618", "", "")]},
    "officetel_wolse_index": {"kind": "officetel", "cycle": "MM", "itm_id": "T001", "itm_nm": "월세가격지수", "tbl_nm": "오피스텔 월세가격지수", "unit": "지수",
                              "tables": [("A_2024_00621", "", "")]},
    "weekly_apt_sale_change": {"kind": "weekly", "cycle": "WK", "itm_id": "T001", "itm_nm": "매매가격변동률", "tbl_nm": "주간 아파트 매매가격변동률", "unit": "%",
                               "tables": [("T244183132827305", "", "")]},
    "weekly_apt_jeonse_change": {"kind": "weekly", "cycle": "WK", "itm_id": "T001", "itm_nm": "전세가격변동률", "tbl_nm": "주간 아파트 전세가격변동률", "unit": "%",
                                 "tables": [("T247713133046872", "", "")]},
    "land_price_change_sigungu": {"kind": "land", "cycle": "MM", "itm_id": "13103890822T1", "itm_nm": "지가변동률", "tbl_nm": "시군구별/용도지역별/이용상황별 지가변동률", "unit": "%",
                                  "tables": [("A_2024_00007", "용도지역", ""), ("A_2024_00008", "이용상황", ""), ("A_2024_00903", "평균", "")]},
}

# R-ONE 시도 표기(경로 첫 단계)
RONE_SIDO_NAME_MAP = {
    "서울특별시": "서울", "부산광역시": "부산", "대구광역시": "대구", "인천광역시": "인천", "광주광역시": "광주",
    "대전광역시": "대전", "울산광역시": "울산", "세종특별자치시": "세종", "경기도": "경기", "강원도": "강원",
    "강원특별자치도": "강원", "충청북도": "충북", "충청남도": "충남", "전라북도": "전북", "전북특별자치도": "전북",
    "전라남도": "전남", "경상북도": "경북", "경상남도": "경남", "제주특별자치도": "제주", "전남광주통합특별시": "전남광주",
}


def rone_request(api_key: str, params: dict, timeout: int = 60) -> list:
    """R-ONE 통계 자료를 페이지 단위로 모두 받는다. 자료가 없으면 빈 목록."""
    rows, page = [], 1
    while True:
        query = {"KEY": api_key, "Type": "json", "pIndex": page, "pSize": RONE_PAGE_SIZE, **params}
        url = RONE_DATA_URL + "?" + urllib.parse.urlencode(query)
        last_error = None
        for attempt in range(3):
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                break
            except Exception as exc:
                last_error = exc
                time.sleep(1.5 * (attempt + 1))
        else:
            raise RuntimeError(f"R-ONE 연결 실패: {last_error}")
        body = data.get("SttsApiTblData") if isinstance(data, dict) else None
        if not body:
            result = (data.get("RESULT") or {}) if isinstance(data, dict) else {}
            code = str(result.get("CODE") or "")
            if code and code != "INFO-200":  # INFO-200: 해당 데이터 없음
                raise RuntimeError(f"R-ONE 오류 {code}: {result.get('MESSAGE')}")
            return rows
        total = int(body[0]["head"][0].get("list_total_count") or 0)
        rows.extend(body[1].get("row") or [])
        if len(rows) >= total or not body[1].get("row"):
            return rows
        page += 1


def _rone_norm(name: str) -> str:
    text = str(name or "").strip().replace(" ", "")
    for suffix in ("특별자치시", "특별자치도", "광역시", "특별시", "시", "군", "구"):
        if text.endswith(suffix) and len(text) > len(suffix):
            return text[: -len(suffix)]
    return text


def _rone_sido_core(sido_name: str) -> str:
    raw = str(sido_name or "").strip()
    return RONE_SIDO_NAME_MAP.get(raw, re.sub(r"(통합특별시|특별자치시|특별자치도|광역시|특별시|도)$", "", raw))


def _rone_month_shift(yyyymm: str, months: int) -> str:
    y, m = int(yyyymm[:4]), int(yyyymm[4:6])
    total = y * 12 + (m - 1) + months
    return f"{total // 12:04d}{total % 12 + 1:02d}"




for _rone_key_name, _rone_spec in RONE_SPECS.items():
    if _rone_key_name in DATASETS:
        DATASETS[_rone_key_name]["rone"] = _rone_spec


MIGRATION_SIDO_TBL_MAP = {
    "11": "DT_1B26B02",
    "26": "DT_1B26B03",
    "27": "DT_1B26B04",
    "28": "DT_1B26B05",
    "29": "DT_1B26B06",
    "30": "DT_1B26B07",
    "31": "DT_1B26B08",
    "36": "DT_1B26B18",
    "41": "DT_1B26B09",
    "42": "DT_1B26B10",
    "43": "DT_1B26B11",
    "44": "DT_1B26B12",
    "45": "DT_1B26B13",
    "46": "DT_1B26B14",
    "47": "DT_1B26B15",
    "48": "DT_1B26B16",
    "50": "DT_1B26B17",
    # 특별자치도 코드 반영
    "51": "DT_1B26B10",  # 강원특별자치도
    "52": "DT_1B26B13",  # 전북특별자치도
}

MIGRATION_SIDO_NAME_TBL_MAP = {
    "서울특별시": "DT_1B26B02",
    "부산광역시": "DT_1B26B03",
    "대구광역시": "DT_1B26B04",
    "인천광역시": "DT_1B26B05",
    "광주광역시": "DT_1B26B06",
    "대전광역시": "DT_1B26B07",
    "울산광역시": "DT_1B26B08",
    "세종특별자치시": "DT_1B26B18",
    "경기도": "DT_1B26B09",
    "강원특별자치도": "DT_1B26B10",
    "충청북도": "DT_1B26B11",
    "충청남도": "DT_1B26B12",
    "전북특별자치도": "DT_1B26B13",
    "전라남도": "DT_1B26B14",
    "경상북도": "DT_1B26B15",
    "경상남도": "DT_1B26B16",
    "제주특별자치도": "DT_1B26B17",
}


GROUPS = {
    "주택동향업무_월간": [
        "population_sigungu", "household_sigungu", "migration_sigungu",
        "unsold_sigungu", "post_completion_unsold",
        "housing_permit", "housing_start", "housing_completion", "housing_sale_supply",
        "house_price_type_sale", "house_price_type_jeonse", "house_price_type_wolse_total",
        "rent_conversion_rate_public",
        "deposit_bank_loan_rate_new",
        "rtms_apt_trade", "rtms_apt_rent", "rtms_rh_trade", "rtms_rh_rent",
        "rtms_sh_trade", "rtms_sh_rent", "rtms_silv_trade",
    ],
    "주택동향업무_주간": [
        # 주간형은 월간 기초통계·공급통계는 유지하고, 비아파트 실거래 API만 제외한다.
        "population_sigungu", "household_sigungu", "migration_sigungu",
        "unsold_sigungu", "post_completion_unsold",
        "housing_permit", "housing_start", "housing_completion", "housing_sale_supply",
        "house_price_type_sale", "house_price_type_jeonse",
        "deposit_bank_loan_rate_new",
        "rtms_apt_trade", "rtms_apt_rent", "rtms_silv_trade",
    ],
    "오피스텔동향업무": [
        "population_sigungu", "household_sigungu", "migration_sigungu",
        "unsold_sigungu", "post_completion_unsold",
        "housing_permit", "housing_start", "housing_completion", "housing_sale_supply",
        "officetel_sale_index", "officetel_jeonse_index", "officetel_wolse_index",
        "officetel_rent_conversion_rate_public",
        "deposit_bank_loan_rate_new",
        "rtms_offi_trade", "rtms_offi_rent",
    ],
    "지가변동률업무": [
        "population_sigungu", "household_sigungu", "migration_sigungu",
        "land_price_change_sigungu",
        "deposit_bank_loan_rate_new",
    ],
}

GROUP_DESCRIPTIONS = {
    "주택동향업무_월간": {
        "title": "주택동향업무(월간)",
        "items": [
            ("population_sigungu", "지역 인구 규모 및 증감 흐름 파악"),
            ("household_sigungu", "세대 증가 여부 및 가구 분화 확인"),
            ("migration_sigungu", "시군구 5자리 코드 기준 연령 5세별 이동자수 최근 14개월 원자료로 최근 13개월 흐름 파악"),
            ("unsold_sigungu", "공급 적체 및 분양시장 부담 확인"),
            ("post_completion_unsold", "준공 이후 누적 미분양 현황 확인"),
            ("housing_permit", "향후 공급 예정 물량 판단"),
            ("housing_start", "실제 공사 착수 물량 확인"),
            ("housing_completion", "신규 입주 가능 물량 확인"),
            ("housing_sale_supply", "공동주택 분양 누계 흐름 확인"),
            ("house_price_type_sale", "월간 매매가격 흐름 파악"),
            ("house_price_type_jeonse", "월간 전세가격 흐름 파악"),
            ("house_price_type_wolse_total", "월세 종합 흐름 참고"),
            ("rent_conversion_rate_public", "시도 단위·물건유형별 공표 전월세전환율을 월세 환산 기준으로 적용"),
            ("deposit_bank_loan_rate_new", "한국은행 KOSIS 예금은행 대출금리 신규취급액 기준 확인"),
            ("rtms_apt_trade", "아파트 매매 실거래 최근 13개월 확인"),
            ("rtms_apt_rent", "아파트 전월세 실거래 최근 13개월 확인"),
            ("rtms_rh_trade", "연립다세대 매매 실거래 최근 13개월 확인"),
            ("rtms_rh_rent", "연립다세대 전월세 실거래 최근 13개월 확인"),
            ("rtms_sh_trade", "단독다가구 매매 실거래 최근 13개월 확인"),
            ("rtms_sh_rent", "단독다가구 전월세 실거래 최근 13개월 확인"),
            ("rtms_silv_trade", "아파트 분양권전매 실거래 최근 13개월 확인"),
        ],
    },
    "주택동향업무_주간": {
        "title": "주택동향업무(주간)",
        "items": [
            ("population_sigungu", "기초 수요 규모 참고"),
            ("household_sigungu", "세대 흐름 참고"),
            ("migration_sigungu", "시군구 5자리 코드 기준 연령 5세별 이동자수 최근 14개월 원자료로 최근 13개월 흐름 확인"),
            ("unsold_sigungu", "공급 부담 참고"),
            ("post_completion_unsold", "누적 미분양 참고"),
            ("housing_permit", "향후 공급 참고"),
            ("housing_start", "공급 진행상황 참고"),
            ("housing_completion", "입주 가능 물량 참고"),
            ("housing_sale_supply", "공동주택 분양 누계 참고"),
            ("house_price_type_sale", "13개월 월간 아파트 매매가격지수 기반 전월 대비 변동률 산출"),
            ("house_price_type_jeonse", "13개월 월간 아파트 전세가격지수 기반 전월 대비 변동률 산출"),
            ("deposit_bank_loan_rate_new", "한국은행 KOSIS 예금은행 대출금리 신규취급액 기준 확인"),
            ("rtms_apt_trade", "아파트 매매 실거래 최근 13개월 참고"),
            ("rtms_apt_rent", "아파트 전월세 실거래 최근 13개월 참고"),
            ("rtms_silv_trade", "아파트 분양권전매 실거래 최근 13개월 참고"),
            ("weekly_non_apt_excluded", "주간형: 연립다세대·단독다가구 실거래 API 제외"),
        ],
    },
    "오피스텔동향업무": {
        "title": "오피스텔동향업무",
        "items": [
            ("population_sigungu", "기초 수요 규모 참고"),
            ("household_sigungu", "세대 분화 및 소형 주거 수요 참고"),
            ("migration_sigungu", "시군구 5자리 코드 기준 연령 5세별 이동자수로 임대수요 흐름 파악"),
            ("unsold_sigungu", "주거형 수익부동산 분양·입주 부담을 판단하기 위한 시군구 미분양 흐름 확인"),
            ("post_completion_unsold", "준공 이후 남은 미분양 재고로 임대·매매시장 부담 여부 확인"),
            ("housing_permit", "향후 주거 공급 예정 물량과 오피스텔 수요 대체 가능성 참고"),
            ("housing_start", "실제 착공 물량을 통해 단기 공급 진행상황 확인"),
            ("housing_completion", "신규 입주 가능 물량과 임대 경쟁 압력 확인"),
            ("housing_sale_supply", "공동주택 분양 누계 흐름으로 신규 주거 공급 부담 참고"),
            ("officetel_sale_index", "오피스텔 매매시장 흐름 확인"),
            ("officetel_jeonse_index", "오피스텔 전세시장 흐름 확인"),
            ("officetel_wolse_index", "오피스텔 월세시장 흐름 확인"),
            ("officetel_rent_conversion_rate_public", "시도 단위·오피스텔 공표 전월세전환율을 오피스텔 월세 환산 기준으로 적용"),
            ("deposit_bank_loan_rate_new", "한국은행 KOSIS 예금은행 대출금리 신규취급액 기준 확인"),
            ("rtms_offi_trade", "오피스텔 매매 실거래 최근 13개월 확인"),
            ("rtms_offi_rent", "오피스텔 전월세 실거래 최근 13개월 확인"),
        ],
    },
    "지가변동률업무": {
        "title": "지가변동률업무",
        "items": [
            ("population_sigungu", "수동 업로드 실거래·지가동향 해석을 보조하는 지역 수요 기반 참고"),
            ("household_sigungu", "생활권 확대 및 세대 흐름 참고"),
            ("migration_sigungu", "시군구 5자리 코드 기준 연령 5세별 이동자수로 지역 활력 참고"),
            ("land_price_change_sigungu", "선택 시군구의 용도지역·이용상황별 지가변동률 최근 13개월 수집"),
            ("deposit_bank_loan_rate_new", "한국은행 KOSIS 예금은행 대출금리 신규취급액 기준 확인"),
        ],
    },
}

# 화면에 표시할 개발완료/업무그룹 통계 키 목록
# GROUPS 정의 이후에 생성해야 NameError가 발생하지 않습니다.
WEEKLY_HOUSING_GROUP = "\uc8fc\ud0dd\ub3d9\ud5a5\uc5c5\ubb34_\uc8fc\uac04"
MONTHLY_HOUSING_GROUP = "\uc8fc\ud0dd\ub3d9\ud5a5\uc5c5\ubb34_\uc6d4\uac04"
MONTHLY_OFFICETEL_GROUP = "\uc624\ud53c\uc2a4\ud154\ub3d9\ud5a5\uc5c5\ubb34"
LAND_PRICE_GROUP = "\uc9c0\uac00\ubcc0\ub3d9\ub960\uc5c5\ubb34"

GROUP_DISPLAY_ORDER = [
    WEEKLY_HOUSING_GROUP,
    MONTHLY_HOUSING_GROUP,
    MONTHLY_OFFICETEL_GROUP,
    LAND_PRICE_GROUP,
]

if MONTHLY_OFFICETEL_GROUP in GROUP_DESCRIPTIONS:
    GROUP_DESCRIPTIONS[MONTHLY_OFFICETEL_GROUP]["title"] = "\uc624\ud53c\uc2a4\ud154\ub3d9\ud5a5\uc5c5\ubb34(\uc6d4\uac04)"


def iter_visible_group_descriptions():
    """Return visible business groups in the fixed UI order."""
    yielded = set()
    for group_name in GROUP_DISPLAY_ORDER:
        if group_name in GROUP_DESCRIPTIONS and group_name not in HIDDEN_GROUP_NAMES:
            yielded.add(group_name)
            yield group_name, GROUP_DESCRIPTIONS[group_name]
    for group_name, info in GROUP_DESCRIPTIONS.items():
        if group_name not in yielded and group_name not in HIDDEN_GROUP_NAMES:
            yield group_name, info


VISIBLE_DATASET_KEYS = sorted({key for group, keys in GROUPS.items() if group not in HIDDEN_GROUP_NAMES for key in keys})

SOURCE_LABELS = {
    "molit_unsold": "국토교통부 API",
    "molit_rtms": "국토교통부 실거래가 API",
}

def dataset_source_label(cfg: dict) -> str:
    source_type = cfg.get("source_type")
    if source_type in SOURCE_LABELS:
        return SOURCE_LABELS[source_type]
    org_id = str(cfg.get("orgId", "") or "")
    if org_id == "301":
        return "KOSIS(한국은행)"
    if org_id or cfg.get("tblId"):
        return "KOSIS"
    return "공공데이터 API"

def dataset_detail_comment(key: str, cfg: dict, group_desc: str = "") -> str:
    source = dataset_source_label(cfg)
    note = str(cfg.get("note", "") or "").strip()
    label = group_desc.strip()
    parts = [f"출처: {source}"]
    if label:
        parts.append(label)
    if note and note != label:
        parts.append(note)
    return " | ".join(parts)

COMMON_RENAME = {
    "PRD_DE": "기간",
    "TBL_NM": "통계표명",
    "ORG_NM": "기관명",
    "DT": "값",
    "DAT": "값",
    "UNIT_NM": "단위",
    "ITM_ID": "항목코드",
    "ITM_NM": "항목명",
    "C1": "분류1코드",
    "C1_NM": "분류1",
    "C2": "분류2코드",
    "C2_NM": "분류2",
    "C3": "분류3코드",
    "C3_NM": "분류3",
    "C4": "분류4코드",
    "C4_NM": "분류4",
    "OBJ_L1": "분류1코드",
    "OBJ_L1_NM": "분류1",
    "OBJ_L2": "분류2코드",
    "OBJ_L2_NM": "분류2",
    "OBJ_L3": "분류3코드",
    "OBJ_L3_NM": "분류3",
    "OBJ_L4": "분류4코드",
    "OBJ_L4_NM": "분류4",
    "OBJ_L": "분류코드",
    "OBJ_L_NM": "분류명",
    "C1_OBJ_NM": "분류1_객체명",
    "C2_OBJ_NM": "분류2_객체명",
}


def config_candidate_paths() -> list[Path]:
    """kosis_app_config.json 탐색 우선순위.
    1) 통합 실행기 루트, 2) 현재 작업폴더, 3) 내부 모듈 폴더, 4) AppData 전역 설정, 5) 사용자 홈
    """
    candidates = []
    try:
        candidates.append(app_base_path().parent / "kosis_app_config.json")
    except Exception:
        pass
    try:
        candidates.append(Path.cwd() / "kosis_app_config.json")
    except Exception:
        pass
    try:
        candidates.append(app_base_path() / "kosis_app_config.json")
    except Exception:
        pass
    candidates.append(CONFIG_PATH)
    try:
        candidates.append(Path.home() / "kosis_app_config.json")
    except Exception:
        pass

    seen = set()
    out = []
    for p in candidates:
        try:
            key = str(p.resolve())
        except Exception:
            key = str(p)
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out


def load_app_config():
    """시작 시 kosis_app_config.json을 자동으로 읽어온다."""
    for path in config_candidate_paths():
        try:
            if path.exists() and path.is_file():
                cfg = json.loads(path.read_text(encoding="utf-8"))
                # 읽은 설정을 AppData 전역 설정과 내부 모듈 쪽에도 동기화
                try:
                    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
                    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception:
                    pass
                return cfg
        except Exception:
            continue
    return {}


def save_app_config(data: dict):
    """종료 또는 선택 변경 시 최종 설정을 저장한다."""
    targets = []
    try:
        targets.append(CONFIG_PATH)
    except Exception:
        pass
    try:
        targets.append(app_base_path().parent / "kosis_app_config.json")
    except Exception:
        pass
    try:
        targets.append(app_base_path() / "kosis_app_config.json")
    except Exception:
        pass

    for path in targets:
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def parse_end_period(end_ym: str | None):
    """종료 기준월/기준일을 datetime으로 변환합니다.

    - 월간·연간 통계: YYYY-MM 입력을 기본으로 사용합니다.
    - 주간아파트 통계: YYYY-MM-DD 또는 YYYYMMDD 입력 시 해당 날짜 기준 직전/당일 월요일을 사용합니다.
    """
    if end_ym:
        s = str(end_ym).strip()
        if re.fullmatch(r"\d{8}", s):
            return datetime.strptime(s, "%Y%m%d")
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
            return datetime.strptime(s, "%Y-%m-%d")
        if re.fullmatch(r"\d{6}", s):
            return datetime.strptime(s + "01", "%Y%m%d")
        return datetime.strptime(s + "-01", "%Y-%m-%d")
    now = datetime.now()
    return datetime(now.year, now.month, 1)


def build_year_ranges(end_ym: str | None, lookback_years: int, prd_se: str):
    end_dt = parse_end_period(end_ym)
    end_year = end_dt.year
    start_year = end_year - lookback_years + 1
    ranges = []
    if prd_se.upper() == "Y":
        for year in range(start_year, end_year + 1):
            suffix = f"{year}_{end_dt.month:02d}" if year == end_year else f"{year}_12"
            ranges.append({"year": str(year), "start": str(year), "end": str(year), "suffix": suffix})
    elif prd_se.upper() == "W":
        for year in range(start_year, end_year + 1):
            suffix = f"{year}_{end_dt.month:02d}" if year == end_year else f"{year}_12"
            ranges.append({"year": str(year), "start": f"{year}01", "end": f"{year}53", "suffix": suffix})
    else:
        for year in range(start_year, end_year + 1):
            end_month = end_dt.month if year == end_year else 12
            ranges.append({"year": str(year), "start": f"{year}01", "end": f"{year}{end_month:02d}", "suffix": f"{year}_{end_month:02d}"})
    return ranges


def month_points(start_ym: str, end_ym: str):
    start = datetime.strptime(start_ym + "01", "%Y%m%d")
    end = datetime.strptime(end_ym + "01", "%Y%m%d")
    points = []
    current = start
    while current <= end:
        points.append(current.strftime("%Y%m"))
        current = current + relativedelta(months=1)
    return points


def extract_year_month_from_prd_de(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    m = re.search(r'(20\d{2})[^0-9]?([01]?\d)', s)
    if m:
        year = int(m.group(1))
        month = int(m.group(2))
        if 1 <= month <= 12:
            return year, month
    digits = re.sub(r'\D', '', s)
    if len(digits) >= 6:
        year = int(digits[:4])
        month = int(digits[4:6])
        if 1 <= month <= 12:
            return year, month
    return None


class EmptyKosisDataError(Exception):
    pass


class InvalidKosisParamsError(Exception):
    pass


def month_range(end_period: str, months_back: int = 12):
    end_dt = parse_end_period(end_period)
    start_dt = end_dt - relativedelta(months=months_back - 1)
    values = []
    current = start_dt
    while current <= end_dt:
        values.append(current.strftime("%Y-%m"))
        current += relativedelta(months=1)
    return values


def clean_rate_value(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text or text in {"-", "", "None", "nan"}:
        return None
    text = text.replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group())
    except Exception:
        return None


def normalize_ym(text):
    raw = str(text).strip()
    m = re.search(r"(20\d{2})[^0-9]?(\d{1,2})", raw)
    if not m:
        return None
    return f"{m.group(1)}-{int(m.group(2)):02d}"


class KosisClient:
    def __init__(self, api_key: str, timeout=DEFAULT_TIMEOUT):
        self.api_key = api_key.strip()
        self.timeout = timeout
        self.session = self._create_session()

    def _create_session(self):
        session = requests.Session()
        session.trust_env = False
        retry = Retry(
            total=2, connect=2, read=2, backoff_factor=1.0,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"], raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry, pool_connections=1, pool_maxsize=1)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        session.headers.update({
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json,text/plain,*/*",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            "Referer": "https://kosis.kr/",
            "Connection": "close",
        })
        return session

    def _build_query(self, params: dict):
        return {"method": "getList", "apiKey": self.api_key, "format": "json", "jsonVD": "Y", **params}

    def _parse_json_result(self, data, url):
        if isinstance(data, dict):
            err = str(data.get("err", ""))
            if err == "30":
                raise EmptyKosisDataError(f"데이터 없음: {url}")
            if err in {"20", "21"}:
                msg = "필수요청변수 누락" if err == "20" else "잘못된 요청 변수"
                raise InvalidKosisParamsError(f"{msg}: {url}")
            if err == "31":
                raise RuntimeError(f"조회 결과가 40,000셀을 초과했습니다.\nURL: {url}")
            raise RuntimeError(f"KOSIS 오류 응답: {data}\nURL: {url}")
        return pd.DataFrame(data)

    def _reset_session(self):
        try:
            self.session.close()
        except Exception:
            pass
        self.session = self._create_session()

    def _sleep_with_jitter(self, attempt: int = 0):
        delay = REQUEST_SLEEP + (attempt * KOSIS_RETRY_BACKOFF) + random.uniform(0.4, 1.2)
        time.sleep(delay)

    def _request_via_requests(self, query: dict, attempt: int = 0) -> pd.DataFrame:
        self._sleep_with_jitter(attempt)
        self._reset_session()
        response = self.session.get(BASE_URL, params=query, timeout=self.timeout)
        response.raise_for_status()
        try:
            data = response.json()
        except Exception:
            raise RuntimeError(f"JSON 파싱 실패\nURL: {response.url}\n응답원문: {response.text[:500]}")
        return self._parse_json_result(data, response.url)

    def _request_via_urllib(self, query: dict, attempt: int = 0) -> pd.DataFrame:
        self._sleep_with_jitter(attempt)
        url = BASE_URL + "?" + urllib.parse.urlencode(query)
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "application/json,text/plain,*/*",
                "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                "Referer": "https://kosis.kr/",
                "Connection": "close",
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
            },
            method="GET",
        )
        context = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=self.timeout[1], context=context) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
        try:
            data = json.loads(raw)
        except Exception:
            raise RuntimeError(f"JSON 파싱 실패(urllib)\nURL: {url}\n응답원문: {raw[:500]}")
        return self._parse_json_result(data, url)

    def request(self, params: dict) -> pd.DataFrame:
        query = self._build_query(params)
        last_requests_error = None
        last_urllib_error = None
        for attempt in range(KOSIS_RETRY_ATTEMPTS):
            try:
                return self._request_via_requests(query, attempt=attempt)
            except (EmptyKosisDataError, InvalidKosisParamsError):
                raise
            except Exception as e1:
                last_requests_error = e1
            try:
                return self._request_via_urllib(query, attempt=attempt)
            except (EmptyKosisDataError, InvalidKosisParamsError):
                raise
            except Exception as e2:
                last_urllib_error = e2
                continue
        raise RuntimeError(
            "연결 오류가 발생했습니다.\n"
            f"requests/urllib 재시도 {KOSIS_RETRY_ATTEMPTS}회 모두 실패했습니다.\n"
            f"요청 파라미터: {params}\n"
            f"requests 오류: {last_requests_error}\n"
            f"urllib 오류: {last_urllib_error}"
        )


    def get_statistics_list(self, vw_cd: str = "MT_ZTITLE", parent_list_id: str | None = None, parent_list_ids=None):
        if parent_list_ids is None:
            parent_list_ids = []
        if parent_list_id:
            parent_list_ids = list(parent_list_ids) + [parent_list_id]
        if not parent_list_ids:
            parent_list_ids = [""]

        def _parse(data, url):
            if isinstance(data, dict):
                err = str(data.get("err", ""))
                if err == "30":
                    return []
                if err in {"20", "21"}:
                    msg = "필수요청변수 누락" if err == "20" else "잘못된 요청 변수"
                    raise InvalidKosisParamsError(f"{msg}: {url}")
                raise RuntimeError(f"KOSIS 목록 오류 응답: {data}\nURL: {url}")
            return data if isinstance(data, list) else []

        all_rows = []
        for pid in parent_list_ids:
            query = {
                "method": "getList",
                "apiKey": self.api_key,
                "vwCd": vw_cd,
                "format": "json",
                "jsonVD": "Y",
            }
            if pid:
                query["parentListId"] = pid
            last_requests_error = None
            last_urllib_error = None
            rows = None
            for attempt in range(KOSIS_RETRY_ATTEMPTS):
                try:
                    self._sleep_with_jitter(attempt)
                    self._reset_session()
                    response = self.session.get(STATISTICS_LIST_URL, params=query, timeout=self.timeout)
                    response.raise_for_status()
                    data = response.json()
                    rows = _parse(data, response.url)
                    break
                except (InvalidKosisParamsError, RuntimeError):
                    raise
                except Exception as e1:
                    last_requests_error = e1
                try:
                    self._sleep_with_jitter(attempt)
                    url = STATISTICS_LIST_URL + "?" + urllib.parse.urlencode(query)
                    req = urllib.request.Request(url, headers={
                        "User-Agent": "Mozilla/5.0",
                        "Accept": "application/json,text/plain,*/*",
                        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                        "Referer": "https://kosis.kr/",
                        "Connection": "close",
                    }, method="GET")
                    context = ssl.create_default_context()
                    with urllib.request.urlopen(req, timeout=self.timeout[1], context=context) as resp:
                        raw = resp.read().decode("utf-8", errors="replace")
                    data = json.loads(raw)
                    rows = _parse(data, url)
                    break
                except (InvalidKosisParamsError, RuntimeError):
                    raise
                except Exception as e2:
                    last_urllib_error = e2
            if rows is None:
                raise RuntimeError(
                    "통계 목록 조회 오류가 발생했습니다.\n"
                    f"parentListId={pid or '(없음)'}\n"
                    f"requests 오류: {last_requests_error}\n"
                    f"urllib 오류: {last_urllib_error}"
                )
            for row in rows:
                row = dict(row)
                row["_PARENT_LIST_ID"] = pid or ""
                all_rows.append(row)
        return all_rows


    def get_meta_itm(self, org_id: str, tbl_id: str):
        query = {
            "method": "getMeta",
            "apiKey": self.api_key,
            "format": "json",
            "jsonVD": "Y",
            "type": "ITM",
            "orgId": str(org_id),
            "tblId": str(tbl_id),
        }

        def parse_meta(data, url):
            if isinstance(data, dict):
                err = str(data.get("err", ""))
                if err == "30":
                    return []
                if err in {"20", "21"}:
                    msg = "필수요청변수 누락" if err == "20" else "잘못된 요청 변수"
                    raise InvalidKosisParamsError(f"{msg}: {url}")
                raise RuntimeError(f"KOSIS 메타 오류 응답: {data}\nURL: {url}")
            return data if isinstance(data, list) else []

        last_requests_error = None
        last_urllib_error = None
        for attempt in range(KOSIS_RETRY_ATTEMPTS):
            try:
                self._sleep_with_jitter(attempt)
                self._reset_session()
                response = self.session.get(BASE_URL, params=query, timeout=self.timeout)
                response.raise_for_status()
                try:
                    data = response.json()
                except Exception:
                    raise RuntimeError(f"메타 JSON 파싱 실패\nURL: {response.url}\n응답원문: {response.text[:500]}")
                return parse_meta(data, response.url)
            except (InvalidKosisParamsError, RuntimeError):
                raise
            except Exception as e1:
                last_requests_error = e1
            try:
                self._sleep_with_jitter(attempt)
                url = BASE_URL + "?" + urllib.parse.urlencode(query)
                req = urllib.request.Request(
                    url,
                    headers={
                        "User-Agent": "Mozilla/5.0",
                        "Accept": "application/json,text/plain,*/*",
                        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                        "Referer": "https://kosis.kr/",
                        "Connection": "close",
                        "Cache-Control": "no-cache",
                        "Pragma": "no-cache",
                    },
                    method="GET",
                )
                context = ssl.create_default_context()
                with urllib.request.urlopen(req, timeout=self.timeout[1], context=context) as resp:
                    raw = resp.read().decode("utf-8", errors="replace")
                try:
                    data = json.loads(raw)
                except Exception:
                    raise RuntimeError(f"메타 JSON 파싱 실패(urllib)\nURL: {url}\n응답원문: {raw[:500]}")
                return parse_meta(data, url)
            except (InvalidKosisParamsError, RuntimeError):
                raise
            except Exception as e2:
                last_urllib_error = e2
                continue
        raise RuntimeError(
            "메타 조회 연결 오류가 발생했습니다.\n"
            f"requests/urllib 재시도 {KOSIS_RETRY_ATTEMPTS}회 모두 실패했습니다.\n"
            f"orgId={org_id}, tblId={tbl_id}\n"
            f"requests 오류: {last_requests_error}\n"
            f"urllib 오류: {last_urllib_error}"
        )


class ExcelExporter:
    @staticmethod
    def save(df: pd.DataFrame, folder: Path, filename: str) -> Path:
        ensure_dir(folder)
        path = folder / filename
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="data", index=False)
        return path


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1360x840")
        self.minsize(1180, 760)
        self.configure(bg=BG_COLOR)
        self._configure_styles()
        config = load_app_config()
        # 제출본: API 키는 내장 키로 고정(사용자 수정 불가)
        self.api_key_var = tk.StringVar(master=self, value=DEFAULT_API_KEY)
        self.ecos_api_key_var = tk.StringVar(master=self, value=config.get("ecos_api_key", ""))
        self.rone_api_key_var = tk.StringVar(master=self, value=config.get("rone_api_key") or DEFAULT_RONE_API_KEY)
        # 제출본: 출력 폴더는 실행 폴더의 output으로 자동 지정
        self.output_dir_var = tk.StringVar(master=self, value=str(default_output_dir()))
        saved_end_period = str(config.get("end_period", "") or "")
        self.end_period_var = tk.StringVar(master=self, value=saved_end_period if saved_end_period in SURVEY_MONTHS else DEFAULT_SURVEY_MONTH)
        self.end_period_display_var = tk.StringVar(master=self, value=survey_month_label(self.end_period_var.get()))
        self.end_period_var.trace_add("write", self._on_end_period_var_changed)
        self.worker_var = tk.IntVar(master=self, value=config.get("worker_count", DEFAULT_WORKERS))
        saved_region_file = config.get("region_file", "")
        # EXE/통합 실행기/사용자 저장 경로에서 법정동 코드 파일을 우선 사용합니다.
        # 사용자가 별도 파일을 지정했고 실제 존재하는 경우에만 그 경로를 유지합니다.
        if saved_region_file and Path(saved_region_file).exists():
            initial_region_file = saved_region_file
        elif DEFAULT_REGION_FILE and Path(DEFAULT_REGION_FILE).exists():
            initial_region_file = str(DEFAULT_REGION_FILE)
        else:
            # 통합 실행기와 같은 폴더 또는 현재 작업폴더에 있는 파일을 추가 탐색
            extra_candidates = [
                app_base_path().parent / DEFAULT_REGION_FILENAME,
                Path.cwd() / DEFAULT_REGION_FILENAME,
                Path.home() / DEFAULT_REGION_FILENAME,
            ]
            found = next((str(p) for p in extra_candidates if p.exists()), str(DEFAULT_REGION_FILE))
            initial_region_file = found
        self.region_file_var = tk.StringVar(master=self, value=initial_region_file)
        # 긴 경로가 Entry에 안 보이는 문제를 보완하기 위한 표시용 라벨 변수
        self.output_path_display_var = tk.StringVar(master=self, value=self.output_dir_var.get())
        self.region_path_display_var = tk.StringVar(master=self, value=initial_region_file)
        self.apply_region_to_all_var = tk.BooleanVar(master=self, value=config.get("apply_region_to_all", True))
        self.housing_sido_report_var = tk.BooleanVar(master=self, value=bool(config.get("housing_sido_report", False)))
        self.officetel_sido_report_var = tk.BooleanVar(master=self, value=bool(config.get("officetel_sido_report", False)))
        self._saved_sido = str(config.get("sido", "") or DEFAULT_SIDO)
        self._saved_sigungu = str(config.get("sigungu", "") or DEFAULT_SIGUNGU)
        self._saved_region_code = str(config.get("region_code", "") or DEFAULT_REGION_CODE)
        self.sido_var = tk.StringVar(master=self)
        self.sigungu_var = tk.StringVar(master=self)
        self.region_code_var = tk.StringVar(master=self)
        self.status_var = tk.StringVar(master=self, value="대기 중")
        self.api_progress_percent_var = tk.StringVar(master=self, value="0%")
        self.current_key = None
        self.dataset_checks = {}
        self.dataset_title_labels = {}
        self.dataset_status_labels = {}
        self.dataset_status_state = {}
        self.stop_event = threading.Event()
        self.is_running = False
        self.run_thread = None
        self._thread_region_context = threading.local()
        self.run_button = None
        self.stop_button = None
        self.log_queue = queue.Queue()
        self.region_df = None
        self.region_map = {}
        # 수집 완료 후 보고서 자동 진행(통합 실행기가 콜백을 연결)
        self.auto_pipeline_var = tk.BooleanVar(master=self, value=bool(config.get("auto_pipeline", True)))
        self.on_collection_finished = None
        self._build_ui()
        self._load_region_master(silent=True)
        try:
            self.protocol("WM_DELETE_WINDOW", self._on_window_close)
        except Exception:
            pass
        self.after(150, self._drain_log_queue)

    def _on_end_period_selected(self, event=None):
        idx = self.end_period_combo.current()
        if 0 <= idx < len(SURVEY_MONTHS):
            self.end_period_var.set(SURVEY_MONTHS[idx])
            self._save_basic_config()

    def _on_end_period_var_changed(self, *_args):
        """외부(통합 실행기·설정 다시 읽기)에서 값이 바뀌어도 콤보박스 범위 안의 값으로 맞춘다."""
        value = self.end_period_var.get().strip()
        if value not in SURVEY_MONTHS:
            self.end_period_var.set(DEFAULT_SURVEY_MONTH)
            return
        label = survey_month_label(value)
        if self.end_period_display_var.get() != label:
            self.end_period_display_var.set(label)

    def _on_window_close(self):
        try:
            self._save_basic_config()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass


    def _configure_styles(self):
        """웹 대시보드형 UI 스타일 설정"""
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        base_font = ("맑은 고딕", 10)
        title_font = ("맑은 고딕", 16, "bold")
        section_font = ("맑은 고딕", 11, "bold")

        style.configure(".", background=BG_COLOR, foreground="#0F172A", font=base_font)
        style.configure("TFrame", background=BG_COLOR)
        style.configure("Dashboard.TFrame", background=BG_COLOR)
        style.configure("Card.TFrame", background=CARD_COLOR, relief="solid", borderwidth=1, bordercolor=BORDER_COLOR)
        style.configure("Header.TFrame", background=DASHBOARD_NAV)
        style.configure("HeaderTitle.TLabel", background=DASHBOARD_NAV, foreground="#FFFFFF", font=title_font)
        style.configure("HeaderSub.TLabel", background=DASHBOARD_NAV, foreground=DASHBOARD_NAV_SUB, font=("맑은 고딕", 9))
        style.configure("MetricTitle.TLabel", background=CARD_COLOR, foreground=TEXT_SUB, font=("맑은 고딕", 9, "bold"))
        style.configure("MetricValue.TLabel", background=CARD_COLOR, foreground=TITLE_COLOR, font=("맑은 고딕", 17, "bold"))
        style.configure("Section.TLabel", background=CARD_COLOR, foreground=TITLE_COLOR, font=section_font)
        style.configure("TLabel", background=BG_COLOR, foreground="#0F172A", font=base_font)
        style.configure("CardLabel.TLabel", background=CARD_COLOR, foreground="#0F172A", font=base_font)
        style.configure("RequiredLabel.TLabel", background=CARD_COLOR, foreground="#B91C1C", font=("맑은 고딕", 10, "bold"))
        style.configure("RequiredHint.TLabel", background=CARD_COLOR, foreground="#B91C1C", font=("맑은 고딕", 9, "bold"))
        style.configure("Sub.TLabel", background=BG_COLOR, foreground=TEXT_SUB, font=("맑은 고딕", 9))
        style.configure("CardSub.TLabel", background=CARD_COLOR, foreground=TEXT_SUB, font=("맑은 고딕", 9))
        style.configure("TLabelframe", background=CARD_COLOR, bordercolor=BORDER_COLOR, relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=CARD_COLOR, foreground=TITLE_COLOR, font=section_font)
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#0F172A", bordercolor=BORDER_COLOR,
                        lightcolor=BORDER_COLOR, darkcolor=BORDER_COLOR, padding=8)
        style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", foreground="#0F172A",
                        bordercolor=BORDER_COLOR, lightcolor=BORDER_COLOR, darkcolor=BORDER_COLOR, padding=6)
        style.configure("Required.TCombobox", fieldbackground="#FFF7ED", background="#FDBA74", foreground="#111827",
                        bordercolor="#EA580C", lightcolor="#EA580C", darkcolor="#EA580C", padding=6)
        style.map("Required.TCombobox", fieldbackground=[("readonly", "#FFF7ED")], background=[("readonly", "#FDBA74")])
        style.configure("TSpinbox", arrowsize=12, fieldbackground="#F8FAFC", padding=5)
        style.configure("TCheckbutton", background=CARD_COLOR, foreground="#0F172A", font=("맑은 고딕", 10))
        style.map("TCheckbutton", background=[("active", CARD_COLOR)], foreground=[("disabled", "#94A3B8")])
        style.configure("TScrollbar", background=SECONDARY_BTN, bordercolor=BORDER_COLOR, arrowcolor=TEXT_SUB)

        style.configure("Primary.TButton",
                        background=PRIMARY_BTN, foreground="#FFFFFF",
                        borderwidth=0, focusthickness=0, focuscolor=PRIMARY_BTN,
                        padding=(14, 9), font=("맑은 고딕", 10, "bold"))
        style.map("Primary.TButton",
                  background=[("active", PRIMARY_BTN_ACTIVE), ("pressed", PRIMARY_BTN_ACTIVE), ("disabled", "#93C5FD")],
                  foreground=[("disabled", "#EFF6FF"), ("!disabled", "#FFFFFF")])

        style.configure("Success.TButton",
                        background=SUCCESS_COLOR, foreground="#FFFFFF", borderwidth=0,
                        padding=(12, 8), font=("맑은 고딕", 10, "bold"))
        style.map("Success.TButton", background=[("active", PRIMARY_BTN_ACTIVE), ("pressed", PRIMARY_BTN_ACTIVE)])

        style.configure("Secondary.TButton",
                        background=SECONDARY_BTN, foreground="#0F172A",
                        borderwidth=1, bordercolor=BORDER_COLOR,
                        focusthickness=0, focuscolor=SECONDARY_BTN,
                        padding=(10, 8), font=("맑은 고딕", 10))
        style.map("Secondary.TButton", background=[("active", SECONDARY_BTN_ACTIVE), ("pressed", SECONDARY_BTN_ACTIVE)])

        style.configure("Danger.TButton",
                        background="#FEE2E2", foreground="#991B1B", borderwidth=1, bordercolor="#FCA5A5",
                        padding=(12, 9), font=("맑은 고딕", 10, "bold"))
        style.map("Danger.TButton", background=[("active", "#FECACA"), ("pressed", "#FECACA")])

        style.configure("TProgressbar", troughcolor="#E2E8F0", background=PRIMARY_BTN,
                        bordercolor=BORDER_COLOR, lightcolor=PRIMARY_BTN, darkcolor=PRIMARY_BTN)
        style.configure("Dashboard.TNotebook", background=CARD_COLOR, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure("Dashboard.TNotebook.Tab", background="#F8FAFC", foreground="#334155",
                        padding=(13, 8), font=("맑은 고딕", 10, "bold"))
        style.map("Dashboard.TNotebook.Tab",
                  background=[("selected", PRIMARY_BTN), ("active", "#DBEAFE")],
                  foreground=[("selected", "#FFFFFF"), ("active", TITLE_COLOR)])
        style.configure("GroupHeader.TLabel", background="#EFF6FF", foreground=TITLE_COLOR,
                        font=("맑은 고딕", 10, "bold"), padding=(10, 6))
        style.configure("ItemDesc.TLabel", background=CARD_COLOR, foreground=TEXT_SUB, font=("맑은 고딕", 9))
        style.configure("StatTitle.TLabel", background=CARD_COLOR, foreground="#111827", font=("맑은 고딕", 10, "bold"))
        style.configure("StatTitleRunning.TLabel", background=CARD_COLOR, foreground="#DC2626", font=("맑은 고딕", 10, "bold"))
        style.configure("StatTitleFailed.TLabel", background=CARD_COLOR, foreground="#B91C1C", font=("맑은 고딕", 10, "bold"))
        style.configure("StatusIdle.TLabel", background=CARD_COLOR, foreground="#94A3B8", font=("맑은 고딕", 9, "bold"))
        style.configure("StatusRunning.TLabel", background="#FEE2E2", foreground="#DC2626", font=("맑은 고딕", 9, "bold"), padding=(8, 3))
        style.configure("StatusDone.TLabel", background="#DCFCE7", foreground="#166534", font=("맑은 고딕", 9, "bold"), padding=(8, 3))
        style.configure("StatusFailed.TLabel", background="#FEE2E2", foreground="#B91C1C", font=("맑은 고딕", 9, "bold"), padding=(8, 3))
        style.configure("StatusSkipped.TLabel", background="#FEF3C7", foreground="#92400E", font=("맑은 고딕", 9, "bold"), padding=(8, 3))

    def _save_basic_config(self, force_clear_region: bool = False):
        """기본 설정 저장.
        시도·시군구·REG가 비어 있을 때 기존 저장 지역값을 실수로 덮어쓰지 않도록 보호한다.
        지역 해제 버튼을 누른 경우에만 force_clear_region=True로 빈 값 저장.
        """
        prev = load_app_config()
        sido = self.sido_var.get().strip()
        sigungu = self.sigungu_var.get().strip()
        region_code = self.region_code_var.get().strip()

        if not force_clear_region and not (sido and sigungu and region_code):
            sido = str(prev.get("sido", "") or "")
            sigungu = str(prev.get("sigungu", "") or "")
            region_code = str(prev.get("region_code", "") or "")

        save_app_config({
            "api_key": self.api_key_var.get().strip(),
            "ecos_api_key": self.ecos_api_key_var.get().strip(),
            "rone_api_key": self.rone_api_key_var.get().strip(),
            "output_dir": self.output_dir_var.get().strip(),
            "end_period": self.end_period_var.get().strip(),
            "worker_count": int(self.worker_var.get()),
            "region_file": self.region_file_var.get().strip(),
            "apply_region_to_all": bool(self.apply_region_to_all_var.get()),
            "housing_sido_report": bool(self.housing_sido_report_var.get()),
            "officetel_sido_report": bool(self.officetel_sido_report_var.get()),
            "auto_pipeline": bool(self.auto_pipeline_var.get()) if hasattr(self, "auto_pipeline_var") else True,
            "sido": sido,
            "sigungu": sigungu,
            "region_code": region_code,
        })

    def _build_ui(self):
        # ---------- 상단 웹 대시보드 헤더 ----------
        header = ttk.Frame(self, padding=(20, 16), style="Header.TFrame")
        header.pack(fill="x")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="부동산 AI 동향 리포터", style="HeaderTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=f"KOSIS 통합수집 · 통계 1개당 엑셀 1개 생성 · {APP_VERSION} · {APP_AUTHOR} · 최종수정일 {APP_UPDATED}",
            style="HeaderSub.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        # 실행 버튼은 업무그룹 패널의 "선택 통계 실행"으로 통일합니다.
        # 상단에는 작업 중단 버튼만 유지해 화면 중복을 줄입니다.
        self.stop_button = ttk.Button(header, text="■ 중단", style="Danger.TButton", command=self._stop_run, state="disabled")
        self.stop_button.grid(row=0, column=1, rowspan=2, sticky="e", padx=(18, 0))

        # ---------- 요약 카드 ----------
        summary = ttk.Frame(self, padding=(16, 14, 16, 6), style="Dashboard.TFrame")
        summary.pack(fill="x")
        for col in range(4):
            summary.columnconfigure(col, weight=1)

        total_stats = len(VISIBLE_DATASET_KEYS)
        active_stats = sum(1 for k in VISIBLE_DATASET_KEYS if not (DATASETS.get(k, {}).get("status") == "hold" or k in HOLD_DATASET_KEYS))
        self._make_metric_card(summary, 0, "화면 표시 통계", f"{total_stats}개", "개발완료 항목만 정리")
        self._make_metric_card(summary, 1, "실행 가능", f"{active_stats}개", "현재 자동 수집 가능한 항목")
        self._make_metric_card(summary, 2, "저장 원칙", "1통계=1엑셀", "data/raw/schema 시트 통합")
        self._make_metric_card(summary, 3, "수집 기준", "최근 13개월", "지표별 최신 발표월 기준")

        # ---------- 페이지 컨트롤러: 자료 수집 / LLM 보고서 ----------
        controller = ttk.Frame(self, padding=(16, 0, 16, 6), style="Dashboard.TFrame")
        controller.pack(fill="x")
        self.api_page_button = ttk.Button(controller, text="① 공공통계·실거래 자료 수집", style="Success.TButton", command=self._show_api_collect_page)
        self.api_page_button.pack(side="left", padx=(0, 8))
        self.llm_page_button = ttk.Button(controller, text="② LLM 기반 시장동향 보고서 생성", style="Secondary.TButton", command=self._show_llm_report_page)
        self.llm_page_button.pack(side="left")
        ttk.Label(controller, text="KOSIS·국토교통부 API 수집부터 LLM 기반 종합 시황 보고서까지 단계별로 검수하며 처리합니다.", style="CardSub.TLabel").pack(side="left", padx=(14, 0))

        # ---------- 본문 레이아웃 ----------
        self.api_collect_main = ttk.Frame(self, padding=(16, 8, 16, 10), style="Dashboard.TFrame")
        main = self.api_collect_main
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1, minsize=360)
        main.columnconfigure(1, weight=2, minsize=470)
        main.columnconfigure(2, weight=2, minsize=430)
        main.rowconfigure(0, weight=1)

        settings_card = ttk.Frame(main, padding=12, style="Card.TFrame")
        settings_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_settings_panel(settings_card)

        list_card = ttk.Frame(main, padding=12, style="Card.TFrame")
        list_card.grid(row=0, column=1, sticky="nsew", padx=8)
        self._build_left_panel(list_card)

        log_card = ttk.Frame(main, padding=12, style="Card.TFrame")
        log_card.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        self._build_right_panel(log_card)

        # ---------- 하단 실행 상태바 ----------
        self.api_collect_footer = ttk.Frame(self, padding=(16, 4, 16, 10), style="Dashboard.TFrame")
        footer = self.api_collect_footer
        footer.pack(fill="x")
        footer.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(footer, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Label(footer, textvariable=self.api_progress_percent_var, style="Sub.TLabel", width=7, anchor="center").grid(row=0, column=1, sticky="e", padx=(0, 12))
        ttk.Label(footer, textvariable=self.status_var, style="Sub.TLabel").grid(row=0, column=2, sticky="e")

        # LLM·보고서 페이지는 통합 실행기에서 내용을 주입합니다.
        self.llm_integrated_frame = ttk.Frame(self, padding=(16, 8, 16, 10), style="Dashboard.TFrame")
        self.llm_integrated_frame.pack(fill="both", expand=True)
        self.llm_integrated_frame.pack_forget()

    def _show_api_collect_page(self):
        """API 통계 수집 페이지 표시."""
        try:
            if hasattr(self, "llm_integrated_frame"):
                self.llm_integrated_frame.pack_forget()
            if hasattr(self, "api_collect_main"):
                self.api_collect_main.pack(fill="both", expand=True)
            if hasattr(self, "api_collect_footer"):
                self.api_collect_footer.pack(fill="x")
            if hasattr(self, "api_page_button"):
                self.api_page_button.configure(style="Success.TButton")
            if hasattr(self, "llm_page_button"):
                self.llm_page_button.configure(style="Secondary.TButton")
            self.status_var.set("공공통계·실거래 자료 수집 페이지")
        except Exception:
            pass

    def _show_llm_report_page(self):
        """LLM·보고서 생성 페이지 표시."""
        try:
            if hasattr(self, "api_collect_main"):
                self.api_collect_main.pack_forget()
            if hasattr(self, "api_collect_footer"):
                self.api_collect_footer.pack_forget()
            if hasattr(self, "llm_integrated_frame"):
                self.llm_integrated_frame.pack(fill="both", expand=True)
            if hasattr(self, "api_page_button"):
                self.api_page_button.configure(style="Secondary.TButton")
            if hasattr(self, "llm_page_button"):
                self.llm_page_button.configure(style="Success.TButton")
            self.status_var.set("LLM 기반 시장동향 보고서 생성 페이지")
        except Exception:
            pass

    def _make_metric_card(self, parent, col, title, value, desc):
        card = ttk.Frame(parent, padding=(13, 10), style="Card.TFrame")
        card.grid(row=0, column=col, sticky="ew", padx=5)
        ttk.Label(card, text=title, style="MetricTitle.TLabel").pack(anchor="w")
        ttk.Label(card, text=value, style="MetricValue.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(card, text=desc, style="CardSub.TLabel").pack(anchor="w", pady=(4, 0))

    def _sync_path_labels(self):
        """출력폴더와 법정동 파일 경로를 화면에 명확하게 표시."""
        try:
            self.output_path_display_var.set(self.output_dir_var.get().strip())
        except Exception:
            pass
        try:
            self.region_path_display_var.set(self.region_file_var.get().strip())
        except Exception:
            pass

    def _first_time_guide_text(self) -> str:
        return (
            "API 자료 수집 처음 사용 안내\n"
            "==============================\n\n"
            "이 화면은 보고서에 필요한 KOSIS·국토교통부·실거래 API 자료를 엑셀로 수집하는 단계입니다.\n\n"
            "1. 수집 설정\n"
            "- KOSIS API KEY와 출력 폴더(실행 폴더의 output)는 자동으로 설정됩니다.\n"
            "- 조사기준월을 선택합니다. 예: 2026년 8월\n"
            "- 시도와 시군구는 필수입니다. 선택하지 않으면 실행할 수 없습니다.\n\n"
            "2. 업무별 분석자료 선택\n"
            "- 가운데 탭에서 주택동향업무(주간), 주택동향업무(월간), 오피스텔동향업무(월간) 중 필요한 업무를 선택합니다.\n"
            "- 체크박스는 보고서에 반영할 통계자료입니다.\n"
            "- [그룹 전체선택]을 누르면 현재 업무 탭의 통계를 한 번에 선택할 수 있습니다.\n\n"
            "3. 선택 자료 수집 실행\n"
            "- [선택 자료 수집 실행]을 누르면 체크된 통계만 순서대로 수집합니다.\n"
            "- 실행 전 선택통계실행 폴더는 자동으로 비워집니다. 다른 지역·다른 기준월 자료가 섞이는 것을 막기 위한 동작입니다.\n"
            "- 진행률은 버튼 아래 프로그레스바와 화면 하단 상태바에서 확인할 수 있습니다.\n"
            "- 실행 중인 항목은 빨간색, 완료/실패 항목은 배지로 표시됩니다.\n\n"
            "4. 수집 완료 후\n"
            "- 저장된 엑셀은 출력 폴더의 선택통계실행 폴더에 생성됩니다.\n"
            "- 통합 실행기 화면으로 돌아가 [수집자료 점검] 후 [LLM 분석자료 생성]을 진행합니다.\n\n"
            "주의\n"
            "- 주간동향은 매주 월요일 기준 공표통계를 사용합니다.\n"
            "- 자료가 미공표된 월/주는 일부 항목이 실패로 표시될 수 있습니다.\n"
            "- 실패 항목이 있으면 우측 로그에서 원인을 확인한 뒤 필요 시 해당 통계만 다시 실행하세요."
        )

    def _show_first_time_guide_popup(self):
        try:
            popup = tk.Toplevel(self)
            popup.title("처음 사용 안내")
            popup.geometry("840x620")
            popup.transient(self)
            container = ttk.Frame(popup, padding=14, style="Card.TFrame")
            container.pack(fill="both", expand=True)
            ttk.Label(container, text="처음 사용 안내", style="Header.TLabel").pack(anchor="w")
            text = tk.Text(container, wrap="word", font=("맑은 고딕", 10), bg="#FFFFFF", fg="#0F172A", relief="solid", bd=1)
            text.pack(fill="both", expand=True, pady=(10, 8))
            text.insert("1.0", self._first_time_guide_text())
            text.configure(state="disabled")
            ttk.Button(container, text="닫기", style="Secondary.TButton", command=popup.destroy).pack(anchor="e")
            try:
                popup.lift()
                popup.focus_force()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("오류", f"처음 사용 안내 표시 실패\n{e}")

    def _build_settings_panel(self, parent):
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(2, weight=0)

        ttk.Label(parent, text="자료 수집 기준 설정", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(parent, text="조사지역과 기준월을 정하면 KOSIS·국토교통부·실거래 API 자료를 보고서 목적에 맞게 수집합니다.", style="CardSub.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 8))
        ttk.Label(parent, text="※ 조사기준월·시도·시군구는 필수 선택 항목입니다.", style="RequiredHint.TLabel").grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 12))

        row = 3
        ttk.Label(parent, text="KOSIS API KEY", style="CardLabel.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        tk.Entry(parent, textvariable=self.api_key_var, show="*", state="readonly", readonlybackground="#E2E8F0", fg="#0F172A", relief="solid", bd=1).grid(row=row, column=1, columnspan=2, sticky="ew", pady=5, ipady=5)
        row += 1
        # 한국부동산원 통계(가격지수·전월세전환율·주간아파트·오피스텔·시군구 지가)는 이 키가 있으면 R-ONE에서 직접 받는다.
        ttk.Label(parent, text="R-ONE API KEY", style="CardLabel.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        rone_entry = tk.Entry(parent, textvariable=self.rone_api_key_var, show="*", fg="#0F172A", relief="solid", bd=1)
        rone_entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5, ipady=5)
        rone_entry.bind("<FocusOut>", lambda _e: self._save_basic_config())
        row += 1
        ttk.Label(parent, text="※ 한국부동산원 통계는 R-ONE에서 먼저 받고, 키가 없거나 실패하면 KOSIS에서 받습니다.", style="CardSub.TLabel").grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 6))
        row += 1

        # 출력 폴더는 실행 폴더의 output으로 자동 지정되므로 경로만 표시한다.
        ttk.Label(parent, text="출력 폴더(자동)", style="CardLabel.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Label(parent, textvariable=self.output_path_display_var, style="CardSub.TLabel", wraplength=340).grid(row=row, column=1, columnspan=2, sticky="w", pady=5)
        row += 1

        ttk.Label(parent, text="* 조사기준월", style="RequiredLabel.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        self.end_period_combo = ttk.Combobox(parent, textvariable=self.end_period_display_var, state="readonly", style="Required.TCombobox",
                                             values=tuple(survey_month_label(v) for v in SURVEY_MONTHS))
        self.end_period_combo.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
        self.end_period_combo.bind("<<ComboboxSelected>>", self._on_end_period_selected)
        row += 1

        ttk.Label(parent, text="동시 수집 수", style="CardLabel.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Spinbox(parent, from_=1, to=5, textvariable=self.worker_var, width=6, command=self._save_basic_config).grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(parent, text="※ 자료별 병렬 실행, 시도취합은 자료 내부 시군구 순차 처리", style="CardSub.TLabel", wraplength=300).grid(row=row, column=2, sticky="w", pady=5)
        row += 1

        option_frame = ttk.Frame(parent, style="Card.TFrame")
        option_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        option_frame.columnconfigure(0, weight=1)
        option_frame.columnconfigure(1, weight=1)
        ttk.Label(option_frame, text="보고서 범위 옵션", style="CardLabel.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 3))
        ttk.Checkbutton(option_frame, text="선택 지역 적용", variable=self.apply_region_to_all_var,
                        command=self._save_basic_config).grid(row=1, column=0, sticky="w", pady=2)
        ttk.Checkbutton(option_frame, text="주택 월간 시도단위", variable=self.housing_sido_report_var,
                        command=self._save_basic_config).grid(row=1, column=1, sticky="w", pady=2)
        ttk.Checkbutton(option_frame, text="오피스텔 시도단위", variable=self.officetel_sido_report_var,
                        command=self._save_basic_config).grid(row=2, column=0, sticky="w", pady=2)
        ttk.Label(option_frame, text="선택 시도 전체 시군구를 수집해 시도_전체 파일로 취합합니다.", style="CardSub.TLabel", wraplength=360).grid(row=2, column=1, sticky="w", pady=2)
        row += 1

        # 법정동 코드는 로컬 DB(local_db/region_codes.sqlite3)에서 자동으로 불러오므로 파일 선택 UI는 표시하지 않는다.
        ttk.Button(parent, text="저장 설정 다시 읽기", style="Secondary.TButton", command=self._reload_saved_config).grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
        row += 1

        ttk.Label(parent, text="* 시도", style="RequiredLabel.TLabel").grid(row=row, column=0, sticky="w", pady=(14, 5))
        self.sido_combo = ttk.Combobox(parent, textvariable=self.sido_var, state="readonly", style="Required.TCombobox")
        self.sido_combo.grid(row=row, column=1, columnspan=2, sticky="ew", pady=(14, 5))
        self.sido_combo.bind("<<ComboboxSelected>>", self._on_sido_change)
        row += 1

        ttk.Label(parent, text="* 시군구", style="RequiredLabel.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        self.sigungu_combo = ttk.Combobox(parent, textvariable=self.sigungu_var, state="readonly", style="Required.TCombobox")
        self.sigungu_combo.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
        self.sigungu_combo.bind("<<ComboboxSelected>>", self._on_sigungu_change)
        row += 1

        ttk.Label(parent, text="REG 5자리", style="CardLabel.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        tk.Entry(parent, textvariable=self.region_code_var, width=12, bg="#F8FAFC", fg="#0F172A", relief="solid", bd=1, state="readonly").grid(row=row, column=1, sticky="w", pady=5, ipady=5)
        ttk.Button(parent, text="지역 해제", style="Secondary.TButton", command=self._clear_region_selection).grid(row=row, column=2, sticky="e", pady=5)
        row += 1
        ttk.Label(parent, text="조사지역 미선택 시 실행할 수 없습니다.", style="RequiredHint.TLabel").grid(row=row, column=0, columnspan=3, sticky="w", pady=(2, 8))
        row += 1

        ttk.Separator(parent, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=14)
        row += 1

        ttk.Button(parent, text="처음 사용 안내", style="Secondary.TButton", command=self._show_first_time_guide_popup).grid(row=row, column=0, columnspan=3, sticky="ew", pady=4)
        row += 1
        ttk.Button(parent, text="통계 설명자료 보기", style="Secondary.TButton", command=self._show_stats_guide_popup).grid(row=row, column=0, columnspan=3, sticky="ew", pady=4)
        self._sync_path_labels()


    def _build_left_panel(self, parent):
        """업무그룹별 간소화 페이지 컨트롤러형 통계 선택 패널."""
        self.group_checks = {}
        self.group_pages = {}
        self.group_key_map = {}
        self.group_select_buttons = {}
        self.group_canvas_map = {}
        self.group_inner_map = {}
        self.dataset_rows = {}
        self.dataset_group_map = {}
        parent.rowconfigure(4, weight=1)
        parent.columnconfigure(0, weight=1)

        ttk.Label(parent, text="업무별 분석자료 선택", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            parent,
            text="보고서에 반영할 통계자료를 선택하세요. 실행 중인 항목과 실패 항목은 화면에서 즉시 확인할 수 있습니다.",
            style="CardSub.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(4, 10))

        action_bar = ttk.Frame(parent, style="Card.TFrame")
        action_bar.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        action_bar.columnconfigure(3, weight=1)
        ttk.Button(action_bar, text="그룹 전체선택", style="Secondary.TButton", command=self._select_current_group).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(action_bar, text="그룹 선택해제", style="Secondary.TButton", command=self._clear_current_group).grid(row=0, column=1, padx=6)
        self.run_button = ttk.Button(action_bar, text="선택 자료 수집 실행", style="Success.TButton", command=self._run_selected)
        self.run_button.grid(row=0, column=2, padx=6)
        self.current_group_label_var = tk.StringVar(master=self, value="")
        ttk.Label(action_bar, textvariable=self.current_group_label_var, style="CardSub.TLabel").grid(row=0, column=3, sticky="e", padx=(12, 0))
        ttk.Checkbutton(
            action_bar,
            text="수집 완료 후 자동 진행(자료점검 → LLM 입력자료 생성)",
            variable=self.auto_pipeline_var,
            command=self._save_basic_config,
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))

        progress_bar = ttk.Frame(parent, style="Card.TFrame")
        progress_bar.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        progress_bar.columnconfigure(0, weight=1)
        self.inline_progress = ttk.Progressbar(progress_bar, mode="determinate", maximum=100)
        self.inline_progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(progress_bar, textvariable=self.api_progress_percent_var, style="CardSub.TLabel", width=7, anchor="center").grid(row=0, column=1, sticky="e")
        ttk.Label(progress_bar, textvariable=self.status_var, style="CardSub.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        notebook = ttk.Notebook(parent, style="Dashboard.TNotebook")
        notebook.grid(row=4, column=0, sticky="nsew")
        self.group_notebook = notebook

        for group_name, info in iter_visible_group_descriptions():
            page = ttk.Frame(notebook, padding=10, style="Card.TFrame")
            page.rowconfigure(1, weight=1)
            page.columnconfigure(0, weight=1)
            notebook.add(page, text=info["title"])
            self.group_pages[str(page)] = group_name

            header = ttk.Frame(page, style="Card.TFrame")
            header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
            header.columnconfigure(1, weight=1)
            ttk.Label(header, text=info["title"], style="GroupHeader.TLabel").grid(row=0, column=0, sticky="w")
            count_label = ttk.Label(header, text="", style="CardSub.TLabel")
            count_label.grid(row=0, column=1, sticky="e")
            self.group_select_buttons[group_name] = count_label

            outer = ttk.Frame(page, style="Card.TFrame")
            outer.grid(row=1, column=0, sticky="nsew")
            outer.rowconfigure(0, weight=1)
            outer.columnconfigure(0, weight=1)

            canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0, bg=CARD_COLOR)
            scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
            inner = ttk.Frame(canvas, style="Card.TFrame")
            inner.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
            canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.bind("<Configure>", lambda e, c=canvas, w=canvas_window: c.itemconfigure(w, width=e.width))
            canvas.grid(row=0, column=0, sticky="nsew")
            scrollbar.grid(row=0, column=1, sticky="ns")
            self.group_canvas_map[group_name] = canvas
            self.group_inner_map[group_name] = inner

            group_keys = []
            visible_items = []
            for key, desc in info["items"]:
                cfg = DATASETS.get(key)
                if not cfg:
                    continue
                is_hold = cfg.get("status") == "hold" or key in HOLD_DATASET_KEYS
                if is_hold:
                    continue
                visible_items.append((key, cfg, desc))

            for idx, (key, cfg, desc) in enumerate(visible_items, start=1):
                row = ttk.Frame(inner, padding=(8, 5), style="Card.TFrame")
                row.grid(row=idx - 1, column=0, sticky="ew")
                row.columnconfigure(1, weight=1)
                row.columnconfigure(2, minsize=70)

                if key not in self.dataset_checks:
                    self.dataset_checks[key] = tk.BooleanVar(master=self, value=False)
                var = self.dataset_checks[key]
                group_keys.append(key)

                chk = ttk.Checkbutton(row, text=f"{idx:02d}", variable=var, command=self._refresh_group_counts)
                chk.grid(row=0, column=0, sticky="w", padx=(0, 8))
                title = ttk.Label(row, text=cfg.get("label", key), style="StatTitle.TLabel")
                title.grid(row=0, column=1, sticky="w")
                self.dataset_title_labels.setdefault(key, []).append(title)
                status_label = ttk.Label(row, text="대기", style="StatusIdle.TLabel")
                status_label.grid(row=0, column=2, sticky="e", padx=(8, 0))
                self.dataset_status_labels.setdefault(key, []).append(status_label)
                source_label = ttk.Label(row, text=f"출처: {dataset_source_label(cfg)}", style="StatusSkipped.TLabel")
                source_label.grid(row=1, column=1, sticky="w", pady=(2, 0))
                desc_label = ttk.Label(row, text=dataset_detail_comment(key, cfg, desc), style="ItemDesc.TLabel", wraplength=470)
                desc_label.grid(row=2, column=1, columnspan=2, sticky="w", pady=(1, 0))
                self.dataset_rows.setdefault(key, []).append(row)
                self.dataset_group_map.setdefault(key, group_name)

            self.group_key_map[group_name] = group_keys

        notebook.bind("<<NotebookTabChanged>>", self._on_group_tab_changed)
        self.after(100, self._refresh_group_counts)
        self.after(120, self._on_group_tab_changed)

    def _scroll_to_dataset(self, key: str):
        """실행 중인 통계 행이 보이도록 업무그룹 탭과 스크롤 위치를 자동 조정."""
        active_group = getattr(self, "active_run_group_name", None)
        if active_group and key in getattr(self, "group_key_map", {}).get(active_group, []):
            group_name = active_group
        else:
            group_name = getattr(self, "dataset_group_map", {}).get(key)
        if not group_name:
            return
        notebook = getattr(self, "group_notebook", None)
        try:
            for page_id, name in getattr(self, "group_pages", {}).items():
                if name == group_name:
                    notebook.select(page_id)
                    break
        except Exception:
            pass
        canvas = getattr(self, "group_canvas_map", {}).get(group_name)
        inner = getattr(self, "group_inner_map", {}).get(group_name)
        rows = getattr(self, "dataset_rows", {}).get(key, [])
        if not canvas or not inner or not rows:
            return
        try:
            self.update_idletasks()
            row = next((candidate for candidate in rows if candidate.master is inner), rows[0])
            y = max(0, row.winfo_y() - 24)
            total_h = max(1, inner.winfo_height())
            canvas.yview_moveto(min(1.0, y / total_h))
        except Exception:
            pass

    def _get_current_group_name(self):
        notebook = getattr(self, "group_notebook", None)
        if not notebook:
            return None
        try:
            current_page = notebook.select()
            return self.group_pages.get(current_page)
        except Exception:
            return None

    def _set_group_selection(self, group_name: str, selected: bool):
        keys = self.group_key_map.get(group_name, GROUPS.get(group_name, []))
        for key in keys:
            if key in HOLD_DATASET_KEYS or DATASETS.get(key, {}).get("status") == "hold":
                continue
            if key in self.dataset_checks:
                self.dataset_checks[key].set(selected)
        self._refresh_group_counts()

    def _select_current_group(self):
        group_name = self._get_current_group_name()
        if group_name:
            self._set_group_selection(group_name, True)

    def _clear_current_group(self):
        group_name = self._get_current_group_name()
        if group_name:
            self._set_group_selection(group_name, False)

    def _run_current_group(self):
        group_name = self._get_current_group_name()
        if not group_name:
            messagebox.showerror("오류", "현재 선택된 업무그룹을 확인할 수 없습니다.")
            return
        self._toggle_all(False)
        self._set_group_selection(group_name, True)
        title = GROUP_DESCRIPTIONS.get(group_name, {}).get("title", group_name)
        self.log(f"[선택실행] {title} 전체선택 후 실행")
        self._run_selected()

    def _on_group_tab_changed(self, event=None):
        group_name = self._get_current_group_name()
        if not group_name:
            return
        title = GROUP_DESCRIPTIONS.get(group_name, {}).get("title", group_name)
        keys = self.group_key_map.get(group_name, [])
        active_count = sum(1 for k in keys if k not in HOLD_DATASET_KEYS and DATASETS.get(k, {}).get("status") != "hold")
        self.current_group_label_var.set(f"현재 그룹: {title} · 실행 가능 {active_count}개")
        self._refresh_group_counts()

    def _refresh_group_counts(self):
        for group_name, label in getattr(self, "group_select_buttons", {}).items():
            keys = self.group_key_map.get(group_name, [])
            active_keys = [k for k in keys if k not in HOLD_DATASET_KEYS and DATASETS.get(k, {}).get("status") != "hold"]
            selected = sum(1 for k in active_keys if self.dataset_checks.get(k) and self.dataset_checks[k].get())
            label.configure(text=f"선택 {selected}/{len(active_keys)}")

    def _reset_dataset_visual_status(self, keys=None):
        target_keys = list(keys) if keys is not None else list(getattr(self, "dataset_title_labels", {}).keys())
        for key in target_keys:
            self._set_dataset_visual_status(key, "idle")

    def _set_dataset_visual_status(self, key: str, status: str, detail: str | None = None):
        """통계별 실행 상태를 화면 체크리스트에 즉시 반영합니다."""
        if not key:
            return

        title_style = "StatTitle.TLabel"
        status_style = "StatusIdle.TLabel"
        status_text = "대기"

        if status == "running":
            title_style = "StatTitleRunning.TLabel"
            status_style = "StatusRunning.TLabel"
            status_text = "진행중"
            self._scroll_to_dataset(key)
        elif status == "done":
            title_style = "StatTitle.TLabel"
            status_style = "StatusDone.TLabel"
            status_text = "완료"
        elif status == "failed":
            title_style = "StatTitleFailed.TLabel"
            status_style = "StatusFailed.TLabel"
            status_text = "실패"
        elif status == "skipped":
            title_style = "StatTitle.TLabel"
            status_style = "StatusSkipped.TLabel"
            status_text = "건너뜀"

        self.dataset_status_state[key] = {"status": status, "detail": detail or ""}
        for label in self.dataset_title_labels.get(key, []):
            try:
                label.configure(style=title_style)
            except Exception:
                pass
        for label in self.dataset_status_labels.get(key, []):
            try:
                label.configure(text=status_text, style=status_style)
            except Exception:
                pass

    def _toggle_all(self, selected: bool):
        for key, var in self.dataset_checks.items():
            if key in HOLD_DATASET_KEYS or DATASETS.get(key, {}).get("status") == "hold":
                var.set(False)
            else:
                var.set(selected)
        self._refresh_group_counts()

    def _build_right_panel(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        ttk.Label(parent, text="실행 로그", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.log_text = tk.Text(
            parent,
            wrap="word",
            bg="#0B1220",
            fg="#E5E7EB",
            relief="flat",
            font=("Consolas", 10),
            padx=14,
            pady=14,
            insertbackground="#E5E7EB",
            selectbackground="#1D4ED8",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

    def _select_dir(self):
        folder = filedialog.askdirectory(parent=self, initialdir=self.output_dir_var.get().strip() or str(Path.cwd()))
        if folder:
            self.output_dir_var.set(folder)
            self._sync_path_labels()
            self._save_basic_config()
            self.update_idletasks()

    def _select_region_file(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="법정동 코드 파일 선택",
            initialdir=str(Path(self.region_file_var.get().strip()).parent) if self.region_file_var.get().strip() else str(Path.cwd()),
            filetypes=[("Excel", "*.xls *.xlsx"), ("All", "*.*")]
        )
        if path:
            self.region_file_var.set(path)
            self._sync_path_labels()
            self._save_basic_config()
            self.update_idletasks()
            # 파일 선택 즉시 법정동 코드 로드 및 시도·시군구 콤보박스 반영
            self._load_region_master(silent=False)

    def _reload_saved_config(self):
        """kosis_app_config.json을 다시 읽어 화면에 반영한다."""
        try:
            cfg = load_app_config()
            self.ecos_api_key_var.set(cfg.get("ecos_api_key", self.ecos_api_key_var.get()))
            self.rone_api_key_var.set(cfg.get("rone_api_key") or self.rone_api_key_var.get() or DEFAULT_RONE_API_KEY)
            self.end_period_var.set(cfg.get("end_period", self.end_period_var.get()))
            self.region_file_var.set(cfg.get("region_file", self.region_file_var.get()))
            self._sync_path_labels()
            self.apply_region_to_all_var.set(bool(cfg.get("apply_region_to_all", self.apply_region_to_all_var.get())))
            self.housing_sido_report_var.set(bool(cfg.get("housing_sido_report", False)))
            self.officetel_sido_report_var.set(bool(cfg.get("officetel_sido_report", False)))
            self._saved_sido = str(cfg.get("sido", "") or DEFAULT_SIDO)
            self._saved_sigungu = str(cfg.get("sigungu", "") or DEFAULT_SIGUNGU)
            self._saved_region_code = str(cfg.get("region_code", "") or DEFAULT_REGION_CODE)
            self._load_region_master(silent=False)
            self.status_var.set("저장 설정 다시 읽기 완료")
        except Exception as e:
            messagebox.showerror("오류", f"저장 설정 다시 읽기 실패\\n{e}")

    def _clear_region_selection(self):
        self.sido_var.set("")
        self.sigungu_var.set("")
        self.region_code_var.set("")
        self.sigungu_combo["values"] = ()
        self.status_var.set("조사지역 미선택")
        try:
            self._save_basic_config(force_clear_region=True)
        except Exception:
            pass

    def _normalize_region_file_path(self, raw_path: str) -> Path:
        """사용자가 직접 입력한 법정동 파일 경로를 정리한다.
        따옴표, 앞뒤 공백, 슬래시 혼용, 환경변수, ~ 경로를 보정한다.
        """
        s = str(raw_path or "").strip().strip('"').strip("'")
        if not s:
            return DEFAULT_REGION_FILE
        try:
            s = os.path.expandvars(s)
        except Exception:
            pass
        try:
            return Path(s).expanduser()
        except Exception:
            return Path(s)

    def _load_region_master(self, silent=True):
        self._sync_path_labels()
        try:
            db_df = load_region_codes_from_db()
        except Exception as e:
            db_df = None
            self.log(f"법정동 코드 로컬 DB 읽기 실패, 엑셀 파일로 대체합니다: {e}")
        if db_df is not None and not db_df.empty:
            self._apply_region_dataframe(db_df, f"로컬 DB: {region_db_path()}", silent)
            return
        # 1순위: 화면에 표시된 경로
        # 2순위: EXE/통합실행기/현재작업폴더/사용자홈에 있는 기본 법정동 코드 파일
        raw_path = self.region_file_var.get().strip()
        candidate = self._normalize_region_file_path(raw_path) if raw_path else DEFAULT_REGION_FILE
        fallback = bundled_resource_path(DEFAULT_REGION_FILENAME)

        # 사용자가 직접 입력한 경로가 없거나 존재하지 않으면 여러 후보를 순차 탐색
        candidates = [candidate, fallback]
        for fname in [DEFAULT_REGION_FILENAME, "전국 법정동 지역.xlsx"]:
            candidates.extend([
                app_base_path() / fname,
                app_base_path().parent / fname,
                Path.cwd() / fname,
                Path.home() / fname,
            ])

        path = None
        for p in candidates:
            try:
                if p and Path(p).exists() and Path(p).is_file():
                    path = Path(p)
                    break
            except Exception:
                continue

        if path and path.exists():
            self.region_file_var.set(str(path))
            self._sync_path_labels()
        else:
            if not silent:
                messagebox.showwarning(
                    "안내",
                    "법정동 코드 파일을 찾을 수 없습니다.\n\n"
                    f"입력/선택 경로: {candidate}\n"
                    f"기본 탐색 경로: {fallback}\n\n"
                    "파일명 예시: 전국 법정동 지역.xls 또는 전국 법정동 지역.xlsx\n"
                    "통합 실행기와 같은 폴더에 두거나 [파일] 버튼으로 직접 선택해 주세요."
                )
            return

        if path.suffix.lower() == ".xls" and xlrd is None:
            if not silent:
                messagebox.showerror(
                    "오류",
                    "현재 Python 환경에 xlrd 패키지가 없어 .xls 법정동 파일을 읽을 수 없습니다.\n\n"
                    "해결 방법 1) pip install xlrd\n"
                    "해결 방법 2) 법정동 파일을 .xlsx로 저장 후 다시 선택"
                )
            return

        try:
            # xlsx는 openpyxl, xls는 xlrd를 우선 사용
            if path.suffix.lower() == ".xlsx":
                df = pd.read_excel(path, engine="openpyxl")
            elif path.suffix.lower() == ".xls":
                df = pd.read_excel(path, engine="xlrd")
            else:
                df = pd.read_excel(path)
            required_cols = ['REG', '시도', '시군구']
            for c in required_cols:
                if c not in df.columns:
                    raise ValueError(f"필수 컬럼 누락: {c}")
        except Exception as e:
            if not silent:
                messagebox.showerror(
                    "오류",
                    f"법정동 코드 파일 로드 실패\n\n파일: {path}\n오류: {e}\n\n"
                    "필수 컬럼은 REG, 시도, 시군구 입니다."
                )
            self.status_var.set("법정동 코드 로드 실패")
            return
        self._apply_region_dataframe(df, f"파일: {path}", silent)

    def _apply_region_dataframe(self, df: pd.DataFrame, source_text: str, silent=True):
        """REG/시도/시군구 데이터프레임을 시도·시군구 콤보박스에 반영한다."""
        required_cols = ['REG', '시도', '시군구']
        try:
            df = df[required_cols].copy()
            df['REG'] = df['REG'].astype(str).str.extract(r'(\d+)')[0].str.zfill(5)
            df = df[df['REG'].str.len() == 5]
            df = df[df['시도'].notna() & df['시군구'].notna()].copy()
            df['시도'] = df['시도'].astype(str).str.strip()
            df['시군구'] = df['시군구'].astype(str).str.strip()
            df = df.drop_duplicates(subset=['REG', '시도', '시군구']).sort_values(['시도', '시군구']).reset_index(drop=True)
            self.region_df = df
            self.region_map = {
                sido: sdf[['시군구', 'REG']].drop_duplicates().sort_values('시군구')
                for sido, sdf in df.groupby('시도')
            }
            self.sido_combo["values"] = tuple(sorted(self.region_map.keys()))

            # 저장된 최종 시도·시군구 선택값 복원
            try:
                if self._saved_sido and self._saved_sido in self.region_map:
                    self.sido_var.set(self._saved_sido)
                    vals = tuple(self.region_map[self._saved_sido]['시군구'].tolist())
                    self.sigungu_combo["values"] = vals
                    if self._saved_sigungu and self._saved_sigungu in vals:
                        self.sigungu_var.set(self._saved_sigungu)
                        self._on_sigungu_change()
                    elif self._saved_region_code:
                        row = df[df["REG"].astype(str).str.zfill(5) == str(self._saved_region_code).zfill(5)]
                        if not row.empty:
                            self.sido_var.set(str(row.iloc[0]["시도"]))
                            vals = tuple(self.region_map[self.sido_var.get()]['시군구'].tolist())
                            self.sigungu_combo["values"] = vals
                            self.sigungu_var.set(str(row.iloc[0]["시군구"]))
                            self.region_code_var.set(str(row.iloc[0]["REG"]).zfill(5))
            except Exception:
                pass

            if not silent:
                self.log(f"법정동 코드 로드 완료: {len(df):,}건 | {source_text}")
                messagebox.showinfo("완료", f"법정동 코드 로드 완료\n{len(df):,}건\n\n{source_text}")
            self.status_var.set(f"법정동 코드 로드 완료: {len(df):,}건")
            self._save_basic_config()
            self.update_idletasks()
        except Exception as e:
            if not silent:
                messagebox.showerror("오류", f"법정동 코드 로드 실패\n\n{source_text}\n오류: {e}")
            self.status_var.set("법정동 코드 로드 실패")

    def _on_sido_change(self, event=None):
        sido = self.sido_var.get().strip()
        self.sigungu_var.set("")
        self.region_code_var.set("")
        if sido and sido in self.region_map:
            vals = tuple(self.region_map[sido]['시군구'].tolist())
            self.sigungu_combo["values"] = vals
        else:
            self.sigungu_combo["values"] = ()
        # 시도만 바뀐 상태에서는 시군구가 비어 있으므로 기존 저장 지역을 덮어쓰지 않음
        try:
            self._save_basic_config()
        except Exception:
            pass

    def _on_sigungu_change(self, event=None):
        sido = self.sido_var.get().strip()
        sigungu = self.sigungu_var.get().strip()
        self.region_code_var.set("")
        if sido and sigungu and sido in self.region_map:
            sdf = self.region_map[sido]
            row = sdf[sdf['시군구'] == sigungu]
            if not row.empty:
                self.region_code_var.set(str(row.iloc[0]['REG']).zfill(5))
                self._saved_sido = sido
                self._saved_sigungu = sigungu
                self._saved_region_code = self.region_code_var.get().strip()
        try:
            self._save_basic_config()
        except Exception:
            pass

    def _get_sido_sigungu_rows(self, sido: str = ""):
        sido = (sido or self.sido_var.get()).strip()
        if not sido or sido not in self.region_map:
            return []
        group_name = getattr(self, "active_run_group_name", None) or self._get_current_group_name()
        gwangju_area_only = (
            sido == "전남광주통합특별시"
            and (
                self._officetel_sido_report_enabled(group_name)
                or self._housing_sido_report_enabled(group_name)
            )
        )
        rows = []
        for _, r in self.region_map[sido].iterrows():
            sigungu = str(r.get("시군구", "")).strip()
            reg = str(r.get("REG", "")).strip().zfill(5)
            if gwangju_area_only and sigungu not in GWANGJU_AREA_SIGUNGU_NAMES:
                continue
            if sigungu and reg:
                rows.append({"시도": sido, "시군구": sigungu, "REG": reg})
        if gwangju_area_only:
            order = {name: idx for idx, name in enumerate(["동구", "서구", "남구", "북구", "광산구"])}
            rows.sort(key=lambda row: order.get(row["시군구"], 999))
        return rows

    def _set_region_selection_values(self, sido: str, sigungu: str = "", reg: str = ""):
        self.sido_var.set(str(sido or "").strip())
        if self.sido_var.get() in self.region_map:
            self.sigungu_combo["values"] = tuple(self.region_map[self.sido_var.get()]["시군구"].tolist())
        self.sigungu_var.set(str(sigungu or "").strip())
        self.region_code_var.set(str(reg or "").strip().zfill(5) if reg else "")

    def _set_thread_region_context(self, sido: str, sigungu: str = "", reg: str = ""):
        self._thread_region_context.sido = str(sido or "").strip()
        self._thread_region_context.sigungu = str(sigungu or "").strip()
        self._thread_region_context.reg = str(reg or "").strip().zfill(5) if reg else ""

    def _clear_thread_region_context(self):
        for name in ("sido", "sigungu", "reg"):
            try:
                delattr(self._thread_region_context, name)
            except Exception:
                pass

    def _current_region_values(self):
        sido = getattr(self._thread_region_context, "sido", None)
        sigungu = getattr(self._thread_region_context, "sigungu", None)
        reg = getattr(self._thread_region_context, "reg", None)
        if sido is not None or sigungu is not None or reg is not None:
            return str(sido or ""), str(sigungu or ""), str(reg or "")
        return self.sido_var.get().strip(), self.sigungu_var.get().strip(), self.region_code_var.get().strip()

    def _officetel_sido_report_enabled(self, group_name: str) -> bool:
        return (
            str(group_name) == MONTHLY_OFFICETEL_GROUP
            and bool(self.officetel_sido_report_var.get())
            and bool(self.sido_var.get().strip())
        )

    def _housing_sido_report_enabled(self, group_name: str) -> bool:
        return (
            str(group_name) == MONTHLY_HOUSING_GROUP
            and bool(self.housing_sido_report_var.get())
            and bool(self.sido_var.get().strip())
        )

    def _sido_report_enabled(self, group_name: str) -> bool:
        return self._officetel_sido_report_enabled(group_name) or self._housing_sido_report_enabled(group_name)

    def _get_kab_sigungu_code(self):
        """주택가격동향조사 KOSIS 지역코드(c###) 반환.

        KOSIS 주택가격동향조사/주간아파트 표는 법정동 REG 5자리 대신
        c501(목포시) 같은 자체 지역코드를 사용한다.
        """
        _sido, sigungu, reg = self._current_region_values()
        code = KAB_SIGUNGU_CODE_MAP.get(sigungu)
        if code:
            self.log(f"[KOSIS지역코드변환] {sigungu} -> {code}")
            return code

        if reg:
            # 보정표에 없는 지역은 임시 fallback. 신규 지역은 KAB_SIGUNGU_CODE_MAP에 추가 권장.
            fallback = "c" + reg[-3:]
            self.log(f"[KOSIS지역코드주의] {sigungu or reg} 매핑 없음 | 임시코드={fallback} | 필요 시 보정표 추가")
            return fallback
        return ""

    def _get_selected_region_code(self, code_length: int):
        _sido, _sigungu, reg = self._current_region_values()
        if not reg:
            return ""
        if code_length == 2:
            return reg[:2]
        if code_length == 3:
            return self._get_kab_sigungu_code()
        return reg[:5]

    def _get_kosis_region_code(self, code_length: int):
        sido, sigungu, reg = self._current_region_values()
        if not reg:
            return ""
        key = (str(sido or "").strip(), str(sigungu or "").strip())
        legacy_reg = KOSIS_LEGACY_REGION_CODE_MAP.get(key) or KOSIS_LEGACY_REGION_CODE_MAP.get(str(reg).zfill(5))
        legacy_sido = KOSIS_LEGACY_SIDO_CODE_MAP.get(key) or KOSIS_LEGACY_SIDO_CODE_MAP.get(str(reg)[:2])
        if legacy_reg or legacy_sido:
            self.log(f"[KOSIS행정구역구코드] {sido} {sigungu} {reg} -> {legacy_sido or str(reg)[:2]}/{legacy_reg or str(reg)[:5]}")
        if code_length == 2:
            return legacy_sido or (legacy_reg[:2] if legacy_reg else reg[:2])
        if code_length == 3:
            return self._get_kab_sigungu_code()
        return legacy_reg or reg[:5]

    def _kosis_region_code_variants(self, code_length: int) -> list:
        """행정구역 개편 전후 KOSIS 지역코드 목록.
        2026-07 통합처럼 특정 월부터 코드가 바뀌는 경우가 있어 구코드·신코드를 모두 사용한다.
        """
        sido, sigungu, reg = self._current_region_values()
        if not reg:
            return []
        legacy = self._get_kosis_region_code(code_length)
        current = reg[:2] if code_length == 2 else reg[:5]
        codes = []
        for c in [legacy, current]:
            c = str(c or "").strip()
            if c and c not in codes:
                codes.append(c)
        return codes

    def _request_with_region_variants(self, client: KosisClient, cfg: dict, params: dict, label: str):
        """구코드로 조회한 뒤, 행정구역 개편으로 코드가 바뀐 월을 신코드로 보완 조회한다."""
        code_length = int(cfg.get("region_code_length", 5) or 5)
        variants = self._kosis_region_code_variants(code_length)
        region_keys = [k for k, v in params.items() if str(k).startswith("objL") and variants and str(v) == variants[0]]
        frames = []
        first_error = None
        invalid_error = None
        codes = variants if (len(variants) > 1 and region_keys) else [None]
        for code in codes:
            call_params = dict(params)
            if code is not None:
                for k in region_keys:
                    call_params[k] = code
            try:
                df = client.request(call_params)
            except EmptyKosisDataError as e:
                first_error = first_error or e
                continue
            except InvalidKosisParamsError as e:
                # 통계표에 등록되지 않은 지역코드(개편 신코드 등)는 KOSIS가 err 21로 거절한다.
                # 다른 코드로 조회가 가능하므로 전체 실패로 처리하지 않는다.
                if code is None:
                    raise
                invalid_error = invalid_error or e
                self.log(f"[지역코드미등록] {label} | 코드={code} | 통계표에 없는 코드라 건너뜀")
                continue
            if df is None or df.empty:
                continue
            if "DT" in df.columns:
                work = df.copy()
                work["DT"] = pd.to_numeric(work["DT"], errors="coerce")
                if work["DT"].notna().sum() == 0 or float(work["DT"].abs().sum(skipna=True) or 0) == 0:
                    # 개편 전 코드가 0값만 남아 있는 경우는 자료 없음으로 본다.
                    self.log(f"[지역코드0값] {label} | 코드={code} | 값이 모두 0이라 다른 코드로 재조회")
                    continue
            frames.append(df)
            if code is not None and len(codes) > 1:
                self.log(f"[지역코드적용] {label} | 코드={code} | {len(df):,}건")
        if not frames:
            if first_error:
                raise first_error
            if invalid_error:
                raise invalid_error
            raise EmptyKosisDataError(f"0건 응답: {label}")
        if len(frames) == 1:
            return frames[0]
        # 같은 월이 두 코드에서 모두 조회되면 먼저 조회한 코드(종전 코드)를 우선하고,
        # 그 코드에 없는 월만 다른 코드로 보완한다.
        if all("PRD_DE" in f.columns for f in frames):
            picked = []
            covered = set()
            for f in frames:
                work = f.copy()
                if "DT" in work.columns:
                    work["_dt"] = pd.to_numeric(work["DT"], errors="coerce").fillna(0).abs()
                    valid_months = {str(m) for m, v in work.groupby(work["PRD_DE"].astype(str))["_dt"].sum().items() if v > 0}
                    work = work.drop(columns=["_dt"])
                else:
                    valid_months = set(work["PRD_DE"].astype(str))
                new_months = valid_months - covered
                if not new_months:
                    continue
                picked.append(work[work["PRD_DE"].astype(str).isin(new_months)])
                covered |= new_months
            frames = picked or frames
        merged = pd.concat(frames, ignore_index=True)
        # 개편 전후 코드가 같은 월에 함께 잡히면, 값이 모두 0인 쪽(이관된 구코드) 행을 제거한다.
        try:
            code_col = next(
                (c for c in ["C1", "C2", "C3"] if c in merged.columns and merged[c].astype(str).isin(variants).any()),
                None,
            )
            if code_col and "PRD_DE" in merged.columns and "DT" in merged.columns:
                work = merged.copy()
                work["_dt"] = pd.to_numeric(work["DT"], errors="coerce").fillna(0).abs()
                sums = work.groupby(["PRD_DE", code_col])["_dt"].sum()
                month_max = sums.groupby(level=0).max()
                drop_keys = {key for key, value in sums.items() if value == 0 and month_max.get(key[0], 0) > 0}
                if drop_keys:
                    keep = [
                        (str(a), str(b)) not in {(str(k[0]), str(k[1])) for k in drop_keys}
                        for a, b in zip(merged["PRD_DE"].astype(str), merged[code_col].astype(str))
                    ]
                    merged = merged[pd.Series(keep, index=merged.index)]
                    self.log(f"[구코드0값행제거] {label} | {len(drop_keys)}개 월·코드 조합 제거")
        except Exception:
            pass
        subset = [c for c in ["PRD_DE", "C1", "C2", "C3", "ITM_ID"] if c in merged.columns]
        if subset:
            merged = merged.drop_duplicates(subset=subset, keep="first")
        return merged.reset_index(drop=True)

    def _get_region_name_suffix(self):
        sido, sigungu, _reg = self._current_region_values()
        if sido and sigungu:
            return f"{sido}_{sigungu}"
        if sido:
            return f"{sido}_전체"
        return "전국_전체"

    def _build_output_filename(self, base_filename: str):
        stem = Path(base_filename).stem
        ext = Path(base_filename).suffix or ".xlsx"
        region_suffix = self._get_region_name_suffix().replace(" ", "")
        stem = stem.replace("__", "_").rstrip("_")
        return f"{stem}_{region_suffix}{ext}"

    def _apply_region_filter(self, cfg: dict):
        # 오피스텔 가격지수 등 KOSIS가 전국/시도 단위로 발표되는 통계는
        # 화면의 법정동 시군구 선택값을 적용하지 않는다.
        if cfg.get("no_region_filter"):
            cfg2 = deepcopy(cfg)
            if cfg2.get("kab_price_region_filter"):
                
                if not self._current_region_values()[2]:
                    raise ValueError(f"{cfg2['label']}은 시도/시군구 콤보박스에서 시군구를 먼저 선택해야 합니다.")
                self.log(f"[전체조회후처리] {cfg2['label']} | objL1=ALL, objL2=ALL 전체 조회 후 선택 시도+시군구만 추출")
            else:
                self.log(f"[지역필터제외] {cfg2['label']} | KOSIS 발표단위 기준(objL1=ALL, objL2=ALL)으로 조회")
            return cfg2

        if cfg.get("require_region") and not self._current_region_values()[2]:
            raise ValueError(f"{cfg['label']}은 시도/시군구 콤보박스에서 시군구를 먼저 선택해야 합니다.")

        if cfg.get("auto_region_probe"):
            cfg2 = deepcopy(cfg)
            if self._current_region_values()[2]:
                region_code = self._get_kosis_region_code(int(cfg2.get("region_code_length", 5)))
                self.log(f"[지역필터대기] {cfg2['label']} | 코드={region_code} | 차원 자동탐색 예정")
            return cfg2

        if not self._current_region_values()[2]:
            return cfg

        if not self.apply_region_to_all_var.get() and not cfg.get("require_region"):
            return cfg

        cfg2 = deepcopy(cfg)
        params = deepcopy(cfg2.get("params", {}))
        key = cfg2.get("region_param_key")
        code_length = int(cfg2.get("region_code_length", 5))
        region_code = self._get_kosis_region_code(code_length)

        if key and region_code:
            params[key] = region_code
            cfg2["params"] = params
            self.log(f"[지역필터] {cfg2['label']} | 코드={region_code} | {key} 적용")
            return cfg2

        if region_code:
            applied_key = None
            for k in ["objL4", "objL3", "objL2", "objL1", "objL"]:
                if params.get(k) == "ALL":
                    params[k] = region_code
                    applied_key = k
                    break
            if applied_key is None and "objL1" not in params:
                params["objL1"] = region_code
                applied_key = "objL1(추가)"
            cfg2["params"] = params
            if applied_key:
                self.log(f"[지역필터] {cfg2['label']} | 코드={region_code} | {applied_key} 적용")
            else:
                self.log(f"[지역필터미적용] {cfg2['label']} | 코드={region_code}")
        return cfg2


    def _apply_group_selection(self, group_name: str):
        selected = bool(self.group_checks[group_name].get())
        for key in GROUPS.get(group_name, []):
            if key in HOLD_DATASET_KEYS:
                continue
            if key in self.dataset_checks:
                self.dataset_checks[key].set(selected)

    def _toggle_all_groups_on(self):
        if getattr(self, "group_checks", None):
            for group_name, var in self.group_checks.items():
                var.set(True)
                self._apply_group_selection(group_name)
        else:
            self._toggle_all(True)
        self._refresh_group_counts()

    def _toggle_all_groups_off(self):
        for group_name, var in getattr(self, "group_checks", {}).items():
            var.set(False)
        for var in self.dataset_checks.values():
            var.set(False)
        self._refresh_group_counts()

    def _build_stats_guide_text(self):
        lines = []
        lines.append("개발완료 통계 리스트 및 업무별 활용 설명")
        lines.append("")
        for group_name, info in iter_visible_group_descriptions():
            lines.append(f"[{info['title']}]")
            for key, desc in info["items"]:
                if key in HOLD_DATASET_KEYS or DATASETS.get(key, {}).get("status") == "hold":
                    continue
                label = DATASETS.get(key, {}).get("label", key)
                lines.append(f"- {label}: {desc}")
            lines.append("")
        lines.append("공통 안내")
        lines.append("- 동일 파일이 있으면 삭제 후 다시 생성됩니다.")
        lines.append("- 엑셀 1개 안에 data, raw, schema 시트가 함께 저장됩니다.")
        lines.append("- 업무그룹별 탭에서 필요한 통계만 선택해 실행할 수 있습니다.")
        return "\n".join(lines)

    def _show_stats_guide_popup(self):
        popup = tk.Toplevel(self)
        popup.title("개발완료 통계 설명자료")
        popup.geometry("980x760")
        popup.configure(bg=BG_COLOR)

        container = ttk.Frame(popup, padding=14)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="개발완료 통계 리스트 및 활용 설명", style="Header.TLabel").pack(anchor="w")
        ttk.Label(container, text="직원 설명용 안내문입니다. 필요한 경우 복사해서 전달할 수 있습니다.", style="Sub.TLabel").pack(anchor="w", pady=(4, 12))

        text_frame = ttk.Labelframe(container, text="설명자료", padding=10)
        text_frame.pack(fill="both", expand=True)

        text_widget = tk.Text(
            text_frame, wrap="word", bg=CARD_COLOR, fg="#111827", relief="flat",
            font=("맑은 고딕", 10), padx=12, pady=12, insertbackground="#111827"
        )
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        guide_text = self._build_stats_guide_text()
        text_widget.insert("1.0", guide_text)

        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill="x", pady=(10, 0))

        def copy_to_clipboard():
            popup.clipboard_clear()
            popup.clipboard_append(text_widget.get("1.0", "end-1c"))
            self.log("[안내] 통계 설명자료를 클립보드에 복사했습니다.")

        ttk.Button(btn_frame, text="클립보드 복사", style="Secondary.TButton", command=copy_to_clipboard).pack(side="left")
        ttk.Button(btn_frame, text="닫기", style="Primary.TButton", command=popup.destroy).pack(side="right")

    def _fetch_migration_sublists(self):
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("오류", "KOSIS API KEY를 먼저 입력하세요.")
            return
        self.log("[이동자수하위조회] 1차/2차 자동 조회 시작")
        try:
            client = KosisClient(api_key)
            seed_rows = [
                {"LIST_ID": pid, "LIST_NM": title, "_DEPTH": 1, "_PARENT_LIST_ID": "A_1", "_SOURCE": "seed"}
                for pid, title in MIGRATION_SUBLIST_IDS
            ]
            rows = []
            for pid, _title in MIGRATION_SUBLIST_IDS:
                self.log(f"[이동자수하위조회] parentListId={pid} 조회")
                sub_rows = client.get_statistics_list(parent_list_id=pid)
                for r in sub_rows:
                    rr = dict(r)
                    rr["_PARENT_LIST_ID"] = pid
                    rr["_DEPTH"] = 2
                    rr["_SOURCE"] = "api"
                    rows.append(rr)
            all_rows = seed_rows + rows
            df = pd.DataFrame(all_rows)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_output_dir = Path(self.output_dir_var.get().strip() or str(Path.cwd() / "output"))
            out_dir = base_output_dir / SIMPLE_RUN_FOLDER / "이동자수_하위목록조회"
            out_dir.mkdir(parents=True, exist_ok=True)
            json_path = out_dir / f"이동자수_하위목록조회_{stamp}.json"
            xlsx_path = out_dir / f"이동자수_하위목록조회_{stamp}.xlsx"
            txt_path = out_dir / f"이동자수_하위목록조회_{stamp}.txt"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(all_rows, f, ensure_ascii=False, indent=2)
            if not df.empty:
                df.to_excel(xlsx_path, index=False)
            report_lines = ["이동자수 계열 하위 목록 1차/2차 자동 조회 결과", ""]
            for pid, title in MIGRATION_SUBLIST_IDS:
                sub = [r for r in rows if str(r.get("_PARENT_LIST_ID", "")) == pid]
                report_lines.append(f"[{pid}] {title} | 하위 {len(sub)}건")
                for r in sub[:100]:
                    tbl_id = r.get("TBL_ID") or ""
                    list_id = r.get("LIST_ID") or ""
                    name = r.get("TBL_NM") or r.get("LIST_NM") or ""
                    org_id = r.get("ORG_ID") or ""
                    rec = r.get("REC_TBL_SE") or ""
                    report_lines.append(f"- 명칭: {name} | TBL_ID: {tbl_id} | LIST_ID: {list_id} | ORG_ID: {org_id} | 추천:{rec}")
                if len(sub) > 100:
                    report_lines.append(f"- ... 나머지 {len(sub) - 100}건은 엑셀/JSON 참고")
                report_lines.append("")
            tbl_rows = [r for r in rows if str(r.get("TBL_ID") or "").strip()]
            report_lines.append(f"실제 TBL_ID 보유 행 수: {len(tbl_rows)}건")
            txt_path.write_text("\n".join(report_lines), encoding="utf-8")
            self.log(f"[이동자수하위조회완료] 총 {len(all_rows)}건 | 실표 {len(tbl_rows)}건 | {xlsx_path}")
            popup = tk.Toplevel(self)
            popup.title("이동자수 계열 하위 목록 1차/2차 자동 조회 결과")
            popup.geometry("980x700")
            popup.configure(bg=BG_COLOR)
            frame = ttk.Frame(popup, padding=14)
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="이동자수 계열 하위 목록 1차/2차 자동 조회 결과", style="Header.TLabel").pack(anchor="w")
            ttk.Label(frame, text=f"JSON / TXT / XLSX 저장 완료: {out_dir}", style="Sub.TLabel").pack(anchor="w", pady=(4, 10))
            text_widget = tk.Text(frame, wrap="word", bg=CARD_COLOR, fg="#111827", relief="flat", font=("맑은 고딕", 10), padx=12, pady=12)
            scroll = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scroll.set)
            text_widget.insert("1.0", "\n".join(report_lines))
            text_widget.pack(side="left", fill="both", expand=True)
            scroll.pack(side="right", fill="y")
        except Exception as e:
            self.log(f"[이동자수하위조회실패] {e}")
            messagebox.showerror("오류", str(e))
    def _sanitize_log_text(self, text: str) -> str:
        """로그에 API URL과 인증키가 노출되지 않도록 정리합니다."""
        try:
            s = str(text)
        except Exception:
            s = ""
        s = re.sub(r"https?://\S+", "[URL 생략]", s)
        s = re.sub(r"URL\s*:\s*\S+", "URL: [생략]", s)
        s = re.sub(r"(apiKey|serviceKey)=([^&\s]+)", r"\1=[생략]", s, flags=re.IGNORECASE)
        return s

    def log(self, text: str):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{now}] {self._sanitize_log_text(text)}")

    def _drain_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
        except queue.Empty:
            pass
        self.after(150, self._drain_log_queue)

    def _update_progress(self, done: int, total: int):
        value = (done / total) * 100 if total else 0
        self.progress["value"] = value
        if hasattr(self, "inline_progress"):
            self.inline_progress["value"] = value
        self.api_progress_percent_var.set(f"{value:.0f}%")
        self.status_var.set(f"진행률 {done}/{total} ({value:.0f}%)")

    def _validate_common_inputs(self):
        api_key = self.api_key_var.get().strip()
        output_dir = self.output_dir_var.get().strip()
        end_period = self.end_period_var.get().strip()

        if not api_key:
            messagebox.showerror("오류", "KOSIS API KEY가 설정되어 있지 않습니다.")
            return None
        if not output_dir:
            output_dir = str(default_output_dir())
            self.output_dir_var.set(output_dir)
        if not end_period:
            messagebox.showwarning("조사기준월 입력 필요", "조사기준월을 입력해 주세요. 예: 2026-08")
            return None
        if end_period and not re.fullmatch(r"\d{4}-\d{2}|\d{4}-\d{2}-\d{2}|\d{8}|\d{6}", end_period):
            messagebox.showerror("오류", "조사기준월은 YYYY-MM 형식으로 입력하세요. 예: 2026-08\n(주간 기준일은 YYYY-MM-DD 형식도 가능)")
            return None
        if not self.sido_var.get().strip() or not self.sigungu_var.get().strip() or not self.region_code_var.get().strip():
            self.status_var.set("조사지역을 선택해 주세요")
            messagebox.showwarning("조사지역 선택 필요", "조사지역을 선택해 주세요.\n\n좌측 필수 항목에서 시도와 시군구를 먼저 선택한 뒤 실행해 주세요.")
            try:
                self.sido_combo.focus_set()
            except Exception:
                pass
            return None

        self._save_basic_config()
        return api_key, output_dir, end_period


    def _clear_simple_run_folder(self, output_dir: Path):
        """선택 통계 실행 전 선택통계실행 폴더를 초기화한다.
        - Excel/브라우저/미리보기/이전 실행 프로세스가 파일을 잡고 있으면 WinError 32가 발생할 수 있다.
        - 삭제 재시도 후, 가능하면 잠금 보관 폴더로 이름을 변경하고 새 폴더를 생성한다.
        - 완전 삭제가 불가능한 경우에는 원인과 조치 방법을 명확히 안내한다.
        """
        target = Path(output_dir) / SIMPLE_RUN_FOLDER

        def _handle_remove_readonly(func, path, exc_info):
            try:
                os.chmod(path, 0o777)
                func(path)
            except Exception:
                pass

        if target.exists():
            last_error = None
            removed = False

            # 파일 핸들이 바로 해제되지 않는 경우가 있어 짧게 여러 번 재시도
            for attempt in range(1, 7):
                try:
                    import gc
                    gc.collect()
                    shutil.rmtree(target, onerror=_handle_remove_readonly)
                    removed = True
                    self.log(f"[기존자료전체삭제] 선택통계실행 폴더 삭제: {target}")
                    break
                except Exception as e:
                    last_error = e
                    self.log(f"[기존자료삭제재시도] {attempt}/6 | {target} | {e}")
                    time.sleep(1.0)

            # 삭제가 계속 실패하면, 폴더명 변경으로 새 실행 폴더를 확보한다.
            # 단, 다른 프로세스가 디렉터리 자체를 강하게 잡고 있으면 rename도 실패할 수 있다.
            if not removed and target.exists():
                archive = Path(output_dir) / f"{SIMPLE_RUN_FOLDER}_잠금보관_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                try:
                    target.rename(archive)
                    self.log(f"[기존자료잠금보관] 삭제 실패 폴더를 보관명으로 변경: {archive}")
                except Exception:
                    msg = (
                        "선택통계실행 폴더 또는 내부 엑셀 파일이 다른 프로세스에서 사용 중입니다.\n"
                        f"대상 폴더: {target}\n\n"
                        "조치 방법:\n"
                        "1) 열려 있는 엑셀 파일을 모두 닫아주세요.\n"
                        "2) 보고서 HTML/PDF 미리보기 창과 탐색기 미리보기 창을 닫아주세요.\n"
                        "3) 이전에 실행한 Python/EXE 프로세스가 있으면 작업 관리자에서 종료해주세요.\n"
                        "4) 다시 실행하면 기존자료 삭제 후 새로 수집됩니다.\n\n"
                        f"원본 오류: {last_error}"
                    )
                    self.log(f"[기존자료삭제실패-프로세스사용중] {target} | {last_error}")
                    raise PermissionError(msg)

        target.mkdir(parents=True, exist_ok=True)
        self.log(f"[자료폴더재생성] 선택통계실행 폴더 재생성: {target}")

    def _get_selected_group_names(self):
        return []

    def _get_group_folders_for_cfg(self, cfg: dict):
        return [SIMPLE_RUN_FOLDER]

    def _remove_existing_file_if_exists(self, path: Path):
        if path.exists():
            try:
                path.unlink()
                self.log(f"[기존파일삭제] {path}")
            except Exception as e:
                self.log(f"[기존파일삭제실패] {path} | {e}")

    def _save_single_stat_excel(self, folder: Path, filename: str, df: pd.DataFrame, dataset_name: str) -> str:
        """통계 1개당 엑셀 1개만 저장합니다.

        기존에는 data 파일 외에 _raw, _schema, AI업로드용 복제 파일이 함께 생성될 수 있었으나,
        현재 버전부터는 하나의 엑셀 파일 안에 필요한 시트만 넣습니다.
        - data: AI/보고서용 표준화 컬럼
        - raw: 원본 컬럼 보존
        - schema: 컬럼 설명
        """
        folder = ensure_dir(folder)
        path = folder / filename
        self._remove_existing_file_if_exists(path)

        ai_df, _rename_map = self._standardize_ai_columns(df)
        schema_df = self._build_schema_dataframe(df, filename, dataset_name)

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            ai_df.to_excel(writer, sheet_name="data", index=False)
            df.to_excel(writer, sheet_name="raw", index=False)
            schema_df.to_excel(writer, sheet_name="schema", index=False)

        self.log(f"[저장완료-통계1개파일1개] {path}")
        return str(path)

    def _read_saved_stat_data_sheet(self, path: Path) -> pd.DataFrame:
        try:
            return pd.read_excel(path, sheet_name="raw")
        except Exception:
            try:
                return pd.read_excel(path, sheet_name="data")
            except Exception:
                return pd.read_excel(path)

    def _sido_total_filename_from_saved(self, first_path: Path, sido: str) -> str:
        stem = Path(first_path).stem
        ext = Path(first_path).suffix or ".xlsx"
        for row in self._get_sido_sigungu_rows(sido):
            suffix = f"_{row['시도']}_{row['시군구']}".replace(" ", "")
            if stem.endswith(suffix):
                stem = stem[: -len(suffix)]
                break
        label = "전남광주통합특별시_광주권역" if sido == "전남광주통합특별시" and len(self._get_sido_sigungu_rows(sido)) == 5 else sido
        return f"{stem}_{label}_전체{ext}".replace(" ", "")

    def _sido_total_single_fetch_required(self, cfg: dict) -> bool:
        if cfg.get("sido_total_single_fetch"):
            return True
        return bool(cfg.get("no_region_filter") and cfg.get("source_type") != "molit_rtms")

    def _run_one_cfg_sido_total(self, client: KosisClient, output_dir: Path, cfg: dict, end_period: str):
        sido = self.sido_var.get().strip()
        rows = self._get_sido_sigungu_rows(sido)
        if not rows:
            raise ValueError("시도 보고서 생성을 위해 선택 시도에 포함된 시군구 목록이 필요합니다.")

        safe_key = re.sub(r"[^0-9A-Za-z가-힣_]+", "_", str(cfg.get("_key") or cfg.get("label") or "dataset"))
        temp_root = ensure_dir(Path(output_dir) / "_tmp_sido_collect" / safe_key)
        try:
            shutil.rmtree(temp_root, ignore_errors=True)
        except Exception:
            pass
        temp_root.mkdir(parents=True, exist_ok=True)

        saved_paths = []
        failed = []
        loop_rows = rows
        if self._sido_total_single_fetch_required(cfg):
            loop_rows = rows[:1]
            self.log(f"[시도수집1회] {cfg.get('label')} | 시도단위/전국단위 발표자료라 시군구 반복 수집 생략")

        self.log(f"[시도수집] {cfg.get('label')} | {sido} | 대상 시군구 {len(loop_rows):,}개")
        for row in loop_rows:
            if self.stop_event.is_set():
                break
            try:
                self._set_thread_region_context(row["시도"], row["시군구"], row["REG"])
                paths = self._run_one_cfg(client, temp_root, cfg, end_period)
                saved_paths.extend(Path(p) for p in paths)
                self.log(f"[시도수집완료] {cfg.get('label')} | {row['시군구']} | 파일 {len(paths)}개")
            except Exception as e:
                failed.append((row["시군구"], str(e)))
                self.log(f"[시도수집실패] {cfg.get('label')} | {row['시군구']} | {self._sanitize_log_text(str(e))}")
            finally:
                self._clear_thread_region_context()

        data_frames = []
        for path in saved_paths:
            try:
                df = self._read_saved_stat_data_sheet(path)
                if not df.empty:
                    data_frames.append(df)
            except Exception as e:
                self.log(f"[시도파일읽기실패] {path.name} | {e}")

        if not data_frames:
            if failed:
                preview = "; ".join([f"{name}: {err}" for name, err in failed[:3]])
                raise RuntimeError(f"시도 전체 수집 결과가 비어 있습니다. 일부 실패: {preview}")
            return []

        combined_df = pd.concat(data_frames, ignore_index=True)
        filename = self._sido_total_filename_from_saved(saved_paths[0], sido)
        final_paths = []
        for group_folder_name in self._get_group_folders_for_cfg(cfg):
            folder = ensure_dir(Path(output_dir) / group_folder_name)
            path = self._save_single_stat_excel(folder, filename, combined_df, cfg["label"])
            final_paths.append(path)
        try:
            shutil.rmtree(temp_root, ignore_errors=True)
        except Exception:
            pass
        if failed:
            self.log(f"[시도부분실패] {cfg.get('label')} | 실패 {len(failed)}개 시군구, 성공 파일 {len(saved_paths)}개")
        return final_paths

    def _run_one_cfg_officetel_sido_total(self, client: KosisClient, output_dir: Path, cfg: dict, end_period: str):
        return self._run_one_cfg_sido_total(client, output_dir, cfg, end_period)


    def _guess_normalized_column(self, original_col: str):
        mapping = {
            "PRD_DE": ("period", "string", "key", "기준시점"),
            "DT": ("value", "number", "metric", "통계값"),
            "DAT": ("value", "number", "metric", "통계값"),
            "ITM_ID": ("item_id", "string", "dimension", "항목코드"),
            "ITM_NM": ("item_name", "string", "dimension", "항목명"),
            "C1": ("region_level_1_code", "string", "dimension", "분류1코드"),
            "C1_NM": ("region_level_1", "string", "dimension", "분류1명"),
            "C2": ("region_level_2_code", "string", "dimension", "분류2코드"),
            "C2_NM": ("region_level_2", "string", "dimension", "분류2명"),
            "C3": ("region_level_3_code", "string", "dimension", "분류3코드"),
            "C3_NM": ("region_level_3", "string", "dimension", "분류3명"),
            "C4": ("region_level_4_code", "string", "dimension", "분류4코드"),
            "C4_NM": ("region_level_4", "string", "dimension", "분류4명"),
            "OBJ_L1": ("obj_level_1_code", "string", "dimension", "분류1코드"),
            "OBJ_L1_NM": ("obj_level_1", "string", "dimension", "분류1명"),
            "OBJ_L2": ("obj_level_2_code", "string", "dimension", "분류2코드"),
            "OBJ_L2_NM": ("obj_level_2", "string", "dimension", "분류2명"),
            "OBJ_L3": ("obj_level_3_code", "string", "dimension", "분류3코드"),
            "OBJ_L3_NM": ("obj_level_3", "string", "dimension", "분류3명"),
            "OBJ_L4": ("obj_level_4_code", "string", "dimension", "분류4코드"),
            "OBJ_L4_NM": ("obj_level_4", "string", "dimension", "분류4명"),
            "TBL_NM": ("table_name", "string", "meta", "통계표명"),
            "ORG_NM": ("org_name", "string", "meta", "기관명"),
            "UNIT_NM": ("unit_name", "string", "meta", "단위"),
            "기간": ("period", "string", "key", "기준시점"),
            "값": ("value", "number", "metric", "통계값"),
            "단위": ("unit_name", "string", "meta", "단위"),
            "항목코드": ("item_id", "string", "dimension", "항목코드"),
            "항목명": ("item_name", "string", "dimension", "항목명"),
            "분류1코드": ("region_level_1_code", "string", "dimension", "분류1코드"),
            "분류1": ("region_level_1", "string", "dimension", "분류1명"),
            "분류2코드": ("region_level_2_code", "string", "dimension", "분류2코드"),
            "분류2": ("region_level_2", "string", "dimension", "분류2명"),
            "분류3코드": ("region_level_3_code", "string", "dimension", "분류3코드"),
            "분류3": ("region_level_3", "string", "dimension", "분류3명"),
            "분류4코드": ("region_level_4_code", "string", "dimension", "분류4코드"),
            "분류4": ("region_level_4", "string", "dimension", "분류4명"),
            "통계표명": ("table_name", "string", "meta", "통계표명"),
            "기관명": ("org_name", "string", "meta", "기관명"),
            "기준월": ("base_month", "string", "key", "기준월"),
            "출처": ("source_name", "string", "meta", "출처"),
            "기준금리": ("base_rate", "number", "metric", "기준금리"),
            "통계표코드": ("stat_code", "string", "meta", "통계표코드"),
            "통계명": ("stat_name", "string", "meta", "통계명"),
            "항목코드": ("item_code", "string", "dimension", "항목코드"),
            "항목명": ("item_name", "string", "dimension", "항목명"),
            "은행명": ("bank_name", "string", "dimension", "은행명"),
            "상품구분": ("product_type", "string", "dimension", "상품구분"),
            "최저금리": ("min_rate", "number", "metric", "최저금리"),
            "최고금리": ("max_rate", "number", "metric", "최고금리"),
            "신규취급액COFIX": ("cofix_new", "number", "metric", "신규취급액COFIX"),
            "잔액COFIX": ("cofix_balance", "number", "metric", "잔액COFIX"),
            "신잔액COFIX": ("cofix_new_balance", "number", "metric", "신잔액COFIX"),
            "비고": ("note", "string", "meta", "비고"),
        }
        if original_col in mapping:
            return mapping[original_col]
        col = str(original_col).strip()
        col_lower = col.lower()
        if "date" in col_lower or "period" in col_lower:
            return ("period", "string", "key", "기준시점")
        if "rate" in col_lower:
            return (col_lower, "number", "metric", "금리지표")
        if "name" in col_lower or col_lower.endswith("_nm"):
            return (col_lower.replace("_nm", "_name"), "string", "dimension", "명칭")
        if "id" in col_lower or col_lower.endswith("_cd"):
            return (col_lower.replace("_cd", "_code"), "string", "dimension", "코드")
        safe = re.sub(r"[^0-9a-zA-Z_]+", "_", col_lower).strip("_")
        return (safe or "column", "string", "dimension", "원본컬럼")

    def _build_schema_dataframe(self, df: pd.DataFrame, source_file: str, dataset_name: str):
        rows = []
        if df is None:
            df = pd.DataFrame()
        for col in df.columns:
            normalized, dtype, role, meaning = self._guess_normalized_column(col)
            example = ""
            if col in df.columns and len(df) > 0:
                try:
                    non_null = df[col].dropna()
                    if len(non_null) > 0:
                        example = str(non_null.iloc[0])
                except Exception:
                    example = ""
            rows.append({
                "source_file": source_file,
                "dataset_name": dataset_name,
                "sheet_name": "data",
                "original_column": col,
                "normalized_column": normalized,
                "data_type": dtype,
                "role": role,
                "meaning": meaning,
                "example": example,
                "use_for_analysis": "Y",
            })
        return pd.DataFrame(rows)


    def _save_ai_upload_file(self, folder: Path, group_folder_name: str, data_filename: str, ai_df: pd.DataFrame):
        ai_upload_root = ensure_dir(folder.parent / "AI_UPLOAD" / group_folder_name)
        ai_xlsx_path = ai_upload_root / data_filename
        self._remove_existing_file_if_exists(ai_xlsx_path)
        with pd.ExcelWriter(ai_xlsx_path, engine="openpyxl") as writer:
            ai_df.to_excel(writer, sheet_name="data", index=False)
        self.log(f"[AI업로드저장] {ai_xlsx_path}")

    def _save_schema_files(self, folder: Path, data_filename: str, dataset_name: str, df: pd.DataFrame):
        schema_df = self._build_schema_dataframe(df, data_filename, dataset_name)
        schema_xlsx = folder / data_filename.replace(".xlsx", "_schema.xlsx")
        schema_json = folder / data_filename.replace(".xlsx", "_schema.json")
        self._remove_existing_file_if_exists(schema_xlsx)
        self._remove_existing_file_if_exists(schema_json)
        with pd.ExcelWriter(schema_xlsx, engine="openpyxl") as writer:
            schema_df.to_excel(writer, sheet_name="schema", index=False)
        schema_payload = {
            "source_file": data_filename,
            "dataset_name": dataset_name,
            "columns": schema_df.to_dict(orient="records"),
        }
        schema_json.write_text(json.dumps(schema_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.log(f"[스키마저장] {schema_xlsx}")
        self.log(f"[스키마저장] {schema_json}")


    def _guess_korean_column_name(self, original_col: str) -> str:
        """엑셀 data 시트의 컬럼명을 업무자가 바로 읽을 수 있는 한글명으로 변환합니다."""
        col = str(original_col).strip()
        mapping = {
            # KOSIS 공통 원본 컬럼
            "PRD_DE": "기간",
            "TBL_NM": "통계표명",
            "ORG_NM": "기관명",
            "DT": "값",
            "DAT": "값",
            "UNIT_NM": "단위",
            "ITM_ID": "항목코드",
            "ITM_NM": "항목명",
            "C1": "분류1코드",
            "C1_NM": "분류1명",
            "C2": "분류2코드",
            "C2_NM": "분류2명",
            "C3": "분류3코드",
            "C3_NM": "분류3명",
            "C4": "분류4코드",
            "C4_NM": "분류4명",
            "OBJ_L1": "분류1코드",
            "OBJ_L1_NM": "분류1명",
            "OBJ_L2": "분류2코드",
            "OBJ_L2_NM": "분류2명",
            "OBJ_L3": "분류3코드",
            "OBJ_L3_NM": "분류3명",
            "OBJ_L4": "분류4코드",
            "OBJ_L4_NM": "분류4명",
            "OBJ_L": "분류코드",
            "OBJ_L_NM": "분류명",
            "C1_OBJ_NM": "분류1객체명",
            "C2_OBJ_NM": "분류2객체명",

            # 기존 내부 표준 컬럼
            "period": "기간",
            "value": "값",
            "item_id": "항목코드",
            "item_name": "항목명",
            "region_level_1_code": "분류1코드",
            "region_level_1": "분류1명",
            "region_level_2_code": "분류2코드",
            "region_level_2": "분류2명",
            "region_level_3_code": "분류3코드",
            "region_level_3": "분류3명",
            "region_level_4_code": "분류4코드",
            "region_level_4": "분류4명",
            "obj_level_1_code": "분류1코드",
            "obj_level_1": "분류1명",
            "obj_level_2_code": "분류2코드",
            "obj_level_2": "분류2명",
            "obj_level_3_code": "분류3코드",
            "obj_level_3": "분류3명",
            "obj_level_4_code": "분류4코드",
            "obj_level_4": "분류4명",
            "table_name": "통계표명",
            "org_name": "기관명",
            "unit_name": "단위",
            "base_month": "기준월",
            "source_name": "출처",
            "stat_code": "통계표코드",
            "stat_name": "통계명",

            # 이미 한글화된 컬럼
            "기간": "기간",
            "값": "값",
            "단위": "단위",
            "항목코드": "항목코드",
            "항목명": "항목명",
            "분류1코드": "분류1코드",
            "분류1": "분류1명",
            "분류1명": "분류1명",
            "분류2코드": "분류2코드",
            "분류2": "분류2명",
            "분류2명": "분류2명",
            "분류3코드": "분류3코드",
            "분류3": "분류3명",
            "분류3명": "분류3명",
            "분류4코드": "분류4코드",
            "분류4": "분류4명",
            "분류4명": "분류4명",
            "통계표명": "통계표명",
            "기관명": "기관명",
            "기준월": "기준월",
            "출처": "출처",
            "통계표코드": "통계표코드",
            "통계명": "통계명",

            # 국토교통부 실거래 공통 컬럼
            "조회월": "조회월",
            "LAWD_CD": "시군구코드",
            "DEAL_YMD": "계약년월",
            "sggCd": "시군구코드",
            "umdNm": "읍면동명",
            "jibun": "지번",
            "aptNm": "아파트명",
            "aptDong": "아파트동",
            "offiNm": "오피스텔명",
            "mhouseNm": "연립다세대명",
            "houseType": "주택유형",
            "dealYear": "계약연도",
            "dealMonth": "계약월",
            "dealDay": "계약일",
            "dealAmount": "거래금액",
            "deposit": "보증금",
            "monthlyRent": "월세",
            "rentMoney": "월세",
            "excluUseAr": "전용면적",
            "totalFloorAr": "연면적",
            "landAr": "토지면적",
            "dealArea": "거래면적",
            "buildYear": "건축년도",
            "floor": "층",
            "rightType": "권리구분",
            "dealType": "거래유형",
            "dealingGbn": "거래유형",
            "estateAgentSggNm": "중개사소재지",
            "cancelDealType": "해제여부",
            "cancelDealDay": "해제사유발생일",
            "reqGbn": "신고구분",
            "rdealerLawdnm": "신고관청",
            "contractTerm": "계약기간",
            "contractType": "계약구분",
            "useRRRight": "갱신요구권사용",
            "preDeposit": "종전보증금",
            "preMonthlyRent": "종전월세",
            "previousDeposit": "종전보증금",
            "previousMonthlyRent": "종전월세",
            "buyerGbn": "매수자구분",
            "sellerGbn": "매도자구분",
            "landUse": "용도지역",
            "dealObject": "거래대상",
            "buildingUse": "건물용도",
            "useArea": "용도지역",
            "induty": "업종",
            "factoryArea": "공장면적",
        }
        if col in mapping:
            return mapping[col]

        lower = col.lower()
        if "date" in lower or "period" in lower or "ymd" in lower:
            return "기간"
        if "amount" in lower or "price" in lower or "money" in lower:
            return "금액"
        if "rent" in lower:
            return "임대료"
        if "area" in lower or lower.endswith("ar"):
            return "면적"
        if "code" in lower or lower.endswith("cd"):
            return "코드"
        if "name" in lower or lower.endswith("nm"):
            return "명칭"
        return col

    def _standardize_ai_columns(self, df: pd.DataFrame):
        """엑셀 data 시트 컬럼을 한글 항목명으로 저장합니다."""
        if df is None:
            return pd.DataFrame(), {}
        df2 = df.copy()
        rename_map = {}
        used = set()

        for col in df2.columns:
            base = self._guess_korean_column_name(col)
            name = base
            i = 2
            while name in used:
                name = f"{base}_{i}"
                i += 1
            used.add(name)
            rename_map[col] = name

        df2 = df2.rename(columns=rename_map)
        return df2, rename_map

    def _normalize_df(self, df: pd.DataFrame):
        if df.empty:
            return df
        rename = {k: v for k, v in COMMON_RENAME.items() if k in df.columns}
        df = df.rename(columns=rename)
        if "값" in df.columns:
            df["값"] = pd.to_numeric(df["값"], errors="coerce")
        sort_cols = [c for c in ["기간", "분류1", "분류2", "분류3", "항목명"] if c in df.columns]
        if sort_cols:
            df = df.sort_values(sort_cols).reset_index(drop=True)
        return df

    def _normalize_kab_region_name(self, value: str) -> str:
        """KOSIS 가격지수 지역명과 법정동 콤보 지역명을 비교하기 위한 간단 정규화."""
        text = str(value or "").strip().replace(" ", "")
        for suffix in ["특별자치시", "특별자치도", "광역시", "특별시", "자치시", "자치도", "시", "군", "구"]:
            if text.endswith(suffix) and len(text) > len(suffix):
                text = text[: -len(suffix)]
                break
        return text

    def _kab_region_name_candidates(self, sigungu_name: str) -> set[str]:
        """선택 시군구명에서 KOSIS 가격지수 지역명 후보를 만듭니다.

        예) 남구 -> 남, 나주시 -> 나주, 청주시 흥덕구 -> {청주시흥덕, 흥덕}
        """
        raw = str(sigungu_name or "").strip()
        compact = raw.replace(" ", "")
        values = {raw, compact, self._normalize_kab_region_name(raw)}
        if " " in raw:
            last = raw.split()[-1]
            values.add(last)
            values.add(self._normalize_kab_region_name(last))
        m = re.search(r"([가-힣]+시)([가-힣]+구)$", compact)
        if m:
            values.add(m.group(2))
            values.add(self._normalize_kab_region_name(m.group(2)))
        return {v for v in values if v}

    def _filter_kab_price_index_selected_sido_df(self, df: pd.DataFrame, cfg: dict):
        """전국주택가격동향조사 유형별 가격지수에서 선택 시도 + 선택 시군구만 후처리 추출합니다."""
        if df is None or df.empty:
            return df

        ctx_sido_name, selected_sigungu_name, ctx_reg = self._current_region_values()
        selected_sido = self._get_kosis_region_code(2) if ctx_reg else ""
        sido_total_mode = self._sido_report_enabled(self._get_current_group_name())
        if not selected_sido:
            return df

        sido_code = KAB_PRICE_INDEX_SIDO_PREFIX_MAP.get(selected_sido)
        if not sido_code:
            self.log(f"[가격지수후처리미지원] {cfg['label']} | 법정동 시도코드={selected_sido} | KAB_PRICE_INDEX_SIDO_PREFIX_MAP에 매핑 필요")
            return df

        code_cols = [c for c in ["C2", "OBJ_L2", "분류2코드"] if c in df.columns]
        name_cols = [c for c in ["C2_NM", "OBJ_L2_NM", "분류2", "분류2명"] if c in df.columns]
        if not code_cols:
            self.log(f"[가격지수후처리없음] {cfg['label']} | 지역코드 컬럼 없음")
            return df

        # 1) 해당 시도는 내부 시도코드와 정확히 일치하는 행만 유지
        sido_mask = None
        matched_col = None
        for col in code_cols:
            cur = df[col].astype(str).str.strip().eq(sido_code)
            if cur.any():
                sido_mask = cur if sido_mask is None else (sido_mask | cur)
                matched_col = matched_col or col

        # 2) 해당 시군구는 같은 시도 코드 계열 안에서 지역명 후보가 일치하는 행만 유지
        sigungu_mask = None
        candidates = self._kab_region_name_candidates(selected_sigungu_name)
        if selected_sigungu_name and name_cols and not sido_total_mode:
            for code_col in code_cols:
                code_series = df[code_col].astype(str).str.strip()
                in_selected_sido = code_series.str.startswith(sido_code) & ~code_series.eq(sido_code)
                for name_col in name_cols:
                    normalized_names = df[name_col].astype(str).map(self._normalize_kab_region_name)
                    cur = in_selected_sido & normalized_names.isin(candidates)
                    if cur.any():
                        sigungu_mask = cur if sigungu_mask is None else (sigungu_mask | cur)
                        matched_col = matched_col or code_col

        mask = sido_mask
        if sigungu_mask is not None and sigungu_mask.any():
            mask = sigungu_mask if mask is None else (mask | sigungu_mask)

        if mask is None or not mask.any():
            self.log(f"[가격지수후처리없음] {cfg['label']} | 시도코드={sido_code} | 시군구={selected_sigungu_name} | 코드컬럼={code_cols}")
            return df

        before = len(df)
        out = df[mask].copy().reset_index(drop=True)
        region_label = ctx_sido_name if sido_total_mode else (f"{ctx_sido_name} + {selected_sigungu_name}" if selected_sigungu_name else ctx_sido_name)
        if selected_sigungu_name and not sido_total_mode and (sigungu_mask is None or not sigungu_mask.any()):
            self.log(f"[가격지수시군구주의] {cfg['label']} | 시도는 추출했으나 시군구명 후보 {sorted(candidates)} 일치 행을 찾지 못했습니다.")
        self.log(f"[가격지수후처리] {cfg['label']} | {region_label} | {before:,}->{len(out):,}건 | 기준컬럼={matched_col}")
        return out

    def _filter_housing_sido_df(self, df: pd.DataFrame, cfg: dict):
        if df is None or df.empty or not cfg.get("post_filter_region"):
            return df
        if cfg.get("kab_price_region_filter"):
            return self._filter_kab_price_index_selected_sido_df(df, cfg)
        selected_sido = self._get_kosis_region_code(2) if self._current_region_values()[2] else ""
        if not selected_sido:
            return df

        code_cols = [c for c in ["C1", "C2", "C3", "C4", "OBJ_L1", "OBJ_L2", "OBJ_L3", "OBJ_L4"] if c in df.columns]
        if not code_cols:
            return df

        def matches(val):
            s = str(val).strip()
            return bool(s) and s.startswith(selected_sido)

        mask = None
        matched_col = None
        for col in code_cols:
            cur = df[col].astype(str).map(matches)
            if cur.any():
                mask = cur if mask is None else (mask | cur)
                matched_col = matched_col or col

        if mask is None or not mask.any():
            self.log(f"[시도후처리없음] {cfg['label']} | 시도코드={selected_sido} | 코드컬럼={code_cols}")
            return df

        before = len(df)
        out = df[mask].copy().reset_index(drop=True)
        self.log(f"[시도후처리] {cfg['label']} | 시도코드={selected_sido} | {before:,}->{len(out):,}건 | 기준컬럼={matched_col}")
        return out

    def _fetch_housing_meta_plan(self, client: KosisClient, cfg: dict):
        cached = cfg.get("_meta_plan")
        if cached:
            return cached

        rows = client.get_meta_itm(cfg["orgId"], cfg["tblId"])
        if not rows:
            raise RuntimeError(f"메타 조회 결과가 비어 있습니다: {cfg['label']}")

        groups = {}
        item_pool = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            seq_raw = row.get("objIdSn") or row.get("OBJ_ID_SN") or row.get("objIdSeq") or row.get("OBJ_ID_SEQ")
            obj_id = str(row.get("objId") or row.get("OBJ_ID") or "").strip()
            obj_nm = str(row.get("objNm") or row.get("OBJ_NM") or "").strip()
            code = str(row.get("itmId") or row.get("ITM_ID") or "").strip()
            code_nm = str(row.get("itmNm") or row.get("ITM_NM") or "").strip()
            up_code = str(row.get("upItmId") or row.get("UP_ITM_ID") or "").strip()
            if not code:
                continue
            rec = {"code": code, "name": code_nm, "up": up_code}
            seq = None
            try:
                if str(seq_raw).strip() != "":
                    seq = int(str(seq_raw).strip())
            except Exception:
                seq = None
            if seq is None:
                item_pool.append(rec)
                continue
            groups.setdefault(seq, {"seq": seq, "obj_id": obj_id, "obj_nm": obj_nm, "codes": []})
            if code not in {x["code"] for x in groups[seq]["codes"]}:
                groups[seq]["codes"].append(rec)

        dimensions = [groups[k] for k in sorted(groups)]
        if not dimensions:
            raise RuntimeError(f"메타에서 분류축을 찾지 못했습니다: {cfg['label']}")

        selected_region = self._get_kosis_region_code(int(cfg.get("region_code_length", 5))) if self._current_region_values()[2] else ""
        region_len = len(selected_region or "")

        region_seq = None
        if selected_region:
            for dim in dimensions:
                codes = [c["code"] for c in dim["codes"]]
                if selected_region in codes:
                    region_seq = dim["seq"]
                    break
        if region_seq is None:
            for dim in dimensions:
                nm = (dim.get("obj_nm") or "") + " " + (dim.get("obj_id") or "")
                nm = nm.lower()
                if region_len == 5 and any(tok in nm for tok in ["시군구", "행정구역", "지역"]):
                    region_seq = dim["seq"]
                    break
                if region_len == 2 and any(tok in nm for tok in ["시도", "시·도", "지역", "행정구역"]):
                    region_seq = dim["seq"]
                    break

        def choose_default_code(dim):
            codes = dim.get("codes", [])
            if not codes:
                return "ALL"
            names = [c.get("name", "") for c in codes]
            values = [c.get("code", "") for c in codes]
            for preferred in ["ALL", "00", "0", "합계", "총계", "계"]:
                for c in codes:
                    if c["code"] == preferred or c["name"] == preferred:
                        return c["code"]
            leaf_codes = [c for c in codes if not c.get("up") or c.get("up") in {"0", "ROOT", "-"}]
            pool = leaf_codes or codes
            return pool[0]["code"]

        item_candidates = []
        seen_items = set()
        for pool in [item_pool]:
            for rec in pool:
                code = rec.get("code", "")
                if code and code not in seen_items:
                    item_candidates.append(code)
                    seen_items.add(code)
        for fallback in ["ALL", "T001", "T01", "1", "0"]:
            if fallback not in seen_items:
                item_candidates.append(fallback)
                seen_items.add(fallback)

        plan = {
            "dimensions": dimensions,
            "region_seq": region_seq,
            "selected_region": selected_region,
            "item_candidates": item_candidates,
        }
        cfg["_meta_plan"] = plan
        dim_desc = ", ".join([f"objL{d['seq']}={d.get('obj_nm') or d.get('obj_id') or '미상'}({len(d.get('codes', []))})" for d in dimensions])
        self.log(f"[메타조회] {cfg['label']} | 축={dim_desc} | 지역축={('objL'+str(region_seq)) if region_seq else '미확정'}")
        return plan

    def _build_housing_meta_candidates(self, client: KosisClient, cfg: dict, ym: str):
        plan = self._fetch_housing_meta_plan(client, cfg)
        base = {
            "orgId": cfg["orgId"],
            "tblId": cfg["tblId"],
            "prdSe": cfg.get("prdSe", "M").upper(),
            "startPrdDe": ym,
            "endPrdDe": ym,
        }
        dims = plan["dimensions"]
        region_seq = plan.get("region_seq")
        selected_region = plan.get("selected_region") or ""
        region_len = len(selected_region)

        def choose_default_code(dim):
            codes = dim.get("codes", [])
            for preferred in ["ALL", "00", "0"]:
                for c in codes:
                    if c["code"] == preferred or c.get("name") == preferred:
                        return c["code"]
            for nm in ["계", "총계", "합계"]:
                for c in codes:
                    if c.get("name") == nm:
                        return c["code"]
            return codes[0]["code"] if codes else "ALL"

        def candidate_region_codes(dim):
            codes = [c["code"] for c in dim.get("codes", [])]
            out = []
            if selected_region:
                if selected_region in codes:
                    out.append(selected_region)
                if region_len == 5:
                    for code in [selected_region[:2], selected_region[:5]]:
                        if code in codes and code not in out:
                            out.append(code)
                if region_len == 2:
                    if selected_region[:2] in codes and selected_region[:2] not in out:
                        out.append(selected_region[:2])
                for code in codes:
                    if code.startswith(selected_region) and code not in out:
                        out.append(code)
            for c in codes[:5]:
                if c not in out:
                    out.append(c)
            return out[:6]

        base_payload = {}
        for dim in dims:
            seq = dim["seq"]
            base_payload[f"objL{seq}"] = choose_default_code(dim)

        candidates = []
        def add_candidate(name, payload):
            merged = {**base, **payload}
            if merged not in [c[1] for c in candidates]:
                candidates.append((name, merged))

        region_variants = [base_payload.copy()]
        if region_seq is not None:
            dim = next(d for d in dims if d["seq"] == region_seq)
            for code in candidate_region_codes(dim):
                payload = base_payload.copy()
                payload[f"objL{region_seq}"] = code
                region_variants.append(payload)

        for idx, payload in enumerate(region_variants, 1):
            add_candidate(f"메타기본-{idx}", payload)
            for item in plan.get("item_candidates", [])[:8]:
                p2 = payload.copy()
                p2["itmId"] = item
                add_candidate(f"메타기본-{idx},itmId={item}", p2)

        return candidates

    def _build_housing_probe_candidates(self, cfg: dict, ym: str):
        base = {
            "orgId": cfg["orgId"],
            "tblId": cfg["tblId"],
            "prdSe": cfg.get("prdSe", "M").upper(),
            "startPrdDe": ym,
            "endPrdDe": ym,
        }

        preferred_length = int(cfg.get("region_code_length", 5))
        reg5 = self._get_kosis_region_code(5)
        reg2 = self._get_kosis_region_code(2)
        preferred_reg = self._get_kosis_region_code(preferred_length) if self._current_region_values()[2] else ""

        region_candidates = []
        if preferred_reg:
            region_candidates.append(preferred_reg)
        if preferred_length == 5:
            region_candidates.extend([reg2, reg5[:2] if reg5 else ""])
        else:
            region_candidates.extend([reg5, reg2])
        region_candidates = [x for i, x in enumerate(region_candidates) if x and x not in region_candidates[:i]]

        candidates = []
        if not region_candidates:
            region_candidates = [None]

        def add_candidate(name, extra):
            payload = {**base, **extra}
            if payload not in [c[1] for c in candidates]:
                candidates.append((name, payload))

        probe_mode = cfg.get("auto_probe_mode", "generic")

        if probe_mode == "permit_sigungu":
            for reg in region_candidates:
                if reg is None:
                    continue
                add_candidate(f"objL1={reg}", {"objL1": reg})
                add_candidate(f"objL2={reg}", {"objL2": reg})
                add_candidate(f"objL1={reg},itmId=ALL", {"objL1": reg, "itmId": "ALL"})
                add_candidate(f"objL2={reg},itmId=ALL", {"objL2": reg, "itmId": "ALL"})
                add_candidate(f"objL1={reg},objL2=ALL", {"objL1": reg, "objL2": "ALL"})
                add_candidate(f"objL1=ALL,objL2={reg}", {"objL1": "ALL", "objL2": reg})
                add_candidate(f"objL1={reg},objL2=ALL,itmId=ALL", {"objL1": reg, "objL2": "ALL", "itmId": "ALL"})
                add_candidate(f"objL1=ALL,objL2={reg},itmId=ALL", {"objL1": "ALL", "objL2": reg, "itmId": "ALL"})
                add_candidate(f"objL1={reg},objL2=ALL,itmId=T001", {"objL1": reg, "objL2": "ALL", "itmId": "T001"})
                add_candidate(f"objL1=ALL,objL2={reg},itmId=T001", {"objL1": "ALL", "objL2": reg, "itmId": "T001"})
            return candidates

        if probe_mode == "sido_monthly_type":
            item_candidates = ["ALL", "T001", "T100", "0", "1"]
            for reg in region_candidates:
                if reg is None:
                    continue
                for item in item_candidates:
                    add_candidate(f"objL1={reg},itmId={item}", {"objL1": reg, "itmId": item})
                    add_candidate(f"objL2={reg},itmId={item}", {"objL2": reg, "itmId": item})
                    add_candidate(f"objL1={reg},objL2=ALL,itmId={item}", {"objL1": reg, "objL2": "ALL", "itmId": item})
                    add_candidate(f"objL1=ALL,objL2={reg},itmId={item}", {"objL1": "ALL", "objL2": reg, "itmId": item})
            return candidates

        base_variants = [
            {},
            {"itmId": "ALL"},
            {"objL1": "ALL"},
            {"objL2": "ALL"},
            {"objL1": "ALL", "itmId": "ALL"},
            {"objL2": "ALL", "itmId": "ALL"},
            {"objL1": "ALL", "objL2": "ALL"},
            {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"},
        ]

        for reg in region_candidates:
            if reg is None:
                for idx, extra in enumerate(base_variants, 1):
                    add_candidate(f"무지역-{idx}", extra)
                continue
            add_candidate(f"objL1={reg}", {"objL1": reg})
            add_candidate(f"objL2={reg}", {"objL2": reg})
            add_candidate(f"objL1={reg},itmId=ALL", {"objL1": reg, "itmId": "ALL"})
            add_candidate(f"objL2={reg},itmId=ALL", {"objL2": reg, "itmId": "ALL"})
            add_candidate(f"objL1={reg},objL2=ALL", {"objL1": reg, "objL2": "ALL"})
            add_candidate(f"objL1=ALL,objL2={reg}", {"objL1": "ALL", "objL2": reg})
            add_candidate(f"objL1={reg},objL2=ALL,itmId=ALL", {"objL1": reg, "objL2": "ALL", "itmId": "ALL"})
            add_candidate(f"objL1=ALL,objL2={reg},itmId=ALL", {"objL1": "ALL", "objL2": reg, "itmId": "ALL"})
        return candidates

    def _request_housing_month_df(self, client: KosisClient, cfg: dict, ym: str):
        if cfg.get("auto_probe_mode") == "all_sido_postfilter":
            candidates = []
            meta_error = None
            try:
                plan = self._fetch_housing_meta_plan(client, cfg)
                base = {
                    "orgId": cfg["orgId"],
                    "tblId": cfg["tblId"],
                    "prdSe": cfg.get("prdSe", "M").upper(),
                    "startPrdDe": ym,
                    "endPrdDe": ym,
                }
                dims = plan.get("dimensions", [])

                def choose_default_code(dim):
                    codes = dim.get("codes", [])
                    for preferred in ["ALL", "00", "0"]:
                        for c in codes:
                            if c.get("code") == preferred or c.get("name") == preferred:
                                return c.get("code")
                    for nm in ["계", "총계", "합계"]:
                        for c in codes:
                            if c.get("name") == nm:
                                return c.get("code")
                    return codes[0].get("code") if codes else "ALL"

                base_payload = {}
                for dim in dims:
                    seq = dim.get("seq")
                    if seq is None:
                        continue
                    base_payload[f"objL{seq}"] = choose_default_code(dim)

                item_candidates = []
                for item in plan.get("item_candidates", []):
                    if item and item not in item_candidates:
                        item_candidates.append(item)
                for fallback in ["ALL", "T001", "T100", "0", "1"]:
                    if fallback not in item_candidates:
                        item_candidates.append(fallback)

                variants = []
                variants.append(("메타기본", base_payload.copy()))
                for item in item_candidates[:10]:
                    payload = base_payload.copy()
                    payload["itmId"] = item
                    variants.append((f"메타+itmId={item}", payload))

                for name, payload in variants:
                    merged = {**base, **payload}
                    if merged not in [c[1] for c in candidates]:
                        candidates.append((name, merged))
            except Exception as e:
                meta_error = e
                self.log(f"[메타기반대기실패] {cfg['label']} | {ym} | 사유={e}")

            fallback_base = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
            }
            fallback_variants = [
                ("fallback-무추가", {}),
                ("fallback-itmId=ALL", {"itmId": "ALL"}),
                ("fallback-objL1=ALL", {"objL1": "ALL"}),
                ("fallback-objL1=ALL,itmId=ALL", {"objL1": "ALL", "itmId": "ALL"}),
                ("fallback-objL1=ALL,objL2=ALL", {"objL1": "ALL", "objL2": "ALL"}),
                ("fallback-objL1=ALL,objL2=ALL,itmId=ALL", {"objL1": "ALL", "objL2": "ALL", "itmId": "ALL"}),
            ]
            for name, payload in fallback_variants:
                merged = {**fallback_base, **payload}
                if merged not in [c[1] for c in candidates]:
                    candidates.append((name, merged))

            last_error = meta_error
            for name, params in candidates:
                try:
                    df = client.request(params)
                    cfg["_resolved_probe_params"] = {k: v for k, v in params.items() if k.startswith("objL") or k == "itmId"}
                    self.log(f"[시도전체조회확정] {cfg['label']} | {ym} | 전략={name} | 파라미터={cfg['_resolved_probe_params']}")
                    return df
                except EmptyKosisDataError:
                    cfg["_resolved_probe_params"] = {k: v for k, v in params.items() if k.startswith("objL") or k == "itmId"}
                    self.log(f"[시도전체조회확정-데이터없음] {cfg['label']} | {ym} | 전략={name} | 파라미터={cfg['_resolved_probe_params']}")
                    raise
                except InvalidKosisParamsError as e:
                    last_error = e
                    self.log(f"[시도전체조회실패] {cfg['label']} | {ym} | 전략={name} | 사유={e}")
                    continue
            raise RuntimeError(f"시도 전체 조회 파라미터 자동 탐색 실패: {cfg['label']} | {ym} | 마지막오류={last_error}")

        resolved = cfg.get("_resolved_probe_params")
        if resolved:
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                **resolved,
            }
            return client.request(params)

        candidates = []
        meta_error = None
        try:
            candidates.extend(self._build_housing_meta_candidates(client, cfg, ym))
        except Exception as e:
            meta_error = e
            self.log(f"[메타기반대기실패] {cfg['label']} | {ym} | 사유={e}")
        candidates.extend(self._build_housing_probe_candidates(cfg, ym))
        invalid_count = 0
        last_error = meta_error
        for idx, (name, params) in enumerate(candidates, 1):
            try:
                df = client.request(params)
                cfg["_resolved_probe_params"] = {k: v for k, v in params.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[차원확정] {cfg['label']} | {ym} | 전략={name} | 파라미터={cfg['_resolved_probe_params']}")
                return df
            except EmptyKosisDataError:
                cfg["_resolved_probe_params"] = {k: v for k, v in params.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[차원확정-데이터없음] {cfg['label']} | {ym} | 전략={name} | 파라미터={cfg['_resolved_probe_params']}")
                raise
            except InvalidKosisParamsError as e:
                invalid_count += 1
                last_error = e
                self.log(f"[차원탐색실패] {cfg['label']} | {ym} | 전략={idx}/{len(candidates)} {name} | 사유={e}")
                continue
        raise RuntimeError(f"표 차원구조 자동 탐색 실패: {cfg['label']} | {ym} | 시도 {len(candidates)}건 | 마지막오류={last_error}")

    def _request_household_sigungu_month_df(self, client: KosisClient, cfg: dict, ym: str):
        region_code = str(cfg.get("params", {}).get("objL1", "")).strip()
        if not region_code or region_code == "ALL":
            raise ValueError("시군구별_주민등록세대수_최근15개월은 시군구 선택이 필요합니다.")

        # 행정구역 개편(예: 2026-07 통합)으로 월별 지역코드가 달라질 수 있어 구코드·신코드를 순서대로 시도한다.
        region_codes = [region_code]
        for alt in self._kosis_region_code_variants(int(cfg.get("region_code_length", 5) or 5)):
            if alt and alt not in region_codes:
                region_codes.append(alt)
        if len(region_codes) > 1 and not cfg.get("_no_region_variants"):
            last_empty = None
            for alt_code in region_codes:
                alt_cfg = dict(cfg)
                alt_cfg["params"] = {**cfg.get("params", {}), "objL1": alt_code}
                alt_cfg["_no_region_variants"] = True  # 재귀 방지
                resolved = cfg.get(f"_resolved_household_params_{alt_code}")
                if resolved:
                    alt_cfg["_resolved_household_params"] = resolved
                try:
                    df = self._request_household_sigungu_month_df(client, alt_cfg, ym)
                except EmptyKosisDataError as e:
                    last_empty = e
                    continue
                if df is not None and not df.empty:
                    if alt_cfg.get("_resolved_household_params"):
                        cfg[f"_resolved_household_params_{alt_code}"] = alt_cfg["_resolved_household_params"]
                    if alt_code != region_code:
                        self.log(f"[세대수지역코드전환] {cfg['label']} | {ym} | 코드={alt_code}")
                    return df
            if last_empty:
                raise last_empty
            raise EmptyKosisDataError(f"0건 응답: {cfg['label']} | {ym}")

        resolved_override = cfg.get("_resolved_household_params")
        if resolved_override:
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                **resolved_override,
            }
            try:
                return client.request(params)
            except (InvalidKosisParamsError, EmptyKosisDataError):
                # 캐시된 파라미터가 특정 월에 더 이상 유효하지 않으면 재탐색
                cfg.pop("_resolved_household_params", None)
                cfg.pop("_resolved_probe_params", None)

        candidate_overrides = [
            {"objL1": region_code},
            {"objL1": region_code, "itmId": "T1"},
            {"objL1": region_code, "itmId": "T01"},
            {"objL1": region_code, "itmId": "1"},
            {"objL1": region_code, "objL2": "00"},
            {"objL1": region_code, "objL2": "0"},
            {"objL1": region_code, "objL2": "ALL"},
            {"objL1": region_code, "objL2": "00", "itmId": "T1"},
            {"objL1": region_code, "objL2": "0", "itmId": "T1"},
        ]

        last_error = None
        for idx, override in enumerate(candidate_overrides, 1):
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                **override,
            }
            try:
                df = client.request(params)
                cfg["_resolved_household_params"] = dict(override)
                cfg["_resolved_probe_params"] = {k: v for k, v in override.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[세대수차원확정] {cfg['label']} | {ym} | 전략={idx}/{len(candidate_overrides)} | 파라미터={cfg['_resolved_probe_params']}")
                return df
            except EmptyKosisDataError:
                cfg["_resolved_household_params"] = dict(override)
                cfg["_resolved_probe_params"] = {k: v for k, v in override.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[세대수차원확정-데이터없음] {cfg['label']} | {ym} | 전략={idx}/{len(candidate_overrides)} | 파라미터={cfg['_resolved_probe_params']}")
                raise
            except InvalidKosisParamsError as e:
                last_error = e
                self.log(f"[세대수차원탐색실패] {cfg['label']} | {ym} | 전략={idx}/{len(candidate_overrides)} | 사유={e}")
                continue

        raise RuntimeError(f"시군구별_주민등록세대수_최근15개월 파라미터 자동 탐색 실패: {cfg['label']} | {ym} | 마지막오류={last_error}")

    def _request_one_person_household_age_year_df(self, client: KosisClient, cfg: dict, year: str):
        region_code = str(cfg.get("params", {}).get("objL2", "") or cfg.get("params", {}).get("objL1", "")).strip()
        if not region_code or region_code == "ALL":
            raise ValueError("시군구별_1인가구_연령별은 시군구 선택이 필요합니다.")

        resolved_override = cfg.get("_resolved_one_person_age_params")
        if resolved_override:
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "Y").upper(),
                "startPrdDe": year,
                "endPrdDe": year,
                **resolved_override,
            }
            try:
                return client.request(params)
            except (InvalidKosisParamsError, EmptyKosisDataError):
                cfg.pop("_resolved_one_person_age_params", None)
                cfg.pop("_resolved_probe_params", None)

        candidate_overrides = [
            {"objL2": region_code},
            {"objL1": region_code},
            {"objL2": region_code, "itmId": "ALL"},
            {"objL1": region_code, "itmId": "ALL"},
            {"objL2": region_code, "itmId": "T1"},
            {"objL1": region_code, "itmId": "T1"},
            {"objL2": region_code, "itmId": "T01"},
            {"objL1": region_code, "itmId": "T01"},
            {"objL2": region_code, "objL1": "ALL"},
            {"objL1": region_code, "objL2": "ALL"},
            {"objL2": region_code, "objL1": "ALL", "itmId": "ALL"},
            {"objL1": region_code, "objL2": "ALL", "itmId": "ALL"},
        ]

        last_error = None
        for idx, override in enumerate(candidate_overrides, 1):
            if getattr(self, "stop_requested", False):
                raise RuntimeError(f"중단요청: {cfg['label']} | {year}")
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "Y").upper(),
                "startPrdDe": year,
                "endPrdDe": year,
                **override,
            }
            try:
                df = client.request(params)
                cfg["_resolved_one_person_age_params"] = dict(override)
                cfg["_resolved_probe_params"] = {k: v for k, v in override.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[1인가구연령차원확정] {cfg['label']} | {year} | 전략={idx}/{len(candidate_overrides)} | 파라미터={cfg['_resolved_probe_params']}")
                return df
            except EmptyKosisDataError:
                cfg["_resolved_one_person_age_params"] = dict(override)
                cfg["_resolved_probe_params"] = {k: v for k, v in override.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[1인가구연령차원확정-데이터없음] {cfg['label']} | {year} | 전략={idx}/{len(candidate_overrides)} | 파라미터={cfg['_resolved_probe_params']}")
                raise
            except InvalidKosisParamsError as e:
                last_error = e
                self.log(f"[1인가구연령차원탐색실패] {cfg['label']} | {year} | 전략={idx}/{len(candidate_overrides)} | 사유={e}")
                continue

        raise RuntimeError(f"시군구별_1인가구_연령별 파라미터 자동 탐색 실패: {cfg['label']} | {year} | 마지막오류={last_error}")

    def _pick_migration_total_item_codes(self, meta_rows):
        items = []
        seen = set()
        for row in meta_rows or []:
            if not isinstance(row, dict):
                continue
            code = str(row.get("itmId") or row.get("ITM_ID") or "").strip()
            name = str(row.get("itmNm") or row.get("ITM_NM") or "").strip()
            if not code or code in seen:
                continue
            items.append({"code": code, "name": name})
            seen.add(code)
        if not items:
            return []

        def has_kw(name, kws):
            n = str(name).strip()
            return any(k in n for k in kws)

        total = [x for x in items if has_kw(x["name"], ["계", "합계", "총계", "전체"])]
        if total:
            return total[:1]
        male = [x for x in items if has_kw(x["name"], ["남"])]
        female = [x for x in items if has_kw(x["name"], ["여"])]
        if male and female:
            return male[:1] + female[:1]
        return items[:2]


    def _pick_migration_sido_tbl_id(self, region_code: str):
        rc = str(region_code or "").strip()
        if rc:
            tbl = MIGRATION_SIDO_TBL_MAP.get(rc[:2])
            if tbl:
                return tbl
        sido_name = str(self.sido_name_var.get() or "").strip()
        if sido_name:
            tbl = MIGRATION_SIDO_NAME_TBL_MAP.get(sido_name)
            if tbl:
                return tbl
        return None

    def _migration_region_code_candidates(self, region_code: str):
        rc = str(region_code or "").strip()
        codes = []
        def add(x):
            x = str(x or '').strip()
            if x and x not in codes:
                codes.append(x)
        add(rc)
        try:
            current_reg = str(self._current_region_values()[2] or "").strip()
            if current_reg:
                add(current_reg[:5])
        except Exception:
            pass
        if len(rc) == 5:
            # 특별자치도 코드 개편 fallback
            if rc.startswith('51'):
                add('42' + rc[2:])
            elif rc.startswith('52'):
                add('45' + rc[2:])
        return codes

    def _filter_migration_region_rows(self, df: pd.DataFrame, region_code: str):
        if df is None or df.empty:
            return df
        code = str(region_code or "").strip()
        sigungu_name = str(self.sigungu_name_var.get() or "").strip()
        sido_name = str(self.sido_name_var.get() or "").strip()
        candidates = []
        if code:
            candidates.append(code)
        if sigungu_name:
            candidates.append(sigungu_name)
        if sido_name and sigungu_name:
            candidates.append(f"{sido_name} {sigungu_name}")
            candidates.append(f"{sido_name}_{sigungu_name}")
        code_cols = [c for c in ["C1", "C2", "C3", "C4"] if c in df.columns]
        name_cols = [c for c in ["C1_NM", "C2_NM", "C3_NM", "C4_NM", "C1_OBJ_NM", "C2_OBJ_NM", "C3_OBJ_NM", "C4_OBJ_NM"] if c in df.columns]
        mask = pd.Series(False, index=df.index)
        for col in code_cols:
            mask = mask | df[col].astype(str).str.strip().isin(candidates)
        for col in name_cols:
            series = df[col].astype(str).str.strip()
            if sigungu_name:
                mask = mask | series.eq(sigungu_name) | series.str.endswith(sigungu_name)
            if code:
                mask = mask | series.eq(code)
        filtered = df[mask].copy()
        if filtered.empty and code_cols:
            # 일부 표는 상위 시도 표 내부에서 C1/C2에 시군구코드를 보관하지 않을 수 있어 앞 2자리 시도코드도 보조 확인
            sido_code = code[:2]
            mask2 = pd.Series(False, index=df.index)
            for col in code_cols:
                mask2 = mask2 | df[col].astype(str).str.strip().eq(sido_code)
            filtered = df[mask2].copy()
        return filtered

    def _filter_weekly_apt_region_df(self, df: pd.DataFrame, cfg: dict):
        if df is None or df.empty:
            return df
        sido_name, sigungu_name, current_reg = self._current_region_values()
        sigungu_name = str(sigungu_name or "").strip()
        sido_name = str(sido_name or "").strip()
        group_name = getattr(self, "active_run_group_name", None) or self._get_current_group_name()
        if self._housing_sido_report_enabled(group_name):
            if not sido_name:
                return df
            sido_core = re.sub(r"(특별자치시|특별자치도|광역시|특별시|자치도|도|시)$", "", sido_name)
            name_candidates = [sido_name]
            if sido_core and sido_core not in name_candidates:
                name_candidates.append(sido_core)
            work = df.copy()
            name_cols = [
                c for c in work.columns
                if str(c).upper().endswith("_NM") or str(c).upper() in {"C1_NM", "C2_NM", "C3_NM", "OBJ_NM"}
            ]
            selected_sido = self._get_kosis_region_code(2) if current_reg else ""
            weekly_prefix = KAB_PRICE_INDEX_SIDO_PREFIX_MAP.get(selected_sido, "")
            code_cols = [c for c in work.columns if str(c).upper() in {"C1", "C2", "C3", "OBJ_ID"}]
            mask = pd.Series(False, index=work.index)
            for col in name_cols:
                series = work[col].astype(str).str.strip()
                compact = series.str.replace(" ", "", regex=False).str.replace("_", "", regex=False)
                for name in name_candidates:
                    mask = mask | series.eq(name) | compact.eq(name.replace(" ", "")) | series.str.contains(name, na=False, regex=False)
            if weekly_prefix:
                for col in code_cols:
                    mask = mask | work[col].astype(str).str.strip().str.startswith(weekly_prefix)
            filtered = work[mask].copy()
            if filtered.empty:
                self.log(f"[주간아파트시도필터주의] {cfg.get('label', '')} | 선택={sido_name} | 원자료 {len(work):,}건 중 매칭 없음, 전체 저장")
                return work
            self.log(f"[주간아파트시도필터] {cfg.get('label', '')} | 선택={sido_name} | {len(work):,}->{len(filtered):,}건")
            return filtered
        if not sigungu_name:
            return df
        sigungu_core = re.sub(r"(시|군|구)$", "", sigungu_name)
        name_candidates = [sigungu_name]
        if sigungu_core and sigungu_core not in name_candidates:
            name_candidates.append(sigungu_core)
        work = df.copy()
        name_cols = [
            c for c in work.columns
            if str(c).upper().endswith("_NM") or str(c).upper() in {"C1_NM", "C2_NM", "C3_NM", "OBJ_NM"}
        ]
        code_cols = [c for c in work.columns if str(c).upper() in {"C1", "C2", "C3", "OBJ_ID"}]
        mask = pd.Series(False, index=work.index)
        for col in name_cols:
            series = work[col].astype(str).str.strip()
            compact = series.str.replace(" ", "", regex=False).str.replace("_", "", regex=False)
            for name in name_candidates:
                mask = mask | series.eq(name) | series.str.endswith(name) | compact.str.endswith(name)
                if sido_name:
                    mask = mask | series.eq(f"{sido_name} {name}") | series.eq(f"{sido_name}_{name}") | compact.eq(f"{sido_name}{name}")

        selected_code = str(getattr(self, "_resolved_weekly_apt_region_code", "") or "").strip()
        if selected_code:
            for col in code_cols:
                mask = mask | work[col].astype(str).str.strip().eq(selected_code)

        filtered = work[mask].copy()
        if filtered.empty:
            self.log(f"[주간아파트지역필터주의] {cfg.get('label', '')} | 선택={sido_name} {sigungu_name} | 원자료 {len(work):,}건 중 매칭 없음, 전체 저장")
            return work
        self.log(f"[주간아파트지역필터] {cfg.get('label', '')} | 선택={sido_name} {sigungu_name} | {len(work):,}->{len(filtered):,}건")
        return filtered

    def _aggregate_migration_region_total_df(self, df: pd.DataFrame, cfg: dict, ym: str, region_code: str):
        filtered = self._filter_migration_region_rows(df, region_code)
        if filtered is None or filtered.empty:
            raise EmptyKosisDataError(f"데이터없음: {cfg['label']} | {ym}")
        work = filtered.copy()
        if "DT" not in work.columns:
            return work
        work["DT"] = pd.to_numeric(work["DT"], errors="coerce")
        work = work[work["DT"].notna()].copy()
        if work.empty:
            raise EmptyKosisDataError(f"데이터없음: {cfg['label']} | {ym}")
        keep_cols = [c for c in ["PRD_DE", "ORG_ID", "TBL_ID", "TBL_NM", "PRD_SE", "LST_CHN_DE", "UNIT_NM"] if c in work.columns]
        if keep_cols:
            agg = work.groupby(keep_cols, dropna=False, as_index=False)["DT"].sum()
        else:
            agg = pd.DataFrame({"DT": [work["DT"].sum()]})
        sido_name, sigungu_name, _ = self._current_region_values()
        region_name = sigungu_name or sido_name or "선택지역"
        agg["ITM_ID"] = "TOTAL"
        agg["ITM_NM"] = "이동자수(지역합산총계)"
        agg["C1"] = region_code
        agg["C1_NM"] = region_name
        agg["C1_OBJ_NM"] = "시군구"
        return agg

    def _aggregate_migration_total_df(self, frames, cfg, ym):
        if not frames:
            raise EmptyKosisDataError(f"데이터없음: {cfg['label']} | {ym}")
        df = pd.concat(frames, ignore_index=True)
        if df.empty or "DT" not in df.columns:
            return df
        work = df.copy()
        work["DT"] = pd.to_numeric(work["DT"], errors="coerce")
        if work["DT"].notna().sum() == 0:
            return df
        keep_cols = [c for c in ["PRD_DE", "ORG_ID", "TBL_ID", "TBL_NM", "PRD_SE", "LST_CHN_DE", "UNIT_NM"] if c in work.columns]
        if not keep_cols:
            keep_cols = ["PRD_DE"] if "PRD_DE" in work.columns else []
        group_cols = keep_cols[:] if keep_cols else None
        if group_cols:
            agg = work.groupby(group_cols, dropna=False, as_index=False)["DT"].sum()
        else:
            agg = pd.DataFrame({"DT": [work["DT"].sum()]})
        sido_name, sigungu_name, current_reg = self._current_region_values()
        region_name = sigungu_name or sido_name or "선택지역"
        if "ITM_ID" in agg.columns:
            agg["ITM_ID"] = "TOTAL"
        else:
            agg.insert(len(agg.columns), "ITM_ID", "TOTAL")
        agg["ITM_NM"] = "이동자수(성별합산총계)"
        agg["C1"] = current_reg
        agg["C1_NM"] = region_name
        agg["C1_OBJ_NM"] = "시군구"
        return agg

    def _request_migration_sigungu_month_df(self, client: KosisClient, cfg: dict, ym: str):
        region_code = self._get_kosis_region_code(cfg.get("region_code_length", 5))

        # 0순위: 제미나이 성공 방식 우선 시도
        base_tbl = "DT_1B26001_A01"
        last_base_error = None
        base_empty_seen = False
        base_invalid_seen = False
        for code in self._migration_region_code_candidates(region_code):
            params = {
                "orgId": cfg["orgId"],
                "tblId": base_tbl,
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                "objL1": code,
                "itmId": "ALL",
            }
            try:
                df_base = client.request(params)
                if df_base is not None and not df_base.empty and "DT" in df_base.columns:
                    work = df_base.copy()
                    work["DT"] = pd.to_numeric(work["DT"], errors="coerce")
                    work = work[work["DT"].notna()].copy()
                    if not work.empty and float(work["DT"].abs().sum(skipna=True) or 0) == 0:
                        self.log(f"[이동자수0값] {cfg['label']} | {ym} | objL1={code} | 값이 모두 0이라 다른 코드로 재조회")
                        work = work.iloc[0:0]
                    if not work.empty:
                        cfg["_resolved_migration_tblId"] = base_tbl
                        cfg["_resolved_migration_params"] = {"objL1": code, "itmId": "ALL"}
                        self.log(f"[이동자수본표확정] {cfg['label']} | {ym} | tblId={base_tbl} | objL1={code},itmId=ALL")
                        return work
                self.log(f"[이동자수본표실패-0건] {cfg['label']} | {ym} | tblId={base_tbl} | objL1={code}")
                last_base_error = EmptyKosisDataError(f"0건 응답: objL1={code}")
                base_empty_seen = True
            except EmptyKosisDataError as e:
                last_base_error = e
                base_empty_seen = True
                self.log(f"[이동자수본표실패] {cfg['label']} | {ym} | tblId={base_tbl} | objL1={code} | 사유={e}")
                continue
            except InvalidKosisParamsError as e:
                last_base_error = e
                base_invalid_seen = True
                self.log(f"[이동자수본표실패] {cfg['label']} | {ym} | tblId={base_tbl} | objL1={code} | 사유={e}")
                continue

        # 본표에서 데이터 없음이 확인된 경우에는 시도표 우회 탐색으로 넘어가지 않음.
        # 최신 월 미공표/반영 지연일 가능성이 높으므로 해당 월만 데이터 없음으로 처리.
        if base_empty_seen and not base_invalid_seen:
            raise EmptyKosisDataError(f"데이터 없음: {cfg['label']} | {ym} | tblId={base_tbl} | objL1 후보={self._migration_region_code_candidates(region_code)}")

        selected_tbl = self._pick_migration_sido_tbl_id(region_code)
        if selected_tbl:
            if cfg.get("_resolved_migration_tblId") != selected_tbl:
                cfg["_resolved_migration_tblId"] = selected_tbl
                self.log(f"[이동자수시도표선택] {cfg['label']} | {ym} | tblId={selected_tbl}")
            candidate_overrides = []
            def _add_override(override):
                if override not in candidate_overrides:
                    candidate_overrides.append(override)

            # 1차: 필수 축 최소 조합 및 itmId(항목코드) 추가 탐색
            itm_candidates = ["ALL", "T1", "T2", "T01", "1", "0"]
            base_regions = [
                {"objL1": region_code},
                {"objL2": region_code},
                {"objL1": region_code, "objL2": "ALL"},
                {"objL2": region_code, "objL1": "ALL"},
            ]

            # 완전 무파라미터/ALL 조합도 소수만 유지
            for override in [
                {},
                {"objL1": "ALL"},
                {"objL2": "ALL"},
                {"objL1": "ALL", "objL2": "ALL"},
            ]:
                _add_override(override)

            for override in base_regions:
                _add_override(override)
                for itm in itm_candidates:
                    _add_override({**override, "itmId": itm})

            # 2차: 시도별 각세별 표는 보통 추가 분류축(objL3/objL4)이 더 필요한 경우가 많음
            obj_total_candidates = ["ALL", "00", "0", "000", "100"]
            for base in base_regions:
                for total_code in obj_total_candidates:
                    _add_override({**base, "objL3": total_code})
                    _add_override({**base, "objL4": total_code})
                    _add_override({**base, "objL3": total_code, "objL4": total_code})
                    for itm in ["ALL", "T1"]:
                        _add_override({**base, "objL3": total_code, "itmId": itm})
                        _add_override({**base, "objL4": total_code, "itmId": itm})
                        _add_override({**base, "objL3": total_code, "objL4": total_code, "itmId": itm})

            max_candidates = 60
            candidate_overrides = candidate_overrides[:max_candidates]
            self.log(f"[이동자수시도표추가탐색] {cfg['label']} | {ym} | tblId={selected_tbl} | 후보수={len(candidate_overrides)}")
            last_error = None
            for idx, override in enumerate(candidate_overrides, 1):
                if self.stop_event.is_set():
                    self.log(f"[중단요청즉시반영] {cfg['label']} | {ym} | 시도표 탐색 중단")
                    raise RuntimeError(f"중단요청: {cfg['label']} | {ym}")
                params = {
                    "orgId": cfg["orgId"],
                    "tblId": selected_tbl,
                    "prdSe": cfg.get("prdSe", "M").upper(),
                    "startPrdDe": ym,
                    "endPrdDe": ym,
                    **override,
                }
                try:
                    df_raw = client.request(params)
                    agg = self._aggregate_migration_region_total_df(df_raw, cfg, ym, region_code)
                    cfg["_resolved_migration_params"] = dict(override)
                    cfg["_resolved_migration_tblId"] = selected_tbl
                    self.log(f"[이동자수차원확정] {cfg['label']} | {ym} | 시도표={selected_tbl} | 전략={idx}/{len(candidate_overrides)} | 파라미터={override}")
                    return agg
                except EmptyKosisDataError:
                    # 0건 응답은 유효 조합으로 확정하지 않는다.
                    self.log(f"[이동자수차원탐색실패-0건] {cfg['label']} | {ym} | 시도표={selected_tbl} | 전략={idx}/{len(candidate_overrides)} | 파라미터={override}")
                    last_error = EmptyKosisDataError(f"0건 응답: {override}")
                    continue
                except InvalidKosisParamsError as e:
                    last_error = e
                    self.log(f"[이동자수차원탐색실패] {cfg['label']} | {ym} | 시도표={selected_tbl} | 전략={idx}/{len(candidate_overrides)} | 사유={e}")
                    continue
            self.log(f"[이동자수시도표실패] {cfg['label']} | {ym} | tblId={selected_tbl} | 마지막오류={last_error}")

        if str(cfg.get("tblId", "")).strip() == "DT_1B26001_A02":
            meta_rows = []
            try:
                meta_rows = client.get_meta_itm(cfg["orgId"], cfg["tblId"])
            except Exception as e:
                self.log(f"[이동자수메타실패] {cfg['label']} | {ym} | 사유={e}")
            target_items = self._pick_migration_total_item_codes(meta_rows)
            if not target_items:
                target_items = [{"code": x, "name": x} for x in ["T1", "T2"]]
            candidate_region_keys = ["objL1", "objL2"]
            last_error = None
            for region_key in candidate_region_keys:
                frames = []
                used = []
                success_any = False
                for item in target_items:
                    params = {
                        "orgId": cfg["orgId"],
                        "tblId": cfg["tblId"],
                        "prdSe": cfg.get("prdSe", "M").upper(),
                        "startPrdDe": ym,
                        "endPrdDe": ym,
                        region_key: region_code,
                        "itmId": item["code"],
                    }
                    try:
                        df_item = client.request(params)
                        success_any = True
                        frames.append(df_item)
                        used.append(f"{region_key}={region_code},itmId={item['code']}")
                    except EmptyKosisDataError:
                        success_any = True
                        used.append(f"{region_key}={region_code},itmId={item['code']}(데이터없음)")
                        continue
                    except InvalidKosisParamsError as e:
                        last_error = e
                        self.log(f"[이동자수차원탐색실패] {cfg['label']} | {ym} | 후보={region_key},{item['code']} | 사유={e}")
                        frames = []
                        used = []
                        success_any = False
                        break
                if success_any:
                    if frames:
                        cfg["_resolved_migration_params"] = {"region_key": region_key, "item_codes": [x["code"] for x in target_items]}
                        self.log(f"[이동자수차원확정] {cfg['label']} | {ym} | 파라미터={used}")
                        return self._aggregate_migration_total_df(frames, cfg, ym)
                    self.log(f"[이동자수차원탐색실패-0건] {cfg['label']} | {ym} | 파라미터={used}")
                    last_error = EmptyKosisDataError(f"데이터없음: {cfg['label']} | {ym}")
                    continue
            raise RuntimeError(f"시군구별_이동자수 파라미터 자동 탐색 실패: {cfg['label']} | {ym} | 마지막오류={last_error}")

        resolved_override = cfg.get("_resolved_migration_params")
        if resolved_override:
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg.get("_resolved_migration_tblId", cfg["tblId"]),
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                **resolved_override,
            }
            try:
                return client.request(params)
            except EmptyKosisDataError:
                raise
            except InvalidKosisParamsError:
                cfg.pop("_resolved_migration_params", None)
                cfg.pop("_resolved_migration_probe_params", None)

        candidate_overrides = [
            {"objL2": region_code, "itmId": "T1"},
            {"objL1": region_code, "itmId": "T1"},
            {"objL2": region_code, "itmId": "T2"},
            {"objL1": region_code, "itmId": "T2"},
            {"objL2": region_code, "itmId": "T3"},
            {"objL1": region_code, "itmId": "T3"},
            {"objL2": region_code, "itmId": "T01"},
            {"objL1": region_code, "itmId": "T01"},
            {"objL2": region_code, "itmId": "1"},
            {"objL1": region_code, "itmId": "1"},
            {"objL2": region_code, "objL1": "00", "itmId": "T1"},
            {"objL1": region_code, "objL2": "00", "itmId": "T1"},
        ]
        last_error = None
        for idx, override in enumerate(candidate_overrides, 1):
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                **override,
            }
            try:
                df = client.request(params)
                cfg["_resolved_migration_params"] = dict(override)
                cfg["_resolved_migration_probe_params"] = {k: v for k, v in override.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[이동자수차원확정] {cfg['label']} | {ym} | 전략={idx}/{len(candidate_overrides)} | 파라미터={cfg['_resolved_migration_probe_params']}")
                return df
            except EmptyKosisDataError:
                self.log(f"[이동자수차원탐색실패-0건] {cfg['label']} | {ym} | 전략={idx}/{len(candidate_overrides)} | 파라미터={override}")
                last_error = EmptyKosisDataError(f"0건 응답: {override}")
                continue
            except InvalidKosisParamsError as e:
                last_error = e
                self.log(f"[이동자수차원탐색실패] {cfg['label']} | {ym} | 전략={idx}/{len(candidate_overrides)} | 사유={e}")
                continue
        raise RuntimeError(f"시군구별_이동자수 파라미터 자동 탐색 실패: {cfg['label']} | {ym} | 마지막오류={last_error}")

    def _fetch_stathtml_dim_options(self, cfg: dict):
        url = "https://kosis.kr/statHtml/statHtml.do"
        params = {
            "orgId": cfg.get("orgId"),
            "tblId": cfg.get("tblId"),
            "vw_cd": "MT_ZTITLE",
            "conn_path": "MT_ZTITLE",
            "lang_mode": "ko",
        }
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        }
        resp = requests.get(url, params=params, headers=headers, timeout=DEFAULT_TIMEOUT)
        resp.raise_for_status()
        html = resp.text
        results = {}
        for key in ["itmId", "objL1", "objL2", "objL3", "objL4"]:
            codes = []
            seen = set()
            for patt in [
                rf"<select[^>]+(?:id|name)=[\"']{key}[\"'][^>]*>(.*?)</select>",
                rf"<select[^>]+(?:id|name)=[\"']{key.lower()}[\"'][^>]*>(.*?)</select>",
            ]:
                for m in re.finditer(patt, html, re.I | re.S):
                    block = m.group(1)
                    for om in re.finditer(r"<option[^>]*value=[\"']([^\"']+)[\"'][^>]*>(.*?)</option>", block, re.I | re.S):
                        code = re.sub(r'\s+', ' ', om.group(1)).strip()
                        name = re.sub(r'<.*?>', '', om.group(2)).strip()
                        if code and code not in seen:
                            seen.add(code)
                            codes.append({"code": code, "name": name})
            # hidden/input fallback
            for im in re.finditer(rf"<input[^>]+(?:id|name)=[\"']{key}[\"'][^>]+value=[\"']([^\"']+)[\"']", html, re.I):
                code = im.group(1).strip()
                if code and code not in seen:
                    seen.add(code)
                    codes.append({"code": code, "name": code})
            if codes:
                results[key] = codes
        return results

    def _pick_total_codes(self, codes):
        if not codes:
            return []
        picked = []
        seen = set()
        preferred_codes = ["ALL", "00", "0", "000", "100", "999", "T001", "T100"]
        preferred_names = ["계", "총계", "합계", "전체", "소계"]
        for c in codes:
            code = str(c.get("code", "")).strip()
            name = str(c.get("name", "")).strip()
            if code in preferred_codes or any(n and n in name for n in preferred_names):
                if code and code not in seen:
                    seen.add(code)
                    picked.append(code)
        for c in codes:
            code = str(c.get("code", "")).strip()
            if code and code not in seen:
                seen.add(code)
                picked.append(code)
            if len(picked) >= 8:
                break
        return picked

    def _region_code_fallbacks(self, region_code: str):
        code = str(region_code or "").strip()
        results = []
        seen = set()

        def add(v):
            v = str(v or "").strip()
            if v and v not in seen:
                seen.add(v)
                results.append(v)

        add(code)
        legacy_code = KOSIS_LEGACY_REGION_CODE_MAP.get(code)
        if legacy_code:
            add(legacy_code)
        if len(code) >= 2:
            add(code[:2])
            legacy_sido = KOSIS_LEGACY_SIDO_CODE_MAP.get(code[:2])
            if legacy_sido:
                add(legacy_sido)
            # 특별자치도 전환 전 구코드 fallback
            if code.startswith("51"):
                add("42" + code[2:] if len(code) == 5 else "42")
            elif code.startswith("52"):
                add("45" + code[2:] if len(code) == 5 else "45")
        return results


    def _molit_sido_short_names(self, sido_name: str, sigungu_name: str = "") -> list:
        """국토부 미분양 공표의 시도 표기 후보.
        2026-07 통합 이후 '전남광주' 표기가 사용되고, 그 이전 월은 종전 '광주'/'전남' 표기를 사용한다.
        """
        names = [self._molit_sido_short_name(sido_name, sigungu_name)]
        if (sido_name or "").strip() == "전남광주통합특별시":
            for extra in ["전남광주", "광주", "전남"]:
                if extra not in names:
                    names.append(extra)
        return names

    def _molit_sido_short_name(self, sido_name: str, sigungu_name: str = "") -> str:
        # 통합 행정구역은 미분양 공표 통계가 아직 종전 광주/전남 명칭을 사용한다.
        if (sido_name or "").strip() == "전남광주통합특별시":
            return "광주" if (sigungu_name or "").strip() in GWANGJU_AREA_SIGUNGU_NAMES else "전남"
        mapping = {
            "서울특별시": "서울", "부산광역시": "부산", "대구광역시": "대구", "인천광역시": "인천",
            "광주광역시": "광주", "대전광역시": "대전", "울산광역시": "울산", "세종특별자치시": "세종",
            "경기도": "경기", "강원특별자치도": "강원", "충청북도": "충북", "충청남도": "충남",
            "전북특별자치도": "전북", "전라남도": "전남", "경상북도": "경북", "경상남도": "경남",
            "제주특별자치도": "제주",
        }
        return mapping.get((sido_name or "").strip(), (sido_name or "").strip())

    def _request_molit_unsold_month_df(self, cfg: dict, ym: str):
        sido_name, sigungu_name, current_reg = self._current_region_values()
        if not current_reg or not sido_name or not sigungu_name:
            raise ValueError(f"{cfg['label']}은 시도/시군구를 먼저 선택해야 합니다.")

        params = {
            "key": MOLIT_DEFAULT_KEY,
            "form_id": cfg.get("molit_form_id", "2082"),
            "style_num": cfg.get("molit_style_num", "128"),
            "start_dt": ym,
            "end_dt": ym,
        }
        try:
            resp = requests.get(MOLIT_OPENAPI_URL, params=params, timeout=DEFAULT_TIMEOUT)
            if resp.status_code >= 500:
                raise EmptyKosisDataError(f"데이터 없음 또는 미공표: {MOLIT_OPENAPI_URL}?form_id={params['form_id']}&start_dt={ym}&end_dt={ym}")
            resp.raise_for_status()
            payload = resp.json()
        except EmptyKosisDataError:
            raise
        except Exception as e:
            raise RuntimeError(f"국토부 OpenAPI 요청 실패: {e}")

        status = ((payload or {}).get("result_status") or {})
        status_code = status.get("status_code")
        if status_code in {"INFO-200", "ERROR-500"}:
            raise EmptyKosisDataError(f"데이터 없음 또는 미공표: {MOLIT_OPENAPI_URL}?form_id={params['form_id']}&start_dt={ym}&end_dt={ym}")
        if status_code != "INFO-000":
            raise RuntimeError(f"국토부 OpenAPI 오류: {status}")

        result_data = (payload or {}).get("result_data") or {}
        form_list = result_data.get("formList") or []
        if not form_list:
            raise EmptyKosisDataError(f"데이터 없음: {MOLIT_OPENAPI_URL}?form_id={params['form_id']}&start_dt={ym}&end_dt={ym}")

        df = pd.DataFrame(form_list)
        if df.empty:
            raise EmptyKosisDataError(f"데이터 없음: {MOLIT_OPENAPI_URL}?form_id={params['form_id']}&start_dt={ym}&end_dt={ym}")

        sido_short = self._molit_sido_short_name(sido_name, sigungu_name)
        sigungu = sigungu_name
        if "구분" in df.columns:
            df["구분"] = df["구분"].astype(str).str.strip()
        if "시군구" in df.columns:
            df["시군구"] = df["시군구"].astype(str).str.strip()

        sido_candidates = self._molit_sido_short_names(sido_name, sigungu_name)
        filtered = df[df.get("구분", "").isin(sido_candidates) & (df.get("시군구", "") == sigungu)].copy()
        if not filtered.empty and str(filtered.iloc[0].get("구분", "")) != sido_short:
            self.log(f"[미분양시도표기전환] {cfg['label']} | {ym} | 표기={filtered.iloc[0].get('구분')}")
        for k, v in (cfg.get("molit_extra_filters") or {}).items():
            if k in filtered.columns:
                filtered[k] = filtered[k].astype(str).str.strip()
                filtered = filtered[filtered[k] == str(v).strip()].copy()
        if filtered.empty:
            raise EmptyKosisDataError(f"데이터 없음: {MOLIT_OPENAPI_URL}?form_id={params['form_id']}&start_dt={ym}&end_dt={ym} | region={sido_short}/{sigungu}")

        value_col = cfg.get("molit_value_col")
        if not value_col or value_col not in filtered.columns:
            candidates = [c for c in filtered.columns if c not in {"date", "구분", "시군구", "규모", "부문"}]
            value_col = candidates[0] if candidates else None
        if not value_col:
            raise RuntimeError(f"국토부 OpenAPI 값 컬럼 탐지 실패: {cfg['label']} | cols={list(filtered.columns)}")

        filtered["ORG_ID"] = cfg.get("orgId", "116")
        filtered["TBL_ID"] = cfg.get("tblId", "DT_MLTM_2082")
        filtered["TBL_NM"] = cfg.get("label", result_data.get("formName", "미분양현황"))
        filtered["PRD_DE"] = filtered.get("date", ym)
        filtered["DT"] = pd.to_numeric(filtered.get(value_col), errors="coerce")
        filtered["UNIT_NM"] = result_data.get("unitName", "호")
        filtered["C1_NM"] = filtered.get("구분", sido_short)
        filtered["C2_NM"] = filtered.get("시군구", sigungu)
        filtered["C1"] = self._get_selected_region_code(2)
        filtered["C2"] = self._get_selected_region_code(5)
        filtered["ITM_NM"] = cfg.get("label", result_data.get("formName", "미분양현황"))
        filtered["ITM_ID"] = cfg.get("molit_item_id", "UNSOLD")
        filtered["PRD_SE"] = "M"
        cols = [c for c in ["date", "구분", "시군구", "규모", "부문", value_col, "DT", "PRD_DE", "C1", "C2", "C1_NM", "C2_NM", "ITM_ID", "ITM_NM", "UNIT_NM", "ORG_ID", "TBL_ID", "TBL_NM"] if c in filtered.columns]
        return filtered[cols].reset_index(drop=True)

    def _build_unsold_candidates(self, cfg: dict, ym: str):
        # auto_region_probe 경로에서는 params에 지역코드가 아직 주입되지 않으므로
        # 선택된 시군구/시도 코드에서 직접 가져오도록 보강
        region_code = str(cfg.get("params", {}).get("objL1", "")).strip()
        if (not region_code or region_code == "ALL") and hasattr(self, "region_code_var"):
            selected = str(self._current_region_values()[2]).strip()
            if selected:
                code_len = int(cfg.get("region_code_length", 5))
                region_code = self._get_kosis_region_code(code_len)
        if not region_code or region_code == "ALL":
            raise ValueError("미분양주택현황_최근15개월은 지역 선택이 필요합니다.")
        region_codes = self._region_code_fallbacks(region_code)
        candidates = []
        base = {
            "orgId": cfg["orgId"],
            "tblId": cfg["tblId"],
            "prdSe": cfg.get("prdSe", "M").upper(),
            "startPrdDe": ym,
            "endPrdDe": ym,
        }
        opt_map = {}
        try:
            opt_map = self._fetch_stathtml_dim_options(cfg)
            self.log(f"[미분양차원탐색] {cfg['label']} | {ym} | 화면옵션 keys={','.join(opt_map.keys()) if opt_map else '없음'}")
        except Exception as e:
            self.log(f"[미분양차원탐색] {cfg['label']} | {ym} | 화면옵션추출실패={e}")
        itm_codes = self._pick_total_codes(opt_map.get("itmId", [])) or ["ALL", "T001", "T100"]
        o2_codes = self._pick_total_codes(opt_map.get("objL2", [])) or ["ALL", "00", "0"]
        o3_codes = self._pick_total_codes(opt_map.get("objL3", [])) or ["ALL", "00", "0"]
        o4_codes = self._pick_total_codes(opt_map.get("objL4", [])) or ["ALL", "00", "0"]

        def add(name, payload):
            merged = {**base, **payload}
            if not any(existing[1] == merged for existing in candidates):
                candidates.append((name, merged))

        for rc in region_codes:
            # region-first minimal
            add(f"objL1={rc}", {"objL1": rc})
            add(f"objL2={rc}", {"objL2": rc})
            # region + itm
            for itm in itm_codes[:5]:
                add(f"objL1={rc},itmId={itm}", {"objL1": rc, "itmId": itm})
                add(f"objL2={rc},itmId={itm}", {"objL2": rc, "itmId": itm})
            # region + one extra dim
            for extra in o2_codes[:4]:
                add(f"objL1={rc},objL2={extra}", {"objL1": rc, "objL2": extra})
            for extra in o3_codes[:4]:
                add(f"objL1={rc},objL3={extra}", {"objL1": rc, "objL3": extra})
            for extra in o4_codes[:4]:
                add(f"objL1={rc},objL4={extra}", {"objL1": rc, "objL4": extra})
            # region + itm + one extra dim
            for itm in itm_codes[:3]:
                for extra in o2_codes[:3]:
                    add(f"objL1={rc},objL2={extra},itmId={itm}", {"objL1": rc, "objL2": extra, "itmId": itm})
                for extra in o3_codes[:3]:
                    add(f"objL1={rc},objL3={extra},itmId={itm}", {"objL1": rc, "objL3": extra, "itmId": itm})
                for extra in o4_codes[:3]:
                    add(f"objL1={rc},objL4={extra},itmId={itm}", {"objL1": rc, "objL4": extra, "itmId": itm})
            # region + multi extra dims
            for e2 in o2_codes[:2]:
                for e3 in o3_codes[:2]:
                    add(f"objL1={rc},objL2={e2},objL3={e3}", {"objL1": rc, "objL2": e2, "objL3": e3})
            for e3 in o3_codes[:2]:
                for e4 in o4_codes[:2]:
                    add(f"objL1={rc},objL3={e3},objL4={e4}", {"objL1": rc, "objL3": e3, "objL4": e4})
        return candidates[:80]

    def _request_unsold_sigungu_month_df(self, client: KosisClient, cfg: dict, ym: str):
        resolved = cfg.get("_resolved_unsold_params")
        if resolved:
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                **resolved,
            }
            try:
                return client.request(params)
            except (InvalidKosisParamsError, EmptyKosisDataError):
                cfg.pop("_resolved_unsold_params", None)

        candidates = self._build_unsold_candidates(cfg, ym)
        last_error = None
        for idx, (name, params) in enumerate(candidates, 1):
            if self.stop_event.is_set():
                raise RuntimeError(f"중단요청: {cfg['label']} | {ym}")
            try:
                df = client.request(params)
                cfg["_resolved_unsold_params"] = {k: v for k, v in params.items() if k.startswith("objL") or k == "itmId"}
                self.log(f"[미분양차원확정] {cfg['label']} | {ym} | 전략={idx}/{len(candidates)} {name} | 파라미터={cfg['_resolved_unsold_params']}")
                return df
            except EmptyKosisDataError as e:
                last_error = e
                self.log(f"[미분양차원탐색실패-0건] {cfg['label']} | {ym} | 전략={idx}/{len(candidates)} {name} | 사유={e}")
                continue
            except InvalidKosisParamsError as e:
                last_error = e
                self.log(f"[미분양차원탐색실패] {cfg['label']} | {ym} | 전략={idx}/{len(candidates)} {name} | 사유={e}")
                continue
        raise RuntimeError(f"미분양주택현황_최근15개월 파라미터 자동 탐색 실패: {cfg['label']} | {ym} | 마지막오류={last_error}")

    def _build_cumulative_point_months(self, end_period: str, years: int = 6) -> list[str]:
        """월별 누계형 전국 통계용 기준월 목록.

        구성: 해당년도 포함 최근 N개 연도.
        - 과거 연도: 매년 12월 누계
        - 당해 연도: 화면에서 선택한 종료월 누계

        예: 종료월 2026-04, years=6 ->
        202112, 202212, 202312, 202412, 202512, 202604
        """
        end_dt = parse_end_period(end_period)
        start_year = end_dt.year - years + 1
        months = [f"{year}12" for year in range(start_year, end_dt.year)]
        months.append(end_dt.strftime("%Y%m"))
        return months

    def _request_cumulative_point_df(self, client: KosisClient, cfg: dict, end_period: str) -> tuple[pd.DataFrame, str]:
        """전국 월별 누계 통계를 최근 6년 기준월만 단건 조회.

        인허가·착공·준공·분양은 월별 누계 통계이므로 전체 월을 받지 않고
        해당년도 포함 최근 6년 중 과거 연도는 12월 누계, 당해연도는
        화면에서 선택한 종료월 누계를 조회한다.
        """
        years = int(cfg.get("cumulative_years", cfg.get("lookback_years", 6)))
        target_months = self._build_cumulative_point_months(end_period, years=years)
        current_year_month = target_months[-1] if target_months else ""
        frames = []
        found_months = []

        def _one_month(ym: str, current_year_point: bool = False):
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": cfg.get("prdSe", "M").upper(),
                "startPrdDe": ym,
                "endPrdDe": ym,
                **cfg.get("params", {}),
            }
            shown_params = {k: v for k, v in params.items() if k.startswith("objL") or k == "itmId"}
            self.log(f"[누계단건조회] {cfg['label']} | {ym} | 파라미터={shown_params}")
            df = client.request(params)
            if not df.empty:
                df = df.copy()
                df["조회구분"] = "당해최종공표월누계" if current_year_point else "연말12월누계"
                df["조회기준월"] = ym
            return df

        for ym in target_months:
            if self.stop_event.is_set():
                break
            is_current_point = ym == current_year_month
            candidates = [ym]
            if is_current_point:
                end_point = datetime.strptime(ym + "01", "%Y%m%d")
                # 선택월 미공표 시 당해연도 안에서 가장 최근 공표 누계월을 사용한다.
                candidates = [
                    (end_point - relativedelta(months=i)).strftime("%Y%m")
                    for i in range(end_point.month)
                ]
            for candidate_ym in candidates:
                try:
                    df = _one_month(candidate_ym, current_year_point=is_current_point)
                    frames.append(df)
                    found_months.append(candidate_ym)
                    if is_current_point and candidate_ym != ym:
                        self.log(f"[당해최종공표월적용] {cfg['label']} | 선택={ym} -> 공표={candidate_ym}")
                    break
                except EmptyKosisDataError:
                    self.log(f"[데이터없음] {cfg['label']} | 누계 기준월 {candidate_ym}")

        frames = [df for df in frames if df is not None and not df.empty]
        if not frames:
            return pd.DataFrame(), "no_data"

        result = pd.concat(frames, ignore_index=True)
        return result, "_".join([m for m in found_months if m])

    # ------------------------------------------------------------------ R-ONE
    def _rone_api_key(self) -> str:
        try:
            var = getattr(self, "rone_api_key_var", None)
            value = var.get().strip() if var is not None else ""
        except Exception:
            value = ""
        return value or str(load_app_config().get("rone_api_key", "") or "").strip() or DEFAULT_RONE_API_KEY

    def _rone_region_paths(self, paths: list) -> list:
        """R-ONE 지역 경로(예: '대구>수성구') 중 선택 지역과 그 상위 지역을 고른다. 시도 보고서는 시도만."""
        sido_name, sigungu_name, _ = self._current_region_values()
        sido_core = _rone_sido_core(sido_name)
        sido_mode = self._sido_report_enabled(self._get_current_group_name()) or not str(sigungu_name or "").strip()
        split = {p: str(p).split(">") for p in paths}
        if sido_mode:
            keep = [p for p, seg in split.items() if seg[0] == sido_core and (len(seg) == 1 or (sido_core == "전남광주" and len(seg) == 2))]
            return keep
        tokens = str(sigungu_name).split()
        last = _rone_norm(tokens[-1])
        firsts = [_rone_norm(t) for t in tokens[:-1]]
        hits = [p for p, seg in split.items()
                if seg[0] == sido_core and len(seg) >= 2 and _rone_norm(seg[-1]) == last
                and all(any(f in _rone_norm(x) for x in seg[:-1]) for f in firsts)]
        if not hits:
            return []
        keep = set()
        for p in hits:
            seg = split[p]
            for i in range(1, len(seg) + 1):
                keep.add(">".join(seg[:i]))
        return [p for p in paths if p in keep]

    def _rone_probe(self, key: str, tbl: str, cycle: str, end_period: str) -> list:
        """지역 목록을 얻기 위해 최근 공표 시점 하나의 전 지역 자료를 받는다."""
        if cycle == "WK":
            return self._rone_latest_rows(key, tbl, cycle)
        for back in range(0, 4):
            ym = _rone_month_shift(end_period[:6], -back)
            rows = rone_request(key, {"STATBL_ID": tbl, "DTACYCLE_CD": cycle, "WRTTIME_IDTFR_ID": ym})
            if rows:
                return rows
        return []

    def _rone_latest_rows(self, key: str, tbl: str, cycle: str) -> list:
        """주간표: 전국(CLS_ID 첫 항목) 시계열에서 최신 주를 찾아 그 주의 전 지역 자료를 받는다."""
        series = rone_request(key, {"STATBL_ID": tbl, "DTACYCLE_CD": cycle, "CLS_ID": "50001"})
        if not series:
            return []
        latest = max(r["WRTTIME_IDTFR_ID"] for r in series)
        return rone_request(key, {"STATBL_ID": tbl, "DTACYCLE_CD": cycle, "WRTTIME_IDTFR_ID": latest})

    def _fetch_rone_df(self, cfg: dict, yr: dict) -> pd.DataFrame:
        spec = cfg.get("rone") or {}
        key = self._rone_api_key()
        if not spec or not key:
            return pd.DataFrame()
        kind, cycle = spec["kind"], spec["cycle"]
        start, end = str(yr.get("start", "")), str(yr.get("end", ""))
        today = datetime.now().strftime("%Y-%m-%d")
        out = []

        def base_row(tbl_id: str, period: str, value) -> dict:
            return {"PRD_DE": period, "PRD_SE": "D" if cycle == "WK" else "M", "DT": value, "ITM_ID": spec["itm_id"], "ITM_NM": spec["itm_nm"],
                    "TBL_ID": cfg["tblId"], "TBL_NM": spec["tbl_nm"], "ORG_ID": "408", "UNIT_NM": spec["unit"], "LST_CHN_DE": today,
                    "출처": f"한국부동산원 R-ONE OpenAPI({tbl_id})"}

        if kind in {"type_region", "conversion"}:
            first_month = _rone_month_shift(start[:6], -1)
            for tbl, type_name, type_code in spec["tables"]:
                probe = self._rone_probe(key, tbl, cycle, end)
                path_to_id = {r["CLS_FULLNM"]: r["CLS_ID"] for r in probe}
                if kind == "conversion":
                    # 공표 전환율 매칭은 시도 단위(분류2코드 a+1~2자리)로 하므로 시도와 통합특별시의 종전 시도(광주·전남)까지 둔다.
                    paths = [p for p in path_to_id if ">" not in p or p in {"전남광주>광주", "전남광주>전남"}]
                else:
                    paths = self._rone_region_paths(list(path_to_id))
                if not paths:
                    raise RuntimeError(f"R-ONE 지역 매칭 실패: {cfg['label']} {tbl}")
                sido_codes = {p: f"a{i + 1}" for i, p in enumerate(sorted(paths))}
                for path in paths:
                    rows = rone_request(key, {"STATBL_ID": tbl, "DTACYCLE_CD": cycle, "CLS_ID": path_to_id[path], "START_WRTTIME": first_month, "END_WRTTIME": end[:6]})
                    for r in rows:
                        rec = base_row(tbl, str(r["WRTTIME_IDTFR_ID"]), r["DTA_VAL"])
                        rec.update({"C1_OBJ_NM": "주택유형별", "C1_NM": type_name, "C1": type_code,
                                    "C2_OBJ_NM": "지역별", "C2_NM": str(path).split(">")[-1],
                                    "C2": sido_codes[path] if kind == "conversion" else f"r{r['CLS_ID']}"})
                        out.append(rec)
        elif kind == "officetel":
            first_month = _rone_month_shift(start[:6], -1)
            tbl = spec["tables"][0][0]
            rows = rone_request(key, {"STATBL_ID": tbl, "DTACYCLE_CD": cycle, "START_WRTTIME": first_month, "END_WRTTIME": end[:6]})
            for r in rows:
                rec = base_row(tbl, str(r["WRTTIME_IDTFR_ID"]), r["DTA_VAL"])
                rec.update({"C1_OBJ_NM": "지역별", "C1_NM": r.get("GRP_NM"), "C1": f"r{r.get('GRP_ID')}",
                            "C2_OBJ_NM": "규모별", "C2_NM": r.get("CLS_NM"), "C2": f"r{r.get('CLS_ID')}"})
                out.append(rec)
        elif kind == "weekly":
            tbl = spec["tables"][0][0]
            probe = self._rone_latest_rows(key, tbl, cycle)
            path_to_id = {r["CLS_FULLNM"]: r["CLS_ID"] for r in probe}
            paths = self._rone_region_paths(list(path_to_id))
            if not paths:
                raise RuntimeError(f"R-ONE 지역 매칭 실패: {cfg['label']}")
            start_day, end_day = start[:8], end[:8]
            for path in paths:
                series = rone_request(key, {"STATBL_ID": tbl, "DTACYCLE_CD": cycle, "CLS_ID": path_to_id[path]})
                series = sorted(series, key=lambda r: str(r["WRTTIME_IDTFR_ID"]))
                prev = None
                for r in series:
                    day = re.sub(r"[^0-9]", "", str(r.get("WRTTIME_DESC") or ""))[:8]
                    value = r.get("DTA_VAL")
                    # 주간 공표 변동률 = 전주 대비 지수 변동률(소수 둘째 자리)
                    if prev not in (None, 0) and value is not None and len(day) == 8 and start_day <= day <= end_day:
                        rec = base_row(tbl, day, round((float(value) / float(prev) - 1) * 100, 2) + 0.0)
                        rec.update({"C1_OBJ_NM": "지역별", "C1_NM": str(path).split(">")[-1], "C1": f"r{path_to_id[path]}"})
                        out.append(rec)
                    prev = value
        elif kind == "land":
            region_code = self._get_kosis_region_code(5)
            for tbl, part, _ in spec["tables"]:
                probe = self._rone_probe(key, tbl, cycle, end)
                if part == "평균":
                    path_to_id = {r["CLS_FULLNM"]: r["CLS_ID"] for r in probe}
                else:
                    path_to_id = {r["GRP_FULLNM"]: r["GRP_ID"] for r in probe if r.get("GRP_FULLNM")}
                paths = self._rone_region_paths(list(path_to_id))
                paths = [p for p in paths if len(str(p).split(">")) == max(len(str(x).split(">")) for x in paths)] if paths else []
                if not paths:
                    raise RuntimeError(f"R-ONE 지역 매칭 실패: {cfg['label']} {tbl}")
                path = paths[0]
                filt = {"CLS_ID": path_to_id[path]} if part == "평균" else {"GRP_ID": path_to_id[path]}
                rows = rone_request(key, {"STATBL_ID": tbl, "DTACYCLE_CD": cycle, "START_WRTTIME": start[:6], "END_WRTTIME": end[:6], **filt})
                for r in rows:
                    if part == "평균" and "변동률" not in str(r.get("ITM_NM")):
                        continue
                    rec = base_row(tbl, str(r["WRTTIME_IDTFR_ID"]), r["DTA_VAL"])
                    rec.update({"C1_OBJ_NM": "시군구별", "C1_NM": str(path).split(">")[-1], "C1": region_code,
                                "C2_OBJ_NM": "용도지역및이용현황별", "C2_NM": "평균" if part == "평균" else r.get("CLS_NM"), "C2": f"r{r.get('CLS_ID')}"})
                    out.append(rec)
        df = pd.DataFrame(out)
        if not df.empty and cycle != "WK":
            df = df[(df["PRD_DE"].astype(str) >= start[:6]) & (df["PRD_DE"].astype(str) <= end[:6])].reset_index(drop=True)
        return df

    def _fetch_year_df(self, client: KosisClient, cfg: dict, yr: dict):
        if cfg.get("rone") and self._rone_api_key():
            try:
                rone_df = self._fetch_rone_df(cfg, yr)
                if not rone_df.empty:
                    self.log(f"[R-ONE] {cfg['label']} | {yr.get('start')}~{yr.get('end')} | {len(rone_df):,}건 (한국부동산원 R-ONE 직접 조회)")
                    return rone_df
                self.log(f"[R-ONE빈자료] {cfg['label']} | KOSIS로 조회")
            except Exception as exc:
                self.log(f"[R-ONE실패] {cfg['label']} | {exc} | KOSIS로 조회")
        prd_se = cfg.get("prdSe", "M").upper()

        if cfg.get("single_period_request") and prd_se in {"M", "D"}:
            params = {
                "orgId": cfg["orgId"],
                "tblId": cfg["tblId"],
                "prdSe": prd_se,
                "startPrdDe": yr["start"],
                "endPrdDe": yr["end"],
                **cfg.get("params", {}),
            }
            shown_params = {k: v for k, v in params.items() if k.startswith("objL") or k == "itmId"}
            self.log(f"[기간단건조회] {cfg['label']} | {yr['start']}~{yr['end']} | 파라미터={shown_params}")
            try:
                return self._request_with_region_variants(client, cfg, params, f"{cfg['label']} {yr['start']}~{yr['end']}")
            except EmptyKosisDataError:
                self.log(f"[데이터없음] {cfg['label']} | {yr['start']}~{yr['end']}")
                return pd.DataFrame()

        if prd_se == "M":
            frames = []
            empty_streak = 0
            for ym in month_points(yr["start"], yr["end"]):
                if self.stop_event.is_set():
                    self.log(f"[중단요청반영] {cfg['label']} | {yr['year']} | 월조회 중단")
                    break
                params = {
                    "orgId": cfg["orgId"],
                    "tblId": cfg["tblId"],
                    "prdSe": prd_se,
                    "startPrdDe": ym,
                    "endPrdDe": ym,
                    **cfg.get("params", {}),
                }
                self.log(f"[월조회] {cfg['label']} | {ym}")
                try:
                    if cfg.get("source_type") == "molit_unsold":
                        df_m = self._request_molit_unsold_month_df(cfg, ym)
                    elif cfg.get("custom_param_probe") == "household_sigungu":
                        df_m = self._request_household_sigungu_month_df(client, cfg, ym)
                    elif cfg.get("custom_param_probe") == "migration_sigungu":
                        df_m = self._request_migration_sigungu_month_df(client, cfg, ym)
                    elif cfg.get("custom_param_probe") == "one_person_household_age":
                        df_m = self._request_one_person_household_age_year_df(client, cfg, ym)
                    elif cfg.get("custom_param_probe") == "unsold_sigungu":
                        df_m = self._request_unsold_sigungu_month_df(client, cfg, ym)
                    elif cfg.get("auto_region_probe"):
                        df_m = self._request_housing_month_df(client, cfg, ym)
                    else:
                        df_m = self._request_with_region_variants(client, cfg, params, f"{cfg['label']} {ym}")
                    frames.append(df_m)
                    empty_streak = 0
                except EmptyKosisDataError:
                    self.log(f"[데이터없음] {cfg['label']} | {ym}")
                    empty_streak += 1
                    if empty_streak >= 2:
                        self.log(f"[중단] {cfg['label']} | {yr['year']} | 연속 {empty_streak}개월 데이터 없음")
                        break
            return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

        params = {
            "orgId": cfg["orgId"],
            "tblId": cfg["tblId"],
            "prdSe": prd_se,
            "startPrdDe": yr["start"],
            "endPrdDe": yr["end"],
            **cfg.get("params", {}),
        }
        self.log(f"[기간조회] {cfg['label']} | {yr['year']} | {yr['start']}~{yr['end']}")
        try:
            if cfg.get("custom_param_probe") == "one_person_household_age":
                return self._request_one_person_household_age_year_df(client, cfg, yr["year"])
            return client.request(params)
        except EmptyKosisDataError:
            self.log(f"[데이터없음] {cfg['label']} | {yr['year']}")
            return pd.DataFrame()

    def _molit_rtms_months(self, end_period: str, recent_months: int) -> list[str]:
        end_dt = parse_end_period(end_period)
        start_dt = end_dt - relativedelta(months=recent_months - 1)
        months = []
        current = start_dt
        while current <= end_dt:
            months.append(current.strftime("%Y%m"))
            current += relativedelta(months=1)
        return months

    def _parse_molit_rtms_xml(self, xml_text: str) -> tuple[list[dict], int]:
        try:
            root = ET.fromstring(xml_text.encode("utf-8") if isinstance(xml_text, str) else xml_text)
        except Exception as e:
            raise RuntimeError(f"국토부 실거래 XML 파싱 실패: {e} / 응답={str(xml_text)[:300]}")

        def _text(path: str, default: str = ""):
            elem = root.find(f".//{path}")
            return (elem.text or default).strip() if elem is not None else default

        result_code = _text("resultCode")
        result_msg = _text("resultMsg")
        if result_code and result_code not in {"00", "000"}:
            raise RuntimeError(f"국토부 실거래 API 오류: resultCode={result_code}, resultMsg={result_msg}")

        try:
            total_count = int(_text("totalCount", "0") or "0")
        except Exception:
            total_count = 0

        rows = []
        for item in root.findall(".//item"):
            row = {}
            for child in list(item):
                tag = child.tag.split("}")[-1]
                row[tag] = (child.text or "").strip()
            if row:
                rows.append(row)
        return rows, total_count

    def _request_molit_rtms_month_df(self, cfg: dict, lawd_cd: str, deal_ymd: str) -> pd.DataFrame:
        endpoint = str(cfg.get("endpoint", "")).strip().strip("/")
        if not endpoint:
            raise ValueError(f"{cfg['label']}의 국토부 실거래 endpoint가 비어 있습니다.")
        url = f"{RTMS_BASE_URL}/{endpoint}"
        all_rows = []
        page_no = 1
        num_rows = 1000
        total_count = None
        while True:
            if self.stop_event.is_set():
                self.log(f"[중단요청반영] {cfg['label']} | {deal_ymd} | 국토부 실거래 조회 중단")
                break
            params = {
                "serviceKey": RTMS_SERVICE_KEY,
                "LAWD_CD": lawd_cd,
                "DEAL_YMD": deal_ymd,
                "pageNo": page_no,
                "numOfRows": num_rows,
            }
            response = None
            rows = []
            total = 0
            last_error = None
            max_attempts = 4
            for attempt in range(1, max_attempts + 1):
                try:
                    self.log(
                        f"[국토부실거래조회] {cfg['label']} | LAWD_CD={lawd_cd} "
                        f"DEAL_YMD={deal_ymd} page={page_no} attempt={attempt}/{max_attempts}"
                    )
                    response = requests.get(url, params=params, timeout=DEFAULT_TIMEOUT)
                    response.raise_for_status()
                    response.encoding = response.apparent_encoding or "utf-8"
                    rows, total = self._parse_molit_rtms_xml(response.text)
                    last_error = None
                    break
                except Exception as exc:
                    last_error = exc
                    self.log(
                        f"[국토부실거래재시도] {cfg['label']} | {lawd_cd} | {deal_ymd} | "
                        f"page={page_no} | attempt={attempt}/{max_attempts}"
                    )
                    if attempt < max_attempts:
                        time.sleep(min(1.5 * attempt, 5.0))
            if last_error is not None:
                raise RuntimeError(
                    f"국토부 실거래 월별 조회 {max_attempts}회 실패: "
                    f"LAWD_CD={lawd_cd}, DEAL_YMD={deal_ymd}, page={page_no}, 오류={last_error}"
                )
            total_count = total if total_count is None else total_count
            all_rows.extend(rows)
            if not rows:
                break
            if total_count is not None and len(all_rows) >= total_count:
                break
            if len(rows) < num_rows:
                break
            page_no += 1
            if page_no > 100:
                self.log(f"[페이지제한] {cfg['label']} | {deal_ymd} | 100페이지 초과로 중단")
                break
            time.sleep(0.2)
        df = pd.DataFrame(all_rows)
        if not df.empty:
            df.insert(0, "자료구분", cfg["label"])
            df.insert(1, "LAWD_CD", lawd_cd)
            df.insert(2, "DEAL_YMD", deal_ymd)
        return df

    def _request_molit_rtms_recent_df(self, cfg: dict, end_period: str) -> tuple[pd.DataFrame, str]:
        """국토교통부 실거래 최근 N개월 수집.

        사용 안정성을 위해 월별 요청을 순차 처리합니다.
        월별 오류·빈 응답은 재확인하고, 최종 실패월이 남으면 불완전 파일 저장을 차단합니다.
        """
        lawd_cd = self._get_selected_region_code(5)
        if not lawd_cd:
            raise ValueError(f"{cfg['label']}은 시도/시군구 선택 후 REG 5자리 코드가 필요합니다.")

        recent_n = int(cfg.get("recent_months", 36))
        months = self._molit_rtms_months(end_period, recent_n)
        suffix = f"{months[0][:4]}_{months[0][4:]}_{months[-1][:4]}_{months[-1][4:]}"
        frames_by_month: dict[str, pd.DataFrame] = {}
        failed_months: list[tuple[str, str]] = []
        confirmed_empty_months: list[str] = []

        self.log(f"[국토부실거래수집] {cfg['label']} | LAWD_CD={lawd_cd} | 최근 {recent_n}개월 | 순차실행")

        for ym in months:
            if self.stop_event.is_set():
                failed_months.append((ym, "중단 요청"))
                self.log(f"[국토부실거래중단] {cfg['label']} | {lawd_cd} | {ym}")
                break
            try:
                df_m = self._request_molit_rtms_month_df(cfg, lawd_cd, ym)
            except Exception as e:
                failed_months.append((ym, self._sanitize_log_text(str(e))))
                self.log(f"[국토부실거래오류-월건너뜀] {cfg['label']} | {lawd_cd} | {ym} | 오류내용은 실패목록 확인")
                continue

            if df_m.empty:
                # 일시적으로 totalCount=0이 반환되는 경우가 있어 한 번 더 확인한 뒤 빈 월로 확정한다.
                self.log(f"[국토부실거래빈응답재확인] {cfg['label']} | {lawd_cd} | {ym}")
                time.sleep(0.8)
                try:
                    df_confirm = self._request_molit_rtms_month_df(cfg, lawd_cd, ym)
                except Exception as e:
                    failed_months.append((ym, self._sanitize_log_text(str(e))))
                    continue
                if not df_confirm.empty:
                    df_m = df_confirm
                    self.log(f"[국토부실거래빈응답복구] {cfg['label']} | {lawd_cd} | {ym} | {len(df_m):,}건")
                else:
                    confirmed_empty_months.append(ym)
                    self.log(f"[국토부실거래데이터없음-재확인완료] {cfg['label']} | {lawd_cd} | {ym}")
                    continue

            if not df_m.empty:
                frames_by_month[ym] = df_m
                self.log(f"[국토부실거래조회완료] {cfg['label']} | {lawd_cd} | {ym} | {len(df_m):,}건")

        if failed_months:
            self.log(f"[국토부실거래실패월최종재수집] {cfg['label']} | 대상={','.join(ym for ym, _ in failed_months)}")
            final_failed: list[tuple[str, str]] = []
            for ym, _reason in failed_months:
                if self.stop_event.is_set():
                    final_failed.append((ym, "중단 요청"))
                    continue
                try:
                    time.sleep(1.0)
                    df_retry = self._request_molit_rtms_month_df(cfg, lawd_cd, ym)
                    if df_retry.empty:
                        time.sleep(0.8)
                        df_retry = self._request_molit_rtms_month_df(cfg, lawd_cd, ym)
                    if df_retry.empty:
                        confirmed_empty_months.append(ym)
                        self.log(f"[국토부실거래실패월빈자료확정] {cfg['label']} | {lawd_cd} | {ym}")
                    else:
                        frames_by_month[ym] = df_retry
                        self.log(f"[국토부실거래실패월복구] {cfg['label']} | {lawd_cd} | {ym} | {len(df_retry):,}건")
                except Exception as exc:
                    final_failed.append((ym, self._sanitize_log_text(str(exc))))
            if final_failed:
                detail = ", ".join(f"{ym}({reason})" for ym, reason in final_failed[:5])
                raise RuntimeError(
                    f"국토부 실거래 수집 실패월이 남아 불완전 파일 저장을 차단했습니다: {detail}"
                )

        collected_or_empty = set(frames_by_month) | set(confirmed_empty_months)
        unaccounted = [ym for ym in months if ym not in collected_or_empty]
        if unaccounted:
            raise RuntimeError(
                "국토부 실거래 월별 완전성 검증 실패로 파일 저장을 차단했습니다: "
                + ", ".join(unaccounted)
            )

        ordered_frames = [frames_by_month[ym] for ym in months if ym in frames_by_month]
        if not ordered_frames:
            return pd.DataFrame(), suffix
        return pd.concat(ordered_frames, ignore_index=True), suffix

    def _run_one_cfg(self, client: KosisClient, output_dir: Path, cfg: dict, end_period: str):


        if cfg.get("source_type") == "molit_rtms":
            df, suffix = self._request_molit_rtms_recent_df(cfg, end_period)
            self.log(f"[조회완료] {cfg['label']} | 최근 {cfg.get('recent_months', 36)}개월 | {len(df):,}건")
            if df.empty:
                self.log(f"[빈파일저장차단] {cfg['label']} | 전체 조회 구간 | 데이터 없음")
                return []
            df = self._normalize_df(df)
            base_filename = cfg["filename"].format(suffix=suffix, year=parse_end_period(end_period).year)
            filename = self._build_output_filename(base_filename)
            saved_paths = []
            for group_folder_name in self._get_group_folders_for_cfg(cfg):
                folder = ensure_dir(output_dir / group_folder_name)
                path = self._save_single_stat_excel(folder, filename, df, cfg["label"])
                saved_paths.append(path)
            return saved_paths

        if not cfg.get("orgId"):
            raise ValueError("orgId가 비어 있습니다.")
        if not cfg.get("tblId"):
            raise ValueError("tblId가 비어 있습니다. KOSIS 표 코드를 먼저 입력해 검증하세요.")
        if not cfg.get("label"):
            raise ValueError("통계명이 비어 있습니다.")

        cfg = self._apply_region_filter(cfg)
        prd_se = cfg.get("prdSe", "M")
        lookback_years = int(cfg.get("lookback_years", 3))
        end_dt = parse_end_period(end_period)

        # 월별 누계형 전국 통계는 연도 전체 반복 조회가 아니라
        # 해당년도 포함 최근 6년 기준으로 과거 연도는 매년 12월, 당해연도는 선택 종료월 누계만 단건 조회
        if cfg.get("cumulative_point_request"):
            df, suffix = self._request_cumulative_point_df(client, cfg, end_period)
            self.log(f"[조회완료] {cfg['label']} | 최근 6년 연말/당해월 누계 기준월 | {len(df):,}건")
            if df.empty:
                self.log(f"[빈파일저장차단] {cfg['label']} | 누계 기준월 | 데이터 없음")
                return []
            df = self._normalize_df(df)
            base_filename = cfg["filename"].format(year=end_dt.year, suffix=suffix)
            filename = self._build_output_filename(base_filename)
            saved_paths = []
            for group_folder_name in self._get_group_folders_for_cfg(cfg):
                folder = ensure_dir(output_dir / group_folder_name)
                path = self._save_single_stat_excel(folder, filename, df, cfg["label"])
                saved_paths.append(path)
            return saved_paths

        # 특정 통계는 연도별 3년 조회가 아니라, 종료월 기준 최근 N개월만 수집
        # 예: 오피스텔/주택동향 월간 가격지수 최근 13개월
        if prd_se.upper() == "M" and cfg.get("recent_months"):
            recent_n = int(cfg.get("recent_months"))
            start_dt = end_dt - relativedelta(months=recent_n - 1)
            year_ranges = [{
                "year": f"{start_dt.year}-{end_dt.year}",
                "start": start_dt.strftime("%Y%m"),
                "end": end_dt.strftime("%Y%m"),
                "suffix": f"{start_dt.strftime('%Y_%m')}_{end_dt.strftime('%Y_%m')}",
            }]
        elif prd_se.upper() == "D" and cfg.get("recent_months"):
            # 주간아파트 지수는 KOSIS에서 prdSe=D, 공표 기준일(매주 월요일) YYYYMMDD 기준으로 조회합니다.
            # 입력값이 YYYY-MM이면 해당 월의 마지막 월요일을 기준으로 하고, 현재 월이면 오늘 기준 직전/당일 월요일을 사용합니다.
            # 입력값이 YYYY-MM-DD 또는 YYYYMMDD이면 그 날짜 기준 직전/당일 월요일을 사용합니다.
            recent_n = int(cfg.get("recent_months"))
            today = datetime.now()
            raw_period = str(end_period or "").strip()
            explicit_day = bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}|\d{8}", raw_period))
            if explicit_day:
                anchor = parse_end_period(raw_period)
            elif end_dt.year == today.year and end_dt.month == today.month:
                anchor = today
            else:
                anchor = end_dt + relativedelta(months=1) - relativedelta(days=1)
            if anchor > today:
                anchor = today
            end_date = anchor - relativedelta(days=anchor.weekday())  # 공표 기준일: 월요일
            start_month = end_date.replace(day=1) - relativedelta(months=recent_n - 1)
            start_date = start_month + relativedelta(days=(7 - start_month.weekday()) % 7)  # 시작월 첫 월요일
            year_ranges = [{
                "year": f"{start_date.year}-{end_date.year}",
                "start": start_date.strftime("%Y%m%d"),
                "end": end_date.strftime("%Y%m%d"),
                "suffix": f"{start_date.strftime('%Y_%m_%d')}_{end_date.strftime('%Y_%m_%d')}",
                "weekly_basis": "MONDAY",
            }]
            self.log(f"[주간기간계산] {cfg['label']} | 매주 월요일 기준 최근 {recent_n}개월 | 기준월/일={raw_period or '현재월'} | {year_ranges[0]['start']}~{year_ranges[0]['end']}")
        else:
            year_ranges = build_year_ranges(end_period, lookback_years, prd_se)

        group_folder_names = self._get_group_folders_for_cfg(cfg)
        saved_paths = []
        collected_frames = []
        suffix_parts = []

        for yr in year_ranges:
            if self.stop_event.is_set():
                self.log(f"[중단요청반영] {cfg['label']} | 남은 연도 건너뜀")
                break
            self.log(f"[시작] {cfg['label']} | {yr['year']} | orgId={cfg['orgId']} tblId={cfg['tblId']}")
            df = self._fetch_year_df(client, cfg, yr)
            self.log(f"[조회완료] {cfg['label']} | {yr['year']} | {len(df):,}건")
            if not df.empty:
                self.log(f"[컬럼] {list(df.columns)}")

            if cfg.get("prdSe", "").upper() == "W" and not df.empty:
                target_year = end_dt.year
                target_month = end_dt.month
                try:
                    current_year = int(str(yr["year"])[:4])
                except Exception:
                    current_year = target_year
                if current_year == target_year and "PRD_DE" in df.columns:
                    ym_pairs = df["PRD_DE"].apply(extract_year_month_from_prd_de)
                    valid_mask = []
                    parsed_count = 0
                    for pair in ym_pairs:
                        if pair is None:
                            valid_mask.append(True)
                        else:
                            parsed_count += 1
                            y, m = pair
                            valid_mask.append((y < target_year) or (y == target_year and m <= target_month))
                    before_cnt = len(df)
                    df = df[pd.Series(valid_mask, index=df.index)]
                    self.log(f"[주간필터] {cfg['label']} | 종료기준월 {target_year}-{target_month:02d}까지 반영 | {before_cnt:,}->{len(df):,}건 | 파싱성공 {parsed_count:,}건")

            if df.empty:
                self.log(f"[빈데이터] {cfg['label']} | {yr['year']} | 조사하지 않는 지역이거나 데이터 없음")
                continue
            from_rone = "출처" in df.columns and df["출처"].astype(str).str.contains("R-ONE", na=False).any()
            if cfg.get("post_filter_region") and not df.empty and not from_rone:
                df = self._filter_housing_sido_df(df, cfg)
            if cfg.get("weekly_apt_region_post_filter") and not df.empty and not from_rone:
                df = self._filter_weekly_apt_region_df(df, cfg)
            df = self._normalize_df(df)
            if not df.empty:
                collected_frames.append(df)
                suffix_parts.append(yr.get("suffix", yr["year"]))

        if not collected_frames:
            self.log(f"[빈파일저장차단] {cfg['label']} | 전체 조회 구간 | 데이터 없음")
            return []

        combined_df = pd.concat(collected_frames, ignore_index=True)
        if len(suffix_parts) == 1:
            combined_suffix = suffix_parts[0]
        else:
            combined_suffix = f"{suffix_parts[0]}_{suffix_parts[-1]}"
        base_filename = cfg["filename"].format(year=end_dt.year, suffix=combined_suffix)
        filename = self._build_output_filename(base_filename)

        for group_folder_name in group_folder_names:
            folder = ensure_dir(output_dir / group_folder_name)
            path = self._save_single_stat_excel(folder, filename, combined_df, cfg["label"])
            saved_paths.append(path)

        return saved_paths


    def _set_running_state(self, running: bool):
        self.is_running = running
        if self.run_button is not None:
            self.run_button.configure(state="disabled" if running else "normal")
        if self.stop_button is not None:
            self.stop_button.configure(state="normal" if running else "disabled")

    def _stop_run(self):
        if not self.is_running:
            return
        self.stop_event.set()
        self.status_var.set("중단 요청됨")
        self.log("[중단요청] 현재 작업을 마친 후 다음 항목부터 중단합니다.")


    def _select_group_tab(self, group_name: str):
        notebook = getattr(self, "group_notebook", None)
        if not notebook or not group_name:
            return
        for page_id, name in self.group_pages.items():
            if name == group_name:
                try:
                    notebook.select(page_id)
                except Exception:
                    pass
                return

    def _finish_collection(self, group_name, saved_count, failed_items):
        """수집 정상 종료 후 완료 안내 또는 보고서 자동 진행."""
        fail_count = len(failed_items)
        preview = "\n".join([f"- {name}: {msg[:120]}" for name, msg in failed_items[:8]])
        if fail_count > 8:
            preview += f"\n... 외 {fail_count - 8}건"
        auto = bool(self.auto_pipeline_var.get()) and callable(self.on_collection_finished)
        if not auto:
            if fail_count:
                messagebox.showwarning("완료 - 실패 항목 있음", f"저장 파일 수: {saved_count}개\n실패 통계: {fail_count}건\n\n{preview}")
            else:
                messagebox.showinfo("완료", f"저장 파일 수: {saved_count}")
            return
        if fail_count and not messagebox.askyesno(
            "완료 - 실패 항목 있음",
            f"저장 파일 수: {saved_count}개\n실패 통계: {fail_count}건\n\n{preview}\n\n실패 항목을 제외하고 보고서 자동 진행을 계속할까요?",
        ):
            return
        self._select_group_tab(group_name)
        try:
            self.on_collection_finished(group_name)
        except Exception as e:
            messagebox.showerror("자동 진행 오류", str(e))

    def _run_selected(self):
        if self.is_running:
            messagebox.showinfo("안내", "이미 실행 중입니다.")
            return

        validated = self._validate_common_inputs()
        if not validated:
            return

        api_key, output_dir, end_period = validated

        # 선택 통계 실행 버튼 클릭 시 기존 선택통계실행 폴더를 무조건 삭제한다.
        # 다른 지역·다른 기준월 파일이 남아 있으면 LLM 보고서 생성 시 혼입될 수 있으므로
        # API 원자료 재수집 전 항상 비운다.
        try:
            self._clear_simple_run_folder(Path(output_dir))
        except Exception as e:
            messagebox.showerror("기존자료 삭제 오류", f"선택통계실행 폴더 삭제 중 오류가 발생했습니다.\n\n{e}")
            return

        current_group_name = self._get_current_group_name()
        if not current_group_name:
            messagebox.showerror("오류", "실행할 업무 탭을 선택할 수 없습니다.")
            return
        sido_total_mode = self._sido_report_enabled(current_group_name)
        if sido_total_mode and not self._get_sido_sigungu_rows():
            messagebox.showerror("오류", "시도 보고서 생성을 위해 법정동 코드 파일과 시도 선택이 필요합니다.")
            return

        current_group_keys = self.group_key_map.get(current_group_name, GROUPS.get(current_group_name, []))
        selected_keys = [
            k for k in current_group_keys
            if self.dataset_checks.get(k) is not None and self.dataset_checks[k].get()
        ]
        if not selected_keys:
            messagebox.showerror("오류", "현재 업무 탭에서 선택된 통계가 없습니다.")
            return

        hold_selected = [k for k in selected_keys if k in HOLD_DATASET_KEYS or DATASETS[k].get("status") == "hold"]
        for k in hold_selected:
            self.log(f"[매핑필요] {DATASETS[k]['label']} | KOSIS OpenAPI 직접수집 제외")
        selected_keys = [k for k in selected_keys if k not in hold_selected]
        if not selected_keys:
            messagebox.showerror("오류", "실행 가능한 통계가 없습니다. 보류 항목은 제외되었습니다.")
            return

        cfgs = []
        for k in selected_keys:
            cfg = deepcopy(DATASETS[k])
            cfg["_key"] = k
            cfgs.append(cfg)
        self.progress["value"] = 0
        if hasattr(self, "inline_progress"):
            self.inline_progress["value"] = 0
        self.api_progress_percent_var.set("0%")
        self._reset_dataset_visual_status()
        self.stop_event.clear()
        self._set_running_state(True)
        self.active_run_group_name = current_group_name
        self.status_var.set("선택 통계 실행 중")
        self.log(f"선택 통계 실행 시작: {len(cfgs)}건")
        try:
            self._show_api_collect_page()
            first_key = cfgs[0].get("_key") if cfgs else None
            if first_key:
                self._scroll_to_dataset(first_key)
        except Exception:
            pass

        finished_info = {}

        def worker():
            total = len(cfgs)
            done = 0
            saved_count = 0
            failed_items = []
            stopped = False
            max_workers = max(1, min(5, int(self.worker_var.get() or 1)))

            def run_one_dataset(cfg: dict):
                if self.stop_event.is_set():
                    return cfg, [], "STOPPED"
                client = KosisClient(api_key)
                key = cfg.get("_key")
                self.after(0, self._set_dataset_visual_status, key, "running")
                self.after(0, lambda name=cfg.get("label", ""): self.status_var.set(f"추출 중: {name}"))
                self.log(f"[진행중] {cfg['label']} 추출 시작")
                if sido_total_mode:
                    paths = self._run_one_cfg_sido_total(client, Path(output_dir), cfg, end_period)
                else:
                    paths = self._run_one_cfg(client, Path(output_dir), cfg, end_period)
                return cfg, paths, None

            try:
                self.log(f"[동시수집설정] 자료별 동시 수집 수={max_workers} | 시도단위취합={'자료별 병렬·자료 내부 시군구 순차' if sido_total_mode else '병렬'}")
                if max_workers <= 1:
                    for cfg in cfgs:
                        if self.stop_event.is_set():
                            stopped = True
                            self.log(f"[중단반영] 남은 통계 {total - done}건은 실행하지 않습니다.")
                            break
                        key = cfg.get("_key")
                        try:
                            _cfg, paths, marker = run_one_dataset(cfg)
                            if marker == "STOPPED":
                                stopped = True
                                break
                            saved_count += len(paths)
                            self.after(0, self._set_dataset_visual_status, key, "done")
                            self.log(f"[완료] {cfg['label']} | 저장 파일 {len(paths)}개")
                        except Exception as e:
                            failed_items.append((cfg.get("label", key), str(e)))
                            self.after(0, self._set_dataset_visual_status, key, "failed", str(e))
                            self.log(f"[실패] {cfg['label']}: {e}")
                        finally:
                            done += 1
                            self.after(0, self._update_progress, done, total)
                else:
                    with ThreadPoolExecutor(max_workers=max_workers) as executor:
                        future_map = {executor.submit(run_one_dataset, cfg): cfg for cfg in cfgs}
                        for future in as_completed(future_map):
                            cfg = future_map[future]
                            key = cfg.get("_key")
                            if self.stop_event.is_set():
                                stopped = True
                            try:
                                _cfg, paths, marker = future.result()
                                if marker == "STOPPED":
                                    stopped = True
                                    self.after(0, self._set_dataset_visual_status, key, "skipped")
                                else:
                                    saved_count += len(paths)
                                    self.after(0, self._set_dataset_visual_status, key, "done")
                                    self.log(f"[완료] {cfg['label']} | 저장 파일 {len(paths)}개")
                            except Exception as e:
                                failed_items.append((cfg.get("label", key), str(e)))
                                self.after(0, self._set_dataset_visual_status, key, "failed", str(e))
                                self.log(f"[실패] {cfg['label']}: {e}")
                            finally:
                                done += 1
                                self.after(0, self._update_progress, done, total)
                            if stopped:
                                for pending in future_map:
                                    pending.cancel()

                final_done = done
                if stopped or self.stop_event.is_set():
                    self.log(f"선택 통계 실행 중단: 완료 {final_done}건 / 저장 파일 {saved_count}개")
                    self.after(0, lambda: self.status_var.set(f"중단됨 ({final_done}/{total})"))
                    self.after(0, lambda: messagebox.showinfo("중단", f"완료 통계: {final_done}건\n저장 파일 수: {saved_count}"))
                else:
                    fail_count = len(failed_items)
                    self.log(f"선택 통계 실행 완료: 저장 파일 {saved_count}개 / 실패 {fail_count}건")
                    if fail_count:
                        self.after(0, lambda: self.status_var.set(f"완료(실패 {fail_count}건 확인 필요)"))
                    else:
                        self.after(0, lambda: self.status_var.set("선택 통계 실행 완료"))
                    finished_info["result"] = (current_group_name, saved_count, list(failed_items))
            except Exception as e:
                self.log(f"[치명적 오류] {e}")
                self.after(0, lambda: messagebox.showerror("오류", str(e)))
                self.after(0, lambda: self.status_var.set("오류"))
            finally:
                self.after(0, lambda: setattr(self, "active_run_group_name", None))
                self.after(0, lambda: self._set_running_state(False))
                if finished_info.get("result"):
                    self.after(50, lambda: self._finish_collection(*finished_info["result"]))

        self.run_thread = threading.Thread(target=worker, daemon=True)
        self.run_thread.start()


if __name__ == "__main__":
    app = App()
    app.mainloop()

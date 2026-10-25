# -*- coding: utf-8 -*-
"""
Land price monthly market report helper.

Creates the LLM analysis package used before writing the final monthly
land-price/land-market HTML report. Numeric calculations are kept in Python;
the LLM receives the resulting JSON and writes only narrative interpretation.
"""

from __future__ import annotations

import argparse
import calendar
import html
import json
import math
import os
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any

import pandas as pd


EOK = 100_000_000


# 반도체 첨단국가산단(광주 군공항 예정지) 영향권 라벨.
# 2026-08 조사 기준으로 송정·도산·신촌·소촌·우산·황룡동을 '반도체 직접 영향권'으로 통합 라벨링한다.
SEMICONDUCTOR_INFLUENCE_DONGS = {
    "송정동": ("direct", "반도체 직접 영향권"),
    "도산동": ("direct", "반도체 직접 영향권"),
    "신촌동": ("direct", "반도체 직접 영향권"),
    "소촌동": ("direct", "반도체 직접 영향권"),
    "우산동": ("direct", "반도체 직접 영향권"),
    "황룡동": ("direct", "반도체 직접 영향권"),
    "선암동": ("linked", "연계 영향권"),
    "운수동": ("linked", "연계 영향권"),
}
SEMICONDUCTOR_LEVEL_LABELS = {
    "direct": "반도체 직접 영향권",
    "linked": "연계 영향권",
    # 과거 산출물 호환용(현재 분류에서는 사용하지 않음)
    "core": "핵심 영향권",
    "adjacent": "인접 영향권",
}
SEMICONDUCTOR_LEVEL_ORDER = ["direct", "core", "adjacent", "linked"]


def semiconductor_dongs(level: str) -> list[str]:
    return [dong for dong, (lvl, _) in SEMICONDUCTOR_INFLUENCE_DONGS.items() if lvl == level]


# 토지거래허가구역 지정(계약일 기준 전·후 비교 기준일).
LAND_PERMIT_ZONE_RULES = {
    "광산구": {
        "effective_date": "2026-07-14",
        "designation_period": "2026-07-14 ~ 2028-07-13",
        "announcement": "국토교통부 공고 제2026-940호",
        "scope": "광주 군공항 호남 반도체 첨단국가산단 예정지 일원(약 364.19㎢)",
    },
}

# 주간 집계 기준: 매주 금요일 작성 → 토요일 시작 ~ 금요일 종료의 7일 창.
WEEK_END_WEEKDAY = 4  # Monday=0 … Friday=4
WEEKLY_WINDOW_COUNT = 13


GWANGSAN_REGIONAL_ZONES = [
    {
        "zone": "송정역·도심전환권",
        "dongs": ["송정동", "송정1동", "송정2동", "도산동", "신흥동", "어룡동", "도호동", "신촌동", "소촌동", "황룡동", "선암동", "운수동"],
        "characteristics": "광주송정역 역세권과 기존 도심, 군공항 부지가 함께 접한 권역으로, 호남권 반도체 첨단국가산단 예정지 인접성이 가장 크게 작용하는 구역이며 송정동·도산동·신촌동·소촌동·황룡동이 반도체 직접 영향권에 해당함",
        "demand_background": "역세권 생활·상업 수요와 기존 시가지 정비 수요에 더해, 반도체 국가산단 예정지·군공항 이전과 연계된 개발 기대 및 토지거래허가구역 규제가 함께 작용함",
        "land_use": "송정역세권의 상업·주거지, 노후 단독주택이 밀집한 기존 시가지, 광주공항·군공항 주변 토지와 황룡강변 생산녹지, 소촌농공단지가 혼재함",
        "rear_demand": "광주송정역 이용객과 원도심 생활수요, 공항·소촌농공단지 종사자, 산단 조성 기대에 따른 관외 매수 문의",
        "transport": "광주송정역·광주공항과 상무대로 등 광역·도심 교통축, 군공항 이전 시 재편이 예상되는 접근축",
        "commercial_area": "송정역세권 원도심 상권과 소촌·도산 생활상권",
        "field_checks": "역세권 정비 진행상황, 상가 공실·유동인구, 토지거래허가 신청·처리 실적, 매도 호가 상향과 매물 회수 여부, 국가산단 지정·군공항 이전의 추진단계",
        "analysis_caution": "역세권 실수요와 반도체 국가산단·군공항 이전 기대를 구분하고, 직접 영향권 표기나 사업 추진단계가 확인되지 않은 기대를 지가 상승의 직접 원인으로 단정하지 않음",
        "industrial_links": ["호남권 반도체 첨단국가산업단지(광주 군공항 부지 후보지)", "소촌농공단지"],
    },
    {
        "zone": "수완·신가 생활권",
        "dongs": ["수완동", "신가동", "신창동", "운남동", "월곡동", "월곡1동", "월곡2동", "우산동", "장덕동", "흑석동"],
        "characteristics": "대규모 공동주택과 상업·생활편의시설이 집적된 광산구의 대표 주거·생활권으로, 권역 내 우산동은 군공항 예정지에 접해 반도체 직접 영향권에 해당함",
        "demand_background": "실거주 기반 주거 수요와 생활형 상업·업무용 토지 수요",
        "land_use": "수완지구의 고밀 주거·상업용지와 장덕·흑석동의 생활편의시설, 외곽 자연녹지 토지가 혼재함",
        "rear_demand": "대규모 공동주택 거주인구, 교육·소매·의료 이용수요와 인접 하남산단 종사자 배후수요",
        "transport": "임방울대로·장신로 등 수완·하남 생활권 간선도로",
        "commercial_area": "수완·장덕 중심상권과 대형 생활편의시설 중심의 광역 생활상권",
        "field_checks": "성숙 상권의 공실·임대수요, 자연녹지 고가거래의 개별성, 주거지와 산업용지 수요의 구분, 우산동 생산녹지의 허가구역 지정 전후 계약 변화",
        "analysis_caution": "성숙 주거지의 소규모 거래와 외곽 개발용 토지 거래를 분리하고 개별 고가거래를 생활권 전반의 상승으로 일반화하지 않음",
        "industrial_links": ["하남일반산업단지"],
    },
    {
        "zone": "첨단·비아 혁신권",
        "dongs": ["첨단1동", "첨단2동", "비아동", "쌍암동", "월계동", "산월동", "도천동"],
        "characteristics": "첨단과학산업단지의 광산업·연구개발 기능과 첨단3지구의 AI 융복합 개발축이 연계되는 혁신산업 권역",
        "demand_background": "연구개발·기업 활동에 따른 업무·상업 수요와 신규 개발 기대",
        "land_use": "첨단과학산단의 연구·산업시설, 첨단지구 상업·업무·배후주거지와 비아·도천동의 개발가능지가 혼재함",
        "rear_demand": "첨단산단 기업·연구기관 종사자, 배후주거 인구와 업무·상업시설 이용수요",
        "transport": "호남고속도로와 첨단·비아권 주요 간선도로 접근성",
        "commercial_area": "첨단지구 상업·업무 중심지와 배후 공동주택 생활상권",
        "field_checks": "첨단3지구의 고시·조성·분양 단계, 기업 입주·고용, 상업시설 공실과 실제 배후수요",
        "analysis_caution": "운영 중인 첨단과학산단과 조성·개발 단계인 첨단3지구를 구분하고 기업 입주·고용이 확인되지 않은 기대수요를 실수요로 간주하지 않음",
        "industrial_links": ["첨단과학국가산업단지", "첨단3지구 AI융복합지구"],
    },
    {
        "zone": "하남·진곡 제조산업권",
        "dongs": ["하남동", "임곡동", "안청동", "진곡동", "고룡동", "오선동"],
        "characteristics": "하남·진곡 산업단지를 중심으로 제조업, 뿌리산업, 자동차 부품 및 산업물류 기능이 집적된 권역",
        "demand_background": "공장·창고·지원시설용 토지 수요와 기존 산단 재생·스마트화 수요",
        "land_use": "하남·진곡산단의 공장·창고·지원시설과 산단 배후 주거·상업지를 중심으로 형성됨",
        "rear_demand": "제조·자동차부품·뿌리산업 기업의 공장·물류 수요와 산단 종사자 배후수요",
        "transport": "산단 내부 간선도로와 무진대로·호남고속도로 등 광역 물류망 연계",
        "commercial_area": "하남산단 지원상권과 하남·흑석 생활권의 근린상권",
        "field_checks": "기업 가동률·입주·증설, 공장 공실과 유휴부지, 고액 산업용지 거래의 대표성",
        "analysis_caution": "산업용지 거래는 기업 투자·가동률·공실·업종별 경기와 함께 검토하고 산단 인접만으로 가격 영향을 단정하지 않음",
        "industrial_links": ["하남일반산업단지", "진곡일반산업단지"],
    },
    {
        "zone": "본량·삼도 서부 농촌권",
        "dongs": ["유계동", "신룡동", "두정동", "광산동", "송촌동", "지죽동", "요기동", "복룡동", "송대동",
                   "하산동", "본덕동", "용봉동", "장수동", "산정동", "송치동", "산수동", "선동", "지산동",
                   "내산동", "대산동", "송학동", "남산동", "왕동", "북산동", "명도동", "지정동", "명화동", "동산동"],
        "characteristics": "광산구 서부의 본량·삼도 일원으로 농경지와 임야, 취락이 넓게 분포하는 농촌권역이며 대면적 저가 토지가 많아 개발·투자 목적 매수가 나타나기 쉬운 지역",
        "demand_background": "영농 목적 실수요와 대면적 토지의 투자·개발 목적 매수, 산정 공공주택지구 등 인접 개발사업에 따른 기대 수요",
        "land_use": "전·답·과수원과 임야, 생산·보전관리지역 토지가 주를 이루고 취락지와 소규모 공장·창고가 산재함",
        "rear_demand": "영농 종사자와 인근 산단·물류시설 종사자, 대면적 토지를 찾는 관외 매수 수요",
        "transport": "광주광역시 서부 외곽을 잇는 지방도와 광주순환도로 접근축",
        "commercial_area": "면 소재지 중심의 소규모 생활상권",
        "field_checks": "농지취득자격증명 등 이용 규제, 개발행위허가 가능 여부, 도로 접합과 실제 이용 상태, 대면적 토지의 매수 목적",
        "analysis_caution": "개별 대면적 거래가 권역 평균 단가를 크게 흔들 수 있으므로 소수 거래를 지역 전반의 가격 흐름으로 일반화하지 않고, 농지 규제와 실제 이용 가능성을 함께 확인함",
        "industrial_links": [],
    },
    {
        "zone": "평동·빛그린 산업확장권",
        "dongs": ["평동", "동곡동", "삼도동", "본량동", "옥동", "월전동", "장록동", "용동", "연산동", "삼거동", "덕림동", "동호동", "양동", "용곡동"],
        "characteristics": "평동산단과 빛그린국가산단을 축으로 금형·제조·미래차 기능이 확장되는 한편 농촌·녹지 토지가 넓게 분포하는 도농복합 권역",
        "demand_background": "산업시설·물류·협력업체 입지 수요와 산업축 주변 장기 개발 기대",
        "land_use": "평동산단의 공장·물류시설과 빛그린 산업축, 주변 농지·녹지·취락이 함께 분포함",
        "rear_demand": "평동·빛그린 입주기업과 협력업체의 산업용지·물류 수요, 산단 종사자 및 농촌 생활수요",
        "transport": "광주도시철도 1호선 평동역과 산단·광역 물류도로망",
        "commercial_area": "평동산단 지원시설과 주변 근린생활형 상권",
        "field_checks": "산단 입주율·기반시설, 농지·녹지 규제와 실제 전용 가능성, 산업용지와 농지 거래의 분리",
        "analysis_caution": "산단 내부·인접 토지와 농지·녹지 거래를 분리하고 입주율, 기반시설, 농지규제와 실제 이용 가능성을 함께 확인함",
        "industrial_links": ["평동일반산업단지", "빛그린국가산업단지"],
    },
]


GWANGSAN_INDUSTRIAL_COMPLEXES = [
    {"name": "하남일반산업단지", "classification": "기존 산업단지", "linked_zone": "하남·진곡 제조산업권", "characteristics": "뿌리산업과 제조·자동차부품 기업이 집적된 기존 산단으로 산단 재생과 스마트공장 연계가 주요 점검 요소", "market_points": "공장·창고 거래, 기업 가동률, 공실, 재생사업 진행상황"},
    {"name": "진곡일반산업단지", "classification": "기존 산업단지", "linked_zone": "하남·진곡 제조산업권", "characteristics": "하남과 빛그린 산업축 사이에서 제조·자동차부품 및 산업물류 연계 가능성을 살필 산단", "market_points": "기업 입주·증설, 산업용지 실수요, 물류 접근성, 미분양·유휴부지"},
    {"name": "평동일반산업단지(1·2·3차)", "classification": "기존 산업단지", "linked_zone": "평동·빛그린 산업확장권", "characteristics": "평동2차 금형특화단지를 포함해 금형·제조업과 시험생산 지원기능이 집적된 산업권", "market_points": "금형·제조기업 투자, 공장·지원시설 거래, 산업물류 및 기반시설 수요"},
    {"name": "소촌농공단지", "classification": "기존 산업단지", "linked_zone": "송정역·도심전환권", "characteristics": "기존 자동차부품 기업 기반을 미래차 산업과 연계해 살필 소규모 산업거점", "market_points": "기존 공장 재편, 부품기업 투자, 주변 주거지와의 용도·환경 제약"},
    {"name": "빛그린국가산업단지", "classification": "기존·확장 산업단지", "linked_zone": "평동·빛그린 산업확장권", "characteristics": "완성차 생산과 친환경·미래차 부품산업 육성을 중심으로 한 광주경제자유구역의 핵심 산업축", "market_points": "완성차·부품기업의 실제 입주·고용, 협력업체 수요, 기반시설 및 잔여용지", "verified_facts": "광산구 삼거동과 함평군 월야면 일원 약 407만㎡(광주 약 184.7만㎡·함평 약 222.4만㎡) 규모로 조성됨. 광주글로벌모터스(GGM)가 2021년 4월 공장을 준공해 캐스퍼를 위탁생산하고 금호타이어·캠스 등 부품기업이 입주함. 2024년 11월 일부 부지가 기회발전특구로 지정되어 모빌리티 특구가 추진됨.", "verified_as_of": "2026-08"},
    {"name": "광주 미래차 국가산업단지", "classification": "신규 국가산업단지(조성 추진)", "linked_zone": "평동·빛그린 산업확장권", "characteristics": "자율주행·미래차 소부장 기업 집적을 목표로 광산구 서남부에 새로 조성되는 국가산단으로, 기존 빛그린 산단의 포화에 따른 산업용지 확장축에 해당함", "market_points": "산단 경계·편입 여부, 보상 착수 시기, 앵커기업 및 소부장 기업 유치 실적, 인접 농경지·생산녹지의 사업용 토지 수요", "verified_facts": "광산구 오운동 일대 약 338만4천㎡(102만평) 규모로, 2023년 3월 정부 신규 국가산업단지 후보지로 선정되어 빛그린국가산단 지정 이후 14년 만에 유치됨. 2028년까지 약 7,000억원을 투입해 조성할 계획이며 자율주행차 부품 소부장 특화단지와 연계 추진됨. 국가산단 지정·보상·착공 등 후속 절차의 진행 상황을 별도로 확인해야 함.", "verified_as_of": "2026-08"},
    {"name": "호남권 반도체 첨단국가산업단지(광주 군공항 부지 후보지)", "classification": "신규 국가산업단지 후보지(지정 절차 진행)", "linked_zone": "송정역·도심전환권", "characteristics": "광주 군공항 부지 일원을 대상으로 한 반도체 첨단국가산단 후보지로, 송정동·도산동·신촌동·소촌동·우산동·황룡동 등 반도체 직접 영향권 읍·면·동과 접하며 군공항 이전 절차와 연동되는 개발축", "market_points": "국가산단 지정·예비타당성조사·보상 등 추진단계, 군공항 이전 절차 진행상황, 토지거래허가구역 내 허가 신청·처리 실적, 개발 기대 수요와 실수요의 구분", "verified_facts": "2026년 7월 6일 정부 발표·언론 보도로 광주 군공항 부지가 호남권 반도체 첨단국가산단 예정지로 낙점된 것이 확인됨. 국토교통부 공고 제2026-940호(2026-07-09)에 따라 예정지 일원 약 364.19㎢가 2026-07-14부터 2028-07-13까지 토지거래허가구역으로 지정되어 효력이 발생함. 국가산단 지정·착공과 군공항 이전은 후속 절차가 남아 있어 확정된 단계로 서술하지 않음.", "verified_as_of": "2026-08"},
    {"name": "첨단과학국가산업단지", "classification": "광산구·북구 연계 산업단지", "linked_zone": "첨단·비아 혁신권", "characteristics": "광산업 연구개발기관과 기업이 집적된 첨단산업 거점으로 상업·업무·배후주거 수요와 연계", "market_points": "기업 가동·연구개발 투자, 업무·상업용지 거래, 배후수요의 실제 변화"},
    {"name": "첨단3지구 AI융복합지구", "classification": "연계 개발지구", "linked_zone": "첨단·비아 혁신권", "characteristics": "AI 융복합과 미래차 시험·실증 기능이 계획된 개발축으로 기존 산단과 달리 추진단계 확인이 필요한 지역", "market_points": "고시·조성·분양·입주 단계, 기반시설 진행률, 기대수요와 실수요의 구분", "verified_facts": "행정구역상 광주 북구 첨단3지구 AI융복합지구(2020년 지정)로 광산구가 아닌 인접·연계 관점에서 서술함. 1단계(2020~2025)로 국가AI데이터센터·실증창업동 등 조성, 2단계 2026년 본격화. 2026년 6월 삼성전자 반도체 패키징 공장 조성 발표(남부권 반도체 혁신벨트 연계).", "verified_as_of": "2026-07"},
]


GWANGSAN_PROFILE_SOURCES = [
    {"name": "광주광역시 산업단지 소방안전 강화 공약자료", "url": "https://www.gwangju.go.kr/home/newgjmayor/down/2025/part4-2/4-8-2.pdf"},
    {"name": "광주광역시 경제자유구역·산단 대개조", "url": "https://www.gwangju.go.kr/economy/contentsView.do?pageId=economy49"},
    {"name": "광주광역시 자동차산업", "url": "https://www.gwangju.go.kr/economy/contentsView.do?pageId=economy54"},
    {"name": "광주광역시 금형산업", "url": "https://www.gwangju.go.kr/economy/contentsView.do?pageId=economy51"},
    {"name": "광주광역시 광산업", "url": "https://www.gwangju.go.kr/economy/contentsView.do?pageId=economy55"},
    {"name": "빛그린 미래차 국가산단·입주기업 현황(2026-07 확인)", "url": "https://www.sedaily.com/NewsView/2DGR4WAS1P"},
    {"name": "광주 미래차 국가산단 조성 계획(광산구 오운동 338만㎡·2028년 목표, 2026-08 확인)", "url": "https://www.etnews.com/20240624000302"},
    {"name": "빛그린국가산업단지 규모·위치(광주 투자환경, 2026-08 확인)", "url": "https://www.investkorea.org/gj-kr/bbs/i-1071/detail.do?ntt_sn=486320"},
    {"name": "첨단3지구 AI·삼성 반도체 패키징(2026 확인)", "url": "https://www.seoul.co.kr/news/economy/estate/2026/04/17/20260417022001"},
]


# 9장 서술 첫 문장에 쓰는 동·조사권역별 한 구절 설명(조사권역 명칭이 있으면 그것을 우선 사용).
GWANGSAN_DONG_PROFILES = {
    "송정동": "광주송정역 역세권의 상업지와 노후 주거지가 섞인 원도심",
    "도산동": "송정역 남측의 저층 단독주택 중심 주거지",
    "도호동": "공항 주변 농경지와 자연취락 위주의 지역",
    "신촌동": "광주공항에 접한 주거지와 농경지",
    "우산동": "공항 인접 저층 주거지와 생산녹지가 섞인 지역",
    "황룡동": "황룡강변 농경지와 취락 위주의 지역",
    "선암동": "황룡강변 자연녹지와 선운지구 인접 주거지",
    "소촌동": "소촌농공단지 등 공장지대와 배후 주거·상업지",
    "도호동,신촌동,우산동": "광주공항을 둘러싼 저층 주거지와 생산녹지·농경지가 섞인 지역",
    "선암동,소촌동": "소촌농공단지 공장지대와 황룡강변 자연녹지가 이어진 지역",
}

# 동별 설명이 없는 지역에 쓰는 권역 요약(brief)과 상승 배경(rise).
GWANGSAN_ZONE_BRIEFS = {
    "송정역·도심전환권": {"brief": "송정역세권과 공항 주변의 기존 시가지", "rise": "역세권 개발과 반도체 국가산단 조성 기대"},
    "수완·신가 생활권": {"brief": "대규모 공동주택과 상업시설이 모인 주거·생활권", "rise": "생활권 실수요와 상업용지 수요"},
    "첨단·비아 혁신권": {"brief": "첨단과학산단과 배후 주거·상업지가 이어진 지역", "rise": "첨단산단 배후 수요와 첨단3지구 개발 기대"},
    "하남·진곡 제조산업권": {"brief": "하남·진곡산단의 공장·창고와 배후지", "rise": "산업용지 수요"},
    "본량·삼도 서부 농촌권": {"brief": "농경지와 임야, 취락이 넓게 분포한 농촌 지역", "rise": "대면적 토지 수요와 인접 개발사업 기대"},
    "평동·빛그린 산업확장권": {"brief": "평동·빛그린 산단과 주변 농지·취락이 섞인 도농복합 지역", "rise": "산단 확장과 미래차 국가산단 조성 기대"},
}


def _field_monitoring_paths(input_dir: Path | None = None) -> list[Path]:
    """중개사무소 주간 모니터링 기록 파일(field_monitoring.json) 탐색 경로."""
    paths: list[Path] = []
    env = os.environ.get("LAND_PRICE_FIELD_MONITORING")
    if env:
        paths.append(Path(env))
    if input_dir:
        paths.append(Path(input_dir) / "field_monitoring.json")
    paths.append(Path.cwd() / "field_monitoring.json")
    try:
        paths.append(Path(__file__).resolve().parent / "field_monitoring.json")
    except NameError:
        pass
    return paths


def load_field_monitoring(region: str, base_month: str, input_dir: Path | None = None) -> dict[str, Any]:
    """조사담당자가 기록한 중개사무소 청취(모니터링) 요지를 읽는다.

    실거래 엑셀에서 나오지 않는 매수 문의·호가·매물 회수 등 현장 정보는 이 파일에만 근거한다.
    파일이 없으면 available=False로 두고, 서술에서는 '현장 확인 필요'로 처리한다.
    """
    for path in _field_monitoring_paths(input_dir):
        try:
            if not path or not path.exists():
                continue
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            return {"available": False, "reason": f"모니터링 파일 읽기 실패: {exc}"}
        entries = data.get("entries", []) if isinstance(data, dict) else []
        region_text = norm_text(region)
        matched: list[dict[str, Any]] = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            keys = entry.get("region_match") or [entry.get("region", "")]
            if not any(norm_text(key) and norm_text(key) in region_text for key in keys):
                continue
            if entry.get("base_month") and str(entry.get("base_month")) != str(base_month):
                continue
            matched.append(entry)
        if matched:
            matched.sort(key=lambda item: str(item.get("as_of") or ""))
            latest = matched[-1]
            return {
                "available": True,
                "source_file": str(path),
                "source": latest.get("source") or "중개사무소 모니터링",
                "as_of": latest.get("as_of"),
                "base_month": latest.get("base_month"),
                "overall": latest.get("overall", ""),
                "groups": latest.get("groups", {}) or {},
                "dongs": latest.get("dongs", {}) or {},
                "price_trend": latest.get("price_trend", ""),
                "outlook": latest.get("outlook", ""),
                "notes": latest.get("notes", ""),
                "history": matched[:-1],
            }
        return {"available": False, "reason": "모니터링 파일에 해당 시군구·기준월 기록 없음", "source_file": str(path)}
    return {"available": False, "reason": "field_monitoring.json 미작성(중개사무소 청취 기록 없음)"}


_REGIONAL_DATA_CACHE: dict[str, Any] | None = None


def _regional_data_paths() -> list[Path]:
    """외부 지역 데이터 파일(regional_profiles.json) 탐색 경로."""
    paths: list[Path] = []
    env = os.environ.get("LAND_PRICE_REGIONAL_PROFILES")
    if env:
        paths.append(Path(env))
    paths.append(Path.cwd() / "regional_profiles.json")
    try:
        paths.append(Path(__file__).resolve().parent / "regional_profiles.json")
    except NameError:
        pass
    return paths


def load_regional_data(force: bool = False) -> dict[str, Any]:
    """편집 가능한 외부 regional_profiles.json을 로드한다(있을 때만).

    코드를 수정하지 않고 시군구를 추가·재정의하기 위한 데이터 파일이다.
    파일이 없거나 손상되면 빈 dict를 돌려 기존 하드코딩 경로로 폴백한다.
    """
    global _REGIONAL_DATA_CACHE
    if _REGIONAL_DATA_CACHE is not None and not force:
        return _REGIONAL_DATA_CACHE
    data: dict[str, Any] = {}
    for path in _regional_data_paths():
        try:
            if path and path.is_file():
                loaded = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    data = loaded
                    break
        except Exception:
            continue
    _REGIONAL_DATA_CACHE = data
    return data


def _match_region_entry(region: str) -> dict[str, Any] | None:
    """region 문자열에 match_keys가 포함되는 외부 데이터 항목을 찾는다(가장 긴 키 우선)."""
    regions = load_regional_data().get("regions") or {}
    if not isinstance(regions, dict):
        return None
    region_text = str(region or "")
    best: tuple[int, dict[str, Any]] | None = None
    for key, entry in regions.items():
        if not isinstance(entry, dict):
            continue
        for match_key in (entry.get("match_keys") or [key]):
            match_key = str(match_key)
            if match_key and match_key in region_text:
                if best is None or len(match_key) > best[0]:
                    best = (len(match_key), entry)
    return best[1] if best else None


def _apply_issue_placeholders(value: Any, display_area: str, issue_scope_note: str) -> Any:
    """외부 데이터의 {display_area}/{issue_scope_note} 자리표시자를 실제 값으로 치환한다."""
    if isinstance(value, str):
        return value.replace("{display_area}", display_area).replace("{issue_scope_note}", issue_scope_note)
    if isinstance(value, list):
        return [_apply_issue_placeholders(v, display_area, issue_scope_note) for v in value]
    if isinstance(value, dict):
        return {k: _apply_issue_placeholders(v, display_area, issue_scope_note) for k, v in value.items()}
    return value


def build_regional_profile(region: str) -> dict[str, Any]:
    """보고서 서술에 사용할 생활권·산업권 기준을 제공한다.

    외부 regional_profiles.json에 해당 시군구가 정의되어 있으면 그 값을 우선 사용하고,
    없으면 코드에 내장된 광산구 기준으로 폴백한다.
    """
    entry = _match_region_entry(region)
    if entry and isinstance(entry.get("profile"), dict):
        prof = entry["profile"]
        return {
            "available": True,
            "region": region,
            "classification_basis": prof.get("classification_basis", ""),
            "zones": prof.get("zones", []),
            "industrial_complexes": prof.get("industrial_complexes", []),
            "sources": prof.get("sources", []),
            "causality_rule": prof.get(
                "causality_rule",
                "산단 또는 개발지구 인접성만으로 거래·지가 상승의 원인이라고 단정하지 않고 실제 거래, 기업 입주·가동, 사업 추진단계를 함께 확인한다.",
            ),
            "analysis_units": prof.get("analysis_units", {}),
            "overlap_rule": prof.get("overlap_rule", ""),
            "dong_profiles": prof.get("dong_profiles", {}),
            "zone_briefs": prof.get("zone_briefs", {}),
        }
    if "광산구" not in str(region or ""):
        return {"available": False, "region": region, "zones": [], "industrial_complexes": [], "sources": []}
    return {
        "available": True,
        "region": region,
        "classification_basis": "행정동 경계를 그대로 합산한 통계권역이 아니라 생활권·산업입지·토지이용 특성을 결합한 보고서 분석권역",
        "zones": GWANGSAN_REGIONAL_ZONES,
        "industrial_complexes": GWANGSAN_INDUSTRIAL_COMPLEXES,
        "sources": GWANGSAN_PROFILE_SOURCES,
        "causality_rule": "산단 또는 개발지구 인접성만으로 거래·지가 상승의 원인이라고 단정하지 않고 실제 거래, 기업 입주·가동, 사업 추진단계를 함께 확인한다.",
        "analysis_units": {
            "price_survey_area": "지가동향 원표의 조사권역 명칭을 그대로 사용",
            "living_zone": "도호·신촌은 송정권, 우산은 수완·신가 생활권 및 공항 인접권으로 구분",
            "semiconductor_influence": "송정동·도산동·신촌동·소촌동·우산동·황룡동 반도체 직접 영향권, 선암동·운수동 연계 영향권",
        },
        "overlap_rule": "도호동·신촌동·우산동이 하나의 가격조사 권역으로 표시되더라도 우산동 생산녹지 거래와 도호·신촌동의 역세권·기존 시가지 수요를 분리해 서술한다.",
        "dong_profiles": GWANGSAN_DONG_PROFILES,
        "zone_briefs": GWANGSAN_ZONE_BRIEFS,
    }


def match_regional_zone(area: Any, profile: dict[str, Any]) -> dict[str, Any] | None:
    normalized = norm_text(area)
    if not normalized:
        return None
    best: tuple[int, dict[str, Any]] | None = None
    for zone in profile.get("zones", []):
        for dong in zone.get("dongs", []):
            token = norm_text(dong)
            if token and (token in normalized or normalized in token):
                score = len(token)
                if best is None or score > best[0]:
                    best = (score, zone)
    return best[1] if best else None


def semiconductor_influence(name: Any, region: str = "") -> dict[str, str] | None:
    """Return the report-only semiconductor influence classification for Gwangsan-gu."""
    if "광산구" not in str(region or ""):
        return None
    normalized = re.sub(r"\s+", "", str(name or ""))
    for dong, (level, label) in SEMICONDUCTOR_INFLUENCE_DONGS.items():
        if normalized == dong or normalized.endswith(dong):
            return {"level": level, "label": label, "dong": dong}
    return None


def norm_text(value: Any) -> str:
    return re.sub(r"\s+", "", "" if value is None else str(value)).strip()


def _josa_euro(word: Any) -> str:
    """받침 유무에 따라 '으로'/'로'를 고른다(ㄹ 받침은 '로')."""
    text = str(word or "").strip()
    if not text:
        return "으로"
    last = text[-1]
    if "가" <= last <= "힣":
        jong = (ord(last) - 0xAC00) % 28
        return "로" if jong in (0, 8) else "으로"
    return "으로"


def as_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    text = str(value).replace(",", "").replace("'", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def ym_compact(base_month: str) -> str:
    return re.sub(r"[^0-9]", "", base_month)[:6]


def previous_ym(base_month: str) -> str:
    y, m = [int(x) for x in base_month.split("-")]
    if m == 1:
        return f"{y - 1}-12"
    return f"{y}-{m - 1:02d}"


def safe_name(text: str) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣_]+", "_", text).strip("_") or "region"


def find_files(input_dir: Path) -> dict[str, list[Path]]:
    files = [p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in {".xlsx", ".xls", ".csv"}]
    buckets = {
        "transactions": [],
        "land_price": [],
        "population": [],
        "household": [],
        "migration": [],
        "rates": [],
        "listings": [],
    }
    for path in files:
        name = path.name
        compact = norm_text(name)
        if ("실거래" in name or "RTMS" in name.upper() or "국토교통부_토지" in name) and "매매" in name:
            buckets["transactions"].append(path)
        elif "전국지가동향" in name or "지가변동률" in name or "지가동향" in name:
            buckets["land_price"].append(path)
        elif "주민등록인구" in name:
            buckets["population"].append(path)
        elif "주민등록세대수" in name or "세대수" in name:
            buckets["household"].append(path)
        elif "이동자수" in name or "인구이동" in name:
            buckets["migration"].append(path)
        elif "대출금리" in name or "예금은행" in name or "loan_rate" in compact.lower():
            buckets["rates"].append(path)
        elif "네이버" in name or "매물" in name:
            buckets["listings"].append(path)
    return buckets


def read_excel_any(path: Path, **kwargs) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, **{k: v for k, v in kwargs.items() if k in {"encoding"}})
    return pd.read_excel(path, **kwargs)


def first_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    compact = {norm_text(c): c for c in df.columns}
    for cand in candidates:
        key = norm_text(cand)
        if key in compact:
            return compact[key]
    for col in df.columns:
        c = norm_text(col)
        if any(norm_text(cand) in c for cand in candidates):
            return col
    return None


def land_permit_rule(region: str) -> dict[str, Any] | None:
    """지역별 토지거래허가구역 지정 기준(계약일 전·후 비교 기준일)을 반환한다."""
    text = str(region or "")
    for key, rule in LAND_PERMIT_ZONE_RULES.items():
        if key in text:
            return {"region_key": key, **rule}
    return None


def week_end_date(value: date) -> date:
    """해당 일자가 속한 주(토~금)의 종료일인 금요일을 반환한다."""
    return value + timedelta(days=(WEEK_END_WEEKDAY - value.weekday()) % 7)


def last_week_end_on_or_before(value: date) -> date:
    """작성 기준일 이전(당일 포함)의 가장 최근 금요일을 반환한다."""
    return value - timedelta(days=(value.weekday() - WEEK_END_WEEKDAY) % 7)


def resolve_as_of_date(base_month: str, as_of: Any = None) -> date:
    """주간 집계의 작성 기준일. 미지정 시 오늘과 기준월 말일 중 이른 날을 쓴다."""
    if as_of:
        parsed = pd.to_datetime(str(as_of), errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
    # 기준월이 지난 뒤에도 주간 작성이 이어지므로 오늘 날짜를 그대로 쓴다.
    # (기준월 다음 달 주차까지 진행 구간으로 표시하기 위함)
    today = datetime.now().date()
    try:
        month_start = pd.Period(base_month, freq="M").start_time.date()
    except Exception:
        return today
    return max(today, month_start)


def _bundle_stats(source: pd.DataFrame) -> dict[str, Any]:
    amount = float(source["_amount"].sum()) if not source.empty else 0.0
    external = source[source["_external"]] if not source.empty else source
    external_amount = float(external["_amount"].sum()) if not external.empty else 0.0
    return {
        "거래건수": int(len(source)),
        "거래금액_억원": round(amount / EOK, 1),
        "외지인거래건수": int(source["_external"].sum()) if not source.empty else 0,
        "외지인거래금액_억원": round(external_amount / EOK, 1),
        "외지인금액비중": round(external_amount / amount * 100, 1) if amount else None,
    }


def summarize_permit_split(bundle: pd.DataFrame, base_month: str, region: str) -> dict[str, Any]:
    """토지거래허가구역 지정 전·후를 계약일 기준으로 나눠 읍·면·동 거래를 집계한다.

    지정 효력월(광산구 2026-07)은 월 안에서 지정 전(1일~13일)·지정 후(14일~말일)로 나누고,
    효력월 이후 달은 전 구간이 지정 후이므로 월 전체를 지정 후 구간으로 표기한다.
    """
    rule = land_permit_rule(region)
    if not rule:
        return {"available": False, "reason": "토지거래허가구역 지정 기준이 정의되지 않은 지역"}
    effective = pd.to_datetime(rule["effective_date"]).date()
    try:
        period = pd.Period(base_month, freq="M")
    except Exception:
        return {"available": False, "reason": "기준월 해석 불가"}
    month_start = period.start_time.date()
    month_end = period.end_time.date()
    if month_end < effective:
        return {
            "available": False,
            "reason": "토지거래허가구역 지정 효력 발생(2026-07-14) 이전 기준월로 월단위 집계만 적용",
            "effective_date": rule["effective_date"],
            "mode": "monthly_only",
        }

    mode = "split" if month_start < effective else "post_only"
    pre_start, pre_end = month_start, effective - timedelta(days=1)
    post_start, post_end = max(month_start, effective), month_end
    dates = bundle["_date"].dt.date

    def period_frame(start: date | None, end: date | None) -> pd.DataFrame:
        if start is None or end is None or start > end:
            return bundle.iloc[0:0]
        return bundle[(dates >= start) & (dates <= end)]

    pre_df = period_frame(pre_start, pre_end) if mode == "split" else bundle.iloc[0:0]
    post_df = period_frame(post_start, post_end)
    pre_days = (pre_end - pre_start).days + 1 if mode == "split" else 0
    post_days = (post_end - post_start).days + 1

    def emd_counts(source: pd.DataFrame) -> dict[str, dict[str, float]]:
        if source.empty:
            return {}
        grouped = source.groupby("_emd", dropna=False).agg(건수=("_amount", "size"), 금액=("_amount", "sum"))
        return {
            str(name): {"건수": int(row["건수"]), "금액_억원": round(float(row["금액"]) / EOK, 1)}
            for name, row in grouped.iterrows()
        }

    pre_map, post_map = emd_counts(pre_df), emd_counts(post_df)
    # 거래가 없더라도 반도체 직접 영향권 읍면동은 0건으로 표시한다.
    names = set(pre_map) | set(post_map)
    if "광산구" in str(region or ""):
        existing = {norm_text(name) for name in names}
        names |= {dong for dong in semiconductor_dongs("direct") if norm_text(dong) not in existing}
    rows: list[dict[str, Any]] = []
    for name in sorted(names):
        pre_item = pre_map.get(name, {"건수": 0, "금액_억원": 0.0})
        post_item = post_map.get(name, {"건수": 0, "금액_억원": 0.0})
        info = semiconductor_influence(name, region)
        pre_daily = round(pre_item["건수"] / pre_days, 2) if pre_days else None
        post_daily = round(post_item["건수"] / post_days, 2) if post_days else None
        rows.append({
            "읍면동": name,
            "영향권등급": info["level"] if info else "",
            "영향권라벨": info["label"] if info else "",
            "지정전_거래건수": pre_item["건수"],
            "지정전_거래금액_억원": pre_item["금액_억원"],
            "지정전_일평균건수": pre_daily,
            "지정후_거래건수": post_item["건수"],
            "지정후_거래금액_억원": post_item["금액_억원"],
            "지정후_일평균건수": post_daily,
            "건수증감": post_item["건수"] - pre_item["건수"],
            "일평균증감률": (
                round((post_daily - pre_daily) / pre_daily * 100, 1)
                if pre_daily not in (None, 0) and post_daily is not None else None
            ),
            "합계건수": pre_item["건수"] + post_item["건수"],
            "합계금액_억원": round(pre_item["금액_억원"] + post_item["금액_억원"], 1),
        })
    rows.sort(key=lambda row: (row["합계건수"], row["합계금액_억원"]), reverse=True)

    # 현장점검 양식 서술 단위(직접/간접 영향권)로도 지정 전·후를 집계한다.
    group_rows: list[dict[str, Any]] = []
    if "광산구" in str(region or ""):
        direct_names = {norm_text(dong) for dong in semiconductor_dongs("direct")}
        for label, scope_text, wanted in [
            ("직접 영향권", "·".join(semiconductor_dongs("direct")), True),
            ("간접 영향권", "직접 영향권 외 읍·면·동", False),
        ]:
            def side(source: pd.DataFrame) -> tuple[int, float]:
                if source.empty:
                    return 0, 0.0
                mask = source["_emd"].map(lambda value: (norm_text(value) in direct_names) == wanted)
                picked = source[mask]
                return int(len(picked)), round(float(picked["_amount"].sum()) / EOK, 1)

            pre_count, pre_amount = side(pre_df)
            post_count, post_amount = side(post_df)
            pre_daily = round(pre_count / pre_days, 2) if pre_days else None
            post_daily = round(post_count / post_days, 2) if post_days else None
            group_rows.append({
                "구분": label,
                "대상": scope_text,
                "지정전_거래건수": pre_count,
                "지정전_거래금액_억원": pre_amount,
                "지정전_일평균건수": pre_daily,
                "지정후_거래건수": post_count,
                "지정후_거래금액_억원": post_amount,
                "지정후_일평균건수": post_daily,
                "건수증감": post_count - pre_count,
                "일평균증감률": (
                    round((post_daily - pre_daily) / pre_daily * 100, 1)
                    if pre_daily not in (None, 0) and post_daily is not None else None
                ),
            })

    pre_total = _bundle_stats(pre_df)
    post_total = _bundle_stats(post_df)
    pre_daily_total = round(pre_total["거래건수"] / pre_days, 2) if pre_days else None
    post_daily_total = round(post_total["거래건수"] / post_days, 2) if post_days else None
    return {
        "available": True,
        "mode": mode,
        "region_rule": rule,
        "effective_date": rule["effective_date"],
        "basis": "실거래 계약일(계약일자) 기준",
        "periods": {
            "pre": {
                "label": f"{pre_start:%Y-%m-%d} ~ {pre_end:%Y-%m-%d}" if mode == "split" else "해당 없음",
                "start": f"{pre_start:%Y-%m-%d}" if mode == "split" else None,
                "end": f"{pre_end:%Y-%m-%d}" if mode == "split" else None,
                "days": pre_days,
                "구분": "지정 전",
            },
            "post": {
                "label": f"{post_start:%Y-%m-%d} ~ {post_end:%Y-%m-%d}",
                "start": f"{post_start:%Y-%m-%d}",
                "end": f"{post_end:%Y-%m-%d}",
                "days": post_days,
                "구분": "지정 후",
            },
        },
        "totals": {
            "pre": {**pre_total, "일평균건수": pre_daily_total, "일수": pre_days},
            "post": {**post_total, "일평균건수": post_daily_total, "일수": post_days},
            "건수증감": post_total["거래건수"] - pre_total["거래건수"],
            "일평균증감률": (
                round((post_daily_total - pre_daily_total) / pre_daily_total * 100, 1)
                if pre_daily_total not in (None, 0) and post_daily_total is not None else None
            ),
        },
        "emd_rows": rows,
        "influence_group_rows": group_rows,
        "note": (
            f"지정 전·후 구간의 일수가 다르므로(전 {pre_days}일·후 {post_days}일) 건수 차이를 그대로 규제 전후 증감으로 해석하지 않아야 함. "
            "또한 계약일 기준 집계는 부동산거래 신고기한(계약일부터 30일) 때문에 최근 계약분이 아직 신고되지 않아 "
            "지정 후 구간이 과소집계될 수 있으므로, 감소폭을 규제 효과로 단정하지 말고 다음 갱신분에서 재확인해야 함."
            if mode == "split" else
            "기준월 전 구간이 토지거래허가구역 지정 이후이므로 지정 후 단일 구간으로 집계함. "
            "계약일 기준 집계는 신고기한(30일) 영향으로 최근 구간이 과소집계될 수 있음."
        ),
    }


def summarize_weekly(
    bundle: pd.DataFrame,
    base_month: str,
    region: str,
    as_of: date,
    weeks: int = WEEKLY_WINDOW_COUNT,
) -> dict[str, Any]:
    """계약일 기준 주간(토~금) 집계. 매주 금요일 작성 일정에 맞춘 최근 N주 창을 만든다."""
    anchor = last_week_end_on_or_before(as_of)
    windows: list[tuple[date, date]] = []
    for index in range(weeks - 1, -1, -1):
        end = anchor - timedelta(weeks=index)
        windows.append((end - timedelta(days=6), end))
    rule = land_permit_rule(region)
    effective = pd.to_datetime(rule["effective_date"]).date() if rule else None
    dates = bundle["_date"].dt.date

    def permit_phase(start: date, end: date) -> str:
        if effective is None:
            return "-"
        if end < effective:
            return "지정 전"
        if start >= effective:
            return "지정 후"
        return "지정 전후 걸침"

    weekly_rows: list[dict[str, Any]] = []
    previous_count: int | None = None
    for start, end in windows:
        sub = bundle[(dates >= start) & (dates <= end)]
        stats = _bundle_stats(sub)
        weekly_rows.append({
            "주차": f"{start:%m/%d}~{end:%m/%d}",
            "시작일": f"{start:%Y-%m-%d}",
            "종료일": f"{end:%Y-%m-%d}",
            "허가구역구분": permit_phase(start, end),
            **stats,
            "직거래건수": int(sub["_direct"].sum()) if not sub.empty else 0,
            "전주대비건수증감": None if previous_count is None else stats["거래건수"] - previous_count,
        })
        previous_count = stats["거래건수"]

    week_labels = [row["주차"] for row in weekly_rows]
    emd_rows: list[dict[str, Any]] = []
    period_start, period_end = windows[0][0], windows[-1][1]
    scope = bundle[(dates >= period_start) & (dates <= period_end)]
    if not scope.empty:
        for emd_name, emd_data in scope.groupby("_emd", dropna=False):
            emd_dates = emd_data["_date"].dt.date
            values = []
            for start, end in windows:
                week_sub = emd_data[(emd_dates >= start) & (emd_dates <= end)]
                values.append({
                    "주차": f"{start:%m/%d}~{end:%m/%d}",
                    "거래건수": int(len(week_sub)),
                    "거래금액_억원": round(float(week_sub["_amount"].sum()) / EOK, 1),
                })
            info = semiconductor_influence(emd_name, region)
            emd_rows.append({
                "읍면동": str(emd_name),
                "영향권등급": info["level"] if info else "",
                "영향권라벨": info["label"] if info else "",
                "주별": values,
                "합계건수": sum(item["거래건수"] for item in values),
                "합계금액_억원": round(sum(item["거래금액_억원"] for item in values), 1),
            })
        emd_rows.sort(key=lambda row: (row["합계건수"], row["합계금액_억원"]), reverse=True)

    # 현장점검 서술용 근거: 최근 계약 건 목록(읍면동별·영향권 구분별)
    def trade_rows(source: pd.DataFrame, limit: int = 5) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        if source is None or source.empty:
            return rows
        for _, item in source.sort_values("_date", ascending=False).head(limit).iterrows():
            area = float(as_number(item.get("_area")) or 0)
            amount = float(as_number(item.get("_amount")) or 0)
            parcel_count = int(as_number(item.get("_parcel_rows")) or 1)
            address = " ".join(
                token for token in [str(item.get("_emd") or ""), str(item.get("_jibun") or "")]
                if token and token != "nan"
            )
            if parcel_count > 1:
                address = f"{address} 외({parcel_count - 1:02d}필지)"
            rows.append({
                "계약일": item.get("_date").strftime("%Y-%m-%d") if pd.notna(item.get("_date")) else "",
                "읍면동": str(item.get("_emd") or ""),
                "소재지": address,
                "물건구분": str(item.get("_property") or ""),
                "용도지역": str(item.get("_zone") or ""),
                "지목": str(item.get("_jimok") or ""),
                "이용상황": str(item.get("_use") or ""),
                "면적_㎡": round(area, 1),
                "거래금액_억원": round(amount / EOK, 2),
                "단가_원㎡": (
                    math.floor((amount / area) / 1_000) * 1_000
                    if parcel_count == 1 and bool(item.get("_land_only")) and area > 0 else None
                ),
                "거래주체": str(item.get("_buyer_type") or "미상"),
                "매수자시군구": str(item.get("_buyer_region") or ""),
                "매수자지역": str(item.get("_buyer_region") or ""),
                "외지인": "Y" if bool(item.get("_external")) else "N",
                "허가구역구분": (
                    "-" if effective is None else
                    "지정 후" if pd.notna(item.get("_date")) and item.get("_date").date() >= effective else "지정 전"
                ),
            })
        return rows

    recent_trades: dict[str, list[dict[str, Any]]] = {}
    group_rows: list[dict[str, Any]] = []
    group_recent_trades: dict[str, list[dict[str, Any]]] = {}
    if "광산구" in str(region or ""):
        direct_names = {norm_text(dong) for dong in semiconductor_dongs("direct")}
        for dong in semiconductor_dongs("direct"):
            target = scope[scope["_emd"].map(lambda value: norm_text(value) == norm_text(dong))] if not scope.empty else scope
            recent_trades[dong] = trade_rows(target)
        # 현장점검 양식 서술 단위: 직접 영향권 6개 동 묶음 / 그 외 읍면동(간접 영향권) 묶음
        is_direct = scope["_emd"].map(lambda value: norm_text(value) in direct_names) if not scope.empty else scope
        groups = [
            ("직접 영향권", "·".join(semiconductor_dongs("direct")), scope[is_direct] if not scope.empty else scope),
            ("간접 영향권", f"{region.split()[-1] if region.split() else region} 직접 영향권 외 읍·면·동",
             scope[~is_direct] if not scope.empty else scope),
        ]
        for label, target_scope_text, frame in groups:
            frame_dates = frame["_date"].dt.date if not frame.empty else None
            values = []
            for start, end in windows:
                week_sub = frame[(frame_dates >= start) & (frame_dates <= end)] if not frame.empty else frame
                values.append({
                    "주차": f"{start:%m/%d}~{end:%m/%d}",
                    "거래건수": int(len(week_sub)),
                    "거래금액_억원": round(float(week_sub["_amount"].sum()) / EOK, 1) if not week_sub.empty else 0.0,
                })
            included = sorted({str(name) for name in frame["_emd"].unique()}) if not frame.empty else []
            group_rows.append({
                "구분": label,
                "대상": target_scope_text,
                "거래확인읍면동": included,
                "주별": values,
                "합계건수": sum(item["거래건수"] for item in values),
                "합계금액_억원": round(sum(item["거래금액_억원"] for item in values), 1),
                "기준주거래건수": values[-1]["거래건수"] if values else 0,
                "기준주거래금액_억원": values[-1]["거래금액_억원"] if values else 0.0,
                "전주거래건수": values[-2]["거래건수"] if len(values) > 1 else 0,
                "외지인거래건수": int(frame["_external"].sum()) if not frame.empty else 0,
            })
            group_recent_trades[label] = trade_rows(frame, 6)

    latest = weekly_rows[-1] if weekly_rows else {}
    return {
        "available": True,
        "basis": "실거래 계약일(계약일자) 기준",
        "week_rule": "토요일 시작 ~ 금요일 종료의 7일 창(매주 금요일 작성 기준)",
        "as_of": f"{as_of:%Y-%m-%d}",
        "report_week": {
            "주차": latest.get("주차"),
            "시작일": latest.get("시작일"),
            "종료일": latest.get("종료일"),
            "허가구역구분": latest.get("허가구역구분"),
        },
        "week_labels": week_labels,
        "weekly_rows": weekly_rows,
        "weekly_emd_rows": emd_rows,
        "emd_recent_trades": recent_trades,
        "influence_group_rows": group_rows,
        "influence_group_recent_trades": group_recent_trades,
        "influence_group_rule": (
            "직접 영향권 = 송정동·도산동·신촌동·소촌동·우산동·황룡동, "
            "간접 영향권 = 해당 시군구 내 직접 영향권 외 읍·면·동"
        ),
        "note": (
            "주간 집계는 계약일 기준이므로 신고기한(계약일부터 30일) 내 미신고 계약이 이후 갱신분에서 추가되어 "
            "최근 2~4주 수치는 과소집계 상태일 수 있음. 최근 주차 감소는 확정된 거래 감소로 단정하지 말고 다음 주 갱신분에서 재확인해야 함."
        ),
    }


def month_week_windows(base_month: str, as_of: date | None = None) -> list[tuple[date, date, str]]:
    """기준월(및 작성일이 속한 이후 진행월)을 금요일 종료 주간 구간으로 나눈다.

    월 경계에서 잘라 각 구간이 한 달 안에만 속하도록 하여 월 합계와 주 합계가 일치한다.
    작성일이 기준월 다음 달이면 그 달의 주 구간도 작성일까지 이어서 만든다.
    """
    period = pd.Period(base_month, freq="M")
    month_start, month_end = period.start_time.date(), period.end_time.date()
    limit = month_end
    if as_of and as_of > month_end:
        limit = as_of
    windows: list[tuple[date, date, str]] = []
    cursor = month_start
    while cursor <= limit:
        current_month_end = pd.Period(f"{cursor:%Y-%m}", freq="M").end_time.date()
        end = min(week_end_date(cursor), current_month_end, limit)
        windows.append((cursor, end, f"{cursor:%m/%d}~{end:%m/%d}"))
        cursor = end + timedelta(days=1)
    return windows


def summarize_period_trend(
    bundle: pd.DataFrame,
    base_month: str,
    region: str,
    months_back: int = 12,
    as_of: date | None = None,
) -> dict[str, Any]:
    """전체 거래 추이: 기준월 이전 N개월은 월단위, 기준월은 주단위(금요일 종료) 행으로 이어 붙인다."""
    period = pd.Period(base_month, freq="M")
    month_labels = [str(period - offset) for offset in range(months_back, 0, -1)]
    windows = month_week_windows(base_month, as_of)
    rule = land_permit_rule(region)
    effective = pd.to_datetime(rule["effective_date"]).date() if rule else None
    dates = bundle["_date"].dt.date

    def phase(start: date, end: date) -> str:
        if effective is None:
            return "-"
        if end < effective:
            return "지정 전"
        if start >= effective:
            return "지정 후"
        return "지정 전후 걸침"

    rows: list[dict[str, Any]] = []
    previous_count: int | None = None
    last_month_count: int | None = None
    for label in month_labels:
        sub = bundle[bundle["_ym_contract"] == label]
        month_period = pd.Period(label, freq="M")
        stats = _bundle_stats(sub)
        rows.append({
            "구분": "월",
            "기간": label,
            "시작일": f"{month_period.start_time.date():%Y-%m-%d}",
            "종료일": f"{month_period.end_time.date():%Y-%m-%d}",
            "허가구역구분": phase(month_period.start_time.date(), month_period.end_time.date()),
            **stats,
            "직거래건수": int(sub["_direct"].sum()) if not sub.empty else 0,
            "전기대비건수증감": None if previous_count is None else stats["거래건수"] - previous_count,
        })
        previous_count = stats["거래건수"]
        last_month_count = stats["거래건수"]
    # 주 구간은 기준월 → (작성일이 다음 달이면) 진행월 순서로 이어지며, 달이 끝날 때마다 소계를 넣는다.
    def week_row(start: date, end: date, label: str) -> dict[str, Any]:
        nonlocal previous_count
        sub = bundle[(dates >= start) & (dates <= end)]
        stats = _bundle_stats(sub)
        row = {
            "구분": "주",
            "기간": label,
            "시작일": f"{start:%Y-%m-%d}",
            "종료일": f"{end:%Y-%m-%d}",
            "허가구역구분": phase(start, end),
            **stats,
            "직거래건수": int(sub["_direct"].sum()) if not sub.empty else 0,
            "전기대비건수증감": None if previous_count is None else stats["거래건수"] - previous_count,
        }
        previous_count = stats["거래건수"]
        return row

    def month_total_row(month_label: str, windows_in_month: list[tuple[date, date, str]],
                        compare_count: int | None, compare_label: str) -> dict[str, Any]:
        month_period = pd.Period(month_label, freq="M")
        start = windows_in_month[0][0]
        end = windows_in_month[-1][1]
        sub = bundle[(dates >= start) & (dates <= end)]
        stats = _bundle_stats(sub)
        partial = end < month_period.end_time.date()
        return {
            "구분": "합계",
            "기간": f"{month_label} {'누적' if partial else '합계'}",
            "시작일": f"{start:%Y-%m-%d}",
            "종료일": f"{end:%Y-%m-%d}",
            "허가구역구분": phase(start, end),
            **stats,
            "직거래건수": int(sub["_direct"].sum()) if not sub.empty else 0,
            "전기대비건수증감": None if compare_count is None else stats["거래건수"] - compare_count,
            "비교기준": compare_label,
            "진행중": partial,
        }

    months_in_windows: list[str] = []
    for start, _, _ in windows:
        label = f"{start:%Y-%m}"
        if label not in months_in_windows:
            months_in_windows.append(label)
    base_total_count: int | None = None
    for month_label in months_in_windows:
        month_windows = [w for w in windows if f"{w[0]:%Y-%m}" == month_label]
        for start, end, label in month_windows:
            rows.append(week_row(start, end, label))
        if month_label == base_month:
            total = month_total_row(
                month_label, month_windows, last_month_count,
                f"직전월({month_labels[-1]}) 대비" if month_labels else "",
            )
            base_total_count = total["거래건수"]
        else:
            total = month_total_row(month_label, month_windows, base_total_count, f"기준월({base_month}) 대비")
        rows.append(total)
        previous_count = total["거래건수"]
    return {
        "available": True,
        "basis": "실거래 계약일(계약일자) 기준",
        "rule": f"기준월({base_month}) 이전 {months_back}개월은 월단위, 기준월은 금요일 종료 주단위와 월 합계로 표시함",
        "month_labels": month_labels,
        "week_labels": [label for _, _, label in windows],
        "rows": rows,
    }


def summarize_emd_month_week(
    bundle: pd.DataFrame,
    base_month: str,
    region: str,
    months_back: int = 12,
    as_of: date | None = None,
) -> dict[str, Any]:
    """읍·면·동별 거래를 '기준월 이전=월단위, 기준월=주단위'로 이어 붙인 표 자료.

    기준월이 바뀌면 직전 기준월은 자동으로 월단위 칸으로 넘어간다.
    """
    period = pd.Period(base_month, freq="M")
    month_labels = [str(period - offset) for offset in range(months_back, 0, -1)]
    windows = month_week_windows(base_month, as_of)
    week_labels = [label for _, _, label in windows]
    dates = bundle["_date"].dt.date
    month_start = period.start_time.date()
    scope_start = pd.Period(month_labels[0], freq="M").start_time.date() if month_labels else month_start
    # 진행월(작성일이 속한 다음 달) 주 구간까지 포함하도록 종료일을 마지막 주 구간에 맞춘다.
    scope_end = windows[-1][1] if windows else period.end_time.date()
    scope = bundle[(dates >= scope_start) & (dates <= scope_end)]

    rows: list[dict[str, Any]] = []
    if not scope.empty:
        for emd_name, emd_data in scope.groupby("_emd", dropna=False):
            emd_dates = emd_data["_date"].dt.date
            monthly = []
            month_amount_raw = 0.0
            for label in month_labels:
                sub = emd_data[emd_data["_ym_contract"] == label]
                amount = float(sub["_amount"].sum()) if not sub.empty else 0.0
                month_amount_raw += amount
                monthly.append({
                    "기간": label,
                    "거래건수": int(len(sub)),
                    "거래금액_억원": round(amount / EOK, 1),
                })
            weekly = []
            week_amount_raw = 0.0
            base_count, base_amount_raw = 0, 0.0
            ongoing_count, ongoing_amount_raw = 0, 0.0
            for start, end, label in windows:
                sub = emd_data[(emd_dates >= start) & (emd_dates <= end)]
                amount = float(sub["_amount"].sum()) if not sub.empty else 0.0
                week_amount_raw += amount
                if f"{start:%Y-%m}" == base_month:
                    base_count += int(len(sub))
                    base_amount_raw += amount
                else:
                    ongoing_count += int(len(sub))
                    ongoing_amount_raw += amount
                weekly.append({
                    "기간": label,
                    "거래건수": int(len(sub)),
                    "거래금액_억원": round(amount / EOK, 1),
                })
            info = semiconductor_influence(emd_name, region)
            rows.append({
                "읍면동": str(emd_name),
                "영향권등급": info["level"] if info else "",
                "영향권라벨": info["label"] if info else "",
                "월별": monthly,
                "주별": weekly,
                "전월건수": next((item["거래건수"] for item in monthly if item["기간"] == str(period - 1)), 0),
                "전월금액_억원": next((item["거래금액_억원"] for item in monthly if item["기간"] == str(period - 1)), 0.0),
                "당월건수": base_count,
                "당월금액_억원": round(base_amount_raw / EOK, 1),
                "전월대비건수증감": base_count - next((item["거래건수"] for item in monthly if item["기간"] == str(period - 1)), 0),
                "진행월건수": ongoing_count,
                "진행월금액_억원": round(ongoing_amount_raw / EOK, 1),
                "합계건수": sum(item["거래건수"] for item in monthly) + sum(item["거래건수"] for item in weekly),
                "합계금액_억원": round((month_amount_raw + week_amount_raw) / EOK, 1),
            })
        rows.sort(key=lambda row: (row["합계건수"], row["합계금액_억원"]), reverse=True)
    return {
        "available": True,
        "basis": "실거래 계약일(계약일자) 기준",
        "rule": f"기준월({base_month}) 이전은 월단위, 기준월은 금요일 종료 주단위로 표시함",
        "month_labels": month_labels,
        "week_labels": week_labels,
        "base_month_week_labels": [label for start, _, label in windows if f"{start:%Y-%m}" == base_month],
        "ongoing_month": next((f"{start:%Y-%m}" for start, _, _ in windows if f"{start:%Y-%m}" != base_month), None),
        "ongoing_week_labels": [label for start, _, label in windows if f"{start:%Y-%m}" != base_month],
        "week_periods": [
            {"기간": label, "시작일": f"{start:%Y-%m-%d}", "종료일": f"{end:%Y-%m-%d}"}
            for start, end, label in windows
        ],
        "rows": rows,
    }


# 지가변동률 실거래 업로드 양식은 두 가지다.
#  · 잠정치: RTMS 신고 원장(실거래_매매_○○.xlsx, 1행 헤더, RTMS_KEY·신고일)
#  · 확정치: 실거래집계 상세내역(2행 헤더, 거래번호·접수일자·취합년월·개별공시지가)
CONFIRMED_TRANSACTION_KEYS = {"거래번호", "취합년월", "접수일자", "처리구분"}
CONFIRMED_DROP_STATUS = {"계약해제", "승인불가"}


def _normalize_confirmed_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """확정치(실거래집계 상세내역) 양식을 잠정치 컬럼 체계로 변환한다.

    계약해제·승인불가 건은 집계에서 제외하고, 매수·매도자 주소에서 시군구를 뽑아
    기존 외지인 판정 로직을 그대로 사용할 수 있게 맞춘다.
    """
    work = df.copy()
    work.columns = [str(c).replace("\n", "").strip() for c in work.columns]
    status = work["처리구분"].astype(str).str.strip() if "처리구분" in work.columns else pd.Series("", index=work.index)
    work = work[~status.isin(CONFIRMED_DROP_STATUS)].copy()
    status = status.reindex(work.index)

    def col(name: str, default: Any = "") -> pd.Series:
        return work[name] if name in work.columns else pd.Series(default, index=work.index)

    def sigungu(series: pd.Series) -> pd.Series:
        text = series.astype(str).str.strip()
        return text.str.split().str[:2].str.join(" ").replace({"": "미상", "nan": "미상", "외": "미상"})

    out = pd.DataFrame(index=work.index)
    out["RTMS_KEY"] = col("거래번호").astype(str).str.strip()
    out["계약일자"] = pd.to_datetime(col("계약일자").astype(str).str.strip(), format="%Y%m%d", errors="coerce")
    out["신고일"] = pd.to_datetime(col("접수일자").astype(str).str.strip(), format="%Y%m%d", errors="coerce")
    out["계약총금액"] = col("총금액").map(as_number)
    out["토지면적"] = col("면적").map(as_number)
    out["건물면적"] = col("건물연면적").map(as_number)
    out["읍면동"] = col("읍면동").astype(str).str.strip()
    out["번지"] = col("지번").astype(str).str.strip()
    out["용도지역"] = col("용도지역").astype(str).str.strip()
    out["현실지목"] = col("지목").astype(str).str.strip()
    out["물건구분"] = col("물건구분").astype(str).str.strip()
    out["거래대상"] = out["물건구분"]
    out["매수자"] = col("거래주체").astype(str).str.strip().replace({"": "미상", "nan": "미상"})
    out["매수자 시군구"] = sigungu(col("매수자주소"))
    out["매도자 시군구"] = sigungu(col("매도자주소"))
    out["직거래구분"] = col("직거래구분").astype(str).str.strip()
    out["시군구"] = col("시군구").astype(str).str.strip()
    out["개별공시지가"] = col("개별공시지가").map(as_number)
    out["취합년월"] = col("취합년월").astype(str).str.strip()
    out["처리구분"] = status
    out["외지인거래"] = col("외지인거래").astype(str).str.strip()
    out["진단결과"] = col("진단결과").astype(str).str.strip()
    out["_source_file"] = col("_source_file")
    return out


def load_transaction_frame(path: Path) -> tuple[pd.DataFrame, str]:
    """실거래 엑셀을 읽고 (정규화된 데이터, 자료구분 '확정치'|'잠정치')을 돌려준다."""
    for header in (1, 0):
        try:
            df = pd.read_excel(path, header=header)
        except Exception:
            continue
        columns = {str(c).replace("\n", "").strip() for c in df.columns}
        if CONFIRMED_TRANSACTION_KEYS <= columns:
            df["_source_file"] = path.name
            return _normalize_confirmed_transactions(df), "확정치"
    df = read_excel_any(path)
    df["_source_file"] = path.name
    return df, "잠정치"


def summarize_transactions(
    paths: list[Path],
    base_month: str,
    region: str = "",
    as_of_date: Any = None,
) -> dict[str, Any]:
    frames: list[pd.DataFrame] = []
    data_types: list[str] = []
    file_types: list[dict[str, str]] = []
    for path in paths:
        try:
            df, data_type = load_transaction_frame(path)
            frames.append(df)
            data_types.append(data_type)
            file_types.append({"파일": path.name, "자료구분": data_type})
        except Exception as exc:
            frames.append(pd.DataFrame([{"_source_file": path.name, "_read_error": str(exc)}]))
            file_types.append({"파일": path.name, "자료구분": "읽기 실패"})
    if not frames:
        return {"available": False, "reason": "실거래 매매 파일 없음"}
    # 같은 업무에서 잠정치와 확정치를 섞으면 시계열이 어긋나므로 확정치를 우선한다.
    unique_types = list(dict.fromkeys(data_types))
    data_type = "확정치" if "확정치" in unique_types else (unique_types[0] if unique_types else "잠정치")
    if len(unique_types) > 1:
        frames = [frame for frame, kind in zip(frames, data_types) if kind == data_type]
    raw = pd.concat(frames, ignore_index=True, sort=False)
    if "_read_error" in raw.columns and raw.dropna(subset=["_read_error"]).shape[0] == len(raw):
        return {"available": False, "reason": "실거래 파일 읽기 실패", "errors": raw["_read_error"].dropna().tolist()}

    # 실거래 집계는 계약일 기준이다. 신고일 계열 열은 계약일 계열이 전혀 없을 때만 사용한다.
    # 본문 통계(3장 등)는 신고일 기준, 부록(허가 전후·주간·현장점검)은 계약일 기준으로 이원 집계한다.
    contract_col = first_col(raw, ["계약일자", "계약일", "계약년월일", "DEAL_YMD"])
    report_col = first_col(raw, ["신고일자", "신고일", "신고년월일"])
    date_col = report_col or contract_col or first_col(raw, ["거래일"])
    contract_date_col = contract_col or date_col
    date_basis = "신고일" if report_col else (norm_text(date_col or "") or "일자 미상")
    amount_col = first_col(raw, ["계약총금액", "거래금액", "거래금액(만원)", "금액"])
    area_col = first_col(raw, ["토지면적", "대지면적", "면적"])
    emd_col = first_col(raw, ["읍면동", "법정동"])
    zone_col = first_col(raw, ["용도지역"])
    use_col = first_col(raw, ["물건구분", "현실지목", "실제지목", "법정지목", "건축물주용도", "건축물용도"])
    jimok_col = first_col(raw, ["현실지목", "실제지목", "법정지목", "지목"])
    buyer_region_col = first_col(raw, ["매수자시군구", "매수자 시군구", "매수자주소"])
    # 거래주체는 실거래 엑셀의 '매수자' 원문(개인·법인·국가등·기타·외국인)을 우선 사용한다.
    # '매수자 주소/시군구/연령대'가 부분일치로 잘못 선택되지 않도록 정확한 열명을 첫 후보로 둔다.
    buyer_type_col = first_col(raw, ["매수자", "매수자구분", "매수자 유형", "매수자유형", "거래주체", "매수주체"])
    direct_col = first_col(raw, ["직거래구분", "직거래"])
    rtms_key_col = first_col(raw, ["RTMS_KEY", "실거래키", "거래키"])
    jibun_col = first_col(raw, ["번지", "지번"])
    property_col = first_col(raw, ["거래대상", "물건구분", "건축물주용도", "현실지목", "법정지목"])

    work = raw.copy()
    if contract_date_col:
        work["_contract_date"] = pd.to_datetime(work[contract_date_col], errors="coerce")
    else:
        work["_contract_date"] = pd.NaT
    if date_col:
        work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
    else:
        work["_report_date"] = work["_contract_date"]
    # 계약일이 비면 신고일로, 신고일이 비면 계약일로 보완한다.
    work["_contract_date"] = work["_contract_date"].fillna(work["_report_date"])
    work["_report_date"] = work["_report_date"].fillna(work["_contract_date"])
    if date_col or contract_date_col:
        # `_date`는 계약일 기준 집계(부록)에서, `_ym`은 신고월 기준 집계(본문)에서 사용한다.
        work["_date"] = work["_contract_date"]
        work["_ym"] = work["_report_date"].dt.to_period("M").astype(str)
        work["_ym_contract"] = work["_contract_date"].dt.to_period("M").astype(str)
    else:
        work["_date"] = pd.NaT
        work["_ym"] = ""
        work["_ym_contract"] = ""
    if work["_date"].isna().all():
        raise ValueError(
            "실거래 엑셀에서 계약일·신고일을 읽지 못했습니다. "
            "실거래 칸에 매물조회 등 다른 자료를 지정하지 않았는지 확인하세요.\n"
            "지정 파일: " + ", ".join(path.name for path in paths)
        )
    if amount_col:
        work["_amount"] = work[amount_col].map(as_number).fillna(0)
        source_text = work.get("_source_file", pd.Series("", index=work.index)).astype(str)
        amount_is_manwon = "만원" in str(amount_col) or (
            norm_text(amount_col) == "거래금액" and source_text.str.contains("국토교통부_토지|RTMS", case=False, regex=True, na=False).any()
        )
        if amount_is_manwon:
            work["_amount"] = work["_amount"] * 10_000
    else:
        work["_amount"] = 0
    if area_col:
        work["_area"] = work[area_col].map(as_number)
    else:
        work["_area"] = None
    if emd_col:
        work["_emd"] = work[emd_col].astype(str).str.strip()
    else:
        work["_emd"] = "미상"
    if zone_col:
        work["_zone"] = work[zone_col].astype(str).str.strip()
    else:
        work["_zone"] = "미상"
    if use_col:
        work["_use"] = work[use_col].astype(str).str.strip()
    else:
        work["_use"] = "미상"
    if buyer_region_col:
        work["_buyer_region"] = work[buyer_region_col].astype(str).str.strip()
        local_name = str(region or "").split()[-1] if str(region or "").split() else ""
        valid_buyer = ~work["_buyer_region"].isin(["", "nan", "None", "미상", "-"])
        work["_external"] = valid_buyer & ~work["_buyer_region"].str.contains(local_name, na=False, regex=False) if local_name else False
    else:
        work["_buyer_region"] = "미상"
        work["_external"] = False
    if buyer_type_col:
        work["_buyer_type"] = work[buyer_type_col].astype(str).str.strip()
        work.loc[work["_buyer_type"].isin(["", "nan", "None", "-"]), "_buyer_type"] = "미상"
    else:
        work["_buyer_type"] = "미상"
    if direct_col:
        work["_direct"] = work[direct_col].astype(str).str.contains("직거래", na=False)
    else:
        work["_direct"] = False

    work["_jibun"] = (
        work[jibun_col].astype(str).str.strip().str.replace(r"^'+", "", regex=True)
        if jibun_col else ""
    )
    work["_property"] = work[property_col].astype(str).str.strip() if property_col else work["_use"]
    work["_jimok"] = work[jimok_col].astype(str).str.strip() if jimok_col else ""
    # 확정치 자료에만 있는 개별공시지가(원/㎡)
    gongsi_col = first_col(raw, ["개별공시지가"])
    work["_gongsi"] = work[gongsi_col].map(as_number) if gongsi_col else None
    property_text = work["_property"].astype(str).str.replace(" ", "", regex=False)
    work["_land_only"] = property_text.str.contains("토지만|토지단독", regex=True, na=False) | property_text.eq("토지")
    if rtms_key_col:
        key_text = work[rtms_key_col].astype(str).str.strip()
        valid_key = ~key_text.isin(["", "nan", "None"])
        # RTMS_KEY의 마지막 번호는 다필지 계약의 개별 필지 순번이다.
        # 예: ...-0002419-1~7은 하나의 계약이므로 마지막 순번을 제거해 묶는다.
        contract_key = key_text.str.replace(r"-\d+$", "", regex=True)
        work["_bundle_key"] = contract_key.where(valid_key, work.index.astype(str))
    else:
        work["_bundle_key"] = work.index.astype(str)

    # 계약총금액은 다필지 각 행에 동일 총액이 반복되므로 합계가 아니라 최댓값 1회를 사용한다.
    # 면적은 필지별 면적이므로 계약묶음 안에서 합산한다.
    bundle = work.groupby("_bundle_key", dropna=False).agg(
        _date=("_date", "min"),
        _ym=("_ym", "first"),
        _ym_contract=("_ym_contract", "first"),
        _report_date=("_report_date", "min"),
        _amount=("_amount", "max"),
        _amount_unique_count=("_amount", "nunique"),
        _area=("_area", "sum"),
        _emd=("_emd", "first"),
        _zone=("_zone", "first"),
        _use=("_use", "first"),
        _buyer_region=("_buyer_region", "first"),
        _buyer_type=("_buyer_type", "first"),
        _external=("_external", "max"),
        _direct=("_direct", "max"),
        _jibun=("_jibun", "first"),
        _property=("_property", "first"),
        _jimok=("_jimok", "first"),
        _gongsi=("_gongsi", "max"),
        _land_only=("_land_only", "max"),
        _parcel_rows=("_bundle_key", "size"),
    ).reset_index()

    # 표시 13개월의 첫 달 증감률 계산을 위해 직전 1개월을 내부적으로 추가 집계한다.
    calc_months = [str(x) for x in pd.period_range(pd.Period(base_month) - 13, pd.Period(base_month), freq="M")]
    monthly_calc = []
    for ym in calc_months:
        sub = bundle[bundle["_ym"] == ym]
        ext = sub[sub["_external"]]
        monthly_calc.append({
            "월": ym,
            "거래건수": int(len(sub)),
            "거래금액_억원": round(float(sub["_amount"].sum()) / EOK, 1),
            "외지인거래건수": int(sub["_external"].sum()),
            "외지인거래금액_억원": round(float(ext["_amount"].sum()) / EOK, 1),
            "외지인금액비중": round(float(ext["_amount"].sum()) / float(sub["_amount"].sum()) * 100, 1) if float(sub["_amount"].sum()) else None,
            "직거래건수": int(sub["_direct"].sum()),
        })

    for i, row in enumerate(monthly_calc):
        prev_row = monthly_calc[i - 1] if i else None
        for key, out_key in [
            ("거래건수", "거래건수증감률"),
            ("거래금액_억원", "거래금액증감률"),
            ("외지인거래건수", "외지인건수증감률"),
            ("외지인거래금액_억원", "외지인금액증감률"),
        ]:
            old = prev_row.get(key) if prev_row else None
            row[out_key] = round((row[key] - old) / old * 100, 1) if old not in (None, 0) else None
    monthly = monthly_calc[-13:]

    current = bundle[bundle["_ym"] == base_month]
    prev = bundle[bundle["_ym"] == previous_ym(base_month)]
    yoy_month = str(pd.Period(base_month, freq="M") - 12)
    yoy = bundle[bundle["_ym"] == yoy_month]

    def grouped(col: str, limit: int = 15, source: pd.DataFrame | None = None) -> list[dict[str, Any]]:
        target_df = current if source is None else source
        if target_df.empty:
            return []
        g = target_df.groupby(col, dropna=False).agg(거래건수=("_amount", "size"), 거래금액=("_amount", "sum")).reset_index()
        total_amount = float(g["거래금액"].sum() or 0)
        total_count = int(g["거래건수"].sum() or 0)
        out = []
        for _, row in g.sort_values("거래금액", ascending=False).head(limit).iterrows():
            out.append({
                "명칭": str(row[col]),
                "거래건수": int(row["거래건수"]),
                "건수비중": round(float(row["거래건수"]) / total_count * 100, 1) if total_count else None,
                "거래금액_억원": round(float(row["거래금액"]) / EOK, 1),
                "금액비중": round(float(row["거래금액"]) / total_amount * 100, 1) if total_amount else None,
            })
        return out

    def pct(cur: float, old: float) -> float | None:
        return round((cur - old) / old * 100, 1) if old else None

    emd_top = grouped("_emd", limit=999)
    for row in emd_top:
        info = semiconductor_influence(row.get("명칭"), region)
        if info:
            row["영향권등급"] = info["level"]
            row["영향권라벨"] = info["label"]

    emd_calc_months = [str(x) for x in pd.period_range(pd.Period(base_month) - 13, pd.Period(base_month), freq="M")]
    emd_months = emd_calc_months[-13:]
    emd_monthly_rows: list[dict[str, Any]] = []
    emd_period = bundle[bundle["_ym"].isin(emd_calc_months)]
    if not emd_period.empty:
        emd_grouped = emd_period.groupby(["_emd", "_ym"], dropna=False).agg(
            거래건수=("_amount", "size"),
            거래금액=("_amount", "sum"),
        ).reset_index()
        for emd_name, emd_data in emd_grouped.groupby("_emd", dropna=False):
            by_month = {str(row["_ym"]): row for _, row in emd_data.iterrows()}
            previous_row = by_month.get(emd_calc_months[0])
            monthly_values = []
            for ym in emd_months:
                month_row = by_month.get(ym)
                monthly_values.append({
                    "월": ym,
                    "거래건수": int(month_row["거래건수"]) if month_row is not None else 0,
                    "거래금액_억원": round(float(month_row["거래금액"]) / EOK, 1) if month_row is not None else 0.0,
                })
            info = semiconductor_influence(emd_name, region)
            emd_monthly_rows.append({
                "읍면동": str(emd_name),
                "영향권등급": info["level"] if info else "",
                "영향권라벨": info["label"] if info else "",
                "_previous_month": emd_calc_months[0],
                "_previous_count": int(previous_row["거래건수"]) if previous_row is not None else 0,
                "_previous_amount_eok": round(float(previous_row["거래금액"]) / EOK, 1) if previous_row is not None else 0.0,
                "월별": monthly_values,
                "합계건수": sum(item["거래건수"] for item in monthly_values),
                "합계금액_억원": round(sum(item["거래금액_억원"] for item in monthly_values), 1),
            })
        emd_monthly_rows.sort(key=lambda row: (row["합계금액_억원"], row["합계건수"]), reverse=True)

    current_external = current[current["_external"]]

    def detailed_trade_rows(source: pd.DataFrame, limit: int) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for rank, (_, row) in enumerate(source.sort_values("_amount", ascending=False).head(limit).iterrows(), 1):
            address = " ".join(x for x in [str(row.get("_emd") or ""), str(row.get("_jibun") or "")] if x and x != "nan")
            parcel_count = int(as_number(row.get("_parcel_rows")) or 1)
            if parcel_count > 1:
                address = f"{address} 외({parcel_count - 1:02d}필지)"
            single_land_only = parcel_count == 1 and bool(row.get("_land_only"))
            contract_date = row.get("_date").strftime("%Y-%m-%d") if pd.notna(row.get("_date")) else ""
            report_date = row.get("_report_date").strftime("%Y-%m-%d") if pd.notna(row.get("_report_date")) else ""
            rows.append({
            "순위": rank,
            "계약일": contract_date,
            "신고일": report_date,
            "소재지": address,
            "물건구분": str(row.get("_property") or ""),
            "용도지역": str(row.get("_zone") or ""),
            "지목": str(row.get("_jimok") or ""),
            "면적_㎡": round(float(row.get("_area") or 0), 1),
            "면적_평": round(float(row.get("_area") or 0) * 0.3025, 1),
            "거래금액_억원": round(float(row.get("_amount") or 0) / EOK, 2),
            "단가_원㎡": (
                math.floor((float(row.get("_amount") or 0) / float(row.get("_area") or 0)) / 1_000) * 1_000
                if single_land_only and float(row.get("_area") or 0) > 0 else None
            ),
            "매수자지역": str(row.get("_buyer_region") or ""),
            "거래주체": str(row.get("_buyer_type") or "미상"),
            "외지인": "Y" if bool(row.get("_external")) else "N",
        })
            # 확정치 자료에는 개별공시지가가 있어 실거래 단가와의 배율을 함께 제시한다.
            unit_price = as_number(rows[-1].get("단가_원㎡"))
            gongsi = as_number(row.get("_gongsi"))
            rows[-1]["공시지가_원㎡"] = int(gongsi) if gongsi else None
            rows[-1]["공시지가배율"] = round(unit_price / gongsi, 2) if unit_price and gongsi else None
        return rows

    top_trades = detailed_trade_rows(current, 10)
    external_top_trades = detailed_trade_rows(current_external, 10)

    resolved_as_of = resolve_as_of_date(base_month, as_of_date)
    permit_split = summarize_permit_split(bundle, base_month, region)
    weekly = summarize_weekly(bundle, base_month, region, resolved_as_of)
    emd_month_week = summarize_emd_month_week(bundle, base_month, region, as_of=resolved_as_of)
    period_trend = summarize_period_trend(bundle, base_month, region, as_of=resolved_as_of)

    # 본문(신고일 기준) 영향권 구분 집계 — 6장 권역별 동향 서술에 사용
    influence_group_report_rows: list[dict[str, Any]] = []
    if "광산구" in str(region or ""):
        direct_set = {norm_text(dong) for dong in semiconductor_dongs("direct")}
        is_direct_report = (
            current["_emd"].map(lambda value: norm_text(value) in direct_set)
            if not current.empty else current
        )
        for label, scope_text, frame in [
            ("직접 영향권", "·".join(semiconductor_dongs("direct")),
             current[is_direct_report] if not current.empty else current),
            ("간접 영향권", "직접 영향권 외 읍·면·동",
             current[~is_direct_report] if not current.empty else current),
        ]:
            stats = _bundle_stats(frame)
            total = stats["거래건수"]
            emd_counts = (
                [
                    {"읍면동": str(name), "거래건수": int(count)}
                    for name, count in frame["_emd"].value_counts().head(6).items()
                ] if not frame.empty else []
            )
            influence_group_report_rows.append({
                "구분": label,
                "대상": scope_text,
                "기준": "신고일",
                "기준월": base_month,
                **stats,
                "건수비중": round(total / len(current) * 100, 1) if len(current) else None,
                "금액비중": (
                    round(stats["거래금액_억원"] / (float(current["_amount"].sum()) / EOK) * 100, 1)
                    if not current.empty and float(current["_amount"].sum()) else None
                ),
                "읍면동_상위": emd_counts,
            })

    # 현장점검 양식 서술 단위(직접/간접 영향권)의 당월 거래 집계
    influence_group_month_rows: list[dict[str, Any]] = []
    if "광산구" in str(region or ""):
        direct_names = {norm_text(dong) for dong in semiconductor_dongs("direct")}
        current_contract = bundle[bundle["_ym_contract"] == base_month]
        is_direct_current = (
            current_contract["_emd"].map(lambda value: norm_text(value) in direct_names)
            if not current_contract.empty else current_contract
        )
        for label, scope_text, frame in [
            ("직접 영향권", "·".join(semiconductor_dongs("direct")),
             current_contract[is_direct_current] if not current_contract.empty else current_contract),
            ("간접 영향권", "직접 영향권 외 읍·면·동",
             current_contract[~is_direct_current] if not current_contract.empty else current_contract),
        ]:
            stats = _bundle_stats(frame)
            total_count = stats["거래건수"]

            def share(count: int) -> float | None:
                return round(count / total_count * 100, 1) if total_count else None

            buyer_mix = (
                [
                    {"거래주체": str(name), "건수": int(count), "비중": share(int(count))}
                    for name, count in frame["_buyer_type"].value_counts().items()
                ] if not frame.empty else []
            )
            zone_mix = (
                [
                    {"용도지역": str(name), "건수": int(count), "비중": share(int(count))}
                    for name, count in frame["_zone"].value_counts().head(5).items()
                ] if not frame.empty else []
            )
            land_only_count = int(frame["_land_only"].sum()) if not frame.empty else 0
            corporate_count = int(frame["_buyer_type"].astype(str).str.contains("법인", na=False).sum()) if not frame.empty else 0
            green_count = (
                int(frame["_zone"].astype(str).str.contains("녹지|관리지역|농림|개발제한", na=False, regex=True).sum())
                if not frame.empty else 0
            )
            influence_group_month_rows.append({
                "구분": label,
                "대상": scope_text,
                "기준월": base_month,
                **stats,
                "외지인건수비중": share(stats["외지인거래건수"]),
                "법인매수건수": corporate_count,
                "법인매수비중": share(corporate_count),
                "토지만거래건수": land_only_count,
                "토지만거래비중": share(land_only_count),
                "녹지·관리·개발제한거래건수": green_count,
                "녹지·관리·개발제한비중": share(green_count),
                "거래주체구성": buyer_mix,
                "용도지역구성": zone_mix,
            })

    return {
        "available": True,
        "source_files": [p.name for p in paths],
        "base_month": base_month,
        "data_type": data_type,
        "data_type_note": (
            "부동산원 실거래집계 상세내역(확정치) — 계약해제·승인불가 건 제외, 거래번호 단위 계약묶음"
            if data_type == "확정치" else
            "실거래 신고 원장(잠정치) — RTMS_KEY 마지막 필지순번 제외 계약묶음"
        ),
        "source_file_types": file_types,
        "mixed_format_dropped": [item for item in file_types if item["자료구분"] not in (data_type, "읽기 실패")],
        "date_basis": date_basis,
        "date_basis_column": str(date_col or ""),
        "contract_date_column": str(contract_date_col or ""),
        "appendix_date_basis": "계약일",
        "aggregation_basis": (
            f"본문 통계는 신고일({date_col or '-'}) 기준, 부록(허가구역 전후·주간·현장점검)은 "
            f"계약일({contract_date_col or '-'}) 기준 / RTMS_KEY 마지막 필지순번 제외 계약묶음 단위"
        ),
        "contract_grouping": {
            "key_rule": "RTMS_KEY 마지막 '-필지순번' 제거",
            "amount_rule": "계약총금액은 계약묶음 내 1회(최댓값), 토지면적은 필지별 합산",
            "multi_parcel_contracts_current": int((current["_parcel_rows"] > 1).sum()),
            "amount_conflict_contracts_current": int((current["_amount_unique_count"] > 1).sum()),
        },
        "current": {
            "거래건수": int(len(current)),
            "거래금액_억원": round(float(current["_amount"].sum()) / EOK, 1),
            "외지인거래건수": int(current["_external"].sum()),
            "외지인거래금액_억원": round(float(current_external["_amount"].sum()) / EOK, 1),
            "외지인금액비중": round(float(current_external["_amount"].sum()) / float(current["_amount"].sum()) * 100, 1) if float(current["_amount"].sum()) else None,
            "직거래건수": int(current["_direct"].sum()),
        },
        "previous": {
            "거래건수": int(len(prev)),
            "거래금액_억원": round(float(prev["_amount"].sum()) / EOK, 1),
        },
        "change_vs_previous": {
            "거래건수증감률": pct(len(current), len(prev)),
            "거래금액증감률": pct(float(current["_amount"].sum()), float(prev["_amount"].sum())),
        },
        "year_over_year": {
            "비교월": yoy_month,
            "거래건수": int(len(yoy)),
            "거래금액_억원": round(float(yoy["_amount"].sum()) / EOK, 1),
            "거래건수증감률": pct(len(current), len(yoy)),
            "거래금액증감률": pct(float(current["_amount"].sum()), float(yoy["_amount"].sum())),
        },
        "monthly_rows": monthly,
        "monthly_rows_with_prev": monthly_calc,
        "emd_top": emd_top,
        "emd_months": emd_months,
        "emd_monthly_rows": emd_monthly_rows,
        "zone_top": grouped("_zone"),
        "use_top": grouped("_use"),
        "jimok_top": grouped("_jimok"),
        "external_buyer_region_top": grouped("_buyer_region", source=current_external),
        "external_use_top": grouped("_use", source=current_external),
        "external_top_trades": external_top_trades,
        "top_trades": top_trades,
        "current_trade_rows": detailed_trade_rows(current, len(current)) if len(current) else [],
        "permit_split": permit_split,
        "weekly": weekly,
        "influence_group_month_rows": influence_group_month_rows,
        "influence_group_report_rows": influence_group_report_rows,
        "emd_month_week": emd_month_week,
        "period_trend": period_trend,
    }


def _parse_kosis_land_price(path: Path, region: str) -> dict[str, Any] | None:
    """API 수집기가 저장한 KOSIS 정규화(data 시트) 지가변동률을 읽는다."""
    try:
        df = pd.read_excel(path, sheet_name="data")
    except Exception:
        return None
    period_col = first_col(df, ["기간", "PRD_DE"])
    value_col = first_col(df, ["값", "DT", "DAT"])
    if not period_col or not value_col:
        return None

    name_cols = [c for c in df.columns if c in {
        "항목명", "ITM_NM", "분류명", "분류1", "분류2", "분류3", "분류4",
        "C1_NM", "C2_NM", "C3_NM", "C4_NM", "분류1_객체명", "분류2_객체명",
    } or str(c).upper().endswith("_NM") or re.fullmatch(r"분류\d+명", str(c))]
    work = df.copy()
    if name_cols and region:
        region_tokens = [region, region.replace("전라남도", "").strip(), region.split()[-1]]
        mask = pd.Series(False, index=work.index)
        for col in name_cols:
            text = work[col].astype(str)
            for token in region_tokens:
                if token:
                    mask |= text.str.contains(token, na=False, regex=False)
        if mask.any():
            work = work[mask].copy()
    work["_month"] = work[period_col].astype(str).str.replace(r"[^0-9]", "", regex=True).str[:6]
    work["_value"] = work[value_col].map(as_number)
    work = work[work["_month"].str.match(r"^20\d{4}$", na=False) & work["_value"].notna()]
    if work.empty:
        return None

    zone_names = ["주거", "상업", "공업", "녹지", "보전관리", "생산관리", "계획관리"]
    use_names = ["전", "답", "주거용", "상업용", "임야", "공업용", "기타"]
    records: list[dict[str, Any]] = []
    for _, row in work.iterrows():
        parts = [str(row.get(c, "")).strip() for c in name_cols if str(row.get(c, "")).strip() not in {"", "nan", "None"}]
        compact_parts = [norm_text(x) for x in parts]
        category = None
        for name in use_names:
            if norm_text(name) in compact_parts:
                category = name
                break
        if category is None:
            for name in zone_names:
                aliases = {norm_text(name), norm_text(f"{name}지역")}
                if aliases.intersection(compact_parts):
                    category = name
                    break
        if category is None and any(x in compact_parts for x in ["평균", "전체", "계", "총계"]):
            category = "평균"
        if category:
            records.append({"월": f"{row['_month'][:4]}-{row['_month'][4:6]}", "구분": category, "값": float(row["_value"])})
    if not records:
        return None

    rec = pd.DataFrame(records).groupby(["월", "구분"], as_index=False)["값"].mean()
    pivot = rec.pivot(index="월", columns="구분", values="값").reset_index().sort_values("월").tail(13)
    monthly_rows = []
    for item in pivot.to_dict("records"):
        monthly_rows.append({k: (round(float(v), 3) if k != "월" and pd.notna(v) else (None if k != "월" else v)) for k, v in item.items()})
    latest = monthly_rows[-1] if monthly_rows else {}
    current = {k: latest.get(k) for k in ["평균", *zone_names, *use_names] if latest.get(k) is not None}
    return {
        "available": True,
        "source_file": path.name,
        "source_format": "KOSIS_API_DT_PLCAHTUSE",
        "latest_month": latest.get("월"),
        "current": current,
        "cumulative": {},
        "monthly_rows": monthly_rows,
        "emd_top": [],
    }


def parse_land_price(path: Path, region: str) -> dict[str, Any]:
    normalized = _parse_kosis_land_price(path, region)
    if normalized:
        return normalized
    raw = pd.read_excel(path, header=None)
    # 원표의 요약행은 시군구명(예: 광산구·나주시)으로, 이하 행은 읍면동 명칭으로 표기된다.
    # 시도명이 어떤 형태로 붙어 오더라도 마지막 토큰(시군구명)으로 매칭한다.
    region_text = str(region or "").strip()
    tokens = [t for t in region_text.split() if t]
    region_key = norm_text(tokens[-1] if tokens else region_text)
    if not region_key:
        region_key = norm_text(region_text)
    rows = []
    for idx, row in raw.iterrows():
        cells = [norm_text(v) for v in row.tolist()]
        if any(region_key and region_key in cell for cell in cells):
            rows.append(idx)
    if not rows:
        return {"available": False, "source_file": path.name, "reason": f"{region} 행을 찾지 못함"}

    # 시군구 요약행은 '읍면동' 열(3번째)에도 시군구명이 들어간 행이다.
    # 원표에 따라 요약행 없이 읍면동(기초권역) 행만 실려 오는 경우가 있으므로,
    # 그때는 첫 읍면동 행을 시군구 평균으로 오인하지 않고 '요약행 없음'으로 처리한다.
    district_rows = [
        idx for idx in rows
        if raw.shape[1] > 2 and norm_text(raw.iat[idx, 2]) == region_key
    ]
    district_row_missing = not district_rows
    base_idx = district_rows[0] if district_rows else rows[0]
    col_names = {
        4: "평균", 5: "주거", 6: "상업", 7: "공업", 8: "녹지",
        9: "보전관리", 10: "생산관리", 11: "계획관리",
        15: "전", 16: "답", 17: "주거용", 18: "상업용", 19: "임야", 20: "공업용", 21: "기타",
    }

    def value_row(idx: int) -> dict[str, float | None]:
        out = {}
        for col, name in col_names.items():
            if col < raw.shape[1]:
                v = as_number(raw.iat[idx, col])
                out[name] = round(v, 3) if v is not None else None
        return out

    # 원표 구조는 두 가지다.
    #  (A) 당월변동률 / 전월변동률 / 당월누계  (2026-07 이후 양식)
    #  (B) 변동률 / 누계                      (기존 양식)
    # 헤더 블록(데이터 시작행 위)의 4열 라벨 순서로 구조를 판별한다.
    header_labels: list[str] = []
    for idx in range(max(0, base_idx - 6), base_idx):
        label = norm_text(raw.iat[idx, 4]) if raw.shape[1] > 4 else ""
        if label and any(key in label for key in ["변동률", "누계"]):
            header_labels.append(label)
    if any("전월" in label for label in header_labels):
        row_labels = ["당월", "전월", "누계"]
    else:
        row_labels = ["당월", "누계"]
    stride = len(row_labels)
    offset = {label: index for index, label in enumerate(row_labels)}

    def offset_row(base: int, label: str) -> dict[str, float | None]:
        if label not in offset:
            return {}
        idx = base + offset[label]
        return value_row(idx) if idx < len(raw) else {}

    current = {} if district_row_missing else offset_row(base_idx, "당월")
    previous = {} if district_row_missing else offset_row(base_idx, "전월")
    cumulative = {} if district_row_missing else offset_row(base_idx, "누계")

    def diff_map(a: dict[str, Any], b: dict[str, Any]) -> dict[str, float | None]:
        out: dict[str, float | None] = {}
        for key in a:
            cur, prev = as_number(a.get(key)), as_number(b.get(key))
            out[key] = round(cur - prev, 3) if cur is not None and prev is not None else None
        return out

    change_vs_previous = diff_map(current, previous) if previous else {}

    # 시군구 요약행이 변동률 순위 사이에 끼어 있는 원표도 있으므로
    # 읍면동 행은 요약행 앞뒤를 가리지 않고 표 전체에서 수집한다(요약행·머리글만 제외).
    emd_rows = []
    skip_names = {"명칭", "시군구", "읍면동", "시도", "집중관리", region, region_key}
    for idx in range(len(raw)):
        name = raw.iat[idx, 2] if raw.shape[1] > 2 else None
        name_text = str(name).strip() if name is not None and not pd.isna(name) else ""
        if not name_text or name_text in skip_names or norm_text(name_text) == region_key:
            continue
        if any(suffix in name_text for suffix in ["읍", "면", "동", "리"]):
            def emd_value(label: str) -> float | None:
                if label not in offset:
                    return None
                target = idx + offset[label]
                if target >= len(raw) or raw.shape[1] <= 4:
                    return None
                value = as_number(raw.iat[target, 4])
                return round(value, 3) if value is not None else None

            val = emd_value("당월")
            prev_val = emd_value("전월")
            cum = emd_value("누계")
            emd_rows.append({
                "읍면동": name_text,
                "당월변동률": val,
                "전월변동률": prev_val,
                "누계변동률": cum,
                "전월대비": round(val - prev_val, 3) if val is not None and prev_val is not None else None,
            })
    # 기초단위 권역 전체를 보존한다. 보고서 상위 5개 표시는 렌더링 단계에서 제한하며,
    # 13개월 거래 시계열 집계에는 원표에 수록된 모든 권역이 필요하다.
    emd_rows = sorted(emd_rows, key=lambda x: (x["당월변동률"] is None, -(x["당월변동률"] or -999)))
    # 파일명(전국지가동향_당월_2026_7.xls)에서 조사 기준월을 확보한다.
    survey_month = None
    name_match = re.search(r"(20\d{2})[^0-9]{1,3}(1[0-2]|0?[1-9])(?!\d)", path.stem)
    if name_match:
        survey_month = f"{int(name_match.group(1))}-{int(name_match.group(2)):02d}"
    return {
        "available": True,
        "source_file": path.name,
        "survey_month": survey_month,
        "row_layout": row_labels,
        "district_row_missing": district_row_missing,
        "district_rate_warning": (
            f"{path.name} 원표에 시군구({region_key}) 요약행이 없어 시군구 평균 변동률을 산출하지 못했습니다. "
            "시군구 합계행이 포함된 원표를 올리거나 시군구 변동률을 직접 지정해야 합니다."
            if district_row_missing else ""
        ),
        "current": current,
        "previous": previous,
        "cumulative": cumulative,
        "change_vs_previous": change_vs_previous,
        "has_previous_column": bool(previous),
        "emd_top": emd_rows,
    }


def apply_district_rate_override(summary: dict[str, Any], override: dict[str, Any] | None) -> dict[str, Any]:
    """원표에 시군구 요약행이 없을 때 조사담당자가 확인한 시군구 변동률을 주입한다.

    override 예: {"평균": 0.106, "누계": 0.824}  (용도지역·이용상황 값도 같은 형식으로 가능)
    """
    if not override or not summary.get("available"):
        return summary
    current = dict(summary.get("current") or {})
    cumulative = dict(summary.get("cumulative") or {})
    for key, value in override.items():
        number = as_number(value)
        if number is None:
            continue
        if str(key).startswith("누계"):
            cumulative[str(key).replace("누계", "").strip() or "평균"] = round(number, 3)
        else:
            current[str(key)] = round(number, 3)
    summary = dict(summary)
    summary["current"] = current
    summary["cumulative"] = cumulative
    summary["district_rate_source"] = "조사담당자 직접 입력(원표에 시군구 요약행 없음)"
    summary["district_row_missing"] = False
    summary["district_rate_warning"] = ""
    return summary


def summarize_land_price(paths: list[Path], region: str) -> dict[str, Any]:
    if not paths:
        return {"available": False, "reason": "지가동향 파일 없음"}
    parsed = [parse_land_price(path, region) for path in paths]
    available = [x for x in parsed if x.get("available")]
    if not available:
        return parsed[0]

    # 당월 지가동향 원표(xls)의 읍면동·누계와 KOSIS 최근 13개월 공표통계를 결합한다.
    # 직접 선택한 당월 원표가 먼저 오더라도 KOSIS 월별 시계열이 사라지지 않도록 한다.
    kosis = next((x for x in available if x.get("source_format") == "KOSIS_API_DT_PLCAHTUSE"), None)
    legacy = next((x for x in available if x.get("source_format") != "KOSIS_API_DT_PLCAHTUSE"), None)
    if not kosis:
        return legacy or available[0]
    if not legacy:
        return kosis

    merged = dict(legacy)
    merged["available"] = True
    merged["source_format"] = "LAND_PRICE_XLS+KOSIS_API_DT_PLCAHTUSE"
    merged["source_files"] = [legacy.get("source_file"), kosis.get("source_file")]
    merged["monthly_rows"] = kosis.get("monthly_rows", [])
    merged["latest_public_month"] = kosis.get("latest_month")
    merged["public_current"] = kosis.get("current", {})
    return merged


def _age_group(value: Any) -> str | None:
    text = str(value).strip()
    if text in {"", "nan", "계", "합계", "총계", "전체"}:
        return None
    nums = re.findall(r"\d+", text)
    if not nums:
        return "70대 이상" if "70대 이상" in text else None
    age = int(nums[0])
    if age >= 70:
        return "70대 이상"
    return "0~9세" if age < 10 else f"{age // 10 * 10}대"


def summarize_age_population(paths: list[Path], base_month: str) -> dict[str, Any]:
    """주택동향 주·월간과 같은 8개 연령대, 최근 13개월 구조를 만든다."""
    collected: list[dict[str, Any]] = []
    for path in paths:
        try:
            try:
                df = pd.read_excel(path, sheet_name="data")
            except Exception:
                df = pd.read_excel(path)
            period_col = first_col(df, ["기간", "PRD_DE"])
            value_col = first_col(df, ["값", "DT"])
            if not period_col or not value_col:
                continue
            candidate_cols = [c for c in df.columns if c not in {period_col, value_col}]
            age_col = next((c for c in candidate_cols if df[c].astype(str).str.contains(r"\d+\s*(?:세|~|대)", regex=True, na=False).any()), None)
            if not age_col:
                continue
            work = df.copy()
            work["_month"] = work[period_col].astype(str).str.replace(r"[^0-9]", "", regex=True).str[:6]
            work["_value"] = work[value_col].map(as_number)
            work["_age"] = work[age_col].map(_age_group)
            item_col = first_col(work, ["항목명", "ITM_NM", "분류1"])
            if item_col:
                item_text = work[item_col].astype(str).str.replace(" ", "", regex=False).str.strip()
                # 남자인구수+여자인구수+총인구수를 함께 합산하면 정확히 2배가 되므로
                # 총인구수(T2) 행만 사용한다.
                total_item = item_text.isin(["총인구수", "총인구", "계", "합계", "총계", "전체"])
                item_code_col = first_col(work, ["항목코드", "ITM_ID", "ITM_ID_NM"])
                if item_code_col and work[item_code_col].astype(str).eq("T2").any():
                    total_item = work[item_code_col].astype(str).eq("T2")
                if total_item.any():
                    work = work[total_item]
            for (month, age), values in work.dropna(subset=["_age", "_value"]).groupby(["_month", "_age"])["_value"]:
                if re.match(r"^20\d{4}$", str(month)):
                    collected.append({"month": f"{month[:4]}-{month[4:6:]}", "age": age, "population": float(values.sum())})
        except Exception:
            continue
    if not collected:
        return {"available": False, "reason": "연령대별 주민등록인구 원자료 없음", "monthly_rows": [], "age_rows": []}
    order = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
    dfr = pd.DataFrame(collected).groupby(["month", "age"], as_index=False)["population"].sum()
    piv = dfr.pivot(index="month", columns="age", values="population").fillna(0).reset_index().sort_values("month")
    for col in order:
        if col not in piv.columns:
            piv[col] = 0
    upto = piv[piv["month"] <= base_month]
    if upto.empty:
        upto = piv
    rows14 = upto[["month", *order]].tail(14).to_dict("records")
    rows13 = rows14[-13:]
    latest = rows13[-1]
    first = rows13[0]
    total = sum(float(latest[c]) for c in order)
    age_rows = [{
        "연령대": c,
        "최신월 인구": int(round(float(latest[c]))),
        "비중": round(float(latest[c]) / total * 100, 1) if total else 0,
        "13개월 증감": int(round(float(latest[c]) - float(first[c]))),
    } for c in order]
    return {"available": True, "latest_month": latest["month"], "monthly_rows_with_prev": rows14, "monthly_rows": rows13, "age_rows": age_rows}


def summarize_migration(paths: list[Path], region: str, base_month: str) -> dict[str, Any]:
    """주택동향과 같은 최근 13개월 총전입·총전출·순이동 구조를 만든다."""
    if not paths:
        return {"available": False, "reason": "인구이동 파일 없음", "table_rows": [], "monthly_rows": []}
    collected: list[dict[str, Any]] = []
    source_files: list[str] = []
    for path in paths:
        try:
            try:
                df = pd.read_excel(path, sheet_name="data")
            except Exception:
                df = pd.read_excel(path)
            period_col = first_col(df, ["기간", "PRD_DE"])
            value_col = first_col(df, ["값", "DT", "DAT"])
            if not period_col or not value_col:
                continue
            item_col = next((c for c in df.columns if df[c].astype(str).str.contains("전입|전출|순이동", regex=True, na=False).any()), None)
            if not item_col:
                continue
            work = df.copy()
            # 여러 지역이 함께 든 파일은 선택 시군구 행만 사용한다.
            if region:
                name_cols = [c for c in work.columns if any(k in str(c) for k in ["지역", "시군구", "행정구역", "분류1", "분류2", "C1_NM", "C2_NM"])]
                mask = pd.Series(False, index=work.index)
                for col in name_cols:
                    mask |= work[col].astype(str).str.contains(region.split()[-1], na=False, regex=False)
                if mask.any():
                    work = work[mask].copy()
            work["_month"] = work[period_col].astype(str).str.replace(r"[^0-9]", "", regex=True).str[:6]
            work["_value"] = work[value_col].map(as_number)
            work["_item"] = work[item_col].astype(str).str.strip()
            age_col = next((c for c in work.columns if c not in {period_col, value_col, item_col} and work[c].astype(str).str.contains(r"\d+\s*(?:세|~|대)", regex=True, na=False).any()), None)
            total_work = work
            if age_col:
                age_text = work[age_col].astype(str).str.strip()
                total_rows = work[age_text.isin(["계", "합계", "총계", "전체"])]
                if not total_rows.empty:
                    total_work = total_rows
            for month, sub in total_work.groupby("_month"):
                if not re.match(r"^20\d{4}$", str(month)):
                    continue
                # 총전입·총전출에는 시도간/시군구간 이동이 이미 포함되어 있으므로
                # 부분일치로 세부항목까지 더하면 중복 합산된다. 총계 항목을 우선 사용한다.
                in_rows = sub[sub["_item"].eq("총전입")]
                out_rows = sub[sub["_item"].eq("총전출")]
                if in_rows.empty:
                    in_rows = sub[sub["_item"].str.contains("전입", na=False) & ~sub["_item"].str.contains("순이동", na=False)]
                if out_rows.empty:
                    out_rows = sub[sub["_item"].str.contains("전출", na=False) & ~sub["_item"].str.contains("순이동", na=False)]
                val_in = in_rows["_value"].sum(min_count=1)
                val_out = out_rows["_value"].sum(min_count=1)
                val_net = sub[sub["_item"].str.contains("순이동", na=False)]["_value"].sum(min_count=1)
                if pd.isna(val_net) and pd.notna(val_in) and pd.notna(val_out):
                    val_net = val_in - val_out
                collected.append({
                    "month": f"{month[:4]}-{month[4:6]}",
                    "in": int(round(val_in)) if pd.notna(val_in) else None,
                    "out": int(round(val_out)) if pd.notna(val_out) else None,
                    "net": int(round(val_net)) if pd.notna(val_net) else None,
                })
            source_files.append(path.name)
        except Exception:
            continue
    if not collected:
        return {"available": False, "reason": "전입·전출·순이동 항목을 찾지 못함", "source_files": source_files, "table_rows": [], "monthly_rows": []}
    dfr = pd.DataFrame(collected).groupby("month", as_index=False).agg({"in": "sum", "out": "sum", "net": "sum"}).sort_values("month")
    dfr = dfr[dfr["month"] <= base_month].tail(13)
    rows = dfr.to_dict("records")
    latest = rows[-1] if rows else {}
    return {
        "available": bool(rows), "source_files": source_files, "latest_month": latest.get("month"),
        "in_migration": latest.get("in"), "out_migration": latest.get("out"), "net_migration": latest.get("net"),
        "current": latest.get("net"), "table_rows": rows, "monthly_rows": rows,
    }


def summarize_simple_timeseries(paths: list[Path], kind: str, base_month: str) -> dict[str, Any]:
    if not paths:
        return {"available": False, "reason": f"{kind} 파일 없음"}
    path = paths[0]
    try:
        df = pd.read_excel(path, sheet_name="data")
    except Exception:
        df = pd.read_excel(path)
    period_col = first_col(df, ["기간", "PRD_DE"])
    value_col = first_col(df, ["값", "DT"])
    item_col = first_col(df, ["항목명", "ITM_NM"])
    age_col = first_col(df, ["분류2명", "분류2"])
    if not period_col or not value_col:
        return {"available": False, "source_file": path.name, "reason": "기간/값 컬럼 없음"}

    work = df.copy()
    work["_period"] = work[period_col].astype(str).str.replace(r"[^0-9]", "", regex=True).str[:6]
    work["_value"] = work[value_col].map(as_number)
    target = ym_compact(base_month)

    if kind == "population" and item_col:
        work = work[work[item_col].astype(str).str.contains("총인구|계", regex=True, na=False)]
    elif kind == "migration" and item_col:
        preferred = work[work[item_col].astype(str).str.contains("순이동", regex=False, na=False)]
        if preferred.empty:
            preferred = work[work[item_col].astype(str).str.contains("이동", regex=False, na=False)]
        if not preferred.empty:
            work = preferred

    if kind in {"population", "migration"} and age_col:
        age_text = work[age_col].astype(str).str.strip()
        total_age = work[age_text.isin(["계", "합계", "총계", "전체"])]
        if not total_age.empty:
            work = total_age

    monthly = work.groupby("_period", dropna=False)["_value"].sum().reset_index().sort_values("_period")
    latest = monthly[monthly["_period"] <= target].tail(14)
    rows_with_prev = [{"월": f"{p[:4]}-{p[4:6]}", "값": int(v) if v is not None and not pd.isna(v) else None}
                      for p, v in latest[["_period", "_value"]].itertuples(index=False)]
    rows = rows_with_prev[-13:]
    current = rows[-1]["값"] if rows else None
    prev = rows[-2]["값"] if len(rows) >= 2 else None
    delta = current - prev if current is not None and prev is not None else None

    extra: dict[str, Any] = {}
    if kind == "migration" and age_col and item_col:
        current_rows = work[work["_period"] == target].copy()
        by_age = current_rows.groupby(age_col, dropna=False)["_value"].sum().sort_values(ascending=False).head(8)
        extra["age_top"] = [{"연령대": str(k), "이동자수": int(v)} for k, v in by_age.items()]

    return {
        "available": True,
        "source_file": path.name,
        "current": current,
        "previous": prev,
        "delta": delta,
        "monthly_rows_with_prev": rows_with_prev,
        "monthly_rows": rows,
        **extra,
    }


def summarize_rates(paths: list[Path], base_month: str) -> dict[str, Any]:
    if not paths:
        return {"available": False, "reason": "대출금리 파일 없음"}
    rate_keys = {
        "mortgage_rate": ["주택담보대출"],
        "jeonse_rate": ["전세자금대출"],
        "sme_rate": ["중소기업대출"],
        "business_rate": ["기업일반자금대출", "기업일반자금"],
    }
    monthly: dict[str, dict[str, float]] = {}
    source_files: list[str] = []
    for path in paths:
        try:
            try:
                df = pd.read_excel(path, sheet_name="data")
            except Exception:
                df = pd.read_excel(path)
            period_col = first_col(df, ["기간", "PRD_DE"])
            value_col = first_col(df, ["값", "DT", "DAT"])
            if not period_col or not value_col:
                continue
            item_col = first_col(df, ["항목명", "ITM_NM", "분류1", "분류2", "C1_NM", "C2_NM"])
            if not item_col or not df[item_col].astype(str).str.contains("주택담보대출|전세자금대출|중소기업대출|기업일반자금", regex=True, na=False).any():
                item_col = next((c for c in df.columns if df[c].astype(str).str.contains("주택담보대출|전세자금대출|중소기업대출|기업일반자금", regex=True, na=False).any()), None)
            if not item_col:
                continue
            work = df.copy()
            work["_month"] = work[period_col].astype(str).str.replace(r"[^0-9]", "", regex=True).str[:6]
            work["_value"] = work[value_col].map(as_number)
            item_text = work[item_col].astype(str).str.strip()
            for out_key, keywords in rate_keys.items():
                mask = pd.Series(False, index=work.index)
                for keyword in keywords:
                    exact = item_text.eq(keyword)
                    mask |= exact if exact.any() else item_text.str.contains(keyword, na=False, regex=False)
                selected = work[mask & work["_value"].notna()]
                for month, value in selected.groupby("_month")["_value"].max().items():
                    if re.match(r"^20\d{4}$", str(month)):
                        monthly.setdefault(str(month), {})[out_key] = float(value)
            source_files.append(path.name)
        except Exception:
            continue
    if not monthly:
        return {"available": False, "reason": "주택동향 표준 대출금리 4개 항목을 찾지 못함", "source_files": source_files}
    rows = []
    for month in sorted(monthly):
        if month > ym_compact(base_month):
            continue
        row: dict[str, Any] = {"month": f"{month[:4]}-{month[4:6]}"}
        row.update({key: monthly[month].get(key) for key in rate_keys})
        rows.append(row)
    dfm = pd.DataFrame(rows).sort_values("month")
    for key in rate_keys:
        dfm[f"{key}_mom"] = pd.to_numeric(dfm.get(key), errors="coerce").diff().round(2)
    rows14 = dfm.tail(14).to_dict("records")
    rows13 = rows14[-13:]
    latest = rows13[-1] if rows13 else {}
    return {
        "available": bool(rows13),
        "source_files": source_files,
        "indicator": "예금은행 대출금리(신규취급액 기준)",
        "unit": "%",
        "latest_month": latest.get("month"),
        "mortgage_rate": latest.get("mortgage_rate"),
        "jeonse_rate": latest.get("jeonse_rate"),
        "sme_rate": latest.get("sme_rate"),
        "business_rate": latest.get("business_rate"),
        "current_rate": latest.get("mortgage_rate"),
        "delta_p": latest.get("mortgage_rate_mom"),
        "table_rows_with_prev": rows14,
        "table_rows": rows13,
        "monthly_rows": rows13,
        "interpretation_note": "대출금리는 토지 매수여력과 관망심리의 배경지표로만 사용하고, 지가변동률의 직접 원인으로 단정하지 않는다.",
    }


def _read_listing_file(path: Path) -> pd.DataFrame | None:
    """네이버 매물조회 엑셀(2행 헤더)을 읽는다. 형식이 다르면 None을 반환한다."""
    for header in (1, 0):
        try:
            df = pd.read_excel(path, header=header)
        except Exception:
            continue
        cols = {norm_text(c) for c in df.columns}
        if {"읍면동", "이용상황"} <= cols or {"읍면동", "매매금액(만원)"} <= cols:
            df.columns = [str(c).strip() for c in df.columns]
            return df
    return None


def summarize_listings(paths: list[Path], base_month: str, region: str = "") -> dict[str, Any]:
    """네이버 매물조회 자료를 읍면동·이용상황·용도지역별로 집계하고 전월 대비 호가 조정을 계산한다."""
    if not paths:
        return {
            "available": False,
            "template_only": True,
            "reason": "매물 엑셀 미지정",
            "has_prev_and_current": False,
            "files": [],
        }
    roles = ["전월", "당월"] if len(paths) >= 2 else ["당월"]
    frames: list[tuple[str, Path, Any]] = []
    for index, path in enumerate(paths):
        role = roles[index] if index < len(roles) else f"추가{index}"
        frames.append((role, path, _read_listing_file(path)))
    rows = [
        {
            "구분": role,
            "source_file": path.name,
            "estimated_listing_count": int(len(df)) if df is not None else None,
            "read_error": None if df is not None else "매물 표준 형식(읍면동·이용상황) 아님",
        }
        for role, path, df in frames
    ]
    parsed = [(role, path, df) for role, path, df in frames if df is not None]
    if not parsed:
        return {
            "available": True,
            "template_only": True,
            "base_month": base_month,
            "files": rows,
            "quality_note": "매물 파일 형식을 인식하지 못해 건수만 표시함.",
            "has_prev_and_current": False,
        }

    direct_names = {norm_text(dong) for dong in semiconductor_dongs("direct")} if "광산구" in str(region or "") else set()
    current_df = parsed[-1][2]
    prev_entry = next((item for item in parsed if item[0] == "전월"), None)
    prev_df = prev_entry[2] if prev_entry else None

    def prep(df):
        work = df.copy()
        def col(name, default=""):
            return work[name] if name in work.columns else pd.Series(default, index=work.index)
        work["_emd"] = col("읍면동").astype(str).str.strip()
        work["_use"] = col("이용상황", "미상").astype(str).str.strip()
        work["_zone"] = col("용도지역", "미상").astype(str).str.strip()
        work["_jimok"] = col("지목").astype(str).str.strip()
        work["_price"] = col("매매금액(만원)", None).map(as_number)
        work["_unit"] = col("매매금액단가(원)", None).map(as_number)
        # 네이버 매물 고유번호는 월마다 재발급되어 전월과 매칭되지 않으므로
        # 소재지(읍면동·지번)+이용상황+대지면적을 동일 매물 판정 키로 사용한다.
        work["_key"] = (
            work["_emd"] + "|" + col("지번").astype(str).str.strip()
            + "|" + work["_use"] + "|" + col("대지면적", "").astype(str).str.strip()
        )
        work["_listing_id"] = col("고유번호").astype(str).str.strip()
        return work

    cur = prep(current_df)
    prev = prep(prev_df) if prev_df is not None else None

    def counts(work, column):
        if work is None or work.empty:
            return {}
        return {str(k): int(v) for k, v in work[column].value_counts().items()}

    def build_rows(column: str, label: str, limit: int) -> list[dict[str, Any]]:
        cur_counts = counts(cur, column)
        prev_counts = counts(prev, column) if prev is not None else {}
        out = []
        # 매물 수가 같으면 이름순으로 두어 실행할 때마다 표 순서가 바뀌지 않게 한다.
        for name in sorted(set(cur_counts) | set(prev_counts), key=lambda x: (-cur_counts.get(x, 0), x)):
            subset = cur[cur[column] == name]
            unit = subset["_unit"].dropna()
            out.append({
                label: name,
                "당월_매물수": cur_counts.get(name, 0),
                "전월_매물수": prev_counts.get(name, 0) if prev is not None else None,
                "증감": (cur_counts.get(name, 0) - prev_counts.get(name, 0)) if prev is not None else None,
                "평균단가_원㎡": int(unit.mean()) if not unit.empty else None,
                "중위단가_원㎡": int(unit.median()) if not unit.empty else None,
                "영향권라벨": "반도체 직접 영향권" if norm_text(name) in direct_names else "",
            })
        return out[:limit]

    price_changes: dict[str, Any] = {}
    emd_changes: dict[str, list[float]] = {}
    if prev is not None:
        cur_price = cur[cur["_key"].astype(bool)].dropna(subset=["_price"]).groupby("_key")["_price"].min()
        prev_price = prev[prev["_key"].astype(bool)].dropna(subset=["_price"]).groupby("_key")["_price"].min()
        common = cur_price.index.intersection(prev_price.index)
        diff = cur_price[common] - prev_price[common]
        detail = cur.drop_duplicates("_key").set_index("_key")
        for key in common:
            base_price = float(prev_price[key] or 0)
            emd_changes.setdefault(str(key).split("|")[0], []).append(float(diff[key]) / base_price * 100 if base_price else 0.0)
        samples = []
        for key in diff[diff != 0].abs().sort_values(ascending=False).head(8).index:
            if key not in detail.index:
                continue
            item = detail.loc[key]
            samples.append({
                "읍면동": str(item.get("_emd") or ""),
                "이용상황": str(item.get("_use") or ""),
                "용도지역": str(item.get("_zone") or ""),
                "지목": str(item.get("_jimok") or ""),
                "전월_호가_만원": int(prev_price[key]),
                "당월_호가_만원": int(cur_price[key]),
                "변동_만원": int(diff[key]),
                "변동률": round(float(diff[key]) / float(prev_price[key]) * 100, 1) if prev_price[key] else None,
            })
        price_changes = {
            "매칭기준": "읍면동+지번+이용상황+대지면적",
            "비교가능_매물수": int(len(common)),
            "호가인상": int((diff > 0).sum()),
            "호가인하": int((diff < 0).sum()),
            "호가유지": int((diff == 0).sum()),
            "평균변동률": round(float((diff / prev_price[common] * 100).mean()), 1) if len(common) else None,
            "신규등록": int(len(cur_price.index.difference(prev_price.index))),
            "매물회수": int(len(prev_price.index.difference(cur_price.index))),
            "조정사례": samples,
        }

    direct_summary: dict[str, Any] = {}
    if direct_names:
        cur_direct = cur[cur["_emd"].map(lambda x: norm_text(x) in direct_names)]
        prev_direct = prev[prev["_emd"].map(lambda x: norm_text(x) in direct_names)] if prev is not None else None
        unit = cur_direct["_unit"].dropna()
        direct_summary = {
            "당월_매물수": int(len(cur_direct)),
            "전월_매물수": int(len(prev_direct)) if prev_direct is not None else None,
            "증감": int(len(cur_direct) - len(prev_direct)) if prev_direct is not None else None,
            "당월비중": round(len(cur_direct) / len(cur) * 100, 1) if len(cur) else None,
            "평균단가_원㎡": int(unit.mean()) if not unit.empty else None,
        }

    cur_emd_counts = counts(cur, "_emd")
    prev_emd_counts = counts(prev, "_emd") if prev is not None else {}
    emd_price_rows = [
        {
            "읍면동": name,
            "당월_매물수": cur_emd_counts.get(name, 0),
            "전월_매물수": prev_emd_counts.get(name, 0) if prev is not None else None,
            "비교가능": len(emd_changes.get(name, [])),
            "호가인상": sum(1 for value in emd_changes.get(name, []) if value > 0),
            "호가인하": sum(1 for value in emd_changes.get(name, []) if value < 0),
            "평균변동률": round(sum(emd_changes[name]) / len(emd_changes[name]), 1) if emd_changes.get(name) else None,
            "이용상황별": {str(k): int(v) for k, v in cur[cur["_emd"] == name]["_use"].value_counts().head(3).items()},
        }
        for name in sorted(set(cur_emd_counts) | set(prev_emd_counts))
    ]
    total_prev = int(len(prev)) if prev is not None else None
    change = {}
    if total_prev is not None:
        change = {
            "전월_매물수": total_prev,
            "당월_매물수": int(len(cur)),
            "증감": int(len(cur) - total_prev),
            "증감률": round((len(cur) - total_prev) / total_prev * 100, 1) if total_prev else None,
        }
    return {
        "available": True,
        "template_only": False,
        "base_month": base_month,
        "source": "네이버 매물조회(매매)",
        "files": rows,
        "total_current": int(len(cur)),
        "total_previous": total_prev,
        "change_vs_previous": change,
        "use_rows": build_rows("_use", "이용상황", 12),
        "emd_rows": build_rows("_emd", "읍면동", 15),
        "zone_rows": build_rows("_zone", "용도지역", 10),
        "price_changes": price_changes,
        "emd_price_rows": emd_price_rows,
        "direct_zone_summary": direct_summary,
        "quality_note": "전월·당월 2개 매물 파일이 있어야 매물 증감과 호가 조정 판단을 산출할 수 있다.",
        "has_prev_and_current": prev is not None,
    }


def existing_paths(paths: list[str | Path] | None) -> list[Path]:
    result: list[Path] = []
    for item in paths or []:
        if not item:
            continue
        path = Path(item)
        if path.exists() and path.is_file():
            result.append(path)
    return result


GWANGJU_GU_NAMES = ["광산구", "동구", "서구", "남구", "북구"]
JEONNAM_CITY_COUNTY_NAMES = [
    "목포시", "여수시", "순천시", "나주시", "광양시",
    "담양군", "곡성군", "구례군", "고흥군", "보성군",
    "화순군", "장흥군", "강진군", "해남군", "영암군",
    "무안군", "함평군", "영광군", "장성군", "완도군",
    "진도군", "신안군",
]
DIRECT_REGIONAL_ISSUE_NAMES = ["광주", "광산", "나주", "장성", "화순"]


def regional_issue_display_area(region: str) -> str:
    region_text = str(region or "")
    for gu in GWANGJU_GU_NAMES:
        if gu in region_text:
            return f"전남광주통합특별시 {gu}"
    for name in JEONNAM_CITY_COUNTY_NAMES:
        if name in region_text:
            return f"전남광주통합특별시 {name}"
    if "광주" in region_text:
        return "전남광주통합특별시 광주권역"
    return "전남광주통합특별시 광주권역"


def build_regional_issues(region: str, base_month: str) -> dict[str, Any]:
    region_text = str(region or "")
    display_area = regional_issue_display_area(region)
    # 외부 regional_profiles.json에 개발·규제 이슈가 정의된 시군구는 그 값을 우선 사용한다.
    entry = _match_region_entry(region)
    if entry and isinstance(entry.get("issues"), dict):
        iss = entry["issues"]
        # 전남광주 밖 시군구는 issues.display_area로 표기 권역을 직접 지정할 수 있다.
        display_area = iss.get("display_area") or display_area
        direct = bool(iss.get("is_direct_issue_area", True))
        issue_scope_note = "직접 영향권" if direct else "광역 배경권"
        built = {
            "available": True,
            "display_area": display_area,
            "issue_scope_note": issue_scope_note,
            "base_month": base_month,
            "events": iss.get("events", []),
            "report_wording_rule": iss.get(
                "report_wording_rule",
                "보고서에서는 result.json의 regional_issues.display_area 값을 그대로 사용한다.",
            ),
            "sources": iss.get("sources", []),
            "source_rule": iss.get("source_rule", ""),
            "regulation_timing_note": iss.get("regulation_timing_note", ""),
        }
        return _apply_issue_placeholders(built, display_area, issue_scope_note)
    is_jeonnam_or_gwangju = (
        any(keyword in region_text for keyword in ["광주", "광산"])
        or any(name in region_text for name in JEONNAM_CITY_COUNTY_NAMES)
    )
    if not is_jeonnam_or_gwangju:
        return {"available": False, "reason": "전남광주통합특별시 광주권역 반도체 클러스터 직접 영향권 외 지역"}
    direct_issue_area = any(keyword in region_text for keyword in DIRECT_REGIONAL_ISSUE_NAMES)
    issue_scope_note = "직접 영향권" if direct_issue_area else "광역 배경권"
    return {
        "available": True,
        "display_area": display_area,
        "issue_scope_note": issue_scope_note,
        "base_month": base_month,
        "events": [
            {
                "title": "호남권 반도체 첨단국가산단 광주 군공항 부지 낙점",
                "event_date": "2026-07-06",
                "project_stage": "부지 낙점·투자 발표 단계 — 광주 군공항 부지(약 250만평)가 후보지로 발표되고 대규모 투자 계획이 공표됨. 국가산단 지정·착공, 군공항 이전, 전력·용수 등 기반 확보는 절차 진행 예정",
                "summary": f"{display_area}은 광주 군공항 부지 중심의 호남권 반도체 첨단국가산단 조성 이슈를 {issue_scope_note} 관점에서 점검할 필요가 있다.",
                "interpretation_caution": "부지 낙점·투자 발표는 확인되나 국가산단 지정·착공·군공항 이전은 절차가 진행 중이므로, 개발 기대를 지가변동률 상승의 직접 원인으로 단정하지 않는다.",
            },
            {
                "title": "광주 군공항 이전 예비이전후보지 선정",
                "event_date": "2026-04-02",
                "project_stage": "예비이전후보지(전남 무안군 망운면 일대) 선정 — 주민투표·유치신청을 거쳐 이전부지 최종 선정 예정, 한미 공동운영기지 협의가 변수",
                "summary": "반도체 산단 조기 착공과 연계해 군공항 이전 논의가 진행되며, 군공항 종전부지 개발 시기는 이전 절차 진행에 연동된다.",
                "interpretation_caution": "이전부지 최종 선정과 종전부지 개발 시기가 확정되지 않았으므로 기대를 실수요로 간주하지 않는다.",
            },
            {
                "title": "토지거래허가구역 지정 효력 발생",
                "event_date": "2026-07-14",
                "project_stage": "고시·효력 발생(확정)",
                "effective_period": "2026-07-14~2028-07-13",
                "summary": f"{display_area} 관련 토지거래허가구역 지정 이슈가 효력을 갖게 되며(예정지 일원 약 364㎢), 일정 면적 초과 토지 거래는 사전 허가와 이용 의무가 수반된다.",
                "interpretation_caution": "거래 규제는 투기성 수요 억제 요인이므로, 시행 전후 거래량·외지인 거래·매물 호가 변화를 분리해 관찰한다.",
            },
            {
                "title": "광주송정역 KTX 투자선도지구 개발",
                "event_date": "2021(투자선도지구 지정)",
                "project_stage": "투자선도지구 지정(2021, 국토교통부) 이후 지구계획 변경·공공개발용 토지비축사업 진행 중 — 토지보상·기반시설 조성 진행률 확인 필요",
                "summary": "광주송정역 주변 약 55만㎡를 지역경제 거점으로 개발하는 사업으로 LH가 사업시행자로 참여하며, 송정역 복합환승 기능과 업무·상업·주거·생활시설 확충이 핵심이다.",
                "interpretation_caution": "지구계획 변경·토지비축·보상 등 후속 절차가 진행 중이므로 개발 기대를 당월 지가 상승의 직접 원인으로 단정하지 않고 토지보상·기반시설 조성 진행률을 확인하며 판단한다.",
            },
            {
                "title": "어등산관광단지·그랜드 스타필드 광주 조성",
                "event_date": "2030~2033(단계 완공 계획)",
                "project_stage": "유원지 부지 세부설계·조성계획 변경 후 기반시설 공사 추진 — 그랜드 스타필드 2030년, 숙박시설 등 2033년 단계적 완공 계획(광주시 공개자료). 실제 착공·공정률 확인 필요",
                "summary": "어등산 일대에 복합쇼핑몰과 관광·휴양·숙박시설을 조성하는 대규모 민간개발사업으로, 운수동·서봉동·산정동과 수완·선운 생활권의 상업·관광 수요에 영향을 줄 수 있다.",
                "interpretation_caution": "민간개발 사업은 착공·공정률에 따라 실현 여부가 좌우되므로 계획 일정을 확정 사실로 보지 않고 실제 착공과 공정률을 중심으로 판단한다.",
            },
            {
                "title": "광주 산정 공공주택지구 조성",
                "event_date": "2023(지구 지정)",
                "project_stage": "공공주택지구 지정 완료(2023) — LH 도시건축통합계획 설계공모(2024), 민·관·공 협의체 출범(2025). 지구계획 승인·보상·착공(LH 자료상 2027년 착공 예정) 등 후속 절차 진행 중이며, 실제 승인·보상 일정은 최신 고시 기준 재확인 필요",
                "summary": "광산구 산정동·장수동 일원 약 168만㎡에 약 1만2,875세대를 공급하는 신도시급 공공주택사업으로, 산정동·장수동과 하남동·흑석동·장덕동 및 어등산·선운 생활권과 연계될 가능성이 있다.",
                "interpretation_caution": "완성된 택지지구가 아니라 지구계획 확정·보상·착공을 확인해야 하는 추진사업이며, 광주지역 주택공급 과잉 우려와 사업 규모 조정 요구도 존재하므로 개발 기대를 당월 지가 상승의 직접 원인으로 적용하기 어렵다.",
            },
        ],
        "report_wording_rule": "보고서에서는 result.json의 regional_issues.display_area 값을 그대로 사용한다.",
        "sources": [
            {"name": "국토교통부 공고(호남 반도체 첨단국가산단 예정지 일원 토지거래허가구역 지정)", "type": "정부 공고", "date": "2026-07-09", "note": "지정기간 2026-07-14~2028-07-13, 예정지 일원 약 364.19㎢. 공고번호는 최신 원문에서 확인 후 표기"},
            {"name": "정부 반도체 클러스터 입지 발표(광주 군공항 부지)", "type": "정부 발표·언론 보도", "date": "2026-07-06", "note": "정부·지자체 원문(보도자료·브리핑) 확인 후 표기"},
            {"name": "국방부 광주 군공항 예비이전후보지 선정", "type": "정부 발표", "date": "2026-04-02"},
            {"name": "국토교통부 광주송정역 투자선도지구 지정 및 지구계획(LH 시행)", "type": "정부 지정·공공사업", "date": "2021", "note": "지구계획 변경·토지비축 진행상황은 최신 고시 확인"},
            {"name": "어등산관광단지·그랜드 스타필드 광주 조성계획(광주시 공개자료)", "type": "지자체 공개자료·민간개발", "note": "단계 완공 2030/2033 계획, 착공·공정률 확인 필요"},
            {"name": "광주 산정 공공주택지구 지정·LH 사업실명제 자료", "type": "정부 지정·공공사업", "date": "2023", "note": "지구계획 승인·보상·2027년 착공 예정 일정은 최신 고시 기준 재확인"},
        ],
        "source_rule": "정부·지자체 원문(공고·보도자료·브리핑)을 우선하고 포털 재전송 페이지는 출처로 사용하지 않는다. 부지 낙점·투자 발표는 확인된 사실로 서술하되, 국가산단 지정·착공·군공항 이전 등 후속 절차는 '진행 예정'으로 구분해 표현한다.",
        "regulation_timing_note": "실거래는 계약일 기준이므로 2026년 7월 14일 이전 계약분은 규제 시행 이후 시장반응으로 해석하지 않는다. 규제 전후 비교는 transactions.permit_split의 지정 전(7월 1~13일)·지정 후(7월 14~31일) 구간과 일평균으로 하되, 신고기한(30일)에 따른 최근 구간 과소집계를 감안해 다음 갱신분에서 재확인한다.",
    }


NEWS_RSS_URL = "https://news.google.com/rss/search"
NEWS_MAX_STORIES_PER_PROJECT = 10
NEWS_HEADLINE_COUNT = 3
NEWS_GENERAL_TOPICS = ("부동산", "개발사업", "아파트 분양", "토지", "재개발")
NEWS_REAL_ESTATE_WORDS = (
    "부동산", "개발", "분양", "아파트", "토지", "산단", "산업단지", "지구", "재개발", "재건축", "택지",
    "주택", "허가구역", "지가", "땅값", "착공", "준공", "군공항", "반도체", "역세권", "입주", "매매", "전세",
)
# 포털 재전송 페이지는 출처로 쓰지 않는다(보고서 news_rules와 동일 기준).
NEWS_PORTAL_SOURCES = ("daum", "다음", "naver", "네이버", "msn", "zum", "네이트", "nate")
NEWS_PUBLIC_SOURCE_KEYWORDS = (
    "정책브리핑", "국토교통부", "국방부", "산업통상", "기획재정부", "광주광역시", "전남광주통합특별시",
    "전라남도", "광산구", "한국토지주택공사", "korea.kr",
)
NEWS_WIRE_SOURCES = ("연합뉴스", "뉴시스", "뉴스1")


def _development_project_paths() -> list[Path]:
    """편집용 개발사업 목록(development_projects.json) 탐색 경로."""
    paths: list[Path] = []
    env = os.environ.get("LAND_PRICE_DEVELOPMENT_PROJECTS")
    if env:
        paths.append(Path(env))
    paths.append(Path.cwd() / "development_projects.json")
    try:
        paths.append(Path(__file__).resolve().parent / "development_projects.json")
    except NameError:
        pass
    return paths


def _match_development_region(region: str) -> dict[str, Any]:
    """development_projects.json에서 region에 해당하는 시군구 항목(가장 긴 match_key 우선)을 찾는다."""
    region_text = str(region or "")
    for path in _development_project_paths():
        try:
            if not path.is_file():
                continue
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        regions = data.get("regions") if isinstance(data, dict) else None
        if not isinstance(regions, dict):
            continue
        best: tuple[int, dict[str, Any]] | None = None
        for key, entry in regions.items():
            if not isinstance(entry, dict) or not isinstance(entry.get("projects"), list):
                continue
            for match_key in (entry.get("match_keys") or [key]):
                match_key = str(match_key)
                if match_key and match_key in region_text and (best is None or len(match_key) > best[0]):
                    best = (len(match_key), entry)
        return best[1] if best else {}
    return {}


def load_development_projects(region: str) -> list[dict[str, Any]]:
    entry = _match_development_region(region)
    return [project for project in entry.get("projects", []) if isinstance(project, dict) and project.get("name")]


def attach_headline_summaries(headlines: list[dict[str, Any]], region: str, base_month: str) -> None:
    """목록 파일의 headline_summaries(담당자가 원문을 확인해 적은 500자 이내 요약)를 제목이 비슷한 기사에 붙인다."""
    def key_of(title: Any) -> str:
        return re.sub(r"[^0-9A-Za-z가-힣]", "", str(title or ""))

    curated = [
        entry for entry in (_match_development_region(region).get("headline_summaries") or [])
        if isinstance(entry, dict) and str(entry.get("month") or "") == str(base_month) and str(entry.get("text") or "").strip()
    ]
    for story in headlines:
        for entry in curated:
            if SequenceMatcher(None, key_of(story.get("title")), key_of(entry.get("title"))).ratio() >= 0.5:
                story["summary"] = str(entry["text"]).strip()
                break


def development_news_window(base_month: str) -> tuple[date, date, list[str]]:
    """수집 기간: 기준월의 전월 1일 ~ 기준월 말일(예: 9월 보고서는 8~9월)."""
    prev = previous_ym(base_month)
    py, pm = [int(x) for x in prev.split("-")]
    by, bm = [int(x) for x in base_month.split("-")]
    return date(py, pm, 1), date(by, bm, calendar.monthrange(by, bm)[1]), [prev, base_month]


def _news_source_type(source: str) -> str:
    text = str(source or "")
    if text.strip() == "LH" or any(key in text for key in NEWS_PUBLIC_SOURCE_KEYWORDS):
        return "공공·보도자료"
    if any(key in text for key in NEWS_WIRE_SOURCES):
        return "통신"
    return "언론"


def _fetch_news_rss(query: str, start: date, end: date, timeout: int = 15) -> list[dict[str, Any]]:
    """뉴스 검색 RSS에서 기간 내 기사 제목·매체·보도일을 읽는다(본문은 수집하지 않음)."""
    full_query = f"{query} after:{start - timedelta(days=1):%Y-%m-%d} before:{end + timedelta(days=1):%Y-%m-%d}"
    url = NEWS_RSS_URL + "?" + urllib.parse.urlencode({"q": full_query, "hl": "ko", "gl": "KR", "ceid": "KR:ko"})
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    nodes: list[Any] = []
    # 검색 RSS는 같은 검색어에도 간헐적으로 빈 목록을 돌려주므로 0건이면 한 번 더 요청한다.
    for attempt in range(2):
        with urllib.request.urlopen(request, timeout=timeout) as response:
            nodes = ET.fromstring(response.read()).findall("./channel/item")
        if nodes:
            break
        time.sleep(2)
    items: list[dict[str, Any]] = []
    for node in nodes:
        source = (node.findtext("source") or "").strip()
        title = (node.findtext("title") or "").strip()
        if source and title.endswith(f" - {source}"):
            title = title[: -len(source) - 3].strip()
        title = re.sub(r"\s*>\s*[^>]{1,10}$", "", title)
        # 여러 소식을 묶은 단신 모음은 개별 사업 진척 근거로 쓰기 어려워 제외한다.
        if re.search(r"外\s*$", title) or re.match(r"\[[^\]]*(소식|로컬|단신|브리핑)\]", title):
            continue
        try:
            published = (parsedate_to_datetime(node.findtext("pubDate") or "") + timedelta(hours=9)).date()
        except Exception:
            continue
        link = (node.findtext("link") or "").strip()
        if not title or not start <= published <= end:
            continue
        if any(portal in source.lower() for portal in NEWS_PORTAL_SOURCES) or not re.search(r"[가-힣]", source):
            continue
        items.append({
            "date": f"{published:%Y-%m-%d}",
            "month": f"{published:%Y-%m}",
            "source": source,
            "source_type": _news_source_type(source),
            "title": title,
            "url": link if link.startswith(("http://", "https://")) else "",
        })
    return items


def _news_title_matches(title: str, must_include: Any) -> bool:
    text = norm_text(title)
    groups = [group for group in (must_include or []) if isinstance(group, list) and group]
    return all(any(norm_text(token) in text for token in group) for group in groups)


def _group_similar_news(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """같은 사안을 여러 매체가 보도한 경우 대표 1건으로 묶고 유사 보도 건수를 남긴다."""
    tier = {"공공·보도자료": 0, "통신": 1, "언론": 2}
    ordered = sorted(items, key=lambda item: (tier.get(item["source_type"], 9), item["date"]))
    stories: list[dict[str, Any]] = []
    for item in ordered:
        key = re.sub(r"[^0-9A-Za-z가-힣]", "", item["title"])
        day = date.fromisoformat(item["date"])
        matched = None
        for story in stories:
            if abs((date.fromisoformat(story["date"]) - day).days) > 2:
                continue
            if SequenceMatcher(None, key, story["_key"]).ratio() >= 0.5:
                matched = story
                break
        if matched:
            matched["similar_count"] += 1
            continue
        stories.append({**item, "similar_count": 0, "_key": key})
    # 여러 매체가 함께 다룬 사안을 먼저 둔다(보도 폭 → 출처 등급 → 최신순).
    stories.sort(key=lambda story: story["date"], reverse=True)
    stories.sort(key=lambda story: (-story["similar_count"], tier.get(story["source_type"], 9)))
    return [{k: v for k, v in story.items() if k != "_key"} for story in stories]


def _pick_monthly_headlines(stories: list[dict[str, Any]], count: int = NEWS_HEADLINE_COUNT) -> list[dict[str, Any]]:
    """기준월 주요 뉴스: 보도 폭이 큰 순으로 고르되 한 사업에 쏠리지 않게 사업당 1건을 우선한다."""
    picked: list[dict[str, Any]] = []
    used_projects: set[str] = set()
    for story in stories:
        related = str(story.get("related_project") or "")
        if related and related in used_projects:
            continue
        picked.append(story)
        used_projects.add(related)
        if len(picked) == count:
            return picked
    for story in stories:
        if story not in picked:
            picked.append(story)
            if len(picked) == count:
                break
    return picked


def _merge_news_cache(cache_path: Path | None, raw: dict[int, list[dict[str, Any]]], names: dict[int, str]) -> None:
    """이전 실행에서 수집한 기사와 합친다(검색 RSS가 간헐적으로 빈 목록을 주므로 실행할수록 누적)."""
    if cache_path is None:
        return
    try:
        cached = json.loads(cache_path.read_text(encoding="utf-8")) if cache_path.is_file() else {}
    except Exception:
        cached = {}
    merged: dict[str, list[dict[str, Any]]] = {}
    for index, name in names.items():
        seen: set[tuple[str, str, str]] = set()
        items: list[dict[str, Any]] = []
        for item in raw[index] + [old for old in cached.get(name, []) if isinstance(old, dict)]:
            key = (str(item.get("source")), str(item.get("title")), str(item.get("date")))
            if key not in seen and all(item.get(field) for field in ("source", "title", "date", "month")):
                seen.add(key)
                items.append(item)
        raw[index] = items
        merged[name] = items
    try:
        cache_path.write_text(json.dumps(merged, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass


def build_development_projects(region: str, base_month: str, cache_dir: Path | None = None) -> dict[str, Any]:
    """시군구 주요 개발사업 목록과 뉴스·보도자료 제목(사업별 전월~기준월, 지역 주요 뉴스는 기준월)을 묶는다.

    수집에 실패해도 보고서 생성은 계속되도록 오류는 errors에만 기록한다.
    cache_dir를 주면 development_news_YYYYMM.json에 수집분을 누적한다.
    """
    projects = load_development_projects(region)
    start, end, months = development_news_window(base_month)
    base_start = date(end.year, end.month, 1)
    district = str(region or "").split()[-1] if str(region or "").split() else ""
    skip_fetch = str(os.environ.get("LAND_PRICE_SKIP_NEWS") or "").strip() in {"1", "true", "TRUE"}
    jobs = [
        (index, str(query))
        for index, project in enumerate(projects)
        for query in (project.get("queries") or [project.get("name")])
        if str(query or "").strip()
    ]
    # index -1: 사업 목록과 무관한 지역 부동산 일반 뉴스(기준월 주요 뉴스 후보)
    if district:
        jobs += [(-1, f"{district} {topic}") for topic in NEWS_GENERAL_TOPICS]
    raw: dict[int, list[dict[str, Any]]] = {index: [] for index in range(-1, len(projects))}
    errors: list[str] = []
    if skip_fetch:
        errors.append("LAND_PRICE_SKIP_NEWS 설정으로 뉴스 수집을 건너뜀")
    else:
        def run(job: tuple[int, str]) -> tuple[int, str, list[dict[str, Any]] | None, str]:
            try:
                return job[0], job[1], _fetch_news_rss(job[1], base_start if job[0] < 0 else start, end), ""
            except Exception as exc:
                return job[0], job[1], None, f"{type(exc).__name__}: {exc}"

        # 동시에 요청하면 일부 검색어가 빈 목록으로 돌아오므로 한 건씩 차례로 요청한다.
        for index, query, items, error in map(run, jobs):
            if items is None:
                owner = projects[index].get("name") if index >= 0 else "지역 주요 뉴스"
                errors.append(f"{owner} | 검색어 '{query}' 수집 실패 | {error}")
            else:
                raw[index].extend(items)
    cache_path = None
    if cache_dir and Path(cache_dir).is_dir():
        cache_path = Path(cache_dir) / f"development_news_{safe_name(region)}_{ym_compact(base_month)}.json"
    _merge_news_cache(
        cache_path, raw,
        {-1: "__지역 주요 뉴스__", **{index: str(project.get("name")) for index, project in enumerate(projects)}},
    )
    built: list[dict[str, Any]] = []
    headline_pool: dict[tuple[str, str], dict[str, Any]] = {}
    for index, project in enumerate(projects):
        seen: set[tuple[str, str]] = set()
        matched: list[dict[str, Any]] = []
        for item in raw[index]:
            key = (item["source"], item["title"])
            if key in seen or not _news_title_matches(item["title"], project.get("must_include")):
                continue
            seen.add(key)
            matched.append(item)
            if item["month"] == base_month:
                headline_pool.setdefault(key, {**item, "related_project": project.get("name")})
        # 전월·기준월이 모두 드러나도록 월별로 보도 폭이 큰 사안부터 절반씩 고르고 남는 자리는 나머지로 채운다.
        stories = _group_similar_news(matched)
        per_month = NEWS_MAX_STORIES_PER_PROJECT // len(months)
        news = [story for month in months for story in [s for s in stories if s["month"] == month][:per_month]]
        news += [story for story in stories if story not in news][:NEWS_MAX_STORIES_PER_PROJECT - len(news)]
        news.sort(key=lambda story: story["date"], reverse=True)
        built.append({
            **{k: project.get(k, "") for k in [
                "name", "registered", "type", "influence_area", "period", "scale", "cost",
                "operator", "representative_lot", "status", "expected_effect",
            ]},
            "history": project.get("history") or [],
            "progress": [
                entry for entry in (project.get("progress") or [])
                if isinstance(entry, dict) and str(entry.get("month") or "") in months
            ],
            "news": news,
            "news_summary": next(
                (str(entry.get("text") or "") for entry in (project.get("news_summaries") or [])
                 if isinstance(entry, dict) and str(entry.get("month") or "") == base_month),
                "",
            ),
            "article_count": len(matched),
            "month_counts": {month: sum(1 for item in matched if item["month"] == month) for month in months},
        })
    # 지역 일반 뉴스는 제목에 시군구명이 있고 부동산·개발 관련어가 있는 기사만 후보로 둔다.
    for item in raw[-1]:
        title = norm_text(item["title"])
        if district and district in title and any(word in title for word in NEWS_REAL_ESTATE_WORDS):
            headline_pool.setdefault((item["source"], item["title"]), {**item, "related_project": ""})
    # 다른 시군만 제목에 나온 기사는 뒤로 보내고, 보도 폭이 같으면 해당 지역명이 있는 기사를 먼저 둔다.
    local_tokens = [token for token in {district, district[:-1] if len(district) > 2 else "", "광주" if "광주" in str(region) else ""} if token]
    other_regions = [name for name in JEONNAM_CITY_COUNTY_NAMES if name != district]

    def headline_rank(story: dict[str, Any]) -> tuple[bool, int, bool]:
        local = any(token in story["title"] for token in local_tokens)
        elsewhere = not local and any(name in story["title"] for name in other_regions)
        return (elsewhere, -story["similar_count"], not local)

    ranked = _group_similar_news(list(headline_pool.values()))
    ranked.sort(key=headline_rank)
    monthly_headlines = _pick_monthly_headlines(ranked)
    attach_headline_summaries(monthly_headlines, region, base_month)
    if not built and not monthly_headlines:
        return {"available": False, "reason": "development_projects.json에 해당 시군구 개발사업 목록이 없고 기준월 뉴스도 수집되지 않음", "errors": errors}
    return {
        "available": True,
        "region": region,
        "base_month": base_month,
        "window": {
            "start": f"{start:%Y-%m-%d}",
            "end": f"{end:%Y-%m-%d}",
            "months": months,
            "label": f"{start.year}년 {start.month}월~{end.year}년 {end.month}월" if start.year != end.year else f"{start.year}년 {start.month}~{end.month}월",
        },
        "collected_at": datetime.now().isoformat(timespec="seconds"),
        "news_source": "뉴스 검색(RSS) 기사 제목·매체·보도일 기준, 포털 재전송 제외, 유사 보도는 대표 1건으로 묶음",
        "news_collected": any(project["news"] for project in built) or bool(monthly_headlines),
        "errors": errors,
        "monthly_headlines": monthly_headlines,
        "projects": built,
        "usage_rule": "수집된 것은 기사 제목이므로 제목에 없는 수치·일정을 보태지 않고, 검토·요구·가능성 보도는 확정 사실과 구분하며, 사업 진척을 지가변동률의 직접 원인으로 단정하지 않는다.",
    }


def build_land_price_judgment(
    transactions: dict[str, Any],
    land_price: dict[str, Any],
    region: str = "",
    regional_profile: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """지가 상위 권역과 당월 실거래를 결합해 LLM용 판단 플래그를 Python에서 확정한다."""
    rate_rows = land_price.get("emd_top", []) if land_price.get("available") else []
    tx_rows = transactions.get("emd_top", []) if transactions.get("available") else []
    rows: list[dict[str, Any]] = []
    profile = regional_profile or build_regional_profile(region)
    for rank, rate_row in enumerate(rate_rows[:5], 1):
        area = str(rate_row.get("읍면동") or "").strip()
        tokens = [norm_text(token) for token in re.split(r"[,·/]", area) if norm_text(token)]
        matched = []
        for tx_row in tx_rows:
            tx_name = norm_text(tx_row.get("명칭"))
            if any(tx_name.startswith(token) or token in tx_name for token in tokens):
                matched.append(tx_row)
        count = int(sum(as_number(row.get("거래건수")) or 0 for row in matched))
        amount = round(sum(as_number(row.get("거래금액_억원")) or 0 for row in matched), 1)
        flag = "실거래 연계" if count >= 8 else "혼합" if count >= 4 else "가격조사 중심"
        basis = {
            "실거래 연계": "가격 변화와 실거래 흐름을 함께 확인",
            "혼합": "가격 변화·실거래 병행, 개별 거래 일반화 주의",
            "가격조사 중심": "거래 희소, 조사담당자의 가격 변화·지역 여건 판단 중심",
        }[flag]
        matched_zone = match_regional_zone(area, profile)
        rows.append({
            "rank": rank,
            "area": area,
            "current_rate": rate_row.get("당월변동률"),
            "cumulative_rate": rate_row.get("누계변동률"),
            "transaction_count": count,
            "transaction_amount_eok": amount,
            "flag": flag,
            "basis": basis,
            "regional_zone": matched_zone.get("zone") if matched_zone else "",
            "local_characteristics": matched_zone.get("characteristics") if matched_zone else "",
            "demand_background": matched_zone.get("demand_background") if matched_zone else "",
            "land_use": matched_zone.get("land_use") if matched_zone else "",
            "rear_demand": matched_zone.get("rear_demand") if matched_zone else "",
            "transport": matched_zone.get("transport") if matched_zone else "",
            "commercial_area": matched_zone.get("commercial_area") if matched_zone else "",
            "field_checks": matched_zone.get("field_checks") if matched_zone else "",
            "industrial_links": matched_zone.get("industrial_links", []) if matched_zone else [],
            "analysis_caution": matched_zone.get("analysis_caution") if matched_zone else "",
        })
    return {
        "available": bool(rows),
        "thresholds": {"실거래 연계": "8건 이상", "혼합": "4~7건", "가격조사 중심": "0~3건"},
        "top_areas": rows,
        "calculation_note": "상위 5개 지역의 플래그는 Python이 당월 RTMS_KEY 계약묶음 거래건수로 계산했으며 LLM은 재계산하거나 변경하지 않는다.",
    }


def _josa(word: Any, with_final: str, without_final: str) -> str:
    """받침 유무에 맞는 조사(은/는, 이/가 등)."""
    tail = re.sub(r"[^0-9A-Za-z가-힣]+$", "", str(word or ""))[-1:]
    if tail and "가" <= tail <= "힣":
        return with_final if (ord(tail) - 0xAC00) % 28 else without_final
    return without_final


def attach_area_evidence(result: dict[str, Any]) -> None:
    """지가 상위 지역별로 9장 서술 근거(지역 설명, 당월 실거래, 매물, 영향권, 상승·하락 배경)를 붙인다."""
    region = str((result.get("meta", {}) or {}).get("region") or "")
    profile = result.get("regional_profile", {}) or {}
    dong_profiles = profile.get("dong_profiles") or {}
    zone_briefs = profile.get("zone_briefs") or {}
    listing_rows = (result.get("listings", {}) or {}).get("emd_price_rows") or []
    trade_rows = (result.get("transactions", {}) or {}).get("current_trade_rows") or []
    for row in (result.get("judgment", {}) or {}).get("top_areas", []) or []:
        area = str(row.get("area") or "")
        dongs = [token.strip() for token in re.split(r"[,·/]", area) if token.strip()]
        tokens = [norm_text(dong) for dong in dongs]

        # 지역 한 구절 설명: 조사권역 명칭 → 단일 동 → 권역 요약 순
        zone_brief = zone_briefs.get(str(row.get("regional_zone") or ""), {}) or {}
        row["description"] = (
            dong_profiles.get(area) or dong_profiles.get(norm_text(area))
            or (dong_profiles.get(dongs[0]) if len(dongs) == 1 else "")
            or zone_brief.get("brief") or ""
        )

        # 당월 실거래(금액 큰 순 3건)
        trades = [item for item in trade_rows if norm_text(str(item.get("소재지") or "").split(" ")[0]) in tokens]
        if trades:
            row["trades"] = [
                {key: item.get(key) for key in ["소재지", "물건구분", "용도지역", "지목", "면적_㎡", "거래금액_억원", "거래주체", "매수자지역", "외지인"]}
                for item in trades[:3]
            ]
            row["external_trade_count"] = sum(1 for item in trades if item.get("외지인") == "Y")

        # 매물 수·호가 조정·주요 이용상황
        matched = [item for item in listing_rows if norm_text(item.get("읍면동")) in tokens]
        if matched:
            comparable = sum(int(item.get("비교가능") or 0) for item in matched)
            has_previous = all(item.get("전월_매물수") is not None for item in matched)
            uses: dict[str, int] = {}
            for item in matched:
                for name, count in (item.get("이용상황별") or {}).items():
                    uses[name] = uses.get(name, 0) + int(count)
            row["listing"] = {
                "당월_매물수": sum(int(item.get("당월_매물수") or 0) for item in matched),
                "전월_매물수": sum(int(item.get("전월_매물수") or 0) for item in matched) if has_previous else None,
                "비교가능": comparable,
                "호가인상": sum(int(item.get("호가인상") or 0) for item in matched),
                "호가인하": sum(int(item.get("호가인하") or 0) for item in matched),
                "평균변동률": round(
                    sum(float(item.get("평균변동률") or 0) * int(item.get("비교가능") or 0) for item in matched) / comparable, 1
                ) if comparable else None,
                "주요_이용상황": ", ".join(f"{name} {count}건" for name, count in sorted(uses.items(), key=lambda kv: -kv[1])[:2]),
            }

        # 영향권과 상승·하락 배경
        levels = {dong: (semiconductor_influence(dong, region) or {}).get("level") for dong in dongs}
        direct = [dong for dong, level in levels.items() if level == "direct"]
        linked = [dong for dong, level in levels.items() if level == "linked"]
        if direct and len(direct) == len(dongs):
            row["influence_note"] = "반도체 직접 영향권"
        elif direct:
            row["influence_note"] = f"반도체 직접 영향권 동({'·'.join(direct)}) 포함"
        elif linked:
            row["influence_note"] = "연계 영향권"
        else:
            row["influence_note"] = ""
        rate = as_number(row.get("current_rate")) or 0
        if rate > 0:
            row["reason_hint"] = (
                "반도체 국가산단 예정지와 맞닿은 입지에 대한 개발 기대" if direct
                else "반도체 국가산단 인접 지역으로 번진 개발 기대" if linked
                else zone_brief.get("rise") or "지역 내 토지 수요"
            )
        elif rate < 0:
            row["reason_hint"] = "토지거래허가구역 지정 이후 줄어든 매수세" if direct else "매수 문의와 거래 감소"
        else:
            row["reason_hint"] = ""
        if not int(as_number(row.get("transaction_count")) or 0) and rate > 0:
            row["basis"] = "당월 실거래 없음, 조사담당자가 확인한 매도 호가 상승 중심 판단"


def build_investigator_emd_narratives(result: dict[str, Any]) -> list[dict[str, str]]:
    """상위 5개 지역: 지역 한 구절 → 실거래·매물 → 상승·하락 사유 순의 200자 이내 서술."""
    rows = (result.get("judgment", {}) or {}).get("top_areas", [])
    narratives: list[dict[str, str]] = []
    for row in rows:
        area = str(row.get("area") or "")
        label = "·".join(token.strip() for token in re.split(r"[,·/]", area) if token.strip())
        rate = as_number(row.get("current_rate")) or 0
        count = int(as_number(row.get("transaction_count")) or 0)
        amount = float(as_number(row.get("transaction_amount_eok")) or 0)
        trades = row.get("trades") or []
        listing = row.get("listing") or {}
        hint = str(row.get("reason_hint") or "")

        intro = f"{label}{_josa(label, '은', '는')} {row['description']}임. " if row.get("description") else ""

        trade_text = ""
        if count and trades:
            first = trades[0]
            raw_kind = str(first.get("물건구분") or "")
            jimok = str(first.get("지목") or "")
            if "토지" in raw_kind:
                kind = f"{jimok} 토지".strip()
            elif raw_kind.startswith("기타"):
                # '기타일반'·'기타건물'은 원자료 분류명이므로 건물이 딸린 거래임을 풀어 쓴다.
                kind = f"{'대지' if jimok == '대' else jimok or '토지'}(건물 포함)"
            else:
                kind = raw_kind
            # 여러 동을 묶은 조사권역은 어느 동의 거래인지 밝힌다.
            trade_dong = str(first.get("소재지") or "").split(" ")[0] if "·" in label else ""
            what = " ".join(part for part in [trade_dong, str(first.get("용도지역") or ""), kind] if part)
            base = f"당월 {what}{' 등' if count > 1 else ''} {count}건({amount:,.1f}억 원)이 거래"
            external = int(row.get("external_trade_count") or 0)
            if not external:
                trade_text = base + "됨. "
            elif external == count:
                trade_text = base + ("됐고 매수자는 외지인임. " if count == 1 else "됐고 매수자는 모두 외지인임. ")
            else:
                trade_text = base + f"됐고 이 중 {external}건은 외지인 매수임. "
        elif count:
            trade_text = f"당월 실거래는 {count}건({amount:,.1f}억 원)임. "

        current_count = int(listing.get("당월_매물수") or 0)
        previous_count = listing.get("전월_매물수")
        raised, lowered = int(listing.get("호가인상") or 0), int(listing.get("호가인하") or 0)
        price_bits = "·".join(bit for bit in [f"인상 {raised}건" if raised else "", f"인하 {lowered}건" if lowered else ""] if bit)
        # 매물이 어떤 이용상황 위주인지(예: 단독주택·상업용)를 함께 적어 지역 특성이 드러나게 한다.
        mix = "·".join(part.rsplit(" ", 1)[0] for part in str(listing.get("주요_이용상황") or "").split(", ") if part)
        listing_rich = listing_full = listing_short = ""
        if current_count:
            def listing_text(mix_text: str) -> str:
                head = f"매물은 {mix_text + ' 위주 ' if mix_text else ''}{current_count}건"
                if previous_count is None:
                    stock, joiner, ending = head, "이며", "임. "
                elif current_count == previous_count:
                    stock, joiner, ending = head + "으로 전월과 같", "고", "음. "
                else:
                    diff = current_count - int(previous_count)
                    stock, joiner, ending = head + f"으로 전월보다 {abs(diff)}건 {'늘었' if diff > 0 else '줄었'}", "고", "음. "
                return stock + (f"{joiner} 동일 매물 호가는 {price_bits}임. " if price_bits else ending)

            listing_full = listing_text("")
            listing_rich = listing_text(mix) if mix else listing_full
            listing_short = f"동일 매물 호가는 {price_bits}임. " if price_bits else ""

        price_up = raised > lowered
        if rate > 0:
            if count and price_up:
                reason = f"{hint}{_josa(hint, '이', '가')} 실거래와 호가 인상으로 이어지며 상승함."
            elif count:
                reason = f"{hint} 속에 실거래가 이어지며 상승함."
            elif price_up:
                reason = f"실거래는 없으나 {hint}{_josa_euro(hint)} 매도 호가가 오르며 상승함."
            elif current_count:
                reason = f"실거래는 없으나 {hint}{_josa(hint, '이', '가')} 매물 가격을 받치며 상승함."
            else:
                reason = f"실거래와 매물은 없으나 {hint}{_josa(hint, '이', '가')} 매도 희망가격에 반영되어 상승함."
        elif rate < 0:
            reason = f"{hint}{_josa_euro(hint)} 하락함." if (count or current_count) else f"실거래와 매물이 없는 가운데 {hint}{_josa_euro(hint)} 하락함."
        else:
            reason = "뚜렷한 변동 요인이 없어 보합세를 보임."

        # 200자를 넘으면 매물 서술부터 줄여 사유 문장이 잘리지 않게 한다.
        candidates = [intro + trade_text + listing_rich + reason, intro + trade_text + listing_full + reason,
                      intro + trade_text + listing_short + reason, intro + trade_text + reason, intro + reason]
        text = next((item for item in candidates if len(item) <= 200), candidates[-1])
        narratives.append({"area": area, "flag": str(row.get("flag") or ""), "narrative": text})
    return narratives


def build_investigator_price_reason_narratives(result: dict[str, Any]) -> dict[str, str]:
    """조사보고서형 종결어미와 변동사유 마무리를 현재 계산값으로 구성한다."""
    lp = result.get("land_price", {}) or {}
    current = lp.get("current", {}) or {}
    rows = (result.get("judgment", {}) or {}).get("top_areas", [])

    def rates(names: list[str]) -> str:
        values = [(name, as_number(current.get(name))) for name in names]
        return ", ".join(f"{name} {value:.3f}%" for name, value in values if value is not None)

    tx = result.get("transactions", {}) or {}

    def named(rows: Any) -> list[dict[str, Any]]:
        return [row for row in (rows or []) if str(row.get("명칭") or "").strip() not in {"", "nan", "None", "미상"}]

    def deal(row: dict[str, Any]) -> str:
        return (
            f"{row.get('명칭')}({int(as_number(row.get('거래건수')) or 0)}건·{float(as_number(row.get('거래금액_억원')) or 0):,.1f}억 원, "
            f"금액비중 {float(as_number(row.get('금액비중')) or 0):.1f}%)"
        )

    def top_rates(names: list[str], count: int = 3) -> str:
        values = [(name, as_number(current.get(name))) for name in names]
        values = sorted([item for item in values if item[1] is not None], key=lambda item: -item[1])[:count]
        return ", ".join(f"{name} {value:.3f}%" for name, value in values)

    def trade_reason(rows: list[dict[str, Any]], unit: str, rate_names: list[str], rate_label: str, extra: str = "") -> str:
        rate_part = top_rates(rate_names)
        rate_sentence = f"{rate_label} 지가변동률은 {rate_part} 순으로 확인됨. " if rate_part else ""
        if not rows:
            return f"당월 {unit} 실거래 자료가 확인되지 않음. {rate_sentence}".strip()
        most = max(rows, key=lambda row: int(as_number(row.get("거래건수")) or 0))
        text = f"당월 실거래는 {unit}별로 " + ", ".join(deal(row) for row in rows[:2]) + " 순으로 금액이 컸음. "
        if most.get("명칭") != rows[0].get("명칭"):
            text += (
                f"건수는 {most.get('명칭')} {int(as_number(most.get('거래건수')) or 0)}건"
                f"(건수비중 {float(as_number(most.get('건수비중')) or 0):.1f}%)이 가장 많아 건수와 금액 분포가 다름. "
            )
        return text + extra + rate_sentence + "소수 계약에 금액이 좌우되므로 실거래 분포와 가격 변화를 함께 확인할 필요가 있음."

    zone_text = trade_reason(named(tx.get("zone_top")), "용도지역", ["주거", "상업", "공업", "녹지", "보전관리", "생산관리", "계획관리"], "용도지역별")
    use_rows = named(tx.get("use_top"))
    use_most = max(use_rows, key=lambda row: int(as_number(row.get("거래건수")) or 0)) if use_rows else None
    use_extra = (
        f"물건구분은 {use_most.get('명칭')} {int(as_number(use_most.get('거래건수')) or 0)}건"
        f"(건수비중 {float(as_number(use_most.get('건수비중')) or 0):.1f}%)이 가장 많음. "
        if use_most else ""
    )
    jimok_rows = named(tx.get("jimok_top"))
    use_text = trade_reason(
        jimok_rows or use_rows, "지목" if jimok_rows else "물건구분",
        ["전", "답", "주거용", "상업용", "공업용", "임야"], "이용상황별", use_extra if jimok_rows else "",
    )
    emd_parts: list[str] = []
    for row in rows:
        area = str(row.get("area") or "")
        rate = float(as_number(row.get("current_rate")) or 0)
        count = int(as_number(row.get("transaction_count")) or 0)
        amount = float(as_number(row.get("transaction_amount_eok")) or 0)
        emd_parts.append(f"{area}은 지가변동률 {rate:.3f}%, 실거래 {count}건·{amount:,.1f}억 원으로 확인됨.")
    return {
        "zone_change_reasoning": zone_text,
        "use_change_reasoning": use_text,
        "emd_change_reasoning": " ".join(emd_parts),
    }


def build_fact_bound_transaction_narratives(result: dict[str, Any]) -> dict[str, str]:
    """TOP10 표의 현재 계산값으로 거래주체·단가가 어긋나지 않는 서술을 만든다."""
    tx = result.get("transactions", {}) or {}
    current = tx.get("current", {}) or {}
    top = tx.get("top_trades", []) or []
    external = tx.get("external_top_trades", []) or []

    def type_summary(rows: list[dict[str, Any]]) -> str:
        counts: dict[str, int] = {}
        for row in rows:
            key = str(row.get("거래주체") or "미상")
            counts[key] = counts.get(key, 0) + 1
        return ", ".join(f"{key} {value}건" for key, value in sorted(counts.items(), key=lambda item: (-item[1], item[0])))

    def num(value: Any) -> float:
        return float(as_number(value) or 0)

    def named(rows: Any) -> list[dict[str, Any]]:
        return [row for row in (rows or []) if str(row.get("명칭") or "").strip() not in {"", "nan", "None", "미상"}]

    region = str((result.get("meta", {}) or {}).get("region") or "")
    tokens = region.split()
    district = tokens[-1] if tokens else "해당 지역"

    def short_region(name: Any) -> str:
        parts = str(name or "").split()
        return parts[-1] if len(parts) > 1 and tokens and parts[0] == tokens[0] else str(name or "")

    # 4장 외지인 거래(500자 이내)
    ext_total = num(current.get("외지인거래금액_억원"))
    ext_top_amount = round(sum(num(row.get("거래금액_억원")) for row in external), 2)
    buyers = named(tx.get("external_buyer_region_top"))
    ext_uses = named(tx.get("external_use_top"))
    ext_parts = [
        f"당월 외지인 거래는 {int(num(current.get('외지인거래건수')))}건·{ext_total:,.1f}억 원으로 "
        f"전체 거래금액의 {num(current.get('외지인금액비중')):.1f}%를 차지함."
    ]
    if buyers:
        ext_parts.append(
            "매수자 지역은 "
            + ", ".join(f"{short_region(row.get('명칭'))} {int(num(row.get('거래건수')))}건·{num(row.get('거래금액_억원')):,.1f}억 원" for row in buyers[:2])
            + f" 순이며 상위 {min(2, len(buyers))}개 지역이 외지인 거래금액의 {sum(num(row.get('금액비중')) for row in buyers[:2]):.1f}%를 차지함."
        )
    if external:
        ext_parts.append(
            f"외지인 고가거래 TOP10의 거래주체는 {type_summary(external)}이며 합계 {ext_top_amount:,.2f}억 원으로 "
            f"외지인 거래금액의 {(ext_top_amount / ext_total * 100 if ext_total else 0):.1f}%에 해당함."
        )
        first = external[0]
        ext_parts.append(
            f"최고가 계약은 {first.get('소재지')}({num(first.get('거래금액_억원')):,.2f}억 원, "
            f"{first.get('거래주체') or '주체 미상'}, {first.get('용도지역') or '용도지역 미상'})임."
        )
        if "광산구" in region:
            direct_count = sum(
                1 for row in external
                if (semiconductor_influence(str(row.get("소재지") or "").split(" ")[0], region) or {}).get("level") == "direct"
            )
            ext_parts.append(f"반도체 직접 영향권 소재 거래는 TOP10 중 {direct_count}건임.")
    if ext_uses:
        ext_parts.append(
            f"물건구분은 {ext_uses[0].get('명칭')}({int(num(ext_uses[0].get('거래건수')))}건, "
            f"금액비중 {num(ext_uses[0].get('금액비중')):.1f}%)의 금액이 가장 큼."
        )
    ext_parts.append(f"개별 고액거래와 특정 매수지역 편중을 {district} 전체 외지인 수요 확대로 일반화하기 어려워 계약목적과 후속 이용상황 확인이 필요함.")

    # 5장 ① 거래 규모·자본 흐름(300자 이내)
    change = tx.get("change_vs_previous", {}) or {}
    yoy = tx.get("year_over_year", {}) or {}
    total_amount = num(current.get("거래금액_억원"))

    def signed(value: Any) -> str:
        number = as_number(value)
        return "-" if number is None else f"{number:+.1f}%"

    scale_parts = [
        f"당월 거래는 {int(num(current.get('거래건수')))}건·{total_amount:,.1f}억 원으로 전월 대비 건수 {signed(change.get('거래건수증감률'))}·"
        f"금액 {signed(change.get('거래금액증감률'))}, 전년동월 대비 건수 {signed(yoy.get('거래건수증감률'))}·금액 {signed(yoy.get('거래금액증감률'))}임."
    ]
    groups = tx.get("influence_group_report_rows") or []
    if groups:
        scale_parts.append(
            ", ".join(
                f"{'반도체 ' if str(row.get('구분')) == '직접 영향권' else ''}{row.get('구분')} {int(num(row.get('거래건수')))}건·"
                f"{num(row.get('거래금액_억원')):,.1f}억 원(금액비중 {num(row.get('금액비중')):.1f}%)"
                for row in groups[:2]
            ) + "임."
        )
    if top:
        top_amount = sum(num(row.get("거래금액_억원")) for row in top)
        scale_parts.append(
            f"고가 실거래 TOP10은 {type_summary(top)}이며 합계 {top_amount:,.1f}억 원으로 전체 거래금액의 "
            f"{(top_amount / total_amount * 100 if total_amount else 0):.1f}%를 차지해 소수 고액 계약에 금액이 좌우되는 구조임."
        )

    # 5장 ② 용도지역·물건구분 분포(200자 이내)
    def spread(rows: list[dict[str, Any]], label: str) -> str:
        if not rows:
            return ""

        def cell(row: dict[str, Any]) -> str:
            return f"{row.get('명칭')}(건수비중 {num(row.get('건수비중')):.1f}%·금액비중 {num(row.get('금액비중')):.1f}%)"

        most = max(rows, key=lambda row: num(row.get("거래건수")))
        if most.get("명칭") == rows[0].get("명칭"):
            return f"{label}는 {cell(rows[0])}의 건수와 금액이 모두 가장 큼."
        return f"{label}는 {cell(rows[0])}의 금액이 가장 크고 {cell(most)}의 건수가 가장 많음."

    zone_use_text = " ".join(
        part for part in [spread(named(tx.get("zone_top")), "용도지역별로"), spread(named(tx.get("use_top")), "물건구분별로")] if part
    )
    return {
        "external_buyer_analysis": " ".join(ext_parts),
        "detailed_transaction_analysis": " ".join(scale_parts),
        "zone_use_transaction_interpretation": zone_use_text,
    }


def build_time_series_context(result: dict[str, Any]) -> dict[str, Any]:
    """매월 같은 기준으로 거래와 지가변동률의 시계열상 위치를 계산한다."""
    tx = result.get("transactions", {}) or {}
    lp = result.get("land_price", {}) or {}

    def series_stats(rows: list[dict[str, Any]], key: str, digits: int = 1) -> dict[str, Any]:
        values = [float(value) for row in rows if (value := as_number(row.get(key))) is not None]
        if not values:
            return {}
        current = values[-1]
        previous = values[-2] if len(values) > 1 else None
        descending = sorted(values, reverse=True)
        return {
            "current": round(current, digits),
            "previous": round(previous, digits) if previous is not None else None,
            "average_3m": round(sum(values[-3:]) / len(values[-3:]), digits),
            "average_period": round(sum(values) / len(values), digits),
            "minimum": round(min(values), digits),
            "maximum": round(max(values), digits),
            "rank_desc": descending.index(current) + 1,
            "observations": len(values),
        }

    tx_rows = list(tx.get("monthly_rows", []) or [])[-13:]
    public_rows = [
        row for row in (lp.get("monthly_rows", []) or [])
        if as_number(row.get("평균")) is not None
    ][-13:]
    transaction = {
        "period_months": len(tx_rows),
        "start_month": tx_rows[0].get("월") if tx_rows else None,
        "end_month": tx_rows[-1].get("월") if tx_rows else None,
        "count": series_stats(tx_rows, "거래건수"),
        "amount_eok": series_stats(tx_rows, "거래금액_억원"),
        "external_count": series_stats(tx_rows, "외지인거래건수"),
        "external_amount_eok": series_stats(tx_rows, "외지인거래금액_억원"),
        "mom": dict(tx.get("change_vs_previous", {}) or {}),
        "yoy": dict(tx.get("year_over_year", {}) or {}),
    }
    land_price = {
        "observations": len(public_rows),
        "start_month": public_rows[0].get("월") if public_rows else None,
        "latest_month": public_rows[-1].get("월") if public_rows else lp.get("latest_public_month"),
        "average_rate": series_stats(public_rows, "평균", 3),
        "latest_change_pp": None,
        "investigator_month": None,
        "investigator_current_rate": as_number((lp.get("current", {}) or {}).get("평균")),
        "investigator_previous_rate": as_number((lp.get("previous", {}) or {}).get("평균")),
        "investigator_change_pp": as_number((lp.get("change_vs_previous", {}) or {}).get("평균")),
        "investigator_cumulative_rate": as_number((lp.get("cumulative", {}) or {}).get("평균")),
    }
    rate_stats = land_price["average_rate"]
    if rate_stats and rate_stats.get("previous") is not None:
        land_price["latest_change_pp"] = round(rate_stats["current"] - rate_stats["previous"], 3)
    # 원표에 당월·전월 열이 함께 있으면 파일의 조사월(당월)을 그대로 쓰고,
    # 당월/누계만 있는 구 양식이면 종전대로 기준월 직전월을 조사월로 본다.
    survey_month = lp.get("survey_month")
    if survey_month:
        land_price["investigator_month"] = survey_month
    else:
        try:
            land_price["investigator_month"] = str(pd.Period(str(result.get("meta", {}).get("base_month")), freq="M") - 1)
        except Exception:
            pass
    return {"transactions": transaction, "land_price": land_price}


def build_time_series_narratives(result: dict[str, Any]) -> dict[str, Any]:
    """13개월 범위와 최근 흐름을 먼저 제시하는 월간 조사보고서 문장을 만든다."""
    context = result.get("time_series_context") or build_time_series_context(result)
    tx = context.get("transactions", {}) or {}
    lp = context.get("land_price", {}) or {}
    count = tx.get("count", {}) or {}
    amount = tx.get("amount_eok", {}) or {}
    ext_count = tx.get("external_count", {}) or {}
    ext_amount = tx.get("external_amount_eok", {}) or {}
    mom = tx.get("mom", {}) or {}
    yoy = tx.get("yoy", {}) or {}
    rate = lp.get("average_rate", {}) or {}
    current_tx = (result.get("transactions", {}) or {}).get("current", {}) or {}

    def month_text(value: Any) -> str:
        match = re.fullmatch(r"(\d{4})-(\d{1,2})", str(value or ""))
        return f"{match.group(1)}년 {int(match.group(2))}월" if match else str(value or "-")

    def direction(value: Any) -> str:
        number = float(as_number(value) or 0)
        return "증가" if number > 0 else "감소" if number < 0 else "보합"

    def comparison(current: Any, benchmark: Any) -> str:
        current_number = float(as_number(current) or 0)
        benchmark_number = float(as_number(benchmark) or 0)
        if not benchmark_number:
            return "비교가 제한되고"
        gap = (current_number - benchmark_number) / abs(benchmark_number) * 100
        if abs(gap) <= 5:
            return "유사한 수준이고"
        return "높고" if gap > 0 else "낮고"

    def rank_text(rank: Any) -> str:
        number = int(as_number(rank) or 0)
        if number == 1:
            return "가장 큰"
        return f"{number}번째로 큰"

    count_mom = float(as_number(mom.get("거래건수증감률")) or 0)
    amount_mom = float(as_number(mom.get("거래금액증감률")) or 0)
    count_yoy = as_number(yoy.get("거래건수증감률"))
    amount_yoy = as_number(yoy.get("거래금액증감률"))
    transaction_text = (
        f"최근 {int(tx.get('period_months') or 0)}개월 거래건수는 {count.get('minimum', 0):,.0f}~{count.get('maximum', 0):,.0f}건 범위에서 등락했으며, "
        f"{month_text(tx.get('end_month'))} {count.get('current', 0):,.0f}건은 최근 3개월 평균 {count.get('average_3m', 0):,.1f}건과 비교해 "
        f"{comparison(count.get('current'), count.get('average_3m'))}, 13개월 평균 {count.get('average_period', 0):,.1f}건을 "
        f"{'상회함' if count.get('current', 0) > count.get('average_period', 0) else '하회함'}. 전월보다 {abs(count_mom):,.1f}% {direction(count_mom)}했고"
    )
    if count_yoy is not None:
        transaction_text += f", 전년동월보다 {abs(float(count_yoy)):,.1f}% {direction(count_yoy)}하여 단기 반등과 전년 대비 수준을 구분해 볼 필요가 있는 것으로 판단됨. "
    else:
        transaction_text += " 전년동월 비교자료가 없어 단기 흐름을 중심으로 확인할 필요가 있는 것으로 판단됨. "
    transaction_text += (
        f"거래금액은 {amount.get('current', 0):,.1f}억 원으로 전월보다 {abs(amount_mom):,.1f}% {direction(amount_mom)}했고, "
        f"최근 3개월 평균 {amount.get('average_3m', 0):,.1f}억 원을 {'상회하며' if amount.get('current', 0) > amount.get('average_3m', 0) else '하회하며'} "
        f"최근 {int(amount.get('observations') or 0)}개월 중 {rank_text(amount.get('rank_desc'))} 규모로 확인됨."
    )
    if amount_yoy is not None:
        transaction_text += f" 전년동월 대비로는 {abs(float(amount_yoy)):,.1f}% {direction(amount_yoy)}한 수준임."
    transaction_text += " 거래건수보다 거래금액의 변동 폭이 큰 경우 일부 고액 계약의 영향을 분리해 해석해야 하는 것으로 분석됨."

    land_price_text = "시군구 지가변동률 시계열 자료가 없어 당월 가격조사 결과를 중심으로 판단함."
    if rate:
        latest_change = float(as_number(lp.get("latest_change_pp")) or 0)
        land_price_text = (
            f"시군구 지가변동률 시계열은 {month_text(lp.get('latest_month'))} {rate.get('current', 0):.3f}%로, "
            f"전월 {rate.get('previous', 0):.3f}%보다 {abs(latest_change):.3f}%p {direction(latest_change)}함. "
            f"최근 3개월 평균은 {rate.get('average_3m', 0):.3f}%이고 확인 가능한 {int(rate.get('observations') or 0)}개월 범위는 "
            f"{rate.get('minimum', 0):.3f}~{rate.get('maximum', 0):.3f}%로 완만한 상승 흐름 안에서 월별 등락이 나타난 것으로 확인됨. "
            f"이와 별도로 조사담당자가 {month_text(lp.get('investigator_month'))} 확인한 평균 변동률은 "
            f"{float(as_number(lp.get('investigator_current_rate')) or 0):.3f}%, 누계 {float(as_number(lp.get('investigator_cumulative_rate')) or 0):.3f}%임. "
            f"{month_text(tx.get('end_month'))} 실거래와 가격조사·시계열 확인시점이 다르므로 직접 인과로 연결하지 않고 이후 흐름에서 지속 여부를 확인할 필요가 있는 것으로 판단됨."
        )

    external_text = (
        f"외지인 거래는 {ext_count.get('current', 0):,.0f}건으로 최근 {int(ext_count.get('observations') or 0)}개월 중 "
        f"{rank_text(ext_count.get('rank_desc'))} 수준이고, 거래금액은 {ext_amount.get('current', 0):,.1f}억 원으로 "
        f"{rank_text(ext_amount.get('rank_desc'))} 규모이며 전체 거래금액의 "
        f"{float(as_number(current_tx.get('외지인금액비중')) or 0):.1f}%를 차지함. 외지인 거래 확대가 전체 시장에 고르게 나타난 것인지 특정 고액 거래에 집중된 것인지 소재지와 매수주체별로 구분해 판단할 필요가 있는 것으로 분석됨."
    )
    market_text = (
        f"{month_text(tx.get('end_month'))} 거래건수는 최근 3개월 평균과 유사한 수준이나 13개월 평균을 하회한 반면, "
        f"거래금액은 최근 3개월 평균을 상회하고 외지인 거래금액 비중도 {float(as_number(current_tx.get('외지인금액비중')) or 0):.1f}%로 높게 나타남. "
        "따라서 시장 전반의 고른 회복으로 단정하기보다 고액·외지인 거래가 당월 금액 증가를 견인한 흐름으로 분석됨. "
        f"가격 측면에서는 {month_text(lp.get('latest_month'))} 시군구 지가변동률과 {month_text(lp.get('investigator_month'))} 가격조사 결과의 방향 및 시차를 함께 살피고, "
        "다음 달 거래건수·거래금액·상승지역 확산 여부를 통해 지속성을 확인해야 하는 것으로 판단됨."
    )
    conclusion_text = (
        market_text + " 개별 고액 거래를 권역 전체의 가격 변화로 일반화하지 않고, 용도지역·이용상황·읍면동별 변동률과 후속 거래의 연속성을 함께 점검해야 함. "
        "본 분석은 가격 변화와 지역시장동향 판단을 지원하기 위한 참고자료이며, 최종 변동률 판단과 보고서 확정은 해당 지역 조사담당자의 검토를 거쳐 수행됨."
    )
    return {
        "transaction_interpretation": transaction_text,
        "land_price_interpretation": land_price_text,
        "external_buyer_interpretation": external_text,
        "market_overview": market_text,
        "executive_summary": market_text,
        "conclusion": conclusion_text,
        "key_takeaways": [
            f"거래건수 {count.get('current', 0):,.0f}건: 최근 3개월 평균 {count.get('average_3m', 0):,.1f}건과 비교",
            f"거래금액 {amount.get('current', 0):,.1f}억 원: 최근 {int(amount.get('observations') or 0)}개월 중 {rank_text(amount.get('rank_desc'))} 규모",
            f"외지인 거래금액 비중 {float(as_number(current_tx.get('외지인금액비중')) or 0):.1f}%: 고액 거래 집중 여부 확인 필요",
        ],
    }


def refresh_fact_bound_narratives(result: dict[str, Any], narratives: dict[str, Any]) -> dict[str, Any]:
    """LLM 해설의 객관 수치 문장을 최신 계산값으로 동기화하고 과거 표현을 정리한다."""
    def change_text(value: Any, unit: str) -> str:
        number = as_number(value)
        if number is None:
            return "증감 미확인"
        direction = "증가" if number > 0 else "감소" if number < 0 else "보합"
        return f"{abs(number):,.0f}{unit} {direction}" if number else direction

    def signed_rate_text(value: Any) -> str:
        number = as_number(value)
        return "증감 미확인" if number is None else f"{number:+.2f}%p"

    refreshed = dict(narratives or {})
    # 해석은 LLM 결과를 우선한다. Python 문장은 LLM 값이 비어 있을 때만
    # 1차 HTML이 공란이 되지 않도록 사용하는 제한적 대체문이다.
    for key, value in build_fact_bound_transaction_narratives(result).items():
        if not refreshed.get(key):
            refreshed[key] = value
    for key, value in build_investigator_price_reason_narratives(result).items():
        if not refreshed.get(key):
            refreshed[key] = value
    if not refreshed.get("emd_area_narratives"):
        refreshed["emd_area_narratives"] = build_investigator_emd_narratives(result)

    pop = result.get("population", {}) or {}
    hh = result.get("households", {}) or {}
    mig = result.get("migration", {}) or {}
    if pop.get("available") or hh.get("available") or mig.get("available"):
        pop_month = pop.get("latest_month") or ((pop.get("monthly_rows") or [{}])[-1].get("월")) or "-"
        hh_month = hh.get("latest_month") or ((hh.get("monthly_rows") or [{}])[-1].get("월")) or "-"
        demographic_text = (
            f"주민등록인구는 {pop_month} 기준 {fmt_num(pop.get('current'), 0, '명')}으로 "
            f"전월보다 {change_text(pop.get('delta'), '명')}하고, 세대수는 {hh_month} 기준 "
            f"{fmt_num(hh.get('current'), 0, '세대')}로 전월보다 {change_text(hh.get('delta'), '세대')}함. "
            f"인구이동은 최신 공표월 {mig.get('latest_month') or '-'} 기준 총전입 {fmt_num(mig.get('in_migration'), 0, '명')}, "
            f"총전출 {fmt_num(mig.get('out_migration'), 0, '명')}, 순이동 {fmt_num(mig.get('net_migration'), 0, '명')}으로 확인됨. "
            "인구·세대·이동은 주거·생활형 토지수요의 배경지표로 활용하되, 산업용·사업용 토지수요와 동일한 방향으로 단정하지 않음."
        )
        if not refreshed.get("population_household_migration_interpretation"):
            refreshed["population_household_migration_interpretation"] = demographic_text
        if not refreshed.get("macro_population_interpretation"):
            refreshed["macro_population_interpretation"] = demographic_text

    rates = result.get("rates", {}) or {}
    rate_rows = rates.get("table_rows", []) or []
    latest_rate = rate_rows[-1] if rate_rows else {}
    if rates.get("available"):
        rate_text = (
            f"{rates.get('latest_month') or '-'} 예금은행 신규취급액 기준 금리는 주택담보대출 "
            f"{fmt_num(rates.get('mortgage_rate'), 2, '%')}(전월 대비 {signed_rate_text(latest_rate.get('mortgage_rate_mom'))}), "
            f"전세자금대출 {fmt_num(rates.get('jeonse_rate'), 2, '%')}(전월 대비 {signed_rate_text(latest_rate.get('jeonse_rate_mom'))}), "
            f"중소기업대출 {fmt_num(rates.get('sme_rate'), 2, '%')}(전월 대비 {signed_rate_text(latest_rate.get('sme_rate_mom'))}), "
            f"기업일반자금대출 {fmt_num(rates.get('business_rate'), 2, '%')}(전월 대비 {signed_rate_text(latest_rate.get('business_rate_mom'))})로 확인됨. "
            "금리 변화는 자금조달 부담과 관망심리의 배경지표로만 사용하며 당월 거래량이나 지가변동률의 직접 원인으로 단정하지 않음."
        )
        if not refreshed.get("rate_interpretation"):
            refreshed["rate_interpretation"] = rate_text

    def clean(value: Any) -> Any:
        if isinstance(value, dict):
            return {key: clean(item) for key, item in value.items()}
        if isinstance(value, list):
            return [clean(item) for item in value]
        if not isinstance(value, str):
            return value
        value = value.replace("공표 가능한 ", "업로드한 당월 지가동향 원표 기준 ")
        value = re.sub(
            r"공표 가능한\s+(\d{4}년\s*\d{1,2}월)\s+지가변동률",
            r"업로드한 당월 지가동향 원표 기준 \1 지가변동률",
            value,
        )
        value = value.replace(
            "용도지역별 2026년 6월 공표변동률",
            "업로드한 당월 지가동향 원표 기준 2026년 6월 용도지역별 변동률",
        )
        value = value.replace(
            "이용상황별 2026년 6월 공표변동률",
            "업로드한 당월 지가동향 원표 기준 2026년 6월 이용상황별 변동률",
        )
        value = value.replace("반도체 산업단지 등 사업용 토지수요", "산업단지 관련 사업용 토지수요")
        value = value.replace("조사 표본지 가격조사 결과", "조사담당자가 가격 변화·실거래·지역시장동향을 종합해 판단한 결과")
        # 반도체 영향권 라벨 통일(2026-08 조사): 핵심·인접 → 반도체 직접 영향권
        value = re.sub(r"(핵심|인접)\s*영향권", "반도체 직접 영향권", value)
        value = value.replace("반도체 반도체 직접 영향권", "반도체 직접 영향권")
        value = value.replace("신고일자 기준", "계약일 기준")
        value = value.replace("조사 표본지", "조사대상 토지")
        value = value.replace("조사 표본가격", "조사담당자가 확인한 가격 변화")
        value = value.replace("조사표본 주도", "가격조사 중심")
        value = value.replace("조사표본", "가격조사")
        value = value.replace("실거래 주도", "실거래 연계")
        value = value.replace("가격조사 주도", "가격조사 중심")
        value = value.replace("공표통계", "지가변동률 통계")
        value = value.replace("공표변동률", "지가변동률")
        value = value.replace("공표 시계열", "시군구 지가변동률 시계열")
        value = value.replace("공표자료", "시군구 시계열 자료")
        value = value.replace("최근 공표 지가변동률", "최근 시군구 지가변동률")
        value = value.replace("공표 지가변동률", "시군구 지가변동률")
        value = value.replace("최신 공표월", "최신 확인월")
        value = value.replace("공표월", "확인월")
        value = value.replace("당월 조사담당자가 확인한 가격 변화은", "당월 지가변동률은")
        value = re.sub(r"(\d{1,2}월) 조사담당자가 확인한 가격 변화은", r"\1 지가변동률은", value)
        value = value.replace("조사담당자가 확인한 가격 변화이", "지가변동률이")
        value = value.replace("조사담당자가 확인한 가격 변화을", "조사담당자가 확인한 가격 변화를")
        value = value.replace("조사담당자의 가격조사 결과", "조사담당자가 가격 변화·실거래·지역시장동향을 종합한 판단 결과")
        value = value.replace("6월 당월 조사값", "2026년 6월 지가변동률")
        value = value.replace("조사 변동률", "지가변동률")
        value = value.replace("Python 판단 결과상", "본 보고서 판단 기준상")
        value = value.replace("판단 플래그", "판단구분")
        value = value.replace(
            "최종 변동률 선정·공표 책임은 담당자에게 있으며",
            "본 분석은 가격 변화와 지역시장동향 판단을 지원하기 위한 참고자료이며, 최종 변동률 판단과 보고서 확정은 해당 지역 조사담당자의 검토를 거쳐 수행되며",
        )
        value = value.replace("표본 특성", "조사대상 토지 특성")
        value = value.replace("광산구의 광산구 군공항", "광산구 군공항")
        value = value.replace("기업대출금리", "기업일반자금대출금리")
        value = value.replace("실수요 배후수요", "배후수요")
        value = value.replace("거래규제", "거래 규제")
        value = value.replace("일반화하지 않을 필요가 있음.", "일반화하기 어려운 것으로 판단됨.")
        value = value.replace("확인할 필요가 있음.", "확인이 필요한 것으로 판단됨.")
        value = value.replace("관찰할 필요가 있음.", "지속적인 관찰이 필요한 것으로 판단됨.")
        value = value.replace("단정하기 어려움.", "단정하기 어려운 것으로 판단됨.")
        value = value.replace("동일한 수요로 묶기 어려움.", "동일한 수요로 보기 어려운 것으로 분석됨.")
        value = value.replace("차이가 큼.", "차이가 큰 것으로 확인됨.")
        value = value.replace("상대적으로 크게 나타남.", "상대적으로 크게 나타난 것으로 분석됨.")
        value = value.replace("순으로 나타남.", "순으로 나타난 것으로 확인됨.")
        value = value.replace("로 하락함.", "로 하락한 것으로 확인됨.")
        value = value.replace("에 달함.", "에 달하는 것으로 확인됨.")
        value = value.replace("해석하기 어려움.", "해석하기 어려운 것으로 판단됨.")
        value = value.replace("표현하기 어려움.", "표현하기 어려운 것으로 판단됨.")
        # "직접 원인으로 단정하지 않음."처럼 앞에 이미 원인 어구가 있으면 흡수하여
        # "직접 원인으로 직접적인 원인으로"가 이중으로 겹치지 않도록 한다.
        value = re.sub(
            r"(?:직접(?:적인)?\s*원인으로\s*)?단정하지 않음\.",
            "직접적인 원인으로 단정하기 어려운 것으로 판단됨.",
            value,
        )
        value = value.replace("작용할 수 있음.", "배경요인으로 작용하는 것으로 판단됨.")
        value = value.replace("해석해야 함.", "구분하여 해석할 필요가 있는 것으로 판단됨.")
        value = value.replace("확인해야 함.", "확인이 필요한 것으로 판단됨.")
        value = value.replace("타당함.", "타당한 것으로 판단됨.")
        value = value.replace("등락함.", "등락한 것으로 확인됨.")
        # 치환 상호작용으로 생길 수 있는 원인 어구 중복을 접는 안전망.
        value = re.sub(r"직접(?:적인)?\s*원인으로\s*(?=직접(?:적인)?\s*원인으로)", "", value)
        value = re.sub(r"(?<=\d)억원", "억 원", value)
        return re.sub(r"\s+$", "", value)

    return clean(refreshed)


def build_result(
    input_dir: Path,
    region: str,
    base_month: str,
    transaction_files: list[str | Path] | None = None,
    listing_files: list[str | Path] | None = None,
    land_price_files: list[str | Path] | None = None,
    as_of_date: Any = None,
    district_rate_override: dict[str, Any] | None = None,
) -> dict[str, Any]:
    buckets = find_files(input_dir) if input_dir and input_dir.exists() else {
        "transactions": [],
        "land_price": [],
        "population": [],
        "household": [],
        "migration": [],
        "rates": [],
        "listings": [],
    }
    explicit_transactions = existing_paths(transaction_files)
    explicit_listings = existing_paths(listing_files)
    explicit_land_price = existing_paths(land_price_files)
    # 직접 파일 선택 화면에서 빈 목록을 전달한 경우에도 폴더 자동 탐색값으로
    # 되돌리지 않는다. 특히 매물 미지정은 의도적인 '양식만 표시' 선택이다.
    if transaction_files is not None:
        buckets["transactions"] = explicit_transactions
    if listing_files is not None:
        buckets["listings"] = explicit_listings
    if land_price_files is not None:
        buckets["land_price"] = explicit_land_price
    resolved_as_of = resolve_as_of_date(base_month, as_of_date)
    transactions = summarize_transactions(buckets["transactions"], base_month, region, as_of_date=resolved_as_of)
    land_price = apply_district_rate_override(summarize_land_price(buckets["land_price"], region), district_rate_override)
    regional_profile = build_regional_profile(region)
    if regional_profile.get("available"):
        zone_totals: dict[str, dict[str, Any]] = {
            str(zone.get("zone")): {"zone": zone.get("zone"), "거래건수": 0, "거래금액_억원": 0.0, "읍면동": []}
            for zone in regional_profile.get("zones", [])
        }
        for row in transactions.get("emd_top", []):
            matched = match_regional_zone(row.get("명칭"), regional_profile)
            if not matched:
                continue
            target = zone_totals[str(matched.get("zone"))]
            target["거래건수"] += int(as_number(row.get("거래건수")) or 0)
            target["거래금액_억원"] += float(as_number(row.get("거래금액_억원")) or 0)
            target["읍면동"].append(str(row.get("명칭")))
        regional_profile["transaction_summaries"] = [
            {**value, "거래금액_억원": round(float(value["거래금액_억원"]), 1)}
            for value in zone_totals.values()
        ]
    # 고가 실거래 서술에서 LLM이 소재지-권역 관계를 추측하지 않도록 Python에서 연결 정보를 확정한다.
    for trade_key in ["top_trades", "external_top_trades"]:
        for trade in transactions.get(trade_key, []):
            matched_zone = match_regional_zone(trade.get("소재지"), regional_profile)
            if matched_zone:
                trade["분석권역"] = matched_zone.get("zone")
                trade["지역특성"] = matched_zone.get("characteristics")
                trade["연계산업거점"] = matched_zone.get("industrial_links", [])
                trade["해석유의"] = matched_zone.get("analysis_caution")
    result = {
        "meta": {
            "report_type": "지가변동률 월간시장동향",
            "region": region,
            "base_month": base_month,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "as_of_date": f"{resolved_as_of:%Y-%m-%d}",
            "report_cadence": "월간 본문 + 주간(매주 금요일) 모니터링 부록",
            "transaction_date_basis": transactions.get("date_basis", "계약일"),
            "input_dir": str(input_dir),
            "input_mode": "direct_files" if any(x is not None for x in [transaction_files, listing_files, land_price_files]) else "folder_scan",
        },
        "processing_roles": {
            "python": "원자료 정제, 13개월 시계열, 전월·전년동월 증감률, 평균·범위·순위·집중도 계산",
            "llm": "Python 계산값과 지역 특성을 근거로 조사보고서 해석 및 자연어 서술",
            "rule": "LLM은 수치를 재계산하지 않고 time_series_context와 result.json의 계산값만 인용한다.",
        },
        "input_files": {k: [str(p) for p in v] for k, v in buckets.items()},
        "transactions": transactions,
        "land_price": land_price,
        "population": summarize_simple_timeseries(buckets["population"], "population", base_month),
        "households": summarize_simple_timeseries(buckets["household"], "household", base_month),
        "age_population": summarize_age_population(buckets["population"], base_month),
        "migration": summarize_migration(buckets["migration"], region, base_month),
        "rates": summarize_rates(buckets["rates"], base_month),
        "listings": summarize_listings(buckets["listings"], base_month, region),
        "field_monitoring": load_field_monitoring(region, base_month, input_dir),
        "regional_issues": build_regional_issues(region, base_month),
        "development_projects": build_development_projects(region, base_month, input_dir),
        "regional_profile": regional_profile,
        "judgment": build_land_price_judgment(transactions, land_price, region, regional_profile),
        "narrative_policy": {
            "sentence_endings": ["~됨", "~확인됨", "~판단됨", "~분석됨", "~조사됨", "~나타났음", "~증가했음", "~감소했음", "~형성됐음", "~상향됐음", "~전환됐음", "~견인함", "~작용함", "~형성됨", "~유지됨"],
            "forbidden_replacements": {
                "압도적": "가장 높은", "미래 가치 선점 수요": "향후 개발 기대를 반영한 수요",
                "광역 자본의 전략적 포트폴리오 편입": "광주권 자본의 자산 배분 및 매수세 유입",
                "강한 하방 경직성": "가격 하락이 제한되는 흐름", "강력한 하방 경직성": "가격 하락이 제한되는 흐름",
                "거래절벽": "거래 소강", "눈치보기": "관망 기조", "폭발적으로": "크게",
                "표준지 평가": "조사담당자의 가격조사", "조사 표본지 가격조사": "조사담당자의 가격조사",
                "조사 표본가격": "조사담당자가 확인한 가격 변화", "조사표본": "가격조사", "공표통계": "지가변동률 통계",
            },
            "regional_four_stage_structure": ["지역 특성", "수요 배경", "당월 거래·개발 기대", "지가 변동 사유와 판단 플래그별 마무리"],
            "gwangsan_regional_rules": ["5개 분석권역별 2~3문장", "기존 산업단지와 연계 개발지구의 추진단계 구분", "산단 인접성과 거래·지가 변동의 인과관계 단정 금지", "상위 5개 지역은 권역·산단 연계와 실제 거래지표를 함께 제시"],
            "news_rules": ["나주시청·전라남도·국토교통부·LH 등 1차 출처 우선", "출처와 발표일 명시", "구상·발표·예타·고시·보상·착공·준공 중 추진단계 명시", "확정성이 낮거나 혼동 가능한 철도·역 신설 사안 제외", "실수요 유입 여부와 제약요인을 함께 서술"],
            "publication_notice": "본 분석은 가격 변화와 지역시장동향 판단을 지원하기 위한 참고자료이며, 최종 변동률 판단과 보고서 확정은 해당 지역 조사담당자의 검토를 거쳐 수행된다.",
        },
        "analysis_rules": [
            "모든 수치·순위·증감률은 result.json 값을 사용한다.",
            "실거래는 신고일이 아니라 계약일(계약일자) 기준으로 집계한다.",
            "2026년 7월 이전은 월단위로 비교하고, 2026년 7월부터는 토지거래허가구역 효력 발생일(2026-07-14)을 기준으로 지정 전(7월 1~13일)·지정 후(7월 14~31일) 구간을 나눠 읍면동 거래건수를 비교한다.",
            "지정 전 13일·지정 후 18일로 구간 일수가 다르므로 건수 비교는 transactions.permit_split의 일평균건수를 함께 인용한다.",
            "계약일 기준 집계는 신고기한(계약일부터 30일) 때문에 최근 구간이 과소집계될 수 있으므로 지정 후 구간·최근 주차의 감소를 규제 효과로 단정하지 않는다.",
            "주간 모니터링은 transactions.weekly의 토요일~금요일 7일 창(매주 금요일 작성) 값을 사용하며, 계약일 기준이므로 최근 주차는 신고 지연으로 변동될 수 있음을 명시한다.",
            "반도체 영향권 표기는 송정동·도산동·신촌동·소촌동·우산동·황룡동을 반도체 직접 영향권, 선암동·운수동을 연계 영향권으로 통일한다.",
            "실거래 거래건수는 RTMS_KEY 마지막 필지순번을 제외한 계약묶음 단위이며, 계약총금액은 묶음당 1회 반영하고 토지면적은 필지별 합산값을 사용한다.",
            "외지인 거래는 선택 시군구와 매수자 시군구가 다른 계약묶음으로 판정한다.",
            "지가변동률은 해당 지역 조사담당자가 가격 변화, 실거래 및 지역시장동향을 종합하여 판단한 결과이며 실거래 가격 변동률과 동일시하지 않는다.",
            "대출금리는 매수여력·관망심리 배경지표로만 사용한다.",
            "매물 전월·당월 파일이 모두 없으면 호가 증감 판단을 제한한다.",
            "거래건수가 적은 읍면동은 방향성 단정 대신 관찰 필요로 표현한다.",
            "지역 이슈는 개발 기대·거래 규제 배경으로만 해석하고 지가변동률의 직접 원인으로 단정하지 않는다.",
            "광주권역 개발·규제 이슈 표기는 regional_issues.display_area 값을 그대로 사용한다.",
            "광산구 보고서는 regional_profile의 5개 권역과 산업단지 목록을 사용하되 산단 인접성만으로 가격 변동의 원인이라고 단정하지 않는다.",
            "월간 서술은 13개월 범위, 최근 3개월 평균, 전월·전년동월 비교, 당월 순위와 특이거래 영향, 지가변동률 확인시점 순으로 작성한다.",
            "월별 수치를 나열하지 않고 고점·저점·방향 전환과 당월 위치를 중심으로 서술한다.",
        ],
    }
    public_rate_by_month = {
        str(row.get("월")): row.get("평균")
        for row in land_price.get("monthly_rows", [])
        if row.get("월")
    }
    for row in transactions.get("monthly_rows", []):
        row["공표변동률"] = public_rate_by_month.get(str(row.get("월")))
    for row in transactions.get("monthly_rows_with_prev", []):
        row["공표변동률"] = public_rate_by_month.get(str(row.get("월")))
    attach_area_evidence(result)
    result["time_series_context"] = build_time_series_context(result)
    return result


# LLM 프롬프트에는 HTML 표·차트 전용 대량 행을 제외한 경량 뷰를 넣는다(정보 손실 없이 토큰 절감).
# 저장되는 result.json 원본은 그대로 유지되어 HTML 렌더링에 사용된다.
# 아래 키들은 프롬프트 규칙이 참조하지 않는 HTML 전용 대량 데이터다.
_LLM_HTML_ONLY_KEYS = {
    "transactions": ("emd_monthly_rows", "monthly_rows_with_prev", "current_trade_rows"),
    # 주간 읍면동 표는 HTML 전용 대량 행이며, 프롬프트에는 weekly.weekly_rows 요약만 넣는다.
    "rates": ("table_rows", "table_rows_with_prev"),
    "age_population": ("monthly_rows_with_prev",),
    # 읍면동별 매물 전체 행은 HTML·집계용이며, 프롬프트에는 judgment.top_areas[].listing 요약만 넣는다.
    "listings": ("emd_price_rows",),
}


def build_llm_data_view(result: dict[str, Any]) -> dict[str, Any]:
    """LLM 프롬프트용 경량 사본. HTML 전용 대량 행(읍면동 월별 표 등)과 경로 메타를 제외한다.

    LLM이 해석에 실제로 쓰는 top·집계·시계열 요약 키는 모두 보존하고,
    HTML 렌더 전용 대량 리스트만 제거한다. 원본 result는 변경하지 않는다.
    """
    view = dict(result)
    for parent, drop_keys in _LLM_HTML_ONLY_KEYS.items():
        node = view.get(parent)
        if isinstance(node, dict) and drop_keys:
            view[parent] = {k: v for k, v in node.items() if k not in drop_keys}
    # 읍면동 단위 대량 행(주간·월/주 혼합 표)은 직접 영향권만 남기고 제외한다.
    weekly = (view.get("transactions") or {}).get("weekly")
    if isinstance(weekly, dict) and weekly.get("weekly_emd_rows"):
        trimmed = {k: v for k, v in weekly.items() if k != "weekly_emd_rows"}
        trimmed["weekly_emd_rows_direct"] = [
            row for row in weekly.get("weekly_emd_rows", [])
            if str(row.get("영향권등급")) == "direct"
        ]
        view["transactions"] = {**view["transactions"], "weekly": trimmed}
    month_week = (view.get("transactions") or {}).get("emd_month_week")
    if isinstance(month_week, dict) and month_week.get("rows"):
        trimmed_mw = {k: v for k, v in month_week.items() if k != "rows"}
        trimmed_mw["rows_direct"] = [
            row for row in month_week.get("rows", [])
            if str(row.get("영향권등급")) == "direct"
        ]
        view["transactions"] = {**view["transactions"], "emd_month_week": trimmed_mw}
    # 기사 링크는 HTML 전용이므로 프롬프트에서는 제목·매체·보도일만 남긴다.
    dev = view.get("development_projects")
    if isinstance(dev, dict) and dev.get("available"):
        def without_url(items: Any) -> list[dict[str, Any]]:
            return [{k: v for k, v in item.items() if k != "url"} for item in (items or [])]

        view["development_projects"] = {
            **dev,
            "monthly_headlines": without_url(dev.get("monthly_headlines")),
            "projects": [{**project, "news": without_url(project.get("news"))} for project in dev.get("projects", [])],
        }
    view.pop("input_files", None)
    return view


def build_llm_prompt(result: dict[str, Any]) -> str:
    schema = {
        "key_takeaways": ["", "", ""],
        "executive_summary": "",
        "transaction_interpretation": "",
        "zone_use_transaction_interpretation": "",
        "land_price_interpretation": "",
        "external_buyer_interpretation": "",
        "listing_interpretation": "",
        "population_household_migration_interpretation": "",
        "macro_population_interpretation": "",
        "rate_interpretation": "",
        "regional_issue_interpretation": "",
        "external_buyer_analysis": "",
        "detailed_transaction_analysis": "",
        "market_overview": "",
        "regional_analysis": "",
        "influence_zone_narratives": [
            {"zone": "반도체 직접 영향권", "narrative": ""},
            {"zone": "간접 영향권", "narrative": ""},
        ],
        "regional_zone_narratives": [{"zone": "", "narrative": ""}],
        "industrial_complex_analysis": "",
        "price_change_reasoning": "",
        "zone_change_reasoning": "",
        "use_change_reasoning": "",
        "emd_change_reasoning": "",
        "emd_area_narratives": [{"area": "", "flag": "", "narrative": ""}],
        "news_items": [{"title": "", "source": "", "source_date": "", "project_stage": "", "summary": "", "impact": ""}],
        "monthly_top_news": [{"title": "", "source": "", "source_date": "", "summary": "", "impact": ""}],
        "development_project_updates": [{"project": "", "stage": "", "monthly_progress": "", "recent_progress": "", "impact": ""}],
        "monitoring_points": ["", "", ""],
        "conclusion": "",
        "weekly_field_survey": {
            "overall": "",
            "group_narratives": [
                {"group": "직접 영향권", "scope": "", "narrative": ""},
                {"group": "간접 영향권", "scope": "", "narrative": ""},
            ],
            "price_trend": "",
            "market_outlook": "",
        },
    }
    return f"""당신은 한국부동산원 지역 조사 담당자의 지가변동률 월간시장동향 보고서 작성을 보조합니다.

아래 result.json은 Python으로 계산한 분석자료입니다. LLM은 숫자를 새로 계산하지 말고, JSON에 있는 값만 사용하여 해석 문장을 작성합니다.

[작성 원칙]
1. 지가변동률은 해당 지역 조사담당자가 가격 변화, 실거래 및 지역시장동향을 종합해 판단한 결과입니다. 실거래 증가만으로 지가 상승을 단정하지 마십시오.
2. 대출금리는 매수여력, 자금조달 부담, 관망심리의 배경지표로만 서술하십시오.
3. 인구·세대수·이동자수는 토지시장 수요 기반의 보조자료로만 연결하십시오.
4. listings.template_only가 true이면 listing_interpretation은 빈 문자열로 두고 매물 증감·호가 판단 문장을 작성하지 마십시오. 매물 파일이 있으나 전월·당월 2개 미만인 경우에만 제한적으로 표현하십시오.
5. 지역 이슈는 개발 기대·거래 규제 배경으로만 서술하고, 지가변동률의 직접 원인으로 단정하지 마십시오.
6. 광주권역 개발·규제 이슈 표기는 반드시 result.json의 regional_issues.display_area 값을 그대로 쓰십시오.
7. 출력은 JSON 객체 하나만 작성하십시오. Markdown 코드블록은 쓰지 마십시오.
8. 업로드 기준 양식의 1~10장과 부록 2를 채울 수 있도록 external_buyer_analysis, detailed_transaction_analysis, market_overview, regional_analysis, regional_zone_narratives, industrial_complex_analysis, zone_change_reasoning, use_change_reasoning, emd_change_reasoning, monitoring_points, conclusion, news_items를 구체적으로 작성하십시오.
9. news_items는 확인된 자료가 있을 때만 작성하고, 확인되지 않은 사업명·일정·금액을 만들지 마십시오.
10. 부동산원 조사보고서 문체를 사용하고 문장은 '~됨', '~확인됨', '~판단됨', '~분석됨', '~견인함', '~작용함', '~형성됨', '~유지됨'으로 끝내십시오. 설명형 존댓말과 구어체는 사용하지 마십시오.
11. narrative_policy.forbidden_replacements의 금지 표현은 사용하지 말고 지정된 대체 표현을 사용하십시오. '조사표본', '조사 표본가격', '조사 표본지', '공표통계', '표준지 평가'라는 표현은 사용하지 마십시오.
12. regional_analysis는 '토지이용·입지 특성 → 실거래 배후수요 → 산업단지·교통·상권 중 해당 요인 → 조사담당자가 확인한 당월 가격 변화와 실거래 → 현장 확인사항과 판단구분별 마무리' 순서로 작성하십시오. 9장 emd_area_narratives는 49번 규칙을 따르십시오.
13. judgment.top_areas의 flag와 basis는 8장 판단 근거 표에만 쓰는 계산 결과이므로 재계산하거나 변경하지 말고, 9장 서술에는 flag 명칭이나 판단 기준 문장을 쓰지 마십시오.
14. emd_area_narratives는 judgment.top_areas와 같은 순서·area·flag로 작성하십시오. 판단 자료가 없으면 빈 배열로 두십시오.
15. news_items는 기준월(meta.base_month) 관점으로 정리하십시오. 기준월 중 발생한 발표·고시·효력 이슈를 먼저 쓰고, 기준월 이전에 지정·착수된 사업은 기준월 시장에 미친 영향과 남은 절차 중심으로 정리하며 source_date는 실제 발표·고시일을 기재하십시오. impact에는 해당 이슈가 기준월 지가변동률·실거래·매물 지표와 어떻게 연결되는지(또는 아직 연결되지 않았는지)를 한 문장 이상 담으십시오. 시청·도청·국토교통부·LH 등 1차 출처를 우선 확인하고 source, source_date, project_stage를 반드시 기재하십시오. 추진단계와 실수요 유입 여부·제약요인을 함께 쓰고, 확정성이 낮거나 혼동 가능한 철도·역 신설 사안은 제외하십시오.
16. 결론에는 '본 분석은 가격 변화와 지역시장동향 판단을 지원하기 위한 참고자료이며, 최종 변동률 판단과 보고서 확정은 해당 지역 조사담당자의 검토를 거쳐 수행됨'이라는 취지를 포함하십시오.
17. external_buyer_analysis에는 transactions.external_top_trades의 외지인 고가 실거래 TOP10을 근거로 매수자 지역·거래주체(실거래 엑셀 '매수자' 컬럼 원문: 개인·법인·국가등·기타·외국인 등)·소재지·용도지역·면적·단일 필지 토지 단가·거래금액 집중도를 서술하되, 단가는 Python에서 천원 미만 절사한 값을 그대로 사용하고 개별 고가거래를 지역 전체 흐름으로 일반화하지 마십시오.
18. detailed_transaction_analysis에는 transactions.top_trades의 전체 고가 실거래 TOP10을 근거로 실거래 엑셀 '매수자' 컬럼에서 추출한 거래주체·면적(㎡/평)·단일 필지이면서 토지만 거래한 경우의 단가를 함께 검토하십시오. 단가는 Python에서 천원 미만 절사한 값을 재계산하지 말고 그대로 사용하십시오. 소재지의 '외(00필지)' 표기는 동일 RTMS_KEY로 묶인 대표 주소 외 추가 필지 수입니다.
19. zone_change_reasoning(용도지역별 변동사유)과 use_change_reasoning(이용상황별 변동사유)은 각각 공백 포함 300자 이내의 완결된 문장으로 작성하십시오. 두 문단을 합치거나 같은 문장을 반복하지 말고, 문장은 '~함.'·'~됨.'·'~판단됨.' 등으로 끝맺으며 분량에 맞추려고 숫자나 어절 중간에서 끊지 마십시오. 서술 순서와 근거는 43번 규칙(실거래 중심)을 따르되, 특정 거래·입지가 지가상승을 '견인함'이라고 단정하지 말고 확인 수준에 따라 '함께 나타남'·'상승 배경으로 일부 작용함'·'현장 시장동향 중심으로 판단됨'으로 마무리하십시오. 자료가 없는 구분은 수치를 만들지 말고 '자료 미확인'으로 두십시오.
20. regional_profile.available이 true이면 regional_zone_narratives를 regional_profile.zones와 같은 순서·zone으로 작성하고 권역당 2~3문장으로 작성하십시오. 각 문단은 지역 특성, 당월 실거래 또는 지가 자료, 수요 배경, 제약요인을 연결하되 자료가 없는 수치는 만들지 마십시오.
21. industrial_complex_analysis는 regional_profile.industrial_complexes의 분류와 특성을 근거로 작성하십시오. 기존 산업단지, 광산구·북구 연계 산업단지, 연계 개발지구를 구분하고 기업 입주·가동·고용 또는 사업 추진단계가 확인되지 않은 기대를 실수요로 표현하지 마십시오. 항목에 verified_facts가 있으면 그 기업·시설·일정은 확인된 사실로 서술하되, verified_as_of 시점 기준임을 감안하고 다른 행정구역 소재(예: 첨단3지구는 광주 북구) 표기를 지키며, 확인된 시설 존재를 곧바로 당월 지가 상승의 원인으로 연결하지 마십시오.
22. detailed_transaction_analysis는 고가 실거래 TOP10의 소재지를 regional_profile의 권역·산업단지와 연결해 구체적으로 설명하되, 단순 인접성만으로 산업단지 수요 또는 가격 상승의 인과관계를 단정하지 마십시오.
23. emd_area_narratives 상위 5개는 49번 규칙에 따라 judgment.top_areas의 description·trades·listing·influence_note·reason_hint를 근거로 작성하십시오.
24. 모든 지역에 산업단지·교통·상권을 기계적으로 나열하지 마십시오. 해당 지역의 토지이용과 실제 거래를 설명하는 요인만 선택하고, '배후수요가 있음'으로 단정하지 말고 공동주택 거주인구·산단 종사자·역 이용객·기업 입주 등 수요 주체를 구체적으로 적으십시오. field_checks는 조사자가 현장에서 확인할 사항으로 문장에 포함하고, 사업 기대와 지가변동의 직접 인과관계는 단정하지 마십시오.
25. transaction_interpretation, land_price_interpretation, market_overview, executive_summary, conclusion은 time_series_context를 우선 사용하십시오. 서술 순서는 '13개월 범위와 당월 위치 → 최근 3개월 평균 → 전월 및 전년동월 비교 → 당월 순위·고액 또는 외지인 거래 집중 영향 → 가격조사와 실거래의 확인시점 차이'로 하되, 13개월의 월별 수치를 차례로 나열하지 마십시오.
26. 장별 역할을 분리하십시오. key_takeaways는 거래량 수준·고액거래에 따른 금액 확대·가격조사와 실거래 시점 차이만 간결하게 제시하고, executive_summary는 시장 국면과 핵심 위험요인, detailed_transaction_analysis는 고가거래 계약구조·매수주체·용도지역·이상치, market_overview와 regional_analysis는 권역별 차별화와 수요구조, conclusion은 최종 판단과 향후 확인사항만 작성하십시오. 같은 수치와 문장을 여러 장에서 반복하지 마십시오.
27. zone_use_transaction_interpretation은 transactions.zone_top과 transactions.use_top만 사용하여 당월 용도지역·이용상황별 건수비중과 금액비중의 차이를 설명하십시오. 13개월 거래범위나 전월·전년동월 증감은 반복하지 마십시오.
28. external_buyer_analysis은 외지인 매수지역·개인/법인 구분·자금 집중도·지역 편중을 중심으로 작성하고, detailed_transaction_analysis은 계약묶음 구조·용도지역·토지만 거래 비중·산업권 연계 가능성·이상치 검토를 중심으로 작성하여 삼거동·우산동 개별 수치를 중복 나열하지 마십시오.
29. regional_profile.analysis_units와 overlap_rule을 지키십시오. 도호동·신촌동의 송정권 특성과 우산동의 수완·신가 생활권 및 생산녹지 거래를 분리하고, 가격조사 권역·생활권·반도체 영향권을 같은 개념으로 혼용하지 마십시오.
30. regional_issues.regulation_timing_note를 conclusion 또는 monitoring_points에 반영하십시오. 실거래는 계약일(계약일자) 기준이므로 2026년 7월 14일 이전 계약분을 규제 시행 이후 시장반응으로 해석하지 말고, 규제 전후 비교는 transactions.permit_split의 지정 전(7월 1~13일)·지정 후(7월 14~31일) 구간값으로 서술하십시오.
31. news_items는 정부·지자체·공공기관 원문을 우선하고 포털 재전송 페이지를 출처로 쓰지 마십시오. regional_issues.events의 project_stage를 그대로 반영하되, 부지 낙점·투자 발표 등 확인된 사실과 국가산단 지정·착공·군공항 이전 등 '진행 예정' 후속 절차를 명확히 구분해 서술하고, 확인되지 않은 단계를 확정된 것으로 표현하지 마십시오.
32. Python 계산값은 표시 자릿수를 그대로 사용하십시오. 거래건수·금액·증감률은 소수점 첫째 자리, 지가변동률은 소수점 셋째 자리로 통일하고 원시 평균값을 불필요하게 길게 쓰지 마십시오.
33. 확신도에 따라 표현 강도를 3단계로 구분하십시오. ① 자료로 직접 확인된 사실(수치·거래·규제 효력 등)은 '확인됨'으로 단정적으로, ② 자료 간 정합적 추론은 '분석됨'·'형성됨'으로, ③ 인과·개발 기대는 '배경으로 일부 작용함'·'단정하기 어려운 것으로 판단됨'으로 서술하십시오. 같은 유보 표현('~단정하지 않음' 등)을 여러 문장에 기계적으로 반복하지 말고 한 서술 안에서 유보는 핵심 1회로 집약하며, 확인된 사실까지 불필요하게 유보체로 흐리지 마십시오.
34. 서술 문체 기준 예시 — 좋은 예: '공업 0.085%로 평균 0.055%를 상회하나 실거래는 2건에 그쳐 금액 비중(28.7%) 대비 표본 구성 변화를 함께 확인할 필요가 있음.'(수치 근거 + 괴리 + 유의 1회) / 피해야 할 예: 근거 없이 '산업수요가 지가를 견인함.', 또는 매 문장 끝에 '~단정하지 않음'을 반복하거나 자료에 없는 수치·사업명을 창작하는 것.
35. 실거래 집계 기준은 신고일이 아니라 계약일입니다. 2026년 7월 이전은 월단위로 비교하고, 2026년 7월부터는 transactions.permit_split(지정 전·후 읍면동 건수·금액·일평균)과 transactions.weekly(토요일~금요일 주간 창, 매주 금요일 작성)를 함께 인용하십시오. 지정 전 13일·지정 후 18일로 일수가 다르므로 건수만 비교하지 말고 일평균건수 또는 일평균증감률을 함께 제시하십시오. 계약일 기준 집계는 부동산거래 신고기한(계약일부터 30일) 때문에 최근 구간이 과소집계 상태일 수 있으므로, 지정 후 구간이나 최근 주차의 감소를 규제 효과 또는 확정된 거래 위축으로 단정하지 말고 다음 갱신분에서 재확인이 필요하다고 명시하십시오.
36. 반도체 영향권 표기는 송정동·도산동·신촌동·소촌동·우산동·황룡동을 '반도체 직접 영향권', 선암동·운수동을 '연계 영향권'으로 통일하십시오. '핵심 영향권'·'인접 영향권' 표현은 사용하지 말고, 직접 영향권 표기는 군공항 예정지 인접성에 따른 보고서 분석 분류일 뿐 개별 동의 수혜나 가격 상승을 뜻하지 않는다는 점을 유지하십시오.
37. weekly_field_survey는 주간 현장점검 보고서(부록 5) 양식에 그대로 들어가는 서술입니다. 서술 단위는 개별 동이 아니라 두 개의 묶음이며, group_narratives에 group="직접 영향권"(송정동·도산동·신촌동·소촌동·우산동·황룡동)과 group="간접 영향권"(해당 시군구 내 직접 영향권 외 읍·면·동)을 이 순서로 각각 1개씩 작성하십시오. 각 묶음 서술은 6~9문장이며 **수치 나열이 아니라 시장 해석**이 중심입니다. 다음 순서로 쓰십시오. (1) 시장 여건: 반도체 첨단국가산단·군공항 이전 등 개발 이슈와 해당 묶음의 위치 관계, regional_profile의 권역·수요 배경(직접 영향권은 예정지 인접성, 간접 영향권은 대체 수요·연계 생활권 관점). (2) 규제 영향: 2026-07-14 토지거래허가구역 지정 이후 계약 흐름이 어떻게 달라졌는지를 transactions.permit_split.influence_group_rows로 해석하되, 허가 절차에 따른 계약 지연·관망과 신고기한 30일에 따른 과소집계를 구분해 서술. (3) 매수 수요의 성격: transactions.influence_group_month_rows의 외지인 건수·금액 비중, 법인 매수 비중, 토지만 거래 비중, 녹지·관리·개발제한 등 대면적 저가 토지 비중, 최다 용도지역을 근거로 '반도체 기대에 따른 외지 자본·법인 매수인지, 실거주·생활권 수요인지'를 구분해 판단. (4) 계약 사례: transactions.weekly.influence_group_recent_trades에서 1~2건만 골라 계약일·소재지·용도지역·면적·금액·단가·거래주체·매수자 시군구·지정 전후 구분을 근거로 제시. (5) 매도자 호가·매물 동향과 매수 문의: **field_monitoring(중개사무소 모니터링 청취 기록)이 available이면 그 내용(overall·groups·dongs·price_trend·outlook)을 1차 근거로 삼아** 문의 증가, 매물 회수, 호가 상향, 관망세, 호가와 매수 희망가격의 격차 등을 서술하고 '중개사무소 청취 기준'임을 밝히십시오. field_monitoring이 없고 listings 자료도 없으면 '매물 자료와 청취 기록이 확인되지 않아 호가 방향은 조사담당자 현장 확인이 필요함'으로 적고, 호가 상승·매물 회수·문의 증가를 자료 없이 지어내지 마십시오. (6) 지가변동률과의 관계 및 다음 주차 확인사항. 수치는 문장당 1~2개까지만 근거로 사용하고 같은 수치를 반복하지 마십시오. 거래가 없으면 수치를 만들지 말고 '최근 13주 신고된 계약 건이 확인되지 않음'으로 적으십시오. overall(시군구 내 전반적인 토지시장 동향)은 4~6문장으로 시군구 전체의 규제·수요 국면을 요약하고 묶음별 서술과 문장을 중복시키지 마십시오. price_trend(매매 호가 및 실거래 현황)은 4~6문장으로 발표·규제 이후의 매도 호가 움직임, 매물 회수·재등록, 매수 문의의 성격(실수요/개발 기대), 호가와 매수 희망가격의 격차, 호가 상승이 실거래가격으로 이어졌는지 여부를 서술하되 field_monitoring 기록과 실거래 계약 사례만 근거로 쓰십시오. market_outlook(향후 시장 전망)은 4~6문장으로 국가산단 최종 지정·군공항 이전·기업 투자 등 후속 절차의 불확실성, 토지거래허가구역의 투기수요 억제 효과, 매도 호가와 매수 관망이 지속될 경우의 거래량 전망, 다음 달 이후 확인할 지표(허가 실적, 거래량, 실거래가격, 매물 회수 여부)를 직접·간접 영향권으로 나눠 제시하되 확정된 가격 방향으로 단정하지 마십시오.
38. land_price.has_previous_column이 true이면 원표에 당월·전월·누계가 함께 있는 양식입니다. land_price_interpretation·zone_change_reasoning·use_change_reasoning·emd_change_reasoning에서 당월 변동률만 쓰지 말고 **전월 변동률과 전월 대비 변동폭(%p, land_price.change_vs_previous·emd_top[].전월대비)**을 함께 인용해 상승·둔화 여부를 판단하십시오. 변동폭 단위는 %p로 쓰고 변동률(%)과 혼용하지 마십시오. 전월 대비 확대·축소가 특정 용도지역·이용상황·읍면동에 집중되는지, 실거래 건수·금액 분포와 어긋나는지를 함께 짚고, 원표에 전월 열이 없으면(구 양식) 전월 대비 서술을 생략하십시오.
39. 6장 ② 권역별 동향은 반도체 영향권 2권역 구조로 작성합니다. influence_zone_narratives에 zone="반도체 직접 영향권"(송정동·도산동·신촌동·소촌동·우산동·황룡동)과 zone="간접 영향권"(해당 시군구 내 직접 영향권 외 읍·면·동)을 이 순서로 각각 1개씩 공백 포함 600자 이내로 작성하십시오. 각 권역 서술은 (1) 군공항 반도체 국가산단 예정지와의 위치 관계·토지이용 특성 → (2) transactions.influence_group_report_rows의 신고일 기준 당월 건수·금액과 건수·금액 비중, 외지인 거래 → (3) 해당 권역에서 거래가 두드러진 읍·면·동을 이름과 건수로 제시 → (4) land_price.emd_top의 당월·전월 변동률로 권역 내 가격 변화 폭과 편차 → (5) transactions.permit_split.influence_group_rows의 계약일 기준 지정 전·후 변화 → (6) 확인 필요사항 순으로 쓰되, 산단 인접성만으로 가격 상승을 단정하지 마십시오. regional_zone_narratives(생활권·산업권 5개 권역)는 세부 참고용으로 유지하되 영향권 서술과 문장을 반복하지 마십시오. market_overview·regional_analysis·key_takeaways도 직접 영향권의 가격·거래·규제 변화를 우선 제시하고 간접 영향권과 대비되도록 작성하십시오.
40. 서술 문체는 담당지역 조사자가 작성하는 현장 보고서 어조로 통일합니다. 문장은 '~임', '~됐음', '~나타났음', '~증가했음', '~조사됨', '~분석됨', '~판단됨' 등 과거형·조사보고체로 끝맺고 설명형 존댓말·구어체는 쓰지 마십시오. 권역·영향권 서술은 (1) 지역 성격과 개발 이슈와의 위치 관계 → (2) 현장 모니터링 결과(문의 증감, 매물 회수·재등록, 매도자·매수자 우위 분위기) → (3) 호가 수준의 구체적 변화(평당·필지 단위 금액, 발표 전후 비교) → (4) 실거래 신고자료 분석 결과(읍·면·동별 전월 대비 건수 변화는 transactions.emd_month_week.rows의 전월건수·당월건수·전월대비건수증감 사용) → (5) 토지거래허가구역 지정 이후 문의·계약·관망세 변화 → (6) 종합 흐름 요약과 다음 확인사항 순으로 이어 쓰되, 각 단락이 수치 나열이 아니라 시장 흐름의 인과로 읽히게 하십시오. (2)(3)의 문의·호가·매물 정보는 field_monitoring(중개사무소 청취 기록)에 있는 내용만 사용하고, 기록이 없으면 해당 단계를 생략하거나 조사담당자 확인 필요로 적으며 평당 단가 등 수치를 창작하지 마십시오.
41. listings.available이 true이고 template_only가 false이면 매물 자료가 실제로 집계된 것이므로 listing_interpretation을 반드시 작성하십시오. total_current·total_previous·change_vs_previous로 전체 매물 증감을, use_rows·emd_rows·zone_rows로 이용상황·읍면동·용도지역별 매물 증감과 평균·중위 호가단가를, price_changes로 동일 매물(읍면동+지번+이용상황+대지면적 매칭)의 호가 인상·인하·유지 건수와 신규 등록·매물 회수, 조정사례를 인용하십시오. direct_zone_summary로 반도체 직접 영향권 매물 수와 비중·전월 대비 증감을 제시하고, 전체 매물이 줄어드는 가운데 직접 영향권 매물이 늘거나 호가가 오르는지 등 대비를 명확히 쓰십시오. 매물 호가는 매도 희망가격이므로 실거래가격과 구분해 서술하십시오.
42. 5장 ① 거래 규모 및 자본 흐름 종합(detailed_transaction_analysis)은 개별 고가거래 나열이 아니라 시장 전체 구조를 공백 포함 300자 이내로 간결하게 요약하십시오. (1) 당월 거래 규모와 전월·전년동월 대비, (2) 반도체 직접 영향권과 간접 영향권의 건수·금액 비중 대비와 그 성격(직접 영향권의 소액·다건 개발 기대 매수 대 간접 영향권 산업권의 대면적 고액 계약), (3) 외지인·법인 등 매수 주체로 본 자본 흐름, (4) 소수 고액 계약에 금액이 좌우되는 구조에 대한 유의 순으로 쓰십시오. 개별 계약의 소재지·면적·단가 등 상세는 ③ 고가 실거래 TOP10 표와 4장에서 다루므로 여기서 반복하지 마십시오. 고가거래 개별 해설이 필요하면 top_trades의 지목(현실지목)과 거래주체·매수자 시군구를 근거로 ③ 항목 서술에서 다루십시오.
43. 7장 ②③ zone_change_reasoning·use_change_reasoning은 실거래를 중심에 놓고 각각 공백 포함 300자 이내로 작성하십시오. zone_change_reasoning은 transactions.zone_top으로 (1) 당월 실거래가 어느 용도지역에 건수·금액으로 집중됐는지, (2) 건수비중과 금액비중이 어긋나는 용도지역과 그 배경이 된 거래의 성격(top_trades의 용도지역·거래주체), (3) 그 거래 분포가 land_price.current의 용도지역별 지가변동률 상·하위와 같은 방향인지 순으로 쓰십시오. use_change_reasoning은 transactions.jimok_top(지목별 실거래)과 use_top(물건구분)으로 (1) 어느 지목·물건구분의 실거래가 건수·금액으로 많았는지, (2) top_trades에서 확인되는 해당 지목 대표 거래의 성격, (3) land_price.current의 이용상황별(전·답·주거용·상업용·공업용·임야) 지가변동률과의 대응 관계 순으로 쓰십시오. 지가변동률 수치는 근거로 1~2개만 인용하고 변동률 나열로 분량을 채우지 말며, 실거래가 없는 구분은 '당월 실거래 없음'으로 적고, 특정 거래가 가격을 '견인함'으로 단정하지 마십시오.
44. 9장 emd_change_reasoning과 emd_area_narratives 상위 5개 지역은 반도체 클러스터 개발 기대를 중심에 놓고 서술하십시오. 각 지역 서술은 첫 문장에서 해당 지역의 토지이용·입지 특성을 한 줄로 간략히 제시하고(예: 역세권 상업·주거지와 노후 단독주택 혼재, 공항 인접 저층 주거지, 강변 농경지·생산녹지 위주 등 judgment.top_areas의 land_use·rear_demand·transport·commercial_area 중 해당 요인만 선택), 그 다음에 개발 기대와 가격·거래 변화를 연결하십시오. 다섯 지역에 같은 권역 설명을 그대로 반복하지 말고 동별로 구분되는 특성만 쓰십시오. 각 지역이 직접 영향권인지 밝히고, 후보지 발표 이후 매수 문의 증가와 매도자 우위 전환, 호가 상향·매물 회수, 발표 직후 거래량 증가와 토지거래허가구역 지정 이후 관망세 전환의 흐름을 지역 여건에 맞게 연결하되, field_monitoring 기록이 없는 호가·문의 정보는 창작하지 말고 실거래·매물 자료로 확인되는 범위에서 쓰십시오.
45. transactions.data_type은 실거래 자료구분입니다. '확정치'는 부동산원 실거래집계 상세내역(계약해제·승인불가 제외, 거래번호 단위 계약묶음, 개별공시지가 포함)이고 '잠정치'는 신고 원장(RTMS_KEY 계약묶음)입니다. executive_summary와 conclusion에서 어느 자료구분으로 집계했는지 한 번 밝히고, 확정치일 때는 top_trades의 공시지가_원㎡·공시지가배율을 근거로 고가 계약의 가격 수준을 개별공시지가와 비교해 서술하십시오(배율은 단일 필지 토지만 거래에서만 산출됨). 잠정치와 확정치는 묶음 기준·면적 산정이 달라 수치가 다르므로 두 자료의 수치를 한 문장에서 비교하거나 합산하지 마십시오.
46. development_projects.available이 true이면 development_project_updates를 development_projects.projects와 같은 순서·사업명(project)으로 작성하십시오. stage(추진단계)는 history·progress에 기록된 경과와 수집 기사 제목으로 확인되는 범위에서 구상·발표·예타·지정(고시)·보상·착공·분양·준공 중 현재 단계를 한 구절로 적고, 확인되지 않으면 '추진단계 확인 필요'로 두십시오. monthly_progress(당월 진행상황)는 개발사업현황 입력란에 그대로 옮길 한 줄 요약으로, 기준월(window.months의 마지막 달)의 progress와 news 제목만 근거로 공백 포함 50자 이내의 명사형 한 구절로 쓰고(예: '군공항 이전 후보지 무안군 망운면 확정'), 기준월 progress가 이미 있으면 그 문구를 그대로 쓰며, 기준월 근거가 없으면 '관련 보도 확인되지 않음'으로 두십시오. recent_progress(최근 2개월 뉴스 요약)는 development_projects.window 기간(전월~기준월)의 news(보도일·매체·제목)와 progress만 근거로 뉴스 내용을 공백 포함 100자 이내로 요약하고(월을 밝힘, 예: '8월 사업기간 1년 연장 추진 보도. 9월 보도는 수집되지 않음.'), 해당 사업에 news_summary가 이미 있으면 그 문구를 그대로 쓰십시오. 수집된 것은 기사 제목뿐이므로 제목에 없는 수치·일정·금액·기업명을 보태지 말고, 제목만으로 사실관계가 불분명하면 '~로 보도됨'·'원문 확인 필요'로 적으십시오. 주민 요구·검토·가능성·전망 보도는 확정된 사실과 구분하고, 해당 기간 news와 progress가 모두 없는 사업은 recent_progress를 '최근 2개월 관련 보도 확인되지 않음'으로 두십시오. impact는 그 진척이 기준월 지가변동률·실거래·매물 지표와 연결되는지(또는 아직 연결되지 않았는지)를 한 문장으로 쓰되 지가 상승의 직접 원인으로 단정하지 마십시오. 다른 시군구 소재 사실(예: 함평군 투자협약, 무안군 이전 후보지)은 소재지를 밝혀 구분하십시오.
47. monthly_top_news는 development_projects.monthly_headlines(기준월에 보도된 해당 시군구 부동산·개발사업 주요 뉴스)와 같은 순서·건수로 작성하십시오. title·source·source_date는 수집된 제목·매체·보도일을 그대로 옮기고, summary는 해당 monthly_headlines 항목에 summary(조사담당자가 원문을 확인해 적어 둔 500자 이내 요약)가 있으면 그 문구를 그대로 옮기고, 없으면 제목에서 확인되는 내용만 한 문장으로 쓰며, impact는 기준월 토지시장(지가변동률·실거래·매물·규제)에 주는 시사점을 한 문장으로 쓰되 제목에 없는 사실을 보태지 마십시오. monthly_headlines가 비어 있으면 빈 배열로 두십시오. news_items(부록 2)에서 기준월 발생 이슈를 쓸 때도 이 수집 기사와 development_projects.projects[].news를 근거로 활용하되 출처는 수집된 매체명과 보도일을 그대로 기재하십시오.
48. 분량 제한(모두 공백 포함. 한도를 넘는 부분은 보고서에서 마지막 완결 문장까지만 표시되므로 한도 안에서 문장을 끝맺으십시오): 4장 external_buyer_analysis 500자 이내 / 5장 detailed_transaction_analysis 300자 이내와 zone_use_transaction_interpretation 200자 이내(5장 합계 500자 이내) / 6장 ② regional_analysis 300자 이내와 influence_zone_narratives 권역별 600자 이내(합계 1,500자 이내) / 7장 zone_change_reasoning·use_change_reasoning 각 300자 이내 / 9장 emd_area_narratives 1~5위 각 200자 이내. 9장의 구성과 문체는 49번 규칙을 따르십시오. 분량이 줄어든 장에서는 개별 계약 나열보다 구조와 핵심 수치 2~4개를 우선하십시오.
49. 9장(emd_change_reasoning·emd_area_narratives)은 이 규칙이 12·23·24·44번보다 우선합니다. emd_area_narratives는 judgment.top_areas와 같은 순서·area로 순위별 공백 포함 200자 이내로 쓰되, (1) 첫 문장에서 그 지역을 한 구절로 간략히 설명하고(top_areas[].description을 바탕으로 하며 다섯 지역에 같은 설명을 반복하지 않음), (2) 당월 실거래(trades: 용도지역·지목·물건구분·금액·외지인 여부)나 매물(listing: 매물 수 증감, 동일 매물의 호가 인상·인하 건수, 주요 이용상황)이 있으면 그 지역의 토지이용 특성과 연결해 무엇이 거래되고 매물로 나왔는지 서술하며, (3) 마지막 문장은 상승 또는 하락 사유로 끝맺으십시오(예: '~개발 기대가 실거래와 호가 인상으로 이어지며 상승함'). 실거래와 매물이 모두 없으면 (2)를 건너뛰고 지역 특성과 reason_hint·influence_note에 맞는 상승·하락 사유를 서술하십시오. 사유는 평서형으로 끝맺되 description·trades·listing·reason_hint·regional_issues·development_projects에 있는 사실 범위에서만 쓰고, 없는 수치·사업명·호가는 만들지 마십시오. 지가변동률 수치는 카드 제목에 표시되므로 본문에서 반복하지 않아도 됩니다. 9장에는 '조사담당자', '현장 확인', '확인이 필요함', '확인할 필요가 있음', '판단구분', '가격조사 중심', '단정하기 어려움' 같은 확인 요청·유보·판단 기준 문장을 쓰지 말고(이런 문장은 보고서에서 삭제됨), emd_change_reasoning도 같은 기준으로 상위 지역에 공통된 변동 사유만 서술하십시오.

[출력 JSON 구조]
{json.dumps(schema, ensure_ascii=False, indent=2)}

[result.json]
{json.dumps(build_llm_data_view(result), ensure_ascii=False, separators=(",", ":"), default=str)}
"""


def h(value: Any) -> str:
    return html.escape("" if value is None else str(value))


def pick(mapping: dict[str, Any], *keys: str, default: Any = None) -> Any:
    for key in keys:
        if isinstance(mapping, dict) and key in mapping:
            return mapping[key]
    return default


def fmt_num(value: Any, digits: int = 1, suffix: str = "") -> str:
    if value is None:
        return "-"
    try:
        v = float(value)
        if math.isnan(v):
            return "-"
        if digits <= 0:
            return f"{v:,.0f}{suffix}"
        return f"{v:,.{digits}f}{suffix}"
    except Exception:
        return f"{value}{suffix}"


def fmt_signed(value: Any, suffix: str = "%") -> str:
    if value is None:
        return "-"
    try:
        v = float(value)
        cls = "up" if v > 0 else "dn" if v < 0 else "muted"
        sign = "+" if v > 0 else ""
        return f'<span class="{cls}">{sign}{v:.1f}{suffix}</span>'
    except Exception:
        return h(value)


def table_html(rows: list[dict[str, Any]], columns: list[tuple[str, str]], empty: str = "자료 없음") -> str:
    if not rows:
        return f'<p class="note">{h(empty)}</p>'
    head = "".join(f"<th>{h(label)}</th>" for _, label in columns)
    body = []
    for i, row in enumerate(rows, 1):
        cls = "hot" if i <= 3 else ""
        cells = "".join(f"<td>{h(row.get(key, '-'))}</td>" for key, _ in columns)
        body.append(f'<tr class="{cls}">{cells}</tr>')
    return f'<table><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table>'


def blank_table_html(columns: list[tuple[str, str]]) -> str:
    """자료를 지정하지 않은 항목은 열 제목과 빈 입력행만 유지한다."""
    head = "".join(f"<th>{h(label)}</th>" for _, label in columns)
    cells = "".join('<td class="blank-cell">&nbsp;</td>' for _ in columns)
    return f'<table class="template-only"><thead><tr>{head}</tr></thead><tbody><tr>{cells}</tr></tbody></table>'


def build_influence_zone_summaries(result: dict[str, Any]) -> list[dict[str, Any]]:
    """6장 권역별 동향을 반도체 직접 영향권 / 간접 영향권 2권역으로 묶어 계산값을 정리한다."""
    region = str((result.get("meta", {}) or {}).get("region") or "")
    if "광산구" not in region:
        return []
    tx = result.get("transactions", {}) or {}
    lp = result.get("land_price", {}) or {}
    profile = result.get("regional_profile", {}) or {}
    report_rows = {str(row.get("구분")): row for row in (tx.get("influence_group_report_rows", []) or [])}
    permit_rows = {str(row.get("구분")): row for row in ((tx.get("permit_split", {}) or {}).get("influence_group_rows", []) or [])}
    contract_rows = {str(row.get("구분")): row for row in (tx.get("influence_group_month_rows", []) or [])}
    direct_dongs = semiconductor_dongs("direct")
    direct_names = {norm_text(dong) for dong in direct_dongs}

    def rate_rows_for(direct: bool) -> list[dict[str, Any]]:
        out = []
        for row in (lp.get("emd_top", []) or []):
            tokens = [norm_text(token) for token in re.split(r"[,·/]", str(row.get("읍면동") or "")) if norm_text(token)]
            if not tokens:
                continue
            matched = any(token in direct_names for token in tokens)
            if matched == direct:
                out.append(row)
        return out

    summaries: list[dict[str, Any]] = []
    for label, direct in [("반도체 직접 영향권", True), ("간접 영향권", False)]:
        key = "직접 영향권" if direct else "간접 영향권"
        rates = rate_rows_for(direct)
        rate_values = [(row.get("읍면동"), as_number(row.get("당월변동률"))) for row in rates]
        rate_values = [(name, value) for name, value in rate_values if value is not None]
        top_rate = max(rate_values, key=lambda item: item[1]) if rate_values else None
        low_rate = min(rate_values, key=lambda item: item[1]) if rate_values else None
        zones = []
        if direct:
            for dong in direct_dongs:
                zone = match_regional_zone(dong, profile) or {}
                if zone.get("zone") and zone.get("zone") not in zones:
                    zones.append(str(zone.get("zone")))
        else:
            for zone in profile.get("zones", []) or []:
                dongs = [norm_text(x) for x in (zone.get("dongs") or [])]
                if any(x not in direct_names for x in dongs):
                    zones.append(str(zone.get("zone")))
        summaries.append({
            "구분": label,
            "대상": "·".join(direct_dongs) if direct else f"{region.split()[-1]} 직접 영향권 외 읍·면·동",
            "구성권역": zones,
            "산업·개발거점": (
                ["호남권 반도체 첨단국가산업단지(광주 군공항 부지 후보지)", "소촌농공단지"] if direct else
                [item.get("name") for item in (profile.get("industrial_complexes", []) or [])
                 if item.get("name") and "반도체" not in str(item.get("name"))][:4]
            ),
            "당월_신고일기준": report_rows.get(key, {}),
            "당월_계약일기준": contract_rows.get(key, {}),
            "허가구역_전후": permit_rows.get(key, {}),
            "지가_최고": {"읍면동": top_rate[0], "당월변동률": top_rate[1]} if top_rate else {},
            "지가_최저": {"읍면동": low_rate[0], "당월변동률": low_rate[1]} if low_rate else {},
            "지가_행": rates[:6],
        })
    return summaries


def build_influence_zone_narratives(result: dict[str, Any]) -> list[dict[str, str]]:
    """직접·간접 영향권 2권역 서술 초안(LLM 미반영 시 표시)."""
    summaries = build_influence_zone_summaries(result)
    lp_current = (result.get("land_price", {}) or {}).get("current", {}) or {}
    out: list[dict[str, str]] = []
    for item in summaries:
        report = item.get("당월_신고일기준", {}) or {}
        permit = item.get("허가구역_전후", {}) or {}
        top, low = item.get("지가_최고", {}), item.get("지가_최저", {})
        direct = item["구분"].startswith("반도체 직접")
        parts = [
            f"{item['구분']}({item['대상']})은 "
            + (
                "광주 군공항 반도체 첨단국가산단 예정지에 접한 읍·면·동으로, 송정역 역세권과 기존 시가지, 공항 주변 토지가 함께 분포하는 구역임."
                if direct else
                "산단 예정지와 직접 접하지 않는 읍·면·동으로, 수완·신가 생활권 주거지와 하남·진곡·평동·빛그린 산업권, 첨단·비아 혁신권이 함께 포함된 구역임."
            ),
            f"당월 실거래는 신고일 기준 {fmt_num(report.get('거래건수'), 0)}건·{fmt_num(report.get('거래금액_억원'), 1)}억 원"
            f"(건수 비중 {fmt_num(report.get('건수비중'), 1)}%·금액 비중 {fmt_num(report.get('금액비중'), 1)}%)이며 "
            f"외지인 거래는 {fmt_num(report.get('외지인거래건수'), 0)}건으로 확인됨.",
        ]
        if top and low:
            if top.get("읍면동") == low.get("읍면동"):
                parts.append(f"당월 지가변동률은 {top.get('읍면동')} {fmt_num(top.get('당월변동률'), 3, '%')}으로 확인됨.")
            else:
                parts.append(
                    f"당월 지가변동률은 {top.get('읍면동')} {fmt_num(top.get('당월변동률'), 3, '%')}에서 "
                    f"{low.get('읍면동')} {fmt_num(low.get('당월변동률'), 3, '%')} 범위로 확인되어 "
                    f"시군구 평균 {fmt_num(lp_current.get('평균'), 3, '%')}과 비교됨."
                )
        if permit:
            parts.append(
                f"계약일 기준 토지거래허가구역 지정 전 {fmt_num(permit.get('지정전_거래건수'), 0)}건에서 "
                f"지정 후 {fmt_num(permit.get('지정후_거래건수'), 0)}건으로 축소된 것으로 확인됨."
            )
        parts.append(
            "산단 예정지 인접성만으로 가격 상승을 단정하지 않고 허가 실적과 실수요 계약 여부를 함께 확인할 필요가 있음."
            if direct else
            "직접 영향권 규제에 따른 대체 수요 이동 여부와 산업권 실수요를 구분해 확인할 필요가 있음."
        )
        out.append({"zone": item["구분"], "narrative": " ".join(parts)})
    return out


def build_field_survey_narratives(result: dict[str, Any]) -> dict[str, Any]:
    """주간 현장점검 보고서 양식의 서술 초안(지역 특성 + 계약일 기준 실거래 내용)을 만든다.

    LLM이 같은 키로 다시 쓰면 그 문장이 우선 사용되고, LLM 결과가 없으면 이 Python 초안이 표시된다.
    수치는 result의 계산값만 쓰고 인과 단정을 피한 문장 구조를 유지한다.
    """
    meta = result.get("meta", {}) or {}
    region = str(meta.get("region") or "조사지역")
    tx = result.get("transactions", {}) or {}
    weekly = tx.get("weekly", {}) or {}
    permit = tx.get("permit_split", {}) or {}
    profile = result.get("regional_profile", {}) or {}
    lp = result.get("land_price", {}) or {}
    lp_current = lp.get("current", {}) if lp.get("available") else {}

    week_rows = weekly.get("weekly_rows", []) or []
    latest = week_rows[-1] if week_rows else {}
    prior = week_rows[-2] if len(week_rows) > 1 else {}
    week_count = len(week_rows)
    emd_weekly = {norm_text(row.get("읍면동")): row for row in (weekly.get("weekly_emd_rows", []) or [])}
    permit_rows = {norm_text(row.get("읍면동")): row for row in (permit.get("emd_rows", []) or [])}
    recent_trades = weekly.get("emd_recent_trades", {}) or {}
    lp_emd = {norm_text(row.get("읍면동")): row for row in (lp.get("emd_top", []) or [])}

    def trend_word(current: Any, previous: Any) -> str:
        cur_value, prev_value = as_number(current), as_number(previous)
        if cur_value is None or prev_value is None:
            return "확인됨"
        if cur_value > prev_value:
            return "전주 대비 증가한 것으로 확인됨"
        if cur_value < prev_value:
            return "전주 대비 감소한 것으로 확인됨"
        return "전주와 같은 수준으로 확인됨"

    overall = (
        f"{region}의 계약일 기준 주간 거래는 기준주({latest.get('주차') or '-'}) "
        f"{fmt_num(latest.get('거래건수'), 0)}건·{fmt_num(latest.get('거래금액_억원'), 1)}억 원으로 "
        f"{trend_word(latest.get('거래건수'), prior.get('거래건수'))}"
        f"(전주 {fmt_num(prior.get('거래건수'), 0)}건). "
        f"최근 {week_count}주 누계는 {fmt_num(sum(as_number(row.get('거래건수')) or 0 for row in week_rows), 0)}건이며 "
        f"기준주 외지인 거래는 {fmt_num(latest.get('외지인거래건수'), 0)}건으로 확인됨. "
    )
    if permit.get("available") and permit.get("mode") == "split":
        totals = permit.get("totals", {}) or {}
        overall += (
            f"토지거래허가구역 지정 전({(permit.get('periods', {}).get('pre', {}) or {}).get('label')}) "
            f"{fmt_num((totals.get('pre') or {}).get('거래건수'), 0)}건·일평균 {fmt_num((totals.get('pre') or {}).get('일평균건수'), 2)}건에서 "
            f"지정 후({(permit.get('periods', {}).get('post', {}) or {}).get('label')}) "
            f"{fmt_num((totals.get('post') or {}).get('거래건수'), 0)}건·일평균 {fmt_num((totals.get('post') or {}).get('일평균건수'), 2)}건으로 집계됨. "
            "다만 계약일 기준 집계는 신고기한(30일) 내 미신고 계약이 이후 갱신분에 추가되므로 감소폭을 규제 효과로 단정하기 어려운 것으로 판단됨. "
        )
    if lp_current.get("평균") is not None:
        overall += f"직전 조사월 지가변동률 평균은 {fmt_num(lp_current.get('평균'), 3, '%')}으로 확인됨."

    # 서술 단위: 직접 영향권 6개 동 묶음 / 그 외 읍·면·동(간접 영향권) 묶음
    group_weekly = {str(row.get("구분")): row for row in (weekly.get("influence_group_rows", []) or [])}
    group_permit = {str(row.get("구분")): row for row in (permit.get("influence_group_rows", []) or [])}
    group_month = {str(row.get("구분")): row for row in (tx.get("influence_group_month_rows", []) or [])}
    group_trades = weekly.get("influence_group_recent_trades", {}) or {}
    direct_dongs = semiconductor_dongs("direct")
    direct_names = {norm_text(dong) for dong in direct_dongs}

    def dong_detail(names: list[str]) -> str:
        pieces = []
        for name in names:
            row = emd_weekly.get(norm_text(name)) or {}
            weeks = row.get("주별", []) or []
            current_week = weeks[-1] if weeks else {}
            pieces.append(
                f"{name} {fmt_num(row.get('합계건수', 0), 0)}건"
                f"(기준주 {fmt_num(current_week.get('거래건수', 0), 0)}건)"
            )
        return ", ".join(pieces)

    def trade_sentence(trades: list[dict[str, Any]]) -> str:
        if not trades:
            return ""
        head = trades[0]
        unit_text = f", 단가 {fmt_num(head.get('단가_원㎡'), 0)}원/㎡" if head.get("단가_원㎡") is not None else ""
        buyer_text = (
            f", 거래주체 {head.get('거래주체') or '미상'}"
            f"·매수자 {head.get('매수자시군구') or head.get('매수자지역') or '미상'}"
            f"({'외지인' if head.get('외지인') == 'Y' else '관내'})"
        )
        return (
            f"최근 계약은 {head.get('계약일')} {head.get('소재지')}({head.get('물건구분')}·{head.get('용도지역')}, "
            f"{fmt_num(head.get('면적_㎡'), 1)}㎡, {fmt_num(head.get('거래금액_억원'), 2)}억 원{unit_text}{buyer_text}, "
            f"{head.get('허가구역구분')} 계약)로 확인되며 최근 {week_count}주 계약 {len(trades)}건을 함께 검토함."
        )

    def permit_sentence(row: dict[str, Any]) -> str:
        if not row or permit.get("mode") != "split":
            return ""
        return (
            f"토지거래허가구역 지정 전 {fmt_num(row.get('지정전_거래건수'), 0)}건"
            f"(일평균 {fmt_num(row.get('지정전_일평균건수'), 2)}건)에서 지정 후 "
            f"{fmt_num(row.get('지정후_거래건수'), 0)}건"
            f"(일평균 {fmt_num(row.get('지정후_일평균건수'), 2)}건)으로 집계됨."
        )

    def rate_sentence(names: list[str]) -> str:
        values = [
            (name, as_number((lp_emd.get(norm_text(name)) or {}).get("당월변동률")))
            for name in names
        ]
        values = [(name, value) for name, value in values if value is not None]
        if not values:
            return ""
        high = max(values, key=lambda item: item[1])
        low = min(values, key=lambda item: item[1])
        if high[0] == low[0]:
            return f"직전 조사월 지가변동률은 {high[0]} {high[1]:.3f}%로 확인됨."
        return (
            f"직전 조사월 지가변동률은 {high[0]} {high[1]:.3f}%로 가장 높고 "
            f"{low[0]} {low[1]:.3f}%로 가장 낮게 확인됨."
        )

    group_month = {str(row.get("구분")): row for row in (tx.get("influence_group_month_rows", []) or [])}
    listings_available = bool((result.get("listings", {}) or {}).get("available"))

    def demand_sentence(label: str) -> str:
        row = group_month.get(label) or {}
        if not row.get("거래건수"):
            return ""
        top_zone = (row.get("용도지역구성") or [{}])[0]
        return (
            f"당월 매수 구성은 외지인 {fmt_num(row.get('외지인거래건수'), 0)}건"
            f"(건수 {fmt_num(row.get('외지인건수비중'), 1)}%·금액 {fmt_num(row.get('외지인금액비중'), 1)}%), "
            f"법인 매수 {fmt_num(row.get('법인매수건수'), 0)}건({fmt_num(row.get('법인매수비중'), 1)}%), "
            f"토지만 거래 {fmt_num(row.get('토지만거래비중'), 1)}%, "
            f"녹지·관리·개발제한 {fmt_num(row.get('녹지·관리·개발제한비중'), 1)}%로 확인되어 "
            f"{top_zone.get('용도지역', '주거지역')} 중심의 계약 구성이 나타남."
        )

    def listing_sentence() -> str:
        if listings_available:
            return "매물 자료를 함께 확인해 호가 조정과 매물 증감을 점검함."
        return (
            "매물 엑셀이 지정되지 않아 매도 호가 방향과 매물 회수·신규 등록 여부는 자료로 확인되지 않으므로 "
            "조사담당자의 중개업소 청취와 현장 확인이 필요함."
        )

    group_narratives: list[dict[str, str]] = []

    # ① 직접 영향권
    direct_row = group_weekly.get("직접 영향권", {})
    zone_labels = []
    for dong in direct_dongs:
        zone = match_regional_zone(dong, profile) or {}
        if zone.get("zone") and zone.get("zone") not in zone_labels:
            zone_labels.append(str(zone.get("zone")))
    direct_parts = [
        f"직접 영향권({'·'.join(direct_dongs)})은 광주 군공항 호남 반도체 첨단국가산단 예정지에 인접한 읍·면·동으로, "
        f"{', '.join(zone_labels) if zone_labels else '송정 생활권'} 중심의 역세권·기존 시가지 수요와 산단 조성 기대가 함께 작용하는 구역임.",
        "2026년 7월 14일 토지거래허가구역 지정으로 허가 절차와 이용목적 제한이 적용되면서 계약 성사 시점이 뒤로 밀리는 여건이 형성됨.",
        permit_sentence(group_permit.get("직접 영향권", {})),
        f"계약일 기준 당월 거래는 {fmt_num((group_month.get('직접 영향권') or {}).get('거래건수', 0), 0)}건·"
        f"{fmt_num((group_month.get('직접 영향권') or {}).get('거래금액_억원', 0), 1)}억 원이며 기준주는 "
        f"{fmt_num(direct_row.get('기준주거래건수', 0), 0)}건으로 "
        f"{trend_word(direct_row.get('기준주거래건수', 0), direct_row.get('전주거래건수', 0))}"
        f"(동별 {dong_detail(direct_dongs)}).",
        demand_sentence("직접 영향권"),
        trade_sentence(group_trades.get("직접 영향권", []) or []),
        listing_sentence(),
        rate_sentence(direct_dongs),
        "개발 기대와 규제 영향이 함께 작용하는 구간이므로 허가 신청·처리 동향과 실수요 계약 비중을 확인하고, "
        "개별 거래나 규제 시행 시점만으로 가격 방향을 단정하지 않음.",
    ]
    group_narratives.append({
        "group": "직접 영향권",
        "scope": "·".join(direct_dongs),
        "narrative": " ".join(part for part in direct_parts if part and part.strip()),
    })

    # ② 간접 영향권
    indirect_row = group_weekly.get("간접 영향권", {})
    indirect_names = [
        str(row.get("읍면동")) for row in (weekly.get("weekly_emd_rows", []) or [])
        if norm_text(row.get("읍면동")) not in direct_names and as_number(row.get("합계건수"))
    ][:6]
    indirect_parts = [
        f"간접 영향권은 {region.split()[-1] if region.split() else region} 내 직접 영향권 외 읍·면·동으로, "
        "산단 예정지와 직접 접하지 않으나 생활권·산업권 연계와 대체 수요 이동 여부를 함께 살필 구역임.",
        f"계약일 기준 당월 거래는 {fmt_num((group_month.get('간접 영향권') or {}).get('거래건수', 0), 0)}건·"
        f"{fmt_num((group_month.get('간접 영향권') or {}).get('거래금액_억원', 0), 1)}억 원이고 "
        f"최근 {week_count}주 거래는 {fmt_num(indirect_row.get('합계건수', 0), 0)}건·"
        f"{fmt_num(indirect_row.get('합계금액_억원', 0), 1)}억 원이며 기준주는 "
        f"{fmt_num(indirect_row.get('기준주거래건수', 0), 0)}건으로 "
        f"{trend_word(indirect_row.get('기준주거래건수', 0), indirect_row.get('전주거래건수', 0))}.",
        (f"거래가 확인된 주요 지역은 {dong_detail(indirect_names)}으로 확인됨." if indirect_names else
         f"최근 {week_count}주 계약이 확인된 읍·면·동은 없음."),
        permit_sentence(group_permit.get("간접 영향권", {})),
        trade_sentence(group_trades.get("간접 영향권", []) or []),
        rate_sentence(indirect_names),
        "직접 영향권 규제에 따른 수요 이동 여부는 단일 주차 수치로 단정하지 않고 주단위 누적 흐름으로 확인할 필요가 있음.",
    ]
    group_narratives.append({
        "group": "간접 영향권",
        "scope": f"{region.split()[-1] if region.split() else region} 직접 영향권 외 읍·면·동",
        "narrative": " ".join(part for part in indirect_parts if part and part.strip()),
    })

    outlook = (
        "토지거래허가구역 지정 이후 계약일 기준 거래는 감소 흐름으로 나타나나 신고기한(30일) 영향이 남아 있어 "
        "확정된 위축으로 단정하기 어려운 것으로 판단됨. 직접 영향권은 허가 대상 여부와 실수요 계약 비중, 외지인 매수 비중, "
        "호가 조정 여부를 주단위로 확인하고, 간접 영향권은 규제에 따른 대체 수요 이동 여부를 누적 주차 흐름으로 점검할 필요가 있음. "
        "산업단지 지정·보상 등 사업 추진단계가 확인되는 시점에 가격 변화와 함께 재점검할 필요가 있음."
    )
    monitoring = result.get("field_monitoring", {}) or {}
    if monitoring.get("available") and monitoring.get("price_trend"):
        price_trend = str(monitoring.get("price_trend"))
    else:
        top = (tx.get("top_trades", []) or [])[:1]
        head = top[0] if top else {}
        price_trend = (
            "매물 엑셀과 중개사무소 청취 기록이 확인되지 않아 매도 호가 방향과 매물 회수 여부는 자료로 확인되지 않음. "
            + (
                f"당월 최고가 계약은 {head.get('계약일')} {head.get('소재지')}"
                f"({head.get('용도지역')}, {fmt_num(head.get('면적_㎡'), 1)}㎡, "
                f"{fmt_num(head.get('거래금액_억원'), 2)}억 원, 거래주체 {head.get('거래주체')}·"
                f"매수자 {head.get('매수자지역')})로 확인됨. " if head else ""
            )
            + "호가 수준과 매수 문의 동향은 조사담당자의 중개업소 청취로 보완할 필요가 있음."
        )

    return {
        "overall": overall.strip(),
        "group_narratives": group_narratives,
        "price_trend": price_trend,
        "market_outlook": outlook,
    }


def default_narratives(result: dict[str, Any]) -> dict[str, Any]:
    tx = result.get("transactions", {})
    lp = result.get("land_price", {})
    listings = result.get("listings", {})
    pop = result.get("population", {})
    hh = result.get("households", {})
    age_population = result.get("age_population", {})
    mig = result.get("migration", {})
    rates = result.get("rates", {})
    regional_issues = result.get("regional_issues", {})
    regional_profile = result.get("regional_profile", {})
    issue_area = regional_issues.get("display_area") or "전남광주통합특별시 광주권역"

    cur = tx.get("current", {}) if tx.get("available") else {}
    chg = tx.get("change_vs_previous", {}) if tx.get("available") else {}
    lp_cur = lp.get("current", {}) if lp.get("available") else {}

    price_reasons = build_investigator_price_reason_narratives(result)

    profile_zone_narratives = [
        {
            "zone": zone.get("zone", ""),
            "narrative": (
                f"{zone.get('characteristics', '')}. 수요 배경은 {zone.get('demand_background', '')}으로 구분됨. "
                f"{zone.get('analysis_caution', '')}."
            ),
        }
        for zone in regional_profile.get("zones", [])
    ]
    industrial_names = ", ".join(item.get("name", "") for item in regional_profile.get("industrial_complexes", []) if item.get("name"))
    fact_transaction_narratives = build_fact_bound_transaction_narratives(result)

    return {
        "weekly_field_survey": build_field_survey_narratives(result),
        "influence_zone_narratives": build_influence_zone_narratives(result),
        "key_takeaways": [
            f"당월 토지 거래는 {pick(cur, '거래건수', default='-')}건, {fmt_num(pick(cur, '거래금액_억원'), 1, '억원')}으로 집계되었고 전월 대비 거래건수는 {fmt_num(pick(chg, '거래건수증감률'), 1, '%')} 변동했습니다.",
            f"당월 지가변동률 평균은 {fmt_num(pick(lp_cur, '평균'), 3, '%')}이며, 용도지역·이용상황별 차이를 함께 확인해야 합니다.",
            "대출금리, 인구, 세대수, 이동자수는 토지시장 방향을 단정하는 지표가 아니라 수요 기반과 매수여력의 배경지표로 해석합니다.",
        ],
        "executive_summary": (
            f"{result['meta']['region']} 토지시장은 업로드된 실거래 엑셀과 당월 지가동향 자료를 기준으로 분석했습니다. "
            f"실거래는 {pick(cur, '거래건수', default='-')}건, {fmt_num(pick(cur, '거래금액_억원'), 1, '억원')} 규모이며, "
            f"외지인 거래는 {pick(cur, '외지인거래건수', default='-')}건으로 집계되었습니다. "
            f"지가변동률은 조사담당자가 가격 변화·실거래·지역시장동향을 종합해 판단한 결과이므로, 거래 집중 지역과 지가 상승 지역은 구분해 해석해야 합니다."
        ),
        "listing_interpretation": (
            ""
            if listings.get("template_only")
            else "매물 파일은 전월·당월 2개가 모두 있을 때 증감 판단이 가능합니다."
            if not listings.get("has_prev_and_current")
            else "전월·당월 매물 파일이 모두 확인되어 매물 증감과 호가 조정 흐름을 함께 점검할 수 있습니다."
        ),
        "population_interpretation": (
            f"인구는 {fmt_num(pop.get('current'), 0, '명')}, 세대수는 {fmt_num(hh.get('current'), 0, '세대')}, "
            f"순이동은 {fmt_num(mig.get('current'), 0, '명')}으로 정리됩니다. 해당 지표는 토지시장 실수요 기반의 보조 자료로 활용합니다."
        ),
        "rate_interpretation": (
            f"예금은행 대출금리는 최신 공표월 {rates.get('latest_month') or '-'} 기준 주택담보대출 {fmt_num(rates.get('mortgage_rate'), 2, '%')}, "
            f"전세자금대출 {fmt_num(rates.get('jeonse_rate'), 2, '%')}, 중소기업대출 {fmt_num(rates.get('sme_rate'), 2, '%')}, "
            f"기업일반자금대출 {fmt_num(rates.get('business_rate'), 2, '%')}로 확인됩니다. "
            "금리는 자금조달 부담과 관망심리를 설명하는 배경지표로만 사용합니다."
            if rates.get("available")
            else "대출금리 파일이 없어 금리 배경지표는 이번 LLM 자료에서 제외되었습니다. 수집기에서 대출금리를 함께 수집하면 rates 블록에 자동 반영됩니다."
        ),
        "zone_change_reasoning": price_reasons["zone_change_reasoning"],
        "use_change_reasoning": price_reasons["use_change_reasoning"],
        "regional_issue_interpretation": (
            f"{issue_area}에서는 광산구 군공항 부지 반도체 클러스터 조성 이슈와 2026년 7월 14일 토지거래허가구역 지정 효력 발생을 함께 점검해야 합니다. "
            "해당 이슈는 개발 기대와 투기성 거래 억제 요인이 동시에 작용할 수 있으므로, 지가변동률의 직접 원인으로 단정하지 않고 거래량·외지인 거래·매물 호가 변화를 분리해 관찰합니다."
            if regional_issues.get("available")
            else ""
        ),
        "regional_analysis": (
            "광산구는 송정역·도심전환, 수완·신가 생활, 첨단·비아 혁신, 하남·진곡 제조산업, 평동·빛그린 산업확장의 5개 분석권역으로 구분함. "
            "권역별 토지이용과 수요 기반이 상이하므로 조사담당자가 확인한 당월 가격 변화와 실거래를 대조하고 개발 기대와 확인된 실수요를 분리해 해석함."
            if regional_profile.get("available") else ""
        ),
        "regional_zone_narratives": profile_zone_narratives,
        "industrial_complex_analysis": (
            f"광산구 산업입지는 {industrial_names}을 중심으로 검토함. 기존 산업단지와 광산구·북구 연계 산업단지, 연계 개발지구를 구분하고 기업 입주·가동·고용과 사업 추진단계가 확인된 범위에서만 토지수요와 연결함."
            if industrial_names else ""
        ),
        "emd_area_narratives": build_investigator_emd_narratives(result),
        **fact_transaction_narratives,
    }


def build_compact_html_report(result: dict[str, Any], narratives: dict[str, Any] | None = None) -> str:
    n = {**default_narratives(result), **(narratives or {})}
    meta = result.get("meta", {})
    tx = result.get("transactions", {})
    lp = result.get("land_price", {})
    pop = result.get("population", {})
    hh = result.get("households", {})
    mig = result.get("migration", {})
    rates = result.get("rates", {})
    listings = result.get("listings", {})
    regional_issues = result.get("regional_issues", {})

    cur = tx.get("current", {}) if tx.get("available") else {}
    prev_chg = tx.get("change_vs_previous", {}) if tx.get("available") else {}
    lp_cur = lp.get("current", {}) if lp.get("available") else {}
    lp_cum = lp.get("cumulative", {}) if lp.get("available") else {}

    monthly_rows = [dict(row) for row in (tx.get("monthly_rows", []) if tx.get("available") else [])]
    public_rate_lookup = {str(row.get("월")): row.get("평균") for row in lp.get("monthly_rows", []) if row.get("월")}
    for row in monthly_rows:
        if row.get("공표변동률") is None:
            row["공표변동률"] = public_rate_lookup.get(str(row.get("월")))
    chart_labels = [r.get("월", "")[-5:] for r in monthly_rows]
    chart_amounts = [r.get("거래금액_억원", 0) for r in monthly_rows]
    chart_counts = [r.get("거래건수", 0) for r in monthly_rows]

    takeaways = n.get("key_takeaways") or []
    takeaways_html = "".join(f"<li>{h(x)}</li>" for x in takeaways)
    if not takeaways_html:
        takeaways_html = "<li>분석 요약 없음</li>"
    listing_narrative = "" if listings.get("template_only") else n.get("listing_interpretation")

    monthly_table = table_html(
        monthly_rows,
        [("월", "월"), ("거래건수", "거래건수"), ("거래금액_억원", "거래금액(억원)"), ("외지인거래건수", "외지인"), ("직거래건수", "직거래")],
    )
    emd_table = table_html(
        tx.get("emd_top", []) if tx.get("available") else [],
        [("명칭", "읍면동"), ("거래건수", "거래건수"), ("거래금액_억원", "거래금액(억원)"), ("금액비중", "금액비중(%)")],
    )
    use_table = table_html(
        tx.get("use_top", []) if tx.get("available") else [],
        [("명칭", "구분"), ("거래건수", "거래건수"), ("거래금액_억원", "거래금액(억원)"), ("금액비중", "금액비중(%)")],
    )
    lp_emd_table = table_html(
        lp.get("emd_top", []) if lp.get("available") else [],
        [("읍면동", "읍면동"), ("당월변동률", "당월변동률(%)"), ("누계변동률", "누계변동률(%)")],
    )
    listing_columns = [("source_file", "매물 파일"), ("estimated_listing_count", "추정 매물건수"), ("read_error", "읽기 오류")]
    listing_table = (
        blank_table_html(listing_columns)
        if listings.get("template_only")
        else table_html(listings.get("files", []) if listings.get("available") else [], listing_columns)
    )
    listing_source_label = "매물 양식(미입력)" if listings.get("template_only") else "매물 엑셀"
    listing_meta_label = "거래·지가 수치는 업로드 및 수집 엑셀에서 Python으로 산출했습니다. 매물은 미지정되어 표 양식만 표시합니다." if listings.get("template_only") else "모든 거래·지가·매물 수치는 업로드 파일 및 수집 엑셀에서 Python으로 산출했습니다."
    issue_rows = regional_issues.get("events", []) if regional_issues.get("available") else []
    issue_table = table_html(
        issue_rows,
        [("event_date", "일자"), ("title", "이슈"), ("summary", "보고서 반영 방향")],
    )
    regional_issue_section = ""
    if regional_issues.get("available"):
        regional_issue_section = f"""
<section>
  <h2>6. 지역 개발·규제 이슈</h2>
  <p>{h(n.get('regional_issue_interpretation'))}</p>
  <div class="grid">
    <div class="card"><div class="t">표기 권역</div><div class="v smallv">{h(regional_issues.get('display_area'))}</div></div>
    <div class="card"><div class="t">허가구역 효력</div><div class="v smallv">2026.07.14</div></div>
    <div class="card"><div class="t">지정기간</div><div class="v smallv">2026.07.14~2028.07.13</div></div>
  </div>
  {issue_table}
  <p class="note">지역 이슈는 개발 기대와 거래 규제 배경으로만 해석하며, 지가변동률의 직접 원인으로 단정하지 않습니다.</p>
</section>
"""

    return f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{h(meta.get('region'))} 토지시장동향 시황보고서 {h(meta.get('base_month'))}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{{--ink:#1a1f2b;--sub:#5b6472;--line:#e3e7ee;--bg:#f6f7f9;--navy:#0f2a52;--gold:#9a7b25;}}
*{{box-sizing:border-box}}
body{{font-family:'Pretendard','Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:var(--ink);line-height:1.75;margin:0;background:var(--bg);font-size:15px}}
.wrap{{max-width:1080px;margin:0 auto;background:#fff;padding-bottom:72px}}
.cover{{background:linear-gradient(135deg,#0f2a52,#1d3f73);color:#fff;padding:54px 56px 46px}}
.kicker{{letter-spacing:5px;font-size:12px;opacity:.8}}
h1{{font-size:32px;margin:14px 0 8px;line-height:1.3}}
.meta{{opacity:.86;font-size:13.5px;border-top:1px solid rgba(255,255,255,.25);padding-top:14px;margin-top:18px}}
.stat{{display:flex;gap:16px;margin-top:24px;flex-wrap:wrap}}
.stat div{{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.18);border-radius:12px;padding:14px 18px;min-width:150px}}
.stat b{{display:block;font-size:23px}}
.stat span{{font-size:12px;opacity:.82}}
section{{padding:30px 56px;border-bottom:1px solid var(--line)}}
h2{{font-size:21px;color:var(--navy);border-left:6px solid var(--gold);padding-left:12px;margin:6px 0 18px}}
h3{{font-size:16px;color:#243a5e;margin:24px 0 8px}}
.kt{{background:#fbfaf4;border:1px solid #ece3c4;border-radius:14px;padding:18px 22px;margin:10px 0}}
.kt li{{margin:7px 0}}
.grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:14px 0}}
.card{{border:1px solid var(--line);border-radius:12px;padding:16px 18px;background:#fff}}
.card .t{{font-size:12.5px;color:var(--sub)}} .card .v{{font-size:22px;font-weight:800;margin:4px 0}}
.card .smallv{{font-size:18px;line-height:1.35}}
.chartbox{{border:1px solid var(--line);border-radius:12px;padding:18px;margin:14px 0;height:300px;position:relative}}
table{{border-collapse:collapse;width:100%;margin:12px 0;font-size:13.5px}}
th,td{{border:1px solid var(--line);padding:7px 9px;text-align:center}}
th{{background:#0f2a52;color:#fff;font-size:12.5px}}
tr.hot{{background:#fff7e6;font-weight:700}}
.template-only td.blank-cell{{height:34px;background:#fff}}
.note{{font-size:12px;color:#7a8493;margin-top:8px}}
.up{{color:#d4111a;font-weight:700}} .dn{{color:#1257d4;font-weight:700}} .muted{{color:#8a93a3}}
.warn{{background:#fff7e6;border:1px solid #edd8a1;border-radius:10px;padding:13px 16px;margin:12px 0}}
@media(max-width:720px){{section,.cover{{padding:22px}}.grid{{grid-template-columns:1fr}}h1{{font-size:24px}}}}
@page{{size:A4 portrait;margin:11mm 9mm 12mm}}
@media print{{html,body{{width:210mm;background:#fff!important}}body{{font-size:9pt;line-height:1.5;-webkit-print-color-adjust:exact;print-color-adjust:exact}}.wrap{{width:100%;max-width:none;margin:0;padding:0}}.cover{{min-height:273mm;break-after:page;padding:24mm 15mm}}section{{padding:8mm 7mm;break-before:page;break-inside:auto;border-bottom:0}}h2,h3{{break-after:avoid}}table{{font-size:7pt;table-layout:fixed}}th,td{{padding:1.2mm 1mm;overflow-wrap:anywhere;word-break:keep-all}}thead{{display:table-header-group}}tr{{break-inside:avoid}}.chartbox{{height:66mm}}}}
</style>
</head>
<body><div class="wrap">
<div class="cover">
  <div class="kicker">LAND MARKET BRIEF</div>
  <h1>{h(meta.get('base_month'))} {h(meta.get('region'))} 토지시장동향 시황보고서</h1>
  <div>업로드 실거래 엑셀 · 당월 지가동향 엑셀 · {listing_source_label} · KOSIS 보조지표 기반</div>
  <div class="stat">
    <div><b>{h(pick(cur, '거래건수', default='-'))}건</b><span>당월 거래건수</span></div>
    <div><b>{fmt_num(pick(cur, '거래금액_억원'), 1, '억원')}</b><span>당월 거래금액</span></div>
    <div><b>{fmt_num(pick(lp_cur, '평균'), 3, '%')}</b><span>당월 지가변동률 평균</span></div>
    <div><b>{fmt_num(rates.get('current_rate'), 3, '%')}</b><span>대출금리</span></div>
  </div>
  <div class="meta">{listing_meta_label} 지가변동률은 조사담당자가 가격 변화·실거래·지역시장동향을 종합해 판단한 결과입니다.</div>
</div>

<section>
  <h2>[Key Takeaways] 당월 시장 핵심 요약</h2>
  <div class="kt"><ol>{takeaways_html}</ol></div>
</section>

<section>
  <h2>1. Executive Summary</h2>
  <p>{h(n.get('executive_summary'))}</p>
  <div class="grid">
    <div class="card"><div class="t">전월 대비 거래건수</div><div class="v">{fmt_signed(pick(prev_chg, '거래건수증감률'))}</div></div>
    <div class="card"><div class="t">전월 대비 거래금액</div><div class="v">{fmt_signed(pick(prev_chg, '거래금액증감률'))}</div></div>
    <div class="card"><div class="t">누계 지가변동률 평균</div><div class="v">{fmt_num(pick(lp_cum, '평균'), 3, '%')}</div></div>
  </div>
</section>

<section>
  <h2>2. 실거래 시장 분석</h2>
  <p>{h(n.get('transaction_interpretation') or '업로드 실거래 엑셀을 기준으로 월별 거래건수, 거래금액, 읍면동별 집중도를 산출했습니다.')}</p>
  <div class="chartbox"><canvas id="trendChart"></canvas></div>
  <h3>최근 13개월 거래 추이</h3>
  {monthly_table}
  <h3>읍면동별 거래금액 상위</h3>
  {emd_table}
  <h3>용도·이용상황별 거래금액 상위</h3>
  {use_table}
</section>

<section>
  <h2>3. 지가변동률 분석</h2>
  <p>{h(n.get('land_price_interpretation') or '당월 지가동향 엑셀에서 시군구 평균, 용도지역별, 이용상황별, 읍면동별 변동률을 추출했습니다.')}</p>
  <div class="grid">
    <div class="card"><div class="t">평균</div><div class="v">{fmt_num(pick(lp_cur, '평균'), 3, '%')}</div></div>
    <div class="card"><div class="t">주거</div><div class="v">{fmt_num(pick(lp_cur, '주거'), 3, '%')}</div></div>
    <div class="card"><div class="t">상업</div><div class="v">{fmt_num(pick(lp_cur, '상업'), 3, '%')}</div></div>
    <div class="card"><div class="t">공업</div><div class="v">{fmt_num(pick(lp_cur, '공업'), 3, '%')}</div></div>
    <div class="card"><div class="t">계획관리</div><div class="v">{fmt_num(pick(lp_cur, '계획관리'), 3, '%')}</div></div>
    <div class="card"><div class="t">주거용</div><div class="v">{fmt_num(pick(lp_cur, '주거용'), 3, '%')}</div></div>
  </div>
  <h3>읍면동별 당월 지가변동률 상위</h3>
  {lp_emd_table}
  <p class="note">지가변동률은 조사담당자가 가격 변화·실거래·지역시장동향을 종합해 판단한 결과입니다. 실거래가 적은 지역은 개별 거래를 일반화하지 않습니다.</p>
</section>

<section>
  <h2>4. 매물 및 호가 흐름</h2>
  {f"<p>{h(listing_narrative)}</p>" if listing_narrative else ""}
  {listing_table}
  {"<p class='note'>매물 엑셀 미지정: 자료값 없이 표 양식만 표시합니다.</p>" if listings.get("template_only") else "<div class='warn'>전월·당월 매물 파일 2개가 모두 있어야 매물 증감과 호가 조정 판단을 정식 산출합니다.</div>" if not listings.get("has_prev_and_current") else ""}
</section>

<section>
  <h2>5. 인구·세대·이동 및 대출금리</h2>
  <p>{h(n.get('population_household_migration_interpretation') or n.get('population_interpretation'))}</p>
  <div class="grid">
    <div class="card"><div class="t">주민등록인구</div><div class="v">{fmt_num(pop.get('current'), 0, '명')}</div><div>전월 대비 {fmt_num(pop.get('delta'), 0, '명')}</div></div>
    <div class="card"><div class="t">주민등록세대수</div><div class="v">{fmt_num(hh.get('current'), 0, '세대')}</div><div>전월 대비 {fmt_num(hh.get('delta'), 0, '세대')}</div></div>
    <div class="card"><div class="t">순이동</div><div class="v">{fmt_num(mig.get('current'), 0, '명')}</div><div>전월 대비 {fmt_num(mig.get('delta'), 0, '명')}</div></div>
  </div>
  <p>{h(n.get('rate_interpretation'))}</p>
</section>

{regional_issue_section}

<section>
  <h2>{'7' if regional_issues.get('available') else '6'}. 결론 및 모니터링 포인트</h2>
  <p>{h(n.get('conclusion') or '거래 집중 지역, 지가변동률 상위 지역, 매물 증감, 대출금리와 인구 흐름을 다음 달에도 함께 점검할 필요가 있습니다.')}</p>
  <div class="warn">
    <b>해석 유의사항</b><br>
    실거래 자료는 신고 기준 거래 흐름이며, 지가변동률은 조사담당자가 가격 변화·실거래·지역시장동향을 종합해 판단한 결과입니다.
  </div>
</section>
</div>
<script>
new Chart(document.getElementById('trendChart'), {{
  type: 'bar',
  data: {{
    labels: {json.dumps(chart_labels, ensure_ascii=False)},
    datasets: [
      {{type:'bar', label:'거래금액(억원)', data:{json.dumps(chart_amounts, ensure_ascii=False)}, backgroundColor:'rgba(15,42,82,.75)', yAxisID:'y'}},
      {{type:'line', label:'거래건수', data:{json.dumps(chart_counts, ensure_ascii=False)}, borderColor:'#9a7b25', backgroundColor:'#9a7b25', yAxisID:'y1', tension:.3}}
    ]
  }},
  options: {{
    responsive:true, maintainAspectRatio:false,
    plugins: {{legend: {{position:'bottom'}}}},
    scales: {{
      y: {{position:'left', title:{{display:true,text:'금액(억원)'}}}},
      y1: {{position:'right', grid:{{drawOnChartArea:false}}, title:{{display:true,text:'건수'}}}}
    }}
  }}
}});
</script>
</body></html>"""


def _change_badge(value: Any) -> str:
    if value is None:
        return '<span class="g">–</span>'
    try:
        number = float(value)
    except Exception:
        return h(value)
    if number > 0:
        return f'<span class="up">▲{number:.1f}%</span>'
    if number < 0:
        return f'<span class="dn">▼{abs(number):.1f}%</span>'
    return '<span class="g">―0.0%</span>'


def _reference_table(
    rows: list[dict[str, Any]],
    columns: list[tuple[str, str, Any]],
    hot_rows: int = 0,
    table_class: str = "",
    blank_rows: int = 1,
) -> str:
    head = "".join(f"<th>{h(label)}</th>" for _, label, _ in columns)
    body: list[str] = []
    if not rows:
        for _ in range(max(1, blank_rows)):
            body.append("<tr>" + "".join('<td class="blank-cell">&nbsp;</td>' for _ in columns) + "</tr>")
    else:
        for i, row in enumerate(rows, 1):
            cells = []
            for key, _, formatter in columns:
                value = row.get(key)
                if callable(formatter):
                    cells.append(f"<td>{formatter(value)}</td>")
                else:
                    cells.append(f"<td>{h(value if value is not None else '-')}</td>")
            cls = ' class="hot"' if hot_rows and i <= hot_rows else ""
            body.append(f"<tr{cls}>{''.join(cells)}</tr>")
    cls_attr = f' class="{h(table_class)}"' if table_class else ""
    return f"<table{cls_attr}><thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table>"


def _fmt_int_html(value: Any) -> str:
    return h(fmt_num(value, 0))


def _emd_label_html(name: Any, region: str) -> str:
    if "광산구" not in str(region or ""):
        return h(name)
    tokens = [str(token).strip() for token in re.split(r"[,·/]", str(name or "")) if str(token).strip()]
    matched: dict[str, list[str]] = {level: [] for level in SEMICONDUCTOR_LEVEL_ORDER}
    labels = SEMICONDUCTOR_LEVEL_LABELS
    for token in tokens:
        normalized = norm_text(token)
        for dong, (level, _) in SEMICONDUCTOR_INFLUENCE_DONGS.items():
            if normalized == norm_text(dong):
                matched.setdefault(level, []).append(dong)
    # 단일 읍면동 셀에서는 배지에 동 이름을 되풀이하지 않는다(예: '송정동 [반도체 직접 영향권]').
    single = len(tokens) == 1 and sum(len(dongs) for dongs in matched.values()) == 1
    badges = "".join(
        f'<span class="influence-badge {h(level)}">{h(labels.get(level, level))}'
        + ("" if single else f': {h("·".join(dongs))}')
        + "</span>"
        for level, dongs in matched.items() if dongs
    )
    return f'<span class="emd-name">{h(name)}</span>{badges}' if badges else h(name)


def _aggregate_basic_area_monthly_rows(
    emd_rows: list[dict[str, Any]],
    basic_area_rows: list[dict[str, Any]],
    months: list[str],
) -> list[dict[str, Any]]:
    """지가동향 원표의 기초단위 권역으로 읍면동 실거래 월별 값을 합산한다."""
    output: list[dict[str, Any]] = []
    for area_row in basic_area_rows:
        area = str(area_row.get("읍면동") or "").strip()
        tokens = [norm_text(token) for token in re.split(r"[,·/]", area) if norm_text(token)]
        matched: list[dict[str, Any]] = []
        for row in emd_rows:
            emd_name = norm_text(row.get("읍면동"))
            if any(emd_name == token or emd_name.startswith(token) or token.startswith(emd_name) for token in tokens):
                matched.append(row)
        monthly = []
        for month in months:
            month_count = 0
            month_amount = 0.0
            for row in matched:
                by_month = {str(item.get("월")): item for item in row.get("월별", [])}
                item = by_month.get(month, {})
                month_count += int(as_number(item.get("거래건수")) or 0)
                month_amount += float(as_number(item.get("거래금액_억원")) or 0)
            monthly.append({"월": month, "거래건수": month_count, "거래금액_억원": round(month_amount, 1)})
        output.append({
            "읍면동": area,
            "_previous_month": next((row.get("_previous_month") for row in matched if row.get("_previous_month")), None),
            "_previous_count": sum(int(as_number(row.get("_previous_count")) or 0) for row in matched),
            "_previous_amount_eok": round(sum(float(as_number(row.get("_previous_amount_eok")) or 0) for row in matched), 1),
            "월별": monthly,
            "합계건수": sum(item["거래건수"] for item in monthly),
            "합계금액_억원": round(sum(item["거래금액_억원"] for item in monthly), 1),
            "당월변동률": area_row.get("당월변동률"),
        })
    return output


def _emd_timeseries_table(
    rows: list[dict[str, Any]],
    months: list[str],
    region: str,
    metric: str,
    first_column_label: str = "읍·면·동",
    show_influence: bool = True,
    include_current_rate: bool = False,
) -> str:
    is_count = metric == "count"
    total_key = "합계건수" if is_count else "합계금액_억원"
    total_label = "합계(건)" if is_count else "합계(억원)"
    rate_header = '<th>당월 변동률(%)</th>' if include_current_rate else ""
    headers = f'<th class="emd-col">{h(first_column_label)}</th>' + rate_header + "".join(f"<th>{h(month)}</th>" for month in months) + f"<th>{total_label}</th>"
    body: list[str] = []
    if not rows:
        body.append(f'<tr><td colspan="{len(months) + 2 + int(include_current_rate)}" class="blank-cell">자료 없음</td></tr>')
    else:
        for row in rows:
            by_month = {str(item.get("월")): item for item in row.get("월별", [])}
            values = []
            previous = as_number(row.get("_previous_count" if is_count else "_previous_amount_eok"))
            for month_index, month in enumerate(months):
                item = by_month.get(month, {})
                value = item.get("거래건수", 0) if is_count else item.get("거래금액_억원", 0)
                number = float(as_number(value) or 0)
                if previous is None:
                    heat_class = "heat-na"
                    marker = '<span class="heat-delta">비교월 없음</span>'
                else:
                    delta = number - float(previous)
                    if delta > 0:
                        heat_class = "heat-up"
                        marker = f'<span class="heat-delta">▲{fmt_num(delta, 0 if is_count else 1)}</span>'
                    elif delta < 0:
                        heat_class = "heat-down"
                        marker = f'<span class="heat-delta">▼{fmt_num(abs(delta), 0 if is_count else 1)}</span>'
                    else:
                        heat_class = "heat-flat"
                        marker = '<span class="heat-delta">―0</span>'
                if month_index == len(months) - 1:
                    heat_class += " heat-latest"
                values.append(
                    f'<td class="{heat_class}"><span class="heat-value">'
                    f'{fmt_num(number, 0 if is_count else 1)}</span>{marker}</td>'
                )
                previous = number
            total = row.get(total_key, 0)
            body.append(
                '<tr><td class="emd-col">'
                + (_emd_label_html(row.get("읍면동"), region) if show_influence else h(row.get("읍면동")))
                + "</td>"
                + (f'<td>{fmt_num(row.get("당월변동률"), 3, "%")}</td>' if include_current_rate else "")
                + "".join(values)
                + f'<td class="total-cell">{fmt_num(total, 0 if is_count else 1)}</td></tr>'
            )
    legend = (
        '<div class="heatmap-legend">'
        '<span class="legend-up">▲ 전월 대비 증가</span>'
        '<span class="legend-down">▼ 전월 대비 감소</span>'
        '<span class="legend-flat">― 보합</span>'
        '<span class="legend-latest">□ 최신월</span>'
        '</div>'
    )
    return legend + f'<div class="wide-table-wrap"><table class="wide-table heatmap-table"><thead><tr>{headers}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'


def _semiconductor_policy_card(region: str) -> str:
    if "광산구" not in str(region or ""):
        return ""
    direct = "·".join(semiconductor_dongs("direct"))
    linked = "·".join(semiconductor_dongs("linked"))
    return f"""
<div class="policy-card">
  <div class="policy-title">정책·개발 이슈 | 호남 반도체 첨단국가산단 및 토지거래허가구역</div>
  <p><b>확인된 사실:</b> 국토교통부 공고 제2026-940호에 따라 광주 군공항의 호남 반도체 첨단국가산단 사업 예정지와 인근 지역이 토지거래허가구역으로 지정되어 2026년 7월 14일부터 효력이 발생함. 지정기간은 2026-07-14 ~ 2028-07-13임.</p>
  <div class="influence-groups"><span class="influence-badge direct">반도체 직접 영향권: {h(direct)}</span><span class="influence-badge linked">연계 영향권: {h(linked)}</span></div>
  <p><b>분류 기준:</b> 공식 사업구역 명칭이 아니라 군공항 예정지와의 인접성, 송정 생활권 및 교통 연계를 기준으로 한 본 보고서용 분석 분류임. 2026년 8월 조사부터 송정동·도산동·신촌동·소촌동·우산동·황룡동을 반도체 직접 영향권으로 통합 표기함.</p>
  <p class="policy-caution"><b>해석 유의:</b> 토지거래허가구역은 투기 방지를 위한 규제구역이며 개별 읍면동의 직접 수혜나 가격 상승을 의미하지 않음. 규제효과와 산업단지 조성 기대효과를 구분하여 거래량·거래금액의 후속 변화를 확인해야 함.</p>
  <p class="policy-source">참고: 국토교통부 공고 제2026-940호</p>
</div>"""


def _permit_split_table(permit: dict[str, Any], region: str) -> str:
    """토지거래허가구역 지정 전·후 읍·면·동 거래건수 표(계약일 기준)."""
    if not permit.get("available"):
        reason = permit.get("reason") or "자료 없음"
        return f'<p class="note">{h(reason)}</p>'
    periods = permit.get("periods", {})
    pre, post = periods.get("pre", {}), periods.get("post", {})
    split_mode = permit.get("mode") == "split"
    headers = (
        '<tr><th rowspan="2" class="emd-col">읍·면·동</th>'
        + (f'<th colspan="2">지정 전 {h(pre.get("label"))} ({h(pre.get("days"))}일)</th>' if split_mode else "")
        + f'<th colspan="2">지정 후 {h(post.get("label"))} ({h(post.get("days"))}일)</th>'
        + "</tr><tr>"
        + ("<th>건수</th><th>금액(억)</th>" if split_mode else "")
        + "<th>건수</th><th>금액(억)</th></tr>"
    )
    body: list[str] = []
    for row in permit.get("emd_rows", []):
        body.append(
            '<tr><td class="emd-col">' + _emd_label_html(row.get("읍면동"), region) + "</td>"
            + ((
                f'<td>{fmt_num(row.get("지정전_거래건수"), 0)}</td>'
                f'<td>{fmt_num(row.get("지정전_거래금액_억원"), 1)}</td>'
            ) if split_mode else "")
            + f'<td>{fmt_num(row.get("지정후_거래건수"), 0)}</td>'
            + f'<td>{fmt_num(row.get("지정후_거래금액_억원"), 1)}</td>'
            + "</tr>"
        )
    if not body:
        span = 5 if split_mode else 3
        body.append(f'<tr><td colspan="{span}" class="blank-cell">해당 기간 거래 없음</td></tr>')
    totals = permit.get("totals", {})
    total_pre, total_post = totals.get("pre", {}), totals.get("post", {})
    foot = (
        '<tr class="total-row"><td class="emd-col">합계</td>'
        + ((
            f'<td>{fmt_num(total_pre.get("거래건수"), 0)}</td>'
            f'<td>{fmt_num(total_pre.get("거래금액_억원"), 1)}</td>'
        ) if split_mode else "")
        + f'<td>{fmt_num(total_post.get("거래건수"), 0)}</td>'
        + f'<td>{fmt_num(total_post.get("거래금액_억원"), 1)}</td>'
        + "</tr>"
    )
    return (
        '<div class="table-wrap permit-split-wrap"><table class="rk permit-split-table"><thead>'
        + headers + f'</thead><tbody>{"".join(body)}</tbody><tfoot>{foot}</tfoot></table></div>'
    )


def _period_trend_table(trend: dict[str, Any]) -> str:
    """전체 거래 추이 표: 기준월 이전은 월단위 행, 기준월은 주단위 행."""
    rows = trend.get("rows", []) if trend.get("available") else []
    if not rows:
        return '<p class="note">기간별 집계 자료 없음</p>'
    body = []
    for row in rows:
        kind = str(row.get("구분") or "월")
        is_week, is_total = kind == "주", kind == "합계"
        row_cls = "row-total" if is_total else "row-week" if is_week else "row-month"
        tag_cls = "tag-total" if is_total else "tag-week" if is_week else "tag-month"
        delta = row.get("전기대비건수증감")
        delta_cls = "delta-up" if (delta or 0) > 0 else "delta-down" if (delta or 0) < 0 else "delta-flat"
        delta_text = "-" if delta is None else f'{"+" if delta > 0 else ""}{fmt_num(delta, 0)}'
        if is_total and row.get("비교기준"):
            delta_text += f'<span class="delta-note">{h(row.get("비교기준"))}</span>'
        body.append(
            f'<tr class="{row_cls}">'
            f'<td><span class="period-tag {tag_cls}">{"계" if is_total else "주" if is_week else "월"}</span> {h(row.get("기간"))}</td>'
            f'<td>{h(row.get("허가구역구분"))}</td>'
            f'<td>{fmt_num(row.get("거래건수"), 0)}</td>'
            f'<td class="{delta_cls}">{delta_text}</td>'
            f'<td>{fmt_num(row.get("거래금액_억원"), 1)}</td>'
            f'<td>{fmt_num(row.get("외지인거래건수"), 0)}</td>'
            f'<td>{fmt_num(row.get("외지인거래금액_억원"), 1)}</td>'
            f'<td>{fmt_num(row.get("직거래건수"), 0)}</td></tr>'
        )
    return (
        '<div class="table-wrap"><table class="rk period-trend-table"><thead><tr>'
        '<th>기간</th><th>허가구역</th><th>거래건수</th><th>전기 대비</th>'
        '<th>거래금액(억)</th><th>외지인 건수</th><th>외지인 금액(억)</th><th>직거래 건수</th></tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>'
    )


def _weekly_trend_table(weekly: dict[str, Any]) -> str:
    """최근 주간(토~금) 거래 추이 표."""
    rows = weekly.get("weekly_rows", []) if weekly.get("available") else []
    if not rows:
        return '<p class="note">주간 집계 자료 없음</p>'
    body = []
    for row in rows:
        delta = row.get("전주대비건수증감")
        delta_cls = "delta-up" if (delta or 0) > 0 else "delta-down" if (delta or 0) < 0 else "delta-flat"
        delta_text = "-" if delta is None else f'{"+" if delta > 0 else ""}{fmt_num(delta, 0)}'
        body.append(
            f'<tr><td>{h(row.get("주차"))}</td><td>{h(row.get("종료일"))}</td>'
            f'<td>{h(row.get("허가구역구분"))}</td>'
            f'<td>{fmt_num(row.get("거래건수"), 0)}</td>'
            f'<td class="{delta_cls}">{delta_text}</td>'
            f'<td>{fmt_num(row.get("거래금액_억원"), 1)}</td>'
            f'<td>{fmt_num(row.get("외지인거래건수"), 0)}</td>'
            f'<td>{fmt_num(row.get("직거래건수"), 0)}</td></tr>'
        )
    return (
        '<div class="table-wrap"><table class="rk weekly-trend-table"><thead><tr>'
        '<th>주차(토~금)</th><th>종료일(금)</th><th>허가구역</th><th>거래건수</th><th>전주 대비</th>'
        '<th>거래금액(억)</th><th>외지인 건수</th><th>직거래 건수</th></tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>'
    )


def _influence_group_weekly_table(weekly: dict[str, Any]) -> str:
    """직접·간접 영향권 구분별 주간 거래건수 표."""
    rows = weekly.get("influence_group_rows", []) if weekly.get("available") else []
    labels = weekly.get("week_labels", []) or []
    if not rows or not labels:
        return '<p class="note">영향권 구분별 주간 집계 자료 없음</p>'
    headers = (
        '<th class="emd-col">구분</th>'
        + "".join(f"<th>{h(label)}</th>" for label in labels)
        + "<th>합계(건)</th><th>합계(억)</th>"
    )
    body = []
    for row in rows:
        by_week = {str(item.get("주차")): item for item in row.get("주별", [])}
        cells = "".join(
            f'<td>{fmt_num((by_week.get(label) or {}).get("거래건수", 0), 0)}</td>' for label in labels
        )
        body.append(
            f'<tr><td class="emd-col"><b>{h(row.get("구분"))}</b><br><span class="group-scope">{h(row.get("대상"))}</span></td>'
            + cells
            + f'<td class="total-cell">{fmt_num(row.get("합계건수"), 0)}</td>'
            + f'<td class="total-cell">{fmt_num(row.get("합계금액_억원"), 1)}</td></tr>'
        )
    return (
        f'<div class="wide-table-wrap"><table class="wide-table weekly-emd-table"><thead><tr>{headers}</tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>'
    )


def _emd_month_week_table(data: dict[str, Any], region: str) -> str:
    """읍·면·동별 거래 표: 기준월 이전은 월단위, 기준월은 주단위(건수·금액 병기)."""
    month_labels = data.get("month_labels", []) or []
    week_labels = data.get("week_labels", []) or []
    base_week_labels = data.get("base_month_week_labels", []) or week_labels
    ongoing_week_labels = data.get("ongoing_week_labels", []) or []
    ongoing_month = data.get("ongoing_month")
    rows = list(data.get("rows", []) or [])
    if not month_labels and not week_labels:
        return '<p class="note">읍·면·동 집계 자료 없음</p>'
    direct_dongs = semiconductor_dongs("direct") if "광산구" in str(region or "") else []
    direct_names = {norm_text(dong) for dong in direct_dongs}
    if direct_dongs:
        existing = {norm_text(row.get("읍면동")): row for row in rows}
        ordered = []
        for dong in direct_dongs:
            ordered.append(existing.get(norm_text(dong)) or {
                "읍면동": dong,
                "월별": [{"기간": label, "거래건수": 0, "거래금액_억원": 0.0} for label in month_labels],
                "주별": [{"기간": label, "거래건수": 0, "거래금액_억원": 0.0} for label in week_labels],
                "당월건수": 0, "당월금액_억원": 0.0, "진행월건수": 0, "진행월금액_억원": 0.0,
                "합계건수": 0, "합계금액_억원": 0.0,
            })
        ordered += [row for row in rows if norm_text(row.get("읍면동")) not in direct_names and row.get("합계건수")]
        rows = ordered

    headers = (
        '<tr><th rowspan="2" class="emd-col">읍·면·동</th>'
        + (f'<th colspan="{len(month_labels)}" class="group-month">기준월 이전 · 월단위</th>' if month_labels else "")
        + (f'<th colspan="{len(base_week_labels)}" class="group-week">기준월 · 주단위(금요일 종료)</th>' if base_week_labels else "")
        + (f'<th colspan="{len(ongoing_week_labels)}" class="group-ongoing">{h(ongoing_month)} 진행 · 주단위</th>'
           if ongoing_week_labels else "")
        + '<th rowspan="2">당월 합계</th>'
        + ('<th rowspan="2">진행월 누적</th>' if ongoing_week_labels else "")
        + '<th rowspan="2">기간 합계</th></tr><tr>'
        + "".join(f'<th class="group-month">{h(label)}</th>' for label in month_labels)
        + "".join(f'<th class="group-week">{h(label)}</th>' for label in base_week_labels)
        + "".join(f'<th class="group-ongoing">{h(label)}</th>' for label in ongoing_week_labels)
        + "</tr>"
    )

    def cell(item: dict[str, Any] | None, kind: str = "month") -> str:
        count = int(as_number((item or {}).get("거래건수")) or 0)
        amount = float(as_number((item or {}).get("거래금액_억원")) or 0)
        cls = {"month": "cell-month", "week": "cell-week", "ongoing": "cell-ongoing"}[kind]
        if not count and not amount:
            return f'<td class="{cls} cell-zero"><span class="cell-count">0</span><span class="cell-amount">0.0</span></td>'
        return (
            f'<td class="{cls}"><span class="cell-count">{fmt_num(count, 0)}</span>'
            f'<span class="cell-amount">{fmt_num(amount, 1)}</span></td>'
        )

    body: list[str] = []
    for row in rows:
        by_month = {str(item.get("기간")): item for item in row.get("월별", [])}
        by_week = {str(item.get("기간")): item for item in row.get("주별", [])}
        body.append(
            '<tr><td class="emd-col">' + _emd_label_html(row.get("읍면동"), region) + "</td>"
            + "".join(cell(by_month.get(label), "month") for label in month_labels)
            + "".join(cell(by_week.get(label), "week") for label in base_week_labels)
            + "".join(cell(by_week.get(label), "ongoing") for label in ongoing_week_labels)
            + f'<td class="total-cell"><span class="cell-count">{fmt_num(row.get("당월건수"), 0)}</span>'
            f'<span class="cell-amount">{fmt_num(row.get("당월금액_억원"), 1)}</span></td>'
            + (f'<td class="total-cell ongoing-total"><span class="cell-count">{fmt_num(row.get("진행월건수"), 0)}</span>'
               f'<span class="cell-amount">{fmt_num(row.get("진행월금액_억원"), 1)}</span></td>' if ongoing_week_labels else "")
            + f'<td class="total-cell"><span class="cell-count">{fmt_num(row.get("합계건수"), 0)}</span>'
            f'<span class="cell-amount">{fmt_num(row.get("합계금액_억원"), 1)}</span></td></tr>'
        )
    if not body:
        body.append(
            f'<tr><td colspan="{len(month_labels) + len(week_labels) + 4}" class="blank-cell">해당 기간 거래 없음</td></tr>'
        )
    legend = '<div class="heatmap-legend"><span>각 칸: 위 = 거래건수(건), 아래 = 거래금액(억원)</span></div>'
    return (
        legend
        + f'<div class="wide-table-wrap"><table class="wide-table month-week-table"><thead>{headers}</thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>'
    )


def _weekly_emd_table(weekly: dict[str, Any], region: str, only_direct: bool = True) -> str:
    """읍·면·동별 주간 거래건수 표(직접 영향권 우선)."""
    rows = weekly.get("weekly_emd_rows", []) if weekly.get("available") else []
    labels = weekly.get("week_labels", [])
    if only_direct:
        direct_names = {norm_text(dong) for dong in semiconductor_dongs("direct")}
        existing = {norm_text(row.get("읍면동")): row for row in rows}
        rows = []
        for dong in semiconductor_dongs("direct"):
            row = existing.get(norm_text(dong))
            rows.append(row or {
                "읍면동": dong,
                "주별": [{"주차": label, "거래건수": 0, "거래금액_억원": 0.0} for label in labels],
                "합계건수": 0,
                "합계금액_억원": 0.0,
            })
        rows += [row for row in (weekly.get("weekly_emd_rows", []) or [])
                 if norm_text(row.get("읍면동")) not in direct_names and row.get("합계건수")]
    if not rows or not labels:
        return '<p class="note">주간 읍·면·동 집계 자료 없음</p>'
    headers = '<th class="emd-col">읍·면·동</th>' + "".join(f"<th>{h(label)}</th>" for label in labels) + "<th>합계(건)</th><th>합계(억)</th>"
    body = []
    for row in rows:
        by_week = {str(item.get("주차")): item for item in row.get("주별", [])}
        cells = "".join(
            f'<td>{fmt_num((by_week.get(label) or {}).get("거래건수", 0), 0)}</td>' for label in labels
        )
        body.append(
            '<tr><td class="emd-col">' + _emd_label_html(row.get("읍면동"), region) + "</td>" + cells
            + f'<td class="total-cell">{fmt_num(row.get("합계건수"), 0)}</td>'
            + f'<td class="total-cell">{fmt_num(row.get("합계금액_억원"), 1)}</td></tr>'
        )
    return (
        f'<div class="wide-table-wrap"><table class="wide-table weekly-emd-table"><thead><tr>{headers}</tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>'
    )


def _field_survey_form_html(result: dict[str, Any], region: str, narratives: dict[str, Any] | None = None) -> str:
    """현장점검 보고서(주간) 작성 양식.

    자동 집계 근거 + LLM(또는 Python 초안) 서술을 함께 싣고, 조사담당자 보완 기입란을 남긴다.
    """
    survey = (narratives or {}).get("weekly_field_survey") or {}
    if not isinstance(survey, dict):
        survey = {}
    survey_groups = {
        norm_text(item.get("group")): str(item.get("narrative") or "")
        for item in (survey.get("group_narratives") or []) if isinstance(item, dict)
    }
    tx = result.get("transactions", {}) or {}
    weekly = tx.get("weekly", {}) or {}
    permit = tx.get("permit_split", {}) or {}
    report_week = weekly.get("report_week", {}) or {}
    week_rows = weekly.get("weekly_rows", []) or []
    latest = week_rows[-1] if week_rows else {}
    previous = week_rows[-2] if len(week_rows) > 1 else {}
    emd_rows = {norm_text(row.get("읍면동")): row for row in (weekly.get("weekly_emd_rows", []) or [])}
    permit_rows = {norm_text(row.get("읍면동")): row for row in (permit.get("emd_rows", []) or [])}

    week_count = len(weekly.get("week_labels", []) or []) or len(week_rows)

    group_weekly = {str(row.get("구분")): row for row in (weekly.get("influence_group_rows", []) or [])}
    group_permit = {str(row.get("구분")): row for row in (permit.get("influence_group_rows", []) or [])}
    group_month = {str(row.get("구분")): row for row in (tx.get("influence_group_month_rows", []) or [])}
    group_trades = weekly.get("influence_group_recent_trades", {}) or {}
    direct_dongs = semiconductor_dongs("direct")
    direct_names = {norm_text(dong) for dong in direct_dongs}
    base_month_text = str((result.get("meta", {}) or {}).get("base_month") or "")
    try:
        month_label = f"{pd.Period(base_month_text, freq='M').year}년 {pd.Period(base_month_text, freq='M').month}월"
    except Exception:
        month_label = base_month_text or "당월"

    def group_fact(label: str) -> str:
        row = group_weekly.get(label, {})
        month_row = group_month.get(label, {})
        pieces = [
            f'기준주 거래 {fmt_num(row.get("기준주거래건수", 0), 0)}건'
            f'({fmt_num(row.get("기준주거래금액_억원", 0), 1)}억)',
            f'전주 {fmt_num(row.get("전주거래건수", 0), 0)}건',
            f'당월({h(month_label)}) 거래 {fmt_num(month_row.get("거래건수", 0), 0)}건·'
            f'{fmt_num(month_row.get("거래금액_억원", 0), 1)}억'
            f'(외지인 {fmt_num(month_row.get("외지인거래건수", 0), 0)}건)',
        ]
        permit_row = group_permit.get(label)
        if permit_row and permit.get("mode") == "split":
            pieces.append(
                f'허가구역 지정 전 {fmt_num(permit_row.get("지정전_거래건수"), 0)}건'
                f'(일평균 {fmt_num(permit_row.get("지정전_일평균건수"), 2)}) → '
                f'지정 후 {fmt_num(permit_row.get("지정후_거래건수"), 0)}건'
                f'(일평균 {fmt_num(permit_row.get("지정후_일평균건수"), 2)})'
            )
        return " · ".join(pieces)

    def dong_breakdown(names: list[str]) -> str:
        if not names:
            return ""
        items = "".join(
            f'<span class="dong-chip">{h(name)} '
            f'{fmt_num((emd_rows.get(norm_text(name)) or {}).get("합계건수", 0), 0)}건'
            f'<i>기준주 {fmt_num((((emd_rows.get(norm_text(name)) or {}).get("주별") or [{}])[-1]).get("거래건수", 0), 0)}건</i>'
            "</span>"
            for name in names
        )
        return f'<div class="form-dongs">{items}</div>'

    def group_trade_list(label: str) -> str:
        trades = group_trades.get(label, []) or []
        if not trades:
            return ""
        items = "".join(
            f'<li>{h(row.get("계약일"))} {h(row.get("소재지"))} · {h(row.get("물건구분"))} · {h(row.get("용도지역"))}'
            + (f' · 지목 {h(row.get("지목"))}' if row.get("지목") else "")
            + f' · {fmt_num(row.get("면적_㎡"), 1)}㎡ · {fmt_num(row.get("거래금액_억원"), 2)}억'
            + (f' · {_fmt_won_html(row.get("단가_원㎡"))}원/㎡' if row.get("단가_원㎡") is not None else "")
            + f' · 거래주체 {h(row.get("거래주체") or "미상")}'
            + f' · 매수자 {h(row.get("매수자시군구") or row.get("매수자지역") or "미상")}'
            + (" (외지인)" if row.get("외지인") == "Y" else " (관내)")
            + f' · {h(row.get("허가구역구분"))}'
            + "</li>"
            for row in trades
        )
        return f'<div class="form-trades"><b>최근 계약 건(계약일 기준)</b><ul>{items}</ul></div>'

    listings_available = bool((result.get("listings", {}) or {}).get("available"))

    def demand_fact(label: str) -> str:
        row = group_month.get(label, {})
        if not row or not row.get("거래건수"):
            return ""
        top_zone = (row.get("용도지역구성") or [{}])[0]
        return (
            f'당월 매수 구성: 외지인 {fmt_num(row.get("외지인거래건수"), 0)}건'
            f'({fmt_num(row.get("외지인건수비중"), 1)}%·금액비중 {fmt_num(row.get("외지인금액비중"), 1)}%) · '
            f'법인 {fmt_num(row.get("법인매수건수"), 0)}건({fmt_num(row.get("법인매수비중"), 1)}%) · '
            f'토지만 거래 {fmt_num(row.get("토지만거래건수"), 0)}건({fmt_num(row.get("토지만거래비중"), 1)}%) · '
            f'녹지·관리·개발제한 {fmt_num(row.get("녹지·관리·개발제한거래건수"), 0)}건'
            f'({fmt_num(row.get("녹지·관리·개발제한비중"), 1)}%) · '
            f'최다 용도지역 {top_zone.get("용도지역", "-")} {fmt_num(top_zone.get("건수"), 0)}건'
        )

    monitoring = result.get("field_monitoring", {}) or {}
    monitoring_dongs = monitoring.get("dongs", {}) if monitoring.get("available") else {}
    monitoring_groups = monitoring.get("groups", {}) if monitoring.get("available") else {}

    def monitoring_block(label: str, names: list[str]) -> str:
        if not monitoring.get("available"):
            return ""
        items: list[str] = []
        group_note = monitoring_groups.get(label)
        if group_note:
            items.append(f"<li><b>{h(label)}</b> {h(group_note)}</li>")
        for name in names:
            note = monitoring_dongs.get(name)
            if note:
                items.append(f"<li><b>{h(name)}</b> {h(note)}</li>")
        if not items:
            return ""
        return (
            '<div class="monitor-note"><div class="monitor-title">중개사무소 모니터링 청취 요지 '
            f'<span class="monitor-src">{h(monitoring.get("source"))}'
            + (f' · 청취기준일 {h(monitoring.get("as_of"))}' if monitoring.get("as_of") else "")
            + f'</span></div><ul>{"".join(items)}</ul></div>'
        )

    def price_check_block(label: str) -> str:
        source_note = (
            "매물 엑셀이 지정되어 있어 호가 자료를 함께 확인함"
            if listings_available else
            "매물 엑셀 미지정으로 호가·매물 수치는 자동 산출되지 않음(중개사무소 청취·현장 확인 항목)"
        )
        return (
            '<div class="price-check"><div class="price-check-title">매도 호가·매물 동향 확인 '
            f'<span class="price-check-note">{h(source_note)}</span></div>'
            '<table class="rk price-check-table"><thead><tr>'
            '<th>매도 호가 방향</th><th>호가 조정 사례(단가·필지)</th><th>매물 회수·신규 등록</th>'
            '<th>매수 문의(반도체 관련 포함)</th><th>중개업소 청취 요지</th></tr></thead>'
            '<tbody><tr><td class="form-blank-cell">상승 / 보합 / 하락</td><td class="form-blank-cell"></td>'
            '<td class="form-blank-cell"></td><td class="form-blank-cell"></td>'
            '<td class="form-blank-cell"></td></tr></tbody></table></div>'
        )

    indirect_names = [
        str(row.get("읍면동")) for row in (weekly.get("weekly_emd_rows", []) or [])
        if norm_text(row.get("읍면동")) not in direct_names and as_number(row.get("합계건수"))
    ][:12]
    group_blocks = "".join(
        f'<div class="form-item"><div class="form-bullet">○ ({h(title)})</div>'
        f'<div class="form-fact">계약일 기준 집계: {h(group_fact(label))}</div>'
        + (f'<div class="form-fact demand-fact">{h(demand_fact(label))}</div>' if demand_fact(label) else "")
        + dong_breakdown(names)
        + monitoring_block(label, names)
        + (f'<div class="form-draft">{h(survey_groups.get(norm_text(label)))}</div>' if survey_groups.get(norm_text(label)) else "")
        + group_trade_list(label)
        + price_check_block(label)
        + '<div class="form-blank tall"></div></div>'
        for label, title, names in [
            ("직접 영향권", f"직접 영향권: {'·'.join(direct_dongs)}", direct_dongs),
            ("간접 영향권", f"간접 영향권: {region.split()[-1] if region.split() else region} 직접 영향권 외 읍·면·동", indirect_names),
        ]
    )
    month_current = tx.get("current", {}) or {}
    overall_fact = (
        f'기준주({h(report_week.get("주차") or "-")}) 전체 거래 {fmt_num(latest.get("거래건수"), 0)}건 / '
        f'{fmt_num(latest.get("거래금액_억원"), 1)}억, 전주 {fmt_num(previous.get("거래건수"), 0)}건, '
        f'당월({h(month_label)}) 거래 {fmt_num(month_current.get("거래건수"), 0)}건·'
        f'{fmt_num(month_current.get("거래금액_억원"), 1)}억'
        f'(외지인 {fmt_num(month_current.get("외지인거래건수"), 0)}건), '
        f'허가구역 구분 {h(latest.get("허가구역구분") or "-")}'
    )
    trade_rows = (tx.get("top_trades", []) or [])[:5]
    trade_body = "".join(
        f'<tr><td>{h(row.get("계약일") or row.get("신고일"))}</td><td>{h(row.get("소재지"))}</td>'
        f'<td>{h(row.get("물건구분"))}</td><td>{h(row.get("용도지역"))}</td><td>{h(row.get("지목"))}</td>'
        f'<td>{fmt_num(row.get("면적_㎡"), 1)}</td><td>{fmt_num(row.get("거래금액_억원"), 2)}</td>'
        f'<td>{_fmt_won_html(row.get("단가_원㎡"))}</td>'
        f'<td>{h(row.get("거래주체") or "미상")}</td>'
        f'<td>{h(row.get("매수자시군구") or row.get("매수자지역") or "미상")}'
        + ("<b class=\"ext-mark\"> (외지인)</b>" if row.get("외지인") == "Y" else "")
        + '</td><td class="form-blank-cell"></td></tr>'
        for row in trade_rows
    ) or '<tr><td colspan="11" class="blank-cell">당월 거래 자료 없음</td></tr>'

    return f"""
<div class="survey-form">
  <div class="form-meta">
    <span><b>조사지역</b> {h(region)}</span>
    <span><b>기준주</b> {h(report_week.get("주차") or "-")} (종료일 {h(report_week.get("종료일") or "-")}, 금요일 작성)</span>
    <span><b>실거래 기준</b> 계약일</span>
  </div>
  <div class="form-box">
    <div class="form-title">□ 토지시장 동향 (가격변동 추이, 실거래 현황, 매매동향 등)</div>
    <div class="form-item"><div class="form-bullet">○ 시군구 내 전반적인 토지시장 동향</div>
      <div class="form-fact">계약일 기준 집계: {overall_fact}</div>
      {f'<div class="monitor-note"><div class="monitor-title">중개사무소 모니터링 청취 요지 <span class="monitor-src">{h(monitoring.get("source"))}' + (f' · 청취기준일 {h(monitoring.get("as_of"))}' if monitoring.get("as_of") else "") + f'</span></div><p>{h(monitoring.get("overall"))}</p></div>' if monitoring.get("available") and monitoring.get("overall") else ""}
      {f'<div class="form-draft">{h(survey.get("overall"))}</div>' if survey.get("overall") else ""}
      <div class="form-blank tall"></div></div>
    {group_blocks}
    <div class="form-item"><div class="form-bullet">* 매매 호가 및 실거래 현황</div>
      {f'<div class="form-draft">{h(survey.get("price_trend"))}</div>' if survey.get("price_trend") else ""}
      {f'<div class="monitor-note"><div class="monitor-title">중개사무소 청취 요지 <span class="monitor-src">{h(monitoring.get("source"))}</span></div><p>{h(monitoring.get("price_trend"))}</p></div>' if monitoring.get("available") and monitoring.get("price_trend") else ""}
      <div class="table-wrap"><table class="rk survey-trade-table"><thead><tr>
        <th>계약일</th><th>소재지</th><th>물건구분</th><th>용도지역</th><th>지목</th><th>면적(㎡)</th>
        <th>거래금액(억)</th><th>단가(원/㎡)</th><th>거래주체</th><th>매수자 시군구</th>
        <th>호가·현장 확인사항</th></tr></thead>
        <tbody>{trade_body}</tbody></table></div>
      <div class="form-blank"></div></div>
  </div>
  <div class="form-box">
    <div class="form-title">□ 향후 시장 전망</div>
    {f'<div class="monitor-note"><div class="monitor-title">중개사무소 청취 요지 <span class="monitor-src">{h(monitoring.get("source"))}</span></div><p>{h(monitoring.get("outlook"))}</p></div>' if monitoring.get("available") and monitoring.get("outlook") else ""}
    {f'<div class="form-draft">{h(survey.get("market_outlook"))}</div>' if survey.get("market_outlook") else ""}
    <div class="form-blank xtall"></div>
  </div>
  <p class="note">※ 업로드 양식의 동별 항목(송정동·도산동·신촌동)을 <b>직접 영향권(송정·도산·신촌·소촌·우산·황룡동)</b>과 <b>간접 영향권(직접 영향권 외 읍·면·동)</b> 두 묶음으로 재구성함. 자동 표기된 수치는 계약일 기준 실거래 집계값이고, 음영 문단은 지역 특성·실거래 내용을 반영한 LLM 서술 초안이며, 호가·현장 확인사항과 최종 문구는 조사담당자가 검토·보완함.</p>
</div>"""


def _fmt_money_html(value: Any) -> str:
    return h(fmt_num(value, 1))


def _fmt_won_html(value: Any) -> str:
    return h(fmt_num(value, 0))


def _fmt_pct_html(value: Any) -> str:
    return h(fmt_num(value, 1, "%"))


def _fmt_rate3_html(value: Any) -> str:
    return h(fmt_num(value, 3, "%"))


def _reference_population_rows(pop: dict[str, Any], hh: dict[str, Any]) -> list[dict[str, Any]]:
    pop_source = pop.get("monthly_rows_with_prev", []) or pop.get("monthly_rows", [])
    hh_source = hh.get("monthly_rows_with_prev", []) or hh.get("monthly_rows", [])
    pop_map = {str(x.get("월")): x.get("값") for x in pop_source}
    hh_map = {str(x.get("월")): x.get("값") for x in hh_source}
    months = sorted(set(pop_map) | set(hh_map))[-14:]
    rows = []
    prev_pop = prev_hh = None
    for month in months:
        pop_value = pop_map.get(month)
        hh_value = hh_map.get(month)
        pop_delta = pop_value - prev_pop if pop_value is not None and prev_pop is not None else None
        hh_delta = hh_value - prev_hh if hh_value is not None and prev_hh is not None else None
        rows.append({
            "월": month,
            "인구수": pop_value,
            "인구증감": pop_delta,
            "인구증감률": round(pop_delta / prev_pop * 100, 1) if pop_delta is not None and prev_pop else None,
            "세대수": hh_value,
            "세대증감": hh_delta,
            "세대증감률": round(hh_delta / prev_hh * 100, 1) if hh_delta is not None and prev_hh else None,
        })
        if pop_value is not None:
            prev_pop = pop_value
        if hh_value is not None:
            prev_hh = hh_value
    return rows


def _signed_html(value: Any) -> str:
    if value is None:
        return "-"
    try:
        number = int(round(float(value)))
    except Exception:
        return h(value)
    cls = "up" if number > 0 else ("dn" if number < 0 else "g")
    sign = "+" if number > 0 else ""
    return f"<span class='{cls}'>{sign}{number:,}</span>"


def _rate_value_html(value: Any) -> str:
    try:
        if value is None or pd.isna(value):
            return "-"
        return f"{float(value):.2f}"
    except Exception:
        return h(value)


def _rate_diff_html(value: Any) -> str:
    try:
        if value is None or pd.isna(value):
            return "-"
        number = float(value)
        cls = "up" if number > 0 else ("dn" if number < 0 else "g")
        return f"<span class='{cls}'>{number:+.2f}</span>"
    except Exception:
        return h(value)


def _count_diff_html(value: Any) -> str:
    """건수·금액 증감은 부호 있는 정수로 표시한다."""
    try:
        if value is None or pd.isna(value):
            return "-"
        number = float(value)
        cls = "up" if number > 0 else ("dn" if number < 0 else "g")
        return f"<span class='{cls}'>{number:+,.0f}</span>"
    except Exception:
        return h(value)


def _rate_diff3_html(value: Any) -> str:
    """지가변동률 전월 대비(%p)는 변동률과 같은 소수점 셋째 자리로 표시한다."""
    try:
        if value is None or pd.isna(value):
            return "-"
        number = float(value)
        cls = "up" if number > 0 else ("dn" if number < 0 else "g")
        return f"<span class='{cls}'>{number:+.3f}</span>"
    except Exception:
        return h(value)


def _rates_standard_table(rates: dict[str, Any]) -> str:
    """주택동향 주·월간의 예금은행 대출금리 13개월 표준표."""
    rows = rates.get("table_rows_with_prev", []) or rates.get("table_rows", [])
    if not rows:
        return "<p class='note'>예금은행 대출금리 원자료 없음.</p>"
    columns = [
        ("mortgage_rate", "주택담보대출"), ("jeonse_rate", "전세자금대출"),
        ("sme_rate", "중소기업대출"), ("business_rate", "기업일반자금대출"),
    ]
    display = rows[-13:]
    prev = rows[-14] if len(rows) > 13 else None
    body = []
    for idx, row in enumerate(display):
        cells = []
        for key, _label in columns:
            diff = row.get(f"{key}_mom")
            if diff is None and prev:
                try:
                    diff = float(row.get(key)) - float(prev.get(key))
                except Exception:
                    diff = None
            cells.append(f"<td>{_rate_value_html(row.get(key))}</td><td>{_rate_diff_html(diff)}</td>")
        cls = " class='latest-row'" if idx == len(display) - 1 else ""
        body.append(f"<tr{cls}><td>{h(row.get('month'))}</td>{''.join(cells)}</tr>")
        prev = row
    head = "".join(f"<th>{label}(%)</th><th>{label} 전월 대비(%p)</th>" for _, label in columns)
    return (
        "<div class='table-wrap rates-standard-table-wrap'><table class='time-series-table rates-standard-table'>"
        f"<thead><tr><th>기준월</th>{head}</tr></thead><tbody>{''.join(body)}</tbody></table></div>"
    )


def _population_migration_table(pop: dict[str, Any], hh: dict[str, Any], mig: dict[str, Any]) -> str:
    """주택동향 주·월간의 인구·세대수·인구이동 통합표 형식."""
    rows = _reference_population_rows(pop, hh)
    display = rows[-13:]
    migration_rows = mig.get("table_rows", []) or mig.get("monthly_rows", [])
    migration_map = {str(x.get("month") or x.get("월")): x for x in migration_rows}
    latest_migration = mig.get("latest_month") or ((migration_rows or [{}])[-1].get("month") or (migration_rows or [{}])[-1].get("월", "-"))
    body = []
    for idx, row in enumerate(display):
        month = str(row.get("월") or "")
        migration_row = migration_map.get(month)
        if migration_row:
            in_html = _fmt_int_html(migration_row.get("in"))
            out_html = _fmt_int_html(migration_row.get("out"))
            net_html = _signed_html(migration_row.get("net") if "net" in migration_row else migration_row.get("값"))
        else:
            in_html = out_html = net_html = "<span class='not-published'>미공표</span>"
        cls = " class='latest-row'" if idx == len(display) - 1 else ""
        body.append(
            f"<tr{cls}><td>{h(month)}</td><td>{_fmt_int_html(row.get('인구수'))}</td>"
            f"<td>{_signed_html(row.get('인구증감'))}</td><td>{_fmt_int_html(row.get('세대수'))}</td>"
            f"<td>{_signed_html(row.get('세대증감'))}</td><td>{in_html}</td><td>{out_html}</td><td>{net_html}</td></tr>"
        )
    return (
        f"<div class='population-publish-note'>인구·세대수는 기준월 자료, 순이동은 최신 공표월(<b>{h(latest_migration)}</b>) 기준으로 표시함.</div>"
        "<div class='table-wrap population-migration-table-wrap'><table class='time-series-table population-migration-table'>"
        "<thead><tr><th>기준월</th><th>총인구수</th><th>전월 대비 증감</th><th>총세대수</th><th>세대 전월 대비 증감</th><th>총 전입</th><th>총 전출</th><th>순이동</th></tr></thead>"
        f"<tbody>{''.join(body)}</tbody></table></div>"
    )


def _age_population_trend_table(age_population: dict[str, Any]) -> str:
    rows = age_population.get("monthly_rows_with_prev", []) or age_population.get("monthly_rows", [])
    if not rows:
        return "<p class='note'>연령대별 주민등록인구 원자료 없음.</p>"
    age_cols = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대 이상"]
    display, prev = rows[-13:], (rows[-14] if len(rows) > 13 else None)
    body = []
    for idx, row in enumerate(display):
        cells = []
        for col in age_cols:
            value = row.get(col)
            pct = None
            try:
                old = float(prev.get(col)) if prev else None
                pct = (float(value) - old) / old * 100 if old else None
            except Exception:
                pass
            rate = "-" if pct is None else f"{pct:+.1f}%"
            rate_cls = "up" if pct and pct > 0 else ("dn" if pct and pct < 0 else "g")
            cells.append(f"<td>{_fmt_int_html(value)} <span class='age-rate {rate_cls}'>({rate})</span></td>")
        cls = " class='latest-row'" if idx == len(display) - 1 else ""
        body.append(f"<tr{cls}><td>{h(row.get('month'))}</td>{''.join(cells)}</tr>")
        prev = row
    return (
        "<div class='table-wrap age-pop-trend-table-wrap'><table class='time-series-table age-pop-trend-table'>"
        f"<thead><tr><th>기준월</th>{''.join(f'<th>{c}</th>' for c in age_cols)}</tr></thead>"
        f"<tbody>{''.join(body)}</tbody></table></div>"
    )


def _age_population_structure_table(age_population: dict[str, Any]) -> str:
    rows = age_population.get("age_rows", [])
    return _reference_table(rows, [
        ("연령대", "연령대", None), ("최신월 인구", "최신월 인구(명)", _fmt_int_html),
        ("비중", "비중(%)", _fmt_pct_html), ("13개월 증감", "13개월 증감(명)", _signed_html),
    ], 5, "age-population-structure-table")


_DEV_NEWS_STYLE = (
    "<style>"
    ".dev-project-table td{vertical-align:top;line-height:1.7;font-size:12.5px;text-align:left}"
    ".dev-project-table td:nth-child(1){font-weight:800;color:#10325f;width:17%}"
    ".dev-project-table td:nth-child(2),.dev-project-table td:nth-child(3){width:21%}"
    ".dev-project-table .dev-meta{display:block;font-weight:500;color:#64748b;font-size:11.5px;margin-top:4px}"
    ".dev-project-table ul,.dev-news-table ul{margin:4px 0 0;padding-left:16px}"
    ".dev-similar{color:#b42318;font-size:11px;font-weight:700;white-space:nowrap}"
    ".dev-count{display:block;color:#64748b;font-size:11.5px;margin-top:4px}"
    ".dev-news-table td{vertical-align:top;font-size:12.5px;line-height:1.6;text-align:left}"
    ".dev-news-table td:nth-child(1){font-weight:800;color:#10325f;width:15%}"
    ".dev-news-table td[rowspan]:nth-child(2){width:27%}"
    ".dev-news-table a,.dev-project-table a,.news a{color:#1d4ed8;text-decoration:none}"
    "</style>"
)


DEV_PROGRESS_MAX_CHARS = 50


def fit_text(text: Any, limit: int) -> str:
    """글자수 제한에 맞추되 가능하면 어절·구두점 경계에서 끊는다."""
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(text) <= limit:
        return text
    cut = text[:limit]
    pos = max(cut.rfind(mark) for mark in " ,·…")
    return (cut[:pos] if pos >= limit * 0.6 else cut).rstrip(" ,·…")


def _news_day_label(value: Any) -> str:
    match = re.match(r"\d{4}-(\d{2})-(\d{2})", str(value or ""))
    return f"{int(match.group(1))}.{int(match.group(2))}." if match else str(value or "-")


def _news_title_html(item: dict[str, Any]) -> str:
    url = str(item.get("url") or "")
    title = h(item.get("title"))
    linked = f'<a href="{h(url)}" target="_blank" rel="noopener">{title}</a>' if url.startswith(("http://", "https://")) else title
    similar = int(as_number(item.get("similar_count")) or 0)
    return linked + (f" <span class='dev-similar'>외 {similar}건</span>" if similar else "")


def _project_news_summary(project: dict[str, Any], update: dict[str, Any]) -> tuple[str, bool]:
    """사업별 뉴스 요약(100자 이내)과 '기사 제목 기준 여부'. 목록 파일 입력값 → LLM 요약 → 보도 폭이 큰 기사 제목 순."""
    summary = str(project.get("news_summary") or update.get("recent_progress") or "").strip()
    title_based = False
    news = project.get("news") or []
    if not summary and news:
        ranked = sorted(news, key=lambda item: -int(as_number(item.get("similar_count")) or 0))
        summary = " / ".join(str(item.get("title") or "") for item in ranked[:2])
        title_based = True
    return limit_text(summary, SECTION_CHAR_LIMITS["development_news_summary"]), title_based


def _development_updates(narratives: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {
        norm_text(item.get("project")): item
        for item in (narratives.get("development_project_updates") or []) if isinstance(item, dict)
    }


def _development_report_html(
    dev: dict[str, Any], narratives: dict[str, Any], region: str, month_title: str, heading_numbers: tuple[str, str],
) -> str:
    """6장에 넣는 기준월 주요 뉴스 3선과 주요 개발사업 진척 현황 표."""
    if not dev.get("available"):
        return ""
    parts: list[str] = [_DEV_NEWS_STYLE]
    headlines = dev.get("monthly_headlines") or []
    top_news = [item for item in (narratives.get("monthly_top_news") or []) if isinstance(item, dict)]
    top_by_title = {norm_text(item.get("title")): item for item in top_news}
    parts.append(f"<h3>{heading_numbers[0]} {h(month_title)} {h(region)} 주요 부동산 뉴스 {len(headlines) or NEWS_HEADLINE_COUNT}선</h3>")
    if headlines:
        for i, item in enumerate(headlines, 1):
            written = top_by_title.get(norm_text(item.get("title"))) or (top_news[i - 1] if len(top_news) == len(headlines) else {})
            related = f" · 관련 사업: {h(item.get('related_project'))}" if item.get("related_project") else ""
            news_summary = limit_text(item.get("summary") or written.get("summary"), SECTION_CHAR_LIMITS["monthly_news_summary"])
            parts.append(
                f"<div class='news'><div class='ttl'>{i}. {_news_title_html(item)}</div>"
                f"<p class='note'>보도일: {h(item.get('date'))} · 매체: {h(item.get('source'))}{related}</p>"
                + (f"<p><b>· 내용 요약:</b> {h(news_summary)}</p>" if news_summary else "")
                + (f"<p><b>· 시사점:</b> {h(written.get('impact'))}</p>" if written.get("impact") else "")
                + "</div>"
            )
    else:
        parts.append("<p class='note'>기준월에 수집된 관련 보도가 없음.</p>")
    projects = dev.get("projects") or []
    if projects:
        window = dev.get("window", {}) or {}
        months = [str(month) for month in (window.get("months") or [])][-2:]
        base = months[-1] if months else ""
        updates = _development_updates(narratives)
        has_impact = any(str(item.get("impact") or "").strip() for item in updates.values())

        def month_progress(project: dict[str, Any], update: dict[str, Any], month: str) -> str:
            """월별 진행상황 한 줄(50자 이내): 목록 파일 입력값 → LLM 요약(기준월) → 최다 보도 제목 순."""
            entered = [str(e.get("text") or "").strip() for e in project.get("progress") or [] if str(e.get("month")) == month]
            entered = [text for text in entered if text]
            if entered:
                return h(fit_text("; ".join(entered), DEV_PROGRESS_MAX_CHARS))
            if month == base and str(update.get("monthly_progress") or "").strip():
                return h(fit_text(update.get("monthly_progress"), DEV_PROGRESS_MAX_CHARS))
            month_news = [item for item in project.get("news") or [] if item.get("month") == month]
            if month_news:
                top = max(month_news, key=lambda item: int(as_number(item.get("similar_count")) or 0))
                return h(fit_text(top.get("title"), DEV_PROGRESS_MAX_CHARS)) + "<span class='dev-meta'>보도 제목 기준</span>"
            return "<span class='dev-meta'>수집된 관련 보도 없음</span>"

        dev_empty = "<span class='dev-meta'>수집된 관련 보도 없음</span>"
        rows: list[str] = []
        for project in projects:
            update = updates.get(norm_text(project.get("name")), {})
            meta_bits = [f"등록 {project.get('registered')}" if project.get("registered") else "", project.get("type"), project.get("status")]
            name_cell = h(project.get("name")) + f"<span class='dev-meta'>{h(' · '.join(str(b) for b in meta_bits if b))}</span>"
            if update.get("stage"):
                name_cell += f"<span class='dev-meta'>추진단계: {h(update.get('stage'))}</span>"
            counts = " · ".join(
                f"{int(str(month)[-2:])}월 {count}건" for month, count in (project.get("month_counts") or {}).items()
            )
            summary, title_based = _project_news_summary(project, update)
            evidence = (
                h(summary) + ("<span class='dev-meta'>보도 제목 기준</span>" if title_based else "")
                + f"<span class='dev-count'>수집 기사: {h(counts)}</span>"
            ) if summary else ""
            rows.append(
                f"<tr><td>{name_cell}</td>"
                + "".join(f"<td>{month_progress(project, update, month)}</td>" for month in reversed(months))
                + f"<td>{evidence or dev_empty}</td>"
                + (f"<td>{h(update.get('impact') or '-')}</td>" if has_impact else "")
                + "</tr>"
            )
        month_heads = "".join(
            f"<th>{'당월' if month == base else '전월'} 진행상황 ({int(month[-2:])}월)</th>" for month in reversed(months)
        )
        parts.append(
            f"<h3>{heading_numbers[1]} 주요 개발사업 진척 현황 (최근 2개월: {h(window.get('label'))})</h3>"
            f"<table class='dev-project-table'><thead><tr><th>사업명</th>{month_heads}"
            f"<th>최근 2개월 뉴스 요약 (100자 이내)</th>{'<th>지가 영향·확인사항</th>' if has_impact else ''}</tr></thead>"
            f"<tbody>{''.join(rows)}</tbody></table>"
        )
    failed = len(dev.get("errors") or [])
    parts.append(
        f"<p class='note'>※ {h(dev.get('news_source'))}. "
        f"기사 제목 기준이므로 사실관계는 원문·1차 출처 확인이 필요하며, 사업별 전체 목록은 부록 2 참조. 수집일 {h(str(dev.get('collected_at') or '')[:10])}"
        + (f" · 일부 검색어 수집 실패 {failed}건" if failed else "") + "</p>"
    )
    return "".join(parts)


def _development_news_appendix_html(dev: dict[str, Any], narratives: dict[str, Any]) -> str:
    """부록 2에 붙이는 사업별 최근 2개월 주요 내용 요약(100자 이내)과 보도 목록."""
    projects = [project for project in (dev.get("projects") or []) if dev.get("available")]
    if not projects:
        return ""
    updates = _development_updates(narratives)
    rows: list[str] = []
    for project in projects:
        news = project.get("news") or []
        summary, title_based = _project_news_summary(project, updates.get(norm_text(project.get("name")), {}))
        summary_html = (
            h(summary) + ("<span class='dev-meta'>보도 제목 기준</span>" if title_based else "")
        ) if summary else "<span class='dev-meta'>수집된 관련 보도 없음</span>"
        span = max(1, len(news))
        lead = f"<td rowspan='{span}'>{h(project.get('name'))}</td><td rowspan='{span}'>{summary_html}</td>"
        if not news:
            rows.append(f"<tr>{lead}<td colspan='3'>-</td></tr>")
            continue
        for i, item in enumerate(news):
            rows.append(
                f"<tr>{lead if i == 0 else ''}<td>{h(item.get('date'))}</td>"
                f"<td>{h(item.get('source'))}{' (공공)' if item.get('source_type') == '공공·보도자료' else ''}</td>"
                f"<td>{_news_title_html(item)}</td></tr>"
            )
    return (
        f"{_DEV_NEWS_STYLE}<h3>참고: 주요 개발사업별 최근 2개월 보도 목록 ({h((dev.get('window') or {}).get('label'))})</h3>"
        "<table class='dev-news-table'><thead><tr><th>사업명</th><th>주요 뉴스 내용 요약 (100자 이내)</th>"
        "<th>보도일</th><th>매체</th><th>기사 제목</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table>"
        f"<p class='note'>※ {h(dev.get('news_source'))}. '외 n건'은 같은 사안을 다룬 다른 매체 보도 수이며, "
        f"사업별 최대 {NEWS_MAX_STORIES_PER_PROJECT}건(월별 보도 폭이 큰 순)을 실음. 요약은 기사 제목에서 확인되는 내용만 담음.</p>"
    )


def build_html_report(result: dict[str, Any], narratives: dict[str, Any] | None = None) -> str:
    """업로드한 나주시 보고서의 장·표·차트 구성을 표준 월간 양식으로 사용한다."""
    n = {**default_narratives(result), **(narratives or {})}
    meta = result.get("meta", {})
    tx = result.get("transactions", {})
    lp = result.get("land_price", {})
    pop = result.get("population", {})
    hh = result.get("households", {})
    age_population = result.get("age_population", {})
    mig = result.get("migration", {})
    rates = result.get("rates", {})
    listings = result.get("listings", {})
    issues = result.get("regional_issues", {})
    regional_profile = result.get("regional_profile", {})
    region = str(meta.get("region") or "조사지역")
    base_month = str(meta.get("base_month") or "")
    tx_data_type = str(tx.get("data_type") or "잠정치")
    tx_data_type_note = str(tx.get("data_type_note") or "")
    has_gongsi = any(row.get("공시지가_원㎡") for row in (tx.get("top_trades", []) or []))
    try:
        period = pd.Period(base_month, freq="M")
        month_title = f"{period.year}년 {period.month}월"
    except Exception:
        month_title = base_month

    def month_label(value: Any) -> str:
        try:
            parsed = pd.Period(str(value), freq="M")
            return f"{parsed.year}년 {parsed.month}월"
        except Exception:
            return str(value or "-")

    current = tx.get("current", {}) if tx.get("available") else {}
    previous = tx.get("previous", {}) if tx.get("available") else {}
    change = tx.get("change_vs_previous", {}) if tx.get("available") else {}
    yoy = tx.get("year_over_year", {}) if tx.get("available") else {}
    lp_current = lp.get("current", {}) if lp.get("available") else {}
    lp_cum = lp.get("cumulative", {}) if lp.get("available") else {}
    latest_public_month = lp.get("latest_public_month") or lp.get("latest_month")
    price_survey_month = ((result.get("time_series_context", {}) or {}).get("land_price", {}) or {}).get("investigator_month")
    if not price_survey_month:
        try:
            price_survey_month = str(pd.Period(base_month, freq="M") - 1)
        except Exception:
            price_survey_month = "-"
    public_rate_availability_note = (
        f"※ 시군구 지가변동률 최신 확인월: {h(latest_public_month)}. 이후 월은 자료 미확인 시 '-'로 표시함."
        if latest_public_month else
        "※ 시군구 지가변동률 파일이 연결되지 않아 변동률을 '-'로 표시함."
    )
    monthly_rows = tx.get("monthly_rows", []) if tx.get("available") else []

    key_takeaways = n.get("key_takeaways") or []
    takeaways_html = "".join(f"<li>{h(x)}</li>" for x in key_takeaways) or "<li>조사담당자 서술 입력 전임.</li>"
    population_rows = _reference_population_rows(pop, hh)
    population_table = _population_migration_table(pop, hh, mig)
    age_population_trend_table = _age_population_trend_table(age_population)
    age_population_structure_table = _age_population_structure_table(age_population)
    rates_standard_table = _rates_standard_table(rates)
    monthly_table = _reference_table(monthly_rows, [
        ("월", "신고월", None),
        ("공표변동률", "시군구 변동률(%)", _fmt_rate3_html),
        ("거래건수", "전체건수", _fmt_int_html),
        ("거래건수증감률", "증감", _change_badge),
        ("거래금액_억원", "전체금액(억)", _fmt_money_html),
        ("거래금액증감률", "증감", _change_badge),
        ("외지인거래건수", "외지건수", _fmt_int_html),
        ("외지인건수증감률", "증감", _change_badge),
        ("외지인거래금액_억원", "외지금액(억)", _fmt_money_html),
        ("외지인금액증감률", "증감", _change_badge),
        ("외지인금액비중", "외지금액비중", _fmt_pct_html),
    ])
    public_rate_rows = lp.get("monthly_rows", []) if lp.get("available") else []
    public_rate_columns = [("월", "확인월", None), ("평균", "시군구 평균", _fmt_rate3_html)]
    for key in ["주거", "상업", "공업", "녹지", "보전관리", "생산관리", "계획관리", "전", "답", "주거용", "상업용", "임야", "공업용", "기타"]:
        if any(row.get(key) is not None for row in public_rate_rows):
            public_rate_columns.append((key, key, _fmt_rate3_html))
    public_rate_table = _reference_table(public_rate_rows, public_rate_columns, 5, "public-rate-table")

    rank_columns = [
        ("순위", "순위", _fmt_int_html), ("명칭", "구분", None),
        ("거래건수", "거래건수", _fmt_int_html), ("건수비중", "건수비중(%)", _fmt_pct_html),
        ("거래금액_억원", "거래금액(억)", _fmt_money_html), ("금액비중", "금액비중(%)", _fmt_pct_html),
    ]
    def ranked(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return [{"순위": i, **row} for i, row in enumerate(rows, 1)]

    # 과거 산출물(신고일 키)로 재생성하는 경우에도 계약일 열이 비지 않도록 보정한다.
    for trade_key in ["top_trades", "external_top_trades"]:
        for trade_row in tx.get(trade_key, []) or []:
            if not trade_row.get("계약일"):
                trade_row["계약일"] = trade_row.get("신고일", "")

    emd_ranked = ranked(tx.get("emd_top", []))
    for row in emd_ranked:
        row["명칭표시"] = _emd_label_html(row.get("명칭"), region)
    emd_rank_columns = [
        ("순위", "순위", _fmt_int_html), ("명칭표시", "읍면동", lambda value: str(value or "")),
        ("거래건수", "거래건수", _fmt_int_html), ("건수비중", "건수비중(%)", _fmt_pct_html),
        ("거래금액_억원", "거래금액(억원)", _fmt_money_html), ("금액비중", "금액비중(%)", _fmt_pct_html),
    ]
    emd_table = _reference_table(emd_ranked, emd_rank_columns, 5, "rk")
    zone_table = _reference_table(ranked(tx.get("zone_top", [])), rank_columns, 5, "rk")
    use_table = _reference_table(ranked(tx.get("use_top", [])), rank_columns, 5, "rk")
    buyer_table = _reference_table(ranked(tx.get("external_buyer_region_top", [])), rank_columns, 5, "rk")
    external_use_table = _reference_table(ranked(tx.get("external_use_top", [])), rank_columns, 5, "rk")
    external_top_trade_table = "<div class='table-wrap top-trade-table-wrap'>" + _reference_table(tx.get("external_top_trades", []), [
        ("순위", "순위", _fmt_int_html), ("계약일", "계약일", None), ("신고일", "신고일", None), ("소재지", "소재지", None),
        ("물건구분", "물건구분", None), ("용도지역", "용도지역", None), ("지목", "지목", None),
        ("면적_㎡", "면적(㎡)", _fmt_money_html), ("면적_평", "면적(평)", _fmt_money_html),
        ("거래금액_억원", "거래금액(억)", _fmt_money_html), ("단가_원㎡", "토지 단가(원/㎡)", _fmt_won_html),
    ] + ([("공시지가_원㎡", "개별공시지가(원/㎡)", _fmt_won_html), ("공시지가배율", "공시지가 배율", _fmt_money_html)] if has_gongsi else []) + [
        ("거래주체", "거래주체", None), ("매수자지역", "외지인 매수자 지역", None),
    ], 10, "rk top-trade-table") + "</div>"
    top_trade_table = "<div class='table-wrap top-trade-table-wrap'>" + _reference_table(tx.get("top_trades", []), [
        ("순위", "순위", _fmt_int_html), ("계약일", "계약일", None), ("신고일", "신고일", None), ("소재지", "소재지", None), ("물건구분", "물건구분", None),
        ("용도지역", "용도지역", None), ("지목", "지목", None), ("면적_㎡", "면적(㎡)", _fmt_money_html), ("면적_평", "면적(평)", _fmt_money_html),
        ("거래금액_억원", "거래금액(억)", _fmt_money_html), ("단가_원㎡", "토지 단가(원/㎡)", _fmt_won_html),
    ] + ([("공시지가_원㎡", "개별공시지가(원/㎡)", _fmt_won_html), ("공시지가배율", "공시지가 배율", _fmt_money_html)] if has_gongsi else []) + [
        ("거래주체", "거래주체", None), ("매수자지역", "매수자 지역", None), ("외지인", "외지인", None),
    ], 10, "rk top-trade-table") + "</div>"

    lp_prev = lp.get("previous", {}) if lp.get("available") else {}
    lp_diff = lp.get("change_vs_previous", {}) if lp.get("available") else {}
    has_prev_rate = bool(lp_prev)

    def rate_rows(names: list[str]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for name in names:
            if lp_current.get(name) is None:
                continue
            rows.append({
                "구분": name,
                "당월": lp_current.get(name),
                "전월": lp_prev.get(name),
                "전월대비": lp_diff.get(name),
                "누계": lp_cum.get(name),
            })
        return rows

    rate_columns = [("구분", "구분", None), ("당월", "당월 변동률", _fmt_rate3_html)]
    if has_prev_rate:
        rate_columns += [("전월", "전월 변동률", _fmt_rate3_html), ("전월대비", "전월 대비(%p)", _rate_diff3_html)]
    rate_columns += [("누계", "누계 변동률", _fmt_rate3_html)]
    zone_rate_rows = rate_rows(["주거", "상업", "공업", "녹지", "보전관리", "생산관리", "계획관리"])
    use_rate_rows = rate_rows(["전", "답", "주거용", "상업용", "임야", "공업용", "기타"])
    lp_zone_table = _reference_table(zone_rate_rows, [("구분", "용도지역", None)] + rate_columns[1:])
    lp_use_table = _reference_table(use_rate_rows, [("구분", "이용상황", None)] + rate_columns[1:])
    lp_emd_rows = [dict(row) for row in (lp.get("emd_top", []) if lp.get("available") else [])][:5]
    for row in lp_emd_rows:
        row["읍면동표시"] = _emd_label_html(row.get("읍면동"), region)
    lp_emd_columns = [
        ("읍면동표시", "읍·면·동", lambda value: str(value or "")),
        ("당월변동률", "당월 변동률", _fmt_rate3_html),
    ]
    if any(row.get("전월변동률") is not None for row in lp_emd_rows):
        lp_emd_columns += [("전월변동률", "전월 변동률", _fmt_rate3_html), ("전월대비", "전월 대비(%p)", _rate_diff3_html)]
    lp_emd_columns += [("누계변동률", "누계 변동률", _fmt_rate3_html)]
    lp_emd_table = _reference_table(lp_emd_rows, lp_emd_columns, 5, "rk")

    # result.json에 Python 계산값으로 저장된 판단 플래그를 표·서술에 공통 사용한다.
    judgment_data = result.get("judgment", {})
    if not judgment_data.get("top_areas") and lp_emd_rows:
        judgment_data = build_land_price_judgment(tx, lp, region, regional_profile)
    judgment_rows: list[dict[str, Any]] = [{
        "순위": row.get("rank"),
        "권역원문": row.get("area"),
        "권역": _emd_label_html(row.get("area"), region),
        "당월": row.get("current_rate"),
        "누계": row.get("cumulative_rate"),
        "거래건수": row.get("transaction_count"),
        "거래금액": row.get("transaction_amount_eok"),
        "판단구분": row.get("flag"),
        "판단근거": row.get("basis"),
        "분석권역": row.get("regional_zone"),
        "지역특성": row.get("local_characteristics"),
        "수요배경": row.get("demand_background"),
        "토지이용": row.get("land_use"),
        "배후수요": row.get("rear_demand"),
        "교통": row.get("transport"),
        "상권": row.get("commercial_area"),
        "현장확인": row.get("field_checks"),
        "산단연계": row.get("industrial_links") or [],
        "유의사항": row.get("analysis_caution"),
    } for row in judgment_data.get("top_areas", [])]
    judgment_table = _reference_table(judgment_rows, [
        ("순위", "순위", _fmt_int_html),
        ("권역", "권역", lambda value: str(value or "")),
        ("당월", "당월 변동률", _fmt_rate3_html),
        ("누계", "누계 변동률", _fmt_rate3_html),
        ("거래건수", "당월 실거래", lambda value: f"{_fmt_int_html(value)}건"),
        ("거래금액", "거래금액(억)", _fmt_money_html),
        ("판단구분", "구분", None),
        ("판단근거", "변동률 판단 근거", None),
    ], 5, "rk")
    area_narratives = n.get("emd_area_narratives") if isinstance(n.get("emd_area_narratives"), list) else []
    area_narrative_map = {norm_text(item.get("area")): str(item.get("narrative") or "").strip() for item in area_narratives if isinstance(item, dict)}

    # LLM 서술이 없거나 확인 요청·판단 기준 문장만 남으면 Python 서술로 대체한다.
    fallback_area_map = {
        norm_text(item.get("area")): item.get("narrative", "")
        for item in build_investigator_emd_narratives({**result, "judgment": judgment_data})
    }

    def area_narrative(row: dict[str, Any]) -> str:
        key = norm_text(row.get("권역원문"))
        written = drop_sentences(area_narrative_map.get(key), SECTION9_BANNED)
        return limit_text(written or fallback_area_map.get(key, ""), SECTION_CHAR_LIMITS["emd_area_narrative"])

    section9_lead_text = drop_sentences(n.get("emd_change_reasoning"), SECTION9_BANNED) or n.get("land_price_interpretation")

    emd_reason_cards = "".join(
        "<div class='zone'>"
        f"<p><b>{row['순위']}. {row['권역']} ({_fmt_rate3_html(row['당월'])})</b> "
        + ("" if str(row["판단구분"]) == "가격조사 중심" else f"<span class='tag'>{h(row['판단구분'])}</span>")
        + "</p><p>"
        + h(area_narrative(row))
        + "</p></div>"
        for row in judgment_rows
    ) or "<p class='note'>읍면동별 지가변동률 자료 없음.</p>"

    listing_parsed = listings.get("available") and not listings.get("template_only")
    listing_count_columns = [
        ("당월_매물수", "당월 매물수", _fmt_int_html),
        ("전월_매물수", "전월 매물수", _fmt_int_html),
        ("증감", "증감", _count_diff_html),
        ("평균단가_원㎡", "평균 호가단가(원/㎡)", _fmt_won_html),
        ("중위단가_원㎡", "중위 호가단가(원/㎡)", _fmt_won_html),
    ]
    if listing_parsed:
        listing_file_table = _reference_table(listings.get("use_rows", []), [("이용상황", "이용상황", None)] + listing_count_columns, 5, "rk")
        emd_rows_listing = [dict(row) for row in listings.get("emd_rows", [])]
        for row in emd_rows_listing:
            row["읍면동표시"] = _emd_label_html(row.get("읍면동"), region)
        listing_blank_emd = _reference_table(
            emd_rows_listing,
            [("읍면동표시", "읍·면·동", lambda value: str(value or ""))] + listing_count_columns, 5, "rk",
        )
        listing_blank_change = _reference_table(listings.get("zone_rows", []), [("용도지역", "용도지역", None)] + listing_count_columns, 5, "rk")
        price_changes = listings.get("price_changes", {}) or {}
        summary_rows = [{
            "구분": "전월 대비",
            "비교가능": price_changes.get("비교가능_매물수"),
            "호가인상": price_changes.get("호가인상"),
            "호가인하": price_changes.get("호가인하"),
            "호가유지": price_changes.get("호가유지"),
            "평균변동률": price_changes.get("평균변동률"),
            "신규등록": price_changes.get("신규등록"),
            "매물회수": price_changes.get("매물회수"),
        }] if price_changes else []
        listing_price_summary = _reference_table(summary_rows, [
            ("구분", "구분", None), ("비교가능", "비교 가능 매물", _fmt_int_html),
            ("호가인상", "호가 인상", _fmt_int_html), ("호가인하", "호가 인하", _fmt_int_html),
            ("호가유지", "호가 유지", _fmt_int_html), ("평균변동률", "평균 변동률(%)", _rate_diff_html),
            ("신규등록", "신규 등록", _fmt_int_html), ("매물회수", "매물 회수", _fmt_int_html),
        ])
        listing_blank_price = listing_price_summary + _reference_table(price_changes.get("조정사례", []), [
            ("읍면동", "읍·면·동", None), ("이용상황", "이용상황", None), ("용도지역", "용도지역", None), ("지목", "지목", None),
            ("전월_호가_만원", "전월 호가(만원)", _fmt_int_html), ("당월_호가_만원", "당월 호가(만원)", _fmt_int_html),
            ("변동_만원", "변동(만원)", _count_diff_html), ("변동률", "변동률(%)", _rate_diff_html),
        ], 5, "rk")
    else:
        listing_file_table = _reference_table(listings.get("files", []) if listings.get("available") else [], [
            ("source_file", "매물 파일", None), ("estimated_listing_count", "추정 매물건수", _fmt_int_html), ("read_error", "읽기 오류", None),
        ])
        listing_blank_emd = _reference_table([], [("순위", "순위", None), ("읍면동", "읍면동", None), ("매물건수", "매물건수", None), ("매물금액", "매물금액(억)", None), ("비중", "건수비중(%)", None)])
        listing_blank_change = _reference_table([], [("읍면동", "읍면동", None), ("증감건수", "증감건수", None)])
        listing_blank_price = _reference_table([], [("순위", "순위", None), ("이용상황", "이용상황", None), ("표본수", "총 표본수", None), ("상향", "상향(%)", None), ("보합", "보합(%)", None), ("하향", "하향(%)", None)])
    listing_change = listings.get("change_vs_previous", {}) or {}
    listing_direct = listings.get("direct_zone_summary", {}) or {}
    listing_headline = (
        f"당월 매물 {fmt_num(listings.get('total_current'), 0)}건 · 전월 {fmt_num(listings.get('total_previous'), 0)}건"
        f"({fmt_num(listing_change.get('증감'), 0)}건, {fmt_num(listing_change.get('증감률'), 1)}%) · "
        f"반도체 직접 영향권 매물 {fmt_num(listing_direct.get('당월_매물수'), 0)}건"
        f"(비중 {fmt_num(listing_direct.get('당월비중'), 1)}%, 전월 대비 {fmt_num(listing_direct.get('증감'), 0)}건)"
        if listing_parsed else ""
    )

    issue_rows = issues.get("events", []) if issues.get("available") else []
    issue_table = _reference_table(issue_rows, [("event_date", "일자", None), ("title", "주요 이슈", None), ("summary", "시장 점검 방향", None)])
    news_items = n.get("news_items") if isinstance(n.get("news_items"), list) else []
    news_cards_html = "".join(
        f"<div class='news'><div class='ttl'>{i}. {h(item.get('title'))}</div>"
        f"<p class='note'>출처: {h(item.get('source') or '-')} · 발표일: {h(item.get('source_date') or '-')} · 추진단계: {h(item.get('project_stage') or '-')}</p>"
        f"<p>{h(item.get('summary'))}</p></div>"
        for i, item in enumerate(news_items, 1) if isinstance(item, dict) and item.get("title")
    )
    def _news_sort_key(item: dict[str, Any]) -> tuple[int, str]:
        """기준월에 발생한 이슈를 먼저, 그다음 발표일 최신순으로 배열한다."""
        date_text = str(item.get("source_date") or "")
        in_base = 0 if date_text.startswith(str(base_month)) else 1
        return (in_base, "" if in_base else date_text)

    def _news_recency(item: dict[str, Any]) -> int:
        """발표일 최신순 정렬값. 기준연도보다 미래인 완공 계획 연도는 정렬에서 뒤로 보낸다."""
        digits = re.sub(r"[^0-9]", "", str(item.get("source_date") or ""))[:8]
        if len(digits) < 4:
            return 0
        try:
            base_year = int(str(base_month)[:4])
        except Exception:
            base_year = 9999
        if int(digits[:4]) > base_year:
            return 0
        return int(digits.ljust(8, "0"))

    news_items = sorted(
        [item for item in news_items if isinstance(item, dict)],
        key=lambda item: (_news_sort_key(item)[0], -_news_recency(item)),
    )
    news_html = "".join(
        f"<h3>{i}. {h(item.get('title'))}"
        + ('<span class="news-flag base-month">기준월 발생</span>'
           if str(item.get('source_date') or '').startswith(str(base_month)) else
           '<span class="news-flag ongoing">진행 중 사업</span>')
        + f"</h3><p><b>· 출처/발표일:</b> {h(item.get('source') or '-')} / {h(item.get('source_date') or '-')} · <b>추진단계:</b> {h(item.get('project_stage') or '-')}</p>"
        f"<p><b>· 내용 요약:</b> {h(item.get('summary'))}</p><p><b>· 지가 영향 및 시사점:</b> {h(item.get('impact'))}</p>"
        for i, item in enumerate(news_items, 1) if isinstance(item, dict)
    )
    if not news_html:
        news_html = issue_table

    zone_narratives = n.get("regional_zone_narratives") if isinstance(n.get("regional_zone_narratives"), list) else []
    zone_narrative_map = {
        norm_text(item.get("zone")): str(item.get("narrative") or "").strip()
        for item in zone_narratives if isinstance(item, dict)
    }
    # ② 권역별 동향은 반도체 직접 영향권 / 간접 영향권 2권역을 상위 구조로 제시한다.
    influence_summaries = build_influence_zone_summaries(result)
    influence_narratives = n.get("influence_zone_narratives")
    if not isinstance(influence_narratives, list) or not influence_narratives:
        influence_narratives = build_influence_zone_narratives(result)
    influence_map = {
        norm_text(item.get("zone")): str(item.get("narrative") or "").strip()
        for item in influence_narratives if isinstance(item, dict)
    }

    def _influence_card(item: dict[str, Any], index: int) -> str:
        report = item.get("당월_신고일기준", {}) or {}
        permit = item.get("허가구역_전후", {}) or {}
        top, low = item.get("지가_최고", {}) or {}, item.get("지가_최저", {}) or {}
        direct = str(item.get("구분", "")).startswith("반도체 직접")
        stat_line = (
            f"<b>당월 실거래(신고일)</b> {fmt_num(report.get('거래건수'), 0)}건 · "
            f"{fmt_num(report.get('거래금액_억원'), 1)}억(건수 {fmt_num(report.get('건수비중'), 1)}% · 금액 {fmt_num(report.get('금액비중'), 1)}%) · "
            f"외지인 {fmt_num(report.get('외지인거래건수'), 0)}건"
        )
        if permit:
            stat_line += (
                f" | <b>허가구역(계약일)</b> 지정 전 {fmt_num(permit.get('지정전_거래건수'), 0)}건 → "
                f"지정 후 {fmt_num(permit.get('지정후_거래건수'), 0)}건"
            )
        if top:
            stat_line += (
                f" | <b>당월 지가변동률</b> {h(top.get('읍면동'))} {fmt_num(top.get('당월변동률'), 3, '%')}"
                + (f" ~ {h(low.get('읍면동'))} {fmt_num(low.get('당월변동률'), 3, '%')}" if low and low.get("읍면동") != top.get("읍면동") else "")
            )
        return (
            f"<div class='regional-zone-card influence-zone-card {'direct' if direct else 'indirect'}'>"
            f"<div class='regional-zone-title'>{index}. {h(item.get('구분'))}</div>"
            f"<div class='regional-zone-dongs'>{h(item.get('대상'))}"
            + (f" · 구성 권역: {h(', '.join(item.get('구성권역', [])))}" if item.get("구성권역") else "")
            + "</div>"
            f"<div class='influence-stat'>{stat_line}</div>"
            f"<p>{h(limit_text(influence_map.get(norm_text(item.get('구분'))), SECTION_CHAR_LIMITS['influence_zone_narrative']))}</p>"
            f"<p class='note'><b>연계 산업·개발거점:</b> {h(', '.join(item.get('산업·개발거점', [])) or '-')}</p>"
            "</div>"
        )

    influence_zone_cards = "".join(
        _influence_card(item, index) for index, item in enumerate(influence_summaries, 1)
    )
    def _zone_dong_html(zone: dict[str, Any]) -> str:
        """세부 권역 카드의 동 목록에 반도체 직접 영향권 동을 표시한다."""
        direct_set = {norm_text(dong) for dong in semiconductor_dongs("direct")} if "광산구" in region else set()
        chips = []
        for dong in zone.get("dongs", []) or []:
            if norm_text(dong) in direct_set:
                chips.append(f"<b class='dong-direct'>{h(dong)}</b>")
            else:
                chips.append(h(dong))
        included = [dong for dong in zone.get("dongs", []) or [] if norm_text(dong) in direct_set]
        note = (
            f"<span class='dong-direct-note'>반도체 직접 영향권 {len(included)}개 동: {h('·'.join(included))}</span>"
            if included else ""
        )
        return "·".join(chips) + note

    regional_zone_cards = "".join(
        "<div class='regional-zone-card'"
        + (" data-direct='1'" if any(norm_text(x) in {norm_text(d) for d in semiconductor_dongs('direct')} for x in (zone.get('dongs') or [])) else "")
        + ">"
        f"<div class='regional-zone-title'>{i}. {h(zone.get('zone'))}</div>"
        f"<div class='regional-zone-dongs'>{_zone_dong_html(zone)}</div>"
        f"<p>{h(zone_narrative_map.get(norm_text(zone.get('zone'))) or zone.get('characteristics'))}</p>"
        f"<p class='note'><b>연계 산업거점:</b> {h(', '.join(zone.get('industrial_links', [])) or '-')} · <b>해석 유의:</b> {h(zone.get('analysis_caution'))}</p>"
        "</div>"
        for i, zone in enumerate(regional_profile.get("zones", []), 1)
    )
    industrial_complex_table = _reference_table(regional_profile.get("industrial_complexes", []), [
        ("name", "산업단지·개발지구", None),
        ("classification", "구분", None),
        ("linked_zone", "연계 권역", None),
        ("characteristics", "산업·입지 특성", None),
        ("market_points", "토지시장 점검사항", None),
    ], 0, "industrial-profile-table")
    industrial_profile_html = (
        f"<h3>③ 광산구 산업단지·연계 개발지구 특성</h3>"
        f"<p>{h(n.get('industrial_complex_analysis'))}</p>{industrial_complex_table}"
        f"<p class='note'>※ 기존 산업단지와 연계 개발지구를 구분함. {h(regional_profile.get('causality_rule'))}</p>"
        if regional_profile.get("available") else ""
    )
    issue_heading_number = "④" if regional_profile.get("available") else "③"
    development = result.get("development_projects", {}) or {}
    development_report_html = _development_report_html(
        development, n, region, month_title,
        (chr(ord(issue_heading_number) + 1), chr(ord(issue_heading_number) + 2)),
    )
    development_news_appendix_html = _development_news_appendix_html(development, n)
    section4_text = limit_text(
        n.get("external_buyer_analysis") or n.get("external_buyer_interpretation"), SECTION_CHAR_LIMITS["external_buyer_analysis"])
    section5_scale_text = limit_text(
        n.get("detailed_transaction_analysis") or n.get("transaction_interpretation"), SECTION_CHAR_LIMITS["detailed_transaction_analysis"])
    section5_zone_use_text = limit_text(
        n.get("zone_use_transaction_interpretation") or n.get("transaction_interpretation"), SECTION_CHAR_LIMITS["zone_use_transaction_interpretation"])
    section6_regional_text = limit_text(
        n.get("regional_analysis") or n.get("regional_issue_interpretation"), SECTION_CHAR_LIMITS["regional_analysis"])
    zone_reason_text = limit_text(
        n.get("zone_change_reasoning") or n.get("price_change_reasoning") or n.get("land_price_interpretation"), SECTION_CHAR_LIMITS["zone_change_reasoning"])
    use_reason_text = limit_text(
        n.get("use_change_reasoning") or n.get("price_change_reasoning") or n.get("land_price_interpretation"), SECTION_CHAR_LIMITS["use_change_reasoning"])

    chart_labels = [str(row.get("월", ""))[-5:] for row in monthly_rows]
    chart_amounts = [row.get("거래금액_억원", 0) or 0 for row in monthly_rows]
    chart_external = [row.get("외지인거래금액_억원", 0) or 0 for row in monthly_rows]
    chart_counts = [row.get("거래건수", 0) or 0 for row in monthly_rows]
    population_chart_rows = population_rows[-13:]
    population_chart_labels = [str(row.get("월", ""))[-5:] for row in population_chart_rows]
    population_chart_values = [row.get("인구수") for row in population_chart_rows]
    household_chart_values = [row.get("세대수") for row in population_chart_rows]
    valid_population_values = [float(v) for v in population_chart_values if v is not None]
    if valid_population_values:
        pop_min, pop_max = min(valid_population_values), max(valid_population_values)
        pop_pad = max((pop_max - pop_min) * 0.18, 120)
        population_axis_min = max(0, pop_min - pop_pad)
        population_axis_max = pop_max + pop_pad
    else:
        population_axis_min = population_axis_max = None
    public_rate_chart_labels = [str(row.get("월", ""))[-5:] for row in public_rate_rows]
    public_rate_chart_values = [row.get("평균") for row in public_rate_rows]
    rate_chart_rows = rates.get("table_rows", [])[-13:] if rates.get("available") else []
    rate_chart_labels = [str(row.get("month", ""))[-5:] for row in rate_chart_rows]
    mortgage_rate_values = [row.get("mortgage_rate") for row in rate_chart_rows]
    jeonse_rate_values = [row.get("jeonse_rate") for row in rate_chart_rows]
    sme_rate_values = [row.get("sme_rate") for row in rate_chart_rows]
    business_rate_values = [row.get("business_rate") for row in rate_chart_rows]
    migration_chart_rows = (mig.get("table_rows", []) or mig.get("monthly_rows", []))[-13:]
    migration_chart_labels = [str(row.get("month") or row.get("월") or "") for row in migration_chart_rows]
    migration_in_values = [row.get("in") for row in migration_chart_rows]
    migration_out_values = [row.get("out") for row in migration_chart_rows]
    migration_net_values = [row.get("net") if "net" in row else row.get("값") for row in migration_chart_rows]
    age_chart_rows = age_population.get("age_rows", []) if age_population.get("available") else []
    age_chart_labels = [row.get("연령대") for row in age_chart_rows]
    age_latest_values = [row.get("최신월 인구") for row in age_chart_rows]
    age_change_values = [row.get("13개월 증감") for row in age_chart_rows]
    age_chart_colors = ["#5BC0EBcc", "#3A86FFcc", "#8338ECcc", "#2A9D8Fcc", "#F4A261cc", "#E76F51cc", "#577590cc", "#D62828cc"][:len(age_chart_rows)]
    age_chart_borders = ["#5BC0EB", "#3A86FF", "#8338EC", "#2A9D8F", "#F4A261", "#E76F51", "#577590", "#D62828"][:len(age_chart_rows)]
    source_listing = "매물 엑셀 미지정(양식만 표시)" if listings.get("template_only") else "매물 엑셀"
    emd_months = tx.get("emd_months", []) if tx.get("available") else []
    emd_monthly_rows = tx.get("emd_monthly_rows", []) if tx.get("available") else []
    cluster_names = list(SEMICONDUCTOR_INFLUENCE_DONGS) if "광산구" in region else []
    emd_row_map = {norm_text(row.get("읍면동")): row for row in emd_monthly_rows}
    cluster_emd_rows = []
    for dong in cluster_names:
        row = emd_row_map.get(norm_text(dong))
        if row:
            cluster_emd_rows.append(row)
        else:
            cluster_emd_rows.append({
                "읍면동": dong,
                "_previous_count": 0,
                "_previous_amount_eok": 0.0,
                "월별": [{"월": month, "거래건수": 0, "거래금액_억원": 0.0} for month in emd_months],
                "합계건수": 0,
                "합계금액_억원": 0.0,
            })
    other_emd_rows = [row for row in emd_monthly_rows if norm_text(row.get("읍면동")) not in {norm_text(x) for x in cluster_names}]
    cluster_count_heatmap = _emd_timeseries_table(
        cluster_emd_rows, emd_months, region, "count", "클러스터 관련 읍·면·동", True, False
    )
    cluster_amount_heatmap = _emd_timeseries_table(
        cluster_emd_rows, emd_months, region, "amount", "클러스터 관련 읍·면·동", True, False
    )
    other_emd_count_heatmap = _emd_timeseries_table(
        other_emd_rows, emd_months, region, "count", "읍·면·동", False, False
    )
    other_emd_amount_heatmap = _emd_timeseries_table(
        other_emd_rows, emd_months, region, "amount", "읍·면·동", False, False
    )
    semiconductor_policy_html = _semiconductor_policy_card(region)
    influence_direct_text = "·".join(semiconductor_dongs("direct"))
    influence_linked_text = "·".join(semiconductor_dongs("linked"))
    # 계약일 기준 주간 집계와 토지거래허가구역 지정 전·후 집계(부록)
    permit_split = tx.get("permit_split", {}) if tx.get("available") else {}
    weekly = tx.get("weekly", {}) if tx.get("available") else {}
    permit_split_table = _permit_split_table(permit_split, region)
    period_trend = tx.get("period_trend", {}) if tx.get("available") else {}
    weekly_trend_table = _period_trend_table(period_trend)
    period_trend_rule = period_trend.get("rule") or ""
    weekly_group_table = _influence_group_weekly_table(weekly)
    emd_month_week = tx.get("emd_month_week", {}) if tx.get("available") else {}
    weekly_emd_table = _emd_month_week_table(emd_month_week, region)
    emd_month_week_rule = emd_month_week.get("rule") or ""
    survey_form_html = _field_survey_form_html(result, region, n)
    permit_rule = permit_split.get("region_rule", {}) if permit_split.get("available") else (land_permit_rule(region) or {})
    permit_periods = permit_split.get("periods", {}) if permit_split.get("available") else {}
    permit_pre_label = (permit_periods.get("pre", {}) or {}).get("label") or "-"
    permit_post_label = (permit_periods.get("post", {}) or {}).get("label") or "-"
    permit_note = permit_split.get("note") or permit_split.get("reason") or ""
    weekly_report_week = (weekly.get("report_week", {}) or {}).get("주차") or "-"
    weekly_as_of = weekly.get("as_of") or str(meta.get("as_of_date") or "-")
    date_basis_text = str(tx.get("date_basis") or meta.get("transaction_date_basis") or "계약일")
    monitoring_points = n.get("monitoring_points") if isinstance(n.get("monitoring_points"), list) else []
    monitoring_html = "".join(f"<li>{h(item)}</li>" for item in monitoring_points if str(item).strip())
    if not monitoring_html:
        monitoring_html = (
            "<li>지가변동률 상위 권역의 가격 변화와 실거래·지역시장동향을 종합하여 점검함.</li>"
            "<li>외지인 거래 비중과 고가·다필지 거래의 지속 여부를 확인함.</li>"
            "<li>전월·당월 매물 자료가 확보되면 매물 증감과 호가 조정을 함께 점검함.</li>"
        )

    return f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{h(month_title)} {h(region)} 토지 시장 동향 시황보고서</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{{--ink:#1a1f2b;--sub:#5b6472;--line:#e3e7ee;--bg:#f6f7f9;--navy:#0f2a52;--gold:#9a7b25}}
*{{box-sizing:border-box}}body{{font-family:'Pretendard','Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:var(--ink);line-height:1.75;margin:0;background:var(--bg);font-size:15px}}
.wrap{{max-width:1080px;margin:0 auto;background:#fff;padding:0 0 80px}}.cover{{background:linear-gradient(135deg,#0f2a52,#1d3f73);color:#fff;padding:54px 56px 46px}}
.cover .kicker{{letter-spacing:5px;font-size:12px;opacity:.8}}.cover h1{{font-size:32px;margin:14px 0 8px;line-height:1.3;font-weight:800}}.cover .meta{{opacity:.85;font-size:13.5px;border-top:1px solid rgba(255,255,255,.25);padding-top:14px;margin-top:18px}}
.cover .stat{{display:flex;gap:18px;margin-top:24px;flex-wrap:wrap}}.cover .stat div{{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.18);border-radius:12px;padding:14px 18px;min-width:180px}}.cover .stat b{{display:block;font-size:23px}}.cover .stat span{{font-size:12px;opacity:.82}}
section{{padding:30px 56px;border-bottom:1px solid var(--line)}}h2{{font-size:21px;color:var(--navy);border-left:6px solid var(--gold);padding-left:12px;margin:6px 0 18px}}h3{{font-size:16.5px;color:#243a5e;margin:24px 0 8px}}p{{margin:9px 0}}
.up{{color:#d4111a;font-weight:700}}.dn{{color:#1257d4;font-weight:700}}.g{{color:#8a93a3}}.kt{{background:#fbfaf4;border:1px solid #ece3c4;border-radius:14px;padding:18px 22px;margin:10px 0}}.kt li{{margin:7px 0}}
table{{border-collapse:collapse;width:100%;margin:12px 0;font-size:13.5px}}th,td{{border:1px solid var(--line);padding:7px 9px;text-align:center}}th{{background:#0f2a52;color:#fff;font-weight:600;font-size:12.5px}}.rk td:nth-child(2){{text-align:left}}tr.hot{{background:#fff7e6;font-weight:700}}tr.hot td{{color:#5a3d00}}td.blank-cell{{height:34px;background:#fff}}
.table-wrap{{overflow-x:auto;margin:10px 0 18px}}.time-series-table{{min-width:900px;margin:0}}.latest-row td{{background:#f2f7ff;font-weight:700}}.population-publish-note{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:9px 12px;margin:8px 0 10px;color:#334155;font-weight:700;font-size:12px}}.not-published{{color:#92400e;background:#fef3c7;border-radius:999px;padding:2px 7px;font-weight:900;font-size:12px}}.population-migration-table{{min-width:820px;table-layout:fixed}}.population-migration-table th,.population-migration-table td{{text-align:center;white-space:nowrap}}.age-pop-trend-table{{min-width:1180px;font-size:12px}}.age-rate{{display:block;font-size:10.5px;font-weight:600}}.age-population-structure-table td:first-child{{font-weight:700}}.public-rate-table{{min-width:1180px;font-size:12px}}.rates-standard-table{{min-width:980px;font-size:11.5px}}.rates-standard-table th,.rates-standard-table td{{text-align:center;white-space:nowrap}}.rates-standard-table .latest-row td{{background:#fff8bf!important;font-weight:900!important}}.rates-standard-table .up{{color:#dc2626!important}}.rates-standard-table .dn{{color:#2563eb!important}}.rates-standard-table .g{{color:#64748b!important}}.top-trade-table{{min-width:1500px;font-size:11.5px;margin:0}}.top-trade-table th,.top-trade-table td{{white-space:nowrap}}
.emd-name{{white-space:nowrap}}.influence-badge{{display:inline-block;margin-left:5px;padding:2px 7px;border-radius:999px;font-size:10.5px;font-weight:800;line-height:1.5;white-space:nowrap}}.influence-badge.direct,.influence-badge.core{{color:#b42318;background:#fee4e2;border:1px solid #fda29b}}.influence-badge.adjacent{{color:#9a6700;background:#fff3c4;border:1px solid #f5c451}}.influence-badge.linked{{color:#175cd3;background:#e8f1ff;border:1px solid #84adff}}
.permit-split-table th{{background:#15365f}}.permit-split-table tfoot .total-row td{{font-weight:900;background:#f7f2e4}}.permit-split-table td.delta-up,.weekly-trend-table td.delta-up{{color:#b91c1c;font-weight:800}}.permit-split-table td.delta-down,.weekly-trend-table td.delta-down{{color:#1d4ed8;font-weight:800}}.permit-split-table td.delta-flat,.weekly-trend-table td.delta-flat{{color:#64748b;font-weight:700}}
.weekly-trend-table td:first-child,.period-trend-table td:first-child{{font-weight:800}}
.period-trend-table td.delta-up{{color:#b91c1c;font-weight:800}}.period-trend-table td.delta-down{{color:#1d4ed8;font-weight:800}}.period-trend-table td.delta-flat{{color:#64748b;font-weight:700}}
.period-trend-table tr.row-week td{{background:#fffaf0}}.period-trend-table tr.row-total td{{background:#f7f2e4;font-weight:900;border-top:2px solid #9a7b25}}.period-trend-table .delta-note{{display:block;font-size:9.5px;font-weight:700;color:#64748b}}
.period-tag{{display:inline-block;min-width:16px;padding:1px 5px;margin-right:5px;border-radius:4px;font-size:10px;font-weight:900;color:#fff}}.period-tag.tag-month{{background:#15365f}}.period-tag.tag-week{{background:#8a3b00}}.period-tag.tag-total{{background:#9a7b25}}.weekly-emd-table{{min-width:1080px}}.weekly-emd-table .total-cell{{font-weight:800;background:#f7f2e4}}
.survey-form{{margin:12px 0}}.survey-form .form-meta{{display:flex;gap:16px;flex-wrap:wrap;font-size:12px;color:#334155;margin-bottom:10px}}
.survey-form .form-box{{border:1.5px solid #0f2a52;border-radius:6px;padding:14px 16px;margin:12px 0;background:#fff}}
.survey-form .form-title{{font-weight:900;color:#0f2a52;font-size:14.5px;margin-bottom:10px}}
.survey-form .form-item{{margin:12px 0}}.survey-form .form-bullet{{font-weight:800;color:#1d4ed8;font-size:13px;margin-bottom:4px}}
.survey-form .form-fact{{font-size:11.5px;color:#475569;background:#f8fafc;border-left:3px solid #94a3b8;padding:5px 9px;margin-bottom:5px;line-height:1.6}}
.survey-form .form-blank{{border:1px dashed #cbd5e1;border-radius:4px;min-height:34px;background:#fff}}.survey-form .form-blank.tall{{min-height:52px}}.survey-form .form-blank.xtall{{min-height:120px}}
.survey-form .survey-trade-table{{font-size:11px}}.survey-form .form-blank-cell{{background:#fcfcfd;min-width:110px}}
.survey-form .form-draft{{background:#f8fbff;border:1px solid #d9e4f5;border-left:4px solid #0f2a52;border-radius:6px;padding:9px 12px;margin-bottom:6px;font-size:12.5px;line-height:1.8;color:#1e293b}}
.survey-form .ext-mark{{color:#b42318}}.survey-form .form-trades{{font-size:11px;color:#475569;margin-bottom:6px}}.survey-form .form-trades ul{{margin:3px 0 0;padding-left:18px}}.survey-form .form-trades li{{line-height:1.6}}
.survey-form .demand-fact{{border-left-color:#c2410c;background:#fff7ed}}
.survey-form .monitor-note{{border:1px solid #cfe3d4;background:#f4fbf6;border-radius:6px;padding:8px 11px;margin:6px 0;font-size:11.5px;line-height:1.7}}.survey-form .monitor-title{{font-weight:900;color:#14532d;margin-bottom:4px}}.survey-form .monitor-src{{font-weight:600;color:#5f7d68;font-size:10.5px}}.survey-form .monitor-note ul{{margin:0;padding-left:18px}}
.survey-form .price-check{{border:1px solid #dbe4f0;border-radius:6px;padding:8px 10px;margin:6px 0}}.survey-form .price-check-title{{font-size:11.5px;font-weight:900;color:#0f2a52;margin-bottom:5px}}.survey-form .price-check-note{{font-weight:600;color:#8a93a3;font-size:10.5px}}
.survey-form .price-check-table{{font-size:11px;margin:0}}.survey-form .price-check-table td{{height:30px;color:#94a3b8}}
.survey-form .form-dongs{{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:6px}}.survey-form .dong-chip{{background:#fff;border:1px solid #cbd5e1;border-radius:999px;padding:2px 9px;font-size:11px;font-weight:800;color:#0f2a52}}.survey-form .dong-chip i{{font-style:normal;font-weight:600;color:#64748b;margin-left:4px}}
.weekly-emd-table .group-scope{{display:block;font-size:9.5px;font-weight:600;color:#64748b;white-space:normal}}
.month-week-table{{min-width:1560px}}.month-week-table th.group-month{{background:#15365f}}.month-week-table th.group-week{{background:#8a3b00}}.month-week-table th.group-ongoing{{background:#166534}}.month-week-table td.cell-ongoing{{background:#f2fbf5}}.month-week-table .ongoing-total{{background:#e8f6ed}}
.month-week-table td.cell-month{{background:#f8fafc}}.month-week-table td.cell-week{{background:#fffaf0}}.month-week-table td.cell-zero{{color:#94a3b8}}
.month-week-table .cell-count{{display:block;font-weight:900;line-height:1.15}}.month-week-table .cell-amount{{display:block;margin-top:2px;font-size:9.5px;font-weight:700;color:#64748b;line-height:1.1}}
.month-week-table .total-cell{{background:#f7f2e4}}.month-week-table .total-cell .cell-amount{{color:#8a6b1f}}
.policy-section{{padding-top:18px;padding-bottom:18px}}.policy-card{{background:#fffaf0;border:1px solid #e2b34b;border-radius:12px;padding:16px 18px;font-size:13px;line-height:1.65}}.policy-card p{{margin:7px 0}}.policy-title{{font-size:16px;font-weight:800;color:#8a3b00;margin-bottom:8px}}.influence-groups{{display:flex;gap:6px;flex-wrap:wrap;margin:8px 0}}.influence-groups .influence-badge{{margin-left:0}}.policy-caution{{background:#fff5eb;border-radius:8px;padding:8px 10px;color:#8a3b00}}.policy-source{{color:#5b6472;font-size:11.5px}}
.wide-table-wrap{{overflow-x:auto;margin:10px 0 18px}}.wide-table{{min-width:1420px;font-size:11.5px;margin:0}}.wide-table th,.wide-table td{{padding:6px 5px}}.wide-table .emd-col{{position:sticky;left:0;z-index:1;min-width:150px;text-align:left;background:#fff}}.wide-table th.emd-col{{z-index:2;background:#0f2a52}}.wide-table .total-cell{{font-weight:800;background:#f7f2e4}}
.heatmap-legend{{display:flex;gap:14px;justify-content:flex-end;align-items:center;flex-wrap:wrap;font-size:11px;color:#475569;margin:5px 2px 8px}}.heatmap-legend .legend-up{{color:#dc2626;font-weight:800}}.heatmap-legend .legend-down{{color:#2563eb;font-weight:800}}.heatmap-legend .legend-flat{{color:#64748b;font-weight:800}}.heatmap-legend .legend-latest{{color:#b45309;font-weight:800}}
.heatmap-table td.heat-up{{background:#fff1f2;color:#b91c1c}}.heatmap-table td.heat-down{{background:#eff6ff;color:#1d4ed8}}.heatmap-table td.heat-flat,.heatmap-table td.heat-na{{background:#f8fafc;color:#475569}}.heatmap-table td.heat-latest{{box-shadow:inset 0 0 0 1.5px #f59e0b}}.heatmap-table .heat-value{{display:block;font-weight:900;line-height:1.15}}.heatmap-table .heat-delta{{display:block;margin-top:2px;font-size:9.5px;font-weight:800;line-height:1.1;white-space:nowrap}}
.heatmap-table .emd-name{{display:block;white-space:normal;font-weight:800;line-height:1.25}}.heatmap-table .influence-badge{{margin:3px 3px 0 0;font-size:9px;white-space:normal}}
.heatmap-table{{table-layout:fixed}}.heatmap-table th:first-child,.heatmap-table td:first-child{{width:180px}}.appendix-heatmap .heatmap-table th:first-child,.appendix-heatmap .heatmap-table td:first-child{{width:125px}}
.sum{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:14px 0}}.card{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,0,0,.03)}}.card .t{{font-size:12.5px;color:var(--sub)}}.card .v{{font-size:22px;font-weight:800;margin:4px 0}}
.chartbox{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:18px;margin:14px 0;height:300px;position:relative}}.chart-box{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:18px 18px 22px;margin:14px 0;min-height:370px;overflow:visible}}.chart-box .chart-subtitle{{font-size:17px;font-weight:900;color:#1e293b;text-align:center;margin-bottom:10px}}.chart-box .chart-wrap{{height:300px;min-height:240px;padding-bottom:30px;position:relative;overflow:visible}}.chart-box .chart-note{{text-align:center;color:#64748b;font-weight:800;font-size:12px;margin-top:5px}}.population-household-chart .chart-wrap{{height:350px;padding-bottom:10px}}.migration-combo-chart{{padding-bottom:22px;min-height:390px}}.migration-combo-chart .chart-wrap{{height:330px;padding-bottom:28px}}.chart-grid-2{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}}.chart-grid-2 .chart-box{{min-width:0}}.rate-card{{border-top-width:4px}}.rate-card.mortgage{{border-top-color:#dc2626}}.rate-card.jeonse{{border-top-color:#2563eb}}.rate-card.sme{{border-top-color:#059669}}.rate-card.business{{border-top-color:#9333ea}}.reason-box{{border:1px solid #dbe4f0;border-left:5px solid #0f2a52;background:#f8fafc;border-radius:10px;padding:14px 17px;margin:10px 0 14px}}.reason-box.use-reason{{border-left-color:#9a7b25;background:#fcfbf6}}.reason-box .reason-label{{font-size:12px;font-weight:900;color:#475569;margin-bottom:5px}}.reason-box p{{margin:0;line-height:1.8}}.news{{background:#f9fbff;border:1px solid #d9e4f5;border-radius:12px;padding:18px 20px;margin:14px 0}}.news .ttl{{font-weight:800;color:#10325f;font-size:15px}}.zone{{border-left:4px solid var(--gold);padding:8px 14px;margin:12px 0;background:#fcfbf6}}.influence-zone-card{{border-top-width:5px}}.influence-zone-card.direct{{border-top-color:#b42318;background:linear-gradient(180deg,#fff5f5 0%,#fff 100%)}}.influence-zone-card.indirect{{border-top-color:#175cd3;background:linear-gradient(180deg,#f5f9ff 0%,#fff 100%)}}.influence-stat{{font-size:11.5px;color:#334155;background:#f8fafc;border-radius:6px;padding:6px 9px;margin:6px 0;line-height:1.7}}.sub-zone-heading{{font-size:13.5px;font-weight:800;color:#475569;margin:16px 0 6px}}.news-flag{{display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;font-size:10.5px;font-weight:800;vertical-align:middle}}.news-flag.base-month{{color:#b42318;background:#fee4e2;border:1px solid #fda29b}}.news-flag.ongoing{{color:#475569;background:#eef2f7;border:1px solid #cbd5e1}}.regional-zone-card[data-direct="1"]{{border-left:4px solid #b42318}}.regional-zone-dongs .dong-direct{{color:#b42318;font-weight:900}}.dong-direct-note{{display:block;margin-top:3px;font-size:10.5px;font-weight:800;color:#b42318}}.regional-zone-card{{border:1px solid #dbe4f0;border-top:4px solid #0f2a52;border-radius:12px;padding:15px 18px;margin:12px 0;background:linear-gradient(180deg,#f8fbff 0%,#fff 100%)}}.regional-zone-title{{font-weight:900;color:#0f2a52;font-size:16px}}.regional-zone-dongs{{font-size:12px;color:#64748b;margin:4px 0 8px}}.regional-zone-card p{{margin:6px 0;line-height:1.75}}.industrial-profile-table th{{background:#15365f}}.industrial-profile-table td:nth-child(1){{font-weight:800;color:#10325f}}.tag{{display:inline-block;background:#0f2a52;color:#fff;font-size:11px;padding:2px 9px;border-radius:20px;margin-left:6px}}.data-limit{{background:#f3f5f8;border:1px solid #d7dde6;border-radius:10px;padding:14px 18px;margin:16px 0}}.note{{font-size:12px;color:#8a93a3;margin-top:8px}}
@media(max-width:720px){{section,.cover{{padding:22px}}.sum,.chart-grid-2{{grid-template-columns:1fr}}.cover h1{{font-size:24px}}}}
@page{{size:A4 portrait;margin:10mm 8mm 11mm}}
@media print{{
html,body{{width:210mm;min-width:0;background:#fff!important}}
body{{margin:0;font-size:8.5pt;line-height:1.48;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
.wrap{{width:100%;max-width:none;margin:0;padding:0}}
.cover{{min-height:276mm;padding:22mm 12mm 16mm;break-after:page;page-break-after:always}}
.cover h1{{font-size:23pt}}.cover .stat{{display:grid;grid-template-columns:1fr 1fr;gap:4mm}}.cover .stat div{{min-width:0;padding:4mm}}
section{{padding:7mm 5mm;border-bottom:0;break-before:page;page-break-before:always;break-inside:auto}}
section.policy-section{{break-before:auto;page-break-before:auto}}
h2{{font-size:15pt;margin:0 0 4mm}}h3{{font-size:11pt;margin:5mm 0 2mm}}h2,h3{{break-after:avoid;page-break-after:avoid}}
p,li{{orphans:3;widows:3}}
.sum{{gap:2.5mm}}.card{{padding:3mm;border-radius:2mm;break-inside:avoid}}.card .v{{font-size:13pt}}
.chart-grid-2{{grid-template-columns:1fr 1fr;gap:3mm}}
.chartbox{{height:72mm;min-height:72mm;padding:3mm;margin:3mm 0 7mm;break-inside:avoid;page-break-inside:avoid}}
.chart-box{{height:82mm;min-height:82mm;padding:3mm;margin:3mm 0 7mm;break-inside:avoid;page-break-inside:avoid}}
.chart-box .chart-wrap{{height:68mm;min-height:68mm;padding-bottom:10mm}}
.population-household-chart,.migration-combo-chart{{height:88mm;min-height:88mm;margin-bottom:8mm}}
.population-household-chart .chart-wrap,.migration-combo-chart .chart-wrap{{height:73mm;min-height:73mm;padding-bottom:12mm}}
.kt,.news,.zone,.regional-zone-card,.reason-box,.data-limit,.policy-card{{break-inside:avoid;page-break-inside:avoid}}
.table-wrap,.wide-table-wrap{{width:100%;overflow:visible;margin:2.5mm 0 4mm}}
table,.time-series-table,.population-migration-table,.age-pop-trend-table,.public-rate-table,.rates-standard-table,.top-trade-table,.wide-table{{width:100%!important;min-width:0!important;max-width:100%!important;margin:2mm 0;font-size:6.2pt;line-height:1.25;table-layout:fixed}}
th,td,.wide-table th,.wide-table td{{padding:1.05mm .65mm;white-space:normal!important;overflow-wrap:anywhere;word-break:keep-all}}
th{{font-size:6pt}}thead{{display:table-header-group}}tfoot{{display:table-footer-group}}tr{{break-inside:avoid;page-break-inside:avoid}}
.top-trade-table{{font-size:5.35pt}}.top-trade-table th{{font-size:5.1pt}}.top-trade-table th,.top-trade-table td{{padding:.85mm .45mm;word-break:break-all}}
.wide-table{{font-size:5.4pt}}.wide-table th,.wide-table td{{padding:.8mm .4mm;word-break:break-all}}.wide-table .emd-col{{position:static;min-width:0;text-align:left}}
.heatmap-legend{{font-size:6.8pt;gap:3mm;margin:1mm 0 2mm}}.heatmap-table .heat-value{{font-size:5.4pt}}.heatmap-table .heat-delta{{font-size:4.8pt;margin-top:.4mm}}.heatmap-table td.heat-up,.heatmap-table td.heat-down,.heatmap-table td.heat-flat,.heatmap-table td.heat-na,.heatmap-table td.heat-latest{{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
.heatmap-table th:first-child,.heatmap-table td:first-child{{width:18%!important}}.heatmap-table th:not(:first-child),.heatmap-table td:not(:first-child){{width:auto!important}}.appendix-heatmap .wide-table-wrap{{break-inside:auto;page-break-inside:auto;overflow:visible}}.appendix-heatmap .heatmap-table{{font-size:5.1pt}}.appendix-heatmap .heatmap-table .heat-value{{font-size:5.1pt}}.appendix-heatmap .heatmap-table .heat-delta{{font-size:4.6pt}}.appendix-heatmap h3{{break-after:avoid;page-break-after:avoid}}
.age-pop-trend-table,.public-rate-table,.rates-standard-table{{font-size:5.7pt}}
.permit-split-table,.weekly-trend-table,.weekly-emd-table,.month-week-table,.survey-trade-table{{font-size:5.8pt}}.weekly-emd-table,.month-week-table{{min-width:0!important}}
.month-week-table .cell-amount{{font-size:4.8pt}}
.survey-form .form-box{{break-inside:avoid;page-break-inside:avoid}}.survey-form .form-draft{{font-size:7.2pt;line-height:1.55;padding:2mm 2.5mm}}.survey-form .form-trades{{font-size:6.4pt}}.survey-form .form-blank{{min-height:9mm}}.survey-form .form-blank.tall{{min-height:14mm}}.survey-form .form-blank.xtall{{min-height:32mm}}.survey-form .form-fact{{font-size:6.4pt}}
.industrial-profile-table{{font-size:5.9pt}}
.note{{font-size:7pt}}
a{{color:inherit;text-decoration:none}}
}}
</style></head><body><div class="wrap">
<div class="cover"><div class="kicker">{h(region.upper())} · LAND MARKET BRIEF</div><h1>{h(month_title)} {h(region)} 토지 시장 동향 시황보고서</h1>
<div>한국부동산원 광주지사 · 부동산시장조사 | 실거래 자료: {h(tx_data_type)} · 본문 기준: 신고일자 · 부록 기준: 계약일 / 계약묶음 단위 집계 · 주간 모니터링(매주 금요일)</div>
<div class="meta"><b>실거래 기준월:</b> {h(month_title)} 신고일 기준(부록은 계약일 기준) · <b>자료구분:</b> {h(tx_data_type)} · <b>주간 기준주:</b> {h(weekly_report_week)}(기준일 {h(weekly_as_of)}) · <b>지가 가격조사 기준월:</b> {h(month_label(price_survey_month))} · <b>최근 확정자료 확인월:</b> {h(month_label(latest_public_month))}</div><div class="stat">
<div><b>{fmt_num(current.get('거래건수'),0,'건')}</b><span>당월 전체 거래건수</span></div><div><b>{fmt_num(current.get('거래금액_억원'),1,'억')}</b><span>당월 전체 거래금액</span></div>
<div><b>{fmt_num(current.get('외지인거래금액_억원'),1,'억')}</b><span>외지인 거래금액(비중 {fmt_num(current.get('외지인금액비중'),1,'%')})</span></div><div><b>{fmt_num(lp_current.get('평균'),3,'%')}</b><span>{h(month_label(price_survey_month))} 지가변동률(누계 {fmt_num(lp_cum.get('평균'),3,'%')})</span></div></div>
<div class="meta">※ RTMS_KEY 마지막 필지순번을 제외한 계약묶음 단위로 집계함. 다필지 계약의 계약총금액은 1회 반영하고 토지면적은 필지별 합산함.</div></div>
<section><h2>[Key Takeaways] 당월 시장 핵심 요약</h2><div class="kt"><ol>{takeaways_html}</ol></div></section>
<section><h2>1. Executive Summary</h2><p>{h(n.get('executive_summary'))}</p></section>
<section><h2>2. 거시 경제 및 인구 동향</h2><h3>1) 예금은행 대출금리 추이 (신규취급액 기준·최근 13개월)</h3><div class="sum"><div class="card rate-card mortgage"><div class="t">주택담보대출</div><div class="v">{fmt_num(rates.get('mortgage_rate'),2,'%')}</div></div><div class="card rate-card jeonse"><div class="t">전세자금대출</div><div class="v">{fmt_num(rates.get('jeonse_rate'),2,'%')}</div></div><div class="card rate-card sme"><div class="t">중소기업대출</div><div class="v">{fmt_num(rates.get('sme_rate'),2,'%')}</div></div><div class="card rate-card business"><div class="t">기업일반자금대출</div><div class="v">{fmt_num(rates.get('business_rate'),2,'%')}</div></div></div><div class="chart-box rates-trend-chart"><div class="chart-subtitle">예금은행 대출금리 추세</div><div class="chart-wrap"><canvas id="ratesChart"></canvas></div></div>{rates_standard_table}<p>{h(n.get('rate_interpretation'))}</p><h3>2) {h(region)} 인구·세대수 추이 (최근 13개월)</h3><div class="chart-box mix-chart population-mix-chart population-household-chart"><div class="chart-subtitle">최근 13개월 인구수·세대수 추이(인구: 선, 세대: 막대, 이중축)</div><div class="chart-wrap"><canvas id="populationChart"></canvas></div><div class="chart-note">인구수는 변화 감지를 위해 좌축을 확대 표시하고, 세대수 막대는 실제 규모 왜곡을 줄이기 위해 우축 0 기준으로 표시함.</div></div><h3>3) 전입·전출·순이동 추이 (최근 13개월)</h3><div class="chart-box migration-combo-chart"><div class="chart-subtitle">최근 13개월 전입·전출·순이동 추이(전입·전출 막대, 순이동 선)</div><div class="chart-wrap"><canvas id="migrationChart"></canvas></div></div>{population_table}<h3>4) 연령대별 주민등록인구 차트</h3><div class="chart-grid-2"><div class="chart-box"><div class="chart-subtitle">최신월 연령대별 주민등록인구</div><div class="chart-wrap"><canvas id="agePopulationLatestChart"></canvas></div></div><div class="chart-box"><div class="chart-subtitle">최근 13개월 연령대별 증감</div><div class="chart-wrap"><canvas id="agePopulationChangeChart"></canvas></div></div></div><h3>5) 연령대별 주민등록인구 추이 (최근 13개월)</h3>{age_population_trend_table}<h3>6) 연령대별 인구 구조</h3>{age_population_structure_table}<p>{h(n.get('macro_population_interpretation') or n.get('population_household_migration_interpretation') or n.get('population_interpretation'))}</p></section>
<section><h2>3. 전체 토지 거래 통계 (기준: 신고일자, 실거래 묶음 단위 · {h(tx_data_type)})</h2><p class="note">※ {h(tx_data_type_note)}</p><div class="sum"><div class="card"><div class="t">당월 전체 거래</div><div class="v">{fmt_num(current.get('거래건수'),0,'건')} / {fmt_num(current.get('거래금액_억원'),1,'억')}</div><div>전월 건수 {_change_badge(change.get('거래건수증감률'))} · 금액 {_change_badge(change.get('거래금액증감률'))}</div></div><div class="card"><div class="t">전년동월 대비</div><div class="v">건수 {_change_badge(yoy.get('거래건수증감률'))}</div><div>금액 {_change_badge(yoy.get('거래금액증감률'))}</div></div><div class="card"><div class="t">당월 외지인 거래</div><div class="v">{fmt_num(current.get('외지인거래건수'),0,'건')} / {fmt_num(current.get('외지인거래금액_억원'),1,'억')}</div><div>금액비중 {fmt_num(current.get('외지인금액비중'),1,'%')}</div></div></div><div class="chartbox"><canvas id="trendChart"></canvas></div><h3>① 최근 13개월 토지 거래 및 외지인 거래 추이</h3>{monthly_table}<p class="note">{public_rate_availability_note}</p><h3>② 읍·면·동별 전체 거래 집계 (금액 비중 내림차순)</h3>{emd_table}<h3>③ 반도체 클러스터 관련 읍·면·동 최근 13개월 거래건수·증감 히트맵</h3><p class="note">반도체 직접 영향권 {h(influence_direct_text)}, 연계 영향권 {h(influence_linked_text)}을 개별 읍·면·동 기준으로 표시함. 각 셀은 거래건수와 직전월 대비 증감 건수를 함께 표시하며 첫 표시월도 내부 직전월과 비교함.</p>{cluster_count_heatmap}<h3>④ 반도체 클러스터 관련 읍·면·동 최근 13개월 거래금액·증감 히트맵</h3><p class="note">단위: 억 원. ③과 동일한 {len(SEMICONDUCTOR_INFLUENCE_DONGS)}개 읍·면·동을 표시하며, 각 셀은 거래금액과 직전월 대비 증감액을 함께 표시함.</p>{cluster_amount_heatmap}<h3>⑤ 용도지역별 거래 집계 (전수)</h3>{zone_table}<h3>⑥ 이용상황별 거래 집계 (전수)</h3>{use_table}<p class="note">※ 반도체 클러스터 관련 지역을 제외한 개별 읍·면·동 13개월 히트맵은 별첨에 제시함. 외지인 매수자 지역·이용상황·고가거래 상세는 4장에서 통합 제시함. 2026년 7월부터의 토지거래허가구역 지정 전·후 읍·면·동 집계와 주간 집계는 부록 3·4에 제시함.</p></section>
<section><h2>4. 외지인 거래 상세 분석</h2><p>{h(section4_text)}</p><h3>① 외지인 매수자 시군구별 집계</h3>{buyer_table}<h3>② 외지인 거래 이용상황별 집계</h3>{external_use_table}<h3>③ 외지인 고가 실거래 TOP10 심층 분석</h3><p class="note">※ 당월 외지인 거래를 RTMS_KEY 마지막 필지순번 제외 계약묶음의 거래총액이 큰 순서로 표시함. 다필지 계약금액은 1회 반영하고 면적은 합산하며, 단가는 단일 필지·토지만 거래인 경우에만 산출하여 천원 미만 절사함.</p>{external_top_trade_table}</section>
<section><h2>5. 실거래 상세 분석</h2><h3>① 거래 규모 및 자본 흐름 종합</h3><p>{h(section5_scale_text)}</p><h3>② 주요 용도지역 및 이용상황별 거래 동향</h3><p>{h(section5_zone_use_text)}</p><h3>③ 고가 실거래 TOP10 심층 분석</h3><p class="note">※ 다수 필지는 RTMS_KEY 마지막 필지순번을 제외한 계약묶음 기준으로 대표 주소 뒤에 외(추가 필지 수)로 표시함. 계약총금액은 묶음당 1회 반영하고 면적은 필지별 합산함. 단가는 단일 필지이면서 토지만 거래한 경우에만 산출하여 천원 미만 절사함.</p>{top_trade_table}</section>
<section><h2>6. 시장동향 종합 및 권역별 특화 분석</h2><h3>① 종합 시장동향</h3><p>{h(n.get('market_overview') or n.get('conclusion'))}</p><h3>② 권역별 동향 (반도체 영향권 구분)</h3><div class="zone"><p>{h(section6_regional_text)}</p></div>{influence_zone_cards}<p class="note">※ 영향권 구분은 군공항 예정지 인접성에 따른 본 보고서용 분석 분류이며, 직접 영향권 표기가 개별 읍·면·동의 수혜나 가격 상승을 의미하지 않음. 실거래 건수·금액은 신고일 기준, 허가구역 전후 비교는 계약일 기준임.</p><h4 class="sub-zone-heading">참고: 생활권·산업권 세부 권역</h4>{regional_zone_cards}{industrial_profile_html}{semiconductor_policy_html}<h3>{issue_heading_number} {h(region)} 부동산 관련 주요 이슈</h3>{news_cards_html or issue_table}{development_report_html}</section>
<section><h2>7. 시군구 지가변동률 및 가격 변동 사유</h2><h3>① 시군구 지가변동률 (최근 13개월)</h3><div class="chartbox"><canvas id="publicRateChart"></canvas></div>{public_rate_table}<p class="note">※ 해당 지역 조사담당자가 조사대상 토지의 가격 변화, 실거래 및 지역시장동향을 종합하여 작성한 지가변동률임.</p><h3>② 용도지역별 가격 변동 및 변동사유</h3><div class="reason-box zone-reason"><div class="reason-label">용도지역별 변동사유 · 실거래 중심 300자 이내</div><p>{h(zone_reason_text)}</p></div>{lp_zone_table}<h3>③ 이용상황별 가격 변동 및 변동사유</h3><div class="reason-box use-reason"><div class="reason-label">이용상황(지목)별 변동사유 · 실거래 중심 300자 이내</div><p>{h(use_reason_text)}</p></div>{lp_use_table}</section>
<section><h2>8. 변동률 판단 근거 종합 (가격 변화·실거래·지역시장동향)</h2><p class="note">※ 조사담당자는 조사대상 토지의 가격 변화와 실거래, 개발사업 추진상황, 산업·교통·상권 등 지역시장동향을 종합하여 변동 배경을 판단함.</p>{judgment_table}</section>
<section><h2>9. 읍·면·동별 상승지역 변동 사유 (지가 변동률 기준)</h2><p>{h(section9_lead_text)}</p>{emd_reason_cards}{lp_emd_table}<p class="note">※ 당월 지가변동률 상위 5개 지역의 지역 특성과 실거래·매물 자료를 종합하여 변동 사유를 서술함.</p></section>
<section><h2>10. 종합 결론 및 모니터링 시사점</h2><p>{h(n.get('conclusion') or n.get('market_overview'))}</p><h3>향후 모니터링 항목</h3><ul>{monitoring_html}</ul><div class="data-limit"><b>■ 데이터 한계 및 해석 시 유의사항</b><ul><li>본문 3~5장 실거래 통계는 신고일자 기준 계약묶음 단위로 집계함. 자료구분은 {h(tx_data_type)}이며 {h(tx_data_type_note)}.</li><li>부록 3~5(토지거래허가구역 전후·주간 모니터링·현장점검 양식)는 규제 시점과 계약 시점을 맞추기 위해 계약일 기준으로 집계하므로 본문 수치와 다를 수 있음.</li><li>계약일 기준 집계는 신고기한(계약일부터 30일) 내 미신고 계약이 이후 갱신 시 추가되므로 최근 구간 수치는 변동될 수 있음.</li><li>지가변동률은 조사담당자가 조사대상 토지의 가격 변화, 실거래 및 지역시장동향을 종합해 판단한 결과임.</li><li>거래가 희소한 권역은 가격 변화와 지역 여건을 중심으로 판단하며 개별 거래를 일반화하지 않음.</li><li>매물 자료가 전월·당월 모두 확보되지 않은 경우 증감·호가 판단을 제한함.</li><li>토지거래허가구역 효력 발생일은 2026년 7월 14일이며, 규제 전후 비교는 계약일 기준 부록 3의 구간 집계(지정 전 {h(permit_pre_label)} · 지정 후 {h(permit_post_label)})로 확인함.</li><li>본 분석은 가격 변화와 지역시장동향 판단을 지원하기 위한 참고자료이며, 최종 변동률 판단과 보고서 확정은 해당 지역 조사담당자의 검토를 거쳐 수행됨.</li></ul></div></section>
<section class="appendix-heatmap"><h2>[별첨] 기타 읍·면·동별 최근 13개월 거래 히트맵</h2><p class="note">※ 본문에서 별도 제시한 반도체 클러스터 관련 7개 지역을 제외하고, 나머지 지역은 기초단위 권역으로 묶지 않은 개별 읍·면·동 기준임. 최근 13개월 합계 거래금액이 큰 순서로 배열함.</p><h3>① 읍·면·동별 거래건수 및 전월 대비 증감</h3>{other_emd_count_heatmap}<h3>② 읍·면·동별 거래금액 및 전월 대비 증감</h3><p class="note">단위: 억 원. 소수점 표시 과정에서 전체 합계와 미세한 반올림 차이가 발생할 수 있음.</p>{other_emd_amount_heatmap}</section>
<section><h2>[부록 1] 매물 분석 ({h(listings.get('source') or '네이버 매물조회')})</h2><p>{h(n.get('listing_interpretation')) if not listings.get('template_only') else '매물 엑셀을 지정하지 않아 수치 없이 기준 양식만 표시함.'}</p>{f'<div class="influence-stat">{h(listing_headline)}</div>' if listing_headline else ''}<h3>① 이용상황별 매물 (전월 대비)</h3>{listing_file_table}<h3>② 읍·면·동별 매물 (전월 대비)</h3>{listing_blank_emd}<h3>③ 용도지역별 매물 (전월 대비)</h3>{listing_blank_change}<h3>④ 동일 매물 호가 조정 현황</h3><p class="note">※ 매물 고유번호가 전월·당월에 모두 있는 매물만 비교하며, 호가 변동 상위 사례를 함께 표시함. 신규 등록·매물 회수는 고유번호 기준 증감임.</p>{listing_blank_price}</section>
<section><h2>[부록 2] {h(region)} 부동산 관련 주요 뉴스 상세 심층 분석 ({h(month_title)} 기준)</h2><p class="note">※ 기준월({h(month_title)}) 중 발생한 이슈를 먼저 배치하고, 이전에 지정·착수돼 진행 중인 사업은 기준월 시장에 미친 영향 관점으로 정리함. 발표일이 확인되지 않은 사업은 추진단계로 표기함.</p>{news_html}{development_news_appendix_html}</section>
<section class="appendix-permit"><h2>[부록 3] 토지거래허가구역 지정 전·후 읍·면·동 거래 현황 (계약일 기준)</h2>
<p class="note">※ {h(permit_rule.get('announcement') or '국토교통부 공고')} 기준 효력 발생일 {h(permit_rule.get('effective_date') or '2026-07-14')}, 지정기간 {h(permit_rule.get('designation_period') or '-')}, 대상 {h(permit_rule.get('scope') or '-')}. 2026년 7월 이전은 월단위로 비교하고, 2026년 7월부터는 계약일 기준으로 지정 전({h(permit_pre_label)})·지정 후({h(permit_post_label)}) 구간을 나눠 집계함.</p>
{permit_split_table}
<p class="note">※ {h(permit_note)} 계약일 기준 집계이므로 신고 지연분이 이후 갱신 시 반영될 수 있으며, 허가구역 지정은 규제조치로서 거래 감소가 곧 가격 하락을 의미하지 않음.</p></section>
<section class="appendix-weekly"><h2>[부록 4] 주간 토지시장 모니터링 (계약일 기준·매주 금요일 작성)</h2>
<p class="note">※ 주간 구간은 {h(weekly.get('week_rule') or '토요일~금요일 7일 창')}이며 작성 기준일은 {h(weekly_as_of)}, 기준주는 {h(weekly_report_week)}임. 실거래 엑셀을 매주 갱신한 뒤 재실행하면 아래 표가 최신 주차까지 자동 갱신됨.</p>
<h3>① 전체 거래 추이 (기준월 이전 월단위 + 기준월 주단위)</h3>
<p class="note">{h(period_trend_rule)}. 기준월이 넘어가면 직전 기준월은 월단위 행으로 자동 이동함. '전기 대비'는 바로 위 행과의 거래건수 차이이며, 월↔주 전환 지점은 기간 길이가 달라 증감폭을 그대로 비교하지 않아야 함.</p>{weekly_trend_table}
<h3>② 영향권 구분별 주간 거래건수</h3>
<p class="note">직접 영향권 {h(influence_direct_text)} 합계와 간접 영향권(직접 영향권 외 읍·면·동) 합계를 주차별로 비교함. 현장점검 보고서(부록 5)의 서술 단위와 동일함.</p>{weekly_group_table}
<h3>③ 읍·면·동별 거래건수·거래금액 (기준월 이전 월단위 + 기준월 주단위)</h3>
<p class="note">{h(emd_month_week_rule)}. 기준월이 넘어가면 직전 기준월은 월단위 칸으로 자동 이동함. 각 칸은 위쪽 거래건수(건), 아래쪽 거래금액(억원)이며 반도체 직접 영향권 {h(influence_direct_text)}을 상단에 고정 표시하고, 그 외 지역은 해당 기간 거래가 있는 읍·면·동만 표시함. 금액은 칸별로 반올림해 표시하므로 세로 합계와 미세한 차이가 발생할 수 있음.</p>{weekly_emd_table}
<p class="note">※ {h(weekly.get('note') or '')}</p></section>
<section class="appendix-form"><h2>[부록 5] 주간 현장점검 보고서 작성 양식</h2>
<p class="note">※ 광주지사 현장점검 보고서 양식을 주간 작성용으로 옮겨 실은 것으로, 계약일 기준 자동 집계값을 근거로 붙이고 서술란은 조사담당자가 직접 기입함.</p>
{survey_form_html}</section>
<section style="border-bottom:none"><p class="note">본 보고서는 실거래 매매자료, 지가변동률 자료, 인구·세대·이동·금리 보조자료 및 {h(source_listing)}을 기초로 해당 지역 조사담당자가 가격 변화와 지역시장동향을 종합하여 작성함.</p></section>
</div><script>
new Chart(document.getElementById('trendChart'),{{type:'bar',data:{{labels:{json.dumps(chart_labels,ensure_ascii=False)},datasets:[{{type:'bar',label:'전체 거래금액(억)',data:{json.dumps(chart_amounts)},backgroundColor:'rgba(15,42,82,.75)',yAxisID:'y',order:2}},{{type:'bar',label:'외지인 거래금액(억)',data:{json.dumps(chart_external)},backgroundColor:'rgba(212,17,26,.65)',yAxisID:'y',order:2}},{{type:'line',label:'전체 거래건수',data:{json.dumps(chart_counts)},borderColor:'#9a7b25',backgroundColor:'#9a7b25',yAxisID:'y1',tension:.3,order:1}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{title:{{display:true,text:'최근 13개월 거래금액·외지인·건수 추이'}},legend:{{position:'bottom'}}}},scales:{{y:{{position:'left',title:{{display:true,text:'금액(억 원)'}}}},y1:{{position:'right',grid:{{drawOnChartArea:false}},title:{{display:true,text:'건수'}}}}}}}}}});
new Chart(document.getElementById('ratesChart'),{{type:'line',data:{{labels:{json.dumps(rate_chart_labels,ensure_ascii=False)},datasets:[{{label:'주택담보대출',data:{json.dumps(mortgage_rate_values)},borderColor:'#dc2626',backgroundColor:'#dc262633',borderWidth:2.5,pointRadius:3,tension:.25,fill:false}},{{label:'전세자금대출',data:{json.dumps(jeonse_rate_values)},borderColor:'#2563eb',backgroundColor:'#2563eb33',borderWidth:2.5,pointRadius:3,tension:.25,fill:false}},{{label:'중소기업대출',data:{json.dumps(sme_rate_values)},borderColor:'#059669',backgroundColor:'#05966933',borderWidth:2.5,pointRadius:3,tension:.25,fill:false}},{{label:'기업일반자금대출',data:{json.dumps(business_rate_values)},borderColor:'#9333ea',backgroundColor:'#9333ea33',borderWidth:2.5,pointRadius:3,tension:.25,fill:false}}]}},options:{{responsive:true,maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},plugins:{{title:{{display:false}},legend:{{position:'bottom',labels:{{boxWidth:12,padding:12}}}}}},scales:{{x:{{grid:{{color:'rgba(148,163,184,.16)'}}}},y:{{title:{{display:true,text:'금리(%)'}},grid:{{color:'rgba(148,163,184,.18)'}}}}}}}}}});
new Chart(document.getElementById('populationChart'),{{type:'bar',data:{{labels:{json.dumps(population_chart_labels,ensure_ascii=False)},datasets:[{{type:'line',label:'총인구수',data:{json.dumps(population_chart_values)},borderColor:'#475569',backgroundColor:'#47556933',pointBackgroundColor:'#475569',pointBorderColor:'#475569',yAxisID:'y',order:1,borderWidth:3.2,pointRadius:5,pointHoverRadius:6,tension:.25,fill:false}},{{type:'bar',label:'총세대수',data:{json.dumps(household_chart_values)},backgroundColor:'#7ee7bdcc',borderColor:'#10b981',borderWidth:1.4,borderRadius:4,barPercentage:.68,categoryPercentage:.72,yAxisID:'y1',order:2}}]}},options:{{responsive:true,maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},layout:{{padding:{{bottom:12,top:4}}}},plugins:{{title:{{display:false}},legend:{{position:'bottom',labels:{{boxWidth:12,padding:12}}}}}},scales:{{x:{{grid:{{color:'rgba(148,163,184,.16)'}}}},y:{{position:'left',min:{json.dumps(population_axis_min)},max:{json.dumps(population_axis_max)},title:{{display:true,text:'인구수(명)'}},grid:{{color:'rgba(148,163,184,.18)'}}}},y1:{{position:'right',beginAtZero:true,grid:{{drawOnChartArea:false}},title:{{display:true,text:'세대수(세대)'}}}}}}}}}});
new Chart(document.getElementById('migrationChart'),{{type:'bar',data:{{labels:{json.dumps(migration_chart_labels,ensure_ascii=False)},datasets:[{{type:'bar',label:'총 전입',data:{json.dumps(migration_in_values)},yAxisID:'y',backgroundColor:'#ef4444cc',borderColor:'#dc2626',borderWidth:1.2,borderRadius:4}},{{type:'bar',label:'총 전출',data:{json.dumps(migration_out_values)},yAxisID:'y',backgroundColor:'#3b82f6cc',borderColor:'#2563eb',borderWidth:1.2,borderRadius:4}},{{type:'line',label:'순이동',data:{json.dumps(migration_net_values)},yAxisID:'y',borderColor:'#111827',backgroundColor:'#11182733',pointBackgroundColor:'#111827',borderWidth:3,pointRadius:3.5,tension:.25}}]}},options:{{responsive:true,maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},layout:{{padding:{{bottom:26,top:4}}}},plugins:{{title:{{display:false}},legend:{{position:'bottom',labels:{{boxWidth:11,padding:16}}}}}},scales:{{x:{{grid:{{color:'rgba(148,163,184,.16)'}}}},y:{{beginAtZero:false,grid:{{color:'rgba(148,163,184,.18)'}}}}}}}}}});
new Chart(document.getElementById('agePopulationLatestChart'),{{type:'bar',data:{{labels:{json.dumps(age_chart_labels,ensure_ascii=False)},datasets:[{{label:'최신월 인구',data:{json.dumps(age_latest_values)},backgroundColor:{json.dumps(age_chart_colors)},borderColor:{json.dumps(age_chart_borders)},borderWidth:1.2,borderRadius:4}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom'}}}},scales:{{y:{{beginAtZero:true,title:{{display:true,text:'인구(명)'}}}}}}}}}});
new Chart(document.getElementById('agePopulationChangeChart'),{{type:'bar',data:{{labels:{json.dumps(age_chart_labels,ensure_ascii=False)},datasets:[{{label:'13개월 증감',data:{json.dumps(age_change_values)},backgroundColor:{json.dumps(age_chart_colors)},borderColor:{json.dumps(age_chart_borders)},borderWidth:1.2,borderRadius:4}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom'}}}},scales:{{y:{{title:{{display:true,text:'증감(명)'}},grid:{{color:'rgba(148,163,184,.18)'}}}}}}}}}});
new Chart(document.getElementById('publicRateChart'),{{type:'line',data:{{labels:{json.dumps(public_rate_chart_labels,ensure_ascii=False)},datasets:[{{label:'시군구 평균 지가변동률',data:{json.dumps(public_rate_chart_values)},borderColor:'#d4111a',backgroundColor:'rgba(212,17,26,.12)',pointBackgroundColor:'#d4111a',borderWidth:3,tension:.25,fill:true}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{title:{{display:true,text:'시군구 지가변동률 최근 13개월'}},legend:{{position:'bottom'}}}},scales:{{y:{{title:{{display:true,text:'변동률(%)'}}}}}}}}}});
</script></body></html>"""


def write_outputs(result: dict[str, Any], output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    region = safe_name(result["meta"]["region"])
    month = result["meta"]["base_month"].replace("-", "")
    stem = f"land_price_monthly_{region}_{month}"
    result_path = output_dir / f"{stem}_result.json"
    prompt_path = output_dir / f"llm_prompt_{region}_지가변동률_{month}.txt"
    html_path = output_dir / f"{region}_토지시장동향_시황보고서_{month}.html"
    result_path.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    prompt_path.write_text(build_llm_prompt(result), encoding="utf-8")
    html_path.write_text(build_html_report(result), encoding="utf-8")
    return {"result_json": result_path, "llm_prompt": prompt_path, "html_report": html_path}


def _sentence_end_positions(text: str) -> list[int]:
    """문장 종결 마침표의 위치를 돌려준다. 숫자 사이 소수점(예: 30.9, 14.0%)은 제외한다."""
    positions: list[int] = []
    for match in re.finditer(r"\.", text):
        i = match.start()
        prev_char = text[i - 1] if i > 0 else ""
        next_char = text[i + 1] if i + 1 < len(text) else " "
        if prev_char.isdigit() and next_char.isdigit():
            continue  # 소수점은 문장 경계가 아니다
        positions.append(i)
    return positions


def trim_reason_text(text: Any, limit: int = 230) -> str:
    """변동사유(200자 내외)를 문장 경계에서 안전하게 줄인다. 숫자·어절 중간에서 끊지 않는다.

    한도는 '200자 내외'를 넉넉히 보아 230자로 두어, 전형적인 2~3문장 사유는 통째로 보존하고
    지나치게 긴 출력만 마지막 완결 문장까지 줄인다.
    """
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(text) <= limit:
        return text
    ends = _sentence_end_positions(text)
    if ends:
        chosen = None
        for end in ends:
            if end + 1 <= limit:
                chosen = end
            else:
                break
        # limit 안에 완결 문장이 없으면 첫 문장 하나는 살려 숫자 중간 절단을 막는다.
        target = chosen if chosen is not None else ends[0]
        return text[: target + 1].strip()
    cut = text.rfind(" ", 0, limit)
    return text[: cut if cut > 0 else limit].rstrip(" ,·") + "…"


# 장별 서술 분량 한도(공백 포함 글자수). 5장은 ①+② 합계 500자, 6장 ②는 총평+2개 영향권 합계 1,500자.
SECTION_CHAR_LIMITS = {
    "external_buyer_analysis": 500,
    "detailed_transaction_analysis": 300,
    "zone_use_transaction_interpretation": 200,
    "regional_analysis": 300,
    "influence_zone_narrative": 600,
    "zone_change_reasoning": 300,
    "use_change_reasoning": 300,
    "emd_area_narrative": 200,
    "development_news_summary": 100,
    "monthly_news_summary": 500,
}


# 9장에서 빼는 문장: 확인 요청·판단 기준류
SECTION9_BANNED = re.compile(r"조사\s*담당자|확인이 필요|확인할 필요|확인 필요|현장 확인|현장조사|판단구분|가격조사 중심|단정하기 어려")


def drop_sentences(text: Any, pattern: re.Pattern) -> str:
    """pattern이 들어간 문장을 통째로 뺀다."""
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    kept: list[str] = []
    start = 0
    for end in _sentence_end_positions(text) + [len(text) - 1]:
        if end < start:
            continue
        sentence = text[start:end + 1].strip()
        start = end + 1
        if sentence and not pattern.search(sentence):
            kept.append(sentence)
    return " ".join(kept)


def limit_text(text: Any, limit: int) -> str:
    """한도 안에서 마지막 완결 문장까지만 남긴다(첫 문장이 한도를 넘으면 어절 경계에서 자름)."""
    trimmed = trim_reason_text(text, limit)
    return trimmed if len(trimmed) <= limit else fit_text(trimmed, limit)


def _flatten_strings(obj: Any):
    """중첩 구조에서 문자열만 순회 산출한다."""
    if isinstance(obj, dict):
        for value in obj.values():
            yield from _flatten_strings(value)
    elif isinstance(obj, list):
        for value in obj:
            yield from _flatten_strings(value)
    elif isinstance(obj, str):
        yield obj


def lint_narratives(narratives: dict[str, Any]) -> list[str]:
    """LLM 서술의 구조·문체·아티팩트 결함을 점검해 경고 문자열 목록을 돌려준다."""
    issues: list[str] = []
    required = [
        "executive_summary", "transaction_interpretation", "land_price_interpretation",
        "zone_change_reasoning", "use_change_reasoning", "market_overview", "conclusion",
    ]
    for key in required:
        value = narratives.get(key)
        if not (isinstance(value, str) and value.strip()):
            issues.append(f"[빈값] 필수 서술 '{key}'가 비어 있음")

    def walk(path: str, value: Any) -> None:
        if isinstance(value, dict):
            for key, item in value.items():
                walk(f"{path}.{key}" if path else key, item)
        elif isinstance(value, list):
            for idx, item in enumerate(value):
                walk(f"{path}[{idx}]", item)
        elif isinstance(value, str) and value.strip():
            text = value
            if re.search(r"직접적?인?\s*원인으로\s*직접적?인?\s*원인으로", text):
                issues.append(f"[원인어구중복] {path}")
            for match in re.finditer(r"([가-힣]{2,6})\s+\1(?=[\s.,])", text):
                issues.append(f"[어절반복] {path}: '{match.group(0)}'")
            if re.search(r"[0-9]\.$", text.rstrip()):
                issues.append(f"[숫자중간종료] {path}: …{text.rstrip()[-24:]}")
            if "{display_area}" in text or "{issue_scope_note}" in text:
                issues.append(f"[자리표시자잔존] {path}")
            for art in ["가격 변화은", "가격 변화이", "광산구의 광산구", "  "]:
                if art in text:
                    issues.append(f"[아티팩트:{art.strip() or '이중공백'}] {path}")

    walk("", narratives)

    style_fields = [
        "executive_summary", "transaction_interpretation", "land_price_interpretation",
        "market_overview", "regional_analysis", "conclusion", "price_change_reasoning",
        "zone_change_reasoning", "use_change_reasoning",
    ]
    for key in style_fields:
        value = narratives.get(key)
        if isinstance(value, str) and value.strip():
            tail = value.rstrip()
            if not re.search(r"(됨|함|음|임)\.?$", tail):
                issues.append(f"[문체] '{key}' 종결어미 확인 필요: …{tail[-16:]}")
    return issues


def audit_narrative_numbers(result: dict[str, Any], narratives: dict[str, Any]) -> list[str]:
    """서술이 인용한 지가변동률(평균·누계·용도·이용상황) 수치를 result.json과 대조한다.

    라벨이 붙은 고정 수치만 검사해 오탐을 최소화한다(연성 경고).
    """
    warns: list[str] = []
    lp = (result.get("land_price", {}) or {})
    current = lp.get("current", {}) or {}
    text = " ".join(_flatten_strings(narratives))

    def near(a: Any, b: Any, tol: float = 0.0011) -> bool:
        n_a, n_b = as_number(a), as_number(b)
        return n_a is not None and n_b is not None and abs(n_a - n_b) <= tol

    # 시군구·용도·이용·읍면동·판단 등 모든 정당한 지가변동률 값 집합
    rate_values: set[float] = set()

    def add_rate(value: Any) -> None:
        number = as_number(value)
        if number is not None:
            rate_values.add(round(number, 3))

    for source in [current, lp.get("previous", {}), lp.get("change_vs_previous", {}),
                   lp.get("cumulative", {}), lp.get("public_current", {})]:
        for value in (source or {}).values():
            add_rate(value)
    for row in (lp.get("emd_top", []) or []):
        if isinstance(row, dict):
            for value in row.values():
                add_rate(value)
    for row in ((result.get("judgment", {}) or {}).get("top_areas", []) or []):
        if isinstance(row, dict):
            add_rate(row.get("current_rate"))
            add_rate(row.get("cumulative_rate"))

    # 매물 호가·거래건수 문맥의 수치는 지가변동률이 아니므로 문장 단위로 제외한다.
    listing_words = ["호가", "매물", "단가", "등록", "회수", "건에서", "건으로", "거래건수",
                     "상향", "하향", "조정 사례", "사례는", "사례가", "억 원", "억원", "만원"]

    # 소수점을 문장 경계로 오인하지 않도록 종결 마침표 위치만 사용한다.
    boundaries = _sentence_end_positions(text)

    def sentence_at(position: int) -> str:
        start = max([0] + [b + 1 for b in boundaries if b < position])
        after = [b for b in boundaries if b >= position]
        end = after[0] if after else len(text)
        return text[start:end]

    def in_listing_context(position: int) -> bool:
        return any(word in sentence_at(position) for word in listing_words)

    def is_rate_context(position: int) -> bool:
        """라벨 정밀 대조는 그 문장이 지가변동률을 말할 때만 적용한다."""
        sentence = sentence_at(position)
        return any(word in sentence for word in ["지가변동률", "변동률", "누계", "%p"]) and not in_listing_context(position)

    # (1) 지가변동률/변동률/누계로 인용된 수치가 어떤 지가 데이터에도 없으면 허수 후보로 표기
    for match in re.finditer(r"(?:지가변동률|변동률|누계)[은는이가]?\s*(-?\d+\.\d+)\s*%", text):
        if in_listing_context(match.start()):
            continue
        number = round(float(match.group(1)), 3)
        if not any(abs(number - known) <= 0.0011 for known in rate_values):
            warns.append(f"'{match.group(0).strip()}' 값이 result 지가 데이터에 없음(허수 후보)")

    # (2) 용도지역·이용상황 라벨에 붙은 당월 변동률이 result 값과 일치하는지(라벨 단위 정밀 대조)
    for label in ["주거", "상업", "공업", "녹지", "보전관리", "생산관리", "계획관리",
                  "전", "답", "주거용", "상업용", "임야", "공업용"]:
        val = current.get(label)
        if val is None:
            continue
        # 같은 라벨이라도 누계·전월 값을 인용할 수 있으므로 세 값 중 하나와 맞으면 통과시킨다.
        allowed = [val, (lp.get("cumulative", {}) or {}).get(label), (lp.get("previous", {}) or {}).get(label)]
        for match in re.finditer(rf"(?<![가-힣]){label}(?![가-힣])\s*(-?\d+\.\d+)\s*%", text):
            if in_listing_context(match.start()):
                continue
            if not any(near(match.group(1), candidate) for candidate in allowed if candidate is not None):
                warns.append(f"'{label} {match.group(1)}%' ≠ result {label} 당월 {val}%")
    return list(dict.fromkeys(warns))


def build_qc_report(result: dict[str, Any], narratives: dict[str, Any]) -> str:
    """검수용 품질 리포트(마크다운) 문자열을 생성한다."""
    lint = lint_narratives(narratives)
    audit = audit_narrative_numbers(result, narratives)
    meta = result.get("meta", {}) or {}
    lines = [
        f"# 지가동향 서술 품질 점검 — {meta.get('region', '')} {meta.get('base_month', '')}",
        "",
        f"- 구조·문체 결함: **{len(lint)}건**",
        f"- 인용 수치 불일치 후보: **{len(audit)}건**",
        "",
        "## 구조·문체 점검",
    ]
    lines += [f"- {x}" for x in lint] or ["- (없음)"]
    lines += ["", "## 인용 수치 대조(연성 경고)"]
    lines += [f"- {x}" for x in audit] or ["- (없음)"]
    lines += ["", "※ 수치 대조는 라벨이 붙은 지가변동률(평균·누계·용도·이용상황)만 검사한 결과이며, 최종 확인은 조사담당자가 원표와 대조해 수행함."]
    return "\n".join(lines)


def regenerate_with_commentary(
    result: dict[str, Any],
    narratives: dict[str, Any],
    html_output_dir: str | Path,
    commentary_output_dir: str | Path | None = None,
    qc_output_dir: str | Path | None = None,
) -> dict[str, Any]:
    """지가변동률 전용 LLM JSON을 적용해 최종 HTML과 해설 JSON을 저장한다."""
    replacements = (result.get("narrative_policy", {}) or {}).get("forbidden_replacements", {})
    def sanitize(value: Any) -> Any:
        if isinstance(value, dict):
            return {key: sanitize(item) for key, item in value.items()}
        if isinstance(value, list):
            return [sanitize(item) for item in value]
        if isinstance(value, str):
            for forbidden, replacement in replacements.items():
                value = value.replace(str(forbidden), str(replacement))
        return value
    narratives = refresh_fact_bound_narratives(result, sanitize(narratives))
    # 거래주체·단가·순위가 원자료 재계산 후에도 과거 LLM 문장에 남지 않도록
    # TOP10 수치 서술은 항상 현재 result.json에서 다시 구성한다.
    for key in ["zone_change_reasoning", "use_change_reasoning"]:
        narratives[key] = limit_text(narratives.get(key), SECTION_CHAR_LIMITS[key])
    html_dir = Path(html_output_dir)
    commentary_dir = Path(commentary_output_dir) if commentary_output_dir else html_dir
    html_dir.mkdir(parents=True, exist_ok=True)
    commentary_dir.mkdir(parents=True, exist_ok=True)
    region = safe_name(result.get("meta", {}).get("region", "지역"))
    month = str(result.get("meta", {}).get("base_month", "")).replace("-", "")
    html_path = html_dir / f"{region}_토지시장동향_시황보고서_{month}_llm.html"
    commentary_path = commentary_dir / f"commentary_land_price_monthly_{region}_{month}_llm.json"
    commentary_path.write_text(json.dumps(narratives, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    html_path.write_text(build_html_report(result, narratives), encoding="utf-8")
    # 검수용 서술 품질 리포트(구조·문체 린트 + 인용 수치 대조)
    qc_dir = Path(qc_output_dir) if qc_output_dir else commentary_dir
    qc_dir.mkdir(parents=True, exist_ok=True)
    qc_path = qc_dir / f"qc_land_price_monthly_{region}_{month}.md"
    qc_lint = lint_narratives(narratives)
    qc_audit = audit_narrative_numbers(result, narratives)
    qc_path.write_text(build_qc_report(result, narratives), encoding="utf-8")
    return {
        "data_pack": result,
        "commentary": narratives,
        "commentary_pack_path": commentary_path,
        "html_path": html_path,
        "html_report": html_path,
        "qc_report_path": qc_path,
        "qc_lint_count": len(qc_lint),
        "qc_audit_count": len(qc_audit),
    }


def run_pipeline(
    input_dir: str | Path,
    region: str,
    base_month: str,
    output_dir: str | Path,
    transaction_files: list[str | Path] | None = None,
    listing_files: list[str | Path] | None = None,
    land_price_files: list[str | Path] | None = None,
    as_of_date: Any = None,
    district_rate_override: dict[str, Any] | None = None,
) -> dict[str, Any]:
    result = build_result(
        Path(input_dir),
        region,
        base_month,
        transaction_files=transaction_files,
        listing_files=listing_files,
        land_price_files=land_price_files,
        as_of_date=as_of_date,
        district_rate_override=district_rate_override,
    )
    paths = write_outputs(result, Path(output_dir))
    return {"result": result, **paths}


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate land-price monthly LLM analysis material.")
    parser.add_argument("--input", required=True, help="Input folder containing Excel files")
    parser.add_argument("--region", required=True, help="Report region, e.g. 나주시")
    parser.add_argument("--base-month", required=True, help="YYYY-MM")
    parser.add_argument("--output", default="OUTPUT/land_price_monthly", help="Output folder")
    parser.add_argument(
        "--as-of",
        default=None,
        help="주간 집계 작성 기준일 YYYY-MM-DD (기본: 오늘과 기준월 말일 중 이른 날, 직전 금요일로 정렬)",
    )
    args = parser.parse_args()
    out = run_pipeline(args.input, args.region, args.base_month, args.output, as_of_date=args.as_of)
    print(f"result_json={out['result_json']}")
    print(f"llm_prompt={out['llm_prompt']}")
    print(f"html_report={out['html_report']}")


if __name__ == "__main__":
    main()

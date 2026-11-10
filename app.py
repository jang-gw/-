# -*- coding: utf-8 -*-
"""한국도시가스협회 월보 공개 드릴다운 대시보드.

사용자 PC에서 생성·검증·게시한 web/data/citygas_public.xlsx만 읽습니다.
공개 엑셀은 원본 Excel의 4번 '분석' 및 5번 '그래프'에서 값만 추출한
2시트 검증본입니다. 원본 5시트 보고서도 검증 목적으로 지원합니다.
AI/API를 호출하지 않으며, 방문자 파일 업로드 기능이 없습니다.

배포: Streamlit Community Cloud의 엔트리포인트 app.py
로컬: python -m streamlit run web/app.py
검사: python web/app.py --self-test
"""
from __future__ import annotations

import argparse
import csv
import html
import io
import math
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

REGIONS = ("대구", "경북", "경남")
ALL = "3개 지역 합산"
METRICS = ("수요가수", "공급량")
METRIC_HEADER = {"수요가수": "수요가수(개)", "공급량": "공급량(천MJ)"}
UNITS = {"수요가수": "개", "공급량": "천MJ"}
CATEGORIES = ("가정용", "일반용", "업무용", "산업용", "열병합", "열전용", "수송용", "연료전지용")
MONTH_RE = re.compile(r"^20\d{2}-(?:0[1-9]|1[0-2])$")
COLORS = {"red": "#E00000", "peach": "#FFA97B", "cream": "#FFECCB", "green": "#9FED72"}
REGION_COLORS = {"대구": COLORS["red"], "경북": COLORS["peach"], "경남": COLORS["green"]}


class DashboardDataError(ValueError):
    """원본 시트의 구조 또는 숫자 검증 실패 시 사용자가 확인할 수 있는 오류."""


def as_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def as_number(value: Any, place: str) -> float:
    if isinstance(value, bool) or value is None:
        raise DashboardDataError(f"{place}: 수치가 없거나 숫자가 아닙니다({value!r}).")
    if isinstance(value, str):
        value = value.strip().replace(",", "")
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise DashboardDataError(f"{place}: 숫자를 읽을 수 없습니다({value!r}).") from exc
    if not math.isfinite(number) or number < 0:
        raise DashboardDataError(f"{place}: 유효하지 않은 음수/무한대입니다({number}).")
    return number


def approx(a: float, b: float, metric: str) -> bool:
    # 수요가수는 정수, 공급량은 소수(천MJ)입니다.
    return math.isclose(a, b, abs_tol=(0.01 if metric == "공급량" else 0.000001), rel_tol=1e-10)


def require(condition: bool, explanation: str) -> None:
    if not condition:
        raise DashboardDataError(explanation)


def parse_analysis(ws) -> dict:
    """표의 행 번호가 달라져도 헤더 이름으로 3개 영역을 추출합니다."""
    rows = list(ws.iter_rows(min_row=1, max_col=8, values_only=True))
    headers = {}
    signatures = {
        "monthly": {"기준월", "지역", "수요가수(개)", "공급량(천MJ)"},
        "company": {"지역", "회사명", "수요가수(개)", "공급량(천MJ)"},
        "category": {"지역", "지표", "용도", "값", "단위"},
    }
    for index, row in enumerate(rows):
        cells = {as_text(v) for v in row if as_text(v)}
        for section, required in signatures.items():
            if required.issubset(cells):
                require(section not in headers, f"분석 시트: '{section}' 헤더가 두 번 발견됐습니다.")
                headers[section] = (index, {as_text(v): j for j, v in enumerate(row) if as_text(v)})
    require(set(headers) == set(signatures),
            "분석 시트의 월별·회사별·용도별 표를 모두 찾지 못했습니다. "
            f"확인된 표: {', '.join(headers)}")
    boundaries = sorted([(pos, name) for name, (pos, _) in headers.items()])
    bounds = {}
    for i, (position, section) in enumerate(boundaries):
        bounds[section] = (position + 1, boundaries[i + 1][0] if i + 1 < len(boundaries) else len(rows))

    def cell(row, header, name):
        return row[headers[name][1][header]]

    monthly: dict[tuple[str, str], dict[str, float]] = {}
    for row in rows[slice(*bounds["monthly"])]:
        month, region = as_text(cell(row, "기준월", "monthly")), as_text(cell(row, "지역", "monthly"))
        if not MONTH_RE.fullmatch(month) or region not in (*REGIONS, ALL):
            continue
        key = (month, region)
        require(key not in monthly, f"분석 시트: 월별 집계 중복 {key}")
        monthly[key] = {metric: as_number(cell(row, METRIC_HEADER[metric], "monthly"), f"분석/{month}/{region}/{metric}")
                        for metric in METRICS}

    companies: dict[tuple[str, str], dict[str, float]] = {}
    for row in rows[slice(*bounds["company"])]:
        region, name = as_text(cell(row, "지역", "company")), as_text(cell(row, "회사명", "company"))
        if region not in REGIONS or not name:
            continue
        key = (region, name)
        require(key not in companies, f"분석 시트: 중복 회사 {key}")
        companies[key] = {metric: as_number(cell(row, METRIC_HEADER[metric], "company"), f"분석/{region}/{name}/{metric}")
                          for metric in METRICS}

    category: dict[tuple[str, str, str], float] = {}
    for row in rows[slice(*bounds["category"])]:
        region = as_text(cell(row, "지역", "category"))
        metric = as_text(cell(row, "지표", "category"))
        purpose = as_text(cell(row, "용도", "category"))
        unit = as_text(cell(row, "단위", "category"))
        if region not in (*REGIONS, ALL) or metric not in METRICS:
            continue
        require(purpose in CATEGORIES, f"분석 시트: 알 수 없는 용도 '{purpose}'")
        require(unit == UNITS[metric], f"분석 시트: '{region}/{metric}/{purpose}' 단위가 '{unit}'입니다.")
        key = (region, metric, purpose)
        require(key not in category, f"분석 시트: 용도별 집계 중복 {key}")
        category[key] = as_number(cell(row, "값", "category"), f"분석/{region}/{metric}/{purpose}")

    require(monthly and companies and category, "분석 시트의 집계 표 중 하나가 비어 있습니다.")
    months = sorted({m for m, r in monthly if r in REGIONS})
    require(months, "분석 시트에 월별 데이터가 없습니다.")
    latest = months[-1]
    for month in months:
        for region in (*REGIONS, ALL):
            require((month, region) in monthly, f"분석 시트: {month}/{region} 집계가 누락됐습니다.")
        for metric in METRICS:
            subtotal = sum(monthly[(month, r)][metric] for r in REGIONS)
            require(approx(subtotal, monthly[(month, ALL)][metric], metric),
                    f"분석 시트: {month}/{metric} 3개 지역 합계와 합산 행이 다릅니다.")
    for region in REGIONS:
        region_companies = [v for (r, _), v in companies.items() if r == region]
        require(region_companies, f"분석 시트: {region}에 회사별 상세 데이터가 없습니다.")
        for metric in METRICS:
            total = sum(v[metric] for v in region_companies)
            require(approx(total, monthly[(latest, region)][metric], metric),
                    f"분석 시트: 최신월 {region}/{metric} 회사 합계({total})가 지역 합계와 다릅니다.")
    for metric in METRICS:
        for region in (*REGIONS, ALL):
            subtotal = sum(v for (r, m, _), v in category.items() if r == region and m == metric)
            require(approx(subtotal, monthly[(latest, region)][metric], metric),
                    f"분석 시트: 최신월 {region}/{metric} 용도별 합계가 지역 합계와 다릅니다.")
        for purpose in CATEGORIES:
            summed = sum(category.get((r, metric, purpose), 0.0) for r in REGIONS)
            require(approx(summed, category.get((ALL, metric, purpose), 0.0), metric),
                    f"분석 시트: {metric}/{purpose} 합산이 지역별 수치와 다릅니다.")

    # 회사 및 용도 상세는 최신월만 수록하는 구조입니다. 제목도 함께 확인합니다.
    for section in ("company", "category"):
        section_row = headers[section][0]
        surrounding = " ".join(as_text(x) for row in rows[max(0, section_row-3):section_row] for x in row)
        found = re.findall(r"20\d{2}-(?:0[1-9]|1[0-2])", surrounding)
        if found:
            require(latest in found, f"분석 시트: '{section}' 구간의 표시 기준월이 최신월 {latest}와 다릅니다.")
    return {"monthly": monthly, "companies": companies, "category": category,
            "months": months, "latest": latest}


def parse_graph(ws, analysis: dict) -> dict:
    """5번 시트에서 숨겨진 차트 원본 테이블 2+8개를 읽고 4번 시트와 대조합니다.

    그려진 차트 객체/렌더링 상태는 사용하지 않습니다. 원본 셀의 숫자만 읽습니다.
    """
    max_col = min(ws.max_column, 200)
    # read_only 워크시트에 cell()을 반복 호출하면 Excel XML을 매번 순회할 수 있어
    # 표 전체를 한 번만 읽습니다. 숨겨진 열도 값은 그대로 포함됩니다.
    matrix = list(ws.iter_rows(min_row=1, max_col=max_col, values_only=True))
    require(bool(matrix), "그래프 시트가 비어 있습니다.")
    top = [as_text(value) for value in matrix[0]]
    month_tables = []
    category_tables = []
    for c in range(max_col):
        if (c + 3 < max_col and top[c] == "기준월" and
                list(top[c+1:c+4]) == list(REGIONS)):
            values = {}
            for row_idx, cells in enumerate(matrix[1:], start=2):
                month = as_text(cells[c])
                if not MONTH_RE.fullmatch(month):
                    if values:
                        break
                    continue
                require(month not in values, f"그래프 시트: {c+1}열 월별 원본 중복 {month}")
                values[month] = {region: as_number(cells[c+i], f"그래프/{c+1}열/{month}/{region}")
                                 for i, region in enumerate(REGIONS, start=1)}
            month_tables.append(values)
        if c + 1 < max_col and top[c] == "용도" and top[c+1] == "값":
            values = {}
            for row_idx, cells in enumerate(matrix[1:101], start=2):
                purpose = as_text(cells[c])
                if purpose not in CATEGORIES:
                    if values:
                        break
                    continue
                require(purpose not in values, f"그래프 시트: {c+1}열 용도 중복 {purpose}")
                values[purpose] = as_number(cells[c+1], f"그래프/{c+1}열/{purpose}")
            category_tables.append(values)

    require(len(month_tables) == 2,
            f"그래프 시트: 월별 원본 표는 2개여야 하지만 {len(month_tables)}개입니다. "
            "오른쪽 숨김 열이 보존됐는지 확인하세요.")
    require(len(category_tables) == 8,
            f"그래프 시트: 용도별 원본 표는 8개여야 하지만 {len(category_tables)}개입니다.")
    matched: dict[str, dict] = {}
    for table in month_tables:
        require(set(table) == set(analysis["months"]), "그래프 시트: 월별 원본의 기준월이 분석 시트와 다릅니다.")
        possible = [metric for metric in METRICS
                    if all(approx(table[month][region], analysis["monthly"][(month, region)][metric], metric)
                           for month in analysis["months"] for region in REGIONS)]
        require(len(possible) == 1, "그래프 시트: 월별 표의 지표를 분석 데이터와 일치시킬 수 없습니다.")
        require(possible[0] not in matched, "그래프 시트: 월별 지표 중복입니다.")
        matched[possible[0]] = table
    require(set(matched) == set(METRICS), "그래프 시트: 두 지표 중 하나가 없습니다.")
    sequence = [(metric, region) for metric in METRICS for region in (ALL, *REGIONS)]
    for values, (metric, region) in zip(category_tables, sequence):
        expected = {purpose: number for (r, m, purpose), number in analysis["category"].items()
                    if r == region and m == metric and number > 0}
        require(set(values) == set(expected),
                f"그래프 시트: {region}/{metric} 용도 항목이 분석 시트와 다릅니다.")
        for purpose, number in values.items():
            require(approx(number, expected[purpose], metric),
                    f"그래프 시트: {region}/{metric}/{purpose}의 그래프 값과 분석 값이 다릅니다.")
    return {"trend": matched,
            "category_tables": {key: value for key, value in zip(sequence, category_tables)}}


def read_report(path_or_bytes: Path | str | bytes | io.BytesIO) -> dict:
    """엑셀 파일을 수정하지 않으며 4·5번 시트만 참조합니다."""
    if isinstance(path_or_bytes, bytes):
        source = io.BytesIO(path_or_bytes)
    else:
        source = path_or_bytes
    try:
        book = load_workbook(source, read_only=True, data_only=True, keep_links=False)
    except Exception as exc:
        raise DashboardDataError(f"Excel 파일을 열 수 없습니다: {exc}") from exc
    try:
        if book.sheetnames == ["분석", "그래프"]:
            # 공개 파일은 원본 5개 시트 중 분석/그래프의 값만 복사한 두 시트입니다.
            analysis_ws, graph_ws = book.worksheets
        else:
            require(len(book.sheetnames) == 5, "원본 보고서는 5개 시트, 공개 파일은 분석/그래프 2개 시트여야 합니다.")
            require(book.sheetnames[3:5] == ["분석", "그래프"],
                    "원본 Excel의 4번·5번 시트는 '분석'과 '그래프'여야 합니다. "
                    f"현재 순서: {book.sheetnames}")
            analysis_ws, graph_ws = book.worksheets[3:5]
        analysis = parse_analysis(analysis_ws)
        graph = parse_graph(graph_ws, analysis)
        analysis["graph"] = graph
        return analysis
    finally:
        book.close()


def desktop_directory() -> Path:
    home = Path(os.environ.get("USERPROFILE", str(Path.home())))
    folders = (home / "Desktop", Path(os.environ.get("OneDrive", str(home / "OneDrive"))) / "Desktop",
               home / "바탕 화면")
    return next((folder for folder in folders if folder.is_dir()), folders[0])


def default_root() -> Path:
    return desktop_directory() / "도시가스월보_자동화"


def discover_reports(root: Path, selected: str | None = None) -> list[Path]:
    reports = root / "보고서"
    candidates: list[Path] = []
    if reports.is_dir():
        candidates = [p.resolve() for p in reports.glob("도시가스월보_최신화_*.xlsx")
                      if p.is_file() and not p.name.startswith("~$")]
    # 기존 수집기의 SQLite에는 마지막 생성 성공한 보고서 경로가 저장됩니다.
    db_path = root / "상태관리" / "citygas.sqlite3"
    preferred: Path | None = None
    if db_path.is_file():
        try:
            uri = f"file:{db_path.resolve().as_posix()}?mode=ro"
            with sqlite3.connect(uri, uri=True, timeout=2) as db:
                result = db.execute("SELECT value FROM meta WHERE key='last_report'").fetchone()
                if result:
                    p = Path(result[0]).resolve()
                    if p.is_file():
                        preferred = p
                        if p not in candidates:
                            candidates.append(p)
        except (OSError, sqlite3.DatabaseError):
            pass  # 보고서 디렉터리 스캔을 계속 사용합니다.
    candidates = sorted(set(candidates), key=lambda p: p.stat().st_mtime_ns, reverse=True)
    if preferred in candidates:
        candidates.remove(preferred)
        candidates.insert(0, preferred)
    if selected:
        explicit = Path(selected).expanduser().resolve()
        if explicit.is_file():
            if explicit in candidates:
                candidates.remove(explicit)
            candidates.insert(0, explicit)
        else:
            raise DashboardDataError(f"--excel로 지정한 파일이 없습니다: {explicit}")
    return candidates


def format_value(value: float, metric: str, *, pct: bool = False) -> str:
    if pct:
        return f"{value * 100:.1f}%"
    return f"{value:,.0f}" if metric == "수요가수" else f"{value:,.3f}"


def delta_text(current: float, prior: float | None, metric: str) -> str:
    if prior is None:
        return "이전월 자료 없음"
    change = current - prior
    percent = f"{(change / prior * 100):+.1f}%" if prior else "증감률 계산 불가"
    change_str = f"{change:+,.0f}" if metric == "수요가수" else f"{change:+,.3f}"
    return f"전월 대비 {change_str} {UNITS[metric]} / {percent}"


def csv_data(headers: list[str], records: list[list]) -> bytes:
    stream = io.StringIO()
    writer = csv.writer(stream)
    writer.writerow(headers)
    writer.writerows(records)
    return stream.getvalue().encode("utf-8-sig")


def previous_month(data: dict, month: str) -> str | None:
    months = data["months"]
    idx = months.index(month)
    # 중간 월 누락이 있으면 인접한 관측을 '전월'이라고 부르지 않습니다.
    if idx == 0:
        return None
    prior = months[idx - 1]
    year, number = map(int, month.split("-"))
    previous = f"{year - 1}-12" if number == 1 else f"{year}-{number - 1:02d}"
    return prior if prior == previous else None


def chart_layout(fig, *, height: int = 340):
    fig.update_layout(height=height, margin=dict(l=12, r=12, t=30, b=20),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="Malgun Gothic, sans-serif", size=12, color="#2A2A29"),
                      legend=dict(orientation="h", y=1.13, x=0),
                      hoverlabel=dict(bgcolor="#FFFFFF", font_size=12))
    fig.update_xaxes(showgrid=False, linecolor="#EBE8E0")
    fig.update_yaxes(showgrid=True, gridcolor="#F2E9DB", zeroline=False)
    return fig


def region_bars(data: dict, month: str, metric: str):
    import plotly.graph_objects as go
    values = [data["monthly"][(month, region)][metric] for region in REGIONS]
    fig = go.Figure(go.Bar(x=list(REGIONS), y=values, marker_color=[REGION_COLORS[r] for r in REGIONS],
                           text=[format_value(v, metric) for v in values], textposition="outside",
                           hovertemplate="%{x}<br>%{y:,.3f} " + UNITS[metric] + "<extra></extra>"))
    fig.update_layout(showlegend=False)
    fig.update_yaxes(title_text=UNITS[metric], rangemode="tozero")
    return chart_layout(fig)


def monthly_lines(data: dict, metric: str, focus_region: str | None = None):
    import plotly.graph_objects as go
    months = data["months"]
    fig = go.Figure()
    for region in (REGIONS if focus_region is None else (focus_region,)):
        # 숫자는 5번 '그래프' 시트의 검증된 원본을 사용합니다.
        values = [data["graph"]["trend"][metric][m][region] for m in months]
        fig.add_trace(go.Scatter(x=months, y=values, mode="lines+markers", name=region,
                                 line=dict(color=REGION_COLORS[region], width=3),
                                 marker=dict(size=8),
                                 hovertemplate="%{x}<br>%{y:,.3f} " + UNITS[metric] + "<extra>" + region + "</extra>"))
    fig.update_yaxes(title_text=UNITS[metric], rangemode="tozero")
    return chart_layout(fig, height=340)


def company_bars(data: dict, region: str, metric: str, *, focus: str | None = None):
    import plotly.graph_objects as go
    rows = [(name, result[metric]) for (r, name), result in data["companies"].items() if r == region]
    rows.sort(key=lambda item: item[1])
    names, values = zip(*rows)
    fill = [COLORS["red"] if focus == name else COLORS["peach"] if focus else REGION_COLORS[region]
            for name in names]
    fig = go.Figure(go.Bar(x=list(values), y=list(names), orientation="h", marker_color=fill,
                           hovertemplate="%{y}<br>%{x:,.3f} " + UNITS[metric] + "<extra></extra>"))
    fig.update_layout(showlegend=False, height=max(280, 58 * len(rows) + 85))
    fig.update_xaxes(title_text=UNITS[metric])
    fig.update_yaxes(title_text="", automargin=True)
    return chart_layout(fig, height=max(280, 58 * len(rows) + 85))


def purpose_donut(data: dict, region: str, metric: str):
    import plotly.graph_objects as go
    values = data["graph"]["category_tables"][(metric, region)]
    purposes = sorted(values, key=lambda x: values[x], reverse=True)
    palette = [COLORS["red"], COLORS["peach"], COLORS["green"], "#F5BA34",
               "#FF7878", "#7EBFCE", "#BEA1E6", "#9CBA86"]
    fig = go.Figure(go.Pie(labels=purposes, values=[values[p] for p in purposes], hole=0.65,
                           marker=dict(colors=palette), textinfo="percent",
                           hovertemplate="%{label}<br>%{value:,.3f} " + UNITS[metric] + " (%{percent})<extra></extra>"))
    fig.update_layout(showlegend=True, legend=dict(orientation="v", x=1.02, y=.5))
    return chart_layout(fig, height=400)


def category_comparison(data: dict, metric: str, purpose: str):
    import plotly.graph_objects as go
    values = [data["category"].get((r, metric, purpose), 0) for r in REGIONS]
    fig = go.Figure(go.Bar(x=list(REGIONS), y=values,
                           marker_color=[REGION_COLORS[r] for r in REGIONS],
                           text=[format_value(v, metric) for v in values], textposition="outside",
                           hovertemplate="%{x}<br>%{y:,.3f} " + UNITS[metric] + "<extra></extra>"))
    fig.update_layout(showlegend=False)
    fig.update_yaxes(title_text=UNITS[metric], rangemode="tozero")
    return chart_layout(fig)


def style_css():
    import streamlit as st
    st.markdown("""<style>
    html, body, [class*="css"], [data-testid="stApp"] {font-family:'Malgun Gothic','맑은 고딕',sans-serif;}
    [data-testid="stAppViewContainer"] {background:#FFFDF8;}
    [data-testid="stSidebar"] {background:#FFECCB;}
    .block-container {max-width:1200px;padding-top:2rem;padding-bottom:3rem;}
    h1,h2,h3 {letter-spacing:-0.03em;}
    .eyebrow {font-size:.79rem;font-weight:800;color:#E00000;letter-spacing:.13em;}
    .hero {font-weight:800;font-size:2.1rem;letter-spacing:-.055em;line-height:1.35;margin:.22rem 0 .4rem;}
    .subtitle {font-size:.9rem;color:#6F6C66;margin-bottom:1rem;}
    .stat-card {background:#FFFFFF;border:1px solid #F0E6D8;border-radius:16px;padding:19px 22px;min-height:139px;box-shadow:0 3px 17px #473c2b09;}
    .stat-label {color:#817D75;font-size:.82rem;font-weight:600;margin-bottom:.5rem;}
    .stat-number {font-size:1.58rem;font-weight:800;letter-spacing:-.045em;line-height:1.22;white-space:normal;overflow-wrap:break-word;}
    .stat-foot {color:#69665E;font-size:.77rem;margin-top:.7rem;}
    .stat-accent {width:38px;height:4px;border-radius:5px;background:#E00000;margin-bottom:13px;}
    .section-hint {font-size:.86rem;color:#6F6C66;margin-bottom:.6rem;}
    .note {padding:12px 16px;border-left:4px solid #FFA97B;background:#FFF3E5;border-radius:7px;font-size:.85rem;}
    div.stButton > button[kind="primary"] {background:#E00000;border-color:#E00000;color:#fff;}
    div.stButton > button {border-radius:10px;font-weight:600;}
    [data-testid="stMetricValue"] {font-weight:800;}
    </style>""", unsafe_allow_html=True)


def safe_html(text: Any) -> str:
    return html.escape(str(text), quote=True)


def stat_card(label: str, value: str, foot: str, color: str | None = None):
    import streamlit as st
    st.markdown(f'''<div class="stat-card"><div class="stat-accent" style="background:{color or COLORS['red']}"></div>
    <div class="stat-label">{safe_html(label)}</div><div class="stat-number">{safe_html(value)}</div>
    <div class="stat-foot">{safe_html(foot)}</div></div>''', unsafe_allow_html=True)


def set_page(view: str, *, region: str | None = None, company: str | None = None):
    import streamlit as st
    st.session_state["view"] = view
    if region is not None:
        st.session_state["region"] = region
    if company is not None:
        st.session_state["company"] = company
    st.rerun()


def show_overview(data: dict, month: str, metric: str):
    import streamlit as st
    st.subheader("전체 현황")
    st.caption("지역을 선택하면 월별 변화, 회사별 실적, 용도별 구성으로 차례로 내려갈 수 있습니다.")
    previous = previous_month(data, month)
    cols = st.columns(2)
    for col, met in zip(cols, METRICS):
        amount = data["monthly"][(month, ALL)][met]
        prior = data["monthly"][(previous, ALL)][met] if previous else None
        with col:
            stat_card(f"{month} · {met}", format_value(amount, met) + " " + UNITS[met],
                      delta_text(amount, prior, met), COLORS["red"] if met == "공급량" else COLORS["green"])
    st.markdown("### 지역을 선택해 상세 분석하기")
    columns = st.columns(3)
    for col, region in zip(columns, REGIONS):
        with col:
            value = data["monthly"][(month, region)][metric]
            whole = data["monthly"][(month, ALL)][metric]
            stat_card(region, format_value(value, metric) + " " + UNITS[metric],
                      f"3개 지역 합계 대비 {value / whole * 100:.1f}%" if whole else "합계 0",
                      REGION_COLORS[region])
            if st.button(f"{region} 상세 보기", key=f"open_{region}", use_container_width=True):
                set_page("region", region=region)
    with st.expander("전체 지역 비교와 월별 추이 보기", expanded=False):
        left, right = st.columns(2)
        with left:
            st.markdown(f"**{month} · 지역 비교**")
            st.plotly_chart(region_bars(data, month, metric), use_container_width=True, key="overview_bar")
        with right:
            st.markdown("**월별 추이**")
            st.plotly_chart(monthly_lines(data, metric), use_container_width=True, key="overview_line")
    st.markdown("<div class='note'>수요가수는 매월 현황이며 합산 누계가 아닙니다. 공급량은 월별 당월 실적(천MJ)입니다.</div>",
                unsafe_allow_html=True)


def show_region(data: dict, region: str, month: str, metric: str):
    import streamlit as st
    if st.button("← 전체 현황", key="back_to_overview"):
        set_page("overview")
    st.subheader(f"{region} · 지역 분석")
    previous = previous_month(data, month)
    cols = st.columns(2)
    for col, met in zip(cols, METRICS):
        value = data["monthly"][(month, region)][met]
        previous_value = data["monthly"][(previous, region)][met] if previous else None
        with col:
            stat_card(f"{month} · {met}", format_value(value, met) + " " + UNITS[met],
                      delta_text(value, previous_value, met), REGION_COLORS[region])
    st.markdown("### 월별 변화")
    st.plotly_chart(monthly_lines(data, metric, region), use_container_width=True, key="region_line")
    st.markdown(f"### 최신월({data['latest']}) 상세로 이동")
    if month != data["latest"]:
        st.info(f"회사별·용도별 상세는 {data['latest']} 자료만 제공됩니다. 현재 선택월({month})의 회사별 세부 실적은 생성하지 않습니다.")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("회사별 실적 및 비중", use_container_width=True, type="primary", key="goto_company_list"):
            if month != data["latest"]:
                st.session_state["pending_month_reset"] = data["latest"]
            set_page("company_list", region=region)
    with c2:
        if st.button("용도별 구성과 지역 비교", use_container_width=True, key="goto_category"):
            if month != data["latest"]:
                st.session_state["pending_month_reset"] = data["latest"]
            set_page("category", region=region)


def show_companies(data: dict, region: str, metric: str):
    import streamlit as st
    if st.button(f"← {region} 분석", key="back_region_companies"):
        set_page("region", region=region)
    st.subheader(f"{region} · 회사별 비교")
    st.caption(f"회사별 데이터는 최신월 {data['latest']}만 수록되어 있습니다. 다른 월의 회사별 실적은 제공하지 않습니다.")
    st.plotly_chart(company_bars(data, region, metric), use_container_width=True, key="company_chart")
    choices = [name for (r, name) in data["companies"] if r == region]
    choices.sort(key=lambda name: data["companies"][(region, name)][metric], reverse=True)
    selected = st.selectbox("상세 조회할 회사 선택", choices, key="selected_company")
    if st.button("선택한 회사 상세", type="primary", key="show_company"):
        set_page("company", region=region, company=selected)
    st.download_button("회사별 표 CSV 내려받기", data=csv_data(
        ["지역", "회사", "수요가수(개)", "공급량(천MJ)", "기준월"],
        [[region, name, data["companies"][(region, name)]["수요가수"],
          data["companies"][(region, name)]["공급량"], data["latest"]] for name in choices]),
        file_name=f"회사별_{region}_{data['latest']}.csv", mime="text/csv", key="companies_csv")


def show_company(data: dict, region: str, company: str, metric: str):
    import streamlit as st
    if st.button(f"← {region} 회사 비교", key="back_to_companies"):
        set_page("company_list", region=region)
    key = (region, company)
    if key not in data["companies"]:
        st.error("선택한 회사가 현재 보고서에 없습니다. 회사 목록으로 돌아가세요.")
        return
    st.subheader(f"{company} · 회사 상세")
    st.caption(f"{data['latest']} · 4번 '분석' 시트의 최신월 회사별 실적")
    cols = st.columns(2)
    for col, met in zip(cols, METRICS):
        value = data["companies"][key][met]
        region_total = data["monthly"][(data["latest"], region)][met]
        with col:
            stat_card(met, f"{format_value(value, met)} {UNITS[met]}",
                      f"{region} 전체 대비 {value / region_total * 100:.2f}%" if region_total else "지역 합계 0",
                      COLORS["red"] if met == "공급량" else COLORS["green"])
    st.markdown(f"### {region} 내 회사 간 {metric} 비교")
    st.plotly_chart(company_bars(data, region, metric, focus=company), use_container_width=True, key="single_company_chart")
    st.markdown("<div class='note'>회사별 월별 추이와 회사별 용도 구성은 현재 지정된 4·5번 시트에 없으므로 표시하지 않습니다.</div>",
                unsafe_allow_html=True)


def show_categories(data: dict, region: str, metric: str):
    import streamlit as st
    if st.button(f"← {region} 지역 분석", key="back_region_cats"):
        set_page("region", region=region)
    st.subheader(f"{region} · 용도별 구성")
    st.caption(f"기준월 {data['latest']} · {metric}({UNITS[metric]}) · 최신월 구성만 제공")
    values = data["graph"]["category_tables"][(metric, region)]
    total = data["monthly"][(data["latest"], region)][metric]
    ranked = sorted(values.items(), key=lambda x: x[1], reverse=True)
    left, right = st.columns([1.45, 1])
    with left:
        st.plotly_chart(purpose_donut(data, region, metric), use_container_width=True, key="cat_donut")
    with right:
        st.markdown("#### 비중 상위 용도")
        for purpose, number in ranked[:4]:
            portion = number / total if total else 0
            st.markdown(f"**{purpose}**　{format_value(number, metric)} {UNITS[metric]} ({portion*100:.1f}%)")
            st.progress(min(1.0, max(0.0, portion)))
    with st.expander("전체 용도별 표 및 CSV", expanded=False):
        st.dataframe([{"용도": purpose, f"{metric}({UNITS[metric]})": value,
                       "구성비(%)": round(value / total * 100, 2) if total else 0}
                      for purpose, value in ranked], use_container_width=True, hide_index=True)
        st.download_button("용도별 표 CSV 내려받기", data=csv_data(
            ["지역", "지표", "용도", "집계값", "단위", "기준월"],
            [[region, metric, purpose, value, UNITS[metric], data["latest"]] for purpose, value in ranked]),
            file_name=f"용도별_{region}_{metric}_{data['latest']}.csv", mime="text/csv", key="cat_csv")
    purpose = st.selectbox("어떤 용도의 지역별 실적을 비교하시겠습니까?", [p for p, _ in ranked], key="chosen_purpose")
    if st.button(f"'{purpose}' 지역별 비교", type="primary", key="open_purpose"):
        set_page("category_region", region=region, company=purpose)


def show_category_region(data: dict, purpose: str, region: str, metric: str):
    import streamlit as st
    if st.button("← 용도별 구성", key="back_to_category"):
        set_page("category", region=region)
    st.subheader(f"{purpose} · 지역별 비교")
    st.caption(f"{data['latest']} · {metric}({UNITS[metric]})")
    value = data["category"].get((ALL, metric, purpose), 0.0)
    stat_card("3개 지역 합산", format_value(value, metric) + " " + UNITS[metric],
              "최신월 용도별 구성 자료", COLORS["red"])
    st.plotly_chart(category_comparison(data, metric, purpose), use_container_width=True, key="purpose_comparison")
    st.caption("이 그래프는 동일 기준월의 지역 간 비교이며, 월별 용도 추이가 아닙니다.")


def render_dashboard():
    import streamlit as st
    st.set_page_config(page_title="도시가스 월보 | 드릴다운", page_icon="📈", layout="wide",
                       initial_sidebar_state="expanded")
    style_css()
    # 공개용 웹앱은 저장소에 포함된 검증된 데이터 파일만 읽습니다.
    # 방문자가 업로드한 Excel이나 개인 PC 경로는 사용하지 않습니다.
    published_file = Path(__file__).resolve().parent / "data" / "citygas_public.xlsx"

    st.markdown("<div class='eyebrow'>CITY GAS · MONTHLY INTELLIGENCE</div>", unsafe_allow_html=True)
    st.markdown("<div class='hero'>도시가스 월보 · 드릴다운 대시보드</div>", unsafe_allow_html=True)
    st.markdown("<div class='subtitle'>Excel의 4번 ‘분석’ + 5번 ‘그래프’ 시트만 활용 · AI 미사용 · 원본 수치 검증</div>",
                unsafe_allow_html=True)
    with st.sidebar:
        st.markdown("### 공개 검증 데이터")
        st.caption("최종 게시된 검증 자료만 제공합니다.")
        if st.button("최신 자료 새로고침", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        st.divider()
        st.caption("자료 원천: 한국도시가스협회 도시가스사업 통계월보")
        st.caption("대상: 대구 · 경북 · 경남")
        st.caption("사용자 파일 업로드 기능은 제공하지 않습니다.")

    if not published_file.is_file():
        st.info("아직 공개용 데이터가 업로드되지 않았습니다. 관리자 PC에서 citygas_run.py를 실행한 뒤 "+
                "web/data/citygas_public.xlsx를 GitHub에 업로드하세요.")
        st.stop()

    @st.cache_data(show_spinner="Excel 분석·그래프 수치 검증 중...")
    def parse_file(path: str, modified_ns: int, size: int):
        return read_report(Path(path))

    try:
        file_info = published_file.stat()
        data = parse_file(str(published_file), file_info.st_mtime_ns, file_info.st_size)
        source_name = published_file.name
        identity = f"published:{file_info.st_mtime_ns}:{file_info.st_size}"
    except (DashboardDataError, OSError, ValueError) as exc:
        st.error(f"게시된 Excel 자료의 검증 실패: {exc}")
        st.caption("4번 분석 / 5번 그래프의 수치 및 구조가 일치하지 않아 대시보드를 표시하지 않습니다.")
        st.stop()

    if st.session_state.get("report_identity") != identity:
        st.session_state["report_identity"] = identity
        st.session_state["view"] = "overview"
        st.session_state["region"] = "대구"
        st.session_state["company"] = ""
        st.session_state.pop("month_control", None)
        for key in ("selected_company", "chosen_purpose"):
            st.session_state.pop(key, None)

    if st.session_state.get("pending_month_reset"):
        st.session_state["month_control"] = st.session_state.pop("pending_month_reset")

    with st.sidebar:
        st.markdown("### 분석 조건")
        if st.session_state.get("month_control") not in data["months"]:
            st.session_state["month_control"] = data["latest"]
        selected_month = st.select_slider("기준월", options=data["months"], key="month_control")
        metric = st.radio("분석 지표", METRICS, horizontal=True, key="metric_option")
        st.divider()
        if st.button("전체 현황으로", use_container_width=True, key="menu_overview"):
            set_page("overview")
        st.caption(f"월보 수록: {data['months'][0]} ~ {data['latest']}")
        st.caption("검증: 월별 합계 · 회사 합계 · 용도 합계 · 그래프 원본 대조")

    st.caption(f"공개 검증 자료: {source_name}   |   데이터 최신월: {data['latest']}")
    view = st.session_state.get("view", "overview")
    region = st.session_state.get("region", "대구")
    if view not in {"overview", "region", "company_list", "company", "category", "category_region"}:
        view = "overview"
    if view == "overview":
        show_overview(data, selected_month, metric)
    elif view == "region":
        show_region(data, region, selected_month, metric)
    elif view == "company_list":
        show_companies(data, region, metric)
    elif view == "company":
        show_company(data, region, st.session_state.get("company", ""), metric)
    elif view == "category":
        show_categories(data, region, metric)
    else:
        show_category_region(data, st.session_state.get("company", ""), region, metric)

    with st.expander("자료 범위와 검증 원칙", expanded=False):
        st.markdown("""- **4번 시트(분석)**: 월별 지역 집계, 최신월 회사별 현황, 최신월 용도별 집계.
- **5번 시트(그래프)**: 숨김 열에 기록된 월별 추이 2개 표, 최신월 용도별 구성 8개 표.
- **자동 교차 검증**: 세 지역 합계, 회사 실적 합계, 용도별 합계, 두 시트의 숫자 일치 여부.
- **제약**: 회사별 월별 추이 및 과거 월별 용도 구성은 이 두 시트에 없으므로 제공하지 않습니다.
- **단위**: 수요가수는 해당 월 현재값(개), 공급량은 그달 실적(천MJ). 누적 자료가 아닙니다.
- **인터넷/AI**: 대시보드는 GitHub에 게시된 검증 Excel만 읽으며 외부 AI API를 사용하지 않습니다.""")


def self_test() -> None:
    """인터넷/Excel COM 없이 실행 가능한 합성 데이터 기반 파서 회귀 테스트."""
    from openpyxl import Workbook
    wb = Workbook()
    wb.active.title = "원본추출"
    wb.create_sheet("분석데이터")
    wb.create_sheet("피벗원본")
    analysis = wb.create_sheet("분석")
    graph = wb.create_sheet("그래프")
    months = ["2026-01", "2026-02"]
    companies = {"대구": ["대성에너지", "테스트가스A"], "경북": ["경북가스A", "경북가스B"],
                 "경남": ["경남가스A", "경남가스B"]}
    company_latest = {r: [(name, 100 * (j+1), 300.25 * (j+1)) for j, name in enumerate(names)]
                      for r, names in companies.items()}
    analysis.append(["도시가스사업통계월보"])
    analysis.append([])
    analysis.append([])
    analysis.append(["지역별 월간 현황 및 전월 대비 변화"])
    analysis.append(["기준월", "지역", "수요가수(개)", "수요가수 전월증감(개)", "수요가수 전월증감률",
                     "공급량(천MJ)", "공급량 전월증감(천MJ)", "공급량 전월증감률"])
    for mi, month in enumerate(months):
        for region in (*REGIONS, ALL):
            multiplier = (mi+1) * (3 if region == ALL else 1)
            analysis.append([month, region, 150 * multiplier, None, None, 450.375 * multiplier, None, None])
    analysis.append(["회사별 최신월 상세 (2026-02)"])
    analysis.append(["지역", "회사명", "수요가수(개)", "지역 내 수요가수 비중", "공급량(천MJ)", "지역 내 공급량 비중"])
    for region in REGIONS:
        for name, demand, supply in company_latest[region]:
            analysis.append([region, name, demand, demand/300, supply, supply/900.75])
    analysis.append(["최신월 용도별 구성 (2026-02)"])
    analysis.append(["지역", "지표", "용도", "값", "단위"])
    for region in (ALL, *REGIONS):
        for metric in METRICS:
            factor = 3 if region == ALL else 1
            val = 300 if metric == "수요가수" else 900.75
            analysis.append([region, metric, "가정용", 0.6*val*factor, UNITS[metric]])
            analysis.append([region, metric, "산업용", 0.4*val*factor, UNITS[metric]])
    for idx, metric in enumerate(METRICS, start=1):
        col = 33 + idx*6
        graph.cell(1, col, "기준월")
        for i, region in enumerate(REGIONS, 1):
            graph.cell(1, col+i, region)
        for i, month in enumerate(months, start=2):
            graph.cell(i, col, month)
            for j, region in enumerate(REGIONS, start=1):
                graph.cell(i, col+j, (150*(i-1) if metric == "수요가수" else 450.375*(i-1)))
    for idx, (metric, region) in enumerate([(m, r) for m in METRICS for r in (ALL, *REGIONS)]):
        col = 50 + idx*3
        graph.cell(1, col, "용도")
        graph.cell(1, col+1, "값")
        factor = 3 if region == ALL else 1
        val = 300 if metric == "수요가수" else 900.75
        graph.cell(2, col, "가정용")
        graph.cell(2, col+1, 0.6*val*factor)
        graph.cell(3, col, "산업용")
        graph.cell(3, col+1, 0.4*val*factor)
    with tempfile.TemporaryDirectory() as tmp:
        file = Path(tmp) / "example.xlsx"
        wb.save(file)
        wb.close()
        report = read_report(file)
        assert report["latest"] == "2026-02"
        assert len(report["graph"]["trend"]) == 2
        assert report["monthly"][("2026-02", ALL)]["수요가수"] == 900
        assert len(report["companies"]) == 6
        assert approx(report["graph"]["category_tables"][("공급량", "경북")]["가정용"],
                      540.45, "공급량")
        # 실제 차트 생성 가능성 및 각 차트 원본 계열도 검증합니다.
        for fig in (region_bars(report, "2026-02", "공급량"),
                    monthly_lines(report, "수요가수"),
                    monthly_lines(report, "공급량", "대구"),
                    company_bars(report, "대구", "공급량"),
                    purpose_donut(report, "대구", "공급량"),
                    category_comparison(report, "수요가수", "산업용")):
            assert fig.data, "차트 데이터가 비어 있습니다."
        # 4번 시트와 5번 시트의 값이 다르면 자동 중단해야 합니다.
        from openpyxl import load_workbook as open_wb
        corrupted = open_wb(file)
        corrupted["그래프"].cell(2, 40, 999999)
        corrupted.save(file)
        corrupted.close()
        try:
            read_report(file)
        except DashboardDataError:
            pass
        else:
            raise AssertionError("그래프 원본 불일치를 놓쳤습니다.")
    print("자체 테스트 통과: 4·5번 시트 파싱 / 합계 / 숨김 차트 원본 / 지역·회사·용도 교차 검증")


def running_in_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx(suppress_warning=True) is not None
    except ImportError:
        return False


if __name__ == "__main__":
    if running_in_streamlit():
        render_dashboard()
    elif "--self-test" in sys.argv:
        self_test()
    else:
        print("이 파일은 웹앱 전용입니다. 로컬 수집·공개파일 생성은 상위 폴더에서 python citygas_run.py 를 실행하세요.")

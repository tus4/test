"""Wide daily time-series sheets (market-data exports) -> compact per-series JSON.

A sheet is treated as a time series when a header row whose first cell is a date
label (``DATES``, ``일자``, ...) is followed by rows carrying strictly ordered dates
in column A and mostly numeric cells elsewhere. Optional metadata rows above the
header (``Category``, ``Ticker``, ``Notation``) describe each column; one label
row directly below the header (e.g. ``3월이하(당일)``) is kept as the "field".

The whole history of every series is published (that is the point of a market
dashboard), so the caps that apply to ordinary tables (``max_rows`` ...) do not
apply here; ``include_columns`` / ``exclude_columns`` still do.
"""

from __future__ import annotations

import bisect
import calendar
import datetime as dt
import fnmatch
import hashlib
import math
import re
import statistics
from collections import Counter
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .analyze import parse_date_text, parse_numeric_text
from .common import fingerprint

DETECT_ROWS = 60           # physical rows inspected to detect the layout
MIN_DATE_ROWS = 5          # dated rows needed (inside the detection window) to call it a time series
MIN_NUMERIC_RATIO = 0.6    # share of non-blank data cells that must be numeric
MAX_LABEL_ROWS = 2         # label rows tolerated between the header and the first dated row

DATE_LABELS = {"dates", "date", "datetime", "day", "일자", "날짜", "기준일", "기준일자", "일시"}
META_LABELS = {
    "category": "category", "분류": "category", "카테고리": "category", "구분": "category",
    "ticker": "ticker", "티커": "ticker", "code": "ticker", "코드": "ticker",
    "notation": "notation", "표기": "notation", "name": "notation", "이름": "notation", "지표명": "notation",
}
CATEGORY_LABELS = {
    "RATE": "금리", "RATES": "금리", "CREDIT": "크레딧", "FX": "환율", "CURRENCY": "환율", "HEDGE": "헤지",
    "MACRO": "매크로", "DEMAND": "수급", "FLOW": "수급", "EQUITY": "주식", "INDEX": "지수", "COMMODITY": "원자재",
    "SPREAD": "스프레드", "VOL": "변동성",
}
# unit: "pct" = percent yield (changes in bp) · "bp" = already basis points · "level" = price/FX level
# (changes in % and absolute) · "pts" = index points · "amount" = flows (sums instead of changes)
UNIT_BY_CATEGORY = {
    "rate": "pct", "rates": "pct", "credit": "pct", "hedge": "pct", "swap": "pct", "yield": "pct",
    "fx": "level", "currency": "level", "commodity": "level", "price": "level",
    "macro": "pts", "index": "pts", "equity": "pts", "vol": "pts", "volatility": "pts",
    "demand": "amount", "flow": "amount", "flows": "amount", "volume": "amount", "supply": "amount",
    "spread": "bp",
}
UNIT_LABELS = {"pct": "%", "bp": "bp", "level": "", "pts": "pt", "amount": ""}
SPREAD_TOKENS = {"sp", "oas", "spread", "스프레드"}
HORIZONS = (("w1", "1주"), ("m1", "1개월"), ("m3", "3개월"), ("ytd", "연초 대비"), ("y1", "1년"))
SNAPSHOT_HORIZONS = (("latest", "최근"), ("w1", "1주 전"), ("m1", "1개월 전"), ("y1", "1년 전"))
TENOR_RE = re.compile(r"^(.*?)[\s_]*(\d+(?:\.\d+)?)\s*([YyMmWwDd])$")  # "-" is not a separator: AA- 1y
_HANGUL = re.compile(r"[가-힣]")
_SLUG_STRIP = re.compile(r"[^A-Za-z0-9]+")


# ---------------------------------------------------------------------------
# cell helpers
# ---------------------------------------------------------------------------
def _cell(row: Optional[Sequence[Any]], j: int) -> Any:
    if row is None or j >= len(row):
        return None
    return row[j]


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _label(value: Any) -> str:
    return value.strip().lower() if isinstance(value, str) else ""


def _as_date(value: Any) -> Optional[dt.date]:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        parsed = parse_date_text(value)
        return parsed.date() if parsed else None
    return None


def _numeric(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        f = float(value)
        return f if math.isfinite(f) else None
    if isinstance(value, str):
        parsed = parse_numeric_text(value)
        return parsed[0] if parsed else None
    return None


def _name_matches(name: str, patterns: Sequence[str]) -> bool:
    low = name.lower()
    return any(fnmatch.fnmatchcase(low, p.lower()) for p in patterns)


def _has_hangul(text: str) -> bool:
    return bool(_HANGUL.search(text))


# ---------------------------------------------------------------------------
# detection
# ---------------------------------------------------------------------------
def detect_timeseries(head: Sequence[Sequence[Any]]) -> Optional[int]:
    """Index (within ``head``) of the header row of a wide time-series sheet, else None."""
    for i, row in enumerate(head):
        if _label(_cell(row, 0)) not in DATE_LABELS:
            continue
        if sum(1 for c in row[1:] if not _is_blank(c)) < 2:
            continue
        dates: List[dt.date] = []
        numeric = other = label_rows = 0
        per_column: Dict[int, List[int]] = {}  # column -> [numeric, other]
        for r in head[i + 1:]:
            a = _cell(r, 0)
            if _is_blank(a):
                continue
            d = _as_date(a)
            if d is None:
                if dates or label_rows >= MAX_LABEL_ROWS:
                    break
                label_rows += 1
                continue
            dates.append(d)
            for j, c in enumerate(r[1:], start=1):
                if _is_blank(c):
                    continue
                tally = per_column.setdefault(j, [0, 0])
                if _numeric(c) is not None:
                    numeric += 1
                    tally[0] += 1
                else:
                    other += 1
                    tally[1] += 1
        if len(dates) < MIN_DATE_ROWS:
            continue
        ascending = all(dates[k] < dates[k + 1] for k in range(len(dates) - 1))
        descending = all(dates[k] > dates[k + 1] for k in range(len(dates) - 1))
        if not (ascending or descending):
            continue
        if numeric == 0 or numeric < MIN_NUMERIC_RATIO * (numeric + other):
            continue
        # a column that is mostly text (region, product, memo ...) makes this a table sorted by date, not a time series
        if any(t[1] >= 3 and t[1] > t[0] for t in per_column.values()):
            continue
        return i
    return None


# ---------------------------------------------------------------------------
# extraction of one sheet
# ---------------------------------------------------------------------------
def extract_timeseries_sheet(ws: Any, sheet_name: str, cfg: Dict[str, Any], head: Sequence[Sequence[Any]], header_idx: int) -> Dict[str, Any]:
    """Read every dated row of ``ws``; returns dates + one column dict per series (values aligned to dates)."""
    header = list(head[header_idx])
    meta: Dict[str, Sequence[Any]] = {}
    for r in head[:header_idx]:
        key = META_LABELS.get(_label(_cell(r, 0)))
        if key and key not in meta:
            meta[key] = r
    field_row: Optional[Sequence[Any]] = None
    for r in head[header_idx + 1: header_idx + 1 + MAX_LABEL_ROWS]:
        a = _cell(r, 0)
        if _is_blank(a) or _as_date(a) is not None:
            break
        field_row = r
        break

    width = max([len(header), len(field_row or ())] + [len(r) for r in meta.values()])
    columns: List[Dict[str, Any]] = []
    excluded: List[str] = []
    for j in range(1, width):
        h = _text(_cell(header, j))
        ticker = _text(_cell(meta.get("ticker"), j))
        notation = _text(_cell(meta.get("notation"), j))
        if not (h or ticker or notation):
            continue
        name = h or notation or ticker
        if cfg["include_columns"] and not _name_matches(name, cfg["include_columns"]):
            excluded.append(name)
            continue
        if _name_matches(name, cfg["exclude_columns"]):
            excluded.append(name)
            continue
        columns.append(
            {
                "col": j,
                "header": name,
                "ticker": ticker,
                "notation": notation,
                "category": _text(_cell(meta.get("category"), j)),
                "field": _text(_cell(field_row, j)),
            }
        )
    max_series = cfg["timeseries"]["max_series"]
    truncated_columns = max(0, len(columns) - max_series)
    columns = columns[:max_series]

    by_date: Dict[dt.date, List[Optional[float]]] = {}
    cols = [c["col"] for c in columns]
    scanned = 0
    max_scan = cfg["max_scan_rows"]
    scan_truncated = False
    for idx, row in enumerate(ws.iter_rows(values_only=True)):
        if idx <= header_idx:
            continue
        a = _cell(row, 0)
        if _is_blank(a):
            continue
        d = _as_date(a)
        if d is None:
            continue  # label rows and stray text
        scanned += 1
        if scanned > max_scan:
            scan_truncated = True
            break
        by_date[d] = [_numeric(_cell(row, j)) for j in cols]  # a repeated date keeps its last row
    dates = sorted(by_date)
    for k, column in enumerate(columns):
        column["values"] = [by_date[d][k] for d in dates]
    columns = [c for c in columns if any(v is not None for v in c["values"])]

    notes: List[str] = []
    if excluded:
        notes.append(f"설정으로 제외한 지표 {len(excluded)}개: " + ", ".join(excluded[:10]) + (" …" if len(excluded) > 10 else ""))
    if truncated_columns:
        notes.append(f"지표 수 상한(max_series={max_series})을 넘는 {truncated_columns}개 열은 공개하지 않습니다.")
    if scan_truncated:
        notes.append(f"처음 {max_scan:,}개 날짜 행만 읽었습니다 (max_scan_rows).")
    iso_dates = [d.isoformat() for d in dates]
    return {
        "name": sheet_name,
        "header_row": header_idx + 1,
        "dates": iso_dates,
        "columns": columns,
        "excluded": len(excluded),
        "notes": notes,
        "fingerprint": fingerprint([iso_dates, [[c["header"], c["values"]] for c in columns]]),
    }


def timeseries_stub(sheet: Dict[str, Any]) -> Dict[str, Any]:
    """The entry an extracted time-series sheet gets in the ordinary (table) dataset."""
    n_series, n_dates = len(sheet["columns"]), len(sheet["dates"])
    span = f" ({sheet['dates'][0]} ~ {sheet['dates'][-1]})" if n_dates else ""
    return {
        "name": sheet["name"],
        "state": "timeseries",
        "title": "",
        "header_row": sheet["header_row"],
        "rows_total": n_dates,
        "rows_scanned": n_dates,
        "rows_published": 0,
        "scan_truncated": False,
        "columns_total": n_series + sheet["excluded"],
        "columns_published": n_series,
        "columns_truncated": 0,
        "columns_excluded": [],
        "cells_truncated": 0,
        "formula_cells": 0,
        "formula_cells_without_cached_value": 0,
        "columns": [],
        "rows": [],
        "charts": [],
        "notes": [f"시계열 데이터입니다 (지표 {n_series}개 · 날짜 {n_dates:,}개{span}). 시계열 대시보드(index.html)에서 표시됩니다."] + list(sheet["notes"]),
        "fingerprint": sheet["fingerprint"],
    }


# ---------------------------------------------------------------------------
# naming, units, cleaning
# ---------------------------------------------------------------------------
def _slug(text: str) -> str:
    return _SLUG_STRIP.sub("-", text).strip("-").lower()[:40]


def series_id(header: str, ticker: str, notation: str) -> str:
    digest = hashlib.sha1(f"{header}|{ticker}|{notation}".encode("utf-8")).hexdigest()[:6]
    return f"{_slug(notation or header) or 's'}-{digest}"


def display_name(header: str, ticker: str, notation: str) -> str:
    """Korean ticker label when there is one, else the header. A tenor in the label that
    contradicts the header's tenor (a copy-paste slip in the source) is replaced by the header's."""
    base = ticker if ticker and _has_hangul(ticker) else (header or notation or ticker)
    m_base, m_head = TENOR_RE.match(base), TENOR_RE.match(header)
    if m_base and m_head and (m_base.group(2), m_base.group(3).upper()) != (m_head.group(2), m_head.group(3).upper()):
        base = f"{m_base.group(1)}_{m_head.group(2)}{m_head.group(3)}"
    return base.replace("_", " ").strip() or header


def infer_unit(category: str, header: str, notation: str, ticker: str, field: str) -> str:
    tokens = set()
    for text in (header, notation, ticker, field):
        tokens.update(t for t in re.split(r"[\s_\-/().]+", text.lower()) if t)
    if tokens & SPREAD_TOKENS:
        return "bp"
    return UNIT_BY_CATEGORY.get(category.strip().lower(), "level")


def infer_decimals(values: Sequence[Optional[float]], unit: str) -> int:
    if unit == "amount":
        return 0
    tail = [v for v in values[-250:] if v is not None]
    best = 0
    for v in tail:
        s = repr(v)
        if "e" in s or "E" in s:
            continue
        if "." in s:
            best = max(best, len(s.split(".")[1].rstrip("0")))
    return min(best, 4)


def clean_series(values: Sequence[Optional[float]], unit: str) -> Tuple[List[Optional[float]], List[int]]:
    """Blank the zeros that mark missing quotes and isolated one-day spikes. Returns (values, indices)."""
    vals = list(values)
    removed: List[int] = []
    if unit == "amount":
        return vals, removed
    nonzero = [i for i, v in enumerate(vals) if v is not None and v != 0]
    if not nonzero:
        return vals, removed
    price_like = unit in ("level", "pts")
    first_nz = nonzero[0]
    for i in range(first_nz):
        if vals[i] == 0 and (price_like or abs(vals[first_nz]) >= 0.5):
            vals[i] = None
            removed.append(i)
    n = len(vals)
    next_nz: List[Optional[float]] = [None] * n
    right: Optional[float] = None
    for i in range(n - 1, -1, -1):
        next_nz[i] = right
        if vals[i] is not None and vals[i] != 0:
            right = vals[i]
    left: Optional[float] = None
    for i in range(n):
        v = vals[i]
        if v is None:
            continue
        if v != 0:
            left = v
            continue
        if i < first_nz:
            continue
        r = next_nz[i]
        if price_like or (left is not None and abs(left) >= 0.5 and (r is None or abs(r) >= 0.5)):
            vals[i] = None
            removed.append(i)
    # isolated spikes: far from a consistent neighbourhood
    for i in range(n):
        v = vals[i]
        if v is None or v == 0:
            continue
        nb = [vals[k] for k in range(max(0, i - 5), min(n, i + 6)) if k != i and vals[k] not in (None, 0)]
        if len(nb) < 4:
            continue
        mags = [abs(x) for x in nb]
        if max(mags) / min(mags) > 1.5 or len({x > 0 for x in nb}) != 1:
            continue
        med = statistics.median(nb)
        ratio = v / med
        if (ratio > 5 or ratio < 0.2) and abs(v - med) > 0.5:
            vals[i] = None
            removed.append(i)
    return vals, sorted(set(removed))


# ---------------------------------------------------------------------------
# statistics
# ---------------------------------------------------------------------------
def shift_months(d: dt.date, months: int) -> dt.date:
    m = d.month - 1 - months
    y = d.year + m // 12
    m = m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def index_at_or_before(dates: Sequence[dt.date], vals: Sequence[Optional[float]], target: dt.date) -> int:
    k = bisect.bisect_right(dates, target) - 1
    while k >= 0 and vals[k] is None:
        k -= 1
    return k


def _point(dates: Sequence[dt.date], vals: Sequence[Optional[float]], k: int) -> Optional[Dict[str, Any]]:
    if k < 0:
        return None
    return {"i": k, "date": dates[k].isoformat(), "value": vals[k]}


def series_stats(dates: Sequence[dt.date], vals: Sequence[Optional[float]], unit: str, spark_points: int) -> Dict[str, Any]:
    last = len(vals) - 1
    while last >= 0 and vals[last] is None:
        last -= 1
    first = 0
    while first <= last and vals[first] is None:
        first += 1
    if last < 0:
        return {"first": None, "last": None, "count": 0, "latest": None}
    latest_date = dates[last]
    prev = last - 1
    while prev >= 0 and vals[prev] is None:
        prev -= 1
    out: Dict[str, Any] = {
        "first": first,
        "last": last,
        "count": sum(1 for v in vals[first: last + 1] if v is not None),
        "latest": _point(dates, vals, last),
    }
    if unit == "amount":
        obs = [(dates[k], vals[k]) for k in range(first, last + 1) if vals[k] is not None]
        out["sums"] = {
            "d1": vals[last],
            "d5": sum(v for _, v in obs[-5:]),
            "d20": sum(v for _, v in obs[-20:]),
            "mtd": sum(v for d, v in obs if (d.year, d.month) == (latest_date.year, latest_date.month)),
            "ytd": sum(v for d, v in obs if d.year == latest_date.year),
        }
    else:
        bases: Dict[str, Optional[Dict[str, Any]]] = {"d1": _point(dates, vals, prev)}
        targets = {
            "w1": latest_date - dt.timedelta(days=7),
            "m1": shift_months(latest_date, 1),
            "m3": shift_months(latest_date, 3),
            "ytd": dt.date(latest_date.year - 1, 12, 31),
            "y1": shift_months(latest_date, 12),
        }
        for key, target in targets.items():
            k = index_at_or_before(dates, vals, target)
            bases[key] = _point(dates, vals, k) if k >= 0 and k < last else None
        out["bases"] = bases
    year_ago = shift_months(latest_date, 12)
    start52 = bisect.bisect_left(dates, year_ago)
    window = [(vals[k], k) for k in range(max(start52, first), last + 1) if vals[k] is not None]
    if window:
        lo, hi = min(window), max(window)
        out["range52"] = {"lo": _point(dates, vals, lo[1]), "hi": _point(dates, vals, hi[1])}
    everything = [(vals[k], k) for k in range(first, last + 1) if vals[k] is not None]
    lo, hi = min(everything), max(everything)
    out["extremes"] = {"lo": _point(dates, vals, lo[1]), "hi": _point(dates, vals, hi[1])}
    spark_src = [k for _, k in window] if window else [k for _, k in everything[-spark_points:]]
    if len(spark_src) > spark_points:
        step = len(spark_src) / float(spark_points)
        picked = sorted({spark_src[int(j * step)] for j in range(spark_points)} | {spark_src[-1]})
    else:
        picked = spark_src
    out["spark"] = [[k, vals[k]] for k in picked]
    return out


# ---------------------------------------------------------------------------
# curves (tenor families)
# ---------------------------------------------------------------------------
def parse_tenor(header: str) -> Optional[Tuple[str, float, str]]:
    m = TENOR_RE.match(header.strip())
    if not m:
        return None
    n, u = float(m.group(2)), m.group(3).upper()
    years = {"Y": n, "M": n / 12.0, "W": n / 52.0, "D": n / 365.0}[u]
    label = (str(int(n)) if n.is_integer() else f"{n:g}") + u
    return m.group(1).strip(" _-").upper(), years, label


def build_curves(series: List[Dict[str, Any]], dates: Sequence[dt.date], values_by_id: Dict[str, List[Optional[float]]], asof: dt.date) -> List[Dict[str, Any]]:
    families: Dict[Tuple[str, str, str], List[Tuple[float, str, Dict[str, Any]]]] = {}
    for s in series:
        parsed = parse_tenor(s["header"])
        if not parsed or s["unit"] == "amount":
            continue
        prefix, years, label = parsed
        families.setdefault((prefix, s["category"], s["unit"]), []).append((years, label, s))
    curves: List[Dict[str, Any]] = []
    targets = {"latest": asof, "w1": asof - dt.timedelta(days=7), "m1": shift_months(asof, 1), "y1": shift_months(asof, 12)}
    for (prefix, category, unit), members in families.items():
        members.sort(key=lambda t: t[0])
        unique: List[Tuple[float, str, Dict[str, Any]]] = []
        seen = set()
        for years, label, s in members:
            if years in seen:
                continue
            seen.add(years)
            unique.append((years, label, s))
        if len(unique) < 3:
            continue
        stems = []
        for _, _, s in unique:
            m = TENOR_RE.match(s["name"])
            stems.append(m.group(1).strip(" _-") if m and m.group(1).strip(" _-") else s["name"])
        label = Counter(stems).most_common(1)[0][0] or prefix
        snapshots = []
        for key, hlabel in SNAPSHOT_HORIZONS:
            vals = []
            snap_dates = []
            for _, _, s in unique:
                k = index_at_or_before(dates, values_by_id[s["id"]], targets[key])
                vals.append(values_by_id[s["id"]][k] if k >= 0 else None)
                if k >= 0:
                    snap_dates.append(dates[k])
            snapshots.append({"key": key, "label": hlabel, "date": max(snap_dates).isoformat() if snap_dates else None, "values": vals})
        curves.append(
            {
                "id": f"curve-{_slug(prefix) or 'x'}-{hashlib.sha1(f'{prefix}|{category}|{unit}'.encode('utf-8')).hexdigest()[:6]}",
                "label": label,
                "prefix": prefix,
                "category": category,
                "unit": unit,
                "tenors": [{"id": s["id"], "label": tl, "years": years, "name": s["name"]} for years, tl, s in unique],
                "snapshots": snapshots,
            }
        )
    counts = Counter(c["label"] for c in curves)
    for c in curves:
        if counts[c["label"]] > 1:
            c["label"] = f"{c['label']} ({c['prefix']})"
    order = {s["id"]: n for n, s in enumerate(series)}
    curves.sort(key=lambda c: order[c["tenors"][0]["id"]])
    return curves


# ---------------------------------------------------------------------------
# the market dataset
# ---------------------------------------------------------------------------
def _match_series(pattern: str, series: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    p = pattern.strip().lower()
    if not p:
        return None
    for key in ("header", "notation", "id", "name", "ticker"):
        for s in series:
            if s.get(key, "").lower() == p:
                return s
    for s in series:
        if s["name"].replace(" ", "_").lower() == p.replace(" ", "_"):
            return s
    return None


def _resolve_list(patterns: Sequence[str], series: List[Dict[str, Any]]) -> Tuple[List[str], List[str]]:
    ids: List[str] = []
    missing: List[str] = []
    for p in patterns:
        s = _match_series(p, series)
        if s is None:
            missing.append(p)
        elif s["id"] not in ids:
            ids.append(s["id"])
    return ids, missing


def _auto_highlights(series: List[Dict[str, Any]], limit: int = 8) -> List[str]:
    by_cat: Dict[str, List[Dict[str, Any]]] = {}
    for s in series:
        by_cat.setdefault(s["category"], []).append(s)
    picked: List[str] = []
    depth = 0
    while len(picked) < limit:
        added = False
        for cat in by_cat:
            if depth < len(by_cat[cat]) and len(picked) < limit:
                picked.append(by_cat[cat][depth]["id"])
                added = True
        if not added:
            break
        depth += 1
    return picked


def build_market(ts_sheets: List[Dict[str, Any]], cfg: Dict[str, Any]) -> Dict[str, Any]:
    """Merge the extracted time-series sheets into one dataset (union of dates, one entry per series)."""
    tcfg = cfg["timeseries"]
    all_dates = sorted({d for sheet in ts_sheets for d in sheet["dates"]})
    dates = [dt.date.fromisoformat(d) for d in all_dates]
    pos = {d: i for i, d in enumerate(all_dates)}
    n = len(dates)

    series: List[Dict[str, Any]] = []
    values_by_id: Dict[str, List[Optional[float]]] = {}
    ids_seen: Dict[str, int] = {}
    cleaned_total = 0
    sources = []
    for sheet in ts_sheets:
        sheet_pos = [pos[d] for d in sheet["dates"]]
        for column in sheet["columns"]:
            aligned: List[Optional[float]] = [None] * n
            for k, v in zip(sheet_pos, column["values"]):
                aligned[k] = v
            unit = infer_unit(column["category"], column["header"], column["notation"], column["ticker"], column["field"])
            removed: List[int] = []
            if tcfg["clean_outliers"]:
                aligned, removed = clean_series(aligned, unit)
            if all(v is None for v in aligned):
                continue
            sid = series_id(column["header"], column["ticker"], column["notation"])
            if sid in ids_seen:
                ids_seen[sid] += 1
                sid = f"{sid}-{ids_seen[sid]}"
            else:
                ids_seen[sid] = 1
            stats = series_stats(dates, aligned, unit, tcfg["spark_points"])
            category = column["category"].strip().upper() or "기타"
            entry: Dict[str, Any] = {
                "id": sid,
                "header": column["header"],
                "code": column["notation"] or column["header"],
                "name": display_name(column["header"], column["ticker"], column["notation"]),
                "ticker": column["ticker"],
                "notation": column["notation"],
                "field": column["field"],
                "category": category,
                "unit": unit,
                "decimals": infer_decimals(aligned, unit),
                "source": {"file": sheet.get("file", ""), "sheet": sheet["name"]},
                "cleaned": len(removed),
                "cleaned_dates": [all_dates[i] for i in removed[:5]],
            }
            entry.update(stats)
            cleaned_total += len(removed)
            series.append(entry)
            values_by_id[sid] = aligned
        sources.append(
            {
                "file": sheet.get("file", ""),
                "sheet": sheet["name"],
                "header_row": sheet["header_row"],
                "series": len(sheet["columns"]),
                "dates": len(sheet["dates"]),
                "first_date": sheet["dates"][0] if sheet["dates"] else None,
                "last_date": sheet["dates"][-1] if sheet["dates"] else None,
                "notes": list(sheet["notes"]),
            }
        )
    names = Counter(s["name"] for s in series)
    for s in series:
        if names[s["name"]] > 1:
            s["name"] = f"{s['name']} ({s['header']})"

    asof = max((dt.date.fromisoformat(s["latest"]["date"]) for s in series), default=None)
    curves = build_curves(series, dates, values_by_id, asof) if asof else []

    highlights, missing_h = _resolve_list(tcfg["highlights"], series)
    if not highlights:
        highlights = _auto_highlights(series)
    featured: List[Dict[str, Any]] = []
    missing_f: List[str] = []
    for group in tcfg["featured"]:
        if isinstance(group, dict):
            title, patterns, mode = group.get("title", ""), group.get("series", []), group.get("mode")
        else:
            title, patterns, mode = "", group, None
        ids, missing = _resolve_list(patterns, series)
        missing_f += missing
        if ids:
            featured.append({"title": title, "series": ids[:6], "mode": mode})
    if not featured:
        by_unit: Dict[str, List[str]] = {}
        for sid in highlights:
            s = next(x for x in series if x["id"] == sid)
            if s["unit"] != "amount":
                by_unit.setdefault(s["unit"], []).append(sid)
        featured = [{"title": "", "series": ids[:4], "mode": None} for ids in by_unit.values()][:4]

    notes: List[str] = []
    if missing_h:
        notes.append("설정의 highlights 중 찾지 못한 지표: " + ", ".join(missing_h))
    if missing_f:
        notes.append("설정의 featured 중 찾지 못한 지표: " + ", ".join(missing_f))
    cats: List[Dict[str, Any]] = []
    for s in series:
        for c in cats:
            if c["key"] == s["category"]:
                c["count"] += 1
                break
        else:
            cats.append({"key": s["category"], "label": CATEGORY_LABELS.get(s["category"], s["category"]), "count": 1})

    market = {
        "schema_version": 1,
        "title": cfg["title"],
        "asof": asof.isoformat() if asof else None,
        "first_date": all_dates[0] if all_dates else None,
        "last_date": all_dates[-1] if all_dates else None,
        "dates": all_dates,
        "sources": sources,
        "categories": cats,
        "series": series,
        "curves": curves,
        "highlights": highlights,
        "featured": featured,
        "cleaning": {"enabled": bool(tcfg["clean_outliers"]), "points": cleaned_total},
        "notes": notes,
        "units": UNIT_LABELS,
    }
    market["fingerprint"] = fingerprint([[s["id"], s["last"], s["count"], s["latest"]] for s in series] + [all_dates[-1:] or None])
    market["_values"] = values_by_id  # stripped before serialisation; written to data/series/<id>.json
    return market


def series_payload(market: Dict[str, Any], entry: Dict[str, Any]) -> Dict[str, Any]:
    vals = market["_values"][entry["id"]]
    first, last = entry["first"], entry["last"]
    return {"id": entry["id"], "start": first, "values": vals[first: last + 1]}


def market_summary(market: Dict[str, Any]) -> Dict[str, Any]:
    """The compact part of the market dataset that goes into dashboard.json and the README block."""
    by_id = {s["id"]: s for s in market["series"]}
    keys = ("id", "name", "code", "category", "unit", "decimals", "latest", "bases", "sums", "cleaned")
    return {
        "asof": market["asof"],
        "first_date": market["first_date"],
        "last_date": market["last_date"],
        "dates": len(market["dates"]),
        "series_count": len(market["series"]),
        "categories": market["categories"],
        "sources": market["sources"],
        "highlights": [{k: by_id[i][k] for k in keys if k in by_id[i]} for i in market["highlights"]],
        "curves": [{"label": c["label"], "tenors": [t["label"] for t in c["tenors"]]} for c in market["curves"]],
        "cleaning": market["cleaning"],
        "notes": market["notes"],
        "fingerprint": market["fingerprint"],
    }


def public_market(market: Dict[str, Any]) -> Dict[str, Any]:
    return {k: v for k, v in market.items() if not k.startswith("_")}

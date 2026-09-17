"""Loading and validating the optional ``dashboard.config.json`` from the data repo.

The config controls WHAT gets published to the public repository. Unknown keys
are rejected on purpose: a misspelled ``exclude_columns`` must not silently
publish more than intended.
"""

from __future__ import annotations

import copy
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from .common import BuildError

CONFIG_FILENAME = "dashboard.config.json"

# Header substrings (case-insensitive) that mark a column as personal data.
# Such columns are dropped automatically unless listed in ``allow_columns``.
DEFAULT_SENSITIVE_PATTERNS: List[str] = [
    "주민", "여권", "전화", "휴대폰", "핸드폰", "연락처", "이메일", "메일주소", "주소", "계좌", "카드번호",
    "비밀번호", "생년월일", "email", "e-mail", "phone", "mobile", "address", "account no", "card number",
    "password", "passwd", "ssn", "passport", "birth",
]

DEFAULTS: Dict[str, Any] = {
    # "rows": publish row data (capped) + aggregates. "aggregates": KPIs and charts only.
    "mode": "rows",
    "title": "데이터 대시보드",
    # File selection: glob patterns matched against the path relative to the data repo root.
    "include_files": ["**/*.xlsx", "**/*.xlsm"],
    "exclude_files": [],
    # Sheet selection: glob patterns matched against the sheet name (case-insensitive).
    "include_sheets": [],
    "exclude_sheets": [],
    # Column selection: glob patterns matched against the header text (case-insensitive).
    "include_columns": [],
    "exclude_columns": [],
    # Automatic exclusion of personal-data columns (by header substring or by value shape:
    # e-mail / phone number / resident registration number). ``allow_columns`` overrides it.
    "auto_exclude_sensitive": True,
    "sensitive_patterns": DEFAULT_SENSITIVE_PATTERNS,
    "allow_columns": [],
    # Caps. Every truncation is stated in the output.
    "max_rows": 1000,          # rows per sheet written to the public JSON
    "max_columns": 50,         # columns per sheet
    "max_scan_rows": 200000,   # rows per sheet used for aggregates / charts
    "max_cell_text": 200,      # characters kept per published text cell
    "preview_rows": 10,        # rows shown in the README preview table
    "preview_columns": 12,     # columns shown in the README preview table
    "header_scan_rows": 20,    # rows inspected to locate the header row
    "max_chart_categories": 12,
    # Aggregates: charts (category labels are published) and top-value lists of text columns
    # (only for columns with <= 50 distinct values, only in "rows" mode).
    "charts": True,
    "publish_top_values": True,
    "readme": {"charts": True, "xychart": False, "preview": True},
    # Fail the run (keeping the old dashboard) when no workbook matches.
    "fail_if_no_workbooks": True,
    # Wide daily time-series sheets (a "DATES"/"일자" header row followed by dated rows, optional
    # Category/Ticker/Notation rows above it) are published as a market dashboard: every series in full.
    "timeseries": {
        "enabled": True,
        "highlights": [],       # series (header / notation / name) shown as KPI tiles, in this order
        "featured": [],         # overview charts: [["UST2Y", "UST10Y"], {"title": "...", "series": [...]}]
        "clean_outliers": True, # blank zeros that mark missing quotes and isolated one-day spikes
        "spark_points": 52,     # points per sparkline (last year)
        "max_series": 1000,     # series per sheet
    },
}

_INT_KEYS = (
    "max_rows",
    "max_columns",
    "max_scan_rows",
    "max_cell_text",
    "preview_rows",
    "preview_columns",
    "header_scan_rows",
    "max_chart_categories",
)
_LIST_KEYS = (
    "include_files",
    "exclude_files",
    "include_sheets",
    "exclude_sheets",
    "include_columns",
    "exclude_columns",
    "sensitive_patterns",
    "allow_columns",
)
_BOOL_KEYS = ("auto_exclude_sensitive", "charts", "publish_top_values", "fail_if_no_workbooks")
_INT_MIN = {"max_columns": 1, "max_cell_text": 1, "header_scan_rows": 1, "max_chart_categories": 2}


def _fail(source: str, message: str) -> None:
    raise BuildError(f"{source}: {message}")


def validate_config(raw: Any, source: str = CONFIG_FILENAME) -> Dict[str, Any]:
    """Merge ``raw`` over DEFAULTS after strict validation. Keys starting with ``_`` are comments."""
    if not isinstance(raw, dict):
        _fail(source, "the top level must be a JSON object")
    raw = {k: v for k, v in raw.items() if not (isinstance(k, str) and k.startswith("_"))}
    unknown = sorted(k for k in raw if k not in DEFAULTS)
    if unknown:
        _fail(
            source,
            "unknown key(s) " + ", ".join(unknown) + ". Allowed keys: " + ", ".join(sorted(DEFAULTS)),
        )
    cfg = copy.deepcopy(DEFAULTS)
    for key, value in raw.items():
        if key == "mode":
            if value not in ("rows", "aggregates"):
                _fail(source, "\"mode\" must be \"rows\" or \"aggregates\"")
        elif key == "title":
            if not isinstance(value, str) or not value.strip():
                _fail(source, "\"title\" must be a non-empty string")
        elif key in _INT_KEYS:
            if isinstance(value, bool) or not isinstance(value, int) or value < _INT_MIN.get(key, 0):
                _fail(source, f"\"{key}\" must be an integer >= {_INT_MIN.get(key, 0)}")
        elif key in _LIST_KEYS:
            if not isinstance(value, list) or not all(isinstance(v, str) and v for v in value):
                _fail(source, f"\"{key}\" must be a list of non-empty strings")
        elif key == "readme":
            if not isinstance(value, dict):
                _fail(source, "\"readme\" must be an object")
            bad = sorted(k for k in value if k not in DEFAULTS["readme"])
            if bad:
                _fail(source, "unknown \"readme\" key(s): " + ", ".join(bad))
            for k, v in value.items():
                if not isinstance(v, bool):
                    _fail(source, f"\"readme.{k}\" must be true or false")
            cfg["readme"].update(value)
            continue
        elif key in _BOOL_KEYS:
            if not isinstance(value, bool):
                _fail(source, f"\"{key}\" must be true or false")
        elif key == "timeseries":
            cfg["timeseries"].update(_validate_timeseries(value, source))
            continue
        cfg[key] = copy.deepcopy(value)
    if cfg["preview_rows"] > cfg["max_rows"] and cfg["mode"] == "rows":
        cfg["preview_rows"] = cfg["max_rows"]
    return cfg


def _validate_timeseries(value: Any, source: str) -> Dict[str, Any]:
    if not isinstance(value, dict):
        _fail(source, "\"timeseries\" must be an object")
    defaults = DEFAULTS["timeseries"]
    bad = sorted(k for k in value if k not in defaults)
    if bad:
        _fail(source, "unknown \"timeseries\" key(s): " + ", ".join(bad) + ". Allowed keys: " + ", ".join(sorted(defaults)))
    out: Dict[str, Any] = {}
    for k, v in value.items():
        if k in ("enabled", "clean_outliers"):
            if not isinstance(v, bool):
                _fail(source, f"\"timeseries.{k}\" must be true or false")
        elif k in ("spark_points", "max_series"):
            if isinstance(v, bool) or not isinstance(v, int) or v < (5 if k == "spark_points" else 1):
                _fail(source, f"\"timeseries.{k}\" must be an integer >= {5 if k == 'spark_points' else 1}")
        elif k == "highlights":
            if not isinstance(v, list) or not all(isinstance(x, str) and x.strip() for x in v):
                _fail(source, "\"timeseries.highlights\" must be a list of non-empty strings")
        elif k == "featured":
            if not isinstance(v, list):
                _fail(source, "\"timeseries.featured\" must be a list")
            for group in v:
                if isinstance(group, dict):
                    extra = sorted(g for g in group if g not in ("title", "series", "mode"))
                    if extra or not isinstance(group.get("title", ""), str) or group.get("mode", "level") not in ("level", "diff", "rebase"):
                        _fail(source, "\"timeseries.featured\" objects allow only \"title\" (string), \"series\" (list) and \"mode\" (level | diff | rebase)")
                    names = group.get("series")
                else:
                    names = group
                if not isinstance(names, list) or not names or not all(isinstance(x, str) and x.strip() for x in names):
                    _fail(source, "\"timeseries.featured\" entries must be non-empty lists of series names")
        out[k] = copy.deepcopy(v)
    return out


def load_config(path: Optional[Path]) -> Dict[str, Any]:
    """Return the effective config. A missing file means DEFAULTS."""
    if path is None or not path.exists():
        cfg = copy.deepcopy(DEFAULTS)
        cfg["_source"] = None
        return cfg
    try:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError) as exc:
        raise BuildError(f"{path.name}: cannot read file: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise BuildError(f"{path.name}: not valid JSON (line {exc.lineno}, column {exc.colno}): {exc.msg}") from exc
    cfg = validate_config(raw, source=path.name)
    cfg["_source"] = path.name
    return cfg


def public_config(cfg: Dict[str, Any]) -> Dict[str, Any]:
    """The subset of the config that is safe and useful to publish."""
    keys: List[str] = [
        "mode",
        "max_rows",
        "max_columns",
        "max_scan_rows",
        "max_cell_text",
        "preview_rows",
        "preview_columns",
        "max_chart_categories",
        "auto_exclude_sensitive",
        "charts",
        "publish_top_values",
        "timeseries",
    ]
    return {k: cfg[k] for k in keys}

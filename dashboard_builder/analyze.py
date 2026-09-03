"""Per-column accumulators, type inference and chart aggregation.

Everything here is streaming-friendly: a column accumulator sees one value at a
time and never stores more than a bounded number of distinct values.
"""

from __future__ import annotations

import datetime as dt
import re
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple

from .common import fmt_plain

MAX_DISTINCT_TRACK = 5000
TOP_VALUES = 12
# Top-value lists are published only for real categoricals (few distinct values).
TOP_VALUES_MAX_DISTINCT = 50

# Excel error literals (openpyxl reports them as strings with data_type "e").
EXCEL_ERRORS = {"#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!", "#CALC!", "#GETTING_DATA"}
# Text placeholders that mean "no value"; treated as blank so one "-" does not turn a numeric column into text.
PLACEHOLDERS = {"-", "–", "—", "n/a", "na", "없음", "해당없음", "해당 없음", "미정", "?"}

_NUMERIC_TEXT = re.compile(r"^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?\s*%?$")
_DATE_TEXT = re.compile(r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$")
_KDATE_TEXT = re.compile(r"^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일$")
# Value shapes that mark a text column as personal data when most of its values match.
_PII_SHAPES = (
    re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$"),                    # e-mail
    re.compile(r"^(\+82[- ]?)?0?\d{1,2}[- .]?\d{3,4}[- .]\d{4}$"),       # Korean phone number
    re.compile(r"^\d{6}[- ]?[1-8]\d{6}$"),                               # resident registration number
)


def parse_numeric_text(text: str) -> Optional[Tuple[float, bool]]:
    """Parse '1,234.5' or '15%' -> (value, is_percent). Percent is returned as a fraction."""
    s = text.strip()
    if not s or not _NUMERIC_TEXT.match(s):
        return None
    is_percent = s.endswith("%")
    s = s.rstrip("%").strip().replace(",", "")
    try:
        value = float(s)
    except ValueError:
        return None
    if is_percent:
        value = value / 100.0
    return value, is_percent


def parse_date_text(text: str) -> Optional[dt.datetime]:
    """Parse '2024-01-05', '2024.1.5', '2024/01/05 09:30' or '2024년 1월 5일' -> datetime, else None."""
    s = text.strip()
    m = _DATE_TEXT.match(s) or _KDATE_TEXT.match(s)
    if not m:
        return None
    g = m.groups()
    try:
        if len(g) > 3 and g[3] is not None:
            return dt.datetime(int(g[0]), int(g[1]), int(g[2]), int(g[3]), int(g[4]), int(g[5] or 0))
        return dt.datetime(int(g[0]), int(g[1]), int(g[2]))
    except ValueError:
        return None


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def is_error_text(value: Any) -> bool:
    return isinstance(value, str) and value.strip().upper() in EXCEL_ERRORS


def is_placeholder(value: Any) -> bool:
    return isinstance(value, str) and value.strip().lower() in PLACEHOLDERS


def looks_like_pii(text: str) -> bool:
    s = text.strip()
    return any(rx.match(s) for rx in _PII_SHAPES)


def iso_datetime(value: Any) -> str:
    """ISO 8601: date only when the time part is midnight."""
    if isinstance(value, dt.datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0 and value.microsecond == 0:
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%dT%H:%M:%S")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S")
    return str(value)


def _num_out(value: float) -> Any:
    """Store integral floats as ints and round the rest to avoid float noise."""
    if abs(value - round(value)) < 1e-9 and abs(value) < 1e15:
        return int(round(value))
    return round(value, 6)


class ColumnAcc:
    """Accumulates type evidence and statistics for one column."""

    def __init__(self, name: str, letter: str, index: int) -> None:
        self.name = name
        self.letter = letter
        self.index = index  # 0-based position in the source sheet
        self.n_nonempty = 0
        self.n_num = 0
        self.n_bool = 0
        self.n_date = 0
        self.n_date_with_time = 0
        self.n_time = 0
        self.n_str = 0
        self.n_str_numeric = 0
        self.n_str_percent = 0
        self.n_str_date = 0
        self.n_str_date_with_time = 0
        self.n_pct_format = 0
        self.n_other = 0
        self.n_error = 0        # Excel error cells (#DIV/0!, #N/A, ...), published as null
        self.n_placeholder = 0  # "-", "N/A", ... published as null
        self.n_pii_shape = 0    # text values shaped like e-mail / phone / resident id
        self.num_sum = 0.0
        self.num_min: Optional[float] = None
        self.num_max: Optional[float] = None
        self.date_min: Optional[dt.datetime] = None
        self.date_max: Optional[dt.datetime] = None
        self.counter: Counter = Counter()
        self.distinct_overflow = False

    # -- feeding -----------------------------------------------------------
    def _count_text(self, text: str) -> None:
        if text in self.counter or len(self.counter) < MAX_DISTINCT_TRACK:
            self.counter[text] += 1
        else:
            self.distinct_overflow = True

    def _num(self, value: float) -> None:
        self.num_sum += value
        self.num_min = value if self.num_min is None else min(self.num_min, value)
        self.num_max = value if self.num_max is None else max(self.num_max, value)

    def _date(self, value: dt.datetime) -> None:
        self.date_min = value if self.date_min is None else min(self.date_min, value)
        self.date_max = value if self.date_max is None else max(self.date_max, value)

    def note_error(self) -> None:
        self.n_error += 1

    def note_placeholder(self) -> None:
        self.n_placeholder += 1

    def feed(self, value: Any, number_format: str = "General") -> None:
        if is_blank(value):
            return
        if is_error_text(value):
            self.n_error += 1
            return
        if is_placeholder(value):
            self.n_placeholder += 1
            return
        self.n_nonempty += 1
        if isinstance(value, bool):
            self.n_bool += 1
            self._count_text("TRUE" if value else "FALSE")
        elif isinstance(value, (int, float)):
            self.n_num += 1
            if "%" in (number_format or ""):
                self.n_pct_format += 1
            self._num(float(value))
            self._count_text(fmt_plain(value))
        elif isinstance(value, dt.datetime):
            self.n_date += 1
            if value.hour or value.minute or value.second or value.microsecond:
                self.n_date_with_time += 1
            self._date(value)
            self._count_text(iso_datetime(value))
        elif isinstance(value, dt.date):
            d = dt.datetime(value.year, value.month, value.day)
            self.n_date += 1
            self._date(d)
            self._count_text(iso_datetime(d))
        elif isinstance(value, dt.time):
            self.n_time += 1
            self._count_text(iso_datetime(value))
        elif isinstance(value, str):
            self.n_str += 1
            parsed = parse_numeric_text(value)
            if parsed is not None:
                self.n_str_numeric += 1
                if parsed[1]:
                    self.n_str_percent += 1
                self._num(parsed[0])
                self._count_text(value.strip())
                return
            parsed_date = parse_date_text(value)
            if parsed_date is not None:
                self.n_str_date += 1
                if parsed_date.hour or parsed_date.minute or parsed_date.second:
                    self.n_str_date_with_time += 1
                self._date(parsed_date)
                self._count_text(iso_datetime(parsed_date))
                return
            if looks_like_pii(value):
                self.n_pii_shape += 1
            self._count_text(value.strip())
        else:
            self.n_other += 1
            self._count_text(str(value))

    # -- inference ---------------------------------------------------------
    def inferred_type(self) -> str:
        """One of: empty, number, percent, date, datetime, time, text."""
        if self.n_nonempty == 0:
            return "empty"
        numeric_like = self.n_num + self.n_str_numeric
        others = self.n_bool + self.n_date + self.n_time + self.n_other + (self.n_str - self.n_str_numeric)
        if numeric_like > 0 and others == 0:
            if (self.n_pct_format + self.n_str_percent) * 2 >= numeric_like:
                return "percent"
            return "number"
        date_like = self.n_date + self.n_str_date
        if date_like > 0 and date_like == self.n_nonempty:
            return "datetime" if (self.n_date_with_time or self.n_str_date_with_time) else "date"
        if self.n_time > 0 and self.n_time == self.n_nonempty:
            return "time"
        return "text"

    def looks_sensitive(self) -> bool:
        """True when most values of a text column are shaped like e-mail / phone / resident id."""
        return self.inferred_type() == "text" and self.n_nonempty >= 3 and self.n_pii_shape * 2 >= self.n_nonempty

    def stats(self, kind: str, top_values: bool = True) -> Dict[str, Any]:
        out: Dict[str, Any] = {"count": self.n_nonempty}
        if kind in ("number", "percent"):
            out.update(
                {
                    "sum": _num_out(self.num_sum),
                    "mean": _num_out(self.num_sum / self.n_nonempty),
                    "min": _num_out(self.num_min if self.num_min is not None else 0.0),
                    "max": _num_out(self.num_max if self.num_max is not None else 0.0),
                }
            )
            if self.n_str_numeric:
                out["coerced_from_text"] = self.n_str_numeric
        elif kind in ("date", "datetime"):
            out["min"] = iso_datetime(self.date_min)
            out["max"] = iso_datetime(self.date_max)
            if self.n_str_date:
                out["coerced_from_text"] = self.n_str_date
        elif kind != "empty":
            distinct = len(self.counter)
            out["distinct"] = distinct
            out["distinct_is_lower_bound"] = self.distinct_overflow
            if top_values and not self.distinct_overflow and distinct <= TOP_VALUES_MAX_DISTINCT:
                out["top"] = [[v, c] for v, c in top_values_of(self.counter, TOP_VALUES)]
        return out


def top_values_of(counter: Counter, limit: int) -> List[Tuple[str, int]]:
    """Deterministic top-N: count desc, then value asc."""
    return sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))[:limit]


# Backwards-compatible alias.
top_values = top_values_of


def convert_value(value: Any, kind: str) -> Any:
    """Convert a raw cell value into its published JSON form for a column of ``kind``."""
    if is_blank(value) or is_error_text(value) or is_placeholder(value):
        return None
    if kind in ("number", "percent"):
        if isinstance(value, bool):
            return fmt_plain(value)
        if isinstance(value, (int, float)):
            return _num_out(float(value))
        if isinstance(value, str):
            parsed = parse_numeric_text(value)
            if parsed is not None:
                return _num_out(parsed[0])
        return fmt_plain(value)
    if kind in ("date", "datetime", "time"):
        if isinstance(value, (dt.datetime, dt.date, dt.time)):
            return iso_datetime(value)
        if isinstance(value, str):
            parsed_date = parse_date_text(value)
            if parsed_date is not None:
                return iso_datetime(parsed_date)
        return fmt_plain(value)
    # text (and anything unexpected): stringify deterministically
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return iso_datetime(value)
    if isinstance(value, str):
        return value.strip()
    return fmt_plain(value)


def month_key(value: Any) -> Optional[str]:
    if isinstance(value, str):
        value = parse_date_text(value)
    if isinstance(value, dt.datetime) or isinstance(value, dt.date):
        return f"{value.year:04d}-{value.month:02d}"
    return None


def day_key(value: Any) -> Optional[str]:
    if isinstance(value, str):
        value = parse_date_text(value)
    if isinstance(value, dt.datetime) or isinstance(value, dt.date):
        return f"{value.year:04d}-{value.month:02d}-{value.day:02d}"
    return None


def numeric_of(value: Any) -> Optional[float]:
    if isinstance(value, bool) or is_blank(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        parsed = parse_numeric_text(value)
        return parsed[0] if parsed else None
    return None


def bucket_categories(counter: Dict[str, float], limit: int, other_label: str = "기타") -> Tuple[List[str], List[Any]]:
    """Top ``limit-1`` categories by value plus an aggregated remainder bucket."""
    items = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))
    if len(items) <= limit:
        return [k for k, _ in items], [_num_out(v) for _, v in items]
    head = items[: limit - 1]
    rest = sum(v for _, v in items[limit - 1 :])
    labels = [k for k, _ in head] + [other_label]
    values = [_num_out(v) for _, v in head] + [_num_out(rest)]
    return labels, values

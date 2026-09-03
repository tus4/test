"""Shared helpers: errors, markers, deterministic JSON and number formatting."""

from __future__ import annotations

import hashlib
import json
from typing import Any

# Markers that delimit the generated block inside README.md.
MARK_START = "<!-- DASHBOARD:START -->"
MARK_END = "<!-- DASHBOARD:END -->"

# Names of the artifacts the build writes into the staging directory.
STAGING_DOCS_DIR = "docs"
STAGING_README_BLOCK = "readme_block.md"
STAGING_SUMMARY = "build_summary.json"


class BuildError(Exception):
    """A fatal, user-facing error. The CLI prints the message and exits 2."""


def canonical_json(obj: Any, pretty: bool = False) -> str:
    """Deterministic JSON: sorted keys, UTF-8 text, fixed separators."""
    if pretty:
        return json.dumps(obj, ensure_ascii=False, sort_keys=True, indent=1) + "\n"
    return json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def fingerprint(obj: Any) -> str:
    """Short, content-derived identifier (first 12 hex chars of SHA-256)."""
    return sha256_text(canonical_json(obj))[:12]


def fmt_number(value: Any) -> str:
    """Format a number for display: thousands separators, at most 2 decimals."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return str(value)
        if abs(value - round(value)) < 1e-9 and abs(value) < 1e15:
            return f"{int(round(value)):,}"
        return f"{value:,.2f}"
    return str(value)


def fmt_plain(value: Any) -> str:
    """Stringify a scalar without thousands separators (used for text cells)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9 and abs(value) < 1e15:
            return str(int(round(value)))
        return repr(value)
    return str(value)


def truncate(text: str, limit: int) -> str:
    text = text.replace("\r", " ").replace("\n", " ")
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)] + "…"

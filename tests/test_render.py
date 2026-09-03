from dashboard_builder.common import MARK_END, MARK_START
from dashboard_builder.config import validate_config
from dashboard_builder.render_readme import (
    README_BLOCK_BUDGET,
    md_cell,
    mermaid_label,
    patch_readme,
    render_block,
    render_pie,
)


def _sheet(name="S", rows=None, columns=None, charts=None, notes=None, state="ok"):
    columns = columns or [
        {"name": "지점|A", "excel_column": "A", "type": "text", "non_empty": 2, "stats": {"count": 2, "distinct": 2, "distinct_is_lower_bound": False, "top": [["x", 1], ["y|z", 1]]}},
        {"name": "비율", "excel_column": "B", "type": "percent", "non_empty": 2, "stats": {"count": 2, "sum": 0.75, "mean": 0.375, "min": 0.25, "max": 0.5}},
    ]
    return {
        "name": name,
        "state": state,
        "title": "제목",
        "header_row": 1,
        "rows_total": len(rows or []),
        "rows_scanned": len(rows or []),
        "rows_published": len(rows or []),
        "scan_truncated": False,
        "columns_total": len(columns),
        "columns_published": len(columns),
        "columns_truncated": 0,
        "columns_excluded": [],
        "formula_cells": 0,
        "formula_cells_without_cached_value": 0,
        "columns": columns,
        "rows": rows or [],
        "charts": charts or [],
        "notes": notes or [],
    }


def _data(sheets, sample=False):
    return {
        "schema_version": 1,
        "title": "T",
        "sample": sample,
        "config": {},
        "config_source": None,
        "fingerprint": "abc123",
        "files": [{"path": "f.xlsx", "fingerprint": "def", "sheets": sheets, "notes": []}],
    }


def test_md_cell_escapes_pipes_and_newlines() -> None:
    assert md_cell("a|b\nc") == "a\\|b c"
    assert md_cell(1234.5) == "1,234.50"
    assert md_cell("x" * 100).endswith("…") and len(md_cell("x" * 100)) == 40


def test_mermaid_label_sanitized() -> None:
    assert mermaid_label('a"b#c;d') == "a'b＃c,d"


def test_pie_skips_non_positive_and_formats_ints() -> None:
    chart = {"kind": "pie", "title": "T", "labels": ["a", "b", "c"], "values": [10, 0, 2.5]}
    text = render_pie(chart)
    assert '"a" : 10' in text and '"c" : 2.5' in text and '"b"' not in text
    assert text.startswith("```mermaid\npie showData")


def test_block_contents_and_percent_preview() -> None:
    cfg = validate_config({})
    sheet = _sheet(rows=[["x", 0.25], ["y|z", 0.5]])
    block = render_block(_data([sheet], sample=True), cfg, "https://example.test/")
    assert block.startswith(MARK_START) and block.endswith(MARK_END)
    assert "샘플 데이터" in block
    assert "https://example.test/" in block
    assert "| 25.0% |" in block  # percent column formatted in preview
    assert "y\\|z" in block  # pipes escaped
    assert "\\\\|" not in block  # ...but only once


def test_block_degrades_to_fit_budget() -> None:
    cfg = validate_config({"preview_rows": 1000, "preview_columns": 50})
    rows = [[f"value {i} " + "가" * 30, 0.5] for i in range(1000)]
    sheets = [_sheet(name=f"S{i}", rows=rows) for i in range(5)]
    block = render_block(_data(sheets), cfg, "u")
    assert len(block) <= README_BLOCK_BUDGET
    assert "생략" in block


def test_patch_readme_preserves_outside_text() -> None:
    before = "# 제목\n\n설명 텍스트\n\n<!-- DASHBOARD:START -->\nold\n<!-- DASHBOARD:END -->\n\n## 아래 내용\n"
    block = f"{MARK_START}\nnew\n{MARK_END}"
    after = patch_readme(before, block)
    assert after == "# 제목\n\n설명 텍스트\n\n<!-- DASHBOARD:START -->\nnew\n<!-- DASHBOARD:END -->\n\n## 아래 내용\n"
    assert patch_readme(after, block) == after  # idempotent


def test_patch_readme_appends_when_markers_missing() -> None:
    block = f"{MARK_START}\nnew\n{MARK_END}"
    assert patch_readme("# 제목\n", block) == "# 제목\n\n" + block + "\n"
    assert patch_readme(None, block).startswith("# 데이터 대시보드")
    crlf = "# T\r\n\r\ntext\r\n"
    assert patch_readme(crlf, block).startswith(crlf)


def test_patch_readme_ignores_markers_mentioned_in_prose() -> None:
    block = f"{MARK_START}\nnew\n{MARK_END}"
    prose = f"# 제목\n\n`{MARK_START}` 와 `{MARK_END}` 사이만 바뀝니다.\n"
    once = patch_readme(prose, block)
    assert once == prose + "\n" + block + "\n"
    twice = patch_readme(once.replace("new", "old"), block)
    assert twice == once  # the real block is replaced, the prose mention is untouched

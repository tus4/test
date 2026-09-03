"""Render the README block (Korean, GitHub-flavoured Markdown + Mermaid) and patch README.md."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from .common import MARK_END, MARK_START, fmt_number, truncate

README_BLOCK_BUDGET = 60_000  # characters; well under GitHub's README rendering limits
CELL_LIMIT = 40
MERMAID_LABEL_LIMIT = 30
BAR_WIDTH = 20

DEFAULT_README_HEADER = (
    "# 데이터 대시보드\n\n"
    "이 저장소는 비공개 데이터 저장소의 Excel 파일로부터 자동 생성된 대시보드를 공개합니다.\n\n"
)

TYPE_LABELS = {
    "number": "숫자",
    "percent": "비율(%)",
    "date": "날짜",
    "datetime": "날짜/시간",
    "time": "시간",
    "text": "텍스트",
    "empty": "비어 있음",
}


def md_cell(value: Any, limit: int = CELL_LIMIT) -> str:
    if value is None:
        return ""
    text = fmt_number(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else str(value)
    text = truncate(text, limit).replace("\\", "\\\\").replace("|", "\\|")
    return text


def mermaid_label(text: str) -> str:
    text = str(text).replace('"', "'").replace("#", "＃").replace(";", ",").replace("%", "％").replace("\n", " ")
    return truncate(text, MERMAID_LABEL_LIMIT)


def format_stat_value(value: Any, kind: str) -> str:
    if value is None:
        return ""
    if kind == "percent" and isinstance(value, (int, float)):
        return f"{value * 100:,.1f}%"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return fmt_number(value)
    return str(value)


def preview_cell(value: Any, kind: str) -> str:
    """Preview-table cell: percent columns as '12.5%', everything else like md_cell."""
    if kind == "percent" and isinstance(value, (int, float)) and not isinstance(value, bool):
        return md_cell(format_stat_value(value, kind))
    return md_cell(value)


def column_summary(col: Dict[str, Any]) -> str:
    """Plain (unescaped) one-line summary; the caller escapes it once for Markdown."""
    kind = col["type"]
    st = col.get("stats", {})
    if kind == "percent":
        return (
            f"평균 {format_stat_value(st.get('mean'), kind)} · 최소 {format_stat_value(st.get('min'), kind)} · "
            f"최대 {format_stat_value(st.get('max'), kind)}"
        )
    if kind == "number":
        return (
            f"합계 {format_stat_value(st.get('sum'), kind)} · 평균 {format_stat_value(st.get('mean'), kind)} · "
            f"최소 {format_stat_value(st.get('min'), kind)} · 최대 {format_stat_value(st.get('max'), kind)}"
        )
    if kind in ("date", "datetime"):
        return f"{st.get('min', '')} ~ {st.get('max', '')}"
    if kind == "empty":
        return "값 없음"
    top = st.get("top") or []
    parts = [f"{truncate(str(v), 20)} ({fmt_number(c)})" for v, c in top[:3]]
    distinct = st.get("distinct", 0)
    prefix = f"고유값 {fmt_number(distinct)}{'+' if st.get('distinct_is_lower_bound') else ''}개"
    return prefix + (" · 상위: " + ", ".join(parts) if parts else "")


def _mermaid_number(value: Any) -> str:
    number = round(float(value), 2)
    return str(int(number)) if number.is_integer() else str(number)


def render_pie(chart: Dict[str, Any]) -> str:
    lines = ["```mermaid", "pie showData", f"    title {mermaid_label(chart['title'])}"]
    for label, value in zip(chart["labels"], chart["values"]):
        if isinstance(value, (int, float)) and value > 0:
            lines.append(f'    "{mermaid_label(label)}" : {_mermaid_number(value)}')
    lines.append("```")
    return "\n".join(lines)


def render_xychart(chart: Dict[str, Any]) -> str:
    labels = ", ".join(f'"{mermaid_label(l)}"' for l in chart["labels"])
    values = ", ".join(_mermaid_number(v) for v in chart["values"])
    series = "line" if chart["kind"] == "line" else "bar"
    return "\n".join(
        [
            "```mermaid",
            "xychart-beta",
            f'    title "{mermaid_label(chart["title"])}"',
            f"    x-axis [{labels}]",
            f'    y-axis "{mermaid_label(chart["measure"])}"',
            f"    {series} [{values}]",
            "```",
        ]
    )


def render_bar_table(chart: Dict[str, Any]) -> str:
    """A bar chart that renders everywhere: Markdown table with block characters."""
    values = [float(v) for v in chart["values"]]
    peak = max((abs(v) for v in values), default=0.0)
    lines = [f"**{md_cell(chart['title'], 120)}**", "", f"| {md_cell(chart['category'])} | {md_cell(chart['measure'])} | |", "|---|---:|:---|"]
    for label, value in zip(chart["labels"], values):
        width = int(round(abs(value) / peak * BAR_WIDTH)) if peak > 0 else 0
        lines.append(f"| {md_cell(label)} | {fmt_number(value)} | {'█' * width} |")
    return "\n".join(lines)


def render_sheet(file_path: str, sheet: Dict[str, Any], cfg: Dict[str, Any], level: int) -> List[str]:
    """level 0 = everything, 1 = no preview, 2 = no column table, 3 = KPIs only."""
    out: List[str] = [f"#### 시트: {md_cell(sheet['name'], 80)}", ""]
    if sheet["state"] == "empty":
        out.append("_빈 시트입니다._")
        out.append("")
        return out
    if sheet.get("title"):
        out.append(f"> 제목 행: {md_cell(sheet['title'], 200)}")
        out.append("")
    type_counts: Dict[str, int] = {}
    for col in sheet["columns"]:
        type_counts[col["type"]] = type_counts.get(col["type"], 0) + 1
    type_text = ", ".join(f"{TYPE_LABELS.get(k, k)} {v}" for k, v in sorted(type_counts.items()))
    out += [
        "| 항목 | 값 |",
        "|---|---:|",
        f"| 총 행 수 | {fmt_number(sheet['rows_total'])} |",
        f"| 공개 행 수 | {fmt_number(sheet['rows_published'])} |",
        f"| 열 수 (공개/전체) | {sheet['columns_published']} / {sheet['columns_total']} |",
        f"| 헤더 행 | {sheet['header_row']} |",
        f"| 열 유형 | {md_cell(type_text, 200)} |",
        "",
    ]
    if level <= 1 and sheet["columns"]:
        out += ["| 열 | 유형 | 값 개수 | 요약 |", "|---|---|---:|---|"]
        for col in sheet["columns"]:
            out.append(
                f"| {md_cell(col['name'])} | {TYPE_LABELS.get(col['type'], col['type'])} | "
                f"{fmt_number(col['non_empty'])} | {md_cell(column_summary(col), 160)} |"
            )
        out.append("")
    if level <= 2 and cfg["readme"]["charts"]:
        for chart in sheet["charts"]:
            if chart["kind"] == "pie":
                out += [render_pie(chart), ""]
            elif chart["kind"] == "bar":
                out += [render_xychart(chart) if cfg["readme"]["xychart"] else render_bar_table(chart), ""]
            elif chart["kind"] == "line" and cfg["readme"]["xychart"]:
                out += [render_xychart(chart), ""]
    if level == 0 and cfg["readme"]["preview"] and sheet["rows"]:
        n_rows = min(cfg["preview_rows"], len(sheet["rows"]))
        n_cols = min(cfg["preview_columns"], len(sheet["columns"]))
        if n_rows > 0 and n_cols > 0:
            out.append(f"**미리보기** (처음 {n_rows}행, {n_cols}열)")
            out.append("")
            kinds = [c["type"] for c in sheet["columns"][:n_cols]]
            out.append("| " + " | ".join(md_cell(c["name"]) for c in sheet["columns"][:n_cols]) + " |")
            out.append("|" + "---|" * n_cols)
            for row in sheet["rows"][:n_rows]:
                out.append("| " + " | ".join(preview_cell(v, k) for v, k in zip(row[:n_cols], kinds)) + " |")
            out.append("")
    for note in sheet.get("notes", []):
        out.append(f"- ℹ️ {md_cell(note, 300)}")
    if sheet.get("notes"):
        out.append("")
    return out


def policy_text(cfg: Dict[str, Any]) -> str:
    if cfg["mode"] == "rows":
        text = (
            f"행 데이터 공개 (시트당 최대 {fmt_number(cfg['max_rows'])}행, {cfg['max_columns']}열, "
            f"셀 텍스트 {cfg['max_cell_text']}자)"
        )
    else:
        text = "집계만 공개 (행 데이터 비공개; 열 이름·통계·차트 범주는 공개)"
    if cfg.get("auto_exclude_sensitive", True):
        text += " · 개인정보로 추정되는 열은 자동 제외"
    return text


def render_block_at_level(data: Dict[str, Any], cfg: Dict[str, Any], pages_url: str, level: int) -> str:
    lines: List[str] = [MARK_START, "<!-- 이 블록은 자동 생성됩니다. 마커 바깥의 내용만 직접 수정하세요. -->"]
    lines.append(f"## 📊 {md_cell(data['title'], 120)}")
    lines.append("")
    if data.get("sample"):
        lines.append("> ⚠️ **샘플 데이터**로 생성된 예시입니다. 실제 데이터를 업로드하면 이 블록이 교체됩니다.")
        lines.append("")
    n_sheets = sum(len(f["sheets"]) for f in data["files"])
    lines += [
        f"- 🔗 인터랙티브 대시보드: {pages_url}",
        f"- 📁 파일 {len(data['files'])}개 · 시트 {n_sheets}개 · 데이터 지문 `{data['fingerprint']}`",
        f"- 🔒 공개 범위: {policy_text(cfg)}",
    ]
    if data.get("excluded_files"):
        lines.append(f"- 🚫 설정으로 제외한 파일: {data['excluded_files']}개 (이름은 공개하지 않음)")
    lines.append("")
    if not data["files"]:
        lines.append("_공개할 워크북이 없습니다._")
        lines.append("")
    for f in data["files"]:
        lines.append(f"### 📁 {md_cell(f['path'], 120)}")
        lines.append("")
        for note in f.get("notes", []):
            lines.append(f"- ℹ️ {md_cell(note, 300)}")
        if f.get("notes"):
            lines.append("")
        for sheet in f["sheets"]:
            lines += render_sheet(f["path"], sheet, cfg, level)
    if level > 0:
        lines.append(f"_README 크기 제한 때문에 일부 내용을 생략했습니다. 전체 내용은 대시보드에서 확인하세요: {pages_url}_")
        lines.append("")
    lines.append(MARK_END)
    return "\n".join(lines)


def render_block(data: Dict[str, Any], cfg: Dict[str, Any], pages_url: str) -> str:
    """Render the README block, degrading detail until it fits the size budget."""
    block = ""
    for level in range(4):
        block = render_block_at_level(data, cfg, pages_url, level)
        if len(block) <= README_BLOCK_BUDGET:
            return block
    return block


def _find_marker_line(text: str, marker: str, start: int = 0) -> Tuple[int, int]:
    """(start, end) offsets of the first line from ``start`` that consists solely of ``marker``.
    A marker mentioned inside prose (e.g. in backticks) does not count. (-1, -1) if absent."""
    pos = start
    while True:
        idx = text.find(marker, pos)
        if idx == -1:
            return -1, -1
        line_start = text.rfind("\n", 0, idx) + 1
        line_end = text.find("\n", idx)
        if line_end == -1:
            line_end = len(text)
        if text[line_start:line_end].strip() == marker:
            return line_start, line_end
        pos = idx + len(marker)


def patch_readme(existing: Optional[str], block: str, default_header: str = DEFAULT_README_HEADER) -> str:
    """Replace the marker-delimited block; text outside the markers is kept byte-for-byte.
    Markers count only when they stand alone on their own line. Missing markers: the block
    is appended. Missing file: a default Korean header is used."""
    text = default_header if existing is None else existing
    start, start_end = _find_marker_line(text, MARK_START)
    if start == -1:
        sep = "" if text.endswith("\n\n") or text == "" else ("\n" if text.endswith("\n") else "\n\n")
        return text + sep + block + "\n"
    end, end_end = _find_marker_line(text, MARK_END, start_end)
    if end == -1:
        return text[:start] + block + "\n"
    return text[:start] + block + text[end_end:]

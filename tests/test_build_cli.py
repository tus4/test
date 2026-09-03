import filecmp
import json
import re
from pathlib import Path

from dashboard_builder import build as build_cli
from dashboard_builder.common import MARK_END, MARK_START

PKG = Path(__file__).resolve().parent.parent / "dashboard_builder"


def _tree(path: Path):
    return {p.relative_to(path).as_posix(): p.read_bytes() for p in sorted(path.rglob("*")) if p.is_file()}


def test_no_pandas_anywhere() -> None:
    for py in list(PKG.glob("*.py")) + list((PKG.parent / "tests").glob("*.py")):
        assert not re.search(r"^\s*(import|from)\s+pandas", py.read_text(encoding="utf-8"), re.M), py


def test_end_to_end_build_is_idempotent(data_dir: Path, tmp_path: Path, capsys) -> None:
    out = tmp_path / "staging"
    rc = build_cli.main(["--source", str(data_dir), "--out", str(out), "--pages-url", "https://tus4.github.io/test/"])
    assert rc == 0
    printed = capsys.readouterr().out
    assert "3 file(s), 5 sheet(s)" in printed
    assert (out / "docs" / "index.html").is_file()
    assert (out / "docs" / ".nojekyll").is_file()
    assert (out / "docs" / "data" / "dashboard.json").is_file()
    assert (out / "docs" / "data" / "dashboard.js").is_file()
    block = (out / "readme_block.md").read_text(encoding="utf-8")
    assert block.startswith(MARK_START) and block.rstrip().endswith(MARK_END)
    assert "```mermaid" in block and "https://tus4.github.io/test/" in block
    data = json.loads((out / "docs" / "data" / "dashboard.json").read_text(encoding="utf-8"))
    assert [f["path"] for f in data["files"]] == ["2024/재고.xlsx", "hard.xlsx", "macro.xlsm"]
    assert all(len(s["rows"]) <= 1000 and len(s["columns"]) <= 50 for f in data["files"] for s in f["sheets"])
    js = (out / "docs" / "data" / "dashboard.js").read_text(encoding="utf-8")
    assert js.startswith("window.DASHBOARD_DATA = ") and json.loads(js[len("window.DASHBOARD_DATA = "):].rstrip(";\n")) == data
    assert not [p for p in (out / "docs").rglob("*") if p.suffix.lower() in (".xlsx", ".xlsm")]

    first = _tree(out)
    assert build_cli.main(["--source", str(data_dir), "--out", str(out), "--pages-url", "https://tus4.github.io/test/"]) == 0
    assert _tree(out) == first  # byte-identical on unchanged input


def test_config_changes_output(data_dir: Path, tmp_path: Path) -> None:
    (data_dir / "dashboard.config.json").write_text(json.dumps({"mode": "aggregates", "title": "집계"}), encoding="utf-8")
    out = tmp_path / "staging"
    assert build_cli.main(["--source", str(data_dir), "--out", str(out)]) == 0
    data = json.loads((out / "docs" / "data" / "dashboard.json").read_text(encoding="utf-8"))
    assert data["title"] == "집계" and data["config"]["mode"] == "aggregates" and data["config_source"] == "dashboard.config.json"
    assert all(s["rows"] == [] for f in data["files"] for s in f["sheets"])
    assert "집계만 공개" in (out / "readme_block.md").read_text(encoding="utf-8")


def test_broken_workbook_fails_loudly_and_keeps_previous_staging(data_dir: Path, tmp_path: Path, capsys) -> None:
    out = tmp_path / "staging"
    assert build_cli.main(["--source", str(data_dir), "--out", str(out)]) == 0
    before = _tree(out)
    (data_dir / "broken.xlsx").write_bytes(b"definitely not a workbook")
    rc = build_cli.main(["--source", str(data_dir), "--out", str(out)])
    err = capsys.readouterr().err
    assert rc == 2
    assert "broken.xlsx" in err and "::error::" in err
    assert _tree(out) == before  # previous build untouched
    assert not (tmp_path / "staging.tmp").exists()


def test_bad_config_fails_loudly(data_dir: Path, tmp_path: Path, capsys) -> None:
    (data_dir / "dashboard.config.json").write_text('{"exclude_colums": ["x"]}', encoding="utf-8")
    assert build_cli.main(["--source", str(data_dir), "--out", str(tmp_path / "s")]) == 2
    assert "unknown key" in capsys.readouterr().err


def test_no_workbooks_fails_by_default(tmp_path: Path, capsys) -> None:
    src = tmp_path / "empty"
    src.mkdir()
    assert build_cli.main(["--source", str(src), "--out", str(tmp_path / "s")]) == 2
    assert "no workbook" in capsys.readouterr().err
    (src / "dashboard.config.json").write_text('{"fail_if_no_workbooks": false}', encoding="utf-8")
    assert build_cli.main(["--source", str(src), "--out", str(tmp_path / "s")]) == 0
    assert "공개할 워크북이 없습니다" in (tmp_path / "s" / "readme_block.md").read_text(encoding="utf-8")


def test_sample_flag_labels_output(data_dir: Path, tmp_path: Path) -> None:
    out = tmp_path / "s"
    assert build_cli.main(["--source", str(data_dir), "--out", str(out), "--sample"]) == 0
    assert "샘플 데이터" in (out / "readme_block.md").read_text(encoding="utf-8")
    assert json.loads((out / "docs" / "data" / "dashboard.json").read_text(encoding="utf-8"))["sample"] is True

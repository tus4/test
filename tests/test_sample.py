from pathlib import Path

from dashboard_builder import sample


def _tree(path: Path):
    return {p.relative_to(path).as_posix(): p.read_bytes() for p in sorted(path.rglob("*")) if p.is_file()}


def test_sample_is_reproducible(tmp_path: Path) -> None:
    a = tmp_path / "a"
    b = tmp_path / "b"
    a.mkdir()
    b.mkdir()
    sample.regenerate(a)
    sample.regenerate(b)
    assert _tree(a) == _tree(b)
    assert (a / "docs" / "index.html").is_file() and (a / "docs" / "data" / "dashboard.js").is_file()
    readme = (a / "README.md").read_text(encoding="utf-8")
    assert "샘플 데이터" in readme and "<!-- DASHBOARD:START -->" in readme

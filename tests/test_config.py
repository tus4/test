import json
from pathlib import Path

import pytest

from dashboard_builder.common import BuildError
from dashboard_builder.config import DEFAULTS, load_config, validate_config


def test_missing_file_gives_defaults(tmp_path: Path) -> None:
    cfg = load_config(tmp_path / "dashboard.config.json")
    assert cfg["mode"] == "rows"
    assert cfg["max_rows"] == DEFAULTS["max_rows"]
    assert cfg["_source"] is None


def test_unknown_key_is_rejected() -> None:
    with pytest.raises(BuildError, match="unknown key"):
        validate_config({"exclude_column": ["x"]})


def test_bad_mode_is_rejected() -> None:
    with pytest.raises(BuildError, match="mode"):
        validate_config({"mode": "everything"})


def test_bad_types_are_rejected() -> None:
    with pytest.raises(BuildError):
        validate_config({"max_rows": "1000"})
    with pytest.raises(BuildError):
        validate_config({"exclude_files": "backup/**"})
    with pytest.raises(BuildError):
        validate_config({"readme": {"chart": True}})


def test_readme_flags_merge() -> None:
    cfg = validate_config({"readme": {"xychart": True}})
    assert cfg["readme"] == {"charts": True, "xychart": True, "preview": True}


def test_invalid_json_names_the_problem(tmp_path: Path) -> None:
    path = tmp_path / "dashboard.config.json"
    path.write_text("{ not json", encoding="utf-8")
    with pytest.raises(BuildError, match="not valid JSON"):
        load_config(path)


def test_file_is_loaded(tmp_path: Path) -> None:
    path = tmp_path / "dashboard.config.json"
    path.write_text(json.dumps({"mode": "aggregates", "max_rows": 5, "preview_rows": 50}), encoding="utf-8")
    cfg = load_config(path)
    assert cfg["mode"] == "aggregates"
    assert cfg["max_rows"] == 5
    assert cfg["_source"] == "dashboard.config.json"

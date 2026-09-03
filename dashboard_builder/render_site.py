"""Write the GitHub Pages site (docs/) from the dashboard template and the dataset."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any, Dict

from .common import BuildError, canonical_json

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "template"
TEMPLATE_INDEX = TEMPLATE_DIR / "index.html"


def write_site(data: Dict[str, Any], docs_dir: Path, template_index: Path = TEMPLATE_INDEX) -> None:
    """Create ``docs_dir`` with index.html, .nojekyll and data/dashboard.{json,js}.

    ``dashboard.js`` assigns the same JSON to ``window.DASHBOARD_DATA`` so the page
    works from ``file://`` (no fetch/XHR needed); ``dashboard.json`` is for other consumers.
    """
    if not template_index.is_file():
        raise BuildError(f"dashboard template not found: {template_index}")
    docs_dir.mkdir(parents=True, exist_ok=True)
    data_dir = docs_dir / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_index, docs_dir / "index.html")
    (docs_dir / ".nojekyll").write_bytes(b"")
    payload = canonical_json(data)
    (data_dir / "dashboard.json").write_text(payload + "\n", encoding="utf-8")
    (data_dir / "dashboard.js").write_text("window.DASHBOARD_DATA = " + payload + ";\n", encoding="utf-8")

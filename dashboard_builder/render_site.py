"""Write the GitHub Pages site (docs/) from the dashboard template and the dataset."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any, Dict

from typing import Optional

from .common import BuildError, canonical_json
from .timeseries import public_market, series_payload

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "template"
TEMPLATE_INDEX = TEMPLATE_DIR / "index.html"
TEMPLATE_MARKET = TEMPLATE_DIR / "market.html"


def write_site(
    data: Dict[str, Any],
    docs_dir: Path,
    template_index: Path = TEMPLATE_INDEX,
    market: Optional[Dict[str, Any]] = None,
    template_market: Path = TEMPLATE_MARKET,
) -> None:
    """Create ``docs_dir`` with index.html, .nojekyll and data/dashboard.{json,js}.

    ``dashboard.js`` assigns the same JSON to ``window.DASHBOARD_DATA`` so the page
    works from ``file://`` (no fetch/XHR needed); ``dashboard.json`` is for other consumers.

    With a market dataset the time-series dashboard becomes ``index.html`` (it fetches
    ``data/market.json`` and one ``data/series/<id>.json`` per plotted series) and the
    table dashboard moves to ``tables.html``.
    """
    if not template_index.is_file():
        raise BuildError(f"dashboard template not found: {template_index}")
    docs_dir.mkdir(parents=True, exist_ok=True)
    data_dir = docs_dir / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    has_market = bool(market and market.get("series"))
    if has_market:
        if not template_market.is_file():
            raise BuildError(f"dashboard template not found: {template_market}")
        shutil.copyfile(template_market, docs_dir / "index.html")
        shutil.copyfile(template_index, docs_dir / "tables.html")
    else:
        shutil.copyfile(template_index, docs_dir / "index.html")
    (docs_dir / ".nojekyll").write_bytes(b"")
    payload = canonical_json(data)
    (data_dir / "dashboard.json").write_text(payload + "\n", encoding="utf-8")
    (data_dir / "dashboard.js").write_text("window.DASHBOARD_DATA = " + payload + ";\n", encoding="utf-8")
    if has_market:
        public = public_market(market)
        public["tables_page"] = "tables.html" if any(s["state"] == "ok" for f in data["files"] for s in f["sheets"]) else None
        (data_dir / "market.json").write_text(canonical_json(public) + "\n", encoding="utf-8")
        series_dir = data_dir / "series"
        series_dir.mkdir(parents=True, exist_ok=True)
        for entry in market["series"]:
            (series_dir / f"{entry['id']}.json").write_text(canonical_json(series_payload(market, entry)) + "\n", encoding="utf-8")

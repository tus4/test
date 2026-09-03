"""CLI: build the dashboard from a data checkout into a staging directory.

    python -m dashboard_builder.build --source ../data1 --out "$RUNNER_TEMP/staging" \
        --pages-url https://tus4.github.io/test/

The staging directory receives ``docs/`` (the Pages site), ``readme_block.md``
(the README block) and ``build_summary.json``. Nothing is written to the public
repository here; ``publish.py`` does that only after this step succeeded.
Exit code 2 = build error (message names the file / sheet).
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, Optional

from .common import (
    STAGING_DOCS_DIR,
    STAGING_README_BLOCK,
    STAGING_SUMMARY,
    BuildError,
    canonical_json,
    fingerprint,
)
from .config import CONFIG_FILENAME, load_config, public_config
from .excel import discover_workbooks_with_excluded, extract_workbook
from .render_readme import render_block
from .render_site import write_site

PAGES_URL_PLACEHOLDER = "(GitHub Pages URL 미설정)"


def build_dataset(source: Path, cfg: Dict[str, Any], sample: bool = False) -> Dict[str, Any]:
    workbooks, excluded_files = discover_workbooks_with_excluded(source, cfg)
    if not workbooks and cfg["fail_if_no_workbooks"]:
        raise BuildError(
            f"no workbook (*.xlsx / *.xlsm) found under '{source}' matching include_files={cfg['include_files']} "
            f"exclude_files={cfg['exclude_files']}. The previously published dashboard is kept. "
            f"Set \"fail_if_no_workbooks\": false in {CONFIG_FILENAME} to publish an empty dashboard instead."
        )
    files = [extract_workbook(path, rel, cfg) for path, rel in workbooks]
    data: Dict[str, Any] = {
        "schema_version": 1,
        "title": cfg["title"],
        "sample": bool(sample),
        "config": public_config(cfg),
        "config_source": cfg.get("_source"),
        "excluded_files": excluded_files,  # count only: names of excluded workbooks are never published
        "files": files,
    }
    data["fingerprint"] = fingerprint([[f["path"], f["fingerprint"]] for f in files])
    return data


def build(source: Path, out: Path, config_path: Optional[Path], pages_url: str, sample: bool = False) -> Dict[str, Any]:
    if not source.is_dir():
        raise BuildError(f"source directory not found: {source}")
    if config_path is None:
        config_path = source / CONFIG_FILENAME
    cfg = load_config(config_path)
    data = build_dataset(source, cfg, sample=sample)
    block = render_block(data, cfg, pages_url or PAGES_URL_PLACEHOLDER)

    tmp = out.parent / (out.name + ".tmp")
    if tmp.exists():
        shutil.rmtree(tmp)
    tmp.mkdir(parents=True)
    write_site(data, tmp / STAGING_DOCS_DIR)
    (tmp / STAGING_README_BLOCK).write_text(block + "\n", encoding="utf-8")
    summary = {
        "fingerprint": data["fingerprint"],
        "mode": cfg["mode"],
        "config_source": cfg.get("_source"),
        "excluded_files": data["excluded_files"],
        "files": [
            {
                "path": f["path"],
                "sheets_hidden": f["sheets_hidden"],
                "sheets_excluded": f["sheets_excluded"],
                "sheets": [
                    {
                        "name": s["name"],
                        "state": s["state"],
                        "rows_total": s["rows_total"],
                        "rows_published": s["rows_published"],
                        "columns_published": s["columns_published"],
                        "columns_excluded": [e["name"] for e in s["columns_excluded"]],
                    }
                    for s in f["sheets"]
                ],
            }
            for f in data["files"]
        ],
        "readme_block_chars": len(block),
    }
    (tmp / STAGING_SUMMARY).write_text(canonical_json(summary, pretty=True), encoding="utf-8")
    if out.exists():
        shutil.rmtree(out)
    tmp.rename(out)
    return summary


def print_summary(summary: Dict[str, Any]) -> None:
    n_sheets = sum(len(f["sheets"]) for f in summary["files"])
    rows = sum(s["rows_published"] for f in summary["files"] for s in f["sheets"])
    print(
        f"Built dashboard: {len(summary['files'])} file(s), {n_sheets} sheet(s), {rows:,} row(s) published, "
        f"mode={summary['mode']}, config={summary['config_source'] or 'defaults'}, "
        f"excluded files={summary['excluded_files']}, "
        f"README block {summary['readme_block_chars']:,} chars, fingerprint {summary['fingerprint']}"
    )
    for f in summary["files"]:
        for s in f["sheets"]:
            excluded = f", excluded columns: {', '.join(s['columns_excluded'])}" if s["columns_excluded"] else ""
            print(
                f"  - {f['path']} / {s['name']}: {s['state']}, rows {s['rows_published']:,}/{s['rows_total']:,}, "
                f"columns {s['columns_published']}{excluded}"
            )
        if f["sheets_hidden"] or f["sheets_excluded"]:
            print(f"    (hidden sheets skipped: {f['sheets_hidden']}, sheets excluded by config: {f['sheets_excluded']})")


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="Build the public dashboard from Excel workbooks.")
    parser.add_argument("--source", required=True, help="checkout of the data repository")
    parser.add_argument("--out", required=True, help="staging directory to (re)create")
    parser.add_argument("--config", default=None, help=f"path to {CONFIG_FILENAME} (default: <source>/{CONFIG_FILENAME})")
    parser.add_argument("--pages-url", default="", help="public dashboard URL shown in the README")
    parser.add_argument("--sample", action="store_true", help="label the output as synthetic sample data")
    args = parser.parse_args(argv)
    try:
        summary = build(
            Path(args.source).resolve(),
            Path(args.out).resolve(),
            Path(args.config).resolve() if args.config else None,
            args.pages_url,
            sample=args.sample,
        )
    except BuildError as exc:
        print(f"::error::{exc}", file=sys.stderr)
        print(f"ERROR: {exc}", file=sys.stderr)
        print("오류: 대시보드를 만들지 못했습니다. 이전에 공개된 대시보드는 그대로 유지됩니다.", file=sys.stderr)
        return 2
    print_summary(summary)
    return 0


if __name__ == "__main__":
    sys.exit(main())

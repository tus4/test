"""CLI: copy a finished staging build into the public repo checkout, commit and push.

    python -m dashboard_builder.publish --staging "$RUNNER_TEMP/staging" --repo ../public \
        --message "Update dashboard (data1@abc1234)"

Safety properties:
* refuses to touch the repo unless the staging build is complete;
* checks that the remote branch is reachable (token / permission problems are reported
  before anything is committed);
* ``docs/`` and the README block are replaced only from that complete build
  (a hand-added ``docs/CNAME`` is preserved);
* no commit when nothing changed (byte-level ``git diff --cached --quiet``);
* ``--dry-run`` shows what would change and never commits or pushes;
* push is a normal (non-force) push; on rejection the checkout is reset to the
  remote branch, the build is re-applied and pushed again (bounded retries).
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Callable, List, Optional

from .common import STAGING_DOCS_DIR, STAGING_README_BLOCK, BuildError
from .render_readme import patch_readme

DEFAULT_AUTHOR_NAME = "github-actions[bot]"
DEFAULT_AUTHOR_EMAIL = "41898282+github-actions[bot]@users.noreply.github.com"
DEFAULT_MESSAGE = "Update dashboard from data repository"
PRESERVED_DOCS_FILES = ("CNAME",)  # user-managed files inside docs/ that survive a rebuild
_CREDENTIAL_IN_URL = re.compile(r"://[^/@\s]+@")


def scrub(text: str) -> str:
    """Never echo credentials that might appear in a remote URL."""
    return _CREDENTIAL_IN_URL.sub("://***@", text or "")


def git(repo: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess:
    proc = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True)
    if check and proc.returncode != 0:
        raise BuildError(f"git {' '.join(args)} failed: {scrub(proc.stderr.strip())}")
    return proc


def check_staging(staging: Path) -> None:
    required = [
        staging / STAGING_DOCS_DIR / "index.html",
        staging / STAGING_DOCS_DIR / "data" / "dashboard.json",
        staging / STAGING_DOCS_DIR / "data" / "dashboard.js",
        staging / STAGING_README_BLOCK,
    ]
    missing = [str(p.relative_to(staging)) for p in required if not p.is_file()]
    if missing:
        raise BuildError("staging build is incomplete, refusing to publish. Missing: " + ", ".join(missing))
    for path in (staging / STAGING_DOCS_DIR).rglob("*"):
        if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
            raise BuildError(f"refusing to publish a workbook file: {path}")


def apply_staging(staging: Path, repo: Path) -> None:
    """Replace repo/docs and the README block with the staging build."""
    check_staging(staging)
    docs = repo / "docs"
    preserved = {}
    for name in PRESERVED_DOCS_FILES:
        if (docs / name).is_file():
            preserved[name] = (docs / name).read_bytes()
    if docs.exists():
        shutil.rmtree(docs)
    shutil.copytree(staging / STAGING_DOCS_DIR, docs)
    for name, content in preserved.items():
        (docs / name).write_bytes(content)
    readme = repo / "README.md"
    existing: Optional[str] = None
    if readme.is_file():
        with open(readme, "r", encoding="utf-8", newline="") as fh:
            existing = fh.read()
    block = (staging / STAGING_README_BLOCK).read_text(encoding="utf-8").rstrip("\n")
    updated = patch_readme(existing, block)
    if updated != existing:
        with open(readme, "w", encoding="utf-8", newline="") as fh:
            fh.write(updated)


def stage_changes(repo: Path) -> bool:
    """git add the published paths; True when the index differs from HEAD."""
    git(repo, "add", "-A", "--", "docs", "README.md")
    has_head = git(repo, "rev-parse", "--verify", "HEAD", check=False).returncode == 0
    if not has_head:
        return git(repo, "diff", "--cached", "--quiet", check=False).returncode != 0
    return git(repo, "diff", "--cached", "--quiet", "HEAD", "--", check=False).returncode != 0


def commit(repo: Path, message: str, name: str, email: str) -> None:
    git(repo, "-c", f"user.name={name}", "-c", f"user.email={email}", "commit", "-q", "-m", message)


def current_branch(repo: Path) -> str:
    branch = git(repo, "rev-parse", "--abbrev-ref", "HEAD").stdout.strip()
    if branch == "HEAD" or not branch:
        raise BuildError("the public repository checkout is in detached HEAD state; check out its default branch")
    return branch


def check_remote(repo: Path, branch: str) -> None:
    """Prove that the remote branch is reachable with the checkout's credentials before committing."""
    proc = git(repo, "ls-remote", "--exit-code", "--heads", "origin", branch, check=False)
    if proc.returncode != 0:
        raise BuildError(
            f"cannot reach branch '{branch}' of the public repository (remote 'origin'): {scrub(proc.stderr.strip())}. "
            "토큰이 만료되었거나, 토큰의 저장소 선택/권한(Contents: Read and write)이 잘못되었거나, "
            "저장소 이름이 틀렸습니다 (SETUP.md 2단계)."
        )


def push_with_retry(
    repo: Path,
    branch: str,
    staging: Path,
    message: str,
    name: str,
    email: str,
    retries: int = 3,
    before_push: Optional[Callable[[int], None]] = None,
) -> str:
    """Push; if the remote moved, rebuild on top of it and retry. Returns 'pushed' or 'up-to-date'."""
    last_error = ""
    for attempt in range(1, retries + 1):
        if before_push is not None:
            before_push(attempt)  # test hook: simulate a concurrent remote commit
        proc = git(repo, "push", "origin", f"HEAD:refs/heads/{branch}", check=False)
        if proc.returncode == 0:
            print(f"Pushed to origin/{branch} (attempt {attempt}).")
            return "pushed"
        last_error = scrub(proc.stderr.strip())
        print(f"::warning::push attempt {attempt} rejected: {last_error}")
        git(repo, "fetch", "origin", branch)
        git(repo, "reset", "--hard", f"origin/{branch}")
        apply_staging(staging, repo)
        if not stage_changes(repo):
            print("Remote branch already contains this build; nothing to push.")
            return "up-to-date"
        commit(repo, message, name, email)
    raise BuildError(f"push to origin/{branch} failed after {retries} attempt(s): {last_error}")


def run(
    staging: Path,
    repo: Path,
    message: str = DEFAULT_MESSAGE,
    push: bool = True,
    retries: int = 3,
    name: str = DEFAULT_AUTHOR_NAME,
    email: str = DEFAULT_AUTHOR_EMAIL,
    before_push: Optional[Callable[[int], None]] = None,
    dry_run: bool = False,
) -> str:
    """Returns 'unchanged', 'dry-run', 'committed' (no push requested), 'pushed' or 'up-to-date'."""
    if not (repo / ".git").exists():
        raise BuildError(f"not a git checkout: {repo}")
    branch = current_branch(repo)
    if push and not dry_run:
        check_remote(repo, branch)
    apply_staging(staging, repo)
    if not stage_changes(repo):
        print("No changes to publish: the dashboard is already up to date.")
        return "unchanged"
    if dry_run:
        stat = git(repo, "diff", "--cached", "--stat", "HEAD", "--", check=False).stdout.strip()
        print("DRY RUN: the following files would be committed and pushed:")
        print(stat)
        print("DRY RUN: nothing was committed or pushed.")
        return "dry-run"
    commit(repo, message, name, email)
    print(f"Committed dashboard update on {branch}.")
    if not push:
        return "committed"
    return push_with_retry(repo, branch, staging, message, name, email, retries=retries, before_push=before_push)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Publish a staging build into the public repository.")
    parser.add_argument("--staging", required=True)
    parser.add_argument("--repo", required=True, help="git checkout of the public repository (on its default branch)")
    parser.add_argument("--message", default=DEFAULT_MESSAGE)
    parser.add_argument("--no-push", action="store_true", help="commit locally only")
    parser.add_argument("--dry-run", action="store_true", help="show what would change; do not commit or push")
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument("--author-name", default=DEFAULT_AUTHOR_NAME)
    parser.add_argument("--author-email", default=DEFAULT_AUTHOR_EMAIL)
    args = parser.parse_args(argv)
    try:
        run(
            Path(args.staging).resolve(),
            Path(args.repo).resolve(),
            message=args.message,
            push=not args.no_push,
            retries=args.retries,
            name=args.author_name,
            email=args.author_email,
            dry_run=args.dry_run,
        )
    except BuildError as exc:
        print(f"::error::{exc}", file=sys.stderr)
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())

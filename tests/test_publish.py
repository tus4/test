"""publish.py against real local git repositories (bare remote + shallow clone)."""

import shutil
import subprocess
from pathlib import Path

import pytest

from dashboard_builder import build as build_cli
from dashboard_builder import publish
from dashboard_builder.common import BuildError

GIT_ID = ["-c", "user.name=t", "-c", "user.email=t@example.com"]


def git(repo: Path, *args: str) -> str:
    return subprocess.run(["git", "-C", str(repo), *GIT_ID, *args], check=True, capture_output=True, text=True).stdout.strip()


@pytest.fixture
def remote_and_clone(tmp_path: Path):
    """A bare 'origin' seeded with README.md + tooling files, and a depth-1 clone of it."""
    seed = tmp_path / "seed"
    seed.mkdir()
    subprocess.run(["git", "init", "-q", "-b", "main", str(seed)], check=True)
    (seed / "README.md").write_text("# 공개 저장소\n\n손으로 쓴 설명입니다.\n", encoding="utf-8")
    (seed / "keep.txt").write_text("keep\n")
    git(seed, "add", "-A")
    git(seed, "commit", "-q", "-m", "seed")
    bare = tmp_path / "origin.git"
    subprocess.run(["git", "clone", "-q", "--bare", str(seed), str(bare)], check=True)
    clone = tmp_path / "public"
    subprocess.run(["git", "clone", "-q", "--depth", "1", f"file://{bare}", str(clone)], check=True)
    return bare, clone


def _staging(data_dir: Path, tmp_path: Path, name: str = "staging") -> Path:
    out = tmp_path / name
    assert build_cli.main(["--source", str(data_dir), "--out", str(out), "--pages-url", "u"]) == 0
    return out


def test_publish_commits_pushes_and_is_idempotent(data_dir: Path, tmp_path: Path, remote_and_clone) -> None:
    bare, clone = remote_and_clone
    staging = _staging(data_dir, tmp_path)
    assert publish.run(staging, clone, message="Update dashboard (data1@abc1234)") == "pushed"
    assert git(bare, "log", "-1", "--format=%s", "main") == "Update dashboard (data1@abc1234)"
    files = git(bare, "ls-tree", "-r", "--name-only", "main").splitlines()
    assert "docs/index.html" in files and "docs/data/dashboard.json" in files and "docs/.nojekyll" in files
    assert not [f for f in files if f.endswith((".xlsx", ".xlsm"))]
    readme = (clone / "README.md").read_text(encoding="utf-8")
    assert readme.startswith("# 공개 저장소\n\n손으로 쓴 설명입니다.\n")
    assert "<!-- DASHBOARD:START -->" in readme

    head_before = git(bare, "rev-parse", "main")
    assert publish.run(staging, clone, message="second") == "unchanged"  # no new commit
    assert git(bare, "rev-parse", "main") == head_before
    assert git(clone, "status", "--porcelain") == ""


def test_publish_replaces_stale_docs_and_keeps_edits_outside_markers(data_dir: Path, tmp_path: Path, remote_and_clone) -> None:
    bare, clone = remote_and_clone
    staging = _staging(data_dir, tmp_path)
    assert publish.run(staging, clone) == "pushed"
    # user edits README outside the markers; an old file lingers in docs/
    readme = clone / "README.md"
    readme.write_text(readme.read_text(encoding="utf-8") + "\n## 추가 설명\n", encoding="utf-8")
    (clone / "docs" / "stale.html").write_text("old")
    git(clone, "add", "-A")
    git(clone, "commit", "-q", "-m", "manual edit")
    git(clone, "push", "-q", "origin", "main")
    (data_dir / "hard.xlsx").unlink()  # data changed -> new build
    staging2 = _staging(data_dir, tmp_path, "staging2")
    assert publish.run(staging2, clone) == "pushed"
    files = git(bare, "ls-tree", "-r", "--name-only", "main").splitlines()
    assert "docs/stale.html" not in files
    text = readme.read_text(encoding="utf-8")
    assert text.endswith("\n## 추가 설명\n") and "hard.xlsx" not in text


def test_push_retry_after_concurrent_remote_commit(data_dir: Path, tmp_path: Path, remote_and_clone) -> None:
    bare, clone = remote_and_clone
    staging = _staging(data_dir, tmp_path)

    def foreign_commit(attempt: int) -> None:
        if attempt != 1:
            return
        other = tmp_path / "other"
        subprocess.run(["git", "clone", "-q", f"file://{bare}", str(other)], check=True)
        (other / "README.md").write_text("# 공개 저장소 (수정됨)\n\n동시 편집.\n", encoding="utf-8")
        git(other, "commit", "-q", "-am", "concurrent edit")
        git(other, "push", "-q", "origin", "main")

    assert publish.run(staging, clone, before_push=foreign_commit) == "pushed"
    log = git(bare, "log", "--format=%s", "main").splitlines()
    assert log[0] == publish.DEFAULT_MESSAGE and log[1] == "concurrent edit"
    text = (clone / "README.md").read_text(encoding="utf-8")
    assert text.startswith("# 공개 저장소 (수정됨)") and "<!-- DASHBOARD:START -->" in text


def test_incomplete_staging_is_refused(tmp_path: Path, remote_and_clone) -> None:
    _, clone = remote_and_clone
    staging = tmp_path / "bad"
    (staging / "docs").mkdir(parents=True)
    with pytest.raises(BuildError, match="incomplete"):
        publish.run(staging, clone)
    assert git(clone, "status", "--porcelain") == ""


def test_workbook_in_staging_is_refused(data_dir: Path, tmp_path: Path, remote_and_clone) -> None:
    _, clone = remote_and_clone
    staging = _staging(data_dir, tmp_path)
    shutil.copy(data_dir / "hard.xlsx", staging / "docs" / "leak.xlsx")
    with pytest.raises(BuildError, match="refusing to publish a workbook"):
        publish.run(staging, clone)

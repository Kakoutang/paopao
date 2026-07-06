#!/usr/bin/env python3
"""Validate the public plugin shell file shape."""

from __future__ import annotations

import fnmatch
import subprocess
import sys
from pathlib import Path

sys.dont_write_bytecode = True

from paopao_file_manifest import PUBLIC_SHELL_FILES

ROOT = Path(__file__).resolve().parents[1]

ALLOWED_TRACKED = {
    ".gitignore",
    ".github/workflows/public-release-guard.yml",
} | set(PUBLIC_SHELL_FILES)

BLOCKED_WORKTREE_PATTERNS = [
    "**/__pycache__/**",
    "**/*.pyc",
    "**/*.pyo",
    "**/*.zip",
    "**/*.tar",
    "**/*.tar.gz",
    "**/*.tgz",
    "**/*.pptx",
    "**/*.ppt",
    "**/*.pdf",
    "**/*.key",
    "**/*.pages",
    "**/*.numbers",
    "**/*.sqlite",
    "**/*.sqlite3",
    "**/*.db",
]


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def tracked_files() -> list[str]:
    out = subprocess.check_output(["git", "ls-files"], cwd=ROOT, text=True)
    return [line.strip() for line in out.splitlines() if line.strip()]


def all_worktree_files() -> list[str]:
    files: list[str] = []
    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        if ".git" in path.parts:
            continue
        files.append(rel(path))
    return sorted(files)


def blocked_pattern(path: str) -> str | None:
    for pattern in BLOCKED_WORKTREE_PATTERNS:
        if fnmatch.fnmatch(path, pattern):
            return pattern
    return None


def main() -> int:
    issues: list[str] = []
    for path in tracked_files():
        if not (ROOT / path).exists():
            continue
        if path not in ALLOWED_TRACKED:
            issues.append(f"tracked file is not allowed in public shell: {path}")

    for path in all_worktree_files():
        if path not in ALLOWED_TRACKED:
            issues.append(f"worktree file is not allowed in public shell: {path}")
        pattern = blocked_pattern(path)
        if pattern:
            issues.append(f"blocked file {path} matched {pattern}")

    if issues:
        print("Public release guard failed:", file=sys.stderr)
        for issue in issues:
            print(f"- {issue}", file=sys.stderr)
        return 1
    print("Public release guard passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

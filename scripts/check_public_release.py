#!/usr/bin/env python3
"""Fail if private Paopao assets are present in the public plugin."""

from __future__ import annotations

import fnmatch
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

ALLOWED_TRACKED = {
    ".codex-plugin/plugin.json",
    ".gitignore",
    "CLAUDE.md",
    "README.md",
    "scripts/check_public_release.py",
    "scripts/paopao_auth.py",
    "scripts/paopao_run.py",
    "scripts/renderer.py",
    "skills/paopao-ppt/SKILL.md",
    "prompts/INDEX.md",
    "prompts/PUBLIC_STYLE.md",
    "prompts/SYSTEM_PROMPT.md",
    "prompts/01C_diagram_with_commentary.md",
    "prompts/02B_dual_chart_with_interpretation_cards.md",
    "prompts/04C_comparison_table_with_summary.md",
    "prompts/07A_executive_summary_scr.md",
    "prompts/08B_initiative_rollout_matrix.md",
    "prompts/09A_chevron_with_detail_rows.md",
    "prompts/14D_headline_metrics_with_charts.md",
    "reference/renderer_guide.md",
    ".github/workflows/public-release-guard.yml",
}

FORBIDDEN_PATTERNS = [
    "docs/**",
    "memory/**",
    "output/**",
    "dist/**",
    "qa/**",
    "pptx/**",
    "image2/**",
    "html/**",
    "spec/**",
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
    "**/pptx_qa.py",
    "**/final_prompt_*.md",
    "**/image2_prompt_*.md",
    "**/generation_request_*.json",
]

FORBIDDEN_TEXT = [
    "paopao-internal",
    "Jenny",
    "/Users/jennytang",
    "SparkDeck",
    "SPARK_DATA_DIR",
    "SPARK_PRESERVED_ASSET",
    "SPARK_CHROMIUM_EXECUTABLE",
]


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def matches_forbidden(path: str) -> str | None:
    for pattern in FORBIDDEN_PATTERNS:
        if fnmatch.fnmatch(path, pattern):
            return pattern
    return None


def tracked_files() -> list[str]:
    out = subprocess.check_output(["git", "ls-files"], cwd=ROOT, text=True)
    return [line.strip() for line in out.splitlines() if line.strip()]


def all_worktree_files() -> list[str]:
    files: list[str] = []
    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        if ".git" in path.parts or "__pycache__" in path.parts:
            continue
        files.append(rel(path))
    return sorted(files)


def text_issues(path: str) -> list[str]:
    if path == "scripts/check_public_release.py":
        return []
    if path in {
        "scripts/renderer.py",
        "scripts/paopao_run.py",
        "reference/renderer_guide.md",
        "prompts/SYSTEM_PROMPT.md",
        "skills/paopao-ppt/SKILL.md",
        "CLAUDE.md",
    }:
        text_forbidden = [
            "paopao-internal",
            "/Users/jennytang",
            "SparkDeck",
            "SPARK_DATA_DIR",
            "SPARK_PRESERVED_ASSET",
            "SPARK_CHROMIUM_EXECUTABLE",
        ]
    else:
        text_forbidden = FORBIDDEN_TEXT
    full = ROOT / path
    if full.suffix.lower() not in {".md", ".py", ".json", ".yml", ".yaml", ".txt"}:
        return []
    try:
        text = full.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return [f"{path}: non-text bytes in a text-like file"]
    issues = []
    for needle in text_forbidden:
        if needle in text:
            issues.append(f"{path}: contains forbidden internal marker {needle!r}")
    return issues


def main() -> int:
    issues: list[str] = []
    for path in tracked_files():
        if path not in ALLOWED_TRACKED:
            issues.append(f"tracked file is not allowed in public shell: {path}")
        pattern = matches_forbidden(path)
        if pattern:
            issues.append(f"tracked forbidden file {path} matched {pattern}")
        issues.extend(text_issues(path))

    for path in all_worktree_files():
        pattern = matches_forbidden(path)
        if pattern:
            issues.append(f"worktree forbidden file {path} matched {pattern}")

    if issues:
        print("Public Paopao release guard failed:", file=sys.stderr)
        for issue in issues:
            print(f"- {issue}", file=sys.stderr)
        return 1
    print("Public Paopao release guard passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""Incrementally update the public paopao plugin files.

This avoids asking the AI agent to reason through reinstall steps or redownload
large directories. It updates only the small managed public runtime files.
"""

from __future__ import annotations

import hashlib
import json
import ssl
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RAW_BASE_TEMPLATE = "https://raw.githubusercontent.com/Kakoutang/paopao/{ref}"
MANAGED_FILES = [
    ".codex-plugin/plugin.json",
    "README.md",
    "prompts/INDEX.md",
    "scripts/check_public_release.py",
    "scripts/paopao_auth.py",
    "scripts/paopao_lab.py",
    "scripts/paopao_run.py",
    "scripts/paopao_codex_assets.py",
    "scripts/paopao_delivery_commands.py",
    "scripts/paopao_html_workflow.py",
    "scripts/paopao_pipeline_state.py",
    "scripts/paopao_token_audit.py",
    "scripts/paopao_update.py",
    "scripts/pptx_qa.py",
    "scripts/renderer.py",
    "skills/paopao-ppt/SKILL.md",
]


def ssl_context() -> ssl.SSLContext:
    try:
        import certifi

        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def remote_ref() -> str:
    try:
        out = subprocess.check_output(
            ["git", "ls-remote", "origin", "refs/heads/main"],
            cwd=ROOT,
            text=True,
            stderr=subprocess.DEVNULL,
            timeout=20,
        ).strip()
        if out:
            return out.split()[0]
    except Exception:
        pass
    return "main"


def fetch(path: str, ref: str) -> bytes:
    url = f"{RAW_BASE_TEMPLATE.format(ref=ref)}/{path}"
    req = urllib.request.Request(url, headers={"User-Agent": "paopao-updater"})
    with urllib.request.urlopen(req, timeout=30, context=ssl_context()) as resp:
        return resp.read()


def main() -> int:
    updated: list[str] = []
    unchanged: list[str] = []
    failed: list[str] = []
    ref = remote_ref()

    for rel in MANAGED_FILES:
        target = ROOT / rel
        try:
            remote = fetch(rel, ref)
            remote_hash = sha256_bytes(remote)
            if target.exists() and sha256_file(target) == remote_hash:
                unchanged.append(rel)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(remote)
            updated.append(rel)
        except (OSError, urllib.error.URLError, urllib.error.HTTPError) as exc:
            failed.append(f"{rel}: {exc}")

    result = {
        "ok": not failed,
        "plugin_root": str(ROOT),
        "remote_ref": ref,
        "updated": updated,
        "unchanged_count": len(unchanged),
        "failed": failed,
        "next_step": "Restart the paopao task." if not failed else "Check network access, then run this updater again.",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())

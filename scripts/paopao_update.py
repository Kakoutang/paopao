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

from paopao_file_manifest import (
    AUTHORIZED_RUNTIME_FILES,
    PUBLIC_SHELL_FILES,
    WORKFLOW_DESTINATION_RELS,
)

ROOT = Path(__file__).resolve().parents[1]
RAW_BASE_TEMPLATE = "https://raw.githubusercontent.com/Kakoutang/paopao/{ref}"
MANAGED_FILES = PUBLIC_SHELL_FILES
AUTHORIZED_WORKFLOW_FILES = AUTHORIZED_RUNTIME_FILES
WORKFLOW_DESTINATIONS = {name: ROOT / rel for name, rel in WORKFLOW_DESTINATION_RELS.items()}


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


def fetch_authorized_runtime() -> tuple[list[str], str]:
    try:
        import paopao_auth
    except Exception as exc:
        return [], f"paopao_auth unavailable: {exc}"
    written: list[str] = []
    for name in AUTHORIZED_WORKFLOW_FILES:
        try:
            result = paopao_auth.fetch_workflow_file(name)
            content = str(result.get("content", "")).strip()
            if not content:
                return written, f"Runtime file is empty: {name}"
            target = WORKFLOW_DESTINATIONS[name]
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(content + "\n", encoding="utf-8")
            written.append(str(target.relative_to(ROOT)))
        except Exception as exc:
            return written, str(exc)
    try:
        catalog = paopao_auth.fetch_prompt_catalog()
        for item in catalog.get("prompts", []):
            name = str(item.get("template", "")).strip()
            if not name.endswith(".md") or "/" in name or "\\" in name or ".." in name:
                continue
            result = paopao_auth.fetch_workflow_file(name)
            content = str(result.get("content", "")).strip()
            if not content:
                return written, f"Prompt file is empty: {name}"
            target = ROOT / "prompts" / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(content + "\n", encoding="utf-8")
            written.append(str(target.relative_to(ROOT)))
    except Exception as exc:
        return written, str(exc)
    return written, ""


def main() -> int:
    updated: list[str] = []
    unchanged: list[str] = []
    failed: list[str] = []
    runtime_written: list[str] = []
    runtime_error = ""
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

    if not failed:
        runtime_written, runtime_error = fetch_authorized_runtime()

    result = {
        "ok": not failed and not runtime_error,
        "public_files_ok": not failed,
        "plugin_root": str(ROOT),
        "remote_ref": ref,
        "updated": updated,
        "unchanged_count": len(unchanged),
        "failed": failed,
        "authorized_runtime_updated": runtime_written,
        "authorized_runtime_error": runtime_error,
        "next_step": (
            "Restart the paopao task. Free preview includes 5 pages and 5 prompts; use an activation code only when upgrading."
            if not failed and not runtime_error
            else (
                "Run: python3 scripts/paopao_run.py update. If this keeps failing, contact support."
                if not failed
                else "Check network access, then run this updater again."
            )
        ),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not failed and not runtime_error else 1


if __name__ == "__main__":
    raise SystemExit(main())

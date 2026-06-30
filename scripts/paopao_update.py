#!/usr/bin/env python3
"""Incrementally update the public paopao plugin files.

This avoids asking the AI agent to reason through reinstall steps or redownload
large directories. It updates only the small managed public runtime files.
"""

from __future__ import annotations

import hashlib
import importlib
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
BUNDLE_CHUNK_SIZE = 80
OPTIONAL_WORKFLOW_FILES = {"paopao_delivery_safety.py"}


def current_workflow_manifest() -> tuple[list[str], dict[str, Path]]:
    """Return the latest authorized runtime list after public files refresh."""
    try:
        manifest = importlib.import_module("paopao_file_manifest")
        manifest = importlib.reload(manifest)
        names = list(getattr(manifest, "AUTHORIZED_RUNTIME_FILES"))
        destinations = {
            str(name): ROOT / str(rel)
            for name, rel in getattr(manifest, "WORKFLOW_DESTINATION_RELS").items()
        }
        return names, destinations
    except Exception:
        return list(AUTHORIZED_WORKFLOW_FILES), dict(WORKFLOW_DESTINATIONS)


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


def workflow_file_missing(exc: Exception) -> bool:
    text = str(exc)
    return "HTTP 404" in text and "Workflow file not found" in text


def fetch_authorized_runtime(*, full_library: bool = False) -> tuple[list[str], str, int]:
    try:
        import paopao_auth
    except Exception as exc:
        return [], f"paopao_auth unavailable: {exc}", 0
    names, workflow_destinations = current_workflow_manifest()
    skipped_prompts = 0
    try:
        catalog = paopao_auth.fetch_prompt_catalog()
        for item in catalog.get("prompts", []):
            name = str(item.get("template", "")).strip()
            if not name.endswith(".md") or "/" in name or "\\" in name or ".." in name:
                continue
            if not full_library and not item.get("free"):
                skipped_prompts += 1
                continue
            names.append(name)
    except Exception as exc:
        return [], str(exc), skipped_prompts

    written: list[str] = []
    try:
        for start in range(0, len(names), BUNDLE_CHUNK_SIZE):
            chunk = names[start:start + BUNDLE_CHUNK_SIZE]
            try:
                result = paopao_auth.fetch_workflow_bundle(chunk)
            except Exception as exc:
                if not workflow_file_missing(exc):
                    raise
                result = {"files": []}
                for name in chunk:
                    try:
                        result["files"].append(paopao_auth.fetch_workflow_file(name))
                    except Exception as file_exc:
                        if name in OPTIONAL_WORKFLOW_FILES and workflow_file_missing(file_exc):
                            continue
                        raise
            for item in result.get("files", []):
                name = str(item.get("name", "")).strip()
                content = str(item.get("content", "")).strip()
                if name in workflow_destinations:
                    target = workflow_destinations[name]
                elif name.endswith(".md") and "/" not in name and "\\" not in name and ".." not in name:
                    target = ROOT / "prompts" / name
                else:
                    return written, f"Unexpected workflow file: {name}"
                if not content:
                    return written, f"Workflow file is empty: {name}"
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_text(content + "\n", encoding="utf-8")
                written.append(str(target.relative_to(ROOT)))
    except Exception as exc:
        return written, str(exc), skipped_prompts
    return written, "", skipped_prompts


def summarize(paths: list[str], sample_size: int = 12) -> dict[str, object]:
    return {
        "count": len(paths),
        "sample": paths[:sample_size],
        "truncated": len(paths) > sample_size,
    }


def main(*, full_library: bool = False) -> int:
    updated: list[str] = []
    unchanged: list[str] = []
    failed: list[str] = []
    runtime_written: list[str] = []
    runtime_error = ""
    skipped_prompts = 0
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
        runtime_written, runtime_error, skipped_prompts = fetch_authorized_runtime(full_library=full_library)

    result = {
        "ok": not failed and not runtime_error,
        "public_files_ok": not failed,
        "plugin_root": str(ROOT),
        "remote_ref": ref,
        "updated": summarize(updated),
        "unchanged_count": len(unchanged),
        "failed": failed,
        "authorized_runtime_updated": summarize(runtime_written),
        "library_mode": "full" if full_library else "quick",
        "authorized_prompts_skipped": skipped_prompts,
        "authorized_runtime_error": runtime_error,
        "next_step": (
            "Restart the paopao task. Paopao is ready to create decks."
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

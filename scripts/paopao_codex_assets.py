#!/usr/bin/env python3
"""Codex session asset extraction helpers for Paopao lab workflows."""

from __future__ import annotations

import base64
import hashlib
import json
import struct
import sys
from pathlib import Path


def codex_session_candidates(explicit_session: str = "") -> list[Path]:
    if explicit_session:
        return [Path(explicit_session).expanduser().resolve()]
    sessions_root = Path.home() / ".codex" / "sessions"
    if not sessions_root.exists():
        return []
    files = [p for p in sessions_root.rglob("*.jsonl") if p.is_file()]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files


def _read_matching_payload(session_path: Path, tool_call_id: str) -> list[dict[str, object]]:
    payloads: list[dict[str, object]] = []
    with session_path.open("r", encoding="utf-8", errors="ignore") as handle:
        for line in handle:
            if tool_call_id not in line:
                continue
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            payload = record.get("payload", {})
            if not isinstance(payload, dict):
                continue
            if payload.get("type") not in {"image_generation_call", "image_generation_end"}:
                continue
            if payload.get("id") != tool_call_id and payload.get("call_id") != tool_call_id:
                continue
            payloads.append(payload)
    return payloads


def decode_codex_imagegen_payload(session_path: Path, tool_call_id: str) -> bytes | None:
    for payload in _read_matching_payload(session_path, tool_call_id):
        encoded = payload.get("result")
        if not isinstance(encoded, str) or not encoded.strip():
            continue
        try:
            return base64.b64decode(encoded, validate=True)
        except Exception:
            continue
    return None


def decode_codex_imagegen_prompt(session_path: Path, tool_call_id: str) -> str | None:
    for payload in _read_matching_payload(session_path, tool_call_id):
        prompt = payload.get("prompt")
        if isinstance(prompt, str) and prompt.strip():
            return prompt
        revised_prompt = payload.get("revised_prompt")
        if isinstance(revised_prompt, str) and revised_prompt.strip():
            return revised_prompt
    return None


def find_codex_imagegen_prompt(tool_call_id: str, session_arg: str = "") -> tuple[str | None, Path | None]:
    for session_path in codex_session_candidates(session_arg):
        if not session_path.exists():
            continue
        prompt = decode_codex_imagegen_prompt(session_path, tool_call_id)
        if prompt is not None:
            return prompt, session_path
    return None, None


def png_size(raw: bytes) -> tuple[int, int] | None:
    if len(raw) < 24 or raw[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    width, height = struct.unpack(">II", raw[16:24])
    return int(width), int(height)


def cmd_extract_codex_imagegen_result(args: object) -> int:
    tool_call_id = str(getattr(args, "tool_call_id", "") or "").strip()
    if len(tool_call_id) < 8:
        print(json.dumps({"ok": False, "issue": "tool-call id is required"}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    output = Path(str(getattr(args, "output", ""))).expanduser().resolve()
    if output.exists() and not getattr(args, "force", False):
        print(json.dumps({
            "ok": False,
            "issue": "output already exists; pass --force to overwrite",
            "output": str(output),
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    matched_session: Path | None = None
    raw: bytes | None = None
    for session_path in codex_session_candidates(str(getattr(args, "session", "") or "")):
        if not session_path.exists():
            continue
        raw = decode_codex_imagegen_payload(session_path, tool_call_id)
        if raw:
            matched_session = session_path
            break
    if not raw or not matched_session:
        print(json.dumps({
            "ok": False,
            "issue": "image generation result not found in Codex session logs",
            "tool_call_id": tool_call_id,
            "hint": "Use a Codex image_generation tool-call id such as ig_..., or pass --session to the correct rollout jsonl.",
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    size = png_size(raw)
    if size is None:
        print(json.dumps({
            "ok": False,
            "issue": "decoded payload is not a valid PNG image",
            "tool_call_id": tool_call_id,
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(raw)
    print(json.dumps({
        "ok": True,
        "output": str(output),
        "tool_call_id": tool_call_id,
        "session": str(matched_session),
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "width": size[0],
        "height": size[1],
    }, ensure_ascii=False, indent=2))
    return 0

#!/usr/bin/env python3
"""Thin public bootstrap for Paopao.

The public package intentionally does not ship private prompts, layout catalogs,
or the direct PPTX runtime. Installations fetch authorized files at runtime;
starter access is created automatically on first use.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

from paopao_file_manifest import AUTHORIZED_RUNTIME_FILES, WORKFLOW_DESTINATION_RELS

PLUGIN_ROOT = Path(__file__).resolve().parents[1]
AUTHORIZED_WORKFLOW_FILES = AUTHORIZED_RUNTIME_FILES


def _load_sibling(name: str):
    sys.path.insert(0, str(PLUGIN_ROOT / "scripts"))
    return __import__(name)


def workflow_destinations() -> dict[str, Path]:
    return {name: PLUGIN_ROOT / rel for name, rel in WORKFLOW_DESTINATION_RELS.items()}


def fetch_workflow_file(name: str, destination: Path) -> None:
    paopao_auth = _load_sibling("paopao_auth")
    try:
        result = paopao_auth.fetch_workflow_file(name)
    except paopao_auth.AuthError as exc:
        raise SystemExit(str(exc)) from exc
    content = str(result.get("content", "")).strip()
    if not content:
        raise SystemExit(f"Workflow file is empty: {name}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(content + "\n", encoding="utf-8")


def fetch_prompt_templates() -> list[str]:
    paopao_auth = _load_sibling("paopao_auth")
    try:
        catalog = paopao_auth.fetch_prompt_catalog()
    except paopao_auth.AuthError as exc:
        raise SystemExit(str(exc)) from exc
    written: list[str] = []
    for item in catalog.get("prompts", []):
        name = str(item.get("template", "")).strip()
        if not name.endswith(".md") or "/" in name or "\\" in name or ".." in name:
            continue
        target = PLUGIN_ROOT / "prompts" / name
        fetch_workflow_file(name, target)
        written.append(str(target.relative_to(PLUGIN_ROOT)))
    return written


def cmd_doctor(_: argparse.Namespace) -> int:
    runtime = PLUGIN_ROOT / "scripts" / "deck_frame.py"
    fetched: list[str] = []
    error = ""
    if not runtime.exists():
        try:
            for name in AUTHORIZED_WORKFLOW_FILES:
                target = workflow_destinations()[name]
                fetch_workflow_file(name, target)
                fetched.append(str(target.relative_to(PLUGIN_ROOT)))
            fetched.extend(fetch_prompt_templates())
        except SystemExit as exc:
            error = str(exc)
    checks = {
        "plugin_root": str(PLUGIN_ROOT),
        "public_bootstrap": True,
        "runtime_present": runtime.exists(),
        "access_ready": True,
        "fetched": fetched,
        "next_step": (
            "Paopao is ready. You can start creating the deck."
            if runtime.exists()
            else "Run: python3 scripts/paopao_run.py update. If this keeps failing, contact support."
        ),
    }
    if error:
        checks["error"] = error
    print(json.dumps(checks, ensure_ascii=False, indent=2))
    return 0 if runtime.exists() else 1


def cmd_fetch_workflow(args: argparse.Namespace) -> int:
    destinations = workflow_destinations()
    names = AUTHORIZED_WORKFLOW_FILES if args.all else [args.name]
    written: list[str] = []
    for name in names:
        if name not in destinations:
            raise SystemExit(f"Unknown workflow file: {name}")
        target = destinations[name]
        fetch_workflow_file(name, target)
        written.append(str(target.relative_to(PLUGIN_ROOT)))
    if args.all:
        written.extend(fetch_prompt_templates())
    print(json.dumps({"ok": True, "written": written}, ensure_ascii=False, indent=2))
    return 0


def cmd_update(_: argparse.Namespace) -> int:
    updater = _load_sibling("paopao_update")
    return updater.main()


def cmd_runtime_required(args: argparse.Namespace) -> int:
    cmd_fetch_workflow(argparse.Namespace(all=True, name="paopao_run.py"))
    os.execv(
        sys.executable,
        [sys.executable, str(PLUGIN_ROOT / "scripts" / "paopao_run.py"), *sys.argv[1:]],
    )
    raise SystemExit("Failed to hand off to refreshed Paopao runtime")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="paopao public bootstrap")
    sub = parser.add_subparsers(dest="command", required=True)

    doctor = sub.add_parser("doctor", help="Check local bootstrap state")
    doctor.set_defaults(func=cmd_doctor)

    update = sub.add_parser("update", help="Update public bootstrap files")
    update.set_defaults(func=cmd_update)

    fetch = sub.add_parser("fetch-workflow", help="Fetch authorized Paopao runtime files")
    fetch.add_argument("--all", action="store_true")
    fetch.add_argument("--name", default="paopao_run.py", choices=sorted(workflow_destinations().keys()))
    fetch.set_defaults(func=cmd_fetch_workflow)

    for name in [
        "init",
        "make-deck",
        "next",
        "check",
        "plan-prompts",
        "finalize-delivery",
        "prepare-direct-build-packets",
        "render-pptx-previews",
    ]:
        command = sub.add_parser(name, help=argparse.SUPPRESS)
        command.set_defaults(func=cmd_runtime_required)
    return parser


def main() -> int:
    args, _unknown = build_parser().parse_known_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""Thin public bootstrap for Paopao.

The public package intentionally does not ship private prompts, layout catalogs,
or the direct PPTX runtime. Licensed installations fetch those files at runtime.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path


PLUGIN_ROOT = Path(__file__).resolve().parents[1]


def _load_sibling(name: str):
    sys.path.insert(0, str(PLUGIN_ROOT / "scripts"))
    return __import__(name)


def workflow_destinations() -> dict[str, Path]:
    return {
        "paopao_run.py": PLUGIN_ROOT / "scripts" / "paopao_run.py",
        "SKILL.md": PLUGIN_ROOT / "skills" / "paopao-ppt" / "SKILL.md",
        "SYSTEM_PROMPT.md": PLUGIN_ROOT / "prompts" / "SYSTEM_PROMPT.md",
        "direct_pptx_guide.md": PLUGIN_ROOT / "reference" / "direct_pptx_guide.md",
        "deck_frame.py": PLUGIN_ROOT / "scripts" / "deck_frame.py",
        "renderer_guide.md": PLUGIN_ROOT / "reference" / "renderer_guide.md",
        "renderer.py": PLUGIN_ROOT / "scripts" / "renderer.py",
    }


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


def cmd_doctor(_: argparse.Namespace) -> int:
    checks = {
        "plugin_root": str(PLUGIN_ROOT),
        "public_bootstrap": True,
        "licensed_runtime_present": (PLUGIN_ROOT / "scripts" / "deck_frame.py").exists(),
        "next_step": "Run fetch-workflow --all after activation to install the licensed runtime.",
    }
    print(json.dumps(checks, ensure_ascii=False, indent=2))
    return 0


def cmd_fetch_workflow(args: argparse.Namespace) -> int:
    destinations = workflow_destinations()
    names = list(destinations) if args.all else [args.name]
    written: list[str] = []
    for name in names:
        if name not in destinations:
            raise SystemExit(f"Unknown workflow file: {name}")
        target = destinations[name]
        fetch_workflow_file(name, target)
        written.append(str(target.relative_to(PLUGIN_ROOT)))
    print(json.dumps({"ok": True, "written": written}, ensure_ascii=False, indent=2))
    return 0


def cmd_update(_: argparse.Namespace) -> int:
    updater = _load_sibling("paopao_update")
    return updater.main()


def cmd_runtime_required(args: argparse.Namespace) -> int:
    raise SystemExit(
        "This public package needs the licensed Paopao runtime before generation.\n"
        "Run: python3 scripts/paopao_run.py fetch-workflow --all\n"
        f"Then rerun your command: {' '.join(['paopao_run.py', *sys.argv[1:]])}"
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="paopao public bootstrap")
    sub = parser.add_subparsers(dest="command", required=True)

    doctor = sub.add_parser("doctor", help="Check local bootstrap state")
    doctor.set_defaults(func=cmd_doctor)

    update = sub.add_parser("update", help="Update public bootstrap files")
    update.set_defaults(func=cmd_update)

    fetch = sub.add_parser("fetch-workflow", help="Fetch licensed Paopao runtime files")
    fetch.add_argument("--all", action="store_true")
    fetch.add_argument("--name", default="paopao_run.py", choices=sorted(workflow_destinations().keys()))
    fetch.set_defaults(func=cmd_fetch_workflow)

    for name in [
        "init",
        "make-deck",
        "next",
        "check",
        "render",
        "finalize-delivery",
        "record-commercial-render",
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

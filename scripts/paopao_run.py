#!/usr/bin/env python3
"""Public helper for the paopao free plugin.

This helper intentionally does not contain Paopao's private commercial
pipeline. It provides only local checks and task scaffolding for the public
free edition.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path


PLUGIN_ROOT = Path(__file__).resolve().parents[1]
PUBLIC_PROMPT = PLUGIN_ROOT / "prompts" / "PUBLIC_STYLE.md"
PROMPT_INDEX = PLUGIN_ROOT / "prompts" / "INDEX.md"
RENDERER = PLUGIN_ROOT / "scripts" / "renderer.py"
RENDERER_GUIDE = PLUGIN_ROOT / "reference" / "renderer_guide.md"
FREE_MAX_SLIDES = int(os.getenv("PAOPAO_FREE_MAX_SLIDES", "15"))


def fail(message: str) -> None:
    raise SystemExit(message)


def validate_pages(pages: int) -> None:
    if pages < 1:
        fail("paopao needs at least 1 slide.")
    if pages > FREE_MAX_SLIDES:
        fail(
            f"paopao 免费版最多支持 {FREE_MAX_SLIDES} 页。"
            "如需更多页数或完整模板库，请联系微信 sugarong_ 获取。\n"
            f"paopao free edition supports up to {FREE_MAX_SLIDES} slides. "
            "For more pages or the full template library, contact WeChat: sugarong_"
        )


def safe_slug(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip())
    return value.strip("-") or "paopao-deck"


def cmd_doctor(_: argparse.Namespace) -> int:
    checks = {
        "plugin_root": str(PLUGIN_ROOT),
        "public_prompt_exists": PUBLIC_PROMPT.exists(),
        "prompt_index_exists": PROMPT_INDEX.exists(),
        "renderer_exists": RENDERER.exists(),
        "renderer_guide_exists": RENDERER_GUIDE.exists(),
        "free_max_slides": FREE_MAX_SLIDES,
        "private_runtime_included": False,
    }
    print(json.dumps(checks, ensure_ascii=False, indent=2))
    required = [
        checks["public_prompt_exists"],
        checks["prompt_index_exists"],
        checks["renderer_exists"],
        checks["renderer_guide_exists"],
    ]
    return 0 if all(required) else 1


def cmd_init(args: argparse.Namespace) -> int:
    validate_pages(args.pages)
    output_root = Path(args.output_root).expanduser().resolve()
    task_dir = output_root / safe_slug(args.name)
    task_dir.mkdir(parents=True, exist_ok=True)
    manifest = {
        "schema": "paopao.public_task.v1",
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "name": args.name,
        "pages": args.pages,
        "language": args.language,
        "focus": args.focus,
        "free_max_slides": FREE_MAX_SLIDES,
        "public_prompt": str(PUBLIC_PROMPT),
        "prompt_index": str(PROMPT_INDEX),
    }
    out = task_dir / "paopao_task.json"
    out.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(task_dir))
    return 0


def cmd_check_pages(args: argparse.Namespace) -> int:
    validate_pages(args.pages)
    print(f"OK: {args.pages} slide(s) within paopao free limit {FREE_MAX_SLIDES}.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="paopao public helper")
    sub = parser.add_subparsers(dest="command", required=True)

    doctor = sub.add_parser("doctor", help="Check public plugin files")
    doctor.set_defaults(func=cmd_doctor)

    init = sub.add_parser("init", help="Create a local public task folder")
    init.add_argument("--name", required=True)
    init.add_argument("--pages", required=True, type=int)
    init.add_argument("--language", default="")
    init.add_argument("--focus", default="")
    init.add_argument("--output-root", default="output")
    init.set_defaults(func=cmd_init)

    check_pages = sub.add_parser("check-pages", help="Enforce free slide limit")
    check_pages.add_argument("--pages", required=True, type=int)
    check_pages.set_defaults(func=cmd_check_pages)

    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    sys.exit(main())

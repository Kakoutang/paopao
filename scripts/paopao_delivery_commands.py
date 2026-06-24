#!/usr/bin/env python3
"""Delivery cleanup, publish, and finalize commands for Paopao."""

from __future__ import annotations

import argparse
import html
import json
import os
import shutil
import time
from pathlib import Path

_CTX_NAMES = ['Path', 'json', 'time', 'shutil', 'argparse', 'os', 'PROMPT_ARCHIVE_ENV', 'PROMPT_ARCHIVE_DEV_ENV', 'PROMPT_PRIVATE_DIR', 'internal_prompt_files', 'prompt_private_path', 'delivery_temp_files', 'expected_pages_from_task', 'pipeline_pass_issues', 'check_pipeline_contract', 'write_pipeline_pass', 'check_delivery_files', 'write_final_delivery_pass', 'final_delivery_pass_path', 'check_pptx_file', 'sha256_file', 'user_visible_quality_summary', 'commercial_render_path', 'commercial_source_of_truth', 'HTML_BROWSER_SOURCE_OF_TRUTH', 'image2_reference_path', 'write_html_prompt_attestation']


def _bind(ctx: object) -> None:
    for name in _CTX_NAMES:
        if name in {"Path", "json", "time", "shutil", "argparse", "os"}:
            continue
        globals()[name] = getattr(ctx, name)


def _cmd_cleanup_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    keep_private = bool(args.keep_private_prompts or os.getenv(PROMPT_ARCHIVE_ENV) == "1")
    if keep_private and not (
        os.getenv("PAOPAO_LOCAL_DEV") == "1" and os.getenv(PROMPT_ARCHIVE_DEV_ENV) == "1"
    ):
        raise SystemExit(
            "Refusing to keep private prompt artifacts outside local development. "
            "Unset --keep-private-prompts/PAOPAO_KEEP_PRIVATE_PROMPTS."
        )
    moved_count = 0
    deleted_prompts: list[str] = []
    private_root = task_dir / PROMPT_PRIVATE_DIR
    if keep_private:
        private_root.mkdir(parents=True, exist_ok=True)
    for path in internal_prompt_files(task_dir):
        src_rel = path.relative_to(task_dir)
        if keep_private:
            dest = prompt_private_path(task_dir, path)
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                dest.unlink()
            shutil.move(str(path), str(dest))
            moved_count += 1
        else:
            path.unlink()
            deleted_prompts.append(str(src_rel))
    if not keep_private and private_root.exists():
        shutil.rmtree(private_root)
    removed_temp_count = 0
    for path in delivery_temp_files(task_dir):
        removed_temp_count += 1
        path.unlink()

    manifest = {
        "policy": "internal prompt and Markdown artifacts are never included in user delivery",
        "kept_private_prompts": keep_private,
        "deleted_prompt_markdown_file_count": len(deleted_prompts),
        "moved_prompt_markdown_file_count": moved_count,
        "removed_temporary_file_count": removed_temp_count,
        "note": (
            "User-facing output must not reveal prompt, Markdown, analysis, QA, or debug artifacts."
        ),
    }
    out = task_dir / "qa" / "delivery_cleanup.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


def _strip_delivery_html_internals(text: str) -> str:
    text = text.replace(' name="paopao-prompt-packet-id"', ' name="paopao-render-id"')
    text = text.replace(" name='paopao-prompt-packet-id'", " name='paopao-render-id'")
    return text


def _write_combined_html(task_dir: Path, delivery_dir: Path, expected: int) -> str:
    slides: list[str] = []
    for idx in range(1, expected + 1):
        html_src = task_dir / "html" / f"slide{idx:02d}.html"
        if not html_src.exists():
            raise SystemExit(f"HTML slide missing: {html_src}")
        slides.append(_strip_delivery_html_internals(html_src.read_text(encoding="utf-8")))

    pages = []
    for idx, slide_html in enumerate(slides, 1):
        pages.append(
            "\n".join([
                f'<section class="page" aria-label="Slide {idx}">',
                f'  <iframe title="Slide {idx}" srcdoc="{html.escape(slide_html, quote=True)}"></iframe>',
                "</section>",
            ])
        )
    combined = "\n".join([
        "<!doctype html>",
        '<html lang="en">',
        "<head>",
        '  <meta charset="utf-8">',
        "  <title>Paopao Deck</title>",
        "  <style>",
        "    body { margin: 0; background: #f3f4f6; font-family: Arial, sans-serif; }",
        "    .page { width: 1920px; height: 1080px; margin: 32px auto; background: #fff; box-shadow: 0 8px 28px rgba(0,0,0,0.12); }",
        "    iframe { width: 1920px; height: 1080px; border: 0; display: block; }",
        "    @media (max-width: 1960px) { .page { transform: scale(calc((100vw - 40px) / 1920)); transform-origin: top center; margin-bottom: calc(32px - (1080px * (1 - ((100vw - 40px) / 1920)))); } }",
        "  </style>",
        "</head>",
        "<body>",
        *pages,
        "</body>",
        "</html>",
        "",
    ])
    target = delivery_dir / "deck.html"
    target.write_text(combined, encoding="utf-8")
    return str(target.relative_to(delivery_dir))


def _cmd_publish_delivery_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    pptx = Path(args.pptx).resolve() if args.pptx else None
    if pptx is None:
        pptx_files = [
            p for p in sorted((task_dir / "pptx").glob("*.pptx"))
            if p.is_file() and not p.name.startswith("~$")
        ]
        if len(pptx_files) != 1:
            raise SystemExit(
                "publish-delivery requires exactly one final PPTX in pptx/ or an explicit --pptx"
            )
        pptx = pptx_files[0].resolve()
    if not pptx.exists() or pptx.suffix.lower() != ".pptx":
        raise SystemExit(f"PPTX missing or invalid: {pptx}")

    expected = expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Initialize task with --pages before publishing.")
    pass_issues = pipeline_pass_issues(task_dir, expected, pptx)
    if pass_issues:
        raise SystemExit(
            "publish-delivery blocked because the full pipeline has not passed:\n- "
            + "\n- ".join(pass_issues)
        )

    delivery_dir = Path(args.output_dir).resolve() if args.output_dir else task_dir / "delivery"
    delivery_dir.mkdir(parents=True, exist_ok=True)
    for path in delivery_dir.iterdir():
        if path.is_file():
            path.unlink()
        elif path.is_dir():
            shutil.rmtree(path)

    target = delivery_dir / pptx.name
    shutil.copy2(pptx, target)
    render_path = commercial_render_path(task_dir)
    source_of_truth = commercial_source_of_truth(task_dir)

    copied_images: list[str] = []
    if expected and bool(getattr(args, "include_slide_images", False)):
        images_dir = delivery_dir / "images"
        images_dir.mkdir(parents=True, exist_ok=True)
        for idx in range(1, expected + 1):
            if source_of_truth == HTML_BROWSER_SOURCE_OF_TRUTH:
                image_src = task_dir / "qa" / "html_source" / f"slide-{idx:02d}.png"
            else:
                image_src = image2_reference_path(task_dir, idx)
            if not image_src.exists():
                if source_of_truth == HTML_BROWSER_SOURCE_OF_TRUTH:
                    continue
                raise SystemExit(f"Selected slide preview missing: {image_src}")
            image_dest = images_dir / f"slide{idx:02d}{image_src.suffix.lower()}"
            shutil.copy2(image_src, image_dest)
            copied_images.append(str(image_dest.relative_to(delivery_dir)))

    if expected:
        for idx in range(1, expected + 1):
            html_src = task_dir / "html" / f"slide{idx:02d}.html"
            if render_path == "html" and not html_src.exists():
                raise SystemExit(f"HTML slide missing: {html_src}")

    delivery_html = ""
    if bool(getattr(args, "include_html", False)):
        if render_path != "html":
            raise SystemExit("--include-html is only supported for HTML-rendered decks")
        delivery_html = _write_combined_html(task_dir, delivery_dir, expected)

    manifest = {
        "published_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "source_pptx": str(pptx),
        "delivery_pptx": str(target),
        "delivery_html": delivery_html,
        "pptx_sha256": sha256_file(target),
        "delivery_images": copied_images,
        "user_visible_summary": user_visible_quality_summary(task_dir, expected),
        "policy": (
            "user-facing delivery contains the PPTX by default; optional combined HTML and slide images "
            "are included only when explicitly requested. Prompt, Markdown, analysis, spec, QA, and debug "
            "files are internal only"
        ),
    }
    manifest_path = task_dir / "qa" / "delivery_publish_manifest.json"
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


def _cmd_finalize_delivery_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    expected = expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Initialize task with --pages before finalizing.")
    pptx = Path(args.pptx).resolve() if args.pptx else None
    if pptx is None:
        pptx_files = [
            p for p in sorted((task_dir / "pptx").glob("*.pptx"))
            if p.is_file() and not p.name.startswith("~$")
        ]
        if len(pptx_files) != 1:
            raise SystemExit(
                "finalize-delivery requires exactly one final PPTX in pptx/ or an explicit --pptx"
            )
        pptx = pptx_files[0].resolve()
    if not pptx.exists() or pptx.suffix.lower() != ".pptx":
        raise SystemExit(f"PPTX missing or invalid: {pptx}")

    existing_pass = final_delivery_pass_path(task_dir)
    if existing_pass.exists():
        try:
            cached = json.loads(existing_pass.read_text(encoding="utf-8"))
            if cached.get("pptx_sha256") == sha256_file(pptx):
                cached["idempotent_cache_hit"] = True
                print(json.dumps(cached, indent=2, ensure_ascii=False))
                return 0
        except Exception:
            pass

    pipeline_issues: list[str] = []
    pipeline_counts = check_pipeline_contract(task_dir, expected, pptx, pipeline_issues)
    if pipeline_issues:
        print(json.dumps({
            "task_dir": str(task_dir),
            "stage": "finalize-pipeline",
            "ok": False,
            "issues": pipeline_issues,
            "counts": pipeline_counts,
        }, indent=2, ensure_ascii=False))
        return 1
    pipeline_receipt = write_pipeline_pass(task_dir, expected, pptx, pipeline_counts)

    publish_args = argparse.Namespace(
        task_dir=str(task_dir),
        pptx=str(pptx),
        output_dir="",
        include_slide_images=bool(getattr(args, "include_slide_images", False)),
        include_html=bool(getattr(args, "include_html", False)),
    )
    _cmd_publish_delivery_impl(ctx, publish_args)
    attestation = write_html_prompt_attestation(task_dir, expected)

    delivery_issues: list[str] = []
    delivery_counts = check_delivery_files(
        task_dir,
        expected,
        delivery_issues,
        require_final_pass=False,
    )
    if delivery_issues:
        print(json.dumps({
            "task_dir": str(task_dir),
            "stage": "finalize-delivery",
            "ok": False,
            "issues": delivery_issues,
            "pipeline_pass": str(pipeline_receipt.relative_to(task_dir)),
            "counts": delivery_counts,
        }, indent=2, ensure_ascii=False))
        return 1

    cleanup_args = argparse.Namespace(
        task_dir=str(task_dir),
        keep_private_prompts=bool(args.keep_private_prompts),
    )
    _cmd_cleanup_impl(ctx, cleanup_args)
    final_receipt = write_final_delivery_pass(task_dir, expected, pptx, delivery_counts)

    final_issues: list[str] = []
    final_counts = check_delivery_files(task_dir, expected, final_issues)
    result = {
        "task_dir": str(task_dir),
        "stage": "finalize-delivery",
        "ok": not final_issues,
        "issues": final_issues,
        "pipeline_pass": str(pipeline_receipt.relative_to(task_dir)),
        "final_delivery_pass": str(final_receipt.relative_to(task_dir)),
        "html_prompt_attestation": str(attestation.relative_to(task_dir)) if attestation else None,
        "delivery_dir": str((task_dir / "delivery").resolve()),
        "counts": final_counts,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0 if not final_issues else 1



def cmd_cleanup(ctx: object, args: object) -> int:
    return _cmd_cleanup_impl(ctx, args)


def cmd_publish_delivery(ctx: object, args: object) -> int:
    return _cmd_publish_delivery_impl(ctx, args)


def cmd_finalize_delivery(ctx: object, args: object) -> int:
    return _cmd_finalize_delivery_impl(ctx, args)

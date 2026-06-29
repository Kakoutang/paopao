#!/usr/bin/env python3
"""HTML-source production workflow for Paopao."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import time
from html import escape
from pathlib import Path

_CTX_NAMES = ['Path', 'json', 're', 'time', 'os', 'subprocess', 'sys', 'expected_pages_from_task', 'COMMERCIAL_RENDER_PATHS', 'COMMERCIAL_SOURCE_OF_TRUTH_VALUES', 'render_manifest_source_of_truth', 'IMAGE2_SOURCE_OF_TRUTH', 'HTML_BROWSER_SOURCE_OF_TRUTH', 'check_image2_files', 'post_image_memory_boundary_issues', 'check_visual_inventory_files', 'check_powerpoint_layout_plan_files', 'check_html_files', 'check_html_generation_manifest', 'check_render_manifest', 'commercial_render_contract_path', 'COMMERCIAL_RENDER_CONTRACT_SCHEMA', 'COMMERCIAL_SIMILARITY_MIN', '_relative_to_task_or_abs', 'sha256_file', 'is_html_source_only_task', 'check_html_source_analysis_files', 'ensure_workflow_file', 'SYSTEM_PROMPT', 'HTML_GENERATION_MANIFEST_SCHEMA', 'HTML_GENERATION_SOURCE', 'html_generation_manifest_path', '_read_json_file', 'read_task_manifest', 'read_text', 'verify_fill_origin', 'html_generation_request_path', 'sha256_text', 'HTML_PROMPT_PACKET_META', 'prompt_zone_contract_for_slide', 'HTML_ZONE_ATTR', 'build_deck_navigation_contract', 'html_compact_packet_path', 'html_compact_provenance_path', 'html_compact_renderer_guide_path', 'compact_deck_navigation_contract', 'html_files_from_task', 'RENDERER', 'render_manifest_path', 'PIPELINE_MODE_HTML_SOURCE_ONLY', 'check_html_reference_fidelity', 'reserve_quota', 'finish_quota', 'PLUGIN_ROOT', 'capture_html_preview']


def _bind(ctx: object) -> None:
    for name in _CTX_NAMES:
        if name in {"Path", "json", "re", "time", "os", "subprocess", "sys"}:
            continue
        globals()[name] = getattr(ctx, name)


def _cmd_record_commercial_render_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    expected = expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Initialize task with --pages before recording commercial render.")
    render_path = str(args.render_path).strip()
    if render_path not in COMMERCIAL_RENDER_PATHS:
        raise SystemExit("--render-path must be html or direct_pptx")
    source_of_truth = str(args.source_of_truth or "").strip() or render_manifest_source_of_truth(task_dir)
    if source_of_truth not in COMMERCIAL_SOURCE_OF_TRUTH_VALUES:
        raise SystemExit("--source-of-truth must be image2_reference or html_browser_render")
    pptx = Path(args.pptx).resolve()
    if not pptx.exists() or pptx.suffix.lower() != ".pptx":
        raise SystemExit(f"PPTX missing or invalid: {pptx}")
    preflight_issues: list[str] = []
    if source_of_truth == IMAGE2_SOURCE_OF_TRUTH:
        check_image2_files(task_dir, expected, preflight_issues)
        preflight_issues.extend(post_image_memory_boundary_issues(task_dir, expected))
    if render_path == "direct_pptx":
        check_visual_inventory_files(task_dir, expected, preflight_issues)
        check_powerpoint_layout_plan_files(task_dir, expected, preflight_issues)
    else:
        check_html_files(task_dir, expected, preflight_issues, source_of_truth=source_of_truth)
        if source_of_truth == HTML_BROWSER_SOURCE_OF_TRUTH:
            check_html_generation_manifest(task_dir, expected, preflight_issues)
        check_render_manifest(task_dir, pptx, preflight_issues)
    if preflight_issues:
        raise SystemExit(
            "record-commercial-render blocked because the image-first reconstruction gates have not passed:\n- "
            + "\n- ".join(preflight_issues[:30])
        )
    contract = {
        "schema": COMMERCIAL_RENDER_CONTRACT_SCHEMA,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "expected_pages": expected,
        "render_path": render_path,
        "source_of_truth": source_of_truth,
        "post_image_inputs_only": source_of_truth == IMAGE2_SOURCE_OF_TRUTH,
        "html_source_only": source_of_truth == HTML_BROWSER_SOURCE_OF_TRUTH,
        "commercial_similarity_min": COMMERCIAL_SIMILARITY_MIN,
        "pptx_path": _relative_to_task_or_abs(task_dir, pptx),
        "pptx_sha256": sha256_file(pptx),
        "actual_preview_dir": "qa/pptx_actual",
        "html_is_debug_only": render_path == "direct_pptx",
        "policy": (
            "Commercial delivery uses the declared source_of_truth as the reconstruction reference. "
            "When source_of_truth=html_browser_render, renderer.py must copy the final browser layout into editable PPTX "
            "without reinterpreting prompts, analysis, Image2, or remembered intent. On the default HTML-source-only "
            "path, final delivery uses one-pass structural/render checks and does not require PPTX preview export."
        ),
    }
    out = commercial_render_contract_path(task_dir)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(contract, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "contract": str(out), "render_path": render_path}, ensure_ascii=False, indent=2))
    return 0


def _cmd_generate_html_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    expected = expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Initialize task with --pages before generating HTML.")
    if not is_html_source_only_task(task_dir):
        raise SystemExit("generate-html is only valid for html_source_only tasks.")
    analysis_issues: list[str] = []
    check_html_source_analysis_files(task_dir, expected, analysis_issues)
    if analysis_issues:
        raise SystemExit("generate-html blocked because analysis gates have not passed:\n- " + "\n- ".join(analysis_issues[:30]))
    slides = list(range(1, expected + 1)) if args.all else [int(args.slide)]
    if any(idx < 1 or idx > expected for idx in slides):
        raise SystemExit(f"--slide must be between 1 and {expected}")

    ensure_workflow_file("SYSTEM_PROMPT.md")
    system_prompt_sha = sha256_file(SYSTEM_PROMPT)
    manifest_path = html_generation_manifest_path(task_dir)
    manifest = _read_json_file(manifest_path) or {
        "schema": HTML_GENERATION_MANIFEST_SCHEMA,
        "generation_source": HTML_GENERATION_SOURCE,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "expected_pages": expected,
        "system_prompt_sha256": system_prompt_sha,
        "slides": [],
    }
    existing: dict[int, dict[str, object]] = {}
    raw_slides = manifest.get("slides")
    if isinstance(raw_slides, list):
        for item in raw_slides:
            if isinstance(item, dict) and isinstance(item.get("slide"), int):
                existing[int(item["slide"])] = item

    written: list[dict[str, object]] = []
    task_manifest = read_task_manifest(task_dir)
    contamination_files = sorted((task_dir / "analysis").glob("html_prompt_*.md"))
    if contamination_files:
        names = [f.name for f in contamination_files]
        raise SystemExit(
            f"Contamination detected: {names} found in analysis/. "
            "html_prompt_XX.md files are not allowed — only final_prompt_XX.md from fill-prompt-template. "
            "Delete these files and re-run."
        )
    for idx in slides:
        final_prompt = task_dir / "analysis" / f"final_prompt_{idx:02d}.md"
        html_path = task_dir / "html" / f"slide{idx:02d}.html"
        final_prompt_text = read_text(final_prompt)
        if not final_prompt_text.strip():
            raise SystemExit(f"final_prompt_{idx:02d}.md missing or empty")
        origin_error = verify_fill_origin(final_prompt)
        if origin_error:
            raise SystemExit(f"generate-html blocked for slide {idx}: {origin_error}")
        prompt_packet = html_generation_request_path(task_dir, idx)
        prompt_packet.parent.mkdir(parents=True, exist_ok=True)
        prompt_packet_id = sha256_text("|".join([
            task_dir.name,
            str(idx),
            str(expected),
            system_prompt_sha,
            sha256_file(final_prompt),
            HTML_GENERATION_SOURCE,
        ]))
        required_marker = f'<meta name="{HTML_PROMPT_PACKET_META}" content="{prompt_packet_id}">'
        packet_text = "\n\n".join([
            "# Paopao Locked HTML Generation Packet",
            f"TASK_NAME: {task_dir.name}",
            f"SLIDE: {idx} / {expected}",
            f"LANGUAGE: {str(task_manifest.get('language', '') or '')}",
            f"FOCUS: {str(task_manifest.get('focus', '') or '')}",
            f"OUTPUT_HTML_PATH: {html_path.relative_to(task_dir)}",
            f"GENERATION_SOURCE: {HTML_GENERATION_SOURCE}",
            f"PROMPT_PACKET_ID: {prompt_packet_id}",
            f"REQUIRED_HTML_MARKER: {required_marker}",
            "HARD_RULE: Generate exactly one complete 16:9 production HTML slide from the SYSTEM_PROMPT and FINAL_PROMPT below. Do not use any other prompt, memory, image reference, template prose, or handwritten substitute.",
            "HARD_RULE: The HTML <head> must include the exact REQUIRED_HTML_MARKER above. This marker binds the HTML to this locked Paopao prompt packet.",
            "HARD_RULE: Return/write only complete HTML. Use real text, CSS boxes, tables, native data-chart blocks, and CSS shapes. No SVG, no whole-slide screenshot background.",
            "HARD_RULE: Follow the reference-style area budget: white/near-white surfaces >=80% of slide area; pale-blue fills (#D9EAF7/#EAF1F8) <=12%; deep-blue fills (#305496/#4472C4) <=8%. Normal modules, cards, tables, and chart panels must default to white.",
            "HARD_RULE: Do not default to blue-filled cards or panels. Pale-blue is only for table headers, one selected row, one focus callout, or a template-required grouped object. Deep-blue is only for narrow functional emphasis such as nav, title rule, small badges, selected bars, or a slim takeaway strip.",
            "HARD_RULE: Reserve a clear bottom safe area for the takeaway/source strip. Main content must end above the takeaway region; do not let tables, cards, arrows, labels, or charts overlap the bottom strip.",
            "PROMPT_ZONE_EXECUTION_CONTRACT:\n" + json.dumps(prompt_zone_contract_for_slide(task_dir, idx), ensure_ascii=False, indent=2, sort_keys=True),
            f"HARD_RULE: Every zone listed in PROMPT_ZONE_EXECUTION_CONTRACT.zones must appear in the HTML as an element with {HTML_ZONE_ATTR}=\"<exact zone name>\".",
            "HARD_RULE: Chart/metric zones must use data-chart with categories and series so PowerPoint receives native editable charts with embedded workbook data.",
            "NAVIGATION_CONTRACT:\n" + json.dumps(build_deck_navigation_contract(task_dir, idx), ensure_ascii=False, indent=2, sort_keys=True),
            "SYSTEM_PROMPT:\n" + read_text(SYSTEM_PROMPT),
            "FINAL_PROMPT:\n" + final_prompt_text,
        ])
        if not prompt_packet.exists() or read_text(prompt_packet) != packet_text:
            prompt_packet.write_text(packet_text, encoding="utf-8")
        compact_packet = _write_html_compact_packet_impl(
            ctx,
            task_dir,
            idx,
            expected,
            prompt_packet_id=prompt_packet_id,
            required_marker=required_marker,
            prompt_packet=prompt_packet,
            final_prompt=final_prompt,
            html_path=html_path,
            task_manifest=task_manifest,
        )
        html_ready = (
            html_path.exists()
            and read_text(html_path).strip()
            and html_path.stat().st_mtime >= prompt_packet.stat().st_mtime
            and required_marker in read_text(html_path)
        )
        agent_prompt_path = task_dir / "qa" / "html_generation_requests" / f"agent_prompt_{idx:02d}.md"
        agent_prompt_path.parent.mkdir(parents=True, exist_ok=True)
        zone_contract = prompt_zone_contract_for_slide(task_dir, idx)
        nav_contract = compact_deck_navigation_contract(task_dir, idx, expected)
        nav_labels = [e["label"] for e in nav_contract.get("labels", []) if e.get("label")]
        agent_prompt_text = "\n\n".join([
            "# Paopao HTML Slide Generation — Final Prompt First",
            "",
            "Generate the slide from the design brief first. Treat it as the source of truth for layout, density, hierarchy, and content. The constraints after the brief are guardrails only; do not let them simplify, shrink, or template-fill the design.",
            "",
            "## Design Brief",
            final_prompt_text,
            "---",
            "## Guardrails After Design",
            f"Write exactly one complete 1920x1080 HTML slide to `{html_path.relative_to(task_dir)}`.",
            f"Include this exact marker in the HTML `<head>`: `{required_marker}`.",
            "Do not read any other files. Do not use memory, old drafts, source PDFs, analysis files, or compact packets.",
            "Keep all visible content editable HTML/CSS: real text, divs, tables, CSS boxes, and native data-chart blocks only. No SVG, canvas, CSS background-image, or whole-slide screenshot background.",
            "Use the final prompt's information density. Do not omit bullets, metrics, stage chips, rows, evidence, or commentary to make layout easier.",
            "Use the reference-style consulting visual language within the locked system palette: dominant white surfaces, black title/body, thin grey/blue rules, compact tables, and only small/narrow blue emphasis. Do not create large pale-blue panels, full-width blue washes, or equal-width blue nav buttons.",
            "Area budget: white/near-white >=80% of slide; pale-blue fills <=12%; deep-blue fills <=8%. Normal sections/cards/panels/tables/chart containers are white by default.",
            "Build a consulting report exhibit, not a UI dashboard: prefer tables, matrices, scorecards, charts, process lanes, value chains, and annotated evidence blocks over equal rounded card grids.",
            "Keep deck chrome stable: top nav, title rule, page number, source, and takeaway style must be consistent; only the active nav state changes.",
            "Reserve a bottom safe area for takeaway/source. No card, table, chart, label, connector, or flow element may overlap or be clipped by the bottom strip.",
            "Use CSS grid/flex for main layout. Avoid absolute positioning for content; absolute is allowed only for fixed chrome such as nav or slide background accents.",
            "Every required zone below must appear once as `data-paopao-zone=\"EXACT_NAME\"`:",
            "ZONES: " + json.dumps(zone_contract.get("zones", []), ensure_ascii=False),
            "The nav bar element (`data-paopao-zone=\"nav\"`) must contain visible navigation labels:",
            "NAV_LABELS: " + json.dumps(nav_labels, ensure_ascii=False),
            f"ACTIVE_SLIDE: {idx} / {expected}",
        ])
        agent_prompt_path.write_text(agent_prompt_text, encoding="utf-8")

        written.append({
            "slide": idx,
            "status": "html_ready_for_register" if html_ready else "prompt_packet_ready",
            "prompt_packet": str(prompt_packet.relative_to(task_dir)),
            "final_prompt": str(final_prompt.relative_to(task_dir)),
            "compact_packet": str(compact_packet.relative_to(task_dir)),
            "agent_prompt": str(agent_prompt_path.relative_to(task_dir)),
            "html_input_mode": "full_final_prompt",
            "prompt_packet_id": prompt_packet_id,
            "required_html_marker": required_marker,
            "html_path": str(html_path.relative_to(task_dir)),
            "next_action": f"Spawn a subagent that reads ONLY {agent_prompt_path.relative_to(task_dir)} and writes {html_path.relative_to(task_dir)}. Then run register-html.",
        })

    manifest.update({
        "schema": HTML_GENERATION_MANIFEST_SCHEMA,
        "generation_source": HTML_GENERATION_SOURCE,
        "expected_pages": expected,
        "system_prompt_sha256": system_prompt_sha,
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "slides": [existing.get(i, {"slide": i, "status": "missing"}) for i in range(1, expected + 1)],
    })
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    pending = list(written)
    print(json.dumps({
        "ok": False,
        "task_dir": str(task_dir),
        "registered_slides": [],
        "pending_slides": pending,
        "manifest": str(manifest_path),
    }, ensure_ascii=False, indent=2))
    return 0


def _html_manifest_existing(manifest: dict[str, object]) -> dict[int, dict[str, object]]:
    existing: dict[int, dict[str, object]] = {}
    raw_slides = manifest.get("slides")
    if isinstance(raw_slides, list):
        for item in raw_slides:
            if isinstance(item, dict) and isinstance(item.get("slide"), int):
                existing[int(item["slide"])] = item
    return existing


_ALLOWED_HTML_CHART_TYPES = {"column", "bar", "line", "doughnut", "scatter"}
_ALLOWED_PAOPAO_COLORS = [
    "#305496",
    "#4472C4",
    "#5B9BD5",
    "#D9EAF7",
    "#EAF1F8",
    "#B4C7E7",
    "#FFFFFF",
    "#1C1917",
    "#666666",
]


def _hex_to_rgb(color: str) -> tuple[int, int, int]:
    value = color.strip().lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def _nearest_paopao_color(color: str) -> str:
    rgb = _hex_to_rgb(color)
    return min(
        _ALLOWED_PAOPAO_COLORS,
        key=lambda allowed: sum((a - b) ** 2 for a, b in zip(rgb, _hex_to_rgb(allowed))),
    )


def _normalize_chart_type(value: str) -> str:
    raw = value.strip().lower().replace("_", "-")
    if raw in _ALLOWED_HTML_CHART_TYPES:
        return raw
    if "line" in raw:
        return "line"
    if "scatter" in raw or "bubble" in raw:
        return "scatter"
    if "doughnut" in raw or "donut" in raw or "pie" in raw:
        return "doughnut"
    if "column" in raw:
        return "column"
    return "bar"


def _reference_nav_labels(task_dir: Path, expected: int) -> list[str]:
    story_path = task_dir / "analysis" / "slide_story.json"
    labels: list[str] = []
    try:
        story = json.loads(read_text(story_path)) if story_path.exists() else {}
    except Exception:
        story = {}
    raw_slides = story.get("slides") if isinstance(story, dict) else None
    if isinstance(raw_slides, list):
        by_slide = {
            int(item.get("slide")): str(item.get("section_name") or "").strip()
            for item in raw_slides
            if isinstance(item, dict) and isinstance(item.get("slide"), int)
        }
        labels = [by_slide.get(pos, "") for pos in range(1, expected + 1)]
    return [label or f"Section {pos}" for pos, label in enumerate(labels or [], 1)] or [
        f"Section {pos}" for pos in range(1, expected + 1)
    ]


def _reference_nav_html(task_dir: Path, idx: int, expected: int) -> str:
    labels = _reference_nav_labels(task_dir, expected)
    items = []
    for pos, label in enumerate(labels, 1):
        active = ' class="active"' if pos == idx else ""
        items.append(
            f'<span{active}><b>{pos:02d}</b><em>{escape(label)}</em></span>'
        )
    return (
        '<div class="paopao-reference-nav-slot" data-paopao-zone="nav">'
        '<nav class="paopao-reference-nav">' + "".join(items) + "</nav>"
        "</div>"
    )


def _reference_nav_css() -> str:
    return """

/* PAOPAO_REFERENCE_NAV_START */
/* Based on the supplied consulting decks. The page header is a system component. */
.slide {
  position: relative !important;
}
.paopao-reference-nav-slot {
  height: 38px !important;
  min-height: 38px !important;
  width: 100% !important;
  display: block !important;
  padding: 0 !important;
  margin: 0 0 10px 0 !important;
  background: transparent !important;
  border: 0 !important;
}
.paopao-reference-nav {
  position: relative !important;
  height: 38px !important;
  min-height: 38px !important;
  width: 100% !important;
  display: flex !important;
  align-items: center !important;
  gap: 18px !important;
  padding: 0 18px !important;
  margin: 0 !important;
  background: #305496 !important;
  color: #FFFFFF !important;
  border: 0 !important;
  font-family: Arial, Helvetica, sans-serif !important;
  font-size: 16px !important;
  line-height: 1.05 !important;
  font-weight: 700 !important;
  letter-spacing: 0 !important;
}
.paopao-reference-nav span {
  height: 38px !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 7px !important;
  padding: 0 !important;
  margin: 0 !important;
  background: transparent !important;
  color: rgba(255,255,255,.68) !important;
  border: 0 !important;
  border-radius: 0 !important;
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  max-width: none !important;
}
.paopao-reference-nav span:not(:last-child)::after {
  content: "·" !important;
  margin-left: 18px !important;
  color: rgba(255,255,255,.45) !important;
}
.paopao-reference-nav b {
  font-style: normal !important;
  font-size: 16px !important;
  font-weight: 800 !important;
  flex: 0 0 auto !important;
}
.paopao-reference-nav em {
  font-style: normal !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  white-space: nowrap !important;
  min-width: 0 !important;
}
.paopao-reference-nav span.active {
  background: transparent !important;
  color: #FFFFFF !important;
  font-weight: 900 !important;
  border-bottom: 3px solid #FFFFFF !important;
}
.paopao-reference-nav span.active b,
.paopao-reference-nav span.active em {
  color: #FFFFFF !important;
}
.paopao-reference-nav span:not(.active) b {
  color: rgba(255,255,255,.82) !important;
}
/* PAOPAO_REFERENCE_NAV_END */
"""


def _replace_first_nav_zone_html(text: str, nav_html: str) -> tuple[str, bool]:
    open_match = re.search(
        r"<(?P<tag>nav|div|section|header)\b(?P<attrs>[^>]*)\bdata-paopao-zone\s*=\s*['\"]nav['\"][^>]*>",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not open_match:
        return text, False
    tag = open_match.group("tag")
    pattern = re.compile(rf"</?{re.escape(tag)}\b[^>]*>", flags=re.IGNORECASE | re.DOTALL)
    depth = 0
    for match in pattern.finditer(text, open_match.start()):
        token = match.group(0)
        if token.startswith("</"):
            depth -= 1
            if depth == 0:
                return text[:open_match.start()] + nav_html + text[match.end():], True
        else:
            depth += 1
    return text, False


def _inject_reference_nav(task_dir: Path, text: str, idx: int, expected: int) -> str:
    nav_html = _reference_nav_html(task_dir, idx, expected)
    text = re.sub(
        r"\s*/\*\s*PAOPAO_REFERENCE_NAV_START\s*\*/.*?/\*\s*PAOPAO_REFERENCE_NAV_END\s*\*/",
        "",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    text = re.sub(
        r"\s*/\*\s*Paopao reference chrome:.*?\.paopao-reference-nav span\.active\s*\{.*?\}\s*",
        "",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    text = re.sub(
        r"\s*/\*\s*Based on the supplied consulting decks\..*?/\*\s*PAOPAO_REFERENCE_NAV_END\s*\*/",
        "",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    text, replaced_nav = _replace_first_nav_zone_html(text, nav_html)
    if not replaced_nav:
        text = re.sub(r"(<body\b[^>]*>)", r"\1\n" + nav_html, text, count=1, flags=re.IGNORECASE)

    style_vars = f"--paopao-slide-count: {max(1, expected)};"
    text = re.sub(
        r"(<(?:main|div|section)\b[^>]*\bclass\s*=\s*['\"][^'\"]*\bslide\b[^'\"]*['\"][^>]*style\s*=\s*['\"])([^'\"]*)(['\"])",
        lambda m: m.group(1) + m.group(2).rstrip("; ") + "; " + style_vars + m.group(3),
        text,
        count=1,
        flags=re.IGNORECASE,
    )
    if "--paopao-slide-count" not in text:
        text = re.sub(
            r"(<(?:main|div|section)\b[^>]*\bclass\s*=\s*['\"][^'\"]*\bslide\b[^'\"]*['\"])",
            r'\1 style="' + escape(style_vars, quote=True) + '"',
            text,
            count=1,
            flags=re.IGNORECASE,
        )
    css = _reference_nav_css()
    if ".paopao-reference-nav" not in text:
        if re.search(r"</style>", text, flags=re.IGNORECASE):
            text = re.sub(r"</style>", css + "\n</style>", text, count=1, flags=re.IGNORECASE)
        else:
            text = re.sub(r"</head>", "<style>" + css + "</style>\n</head>", text, count=1, flags=re.IGNORECASE)
    return text


def _normalize_html_for_register(task_dir: Path, idx: int, html_path: Path, required_marker: str) -> bool:
    text = read_text(html_path)
    original = text
    expected = expected_pages_from_task(task_dir) or idx

    text = re.sub(
        rf"\s*<meta\s+name\s*=\s*['\"]{re.escape(HTML_PROMPT_PACKET_META)}['\"]\s+content\s*=\s*['\"][^'\"]*['\"]\s*/?>",
        "",
        text,
        flags=re.IGNORECASE,
    )
    if required_marker not in text:
        if re.search(r"<head\b[^>]*>", text, flags=re.IGNORECASE):
            text = re.sub(r"(<head\b[^>]*>)", r"\1\n  " + required_marker, text, count=1, flags=re.IGNORECASE)
        else:
            text = re.sub(r"(<html\b[^>]*>)", r"\1\n<head>\n  " + required_marker + "\n</head>", text, count=1, flags=re.IGNORECASE)

    def color_repl(match: re.Match[str]) -> str:
        raw = match.group(0)
        normalized = raw.upper()
        if len(normalized) == 4:
            normalized = "#" + "".join(ch * 2 for ch in normalized[1:])
        if normalized in _ALLOWED_PAOPAO_COLORS:
            return normalized
        return _nearest_paopao_color(normalized)

    text = re.sub(r"#[0-9A-Fa-f]{3,6}\b", color_repl, text)

    language = str(read_task_manifest(task_dir).get("language", "") or "").lower()
    if any(token in language for token in ["中文", "汉语", "chinese", "mandarin", "zh", "cn"]):
        replacements = {
            "Key takeaway:": "关键结论：",
            "Key Takeaway:": "关键结论：",
            "key takeaway:": "关键结论：",
            "Source:": "来源：",
            "source:": "来源：",
            "Situation": "现状",
            "Complication": "矛盾",
            "Resolution": "行动",
            "Executive summary": "执行摘要",
            "Agenda": "目录",
        }
        for old, new in replacements.items():
            text = text.replace(old, new)

    def chart_repl(match: re.Match[str]) -> str:
        quote = match.group("quote")
        normalized = _normalize_chart_type(match.group("value"))
        return f"data-chart={quote}{normalized}{quote}"

    text = re.sub(
        r"data-chart\s*=\s*(?P<quote>['\"])(?P<value>[^'\"]+)(?P=quote)",
        chart_repl,
        text,
        flags=re.IGNORECASE,
    )
    text = _inject_reference_nav(task_dir, text, idx, expected)

    if text != original:
        html_path.write_text(text, encoding="utf-8")
        return True
    return False


def _cmd_register_html_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    expected = expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Initialize task with --pages before registering HTML.")
    if not is_html_source_only_task(task_dir):
        raise SystemExit("register-html is only valid for html_source_only tasks.")
    slides = list(range(1, expected + 1)) if args.all else [int(args.slide)]
    if any(idx < 1 or idx > expected for idx in slides):
        raise SystemExit(f"--slide must be between 1 and {expected}")

    ensure_workflow_file("SYSTEM_PROMPT.md")
    system_prompt_sha = sha256_file(SYSTEM_PROMPT)
    manifest_path = html_generation_manifest_path(task_dir)
    manifest = _read_json_file(manifest_path) or {
        "schema": HTML_GENERATION_MANIFEST_SCHEMA,
        "generation_source": HTML_GENERATION_SOURCE,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "expected_pages": expected,
        "system_prompt_sha256": system_prompt_sha,
        "slides": [],
    }
    existing = _html_manifest_existing(manifest)
    registered: list[dict[str, object]] = []
    for idx in slides:
        final_prompt = task_dir / "analysis" / f"final_prompt_{idx:02d}.md"
        html_path = task_dir / "html" / f"slide{idx:02d}.html"
        prompt_packet = html_generation_request_path(task_dir, idx)
        compact_packet = html_compact_packet_path(task_dir, idx)
        compact_provenance = html_compact_provenance_path(task_dir, idx)
        if not final_prompt.exists() or not read_text(final_prompt).strip():
            raise SystemExit(f"final_prompt_{idx:02d}.md missing or empty")
        if not prompt_packet.exists():
            raise SystemExit(f"html_prompt_packet_{idx:02d}.md missing; run generate-html first")
        if not compact_packet.exists():
            raise SystemExit(f"html_compact_packet_{idx:02d}.md missing; run generate-html or compact-html-packet first")
        if not compact_provenance.exists():
            raise SystemExit(f"html_compact_provenance_{idx:02d}.json missing; run generate-html first")
        if not html_path.exists() or not read_text(html_path).strip():
            raise SystemExit(f"html/slide{idx:02d}.html missing or empty")
        prompt_packet_id = sha256_text("|".join([
            task_dir.name,
            str(idx),
            str(expected),
            system_prompt_sha,
            sha256_file(final_prompt),
            HTML_GENERATION_SOURCE,
        ]))
        required_marker = f'<meta name="{HTML_PROMPT_PACKET_META}" content="{prompt_packet_id}">'
        _normalize_html_for_register(task_dir, idx, html_path, required_marker)
        html_text = read_text(html_path)
        if required_marker not in html_text:
            raise SystemExit(f"html/slide{idx:02d}.html missing required prompt packet marker: {required_marker}")
        issues: list[str] = []
        check_html_files(
            task_dir,
            expected,
            issues,
            source_of_truth=HTML_BROWSER_SOURCE_OF_TRUTH,
        )
        slide_issue_prefixes = (f"slide{idx:02d}.html", f"html/slide{idx:02d}.html", f"slide {idx}")
        slide_issues = [
            issue for issue in issues
            if any(prefix.lower() in issue.lower() for prefix in slide_issue_prefixes)
        ]
        if slide_issues:
            raise SystemExit("register-html blocked by HTML issues:\n- " + "\n- ".join(slide_issues[:20]))
        entry = {
            "slide": idx,
            "html_path": str(html_path.relative_to(task_dir)),
            "html_sha256": sha256_file(html_path),
            "final_prompt_path": str(final_prompt.relative_to(task_dir)),
            "final_prompt_sha256": sha256_file(final_prompt),
            "system_prompt_sha256": system_prompt_sha,
            "generated_with_system_prompt": True,
            "generation_source": HTML_GENERATION_SOURCE,
            "prompt_packet_id": prompt_packet_id,
            "prompt_packet_path": str(prompt_packet.relative_to(task_dir)),
            "prompt_packet_sha256": sha256_file(prompt_packet),
            "compact_packet_path": str(compact_packet.relative_to(task_dir)),
            "compact_packet_sha256": sha256_file(compact_packet),
            "compact_provenance_path": str(compact_provenance.relative_to(task_dir)),
            "compact_provenance_sha256": sha256_file(compact_provenance),
            "generator": "codex_host_model_full_final_prompt",
            "html_input_mode": "full_final_prompt",
            "registered_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        }
        existing[idx] = entry
        registered.append(entry)
    manifest.update({
        "schema": HTML_GENERATION_MANIFEST_SCHEMA,
        "generation_source": HTML_GENERATION_SOURCE,
        "expected_pages": expected,
        "system_prompt_sha256": system_prompt_sha,
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "slides": [existing.get(i, {"slide": i, "status": "missing"}) for i in range(1, expected + 1)],
    })
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "ok": True,
        "task_dir": str(task_dir),
        "registered_slides": [item["slide"] for item in registered],
        "manifest": str(manifest_path),
    }, ensure_ascii=False, indent=2))
    return 0


def compact_renderer_guide_text(system_prompt_text: str = "") -> str:
    """Full visual design rules + renderer technical contract for subagents.

    Subagents have their own context window — no need to compress.
    When system_prompt_text is provided, includes the complete SYSTEM_PROMPT
    visual language so the subagent produces the same quality as the full
    prompt packet path.
    """
    parts = [
        "# Paopao HTML Output Contract + Visual Design System",
        "",
        "Use this with one slide's full final_prompt_XX.md. The final_prompt is the design brief; this file provides the complete visual design system and editable-PPTX output contract.",
    ]
    if system_prompt_text.strip():
        parts.extend([
            "",
            "---",
            "",
            "## SYSTEM PROMPT (complete visual design rules)",
            "",
            system_prompt_text,
        ])
    parts.extend([
        "",
        "---",
        "",
        "## Hard Technical Requirements (HTML → PPTX renderer)",
        "",
        "1. Canvas is 1920x1080; root slide must be one overflow:hidden flex column.",
        "2. Structure order: visible nav, title, content, takeaway, source.",
        "3. Every required prompt zone must appear with data-paopao-zone=\"<exact zone name>\".",
        "4. Use real text, divs, tables, and simple CSS shapes; do not bake editable content into images.",
        "5. No whole-slide image or screenshot background. No inline SVG, canvas, CSS background-image, or unmarked <img>.",
        "6. Keep takeaway and source visible; no content may overlap or be clipped.",
        "7. Reserve a clear bottom safe area for takeaway/source. Main content must end above the takeaway region.",
        "8. Follow the reference-style area budget: white/near-white surfaces >=80% of slide area; pale-blue fills (#D9EAF7/#EAF1F8) <=12%; deep-blue fills (#305496/#4472C4) <=8%.",
        "9. Do not default to blue-filled cards or panels. Normal modules, tables, charts, and cards are white; pale-blue is only for table headers, one selected row, one focus callout, or a template-required grouped object. Deep-blue is only for narrow emphasis.",
        "10. Prefer consulting report exhibits: tables, matrices, scorecards, charts, process lanes, value chains, and annotated evidence blocks. Avoid UI-dashboard card grids unless explicitly requested.",
        "11. Nav must be a thin directory row or breadcrumb, not equal-width blue buttons. Takeaway must be slim, not a large banner. Deck chrome must be stable across slides.",
        "",
        "## Semantic Markers",
        "",
        "- KPI/card/callout/diagram regions: `data-paopao-component=\"kpi-card|callout|diagram-node\"`",
        "- Tables: real `<table>` with `data-paopao-component=\"table\"`",
        "- Charts: `data-paopao-component=\"native-chart\"` + `data-chart` + `data-categories` + `data-series` with visible fallback",
        "- Process/chevron: `data-paopao-component=\"chevron-flow\"`; avoid SVG-only arrows",
        "",
    ])
    return "\n".join(parts)


def _extract_prompt_header(final_prompt_text: str) -> dict[str, str]:
    fields = {
        "fill_origin": "",
        "prompt_template": "",
        "layout_name": "",
        "title": "",
        "bottom": "",
        "source": "",
    }
    for raw_line in final_prompt_text.splitlines():
        line = raw_line.strip()
        upper = line.upper()
        if upper.startswith("FILL_ORIGIN:"):
            fields["fill_origin"] = line.split(":", 1)[1].strip()
        elif upper.startswith("PROMPT_TEMPLATE:"):
            fields["prompt_template"] = line.split(":", 1)[1].strip().strip('"')
        elif upper.startswith("LAYOUT_NAME:"):
            fields["layout_name"] = line.split(":", 1)[1].strip().strip('"')
        elif upper.startswith("TITLE:") and not fields["title"]:
            fields["title"] = line.split(":", 1)[1].strip().strip('"')
        elif upper.startswith("BOTTOM:") and not fields["bottom"]:
            fields["bottom"] = line.split(":", 1)[1].strip().strip('"')
        elif upper.startswith("SOURCE:") and not fields["source"]:
            fields["source"] = line.split(":", 1)[1].strip()
    return fields


def _match_zone_label(line: str, zone_names: list[str]) -> str | None:
    label = line.strip()
    if not label or ":" not in label:
        return None
    left = label.split(":", 1)[0].strip()
    left_upper = left.upper()
    for zone in zone_names:
        zone_upper = zone.upper()
        if left_upper == zone_upper or left_upper.startswith(zone_upper + " "):
            return zone
    return None


def _split_compact_items(text: str) -> list[str]:
    parts = []
    for part in re.split(r"\s*[;|]\s*", text or ""):
        cleaned = part.strip()
        if cleaned.startswith("- ") or cleaned.startswith("• "):
            cleaned = cleaned[2:].strip()
        parts.append(cleaned)
    return [part for part in parts if part]


def _number_lead(text: str) -> tuple[str, str]:
    match = re.match(r"^([+-]?\d[\d,.]*(?:\.\d+)?%?|[+-]?\d[\d,.]*(?:\.\d+)?)\s*(.*)$", text.strip())
    if not match:
        return "", text.strip()
    return match.group(1).strip(), match.group(2).strip(" :-")


def _chart_specs_from_text(text: str) -> list[dict[str, object]]:
    specs: list[dict[str, object]] = []
    for raw in re.split(r"\bChart\s+\d+\s*:", text or "", flags=re.IGNORECASE):
        chunk = raw.strip(" .")
        if not chunk:
            continue
        years = re.findall(r"\b(20\d{2})\b", chunk)
        numbers = re.findall(r"[+-]?\d[\d,]*(?:\.\d+)?%?", chunk)
        values = [n for n in numbers if n not in years]
        specs.append({
            "chart_type": "bar",
            "description": chunk,
            "categories": years[:4],
            "values": values[:6],
            "html_hint": "Use visible bars/lines with labels; do not replace this chart with plain text only.",
        })
    return specs


def _detail_rows_from_text(text: str) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    row_markers = [
        ("Key metrics", r"Key Metrics:\s*(.*?)(?=\s*Key Activities:|\s*Status/Risk:|$)"),
        ("Key activities", r"Key Activities:\s*(.*?)(?=\s*Status/Risk:|$)"),
        ("Status / risk", r"Status/Risk:\s*(.*)$"),
    ]
    for label, pattern in row_markers:
        match = re.search(pattern, text or "", flags=re.IGNORECASE | re.DOTALL)
        if match:
            role = _detail_row_role(label)
            raw_cells = _split_detail_cells(match.group(1))
            rows.append({
                "label": label,
                "role": role,
                "cells": [
                    _detail_cell_spec(cell, role, idx)
                    for idx, cell in enumerate(raw_cells)
                ],
            })
    return rows


def _detail_row_role(label: str) -> str:
    lower = label.lower()
    if "metric" in lower:
        return "kpi"
    if "risk" in lower or "status" in lower:
        return "risk"
    return "activity"


def _split_detail_cells(text: str) -> list[str]:
    raw = text or ""
    if "|" in raw:
        return [part.strip(" .") for part in raw.split("|") if part.strip(" .")]
    return _split_compact_items(raw)


def _extract_metric_tokens(text: str) -> list[str]:
    raw = text or ""
    patterns = [
        r"20\d{2}\s*-\s*20\d{2}",
        r"[+-]?\d[\d,]*(?:\.\d+)?%",
        r"\d[\d,]*(?:\.\d+)?\s*tCO2e(?:\s+per\s+vehicle)?",
        r"\d[\d,]*(?:\.\d+)?\s*vehicles?",
        r"\d[\d,]*(?:\.\d+)?\s*credits?",
        r"\d(?:\.\d+)?C\b",
    ]
    found: list[str] = []
    for pattern in patterns:
        for match in re.findall(pattern, raw, flags=re.IGNORECASE):
            token = re.sub(r"\s+", " ", str(match)).strip()
            if token and token not in found:
                found.append(token)
    return found


def _detail_cell_spec(text: str, role: str, index: int) -> dict[str, object]:
    clean = re.sub(r"\s+", " ", (text or "").strip(" ."))
    label = ""
    body = clean
    if ":" in clean:
        label, body = [part.strip() for part in clean.split(":", 1)]
    metrics = _extract_metric_tokens(body or clean)
    spec: dict[str, object] = {
        "text": clean,
        "cell_role": role,
        "stage_index": index + 1,
    }
    if label:
        spec["label"] = label
    if role == "kpi":
        spec.update({
            "primary_metric": metrics[0] if metrics else "",
            "secondary_metrics": metrics[1:3],
            "body": body,
            "emphasis": "large_metric",
            "visual_treatment": "large blue metric, short label, optional badge",
        })
    elif role == "risk":
        severity = "watch"
        lower = clean.lower()
        if any(word in lower for word in ["risk", "depends", "delayed", "exposed", "must"]):
            severity = "caution"
        spec.update({
            "severity": severity,
            "emphasis": "muted_caution",
            "visual_treatment": "small grey text with caution accent; do not compete with KPI row",
        })
    else:
        spec.update({
            "emphasis": "action_text",
            "visual_treatment": "compact body text; align to matching stage column",
        })
    return spec


def _zone_render_hint(zone: str, content: str, layout_name: str) -> dict[str, object]:
    zone_upper = zone.upper()
    hint: dict[str, object] = {"role": "text"}
    if "METRIC" in zone_upper and "CARD" in zone_upper:
        cards = []
        for item in _split_compact_items(content):
            metric, label = _number_lead(item)
            cards.append({"metric": metric, "label": label or item})
        hint.update({
            "role": "metric_cards",
            "cards": cards,
            "html_hint": "Render as equal-width metric cards with large metric, concise label, and a small badge only when useful.",
        })
    elif "CHART" in zone_upper:
        hint.update({
            "role": "charts",
            "charts": _chart_specs_from_text(content),
            "html_hint": "Build visible chart geometry from the extracted categories/values; avoid blank chart panels.",
        })
    elif "STAGE HEADER" in zone_upper:
        hint.update({
            "role": "process_stages",
            "stages": _split_compact_items(content),
            "html_hint": "Render as connected chevrons or a clear horizontal process row.",
        })
    elif "DETAIL GRID" in zone_upper:
        rows = _detail_rows_from_text(content)
        stage_count = max((len(row.get("cells", [])) for row in rows), default=0)
        hint.update({
            "role": "detail_grid",
            "visual_structure": "stage_matrix",
            "stage_count": stage_count,
            "rows": rows,
            "emphasis_rules": [
                "Key metrics row is the hero row: show primary_metric large in blue, with label/badge above and body below.",
                "Key activities row is supporting text: compact, aligned to the same stage columns, no large numbers.",
                "Status / risk row is muted: grey text, smaller type, optional caution accent; it should not dominate the slide.",
                "Columns must align to the stage header row; avoid a plain spreadsheet look by adding KPI badges and vertical accents.",
            ],
            "html_hint": "Render as a visually tiered stage matrix, not a generic table. Make KPI cells visually strongest, activity cells explanatory, and risk cells muted/cautionary.",
        })
    elif "COMMENTARY" in zone_upper:
        hint.update({
            "role": "commentary",
            "html_hint": "Render as a compact commentary band or cards, not a large empty panel.",
        })
    elif zone_upper in {"TITLE", "BOTTOM", "SOURCE"}:
        hint["role"] = zone_upper.lower()
    if layout_name:
        hint["layout_name"] = layout_name
    return hint


def _extract_final_prompt_execution_zones(
    final_prompt_text: str,
    zone_contract: dict[str, object],
    header: dict[str, str],
) -> list[dict[str, object]]:
    zones = zone_contract.get("zones")
    if not isinstance(zones, list):
        zones = []
    zone_names = [str(item.get("zone", "")).strip() for item in zones if isinstance(item, dict) and item.get("zone")]
    filled: dict[str, str] = {}
    current_zone: str | None = None
    capturing_zone: str | None = None
    buffer: list[str] = []

    def flush() -> None:
        nonlocal capturing_zone, buffer
        if capturing_zone:
            content = "\n".join(buffer).strip()
            if content:
                filled[capturing_zone] = content
        capturing_zone = None
        buffer = []

    hard_stop_prefixes = (
        "DESIGN:",
        "VISUAL STYLE:",
        "LAYOUT_NAME:",
        "WHEN_TO_USE:",
        "DATA_REQUIRES:",
        "PROMPT_TEMPLATE:",
        "FILL_ORIGIN:",
    )
    for raw_line in final_prompt_text.splitlines():
        stripped = raw_line.strip()
        matched_zone = _match_zone_label(stripped, zone_names)
        if matched_zone:
            flush()
            current_zone = matched_zone
            continue
        if stripped.upper() == "FILLED_CONTENT:":
            flush()
            capturing_zone = current_zone
            buffer = []
            continue
        if capturing_zone and stripped.upper().startswith(hard_stop_prefixes):
            flush()
            current_zone = None
            continue
        if capturing_zone:
            buffer.append(raw_line)
    flush()

    fallback = {
        "TITLE": header.get("title", ""),
        "BOTTOM": header.get("bottom", ""),
        "Source": header.get("source", ""),
        "SOURCE": header.get("source", ""),
    }
    layout_name = str(zone_contract.get("layout_name") or header.get("layout_name") or "")
    execution_zones: list[dict[str, object]] = []
    for zone in zone_names:
        content = filled.get(zone) or fallback.get(zone) or ""
        execution_zones.append({
            "zone": zone,
            "content": content,
            "render_hint": _zone_render_hint(zone, content, layout_name),
        })
    return execution_zones


def _write_compact_renderer_guide_impl(ctx: object, task_dir: Path) -> Path:
    _bind(ctx)
    path = html_compact_renderer_guide_path(task_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    text = compact_renderer_guide_text(read_text(SYSTEM_PROMPT))
    if not path.exists() or read_text(path) != text:
        path.write_text(text, encoding="utf-8")
    return path


def _write_html_compact_packet_impl(
    ctx: object,
    task_dir: Path,
    idx: int,
    expected: int,
    *,
    prompt_packet_id: str,
    required_marker: str,
    prompt_packet: Path,
    final_prompt: Path,
    html_path: Path,
    task_manifest: dict[str, object],
) -> Path:
    _bind(ctx)
    final_prompt_text = read_text(final_prompt)
    header = _extract_prompt_header(final_prompt_text)
    guide_path = _write_compact_renderer_guide_impl(ctx, task_dir)
    zone_contract = prompt_zone_contract_for_slide(task_dir, idx)
    execution_zones = _extract_final_prompt_execution_zones(final_prompt_text, zone_contract, header)
    provenance = {
        "schema": "paopao.html_compact_provenance.v1",
        "task_name": task_dir.name,
        "slide": idx,
        "full_prompt_packet_path": str(prompt_packet.relative_to(task_dir)),
        "full_prompt_packet_sha256": sha256_file(prompt_packet) if prompt_packet.exists() else "",
        "prompt_packet_id": prompt_packet_id,
        "renderer_compact_guide": str(guide_path.relative_to(task_dir)),
        "renderer_compact_guide_sha256": sha256_file(guide_path),
        "final_prompt_path": str(final_prompt.relative_to(task_dir)),
        "final_prompt_sha256": sha256_file(final_prompt),
        "prompt_meta": {
            "fill_origin": header.get("fill_origin", ""),
            "prompt_template": header.get("prompt_template", ""),
            "layout_name": header.get("layout_name", ""),
        },
    }
    provenance_path = html_compact_provenance_path(task_dir, idx)
    provenance_path.parent.mkdir(parents=True, exist_ok=True)
    provenance_text = json.dumps(provenance, ensure_ascii=False, indent=2, sort_keys=True)
    if not provenance_path.exists() or read_text(provenance_path) != provenance_text:
        provenance_path.write_text(provenance_text, encoding="utf-8")
    compact = {
        "schema": "paopao.html_compact_packet.v1",
        "purpose": "Executable HTML work order for the agent. Provenance/hash fields are intentionally kept out of this file.",
        "slide": idx,
        "expected_pages": expected,
        "language": str(task_manifest.get("language", "") or ""),
        "output_html_path": str(html_path.relative_to(task_dir)),
        "required_html_marker": required_marker,
        "renderer_compact_guide": str(guide_path.relative_to(task_dir)),
        "layout_name": header.get("layout_name", "") or str(zone_contract.get("layout_name") or ""),
        "navigation": compact_deck_navigation_contract(task_dir, idx, expected),
        "execution_zones": execution_zones,
        "composer_rules": [
            "Generate one complete HTML slide.",
            "Use required_html_marker in <head>.",
            "Use execution_zones only; do not reopen final_prompt, analysis, PDF/source, old drafts, or memory.",
            "Charts must use render_hint.charts and visible marks; no blank chart panels.",
            "Metric/process/grid zones must use render_hint.cards/stages/rows directly.",
            "For detail_grid, follow render_hint.visual_structure and emphasis_rules; do not render a flat spreadsheet.",
            f"Every required zone must appear as an element with {HTML_ZONE_ATTR}.",
            "Use compact renderer guide only unless blocked.",
        ],
    }
    text = "\n".join([
        "# Paopao Compact HTML Packet",
        "```json",
        json.dumps(compact, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
        "```",
        "",
    ])
    path = html_compact_packet_path(task_dir, idx)
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists() or read_text(path) != text:
        path.write_text(text, encoding="utf-8")
    return path


def _cmd_compact_html_packet_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    expected = expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Initialize task with --pages first.")
    if not is_html_source_only_task(task_dir):
        raise SystemExit("compact-html-packet is only valid for html_source_only tasks.")
    slides = list(range(1, expected + 1)) if args.all else [int(args.slide)]
    task_manifest = read_task_manifest(task_dir)
    system_prompt = ensure_workflow_file("SYSTEM_PROMPT.md")
    system_prompt_sha = sha256_file(system_prompt)
    written = []
    for idx in slides:
        if idx < 1 or idx > expected:
            raise SystemExit(f"--slide must be between 1 and {expected}")
        final_prompt = task_dir / "analysis" / f"final_prompt_{idx:02d}.md"
        if not final_prompt.exists() or not read_text(final_prompt).strip():
            raise SystemExit(f"final_prompt_{idx:02d}.md missing or empty")
        html_path = task_dir / "html" / f"slide{idx:02d}.html"
        prompt_packet = html_generation_request_path(task_dir, idx)
        if not prompt_packet.exists():
            raise SystemExit(
                f"Full locked packet missing for slide {idx}. Run generate-html first: "
                f"paopao_run.py generate-html --task-dir {task_dir} --slide {idx}"
            )
        prompt_packet_id = sha256_text("|".join([
            task_dir.name,
            str(idx),
            str(expected),
            system_prompt_sha,
            sha256_file(final_prompt),
            HTML_GENERATION_SOURCE,
        ]))
        required_marker = f'<meta name="{HTML_PROMPT_PACKET_META}" content="{prompt_packet_id}">'
        compact_path = _write_html_compact_packet_impl(
            ctx,
            task_dir,
            idx,
            expected,
            prompt_packet_id=prompt_packet_id,
            required_marker=required_marker,
            prompt_packet=prompt_packet,
            final_prompt=final_prompt,
            html_path=html_path,
            task_manifest=task_manifest,
        )
        written.append({
            "slide": idx,
            "compact_packet": str(compact_path.relative_to(task_dir)),
            "compact_packet_sha256": sha256_file(compact_path),
        })
    print(json.dumps({
        "ok": True,
        "task_dir": str(task_dir),
        "compact_renderer_guide": str(_write_compact_renderer_guide_impl(ctx, task_dir).relative_to(task_dir)),
        "slides": written,
    }, ensure_ascii=False, indent=2))
    return 0


def _cmd_compact_renderer_guide_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve() if args.task_dir else Path.cwd()
    text = compact_renderer_guide_text(read_text(SYSTEM_PROMPT))
    if args.output:
        out = Path(args.output)
        if not out.is_absolute():
            out = task_dir / out
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding="utf-8")
        print(json.dumps({"ok": True, "output": str(out), "sha256": sha256_file(out)}, ensure_ascii=False, indent=2))
    else:
        print(text)
    return 0


def _cmd_render_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    pptx = Path(args.pptx).resolve()
    html_files = [Path(p).resolve() for p in args.html] if args.html else html_files_from_task(task_dir)
    expected = expected_pages_from_task(task_dir) or len(html_files)
    preflight_issues: list[str] = []
    html_source_only = bool(args.html_source_only or is_html_source_only_task(task_dir))
    source_of_truth = HTML_BROWSER_SOURCE_OF_TRUTH if html_source_only else IMAGE2_SOURCE_OF_TRUTH
    check_html_files(task_dir, expected, preflight_issues, source_of_truth=source_of_truth)
    if html_source_only:
        check_html_generation_manifest(task_dir, expected, preflight_issues)
    if not html_source_only:
        check_image2_files(task_dir, expected, preflight_issues)
        check_html_reference_fidelity(task_dir, expected, preflight_issues, html_files=html_files)
    if preflight_issues:
        print(json.dumps({
            "task_dir": str(task_dir),
            "stage": "render-preflight",
            "expected_pages": expected,
            "ok": False,
            "issues": preflight_issues,
        }, indent=2, ensure_ascii=False))
        return 1
    reservation_id = reserve_quota(task_dir, len(html_files))

    renderer_path = ensure_workflow_file("renderer.py")
    cmd = [sys.executable, str(renderer_path), *map(str, html_files), "--pptx", str(pptx)]
    if args.pdf:
        cmd.extend(["--pdf", str(Path(args.pdf).resolve())])

    env = None
    proc = subprocess.run(cmd, cwd=str(PLUGIN_ROOT), text=True, capture_output=True, env=env)
    if proc.stdout:
        print(proc.stdout)
    if proc.stderr:
        print(proc.stderr, file=sys.stderr)
    pptx_generated = pptx.exists()
    succeeded = proc.returncode == 0 and pptx_generated
    finish_quota(reservation_id, succeeded)
    if pptx_generated:
        html_manifest: list[dict[str, object]] = []
        for idx, path in enumerate(html_files, 1):
            entry: dict[str, object] = {
                "path": str(path),
                "sha256": sha256_file(path),
            }
            if html_source_only:
                preview = task_dir / "qa" / "html_source" / f"slide-{idx:02d}.png"
                error = capture_html_preview(path, preview)
                if error:
                    print(error, file=sys.stderr)
                elif preview.exists():
                    entry["browser_preview_path"] = str(preview.relative_to(task_dir))
                    entry["browser_preview_sha256"] = sha256_file(preview)
            html_manifest.append(entry)
        manifest = {
            "rendered_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
            "renderer": str(RENDERER.resolve()),
            "task_dir": str(task_dir),
            "source_of_truth": HTML_BROWSER_SOURCE_OF_TRUTH if html_source_only else IMAGE2_SOURCE_OF_TRUTH,
            "html_source_only": bool(html_source_only),
            "pptx_path": str(pptx),
            "pptx_sha256": sha256_file(pptx),
            "html": html_manifest,
        }
        out = render_manifest_path(task_dir)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return proc.returncode



def cmd_record_commercial_render(ctx: object, args: object) -> int:
    return _cmd_record_commercial_render_impl(ctx, args)


def cmd_generate_html(ctx: object, args: object) -> int:
    return _cmd_generate_html_impl(ctx, args)


def cmd_register_html(ctx: object, args: object) -> int:
    return _cmd_register_html_impl(ctx, args)


def write_compact_renderer_guide(ctx: object, task_dir: Path) -> Path:
    return _write_compact_renderer_guide_impl(ctx, task_dir)


def write_html_compact_packet(ctx: object, task_dir: Path, idx: int, expected: int, **kwargs: object) -> Path:
    return _write_html_compact_packet_impl(ctx, task_dir, idx, expected, **kwargs)


def cmd_compact_html_packet(ctx: object, args: object) -> int:
    return _cmd_compact_html_packet_impl(ctx, args)


def cmd_compact_renderer_guide(ctx: object, args: object) -> int:
    return _cmd_compact_renderer_guide_impl(ctx, args)


def cmd_render(ctx: object, args: object) -> int:
    return _cmd_render_impl(ctx, args)

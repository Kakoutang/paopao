#!/usr/bin/env python3
"""Pipeline state and next-step controller for Paopao tasks."""

from __future__ import annotations

import json
import re
from pathlib import Path

_CTX_NAMES = ['Path', 'json', 're', 'is_html_source_only_task', 'check_html_source_analysis_files', 'check_analysis_files', 'check_image2_files', 'post_image_memory_boundary_issues', 'check_render_path_profile', 'commercial_render_path', 'check_pptx_file', 'check_commercial_render_contract', 'check_pipeline_contract', 'check_delivery_files', 'expected_pages_from_task', 'task_pipeline_mode', 'final_delivery_pass_path', 'check_html_files', 'HTML_BROWSER_SOURCE_OF_TRUTH', 'check_render_manifest', 'commercial_render_contract_path', 'IMAGE2_SOURCE_OF_TRUTH', 'DELIVERY_REVIEW_ACCEPTED_STATUSES', 'paopao_auth']


def _bind(ctx: object) -> None:
    for name in _CTX_NAMES:
        if name in {"Path", "json", "re"}:
            continue
        globals()[name] = getattr(ctx, name)


def _task_stage_issues_impl(
    ctx: object,
    task_dir: Path,
    expected: int,
    stage: str,
    pptx: Path | None = None,
) -> tuple[list[str], dict[str, object]]:
    _bind(ctx)
    issues: list[str] = []
    counts: dict[str, object] = {}
    html_source_only = is_html_source_only_task(task_dir)
    if stage == "analysis":
        if html_source_only:
            check_html_source_analysis_files(task_dir, expected, issues)
        else:
            check_analysis_files(task_dir, expected, issues)
    elif stage == "image2":
        if html_source_only:
            return issues, {"image2_reference_count": "not_required_for_html_source_only"}
        check_analysis_files(task_dir, expected, issues)
        counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
    elif stage == "memory_boundary":
        if html_source_only:
            return issues, {"memory_boundary": "not_required_for_html_source_only"}
        check_analysis_files(task_dir, expected, issues)
        counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
        issues.extend(post_image_memory_boundary_issues(task_dir, expected))
    elif stage == "html":
        if html_source_only:
            check_html_source_analysis_files(task_dir, expected, issues)
            counts["image2_reference_count"] = "not_required_for_html_source_only"
        else:
            check_analysis_files(task_dir, expected, issues)
            counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
        counts.update(check_render_path_profile(
            task_dir,
            expected,
            commercial_render_path(task_dir),
            issues,
        ))
    elif stage == "pptx":
        if pptx is None:
            pptx_files = [
                p for p in sorted((task_dir / "pptx").glob("*.pptx"))
                if p.is_file() and not p.name.startswith("~$")
            ]
            pptx = pptx_files[-1].resolve() if pptx_files else task_dir / "pptx" / "missing.pptx"
        if html_source_only:
            check_html_source_analysis_files(task_dir, expected, issues)
            pptx_summary = check_pptx_file(pptx, expected, issues)
            if pptx_summary is not None:
                counts["pptx"] = str(pptx)
                counts["pptx_summary"] = pptx_summary
            counts["commercial_render_contract"] = check_commercial_render_contract(task_dir, expected, pptx, issues)
            counts.update(check_render_path_profile(
                task_dir,
                expected,
                "html",
                issues,
                pptx=pptx,
                pptx_summary=pptx_summary,
            ))
        else:
            counts.update(check_pipeline_contract(task_dir, expected, pptx, issues))
    elif stage == "delivery":
        counts.update(check_delivery_files(task_dir, expected, issues))
    else:
        issues.append(f"Unknown task stage: {stage}")
    return issues, counts


def _task_controller_status_impl(
    ctx: object,
    task_dir: Path,
    expected: int,
    pptx: Path | None = None,
) -> dict[str, object]:
    _bind(ctx)
    if is_html_source_only_task(task_dir):
        stages = [
            {
                "id": "analysis",
                "label": "Analysis and locked Paopao prompts",
                "next_action": "Analysis employee may read PDF/source plus analysis/evidence_pool.json. First run extract-evidence-pool, then complete analysis_report.md, slide_story.json, prompt_selection_plan.json, prompt_selection_audit.md, and final_prompt_XX.md from the Paopao prompt library; then run check --stage analysis.",
            },
            {
                "id": "html",
                "label": "Signed HTML source pages",
                "next_action": "HTML employee reads only analysis/final_prompt_XX.md and renderer_compact_guide.md. Do not read PDF/source, analysis_report, full SKILL, full renderer_guide, prior drafts, or other pipeline materials. Run generate-html, compose HTML from the full final_prompt plus compact guide, then run register-html. compact_packet is optional economy/debug input only.",
            },
            {
                "id": "pptx",
                "label": "HTML-source-only PPTX render",
                "next_action": "Render employee reads only HTML files and command check results. Do not read PDF/source, analysis, final_prompt, compact/full packets, prior drafts, previews, or other pipeline materials. Run render --html-source-only, then record-commercial-render --source-of-truth html_browser_render. Do not export or inspect PPTX previews on the default path.",
            },
            {
                "id": "delivery",
                "label": "Final delivery",
                "next_action": "Delivery employee uses only delivery gates and command check results; do not reopen PDF/source, analysis, final_prompt, previews, prior drafts, or other pipeline materials. Run finalize-delivery after structural PPTX/render checks pass.",
            },
        ]
    else:
        stages = [
        {
            "id": "analysis",
            "label": "Analysis and slide story",
            "next_action": "Complete analysis_report.md, slide_story.json, prompt_selection_plan.json, final_prompt_XX.md, then run check --stage analysis.",
        },
        {
            "id": "image2",
            "label": "Selected Image2 references",
            "next_action": "Generate, register, review, and user-approve exactly one Image2 reference per requested slide; then run check --stage image2.",
        },
        {
            "id": "memory_boundary",
            "label": "Post-image memory boundary",
            "next_action": "Run forget-after-image2, then reconstruct only from the selected images.",
        },
        {
            "id": "html",
            "label": "Image-derived reconstruction source",
            "next_action": "Open each selected Image2 reference and hand-author custom HTML/CSS for renderer.py.",
        },
        {
            "id": "pptx",
            "label": "Rendered PPTX and QA",
            "next_action": "Render the declared commercial path, open the real PPTX in PowerPoint, complete PowerPoint and fidelity reviews, then run check --stage pipeline.",
        },
        {
            "id": "delivery",
            "label": "Final delivery",
            "next_action": "Run finalize-delivery. Do not reply with delivery links until qa/final_delivery_pass.json exists and check --stage delivery passes.",
        },
        ]
    stage_results: list[dict[str, object]] = []
    first_blocked: dict[str, object] | None = None
    for stage in stages:
        issues, counts = _task_stage_issues_impl(ctx, task_dir, expected, str(stage["id"]), pptx)
        result = {
            "id": stage["id"],
            "label": stage["label"],
            "ok": not issues,
            "issue_count": len(issues),
            "issues": issues[:20],
            "counts": counts,
        }
        stage_results.append(result)
        if issues and first_blocked is None:
            first_blocked = {
                "stage": stage["id"],
                "label": stage["label"],
                "next_action": stage["next_action"],
                "issues": issues[:20],
            }
            break

    deliverable = first_blocked is None
    return {
        "schema": "paopao.task_controller_status.v1",
        "task_dir": str(task_dir),
        "expected_pages": expected,
        "deliverable": deliverable,
        "blocked": first_blocked,
        "completed_stages": [stage["id"] for stage in stage_results if stage.get("ok")],
        "stage_results": stage_results,
        "policy": (
            "Paopao tasks are not deliverable by agent judgment. Delivery is valid only through "
            "finalize-delivery and a current qa/final_delivery_pass.json."
        ),
    }


def _cmd_run_task_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")
    pptx = Path(args.pptx).resolve() if args.pptx else None
    status = _task_controller_status_impl(ctx, task_dir, expected, pptx)
    out = task_dir / "qa" / "task_controller_status.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(status, indent=2, ensure_ascii=False))
    return 0 if status["deliverable"] else 1


def _pipeline_step_state_impl(ctx: object, task_dir: Path, expected: int) -> dict[str, object]:
    _bind(ctx)
    """Determine the precise pipeline step the task is at and what to do next."""

    state: dict[str, object] = {
        "task_dir": str(task_dir),
        "expected_pages": expected,
        "pipeline_mode": task_pipeline_mode(task_dir),
    }

    delivery_dir = task_dir / "delivery"
    delivery_pptx = sorted(delivery_dir.glob("*.pptx")) if delivery_dir.exists() else []
    if final_delivery_pass_path(task_dir).exists() and delivery_pptx:
        state["step"] = "finalize"
        state["step_number"] = 99
        state["instruction"] = "All checks passed. Delivery is ready."
        return state

    if is_html_source_only_task(task_dir):
        analysis_issues: list[str] = []
        check_html_source_analysis_files(task_dir, expected, analysis_issues)
        if analysis_issues:
            state["step"] = "analysis"
            state["step_number"] = 1
            state["instruction"] = (
                "Analysis employee context boundary: you may read PDF/source files and analysis/evidence_pool.json. "
                "Downstream employees must not reopen PDF/source.\n"
                "Produce Paopao prompt-library artifacts:\n"
                "  0. Run: paopao_run.py extract-evidence-pool --task-dir <dir>\n"
                "  1. analysis/analysis_report.md — synthesized report and source-backed conclusions; cite evidence_pool facts where possible\n"
                "  2. analysis/slide_story.json — one role/brief per slide\n"
                "  3. Run: paopao_run.py plan-prompts --task-dir <dir>\n"
                "     This selects a prompt template for each page and outputs fill_zones for each.\n"
                "  4. For EACH page, read that page's fill_zones from prompt_selection_plan.json.\n"
                "     Each zone has a name and instructions describing what data to fill.\n"
                "     Extract matching data from analysis_report.md for EVERY zone — not just TITLE.\n"
                "     Then run:\n"
                "     paopao_run.py fill-prompt-template --template <selected_template.md> "
                "--fills '<zone JSON>' --output analysis/final_prompt_XX.md\n"
                "     The --fills JSON must include ALL zones listed in fill_zones.\n"
                "     Do NOT write final_prompt_XX.md by hand.\n"
                "  5. analysis/prompt_selection_audit.md — why each template was selected\n"
                "Do not create analysis/html_prompt_XX.md.\n"
                "When done, run: paopao_run.py next --task-dir <dir>"
            )
            state["issues"] = analysis_issues
            return state

        html_issues: list[str] = []
        html_count = check_html_files(
            task_dir,
            expected,
            html_issues,
            source_of_truth=HTML_BROWSER_SOURCE_OF_TRUTH,
        )
        if html_count < expected or html_issues:
            done_slides = sorted(
                int(p.stem.replace("slide", ""))
                for p in (task_dir / "html").glob("slide*.html")
                if p.is_file() and re.match(r"slide\d+$", p.stem)
            )
            missing = [i for i in range(1, expected + 1) if i not in done_slides]
            slide_idx = missing[0] if missing else 1
            state["step"] = "html_source"
            state["step_number"] = 2
            state["current_slide"] = slide_idx
            agent_prompt_rel = f"qa/html_generation_requests/agent_prompt_{slide_idx:02d}.md"
            state["instruction"] = (
                f"Run these 3 commands in order for slide {slide_idx}:\n"
                f"  1. paopao_run.py generate-html --task-dir {task_dir} --slide {slide_idx}\n"
                f"  2. Spawn an Agent with this prompt: \"Read {task_dir}/{agent_prompt_rel} and follow its instructions. Do not read any other files.\"\n"
                f"  3. paopao_run.py register-html --task-dir {task_dir} --slide {slide_idx}\n"
                f"Then run: paopao_run.py next --task-dir {task_dir}"
            )
            state["agent_prompt_file"] = str(task_dir / agent_prompt_rel)
            state["subagent_prompt"] = (
                f"Read {task_dir}/{agent_prompt_rel} and follow its instructions exactly. "
                "Do not read any other files."
            )
            if html_issues:
                state["issues"] = html_issues
            return state

        pptx_dir = task_dir / "pptx"
        pptx_files = sorted(p for p in pptx_dir.glob("*.pptx") if p.is_file() and not p.name.startswith("~$"))
        pptx = pptx_files[-1] if pptx_files else pptx_dir / "deck.pptx"
        render_issues: list[str] = []
        if pptx.exists():
            check_render_manifest(task_dir, pptx, render_issues)
        if not pptx.exists() or render_issues:
            state["step"] = "render"
            state["step_number"] = 3
            state["instruction"] = (
                "Render employee context boundary: inspect only HTML files and command check results. "
                "Do not read PDF/source, analysis, final_prompt, prompt packets, prior drafts, or other pipeline materials.\n"
                "Convert HTML to editable PPTX through the HTML-source-only renderer:\n"
                "  1. Run: paopao_run.py render --task-dir <dir> --pptx pptx/deck.pptx --html-source-only\n"
                "  2. Run: paopao_run.py record-commercial-render --task-dir <dir> --render-path html "
                "--source-of-truth html_browser_render --pptx pptx/deck.pptx\n"
                "This binds the PPTX to the current HTML files and render hashes. Do not export, open, or compare previews on the default path."
            )
            if render_issues:
                state["issues"] = render_issues
            return state

        contract_issues: list[str] = []
        check_commercial_render_contract(task_dir, expected, pptx, contract_issues)
        if contract_issues:
            state["step"] = "record_render"
            state["step_number"] = 3
            state["instruction"] = (
                "Render employee context boundary: inspect only HTML files and command check results. "
                "Do not read PDF/source, analysis, final_prompt, prompt packets, prior drafts, or other pipeline materials.\n"
                "Record the HTML-source-only commercial render contract:\n"
                "  Run: paopao_run.py record-commercial-render --task-dir <dir> --render-path html "
                "--source-of-truth html_browser_render --pptx pptx/deck.pptx\n"
                "Do not export, open, or compare previews on the default path."
            )
            state["issues"] = contract_issues
            return state

        delivery_dir = task_dir / "delivery"
        delivery_pptx = sorted(delivery_dir.glob("*.pptx")) if delivery_dir.exists() else []
        if delivery_pptx:
            state["step"] = "finalize"
            state["step_number"] = 5
            state["instruction"] = "All checks passed. Delivery is ready."
            return state

        state["step"] = "finalize"
        state["step_number"] = 4
        state["instruction"] = (
            "Delivery employee context boundary: inspect only command check results. "
            "Do not read PDF/source, analysis, final_prompt, prompt packets, prior drafts, or other pipeline materials.\n"
            "Default HTML-source-only production is one-pass: do not export, open, or compare previews.\n"
            "Run:\n"
            "  paopao_run.py finalize-delivery --task-dir <dir> --pptx pptx/deck.pptx"
        )
        return state

    analysis_issues: list[str] = []
    check_analysis_files(task_dir, expected, analysis_issues)
    if analysis_issues:
        state["step"] = "analysis"
        state["step_number"] = 1
        state["instruction"] = (
            "Read source materials and produce:\n"
            "  1. analysis/analysis_report.md — key data with sources\n"
            "  2. analysis/slide_story.json — arc + per-page conclusions\n"
            "  3. Run: paopao_run.py plan-prompts --task-dir <dir>\n"
            "     This selects a prompt template for each page and outputs fill_zones for each.\n"
            "  4. For EACH page, read that page's fill_zones from prompt_selection_plan.json.\n"
            "     Each zone has a name and instructions describing what data to fill.\n"
            "     Extract matching data from analysis_report.md for EVERY zone — not just TITLE.\n"
            "     Then run:\n"
            "     paopao_run.py fill-prompt-template --template <selected_template.md> "
            "--fills '<zone JSON>' --output analysis/final_prompt_XX.md\n"
            "     The --fills JSON must include ALL zones listed in fill_zones.\n"
            "     Do NOT write final_prompt_XX.md by hand.\n"
            "  5. analysis/prompt_selection_audit.md\n"
            "When done, run: paopao_run.py next --task-dir <dir>"
        )
        state["issues"] = analysis_issues
        return state

    image2_issues: list[str] = []
    image2_count = check_image2_files(task_dir, expected, image2_issues)
    def _is_style_review_issue(issue: str) -> bool:
        lowered = issue.lower()
        return "image2_style_review" in lowered or "image2 style review" in lowered

    def _is_user_review_issue(issue: str) -> bool:
        lowered = issue.lower()
        return (
            "image2_user_review" in lowered
            or "user image review" in lowered
            or "user_approved" in lowered
        )

    style_review_issues = [i for i in image2_issues if _is_style_review_issue(i)]
    user_review_issues = [i for i in image2_issues if _is_user_review_issue(i)]
    image2_provenance_issues = [
        i for i in image2_issues
        if not _is_style_review_issue(i) and not _is_user_review_issue(i)
    ]
    if image2_count < expected or image2_provenance_issues:
        done_slides = sorted(
            int(p.stem.split("_")[-1])
            for p in (task_dir / "image2").glob("image2_reference_*.png")
            if p.is_file()
        )
        missing = [i for i in range(1, expected + 1) if i not in done_slides]
        if not (task_dir / "image2" / "image2_generation_manifest.json").exists():
            state["step"] = "image2_prepare"
            state["step_number"] = 2
            state["instruction"] = (
                "Run: paopao_lab.py prepare-image2-prompts --task-dir <dir>\n"
                "This locks the per-slide prompts for image generation."
            )
        elif missing:
            slide = missing[0]
            state["step"] = "image2_generate"
            state["step_number"] = 2
            state["current_slide"] = slide
            state["instruction"] = (
                f"Generate Image2 for slide {slide}:\n"
                f"  1. Run: paopao_lab.py start-image2-generation --task-dir <dir> --slide {slide}\n"
                f"  2. Use a controlled generator that reads image2/generation_request_{slide:02d}.json prompt_text exactly. "
                f"Do not paste, summarize, compress, translate, or rewrite the prompt in chat.\n"
                f"  3. Register the actual original image-generation output only. Do not register a manually redrawn, simplified, "
                f"smoke-test, screenshot, or replacement reference.\n"
                f"  4. Register: paopao_lab.py register-image2-reference --task-dir <dir> "
                f"--slide {slide} --image <output_path> --generation-request image2/generation_request_{slide:02d}.json "
                f"--generated-prompt-sha256 <sha> --source image_gen_builtin --tool-call-id <id> "
                f"--controlled-generation image2/controlled_generation_{slide:02d}.json\n"
                f"When done, run: paopao_lab.py next --task-dir <dir>"
            )
        else:
            state["step"] = "image2_fix"
            state["step_number"] = 2
            state["instruction"] = "Fix Image2 registration issues listed below, then run next again."
            state["issues"] = image2_provenance_issues
        return state

    if style_review_issues:
        state["step"] = "image2_style_review"
        state["step_number"] = 3
        state["instruction"] = (
            "Open every selected Image2 reference and record actual visual style evidence in "
            "qa/image2_style_review.json.\n"
            "Check each slide for: aspect_ratio_16_9, house_style_reference, palette_discipline, "
            "clean_background, linework_and_borders, material_simplicity, title_weight, "
            "module_density, takeaway, color_hierarchy, and icons.\n"
            "Do not proceed from the prompt text; inspect the actual generated images.\n"
            "When done, run: paopao_lab.py next --task-dir <dir>"
        )
        state["issues"] = style_review_issues
        return state

    user_review_path = task_dir / "qa" / "image2_user_review.json"
    if not user_review_path.exists() or user_review_issues:
        state["step"] = "image2_user_review"
        state["step_number"] = 4
        state["instruction"] = (
            "Show all Image2 reference images to the user for approval.\n"
            "After user responds, run:\n"
            "  paopao_lab.py record-image2-user-review --task-dir <dir> --approved <yes|no> --feedback '<text>'\n"
            "Then run: paopao_lab.py next --task-dir <dir>"
        )
        if user_review_issues:
            state["issues"] = user_review_issues
        return state
    else:
        try:
            review_data = json.loads(user_review_path.read_text(encoding="utf-8"))
            if review_data.get("user_approved") is not True:
                state["step"] = "image2_user_review"
                state["step_number"] = 4
                state["instruction"] = (
                    "User requested changes. Regenerate the affected Image2 references, "
                    "then record a new approval.\n"
                    f"User feedback: {review_data.get('user_feedback', '')}"
                )
                return state
        except Exception:
            pass

    boundary_path = task_dir / "qa" / "post_image_memory_boundary.json"
    if not boundary_path.exists():
        state["step"] = "memory_boundary"
        state["step_number"] = 5
        state["instruction"] = (
            "Run: paopao_lab.py forget-after-image2 --task-dir <dir>\n"
            "This enforces the memory boundary — after this point, reconstruction "
            "must be based solely on the selected Image2 reference images."
        )
        return state

    html_issues: list[str] = []
    html_count = check_html_files(task_dir, expected, html_issues)
    if html_count < expected:
        done_slides = sorted(
            int(p.stem.replace("slide", ""))
            for p in (task_dir / "html").glob("slide*.html")
            if p.is_file() and re.match(r"slide\d+$", p.stem)
        )
        missing = [i for i in range(1, expected + 1) if i not in done_slides]
        slide_idx = missing[0] if missing else 1
        state["step"] = "direct_painter"
        state["step_number"] = 6
        state["current_slide"] = slide_idx
        state["instruction"] = (
            f"Create custom HTML/CSS for slide {slide_idx} from the approved Image2 reference:\n"
            f"  1. Open image2/image2_reference_{slide_idx:02d}.png and use it as the only visual source.\n"
            f"  2. Hand-author html/slide{slide_idx:02d}.html to match the reference page: nav, title, content modules, "
            f"charts/tables/icons/connectors, takeaway, source, colors, spacing, and typography.\n"
            f"  3. Use editable HTML primitives that renderer.py can convert: text boxes, divs, tables, chart data blocks, "
            f"and marked small image/icon assets when needed. Never use a whole-slide screenshot.\n"
            f"  4. Do not write python-pptx. Do not use compile_object_graph, object_graph, observation, visual_inventory, "
            f"or layout_plan as the production path.\n"
            f"When done, run: paopao_lab.py next --task-dir <dir>"
        )
        return state
    if html_issues:
        state["step"] = "direct_painter_fix"
        state["step_number"] = 6
        state["instruction"] = (
            "Fix the custom HTML/CSS issues below by reopening the approved Image2 references and editing the relevant "
            "html/slideXX.html files. Do not create observation, visual_inventory, layout_plan, object_graph, or "
            "python-pptx painter files for production."
        )
        state["issues"] = html_issues
        return state

    pptx_dir = task_dir / "pptx"
    pptx_files = sorted(p for p in pptx_dir.glob("*.pptx") if p.is_file() and not p.name.startswith("~$"))
    pptx = pptx_files[-1] if pptx_files else pptx_dir / "deck.pptx"
    render_issues: list[str] = []
    if pptx.exists():
        check_render_manifest(task_dir, pptx, render_issues)
    if not pptx.exists() or render_issues:
        state["step"] = "render"
        state["step_number"] = 7
        state["instruction"] = (
            "Convert the hand-authored HTML/CSS into editable PPTX through the production renderer:\n"
            "  1. Run: paopao_lab.py render --task-dir <dir> --pptx pptx/deck.pptx\n"
            "  2. Then run: paopao_lab.py record-commercial-render --task-dir <dir> --render-path html --pptx pptx/deck.pptx\n"
            "renderer.py is the format converter for this production path. Do not replace it with python-pptx or "
            "compile_object_graph.\n"
            "When done, run: paopao_lab.py next --task-dir <dir>"
        )
        if render_issues:
            state["issues"] = render_issues
        return state

    contract_issues: list[str] = []
    check_commercial_render_contract(task_dir, expected, pptx, contract_issues)
    if contract_issues:
        state["step"] = "record_render"
        state["step_number"] = 7
        state["instruction"] = (
            "Record the production render contract before PPTX preview export:\n"
            "  Run: paopao_run.py record-commercial-render --task-dir <dir> --render-path html --pptx pptx/deck.pptx\n"
            "This binds the final PPTX hash to the approved Image2 references and the HTML renderer path."
        )
        state["issues"] = contract_issues
        return state

    actual_dir = task_dir / "qa" / "pptx_actual"
    actual_pngs = sorted(actual_dir.glob("slide-*.png"))
    if len(actual_pngs) < expected:
        state["step"] = "pptx_export"
        state["step_number"] = 8
        state["instruction"] = (
            "Export PPTX slide previews:\n"
            "  1. Open the generated PPTX in PowerPoint\n"
            "  2. Export each slide as PNG to qa/pptx_actual/slide-1.png, slide-2.png, ...\n"
            "  3. Also export a PDF to qa/pptx_actual/actual.pdf\n"
            "When done, run: paopao_run.py next --task-dir <dir>"
        )
        return state

    fidelity_path = task_dir / "qa" / "fidelity_review.json"
    fidelity_ok = True
    if not fidelity_path.exists():
        fidelity_ok = False
    else:
        try:
            fidelity_data = json.loads(fidelity_path.read_text(encoding="utf-8"))
            slides = fidelity_data.get("slides", [])
            evidences = [str(s.get("evidence", "")) for s in slides if isinstance(s, dict)]
            if len(set(evidences)) < len(evidences):
                fidelity_ok = False
            for s in slides:
                if not isinstance(s, dict):
                    fidelity_ok = False
                    continue
                status = str(s.get("status", "")).lower().strip()
                if status not in DELIVERY_REVIEW_ACCEPTED_STATUSES:
                    fidelity_ok = False
                if s.get("compared") is not True:
                    fidelity_ok = False
                if len(str(s.get("evidence", ""))) < 80:
                    fidelity_ok = False
        except Exception:
            fidelity_ok = False

    if not fidelity_ok:
        state["step"] = "fidelity_review"
        state["step_number"] = 9
        state["instruction"] = (
            "Compare EACH slide's PPTX preview against its Image2 reference:\n"
            "  For each slide (1 to {expected}):\n"
            "    1. Open image2/image2_reference_XX.png\n"
            "    2. Open qa/pptx_actual/slide-X.png\n"
            "    3. Write a UNIQUE, SPECIFIC comparison for this slide:\n"
            "       - What matches well\n"
            "       - What differs (position, size, color, text, missing elements)\n"
            "       - Whether it needs fixing\n"
            "  Save as qa/fidelity_review.json with per-slide evidence.\n"
            "  EACH slide MUST have different evidence text — no copy-paste.\n"
            "  Evidence must be at least 80 characters per slide.\n"
            "If any slide has issues, fix the relevant html/slideXX.html, re-render, re-export, then re-review.\n"
            "When done, run: paopao_run.py next --task-dir <dir>"
        ).format(expected=expected)
        return state

    ppt_review_path = task_dir / "qa" / "powerpoint_review.json"
    if not ppt_review_path.exists():
        state["step"] = "powerpoint_review"
        state["step_number"] = 10
        state["instruction"] = (
            "Open the PPTX in PowerPoint and inspect the editing interface:\n"
            "  For each slide:\n"
            "    - Is text editable and properly formatted?\n"
            "    - Are charts native (double-click opens data editor)?\n"
            "    - Are there overlapping elements or clipped text?\n"
            "    - Does the nav text appear centered?\n"
            "  Save as qa/powerpoint_review.json with actual_pptx_opened: true.\n"
            "  Each slide must have unique evidence.\n"
            "When done, run: paopao_run.py next --task-dir <dir>"
        )
        return state

    pipeline_issues: list[str] = []
    pipeline_counts: dict[str, object] = {}
    pipeline_counts.update(check_pipeline_contract(task_dir, expected, pptx, pipeline_issues))
    if pipeline_issues:
        fixable = [i for i in pipeline_issues if "similarity" in i.lower() or "fidelity" in i.lower()]
        structural = [i for i in pipeline_issues if i not in fixable]
        if fixable and not structural:
            state["step"] = "iterate_pptx"
            state["step_number"] = 11
            state["instruction"] = (
                "Similarity scores are below threshold. Fix the custom HTML/CSS output:\n"
                "  1. Open EACH failing slide's Image2 reference and PPTX preview side by side\n"
                "  2. Identify specific differences (missing elements, wrong positions, wrong text)\n"
                "  3. Update the relevant html/slideXX.html to fix those differences\n"
                "  4. Regenerate pptx/deck.pptx through paopao_run.py render, not python-pptx or compile_object_graph\n"
                "  5. Re-export PNGs, redo fidelity review\n"
                "  5. Run: paopao_run.py next --task-dir <dir>\n"
                "Failing checks:"
            )
            state["issues"] = fixable
        else:
            state["step"] = "pipeline_fix"
            state["step_number"] = 11
            state["instruction"] = "Fix pipeline issues listed below."
            state["issues"] = pipeline_issues
        return state

    state["step"] = "finalize"
    state["step_number"] = 12
    state["instruction"] = (
        "All checks passed. Finalize delivery:\n"
        "  Run: paopao_run.py finalize-delivery --task-dir <dir> --pptx <path>\n"
        "This runs the final gate and packages the deliverable."
    )
    return state


def _fetch_server_notices() -> dict[str, str]:
    result: dict[str, str] = {}
    try:
        data = paopao_auth.request_json("GET", f"{paopao_auth.server_url()}/health")
        notice = data.get("update_notice")
        if isinstance(notice, dict):
            msg = str(notice.get("message", "")).strip()
            if msg:
                result["update_notice"] = msg
    except Exception:
        pass
    return result


def _cmd_next_impl(ctx: object, args: object) -> int:
    _bind(ctx)
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")

    state = _pipeline_step_state_impl(ctx, task_dir, expected)

    notices = _fetch_server_notices()
    if notices.get("update_notice"):
        state["update_notice"] = notices["update_notice"]
        state["update_notice_instruction"] = (
            "STOP: show the update_notice message to the user verbatim before continuing. "
            "Do not start generating until the user has updated or explicitly chosen to skip."
        )

    state_path = task_dir / "qa" / "pipeline_state.json"
    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(state, indent=2, ensure_ascii=False))
    return 0



def task_stage_issues(ctx: object, task_dir: Path, expected: int, stage: str, pptx: Path | None = None) -> tuple[list[str], dict[str, object]]:
    return _task_stage_issues_impl(ctx, task_dir, expected, stage, pptx)


def task_controller_status(ctx: object, task_dir: Path, expected: int, pptx: Path | None = None) -> dict[str, object]:
    return _task_controller_status_impl(ctx, task_dir, expected, pptx)


def pipeline_step_state(ctx: object, task_dir: Path, expected: int) -> dict[str, object]:
    return _pipeline_step_state_impl(ctx, task_dir, expected)


def cmd_run_task(ctx: object, args: object) -> int:
    return _cmd_run_task_impl(ctx, args)


def cmd_next(ctx: object, args: object) -> int:
    return _cmd_next_impl(ctx, args)

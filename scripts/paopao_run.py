#!/usr/bin/env python3
"""Thin public bootstrap for Paopao."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from paopao_file_manifest import AUTHORIZED_RUNTIME_FILES, WORKFLOW_DESTINATION_RELS

PLUGIN_ROOT = Path(__file__).resolve().parents[1]
AUTHORIZED_WORKFLOW_FILES = AUTHORIZED_RUNTIME_FILES
BUNDLE_CHUNK_SIZE = 80
OPTIONAL_WORKFLOW_FILES = {"paopao_delivery_safety.py"}
CATALOG_EMPTY_ERROR = (
    "Authorized prompt catalog is empty. Paopao service did not return any "
    "page templates for this access token; please update Paopao service package."
)
SERVER_JOB_ARTIFACTS = {"deck.pptx", "analysis_report.md"}
JOB_SPEC_SCHEMA_VERSION = "paopao.job_spec.v1"
JOB_SPEC_INTENTS = [
    "summary_scr",
    "trend_chart",
    "comparison",
    "process",
    "kpi_summary",
    "table_analysis",
    "structure_map",
    "decision_memo",
]
INTENT_TITLES = {
    "summary_scr": "Executive logic",
    "trend_chart": "Signal trend",
    "comparison": "Option comparison",
    "process": "Execution path",
    "kpi_summary": "KPI dashboard",
    "table_analysis": "Evidence matrix",
    "structure_map": "Structure map",
    "decision_memo": "Decision memo",
}


def _s(*parts: str) -> str:
    return "".join(parts)


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


def workflow_file_missing(exc: Exception) -> bool:
    text = str(exc)
    return "HTTP 404" in text and "Workflow file not found" in text


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


def fetch_workflow_bundle(names: list[str]) -> list[str]:
    paopao_auth = _load_sibling("paopao_auth")
    written: list[str] = []
    for start in range(0, len(names), BUNDLE_CHUNK_SIZE):
        chunk = names[start:start + BUNDLE_CHUNK_SIZE]
        try:
            result = paopao_auth.fetch_workflow_bundle(chunk)
        except paopao_auth.AuthError as exc:
            if not workflow_file_missing(exc):
                raise SystemExit(str(exc)) from exc
            result = {"files": []}
            for name in chunk:
                try:
                    result["files"].append(paopao_auth.fetch_workflow_file(name))
                except paopao_auth.AuthError as file_exc:
                    if name in OPTIONAL_WORKFLOW_FILES and workflow_file_missing(file_exc):
                        continue
                    raise SystemExit(str(file_exc)) from file_exc
        for item in result.get("files", []):
            name = str(item.get("name", "")).strip()
            content = str(item.get("content", "")).strip()
            if not content:
                raise SystemExit(f"Workflow file is empty: {name}")
            if name in WORKFLOW_DESTINATION_RELS:
                target = PLUGIN_ROOT / WORKFLOW_DESTINATION_RELS[name]
            elif name.endswith(".md") and "/" not in name and "\\" not in name and ".." not in name:
                target = PLUGIN_ROOT / "prompts" / name
            else:
                raise SystemExit(f"Unknown workflow file: {name}")
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(content + "\n", encoding="utf-8")
            written.append(str(target.relative_to(PLUGIN_ROOT)))
    return written


def assert_downloaded_prompts(expected: list[str], written: list[str]) -> None:
    written_prompts = {
        Path(label).name
        for label in written
        if Path(label).parent.name == "prompts" and Path(label).suffix == ".md"
    }
    missing = sorted(set(expected) - written_prompts)
    if missing:
        sample = ", ".join(missing[:8])
        suffix = "..." if len(missing) > 8 else ""
        raise SystemExit(
            "Authorized prompt library did not download completely: "
            f"missing {len(missing)} template(s): {sample}{suffix}"
        )


def authorized_prompt_names(*, full_library: bool = False) -> list[str]:
    paopao_auth = _load_sibling("paopao_auth")
    try:
        catalog = paopao_auth.fetch_prompt_catalog()
    except paopao_auth.AuthError as exc:
        raise SystemExit(str(exc)) from exc
    names: list[str] = []
    for item in catalog.get("prompts", []):
        name = str(item.get("template", "")).strip()
        if name.endswith(".md") and "/" not in name and "\\" not in name and ".." not in name:
            # Render already returns only the templates this token may use.
            # The public bootstrap should not downgrade paid users to the
            # starter/free prompt subset.
            names.append(name)
    if not names:
        raise SystemExit(CATALOG_EMPTY_ERROR)
    return names


def summarize(paths: list[str], sample_size: int = 12) -> dict[str, object]:
    return {
        "count": len(paths),
        "sample": paths[:sample_size],
        "truncated": len(paths) > sample_size,
    }


def slugify(name: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9._-]+", "-", str(name or "").strip()).strip("-").lower()
    return slug or "paopao-task"


def current_plan() -> str:
    paopao_auth = _load_sibling("paopao_auth")
    try:
        status = paopao_auth.status()
    except paopao_auth.AuthError as exc:
        raise SystemExit(str(exc)) from exc
    license_data = status.get("license", {}) if isinstance(status.get("license", {}), dict) else {}
    return str(license_data.get("plan", "") or "").strip().lower()


def should_use_server_job() -> bool:
    plan = current_plan()
    return plan.startswith("free_preview") or plan.startswith("starter")


def extract_text_from_source(path_value: str) -> str:
    if not path_value:
        return ""
    path = Path(path_value).expanduser().resolve()
    if not path.exists() or not path.is_file():
        raise SystemExit(f"Source file not found: {path}")
    if path.suffix.lower() in {".txt", ".md", ".csv"}:
        return path.read_text(encoding="utf-8", errors="replace")
    if path.suffix.lower() == ".pdf":
        tool = shutil.which("pdftotext")
        if tool:
            result = subprocess.run(
                [tool, str(path), "-"],
                check=False,
                capture_output=True,
                text=True,
                timeout=60,
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        return f"PDF source file provided locally: {path.name}. Text extraction was unavailable in this environment."
    raw = path.read_bytes()[:120000]
    return raw.decode("utf-8", errors="ignore")


def clip_text(text: str, limit: int) -> str:
    clean = re.sub(r"\s+", " ", str(text or "")).strip()
    return clean if len(clean) <= limit else clean[:limit - 3].rstrip() + "..."


def compact_lines(text: str, limit: int = 18, item_limit: int = 112) -> list[str]:
    clean = re.sub(r"\s+", " ", text or "").strip()
    if not clean:
        return []
    pieces = re.split(r"(?<=[。！？.!?])\s+|[;\n]+", clean)
    lines: list[str] = []
    for piece in pieces:
        item = piece.strip(" -•\t")
        if not item:
            continue
        lines.append(clip_text(item, item_limit))
        if len(lines) >= limit:
            break
    return lines


def _sample_numbers(seed: int, count: int) -> list[float]:
    return [float(max(1, ((seed + i * 3) % 9) + 2)) for i in range(count)]


def build_job_spec(source_text: str, pages: int, focus: str, language: str) -> dict[str, object]:
    lines = compact_lines(source_text or focus, limit=max(24, pages * 3))
    if not lines:
        lines = [focus or "Paopao direction deck"]
    spec_pages: list[dict[str, object]] = []
    for idx in range(1, pages + 1):
        intent = JOB_SPEC_INTENTS[(idx - 1) % len(JOB_SPEC_INTENTS)]
        start = (idx - 1) * 2
        bullets = lines[start:start + 5] or lines[:5]
        takeaway = clip_text(bullets[0] if bullets else (focus or f"Page {idx} key message"), 170)
        page: dict[str, object] = {
            "intent": intent,
            "title": clip_text(f"Page {idx}: {INTENT_TITLES.get(intent, 'Direction')}", 86),
            "takeaway": takeaway,
            "bullets": [clip_text(line, 112) for line in bullets[:5]],
            "source": "Source: user-provided local material",
        }
        if intent == "summary_scr":
            page["scr"] = {
                "situation": clip_text(bullets[0] if len(bullets) > 0 else takeaway, 130),
                "complication": clip_text(bullets[1] if len(bullets) > 1 else takeaway, 130),
                "resolution": clip_text(bullets[2] if len(bullets) > 2 else takeaway, 130),
            }
        if intent == "trend_chart":
            page["chart"] = {
                "chart_type": "line",
                "unit": "index",
                "categories": ["T1", "T2", "T3", "T4"],
                "series": [{"name": "Signal", "values": _sample_numbers(idx, 4)}],
            }
        if intent == "comparison":
            page["comparison"] = [
                {"label": "Option A", "points": [clip_text(line, 86) for line in bullets[:3]]},
                {"label": "Option B", "points": [clip_text(line, 86) for line in (bullets[2:5] or bullets[:3])]},
            ]
        if intent == "process":
            page["steps"] = [clip_text(line, 64) for line in bullets[:5]]
        if intent == "kpi_summary":
            page["kpis"] = [
                {"label": "Priority", "value": str(idx), "note": clip_text(bullets[0] if bullets else "", 56)},
                {"label": "Signal", "value": "High", "note": clip_text(bullets[1] if len(bullets) > 1 else "", 56)},
                {"label": "Action", "value": "Next", "note": clip_text(bullets[2] if len(bullets) > 2 else "", 56)},
            ]
        if intent == "table_analysis":
            page["table"] = {
                "headers": ["Theme", "Evidence", "Implication"],
                "rows": [[f"Item {i + 1}", line[:42], "Review"] for i, line in enumerate(bullets[:4])],
            }
        if intent in {"structure_map", "decision_memo"}:
            page["comparison"] = [{"label": f"Area {i + 1}", "points": [clip_text(line, 86)]} for i, line in enumerate(bullets[:4])]
        spec_pages.append(page)
    return {
        "schema_version": JOB_SPEC_SCHEMA_VERSION,
        "language": language,
        "focus": focus,
        "pages": spec_pages,
        "client_note": "Public job spec contains page intent, title, takeaway, content blocks, and source notes only.",
    }


def write_job_spec_prompt(task_dir: Path) -> None:
    prompt = {
        "schema_version": JOB_SPEC_SCHEMA_VERSION,
        "instructions": [
            "Create one page object per slide.",
            "Use only public intent values such as comparison, trend_chart, process, kpi_summary, structure_map, summary_scr, table_analysis, decision_memo.",
            "For each page provide title, takeaway, bullets, optional table, optional chart, optional comparison, optional steps, optional kpis, optional scr, and source.",
            "Keep title under 90 characters, takeaway under 180 characters, bullets under 120 characters, tables at 5 rows by 4 columns, charts at 6 categories by 3 series.",
        ],
        "page_fields": [
            "intent",
            "title",
            "takeaway",
            "bullets",
            "table",
            "chart",
            "comparison",
            "steps",
            "kpis",
            "scr",
            "source",
        ],
    }
    analysis_dir = task_dir / "analysis"
    analysis_dir.mkdir(parents=True, exist_ok=True)
    (analysis_dir / "job_spec_prompt.json").write_text(json.dumps(prompt, ensure_ascii=False, indent=2), encoding="utf-8")


def load_job_spec(path_value: str) -> dict[str, object] | None:
    if not path_value:
        return None
    path = Path(path_value).expanduser()
    if not path.exists():
        raise SystemExit(f"job spec not found: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise SystemExit("job spec must be a JSON object.")
    return data


def cmd_make_deck_server(args: argparse.Namespace) -> int:
    paopao_auth = _load_sibling("paopao_auth")
    if not args.name:
        raise SystemExit("make-deck requires --name.")
    if not args.pages:
        raise SystemExit("make-deck requires --pages.")
    pages = int(args.pages)
    task_dir = Path(args.output_root).resolve() / slugify(args.name)
    delivery_dir = task_dir / "delivery"
    delivery_dir.mkdir(parents=True, exist_ok=True)
    source_text = extract_text_from_source(str(args.source or ""))
    lines = compact_lines(source_text, limit=28)
    analysis_summary = "\n".join(f"- {line}" for line in lines) or str(args.focus or "")
    job_spec = load_job_spec(str(getattr(args, "job_spec", "") or "")) or build_job_spec(
        source_text,
        min(pages, 8),
        str(args.focus or ""),
        str(args.language or ""),
    )
    payload = {
        "client_job_id": f"{task_dir.name}-{int(time.time())}",
        "pages": pages,
        "language": str(args.language or ""),
        "focus": str(args.focus or ""),
        "job_spec": job_spec,
    }
    write_job_spec_prompt(task_dir)
    try:
        job = paopao_auth.submit_deck_job(payload)
        job_id = str(job.get("job_id", ""))
        status = job if str(job.get("status")) == "done" else paopao_auth.fetch_deck_job(job_id)
        if str(status.get("status")) != "done":
            raise SystemExit(json.dumps(status, ensure_ascii=False, indent=2))
        artifacts = [
            name for name in status.get("artifacts", job.get("artifacts", []))
            if isinstance(name, str) and Path(name).name in SERVER_JOB_ARTIFACTS
        ]
        for name in artifacts:
            data = paopao_auth.fetch_deck_job_artifact(job_id, name)
            (delivery_dir / Path(name).name).write_bytes(data)
    except paopao_auth.AuthError as exc:
        raise SystemExit(str(exc)) from exc
    manifest = {
        "task_name": task_dir.name,
        "page_count": min(pages, 8),
        "requested_page_count": pages,
        "language": str(args.language or ""),
        "focus": str(args.focus or ""),
        "status": "delivered",
        "pipeline_mode": "server_job",
        "server_job_id": job_id,
        "delivery_dir": str(delivery_dir),
        "delivery_files": sorted(path.name for path in delivery_dir.iterdir() if path.is_file()),
    }
    task_dir.mkdir(parents=True, exist_ok=True)
    (task_dir / "paopao_task.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "ok": True,
        "task_dir": str(task_dir),
        "delivery_dir": str(delivery_dir),
        "files": manifest["delivery_files"],
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_doctor(_: argparse.Namespace) -> int:
    runtime = PLUGIN_ROOT / "scripts" / _s("deck", "_frame.py")
    fetched: list[str] = []
    error = ""
    if not runtime.exists():
        try:
            names = [*AUTHORIZED_WORKFLOW_FILES, *authorized_prompt_names()]
            fetched = fetch_workflow_bundle(names)
        except SystemExit as exc:
            error = str(exc)
    checks = {
        "plugin_root": str(PLUGIN_ROOT),
        "public_bootstrap": True,
        "runtime_present": runtime.exists(),
        "access_ready": True,
        "fetched": summarize(fetched),
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
    for name in names:
        if name not in destinations:
            raise SystemExit(f"Unknown workflow file: {name}")
    if args.all:
        prompt_names = authorized_prompt_names(full_library=bool(getattr(args, "full_library", False)))
        names = [*names, *prompt_names]
    else:
        prompt_names = []
    written = fetch_workflow_bundle(list(names))
    if prompt_names:
        assert_downloaded_prompts(prompt_names, written)
    print(json.dumps({
        "ok": True,
        "library_mode": "server_authorized",
        "written": summarize(written),
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_update(args: argparse.Namespace) -> int:
    return cmd_fetch_workflow(
        argparse.Namespace(
            all=True,
            name="paopao_run.py",
            full_library=bool(getattr(args, "full_library", False)),
        )
    )


def cmd_runtime_required(args: argparse.Namespace) -> int:
    if getattr(args, "command", "") == "make-deck" and should_use_server_job():
        return cmd_make_deck_server(args)
    cmd_fetch_workflow(argparse.Namespace(all=True, name="paopao_run.py", full_library=True))
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
    update.add_argument("--full-library", action="store_true", help="Also refresh every authorized prompt template")
    update.set_defaults(func=cmd_update)

    fetch = sub.add_parser("fetch-workflow", help="Fetch authorized Paopao runtime files")
    fetch.add_argument("--all", action="store_true")
    fetch.add_argument("--full-library", action="store_true", help="Also fetch every authorized prompt template")
    fetch.add_argument("--name", default="paopao_run.py", choices=sorted(workflow_destinations().keys()))
    fetch.set_defaults(func=cmd_fetch_workflow)

    init = sub.add_parser("init", help=argparse.SUPPRESS)
    init.set_defaults(func=cmd_runtime_required, command="init")

    make_deck = sub.add_parser("make-deck", help=argparse.SUPPRESS)
    make_deck.add_argument("--name", default="")
    make_deck.add_argument("--source", default="")
    make_deck.add_argument("--pages", type=int, default=0)
    make_deck.add_argument("--language", default="")
    make_deck.add_argument("--focus", default="")
    make_deck.add_argument("--output-root", default="output")
    make_deck.add_argument("--pipeline-mode", default="direct_pptx")
    make_deck.add_argument("--job-spec", default="")
    make_deck.set_defaults(func=cmd_runtime_required, command="make-deck")

    for name in [
        "next",
        "check",
        "plan-prompts",
        "finalize-delivery",
        _s("prepare-direct-build-", "pack", "ets"),
        "render-pptx-previews",
    ]:
        command = sub.add_parser(name, help=argparse.SUPPRESS)
        command.set_defaults(func=cmd_runtime_required, command=name)
    return parser


def main() -> int:
    args, _unknown = build_parser().parse_known_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())

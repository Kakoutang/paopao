#!/usr/bin/env python3
"""Token-audit helpers for Paopao lab diagnostics."""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path


TOKEN_AUDIT_STAGES = [
    "setup_update",
    "analysis_prompting",
    "image2_generation",
    "visual_measurement",
    "direct_reconstruction",
    "html_reconstruction",
    "qa_review",
    "release_deploy",
    "other",
]


def _empty_usage() -> dict[str, int]:
    return {key: 0 for key in [
        "input_tokens",
        "cached_input_tokens",
        "output_tokens",
        "reasoning_output_tokens",
        "total_tokens",
    ]}


def _add_usage(target: dict[str, int], usage: object) -> None:
    if not isinstance(usage, dict):
        return
    for key in target:
        value = usage.get(key)
        if isinstance(value, (int, float)):
            target[key] += int(value)


def _codex_sessions(session_arg: str = "", recent: int = 1) -> list[Path]:
    raw = str(session_arg or "").strip()
    root = Path.home() / ".codex" / "sessions"
    candidates = sorted(
        (p for p in root.rglob("*.jsonl") if p.is_file()) if root.exists() else [],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if raw:
        path = Path(raw).expanduser()
        if path.exists():
            return [path.resolve()]
        return [p for p in candidates if raw in p.name or raw in str(p)][: max(1, recent)]
    return candidates[: max(1, recent)]


def _extract_int(pattern: str, text: str) -> int | None:
    match = re.search(pattern, text)
    return int(match.group(1)) if match else None


def _stage_from_text(text: str) -> str:
    lower = text.lower()
    groups = [
        ("release_deploy", [
            "git push", "render deploy", "check_public_release", "raw.githubusercontent.com",
            "deploy hook", "public repo", "internal repo", "public release",
        ]),
        ("qa_review", [
            "fidelity_review", "powerpoint_review", "check --stage pptx", "check --stage pipeline",
            "check --stage delivery", "pptx_actual", "soffice", "libreoffice", "pdftoppm",
            "compare_", "qa/", "similarity", "visual check",
        ]),
        ("direct_reconstruction", ["build-powerpoint-layout-plan", "direct_pptx", "native_chart", "native_table"]),
        ("html_reconstruction", [
            " render ", "renderer.py", "html_renderer", "html_reference", "render_manifest",
            "playwright", "html slide", "html path",
        ]),
        ("visual_measurement", [
            "visual_inventory", "visual_contract", "visual_blueprint", "record-image2-observation",
            "extract-image2-contract", "measurement", "bbox", "component_parts",
        ]),
        ("image2_generation", [
            "image2", "image_generation", "imagegen", "prepare-image2", "register-image2",
            "extract-codex-imagegen-result", "generation_request",
        ]),
        ("analysis_prompting", [
            "analysis_report", "slide_story", "plan-prompts", "fill-prompt-template", "final_prompt",
            "prompt_selection", "source analysis",
        ]),
        ("setup_update", [
            "doctor", "fetch-workflow", "paopao_run.py update", "make-deck", "run-task",
            "read skill", "skill.md", "git status", "git log", "rg --files",
        ]),
    ]
    for stage, needles in groups:
        if any(token in lower for token in needles):
            return stage
    return "other"


def _merge_stage(existing: str, incoming: str) -> str:
    if existing == "other":
        return incoming
    if incoming == "other" or existing == incoming:
        return existing
    priority = {
        "release_deploy": 90,
        "qa_review": 80,
        "direct_reconstruction": 70,
        "html_reconstruction": 70,
        "visual_measurement": 60,
        "image2_generation": 50,
        "analysis_prompting": 40,
        "setup_update": 30,
        "other": 0,
    }
    return incoming if priority.get(incoming, 0) > priority.get(existing, 0) else existing


def _read_session(session_path: Path) -> dict[str, object]:
    turns: dict[str, dict[str, object]] = {}
    current_turn = "unknown"
    session_meta: dict[str, object] = {"path": str(session_path)}
    command_counts: dict[str, int] = {}
    large_outputs: list[dict[str, object]] = []
    failed_commands: list[dict[str, object]] = []
    path_profile_hits = {"direct_pptx": 0, "html": 0, "render_manifest_blockers": 0}
    token_events = 0

    def turn(turn_id: str) -> dict[str, object]:
        return turns.setdefault(turn_id, {
            "turn_id": turn_id,
            "stage": "other",
            "usage": _empty_usage(),
            "snippets": [],
            "commands": [],
            "tool_output_tokens": 0,
            "failed_command_count": 0,
        })

    with session_path.open("r", encoding="utf-8", errors="ignore") as handle:
        for line_no, line in enumerate(handle, 1):
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            payload = record.get("payload")
            if record.get("type") == "session_meta" and isinstance(payload, dict):
                session_meta.update({
                    "id": payload.get("id"),
                    "cwd": payload.get("cwd"),
                    "created_at": payload.get("timestamp"),
                    "originator": payload.get("originator"),
                })
                continue
            if not isinstance(payload, dict):
                continue
            if record.get("type") == "event_msg":
                event_type = payload.get("type")
                if event_type == "task_started" and payload.get("turn_id"):
                    current_turn = str(payload.get("turn_id"))
                    turn(current_turn)
                if event_type == "token_count":
                    info = payload.get("info")
                    usage = info.get("last_token_usage") if isinstance(info, dict) else None
                    if isinstance(usage, dict):
                        token_events += 1
                        _add_usage(turn(current_turn)["usage"], usage)  # type: ignore[arg-type]
                    continue
            if record.get("type") != "response_item":
                continue
            metadata = payload.get("metadata")
            if isinstance(metadata, dict) and metadata.get("turn_id"):
                current_turn = str(metadata.get("turn_id"))
            item = turn(current_turn)
            snippets = item["snippets"]  # type: ignore[assignment]
            ptype = str(payload.get("type", ""))
            role = str(payload.get("role", ""))
            if ptype == "function_call":
                name = str(payload.get("name", ""))
                args = str(payload.get("arguments", ""))
                command_key = name
                try:
                    parsed = json.loads(args)
                    if isinstance(parsed, dict) and parsed.get("cmd"):
                        command_key = str(parsed.get("cmd", "")).strip().split("\n", 1)[0][:180]
                except Exception:
                    pass
                command_counts[command_key] = command_counts.get(command_key, 0) + 1
                item["commands"].append(command_key)  # type: ignore[index]
                snippets.append(f"{name} {args[:600]}")
                item["stage"] = _merge_stage(str(item["stage"]), _stage_from_text(f"{name} {args}"))
            elif ptype == "function_call_output":
                output = str(payload.get("output", ""))
                original_tokens = _extract_int(r"Original token count:\s*(\d+)", output)
                if original_tokens:
                    item["tool_output_tokens"] = int(item["tool_output_tokens"]) + original_tokens
                    if original_tokens >= 1500:
                        large_outputs.append({
                            "line": line_no,
                            "turn_id": current_turn,
                            "tokens": original_tokens,
                            "stage": item["stage"],
                            "preview": output[:220].replace("\n", "\\n"),
                        })
                exit_code = _extract_int(r"Process exited with code\s+(-?\d+)", output)
                if exit_code not in (None, 0):
                    item["failed_command_count"] = int(item["failed_command_count"]) + 1
                    failed_commands.append({
                        "line": line_no,
                        "turn_id": current_turn,
                        "exit_code": exit_code,
                        "stage": item["stage"],
                        "preview": output[:260].replace("\n", "\\n"),
                    })
                if "render_manifest" in output and "not_required_for_direct_pptx" not in output:
                    path_profile_hits["render_manifest_blockers"] += 1
                if "direct_pptx" in output:
                    path_profile_hits["direct_pptx"] += 1
                if "html_reference_fidelity" in output or "render_manifest" in output:
                    path_profile_hits["html"] += 1
                snippets.append(output[:500])
                item["stage"] = _merge_stage(str(item["stage"]), _stage_from_text(output))
            elif ptype == "message":
                parts = payload.get("content")
                if isinstance(parts, list):
                    text = "\n".join(str(part.get("text", "")) for part in parts if isinstance(part, dict))
                else:
                    text = str(parts or "")
                if role in {"user", "assistant"} or text:
                    snippets.append(text[:700])
                    item["stage"] = _merge_stage(str(item["stage"]), _stage_from_text(text))

    totals = _empty_usage()
    by_stage = {stage: _empty_usage() for stage in TOKEN_AUDIT_STAGES}
    stage_turns = {stage: 0 for stage in TOKEN_AUDIT_STAGES}
    for data in turns.values():
        stage = str(data.get("stage") or "other")
        if stage not in by_stage:
            stage = "other"
        usage = data.get("usage")
        if isinstance(usage, dict):
            _add_usage(totals, usage)
            _add_usage(by_stage[stage], usage)
        stage_turns[stage] += 1
    repeated = [
        {"command": command, "count": count}
        for command, count in sorted(command_counts.items(), key=lambda item: (-item[1], item[0]))
        if count >= 3
    ][:20]
    return {
        "session": session_meta,
        "ok": token_events > 0,
        "token_events": token_events,
        "total_usage": totals,
        "by_stage": by_stage,
        "stage_turns": stage_turns,
        "turn_count": len(turns),
        "large_tool_outputs": large_outputs[:30],
        "failed_commands": failed_commands[:30],
        "repeated_commands": repeated,
        "path_profile_hits": path_profile_hits,
    }


def _recommendations(report: dict[str, object]) -> list[str]:
    recs: list[str] = []
    by_stage = report.get("by_stage") if isinstance(report.get("by_stage"), dict) else {}
    ranked = sorted(
        ((stage, data.get("total_tokens", 0)) for stage, data in by_stage.items() if isinstance(data, dict)),
        key=lambda item: int(item[1]),
        reverse=True,
    )
    if ranked and ranked[0][1]:
        recs.append(f"largest_stage={ranked[0][0]} uses {ranked[0][1]} tokens; optimize this stage first")
    if report.get("large_tool_outputs"):
        recs.append("large_tool_outputs detected; cap command max_output_tokens and avoid dumping full SKILL/spec files after first read")
    if report.get("repeated_commands"):
        recs.append("repeated_commands detected; cache stage check results and batch rebuild/check commands")
    if report.get("failed_commands"):
        recs.append("failed_commands detected; move these blockers earlier into the controller before expensive generation")
    hits = report.get("path_profile_hits")
    if isinstance(hits, dict) and int(hits.get("render_manifest_blockers", 0) or 0) > 0:
        recs.append("render_manifest blockers appeared; keep production on the html renderer path and refresh stale PPTX renders earlier")
    return recs or ["no obvious waste pattern found; compare against another session for A/B insight"]


def _write_markdown(report: dict[str, object], output: Path) -> None:
    session = report.get("session") if isinstance(report.get("session"), dict) else {}
    total = report.get("total_usage") if isinstance(report.get("total_usage"), dict) else {}
    by_stage = report.get("by_stage") if isinstance(report.get("by_stage"), dict) else {}
    stage_turns = report.get("stage_turns") if isinstance(report.get("stage_turns"), dict) else {}
    lines = [
        "# Paopao Token Audit", "",
        "## Session",
        f"- id: {session.get('id', '')}",
        f"- path: {session.get('path', '')}",
        f"- cwd: {session.get('cwd', '')}", "",
        "## Total",
        f"- input_tokens: {total.get('input_tokens', 0)}",
        f"- cached_input_tokens: {total.get('cached_input_tokens', 0)}",
        f"- output_tokens: {total.get('output_tokens', 0)}",
        f"- reasoning_output_tokens: {total.get('reasoning_output_tokens', 0)}",
        f"- total_tokens: {total.get('total_tokens', 0)}", "",
        "## By Stage",
        "| Stage | Turns | Input | Cached input | Output | Reasoning | Total |",
        "|---|---:|---:|---:|---:|---:|---:|",
    ]
    for stage in TOKEN_AUDIT_STAGES:
        data = by_stage.get(stage, {}) if isinstance(by_stage, dict) else {}
        if not isinstance(data, dict):
            data = {}
        lines.append(
            f"| {stage} | {stage_turns.get(stage, 0)} | {data.get('input_tokens', 0)} | "
            f"{data.get('cached_input_tokens', 0)} | {data.get('output_tokens', 0)} | "
            f"{data.get('reasoning_output_tokens', 0)} | {data.get('total_tokens', 0)} |"
        )
    large_outputs = report.get("large_tool_outputs") if isinstance(report.get("large_tool_outputs"), list) else []
    failed_commands = report.get("failed_commands") if isinstance(report.get("failed_commands"), list) else []
    repeated_commands = report.get("repeated_commands") if isinstance(report.get("repeated_commands"), list) else []
    lines.extend([
        "", "## Waste Signals",
        f"- large_tool_outputs: {len(large_outputs)}",
        f"- failed_commands: {len(failed_commands)}",
        f"- repeated_commands: {len(repeated_commands)}",
        "", "## Recommendations",
    ])
    lines.extend(f"- {rec}" for rec in _recommendations(report))
    if repeated_commands:
        lines.extend(["", "## Top Repeated Commands"])
        lines.extend(
            f"- {item.get('count')}x `{item.get('command')}`"
            for item in repeated_commands[:10]
            if isinstance(item, dict)
        )
    if large_outputs:
        lines.extend(["", "## Large Tool Outputs"])
        lines.extend(
            f"- {item.get('tokens')} tokens, stage={item.get('stage')}, line={item.get('line')}"
            for item in large_outputs[:10]
            if isinstance(item, dict)
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")


def cmd_token_audit(args: object) -> int:
    sessions = _codex_sessions(str(getattr(args, "session", "")), int(getattr(args, "recent", 1)))
    if not sessions:
        print(json.dumps({
            "ok": False,
            "issue": "no Codex session jsonl files found",
            "hint": "Pass --session with a rollout jsonl path or session id.",
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    reports = [_read_session(path) for path in sessions if path.exists()]
    if not reports:
        print(json.dumps({
            "ok": False,
            "issue": "session path did not exist",
            "sessions": [str(path) for path in sessions],
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    combined = {
        "schema": "paopao.token_audit.v1",
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "ok": any(report.get("ok") for report in reports),
        "session_count": len(reports),
        "reports": reports,
    }
    json_path = str(getattr(args, "json", "") or "")
    output_path = str(getattr(args, "output", "") or "")
    if json_path:
        out = Path(json_path).expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(combined, ensure_ascii=False, indent=2), encoding="utf-8")
    if output_path:
        out = Path(output_path).expanduser().resolve()
        if len(reports) == 1:
            _write_markdown(reports[0], out)
        else:
            out.parent.mkdir(parents=True, exist_ok=True)
            lines = ["# Paopao Token Audit", ""]
            for idx, report in enumerate(reports, 1):
                session = report.get("session") if isinstance(report.get("session"), dict) else {}
                total = report.get("total_usage") if isinstance(report.get("total_usage"), dict) else {}
                lines.append(f"## Session {idx}: {session.get('id', session.get('path', ''))}")
                lines.append(f"- total_tokens: {total.get('total_tokens', 0)}")
                lines.extend(f"- {rec}" for rec in _recommendations(report))
                lines.append("")
            out.write_text("\n".join(lines), encoding="utf-8")
    summary = {
        "ok": combined["ok"],
        "session_count": len(reports),
        "sessions": [
            {
                "id": (r.get("session") or {}).get("id") if isinstance(r.get("session"), dict) else None,
                "path": (r.get("session") or {}).get("path") if isinstance(r.get("session"), dict) else None,
                "total_tokens": (r.get("total_usage") or {}).get("total_tokens") if isinstance(r.get("total_usage"), dict) else 0,
                "recommendations": _recommendations(r),
            }
            for r in reports
        ],
        "json": str(Path(json_path).expanduser().resolve()) if json_path else "",
        "output": str(Path(output_path).expanduser().resolve()) if output_path else "",
    }
    if getattr(args, "fail_on_waste", False):
        failures: list[str] = []
        for report in reports:
            session = report.get("session") if isinstance(report.get("session"), dict) else {}
            sid = str(session.get("id") or session.get("path") or "unknown")
            if report.get("large_tool_outputs"):
                failures.append(f"{sid}: large tool outputs")
            if report.get("failed_commands"):
                failures.append(f"{sid}: failed commands")
            if report.get("repeated_commands"):
                failures.append(f"{sid}: repeated commands")
        if failures:
            summary["ok"] = False
            summary["waste_failures"] = failures
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 2
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if combined["ok"] else 1

#!/usr/bin/env python3
"""Local helpers for paopao tasks.

This script does not call an LLM. Codex performs the reasoning workflow from
the skill instructions; this helper creates stable task folders, validates the
Image2-to-editable-PPTX commercial contract, and renders HTML when the declared
commercial path uses the HTML renderer.
"""

from __future__ import annotations

import argparse
import contextlib
import hashlib
import io
import importlib.util
import json
import os
import random
import re
import shutil
import struct
import subprocess
import sys
import time
from pathlib import Path

import paopao_auth


PLUGIN_ROOT = Path(__file__).resolve().parents[1]
RENDERER = PLUGIN_ROOT / "scripts" / "renderer.py"
SYSTEM_PROMPT = PLUGIN_ROOT / "prompts" / "SYSTEM_PROMPT.md"
PROMPT_LIBRARY_DIR = PLUGIN_ROOT / "prompts"


PROMPT_REQUIRED_MARKERS = ["TITLE", "BOTTOM", "Source", "DESIGN"]
PROMPT_TEMPLATE_RE = re.compile(r"^\s*PROMPT_TEMPLATE\s*:\s*(?P<name>[A-Za-z0-9._-]+\.md)\s*$", re.MULTILINE)
LAYOUT_NAME_RE = re.compile(r"^\s*LAYOUT_NAME\s*:\s*(?P<name>[A-Za-z0-9._-]+)\s*$", re.MULTILINE)
PROMPT_FORBIDDEN_PATTERNS = [
    "Create a McKinsey-style consulting slide.",
]
PLACEHOLDER_PATTERNS = [
    "add bullets here",
    "use relevant data",
    "tbd",
    "todo",
    "placeholder",
    "占位",
    "待补充",
    "相关数据",
]

PROMPT_INTERNAL_PATTERNS = [
    "analysis/final_prompt_*.md",
    "analysis/prompt_selection_audit.md",
    "**/*prompt*.md",
    "**/*Prompt*.md",
]
PROMPT_PRIVATE_DIR = Path("qa") / "private_prompts"
PROMPT_ARCHIVE_ENV = "PAOPAO_KEEP_PRIVATE_PROMPTS"
TASK_LOCAL_GITIGNORE = """# Paopao build workspace: internal by default.
# The user-facing artifacts are published under delivery/ after final QA.
*
!.gitignore
!delivery/
!delivery/**
"""
DELIVERY_TEMP_PATTERNS = [
    "~$*.pptx",
    "~$*.ppt",
    "*.tmp",
]
DELIVERY_ALLOWED_SUFFIXES = {
    ".pptx",
    ".html",
    ".png",
    ".jpg",
    ".jpeg",
    ".webp",
    ".gif",
    ".css",
}
DELIVERY_FORBIDDEN_NAME_PARTS = {
    "prompt",
    "analysis",
    "spec",
    "qa",
    "final_prompt",
    "image2_prompt",
    "generation_request",
    "manifest",
    "audit",
    "contract",
}

POWERPOINT_REVIEW_MIN_CHECKS = [
    "actual_powerpoint",
    "text_visible",
    "numbers_visible",
    "icons_visible",
    "layout_match",
    "no_overlap",
]

FIDELITY_REVIEW_MIN_CHECKS = [
    "nav",
    "title",
    "module_geometry",
    "icons",
    "takeaway",
    "color_hierarchy",
]
MIN_FIDELITY_IMAGE_SCORE = float(os.getenv("PAOPAO_MIN_FIDELITY_IMAGE_SCORE", "0.82"))
MIN_HTML_REFERENCE_IMAGE_SCORE = float(
    os.getenv("PAOPAO_MIN_HTML_REFERENCE_IMAGE_SCORE", str(MIN_FIDELITY_IMAGE_SCORE))
)
COMMERCIAL_SIMILARITY_MIN = float(os.getenv("PAOPAO_COMMERCIAL_SIMILARITY_MIN", "0.95"))
REGION_GEOMETRY_CENTER_TOLERANCE_PX = float(os.getenv("PAOPAO_REGION_GEOMETRY_CENTER_TOLERANCE_PX", "64"))
REGION_GEOMETRY_SIZE_TOLERANCE_PX = float(os.getenv("PAOPAO_REGION_GEOMETRY_SIZE_TOLERANCE_PX", "80"))
REGION_GEOMETRY_SIZE_TOLERANCE_RATIO = float(os.getenv("PAOPAO_REGION_GEOMETRY_SIZE_TOLERANCE_RATIO", "0.25"))
IMAGE2_STYLE_REVIEW_MIN_CHECKS = [
    "aspect_ratio_16_9",
    "house_style_reference",
    "palette_discipline",
    "clean_background",
    "linework_and_borders",
    "material_simplicity",
    "title_weight",
    "module_density",
    "takeaway",
    "color_hierarchy",
    "icons",
]
IMAGE2_USER_REVIEW_SCHEMA = "paopao.image2_user_review.v1"
PIPELINE_PASS_SCHEMA = "paopao.pipeline_pass.v1"
FINAL_DELIVERY_PASS_SCHEMA = "paopao.final_delivery_pass.v1"
COMMERCIAL_RENDER_CONTRACT_SCHEMA = "paopao.commercial_render_contract.v1"
COMMERCIAL_RENDER_PATHS = {"html", "direct_pptx"}

VISUAL_CONTRACT_REQUIRED_REGIONS = [
    "nav",
    "title",
    "content",
    "takeaway",
    "source",
]
VISUAL_CONTRACT_ROOT_REGION_IDS = set(VISUAL_CONTRACT_REQUIRED_REGIONS)
MIN_VISUAL_CONTRACT_REGIONS = int(os.getenv("PAOPAO_MIN_VISUAL_CONTRACT_REGIONS", "9"))
MIN_VISUAL_CONTRACT_DETAIL_REGIONS = int(os.getenv("PAOPAO_MIN_VISUAL_CONTRACT_DETAIL_REGIONS", "4"))
VISUAL_CONTRACT_BORDER_STYLES = {
    "none",
    "solid",
    "dashed",
    "dotted",
    "mixed",
}
IMAGE_ONLY_RECONSTRUCTION_SOURCE = "image2_reference_only"
POST_IMAGE_FORBIDDEN_MEMORY_MARKERS = [
    "analysis_report",
    "final_prompt",
    "prompt_selection",
    "prompt_selection_audit",
    "prompt_template",
    "layout_name",
    "system_prompt",
    "image2_prompt",
    "prompt-library",
    "prompt library",
    "selected prompt",
    "template-backed",
    "filled prompt",
]

IMAGE2_LOCKED_TAIL = """\

IMAGE2 GENERATION LOCK:
- Use the full slide prompt above; do not summarize, rewrite, compress, or substitute it.
- Canvas must be landscape 16:9, target 1920x1080. Reject portrait, square, 4:3, or 3:2 outputs.
- Preserve the locked Paopao visual language exactly: white page background, black bold title, deep-blue accents, thin blue/grey linework, compact editable-looking consulting materials, a slim top directory bar, and a slim bottom takeaway strip.
- Keep white as the dominant surface. Use pale blue only as a local highlight, table header tint, small callout fill, or clearly needed grouping aid; do not wash large content bands, whole modules, or the slide canvas in pale blue unless the selected template explicitly requires a filled analytical object.
- Icons are allowed but not automatic. Add an icon only when the selected slide prompt explicitly asks for an icon or when a single semantic marker materially improves comprehension. Do not add large line icons to every band, card, bullet, takeaway, or navigation item.
- The top navigation should look like a thin consulting directory bar, not equal-width filled web tabs. Use compact text labels, subtle separators, and a right-side page number when useful.
- The bottom takeaway should be a thin text strip, not a large illustrated banner. Do not add a lightbulb, target, arrow, or other icon to the takeaway unless the slide prompt explicitly requires it.
- Preserve the selected prompt-template layout diversity. Do not force SCR, matrix, sidebar, or decision-rail composition unless the selected prompt template calls for it.
- Do not output decorative backgrounds, heavy gradients, thick random frames, ornamental boxes, oversized icon art, or busy linework that is not present in the selected layout.
- The output must be a rebuildable PPT reference, not a poster, photo, illustration, hero image, or decorative mockup.
- Avoid extra borders, dashed boxes, guide lines, decorative frames, gradients, or complex imagery unless explicitly required by the slide prompt.
"""

PROMPT_SCAFFOLD_FAMILY_BY_PREFIX = {
    "01": "two-column",
    "02": "t-shape",
    "03": "inverted-t",
    "04": "two-row",
    "05": "three-column",
    "06": "four-quadrant",
    "07": "row-stack",
    "08": "full-table",
    "09": "process-flow",
    "10": "timeline",
    "11": "waterfall",
    "12": "pyramid-funnel",
    "13": "center-radial",
    "14": "dashboard",
    "15": "kpi-hero",
    "16": "large-grid",
    "17": "divider-cover",
    "18": "swimlane",
}
PROMPT_VISUAL_GRAMMAR_BY_FAMILY = {
    "two-column": "split-panel",
    "t-shape": "split-panel",
    "inverted-t": "split-panel",
    "two-row": "stacked-bands",
    "three-column": "column-cards",
    "four-quadrant": "quadrant",
    "row-stack": "stacked-bands",
    "full-table": "matrix-table",
    "process-flow": "flow",
    "timeline": "flow",
    "waterfall": "bridge-waterfall",
    "pyramid-funnel": "hierarchy",
    "center-radial": "radial-system",
    "dashboard": "metric-dashboard",
    "kpi-hero": "metric-dashboard",
    "large-grid": "matrix-table",
    "divider-cover": "section-break",
    "swimlane": "flow",
}
PROMPT_ROLE_FAMILY_PREFERENCES = [
    (
        re.compile(r"市场|空间|规模|渗透|market|space|sizing|规模", re.IGNORECASE),
        {"dashboard": 34, "kpi-hero": 26, "pyramid-funnel": 18, "two-column": 12, "waterfall": -14},
    ),
    (
        re.compile(r"增长|驱动|需求|driver|growth|demand|acceleration", re.IGNORECASE),
        {"center-radial": 36, "process-flow": 26, "four-quadrant": 20, "row-stack": 14, "waterfall": -18},
    ),
    (
        re.compile(r"细分|赛道|品类|对比|segment|category|compare|matrix", re.IGNORECASE),
        {"full-table": 38, "large-grid": 28, "three-column": 18, "dashboard": 10},
    ),
    (
        re.compile(r"机会落点|产业链|供应链|链路|主体|value chain|supply chain|winners?|map", re.IGNORECASE),
        {"process-flow": 34, "swimlane": 28, "center-radial": 20, "dashboard": 12},
    ),
    (
        re.compile(r"风险|约束|risk|priority|prioriti[sz]ation|评估|score", re.IGNORECASE),
        {"four-quadrant": 30, "full-table": 22, "large-grid": 18},
    ),
]
PROMPT_TEMPLATE_KEYWORD_BONUSES = [
    (re.compile(r"tam|sam|som|市场|空间|规模|market", re.IGNORECASE), "12D_market_tam_sam_som.md", 30),
    (re.compile(r"驱动|driver|增长|growth|需求", re.IGNORECASE), "13D_hub_spoke_ecosystem.md", 24),
    (re.compile(r"驱动|driver|增长|growth|需求", re.IGNORECASE), "02D_flow_diagram_with_detail_panels.md", 18),
    (re.compile(r"细分|赛道|品类|segment|category", re.IGNORECASE), "08F_trend_matrix_grid.md", 28),
    (re.compile(r"细分|赛道|品类|segment|category", re.IGNORECASE), "16A_segment_attribute_matrix.md", 22),
    (re.compile(r"产业链|供应链|链路|value chain|supply chain", re.IGNORECASE), "09C_value_chain_decomposition.md", 28),
    (re.compile(r"机会落点|主体|winners?|stakeholder", re.IGNORECASE), "13E_stakeholder_map.md", 18),
]

PROMPT_SELECTION_PLAN_SCHEMA = "paopao.prompt_selection_plan.v1"
PROMPT_SELECTION_PLAN_PATH = Path("analysis") / "prompt_selection_plan.json"
PROMPT_DATA_TIER1_LABELS = {
    "chain_stages",
    "two_axis_entities",
    "funnel_stages",
    "metric_tree",
    "generation_phases",
    "case_study",
    "scored_alternatives",
    "demographic_data",
    "process_stages",
    "scoring_dimensions",
    "kpi_metrics",
}
PROMPT_DATA_SIGNAL_PATTERNS: dict[str, re.Pattern[str]] = {
    "chain_stages": re.compile(
        r"上游|中游|下游|upstream|midstream|downstream|value\s+chain|供应链|产业链|价值链",
        re.IGNORECASE,
    ),
    "two_axis_entities": re.compile(
        r"象限|quadrant|scatter|positioning\s+map|2\s*[x×]\s*2|two\s+axes|两个维度",
        re.IGNORECASE,
    ),
    "funnel_stages": re.compile(r"漏斗|转化率|drop.?off|conversion|customer\s+journey|客户旅程", re.IGNORECASE),
    "metric_tree": re.compile(r"指标分解|driver\s+tree|decomposition|拆解|分解|root\s+cause", re.IGNORECASE),
    "generation_phases": re.compile(r"第[一二三四]代|generation|代际|技术演进|phase\s+[1-4ivx]", re.IGNORECASE),
    "case_study": re.compile(r"案例|case\s+study|发展历程|company\s+history|milestone", re.IGNORECASE),
    "scored_alternatives": re.compile(r"评分|打分|scoring|weighted\s+criteria|综合评估|评价标准", re.IGNORECASE),
    "demographic_data": re.compile(r"人口|age\s+distribution|gender|年龄|性别|population", re.IGNORECASE),
    "named_competitors": re.compile(r"竞争|market\s+share|市场份额|品牌|competitor|players?|CR[358]|集中度", re.IGNORECASE),
    "segment_profiles": re.compile(r"细分|segment|persona|用户画像|客户群|品类|需求结构", re.IGNORECASE),
    "geographic_markets": re.compile(r"国家|地区|城市|省|regional|geographic|countries|markets", re.IGNORECASE),
    "initiative_portfolio": re.compile(r"举措|initiative|行动|roadmap|workstream|战略路径", re.IGNORECASE),
    "pain_points": re.compile(r"痛点|challenge|risk|风险|问题|瓶颈|headwind", re.IGNORECASE),
    "time_series": re.compile(r"\d{4}[-—~至]\d{4}|CAGR|同比|YoY|forecast|预测|trend|趋势", re.IGNORECASE),
    "financial_dual_metric": re.compile(r"收入|利润|营收|毛利|margin|EBIT|revenue|量价", re.IGNORECASE),
    "scenario_forecast": re.compile(r"情景|scenario|base\s+case|optimistic|pessimistic|乐观|悲观", re.IGNORECASE),
    "kpi_actuals": re.compile(r"KPI|关键绩效|scorecard|target|actual|达成率", re.IGNORECASE),
    "kpi_metrics": re.compile(r"KPI|关键指标|metric|dashboard|核心指标|scorecard", re.IGNORECASE),
    "cost_breakdown": re.compile(r"成本|cost|margin|费用|利润率", re.IGNORECASE),
    "event_annotations": re.compile(r"政策|regulation|事件|milestone|里程碑|监管", re.IGNORECASE),
    "named_drivers": re.compile(r"驱动|driver|growth\s+driver|因素|pull\s+factor|push\s+factor", re.IGNORECASE),
    "process_stages": re.compile(r"流程|步骤|process|stage|journey|阶段|路径", re.IGNORECASE),
    "scoring_dimensions": re.compile(r"维度|criteria|dimension|评估项|评分标准", re.IGNORECASE),
}


IMAGE2_MANIFEST_SCHEMA = "paopao.image2_generation_manifest.v2"
IMAGE2_GENERATION_REQUEST_SCHEMA = "paopao.image2_generation_request.v1"
IMAGE2_VERIFICATION_METHOD = "prompt_sha_attestation"
IMAGE2_PROMPT_TRANSFER_METHOD = "exact_locked_prompt_file_text"
IMAGE2_TARGET_ASPECT = 16 / 9
IMAGE2_ASPECT_TOLERANCE = 0.025
IMAGE2_ALLOWED_SOURCE_KINDS = {
    "image_gen_builtin",
    "image_gen_cli",
}
IMAGE2_FORBIDDEN_SOURCE_PARTS = {
    "qa",
    "pptx_actual",
    "html",
    "pptx",
    "delivery",
}
IMAGE2_FORBIDDEN_SOURCE_NAME_PARTS = {
    "html_preview",
    "html_previews",
    "final_preview",
    "final_previews",
    "contact_sheet",
}

IMAGE2_OBSERVATION_SCHEMA = "paopao.image2_observation.v1"
VISUAL_CONTRACT_SCHEMA = "paopao.visual_contract.v1"
VISUAL_MEASUREMENT_SCHEMA = "paopao.visual_measurement.v1"
POST_IMAGE_DERIVATION_METHOD = "fresh_visual_observation_record"
POST_IMAGE_MEMORY_BOUNDARY_SCHEMA = "paopao.post_image_memory_boundary.v1"

ICON_PLACEHOLDER_WORDS = {
    "AI", "API", "AR", "BANK", "CASE", "CN", "CO2", "CORE", "DB", "DOC",
    "ECO", "EXT", "FIND", "IP", "KPI", "MAP", "METH", "NAV", "NET", "OK",
    "PROD", "SCR", "STD", "SYN", "SYS", "TOOL", "UP", "USER",
}
ICON_CLASS_EXACT_HINTS = {
    "glyph",
    "ibox",
    "icon",
    "iconbox",
    "icon-box",
    "icon_box",
    "ricon",
    "symbol",
}
ICON_CLASS_EDGE_HINTS = (
    "icon-",
    "icon_",
    "-icon",
    "_icon",
    "glyph-",
    "glyph_",
    "-glyph",
    "_glyph",
)


def auth_should_run() -> bool:
    if os.getenv("PAOPAO_LOCAL_DEV") == "1":
        return True
    if os.getenv("PAOPAO_AUTH_REQUIRED") == "1":
        return True
    if open_preview_enabled() and not has_local_license():
        return False
    return (
        bool(os.getenv("PAOPAO_AUTH_URL"))
        or paopao_auth.LICENSE_PATH.exists()
    )


def open_preview_enabled() -> bool:
    return os.getenv("PAOPAO_OPEN_PREVIEW", "1") != "0"


def free_max_slides() -> int:
    raw = os.getenv("PAOPAO_FREE_MAX_SLIDES", "10").strip()
    try:
        return 10 if int(raw) <= 0 else min(10, int(raw))
    except ValueError:
        return 10


def has_local_license() -> bool:
    try:
        data = paopao_auth.read_license()
    except Exception:
        return False
    return bool(data.get("token") and data.get("server_url"))


def reserve_quota(task_dir: Path, pages: int) -> str:
    if os.getenv("PAOPAO_LOCAL_DEV") == "1":
        result = paopao_auth.reserve(job_id=f"{task_dir.name}-{int(time.time())}", pages=pages)
        return str(result.get("reservation_id", ""))
    free_limit = free_max_slides()
    if free_limit and pages > free_limit:
        raise SystemExit(
            f"paopao 免费版最多支持 {free_limit} 页。"
            "如需更多页数或完整模板库，请联系微信 sugarong_ 获取。\n"
            f"paopao public edition supports up to {free_limit} slides. "
            "For more pages or the full template library, contact WeChat: sugarong_"
        )
    if not has_local_license():
        if open_preview_enabled() and free_limit == 0:
            return ""
        if free_limit and pages <= free_limit:
            return ""
        if auth_should_run() or free_limit:
            if free_limit == 0:
                raise SystemExit(
                    "Paopao requires an active license. "
                    "Activate with scripts/paopao_auth.py activate before rendering."
                )
            raise SystemExit(
                f"Paopao free mode supports up to {free_limit} slides. "
                "Activate a license with scripts/paopao_auth.py activate to render larger decks."
            )
    if not auth_should_run():
        return ""
    job_id = f"{task_dir.name}-{int(time.time())}"
    result = paopao_auth.reserve(job_id=job_id, pages=pages)
    return str(result.get("reservation_id", ""))


def finish_quota(reservation_id: str, succeeded: bool) -> None:
    if not reservation_id:
        return
    command = "commit" if succeeded else "cancel"
    paopao_auth.finish_reservation(command, reservation_id)


def slugify(name: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9._-]+", "-", name.strip()).strip("-").lower()
    return slug or "paopao-task"


def write_task_local_gitignore(task_dir: Path) -> None:
    """Hide internal task artifacts from Codex/Git change surfaces."""
    path = task_dir / ".gitignore"
    current = path.read_text(encoding="utf-8") if path.exists() else ""
    if current == TASK_LOCAL_GITIGNORE:
        return
    path.write_text(TASK_LOCAL_GITIGNORE, encoding="utf-8")


def enforce_public_page_limit(pages: int | None) -> None:
    if pages and free_max_slides() and pages > free_max_slides():
        raise SystemExit(
            f"paopao 免费版最多支持 {free_max_slides()} 页。"
            "如需更多页数或完整模板库，请联系微信 sugarong_ 获取。\n"
            f"paopao free tier supports up to {free_max_slides()} slides. "
            "For more pages or the full template library, contact WeChat: sugarong_"
        )


def create_task_dir(
    *,
    name: str,
    output_root: str | Path,
    pages: int | None,
    language: str,
    focus: str,
) -> Path:
    enforce_public_page_limit(pages)
    task_name = slugify(name)
    root = Path(output_root).resolve() / task_name
    for child in [
        "source",
        "analysis",
        "image2",
        "spec",
        "html/assets",
        "pptx",
        "qa/pptx_actual",
    ]:
        (root / child).mkdir(parents=True, exist_ok=True)

    manifest = {
        "task_name": task_name,
        "page_count": pages,
        "language": language,
        "focus": focus,
        "status": "initialized",
    }
    (root / "paopao_task.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    write_task_local_gitignore(root)
    return root


def cmd_init(args: argparse.Namespace) -> int:
    root = create_task_dir(
        name=args.name,
        output_root=args.output_root,
        pages=args.pages,
        language=args.language,
        focus=args.focus,
    )
    print(root)
    return 0


def html_files_from_task(task_dir: Path) -> list[Path]:
    html_dir = task_dir / "html"
    files = sorted(html_dir.glob("slide*.html"))
    if not files:
        raise SystemExit(f"No slide*.html files found in {html_dir}")
    return files


def expected_pages_from_task(task_dir: Path) -> int | None:
    data = read_task_manifest(task_dir)
    pages = data.get("page_count")
    return pages if isinstance(pages, int) and pages > 0 else None


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception:
        return ""


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def pipeline_pass_path(task_dir: Path) -> Path:
    return task_dir / "qa" / "pipeline_pass.json"


def final_delivery_pass_path(task_dir: Path) -> Path:
    return task_dir / "qa" / "final_delivery_pass.json"


def commercial_render_contract_path(task_dir: Path) -> Path:
    return task_dir / "qa" / "commercial_render_contract.json"


def _relative_to_task_or_abs(task_dir: Path, path: Path) -> str:
    try:
        return str(path.resolve().relative_to(task_dir.resolve()))
    except ValueError:
        return str(path.resolve())


def _resolve_task_path(task_dir: Path, value: object) -> Path | None:
    if not isinstance(value, str) or not value:
        return None
    path = Path(value)
    if path.is_absolute():
        return path
    return task_dir / path


def _read_json_file(path: Path) -> dict[str, object] | None:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def read_task_manifest(task_dir: Path) -> dict[str, object]:
    manifest = task_dir / "paopao_task.json"
    if not manifest.exists():
        return {}
    try:
        data = json.loads(manifest.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def requested_language_family(task_dir: Path) -> str:
    raw = str(read_task_manifest(task_dir).get("language", "")).strip().lower()
    if not raw:
        return ""
    if any(token in raw for token in ["中文", "汉语", "chinese", "mandarin", "zh", "cn"]):
        return "zh"
    if any(token in raw for token in ["english", "英文", "英语", "en"]):
        return "en"
    return ""


def visible_text_from_html(text: str) -> str:
    text = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", text)
    text = re.sub(r"(?is)<!--.*?-->", " ", text)
    text = re.sub(r"(?is)<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def cjk_char_count(text: str) -> int:
    return sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")


def english_word_count(text: str) -> int:
    count = 0
    for word in re.findall(r"\b[A-Za-z][A-Za-z'-]{2,}\b", text):
        if word.isupper() and len(word) <= 5:
            continue
        count += 1
    return count


def language_consistency_issues(text: str, label: str, language_family: str) -> list[str]:
    if not language_family:
        return []
    visible = visible_text_from_html(text)
    cjk = cjk_char_count(visible)
    eng = english_word_count(visible)
    issues: list[str] = []
    if language_family == "en" and cjk:
        issues.append(f"{label}: target language is English but visible text contains CJK characters")
    if language_family == "zh":
        forbidden_labels = [
            "key takeaway",
            "source:",
            "situation",
            "complication",
            "resolution",
            "agenda",
            "executive summary",
        ]
        lower = visible.lower()
        found = [item for item in forbidden_labels if item in lower]
        if found:
            issues.append(
                f"{label}: target language is Chinese but visible text contains English UI labels: "
                + ", ".join(found)
            )
        if cjk >= 20 and eng > max(10, cjk // 8):
            issues.append(
                f"{label}: target language is Chinese but visible text appears mixed "
                f"({eng} English words vs {cjk} Chinese characters)"
            )
    return issues


def extract_prompt_title(text: str) -> str:
    match = re.search(r"^\s*TITLE\s*:\s*(?P<title>.+?)\s*$", text, flags=re.MULTILINE)
    if not match:
        return ""
    title = match.group("title").strip()
    title = title.strip("\"'")
    return re.sub(r"\s+", " ", title).strip()


def compact_nav_label(title: str, idx: int, language_family: str) -> str:
    if not title:
        return f"{idx:02d}"
    if language_family == "zh" or cjk_char_count(title) >= 4:
        cjk = "".join(ch for ch in title if "\u4e00" <= ch <= "\u9fff")
        return cjk[:6] or f"{idx:02d}"
    words = [
        word
        for word in re.findall(r"[A-Za-z0-9][A-Za-z0-9&/-]*", title)
        if word.lower() not in {"the", "and", "for", "with", "from", "into", "that", "this"}
    ]
    label = " ".join(words[:2]).strip()
    return label[:22] if label else f"{idx:02d}"


def deck_navigation_labels(task_dir: Path, expected: int) -> list[str]:
    manifest = read_task_manifest(task_dir)
    raw_labels = manifest.get("navigation_labels")
    if isinstance(raw_labels, list) and len(raw_labels) == expected:
        labels = [str(label).strip() for label in raw_labels]
        if all(labels):
            return labels
    language_family = requested_language_family(task_dir)
    labels: list[str] = []
    for idx in range(1, expected + 1):
        prompt_text = read_text(task_dir / "analysis" / f"final_prompt_{idx:02d}.md")
        title = extract_prompt_title(prompt_text)
        labels.append(compact_nav_label(title, idx, language_family))
    return labels


def build_deck_navigation_contract(task_dir: Path, idx: int) -> str:
    expected = expected_pages_from_task(task_dir)
    if not expected:
        return ""
    labels = deck_navigation_labels(task_dir, expected)
    active = labels[idx - 1] if 1 <= idx <= len(labels) else f"{idx:02d}"
    label_lines = "\n".join(
        f"- {pos:02d}: {label}{' [ACTIVE]' if pos == idx else ''}"
        for pos, label in enumerate(labels, 1)
    )
    return f"""

DECK NAVIGATION CONTRACT:
- This deck uses a persistent top directory strip on every slide. Do not omit it on summary, dashboard, chart, matrix, SCR, cover-like, or dense data slides.
- Place the navigation before the title as a slim full-width bar, approx. 36-42 px high on a 1920x1080 canvas.
- Use the same item count, order, labels, color system, and page number on every slide; only the active item changes.
- Visual style: one continuous deep-blue bar with compact directory text. Do not render labels as large equal-width filled web tabs or boxed buttons.
- Active slide: {idx:02d} / {expected}, active tab label: "{active}".
- Navigation labels:
{label_lines}
- Use #305496 for the bar, white text, subtle separators such as dots or thin ticks, and a small right-aligned page number. Active label may be bold, underlined, or lightly accented with #4472C4, but should not become a large filled rectangle.
- The visual contract must record nav as a visible region with non-zero bbox and the actual labels; HTML must bind it as data-ref-id="nav" using semantic nav children.
""".rstrip()


def image_dimensions(path: Path) -> tuple[int, int] | None:
    try:
        with path.open("rb") as f:
            header = f.read(32)
            if header.startswith(b"\x89PNG\r\n\x1a\n") and header[12:16] == b"IHDR":
                width, height = struct.unpack(">II", header[16:24])
                return int(width), int(height)
            if header.startswith(b"\xff\xd8"):
                f.seek(2)
                while True:
                    marker_start = f.read(1)
                    if not marker_start:
                        return None
                    if marker_start != b"\xff":
                        continue
                    marker = f.read(1)
                    while marker == b"\xff":
                        marker = f.read(1)
                    if marker in {b"\xc0", b"\xc1", b"\xc2", b"\xc3"}:
                        length = struct.unpack(">H", f.read(2))[0]
                        data = f.read(length - 2)
                        height, width = struct.unpack(">HH", data[1:5])
                        return int(width), int(height)
                    if marker in {b"\xd8", b"\xd9"}:
                        continue
                    length_data = f.read(2)
                    if len(length_data) != 2:
                        return None
                    length = struct.unpack(">H", length_data)[0]
                    f.seek(length - 2, os.SEEK_CUR)
    except Exception:
        return None
    return None


def image_similarity_score(reference: Path, actual: Path) -> float | None:
    """Return a rough structural similarity score in [0, 1] for two slide images.

    This is intentionally simple and local-only: it catches gross drift between
    the selected Image2 reference and final PPTX preview before delivery. Agents
    can still add human fidelity evidence, but cannot make a rough rebuild pass
    by writing a positive JSON review.
    """
    try:
        from PIL import Image, ImageChops, ImageFilter, ImageStat
    except Exception:
        return None
    try:
        with Image.open(reference) as ref_img, Image.open(actual) as actual_img:
            size = (320, 180)
            # Compare coarse structure rather than glyph-level antialiasing.
            # Image2 references and browser/PPT renders rarely share identical
            # font rasterization, so unblurred edge maps over-penalize otherwise
            # faithful editable rebuilds. A light blur keeps module geometry,
            # bands, charts, and large title placement visible while damping
            # text-edge noise.
            ref = ref_img.convert("L").resize(size).filter(ImageFilter.GaussianBlur(1))
            act = actual_img.convert("L").resize(size).filter(ImageFilter.GaussianBlur(1))
            ref_edges = ref.filter(ImageFilter.FIND_EDGES).filter(ImageFilter.GaussianBlur(1))
            act_edges = act.filter(ImageFilter.FIND_EDGES).filter(ImageFilter.GaussianBlur(1))
            pixel_diff = ImageChops.difference(ref, act)
            edge_diff = ImageChops.difference(ref_edges, act_edges)
            pixel_rms = ImageStat.Stat(pixel_diff).rms[0] / 255
            edge_rms = ImageStat.Stat(edge_diff).rms[0] / 255
            # Blend tonal and structural similarity. Edge score is weighted more
            # heavily because layout drift is usually worse than minor text/color
            # antialiasing differences.
            score = 1 - (0.50 * pixel_rms + 0.50 * edge_rms)
            return max(0.0, min(1.0, float(score)))
    except Exception:
        return None


def html_reference_preview_path(task_dir: Path, idx: int) -> Path:
    return task_dir / "qa" / "html_reference" / f"slide-{idx:02d}.png"


def html_path_for_slide(task_dir: Path, idx: int) -> Path:
    return task_dir / "html" / f"slide{idx:02d}.html"


def capture_html_preview(html: Path, output: Path) -> str | None:
    """Render a slide HTML file to a PNG for pre-PPTX reference fidelity checks."""
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        return f"Playwright unavailable for HTML reference preview: {exc}"
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                viewport={"width": 1920, "height": 1080},
                device_scale_factor=1,
            )
            page.goto(html.resolve().as_uri(), wait_until="networkidle")
            page.wait_for_timeout(100)
            slide = page.locator(".slide").first
            if slide.count():
                slide.screenshot(path=str(output))
            else:
                page.screenshot(path=str(output), full_page=False)
            browser.close()
        return None
    except Exception as exc:
        return f"HTML reference preview failed for {html.name}: {exc}"


def check_html_reference_fidelity(
    task_dir: Path,
    expected: int,
    issues: list[str],
    *,
    html_files: list[Path] | None = None,
) -> dict[str, object]:
    """Block rough HTML rebuilds before PPTX render by comparing HTML PNGs to Image2.

    The visual contract and HTML are both agent-authored, so contract-vs-HTML
    checks can become self-consistent even when the result no longer resembles
    the selected reference. This gate compares pixels before rendering PPTX.
    """
    html_by_idx: dict[int, Path] = {}
    for html in html_files or sorted((task_dir / "html").glob("slide*.html")):
        match = re.search(r"slide(\d+)\.html$", html.name)
        if match:
            html_by_idx[int(match.group(1))] = html

    scores: list[dict[str, object]] = []
    for idx in range(1, expected + 1):
        ref = image2_reference_path(task_dir, idx)
        html = html_by_idx.get(idx, html_path_for_slide(task_dir, idx))
        if not ref.exists() or not html.exists():
            continue
        preview = html_reference_preview_path(task_dir, idx)
        error = capture_html_preview(html, preview)
        if error:
            issues.append(f"slide{idx:02d}.html: {error}")
            continue
        score = image_similarity_score(ref, preview)
        if score is None:
            issues.append(
                f"slide{idx:02d}.html: cannot compute HTML-vs-Image2 similarity; "
                "Pillow and readable PNG previews are required before PPTX render"
            )
            continue
        rounded = round(score, 4)
        scores.append({
            "slide": idx,
            "html_preview_path": str(preview.relative_to(task_dir)),
            "image_similarity_score": rounded,
        })
        if score < MIN_HTML_REFERENCE_IMAGE_SCORE:
            issues.append(
                f"slide{idx:02d}.html: HTML preview image similarity {score:.3f} below minimum "
                f"{MIN_HTML_REFERENCE_IMAGE_SCORE:.3f}; rebuild HTML from the Image2 reference before rendering PPTX"
            )
    return {
        "min_html_reference_image_score": MIN_HTML_REFERENCE_IMAGE_SCORE,
        "image_similarity_scores": scores,
    }


def is_image2_widescreen(path: Path) -> tuple[bool, str]:
    dims = image_dimensions(path)
    if dims is None:
        return False, "cannot read image dimensions"
    width, height = dims
    if width <= 0 or height <= 0:
        return False, f"invalid image dimensions {width}x{height}"
    aspect = width / height
    if abs(aspect - IMAGE2_TARGET_ASPECT) > IMAGE2_ASPECT_TOLERANCE:
        return False, f"image must be 16:9 landscape; found {width}x{height} aspect={aspect:.3f}"
    return True, f"{width}x{height}"


def has_placeholder(text: str) -> bool:
    lower = text.lower()
    return any(pattern in lower for pattern in PLACEHOLDER_PATTERNS)


_remote_prompt_cache: dict[str, str] = {}


def prompt_template_path(name: str) -> Path:
    return PROMPT_LIBRARY_DIR / name


def read_prompt_template(name: str) -> str:
    local = PROMPT_LIBRARY_DIR / name
    if local.exists():
        return read_text(local)
    if name in _remote_prompt_cache:
        return _remote_prompt_cache[name]
    try:
        content = paopao_auth.fetch_prompt_content(name)
        _remote_prompt_cache[name] = content
        return content
    except Exception:
        pass
    return ""


def prompt_selection_plan_path(task_dir: Path) -> Path:
    return task_dir / PROMPT_SELECTION_PLAN_PATH


def selected_prompt_template(text: str) -> str | None:
    match = PROMPT_TEMPLATE_RE.search(text)
    return match.group("name") if match else None


def prompt_scaffold_family(template_name: str) -> str:
    prefix = template_name[:2]
    return PROMPT_SCAFFOLD_FAMILY_BY_PREFIX.get(prefix, "unknown")


def prompt_visual_grammar(family: str) -> str:
    return PROMPT_VISUAL_GRAMMAR_BY_FAMILY.get(family, family or "unknown")


def extract_prompt_field(text: str, field_name: str) -> str:
    pattern = rf"^{re.escape(field_name)}\s*:\s*(.+?)(?=^[A-Z_]+:|\Z)"
    match = re.search(pattern, text, re.MULTILINE | re.DOTALL)
    if not match:
        return ""
    return re.sub(r"\s+", " ", match.group(1).strip())


def parse_data_requires(text: str) -> list[str]:
    raw = extract_prompt_field(text, "DATA_REQUIRES")
    if not raw:
        return []
    values = []
    for part in re.split(r"[,;，、]\s*", raw):
        label = part.strip()
        if re.fullmatch(r"[a-z_]+", label):
            values.append(label)
    return values


def _load_local_prompt_catalog() -> list[dict[str, object]]:
    catalog: list[dict[str, object]] = []
    for path in sorted(PROMPT_LIBRARY_DIR.glob("*.md")):
        if path.name in {"SYSTEM_PROMPT.md", "INDEX.md"}:
            continue
        text = read_text(path)
        layout_match = LAYOUT_NAME_RE.search(text)
        layout_name = layout_match.group("name") if layout_match else path.stem
        when = extract_prompt_field(text, "WHEN_TO_USE") or extract_prompt_field(text, "LAYOUT")
        catalog.append({
            "template": path.name,
            "layout_name": layout_name,
            "family": prompt_scaffold_family(path.name),
            "when_to_use": when[:320],
            "data_requires": parse_data_requires(text),
        })
    return catalog


def load_prompt_catalog() -> list[dict[str, object]]:
    local = _load_local_prompt_catalog()
    local_names = {str(e["template"]) for e in local}
    try:
        remote = paopao_auth.fetch_prompt_catalog()
        for entry in remote:
            name = entry.get("template", "")
            if name and name not in local_names:
                dr = entry.get("data_requires", "")
                local.append({
                    "template": name,
                    "layout_name": entry.get("layout_name", ""),
                    "family": prompt_scaffold_family(name),
                    "when_to_use": str(entry.get("when_to_use", ""))[:320],
                    "data_requires": [t.strip() for t in dr.split(",") if t.strip()] if isinstance(dr, str) else dr,
                })
    except Exception:
        pass
    return local


def extract_report_signals(analysis_report: str) -> dict[str, int]:
    return {
        label: len(pattern.findall(analysis_report))
        for label, pattern in PROMPT_DATA_SIGNAL_PATTERNS.items()
    }


def prompt_is_data_compatible(entry: dict[str, object], report_signals: dict[str, int]) -> tuple[bool, list[str]]:
    missing = [
        label for label in entry.get("data_requires", [])
        if label in PROMPT_DATA_TIER1_LABELS and report_signals.get(str(label), 0) == 0
    ]
    return not missing, [str(label) for label in missing]


def prompt_story_text(story: dict[str, str], topic: str) -> str:
    return " ".join([
        topic,
        story.get("section_name", ""),
        story.get("role", ""),
        story.get("brief", ""),
    ]).strip()


def prompt_role_base_score(entry: dict[str, object], story_text: str, slide_idx: int, expected: int) -> int:
    template = str(entry.get("template", ""))
    family = str(entry.get("family", ""))
    score = 0
    for pattern, family_scores in PROMPT_ROLE_FAMILY_PREFERENCES:
        if pattern.search(story_text):
            score += int(family_scores.get(family, 0))
    for pattern, target_template, bonus in PROMPT_TEMPLATE_KEYWORD_BONUSES:
        if template == target_template and pattern.search(story_text):
            score += bonus
    if slide_idx == 1 and family in {"dashboard", "kpi-hero", "pyramid-funnel", "two-column"}:
        score += 8
    if slide_idx == expected and family in {"process-flow", "swimlane", "center-radial", "dashboard"}:
        score += 8
    if family == "divider-cover" and expected <= 5:
        score -= 40
    return score


def prompt_candidate_score(
    entry: dict[str, object],
    story_text: str,
    slide_idx: int,
    expected: int,
    used_templates: set[str],
    used_families: set[str],
    used_grammars: set[str],
    previous_family: str,
    previous_grammar: str,
) -> tuple[int, list[str]]:
    template = str(entry.get("template", ""))
    family = str(entry.get("family", ""))
    grammar = prompt_visual_grammar(family)
    score = prompt_role_base_score(entry, story_text, slide_idx, expected)
    reasons: list[str] = []
    if score:
        reasons.append(f"role_fit={score}")
    if template in used_templates:
        score -= 1000
        reasons.append("repeat_template=-1000")
    if family in used_families:
        score -= 180
        reasons.append("repeat_family=-180")
    if grammar in used_grammars:
        score -= 70
        reasons.append("repeat_visual_grammar=-70")
    if family == previous_family:
        score -= 420
        reasons.append("adjacent_family_collision=-420")
    elif grammar == previous_grammar:
        score -= 260
        reasons.append("adjacent_visual_grammar_collision=-260")
    return score, reasons


def load_slide_story(task_dir: Path) -> dict[int, dict[str, str]]:
    path = task_dir / "analysis" / "slide_story.json"
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    slides = data.get("slides") if isinstance(data, dict) else data
    if not isinstance(slides, list):
        return {}
    out: dict[int, dict[str, str]] = {}
    for item in slides:
        if not isinstance(item, dict):
            continue
        slide_value = item.get("slide")
        if not isinstance(slide_value, int):
            slide_value = item.get("slide_number")
        if not isinstance(slide_value, int):
            continue
        out[int(slide_value)] = {
            "brief": str(item.get("brief", "") or item.get("claim", "") or item.get("title", "")),
            "role": str(item.get("role", "") or item.get("slide_role", "")),
            "section_name": str(item.get("section_name", "")),
        }
    return out


def select_prompt_plan(task_dir: Path, expected: int, topic: str = "") -> dict[str, object]:
    analysis_text = read_text(task_dir / "analysis" / "analysis_report.md")
    slide_story = load_slide_story(task_dir)
    catalog = load_prompt_catalog()
    report_signals = extract_report_signals(analysis_text + "\n" + topic)
    rejected: list[dict[str, object]] = []
    compatible: list[dict[str, object]] = []
    for entry in catalog:
        ok, missing = prompt_is_data_compatible(entry, report_signals)
        if ok:
            compatible.append(entry)
        else:
            rejected.append({
                "template": entry.get("template"),
                "layout_name": entry.get("layout_name"),
                "family": entry.get("family"),
                "missing_tier1_data": missing,
            })
    if not compatible:
        compatible = list(catalog)
    elif len(compatible) < min(3, len(catalog)):
        compatible = list(catalog)

    selected: list[dict[str, object]] = []
    used_templates: set[str] = set()
    used_families: set[str] = set()
    used_grammars: set[str] = set()
    previous_family = ""
    previous_grammar = ""
    for idx in range(1, expected + 1):
        story = slide_story.get(idx, {})
        story_text = prompt_story_text(story, topic)
        scored: list[tuple[int, str, dict[str, object], list[str]]] = []
        for entry in compatible:
            score, reasons = prompt_candidate_score(
                entry,
                story_text,
                idx,
                expected,
                used_templates,
                used_families,
                used_grammars,
                previous_family,
                previous_grammar,
            )
            scored.append((score, str(entry.get("template", "")), entry, reasons))
        scored.sort(key=lambda item: (-item[0], item[1]))
        chosen_score, _, chosen, chosen_reasons = scored[0]
        candidate_pool = [entry for _, _, entry, _ in scored[:3]]
        if len(candidate_pool) < 3:
            candidate_pool.extend(entry for entry in compatible if entry not in candidate_pool)
        if len(candidate_pool) < 3:
            candidate_pool.extend(entry for entry in catalog if entry not in candidate_pool)
        candidate_pool = candidate_pool[:3]
        score_by_template = {
            str(scored_entry.get("template", "")): score
            for score, _, scored_entry, _ in scored
        }
        candidates = [
            {
                "rank": rank + 1,
                "template": candidate["template"],
                "layout_name": candidate["layout_name"],
                "family": candidate["family"],
                "visual_grammar": prompt_visual_grammar(str(candidate["family"])),
                "selection_method": "role_fit_with_visual_diversity",
                "score": score_by_template.get(str(candidate.get("template", "")), 0),
                "when_to_use": candidate.get("when_to_use", ""),
            }
            for rank, candidate in enumerate(candidate_pool)
        ]
        family = str(chosen["family"])
        grammar = prompt_visual_grammar(family)
        selected.append({
            "slide": idx,
            "section_name": story.get("section_name", ""),
            "role": story.get("role", ""),
            "brief": story.get("brief", ""),
            "selected_template": chosen["template"],
            "layout_name": chosen["layout_name"],
            "family": family,
            "visual_grammar": grammar,
            "selection_method": "role_fit_with_visual_diversity",
            "score": chosen_score,
            "score_reasons": chosen_reasons,
            "candidates": candidates,
        })
        used_templates.add(str(chosen["template"]))
        used_families.add(family)
        used_grammars.add(grammar)
        previous_family = family
        previous_grammar = grammar

    return {
        "schema": PROMPT_SELECTION_PLAN_SCHEMA,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "task_dir": str(task_dir),
        "expected_pages": expected,
        "topic": topic,
        "catalog_size": len(catalog),
        "compatible_catalog_size": len(compatible),
        "rejected_for_missing_tier1_data": rejected[:40],
        "report_signals": report_signals,
        "selection_rule": (
            "deterministic role-fit scoring with tier-1 data compatibility, template uniqueness, "
            "scaffold-family diversity, and adjacent visual-grammar collision penalties"
        ),
        "slide_story_path": "analysis/slide_story.json" if slide_story else None,
        "slides": selected,
    }


def cmd_plan_prompts(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")
    plan = select_prompt_plan(task_dir, int(expected), topic=str(args.topic or ""))
    out = prompt_selection_plan_path(task_dir)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "plan": str(out), "selected": plan["slides"]}, ensure_ascii=False, indent=2))
    return 0


def prompt_template_issue(text: str, prompt_name: str) -> list[str]:
    issues: list[str] = []
    match = PROMPT_TEMPLATE_RE.search(text)
    if not match:
        return [
            f"{prompt_name} missing PROMPT_TEMPLATE: <prompt-library-file>.md; final prompts must be filled from plugins/paopao-codex-plugin/prompts"
        ]
    template_name = match.group("name")
    template = prompt_template_path(template_name)
    template_text = read_prompt_template(template_name)
    if (not template_text) or template.name == SYSTEM_PROMPT.name:
        return [f"{prompt_name} PROMPT_TEMPLATE does not exist in prompt library: {template_name}"]
    layout_match = LAYOUT_NAME_RE.search(template_text)
    if not layout_match:
        issues.append(f"{prompt_name} PROMPT_TEMPLATE {template_name} missing LAYOUT_NAME in library")
        return issues
    layout_name = layout_match.group("name")
    if f"LAYOUT_NAME: {layout_name}" not in text:
        issues.append(
            f"{prompt_name} must include LAYOUT_NAME: {layout_name} from PROMPT_TEMPLATE {template_name}"
        )
    if layout_name not in text:
        issues.append(f"{prompt_name} does not reference selected prompt-library layout name: {layout_name}")
    return issues


def prompt_selection_diversity_issues(selections: list[tuple[int, str, str]]) -> list[str]:
    issues: list[str] = []
    if len(selections) < 2:
        return issues
    template_seen: dict[str, int] = {}
    family_seen: dict[str, int] = {}
    previous_idx = 0
    previous_family = ""
    previous_grammar = ""
    for idx, template_name, family in selections:
        grammar = prompt_visual_grammar(family)
        if template_name in template_seen:
            issues.append(
                f"final_prompt_{idx:02d}.md repeats PROMPT_TEMPLATE {template_name}; "
                f"already used on slide {template_seen[template_name]:02d}. Select a different prompt-library template unless the audit gives a content-specific exception."
            )
        else:
            template_seen[template_name] = idx
        if family in family_seen:
            issues.append(
                f"prompt selection plan slide {idx} repeats scaffold family {family}; "
                f"already used on slide {family_seen[family]:02d}. Select a visually distinct prompt family."
            )
        else:
            family_seen[family] = idx
        if previous_idx and family == previous_family:
            issues.append(
                f"prompt selection plan slides {previous_idx} and {idx} use the same scaffold family {family}; adjacent slides must not share the same page structure."
            )
        if previous_idx and grammar == previous_grammar:
            issues.append(
                f"prompt selection plan slides {previous_idx} and {idx} share visual grammar {grammar}; adjacent slides must have visibly different composition."
            )
        previous_idx = idx
        previous_family = family
        previous_grammar = grammar
    return issues

def prompt_selection_plan_issues(task_dir: Path, expected: int, selections: list[tuple[int, str, str]]) -> list[str]:
    issues: list[str] = []
    plan_path = prompt_selection_plan_path(task_dir)
    if not plan_path.exists():
        issues.append("analysis/prompt_selection_plan.json missing; run plan-prompts before writing final_prompt_XX.md")
        return issues
    try:
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return [f"analysis/prompt_selection_plan.json cannot be parsed: {exc}"]
    if not isinstance(plan, dict):
        return ["analysis/prompt_selection_plan.json must be a JSON object"]
    if plan.get("schema") != PROMPT_SELECTION_PLAN_SCHEMA:
        issues.append(f"analysis/prompt_selection_plan.json schema must be {PROMPT_SELECTION_PLAN_SCHEMA}")
    if int(plan.get("expected_pages", 0) or 0) != expected:
        issues.append("analysis/prompt_selection_plan.json expected_pages does not match task page count")
    if not plan.get("slide_story_path"):
        issues.append("analysis/slide_story.json missing or not used; write per-slide role/brief before running plan-prompts")
    slides = plan.get("slides")
    if not isinstance(slides, list) or len(slides) != expected:
        issues.append(f"analysis/prompt_selection_plan.json must contain exactly {expected} slide selections")
        return issues
    selected_by_slide: dict[int, str] = {}
    for item in slides:
        if not isinstance(item, dict):
            continue
        slide = item.get("slide")
        template = str(item.get("selected_template", "")).strip()
        family = str(item.get("family", "")).strip()
        visual_grammar = str(item.get("visual_grammar", "")).strip()
        candidates = item.get("candidates")
        if not isinstance(slide, int):
            issues.append("prompt selection plan slide entry missing numeric slide")
            continue
        if not template:
            issues.append(f"prompt selection plan slide {slide}: selected_template missing")
            continue
        if prompt_scaffold_family(template) != family:
            issues.append(f"prompt selection plan slide {slide}: family does not match selected_template prefix")
        expected_grammar = prompt_visual_grammar(family)
        if visual_grammar and visual_grammar != expected_grammar:
            issues.append(f"prompt selection plan slide {slide}: visual_grammar does not match family")
        if not isinstance(candidates, list) or len(candidates) < 3:
            issues.append(f"prompt selection plan slide {slide}: must record at least 3 candidates")
        selected_by_slide[slide] = template

    actual_by_slide = {idx: template for idx, template, _family in selections}
    for idx in range(1, expected + 1):
        planned = selected_by_slide.get(idx)
        actual = actual_by_slide.get(idx)
        if planned and actual and planned != actual:
            issues.append(
                f"final_prompt_{idx:02d}.md uses {actual}, but prompt_selection_plan.json selected {planned}; "
                "rerun plan-prompts or revise the final prompt to match the selected template"
            )
    issues.extend(prompt_selection_diversity_issues([
        (idx, selected_by_slide[idx], prompt_scaffold_family(selected_by_slide[idx]))
        for idx in sorted(selected_by_slide)
    ]))
    return issues


def image2_prompt_path(task_dir: Path, idx: int) -> Path:
    return task_dir / "image2" / f"image2_prompt_{idx:02d}.md"


def image2_generation_request_path(task_dir: Path, idx: int) -> Path:
    return task_dir / "image2" / f"generation_request_{idx:02d}.json"


def image2_manifest_path(task_dir: Path) -> Path:
    return task_dir / "image2" / "image2_generation_manifest.json"


def image2_reference_path(task_dir: Path, idx: int) -> Path:
    return task_dir / "image2" / f"image2_reference_{idx:02d}.png"


def image2_style_review_path(task_dir: Path) -> Path:
    return task_dir / "qa" / "image2_style_review.json"


def image2_user_review_path(task_dir: Path) -> Path:
    return task_dir / "qa" / "image2_user_review.json"


def image2_observation_path(task_dir: Path, idx: int) -> Path:
    return task_dir / "spec" / f"slide{idx:02d}_image_observation.json"


def post_image_memory_boundary_path(task_dir: Path) -> Path:
    return task_dir / "qa" / "post_image_memory_boundary.json"


def path_contains_any(path: Path, parts: set[str]) -> bool:
    normalized = [part.lower() for part in path.parts]
    return any(part.lower() in normalized for part in parts)


def name_contains_any(path: Path, fragments: set[str]) -> bool:
    name = path.name.lower()
    return any(fragment.lower() in name for fragment in fragments)


def looks_like_rendered_preview_image(path: Path) -> bool:
    lower_parts = [part.lower() for part in path.parts]
    lower_name = path.name.lower()
    if "pptx_actual" in lower_parts or "qa" in lower_parts:
        return True
    if lower_name.startswith("slide_") or lower_name.startswith("slide-"):
        return True
    return name_contains_any(path, IMAGE2_FORBIDDEN_SOURCE_NAME_PARTS)


def file_is_after_reference(task_dir: Path, idx: int, path: Path) -> bool:
    ref = image2_reference_path(task_dir, idx)
    if not path.exists() or not ref.exists():
        return True
    return path.stat().st_mtime >= ref.stat().st_mtime


def file_is_after_memory_boundary(task_dir: Path, path: Path) -> bool:
    boundary = post_image_memory_boundary_path(task_dir)
    if not path.exists() or not boundary.exists():
        return True
    return path.stat().st_mtime >= boundary.stat().st_mtime


def image2_reference_provenance_issues(task_dir: Path, idx: int, entry: dict[str, object]) -> list[str]:
    issues: list[str] = []
    source_kind = str(entry.get("registration_source_kind", "")).strip()
    tool_call_id = str(entry.get("tool_call_id", "")).strip()
    registered_source_raw = str(entry.get("registered_source", "")).strip()
    if source_kind not in IMAGE2_ALLOWED_SOURCE_KINDS:
        issues.append(
            f"image2_generation_manifest slide {idx}: registration_source_kind must be one of "
            f"{', '.join(sorted(IMAGE2_ALLOWED_SOURCE_KINDS))}; found {source_kind or '<missing>'}"
        )
    if len(tool_call_id) < 8:
        issues.append(
            f"image2_generation_manifest slide {idx}: tool_call_id/provenance id is required and must not be blank"
        )
    if not registered_source_raw:
        issues.append(f"image2_generation_manifest slide {idx}: registered_source missing")
        return issues
    registered_source = Path(registered_source_raw)
    if not registered_source.is_absolute():
        registered_source = task_dir / registered_source
    if registered_source.resolve() == image2_reference_path(task_dir, idx).resolve():
        issues.append(
            f"image2_generation_manifest slide {idx}: registered_source is the final image2_reference target; "
            "it must point to the original image-generation output"
        )
    try:
        rel = registered_source.resolve().relative_to(task_dir.resolve())
        issues.append(
            f"image2_generation_manifest slide {idx}: registered_source is inside the task directory; "
            "register from the original generated image outside output/<task>"
        )
        if path_contains_any(rel, IMAGE2_FORBIDDEN_SOURCE_PARTS) or name_contains_any(rel, IMAGE2_FORBIDDEN_SOURCE_NAME_PARTS):
            issues.append(
                f"image2_generation_manifest slide {idx}: registered_source points inside a build/preview folder, "
                "not an image-generation output"
            )
    except Exception:
        pass
    if looks_like_rendered_preview_image(registered_source):
        issues.append(
            f"image2_generation_manifest slide {idx}: registered_source looks like an HTML/PPTX preview screenshot"
        )
    return issues


def image2_generation_request_issues(task_dir: Path, idx: int, entry: dict[str, object]) -> list[str]:
    issues: list[str] = []
    prompt_path = image2_prompt_path(task_dir, idx)
    request_path = image2_generation_request_path(task_dir, idx)
    if not request_path.exists():
        return [f"image2 generation request slide {idx}: generation_request_{idx:02d}.json missing"]
    data = load_generation_request(request_path)
    if data is None:
        return [f"image2 generation request slide {idx}: file cannot be parsed"]
    prompt_text = read_text(prompt_path)
    prompt_sha = sha256_text(prompt_text) if prompt_text else str(entry.get("image2_prompt_sha256", ""))
    request_prompt_text = str(data.get("prompt_text", ""))
    if data.get("schema") != IMAGE2_GENERATION_REQUEST_SCHEMA:
        issues.append(f"image2 generation request slide {idx}: schema must be {IMAGE2_GENERATION_REQUEST_SCHEMA}")
    if data.get("slide") != idx:
        issues.append(f"image2 generation request slide {idx}: slide mismatch")
    if data.get("prompt_source_path") != f"image2/image2_prompt_{idx:02d}.md":
        issues.append(f"image2 generation request slide {idx}: prompt_source_path mismatch")
    if data.get("prompt_sha256") != prompt_sha:
        issues.append(f"image2 generation request slide {idx}: prompt_sha256 mismatch")
    if not request_prompt_text.strip():
        issues.append(f"image2 generation request slide {idx}: prompt_text missing")
    elif sha256_text(request_prompt_text) != prompt_sha:
        issues.append(
            f"image2 generation request slide {idx}: prompt_text must exactly equal image2_prompt_{idx:02d}.md"
        )
    if data.get("prompt_transfer_method") != IMAGE2_PROMPT_TRANSFER_METHOD:
        issues.append(
            f"image2 generation request slide {idx}: prompt_transfer_method must be {IMAGE2_PROMPT_TRANSFER_METHOD}"
        )
    if data.get("manual_prompt_rewrite_allowed") is not False:
        issues.append(f"image2 generation request slide {idx}: manual_prompt_rewrite_allowed must be false")
    if entry.get("generation_request_path") != f"image2/generation_request_{idx:02d}.json":
        issues.append(f"image2_generation_manifest slide {idx}: generation_request_path mismatch")
    if entry.get("generation_request_sha256") != sha256_file(request_path):
        issues.append(f"image2_generation_manifest slide {idx}: generation_request_sha256 mismatch")
    if entry.get("generation_request_id") != data.get("request_id"):
        issues.append(f"image2_generation_manifest slide {idx}: generation_request_id mismatch")
    if entry.get("prompt_transfer_method") != IMAGE2_PROMPT_TRANSFER_METHOD:
        issues.append(
            f"image2_generation_manifest slide {idx}: prompt_transfer_method must be {IMAGE2_PROMPT_TRANSFER_METHOD}"
        )
    return issues


def load_image2_manifest(task_dir: Path) -> dict[str, object]:
    manifest_path = image2_manifest_path(task_dir)
    if not manifest_path.exists():
        return {}
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def manifest_entries_by_slide(manifest: dict[str, object]) -> dict[int, dict[str, object]]:
    entries: dict[int, dict[str, object]] = {}
    slides = manifest.get("slides")
    if not isinstance(slides, list):
        return entries
    for entry in slides:
        if not isinstance(entry, dict):
            continue
        slide = entry.get("slide")
        if isinstance(slide, int):
            entries[slide] = entry
    return entries


def prompt_private_path(task_dir: Path, path: Path) -> Path:
    return task_dir / PROMPT_PRIVATE_DIR / path.relative_to(task_dir)


def build_image2_prompt_text(task_dir: Path, idx: int) -> str:
    system_text = read_text(SYSTEM_PROMPT)
    slide_prompt = read_text(task_dir / "analysis" / f"final_prompt_{idx:02d}.md")
    navigation_contract = build_deck_navigation_contract(task_dir, idx)
    if not system_text.strip():
        raise SystemExit(f"System prompt missing or empty: {SYSTEM_PROMPT}")
    if not slide_prompt.strip():
        raise SystemExit(f"final_prompt_{idx:02d}.md missing or empty")
    return (
        f"{system_text.rstrip()}\n\n"
        "---\n\n"
        f"{slide_prompt.rstrip()}\n"
        f"{navigation_contract}\n"
        f"{IMAGE2_LOCKED_TAIL}"
    )


def build_image2_generation_request(task_dir: Path, idx: int, prompt_text: str) -> dict[str, object]:
    prompt_path = image2_prompt_path(task_dir, idx)
    prompt_sha = sha256_text(prompt_text)
    request_seed = f"{task_dir.name}:{idx}:{prompt_sha}:{IMAGE2_PROMPT_TRANSFER_METHOD}"
    request_id = sha256_text(request_seed)[:24]
    return {
        "schema": IMAGE2_GENERATION_REQUEST_SCHEMA,
        "request_id": request_id,
        "slide": idx,
        "prompt_source_path": str(prompt_path.relative_to(task_dir)),
        "prompt_sha256": prompt_sha,
        "prompt_transfer_method": IMAGE2_PROMPT_TRANSFER_METHOD,
        "manual_prompt_rewrite_allowed": False,
        "image_generation_input_contract": (
            "Use prompt_text exactly as the image-generation prompt. Do not summarize, "
            "compress, translate, rewrite, or add a separate chat-authored prompt."
        ),
        "prompt_text": prompt_text,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
    }


def load_generation_request(path: Path) -> dict[str, object] | None:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def cmd_prepare_image2_prompts(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = expected_pages_from_task(task_dir)
    if expected is None:
        print("Cannot determine expected page count from paopao_task.json", file=sys.stderr)
        return 1
    analysis_issues: list[str] = []
    check_analysis_files(task_dir, expected, analysis_issues)
    if analysis_issues:
        print(json.dumps({
            "task_dir": str(task_dir),
            "stage": "prepare-image2-preflight",
            "ok": False,
            "issues": analysis_issues,
        }, indent=2, ensure_ascii=False))
        return 1

    image2_dir = task_dir / "image2"
    image2_dir.mkdir(parents=True, exist_ok=True)
    previous = manifest_entries_by_slide(load_image2_manifest(task_dir))
    slides: list[dict[str, object]] = []
    for idx in range(1, expected + 1):
        final_prompt = task_dir / "analysis" / f"final_prompt_{idx:02d}.md"
        prompt_text = build_image2_prompt_text(task_dir, idx)
        prompt_path = image2_prompt_path(task_dir, idx)
        prompt_path.write_text(prompt_text, encoding="utf-8")
        generation_request = build_image2_generation_request(task_dir, idx, prompt_text)
        request_path = image2_generation_request_path(task_dir, idx)
        request_path.write_text(
            json.dumps(generation_request, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        image_path = image2_reference_path(task_dir, idx)
        prompt_sha = sha256_text(prompt_text)
        final_prompt_sha = sha256_file(final_prompt) if final_prompt.exists() else None
        system_prompt_sha = sha256_file(SYSTEM_PROMPT)
        request_sha = sha256_file(request_path)
        image_sha = sha256_file(image_path) if image_path.exists() else None
        base_entry: dict[str, object] = {
            "slide": idx,
            "status": "prompt_prepared",
            "final_prompt_path": str(final_prompt.relative_to(task_dir)),
            "final_prompt_sha256": final_prompt_sha,
            "system_prompt_path": str(SYSTEM_PROMPT),
            "system_prompt_sha256": system_prompt_sha,
            "image2_prompt_path": str(prompt_path.relative_to(task_dir)),
            "image2_prompt_sha256": prompt_sha,
            "generation_request_path": str(request_path.relative_to(task_dir)),
            "generation_request_sha256": request_sha,
            "generation_request_id": generation_request["request_id"],
            "prompt_transfer_method": IMAGE2_PROMPT_TRANSFER_METHOD,
            "image_path": str(image_path.relative_to(task_dir)),
            "image_sha256": image_sha,
            "generated_from_image2_prompt_file": False,
            "verification_method": None,
            "manual_short_prompt_allowed": False,
            "prepared_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        }

        old = previous.get(idx, {})
        old_registered = (
            old.get("status") == "registered_verified"
            and old.get("generated_from_image2_prompt_file") is True
            and old.get("verification_method") == IMAGE2_VERIFICATION_METHOD
            and old.get("image2_prompt_sha256") == prompt_sha
            and old.get("generation_request_sha256") == request_sha
            and old.get("generation_request_id") == generation_request["request_id"]
            and old.get("prompt_transfer_method") == IMAGE2_PROMPT_TRANSFER_METHOD
            and old.get("final_prompt_sha256") == final_prompt_sha
            and old.get("system_prompt_sha256") == system_prompt_sha
            and image_sha is not None
            and old.get("image_sha256") == image_sha
        )
        if old_registered:
            preserved = {**base_entry, **old}
            preserved.update({
                "image2_prompt_sha256": prompt_sha,
                "generation_request_sha256": request_sha,
                "generation_request_id": generation_request["request_id"],
                "prompt_transfer_method": IMAGE2_PROMPT_TRANSFER_METHOD,
                "final_prompt_sha256": final_prompt_sha,
                "system_prompt_sha256": system_prompt_sha,
                "image_sha256": image_sha,
                "generated_from_image2_prompt_file": True,
                "verification_method": IMAGE2_VERIFICATION_METHOD,
                "status": "registered_verified",
            })
            slides.append(preserved)
        else:
            if image_sha is not None:
                base_entry["image_status"] = "present_but_unregistered"
            else:
                base_entry["image_status"] = "missing"
            slides.append(base_entry)

    manifest = {
        "schema": IMAGE2_MANIFEST_SCHEMA,
        "task_dir": str(task_dir),
        "expected_pages": expected,
        "rule": "prepare-image2-prompts only locks prompts. register-image2-reference is required before HTML/render and must attest the exact image2_prompt_XX.md sha.",
        "slides": slides,
    }
    image2_manifest_path(task_dir).write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


def cmd_record_image2_user_review(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")
    approved = str(args.approved).lower() in {"1", "true", "yes", "y", "approved", "ok"}
    feedback = str(args.feedback or "").strip()
    if not feedback:
        raise SystemExit("--feedback is required; record the user's approval note or requested changes")

    slides: list[dict[str, object]] = []
    missing: list[str] = []
    for idx in range(1, expected + 1):
        ref = image2_reference_path(task_dir, idx)
        if not ref.exists():
            missing.append(str(ref))
            continue
        slides.append({
            "slide": idx,
            "status": "approved" if approved else "changes_requested",
            "reference_path": str(ref.relative_to(task_dir)),
            "image_sha256": sha256_file(ref),
        })
    if missing:
        print(json.dumps({
            "ok": False,
            "issue": "missing selected image references",
            "missing": missing,
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1

    review = {
        "schema": IMAGE2_USER_REVIEW_SCHEMA,
        "reviewed_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "reviewed_after_image_generation": True,
        "user_approved": approved,
        "user_feedback": feedback,
        "slides": slides,
    }
    out = image2_user_review_path(task_dir)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(review, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": approved, "review": str(out)}, ensure_ascii=False, indent=2))
    return 0 if approved else 1


def write_post_image_memory_boundary(task_dir: Path, expected: int) -> Path:
    review_issues: list[str] = []
    check_image2_user_review(task_dir, expected, review_issues)
    if review_issues:
        raise SystemExit(
            "post-image memory boundary requires current approved Image2 user review:\n- "
            + "\n- ".join(review_issues)
        )
    slides: list[dict[str, object]] = []
    for idx in range(1, expected + 1):
        ref = image2_reference_path(task_dir, idx)
        if not ref.exists():
            raise SystemExit(f"Missing selected Image2 reference: {ref}")
        slides.append({
            "slide": idx,
            "reference_path": str(ref.relative_to(task_dir)),
            "image_sha256": sha256_file(ref),
        })
    boundary = {
        "schema": POST_IMAGE_MEMORY_BOUNDARY_SCHEMA,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "expected_pages": expected,
        "reconstruction_source": IMAGE_ONLY_RECONSTRUCTION_SOURCE,
        "prompt_context_discarded": True,
        "observed_as_fresh_image": True,
        "forbidden_after_boundary": sorted(POST_IMAGE_FORBIDDEN_MEMORY_MARKERS),
        "slides": slides,
        "policy": (
            "After this point, visual observation, contracts, specs, and HTML must be authored "
            "as if the selected Image2 references were unfamiliar external screenshots. "
            "Do not use final prompts, prompt templates, analysis notes, or remembered intent as visual evidence."
        ),
    }
    out = post_image_memory_boundary_path(task_dir)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(boundary, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def post_image_memory_boundary_issues(task_dir: Path, expected: int) -> list[str]:
    path = post_image_memory_boundary_path(task_dir)
    if not path.exists():
        return [
            "qa/post_image_memory_boundary.json missing; lock the post-image memory boundary before observation/spec/HTML"
        ]
    data = _read_json_file(path)
    if data is None:
        return ["qa/post_image_memory_boundary.json cannot be parsed"]
    issues: list[str] = []
    if data.get("schema") != POST_IMAGE_MEMORY_BOUNDARY_SCHEMA:
        issues.append(f"qa/post_image_memory_boundary.json schema must be {POST_IMAGE_MEMORY_BOUNDARY_SCHEMA}")
    if data.get("expected_pages") != expected:
        issues.append("qa/post_image_memory_boundary.json expected_pages does not match current task")
    if data.get("reconstruction_source") != IMAGE_ONLY_RECONSTRUCTION_SOURCE:
        issues.append(f"qa/post_image_memory_boundary.json reconstruction_source must be {IMAGE_ONLY_RECONSTRUCTION_SOURCE}")
    if data.get("prompt_context_discarded") is not True:
        issues.append("qa/post_image_memory_boundary.json prompt_context_discarded must be true")
    if data.get("observed_as_fresh_image") is not True:
        issues.append("qa/post_image_memory_boundary.json observed_as_fresh_image must be true")
    slides = data.get("slides")
    if not isinstance(slides, list) or len(slides) != expected:
        found = len(slides) if isinstance(slides, list) else 0
        issues.append(f"qa/post_image_memory_boundary.json slides: expected {expected}, found {found}")
        return issues
    for idx in range(1, expected + 1):
        entry = slides[idx - 1] if idx - 1 < len(slides) else {}
        if not isinstance(entry, dict):
            issues.append(f"qa/post_image_memory_boundary.json slide {idx}: entry must be an object")
            continue
        ref = image2_reference_path(task_dir, idx)
        if entry.get("reference_path") != f"image2/image2_reference_{idx:02d}.png":
            issues.append(f"qa/post_image_memory_boundary.json slide {idx}: reference_path mismatch")
        if ref.exists() and entry.get("image_sha256") != sha256_file(ref):
            issues.append(f"qa/post_image_memory_boundary.json slide {idx}: image_sha256 mismatch; rerun user review and memory boundary")
        if ref.exists() and path.stat().st_mtime < ref.stat().st_mtime:
            issues.append(f"qa/post_image_memory_boundary.json is older than image2_reference_{idx:02d}.png")
    return issues


def cmd_forget_after_image2(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")
    out = write_post_image_memory_boundary(task_dir, expected)
    print(json.dumps({
        "ok": True,
        "boundary": str(out),
        "policy": "post-image reconstruction must use selected images only",
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_record_image2_observation(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    idx = int(args.slide)
    expected = expected_pages_from_task(task_dir)
    if expected is not None and (idx < 1 or idx > expected):
        raise SystemExit(f"Slide {idx} is outside expected page count {expected}")
    evidence = str(args.evidence or "").strip()
    if len(evidence) < 120:
        raise SystemExit("--evidence must be a concrete fresh visual observation of at least 120 characters")
    forbidden_evidence = post_image_memory_markers(evidence)
    if forbidden_evidence:
        raise SystemExit(
            "--evidence must describe only visible image facts and cannot mention upstream memory markers: "
            + ", ".join(forbidden_evidence)
        )
    ref = image2_reference_path(task_dir, idx)
    if not ref.exists():
        raise SystemExit(f"Missing selected Image2 reference: {ref}")

    review_issues: list[str] = []
    if expected is not None:
        check_image2_user_review(task_dir, expected, review_issues)
    if review_issues:
        print(
            json.dumps({
                "ok": False,
                "issue": "user image review must be current and approved before recording fresh image observation",
                "issues": review_issues,
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    if expected is not None:
        boundary_issues = post_image_memory_boundary_issues(task_dir, expected)
        if boundary_issues:
            write_post_image_memory_boundary(task_dir, expected)

    observation_id = sha256_text(f"{task_dir.name}:{idx}:{sha256_file(ref)}:{evidence}")[:24]
    record = {
        "schema": IMAGE2_OBSERVATION_SCHEMA,
        "observation_id": observation_id,
        "slide": idx,
        "reference_path": str(ref.relative_to(task_dir)),
        "reference_sha256": sha256_file(ref),
        "observed_from_reference": True,
        "reconstruction_source": IMAGE_ONLY_RECONSTRUCTION_SOURCE,
        "prompt_context_discarded": True,
        "observed_as_fresh_image": True,
        "derivation_method": POST_IMAGE_DERIVATION_METHOD,
        "observation_evidence": evidence,
        "recorded_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
    }
    out = image2_observation_path(task_dir, idx)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "observation": str(out), "observation_id": observation_id}, ensure_ascii=False, indent=2))
    return 0


def _hex_from_rgb(rgb: tuple[int, int, int]) -> str:
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def _crop_average_hex(image: object, bbox: list[int]) -> str:
    try:
        from PIL import ImageStat
    except Exception:
        return ""
    x, y, w, h = bbox
    if w <= 0 or h <= 0:
        return ""
    crop = image.crop((x, y, x + w, y + h)).convert("RGB")  # type: ignore[attr-defined]
    stat = ImageStat.Stat(crop.resize((1, 1)))
    rgb = tuple(max(0, min(255, int(round(v)))) for v in stat.mean[:3])
    return _hex_from_rgb(rgb)  # type: ignore[arg-type]


def _bbox_overlap_ratio(a: list[int], b: list[int]) -> float:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix1, iy1 = max(ax, bx), max(ay, by)
    ix2, iy2 = min(ax + aw, bx + bw), min(ay + ah, by + bh)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    inter = (ix2 - ix1) * (iy2 - iy1)
    denom = max(1, min(aw * ah, bw * bh))
    return inter / denom


def _connected_visual_components(reference: Path) -> tuple[list[dict[str, object]], tuple[int, int]]:
    """Extract coarse visual blobs from the reference image using only local pixels.

    This is deliberately deterministic and dependency-light. It does not claim OCR
    accuracy; its job is to stop the post-image stage from inventing module geometry
    before HTML is written.
    """
    Image = _load_pil()
    try:
        from PIL import ImageFilter
    except Exception as exc:
        raise SystemExit(f"Pillow ImageFilter is required for image contract extraction: {exc}")

    with Image.open(reference) as src:
        img = src.convert("RGB")
        src_w, src_h = img.size
        small_w, small_h = 480, 270
        small = img.resize((small_w, small_h))

        corner_points = [
            small.getpixel((0, 0)),
            small.getpixel((small_w - 1, 0)),
            small.getpixel((0, small_h - 1)),
            small.getpixel((small_w - 1, small_h - 1)),
            small.getpixel((small_w // 2, small_h - 1)),
        ]
        bg = tuple(sorted(point[channel] for point in corner_points)[len(corner_points) // 2] for channel in range(3))

        mask = Image.new("L", (small_w, small_h), 0)
        pix = mask.load()
        for y in range(small_h):
            for x in range(small_w):
                r, g, b = small.getpixel((x, y))
                diff = abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2])
                dark_text = (r + g + b) < 680 and max(r, g, b) - min(r, g, b) < 90
                colored = max(r, g, b) - min(r, g, b) > 35
                if diff > 36 or dark_text or colored:
                    pix[x, y] = 255

        # Dilation joins text, rules, and icon strokes into module-scale blobs.
        mask = mask.filter(ImageFilter.MaxFilter(11)).filter(ImageFilter.MinFilter(3))
        data = mask.load()
        seen = bytearray(small_w * small_h)
        components: list[dict[str, object]] = []
        scale_x = src_w / small_w
        scale_y = src_h / small_h

        for start_y in range(small_h):
            for start_x in range(small_w):
                offset = start_y * small_w + start_x
                if seen[offset] or data[start_x, start_y] == 0:
                    continue
                stack = [(start_x, start_y)]
                seen[offset] = 1
                min_x = max_x = start_x
                min_y = max_y = start_y
                count = 0
                while stack:
                    x, y = stack.pop()
                    count += 1
                    min_x, max_x = min(min_x, x), max(max_x, x)
                    min_y, max_y = min(min_y, y), max(max_y, y)
                    for nx, ny in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
                        if nx < 0 or ny < 0 or nx >= small_w or ny >= small_h:
                            continue
                        n_offset = ny * small_w + nx
                        if seen[n_offset] or data[nx, ny] == 0:
                            continue
                        seen[n_offset] = 1
                        stack.append((nx, ny))

                w = max_x - min_x + 1
                h = max_y - min_y + 1
                if count < 18 or w < 5 or h < 4:
                    continue
                bbox = [
                    int(round(min_x * scale_x)),
                    int(round(min_y * scale_y)),
                    int(round(w * scale_x)),
                    int(round(h * scale_y)),
                ]
                if bbox[2] < 24 or bbox[3] < 18:
                    continue
                fill = _crop_average_hex(img, bbox)
                components.append({
                    "bbox": bbox,
                    "area": bbox[2] * bbox[3],
                    "fill": fill,
                    "small_pixel_count": count,
                })

    components.sort(key=lambda item: (item["bbox"][1], item["bbox"][0]))  # type: ignore[index]
    return components, (src_w, src_h)


def _horizontal_color_bands(reference: Path) -> list[dict[str, object]]:
    Image = _load_pil()
    with Image.open(reference) as src:
        img = src.convert("RGB")
        src_w, src_h = img.size
        small_w, small_h = 480, 270
        small = img.resize((small_w, small_h))
        corner_points = [
            small.getpixel((0, 0)),
            small.getpixel((small_w - 1, 0)),
            small.getpixel((0, small_h - 1)),
            small.getpixel((small_w - 1, small_h - 1)),
        ]
        bg = tuple(sorted(point[channel] for point in corner_points)[len(corner_points) // 2] for channel in range(3))
        active_rows: list[int] = []
        for y in range(small_h):
            active = 0
            for x in range(small_w):
                r, g, b = small.getpixel((x, y))
                diff = abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2])
                saturated = max(r, g, b) - min(r, g, b) > 32
                dark = r + g + b < 650
                if diff > 45 and (saturated or dark):
                    active += 1
            if active >= small_w * 0.42:
                active_rows.append(y)

        bands: list[dict[str, object]] = []
        if not active_rows:
            return bands
        start = prev = active_rows[0]
        for y in active_rows[1:] + [-1]:
            if y == prev + 1:
                prev = y
                continue
            height = prev - start + 1
            if height >= 4:
                bbox = [
                    0,
                    int(round(start * src_h / small_h)),
                    1920,
                    int(round(height * 1080 / small_h)),
                ]
                fill = _crop_average_hex(img, [
                    0,
                    int(round(start * src_h / small_h)),
                    src_w,
                    max(1, int(round(height * src_h / small_h))),
                ])
                bands.append({"bbox": bbox, "fill": fill})
            start = prev = y
    return bands


def _region(
    rid: str,
    role: str,
    bbox: list[int],
    rtype: str,
    border_style: str,
    fill: str,
    text: list[str] | None = None,
    icon_semantics: list[str] | None = None,
    confidence: str = "medium",
) -> dict[str, object]:
    return {
        "id": rid,
        "role": role,
        "bbox": bbox,
        "type": rtype,
        "border_style": border_style,
        "fill": fill or "#FFFFFF",
        "text": text or [],
        "icon_semantics": icon_semantics or [],
        "extraction_confidence": confidence,
    }


def extract_visual_contract_regions(reference: Path) -> tuple[list[dict[str, object]], dict[str, object]]:
    components, (img_w, img_h) = _connected_visual_components(reference)
    bands = _horizontal_color_bands(reference)
    sx = 1920 / img_w
    sy = 1080 / img_h

    def scale_bbox(bbox: list[int]) -> list[int]:
        x, y, w, h = bbox
        return [
            int(round(x * sx)),
            int(round(y * sy)),
            int(round(w * sx)),
            int(round(h * sy)),
        ]

    scaled = [
        {**comp, "bbox": scale_bbox(comp["bbox"])}  # type: ignore[index]
        for comp in components
    ]
    warnings: list[str] = []

    wide = [comp for comp in scaled if comp["bbox"][2] >= 900]  # type: ignore[index]
    nav_comp = next((comp for comp in wide if comp["bbox"][1] <= 90 and comp["bbox"][3] <= 120), None)  # type: ignore[index]
    takeaway_comp = next(
        (
            comp for comp in sorted(wide, key=lambda item: item["bbox"][1], reverse=True)  # type: ignore[index]
            if comp["bbox"][1] >= 820 and comp["bbox"][3] <= 130  # type: ignore[index]
        ),
        None,
    )
    source_comp = next(
        (
            comp for comp in sorted(scaled, key=lambda item: item["bbox"][1], reverse=True)  # type: ignore[index]
            if comp["bbox"][1] >= 970 and comp["bbox"][3] <= 70  # type: ignore[index]
        ),
        None,
    )

    nav_band = next((band for band in bands if band["bbox"][1] <= 80 and 18 <= band["bbox"][3] <= 90), None)  # type: ignore[index]
    takeaway_band = next(
        (
            band for band in sorted(bands, key=lambda item: item["bbox"][1], reverse=True)  # type: ignore[index]
            if band["bbox"][1] >= 850 and 18 <= band["bbox"][3] <= 110  # type: ignore[index]
        ),
        None,
    )

    nav_bbox = nav_band["bbox"] if nav_band else (nav_comp["bbox"] if nav_comp else [0, 0, 1920, 42])  # type: ignore[index]
    if nav_comp is None:
        warnings.append("nav region fell back to the standard top-strip location; verify visible nav labels manually")
    takeaway_bbox = takeaway_band["bbox"] if takeaway_band else (takeaway_comp["bbox"] if takeaway_comp else [60, 930, 1800, 48])  # type: ignore[index]
    if takeaway_band is None and takeaway_comp is None:
        warnings.append("takeaway region fell back to the standard bottom-strip location; verify against the reference")

    title_candidates = [
        comp for comp in scaled
        if comp["bbox"][1] >= nav_bbox[1] + nav_bbox[3] and comp["bbox"][1] <= 210  # type: ignore[index]
    ]
    if title_candidates:
        x1 = min(comp["bbox"][0] for comp in title_candidates)  # type: ignore[index]
        y1 = min(comp["bbox"][1] for comp in title_candidates)  # type: ignore[index]
        x2 = max(comp["bbox"][0] + comp["bbox"][2] for comp in title_candidates)  # type: ignore[index]
        y2 = max(comp["bbox"][1] + comp["bbox"][3] for comp in title_candidates)  # type: ignore[index]
        title_bbox = [max(0, x1 - 20), max(nav_bbox[1] + nav_bbox[3], y1 - 20), min(1920, x2 - x1 + 40), min(170, y2 - y1 + 40)]
    else:
        title_bbox = [60, nav_bbox[1] + nav_bbox[3] + 28, 1800, 120]
        warnings.append("title region was not separable from pixels; verify title bbox manually")

    content_top = max(title_bbox[1] + title_bbox[3] + 20, 190)
    content_bottom = min(takeaway_bbox[1] - 18, 930)
    content_components = [
        comp for comp in scaled
        if comp["bbox"][1] >= content_top - 10  # type: ignore[index]
        and comp["bbox"][1] + comp["bbox"][3] <= content_bottom + 10  # type: ignore[index]
        and _bbox_overlap_ratio(comp["bbox"], nav_bbox) < 0.2  # type: ignore[arg-type]
        and _bbox_overlap_ratio(comp["bbox"], takeaway_bbox) < 0.2  # type: ignore[arg-type]
    ]
    if content_components:
        x1 = min(comp["bbox"][0] for comp in content_components)  # type: ignore[index]
        y1 = min(comp["bbox"][1] for comp in content_components)  # type: ignore[index]
        x2 = max(comp["bbox"][0] + comp["bbox"][2] for comp in content_components)  # type: ignore[index]
        y2 = max(comp["bbox"][1] + comp["bbox"][3] for comp in content_components)  # type: ignore[index]
        content_bbox = [max(40, x1 - 24), max(content_top, y1 - 24), min(1840, x2 - x1 + 48), max(80, min(content_bottom - content_top, y2 - y1 + 48))]
    else:
        content_bbox = [60, content_top, 1800, max(80, content_bottom - content_top)]
        warnings.append("content modules were not separable from pixels; generated detail fallback slots")

    detail_components = []
    for comp in content_components:
        bbox = comp["bbox"]  # type: ignore[index]
        if bbox[2] < 80 or bbox[3] < 36:
            continue
        if _bbox_overlap_ratio(bbox, content_bbox) > 0.92 and bbox[2] > content_bbox[2] * 0.8:
            continue
        detail_components.append(comp)

    # Merge near-duplicate nested blobs by keeping the larger one.
    deduped: list[dict[str, object]] = []
    for comp in sorted(detail_components, key=lambda item: item["area"], reverse=True):  # type: ignore[index]
        bbox = comp["bbox"]  # type: ignore[index]
        if any(_bbox_overlap_ratio(bbox, kept["bbox"]) > 0.72 for kept in deduped):  # type: ignore[arg-type,index]
            continue
        deduped.append(comp)
    detail_components = sorted(deduped[:18], key=lambda item: (item["bbox"][1], item["bbox"][0]))  # type: ignore[index]

    if len(detail_components) < MIN_VISUAL_CONTRACT_DETAIL_REGIONS:
        warnings.append(
            f"only {len(detail_components)} separated detail module(s) detected; added grid fallback detail regions that require visual correction"
        )
        detail_components = []
        cols = 2
        rows = 2
        gap_x = 36
        gap_y = 28
        cell_w = int((content_bbox[2] - gap_x * (cols + 1)) / cols)
        cell_h = int((content_bbox[3] - gap_y * (rows + 1)) / rows)
        for row in range(rows):
            for col in range(cols):
                bbox = [
                    content_bbox[0] + gap_x + col * (cell_w + gap_x),
                    content_bbox[1] + gap_y + row * (cell_h + gap_y),
                    max(80, cell_w),
                    max(48, cell_h),
                ]
                detail_components.append({"bbox": bbox, "area": bbox[2] * bbox[3], "fill": "#FFFFFF", "fallback": True})

    regions: list[dict[str, object]] = [
        _region(
            "nav",
            "nav",
            nav_bbox,
            "nav",
            "none",
            str(nav_band.get("fill", nav_comp.get("fill", "#305496") if nav_comp else "#305496")) if nav_band else (str(nav_comp.get("fill", "#305496")) if nav_comp else "#305496"),
            ["visible top navigation text; verify OCR manually"],
            [],
            "high" if nav_band or nav_comp else "low",
        ),
        _region(
            "title",
            "title",
            title_bbox,
            "title",
            "none",
            "#FFFFFF",
            ["visible title text; transcribe in spec before HTML"],
            [],
            "medium" if title_candidates else "low",
        ),
        _region(
            "content",
            "content",
            content_bbox,
            "content",
            "none",
            "#FFFFFF",
            ["content region bounding all detected modules"],
            [],
            "high" if content_components else "low",
        ),
    ]

    for idx, comp in enumerate(detail_components, 1):
        bbox = comp["bbox"]  # type: ignore[index]
        fallback = bool(comp.get("fallback"))
        regions.append(_region(
            f"detail_{idx:02d}",
            "observed_detail_module" if not fallback else "fallback_detail_slot",
            bbox,  # type: ignore[arg-type]
            "card_or_module",
            "solid",
            str(comp.get("fill", "#FFFFFF")),
            ["visible module text requires transcription"] if not fallback else [],
            ["image-extracted visual module"] if not fallback else ["fallback slot after weak pixel separation"],
            "medium" if not fallback else "low",
        ))

    regions.extend([
        _region(
            "takeaway",
            "takeaway",
            takeaway_bbox,
            "takeaway",
            "none",
            str(takeaway_band.get("fill", takeaway_comp.get("fill", "#305496") if takeaway_comp else "#305496")) if takeaway_band else (str(takeaway_comp.get("fill", "#305496")) if takeaway_comp else "#305496"),
            ["visible takeaway text; transcribe in spec before HTML"],
            [],
            "high" if takeaway_band or takeaway_comp else "low",
        ),
        _region(
            "source",
            "source",
            source_comp["bbox"] if source_comp else [60, 1008, 1800, 34],  # type: ignore[index]
            "source",
            "none",
            "#FFFFFF",
            ["visible source line; transcribe in spec before HTML"],
            [],
            "medium" if source_comp else "low",
        ),
    ])

    diagnostics = {
        "image_dimensions": [img_w, img_h],
        "component_count": len(components),
        "horizontal_band_count": len(bands),
        "detail_region_count": len([r for r in regions if not _is_visual_contract_root_region(r)]),
        "warnings": warnings,
        "ocr_status": "not_available_in_local_extractor",
    }
    return regions, diagnostics


def cmd_extract_image2_contract(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    idx = int(args.slide)
    ref = image2_reference_path(task_dir, idx)
    if not ref.exists():
        raise SystemExit(f"Missing selected Image2 reference: {ref}")

    observation, observation_issues = image_observation_record_issues(task_dir, idx)
    if observation_issues:
        print(json.dumps({
            "ok": False,
            "issue": "fresh image observation is required before extracting visual contract",
            "issues": observation_issues,
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    if observation is None:
        raise SystemExit(f"Missing image observation for slide {idx}")

    regions, diagnostics = extract_visual_contract_regions(ref)
    obs_path = image2_observation_path(task_dir, idx)
    contract = {
        "schema": VISUAL_CONTRACT_SCHEMA,
        "measurement_path": str(visual_measurement_path(task_dir, idx).relative_to(task_dir)),
        "reference_path": str(ref.relative_to(task_dir)),
        "observed_from_reference": True,
        "reconstruction_source": IMAGE_ONLY_RECONSTRUCTION_SOURCE,
        "prompt_context_discarded": True,
        "observed_as_fresh_image": True,
        "derivation_method": POST_IMAGE_DERIVATION_METHOD,
        "contract_extraction_method": "local_image_region_segmentation_v1",
        "contract_extraction_source": "image_pixels",
        "observation_record_path": str(obs_path.relative_to(task_dir)),
        "observation_record_sha256": sha256_file(obs_path),
        "observation_id": observation.get("observation_id"),
        "observation_evidence": observation.get("observation_evidence"),
        "regions": regions,
        "extraction_diagnostics": diagnostics,
    }
    measurement = {
        **contract,
        "schema": VISUAL_MEASUREMENT_SCHEMA,
        "contract_path": str(visual_contract_path(task_dir, idx).relative_to(task_dir)),
        "canvas": {"width": 1920, "height": 1080},
        "color_samples": {
            "primary_nav_or_takeaway": next(
                (str(region.get("fill", "")) for region in regions if region.get("id") in {"nav", "takeaway"}),
                "",
            ),
            "background": "#FFFFFF",
        },
        "font_size_estimates": {
            "nav": "estimate from visible nav band; verify manually",
            "title": "estimate from visible title hierarchy; verify manually",
            "body": "estimate from module text density; verify manually",
            "source": "estimate from source line density; verify manually",
        },
        "text_transcription_status": "manual_transcription_required_from_reference_image",
    }
    out = visual_contract_path(task_dir, idx)
    measurement_out = visual_measurement_path(task_dir, idx)
    if out.exists() and not args.force:
        print(json.dumps({
            "ok": False,
            "issue": f"{out.relative_to(task_dir)} already exists; pass --force to overwrite",
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    if measurement_out.exists() and not args.force:
        print(json.dumps({
            "ok": False,
            "issue": f"{measurement_out.relative_to(task_dir)} already exists; pass --force to overwrite",
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(contract, ensure_ascii=False, indent=2), encoding="utf-8")
    measurement_out.write_text(json.dumps(measurement, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "ok": True,
        "contract": str(out),
        "measurement": str(measurement_out),
        "regions": len(regions),
        "detail_regions": diagnostics["detail_region_count"],
        "warnings": diagnostics["warnings"],
    }, ensure_ascii=False, indent=2))
    return 0


def cmd_register_image2_reference(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    idx = int(args.slide)
    source_kind = str(args.source or "").strip()
    tool_call_id = str(args.tool_call_id or "").strip()
    if source_kind not in IMAGE2_ALLOWED_SOURCE_KINDS:
        print(
            json.dumps({
                "ok": False,
                "issue": "registration source must be an explicit image-generation source",
                "allowed_sources": sorted(IMAGE2_ALLOWED_SOURCE_KINDS),
                "supplied_source": source_kind or "<missing>",
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    if len(tool_call_id) < 8:
        print(
            json.dumps({
                "ok": False,
                "issue": "tool-call/provenance id is required; do not register local previews as Image2 references",
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    prompt_path = image2_prompt_path(task_dir, idx)
    if not prompt_path.exists():
        print(
            json.dumps({
                "ok": False,
                "issue": f"image2_prompt_{idx:02d}.md missing; run prepare-image2-prompts first",
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1

    prompt_sha = sha256_text(read_text(prompt_path))
    supplied_sha = str(args.generated_prompt_sha256).strip().lower()
    if supplied_sha != prompt_sha:
        print(
            json.dumps({
                "ok": False,
                "issue": "generated prompt sha does not match the locked image2_prompt file",
                "expected_image2_prompt_sha256": prompt_sha,
                "supplied_generated_prompt_sha256": supplied_sha,
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1

    request_path = Path(args.generation_request).resolve()
    expected_request_path = image2_generation_request_path(task_dir, idx).resolve()
    if request_path != expected_request_path:
        print(
            json.dumps({
                "ok": False,
                "issue": "generation request path must be the locked per-slide request produced by prepare-image2-prompts",
                "expected_generation_request": str(expected_request_path),
                "supplied_generation_request": str(request_path),
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    generation_request = load_generation_request(request_path)
    if generation_request is None:
        print(
            json.dumps({
                "ok": False,
                "issue": "generation request is missing or cannot be parsed",
                "generation_request": str(request_path),
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    request_prompt_text = str(generation_request.get("prompt_text", ""))
    request_issues: list[str] = []
    if generation_request.get("schema") != IMAGE2_GENERATION_REQUEST_SCHEMA:
        request_issues.append(f"schema must be {IMAGE2_GENERATION_REQUEST_SCHEMA}")
    if generation_request.get("slide") != idx:
        request_issues.append("slide mismatch")
    if generation_request.get("prompt_source_path") != f"image2/image2_prompt_{idx:02d}.md":
        request_issues.append("prompt_source_path mismatch")
    if generation_request.get("prompt_sha256") != prompt_sha:
        request_issues.append("prompt_sha256 mismatch")
    if sha256_text(request_prompt_text) != prompt_sha:
        request_issues.append("prompt_text does not exactly match the locked image2_prompt file")
    if generation_request.get("prompt_transfer_method") != IMAGE2_PROMPT_TRANSFER_METHOD:
        request_issues.append(f"prompt_transfer_method must be {IMAGE2_PROMPT_TRANSFER_METHOD}")
    if generation_request.get("manual_prompt_rewrite_allowed") is not False:
        request_issues.append("manual_prompt_rewrite_allowed must be false")
    if request_issues:
        print(
            json.dumps({
                "ok": False,
                "issue": "generation request does not match the locked prompt handoff",
                "request_issues": request_issues,
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1

    source_image = Path(args.image).resolve()
    target = image2_reference_path(task_dir, idx)
    if source_image == target.resolve():
        print(
            json.dumps({
                "ok": False,
                "issue": "register from the original image-generation output path, not the final image2_reference target",
                "target": str(target),
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    try:
        source_image.relative_to(task_dir)
        print(
            json.dumps({
                "ok": False,
                "issue": "register source must be outside the task directory; generated image will be copied into image2/",
                "image": str(source_image),
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    except ValueError:
        pass
    suspicious_source = looks_like_rendered_preview_image(source_image)
    try:
        rel_source = source_image.relative_to(task_dir)
        if path_contains_any(rel_source, IMAGE2_FORBIDDEN_SOURCE_PARTS) or name_contains_any(rel_source, IMAGE2_FORBIDDEN_SOURCE_NAME_PARTS):
            suspicious_source = True
    except ValueError:
        pass
    if suspicious_source:
        print(
            json.dumps({
                "ok": False,
                "issue": "image2 reference source looks like an HTML/PPTX/QA preview, not a generated visual reference",
                "image": str(source_image),
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    if not source_image.exists() or not source_image.is_file():
        print(json.dumps({"ok": False, "issue": f"image file missing: {source_image}"}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    widescreen_ok, widescreen_detail = is_image2_widescreen(source_image)
    if not widescreen_ok:
        print(
            json.dumps({
                "ok": False,
                "issue": f"image2 reference rejected: {widescreen_detail}",
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    target.parent.mkdir(parents=True, exist_ok=True)
    if source_image.resolve() != target.resolve():
        shutil.copy2(source_image, target)

    manifest = load_image2_manifest(task_dir)
    if manifest.get("schema") != IMAGE2_MANIFEST_SCHEMA:
        print(
            json.dumps({
                "ok": False,
                "issue": "image2_generation_manifest.json missing or stale; run prepare-image2-prompts first",
            }, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1
    entries = manifest_entries_by_slide(manifest)
    entry = entries.get(idx)
    if not entry:
        print(json.dumps({"ok": False, "issue": f"manifest entry for slide {idx} missing"}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1

    final_prompt = task_dir / "analysis" / f"final_prompt_{idx:02d}.md"
    entry.update({
        "status": "registered_verified",
        "generated_from_image2_prompt_file": True,
        "verification_method": IMAGE2_VERIFICATION_METHOD,
        "prompt_transfer_method": IMAGE2_PROMPT_TRANSFER_METHOD,
        "manual_short_prompt_allowed": False,
        "image2_prompt_sha256": prompt_sha,
        "generation_request_path": str(request_path.relative_to(task_dir)),
        "generation_request_sha256": sha256_file(request_path),
        "generation_request_id": str(generation_request.get("request_id", "")),
        "final_prompt_sha256": sha256_file(final_prompt) if final_prompt.exists() else None,
        "system_prompt_sha256": sha256_file(SYSTEM_PROMPT),
        "image_sha256": sha256_file(target),
        "image_path": str(target.relative_to(task_dir)),
        "image_status": "registered_verified",
        "registered_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "registered_source": str(source_image),
        "registration_source_kind": source_kind,
        "tool_call_id": tool_call_id,
    })
    manifest["slides"] = [entries.get(i, {}) for i in range(1, int(manifest.get("expected_pages", idx)) + 1)]
    image2_manifest_path(task_dir).write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps({"ok": True, "registered": entry}, ensure_ascii=False, indent=2))
    return 0


def check_analysis_files(task_dir: Path, expected: int, issues: list[str]) -> None:
    analysis = task_dir / "analysis" / "analysis_report.md"
    audit = task_dir / "analysis" / "prompt_selection_audit.md"

    analysis_text = read_text(analysis)
    if not analysis.exists():
        issues.append("analysis_report.md missing")
    elif len(analysis_text.strip()) < 1200:
        issues.append("analysis_report.md too short; looks like a placeholder")
    elif has_placeholder(analysis_text):
        issues.append("analysis_report.md contains placeholder language")
    else:
        analysis_lower = analysis_text.lower()
        required_sections = {
            "source inventory": ["source inventory", "资料清单", "来源清单"],
            "conflict resolution": ["conflict", "冲突", "矛盾", "分歧"],
            "codex judgment": ["codex judgment", "independent judgment", "独立判断", "综合判断"],
            "cross validation": ["cross validation", "cross-check", "交叉验证", "外部验证"],
            "limits": ["known limits", "limits", "局限", "限制"],
        }
        for section, needles in required_sections.items():
            if not any(needle in analysis_lower for needle in needles):
                issues.append(
                    f"analysis_report.md missing {section}; analyze source conflicts, Codex judgment, and cross-validation before slide prompts"
                )

    audit_text = read_text(audit)
    if not audit.exists():
        issues.append("prompt_selection_audit.md missing")
    elif len(audit_text.strip()) < 500:
        issues.append("prompt_selection_audit.md too short; must justify layout choices")
    elif has_placeholder(audit_text):
        issues.append("prompt_selection_audit.md contains placeholder language")

    prompt_selections: list[tuple[int, str, str]] = []
    for idx in range(1, expected + 1):
        prompt = task_dir / "analysis" / f"final_prompt_{idx:02d}.md"
        text = read_text(prompt)
        if not prompt.exists():
            issues.append(f"final_prompt_{idx:02d}.md missing")
            continue
        template_name = selected_prompt_template(text)
        if template_name:
            prompt_selections.append((idx, template_name, prompt_scaffold_family(template_name)))
        if len(text.strip()) < 800:
            issues.append(f"final_prompt_{idx:02d}.md too short; must be a complete filled prompt")
        if has_placeholder(text):
            issues.append(f"final_prompt_{idx:02d}.md contains placeholder language")
        for pattern in PROMPT_FORBIDDEN_PATTERNS:
            if pattern in text:
                issues.append(
                    f"final_prompt_{idx:02d}.md contains forbidden legacy style phrase: {pattern}"
                )
        if "LAYOUT" not in text and "LAYOUT_NAME" not in text:
            issues.append(f"final_prompt_{idx:02d}.md missing LAYOUT/LAYOUT_NAME")
        issues.extend(prompt_template_issue(text, f"final_prompt_{idx:02d}.md"))
        for marker in PROMPT_REQUIRED_MARKERS:
            if marker not in text:
                issues.append(f"final_prompt_{idx:02d}.md missing {marker}")
        if "VISUAL STYLE" not in text and "compact consulting" not in text.lower():
            issues.append(f"final_prompt_{idx:02d}.md missing compact consulting visual-style instruction")
    issues.extend(prompt_selection_diversity_issues(prompt_selections))
    issues.extend(prompt_selection_plan_issues(task_dir, expected, prompt_selections))


def check_image2_style_review(task_dir: Path, expected: int, issues: list[str]) -> None:
    review_path = image2_style_review_path(task_dir)
    if not review_path.exists():
        issues.append(
            "qa/image2_style_review.json missing; open each Image2 reference and record style evidence before HTML"
        )
        return
    try:
        data = json.loads(review_path.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"qa/image2_style_review.json cannot be parsed: {exc}")
        return

    if not isinstance(data, dict):
        issues.append("qa/image2_style_review.json must be a JSON object")
        return
    if data.get("actual_images_opened") is not True:
        issues.append("qa/image2_style_review.json actual_images_opened must be true")
    if data.get("style_references_compared") is not True:
        issues.append("qa/image2_style_review.json style_references_compared must be true")
    for idx in range(1, expected + 1):
        if not file_is_after_reference(task_dir, idx, review_path):
            issues.append(
                f"qa/image2_style_review.json is older than image2_reference_{idx:02d}.png; "
                "style review must be recorded after opening the selected Image2 references"
            )

    slides = data.get("slides")
    if not isinstance(slides, list) or len(slides) != expected:
        found = len(slides) if isinstance(slides, list) else 0
        issues.append(f"qa/image2_style_review.json slides: expected {expected}, found {found}")
        return

    seen: set[int] = set()
    required_checks = set(IMAGE2_STYLE_REVIEW_MIN_CHECKS)
    for pos, entry in enumerate(slides, 1):
        if not isinstance(entry, dict):
            issues.append(f"image2 style review slide entry {pos}: must be an object")
            continue
        slide = entry.get("slide")
        if not isinstance(slide, int):
            issues.append(f"image2 style review entry {pos}: slide must be an integer")
            continue
        seen.add(slide)
        if slide < 1 or slide > expected:
            issues.append(f"image2 style review slide {slide}: out of range for expected page count")
            continue
        if entry.get("status") != "pass":
            issues.append(f"image2 style review slide {slide}: status must be pass")
        if entry.get("compared") is not True:
            issues.append(f"image2 style review slide {slide}: compared must be true")

        ref = _normalize_ref_path(task_dir, entry.get("reference_path"))
        expected_ref = image2_reference_path(task_dir, slide).resolve()
        if ref is None or not ref.exists():
            issues.append(f"image2 style review slide {slide}: reference_path missing or does not exist")
        elif ref.resolve() != expected_ref:
            issues.append(
                f"image2 style review slide {slide}: reference_path must point to image2_reference_{slide:02d}.png"
            )

        dims = str(entry.get("image_dimensions", "")).strip()
        if ref is not None and ref.exists():
            actual_dims = image_dimensions(ref)
            if actual_dims is not None:
                actual_dims_text = f"{actual_dims[0]}x{actual_dims[1]}"
                if dims != actual_dims_text:
                    issues.append(
                        f"image2 style review slide {slide}: image_dimensions must be {actual_dims_text}, found {dims or '<missing>'}"
                    )
        elif not dims:
            issues.append(f"image2 style review slide {slide}: image_dimensions missing")

        evidence = str(entry.get("evidence", "")).strip()
        if len(evidence) < 120:
            issues.append(
                f"image2 style review slide {slide}: evidence must be a concrete observed comparison of at least 120 characters"
            )

        checks = entry.get("dimensions_checked")
        check_set = set(checks) if isinstance(checks, list) and all(isinstance(v, str) for v in checks) else set()
        missing = sorted(required_checks - check_set)
        if missing:
            issues.append(
                f"image2 style review slide {slide}: missing dimensions_checked values: {', '.join(missing)}"
            )
        reject_reasons = entry.get("reject_reasons")
        if reject_reasons not in ([], None):
            issues.append(f"image2 style review slide {slide}: reject_reasons must be empty for a passing reference")

    missing_slides = sorted(set(range(1, expected + 1)) - seen)
    if missing_slides:
        issues.append(f"qa/image2_style_review.json missing slide reviews: {', '.join(map(str, missing_slides))}")


def check_image2_user_review(task_dir: Path, expected: int, issues: list[str]) -> None:
    review_path = image2_user_review_path(task_dir)
    if not review_path.exists():
        issues.append(
            "qa/image2_user_review.json missing; pause after image generation, show the selected page previews to the user, "
            "record explicit approval or feedback, and do not write HTML before this gate passes"
        )
        return
    try:
        data = json.loads(review_path.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"qa/image2_user_review.json cannot be parsed: {exc}")
        return
    if not isinstance(data, dict):
        issues.append("qa/image2_user_review.json must be a JSON object")
        return
    if data.get("schema") != IMAGE2_USER_REVIEW_SCHEMA:
        issues.append(f"qa/image2_user_review.json schema must be {IMAGE2_USER_REVIEW_SCHEMA}")
    if data.get("user_approved") is not True:
        issues.append("qa/image2_user_review.json user_approved must be true before HTML/spec authoring")
    if data.get("reviewed_after_image_generation") is not True:
        issues.append("qa/image2_user_review.json reviewed_after_image_generation must be true")
    feedback = str(data.get("user_feedback", "")).strip()
    if len(feedback) < 2:
        issues.append("qa/image2_user_review.json user_feedback must record the user's approval or requested changes")

    slides = data.get("slides")
    if not isinstance(slides, list) or len(slides) != expected:
        found = len(slides) if isinstance(slides, list) else 0
        issues.append(f"qa/image2_user_review.json slides: expected {expected}, found {found}")
        return

    seen: set[int] = set()
    for pos, entry in enumerate(slides, 1):
        if not isinstance(entry, dict):
            issues.append(f"user image review slide entry {pos}: must be an object")
            continue
        slide = entry.get("slide")
        if not isinstance(slide, int):
            issues.append(f"user image review entry {pos}: slide must be an integer")
            continue
        seen.add(slide)
        if slide < 1 or slide > expected:
            issues.append(f"user image review slide {slide}: out of range for expected page count")
            continue
        if entry.get("status") != "approved":
            issues.append(f"user image review slide {slide}: status must be approved; regenerate before HTML if changes are requested")
        ref = _normalize_ref_path(task_dir, entry.get("reference_path"))
        expected_ref = image2_reference_path(task_dir, slide).resolve()
        if ref is None or not ref.exists():
            issues.append(f"user image review slide {slide}: reference_path missing or does not exist")
            continue
        if ref.resolve() != expected_ref:
            issues.append(f"user image review slide {slide}: reference_path must point to image2_reference_{slide:02d}.png")
        expected_sha = sha256_file(expected_ref)
        if entry.get("image_sha256") != expected_sha:
            issues.append(
                f"user image review slide {slide}: image_sha256 must match the current selected reference; "
                "rerun user review after regenerating or replacing an image"
            )
        if not file_is_after_reference(task_dir, slide, review_path):
            issues.append(
                f"qa/image2_user_review.json is older than image2_reference_{slide:02d}.png; "
                "user review must happen after the current selected reference is generated"
            )
    missing_slides = sorted(set(range(1, expected + 1)) - seen)
    if missing_slides:
        issues.append(f"qa/image2_user_review.json missing slide reviews: {', '.join(map(str, missing_slides))}")


def check_image2_files(
    task_dir: Path,
    expected: int,
    issues: list[str],
    *,
    require_prompt_files: bool = True,
) -> int:
    image_refs = sorted((task_dir / "image2").glob("image2_reference_*.png"))
    if len(image_refs) != expected:
        issues.append(f"selected Image2 references: expected {expected}, found {len(image_refs)}")
    manifest_path = image2_manifest_path(task_dir)
    manifest: dict[str, object] | None = None
    if not manifest_path.exists():
        issues.append(
            "image2_generation_manifest.json missing; run prepare-image2-prompts and generate each image from image2_prompt_XX.md"
        )
    else:
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception as exc:
            issues.append(f"image2_generation_manifest.json cannot be parsed: {exc}")
            manifest = None
        if isinstance(manifest, dict):
            if manifest.get("schema") != IMAGE2_MANIFEST_SCHEMA:
                issues.append(
                    f"image2_generation_manifest.json has wrong schema; expected {IMAGE2_MANIFEST_SCHEMA}"
                )
            slides = manifest.get("slides")
            if not isinstance(slides, list) or len(slides) != expected:
                found = len(slides) if isinstance(slides, list) else 0
                issues.append(f"image2_generation_manifest slides: expected {expected}, found {found}")
    for idx in range(1, expected + 1):
        ref = image2_reference_path(task_dir, idx)
        if not ref.exists():
            continue
        if ref.stat().st_size < 10_000:
            issues.append(f"image2_reference_{idx:02d}.png is suspiciously small")
        widescreen_ok, widescreen_detail = is_image2_widescreen(ref)
        if not widescreen_ok:
            issues.append(f"image2_reference_{idx:02d}.png {widescreen_detail}")
        prompt_path = image2_prompt_path(task_dir, idx)
        if require_prompt_files and not prompt_path.exists():
            issues.append(f"image2_prompt_{idx:02d}.md missing; Image2 must be generated from the locked prompt file")
        elif prompt_path.exists() and require_prompt_files:
            prompt_text = read_text(prompt_path)
            expected_prompt = build_image2_prompt_text(task_dir, idx)
            if sha256_text(prompt_text) != sha256_text(expected_prompt):
                issues.append(
                    f"image2_prompt_{idx:02d}.md does not match SYSTEM_PROMPT + final_prompt_{idx:02d}.md + locked Image2 tail"
                )
            if "IMAGE2 GENERATION LOCK" not in prompt_text:
                issues.append(f"image2_prompt_{idx:02d}.md missing IMAGE2 GENERATION LOCK")

        if isinstance(manifest, dict) and isinstance(manifest.get("slides"), list):
            entry = manifest["slides"][idx - 1] if idx - 1 < len(manifest["slides"]) else {}
            if not isinstance(entry, dict):
                issues.append(f"image2_generation_manifest slide {idx}: entry must be an object")
                continue
            if entry.get("manual_short_prompt_allowed") is not False:
                issues.append(f"image2_generation_manifest slide {idx}: manual_short_prompt_allowed must be false")
            if entry.get("status") != "registered_verified":
                issues.append(
                    f"image2_generation_manifest slide {idx}: status must be registered_verified after register-image2-reference"
                )
            if entry.get("generated_from_image2_prompt_file") is not True:
                issues.append(f"image2_generation_manifest slide {idx}: generated_from_image2_prompt_file must be true")
            if entry.get("verification_method") != IMAGE2_VERIFICATION_METHOD:
                issues.append(
                    f"image2_generation_manifest slide {idx}: verification_method must be {IMAGE2_VERIFICATION_METHOD}"
                )
            issues.extend(image2_generation_request_issues(task_dir, idx, entry))
            issues.extend(image2_reference_provenance_issues(task_dir, idx, entry))
            if entry.get("image2_prompt_path") != f"image2/image2_prompt_{idx:02d}.md":
                issues.append(f"image2_generation_manifest slide {idx}: image2_prompt_path mismatch")
            if entry.get("image_path") != f"image2/image2_reference_{idx:02d}.png":
                issues.append(f"image2_generation_manifest slide {idx}: image_path mismatch")
            if prompt_path.exists() and entry.get("image2_prompt_sha256") != sha256_text(read_text(prompt_path)):
                issues.append(f"image2_generation_manifest slide {idx}: image2_prompt_sha256 mismatch")
            elif not prompt_path.exists() and not isinstance(entry.get("image2_prompt_sha256"), str):
                issues.append(f"image2_generation_manifest slide {idx}: image2_prompt_sha256 missing")
            if ref.exists() and entry.get("image_sha256") != sha256_file(ref):
                issues.append(f"image2_generation_manifest slide {idx}: image_sha256 mismatch")
            final_prompt = task_dir / "analysis" / f"final_prompt_{idx:02d}.md"
            if require_prompt_files and final_prompt.exists() and entry.get("final_prompt_sha256") != sha256_file(final_prompt):
                issues.append(f"image2_generation_manifest slide {idx}: final_prompt_sha256 mismatch")
            if require_prompt_files and SYSTEM_PROMPT.exists() and entry.get("system_prompt_sha256") != sha256_file(SYSTEM_PROMPT):
                issues.append(f"image2_generation_manifest slide {idx}: system_prompt_sha256 mismatch")
    if image_refs:
        check_image2_style_review(task_dir, expected, issues)
        check_image2_user_review(task_dir, expected, issues)
    return len(image_refs)


def check_spec_files(task_dir: Path, expected: int, issues: list[str]) -> int:
    spec_refs = sorted((task_dir / "spec").glob("slide*_spec.md"))
    if len(spec_refs) != expected:
        issues.append(f"spec files: expected {expected}, found {len(spec_refs)}")
    issues.extend(post_image_memory_boundary_issues(task_dir, expected))
    required = [
        "canvas",
        "element",
        "layout",
        "icon",
        "risk",
        "checklist",
        "reference",
        "html",
    ]
    for idx in range(1, expected + 1):
        spec = task_dir / "spec" / f"slide{idx:02d}_spec.md"
        raw_text = read_text(spec)
        text = raw_text.lower()
        if not spec.exists():
            continue
        observation, observation_issues = image_observation_record_issues(task_dir, idx)
        issues.extend(observation_issues)
        if len(text.strip()) < 500:
            issues.append(f"slide{idx:02d}_spec.md too short")
        if not file_is_after_reference(task_dir, idx, spec):
            issues.append(
                f"slide{idx:02d}_spec.md is older than image2_reference_{idx:02d}.png; "
                "spec must be written after observing the selected Image2 reference"
            )
        if not file_is_after_memory_boundary(task_dir, spec):
            issues.append(
                f"slide{idx:02d}_spec.md is older than qa/post_image_memory_boundary.json; "
                "rewrite spec after the forced post-image memory reset"
            )
        for marker in required:
            if marker not in text:
                issues.append(f"slide{idx:02d}_spec.md missing {marker} section/detail")
        if IMAGE_ONLY_RECONSTRUCTION_SOURCE not in text:
            issues.append(
                f"slide{idx:02d}_spec.md missing reconstruction boundary "
                f"'{IMAGE_ONLY_RECONSTRUCTION_SOURCE}'; after Image2 approval, specs must be written from the image as a fresh artifact"
            )
        if "prompt_context_discarded: true" not in text and "prompt context discarded: true" not in text:
            issues.append(
                f"slide{idx:02d}_spec.md must declare prompt_context_discarded: true; "
                "post-image spec authoring must not use final prompts, prompt-library memory, or analysis notes as a design source"
            )
        expected_observation_ref = f"slide{idx:02d}_image_observation.json"
        if expected_observation_ref not in raw_text:
            issues.append(
                f"slide{idx:02d}_spec.md must reference {expected_observation_ref}; "
                "post-image spec must derive from the fresh visual observation record"
            )
        if observation is not None:
            observation_id = str(observation.get("observation_id", "")).strip()
            if observation_id and observation_id not in raw_text:
                issues.append(
                    f"slide{idx:02d}_spec.md must include observation_id {observation_id}; "
                    "spec and visual contract must bind to the same observed image record"
                )
        forbidden = post_image_memory_markers(raw_text)
        if forbidden:
            issues.append(
                f"slide{idx:02d}_spec.md contains upstream prompt/analysis memory markers after Image2 approval: "
                + ", ".join(forbidden)
            )
    return len(spec_refs)


def visual_contract_path(task_dir: Path, idx: int) -> Path:
    return task_dir / "spec" / f"slide{idx:02d}_visual_contract.json"


def visual_measurement_path(task_dir: Path, idx: int) -> Path:
    return task_dir / "spec" / f"slide{idx:02d}_visual_measurement.json"


def _visual_region_ids(data: dict[str, object]) -> list[str]:
    regions = data.get("regions")
    if not isinstance(regions, list):
        return []
    ids: list[str] = []
    for region in regions:
        if isinstance(region, dict):
            rid = str(region.get("id", "")).strip()
            if rid:
                ids.append(rid)
    return ids


def _visual_measurement_quality_issues(label: str, data: dict[str, object]) -> list[str]:
    issues: list[str] = []
    regions = data.get("regions")
    if not isinstance(regions, list):
        return issues
    diagnostics = data.get("extraction_diagnostics")
    warnings: list[str] = []
    if isinstance(diagnostics, dict):
        raw_warnings = diagnostics.get("warnings")
        if isinstance(raw_warnings, list):
            warnings = [str(item).lower() for item in raw_warnings]
        if diagnostics.get("ocr_status") in {"not_available", "not_available_in_local_extractor"}:
            issues.append(
                f"{label} OCR/text extraction is unavailable; do not call this a finished visual measurement"
            )
        detail_count = diagnostics.get("detail_region_count")
        if isinstance(detail_count, int) and detail_count < MIN_VISUAL_CONTRACT_DETAIL_REGIONS:
            issues.append(
                f"{label} extracted only {detail_count} detail region(s); visual measurement must identify real visible modules"
            )
    if any("fallback" in warning or "not separable" in warning for warning in warnings):
        issues.append(
            f"{label} used fallback or non-separable geometry; stop and create real image-derived bboxes before HTML/PPTX"
        )
    for pos, region in enumerate(regions, 1):
        if not isinstance(region, dict):
            continue
        rid = str(region.get("id", pos)).strip()
        role = str(region.get("role", "")).lower()
        rtype = str(region.get("type", "")).lower()
        text_items = _region_text_items(region)
        icon_items = _region_icon_items(region)
        placeholders = [
            item for item in text_items + icon_items
            if any(token in item.lower() for token in [
                "fallback",
                "requires transcription",
                "verify manually",
                "visible module text",
                "visible content module reconstructed",
            ])
        ]
        if "fallback" in role or "fallback" in rtype:
            issues.append(f"{label} region {rid} is a fallback slot, not measured geometry")
        if placeholders:
            issues.append(f"{label} region {rid} still contains placeholder measurement text")
    return issues


def visual_measurement_issues(
    task_dir: Path,
    idx: int,
    contract_data: dict[str, object],
    observation: dict[str, object] | None,
) -> list[str]:
    issues: list[str] = []
    measurement = visual_measurement_path(task_dir, idx)
    if not measurement.exists():
        return [
            f"slide{idx:02d}_visual_measurement.json missing; before HTML, create a visible measurement file "
            "from the selected Image2 reference and use it as the only reconstruction source"
        ]
    try:
        raw = measurement.read_text(encoding="utf-8")
        data = json.loads(raw)
    except Exception as exc:
        return [f"slide{idx:02d}_visual_measurement.json cannot be parsed: {exc}"]
    if not isinstance(data, dict):
        return [f"slide{idx:02d}_visual_measurement.json must be a JSON object"]
    if data.get("schema") != VISUAL_MEASUREMENT_SCHEMA:
        issues.append(f"slide{idx:02d}_visual_measurement.json schema must be {VISUAL_MEASUREMENT_SCHEMA}")
    if data.get("reference_path") != f"image2/image2_reference_{idx:02d}.png":
        issues.append(f"slide{idx:02d}_visual_measurement.json reference_path mismatch")
    if data.get("reconstruction_source") != IMAGE_ONLY_RECONSTRUCTION_SOURCE:
        issues.append(
            f"slide{idx:02d}_visual_measurement.json reconstruction_source must be {IMAGE_ONLY_RECONSTRUCTION_SOURCE}"
        )
    if data.get("prompt_context_discarded") is not True:
        issues.append(f"slide{idx:02d}_visual_measurement.json prompt_context_discarded must be true")
    if data.get("observed_as_fresh_image") is not True:
        issues.append(f"slide{idx:02d}_visual_measurement.json observed_as_fresh_image must be true")
    if data.get("contract_extraction_source") != "image_pixels":
        issues.append(f"slide{idx:02d}_visual_measurement.json contract_extraction_source must be image_pixels")
    if observation is not None and data.get("observation_id") != observation.get("observation_id"):
        issues.append(f"slide{idx:02d}_visual_measurement.json observation_id must match the fresh image observation record")
    if _visual_region_ids(data) != _visual_region_ids(contract_data):
        issues.append(
            f"slide{idx:02d}_visual_measurement.json region ids must exactly match slide{idx:02d}_visual_contract.json"
        )
    issues.extend(_visual_measurement_quality_issues(f"slide{idx:02d}_visual_measurement.json", data))
    required_measurement_keys = [
        "canvas",
        "regions",
        "color_samples",
        "font_size_estimates",
        "text_transcription_status",
    ]
    for key in required_measurement_keys:
        if key not in data:
            issues.append(f"slide{idx:02d}_visual_measurement.json missing {key}")
    forbidden = post_image_memory_markers(raw)
    if forbidden:
        issues.append(
            f"slide{idx:02d}_visual_measurement.json contains upstream prompt/analysis memory markers after Image2 approval: "
            + ", ".join(forbidden)
        )
    if not file_is_after_reference(task_dir, idx, measurement):
        issues.append(
            f"slide{idx:02d}_visual_measurement.json is older than image2_reference_{idx:02d}.png; "
            "recreate measurement after observing the selected image"
        )
    if not file_is_after_memory_boundary(task_dir, measurement):
        issues.append(
            f"slide{idx:02d}_visual_measurement.json is older than qa/post_image_memory_boundary.json; "
            "recreate measurement after the forced post-image memory reset"
        )
    return issues


def _normalize_ref_path(task_dir: Path, raw: object) -> Path | None:
    if not isinstance(raw, str) or not raw.strip():
        return None
    path = Path(raw)
    if not path.is_absolute():
        path = task_dir / path
    return path


def _region_ids_from_html(text: str) -> set[str]:
    ids: set[str] = set()
    for match in re.finditer(
        r"data-ref-id\s*=\s*['\"]([^'\"]+)['\"]",
        text,
        flags=re.IGNORECASE,
    ):
        value = match.group(1).strip()
        if value:
            ids.add(value)
    return ids


def post_image_memory_markers(text: str) -> list[str]:
    """Find prompt/analysis memory references that are forbidden after Image2 approval."""
    lowered = text.lower()
    return sorted({marker for marker in POST_IMAGE_FORBIDDEN_MEMORY_MARKERS if marker in lowered})


def image_observation_record_issues(task_dir: Path, idx: int) -> tuple[dict[str, object] | None, list[str]]:
    issues: list[str] = []
    path = image2_observation_path(task_dir, idx)
    if not path.exists():
        return None, [f"slide{idx:02d}_image_observation.json missing; record fresh image observation before spec/HTML"]
    expected = expected_pages_from_task(task_dir)
    if expected:
        issues.extend(post_image_memory_boundary_issues(task_dir, expected))
    try:
        raw = path.read_text(encoding="utf-8")
        data = json.loads(raw)
    except Exception as exc:
        return None, [f"slide{idx:02d}_image_observation.json cannot be parsed: {exc}"]
    if not isinstance(data, dict):
        return None, [f"slide{idx:02d}_image_observation.json must be a JSON object"]
    ref = image2_reference_path(task_dir, idx)
    if data.get("schema") != IMAGE2_OBSERVATION_SCHEMA:
        issues.append(f"slide{idx:02d}_image_observation.json schema must be {IMAGE2_OBSERVATION_SCHEMA}")
    if data.get("slide") != idx:
        issues.append(f"slide{idx:02d}_image_observation.json slide mismatch")
    if data.get("reference_path") != f"image2/image2_reference_{idx:02d}.png":
        issues.append(f"slide{idx:02d}_image_observation.json reference_path mismatch")
    if ref.exists() and data.get("reference_sha256") != sha256_file(ref):
        issues.append(f"slide{idx:02d}_image_observation.json reference_sha256 mismatch")
    if data.get("observed_from_reference") is not True:
        issues.append(f"slide{idx:02d}_image_observation.json observed_from_reference must be true")
    if data.get("reconstruction_source") != IMAGE_ONLY_RECONSTRUCTION_SOURCE:
        issues.append(f"slide{idx:02d}_image_observation.json reconstruction_source must be {IMAGE_ONLY_RECONSTRUCTION_SOURCE}")
    if data.get("prompt_context_discarded") is not True:
        issues.append(f"slide{idx:02d}_image_observation.json prompt_context_discarded must be true")
    if data.get("observed_as_fresh_image") is not True:
        issues.append(f"slide{idx:02d}_image_observation.json observed_as_fresh_image must be true")
    if data.get("derivation_method") != POST_IMAGE_DERIVATION_METHOD:
        issues.append(f"slide{idx:02d}_image_observation.json derivation_method must be {POST_IMAGE_DERIVATION_METHOD}")
    evidence = str(data.get("observation_evidence", "")).strip()
    if len(evidence) < 120:
        issues.append(f"slide{idx:02d}_image_observation.json observation_evidence must be at least 120 characters")
    forbidden = post_image_memory_markers(raw)
    if forbidden:
        issues.append(
            f"slide{idx:02d}_image_observation.json contains upstream prompt/analysis memory markers: "
            + ", ".join(forbidden)
        )
    if not file_is_after_reference(task_dir, idx, path):
        issues.append(
            f"slide{idx:02d}_image_observation.json is older than image2_reference_{idx:02d}.png; "
            "record observation after reopening the selected reference"
        )
    if not file_is_after_memory_boundary(task_dir, path):
        issues.append(
            f"slide{idx:02d}_image_observation.json is older than qa/post_image_memory_boundary.json; "
            "record observation after the forced post-image memory reset"
        )
    return data, issues


def _large_visual_class_count(text: str) -> int:
    """Count common visual containers that should be bound to visual contract ids."""
    count = 0
    for match in re.finditer(r"<(?P<tag>div|section|aside|main|nav)\b(?P<attrs>[^>]*)>", text, flags=re.IGNORECASE):
        attrs = match.group("attrs")
        if "data-ref-id" in attrs.lower():
            continue
        class_match = re.search(r"class\s*=\s*['\"]([^'\"]+)['\"]", attrs, flags=re.IGNORECASE)
        if not class_match:
            continue
        classes = class_match.group(1).lower()
        if any(token in classes for token in [
            "panel", "card", "kpi", "rail", "band", "tablebox", "takeaway", "chartwrap", "sidecard", "info"
        ]):
            count += 1
    return count


def _is_visual_contract_root_region(region: dict[str, object]) -> bool:
    rid = str(region.get("id", "")).strip().lower()
    role = str(region.get("role", "")).strip().lower()
    return rid in VISUAL_CONTRACT_ROOT_REGION_IDS or role in VISUAL_CONTRACT_ROOT_REGION_IDS


def _region_text_items(region: dict[str, object]) -> list[str]:
    raw = region.get("text")
    if not isinstance(raw, list):
        return []
    return [str(item).strip() for item in raw if str(item).strip()]


def _region_icon_items(region: dict[str, object]) -> list[str]:
    raw = region.get("icon_semantics")
    if not isinstance(raw, list):
        return []
    return [str(item).strip() for item in raw if str(item).strip()]


def _bbox_is_numeric(bbox: object) -> bool:
    return (
        isinstance(bbox, list)
        and len(bbox) == 4
        and all(isinstance(v, (int, float)) for v in bbox)
    )


def _contract_region_bbox_map(regions: list[object]) -> dict[str, list[float]]:
    out: dict[str, list[float]] = {}
    for region in regions:
        if not isinstance(region, dict):
            continue
        rid = str(region.get("id", "")).strip()
        bbox = region.get("bbox")
        if not rid or not _bbox_is_numeric(bbox):
            continue
        out[rid] = [float(v) for v in bbox]  # type: ignore[arg-type]
    return out


def _actual_html_ref_bboxes(html: Path) -> tuple[dict[str, list[float]] | None, str | None]:
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        return None, f"Playwright unavailable for HTML geometry audit: {exc}"
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                viewport={"width": 1920, "height": 1080},
                device_scale_factor=1,
            )
            page.goto(html.resolve().as_uri(), wait_until="networkidle")
            page.wait_for_timeout(100)
            data = page.evaluate(
                """
                () => {
                  const groups = {};
                  for (const el of document.querySelectorAll('[data-ref-id]')) {
                    const id = (el.getAttribute('data-ref-id') || '').trim();
                    if (!id) continue;
                    const style = window.getComputedStyle(el);
                    if (style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity || 1) === 0) continue;
                    const rect = el.getBoundingClientRect();
                    if (rect.width <= 0 || rect.height <= 0) continue;
                    if (!groups[id]) groups[id] = [];
                    groups[id].push([rect.left, rect.top, rect.width, rect.height]);
                  }
                  const out = {};
                  for (const [id, rects] of Object.entries(groups)) {
                    let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
                    for (const [x, y, w, h] of rects) {
                      x1 = Math.min(x1, x);
                      y1 = Math.min(y1, y);
                      x2 = Math.max(x2, x + w);
                      y2 = Math.max(y2, y + h);
                    }
                    out[id] = [x1, y1, x2 - x1, y2 - y1];
                  }
                  return out;
                }
                """
            )
            browser.close()
        if not isinstance(data, dict):
            return None, "HTML geometry audit returned invalid data"
        out: dict[str, list[float]] = {}
        for key, value in data.items():
            if isinstance(key, str) and _bbox_is_numeric(value):
                out[key] = [float(v) for v in value]
        return out, None
    except Exception as exc:
        return None, f"HTML geometry audit failed for {html.name}: {exc}"


def html_region_geometry_issues(
    html: Path,
    expected_bboxes: dict[str, list[float]],
) -> list[str]:
    actual_bboxes, error = _actual_html_ref_bboxes(html)
    if error:
        return [f"{html.name}: {error}"]
    if actual_bboxes is None:
        return [f"{html.name}: HTML geometry audit did not return element bboxes"]
    issues: list[str] = []
    for rid, expected in expected_bboxes.items():
        actual = actual_bboxes.get(rid)
        if actual is None:
            issues.append(f"{html.name}: data-ref-id '{rid}' is not visible or has zero geometry in the HTML render")
            continue
        ex, ey, ew, eh = expected
        ax, ay, aw, ah = actual
        ecx, ecy = ex + ew / 2, ey + eh / 2
        acx, acy = ax + aw / 2, ay + ah / 2
        center_tol_x = max(REGION_GEOMETRY_CENTER_TOLERANCE_PX, ew * 0.08)
        center_tol_y = max(REGION_GEOMETRY_CENTER_TOLERANCE_PX, eh * 0.08)
        size_tol_w = max(REGION_GEOMETRY_SIZE_TOLERANCE_PX, ew * REGION_GEOMETRY_SIZE_TOLERANCE_RATIO)
        size_tol_h = max(REGION_GEOMETRY_SIZE_TOLERANCE_PX, eh * REGION_GEOMETRY_SIZE_TOLERANCE_RATIO)
        if abs(acx - ecx) > center_tol_x or abs(acy - ecy) > center_tol_y:
            issues.append(
                f"{html.name}: data-ref-id '{rid}' center drift is too large "
                f"(expected center {ecx:.0f},{ecy:.0f}; actual {acx:.0f},{acy:.0f})"
            )
        if abs(aw - ew) > size_tol_w or abs(ah - eh) > size_tol_h:
            issues.append(
                f"{html.name}: data-ref-id '{rid}' size drift is too large "
                f"(expected {ew:.0f}x{eh:.0f}; actual {aw:.0f}x{ah:.0f})"
            )
    return issues


def check_visual_contract_files(task_dir: Path, expected: int, issues: list[str]) -> int:
    contracts = sorted((task_dir / "spec").glob("slide*_visual_contract.json"))
    if len(contracts) != expected:
        issues.append(f"visual contracts: expected {expected}, found {len(contracts)}")
    issues.extend(post_image_memory_boundary_issues(task_dir, expected))

    for idx in range(1, expected + 1):
        contract = visual_contract_path(task_dir, idx)
        if not contract.exists():
            issues.append(
                f"slide{idx:02d}_visual_contract.json missing; reconstruction source must be based "
                "on observed Image2 geometry"
            )
            continue
        try:
            raw_contract_text = contract.read_text(encoding="utf-8")
            data = json.loads(raw_contract_text)
        except Exception as exc:
            issues.append(f"slide{idx:02d}_visual_contract.json cannot be parsed: {exc}")
            continue

        ref = _normalize_ref_path(task_dir, data.get("reference_path"))
        expected_ref = (task_dir / "image2" / f"image2_reference_{idx:02d}.png").resolve()
        observation, observation_issues = image_observation_record_issues(task_dir, idx)
        issues.extend(observation_issues)
        issues.extend(visual_measurement_issues(task_dir, idx, data, observation))
        if ref is None or not ref.exists():
            issues.append(f"slide{idx:02d}_visual_contract.json reference_path missing or does not exist")
        elif ref.resolve() != expected_ref:
            issues.append(
                f"slide{idx:02d}_visual_contract.json reference_path must point to image2_reference_{idx:02d}.png"
            )
        if data.get("observed_from_reference") is not True:
            issues.append(f"slide{idx:02d}_visual_contract.json observed_from_reference must be true")
        if data.get("reconstruction_source") != IMAGE_ONLY_RECONSTRUCTION_SOURCE:
            issues.append(
                f"slide{idx:02d}_visual_contract.json reconstruction_source must be "
                f"{IMAGE_ONLY_RECONSTRUCTION_SOURCE}; visual contracts must treat the image as a fresh, standalone artifact"
            )
        if data.get("prompt_context_discarded") is not True:
            issues.append(
                f"slide{idx:02d}_visual_contract.json prompt_context_discarded must be true; "
                "do not carry prompt-library, final_prompt, or analysis memory into HTML reconstruction"
            )
        if data.get("observed_as_fresh_image") is not True:
            issues.append(
                f"slide{idx:02d}_visual_contract.json observed_as_fresh_image must be true; "
                "the contract must be authored as if the image were newly supplied with no prior production history"
            )
        if data.get("derivation_method") != POST_IMAGE_DERIVATION_METHOD:
            issues.append(
                f"slide{idx:02d}_visual_contract.json derivation_method must be {POST_IMAGE_DERIVATION_METHOD}; "
                "visual contract must derive from the fresh image observation record"
            )
        if data.get("contract_extraction_source") != "image_pixels":
            issues.append(
                f"slide{idx:02d}_visual_contract.json contract_extraction_source must be image_pixels; "
                "run extract-image2-contract after the fresh observation so initial geometry comes from the selected image"
            )
        if not str(data.get("contract_extraction_method", "")).strip():
            issues.append(
                f"slide{idx:02d}_visual_contract.json contract_extraction_method missing; "
                "visual contracts must record the image extraction path used before manual correction"
            )
        issues.extend(_visual_measurement_quality_issues(f"slide{idx:02d}_visual_contract.json", data))
        expected_observation_path = image2_observation_path(task_dir, idx)
        if data.get("observation_record_path") != str(expected_observation_path.relative_to(task_dir)):
            issues.append(
                f"slide{idx:02d}_visual_contract.json observation_record_path must point to "
                f"{expected_observation_path.relative_to(task_dir)}"
            )
        if expected_observation_path.exists() and data.get("observation_record_sha256") != sha256_file(expected_observation_path):
            issues.append(f"slide{idx:02d}_visual_contract.json observation_record_sha256 mismatch")
        if observation is not None and data.get("observation_id") != observation.get("observation_id"):
            issues.append(f"slide{idx:02d}_visual_contract.json observation_id must match the fresh image observation record")
        expected_measurement = visual_measurement_path(task_dir, idx)
        if data.get("measurement_path") != str(expected_measurement.relative_to(task_dir)):
            issues.append(
                f"slide{idx:02d}_visual_contract.json measurement_path must point to "
                f"{expected_measurement.relative_to(task_dir)}"
            )
        observation_evidence = str(data.get("observation_evidence", "")).strip()
        if len(observation_evidence) < 120:
            issues.append(
                f"slide{idx:02d}_visual_contract.json observation_evidence must describe fresh visual observation in at least 120 characters"
            )
        forbidden_contract = post_image_memory_markers(raw_contract_text)
        if forbidden_contract:
            issues.append(
                f"slide{idx:02d}_visual_contract.json contains upstream prompt/analysis memory markers after Image2 approval: "
                + ", ".join(forbidden_contract)
            )
        if not file_is_after_reference(task_dir, idx, contract):
            issues.append(
                f"slide{idx:02d}_visual_contract.json is older than image2_reference_{idx:02d}.png; "
                "visual contract must be created after observing the selected Image2 reference"
            )
        if not file_is_after_memory_boundary(task_dir, contract):
            issues.append(
                f"slide{idx:02d}_visual_contract.json is older than qa/post_image_memory_boundary.json; "
                "re-extract the contract after the forced post-image memory reset"
            )

        regions = data.get("regions")
        if not isinstance(regions, list) or len(regions) < MIN_VISUAL_CONTRACT_REGIONS:
            found = len(regions) if isinstance(regions, list) else 0
            issues.append(
                f"slide{idx:02d}_visual_contract.json is too coarse: expected at least "
                f"{MIN_VISUAL_CONTRACT_REGIONS} observed visual regions, found {found}. "
                "Do not collapse the reference into only nav/title/content/takeaway/source; "
                "declare each visible row, card, chart, table, KPI, rail, or action module as its own region before HTML."
            )
            continue

        seen_ids: set[str] = set()
        border_styles: set[str] = set()
        region_names: set[str] = set()
        detail_region_count = 0
        for pos, region in enumerate(regions, 1):
            if not isinstance(region, dict):
                issues.append(f"slide{idx:02d}_visual_contract.json region {pos} must be an object")
                continue
            rid = str(region.get("id", "")).strip()
            if not rid:
                issues.append(f"slide{idx:02d}_visual_contract.json region {pos} missing id")
            elif rid in seen_ids:
                issues.append(f"slide{idx:02d}_visual_contract.json duplicate region id: {rid}")
            else:
                seen_ids.add(rid)
            name = str(region.get("role", region.get("id", ""))).lower()
            region_names.add(name)
            bbox = region.get("bbox")
            if not _bbox_is_numeric(bbox):
                issues.append(f"slide{idx:02d}_visual_contract.json region {rid or pos} must include numeric bbox [x,y,w,h]")
            else:
                x, y, w, h = [float(v) for v in bbox]
                if w <= 0 or h <= 0:
                    issues.append(f"slide{idx:02d}_visual_contract.json region {rid or pos} bbox width/height must be positive")
                if x < -2 or y < -2 or x + w > 1922 or y + h > 1082:
                    issues.append(f"slide{idx:02d}_visual_contract.json region {rid or pos} bbox must stay within the 1920x1080 canvas")
            border_style = str(region.get("border_style", "")).lower().strip()
            if border_style not in VISUAL_CONTRACT_BORDER_STYLES:
                issues.append(
                    f"slide{idx:02d}_visual_contract.json region {rid or pos} has invalid border_style: {border_style or '<missing>'}"
                )
            else:
                border_styles.add(border_style)
            if "text" not in region:
                issues.append(f"slide{idx:02d}_visual_contract.json region {rid or pos} missing text list")
            elif "nav" in name:
                nav_text = " ".join(str(v) for v in region.get("text", []) if str(v).strip()).strip()
                if len(nav_text) < 2 or nav_text.lower() in {"nav", "navigation", "tab", "tabs"}:
                    issues.append(
                        f"slide{idx:02d}_visual_contract.json nav region must record the actual visible nav/tab text"
                    )
                if border_style == "none" and str(region.get("type", "")).lower().strip() in {"not_visible", "hidden", "none"}:
                    issues.append(
                        f"slide{idx:02d}_visual_contract.json nav cannot be hidden; deck navigation is a persistent visible region"
                    )
                if isinstance(bbox, list) and len(bbox) == 4 and all(isinstance(v, (int, float)) for v in bbox):
                    if bbox[2] <= 0 or bbox[3] <= 0:
                        issues.append(
                            f"slide{idx:02d}_visual_contract.json nav bbox must be non-zero; deck navigation is mandatory"
                        )
            if not _is_visual_contract_root_region(region):
                detail_region_count += 1
                if not _region_text_items(region) and not _region_icon_items(region):
                    issues.append(
                        f"slide{idx:02d}_visual_contract.json detail region {rid or pos} must record visible text "
                        "or icon_semantics from the reference; empty detail regions let HTML drift into a new design"
                    )

        if detail_region_count < MIN_VISUAL_CONTRACT_DETAIL_REGIONS:
            issues.append(
                f"slide{idx:02d}_visual_contract.json is too coarse: expected at least "
                f"{MIN_VISUAL_CONTRACT_DETAIL_REGIONS} non-root detail regions, found {detail_region_count}. "
                "The post-image contract must enumerate visible modules inside content instead of using one broad content box."
            )

        for required in VISUAL_CONTRACT_REQUIRED_REGIONS:
            if not any(required in name for name in region_names):
                issues.append(f"slide{idx:02d}_visual_contract.json missing required observed region role/id containing '{required}'")

        html = task_dir / "html" / f"slide{idx:02d}.html"
        html_text = read_text(html)
        if observation is not None:
            observation_id = str(observation.get("observation_id", "")).strip()
            if observation_id and observation_id not in html_text:
                issues.append(
                    f"slide{idx:02d}.html must include observation_id {observation_id}; "
                    "HTML must bind to the same fresh image observation record as the visual contract"
                )
        forbidden_html = post_image_memory_markers(html_text)
        if forbidden_html:
            issues.append(
                f"slide{idx:02d}.html contains upstream prompt/analysis memory markers after Image2 approval: "
                + ", ".join(forbidden_html)
            )
        html_ref_ids = _region_ids_from_html(html_text)
        missing_in_html = sorted(seen_ids - html_ref_ids)
        extra_in_html = sorted(html_ref_ids - seen_ids)
        if missing_in_html:
            issues.append(f"slide{idx:02d}.html missing data-ref-id bindings for visual regions: {', '.join(missing_in_html)}")
        if extra_in_html:
            issues.append(f"slide{idx:02d}.html has data-ref-id values not declared in visual contract: {', '.join(extra_in_html)}")
        if "dashed" in html_text.lower() and "dashed" not in border_styles:
            issues.append(f"slide{idx:02d}.html uses dashed borders, but visual contract declares no dashed border")
        unbound = _large_visual_class_count(html_text)
        if unbound:
            issues.append(f"slide{idx:02d}.html has {unbound} large visual container(s) without data-ref-id binding")
        issues.extend(html_region_geometry_issues(html, _contract_region_bbox_map(regions)))

    return len(contracts)


def validate_slide_review(
    data: dict[str, object],
    expected: int,
    issues: list[str],
    label: str,
    require_powerpoint: bool = False,
) -> dict[str, object] | None:
    slides = data.get("slides")
    if not isinstance(slides, list):
        issues.append(f"{label} missing slides list")
        return None
    if len(slides) != expected:
        issues.append(f"{label} slides: expected {expected}, found {len(slides)}")

    passed = 0
    for idx in range(1, expected + 1):
        entry = slides[idx - 1] if idx - 1 < len(slides) else {}
        if not isinstance(entry, dict):
            issues.append(f"{label} slide {idx}: entry must be an object")
            continue
        status = str(entry.get("status", "")).lower()
        compared = entry.get("compared")
        dimensions = entry.get("dimensions_checked")
        if status != "pass":
            issues.append(f"{label} slide {idx}: status must be pass")
        else:
            passed += 1
        if compared is not True:
            issues.append(f"{label} slide {idx}: compared must be true")
        if not isinstance(dimensions, list) or len(dimensions) < 5:
            issues.append(f"{label} slide {idx}: dimensions_checked must list at least 5 visual checks")
        if require_powerpoint:
            actual = entry.get("actual_powerpoint_opened")
            if actual is not True:
                issues.append(f"{label} slide {idx}: actual_powerpoint_opened must be true")
            checks = [str(item).lower() for item in dimensions] if isinstance(dimensions, list) else []
            for required in POWERPOINT_REVIEW_MIN_CHECKS:
                if not any(required in item for item in checks):
                    issues.append(f"{label} slide {idx}: dimensions_checked missing {required}")
    return {"passed_slides": passed}


def html_has_unmarked_img(text: str) -> bool:
    for match in re.finditer(r"<img\b[^>]*>", text, flags=re.IGNORECASE):
        tag = match.group(0)
        if "data-pptx-image" not in tag:
            return True
    return False


def strip_tags(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def class_suggests_icon(cls: str) -> bool:
    normalized = cls.strip().lower()
    if not normalized:
        return False
    if normalized in ICON_CLASS_EXACT_HINTS:
        return True
    return any(normalized.startswith(edge) or normalized.endswith(edge) for edge in ICON_CLASS_EDGE_HINTS)


def html_icon_placeholder_issues(text: str, name: str) -> list[str]:
    """Flag icon boxes that contain only short uppercase labels.

    Letter badges such as A/B/C section markers are allowed when their class is
    explicitly badge/letter/tab. Icon containers must be real visual symbols,
    a marked local image asset, or recognizable editable geometry.
    """
    issues: list[str] = []
    element_re = re.compile(
        r"<(?P<tag>div|span|i)\b(?P<attrs>[^>]*)>(?P<body>[^<>]{1,80})</(?P=tag)>",
        flags=re.IGNORECASE | re.DOTALL,
    )
    for match in element_re.finditer(text):
        attrs = match.group("attrs")
        body = match.group("body")
        class_match = re.search(r"class\s*=\s*['\"]([^'\"]+)['\"]", attrs, flags=re.IGNORECASE)
        classes = class_match.group(1).lower().split() if class_match else []
        if not any(class_suggests_icon(cls) for cls in classes):
            continue
        if any(cls in {"badge", "letter", "tab", "nav-badge", "section-badge"} for cls in classes):
            continue
        if "data-pptx-image" in attrs.lower() or "data-pptx-image" in body.lower():
            continue
        label = strip_tags(body)
        normalized = re.sub(r"[^A-Za-z0-9]+", "", label).upper()
        if not normalized:
            continue
        if normalized in ICON_PLACEHOLDER_WORDS or (
            len(normalized) <= 6 and normalized.isupper() and not normalized.isdigit()
        ):
            issues.append(
                f"{name}: icon-like element class='{class_match.group(1) if class_match else ''}' "
                f"uses text placeholder '{label}' instead of a recognizable icon or data-pptx-image asset"
            )
    return issues


def html_asset_issues(text: str, html: Path) -> list[str]:
    issues: list[str] = []
    for match in re.finditer(
        r"data-pptx-image\s*=\s*['\"]([^'\"]+)['\"]",
        text,
        flags=re.IGNORECASE,
    ):
        raw = match.group(1).strip()
        if not raw:
            issues.append(f"{html.name}: empty data-pptx-image asset path")
            continue
        path = Path(raw)
        if not path.is_absolute():
            path = html.parent / path
        if not path.exists():
            issues.append(f"{html.name}: data-pptx-image asset missing: {raw}")
            continue
        try:
            size = path.stat().st_size
        except Exception:
            size = 0
        if size < 500:
            issues.append(f"{html.name}: data-pptx-image asset suspiciously small: {raw}")
        if size > 2_000_000:
            issues.append(f"{html.name}: data-pptx-image asset too large for an icon/local crop: {raw}")
        issues.extend(icon_asset_background_issues(path, html.name, raw))
    return issues


def _load_pil():
    try:
        from PIL import Image
    except Exception as exc:
        raise RuntimeError(
            "Pillow is required for icon crop cleanup. Install pillow or use the bundled workspace Python."
        ) from exc
    return Image


def _parse_box(value: str) -> tuple[int, int, int, int]:
    parts = [p.strip() for p in re.split(r"[, ]+", value.strip()) if p.strip()]
    if len(parts) != 4:
        raise ValueError("--box must contain four integers: x,y,w,h")
    x, y, w, h = [int(float(p)) for p in parts]
    if w <= 0 or h <= 0:
        raise ValueError("--box width and height must be positive")
    return x, y, w, h


def _clamp_box(x: int, y: int, w: int, h: int, width: int, height: int) -> tuple[int, int, int, int]:
    x1 = max(0, min(width, x))
    y1 = max(0, min(height, y))
    x2 = max(0, min(width, x + w))
    y2 = max(0, min(height, y + h))
    if x2 <= x1 or y2 <= y1:
        raise ValueError("expanded crop box is outside the source image")
    return x1, y1, x2, y2


def _median_color(samples: list[tuple[int, int, int, int]]) -> tuple[int, int, int]:
    opaque = [s for s in samples if s[3] > 20]
    if not opaque:
        return (255, 255, 255)
    channels = []
    for idx in range(3):
        values = sorted(s[idx] for s in opaque)
        channels.append(values[len(values) // 2])
    return tuple(channels)  # type: ignore[return-value]


def _color_distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return sum((int(a[i]) - int(b[i])) ** 2 for i in range(3)) ** 0.5


def _detect_bg_color(im) -> tuple[int, int, int]:
    px = im.load()
    w, h = im.size
    samples: list[tuple[int, int, int, int]] = []
    edge = max(2, min(w, h) // 10)
    for x in range(w):
        for y in list(range(edge)) + list(range(max(edge, h - edge), h)):
            samples.append(px[x, y])
    for y in range(h):
        for x in list(range(edge)) + list(range(max(edge, w - edge), w)):
            samples.append(px[x, y])
    return _median_color(samples)


def clean_icon_crop_image(source: Path, box: tuple[int, int, int, int], output: Path, *, expand: int = 14, padding: int = 10, threshold: int = 34, min_canvas: int = 96) -> dict[str, object]:
    Image = _load_pil()
    src = Image.open(source).convert("RGBA")
    x, y, w, h = box
    x1, y1, x2, y2 = _clamp_box(x - expand, y - expand, w + expand * 2, h + expand * 2, src.width, src.height)
    crop = src.crop((x1, y1, x2, y2))
    bg = _detect_bg_color(crop)
    pixels = crop.load()
    for py in range(crop.height):
        for px in range(crop.width):
            r, g, b, a = pixels[px, py]
            if a == 0:
                continue
            dist = _color_distance((r, g, b), bg)
            if dist <= threshold:
                pixels[px, py] = (r, g, b, 0)
            elif dist <= threshold * 1.8:
                alpha = int(a * min(1.0, (dist - threshold) / max(1, threshold * 0.8)))
                pixels[px, py] = (r, g, b, alpha)
    bbox = crop.getchannel("A").getbbox()
    if bbox:
        crop = crop.crop(bbox)
    side = max(min_canvas, crop.width + padding * 2, crop.height + padding * 2)
    packed = Image.new("RGBA", (side, side), (255, 255, 255, 0))
    packed.alpha_composite(crop, ((side - crop.width) // 2, (side - crop.height) // 2))
    output.parent.mkdir(parents=True, exist_ok=True)
    packed.save(output)
    return {
        "ok": True,
        "source": str(source),
        "output": str(output),
        "input_box_xywh": [x, y, w, h],
        "expanded_box_xyxy": [x1, y1, x2, y2],
        "background_rgb_removed": list(bg),
        "output_size": list(packed.size),
    }


def icon_asset_background_issues(path: Path, html_name: str, raw: str) -> list[str]:
    if path.suffix.lower() not in {".png", ".webp"}:
        return []
    try:
        Image = _load_pil()
        im = Image.open(path).convert("RGBA")
    except Exception:
        return []
    if im.width < 16 or im.height < 16:
        return []
    corners = [
        im.getpixel((0, 0)),
        im.getpixel((im.width - 1, 0)),
        im.getpixel((0, im.height - 1)),
        im.getpixel((im.width - 1, im.height - 1)),
    ]
    opaque_corners = [c for c in corners if c[3] > 245]
    if len(opaque_corners) < 3:
        return []
    bg = _median_color(opaque_corners)
    similar = sum(1 for c in opaque_corners if _color_distance((c[0], c[1], c[2]), bg) < 8)
    if similar >= 3:
        return [
            f"{html_name}: data-pptx-image icon asset has opaque corner background: {raw}; "
            "run clean-icon-crop or export with transparent background"
        ]
    return []


def html_takeaway_style_issues(text: str, name: str) -> list[str]:
    issues: list[str] = []
    if "takeaway" not in text.lower() and "data-ref-id=\"takeaway\"" not in text.lower():
        return issues
    blocks: list[str] = []
    for match in re.finditer(r"(?P<selector>[^{}]*(?:takeaway|data-ref-id\s*=\s*['\"]takeaway['\"])[^{}]*)\{(?P<body>[^{}]+)\}", text, flags=re.IGNORECASE):
        blocks.append(match.group("body"))
    if not blocks:
        issues.append(f"{name}: takeaway must have an explicit CSS rule bound to the observed reference")
        return issues
    style = " ".join(blocks).lower()
    heights = [
        int(match.group(1))
        for match in re.finditer(r"(?:min-height|height)\s*:\s*(\d+)px", style, flags=re.IGNORECASE)
    ]
    if not heights:
        issues.append(f"{name}: takeaway CSS must declare height/min-height in px")
    elif max(heights) < 36:
        issues.append(f"{name}: takeaway height must be at least 36px so the slim bottom strip remains legible")
    elif max(heights) > 56:
        issues.append(f"{name}: takeaway height should stay slim, normally 36-48px; avoid large illustrated bottom banners")
    if "#305496" not in style and "rgb(48,84,150)" not in style.replace(" ", ""):
        issues.append(f"{name}: takeaway must use the locked deep-blue #305496 system fill")
    return issues


def html_nav_issues(text: str, name: str) -> list[str]:
    issues: list[str] = []
    nav_match = re.search(
        r"<nav\b(?P<attrs>[^>]*)>(?P<body>.*?)</nav>",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if nav_match:
        attrs = nav_match.group("attrs")
        body = nav_match.group("body")
    else:
        nav_match = re.search(
            r"<(?P<tag>div|section|header)\b(?P<attrs>[^>]*class\s*=\s*['\"][^'\"]*\b(?:nav|navbar|navigation|tabs|tabbar|breadcrumb)\b[^'\"]*['\"][^>]*)>(?P<body>.*?)</(?P=tag)>",
            text,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if not nav_match:
            return [f"{name}: missing semantic nav/.nav container"]
        attrs = nav_match.group("attrs")
        body = nav_match.group("body")

    nav_text = strip_tags(body)
    if len(nav_text) < 2:
        issues.append(f"{name}: nav/.nav container has no visible navigation text")
    if "data-ref-id" not in attrs.lower() or "nav" not in attrs.lower():
        issues.append(f"{name}: nav/.nav container must bind to data-ref-id=\"nav\"")
    if not re.search(r"class\s*=\s*['\"][^'\"]*\b(?:tab|nav-item|breadcrumb-item|active)\b", body, flags=re.IGNORECASE):
        issues.append(f"{name}: navigation should expose tab/nav-item/breadcrumb elements so PPTX keeps the bar")
    if not re.search(r"(align-items\s*:\s*center|text-align\s*:\s*center|justify-content\s*:\s*center)", text, flags=re.IGNORECASE):
        issues.append(f"{name}: navigation text should be explicitly centered for PPTX conversion")
    return issues


def check_html_files(task_dir: Path, expected: int, issues: list[str]) -> int:
    html_refs = sorted((task_dir / "html").glob("slide*.html"))
    if len(html_refs) != expected:
        issues.append(f"HTML slides: expected {expected}, found {len(html_refs)}")
    issues.extend(post_image_memory_boundary_issues(task_dir, expected))

    language_family = requested_language_family(task_dir)
    for html in html_refs:
        text = read_text(html)
        lower = text.lower()
        name = html.name
        if "<svg" in lower:
            issues.append(f"{name}: inline SVG is forbidden")
        if "background-image" in lower:
            issues.append(f"{name}: CSS background-image is forbidden")
        if html_has_unmarked_img(text):
            issues.append(f"{name}: unmarked <img> found; use data-pptx-image for local assets")
        issues.extend(html_icon_placeholder_issues(text, name))
        issues.extend(html_asset_issues(text, html))
        issues.extend(html_takeaway_style_issues(text, name))
        if "data-pptx-image" in lower and ("1920" in lower and "1080" in lower and "whole-slide" in lower):
            issues.append(f"{name}: possible whole-slide preserved image")
        if "font-family" not in lower:
            issues.append(f"{name}: missing explicit font-family")
        elif not any(font.lower() in lower for font in ["arial", "microsoft yahei", "pingfang sc"]):
            issues.append(f"{name}: font-family should include Arial and Chinese-safe fallback")
        issues.extend(html_nav_issues(text, name))
        issues.extend(language_consistency_issues(text, name, language_family))
        if "takeaway" not in lower and "take" not in lower:
            issues.append(f"{name}: missing takeaway bar/class")
        idx_match = re.search(r"slide(\d+)\.html$", name)
        if idx_match:
            idx = int(idx_match.group(1))
            if not file_is_after_reference(task_dir, idx, html):
                issues.append(
                    f"{name}: HTML is older than image2_reference_{idx:02d}.png; "
                    "HTML must be authored after observing the selected Image2 reference"
                )
            if not file_is_after_memory_boundary(task_dir, html):
                issues.append(
                    f"{name}: HTML is older than qa/post_image_memory_boundary.json; "
                    "rewrite HTML after the forced post-image memory reset"
                )
    return len(html_refs)


def check_pptx_file(pptx: Path, expected: int, issues: list[str]) -> dict[str, int] | None:
    if not pptx.exists():
        issues.append(f"PPTX missing: {pptx}")
        return None
    try:
        from pptx import Presentation
    except Exception:
        issues.append("python-pptx unavailable; cannot inspect PPTX")
        return None

    try:
        prs = Presentation(pptx)
    except Exception as exc:
        issues.append(f"PPTX cannot be opened: {exc}")
        return None

    slide_count = len(prs.slides)
    if slide_count != expected:
        issues.append(f"PPTX slides: expected {expected}, found {slide_count}")

    picture_count = 0
    text_shape_count = 0
    suspicious_full_slide_images = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == 13:
                picture_count += 1
                if shape.width >= prs.slide_width * 0.9 and shape.height >= prs.slide_height * 0.9:
                    suspicious_full_slide_images += 1
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                text_shape_count += 1

    if suspicious_full_slide_images:
        issues.append(f"PPTX contains {suspicious_full_slide_images} possible whole-slide image background(s)")
    if text_shape_count < expected * 6:
        issues.append("PPTX has too few editable text shapes; likely flattened or incomplete")

    return {
        "slides": slide_count,
        "pictures": picture_count,
        "text_shapes": text_shape_count,
        "suspicious_full_slide_images": suspicious_full_slide_images,
    }


def render_manifest_path(task_dir: Path) -> Path:
    return task_dir / "qa" / "render_manifest.json"


def check_render_manifest(
    task_dir: Path,
    pptx: Path | None,
    issues: list[str],
) -> dict[str, object] | None:
    manifest_path = render_manifest_path(task_dir)
    if not manifest_path.exists():
        issues.append("qa/render_manifest.json missing; render with paopao_run.py render before delivery")
        return None
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"qa/render_manifest.json cannot be parsed: {exc}")
        return None

    recorded_pptx = Path(str(data.get("pptx_path", "")))
    if pptx is None:
        pptx = recorded_pptx
    else:
        pptx = pptx.resolve()
    if not str(recorded_pptx).endswith(".pptx"):
        issues.append("qa/render_manifest.json missing pptx_path")
    elif recorded_pptx.resolve() != pptx.resolve():
        issues.append(
            f"render manifest PPTX mismatch: manifest has {recorded_pptx}, checked {pptx}"
        )
    if not pptx.exists():
        issues.append(f"render manifest PPTX missing: {pptx}")
        return {"path": str(manifest_path.resolve()), "pptx_hash_matches": False}

    current_hash = sha256_file(pptx)
    recorded_hash = str(data.get("pptx_sha256", ""))
    if current_hash != recorded_hash:
        issues.append(
            "PPTX does not match qa/render_manifest.json; rerender from the current HTML source before delivery"
        )
    return {
        "path": str(manifest_path.resolve()),
        "pptx_hash_matches": current_hash == recorded_hash,
        "pptx": str(pptx),
    }


def load_commercial_render_contract(task_dir: Path) -> dict[str, object]:
    path = commercial_render_contract_path(task_dir)
    if not path.exists():
        return {}
    data = _read_json_file(path)
    return data or {}


def commercial_render_path(task_dir: Path) -> str:
    data = load_commercial_render_contract(task_dir)
    path = str(data.get("render_path", "")).strip()
    return path if path in COMMERCIAL_RENDER_PATHS else "html"


def check_commercial_render_contract(
    task_dir: Path,
    expected: int,
    pptx: Path,
    issues: list[str],
) -> dict[str, object] | None:
    path = commercial_render_contract_path(task_dir)
    if not path.exists():
        issues.append(
            "qa/commercial_render_contract.json missing; commercial delivery must declare html or direct_pptx "
            "as the render path and bind the final PPTX to Image2 references"
        )
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"qa/commercial_render_contract.json cannot be parsed: {exc}")
        return None
    if not isinstance(data, dict):
        issues.append("qa/commercial_render_contract.json must be a JSON object")
        return None

    render_path = str(data.get("render_path", "")).strip()
    if data.get("schema") != COMMERCIAL_RENDER_CONTRACT_SCHEMA:
        issues.append(
            f"qa/commercial_render_contract.json schema must be {COMMERCIAL_RENDER_CONTRACT_SCHEMA}"
        )
    if render_path not in COMMERCIAL_RENDER_PATHS:
        issues.append("qa/commercial_render_contract.json render_path must be html or direct_pptx")
    if data.get("source_of_truth") != "image2_reference":
        issues.append("qa/commercial_render_contract.json source_of_truth must be image2_reference")
    if data.get("post_image_inputs_only") is not True:
        issues.append("qa/commercial_render_contract.json post_image_inputs_only must be true")
    try:
        min_score = float(data.get("commercial_similarity_min", 0))
    except Exception:
        min_score = 0.0
    if min_score < COMMERCIAL_SIMILARITY_MIN:
        issues.append(
            f"qa/commercial_render_contract.json commercial_similarity_min must be at least "
            f"{COMMERCIAL_SIMILARITY_MIN:.2f}"
        )
    recorded_pptx = _resolve_task_path(task_dir, data.get("pptx_path"))
    if recorded_pptx is None or recorded_pptx.suffix.lower() != ".pptx":
        issues.append("qa/commercial_render_contract.json pptx_path must point to the final PPTX")
    elif recorded_pptx.resolve() != pptx.resolve():
        issues.append("qa/commercial_render_contract.json pptx_path does not match the checked PPTX")
    if pptx.exists() and data.get("pptx_sha256") != sha256_file(pptx):
        issues.append("qa/commercial_render_contract.json pptx_sha256 does not match the checked PPTX")
    if data.get("actual_preview_dir") != "qa/pptx_actual":
        issues.append("qa/commercial_render_contract.json actual_preview_dir must be qa/pptx_actual")
    if render_path == "direct_pptx" and data.get("html_is_debug_only") is not True:
        issues.append("qa/commercial_render_contract.json direct_pptx path must set html_is_debug_only true")
    if render_path == "html" and data.get("html_is_debug_only") is True:
        issues.append("qa/commercial_render_contract.json html path cannot mark HTML as debug-only")

    return {
        "path": str(path.resolve()),
        "render_path": render_path or None,
        "commercial_similarity_min": min_score,
        "expected_pages": expected,
    }


def check_commercial_similarity(
    task_dir: Path,
    expected: int,
    issues: list[str],
) -> dict[str, object]:
    review = task_dir / "qa" / "fidelity_review.json"
    scores: list[dict[str, object]] = []
    if not review.exists():
        issues.append("qa/fidelity_review.json missing; commercial gate needs real PPTX preview comparisons")
        return {"min_required": COMMERCIAL_SIMILARITY_MIN, "scores": scores}
    try:
        data = json.loads(review.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"qa/fidelity_review.json cannot be parsed for commercial gate: {exc}")
        return {"min_required": COMMERCIAL_SIMILARITY_MIN, "scores": scores}
    slides = data.get("slides")
    if not isinstance(slides, list):
        issues.append("qa/fidelity_review.json missing slides list for commercial gate")
        return {"min_required": COMMERCIAL_SIMILARITY_MIN, "scores": scores}
    for idx in range(1, expected + 1):
        entry = slides[idx - 1] if idx - 1 < len(slides) else {}
        if not isinstance(entry, dict):
            issues.append(f"commercial gate slide {idx}: fidelity review entry missing")
            continue
        ref = _resolve_review_path(task_dir, entry.get("reference_path"))
        actual = _resolve_review_path(task_dir, entry.get("actual_preview_path"))
        if ref is None or actual is None or not ref.exists() or not actual.exists():
            issues.append(f"commercial gate slide {idx}: reference and actual PPTX preview images are required")
            continue
        if not _path_is_under(actual, task_dir / "qa" / "pptx_actual"):
            issues.append(f"commercial gate slide {idx}: actual preview must be under qa/pptx_actual")
            continue
        if _looks_like_html_reference_preview(task_dir, idx, actual):
            issues.append(f"commercial gate slide {idx}: actual preview cannot be copied from HTML preview")
            continue
        score = image_similarity_score(ref, actual)
        if score is None:
            issues.append(f"commercial gate slide {idx}: cannot compute Image2-vs-PowerPoint preview score")
            continue
        rounded = round(score, 4)
        scores.append({"slide": idx, "commercial_similarity_score": rounded})
        if score < COMMERCIAL_SIMILARITY_MIN:
            issues.append(
                f"commercial gate slide {idx}: final PowerPoint preview similarity {score:.3f} below "
                f"commercial minimum {COMMERCIAL_SIMILARITY_MIN:.3f}"
            )
    return {"min_required": COMMERCIAL_SIMILARITY_MIN, "scores": scores}


def internal_prompt_files(task_dir: Path) -> list[Path]:
    files: dict[Path, None] = {}
    for pattern in PROMPT_INTERNAL_PATTERNS:
        for path in task_dir.glob(pattern):
            if not path.is_file():
                continue
            resolved = path.resolve()
            files[resolved] = None
    return sorted(files)


def prompt_delivery_files(task_dir: Path) -> list[Path]:
    delivery_dir = task_dir / "delivery"
    if not delivery_dir.exists():
        return []
    return sorted(
        path.resolve()
        for path in delivery_dir.rglob("*")
        if path.is_file() and "prompt" in path.name.lower()
    )


def delivery_forbidden_user_file(path: Path) -> bool:
    lowered_name = path.name.lower()
    lowered_parts = {part.lower() for part in path.parts}
    if path.suffix.lower() not in DELIVERY_ALLOWED_SUFFIXES:
        return True
    if path.name.startswith("~$"):
        return True
    if path.suffix.lower() == ".pptx":
        return False
    if path.suffix.lower() in {".md", ".markdown", ".json"}:
        return True
    return any(part in lowered_name or part in lowered_parts for part in DELIVERY_FORBIDDEN_NAME_PARTS)


def delivery_temp_files(task_dir: Path) -> list[Path]:
    files: dict[Path, None] = {}
    for pattern in DELIVERY_TEMP_PATTERNS:
        for path in task_dir.glob(f"**/{pattern}"):
            if path.is_file():
                files[path.resolve()] = None
    return sorted(files)


def _review_path_exists(task_dir: Path, raw: object) -> bool:
    if not isinstance(raw, str) or not raw.strip():
        return False
    path = Path(raw)
    if not path.is_absolute():
        path = task_dir / path
    return path.exists() and path.is_file()


def _resolve_review_path(task_dir: Path, raw: object) -> Path | None:
    if not isinstance(raw, str) or not raw.strip():
        return None
    path = Path(raw)
    if not path.is_absolute():
        path = task_dir / path
    return path.resolve()


def _path_is_under(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def _looks_like_html_reference_preview(task_dir: Path, idx: int, path: Path) -> bool:
    html_preview = html_reference_preview_path(task_dir, idx)
    html_preview_dir = task_dir / "qa" / "html_reference"
    if _path_is_under(path, html_preview_dir):
        return True
    if html_preview.exists() and path.exists() and path.is_file():
        try:
            return sha256_file(path) == sha256_file(html_preview)
        except Exception:
            return False
    return False


def check_fidelity_review(task_dir: Path, expected: int, issues: list[str]) -> dict[str, object] | None:
    review = task_dir / "qa" / "fidelity_review.json"
    if not review.exists():
        issues.append("qa/fidelity_review.json missing; compare final PPTX against Image2 slide by slide before delivery")
        return None
    try:
        data = json.loads(review.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"qa/fidelity_review.json cannot be parsed: {exc}")
        return None

    summary = validate_slide_review(data, expected, issues, "fidelity review")
    if summary is None:
        return None
    slides = data.get("slides", [])
    scored_slides: list[dict[str, object]] = []
    for idx in range(1, expected + 1):
        entry = slides[idx - 1] if idx - 1 < len(slides) else {}
        dimensions = entry.get("dimensions_checked") if isinstance(entry, dict) else []
        checks = [str(item).lower() for item in dimensions] if isinstance(dimensions, list) else []
        for required in FIDELITY_REVIEW_MIN_CHECKS:
            if not any(required in item for item in checks):
                issues.append(f"fidelity review slide {idx}: dimensions_checked missing {required}")
        if isinstance(entry, dict):
            reference_raw = entry.get("reference_path")
            actual_raw = entry.get("actual_preview_path")
            if not _review_path_exists(task_dir, reference_raw):
                issues.append(
                    f"fidelity review slide {idx}: reference_path must point to the selected Image2 reference"
                )
            if not _review_path_exists(task_dir, actual_raw):
                issues.append(
                    f"fidelity review slide {idx}: actual_preview_path must point to a real PPTX preview image"
                )
            else:
                actual_resolved = _resolve_review_path(task_dir, actual_raw)
                if actual_resolved is not None:
                    pptx_actual_dir = task_dir / "qa" / "pptx_actual"
                    if not _path_is_under(actual_resolved, pptx_actual_dir):
                        issues.append(
                            f"fidelity review slide {idx}: actual_preview_path must be exported from the final PPTX "
                            "under qa/pptx_actual/, not from HTML/browser preview output"
                        )
                    if _looks_like_html_reference_preview(task_dir, idx, actual_resolved):
                        issues.append(
                            f"fidelity review slide {idx}: actual_preview_path matches the HTML reference preview; "
                            "export and review the actual PowerPoint-rendered slide instead"
                        )
            evidence = entry.get("evidence")
            if not isinstance(evidence, str) or len(evidence.strip()) < 80:
                issues.append(
                    f"fidelity review slide {idx}: evidence must describe concrete visual comparisons, not a generic pass"
                )
            if _review_path_exists(task_dir, reference_raw) and _review_path_exists(task_dir, actual_raw):
                reference_path = Path(reference_raw) if isinstance(reference_raw, str) else Path()
                actual_path = Path(actual_raw) if isinstance(actual_raw, str) else Path()
                if not reference_path.is_absolute():
                    reference_path = task_dir / reference_path
                if not actual_path.is_absolute():
                    actual_path = task_dir / actual_path
                score = image_similarity_score(reference_path, actual_path)
                if score is None:
                    issues.append(
                        f"fidelity review slide {idx}: cannot compute image similarity; Pillow and readable PNG/JPG previews are required"
                    )
                else:
                    scored_slides.append({"slide": idx, "image_similarity_score": round(score, 4)})
                    recorded = entry.get("image_similarity_score")
                    if recorded is not None:
                        try:
                            recorded_score = float(recorded)
                            if abs(recorded_score - score) > 0.03:
                                issues.append(
                                    f"fidelity review slide {idx}: recorded image_similarity_score {recorded_score:.3f} "
                                    f"does not match computed score {score:.3f}"
                                )
                        except Exception:
                            issues.append(f"fidelity review slide {idx}: image_similarity_score must be numeric when present")
                    if score < MIN_FIDELITY_IMAGE_SCORE:
                        issues.append(
                            f"fidelity review slide {idx}: image similarity {score:.3f} below minimum "
                            f"{MIN_FIDELITY_IMAGE_SCORE:.3f}; refill the measured reconstruction source and rerender "
                            "instead of delivering a rough interpretation"
                        )
    return {
        "path": str(review.resolve()),
        "min_image_similarity_score": MIN_FIDELITY_IMAGE_SCORE,
        "image_similarity_scores": scored_slides,
        **summary,
    }


def check_powerpoint_review(task_dir: Path, expected: int, issues: list[str]) -> dict[str, object] | None:
    review = task_dir / "qa" / "powerpoint_review.json"
    if not review.exists():
        issues.append("qa/powerpoint_review.json missing; open the actual PPTX in PowerPoint and inspect it slide by slide")
        return None
    try:
        data = json.loads(review.read_text(encoding="utf-8"))
    except Exception as exc:
        issues.append(f"qa/powerpoint_review.json cannot be parsed: {exc}")
        return None

    if data.get("actual_pptx_opened") is not True:
        issues.append("qa/powerpoint_review.json actual_pptx_opened must be true")
    if not str(data.get("pptx_path", "")).endswith(".pptx"):
        issues.append("qa/powerpoint_review.json must include pptx_path")

    summary = validate_slide_review(
        data,
        expected,
        issues,
        "PowerPoint review",
        require_powerpoint=True,
    )
    if summary is None:
        return None
    return {"path": str(review.resolve()), **summary}


def check_delivery_files(
    task_dir: Path,
    expected: int,
    issues: list[str],
    *,
    require_final_pass: bool = True,
) -> dict[str, object]:
    prompt_files = prompt_delivery_files(task_dir)
    temp_files = delivery_temp_files(task_dir)
    if prompt_files:
        issues.append(
            "prompt Markdown files still present after delivery cleanup: "
            + ", ".join(str(p.relative_to(task_dir)) for p in prompt_files)
        )
    if temp_files:
        issues.append(
            "temporary Office/cache files still present in delivery tree: "
            + ", ".join(str(p.relative_to(task_dir)) for p in temp_files)
        )
    pptx_dir = task_dir / "pptx"
    exposed_pptx = [
        p for p in sorted(pptx_dir.glob("*.pptx"))
        if p.is_file() and not p.name.startswith("~$")
    ]
    if len(exposed_pptx) > 1:
        issues.append(
            "multiple top-level PPTX files exposed; move drafts to pptx/_drafts and leave only the final delivery file: "
            + ", ".join(p.name for p in exposed_pptx)
        )
    check_image2_files(task_dir, expected, issues, require_prompt_files=False)
    spec_count = check_spec_files(task_dir, expected, issues)
    visual_contract_count = check_visual_contract_files(task_dir, expected, issues)
    html_count = check_html_files(task_dir, expected, issues)
    final_pptx = exposed_pptx[0].resolve() if len(exposed_pptx) == 1 else None
    pptx_summary = check_pptx_file(final_pptx, expected, issues) if final_pptx else None
    commercial = check_commercial_render_contract(task_dir, expected, final_pptx, issues) if final_pptx else None
    render_path = commercial.get("render_path") if isinstance(commercial, dict) else commercial_render_path(task_dir)
    if render_path == "html":
        render_manifest: object = check_render_manifest(task_dir, final_pptx, issues)
    else:
        render_manifest = "not_required_for_direct_pptx"
    fidelity = check_fidelity_review(task_dir, expected, issues)
    commercial_similarity = check_commercial_similarity(task_dir, expected, issues)
    powerpoint = check_powerpoint_review(task_dir, expected, issues)
    delivery_dir = task_dir / "delivery"
    delivery_files: list[Path] = []
    if delivery_dir.exists():
        delivery_files = [p for p in sorted(delivery_dir.rglob("*")) if p.is_file()]
        pptx_delivery = [
            p for p in delivery_files
            if p.suffix.lower() == ".pptx" and not p.name.startswith("~$")
        ]
        html_delivery = [
            p for p in delivery_files
            if p.suffix.lower() == ".html" and "html" in {part.lower() for part in p.relative_to(delivery_dir).parts}
        ]
        image_delivery = [
            p for p in delivery_files
            if p.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".gif"}
            and "images" in {part.lower() for part in p.relative_to(delivery_dir).parts}
        ]
        forbidden_delivery = [
            p for p in delivery_files
            if delivery_forbidden_user_file(p.relative_to(delivery_dir))
        ]
        if len(pptx_delivery) != 1:
            issues.append(
                "delivery directory must contain exactly one user-facing PPTX: "
                + str(delivery_dir)
            )
        if render_path == "html" and len(html_delivery) != expected:
            issues.append(
                f"HTML render path delivery must contain exactly {expected} user-facing HTML slide(s) under delivery/html; "
                f"found {len(html_delivery)}"
            )
        if len(image_delivery) != expected:
            issues.append(
                f"delivery directory must contain exactly {expected} user-facing slide image(s) under delivery/images; "
                f"found {len(image_delivery)}"
            )
        if forbidden_delivery:
            issues.append(
                "delivery directory contains internal or temporary files: "
                + ", ".join(str(p.relative_to(delivery_dir)) for p in forbidden_delivery)
            )
    if require_final_pass:
        issues.extend(final_delivery_pass_issues(task_dir, expected, final_pptx))
    return {
        "prompt_markdown_files": [str(p.relative_to(task_dir)) for p in prompt_files],
        "temporary_files": [str(p.relative_to(task_dir)) for p in temp_files],
        "top_level_pptx_files": [p.name for p in exposed_pptx],
        "spec_count": spec_count,
        "visual_contract_count": visual_contract_count,
        "html_slide_count": html_count,
        "pptx_summary": pptx_summary,
        "commercial_render_contract": commercial,
        "render_manifest": render_manifest,
        "fidelity_review": fidelity,
        "commercial_similarity": commercial_similarity,
        "powerpoint_review": powerpoint,
        "delivery_files": [
            str(p.relative_to(delivery_dir)) for p in delivery_files
        ] if delivery_files else [],
        "delivery_contract": {
            "pptx": 1,
            "html_slides": expected if render_path == "html" else "debug_optional",
            "images": expected,
            "forbidden": "prompt/markdown/json/analysis/spec/qa/internal files",
        },
        "final_delivery_pass": str(final_delivery_pass_path(task_dir).relative_to(task_dir)),
    }


def check_pipeline_contract(
    task_dir: Path,
    expected: int,
    pptx: Path | None,
    issues: list[str],
) -> dict[str, object]:
    counts: dict[str, object] = {}
    check_analysis_files(task_dir, expected, issues)
    counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
    counts["spec_count"] = check_spec_files(task_dir, expected, issues)
    counts["visual_contract_count"] = check_visual_contract_files(task_dir, expected, issues)
    if pptx is None:
        pptx_files = sorted((task_dir / "pptx").glob("*.pptx"))
        pptx = pptx_files[-1].resolve() if pptx_files else task_dir / "pptx" / "missing.pptx"
    commercial = check_commercial_render_contract(task_dir, expected, pptx, issues)
    counts["commercial_render_contract"] = commercial
    render_path = commercial.get("render_path") if isinstance(commercial, dict) else commercial_render_path(task_dir)
    if render_path == "html":
        counts["html_slide_count"] = check_html_files(task_dir, expected, issues)
        counts["html_reference_fidelity"] = check_html_reference_fidelity(task_dir, expected, issues)
        counts["render_manifest"] = check_render_manifest(task_dir, pptx, issues)
    else:
        counts["html_slide_count"] = "debug_optional"
        counts["html_reference_fidelity"] = "not_required_for_direct_pptx"
        counts["render_manifest"] = "not_required_for_direct_pptx"
    pptx_summary = check_pptx_file(pptx, expected, issues)
    if pptx_summary is not None:
        counts["pptx"] = str(pptx)
        counts["pptx_summary"] = pptx_summary
    counts["powerpoint_review"] = check_powerpoint_review(task_dir, expected, issues)
    counts["fidelity_review"] = check_fidelity_review(task_dir, expected, issues)
    counts["commercial_similarity"] = check_commercial_similarity(task_dir, expected, issues)
    return counts


def write_pipeline_pass(
    task_dir: Path,
    expected: int,
    pptx: Path,
    pipeline_counts: dict[str, object],
) -> Path:
    pptx = pptx.resolve()
    receipt = {
        "schema": PIPELINE_PASS_SCHEMA,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "task_dir": str(task_dir.resolve()),
        "expected_pages": expected,
        "source_pptx": _relative_to_task_or_abs(task_dir, pptx),
        "pptx_sha256": sha256_file(pptx),
        "pipeline_counts": pipeline_counts,
        "policy": (
            "Delivery publishing is allowed only after analysis, Image2 references, fresh "
            "visual contracts, HTML, render manifest, PowerPoint review, and fidelity review pass."
        ),
    }
    path = pipeline_pass_path(task_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(receipt, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def pipeline_pass_issues(task_dir: Path, expected: int, pptx: Path) -> list[str]:
    path = pipeline_pass_path(task_dir)
    if not path.exists():
        return [
            "qa/pipeline_pass.json missing; run finalize-delivery after the full pipeline check passes"
        ]
    data = _read_json_file(path)
    if data is None:
        return ["qa/pipeline_pass.json cannot be parsed"]
    issues: list[str] = []
    if data.get("schema") != PIPELINE_PASS_SCHEMA:
        issues.append(f"qa/pipeline_pass.json schema must be {PIPELINE_PASS_SCHEMA}")
    if data.get("expected_pages") != expected:
        issues.append("qa/pipeline_pass.json expected_pages does not match current task")
    recorded_pptx = _resolve_task_path(task_dir, data.get("source_pptx"))
    if recorded_pptx is None or not recorded_pptx.exists():
        issues.append("qa/pipeline_pass.json source_pptx is missing or invalid")
    elif recorded_pptx.resolve() != pptx.resolve():
        issues.append("qa/pipeline_pass.json source_pptx does not match the requested PPTX")
    if not pptx.exists():
        issues.append(f"PPTX missing or invalid: {pptx}")
    elif data.get("pptx_sha256") != sha256_file(pptx):
        issues.append("qa/pipeline_pass.json pptx_sha256 does not match the current PPTX")
    if not isinstance(data.get("pipeline_counts"), dict):
        issues.append("qa/pipeline_pass.json pipeline_counts must be present")
    return issues


def write_final_delivery_pass(
    task_dir: Path,
    expected: int,
    pptx: Path,
    delivery_counts: dict[str, object],
) -> Path:
    delivery_dir = task_dir / "delivery"
    delivery_files = [
        _relative_to_task_or_abs(task_dir, p)
        for p in sorted(delivery_dir.rglob("*"))
        if p.is_file()
    ] if delivery_dir.exists() else []
    receipt = {
        "schema": FINAL_DELIVERY_PASS_SCHEMA,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "task_dir": str(task_dir.resolve()),
        "expected_pages": expected,
        "source_pptx": _relative_to_task_or_abs(task_dir, pptx.resolve()),
        "pptx_sha256": sha256_file(pptx.resolve()),
        "delivery_files": delivery_files,
        "delivery_counts": delivery_counts,
        "policy": (
            "Final response may link only to files under delivery/ after this pass is current."
        ),
    }
    path = final_delivery_pass_path(task_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(receipt, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def final_delivery_pass_issues(task_dir: Path, expected: int, pptx: Path | None) -> list[str]:
    path = final_delivery_pass_path(task_dir)
    if not path.exists():
        return [
            "qa/final_delivery_pass.json missing; use finalize-delivery before replying to the user"
        ]
    data = _read_json_file(path)
    if data is None:
        return ["qa/final_delivery_pass.json cannot be parsed"]
    issues: list[str] = []
    if data.get("schema") != FINAL_DELIVERY_PASS_SCHEMA:
        issues.append(f"qa/final_delivery_pass.json schema must be {FINAL_DELIVERY_PASS_SCHEMA}")
    if data.get("expected_pages") != expected:
        issues.append("qa/final_delivery_pass.json expected_pages does not match current task")
    if pptx is not None:
        recorded_pptx = _resolve_task_path(task_dir, data.get("source_pptx"))
        if recorded_pptx is None or not recorded_pptx.exists():
            issues.append("qa/final_delivery_pass.json source_pptx is missing or invalid")
        elif recorded_pptx.resolve() != pptx.resolve():
            issues.append("qa/final_delivery_pass.json source_pptx does not match the current PPTX")
        elif data.get("pptx_sha256") != sha256_file(pptx):
            issues.append("qa/final_delivery_pass.json pptx_sha256 does not match the current PPTX")
    if not isinstance(data.get("delivery_files"), list):
        issues.append("qa/final_delivery_pass.json delivery_files must be present")
    return issues


def classify_audit_issue(issue: str) -> str:
    lower = issue.lower()
    if any(token in lower for token in [
        "image2",
        "registered_source",
        "provenance",
        "tool_call_id",
        "render_manifest",
        "powerpoint_review",
        "fidelity_review",
        "whole-slide",
        "missing",
        "mismatch",
        "older than",
        "delivery directory",
        "temporary",
        "prompt markdown",
    ]):
        return "blocker"
    if any(token in lower for token in ["too short", "placeholder", "font", "takeaway", "icon"]):
        return "warning"
    return "info"


def cmd_audit_task(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")
    pptx = Path(args.pptx).resolve() if args.pptx else None

    stage_results: dict[str, dict[str, object]] = {}
    for stage in ["analysis", "image2", "html", "pptx", "pipeline", "delivery"]:
        issues: list[str] = []
        counts: dict[str, object] = {}
        if stage == "analysis":
            check_analysis_files(task_dir, expected, issues)
        elif stage == "image2":
            check_analysis_files(task_dir, expected, issues)
            counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
        elif stage == "html":
            check_analysis_files(task_dir, expected, issues)
            counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
            counts["spec_count"] = check_spec_files(task_dir, expected, issues)
            counts["visual_contract_count"] = check_visual_contract_files(task_dir, expected, issues)
            counts["html_slide_count"] = check_html_files(task_dir, expected, issues)
            counts["html_reference_fidelity"] = check_html_reference_fidelity(task_dir, expected, issues)
        elif stage == "pptx":
            check_analysis_files(task_dir, expected, issues)
            counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
            counts["spec_count"] = check_spec_files(task_dir, expected, issues)
            counts["visual_contract_count"] = check_visual_contract_files(task_dir, expected, issues)
            counts["html_slide_count"] = check_html_files(task_dir, expected, issues)
            counts["html_reference_fidelity"] = check_html_reference_fidelity(task_dir, expected, issues)
            check_pptx = pptx
            if check_pptx is None:
                pptx_files = sorted((task_dir / "pptx").glob("*.pptx"))
                check_pptx = pptx_files[-1].resolve() if pptx_files else task_dir / "pptx" / "missing.pptx"
            pptx_summary = check_pptx_file(check_pptx, expected, issues)
            if pptx_summary is not None:
                counts["pptx"] = str(check_pptx)
                counts["pptx_summary"] = pptx_summary
            counts["render_manifest"] = check_render_manifest(task_dir, check_pptx, issues)
        elif stage == "pipeline":
            counts.update(check_pipeline_contract(task_dir, expected, pptx, issues))
        elif stage == "delivery":
            counts.update(check_delivery_files(task_dir, expected, issues))

        grouped: dict[str, list[str]] = {"blocker": [], "warning": [], "info": []}
        for issue in issues:
            grouped[classify_audit_issue(issue)].append(issue)
        stage_results[stage] = {
            "ok": not issues,
            "issue_count": len(issues),
            "issues_by_severity": grouped,
            "counts": counts,
        }

    blockers = sum(
        len(result["issues_by_severity"]["blocker"])
        for result in stage_results.values()
        if isinstance(result.get("issues_by_severity"), dict)
    )
    warnings = sum(
        len(result["issues_by_severity"]["warning"])
        for result in stage_results.values()
        if isinstance(result.get("issues_by_severity"), dict)
    )
    result = {
        "task_dir": str(task_dir),
        "expected_pages": expected,
        "deliverable": blockers == 0 and all(stage_results[s]["ok"] for s in ["pipeline", "delivery"]),
        "blockers": blockers,
        "warnings": warnings,
        "stages": stage_results,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0 if result["deliverable"] else 1


def cmd_check(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")

    stage = args.stage
    issues: list[str] = []
    counts: dict[str, object] = {}

    if stage in {"analysis", "image2", "html", "pptx", "all"}:
        check_analysis_files(task_dir, expected, issues)
    if stage in {"image2", "html", "pptx", "all"}:
        counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
    if stage in {"html", "pptx", "all"}:
        counts["spec_count"] = check_spec_files(task_dir, expected, issues)
        counts["visual_contract_count"] = check_visual_contract_files(task_dir, expected, issues)
    if stage == "html" or (stage in {"pptx", "all"} and commercial_render_path(task_dir) == "html"):
        counts["html_slide_count"] = check_html_files(task_dir, expected, issues)
        counts["html_reference_fidelity"] = check_html_reference_fidelity(task_dir, expected, issues)
    if stage in {"pptx", "all"}:
        pptx = Path(args.pptx).resolve() if args.pptx else None
        if pptx is None:
            pptx_files = sorted((task_dir / "pptx").glob("*.pptx"))
            pptx = pptx_files[-1].resolve() if pptx_files else task_dir / "pptx" / "missing.pptx"
        counts["commercial_render_contract"] = check_commercial_render_contract(task_dir, expected, pptx, issues)
        pptx_summary = check_pptx_file(pptx, expected, issues)
        if pptx_summary is not None:
            counts["pptx"] = str(pptx)
            counts["pptx_summary"] = pptx_summary
        if commercial_render_path(task_dir) == "html":
            counts["render_manifest"] = check_render_manifest(task_dir, pptx, issues)
        else:
            counts["render_manifest"] = "not_required_for_direct_pptx"
        counts["commercial_similarity"] = check_commercial_similarity(task_dir, expected, issues)
    if stage in {"delivery"}:
        counts.update(check_delivery_files(task_dir, expected, issues))
    if stage in {"pipeline"}:
        pptx = Path(args.pptx).resolve() if args.pptx else None
        counts.update(check_pipeline_contract(task_dir, expected, pptx, issues))

    result = {
        "task_dir": str(task_dir),
        "stage": stage,
        "expected_pages": expected,
        "ok": not issues,
        "issues": issues,
        **counts,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0 if not issues else 1


def task_stage_issues(
    task_dir: Path,
    expected: int,
    stage: str,
    pptx: Path | None = None,
) -> tuple[list[str], dict[str, object]]:
    issues: list[str] = []
    counts: dict[str, object] = {}
    if stage == "analysis":
        check_analysis_files(task_dir, expected, issues)
    elif stage == "image2":
        check_analysis_files(task_dir, expected, issues)
        counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
    elif stage == "memory_boundary":
        check_analysis_files(task_dir, expected, issues)
        counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
        issues.extend(post_image_memory_boundary_issues(task_dir, expected))
    elif stage == "html":
        check_analysis_files(task_dir, expected, issues)
        counts["image2_reference_count"] = check_image2_files(task_dir, expected, issues)
        counts["spec_count"] = check_spec_files(task_dir, expected, issues)
        counts["visual_contract_count"] = check_visual_contract_files(task_dir, expected, issues)
        if commercial_render_path(task_dir) == "html":
            counts["html_slide_count"] = check_html_files(task_dir, expected, issues)
            counts["html_reference_fidelity"] = check_html_reference_fidelity(task_dir, expected, issues)
        else:
            counts["html_slide_count"] = "debug_optional"
            counts["html_reference_fidelity"] = "not_required_for_direct_pptx"
    elif stage == "pptx":
        if pptx is None:
            pptx_files = [
                p for p in sorted((task_dir / "pptx").glob("*.pptx"))
                if p.is_file() and not p.name.startswith("~$")
            ]
            pptx = pptx_files[-1].resolve() if pptx_files else task_dir / "pptx" / "missing.pptx"
        counts.update(check_pipeline_contract(task_dir, expected, pptx, issues))
    elif stage == "delivery":
        counts.update(check_delivery_files(task_dir, expected, issues))
    else:
        issues.append(f"Unknown task stage: {stage}")
    return issues, counts


def task_controller_status(
    task_dir: Path,
    expected: int,
    pptx: Path | None = None,
) -> dict[str, object]:
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
            "next_action": "Run forget-after-image2, then record fresh image observations from the selected images only.",
        },
        {
            "id": "html",
            "label": "Image-derived reconstruction source",
            "next_action": "Record observations, extract real visual measurements/contracts, then build the declared commercial render source from the selected images.",
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
        issues, counts = task_stage_issues(task_dir, expected, str(stage["id"]), pptx)
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


def copy_runtime_sources(task_dir: Path, sources: list[str]) -> list[dict[str, object]]:
    copied: list[dict[str, object]] = []
    source_dir = task_dir / "source"
    source_dir.mkdir(parents=True, exist_ok=True)
    for raw in sources:
        src = Path(raw).expanduser().resolve()
        if not src.exists():
            copied.append({"source": str(src), "ok": False, "error": "missing"})
            continue
        if src.is_dir():
            target = source_dir / src.name
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(src, target)
            copied.append({"source": str(src), "target": str(target), "ok": True, "type": "directory"})
        else:
            target = source_dir / src.name
            shutil.copy2(src, target)
            copied.append(
                {
                    "source": str(src),
                    "target": str(target),
                    "ok": True,
                    "type": "file",
                    "sha256": sha256_file(target),
                }
            )
    return copied


def write_public_runtime_state(task_dir: Path, state: dict[str, object]) -> None:
    runtime_dir = task_dir / "qa"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    (runtime_dir / "public_runtime_state.json").write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def slide_story_exists(task_dir: Path) -> bool:
    path = task_dir / "analysis" / "slide_story.json"
    if not path.exists():
        return False
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return False
    slides = data.get("slides") if isinstance(data, dict) else data
    return isinstance(slides, list) and bool(slides)


def maybe_run_prompt_plan(task_dir: Path, expected: int, topic: str) -> dict[str, object] | None:
    plan_path = prompt_selection_plan_path(task_dir)
    if plan_path.exists() or not slide_story_exists(task_dir):
        return None
    args = argparse.Namespace(task_dir=str(task_dir), pages=expected, topic=topic)
    stdout = io.StringIO()
    with contextlib.redirect_stdout(stdout):
        rc = cmd_plan_prompts(args)
    return {
        "action": "plan-prompts",
        "return_code": rc,
        "path": str(plan_path),
        "stdout_captured": bool(stdout.getvalue().strip()),
    }


def maybe_prepare_image2_requests(task_dir: Path) -> dict[str, object] | None:
    manifest = task_dir / "image2" / "image2_generation_manifest.json"
    if manifest.exists():
        return None
    expected = expected_pages_from_task(task_dir)
    if not expected:
        return None
    issues: list[str] = []
    check_analysis_files(task_dir, expected, issues)
    if issues:
        return None
    stdout = io.StringIO()
    with contextlib.redirect_stdout(stdout):
        rc = cmd_prepare_image2_prompts(argparse.Namespace(task_dir=str(task_dir)))
    return {
        "action": "prepare-image2-prompts",
        "return_code": rc,
        "path": str(manifest),
        "stdout_captured": bool(stdout.getvalue().strip()),
    }


def public_runtime_next_action(task_dir: Path, status: dict[str, object]) -> dict[str, object]:
    blocked = status.get("blocked") if isinstance(status.get("blocked"), dict) else None
    if not blocked:
        return {
            "type": "deliverable",
            "message": "Final delivery gate is complete. Reply only with files under delivery/.",
        }
    stage = str(blocked.get("stage", ""))
    expected = int(status.get("expected_pages", 0) or 0)
    if stage == "analysis":
        return {
            "type": "agent_required",
            "stage": "analysis",
            "message": (
                "Read source files, write analysis_report.md and slide_story.json, rerun make-deck "
                "so the runtime can select templates, then write final_prompt_XX.md from the selected plan."
            ),
            "must_create": [
                "analysis/analysis_report.md",
                "analysis/slide_story.json",
                "analysis/prompt_selection_audit.md",
                *[f"analysis/final_prompt_{idx:02d}.md" for idx in range(1, expected + 1)],
            ],
            "continue_command": f"python3 scripts/paopao_run.py make-deck --task-dir {task_dir}",
            "forbidden": [
                "Do not write HTML or PPTX before analysis check passes.",
                "Do not show prompt Markdown files to the user.",
            ],
        }
    if stage == "image2":
        requests = [f"image2/generation_request_{idx:02d}.json" for idx in range(1, expected + 1)]
        register_commands = [
            (
                "python3 scripts/paopao_run.py register-image2-reference "
                f"--task-dir {task_dir} --slide {idx} "
                f"--image <generated_image_{idx:02d}_path_outside_task_dir> "
                f"--generation-request {task_dir / 'image2' / f'generation_request_{idx:02d}.json'} "
                f"--generated-prompt-sha256 <sha256_of_image2_prompt_{idx:02d}.md> "
                "--source image_gen_builtin --tool-call-id <image_generation_artifact_id>"
            )
            for idx in range(1, expected + 1)
        ]
        return {
            "type": "image_generation_required",
            "stage": "image2",
            "message": (
                "Generate exactly one reference image per slide from generation_request_XX.json "
                "prompt_text, register each image, then record style review and ask the user to approve previews."
            ),
            "generation_requests": requests,
            "register_commands": register_commands,
            "continue_command": f"python3 scripts/paopao_run.py make-deck --task-dir {task_dir}",
            "forbidden": [
                "Do not summarize or rewrite prompt_text for image generation.",
                "Do not use HTML/PPTX/browser previews as Image2 references.",
                "Do not continue to reconstruction before user approval is recorded.",
            ],
        }
    if stage == "memory_boundary":
        return {
            "type": "agent_required",
            "stage": "memory_boundary",
            "message": (
                "Run forget-after-image2 after user approval, then reopen each selected image and "
                "record fresh observations from the image only."
            ),
            "commands": [
                f"python3 scripts/paopao_run.py forget-after-image2 --task-dir {task_dir}",
                *[
                    (
                        "python3 scripts/paopao_run.py record-image2-observation "
                        f"--task-dir {task_dir} --slide {idx} --evidence <fresh visual observation>"
                    )
                    for idx in range(1, expected + 1)
                ],
            ],
            "continue_command": f"python3 scripts/paopao_run.py make-deck --task-dir {task_dir}",
        }
    if stage == "html":
        return {
            "type": "agent_required",
            "stage": "reconstruction",
            "message": (
                "Extract image-derived contracts, write measurement/spec files, declare html or direct_pptx, "
                "then build the declared editable source from the selected images only."
            ),
            "commands": [
                *[
                    f"python3 scripts/paopao_run.py extract-image2-contract --task-dir {task_dir} --slide {idx}"
                    for idx in range(1, expected + 1)
                ],
                (
                    "python3 scripts/paopao_run.py record-commercial-render "
                    f"--task-dir {task_dir} --render-path <html|direct_pptx> --pptx <final_pptx_path>"
                ),
            ],
            "continue_command": f"python3 scripts/paopao_run.py make-deck --task-dir {task_dir}",
            "forbidden": [
                "Do not use final_prompt, image2_prompt, analysis_report, or remembered intent as visual inputs.",
                "Do not render whole-slide images as PPT backgrounds.",
            ],
        }
    if stage == "pptx":
        return {
            "type": "agent_required",
            "stage": "pptx_qa",
            "message": (
                "Render the declared editable PPTX, open the real PPTX in PowerPoint, record PowerPoint "
                "and fidelity reviews, then rerun make-deck."
            ),
            "continue_command": f"python3 scripts/paopao_run.py make-deck --task-dir {task_dir}",
        }
    if stage == "delivery":
        return {
            "type": "finalize_required",
            "stage": "delivery",
            "message": "Run finalize-delivery; only files under delivery/ are user-facing.",
            "command": f"python3 scripts/paopao_run.py finalize-delivery --task-dir {task_dir} --pptx <final_pptx_path>",
        }
    return {
        "type": "blocked",
        "stage": stage,
        "message": str(blocked.get("next_action", "")),
        "continue_command": f"python3 scripts/paopao_run.py make-deck --task-dir {task_dir}",
    }


def cmd_make_deck(args: argparse.Namespace) -> int:
    if args.task_dir:
        task_dir = Path(args.task_dir).expanduser().resolve()
        if not task_dir.exists():
            raise SystemExit(f"Task directory does not exist: {task_dir}")
    else:
        if not args.name:
            raise SystemExit("make-deck requires --name when --task-dir is not supplied.")
        if not args.pages:
            raise SystemExit("make-deck requires --pages when creating a new task.")
        task_dir = create_task_dir(
            name=args.name,
            output_root=args.output_root,
            pages=args.pages,
            language=args.language,
            focus=args.focus,
        )

    copied_sources = copy_runtime_sources(task_dir, args.source or [])
    if any(not item.get("ok") for item in copied_sources):
        state = {
            "schema": "paopao.public_runtime_state.v1",
            "task_dir": str(task_dir),
            "ok": False,
            "blocked": {
                "stage": "source",
                "issues": copied_sources,
            },
        }
        write_public_runtime_state(task_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 1

    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize the task with --pages.")
    enforce_public_page_limit(expected)

    automatic_actions: list[dict[str, object]] = []
    planned = maybe_run_prompt_plan(task_dir, expected, args.topic or args.focus or task_dir.name)
    if planned:
        automatic_actions.append(planned)
    prepared = maybe_prepare_image2_requests(task_dir)
    if prepared:
        automatic_actions.append(prepared)

    pptx = Path(args.pptx).expanduser().resolve() if args.pptx else None
    status = task_controller_status(task_dir, expected, pptx)
    next_action = public_runtime_next_action(task_dir, status)
    state = {
        "schema": "paopao.public_runtime_state.v1",
        "runtime": "public_make_deck",
        "task_dir": str(task_dir),
        "expected_pages": expected,
        "public_limits": {
            "max_slides": free_max_slides(),
            "prompt_templates": len(load_prompt_catalog()),
        },
        "copied_sources": copied_sources,
        "automatic_actions": automatic_actions,
        "status": status,
        "next_action": next_action,
        "user_visible_output_policy": (
            "Only delivery/ is user-facing. Prompt, analysis, spec, image2 request, and QA files are internal."
        ),
    }
    write_public_runtime_state(task_dir, state)
    print(json.dumps(state, ensure_ascii=False, indent=2))
    return 0 if status.get("deliverable") else 2


def cmd_run_task(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = args.pages or expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Pass --pages or initialize task with --pages.")
    pptx = Path(args.pptx).resolve() if args.pptx else None
    status = task_controller_status(task_dir, expected, pptx)
    out = task_dir / "qa" / "task_controller_status.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(status, indent=2, ensure_ascii=False))
    return 0 if status["deliverable"] else 1


def cmd_cleanup(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    keep_private = bool(args.keep_private_prompts or os.getenv(PROMPT_ARCHIVE_ENV) == "1")
    moved: list[dict[str, str]] = []
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
            moved.append({
                "from": str(src_rel),
                "to": str(dest.relative_to(task_dir)),
            })
        else:
            path.unlink()
            deleted_prompts.append(str(src_rel))
    if not keep_private and private_root.exists():
        shutil.rmtree(private_root)
    removed_temp: list[str] = []
    for path in delivery_temp_files(task_dir):
        removed_temp.append(str(path.relative_to(task_dir)))
        path.unlink()

    manifest = {
        "policy": "prompt Markdown artifacts are not kept in user task output by default",
        "kept_private_prompts": keep_private,
        "deleted_prompt_markdown_files": deleted_prompts,
        "moved_prompt_markdown_files": moved,
        "removed_temporary_files": removed_temp,
        "note": (
            "Set --keep-private-prompts or PAOPAO_KEEP_PRIVATE_PROMPTS=1 only for local debugging; "
            "normal user output must not contain final_prompt or image2_prompt files."
        ),
    }
    out = task_dir / "qa" / "delivery_cleanup.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


def cmd_publish_delivery(args: argparse.Namespace) -> int:
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
    images_dir = delivery_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)
    html_delivery_dir = delivery_dir / "html"
    html_delivery_dir.mkdir(parents=True, exist_ok=True)
    render_path = commercial_render_path(task_dir)

    copied_images: list[str] = []
    copied_html: list[str] = []
    if expected:
        for idx in range(1, expected + 1):
            image_src = image2_reference_path(task_dir, idx)
            if not image_src.exists():
                raise SystemExit(f"Selected slide image missing: {image_src}")
            image_dest = images_dir / f"slide{idx:02d}{image_src.suffix.lower()}"
            shutil.copy2(image_src, image_dest)
            copied_images.append(str(image_dest.relative_to(delivery_dir)))

            html_src = task_dir / "html" / f"slide{idx:02d}.html"
            if html_src.exists():
                html_dest = html_delivery_dir / html_src.name
                shutil.copy2(html_src, html_dest)
                copied_html.append(str(html_dest.relative_to(delivery_dir)))
            elif render_path == "html":
                raise SystemExit(f"HTML slide missing: {html_src}")

    html_assets = task_dir / "html" / "assets"
    copied_assets: list[str] = []
    if html_assets.exists():
        assets_dest = html_delivery_dir / "assets"
        assets_dest.mkdir(parents=True, exist_ok=True)
        for asset in sorted(p for p in html_assets.rglob("*") if p.is_file()):
            rel = asset.relative_to(html_assets)
            if delivery_forbidden_user_file(Path("html") / "assets" / rel):
                continue
            target_asset = assets_dest / rel
            target_asset.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(asset, target_asset)
            copied_assets.append(str(target_asset.relative_to(delivery_dir)))

    manifest = {
        "published_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "source_pptx": str(pptx),
        "delivery_pptx": str(target),
        "pptx_sha256": sha256_file(target),
        "delivery_images": copied_images,
        "delivery_html": copied_html,
        "delivery_html_assets": copied_assets,
        "policy": (
            "user-facing delivery contains only PPTX, slide images, and optional HTML/assets; "
            "prompt, analysis, spec, QA, and other internal files are excluded"
        ),
    }
    manifest_path = task_dir / "qa" / "delivery_publish_manifest.json"
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


def cmd_finalize_delivery(args: argparse.Namespace) -> int:
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

    cleanup_args = argparse.Namespace(
        task_dir=str(task_dir),
        keep_private_prompts=bool(args.keep_private_prompts),
    )
    cmd_cleanup(cleanup_args)

    publish_args = argparse.Namespace(
        task_dir=str(task_dir),
        pptx=str(pptx),
        output_dir="",
    )
    cmd_publish_delivery(publish_args)

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
        "delivery_dir": str((task_dir / "delivery").resolve()),
        "counts": final_counts,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0 if not final_issues else 1


def cmd_clean_icon_crop(args: argparse.Namespace) -> int:
    try:
        box = _parse_box(args.box)
        result = clean_icon_crop_image(
            Path(args.image).resolve(),
            box,
            Path(args.output).resolve(),
            expand=args.expand,
            padding=args.padding,
            threshold=args.threshold,
            min_canvas=args.min_canvas,
        )
    except Exception as exc:
        print(json.dumps({"ok": False, "issue": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def cmd_record_commercial_render(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    expected = expected_pages_from_task(task_dir)
    if not expected:
        raise SystemExit("Missing expected page count. Initialize task with --pages before recording commercial render.")
    render_path = str(args.render_path).strip()
    if render_path not in COMMERCIAL_RENDER_PATHS:
        raise SystemExit("--render-path must be html or direct_pptx")
    pptx = Path(args.pptx).resolve()
    if not pptx.exists() or pptx.suffix.lower() != ".pptx":
        raise SystemExit(f"PPTX missing or invalid: {pptx}")
    contract = {
        "schema": COMMERCIAL_RENDER_CONTRACT_SCHEMA,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "expected_pages": expected,
        "render_path": render_path,
        "source_of_truth": "image2_reference",
        "post_image_inputs_only": True,
        "commercial_similarity_min": COMMERCIAL_SIMILARITY_MIN,
        "pptx_path": _relative_to_task_or_abs(task_dir, pptx),
        "pptx_sha256": sha256_file(pptx),
        "actual_preview_dir": "qa/pptx_actual",
        "html_is_debug_only": render_path == "direct_pptx",
        "policy": (
            "Commercial delivery is judged by selected Image2 references versus real PowerPoint previews. "
            "HTML is a production path only when render_path=html; otherwise it is debug-only."
        ),
    }
    out = commercial_render_contract_path(task_dir)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(contract, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "contract": str(out), "render_path": render_path}, ensure_ascii=False, indent=2))
    return 0


def cmd_render(args: argparse.Namespace) -> int:
    task_dir = Path(args.task_dir).resolve()
    pptx = Path(args.pptx).resolve()
    html_files = [Path(p).resolve() for p in args.html] if args.html else html_files_from_task(task_dir)
    expected = expected_pages_from_task(task_dir) or len(html_files)
    preflight_issues: list[str] = []
    check_image2_files(task_dir, expected, preflight_issues)
    check_spec_files(task_dir, expected, preflight_issues)
    check_visual_contract_files(task_dir, expected, preflight_issues)
    check_html_files(task_dir, expected, preflight_issues)
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

    cmd = [sys.executable, str(RENDERER), *map(str, html_files), "--pptx", str(pptx)]
    if args.pdf:
        cmd.extend(["--pdf", str(Path(args.pdf).resolve())])

    env = None
    proc = subprocess.run(cmd, cwd=str(PLUGIN_ROOT), text=True, capture_output=True, env=env)
    if proc.stdout:
        print(proc.stdout)
    if proc.stderr:
        print(proc.stderr, file=sys.stderr)
    succeeded = proc.returncode == 0 and pptx.exists()
    finish_quota(reservation_id, succeeded)
    if succeeded:
        manifest = {
            "rendered_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
            "renderer": str(RENDERER.resolve()),
            "task_dir": str(task_dir),
            "pptx_path": str(pptx),
            "pptx_sha256": sha256_file(pptx),
            "html": [
                {
                    "path": str(path),
                    "sha256": sha256_file(path),
                }
                for path in html_files
            ],
        }
        out = render_manifest_path(task_dir)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return proc.returncode


def cmd_package(args: argparse.Namespace) -> int:
    src = PLUGIN_ROOT
    out = Path(args.output).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()
    base = out.parent / out.name.removesuffix(".zip")
    shutil.make_archive(str(base), "zip", root_dir=str(src.parent), base_dir=src.name)
    built = Path(str(base) + ".zip")
    if built != out:
        built.replace(out)
    print(out)
    return 0


def cmd_doctor(_: argparse.Namespace) -> int:
    modules = ["playwright", "pptx", "lxml"]
    module_checks = {
        name: importlib.util.find_spec(name) is not None
        for name in modules
    }
    chromium_hint = (
        "If Playwright browsers are missing, run: python3 -m playwright install chromium"
    )
    checks = {
        "plugin_root": str(PLUGIN_ROOT),
        "renderer_exists": RENDERER.exists(),
        "prompts_exists": (PLUGIN_ROOT / "prompts").exists(),
        "renderer_guide_exists": (PLUGIN_ROOT / "reference" / "renderer_guide.md").exists(),
        "python_modules": module_checks,
        "powerpoint_qa": "Open the generated PPTX in PowerPoint for final visual QA.",
        "chromium_hint": chromium_hint,
    }
    print(json.dumps(checks, indent=2, ensure_ascii=False))
    required_files_ok = all(v for k, v in checks.items() if k.endswith("_exists"))
    modules_ok = all(module_checks.values())
    return 0 if required_files_ok and modules_ok else 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="paopao helper")
    sub = parser.add_subparsers(dest="command", required=True)

    init = sub.add_parser("init", help="Create a Paopao task folder")
    init.add_argument("--name", required=True)
    init.add_argument("--output-root", default="output")
    init.add_argument("--pages", type=int, default=None)
    init.add_argument("--language", default="")
    init.add_argument("--focus", default="")
    init.set_defaults(func=cmd_init)

    make_deck = sub.add_parser(
        "make-deck",
        help="Public runtime controller: create or continue a deck task without skipping required gates",
    )
    make_deck.add_argument("--task-dir", default="", help="Existing task directory to continue")
    make_deck.add_argument("--name", default="", help="Task name when creating a new task")
    make_deck.add_argument("--output-root", default="output")
    make_deck.add_argument("--source", action="append", default=[], help="Source file or folder to copy into task/source")
    make_deck.add_argument("--pages", type=int, default=None)
    make_deck.add_argument("--language", default="")
    make_deck.add_argument("--focus", default="")
    make_deck.add_argument("--topic", default="", help="Deck topic for deterministic prompt-template planning")
    make_deck.add_argument("--pptx", default="", help="Final PPTX path when continuing late-stage QA/finalization")
    make_deck.set_defaults(func=cmd_make_deck)

    plan_prompts = sub.add_parser(
        "plan-prompts",
        help="Select prompt-library templates for every slide before final_prompt_XX.md is written",
    )
    plan_prompts.add_argument("--task-dir", required=True)
    plan_prompts.add_argument("--pages", type=int, default=None)
    plan_prompts.add_argument("--topic", default="")
    plan_prompts.set_defaults(func=cmd_plan_prompts)

    render = sub.add_parser("render", help="Render task HTML slides to PPTX")
    render.add_argument("--task-dir", required=True)
    render.add_argument("--pptx", required=True)
    render.add_argument("--pdf", default="")
    render.add_argument("html", nargs="*")
    render.set_defaults(func=cmd_render)

    image2 = sub.add_parser(
        "prepare-image2-prompts",
        help="Build locked per-slide Image2 prompt files and provenance manifest",
    )
    image2.add_argument("--task-dir", required=True)
    image2.set_defaults(func=cmd_prepare_image2_prompts)

    user_review = sub.add_parser(
        "record-image2-user-review",
        help="Record user approval or requested changes after showing selected Image2 references",
    )
    user_review.add_argument("--task-dir", required=True)
    user_review.add_argument("--pages", type=int, default=None)
    user_review.add_argument(
        "--approved",
        required=True,
        choices=["yes", "no", "true", "false", "approved", "changes_requested"],
    )
    user_review.add_argument("--feedback", required=True)
    user_review.set_defaults(func=cmd_record_image2_user_review)

    forget = sub.add_parser(
        "forget-after-image2",
        help="Lock the post-image memory boundary so reconstruction can use selected images only",
    )
    forget.add_argument("--task-dir", required=True)
    forget.add_argument("--pages", type=int, default=None)
    forget.set_defaults(func=cmd_forget_after_image2)

    observation = sub.add_parser(
        "record-image2-observation",
        help="Record the fresh visual observation used as the only source for post-approval reconstruction",
    )
    observation.add_argument("--task-dir", required=True)
    observation.add_argument("--slide", type=int, required=True)
    observation.add_argument(
        "--evidence",
        required=True,
        help="Concrete visual observation from reopening the selected image; at least 120 characters.",
    )
    observation.set_defaults(func=cmd_record_image2_observation)

    extract_contract = sub.add_parser(
        "extract-image2-contract",
        help="Generate an initial visual contract from the selected Image2 reference pixels",
    )
    extract_contract.add_argument("--task-dir", required=True)
    extract_contract.add_argument("--slide", type=int, required=True)
    extract_contract.add_argument(
        "--force",
        action="store_true",
        help="Overwrite an existing slideXX_visual_contract.json.",
    )
    extract_contract.set_defaults(func=cmd_extract_image2_contract)

    register_image2 = sub.add_parser(
        "register-image2-reference",
        help="Register one generated Image2 reference with locked prompt-sha provenance",
    )
    register_image2.add_argument("--task-dir", required=True)
    register_image2.add_argument("--slide", type=int, required=True)
    register_image2.add_argument("--image", required=True)
    register_image2.add_argument(
        "--generation-request",
        required=True,
        help="Locked generation_request_XX.json whose prompt_text was used exactly for image generation.",
    )
    register_image2.add_argument("--generated-prompt-sha256", required=True)
    register_image2.add_argument(
        "--source",
        required=True,
        choices=sorted(IMAGE2_ALLOWED_SOURCE_KINDS),
        help="Provenance source for the selected reference; local screenshots/previews are rejected.",
    )
    register_image2.add_argument(
        "--tool-call-id",
        required=True,
        help="Non-empty image generation tool call, job, or artifact id used for provenance audit.",
    )
    register_image2.set_defaults(func=cmd_register_image2_reference)

    check = sub.add_parser("check", help="Validate Paopao pipeline stage invariants")
    check.add_argument("--task-dir", required=True)
    check.add_argument("--pages", type=int, default=None)
    check.add_argument(
        "--stage",
        choices=["analysis", "image2", "html", "pptx", "all", "pipeline", "delivery"],
        default="all",
        help="Pipeline stage to validate. Later stages include earlier checks.",
    )
    check.add_argument("--pptx", default="", help="Optional PPTX path for stage=pptx/all")
    check.set_defaults(func=cmd_check)

    run_task = sub.add_parser(
        "run-task",
        help="Report the single allowed next step for a Paopao task; final delivery still requires finalize-delivery",
    )
    run_task.add_argument("--task-dir", required=True)
    run_task.add_argument("--pages", type=int, default=None)
    run_task.add_argument("--pptx", default="")
    run_task.set_defaults(func=cmd_run_task)

    audit = sub.add_parser("audit-task", help="Run a full task audit and summarize blockers before delivery")
    audit.add_argument("--task-dir", required=True)
    audit.add_argument("--pages", type=int, default=None)
    audit.add_argument("--pptx", default="", help="Optional PPTX path to audit")
    audit.set_defaults(func=cmd_audit_task)

    cleanup = sub.add_parser("cleanup-delivery", help="Remove prompt Markdown artifacts before delivery")
    cleanup.add_argument("--task-dir", required=True)
    cleanup.add_argument(
        "--keep-private-prompts",
        action="store_true",
        help="Debug only: move prompt artifacts under qa/private_prompts instead of deleting them from task output",
    )
    cleanup.set_defaults(func=cmd_cleanup)

    publish = sub.add_parser("publish-delivery", help="Publish user-facing PPTX, slide images, and HTML to delivery/")
    publish.add_argument("--task-dir", required=True)
    publish.add_argument("--pptx", default="")
    publish.add_argument("--output-dir", default="")
    publish.set_defaults(func=cmd_publish_delivery)

    finalize = sub.add_parser(
        "finalize-delivery",
        help="Run final pipeline gate, cleanup, publish, and delivery gate in one required release step",
    )
    finalize.add_argument("--task-dir", required=True)
    finalize.add_argument("--pptx", default="")
    finalize.add_argument(
        "--keep-private-prompts",
        action="store_true",
        help="Debug only: move prompt artifacts under qa/private_prompts instead of deleting them",
    )
    finalize.set_defaults(func=cmd_finalize_delivery)

    clean_icon = sub.add_parser(
        "clean-icon-crop",
        help="Crop an icon from a reference image, remove detected corner background, and export a transparent PNG.",
    )
    clean_icon.add_argument("--image", required=True, help="Source reference image path")
    clean_icon.add_argument("--box", required=True, help="Crop box as x,y,w,h in source image pixels")
    clean_icon.add_argument("--output", required=True, help="Output PNG path, usually output/<task>/html/assets/<name>.png")
    clean_icon.add_argument("--expand", type=int, default=14, help="Pixels to expand around the supplied box before cleanup")
    clean_icon.add_argument("--padding", type=int, default=10, help="Transparent padding around the cleaned icon")
    clean_icon.add_argument("--threshold", type=int, default=34, help="RGB distance threshold for removing detected corner background")
    clean_icon.add_argument("--min-canvas", type=int, default=96, help="Minimum square output canvas size")
    clean_icon.set_defaults(func=cmd_clean_icon_crop)

    commercial_render = sub.add_parser(
        "record-commercial-render",
        help="Declare whether the commercial PPTX was produced through html or direct_pptx and bind its hash",
    )
    commercial_render.add_argument("--task-dir", required=True)
    commercial_render.add_argument("--render-path", required=True, choices=sorted(COMMERCIAL_RENDER_PATHS))
    commercial_render.add_argument("--pptx", required=True)
    commercial_render.set_defaults(func=cmd_record_commercial_render)

    package = sub.add_parser("package", help="Build a zip package of this plugin")
    package.add_argument("--output", required=True)
    package.set_defaults(func=cmd_package)

    doctor = sub.add_parser("doctor", help="Check local plugin files")
    doctor.set_defaults(func=cmd_doctor)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())

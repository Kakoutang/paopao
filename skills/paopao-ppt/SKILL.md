---
name: "paopao-ppt"
description: "Use Paopao to turn PDFs, reports, papers, and reference images into editable consulting-style PPTX decks. Trigger when the user asks to make, generate, create, or package a PPT/PPTX/slides/deck from source documents, especially when they want editable output rather than image slides."
---

# Paopao PPT

paopao creates editable PowerPoint decks from source documents.

## Hard Rules

- Do not freehand the workflow from memory. At the start of every Paopao task and after each major artifact is created, run `paopao_run.py run-task --task-dir output/<task-name>` and follow the blocked `next_action`. If `run-task` is blocked, do not skip ahead to later stages.
- Final output must be `.pptx`.
- Final PPTX must be editable. Never use a whole-slide screenshot, Image2, PNG, JPG, or PDF as the slide background.
- Image2 visual references are required before any reconstruction source. Do not skip Image2, do not substitute old drafts, and do not write HTML or PPTX from memory.
- Image2 visual references must be landscape 16:9 widescreen, target 1920x1080. Portrait, square, 4:3, and 3:2 references are invalid and must be regenerated.
- Each `final_prompt_XX.md` must be filled from one concrete prompt-library file in `<plugin-root>/prompts/` and declare `PROMPT_TEMPLATE: <file>.md` plus the matching `LAYOUT_NAME`. Handwritten layout prose that merely resembles the style is invalid.
- Prompt templates must be selected through `analysis/slide_story.json` and `plan-prompts` before writing final prompts. Do not manually browse the prompt library and choose by habit; final prompts must match `analysis/prompt_selection_plan.json`.
- Each Image2 reference must be generated from the corresponding template-backed `final_prompt_XX.md` plus the Paopao compact consulting style rules below. Generic image prompts or decorative mockups are invalid.
- Preserve prompt-library layout diversity. The visual style rules control palette, typography, linework, borders, icons, background, and material simplicity; they must not force every slide into one repeated composition.
- The only valid input to the image generation tool is the `prompt_text` field in `output/<task-name>/image2/generation_request_XX.json`, produced by `prepare-image2-prompts`; that field must exactly equal `output/<task-name>/image2/image2_prompt_XX.md`. Do not handwrite, summarize, compress, translate, or improvise image-generation prompts in chat/tool calls.
- After saving selected Image2 references, register each one with `register-image2-reference` and the exact `image2_prompt_XX.md` sha. Do not rerun `prepare-image2-prompts` as a substitute for registration. `check --stage image2`, `check --stage html`, and `render` must fail if locked Image2 provenance is missing, stale, or unregistered.
- After Image2 is generated, treat each generated image as an unfamiliar external reference. Open it again, observe it, record `qa/image2_style_review.json`, then reconstruct only from image-derived measurement.
- After Image2 references are generated and internally reviewed, pause and show the selected page previews to the user. Do not write specs, HTML, or render PPTX until the user explicitly approves the previews and `record-image2-user-review` has recorded that approval. If the user requests changes, regenerate the affected references first, then ask again.
- After user approval and before writing specs, run `forget-after-image2` to lock the post-image memory boundary. From that point on, do not consult or reuse `final_prompt_XX.md`, `image2_prompt_XX.md`, prompt templates, analysis notes, or remembered intent for visual reconstruction. `record-image2-observation` also auto-creates this boundary if it is missing, but the preferred workflow is to lock it explicitly immediately after approval.
- After the memory boundary is locked, run `record-image2-observation` once per slide from fresh visual inspection of the selected image. Specs, visual contracts, HTML, and direct PPTX render plans must bind to that observation record by `observation_id`; do not let them derive directly from prompts, analysis, or memory.
- After recording the fresh observation, run `extract-image2-contract` for each slide before writing or editing the visual contract. This command creates the initial region inventory from the selected image pixels. Manual edits may correct OCR/text labels or refine regions, but the starting geometry must come from this image-extraction step, not from prompt memory.
- After user approval, treat each selected Image2 reference as if it were a brand-new user-supplied screenshot with no production history. Discard prompt, prompt-template, analysis, and layout-library memory for visual reconstruction. The spec, visual contract, and HTML must be authored from fresh observation of the actual image only, not from `final_prompt_XX.md`, `image2_prompt_XX.md`, `analysis_report.md`, or remembered slide intent.
- Each visual contract must declare `reconstruction_source: "image2_reference_only"`, `prompt_context_discarded: true`, `observed_as_fresh_image: true`, `derivation_method: "fresh_visual_observation_record"`, `observation_record_path`, `observation_record_sha256`, `observation_id`, and a concrete `observation_evidence` paragraph. Each spec, HTML page, or direct PPTX render plan must also bind to the same `observation_id`. If post-image artifacts mention prompt templates, final prompts, prompt-selection audits, analysis reports, or image2 prompt files as design inputs, the workflow is invalid.
- The number of selected Image2 references must equal the requested slide count exactly. If the user asked for 4 pages, use exactly 4 `image2_reference_XX.png` files and exactly 4 reconstructed PPTX slides. Extra generated images are discarded and must not become slides.
- Never continue after a failed stage. If analysis, prompt audit, Image2 count, spec count, declared reconstruction source, render, or actual PPTX QA fails, stop, fix, and rerun the failed stage before moving on.
- After Image2 approval, the commercial reconstruction path must be declared in `qa/commercial_render_contract.json` as either `render_path: "html"` or `render_path: "direct_pptx"`. Both paths must use the selected Image2 references and image-derived measurement as the only visual source.
- HTML is a production source only when the commercial render contract says `render_path: "html"`. When the contract says `direct_pptx`, HTML is optional debug output and must not be treated as the source of truth.
- Direct PPTX is valid only when it is rebuilt from the image-derived measurement/visual contract, remains editable, and records the same observation binding as the spec. Do not use python-pptx or any other tool to redesign from prompt/story memory.
- Author the declared reconstruction source one slide at a time from the observed image/spec. Do not generate HTML or PPTX mechanically from a generic script or from remembered prompt intent.
- If using the HTML path, the HTML-stage check captures each HTML slide as an image and compares it against the selected Image2 reference before PPTX rendering. If HTML/reference similarity is below the configured threshold, rewrite the HTML from the image-derived contract; do not render PPTX and discover the drift later.
- If the PPTX is visibly unlike the Image2 reference, refill the failed source layer: reopen the Image2 reference, repair the measurement/visual contract, fix the declared reconstruction source, rerender the PPTX, and rerun QA. If the reference overuses decorative icons, regenerate the reference with fewer, more natural icons before reconstruction. The final delivery PPTX must be produced through the current declared commercial path and pass the real-PowerPoint similarity gate.
- For QA, do not use the renderer's HTML PDF preview as the final proof. Open the actual PPTX in PowerPoint and inspect it visually. Do not require LibreOffice, PDF conversion, or PNG export for normal QA.
- Navigation bars must use `<nav>` or `.nav/.navbar/.navigation/.tabs/.tabbar/.breadcrumb`; tab elements must use `.tab` or a semantic equivalent. This lets the renderer force nav text vertical centering in PPTX.
- Navigation is deck-level chrome, not a per-slide optional design choice. Every slide must use the same visible top navigation strip with the same tab count/order/labels and only the active tab changing. The locked Image2 prompt adds this deck navigation contract before reference generation; do not remove it because a selected slide template is dense, dashboard-like, or cover-like.
- Icons visible in the reference must remain visible in PPTX, either as small `data-pptx-image` assets or recognizable editable shapes.
- Icons are allowed, but never automatic decoration. Prefer text labels, numbered badges, thin dividers, status dots, Harvey balls, or KPI markers when they communicate the same meaning. Use simple editable line icons only when they add clear semantic value or the selected layout requires them; use small preserved image assets only when a necessary icon cannot be recreated cleanly. Icons must be sparse, simple, monochrome blue/grey, and visually similar to the reference. Do not add icons to every band, card, bullet, nav item, or takeaway by default.
- Short uppercase text inside an icon box (`ECO`, `SYS`, `BANK`, `PROD`, `NET`, `UP`, etc.) is an invalid icon placeholder unless the reference itself is a letter/number badge. If an icon cannot be cleanly rebuilt, crop that small icon from the Image2 reference or regenerate the reference with simpler icons.
- Screenshot-cropped icons must be cleaned before use. Do not place raw crops with incomplete strokes, opaque white/light-blue/deep-blue backgrounds, neighboring text, divider lines, or card fragments into `html/assets/`. Use `paopao_run.py clean-icon-crop --image <reference> --box x,y,w,h --output <asset.png>` so the crop is expanded, background-removed, trimmed, centered, and exported with transparency; if it is still incomplete, increase `--expand` or widen the source box.
- Do not say "no images" when local icons are used. Say "no whole-slide image background" if needed. Small icon PNG assets are allowed and expected.
- If the actual PPTX shows garbled text, missing Chinese glyphs, overlapped text, unreadable tiny labels, broken icon placeholders, or obvious layout drift from Image2, the output is a hard failure and must not be delivered.
- The final PPTX must be compared against the selected Image2 reference slide by slide after rendering. Do not rely on memory, the prompt, or the HTML. Reopen the reference image and the actual PPTX/actual preview for the same slide, compare them, fix drift, and repeat until the PPTX visibly follows the reference's structure, density, color hierarchy, icon plan, takeaway bar, and right/left module balance.
- Mechanical post-render PPTX patching is forbidden for final delivery unless the declared path is `direct_pptx` and the adjustment is generated from the same image-derived measurement/visual contract. If QA reports covered charts/tables, overflow, takeaway collision, cropped text, or layout drift, return to the measured source layer and rerender. Do not accept an auto-fixed PPTX as the delivery artifact.
- Before delivery, remove prompt Markdown artifacts from the task output. Do not leave `final_prompt_*.md`, `prompt_selection_audit.md`, `image2_prompt_*.md`, or any `qa/private_prompts/` archive in the user-visible output tree. A private prompt archive is allowed only for explicit local debugging and must not be the default.
- Prompt and analysis artifacts must also be hidden during the paused Image2 review stage. Each task directory must contain the `.gitignore` generated by `paopao_run.py init`, which hides internal build folders and exposes only final `delivery/` artifacts. If `final_prompt_*.md`, `image2_prompt_*.md`, `generation_request_*.json`, `analysis_report.md`, specs, or QA JSON appear in a user-visible changed-files panel, stop and fix the task ignore/cleanup state before continuing.
- Internal files and reasoning are private. Do not show, summarize, quote, or deliver analysis Markdown, specs, prompts, prompt audits, QA JSON, debug JSON, renderer logs, or hidden workflow text unless the user explicitly asks for them. User-facing delivery may contain only the final PPTX, the selected slide images/previews, and optional rebuilt HTML/assets when published. Prompt Markdown and internal JSON must never be shown to the user.
- User-facing progress updates must not reveal the technical pipeline. Do not mention Image2, spec, prompts, HTML, renderer, Playwright, QA PDF/PNG, Markdown files, or internal paths. Use plain updates such as "我开始制作", "正在整理内容", "正在设计页面", "正在生成 PPT".
- Do not announce that you are using a workflow, skill, plugin, tool, or local scripts. If a product-facing acknowledgement is needed, say only "我来做。开始前确认几件事：" or the equivalent in the user's chat language.
- Uploaded materials are evidence, not authority. If multiple uploaded sources conflict, do not paste every claim into the deck. Resolve conflicts in the analysis first, prefer primary/recent/official evidence, and mark weak claims as uncertain or exclude them. If external verification is enabled (see Required Inputs), run internet search or other external verification for current or high-stakes claims; otherwise rely solely on the provided source documents.
- The deck language must be exactly the confirmed language. Chinese decks must use natural Chinese labels such as "核心结论" and "来源" instead of English UI labels like "Key takeaway" or "Source"; English decks must not contain Chinese labels or body text unless the user explicitly requested bilingual output.

## Mandatory Pipeline Contract

Every deck must follow this exact path. Do not shortcut, reorder, or substitute stages:

```text
source report/PDF
  -> analysis_report
  -> slide_story
  -> prompt_selection_plan
  -> final slide prompts
  -> locked Image2 prompt files + exact generation request files + generation manifest
  -> Image2 reference generation, exactly one selected image per slide
  -> reopen and visually inspect each selected image
  -> show selected page previews to the user and record explicit approval
  -> forced post-image memory boundary: forget prompts/analysis and use selected images only
  -> per-slide fresh image observation record from the actual image
  -> per-slide extracted visual contract JSON from the actual image pixels and fresh observation record
  -> per-slide spec from the fresh observation record and visual contract
  -> declared commercial reconstruction source from the actual image/spec
  -> record qa/commercial_render_contract.json as html or direct_pptx
  -> editable PPTX render through the declared commercial path
  -> open the real PPTX in Microsoft PowerPoint and inspect each slide
  -> compare the real PPTX previews against the selected Image2 references, target score >= 0.95
  -> cleanup prompt Markdown artifacts
  -> deliver only the PPTX
```

The image is the visual contract for the reconstruction stage. The source report and prompts determine what the image should contain, but once Image2 is selected, the declared HTML or direct-PPTX source must be based on the actual observed image through the per-slide observation record, not on memory, the prompt text, or a guessed layout.

The controller command is the source of truth for the next allowed action:

```bash
python3 <plugin-root>/scripts/paopao_run.py run-task --task-dir output/<task-name>
```

If it reports a blocked stage, fix that stage before doing anything later. Do not manually publish, render, write HTML, or reply with delivery links from agent judgment alone.

The post-approval reconstruction boundary is deliberately strict: after the user approves selected page previews, imagine those images arrived in a new blank thread with no report, no prompts, no selected templates, and no memory of why they were made. Reopen the current image, observe it as an external artifact, record only visible geometry/text/style into the visual contract, then write the spec and declared reconstruction source from that observation. Do not look back at prompt Markdown to decide composition, module count, nav labels, icon plan, table structure, or text placement.

The deck navigation strip is a global visual contract across slides. If an Image2 reference omits the top nav, reject and regenerate it from the locked prompt; do not write `nav` as hidden, zero-height, empty, or `not_visible` in the visual contract or HTML.

User approval of the selected page previews is a hard boundary between visual direction and rebuild. This boundary exists because the next stage must rebuild from actual images, not from prompt memory. If approval is missing, stale, or attached to different image hashes, stop and ask the user to review the current previews again.

The actual PowerPoint file is the final contract for delivery. Browser previews, HTML screenshots, PDF exports, and PNG renders are useful debugging aids, but they cannot replace opening the PPTX in PowerPoint. If PowerPoint shows missing numbers, garbled Chinese, shifted modules, broken icons, or visible drift from the image reference, the deck is not deliverable.

## Refill Loop For Commercial Quality

Do not treat after-the-fact repair as part of the product. Each slide must pass the loop before delivery:

```text
Image2 reference
  -> observed spec
  -> declared reconstruction source
  -> editable PPTX render
  -> actual PPTX visual check
  -> compare with Image2, target score >= 0.95
  -> if drift exists: refill measurement/source and rerender
```

Refill means returning to the failed source layer. If the declared reconstruction source does not look like the image, repair the visual measurement and source. If the Image2 is hard to reproduce or overuses decorative icons, regenerate Image2 with fewer icons and simpler layout instructions. If PowerPoint rendering differs from the declared source, change that source or renderer behavior until the real PPTX preview matches. Do not keep a bad PPTX and call it deliverable.

Common hard-fail drift patterns that require refill:

- Image2 has real icons but the reconstructed PPTX uses text placeholders.
- The reconstruction roughly follows the topic but not the reference geometry.
- PPTX looks different from the declared source because of font shrink, white masks, clipped labels, or PowerPoint text layout.
- QA is performed against PDF/browser previews instead of the real PPTX.
- Multiple draft PPTX files remain next to the delivery file and the user can accidentally open the wrong one.

## Required Inputs

Before reading the source, running commands, or starting work, confirm all three:

- Page count
- Language
- Main focus, use case, or audience

If page count is missing, ask "想做几页？" Do not suggest a default page count.

If language is missing, ask "偏好哪一种语言？" Do not imply only Chinese or English; any language is acceptable.

If focus, use case, audience, or emphasis is missing, ask "有什么想突出的重点、用途或特殊偏好？" Do not infer a default such as investment, management briefing, sales, roadshow, academic, or brand style.

Optionally ask about external research: "是否需要我仅基于你提供的资料制作？（默认只用你的资料，不额外搜索外部信息）" The default is source-only mode (no external search or expansion beyond the provided documents). If the user explicitly requests external research or says "可以搜索", enable external verification. Do not expand, supplement, or search beyond the provided sources unless the user opts in.

Ask only the missing questions. For example, if the user says "用这个 PDF 帮我做 3 页 PPT", page count is known but language and focus are missing, so reply:

```text
我来做。开始前还需要确认两点：
1. 偏好哪一种语言？
2. 有什么想突出的重点、用途或特殊偏好？
```

The chat reply language should follow the user's message language. The deck language is a separate question and must not be inferred from chat language.

Never reply with a detailed technical confirmation such as "I will extract the PDF, create prompts, write HTML, and render PPTX." A short confirmation is enough: "收到，我按这个页数和语言做。"

## Standard Workflow

0. Run a local environment check before serious work:

```bash
python3 <plugin-root>/scripts/paopao_run.py doctor
```

If required local dependencies are missing, stop early and tell the user what to install. Do not run until halfway and fail because Playwright, Chromium, or PowerPoint cannot be used.

1. Create task folder:

```bash
python3 <plugin-root>/scripts/paopao_run.py init --name <task-name>
```

Always include the confirmed inputs when they are known:

```bash
python3 <plugin-root>/scripts/paopao_run.py init \
  --name <task-name> \
  --pages <requested-page-count> \
  --language <requested-language> \
  --focus <confirmed-focus>
```

2. Read the source document and write:

```text
output/<task-name>/analysis/analysis_report.md
output/<task-name>/analysis/slide_story.json
```

Include only facts found in the sources or verified externally. Do not invent data.
The analysis must be substantive, not a placeholder: include source inventory, fact bank, conflict resolution, Codex independent judgment, page story, and known limits. If external verification is enabled, also include cross-validation with external sources. It must contain enough sourced facts to support every requested slide.
`slide_story.json` must contain one entry per requested slide with the slide number, section name, slide role, and concrete brief/claim. This file is the input to deterministic prompt-template selection.

3. Generate the prompt-selection plan:

```bash
python3 <plugin-root>/scripts/paopao_run.py plan-prompts \
  --task-dir output/<task-name> \
  --topic "<deck-topic>"
```

This writes:

```text
output/<task-name>/analysis/prompt_selection_plan.json
```

The plan must record the selected prompt-library template, scaffold family, and at least three candidates per slide. Use the selected template from this plan for each final prompt.

4. Create page claims and final prompts:

```text
output/<task-name>/analysis/prompt_selection_audit.md
output/<task-name>/analysis/final_prompt_01.md
...
```

Use the prompt library file selected in `prompt_selection_plan.json`. Each slide must use the planned layout annotation file; rerun `plan-prompts` if the story changes.
Every `final_prompt_XX.md` must begin with these two lines:

```text
PROMPT_TEMPLATE: <selected-prompt-library-file>.md
LAYOUT_NAME: <the exact LAYOUT_NAME from that file>
```

Each `final_prompt_XX.md` must be a complete filled prompt, not a summary. It must include title, selected layout, zones/sections, exact bullets/data, bottom takeaway, source line, and design line. A file with only a topic, outline, or "create a slide" placeholder is invalid.

Run:

```bash
python3 <plugin-root>/scripts/paopao_run.py check --task-dir output/<task-name> --stage analysis
```

Do not generate Image2 until this check passes analysis, prompt-selection-plan, and prompt completeness.

Then create locked image-generation inputs:

```bash
python3 <plugin-root>/scripts/paopao_run.py prepare-image2-prompts \
  --task-dir output/<task-name>
```

Use the resulting files as the exact prompt text for the image generation tool:

```text
output/<task-name>/image2/image2_prompt_01.md
output/<task-name>/image2/image2_prompt_02.md
output/<task-name>/image2/generation_request_01.json
output/<task-name>/image2/generation_request_02.json
...
```

Do not use a rewritten chat summary, shortened prompt, or ad hoc prompt. The `generation_request_XX.json` file is the system-level handoff for reference-image generation. Its `prompt_text` field must be passed exactly as the image-generation prompt and must match the locked `image2_prompt_XX.md` by sha.

5. Generate exactly one selected Image2 reference per requested slide with the built-in image generation tool, using only the exact `prompt_text` value from `generation_request_XX.json` as the prompt input.

Save selected references to:

```text
output/<task-name>/image2/image2_reference_01.png
output/<task-name>/image2/image2_reference_02.png
...
```

The count of selected references must equal the requested page count. Do not generate or keep an extra selected reference.

After each selected reference is generated, register it from the original image-generation output path with the exact sha of the prompt file used for generation. Do not pre-copy a local screenshot, HTML preview, PPTX preview, or already-renamed `image2_reference_XX.png` into `image2/` and then register it. The register command copies the verified generated source into the selected reference path:

```bash
python3 <plugin-root>/scripts/paopao_run.py register-image2-reference \
  --task-dir output/<task-name> \
  --slide 1 \
  --image <original-generated-image-output-path-outside-task-dir> \
  --generation-request output/<task-name>/image2/generation_request_01.json \
  --generated-prompt-sha256 <sha256-of-image2_prompt_01.md> \
  --source image_gen_builtin \
  --tool-call-id <image-generation-tool-call-or-artifact-id>
```

Repeat for every slide. `prepare-image2-prompts` only locks prompt files; it does not prove an image was generated from them.

Run:

```bash
python3 <plugin-root>/scripts/paopao_run.py check --task-dir output/<task-name> --stage image2
```

If the Image2 reference count does not equal the requested page count, or any reference is present but unregistered, fix the selected references before writing HTML.
Also inspect the Image2 references for the Paopao visual language. Reject and regenerate any reference that is poster-like, sparse, decorative, has complex photo backgrounds, has illegible/garbled text, uses too many colors, has heavy gradients/textures/shadows, uses thick or random extra frames, has messy linework, uses oversized illustrative icons, or cannot be rebuilt mostly with editable PPT elements.

Before writing HTML, create:

```text
output/<task-name>/qa/image2_style_review.json
```

It must record that the actual generated images were opened and compared against the Paopao visual language, with one passing entry per slide. Each entry must include the selected reference path, real image dimensions, concrete observed evidence, checked dimensions including `aspect_ratio_16_9`, `house_style_reference`, `palette_discipline`, `clean_background`, `linework_and_borders`, `material_simplicity`, `title_weight`, `module_density`, `takeaway`, `color_hierarchy`, and `icons`, plus an empty `reject_reasons` list. `check --stage image2` fails without this file.

Before writing specs or HTML, show the selected page previews to the user in the chat. Ask whether the visual direction is approved and which pages need changes. If the user asks for changes, regenerate those page references and repeat the preview step. If the user approves, record the approval:

```bash
python3 <plugin-root>/scripts/paopao_run.py record-image2-user-review \
  --task-dir output/<task-name> \
  --approved yes \
  --feedback "<brief user approval note>"
```

Then rerun:

```bash
python3 <plugin-root>/scripts/paopao_run.py check --task-dir output/<task-name> --stage image2
```

This gate binds the user's approval to the current image hashes. If any reference is regenerated or replaced, the user review becomes stale and must be recorded again. Do not proceed to visual contracts, specs, HTML, or render until this gate passes.

6. Immediately after user approval, lock the post-image memory boundary:

```bash
python3 <plugin-root>/scripts/paopao_run.py forget-after-image2 \
  --task-dir output/<task-name>
```

This writes `qa/post_image_memory_boundary.json`. Every observation record, visual contract, spec, and HTML slide must be created after this file and must treat the selected image as an unfamiliar external screenshot. If this boundary is missing, stale, or older than regenerated references, post-image checks fail.

7. For each slide, open the current Image2 reference again after user approval and first record the fresh image observation:

```bash
python3 <plugin-root>/scripts/paopao_run.py record-image2-observation \
  --task-dir output/<task-name> \
  --slide 1 \
  --evidence "<concrete visual observation from the selected image only>"
```

Then write a machine-checkable visual contract:

```text
output/<task-name>/spec/slide01_visual_contract.json
```

Start by extracting the draft contract from the actual reference image:

```bash
python3 <plugin-root>/scripts/paopao_run.py extract-image2-contract \
  --task-dir output/<task-name> \
  --slide 1
```

Use `--force` only when deliberately replacing a stale or incorrect extracted contract. The extractor is geometry-first: it reads local pixels to identify nav/title/content/detail/takeaway/source regions, but it does not guarantee perfect OCR. After extraction, reopen the image and correct visible text labels, icon semantics, and any weak/low-confidence regions in the JSON before writing the spec. Do not replace extracted geometry with a prompt-template layout from memory.

The visual contract is the program-side proof that HTML is based on the actual image. It must include:
- `reference_path`: the selected `image2/image2_reference_XX.png`
- `observed_from_reference: true`
- `reconstruction_source: "image2_reference_only"`
- `prompt_context_discarded: true`
- `observed_as_fresh_image: true`
- `derivation_method: "fresh_visual_observation_record"`
- `observation_record_path`: the selected `spec/slideXX_image_observation.json`
- `observation_record_sha256`: sha of that observation record
- `observation_id`: matching the observation record
- `observation_evidence`: a concrete paragraph describing what was seen in the image, written without referencing prompts, prompt templates, analysis reports, or remembered intent
- `regions`: one object per visible major region, each with `id`, `role`, numeric `bbox: [x,y,w,h]`, `type`, `border_style`, `fill`, `text`, and `icon_semantics`
- Required region roles/ids: nav, title, content, takeaway, source

Use `border_style: "none"` when the image has no border. Use `dashed` only if the selected Image2 reference visibly has a dashed border in that region. Do not infer borders from a reusable component pattern.
Do not include `PROMPT_TEMPLATE`, `LAYOUT_NAME`, `final_prompt`, `analysis_report`, `image2_prompt`, prompt-library names, or prompt-derived rationale in the visual contract.

Then write the human-readable spec:

```text
output/<task-name>/spec/slide01_spec.md
```

Each spec must include canvas contract, element inventory, layout grid, component specs, icon plan, conversion risks, and HTML checklist. The spec must summarize the visual contract rather than replacing it.
Each spec must explicitly list:
- The reconstruction boundary: `reconstruction_source: image2_reference_only` and `prompt_context_discarded: true`
- The observation binding: `slideXX_image_observation.json` and the matching `observation_id`
- The visible title text and line count
- Main module geometry and density
- Every icon and how it will be rebuilt or preserved
- Which text must be editable
- Any local image crop allowed and why
- Chinese font choice and expected PPTX conversion risk
The spec must describe what is visible in the image. It must not cite prompt templates, final prompts, prompt-selection audits, analysis reports, or image2 prompt files as layout/design evidence.

If the spec says an icon will be preserved as an asset, save the asset under `html/assets/` and use `data-pptx-image` in HTML. If the icon will be editable, the HTML must contain actual visible geometry or a recognizable symbol, not an uppercase abbreviation.
If preserving an icon as a screenshot crop, create the asset with `clean-icon-crop`; raw screenshot crops with visible background blocks or clipped strokes are invalid.

8. Author the declared reconstruction source from the observed image/spec:

```text
output/<task-name>/html/slide01.html
```

Before writing HTML or a direct-PPTX render plan, each slide must have both `slideXX_visual_measurement.json` and `slideXX_visual_contract.json`. Treat the measurement file as the "first time seeing this image" record: it must come from the selected Image2 reference only, after the post-image memory boundary, and it must include canvas size, region bboxes, color samples, font-size estimates, and text transcription status. Do not use upstream prompt intent, source-material analysis, or slide-story memory to invent layout.

Every major visual container in the declared reconstruction source must bind to the measured visual contract. HTML uses `data-ref-id="<region id>"`; direct-PPTX render plans must record the same region id in their source metadata/comments or sidecar plan. Do not create large cards, panels, rails, chart wrappers, tables, takeaways, or decorative borders that are not declared in `slideXX_visual_measurement.json` and `slideXX_visual_contract.json`. If a source contains `dashed` but the visual contract has no dashed border, the render preflight fails.

Choose the commercial render path deliberately:

- Use `render_path: "html"` when the HTML renderer can preserve the measured layout and final PPTX preview.
- Use `render_path: "direct_pptx"` when direct shape/text placement is more stable for matching the reference. HTML may still be created for debugging, but it is not a required delivery source.

Record the chosen path before final PPTX QA:

```bash
python3 <plugin-root>/scripts/paopao_run.py record-commercial-render \
  --task-dir output/<task-name> \
  --render-path html \
  --pptx output/<task-name>/pptx/<deck-name>.pptx
```

Use `--render-path direct_pptx` for a direct editable-PPTX build. The contract requires `source_of_truth: image2_reference`, `post_image_inputs_only: true`, final PPTX hash binding, `actual_preview_dir: "qa/pptx_actual"`, and `commercial_similarity_min` at least `0.95`.

If using the HTML path, write one HTML file per slide:

```text
output/<task-name>/html/slide01.html
```

Use the renderer guide at:

```text
<plugin-root>/reference/renderer_guide.md
```

Run the HTML-stage check only when the declared path is HTML. If the HTML slide count does not equal the requested page count, or any spec/HTML lint fails, fix it before rendering.

```bash
python3 <plugin-root>/scripts/paopao_run.py check --task-dir output/<task-name> --stage html
```

This gate also renders internal HTML preview PNGs under `qa/html_reference/` and compares them against the selected `image2_reference_XX.png` files. Low similarity means the HTML is a rough interpretation of the reference image; refill the HTML/assets from the image-derived visual contract before rendering PPTX.

The check must also pass prompt/spec/HTML lint: no whole-slide image, no SVG root artwork, no unmarked `<img>`, no CSS background images, no unsupported decorative effects, and no missing spec or final prompt.
The check also requires one `slideXX_visual_measurement.json` and one `slideXX_visual_contract.json` per slide, with matching region ids and matching `data-ref-id` bindings in HTML.
The HTML check also rejects text-only icon placeholders inside icon containers. If this fails, add real icon geometry or a marked local icon asset before rendering.

9. Render PPTX through the declared commercial path:

```bash
python3 <plugin-root>/scripts/paopao_run.py render \
  --task-dir output/<task-name> \
  --pptx output/<task-name>/pptx/<deck-name>.pptx
```

For `direct_pptx`, use the image-derived measurement and visual contract to place editable PowerPoint shapes/text directly. Do not write direct-PPTX coordinates from prompt-library memory, analysis notes, or a generic deck template. After creating the PPTX, run `record-commercial-render --render-path direct_pptx` so the pipeline knows HTML/render_manifest are not required.

10. QA the actual PPTX output by opening the PPTX in PowerPoint and visually inspecting it. Do not ask the user to install LibreOffice for normal QA. Do not use HTML PDF, renderer PDF, or PNG screenshots as the final proof.
Before PowerPoint inspection, run:

```bash
python3 <plugin-root>/scripts/paopao_run.py check \
  --task-dir output/<task-name> \
  --stage pptx \
  --pptx output/<task-name>/pptx/<deck-name>.pptx
```

QA must verify:
- PPTX opens normally and has exactly the requested page count.
- Main title, subtitles, labels, bullets, source, and takeaway text are editable.
- Chinese text renders normally with no garbling, tofu boxes, missing glyphs, or unexpected font substitution.
- No whole-slide image background exists.
- Main modules are aligned, compact, readable, and visually similar to Image2.
- Icons are visible and semantically recognizable.
- No text overlaps, table/takeaway collision, white-block masking, or broken nav centering.

If any item fails, fix and rerender. Do not deliver a known-bad PPTX.
If a PowerPoint issue is caused by the declared source-to-PPTX conversion, refill that source and rerender. Do not use final PPTX micro-adjustments as the primary fix unless the declared path is direct_pptx and the adjustment is generated from the same image-derived measurement.

Record the completed PowerPoint inspection in:

```text
output/<task-name>/qa/powerpoint_review.json
```

The JSON must state `actual_pptx_opened: true`, include the `.pptx` path, and contain one slide object per requested page with `status: "pass"`, `compared: true`, `actual_powerpoint_opened: true`, and `dimensions_checked` covering actual PowerPoint opening, text visibility, number visibility, layout match, and no overlap.

11. Run the reference-fidelity loop. For each slide, reopen the selected Image2 reference and the final PPTX/actual preview for that same slide. Compare:
- title line breaks and visual weight
- nav/tab position and active state
- module count, geometry, density, and whitespace
- color hierarchy and border weights
- icon presence, size, and semantic recognizability
- chart/table/card proportions
- bottom takeaway bar placement and source line

If the PPTX looks like a rough interpretation rather than a close rebuild of the Image2 reference, fix the measurement and declared reconstruction source, then rerender. Record the completed comparison in:

```text
output/<task-name>/qa/fidelity_review.json
```

The JSON should contain one object per slide with `status: "pass"`, `compared: true`, `reference_path`, `actual_preview_path`, a concrete `evidence` sentence, and a `dimensions_checked` list. `reference_path` must point to the selected `image2_reference_XX.png`; `actual_preview_path` must point to a real final-PPTX preview image for the same slide under `qa/pptx_actual/`. Never use `qa/html_reference/`, browser HTML screenshots, or copied HTML preview images as final-PPTX evidence. Do not use Markdown for this record.
Each slide's `dimensions_checked` must explicitly mention nav, title, module_geometry, icons, takeaway, and color_hierarchy, plus any slide-specific chart/table/rail checks. A generic "looks good" entry is invalid.
The commercial similarity gate recomputes Image2-vs-real-PowerPoint-preview similarity and fails below `0.95`. This score is a delivery threshold, not a design navigator; the source must already be measured from the image before this gate runs.

Before delivery, run the single mandatory finalization gate:

```bash
python3 <plugin-root>/scripts/paopao_run.py finalize-delivery \
  --task-dir output/<task-name> \
  --pptx output/<task-name>/pptx/<deck-name>.pptx
```

This command runs the full pipeline gate, writes `qa/pipeline_pass.json`, removes prompt Markdown artifacts, publishes the user-facing delivery folder, verifies delivery, and writes `qa/final_delivery_pass.json`. If this gate fails, do not deliver or reply with file links. Fix the failed stage and rerun it.

12. Finalize delivery. By default this deletes prompt artifacts from the task output; use `--keep-private-prompts` only for explicit local debugging:

```bash
python3 <plugin-root>/scripts/paopao_run.py finalize-delivery \
  --task-dir output/<task-name> \
  --pptx output/<task-name>/pptx/<deck-name>.pptx
```

If finalization finds any prompt Markdown files, internal JSON/Markdown, temporary Office files such as `~$*.pptx`, missing PowerPoint review, missing fidelity review, missing delivery images, required delivery HTML for the HTML path, or missing pass receipts, fix before replying to the user.
Final response must link only to user-facing files under `output/<task-name>/delivery/`, not to files under the build folders. The `delivery/` folder is the only user-facing output folder. It may contain exactly the final PPTX, `images/` slide images, and optional `html/` rebuilt HTML/assets. `analysis/`, `spec/`, build `html/`, build `image2/`, and `qa/` are internal build folders.

## Paopao Visual Language

All new Image2 references and PPTX slides should preserve the chosen prompt-library layout while applying the visual language shown in the user's preferred references:

- One strong black action title at the top, usually 1-2 lines, with a quantified anchor when possible.
- Thin top navigation should look like a consulting directory bar: full-width deep-blue strip around 36-42px high, compact text labels, subtle separators, and page number on the right. It should not look like large equal-width web tabs or filled buttons.
- White page background and mostly white content surfaces; deep-blue accents; pale-blue only as local highlight or necessary grouping tint; neutral grey secondary text/rules; black title and primary body text.
- Thin, precise blue/grey rules and borders; no thick random frames, dashed guide boxes, ornamental dividers, or unnecessary nested line boxes.
- Compact cards/tables/charts/callouts with rounded corners no larger than 8px, white surfaces by default, light-blue tint only where it improves grouping or focus, and strong hierarchy.
- Bottom takeaway should be a slim text strip, usually 36-48px high, not a large illustrated banner.
- Icons should feel natural and intentional. Use them only for clear semantic anchors, not as decoration for every card or bullet. When used, icons should be small, functional, monochrome blue/grey, and easy to recreate as editable shapes.
- Avoid decorative gradients, large empty hero areas, photorealistic imagery, 3D, bokeh, complex background textures, heavy shadows, glass effects, and oversized sparse cards.
- Preserve layout diversity from the prompt library; do not force a repeated SCR/matrix/sidebar/right-rail composition unless the selected annotation itself requires it.
- Keep text readable: no paragraph walls, no tiny labels, no more than 3 bullet lines per small card.

## Folder Contract

Each task should use:

```text
output/<task-name>/
  analysis/
  image2/
  spec/
  html/
    assets/
  pptx/
  qa/
    pptx_actual/
```

## HTML Rules

These rules apply when `qa/commercial_render_contract.json` declares `render_path: "html"` or when HTML is created as optional debug output for a direct-PPTX build.

- Canvas: 1920x1080.
- Root layout: nav / title / content / takeaway / source.
- Nav: persistent visible top strip on every slide, 36-42px high on a 1920x1080 canvas. Same labels and order throughout the deck, active item only changes by page. Never set nav height to 0, omit nav text, or mark nav as not visible. Do not render nav as large equal-width filled tabs.
- Use the nine-color palette in `reference/renderer_guide.md`.
- Use Arial for English. For Chinese decks, use `Microsoft YaHei`, `PingFang SC`, or Arial fallback; do not use rare fonts.
- Prefer flex/grid layouts with stable heights and widths.
- Use `data-chart` for editable charts.
- Use `data-pptx-image` only for small icons or complex local assets, never whole slides.
- Explicitly center nav text with `text-align:center; display:flex; align-items:center; justify-content:center;`.

## Delivery

Final response should include only the published user-facing delivery paths requested by the user, normally `delivery/*.pptx`, `delivery/images/`, and `delivery/html/` only when HTML was part of the declared path or published as optional debug output, plus a short verification note and any known residual limitation. Do not reply with delivery links until `qa/final_delivery_pass.json` exists and `check --stage delivery` passes. Do not include internal Markdown, prompts, specs, QA JSON, renderer logs, or hidden workflow details unless explicitly requested. Never include or leave generated prompt Markdown artifacts in the delivered task folder.

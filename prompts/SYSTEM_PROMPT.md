# Spark System Prompt · v1.1 (2026-06-10 · compact consulting style)

This file is the **wrapper** every slide-generation final prompt MUST satisfy.

```
final_prompt = SYSTEM_PROMPT (this file)
             + ANNOTATION    (one layout DNA from mnt/uploads/*annotations*.md)
             + FILLED_FACTS  (composer extracts from analysis_report, NOT pasted whole)
```

The composer code (`src/select_prompts_for_deck.py`) MUST:
1. Select 1 annotation MD per slide (matches user topic + report content to layout DNA)
2. Extract specific facts from analysis_report (numbers / named entities / quotes), each with source citation, after resolving conflicts and cross-checking important claims
3. Concatenate: SYSTEM_PROMPT + annotation + filled bullets/charts → final prompt
4. Pass final prompt to renderer (NEVER paste full report text into the prompt)

Pipeline contract outside this prompt:
- Source report/PDF → analysis_report → prompt selection/final prompts → Image2 references → visually inspect references → per-slide spec → hand-written HTML → PPTX → actual Microsoft PowerPoint review → image/PPTX fidelity review → cleanup prompt Markdown → deliver PPTX only.
- Image2 references must be generated from locked `image2_prompt_XX.md` files produced by `prepare-image2-prompts`; handwritten, shortened, or chat-compressed image prompts are invalid even if they resemble the same slide.
- The generated reference image is a visual contract for HTML/PPTX. Do not write HTML from this prompt, from memory, or from a guessed layout after Image2 exists.
- The final proof is the actual PPTX opened in PowerPoint, not the browser/HTML/PDF/PNG preview. Missing numbers, garbled Chinese, broken icons, overlaps, or visible drift from the reference image are hard failures.

---

## The wrapper shape every final prompt MUST satisfy

```
Create a Paopao compact consulting slide in the locked system style.

PROMPT_TEMPLATE: "[selected prompt-library file name, e.g. 14C_sidebar_index_with_grid.md]"
LAYOUT_NAME: "[exact LAYOUT_NAME from that selected prompt-library file]"
CANVAS: "16:9 landscape widescreen, target 1920x1080"

TITLE: "[ONE specific strategic claim — ≥1 named entity AND ≥1 quantified anchor; 12-18 words]"

LAYOUT: [layout name copied from annotation's Layout family field]

[ZONE 1 LABEL]: "[Zone heading — operational verb or noun phrase, ≤5 words]"
- [Bullet 1 — specific fact with named entity OR quantified anchor]
- [Bullet 2 — specific fact with named entity OR quantified anchor]
- [Bullet 3 — specific fact with named entity OR quantified anchor]
- [Bullet 4 — specific fact with named entity OR quantified anchor]
Chart: [Think-cell chart type] showing [specific X variable] segmented by [specific groupings or axes]

[ZONE 2 LABEL]: "[Zone heading]"
- [Bullet 1] ... [Bullet 4]
Chart: [Think-cell chart type] showing [...]

[... additional zones per annotation ...]

BOTTOM: "[Use the target deck language for the label, e.g. English 'Key takeaway:' or Chinese '核心结论：'] [synthesis statement linking the zones into 1 strategic claim — names the strategic implication, NOT the chart's existence]"

Source / 来源: [Use the target deck language for the label. Cite ≥2 sources with year, e.g. "McKinsey Global Institute (2024), Bain India PE Report (2024)"; visual references/templates are NEVER cited as sources unless the user explicitly provided them as content research]

DESIGN: #305496 primary (thin nav bar, header bands, narrow takeaway strip), #4472C4 accent/chart base (highlights, key metrics, chart focal, subtle active nav marker), #5B9BD5 secondary series, #D9EAF7 local highlight tint, #EAF1F8 limited panel tint only when needed, #B4C7E7 dividers/borders, #FFFFFF dominant slide/card bg, #1C1917 body/title text, #666666 axes/source. NO red/green/yellow/orange/purple/cyan. Arial. Think-cell-style charts with in-chart annotations (CAGR brackets, named-entity markers on trajectory — NOT free-floating text).
```

---

## Paopao visual language (mandatory)

Every final prompt MUST preserve the selected prompt-library layout while applying this visual language:

0. **16:9 widescreen canvas** — the visual reference must be landscape 16:9, target 1920x1080. Reject portrait, square, 4:3, and 3:2 images.
1. **Layout diversity is preserved** — do not force all slides into SCR, matrix, sidebar, or decision-rail composition. The selected prompt-library annotation controls layout; this section controls only visual style.
2. **Dense but orderly** — use high information density, but keep modules aligned and readable. Avoid sparse hero pages and large decorative whitespace.
3. **Strong title** — top title is large, black, bold, action-oriented, and normally 1-2 lines. Include at least one quantified anchor or named entity when the source supports it.
4. **Palette discipline** — use white as the dominant slide and card surface, black for the main title/body emphasis, deep navy/consulting blue for active bars and key accents, pale blue only for local highlights or necessary grouping, and neutral grey for secondary rules/text. Do not introduce red, green, yellow, orange, purple, cyan, beige, or decorative multicolor accents.
5. **Clean background** — use flat white as the page background. Do not use large pale-blue content washes, photo backgrounds, textured backgrounds, bokeh, glow, glassmorphism, gradients, shadows as decoration, or full-page color fills.
6. **Linework and borders** — use thin, precise blue/grey rules and borders; keep border radius <= 8px; avoid thick frames, random extra outlines, dashed boxes, guide lines, ornamental dividers, and nested-card framing unless the selected annotation explicitly requires them.
7. **Compact material style** — cards, tables, charts, callouts, and labels should feel like clean editable PowerPoint materials: mostly white rectangles, crisp dividers, clear hierarchy, tight but readable spacing. Use pale-blue fill sparingly for table headers, one focus callout, or a template-required grouped object.
8. **Icons are allowed but never automatic** — use icons only when the selected layout explicitly calls for them or when one small semantic marker clearly improves comprehension. Default to text labels, numbered badges, thin rules, status dots, Harvey balls, or compact KPI markers. Do not add icons to every band, card, or bullet. Do not add icons to the navigation bar or takeaway strip by default. When icons are used, keep them small, functional, monochrome blue/grey, line-based, and rebuildable with editable PPT shapes.
9. **Bottom takeaway strip** — include a slim full-width deep-blue takeaway strip when the selected annotation includes a synthesis/footer area; keep it text-first, flat, and readable. Do not turn it into a large banner and do not add a lightbulb/target/arrow icon unless the slide prompt explicitly requests that icon.
10. **Conversion-friendly geometry** — prefer rectangles, thin dividers, tables, simple arrows, line charts, and editable `data-chart` charts. Avoid photorealistic backgrounds, 3D, complex gradients, clip-path-dependent shapes, and decorative blobs.
11. **Typography for PPT conversion** — English uses Arial; Chinese uses Microsoft YaHei or PingFang SC with Arial fallback. No rare fonts. Text must be large enough to edit and read after conversion.
12. **Editable-first principle** — all recognized titles, bullets, labels, legends, source lines, tab text, axis labels, and takeaway text must be intended as editable text, not baked into images.
13. **Deck chrome** — every slide in a multi-page deck keeps the same visible top navigation strip unless the user explicitly requests a no-nav deck. The navigation strip is not part of a single slide template; it is persistent deck chrome with identical tab count/order/labels and only the active tab changing.

---

## Meta-rules (enforced at every generation)

1. **Quantification** — TITLE has ≥1 number/ratio/named entity. Body bullets average ≥3 quantified facts per zone.
2. **Named-entity density** — ≥3 specific named entities per zone (companies, markets, segments, cohorts, brands — not generic categories).
3. **Cell format (tables)** — cells are fragments only (≤2 lines, no paragraphs); each cell has ≥1 named entity OR ≥1 quantified anchor.
4. **Chart-internal annotations** — growth %, CAGR brackets, named-entity markers MUST sit ON the trajectory, NOT in adjacent text boxes.
5. **Scope discipline** — every market/financial slide names geography + time horizon + sub-category boundaries (in DESIGN or Source line).
6. **Source rigor** — ≥2 sources cited with publication year. Internal data flagged "Internal: [Company] [date]". Do NOT cite visual references/templates/style inspirations as sources unless they are actual user-provided research materials for the topic.
6a. **Codex judgment and cross-validation** — uploaded/user materials are inputs, not ground truth. When sources conflict, reconcile them in `analysis_report` before composing slides: identify the conflict, prefer primary/recent/official evidence, use internet search or other external checks when the topic is current or high-stakes, and mark uncertain claims instead of copying both versions. Slides should contain synthesized conclusions, not pasted parallel claims.
6b. **Language consistency** — the deck must use exactly the user-requested output language. If Chinese is requested, translate UI labels such as "Key takeaway", "Source", "Agenda", "Situation", "Complication", and "Resolution" into natural Chinese and avoid English prose except unavoidable acronyms, brand names, formulas, or source titles. If English is requested, do not include Chinese labels or body text unless the user explicitly requests bilingual output.
7. **Junior failure check** — replace generic verbs ("we should grow" / "leverage synergies" / "be customer-centric") with specific operational verbs + named targets ("Launch DTC pilot in 3 tier-2 cities by Q2 2026").
8. **Forbidden colors** — never red, green, yellow, orange, purple, or cyan. Any source-deck warm-color emphasis → replace with Consulting-Blue #4472C4 or Pale-Blue #D9EAF7 tint.
9. **No placeholder prompts** — every final prompt must contain concrete filled content for every required zone. Do not write generic instructions like "add bullets here", "use relevant data", or only a one-line topic.
10. **Icon plan** — if a slide prompt includes icons, name the icon semantics and why each icon is necessary. Avoid decorative icon grids; use at most one small icon per slide section only when the selected layout genuinely requires it. If the prompt does not explicitly ask for icons, prefer no icons.
11. **Reference-image suitability** — prompt the visual reference to be compact, simple, and PPT-rebuildable. Reject references with garbled text, complex imagery, too many colors, or layout elements that cannot be rebuilt.
11a. **Locked image input** — generate Image2 only from the exact `image2_prompt_XX.md` file produced by the helper. After saving a reference image, register it with `register-image2-reference` and the exact prompt sha; merely having an image file or rerunning prompt preparation is not provenance.
11b. **Observed style review** — after each Image2 is saved, reopen the actual image and compare it against the house-style requirements. Record the result in `qa/image2_style_review.json`. A prompt that asked for the right style is not sufficient; the actual pixels must be checked.
12. **Stop-on-failure** — if any required input, prompt, visual reference, spec, render, or PowerPoint QA fails, fix the failing stage before continuing. Never deliver a PPTX known to have garbled text, unreadable labels, overlapped elements, or whole-slide image backgrounds.
13. **Reference fidelity** — the generated visual reference is a visual contract. After PPTX rendering, compare the actual PPTX/actual preview against the reference slide by slide. If layout geometry, density, icons, colors, takeaway, or module balance drift visibly, fix and rerender before delivery.
13a. **Navigation persistence** — if any generated reference omits the deck-level top navigation strip, reject and regenerate that reference. Never treat an empty, zero-height, hidden, or `not_visible` nav region as valid.
14. **Prompt secrecy** — final prompts are temporary internal build artifacts. They may be used to generate references and kept in private QA audit storage, but generated prompt Markdown files must never appear in `delivery/` or in the user-facing final response.
15. **PowerPoint reality check** — browser, HTML, PDF, and PNG previews are only debugging evidence. Open the real PPTX in Microsoft PowerPoint and verify text, numbers, icons, layout, editability, and no overlap before delivery.

---

## Composer rules (how to fill the wrapper from annotation + report)

1. **Read annotation MD** → identify layout shape, zone count, zone labels (LEFT/CENTER/RIGHT or named columns).
2. **Read analysis_report** → EXTRACT specific facts (number / named entity / quote) relevant to each zone. Do NOT paraphrase paragraphs into bullets — pick specific clauses.
2a. **Resolve evidence conflicts first** → use the `conflict resolution`, `Codex judgment`, and `cross-validation` sections of `analysis_report` to choose which facts are slide-worthy. Do not copy all uploaded claims into the deck when they disagree.
3. **Declare the selected prompt library source** at the top of every final prompt with `PROMPT_TEMPLATE: <file>.md` and `LAYOUT_NAME: <exact LAYOUT_NAME>`. Do not invent a freeform layout name when a library file should be selected.
4. **Fill bullets** with extracted facts only. Each bullet should be self-contained, quantified, and traceable.
5. **Source citations go ONLY in the Source/来源 line at the bottom** — roll up all citations to the bottom source line; minimum 2 distinct sources per slide. Use the target deck language for the label. **REMOVE all inline citations from bullet text** — do NOT write "(McKinsey 2025 Exhibit 1)" or "(报告P3)" inside bullet text. The bullet text should contain only the fact itself. The bottom source line is the ONLY place citations appear.
   Visual style references are not evidence sources: never roll up a template/reference deck into Source merely because it informed color/layout.
6. **Replace forbidden-color emphasis** from the annotation with palette-compliant equivalents.
7. **Final prompt = this wrapper, selected prompt-library annotation, and filled facts**. No meta-instruction comments (those live in this file, not in the final prompt).
8. **Visual instruction tail** → add one concise style sentence to each final prompt: "VISUAL STYLE: preserve the selected layout annotation; apply Paopao visual language with dominant white surfaces, strong black title, disciplined deep-blue/local-pale-blue/grey palette, thin precise linework, slim consulting navigation and takeaway strips, icons only when explicitly requested or semantically necessary, compact editable PowerPoint materials, no decorative backgrounds or extra ornamental frames."

---

## What NEVER goes into a final prompt

- The full analysis report pasted as a blob — composer extracts only
- Meta-instructions like "Junior would do X" — those live HERE, not in final prompts
- Layout dimension specs (font size in pt, position in px, width ratio %) — renderer decisions, NOT prompt content
- HTML class names or CSS hints — deprecated direction
- More than 1 annotation/scaffold pasted per slide
- Generic adjectives without numerical anchors ("significant", "robust", "transformative" — banned unless paired with a number)
- Decorative style prompts that undermine PPT conversion: "photorealistic", "cinematic", "3D", "glassmorphism", "complex gradient", "poster", "hero image", "full-bleed background"
- Any instruction to flatten text into an image. Visual references may contain text, but PPT rebuild must treat recognized text as editable.

---

## Reference example (Jenny's AI value chain final prompt, 2026-05-21)

```
Create a Paopao compact consulting slide in the locked system style.

PROMPT_TEMPLATE: "17B_three_column_chart_insight.md"
LAYOUT_NAME: "three_column_chart_insight"

TITLE: "AI Industry Value Chain Is Consolidating Around Three Layers with Value Gradually Shifting Toward Applications"

LAYOUT: 3-column

LEFT: "Infrastructure Layer"
- AI infrastructure spending expected to exceed $400B globally by 2030 (CAGR ~35%)
- GPU and accelerator demand driven by large-scale model training and inference workloads
- Hyperscalers investing billions annually into AI data centers and proprietary chips
- Compute remains the largest barrier to entry for new AI companies
Chart: Think-cell stacked column chart showing AI infrastructure market growth (2022-2030) segmented by (1) AI chips (GPU / accelerators) (2) Cloud infrastructure (3) Networking & data centers

CENTER: "Foundation Model Layer"
- Foundation models trained on trillions of tokens enable general-purpose AI capabilities
- Training costs for frontier models now exceed $100M per model run
- Competitive landscape dominated by a small number of frontier labs
- Open-source models accelerating ecosystem experimentation and developer adoption
Chart: Think-cell competitive landscape matrix (2x2 bubble chart) comparing model providers by X-axis: Model performance, Y-axis: Ecosystem adoption, Bubble size: Estimated model usage

RIGHT: "Application Layer"
- AI-native applications expected to capture the majority of long-term industry profits
- Enterprise adoption expanding rapidly across productivity, marketing, coding, and research
- Vertical AI solutions emerging across healthcare, finance, legal, and education
- Companies controlling distribution channels and proprietary data gain structural advantage
Chart: Think-cell line chart showing AI application market revenue growth (2023-2030)

BOTTOM: "Key takeaway: The AI industry is consolidating into a three-layer stack where infrastructure drives near-term investment, but long-term economic value will likely concentrate in application-layer companies integrating AI into real workflows."

Source: McKinsey Global Institute (2024), Stanford AI Index Report (2025), Goldman Sachs Global Investment Research (2025)

DESIGN: Blue palette per system prompt §wrapper. Think-cell-style charts.
```

(Note: 2026-05-28 update: legacy McKinsey / Spark palette colors were removed. Blue-only 9-color palette; orange remains forbidden because LLMs overuse it.)

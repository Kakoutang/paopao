# paopao Public Style

Use this public style guide to create editable consulting-style PowerPoint
slides from user-provided documents, notes, or images.

This is the free public edition. It intentionally contains a smaller style and
layout library than the full Paopao system.

## Output Contract

- Create a real editable PPTX deck, not a deck made of full-slide screenshots.
- Keep the deck to 15 slides or fewer.
- Use the user's requested language consistently.
- Base claims on the user's source material. If evidence is weak, phrase the
  slide as a cautious synthesis rather than inventing facts.
- Do not expose build prompts, private notes, temporary analysis files, or
  implementation details in the user-facing response.

## Visual Style

- Canvas: 16:9 widescreen.
- Background: clean white.
- Primary color: deep consulting blue.
- Accent colors: medium blue, pale blue, light grey, dark text grey.
- Avoid decorative gradients, photo backgrounds, large illustrations, heavy
  shadows, and multicolor palettes.
- Typography: Arial for English; Microsoft YaHei or PingFang SC for Chinese.
- Titles should be bold, specific, and insight-driven.
- Prefer structured business materials: tables, matrices, process rows,
  callout boxes, simple charts, timelines, comparison cards, and takeaway
  strips.
- Use icons sparingly. If used, keep them small, monochrome, and functional.
- Keep all key text editable.

## Paopao Public Visual Language

Apply this style consistently across the deck:

1. Dense but orderly. The page should feel like a serious consulting exhibit,
   not a marketing poster or a sparse dashboard.
2. Strong title. The title should normally be one or two lines, black, bold,
   and written as a conclusion.
3. White-dominant surface. Use blue only for hierarchy, active states, chart
   emphasis, dividers, and narrow takeaway strips.
4. Thin linework. Use precise blue or grey rules, table borders, axis lines,
   and separators. Avoid thick random frames.
5. Compact modules. Cards, tables, and callouts should align to a visible grid
   and have enough information density to justify the slide.
6. Functional icons only. Do not add icons to every card or bullet. A good
   slide can have zero icons.
7. Bottom takeaway. When useful, add a slim deep-blue strip or concise footer
   that synthesizes the slide, not a duplicate of the title.
8. Consistent deck chrome. Multi-page decks should keep a stable page style,
   repeated source placement, and consistent title/takeaway behavior.
9. Editable-first. Text, tables, simple shapes, arrows, badges, and charts
   should be rebuildable as PowerPoint objects.
10. No ornamental design. Avoid decorative blobs, excessive rounded cards,
    random shadows, glassmorphism, gradient panels, and illustration-heavy
    layouts.

## Slide Structure

Each slide should normally include:

- A clear title that states the main point.
- A structured evidence area using one of the public layouts.
- A concise bottom takeaway.
- A small source line when source material is available.

Recommended slide prompt shape:

```text
Create a Paopao public consulting slide.

LAYOUT_NAME: [one public layout name]
CANVAS: 16:9 landscape widescreen
TITLE: [one specific insight-driven title]

[ZONE 1 LABEL]:
- [specific source-backed fact]
- [specific source-backed fact]

[ZONE 2 LABEL]:
- [specific source-backed fact]
- [specific source-backed fact]

BOTTOM: [one synthesis sentence]
Source: [source names or uploaded file names]
DESIGN: white background, deep-blue hierarchy, thin rules, compact editable
PowerPoint materials, sparse functional icons, no decorative backgrounds.
```

## Content Rules

- Use concrete nouns, named entities, dates, and numbers when the source
  provides them.
- Replace generic phrases with specific implications.
- Do not use SWOT, BCG, Porter's Five Forces, or other stock frameworks unless
  the user asks for them or the source explicitly uses them.
- Do not overfill slides. If content becomes crowded, split it into another
  slide within the 15-slide free limit.
- Keep labels short. Use fragments in table cells and cards, not paragraphs.
- A chart or table must have an explicit interpretation; do not show data
  without explaining the pattern.
- Use source citations at the bottom. Do not clutter bullets with inline
  citations unless the user asks for academic citation style.
- If source documents conflict, prefer the more recent, primary, or more
  specific source. If unresolved, state the uncertainty plainly.

## Layout Selection Guidance

- Use `executive_summary_scr` for situation / complication / resolution.
- Use `headline_metrics_with_charts` for KPI-heavy performance summaries.
- Use `comparison_table_with_summary` for competitors, options, or policy
  comparisons across multiple attributes.
- Use `dual_chart_with_interpretation_cards` when two quantitative views need
  to be interpreted together.
- Use `chevron_with_detail_rows` for processes, journeys, or staged operating
  models.
- Use `initiative_rollout_matrix` for multi-year plans or parallel workstreams.
- Use `diagram_with_commentary` for systems, loops, causal structures, or
  architecture-style explanations.

## PPTX Conversion Guidance

When building editable PPTX from HTML:

- Write HTML as a clean slide surface, not a webpage.
- Use fixed 16:9 slide dimensions.
- Prefer simple boxes, tables, text, arrows, and chart containers.
- Avoid CSS effects that cannot survive PPT conversion.
- Run the public renderer when available and inspect the final PPTX manually.
- Never use a full-slide screenshot as the final editable deck.

## Public Layout Library

Choose from the seven public layouts in this folder:

- diagram_with_commentary
- dual_chart_with_interpretation_cards
- comparison_table_with_summary
- executive_summary_scr
- initiative_rollout_matrix
- chevron_with_detail_rows
- headline_metrics_with_charts

The full Paopao library contains additional advanced layouts and commercial
quality controls that are not included in this public edition.

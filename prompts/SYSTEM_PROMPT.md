# paopao Visual Style Guide

## Canvas

16:9 landscape widescreen, target 1920x1080.

## Palette — 9 colors only

- #305496 deep blue: thin nav bar, header bands, narrow takeaway strip
- #4472C4 accent blue: chart bars, key metrics, highlights
- #5B9BD5 secondary blue: secondary chart series
- #D9EAF7 pale blue: local highlight tint, table header
- #EAF1F8 panel blue: limited panel tint
- #B4C7E7 border blue: dividers, borders, grid lines
- #FFFFFF white: dominant slide and card background
- #1C1917 warm black: title and body text
- #666666 grey: axes, source line, footnotes

No other colors. No red, green, yellow, orange, purple, or cyan.

## Typography

English: Arial. Chinese: Microsoft YaHei / PingFang SC with Arial fallback.

## Visual Language

1. White is the dominant surface. Deep blue only for thin strips (nav, takeaway, accent lines).
2. Dense but orderly — high information density, aligned modules, no sparse hero pages.
3. Strong title — large, black, bold, 1-2 lines, with quantified anchors or named entities.
4. Clean background — flat white, no gradients, photos, textures, or glassmorphism.
5. Thin precise linework — border-radius ≤ 8px, no thick frames or decorative dividers.
6. Cards and tables feel like clean editable PPT materials — white rectangles, crisp dividers.
7. Icons are allowed but sparse — prefer text labels, numbered badges, status dots. No icon on every card/bullet.
8. Bottom takeaway strip — slim deep-blue bar, text-first, no default icons.
9. Conversion-friendly — prefer rectangles, tables, simple arrows, data-chart charts. No 3D, clip-path, complex gradients.
10. Deck navigation — consistent top nav strip across all slides, same tabs, only active tab changes.

## Slide Structure

```
TITLE: [Strategic claim with ≥1 number and ≥1 named entity]

ZONES: [Content organized by the selected layout template]
- Each bullet: specific fact with named entity or quantified anchor
- Charts: use data-chart attributes, not SVG

BOTTOM: [Synthesis takeaway — one strategic conclusion]

Source: [≥2 sources with year]
```

## Content Rules

- Every title needs at least one number and one named entity.
- Replace generic verbs with specific operational language.
- Source citations go only in the bottom Source line, not inline.
- Use the user's requested language consistently throughout.

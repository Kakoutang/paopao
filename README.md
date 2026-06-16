# paopao

paopao public edition turns source documents, notes, and reference materials
into editable consulting-style PowerPoint decks.

泡泡公开版可以把文档、笔记和参考资料整理成可编辑的咨询风格 PPT。

## Free Edition

- 7 public layout templates
- Same Image2-first production workflow as local Paopao
- Prompt-backed visual reference generation before PPTX reconstruction
- HTML-to-editable-PPTX renderer and renderer guide
- Full pipeline gates: analysis, prompt selection, image references, user
  approval, image-only reconstruction, PowerPoint QA, final delivery cleanup
- Up to 15 slides per deck
- Editable PPTX output; no whole-slide screenshot backgrounds
- Local files stay in the user's workspace

The full Paopao system has a larger private prompt library and broader advanced
layout coverage. The public edition keeps the production workflow but limits
template count and deck length.

完整版包含更大的私有模板库和更多高级版式。公开版保留生产流程，但限制模板数量和页数。

## Setup

Install the plugin in Codex, then ask paopao to create a deck.

Optional local helper check:

```bash
python3 scripts/paopao_run.py doctor
```

## Usage

Examples:

- `用这个 PDF 帮我做 5 页中文 PPT，重点是管理层汇报`
- `Make a 4-page English deck from this report for an investor briefing`

If slide count, language, or focus is missing, paopao will ask before starting.

## Full Version

For the full version, contact WeChat: `sugarong_`.

## Privacy

Source documents stay in your local workspace. This public plugin does not send
files to a Paopao server.

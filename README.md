# paopao

paopao helps create editable PowerPoint decks from PDFs, papers, reports, and reference materials inside your local AI workspace.

## Public Preview

This public preview is intended for early feedback.

- Up to 10 slides per deck.
- Output is editable `.pptx`.
- Works locally in your AI workspace.
- Quality is being updated frequently; if a result looks imperfect, try again later or reduce the page count for a cleaner first pass.

For the full commercial version or higher limits, contact WeChat: `sugarong_`.

## Install

Install this repository as a Codex plugin, then mention paopao when asking for a deck.

Example:

```text
Use paopao to make a 5-page English PPT from this PDF.
```

Before generation, paopao may ask you to confirm:

- Page count
- Language
- Focus, use case, or audience

## Included

- Codex plugin manifest
- Paopao PPT skill entry
- Local runtime scripts
- Public preview template index

## Privacy

- **您的原始文件（PDF、文档等）和生成的 PPT 始终保留在您的本地环境中，不会上传到任何服务器。**
- 在生成过程中，paopao 会将每页幻灯片的摘要信息（标题、关键数据点、结论等即将出现在 PPT 上的内容）发送至 paopao 设计服务器，用于版式匹配和页面设计。这些信息不会被存储或用于其他用途。
- paopao 还会联系许可证服务器进行激活验证和运行时更新，但不会传输您的文档或生成的 PPT 文件。

Your source files (PDFs, documents) and generated decks stay in your local workspace and are never uploaded. During generation, paopao sends per-slide summary data (titles, key figures, conclusions — content that will appear on the slides) to the paopao design server for layout matching and page design. This data is not stored or used for other purposes. paopao also contacts the license server for activation and runtime updates, but never transmits your documents or generated PPTX files.

## Limits

The public preview is intentionally limited. It may not match the full commercial version on page count, template variety, or reconstruction quality.

## Support

For feedback, access issues, or upgrades, contact WeChat: `sugarong_`.

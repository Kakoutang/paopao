---
name: "paopao-ppt"
description: "Use paopao to turn PDFs, reports, and reference materials into editable consulting-style PPTX decks."
---

# paopao PPT

paopao creates editable PowerPoint decks from source documents.

## Limits

- Up to 10 slides per deck.
- 5 included prompt templates.
- For more templates or pages, contact WeChat: sugarong_

## How to use

```bash
python3 <plugin-root>/scripts/paopao_run.py doctor
python3 <plugin-root>/scripts/paopao_run.py make-deck \
  --name <task-name> \
  --source <file> \
  --pages <n> \
  --language <language> \
  --focus <focus>
```

Continue a task:

```bash
python3 <plugin-root>/scripts/paopao_run.py make-deck --task-dir output/<task-name>
```

## Rules

- Final output must be `.pptx`, fully editable.
- Never use a whole-slide image as background.
- Follow `make-deck` prompts — it will tell you what to do next.
- Visual references are generated before building slides.
- Open the final PPTX in PowerPoint to verify before delivery.

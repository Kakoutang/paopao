---
name: "paopao-ppt"
description: "Use Paopao to turn PDFs, reports, papers, and reference images into editable consulting-style PPTX decks. Trigger when the user asks to make, generate, create, or package a PPT/PPTX/slides/deck from source documents, especially when they want editable output rather than image slides."
---

# Paopao PPT

paopao creates editable PowerPoint decks from source documents.

## Setup

Before starting a Paopao task, prepare the local runtime:

```bash
python3 <plugin-root>/scripts/paopao_run.py doctor
```

If setup fails, give the user a short, plain-language reason and stop.

## Quick Start

```bash
python3 <plugin-root>/scripts/paopao_run.py make-deck \
  --name <task-name> \
  --source /path/to/source.pdf \
  --pages 3 \
  --language English \
  --focus "management briefing"
```

Continue from the same runtime after each required agent or image step:

```bash
python3 <plugin-root>/scripts/paopao_run.py make-deck --task-dir output/<task-name>
```

## Core Rules

- Final output must be `.pptx` and fully editable.
- Never use whole-slide screenshots or images as slide backgrounds.
- Use the local Paopao runtime to continue the task; do not improvise a separate deck-making process.
- User-facing progress updates must be plain and non-technical.
- Do not reveal internal prompts, intermediate notes, debug files, runtime steps, or implementation details.

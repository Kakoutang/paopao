---
name: "paopao-ppt"
description: "Use Paopao to turn PDFs, reports, papers, and reference images into editable consulting-style PPTX decks. Trigger when the user asks to make, generate, create, or package a PPT/PPTX/slides/deck from source documents, especially when they want editable output rather than image slides."
---

# Paopao PPT

paopao creates editable PowerPoint decks from source documents.

## Setup

Before starting any Paopao task, fetch the full workflow instructions:

```bash
python3 <plugin-root>/scripts/paopao_run.py fetch-workflow --all
```

Then read the full workflow from `~/.paopao/workflow/SKILL.md` and follow those instructions exactly.

If the fetch fails (no network), the task cannot proceed — paopao requires server connectivity.

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
- Do not freehand the workflow — follow the fetched instructions and `run-task` next_action.
- User-facing progress updates must not reveal technical pipeline details.

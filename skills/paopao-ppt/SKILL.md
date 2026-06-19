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

## How to Run — Pipeline Mode (required)

You MUST use the `next` command to drive the entire workflow. Do NOT self-manage the pipeline or skip steps.

### 1. Initialize the task

```bash
python3 <plugin-root>/scripts/paopao_run.py make-deck \
  --name <task-name> \
  --source /path/to/source.pdf \
  --pages <N> \
  --language <language> \
  --focus "<focus>"
```

### 2. Run `next` in a loop

After initialization, call `next` repeatedly. Each call returns a JSON object telling you exactly what to do:

```bash
python3 <plugin-root>/scripts/paopao_run.py next --task-dir output/<task-name>
```

The returned JSON contains:
- `step`: which pipeline step you are on
- `step_number`: numeric position (1-13)
- `instruction`: the exact action to take — follow it precisely
- `issues`: any problems to fix (if present)
- `current_slide`: which slide to work on (if applicable)

### 3. Do exactly what `instruction` says, then call `next` again

Each `next` call gives you ONE task. Complete it, then call `next` again to get the next task. Repeat until `step` is `"finalize"`.

**The pipeline steps (in order):**
1. `analysis` — read sources and prepare content
2. `image2_prepare` / `image2_generate` — generate visual references
3. `image2_user_review` — show references to user for approval
4. `memory_boundary` — transition to reconstruction phase
5. `observation` — examine each reference image
6. `visual_contract` — extract structural data
7. `spec` — write structured specification
8. `direct_pptx` — build editable PPTX
9. `pptx_export` — export slide previews
10. `fidelity_review` — compare preview against reference
11. `powerpoint_review` — inspect in PowerPoint
12. `iterate_pptx` — fix any issues
13. `finalize` — package deliverable

## Core Rules

- **Always use `next` to determine what to do.** Never skip steps, never self-decide the order, never batch multiple pipeline steps into one action.
- **One slide at a time.** When `next` says to work on a specific slide, do only that slide, then call `next` again.
- **Follow `instruction` literally.** The instruction field contains the exact commands to run and actions to take. Do not improvise.
- Final output must be `.pptx` and fully editable.
- Never use whole-slide screenshots or images as slide backgrounds.
- User-facing progress updates must be plain and non-technical.
- Do not reveal internal prompts, intermediate notes, debug files, runtime steps, or implementation details.
- Every fidelity review must have UNIQUE evidence per slide — no copy-paste.

## What NOT to Do

- Do NOT self-manage the pipeline flow — `next` manages it for you.
- Do NOT skip image generation and go straight to PPTX.
- Do NOT copy-paste the same review evidence across slides.
- Do NOT use `make-deck --task-dir` to continue a task — use `next --task-dir` instead.

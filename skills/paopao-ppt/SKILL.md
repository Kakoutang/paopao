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
1. `analysis` — read sources, produce analysis report, slide story, final prompts
2. `image2_prepare` / `image2_generate` — generate Image2 visual references
3. `image2_user_review` — show references to user for approval
4. `memory_boundary` — forget analysis context, work only from images
5. `observation` — describe every visible element in each reference image
6. `visual_contract` — extract structural contract from each image
7. `spec` — write structured specification for each slide
8. `direct_pptx` — build python-pptx painter by looking at reference images
9. `pptx_export` — export slide previews from PowerPoint
10. `fidelity_review` — compare each preview against its reference image
11. `powerpoint_review` — inspect PPTX in PowerPoint editing interface
12. `iterate_pptx` — fix any failing similarity checks
13. `finalize` — run final gate and package deliverable

## Core Rules

- **Always use `next` to determine what to do.** Never skip steps, never self-decide the order, never batch multiple pipeline steps into one action.
- **One slide at a time.** When `next` says to work on a specific slide, do only that slide, then call `next` again.
- **Follow `instruction` literally.** The instruction field contains the exact commands to run and actions to take. Do not improvise.
- Final output must be `.pptx` and fully editable.
- Never use whole-slide screenshots or images as slide backgrounds.
- User-facing progress updates must be plain and non-technical.
- Do not reveal internal prompts, intermediate notes, debug files, runtime steps, or implementation details.
- Every fidelity review and powerpoint review must have UNIQUE evidence per slide — no copy-paste.
- Specs must contain an Element Inventory table with concrete element IDs, types, text, positions, and sizes.

## What NOT to Do

- Do NOT self-manage the pipeline flow — `next` manages it for you.
- Do NOT skip Image2 generation and go straight to PPTX.
- Do NOT write a single monolithic script that produces all artifacts at once.
- Do NOT copy-paste the same review evidence across slides.
- Do NOT write boilerplate specs — each slide spec must reflect its actual reference image.
- Do NOT use `make-deck --task-dir` to continue a task — use `next --task-dir` instead.

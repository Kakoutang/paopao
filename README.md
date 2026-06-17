# paopao

paopao is a local plugin for creating editable consulting-style
PPTX decks from PDFs, papers, reports, and reference images.

This MVP does not run a web app and does not call Paopao-owned model APIs. The
user's local AI workspace performs the reasoning workflow.

## What Is Included

- `skills/paopao-ppt/SKILL.md`: deck workflow entry point.
- `scripts/renderer.py`: HTML to editable PPTX renderer for the HTML commercial path.
- `scripts/pptx_qa.py`: mechanical PPTX validation and renderer-safety checks.
- `scripts/paopao_run.py`: public runtime controller, task initialization, commercial path validation, rendering, and packaging helper.
- `prompts/`: layout annotation index (templates served from production server).
- `reference/`: renderer guide (fetched from production server at runtime).

## Public Edition Limits

This public GitHub edition uses the same production workflow as local Paopao.
The public limits are:

- 5 included prompt templates instead of the full private prompt library.
- Up to 10 slides per deck.

Paopao is updating quickly right now. If an output looks like an older workflow
or skips visual references, pull/reinstall the latest GitHub version and try
again; quality may improve as updates land.

## Quality Gates

The plugin enforces the commercial delivery path with local checks:

- `make-deck` is the public runtime entrypoint. It creates or continues a task,
  copies source files into `source/`, auto-runs deterministic steps such as
  prompt-template planning and locked image-request preparation, and stops with
  a machine-readable `next_action` whenever model reasoning, image generation,
  user approval, or visual reconstruction is required.
- A commercial render contract must declare either `html` or `direct_pptx` as the editable reconstruction path, with Image2 as the source of truth.
- Direct PPTX output is allowed only when it is built from image-derived measurement/visual contracts and passes the real PowerPoint preview gate.
- HTML cannot use short text placeholders as icons inside icon containers when HTML is the declared path.
- Screenshot-cropped icon PNGs are checked for opaque corner backgrounds. Use `python3 scripts/paopao_run.py clean-icon-crop --image <reference> --box x,y,w,h --output <asset.png>` to expand the crop, remove the detected background, trim to the real icon, and export a transparent PNG.
- Final QA must explicitly compare title, module geometry, icons, takeaway, and color hierarchy against the generated visual reference, using real PPTX previews rather than HTML/browser previews.
- PowerPoint QA must confirm actual PPTX opening, visible text/numbers/icons, layout match, and no overlap.
- Delivery cleanup rejects exposed prompt Markdown files.
- Delivery cleanup rejects multiple top-level PPTX drafts; only the final PPTX should remain visible.

## Quick Test

```bash
python3 scripts/paopao_run.py doctor
python3 scripts/paopao_run.py make-deck \
  --name demo \
  --source /path/to/source.pdf \
  --pages 3 \
  --language English \
  --focus "management briefing"
```

Continue from the same runtime after each required agent or image step:

```bash
python3 scripts/paopao_run.py make-deck --task-dir output/demo
```

The command writes `qa/public_runtime_state.json`. Only `delivery/` is
user-facing after finalization; prompt, analysis, spec, image request, and QA
files are internal build artifacts.

## Open Preview And Licensing

The commercial plugin supports activation-code licensing through the companion
license API in `auth_server/`.

Open preview mode:

- Early users can render without activation while `PAOPAO_OPEN_PREVIEW=1`.
- This is the default during the initial feedback window.
- Set `PAOPAO_OPEN_PREVIEW=0` when you are ready to turn paid licensing back on.

Paid mode:

- Set `PAOPAO_FREE_MAX_SLIDES=10` to allow 10 slides without activation.
- Larger decks require a paid license and page quota.

Activate a paid installation:

```bash
PAOPAO_AUTH_URL=https://your-render-service.onrender.com \
python3 scripts/paopao_auth.py activate --code PAOPAO-PRO-XXXX
```

Check status:

```bash
python3 scripts/paopao_auth.py status
```

When `PAOPAO_AUTH_REQUIRED=1` is set in a commercial package, rendering reserves
page quota before generation. Successful jobs commit the quota; failed jobs
cancel the reservation so users are not charged for failed output.

Manual license fulfillment:

```bash
python3 ../../scripts/issue_paopao_license.py \
  --server-url "$PAOPAO_AUTH_URL" \
  --admin-key-id "$PAOPAO_ADMIN_KEY_ID" \
  --admin-private-key "$PAOPAO_ADMIN_PRIVATE_KEY" \
  --email user@example.com \
  --plan pro_monthly \
  --price-code early_bird_19_usd_monthly
```

For local development only, set:

```bash
PAOPAO_LOCAL_DEV=1
```

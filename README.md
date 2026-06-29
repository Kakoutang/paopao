# paopao

paopao is a local plugin for creating editable consulting-style
PPTX decks from PDFs, papers, reports, and reference images.

This MVP does not run a web app and does not call Paopao-owned model APIs. The
user's local AI workspace performs the reasoning workflow.

## What Is Included

- `skills/paopao-ppt/SKILL.md`: the local deck workflow.
- `scripts/paopao_run.py`: task initialization, workflow validation, rendering, and packaging helper.
- `scripts/pptx_qa.py`: mechanical PPTX validation and renderer-safety checks.
- `prompts/INDEX.md`: public prompt catalog index. Full templates and runtime files are fetched through the licensed workflow service.

## Quality Gates

The plugin enforces the commercial delivery path with local checks:

- The default production path is direct editable PPTX built from locked workflow packets and the authorized runtime.
- HTML rendering is a legacy path for explicit browser/HTML requests, not the default PPTX path.
- Screenshot-cropped icon PNGs are checked for opaque corner backgrounds. Use `python3 scripts/paopao_run.py clean-icon-crop --image <reference> --box x,y,w,h --output <asset.png>` to expand the crop, remove the detected background, trim to the real icon, and export a transparent PNG.
- Final QA must inspect real PPTX previews rather than relying on code success alone.
- PowerPoint QA must confirm actual PPTX opening, visible text/numbers/icons, layout match, and no overlap.
- Delivery cleanup rejects exposed prompt Markdown files.
- Delivery cleanup rejects multiple top-level PPTX drafts; only the final PPTX should remain visible.

## Quick Test

```bash
python3 scripts/paopao_run.py doctor
python3 scripts/paopao_run.py init --name demo --pages 3 --language English --pipeline-mode direct_pptx
python3 scripts/paopao_run.py next --task-dir output/demo
```

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
python3 scripts/paopao_auth.py activate --code PAOPAO-PLAN-XXXX
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

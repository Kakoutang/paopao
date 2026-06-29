# paopao

paopao is a local plugin for creating editable consulting-style
PPTX decks from PDFs, papers, reports, and reference images.

This MVP does not run a web app and does not call Paopao-owned model APIs. The
user's local AI workspace performs the reasoning workflow, while Paopao's
service gates the private runtime, prompt access, and page quota.

## What Is Included

- `skills/paopao-ppt/SKILL.md`: the local deck workflow.
- `scripts/paopao_run.py`: task initialization, workflow validation, rendering, and packaging helper.
- `scripts/pptx_qa.py`: mechanical PPTX validation and renderer-safety checks.
- `prompts/INDEX.md`: public prompt catalog index. Full templates and runtime files are fetched through the authorized workflow service.

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

## Free Preview And Licensing

The commercial plugin supports activation-code licensing through the companion
license API in `auth_server/`.

Free preview mode:

- First run automatically creates a `free_preview` token while `PAOPAO_OPEN_PREVIEW=1`.
- Free preview users do not enter an activation code.
- The free preview includes 10 pages and 5 prompts.
- Set `PAOPAO_OPEN_PREVIEW=0` only if you want to disable automatic free preview issuance.

Paid mode:

- Larger decks or larger prompt pools require a paid activation code and page quota.

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

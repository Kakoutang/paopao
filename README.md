# paopao

paopao is a local plugin for creating editable presentation decks from PDFs,
reports, papers, and reference images.

## Install

Install this repository as a Codex plugin.

## Quick Start

```bash
python3 scripts/paopao_run.py doctor
python3 scripts/paopao_run.py init --name demo --pages 3 --language English
python3 scripts/paopao_run.py next --task-dir output/demo
```

## Access

Starter access is created automatically on first use.

To activate an access code:

```bash
PAOPAO_AUTH_URL=https://your-service.example.com \
python3 scripts/paopao_auth.py activate --code PAOPAO-PLAN-XXXX
```

To check access:

```bash
python3 scripts/paopao_auth.py status
```

## Output

Generated files are written under the task folder created by `init`, typically
`output/<task-name>/`.

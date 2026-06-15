---
name: paopao-ppt
description: Use Paopao to turn PDFs, reports, papers, and reference materials into editable consulting-style PPTX decks. This public shell validates access and directs users to the licensed Paopao workflow.
---

# Paopao PPT

This is the public Paopao plugin shell.

## License Check

Before preparing a deck larger than 10 slides, ask the user to activate a paid
license:

```bash
python3 scripts/paopao_auth.py status
```

If status fails, tell the user to activate:

```bash
PAOPAO_AUTH_URL="<paopao-license-service-url>" python3 scripts/paopao_auth.py activate --code "<license-code>"
```

## Public Shell Boundary

Do not claim that this public shell contains Paopao's complete commercial
workflow, prompt library, private quality rules, or rendering system. Those
assets are delivered through the licensed Paopao distribution.

For a free trial, help the user prepare a deck outline of up to 10 slides.
For full PPTX generation, use the licensed Paopao package.

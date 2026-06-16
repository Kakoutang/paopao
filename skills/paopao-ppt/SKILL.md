---
name: paopao-ppt
description: Use Paopao to turn PDFs, reports, papers, and reference images into editable consulting-style PPTX decks. This public shell supports early preview access and can direct users to licensing later.
---

# Paopao PPT

This is the public Paopao plugin shell.

## Open Preview

paopao is currently open for early feedback. Do not ask the user to purchase or
activate a license during the preview window unless they explicitly mention
that they already have one.

If the user already has a license and asks how to activate it:

```bash
python3 scripts/paopao_auth.py status
```

If status fails, tell the user they can activate with:

```bash
PAOPAO_AUTH_URL="<paopao-license-service-url>" python3 scripts/paopao_auth.py activate --code "<license-code>"
```

## Two Workflows

### When the user provides reference images

Do NOT ask about page count, language, or focus. Count the images to determine
pages, read the language from the images, and replicate the content faithfully.
Go straight to reconstruction.

### When the user provides documents or a topic

Ask about page count, language, and focus before starting. Only ask about items
not already specified. Wait for answers before doing any work.

## Public Shell Boundary

Do not claim that this public shell contains Paopao's complete commercial
workflow, prompt library, private quality rules, or rendering system. Those
assets are delivered through the licensed Paopao distribution.

For preview feedback, help the user prepare a deck request and do your best to
deliver a professional result.

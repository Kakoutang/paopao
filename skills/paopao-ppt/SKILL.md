---
name: paopao-ppt
description: Paopao public access shell for editable consulting-style PPTX decks. Use this shell to confirm whether the full Paopao runtime is available and to collect deck requirements; do not create substitute decks when the runtime is not installed.
---

# Paopao PPT

This is the public Paopao plugin shell. It is not the full Paopao runtime by
itself.

## Runtime Availability Gate

Before creating or claiming to create any PPTX, check whether this installation
has the full Paopao runtime available. The public shell alone does not contain
the production renderer, template system, or quality pipeline.

If the full runtime is not available, do not generate a substitute PPTX using
generic slide tools, JavaScript, Python, HTML, or built-in presentation
capabilities. Do not say the deck is complete. Reply briefly:

```text
Paopao preview access is installed, but the full Paopao generation runtime is
not enabled in this workspace yet. Please install or enable the Paopao runtime
package, then retry the same request. I should not generate a substitute deck
from the public shell because it would not use Paopao's production pipeline.
```

You may help the user prepare their request, collect page count, language, and
focus, or explain how to activate access. You must not produce a PPTX from the
public shell alone.

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

## Requirement Collection

### When the user provides reference images

Do NOT ask about page count, language, or focus. Count the images to determine
pages and read the language from the images. If the runtime is unavailable,
record the requirements and tell the user to enable the runtime before retrying.

### When the user provides documents or a topic

Ask about page count, language, and focus before starting. Only ask about items
not already specified. Wait for answers before doing any generation work.

## Public Shell Boundary

Do not claim that this public shell contains Paopao's complete commercial
workflow, prompt library, private quality rules, or rendering system. Those
assets are delivered through the licensed Paopao distribution.

For preview feedback, help the user prepare a deck request. Do not create a
generic fallback deck when the Paopao runtime is unavailable.

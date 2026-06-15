# paopao

paopao helps turn PDFs, reports, papers, and reference materials into
editable consulting-style PowerPoint decks.

This public repository contains the installable plugin shell:

- plugin metadata
- public usage instructions
- license activation client
- lightweight workflow entrypoint

The full commercial workflow, template library, private rules, and update logic
are delivered through Paopao's license service after activation.

## Free Trial

You can try Paopao with decks up to 10 slides. Larger decks require a paid
license.

## Activate A License

After purchase, you will receive a license key from Paopao.

From this plugin folder:

```bash
PAOPAO_AUTH_URL="https://your-paopao-license-service.example.com" \
python3 scripts/paopao_auth.py activate --code "PAOPAO-PRO-XXXX"
```

Check license status:

```bash
python3 scripts/paopao_auth.py status
```

Remove local activation:

```bash
python3 scripts/paopao_auth.py logout
```

## Privacy

Your source documents stay in your local Codex environment. The license service
tracks license status, device activation, plan, and page quota metadata. It
does not need your source PDFs to validate your license.

## Support

For license help, billing questions, or commercial access, contact Paopao.

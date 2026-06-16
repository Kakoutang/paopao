# paopao

paopao helps turn PDFs, reports, papers, and reference materials into
editable consulting-style PowerPoint decks.

This public repository contains the installable plugin shell:

- plugin metadata
- public usage instructions
- license activation client
- lightweight workflow entrypoint

The full commercial workflow, template library, private rules, and update logic
are delivered through Paopao's managed distribution.

## Open Preview

paopao is currently open for early feedback. During this preview window, you do
not need a license to try it.

After the preview period, paid licensing may be enabled for larger decks or
commercial use.

## Activate A License

If you already received a license key from Paopao, activate it from this plugin
folder:


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

Your source documents stay in your local AI workspace. The license service
tracks license status, device activation, plan, and page quota metadata. It
does not need your source PDFs to validate your license.

## Support

For preview feedback, license help, or commercial access, contact Paopao.

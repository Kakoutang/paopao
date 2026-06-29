"""Paopao renderer bootstrap: fetch the licensed engine and run it."""

from __future__ import annotations

import json
import os
import ssl
import sys
import time
import urllib.request
from pathlib import Path


_CACHE = Path(os.getenv("PAOPAO_CONFIG_DIR", Path.home() / ".paopao")) / "cache"
_CACHED = _CACHE / "_renderer_engine.py"
_SERVER = os.getenv("PAOPAO_AUTH_URL", "").rstrip("/") or "https://paopao-license-api.onrender.com"
_MAX_AGE = 86400


def _ssl() -> ssl.SSLContext:
    try:
        import certifi

        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _token() -> str:
    try:
        cfg = Path(os.getenv("PAOPAO_CONFIG_DIR", Path.home() / ".paopao")) / "license.json"
        if cfg.exists():
            return str(json.loads(cfg.read_text(encoding="utf-8")).get("token", "") or "")
    except Exception:
        pass
    return ""


def _fetch() -> None:
    _CACHE.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(f"{_SERVER}/workflow/renderer.py")
    token = _token()
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=30, context=_ssl()) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    content = str(data.get("content", "") or "")
    if not content or len(content) < 500:
        raise RuntimeError("Server returned empty renderer")
    _CACHED.write_text(content, encoding="utf-8")


def _ensure() -> Path:
    needs = not _CACHED.exists() or _CACHED.stat().st_size < 500
    if not needs and _MAX_AGE > 0 and time.time() - _CACHED.stat().st_mtime > _MAX_AGE:
        needs = True
    if needs:
        _fetch()
    return _CACHED


if __name__ == "__main__":
    engine = _ensure()
    sys.argv[0] = str(engine)
    code = compile(engine.read_text(encoding="utf-8"), str(engine), "exec")
    exec(code)

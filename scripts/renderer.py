"""Paopao renderer bootstrap — downloads the engine from server and runs it."""
import hashlib
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


def _ssl():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _fetch():
    _CACHE.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(f"{_SERVER}/workflow/renderer.py")
    token = ""
    try:
        cfg = Path(os.getenv("PAOPAO_CONFIG_DIR", Path.home() / ".paopao")) / "license.json"
        if cfg.exists():
            token = json.loads(cfg.read_text()).get("token", "")
    except Exception:
        pass
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=30, context=_ssl()) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    content = data.get("content", "")
    if not content or len(content) < 500:
        raise RuntimeError("Server returned empty renderer")
    _CACHED.write_text(content, encoding="utf-8")


def _ensure():
    needs = not _CACHED.exists() or _CACHED.stat().st_size < 500
    if not needs and _MAX_AGE > 0:
        if time.time() - _CACHED.stat().st_mtime > _MAX_AGE:
            needs = True
    if needs:
        _fetch()
    return _CACHED


if __name__ == "__main__":
    engine = _ensure()
    sys.argv[0] = str(engine)
    code = compile(engine.read_text(encoding="utf-8"), str(engine), "exec")
    exec(code)

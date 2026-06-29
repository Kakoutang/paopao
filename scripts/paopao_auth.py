#!/usr/bin/env python3
"""Access and design-service client for paopao.

The plugin stores only a signed token and public access summary locally.
Source files remain on the user's machine. The public plugin automatically
creates a starter token on first use, then fetches only the runtime and
prompt files available to that token or access code.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import socket
import ssl
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any


PLUGIN_VERSION = "0.3.1"
WORKFLOW_PROTOCOL_VERSION = "direct-pptx-packets-v1"
CONFIG_DIR = Path(os.getenv("PAOPAO_CONFIG_DIR", Path.home() / ".paopao"))
LICENSE_PATH = CONFIG_DIR / "license.json"
DEFAULT_SERVER_URL = "https://paopao-license-api.onrender.com"
DEFAULT_TIMEOUT = 20
UPDATE_INSTRUCTION_ZH = "请先运行 paopao 的增量更新脚本：python3 scripts/paopao_update.py，然后重新开始生成 PPT；如果没有这个脚本，再重新下载安装最新版 paopao 插件。"
UPDATE_INSTRUCTION_EN = "Please run the paopao incremental updater first: python3 scripts/paopao_update.py, then restart PPT generation. If the updater is unavailable, reinstall the latest paopao plugin."


class AuthError(RuntimeError):
    pass


def ssl_context() -> ssl.SSLContext:
    try:
        import certifi

        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def device_id() -> str:
    raw = "|".join(
        [
            platform.system(),
            platform.release(),
            platform.machine(),
            socket.gethostname(),
            str(Path.home()),
        ]
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]


def read_license() -> dict[str, Any]:
    if not LICENSE_PATH.exists():
        return {}
    return json.loads(LICENSE_PATH.read_text(encoding="utf-8"))


def write_license(data: dict[str, Any]) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    LICENSE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        LICENSE_PATH.chmod(0o600)
    except OSError:
        pass


def server_url(value: str = "") -> str:
    url = value or os.getenv("PAOPAO_AUTH_URL") or read_license().get("server_url", "") or DEFAULT_SERVER_URL
    return str(url).rstrip("/")


def auth_token() -> str:
    return str(read_license().get("token", "") or "")


def request_json(method: str, url: str, payload: dict[str, Any] | None = None, token: str = "") -> dict[str, Any]:
    data = None
    headers = {
        "Content-Type": "application/json",
        "X-Paopao-Plugin-Version": PLUGIN_VERSION,
        "X-Paopao-Workflow-Protocol": WORKFLOW_PROTOCOL_VERSION,
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=DEFAULT_TIMEOUT, context=ssl_context()) as resp:
            body = resp.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        if exc.code == 426:
            message = ""
            try:
                parsed = json.loads(detail)
                message = str(parsed.get("detail", "")).strip()
            except Exception:
                message = detail.strip()
            raise AuthError(
                "paopao 插件需要更新后才能继续生成。\n"
                f"{message}\n"
                "如果你不熟悉操作，请直接把这句话发给 Codex：\n"
                f"{UPDATE_INSTRUCTION_ZH}\n"
                f"English: {UPDATE_INSTRUCTION_EN}"
            ) from exc
        message = ""
        try:
            parsed = json.loads(detail)
            message = str(parsed.get("detail", "")).strip()
        except Exception:
            message = ""
        if not message:
            if "<html" in detail.lower() or "<!doctype html" in detail.lower():
                message = "paopao 服务暂时不可用，请稍后重试。"
            else:
                message = detail.strip()
        if len(message) > 500:
            message = message[:500].rstrip() + "..."
        raise AuthError(f"paopao service rejected request: HTTP {exc.code} {message}") from exc
    except urllib.error.URLError as exc:
        raise AuthError(
            "无法连接 paopao 服务，请检查网络后重试。"
            f"English: cannot reach paopao service: {exc}"
        ) from exc


def activate(code: str, url: str) -> dict[str, Any]:
    base = server_url(url)
    if not base:
        raise AuthError("missing paopao service URL. Set PAOPAO_AUTH_URL or pass --server-url.")
    result = request_json(
        "POST",
        f"{base}/activate",
        {
            "code": code,
            "device_id": device_id(),
            "plugin_version": PLUGIN_VERSION,
        },
    )
    data = {
        "server_url": base,
        "token": result["token"],
        "device_id": device_id(),
        "activated_at": int(time.time()),
        "license": result.get("license", {}),
    }
    write_license(data)
    return data


def open_preview_enabled() -> bool:
    return os.getenv("PAOPAO_OPEN_PREVIEW", "1") != "0"


def activate_preview(url: str = "") -> dict[str, Any]:
    base = server_url(url)
    if not base:
        raise AuthError("missing paopao service URL. Set PAOPAO_AUTH_URL or pass --server-url.")
    result = request_json(
        "POST",
        f"{base}/preview/activate",
        {
            "device_id": device_id(),
            "plugin_version": PLUGIN_VERSION,
        },
    )
    data = {
        "server_url": base,
        "token": result["token"],
        "device_id": device_id(),
        "activated_at": int(time.time()),
        "access_mode": "free_preview",
        "license": result.get("license", {}),
    }
    write_license(data)
    return data


def ensure_preview_access() -> dict[str, Any]:
    data = read_license()
    if data.get("token") and data.get("server_url"):
        return data
    if not open_preview_enabled():
        raise AuthError(
            "Paopao starter access is disabled (PAOPAO_OPEN_PREVIEW=0). "
            "Enable it or activate an access code with scripts/paopao_auth.py activate --code <code>."
        )
    return activate_preview()


def status(allow_dev: bool = True) -> dict[str, Any]:
    data = read_license()
    token = data.get("token", "")
    base = server_url()
    if not token or not base:
        if allow_dev and os.getenv("PAOPAO_LOCAL_DEV") == "1":
            return {
                "licensed": True,
                "mode": "local-dev",
                "message": "PAOPAO_LOCAL_DEV=1 bypass is enabled.",
            }
        if open_preview_enabled():
            data = ensure_preview_access()
            token = data.get("token", "")
            base = server_url()
        if not token or not base:
            raise AuthError("paopao access is unavailable. Please update paopao or contact support if this keeps happening.")
    result = request_json("GET", f"{base}/license/status", token=token)
    data["license"] = result.get("license", data.get("license", {}))
    data["last_status_at"] = int(time.time())
    write_license(data)
    return result


def require_license() -> dict[str, Any]:
    return status(allow_dev=True)


def reserve(job_id: str, pages: int) -> dict[str, Any]:
    data = read_license()
    token = data.get("token", "")
    base = server_url()
    if os.getenv("PAOPAO_LOCAL_DEV") == "1" and (not token or not base):
        return {"reservation_id": "local-dev", "license": {"remaining_pages": 999999}}
    if not token or not base:
        if open_preview_enabled():
            data = ensure_preview_access()
            token = data.get("token", "")
            base = server_url()
    if not token or not base:
        raise AuthError("paopao access is unavailable. Please update paopao or contact support if this keeps happening.")
    result = request_json(
        "POST",
        f"{base}/usage/reserve",
        {"job_id": job_id, "pages": pages, "plugin_version": PLUGIN_VERSION},
        token=token,
    )
    data["license"] = result.get("license", data.get("license", {}))
    write_license(data)
    return result


def fetch_prompt_catalog() -> dict[str, Any]:
    if not auth_token() and open_preview_enabled():
        ensure_preview_access()
    return request_json("GET", f"{server_url()}/prompts/catalog", token=auth_token())


def fill_prompt_template(template: str, fills: dict[str, str]) -> dict[str, Any]:
    if not auth_token() and open_preview_enabled():
        ensure_preview_access()
    return request_json(
        "POST",
        f"{server_url()}/prompts/fill",
        {"template": template, "fills": fills},
        token=auth_token(),
    )


def fetch_workflow_file(name: str) -> dict[str, Any]:
    if not auth_token() and open_preview_enabled():
        ensure_preview_access()
    return request_json("GET", f"{server_url()}/workflow/{name}", token=auth_token())


def finish_reservation(command: str, reservation_id: str) -> dict[str, Any]:
    if reservation_id == "local-dev":
        return {"ok": True, "license": {"remaining_pages": 999999}}
    data = read_license()
    token = data.get("token", "")
    base = server_url()
    if not token or not base:
        raise AuthError("paopao access is unavailable. Please update paopao or contact support if this keeps happening.")
    result = request_json(
        "POST",
        f"{base}/usage/{command}",
        {"reservation_id": reservation_id},
        token=token,
    )
    data["license"] = result.get("license", data.get("license", {}))
    write_license(data)
    return result


def logout() -> None:
    if LICENSE_PATH.exists():
        LICENSE_PATH.unlink()


def print_json(data: dict[str, Any]) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Paopao access client")
    sub = parser.add_subparsers(dest="command", required=True)

    activate_cmd = sub.add_parser("activate", help="Activate an access code for this installation")
    activate_cmd.add_argument("--code", required=True)
    activate_cmd.add_argument("--server-url", default="")

    preview_cmd = sub.add_parser("preview", help="Enable starter access for this installation")
    preview_cmd.add_argument("--server-url", default="")

    sub.add_parser("status", help="Check paopao access status")
    sub.add_parser("require", help="Fail if paopao access is unavailable")

    reserve_cmd = sub.add_parser("reserve", help="Reserve page access before a job")
    reserve_cmd.add_argument("--job-id", required=True)
    reserve_cmd.add_argument("--pages", required=True, type=int)

    for name in ["commit", "cancel"]:
        cmd = sub.add_parser(name, help=f"{name.title()} a reserved job")
        cmd.add_argument("--reservation-id", required=True)

    sub.add_parser("logout", help="Remove local activation token")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        if args.command == "activate":
            print_json(activate(args.code, args.server_url))
        elif args.command == "preview":
            print_json(activate_preview(args.server_url))
        elif args.command == "status":
            print_json(status())
        elif args.command == "require":
            print_json(require_license())
        elif args.command == "reserve":
            print_json(reserve(args.job_id, args.pages))
        elif args.command in {"commit", "cancel"}:
            print_json(finish_reservation(args.command, args.reservation_id))
        elif args.command == "logout":
            logout()
            print_json({"ok": True})
        return 0
    except AuthError as exc:
        print(f"Paopao access error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

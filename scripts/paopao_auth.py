#!/usr/bin/env python3
"""License client for Paopao for Codex.

The plugin stores only a signed token and public license summary locally.
Source documents remain on the user's machine; the license server receives
only device, plan, quota, and job-page metadata.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import socket
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any


PLUGIN_VERSION = "0.1.0"
CONFIG_DIR = Path(os.getenv("PAOPAO_CONFIG_DIR", Path.home() / ".paopao"))
LICENSE_PATH = CONFIG_DIR / "license.json"
DEFAULT_TIMEOUT = 20


class AuthError(RuntimeError):
    pass


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
    url = value or os.getenv("PAOPAO_AUTH_URL") or read_license().get("server_url", "")
    return str(url).rstrip("/")


def request_json(method: str, url: str, payload: dict[str, Any] | None = None, token: str = "") -> dict[str, Any]:
    data = None
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=DEFAULT_TIMEOUT) as resp:
            body = resp.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise AuthError(f"license server rejected request: HTTP {exc.code} {detail}") from exc
    except urllib.error.URLError as exc:
        raise AuthError(f"cannot reach license server: {exc}") from exc


def activate(code: str, url: str) -> dict[str, Any]:
    base = server_url(url)
    if not base:
        raise AuthError("missing license server URL. Set PAOPAO_AUTH_URL or pass --server-url.")
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
        raise AuthError("not activated. Run paopao_auth.py activate first.")
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
        raise AuthError("not activated. Run paopao_auth.py activate first.")
    result = request_json(
        "POST",
        f"{base}/usage/reserve",
        {"job_id": job_id, "pages": pages},
        token=token,
    )
    data["license"] = result.get("license", data.get("license", {}))
    write_license(data)
    return result


def finish_reservation(command: str, reservation_id: str) -> dict[str, Any]:
    if reservation_id == "local-dev":
        return {"ok": True, "license": {"remaining_pages": 999999}}
    data = read_license()
    token = data.get("token", "")
    base = server_url()
    if not token or not base:
        raise AuthError("not activated. Run paopao_auth.py activate first.")
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
    parser = argparse.ArgumentParser(description="Paopao license client")
    sub = parser.add_subparsers(dest="command", required=True)

    activate_cmd = sub.add_parser("activate", help="Activate this plugin installation")
    activate_cmd.add_argument("--code", required=True)
    activate_cmd.add_argument("--server-url", default="")

    sub.add_parser("status", help="Check license status")
    sub.add_parser("require", help="Fail if the plugin is not licensed")

    reserve_cmd = sub.add_parser("reserve", help="Reserve page quota before a job")
    reserve_cmd.add_argument("--job-id", required=True)
    reserve_cmd.add_argument("--pages", required=True, type=int)

    for name in ["commit", "cancel"]:
        cmd = sub.add_parser(name, help=f"{name.title()} a reserved job quota")
        cmd.add_argument("--reservation-id", required=True)

    sub.add_parser("logout", help="Remove local activation token")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        if args.command == "activate":
            print_json(activate(args.code, args.server_url))
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
        print(f"Paopao license error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

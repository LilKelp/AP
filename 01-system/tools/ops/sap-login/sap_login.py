"""
SAP GUI login helper (SAP GUI scripting).

Purpose
- Ensure a SAP GUI session is logged in for downstream scripted pipelines.
- Defaults to reading SAP connection details from `01-system/configs/apis/API-Keys.md`.

Notes
- Never prints or writes the SAP password.
- Writes machine-readable results under `03-outputs/sap-login/`.

Usage:
  python 01-system/tools/ops/sap-login/sap_login.py
  python 01-system/tools/ops/sap-login/sap_login.py --entry "ECP(1)" --client 800 --user AZHAO
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

import win32com.client

BASE_DIR = Path(__file__).resolve().parents[4]
DEFAULT_CONFIG_PATH = BASE_DIR / "01-system" / "configs" / "apis" / "API-Keys.md"
DEFAULT_OUTPUT_ROOT = BASE_DIR / "03-outputs" / "sap-login"


@dataclass(frozen=True)
class SapLoginConfig:
    entry: str
    client: str
    user: str
    password: str | None


def parse_kv_file(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    if not path.exists():
        return data
    for raw_line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if (
            (value.startswith('"') and value.endswith('"'))
            or (value.startswith("'") and value.endswith("'"))
        ):
            value = value[1:-1]
        data[key] = value
    return data


def is_placeholder(value: str | None) -> bool:
    if value is None:
        return True
    stripped = value.strip()
    return stripped in {"", "...", "sk-..."}


def resolve_saplogon_exe(explicit_path: str | None) -> Path | None:
    if explicit_path:
        candidate = Path(explicit_path)
        return candidate if candidate.exists() else None

    candidates = [
        Path(os.environ.get("ProgramFiles", "")) / "SAP" / "FrontEnd" / "SAPgui" / "saplogon.exe",
        Path(os.environ.get("ProgramFiles(x86)", "")) / "SAP" / "FrontEnd" / "SAPgui" / "saplogon.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def start_saplogon(saplogon_path: Path | None) -> None:
    if saplogon_path is None:
        return
    try:
        subprocess.Popen(
            [str(saplogon_path)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0)
            | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
        )
    except Exception:
        # Best effort only; connection may still exist.
        return


def get_scripting_engine(ensure_started: bool, saplogon_path: Path | None) -> object:
    try:
        sapgui = win32com.client.GetObject("SAPGUI")
        return sapgui.GetScriptingEngine
    except Exception:
        pass

    if ensure_started:
        start_saplogon(saplogon_path)
        time.sleep(1.5)
        try:
            sapgui = win32com.client.GetObject("SAPGUI")
            return sapgui.GetScriptingEngine
        except Exception:
            pass

    # Fallback: scripting controller can exist even when SAPGUI ROT isn't populated.
    ctrl = win32com.client.Dispatch("Sapgui.ScriptingCtrl.1")
    return ctrl.GetScriptingEngine()


def wait_until(fn, timeout_s: float, sleep_s: float = 0.25):
    deadline = time.time() + timeout_s
    last_exc: Exception | None = None
    while time.time() < deadline:
        try:
            value = fn()
            if value:
                return value
        except Exception as exc:
            last_exc = exc
        time.sleep(sleep_s)
    if last_exc:
        raise last_exc
    return None


def get_first_session(connection):
    for attr in ("Sessions", "Children"):
        try:
            col = getattr(connection, attr)
            if col is None:
                continue
            try:
                if int(col.Count) > 0:
                    if hasattr(col, "Item"):
                        return col.Item(0)
                    return col(0)
            except Exception:
                pass
            try:
                return col(0)
            except Exception:
                pass
        except Exception:
            continue
    return None


def iter_collection(collection):
    try:
        count = int(collection.Count)
    except Exception:
        count = 0

    if count:
        for idx in range(count):
            try:
                if hasattr(collection, "Item"):
                    yield collection.Item(idx)
                else:
                    yield collection(idx)
            except Exception:
                continue
        return

    try:
        for item in collection:
            yield item
    except Exception:
        return


def find_existing_logged_in_session(app, cfg: SapLoginConfig):
    try:
        connections = app.Connections
    except Exception:
        return None

    for connection in iter_collection(connections):
        session = get_first_session(connection)
        if session is None:
            continue
        try:
            info = session_info(session)
            if info.get("user") and info.get("client") == cfg.client:
                if info.get("user", "").upper() == cfg.user.upper():
                    return session
        except Exception:
            continue
    return None


def session_info(session) -> dict[str, str]:
    info = session.Info
    return {
        "system_name": str(getattr(info, "SystemName", "")).strip(),
        "client": str(getattr(info, "Client", "")).strip(),
        "user": str(getattr(info, "User", "")).strip(),
    }


def is_logged_in(session) -> bool:
    try:
        info = session_info(session)
        if info.get("user"):
            return True
    except Exception:
        pass
    try:
        # If the password field doesn't exist, we're likely past the logon screen.
        return session.findById("wnd[0]/usr/pwdRSYST-BCODE", False) is None
    except Exception:
        return True


def try_press_default_dialog_button(session) -> bool:
    try:
        if session.ActiveWindow is not None and session.ActiveWindow.Name == "wnd[1]":
            btn = session.findById("wnd[1]/tbar[0]/btn[0]", False)
            if btn is not None:
                btn.press()
                return True
    except Exception:
        return False
    return False


def perform_login(session, cfg: SapLoginConfig, timeout_s: float) -> dict[str, str]:
    # If already logged in, do nothing.
    if is_logged_in(session):
        return {"mode": "already_logged_in"}

    if not cfg.password or is_placeholder(cfg.password):
        raise ValueError("SAP password not provided (set SAP_PASSWORD or pass --password).")

    # Fill the SAP logon screen.
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = cfg.client
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = cfg.user
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = cfg.password
    try:
        lang = session.findById("wnd[0]/usr/txtRSYST-LANGU", False)
        if lang is not None:
            lang.text = "EN"
    except Exception:
        pass

    session.findById("wnd[0]").sendVKey(0)

    # Handle possible dialogs (multi-logon/info) with best-effort default continue.
    dialog_deadline = time.time() + min(20.0, timeout_s)
    while time.time() < dialog_deadline:
        if not try_press_default_dialog_button(session):
            time.sleep(0.5)

    # Wait until login completes.
    wait_until(lambda: is_logged_in(session), timeout_s=timeout_s, sleep_s=0.5)
    return {"mode": "login"}


def write_result(output_root: Path, result: dict) -> Path:
    run_id = result.get("run_id") or datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    run_dir = output_root / "runs" / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    result_path = run_dir / "result.json"
    result_path.write_text(json.dumps(result, indent=2, sort_keys=True), encoding="utf-8")
    (output_root / "latest.json").write_text(
        json.dumps(result, indent=2, sort_keys=True), encoding="utf-8"
    )
    return result_path


def build_config(args: argparse.Namespace) -> SapLoginConfig:
    cfg = parse_kv_file(Path(args.config))
    entry = (args.entry or cfg.get("SAP_LOGON_ENTRY") or "").strip()
    client = str(args.client or cfg.get("SAP_CLIENT") or "").strip()
    user = (args.user or cfg.get("SAP_USER") or "").strip()
    password = args.password if args.password is not None else cfg.get("SAP_PASSWORD")

    missing = []
    if not entry:
        missing.append("SAP_LOGON_ENTRY (or --entry)")
    if not client:
        missing.append("SAP_CLIENT (or --client)")
    if not user:
        missing.append("SAP_USER (or --user)")
    if missing:
        raise ValueError("Missing required SAP config: " + ", ".join(missing))

    return SapLoginConfig(entry=entry, client=client, user=user, password=password)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Ensure SAP GUI is logged in via scripting.")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG_PATH))
    parser.add_argument("--entry", help="SAP Logon entry name/description.")
    parser.add_argument("--client", help="SAP client, e.g. 800.")
    parser.add_argument("--user", help="SAP username.")
    parser.add_argument("--password", help="SAP password (avoid; prefer config/SSO).")
    parser.add_argument("--timeout-s", type=float, default=120.0)
    parser.add_argument(
        "--no-start-saplogon",
        action="store_true",
        help="Do not attempt to start saplogon.exe if SAPGUI ROT is unavailable.",
    )
    parser.add_argument("--saplogon-path", help="Override saplogon.exe path.")
    parser.add_argument("--output-root", default=str(DEFAULT_OUTPUT_ROOT))
    parser.add_argument("--print-json", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output_root = Path(args.output_root)
    output_root.mkdir(parents=True, exist_ok=True)

    started = datetime.now(timezone.utc)
    run_id = started.strftime("%Y%m%d_%H%M%S")

    cfg = build_config(args)
    saplogon_path = resolve_saplogon_exe(args.saplogon_path)

    result: dict = {
        "ok": False,
        "run_id": run_id,
        "started_utc": started.isoformat(),
        "sap_logon_entry": cfg.entry,
        "client": cfg.client,
        "user": cfg.user,
    }

    try:
        app = get_scripting_engine(
            ensure_started=not args.no_start_saplogon, saplogon_path=saplogon_path
        )
        session = find_existing_logged_in_session(app, cfg)
        if session is not None:
            result.update({"mode": "reused_existing_session"})
        else:
            connection = app.OpenConnection(cfg.entry, True)
            session = wait_until(
                lambda: get_first_session(connection),
                timeout_s=max(5.0, args.timeout_s),
                sleep_s=0.25,
            )
            if session is None:
                raise RuntimeError("SAP session did not appear after OpenConnection().")

            login_meta = perform_login(session, cfg, timeout_s=args.timeout_s)
            result.update(login_meta)

        info = session_info(session)
        result.update(
            {
                "system_name": info.get("system_name", ""),
                "logged_in_user": info.get("user", ""),
                "logged_in_client": info.get("client", ""),
                "ok": True,
            }
        )
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
        try:
            # Best-effort: capture SAP status bar message (no secrets) if available.
            sbar = session.findById("wnd[0]/sbar", False)  # type: ignore[name-defined]
            if sbar is not None:
                result["sap_status"] = str(getattr(sbar, "Text", "")).strip()
        except Exception:
            pass

    result_path = write_result(output_root, result)
    if args.print_json:
        print(json.dumps(result, indent=2, sort_keys=True))
    else:
        if result.get("ok"):
            print(
                f"[OK] SAP login ready: {result.get('system_name','')} "
                f"client {result.get('logged_in_client','')} user {result.get('logged_in_user','')}"
            )
        else:
            print(f"[ERROR] SAP login failed: {result.get('error','Unknown error')}")
        print(f"[INFO] Result: {result_path.relative_to(BASE_DIR)}")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())

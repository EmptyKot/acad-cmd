from __future__ import annotations

import argparse
import json
import os
import re
import time
from pathlib import Path

import pywintypes
import win32con
import win32file
import win32pipe

from acad_cmd.autocad_bridge import AutoCADBridge


START_EVENTS = {"command_will_start"}
COMPLETION_EVENTS = {"command_ended", "command_cancelled", "command_failed"}


def _resolve_pipe_name(pid: int | None, explicit_name: str | None) -> str:
    if explicit_name:
        return explicit_name
    if pid is None:
        raise RuntimeError("Cannot resolve pipe name: pid is unknown")
    return f"acad-event-bridge-{pid}"


def _load_plugin_if_requested(bridge: AutoCADBridge, dll_path: Path | None) -> None:
    if dll_path is None:
        return
    if not dll_path.exists():
        raise FileNotFoundError(f"Plugin DLL not found: {dll_path}")

    path_for_lisp = str(dll_path).replace("\\", "/")
    cmd = f'(command "_.NETLOAD" "{path_for_lisp}")\n(princ)'
    bridge.send_command(cmd)
    wr = bridge.wait_for_idle(timeout_sec=40.0, poll_interval_sec=0.2)
    if not wr.completed:
        raise RuntimeError(f"NETLOAD did not complete: needs_input={wr.needs_input}")


def _open_pipe(pipe_name: str, timeout_sec: float):
    pipe_path = rf"\\.\pipe\{pipe_name}"
    timeout_ms = max(100, int(timeout_sec * 1000))
    win32pipe.WaitNamedPipe(pipe_path, timeout_ms)
    return win32file.CreateFile(
        pipe_path,
        win32con.GENERIC_READ,
        0,
        None,
        win32con.OPEN_EXISTING,
        0,
        None,
    )


def _normalize_command_name(raw: str | None) -> str:
    if not raw:
        return ""
    token = raw.strip().split(maxsplit=1)[0]
    token = token.lstrip("._")
    token = re.sub(r"[^A-Za-z0-9_]", "", token)
    return token.upper()


def _read_messages(handle, timeout_sec: float) -> list[dict]:
    deadline = time.time() + timeout_sec
    buffer = b""
    messages: list[dict] = []
    while time.time() < deadline:
        try:
            _hr, data = win32file.ReadFile(handle, 4096, None)
        except pywintypes.error as err:
            if int(err.winerror or 0) == 109:
                break
            raise
        if not data:
            time.sleep(0.05)
            continue
        buffer += bytes(data)
        while b"\n" in buffer:
            line_raw, buffer = buffer.split(b"\n", 1)
            line_txt = line_raw.decode("utf-8", errors="replace").strip()
            if not line_txt:
                continue
            obj = json.loads(line_txt)
            if isinstance(obj, dict):
                messages.append(obj)
    return messages


def _find_command_lifecycle(messages: list[dict], command_name: str) -> tuple[dict | None, dict | None]:
    normalized_target = _normalize_command_name(command_name)
    start_msg = None
    end_msg = None
    for msg in messages:
        if str(msg.get("type") or "") != "event":
            continue
        event_name = str(msg.get("event") or "")
        payload = msg.get("payload") or {}
        payload_name = ""
        if isinstance(payload, dict):
            payload_name = _normalize_command_name(str(payload.get("name") or ""))
        if normalized_target and payload_name and payload_name != normalized_target:
            continue
        if event_name in START_EVENTS and start_msg is None:
            start_msg = msg
            continue
        if start_msg is not None and event_name in COMPLETION_EVENTS:
            end_msg = msg
            break
    return start_msg, end_msg


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke-check command lifecycle events via bridge pipe.")
    parser.add_argument("--pid", type=int, default=None, help="AutoCAD PID (optional)")
    parser.add_argument("--pipe-name", type=str, default=None, help="Explicit pipe name (optional)")
    parser.add_argument("--timeout-sec", type=float, default=12.0, help="Pipe read timeout")
    parser.add_argument("--command", type=str, default="_.REGEN", help="AutoCAD command to execute")
    parser.add_argument(
        "--netload-dll",
        type=str,
        default=None,
        help="Optional DLL path for NETLOAD before pipe connect",
    )
    args = parser.parse_args()

    bridge = AutoCADBridge()
    if not bridge.connect():
        raise RuntimeError("Failed to connect to AutoCAD")

    if args.netload_dll:
        _load_plugin_if_requested(bridge, Path(os.path.expanduser(args.netload_dll)).resolve())

    pid = args.pid
    if pid is None:
        snap = bridge.get_status_snapshot()
        pid_raw = snap.get("acad_pid")
        pid = int(pid_raw) if pid_raw is not None else None
    pipe_name = _resolve_pipe_name(pid=pid, explicit_name=args.pipe_name)

    handle = _open_pipe(pipe_name=pipe_name, timeout_sec=float(args.timeout_sec))
    try:
        command_text = str(args.command or "").rstrip() + "\n"
        bridge.send_command(command_text)
        wr = bridge.wait_for_idle(timeout_sec=20.0, poll_interval_sec=0.1)
        if not wr.completed:
            raise RuntimeError(f"Command did not complete in smoke run: needs_input={wr.needs_input}")
        messages = _read_messages(handle, timeout_sec=float(args.timeout_sec))
    finally:
        win32file.CloseHandle(handle)

    start_msg, end_msg = _find_command_lifecycle(messages, command_name=str(args.command))
    if start_msg is None or end_msg is None:
        sample = [m for m in messages if isinstance(m, dict)][:12]
        raise RuntimeError(
            "Command lifecycle not found for target command. "
            f"start_found={start_msg is not None}, end_found={end_msg is not None}, sample={sample}"
        )

    out = {
        "pipe_name": pipe_name,
        "command": args.command,
        "start_event": start_msg,
        "completion_event": end_msg,
        "messages_total": len(messages),
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

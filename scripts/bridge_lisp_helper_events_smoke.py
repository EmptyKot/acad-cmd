from __future__ import annotations

import argparse
import json
import os
import time
import uuid
from pathlib import Path

import pywintypes
import win32con
import win32file
import win32pipe

from acad_cmd.autocad_bridge import AutoCADBridge


REQUIRED_EVENTS = {
    "lisp_will_start",
    "lisp_ended",
    "unknown_command",
    "implied_selection_changed",
}


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


def _read_messages_until(handle, timeout_sec: float, required_events: set[str]) -> list[dict]:
    deadline = time.time() + timeout_sec
    buffer = b""
    messages: list[dict] = []
    seen: set[str] = set()

    while time.time() < deadline and not required_events.issubset(seen):
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
            if not isinstance(obj, dict):
                continue
            messages.append(obj)
            if str(obj.get("type") or "") == "event":
                event_name = str(obj.get("event") or "")
                if event_name:
                    seen.add(event_name)
    return messages


def _trigger_flow(bridge: AutoCADBridge) -> dict:
    # 1) Basic LISP run.
    bridge.send_command("(princ)\n")
    wr1 = bridge.wait_for_idle(timeout_sec=10.0, poll_interval_sec=0.1)
    if not wr1.completed:
        raise RuntimeError(f"LISP smoke command did not complete: needs_input={wr1.needs_input}")

    # 2) Unknown command.
    unknown = f"ZZZ_UNKNOWN_{uuid.uuid4().hex[:8].upper()}"
    bridge.send_command(unknown + "\n")
    wr2 = bridge.wait_for_idle(timeout_sec=10.0, poll_interval_sec=0.1)
    if not wr2.completed:
        raise RuntimeError(f"Unknown command smoke did not complete: needs_input={wr2.needs_input}")

    # 3) Implied selection changed via PickFirst mutation.
    expr = (
        "(progn "
        "(setq e (entmakex (list (cons 0 \"LINE\") (cons 10 '(0.0 0.0 0.0)) (cons 11 '(10.0 0.0 0.0))))) "
        "(if e (progn (sssetfirst nil (ssadd e)) (sssetfirst nil nil))) "
        "(princ))"
    )
    bridge.send_command(expr + "\n")
    wr3 = bridge.wait_for_idle(timeout_sec=15.0, poll_interval_sec=0.1)
    if not wr3.completed:
        raise RuntimeError(f"Implied selection smoke did not complete: needs_input={wr3.needs_input}")

    return {"unknown_command": unknown}


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke-check LISP/helper events via bridge pipe.")
    parser.add_argument("--pid", type=int, default=None, help="AutoCAD PID (optional)")
    parser.add_argument("--pipe-name", type=str, default=None, help="Explicit pipe name (optional)")
    parser.add_argument("--timeout-sec", type=float, default=14.0, help="Pipe read timeout")
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
        flow = _trigger_flow(bridge)
        messages = _read_messages_until(
            handle,
            timeout_sec=float(args.timeout_sec),
            required_events=REQUIRED_EVENTS,
        )
    finally:
        win32file.CloseHandle(handle)

    event_messages = [m for m in messages if str(m.get("type") or "") == "event"]
    seen = {str(m.get("event") or "") for m in event_messages if m.get("event")}
    missing = sorted(REQUIRED_EVENTS.difference(seen))
    if missing:
        raise RuntimeError(f"Missing expected events: {missing}. Seen={sorted(seen)}")

    out = {
        "pipe_name": pipe_name,
        "flow": flow,
        "events_seen": sorted(seen),
        "event_count": len(event_messages),
        "sample_events": event_messages[:12],
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

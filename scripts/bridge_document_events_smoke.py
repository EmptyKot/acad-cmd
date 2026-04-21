from __future__ import annotations

import argparse
import json
import os
import time
from pathlib import Path

import pywintypes
import win32con
import win32file
import win32pipe

from acad_cmd.autocad_bridge import AutoCADBridge


REQUIRED_EVENTS = {"document_created", "document_activated", "document_destroyed"}


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


def _trigger_document_flow(bridge: AutoCADBridge) -> dict:
    acad = bridge.acad
    docs = acad.Documents
    base_doc = acad.ActiveDocument
    base_name = str(getattr(base_doc, "Name", "") or "")

    new_doc = docs.Add()
    wr1 = bridge.wait_for_idle(timeout_sec=20.0, poll_interval_sec=0.1)
    if not wr1.completed:
        raise RuntimeError(f"After Documents.Add AutoCAD not idle: needs_input={wr1.needs_input}")

    new_name = str(getattr(new_doc, "Name", "") or "")

    # Switch back to original doc (if available) to force activation event.
    try:
        if base_doc is not None:
            base_doc.Activate()
            wr2 = bridge.wait_for_idle(timeout_sec=10.0, poll_interval_sec=0.1)
            if not wr2.completed:
                raise RuntimeError(
                    f"After base_doc.Activate AutoCAD not idle: needs_input={wr2.needs_input}"
                )
    except Exception:
        # Keep smoke resilient: activation event may still be present from Add().
        pass

    # Close the temporary unsaved drawing (discard changes).
    new_doc.Close(False)
    wr3 = bridge.wait_for_idle(timeout_sec=20.0, poll_interval_sec=0.1)
    if not wr3.completed:
        raise RuntimeError(f"After new_doc.Close AutoCAD not idle: needs_input={wr3.needs_input}")

    return {"base_doc": base_name, "new_doc": new_name}


def _read_messages_until(
    handle,
    timeout_sec: float,
    required_events: set[str],
) -> list[dict]:
    deadline = time.time() + timeout_sec
    messages: list[dict] = []
    events_seen: set[str] = set()
    buffer = b""

    while time.time() < deadline and not required_events.issubset(events_seen):
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
                ev_name = str(obj.get("event") or "")
                if ev_name:
                    events_seen.add(ev_name)

    return messages


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke-check document lifecycle events via bridge pipe.")
    parser.add_argument("--pid", type=int, default=None, help="AutoCAD PID (optional)")
    parser.add_argument("--pipe-name", type=str, default=None, help="Explicit pipe name (optional)")
    parser.add_argument("--timeout-sec", type=float, default=12.0, help="Pipe read timeout")
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
        flow = _trigger_document_flow(bridge)
        messages = _read_messages_until(
            handle,
            timeout_sec=float(args.timeout_sec),
            required_events=REQUIRED_EVENTS,
        )
    finally:
        win32file.CloseHandle(handle)

    event_messages = [m for m in messages if str(m.get("type") or "") == "event"]
    event_names = [str(m.get("event") or "") for m in event_messages if m.get("event")]
    seen = set(event_names)
    missing = sorted(REQUIRED_EVENTS.difference(seen))
    if missing:
        raise RuntimeError(f"Missing document events: {missing}. Seen: {sorted(seen)}")

    out = {
        "pipe_name": pipe_name,
        "flow": flow,
        "events_seen": sorted(seen),
        "event_count": len(event_messages),
        "sample_events": event_messages[:6],
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

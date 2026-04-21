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
    wr = bridge.wait_for_idle(timeout_sec=30.0, poll_interval_sec=0.2)
    if not wr.completed:
        raise RuntimeError(f"NETLOAD did not complete: needs_input={wr.needs_input}")


def _read_messages(pipe_name: str, timeout_sec: float, expected_messages: int) -> list[dict]:
    pipe_path = rf"\\.\pipe\{pipe_name}"
    timeout_ms = max(100, int(timeout_sec * 1000))

    # Wait for server endpoint.
    win32pipe.WaitNamedPipe(pipe_path, timeout_ms)

    handle = win32file.CreateFile(
        pipe_path,
        win32con.GENERIC_READ,
        0,
        None,
        win32con.OPEN_EXISTING,
        0,
        None,
    )
    try:
        deadline = time.time() + timeout_sec
        chunks: list[bytes] = []
        lines: list[dict] = []
        while time.time() < deadline:
            try:
                _hr, data = win32file.ReadFile(handle, 4096, None)
            except pywintypes.error as err:
                # 109 = ERROR_BROKEN_PIPE; parse what we already have.
                if int(err.winerror or 0) == 109:
                    break
                raise
            if data:
                chunks.append(bytes(data))
                raw = b"".join(chunks)
                while b"\n" in raw:
                    line_raw, raw = raw.split(b"\n", 1)
                    line_txt = line_raw.decode("utf-8", errors="replace").strip()
                    if not line_txt:
                        continue
                    obj = json.loads(line_txt)
                    if not isinstance(obj, dict):
                        raise RuntimeError(f"Unexpected message payload type: {type(obj)}")
                    lines.append(obj)
                    if len(lines) >= expected_messages:
                        return lines
                chunks = [raw] if raw else []
            else:
                time.sleep(0.05)

        if not lines:
            raise RuntimeError("No messages received from pipe")
        return lines
    finally:
        win32file.CloseHandle(handle)


def main() -> int:
    parser = argparse.ArgumentParser(description="Connect to AcadEventBridge named pipe and read hello.")
    parser.add_argument("--pid", type=int, default=None, help="AutoCAD PID (optional)")
    parser.add_argument("--pipe-name", type=str, default=None, help="Explicit pipe name (optional)")
    parser.add_argument("--timeout-sec", type=float, default=5.0, help="Pipe wait/read timeout")
    parser.add_argument(
        "--expect-heartbeat",
        action="store_true",
        help="Also require heartbeat after hello",
    )
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
    expected = 2 if args.expect_heartbeat else 1
    messages = _read_messages(pipe_name=pipe_name, timeout_sec=float(args.timeout_sec), expected_messages=expected)

    hello = messages[0]
    if str(hello.get("type") or "") != "hello":
        raise RuntimeError(f"First message is not hello: {hello}")

    out: dict = {"pipe_name": pipe_name, "hello": hello}
    if args.expect_heartbeat:
        heartbeat = None
        for msg in messages[1:]:
            if str(msg.get("type") or "") == "heartbeat":
                heartbeat = msg
                break
        if heartbeat is None:
            raise RuntimeError(f"Heartbeat not received. Messages: {messages}")
        for field in ("seq", "ts", "busy", "queue_depth"):
            if field not in heartbeat:
                raise RuntimeError(f"Heartbeat missing field '{field}': {heartbeat}")
        out["heartbeat"] = heartbeat

    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

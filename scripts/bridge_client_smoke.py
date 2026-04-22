from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Optional

from acad_cmd.autocad_bridge import AutoCADBridge
from acad_cmd.bridge_plugin_client import EventBridgeClient


def _resolve_pipe_name(pid: Optional[int], explicit_name: Optional[str]) -> str:
    if explicit_name:
        return explicit_name
    if pid is None:
        raise RuntimeError("Cannot resolve pipe name: pid is unknown")
    return EventBridgeClient.pipe_name_for_pid(int(pid))


def _load_plugin_if_requested(bridge: AutoCADBridge, dll_path: Optional[Path]) -> None:
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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Smoke check for EventBridgeClient: connect and wait for hello/heartbeat."
    )
    parser.add_argument("--pid", type=int, default=None, help="AutoCAD PID (optional)")
    parser.add_argument("--pipe-name", type=str, default=None, help="Explicit pipe name (optional)")
    parser.add_argument("--timeout-sec", type=float, default=8.0, help="Wait timeout for hello/heartbeat")
    parser.add_argument("--netload-dll", type=str, default=None, help="Optional DLL path for NETLOAD")
    args = parser.parse_args()

    bridge = AutoCADBridge()
    if not bridge.connect():
        raise RuntimeError("Failed to connect to AutoCAD")

    dll_path = Path(os.path.expanduser(args.netload_dll)).resolve() if args.netload_dll else None
    _load_plugin_if_requested(bridge, dll_path)

    pid = args.pid
    if pid is None:
        snap = bridge.get_status_snapshot()
        pid_raw = snap.get("acad_pid")
        pid = int(pid_raw) if pid_raw is not None else None

    pipe_name = _resolve_pipe_name(pid=pid, explicit_name=args.pipe_name)
    client = EventBridgeClient(pipe_name=pipe_name, connect_timeout_sec=1.0)
    client.start()
    try:
        hello_ok = client.wait_for_hello(timeout_sec=float(args.timeout_sec))
        heartbeat_ok = client.wait_for_heartbeat(timeout_sec=float(args.timeout_sec))
        out = {
            "pipe_name": pipe_name,
            "hello_ok": bool(hello_ok),
            "heartbeat_ok": bool(heartbeat_ok),
            "snapshot": client.snapshot().to_dict(),
        }
        print(json.dumps(out, ensure_ascii=False, indent=2))
        if not hello_ok or not heartbeat_ok:
            raise RuntimeError(f"Bridge check failed: hello_ok={hello_ok}, heartbeat_ok={heartbeat_ok}")
        return 0
    finally:
        client.stop()


if __name__ == "__main__":
    raise SystemExit(main())

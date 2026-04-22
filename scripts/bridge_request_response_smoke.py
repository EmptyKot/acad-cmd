from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
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


def _sha256_hex(path: Path) -> Optional[str]:
    try:
        h = hashlib.sha256()
        with path.open("rb") as fh:
            while True:
                chunk = fh.read(1024 * 1024)
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def _load_plugin_if_requested(bridge: AutoCADBridge, dll_path: Optional[Path]) -> None:
    if dll_path is None:
        return
    if not dll_path.exists():
        raise FileNotFoundError(f"Plugin DLL not found: {dll_path}")

    source_path = str(dll_path.resolve())
    load_path = source_path
    try:
        source_path.encode("ascii")
    except Exception:
        cache_dir = Path(r"C:\Temp\acad_event_bridge_cache")
        cache_dir.mkdir(parents=True, exist_ok=True)
        stat = dll_path.stat()
        src_hash = _sha256_hex(dll_path)
        hash_part = (src_hash[:12] if src_hash else "nohash")
        staged = cache_dir / f"AcadEventBridge_{hash_part}_{int(stat.st_mtime):x}_{int(stat.st_size):x}.dll"
        need_copy = True
        if staged.exists():
            if src_hash:
                need_copy = _sha256_hex(staged) != src_hash
            else:
                need_copy = staged.stat().st_size != stat.st_size
        if need_copy:
            shutil.copy2(str(dll_path), str(staged))
        load_path = str(staged)

    path_for_lisp = load_path.replace("\\", "/")
    cmd = f'(command "_.NETLOAD" "{path_for_lisp}")\n(princ)'
    bridge.send_command(cmd)
    wr = bridge.wait_for_idle(timeout_sec=30.0, poll_interval_sec=0.2)
    if not wr.completed:
        raise RuntimeError(f"NETLOAD did not complete: needs_input={wr.needs_input}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke-check bridge request/response over named pipe.")
    parser.add_argument("--pid", type=int, default=None, help="AutoCAD PID (optional)")
    parser.add_argument("--pipe-name", type=str, default=None, help="Explicit pipe name (optional)")
    parser.add_argument("--timeout-sec", type=float, default=8.0, help="Hello/heartbeat timeout")
    parser.add_argument("--request-timeout-sec", type=float, default=2.0, help="Request timeout")
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
        if not hello_ok or not heartbeat_ok:
            raise RuntimeError(f"Bridge check failed: hello_ok={hello_ok}, heartbeat_ok={heartbeat_ok}")

        ping = client.request_ping(timeout_sec=float(args.request_timeout_sec))
        status = client.request_status(timeout_sec=float(args.request_timeout_sec))

        if not bool(ping.get("ok", False)):
            raise RuntimeError(f"Ping request failed: {ping}")
        if not bool(status.get("ok", False)):
            raise RuntimeError(f"Status request failed: {status}")

        out = {
            "pipe_name": pipe_name,
            "snapshot": client.snapshot().to_dict(),
            "ping": ping,
            "status": status,
        }
        print(json.dumps(out, ensure_ascii=False, indent=2))
        return 0
    finally:
        client.stop()


if __name__ == "__main__":
    raise SystemExit(main())

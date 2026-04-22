from __future__ import annotations

import argparse
import json
import time
from typing import Optional

from acad_cmd.autocad_bridge import AutoCADBridge
from acad_cmd.bridge_plugin_client import EventBridgeClient
from acad_cmd.command_waiter import CommandWaiter


def _resolve_pipe_name(pid: Optional[int], explicit_name: Optional[str]) -> str:
    if explicit_name:
        return explicit_name
    if pid is None:
        raise RuntimeError("Cannot resolve pipe name: pid is unknown")
    return EventBridgeClient.pipe_name_for_pid(int(pid))


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke check for CommandWaiter event-first completion.")
    parser.add_argument("--pid", type=int, default=None, help="AutoCAD PID (optional)")
    parser.add_argument("--pipe-name", type=str, default=None, help="Explicit pipe name (optional)")
    parser.add_argument("--command", type=str, default="_.REGEN ", help="Command to send to AutoCAD")
    parser.add_argument("--timeout-sec", type=float, default=12.0, help="Wait timeout")
    args = parser.parse_args()

    bridge = AutoCADBridge()
    if not bridge.connect():
        raise RuntimeError("Failed to connect to AutoCAD")

    pid = args.pid
    if pid is None:
        snap = bridge.get_status_snapshot()
        pid_raw = snap.get("acad_pid")
        pid = int(pid_raw) if pid_raw is not None else None

    pipe_name = _resolve_pipe_name(pid=pid, explicit_name=args.pipe_name)
    client = EventBridgeClient(pipe_name=pipe_name, connect_timeout_sec=1.0)
    client.start()

    try:
        if not client.wait_for_hello(timeout_sec=8.0):
            raise RuntimeError("Bridge hello timeout")
        if not client.wait_for_heartbeat(timeout_sec=8.0):
            raise RuntimeError("Bridge heartbeat timeout")

        base_seq = client.snapshot().last_seq
        command_id = bridge.send_command(args.command)

        waiter = CommandWaiter()
        result = waiter.wait_for_completion(
            bridge_client=client,
            after_seq=base_seq,
            timeout_sec=float(args.timeout_sec),
            fallback_wait=lambda t: bridge.wait_for_idle(timeout_sec=t, poll_interval_sec=0.1),
        )

        out = {
            "command_id": command_id,
            "sent": args.command,
            "base_seq": base_seq,
            "wait_result": result.to_dict(),
            "event_bridge_snapshot": client.snapshot().to_dict(),
            "event_state_snapshot": client.event_state_snapshot().to_dict(),
        }
        print(json.dumps(out, ensure_ascii=False, indent=2))

        if not result.completed:
            raise RuntimeError(f"CommandWaiter did not complete: {result}")
        if result.source != "event_stream":
            raise RuntimeError(f"CommandWaiter did not use event stream: source={result.source}")
        if result.completion_event not in {
            "command_ended",
            "command_cancelled",
            "command_failed",
            "lisp_ended",
            "lisp_cancelled",
        }:
            raise RuntimeError(f"Unexpected completion event: {result.completion_event}")

        # Short delay to let following heartbeats settle before returning.
        time.sleep(0.1)
        return 0
    finally:
        client.stop()


if __name__ == "__main__":
    raise SystemExit(main())

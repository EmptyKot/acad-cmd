from __future__ import annotations

import argparse
import json
import os

# Must be set before importing tools module (AppState reads env at import time).
os.environ["AUTOCAD_MCP_EVENT_BRIDGE_ENABLED"] = "1"

from acad_cmd import tools  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke check for send_command event-first waiting.")
    parser.add_argument("--command", type=str, default="_.REGEN ", help="Command to send")
    parser.add_argument("--timeout-sec", type=float, default=12.0, help="Command timeout")
    args = parser.parse_args()

    status_before = tools.get_status(None)
    event_bridge_before = status_before.get("event_bridge") or {}
    if not bool(event_bridge_before.get("connected")):
        raise RuntimeError(f"Bridge is not connected in get_status: {event_bridge_before}")

    result = tools.send_command(
        None,
        command=args.command,
        timeout_sec=float(args.timeout_sec),
        wait=True,
        poll_interval_sec=0.1,
    )

    status_after = tools.get_status(None)
    out = {
        "status_before_event_bridge": event_bridge_before,
        "send_command_result": {
            "completed": result.get("completed"),
            "needs_input": result.get("needs_input"),
            "wait_source": result.get("wait_source"),
            "wait_completion_event": result.get("wait_completion_event"),
            "wait_completion_seq": result.get("wait_completion_seq"),
            "wait_started_seen": result.get("wait_started_seen"),
            "wait_fallback_used": result.get("wait_fallback_used"),
        },
        "status_after_event_bridge": status_after.get("event_bridge"),
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))

    if not bool(result.get("completed")):
        raise RuntimeError(f"Command did not complete: {result}")
    if str(result.get("wait_source") or "") != "event_stream":
        raise RuntimeError(f"send_command did not use event stream: {result}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

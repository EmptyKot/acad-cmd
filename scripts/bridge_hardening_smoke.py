from __future__ import annotations

import json
import os
import time

# Must be set before importing tools module (AppState reads env at import time).
os.environ["AUTOCAD_MCP_EVENT_BRIDGE_ENABLED"] = "1"

from acad_cmd import tools  # noqa: E402


def _compact(result: dict) -> dict:
    return {
        "completed": result.get("completed"),
        "needs_input": result.get("needs_input"),
        "wait_source": result.get("wait_source"),
        "wait_completion_event": result.get("wait_completion_event"),
        "wait_fallback_used": result.get("wait_fallback_used"),
        "bridge_wait_prepare_issue": result.get("bridge_wait_prepare_issue"),
    }


def _send_regen(timeout_sec: float = 10.0) -> dict:
    return tools.send_command(
        None,
        command="_.REGEN ",
        timeout_sec=timeout_sec,
        wait=True,
        poll_interval_sec=0.1,
    )


def main() -> int:
    status0 = tools.get_status(None)
    eb0 = status0.get("event_bridge") or {}
    if not bool(eb0.get("connected")):
        raise RuntimeError(f"Bridge is not connected before hardening smoke: {eb0}")

    # 1) Reconnect smoke: force local client stop, then verify send_command still works.
    if tools.state.event_bridge_client is not None:
        tools.state.event_bridge_client.stop()
    reconnect_result = _send_regen(timeout_sec=12.0)

    # 2) Heartbeat-timeout degradation smoke.
    old_timeout = float(tools.state.event_bridge_heartbeat_timeout_sec)
    tools.state.event_bridge_heartbeat_timeout_sec = 0.1
    timeout_result = None
    for _ in range(5):
        time.sleep(0.35)
        candidate = _send_regen(timeout_sec=8.0)
        if str(candidate.get("wait_source") or "").startswith("fallback_"):
            timeout_result = candidate
            break
    if timeout_result is None:
        timeout_result = candidate

    # 3) Restore timeout and verify recovery back to event stream.
    tools.state.event_bridge_heartbeat_timeout_sec = old_timeout
    time.sleep(0.4)
    recovery_result = _send_regen(timeout_sec=12.0)

    status_end = tools.get_status(None)
    out = {
        "status_before_event_bridge": eb0,
        "reconnect_result": _compact(reconnect_result),
        "heartbeat_timeout_result": _compact(timeout_result),
        "recovery_result": _compact(recovery_result),
        "status_after_event_bridge": status_end.get("event_bridge"),
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))

    if not bool(reconnect_result.get("completed")):
        raise RuntimeError(f"Reconnect scenario did not complete: {reconnect_result}")
    if not bool(timeout_result.get("wait_fallback_used")):
        raise RuntimeError(f"Heartbeat-timeout scenario did not degrade to fallback: {timeout_result}")
    if not bool(recovery_result.get("completed")):
        raise RuntimeError(f"Recovery scenario did not complete: {recovery_result}")
    if str(recovery_result.get("wait_source") or "") != "event_stream":
        raise RuntimeError(f"Recovery scenario did not return to event_stream: {recovery_result}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

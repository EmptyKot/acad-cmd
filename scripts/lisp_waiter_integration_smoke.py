from __future__ import annotations

import json
import os
import tempfile

# Must be set before importing tools module (AppState reads env at import time).
os.environ["AUTOCAD_MCP_EVENT_BRIDGE_ENABLED"] = "1"

from acad_cmd import tools  # noqa: E402


LISP_COMPLETION_EVENTS = {"lisp_ended", "lisp_cancelled"}


def _write_temp_lisp_cp1251() -> str:
    fd, path = tempfile.mkstemp(prefix="aeb_step15_", suffix=".lsp")
    os.close(fd)
    script = '(prompt "\\n[MCP:LOAD-LISP-OK]")\n(princ)\n'
    with open(path, "wb") as f:
        f.write(script.encode("cp1251", errors="replace"))
    return path


def _compact_result(r: dict) -> dict:
    return {
        "completed": r.get("completed"),
        "needs_input": r.get("needs_input"),
        "wait_source": r.get("wait_source"),
        "wait_completion_event": r.get("wait_completion_event"),
        "wait_completion_seq": r.get("wait_completion_seq"),
        "wait_started_seen": r.get("wait_started_seen"),
        "wait_fallback_used": r.get("wait_fallback_used"),
    }


def main() -> int:
    status_before = tools.get_status(None)
    event_bridge = status_before.get("event_bridge") or {}
    if not bool(event_bridge.get("connected")):
        raise RuntimeError(f"Bridge is not connected in get_status: {event_bridge}")

    temp_lsp = _write_temp_lisp_cp1251()
    try:
        load_result = tools.load_lisp_file(None, path=temp_lsp, timeout_sec=12.0, wait=True)
        run_result = tools.run_lisp(None, expr="(princ)", timeout_sec=12.0, wait=True)
    finally:
        try:
            os.remove(temp_lsp)
        except Exception:
            pass

    out = {
        "status_before_event_bridge": event_bridge,
        "load_lisp_file_result": _compact_result(load_result),
        "run_lisp_result": _compact_result(run_result),
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))

    for label, result in (("load_lisp_file", load_result), ("run_lisp", run_result)):
        if not bool(result.get("completed")):
            raise RuntimeError(f"{label} did not complete: {result}")
        if str(result.get("wait_source") or "") != "event_stream":
            raise RuntimeError(f"{label} did not use event stream: {result}")
        if str(result.get("wait_completion_event") or "") not in LISP_COMPLETION_EVENTS:
            raise RuntimeError(f"{label} has non-LISP completion event: {result}")
        if bool(result.get("wait_fallback_used")):
            raise RuntimeError(f"{label} unexpectedly used fallback: {result}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

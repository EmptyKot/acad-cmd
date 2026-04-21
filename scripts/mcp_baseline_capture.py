from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import json
import os
import re

import anyio

from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _tail_text(path: Path, max_chars: int = 8000) -> str:
    try:
        data = path.read_bytes()
    except Exception:
        return ""
    try:
        txt = data.decode("utf-8", "replace")
    except Exception:
        txt = data.decode("latin1", "replace")
    return txt[-max_chars:]


def _unwrap_result(result) -> dict:
    payload = result.structuredContent if result.structuredContent is not None else None
    if payload is None:
        for item in result.content or []:
            txt = getattr(item, "text", None)
            if not isinstance(txt, str):
                continue
            try:
                payload = json.loads(txt)
                break
            except Exception:
                continue
    if isinstance(payload, dict) and isinstance(payload.get("result"), dict):
        payload = payload["result"]
    if not isinstance(payload, dict):
        raise RuntimeError(f"Unexpected tool payload: {result.content!r}")
    return payload


def _is_likely_logfilemode_enabled(text: str) -> bool:
    # Typical command-line echo for (getvar "LOGFILEMODE") includes a standalone "1".
    return bool(re.search(r"(?<!\d)1(?!\d)", text or ""))


async def main() -> None:
    root = Path(__file__).resolve().parents[1]
    py = root / ".venv" / "Scripts" / "python.exe"
    if not py.exists():
        raise SystemExit(f"Missing python exe: {py}")

    out_dir = root / "out" / "baseline"
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / f"acad_cmd_baseline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"

    params = StdioServerParameters(
        command=str(py),
        args=["-m", "acad_cmd.server"],
        env={
            **os.environ,
            "PYTHONPATH": str(root / "src"),
        },
        cwd=str(root),
        encoding="utf-8",
        encoding_error_handler="replace",
    )

    err_path = root / "logs" / "mcp_baseline_capture_stderr.log"
    err_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        err_path.write_text("", encoding="utf-8", errors="ignore")
    except Exception:
        pass

    report: dict = {
        "ts": _utc_now_iso(),
        "roadmap_step": 1,
        "name": "acad-cmd baseline before event bridge changes",
        "checks": {},
        "status": "failed",
    }

    try:
        with err_path.open("a", encoding="utf-8", errors="replace") as err_f:
            async with stdio_client(params, errlog=err_f) as (read_stream, write_stream):
                async with ClientSession(read_stream, write_stream) as session:
                    with anyio.fail_after(30):
                        await session.initialize()

                    with anyio.fail_after(30):
                        listed = await session.list_tools()
                    tool_names = sorted(t.name for t in listed.tools)
                    required = [
                        "get_status",
                        "send_command",
                        "run_lisp",
                        "start_logging",
                        "get_new_output_since",
                        "get_last_output",
                    ]
                    missing = [x for x in required if x not in tool_names]
                    if missing:
                        raise RuntimeError(f"Missing required tools for baseline: {missing}")
                    report["checks"]["tools"] = {"ok": True, "required": required}

                    with anyio.fail_after(30):
                        r_status_1 = await session.call_tool("get_status", {})
                    if r_status_1.isError:
                        raise RuntimeError(f"get_status failed: {r_status_1.content}")
                    status_1 = _unwrap_result(r_status_1)
                    if not bool(status_1.get("connected")):
                        raise RuntimeError(f"AutoCAD is not connected: {status_1}")
                    report["checks"]["get_status_before"] = {
                        "ok": True,
                        "connected": bool(status_1.get("connected")),
                        "dwg": status_1.get("dwg"),
                        "acadver": status_1.get("acadver"),
                        "acad_pid": status_1.get("acad_pid"),
                        "busy": bool(status_1.get("busy")),
                        "stale": bool(status_1.get("stale")),
                        "source": status_1.get("source"),
                        "event_bridge": status_1.get("event_bridge"),
                    }

                    with anyio.fail_after(30):
                        r_start = await session.call_tool("start_logging", {"mode": "logfile", "reset": False})
                    if r_start.isError:
                        raise RuntimeError(f"start_logging(logfile) failed: {r_start.content}")
                    started = _unwrap_result(r_start)
                    stream_id = started.get("stream_id")
                    cursor0 = int(started.get("cursor", 0) or 0)
                    logfile_path = started.get("logfile_path")
                    if not stream_id:
                        raise RuntimeError(f"start_logging returned empty stream_id: {started}")
                    report["checks"]["start_logging"] = {
                        "ok": True,
                        "stream_id": stream_id,
                        "cursor": cursor0,
                        "logfile_path": logfile_path,
                    }

                    send_expr = '(princ "\\n[MCP_BASELINE_SEND_COMMAND]\\n")(princ)'
                    with anyio.fail_after(45):
                        r_send = await session.call_tool(
                            "send_command",
                            {"command": send_expr, "wait": True, "timeout_sec": 10.0, "poll_interval_sec": 0.1},
                        )
                    if r_send.isError:
                        raise RuntimeError(f"send_command failed: {r_send.content}")
                    send_payload = _unwrap_result(r_send)
                    log_block = send_payload.get("log") or {}
                    send_log_text = str(log_block.get("text") or "")
                    report["checks"]["send_command"] = {
                        "ok": bool(send_payload.get("completed")) and not bool(send_payload.get("needs_input")),
                        "completed": bool(send_payload.get("completed")),
                        "needs_input": bool(send_payload.get("needs_input")),
                        "command_id": send_payload.get("command_id"),
                        "log_bytes": len(send_log_text.encode("utf-8", errors="replace")),
                        "marker_seen": "[MCP_BASELINE_SEND_COMMAND]" in send_log_text,
                    }

                    with anyio.fail_after(45):
                        r_lisp = await session.call_tool(
                            "run_lisp",
                            {"expr": "(getvar 'ACADVER)", "wait": True, "timeout_sec": 10.0},
                        )
                    if r_lisp.isError:
                        raise RuntimeError(f"run_lisp(ACADVER) failed: {r_lisp.content}")
                    lisp_payload = _unwrap_result(r_lisp)
                    report["checks"]["run_lisp"] = {
                        "ok": bool(lisp_payload.get("completed")) and not bool(lisp_payload.get("needs_input")),
                        "completed": bool(lisp_payload.get("completed")),
                        "needs_input": bool(lisp_payload.get("needs_input")),
                        "marker_id": lisp_payload.get("marker_id"),
                        "last_prompt": lisp_payload.get("last_prompt"),
                    }

                    # Explicit LOGFILEMODE probe: capture command line echo via logfile.
                    with anyio.fail_after(45):
                        r_lisp_logmode = await session.call_tool(
                            "run_lisp",
                            {"expr": '(getvar "LOGFILEMODE")', "wait": True, "timeout_sec": 10.0},
                        )
                    if r_lisp_logmode.isError:
                        raise RuntimeError(f"run_lisp(LOGFILEMODE) failed: {r_lisp_logmode.content}")
                    lisp_logmode_payload = _unwrap_result(r_lisp_logmode)
                    lisp_logmode_text = str((lisp_logmode_payload.get("log") or {}).get("text") or "")

                    with anyio.fail_after(30):
                        r_new = await session.call_tool(
                            "get_new_output_since",
                            {"stream_id": stream_id, "cursor": cursor0, "max_bytes": 65536},
                        )
                    if r_new.isError:
                        raise RuntimeError(f"get_new_output_since failed: {r_new.content}")
                    new_payload = _unwrap_result(r_new)
                    new_text = str(new_payload.get("text") or "")
                    report["checks"]["logfile_growth"] = {
                        "ok": int(new_payload.get("new_cursor") or cursor0) > cursor0,
                        "cursor_before": cursor0,
                        "cursor_after": int(new_payload.get("new_cursor") or cursor0),
                        "bytes": len(new_text.encode("utf-8", errors="replace")),
                        "truncated": bool(new_payload.get("truncated")),
                    }
                    report["checks"]["logfilemode_probe"] = {
                        "ok": _is_likely_logfilemode_enabled(lisp_logmode_text + "\n" + new_text),
                        "probe": '(getvar "LOGFILEMODE")',
                    }

                    with anyio.fail_after(30):
                        r_last = await session.call_tool("get_last_output", {"source": "logfile"})
                    if r_last.isError:
                        raise RuntimeError(f"get_last_output(logfile) failed: {r_last.content}")
                    last_payload = _unwrap_result(r_last)
                    report["checks"]["get_last_output_logfile"] = {
                        "ok": isinstance(last_payload.get("text"), str),
                        "source": last_payload.get("source"),
                        "bytes": len(str(last_payload.get("text") or "").encode("utf-8", errors="replace")),
                    }

                    with anyio.fail_after(30):
                        r_status_2 = await session.call_tool("get_status", {})
                    if r_status_2.isError:
                        raise RuntimeError(f"get_status(after) failed: {r_status_2.content}")
                    status_2 = _unwrap_result(r_status_2)
                    report["checks"]["get_status_after"] = {
                        "ok": bool(status_2.get("connected")),
                        "connected": bool(status_2.get("connected")),
                        "dwg": status_2.get("dwg"),
                        "acadver": status_2.get("acadver"),
                        "acad_pid": status_2.get("acad_pid"),
                        "busy": bool(status_2.get("busy")),
                        "stale": bool(status_2.get("stale")),
                        "source": status_2.get("source"),
                        "event_bridge": status_2.get("event_bridge"),
                    }

                    report["status"] = "passed"
    except Exception as exc:
        report["status"] = "failed"
        report["error"] = f"{type(exc).__name__}: {exc}"
        tail = _tail_text(err_path)
        if tail.strip():
            report["server_stderr_tail"] = tail
        raise
    finally:
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"baseline report: {report_path}")
        print(f"status: {report.get('status')}")


if __name__ == "__main__":
    anyio.run(main)

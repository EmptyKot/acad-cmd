import os
import uuid
from dataclasses import dataclass
from typing import Any, Dict, Optional

from mcp.server.fastmcp import FastMCP, Context

from .autocad_bridge import AutoCADBridge
from .lisp import build_load_lisp_command, build_run_lisp_script, lisp_quote_string
from .output_log import OutputStreamManager
from .session_log import SessionLogger, iso_now


DEFAULT_LOG_DIR = os.path.join(os.getcwd(), "logs", "acad-cmd")


@dataclass
class AppState:
    session_id: str
    bridge: AutoCADBridge
    streams: OutputStreamManager
    audit: SessionLogger


def _make_state() -> AppState:
    session_id = str(uuid.uuid4())
    base_dir = os.path.join(DEFAULT_LOG_DIR, session_id)
    os.makedirs(base_dir, exist_ok=True)
    audit_path = os.path.join(base_dir, "session.jsonl")
    return AppState(
        session_id=session_id,
        bridge=AutoCADBridge(),
        streams=OutputStreamManager(base_dir=base_dir),
        audit=SessionLogger(path=audit_path, session_id=session_id),
    )


state = _make_state()
mcp = FastMCP("acad-cmd")


def _ensure_connected() -> None:
    ok = state.bridge.ensure_connection()
    if not ok:
        raise RuntimeError("Failed to connect to AutoCAD via COM")


def _default_logfile_path() -> str:
    return os.path.join(state.streams.base_dir, "acad-commandline.log")


def _get_current_logfilename() -> Optional[str]:
    try:
        v = state.bridge.get_variable("LOGFILENAME")
        s = str(v) if v is not None else ""
        return s or None
    except Exception:
        return None


@mcp.tool()
def get_status(ctx: Context) -> Dict[str, Any]:
    connected = state.bridge.ensure_connection()
    dwg = state.bridge.get_dwg_label() if connected else None
    acadver = None
    hwnd = None
    pid = None
    if connected:
        try:
            acadver = str(state.bridge.get_variable("ACADVER"))
        except Exception:
            acadver = None
        try:
            hwnd = int(getattr(state.bridge.acad, "HWND", 0) or 0)
        except Exception:
            hwnd = None
        if hwnd:
            try:
                import win32process

                _tid, pidv = win32process.GetWindowThreadProcessId(hwnd)
                pid = int(pidv)
            except Exception:
                pid = None
    default_stream = state.streams.get_default()
    return {
        "ts": iso_now(),
        "session_id": state.session_id,
        "connected": connected,
        "dwg": dwg,
        "acadver": acadver,
        "acad_hwnd": hwnd,
        "acad_pid": pid,
        "default_stream": (
            {
                "stream_id": default_stream.stream_id,
                "mode": default_stream.mode,
                "logfile_path": default_stream.logfile_path,
                "cursor": default_stream.cursor,
            }
            if default_stream
            else None
        ),
    }


@mcp.tool()
def start_logging(
    ctx: Context,
    mode: str = "logfile",
    logfile_path: Optional[str] = None,
    reset: bool = False,
) -> Dict[str, Any]:
    _ensure_connected()
    if mode not in ("logfile", "lastprompt"):
        raise ValueError("mode must be 'logfile' or 'lastprompt'")

    stream_id = str(uuid.uuid4())
    dwg = state.bridge.get_dwg_label()

    if mode == "lastprompt":
        # Logical stream for clients that only want LASTPROMPT.
        state.streams.start_lastprompt_stream(stream_id=stream_id)
        state.audit.log("start_logging", {"mode": mode}, dwg=dwg)
        return {"stream_id": stream_id, "mode": mode, "logfile_path": None, "cursor": 0}

    # Choose logfile path.
    # If caller didn't provide a path, prefer AutoCAD's current LOGFILENAME.
    # This avoids issues where AutoCAD refuses to write to paths with
    # non-ASCII characters (common when the workspace path contains Cyrillic).
    path = logfile_path
    if not path:
        path = _get_current_logfilename() or _default_logfile_path()

    if path:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
        except Exception:
            pass

    # Enable AutoCAD logfile output.
    # LOGFILENAME must be set before enabling LOGFILEMODE in some setups.
    if logfile_path:
        # Only attempt to override LOGFILENAME if the user explicitly asked.
        try:
            state.bridge.set_variable("LOGFILENAME", path)
            state.bridge.set_variable("LOGFILEMODE", 1)
        except Exception:
            # Fallback via AutoLISP setvar (some environments block COM SetVariable).
            path_norm = path.replace("\\", "/")
            lsp = "\n".join(
                [
                    f'(setvar "LOGFILENAME" "{lisp_quote_string(path_norm)}")',
                    '(setvar "LOGFILEMODE" 1)',
                    '(princ)',
                ]
            )
            state.bridge.send_command(lsp)
    else:
        # Keep current LOGFILENAME; just ensure LOGFILEMODE is enabled.
        try:
            state.bridge.set_variable("LOGFILEMODE", 1)
        except Exception:
            try:
                state.bridge.send_command('(setvar "LOGFILEMODE" 1)\n(princ)')
            except Exception:
                pass

    # Refresh the effective path (AutoCAD may normalize/override it).
    if not logfile_path:
        path = _get_current_logfilename() or path

    cursor = 0
    if not reset and os.path.exists(path):
        try:
            cursor = os.path.getsize(path)
        except Exception:
            cursor = 0

    state.streams.start_logfile_stream(
        stream_id=stream_id,
        logfile_path=path,
        cursor=cursor,
        started_by_server=True,
    )

    state.audit.log("start_logging", {"mode": mode, "logfile_path": path, "cursor": cursor}, dwg=dwg)
    return {"stream_id": stream_id, "mode": mode, "logfile_path": path, "cursor": cursor}


@mcp.tool()
def stop_logging(ctx: Context, stream_id: str) -> Dict[str, Any]:
    _ensure_connected()
    s = state.streams.get(stream_id)
    stopped = state.streams.stop(stream_id)

    # Best-effort: if we stopped a logfile stream started by us and
    # there are no remaining logfile streams, disable AutoCAD logging.
    if stopped and s and s.mode == "logfile" and s.started_by_server:
        remaining_logfile = False
        default_stream = state.streams.get_default()
        if default_stream and default_stream.mode == "logfile":
            remaining_logfile = True
        if not remaining_logfile:
            try:
                state.bridge.set_variable("LOGFILEMODE", 0)
            except Exception:
                # Fallback via AutoLISP
                try:
                    state.bridge.send_command('(setvar "LOGFILEMODE" 0)\n(princ)')
                except Exception:
                    pass

    dwg = state.bridge.get_dwg_label()
    state.audit.log("stop_logging", {"stream_id": stream_id, "stopped": stopped}, dwg=dwg)
    return {"stream_id": stream_id, "stopped": stopped}


@mcp.tool()
def get_new_output_since(
    ctx: Context,
    stream_id: str,
    cursor: int,
    max_bytes: int = 65536,
) -> Dict[str, Any]:
    _ensure_connected()
    text, new_cursor, truncated = state.streams.read_new(stream_id, cursor, max_bytes)
    dwg = state.bridge.get_dwg_label()
    state.audit.log(
        "get_new_output_since",
        {"stream_id": stream_id, "cursor": cursor, "new_cursor": new_cursor, "bytes": len(text)},
        dwg=dwg,
    )
    return {
        "dwg": dwg,
        "text": text,
        "new_cursor": new_cursor,
        "truncated": truncated,
    }


@mcp.tool()
def get_last_output(ctx: Context, source: str = "lastprompt") -> Dict[str, Any]:
    _ensure_connected()
    dwg = state.bridge.get_dwg_label()

    if source == "logfile":
        s = state.streams.get_default()
        if not s:
            return {"dwg": dwg, "text": "", "timestamp": iso_now(), "source": source}
        text = state.streams.read_tail(s.stream_id)
        state.audit.log("get_last_output", {"source": source, "bytes": len(text)}, dwg=dwg)
        return {"dwg": dwg, "text": text, "timestamp": iso_now(), "source": source}

    text = state.bridge.get_last_prompt()
    state.audit.log("get_last_output", {"source": "lastprompt", "bytes": len(text)}, dwg=dwg)
    return {"dwg": dwg, "text": text, "timestamp": iso_now(), "source": "lastprompt"}


@mcp.tool()
def send_command(
    ctx: Context,
    command: str,
    wait: bool = True,
    timeout_sec: float = 10.0,
    poll_interval_sec: float = 0.1,
) -> Dict[str, Any]:
    _ensure_connected()
    dwg = state.bridge.get_dwg_label()
    command_id = state.bridge.send_command(command)

    state.audit.log(
        "send_command",
        {"command_id": command_id, "command": command, "wait": wait, "timeout_sec": timeout_sec},
        dwg=dwg,
    )

    completed = True
    needs_input = False

    if wait:
        wr = state.bridge.wait_for_idle(timeout_sec=timeout_sec, poll_interval_sec=poll_interval_sec)
        completed = wr.completed
        needs_input = wr.needs_input

    last_prompt = state.bridge.get_last_prompt()

    # If we have an active logfile stream, also return new output.
    stream = state.streams.get_default()
    log_block = None
    if stream and stream.mode == "logfile" and stream.logfile_path:
        text, new_cursor, truncated = state.streams.read_new(stream.stream_id, stream.cursor, 65536)
        log_block = {
            "stream_id": stream.stream_id,
            "cursor": new_cursor,
            "text": text,
            "truncated": truncated,
        }

    state.audit.log(
        "send_command_result",
        {
            "command_id": command_id,
            "completed": completed,
            "needs_input": needs_input,
            "last_prompt": last_prompt,
            "has_log": bool(log_block),
        },
        dwg=dwg,
    )

    return {
        "command_id": command_id,
        "dwg": dwg,
        "sent": command,
        "completed": completed,
        "needs_input": needs_input,
        "last_prompt": last_prompt,
        "log": log_block,
    }


@mcp.tool()
def load_lisp_file(
    ctx: Context,
    path: str,
    wait: bool = True,
    timeout_sec: float = 10.0,
) -> Dict[str, Any]:
    _ensure_connected()
    dwg = state.bridge.get_dwg_label()
    cmd = build_load_lisp_command(path)
    state.audit.log("load_lisp_file", {"path": path, "command": cmd}, dwg=dwg)
    return send_command(ctx, cmd, wait=wait, timeout_sec=timeout_sec)


@mcp.tool()
def run_lisp(
    ctx: Context,
    expr: str,
    wait: bool = True,
    timeout_sec: float = 10.0,
) -> Dict[str, Any]:
    _ensure_connected()
    dwg = state.bridge.get_dwg_label()
    marker_id = str(uuid.uuid4())
    script = build_run_lisp_script(expr, marker_id)
    state.audit.log("run_lisp", {"expr": expr, "marker_id": marker_id}, dwg=dwg)
    result = send_command(ctx, script, wait=wait, timeout_sec=timeout_sec)
    result["marker_id"] = marker_id
    return result


def main() -> None:
    # Run MCP over stdio (FastMCP default)
    mcp.run()


if __name__ == "__main__":
    main()

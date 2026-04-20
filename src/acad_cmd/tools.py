import os

import uuid
from dataclasses import dataclass
from typing import Any, Dict, Optional

from mcp.server.fastmcp import FastMCP, Context

from .autocad_bridge import AutoCADBridge
from .lisp import build_load_lisp_command, build_run_lisp_script, lisp_quote_string
from .output_log import OutputStreamManager
from .session_log import SessionLogger, iso_now
from .lisp_lib import (
    MCP_DICT_LISP_LIB as _MCP_DICT_LISP_LIB,
    MCP_SELECTION_LISP_LIB as _MCP_SELECTION_LISP_LIB,
    lisp_concat as _lisp_concat,
    lisp_string as _lisp_string,
    lisp_typed_values as _lisp_typed_values,
    strip_ok as _strip_ok,
)
from .protocol import extract_mcp_json as _extract_mcp_json
from .selection import collect_selection_stream_lite as _collect_selection_stream_lite


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


def _ensure_logfile_stream(ctx: Context) -> Optional[str]:
    """Ensure default output stream is a logfile stream; return temp stream_id if created."""

    s = state.streams.get_default()
    if s and s.mode == "logfile" and s.logfile_path:
        return None
    r = start_logging(ctx, mode="logfile")
    return str(r.get("stream_id"))


def _run_lisp_json(ctx: Context, expr: str, *, timeout_sec: float = 10.0) -> Dict[str, Any]:
    """Run a LISP expr that prints one [MCP:JSON]{...} line."""

    temp_stream_id = _ensure_logfile_stream(ctx)
    try:
        r = run_lisp(ctx, expr, wait=True, timeout_sec=timeout_sec)
        log_block = r.get("log") or {}
        text = str(log_block.get("text") or "")
        try:
            obj = _extract_mcp_json(text)
        except Exception:
            # Fallback: sometimes the logfile chunk returned by send_command()
            # does not include the marker yet. Try the logfile tail.
            tail = get_last_output(ctx, source="logfile")
            obj = _extract_mcp_json(str(tail.get("text") or ""))
        if obj.get("ok") is False:
            msg = obj.get("error") or "Unknown AutoLISP error"
            raise RuntimeError(str(msg))
        return obj
    finally:
        if temp_stream_id:
            try:
                stop_logging(ctx, temp_stream_id)
            except Exception:
                pass


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


@mcp.tool()
def dict_list(ctx: Context) -> Dict[str, Any]:
    """List top-level dictionaries from Named Objects Dictionary."""

    _ensure_connected()
    expr = _lisp_concat(_MCP_DICT_LISP_LIB, "(mcp-dict-list)\n")
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def dict_keys(ctx: Context, dict_name: str) -> Dict[str, Any]:
    """List keys (and entry types) in a named dictionary."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    expr = _lisp_concat(_MCP_DICT_LISP_LIB, f"(mcp-dict-keys {_lisp_string(dict_name)})\n")
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def dict_xrecord_get(ctx: Context, dict_name: str, key: str) -> Dict[str, Any]:
    """Read XRecord data from a named dictionary by key."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    if not key:
        raise ValueError("key must be non-empty")
    expr = _lisp_concat(
        _MCP_DICT_LISP_LIB,
        f"(mcp-xrecord-get {_lisp_string(dict_name)} {_lisp_string(key)})\n",
    )
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def dict_xrecord_set(
    ctx: Context,
    dict_name: str,
    key: str,
    values: Any,
    overwrite: bool = True,
) -> Dict[str, Any]:
    """Write XRecord data into a named dictionary under key."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    if not key:
        raise ValueError("key must be non-empty")
    values_expr = _lisp_typed_values(values)
    ow = "T" if overwrite else "nil"
    expr = _lisp_concat(
        _MCP_DICT_LISP_LIB,
        f"(mcp-xrecord-set {_lisp_string(dict_name)} {_lisp_string(key)} {values_expr} {ow})\n",
    )
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def dict_xrecord_delete(ctx: Context, dict_name: str, key: str) -> Dict[str, Any]:
    """Delete an XRecord entry from a named dictionary."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    if not key:
        raise ValueError("key must be non-empty")
    expr = _lisp_concat(
        _MCP_DICT_LISP_LIB,
        f"(mcp-xrecord-delete {_lisp_string(dict_name)} {_lisp_string(key)})\n",
    )
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def dict_delete(ctx: Context, dict_name: str, recursive: bool = True) -> Dict[str, Any]:
    """Delete a named dictionary from the Named Objects Dictionary."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    rec = "T" if recursive else "nil"
    expr = _lisp_concat(_MCP_DICT_LISP_LIB, f"(mcp-dict-delete {_lisp_string(dict_name)} {rec})\n")
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def selection(
    ctx: Context,
    timeout_sec: float = 300.0,
    prompt: Optional[str] = None,
    filter: Any = None,
    max_objects: Optional[int] = None,
    alert_message: Optional[str] = None,
) -> Dict[str, Any]:
    """Get currently selected objects, or prompt the user to select objects.

    Returns only handle + type for each selected object.
    """

    _ensure_connected()
    dwg = state.bridge.get_dwg_label()

    temp_stream_id = _ensure_logfile_stream(ctx)
    try:
        stream = state.streams.get_default()
        if not stream:
            raise RuntimeError("No default stream")

        mo = int(max_objects) if max_objects is not None else -1

        # 1) Try implied (PickFirst) selection.
        req_id1 = str(uuid.uuid4())
        cursor0 = int(stream.cursor)
        expr1 = _lisp_concat(
            _MCP_SELECTION_LISP_LIB,
            f"(mcp-selection-implied-lite {_lisp_string(req_id1)} {mo})\n",
        )
        r1 = send_command(ctx, expr1, wait=True, timeout_sec=min(10.0, float(timeout_sec)))
        log_block1 = r1.get("log") or {}
        initial_text1 = str(log_block1.get("text") or "")
        cursor1 = log_block1.get("cursor")
        out1 = _collect_selection_stream_lite(
            streams=state.streams,
            stream_id=stream.stream_id,
            req_id=req_id1,
            timeout_sec=min(10.0, float(timeout_sec)),
            initial_text=initial_text1,
            cursor=int(cursor1) if cursor1 is not None else cursor0,
        )
        out1["dwg"] = dwg
        state.audit.log(
            "selection",
            {
                "phase": "implied",
                "req_id": req_id1,
                "max_objects": max_objects,
                "count": out1.get("count"),
                "timed_out": out1.get("timed_out"),
            },
            dwg=dwg,
        )

        if not out1.get("timed_out") and int(out1.get("count") or 0) > 0:
            return out1

        # 2) If nothing selected, prompt interactively.
        try:
            cmdactive = int(state.bridge.get_variable("CMDACTIVE") or 0)
        except Exception:
            cmdactive = 0
        if cmdactive != 0:
            raise RuntimeError(f"AutoCAD is busy (CMDACTIVE={cmdactive}); cannot prompt for selection")

        req_id2 = str(uuid.uuid4())
        prompt_expr = _lisp_string(prompt) if prompt else "nil"
        alert_expr = _lisp_string(alert_message) if alert_message else "nil"
        filter_expr = _lisp_typed_values(filter) if filter is not None else "nil"

        expr2 = _lisp_concat(
            _MCP_SELECTION_LISP_LIB,
            f"(mcp-selection-prompt-lite {_lisp_string(req_id2)} {prompt_expr} {alert_expr} {filter_expr} {mo})\n",
        )

        # Critical: interactive ssget must be the last input in this SendCommand.
        r2 = send_command(ctx, expr2, wait=False, timeout_sec=0.1)
        log_block2 = r2.get("log") or {}
        initial_text2 = str(log_block2.get("text") or "")
        cursor2 = log_block2.get("cursor")
        out2 = _collect_selection_stream_lite(
            streams=state.streams,
            stream_id=stream.stream_id,
            req_id=req_id2,
            timeout_sec=float(timeout_sec),
            initial_text=initial_text2,
            cursor=int(cursor2) if cursor2 is not None else int(out1.get("cursor") or stream.cursor),
        )
        out2["dwg"] = dwg
        state.audit.log(
            "selection",
            {
                "phase": "prompt",
                "req_id": req_id2,
                "timeout_sec": timeout_sec,
                "prompt": prompt,
                "has_alert_message": bool(alert_message),
                "has_filter": filter is not None,
                "max_objects": max_objects,
                "count": out2.get("count"),
                "timed_out": out2.get("timed_out"),
            },
            dwg=dwg,
        )
        return out2
    finally:
        if temp_stream_id:
            try:
                stop_logging(ctx, temp_stream_id)
            except Exception:
                pass


def main() -> None:
    # Run MCP over stdio (FastMCP default)
    mcp.run()


if __name__ == "__main__":
    main()

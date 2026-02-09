import os
import json

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


_MCP_JSON_MARKER = "[MCP:JSON]"


def _extract_mcp_json(text: str) -> Dict[str, Any]:
    """Extract and parse the last MCP JSON marker from logfile output."""

    if not text:
        raise RuntimeError("No output text to parse")

    # Prefer line-based extraction: we intentionally print one marker per line.
    last_line = None
    for line in text.splitlines():
        if _MCP_JSON_MARKER in line:
            last_line = line

    if last_line is None:
        raise RuntimeError("MCP JSON marker not found in output")

    idx = last_line.rfind(_MCP_JSON_MARKER)
    payload = last_line[idx + len(_MCP_JSON_MARKER) :].strip()
    if not payload:
        raise RuntimeError("MCP JSON marker present but payload is empty")

    try:
        obj = json.loads(payload)
    except Exception as e:
        raise RuntimeError(f"Failed to parse MCP JSON payload: {e}")

    if not isinstance(obj, dict):
        raise RuntimeError("MCP JSON payload is not an object")
    return obj


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


_MCP_DICT_LISP_LIB = r"""
(progn
  (if (not (fboundp 'mcp--emit-json))
    (progn
      (defun mcp--json-escape (s / i c out)
        (setq out "")
        (setq i 1)
        (while (<= i (strlen s))
          (setq c (substr s i 1))
          (cond
            ((= c "\\") (setq out (strcat out "\\\\")))
            ((= c "\"") (setq out (strcat out "\\\"")))
            (T (setq out (strcat out c)))
          )
          (setq i (+ i 1))
        )
        out
      )

      (defun mcp--json-quote (s)
        (strcat "\"" (mcp--json-escape s) "\"")
      )

      (defun mcp--json-real (r)
        (vl-string-right-trim "." (vl-string-right-trim "0" (rtos r 2 15)))
      )

      (defun mcp--json-value (v)
        (cond
          ((= v T) "true")
          ((= v nil) "false")
          ((and (= (type v) 'SYM) (= (strcase (vl-symbol-name v)) "MCPNULL")) "null")
          ((= (type v) 'STR) (mcp--json-quote v))
          ((= (type v) 'INT) (itoa v))
          ((= (type v) 'REAL) (mcp--json-real v))
          ((= (type v) 'LIST) (mcp--json-arr v))
          (T (mcp--json-quote (vl-princ-to-string v)))
        )
      )

      (defun mcp--json-arr (lst / out first)
        (setq out "[")
        (setq first T)
        (foreach v lst
          (if first
            (setq first nil)
            (setq out (strcat out ","))
          )
          (setq out (strcat out (mcp--json-value v)))
        )
        (setq out (strcat out "]"))
        out
      )

      (defun mcp--emit-json (json)
        (prompt (strcat "\n" "[MCP:JSON]" json))
        (princ)
      )

      (defun mcp--emit-ok (body)
        (mcp--emit-json (strcat "{\"ok\":true" body "}"))
      )

      (defun mcp--emit-err (msg)
        (mcp--emit-json (strcat "{\"ok\":false,\"error\":" (mcp--json-value msg) "}"))
      )

      (defun mcp--nod () (namedobjdict))

      (defun mcp--is-system-name (name / u)
        (setq u (strcase name))
        (or (wcmatch u "ACAD_*") (wcmatch u "AEC_*") (wcmatch u "ADSK_*") (wcmatch u "A$*"))
      )

      (defun mcp--dict-by-name (name / nod r)
        (setq nod (mcp--nod))
        (setq r (dictsearch nod name))
        (if (and r (= (cdr (assoc 0 r)) "DICTIONARY"))
          (cdr (assoc -1 r))
          nil
        )
      )

      (defun mcp--dict-entry-pairs (d / el out key)
        ;; Returns list of (key . ename) from DICTIONARY entity list.
        (setq el (entget d))
        (setq out nil)
        (while el
          (if (= (caar el) 3)
            (progn
              (setq key (cdar el))
              (setq el (cdr el))
              (while (and el (/= (caar el) 350))
                (setq el (cdr el))
              )
              (if el
                (progn
                  (setq out (cons (cons key (cdar el)) out))
                  (setq el (cdr el))
                )
              )
            )
            (setq el (cdr el))
          )
        )
        (reverse out)
      )

      (defun mcp--ensure-dict (name / d)
        (setq d (mcp--dict-by-name name))
        (if d
          d
          (progn
            (setq d (entmakex (list (cons 0 "DICTIONARY") (cons 100 "AcDbDictionary"))))
            (dictadd (mcp--nod) name d)
            d
          )
        )
      )

      (defun mcp--xrec-by-key (d key / r e)
        (setq r (dictsearch d key))
        (if (and r (= (cdr (assoc 0 r)) "XRECORD"))
          (cdr (assoc -1 r))
          nil
        )
      )

      (defun mcp--xrec-filter-pairs (pairs / out)
        (setq out nil)
        (foreach p pairs
          (if (and (numberp (car p))
                   (>= (car p) 1)
                   (/= (car p) 5)
                   (/= (car p) 100)
                   (/= (car p) 102)
                   (/= (car p) 280)
                   (/= (car p) 330)
                   (/= (car p) 360))
            (setq out (cons p out))
          )
        )
        (reverse out)
      )

      (defun mcp--xrec-read (e / pairs)
        (setq pairs (entget e))
        (mcp--xrec-filter-pairs pairs)
      )

      (defun mcp--json-xrec-values (pairs / out first)
        ;; pairs: list of (code . value) -> JSON [[code,value],...]
        (setq out "[")
        (setq first T)
        (foreach p pairs
          (if first
            (setq first nil)
            (setq out (strcat out ","))
          )
          (setq out (strcat out "[" (itoa (car p)) "," (mcp--json-value (cdr p)) "]"))
        )
        (setq out (strcat out "]"))
        out
      )

      (defun mcp--dicts-json (/ nod it out first name obj etype isSys reason)
        (setq nod (mcp--nod))
        (setq it (mcp--dict-entry-pairs nod))
        (setq out "[")
        (setq first T)
        (foreach kv it
          (setq name (car kv))
          (setq obj (cdr kv))
          (setq etype (if obj (cdr (assoc 0 (entget obj))) ""))
          (if (= etype "DICTIONARY")
            (progn
              (setq isSys (mcp--is-system-name name))
              (setq reason (if isSys "prefix" 'MCPNULL))
              (if first
                (setq first nil)
                (setq out (strcat out ","))
              )
              (setq out
                (strcat out
                  "{\"name\":" (mcp--json-value name)
                  ",\"is_system_guess\":" (mcp--json-value isSys)
                  ",\"system_reason\":" (mcp--json-value reason)
                  "}"
                )
              )
            )
          )
        )
        (setq out (strcat out "]"))
        out
      )

      (defun mcp-dict-list ()
        (mcp--emit-json (strcat "{\"ok\":true,\"dicts\":" (mcp--dicts-json) "}"))
      )

      (defun mcp-dict-keys (dictName / d it entries keys first k obj etype)
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"found\":false,\"keys\":[],\"entries\":[]}")
          (progn
            (setq it (mcp--dict-entry-pairs d))
            (setq entries "[")
            (setq keys "[")
            (setq first T)
            (foreach kv it
              (setq k (car kv))
              (setq obj (cdr kv))
              (setq etype (if obj (cdr (assoc 0 (entget obj))) 'MCPNULL))
              (if first
                (setq first nil)
                (progn
                  (setq entries (strcat entries ","))
                  (setq keys (strcat keys ","))
                )
              )
              (setq entries (strcat entries "{\"key\":" (mcp--json-value k) ",\"type\":" (mcp--json-value etype) "}"))
              (setq keys (strcat keys (mcp--json-value k)))
            )
            (setq entries (strcat entries "]"))
            (setq keys (strcat keys "]"))
            (mcp--emit-json (strcat "{\"ok\":true,\"found\":true,\"keys\":" keys ",\"entries\":" entries "}"))
          )
        )
      )

      (defun mcp-xrecord-get (dictName key / d x pairs)
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"found\":false,\"values\":[]}")
          (progn
            (setq x (mcp--xrec-by-key d key))
            (if (not x)
              (mcp--emit-json "{\"ok\":true,\"found\":false,\"values\":[]}")
              (progn
                (setq pairs (mcp--xrec-read x))
                (mcp--emit-json (strcat "{\"ok\":true,\"found\":true,\"values\":" (mcp--json-xrec-values pairs) "}"))
              )
            )
          )
        )
      )

      (defun mcp-xrecord-set (dictName key values overwrite / d old xrec)
        (setq d (mcp--ensure-dict dictName))
        (setq old (mcp--xrec-by-key d key))
        (if old
          (if overwrite
            (progn
              (dictremove d key)
              (entdel old)
            )
            (progn
              (mcp--emit-err "Key already exists")
              (setq d nil)
            )
          )
        )
        (if d
          (progn
            (setq xrec (entmakex (append (list (cons 0 "XRECORD") (cons 100 "AcDbXrecord")) values)))
            (dictadd d key xrec)
            (mcp--emit-json "{\"ok\":true,\"written\":true}")
          )
        )
      )

      (defun mcp-xrecord-delete (dictName key / d old)
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"deleted\":false}")
          (progn
            (setq old (mcp--xrec-by-key d key))
            (if (not old)
              (mcp--emit-json "{\"ok\":true,\"deleted\":false}")
              (progn
                (dictremove d key)
                (entdel old)
                (mcp--emit-json "{\"ok\":true,\"deleted\":true}")
              )
            )
          )
        )
      )

      (defun mcp-dict-delete (dictName recursive / nod d it k obj n)
        (setq nod (mcp--nod))
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"deleted\":false,\"deleted_entries\":0}")
          (progn
            (setq n 0)
            (setq it (mcp--dict-entry-pairs d))
            (if (and (not recursive) it)
              (progn
                (mcp--emit-err "Dictionary not empty (set recursive=true to delete)")
                (setq d nil)
              )
            )
            (if d
              (progn
                (foreach kv it
                  (setq k (car kv))
                  (setq obj (cdr kv))
                  (if k (dictremove d k))
                  (if obj (entdel obj))
                  (setq n (+ n 1))
                )
                (dictremove nod dictName)
                (entdel d)
                (mcp--emit-json (strcat "{\"ok\":true,\"deleted\":true,\"deleted_entries\":" (itoa n) "}"))
              )
            )
          )
        )
      )
    )
  )
)
"""


def _lisp_string(s: str) -> str:
    return '"' + lisp_quote_string(s) + '"'


def _lisp_typed_values(values: Any) -> str:
    """Convert [{code,value},...] into a LISP list of dotted pairs."""

    if values is None:
        return "'()"
    if not isinstance(values, list):
        raise ValueError("values must be a list")

    parts: list[str] = []
    for i, item in enumerate(values):
        if not isinstance(item, dict):
            raise ValueError(f"values[{i}] must be an object")
        if "code" not in item or "value" not in item:
            raise ValueError(f"values[{i}] must have 'code' and 'value'")
        code = item["code"]
        val = item["value"]
        if not isinstance(code, int):
            raise ValueError(f"values[{i}].code must be integer")

        if isinstance(val, str):
            v = _lisp_string(val)
        elif isinstance(val, bool):
            v = "T" if val else "nil"
        elif isinstance(val, int) or isinstance(val, float):
            v = str(val)
        elif isinstance(val, (list, tuple)):
            # Point/list of numbers
            nums: list[str] = []
            for j, n in enumerate(val):
                if not isinstance(n, (int, float)):
                    raise ValueError(f"values[{i}].value[{j}] must be number")
                nums.append(str(float(n)))
            v = "(" + " ".join(nums) + ")"
        elif val is None:
            # No 'null' in LISP, store as empty string marker
            v = "nil"
        else:
            raise ValueError(f"values[{i}].value has unsupported type")

        parts.append(f"(cons {code} {v})")

    return "(list " + " ".join(parts) + ")"


def _strip_ok(obj: Dict[str, Any]) -> Dict[str, Any]:
    if "ok" not in obj:
        return obj
    out = dict(obj)
    out.pop("ok", None)
    return out


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
    expr = _MCP_DICT_LISP_LIB + "\n(mcp-dict-list)\n"
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def dict_keys(ctx: Context, dict_name: str) -> Dict[str, Any]:
    """List keys (and entry types) in a named dictionary."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    expr = _MCP_DICT_LISP_LIB + f"\n(mcp-dict-keys {_lisp_string(dict_name)})\n"
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
    expr = _MCP_DICT_LISP_LIB + f"\n(mcp-xrecord-get {_lisp_string(dict_name)} {_lisp_string(key)})\n"
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
    expr = _MCP_DICT_LISP_LIB + f"\n(mcp-xrecord-set {_lisp_string(dict_name)} {_lisp_string(key)} {values_expr} {ow})\n"
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
    expr = _MCP_DICT_LISP_LIB + f"\n(mcp-xrecord-delete {_lisp_string(dict_name)} {_lisp_string(key)})\n"
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


@mcp.tool()
def dict_delete(ctx: Context, dict_name: str, recursive: bool = True) -> Dict[str, Any]:
    """Delete a named dictionary from the Named Objects Dictionary."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    rec = "T" if recursive else "nil"
    expr = _MCP_DICT_LISP_LIB + f"\n(mcp-dict-delete {_lisp_string(dict_name)} {rec})\n"
    obj = _run_lisp_json(ctx, expr)
    return _strip_ok(obj)


def main() -> None:
    # Run MCP over stdio (FastMCP default)
    mcp.run()


if __name__ == "__main__":
    main()

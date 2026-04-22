import os
import shutil
import hashlib

import uuid
from dataclasses import dataclass
from typing import Any, Dict, Optional

from mcp.server.fastmcp import FastMCP, Context

from .autocad_bridge import AutoCADBridge
from .bridge_plugin_client import EventBridgeClient
from .command_waiter import (
    CommandWaiter,
    LISP_COMPLETION_EVENTS,
    LISP_START_EVENTS,
)
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
MIN_TIMEOUT_SEC = 0.1
MAX_TIMEOUT_SEC = 1800.0
MIN_POLL_INTERVAL_SEC = 0.01
MAX_POLL_INTERVAL_SEC = 5.0
EVENT_BRIDGE_ENABLED_ENV = "AUTOCAD_MCP_EVENT_BRIDGE_ENABLED"
EVENT_BRIDGE_HEARTBEAT_TIMEOUT_ENV = "AUTOCAD_MCP_EVENT_BRIDGE_HEARTBEAT_TIMEOUT_SEC"
EVENT_BRIDGE_MAX_DROPPED_FOR_WAIT_ENV = "AUTOCAD_MCP_EVENT_BRIDGE_MAX_DROPPED_FOR_WAIT"
EVENT_BRIDGE_OBJECT_EVENTS_ENABLED_ENV = "AUTOCAD_MCP_EVENT_BRIDGE_OBJECT_EVENTS_ENABLED"
EVENT_BRIDGE_AUTOLOAD_ENV = "AUTOCAD_MCP_EVENT_BRIDGE_AUTOLOAD"
EVENT_BRIDGE_AUTOLOAD_DLL_ENV = "AUTOCAD_MCP_EVENT_BRIDGE_AUTOLOAD_DLL"
DEFAULT_EVENT_BRIDGE_HEARTBEAT_TIMEOUT_SEC = 6.0
DEFAULT_EVENT_BRIDGE_MAX_DROPPED_FOR_WAIT = 0


@dataclass
class AppState:
    session_id: str
    bridge: AutoCADBridge
    streams: OutputStreamManager
    audit: SessionLogger
    event_bridge_enabled: bool
    event_bridge_client: Optional[EventBridgeClient]
    event_bridge_pid: Optional[int]
    event_bridge_heartbeat_timeout_sec: float
    event_bridge_max_dropped_for_wait: int
    event_bridge_object_events_requested: bool
    event_bridge_autoload_enabled: bool
    event_bridge_autoload_dll: Optional[str]
    event_bridge_last_prepare_issue: Optional[str]


def _parse_bool_env(name: str, default: bool = False) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return bool(default)
    if raw in ("1", "true", "yes", "on"):
        return True
    if raw in ("0", "false", "no", "off"):
        return False
    return bool(default)


def _parse_float_env(name: str, default: float) -> float:
    raw = (os.environ.get(name) or "").strip()
    if not raw:
        return float(default)
    try:
        value = float(raw)
    except Exception:
        return float(default)
    if value <= 0:
        return float(default)
    return float(value)


def _parse_int_env(name: str, default: int) -> int:
    raw = (os.environ.get(name) or "").strip()
    if not raw:
        return int(default)
    try:
        value = int(raw)
    except Exception:
        return int(default)
    if value < 0:
        return int(default)
    return int(value)


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
        event_bridge_enabled=_parse_bool_env(EVENT_BRIDGE_ENABLED_ENV, default=False),
        event_bridge_client=None,
        event_bridge_pid=None,
        event_bridge_heartbeat_timeout_sec=_parse_float_env(
            EVENT_BRIDGE_HEARTBEAT_TIMEOUT_ENV,
            default=DEFAULT_EVENT_BRIDGE_HEARTBEAT_TIMEOUT_SEC,
        ),
        event_bridge_max_dropped_for_wait=_parse_int_env(
            EVENT_BRIDGE_MAX_DROPPED_FOR_WAIT_ENV,
            default=DEFAULT_EVENT_BRIDGE_MAX_DROPPED_FOR_WAIT,
        ),
        event_bridge_object_events_requested=_parse_bool_env(
            EVENT_BRIDGE_OBJECT_EVENTS_ENABLED_ENV,
            default=True,
        ),
        event_bridge_autoload_enabled=_parse_bool_env(
            EVENT_BRIDGE_AUTOLOAD_ENV,
            default=True,
        ),
        event_bridge_autoload_dll=(os.environ.get(EVENT_BRIDGE_AUTOLOAD_DLL_ENV) or "").strip() or None,
        event_bridge_last_prepare_issue=None,
    )


state = _make_state()
mcp = FastMCP("acad-cmd")


def _ensure_logfile_stream(ctx: Context) -> None:
    """Ensure default output stream is a logfile stream."""

    s = state.streams.get_default()
    if s and s.mode == "logfile" and s.logfile_path:
        return
    start_logging(ctx, mode="logfile")


def _normalize_timeout_sec(timeout_sec: Any) -> float:
    try:
        value = float(timeout_sec)
    except Exception as exc:
        raise ValueError(
            f"timeout_sec must be a number in range [{MIN_TIMEOUT_SEC}, {MAX_TIMEOUT_SEC}]"
        ) from exc
    if value < MIN_TIMEOUT_SEC or value > MAX_TIMEOUT_SEC:
        raise ValueError(f"timeout_sec must be in range [{MIN_TIMEOUT_SEC}, {MAX_TIMEOUT_SEC}]")
    return value


def _normalize_poll_interval_sec(poll_interval_sec: Any) -> float:
    try:
        value = float(poll_interval_sec)
    except Exception as exc:
        raise ValueError(
            f"poll_interval_sec must be a number in range [{MIN_POLL_INTERVAL_SEC}, {MAX_POLL_INTERVAL_SEC}]"
        ) from exc
    if value < MIN_POLL_INTERVAL_SEC or value > MAX_POLL_INTERVAL_SEC:
        raise ValueError(
            f"poll_interval_sec must be in range [{MIN_POLL_INTERVAL_SEC}, {MAX_POLL_INTERVAL_SEC}]"
        )
    return value


def _run_lisp_json(ctx: Context, expr: str, *, timeout_sec: float) -> Dict[str, Any]:
    """Run a LISP expr that prints one [MCP:JSON]{...} line."""

    timeout_sec = _normalize_timeout_sec(timeout_sec)
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


def _repo_root() -> str:
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def _resolve_bridge_dll_path() -> Optional[str]:
    explicit = (state.event_bridge_autoload_dll or "").strip()
    if explicit:
        path = os.path.abspath(os.path.expanduser(os.path.expandvars(explicit)))
        return path if os.path.isfile(path) else None

    base_dir = os.path.join(_repo_root(), "plugins", "AcadEventBridge", "bin", "Debug", "net48")
    if not os.path.isdir(base_dir):
        return None

    candidates: list[str] = []
    primary = os.path.join(base_dir, "AcadEventBridge.dll")

    hotbuild_dirs: list[tuple[float, str]] = []
    try:
        for name in os.listdir(base_dir):
            if not str(name).lower().startswith("hotbuild"):
                continue
            d = os.path.join(base_dir, name)
            if not os.path.isdir(d):
                continue
            dll = os.path.join(d, "AcadEventBridge.dll")
            if not os.path.isfile(dll):
                continue
            try:
                mtime = os.path.getmtime(dll)
            except Exception:
                mtime = 0.0
            hotbuild_dirs.append((mtime, dll))
    except Exception:
        pass

    for _mtime, dll in sorted(hotbuild_dirs, key=lambda x: x[0], reverse=True):
        candidates.append(dll)

    if os.path.isfile(primary):
        candidates.append(primary)

    return candidates[0] if candidates else None


def _is_ascii_path(path: str) -> bool:
    try:
        str(path).encode("ascii")
        return True
    except Exception:
        return False


def _candidate_ascii_cache_dirs() -> list[str]:
    out: list[str] = []
    env_cache = (os.environ.get("AUTOCAD_MCP_NETLOAD_CACHE_DIR") or "").strip()
    local_app_data = (os.environ.get("LOCALAPPDATA") or "").strip()
    temp_dir = (os.environ.get("TEMP") or "").strip()
    for raw in (
        env_cache,
        r"C:\Temp\acad_event_bridge_cache",
        (os.path.join(local_app_data, "acad_event_bridge_cache") if local_app_data else ""),
        (os.path.join(temp_dir, "acad_event_bridge_cache") if temp_dir else ""),
        r"C:\Windows\Temp\acad_event_bridge_cache",
    ):
        if not raw:
            continue
        path = os.path.abspath(os.path.expanduser(os.path.expandvars(raw)))
        if not _is_ascii_path(path):
            continue
        if path not in out:
            out.append(path)
    return out


def _file_sha256_hex(path: str) -> Optional[str]:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            while True:
                chunk = fh.read(1024 * 1024)
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def _prepare_bridge_dll_for_netload(dll_path: str) -> Optional[str]:
    src = os.path.abspath(os.path.expanduser(os.path.expandvars(dll_path)))
    if not os.path.isfile(src):
        return None
    if _is_ascii_path(src):
        return src

    try:
        mtime = int(os.path.getmtime(src))
    except Exception:
        mtime = 0
    try:
        size = int(os.path.getsize(src))
    except Exception:
        size = 0
    src_hash = _file_sha256_hex(src)
    hash_part = (src_hash[:12] if src_hash else "nohash")
    stamp = f"{hash_part}_{mtime:x}_{size:x}"
    filename = f"AcadEventBridge_{stamp}.dll"

    for cache_dir in _candidate_ascii_cache_dirs():
        try:
            os.makedirs(cache_dir, exist_ok=True)
            dst = os.path.join(cache_dir, filename)
            need_copy = True
            if os.path.isfile(dst):
                if src_hash:
                    need_copy = _file_sha256_hex(dst) != src_hash
                else:
                    need_copy = os.path.getsize(dst) != size
            if need_copy:
                shutil.copy2(src, dst)
            if os.path.isfile(dst):
                if src_hash and _file_sha256_hex(dst) != src_hash:
                    continue
                return dst
        except Exception:
            continue

    return None


def _build_netload_command(dll_path: str) -> str:
    normalized = str(dll_path).replace("\\", "/")
    return f'(command "_.NETLOAD" "{lisp_quote_string(normalized)}")\n(princ)'


def _try_netload_bridge_plugin(*, timeout_sec: float) -> bool:
    dll_path = _resolve_bridge_dll_path()
    if not dll_path:
        state.event_bridge_last_prepare_issue = "bridge_autoload_dll_not_found"
        return False
    load_path = _prepare_bridge_dll_for_netload(dll_path)
    if not load_path:
        state.event_bridge_last_prepare_issue = "bridge_autoload_ascii_path_unavailable"
        return False

    try:
        state.bridge.send_command(_build_netload_command(load_path))
        wr = state.bridge.wait_for_idle(timeout_sec=max(5.0, float(timeout_sec)), poll_interval_sec=0.2)
        if not bool(getattr(wr, "completed", False)):
            state.event_bridge_last_prepare_issue = "bridge_autoload_timeout"
            return False
        state.event_bridge_last_prepare_issue = None
        return True
    except Exception:
        state.event_bridge_last_prepare_issue = "bridge_autoload_failed"
        return False


def _try_set_bridge_object_events(*, enabled: bool, timeout_sec: float) -> bool:
    cmd = "AEB_OBJECT_EVENTS_ON" if enabled else "AEB_OBJECT_EVENTS_OFF"
    try:
        state.bridge.send_command(cmd + "\n")
        wr = state.bridge.wait_for_idle(timeout_sec=max(3.0, float(timeout_sec)), poll_interval_sec=0.1)
        return bool(getattr(wr, "completed", False))
    except Exception:
        return False


def _stop_event_bridge_client() -> None:
    client = state.event_bridge_client
    if client is not None:
        try:
            client.stop()
        except Exception:
            pass
    state.event_bridge_client = None
    state.event_bridge_pid = None


def _bridge_client_is_heartbeat_stale(client: EventBridgeClient) -> bool:
    try:
        snap = client.snapshot()
    except Exception:
        return False
    if not bool(snap.connected):
        return False
    if snap.last_heartbeat is None:
        return False
    age = snap.heartbeat_age_sec
    if age is None:
        return False
    return float(age) > float(state.event_bridge_heartbeat_timeout_sec)


def _ensure_event_bridge_client(pid: Optional[int]) -> Optional[EventBridgeClient]:
    if not state.event_bridge_enabled:
        _stop_event_bridge_client()
        return None
    try:
        pid_int = int(pid) if pid is not None else None
    except Exception:
        pid_int = None
    if pid_int is None or pid_int <= 0:
        _stop_event_bridge_client()
        return None

    current = state.event_bridge_client
    current_pid = state.event_bridge_pid
    if current is not None and current_pid == pid_int:
        current.start()
        if not current.is_connected():
            try:
                current.wait_for_hello(timeout_sec=0.6)
            except Exception:
                pass
        if current.is_connected() and not _bridge_client_is_heartbeat_stale(current):
            return current
        try:
            current.stop()
        except Exception:
            pass

    if current is not None:
        try:
            current.stop()
        except Exception:
            pass

    pipe_name = state.bridge.get_pipe_name(pid=pid_int) or EventBridgeClient.pipe_name_for_pid(pid_int)
    client = EventBridgeClient(pipe_name=pipe_name)
    client.start()
    state.event_bridge_client = client
    state.event_bridge_pid = pid_int
    return client


def _ensure_event_bridge_ready(
    *,
    pid: Optional[int],
    timeout_sec: float,
    allow_autoload: bool = True,
    ensure_object_events: bool = True,
) -> Optional[EventBridgeClient]:
    client = _ensure_event_bridge_client(pid)
    if client is not None and not client.is_connected():
        try:
            client.wait_for_hello(timeout_sec=0.7)
        except Exception:
            pass

    if (client is None or not client.is_connected()) and allow_autoload and state.event_bridge_autoload_enabled:
        if _try_netload_bridge_plugin(timeout_sec=min(40.0, max(8.0, timeout_sec))):
            client = _ensure_event_bridge_client(pid)
            if client is not None and not client.is_connected():
                try:
                    client.wait_for_hello(timeout_sec=1.2)
                except Exception:
                    pass

    if client is None or not client.is_connected():
        if state.event_bridge_last_prepare_issue is None:
            state.event_bridge_last_prepare_issue = "bridge_client_disconnected"
        return client

    if not ensure_object_events:
        return client

    desired = bool(state.event_bridge_object_events_requested)
    current = client.snapshot().object_events_enabled
    if current is None:
        # Older plugin, or bridge state unknown: best-effort command anyway.
        current = False
    if bool(current) == desired:
        return client

    if not _try_set_bridge_object_events(enabled=desired, timeout_sec=min(12.0, max(3.0, timeout_sec))):
        state.event_bridge_last_prepare_issue = "bridge_object_events_toggle_failed"
        return client

    # Reconnect client to refresh hello snapshot with updated object_events_enabled.
    try:
        client.stop()
        client.start()
        client.wait_for_hello(timeout_sec=1.0)
    except Exception:
        pass
    return client


def _fetch_bridge_service_status(
    bridge_client: Optional[EventBridgeClient],
    *,
    timeout_sec: float = 0.8,
) -> Optional[Dict[str, Any]]:
    if bridge_client is None:
        return None
    if not bridge_client.request_response_available():
        return None
    try:
        raw = bridge_client.request_status(timeout_sec=timeout_sec)
    except Exception:
        return None
    if not isinstance(raw, dict):
        return None

    payload = raw.get("payload")
    payload_dict = payload if isinstance(payload, dict) else None
    return {
        "id": raw.get("id"),
        "ok": bool(raw.get("ok", False)),
        "ts": raw.get("ts"),
        "error": raw.get("error"),
        "payload": payload_dict,
    }


def _prepare_bridge_wait_context(
    *,
    timeout_sec: float,
) -> tuple[Optional[EventBridgeClient], Optional[int]]:
    """Return bridge client + sequence baseline for command waiting."""

    state.event_bridge_last_prepare_issue = None
    if not state.event_bridge_enabled:
        state.event_bridge_last_prepare_issue = "bridge_disabled"
        return None, None

    try:
        pid = state.bridge.get_pid()
    except Exception:
        pid = None
    try:
        pid_int = int(pid) if pid is not None else None
    except Exception:
        pid_int = None
    if pid_int is None or pid_int <= 0:
        state.event_bridge_last_prepare_issue = "acad_pid_unavailable"
        return None, None

    bridge_client = _ensure_event_bridge_ready(
        pid=pid_int,
        timeout_sec=timeout_sec,
        allow_autoload=True,
        ensure_object_events=True,
    )
    if bridge_client is None:
        state.event_bridge_last_prepare_issue = "bridge_client_unavailable"
        return None, None
    if not bridge_client.is_connected():
        try:
            bridge_client.wait_for_hello(timeout_sec=0.8)
        except Exception:
            pass
        if not bridge_client.is_connected():
            state.event_bridge_last_prepare_issue = "bridge_client_disconnected"
            return None, None

    try:
        wait_budget = min(
            max(0.3, timeout_sec * 0.2),
            max(0.3, state.event_bridge_heartbeat_timeout_sec),
        )
        bridge_client.wait_for_heartbeat(timeout_sec=wait_budget)
    except Exception:
        state.event_bridge_last_prepare_issue = "heartbeat_wait_failed"
        return None, None
    if not bridge_client.heartbeat_is_fresh(state.event_bridge_heartbeat_timeout_sec):
        state.event_bridge_last_prepare_issue = "heartbeat_timeout"
        return None, None

    try:
        snap = bridge_client.snapshot()
        dropped_count = int(snap.dropped_count or 0)
        if dropped_count > int(state.event_bridge_max_dropped_for_wait):
            state.event_bridge_last_prepare_issue = f"dropped_count_exceeded:{dropped_count}"
            return None, None

        # Drain backlog before taking baseline sequence to avoid stale events.
        bridge_client.drain_messages()
        snap = bridge_client.snapshot()
        after_seq = snap.last_seq
        if after_seq is None:
            wait_budget = min(
                max(0.3, timeout_sec * 0.2),
                max(0.3, state.event_bridge_heartbeat_timeout_sec),
            )
            bridge_client.wait_for_heartbeat(timeout_sec=wait_budget)
            after_seq = bridge_client.snapshot().last_seq
        if after_seq is None:
            state.event_bridge_last_prepare_issue = "baseline_seq_unavailable"
            return None, None
        state.event_bridge_last_prepare_issue = None
        return bridge_client, int(after_seq)
    except Exception:
        state.event_bridge_last_prepare_issue = "bridge_snapshot_failed"
        return None, None


def _wait_for_completion_with_fallback(
    *,
    bridge_client: Optional[EventBridgeClient],
    wait_after_seq: Optional[int],
    timeout_sec: float,
    poll_interval_sec: float,
    start_events: Optional[set[str]] = None,
    completion_events: Optional[set[str]] = None,
):
    waiter = CommandWaiter(heartbeat_timeout_sec=state.event_bridge_heartbeat_timeout_sec)
    wait_result = waiter.wait_for_completion(
        bridge_client=bridge_client,
        after_seq=wait_after_seq,
        timeout_sec=timeout_sec,
        fallback_wait=lambda t: state.bridge.wait_for_idle(
            timeout_sec=t,
            poll_interval_sec=poll_interval_sec,
        ),
        start_events=start_events,
        completion_events=completion_events,
    )
    if bridge_client is not None and wait_result.source in (
        "fallback_bridge_disconnected",
        "fallback_after_event_timeout",
    ):
        try:
            bridge_client.stop()
            bridge_client.start()
        except Exception:
            pass
    return wait_result


def _apply_wait_result_to_command_result(
    *,
    result: Dict[str, Any],
    wait_result: Any,
) -> Dict[str, Any]:
    result["completed"] = bool(wait_result.completed)
    result["needs_input"] = bool(wait_result.needs_input)
    result["wait_source"] = wait_result.source
    result["wait_completion_event"] = wait_result.completion_event
    result["wait_completion_seq"] = wait_result.completion_seq
    result["wait_started_seen"] = wait_result.started_seen
    result["wait_fallback_used"] = wait_result.fallback_used
    return result


def _refresh_waited_command_output(result: Dict[str, Any]) -> Dict[str, Any]:
    result["last_prompt"] = state.bridge.get_last_prompt()
    stream = state.streams.get_default()
    if stream and stream.mode == "logfile" and stream.logfile_path:
        text, new_cursor, truncated = state.streams.read_new(stream.stream_id, stream.cursor, 65536)
        result["log"] = {
            "stream_id": stream.stream_id,
            "cursor": new_cursor,
            "text": text,
            "truncated": truncated,
        }
    return result


@mcp.tool()
def get_status(ctx: Context) -> Dict[str, Any]:
    snapshot = None
    if hasattr(state.bridge, "get_status_snapshot"):
        try:
            snapshot = state.bridge.get_status_snapshot()
        except Exception:
            snapshot = None

    if isinstance(snapshot, dict):
        connected = bool(snapshot.get("connected"))
        dwg = snapshot.get("dwg")
        acadver = snapshot.get("acadver")
        hwnd = snapshot.get("acad_hwnd")
        pid = snapshot.get("acad_pid")
        cmdactive = snapshot.get("cmdactive")
        busy = bool(snapshot.get("busy"))
        stale = bool(snapshot.get("stale"))
        source = str(snapshot.get("source") or "live")
        error_class = snapshot.get("error_class")
        error_message = snapshot.get("error_message")
        locked_major = snapshot.get("locked_major")
        bound_progid = snapshot.get("bound_progid")
    else:
        connected = state.bridge.ensure_connection()
        dwg = state.bridge.get_dwg_label() if connected else None
        acadver = None
        hwnd = None
        pid = None
        cmdactive = None
        busy = False
        stale = False
        source = "live" if connected else "none"
        error_class = None
        error_message = None
        locked_major = None
        bound_progid = None
        if connected:
            try:
                acadver = str(state.bridge.get_variable("ACADVER"))
            except Exception:
                acadver = None
            try:
                hwnd = state.bridge.get_hwnd()
            except Exception:
                hwnd = None
            try:
                pid = state.bridge.get_pid()
            except Exception:
                pid = None

    default_stream = state.streams.get_default()
    event_bridge_enabled = bool(getattr(state, "event_bridge_enabled", False))
    bridge_client = _ensure_event_bridge_ready(
        pid=pid if connected else None,
        timeout_sec=4.0,
        allow_autoload=True,
        ensure_object_events=True,
    )
    event_bridge: Dict[str, Any] = {
        "enabled": event_bridge_enabled,
        "available": False,
        "connected": False,
        "autoload_enabled": state.event_bridge_autoload_enabled,
        "autoload_dll": _resolve_bridge_dll_path(),
        "request_response_available": False,
        "heartbeat_timeout_sec": state.event_bridge_heartbeat_timeout_sec,
        "max_dropped_for_wait": state.event_bridge_max_dropped_for_wait,
        "object_events_requested": state.event_bridge_object_events_requested,
        "degraded_reason": state.event_bridge_last_prepare_issue,
    }
    if bridge_client is not None:
        try:
            bridge_client.wait_for_hello(timeout_sec=0.35)
        except Exception:
            pass
        snap = bridge_client.snapshot()
        service_status = _fetch_bridge_service_status(bridge_client, timeout_sec=0.8)
        event_bridge = {
            "enabled": event_bridge_enabled,
            "available": bool(snap.available),
            "connected": bool(snap.connected),
            "pipe_name": snap.pipe_name,
            "plugin_version": snap.plugin_version,
            "last_seq": snap.last_seq,
            "last_heartbeat": snap.last_heartbeat,
            "plugin_pid": snap.plugin_pid,
            "autoload_enabled": state.event_bridge_autoload_enabled,
            "autoload_dll": _resolve_bridge_dll_path(),
            "object_events_enabled": snap.object_events_enabled,
            "request_response_available": snap.request_response_available,
            "object_events_requested": state.event_bridge_object_events_requested,
            "busy": snap.busy,
            "heartbeat_age_sec": snap.heartbeat_age_sec,
            "queue_depth": snap.queue_depth,
            "dropped_count": snap.dropped_count,
            "last_error": snap.last_error,
            "service_status": service_status,
            "heartbeat_timeout_sec": state.event_bridge_heartbeat_timeout_sec,
            "max_dropped_for_wait": state.event_bridge_max_dropped_for_wait,
            "degraded_reason": state.event_bridge_last_prepare_issue,
        }
    elif event_bridge_enabled and connected:
        pipe_name = state.bridge.get_pipe_name(pid=pid) if pid else None
        if pipe_name:
            event_bridge["pipe_name"] = pipe_name

    return {
        "ts": iso_now(),
        "session_id": state.session_id,
        "connected": connected,
        "busy": busy,
        "stale": stale,
        "source": source,
        "error_class": error_class,
        "error_message": error_message,
        "dwg": dwg,
        "acadver": acadver,
        "acad_hwnd": hwnd,
        "acad_pid": pid,
        "cmdactive": cmdactive,
        "locked_major": locked_major,
        "bound_progid": bound_progid,
        "event_bridge": event_bridge,
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
def open_drawing(
    ctx: Context,
    path: str,
    timeout_sec: float,
    read_only: bool = False,
) -> Dict[str, Any]:
    _ensure_connected()
    timeout_sec = _normalize_timeout_sec(timeout_sec)
    dwg_before = state.bridge.get_dwg_label()

    opened = state.bridge.open_drawing(path=path, timeout_sec=timeout_sec, read_only=read_only)
    dwg_after = state.bridge.get_dwg_label()
    _ensure_logfile_stream(ctx)

    result = {
        "path": opened.get("path"),
        "dwg_before": dwg_before,
        "dwg": dwg_after,
        "dwg_opened": opened.get("dwg"),
        "already_open": bool(opened.get("already_open")),
        "opened": bool(opened.get("opened")),
        "activated": bool(opened.get("activated")),
        "read_only": bool(opened.get("read_only")),
        "timeout_sec": timeout_sec,
    }
    state.audit.log(
        "open_drawing",
        {
            "path": result["path"],
            "dwg_before": dwg_before,
            "dwg_after": dwg_after,
            "already_open": result["already_open"],
            "opened": result["opened"],
            "activated": result["activated"],
            "read_only": result["read_only"],
            "timeout_sec": timeout_sec,
        },
        dwg=dwg_after,
    )
    return result


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
def get_last_output(ctx: Context, source: str = "logfile") -> Dict[str, Any]:
    _ensure_connected()
    dwg = state.bridge.get_dwg_label()

    if source not in ("logfile", "lastprompt"):
        raise ValueError("source must be 'logfile' or 'lastprompt'")

    if source == "logfile":
        _ensure_logfile_stream(ctx)
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
    timeout_sec: float,
    wait: bool = True,
    poll_interval_sec: float = 0.1,
) -> Dict[str, Any]:
    _ensure_connected()
    _ensure_logfile_stream(ctx)
    timeout_sec = _normalize_timeout_sec(timeout_sec)
    poll_interval_sec = _normalize_poll_interval_sec(poll_interval_sec)
    dwg = state.bridge.get_dwg_label()

    bridge_client: Optional[EventBridgeClient] = None
    wait_after_seq: Optional[int] = None
    if wait:
        bridge_client, wait_after_seq = _prepare_bridge_wait_context(timeout_sec=timeout_sec)

    command_id = state.bridge.send_command(command)

    state.audit.log(
        "send_command",
        {
            "command_id": command_id,
            "command": command,
            "wait": wait,
            "timeout_sec": timeout_sec,
            "bridge_wait_enabled": bool(bridge_client is not None),
            "bridge_wait_after_seq": wait_after_seq,
            "bridge_wait_prepare_issue": state.event_bridge_last_prepare_issue,
        },
        dwg=dwg,
    )

    completed = True
    needs_input = False
    wait_source = "no_wait"
    wait_completion_event = None
    wait_completion_seq = None
    wait_started_seen = False
    wait_fallback_used = False

    if wait:
        wait_result = _wait_for_completion_with_fallback(
            bridge_client=bridge_client,
            wait_after_seq=wait_after_seq,
            timeout_sec=timeout_sec,
            poll_interval_sec=poll_interval_sec,
        )
        completed = wait_result.completed
        needs_input = wait_result.needs_input
        wait_source = wait_result.source
        wait_completion_event = wait_result.completion_event
        wait_completion_seq = wait_result.completion_seq
        wait_started_seen = wait_result.started_seen
        wait_fallback_used = wait_result.fallback_used

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
            "wait_source": wait_source,
            "wait_completion_event": wait_completion_event,
            "wait_completion_seq": wait_completion_seq,
            "wait_started_seen": wait_started_seen,
            "wait_fallback_used": wait_fallback_used,
            "bridge_wait_prepare_issue": state.event_bridge_last_prepare_issue,
        },
        dwg=dwg,
    )

    return {
        "command_id": command_id,
        "dwg": dwg,
        "sent": command,
        "completed": completed,
        "needs_input": needs_input,
        "wait_source": wait_source,
        "wait_completion_event": wait_completion_event,
        "wait_completion_seq": wait_completion_seq,
        "wait_started_seen": wait_started_seen,
        "wait_fallback_used": wait_fallback_used,
        "bridge_wait_prepare_issue": state.event_bridge_last_prepare_issue,
        "last_prompt": last_prompt,
        "log": log_block,
    }


@mcp.tool()
def load_lisp_file(
    ctx: Context,
    path: str,
    timeout_sec: float,
    wait: bool = True,
) -> Dict[str, Any]:
    _ensure_connected()
    timeout_sec = _normalize_timeout_sec(timeout_sec)
    dwg = state.bridge.get_dwg_label()
    cmd = build_load_lisp_command(path)
    state.audit.log("load_lisp_file", {"path": path, "command": cmd}, dwg=dwg)

    if not wait:
        return send_command(ctx, cmd, wait=False, timeout_sec=timeout_sec)

    bridge_client, wait_after_seq = _prepare_bridge_wait_context(timeout_sec=timeout_sec)
    result = send_command(ctx, cmd, wait=False, timeout_sec=timeout_sec)
    wait_result = _wait_for_completion_with_fallback(
        bridge_client=bridge_client,
        wait_after_seq=wait_after_seq,
        timeout_sec=timeout_sec,
        poll_interval_sec=0.1,
        start_events=LISP_START_EVENTS,
        completion_events=LISP_COMPLETION_EVENTS,
    )
    result = _apply_wait_result_to_command_result(result=result, wait_result=wait_result)
    result = _refresh_waited_command_output(result)
    state.audit.log(
        "load_lisp_file_wait_result",
        {
            "path": path,
            "completed": result.get("completed"),
            "needs_input": result.get("needs_input"),
            "wait_source": result.get("wait_source"),
            "wait_completion_event": result.get("wait_completion_event"),
            "wait_completion_seq": result.get("wait_completion_seq"),
            "wait_fallback_used": result.get("wait_fallback_used"),
        },
        dwg=dwg,
    )
    return result


@mcp.tool()
def run_lisp(
    ctx: Context,
    expr: str,
    timeout_sec: float,
    wait: bool = True,
) -> Dict[str, Any]:
    _ensure_connected()
    timeout_sec = _normalize_timeout_sec(timeout_sec)
    dwg = state.bridge.get_dwg_label()
    marker_id = str(uuid.uuid4())
    script = build_run_lisp_script(expr, marker_id)
    state.audit.log("run_lisp", {"expr": expr, "marker_id": marker_id}, dwg=dwg)

    if not wait:
        result = send_command(ctx, script, wait=False, timeout_sec=timeout_sec)
        result["marker_id"] = marker_id
        return result

    bridge_client, wait_after_seq = _prepare_bridge_wait_context(timeout_sec=timeout_sec)
    result = send_command(ctx, script, wait=False, timeout_sec=timeout_sec)
    wait_result = _wait_for_completion_with_fallback(
        bridge_client=bridge_client,
        wait_after_seq=wait_after_seq,
        timeout_sec=timeout_sec,
        poll_interval_sec=0.1,
        start_events=LISP_START_EVENTS,
        completion_events=LISP_COMPLETION_EVENTS,
    )
    result = _apply_wait_result_to_command_result(result=result, wait_result=wait_result)
    result = _refresh_waited_command_output(result)
    state.audit.log(
        "run_lisp_wait_result",
        {
            "marker_id": marker_id,
            "completed": result.get("completed"),
            "needs_input": result.get("needs_input"),
            "wait_source": result.get("wait_source"),
            "wait_completion_event": result.get("wait_completion_event"),
            "wait_completion_seq": result.get("wait_completion_seq"),
            "wait_fallback_used": result.get("wait_fallback_used"),
        },
        dwg=dwg,
    )
    result["marker_id"] = marker_id
    return result


@mcp.tool()
def dict_list(ctx: Context, timeout_sec: float) -> Dict[str, Any]:
    """List top-level dictionaries from Named Objects Dictionary."""

    _ensure_connected()
    expr = _lisp_concat(_MCP_DICT_LISP_LIB, "(mcp-dict-list)\n")
    obj = _run_lisp_json(ctx, expr, timeout_sec=timeout_sec)
    return _strip_ok(obj)


@mcp.tool()
def dict_keys(ctx: Context, dict_name: str, timeout_sec: float) -> Dict[str, Any]:
    """List keys (and entry types) in a named dictionary."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    expr = _lisp_concat(_MCP_DICT_LISP_LIB, f"(mcp-dict-keys {_lisp_string(dict_name)})\n")
    obj = _run_lisp_json(ctx, expr, timeout_sec=timeout_sec)
    return _strip_ok(obj)


@mcp.tool()
def dict_xrecord_get(ctx: Context, dict_name: str, key: str, timeout_sec: float) -> Dict[str, Any]:
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
    obj = _run_lisp_json(ctx, expr, timeout_sec=timeout_sec)
    return _strip_ok(obj)


@mcp.tool()
def dict_xrecord_set(
    ctx: Context,
    dict_name: str,
    key: str,
    values: Any,
    timeout_sec: float,
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
    obj = _run_lisp_json(ctx, expr, timeout_sec=timeout_sec)
    return _strip_ok(obj)


@mcp.tool()
def dict_xrecord_delete(ctx: Context, dict_name: str, key: str, timeout_sec: float) -> Dict[str, Any]:
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
    obj = _run_lisp_json(ctx, expr, timeout_sec=timeout_sec)
    return _strip_ok(obj)


@mcp.tool()
def dict_delete(ctx: Context, dict_name: str, timeout_sec: float, recursive: bool = True) -> Dict[str, Any]:
    """Delete a named dictionary from the Named Objects Dictionary."""

    _ensure_connected()
    if not dict_name:
        raise ValueError("dict_name must be non-empty")
    rec = "T" if recursive else "nil"
    expr = _lisp_concat(_MCP_DICT_LISP_LIB, f"(mcp-dict-delete {_lisp_string(dict_name)} {rec})\n")
    obj = _run_lisp_json(ctx, expr, timeout_sec=timeout_sec)
    return _strip_ok(obj)


@mcp.tool()
def selection(
    ctx: Context,
    timeout_sec: float,
    prompt: Optional[str] = None,
    filter: Any = None,
    max_objects: Optional[int] = None,
    alert_message: Optional[str] = None,
) -> Dict[str, Any]:
    """Get currently selected objects, or prompt the user to select objects.

    Returns only handle + type for each selected object.
    """

    _ensure_connected()
    _ensure_logfile_stream(ctx)
    timeout_sec = _normalize_timeout_sec(timeout_sec)
    dwg = state.bridge.get_dwg_label()

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
    r1 = send_command(ctx, expr1, wait=True, timeout_sec=timeout_sec)
    log_block1 = r1.get("log") or {}
    initial_text1 = str(log_block1.get("text") or "")
    cursor1 = log_block1.get("cursor")
    out1 = _collect_selection_stream_lite(
        streams=state.streams,
        stream_id=stream.stream_id,
        req_id=req_id1,
        timeout_sec=timeout_sec,
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
        timeout_sec=timeout_sec,
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


def main() -> None:
    # Run MCP over stdio (FastMCP default)
    mcp.run()


if __name__ == "__main__":
    main()

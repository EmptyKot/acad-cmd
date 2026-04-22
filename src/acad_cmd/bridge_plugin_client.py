from __future__ import annotations

import json
import threading
import time
import uuid
from collections import deque
from dataclasses import dataclass
from typing import Any, Deque, Dict, List, Optional

import pywintypes
import win32con
import win32file
import win32pipe

from .event_state import EventState, EventStateSnapshot
from .protocol import consume_complete_lines


ERROR_BROKEN_PIPE = 109
ERROR_FILE_NOT_FOUND = 2
ERROR_INVALID_HANDLE = 6
ERROR_NO_DATA = 232
ERROR_SEM_TIMEOUT = 121


@dataclass(frozen=True)
class EventBridgeSnapshot:
    available: bool
    connected: bool
    pipe_name: str
    protocol: Optional[int]
    plugin: Optional[str]
    plugin_version: Optional[str]
    plugin_pid: Optional[int]
    object_events_enabled: Optional[bool]
    request_response_available: bool
    last_heartbeat: Optional[str]
    last_seq: Optional[int]
    busy: Optional[bool]
    command_depth: Optional[int]
    lisp_depth: Optional[int]
    active_doc_id: Optional[str]
    queue_depth: Optional[int]
    dropped_count: Optional[int]
    last_command_event: Optional[str]
    last_command_name: Optional[str]
    last_command_doc_id: Optional[str]
    last_command_doc_name: Optional[str]
    last_command_ts: Optional[str]
    last_error: Optional[str]
    heartbeat_age_sec: Optional[float]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "available": self.available,
            "connected": self.connected,
            "pipe_name": self.pipe_name,
            "protocol": self.protocol,
            "plugin": self.plugin,
            "plugin_version": self.plugin_version,
            "plugin_pid": self.plugin_pid,
            "object_events_enabled": self.object_events_enabled,
            "request_response_available": self.request_response_available,
            "last_heartbeat": self.last_heartbeat,
            "last_seq": self.last_seq,
            "busy": self.busy,
            "command_depth": self.command_depth,
            "lisp_depth": self.lisp_depth,
            "active_doc_id": self.active_doc_id,
            "queue_depth": self.queue_depth,
            "dropped_count": self.dropped_count,
            "last_command_event": self.last_command_event,
            "last_command_name": self.last_command_name,
            "last_command_doc_id": self.last_command_doc_id,
            "last_command_doc_name": self.last_command_doc_name,
            "last_command_ts": self.last_command_ts,
            "last_error": self.last_error,
            "heartbeat_age_sec": self.heartbeat_age_sec,
        }


class EventBridgeClient:
    """Named-pipe client for AcadEventBridge NDJSON stream (+ request/response)."""

    def __init__(
        self,
        pipe_name: str,
        *,
        connect_timeout_sec: float = 0.5,
        reconnect_delay_sec: float = 0.25,
        read_chunk_size: int = 4096,
        max_buffered_messages: int = 4096,
    ) -> None:
        if not pipe_name:
            raise ValueError("pipe_name must be non-empty")
        if connect_timeout_sec <= 0:
            raise ValueError("connect_timeout_sec must be > 0")
        if reconnect_delay_sec < 0:
            raise ValueError("reconnect_delay_sec must be >= 0")
        if read_chunk_size < 256:
            raise ValueError("read_chunk_size must be >= 256")
        if max_buffered_messages < 1:
            raise ValueError("max_buffered_messages must be >= 1")

        self.pipe_name = pipe_name
        self._pipe_path = rf"\\.\pipe\{pipe_name}"
        self._connect_timeout_sec = float(connect_timeout_sec)
        self._reconnect_delay_sec = float(reconnect_delay_sec)
        self._read_chunk_size = int(read_chunk_size)
        self._max_buffered_messages = int(max_buffered_messages)

        self._lock = threading.RLock()
        self._messages_cv = threading.Condition(self._lock)
        self._write_lock = threading.RLock()
        self._messages: Deque[Dict[str, Any]] = deque(maxlen=self._max_buffered_messages)
        self._responses: Dict[str, Dict[str, Any]] = {}
        self._reader_thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._read_handle = None

        self._state = EventState()
        self._available = False
        self._connected = False
        self._protocol: Optional[int] = None
        self._plugin: Optional[str] = None
        self._plugin_version: Optional[str] = None
        self._plugin_pid: Optional[int] = None
        self._object_events_enabled: Optional[bool] = None
        self._request_response_available: bool = False
        self._last_error: Optional[str] = None

    @staticmethod
    def pipe_name_for_pid(pid: int) -> str:
        if int(pid) <= 0:
            raise ValueError("pid must be > 0")
        return f"acad-event-bridge-{int(pid)}"

    def start(self) -> None:
        with self._lock:
            if self._reader_thread and self._reader_thread.is_alive():
                return
            self._stop_event.clear()
            self._reader_thread = threading.Thread(
                target=self._reader_loop,
                name="AcadEventBridgeClient.Reader",
                daemon=True,
            )
            self._reader_thread.start()

    def stop(self, *, join_timeout_sec: float = 2.0) -> None:
        self._stop_event.set()
        self._close_read_handle()
        t = None
        with self._lock:
            t = self._reader_thread
        if t is not None and t.is_alive():
            t.join(timeout=max(0.0, float(join_timeout_sec)))
        with self._lock:
            self._reader_thread = None
            self._connected = False
            self._request_response_available = False
            self._responses.clear()
            self._messages_cv.notify_all()

    def is_running(self) -> bool:
        with self._lock:
            return bool(self._reader_thread and self._reader_thread.is_alive())

    def is_connected(self) -> bool:
        with self._lock:
            return bool(self._connected)

    def heartbeat_is_fresh(self, timeout_sec: float) -> bool:
        with self._lock:
            if not self._connected:
                return False
        return self._state.heartbeat_is_fresh(timeout_sec)

    def wait_for_hello(self, timeout_sec: float) -> bool:
        deadline = time.monotonic() + max(0.0, float(timeout_sec))
        with self._messages_cv:
            while True:
                if self._plugin_version is not None:
                    return True
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                self._messages_cv.wait(timeout=remaining)

    def wait_for_heartbeat(self, timeout_sec: float) -> bool:
        deadline = time.monotonic() + max(0.0, float(timeout_sec))
        with self._messages_cv:
            while True:
                if self._state.has_heartbeat():
                    return True
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                self._messages_cv.wait(timeout=remaining)

    def drain_messages(self) -> List[Dict[str, Any]]:
        with self._messages_cv:
            out = list(self._messages)
            self._messages.clear()
            return out

    def request(
        self,
        method: str,
        payload: Optional[Dict[str, Any]] = None,
        *,
        timeout_sec: float = 2.0,
    ) -> Dict[str, Any]:
        req_method = str(method or "").strip()
        if not req_method:
            raise ValueError("method must be non-empty")
        if timeout_sec <= 0:
            raise ValueError("timeout_sec must be > 0")
        if not self.request_response_available():
            raise RuntimeError("request/response is not available for this bridge connection")

        request_id = uuid.uuid4().hex
        req_payload = payload if isinstance(payload, dict) else {}
        req_obj = {
            "type": "request",
            "id": request_id,
            "method": req_method,
            "payload": req_payload,
        }
        req_line = (json.dumps(req_obj, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")

        return self._request_over_active_connection(
            request_id=request_id,
            req_method=req_method,
            req_line=req_line,
            timeout_sec=float(timeout_sec),
        )

    def _request_over_active_connection(
        self,
        *,
        request_id: str,
        req_method: str,
        req_line: bytes,
        timeout_sec: float,
    ) -> Dict[str, Any]:
        with self._messages_cv:
            handle = self._read_handle
            connected = self._connected
            self._responses.pop(request_id, None)
        if not connected or handle is None:
            raise RuntimeError("event bridge is not connected")

        try:
            with self._write_lock:
                hr, written = win32file.WriteFile(handle, req_line, None)
                if int(hr or 0) != 0:
                    raise RuntimeError(f"pipe request write returned hr={hr}")
                try:
                    written_count = int(written)
                except Exception:
                    written_count = len(req_line)
                if written_count < len(req_line):
                    raise RuntimeError(
                        f"pipe request write truncated: {written_count}/{len(req_line)} bytes"
                    )
                try:
                    win32file.FlushFileBuffers(handle)
                except Exception:
                    # Ignore flush failures; some pipe servers don't support explicit flush.
                    pass
        except Exception as err:
            raise RuntimeError(f"pipe request write failed: {err}") from err

        deadline = time.monotonic() + float(timeout_sec)
        with self._messages_cv:
            while True:
                response = self._responses.pop(request_id, None)
                if response is not None:
                    return response
                if not self._connected:
                    raise RuntimeError("event bridge disconnected while waiting for response")
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise TimeoutError(f"event bridge request timeout: method={req_method}")
                self._messages_cv.wait(timeout=remaining)

    def _request_via_reconnect(
        self,
        *,
        request_id: str,
        req_method: str,
        req_line: bytes,
        timeout_sec: float,
    ) -> Dict[str, Any]:
        was_running = self.is_running()
        with self._write_lock:
            if was_running:
                self.stop(join_timeout_sec=min(1.0, max(0.1, timeout_sec / 4.0)))
            try:
                return self._request_once(request_id=request_id, req_method=req_method, req_line=req_line, timeout_sec=timeout_sec)
            finally:
                if was_running:
                    self.start()

    def _request_once(
        self,
        *,
        request_id: str,
        req_method: str,
        req_line: bytes,
        timeout_sec: float,
    ) -> Dict[str, Any]:
        deadline = time.monotonic() + float(timeout_sec)
        handle = None
        last_err: Optional[Exception] = None

        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            timeout_ms = max(50, int(min(remaining, 0.75) * 1000))

            try:
                win32pipe.WaitNamedPipe(self._pipe_path, timeout_ms)
                handle = win32file.CreateFile(
                    self._pipe_path,
                    win32con.GENERIC_READ | win32con.GENERIC_WRITE,
                    0,
                    None,
                    win32con.OPEN_EXISTING,
                    0,
                    None,
                )
                break
            except pywintypes.error as err:
                last_err = err
                winerror = int(getattr(err, "winerror", 0) or 0)
                if winerror in (ERROR_FILE_NOT_FOUND, ERROR_SEM_TIMEOUT):
                    time.sleep(min(0.05, max(0.01, remaining)))
                    continue
                raise RuntimeError(f"pipe request connect failed: {err}") from err

        if handle is None:
            if last_err is not None:
                raise RuntimeError(f"pipe request connect failed: {last_err}") from last_err
            raise TimeoutError(f"event bridge request timeout: method={req_method}")

        buf = ""
        try:
            hr, written = win32file.WriteFile(handle, req_line, None)
            if int(hr or 0) != 0:
                raise RuntimeError(f"pipe request write returned hr={hr}")
            try:
                written_count = int(written)
            except Exception:
                written_count = len(req_line)
            if written_count < len(req_line):
                raise RuntimeError(
                    f"pipe request write truncated: {written_count}/{len(req_line)} bytes"
                )
            try:
                win32file.FlushFileBuffers(handle)
            except Exception:
                pass

            while True:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise TimeoutError(f"event bridge request timeout: method={req_method}")

                try:
                    _peek_data, available_bytes, _peek_left = win32pipe.PeekNamedPipe(handle, 0)
                    available = int(available_bytes)
                except pywintypes.error as err:
                    winerror = int(getattr(err, "winerror", 0) or 0)
                    if winerror in (ERROR_BROKEN_PIPE, ERROR_INVALID_HANDLE, ERROR_NO_DATA):
                        raise RuntimeError("event bridge disconnected while waiting for response") from err
                    raise RuntimeError(f"pipe request peek failed (winerror={winerror}): {err}") from err

                if available <= 0:
                    time.sleep(min(0.05, max(0.01, remaining)))
                    continue

                _hr, data = win32file.ReadFile(handle, min(self._read_chunk_size, available), None)
                if not data:
                    time.sleep(0.01)
                    continue

                buf += bytes(data).decode("utf-8", errors="replace")
                lines, buf = consume_complete_lines(buf)
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        obj = json.loads(line)
                    except Exception:
                        continue
                    if not isinstance(obj, dict):
                        continue
                    msg_type = str(obj.get("type") or "").strip().lower()
                    if msg_type != "response":
                        continue
                    if str(obj.get("id") or "").strip() == request_id:
                        return obj
        finally:
            try:
                win32file.CloseHandle(handle)
            except Exception:
                pass

    def request_ping(self, *, timeout_sec: float = 1.5) -> Dict[str, Any]:
        return self.request("ping", {}, timeout_sec=timeout_sec)

    def request_status(self, *, timeout_sec: float = 1.5) -> Dict[str, Any]:
        return self.request("status", {}, timeout_sec=timeout_sec)

    def request_response_available(self) -> bool:
        with self._lock:
            return bool(self._connected) and bool(self._request_response_available)

    def event_state_snapshot(self) -> EventStateSnapshot:
        return self._state.snapshot()

    def snapshot(self) -> EventBridgeSnapshot:
        state_snapshot = self._state.snapshot()
        with self._lock:
            return EventBridgeSnapshot(
                available=self._available,
                connected=self._connected,
                pipe_name=self.pipe_name,
                protocol=self._protocol,
                plugin=self._plugin,
                plugin_version=self._plugin_version,
                plugin_pid=self._plugin_pid,
                object_events_enabled=self._object_events_enabled,
                request_response_available=self._request_response_available,
                last_heartbeat=state_snapshot.last_heartbeat,
                last_seq=state_snapshot.last_seq,
                busy=state_snapshot.busy,
                command_depth=state_snapshot.command_depth,
                lisp_depth=state_snapshot.lisp_depth,
                active_doc_id=state_snapshot.active_doc_id,
                queue_depth=state_snapshot.queue_depth,
                dropped_count=state_snapshot.dropped_count,
                last_command_event=state_snapshot.last_command_event,
                last_command_name=state_snapshot.last_command_name,
                last_command_doc_id=state_snapshot.last_command_doc_id,
                last_command_doc_name=state_snapshot.last_command_doc_name,
                last_command_ts=state_snapshot.last_command_ts,
                last_error=self._last_error,
                heartbeat_age_sec=state_snapshot.heartbeat_age_sec,
            )

    def _reader_loop(self) -> None:
        while not self._stop_event.is_set():
            handle = self._connect_pipe_once()
            if handle is None:
                if self._stop_event.wait(self._reconnect_delay_sec):
                    break
                continue

            try:
                self._run_read_loop(handle)
            finally:
                self._mark_disconnected(available=False)
                self._close_read_handle()

            if self._stop_event.wait(self._reconnect_delay_sec):
                break

    def _connect_pipe_once(self):
        timeout_ms = max(50, int(self._connect_timeout_sec * 1000))
        try:
            win32pipe.WaitNamedPipe(self._pipe_path, timeout_ms)
        except pywintypes.error as err:
            self._set_connect_error(err)
            return None

        request_response_available = True
        try:
            handle = win32file.CreateFile(
                self._pipe_path,
                win32con.GENERIC_READ | win32con.GENERIC_WRITE,
                0,
                None,
                win32con.OPEN_EXISTING,
                0,
                None,
            )
        except pywintypes.error as err:
            winerror = int(getattr(err, "winerror", 0) or 0)
            if winerror != 5:
                self._set_connect_error(err)
                return None
            try:
                handle = win32file.CreateFile(
                    self._pipe_path,
                    win32con.GENERIC_READ,
                    0,
                    None,
                    win32con.OPEN_EXISTING,
                    0,
                    None,
                )
                request_response_available = False
            except pywintypes.error as err2:
                self._set_connect_error(err2)
                return None

        with self._messages_cv:
            self._available = True
            self._connected = True
            self._request_response_available = bool(request_response_available)
            self._last_error = None
            self._read_handle = handle
            self._messages_cv.notify_all()
        return handle

    def _run_read_loop(self, handle) -> None:
        buf = ""
        while not self._stop_event.is_set():
            try:
                _peek_data, available_bytes, _peek_left = win32pipe.PeekNamedPipe(handle, 0)
                available = int(available_bytes)
            except pywintypes.error as err:
                winerror = int(getattr(err, "winerror", 0) or 0)
                if winerror in (ERROR_BROKEN_PIPE, ERROR_INVALID_HANDLE, ERROR_NO_DATA):
                    return
                self._set_runtime_error(f"pipe peek failed (winerror={winerror}): {err}")
                return
            except Exception as err:
                self._set_runtime_error(f"pipe peek failed: {err}")
                return

            if available <= 0:
                time.sleep(0.02)
                continue

            try:
                _hr, data = win32file.ReadFile(handle, min(self._read_chunk_size, available), None)
            except pywintypes.error as err:
                winerror = int(getattr(err, "winerror", 0) or 0)
                if winerror in (ERROR_BROKEN_PIPE, ERROR_INVALID_HANDLE, ERROR_NO_DATA):
                    return
                self._set_runtime_error(f"pipe read failed (winerror={winerror}): {err}")
                return
            except Exception as err:
                self._set_runtime_error(f"pipe read failed: {err}")
                return

            if not data:
                continue

            chunk = bytes(data).decode("utf-8", errors="replace")
            buf += chunk
            lines, buf = consume_complete_lines(buf)
            for line in lines:
                self._handle_line(line)

    def _handle_line(self, line: str) -> None:
        line = line.strip()
        if not line:
            return

        try:
            obj = json.loads(line)
        except Exception as err:
            self._set_runtime_error(f"ndjson parse failed: {err}")
            return
        if not isinstance(obj, dict):
            self._set_runtime_error("ndjson payload is not an object")
            return

        self._apply_message(obj)
        with self._messages_cv:
            self._messages.append(obj)
            self._messages_cv.notify_all()

    def _apply_message(self, msg: Dict[str, Any]) -> None:
        msg_type = str(msg.get("type") or "").strip().lower()
        with self._lock:
            if msg_type == "hello":
                try:
                    self._protocol = int(msg.get("protocol")) if msg.get("protocol") is not None else None
                except Exception:
                    self._protocol = None
                plugin_raw = msg.get("plugin")
                self._plugin = str(plugin_raw).strip() if plugin_raw is not None else None
                if self._plugin == "":
                    self._plugin = None
                version_raw = msg.get("version")
                self._plugin_version = str(version_raw).strip() if version_raw is not None else None
                if self._plugin_version == "":
                    self._plugin_version = None
                try:
                    self._plugin_pid = int(msg.get("pid")) if msg.get("pid") is not None else None
                except Exception:
                    self._plugin_pid = None
                object_events_raw = msg.get("object_events_enabled")
                if isinstance(object_events_raw, bool):
                    self._object_events_enabled = object_events_raw
                elif object_events_raw is None:
                    self._object_events_enabled = None
                else:
                    s = str(object_events_raw).strip().lower()
                    if s in ("1", "true", "yes", "on"):
                        self._object_events_enabled = True
                    elif s in ("0", "false", "no", "off"):
                        self._object_events_enabled = False
            elif msg_type == "response":
                resp_id = str(msg.get("id") or "").strip()
                if resp_id:
                    self._responses[resp_id] = msg
        self._state.apply_message(msg)

    def _set_connect_error(self, err: pywintypes.error) -> None:
        winerror = int(getattr(err, "winerror", 0) or 0)
        available = winerror != ERROR_FILE_NOT_FOUND
        self._mark_disconnected(available=available)
        self._set_runtime_error(f"pipe connect failed (winerror={winerror}): {err}")

    def _set_runtime_error(self, message: str) -> None:
        with self._messages_cv:
            self._last_error = str(message)
            self._messages_cv.notify_all()

    def _mark_disconnected(self, *, available: bool) -> None:
        with self._messages_cv:
            self._available = bool(available)
            self._connected = False
            self._request_response_available = False
            self._responses.clear()
            self._messages_cv.notify_all()

    def _close_read_handle(self) -> None:
        handle = None
        with self._lock:
            handle = self._read_handle
            self._read_handle = None
        if handle is not None:
            try:
                win32file.CloseHandle(handle)
            except Exception:
                pass

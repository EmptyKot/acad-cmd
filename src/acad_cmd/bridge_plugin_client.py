from __future__ import annotations

import json
import threading
import time
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


@dataclass(frozen=True)
class EventBridgeSnapshot:
    available: bool
    connected: bool
    pipe_name: str
    protocol: Optional[int]
    plugin: Optional[str]
    plugin_version: Optional[str]
    plugin_pid: Optional[int]
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
    """Named-pipe client for AcadEventBridge (read-only NDJSON stream)."""

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
        self._messages: Deque[Dict[str, Any]] = deque(maxlen=self._max_buffered_messages)
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
        except pywintypes.error as err:
            self._set_connect_error(err)
            return None

        with self._messages_cv:
            self._available = True
            self._connected = True
            self._last_error = None
            self._read_handle = handle
            self._messages_cv.notify_all()
        return handle

    def _run_read_loop(self, handle) -> None:
        buf = ""
        while not self._stop_event.is_set():
            try:
                _hr, data = win32file.ReadFile(handle, self._read_chunk_size, None)
            except pywintypes.error as err:
                winerror = int(getattr(err, "winerror", 0) or 0)
                if winerror in (ERROR_BROKEN_PIPE, ERROR_INVALID_HANDLE):
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

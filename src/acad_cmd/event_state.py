from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from typing import Any, Dict, Optional


def _coerce_int(value: Any) -> Optional[int]:
    try:
        if value is None:
            return None
        return int(value)
    except Exception:
        return None


def _coerce_bool(value: Any) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        v = value.strip().lower()
        if v in ("1", "true", "yes", "on"):
            return True
        if v in ("0", "false", "no", "off"):
            return False
    return None


def _coerce_str(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    return s or None


@dataclass(frozen=True)
class EventStateSnapshot:
    last_seq: Optional[int]
    busy: Optional[bool]
    command_depth: Optional[int]
    lisp_depth: Optional[int]
    active_doc_id: Optional[str]
    last_heartbeat: Optional[str]
    heartbeat_age_sec: Optional[float]
    queue_depth: Optional[int]
    dropped_count: Optional[int]
    last_command_event: Optional[str]
    last_command_name: Optional[str]
    last_command_doc_id: Optional[str]
    last_command_doc_name: Optional[str]
    last_command_ts: Optional[str]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "last_seq": self.last_seq,
            "busy": self.busy,
            "command_depth": self.command_depth,
            "lisp_depth": self.lisp_depth,
            "active_doc_id": self.active_doc_id,
            "last_heartbeat": self.last_heartbeat,
            "heartbeat_age_sec": self.heartbeat_age_sec,
            "queue_depth": self.queue_depth,
            "dropped_count": self.dropped_count,
            "last_command_event": self.last_command_event,
            "last_command_name": self.last_command_name,
            "last_command_doc_id": self.last_command_doc_id,
            "last_command_doc_name": self.last_command_doc_name,
            "last_command_ts": self.last_command_ts,
        }


class EventState:
    """Bridge message state snapshot used by Python-side business logic."""

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._last_seq: Optional[int] = None
        self._busy: Optional[bool] = None
        self._command_depth: Optional[int] = None
        self._lisp_depth: Optional[int] = None
        self._active_doc_id: Optional[str] = None
        self._last_heartbeat: Optional[str] = None
        self._last_heartbeat_mono: Optional[float] = None
        self._queue_depth: Optional[int] = None
        self._dropped_count: Optional[int] = None

        self._last_command_event: Optional[str] = None
        self._last_command_name: Optional[str] = None
        self._last_command_doc_id: Optional[str] = None
        self._last_command_doc_name: Optional[str] = None
        self._last_command_ts: Optional[str] = None

    def apply_message(self, msg: Dict[str, Any]) -> None:
        msg_type = str(msg.get("type") or "").strip().lower()
        with self._lock:
            seq = _coerce_int(msg.get("seq"))
            if seq is not None and (self._last_seq is None or seq > self._last_seq):
                self._last_seq = seq

            if msg_type == "heartbeat":
                self._last_heartbeat = _coerce_str(msg.get("ts"))
                self._last_heartbeat_mono = time.monotonic()
                self._busy = _coerce_bool(msg.get("busy"))
                self._command_depth = _coerce_int(msg.get("command_depth"))
                self._lisp_depth = _coerce_int(msg.get("lisp_depth"))
                self._active_doc_id = _coerce_str(msg.get("active_doc_id"))
                self._queue_depth = _coerce_int(msg.get("queue_depth"))
                self._dropped_count = _coerce_int(msg.get("dropped_count"))
                return

            if msg_type != "event":
                return

            event_name = str(msg.get("event") or "").strip().lower()
            if not event_name.startswith("command_"):
                return

            payload = msg.get("payload")
            payload_dict = payload if isinstance(payload, dict) else {}

            self._last_command_event = event_name
            self._last_command_name = _coerce_str(payload_dict.get("name"))
            self._last_command_doc_id = _coerce_str(msg.get("doc_id"))
            self._last_command_doc_name = _coerce_str(msg.get("doc_name"))
            self._last_command_ts = _coerce_str(msg.get("ts"))

            payload_busy = _coerce_bool(payload_dict.get("busy"))
            if payload_busy is not None:
                self._busy = payload_busy
            payload_command_depth = _coerce_int(payload_dict.get("command_depth"))
            if payload_command_depth is not None:
                self._command_depth = payload_command_depth
            payload_lisp_depth = _coerce_int(payload_dict.get("lisp_depth"))
            if payload_lisp_depth is not None:
                self._lisp_depth = payload_lisp_depth

    def has_heartbeat(self) -> bool:
        with self._lock:
            return self._last_heartbeat is not None

    def heartbeat_is_fresh(self, timeout_sec: float) -> bool:
        timeout_sec = float(timeout_sec)
        if timeout_sec <= 0:
            return False
        with self._lock:
            if self._last_heartbeat_mono is None:
                return False
            age = time.monotonic() - self._last_heartbeat_mono
            return age <= timeout_sec

    def snapshot(self) -> EventStateSnapshot:
        with self._lock:
            heartbeat_age_sec = None
            if self._last_heartbeat_mono is not None:
                heartbeat_age_sec = max(0.0, time.monotonic() - self._last_heartbeat_mono)

            return EventStateSnapshot(
                last_seq=self._last_seq,
                busy=self._busy,
                command_depth=self._command_depth,
                lisp_depth=self._lisp_depth,
                active_doc_id=self._active_doc_id,
                last_heartbeat=self._last_heartbeat,
                heartbeat_age_sec=heartbeat_age_sec,
                queue_depth=self._queue_depth,
                dropped_count=self._dropped_count,
                last_command_event=self._last_command_event,
                last_command_name=self._last_command_name,
                last_command_doc_id=self._last_command_doc_id,
                last_command_doc_name=self._last_command_doc_name,
                last_command_ts=self._last_command_ts,
            )

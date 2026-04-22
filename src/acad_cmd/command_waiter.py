from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any, Callable, Dict, Optional, Tuple

from .bridge_plugin_client import EventBridgeClient


COMMAND_START_EVENTS = {"command_will_start"}
COMMAND_COMPLETION_EVENTS = {
    "command_ended",
    "command_cancelled",
    "command_failed",
}
LISP_START_EVENTS = {"lisp_will_start"}
LISP_COMPLETION_EVENTS = {
    "lisp_ended",
    "lisp_cancelled",
}
START_EVENTS = COMMAND_START_EVENTS | LISP_START_EVENTS
COMPLETION_EVENTS = COMMAND_COMPLETION_EVENTS | LISP_COMPLETION_EVENTS


def _normalize_events(value: Optional[set[str]], fallback: set[str]) -> set[str]:
    if not value:
        return set(fallback)
    out: set[str] = set()
    for item in value:
        s = str(item).strip().lower()
        if s:
            out.add(s)
    return out or set(fallback)


def _coerce_int(value: Any) -> Optional[int]:
    try:
        if value is None:
            return None
        return int(value)
    except Exception:
        return None


@dataclass(frozen=True)
class CommandWaitResult:
    completed: bool
    needs_input: bool
    source: str
    completion_event: Optional[str]
    completion_seq: Optional[int]
    started_seen: bool
    fallback_used: bool
    bridge_connected: bool
    quiescent: Optional[bool]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "completed": self.completed,
            "needs_input": self.needs_input,
            "source": self.source,
            "completion_event": self.completion_event,
            "completion_seq": self.completion_seq,
            "started_seen": self.started_seen,
            "fallback_used": self.fallback_used,
            "bridge_connected": self.bridge_connected,
            "quiescent": self.quiescent,
        }


class CommandWaiter:
    """Wait for command completion with event-first strategy and COM fallback."""

    def __init__(
        self,
        *,
        heartbeat_timeout_sec: float = 6.0,
        poll_interval_sec: float = 0.05,
    ) -> None:
        if heartbeat_timeout_sec <= 0:
            raise ValueError("heartbeat_timeout_sec must be > 0")
        if poll_interval_sec <= 0:
            raise ValueError("poll_interval_sec must be > 0")
        self._heartbeat_timeout_sec = float(heartbeat_timeout_sec)
        self._poll_interval_sec = float(poll_interval_sec)

    def wait_for_completion(
        self,
        *,
        bridge_client: Optional[EventBridgeClient],
        after_seq: Optional[int],
        timeout_sec: float,
        fallback_wait: Callable[[float], Any],
        start_events: Optional[set[str]] = None,
        completion_events: Optional[set[str]] = None,
    ) -> CommandWaitResult:
        if timeout_sec <= 0:
            raise ValueError("timeout_sec must be > 0")

        deadline = time.monotonic() + float(timeout_sec)
        seq_threshold = _coerce_int(after_seq) if after_seq is not None else None
        wait_start_events = _normalize_events(start_events, START_EVENTS)
        wait_completion_events = _normalize_events(completion_events, COMPLETION_EVENTS)

        if bridge_client is None:
            return self._run_fallback(
                timeout_sec=timeout_sec,
                fallback_wait=fallback_wait,
                source="fallback_no_bridge",
            )

        started_seen = False

        while time.monotonic() < deadline:
            if not bridge_client.is_connected():
                return self._run_fallback(
                    timeout_sec=self._remaining_timeout(deadline),
                    fallback_wait=fallback_wait,
                    source="fallback_bridge_disconnected",
                )

            messages = bridge_client.drain_messages()
            completed, completion_event, completion_seq, started_seen = self._scan_messages(
                messages=messages,
                after_seq=seq_threshold,
                started_seen=started_seen,
                start_events=wait_start_events,
                completion_events=wait_completion_events,
            )
            if completed:
                return CommandWaitResult(
                    completed=True,
                    needs_input=False,
                    source="event_stream",
                    completion_event=completion_event,
                    completion_seq=completion_seq,
                    started_seen=started_seen,
                    fallback_used=False,
                    bridge_connected=True,
                    quiescent=None,
                )

            time.sleep(self._poll_interval_sec)

        # Timeout on event stream path.
        snap = bridge_client.snapshot()
        bridge_live = bool(snap.connected) and (
            snap.heartbeat_age_sec is None or snap.heartbeat_age_sec <= self._heartbeat_timeout_sec
        )
        if bridge_live and bool(snap.busy):
            return CommandWaitResult(
                completed=False,
                needs_input=True,
                source="event_timeout_busy",
                completion_event=None,
                completion_seq=None,
                started_seen=started_seen,
                fallback_used=False,
                bridge_connected=True,
                quiescent=None,
            )

        return self._run_fallback(
            timeout_sec=self._remaining_timeout(deadline),
            fallback_wait=fallback_wait,
            source="fallback_after_event_timeout",
            started_seen=started_seen,
        )

    def _scan_messages(
        self,
        *,
        messages: list[Dict[str, Any]],
        after_seq: Optional[int],
        started_seen: bool,
        start_events: set[str],
        completion_events: set[str],
    ) -> Tuple[bool, Optional[str], Optional[int], bool]:
        started = bool(started_seen)
        for msg in messages:
            if not isinstance(msg, dict):
                continue
            if str(msg.get("type") or "").strip().lower() != "event":
                continue
            seq = _coerce_int(msg.get("seq"))
            if after_seq is not None and seq is not None and seq <= after_seq:
                continue

            event_name = str(msg.get("event") or "").strip().lower()
            if event_name in start_events:
                started = True
            if event_name in completion_events:
                return True, event_name, seq, started
        return False, None, None, started

    def _run_fallback(
        self,
        *,
        timeout_sec: float,
        fallback_wait: Callable[[float], Any],
        source: str,
        started_seen: bool = False,
    ) -> CommandWaitResult:
        timeout_value = max(0.1, float(timeout_sec))
        wr = fallback_wait(timeout_value)
        completed = bool(getattr(wr, "completed", False))
        needs_input = bool(getattr(wr, "needs_input", False))
        quiescent = getattr(wr, "quiescent", None)
        try:
            quiescent_bool = bool(quiescent) if quiescent is not None else None
        except Exception:
            quiescent_bool = None
        return CommandWaitResult(
            completed=completed,
            needs_input=needs_input,
            source=source,
            completion_event=None,
            completion_seq=None,
            started_seen=bool(started_seen),
            fallback_used=True,
            bridge_connected=False,
            quiescent=quiescent_bool,
        )

    @staticmethod
    def _remaining_timeout(deadline_mono: float) -> float:
        return max(0.1, deadline_mono - time.monotonic())

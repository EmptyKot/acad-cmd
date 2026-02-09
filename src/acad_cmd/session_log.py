import json
import os
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, Optional


def iso_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


@dataclass
class SessionLogger:
    path: str
    session_id: str

    def __post_init__(self) -> None:
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        self._lock = threading.Lock()

    def log(self, event: str, payload: Dict[str, Any], *, dwg: Optional[str] = None) -> None:
        row = {
            "ts": iso_now(),
            "session_id": self.session_id,
            "event": event,
            "dwg": dwg,
            "payload": payload,
        }
        line = json.dumps(row, ensure_ascii=True)
        with self._lock:
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
                f.flush()
        # Best-effort: give AutoCAD/COM some breathing room
        time.sleep(0)

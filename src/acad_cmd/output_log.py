import locale
import os
from collections import deque
from dataclasses import dataclass
from typing import Deque, Dict, Optional, Tuple


def _preferred_text_encoding() -> str:
    enc = locale.getpreferredencoding(False) or "utf-8"
    return enc


@dataclass
class OutputStream:
    stream_id: str
    mode: str  # logfile|lastprompt
    logfile_path: Optional[str]
    cursor: int
    ring: Deque[str]
    started_by_server: bool = True


class OutputStreamManager:
    def __init__(self, base_dir: str, ring_max_chunks: int = 200) -> None:
        self.base_dir = base_dir
        self.ring_max_chunks = ring_max_chunks
        os.makedirs(self.base_dir, exist_ok=True)
        self._streams: Dict[str, OutputStream] = {}
        self._default_stream_id: Optional[str] = None

    def get_default(self) -> Optional[OutputStream]:
        if self._default_stream_id is None:
            return None
        return self._streams.get(self._default_stream_id)

    def get(self, stream_id: str) -> Optional[OutputStream]:
        return self._streams.get(stream_id)

    def start_logfile_stream(self, *, stream_id: str, logfile_path: str, cursor: int, started_by_server: bool) -> OutputStream:
        ring: Deque[str] = deque(maxlen=self.ring_max_chunks)
        s = OutputStream(
            stream_id=stream_id,
            mode="logfile",
            logfile_path=logfile_path,
            cursor=cursor,
            ring=ring,
            started_by_server=started_by_server,
        )
        self._streams[stream_id] = s
        self._default_stream_id = stream_id
        return s

    def start_lastprompt_stream(self, *, stream_id: str) -> OutputStream:
        ring: Deque[str] = deque(maxlen=self.ring_max_chunks)
        s = OutputStream(
            stream_id=stream_id,
            mode="lastprompt",
            logfile_path=None,
            cursor=0,
            ring=ring,
            started_by_server=False,
        )
        self._streams[stream_id] = s
        self._default_stream_id = stream_id
        return s

    def stop(self, stream_id: str) -> bool:
        if stream_id not in self._streams:
            return False
        del self._streams[stream_id]
        if self._default_stream_id == stream_id:
            self._default_stream_id = next(iter(self._streams), None)
        return True

    def read_new(self, stream_id: str, cursor: int, max_bytes: int) -> Tuple[str, int, bool]:
        s = self._streams.get(stream_id)
        if s is None or s.mode != "logfile" or not s.logfile_path:
            return "", cursor, False

        path = s.logfile_path
        if not os.path.exists(path):
            return "", cursor, False

        file_size = os.path.getsize(path)
        if cursor > file_size:
            cursor = file_size

        to_read = min(max_bytes, file_size - cursor)
        if to_read <= 0:
            return "", cursor, False

        with open(path, "rb") as f:
            f.seek(cursor)
            data = f.read(to_read)
        new_cursor = cursor + len(data)
        truncated = new_cursor < file_size and len(data) == max_bytes

        enc = _preferred_text_encoding()
        try:
            text = data.decode(enc, errors="replace")
        except Exception:
            text = data.decode("utf-8", errors="replace")

        if text:
            s.ring.append(text)
            s.cursor = new_cursor

        return text, new_cursor, truncated

    def read_tail(self, stream_id: str, tail_bytes: int = 8192) -> str:
        s = self._streams.get(stream_id)
        if s is None or s.mode != "logfile" or not s.logfile_path:
            return ""
        path = s.logfile_path
        if not os.path.exists(path):
            return ""
        size = os.path.getsize(path)
        start = max(0, size - tail_bytes)
        with open(path, "rb") as f:
            f.seek(start)
            data = f.read(size - start)
        enc = _preferred_text_encoding()
        try:
            return data.decode(enc, errors="replace")
        except Exception:
            return data.decode("utf-8", errors="replace")

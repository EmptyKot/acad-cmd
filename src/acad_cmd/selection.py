import time
from typing import Any, Dict, List, Optional

from .output_log import OutputStreamManager
from .protocol import consume_complete_lines, extract_mcp_json_messages


def collect_selection_stream_lite(
    *,
    streams: OutputStreamManager,
    stream_id: str,
    req_id: str,
    timeout_sec: float,
    poll_interval_sec: float = 0.2,
    max_bytes: int = 65536,
    initial_text: str = "",
    cursor: Optional[int] = None,
) -> Dict[str, Any]:
    """Collect streamed selection messages emitted as [MCP:JSON] lines.

    Lite variant: returns only handle + type for each object.
    """

    stream = streams.get(stream_id)
    if not stream or stream.mode != "logfile" or not stream.logfile_path:
        raise RuntimeError("No active logfile stream")

    cur = int(cursor if cursor is not None else stream.cursor)
    buf = ""

    started: Optional[Dict[str, Any]] = None
    items: Dict[int, Dict[str, Any]] = {}
    order: List[int] = []
    timed_out = False

    def _handle_msgs(msgs: List[Dict[str, Any]]) -> bool:
        nonlocal started
        for m in msgs:
            if not isinstance(m, dict):
                continue
            if m.get("req_id") != req_id:
                continue
            if m.get("ok") is False:
                raise RuntimeError(str(m.get("error") or "Unknown AutoLISP error"))
            ev = m.get("event")
            if ev == "start":
                started = m
            elif ev == "item_begin":
                i = int(m.get("i") or 0)
                items[i] = {
                    "handle": m.get("handle"),
                    "type": m.get("type"),
                }
                if i not in order:
                    order.append(i)
            elif ev == "done":
                return True
        return False

    if initial_text:
        msgs = extract_mcp_json_messages(initial_text)
        if _handle_msgs(msgs):
            started_local = started
            objs = [items[i] for i in sorted(order)]
            count = None
            errno = None
            if isinstance(started_local, dict):
                try:
                    count = int(started_local.get("count"))
                except Exception:
                    count = None
                try:
                    errno = int(started_local.get("errno"))
                except Exception:
                    errno = None
            return {
                "req_id": req_id,
                "count": count if count is not None else len(objs),
                "errno": errno,
                "objects": objs,
                "timed_out": False,
                "cursor": cur,
            }

    t0 = time.time()
    while True:
        if time.time() - t0 >= timeout_sec:
            timed_out = True
            break

        text, new_cursor, _tr = streams.read_new(stream_id, cur, max_bytes)
        cur = int(new_cursor)
        if text:
            buf += text
            lines, buf = consume_complete_lines(buf)
            if lines:
                msgs = extract_mcp_json_messages("\n".join(lines))
                if _handle_msgs(msgs):
                    break
        else:
            time.sleep(poll_interval_sec)

    objs = [items[i] for i in sorted(order)]

    count = None
    errno = None
    if isinstance(started, dict):
        try:
            count = int(started.get("count"))
        except Exception:
            count = None
        try:
            errno = int(started.get("errno"))
        except Exception:
            errno = None

    return {
        "req_id": req_id,
        "count": count if count is not None else len(objs),
        "errno": errno,
        "objects": objs,
        "timed_out": timed_out,
        "cursor": cur,
    }

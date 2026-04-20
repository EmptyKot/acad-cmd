from typing import Any, Dict, List, Tuple

MCP_JSON_MARKER = "[MCP:JSON]"


def extract_mcp_json(text: str) -> Dict[str, Any]:
    """Extract and parse the last MCP JSON marker from logfile output."""

    if not text:
        raise RuntimeError("No output text to parse")

    last_line = None
    for line in text.splitlines():
        if MCP_JSON_MARKER in line:
            last_line = line

    if last_line is None:
        raise RuntimeError("MCP JSON marker not found in output")

    idx = last_line.rfind(MCP_JSON_MARKER)
    payload = last_line[idx + len(MCP_JSON_MARKER) :].strip()
    if not payload:
        raise RuntimeError("MCP JSON marker present but payload is empty")

    import json

    try:
        obj = json.loads(payload)
    except Exception as e:
        raise RuntimeError(f"Failed to parse MCP JSON payload: {e}")

    if not isinstance(obj, dict):
        raise RuntimeError("MCP JSON payload is not an object")
    return obj


def extract_mcp_json_messages(text: str) -> List[Dict[str, Any]]:
    """Extract all MCP JSON marker payloads from text."""

    out: List[Dict[str, Any]] = []
    if not text:
        return out

    import json

    for line in text.splitlines():
        if MCP_JSON_MARKER not in line:
            continue
        idx = line.rfind(MCP_JSON_MARKER)
        payload = line[idx + len(MCP_JSON_MARKER) :].strip()
        if not payload:
            continue
        try:
            obj = json.loads(payload)
        except Exception:
            continue
        if isinstance(obj, dict):
            out.append(obj)
    return out


def consume_complete_lines(buf: str) -> Tuple[List[str], str]:
    """Split buffer into complete lines and a remainder (no trailing newline)."""

    if not buf:
        return [], ""
    last_nl = buf.rfind("\n")
    if last_nl < 0:
        return [], buf
    chunk = buf[: last_nl + 1]
    rest = buf[last_nl + 1 :]
    return chunk.splitlines(), rest

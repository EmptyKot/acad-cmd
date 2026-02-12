from __future__ import annotations

from pathlib import Path
import json
import sys

import anyio

from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


def _unwrap_result(r):
    payload = r.structuredContent if r.structuredContent is not None else None
    if payload is None:
        for item in r.content or []:
            txt = getattr(item, "text", None)
            if isinstance(txt, str):
                try:
                    payload = json.loads(txt)
                    break
                except Exception:
                    continue

    if isinstance(payload, dict) and isinstance(payload.get("result"), dict):
        payload = payload["result"]
    return payload


def _validate_selection_payload(payload: dict) -> tuple[bool, str]:
    if not isinstance(payload, dict):
        return False, "payload_not_object"
    objs = payload.get("objects")
    if objs is None:
        return False, "missing_objects"
    if not isinstance(objs, list):
        return False, "objects_not_list"
    for i, o in enumerate(objs[:10]):
        if not isinstance(o, dict):
            return False, f"object_{i}_not_object"
        if "handle" not in o or "type" not in o:
            return False, f"object_{i}_missing_handle_or_type"
        if o.get("handle") is not None and not isinstance(o.get("handle"), str):
            return False, f"object_{i}_handle_not_str"
        if o.get("type") is not None and not isinstance(o.get("type"), str):
            return False, f"object_{i}_type_not_str"
    return True, "ok"


async def main() -> None:
    root = Path(__file__).resolve().parents[1]
    py = root / ".venv" / "Scripts" / "python.exe"

    params = StdioServerParameters(
        command=str(py),
        args=["-m", "acad_cmd.server"],
        cwd=str(root),
        encoding="utf-8",
        encoding_error_handler="replace",
    )

    async with stdio_client(params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            with anyio.fail_after(30):
                await session.initialize()

            try:
                with anyio.fail_after(75):
                    r = await session.call_tool(
                        "selection",
                        {
                            "timeout_sec": 60,
                            "prompt": "Select 1 object and press Enter",
                            "max_objects": 1,
                        },
                    )
                if r.isError:
                    raise RuntimeError(f"selection error: {r.content}")

                payload = _unwrap_result(r)
                if not isinstance(payload, dict):
                    raise RuntimeError(f"selection unexpected result: {payload!r}")

                ok, reason = _validate_selection_payload(payload)
                out = {
                    "ok": True,
                    "path": "selection",
                    "count": payload.get("count"),
                    "timed_out": payload.get("timed_out"),
                    "payload_valid": bool(ok),
                    "payload_valid_reason": reason,
                    "objects_sample": (payload.get("objects") or [])[:3],
                }
                sys.stdout.write(json.dumps(out, ensure_ascii=True) + "\n")

            except TimeoutError:
                sys.stdout.write(json.dumps({"ok": False, "error": "selection_call_timeout"}, ensure_ascii=True) + "\n")


if __name__ == "__main__":
    anyio.run(main)

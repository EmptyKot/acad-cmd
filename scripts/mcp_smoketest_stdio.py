from __future__ import annotations

from pathlib import Path
import json

import anyio

from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


def _tail_text(path: Path, max_chars: int = 8000) -> str:
    try:
        data = path.read_bytes()
    except Exception:
        return ""
    try:
        txt = data.decode("utf-8", "replace")
    except Exception:
        txt = data.decode("latin1", "replace")
    return txt[-max_chars:]


async def main() -> None:
    root = Path(__file__).resolve().parents[1]
    py = root / ".venv" / "Scripts" / "python.exe"

    print(f"root={root}", flush=True)
    print(f"python={py}", flush=True)

    if not py.exists():
        raise SystemExit(f"Missing python exe: {py}")

    params = StdioServerParameters(
        command=str(py),
        args=["-m", "acad_cmd.server"],
        cwd=str(root),
        encoding="utf-8",
        encoding_error_handler="replace",
    )

    err_path = root / "logs" / "mcp_smoketest_stderr.log"
    err_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        err_path.write_text("", encoding="utf-8", errors="ignore")
    except Exception:
        pass

    try:
        with err_path.open("a", encoding="utf-8", errors="replace") as err_f:
            async with stdio_client(params, errlog=err_f) as (read_stream, write_stream):
                print(f"spawned stdio server (stderr -> {err_path})", flush=True)

                async with ClientSession(read_stream, write_stream) as session:
                    print("initialize...", flush=True)
                    with anyio.fail_after(30):
                        await session.initialize()

                    print("list_tools...", flush=True)
                    with anyio.fail_after(30):
                        tools = await session.list_tools()

                    tool_names = [t.name for t in tools.tools]
                    print("tools:", ", ".join(sorted(tool_names)), flush=True)

                    print("start_logging...", flush=True)
                    with anyio.fail_after(30):
                        r = await session.call_tool("start_logging", {"mode": "logfile", "reset": True})
                    if r.isError:
                        raise RuntimeError(f"start_logging error: {r.content}")

                    payload = None
                    if r.structuredContent is not None:
                        payload = r.structuredContent
                    else:
                        # FastMCP commonly returns a single text item with JSON.
                        for item in r.content or []:
                            txt = getattr(item, "text", None)
                            if not isinstance(txt, str):
                                continue
                            try:
                                payload = json.loads(txt)
                                break
                            except Exception:
                                continue

                    if not isinstance(payload, dict):
                        raise RuntimeError(f"start_logging: unexpected result: {r.content}")

                    # FastMCP wraps tool returns into {"result": {...}}
                    if isinstance(payload.get("result"), dict):
                        payload = payload["result"]

                    stream_id = payload.get("stream_id")
                    cursor = int(payload.get("cursor", 0) or 0)
                    print(f"start_logging: stream_id={stream_id} cursor={cursor}", flush=True)

                    print("run_lisp...", flush=True)
                    with anyio.fail_after(30):
                        r = await session.call_tool(
                            "run_lisp",
                            {"expr": "(getvar 'ACADVER)", "wait": True, "timeout_sec": 10.0},
                        )
                    if r.isError:
                        raise RuntimeError(f"run_lisp error: {r.content}")

                    rc = None
                    if r.structuredContent is not None:
                        rc = r.structuredContent
                    else:
                        for item in r.content or []:
                            txt = getattr(item, "text", None)
                            if not isinstance(txt, str):
                                continue
                            try:
                                rc = json.loads(txt)
                                break
                            except Exception:
                                continue

                    if not isinstance(rc, dict):
                        raise RuntimeError(f"run_lisp: unexpected result: {r.content}")

                    if isinstance(rc.get("result"), dict):
                        rc = rc["result"]

                    last_prompt = rc.get("last_prompt", "")
                    print("last_prompt:", last_prompt, flush=True)

                    log_block = rc.get("log") or {}
                    if log_block.get("text"):
                        print("---log from run_lisp---", flush=True)
                        print(str(log_block.get("text")).strip(), flush=True)
                        cursor = int(log_block.get("cursor", cursor))

                    if stream_id:
                        # Give AutoCAD a moment to flush LOGFILE
                        await anyio.sleep(0.5)
                        print("get_new_output_since...", flush=True)
                        with anyio.fail_after(30):
                            r2 = await session.call_tool(
                                "get_new_output_since",
                                {"stream_id": stream_id, "cursor": cursor, "max_bytes": 65536},
                            )
                        if r2.isError:
                            raise RuntimeError(f"get_new_output_since error: {r2.content}")

                        oc = None
                        if r2.structuredContent is not None:
                            oc = r2.structuredContent
                        else:
                            for item in r2.content or []:
                                txt = getattr(item, "text", None)
                                if not isinstance(txt, str):
                                    continue
                                try:
                                    oc = json.loads(txt)
                                    break
                                except Exception:
                                    continue

                        if not isinstance(oc, dict):
                            raise RuntimeError(f"get_new_output_since: unexpected result: {r2.content}")

                        if isinstance(oc.get("result"), dict):
                            oc = oc["result"]

                        text = (oc.get("text") or "").strip()
                        if text:
                            print("---new output since cursor---", flush=True)
                            print(text, flush=True)

    except Exception:
        tail = _tail_text(err_path)
        if tail.strip():
            print("---server stderr tail---", flush=True)
            print(tail.strip(), flush=True)
        raise


if __name__ == "__main__":
    anyio.run(main)

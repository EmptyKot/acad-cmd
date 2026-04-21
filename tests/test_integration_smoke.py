from __future__ import annotations

from pathlib import Path
import json
import os
import unittest
import uuid


def _unwrap_result(result):
    payload = result.structuredContent if result.structuredContent is not None else None
    if payload is None:
        for item in result.content or []:
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


async def _run_smoke() -> None:
    import anyio
    from mcp.client.session import ClientSession
    from mcp.client.stdio import StdioServerParameters, stdio_client

    root = Path(__file__).resolve().parents[1]
    py = root / ".venv" / "Scripts" / "python.exe"
    if not py.exists():
        raise unittest.SkipTest(f"Missing python exe: {py}")

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

            with anyio.fail_after(30):
                listed = await session.list_tools()
            names = {t.name for t in listed.tools}
            for required in ("get_status", "open_drawing", "dict_list", "dict_xrecord_set", "dict_xrecord_get", "selection"):
                if required not in names:
                    raise AssertionError(f"Missing tool: {required}")

            with anyio.fail_after(30):
                rs = await session.call_tool("get_status", {})
            if rs.isError:
                raise AssertionError(f"get_status error: {rs.content}")
            status = _unwrap_result(rs)
            if not isinstance(status, dict) or not status.get("connected"):
                raise AssertionError(f"AutoCAD not connected: {status!r}")

            with anyio.fail_after(30):
                rl = await session.call_tool("start_logging", {"mode": "logfile", "reset": False})
            if rl.isError:
                raise AssertionError(f"start_logging error: {rl.content}")
            started = _unwrap_result(rl)
            if not isinstance(started, dict):
                raise AssertionError(f"start_logging payload invalid: {started!r}")

            stream_id = started.get("stream_id")
            try:
                with anyio.fail_after(30):
                    rr = await session.call_tool("run_lisp", {"expr": "(getvar 'ACADVER)", "wait": True, "timeout_sec": 10.0})
                if rr.isError:
                    raise AssertionError(f"run_lisp error: {rr.content}")

                dict_name = f"MCP_SMOKE_{uuid.uuid4().hex[:8]}"
                key = "SMOKE_KEY"
                values = [{"code": 1, "value": "hello"}, {"code": 90, "value": 42}]

                with anyio.fail_after(30):
                    r_set = await session.call_tool(
                        "dict_xrecord_set",
                        {
                            "dict_name": dict_name,
                            "key": key,
                            "values": values,
                            "timeout_sec": 10.0,
                            "overwrite": True,
                        },
                    )
                if r_set.isError:
                    raise AssertionError(f"dict_xrecord_set error: {r_set.content}")
                p_set = _unwrap_result(r_set)
                if not isinstance(p_set, dict) or not p_set.get("written"):
                    raise AssertionError(f"dict_xrecord_set unexpected payload: {p_set!r}")

                with anyio.fail_after(30):
                    r_get = await session.call_tool(
                        "dict_xrecord_get",
                        {"dict_name": dict_name, "key": key, "timeout_sec": 10.0},
                    )
                if r_get.isError:
                    raise AssertionError(f"dict_xrecord_get error: {r_get.content}")
                p_get = _unwrap_result(r_get)
                if not isinstance(p_get, dict) or not p_get.get("found"):
                    raise AssertionError(f"dict_xrecord_get unexpected payload: {p_get!r}")

                # Create and preselect one entity so selection() returns from PickFirst path.
                setup_expr = (
                    "(progn "
                    "(setq e (entmakex (list (cons 0 \\\"LINE\\\") (cons 10 '(0.0 0.0 0.0)) (cons 11 '(10.0 0.0 0.0))))) "
                    "(if e (sssetfirst nil (ssadd e))) "
                    "(princ))"
                )
                with anyio.fail_after(30):
                    r_setup = await session.call_tool("run_lisp", {"expr": setup_expr, "wait": True, "timeout_sec": 10.0})
                if r_setup.isError:
                    raise AssertionError(f"selection setup run_lisp error: {r_setup.content}")

                with anyio.fail_after(45):
                    r_sel = await session.call_tool("selection", {"timeout_sec": 10.0, "max_objects": 1})
                if r_sel.isError:
                    raise AssertionError(f"selection error: {r_sel.content}")
                p_sel = _unwrap_result(r_sel)
                if not isinstance(p_sel, dict):
                    raise AssertionError(f"selection payload invalid: {p_sel!r}")
                if p_sel.get("timed_out"):
                    raise AssertionError(f"selection timed out: {p_sel!r}")
                if int(p_sel.get("count") or 0) < 1:
                    raise AssertionError(f"selection empty: {p_sel!r}")
                objs = p_sel.get("objects") or []
                if not objs or not isinstance(objs[0], dict) or not objs[0].get("handle") or not objs[0].get("type"):
                    raise AssertionError(f"selection objects invalid: {p_sel!r}")
            finally:
                if stream_id:
                    with anyio.move_on_after(15):
                        await session.call_tool("stop_logging", {"stream_id": stream_id})


class IntegrationSmokeTests(unittest.TestCase):
    @unittest.skipUnless(
        os.environ.get("ACAD_MCP_RUN_INTEGRATION", "").strip() == "1",
        "Set ACAD_MCP_RUN_INTEGRATION=1 to run real AutoCAD smoke tests",
    )
    def test_stdio_smoke_real_autocad(self) -> None:
        import anyio

        anyio.run(_run_smoke)


if __name__ == "__main__":
    unittest.main()

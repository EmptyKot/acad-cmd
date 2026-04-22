from collections import deque
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from acad_cmd.autocad_bridge import WaitResult
from acad_cmd.output_log import OutputStream
import acad_cmd.tools as tools


class _FakeAudit:
    def __init__(self) -> None:
        self.events = []

    def log(self, event, payload, *, dwg=None) -> None:
        self.events.append((event, payload, dwg))


class _FakeBridge:
    def __init__(self) -> None:
        self.acad = SimpleNamespace(HWND=0)
        self.last_sent = None
        self._dwg = "D:/tmp/Test1.dwg"
        self._status_snapshot = {
            "connected": True,
            "busy": False,
            "stale": False,
            "source": "live",
            "error_class": None,
            "error_message": None,
            "dwg": self._dwg,
            "acadver": "24.0s (LMS Tech)",
            "acad_hwnd": 0,
            "acad_pid": None,
            "cmdactive": 0,
            "locked_major": 24,
            "bound_progid": "AutoCAD.Application.24",
        }

    def ensure_connection(self) -> bool:
        return True

    def get_status_snapshot(self):
        out = dict(self._status_snapshot)
        out["dwg"] = self._dwg
        return out

    def get_dwg_label(self):
        return self._dwg

    def get_variable(self, name: str):
        if name == "ACADVER":
            return "24.0s (LMS Tech)"
        if name == "CMDACTIVE":
            return 0
        return None

    def send_command(self, command: str) -> str:
        self.last_sent = command
        return "cmd-1"

    def wait_for_idle(self, timeout_sec: float, poll_interval_sec: float = 0.1) -> WaitResult:
        return WaitResult(completed=True, needs_input=False, quiescent=True)

    def get_last_prompt(self) -> str:
        return "Command:"

    def open_drawing(self, path: str, *, timeout_sec: float = 30.0, read_only: bool = False):
        norm = path.replace("\\", "/")
        self._dwg = norm
        return {
            "path": norm,
            "dwg": norm,
            "already_open": False,
            "opened": True,
            "activated": True,
            "read_only": bool(read_only),
        }


class _FakeStreams:
    def __init__(self) -> None:
        self.default = OutputStream(
            stream_id="stream-1",
            mode="logfile",
            logfile_path="D:/tmp/acad.log",
            cursor=100,
            ring=deque(maxlen=16),
            started_by_server=True,
        )
        self.by_id = {self.default.stream_id: self.default}

    def get_default(self):
        return self.default

    def get(self, stream_id: str):
        return self.by_id.get(stream_id)

    def read_new(self, stream_id: str, cursor: int, max_bytes: int):
        return "line-1", 106, False

    def read_tail(self, stream_id: str, tail_bytes: int = 8192):
        return "tail-1"


@contextmanager
def _patched_state(temp_state):
    old_state = tools.state
    tools.state = temp_state
    try:
        yield
    finally:
        tools.state = old_state


def _make_fake_state():
    return SimpleNamespace(
        session_id="session-test",
        bridge=_FakeBridge(),
        streams=_FakeStreams(),
        audit=_FakeAudit(),
        event_bridge_enabled=False,
        event_bridge_client=None,
        event_bridge_pid=None,
        event_bridge_heartbeat_timeout_sec=6.0,
        event_bridge_max_dropped_for_wait=0,
        event_bridge_object_events_requested=False,
        event_bridge_autoload_enabled=False,
        event_bridge_autoload_dll=None,
        event_bridge_last_prepare_issue=None,
    )


class ToolsContractTests(unittest.TestCase):
    def test_get_status_contract(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            res = tools.get_status(None)

        self.assertIn("session_id", res)
        self.assertIn("connected", res)
        self.assertIn("dwg", res)
        self.assertIn("acadver", res)
        self.assertIn("default_stream", res)
        self.assertIn("busy", res)
        self.assertIn("stale", res)
        self.assertIn("source", res)
        self.assertIn("event_bridge", res)
        self.assertEqual(res["session_id"], "session-test")
        self.assertEqual(res["connected"], True)
        self.assertEqual(res["dwg"], "D:/tmp/Test1.dwg")
        self.assertEqual(res["source"], "live")
        self.assertEqual(res["event_bridge"]["enabled"], False)
        self.assertEqual(res["event_bridge"]["available"], False)
        self.assertEqual(res["event_bridge"]["connected"], False)

    def test_parse_bool_env(self) -> None:
        with patch.dict("os.environ", {"AUTOCAD_MCP_EVENT_BRIDGE_ENABLED": "1"}, clear=False):
            self.assertTrue(tools._parse_bool_env("AUTOCAD_MCP_EVENT_BRIDGE_ENABLED", default=False))
        with patch.dict("os.environ", {"AUTOCAD_MCP_EVENT_BRIDGE_ENABLED": "true"}, clear=False):
            self.assertTrue(tools._parse_bool_env("AUTOCAD_MCP_EVENT_BRIDGE_ENABLED", default=False))
        with patch.dict("os.environ", {"AUTOCAD_MCP_EVENT_BRIDGE_ENABLED": "0"}, clear=False):
            self.assertFalse(tools._parse_bool_env("AUTOCAD_MCP_EVENT_BRIDGE_ENABLED", default=True))
        with patch.dict("os.environ", {"AUTOCAD_MCP_EVENT_BRIDGE_ENABLED": "off"}, clear=False):
            self.assertFalse(tools._parse_bool_env("AUTOCAD_MCP_EVENT_BRIDGE_ENABLED", default=True))
        with patch.dict("os.environ", {"AUTOCAD_MCP_EVENT_BRIDGE_ENABLED": "unexpected"}, clear=False):
            self.assertFalse(tools._parse_bool_env("AUTOCAD_MCP_EVENT_BRIDGE_ENABLED", default=False))

    def test_send_command_contract_with_log_block(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            res = tools.send_command(None, command="._LINE", wait=True, timeout_sec=1.0)

        self.assertEqual(res["command_id"], "cmd-1")
        self.assertEqual(res["dwg"], "D:/tmp/Test1.dwg")
        self.assertEqual(res["sent"], "._LINE")
        self.assertEqual(res["completed"], True)
        self.assertEqual(res["needs_input"], False)
        self.assertEqual(res["last_prompt"], "Command:")
        self.assertIsInstance(res["log"], dict)
        self.assertEqual(res["log"]["stream_id"], "stream-1")
        self.assertEqual(res["log"]["cursor"], 106)
        self.assertEqual(res["log"]["text"], "line-1")

        logged_names = [e[0] for e in fake.audit.events]
        self.assertIn("send_command", logged_names)
        self.assertIn("send_command_result", logged_names)

    def test_open_drawing_contract(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            res = tools.open_drawing(None, path="D:/tmp/NewProject.dwg", timeout_sec=5.0, read_only=True)

        self.assertEqual(res["path"], "D:/tmp/NewProject.dwg")
        self.assertEqual(res["dwg_before"], "D:/tmp/Test1.dwg")
        self.assertEqual(res["dwg"], "D:/tmp/NewProject.dwg")
        self.assertEqual(res["dwg_opened"], "D:/tmp/NewProject.dwg")
        self.assertEqual(res["opened"], True)
        self.assertEqual(res["already_open"], False)
        self.assertEqual(res["activated"], True)
        self.assertEqual(res["read_only"], True)

        logged_names = [e[0] for e in fake.audit.events]
        self.assertIn("open_drawing", logged_names)

    def test_get_last_output_defaults_to_logfile(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            res = tools.get_last_output(None)

        self.assertEqual(res["source"], "logfile")
        self.assertEqual(res["text"], "tail-1")
        self.assertEqual(res["dwg"], "D:/tmp/Test1.dwg")

    def test_get_last_output_lastprompt_legacy_source(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            res = tools.get_last_output(None, source="lastprompt")

        self.assertEqual(res["source"], "lastprompt")
        self.assertEqual(res["text"], "Command:")
        self.assertEqual(res["dwg"], "D:/tmp/Test1.dwg")

    def test_dict_list_contract_and_expression(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            with patch("acad_cmd.tools._run_lisp_json", return_value={"ok": True, "dicts": [{"name": "X"}]}) as run_json:
                res = tools.dict_list(None, timeout_sec=3.0)

        self.assertEqual(res, {"dicts": [{"name": "X"}]})
        expr = run_json.call_args.args[1]
        self.assertIn("(mcp-dict-list)", expr)

    def test_dict_xrecord_set_contract_and_expression(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            with patch("acad_cmd.tools._run_lisp_json", return_value={"ok": True, "written": True}) as run_json:
                res = tools.dict_xrecord_set(
                    None,
                    dict_name="MCP_TEST",
                    key="K1",
                    values=[
                        {"code": 1, "value": "abc"},
                        {"code": 90, "value": 42},
                    ],
                    timeout_sec=3.0,
                    overwrite=False,
                )

        self.assertEqual(res, {"written": True})
        expr = run_json.call_args.args[1]
        self.assertIn("(mcp-xrecord-set", expr)
        self.assertIn("\"MCP_TEST\"", expr)
        self.assertIn("\"K1\"", expr)
        self.assertIn("(cons 1 \"abc\")", expr)
        self.assertIn("(cons 90 42)", expr)
        self.assertIn("nil)", expr)

    def test_dict_validates_required_args(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            with self.assertRaisesRegex(ValueError, "dict_name must be non-empty"):
                tools.dict_keys(None, dict_name="", timeout_sec=3.0)
            with self.assertRaisesRegex(ValueError, "key must be non-empty"):
                tools.dict_xrecord_get(None, dict_name="A", key="", timeout_sec=3.0)

    def test_timeout_validation(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            with self.assertRaisesRegex(ValueError, "timeout_sec must be in range"):
                tools.send_command(None, command="._LINE", timeout_sec=0.0, wait=True)
            with self.assertRaisesRegex(ValueError, "poll_interval_sec must be in range"):
                tools.send_command(None, command="._LINE", timeout_sec=1.0, wait=True, poll_interval_sec=10.0)
            with self.assertRaisesRegex(ValueError, "timeout_sec must be in range"):
                tools.dict_list(None, timeout_sec=2000.0)

    def test_selection_contract_implied_pickfirst(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            with patch("acad_cmd.tools._collect_selection_stream_lite") as collect_sel:
                collect_sel.return_value = {
                    "req_id": "r1",
                    "count": 1,
                    "errno": 0,
                    "objects": [{"handle": "10A", "type": "LINE"}],
                    "timed_out": False,
                    "cursor": 106,
                }
                with patch("acad_cmd.tools.send_command", return_value={"log": {"text": "", "cursor": 106}}) as send_cmd:
                    out = tools.selection(None, timeout_sec=2.0, max_objects=1, alert_message="Select one LINE")

        self.assertEqual(out["dwg"], "D:/tmp/Test1.dwg")
        self.assertEqual(out["count"], 1)
        self.assertEqual(out["objects"][0]["handle"], "10A")
        self.assertEqual(send_cmd.call_count, 1)
        first_expr = send_cmd.call_args_list[0].args[1]
        self.assertIn("(mcp-selection-implied-lite", first_expr)
        self.assertNotIn("(mcp-selection-prompt-lite", first_expr)
        self.assertNotIn("Select one LINE", first_expr)

        phases = [payload["phase"] for event, payload, _ in fake.audit.events if event == "selection"]
        self.assertEqual(phases, ["implied"])

    def test_selection_contract_prompt_fallback(self) -> None:
        fake = _make_fake_state()
        with _patched_state(fake):
            with patch("acad_cmd.tools._collect_selection_stream_lite") as collect_sel:
                collect_sel.side_effect = [
                    {
                        "req_id": "r1",
                        "count": 0,
                        "errno": 0,
                        "objects": [],
                        "timed_out": False,
                        "cursor": 106,
                    },
                    {
                        "req_id": "r2",
                        "count": 1,
                        "errno": 0,
                        "objects": [{"handle": "20B", "type": "CIRCLE"}],
                        "timed_out": False,
                        "cursor": 110,
                    },
                ]
                with patch("acad_cmd.tools.send_command") as send_cmd:
                    send_cmd.side_effect = [
                        {"log": {"text": "", "cursor": 106}},
                        {"log": {"text": "", "cursor": 110}},
                    ]
                    out = tools.selection(
                        None,
                        timeout_sec=3.0,
                        prompt="Pick one",
                        max_objects=1,
                        alert_message="Select exactly one circle",
                    )

        self.assertEqual(out["count"], 1)
        self.assertEqual(out["objects"][0]["type"], "CIRCLE")
        self.assertEqual(send_cmd.call_count, 2)
        prompt_expr = send_cmd.call_args_list[1].args[1]
        self.assertIn("(mcp-selection-prompt-lite", prompt_expr)
        self.assertIn("\"Select exactly one circle\"", prompt_expr)
        phases = [payload["phase"] for event, payload, _ in fake.audit.events if event == "selection"]
        self.assertEqual(phases, ["implied", "prompt"])


if __name__ == "__main__":
    unittest.main()

from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from acad_cmd.protocol import consume_complete_lines, extract_mcp_json, extract_mcp_json_messages


class ProtocolTests(unittest.TestCase):
    def test_extract_last_json_marker(self) -> None:
        text = (
            "noise\n"
            "[MCP:JSON]{\"ok\":true,\"step\":1}\n"
            "more noise\n"
            "[MCP:JSON]{\"ok\":true,\"step\":2}\n"
        )
        obj = extract_mcp_json(text)
        self.assertEqual(obj.get("ok"), True)
        self.assertEqual(obj.get("step"), 2)

    def test_extract_messages_skips_invalid(self) -> None:
        text = (
            "[MCP:JSON]{\"ok\":true,\"event\":\"start\"}\n"
            "[MCP:JSON]{bad json}\n"
            "plain\n"
            "[MCP:JSON]{\"ok\":true,\"event\":\"done\"}\n"
        )
        msgs = extract_mcp_json_messages(text)
        self.assertEqual(len(msgs), 2)
        self.assertEqual(msgs[0].get("event"), "start")
        self.assertEqual(msgs[1].get("event"), "done")

    def test_consume_complete_lines(self) -> None:
        lines, rest = consume_complete_lines("a\nb\npartial")
        self.assertEqual(lines, ["a", "b"])
        self.assertEqual(rest, "partial")


if __name__ == "__main__":
    unittest.main()

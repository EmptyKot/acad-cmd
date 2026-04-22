# acad-cmd (AutoCAD MCP server)

Local MCP server (stdio / JSON-RPC) that connects to AutoCAD on Windows via COM (pywin32) and exposes command-line I/O as MCP tools.

What it does:

- send text to the AutoCAD command line (`SendCommand`)
- open DWG files and switch active AutoCAD context to the opened drawing
- use persistent full command history logging via AutoCAD `LOGFILEMODE` / `LOGFILENAME` (primary output source)
- optionally read `LASTPROMPT` for legacy compatibility
- optionally use in-process `.NET` event bridge (`AcadEventBridge`) for structured command/document lifecycle events
- use event-first completion waits (with automatic fallback to COM idle wait)
- write an audit log (JSONL) for every tool call

## Requirements

- Windows
- AutoCAD (tested primarily with AutoCAD 2021; other versions may work)
- Python 3.10+

## Compatibility (current baseline)

| Component | Status |
| --- | --- |
| Windows 10/11 | Supported |
| AutoCAD 2021 (major 24) | Primary tested target |
| Python 3.10+ | Supported |

## Install

```bat
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install .
```

If your `python` command opens the Microsoft Store, use `py -3.11` as above.

## Run (standalone)

Starting AutoCAD first is recommended (and most reliable), then:

```bat
acad-cmd
```

Or use the provided helper script (creates `.venv` and installs on first run):

```bat
start_server.bat
```

If you want the server to launch AutoCAD, set `AUTOCAD_MCP_ACAD_EXE` to the full path
to `acad.exe` (the server will start the process and then attach via COM).

If you have multiple Autodesk products installed/running (e.g. Civil 3D and AutoCAD)
and want to force a specific AutoCAD major, set `AUTOCAD_MCP_TARGET_MAJOR`.

Example: AutoCAD 2021 is major `24`:

```bat
set AUTOCAD_MCP_TARGET_MAJOR=24
```

Important connection behavior:

- AutoCAD instances are discovered via COM (first `GetActiveObject`, then optionally `Dispatch`).
- Some installations do not expose a normally-launched UI instance to `GetActiveObject`.
  In that case, `Dispatch` can attach to a running instance or spawn a new automation-enabled instance.
- To disable spawning a new AutoCAD instance, set `AUTOCAD_MCP_ALLOW_NEW_INSTANCE=0`.
- If you want a normally-launched UI instance to be attachable, start AutoCAD with automation enabled
  (commonly `acad.exe /automation`, but this can vary by installation).

Runtime logs are written under `logs/acad-cmd/<session_id>/`.

## Event Bridge (optional)

`acad-cmd` can use an in-process AutoCAD plugin (`AcadEventBridge`) over named pipe for structured events.

- command/lisp completion can be resolved from event stream (faster and more deterministic than text-only checks)
- document lifecycle events are available in status/diagnostics
- if bridge is unavailable or degraded, server automatically falls back to legacy COM idle wait and logfile path

Plugin source and smoke commands are in `plugins/AcadEventBridge/README.md`.

## Configuration (environment variables)

Connection / version selection:

- `AUTOCAD_MCP_TARGET_MAJOR` (optional): pin AutoCAD major version (e.g. `24` for AutoCAD 2021).
- `AUTOCAD_MCP_ALLOW_NEW_INSTANCE` (default: allow): set to `0` to prevent spawning a new `acad.exe` via COM activation.
- `AUTOCAD_MCP_USE_DISPATCH` (default: off unless `AUTOCAD_MCP_TARGET_MAJOR` is set): force trying `Dispatch` activation.
- `AUTOCAD_MCP_PREFER_CURVER` (default: off): prefer registry `CurVer` ProgID when resolving AutoCAD version.
- `AUTOCAD_MCP_EVENT_BRIDGE_ENABLED` (default: `0`): enables bridge integration path.

Event bridge behavior:

- `AUTOCAD_MCP_EVENT_BRIDGE_HEARTBEAT_TIMEOUT_SEC` (default: `6.0`): heartbeat freshness threshold for bridge waits.
- `AUTOCAD_MCP_EVENT_BRIDGE_MAX_DROPPED_FOR_WAIT` (default: `0`): max tolerated dropped queue messages before degrading to fallback wait.
- `AUTOCAD_MCP_EVENT_BRIDGE_OBJECT_EVENTS_ENABLED` (default: `1`): desired plugin-side object events mode (`AEB_OBJECT_EVENTS_ON/OFF` sync).
- `AUTOCAD_MCP_EVENT_BRIDGE_AUTOLOAD` (default: `1`): when bridge is enabled, allow best-effort plugin autoload (`NETLOAD`) if pipe is unavailable.
- `AUTOCAD_MCP_EVENT_BRIDGE_AUTOLOAD_DLL` (optional): explicit path to `AcadEventBridge.dll` for autoload; otherwise server tries default build locations.

Launching AutoCAD:

- `AUTOCAD_MCP_ACAD_EXE` (optional): full path to `acad.exe` to explicitly launch AutoCAD.
- `AUTOCAD_MCP_ACAD_ARGS` (optional): extra args passed to `acad.exe` when launching.
- `AUTOCAD_MCP_LAUNCH_WAIT_SEC` (default: `30`): how long to wait for AutoCAD to start before retrying COM attach.

## Claude Desktop config example

`%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "acad-cmd": {
      "command": "C:/path/to/project/.venv/Scripts/python.exe",
      "args": ["-m", "acad_cmd.server"]
    }
  }
}
```

Notes:

- `.venv` is not committed to git and will not appear after cloning; create it locally (or run `start_server.bat`).
- If you want the server to auto-launch AutoCAD, set `AUTOCAD_MCP_ACAD_EXE` to the full path to `acad.exe`.

## Tools

All tools return JSON (FastMCP commonly wraps results as `{ "result": ... }`).

- `get_status()`
  - returns connection info (DWG label, `ACADVER`, window handle / PID when available) and default stream details
  - includes stability fields for busy/degraded states: `busy`, `stale`, `source`, `error_class`, `cmdactive`
  - includes binding info to avoid cross-version drift: `locked_major`, `bound_progid`
  - includes `event_bridge` diagnostics: availability/connectivity, pipe/plugin metadata, heartbeat/queue state, degradation reason, and optional service status
- `send_command(command, timeout_sec, wait=true, poll_interval_sec=0.1)`
  - sends raw command line text
  - when `wait=true`: uses bridge event-first completion wait when available, otherwise falls back to COM idle wait
  - ensures a default logfile stream and returns a `log` block with new output and updated cursor
  - returns wait diagnostics (`wait_source`, `wait_completion_event`, `wait_completion_seq`, `wait_fallback_used`, `bridge_wait_prepare_issue`)
- `open_drawing(path, timeout_sec, read_only=false)`
  - opens a DWG and guarantees active context switch to the target document (or errors on timeout)
  - returns `dwg_before`, `dwg`, `already_open`, `opened`, `activated`
- `get_last_output(source=lastprompt|logfile)`
  - default source is `logfile`
  - `logfile`: returns a tail of the current default logfile stream (auto-starts logfile stream if needed)
  - `lastprompt`: reads `LASTPROMPT` (legacy/fallback source)
- `start_logging(mode=logfile|lastprompt, logfile_path=null, reset=false)`
  - starts a stream and returns `{stream_id, cursor, ...}`
  - `logfile` mode enables `LOGFILEMODE` and tracks `LOGFILENAME`
  - `lastprompt` mode is kept only for backward compatibility
  - if `logfile_path` is not provided, the server prefers AutoCAD's current `LOGFILENAME` to avoid path issues
- `get_new_output_since(stream_id, cursor, max_bytes=65536)`
  - reads appended logfile bytes and returns `{text, new_cursor, truncated}`
- `stop_logging(stream_id)`
  - stops a stream; best-effort disables `LOGFILEMODE` when the last server-started logfile stream is stopped
- `load_lisp_file(path, timeout_sec, wait=true)`
  - sends `(load "...")` (path normalized for AutoCAD)
  - when `wait=true`: uses bridge LISP completion events first, then COM fallback if needed
- `run_lisp(expr, timeout_sec, wait=true)`
  - executes an AutoLISP expression/script via `SendCommand` with start/end markers in the command history
  - when `wait=true`: uses bridge LISP completion events first, then COM fallback if needed
- `selection(timeout_sec, prompt=null, filter=null, max_objects=null, alert_message=null)`
  - returns currently selected objects (PickFirst); if none, prompts the user to select objects
  - when `alert_message` is provided and interactive selection is needed, shows standard AutoCAD `alert`
  - returns only `handle` + `type` for each object
- `dict_list/dict_keys/dict_xrecord_get/dict_xrecord_set/dict_xrecord_delete/dict_delete(..., timeout_sec)`
  - dictionary tools now require explicit `timeout_sec` as well
  - timeout range for all waiting tools: `0.1..1800` seconds

## Troubleshooting notes

- AutoLISP file loading can be blocked by AutoCAD security settings.
  - Add the folder with your `.lsp` files to AutoCAD **Trusted Locations**.
  - Check `SECURELOAD` behavior (do not weaken security globally for production).
- If COM calls fail with "callee busy", the server retries with backoff.
- `LOGFILEMODE`/`LOGFILENAME` writes a file in AutoCAD's current codepage; decoding uses your Windows preferred encoding with fallback.

## Development / smoke tests

- `scripts/mcp_smoketest_stdio.py`: spawns the server over stdio, lists tools, starts logging, runs a small LISP expression.
- `scripts/mcp_sanity_acadver.py`: direct COM sanity check that `LOGFILEMODE` output grows after sending `(getvar 'ACADVER)`.
- `scripts/mcp_baseline_capture.py`: roadmap step 1 baseline capture (`get_status`, `send_command`, `run_lisp`, `LOGFILEMODE`) with JSON report in `out/baseline/`.

Run unit tests:

```bat
python -m unittest discover -s tests -p "test_*.py"
```

Run integration tests (real AutoCAD required):

```bat
set ACAD_MCP_RUN_INTEGRATION=1
python -m unittest discover -s tests -p "test_*.py"
```

## Repository Docs

- Contribution guide: `CONTRIBUTING.md`
- Security policy: `SECURITY.md`
- License: `LICENSE`

# acad-cmd (AutoCAD MCP server)

Local MCP server (stdio / JSON-RPC) that connects to AutoCAD on Windows via COM (pywin32) and exposes command-line I/O as MCP tools.

What it does:

- send text to the AutoCAD command line (`SendCommand`)
- read output via `LASTPROMPT` (low overhead)
- stream full command history via AutoCAD `LOGFILEMODE` / `LOGFILENAME`
- write an audit log (JSONL) for every tool call

## Requirements

- Windows
- AutoCAD (tested primarily with AutoCAD 2021; other versions may work)
- Python 3.10+

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

## Configuration (environment variables)

Connection / version selection:

- `AUTOCAD_MCP_TARGET_MAJOR` (optional): pin AutoCAD major version (e.g. `24` for AutoCAD 2021).
- `AUTOCAD_MCP_ALLOW_NEW_INSTANCE` (default: allow): set to `0` to prevent spawning a new `acad.exe` via COM activation.
- `AUTOCAD_MCP_USE_DISPATCH` (default: off unless `AUTOCAD_MCP_TARGET_MAJOR` is set): force trying `Dispatch` activation.
- `AUTOCAD_MCP_PREFER_CURVER` (default: off): prefer registry `CurVer` ProgID when resolving AutoCAD version.

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
- `send_command(command, wait=true, timeout_sec=10, poll_interval_sec=0.1)`
  - sends raw command line text; when `wait=true` waits until AutoCAD is idle or timeout
  - if a default logfile stream is active, also returns a `log` block with new output and updated cursor
- `get_last_output(source=lastprompt|logfile)`
  - `lastprompt`: reads `LASTPROMPT`
  - `logfile`: returns a tail of the current default logfile stream
- `start_logging(mode=logfile|lastprompt, logfile_path=null, reset=false)`
  - starts a stream and returns `{stream_id, cursor, ...}`
  - `logfile` mode enables `LOGFILEMODE` and tracks `LOGFILENAME`
  - if `logfile_path` is not provided, the server prefers AutoCAD's current `LOGFILENAME` to avoid path issues
- `get_new_output_since(stream_id, cursor, max_bytes=65536)`
  - reads appended logfile bytes and returns `{text, new_cursor, truncated}`
- `stop_logging(stream_id)`
  - stops a stream; best-effort disables `LOGFILEMODE` when the last server-started logfile stream is stopped
- `load_lisp_file(path, wait=true, timeout_sec=10)`
  - sends `(load "...")` (path normalized for AutoCAD)
- `run_lisp(expr, wait=true, timeout_sec=10)`
  - executes an AutoLISP expression/script via `SendCommand` with start/end markers in the command history
- `selection(timeout_sec=300, prompt=null, filter=null, max_objects=null)`
  - returns currently selected objects (PickFirst); if none, prompts the user to select objects
  - returns only `handle` + `type` for each object

## Troubleshooting notes

- AutoLISP file loading can be blocked by AutoCAD security settings.
  - Add the folder with your `.lsp` files to AutoCAD **Trusted Locations**.
  - Check `SECURELOAD` behavior (do not weaken security globally for production).
- If COM calls fail with "callee busy", the server retries with backoff.
- `LOGFILEMODE`/`LOGFILENAME` writes a file in AutoCAD's current codepage; decoding uses your Windows preferred encoding with fallback.

## Development / smoke tests

- `scripts/mcp_smoketest_stdio.py`: spawns the server over stdio, lists tools, starts logging, runs a small LISP expression.
- `scripts/mcp_sanity_acadver.py`: direct COM sanity check that `LOGFILEMODE` output grows after sending `(getvar 'ACADVER)`.

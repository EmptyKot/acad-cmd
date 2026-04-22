# AcadEventBridge (skeleton)

Current milestone: roadmap step 4 (`PipeServer` + `hello`).

## Object Events (opt-in)

Object-level events are disabled by default and can be enabled:

1) via environment variable before AutoCAD start:

```bat
set AUTOCAD_MCP_EVENT_BRIDGE_OBJECT_EVENTS_ENABLED=1
```

2) at runtime via plugin debug commands (no restart):

```text
AEB_OBJECT_EVENTS_ON
AEB_OBJECT_EVENTS_OFF
```

Python MCP server uses runtime commands to align plugin state automatically, so manual AutoCAD setup is not required in normal MCP flow.
`hello` contains `object_events_enabled`, and `AEB_STATUS` prints `object_events_enabled=...`.

## Build

```bat
dotnet build plugins\AcadEventBridge\AcadEventBridge.csproj -c Debug
```

If `AcadEventBridge.dll` is already loaded in AutoCAD, the default output file is locked.
Use a temporary output directory for rebuilds while AutoCAD is running:

```bat
dotnet build plugins\AcadEventBridge\AcadEventBridge.csproj -c Debug -p:OutDir=bin\Debug\net48\hotbuild\
```

## NETLOAD

Load from AutoCAD command line:

```text
NETLOAD
```

Then pick:

```text
plugins\AcadEventBridge\bin\Debug\net48\AcadEventBridge.dll
```

Smoke command (after loading):

```text
AEB_PING
```

`AEB_STATUS` should print pipe diagnostics:

```text
AcadEventBridge loaded. pipe_name=acad-event-bridge-<pid>, pipe_running=True, last_seq=..., queue_depth=..., dropped_count=...
```

If you still see old output (`AcadEventBridge loaded.` only), restart AutoCAD before retesting:
AutoCAD keeps the first loaded assembly in memory.

## Pipe hello smoke

From repo root:

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_pipe_hello_smoke.py
```

Expected output contains:

```json
{"hello":{"type":"hello","protocol":1,"plugin":"AcadEventBridge",...}}
```

Heartbeat smoke (step 5):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_pipe_hello_smoke.py --expect-heartbeat --timeout-sec 8
```

Document lifecycle smoke (step 7):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_document_events_smoke.py --timeout-sec 12
```

Command lifecycle smoke (step 8):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_command_events_smoke.py --command "_.REGEN" --timeout-sec 12
```

LISP/helper lifecycle smoke (step 9):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_lisp_helper_events_smoke.py --timeout-sec 14
```

Python bridge client smoke (step 10):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_client_smoke.py --timeout-sec 8
```

Command waiter smoke (step 13):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\command_waiter_smoke.py --timeout-sec 12
```

send_command integration smoke (step 14):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\send_command_event_wait_smoke.py --timeout-sec 12
```

LISP waiter integration smoke (step 15):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\lisp_waiter_integration_smoke.py
```

Bridge hardening smoke (step 16):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_hardening_smoke.py
```

Object events smoke (step 18, opt-in):

```bat
set PYTHONPATH=%CD%\src
set AUTOCAD_MCP_EVENT_BRIDGE_OBJECT_EVENTS_ENABLED=1
.venv\Scripts\python.exe scripts\bridge_object_events_smoke.py
```

Request/response smoke (step 20):

```bat
set PYTHONPATH=%CD%\src
.venv\Scripts\python.exe scripts\bridge_request_response_smoke.py
```

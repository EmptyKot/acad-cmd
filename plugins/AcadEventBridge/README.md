# AcadEventBridge (skeleton)

Current milestone: roadmap step 4 (`PipeServer` + `hello`).

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

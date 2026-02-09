@echo off
setlocal enableextensions

set "ROOT=%~dp0"
set "PY=%ROOT%.venv\Scripts\python.exe"

if not exist "%PY%" (
  echo Creating venv in %ROOT%.venv ...
  py -3.11 -m venv "%ROOT%.venv" 1>nul 2>nul
  if errorlevel 1 (
    py -3 -m venv "%ROOT%.venv"
    if errorlevel 1 (
      echo ERROR: failed to create venv. Install Python 3.10+ and ensure the launcher 'py' is available.
      exit /b 1
    )
  )
)

echo Installing/Updating package...
"%PY%" -m pip install -U pip 1>nul
"%PY%" -m pip install .
if errorlevel 1 (
  echo ERROR: pip install failed.
  exit /b 1
)

echo Starting AutoCAD MCP server (stdio)...
"%PY%" -m acad_cmd.server

endlocal

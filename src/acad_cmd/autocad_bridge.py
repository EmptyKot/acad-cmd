import os
import re
import subprocess
import time
import uuid
from dataclasses import dataclass
from typing import Any, Optional, Tuple

import pythoncom
import pywintypes
import win32com.client
import win32process

try:
    import winreg  # type: ignore
except Exception:  # pragma: no cover
    winreg = None  # type: ignore


RPC_E_CALL_REJECTED = -2147418111


def _com_init() -> None:
    """Initialize COM for the current thread.

    FastMCP tool calls may run on a thread pool. In pywin32, each thread that
    touches COM must call CoInitialize() (it's safe to call multiple times).
    Missing initialization can lead to hangs/crashes when automating AutoCAD.
    """

    try:
        pythoncom.CoInitialize()
    except Exception:
        # Best-effort: if COM is already initialized (or cannot be), proceed.
        pass


_ACADVER_RE = re.compile(r"(?P<major>\d+)\.(?P<minor>\d+)")


def _parse_acadver_major(v: Any) -> Optional[int]:
    """Extract major version from AutoCAD ACADVER value.

    Examples: '23.1s (LMS Tech)' -> 23, '24.0' -> 24.
    """

    if v is None:
        return None
    s = str(v).strip().strip('"').strip("'")
    m = _ACADVER_RE.search(s)
    if not m:
        return None
    try:
        return int(m.group("major"))
    except Exception:
        return None


def _get_target_major() -> Optional[int]:
    # AutoCAD 2021 corresponds to major version 24.*
    raw = (os.environ.get("AUTOCAD_MCP_TARGET_MAJOR") or "").strip()
    if not raw:
        return None
    try:
        return int(raw)
    except Exception:
        return None


def _allow_new_instance(target_major: Optional[int]) -> bool:
    # If explicitly configured, obey it.
    raw = (os.environ.get("AUTOCAD_MCP_ALLOW_NEW_INSTANCE") or "").strip().lower()
    if raw:
        return raw in ("1", "true", "yes")

    # Default: allow launching a new automation-enabled AutoCAD instance.
    # Many AutoCAD installs don't expose a running instance via GetActiveObject(),
    # so attach-only defaults are fragile.
    return True


def _tasklist_pids(image_name: str) -> Tuple[int, ...]:
    """Return process IDs for a given image name (best-effort)."""

    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/FO", "CSV", "/NH"],
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception:
        return ()

    pids: list[int] = []
    for line in out.splitlines():
        line = line.strip()
        if not line or line.startswith("INFO:"):
            continue
        # CSV: "Image Name","PID",...
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) < 2:
            continue
        try:
            pids.append(int(parts[1]))
        except Exception:
            continue
    return tuple(pids)


def _get_hwnd_pid(hwnd: Any) -> Optional[int]:
    try:
        h = int(hwnd)
        if h <= 0:
            return None
        _tid, pid = win32process.GetWindowThreadProcessId(h)
        return int(pid)
    except Exception:
        return None


def _is_callee_busy(err: Exception) -> bool:
    if isinstance(err, pywintypes.com_error):
        hr_attr = getattr(err, "hresult", None)
        if hr_attr is not None:
            try:
                return int(hr_attr) == RPC_E_CALL_REJECTED
            except Exception:
                pass
        # Sometimes hresult is in args[0]
        try:
            hr = int(err.args[0])
            return hr == RPC_E_CALL_REJECTED
        except Exception:
            return False
    return False


def com_retry(fn, *, retries: int = 15, base_delay: float = 0.05, max_delay: float = 0.8):
    _com_init()
    delay = base_delay
    last = None
    for _ in range(retries):
        try:
            return fn()
        except Exception as e:
            last = e
            if not _is_callee_busy(e):
                raise
            time.sleep(delay)
            delay = min(max_delay, delay * 1.6)
    if last:
        raise last


@dataclass
class WaitResult:
    completed: bool
    needs_input: bool
    quiescent: bool


class AutoCADBridge:
    def __init__(self) -> None:
        self._acad = None
        self._doc = None
        self._connected = False

    def _get_acad_progids(self) -> Tuple[str, ...]:
        """Return ProgIDs to try, in preferred order.

        Key behavior:
        - Prefer versioned ProgIDs (AutoCAD.Application.XX) from newest to oldest.
        - Only try the unversioned ProgID last.

        Rationale: CurVer/unversioned ProgID can be hijacked by Civil 3D installs
        and may attach to an older product even if a newer AutoCAD is running.
        """

        target_major = _get_target_major()

        progids: list[str] = []

        # If user pinned a version, try it first.
        if target_major:
            progids.append(f"AutoCAD.Application.{target_major}")

        # Try a small range of recent versions.
        # AutoCAD 2020..2026 typically maps to Application.23..29.
        for v in range(30, 18, -1):
            p = f"AutoCAD.Application.{v}"
            if p not in progids:
                progids.append(p)

        # CurVer (optional) - some setups only register this.
        prefer_curver = (os.environ.get("AUTOCAD_MCP_PREFER_CURVER") or "").strip().lower() in ("1", "true", "yes")
        if prefer_curver and winreg is not None:
            for root, key_path in (
                (winreg.HKEY_CLASSES_ROOT, r"AutoCAD.Application\\CurVer"),
                (winreg.HKEY_CURRENT_USER, r"Software\\Classes\\AutoCAD.Application\\CurVer"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\Classes\\AutoCAD.Application\\CurVer"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\Classes\\Wow6432Node\\AutoCAD.Application\\CurVer"),
            ):
                try:
                    with winreg.OpenKey(root, key_path) as k:
                        v, _ = winreg.QueryValueEx(k, "")
                        if isinstance(v, str) and v.strip() and v.strip() not in progids:
                            progids.append(v.strip())
                            break
                except Exception:
                    continue

        # Unversioned last.
        progids.append("AutoCAD.Application")
        return tuple(progids)

    def connect(self, *, attach_or_launch: bool = True, visible: bool = True) -> bool:
        _com_init()

        target_major = _get_target_major()
        allow_new = _allow_new_instance(target_major)

        def _attach(progid: str):
            self._acad = win32com.client.GetActiveObject(progid)
            self._acad.Visible = bool(visible)
            doc = self._acad.ActiveDocument
            self._doc = doc
            _ = str(doc.Name)

            if target_major is not None:
                try:
                    acadver = com_retry(lambda: doc.GetVariable("ACADVER"))
                    major = _parse_acadver_major(acadver)
                    if major is None or major != target_major:
                        return False
                except Exception:
                    return False

            return True

        for progid in self._get_acad_progids():
            try:
                ok = com_retry(lambda: _attach(progid))

                self._connected = bool(ok)
                if self._connected:
                    return True
            except Exception:
                continue

        # Fallback: some AutoCAD versions do not register an active object in the ROT
        # when launched normally. In that case, GetActiveObject() fails even though
        # AutoCAD is running. Dispatch() can attach to the running instance OR start
        # a new automation-enabled instance.
        if attach_or_launch:
            use_dispatch = target_major is not None or (os.environ.get("AUTOCAD_MCP_USE_DISPATCH") or "").strip().lower() in (
                "1",
                "true",
                "yes",
            )
            if use_dispatch:
                before = set(_tasklist_pids("acad.exe"))
                for progid in self._get_acad_progids():
                    # Avoid unversioned Dispatch() by default (may start the wrong product).
                    if "." not in progid:
                        continue
                    try:
                        _com_init()
                        self._acad = win32com.client.Dispatch(progid)
                        self._acad.Visible = bool(visible)
                        spawned_pid = _get_hwnd_pid(getattr(self._acad, "HWND", None))
                        after = set(_tasklist_pids("acad.exe"))

                        # If Dispatch spawned a new acad.exe and user doesn't allow it,
                        # immediately close it and keep searching.
                        if not allow_new and spawned_pid is not None and spawned_pid not in before and len(after) > len(before):
                            try:
                                self._acad.Quit()
                            except Exception:
                                pass
                            self._acad = None
                            self._doc = None
                            self._connected = False
                            continue

                        doc = self._acad.ActiveDocument
                        self._doc = doc
                        _ = str(doc.Name)

                        if target_major is not None:
                            acadver = com_retry(lambda: doc.GetVariable("ACADVER"))
                            major = _parse_acadver_major(acadver)
                            if major is None or major != target_major:
                                continue

                        self._connected = True
                        return True
                    except Exception:
                        continue

        # Optional launch: avoid win32com Dispatch() here.
        # In practice, COM activation via Dispatch can hang indefinitely and may
        # destabilize AutoCAD (and the host process). For reliability, only
        # support launching AutoCAD via an explicit executable path, then attach.
        if attach_or_launch:
            acad_exe = (os.environ.get("AUTOCAD_MCP_ACAD_EXE") or "").strip().strip('"')
            if acad_exe and os.path.exists(acad_exe):
                try:
                    extra = (os.environ.get("AUTOCAD_MCP_ACAD_ARGS") or "").strip()
                    args = [acad_exe]
                    if extra:
                        try:
                            import shlex

                            args.extend(shlex.split(extra, posix=False))
                        except Exception:
                            # Last resort: split on whitespace
                            args.extend([p for p in extra.split() if p.strip()])
                    subprocess.Popen(args, close_fds=True)
                except Exception:
                    pass

                # Give AutoCAD time to start and register in ROT, then retry attach.
                try:
                    wait_sec = float((os.environ.get("AUTOCAD_MCP_LAUNCH_WAIT_SEC") or "30").strip())
                except Exception:
                    wait_sec = 30.0

                t0 = time.time()
                while time.time() - t0 < wait_sec:
                    for progid in self._get_acad_progids():
                        try:
                            ok = com_retry(lambda: _attach(progid))
                            self._connected = bool(ok)
                            if self._connected:
                                return True
                        except Exception:
                            continue
                    time.sleep(0.5)

        self._connected = False
        return False

    def ensure_connection(self) -> bool:
        _com_init()
        if not self._connected or self._acad is None or self._doc is None:
            return self.connect(attach_or_launch=True)
        try:
            _ = str(self._doc.Name)
            return True
        except Exception:
            self._connected = False
            return self.connect(attach_or_launch=True)

    @property
    def acad(self) -> Any:
        _com_init()
        if not self.ensure_connection():
            raise RuntimeError("Not connected to AutoCAD")
        return self._acad

    @property
    def doc(self) -> Any:
        _com_init()
        if not self.ensure_connection():
            raise RuntimeError("Not connected to AutoCAD")
        return self._doc

    def get_dwg_label(self) -> Optional[str]:
        _com_init()
        try:
            name = str(self.doc.Name)
            path = str(self.doc.Path) if getattr(self.doc, "Path", None) else ""
            if path:
                return os.path.join(path, name)
            return name
        except Exception:
            return None

    def get_variable(self, name: str) -> Any:
        _com_init()
        def _op():
            return self.doc.GetVariable(name)
        return com_retry(_op)

    def set_variable(self, name: str, value: Any) -> None:
        _com_init()
        def _op():
            self.doc.SetVariable(name, value)
        com_retry(_op)

    def send_command(self, command: str) -> str:
        _com_init()
        cmd = command
        if not cmd.endswith("\n"):
            cmd += "\n"
        command_id = str(uuid.uuid4())

        def _op():
            self.doc.SendCommand(cmd)
            return True

        com_retry(_op)
        return command_id

    def wait_for_idle(self, timeout_sec: float, poll_interval_sec: float = 0.1) -> WaitResult:
        _com_init()
        t0 = time.time()
        last_quiescent = False

        while True:
            try:
                state = self.acad.GetAcadState()
                is_quiescent = bool(state.IsQuiescent)
            except Exception:
                is_quiescent = False

            try:
                cmdactive = int(self.get_variable("CMDACTIVE"))
            except Exception:
                cmdactive = 999

            last_quiescent = is_quiescent
            if is_quiescent and cmdactive == 0:
                return WaitResult(completed=True, needs_input=False, quiescent=True)

            if time.time() - t0 >= timeout_sec:
                # Not idle. Likely waiting for input or long running.
                needs_input = cmdactive != 0
                return WaitResult(completed=False, needs_input=needs_input, quiescent=is_quiescent)

            time.sleep(poll_interval_sec)

    def get_last_prompt(self) -> str:
        _com_init()
        try:
            v = self.get_variable("LASTPROMPT")
            return str(v) if v is not None else ""
        except Exception:
            return ""

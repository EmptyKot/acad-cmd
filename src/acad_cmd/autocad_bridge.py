import os
import re
import subprocess
import time
import uuid
from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple

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


def _normalize_fs_path(path: str) -> str:
    return os.path.normcase(os.path.abspath(os.path.normpath(path)))


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


def _allow_new_instance() -> bool:
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
        self._locked_major: Optional[int] = None
        self._bound_progid: Optional[str] = None
        self._last_status_cache: Dict[str, Any] = {}
        self._last_error_class: Optional[str] = None
        self._last_error_message: Optional[str] = None

    def _set_error(self, error_class: str, message: str) -> None:
        self._last_error_class = str(error_class)
        self._last_error_message = str(message)

    def _clear_error(self) -> None:
        self._last_error_class = None
        self._last_error_message = None

    def _desired_major(self) -> Optional[int]:
        return _get_target_major() or self._locked_major

    def _remember_binding(self, progid: str, major: Optional[int]) -> None:
        self._bound_progid = progid
        if major is not None:
            self._locked_major = int(major)

    def _get_variable_direct(self, name: str) -> Any:
        if self._doc is None:
            raise RuntimeError("Not connected to AutoCAD")

        def _op():
            return self._doc.GetVariable(name)

        return com_retry(_op)

    def _read_dwg_label_direct(self) -> Optional[str]:
        if self._doc is None:
            return None
        return self._doc_dwg_label(self._doc)

    def _doc_dwg_label(self, doc: Any) -> Optional[str]:
        if doc is None:
            return None
        name = str(com_retry(lambda: doc.Name))
        path = str(com_retry(lambda: doc.Path)) if getattr(doc, "Path", None) else ""
        if path:
            return os.path.join(path, name)
        return name

    def _safe_hwnd(self) -> Optional[int]:
        if self._acad is None:
            return None
        try:
            h = int(com_retry(lambda: getattr(self._acad, "HWND", 0)) or 0)
        except Exception:
            return None
        return h if h > 0 else None

    def _get_acad_progids(self) -> Tuple[str, ...]:
        """Return ProgIDs to try, in preferred order.

        Key behavior:
        - Prefer versioned ProgIDs (AutoCAD.Application.XX) from newest to oldest.
        - Only try the unversioned ProgID last.

        Rationale: CurVer/unversioned ProgID can be hijacked by Civil 3D installs
        and may attach to an older product even if a newer AutoCAD is running.
        """

        target_major = self._desired_major()

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

        target_major = self._desired_major()
        allow_new = _allow_new_instance()

        def _attach(progid: str):
            self._acad = win32com.client.GetActiveObject(progid)
            self._acad.Visible = bool(visible)
            doc = self._acad.ActiveDocument
            self._doc = doc
            _ = str(doc.Name)

            try:
                acadver = com_retry(lambda: doc.GetVariable("ACADVER"))
                major = _parse_acadver_major(acadver)
            except Exception:
                major = None

            if target_major is not None and major != target_major:
                self._acad = None
                self._doc = None
                return False

            self._remember_binding(progid, major)

            return True

        for progid in self._get_acad_progids():
            try:
                ok = com_retry(lambda: _attach(progid))

                self._connected = bool(ok)
                if self._connected:
                    self._clear_error()
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

                        try:
                            acadver = com_retry(lambda: doc.GetVariable("ACADVER"))
                            major = _parse_acadver_major(acadver)
                        except Exception:
                            major = None

                        if target_major is not None and major != target_major:
                            continue

                        self._remember_binding(progid, major)
                        self._connected = True
                        self._clear_error()
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
                                self._clear_error()
                                return True
                        except Exception:
                            continue
                    time.sleep(0.5)

        self._connected = False
        self._set_error("connect_failed", "Failed to connect to AutoCAD")
        return False

    def ensure_connection(self) -> bool:
        _com_init()
        if not self._connected or self._acad is None or self._doc is None:
            return self.connect(attach_or_launch=True)
        try:
            _ = com_retry(lambda: str(self._doc.Name), retries=3, base_delay=0.03, max_delay=0.2)
            self._clear_error()
            return True
        except Exception as e:
            if _is_callee_busy(e):
                self._set_error("busy", str(e))
                # AutoCAD can reject calls while UI thread is busy; treat as connected.
                return True
            self._connected = False
            self._set_error("disconnected", str(e))
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
        if not self.ensure_connection():
            return str(self._last_status_cache.get("dwg") or "") or None
        try:
            return self._read_dwg_label_direct()
        except Exception as e:
            if _is_callee_busy(e):
                cached = str(self._last_status_cache.get("dwg") or "")
                return cached or None
            return None

    def get_variable(self, name: str) -> Any:
        _com_init()
        if not self.ensure_connection():
            raise RuntimeError("Not connected to AutoCAD")
        return self._get_variable_direct(name)

    def set_variable(self, name: str, value: Any) -> None:
        _com_init()

        if not self.ensure_connection():
            raise RuntimeError("Not connected to AutoCAD")

        def _op():
            if self._doc is None:
                raise RuntimeError("Not connected to AutoCAD")
            self._doc.SetVariable(name, value)

        com_retry(_op)

    def send_command(self, command: str) -> str:
        _com_init()
        if not self.ensure_connection():
            raise RuntimeError("Not connected to AutoCAD")
        cmd = command
        if not cmd.endswith("\n"):
            cmd += "\n"
        command_id = str(uuid.uuid4())

        def _op():
            if self._doc is None:
                raise RuntimeError("Not connected to AutoCAD")
            self._doc.SendCommand(cmd)
            return True

        com_retry(_op)
        return command_id

    def open_drawing(
        self,
        path: str,
        *,
        timeout_sec: float = 30.0,
        poll_interval_sec: float = 0.2,
        read_only: bool = False,
    ) -> Dict[str, Any]:
        _com_init()
        if not path or not str(path).strip():
            raise ValueError("path must be non-empty")
        if timeout_sec <= 0:
            raise ValueError("timeout_sec must be > 0")
        if poll_interval_sec <= 0:
            raise ValueError("poll_interval_sec must be > 0")

        target_path = _normalize_fs_path(os.path.expandvars(os.path.expanduser(str(path))))
        if not os.path.isfile(target_path):
            raise FileNotFoundError(f"Drawing file not found: {target_path}")

        if not self.ensure_connection():
            raise RuntimeError("Not connected to AutoCAD")
        if self._acad is None:
            raise RuntimeError("Not connected to AutoCAD")

        existing = None
        already_open = False
        opened_now = False

        def _list_docs():
            if self._acad is None:
                return []
            return [doc for doc in self._acad.Documents]

        for doc in com_retry(_list_docs):
            try:
                label = self._doc_dwg_label(doc)
            except Exception:
                continue
            if label and _normalize_fs_path(label) == target_path:
                existing = doc
                already_open = True
                break

        if existing is None:
            def _open():
                if self._acad is None:
                    raise RuntimeError("Not connected to AutoCAD")
                docs = self._acad.Documents
                if read_only:
                    return docs.Open(target_path, True)
                return docs.Open(target_path)

            existing = com_retry(_open)
            opened_now = True

        def _activate():
            if existing is None:
                raise RuntimeError("Internal error: document handle missing")
            existing.Activate()
            return True

        com_retry(_activate)

        t0 = time.time()
        last_active = None
        activated = False

        while True:
            active_doc = com_retry(lambda: self._acad.ActiveDocument)
            self._doc = active_doc
            last_active = self._doc_dwg_label(active_doc)

            if last_active and _normalize_fs_path(last_active) == target_path:
                activated = True
                break

            if time.time() - t0 >= timeout_sec:
                break

            try:
                com_retry(_activate, retries=3, base_delay=0.02, max_delay=0.15)
            except Exception:
                pass
            time.sleep(poll_interval_sec)

        if not activated:
            msg = (
                f"Failed to activate opened drawing within {timeout_sec:.1f}s. "
                f"target='{target_path}', active='{last_active}'"
            )
            self._set_error("open_drawing_activate_failed", msg)
            raise RuntimeError(msg)

        self._clear_error()
        return {
            "path": target_path,
            "dwg": last_active,
            "already_open": already_open,
            "opened": opened_now,
            "activated": True,
            "read_only": bool(read_only),
        }

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

    def get_status_snapshot(self) -> Dict[str, Any]:
        _com_init()
        connected = self.ensure_connection()
        source = "live"
        stale = False
        busy = False
        error_class = self._last_error_class
        error_message = self._last_error_message

        if not connected:
            if self._last_status_cache:
                out = dict(self._last_status_cache)
                out.update(
                    {
                        "connected": False,
                        "busy": False,
                        "stale": True,
                        "source": "cache",
                        "error_class": error_class or "disconnected",
                        "error_message": error_message,
                        "locked_major": self._locked_major,
                        "bound_progid": self._bound_progid,
                    }
                )
                return out
            return {
                "connected": False,
                "busy": False,
                "stale": True,
                "source": "none",
                "error_class": error_class or "disconnected",
                "error_message": error_message,
                "dwg": None,
                "acadver": None,
                "acad_hwnd": None,
                "acad_pid": None,
                "cmdactive": None,
                "locked_major": self._locked_major,
                "bound_progid": self._bound_progid,
            }

        try:
            dwg = self._read_dwg_label_direct()
            acadver_val = self._get_variable_direct("ACADVER")
            acadver = str(acadver_val) if acadver_val is not None else None
            cmdactive_val = self._get_variable_direct("CMDACTIVE")
            try:
                cmdactive = int(cmdactive_val)
            except Exception:
                cmdactive = None

            hwnd = self._safe_hwnd()
            pid = _get_hwnd_pid(hwnd) if hwnd else None
            major = _parse_acadver_major(acadver)
            if self._locked_major is None and major is not None:
                self._locked_major = major

            out = {
                "connected": True,
                "busy": False,
                "stale": False,
                "source": "live",
                "error_class": None,
                "error_message": None,
                "dwg": dwg,
                "acadver": acadver,
                "acad_hwnd": hwnd,
                "acad_pid": pid,
                "cmdactive": cmdactive,
                "locked_major": self._locked_major,
                "bound_progid": self._bound_progid,
            }
            self._last_status_cache = dict(out)
            return out
        except Exception as e:
            busy = _is_callee_busy(e)
            error_class = "busy" if busy else "status_read_failed"
            source = "cache" if self._last_status_cache else "none"
            stale = True
            self._set_error(error_class, str(e))

            if self._last_status_cache:
                out = dict(self._last_status_cache)
                out.update(
                    {
                        "connected": True,
                        "busy": busy,
                        "stale": stale,
                        "source": source,
                        "error_class": error_class,
                        "error_message": str(e),
                        "locked_major": self._locked_major,
                        "bound_progid": self._bound_progid,
                    }
                )
                return out

            return {
                "connected": True,
                "busy": busy,
                "stale": stale,
                "source": source,
                "error_class": error_class,
                "error_message": str(e),
                "dwg": None,
                "acadver": None,
                "acad_hwnd": None,
                "acad_pid": None,
                "cmdactive": None,
                "locked_major": self._locked_major,
                "bound_progid": self._bound_progid,
            }

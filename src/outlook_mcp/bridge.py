"""Persistent STA thread for Outlook COM operations.

All COM calls run on a dedicated Single-Threaded Apartment (STA) thread
holding one ``Outlook.Application`` Dispatch handle. The MCP event loop
schedules work via :meth:`OutlookBridge.call` and awaits the result, so
the event loop never blocks on COM and COM never sees a non-STA thread.

Why one persistent thread instead of a fresh client per call:
* Reuses the same Dispatch — no per-call ``CoCreateInstance`` cost.
* Survives Outlook auto-launching (Dispatch will spawn ``OUTLOOK.EXE``
  if it isn't running, and the same connection serves every later call).
* Matches Outlook's strict STA threading model.
"""

from __future__ import annotations

import asyncio
import logging
import queue
import subprocess
import threading
import time
from concurrent.futures import Future
from dataclasses import dataclass
from typing import Any, Callable

from outlook_mcp.errors import OutlookError, WriteOutcomeUnknown, is_disconnect_error

logger = logging.getLogger("outlook_mcp.bridge")

# Cold-launching Outlook from a fully closed state can take 15–25s before
# the MAPI session answers `CurrentUser.Name`. The COM thread polls up to
# `_ATTACH_TIMEOUT_SEC`; the main thread waits slightly longer in start().
_ATTACH_TIMEOUT_SEC = 60
_READY_TIMEOUT_SEC = _ATTACH_TIMEOUT_SEC + 5
_CALL_TIMEOUT_SEC = 60


def _find_outlook_exe() -> str | None:
    """Look up OUTLOOK.EXE via the App Paths registry key."""
    try:
        import winreg
    except ImportError:
        return None
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(
                hive,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
            ) as key:
                path, _ = winreg.QueryValueEx(key, None)
                if path:
                    return path
        except OSError:
            continue
    return None


def _launch_outlook() -> bool:
    """Spawn OUTLOOK.EXE detached from this process. Returns True if started."""
    path = _find_outlook_exe()
    if not path:
        logger.warning("Could not locate OUTLOOK.EXE in registry App Paths")
        return False
    try:
        flags = 0
        # DETACHED_PROCESS so Outlook outlives this MCP process.
        flags |= getattr(subprocess, "DETACHED_PROCESS", 0)
        flags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        subprocess.Popen(  # noqa: S603 - path comes from HKLM registry
            [path],
            creationflags=flags,
            close_fds=True,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        logger.info("Launched classic Outlook")
        return True
    except OSError:
        logger.warning("Could not launch classic Outlook")
        return False


_CHECK_WINDOW = (
    "Open classic Outlook in your unlocked Windows session and check for a "
    "profile selection, sign-in or security dialog."
)


@dataclass
class _Call:
    func: Callable[..., Any]
    args: tuple
    kwargs: dict
    retry_on_disconnect: bool
    session: object
    deadline: float
    result: Future


class OutlookBridge:
    """Own the STA and expose only primitive readiness information to callers."""

    def __init__(self) -> None:
        self._thread: threading.Thread | None = None
        self._queue: queue.Queue[_Call | None] = queue.Queue()
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._state = "closed"
        # A fresh token invalidates work queued for an earlier Outlook session.
        self._session: object | None = None
        self._outlook: Any = None
        self._namespace: Any = None
        self._profile_name: str | None = None

    def start(self, *, wait: bool = True) -> None:
        """Start one STA. STDIO can serve diagnostics while Dispatch is blocked."""
        if self._thread is not None:
            raise OutlookError("The Outlook bridge has already been started.")
        self._state = "starting"
        self._thread = threading.Thread(target=self._run, daemon=True, name="outlook-com")
        self._thread.start()
        if wait:
            self._ready.wait(timeout=_READY_TIMEOUT_SEC)
            if self._state != "ready":
                raise self._not_ready(self._state)

    @staticmethod
    def _not_ready(state: str) -> OutlookError:
        return OutlookError(
            f"Outlook is {state}. The operation did not start. {_CHECK_WINDOW} "
            "Check readiness again before submitting the operation."
        )

    def _attach(self) -> bool:
        """Attach on the STA. Never log on to, select or switch a profile."""
        from win32com.client import dynamic

        self._session = self._namespace = self._outlook = self._profile_name = None
        deadline = time.monotonic() + _ATTACH_TIMEOUT_SEC
        launch_attempted = False
        while time.monotonic() < deadline and not self._shutdown.is_set():
            outlook = None
            try:
                # Late binding avoids thread-affine generated typelib wrappers.
                outlook = dynamic.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                # Dispatch alone can succeed before the profile is usable.
                namespace.CurrentUser.Name
            except Exception:
                if outlook is None and not launch_attempted:
                    logger.info("Outlook is closed or still starting. Checking the interactive session.")
                    _launch_outlook()
                    launch_attempted = True
                self._shutdown.wait(1.5)
                continue
            try:
                name = namespace.CurrentProfileName
                self._profile_name = name if isinstance(name, str) and name else None
            except Exception:
                self._profile_name = None
            self._outlook, self._namespace = outlook, namespace
            self._session = object()
            self._state = "ready"
            logger.info("Outlook session ready")
            return True
        self._state = "unavailable"
        logger.warning("Outlook is unavailable. %s", _CHECK_WINDOW)
        return False

    def _execute(self, job: _Call) -> Any:
        if job.session is not self._session or time.monotonic() >= job.deadline:
            raise OutlookError("Outlook session changed or the request expired. The operation did not start.")
        try:
            return job.func(self._outlook, self._namespace, *job.args, **job.kwargs)
        except Exception as exc:
            disconnected = is_disconnect_error(exc) or self._connection_lost()
            if disconnected:
                self._state = "reconnecting"
                logger.warning("Outlook disconnected. Reconnecting to the interactive session.")
            if not job.retry_on_disconnect:
                if isinstance(exc, WriteOutcomeUnknown):
                    raise
                if disconnected or not isinstance(exc, OutlookError):
                    raise WriteOutcomeUnknown() from None
                raise
            if not disconnected:
                raise

        previous_profile = self._profile_name
        if not self._attach():
            raise OutlookError(f"Outlook was closed and could not be relaunched. {_CHECK_WINDOW}")
        if not previous_profile or previous_profile != self._profile_name:
            raise OutlookError("Cannot confirm the same Outlook profile. Check the active profile before retrying.")
        if time.monotonic() >= job.deadline or self._shutdown.is_set():
            raise OutlookError("Outlook reconnected after the read expired. Request the read again.")
        try:
            return job.func(self._outlook, self._namespace, *job.args, **job.kwargs)
        except Exception as exc:
            if is_disconnect_error(exc) or self._connection_lost():
                self._state = "reconnecting"
                raise OutlookError(f"Outlook disconnected again during the read. {_CHECK_WINDOW}") from None
            raise

    def _connection_lost(self) -> bool:
        """Probe the cached namespace proxy for a dead Outlook. COM thread only.

        A stale proxy doesn't always surface as a disconnect HRESULT:
        pywin32's dynamic ``__getattr__`` swallows the ``com_error`` from
        ``GetIDsOfNames`` and raises a bare ``AttributeError`` such as
        ``GetNamespace.GetDefaultFolder`` with no cause chain. So after any
        failure, ask the raw ``IDispatch`` directly — it raises the real
        RPC error if Outlook was closed or restarted under us.
        """
        if self._namespace is None:
            return True
        try:
            self._namespace._oleobj_.GetIDsOfNames("CurrentUser")
        except BaseException as exc:  # noqa: BLE001
            return is_disconnect_error(exc)
        return False

    def _run(self) -> None:
        try:
            import pythoncom

            pythoncom.CoInitializeEx(
                pythoncom.COINIT_APARTMENTTHREADED | pythoncom.COINIT_DISABLE_OLE1DDE
            )
        except Exception:
            self._state = "unavailable"
            self._ready.set()
            logger.error("Cannot initialize Outlook COM. Check the Windows installation.")
            return
        try:
            self._attach()
            self._ready.set()
            while not self._shutdown.is_set():
                job = self._queue.get()
                if self._shutdown.is_set():
                    break
                if job is None:
                    # A request observed a failed attach and asked us to try again.
                    if self._state == "reconnecting":
                        self._attach()
                    continue
                if not job.result.set_running_or_notify_cancel():
                    continue
                try:
                    value = self._execute(job)
                except Exception as exc:
                    job.result.set_exception(exc)
                else:
                    job.result.set_result(value)
                if self._state == "reconnecting" and self._session is not None:
                    self._attach()
        finally:
            self._outlook = self._namespace = self._session = None
            self._state = "closed"
            self._ready.set()
            pythoncom.CoUninitialize()

    async def call(self, func: Callable[..., Any], *args: Any,
                   retry_on_disconnect: bool = True, **kwargs: Any) -> Any:
        """Bound the wait and atomically cancel queued work before it can start."""
        session = self._session
        state = self._state
        if state != "ready" or session is None or self._shutdown.is_set():
            if state == "unavailable" and self._thread and self._thread.is_alive():
                # All MCP calls enter here on the event loop. Coalesce recovery
                # requests before waking the STA. Never enqueue the operation.
                self._state = "reconnecting"
                self._queue.put(None)
            raise self._not_ready(state)

        job = _Call(func, args, kwargs, retry_on_disconnect, session,
                    time.monotonic() + _CALL_TIMEOUT_SEC, Future())
        result = asyncio.wrap_future(job.result)
        self._queue.put(job)
        try:
            done, _ = await asyncio.wait({result}, timeout=_CALL_TIMEOUT_SEC)
            if done:
                return result.result()
            if job.result.cancel():
                raise OutlookError(f"Outlook timed out in the queue. The operation did not start. {_CHECK_WINDOW}")
            if job.result.done():
                return job.result.result()
            if not retry_on_disconnect:
                raise WriteOutcomeUnknown(f"Outlook timed out during the operation. {_CHECK_WINDOW}")
            raise OutlookError(f"The Outlook read timed out. {_CHECK_WINDOW}")
        finally:
            # Also withdraw a queued request when its MCP client cancels.
            job.result.cancel()
            result.cancel()

    def stop(self) -> None:
        self._shutdown.set()
        self._queue.put(None)
        if self._thread is not None:
            self._thread.join(timeout=5)

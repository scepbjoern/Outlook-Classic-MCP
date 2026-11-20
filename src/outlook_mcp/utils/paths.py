"""Validate user-supplied paths and sender-supplied attachment names."""

from __future__ import annotations

import os
from pathlib import Path, PureWindowsPath

from outlook_mcp.errors import OutlookError

_DEVICE_NAMES = {"CON", "PRN", "AUX", "NUL", "CLOCK$", "CONIN$", "CONOUT$"} | {
    prefix + digit for prefix in ("COM", "LPT") for digit in "123456789¹²³"
}


def validate_filename(name: str) -> str:
    """Accept a single filename that Windows can represent unambiguously."""
    if (
        not name
        or name.endswith((" ", "."))
        or any(ord(c) < 32 or c in '<>:"/\\|?*' for c in name)
        or name.lstrip(".").partition(".")[0].rstrip(" ").upper() in _DEVICE_NAMES
    ):
        raise OutlookError("Invalid or reserved Windows filename.")
    return name


def _validated_path(raw: str) -> Path:
    if not raw:
        raise OutlookError("Path is empty.")
    allow_any = os.environ.get("OUTLOOK_MCP_ALLOW_ANY_PATH", "").strip() in ("1", "true", "yes")
    windows = PureWindowsPath(raw)
    spelling = raw.replace("/", "\\")
    if spelling.startswith(("\\\\?\\", "\\\\.\\")):
        raise OutlookError("Windows device paths are not supported.")
    if spelling.startswith("\\\\") and not allow_any:
        raise OutlookError("Network paths require OUTLOOK_MCP_ALLOW_ANY_PATH=1.")
    path = Path(raw)
    if not path.is_absolute():
        raise OutlookError("Path must be absolute.")
    for part in windows.parts:
        if part != windows.anchor:
            validate_filename(part)
    # Check the lexical boundary before touching the supplied path, then check
    # the actual destination after resolving links and Windows junctions.
    if not allow_any:
        profile = Path(os.environ.get("USERPROFILE") or Path.home())
        if not path.is_relative_to(profile):
            raise OutlookError("Path is outside the user profile.")
    try:
        target = path.resolve()
        if not allow_any and not target.is_relative_to(profile.resolve()):
            raise OutlookError("Path is outside the user profile.")
    except (OSError, RuntimeError) as exc:
        raise OutlookError("Could not resolve the filesystem path.") from exc
    return target


def validate_attachment_path(raw: str) -> str:
    """Return the validated absolute path of an existing file."""
    target = _validated_path(raw)
    if not target.is_file():
        raise OutlookError("Attachment not found. Provide an existing file.")
    return str(target)


def validate_output_dir(raw: str) -> str:
    """Validate the destination before creating any directories."""
    target = _validated_path(raw)
    target.mkdir(parents=True, exist_ok=True)
    return str(target)

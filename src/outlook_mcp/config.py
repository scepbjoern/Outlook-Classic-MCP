"""Local tool and mailbox policy for the Outlook MCP server.

This module intentionally has no Outlook, MCP or third-party dependencies so a
broken configuration can reject startup before Outlook COM is initialized.
"""

from __future__ import annotations

import json
import os
import re
import stat
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

# The explicit inventory also makes an unreviewed new MCP tool fail closed:
# adding a tool decorator requires updating its group's inventory.
TOOL_GROUPS: dict[str, tuple[str, ...]] = {
    "mail": (
        "outlook_list_mails", "outlook_search_mails", "outlook_get_mail",
        "outlook_send_mail", "outlook_reply_mail", "outlook_forward_mail",
        "outlook_move_mail", "outlook_delete_mail", "outlook_mark_mail",
        "outlook_save_attachments",
    ),
    "folders": ("outlook_list_folders", "outlook_create_folder"),
    "calendar": (
        "outlook_list_events", "outlook_get_event", "outlook_create_event",
        "outlook_update_event", "outlook_delete_event", "outlook_respond_event",
    ),
    "contacts": (
        "outlook_list_contacts", "outlook_search_contacts",
        "outlook_resolve_name", "outlook_get_contact",
    ),
    "tasks": ("outlook_list_tasks", "outlook_create_task", "outlook_complete_task"),
    "categories": ("outlook_list_categories", "outlook_set_category"),
    "rules": ("outlook_list_rules", "outlook_toggle_rule"),
    "out_of_office": ("outlook_get_out_of_office",),
    "account": ("outlook_whoami",),
}
ALL_TOOLS = frozenset(name for names in TOOL_GROUPS.values() for name in names)
MAX_CONFIG_BYTES = 1024 * 1024


class ToolConfigError(ValueError):
    """Tool policy is invalid or could not be read; refuse to start."""


def normalize_identifier(value: str) -> str:
    """Fold ASCII case only. Distinct international mailbox names stay distinct."""
    return value.translate(_ASCII_LOWER)


_ASCII_LOWER = str.maketrans("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz")
_SMTP_ATOM = r"[\w!#$%&'*+/=?^`{|}~-]+"
_DOMAIN_LABEL = r"[^\W_](?:[^\W_]|-(?=[^\W_]|-))*"
_SMTP_ADDRESS = re.compile(rf"{_SMTP_ATOM}(?:\.{_SMTP_ATOM})*@{_DOMAIN_LABEL}(?:\.{_DOMAIN_LABEL})*")


def is_smtp_address(value: str) -> bool:
    """Accept a bare dot-atom SMTP address instead of Outlook display-name syntax."""
    return _SMTP_ADDRESS.fullmatch(value) is not None


@dataclass(frozen=True)
class AccountPolicy:
    read: frozenset[str] = field(default_factory=frozenset)
    write: frozenset[str] = field(default_factory=frozenset)


@dataclass(frozen=True)
class ToolPolicy:
    """Allowlist constraints intersect; deny rules always take precedence.

    ``None`` means no allowlist restriction; an explicit empty allowlist
    means no tools/groups are allowed.
    """

    enabled_groups: frozenset[str] | None = None
    disabled_groups: frozenset[str] = field(default_factory=frozenset)
    enabled_tools: frozenset[str] | None = None
    disabled_tools: frozenset[str] = field(default_factory=frozenset)
    accounts: AccountPolicy | None = None

    def allows(self, group: str, name: str) -> bool:
        if group not in TOOL_GROUPS or name not in TOOL_GROUPS[group]:
            raise ToolConfigError(f"Unrecognised registered tool {name!r} in group {group!r}")
        return (
            group not in self.disabled_groups
            and name not in self.disabled_tools
            and (self.enabled_groups is None or group in self.enabled_groups)
            and (self.enabled_tools is None or name in self.enabled_tools)
        )


def default_config_path() -> Path:
    """Platform-specific user config path (never relative to the repository)."""
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or (Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config"))
    return base / "outlook-classic-mcp" / "config.json"


def _unique_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ToolConfigError(f"Duplicate JSON key: {key!r}")
        result[key] = value
    return result


def _object(value: Any, label: str, allowed: set[str]) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ToolConfigError(f"{label} must be a JSON object")
    unknown = set(value) - allowed
    if unknown:
        raise ToolConfigError(f"Unknown {label} field(s): {', '.join(sorted(unknown))}")
    return value


def _names(obj: dict[str, Any], key: str, valid: frozenset[str] | set[str]) -> frozenset[str] | None:
    if key not in obj:
        return None
    values = obj[key]
    if not isinstance(values, list) or any(not isinstance(name, str) for name in values):
        raise ToolConfigError(f"tools.{key} must be an array of names")
    if len(values) != len(set(values)):
        raise ToolConfigError(f"tools.{key} contains duplicate names")
    invalid = set(values) - valid
    if invalid:
        raise ToolConfigError(f"Unknown name(s) in tools.{key}: {', '.join(sorted(invalid))}")
    return frozenset(values)


def _smtp_addresses(obj: dict[str, Any], key: str) -> frozenset[str]:
    values = obj.get(key, [])
    if not isinstance(values, list) or any(
        not isinstance(value, str) or not is_smtp_address(value) for value in values
    ):
        raise ToolConfigError(f"accounts.{key} must be an array of exact SMTP addresses")
    normalized = frozenset(normalize_identifier(value) for value in values)
    if len(normalized) != len(values):
        raise ToolConfigError(f"accounts.{key} contains duplicate SMTP addresses")
    return normalized


def parse_tool_policy(raw: str) -> ToolPolicy:
    """Parse and validate the complete configuration, rejecting unknown fields."""
    try:
        data = json.loads(raw, object_pairs_hook=_unique_keys)
    except ToolConfigError:
        raise
    except (ValueError, RecursionError) as exc:
        raise ToolConfigError(f"Invalid JSON configuration: {exc}") from exc
    root = _object(data, "config", {"version", "tools", "accounts"})
    version = root.get("version", 1)
    if type(version) is not int or version not in (1, 2):
        raise ToolConfigError("config.version must be 1 or 2")
    account_policy = None
    if version == 2:
        accounts = _object(root.get("accounts"), "accounts", {"read", "write"})
        account_policy = AccountPolicy(_smtp_addresses(accounts, "read"), _smtp_addresses(accounts, "write"))
    elif "accounts" in root:
        raise ToolConfigError("accounts requires config.version 2")
    tools = _object(
        root.get("tools", {}),
        "tools",
        {"enabled_groups", "disabled_groups", "enabled_tools", "disabled_tools"},
    )
    groups = frozenset(TOOL_GROUPS)
    enabled_groups = _names(tools, "enabled_groups", groups)
    disabled_groups = _names(tools, "disabled_groups", groups)
    enabled_tools = _names(tools, "enabled_tools", ALL_TOOLS)
    disabled_tools = _names(tools, "disabled_tools", ALL_TOOLS)
    return ToolPolicy(
        enabled_groups=enabled_groups,
        disabled_groups=disabled_groups or frozenset(),
        enabled_tools=enabled_tools,
        disabled_tools=disabled_tools or frozenset(),
        accounts=account_policy,
    )


def _missing_default(path: Path) -> bool:
    """Accept missing components only below a usable directory.

    exists()/is_symlink() can hide filesystem errors. lstat() also
    distinguishes a dangling link from a genuinely absent policy.
    """
    for candidate in (path, *path.parents):
        try:
            candidate.lstat()
        except FileNotFoundError:
            continue
        if candidate == path:
            return False
        if not stat.S_ISDIR(candidate.stat().st_mode):
            raise ToolConfigError(f"Config parent is not a directory: {candidate}")
        return True
    raise ToolConfigError(f"Cannot locate a directory for tool config at {path}")


def load_tool_policy() -> ToolPolicy:
    """Load optional user config or fail closed on explicit/broken config.

    A missing *default* file means backwards-compatible unrestricted tools.
    A missing explicit OUTLOOK_MCP_CONFIG, invalid contents or I/O errors
    always abort startup rather than silently enabling more tools.
    """
    override = os.environ.get("OUTLOOK_MCP_CONFIG")
    if override is not None and not override.strip():
        raise ToolConfigError("OUTLOOK_MCP_CONFIG cannot be empty")
    try:
        path = Path(override).expanduser() if override is not None else default_config_path()
    except RuntimeError as exc:
        raise ToolConfigError(f"Cannot resolve tool config path: {exc}") from exc
    try:
        if override is None and _missing_default(path):
            return ToolPolicy()
        with path.open("rb") as stream:
            raw = stream.read(MAX_CONFIG_BYTES + 1)
    except OSError as exc:
        raise ToolConfigError(f"Cannot read tool config at {path}: {exc.strerror or exc}") from exc
    if len(raw) > MAX_CONFIG_BYTES:
        raise ToolConfigError(f"Tool config at {path} exceeds {MAX_CONFIG_BYTES} bytes")
    try:
        content = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ToolConfigError(f"Tool config at {path} is not UTF-8") from exc
    return parse_tool_policy(content)

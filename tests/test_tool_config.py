"""Tool config: strict startup validation and actual FastMCP registration.

Only synthetic configuration and tool metadata are used; Outlook is never started.
"""

from __future__ import annotations

import asyncio
import json

import pytest

from outlook_mcp.config import (
    ALL_TOOLS,
    TOOL_GROUPS,
    ToolConfigError,
    ToolPolicy,
    load_tool_policy,
    parse_tool_policy,
)
from outlook_mcp.server import build_server


@pytest.fixture(autouse=True)
def isolated_config(monkeypatch, tmp_path):
    # Never inspect a developer's actual default Windows/Unix config file.
    monkeypatch.delenv("OUTLOOK_MCP_CONFIG", raising=False)
    monkeypatch.setattr(
        "outlook_mcp.config.default_config_path",
        lambda: tmp_path / "missing" / "config.json",
    )


def _tool_names(policy=None):
    mcp, _bridge = build_server(policy=policy)
    return {tool.name for tool in asyncio.run(mcp.list_tools())}


def test_no_config_preserves_all_original_tools():
    assert len(ALL_TOOLS) == 31
    assert _tool_names() == ALL_TOOLS


def test_disable_single_mail_tool_leaves_other_mail_tools():
    policy = parse_tool_policy('{"tools":{"disabled_tools":["outlook_delete_mail"]}}')
    names = _tool_names(policy)
    assert names == ALL_TOOLS - {"outlook_delete_mail"}
    assert set(TOOL_GROUPS["mail"]) - {"outlook_delete_mail"} <= names


def test_disabled_groups_and_allowlists():
    policy = parse_tool_policy('{"tools":{"disabled_groups":["mail", "rules"]}}')
    assert _tool_names(policy) == ALL_TOOLS - set(TOOL_GROUPS["mail"]) - set(TOOL_GROUPS["rules"])

    only_calendar = parse_tool_policy('{"tools":{"enabled_groups":["calendar"]}}')
    assert _tool_names(only_calendar) == set(TOOL_GROUPS["calendar"])

    only_one = parse_tool_policy('{"tools":{"enabled_tools":["outlook_list_events"]}}')
    assert _tool_names(only_one) == {"outlook_list_events"}


def test_deny_wins_and_allowlists_intersect():
    policy = parse_tool_policy(json.dumps({"tools": {
        "enabled_groups": ["mail", "calendar"],
        "disabled_groups": ["mail"],
        "enabled_tools": ["outlook_delete_mail", "outlook_list_events"],
        "disabled_tools": ["outlook_delete_mail"],
    }}))
    assert _tool_names(policy) == {"outlook_list_events"}
    assert _tool_names(parse_tool_policy('{"tools":{"enabled_groups":[]}}')) == set()
    assert _tool_names(parse_tool_policy('{"tools":{"enabled_tools":[]}}')) == set()


@pytest.mark.parametrize("rules", [
    {"enabled_tools": ["outlook_delete_mail"], "disabled_groups": ["mail"]},
    {"enabled_groups": ["mail"], "enabled_tools": ["outlook_list_events"]},
    {"enabled_tools": ["outlook_delete_mail"], "disabled_tools": ["outlook_delete_mail"]},
])
def test_conflicting_rules_leave_no_tools(rules):
    assert _tool_names(parse_tool_policy(json.dumps({"tools": rules}))) == set()


@pytest.mark.parametrize("raw", [
    "not json", "[]", "null", "{", '{"version":true}', '{"version":2}',
    '{"tools":null}', '{"unknown":true}', '{"tools":{"disabled_tool":[]}}',
    '{"tools":{"disabled_groups":["not_a_group"]}}',
    '{"tools":{"disabled_tools":["outlook_not_real"]}}',
    '{"tools":{"enabled_tools":"outlook_list_mails"}}',
    '{"tools":{"enabled_groups":[null]}}',
    '{"tools":{"disabled_tools":["outlook_delete_mail","outlook_delete_mail"]}}',
    '{"tools":{"disabled_tools":["outlook_delete_mail"],"disabled_tools":[]}}',
    '{"tools":{},"tools":{"disabled_tools":[]}}',
])
def test_malformed_or_unknown_config_fails_closed(raw):
    with pytest.raises(ToolConfigError):
        parse_tool_policy(raw)


def test_missing_default_config_is_optional():
    assert load_tool_policy() == ToolPolicy()


def test_explicit_override_is_used(monkeypatch, tmp_path):
    path = tmp_path / "synthetic.json"
    path.write_text('{"tools":{"disabled_tools":["outlook_delete_mail"]}}', encoding="utf-8")
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG", str(path))
    assert "outlook_delete_mail" not in _tool_names(load_tool_policy())


def test_invalid_explicit_path_or_file_is_not_ignored(monkeypatch, tmp_path):
    path = tmp_path / "missing.json"
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG", str(path))
    with pytest.raises(ToolConfigError):
        load_tool_policy()
    path.write_text('{"tools":{"disabled_tool":["outlook_delete_mail"]}}', encoding="utf-8")
    with pytest.raises(ToolConfigError):
        load_tool_policy()
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG", " ")
    with pytest.raises(ToolConfigError):
        load_tool_policy()


def test_read_errors_and_size_limit_fail_closed(monkeypatch, tmp_path):
    path = tmp_path / "synthetic.json"
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG", str(path))
    path.write_bytes(b"\xff")
    with pytest.raises(ToolConfigError):
        load_tool_policy()
    path.write_bytes(b" " * (1024 * 1024 + 1))
    with pytest.raises(ToolConfigError):
        load_tool_policy()


def test_decorator_must_be_in_inventory():
    from outlook_mcp.tools import _PolicyToolRegistration

    class StubServer:
        def tool(self, **kwargs):
            return lambda fn: fn

    registrar = _PolicyToolRegistration(StubServer(), "mail", ToolPolicy())
    with pytest.raises(ToolConfigError):
        registrar.tool(name="outlook_unreviewed_tool")
    with pytest.raises(ToolConfigError):
        registrar.tool()


def test_invalid_config_prevents_server_construction(monkeypatch, tmp_path):
    path = tmp_path / "synthetic.json"
    path.write_text('{"tools":{"disabled_tool":["outlook_delete_mail"]}}', encoding="utf-8")
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG", str(path))
    with pytest.raises(ToolConfigError):
        build_server()


def test_malformed_default_config_is_not_ignored(monkeypatch):
    import outlook_mcp.config as config_module

    path = config_module.default_config_path()
    path.parent.mkdir(parents=True)
    path.write_text('{"tools":{"disabled_groups":["typo"]}}', encoding="utf-8")
    with pytest.raises(ToolConfigError):
        load_tool_policy()


def test_unusable_default_directory_fails_closed():
    from outlook_mcp.config import default_config_path

    default_config_path().parent.touch()
    with pytest.raises(ToolConfigError):
        load_tool_policy()


def test_default_path_permission_error_fails_closed(monkeypatch):
    from pathlib import Path

    def deny_access(path):
        raise PermissionError("synthetic access denied")

    monkeypatch.setattr(Path, "lstat", deny_access)
    with pytest.raises(ToolConfigError, match="Cannot read"):
        load_tool_policy()


def test_deeply_nested_non_object_config_is_rejected():
    # Some CPython decoders can parse this depth; the root schema still rejects it.
    with pytest.raises(ToolConfigError):
        parse_tool_policy("[" * 2000 + "]" * 2000)


def test_json_decoder_limits_are_config_errors(monkeypatch):
    def exceeded_limit(*args, **kwargs):
        raise RecursionError("synthetic decoder limit")

    monkeypatch.setattr(json, "loads", exceeded_limit)
    with pytest.raises(ToolConfigError, match="Invalid JSON") as error:
        parse_tool_policy("{}")
    assert isinstance(error.value.__cause__, RecursionError)

"""Exercise tool discovery, calls and config reload through a real MCP client."""

import asyncio
import json
import subprocess
import sys
from contextlib import asynccontextmanager
from datetime import timedelta
from pathlib import Path

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import get_default_environment, stdio_client

from outlook_mcp.config import ALL_TOOLS, TOOL_GROUPS

FIXTURE = Path(__file__).with_name("stdio_fixture.py")


@asynccontextmanager
async def session_at(tmp_path):
    params = StdioServerParameters(
        command=sys.executable,
        args=[str(FIXTURE), str(tmp_path / "bridge.log")],
        env={"APPDATA": str(tmp_path), "XDG_CONFIG_HOME": str(tmp_path)},
    )
    async with stdio_client(params) as streams:
        async with ClientSession(*streams, read_timeout_seconds=timedelta(seconds=20)) as session:
            await session.initialize()
            yield session


def test_discovery_and_calls_across_restarts(tmp_path):
    config = tmp_path / "outlook-classic-mcp" / "config.json"
    config.parent.mkdir()
    bridge_log = tmp_path / "bridge.log"
    delete = "outlook_delete_mail"

    async def exercise():
        async with session_at(tmp_path) as session:
            original = {tool.name: tool for tool in (await session.list_tools()).tools}
            assert len(original) == 31
            assert original.keys() == ALL_TOOLS
            resources = (await session.list_resources()).resources
            config.write_text(json.dumps({"tools": {"disabled_tools": [delete]}}), encoding="utf-8")
            assert {tool.name for tool in (await session.list_tools()).tools} == original.keys()

        async with session_at(tmp_path) as session:
            available = {tool.name: tool for tool in (await session.list_tools()).tools}
            assert available == {name: tool for name, tool in original.items() if name != delete}
            assert (await session.list_resources()).resources == resources
            uri = available["outlook_list_mails"].meta["ui"]["resourceUri"]
            assert "<!DOCTYPE html>" in (await session.read_resource(uri)).contents[0].text
            denied = await session.call_tool(delete, {"entry_id": "synthetic-id"})
            assert denied.isError and f"Unknown tool: {delete}" in denied.content[0].text
            assert bridge_log.read_text(encoding="utf-8").splitlines() == ["constructed"]
            allowed = await session.call_tool("outlook_whoami", {"response_format": "json"})
            assert not allowed.isError
            assert json.loads(allowed.content[0].text) == {"user": "Synthetic User"}
            assert bridge_log.read_text(encoding="utf-8").splitlines() == ["constructed", "whoami"]

        config.write_text('{"tools":{"disabled_groups":["mail"]}}', encoding="utf-8")
        async with session_at(tmp_path) as session:
            names = {tool.name for tool in (await session.list_tools()).tools}
            assert names == ALL_TOOLS - set(TOOL_GROUPS["mail"])
            denied = await session.call_tool(delete, {"entry_id": "synthetic-id"})
            assert denied.isError and f"Unknown tool: {delete}" in denied.content[0].text
            assert bridge_log.read_text(encoding="utf-8").splitlines() == ["constructed"]

    asyncio.run(exercise())


@pytest.mark.parametrize("contents", [None, '{"tools":{"disabled_tool":[]}}'], ids=["missing", "invalid"])
def test_explicit_bad_config_exits_before_bridge(tmp_path, contents):
    path = tmp_path / "policy.json"
    if contents is not None:
        path.write_text(contents, encoding="utf-8")
    log = tmp_path / "bridge.log"
    result = subprocess.run(
        [sys.executable, str(FIXTURE), str(log)],
        env={**get_default_environment(), "OUTLOOK_MCP_CONFIG": str(path)},
        input=b"", capture_output=True, timeout=20,
    )
    assert result.returncode == 2 and not result.stdout
    assert b"Invalid Outlook MCP configuration" in result.stderr
    assert not log.exists()

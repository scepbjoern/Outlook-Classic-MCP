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

# Independent of the registration table: renaming even one tool needs review.
MAIL_NAMES = {"outlook_" + name for name in (
    "list_mails search_mails get_mail send_mail reply_mail forward_mail move_mail delete_mail "
    "mark_mail save_attachments"
).split()}
REMINDER_NAMES = {"outlook_list_reminders", "outlook_preview_reminder_action", "outlook_apply_reminder_action"}
REVIEWED_NAMES = MAIL_NAMES | REMINDER_NAMES | {"outlook_" + name for name in (
    "list_folders create_folder list_events get_event create_event update_event delete_event "
    "respond_event list_contacts search_contacts resolve_name get_contact list_tasks create_task "
    "complete_task query_due list_categories set_category list_rules toggle_rule get_out_of_office whoami"
).split()}

FIXTURE = Path(__file__).with_name("stdio_fixture.py")


@asynccontextmanager
async def session_at(tmp_path, *, errlog=None, log_level="INFO"):
    params = StdioServerParameters(
        command=sys.executable,
        args=[str(FIXTURE), str(tmp_path / "bridge.log")],
        env={"APPDATA": str(tmp_path), "XDG_CONFIG_HOME": str(tmp_path), "OUTLOOK_MCP_LOG": log_level},
    )
    async with stdio_client(params, errlog=errlog if errlog is not None else sys.stderr) as streams:
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
            assert len(original) == 35
            assert original.keys() == REVIEWED_NAMES
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
            identity = json.loads(allowed.content[0].text)
            assert identity["requires_account_selection"]
            assert identity["readiness"] == "ready" and identity["accounts"] == []
            assert identity["profile_name"] is None and identity["current_user"] is None
            assert bridge_log.read_text(encoding="utf-8").splitlines() == ["constructed", "whoami"]

        config.write_text('{"tools":{"disabled_groups":["mail"]}}', encoding="utf-8")
        async with session_at(tmp_path) as session:
            available = {tool.name: tool for tool in (await session.list_tools()).tools}
            blocked = MAIL_NAMES | REMINDER_NAMES | {"outlook_query_due"}
            assert available == {name: tool for name, tool in original.items() if name not in blocked}
            denied = await session.call_tool(delete, {"entry_id": "synthetic-id"})
            assert denied.isError and f"Unknown tool: {delete}" in denied.content[0].text
            assert bridge_log.read_text(encoding="utf-8").splitlines() == ["constructed"]

    asyncio.run(exercise())


def test_registered_tool_rejects_forbidden_store_over_stdio(tmp_path):
    config = tmp_path / "outlook-classic-mcp" / "config.json"
    config.parent.mkdir()
    config.write_text('{"version":2,"accounts":{"read":["b@example.test"]}}', encoding="utf-8")

    async def exercise():
        async with session_at(tmp_path) as session:
            assert "outlook_list_events" in {tool.name for tool in (await session.list_tools()).tools}
            result = await session.call_tool("outlook_list_events", {"store_id": "store-a"})
            assert result.isError and "does not grant read" in result.content[0].text

    asyncio.run(exercise())


def test_read_only_calendar_walkthrough_then_create(tmp_path):
    config = tmp_path / "outlook-classic-mcp" / "config.json"
    config.parent.mkdir()
    read_tools = {"outlook_whoami", "outlook_list_events", "outlook_get_event"}
    policy = {
        "version": 2,
        "tools": {"enabled_tools": sorted(read_tools)},
        "accounts": {"read": ["a@example.test", "b@example.test"], "write": []},
        "calendar_privacy": {"default": "busy_only", "accounts": {"b@example.test": "exclude"}},
    }
    config.write_text(json.dumps(policy), encoding="utf-8")

    async def exercise():
        async with session_at(tmp_path) as session:
            assert {t.name for t in (await session.list_tools()).tools} == read_tools
            result = await session.call_tool("outlook_whoami", {"response_format": "json"})
            identity = json.loads(result.content[0].text)
            assert identity["requires_account_selection"]
            accounts = {row["smtp_address"]: row for row in identity["accounts"]}
            assert set(accounts) == {"a@example.test", "b@example.test"}
            assert all(row["read"] and not row["write"] for row in accounts.values())
            query = {"start": "2026-11-23", "end": "2026-11-24", "limit": 20, "response_format": "json"}
            for address, row in accounts.items():
                listing = await session.call_tool("outlook_list_events", {**query, "store_id": row["store_id"]})
                assert not listing.isError
                items = listing.structuredContent["items"]
                if address == "a@example.test":
                    slot, = items
                    assert set(slot) == {"subject", "start", "end", "busy_status"}
                    assert slot["subject"] == "Appointment" and slot["busy_status"] == 2
                else:
                    assert items == []

        policy["tools"]["enabled_tools"].append("outlook_create_event")
        policy["accounts"]["write"] = ["b@example.test"]
        config.write_text(json.dumps(policy), encoding="utf-8")
        async with session_at(tmp_path) as session:
            assert {t.name for t in (await session.list_tools()).tools} == read_tools | {"outlook_create_event"}
            identity = await session.call_tool("outlook_whoami", {"response_format": "json"})
            rows = json.loads(identity.content[0].text)["accounts"]
            target, = [row for row in rows if row["write"]]
            assert target["smtp_address"] == "b@example.test"
            result = await session.call_tool("outlook_create_event", {
                "store_id": target["store_id"], "subject": "Focus block",
                "start": "2026-11-24T10:00:00+01:00", "end": "2026-11-24T10:30:00+01:00",
                "sensitivity": 2, "reminder_minutes": 15,
            })
            assert not result.isError
            receipt = json.loads(result.content[0].text)
            assert receipt["status"] == "created"
            assert receipt["local_outcome"] == "saved" and receipt["exchange_sync"] == "unverified"
            assert {"entry_id", "subject", "store_id", "start", "end", "busy_status"}.isdisjoint(receipt)
            assert "Focus block" not in result.model_dump_json()

    asyncio.run(exercise())


def test_calendar_privacy_over_stdio(tmp_path):
    config = tmp_path / "outlook-classic-mcp" / "config.json"
    config.parent.mkdir()
    policy = {"version": 2, "accounts": {"read": ["a@example.test", "b@example.test"]},
              "calendar_privacy": {"accounts": {"B@example.test": "exclude"}}}
    config.write_text(json.dumps(policy), encoding="utf-8")
    known_id = "synthetic-shared-id"
    secret = "synthetic-private-content"

    async def exercise():
        async with session_at(tmp_path) as session:
            for fmt in ("json", "markdown"):
                for store, mode in (("store-a", "busy_only"), ("store-b", "exclude")):
                    args = {"store_id": store, "response_format": fmt}
                    listing = await session.call_tool("outlook_list_events", {**args, "start": "2026-11-23", "end": "2026-11-24"})
                    detail = await session.call_tool("outlook_get_event", {**args, "entry_id": known_id})
                    for result in (listing, detail):
                        wire = result.model_dump_json(by_alias=True)
                        assert secret not in wire and known_id not in wire
                    assert not listing.isError
                    if mode == "exclude":
                        assert listing.structuredContent["count"] == 0 and listing.structuredContent["items"] == []
                        missing = await session.call_tool("outlook_get_event", {**args, "entry_id": "missing-id"})
                        assert detail.isError and detail == missing
                    else:
                        row, = listing.structuredContent["items"]
                        assert not detail.isError and detail.structuredContent == row
                        assert set(row) == {"subject", "start", "end", "busy_status"}
                        assert row["subject"] == "Appointment" and row["busy_status"] == 2
                        assert "Appointment" in detail.content[0].text and "Appointment" in listing.content[0].text
                        if fmt == "json":
                            assert json.loads(detail.content[0].text) == row
        policy["calendar_privacy"]["accounts"]["B@example.test"] = "full"
        config.write_text(json.dumps(policy), encoding="utf-8")
        async with session_at(tmp_path) as session:
            detail = await session.call_tool("outlook_get_event", {"account": "b@example.test", "entry_id": known_id})
            assert not detail.isError and detail.structuredContent["body"] == secret

    asyncio.run(exercise())


def test_private_create_is_absent_from_stdio_debug_log(tmp_path):
    config = tmp_path / "outlook-classic-mcp" / "config.json"
    config.parent.mkdir()
    config.write_text('{"version":2,"accounts":{"write":["b@example.test"]}}', encoding="utf-8")
    secret = "synthetic-private-debug-subject"
    log = tmp_path / "server.log"

    async def exercise():
        with log.open("w", encoding="utf-8") as stderr:
            async with session_at(tmp_path, errlog=stderr, log_level="DEBUG") as session:
                result = await session.call_tool("outlook_create_event", {
                    "store_id": "store-b", "sensitivity": 2, "subject": secret, "body": secret,
                    "start": "2026-11-23T10:00", "end": "2026-11-23T11:00",
                })
                assert not result.isError and secret not in result.model_dump_json()
                receipt = json.loads(result.content[0].text)
                assert receipt["local_outcome"] == "saved" and receipt["exchange_sync"] == "unverified"

    asyncio.run(exercise())
    trace = log.read_text(encoding="utf-8", errors="replace")
    assert "Starting outlook_mcp" in trace and secret not in trace


def test_calendar_sensitivity_schema_and_validation_over_stdio(tmp_path):
    async def exercise():
        async with session_at(tmp_path) as session:
            schemas = {tool.name: tool.inputSchema for tool in (await session.list_tools()).tools}
            for name, args in [
                ("outlook_create_event", {"subject": "Synthetic", "start": "2026-11-23", "end": "2026-11-24"}),
                ("outlook_update_event", {"entry_id": "event-b"}),
            ]:
                schema = schemas[name]
                field = schema["properties"]["sensitivity"]
                assert "sensitivity" not in schema["required"]
                assert (field["type"], field["minimum"], field["maximum"]) == ("integer", 0, 3)
                result = await session.call_tool(name, {**args, "sensitivity": True})
                assert result.isError and "sensitivity" in result.content[0].text
            assert (tmp_path / "bridge.log").read_text(encoding="utf-8").splitlines() == ["constructed"]

    asyncio.run(exercise())


@pytest.mark.parametrize("contents", [
    None, '{"tools":{"disabled_tool":[]}}', '{"version":2,"accounts":{"write":true}}',
], ids=["missing", "invalid", "invalid-account-policy"])
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

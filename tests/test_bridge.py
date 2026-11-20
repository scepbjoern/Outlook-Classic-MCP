"""Tests for the bridge reconnect cycle, with COM faked out.

These run the real OutlookBridge thread machinery (CoInitializeEx and
all) but replace ``dynamic.Dispatch`` so no Outlook is needed.
"""

from __future__ import annotations

import asyncio
import sys
import threading
from unittest.mock import Mock
from types import SimpleNamespace

import pytest

from outlook_mcp.errors import OutlookError

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="bridge requires pywin32"
)


def _fake_outlook():
    namespace = SimpleNamespace(CurrentProfileName="Synthetic profile", CurrentUser=SimpleNamespace(Name="Test User"))
    return SimpleNamespace(GetNamespace=lambda kind: namespace)


def _rpc_unavailable() -> Exception:
    import pythoncom

    return pythoncom.com_error(
        -2147023174, "The RPC server is unavailable.", None, None
    )


@pytest.fixture
def bridge(monkeypatch):
    from win32com.client import dynamic

    from outlook_mcp.bridge import OutlookBridge

    monkeypatch.setattr(dynamic, "Dispatch", lambda name: _fake_outlook())
    monkeypatch.setattr("outlook_mcp.bridge._launch_outlook", lambda: False)
    b = OutlookBridge()
    b.start()
    yield b
    b.stop()


def test_call_runs_on_com_thread(bridge):
    result = asyncio.run(bridge.call(lambda outlook, ns: ns.CurrentUser.Name))
    assert result == "Test User"


def test_reconnects_and_retries_after_disconnect(bridge):
    attempts = []

    def flaky(outlook, namespace):
        attempts.append(namespace)
        if len(attempts) == 1:
            raise _rpc_unavailable()
        return "recovered"

    assert asyncio.run(bridge.call(flaky)) == "recovered"
    # Retried once, against a freshly attached namespace.
    assert len(attempts) == 2
    assert attempts[0] is not attempts[1]


def test_non_disconnect_errors_are_not_retried(bridge):
    attempts = []

    def broken(outlook, namespace):
        attempts.append(1)
        raise ValueError("bad input")

    with pytest.raises(ValueError):
        asyncio.run(bridge.call(broken))
    assert len(attempts) == 1


def test_write_is_not_replayed_after_disconnect(bridge):
    attempts = []

    def interrupted_write(outlook, namespace):
        attempts.append(namespace)
        raise _rpc_unavailable()

    with pytest.raises(OutlookError, match="outcome is unknown"):
        asyncio.run(bridge.call(interrupted_write, retry_on_disconnect=False))
    assert len(attempts) == 1


def test_reconnects_when_stale_proxy_raises_bare_attribute_error(bridge):
    # Outlook restarted under us: pywin32's dynamic __getattr__ swallows the
    # RPC error and raises a bare AttributeError. The raw IDispatch probe
    # still reports the disconnect, so the bridge must re-attach.
    def dead_get_ids(name):
        raise _rpc_unavailable()

    bridge._namespace._oleobj_ = SimpleNamespace(GetIDsOfNames=dead_get_ids)
    attempts = []

    def stale(outlook, namespace):
        attempts.append(namespace)
        if len(attempts) == 1:
            raise AttributeError("GetNamespace.GetDefaultFolder")
        return "recovered"

    assert asyncio.run(bridge.call(stale)) == "recovered"
    assert len(attempts) == 2
    assert attempts[0] is not attempts[1]


def test_attribute_error_on_live_proxy_is_not_retried(bridge):
    bridge._namespace._oleobj_ = SimpleNamespace(GetIDsOfNames=lambda name: 1)
    attempts = []

    def typo(outlook, namespace):
        attempts.append(1)
        raise AttributeError("GetNamespace.NoSuchThing")

    with pytest.raises(AttributeError):
        asyncio.run(bridge.call(typo))
    assert len(attempts) == 1


def test_error_surfaces_when_reattach_fails(bridge, monkeypatch):
    from win32com.client import dynamic

    import outlook_mcp.bridge as bridge_mod

    # After the disconnect, make re-attach fail fast: Dispatch raises and
    # the attach deadline is immediate.
    def dead_dispatch(name):
        raise _rpc_unavailable()

    monkeypatch.setattr(dynamic, "Dispatch", dead_dispatch)
    monkeypatch.setattr(bridge_mod, "_ATTACH_TIMEOUT_SEC", 0)

    def dies(outlook, namespace):
        raise _rpc_unavailable()

    with pytest.raises(OutlookError, match="could not be relaunched"):
        asyncio.run(bridge.call(dies))


def test_blocked_profile_prompt_does_not_block_stdio_or_queue_a_write(monkeypatch):
    from win32com.client import dynamic

    from outlook_mcp.bridge import OutlookBridge

    entered, release = threading.Event(), threading.Event()
    application = _fake_outlook()

    def dispatch(name):
        entered.set()
        assert release.wait(5)
        return application

    monkeypatch.setattr(dynamic, "Dispatch", dispatch)
    write = Mock()
    bridge = OutlookBridge()
    bridge.start(wait=False)
    try:
        assert entered.wait(2)
        with pytest.raises(OutlookError, match="starting.*did not start.*profile selection"):
            asyncio.run(bridge.call(write, retry_on_disconnect=False))
        release.set()
        assert bridge._ready.wait(2)
        assert asyncio.run(bridge.call(lambda app, ns: ns.CurrentProfileName)) == "Synthetic profile"
        write.assert_not_called()
    finally:
        release.set()
        bridge.stop()


def test_cold_launch_attempt_then_ready(monkeypatch):
    from win32com.client import dynamic

    from outlook_mcp.bridge import OutlookBridge

    launch = Mock(return_value=True)
    dispatch = Mock(side_effect=[_rpc_unavailable(), _fake_outlook()])
    monkeypatch.setattr(dynamic, "Dispatch", dispatch)
    monkeypatch.setattr("outlook_mcp.bridge._launch_outlook", launch)
    bridge = OutlookBridge()
    try:
        bridge.start()
        launch.assert_called_once_with()
        assert dispatch.call_count == 2
    finally:
        bridge.stop()


@pytest.mark.parametrize("cancel_request", [False, True])
def test_expired_or_cancelled_queued_write_never_reaches_provider(bridge, monkeypatch, cancel_request):
    import outlook_mcp.bridge as bridge_mod

    entered, release = threading.Event(), threading.Event()
    write = Mock()

    def hold_worker(app, ns):
        entered.set()
        assert release.wait(5)

    async def exercise():
        first = asyncio.create_task(bridge.call(hold_worker))
        try:
            assert await asyncio.to_thread(entered.wait, 2)
            monkeypatch.setattr(bridge_mod, "_CALL_TIMEOUT_SEC", 0.05)
            second = asyncio.create_task(bridge.call(write, retry_on_disconnect=False))
            await asyncio.sleep(0)
            if cancel_request:
                second.cancel()
                with pytest.raises(asyncio.CancelledError):
                    await second
            else:
                with pytest.raises(OutlookError, match="timed out in the queue.*did not start"):
                    await second
        finally:
            release.set()
            await first
        monkeypatch.setattr(bridge_mod, "_CALL_TIMEOUT_SEC", 2)
        await bridge.call(lambda app, ns: None)
        write.assert_not_called()

    asyncio.run(exercise())


def test_inflight_write_timeout_reports_unknown_and_never_replays(bridge, monkeypatch):
    import outlook_mcp.bridge as bridge_mod

    release = threading.Event()
    effects = []

    def slow_write(app, ns):
        effects.append(1)
        assert release.wait(5)

    monkeypatch.setattr(bridge_mod, "_CALL_TIMEOUT_SEC", 0.1)
    try:
        with pytest.raises(OutlookError, match="timed out.*outcome is unknown.*Do not repeat"):
            asyncio.run(bridge.call(slow_write, retry_on_disconnect=False))
    finally:
        release.set()
    monkeypatch.setattr(bridge_mod, "_CALL_TIMEOUT_SEC", 2)
    asyncio.run(bridge.call(lambda app, ns: None))
    assert effects == [1]


def test_reconnect_blocks_other_writes_and_refuses_changed_profile(bridge, monkeypatch):
    from win32com.client import dynamic

    entered, release = threading.Event(), threading.Event()
    other = _fake_outlook()
    other.GetNamespace("MAPI").CurrentProfileName = "Other synthetic profile"

    def reconnect(name):
        entered.set()
        assert release.wait(5)
        return other

    monkeypatch.setattr(dynamic, "Dispatch", reconnect)
    read = Mock(side_effect=_rpc_unavailable())
    write = Mock()

    async def exercise():
        first = asyncio.create_task(bridge.call(read))
        try:
            assert await asyncio.to_thread(entered.wait, 2)
            with pytest.raises(OutlookError, match="reconnecting.*did not start"):
                await bridge.call(write, retry_on_disconnect=False)
        finally:
            release.set()
        with pytest.raises(OutlookError, match="same Outlook profile"):
            await first

    asyncio.run(exercise())
    assert read.call_count == 1
    write.assert_not_called()


@pytest.mark.parametrize("tool, args, action", [
    ("send_mail", {"subject": "Synthetic", "body": "Example", "to": ["guest@example.test"]}, "Send"),
    ("create_event", {"subject": "Synthetic", "start": "2026-11-23", "end": "2026-11-24"}, "Save"),
    ("create_event", {"subject": "Synthetic", "start": "2026-11-23", "end": "2026-11-24",
                      "attendees": ["guest@example.test"]}, "Send"),
])
def test_registered_writes_are_never_replayed_after_lost_ack(monkeypatch, tool, args, action):
    from win32com.client import dynamic

    from outlook_mcp.config import ToolPolicy
    from outlook_mcp.server import build_server
    from tests.test_mailbox import make_profile

    ns = make_profile(("a",))
    ns.CurrentProfileName = "Synthetic profile"
    monkeypatch.setattr(dynamic, "Dispatch", lambda name: SimpleNamespace(GetNamespace=lambda kind: ns))
    server, bridge = build_server(policy=ToolPolicy())
    items = ns.Stores[0].GetDefaultFolder.return_value.Items
    item = items.Add.return_value
    effect = getattr(item, action)
    effect.side_effect = _rpc_unavailable()
    bridge.start()
    try:
        with pytest.raises(Exception, match="outcome is unknown") as error:
            asyncio.run(server.call_tool("outlook_" + tool, args))
        if "attendees" in args:
            assert "saved the appointment locally" in str(error.value)
            item.Save.assert_called_once_with()
        items.Add.assert_called_once()
        effect.assert_called_once_with()
    finally:
        bridge.stop()


def test_unavailable_attach_recovers_without_running_triggering_write(bridge, monkeypatch):
    import outlook_mcp.bridge as bridge_mod
    from win32com.client import dynamic

    monkeypatch.setattr(bridge_mod, "_ATTACH_TIMEOUT_SEC", 0)
    with pytest.raises(OutlookError, match="could not be relaunched"):
        asyncio.run(bridge.call(Mock(side_effect=_rpc_unavailable())))
    attached = threading.Event()
    monkeypatch.setattr(bridge_mod, "_ATTACH_TIMEOUT_SEC", 2)
    def dispatch(name):
        attached.set()
        return _fake_outlook()
    monkeypatch.setattr(dynamic, "Dispatch", dispatch)
    write = Mock()
    with pytest.raises(OutlookError, match="unavailable.*did not start"):
        asyncio.run(bridge.call(write, retry_on_disconnect=False))
    assert attached.wait(2)
    # A readiness probe may race the last instructions in _attach.
    async def wait_ready():
        for _ in range(100):
            try:
                return await bridge.call(lambda app, ns: ns.CurrentProfileName)
            except OutlookError as exc:
                assert "reconnecting" in str(exc)
                await asyncio.sleep(0.01)
        pytest.fail("worker did not recover")
    assert asyncio.run(wait_ready()) == "Synthetic profile"
    write.assert_not_called()


def test_queued_write_cannot_cross_session_boundary(bridge):
    entered, release = threading.Event(), threading.Event()
    write = Mock()
    def restart(*_):
        entered.set()
        assert release.wait(5)
        raise _rpc_unavailable()
    read = Mock(side_effect=[None, "ok"])
    def interrupted(app, ns):
        return read() or restart()

    async def exercise():
        first = asyncio.create_task(bridge.call(interrupted))
        try:
            assert await asyncio.to_thread(entered.wait, 2)
            second = asyncio.create_task(bridge.call(write, retry_on_disconnect=False))
            await asyncio.sleep(0)
        finally:
            release.set()
        assert await first == "ok"
        with pytest.raises(OutlookError, match="session changed.*did not start"):
            await second
    asyncio.run(exercise())
    write.assert_not_called()


def test_replay_rechecks_exact_store_even_if_profile_name_is_unchanged(monkeypatch):
    from win32com.client import dynamic
    from outlook_mcp.server import build_server
    from tests.test_mailbox import make_profile

    original, replacement = make_profile(("a",)), make_profile(("b",))
    for ns in (original, replacement):
        ns.CurrentProfileName = "Synthetic profile"
    original.Stores[0].GetDefaultFolder.side_effect = _rpc_unavailable()
    applications = [SimpleNamespace(GetNamespace=Mock(return_value=ns)) for ns in (original, replacement)]
    monkeypatch.setattr(dynamic, "Dispatch", Mock(side_effect=applications))
    server, bridge = build_server()
    bridge.start()
    try:
        with pytest.raises(Exception, match="Unknown Outlook (account|store_id)"):
            asyncio.run(server.call_tool("outlook_list_mails", {}))
        replacement.Stores[0].GetDefaultFolder.assert_not_called()
    finally:
        bridge.stop()


def test_export_is_read_authorized_but_never_replayed(monkeypatch, tmp_path):
    from win32com.client import dynamic
    from outlook_mcp.config import parse_tool_policy
    from outlook_mcp.constants import OL_CLASS_MAIL
    from outlook_mcp.server import build_server
    from tests.test_mailbox import make_profile

    ns = make_profile(("a",))
    ns.CurrentProfileName = "Synthetic profile"
    item = ns.Stores[0].GetDefaultFolder.return_value.Items.Add.return_value
    item.Class = OL_CLASS_MAIL
    attachment = SimpleNamespace(FileName="synthetic.txt", SaveAsFile=Mock(side_effect=_rpc_unavailable()))
    item.Attachments = [attachment]
    monkeypatch.setattr(dynamic, "Dispatch", lambda name: SimpleNamespace(GetNamespace=lambda kind: ns))
    monkeypatch.setattr("outlook_mcp.client.mail.validate_output_dir", lambda path: str(tmp_path))
    policy = parse_tool_policy('{"version":2,"accounts":{"read":["a@example.test"]}}')
    server, bridge = build_server(policy=policy)
    bridge.start()
    try:
        with pytest.raises(Exception, match="outcome is unknown"):
            asyncio.run(server.call_tool("outlook_save_attachments", {
                "entry_id": item.EntryID, "output_dir": str(tmp_path),
            }))
        attachment.SaveAsFile.assert_called_once()
    finally:
        bridge.stop()

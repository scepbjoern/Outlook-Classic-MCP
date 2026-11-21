"""Registered reminder tools against a synthetic, changing profile-wide collection."""

import asyncio
import json
import sys
from datetime import datetime, timedelta, timezone
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from outlook_mcp.client import reminder_actions
from .test_mailbox import Collection, make_profile, registered
from .test_reminders import reminder


STORE = {"store_id": "store-b"}


def alert(key, kind=26, **kwargs):
    result = reminder(key, **kwargs)
    result.Item.Class = kind
    result.Item.EntryID = key
    result.Dismiss = Mock(side_effect=lambda: setattr(result, "IsVisible", False))
    result.Snooze = Mock(side_effect=lambda minutes: setattr(result, "IsVisible", False))
    return result


def profile(*alerts, **permissions):
    ns = make_profile()
    ns.Application.Reminders = Collection(alerts)
    return ns, registered(ns, **permissions)


def listed(call, **kwargs):
    return call("list_reminders", **STORE, **kwargs)["items"]


def preview(call, selection=None, **kwargs):
    if selection is None:
        selection = listed(call)[0]["selection"]
    return call("preview_reminder_action", **STORE, selection=selection, **({"action": "dismiss"} | kwargs))


def apply(call, receipt, **kwargs):
    return call("apply_reminder_action", **(STORE | {"confirmed": True} | kwargs),
                confirmation_token=receipt["confirmation_token"])


@pytest.mark.parametrize("kind", [26, 43, 48])
def test_preview_dismiss_exactly_one_despite_reorder(kind):
    selected, other = alert("a", kind), alert("b", kind)
    before = vars(selected.Item).copy()
    ns, call = profile(selected, other)
    receipt = preview(call)
    assert receipt["account"] == "b@example.test" and receipt["reminder"]["subject"] == "a"
    assert receipt["action"] == "dismiss" and receipt["until_local"] is None
    with pytest.raises(Exception, match="confirm"):
        apply(call, receipt, confirmed=False)
    selected.Dismiss.assert_not_called()
    ns.Application.Reminders.reverse()
    result = apply(call, receipt)
    assert result["local_outcome"] == "completed" and result["exchange_sync"] == "unverified"
    selected.Dismiss.assert_called_once_with()
    other.Dismiss.assert_not_called()
    assert [r["subject"] for r in listed(call)] == ["b"]
    assert vars(selected.Item) == before
    with pytest.raises(Exception, match="confirmation"):
        apply(call, receipt)


@pytest.mark.parametrize("change", ["gone", "invisible", "store", "id", "next", "original", "subject",
                                    "privacy", "occurrence", "duplicate", "broken collection"])
def test_changed_selection_requires_a_new_preview(change):
    selected = alert("a", state=2)
    selected.Item.StartUTC = datetime(2026, 11, 7, 15)
    ns, call = profile(selected)
    receipt = preview(call)
    assert receipt["reminder"]["original_start"] == "2026-11-07T15:00:00+00:00"
    collection = ns.Application.Reminders
    if change == "gone":
        collection.clear()
    elif change == "invisible":
        selected.IsVisible = False
    elif change == "store":
        selected.Item.Parent.StoreID = "store-a"
    elif change == "id":
        selected.Item.EntryID = "different"
    elif change == "next":
        selected.NextReminderDate += timedelta(minutes=10)
    elif change == "original":
        selected.OriginalReminderDate += timedelta(minutes=10)
    elif change == "subject":
        selected.Item.Subject = "changed"
    elif change == "privacy":
        selected.Item.Sensitivity = 2
    elif change == "occurrence":
        selected.Item.StartUTC += timedelta(days=7)
    elif change == "duplicate":
        collection.append(selected)
    else:
        collection.append(object())
    with pytest.raises(Exception, match="Refresh"):
        apply(call, receipt)
    selected.Dismiss.assert_not_called()


def test_snooze_retains_preview_minute_and_updates_only_selected_alert(monkeypatch):
    selected, other = alert("a"), alert("b", 43)
    clock = SimpleNamespace(time=Mock(return_value=1795258831), monotonic=Mock(return_value=10))
    monkeypatch.setattr(reminder_actions, "time", clock)
    ns, call = profile(selected, other)
    receipt = preview(call, action="snooze", delay_minutes=1440)
    target = datetime.fromisoformat(receipt["until_local"])
    assert target.timestamp() == (1795258831 // 60 + 1440) * 60
    clock.time.return_value += 119
    selected.Snooze.side_effect = lambda minutes: (
        setattr(selected, "IsVisible", False), setattr(selected, "NextReminderDate", target.replace(tzinfo=None)))
    result = apply(call, receipt)
    assert result["until_local"] == receipt["until_local"]
    selected.Snooze.assert_called_once_with(1438)
    other.Snooze.assert_not_called()
    assert len(listed(call)) == 1
    assert {r["subject"]: r["status"] for r in listed(call, mode="all")} == {"a": "snoozed", "b": "active"}


def test_absolute_target_and_expired_confirmation(monkeypatch):
    selected = alert("a")
    _, call = profile(selected)
    now = datetime.now(timezone.utc).replace(second=0, microsecond=0)
    clock = SimpleNamespace(time=Mock(return_value=now.timestamp()), monotonic=Mock(return_value=10))
    monkeypatch.setattr(reminder_actions, "time", clock)
    until = (now + timedelta(days=1)).astimezone().isoformat(timespec="minutes")
    receipt = preview(call, action="snooze", until_local=until)
    assert receipt["until_local"] == until
    clock.monotonic.return_value += 301
    with pytest.raises(Exception, match="expired"):
        apply(call, receipt)
    receipt = preview(call, action="snooze", delay_minutes=1)
    clock.time.return_value += 65
    with pytest.raises(Exception, match="passed"):
        apply(call, receipt)
    selected.Snooze.assert_not_called()


@pytest.mark.parametrize("args", [{"action": "snooze"}, {"action": "dismiss", "delay_minutes": 1},
                                  {"action": "snooze", "until_local": "2030-01-01T09:00"},
                                  {"action": "snooze", "until_local": "1999-01-01T09:00+00:00"},
                                  {"action": "snooze", "until_local": "tomorrow", "delay_minutes": 1}])
def test_invalid_action_and_target(args):
    selected = alert("a")
    _, call = profile(selected)
    with pytest.raises(Exception):
        preview(call, **args)
    selected.Snooze.assert_not_called()
    selected.Dismiss.assert_not_called()


def test_store_binding_permissions_and_explicit_selection():
    b, a = alert("same"), alert("same", store="store-a")
    _, call = profile(b, a, write=("a", "b"))
    receipt = preview(call)
    with pytest.raises(Exception, match="mismatched"):
        apply(call, receipt, store_id="store-a")
    a.Dismiss.assert_not_called()
    b.Dismiss.assert_not_called()
    _, read_only = profile(b, write=())
    with pytest.raises(Exception, match="read and write"):
        apply(read_only, preview(read_only))
    _, write_only = profile(b, read=())
    with pytest.raises(Exception, match="read"):
        listed(write_only)
    with pytest.raises(Exception, match="read and write"):
        apply(write_only, receipt)
    with pytest.raises(Exception, match="Select"):
        call("list_reminders")


def test_projection_status_privacy_and_listing_only_rows():
    class Foreign:
        Parent = SimpleNamespace(StoreID="store-a")

        def __getattr__(self, name):
            pytest.fail("Foreign property read: " + name)

    class PrivateReminder:
        Item = alert("private title", sensitivity=2).Item

        def __getattr__(self, name):
            pytest.fail("Private reminder property read: " + name)

    master = alert("hidden master title", state=1)
    master.Item.GetRecurrencePattern = lambda: SimpleNamespace(Exceptions=Collection())
    private_mail = alert("private mail", 43, sensitivity=2, snooze=20, visible=False)
    task = alert("future task", 48, visible=False)
    _, call = profile(master, private_mail, task, PrivateReminder(), SimpleNamespace(Item=Foreign()))
    rows = listed(call, mode="all")
    assert len(rows) == 3
    by_type = {r["item_type"]: r for r in rows}
    assert by_type["calendar"]["subject"] is None and by_type["calendar"]["appointment_start"] is None
    assert by_type["mail"]["subject"] is None and by_type["mail"]["status"] == "snoozed"
    assert by_type["task"]["status"] == "scheduled" and by_type["task"]["subject"] == "future task"
    assert all(r["selection"] is None for r in rows)


def test_partial_collection_and_missing_slot_are_readable_but_not_actionable():
    good, no_slot = alert("a"), alert("b", state=3)
    ns, call = profile(good, no_slot)
    rows = listed(call)
    assert rows[0]["selection"] and rows[1]["selection"] is None
    ns.Application.Reminders.append(object())
    result = call("list_reminders", **STORE)
    assert result["coverage"] == "partial" and result["count"] == 2
    assert all(row["selection"] is None for row in result["items"])


def test_modified_exception_is_bound_to_original_slot():
    moved = alert("moved", state=3)
    original = datetime(2026, 10, 25, 1, 30)
    accessor = moved.Item.PropertyAccessor = Mock()
    accessor.GetProperty.return_value = original
    _, call = profile(moved)
    receipt = preview(call)
    assert receipt["reminder"]["is_exception"] is True
    assert receipt["reminder"]["original_start"] == "2026-10-25T01:30:00+00:00"
    accessor.GetProperty.return_value += timedelta(hours=1)
    with pytest.raises(Exception, match="changed"):
        apply(call, receipt)
    moved.Dismiss.assert_not_called()


def test_new_preview_supersedes_prior_action_on_same_selection():
    selected = alert("a")
    _, call = profile(selected)
    dismiss = preview(call)
    snooze = preview(call, action="snooze", delay_minutes=30)
    with pytest.raises(Exception, match="confirmation"):
        apply(call, dismiss)
    apply(call, snooze)
    selected.Dismiss.assert_not_called()
    selected.Snooze.assert_called_once()


@pytest.mark.skipif(sys.platform != "win32", reason="exercises the Windows STA")
def test_disconnect_during_dismiss_consumes_confirmation_without_replaying(monkeypatch):
    from win32com.client import dynamic
    from outlook_mcp.server import build_server
    from .test_bridge import _rpc_unavailable

    ns = make_profile(names=("b",))
    selected = alert("a")
    ns.Application.Reminders = Collection([selected])
    ns.CurrentProfileName = "Synthetic"
    selected.Dismiss.side_effect = _rpc_unavailable()
    monkeypatch.setattr(dynamic, "Dispatch", lambda _: SimpleNamespace(GetNamespace=lambda _: ns))
    server, bridge = build_server()

    def call(name, **kwargs):
        result = asyncio.run(server.call_tool("outlook_" + name, kwargs))
        content = result[0] if isinstance(result, tuple) else result
        return json.loads(content[0].text)

    bridge.start()
    try:
        receipt = preview(call)
        with pytest.raises(Exception, match="outcome is unknown"):
            apply(call, receipt)
        with pytest.raises(Exception, match="confirmation"):
            apply(call, receipt)
        selected.Dismiss.assert_called_once_with()
    finally:
        bridge.stop()


def test_combined_read_paginates_and_preserves_account_and_occurrence_for_preview():
    first, second = alert("same", store="store-a"), alert("same", state=2)
    second.Item.StartUTC = datetime(2026, 11, 22, 10)
    ns, call = profile(first, second)
    page = call("list_reminders", all_readable=True, limit=1)
    following = call("list_reminders", all_readable=True, limit=1, offset=page["next_offset"])
    assert page["coverage"] == following["coverage"] == "complete"
    assert {m["store_id"] for m in page["stores"]} == {"store-a", "store-b"}
    rows = {r["store_id"]: r for r in page["items"] + following["items"]}
    assert rows["store-a"]["selection"] != rows["store-b"]["selection"]
    receipt = preview(call, rows["store-b"]["selection"])
    assert receipt["reminder"]["original_start"] == "2026-11-22T10:00:00+00:00"
    apply(call, receipt)
    first.Dismiss.assert_not_called()
    second.Dismiss.assert_called_once()
    with pytest.raises(Exception, match="either"):
        call("list_reminders", all_readable=True, **STORE)


def test_foreign_mapi_alias_is_skipped_before_details_and_unknown_owner_remains_partial():
    class Foreign:
        Parent = SimpleNamespace(StoreID=None)
        PropertyAccessor = Mock()

        def __getattr__(self, name):
            pytest.fail("Accessed foreign item field " + name)

    foreign = Foreign()
    foreign.PropertyAccessor.BinaryToString.return_value = "aliased-a"
    ns, call = profile(SimpleNamespace(Item=foreign), alert("permitted"), read=("b",))
    ns.CompareEntryIDs.side_effect = lambda raw, known: raw == "aliased-a" and known == "store-a"
    for arguments in (STORE, {"all_readable": True}):
        result = call("list_reminders", **arguments)
        assert result["coverage"] == "complete"
        row, = result["items"]
        assert row["subject"] == "permitted" and row["selection"]
    assert result["stores"] == [{"account": "b@example.test", "store_id": "store-b"}]
    assert preview(call, row["selection"])["store_id"] == "store-b"
    foreign.PropertyAccessor.BinaryToString.return_value = "unverifiable"
    for arguments in (STORE, {"all_readable": True}):
        result = call("list_reminders", **arguments)
        assert result["coverage"] == "partial" and result["items"][0]["selection"] is None


def test_aggregate_respects_account_privacy_and_rejects_zero_readable_stores():
    a, b = alert("hidden", store="store-a", sensitivity=2), alert("private", sensitivity=2)
    ns, call = profile(a, b, privacy={"default": "busy_only", "accounts": {"b@example.test": "full"}})
    row, = call("list_reminders", all_readable=True)["items"]
    assert row["store_id"] == "store-b" and row["subject"] is None
    with pytest.raises(Exception, match="No readable"):
        registered(ns, read=())("list_reminders", all_readable=True)

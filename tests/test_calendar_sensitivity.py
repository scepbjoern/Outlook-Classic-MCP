"""Native appointment sensitivity through the registered tools and synthetic Outlook."""

from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from outlook_mcp.client import calendar
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.formatting import format_response
from tests.test_mailbox import make_profile, registered


NEW_EVENT = {"subject": "Synthetic appointment", "start": "2026-11-23T10:00", "end": "2026-11-23T11:00"}
STORE = {"store_id": "store-b"}
EXISTING = {"entry_id": "event-b", **STORE}


@pytest.mark.parametrize("value,label", [(0, "Normal"), (1, "Personal"), (2, "Private"), (3, "Confidential")])
def test_create_and_read_each_level(value, label):
    ns = make_profile()
    event = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    event.Sensitivity = 3 if value == 0 else 0
    event.Categories = "Privat"
    saved = []
    event.Save.side_effect = lambda: saved.append(event.Sensitivity)
    call = registered(ns)
    call("create_event", sensitivity=value, **NEW_EVENT, **STORE)
    assert saved == [value]

    listing = call("list_events", **STORE)
    detail = call("get_event", **EXISTING)
    assert detail["categories"] == "Privat"
    for row in (listing["items"][0], detail):
        assert row["sensitivity"] == value
        assert row["sensitivity_label"] == label
    for payload in (listing, detail):
        assert f"{label} ({value})" in format_response(payload, "markdown")


@pytest.mark.parametrize("recurring", [False, True], ids=["single", "series"])
def test_update_preserves_private_on_move_and_accepts_explicit_zero(recurring):
    ns = make_profile()
    event = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    event.Sensitivity = 3  # Synthetic provider default, preserved on create.
    saved = []
    event.Save.side_effect = lambda: saved.append(event.Sensitivity)
    call = registered(ns)
    args = {"recurrence": {"type": "weekly", "occurrences": 3}} if recurring else {}
    call("create_event", **NEW_EVENT, **STORE, **args)
    if recurring:
        assert event.GetRecurrencePattern.return_value.Occurrences == 3
        event.RecurrenceState = 1  # Master resolved by EntryID.
    call("update_event", sensitivity=2, **EXISTING)
    call("update_event", start="2026-11-24T10:00", end="2026-11-24T11:00", **EXISTING)
    detail = call("get_event", **EXISTING)
    assert detail["start"].startswith("2026-11-24T10:00")
    assert detail["sensitivity_label"] == "Private"
    call("update_event", sensitivity=0, **EXISTING)
    assert saved == [3, 2, 2, 0]
    assert event.Categories == ""


def test_series_and_expanded_occurrences_report_their_own_values():
    ns = make_profile()
    items = ns.Stores[1].GetDefaultFolder.return_value.Items
    master = items.Add.return_value
    call = registered(ns)
    call("create_event", recurrence={"type": "weekly", "occurrences": 2}, sensitivity=2, **NEW_EVENT, **STORE)
    master.RecurrenceState = 1
    occurrences = [
        SimpleNamespace(Parent=master.Parent, EntryID=master.EntryID, RecurrenceState=state, Sensitivity=value)
        for state, value in [(2, 2), (3, 1)]  # Regular occurrence and modified exception.
    ]
    items.IncludeRecurrences = False
    items.Restrict.side_effect = lambda query: occurrences if items.IncludeRecurrences else [master]
    masters = call("list_events", include_recurrences=False, **STORE)["items"]
    assert [row["sensitivity"] for row in masters] == [2]
    expanded = call("list_events", include_recurrences=True, **STORE)["items"]
    assert [(row["sensitivity"], row["sensitivity_label"]) for row in expanded] == [(2, "Private"), (1, "Personal")]


@pytest.mark.parametrize("value", [-1, 4, True, False, 2.0, 1.5, "2", "Private", "null", None])
@pytest.mark.parametrize("tool,args", [("create_event", NEW_EVENT), ("update_event", {"entry_id": "event-b", "subject": "Changed"})])
def test_invalid_explicit_sensitivity_never_reaches_outlook(tool, args, value):
    ns = make_profile()
    with pytest.raises(Exception, match="sensitivity"):
        registered(ns)(tool, sensitivity=value, **args, **STORE)
    assert not ns.mock_calls
    for store in ns.Stores:
        store.GetDefaultFolder.assert_not_called()
        assert store.GetDefaultFolder.return_value.Items.Add.return_value.Subject == store.DisplayName


@pytest.mark.parametrize("operation,args", [(calendar.create_event, NEW_EVENT), (calendar.update_event, {"entry_id": "event-b"})])
def test_client_rejects_invalid_sensitivity_before_outlook_access(operation, args):
    namespace, mailbox = MagicMock(), MagicMock()
    with pytest.raises(OutlookError, match="sensitivity"):
        operation(namespace, mailbox, sensitivity=False, **args)
    assert not namespace.mock_calls and not mailbox.mock_calls


def test_unavailable_or_unknown_sensitivity_is_never_reported_as_normal():
    for event, code in [(SimpleNamespace(), None), (SimpleNamespace(Sensitivity=9), 9)]:
        summary = calendar._event_summary(event)
        assert summary["sensitivity"] == code and summary["sensitivity_label"] == "Unknown"

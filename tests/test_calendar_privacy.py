"""Only synthetic stores, appointments and COM failures."""

import json
from datetime import datetime
from types import SimpleNamespace
from unittest.mock import PropertyMock

import pytest

from outlook_mcp.client.calendar_privacy import availability, classify_event
from outlook_mcp.config import PrivacyMode, ToolConfigError, parse_tool_policy
from tests.test_mailbox import WINDOW, Collection, make_profile, registered


SELECT = {"store_id": "store-b"}
SECRET = "fictional-private-marker"
START, END = datetime(2026, 11, 23, 10), datetime(2026, 11, 23, 11)
SLOT = {"subject": "Appointment", "start": START.astimezone().isoformat(),
        "end": END.astimezone().isoformat(), "busy_status": 0}


def appointment(ns):
    event = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    event.Sensitivity, event.BusyStatus = 2, 0
    event.Subject = event.Body = event.Location = event.Organizer = event.Categories = SECRET
    event.Start, event.End = START, END
    return event


@pytest.mark.parametrize("sensitivity", [0, 1, 3])
def test_busy_only_can_disclose_nonprivate_calendar_previews(sensitivity):
    ns = make_profile()
    event = appointment(ns)
    event.Sensitivity = sensitivity
    result = registered(ns, write=())("list_events", **WINDOW, **SELECT)
    row, = result["items"]
    assert row["sensitivity"] == sensitivity
    assert row["preview"] == event.Body and row["subject"] == event.Subject


@pytest.mark.parametrize("raw", ["{}", '{"version":1}', '{"version":2,"accounts":{}}'])
def test_legacy_default(raw):
    assert parse_tool_policy(raw).calendar_privacy.for_account("a@example.test") == PrivacyMode.BUSY_ONLY


@pytest.mark.parametrize("setting", [
    None, {"default": None}, {"default": "Busy_Only"}, {"accounts": []},
    {"accounts": {"A@example.test": "full", "a@example.test": "exclude"}},
    {"accounts": {"Mailbox A": "full"}}, {"accounts": {"a@example.test": 2}},
    {"defaults": "full"},
])
def test_bad_privacy_settings_fail_startup(setting):
    with pytest.raises(ToolConfigError):
        parse_tool_policy(json.dumps({"version": 2, "accounts": {}, "calendar_privacy": setting}))


def test_per_account_modes_require_v2_and_cannot_lift_global_exclude():
    config = {"calendar_privacy": {"default": "exclude", "accounts": {"a@example.test": "full"}}}
    with pytest.raises(ToolConfigError):
        parse_tool_policy(json.dumps(config))
    policy = parse_tool_policy(json.dumps({**config, "version": 2, "accounts": {}}))
    assert policy.calendar_privacy.for_account("A@example.test") == PrivacyMode.EXCLUDE


@pytest.mark.parametrize("mode", ["busy_only", "exclude", "full"])
def test_all_actions_and_explicit_downgrade(mode):
    ns = make_profile()
    ev = appointment(ns)
    call = registered(ns, privacy={"default": mode})
    by_id = {**SELECT, "entry_id": ev.EntryID}
    created = call("create_event", subject=SECRET, body=SECRET, sensitivity=2, busy_status=0,
                   start=START.isoformat(), end=END.isoformat(), **SELECT)
    changed = call("update_event", location=SECRET, **by_id)
    category = call("set_category", categories=SECRET, **by_id)
    response = call("respond_event", response="accept", send_response=False, **by_id)
    deleted = call("delete_event", **by_id)
    for receipt in (created, changed, category, response, deleted):
        assert receipt.pop("exchange_sync") == "unverified"
        assert receipt.pop("local_outcome") == ("saved" if receipt is created else "completed")
    if mode == "full":
        assert created["subject"] == category["categories"] == deleted["subject"] == SECRET
        assert changed["entry_id"] == ev.EntryID
        assert response["response"] == "accept"
    else:
        assert created == ({"status": "created", **SLOT} if mode == "busy_only" else {"status": "created"})
        assert changed == category == {"status": "updated"}
        assert response == {"status": "responded"} and deleted == {"status": "deleted"}
    ev.Delete.assert_called_once()
    ev.Respond.assert_called_once()
    # A caller with write permission can explicitly change the native marker.
    ns = make_profile()
    appointment(ns)
    call = registered(ns, privacy={"default": mode})
    call("update_event", sensitivity=0, **by_id)
    assert call("get_event", **by_id)["subject"] == SECRET


@pytest.mark.parametrize("value", [None, False, True, 2.0, 9, "0", [], SECRET])
def test_unreliable_sensitivity_and_malformed_availability(value):
    ev = SimpleNamespace(Sensitivity=value, RecurrenceState=0, BusyStatus=SECRET, Start=SECRET, End=SECRET)
    for mode in PrivacyMode:
        assert classify_event(ev, mode).mode == max(PrivacyMode.BUSY_ONLY, mode)
    assert availability(ev) == {"subject": "Appointment", "busy_status": None, "start": None, "end": None}


@pytest.mark.parametrize("mode", ["full", "busy_only", "exclude"])
def test_unreadable_property_and_no_reads_of_private_fields(mode):
    ns = make_profile()
    ev = appointment(ns)
    type(ev).Sensitivity = PropertyMock(side_effect=RuntimeError(SECRET))
    forbidden = PropertyMock(side_effect=AssertionError("unexpected private field read"))
    for field in ("Subject", "Body", "Recipients", "Location", "Organizer", "Categories"):
        setattr(type(ev), field, forbidden)
    call = registered(ns, privacy={"default": mode})
    result = call("list_events", **WINDOW, **SELECT)
    if mode == "exclude":
        assert result["items"] == [] and result["count"] == 0
        with pytest.raises(Exception, match="Item not found"):
            call("get_event", entry_id=ev.EntryID, **SELECT)
    else:
        assert result["items"] == [SLOT]
        assert call("get_event", entry_id=ev.EntryID, **SELECT) == SLOT
    forbidden.assert_not_called()


@pytest.mark.parametrize("mode", ["busy_only", "exclude", "full"])
def test_private_exception_and_normal_master_share_id(mode):
    ns = make_profile()
    master = appointment(ns)
    master.RecurrenceState, master.Sensitivity = 1, 0
    occurrence = SimpleNamespace(EntryID=master.EntryID, Parent=master.Parent, RecurrenceState=3,
                                 Sensitivity=2, Subject=SECRET, BusyStatus=0, Start=START, End=END)
    master.GetRecurrencePattern.return_value.Exceptions = Collection([
        SimpleNamespace(Deleted=True), SimpleNamespace(Deleted=False, AppointmentItem=occurrence),
    ])
    items = master.Parent.Items
    items.IncludeRecurrences = False
    items.Restrict.side_effect = lambda _: [occurrence] if items.IncludeRecurrences else [master]
    call = registered(ns, privacy={"default": mode})
    for expanded in (False, True):
        rows = call("list_events", **WINDOW, include_recurrences=expanded, **SELECT)["items"]
        if mode == "exclude":
            assert rows == []
        elif mode == "busy_only":
            assert rows == [SLOT]
        else:
            assert rows[0]["subject"] == SECRET
    if mode == "exclude":
        with pytest.raises(Exception, match="Item not found"):
            call("get_event", entry_id=occurrence.EntryID, **SELECT)
    else:
        result = call("get_event", entry_id=occurrence.EntryID, **SELECT)
        if mode == "busy_only":
            assert result == SLOT
        else:
            assert result["body"] == SECRET


def test_inaccessible_exception_and_unknown_recurrence_state_fail_closed():
    ns = make_profile()
    ev = appointment(ns)
    ev.Sensitivity, ev.RecurrenceState = 0, 1
    ev.GetRecurrencePattern.side_effect = RuntimeError(SECRET)
    for state in (1, None, False, "0"):
        ev.RecurrenceState = state
        assert classify_event(ev, PrivacyMode.FULL).mode == PrivacyMode.BUSY_ONLY
        assert classify_event(ev, PrivacyMode.EXCLUDE).mode == PrivacyMode.EXCLUDE


def test_hidden_row_does_not_consume_limit_or_leak_in_mail_search():
    ns = make_profile()
    ev = appointment(ns)
    visible = SimpleNamespace(Parent=ev.Parent, Sensitivity=0, RecurrenceState=0, Subject="Visible",
                              Start=START, End=END)
    ev.Parent.Items.Restrict.return_value = [ev, visible]
    call = registered(ns, privacy={"default": "exclude"})
    result = call("list_events", **WINDOW, limit=1, **SELECT)
    assert [r["subject"] for r in result["items"]] == ["Visible"] and result["count"] == 1
    ev.Parent.Name = "Synthetic calendar"
    assert call("search_mails", folder="calendar", query="fictional", **SELECT)["items"] == []
    with pytest.raises(Exception, match="Item not found"):
        call("get_mail", entry_id=ev.EntryID, **SELECT)


def test_provider_failure_does_not_enter_response_or_logs(caplog):
    ns = make_profile()
    ev = appointment(ns)
    ev.Delete.side_effect = RuntimeError(SECRET)
    with pytest.raises(Exception, match="could not complete") as error:
        registered(ns)("delete_event", entry_id=ev.EntryID, **SELECT)
    assert SECRET not in str(error.value) + caplog.text


@pytest.mark.parametrize("mode,expected", [
    ("full", {"status": "created", **SLOT}), ("exclude", {"status": "created"}),
])
def test_conflicting_saved_sensitivity_is_unreliable(mode, expected):
    ns = make_profile()
    ev = appointment(ns)
    ev.Save.side_effect = lambda: setattr(ev, "Sensitivity", 0)
    result = registered(ns, privacy={"default": mode})(
        "create_event", subject=SECRET, sensitivity=2, busy_status=0, start=START.isoformat(), end=END.isoformat(), **SELECT)
    assert result == {**expected, "local_outcome": "saved", "exchange_sync": "unverified"}


def test_folder_metadata_hides_raw_counts_under_exclude():
    ns = make_profile()
    root = ns.Stores[1].GetRootFolder.return_value
    root.StoreID, root.Name, root.DefaultItemType = "store-b", "Synthetic folder", 1
    root.Items.Count, root.UnReadItemCount = 37, 29
    root.Folders = []
    folders = registered(ns, privacy={"default": "exclude"})(
        "list_folders", max_depth=1, response_format="json", **SELECT)
    row, = folders["items"]
    assert row["item_count"] is None and row["unread_count"] is None
